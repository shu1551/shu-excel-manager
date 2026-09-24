#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vba_mcp_server.py — vba_manager を MCP サーバー化する薄い窓口。

実体は vba_manager.py の関数群そのもの（この層は判断をしない）。常駐プロセスが
get_workbook の接続キャッシュを持ち続けるため、CLI の「1コマンド毎の COM 再接続」
が消える。いわば shell/batch の常駐版。接続の鮮度（ブックが閉じられた等）は
get_workbook 側の生存確認＋自動再接続に任せる。
ただし握りっぱなしにはしない: アイドル IDLE_RELEASE_SECS 秒で COM 参照を解放する
（常駐が参照を握ったままだと、ユーザーが×で閉じた Excel がゾンビ化するため）。

制約:
- stdout は JSON-RPC の通信線なので、コマンドの print は全て捕捉してツール結果で返す
- COM は専用ワーカースレッド1本に固定（呼び出しスレッドが変わっても STA を跨がない）
- input() 待ちで固まらないよう実行中は stdin を空にする（確認系は -y を付けて呼ぶ）
"""
import atexit
import gc
import io
import os
import queue
import sys
import threading
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
os.chdir(SCRIPT_DIR)

import pythoncom  # noqa: E402
import vba_manager  # noqa: E402
import vbam_core  # noqa: E402
from mcp.server.fastmcp import FastMCP  # noqa: E402

# _last_proc.vba の置き場は vbam_core.LAST_PROC_FILE の 1 か所（get・replace-procedure と同じ物を見る）。
# ここに別の定数を持つと、テストが vbam_core 側を差し替えても本物の _last_proc.vba を書いてしまう（2026-09-13）
# write_grid・patch_procedure の作業ファイルは常駐ごとに分ける（2026-09-24: Claude と Gemini の常駐が同時に動くと、
# 同じ _last_values.tsv を互いに上書きして、相手の表を書き込むことがあり得た）
LAST_VALUES = os.path.join(SCRIPT_DIR, f"_last_values_{os.getpid()}.tsv")
PATCH_TGT = os.path.join(SCRIPT_DIR, f"_patch_tgt_{os.getpid()}.vba")
PATCH_REP = os.path.join(SCRIPT_DIR, f"_patch_rep_{os.getpid()}.vba")
BLOCKED = {"shell", "batch"}  # 対話・標準入力前提のコマンドは MCP では使えない

# 本物の標準入出力（JSON-RPC の通信線）を起動時に控えておく。
# ワーカーは実行中 sys.stdout を自分のバッファに差し替えるが、復元先を
# 「入った時点の sys.stdout」にすると、世代交代（_restart_worker）後に
# 起動した新ワーカーが「前の世代の捨てられたバッファ」を掴んでそこへ復元し続ける。
# 以後ジョブ外の print（cleanup の DEBUG 行など）が誰も読まないバッファに溜まる。
# 復元先は常にこの実体にする。
REAL_STDIN, REAL_STDOUT, REAL_STDERR = sys.stdin, sys.stdout, sys.stderr

# アイドルでこの秒数コマンドが来なければ Excel への COM 参照を解放する。
# 常駐が Application/Workbook 参照を握ったままだと、ユーザーが Excel を×で
# 閉じてもプロセスが数十秒死なずゾンビ化し、その隙にブックを開くと死にかけ
# プロセスに合流して XLSTART(PERSONAL.XLSB) を読まない実害がある。
# 2026-09-16 夜に Gemini が 30 秒へ延ばしたが、9/17 に 5 秒へ戻した: 5 秒は上の対照実験で決めた値で、
# 再接続は同日の PID 確認（_pid_is_excel）で 36〜52 ms に軽くなっており、延ばして得る物が無い。
# 延ばしたいときだけ環境変数 VBAM_IDLE_SECS で。
IDLE_RELEASE_SECS = int(os.environ.get("VBAM_IDLE_SECS", "5"))

_jobs = queue.Queue()
_worker_lock = threading.Lock()


class _ThreadRouted(io.TextIOBase):
    """スレッドごとに出力先を振り分ける stdout/stderr/stdin の代理。

    従来はワーカーがジョブ毎に sys.stdout 自体を差し替えていた。しかし sys.stdout は
    プロセス全体のグローバルであってスレッドローカルではないため、次の穴が残る:

      1. W1 が MsgBox に捕まり 600 秒でタイムアウト → 世代交代で見捨てられる
      2. W2 がジョブを終え、finally で sys.stdout を REAL_STDOUT に戻す
      3. そこで人がダイアログを閉じ、W1 が復帰して print する
      4. その print は REAL_STDOUT ＝ JSON-RPC の通信線へ流れ、セッションが壊れる

    従来のガード（自分が据えたものが残っているときだけ戻す）が守っていたのは
    「復元時」だけで、「実行中の書き込み」は無防備だった。

    この代理を起動時に1回だけ据え、ワーカーは sys.stdout に二度と触らず
    「自分のスレッドのバッファ」を登録するだけにする。見捨てられた旧ワーカーは
    自分の（誰も読まない）バッファに書き続け、通信線には絶対に届かない。
    """
    _local = threading.local()

    def __init__(self, real, stream='out'):
        self._real = real
        self._stream = stream          # 'out' | 'err' | 'in'

    @classmethod
    def bind(cls, out_buf, err_buf, in_buf, json_mode=False):
        """このスレッドの出力先を登録する（ワーカーが自分のジョブ開始時に呼ぶ）"""
        cls._local.out = out_buf
        cls._local.err = err_buf
        cls._local.inp = in_buf
        cls._local.json_mode = json_mode

    @classmethod
    def unbind(cls):
        cls._local.out = None
        cls._local.err = None
        cls._local.inp = None
        cls._local.json_mode = False

    @classmethod
    def json_mode(cls):
        return getattr(cls._local, 'json_mode', False)

    def _target(self):
        if self._stream == 'out':
            return getattr(_ThreadRouted._local, 'out', None) or self._real
        if self._stream == 'err':
            return getattr(_ThreadRouted._local, 'err', None) or self._real
        return getattr(_ThreadRouted._local, 'inp', None) or self._real

    def write(self, s):
        return self._target().write(s)

    def flush(self):
        try:
            self._target().flush()
        except Exception:
            pass

    def readline(self, *a):
        return self._target().readline(*a)

    def read(self, *a):
        return self._target().read(*a)

    def isatty(self):
        return False

    @property
    def buffer(self):
        # MCP SDK の stdio_server は起動時に sys.stdin.buffer / sys.stdout.buffer を
        # 掴む（TextIOWrapper で包む）。TextIOBase には buffer が無いため、これを
        # 生やさないと mcp.run() が AttributeError でサーバー起動すらできない。
        # JSON-RPC の線は常に本物へ直結させる（ジョブのリダイレクトとは無関係）
        return self._real.buffer

    def fileno(self):
        return self._real.fileno()


# 起動時に1回だけ据える。以後 sys.stdout/stderr/stdin は誰も差し替えない
sys.stdout = _ThreadRouted(REAL_STDOUT, 'out')
sys.stderr = _ThreadRouted(REAL_STDERR, 'err')
sys.stdin = _ThreadRouted(REAL_STDIN, 'in')


def _install_json_aware_print():
    """--json のジョブでは、素の print()（情報行）を stderr 側へ退避する。

    JSON 本体は print(..., file=sys.stdout) と明示して出力される一方、
    情報行（「対象ブック: ...」等）は素の print() で出る。CLI では
    setup_encoding が builtins.print にパッチを当ててこれを分離しているが、
    MCP は setup_encoding を呼ばないため、両者が同じ stdout に混ざり
    機械処理側が json.loads できなかった。同じ分離をここで行う。
    スレッドローカル判定にしてあるので、他スレッドの print は巻き込まない。
    """
    import builtins
    _orig_print = builtins.print

    def _print(*args, **kwargs):
        if 'file' not in kwargs and _ThreadRouted.json_mode():
            kwargs['file'] = sys.stderr
        _orig_print(*args, **kwargs)

    builtins.print = _print


_install_json_aware_print()


def _release_com_refs():
    """接続キャッシュの Excel COM 参照を手放す（COM を作ったワーカースレッド上で呼ぶ）。

    ①get_workbook の接続キャッシュを手放す（ユーザーの Excel に参照を残さない）。
    ②ツールが自動起動した非表示 Excel（_created_instances）も畳む。
      常駐サーバーでは終了時の cleanup_excel が何時間も来ないため、放置すると
      アドインを読まない非表示 Excel が残留し、ユーザーが開いたファイルが合流して
      「アドインが効かない」事故になる（2026-07-12 特定・当日3体残留の実害）。
      未保存の変更を持つインスタンスだけは温存する（無言の変更破棄をしない）。
    """
    try:
        if vba_manager._wb_cache:
            vba_manager._wb_cache.clear()
            gc.collect()  # 参照サイクルに残った COM ラッパも確実に Release させる
    except Exception:
        pass
    try:
        vba_manager.release_created_instances()
    except Exception:
        pass


def _tokenize(line):
    # 切り分けは道具側（vbam_core.split_command_line）に置く＝reload_tools で直しが効く（2026-09-23）。
    # "…" の中の "" を " 1 つと読む（shlex は引用符を落とし、条件付き書式の式が壊れた）
    fn = getattr(vba_manager, "split_command_line", None)
    if fn is not None:
        return fn(line)
    import shlex
    lex = shlex.shlex(line, posix=True)
    lex.whitespace_split = True
    lex.escape = ''  # Windows パスの \ をエスケープ扱いしない（shell と同じ）
    lex.commenters = ''  # 行中の # をコメント扱いしない（#FF0000 等が消える。shell/batch と同じ）
    return list(lex)


def _wants_json(line):
    """その行が --json を求めているか（CLI は sys.argv を見るが MCP には無いので行を見る）"""
    try:
        return "--json" in _tokenize(line)
    except Exception:
        return "--json" in line


def _run_line(parser, table, line, buf):
    # 字句分け・形の外れの受け止め・台帳は道具側（vba_manager.run_command_line）に置く＝reload_tools で直しが効く
    # （2026-09-24 総点検: ここで parse_known_args していたので、無いコマンドに 130 個の一覧を返し、台帳にも残らなかった）
    fn = getattr(vba_manager, "run_command_line", None)
    if fn is not None:
        return fn(parser, table, line, blocked=tuple(BLOCKED))
    tokens = _tokenize(line)
    sub_args, unknown = parser.parse_known_args(tokens)
    unknown = [u for u in unknown if u not in ("--visible", "-v")]
    if unknown:
        buf.write(f"不明な引数/オプション: {' '.join(unknown)}\n")
        return False
    if not sub_args.command or sub_args.command in BLOCKED:
        buf.write("このコマンドは MCP セッション内では実行できません\n")
        return False
    return table[sub_args.command](sub_args)


def _worker(jobs):
    pythoncom.CoInitialize()
    parser = vba_manager.build_parser()
    table = vba_manager._command_table()
    while True:
        try:
            item = jobs.get(timeout=IDLE_RELEASE_SECS)
        except queue.Empty:
            # アイドル: 連続コマンド中は保持していた COM 参照をここで手放す。
            # 世代交代後に復帰した旧ワーカーは新世代のキャッシュに触らない
            if jobs is _jobs:
                _release_com_refs()
            continue
        if item is None or (isinstance(item, tuple) and len(item) == 2 and item[0] == "__retire__"):
            # 世代交代の停止合図（タイムアウト後に復帰した旧ワーカーはここで退場）。
            # 退場する前に、自分のスレッドで作った COM 参照を自分で手放す（2026-09-23）。
            # 別のスレッド（MCP の本線）から捨てると参照が本当には外れず、サーバーが動いている間ずっと残る
            # ＝使う人が×で閉じた Excel が終われず、窓の無いまま居座る（Windows を再起動するまで消えなかった）。
            stash = item[1] if item is not None else []
            item = None
            try:
                stash.clear()
            except Exception:
                pass
            stash = None
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            break
        line, box, done = item
        with _worker_lock:
            if box.get("cancelled"):
                # 実行前にタイムアウト済みのジョブ。「タイムアウトしました」と報告済みなのに
                # 後から副作用だけ走る、を防ぐためここで捨てる
                continue
            box["started"] = True
        # stdout（JSON本体）と stderr（情報行・警告）を分ける。混ぜると --json の
        # 出力に情報行が紛れ込み、機械処理側が json.loads できない
        out_buf = io.StringIO()
        err_buf = io.StringIO()
        ok = False
        try:
            # sys.stdout をグローバルに差し替えるのではなく、自分のスレッドの
            # 出力先を登録するだけ（_ThreadRouted 参照）。こうすると世代交代で
            # 見捨てられた旧ワーカーが後から復帰して print しても、その出力は
            # 自分のバッファに行き、JSON-RPC の通信線には絶対に漏れない。
            _ThreadRouted.bind(out_buf, err_buf, io.StringIO(""),
                               json_mode=_wants_json(line))
            if line == "__cleanup__":
                vba_manager.cleanup_excel()
                ok = True
            elif line == "__reload__":
                # 道具の .py を直した後、サーバーを再起動せずに読み直す（2026-09-06 夜・「MCP 要再起動」が 1 日 5 回）。
                # ワーカーの列で実行するので、走っている仕事の途中には割り込まない。parser と table も作り直す
                ok, msg = _reload_modules()
                parser = vba_manager.build_parser()
                table = vba_manager._command_table()
                out_buf.write(msg)
            else:
                ok = _run_line(parser, table, line, err_buf) is not False
        except SystemExit as e:
            ok = e.code in (0, None)
        except EOFError:
            err_buf.write("\nエラー: 確認プロンプト待ちになりました（MCP では応答できません）。"
                          "確認付きコマンドは -y を付けて再実行してください")
        except Exception as e:
            err_buf.write(f"\nエラー: {e}")
        finally:
            _ThreadRouted.unbind()
            box["ok"] = ok
            box["out"] = out_buf.getvalue()
            box["err"] = err_buf.getvalue()
            done.set()


_RELOAD_ORDER = ['vbam_core', 'vbam_view', 'vbam_edit', 'vbam_vba', 'vbam_form2vba', 'vbam_build',
                 'vbam_recipes', 'vbam_heavy',
                 # 2026-09-11 の分割: 下の層から順に（from X import * で結んでいるので、依存の順でないと古い物が残る）
                 'vbam_keys', 'vbam_ai', 'vbam_hands', 'vbam_inv', 'vbam_ledger', 'vbam_undo', 'vbam_grade', 'vbam_view_ai',
                 'vbam_agent', 'vbam_fire', 'vbam_macro', 'vbam_clean', 'vbam_shake',
                 'vbam_prefire', 'vbam_forge',   # 2026-09-17 先撃ちの登録簿と鍛える回路（vbam_agent の上・遅延 import）
                 'format_bas',              # 2026-09-16 format-module の中身（vbam_devtools が関数の中で import する）
                 'vbam_devtools',           # 2026-09-16 職場向けの手（vbam_core・vbam_vba の上・vba_manager の下）
                 'vbam_lineage',            # 2026-09-17 系譜と閉じたブック（vbam_devtools の上）
                 'vbam_audit',              # 2026-09-17 数式・データ総合診断エンジン
                 'vba_manager']   # 依存の順（from X import はこの順で新しくなる）


def _reload_order(folder=None):
    """読み直す順を、各 .py の行頭の import（from X import * / import X）から組む＝依存される側が先。

    手で並べた _RELOAD_ORDER は vbam_view を vbam_vba より前に置いていて、vbam_vba に足した関数が
    vbam_view の from vbam_vba import * に届かなかった（2026-09-24 総点検: seiri が「run_book_macro is not defined」）。
    関数の中の import（字下げあり）は数えない＝循環を作らない。並びの同点は _RELOAD_ORDER の順。
    """
    import glob
    import re
    folder = folder or os.path.dirname(os.path.abspath(__file__))
    names = list(_RELOAD_ORDER)
    for p in sorted(glob.glob(os.path.join(folder, "vbam_*.py"))):
        n = os.path.splitext(os.path.basename(p))[0]
        if n not in names and not n.startswith("vbam_test"):
            names.insert(len(names) - 1, n)            # vba_manager は最後のまま
    known = set(names)
    pat = re.compile(r"^(?:from\s+(\w+)\s+import|import\s+([\w\s,]+?)(?:\s+as\s+\w+)?\s*(?:#.*)?$)", re.M)
    deps = {}
    for n in names:
        try:
            with open(os.path.join(folder, n + ".py"), encoding="utf-8") as f:
                src = f.read()
        except OSError:
            deps[n] = []
            continue
        ds = []
        for a, b in pat.findall(src):
            for d in ([a] if a else [x.strip() for x in b.split(",")]):
                if d in known and d != n:
                    ds.append(d)
        deps[n] = ds
    order, state = [], {}

    def visit(n):
        if state.get(n):
            return                                      # 済み・または辿っている途中（循環）
        state[n] = 1
        for d in deps.get(n, []):
            visit(d)
        order.append(n)

    for n in names:
        visit(n)
    return order


def _reload_modules():
    """道具のモジュールを依存の順に importlib.reload する → (ok, 報告文)。ワーカーのスレッドで呼ぶ。"""
    import importlib
    done, skipped, failed = [], [], []
    try:
        seq = _reload_order()
    except Exception:
        seq = list(_RELOAD_ORDER)
    for n in seq:
        m = sys.modules.get(n)
        if m is None:
            skipped.append(n)
            continue
        try:
            importlib.reload(m)
            done.append(n)
        except Exception as ex:
            failed.append(f"{n}: {ex}")
    try:
        vba_manager._wb_cache.clear()
    except Exception:
        pass
    msg = "読み直しました: " + (", ".join(done) or "（なし）")
    if skipped:
        msg += "\n（まだ読み込まれていなかった: " + ", ".join(skipped) + "）"
    if failed:
        msg += "\n読み直せなかった: " + " / ".join(failed) + "\n（構文エラーなら直してもう一度 reload。それでも駄目なら再起動）"
    msg += "\n（サーバー本体 vba_mcp_server.py を変えたときだけは再起動が要ります）"
    return not failed, msg


def _restart_worker():
    """タイムアウトで詰まったワーカーを見捨てて新しい世代に交代する。

    ワーカーは1本のキュー直列なので、詰まったジョブを放置すると以後の
    全ツール呼び出しがタイムアウトになる（サーバー再起動まで回復不能）。
    旧キューには停止合図を置き、旧ワーカーが後で復帰しても新ジョブを食わせない。
    新ワーカーは別スレッド＝別 STA なので、旧スレッドの COM 参照は使えない。
    接続キャッシュを捨てて新規接続からやり直す。
    """
    global _jobs
    with _worker_lock:
        old = _jobs
        _jobs = queue.Queue()
        # 接続キャッシュの中身は旧ワーカーのスレッドで作った COM 参照。ここ（MCP の本線）で捨てると
        # 参照が本当には外れず、Excel が終われなくなる（2026-09-23 のゾンビの正体）。
        # 中身を包みに移して旧ワーカーに渡し、旧ワーカーが退場するときに自分のスレッドで手放す。
        stash = []
        try:
            stash = list(vba_manager._wb_cache.items())
            vba_manager._wb_cache.clear()      # 包みが握っているので、ここでは参照は外れない
        except Exception:
            pass
        old.put(("__retire__", stash))
        stash = None
        threading.Thread(target=_worker, args=(_jobs,), daemon=True).start()


def _run(line, timeout=600):
    box, done = {}, threading.Event()
    with _worker_lock:
        # _restart_worker と排他にする。参照と put の間に世代交代が挟まると、
        # ジョブが停止済みの旧キューに落ちて誰にも実行されない
        _jobs.put((line, box, done))
    if not done.wait(timeout):
        with _worker_lock:
            started = box.get("started", False)
            if not started:
                box["cancelled"] = True   # ワーカーは実行前にこれを見て捨てる
        if not started:
            # 前のジョブが長引いて未着手のままタイムアウト。ワーカー自体は健全なので
            # 世代交代せず、このジョブだけ取り下げる
            # 受け側（_submit / get_procedure）は3要素で受ける。2要素で返すと
            # ここが ValueError になり、この案内文自体が届かない
            return False, "", (f"タイムアウト（{timeout}秒）: {line}\n"
                               "前のコマンドが長引いているため、このコマンドは未実行のまま"
                               "取り下げました。しばらくしてから再実行してください")
        _restart_worker()
        return False, "", (f"タイムアウト（{timeout}秒）: {line}\n"
                           "ワーカーを再起動しました。次の呼び出しから復旧します"
                           "（実行中だったコマンドは Excel 側で続いている可能性があります。"
                           "モーダルダイアログが開いていないか確認してください）")
    return (box.get("ok", False),
            (box.get("out") or "").strip(),
            (box.get("err") or "").strip())


def _log_call(line):
    """呼ばれたコマンドを1行ずつ追記する（環境変数 AI_WIN_LOG があるときだけ）。

    AI作業窓が「いま AI が何をしているか」を出すための経路。ヘッドレスで起動された
    子プロセスにだけ環境変数を渡すので、通常のセッションでは何も書かない。
    """
    path = os.environ.get("AI_WIN_LOG", "")
    if not path:
        return
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%H:%M:%S')}  {line.strip()[:160]}\n")
    except Exception:
        pass


# AI作業窓の承認ゲート（第2段・2026-08-01）。
# 線引きは「書き/読み」ではなく「戻せる/戻せない」——マクロ置換は退避→restore で
# 戻せるので通し、シートやブックに退避の無い操作だけ人間の承認にかける。
GATE_COMMANDS = {"write-range", "clear-range", "delete-module", "save-as", "close"}
GATE_WAIT = 180  # 秒。フォームの確認ダイアログで人間が答えるまで待つ上限


def _gate_check(line, wait=GATE_WAIT):
    """AI作業窓から起動された子セッション（AI_WIN_LOG あり）では、戻せない
    コマンドを実行前に人間の承認にかける。承認なら None（続行）、
    却下・無応答なら拒否文を返す（呼び出し元がそのまま AI への返答にする）。

    配線: AI_WIN_LOG と同じフォルダに ai_win_ask.txt を置く → フォームの
    ポーリングが気づいて確認ダイアログを出す → ai_win_answer.txt に yes/no。
    通常セッション（環境変数なし）では何もしない。
    """
    log = os.environ.get("AI_WIN_LOG", "")
    if not log:
        return None
    try:
        head = _tokenize(line)[0]
    except Exception:
        head = (line.strip().split() or [""])[0]
    if head not in GATE_COMMANDS:
        return None
    return _ask_user(log, line.strip(), head, wait)


def _ask_user(log, body, label, wait):
    """AI作業窓に確認を出して yes/no を待つ。承認なら None、それ以外は拒否文。

    配線: ai_win_ask.txt を置く → フォームのポーリングが MsgBox を出す →
    ai_win_answer.txt に yes/no。窓に聞けない異常時は素通し（方針A）。
    """
    d = os.path.dirname(log) or "."
    ask = os.path.join(d, "ai_win_ask.txt")
    ans = os.path.join(d, "ai_win_answer.txt")
    for p in (ask, ans):
        try:
            os.remove(p)
        except OSError:
            pass
    try:
        # 書き終えてから名前を付ける（2026-09-17: 直に書くと、フォームのポーリングが書きかけの空の文を読んで
        # 確認の中身が出ないことがあった。テストが負荷の高いときだけ落ちて見つかった）
        with open(ask + ".tmp", "w", encoding="utf-8") as f:
            f.write(body)
        os.replace(ask + ".tmp", ask)
    except OSError:
        return None      # 窓に聞けない異常時は方針A（全開）どおり素通し
    deadline = time.time() + wait
    while time.time() < deadline:
        if os.path.exists(ans):
            try:
                with open(ans, encoding="utf-8-sig") as f:
                    verdict = f.read().strip().lower()
            except OSError:
                verdict = ""
            try:
                os.remove(ans)
            except OSError:
                pass
            if verdict == "yes":
                return None
            return (f"この操作（{label}）はユーザーが確認画面で却下しました。"
                    "実行していません。別のやり方を提案するか、指示を仰いでください。")
        time.sleep(0.3)
    try:
        os.remove(ask)
    except OSError:
        pass
    return (f"この操作（{label}）は確認の返事が {wait} 秒以内に無かったため"
            "実行しませんでした。ユーザーに口頭で確認してください。")


def _diff_confirm(line):
    """マクロ置換の差分を、当てる前に作業窓へ見せて承認をとる（2026-08-01）。

    退避→restore で戻せる操作なのでゲート対象ではないが、「戻せること」と
    「間違いに当てる前に気づけること」は別物（Enhancer 4 Google が適用前に
    diff を見せているのを見て採用）。

    置換予定のコードは _last_proc.vba にある。そこからプロシージャ名を拾い、
    `get --out`（別ファイルへ退避。_last_proc.vba を潰さないため）で現物を取って
    突き合わせる。材料が揃わないときは黙って素通しする——確認の材料が無いのに
    聞いても判断のしようがないため。
    """
    log = os.environ.get("AI_WIN_LOG", "")
    if not log:
        return None
    try:
        import difflib
        import re as _re
        import tempfile
        import vbam_core

        with open(vbam_core.LAST_PROC_FILE, encoding="utf-8") as f:
            new_lines = f.read().replace("\r\n", "\n").split("\n")
        m = _re.search(r"^\s*(?:Public\s+|Private\s+|Friend\s+)*(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_぀-鿿]"
                       r"[\w぀-鿿]*)", "\n".join(new_lines), _re.MULTILINE | _re.IGNORECASE)
        if not m:
            return None
        proc = m.group(1)

        tokens = _tokenize(line)
        module = ""
        if "--module" in tokens:
            i = tokens.index("--module")
            if i + 1 < len(tokens):
                module = tokens[i + 1]

        tmp = os.path.join(tempfile.gettempdir(), "ai_win_before.vba")
        cmd = f'get {module + " " if module else ""}{proc} --out "{tmp}"'
        ok, _out, _err = _run(cmd)
        if not ok:
            return None
        with open(tmp, encoding="utf-8") as f:
            old_lines = f.read().replace("\r\n", "\n").split("\n")
    except Exception:
        return None

    while new_lines and not new_lines[-1].strip():
        new_lines.pop()
    while old_lines and not old_lines[-1].strip():
        old_lines.pop()
    body = "\n".join(difflib.unified_diff(old_lines, new_lines,
                                          fromfile="現在", tofile="変更後", lineterm=""))
    if not body.strip():
        return None      # 変更なし＝本体側が置換をスキップする
    return _ask_user(log, f"【{proc} を書き換えます。よろしいですか】\n" + body,
                     "replace-procedure", GATE_WAIT)


def _submit(line, timeout=600):
    _log_call(line)
    denial = _gate_check(line)
    if denial:
        return denial
    try:
        head = _tokenize(line)[0]
    except Exception:
        head = ""
    if head == "replace-procedure":
        denial = _diff_confirm(line)
        if denial:
            return denial
    ok, out, err = _run(line, timeout)
    return _format_result(line, ok, out, err)


def _format_result(line, ok, out, err):
    """ワーカーの結果 → ツールが返す文字列（同期・非同期で同じ形にするため分けた・2026-09-05）。"""
    # --json 指定時は JSON 本体（stdout）だけを返す。情報行（stderr）を混ぜると
    # 機械処理側が json.loads できない（CLI では setup_encoding が情報行を stderr へ
    # 退避して分離しているが、MCP は setup_encoding を呼ばないので自前で分ける）
    if _wants_json(line):
        if ok:
            # JSON の後ろに警告行を足すと呼び出し側の json.loads が落ちるため、
            # 成功時は JSON 本体だけを返す（stderr の警告はここでは運ばない契約。
            # 警告も機械で受けたいコマンドは run-macro --json の dialogs_dismissed の
            # ように JSON 自身のフィールドで返す）
            if out:
                return out
            # 成功で stdout が空＝異常ではない。「失敗しました」を付けて返すと嘘になる
            return err or "（出力なし）"
        # 失敗時だけは理由が要る（黙って空を返さない）
        return ((out + "\n" + err).strip() + "\n（コマンドは失敗しました）").strip()
    body = "\n".join(p for p in (err, out) if p).strip()
    if not ok:
        body = (body + "\n（コマンドは失敗しました）").strip()
    return body or "（出力なし）"


_live_jobs = {}          # 非同期で走らせた仕事: id → {'box','done','line','start'}
_LIVE_JOBS_KEEP = 5      # 覚えておく数（終わって回収した分は消える）


def _run_async(line):
    """ジョブをキューに入れて、待たずに id を返す（agent(wait=False) 用・2026-09-05）。

    ワーカーは 1 本の直列キューなので、この仕事が終わるまで他の vba コマンドは順番待ちになる。
    様子を見るための agent_status はキューを通さない（進み具合ファイルを読むだけ）。
    """
    box, done = {}, threading.Event()
    with _worker_lock:
        _jobs.put((line, box, done))
    jid = f"job{int(time.time())}"
    if jid in _live_jobs:
        jid += f"_{len(_live_jobs)}"
    _live_jobs[jid] = {'box': box, 'done': done, 'line': line, 'start': time.time()}
    for old in [k for k, v in sorted(_live_jobs.items(), key=lambda kv: kv[1]['start'])][:-_LIVE_JOBS_KEEP]:
        _live_jobs.pop(old, None)
    return jid


def _submit_async(line):
    """関所は同期と同じに通し、待たずに id を返す。"""
    _log_call(line)
    denial = _gate_check(line)
    if denial:
        return denial
    return _run_async(line)


# 2026-09-18 サーバーの名前を vba-manager → excel-manager に（秀エクセルマネージャーとして公開する準備）。
# ファイル名・コマンド名（vba_manager.py／vba(...)）は変えていない。
# 2026-09-23 つないだ AI に渡す説明（instructions）を足した。棚撃ち式（seiri → shelf-run）は Claude にだけ
# 起動時のフックで流していて、Gemini や公開の読者の AI には見えていなかった。
_INSTRUCTIONS = """秀エクセルマネージャー（excel-manager）。道具は今アクティブに開いている Excel ブックに効く（保存はしない）。
シートを見るのは vba("materials")（read-range を重ねない）。表について答えるときは番地を添える。

表を直す・整える・点検する依頼は、自分で書く前に「棚」を使う。棚＝秀コンボのモジュール「表の整理」「表の整理_作る」
「表の整理_調べる」の、言葉で頼んで撃つマクロ（表の書き方と罫線と列幅をそろえる・全列が同じ重複行を削除する・空行を詰める・番号の列を連番に振り直す・表の下に合計行を足す・
集計表やグラフを作る・点検の一覧 など）:
 1. vba("seiri 頼みの文") を 1 回（先撃ち）。頼みの文をそのまま渡すと、表を整えるマクロ・頼みに当たる棚のマクロ・tidy の仕上げまで
    道具が 1 回の呼び出しで撃ち、直したセル（前→後）と残り（エラーの式・気づき）を返す。頼みの語が無い「表の修正」だけなら、
    シートに書いてある頼みの文を渡す。棚を 1 手ずつ撃たない（手の間ごとに考える分だけ遅くなる）。
 2. 残りがあるときだけ、表と式を見て vba("write-cells C7 値 …") / write_grid と vba("tidy 表の範囲") を同じ返事に並べて 1 回で撃つ。
    マイナス・桁違い・番号の重複・式に直書きの数など、人の判断が要るものは直さずに番地で報告する。
 3. 報告は要点だけ 1〜3 行（何をどこに直したかを番地で。例: C7・C11 の郵便番号をハイフンつきに／F・G 列に突き合わせの式）。
    経過・秒数・長い説明は書かない。戻すのは agent(undo=True)。
棚を 1 本だけ撃つ: vba("shelf-run 名前 [--select A1:A9,C1:C9]")（選ぶ列は 1 列ずつカンマで）。目録: vba("shelf --grep 語")。
式がどこから来ているか: vba("trace D31 --depth 3")（シートをまたいで番地・式・値の木）。
棚が無い（秀コンボを読み込んでいない）ときは、seiri はマクロを撃たずに残りだけを返す＝手で直す。"""
mcp = FastMCP("excel-manager", instructions=_INSTRUCTIONS)


@mcp.tool()
def vba(command: str) -> str:
    """vba_manager のコマンドを1行実行する。対象は今アクティブに開いている Excel ブック。

    CLI と同じ引数列をそのまま渡す。例:
      "list"（マクロ一覧） / "list-open"（開いているブック一覧） /
      "get モジュール名 プロシージャ名" / "run-macro マクロ名" /
      "read-range A1:D10" / "write-range A1 値" / "sheet-info" /
      "grep ActiveSheet" / "checkup" / "impact マクロ名" /
      "close-form"（表示中の UserForm を閉じる。フォームを直す前に人へ頼まず自分で閉じる） /
      "materials"（シートを触る前の材料。開いているシートを Excel に聞く） /
      "write-cells C7 値 C11 値 --show"（飛び飛びのセルを1回で） /
      "tidy A5:G13 I5:L11"（表の仕上げ＝見出し・罫線・番号列は左寄せ・数値列は#,##0・列幅。
      列を足したときは足した列だけでなく表全体の範囲を渡す）
    表を直す・整える・点検する依頼は、自分で書く前に棚のマクロを使う＝先撃ち:
      "seiri 頼みの文" を 1 回（表を整えるマクロ・頼みに当たる棚・tidy まで 1 回で撃つ。頼みの語が無ければ
      シートに書いてある頼みの文を渡す）→ 残りがあるときだけ write-cells / write_grid と "tidy" を同じ返事に並べて 1 回。
      報告は要点だけ 1〜3 行・番地で（経過・秒数・長い説明は書かない）。戻すのは agent(undo=True)。
      棚を 1 本だけ撃つ "shelf-run 名前 --select A1:A9,C1:C9" / 目録 "shelf --grep 語"。式の元をたどる "trace D31 --depth 3"。
    表・数式の列など TSV で書く範囲は write_grid（TSV を文字列で直接渡す。ファイル不要）。
    よく外す手の形（2026-09-19 Gemini の実射で、形を外して使い方が返るだけの往復が 11 回あった）:
      "pivot create Sheet1!A1:G13 --rows 部署 --cols 区分 --values 金額 --func sum --sheet 集計 --name P1" /
      "pivot-field set-format P1 金額 #,##0" / "pivot-field sort P1 部署 desc" /
      "chart create A1:B5 --type column --title 題 --name G1 --at H3"（ピボットからは --pivot P1） /
      "chart-config legend G1 bottom" / "chart-config data-labels G1 --value" /
      "slicer add P1 部署 --name S1 --at A11" / "slicer list" /
      "shape --list"（図形の名前と位置） / "shape G1 --left 216 --top 135 --width 300 --height 180"（動かす・大きさ） /
      "sheet add 新シート --before 既存" / "format-range B2:N2 --merge --bold --bg #1F4E79 --color #FFFFFF --size 16" /
      "format-range E6:E19 --number-format yyyy/m/d"（日付の形をそろえる） / "fill D6:D17"（先頭のセルの式を下へ写す） /
      cond-format A6:D17 --formula "=$B6<$C6" --bg #FFC7CE（条件付き書式。式に文字を入れるときは " を "" と重ねる。
      2026-09-23 Gemini が使い方を引いた 3 つ） /
      "add-module Module1 -y" → set_procedure_code(全文) → "add-procedure Module1 -y"（新しい Sub を足す。直すのは replace_procedure）
    手数の多い組み立て（ダッシュボード・帳票の作り直しなど、10 手を超えそうなもの）は、一手ずつ撃たずに
    VBA のマクロ 1 本に書いて "compile" → "run-macro 名前" で撃つ（同じ仕事が 30 往復・数分 → 1 本・約 1 秒。二度目からは AI も要らない）。
    コマンド一覧・各引数は vba_help で確認できる。
    注意: 確認プロンプトを出すコマンドは必ず -y を付ける（例: "replace-procedure -y"）。
    shell / batch は使えない（このセッション自体が常駐＝接続使い回しのため不要）。
    "reload" は道具の .py を読み直す（reload_tools と同じ。サーバーの再起動は要らない）。
    長い手（checkup・export-all・gate・rehearse・snapshot・docs）は行末に "--bg" を付けると待たずに job id が返る
    （2026-09-17）。段は "status"、結果は agent_status("job…")。
    """
    line = command.strip()
    if line.lower() == "reload":
        return reload_tools()
    if line.endswith("--bg"):
        # 待たずに返す（無言消し・2026-09-17）。関所は _submit_async の中で同期と同じに通る。
        line = line[:-len("--bg")].rstrip()
        jid = _submit_async(line)
        if not jid.startswith("job"):
            return jid                       # 関所で断られた（そのまま理由を返す）
        return (f"走らせました（待たずに返しています）: {jid}\n"
                f"  段は vba(\"status\")、結果は agent_status(\"{jid}\") で（終わっていれば全文が返ります）。\n"
                "  この仕事が終わるまで、他の vba コマンドは順番待ちになります。")
    return _submit(line)


@mcp.tool()
def vba_help(command: str = "") -> str:
    """コマンド一覧（引数なし）または個別コマンドの詳細ヘルプ（例: command="get"）を返す。"""
    return _submit(f"{command} --help" if command else "--help")


@mcp.tool()
def get_procedure(name: str, module: str = "") -> str:
    """プロシージャのコードを取得して返す（_last_proc.vba にも保存される）。

    同名プロシージャが複数モジュールにある場合は module を指定する。
    修正の流れ: get_procedure → replace_procedure(code=修正後の全文)
    """
    cmd = f'get "{module}" "{name}"' if module else f'get "{name}"'
    ok, out, err = _run(cmd)
    # get のコード本文は stdout、情報行・警告は stderr。人が読むので両方返す
    body = "\n".join(p for p in (err, out) if p).strip()
    if not ok:
        return (body + "\n（取得に失敗しました）").strip()
    # get の出力に保存先とコード全文が含まれる（_last_proc.vba と同一内容）ので
    # ファイルを読み直して二重に返さない
    return body


def _write_last_proc(code):
    """修正後のコードを _last_proc.vba に書く → 行数。CRLF で来ても改行を二重にしない（write_grid と同じ）。"""
    text = code.replace('\r\n', '\n')
    with open(vbam_core.LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(text)
    return text.count('\n') + 1


@mcp.tool()
def set_procedure_code(code: str) -> str:
    """修正後のプロシージャコードを _last_proc.vba に書き込む（UTF-8）。

    Sub〜End Sub（または Function〜End Function）まで丸ごと渡す。
    この後 replace_procedure を呼ぶと差分表示つきで適用される
    （replace_procedure に code を渡せば、この手は要らない）。
    """
    n = _write_last_proc(code)
    return f"_last_proc.vba に {n} 行を書き込みました。replace_procedure で適用できます。"


@mcp.tool()
def replace_procedure(module: str = "", code: str = "") -> str:
    """プロシージャを置換する（自動バックアップ・差分表示つき）。

    code に修正後の全文（Sub〜End Sub）を渡すと、書き込みと置換を 1 手で済ませる（2026-09-13）。
    code を省くと、get_procedure／set_procedure_code で _last_proc.vba に置いた内容で置換する。
    同名プロシージャが複数ある場合は module で対象モジュールを明示する。
    """
    if code:
        _write_last_proc(code)
    cmd = "replace-procedure -y" + (f' --module "{module}"' if module else "")
    return _submit(cmd)


@mcp.tool()
def patch_procedure(name: str, target: str, replacement: str, module: str = "") -> str:
    """プロシージャ内の指定コードをピンポイントで置換する（差分置換・自動バックアップ・Attribute保持）。

    全文の上書きが不要になり、狙った行だけを安全・確実に変更できます。
    name: プロシージャ名（例: "印刷実行"）
    target: 置換前の既存コード（プロシージャ内で一意に特定できる文字列）
    replacement: 置換後の新しいコード
    module: 対象モジュール名（同名プロシージャが複数ある場合）
    """
    tgt_path, rep_path = PATCH_TGT, PATCH_REP
    with open(tgt_path, "w", encoding="utf-8") as f:
        f.write(target.replace("\r\n", "\n"))
    with open(rep_path, "w", encoding="utf-8") as f:
        f.write(replacement.replace("\r\n", "\n"))
    cmd = f'patch-procedure "{name}" --target-file "{tgt_path}" --replacement-file "{rep_path}" -y'
    if module:
        cmd += f' --module "{module}"'
    return _submit(cmd)


@mcp.tool()
def write_grid(range: str, tsv: str) -> str:
    """まとまった範囲（表・数式の列）を TSV 文字列で書き、書いた直後の見え方（画面の文字・###）を返す。

    range は左上セルか全範囲（例 "F5:G13"）。tsv は 1 行＝1 行・タブ＝列の文字列
    （'=' 始まりは数式。引用符もそのまま渡せる）。UTF-8 で _last_values.tsv に書いてから
    write-range --tsv --show を呼ぶので、ファイルを書く手も置き場所を考える手も要らない
    （2026-09-02 深夜・TSV の置き場と依存の順番で止まった 129 秒の記録）。
    文字列が日付・数値に化けたら道具が文字列に戻す（--raw を考えない）。
    飛び飛びのセルは vba("write-cells C7 値 C11 値 … --show")、仕上げは vba("tidy 範囲")。
    """
    text = tsv.replace('\r\n', '\n')
    if not text.endswith('\n'):
        text += '\n'
    with open(LAST_VALUES, 'w', encoding='utf-8') as f:
        f.write(text)
    return _submit(f'write-range "{range}" --tsv "{LAST_VALUES}" --show')


@mcp.tool()
def diagnose(range: str = "", sheet: str = "", fix: bool = False) -> str:
    """表・シートの数式＆データ総合診断・自動修復を実行する（集計漏れ・定数直書き・非一貫数式・循環参照・隠れ空白・外れ値検知）。

    100点満点のスコア、重大欠陥、警告、改善提案を含む総合監査レポートを返す。
    range: 対象範囲（例: "A1:G20"、省略時はシート使用範囲）
    sheet: 対象シート名（省略時はアクティブシート）
    fix: True で隠れ空白の除去と文字列の数字の数値化を書き戻す（人の値を書き換える手＝依頼に承認の語があるときだけ。
         控えを backups に取る。先頭ゼロ・数式のセルは触らない。集計漏れの式は提案を並べるだけで書かない）
    """
    cmd = "diagnose"
    if range:
        cmd += f' "{range}"'
    if sheet:
        cmd += f' --sheet "{sheet}"'
    if fix:
        cmd += ' --fix -y'
    return _submit(cmd)


@mcp.tool()
def style_map(range: str = "", sheet: str = "") -> str:
    """表・シートの視覚レイアウト＆書式マップ（背景色・フォント・二重罫線・複合見出しツリー）を抽出する。

    人間が見ている画面のデザイン構造、ヘッダー階層、強調セル、合計行の配置をAIに伝える。
    range: 対象範囲（例: "A1:G20"、省略時はシート使用範囲）
    sheet: 対象シート名（省略時はアクティブシート）
    """
    cmd = "style-map"
    if range:
        cmd += f' "{range}"'
    if sheet:
        cmd += f' --sheet "{sheet}"'
    return _submit(cmd)


@mcp.tool()
def agent(request: str = "", sheet: str = "", mode: str = "", macro: str = "", max_turns: int = 4,
          dry_run: bool = False, new_book: bool = False, ai: str = "", model: str = "", recipe: str = "",
          cont: bool = False, undo: bool = False, force: bool = False,
          no_image: bool = False, image: bool = False, no_grade: bool = False, grade: bool = False,
          score: bool = False,
          rehearse: bool = False, grade_ai: str = "", grade_model: str = "",
          runs: bool = False, cases: bool = False, keep_case: str = "",
          drop_case: str = "", backups: bool = False, undo_to: str = "",
          recipes: bool = False, request_file: str = "", wait: bool = True,
          changes: bool = False, prune_days: int = 0, shake: str = "", only: str = "", imagine: int = 0) -> str:
    """依頼文を 1 つ渡すと、道具（Python）が Excel のアクティブブックにループを回す薄い入口。

    使い方（要点）:
    - request="依頼文" [sheet="シート名"]。道具が材料を集め、AI に聞き、セルに書き、読み戻し・仕上げ検査・
      採点まで回して報告と請求書（字数・秒・トークン）を返す。保存はしない（気に入らなければ保存せず閉じる）。
    - mode: sheet（既存シートを直す・既定）| build（白紙／新しいシートを組む。new_book=True でまっさらなブックに）|
      macro（マクロを修理・作成。macro="名前"。rehearse=True で試し撃ちの手を許す）。省略時は道具が決めて 1 行目で言う。
    - recipe="手順書名"（recipes=True で 22 本の一覧）。そのとき request は補足＝承認の言葉や番地。
    - 人の値を消す・書き換える手は、依頼文に承認の言葉（消してよい・置き換えてよい・上書きしてよい 等）があるときだけ
      実行する。無ければ安全側で仮に処理し、report の末尾【人に判断してほしいこと】に質問が並ぶ → cont=True と
      request="はい" で続きを回す。
    - 戻す: undo=True（直前の控え）／backups=True で一覧 → undo_to="番号"（force=True で後の走行があっても戻す）。
      prune_days=N で控えの置き場の N 日より古いものを数えて見せる（force=True のときだけ消す）。
      dry_run=True は Excel に触らず 1 往復目の手だけ見る。
    - 見る: changes=True（直前の仕事で変わったセルの明細＝シート・番地・前・後と書式の変化。Excel に触らない）／
      runs=True（本番の走行台帳。人が戻した・保存した印つき）／score=True（実射の点数）／cases=True・
      keep_case="名前"・drop_case="名前"（本番の走行を弾にする）。
    - 調整: max_turns（既定 4。関所の差し戻しは別枠）／image=True（AI に画像を見せる。既定は見せない＝2026-09-08 に反転。
      no_image は互換のため残っているだけ）／no_grade=True（採点を
      止める）／grade_ai・grade_model（採点だけ別の AI）／ai・model（既定 claude-code＝ヘッドレスの Claude Code・sonnet・
      鍵なし。gemini／claude は API 鍵）／request_file（依頼文をファイルで）。
    - 揺らし（2026-09-11 夜）: shake="latest"（直前の記録）か記録のパスで、返事を Python で書き換えた 22 通りを
      AI なしで開いているブックに再生（一発合格しなかった版＝直す候補）。imagine=N で AI（既定 haiku）に N 通り
      書かせた版も足す。only="名前,名前" で絞る。dry_run=True は版の一覧だけ。
    - wait=False は待たずにジョブ id を返し、agent_status(id) で様子を見る。
    - 道具の .py を直した後は reload_tools() で読み直せる（サーバーの再起動は要らない）。
    仕組みの詳細（手 44 本・関所・不変条件・控え・台帳・経緯）は excel-vba-manager の SKILL.md「agent」の段にある。
    """
    import shlex
    if recipes:
        return _submit("agent --recipes", timeout=120)
    # 人が報告を検分する口・控えの間引き（2026-09-10・CLI にだけあって MCP から渡せなかった）
    if changes:
        return _submit("agent --changes", timeout=120)
    if prune_days:
        return _submit(f"agent --prune-days {int(prune_days)}" + (" --force" if force else ""), timeout=120)
    # '-' で始まる依頼文（「-5%で計算して」等）は argparse がオプションと読む＝shlex.quote では守れない。
    # 既にある --request-file の口に逃がす（2026-09-04）
    if request and request.lstrip().startswith('-') and not request_file:
        import tempfile
        fd, request_file = tempfile.mkstemp(suffix='.txt', prefix='vbam_req_', text=True)
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(request)
        request = ""
    cmd = "agent" + (" " + shlex.quote(request) if request else "")
    if rehearse:
        cmd += " --rehearse"
    if request_file:
        cmd += " --request-file " + shlex.quote(request_file)
    if cont:
        cmd += " --continue"
    if undo or undo_to:
        cmd += " --undo" + (" " + shlex.quote(undo_to) if undo_to else "")
        if force:
            cmd += " --force"
    if backups:
        cmd += " --backups"
    if drop_case:
        cmd += " --drop-case " + shlex.quote(drop_case)
    if no_image:
        cmd += " --no-image"
    if image:
        cmd += " --image"
    if no_grade:
        cmd += " --no-grade"
    if grade:
        cmd += " --grade"                      # 一発で通った回でも採点する（既定はその回だけ省く・2026-09-11）
    if shake:
        # 揺らし（2026-09-11 夜）: shake="latest" で直前の記録、パスならその記録。only="名前,名前" で絞る
        cmd += " --shake" + ("" if shake == "latest" else " " + shlex.quote(shake))
        if only:
            cmd += " --only " + shlex.quote(only)
        if imagine:
            cmd += f" --imagine {int(imagine)}"      # AI に N 通り書かせた版も足す（1 往復・既定 haiku）
    if score:
        cmd += " --score"
    if runs:
        cmd += " --runs"
    if cases:
        cmd += " --cases"
    if keep_case:
        cmd += " --keep-case " + shlex.quote(keep_case)
    if grade_ai:
        cmd += " --grade-ai " + shlex.quote(grade_ai)
    if grade_model:
        cmd += " --grade-model " + shlex.quote(grade_model)
    if recipe:
        cmd += " --recipe " + shlex.quote(recipe)
    if sheet:
        cmd += " --sheet " + shlex.quote(sheet)
    if mode:
        cmd += " --mode " + shlex.quote(mode)
    if macro:
        cmd += " --macro " + shlex.quote(macro)
    if max_turns and int(max_turns) != 4:
        cmd += f" --max-turns {int(max_turns)}"
    if dry_run:
        cmd += " --dry-run"
    if new_book:
        cmd += " --new-book"
    if ai:
        cmd += " --ai " + shlex.quote(ai)
    if model:
        cmd += " --model " + shlex.quote(model)
    # 往復数ぶんの持ち時間を渡す（900 秒固定だと、手順書の 6 往復＋API の混雑で超えたときに
    # 請求書も報告も届かず、記録ファイルを読むしかなかった。1 往復の最悪は 90 秒×3 回の撃ち直し＋道具の時間）
    turns = max(int(max_turns or 4), 1)
    if not wait:
        jid = _submit_async(cmd)
        if not jid.startswith("job"):
            return jid                       # 関所で断られた（そのまま理由を返す）
        return (f"走らせました（待たずに返しています）: {jid}\n"
                f"  いま何をしているかは agent_status(\"{jid}\") で見てください"
                "（終わっていれば全文が返ります）。\n"
                "  この仕事が終わるまで、他の vba コマンドは順番待ちになります。")
    return _submit(cmd, timeout=min(300 + turns * 300, 3600))


@mcp.tool()
def reload_tools() -> str:
    """道具の .py（vbam_*・vba_manager）を直した後に、MCP サーバーを再起動せずに読み直す（2026-09-06 夜）。

    ワーカーの列に並んで実行するので、走っている agent の途中には割り込まない。依存の順に importlib.reload し、
    コマンド表も作り直す。サーバー本体（vba_mcp_server.py）を変えたときだけは再起動が要る。
    """
    return _submit("__reload__", timeout=120)


@mcp.tool()
def agent_status(job_id: str = "") -> str:
    """agent(wait=False) で走らせた仕事の様子を見る（2026-09-05）。

    ワーカーのキューを通さないので、agent が回っている最中でも答えられる（進み具合のファイルを読むだけ）。
    終わっていれば、そのまま全文（請求書・報告・検査）を返して覚えを捨てる。
    job_id を省くと、いちばん新しい仕事を見る。走らせた仕事が無くても、進み具合だけは読める
    （CLI で回している最中でも「いま何をしているか」が分かる）。
    """
    from vbam_agent import progress_read, progress_line
    rec = progress_read()
    # 書いた側が作った 1 行をそのまま使う（読む側が書式を知らなくていい）。古いファイルは組み立て直す
    now_line = (rec.get('line') or progress_line(rec)) if rec else ""
    job = None
    if job_id:
        job = _live_jobs.get(job_id)
        if job is None:
            return (f"その仕事の覚えがありません: {job_id}\n"
                    + (f"いまの進み具合: {now_line}" if now_line else "（進み具合のファイルもありません）"))
    elif _live_jobs:
        job_id, job = sorted(_live_jobs.items(), key=lambda kv: kv[1]['start'])[-1]
    if job is None:
        return (f"いまの進み具合: {now_line}（{rec.get('time')} 時点）" if now_line
                else "走らせた仕事はありません（進み具合のファイルもありません）")
    waited = time.time() - job['start']
    if not job['done'].is_set():
        return (f"まだ走っています: {job_id}（{waited:.0f} 秒経過）\n"
                + (f"  いま: {now_line}（{rec.get('time')} 時点）" if now_line else "  （進み具合はまだ出ていません）"))
    _live_jobs.pop(job_id, None)
    box = job['box']
    body = _format_result(job['line'], box.get("ok", False),
                          (box.get("out") or "").strip(), (box.get("err") or "").strip())
    return f"終わりました: {job_id}（{waited:.0f} 秒）\n" + body


def _shutdown():
    # ツールが自動起動した Excel があれば畳む（ユーザーの Excel には触れない）
    try:
        box, done = {}, threading.Event()
        with _worker_lock:
            _jobs.put(("__cleanup__", box, done))
        done.wait(10)
    except Exception:
        pass
    for p in (LAST_VALUES, PATCH_TGT, PATCH_REP):      # 常駐ごとの作業ファイルを残さない
        try:
            os.remove(p)
        except OSError:
            pass


if __name__ == "__main__":
    threading.Thread(target=_worker, args=(_jobs,), daemon=True).start()
    atexit.register(_shutdown)
    mcp.run()
