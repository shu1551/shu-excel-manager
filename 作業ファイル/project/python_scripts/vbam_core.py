# -*- coding: utf-8 -*-
"""vbam_core.py — vba_manager 分割パート: COM接続・共通ユーティリティ・ガード（衝突ガード/消滅待ち/共有ヘルパー）

vba_manager.py から機械分割（2026-07-12）。単体で実行せず、vba_manager.py 経由で使う。
"""
import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

# ---- パス定数 ----
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR  = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', 'backups'))
LAST_PROC_FILE = os.path.join(SCRIPT_DIR, '_last_proc.vba')   # get の出力先
XL_EXTS     = ('.xlsm', '.xlam', '.xlsx', '.xls', '.xlsb')


# ================================================================
# ユーティリティ
# ================================================================

# このツールが自動起動した Excel インスタンスの記録。
# batch/shell で「未起動の別ファイル」を複数触ると DispatchEx が複数走るため、
# スカラー1個だと最後の1台しか始末できずゾンビが残る（2026-07-09 点検で発見）。
# 起動した全台を積んで、cleanup_excel で全部閉じる。
_created_instances = []          # [{"xl": <COM>, "pid": int|None}, ...]
# 後方互換の別名（コード内の他参照・デバッグ表示用に「最後の1台」を指す）
_created_xl = None
_created_xl_pid = None
# 直前の _get_workbook_uncached が「このツール自身でブックを開いた」かどうか。
# 健診モード(readonly)の後始末は自分が開いたブックにしか許されない
# （ユーザーが読み取り専用で開いていたブックを閉じる事故の防止・2026-07-14）
_last_open_by_tool = False
# 直前に掴んだ Excel の PID（2026-09-16）。_any_excel_process は tasklist（実測 36〜52 ミリ秒）を撃つ前に
# この PID がまだ EXCEL.EXE なら「いる」と即答する。常駐の MCP は 5 秒無操作で接続を手放すので、
# 会話からの呼び出しはほぼ毎回ここを通る＝1 呼び出しあたり数十ミリ秒が消える
_last_excel_pid = None
# 直前のコマンドが掴んだブックの名前（呼び出し台帳 _calls.jsonl に書く。コマンドの入口で None に戻す）
_last_book_name = None


def setup_encoding():
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')   # --json 時の情報行が stderr に行くため同様に固定
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    pythoncom.CoInitialize()

    # --json オプションが引数に含まれている場合、通常の print を標準エラー出力にリダイレクトする
    if "--json" in sys.argv:
        import builtins
        _orig_print = builtins.print
        def custom_print(*args, **kwargs):
            if 'file' not in kwargs:
                kwargs['file'] = sys.stderr
            _orig_print(*args, **kwargs)
        builtins.print = custom_print


def looks_like_xl_file(s):
    """文字列がExcelファイルパスっぽいか判定。

    Excel拡張子で終わるものだけを対象ブックと見なす。
    以前は「パス区切りを含む」「ファイルとして存在する」でも True にしていたが、
    それだと `replace-procedure fix.vba` の fix.vba が対象ブック扱いになり、
    無関係ファイルのオープンや古い _last_proc.vba への静かなフォールバックを招く。
    """
    if not s:
        return False
    return s.lower().endswith(XL_EXTS)


def smart_path_resolve(filename):
    """ファイルパスを柔軟に解決"""
    if not filename:
        return None
    if os.path.exists(filename):
        return os.path.abspath(filename)
    for d in [os.getcwd(), SCRIPT_DIR,
              os.path.join(SCRIPT_DIR, '..', '..'),
              os.path.join(SCRIPT_DIR, '..', '..', '..')]:
        c = os.path.join(os.path.abspath(d), filename)
        if os.path.exists(c):
            # 2026-09-19: 「作業ファイル/test.xlsx」のように / で渡されると、/ が混ざったままのパスを返していた。
            # Excel の FullName は \ なので、open が「開いたのに 30 秒待って失敗」と誤報した（Gemini の実射の記録）
            return os.path.normpath(c)
    return None


def note_if_macro_free_book(wb):
    """マクロを入れたブックが .xlsx などマクロを持てない形式なら、1 行で知らせる（保存してもファイルには残らない）。

    2026-09-19: Gemini も Claude も test_pivot.xlsx にマクロを足し、道具は「保存しました」と言ったが、
    .xlsx はマクロを持てない＝開いているあいだ動くだけで、閉じると消える。人が後で「無い」と気づく形だった。"""
    try:
        name = str(wb.Name)
    except Exception:
        return False
    if name.lower().endswith(('.xlsx', '.xltx', '.csv')):
        print(f"⚠ {name} はマクロを持てない形式です。保存しても、いま入れたマクロはファイルに残りません"
              "（開いているあいだは動きます）。残すには .xlsm で保存し直してください。")
        return True
    return False


def same_path(a, b):
    """2 つのパスが同じファイルを指すか（/ と \\・大文字小文字・.. の違いをならして比べる）。"""
    if not a or not b:
        return False
    try:
        return os.path.normcase(os.path.normpath(os.path.abspath(a))) == os.path.normcase(os.path.normpath(os.path.abspath(b)))
    except Exception:
        return str(a).lower() == str(b).lower()


def parse_target_and_rest(posargs):
    """
    posargs の先頭が Excel ファイルなら target_file に、残りを rest に返す。
    そうでなければ target_file=None で全部 rest に返す。
    """
    if posargs and looks_like_xl_file(posargs[0]):
        return posargs[0], list(posargs[1:])
    return None, list(posargs)


# Excel が「今は応答できない」ときに返す HRESULT（参照が死んだわけではない）。
# マクロ実行中・モーダルダイアログ表示中・セル編集中は普通にこれが返る。
_COM_BUSY_HRESULTS = (
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
    -2147417847,  # 0x80010109 RPC_E_SERVERCALL_REJECTED
    -2147417851,  # 0x80010105 RPC_E_SERVERFAULT
    -2146777998,  # 0x800AC472 VBA_E_IGNORE（Excel がセル編集中・モーダル表示中に返す
                  #             一番ありふれた「今は無理」。これを外すと、ビジーな
                  #             Excel を「死んでいる」と誤判定して強制終了し、
                  #             未保存の変更を捨てる＝温存の約束を破る）
)


def _com_is_busy(ex):
    """COM 例外が「Excel がビジー（＝生きているが今は応答できない）」かどうか。

    判定できない例外（COM 以外・HRESULT が読めない）は False を返すが、
    呼び出し側はそれを「死んでいる」と決めつけずに扱うこと。
    """
    for val in (getattr(ex, 'hresult', None),) + tuple(getattr(ex, 'args', ()) or ()):
        if isinstance(val, int) and val in _COM_BUSY_HRESULTS:
            return True
    return False


def _pid_is_excel(pid):
    """その PID が今も EXCEL.EXE かを確認する（強制終了の前の身元確認）。

    Windows は終了したプロセスの PID を再利用する。参照が死んだ＝プロセスは
    もう無い、という状況で PID を撃つと、その番号を引き継いだ無関係のプロセスを
    殺しうる。確認できないときは False（撃たない）を返す。
    """
    if pid is None:
        return False
    try:
        import win32api
        import win32con
        import win32process
        h = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        try:
            name = win32process.GetModuleFileNameEx(h, 0)
        finally:
            win32api.CloseHandle(h)
        return os.path.basename(name).lower() == 'excel.exe'
    except Exception:
        return False


def _remember_excel_pid(xl):
    """掴んだ Excel の PID を控える（次の _any_excel_process が tasklist を撃たずに済む）。失敗は黙って無視。"""
    global _last_excel_pid
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(int(xl.Hwnd))
        _last_excel_pid = pid or None
    except Exception:
        pass


def _any_excel_process():
    """EXCEL.EXE が1つでも動いているか（COM を一切使わない・PID の身元確認→tasklist）。

    Excel が本当に未起動のとき、ROT 走査や GetActiveObject などの COM 呼び出しを
    試みると応答待ちで長く戻らないことがある（agent --fire が Excel 未起動時に
    止まる・2026-09-03 記録）。ここで False と分かればその前に即エラーにできる。
    2026-09-16: 前回掴んだ Excel の PID がまだ EXCEL.EXE なら tasklist を撃たずに True
    （1 ミリ秒未満）。PID が死んでいる／控えが無いときだけ従来どおり tasklist。
    tasklist 自体が失敗したときは「いるかもしれない」で True を返し、
    従来どおり COM 側の判定に委ねる（安全側に倒す＝誤って早期エラーにしない）。
    """
    if _pid_is_excel(_last_excel_pid):
        return True
    try:
        import subprocess
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                             capture_output=True, encoding="cp932", errors="replace", timeout=10)
        return "EXCEL.EXE" in (out.stdout or "")
    except Exception:
        return True


def release_created_instances(only_saved=True):
    """自動起動した Excel インスタンスのうち、後始末してよいものを閉じて終了する。

    常駐プロセス（MCPサーバー）では終了時の cleanup_excel が何時間も来ないため、
    パス指定で自動起動した非表示 Excel がセッション中ずっと残留する。COM 起動の
    Excel はアドイン・PERSONAL.XLSB を読まないため、残留中にユーザーが開いた
    ファイルがそこへ合流すると「アドインが効かない・マクロが効かない」状態になる
    （2026-07-12 特定。6/13 ゾンビ事故と同症状。当日実測で非表示3体が残留していた）。
    MCP のアイドル解放（5秒）からこれを呼んで畳む。

    only_saved=True では未保存の変更を持つインスタンスは温存する
    （アイドルのたびに無言で変更破棄をしないため。それらは終了時の
    cleanup_excel が警告つきで始末する）。戻り値: (閉じた台数, 温存した台数)
    """
    global _created_instances, _created_xl, _created_xl_pid
    import gc
    import os
    import signal
    keep = []
    closed = 0
    for inst in _created_instances:
        xl = inst.get("xl")
        pid = inst.get("pid")
        alive = True
        busy = False
        dirty = False
        try:
            wbs = xl.Workbooks
            for i in range(wbs.Count, 0, -1):
                try:
                    if not wbs.Item(i).Saved:
                        dirty = True
                        break
                except Exception as ex2:
                    # Saved の取得が COM のビジー拒否で弾かれることがある
                    # （セル編集中・モーダル表示中に普通に起きる）。ここで握って
                    # 素通りさせると「未保存なし」と誤断し、下で Close(SaveChanges=False)
                    # → Quit → kill まで進んで、未保存の変更を無言で捨てる。
                    # 判定できないときは安全side（温存）に倒す。
                    if _com_is_busy(ex2):
                        busy = True
                    else:
                        dirty = True
                    break
        except Exception as ex:
            if _com_is_busy(ex):
                # マクロ実行中・モーダル表示中。生きているし、未保存かどうかも
                # 今は判定できない。ここで畳むと only_saved の約束（未保存は温存）を
                # 破って変更を捨てるので、触らず次のアイドルに回す
                busy = True
            else:
                alive = False  # 参照が本当に死んでいる（手動で閉じられた等）
        if busy:
            keep.append(inst)
            continue
        if alive and only_saved and dirty:
            keep.append(inst)
            continue
        if alive:
            try:
                wbs = xl.Workbooks
                for i in range(wbs.Count, 0, -1):
                    try:
                        wbs.Item(i).Close(SaveChanges=False)
                    except Exception:
                        pass
                xl.Quit()
            except Exception:
                pass
        # Quit がゾンビ化しても残さない（cleanup_excel と同じ最終手段）。ただし
        # PID の身元を確認してから撃つ（再利用された PID の巻き添えを防ぐ）
        if _pid_is_excel(pid):
            try:
                os.kill(pid, signal.SIGTERM)
            except Exception:
                pass
        closed += 1
    _created_instances = keep
    if not keep:
        _created_xl = None
        _created_xl_pid = None
    gc.collect()
    return closed, len(keep)


def release_instance(pid):
    """自動起動した Excel のうち、この PID の 1 台だけを畳む（ほかの台には触らない）。→ 畳んだ台数。

    2026-09-17: 鍛える回路・先撃ちの試し撃ちが、自分の台を畳むのに release_created_instances を呼んでいたので、
    試験（exam）や入口の確かめのように「道具が起こした Excel の中で agent を回す」と、その Excel まで畳んで
    「RPC サーバーを利用できません」で止まった。"""
    global _created_instances, _created_xl, _created_xl_pid
    if not pid:
        return 0
    mine = [i for i in _created_instances if i.get("pid") == pid]
    others = [i for i in _created_instances if i.get("pid") != pid]
    if not mine:
        return 0
    last_xl, last_pid = _created_xl, _created_xl_pid
    _created_instances = mine
    try:
        closed, _kept = release_created_instances(only_saved=False)
    finally:
        _created_instances = others + _created_instances
        if last_pid == pid:
            _created_xl = others[-1]["xl"] if others else None
            _created_xl_pid = others[-1]["pid"] if others else None
        else:
            _created_xl, _created_xl_pid = last_xl, last_pid
    return closed


def cleanup_excel():
    """新規起動されたExcelインスタンスがあれば（複数でも）全て終了し、COMを初期化解除する"""
    global _created_instances, _created_xl, _created_xl_pid
    import gc
    import os
    import signal
    # 自動起動したインスタンスが無ければ何もしない回なので、DEBUG も出さない
    # （全コマンドの末尾に毎回2行のノイズが出ていた）
    if _created_instances:
        print(f"[DEBUG] cleanup_excel called. instances: {len(_created_instances)}")

    # Python側のCOM参照を解放するためにGCを強制実行
    gc.collect()

    for inst in _created_instances:
        xl = inst.get("xl")
        pid = inst.get("pid")
        if xl is not None:
            try:
                print("[DEBUG] Closing open workbooks...")
                try:
                    # 切断エラーを回避しつつ、逆順にブックを閉じる
                    wbs = xl.Workbooks
                    for i in range(wbs.Count, 0, -1):
                        try:
                            wb = wbs.Item(i)
                            name = wb.Name
                            try:
                                if not wb.Saved:
                                    # 自動起動経路では閉じる＝未保存変更の破棄。無言だと
                                    # 「書いたつもりが消えていた」になるため明示する
                                    print(f"⚠ 未保存の変更を破棄して閉じます: {name}")
                                    print("  （保存したい場合は同じ batch 内で save を実行するか、"
                                          "Excel を先に起動してから作業してください）")
                            except Exception:
                                pass
                            wb.Close(SaveChanges=False)
                            print(f"[DEBUG] Workbook closed: {name}")
                        except Exception as ex:
                            print(f"[DEBUG] Failed to close wb {i}: {ex}")
                except Exception as ex:
                    print(f"[DEBUG] Failed to access Workbooks: {ex}")

                print("[DEBUG] Calling xl.Quit()...")
                xl.Quit()
                print("[DEBUG] xl.Quit() completed.")
            except Exception as ex:
                print(f"[DEBUG] Error during Excel cleanup: {ex}")

        # 新規起動したPIDが存在する場合は強制クリーンアップ。
        # 撃つ前に「その PID が今も EXCEL.EXE か」を確認する（既に終了していた場合、
        # Windows は PID を再利用するため無関係のプロセスを殺しうる）
        if pid is not None:
            if not _pid_is_excel(pid):
                print(f"[DEBUG] PID {pid} は既に EXCEL.EXE ではありません（終了済み）。強制終了はしません。")
            else:
                try:
                    print(f"[DEBUG] Force-killing Excel process (PID: {pid})...")
                    os.kill(pid, signal.SIGTERM)
                    print("[DEBUG] Excel process force-killed successfully.")
                except Exception as ex:
                    print(f"[DEBUG] Excel process force-kill failed or already exited: {ex}")

    _created_instances = []
    _created_xl = None
    _created_xl_pid = None

    # 最後の解放
    gc.collect()
    try:
        pythoncom.CoUninitialize()
    except Exception as ex:
        print(f"[DEBUG] CoUninitialize failed: {ex}")


def load_excel_addins_and_personal(xl):
    """新規起動されたExcelにアドインとPERSONAL.XLSBをロードする"""
    import os
    import time

    # 警告を非表示にする
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    loaded_any = False

    # 1. アドインのロード
    try:
        for addin in xl.AddIns:
            if addin.Installed:
                try:
                    # すでに開いていなければ開く
                    opened = False
                    for wb in xl.Workbooks:
                        if wb.Name.lower() == os.path.basename(addin.FullName).lower():
                            opened = True
                            break
                    if not opened:
                        xl.Workbooks.Open(addin.FullName)
                        print(f"[DEBUG] Loaded addin: {addin.Name}")
                        loaded_any = True
                except Exception as ex:
                    # Excelがビジー状態の時のリトライ (0x800ac472)
                    if "800ac472" in str(ex):
                        time.sleep(0.5)
                        try:
                            xl.Workbooks.Open(addin.FullName)
                            print(f"[DEBUG] Loaded addin (retry): {addin.Name}")
                            loaded_any = True
                        except Exception as ex2:
                            print(f"[DEBUG] Failed to load addin {addin.Name} after retry: {ex2}")
                    else:
                        print(f"[DEBUG] Failed to load addin {addin.Name}: {ex}")
    except Exception as ex:
        print(f"[DEBUG] Failed to access AddIns: {ex}")

    # 2. PERSONAL.XLSB のロード
    try:
        startup_path = xl.StartupPath
        if startup_path:
            for fname in ["PERSONAL.XLSB", "personal.xlsb", "PERSONAL.XLS", "personal.xls"]:
                p_path = os.path.join(startup_path, fname)
                if os.path.exists(p_path):
                    opened = False
                    for wb in xl.Workbooks:
                        if wb.Name.lower() == fname.lower():
                            opened = True
                            break
                    if not opened:
                        try:
                            xl.Workbooks.Open(p_path)
                            print(f"[DEBUG] Loaded personal macro book: {fname}")
                            loaded_any = True
                        except Exception as ex:
                            if "800ac472" in str(ex):
                                time.sleep(0.5)
                                try:
                                    xl.Workbooks.Open(p_path)
                                    print(f"[DEBUG] Loaded personal macro book (retry): {fname}")
                                    loaded_any = True
                                except Exception as ex2:
                                    print(f"[DEBUG] Failed to load personal macro book {fname} after retry: {ex2}")
                            else:
                                print(f"[DEBUG] Failed to load personal macro book {fname}: {ex}")
                    break
    except Exception as ex:
        print(f"[DEBUG] Failed to load personal macro book: {ex}")

    # ロード処理が走った場合は待機
    if loaded_any:
        time.sleep(1.0)

    try:
        xl.DisplayAlerts = True
    except Exception:
        pass


def _get_active_excel():
    """起動中の Excel.Application に late-binding で接続する(gencache 非経由)。

    win32com.client.GetActiveObject は gen_py キャッシュを通るため、キャッシュ破損
    (例: module '...' has no attribute 'CLSIDToClassMap')で例外になり、Excel が実際は
    開いているのに「起動していない」と誤判定する。pythoncom.GetActiveObject で生の
    インスタンスを掴んで dynamic.Dispatch で包めばキャッシュに一切依存しない。
    Excel 未起動時は例外を送出する(呼び出し側で捕捉する)。
    """
    clsid = pywintypes.IID("Excel.Application")
    unk = pythoncom.GetActiveObject(clsid)
    disp = unk.QueryInterface(pythoncom.IID_IDispatch)
    return win32com.client.dynamic.Dispatch(disp)


def _running_excel_workbooks():
    """Running Object Table を走査し、起動中の全 Excel の全ブックを返す。

    GetActiveObject は ROT 先頭の1インスタンスしか返さず、非表示ゾンビ Excel
    (ブック0個)を掴むことがある。ROT を直接舐めれば、実際にブックを開いている
    インスタンスだけ拾える(ゾンビはブックを持たないので現れない)。
    失敗時は [] を返し、呼び出し側は GetActiveObject にフォールバックする。
    """
    wbs = []
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        monikers = list(rot)
    except Exception:
        return wbs
    for mon in monikers:
        try:
            disp = mon.GetDisplayName(ctx, None)
        except Exception:
            continue
        # Excel のブックはフルパス(...\xxx.xlsm 等)で ROT 登録される
        if not disp or not re.search(r'\.xl\w{1,4}$', disp, re.IGNORECASE):
            continue
        try:
            obj = rot.GetObject(mon)
            wb = win32com.client.dynamic.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            _ = wb.FullName  # ブックか確認(違えば例外)
            wbs.append(wb)
        except Exception:
            continue
    return wbs


# get_workbook の接続キャッシュ。通常の1コマンド1プロセスでは1回しか呼ばれないので
# 実質無関係だが、batch モードでは全行が同じCOM接続を使い回せる（1コマンド毎の
# COM再接続＝ROT走査が一番重い、という実測への構造的な解）。
_wb_cache = {}

# 固い固定（pinned_workbook のあいだだけ True）。立っているあいだは「アクティブブックの切替に追従」を止める。
# 道具が自分で起こしたブック（練習台・別ブックの組み立て）を触っている最中に、人が Excel 側で別のブックを
# 選んでも流れないため（2026-09-08）。
_wb_pin_hard = False


def pinned_workbook(wb):
    """with pinned_workbook(wb): のあいだ、道具の「アクティブブック」をこの wb に固く固定する。

    cmd_* 関数を道具の内側から呼ぶとき（設計図の組み立て・実射の練習台）に使う。
    これを付けないと、対象ファイルを名指ししない Namespace で呼んだ cmd_* が
    get_workbook(None) の「アクティブブック自動検出」に落ち、**人が開いているブックを掴む**
    （2026-09-08・--fire の 1 本目が人のブックで tidy を撃とうとして落ちた。
    シート名がたまたま一致していれば、人のブックの書式を書き換えていた）。

    抜けるときは前の状態に戻す＝常駐の MCP で固定が残らない。
    """
    import contextlib

    @contextlib.contextmanager
    def _ctx():
        global _wb_pin_hard
        prev_entry = _wb_cache.get("__active__")
        prev_hard = _wb_pin_hard
        pin_active_workbook(wb)
        _wb_pin_hard = True
        try:
            yield
        finally:
            _wb_pin_hard = prev_hard
            if prev_entry is None:
                _wb_cache.pop("__active__", None)
            else:
                _wb_cache["__active__"] = prev_entry

    return _ctx()


def _open_maybe_repair(xl, target_path, args=()):
    """Workbooks.Open(target_path, *args)。失敗したら CorruptLoad=xlRepairFile（15番目の引数）で開き直す。

    Google スプレッドシートの書き出し xlsx 等、中身は正常（openpyxl は読める）なのに素の Open が
    COM から弾かれるファイルがある。Excel の GUI で「修復して開く」を選ぶのと同じ動きを自動でする
    （2026-09-05・本物のブックの実射で発見。エラーは HRESULT -2146827284）。
    """
    try:
        return xl.Workbooks.Open(target_path, *args)
    except Exception:
        full = [target_path] + list(args) + [None] * (15 - 1 - len(args))
        full[-1] = 1                                    # CorruptLoad=xlRepairFile
        wb = xl.Workbooks.Open(*full)
        print("（素の Open が失敗したため、修復モード（CorruptLoad=xlRepairFile）で開き直しました）")
        return wb


def _screen_book(app):
    """画面に出ている窓のブック（ActiveWindow.Parent）。取れなければ ActiveWorkbook。

    2026-09-13: アドインの更新登録（.xlsm を .xlam に SaveAs）の後、ActiveWorkbook が画面から消えた
    秀コンボ.xlsm を指したまま残り、お試し版のテスト用1 を見ているのに 秀コンボ.xlsm の目次 を対象にした。
    人が見ているのは窓なので、窓のブックを先に取る。
    """
    try:
        w = app.ActiveWindow
        if w is not None and bool(w.Visible):
            return w.Parent
    except Exception:
        pass
    return app.ActiveWorkbook


def get_workbook(target_file_arg=None, load_addins=False, readonly=False):
    """get_workbook（接続キャッシュつきの入口）。戻り値: (xl, wb)

    readonly=True は診断系コマンド用。未起動ブックを自動で開く場合に
    「読み取り専用＋イベント無効」で開く（Workbook_Open 等を起こさない健診モード）。
    既に開いているブックには影響しない。
    """
    key = "__active__"
    if target_file_arg:
        resolved = smart_path_resolve(target_file_arg)
        key = resolved.lower() if resolved else target_file_arg.lower()
    if key in _wb_cache:
        xl, wb, was_ro = _wb_cache[key]
        try:
            _ = wb.Name          # 生存確認（閉じられていたら再接続）
            if key == "__active__" and not _wb_pin_hard:
                # shell/batch/MCP 等の1接続セッション中にユーザーが Excel 側で
                # 別ブックをアクティブにした場合、キャッシュした旧ブックのまま
                # 破壊コマンドが走ると対象取り違えになる。毎回同一性を確認して追従する
                # （pinned_workbook のあいだは追従しない＝道具が自分の台を触っている最中）
                try:
                    cur = _screen_book(xl)
                    if cur is not None and cur.Name != wb.Name:
                        print(f"対象ブック: {cur.Name}（アクティブブックの切替に追従）")
                        wb = cur
                        _wb_cache[key] = (xl, wb, was_ro)
                except Exception:
                    pass
            if was_ro and not readonly:
                # 健診モード（読み取り専用・イベント無効）で「このツールが」開いた
                # ブックを、同じセッション（batch/shell/MCP は接続キャッシュを持ち越す）で
                # 書き込み系に渡すと、編集が分かりにくい COM エラーで失敗する。
                # 例: get <path> → replace-procedure <path>
                # 警告だけでは直らないので、書き込みが要るときは通常モードで開き直す。
                # ※ was_ro はツール自身が開いたときにしか立たない（_last_open_by_tool）。
                #   ユーザーが読み取り専用で開いていたブックをここで閉じてはならない。
                print("（診断用に読み取り専用で開いていたブックを、書き込みのため通常モードで開き直します）")
                try:
                    wb.Close(SaveChanges=False)
                except Exception as ex:
                    # 閉じられない（ビジー・モーダル表示中等）のに開き直したことにすると、
                    # 読み取り専用のまま書き込みに進み、Save で落ちるか黙って捨てられる
                    print(f"エラー: 読み取り専用で開いたブックを閉じられませんでした: {ex}")
                    print("  Excel がビジー（マクロ実行中・ダイアログ表示中）の可能性があります。")
                    print("  少し待ってから再実行してください。")
                    raise ReopenFailed(str(ex))
                try:
                    xl.EnableEvents = True   # 健診モードで切っていたイベントを戻す
                except Exception as ex:
                    print(f"⚠ EnableEvents を戻せませんでした: {ex}", file=sys.stderr)
                    print("  このExcelではイベントマクロ(Worksheet_Change等)が動きません。",
                          file=sys.stderr)
                del _wb_cache[key]
                xl, wb = _get_workbook_uncached(target_file_arg, load_addins, False)
                # 本当に書き込み可能になったかを実測する（「開き直しました」と言いながら
                # 読み取り専用のまま、を防ぐ）
                try:
                    if bool(wb.ReadOnly):
                        raise ReopenFailed(
                            "開き直しましたが、まだ読み取り専用です"
                            "（ファイルが読み取り専用属性、または他で開かれています）")
                except pywintypes.com_error:
                    pass
                _wb_cache[key] = (xl, wb, False)
                return xl, wb
            if load_addins:
                load_excel_addins_and_personal(xl)
            _note_book(wb)
            return xl, wb
        except ReopenFailed:
            # 「読み取り専用のまま書き込みに進む」のを防ぐために上げた例外。
            # ここで握ると下の _get_workbook_uncached が既に開いている（＝閉じ損ねた／
            # 読み取り専用のままの）ブックを見つけて返してしまい、防ごうとした事故が
            # そのまま起きる。キャッシュだけ捨てて、失敗は呼び出し元へ伝える。
            _wb_cache.pop(key, None)
            raise
        except Exception:
            _wb_cache.pop(key, None)   # 開き直しの途中で既に消していることがある（二重 del の KeyError 防止）
    xl, wb = _get_workbook_uncached(target_file_arg, load_addins, readonly)
    # 健診モードの印は「このツールが今このブックを開いた場合」だけ立てる。
    # wb.ReadOnly だけで判定すると、ユーザーが読み取り専用で開いていたブック
    # （読み取り専用属性・読み取り専用推奨・他者がロック中・「読み取り専用で開く」を
    # 選択）を健診モードと誤認し、後で書き込み系が来たときに勝手に Close して
    # 開き直す＝ユーザーのウィンドウが消える（2026-07-14 実機で再現）。
    opened_ro = readonly and target_file_arg is not None and _last_open_by_tool
    try:
        opened_ro = opened_ro and bool(wb.ReadOnly)
    except Exception:
        pass
    _wb_cache[key] = (xl, wb, opened_ro)
    _note_book(wb)
    return xl, wb


def _note_book(wb):
    """このコマンドが掴んだブックの名前を控える（呼び出し台帳用・2026-09-16）。読めなければ何もしない。"""
    global _last_book_name
    try:
        _last_book_name = str(wb.Name)
    except Exception:
        pass


def _get_workbook_uncached(target_file_arg=None, load_addins=False, readonly=False):
    """
    target_file_arg が None/空 → アクティブExcelブックを自動使用
    それ以外 → 既に開いているか確認、なければ新規オープン
    戻り値: (xl, wb)

    副作用: グローバル _last_open_by_tool に「このツールが今開いたのか
    （True）／既に開いていたブックを掴んだだけか（False）」を記録する。
    健診モード(readonly)の後始末は自分が開いたブックにしか許されないため
    （ユーザーが読み取り専用で開いていたブックを閉じる事故の防止）。
    """
    global _last_open_by_tool
    _last_open_by_tool = False
    pythoncom.CoInitialize()

    if not target_file_arg:
        # Excel が1つも起動していないなら、COM に触る前にここで打ち切る
        # （ROT 走査や GetActiveObject の応答待ちで止まるのを避ける）
        if not _any_excel_process():
            raise Exception(
                "Excel が起動していません。\n"
                "  ・Excel で対象ブックを開いてから再実行してください。")
        # ROT 全インスタンス横断で、実ブックを持つ Excel を選ぶ(ゾンビ自動回避)
        wb = None
        for cand in _running_excel_workbooks():
            try:
                app = cand.Application
                if app.Visible and app.ActiveWorkbook is not None:
                    wb = _screen_book(app)    # 可視インスタンスの「画面に出ている窓」のブックを最優先（2026-09-13）
                    break
                if wb is None:
                    try:
                        if cand.IsAddin or not cand.Windows(1).Visible:
                            continue           # アドイン／非表示ブック(PERSONAL.XLSB 等)は暫定候補にしない
                    except Exception:
                        pass
                    wb = cand                  # 暫定: 実ブックを持つ最初のもの
            except Exception:
                continue
        if wb is None:
            # ROT に出ない稀ケース(未保存ブック等)は GetActiveObject で再挑戦
            try:
                wb = _get_active_excel().ActiveWorkbook
            except Exception:
                wb = None
        if wb is None:
            raise Exception(
                "起動中の Excel に開いているブックが見つかりません。\n"
                "  ・Excel で対象ブックを開いてから再実行してください。\n"
                "  ・非表示のゾンビ EXCEL.EXE が残っている場合があります。"
                "タスクマネージャーで余分な EXCEL.EXE を終了し、対象ブックを開いて再実行してください。\n"
                "  ※ COM 接続できないからといって .bas を手書きスクリプトで処理しないこと"
                "(改行二重化の原因)。")
        xl = wb.Application
        _remember_excel_pid(xl)
        print(f"対象ブック: {wb.Name}  (アクティブブック自動検出)")
        if load_addins:
            load_excel_addins_and_personal(xl)
        return xl, wb

    target_path = smart_path_resolve(target_file_arg)
    if not target_path:
        raise Exception(f"ファイルが見つかりません: {target_file_arg}")

    # 既に開いているか確認(全 Excel インスタンス横断)
    excel_running = False
    for wb in _running_excel_workbooks():
        excel_running = True
        try:
            if same_path(wb.FullName, target_path):
                xl = wb.Application
                _remember_excel_pid(xl)
                print(f"対象ブック: {wb.Name}  (既に開いています)")
                if load_addins:
                    load_excel_addins_and_personal(xl)
                return xl, wb
        except Exception:
            continue
    if not excel_running:
        try:
            xl_fallback = _get_active_excel()
            excel_running = True
            # ROT 走査が失敗していても実際は開いているケースの二重オープン防止:
            # GetActiveObject で掴んだインスタンスの Workbooks も確認する
            try:
                for wb in xl_fallback.Workbooks:
                    if same_path(wb.FullName, target_path):
                        _remember_excel_pid(xl_fallback)
                        print(f"対象ブック: {wb.Name}  (既に開いています)")
                        if load_addins:
                            load_excel_addins_and_personal(xl_fallback)
                        return xl_fallback, wb
            except Exception:
                pass
        except Exception:
            excel_running = False

    # 新規オープン
    # ★必ず DispatchEx を使う。Dispatch は既存インスタンスがあるとそこに接続してしまい、
    #   「自分が起動した Excel」と誤認 → 後始末でユーザーの Excel ごと閉じる大事故になる
    #   （2026-07-03 実害。ユーザーのブックを巻き込んで Quit した）
    xl = win32com.client.DispatchEx("Excel.Application")
    global _created_xl, _created_xl_pid
    _created_xl = xl
    try:
        import win32process
        hwnd = xl.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        _created_xl_pid = pid
        global _last_excel_pid
        _last_excel_pid = pid or None
        print(f"[DEBUG] Excel process PID detected: {_created_xl_pid}")
    except Exception as ex:
        _created_xl_pid = None
        print(f"[DEBUG] Failed to detect Excel PID: {ex}")
    # 起動した全台を記録（batch/shell で複数起動しても cleanup で全部閉じるため）
    _created_instances.append({"xl": xl, "pid": _created_xl_pid})
    xl.Visible = "--visible" in sys.argv or "-v" in sys.argv

    if load_addins:
        load_excel_addins_and_personal(xl)

    if readonly:
        # 健診モード: Workbook_Open 等の自動実行を起こさず、リンク更新もせず、
        # 読み取り専用で開く（「診察に行ったら患者を起こしてしまった」の防止）
        try:
            xl.EnableEvents = False
        except Exception:
            pass
        wb = _open_maybe_repair(xl, target_path, (0, True))   # UpdateLinks=0, ReadOnly=True
        _last_open_by_tool = True
        print(f"対象ブック: {wb.Name}  (新規オープン・読み取り専用・イベント無効=健診モード)")
        return xl, wb

    wb = _open_maybe_repair(xl, target_path)
    _last_open_by_tool = True
    print(f"対象ブック: {wb.Name}  (新規オープン)")
    if not excel_running and not load_addins:
        # COM起動のExcelは起動処理が走らず、アドインや PERSONAL.XLSB が読み込まれない
        print("注意: Excelが未起動だったため自動化用に新規起動しました。")
        print("      このExcelにはアドイン(秀.xlam等)・PERSONAL.XLSB が読み込まれていません。")
        print("      普段使いにはこのウィンドウを閉じて、手動起動した Excel で開き直してください。")
    return xl, wb


def pin_active_workbook(wb):
    """このプロセスの「アクティブブック」をこのブックに固定する（未保存のブック用）。

    2026-09-06 深夜: 実射が自分で起こした Excel に練習台（未保存の Book1）を作ったところ、
    中のコマンドが get_workbook(None) で「アクティブブック」を探し、GetActiveObject が
    人の空の Excel を返して 45 本すべて「開いているブックが見つかりません」で落ちた。
    未保存のブックはフルパスで名指しできないので、ここで直に据える。
    """
    if wb is None:
        _wb_cache.pop("__active__", None)
        return
    try:
        _wb_cache["__active__"] = (wb.Application, wb, False)
    except Exception:
        pass


def get_or_start_excel(visible=True):
    """実射・練習台のための Excel を得る（開いているブックが無くても止まらない）。

    実射（agent --fire）は自分で Workbooks.Add して練習台を作るので、
    「人が開いているブック」は要らない。それなのに get_workbook を通していたため、
    Excel を開いていない夜間の撃ち込みが「開いているブックが見つかりません」で
    1 本も撃てずに終わっていた（2026-09-06 深夜）。

    開いているブックがあればその Excel を使い（人の Excel を勝手に増やさない）、
    無ければ DispatchEx で自分の台を起こす。起こした台は _created_instances に
    記録されるので release_created_instances で畳まれる。
    既定は可視（撃っている様子が見えるほうを採る）。
    """
    global _created_xl, _created_xl_pid
    pythoncom.CoInitialize()
    # 人の Excel は借りない（2026-09-06 深夜）。借りた実射が練習台を人の窓に出し入れし、
    # 終わったあと人の Excel がブック 0 冊の真っ黒な窓になった＝「画面が真っ黒になってるんですが」。
    # 実射は自分の台で完結させる。
    xl = win32com.client.DispatchEx("Excel.Application")
    _created_xl = xl
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
        _created_xl_pid = pid
    except Exception:
        _created_xl_pid = None
    _created_instances.append({"xl": xl, "pid": _created_xl_pid})
    try:
        xl.Visible = bool(visible)
        xl.DisplayAlerts = False
    except Exception:
        pass
    print("実射の Excel: 自分の台を起こしました（人の Excel には触りません。終わったら畳みます）")
    return xl


def make_backup(wb_fullname, label):
    """バックアップを作成（タイムスタンプ付き・同系列は直近5世代まで保持）。

    成功時はバックアップパス、失敗時は None を返す。
    破壊的操作（replace-procedure / replace-module）の呼び出し元は
    None のとき停止する（--force で続行可）。
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ext = os.path.splitext(wb_fullname)[1] or '.xlsm'
    stamp = time.strftime("%Y%m%d_%H%M%S")
    # フォルダ違いの同名ブックが同じ系列に混ざると、世代の間引きが
    # もう一方のブックのバックアップを消してしまう。フォルダの短いタグで系列を分ける
    dirtag = format(zlib.crc32(os.path.dirname(os.path.abspath(wb_fullname))
                               .lower().encode('utf-8')) & 0xffff, '04x')
    prefix = os.path.basename(wb_fullname) + f".backup_before_{label}_{dirtag}_"
    backup_path = os.path.join(BACKUP_DIR, prefix + stamp + ext)
    # 同一秒内の連続バックアップ（batch での連続置換等）を黙って上書きしない
    seq = 1
    while os.path.exists(backup_path):
        seq += 1
        backup_path = os.path.join(BACKUP_DIR, f"{prefix}{stamp}_{seq}{ext}")
    try:
        shutil.copy2(wb_fullname, backup_path)
        print(f"バックアップ作成: backups/{os.path.basename(backup_path)}")
    except Exception as e:
        print(f"警告: バックアップ失敗 ({e})")
        return None
    # 同じ系列（同ブック・同ラベル）の古い世代を間引いて5世代までにする
    try:
        olds = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix))
        for f in olds[:-5]:
            os.remove(os.path.join(BACKUP_DIR, f))
    except Exception:
        pass
    return backup_path


def _remove_export_artifacts(bas_path):
    """一時 Export の後始末。

    フォーム（UserForm）を Export すると同名の .frx が必ず併産されるが、
    従来は .bas しか消しておらず _tmp_*.frx が溜まり続けていた（構造的な掃除漏れ）。
    ペアで削除する。
    """
    for p in (bas_path, os.path.splitext(bas_path)[0] + '.frx'):
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass


def read_code_file(path):
    """コードファイルを UTF-8 / CP932 で読み込む（改行は \\n に正規化）。

    改行二重化(\\r\\r\\n)の水際検知＝多層防御の①層。テキストモード読みでは
    \\r\\r\\n が \\n\\n（空行）に化けて検知不能になるため、バイトで検知してから復号する。
    外部エディタ等で二重化済みのファイルもここで畳まれ、素通しでモジュールに入らない。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    if b'\r\r' in raw:
        print(f"⚠ {os.path.basename(path)} に改行の二重化(\\r\\r\\n)を検知しました。正規化して読み込みます。")
    for enc in ['utf-8-sig', 'utf-8', 'cp932']:
        try:
            text = raw.decode(enc)
        except Exception:
            continue
        # \r\n / \r\r\n / lone \r をすべて \n に畳む（従来のテキストモード読みと互換）
        return re.sub(r'\r+\n?', '\n', text)
    raise Exception(f"ファイルを読み込めません: {path}")


def validate_bas_encoding(path):
    """.bas が CP932 で安全にインポートできるかの水際チェック。

    過去に繰り返された「CP932 の .bas を UTF-8 で上書きして日本語を壊す」事故を
    インポート前に機械的に検知する。日本語を含む CP932 のバイト列が偶然
    完全な UTF-8 として解釈できることはまず無いので、
    『非ASCIIを含むのに UTF-8 として読める』＝ UTF-8 化の疑い濃厚、として弾く。
    """
    with open(path, 'rb') as f:
        data = f.read()
    name = os.path.basename(path)
    if data.startswith(b'\xef\xbb\xbf'):
        print(f"エラー: {name} に UTF-8 BOM があります。")
        print("  この .bas は UTF-8 で保存されています。CP932(Shift-JIS) に戻してから実行してください。")
        return False
    if any(b > 0x7F for b in data):
        try:
            data.decode('cp932')
            cp932_ok = True
        except UnicodeDecodeError:
            cp932_ok = False
        try:
            data.decode('utf-8')
            utf8_ok = True
        except UnicodeDecodeError:
            utf8_ok = False
        if not cp932_ok:
            print(f"エラー: {name} は CP932 として読めません。エンコーディングを確認してください。")
            return False
        if utf8_ok:
            print(f"エラー: {name} は UTF-8 で保存されている疑いが濃厚です。")
            print("  このままインポートすると日本語が文字化けします。CP932(Shift-JIS) に変換してから実行してください。")
            return False
    return True


def normalize_bas_newlines(path):
    """インポート前に .bas の改行を正規 CRLF に揃える（改行二重化アーティファクトの水際修正）。

    「export → 編集 → テキストモードで書き戻し」の経路では、\\r\\n の各 \\n の前に
    余分な \\r が足されて \\r\\r\\n になることがある。これを VBA の Import に通すと
    1行おきに空行が挟まり、モジュールの行数が倍に膨れる（過去に繰り返した二重化事故の正体）。
    validate_bas_encoding が「文字コード事故」を弾くのと対になる「改行事故」の水際チェック。

    返り値 (fixed_bytes, raw_bytes, was_fixed):
        was_fixed=False なら元と完全一致＝無加工。True なら二重化等を検知・修正済み。

    クリーンな CRLF ファイルには冪等（バイト列が変わらないので was_fixed=False）。
    空行 (\\r\\n\\r\\n) は各 \\r\\n が個別にマッチするため保持される。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('cp932')
    # \r を1個以上含む改行（\r\n / \r\r\n / \r\r\r\n / lone \r）を一旦 \n に畳む
    norm = re.sub(r'\r+\n?', '\n', text)
    # lone \n（Unix改行）も含め、すべて正規 CRLF へ揃える
    norm = norm.replace('\n', '\r\n')
    fixed = norm.encode('cp932')
    return fixed, raw, (fixed != raw)


def validate_vba_code(code, force=False):
    """VBAコードの簡易バリデーション（構文・エンコード）"""
    # 1. CP932エンコード検証
    try:
        code.encode('cp932')
    except UnicodeEncodeError as e:
        bad_char = code[e.start:e.end]
        print(f"エラー: CP932(Shift_JIS)でエンコードできない文字 '{bad_char}' (インデックス: {e.start}) が含まれています。")
        print("VBAマクロでは文字化けやインポートエラーの原因となるため、修正してください。")
        if not force:
            return False
        print("警告: --force が指定されているため、検証エラーを無視して処理を続行します。")

    # 2. 簡易構文チェック (Sub/End Sub, Function/End Function の対のチェック)
    # コメント行を除外して検索。マルチステートメント行（Sub x(): End Sub 等）は
    # ':' で文に分割してから数える（1行書きを「End Sub 不足」と誤警告しない）
    clean_lines = []
    for line in code.split('\n'):
        stripped = line.strip()
        if stripped.startswith("'") or stripped.lower().startswith("rem "):
            continue
        blanked = re.sub(r'"[^"]*"', '""', stripped)   # 文字列内の ':' で誤分割しない
        blanked = blanked.split("'", 1)[0]              # 行末コメント内の ': Sub' 等を数えない
        for seg in blanked.split(':'):
            clean_lines.append(seg.strip())
    clean_code = '\n'.join(clean_lines)

    decl_sub = len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?\bSub\b', clean_code, re.IGNORECASE | re.MULTILINE))
    end_sub = len(re.findall(r'^\s*End\s+Sub\b', clean_code, re.IGNORECASE | re.MULTILINE))
    decl_func = len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?\bFunction\b', clean_code, re.IGNORECASE | re.MULTILINE))
    end_func = len(re.findall(r'^\s*End\s+Function\b', clean_code, re.IGNORECASE | re.MULTILINE))

    errors = []
    if decl_sub != end_sub:
        errors.append(f"Sub宣言数 ({decl_sub}) と End Sub数 ({end_sub}) が一致しません。")
    if decl_func != end_func:
        errors.append(f"Function宣言数 ({decl_func}) と End Function数 ({end_func}) が一致しません。")

    # 3. プロシージャ名の識別子チェック（`_tmp検証` 級の先頭 _ 注入が繰り返された定番事故。
    #    AddFromString/InsertLines は構文検査をせず成功報告になるため、ここが水際）
    for ln, _name, reason in _find_invalid_procedure_names(code):
        errors.append(f"行{ln}: プロシージャ名が VBA の識別子規則に反しています: {reason}")

    if errors:
        for err in errors:
            print(f"構文エラー警告: {err}")
        if not force:
            print("エラー: 構文チェックに失敗しました。(--force で無視して実行可能)")
            return False
        print("警告: --force が指定されているため、構文エラーを無視して処理を続行します。")
    return True


def make_module_backup(wb, module_name):
    """モジュール単位のバックアップを .bas 形式で保存"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    base_wb_name = os.path.splitext(os.path.basename(wb.FullName))[0]
    backup_filename = f"{base_wb_name}_{module_name}_{timestamp}.bas"
    backup_path = os.path.join(BACKUP_DIR, backup_filename)

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            try:
                comp.Export(backup_path)
                print(f"モジュールバックアップ作成: backups/{backup_filename}")
                return backup_path
            except Exception as e:
                print(f"警告: モジュールバックアップ失敗 ({e})")
    return None


class ModuleNameCollisionError(Exception):
    """Remove+Import の名前衝突で、モジュールを期待名で取り込めなかったことを示す。

    Import 自体は成功しており、置換後のコードは actual_name（連番付き別名）側に
    存在している。呼び出し側はバックアップの再 Import で「復旧」してはいけない
    （同じ VB_Name の .bas を重ねると連番モジュールがさらに増えるだけ）。
    """
    def __init__(self, expected_name, actual_name, message):
        super().__init__(message)
        self.expected_name = expected_name
        self.actual_name = actual_name


def _find_component(wb, name):
    """VBComponent を名前（大小無視）で探す。無ければ None"""
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == name.lower():
            return c
    return None


def _wait_component_gone(wb, name, timeout=15.0, interval=0.05):
    """Remove 発行後、同名コンポーネントが VBProject から実際に消えるまで待つ。

    VBE の Remove は、対象モジュールのプロシージャが実行中（メニューや
    ショートカット経由の呼び出し中）などの場合に遅延完了する。消える前に
    Import すると名前衝突で「shu0051」のような連番付き別名で取り込まれる
    （2026-07-11 深夜の shu005 消滅事故の直接原因）。
    戻り値: 消えたら True / timeout まで残っていたら False
    """
    deadline = time.monotonic() + timeout
    while True:
        pythoncom.PumpWaitingMessages()
        if _find_component(wb, name) is None:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(interval)


def _reregister_shortcuts_from_bas(wb, import_path):
    """Import した .bas の Attribute VB_Invoke_Func からショートカットをその場で再登録する。

    ショートカット（Attribute VB_ProcData.VB_Invoke_Func）が Excel に登録されるのは
    「ブックを開いた瞬間」だけ。Remove+Import はモジュールごと登録を剥がすため、
    Import 直後のセッションではショートカットが無反応になり、閉じて開き直すまで
    直らない（2026-07-12 発覚。修正のたびにマクロが効かなくなる症状の正体）。
    ここで MacroOptions により実行中セッションへ再登録して、開き直し不要にする。
    大文字=Ctrl+Shift+キー / 小文字=Ctrl+キー（MacroOptions の仕様どおり渡す）。

    ⚠ ロード中アドインが同名プロシージャを持つ場合は掛け直さない（2026-08-25）。
      ショートカットキーは Excel セッションで1つの持ち主しか持てない。編集中のブックへ
      付け替えると、アドインから持ち主を奪う形になる。そのブックが退場した瞬間
      （例: アドインの更新登録＝ブックを .xlam として保存し直す運用）キーの行き先が消え、
      押すと「オートメーション エラー／例外が発生しました」になる。閉じて開き直すまで
      直らないので、修理のたびに再発していた。アドインはブックより長生きするので、
      持ち主はアドインに残す。アドインが居ないときは従来どおり編集ブックへ掛ける。
    """
    try:
        with open(import_path, 'rb') as f:
            text = f.read().decode('cp932', errors='replace')
    except Exception:
        return
    pairs = re.findall(
        r'^Attribute\s+([^\s.]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"(.)\\n14"',
        text, re.MULTILINE)
    if not pairs:
        return
    xl = wb.Application
    for proc, key in pairs:
        if not key.strip():
            continue  # 空白キー（キー未割当の名残）は再登録しない
        owner = _addin_project_owning_proc(wb, proc)
        if owner:
            print(f"  ショートカット再登録を見送り: {proc} "
                  f"（アドイン {owner} が同名を持つ。持ち主を奪うと更新登録後に宙に浮く）")
            continue
        try:
            xl.MacroOptions(Macro=f"'{wb.Name}'!{proc}",
                            HasShortcutKey=True, ShortcutKey=key)
            label = f"Ctrl+Shift+{key.upper()}" if key.isupper() else f"Ctrl+{key}"
            print(f"  ショートカット再登録: {label} → {proc}")
        except Exception as ex:
            print(f"  ⚠ ショートカット再登録に失敗: {proc} ({ex})")


def _addin_project_owning_proc(wb, proc):
    """ロード中のアドイン(.xlam/.xla)のうち、指定プロシージャを持つものの名前を返す。

    無ければ None。汎用ツールなので特定ブック名には依存せず、拡張子で判定する。
    自分自身（wb）は対象外。判定できないときは None（＝従来どおり掛け直す）。
    """
    pat = re.compile(
        r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s+' +
        re.escape(proc) + r'\s*\(', re.IGNORECASE | re.MULTILINE)
    try:
        projects = list(wb.Application.VBE.VBProjects)
        self_name = wb.Name.lower()
    except Exception:
        return None
    for p in projects:
        try:
            fn = p.Filename
        except Exception:
            continue  # 未保存等でファイル名が取れないものは対象外
        base = os.path.basename(fn)
        if base.lower() == self_name:
            continue
        if os.path.splitext(base)[1].lower() not in ('.xlam', '.xla'):
            continue
        try:
            for comp in p.VBComponents:
                try:
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0 and pat.search(cm.Lines(1, cm.CountOfLines)):
                        return base
                except Exception:
                    continue
        except Exception:
            continue
    return None


REVIVAL_MACRO = 'ショートカット復活'


def _run_shortcut_revival(wb):
    """モジュール入れ直し後、「ショートカット復活」マクロを持つプロジェクトがあれば自動実行する。

    _reregister_shortcuts_from_bas が掛け直せるのは入れ直したモジュール自身の
    Attribute 分だけで、アドイン等が別途登録しているショートカットは入れ直しの
    たびに死ぬ（2026-07-25 発覚: ファイル一覧.xlsm の全検索モジュールを置換したら
    秀.xlam 側の「保存で閉じる」が無反応になった）。
    汎用ツールなので特定ブック名には依存せず、「ショートカット復活」という公開
    Sub を持つロード中プロジェクトを探して呼ぶ、というマクロ名の規約で対処する。
    見つからなければ無言で何もしない。実行失敗でも本体の置換は成功扱い
    （ショートカットはブックを開き直せば直る種類の症状のため）。
    """
    pattern = re.compile(
        r'^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+' + re.escape(REVIVAL_MACRO) + r'\s*\(',
        re.IGNORECASE | re.MULTILINE)
    try:
        xl = wb.Application
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return

    # アドインを最後に回す（2026-08-25）。復活マクロは自分の ThisWorkbook 宛に
    # MacroOptions を打つので、同名マクロを持つブックとアドインが両方ロードされていると
    # 「後に走ったほうが持ち主」になる。列挙順は運任せなので、長生きするアドインを
    # 最後に走らせて持ち主を確定させる。
    def _is_addin(p):
        try:
            return os.path.splitext(os.path.basename(p.Filename))[1].lower() in ('.xlam', '.xla')
        except Exception:
            return False
    projects.sort(key=_is_addin)

    for p in projects:
        try:
            book_name = os.path.basename(p.Filename)
        except Exception:
            continue  # 未保存等でブック名が取れないプロジェクトは対象外
        hit = False
        try:
            for comp in p.VBComponents:
                try:
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0 and pattern.search(cm.Lines(1, cm.CountOfLines)):
                        hit = True
                        break
                except Exception:
                    continue
        except Exception:
            continue
        if not hit:
            continue
        try:
            quoted = book_name.replace("'", "''")
            xl.Application.Run(f"'{quoted}'!{REVIVAL_MACRO}")
            print(f"  ショートカット復活 実行: {book_name}")
        except Exception as ex:
            print(f"  ⚠ ショートカット復活の実行に失敗: {book_name} ({ex})")


def _import_module_verified(wb, import_path, expected_name,
                            ghost_timeout=15.0, rename_timeout=20.0, settle=0.1):
    """Import ＋ 取り込み実名の検証。Remove+Import 系3経路の共用ガード。

    2026-07-11 深夜の実害: replace-procedure（Attribute経路）で shu005 を置換した際、
    Remove の遅延完了中に Import が走って名前衝突し「shu0051」として取り込まれた。
    ツールは成功と報告し、shu005 の消滅は翌日まで発覚しなかった。
    同じ事故を二度と起こさないため、ここで
      (1) Import 前: expected_name の残骸が消えたことを確認（遅延 Remove 対策）。
          timeout しても Import 自体は行う（見送ると遅延 Remove だけが後から完了して
          モジュールが完全消滅するため、衝突覚悟で取り込んで (3) で回復する）。
      (2) Import 後: 返ってきた VBComponent の実名を expected_name と照合。
      (3) 不一致（連番付き別名等）なら旧名の消滅を待って改名で自動回復。
          回復できなければ ModuleNameCollisionError（黙って成功と報告しない）。
    成功時は Import 済みコンポーネントを返す（Save は呼び出し側の責務）。
    """
    if not _wait_component_gone(wb, expected_name, timeout=ghost_timeout):
        print(f"⚠ 旧 '{expected_name}' の Remove がまだ完了していません"
              f"（同モジュールのコードが実行中の可能性）。実名検証つきで Import を続行します。")
    imported = wb.VBProject.VBComponents.Import(import_path)
    if settle:
        time.sleep(settle)
    pythoncom.PumpWaitingMessages()
    actual = imported.Name
    if actual.lower() == expected_name.lower():
        _reregister_shortcuts_from_bas(wb, import_path)
        _run_shortcut_revival(wb)
        return imported
    # 名前衝突を検出（例: shu005 → shu0051）。旧名が空くのを待って改名で回復する
    print(f"⚠ 名前衝突を検出: '{expected_name}' が '{actual}' として取り込まれました。"
          f"旧モジュールの消滅を待って改名で回復します...")
    if not _wait_component_gone(wb, expected_name, timeout=rename_timeout):
        raise ModuleNameCollisionError(
            expected_name, actual,
            f"旧 '{expected_name}' が残ったままのため '{actual}' を改名できません")
    try:
        imported.Name = expected_name
    except Exception as ex:
        raise ModuleNameCollisionError(
            expected_name, actual, f"'{actual}' → '{expected_name}' の改名に失敗: {ex}")
    if imported.Name.lower() != expected_name.lower():
        raise ModuleNameCollisionError(
            expected_name, actual, f"改名後の実名が '{imported.Name}' のままです")
    print(f"回復成功: '{actual}' → '{expected_name}' に改名しました。")
    _reregister_shortcuts_from_bas(wb, import_path)
    _run_shortcut_revival(wb)
    return imported


def _print_collision_guidance(ex, module_name, module_backup, err=False):
    """名前衝突が自動回復できなかったときの案内（3経路共通・黙って成功にしない）"""
    out = sys.stderr if err else sys.stdout
    print(f"エラー: 名前衝突からの自動回復に失敗しました: {ex}", file=out)
    print(f"  ⚠ 置換後のコードは別名モジュール '{ex.actual_name}' として存在しています。", file=out)
    print(f"  ⚠ バックアップの再 Import はしないでください（VB_Name 衝突で連番モジュールが増えます）。", file=out)
    print(f"  対処: 旧 '{module_name}' が消えているのを確認してから、VBE のプロパティウィンドウで", file=out)
    print(f"        '{ex.actual_name}' の (オブジェクト名) を '{module_name}' に改名してください。", file=out)
    if module_backup:
        print(f"  置換前の内容のバックアップ: {module_backup}", file=out)


def _save_with_retry(wb, attempts=5, delay=0.6):
    """Import 直後の保存だけはビジー拒否で諦めない。

    Excel はセル編集中・メニュー展開中・モーダル表示中に COM 呼び出しを
    RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER / VBA_E_IGNORE で弾く
    （_com_is_busy 参照）。Remove+Import 直後にこれを食らうと、モジュールは
    正しく入れ替わっているのにブックだけ未保存で残る。一過性の拒否がほとんど
    なので、少し待って撃ち直す。恒久的な失敗（読み取り専用・ロック等）は
    そのまま送出して呼び出し側に判断させる。
    """
    last = None
    for i in range(attempts):
        try:
            wb.Save()
            return
        except Exception as ex:
            last = ex
            if not _com_is_busy(ex) or i == attempts - 1:
                raise
            print(f"  Excel がビジーのため保存を再試行します（{i + 1}/{attempts - 1}）...")
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    if last:
        raise last


def _print_save_failed_guidance(module_name, err=False):
    """Import は成功したが、その後の『保存』で失敗したときの案内。

    2026-07-14 に発見した実害筋: Remove+Import 系3経路の except が removed
    フラグしか見ておらず、Import 成功後に wb.Save() が失敗すると「モジュールが
    消えた」と誤認してバックアップを再 Import していた。期待名のモジュールは
    既に正しく存在するので _wait_component_gone は空振りし（15秒）、VB_Name 衝突で
    連番別名として取り込まれ、改名待ちも空振りして（20秒）例外になる——つまり
    ツール自身が「旧コード入りの連番モジュール」を生んで 35 秒待たせていた。
    ここに来たら再 Import は絶対にしない。壊れているのは『保存』だけである。
    """
    out = sys.stderr if err else sys.stdout
    print(f"  ⚠ モジュール '{module_name}' の取り込みは成功しています"
          f"（開いているブックの中身は新しいコードに置き換わっています）。", file=out)
    print(f"  ⚠ 失敗したのは『保存』だけです。バックアップの再 Import はしません"
          f"（すると連番モジュールが増えるだけです）。", file=out)
    print(f"  対処: Excel のダイアログ・セル編集モードを解除してから、", file=out)
    print(f"        Excel で保存する（Ctrl+S）か、同じコマンドをもう一度実行してください。", file=out)
    print(f"  ⚠ 保存せずにブックを閉じると、この置換内容は失われます。", file=out)




# ================================================================
# 共有ヘルパー（分割時に後続パートから移設・前方参照解消）
# ================================================================

def check_vba_identifier(name):
    """VBA 識別子として無効なら理由（文字列）を返す。有効なら None。

    先頭 `_` の名前（例: `_tmp検証`）は VBE が黙って受け入れるがコンパイルで死ぬ。
    AddFromString / InsertLines は構文検査をしないため、注入自体は成功報告になり
    事故が繰り返された。VBA の識別子は英字か日本語などの文字で始まる必要があり、
    `_`・数字・記号では始められない。注入前にここで機械的に止める。
    """
    if not name:
        return "名前が空です"
    if name[0] == '_':
        suggestion = name.lstrip('_') or 'tmp'
        return (f"'{name}' は _ 始まりです。VBA の識別子は _ で始められません"
                f"（英字か日本語で始める。例: '{suggestion}'）")
    if name[0].isdigit():
        return f"'{name}' は数字始まりです。VBA の識別子は英字か日本語で始めてください"
    bad = re.sub(r'\w', '', name)
    if bad:
        return f"'{name}' に識別子に使えない文字が含まれています: {bad}"
    if len(name) > 255:
        return f"'{name}' が長すぎます（{len(name)}文字。VBA の上限は255文字）"
    return None

def _find_invalid_procedure_names(norm_text):
    """Sub/Function/Property 宣言の名前が VBA 識別子規則に反するものを列挙。

    コメント行は対象外。Declare 宣言は行頭トークンが合わないので元から素通り
    （外部 API 名は別規則のため対象にしない）。
    戻り値: [(行番号, 名前, 理由), ...]
    """
    decl_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s\(]+)',
        re.IGNORECASE)
    hits = []
    for idx, line in enumerate(norm_text.split('\n'), 1):
        s = line.strip()
        if s.startswith("'") or s.lower().startswith('rem '):
            continue
        m = decl_pattern.match(line)
        if m:
            reason = check_vba_identifier(m.group(1))
            if reason:
                hits.append((idx, m.group(1), reason))
    return hits

def _col_letter(n):
    """列番号(1始まり)を A, B, ... Z, AA, ... に変換"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _reject_extra_args(rest, used, usage):
    """位置引数の食い残しをエラーにする。

    黙って捨てると `clear-range Sheet1 A1:B2` のように「第2引数のつもりの範囲」が
    無視され、シート全域が対象になる事故（過去に実害）につながる。
    """
    if len(rest) > used:
        print(f"エラー: 余分な引数があります: {' '.join(str(a) for a in rest[used:])}")
        print(f"  {usage}")
        return True
    return False

def _cell_str(v):
    """セル値を表示用文字列に"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime):
        # pywintypes の日時は str() だと "+00:00"（TZ）が付き、TSV 往復の
        # 書き戻しで日付がテキスト化する。素直な表記に整える
        # （_coerce_cell がこの表記を日付として復元する＝往復が一致する）
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime('%Y/%m/%d')
        return v.strftime('%Y/%m/%d %H:%M:%S')
    return str(v)

def _range_values_2d(rng, use_formula=False):
    """Range の値を 2次元リストへ正規化（--tsv / --json 用）"""
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return [['']]
    if not isinstance(raw, tuple):
        return [[raw]]
    return [list(r) if isinstance(r, tuple) else [r] for r in raw]

_LAST_VALUES_FILE = os.path.join(SCRIPT_DIR, '_last_values.tsv')   # write-range のグリッド入力

_LAST_SNAPSHOT_FILE = os.path.join(SCRIPT_DIR, '_last_snapshot.json')  # snapshot の意味構造JSON出力先

_JOB_CLOCK_FILE = os.path.join(SCRIPT_DIR, '_job_clock.json')   # materials が押す仕事の時計（tidy／write が経過秒を出す）


def job_clock_start(label=""):
    """仕事の時計を押す（materials が呼ぶ）。

    同じ仕事の実測が 44 秒→18.5 秒で、道具の中は 5 秒で不変、差は全部往復のあいだの考え込みだった
    （2026-09-02 夜）。測らないと戻るので道具が自分で測り、tidy／write-range／write-cells が
    「materials から N 秒」を出す。1 時間より古い時計は無視する（job_clock_elapsed）。
    """
    import json
    try:
        with open(_JOB_CLOCK_FILE, 'w', encoding='utf-8') as f:
            json.dump({"start": time.time(), "label": label}, f, ensure_ascii=False)
    except Exception:
        pass


def job_clock_elapsed(max_age=3600):
    """時計を押してからの秒数。押していない／max_age 秒より古い／読めないときは None。"""
    import json
    try:
        with open(_JOB_CLOCK_FILE, 'r', encoding='utf-8') as f:
            sec = time.time() - float(json.load(f).get("start", 0))
    except Exception:
        return None
    if sec < 0 or sec > max_age:
        return None
    return sec


def job_clock_get():
    """いまの時計の中身をそのまま返す（退避用）。押していない／読めないときは None。"""
    import json
    try:
        with open(_JOB_CLOCK_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def job_clock_set(d):
    """job_clock_get で退避した中身に時計を戻す（d が None なら何もしない）。"""
    import json
    if not d:
        return
    try:
        with open(_JOB_CLOCK_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception:
        pass


def job_clock_note():
    """経過を 1 行にする（時計が無ければ空文字）。書き込み系コマンドの末尾に付ける。"""
    sec = job_clock_elapsed()
    return "" if sec is None else f"経過: materials から {sec:.1f} 秒"


# ---- 呼び出し台帳（2026-09-16）----
# コマンドを 1 回撃つたびに 1 行足す（コマンド名・引数の頭・道具の中の秒・成否・対象ブック・経路）。
# それまでは AI 作業窓（AI_WIN_LOG）から呼んだときしか記録が無く、会話からの呼び出しがどれだけ速いか・
# どこで遅いかを数字で言えなかった。`stats` が集計する。道具の外（考える・読む・人の時間）は、
# 前の呼び出しが終わってから次を撃つまでの「間」として同じ台帳から出る
_CALL_LOG_FILE = os.path.join(SCRIPT_DIR, '_calls.jsonl')
_CALL_LOG_ARGS_MAX = 80        # 引数は頭だけ（セルの値をだらだら残さない）
_CALL_GAP_MAX = 600            # 「間」に数える上限（秒）。これより空いたら別の仕事

# ---- コマンドの進み具合（2026-09-16）----
# 長いコマンド（export-all・checkup・gate・rehearse・materials）は MCP 越しだと終わるまで無言。
# 台帳の包み（vba_manager._logged）が「始めた／終わった」を、長い手が途中の段を note に書き、`status` が読む。
# agent の進み具合（_agent_progress.json・vbam_ai が書く）とは別のファイル＝本体は AI の module を import しない。
_CMD_PROGRESS_FILE = os.path.join(SCRIPT_DIR, '_cmd_progress.json')
_AGENT_PROGRESS_PATH = os.path.join(SCRIPT_DIR, '_agent_progress.json')


def cmd_progress_write(state, cmd=None, args=None, ok=None, sec=None, note=None):
    """進み具合を 1 ファイルに上書きする。書けなくても仕事は止めない。"""
    import json
    try:
        rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'state': state, 'pid': os.getpid()}
        if cmd is not None:
            rec['cmd'] = str(cmd)
        if args is not None:
            pos = list(getattr(args, 'posargs', None) or [])
            head = " ".join(str(p) for p in pos[:3])
            rec['args'] = head[:_CALL_LOG_ARGS_MAX]
        if ok is not None:
            rec['ok'] = bool(ok)
        if sec is not None:
            rec['sec'] = round(float(sec), 3)
        if note:
            rec['note'] = str(note)
        with open(_CMD_PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(rec, f, ensure_ascii=False)
        return rec
    except Exception:
        return None


def cmd_progress_note(note):
    """走っているコマンドの「いまの段」を書き足す（始めた記録が無ければ何もしない）。"""
    import json
    try:
        rec = cmd_progress_read()
        if not rec or rec.get('state') != 'run' or int(rec.get('pid') or 0) != os.getpid():
            return
        rec['note'] = str(note)
        rec['note_time'] = time.strftime('%H:%M:%S')
        with open(_CMD_PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(rec, f, ensure_ascii=False)
    except Exception:
        pass


def cmd_progress_read():
    """進み具合のファイルを読む（無ければ None）。"""
    import json
    try:
        with open(_CMD_PROGRESS_FILE, encoding='utf-8') as f:
            rec = json.load(f)
        return rec if isinstance(rec, dict) else None
    except (OSError, ValueError):
        return None


def _pid_alive(pid):
    """その PID のプロセスが生きているか（Windows）。読めなければ False。"""
    if not pid:
        return False
    try:
        import win32api
        import win32con
        import win32process
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, int(pid))
        try:
            return win32process.GetExitCodeProcess(h) == 259        # STILL_ACTIVE
        finally:
            win32api.CloseHandle(h)
    except Exception:
        return False


def call_log_write(cmd, args, sec, ok, via="cli"):
    """呼び出し台帳に 1 行足す。失敗しても何も起こさない（台帳のために仕事を止めない）。"""
    import json
    try:
        pos = list(getattr(args, 'posargs', None) or [])
        head = " ".join(str(p) for p in pos[:3])
        if len(head) > _CALL_LOG_ARGS_MAX:
            head = head[:_CALL_LOG_ARGS_MAX] + "…"
        rec = {"time": time.strftime('%Y-%m-%d %H:%M:%S'), "cmd": str(cmd), "args": head,
               "sec": round(float(sec), 3), "ok": (None if ok is None else bool(ok)),
               "book": _last_book_name, "via": via}
        with open(_CALL_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass


def call_log_read(days=7):
    """台帳の直近 days 日ぶんを古い順に返す（読めない行は飛ばす）。"""
    import json
    since = time.time() - float(days) * 86400
    out = []
    try:
        with open(_CALL_LOG_FILE, 'r', encoding='utf-8') as f:
            for ln in f:
                try:
                    r = json.loads(ln)
                    t = time.mktime(time.strptime(r["time"], '%Y-%m-%d %H:%M:%S'))
                except Exception:
                    continue
                if t >= since:
                    r["_t"] = t
                    out.append(r)
    except OSError:
        return []
    return out


def call_log_summary(recs, top=20, slow=8):
    """台帳の行 → 集計（純 Python）。コマンド別・遅かった呼び出し・呼び出しの「間」。"""
    by = {}
    for r in recs:
        e = by.setdefault(r.get("cmd", "?"), {"cmd": r.get("cmd", "?"), "count": 0, "sec": 0.0, "max": 0.0, "fail": 0})
        s = float(r.get("sec") or 0)
        e["count"] += 1
        e["sec"] += s
        e["max"] = max(e["max"], s)
        if r.get("ok") is False:
            e["fail"] += 1
    cmds = sorted(by.values(), key=lambda e: -e["count"])
    for e in cmds:
        e["avg"] = e["sec"] / e["count"] if e["count"] else 0.0
    slowest = sorted(recs, key=lambda r: -float(r.get("sec") or 0))[:slow]
    gaps = []
    for a, b in zip(recs, recs[1:]):
        g = float(b["_t"]) - float(b.get("sec") or 0) - float(a["_t"])
        if 0 <= g <= _CALL_GAP_MAX:
            gaps.append(g)
    gaps.sort()
    med = gaps[len(gaps) // 2] if gaps else None
    return {"count": len(recs), "fail": sum(1 for r in recs if r.get("ok") is False),
            "tool_sec": sum(float(r.get("sec") or 0) for r in recs),
            "books": sorted({r["book"] for r in recs if r.get("book")}),
            "commands": cmds[:top], "slowest": slowest,
            "gap_count": len(gaps), "gap_median": med,
            "gap_mean": (sum(gaps) / len(gaps)) if gaps else None, "gap_sec": sum(gaps)}


def cmd_stats(args):
    """呼び出し台帳の集計: stats [--days N] [--top N] [--slow N] [--json]（Excel に触らない）"""
    import json
    days = int(getattr(args, 'days', None) or 7)
    top = int(getattr(args, 'top', None) or 20)
    slow = int(getattr(args, 'slow', None) or 8)
    recs = call_log_read(days)
    if not recs:
        print(f"呼び出し台帳に直近 {days} 日の記録がありません（コマンドを撃つと {os.path.basename(_CALL_LOG_FILE)} に溜まります）")
        return True
    s = call_log_summary(recs, top=top, slow=slow)
    if getattr(args, 'json', False):
        out = dict(s)
        out["days"] = days
        out["slowest"] = [{k: v for k, v in r.items() if k != "_t"} for r in s["slowest"]]
        print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        return True
    print(f"呼び出し台帳: 直近 {days} 日  {s['count']} 回・失敗 {s['fail']} 回・道具の中の合計 {s['tool_sec']:.1f} 秒"
          + (f"・対象ブック: {', '.join(s['books'][:6])}" + ("…" if len(s['books']) > 6 else "") if s['books'] else ""))
    print("  コマンド" + " " * 12 + "   回数   平均秒   最長秒  失敗")     # 全角は 2 桁幅＝手で揃える
    for e in s["commands"]:
        print(f"  {e['cmd']:<20} {e['count']:>5} {e['avg']:>7.2f} {e['max']:>7.2f} {e['fail']:>4}")
    print(f"遅かった呼び出し（上位 {len(s['slowest'])}）:")
    for r in s["slowest"]:
        book = f"  [{r['book']}]" if r.get("book") else ""
        print(f"  {r['time'][5:]}  {r['cmd']} {r.get('args', '')}{book}  {float(r.get('sec') or 0):.2f} 秒")
    if s["gap_median"] is not None:
        print(f"呼び出しの間（前が終わってから次を撃つまで・{_CALL_GAP_MAX // 60} 分未満の {s['gap_count']} 本）: "
              f"中央値 {s['gap_median']:.1f} 秒・平均 {s['gap_mean']:.1f} 秒・合計 {s['gap_sec']:.0f} 秒")
        print("  ＝道具の外（AI が考える・読む・人の時間）。道具の中の合計と比べると、どちらを縮めるべきかが分かる")
    return True

def _coerce_cell(s):
    """文字列をセル値に変換。'='始まりは数式、数値は数値、空は None。

    先頭が 0 の数字（郵便番号 "007"・伝票番号 "0004"）は文字のまま返す（2026-09-23）。
    """
    if s is None or s == '':
        return None
    if s.startswith('='):
        return s                      # 数式 (.Value への代入で Excel が数式と解釈)
    if s.startswith("'"):
        return s[1:]                  # Excel と同じ: 先頭の ' は「文字のまま」（日付・数値に見えても変換しない。' は残らない）
    if not s.isascii():
        return s                      # 全角数字「１２３」等は Excel と同じくテキストのまま
    # 日付表記（_cell_str の出力と同じ形）は datetime で書き戻す。
    # 文字列のまま .Value に入れるとテキスト格納になり、TSV 往復で日付列が壊れる。
    # tzinfo=UTC を付ける: 素の datetime は pywin32 が「ローカル時刻」とみなして UTC に直してから
    # Excel に渡すので、日本では 9 時間前（前日 15:00）にずれる（2026-09-04 実射で発覚）
    m = re.fullmatch(
        r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})'
        r'(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?', s)
    if m:
        try:
            return datetime.datetime(
                int(m.group(1)), int(m.group(2)), int(m.group(3)),
                int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0),
                tzinfo=datetime.timezone.utc)
        except ValueError:
            return s                  # 2026/99/99 のような非実在日付は文字列のまま
    if re.fullmatch(r'0\d+', s):
        # 先頭が 0 の数字（伝票番号 0004・郵便番号 007）は番号＝文字のまま書く。数にすると 0 が消え、
        # 書いた後の「化けたら文字に戻す」も「意図した変換」と見て戻さなかった（2026-09-23 通しの実測 2 の A5）
        return s
    if re.fullmatch(r'-?\d+', s):
        try:
            return int(s)
        except ValueError:
            return s
    # Python の float() は "1_000" や前後空白も受理してしまい、
    # "2026_07" のようなテキストIDが黙って数値化される。桁区切り表記や
    # 空白付きは文字列のまま扱う
    if '_' in s or s != s.strip():
        return s
    try:
        f = float(s)
    except ValueError:
        return s
    # "nan"/"inf" は float 化すると Excel 上でエラー値になるため文字列のまま
    if f != f or f in (float('inf'), float('-inf')):
        return s
    return f


# ================================================================
# 保護シート対策（2026-07-13）
#
# UserInterfaceOnly:=True で保護したシートでも、外部COM(このツール)からの
# ClearContents/Clear/CopyPicture・Shapes操作の一部は例外的にブロックされる
# （Value代入は素通りするが挙動が不統一）。VBAマクロ実行(run-macro)経由の
# 変更だけは正しく素通りするため、これは「マクロ実行か外部COMか」で
# UserInterfaceOnly の適用され方が違うことに起因する。
# 対策として、保護されたシートのある操作の前後で一時解除→元の設定で再保護する。
# ================================================================

# Protect()/Protection オブジェクトの真偽値フラグ（DrawingObjects/Contents/Scenarios/
# UserInterfaceOnly の4つ以外）。Worksheet.Protection.AllowXxx で現在値を読める。
# これを記録・復元しないと、並べ替え許可や行列挿入許可などブック固有のカスタム設定が
# 一時解除→再保護の往復で既定値(False=不許可)に巻き戻ってしまう。
_PROTECTION_ALLOW_FLAGS = (
    'AllowFormattingCells', 'AllowFormattingColumns', 'AllowFormattingRows',
    'AllowInsertingColumns', 'AllowInsertingRows', 'AllowInsertingHyperlinks',
    'AllowDeletingColumns', 'AllowDeletingRows',
    'AllowSorting', 'AllowFiltering', 'AllowUsingPivotTables',
)


def _unprotect_all_sheets(wb):
    """ブック内の保護されたシートを全て記録して一時解除する。

    戻り値は再保護用の (ws, drawing, contents, scenarios, ui_only, allow_kwargs) の
    リスト。allow_kwargs は並べ替え許可・行列挿入許可など細かい許可設定の辞書
    （Protect() にそのまま **allow_kwargs で渡せる）。
    パスワード保護等で解除できないシートは諦めて記録しない（そのシートは
    保護されたままなので、そのシートを触る操作は従来どおり失敗しうる）。
    """
    saved = []
    try:
        sheets = list(wb.Worksheets)
    except Exception:
        return saved
    for ws in sheets:
        try:
            protected = bool(ws.ProtectContents or ws.ProtectDrawingObjects
                              or ws.ProtectScenarios)
        except Exception:
            protected = False
        if not protected:
            continue
        try:
            drawing = bool(ws.ProtectDrawingObjects)
        except Exception:
            drawing = False
        try:
            contents = bool(ws.ProtectContents)
        except Exception:
            contents = False
        try:
            scenarios = bool(ws.ProtectScenarios)
        except Exception:
            scenarios = False
        try:
            ui_only = bool(ws.ProtectionMode)
        except Exception:
            ui_only = False
        allow_kwargs = {}
        try:
            prot = ws.Protection
            for flag in _PROTECTION_ALLOW_FLAGS:
                try:
                    allow_kwargs[flag] = bool(getattr(prot, flag))
                except Exception:
                    pass
        except Exception:
            pass
        try:
            # Password を省略すると、パスワード保護シートに対して Excel は例外ではなく
            # 「パスワード入力ダイアログ」を出し、COM 呼び出しがそこでブロックする
            # （DisplayAlerts=False でも抑止されない）。この解除は dialog_safe の
            # ウォッチャーが起動する前（protect_safe が外側）に走るため、誰も解除できず
            # 人が手で閉じるまで戻らない＝無言ハングになる。
            # ダミーを渡せば「パスワード相違」で確実に例外化し、下の警告経路に落ちる。
            # （パスワード無しの保護シートに対しては、渡しても普通に解除できる）
            ws.Unprotect(Password=_UNPROTECT_PROBE_PW)
            saved.append((ws, drawing, contents, scenarios, ui_only, allow_kwargs))
        except Exception as ex:
            # 解除できなかった＝そのシートは保護されたまま。以後の操作はそこで失敗するが、
            # 従来は完全に無言だったため「なぜ失敗したのか」が分からなかった
            print(f"⚠ シート '{_ws_name(ws)}' の保護を一時解除できませんでした"
                  f"（パスワード保護の可能性）: {ex}", file=sys.stderr)
            print("  このシートを触る操作は失敗します。Excel 側で先に保護を解除してください。",
                  file=sys.stderr)
    return saved


class ReopenFailed(Exception):
    """読み取り専用で開いたブックの「開き直し」に失敗した。

    get_workbook のキャッシュ生存確認は例外を握って再接続に落ちる作りだが、
    この失敗だけは握りつぶしてはいけない（握ると読み取り専用のままのブックを
    返し、呼び出し元の replace-procedure 等がそのまま書き込みに進む）。
    生存確認の失敗と区別するための専用型。
    """


# Unprotect に渡すダミーパスワード。Password 省略時のパスワード入力ダイアログ
# （＝無言ハング）を避け、確実に例外へ倒すためだけに使う。解錠が目的ではない。
_UNPROTECT_PROBE_PW = "__vbam_probe__"


def _ws_name(ws):
    """シート名（取れなければ '?'）。エラーメッセージ用"""
    try:
        return ws.Name
    except Exception:
        return "?"


def _reprotect_sheets(saved):
    """_unprotect_all_sheets で記録した設定どおりに再保護する。

    再保護に失敗したら必ず報告する。黙って諦めると「ブックの保護が外れたまま
    コマンドは成功終了」になり、保護されているつもりのブックが無防備になる。
    """
    for ws, drawing, contents, scenarios, ui_only, allow_kwargs in saved:
        try:
            ws.Protect(DrawingObjects=drawing, Contents=contents,
                       Scenarios=scenarios, UserInterfaceOnly=ui_only,
                       **allow_kwargs)
        except Exception:
            try:
                # allow_kwargs の一部が今の Excel バージョンで受理されない場合の保険
                ws.Protect(DrawingObjects=drawing, Contents=contents,
                           Scenarios=scenarios, UserInterfaceOnly=ui_only)
                print(f"⚠ シート '{_ws_name(ws)}' を再保護しましたが、"
                      "細かい許可設定（並べ替え許可など）は復元できませんでした。",
                      file=sys.stderr)
            except Exception as ex:
                print(f"⚠ シート '{_ws_name(ws)}' の再保護に失敗しました: {ex}",
                      file=sys.stderr)
                print("  このシートは保護が外れたままです。Excel 側で保護し直してください。",
                      file=sys.stderr)


# 実行中の protected_sheets_guard（forget_protection から参照する）
_active_guards = []


def forget_protection(ws):
    """このシートを「ガードの再保護対象」から外す（保護解除そのものが目的の操作用）。

    ガードは入口で全保護シートを一時解除し、出口で記録どおり再保護する。
    そのため sheet unprotect のように「保護を外すこと自体が目的」の操作は、
    出口で保護を戻されて無言で効かなくなる。この関数で記録から落としておく。
    """
    try:
        name = ws.Name
    except Exception:
        return
    for g in _active_guards:
        kept = []
        for entry in g._saved:
            try:
                if entry[0].Name == name:
                    continue
            except Exception:
                pass
            kept.append(entry)
        g._saved = kept


class protected_sheets_guard:
    """保護されたシートのあるブックに対する外部COM操作を安全に行う with 文。

    シート保護に加えてブック構造保護（Protect Structure。シートの追加・削除・
    移動・改名を外部COMからもブロックする）も同じ流儀で一時解除→復元する。
    パスワード付きで解除できない場合は諦めて従来どおり（操作は失敗しうる）。

    使い方: with protected_sheets_guard(wb): ...操作...
    """
    def __init__(self, wb):
        self.wb = wb
        self._saved = []
        self._wb_saved = None  # (structure, windows) 解除できた場合のみ

    def __enter__(self):
        try:
            structure = bool(self.wb.ProtectStructure)
            windows = bool(self.wb.ProtectWindows)
        except Exception:
            structure = windows = False
        if structure or windows:
            try:
                # 引数なしの Unprotect はパスワード保護時にダイアログを出して固まる
                # （_unprotect_all_sheets と同じ理由。ダミーで確実に例外化させる）
                self.wb.Unprotect(Password=_UNPROTECT_PROBE_PW)
                self._wb_saved = (structure, windows)
            except Exception as ex:
                print(f"⚠ ブックの構造保護を一時解除できませんでした"
                      f"（パスワード保護の可能性）: {ex}", file=sys.stderr)
                print("  シートの追加・削除・改名を伴う操作は失敗します。", file=sys.stderr)
        self._saved = _unprotect_all_sheets(self.wb)
        _active_guards.append(self)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            _active_guards.remove(self)
        except ValueError:
            pass
        _reprotect_sheets(self._saved)
        if self._wb_saved:
            try:
                self.wb.Protect(Structure=self._wb_saved[0],
                                Windows=self._wb_saved[1])
            except Exception as ex:
                # 従来は完全に無言だった。復元できないと「保護してあるはずのブックの
                # 構造保護が、いつの間にか外れている」状態でコマンドが成功終了する。
                # シート保護の再保護（_reprotect_sheets）は必ず報告するので、対称に揃える
                print(f"⚠ ブックの構造保護を復元できませんでした: {ex}", file=sys.stderr)
                print("  このブックは構造保護が外れたままです。Excel 側で保護し直してください。",
                      file=sys.stderr)
        return False


def protect_safe(cmd_func):
    """cmd_* 関数を「対象ブックの保護シートを一時解除してから実行」に変える decorator。

    対象ブックは他の cmd_* 関数と同じ流儀（posargs先頭がExcelファイルなら
    それを対象、なければアクティブブック）で解決する。get_workbook は
    接続キャッシュを持つため、ここでの解決が二重コストにはならない。
    """
    import functools

    @functools.wraps(cmd_func)
    def wrapper(args):
        try:
            target_file, _rest = parse_target_and_rest(getattr(args, 'posargs', []) or [])
            xl, wb = get_workbook(target_file)
        except Exception:
            # ブック解決に失敗した場合は元の関数にそのまま委ね、
            # そちらのエラーメッセージを出させる
            return cmd_func(args)
        with protected_sheets_guard(wb):
            return cmd_func(args)
    return wrapper


__all__ = [
    'BACKUP_DIR',
    'LAST_PROC_FILE',
    'ModuleNameCollisionError',
    'ReopenFailed',
    'SCRIPT_DIR',
    'XL_EXTS',
    '_LAST_SNAPSHOT_FILE',
    '_LAST_VALUES_FILE',
    '_JOB_CLOCK_FILE',
    '_CALL_LOG_FILE',
    '_CMD_PROGRESS_FILE',
    '_AGENT_PROGRESS_PATH',
    'cmd_progress_write',
    'cmd_progress_note',
    'cmd_progress_read',
    '_pid_alive',
    '_last_excel_pid',
    '_last_book_name',
    '_remember_excel_pid',
    '_note_book',
    'call_log_write',
    'call_log_read',
    'call_log_summary',
    'cmd_stats',
    '_cell_str',
    '_coerce_cell',
    '_col_letter',
    '_created_instances',
    '_created_xl',
    '_created_xl_pid',
    '_last_open_by_tool',
    '_find_component',
    '_find_invalid_procedure_names',
    '_get_active_excel',
    '_get_workbook_uncached',
    '_import_module_verified',
    '_print_collision_guidance',
    '_print_save_failed_guidance',
    '_save_with_retry',
    '_range_values_2d',
    '_reject_extra_args',
    '_remove_export_artifacts',
    '_running_excel_workbooks',
    '_wait_component_gone',
    '_wb_cache',
    '_unprotect_all_sheets',
    '_reprotect_sheets',
    '_active_guards',
    '_com_is_busy',
    '_pid_is_excel',
    '_ws_name',
    'forget_protection',
    'protected_sheets_guard',
    'protect_safe',
    'argparse',
    'check_vba_identifier',
    'cleanup_excel',
    'release_created_instances',
    'get_or_start_excel',
    'pin_active_workbook',
    'pinned_workbook',
    'datetime',
    'get_workbook',
    'job_clock_elapsed',
    'job_clock_get',
    'job_clock_note',
    'job_clock_set',
    'job_clock_start',
    'load_excel_addins_and_personal',
    'looks_like_xl_file',
    'make_backup',
    'make_module_backup',
    'normalize_bas_newlines',
    'os',
    'parse_target_and_rest',
    'pythoncom',
    'pywintypes',
    're',
    'note_if_macro_free_book',   # 2026-09-19（一覧に入れ忘れると import * の側から見えず、open の中では NameError が黙って飲まれた）
    'read_code_file',
    'release_instance',          # 2026-09-17 に足したとき一覧に入れ忘れていた（見張りのテストが 9/19 に見つけた）
    'same_path',
    'setup_encoding',
    'shutil',
    'smart_path_resolve',
    'sys',
    'time',
    'unicodedata',
    'validate_bas_encoding',
    'validate_vba_code',
    'win32com',
    'zlib',
]
