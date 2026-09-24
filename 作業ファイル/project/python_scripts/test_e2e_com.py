"""実機の Excel / VBE を使う E2E テスト（オプトイン）。

  走らせ方:  py -m pytest test_e2e_com.py --run-e2e -q
  既定では   py -m pytest -q          ← 純ロジックだけ（1〜2秒）でスキップされる

■ なぜ正式ファイルなのか（2026-07-14）
E2E がセッションごとの使い捨てスクリプトにしか無く、毎回書き直されていた。
その結果、同じ穴を何度も踏み直し「テストがまた落ちる」を繰り返していた。
ここに置いてある限り、書き直しは起きない。

■ 安全上の約束（過去の事故から）
  - シュウさんの Excel には絶対に触らない。既に Excel が起動していたら skip する
    （理由を名指しして止める。曖昧に落ちない）
  - ブックは毎回ユニークな一時パスに作る（固定パスは前回の残骸と衝突する）
  - _last_proc.vba は共有ファイル。退避して、終わったら必ず戻す
  - Excel は DispatchEx で起こし、必ず畳む（ゾンビを残さない）
"""
import json
import os
import subprocess
import sys
import time

import pytest

pytestmark = pytest.mark.e2e

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBM = os.path.join(SCRIPT_DIR, "vba_manager.py")
LAST_PROC = os.path.join(SCRIPT_DIR, "_last_proc.vba")

MOD_NAME = "TestMod"

MOD_CODE = (
    'Sub 一行完結(): Debug.Print 1: End Sub\r\n'
    'Sub 合計を出す()\r\n'
    '    Range("C1").Value = Range("A1").Value + Range("B1").Value\r\n'
    'End Sub\r\n'
    # --input-text の的: VBA の InputBox()（#32770/Edit）と Excel 内蔵の
    # Application.InputBox（bosa_sdm_XL9/EDTBX）を1本で両方出す
    'Sub 入力箱()\r\n'
    '    Range("D1").Value = InputBox("名前は？", "確認")\r\n'
    '    Range("D2").Value = Application.InputBox("範囲", "範囲", Type:=8).Address(False, False)\r\n'
    'End Sub\r\n'
)

# 閉じたブックを覗いただけで Workbook_Open が走ったら、その痕跡がセルに残る
THISWB_CODE = (
    'Private Sub Workbook_Open()\r\n'
    '    Worksheets("本体").Range("Z1").Value = "OPENED"\r\n'
    'End Sub\r\n'
)


# ================================================================
# 土台
# ================================================================

def _excel_is_running():
    """EXCEL.EXE が動いているか（シュウさんの作業中 Excel を巻き込まないための門番）"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True, encoding="cp932", errors="replace", timeout=30)
        return "EXCEL.EXE" in (out.stdout or "")
    except Exception:
        return False


def _wait_no_excel(limit=90.0):
    """EXCEL.EXE が完全に消えるまで待つ。消えたら True。

    Excel の死は非同期で、Quit を撃ってから実測 2〜3 秒（負荷次第でもっと）かかる。
    死にきる前に次の操作を始めると、vba_manager は GetActiveObject で
    その「まだ死んでいない Excel」に相乗りしてしまい、挙動が変わる。
    これがテストが実行ごとにコケたり通ったりする正体だった（2026-07-14 特定）。
    待ち時間を延ばして誤魔化すのではなく、毎回「無人」を確認してから次へ進む。
    """
    t0 = time.time()
    while time.time() - t0 < limit:
        if not _excel_is_running():
            return True
        time.sleep(0.25)
    return False


@pytest.fixture(scope="session", autouse=True)
def guard_users_excel():
    """作業中の Excel があるなら、理由を言って止まる（曖昧に落ちない）

    見えている Excel を閉じてもこれが出る場合は、非表示・ブック0 の残骸 Excel が
    居座っている（過去に「アドインが効かない」を起こした型）。
    list-open / タスクマネージャで確認して始末してから再実行する。
    """
    if _excel_is_running():
        pytest.skip(
            "Excel が起動しています。開いているブックを巻き込まないため E2E は行いません。\n"
            "  Excel を閉じてから --run-e2e してください。\n"
            "  閉じたはずなのに出る場合は、非表示の残骸 Excel が居ます"
            "（タスクマネージャの EXCEL.EXE を確認）。")
    yield


@pytest.fixture(scope="session", autouse=True)
def preserve_last_proc():
    """_last_proc.vba は MCP サーバー・GUI と共有の受け渡しファイル。退避して必ず戻す"""
    saved = None
    if os.path.exists(LAST_PROC):
        with open(LAST_PROC, "rb") as f:
            saved = f.read()
    yield
    if saved is not None:
        with open(LAST_PROC, "wb") as f:
            f.write(saved)
    elif os.path.exists(LAST_PROC):
        os.remove(LAST_PROC)


@pytest.fixture(scope="session", autouse=True)
def com_apartment():
    """COM の初期化はセッションで1回だけ。CoUninitialize は「しない」。

    2026-07-14 に実測して分かった2つの罠:
      1. CoInitialize / CoUninitialize はスレッドごとの参照カウント。helper のたびに
         対で呼ぶと vbam_core.cleanup_excel() が内部で呼ぶ CoUninitialize と釣り合わず、
         生きた COM 参照を抱えたまま COM を落として**プロセスごと即死**する。
      2. 終了時に自分で CoUninitialize しても、pywin32 側に残った参照との兼ね合いで
         インタプリタ終了時に落ち、全テストが通っているのに**終了コードが 1 になる**。
    COM アパートメントの後始末はプロセス終了時に OS に任せるのが安全。
    """
    import pythoncom
    pythoncom.CoInitialize()
    yield
    # あえて CoUninitialize しない（上記2）。ゾンビ Excel は no_zombie_excel が見張る。


@pytest.fixture(autouse=True)
def no_zombie_excel():
    """各テストの入口で無人を再確認し、出口で Excel を残していないことを確かめる。

    無人確認がセッション開始の1回だけだと、E2E 実行中にシュウさんが Excel を
    開いた場合、以後のテストがその Excel に GetActiveObject で相乗りし、
    さらにこの見張りが本人の Excel を「ゾンビ」と誤認して残り全テスト×最大90秒の
    空回り＋誤失格になる。入口で検知したら巻き込まずに skip する。
    """
    if _excel_is_running():
        pytest.skip("Excel が起動しています（E2E 実行中に開かれた可能性）。"
                    "作業中の Excel を巻き込まないため、このテストは行いません。")
    yield
    if not _wait_no_excel():
        pytest.fail("テスト後に EXCEL.EXE が残り続けています（ゾンビ Excel か、"
                    "テスト中に開かれた作業用 Excel）。後始末を確認してください。")


def _new_excel():
    """新インスタンス（DispatchEx）。COM の初期化はセッション側で済んでいる"""
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


# 後始末の作法（2026-07-14 に実測して分かったこと）:
#   xl.Quit() を撃っても、Python 側が COM 参照を1つでも握っている限り EXCEL.EXE は
#   死なない（プロセス終了時にようやく消える＝テスト中はゾンビに見える）。
#   ヘルパ関数に xl を渡して中で gc しても、呼び出し元の変数がまだ生きているので効かない。
#   参照を持っている当人が None にしてから gc.collect() すること。実測で 2.4 秒で消える。
#   そのため各所で quit → None → gc を「その場で」書く（関数に切り出さない）。


@pytest.fixture
def book(tmp_path):
    """使い捨ての .xlsm を1つ作って渡す（毎回ユニークなパス＝残骸と衝突しない）"""
    import gc
    path = str(tmp_path / "vbam_e2e.xlsm")
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "本体"
        ws.Range("A1").Value = 2
        ws.Range("B1").Value = 3
        comp = wb.VBProject.VBComponents.Add(1)   # 1 = 標準モジュール
        comp.Name = MOD_NAME
        comp.CodeModule.AddFromString(MOD_CODE)
        wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(THISWB_CODE)
        wb.SaveAs(path, FileFormat=52)            # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(SaveChanges=False)
        comp = ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    # 完全に消えるまで待ってから本番へ。生き残った Excel が居ると、
    # vba_manager がそいつに相乗りして挙動が変わる（テストがブレる原因）
    assert _wait_no_excel(), "ブック作成に使った Excel が消えません（後始末が壊れています）"
    yield path


def peek(path):
    """ブックを読み取り専用・イベント無効で覗く（この検査自体が Workbook_Open を起こさない）"""
    import gc
    xl = _new_excel()
    xl.EnableEvents = False
    wb = None
    try:
        wb = xl.Workbooks.Open(path, 0, True)
        modules = sorted(c.Name for c in wb.VBProject.VBComponents)
        cm = wb.VBProject.VBComponents(MOD_NAME).CodeModule
        code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        ws = wb.Worksheets("本体")
        result = {
            "modules": modules,
            "code": code,
            "Z1": ws.Range("Z1").Value,
            "C1": ws.Range("C1").Value,
            "D1": ws.Range("D1").Value,
            "D2": ws.Range("D2").Value,
        }
        cm = ws = None
        wb.Close(SaveChanges=False)
        wb = None
        return result
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()   # 覗きに使った Excel が消えるまで待つ（次の操作を汚さない）


def run_cli(*args, timeout=180):
    """vba_manager.py を別プロセスで叩く（実際の使われ方に一番近い）"""
    p = subprocess.run([sys.executable, VBM, *args], capture_output=True,
                       encoding="utf-8", errors="replace", cwd=SCRIPT_DIR, timeout=timeout)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


# ================================================================
# 本命: 「Import 成功後に保存が失敗しても、連番モジュールを作らない」
# 2026-07-14 に発見・修正した実害の回帰。実機の VBE で確かめる。
# ================================================================

# この検証だけは「別プロセス」で走らせる。
# vbam_core を pytest と同じプロセスで動かすと、終了時の COM 後始末とかち合って
# インタプリタが落ち（全テストが通っているのに終了コード 1 になる）、
# さらに Excel の COM 参照が居残ってゾンビ扱いになる（2026-07-14 実測）。
# プロセスを分ければ COM はプロセス終了で確実に片付く。実際の使われ方にも近い。
_PROBE = '''# -*- coding: utf-8 -*-
import argparse, json, sys
sys.path.insert(0, r"{scripts}")
import vbam_core, vbam_vba

# vba_manager.py の入口と同じく stdout/stderr を UTF-8 にする。
# これを忘れると cp932 のままになり、案内文の「⚠」（cp932 に無い）で
# UnicodeEncodeError になる。ライブラリとして import するときの作法。
vbam_core.setup_encoding()

book, new_bas, mod = sys.argv[1], sys.argv[2], sys.argv[3]

# Import が終わった直後の保存だけを確実に失敗させる（ディスク要因の保存失敗を模す）
def boom(wb, *a, **kw):
    raise RuntimeError("模擬: 保存に失敗しました")
vbam_vba._save_with_retry = boom

try:
    args = argparse.Namespace(posargs=[book, mod, new_bas], yes=True, force=False)
    ok = vbam_vba.cmd_replace_module(args)

    # 開いているブックの中身をそのまま見る（保存されていないのでファイルには出ない）
    xl, wb = vbam_core.get_workbook(book)
    names = sorted(c.Name for c in wb.VBProject.VBComponents)
    cm = wb.VBProject.VBComponents(mod).CodeModule
    code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
    cm = wb = xl = None

    print("__RESULT__" + json.dumps(
        {{"ok": bool(ok), "modules": names, "has_new_code": "999" in code}},
        ensure_ascii=False))
finally:
    # 起こした Excel を PID ごと確実に始末する（非表示・ブック0 の Excel を残さない）。
    # 未保存のまま破棄＝ディスクには一切触らせない。
    # cleanup_excel は内部で CoUninitialize するが、この使い捨てプロセスは
    # ここで終わるので問題ない（pytest と同じプロセスで呼ぶと落ちる）
    try:
        vbam_core._wb_cache.clear()
    except Exception:
        pass
    vbam_core.cleanup_excel()
'''


def test_save_failure_after_import_does_not_create_numbered_module(book, tmp_path):
    """保存だけ失敗したとき、バックアップを再 Import して連番モジュールを産まないこと。

    旧コードは except が removed フラグしか見ておらず、Import 成功後の Save 失敗を
    「モジュールが消えた」と誤認してバックアップを重ね Import していた。正しい名前の
    モジュールが既に在るので必ず衝突し、ツール自身が「旧コード入りの TestMod1」を
    産んで 35 秒待たせた末に失敗を告げていた（実機の VBE で連番化を再現済み）。

    ここでは Import が終わった直後の保存だけを確実に失敗させて、その分岐を撃つ。
    """
    new_bas = str(tmp_path / "new_mod.bas")
    new_code = (
        f'Attribute VB_Name = "{MOD_NAME}"\r\n'
        'Sub 一行完結(): Debug.Print 1: End Sub\r\n'
        'Sub 合計を出す()\r\n'
        '    Range("C1").Value = 999\r\n'   # 置換後だと分かる印
        'End Sub\r\n'
    )
    with open(new_bas, "wb") as f:
        f.write(new_code.encode("cp932"))

    probe = str(tmp_path / "probe_save_fail.py")
    with open(probe, "w", encoding="utf-8") as f:
        f.write(_PROBE.format(scripts=SCRIPT_DIR))

    p = subprocess.run([sys.executable, probe, book, new_bas, MOD_NAME],
                       capture_output=True, encoding="utf-8", errors="replace",
                       cwd=SCRIPT_DIR, timeout=300)
    out = (p.stdout or "") + (p.stderr or "")
    line = next((l for l in out.splitlines() if l.startswith("__RESULT__")), None)
    assert line, f"検証プロセスが結果を返さなかった:\n{out}"
    result = __import__("json").loads(line[len("__RESULT__"):])

    assert result["ok"] is False, f"保存に失敗したのに成功と報告している:\n{out}"

    names = result["modules"]
    numbered = [n for n in names
                if n.lower() != MOD_NAME.lower() and n.lower().startswith(MOD_NAME.lower())]
    assert not numbered, (
        f"連番モジュールが生まれている: {numbered}\n"
        "  Import 成功後の Save 失敗で、ツールがバックアップを再 Import している"
        f"（imported フラグのガードが外れた）。\n{out}")

    # 置換自体は済んでいる＝「消えた」のではなく「保存できなかった」だけ
    assert MOD_NAME in names, f"モジュールが消えている:\n{out}"
    assert result["has_new_code"], f"取り込みは成功しているはずなのに新コードが入っていない:\n{out}"
    # 誤誘導せず「保存だけ失敗した」と伝えていること
    assert "保存" in out


# ================================================================
# 回帰: これまでの修正を壊していないこと
# ================================================================

def test_list_does_not_trigger_workbook_open(book):
    """読み取り系（list）が、閉じたブックの Workbook_Open を起こさないこと"""
    rc, out = run_cli("list", book)
    assert rc == 0, out
    assert peek(book)["Z1"] is None, "list で Workbook_Open が走っている（健診モードが壊れた）"


def test_run_macro_executes_target_macro(book):
    """run-macro が対象ブックのマクロを実際に動かすこと"""
    batch = os.path.join(os.path.dirname(book), "_batch.txt")
    with open(batch, "w", encoding="utf-8") as f:
        f.write(f'run-macro "{book}" 合計を出す\n')
        f.write(f'save "{book}"\n')
    rc, out = run_cli("batch", batch)
    assert rc == 0, out
    assert peek(book)["C1"] == 5, "A1(2)+B1(3)=5 になっていない"


def test_run_macro_input_text_answers_inputbox(book):
    """--input-text が InputBox に値を入れて確定し、マクロを先へ進ませること（2026-08-23）。

    VBA の InputBox()（#32770 / Edit）と Excel 内蔵の Application.InputBox
    （bosa_sdm_XL9 / EDTBX）の両方に、出た順で値が入ること。
    入れた値は報告に出ること（黙って入れない）。
    """
    batch = os.path.join(os.path.dirname(book), "_batch_input.txt")
    with open(batch, "w", encoding="utf-8") as f:
        f.write(f'run-macro "{book}" 入力箱 --input-text 太郎 --input-text B2:C3\n')
        f.write(f'save "{book}"\n')
    rc, out = run_cli("batch", batch)
    assert rc == 0, out
    assert "「太郎」" in out and "「B2:C3」" in out, f"入れた値が報告に出ていない:\n{out}"
    got = peek(book)
    assert got["D1"] == "太郎", f"InputBox() に値が入っていない: {got['D1']!r}"
    assert got["D2"] == "B2:C3", f"Application.InputBox に値が入っていない: {got['D2']!r}"


def test_replace_procedure_rejects_two_procedures(book):
    """コードファイルに2本入っていたら弾くこと（黙って両方入れない）"""
    with open(LAST_PROC, "w", encoding="utf-8") as f:
        f.write('Sub 一行完結(): Debug.Print 99: End Sub\n'
                'Sub 合計を出す()\n    Debug.Print 2\nEnd Sub\n')
    rc, out = run_cli("replace-procedure", book, "-y")
    assert rc == 1, f"2本入りを通してしまった: {out}"
    assert "2 本のプロシージャ" in out, out


def test_replace_procedure_keeps_other_procedures(book):
    """1行完結 Sub を巻き込んで消さないこと（過去の実害の回帰）"""
    with open(LAST_PROC, "w", encoding="utf-8") as f:
        f.write('Sub 合計を出す()\n    Range("C1").Value = 7\nEnd Sub\n')
    rc, out = run_cli("replace-procedure", book, "-y")
    assert rc == 0, out
    code = peek(book)["code"]
    assert "Sub 一行完結()" in code, "隣の1行完結 Sub が消えている"
    assert "Sub 合計を出す()" in code
    assert "7" in code


def test_read_only_commands_work(book):
    """grep / list-modules / list-shortcuts が動くこと（健診モード化の回帰）"""
    rc1, _ = run_cli("grep", book, "Debug.Print")
    rc2, out2 = run_cli("list-modules", book)
    rc3, _ = run_cli("list-shortcuts", book)
    assert (rc1, rc2, rc3) == (0, 0, 0)
    assert MOD_NAME in out2


def test_snapshot_and_diff(book, tmp_path):
    """snapshot / snapshot-diff が動き、変更が無ければ差分なしと言うこと"""
    s1 = str(tmp_path / "s1.json")
    s2 = str(tmp_path / "s2.json")
    rc1, _ = run_cli("snapshot", book, "--out", s1)
    rc2, _ = run_cli("snapshot", book, "--out", s2)
    rc3, out3 = run_cli("snapshot-diff", s1, s2)
    assert (rc1, rc2, rc3) == (0, 0, 0)
    assert "差分なし" in out3, out3


def test_delete_module_removes_it(book):
    """delete-module が実際に消し、保存すること"""
    import gc
    rc, out = run_cli("delete-module", book, MOD_NAME, "-y")
    assert rc == 0, out
    xl = _new_excel()
    wb = None
    try:
        wb = xl.Workbooks.Open(book, 0, True)
        names = [c.Name for c in wb.VBProject.VBComponents]
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()

    assert MOD_NAME not in names, "delete-module が「削除完了」と言ったのに残っている"


# ================================================================
# 2026-07-15 追加: rehearse（予行演習run）と open / close
# ================================================================

def test_rehearse_runs_macro_on_copy_only(book, tmp_path):
    """rehearse がコピーだけを撃ち、本体には指一本触れないこと。

    確かめる約束:
      - 本体は無傷（C1 空のまま・Workbook_Open も起きていない）
      - 結果コピーには A1(2)+B1(3)=5 が入っている
      - コピーを開く瞬間も Workbook_Open を起こしていない（Z1 が空）
      - 差分レポートに C1 の変化が事実として出る
    """
    copy_path = str(tmp_path / "rehearsal.xlsm")
    rc, out = run_cli("rehearse", book, "合計を出す", "--out", copy_path)
    assert rc == 0, out
    assert "マクロ実行: 成功" in out, out
    assert "C1" in out, f"差分レポートに C1 の変化が出ていない:\n{out}"

    orig = peek(book)
    assert orig["C1"] is None, "本体が書き換わっている（予行演習が本体を撃った）"
    assert orig["Z1"] is None, "本体の Workbook_Open が走っている"

    assert os.path.exists(copy_path), "結果コピーが残っていない"
    res = peek(copy_path)
    assert res["C1"] == 5, "結果コピーにマクロの効果が入っていない"
    assert res["Z1"] is None, "コピーを開く瞬間に Workbook_Open が走っている"
    assert os.path.exists(copy_path + ".before.json")
    assert os.path.exists(copy_path + ".after.json")


# ================================================================
# 2026-08-23 追加: gate（関所＝停止条件）・--timeout・rehearse のハーネス
# ================================================================

GATE_MOD = "GateMod"
GATE_CODE = (
    'Sub テスト成功()\r\n'
    '    If 1 + 1 <> 2 Then Err.Raise 5, , "算数が壊れた"\r\n'
    'End Sub\r\n'
    'Sub テスト失敗()\r\n'
    '    Err.Raise 5, , "期待3 実際=2"\r\n'
    'End Sub\r\n'
    'Sub テスト無限()\r\n'
    '    Do\r\n'
    '    Loop\r\n'
    'End Sub\r\n'
    'Sub 落ちる()\r\n'
    '    Err.Raise 5, , "わざと落とす"\r\n'
    'End Sub\r\n'
)


@pytest.fixture
def gate_book(tmp_path):
    """gate / rehearse 用の使い捨て .xlsm（テスト Sub 入り・毎回ユニークなパス）"""
    import gc
    path = str(tmp_path / "vbam_gate.xlsm")
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Add()
        wb.Worksheets(1).Name = "本体"
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = GATE_MOD
        comp.CodeModule.AddFromString(GATE_CODE)
        wb.SaveAs(path, FileFormat=52)
        wb.Close(SaveChanges=False)
        comp = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    assert _wait_no_excel(), "ブック作成に使った Excel が消えません（後始末が壊れています）"
    yield path


def test_gate_passes_when_filtered_tests_all_succeed(gate_book):
    rc, out = run_cli("gate", gate_book, "成功")
    assert rc == 0, out
    assert "判定: 通過（PASS）" in out, out
    assert "テスト 1/1 成功" in out, out


def test_gate_fails_on_failing_test_and_reports_reason_in_json(gate_book):
    rc, out = run_cli("gate", gate_book, "失敗", "--json")
    assert rc == 1, out
    doc = json.loads([l for l in out.splitlines() if l.startswith("{")][-1])
    assert doc["pass"] is False
    assert doc["tests"]["failed"][0]["name"] == "テスト失敗"
    assert "実行時エラー 5" in doc["tests"]["failed"][0]["error"]
    assert any("NG" in r for r in doc["reasons"])


def test_gate_fails_when_no_tests_exist(book):
    # MOD_CODE にはテスト Sub が無い＝「合否を言えない」ものは通さない（テスト 0 本は FAIL）
    rc, out = run_cli("gate", book)
    assert rc == 1, out
    assert "テストが 0 本" in out, out


def test_gate_timeout_kills_only_its_own_excel(gate_book):
    rc, out = run_cli("gate", gate_book, "無限", "--timeout", "4")
    assert rc == 1, out
    assert "時間切れ" in out and "強制終了" in out, out
    assert _wait_no_excel(), "時間切れ後に EXCEL.EXE が残っています（演習用 Excel の始末漏れ）"


def test_rehearse_reports_runtime_error_via_harness(gate_book):
    rc, out = run_cli("rehearse", gate_book, "落ちる", "--discard")
    assert rc == 1, out
    assert "マクロ実行: 失敗  実行時エラー 5: わざと落とす" in out, out


def test_rehearse_timeout_kills_only_its_own_excel(gate_book):
    rc, out = run_cli("rehearse", gate_book, "テスト無限", "--timeout", "4", "--discard")
    assert rc == 1, out
    assert "時間切れ" in out and "強制終了" in out, out
    assert _wait_no_excel(), "時間切れ後に EXCEL.EXE が残っています（演習用 Excel の始末漏れ）"


def test_open_and_close_on_running_instance(book):
    """open の二重オープン防止と、close の鎧（-y 必須・--no-save で破棄）を実機で確かめる"""
    import gc
    xl = _new_excel()
    wb = None
    try:
        wb = xl.Workbooks.Open(book)

        # open: 既に開いているブックは二重に開かず、非表示に居ることを警告する
        rc, out = run_cli("open", book)
        assert rc == 0, out
        assert "既に開いています" in out, out
        assert "非表示" in out, out
        assert xl.Workbooks.Count == 1

        # close: -y なしの非対話は鎧③（確認プロンプト）で止まる
        p = subprocess.run(
            [sys.executable, VBM, "close", os.path.basename(book), "--no-save"],
            capture_output=True, encoding="utf-8", errors="replace",
            cwd=SCRIPT_DIR, timeout=180, stdin=subprocess.DEVNULL)
        out = (p.stdout or "") + (p.stderr or "")
        assert p.returncode != 0, f"確認なしで閉じてしまった:\n{out}"
        assert "非対話" in out, out
        assert xl.Workbooks.Count == 1, "確認で止まったはずのブックが閉じられている"

        # close: 名指し＋--no-save＋-y。未保存の変更は破棄され、ディスクは無傷
        wb.Worksheets("本体").Range("C1").Value = 123   # 未保存の変更
        rc, out = run_cli("close", os.path.basename(book), "--no-save", "-y")
        assert rc == 0, out
        assert "閉じました" in out, out
        assert "破棄" in out, "未保存の変更を破棄する警告が出ていない"
        wb = None
        assert xl.Workbooks.Count == 0, "ブックが閉じられていない"
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()

    assert peek(book)["C1"] is None, "--no-save なのに変更がディスクへ保存されている"


# ================================================================
# tidy が既存の表示形式を壊さないこと（2026-09-08 の回帰）
#
# 見出しの書式を複写する xlPasteFormats は表示形式まで複写する。集計ブロック
# （左が「合計」・右が 203,250）や年を見出しにした表に当てると #,##0 が General に戻り、
# カンマの消えた見た目になる。仕上げ検査も見出し行はデータとして見ておらず、
# 「桁区切り抜け」を素通りさせていた（本番の走行で実際に通った）。
# ================================================================

def _tidy_probe_sheet(ws):
    """集計ブロック（先頭行が数値）と、年を見出しにした表を 1 枚に置く"""
    ws.Range("F27").Value = "合計"
    ws.Range("F27").Font.Bold = True
    ws.Range("G27").Value = 203250
    ws.Range("G27").NumberFormat = "#,##0"
    ws.Range("F28").Value = "平均"
    ws.Range("F28").Font.Bold = True
    ws.Range("G28").Value = 10162.5
    ws.Range("G28").NumberFormat = "#,##0.0"
    ws.Range("A1").Value = "地域"
    ws.Range("A1").Font.Bold = True
    for col, year in (("B", 2025), ("C", 2026)):
        ws.Range(f"{col}1").Value = year
        ws.Range(f"{col}1").NumberFormat = "0"
    ws.Range("A2").Value = "架空県北"
    ws.Range("B2").Value = 1234
    ws.Range("C2").Value = 5678
    ws.Range("A3").Value = "架空県南"
    ws.Range("B3").Value = 2345
    ws.Range("C3").Value = 6789


def test_tidy_keeps_number_format_of_numeric_header_cells(book):
    """tidy の見出し複写が、先頭行の数値セルの表示形式を潰さないこと"""
    import argparse
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_core import pinned_workbook
    from vbam_edit import cmd_tidy

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets("本体")
        _tidy_probe_sheet(ws)
        ns = argparse.Namespace(posargs=["F27:G28", "A1:C3"], header_from=None, bg=None,
                                sheet_opt="本体", no_header=False, no_border=False,
                                no_col_format=False, no_autofit=False, min_width=None, max_width=None)
        with pinned_workbook(wb):
            assert cmd_tidy(ns)
        assert ws.Range("G27").NumberFormat == "#,##0", "合計セルの桁区切りが消えた"
        assert ws.Range("G28").NumberFormat == "#,##0.0"
        assert ws.Range("B1").NumberFormat == "0", "年の見出しの表示形式が消えた"
        assert ws.Range("C1").NumberFormat == "0"
        # 見出しらしさ（太字）は当たっている＝表示形式を守るために整形をやめてはいない
        assert ws.Range("B1").Font.Bold is True
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_tidy_money_stays_whole_and_blank_rows_get_no_border_20260913(book):
    """お試し版テスト用1 の形: 明細 20 行＋空の行＋合計・平均（F:G だけ）。

    前は平均 10,162.5 の 1 つで金額の列が #,##0.00 になり、空の行と合計欄の左にも罫線が付いた。
    """
    import argparse
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_core import pinned_workbook
    from vbam_edit import cmd_tidy

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets.Add()
        ws.Name = "明細"
        for j, h in enumerate(["日付", "商品", "分類", "担当", "数量", "単価", "金額"], start=1):
            ws.Cells(5, j).Value = h
        ws.Range("A5:G5").Font.Bold = True
        for i in range(20):
            r = 6 + i
            ws.Cells(r, 1).Formula = f"2026/8/{i + 1}"
            ws.Cells(r, 2).Value = "えんぴつ"
            ws.Cells(r, 3).Value = "筆記具"
            ws.Cells(r, 4).Value = "佐藤"
            ws.Cells(r, 5).Value = 100 + i
            ws.Cells(r, 6).Value = 60
            ws.Cells(r, 7).Formula = f"=E{r}*F{r}"
        ws.Range("F27").Value = "合計"
        ws.Range("G27").Formula = "=SUM(G6:G25)"
        ws.Range("F28").Value = "平均"
        ws.Range("G28").Formula = "=AVERAGE(G6:G25)"            # 6,570 … 小数が出る数に
        ws.Range("E25").Value = 121
        ns = argparse.Namespace(posargs=["A5:G28"], header_from=None, bg=None,
                                sheet_opt="明細", no_header=False, no_border=False,
                                no_col_format=False, no_autofit=False, min_width=None, max_width=None)
        with pinned_workbook(wb):
            assert cmd_tidy(ns)
        assert ws.Range("G6").NumberFormat == "#,##0", "金額が小数の形になった"
        assert ws.Range("G28").NumberFormat == "#,##0.0#", "平均の小数が隠れた"
        assert ws.Range("A26").Borders(7).LineStyle == -4142, "空の行に罫線が付いた"      # xlEdgeLeft
        assert ws.Range("E27").Borders(7).LineStyle == -4142, "合計欄の左の空欄に罫線が付いた"
        assert ws.Range("F27").Borders(7).LineStyle == 1, "合計欄に罫線が無い"
        assert ws.Range("A25").Borders(9).LineStyle == 1, "明細の下端に罫線が無い"          # xlEdgeBottom
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_audit_table_reports_comma_gap_in_numeric_header_row(book):
    """仕上げ検査が「先頭行の数値だけ桁区切りが無い」を見つけ、直った表では黙ること"""
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_edit import audit_table

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets("本体")
        ws.Range("F27").Value = "合計"
        ws.Range("F27").Font.Bold = True
        ws.Range("G27").Value = 203250                     # 表示形式なし＝壊れた後の姿
        ws.Range("F28").Value = "平均"
        ws.Range("F28").Font.Bold = True
        ws.Range("G28").Value = 10162.5
        ws.Range("G28").NumberFormat = "#,##0.0"
        ws.Range("F27:G28").Borders.LineStyle = 1
        bad = audit_table(ws, ws.Range("F27:G28"))
        assert any("桁区切り" in s and "G27" in s for s in bad), f"見逃した: {bad}"
        ws.Range("G27").NumberFormat = "#,##0"
        assert not audit_table(ws, ws.Range("F27:G28")), "直っている表に指摘が出た"
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_harvest_trims_big_sheets_instead_of_leaving_real_data(tmp_path):
    """写し取り（--harvest）が、大きいシートを実データのまま複製しないこと（2026-09-08 の回帰）。

    直す前は _INV_MAX_CELLS を超えるシートを「見送り」して素通りさせ、複製に本物の値が残ったまま
    「元の中身が残ったセルは 0 個」と報告していた（渡せない物を渡せると言っていた）。
    """
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    import vbam_agent as va

    secret = "本物の氏名"
    src = str(tmp_path / "本物.xlsx")
    n = 400
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "明細"
        ws.Range("A1:C1").Value = (("氏名", "住所", "金額"),)
        ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 3)).Value = tuple(
            (f"{secret}{i}", f"架空県{i}丁目", 1000 + i) for i in range(1, n + 1))
        wb.SaveAs(src, FileFormat=51)
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    assert _wait_no_excel()

    out = str(tmp_path / "練習台.xlsx")
    keep = va._INV_MAX_CELLS
    va._INV_MAX_CELLS = 300                     # 70,000 行を作らずに「大きすぎる」道を通す
    try:
        assert va.harvest_book(src, out)
    finally:
        va._INV_MAX_CELLS = keep

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(out)
        ws = wb.Worksheets("明細")
        ur = ws.UsedRange
        left = []
        for r in range(1, int(ur.Rows.Count) + 1):
            for c in range(1, int(ur.Columns.Count) + 1):
                v = ws.Cells(r, c).Value
                if isinstance(v, str) and secret in v:
                    left.append(str(ws.Cells(r, c).Address).replace('$', ''))
        assert int(ur.Rows.Count) < n, "切り詰められていない"
        assert not left, f"実データが残っている: {left[:5]}"
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_autofilter_clears_a_previous_filter_instead_of_doing_nothing(book):
    """列を書かない autofilter が、前の絞り込みを残したまま黙って通らないこと（2026-09-08 の回帰）。

    実射「鈴木さんの分だけ」が、前の人の絞り込みが効いたまま合格していた（道具が ok を返し、
    手が全部通った回として done まで通った）。かといって失敗で突き返すと、「解除したい」ときに
    同じ手を繰り返して往復を使い切る（実射「テーブル化」）。絞り込みは見せ方＝解いて報告する。
    """
    import argparse
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_core import pinned_workbook
    from vbam_edit import cmd_autofilter

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets("本体")
        ws.Cells.Clear()
        ws.Range("A1:C1").Value = (("会員番号", "氏名", "担当"),)
        ws.Range("A2:C5").Value = (("0001", "青木 誠", "佐藤"), ("0002", "石川 恵", "鈴木"),
                                   ("0003", "上田 学", "佐藤"), ("0004", "遠藤 香", "高橋"))
        ws.Range("A1:C5").AutoFilter(3, "佐藤")          # 前の人の絞り込み

        def shoot():
            ns = argparse.Namespace(posargs=["A1:C5"], sheet_opt="本体", column=None,
                                    equals=None, show_all=False)
            with pinned_workbook(wb):
                return cmd_autofilter(ns)

        assert shoot() is True
        assert not any(ws.Rows(r).Hidden for r in range(2, 6)), "前の絞り込みが残ったまま"
        assert shoot() is True, "▼があるだけの回で落ちている"
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_row_height_keeps_manually_hidden_rows_hidden(book):
    """行高を入れても、人が手で隠した行を表に出さないこと（2026-09-08 の回帰）。

    隠す＝高さ 0 なので、範囲に行高を入れると隠れた行が見えるようになる。実射の手順書
    「見た目を直す」が「頼んでいない変化: 隠れていた行が変わった（4・5 行目）」で落ちていた。
    """
    import argparse
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_core import pinned_workbook
    from vbam_edit import cmd_format_range

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets("本体")
        ws.Cells.Clear()
        for i in range(1, 8):
            ws.Cells(i, 1).Value = f"行{i}"
        ws.Rows("4:5").Hidden = True
        ns = argparse.Namespace(posargs=["A1:A7"], sheet_opt="本体", row_height=18,
                                col_width=None, border=None, merge=False, unmerge=False,
                                align=None, number_format=None, autofit=False, bold=False,
                                bg=None, wrap=False, show=False)
        with pinned_workbook(wb):
            cmd_format_range(ns)
        assert [r for r in range(1, 8) if ws.Rows(r).Hidden] == [4, 5], "隠した行が表に出た"
        assert abs(float(ws.Rows(6).RowHeight) - 18) < 0.6, "行高が入っていない"
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


def test_sort_shows_hidden_rows_so_they_are_not_left_behind(book):
    """隠れた行があるとき、並べ替えがその行を置き去りにしないこと（2026-09-08 の回帰）。

    Excel の並べ替えは見えている行しか動かさない。実射「多い順に並べて」が、手で隠した 2 行を
    末尾に残したまま「合格」で通っていた（仕上げ検査も終わりの検査も見つけられない）。
    """
    import argparse
    import gc

    sys.path.insert(0, SCRIPT_DIR)
    from vbam_core import pinned_workbook
    from vbam_edit import cmd_sort

    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        ws = wb.Worksheets("本体")
        ws.Cells.Clear()
        ws.Range("A1:B1").Value = (("氏名", "金額"),)
        ws.Range("A2:B9").Value = (("青木 誠", 1250000), ("石川 恵", 980000), ("上田 学", 1430000),
                                   ("遠藤 香", 760000), ("大西 亮", 2100000), ("加藤 環", 540000),
                                   ("木村 悠", 1875000), ("小島 望", 320000))
        ws.Rows("8:9").Hidden = True                    # 人が手で隠した 2 行
        ns = argparse.Namespace(posargs=["A1:B9"], sheet_opt="本体", key="B", desc=True,
                                header=True, no_header=False, whole_sheet=False)
        with pinned_workbook(wb):
            assert cmd_sort(ns)
        got = [ws.Cells(r, 2).Value for r in range(2, 10)]
        assert got == sorted(got, reverse=True), f"隠れた行が取り残された: {got}"
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()


# ================================================================
# agent の往復を本物の Excel で再生する（2026-09-10 夜）
#
# test_tools.py は偽のブックでループの判断を見る。ここでは記録した AI の返事をそのまま再生して
# （AI は呼ばない＝課金ゼロ）、本物の Excel で「材料を読む → 控えを取る → 手を実行する → 関所 → 報告」
# まで回し、最後に --undo で控えの姿へ戻るところまで確かめる。
# 道具が自分で書く状態ファイル（台帳・控え・記録・覚書・時計・画像）は全部一時フォルダへ逃がす
# （本物の台帳は「実運用の量」の数字＝E2E の走行が 1 行でも混ざると嘘になる）。
# 別プロセスで回す理由は _PROBE と同じ（pytest と同じプロセスで COM を畳むと落ちる）。
# ================================================================

_REPLAY_PROBE = '''# -*- coding: utf-8 -*-
import argparse, json, os, sys, types
sys.path.insert(0, r"{scripts}")
import vbam_core
vbam_core.setup_encoding()
import vba_manager, vbam_agent as va          # _run_cmd が後から読むコマンド表も、逃がす前に読み込んでおく

book, src, state = sys.argv[1], sys.argv[2], sys.argv[3]
real = vbam_core.SCRIPT_DIR
backups = os.path.join(state, "backups")
for name, mod in list(sys.modules.items()):
    if mod is None or not name.startswith(("vbam_", "vba_manager")):
        continue
    for key, val in list(vars(mod).items()):
        if key == "BACKUP_DIR":
            setattr(mod, key, backups)
        elif isinstance(val, str) and val.startswith(real):
            setattr(mod, key, state + val[len(real):])
        elif isinstance(val, types.FunctionType) and val.__defaults__:
            val.__defaults__ = tuple(state + d[len(real):] if isinstance(d, str) and d.startswith(real) else d
                                     for d in val.__defaults__)

out = {{}}
try:
    args = argparse.Namespace(sheet_opt=None, max_turns=None, dry_run=False, image=False, no_grade=True)
    out["ok"] = bool(va._cmd_agent_replay(args, book, [], src))
    xl, wb = vbam_core.get_workbook(book)
    ws = wb.Worksheets("名簿")
    out["after"] = {{"A6": ws.Range("A6").Value, "C6f": ws.Range("C6").Formula, "C6": ws.Range("C6").Value,
                    "C6fmt": ws.Range("C6").NumberFormat, "A1bold": bool(ws.Range("A1").Font.Bold),
                    "inner": ws.Range("A1:C6").Borders(12).LineStyle}}
    out["undo"] = bool(va.undo_agent(book))
    ws = wb.Worksheets("名簿")
    out["back"] = {{"A6": ws.Range("A6").Value, "C6f": ws.Range("C6").Formula,
                   "A1bold": bool(ws.Range("A1").Font.Bold), "C2": ws.Range("C2").Value}}
    out["saved"] = bool(wb.Saved)
    with open(va._AGENT_RUNS_FILE, encoding="utf-8") as f:
        out["runs"] = [json.loads(l) for l in f if l.strip()]
    out["backups"] = sorted(os.listdir(backups))
    ws = wb = xl = None
finally:
    print("__RESULT__" + json.dumps(out, ensure_ascii=False, default=str))
    try:
        vbam_core._wb_cache.clear()
    except Exception:
        pass
    vbam_core.cleanup_excel()
'''


def test_agent_replay_runs_the_whole_loop_in_real_excel_and_undo_restores_it(tmp_path):
    """記録した返事の再生で、本物の Excel が 1 周する（AI ゼロ）。書いた物は --undo で控えの姿に戻る。"""
    import gc
    import hashlib
    book = str(tmp_path / "e2e_meibo.xlsx")
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "名簿"
        ws.Range("A1:C5").Value = (("氏名", "区分", "金額"), ("佐藤", "正", 1200), ("鈴木", "準", 980),
                                   ("田中", "正", 1500), ("高橋", "準", 700))
        wb.SaveAs(book, FileFormat=51)                 # 51 = xlOpenXMLWorkbook
        wb.Close(SaveChanges=False)
        ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    assert _wait_no_excel(), "ブック作成に使った Excel が消えません（後始末が壊れています）"
    with open(book, "rb") as f:
        digest = hashlib.sha1(f.read()).hexdigest()

    plan = [{"item": "表の下に合計行を足す", "state": "未"}, {"item": "表の書き方と罫線と列幅をそろえる", "state": "未"}]
    first = {"say": "合計行を書いて整えます", "plan": plan,
             "actions": [{"op": "write_cells", "cells": {"A6": "合計", "C6": "=SUM(C2:C5)"}},
                         {"op": "tidy", "ranges": ["A1:C6"]}], "done": False}
    done = {"say": "終わりました", "plan": [dict(p, state="済") for p in plan], "actions": [], "done": True,
            "report": "【やったこと】\n・A6 に「合計」、C6 に =SUM(C2:C5) を書いた\n・A1:C6 に tidy を当てた\n"
                      "【できなかったこと】\nなし"}
    src = tmp_path / "record.jsonl"
    with open(src, "w", encoding="utf-8") as f:
        f.write(json.dumps({"meta": {"book": "e2e_meibo.xlsx", "sheet": "名簿", "mode": "sheet",
                                     "request": "金額の合計行を足して、表を整えて", "ai": "gemini",
                                     "model": "replay", "max_turns": 6}}, ensure_ascii=False) + "\n")
        # 関所に 2 回まで差し戻されても尽きないように、done の返事を 3 本並べる
        for i, reply in enumerate([first, done, done, done], 1):
            f.write(json.dumps({"turn": i, "prompt": "", "reply": json.dumps(reply, ensure_ascii=False)},
                               ensure_ascii=False) + "\n")

    watched = [os.path.join(SCRIPT_DIR, n) for n in ("_agent_runs.jsonl", "_last_agent_undo.json", "_agent_notes.json",
                                                     "_last_agent_log.jsonl", "_agent_progress.json",
                                                     "_last_agent_changes.tsv", "_job_clock.json")]

    def stamp():
        return {p: (os.path.getsize(p), os.path.getmtime(p)) if os.path.exists(p) else None for p in watched}
    real_backups = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "backups"))

    def mine():
        return sorted(n for n in os.listdir(real_backups) if n.startswith("e2e_meibo")) \
            if os.path.isdir(real_backups) else []
    before, mine_before = stamp(), mine()

    state = tmp_path / "state"
    state.mkdir()
    probe = tmp_path / "probe_replay.py"
    probe.write_text(_REPLAY_PROBE.format(scripts=SCRIPT_DIR), encoding="utf-8")
    p = subprocess.run([sys.executable, str(probe), book, str(src), str(state)], capture_output=True,
                       encoding="utf-8", errors="replace", cwd=SCRIPT_DIR, timeout=300)
    out = (p.stdout or "") + (p.stderr or "")
    line = next((l for l in out.splitlines() if l.startswith("__RESULT__")), None)
    assert line, f"検証プロセスが結果を返さなかった:\n{out}"
    r = json.loads(line[len("__RESULT__"):])
    assert r.get("ok") is True, f"再生した往復が合格しなかった:\n{out}"
    after = r["after"]
    assert (after["A6"], after["C6f"], after["C6"]) == ("合計", "=SUM(C2:C5)", 4380), out
    assert "#,##0" in after["C6fmt"] and after["A1bold"] is True and after["inner"] == 1, out   # tidy が当たった
    assert r["undo"] is True, out
    back = r["back"]
    assert (back["A6"], back["C6f"], back["A1bold"], back["C2"]) == (None, "", False, 1200), out
    assert r["saved"] is False                          # 道具は保存しない（戻したのも Excel の中だけ）
    assert [x.get("mode") for x in r["runs"] if not x.get("undo")] == ["sheet"], out
    assert any(x.get("undo") for x in r["runs"]), "台帳に「人が戻した」の印が無い"
    assert any(n.endswith(".xlsx") for n in r["backups"]), out
    assert "記録が尽きました" not in out
    # ファイルは 1 バイトも変わっていない・本物の台帳や覚書や控え置き場も汚していない
    with open(book, "rb") as f:
        assert hashlib.sha1(f.read()).hexdigest() == digest
    now = stamp()
    assert now == before, f"本物の状態ファイルが書き換わった: {[q for q in watched if now[q] != before[q]]}"
    assert mine() == mine_before, "本物の控え置き場に E2E の控えが残った"


def test_compile_names_the_failing_line(book):
    """全体コンパイルが落ちた行を名指しする（2026-09-17）。ラベル無しの On Error GoTo を仕込んで押す。

    手で確かめる手順（Excel が開いていて E2E が走らないとき）: 練習台 fixtures/ポスター_コンパイルエラー3件.xlsm を
    Desktop へ写して開き `compile` → `[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― …行ラベルが定義されていません。`
    """
    import gc
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Open(book)
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = "壊れ"
        comp.CodeModule.AddFromString("Sub 落ちる()\r\n    On Error GoTo Errorgo\r\n    x = 1\r\nEnd Sub\r\n")
        try:
            vis_before = bool(xl.VBE.MainWindow.Visible)
        except Exception:
            vis_before = None
        res = vbam_vba._compile_vbproject(xl, wb)
        assert res["ok"] is False and res["state"] == "error", res
        assert res.get("module") == "壊れ" and res.get("proc") == "落ちる" and res.get("line") == 2, res
        assert "[壊れ] 落ちる:2: On Error GoTo Errorgo ― " in res["detail"] and "ラベル" in res["detail"], res
        if vis_before is False:
            assert bool(xl.VBE.MainWindow.Visible) is False, "押す前に隠れていた VBE の窓は戻す"
        wb.Close(SaveChanges=False)
        comp = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    assert _wait_no_excel()
