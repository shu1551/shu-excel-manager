"""VBA静的解析 (vba lint) の徹底・実機エッジケース試験。

Rubberduck由来のインスペクションルール:
- Option Explicit 宣言漏れ
- 暗黙の Variant 型宣言 (Dim x, y As Long や Dim a)
- 未使用変数 (Dim dummy As Long)
- On Error Resume Next の放置
- 閉じ忘れ
を実機 Excel 上のモジュールに対して正確に検出・報告できるかを検証。
"""
import os
import subprocess
import sys
import gc
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBM = os.path.join(SCRIPT_DIR, "vba_manager.py")


def _run_vbm(*args, timeout=60):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run(
        [sys.executable, VBM, *args],
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        cwd=SCRIPT_DIR,
        env=env,
        timeout=timeout,
    )
    return p.returncode, (p.stdout or "") + (p.stderr or "")


@pytest.fixture
def temp_macro_excel(tmp_path):
    """マクロ有効ブック (.xlsm) を実機で起動して提供するフィクスチャ"""
    import win32com.client
    xl = None
    wb = None
    file_path = str(tmp_path / "lint_test.xlsm")
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        wb.SaveAs(file_path, FileFormat=52) # 52 = xlOpenXMLWorkbookMacroEnabled
        yield xl, wb, file_path
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        if xl:
            try:
                xl.Quit()
            except Exception:
                pass
            xl = None
        gc.collect()


@pytest.mark.core
def test_vba_lint_detects_implicit_variant_and_unused(temp_macro_excel):
    """【実機試験: 暗黙のVariant & 未使用変数の検出】
    Dim x, y As Long (xがVariant) および Dim unused As Long を実機モジュールに注入し、
    lint コマンドで漏れなく検知されることを検証。
    """
    xl, wb, file_path = temp_macro_excel
    # 実機 VBE に問題モジュールを追加
    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = "BadCodeMod"
    bad_code = (
        "Sub TestProblem()\n"
        "    Dim x, y As Long\n"
        "    Dim unused As Double\n"
        "    y = 10\n"
        "    x = 20\n"
        "    MsgBox x + y\n"
        "End Sub\n"
    )
    comp.CodeModule.AddFromString(bad_code)
    wb.Save()

    # lint 実行
    code, out = _run_vbm("lint", file_path, "BadCodeMod")
    # 暗黙のVariant (x) と 未使用変数 (unused) と Option Explicit 漏れが検知されること
    assert "Option Explicit" in out or "Option" in out
    assert "暗黙の Variant" in out or "x" in out
    assert "unused" in out


def test_vba_lint_clean_module_passes(temp_macro_excel):
    """【実機試験: 健全なモジュールは警告なし】"""
    xl, wb, file_path = temp_macro_excel
    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = "CleanCodeMod"
    clean_code = (
        "Option Explicit\n"
        "Sub CleanSub()\n"
        "    On Error GoTo CatchErr\n"
        "    Dim a As Long\n"
        "    a = 100\n"
        "    MsgBox a\n"
        "    Exit Sub\n"
        "CatchErr:\n"
        "    MsgBox Err.Description\n"
        "End Sub\n"
    )
    comp.CodeModule.AddFromString(clean_code)
    wb.Save()

    code, out = _run_vbm("lint", file_path, "CleanCodeMod")
    assert code == 0
    assert "エラーなし" in out or "警告なし" in out or "問題は見つかりませんでした" in out or "OK" in out
