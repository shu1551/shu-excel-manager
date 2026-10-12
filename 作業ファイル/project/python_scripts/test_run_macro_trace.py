"""マクロ実行時エラー (run-macro) の詳細スタックトレース・実機検証試験。

VBA マクロ実行時に発生したエラーについて:
- 発生モジュール名
- プロシージャ名
- エラー行番号 (Erl がある場合) / プロシージャ開始行
- エラー番号・説明
がハーネス経由で正確に捕捉・報告されるかを本物の Excel 上で検証。
"""
import os
import subprocess
import sys
import json
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
    return p.returncode, (p.stdout or ""), (p.stderr or "")


@pytest.fixture
def temp_macro_excel(tmp_path):
    """マクロ有効ブック (.xlsm) を実機で起動して提供するフィクスチャ"""
    import win32com.client
    xl = None
    wb = None
    file_path = str(tmp_path / "trace_test.xlsm")
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


def test_run_macro_zero_divide_trace(temp_macro_excel):
    """ゼロ除算エラー時に行番号ラベル (Erl) とモジュール・プロシージャが特定できるか"""
    xl, wb, file_path = temp_macro_excel

    # モジュール注入
    comp = wb.VBProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    comp.Name = "TestErrorMod"
    code = (
        "Sub TestDivZero()\n"
        "10:  Dim a As Long, b As Long\n"
        "20:  a = 100\n"
        "30:  b = 0\n"
        "40:  Dim c As Long\n"
        "50:  c = a / b\n"
        "End Sub\n"
    )
    comp.CodeModule.AddFromString(code)
    wb.Save()

    # run-macro --json を実行
    code, out, err = _run_vbm("run-macro", file_path, "TestDivZero", "--json")

    # JSONパース（ログ行を除去してJSON部分を切り出し）
    json_start = out.find("{")
    assert json_start != -1, f"JSON出力が見つかりません: out={out}, err={err}"
    data = json.loads(out[json_start:])

    assert data["success"] is False
    assert "vba_error" in data
    ve = data["vba_error"]

    assert str(ve["number"]) == "11" # 0 で除算
    assert ve["module"] == "TestErrorMod"
    assert ve["procedure"] == "TestDivZero"
    assert ve["proc_start_line"] > 0 # プロシージャ開始行が特定されている
    assert "TestErrorMod.TestDivZero" in ve["location"]


def test_run_macro_runtime_1004_trace(temp_macro_excel):
    """行番号なしの実行時エラーでモジュール・プロシージャ・開始行が特定できるか"""
    xl, wb, file_path = temp_macro_excel

    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = "TestMod2"
    code = (
        "Sub TestNoSheet()\n"
        "    Dim ws As Worksheet\n"
        "    Set ws = Worksheets(\"存在しないシート名_99999\")\n"
        "End Sub\n"
    )
    comp.CodeModule.AddFromString(code)
    wb.Save()

    code, out, err = _run_vbm("run-macro", file_path, "TestNoSheet", "--json")

    json_start = out.find("{")
    assert json_start != -1, f"JSON出力が見つかりません: out={out}, err={err}"
    data = json.loads(out[json_start:])

    assert data["success"] is False
    assert "vba_error" in data
    ve = data["vba_error"]

    assert str(ve["number"]) == "9" # インデックスが有効範囲にありません
    assert ve["module"] == "TestMod2"
    assert ve["procedure"] == "TestNoSheet"
    assert ve["erl"] == 0 # ラベルなし
    assert ve["proc_start_line"] > 0
    assert "TestMod2.TestNoSheet" in ve["location"]


def test_run_macro_success_intact(temp_macro_excel):
    """正常終了するマクロで戻り値や成功ステータスが正しく返るか"""
    xl, wb, file_path = temp_macro_excel

    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = "TestSuccessMod"
    code = (
        "Function CalcAdd() As Long\n"
        "    CalcAdd = 20 + 22\n"
        "End Function\n"
    )
    comp.CodeModule.AddFromString(code)
    wb.Save()

    code, out, err = _run_vbm("run-macro", file_path, "CalcAdd", "--json")

    json_start = out.find("{")
    assert json_start != -1, f"JSON出力が見つかりません: out={out}, err={err}"
    data = json.loads(out[json_start:])

    assert data["success"] is True
    assert data["result"] == "42"
