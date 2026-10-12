"""条件付き書式の逆解析 (list)、重複・ゾンビ診断 (audit)、クレンジング (clean) の実機検証試験。

コピペ分裂病（同じルールが細切れに多重登録される現象）や #REF! ゾンビルールを、
実機 Excel 上で正確に検出・統合・掃除できるかを物理的に検証。
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
def temp_excel_sheet(tmp_path):
    """実機 Excel を起動して一時ブック・シートを提供するフィクスチャ"""
    import win32com.client
    xl = None
    wb = None
    file_path = str(tmp_path / "cond_test.xlsx")
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        wb.SaveAs(file_path)
        ws = wb.ActiveSheet
        yield xl, wb, ws, file_path
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


def test_cond_format_list_and_clean_audit(temp_excel_sheet):
    """健全な条件付き書式の逆解析 (list) と診断 (audit) が正しく動くか"""
    xl, wb, ws, file_path = temp_excel_sheet

    # ルール1: A1:A10 に セル値 > 100
    r1 = ws.Range("A1:A10")
    fc1 = r1.FormatConditions.Add(Type=1, Operator=5, Formula1="100") # Operator=5 (greater)
    fc1.Interior.Color = 255 # 赤

    # ルール2: B1:B10 に 数式 =B1<50
    r2 = ws.Range("B1:B10")
    fc2 = r2.FormatConditions.Add(Type=2, Formula1="=B1<50") # Expression
    fc2.Font.Bold = True

    wb.Save()

    # 1. list --json
    code, out, err = _run_vbm("cond-format", file_path, "list", "--json")
    json_start = out.find("{")
    assert json_start != -1, f"JSON出力が見つかりません: out={out}"
    data = json.loads(out[json_start:])

    assert data["total"] == 2
    types = [r["type"] for r in data["rules"]]
    assert "CellValue" in types
    assert "Expression" in types

    # 2. audit --json (健全な状態)
    code, out, err = _run_vbm("cond-format", file_path, "audit", "--json")
    json_start = out.find("{")
    data = json.loads(out[json_start:])

    assert data["total_rules"] == 2
    assert data["is_clean"] is True
    assert len(data["ref_errors"]) == 0
    assert len(data["exact_duplicates"]) == 0
    assert len(data["fragmented_groups"]) == 0


def test_cond_format_detects_duplicates_and_ref_zombies(temp_excel_sheet):
    """細切れ分裂・完全重複・#REF!ゾンビの検出と、clean による一括統合・削除"""
    xl, wb, ws, file_path = temp_excel_sheet

    # 1. 細切れ分裂: 同じルールを A1:A5 と A6:A10 に別々に登録
    fc1 = ws.Range("A1:A5").FormatConditions.Add(Type=2, Formula1="=A1>0")
    fc1.Interior.Color = 65535 # 黄色

    fc2 = ws.Range("A6:A10").FormatConditions.Add(Type=2, Formula1="=A1>0")
    fc2.Interior.Color = 65535 # 黄色

    # 2. 完全重複: A1:A5 に全く同じルールをもう1つ登録
    fc3 = ws.Range("A1:A5").FormatConditions.Add(Type=2, Formula1="=A1>0")
    fc3.Interior.Color = 65535 # 黄色

    # 3. #REF! ゾンビ: 壊れた参照を含むルール
    fc_zombie = ws.Range("C1:C10").FormatConditions.Add(Type=2, Formula1="=#REF!>10")
    fc_zombie.Interior.Color = 255

    wb.Save()

    # 診断 (audit)
    code, out, err = _run_vbm("cond-format", file_path, "audit", "--json")
    json_start = out.find("{")
    data = json.loads(out[json_start:])

    assert data["is_clean"] is False
    assert data["total_rules"] == 4
    assert len(data["ref_errors"]) == 1 # fc_zombie
    assert len(data["exact_duplicates"]) == 1 # fc3
    assert len(data["fragmented_groups"]) >= 1 # fc1 と fc2 の細切れ

    # クレンジング (clean)
    code, out, err = _run_vbm("cond-format", file_path, "clean", "--json")
    json_start = out.find("{")
    assert json_start != -1, f"JSON出力が見つかりません: code={code}, out={out}, err={err}"
    cdata = json.loads(out[json_start:])

    assert cdata["success"] is True
    assert cdata["deleted_ref"] == 1 # ゾンビ削除
    assert cdata["deleted_dups"] >= 1 # 重複・細切れ削除
    assert cdata["after"] < cdata["before"]

    # クレンジング後の再診断 (audit) → 健全化されていること
    code, out, err = _run_vbm("cond-format", file_path, "audit", "--json")
    json_start = out.find("{")
    adata = json.loads(out[json_start:])

    assert adata["is_clean"] is True
    assert len(adata["ref_errors"]) == 0
    assert len(adata["exact_duplicates"]) == 0
    assert len(adata["fragmented_groups"]) == 0
