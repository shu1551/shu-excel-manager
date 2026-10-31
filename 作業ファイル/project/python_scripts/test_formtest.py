"""
098: form-test の単体テストスイート
自動テスト台本（パース、手ごとの期待照合、外れ報告の形式、表の退避と復元、自動台本生成）の検証
"""
import pytest
from unittest.mock import MagicMock
from vbam_formtest import (
    parse_test_script,
    backup_sheet_data,
    restore_sheet_data,
    execute_step,
    generate_auto_test_script
)


def test_parse_test_script_valid():
    """正常な台本 JSON / 辞書のパース"""
    valid_dict = {
        "book": "Book1.xlsx",
        "steps": [
            {"open": "F_顧客入力"},
            {"close": "F_顧客入力"}
        ]
    }
    parsed = parse_test_script(valid_dict)
    assert parsed["book"] == "Book1.xlsx"
    assert len(parsed["steps"]) == 2


def test_parse_test_script_invalid():
    """不正な台本形式での例外発生"""
    with pytest.raises(ValueError, match="最上位はオブジェクト"):
        parse_test_script(["not", "dict"])

    with pytest.raises(ValueError, match="steps リストが必要"):
        parse_test_script({"book": "test.xlsx"})


def test_generate_auto_test_script():
    """設計図から自動テスト台本が正しく組み立てられる"""
    cols = [{"name": "txt氏名", "header": "氏名"}]
    script = generate_auto_test_script("F_顧客入力", "顧客入力", "顧客", cols, {"txt氏名": "山田"})

    assert script["form"] == "F_顧客入力"
    assert script["sheet"] == "顧客"
    assert len(script["steps"]) >= 4
    op_names = [list(s.keys())[0] for s in script["steps"]]
    assert "open" in op_names
    assert "set" in op_names
    assert "click" in op_names
    assert "close" in op_names


def test_backup_and_restore_sheet():
    """シートの UsedRange のバックアップと復元"""
    ws = MagicMock()
    rng = MagicMock()
    rng.Address = "$A$1:$C$10"
    rng.Value2 = (("H1", "H2", "H3"), (1, 2, 3))
    rng.Formula = (("H1", "H2", "H3"), (1, 2, 3))
    ws.UsedRange = rng
    ws.Range.return_value = rng

    # 退避
    backup = backup_sheet_data(ws)
    assert backup["address"] == "$A$1:$C$10"
    assert backup["values"] == (("H1", "H2", "H3"), (1, 2, 3))

    # 復元
    restore_sheet_data(ws, backup)
    ws.Range.assert_called_with("$A$1:$C$10")
    assert rng.Value2 == (("H1", "H2", "H3"), (1, 2, 3))


def test_execute_step_expect_cell_match():
    """expect_cell で値と型が一致したときの成功"""
    wb = MagicMock()
    ws = MagicMock()
    wb.ActiveSheet = ws
    ws.UsedRange.Columns.Count = 2

    # 見出し
    def cells_func(r, c):
        mock_cell = MagicMock()
        if r == 1:
            mock_cell.Value = "顧客名" if c == 1 else "売上"
        else:
            mock_cell.Value2 = "山田" if c == 1 else 1000
            mock_cell.NumberFormatLocal = "@" if c == 1 else "#,##0"
        return mock_cell

    ws.Cells.side_effect = cells_func
    ws.Cells(ws.Rows.Count, 1).End.return_value.Row = 2

    step = {
        "expect_cell": {
            "row": 2,
            "values": {"顧客名": "山田", "売上": 1000},
            "types": {"顧客名": "string"}
        }
    }
    success, msg = execute_step(step, 5, wb, ws_target=ws)
    assert success is True
    assert msg == "OK"


def test_execute_step_expect_cell_mismatch_reports_step_index():
    """expect_cell で値不一致のとき、手番号・期待・実際が報告される"""
    wb = MagicMock()
    ws = MagicMock()
    wb.ActiveSheet = ws
    ws.UsedRange.Columns.Count = 1

    def cells_func(r, c):
        mock_cell = MagicMock()
        mock_cell.Value = "氏名" if r == 1 else "田中"
        mock_cell.Value2 = "田中" if r != 1 else "氏名"
        return mock_cell

    ws.Cells.side_effect = cells_func
    step = {
        "expect_cell": {
            "row": 2,
            "values": {"氏名": "山田"}
        }
    }
    success, msg = execute_step(step, 3, wb, ws_target=ws)
    assert success is False
    assert "手 3 (expect_cell)" in msg
    assert "期待=山田" in msg
    assert "実際=田中" in msg


def test_execute_step_expect_rows_mismatch():
    """expect_rows で行数不一致のときのエラー報告"""
    wb = MagicMock()
    ws = MagicMock()
    wb.ActiveSheet = ws
    # データ行数 3（last_row = 4）
    ws.Cells(ws.Rows.Count, 1).End.return_value.Row = 4

    step = {"expect_rows": 5}
    success, msg = execute_step(step, 7, wb, ws_target=ws)
    assert success is False
    assert "手 7 (expect_rows)" in msg
    assert "期待=5" in msg
    assert "実際=3" in msg


def test_generate_auto_test_script_expect_rows():
    """3. generate_auto_test_script(..., base_rows=3) の expect_rows が 4 になること"""
    script = generate_auto_test_script("F_Test", "Test Form", "Sheet1", [], base_rows=3)
    step_expect = [s for s in script["steps"] if "expect_rows" in s]
    assert len(step_expect) == 1
    assert step_expect[0]["expect_rows"] == 4


def test_restore_sheet_data_clears_used_range():
    """4. restore_sheet_data で ws.UsedRange.ClearContents() が呼ばれ、増えた行が消えること"""
    ws = MagicMock()
    backup = {
        "address": "A1:E4",
        "values": [["ID", "氏名"]],
        "formulas": None
    }
    restore_sheet_data(ws, backup)
    ws.UsedRange.ClearContents.assert_called_once()
    ws.Range.assert_called_with("A1:E4")
