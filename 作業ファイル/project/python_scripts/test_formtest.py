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
    # 消してから書き戻す順（逆だと書き戻した値まで消える）
    names = [c[0] for c in ws.mock_calls]
    assert names.index("UsedRange.ClearContents") < names.index("Range")


def test_run_form_test_refuses_main_tool_workbook(monkeypatch):
    """道具の本体ブック（秀コンボ）では台本を動かさない（登録ボタンを実際に押すため）"""
    import vbam_core
    from vbam_formtest import run_form_test
    wb = MagicMock()
    wb.Name = "秀コンボ.xlsm"
    monkeypatch.setattr(vbam_core, "get_workbook", lambda f=None: (MagicMock(), wb))
    ok, msg, _ = run_form_test({"steps": [{"click": "btnOK", "form": "F_x"}]})
    assert ok is False
    assert "秀コンボ" in msg
    wb.ActiveSheet.UsedRange.ClearContents.assert_not_called()


def test_run_form_test_respects_script_sheet_name(monkeypatch):
    """101: run_form_test が台本の sheet 指定を尊重し ActiveSheet で上書きしないこと"""
    from unittest.mock import MagicMock
    import vbam_formtest

    mock_xl = MagicMock()
    mock_wb = MagicMock()
    mock_wb.Name = "BookTest"
    mock_ws_active = MagicMock()
    mock_ws_active.Name = "ActiveSheet"
    mock_ws_target = MagicMock()
    mock_ws_target.Name = "TargetSheet"

    mock_wb.ActiveSheet = mock_ws_active
    mock_wb.Worksheets.side_effect = lambda name: mock_ws_target if name == "TargetSheet" else mock_ws_active

    import vbam_core
    monkeypatch.setattr(vbam_core, "get_workbook", lambda book: (mock_xl, mock_wb))
    monkeypatch.setattr(vbam_formtest, "backup_sheet_data", lambda ws: {"address": "A1:B2", "values": [["a"]], "formulas": None})
    monkeypatch.setattr(vbam_formtest, "restore_sheet_data", lambda ws, backup: None)

    script = {
        "book": "BookTest",
        "form": "F_Test",
        "caption": "Test",
        "sheet": "TargetSheet",
        "steps": []
    }
    ok, msg, _ = vbam_formtest.run_form_test(script)
    assert ok is True
    # Worksheets("TargetSheet") が呼ばれたこと
    mock_wb.Worksheets.assert_called_with("TargetSheet")


def test_get_workbook_finds_unregistered_workbook_via_active_excel():
    """ROTに載っていない未保存ブック(Book1等)が名前指定された時、_get_active_excel経由でヒットすることを検証"""
    import vbam_core
    from unittest.mock import MagicMock, patch

    mock_wb = MagicMock()
    mock_wb.Name = "Book3"
    mock_wb.Application = MagicMock()

    mock_xl = MagicMock()
    mock_xl.Workbooks = [mock_wb]

    with patch("vbam_core.smart_path_resolve", return_value=None), \
         patch("vbam_core._running_excel_workbooks", return_value=[]), \
         patch("vbam_core._get_active_excel", return_value=mock_xl), \
         patch("vbam_core._remember_excel_pid"):
        xl, wb = vbam_core._get_workbook_uncached("Book3")
        assert wb.Name == "Book3"


def test_execute_step_set_sleeps_after_completion():
    """form-set完了後にChangeイベント待ちのためウェイトが入ることを検証"""
    import vbam_formtest
    from unittest.mock import MagicMock, patch

    mock_wb = MagicMock()
    step = {"set": {"txtName": "test"}, "form": "MyForm"}

    with patch("vbam_edit.cmd_form_set", return_value=True), \
         patch("time.sleep") as mock_sleep:
        ok, msg = vbam_formtest.execute_step(step, 1, mock_wb)
        assert ok is True
        mock_sleep.assert_called_with(0.15)


def test_cmd_form_click_sends_double_stroke_for_focus_transition():
    """twice_flag（form-test の台本）のときだけ LBUTTONDOWN が 2 回送られる"""
    import vbam_edit
    import json
    from unittest.mock import MagicMock, patch, mock_open

    args = MagicMock()
    args.posargs = ["F_T8", "btnOK"]
    args.learn_flag = False
    args.list_flag = False
    args.here_flag = True
    args.twice_flag = True

    geo_data = {"btnOK": [100.0, 100.0, 50.0, 20.0, "OK"]}

    with patch("vbam_edit._list_form_windows", return_value=[(12345, "データ登録")]), \
         patch("vbam_edit._form_geometry_path", return_value="_dummy_geo.json"), \
         patch("os.path.isfile", return_value=True), \
         patch("builtins.open", mock_open(read_data=json.dumps(geo_data))), \
         patch("win32gui.EnumChildWindows", side_effect=lambda hwnd, cb, acc: acc.append(9999)), \
         patch("win32gui.ChildWindowFromPointEx", return_value=9999), \
         patch("win32gui.SendMessage"), \
         patch("win32gui.PostMessage") as mock_post, \
         patch("win32gui.GetClientRect", return_value=(0, 0, 500, 400)), \
         patch("win32gui.GetClassName", return_value="F3 Server"), \
         patch("time.sleep"):
        res = vbam_edit.cmd_form_click(args)
        assert res is True
        # WM_LBUTTONDOWN (0x0201) が2回送られていることを確認
        down_calls = [c for c in mock_post.call_args_list if c[0][1] == 0x0201]
        assert len(down_calls) == 2




def test_cmd_form_click_default_is_single_stroke():
    """既定では 1 回だけ（任意のフォームの切り替え・削除ボタンを 2 回効かせない）"""
    import vbam_edit
    import json
    from unittest.mock import MagicMock, patch, mock_open

    args = MagicMock()
    args.posargs = ["F_T8", "btnOK"]
    args.learn_flag = False
    args.list_flag = False
    args.here_flag = True

    geo_data = {"btnOK": [100.0, 100.0, 50.0, 20.0, "OK"]}

    with patch("vbam_edit._list_form_windows", return_value=[(12345, "データ登録")]), \
         patch("vbam_edit._form_geometry_path", return_value="_dummy_geo.json"), \
         patch("os.path.isfile", return_value=True), \
         patch("builtins.open", mock_open(read_data=json.dumps(geo_data))), \
         patch("win32gui.EnumChildWindows", side_effect=lambda hwnd, cb, acc: acc.append(9999)), \
         patch("win32gui.ChildWindowFromPointEx", return_value=9999), \
         patch("win32gui.SendMessage"), \
         patch("win32gui.PostMessage") as mock_post, \
         patch("win32gui.GetClientRect", return_value=(0, 0, 500, 400)), \
         patch("win32gui.GetClassName", return_value="F3 Server"), \
         patch("time.sleep"):
        res = vbam_edit.cmd_form_click(args)
        assert res is True
        # WM_LBUTTONDOWN (0x0201) が2回送られていることを確認
        down_calls = [c for c in mock_post.call_args_list if c[0][1] == 0x0201]
        assert len(down_calls) == 1

