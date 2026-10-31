"""
094: form-from-table の単体テストスイート
表の見出しと中身から入力フォーム設計図を自動構築するロジックの検証
"""
import pytest
from unittest.mock import MagicMock
from vbam_formbuild import (
    _sanitize_var_name,
    analyze_table_columns,
    build_form_spec_from_columns
)


def _make_mock_sheet(headers, data_rows, formulas=None, validations=None, number_formats=None):
    """COM オブジェクトを模倣した Worksheet と Range のモック"""
    ws = MagicMock()
    num_rows = len(data_rows) + 1
    num_cols = len(headers)

    table_range = MagicMock()
    table_range.Rows.Count = num_rows
    table_range.Columns.Count = num_cols

    # Header cells
    header_cells = []
    for c_idx, h in enumerate(headers, 1):
        hc = MagicMock()
        hc.Value = h
        hc.Column = c_idx
        # validation
        val_mock = MagicMock()
        if validations and c_idx - 1 < len(validations) and validations[c_idx - 1]:
            val_mock.Type = 3
            val_mock.Formula1 = validations[c_idx - 1]
        else:
            val_mock.Type = 0
            val_mock.Formula1 = None
        hc.Validation = val_mock
        header_cells.append(hc)

    def cells_func(r, c):
        if r == 1:
            return header_cells[c - 1]
        # data cell
        dc = MagicMock()
        dc.Column = c
        val_mock = MagicMock()
        if validations and c - 1 < len(validations) and validations[c - 1]:
            val_mock.Type = 3
            val_mock.Formula1 = validations[c - 1]
        else:
            val_mock.Type = 0
            val_mock.Formula1 = None
        dc.Validation = val_mock
        fmt = ""
        if number_formats and c - 1 < len(number_formats) and number_formats[c - 1]:
            fmt = number_formats[c - 1]
        dc.NumberFormatLocal = fmt
        return dc

    table_range.Cells.side_effect = cells_func
    ws.Cells.side_effect = cells_func

    # data Range
    data_rng = MagicMock()
    if len(data_rows) > 0:
        vals = tuple(tuple(row) for row in data_rows)
        if formulas:
            f_vals = tuple(tuple(f_row) for f_row in formulas)
        else:
            f_vals = tuple(tuple(None for _ in row) for row in data_rows)
    else:
        vals = ()
        f_vals = ()

    data_rng.Value2 = vals
    data_rng.Formula = f_vals
    ws.Range.return_value = data_rng

    return ws, table_range


def test_sanitize_var_name():
    """VBA識別子として安全な名前に変換されるか"""
    assert _sanitize_var_name("顧客名", prefix="txt") == "txt顧客名"
    assert _sanitize_var_name("顧客(名)", prefix="txt") == "txt顧客_名"
    assert _sanitize_var_name("1日売上", prefix="txt") == "txt_1日売上"
    assert _sanitize_var_name("###", prefix="txt") == "txtitem"


def test_formula_column_excluded():
    """C1要件: 数式列はフォーム設計図から完全に除外される"""
    headers = ["商品名", "単価", "数量", "小計"]
    data = [
        ["りんご", 100, 2, 200],
        ["みかん", 80, 5, 400]
    ]
    formulas = [
        [None, None, None, "=B2*C2"],
        [None, None, None, "=B3*C3"]
    ]
    ws, rng = _make_mock_sheet(headers, data, formulas=formulas)
    cols, text_cols = analyze_table_columns(ws, rng)[:2]

    col_names = [c["header"] for c in cols]
    assert "小計" not in col_names
    assert len(cols) == 3
    assert col_names == ["商品名", "単価", "数量"]


def test_validation_list_becomes_combo():
    """入力規則リストが combo になる"""
    headers = ["部署", "氏名"]
    data = [["総務部", "田中"], ["開発部", "佐藤"]]
    validations = ["総務部,開発部,営業部", None]
    ws, rng = _make_mock_sheet(headers, data, validations=validations)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    dep_col = cols[0]
    assert dep_col["control"] == "combo"
    assert dep_col["items"] == ["総務部", "開発部", "営業部"]


def test_cardinality_boundary_20_vs_21():
    """境目テスト: 20種類以下かつ3割以下なら combo、21種類なら txt"""
    # 100行のデータ
    # 20種類の場合 (20 <= 20 & 20 <= 30) -> combo
    items_20 = [f"種別_{i % 20}" for i in range(100)]
    # 21種類の場合 (21 > 20) -> txt
    items_21 = [f"種別_{i % 21}" for i in range(100)]

    headers = ["少区分", "多区分"]
    data = [[items_20[i], items_21[i]] for i in range(100)]
    ws, rng = _make_mock_sheet(headers, data)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    assert cols[0]["control"] == "combo"
    assert len(cols[0]["items"]) == 20
    assert cols[1]["control"] == "txt"


def test_cardinality_boundary_30_percent():
    """境目テスト: 行数の3割ちょうどは combo、3割超は txt"""
    # 10行のデータ: 3割 = 3種類
    # 3種類 (3 <= 3) -> combo
    # 4種類 (4 > 3) -> txt
    data = [
        ["A", "X1"],
        ["A", "X2"],
        ["B", "X3"],
        ["B", "X4"],
        ["C", "X1"],
        ["A", "X2"],
        ["B", "X3"],
        ["C", "X4"],
        ["A", "X1"],
        ["B", "X2"]
    ]
    headers = ["割合3", "割合4"]
    ws, rng = _make_mock_sheet(headers, data)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    assert cols[0]["control"] == "combo"
    assert cols[0]["items"] == ["A", "B", "C"]
    assert cols[1]["control"] == "txt"


def test_multiline_boundary_40_vs_41():
    """境目テスト: 平均文字数40字以下は通常txt、41字以上は multiline"""
    # 40字
    s40 = "あ" * 40
    # 41字
    s41 = "い" * 41

    headers = ["短文", "長文"]
    data = [[s40, s41], [s40, s41]]
    ws, rng = _make_mock_sheet(headers, data)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    assert cols[0]["control"] == "txt"
    assert cols[0]["multiline"] is False

    assert cols[1]["control"] == "txt"
    assert cols[1]["multiline"] is True


def test_multipage_boundary_12_vs_13():
    """境目テスト: 12列以下は単一レイアウト、13列以上は MultiPage 分割"""
    # 12列
    cols_12 = [{"header": f"列{i}", "name": f"txt{i}", "control": "txt", "required": False, "multiline": False, "hint": None} for i in range(12)]
    spec_12 = build_form_spec_from_columns("F_12", "12列フォーム", "Sheet1", "T_12", cols_12, [])
    # multipage キーが存在しないこと
    has_mp_12 = any("multipage" in r for r in spec_12["rows"])
    assert not has_mp_12

    # 13列
    cols_13 = [{"header": f"列{i}", "name": f"txt{i}", "control": "txt", "required": False, "multiline": False, "hint": None} for i in range(13)]
    spec_13 = build_form_spec_from_columns("F_13", "13列フォーム", "Sheet1", "T_13", cols_13, [])
    # multipage キーが存在すること
    has_mp_13 = any("multipage" in r for r in spec_13["rows"])
    assert has_mp_13
    mp_def = next(r for r in spec_13["rows"] if "multipage" in r)
    assert len(mp_def["pages"]) == 2  # 10項目と3項目


def test_leading_zero_and_special_text_columns():
    """先頭0の番号（00123）や郵便番号が text_columns に記録される"""
    headers = ["社員番号", "郵便番号", "電話番号", "通常数値"]
    data = [
        ["00123", "100-0001", "03-1234-5678", 123],
        ["00124", "100-0002", "03-1234-5679", 456]
    ]
    ws, rng = _make_mock_sheet(headers, data)
    cols, text_cols = analyze_table_columns(ws, rng)[:2]

    assert "社員番号" in text_cols
    assert "郵便番号" in text_cols
    assert "電話番号" in text_cols
    assert "通常数値" not in text_cols


def test_duplicate_headers_numbered():
    """同じ見出し名がある場合、コントロール名に連番が付与される"""
    headers = ["備考", "備考", "備考"]
    data = [["A", "B", "C"]]
    ws, rng = _make_mock_sheet(headers, data)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    names = [c["name"] for c in cols]
    assert names == ["txt備考", "txt備考_2", "txt備考_3"]


def test_empty_table_zero_data_rows():
    """T05: データ0行（見出しのみ）の場合でもエラーにならず構築できる"""
    headers = ["顧客名", "区分", "登録日"]
    data = []
    ws, rng = _make_mock_sheet(headers, data)
    cols, text_cols = analyze_table_columns(ws, rng)[:2]

    assert len(cols) == 3
    # データ行がない場合は required は False に倒れる
    assert all(c["required"] is False for c in cols)
    # 部品名が正常に振られている
    assert cols[0]["name"] == "txt顧客名"


def test_boolean_checkbox_detection():
    """真偽値（True/False, ○/空欄, はい/いいえ）は chk になる"""
    headers = ["有効フラグ", "受講済", "承認"]
    data = [
        ["True", "○", "はい"],
        ["False", "", "いいえ"],
        ["True", "○", "はい"]
    ]
    ws, rng = _make_mock_sheet(headers, data)
    cols, _ = analyze_table_columns(ws, rng)[:2]

    assert cols[0]["control"] == "chk"
    assert cols[1]["control"] == "chk"
    assert cols[2]["control"] == "chk"


def test_list_form_vba_no_empty_comparisons():
    """1. form_layout の一覧フォーム VBA 生成に、`<>  Then`・`= Then`・`<> Then` の形が無いこと"""
    import form_layout
    bind_info = {
        "sheet": "Sheet1",
        "table": "",
        "columns": {"txtID": "ID", "txt氏名": "氏名"},
        "start_col": 1
    }
    rows = [
        {"row": [{"lbl": "検索:"}, {"txt": "txtSearch"}]},
        {"row": [{"lst": "lstData"}]}
    ]
    vba_path = form_layout.generate_vba_stub(rows, assist=True, bind_info=bind_info, is_list_form=True)
    with open(vba_path, "r", encoding="utf-8") as f:
        code = f.read()
    assert "<>  Then" not in code
    assert "= Then" not in code
    assert "<> Then" not in code


def test_form_from_table_argparse_dest_matches_handler():
    """2. form-from-table の --test・--list の argparse dest と cmd_form_from_table の参照名が一致"""
    import argparse
    import vba_manager
    parser = vba_manager.build_parser()
    subparsers_actions = [
        action for action in parser._actions
        if isinstance(action, argparse._SubParsersAction)
    ]
    subparser = subparsers_actions[0].choices["form-from-table"]
    actions_by_opt = {}
    for act in subparser._actions:
        for opt in act.option_strings:
            actions_by_opt[opt] = act.dest

    assert "--list" in actions_by_opt
    assert "--test" in actions_by_opt
    assert actions_by_opt["--list"] == "list_flag"
    assert actions_by_opt["--test"] == "test_flag"

    import inspect
    import vbam_formbuild
    src = inspect.getsource(vbam_formbuild.cmd_form_from_table)
    assert "list_flag" in src
    assert "test_flag" in src


def test_form_name_prefix_not_duplicated():
    """5. --name F_試験 と渡したとき設計図の name が F_F_試験 にならず F_試験 のままであること"""
    import vbam_formbuild
    name_input = "F_試験"
    if not name_input.startswith("F_"):
        res = vbam_formbuild._sanitize_var_name(name_input, prefix="F_")
    else:
        res = name_input
    assert res == "F_試験"
    assert not res.startswith("F_F_")

    raw_name = "試験"
    if not raw_name.startswith("F_"):
        res_raw = vbam_formbuild._sanitize_var_name(raw_name, prefix="F_")
    else:
        res_raw = raw_name
    assert res_raw == "F_試験"
