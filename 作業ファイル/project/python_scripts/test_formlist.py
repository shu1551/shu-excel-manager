"""
097: form-list の単体テストスイート
一覧と検索の型（設計図生成、先頭6列制限、列幅計算、VBA絞り込みコード、キー遷移）の検証
"""
import pytest
import tempfile
import os
from vbam_formbuild import build_list_form_spec
from form_layout import generate_vba_stub, _generate_list_form_vba


def test_build_list_form_spec_basic():
    """一覧フォーム設計図が正しく組み立てられる（kind: list, txtSearch, lstData, lblCount）"""
    cols_spec = [
        {"name": "txt1", "header": "氏名", "is_date": False, "is_currency": False},
        {"name": "txt2", "header": "区分", "is_date": False, "is_currency": False},
        {"name": "txt3", "header": "売上", "is_date": False, "is_currency": True}
    ]
    spec = build_list_form_spec("F_顧客一覧", "顧客", "顧客シート", "T_顧客", cols_spec)

    assert spec["name"] == "F_顧客一覧"
    assert spec["kind"] == "list"
    assert spec["bind"]["sheet"] == "顧客シート"
    assert spec["bind"]["table"] == "T_顧客"
    assert spec["bind"]["entry_form"] == "F_顧客入力"

    # コントロールの存在確認
    row_elems = [e for r in spec["rows"] if "row" in r for e in r["row"]]
    names = [e.get("txt") or e.get("lst") or e.get("name") for e in row_elems]
    assert "txtSearch" in names
    assert "lstData" in names
    assert "lblCount" in names


def test_list_form_max_columns_and_widths():
    """6列を超える表の場合、ListBox は先頭6列に制限され、列幅が計算される"""
    cols_spec = [
        {"name": f"col_{i}", "header": f"項目_{i}", "is_date": (i == 2), "is_currency": (i == 3)}
        for i in range(15)
    ]
    spec = build_list_form_spec("F_15列一覧", "15列", "Sheet1", "T_15", cols_spec)

    lst_elem = next(e for r in spec["rows"] if "row" in r for e in r["row"] if "lst" in e)
    assert lst_elem["column_count"] == 6
    widths = lst_elem["column_widths"].split(";")
    assert len(widths) == 6
    # 日付は 80（yyyy/mm/dd が最後まで見える幅）、金額は 64、他は 80
    assert widths[2] == "80"
    assert widths[3] == "64"
    assert widths[0] == "80"


def test_list_form_vba_stub_generation():
    """一覧検索用 VBA コード（LoadData, RefreshList, txtSearch_Change）が生成される"""
    bind = {
        "sheet": "顧客シート",
        "table": "T_顧客",
        "columns": {"txt1": "氏名", "txt2": "区分"},
        "start_col": 1,
        "entry_form": "F_顧客入力"
    }
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name

    _generate_list_form_vba(bind, out_path=tmp_name)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass

    assert 'Private Const M_SHEET As String = "顧客シート"' in code
    assert 'Private Const M_TABLE As String = "T_顧客"' in code
    assert 'Private Sub UserForm_Initialize()' in code
    assert 'Private Sub LoadData()' in code
    assert 'Private Sub RefreshList(ByVal keyword As String)' in code
    assert 'Private Sub txtSearch_Change()' in code
    assert 'F_顧客入力.Show' in code


def test_list_form_case_width_insensitive_search():
    """検索ロジックで StrConv(vbNarrow) と LCase による全角半角・大文字小文字無視が組まれている"""
    bind = {"sheet": "S1", "table": "T1", "columns": {}, "start_col": 1}
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name
    _generate_list_form_vba(bind, out_path=tmp_name)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass

    assert "kw = LCase(StrConv(Trim(keyword), vbNarrow))" in code
    assert "StrConv(CellStr(mAllData(r, c)), vbNarrow)" in code
    assert "InStr(1, LCase(rowText), kw)" in code


def test_list_form_zero_data_handling():
    """データ0行の場合に 0 件中 0 件 と表示して安全に抜けるガードがある"""
    bind = {"sheet": "S1", "table": "T1", "columns": {}, "start_col": 1}
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name
    _generate_list_form_vba(bind, out_path=tmp_name)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass

    assert "If mRowCount = 0 Then" in code
    assert 'lblCount.Caption = "0 件中 0 件"' in code


def test_list_form_keys_navigation():
    """検索欄で ↓ キーを押すと ListBox へ遷移し、Enter で btnEdit が呼ばれる"""
    bind = {"sheet": "S1", "table": "T1", "columns": {}, "start_col": 1}
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name
    _generate_list_form_vba(bind, out_path=tmp_name)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass

    assert "Private Sub txtSearch_KeyDown" in code
    assert "If KeyCode = vbKeyDown Then" in code
    assert "lstData.SetFocus" in code
    assert "ElseIf KeyCode = vbKeyReturn Then" in code
    assert "Call btnEdit_Click" in code
