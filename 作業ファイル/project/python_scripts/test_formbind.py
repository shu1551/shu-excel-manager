"""
095: form-bind の単体テストスイート
表とフォームをつなぐ VBA 登録コード生成（型保全・必須バリデーション・テーブル/範囲対応）の検証
"""
import pytest
from form_layout import _generate_bind_click_code, generate_vba_stub, row, lbl, txt, combo, chk, button_bar, ok, cancel


def _get_stub_code(rows, bind_info=None):
    import tempfile, os
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name
    generate_vba_stub(rows, out_path=tmp_name, assist=False, bind_info=bind_info)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass
    return code


def test_bind_vba_table_basic():
    """テーブル（ListObject）への行追加コードが生成される"""
    rows = [
        row(lbl("氏名"), txt("txtName", required=True)),
        row(lbl("区分"), combo("cmbKind", items=["法人", "個人"])),
        button_bar(ok("btnOK", "登録"), cancel("btnCancel", "閉じる"))
    ]
    bind = {
        "sheet": "顧客シート",
        "table": "T_顧客",
        "columns": {"txtName": "氏名", "cmbKind": "区分"},
        "text_columns": [],
        "column_types": {"txtName": "text", "cmbKind": {"type": "combo", "items": ["法人", "個人"]}},
        "column_positions": {"txtName": 1, "cmbKind": 2},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'Set lo = ws.ListObjects("T_顧客")' in code
    assert 'Set lr = lo.ListRows.Add' in code
    assert 'lr.Range.Cells(1, 1).Value = txtName.Value' in code
    assert 'lr.Range.Cells(1, 2).Value = cmbKind.Value' in code
    assert 'Application.StatusBar = "登録しました (" & lr.Index & "行目)"' in code


def test_bind_vba_range_basic():
    """範囲（Range）の場合の末尾行特定と代入コードが生成される"""
    rows = [
        row(lbl("商品名"), txt("txtItem")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "売上",
        "table": "",
        "columns": {"txtItem": "商品名"},
        "text_columns": [],
        "column_types": {"txtItem": "text"},
        "column_positions": {"txtItem": 1},
        "formula_columns": [],
        "start_col": 2  # B列始まり
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'targetRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1' in code
    assert 'ws.Cells(targetRow, 2).Value = txtItem.Value' in code


def test_bind_vba_text_columns_protection():
    """先頭0落ち番号（00123、電話、郵便）が文字列型保護される"""
    rows = [
        row(lbl("社員番号"), txt("txtEmpNo")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "社員",
        "table": "T_社員",
        "columns": {"txtEmpNo": "社員番号"},
        "text_columns": ["社員番号"],
        "column_types": {"txtEmpNo": "text"},
        "column_positions": {"txtEmpNo": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'lr.Range.Cells(1, 1).NumberFormatLocal = "@"' in code
    assert 'lr.Range.Cells(1, 1).Value = CStr(txtEmpNo.Value)' in code


def test_bind_vba_date_conversion():
    """日付列は IsDate 検証と CDate() 変換が行われる"""
    rows = [
        row(lbl("取引日"), txt("txtDate")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"txtDate": "取引日"},
        "text_columns": [],
        "column_types": {"txtDate": "date"},
        "column_positions": {"txtDate": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If Not IsDate(txtDate.Value) Then' in code
    assert 'lr.Range.Cells(1, 1).Value = CDate(txtDate.Value)' in code


def test_bind_vba_number_conversion():
    """数値・金額列は IsNumeric 検証と CDbl(Replace) 変換が行われる"""
    rows = [
        row(lbl("金額"), txt("txtAmount")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"txtAmount": "金額"},
        "text_columns": [],
        "column_types": {"txtAmount": "number"},
        "column_positions": {"txtAmount": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If Not IsNumeric(Replace(Trim(txtAmount.Value), ",", "")) Then' in code
    assert 'lr.Range.Cells(1, 1).Value = CDbl(Replace(txtAmount.Value, ",", ""))' in code


def test_bind_vba_boolean_style():
    """真偽値（chk）は表の流儀（○/空欄）に変換されて代入される"""
    rows = [
        row(lbl("有効"), chk("chkActive", "有効")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"chkActive": "有効"},
        "text_columns": [],
        "column_types": {"chkActive": {"type": "bool", "true_val": "○", "false_val": ""}},
        "column_positions": {"chkActive": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If chkActive.Value = True Then' in code
    assert 'lr.Range.Cells(1, 1).Value = "○"' in code
    assert 'lr.Range.Cells(1, 1).Value = ""' in code


def test_bind_vba_formula_column_copy():
    """範囲の場合、数式列は上行の式がコピーされる"""
    rows = [
        row(lbl("単価"), txt("txtPrice")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "",
        "columns": {"txtPrice": "単価"},
        "text_columns": [],
        "column_types": {"txtPrice": "number"},
        "column_positions": {"txtPrice": 1},
        "formula_columns": [2],  # 2列目が小計数式
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If targetRow > 2 Then' in code
    assert 'ws.Cells(targetRow - 1, 2).Copy ws.Cells(targetRow, 2)' in code


def test_bind_vba_required_validation():
    """必須項目は空欄時に警告を表示して中断する"""
    rows = [
        row(lbl("顧客名"), txt("txtName", required=True)),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"txtName": "顧客名"},
        "text_columns": [],
        "column_types": {"txtName": "text"},
        "column_positions": {"txtName": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If Trim(txtName.Value) = "" Then' in code
    assert 'MsgBox "顧客名 を入力してください", vbExclamation' in code
    assert 'txtName.SetFocus' in code
    assert 'Exit Sub' in code


def test_bind_vba_validation_list_check():
    """入力規則リストの候補外入力は弾かれる"""
    rows = [
        row(lbl("部署"), combo("cmbDept", items=["総務", "営業"])),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"cmbDept": "部署"},
        "text_columns": [],
        "column_types": {"cmbDept": {"type": "combo", "items": ["総務", "営業"]}},
        "column_positions": {"cmbDept": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'cmbDept.Value <> "総務" And cmbDept.Value <> "営業"' in code
    assert 'MsgBox "【部署】は一覧から選択してください", vbExclamation' in code


def test_bind_vba_sheet_protection_guard():
    """シート保護されている場合は登録を拒否する"""
    rows = [
        row(lbl("項目"), txt("txtItem")),
        button_bar(ok("btnOK", "登録"))
    ]
    bind = {
        "sheet": "Sheet1",
        "table": "T1",
        "columns": {"txtItem": "項目"},
        "text_columns": [],
        "column_types": {"txtItem": "text"},
        "column_positions": {"txtItem": 1},
        "formula_columns": [],
        "start_col": 1
    }
    code = _get_stub_code(rows, bind_info=bind)

    assert 'If ws.ProtectContents Then' in code
    assert 'MsgBox "シートが保護されているため登録できません。", vbCritical' in code


def test_bind_no_bind_flag():
    """bind_info が None のときは TODO コメントのままになる"""
    rows = [
        row(lbl("項目"), txt("txtItem")),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, bind_info=None)

    assert "' TODO: 登録 の処理" in code
    assert "ListRows.Add" not in code
