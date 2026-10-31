"""
096: form-keys の単体テストスイート
キー操作（Activate時の初期フォーカス、Enter遷移・登録、Esc、MultiPage巡回、form-key）の検証
"""
import pytest
import tempfile
import os
from unittest.mock import MagicMock
from form_layout import generate_vba_stub, row, lbl, txt, combo, button_bar, ok, cancel, multipage, page
from vbam_edit import cmd_form_key


def _get_stub_code(rows, keys=True):
    with tempfile.NamedTemporaryFile("w+", suffix=".vba", delete=False, encoding="utf-8") as tf:
        tmp_name = tf.name
    generate_vba_stub(rows, out_path=tmp_name, assist=False, bind_info=None, keys=keys)
    with open(tmp_name, "r", encoding="utf-8") as f:
        code = f.read()
    try:
        os.remove(tmp_name)
    except Exception:
        pass
    return code


def test_keys_vba_activate_setfocus():
    """keys=True のとき UserForm_Activate が生成され最初の欄に SetFocus される"""
    rows = [
        row(lbl("氏名"), txt("txtName")),
        row(lbl("区分"), combo("cmbKind", items=["A", "B"])),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=True)

    assert "Private Sub UserForm_Activate()" in code
    assert "txtName.SetFocus" in code


def test_keys_vba_keydown_enter_navigation():
    """1行欄で Enter を押すと次の入力欄へ SetFocus する KeyDown が生成される"""
    rows = [
        row(lbl("氏名"), txt("txtName")),
        row(lbl("電話"), txt("txtPhone")),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=True)

    assert "Private Sub txtName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" in code
    assert "If KeyCode = vbKeyReturn And Shift = 0 Then" in code
    assert "KeyCode = 0" in code
    assert "txtPhone.SetFocus" in code


def test_keys_vba_keydown_last_control_submit():
    """最後の1行欄で Enter を押すと Call btnOK_Click（登録）が呼ばれる"""
    rows = [
        row(lbl("氏名"), txt("txtName")),
        row(lbl("電話"), txt("txtPhone")),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=True)

    assert "Private Sub txtPhone_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" in code
    assert "Call btnOK_Click" in code


def test_keys_vba_keys_false():
    """keys=False のときは Activate も KeyDown も生成されない"""
    rows = [
        row(lbl("氏名"), txt("txtName")),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=False)

    assert "Private Sub UserForm_Activate()" not in code
    assert "txtName_KeyDown" not in code


def test_keys_vba_multiline_excluded():
    """複数行 TextBox は EnterKeyBehavior が有効で KeyDown 遷移から除外される"""
    rows = [
        row(lbl("氏名"), txt("txtName")),
        row(lbl("備考"), txt("txtMemo", multiline=True)),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=True)

    # txtMemo は multiline なので KeyDown は生成されない
    assert "Private Sub txtMemo_KeyDown" not in code
    # txtName が最後の1行コントロールとなるため、txtName から btnOK_Click が呼ばれる
    assert "Private Sub txtName_KeyDown" in code
    assert "Call btnOK_Click" in code


def test_keys_vba_multipage_ctrl_tab():
    """MultiPage がある場合、Ctrl+Tab でページを巡回するコードが生成される"""
    rows = [
        multipage("mpMain",
            page("基本", row(lbl("氏名"), txt("txtName"))),
            page("詳細", row(lbl("住所"), txt("txtAddress")))
        ),
        button_bar(ok("btnOK", "登録"))
    ]
    code = _get_stub_code(rows, keys=True)

    assert "ElseIf Shift = 2 And KeyCode = vbKeyTab Then" in code
    assert "mpMain.Value = (mpMain.Value + 1) Mod mpMain.Pages.Count" in code


def test_form_key_cmd_args_validation(capsys):
    """form-key コマンドの引数不足・未対応キーの検査"""
    # 引数なし
    args = MagicMock()
    args.posargs = []
    res = cmd_form_key(args)
    assert res is False
    out, _ = capsys.readouterr()
    assert "エラー" in out

    # 未対応キー名
    args.posargs = ["F_Dummy", "invalid_key"]
    # フォームが存在しない場合
    res = cmd_form_key(args)
    assert res is False
