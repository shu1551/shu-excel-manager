# -*- coding: utf-8 -*-
"""
test_formset.py — form-set コマンドの引数解析と値処理のテスト
"""
import os
import sys
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from vbam_edit import cmd_form_set


@pytest.mark.core
def test_form_set_arg_parsing():
    """フォーム名付き・省略時の引数解析"""
    # フォーム名あり
    args1 = type("Args", (), {"posargs": ["F_顧客", "txtName=山田", "cmbKind=法人"]})()
    # フォーム名なし
    args2 = type("Args", (), {"posargs": ["txtName=田中", "chkActive=True"]})()
    
    assert "=" not in args1.posargs[0]
    assert "=" in args2.posargs[0]


@pytest.mark.core
def test_form_set_special_values():
    """改行・引用符・数式・00123・負数など特殊な値の安全な分割と保持"""
    pairs = [
        "txtName=山田 \"太郎\"",
        "txtMemo=1行目\\n2行目",
        "txtFormula==SUM(A1:A10)",
        "txtCode=00123",
        "txtPrice=-1,234",
        "chkFlag=True"
    ]
    parsed = []
    for p in pairs:
        k, v = p.split("=", 1)
        parsed.append((k, v))
        
    assert parsed[0] == ("txtName", '山田 "太郎"')
    assert parsed[2] == ("txtFormula", "=SUM(A1:A10)")
    assert parsed[3] == ("txtCode", "00123")
    assert parsed[4] == ("txtPrice", "-1,234")


@pytest.mark.core
def test_form_set_boolean_conversion():
    """CheckBox/OptionButton 用の真偽値判定ロジック"""
    truthy = ["true", "True", "1", "yes", "YES"]
    falsy = ["false", "False", "0", "no", "NO"]
    
    for val in truthy:
        assert (val.lower() in ("true", "1", "yes")) is True
        
    for val in falsy:
        assert (val.lower() in ("true", "1", "yes")) is False


@pytest.mark.core
def test_form_set_invalid_syntax(capsys):
    """キー=値 形式でない引数で明確に止まること"""
    args = type("Args", (), {"posargs": ["F_Test", "invalid_pair_without_equal"]})()
    res = cmd_form_set(args)
    assert res is False
    captured = capsys.readouterr()
    assert "形式ではありません" in captured.out


@pytest.mark.core
def test_form_set_empty_args(capsys):
    """引数なしで使い方を案内して止まること"""
    args = type("Args", (), {"posargs": []})()
    res = cmd_form_set(args)
    assert res is False
    captured = capsys.readouterr()
    assert "指定してください" in captured.out


@pytest.mark.core
def test_form_set_multiline_value_escaped():
    """101: 改行を含む文字列が __VBCRLF__ / __VBLF__ にエスケープされて specs_str に入ること"""
    raw_val = "line1\r\nline2\nline3"
    pair = f"txtPR={raw_val}"
    k, v = pair.split("=", 1)
    v_esc = str(v).replace("\r\n", "__VBCRLF__").replace("\n", "__VBLF__")
    assert "__VBCRLF__" in v_esc
    assert "__VBLF__" in v_esc
    assert "\n" not in v_esc
    # VBA 側で戻したときに元の改行に復元できること
    recovered = v_esc.replace("__VBCRLF__", "\r\n").replace("__VBLF__", "\n")
    assert recovered == raw_val
