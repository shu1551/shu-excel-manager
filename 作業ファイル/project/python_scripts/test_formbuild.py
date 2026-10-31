# -*- coding: utf-8 -*-
"""
test_formbuild.py — form-build コマンドと設計図パース機能のテスト
"""
import os
import sys
import json
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from vbam_formbuild import _validate_and_parse_spec, cmd_form_build, SAMPLE_SPEC
import form_layout


@pytest.mark.core
def test_spec_to_layout_all_parts():
    """各種部品（heading, lbl, txt, combo, lst, chk, opt_group, btn, ok, cancel, frame, multipage, spacer）が正しくパースされる"""
    spec = {
        "name": "F_AllParts",
        "caption": "全部品テスト",
        "rows": [
            {"heading": "基本情報"},
            {"row": [{"lbl": "名前"}, {"txt": "txtName", "required": True}]},
            {"row": [{"lbl": "区分"}, {"combo": "cmbKind", "items": ["A", "B"]}]},
            {"row": [{"lbl": "一覧"}, {"lst": "lstItems", "items": ["1", "2"]}]},
            {"row": [{"chk": "chkValid", "caption": "有効"}]},
            {"row": [{"opt_group": [["opt1", "高"], ["opt2", "低"]], "group": "grpP"}]},
            {"frame": "枠グループ", "rows": [
                {"row": [{"lbl": "備考"}, {"txt": "txtMemo"}]}
            ]},
            {"multipage": "mpMain", "pages": [
                {"page": "タブ1", "rows": [{"row": [{"lbl": "T1"}, {"txt": "txtT1"}]}]},
                {"page": "タブ2", "rows": [{"row": [{"lbl": "T2"}, {"txt": "txtT2"}]}]}
            ]},
            {"spacer": 16},
            {"button_bar": [{"ok": "btnOK", "caption": "登録"}, {"cancel": "btnCancel", "caption": "閉じる"}]}
        ]
    }
    form_name, caption, width, rows, *rest = _validate_and_parse_spec(spec)
    assert form_name == "F_AllParts"
    assert caption == "全部品テスト"
    assert len(rows) == 10
    
    placements, cw, ch = form_layout.compute_layout(rows)
    assert len(placements) > 0
    assert cw > 0
    assert ch > 0


@pytest.mark.core
def test_unknown_part_rejected():
    """知らない部品名が含まれる場合、行番号と部品名を明示して停止する"""
    spec = {
        "name": "F_Unknown",
        "rows": [
            {"row": [{"lbl": "名前"}, {"txt": "txtName"}]},
            {"slider": "sldVolume"}
        ]
    }
    with pytest.raises(ValueError) as excinfo:
        _validate_and_parse_spec(spec)
    err_msg = str(excinfo.value)
    assert "2 行目" in err_msg
    assert "slider が分からない" in err_msg
    assert "使える部品" in err_msg


@pytest.mark.core
def test_unknown_elem_in_row_rejected():
    """row の中に知らない要素が含まれる場合も停止する"""
    spec = {
        "name": "F_UnknownElem",
        "rows": [
            {"row": [{"lbl": "名前"}, {"gauge": "gauProgress"}]}
        ]
    }
    with pytest.raises(ValueError) as excinfo:
        _validate_and_parse_spec(spec)
    err_msg = str(excinfo.value)
    assert "gauge が分からない" in err_msg


@pytest.mark.core
def test_broken_json_handled(tmp_path):
    """構文エラーの壊れた JSON で行番号と字番号を言う"""
    broken_file = tmp_path / "broken.json"
    broken_file.write_text('{\n  "name": "F_Broken",\n  "rows": [\n', encoding="utf-8")
    
    class Args:
        sample = False
        dry_run = True
        posargs = [str(broken_file)]
        spec_file = None
        excel_file = None

    import io
    from contextlib import redirect_stdout
    out = io.StringIO()
    with redirect_stdout(out):
        cmd_form_build(Args())
    result = out.getvalue()
    assert "構文エラー" in result
    assert "行目" in result


@pytest.mark.core
def test_sample_spec_valid():
    """SAMPLE_SPEC が正常にパース可能であること"""
    form_name, caption, width, rows, *rest = _validate_and_parse_spec(SAMPLE_SPEC)
    assert form_name == "F_顧客入力"
    assert len(rows) == 6
    placements, cw, ch = form_layout.compute_layout(rows)
    assert len(placements) == 9


@pytest.mark.core
def test_fields_format_backward_compatibility():
    """fields 形式の設計図も互換レイアウトに変換できる"""
    spec = {
        "form_name": "F_Compat",
        "title": "互換テスト",
        "fields": [
            {"name": "col1", "control": "textbox", "label": "列1", "required": True},
            {"name": "col2", "control": "combobox", "label": "列2", "items": ["X", "Y"]},
            {"name": "col3", "control": "checkbox", "label": "列3"}
        ],
        "buttons": [
            {"name": "btnSubmit", "caption": "送信", "action": "submit"}
        ]
    }
    form_name, caption, width, rows, *rest = _validate_and_parse_spec(spec)
    assert form_name == "F_Compat"
    assert caption == "互換テスト"
    assert len(rows) == 4
    placements, cw, ch = form_layout.compute_layout(rows)
    assert len(placements) > 0


@pytest.mark.core
def test_missing_form_name():
    """name が指定されていない場合にエラーになること"""
    spec = {"rows": [{"row": [{"lbl": "名前"}]}]}
    with pytest.raises(ValueError) as excinfo:
        _validate_and_parse_spec(spec)
    assert "フォーム名 'name' が指定されていません" in str(excinfo.value)


@pytest.mark.core
def test_contrast_ratio_calculation():
    """白地に黒のコントラスト比は 21.0、十分な比率になる"""
    from form_inspect import contrast_ratio
    # 白 (0xFFFFFF) と 黒 (0x000000)
    cr = contrast_ratio(0xFFFFFF, 0x000000)
    assert cr >= 20.0
    # システム色: BTNFACE と BTNTEXT
    cr_sys = contrast_ratio("&H8000000F&", "&H80000012&")
    assert cr_sys >= 4.5


@pytest.mark.core
def test_lint_detects_low_contrast():
    """明るさ比 4.5 未満（コントラスト不足）を lint が拾う"""
    from form_inspect import lint_form
    info = {"width": 300, "height": 200, "inside_width": 290, "inside_height": 180}
    # 暗いグレー (0x333333) に黒文字 (0x000000) -> コントラスト比不足
    controls = [{
        "name": "lblDark", "left": 10, "top": 10, "width": 80, "height": 20,
        "type": "Label", "caption": "暗い文字", "backcolor": 0x333333, "forecolor": 0x000000
    }]
    findings = lint_form("F_Test", info, controls)
    assert any("コントラスト比不足" in f for f in findings)


@pytest.mark.core
def test_lint_detects_screen_overflow():
    """画面サイズ 1366x768 を超える巨大フォームを lint が拾う"""
    from form_inspect import lint_form
    info = {"width": 1400, "height": 800, "inside_width": 1390, "inside_height": 780}
    controls = [{"name": "lbl1", "left": 10, "top": 10, "width": 50, "height": 20, "type": "Label"}]
    findings = lint_form("F_Huge", info, controls)
    assert any("画面幅超過" in f for f in findings)
    assert any("画面高さ超過" in f for f in findings)


@pytest.mark.core
def test_lint_detects_text_truncation():
    """コントロール幅を超える長い caption（字の切れ）を lint が拾う"""
    from form_inspect import lint_form
    info = {"width": 300, "height": 200, "inside_width": 290, "inside_height": 180}
    # 幅 30pt に対して 20文字の長いラベル -> 確実に字切れ
    controls = [{
        "name": "lblTrunc", "left": 10, "top": 10, "width": 30, "height": 18,
        "type": "Label", "caption": "これは確実に幅が足りずに切れる長い見出しラベルです",
        "font_size": 11
    }]
    findings = lint_form("F_Test", info, controls)
    assert any("字の切れ" in f for f in findings)


@pytest.mark.core
def test_lint_detects_button_width_inconsistency():
    """水平帯でボタンの幅がバラバラな場合に lint が拾う"""
    from form_inspect import lint_form
    info = {"width": 400, "height": 200, "inside_width": 390, "inside_height": 180}
    controls = [
        {"name": "btn1", "left": 100, "top": 150, "width": 60, "height": 24, "type": "CommandButton"},
        {"name": "btn2", "left": 180, "top": 150, "width": 100, "height": 24, "type": "CommandButton"}
    ]
    findings = lint_form("F_Test", info, controls)
    assert any("ボタン幅不揃い" in f for f in findings)


@pytest.mark.core
def test_theme_spec_parsing():
    """設計図の theme / assist メタデータが正しく受け取れる"""
    spec_plain = {
        "name": "F_Plain",
        "theme": "plain",
        "assist": False,
        "rows": [{"row": [{"lbl": "名前"}, {"txt": "txtName"}]}]
    }
    _, _, _, _, meta = _validate_and_parse_spec(spec_plain)
    assert meta["theme"] == "plain"
    assert meta["assist"] is False

    spec_default = {
        "name": "F_Default",
        "rows": [{"row": [{"lbl": "名前"}, {"txt": "txtName"}]}]
    }
    _, _, _, _, meta_def = _validate_and_parse_spec(spec_default)
    assert meta_def["theme"] == "秀"
    assert meta_def["assist"] is True


@pytest.mark.core
def test_assist_stub_generation(tmp_path):
    """assist=True のとき Enter/Exit イベント雛形が生成され、False では生成されない"""
    from form_layout import generate_vba_stub, row, lbl, txt
    rows = [row(lbl("名前"), txt("txtName"))]
    
    out_true = tmp_path / "stub_true.vba"
    generate_vba_stub(rows, out_path=str(out_true), assist=True)
    code_true = out_true.read_text(encoding="utf-8")
    assert "txtName_Enter()" in code_true
    assert "txtName_Exit(" in code_true

    out_false = tmp_path / "stub_false.vba"
    generate_vba_stub(rows, out_path=str(out_false), assist=False)
    code_false = out_false.read_text(encoding="utf-8")
    assert "txtName_Enter()" not in code_false
    assert "txtName_Exit(" not in code_false


@pytest.mark.core
def test_forbidden_dark_contrast_rule():
    """やってはいけない形: 選択や強調で黒背景や濃い背景に暗い文字でコントラスト不足になるのを lint で検知"""
    from form_inspect import lint_form
    info = {"width": 300, "height": 200, "inside_width": 290, "inside_height": 180}
    # 濃い背景 (&H00330000&) に黒文字
    controls = [{
        "name": "btnDark", "left": 50, "top": 50, "width": 80, "height": 24,
        "type": "CommandButton", "caption": "強調", "backcolor": 0x00330000, "forecolor": 0x000000
    }]
    findings = lint_form("F_Test", info, controls)
    assert any("コントラスト比不足" in f for f in findings)


def test_generate_vba_stub_with_date_column_no_syntax_error():
    """101: generate_vba_stub に入社日などの日付列を渡したとき InStr の除算エラーが起きないこと"""
    import form_layout
    rows = [
        form_layout.row(form_layout.lbl("入社日"), form_layout.txt("txt入社日"))
    ]
    bind_info = {
        "sheet": "Sheet1",
        "columns": {"txt入社日": "入社日"}
    }
    vba_file = form_layout.generate_vba_stub(rows, assist=True, bind_info=bind_info, keys=True)
    with open(vba_file, "r", encoding="utf-8") as f:
        code = f.read()
    assert 'InStr(sDate, "/")' in code


def test_analyze_columns_handles_sparse_empty_cells():
    """空欄セルを含む列でも非空セルから日付判定が正しく行われることを検証"""
    import vbam_formbuild
    from unittest.mock import MagicMock

    mock_ws = MagicMock()
    mock_rng = MagicMock()
    mock_rng.Rows.Count = 4
    mock_rng.Columns.Count = 2

    def cell_side_effect(r, c):
        cell = MagicMock()
        cell.NumberFormatLocal = "G/標準"
        cell.Validation.Type = 0
        cell.Column = c
        cell.Value = "ID" if c == 1 else "完了日"
        return cell

    mock_rng.Cells.side_effect = cell_side_effect

    mock_data_rng = MagicMock()
    mock_data_rng.Value2 = (
        ("A001", ""),
        ("A002", ""),
        ("A003", "2026/02/15"),
    )
    mock_data_rng.Formula = (
        ("A001", ""),
        ("A002", ""),
        ("A003", "2026/02/15"),
    )
    mock_ws.Range.return_value = mock_data_rng

    cols_spec, text_cols, formula_cols = vbam_formbuild.analyze_table_columns(mock_ws, mock_rng)
    assert len(cols_spec) == 2
    date_col = cols_spec[1]
    assert date_col["is_date"] is True


def test_form_from_table_aborts_when_build_fails():
    """フォーム構築失敗時にテストを実行せず中断することを検証"""
    import vbam_formbuild
    from unittest.mock import MagicMock, patch

    mock_wb = MagicMock()
    mock_ws = MagicMock()
    mock_ws.ProtectContents = False
    mock_ws.Name = "Sheet1"
    mock_wb.ActiveSheet = mock_ws
    mock_rng = MagicMock()
    mock_rng.Rows.Count = 3
    mock_rng.Columns.Count = 2
    mock_rng.Cells.Count = 6
    mock_ws.Range.return_value = mock_rng

    args = MagicMock()
    args.posargs = ["A1:B3"]
    args.build = True
    args.test_flag = True
    args.dry_run = False
    args.list_flag = False

    with patch("vbam_formbuild.get_workbook", return_value=(None, mock_wb)), \
         patch("vbam_formbuild.analyze_table_columns", return_value=([{"header": "A", "name": "txtA", "control": "txt", "required": False, "hint": ""}], [], [])), \
         patch("vbam_formbuild.cmd_form_build", return_value=False), \
         patch("vbam_formtest.run_form_test") as mock_run_test:
        ok = vbam_formbuild.cmd_form_from_table(args)
        assert ok is False
        mock_run_test.assert_not_called()
