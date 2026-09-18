"""COM 不要の純粋ロジックの自動テスト（本体: vba_manager と Excel を触る道具）。

実行: このフォルダで `py -m pytest test_tools.py -v`
Excel には一切接続しない（COM オブジェクトはダミーで代用）。
AI の機能（agent・試験・揺らし・手順書・キー）のテストは test_agent.py（2026-09-12 に分けた）。
"""
import os
import sys
import json
import inspect

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vba_manager as vm
import form_layout as fl
import form_inspect as fi
import vbam_clean as vc
import vbam_hands as vh


# ================================================================
# vba_manager: 改行・エンコーディングの多層ガード
# ================================================================

def test_normalize_bas_newlines_idempotent(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub A()\r\nEnd Sub\r\n".encode('cp932'))
    fixed, raw, was = vm.normalize_bas_newlines(str(p))
    assert not was
    assert fixed == raw


def test_normalize_bas_newlines_fixes_doubling(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub A()\r\r\nEnd Sub\r\r\n".encode('cp932'))
    fixed, raw, was = vm.normalize_bas_newlines(str(p))
    assert was
    assert fixed == "Sub A()\r\nEnd Sub\r\n".encode('cp932')


def test_read_code_file_collapses_doubling(tmp_path):
    p = tmp_path / "a.vba"
    p.write_bytes('Sub A()\r\r\n    MsgBox "x"\r\r\nEnd Sub\r\r\n'.encode('cp932'))
    assert vm.read_code_file(str(p)) == 'Sub A()\n    MsgBox "x"\nEnd Sub\n'


def test_validate_bas_encoding_rejects_utf8_japanese(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub あ()\r\nEnd Sub\r\n".encode('utf-8'))
    assert vm.validate_bas_encoding(str(p)) is False


def test_validate_bas_encoding_accepts_cp932(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub あ()\r\nEnd Sub\r\n".encode('cp932'))
    assert vm.validate_bas_encoding(str(p)) is True


def test_duplicate_procedures_detected():
    code = "Sub A()\nEnd Sub\nSub A()\nEnd Sub\n"
    assert "A" in vm._find_duplicate_procedures(code)


def test_parse_module_blocks_end_sub_comment():
    # End Sub の末尾コメントでブロック境界がずれないこと（隣Sub消失バグの回帰）
    bas = ('Attribute VB_Name = "M"\r\n'
           "Sub A()\r\n"
           "End Sub ' comment\r\n"
           "Sub B()\r\n"
           "End Sub\r\n")
    header, blocks, trailing = vm._parse_module_blocks(bas)
    assert [b['name'] for b in blocks] == ['A', 'B']


def test_looks_like_xl_file():
    assert vm.looks_like_xl_file("a.xlsm")
    assert vm.looks_like_xl_file(r"C:\path\b.xlam")
    assert not vm.looks_like_xl_file("fix.vba")          # 位置引数の罠の回帰
    assert not vm.looks_like_xl_file(r"C:\tmp\fix.vba")


def test_coerce_cell():
    assert vm._coerce_cell("12") == 12
    assert vm._coerce_cell("1.5") == 1.5
    assert vm._coerce_cell("=SUM(A1)") == "=SUM(A1)"
    assert vm._coerce_cell("nan") == "nan"     # Excel でエラー値化するため文字列のまま
    assert vm._coerce_cell("inf") == "inf"
    assert vm._coerce_cell("") is None
    # 全角数字は Excel と同じくテキストのまま（int()/float() が通してしまう経路を塞ぐ）
    assert vm._coerce_cell("１２３") == "１２３"
    assert vm._coerce_cell("１.５") == "１.５"
    assert vm._coerce_cell("２０２６/１/１") == "２０２６/１/１"


# ================================================================
# form_layout: レイアウト計算の不変条件
# ================================================================

def _std_rows():
    return [
        fl.row(fl.lbl("名前"), fl.txt("txtName")),
        fl.row(fl.lbl("区分"), fl.combo("cmbKind")),
        fl.spacer(),
        fl.button_bar(fl.ok("btnOK"), fl.cancel("btnCancel")),
    ]


def test_layout_alignment():
    pl, cw, ch = fl.compute_layout(_std_rows())
    lefts = {l for e, l, t, w, h in pl if e['kind'] in ('txt', 'combo')}
    rights = {l + w for e, l, t, w, h in pl if e['kind'] in ('txt', 'combo')}
    assert len(lefts) == 1, "入力の左端は1本に揃う"
    assert len(rights) == 1, "入力の右端は1本に揃う"


def test_layout_button_bar_uniform_and_right_aligned():
    pl, cw, ch = fl.compute_layout(_std_rows())
    sizes = {(w, h) for e, l, t, w, h in pl if e['kind'] == 'btn'}
    assert len(sizes) == 1, "ボタンバーは同サイズ"
    right = max(l + w for e, l, t, w, h in pl if e['kind'] == 'btn')
    assert right == fl.STYLE['pad'] + cw, "ボタンバーはコンテンツ右端に揃う"


def test_frame_children_fit():
    # frame 内の最終行がクリップされない（frame_top 足し忘れバグの回帰）
    rows = [fl.frame("G",
                     fl.row(fl.lbl("A"), fl.txt("t1")),
                     fl.row(fl.chk("c1", "チェック")))]
    pl, cw, ch = fl.compute_layout(rows)
    fe, l, t, w, h = pl[0]
    for ce, cl, ct, cw2, ch2 in fe['children']:
        assert ct + ch2 <= h, f"{ce.get('name')} が frame からはみ出している"


def test_multipage_width_included():
    # multipage がコンテンツ幅の自動計算に入っている（計算漏れバグの回帰）
    rows = [fl.multipage("mp",
                         fl.page("P1", fl.row(fl.lbl("ラベル"), fl.txt("t", width=200))))]
    pl, cw, ch = fl.compute_layout(rows)
    assert cw >= 200


def test_required_star_on_label():
    rows = [fl.row(fl.lbl("名前"), fl.txt("txtName", required=True))]
    pl, cw, ch = fl.compute_layout(rows)
    labels = [e for e, *_ in pl if e['kind'] == 'lbl']
    assert any('＊' in (e.get('caption') or '') for e in labels)


def test_stub_required_spin_cancel(tmp_path):
    rows = [
        fl.row(fl.lbl("名前"), fl.txt("txtName", required=True)),
        fl.row(fl.lbl("数"), fl.spin_txt("txtQty")),
        fl.button_bar(fl.ok("btnGo", "実行"), fl.cancel("btnClose")),
    ]
    out = fl.generate_vba_stub(rows, str(tmp_path / "s.vba"))
    code = open(out, encoding='utf-8').read()
    assert 'If Trim(txtName.Value) = ""' in code, "必須チェックの雛形"
    assert "txtQtySpin_Change" in code, "スピン連動イベント"
    assert "Unload Me" in code, "キャンセルの雛形"


def test_refedit_is_composite_with_pick_stub(tmp_path):
    rows = [fl.row(fl.lbl("範囲"), fl.refedit("refX")),   # refedit は複合部品（rowが展開）
            fl.button_bar(fl.ok("btnGo"))]
    out = fl.generate_vba_stub(rows, str(tmp_path / "s.vba"))
    code = open(out, encoding='utf-8').read()
    assert "Application.InputBox" in code, "範囲選択ハンドラの雛形"


# ================================================================
# form_inspect: lint とリバースの機械判定
# ================================================================

def _ctl(name, type_, l, t, w, h, parent="F", **kw):
    d = dict(name=name, type=type_, left=l, top=t, width=w, height=h,
             caption=kw.pop('caption', None), font_size=kw.pop('font_size', 12.0),
             parent=parent, tab_index=kw.pop('tab_index', None),
             bold=kw.pop('bold', None))
    d.update(kw)
    return d


def test_lint_detects_overlap_and_out_of_bounds():
    info = {"inside_width": 200, "inside_height": 100}
    ctrls = [
        _ctl("a", "TextBox", 10, 10, 100, 22),
        _ctl("b", "TextBox", 50, 12, 100, 22),      # a と重なる
        _ctl("c", "TextBox", 150, 90, 100, 22),     # 右下にはみ出す
    ]
    findings = fi.lint_form("F", info, ctrls)
    assert any("重なり" in s for s in findings)
    assert any("はみ出し" in s for s in findings)


def test_lint_clean_form_has_no_findings():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [
        _ctl("lbl1", "Label", 12, 14, 40, 18),
        _ctl("txtA", "TextBox", 60, 12, 200, 22),
        _ctl("lbl2", "Label", 12, 44, 40, 18),
        _ctl("txtB", "TextBox", 60, 42, 200, 22),
    ]
    assert fi.lint_form("F", info, ctrls) == []


def test_lint_orphan_handler():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [_ctl("btnGo", "CommandButton", 10, 10, 72, 24,
                  default=True, cancel=True, accelerator=None)]
    code = ("Private Sub btnGone_Click()\nEnd Sub\n"
            "Private Sub btnGo_Click()\nEnd Sub\n")
    findings = fi.lint_form("F", info, ctrls, code=code)
    assert any("孤児ハンドラ" in s and "btnGone" in s for s in findings)
    assert not any("btnGo'" in s for s in findings)


def test_lint_missing_click_handler():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [_ctl("btnGo", "CommandButton", 10, 10, 72, 24,
                  default=True, cancel=True, accelerator=None)]
    findings = fi.lint_form("F", info, ctrls, code="Private Sub UserForm_Initialize()\nEnd Sub\n")
    assert any("Click ハンドラ未実装" in s for s in findings)


def test_type_normalize():
    assert fi._normalize_type("IMdcText") == "TextBox"
    assert fi._normalize_type("ILabelControl") == "Label"
    assert fi._normalize_type("IMultiPage") == "MultiPage"
    assert fi._normalize_type("ICommandButton") == "CommandButton"
    assert fi._normalize_type("TextBox") == "TextBox"


def test_cluster_rows_groups_by_center():
    items = [
        _ctl("lbl1", "Label", 12, 14, 40, 18),      # 中央 23
        _ctl("txtA", "TextBox", 60, 12, 200, 22),   # 中央 23 → 同じ行
        _ctl("txtB", "TextBox", 60, 42, 200, 22),   # 別の行
    ]
    rows = fi._cluster_rows(items)
    assert len(rows) == 2
    assert [c["name"] for c in rows[0]] == ["lbl1", "txtA"]


# ================================================================
# vba_manager: 健康診断（checkup）強化まわりの純粋ロジック
# ================================================================

def _mod(name, type_, code, procs=()):
    return {'name': name, 'type': type_, 'type_name': '',
            'total_lines': code.count('\r\n'), 'procs': list(procs), 'code': code}


def test_strip_vba_comment_keeps_quote_in_string():
    assert vm._strip_vba_comment('x = "a\'b" \' comment') == 'x = "a\'b" '
    assert vm._strip_vba_comment("' 全部コメント") == ""


def test_extra_scans_hardcoded_path_and_comment_excluded():
    code = ('Sub A()\r\n'
            '    p = "C:\\data\\in.csv"\r\n'
            '    \' 例: "C:\\old\\path"\r\n'
            '    On Error Resume Next\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 5}])]}
    res = vm._extra_code_scans(inv)
    assert [(m, p, path) for m, p, _, path in res['hardcoded_paths']] == \
        [('M1', 'A', 'C:\\data\\in.csv')]
    assert [(m, p) for m, p, _ in res['error_resume']] == [('M1', 'A')]
    assert res['no_option_explicit'] == ['M1']


def test_extra_scans_auto_exec_and_option_explicit():
    code = 'Option Explicit\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n'
    inv = {'modules': [_mod('ThisWorkbook', 100, code,
                            [{'name': 'Workbook_Open', 'lines': 3}]),
                       _mod('Module1', 1, 'Sub Auto_Open()\r\nEnd Sub\r\n',
                            [{'name': 'Auto_Open', 'lines': 2}])]}
    res = vm._extra_code_scans(inv)
    assert {n for _, n, _ in res['auto_exec']} == {'Workbook_Open', 'Auto_Open'}
    assert res['no_option_explicit'] == ['Module1']


def test_checkup_diff_line_shift_is_not_a_change():
    prev = {'time': '2026-07-03 23:00', 'keys': ['未解決Call: [M1] 親 → 子'],
            'procs': {'M1': ['親']}, 'total_lines': 100, 'sheets': ['S1'], 'forms': []}
    cur = {'time': '2026-07-04 09:00', 'keys': ['未解決Call: [M1] 親 → 子'],
           'procs': {'M1': ['親']}, 'total_lines': 100, 'sheets': ['S1'], 'forms': []}
    assert vm._checkup_diff(prev, cur)['changed'] is False


def test_checkup_diff_detects_new_resolved_and_growth():
    prev = {'time': 't0', 'keys': ['重複プロシージャ: [M1] A'],
            'procs': {'M1': ['A']}, 'total_lines': 50, 'sheets': ['S1'], 'forms': []}
    cur = {'time': 't1', 'keys': ['未解決Call: [M2] B → 消えた子'],
           'procs': {'M1': ['A'], 'M2': ['B']}, 'total_lines': 80,
           'sheets': ['S1', 'S2'], 'forms': ['F_New']}
    d = vm._checkup_diff(prev, cur)
    assert d['changed']
    assert d['new'] == ['未解決Call: [M2] B → 消えた子']
    assert d['resolved'] == ['重複プロシージャ: [M1] A']
    assert d['procs_added'] == ['[M2] B']
    assert d['lines_delta'] == 30
    assert d['sheets_added'] == ['S2'] and d['forms_added'] == ['F_New']


def test_checkup_diff_first_run_returns_none():
    assert vm._checkup_diff(None, {'keys': [], 'procs': {}, 'total_lines': 0,
                                   'sheets': [], 'forms': []}) is None


def test_extra_scans_destructive_and_no_restore():
    code = ('Sub 掃除()\r\n'
            '    Application.ScreenUpdating = False\r\n'
            '    Kill "C:\\tmp\\old.txt"\r\n'
            '    Worksheets("作業").Delete\r\n'
            '    ActiveSheet.Rows("2:" & lastRow).Delete\r\n'
            'End Sub\r\n'
            'Sub 正常()\r\n'
            '    Application.ScreenUpdating = False\r\n'
            '    Application.ScreenUpdating = True\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code,
                            [{'name': '掃除', 'lines': 6}, {'name': '正常', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    labels = {(p, lab) for _, p, _, lab, _ in res['destructive']}
    assert ('掃除', 'ファイル/フォルダ削除') in labels
    assert ('掃除', 'シート削除') in labels
    assert ('掃除', '行/列の削除') in labels
    # 行削除の行が「シート削除」と誤ラベルされないこと
    row_line_labels = {lab for _, _, ln, lab, _ in res['destructive'] if ln == 5}
    assert row_line_labels == {'行/列の削除'}
    assert [(m, p) for m, p, _ in res['no_restore']] == [('M1', '掃除')]


def test_extra_scans_no_restore_ignores_comment_and_string():
    code = ('Sub A()\r\n'
            '    \' Application.ScreenUpdating = False と書いたコメント\r\n'
            '    s = "Kill されそうな文字列"\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    assert res['no_restore'] == [] and res['destructive'] == []


def test_extra_scans_enableevents_no_restore():
    code = ('Sub 事故りがち()\r\n'
            '    Application.EnableEvents = False\r\n'
            'End Sub\r\n'
            'Sub 正しい()\r\n'
            '    Application.EnableEvents = False\r\n'
            '    Application.EnableEvents = True\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': '事故りがち', 'lines': 3},
                                            {'name': '正しい', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    assert [(m, p) for m, p, _ in res['no_restore']] == [('M1', '事故りがち')]
    assert 'EnableEvents' in res['no_restore'][0][2]


def test_checkup_rating():
    assert vm._checkup_rating(0, 0) == "A（異常なし）"
    assert vm._checkup_rating(5, 0) == "B（軽度所見）"
    assert vm._checkup_rating(5, 2) == "C（要確認）"


def test_analyze_calls_declared_api_is_not_unresolved():
    code = ('Private Declare PtrSafe Sub MoveMemory Lib "kernel32" '
            'Alias "RtlMoveMemory" (d As LongPtr, s As LongPtr, ByVal n As LongPtr)\r\n'
            'Sub A()\r\n'
            '    Call MoveMemory(1, 2, 3)\r\n'
            '    Call 存在しない子\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._analyze_calls(inv)
    assert [u[2] for u in res['unresolved']] == ['存在しない子']


def test_analyze_calls_object_method_and_qualified_call():
    code = ('Sub A()\r\n'
            '    Call wCompo.Export(sFilePath)\r\n'          # オブジェクトのメソッド＝対象外
            '    Call M1.子マクロ\r\n'                        # モジュール修飾の実マクロ＝辺
            'End Sub\r\n'
            'Sub 子マクロ()\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4},
                                            {'name': '子マクロ', 'lines': 2}])]}
    res = vm._analyze_calls(inv)
    assert res['unresolved'] == []
    assert res['edges'][('M1', 'A')] == {'子マクロ'}


def test_analyze_calls_dynamic_run_not_unresolved():
    code = ('Sub メニュー実行()\r\n'
            '    Application.Run "\'PERSONAL.XLSB\'!" & AAA\r\n'   # 動的＝未解決にしない
            '    Application.Run "\'" & ZZZ & "\'!" & AAA\r\n'     # 動的
            '    Application.Run "実在マクロ"\r\n'                  # 静的＝辺
            '    Application.Run "居ないマクロ"\r\n'                # 静的＝未解決
            'End Sub\r\n'
            'Sub 実在マクロ()\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'メニュー実行', 'lines': 6},
                                            {'name': '実在マクロ', 'lines': 2}])]}
    res = vm._analyze_calls(inv)
    assert [u[2] for u in res['unresolved']] == ['Run "居ないマクロ"']
    assert len(res['dynamic_runs']) == 2
    assert res['edges'][('M1', 'メニュー実行')] >= {'実在マクロ'}


def test_analyze_calls_commented_run_is_ignored():
    # コメントアウトされた Application.Run を未解決Callにしない（総点検で発見）
    code = ('Sub A()\r\n'
            "    ' Application.Run \"居ないマクロ\"  ←コメント行\r\n"
            '    x = "文字列の中の Application.Run も無視"\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._analyze_calls(inv)
    assert res['unresolved'] == []
    assert res['dynamic_runs'] == []


def test_checkup_diff_survives_missing_fields():
    # 履歴ファイルは外部データ＝フィールド欠損でも落ちない
    d = vm._checkup_diff({'time': 't0'}, {'time': 't1'})
    assert d['changed'] is False


def test_analyze_calls_onaction_macro_is_not_orphan():
    inv = {'modules': [_mod('M1', 1,
                            'Sub ボタン処理()\r\nEnd Sub\r\nSub 本当の孤立()\r\nEnd Sub\r\n',
                            [{'name': 'ボタン処理', 'lines': 2},
                             {'name': '本当の孤立', 'lines': 2}])],
           'onaction': [('メニュー', '角丸四角形 1', 'ボタン処理')]}
    res = vm._analyze_calls(inv)
    assert res['onaction'] == {'ボタン処理': ['メニュー/角丸四角形 1']}
    assert [n for _, n in res['orphans']] == ['本当の孤立']


def test_validate_vba_code_single_line_sub():
    # 1行書き Sub x(): End Sub を「End Sub 不足」と誤警告しない
    assert vm.validate_vba_code('Sub x(): End Sub\n') is True
    # 文字列内の ':' で誤分割しない
    assert vm.validate_vba_code('Sub y()\n    s = "a:b"\nEnd Sub\n') is True
    # 本当に対応が取れていないものは引き続き検出する
    assert vm.validate_vba_code('Sub z()\n', force=False) is False
    # 行末コメント内の ': Sub' / ': End Sub' を宣言として数えない
    assert vm.validate_vba_code("Sub a()\n    Call Foo ' 注: Sub Foo は別モジュール\nEnd Sub\n") is True
    assert vm.validate_vba_code("Sub b()\n    x = 1 ' TODO: End Sub まで見直す\nEnd Sub\n") is True


# ================================================================
# 2026-07-10 総点検の回帰テスト
# ================================================================

def test_parse_module_blocks_one_liner_and_property():
    # 1行完結 Sub（＋随伴Attribute）が次のプロシージャを巻き込まないこと
    bas = ('Attribute VB_Name = "M"\r\n'
           'Sub S1(): Call Main: End Sub\r\n'
           'Attribute S1.VB_ProcData.VB_Invoke_Func = "q\\n14"\r\n'
           'Sub Main()\r\n'
           'End Sub\r\n'
           'Property Get V() As Long\r\n'
           '    V = 1\r\n'
           'End Property\r\n'
           'Sub Last()\r\n'
           'End Sub\r\n')
    header, blocks, trailing = vm._parse_module_blocks(bas)
    assert [b['name'] for b in blocks] == ['S1', 'Main', 'V', 'Last']
    assert [b['kind'] for b in blocks] == ['sub', 'sub', 'property get', 'sub']
    assert any('VB_Invoke_Func' in ln for ln in blocks[0]['lines'])
    assert vm._write_module(header, blocks, trailing) == bas


def test_analyze_calls_run_module_qualified_resolves():
    # Application.Run "モジュール名.マクロ名" が未解決（C判定）に落ちないこと
    inv = {'modules': [
        _mod('M1', 1, 'Sub A()\r\n    Application.Run "M2.B"\r\nEnd Sub\r\n',
             [{'name': 'A', 'lines': 3}]),
        _mod('M2', 1, 'Sub B()\r\nEnd Sub\r\n', [{'name': 'B', 'lines': 2}]),
    ], 'onaction': []}
    res = vm._analyze_calls(inv)
    assert not res['unresolved'], res['unresolved']
    assert 'B' in res['edges'][('M1', 'A')]


def test_analyze_calls_onaction_module_qualified_resolves():
    # OnAction の「モジュール名.マクロ名」形式も孤立扱いしないこと
    inv = {'modules': [_mod('M1', 1, 'Sub C()\r\nEnd Sub\r\n',
                            [{'name': 'C', 'lines': 2}])],
           'onaction': [('メニュー', '四角形 1', 'M1.C')]}
    res = vm._analyze_calls(inv)
    assert res['onaction'] == {'C': ['メニュー/四角形 1']}
    assert not res['orphans']


def test_layout_button_bar_fits_content_width():
    # 幅の違うボタン（OK/キャンセル）でもバーが左にあふれないこと
    rows = [fl.button_bar(fl.ok("btnOK", "OK"), fl.cancel("btnCancel", "キャンセル"))]
    pl, cw, ch = fl.compute_layout(rows)
    assert min(l for e, l, t, w, h in pl) >= fl.STYLE['pad']
    assert max(l + w for e, l, t, w, h in pl) <= fl.STYLE['pad'] + cw + 0.01


def test_frame_autoname_deterministic():
    # hash() 乱数化で再buildのたび frame 名が変わらないこと
    a = fl.frame("設定", fl.row(fl.lbl("A"), fl.txt("t1")))
    b = fl.frame("設定", fl.row(fl.lbl("B"), fl.txt("t2")))
    c = fl.frame("別枠", fl.row(fl.lbl("C"), fl.txt("t3")))
    assert a['name'] == b['name']
    assert a['name'] != c['name']


def test_multipage_frame_children_present_in_layout():
    # page 内 frame の子が配置計画に含まれること（build 消失バグの回帰）
    rows = [fl.multipage("mp",
                         fl.page("基本", fl.frame("G", fl.row(fl.lbl("N"), fl.txt("t")))),
                         fl.page("詳細", fl.row(fl.lbl("M"), fl.txt("u"))))]
    pl, cw, ch = fl.compute_layout(rows)
    mp_e = next(e for e, *_ in pl if e['kind'] == 'multipage')
    frames = [ce for ce, *_ in mp_e['pages_layout'][0]['children'] if ce['kind'] == 'frame']
    assert frames and len(frames[0]['children']) == 2


# ================================================================
# vba_manager: Remove+Import の名前衝突ガード（2026-07-11 shu005→shu0051 事故の回帰）
# フェイク VBE で「Remove 遅延完了中の Import は連番付き別名になる」挙動を模す。
# 実 Excel での正常系は E2E（実機テスト）で担保。
# ================================================================

class _FakeComp:
    def __init__(self, name):
        self.Name = name


class _FakeVBComponents:
    """VBE の Import 挙動を模す: 同名が既に居ると連番付き別名で取り込まれる。

    ghost_clears_after_iters: 列挙がその回数を超えたらゴースト（遅延 Remove 中の
    旧モジュール）を消す。None なら永遠に残る（回復不能ケース）。
    """
    def __init__(self, names, base_name, ghost=None, ghost_clears_after_iters=None):
        self.comps = [_FakeComp(n) for n in names]
        self._base = base_name
        self._ghost = ghost
        self._clear_after = ghost_clears_after_iters
        self._iters = 0

    def __iter__(self):
        self._iters += 1
        if (self._ghost is not None and self._clear_after is not None
                and self._iters > self._clear_after):
            self.comps = [c for c in self.comps if c.Name != self._ghost]
            self._ghost = None
        return iter(list(self.comps))

    def Import(self, path):
        name = self._base
        if any(c.Name.lower() == name.lower() for c in self.comps):
            name = name + "1"          # VBE の連番リネーム
        c = _FakeComp(name)
        self.comps.append(c)
        return c


class _FakeWB:
    def __init__(self, components):
        self.VBProject = type("VBP", (), {"VBComponents": components})()


def test_import_verified_normal_returns_expected_name():
    comps = _FakeVBComponents(["OtherMod"], "TestMod")
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=0.05, rename_timeout=0.05, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["OtherMod", "TestMod"]


def test_import_verified_waits_for_ghost_and_avoids_collision():
    # 遅延 Remove 中のゴーストが待機中に消える → 衝突せず期待名で取り込まれる
    comps = _FakeVBComponents(["TestMod"], "TestMod",
                              ghost="TestMod", ghost_clears_after_iters=2)
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=3.0, rename_timeout=0.05, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["TestMod"]


def test_import_verified_collision_recovers_by_rename():
    # ゴーストが Import 後まで残る → TestMod1 で衝突 → 消滅を待って改名回復
    comps = _FakeVBComponents(["TestMod"], "TestMod",
                              ghost="TestMod", ghost_clears_after_iters=2)
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=0.01, rename_timeout=3.0, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["TestMod"]


def test_import_verified_unrecoverable_raises_not_silent_success():
    # ゴーストが消えない → 黙って成功にせず ModuleNameCollisionError
    comps = _FakeVBComponents(["TestMod"], "TestMod", ghost="TestMod")
    wb = _FakeWB(comps)
    try:
        vm._import_module_verified(wb, "x.bas", "TestMod",
                                   ghost_timeout=0.01, rename_timeout=0.05, settle=0)
        assert False, "ModuleNameCollisionError が飛ぶべき"
    except vm.ModuleNameCollisionError as ex:
        assert ex.expected_name == "TestMod"
        assert ex.actual_name == "TestMod1"
    # 衝突した別名側にコードが残っている（バックアップ再Importさせないための前提）
    assert any(c.Name == "TestMod1" for c in comps.comps)


def test_backup_reimport_only_when_import_actually_failed():
    """Import 成功後の Save 失敗で、バックアップを再 Import してはいけない（回帰）。

    2026-07-14 発見の実害: Remove+Import 系3経路の except が removed フラグしか
    見ておらず、Import が成功した後の wb.Save() が失敗すると「モジュールが消えた」と
    誤認してバックアップを重ね Import していた。期待名のモジュールは既に正しく
    存在するので _wait_component_gone は空振りし（15秒）、VB_Name 衝突で連番別名として
    取り込まれ、改名待ちも空振りして（20秒）例外になる——つまりツール自身が
    「旧コード入りの連番モジュール」を生み、35秒待たせたうえで失敗していた。

    復旧（バックアップの再 Import）は Import 自体が失敗したときだけ通ること。
    構文木で「復旧呼び出しを囲む if の条件に imported が出てくるか」を機械照合する。
    """
    import ast as _ast
    base = os.path.dirname(os.path.abspath(__file__))
    targets = ['vbam_vba.py', 'optimize_vba_modules.py']

    def _func_name(node):
        f = node.func
        return getattr(f, 'id', None) or getattr(f, 'attr', None) or ''

    total = 0
    for fname in targets:
        with open(os.path.join(base, fname), encoding='utf-8') as f:
            tree = _ast.parse(f.read())
        parents = {}
        for parent in _ast.walk(tree):
            for child in _ast.iter_child_nodes(parent):
                parents[child] = parent

        recoveries = [
            n for n in _ast.walk(tree)
            if isinstance(n, _ast.Call)
            and _func_name(n) == '_import_module_verified'
            and len(n.args) >= 2
            and getattr(n.args[1], 'id', '') == 'module_backup'
        ]
        for call in recoveries:
            names = set()
            node = call
            while node in parents:
                node = parents[node]
                if isinstance(node, _ast.If):
                    names |= {n.id for n in _ast.walk(node.test)
                              if isinstance(n, _ast.Name)}
            assert 'imported' in names, (
                f"{fname}: バックアップの再 Import が imported フラグで守られていない。\n"
                "  Import 成功後の Save 失敗で、ツール自身が連番モジュール"
                "（旧コード入り）を作る経路が復活している。")
        total += len(recoveries)

    # replace-procedure(Attribute経路) / replace-module / reorder / optimize の4経路
    assert total == 4, f"復旧経路の数が想定と違う（台帳の更新が必要）: {total}"


# ================================================================
# ダイアログ自動解除の報告（マクロ発火中のMsgBoxを黙って握りつぶさない）
# 2026-07-11: write-range→Worksheet_Change→MsgBox で無言ハングした実害の対策。
# COM/win32 部は実機E2Eで担保、ここは「検出時に報告文を出す」純粋部分の回帰。
# ================================================================

class _FakeWatcher:
    def __init__(self, count, last=""):
        self.count = count
        self.last_text = last


def test_dialog_note_empty_when_no_dialog():
    assert vm._dialog_watcher_note(_FakeWatcher(0), None) == ""
    assert vm._dialog_watcher_note(None, None) == ""


def test_dialog_note_safe_mode_reports_body():
    note = vm._dialog_watcher_note(_FakeWatcher(1, "B1が変わりました"), None)
    assert "1件" in note
    assert "安全側" in note
    assert "B1が変わりました" in note


def test_dialog_note_explicit_mode_says_specified_button():
    note = vm._dialog_watcher_note(_FakeWatcher(2, "確認"), "ok")
    assert "2件" in note
    assert "指定ボタン" in note


# --input-text（2026-08-23）: InputBox に値を入れて確定したら「何を入れたか」を必ず報告する
class _FakeWatcherWithInputs(_FakeWatcher):
    def __init__(self, count, last="", inputs=()):
        super().__init__(count, last)
        self.inputs = list(inputs)


def test_dialog_note_reports_entered_inputs_in_order():
    note = vm._dialog_watcher_note(_FakeWatcherWithInputs(2, "名前は？", ["太郎", "B2:C3"]), None)
    assert "2件" in note
    assert "--input-text" in note
    assert note.index("「太郎」") < note.index("「B2:C3」")


def test_dialog_note_without_inputs_attr_is_unchanged():
    # 旧来の watcher（inputs 属性なし）でも落ちず、入力欄の行も混ざらない
    note = vm._dialog_watcher_note(_FakeWatcher(1, "x"), "ok")
    assert "input-text" not in note


def test_run_macro_parser_collects_repeated_input_text():
    a = vm.build_parser().parse_args(
        ["run-macro", "M", "--input-text", "太郎", "--input-text", "B2:C3"])
    assert a.input_text == ["太郎", "B2:C3"]
    assert vm.build_parser().parse_args(["run-macro", "M"]).input_text is None
    assert vm.build_parser().parse_args(["test", "--input-text", "1"]).input_text == ["1"]
    assert vm.build_parser().parse_args(["rehearse", "M", "--input-text", "x"]).input_text == ["x"]


# ================================================================
# 2026-08-23: 停止条件（gate／関所）と --timeout の配線（COM 不要の部分）
# ================================================================

def test_gate_is_wired_with_defaults():
    table = vm._command_table()
    assert "gate" in table and "関所" in table and table["gate"] is table["関所"]
    ns, unknown = vm.build_parser().parse_known_args(["gate"])
    assert ns.command == "gate" and not unknown
    assert float(ns.timeout) == 120 and ns.module is None and not ns.keep and not ns.json
    ns, unknown = vm.build_parser().parse_known_args(
        ["関所", "a.xlsm", "合計", "--module", "M", "--addins", "--timeout", "30",
         "--auto-dialog", "ok", "--input-text", "x", "--keep", "--json"])
    assert not unknown and ns.posargs == ["a.xlsm", "合計"]
    assert ns.module == "M" and ns.addins and ns.timeout == "30" and ns.auto_dialog == "ok"
    assert ns.input_text == ["x"] and ns.keep and ns.json


# ================================================================
# 2026-08-28: 全体コンパイル（compile／全体コンパイル）と gate への組み込み
#   check（静的検査）は「呼ばれないマクロの中」を見ない。VBA はプロシージャ単位の
#   オンデマンド コンパイルなので、存在しない処理を呼ぶ Sub は実行するまで誰も気づかない。
#   実機では VBE の Id 578 を押して捕まえる。ここは COM 不要の判定部分の回帰。
# ================================================================

class _FakeCompileWatcher(_FakeWatcher):
    """_FakeWatcher に .stop() を足しただけ（_compile_vbproject は必ず止める）。"""

    def stop(self):
        pass


class _FakeCtls:
    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class _FakeCtl:
    """VBE の CommandBar コントロールのふり。Type=10 はポップアップ（デバッグ メニュー）。"""

    def __init__(self, id_=0, type_=1, children=(), enabled=True, disable_on_execute=False):
        self.Id = id_
        self.Type = type_
        self.Enabled = enabled
        self._children = list(children)
        self.executed = 0
        self._disable_on_execute = disable_on_execute

    @property
    def Controls(self):
        return _FakeCtls(self._children)

    def Execute(self):
        self.executed += 1
        if self._disable_on_execute:
            self.Enabled = False          # 通ったら「もう押す必要がない」＝False に落ちる


class _FakeBar:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Controls(self):
        return _FakeCtls(self._items)


class _FakeProj:
    def __init__(self, filename):
        self.FileName = filename


class _FakeVbe:
    def __init__(self, bars, projects=()):
        self.CommandBars = list(bars)
        self.VBProjects = list(projects)
        self.ActiveVBProject = None


class _FakeXl:
    def __init__(self, vbe):
        self.VBE = vbe


class _FakeWb:
    def __init__(self, fullname, name="本.xlsm"):
        self.FullName = fullname
        self.Name = name


def _compile_fixture(ctl, path=r"C:\x\本.xlsm"):
    """デバッグ メニュー（ポップアップ）の中に ctl を置いた VBE 一式を作る。"""
    menu = _FakeCtl(id_=1, type_=10, children=[_FakeCtl(id_=999), ctl])
    vbe = _FakeVbe([_FakeBar([menu])], [_FakeProj(path)])
    return _FakeXl(vbe), _FakeWb(path)


def test_find_vbe_compile_control_digs_into_popup():
    target = _FakeCtl(id_=578)
    xl, _ = _compile_fixture(target)
    assert vm._find_vbe_compile_control(xl.VBE) is target


def test_find_vbe_compile_control_returns_none_when_absent():
    xl, _ = _compile_fixture(_FakeCtl(id_=577))
    assert vm._find_vbe_compile_control(xl.VBE) is None


def test_compile_vbproject_says_already_when_button_disabled(monkeypatch):
    ctl = _FakeCtl(id_=578, enabled=False)
    xl, wb = _compile_fixture(ctl)
    res = vm._compile_vbproject(xl, wb)
    assert res["ok"] and res["state"] == "already"
    assert ctl.executed == 0          # 押す必要がないので押さない


def test_compile_vbproject_ok_when_button_goes_disabled(monkeypatch):
    monkeypatch.setattr(vm.time, "sleep", lambda *_: None)
    # 差し替えは実体のある vbam_vba 側に当てる（vba_manager 側は再エクスポートの別名）
    import vbam_vba
    monkeypatch.setattr(vbam_vba, "_start_dialog_watcher", lambda *_a, **_k: _FakeCompileWatcher(0))
    ctl = _FakeCtl(id_=578, enabled=True, disable_on_execute=True)
    xl, wb = _compile_fixture(ctl)
    res = vm._compile_vbproject(xl, wb)
    assert res["ok"] and res["state"] == "compiled"
    assert ctl.executed == 1


def test_compile_vbproject_fails_when_dialog_fired(monkeypatch):
    # コンパイルエラーは VBE がモーダルの MsgBox で知らせる＝本文を理由として持ち帰る
    monkeypatch.setattr(vm.time, "sleep", lambda *_: None)
    import vbam_vba
    monkeypatch.setattr(vbam_vba, "_start_dialog_watcher",
                        lambda *_a, **_k: _FakeCompileWatcher(1, "コンパイル エラー:\n\nSub または Function が定義されていません。"))
    ctl = _FakeCtl(id_=578, enabled=True)          # 落ちたので Enabled は True のまま
    xl, wb = _compile_fixture(ctl)
    res = vm._compile_vbproject(xl, wb)
    assert not res["ok"] and res["state"] == "error"
    assert "Sub または Function が定義されていません" in res["detail"]
    assert "\n" not in res["detail"]               # 1行に畳んで理由欄に載せる


def test_compile_vbproject_unavailable_is_not_a_failure():
    # 見られなかったことを「落ちた」にはしない（VBOM 未信頼のブックで全部 FAIL になるのを避ける）
    xl, wb = _compile_fixture(_FakeCtl(id_=577))   # Id 578 が無い環境
    res = vm._compile_vbproject(xl, wb)
    assert res["ok"] and res["state"] == "unavailable"

    class _NoVbe:
        @property
        def VBE(self):
            raise RuntimeError("VBOM が信頼されていません")

    res2 = vm._compile_vbproject(_NoVbe(), wb)
    assert res2["ok"] and res2["state"] == "unavailable"


def test_compile_command_is_wired():
    table = vm._command_table()
    assert "compile" in table and "全体コンパイル" in table
    assert table["compile"] is table["全体コンパイル"]
    ns, unknown = vm.build_parser().parse_known_args(["compile"])
    assert ns.command == "compile" and not unknown and not ns.json
    ns, unknown = vm.build_parser().parse_known_args(["全体コンパイル", "a.xlsm", "--json"])
    assert not unknown and ns.posargs == ["a.xlsm"] and ns.json


def test_timeout_option_parsed_on_run_macro_test_rehearse():
    # 既定は None（従来どおり無制限）。付けたときだけ効く＝既定の挙動を変えない
    p = vm.build_parser()
    assert p.parse_args(["run-macro", "M", "--timeout", "5"]).timeout == "5"
    assert p.parse_args(["run-macro", "M"]).timeout is None
    assert p.parse_args(["test", "--timeout", "5"]).timeout == "5"
    assert p.parse_args(["test"]).timeout is None
    assert p.parse_args(["rehearse", "M", "--timeout", "5"]).timeout == "5"
    assert p.parse_args(["rehearse", "M"]).timeout is None


def test_gate_rejects_extra_args_before_touching_excel(capsys):
    import argparse
    assert vm.cmd_gate(argparse.Namespace(posargs=["a", "b"])) is False
    assert "余分な引数" in capsys.readouterr().out


def test_call_watchdog_is_noop_without_seconds():
    for sec in (None, 0, -1, "abc"):
        w = vm._start_call_watchdog(sec)
        assert w.fired is False and w.seconds is None
        w.stop()


def test_call_watchdog_fires_after_seconds_and_checks_pid_identity():
    import time as _t
    # 存在しない PID＝EXCEL.EXE ではないので撃たない（fired は立つ）。COM 不要
    w = vm._start_call_watchdog(0.2, owned_pid=0x7FFFFFFF)
    assert w.fired is False
    _t.sleep(0.6)
    assert w.fired is True and w.seconds == 0.2
    w.stop()
    # 時間内に stop すれば立たない
    w = vm._start_call_watchdog(5, owned_pid=0x7FFFFFFF)
    w.stop()
    _t.sleep(0.1)
    assert w.fired is False


def test_vba_tolerant_pattern_ignores_case_and_spacing():
    # VBA は保存時に識別子の大小文字を初出に揃え、演算子前後の空白を整える。
    # code-replace の探し直しパターンは、その字面のずれを吸収し、別の意味には当たらない
    import vbam_vba
    p = vbam_vba._vba_tolerant_pattern('.Value - ActiveSheet')
    assert p.search('x.value - ActiveSheet.Range("B1").value')     # 大小文字（2026-08-23 実害）
    assert p.search('x.value-ActiveSheet')                          # 演算子前後の空白なし
    assert p.search('x.VALUE  -  activesheet')                      # 空白の数
    assert not p.search('x.value + ActiveSheet')                    # 別の演算子には当たらない
    q = vbam_vba._vba_tolerant_pattern('Range("A1").Value')
    assert q.search('range("a1").value')                            # 正規表現の特殊文字は字面どおり
    assert not q.search('range("Ax").value')


def test_json_line_sieve_keeps_human_lines_and_captures_json(capsys):
    import vbam_vba
    with vbam_vba._JsonLineSieve() as sieve:
        print("○ テストA  [Mod]  (0.01秒)")
        print('{"success": true, "total": 1}')
        print("結果: 1/1 成功")
    out = capsys.readouterr().out
    assert "○ テストA" in out and "結果: 1/1" in out
    assert '"success"' not in out
    assert sieve.last == {"success": True, "total": 1}


# ================================================================
# _collect_shapes: OnAction のブック修飾を落とさない（wiring の外部ブック判定の素）
# ================================================================

class _FakeShape:
    def __init__(self, name, onaction):
        self.Name = name
        self.Type = 1
        self.OnAction = onaction
        self.Left = 1.0
        self.Top = 2.0


def test_collect_shapes_keeps_onaction_book_qualifier():
    out = []
    vm._collect_shapes([_FakeShape("B1", "'秀 テスト.xlam'!Macro1"),
                        _FakeShape("B2", "Macro2"),
                        _FakeShape("B3", "ファイル一覧.xlsm!全検索開始")], out)
    assert out[0]["onaction"] == "Macro1"
    assert out[0]["onaction_book"] == "秀 テスト.xlam"
    assert out[1]["onaction"] == "Macro2"
    assert "onaction_book" not in out[1]
    assert out[2]["onaction"] == "全検索開始"
    assert out[2]["onaction_book"] == "ファイル一覧.xlsm"


# ================================================================
# snapshot-diff: 結合未走査の snapshot に疑似差分（結合追加/解除）を出さない
# ================================================================

def _write_snap(path, sheets):
    import json
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({"success": True, "book": "T.xlsm", "sheets": sheets},
                  f, ensure_ascii=False)


def _diff_args(old_path, new_path):
    import argparse
    return argparse.Namespace(posargs=[str(old_path), str(new_path)], max_opt=None)


def test_snapshot_diff_merged_unscanned_suppresses_pseudo_diff(tmp_path, capsys):
    # 旧=結合未走査 / 新=結合あり。セル差分でシートは表示されるが、
    # 「結合追加」の疑似差分は出さず、比較していない旨を注記する
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "1行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "x"}}],
                           "merged_skipped_cells": 99999}})
    _write_snap(p2, {"S": {"dims": "1行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "y"}}],
                           "merged": ["$A$5:$B$5"]}})
    assert vm.cmd_snapshot_diff(_diff_args(p1, p2)) is True
    out = capsys.readouterr().out
    assert "結合追加" not in out
    assert "結合の差分は比較していない" in out
    assert "セル変更: 1件" in out


def test_snapshot_diff_merged_unscanned_note_survives_hidden_sheet(tmp_path, capsys):
    # 結合以外に差分がなくシート自体が非表示でも、「比較できていない」事実は落とさない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "1行 x 1列", "cells": [],
                           "merged_skipped_cells": 99999}})
    _write_snap(p2, {"S": {"dims": "1行 x 1列", "cells": [],
                           "merged": ["$A$1:$B$1"]}})
    assert vm.cmd_snapshot_diff(_diff_args(p1, p2)) is True
    out = capsys.readouterr().out
    assert "結合の差分を比較できていません" in out
    assert "S" in out.split("比較できていません:")[-1]
    assert "差分なし" in out


def test_trailing_spacer_counts_fully():
    # 末尾 spacer が gap_y ぶん目減りしないこと（中間の spacer と同じ意味論）
    base = [fl.row(fl.lbl("X"), fl.txt("tX"))]
    _, _, h_a = fl.compute_layout(base)
    _, _, h_b = fl.compute_layout(base + [fl.spacer(24)])
    assert abs((h_b - h_a) - (fl.STYLE['gap_y'] + 24)) < 0.01


# ================================================================
# vba_manager: VBA 識別子ガード（先頭 _ の Sub 注入事故 `_tmp検証` の回帰）
# ================================================================

def test_check_vba_identifier_rejects_leading_underscore():
    assert vm.check_vba_identifier("_tmp検証") is not None


def test_check_vba_identifier_rejects_leading_digit():
    assert vm.check_vba_identifier("1テスト") is not None


def test_check_vba_identifier_rejects_symbols():
    assert vm.check_vba_identifier("foo-bar") is not None


def test_check_vba_identifier_accepts_normal_names():
    for name in ("tmp検証", "テスト検証", "Btn_Click", "UserForm_Initialize", "A1"):
        assert vm.check_vba_identifier(name) is None, name


def test_find_invalid_procedure_names_hits_declaration():
    code = "Sub _tmp検証()\nEnd Sub\n"
    hits = vm._find_invalid_procedure_names(code)
    assert len(hits) == 1
    assert hits[0][1] == "_tmp検証"


def test_find_invalid_procedure_names_ignores_comments_and_events():
    code = ("' Sub _コメントは対象外()\n"
            "Private Sub CommandButton1_Click()\n"
            "End Sub\n"
            "Property Get 値()\n"
            "End Property\n")
    assert vm._find_invalid_procedure_names(code) == []


def test_validate_vba_code_rejects_underscore_name():
    assert vm.validate_vba_code("Sub _tmp検証()\nEnd Sub\n") is False


def test_validate_vba_code_accepts_valid_japanese_name():
    assert vm.validate_vba_code("Sub tmp検証()\nEnd Sub\n") is True


def test_check_bas_rejects_invalid_identifier(tmp_path):
    p = tmp_path / "m.bas"
    p.write_bytes('Attribute VB_Name = "M"\r\nSub _tmp検証()\r\nEnd Sub\r\n'.encode('cp932'))
    assert vm._check_bas_one(str(p)) is False


# ================================================================
# 注入経路の台帳（機械検査）
# ================================================================

def test_injection_route_ledger():
    """VBA へコードを入れる注入プリミティブの台帳。

    2026-07-12 の識別子ガード配線で一番時間を食ったのは「注入経路の洗い出し」
    だった。この台帳が現物と一致する限り、次回の穴塞ぎで経路の再調査は不要。
    新しい注入点を足す手順: ①ガードを配線する（または安全な理由を確認する）
    ②下の EXPECTED に理由コメントつきで登録する。
    未登録の注入点が現れたらこのテストが落ちる＝ガード無しの新経路の検知器。
    """
    import re as _re
    base = os.path.dirname(os.path.abspath(__file__))
    PRIM = _re.compile(r'\.AddFromString\(|\.InsertLines\(|VBComponents\.Import\(')
    DEF = _re.compile(r'^\s*def\s+(\w+)')
    EXPECTED = {
        # (ファイル, 関数): ガードの所在／安全な理由
        ('form_builder.py', 'inject_vba'),          # 注入前に識別子検査（既存コード削除より前）
        ('form_inspect.py', 'render_form_png'),     # 機械固定名 tmpFormInspect* のみ＝安全
        ('form_layout.py', 'build_form'),           # 起動マクロ名を check_vba_identifier で検査
        ('form_tool.py', 'cmd_copy_form'),          # 新フォーム名を check_vba_identifier で検査
        ('live_sync_vba.py', 'update_module'),      # 既存コード往復＋固定名マクロ追記のみ＝新規名の流入なし
                                                    # （2026-07-14: 失敗時に旧コードを書き戻すため関数へ切り出し）
        # 2026-07-14: optimize_vba_modules.apply_module は素の VBComponents.Import をやめ、
        # vbam_core._import_module_verified（実名検証つき）に委譲したので注入経路ではなくなった
        # 2026-07-12 分割: vba_manager.py の実装は vbam_core/vbam_vba 等へ移動（入口は不変）
        ('vbam_core.py', '_import_module_verified'),  # 取込の中央関数（名前衝突ガード）。内容の識別子検査は呼び元
        ('vbam_vba.py', 'cmd_replace_procedure'),     # validate_vba_code で識別子検査
        ('vbam_vba.py', 'cmd_add_procedure'),         # validate_vba_code で識別子検査
        ('vbam_vba.py', 'cmd_test'),                  # 機械固定名 VMT_n ハーネス＝安全
        # 2026-08-23: run-macro のハーネス（実行時エラーをダイアログにせず番号・説明で持ち帰る）。
        # 生成コードに埋めるのは呼び出し名だけで、check_vba_identifier を通し、さらに
        # 宣言（Sub/Function 名）と実在モジュール名に一致したときしか注入しない。固定名 VMR
        ('vbam_vba.py', '_prepare_run_harness'),
        # 2026-08-17: form-to-vba（UserForm→作成マクロ）。捨てブックに作成マクロを入れて
        # 実際に組み立て、元と照合するだけの検証経路。入れる先は使い捨ての新規ブックで、
        # 中身は自分が生成した VBA（フォーム名・プロシージャ名とも元フォーム由来）＝
        # 外から任意の名前が流入しない。元ブックには一切書かない
        ('vbam_form2vba.py', '_f2v_verify'),
        ('test_e2e_com.py', 'book'),                  # E2E の使い捨てブック生成。固定リテラルの
                                                      # テストコードのみ＝外から名前が流入しない
        ('test_e2e_com.py', 'gate_book'),             # 同上（2026-08-23 gate / rehearse 用。固定リテラルのみ）
        # 2026-09-03: agent --fire（macro）の練習台。一時フォルダの使い捨てブックに、固定リテラルの
        # 壊れたマクロ（_FIRE_MACRO_CASES）を植えて修理させるだけ＝外から名前が流入しない。
        # AI が返す書き換えは replace-procedure / code-replace 経由（識別子検査はそちら）
        # 2026-09-11: 実射は vbam_fire.py へ切り出した（中身は同じ）
        ('vbam_fire.py', '_fire_macro_cases'),
        # 2026-09-11 夜: agent --exam --by-macro。AI の代わりにマクロで試験を撃つとき、道具が起こした Excel の
        # 使い捨ての新しいブックへ .bas を取り込む。取り込む前に exam() が _check_bas_one（文字コード・識別子）を通す。
        # 人のブックには書かない（台本は子プロセスの文字列＝_probe_source の中）
        ('vbam_exam.py', '_probe_source'),
        # 2026-09-17: compile の行名指しの E2E。使い捨てブック（book fixture）に固定リテラルの壊れた Sub を
        # 植えて全体コンパイルを押すだけ＝外から名前が流入しない。人のブックには書かない
        ('test_e2e_com.py', 'test_compile_names_the_failing_line'),
        # 2026-09-17: 鍛える回路（agent --forge）。AI が書いたマクロを、道具が起こした Excel の使い捨ての新しいブックへ
        # .bas で取り込み、撃つ前のブックの写しに撃って値で突き合わせる。取り込む前に _check_bas（文字コード・識別子）と
        # _validate_code（Sub 1 本・引数なし・Function なし・既定の名前と重ねない）を通す。人のブックには書かない。
        # 合格したコードの登録は add-procedure／replace-procedure 経由（識別子検査はそちら）
        # 2026-09-17 夕: 別の表でも試すため、1 つの Excel で写しを順に撃つ _run_on_copies に移した（中身は同じ）
        ('vbam_forge.py', '_run_on_copies'),
        # 2026-09-17 夕: 先撃ちの試し撃ち。人の Excel で撃つ前に、道具が起こした Excel の使い捨てのブックへ、人の Excel に
        # 既に登録されている「表の整理」の本文と固定の台（鍛冶_撃つN）を入れ、ブックの写しに撃つ。台に埋める名前は
        # 登録簿（Sub 宣言の正規表現で取った名前）と既定のマクロ名だけ。人のブックには書かない
        ('vbam_forge.py', 'rehearse_steps'),
        # 2026-09-17 夜: 修理の試験（agent --mend）の子プロセス。前の表の写し（一時の置き場）に、鍛えた台帳の
        # マクロを Python が 1〜2 行だけ壊したコードを入れて、macro モードに直させる。Sub 名・識別子は鍛えたとき
        # （_validate_code・check-bas を通った物）のまま変えない。人のブックには書かない（台本は _probe_source の中）
        ('vbam_mend.py', '_probe_source'),
        # 同じ夜: 表を整える を修理の試験に載せる正解づくり。使い捨ての新しいブックに 表の整理.bas から抜いた元の Sub と
        # 固定の台（鍛冶_撃つ・名前は Sub 宣言から取った 表を整える だけ）を入れ、試験の表の写しに撃つ。人のブックには書かない
        ('vbam_mend.py', '_capture'),
    }
    found = set()
    for fname in sorted(os.listdir(base)):
        if not fname.endswith('.py') or fname == os.path.basename(__file__):
            continue
        cur = '(module)'
        with open(os.path.join(base, fname), encoding='utf-8', errors='replace') as f:
            for line in f:
                m = DEF.match(line)
                if m:
                    cur = m.group(1)
                if PRIM.search(line):
                    found.add((fname, cur))
    new_routes = found - EXPECTED
    # 台帳のファイルが現物に無いのは「経路が消えた」ではなく「ファイルごと無い」
    # （例: live_sync_vba.py は 2026-07-16 にローカル専用へ退役＝公開リポに入らない。
    #   クローン先ではファイル自体が無いのが正常）。ファイルは在るのに注入が
    # 消えた場合だけ「台帳の掃除」として検知する
    gone = {e for e in (EXPECTED - found)
            if os.path.exists(os.path.join(base, e[0]))}
    assert not new_routes, (
        f"台帳に無い注入経路: {sorted(new_routes)}\n"
        "  ガード（check_vba_identifier / _find_invalid_procedure_names）を配線してから"
        "台帳に理由コメントつきで登録すること。")
    assert not gone, f"台帳にあるのに現物に無い注入経路（台帳の掃除が必要）: {sorted(gone)}"


def test_copy_form_rejects_leading_underscore():
    # 旧regexは先頭 _ を素通しした（\w−数字＝英字＋_）。COM接続前に止まることの回帰
    import argparse
    import pytest
    import form_tool
    with pytest.raises(SystemExit):
        form_tool.cmd_copy_form(argparse.Namespace(form="F_X", new="_F_X2"))


# ================================================================
# 2026-07-15 点検: 「1行完結 Sub」の残存経路（delete-procedure / call-graph）
# ================================================================

class _FakeCodeModule:
    """ProcStartLine/ProcCountLines の領域に次の宣言行が食い込む状況を再現する。"""
    def __init__(self, lines):
        self._lines = lines

    def Lines(self, start, count):
        return '\r\n'.join(self._lines[start - 1:start - 1 + count])


def test_narrow_proc_range_excludes_next_declaration():
    # 1行完結 Sub の直後のプロシージャ。ProcCountLines が次の宣言行まで含んでしまう
    # （count=3 が「Sub 次の処理()」まで食い込んだ状態）
    cm = _FakeCodeModule([
        'Sub 一行版(): Call Main: End Sub',   # 1行目
        '',                                    # 2行目
        'Sub 次の処理()',                      # 3行目 ← 食い込み
        '    MsgBox 1',
        'End Sub',
    ])
    start, count = vm._narrow_proc_range(cm, 1, 3)
    assert (start, count) == (1, 1), "次の宣言行を削除範囲に含めてはいけない"
    # 絞った範囲は1行完結 Sub 自身のみ
    assert cm.Lines(start, count) == 'Sub 一行版(): Call Main: End Sub'


def test_narrow_proc_range_keeps_normal_proc_intact():
    cm = _FakeCodeModule([
        'Sub 普通()',
        '    MsgBox 1',
        'End Sub',
        '',
    ])
    start, count = vm._narrow_proc_range(cm, 1, 4)
    assert (start, count) == (1, 3)


def test_inline_body_after_decl():
    f = vm._inline_body_after_decl
    # 1行完結 Sub → 本体を返す
    raw = 'Sub X(): Call Main: End Sub'
    assert f(raw, raw.index('X') + 1).strip() == 'Call Main: End Sub'
    # 普通の宣言行 → 空
    raw2 = 'Sub Y(a As String)'
    assert f(raw2, raw2.index('Y') + 1) == ''
    # 引数の既定値の文字列に ':' や ')' があっても誤らない
    raw3 = 'Sub Z(Optional s As String = "a:b)"): Call Main: End Sub'
    assert f(raw3, raw3.index('Z') + 1).strip() == 'Call Main: End Sub'
    # 戻り型指定つき Function の1行完結
    raw4 = 'Function F() As String: F = "x": End Function'
    assert 'F = ' in f(raw4, raw4.index('F') + 1)


def _inv(code, procs):
    return {'modules': [{'name': 'Mod1', 'type': 1, 'code': code,
                         'procs': [{'name': p} for p in procs]}],
            'onaction': []}


def test_analyze_calls_sees_body_of_one_line_sub():
    # 「Sub ボタン18_Click(): Call 印刷実行: End Sub」の Call を見落とすと、
    # 印刷実行が「どこからも呼ばれていない」に誤って載る
    code = '\r\n'.join([
        'Sub ボタン18_Click(): Call 印刷実行: End Sub',
        '',
        'Sub 印刷実行()',
        '    MsgBox 1',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['ボタン18_Click', '印刷実行']))
    assert ('Mod1', 'ボタン18_Click') in res['edges'], "1行完結Subの Call が計上されていない"
    assert '印刷実行' in res['edges'][('Mod1', 'ボタン18_Click')]
    assert '印刷実行' not in res['orphans'], "呼ばれているのに孤立扱いになっている"


def test_analyze_calls_detects_typo_in_one_line_sub():
    # 1行完結 Sub の中の誤記（印刷実効）が未解決Callとして検出されること
    code = '\r\n'.join([
        'Sub ボタン18_Click(): Call 印刷実効: End Sub',
        '',
        'Sub 印刷実行()',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['ボタン18_Click', '印刷実行']))
    names = [u[2] for u in res['unresolved']]
    assert '印刷実効' in names, "1行完結Subの中の誤記が検出されていない"


def test_analyze_calls_one_line_sub_does_not_capture_following_lines():
    # 1行完結 Sub は宣言行で閉じている。後続行の呼び出しをその Sub の名で
    # 計上し続けてはいけない（cur が閉じられているか）
    code = '\r\n'.join([
        'Sub 一行版(): Call A: End Sub',
        '',
        'Sub 別の処理()',
        '    Call B',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['一行版', '別の処理', 'A', 'B']))
    assert 'B' not in res['edges'].get(('Mod1', '一行版'), set()), \
        "1行完結Subの後続行が、その Sub の呼び出しとして計上されている"
    assert 'B' in res['edges'][('Mod1', '別の処理')]


# ================================================================
# 2026-07-15 点検: snapshot-diff が「読めなかった」を「消えた」と誤報する件
# （ヘルパーは既存の _write_snap(path, sheets) / _diff_args(old, new) を使う）
# ================================================================

def test_snapshot_diff_does_not_report_unreadable_shapes_as_deleted(tmp_path, capsys):
    # 旧: 図形3つを正常に読めた / 新: 図形の列挙に失敗（shapes_error だけ）
    # → 図形は1つも消えていない。「図形削除: 3件」と言ってはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "ボタン1"}, {"name": "ボタン2"},
                                      {"name": "ボタン3"}]}})
    _write_snap(p2, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes_error": "COM エラー: 図形を列挙できません"}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "図形削除" not in out, "読めなかっただけの図形を「削除」と誤報している"
    assert "比較していません" in out or "比較できていません" in out


def test_snapshot_diff_never_claims_match_when_unreadable(tmp_path, capsys):
    # 両側とも図形が読めない → 差分ゼロに見えるが「一致」と断言してはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    info = {"dims": "10行 x 3列", "cells": [], "shapes_error": "COM エラー"}
    _write_snap(p1, {"S": dict(info)})
    _write_snap(p2, {"S": dict(info)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "一致" not in out, "読めていないのに「一致」と断言している"
    assert "比較できていません" in out


def test_snapshot_diff_does_not_report_unreadable_cells_as_deleted(tmp_path, capsys):
    # 旧: セルを読めた / 新: セルの読み取りに失敗 → 「セル削除」と言ってはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "2行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "あ"}},
                                     {"r": 2, "c": {"A": "い"}}]}})
    _write_snap(p2, {"S": {"dims": "2行 x 1列", "cells": [],
                           "cells_error": "COM エラー: 値を読めません"}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "セル削除" not in out, "読めなかっただけのセルを「削除」と誤報している"


def test_snapshot_diff_detects_shape_resize(tmp_path, capsys):
    # ボタンを 1x1pt に潰す変更を「差分なし」と言ってはいけない（図形のサイズの目）
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "実行ボタン", "l": 10, "t": 10,
                                       "w": 96, "h": 24}]}})
    _write_snap(p2, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "実行ボタン", "l": 10, "t": 10,
                                       "w": 1, "h": 1}]}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "図形変更" in out, "図形のサイズ変更を見落としている"
    assert "大きさ" in out


def test_snapshot_diff_match_message_does_not_overclaim(tmp_path, capsys):
    # 本当に一致しているときも、見ていない書式まで一致したかのように言わない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    info = {"dims": "1行 x 1列", "cells": [{"r": 1, "c": {"A": "x"}}]}
    _write_snap(p1, {"S": dict(info)})
    _write_snap(p2, {"S": dict(info)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "差分なし" in out
    assert "書式" in out, "書式が比較対象外であることを明示していない"


# ================================================================
# 2026-07-15: 書式の目（四つ目の目）。クリーン化で書式が飛んでも
# 「差分なし（一致）」と言っていたのを塞いだ回帰テスト
# ================================================================

def _fmt_sheet(sheet=None, cols=None, rows=None):
    """format ブロックつきのシート情報を作る"""
    return {"dims": "11行 x 4列", "cells": [{"r": 1, "c": {"A": "商品"}}],
            "format": {"sheet": sheet or {}, "cols": cols or {}, "rows": rows or {}}}


def test_snapshot_diff_detects_wiped_formatting(tmp_path, capsys):
    # クリーン化マクロが書式だけを吹き飛ばした（値・結合・図形は無傷）。
    # 従来はこれを「差分なし（一致）」と報告していた＝ポスター.xlsm 破壊の再演
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"売上表": _fmt_sheet(
        sheet={"bold": None, "border": 1, "fill": 0, "numfmt": None},
        cols={"A": {"bold": None, "border": 1, "w": 18.0}},
        rows={"1": {"bold": True, "border": 1, "h": 24.0}})})
    _write_snap(p2, {"売上表": _fmt_sheet(
        sheet={"bold": False, "border": -4142, "fill": 16777215, "numfmt": "G/標準"},
        cols={"A": {"bold": False, "border": -4142, "w": 8.42}},
        rows={"1": {"bold": False, "border": -4142, "h": 18.0}})})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" in out, "書式が吹き飛んだのに検出していない"
    assert "罫線" in out, "罫線の消失を検出していない"
    assert "一致" not in out
    # シート全体の一撃を列・行ぶん繰り返して騒がしくしない（読めない検分は役に立たない）
    assert "列幅" in out, "局所的な変化（列幅）は出すこと"


def test_snapshot_diff_format_noise_is_suppressed(tmp_path, capsys):
    # シート全体で起きた変化を、各列でも繰り返さない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    cols_old = {c: {"border": 1, "w": 8.38} for c in "ABCD"}
    cols_new = {c: {"border": -4142, "w": 8.38} for c in "ABCD"}
    _write_snap(p1, {"S": _fmt_sheet(sheet={"border": 1}, cols=cols_old)})
    _write_snap(p2, {"S": _fmt_sheet(sheet={"border": -4142}, cols=cols_new)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "シート全体: 罫線" in out
    # A列〜D列で同じ内容を繰り返さない
    assert "A列: 罫線" not in out, "シート全体の変化を列ぶん繰り返している（騒がしい）"


def test_snapshot_diff_format_local_change_survives(tmp_path, capsys):
    # 局所的な変化（その列だけ違う）は、抑制に巻き込まれず出ること
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"border": 1},
                                     cols={"A": {"border": 1}, "B": {"border": 1}})})
    _write_snap(p2, {"S": _fmt_sheet(sheet={"border": -4142},
                                     cols={"A": {"border": -4142}, "B": {"border": 9}})})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "B列: 罫線" in out, "列だけの局所的な変化が消えている"
    assert "A列: 罫線" not in out


def test_snapshot_diff_does_not_report_unreadable_format(tmp_path, capsys):
    # 書式を読めなかっただけの snapshot を「書式が変わった」と誤報しない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"bold": True})})
    s2 = {"dims": "11行 x 4列", "cells": [{"r": 1, "c": {"A": "商品"}}],
          "format_error": "COM エラー: 書式を読めません"}
    _write_snap(p2, {"S": s2})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" not in out, "読めなかっただけの書式を「変更」と誤報している"
    assert ("比較していません" in out or "比較できていません" in out)
    assert "一致" not in out, "書式を読めていないのに「一致」と言っている"


def test_snapshot_diff_format_missing_on_one_side(tmp_path, capsys):
    # 片方が --no-format で採られている → 書式は比較できないと明示する
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"bold": True})})
    _write_snap(p2, {"S": {"dims": "11行 x 4列",
                           "cells": [{"r": 1, "c": {"A": "商品"}}]}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" not in out
    assert ("比較していません" in out or "比較できていません" in out)


# ================================================================
# open / close / rehearse（2026-07-15 追加）: COM に触る前の門番
# ================================================================

import argparse as _ap


def test_open_requires_path(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=[])) is False
    assert "使い方" in capsys.readouterr().out


def test_open_rejects_multiple_paths(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=["a.xlsm", "b.xlsm"])) is False
    assert "1つだけ" in capsys.readouterr().out


def test_open_rejects_missing_file(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=["__vbam_unit_no_such__.xlsm"])) is False
    assert "見つかりません" in capsys.readouterr().out


def test_close_requires_name(capsys):
    ns = _ap.Namespace(posargs=[], save_flag=True, no_save_flag=False, yes=True)
    assert vm.cmd_close(ns) is False
    assert "使い方" in capsys.readouterr().out


def test_close_requires_save_policy(capsys):
    # 保存方針の明示は鎧の一部。無指定も両指定も門前払いする
    ns = _ap.Namespace(posargs=["a.xlsm"], save_flag=False, no_save_flag=False, yes=True)
    assert vm.cmd_close(ns) is False
    out = capsys.readouterr().out
    assert "--save" in out and "--no-save" in out

    ns = _ap.Namespace(posargs=["a.xlsm"], save_flag=True, no_save_flag=True, yes=True)
    assert vm.cmd_close(ns) is False


def test_close_not_open_book_fails(capsys):
    # 開いていないブック名は「開いていません」で止まる（保存方針が正しくても）
    ns = _ap.Namespace(posargs=["__vbam_unit_no_such__.xlsm"],
                       save_flag=False, no_save_flag=True, yes=True)
    assert vm.cmd_close(ns) is False
    assert "開いていません" in capsys.readouterr().out


def test_close_form_rejects_multiple_targets(capsys):
    ns = _ap.Namespace(posargs=["A", "B"], list_flag=False, wait_opt=None)
    assert vm.cmd_close_form(ns) is False
    assert "1つだけ" in capsys.readouterr().out


def test_close_form_pick_exact_then_partial():
    # キャプション選別は純粋関数：完全一致（大小無視）が優先、無ければ部分一致、空指定は全部
    forms = [(1, "Excelコンボ"), (2, "AIマクロ"), (3, "Excelコンボ（裏）")]
    assert vm._pick_form_windows(forms, "") == forms
    assert vm._pick_form_windows(forms, "excelコンボ") == [(1, "Excelコンボ")]
    assert vm._pick_form_windows(forms, "コンボ") == [(1, "Excelコンボ"), (3, "Excelコンボ（裏）")]
    assert vm._pick_form_windows(forms, "無い") == []


def test_close_form_is_wired():
    assert "close-form" in vm._command_table()
    parser = vm.build_parser()
    ns, unknown = parser.parse_known_args(["close-form", "Excelコンボ", "--list", "--wait", "2"])
    assert ns.command == "close-form" and ns.list_flag and ns.wait_opt == "2" and not unknown
    assert ns.posargs == ["Excelコンボ"]


def test_rehearse_requires_macro_name(capsys):
    assert vm.cmd_rehearse(_ap.Namespace(posargs=[])) is False
    assert "使い方" in capsys.readouterr().out


def test_new_commands_are_wired():
    # コマンド表と argparse の両方に配線されていること（表だけ・パーサだけの片肺を防ぐ）
    table = vm._command_table()
    for name in ("open", "close", "rehearse", "予行演習"):
        assert name in table
    parser = vm.build_parser()
    ns, unknown = parser.parse_known_args(
        ["close", "a.xlsm", "--no-save", "-y"])
    assert ns.command == "close" and ns.no_save_flag and ns.yes and not unknown
    ns, unknown = parser.parse_known_args(
        ["rehearse", "マクロA", "--addins", "--discard", "--max", "5"])
    assert ns.command == "rehearse" and ns.addins and ns.discard
    assert ns.max_opt == "5" and not unknown
    ns, unknown = parser.parse_known_args(["open", "a.xlsm"])
    assert ns.command == "open" and not unknown


# ================================================================
# vba_mcp_server: AI作業窓の関所（実行ログ・承認ゲート・差分確認）
# ================================================================
# 2026-08-26 追加。この層は「AI に任せる範囲が広がったこと」への備えなので、
# 壊れても普通の使い方では気づけない（AI_WIN_LOG が無いと全部素通しするため）。
# 素通しの向き——窓に聞けない異常時はフェイルオープン（方針A）、返事が無いときは
# フェイルクローズ（実行しない）——まで含めて機械で押さえる。
import re as _re                                                    # noqa: E402
import tempfile as _tf                                              # noqa: E402
import threading as _th                                             # noqa: E402
import time as _time                                                # noqa: E402

import pytest as _pytest                                            # noqa: E402

_cwd_before_srv_import = os.getcwd()
try:
    import vba_mcp_server as _srv
except Exception:                    # mcp SDK 未導入の環境ではこの節ごと飛ばす
    _srv = None
finally:
    # vba_mcp_server は import 時に os.chdir(SCRIPT_DIR) する。他のテストへ
    # 作業ディレクトリの変更を持ち込まないよう、ここで必ず戻す。
    os.chdir(_cwd_before_srv_import)

_needs_srv = _pytest.mark.skipif(
    _srv is None, reason="mcp SDK 未導入のため vba_mcp_server を import できない")


def _aiwin(tmp_path, monkeypatch):
    """AI作業窓から起動された子セッションを装う（AI_WIN_LOG を張る）。"""
    log = tmp_path / "ai_win.log"
    monkeypatch.setenv("AI_WIN_LOG", str(log))
    return log


def _answer_when_asked(log, verdict, encoding="utf-8", capture=None):
    """フォーム側ポーリングの代役。ask が置かれたら answer を書く。

    capture に dict を渡すと、ask に書かれた本文を capture['body'] に残す。
    """
    d = os.path.dirname(str(log)) or "."
    ask = os.path.join(d, "ai_win_ask.txt")
    ans = os.path.join(d, "ai_win_answer.txt")

    def _poll():
        deadline = _time.time() + 5
        while _time.time() < deadline:
            if os.path.exists(ask):
                if capture is not None:
                    try:
                        with open(ask, encoding="utf-8") as f:
                            capture["body"] = f.read()
                    except OSError:
                        pass
                with open(ans, "w", encoding=encoding) as f:
                    f.write(verdict)
                return
            _time.sleep(0.01)

    t = _th.Thread(target=_poll, daemon=True)
    t.start()
    return t


# ---- _log_call（実行ログ） ----

@_needs_srv
def test_log_call_writes_nothing_without_env(tmp_path, monkeypatch):
    monkeypatch.delenv("AI_WIN_LOG", raising=False)
    log = tmp_path / "ai_win.log"
    _srv._log_call("write-range A1 999")
    assert not log.exists()
    assert not list(tmp_path.iterdir())


@_needs_srv
def test_log_call_appends_time_and_command(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    _srv._log_call("  list-modules  ")
    _srv._log_call("read-range A1:B2")
    lines = log.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2                       # 上書きでなく追記
    assert _re.match(r"^\d\d:\d\d:\d\d  list-modules$", lines[0])
    assert lines[1].endswith("  read-range A1:B2")


@_needs_srv
def test_log_call_truncates_long_command(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    _srv._log_call("grep " + "あ" * 300)
    body = log.read_text(encoding="utf-8").rstrip("\n")
    assert len(body.split("  ", 1)[1]) == 160    # 窓の1行に収める


@_needs_srv
def test_log_call_swallows_write_errors(tmp_path, monkeypatch):
    # 書けない場所を指されても、本来の仕事（コマンド実行）を巻き込んで落とさない
    monkeypatch.setenv("AI_WIN_LOG", str(tmp_path / "no" / "such" / "x.log"))
    _srv._log_call("list")


# ---- GATE_COMMANDS / _gate_check（承認ゲート） ----

@_needs_srv
def test_gate_commands_are_exactly_the_irreversible_five():
    # 線引きは「読み/書き」ではなく「戻せる/戻せない」。退避→restore で戻せる
    # 置換系をここへ足すと、線引きが崩れたことに誰も気づけない
    assert _srv.GATE_COMMANDS == {
        "write-range", "clear-range", "delete-module", "save-as", "close"}
    assert "replace-procedure" not in _srv.GATE_COMMANDS
    assert "add-procedure" not in _srv.GATE_COMMANDS
    assert "restore" not in _srv.GATE_COMMANDS
    assert _srv.GATE_WAIT == 180


@_needs_srv
def test_gate_check_is_noop_without_env(monkeypatch):
    monkeypatch.delenv("AI_WIN_LOG", raising=False)
    assert _srv._gate_check("write-range A1 999") is None
    assert _srv._gate_check("delete-module shu009 -y") is None


@_needs_srv
def test_gate_check_passes_non_gated_commands(tmp_path, monkeypatch):
    _aiwin(tmp_path, monkeypatch)
    # 対象外は窓に聞かずその場で通す（聞きに行くと 180 秒固まる）
    assert _srv._gate_check("read-range A1:B2") is None
    assert _srv._gate_check("replace-procedure --module shu003 -y") is None
    assert _srv._gate_check("list-modules") is None


@_needs_srv
def test_gate_check_blocks_when_user_rejects(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    _answer_when_asked(log, "no")
    msg = _srv._gate_check("write-range A1 999", wait=5)
    assert msg is not None
    assert "write-range" in msg and "却下" in msg
    assert "実行していません" in msg          # AI に「やっていない」と伝わる文面


@_needs_srv
def test_gate_check_proceeds_when_user_approves(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    _answer_when_asked(log, "yes")
    assert _srv._gate_check("clear-range A1:B9", wait=5) is None


@_needs_srv
def test_gate_check_accepts_bom_answer(tmp_path, monkeypatch):
    # 答えを書くのは VBA 側。BOM 付きで書かれても yes と読めること
    log = _aiwin(tmp_path, monkeypatch)
    _answer_when_asked(log, "yes", encoding="utf-8-sig")
    assert _srv._gate_check("save-as out.xlsm", wait=5) is None


@_needs_srv
def test_gate_check_fails_closed_when_no_answer(tmp_path, monkeypatch):
    # 返事が無いときは実行しない側に倒す（黙って実行するのが最悪）
    _aiwin(tmp_path, monkeypatch)
    msg = _srv._gate_check("delete-module shu009 -y", wait=0.5)
    assert msg is not None
    assert "delete-module" in msg and "実行しませんでした" in msg
    assert not (tmp_path / "ai_win_ask.txt").exists()   # 聞いた跡は片づける


@_needs_srv
def test_gate_check_fails_open_when_window_unreachable(tmp_path, monkeypatch):
    # 窓に聞けない異常（ask が書けない）は素通し＝方針A。ここを閉じると
    # 窓が無い経路で道具が丸ごと使えなくなる
    monkeypatch.setenv("AI_WIN_LOG", str(tmp_path / "no" / "such" / "ai_win.log"))
    assert _srv._gate_check("close a.xlsm --no-save -y", wait=0.5) is None


@_needs_srv
def test_ask_user_ignores_stale_answer(tmp_path, monkeypatch):
    # 前回の答えが残っていても拾わない（残骸の yes で素通ししたら承認の意味が消える）
    _aiwin(tmp_path, monkeypatch)
    (tmp_path / "ai_win_answer.txt").write_text("yes", encoding="utf-8")
    msg = _srv._gate_check("write-range A1 1", wait=0.5)
    assert msg is not None


@_needs_srv
def test_ask_user_shows_the_actual_command(tmp_path, monkeypatch):
    # 人が見るのは「これから実行される行そのもの」でなければ承認にならない
    log = _aiwin(tmp_path, monkeypatch)
    seen = {}
    _answer_when_asked(log, "no", capture=seen)
    _srv._gate_check("  write-range A1 42  ", wait=5)
    assert seen.get("body") == "write-range A1 42"


# ---- _diff_confirm（当てる前に差分を見せる） ----

def _prepare_diff(tmp_path, monkeypatch, before, after):
    """_last_proc.vba と `get --out` の現物を差し替えて差分確認を撃てる形にする。"""
    import vbam_core
    last = tmp_path / "_last_proc.vba"
    last.write_text(after, encoding="utf-8")
    monkeypatch.setattr(vbam_core, "LAST_PROC_FILE", str(last))
    monkeypatch.setattr(_tf, "gettempdir", lambda: str(tmp_path))
    seen_cmd = []

    def _fake_run(cmd, *a, **k):
        seen_cmd.append(cmd)
        m = _re.search(r'--out "([^"]+)"', cmd)
        with open(m.group(1), "w", encoding="utf-8") as f:
            f.write(before)
        return True, "", ""

    monkeypatch.setattr(_srv, "_run", _fake_run)
    return seen_cmd


@_needs_srv
def test_diff_confirm_is_noop_without_env(monkeypatch):
    monkeypatch.delenv("AI_WIN_LOG", raising=False)
    assert _srv._diff_confirm("replace-procedure -y") is None


@_needs_srv
def test_diff_confirm_shows_diff_and_blocks_on_reject(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    _prepare_diff(tmp_path, monkeypatch,
                  before="Sub 合計を出す()\n    a = 1\nEnd Sub\n",
                  after="Sub 合計を出す()\n    a = 2\nEnd Sub\n")
    seen = {}
    _answer_when_asked(log, "no", capture=seen)
    msg = _srv._diff_confirm("replace-procedure --module shu003 -y")
    assert msg is not None and "replace-procedure" in msg
    body = seen.get("body", "")
    assert "合計を出す を書き換えます" in body      # 日本語のプロシージャ名を拾えている
    assert "-    a = 1" in body and "+    a = 2" in body


@_needs_srv
def test_diff_confirm_skips_when_nothing_changes(tmp_path, monkeypatch):
    # 中身が同じなら聞かない（本体側が置換をスキップする）
    _aiwin(tmp_path, monkeypatch)
    same = "Sub A()\n    x = 1\nEnd Sub\n"
    _prepare_diff(tmp_path, monkeypatch, before=same, after=same + "\n\n")
    assert _srv._diff_confirm("replace-procedure -y") is None


@_needs_srv
def test_diff_confirm_passes_module_to_get(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    seen_cmd = _prepare_diff(tmp_path, monkeypatch,
                             before="Sub A()\n    x = 1\nEnd Sub\n",
                             after="Sub A()\n    x = 2\nEnd Sub\n")
    _answer_when_asked(log, "yes")
    assert _srv._diff_confirm("replace-procedure --module shu003 -y") is None
    assert seen_cmd and seen_cmd[0].startswith('get shu003 A --out "')


@_needs_srv
def test_diff_confirm_stays_quiet_without_material(tmp_path, monkeypatch):
    # 置換予定のコードが読めない／プロシージャ名が拾えないときは黙って素通し
    import vbam_core
    _aiwin(tmp_path, monkeypatch)
    monkeypatch.setattr(vbam_core, "LAST_PROC_FILE", str(tmp_path / "無い.vba"))
    assert _srv._diff_confirm("replace-procedure -y") is None
    empty = tmp_path / "_last_proc.vba"
    empty.write_text("' コメントだけ\n", encoding="utf-8")
    monkeypatch.setattr(vbam_core, "LAST_PROC_FILE", str(empty))
    assert _srv._diff_confirm("replace-procedure -y") is None


# ---- _submit（関所を通す順番） ----

@_needs_srv
def test_submit_denial_prevents_execution(tmp_path, monkeypatch):
    log = _aiwin(tmp_path, monkeypatch)
    ran = []
    monkeypatch.setattr(_srv, "_run",
                        lambda line, timeout=600: (ran.append(line), (True, "", ""))[1])
    _answer_when_asked(log, "no")
    out = _srv._submit("write-range A1 999")
    assert "却下" in out
    assert ran == []                                   # 却下＝一行も実行しない
    assert "write-range A1 999" in log.read_text(encoding="utf-8")  # 記録は残す


@_needs_srv
def test_submit_runs_when_not_gated(monkeypatch):
    monkeypatch.delenv("AI_WIN_LOG", raising=False)
    ran = []
    monkeypatch.setattr(_srv, "_run",
                        lambda line, timeout=600: (ran.append(line), (True, "ok", ""))[1])
    out = _srv._submit("list-modules")
    assert ran == ["list-modules"]
    assert "ok" in out


# ================================================================
# xlflow から輸入した診断4本と新コマンド5本（2026-08-28）
#   VBM008 Application の状態を戻さない / VBM009 呼び先が状態を変えたまま返る
#   VBM010 モジュール変数の書き手が散る / VBM011 Find/Replace の引数省略
# ================================================================

import vbam_vba as vv


_DIAG_SRC = '''
Public gLog As String
Private gCount As Long
Public Const MAX_N = 10

Sub 戻し忘れ()
    Application.ScreenUpdating = False
    gLog = "書いた"
End Sub

Sub 戻している()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub 後始末だけ()
    Application.ScreenUpdating = True
End Sub

Sub 呼ぶ側()
    Call 戻し忘れ
    gLog = "呼んだ"
End Sub

Sub 計算モード()
    Application.Calculation = xlCalculationManual
End Sub

Sub 検索_省略()
    Set c = Range("A:A").Find(What:="あ")
End Sub

Sub 検索_明示()
    Set c = Range("A:A").Find(What:="あ", LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, MatchCase:=False)
End Sub

Sub 文字列の見せかけ()
    s = "Application.ScreenUpdating = False という説明"
    gCount = 1
End Sub
'''


def _diag_lines():
    return _DIAG_SRC.replace('\r\n', '\n').split('\n')


def test_clean_vba_line_survives_apostrophe_in_string():
    # 先に ' で切ると "Don't" で行が壊れる（check 本体が踏んだ罠の回帰）
    assert vv._clean_vba_line("""MsgBox "Don't": End Sub""") == '''MsgBox "": End Sub'''


def test_logical_line_folds_continuations():
    lines = ['Foo A, _', '    B, _', '    C', 'Next']
    stmt, used = vv._logical_line(lines, 0)
    assert used == 3
    assert 'A,' in stmt and 'B,' in stmt and 'C' in stmt


def test_split_procedures_finds_every_proc():
    procs = vv._split_procedures(_diag_lines())
    names = [p["name"] for p in procs]
    assert names == ['戻し忘れ', '戻している', '後始末だけ', '呼ぶ側', '計算モード',
                     '検索_省略', '検索_明示', '文字列の見せかけ']


def test_split_procedures_handles_property_and_oneliner():
    lines = ['Property Get X() As Long', '    X = 1', 'End Property',
             'Sub Y(): End Sub']
    procs = vv._split_procedures(lines)
    assert [(p["name"], p["kind"]) for p in procs] == [('X', 'property'), ('Y', 'sub')]


# ---- VBM008 ----

def test_vbm008_flags_unrestored_state():
    lines = _diag_lines()
    leaks = vv._diag_app_state(lines, vv._split_procedures(lines))
    assert set(leaks) == {'戻し忘れ', '計算モード'}
    assert leaks['戻し忘れ'][0][0] == 'ScreenUpdating'
    assert leaks['計算モード'][0][0] == 'Calculation'


def test_vbm008_silent_when_restored():
    lines = _diag_lines()
    leaks = vv._diag_app_state(lines, vv._split_procedures(lines))
    # 変えて戻している / 戻す側だけ書いた後始末 / 文字列の中 は鳴らない
    assert '戻している' not in leaks
    assert '後始末だけ' not in leaks
    assert '文字列の見せかけ' not in leaks


# ---- VBM009 ----

def test_vbm009_reports_caller_of_leaking_helper():
    lines = _diag_lines()
    procs = vv._split_procedures(lines)
    leaks = vv._diag_app_state(lines, procs)
    book = {("M1", p["name"]): lines[p["start"]:p["end"] + 1] for p in procs}
    lbp = {("M1", n): pr for n, pr in leaks.items()}
    hits = vv._diag_state_callers(book, lbp)
    assert [(h["caller"], h["callee"]) for h in hits] == [('呼ぶ側', '戻し忘れ')]
    assert hits[0]["props"] == ['Application.ScreenUpdating']


def test_vbm009_silent_without_leaks():
    assert vv._diag_state_callers({("M1", "A"): ["Call B"]}, {}) == []


# ---- VBM010 ----

def test_vbm010_flags_scattered_writers_only():
    lines = _diag_lines()
    found = vv._diag_module_state(lines, vv._split_procedures(lines))
    names = {v for v, _ln, _w in found}
    assert names == {'gLog'}                 # gCount は書き手1か所、MAX_N は Const
    writers = [w for v, _ln, w in found if v == 'gLog'][0]
    assert set(writers) == {'戻し忘れ', '呼ぶ側'}


# ---- VBM011 ----

def test_vbm011_flags_find_without_lookat():
    lines = _diag_lines()
    found = vv._diag_stateful_find(lines, vv._split_procedures(lines))
    assert [f[0] for f in found] == ['検索_省略']
    assert found[0][1] == 'Find'
    assert found[0][3] == ['LookAt', 'SearchOrder', 'MatchCase']


def test_vbm011_silent_when_args_span_continuation():
    # 継続行を畳まないと「LookAt が無い」と誤報する（畳めていることの回帰）
    lines = _diag_lines()
    found = vv._diag_stateful_find(lines, vv._split_procedures(lines))
    assert '検索_明示' not in [f[0] for f in found]


# ---- inspect-gui ----

def test_scan_gui_boundaries_catches_blocking_calls():
    lines = ['Sub A()', '    MsgBox "x"', '    v = InputBox("y")',
             '    Application.FileDialog(msoFileDialogFilePicker).Show',
             '    f = Application.GetOpenFilename()', 'End Sub']
    hits = vv._scan_gui_boundaries(lines, vv._split_procedures(lines))
    kinds = [h["kind"] for h in hits]
    assert 'MsgBox' in kinds and 'InputBox' in kinds
    assert 'Application.FileDialog' in kinds and 'GetOpenFilename' in kinds


def test_scan_gui_boundaries_ignores_showalldata_and_modeless():
    lines = ['Sub A()', '    ws.ShowAllData', '    UserForm1.Show vbModeless',
             '    s = "MsgBox は文字列の中"', 'End Sub']
    assert vv._scan_gui_boundaries(lines, vv._split_procedures(lines)) == []


def test_scan_gui_boundaries_catches_modal_show():
    lines = ['Sub A()', '    UserForm1.Show', 'End Sub']
    hits = vv._scan_gui_boundaries(lines, vv._split_procedures(lines))
    assert [h["kind"] for h in hits] == ['モーダルの .Show']


def test_reachable_procs_follows_calls_only():
    bodies = {'親': ['Sub 親()', '    Call 子', 'End Sub'],
              '子': ['Sub 子()', '    孫', 'End Sub'],
              '孫': ['Sub 孫()', 'End Sub'],
              '他人': ['Sub 他人()', 'End Sub']}
    assert vv._reachable_procs(bodies, '親') == {'親', '子', '孫'}
    assert vv._reachable_procs(bodies, '孫') == {'孫'}
    assert vv._reachable_procs(bodies, '居ない') == set()


def test_reachable_procs_ignores_member_access():
    # .子 はメンバー参照であって呼び出しではない
    bodies = {'親': ['Sub 親()', '    obj.子 = 1', 'End Sub'],
              '子': ['Sub 子()', 'End Sub']}
    assert vv._reachable_procs(bodies, '親') == {'親'}


# ---- materials / tidy / --show（2026-09-02 Excelコンボから輸入した「手」）----
def test_parser_accepts_materials_and_tidy():
    p = vm.build_parser()
    ns = p.parse_args(["materials", "Sheet1", "--rows", "3"])
    assert ns.command == "materials" and ns.posargs == ["Sheet1"] and ns.rows == "3"
    ns = p.parse_args(["tidy", "A5:G13", "--header-from", "A5", "--no-border", "--max-width", "40"])
    assert ns.command == "tidy" and ns.header_from == "A5" and ns.no_border and ns.max_width == "40"
    table = vm._command_table()
    assert vm.raw_command(table["materials"]) is vm.cmd_materials     # 表の実装は台帳の包みつき（2026-09-16）
    assert vm.raw_command(table["tidy"]) is vm.cmd_tidy


def test_dirt_notes_counts_what_the_grid_cannot_show():
    """materials の「気づき」: 空白の幅・半角カナ・全角数字・重複・文字の日付・空欄を道具が数える（2026-09-07）。

    格子の表示では見えないものを頭で埋めようとして read-range と --raw を 2 往復挟み、
    書き始めるまで 100 秒かかった。道具が数えれば、書く前に調べる理由が消える。
    """
    import datetime as _dt
    grid = [["テスト用データ2", None, None, None, None],
            [None] * 5,
            ["例：指示", None, None, None, None],
            [None] * 5,
            ["会員番号", "氏名", "フリガナ", "種別", "入会日"],
            [1001, "佐藤　一郎", "ｻﾄｳ ｲﾁﾛｳ", "正会員", _dt.datetime(2024, 4, 1)],
            [1002, "鈴木 花子", "スズキ　ハナコ", "正会員", "2024/5/12"],
            [1003, "田中　　次郎", "タナカ ジロウ", "準会員", _dt.datetime(2024, 6, 3)],
            [1003, "田中　　次郎", "タナカ ジロウ", "準会員", _dt.datetime(2024, 6, 3)],
            ["１００５", " 伊藤　さくら", None, "準会員", _dt.datetime(2025, 1, 15)],
            [1006, "渡辺　健", "ワタナベ ケン", "12", _dt.datetime(2025, 2, 8)],
            [1001, "佐藤　太郎", "サトウ タロウ", "学生", _dt.datetime(2025, 3, 8)]]
    text = "\n".join(vm.dirt_notes(grid, r0=1, c0=1, col_formats={4: {"yyyy/mm/dd", "yyyy/m/d"}}))
    assert "重複行: 行9 = 行8（1003）" in text
    assert "番号列の重複: A12 = A6（1001）" in text          # 行ごと同じ行（9）は番号列の重複には数えない
    assert "全角数字: A10" in text
    assert "半角カナ: C6" in text
    assert "端に空白 B10" in text and "連続 B8" in text
    assert "B列は全角と半角の区切りが混在（半角 B7）" in text
    assert "C列は全角と半角の区切りが混在（全角 C7）" in text
    assert "文字の数字（番号列は除く）: D11" in text          # 番号列 A の 1001 等は数えない
    assert "文字の日付: E7" in text
    assert "日付の表示形式が混在: E列（yyyy/m/d, yyyy/mm/dd）" in text
    assert "表の中の空欄: C10" in text
    # きれいな表は 1 行で「なし」
    clean = [["会員番号", "氏名"], [1, "佐藤　一郎"], [2, "鈴木　花子"], [3, "田中　次郎"]]
    assert vm.dirt_notes(clean) == ["気づき（表の汚れ）: なし（重複・全角数字・半角カナ・空白の乱れ・"
                                   "文字の数字/日付・空欄・文字として入った式、いずれも無し）"]
    # 文字として入った式（表示形式が @ の列に式を入れた形。Excel は計算しない・2026-09-09）
    astext = [["品目", "金額"], ["a", "=B2*2"], ["b", 200], ["c", 300]]
    assert any("式が文字として入っているセル（計算されていません）: B2" in n for n in vm.dirt_notes(astext))
    assert vm.dirt_notes([]) == []
    src = inspect.getsource(vm.cmd_materials)
    assert "dirt_notes(" in src and "_date_col_formats(" in src
    assert "formula_notes(" in src        # 数式の気づきも materials が出す（2026-09-09）


def test_parser_show_flag_on_write_commands():
    p = vm.build_parser()
    for argv in (["write-range", "A1", "x", "--show"], ["fill", "A1:A5", "--show"],
                 ["copy-range", "A1", "B1", "--show"], ["format-range", "A1:B2", "--bold", "--show"]):
        ns, unknown = p.parse_known_args(argv)
        assert ns.show is True and not unknown, argv


# ---- write-cells / tidy 複数範囲 / 仕事の時計（2026-09-02 夜・往復数を減らす改修）----
def test_parser_accepts_write_cells_and_multi_range_tidy():
    p = vm.build_parser()
    ns, unknown = p.parse_known_args(["write-cells", "C7", "000-0002", "D7", "x y", "--show"])
    assert ns.command == "write-cells" and not unknown
    assert ns.posargs == ["C7", "000-0002", "D7", "x y"] and ns.show is True
    ns = p.parse_args(["write-cells", "--tsv", "cells.tsv", "--raw", "--sheet", "S"])
    assert ns.tsv_opt == "cells.tsv" and ns.raw and ns.sheet_opt == "S" and ns.posargs == []
    ns = p.parse_args(["tidy", "A5:G13", "I5:L11"])
    assert ns.posargs == ["A5:G13", "I5:L11"] and ns.header_from is None
    table = vm._command_table()
    assert vm.raw_command(table["write-cells"]) is vm.cmd_write_cells


# ---- tidy の列の型（2026-09-02 深夜・「会員番号は左寄せ・金額はカンマが当然。道具でそうなるようにしろ」）----
import vbam_edit as ve                                              # noqa: E402


def test_column_style_id_number_skip():
    cs = ve._column_style
    assert cs("会員番号", [2001.0, 2002.0]) == 'id'
    assert cs("郵便番号", ["000-0001", "000-0002"]) == 'id'
    assert cs("ＩＤ", [1.0, 2.0]) == 'id'                    # 全角も NFKC で拾う
    assert cs("No.", [1.0]) == 'id' and cs("商品コード", [1.0]) == 'id'
    assert cs("備考", [1.0, 2.0]) == 'int'                    # No/ID の誤検知はしない
    assert cs("金額", [12000.0, "", None, 7000.0]) == 'int'   # IFERROR の "" と空欄は無視
    assert cs("月額", [1200.0]) == 'int'                      # 「月」で終わらない＝金額
    assert cs("単価", [12.5, 3.0]) == 'dec'
    assert cs("年度", [2024.0, 2025.0]) is None               # 2,024 にしない
    assert cs("入会年", [2024.0]) is None and cs("日", [1.0]) is None
    assert cs("達成率", [0.5]) is None and cs("割合(%)", [50.0]) is None
    assert cs("申込コース", ["年間", "半年"]) is None          # 文字列の列は触らない
    assert cs("金額", [1.0, "x"]) is None                     # 混在も触らない
    assert cs("金額", ["", None]) is None and cs(None, []) is None
    assert cs("フラグ", [True, False]) is None


def test_parser_accepts_no_col_format():
    p = vm.build_parser()
    ns = p.parse_args(["tidy", "A5:G13", "--no-col-format"])
    assert ns.no_col_format is True
    ns = p.parse_args(["tidy", "A5"])
    assert ns.no_col_format is False


def test_write_cells_rejects_unpaired_args(capsys):
    ns = _ap.Namespace(posargs=["C7"], tsv_opt=None, raw=False, sheet_opt=None, show=False)
    assert vm.cmd_write_cells(ns) is False
    out = capsys.readouterr().out
    assert "使い方" in out and "組になっていません" in out
    ns = _ap.Namespace(posargs=["C7", "v"], tsv_opt="x.tsv", raw=False, sheet_opt=None, show=False)
    assert vm.cmd_write_cells(ns) is False
    assert "同時に指定できません" in capsys.readouterr().out


def test_job_clock_roundtrip(monkeypatch, tmp_path):
    import vbam_core as vc
    monkeypatch.setattr(vc, "_JOB_CLOCK_FILE", str(tmp_path / "clock.json"))
    assert vc.job_clock_elapsed() is None and vc.job_clock_note() == ""
    vc.job_clock_start("book!sheet")
    sec = vc.job_clock_elapsed()
    assert sec is not None and 0 <= sec < 5
    assert vc.job_clock_note().startswith("経過: materials から")
    # 1時間より古い時計は無視（前の仕事の時計を今の仕事に足さない）
    assert vc.job_clock_elapsed(max_age=0) is None


# ---- Excel未起動の早期検知（2026-09-03・agent --fire が止まる穴を塞ぐ）----

class _FakeCompleted:
    def __init__(self, stdout):
        self.stdout = stdout


def test_any_excel_process_false_when_tasklist_has_no_excel(monkeypatch):
    import vbam_core as vc
    # 2026-09-16 の PID の控え: 前のテストが本物の Excel の PID を控えていると tasklist を撃たずに True になる
    # （本物の Excel が開いているときだけ落ちる順序依存）。控えを空にしてから tasklist の道を試す
    monkeypatch.setattr(vc, "_last_excel_pid", None)
    monkeypatch.setattr(
        "subprocess.run",
        lambda *a, **k: _FakeCompleted("情報: 条件に一致する実行中のタスクはありません。"))
    assert vc._any_excel_process() is False


def test_any_excel_process_true_when_tasklist_has_excel(monkeypatch):
    import vbam_core as vc
    monkeypatch.setattr(
        "subprocess.run",
        lambda *a, **k: _FakeCompleted('"EXCEL.EXE","1234","Console","1","123,456 K"'))
    assert vc._any_excel_process() is True


def test_any_excel_process_true_when_tasklist_itself_fails(monkeypatch):
    # tasklist が失敗＝「いるかどうか分からない」なので安全側（True）に倒す
    import vbam_core as vc

    def _boom(*a, **k):
        raise OSError("tasklist 不在")
    monkeypatch.setattr("subprocess.run", _boom)
    assert vc._any_excel_process() is True


def test_get_workbook_uncached_fails_fast_without_com_when_excel_absent(monkeypatch):
    # Excel が1つも無いと分かったら、ROT走査やGetActiveObject等のCOM呼び出しに
    # 進む前に打ち切る（進んだら止まりうる＝2026-09-03 の穴）。COM側の関数が
    # 呼ばれたらテスト自体を失敗させて、打ち切りが本当に先に効くことを確かめる
    import vbam_core as vc
    monkeypatch.setattr(vc, "_any_excel_process", lambda: False)

    def _must_not_be_called(*a, **k):
        raise AssertionError("Excel未起動と分かった後もCOM呼び出しに進んでいる")
    monkeypatch.setattr(vc, "_running_excel_workbooks", _must_not_be_called)
    monkeypatch.setattr(vc, "_get_active_excel", _must_not_be_called)
    try:
        vc._get_workbook_uncached(None)
        assert False, "例外が発生するはず"
    except Exception as ex:
        assert "起動していません" in str(ex)


def test_open_maybe_repair_retries_with_corruptload_on_failure(capsys):
    """9-1: Google スプレッドシート書き出しの xlsx 等、素の Open が COM から弾かれるファイルがある
    （中身は正常＝openpyxl は読める）。CorruptLoad=xlRepairFile で開き直すと通る（2026-09-05・実射で発見）。
    """
    import vbam_core as vc
    calls = []

    class FakeWorkbooks:
        def Open(self, *args):
            calls.append(args)
            if len(args) <= 3:                 # 素の呼び出し（target_path, 0, True 等）は失敗させる
                raise Exception("COM から弾かれた")
            return "wb-repaired"

    class FakeXl:
        Workbooks = FakeWorkbooks()

    wb = vc._open_maybe_repair(FakeXl(), r"C:\x\本物.xlsx", (0, True))
    assert wb == "wb-repaired"
    assert len(calls) == 2
    assert calls[0] == (r"C:\x\本物.xlsx", 0, True)          # 1 回目は素の呼び出しのまま（既存動作を変えない）
    assert len(calls[1]) == 15 and calls[1][-1] == 1          # 2 回目は 15 引数・CorruptLoad=1
    assert calls[1][:3] == (r"C:\x\本物.xlsx", 0, True)       # 元の引数はそのまま先頭に残る
    assert "修復モード" in capsys.readouterr().out

    # 素の呼び出しが最初から通るときは 1 回だけ・そのまま返す（引数無しの呼び出し形も確認）
    calls.clear()

    class FakeWorkbooksOk:
        def Open(self, *args):
            calls.append(args)
            return "wb-ok"
    FakeXl.Workbooks = FakeWorkbooksOk()
    wb = vc._open_maybe_repair(FakeXl(), r"C:\x\本物.xlsx")
    assert wb == "wb-ok" and calls == [(r"C:\x\本物.xlsx",)]


# ---- build-sheet（2026-09-03・設計図からシートを組み上げる「作る犬」の手1）----

def test_build_sheet_is_wired():
    # 表と argparse の両方に配線（片肺を防ぐ）
    import vbam_build as vb
    table = vm._command_table()
    assert vm.raw_command(table["build-sheet"]) is vb.cmd_build_sheet
    p = vm.build_parser()
    ns, unknown = p.parse_known_args(["build-sheet", "spec.json", "--overwrite", "--dry-run"])
    assert ns.command == "build-sheet" and ns.overwrite and ns.dry_run and not unknown
    ns = p.parse_args(["build-sheet", "--sample"])
    assert ns.sample and ns.posargs == []


def test_plan_sheet_positions_and_formulas():
    import vbam_build as vb
    p = vb.plan_sheet(vb.SAMPLE_SPEC)
    # 題名があるので見出しは 3 行目、データは 4〜23 行、合計行は 24
    assert p['header_row'] == 3 and p['first'] == 4 and p['last'] == 23 and p['total_row'] == 24
    assert p['header_addr'] == 'A3:H3' and p['full_addr'] == 'A3:H24'
    letters = {c['name']: c['letter'] for c in p['columns']}
    assert letters['数量'] == 'E' and letters['単価'] == 'F' and letters['金額'] == 'G'
    # {列名} は同じ行の実番地へ
    assert p['grid'][0][6] == '=IF(E4="","",E4*F4)' and p['grid'][19][6] == '=IF(E23="","",E23*F23)'
    # 合計行は SUM（データ行の範囲）
    assert ('E', '=SUM(E4:E23)') in p['total_cells'] and ('G', '=SUM(G4:G23)') in p['total_cells']
    # 型 → 表示形式、選択肢、名前定義は見出し込みの表
    fmt = {c['name']: c['format'] for c in p['columns']}
    assert fmt['会員番号'] == '@' and fmt['金額'] == '#,##0' and fmt['入会日'] == 'yyyy/m/d'
    assert p['validations'] == [('D4:D23', ['正会員', '準会員', '賛助'])]
    assert p['names'] == [('会員名簿範囲', 'A3:H23')]
    assert p['cond_format'][0]['addr'] == 'G4:G23' and p['cond_format'][0]['rule'] == ('gt', 100000)
    assert p['print']['area'] == 'A1:H24' and p['print']['title_rows'] == '3:3'


def test_api_call_with_retry_only_on_busy_codes(monkeypatch):
    # 429・5xx だけ待って撃ち直す（2 回まで）。401 や 400 は撃ち直さない（同じ返事が返るだけ）
    import io
    import urllib.error
    import vbam_build as vb
    waits = []
    monkeypatch.setattr(vb.time, "sleep", lambda s: waits.append(s))

    def http_error(code):
        return urllib.error.HTTPError("http://x", code, "err", {}, io.BytesIO(b""))

    calls = {'n': 0}

    def flaky():
        calls['n'] += 1
        if calls['n'] < 3:
            raise http_error(503)
        return "ok"

    assert vb._api_call_with_retry(flaky, "gemini") == "ok"
    assert calls['n'] == 3 and waits == list(vb._API_RETRY_WAITS)          # 2 回待って 3 回目で通る
    waits.clear()
    calls['n'] = 0

    def always_busy():
        calls['n'] += 1
        raise http_error(429)

    import pytest
    with pytest.raises(urllib.error.HTTPError):
        vb._api_call_with_retry(always_busy, "gemini")
    assert calls['n'] == len(vb._API_RETRY_WAITS) + 1                       # 撃ち直しは 2 回まで
    waits.clear()
    calls['n'] = 0

    def bad_key():
        calls['n'] += 1
        raise http_error(401)

    with pytest.raises(urllib.error.HTTPError):
        vb._api_call_with_retry(bad_key, "gemini")
    assert calls['n'] == 1 and waits == []                                   # 401 は即エラー・待たない


def test_build_book_materials_show_headers():
    # 2026-09-04: 手元の様子はシート名だけでなく、使用範囲と見出しまで（既存の表を参照する数式を組めるように）
    import vbam_build as vb

    class FakeRange:
        def __init__(self, addr, rows, cols, first):
            self.Address, self._first = addr, first
            self.Rows = type("R", (), {"Count": rows, "__call__": lambda s, i: type("X", (), {"Value": first})})()
            self.Columns = type("C", (), {"Count": cols})()

        def Cells(self, r, c):
            return type("X", (), {"Value": None})()

    class FakeSheet:
        def __init__(self, name, ur):
            self.Name, self.UsedRange = name, ur

    class FakeWB:
        Sheets = [FakeSheet("目次", FakeRange("$A$1:$A$1", 1, 1, None)),
                  FakeSheet("売上", FakeRange("$A$1:$D$46", 46, 4, (("日付", "地域", "商品", "売上"),)))]
    text = vb._book_materials(FakeWB())
    assert "既存のシート: 目次, 売上" in text
    assert "目次: （空）" in text
    assert "売上: A1:D46（46 行 × 4 列）" in text and "見出し: 日付 | 地域 | 商品 | 売上" in text
    assert vb._book_materials(type("E", (), {"Sheets": []})()) is None


def test_recipes_all_plan_and_are_wired():
    # 手5: 手順書5本は全部 plan_sheet を通る（Excel 不要）。請求明細は合計行の下に消費税・税込合計が続く
    import vbam_build as vb
    assert set(vb.RECIPES) == {"名簿", "備品台帳", "月次集計表", "請求明細", "出勤簿"}
    for n, r in vb.RECIPES.items():
        assert r.get("_use"), n
        p = vb.plan_sheet(r)
        assert p['sheet'] == r['sheet'] and p['columns'], n
    p = vb.plan_sheet(vb.RECIPES["請求明細"])
    assert p['total_row'] == 14 and p['end_row'] == 16 and p['full_addr'] == 'A3:G16'
    rows = {r: v for L, v, r in p['extra_cells'] if str(v).startswith('=')}
    assert rows[15] == '=ROUNDDOWN(F14*0.1,0)' and rows[16] == '=F14+F15'
    # 月次集計表: 年計は同じ行の 4月〜3月（B〜M）を足す
    p = vb.plan_sheet(vb.RECIPES["月次集計表"])
    assert p['grid'][0][13] == '=IF(COUNT(B4:M4)=0,"",SUM(B4:M4))'
    # 出勤簿: 曜日は数式なので '@' を当てない（数式が文字として入る事故の予防）
    p = vb.plan_sheet(vb.RECIPES["出勤簿"])
    yobi = [c for c in p['columns'] if c['name'] == '曜日'][0]
    assert yobi['formula'] and yobi['format'] is None
    # 配線
    ns, unknown = vm.build_parser().parse_known_args(["build-sheet", "--fire", "名簿", "請求明細"])
    assert ns.fire and ns.posargs == ["名簿", "請求明細"] and not unknown
    ns = vm.build_parser().parse_args(["build-sheet", "--recipe", "出勤簿", "--new-book"])
    assert ns.recipe == "出勤簿" and ns.new_book


def test_plan_sheet_rejects_bad_spec():
    import vbam_build as vb
    import pytest
    with pytest.raises(ValueError):
        vb.plan_sheet({"columns": [{"name": "a"}]})                     # sheet なし
    with pytest.raises(ValueError):
        vb.plan_sheet({"sheet": "s"})                                    # columns なし
    with pytest.raises(ValueError):
        vb.plan_sheet({"sheet": "s", "columns": [{"name": "a"}, {"name": "a"}]})   # 列名重複
    with pytest.raises(ValueError):
        vb.plan_sheet({"sheet": "s", "columns": [{"name": "a", "formula": "={b}*2"}]})  # 未知の列名
    with pytest.raises(ValueError):
        vb.plan_sheet({"sheet": "s", "columns": [{"name": "a", "type": "money"}]})     # 不明な型
    # 題名なし・データ付きなら見出し 1 行目、行数はデータの行数
    p = vb.plan_sheet({"sheet": "s", "columns": ["a", "b"], "data": [[1, 2], [3, 4], [5, 6]]})
    assert p['header_row'] == 1 and p['nrows'] == 3 and p['full_addr'] == 'A1:B4' and p['grid'][2] == [5, 6]


def test_write_helpers_exposed_by_recipe_fire():
    # 2026-09-04 の手順書の実射で出た 3 つの穴: 日付の 9 時間ずれ／先頭 ' で文字のまま／FILTER の暗黙の交差
    import datetime
    import vbam_core as vc
    import vbam_edit as ve
    d = vc._coerce_cell("2026/01/05")
    assert isinstance(d, datetime.datetime) and d.tzinfo is datetime.timezone.utc and (d.year, d.month, d.day) == (2026, 1, 5)
    assert vc._coerce_cell("'2026/01/05") == "2026/01/05" and vc._coerce_cell("'0001") == "0001"
    assert vc._coerce_cell("'") == "" and vc._coerce_cell("=SUM(A1)") == "=SUM(A1)"
    assert ve._is_dynamic_formula('=FILTER(A2:E13,C2:C13="正会員","該当なし")')
    assert ve._is_dynamic_formula("=xlookup(A1,B:B,C:C)") and ve._is_dynamic_formula("=LET(x,1,x)")
    assert not ve._is_dynamic_formula("=SUM(A1:A3)") and not ve._is_dynamic_formula("FILTER") and not ve._is_dynamic_formula(12)

    class _Cell:
        def __init__(self, fmt="@"):
            self.NumberFormat = fmt
            self.Value = None
            self.Formula2 = None
    c = _Cell()
    ve._untext_one(c, 12.0)
    assert c.NumberFormat == "General"
    c = _Cell()
    ve._untext_one(c, "0001")                      # 文字を書くときは @ を守る（先頭ゼロ）
    assert c.NumberFormat == "@"
    c = _Cell("General")
    ve._write_value(c, "=FILTER(A1:A3,B1:B3=1)")
    assert c.Formula2 and c.Value is None
    ve._write_value(c, "=SUM(A1)")
    assert c.Value == "=SUM(A1)"


# ---- capabilities / rules ----

_JA_ALIASES = {'健康診断', '影響範囲', '予行演習', 'テスト',
               '全体コンパイル', '関所', 'フォーム書き出し', '配線図', 'エージェント',
               'キー設定', 'キー削除'}


def test_capabilities_classifies_every_command():
    # コマンドを足して分類を忘れたらここで落ちる（分類漏れの番人）
    table = set(vm._command_table())
    listed = {n for items in vv._CAPABILITIES.values() for n, _ in items}
    assert not (listed - table), f"表に無いコマンドを載せている: {sorted(listed - table)}"
    missing = table - listed - _JA_ALIASES
    assert not missing, f"安全性の分類が無いコマンド: {sorted(missing)}"


def test_capabilities_names_are_unique():
    names = [n for items in vv._CAPABILITIES.values() for n, _ in items]
    assert len(names) == len(set(names))


def test_capabilities_destructive_set_is_explicit():
    d = {n for n, _ in vv._CAPABILITIES["destructive"]}
    assert d == {'delete-procedure', 'delete-module', 'clear-range',
                 'close', 'close-form', 'vbe-reset',
                 'clear-key',          # 2026-09-04: APIキーの削除（ブックではなく環境変数を消す）
                 'shape'}              # 2026-09-04: 図形の削除（--delete。消せるのは図形だけ）


def test_rules_codes_are_unique_and_sequential():
    codes = [r[0] for r in vv._CHECK_RULES]
    assert codes == sorted(set(codes))
    assert codes[0] == 'VBM001' and codes[-1] == 'VBM016'      # 2026-09-17: VBM012〜014、同日 VBM015・016（clean-vba を畳んだ）


def test_rules_vbm012_to_014_are_errors():
    d = {r[0]: r for r in vv._CHECK_RULES}
    for c in ('VBM012', 'VBM013', 'VBM014'):
        assert d[c][1] == 'error', c
    assert 'ラベル' in d['VBM012'][2] and 'On Err' in d['VBM013'][2] and '存在しない' in d['VBM014'][2]


def test_rules_imported_four_cite_their_origin():
    imported = {r[0]: r[4] for r in vv._CHECK_RULES if r[4]}
    assert imported == {'VBM008': 'xlflow VBA203', 'VBM009': 'xlflow VBA221',
                        'VBM010': 'xlflow VBA240', 'VBM011': 'xlflow VBA215'}


def test_xlflow_commands_are_wired():
    # 表と argparse の両方に配線されていること（片肺を防ぐ）
    table = vm._command_table()
    for name in ("inspect-gui", "capabilities", "rules", "metrics", "process"):
        assert name in table
    parser = vm.build_parser()
    ns, unknown = parser.parse_known_args(["inspect-gui", "a.xlsm", "macroA", "--json"])
    assert ns.command == "inspect-gui" and ns.json and not unknown
    ns, unknown = parser.parse_known_args(["metrics", "--top", "5"])
    assert ns.command == "metrics" and ns.top == 5 and not unknown
    ns, unknown = parser.parse_known_args(["process", "--kill", "1234"])
    assert ns.command == "process" and ns.kill == 1234 and not unknown
    ns, unknown = parser.parse_known_args(["capabilities", "run-macro"])
    assert ns.command == "capabilities" and not unknown


# ---- 文字列の化け戻し / write_grid（2026-09-02 深夜・決め直しの 129 秒を道具に焼いた改修）----
class _FakeCell:
    def __init__(self, readback, addr="$H$6"):
        self.Value = readback           # 書いた直後に Excel が返す値（読み替え後）
        self.NumberFormat = "General"
        self.Address = addr


def test_kept_as_text_restores_only_coerced_strings():
    import datetime as _dt
    import vbam_edit as ve
    # "1-2" を書いたら Excel が日付にした → 表示形式 @ で文字列に戻す
    c = _FakeCell(_dt.datetime(2026, 1, 2))
    assert ve._kept_as_text(c, "1-2") is True
    assert c.NumberFormat == "@" and c.Value == "1-2"
    # "1,000" → 1000 も同じ
    c = _FakeCell(1000.0)
    assert ve._kept_as_text(c, "1,000") is True and c.Value == "1,000"
    # 文字列のまま残った → 触らない
    c = _FakeCell("000-0002")
    assert ve._kept_as_text(c, "000-0002") is False and c.NumberFormat == "General"
    # 数式・数値（意図した変換）・真偽値・空は対象外
    assert ve._kept_as_text(_FakeCell(3), 3) is False
    assert ve._kept_as_text(_FakeCell(5), "=A1") is False
    assert ve._kept_as_text(_FakeCell(True), "TRUE") is False
    assert ve._kept_as_text(_FakeCell(None), "") is False


class _FakeWS:
    """Range(...).Value の読み戻しを 1 回で返し、Cells(r,c) を番地つきで配る"""
    def __init__(self, readback):
        self._rb = readback
        self.cells = {}
        self.range_reads = 0

    def Cells(self, r, c):
        return self.cells.setdefault((r, c), _FakeCell(None, addr=f"$R{r}C{c}"))

    def Range(self, a, b):
        self.range_reads += 1
        return _FakeCell(self._rb)


def test_keep_grid_text_fixes_only_coerced_cells_with_one_readback():
    import datetime as _dt
    import vbam_edit as ve
    grid = (("申込コース", "1-2"), ("=A1", "3/4"), ("000-0002", 7))
    # 読み戻し: "1-2" と "3/4" だけ Excel が日付にした
    rb = (("申込コース", _dt.datetime(2026, 1, 2)),
          ("=A1", _dt.datetime(2026, 3, 4)),
          ("000-0002", 7))
    ws = _FakeWS(rb)
    fixed = ve._keep_grid_text(ws, 5, 6, grid)
    assert fixed == ["R5C7", "R6C7"]
    assert ws.range_reads == 1                      # 読み戻しは範囲 1 回
    assert ws.cells[(5, 7)].Value == "1-2" and ws.cells[(5, 7)].NumberFormat == "@"
    assert ws.cells[(6, 7)].Value == "3/4" and ws.cells[(6, 7)].NumberFormat == "@"
    untouched = [k for k in ws.cells if k not in {(5, 7), (6, 7)}]
    assert all(ws.cells[k].NumberFormat == "General" and ws.cells[k].Value is None
               for k in untouched)                  # 化けていないセルには触らない
    # 1x1（読み戻しがスカラー）も同じ経路で戻る
    ws = _FakeWS(_dt.datetime(2026, 1, 2))
    assert ve._keep_grid_text(ws, 1, 1, (("1-2",),)) == ["R1C1"]
    # 全部数式／数値なら読み戻しすらしない
    ws = _FakeWS(None)
    assert ve._keep_grid_text(ws, 1, 1, (("=A1", 2),)) == [] and ws.range_reads == 0


def test_keep_grid_text_leaves_typed_columns_alone():
    """2026-09-09: build が設計図で [date] と決めた列まで文字に戻していた。

    戻すと日付が文字になり、ピボットの月別まとめ（group_date）ができない（Excel は文字の日付を
    月でまとめられない）。守るのは文字列列と型を決めなかった列だけ＝先頭ゼロは今までどおり守る。
    """
    import datetime as _dt
    import vbam_edit as ve
    grid = (("0001", "2026-01-05"), ("0002", "2026-01-07"))
    rb = (("0001", _dt.datetime(2026, 1, 5)), ("0002", _dt.datetime(2026, 1, 7)))
    # 列 1（0 起点）は [date] と決めた列＝戻さない。列 0（会員番号・文字列列）は守る
    ws = _FakeWS(rb)
    assert ve._keep_grid_text(ws, 1, 1, grid, skip_cols=(1,)) == []
    assert ws.cells == {} or all(c.Value is None for c in ws.cells.values())
    # skip_cols を渡さなければ今までどおり戻す（既存の呼び出し側は変わらない）
    ws2 = _FakeWS(rb)
    assert ve._keep_grid_text(ws2, 1, 1, grid) == ["R1C2", "R2C2"]
    assert ws2.cells[(1, 2)].Value == "2026-01-05" and ws2.cells[(1, 2)].NumberFormat == "@"


# ================================================================
# APIキーの預かり（2026-09-04）: set-key / clear-key
# ================================================================
import io                                                            # noqa: E402


@_needs_srv
def test_write_grid_writes_tsv_and_calls_write_range_show(tmp_path, monkeypatch):
    tsv_path = tmp_path / "_last_values.tsv"
    monkeypatch.setattr(_srv, "LAST_VALUES", str(tsv_path))
    seen = {}

    def _fake_submit(line, timeout=600):
        seen["line"] = line
        return "ok"
    monkeypatch.setattr(_srv, "_submit", _fake_submit)
    out = _srv.write_grid(
        "F5:G13", '申込コース\t金額\r\n=IFERROR(VLOOKUP($A6,申込一覧,2,FALSE),"申込なし")\t12000')
    assert out == "ok"
    body = tsv_path.read_text(encoding="utf-8")
    assert body == '申込コース\t金額\n=IFERROR(VLOOKUP($A6,申込一覧,2,FALSE),"申込なし")\t12000\n'
    assert seen["line"] == f'write-range "F5:G13" --tsv "{tsv_path}" --show'


@_needs_srv
def test_replace_procedure_takes_code_and_replaces_in_one_call(tmp_path, monkeypatch):
    """set_procedure_code → replace_procedure の 2 手を、code を渡せば 1 手に（2026-09-13・Gemini の提案）。
    書き込み先は vbam_core.LAST_PROC_FILE の 1 か所。CRLF で来ても改行を二重にしない。"""
    import vbam_core
    last = tmp_path / "_last_proc.vba"
    monkeypatch.setattr(vbam_core, "LAST_PROC_FILE", str(last))
    seen = []

    def _fake_submit(line, timeout=600):
        seen.append((line, last.read_text(encoding="utf-8")))   # 置換の時点でファイルに何が入っているか
        return "ok"
    monkeypatch.setattr(_srv, "_submit", _fake_submit)
    assert _srv.replace_procedure(module="Module1", code="Sub A()\r\n    x = 2\r\nEnd Sub") == "ok"
    assert seen == [('replace-procedure -y --module "Module1"', "Sub A()\n    x = 2\nEnd Sub")]
    # code を省けば今までどおり: set_procedure_code で置いた中身をそのまま当てる
    seen.clear()
    _srv.set_procedure_code("Sub B()\nEnd Sub")
    assert _srv.replace_procedure() == "ok"
    assert seen == [("replace-procedure -y", "Sub B()\nEnd Sub")]


def test_mcp_agent_escapes_a_request_that_starts_with_a_dash(tmp_path, monkeypatch):
    """'-5%で計算して' が argparse にオプションと読まれていた（shlex.quote では守れない）。"""
    import vba_mcp_server as ms
    seen = {}

    def fake_submit(line, timeout=600):
        seen['line'] = line
        return "ok"

    monkeypatch.setattr(ms, "_submit", fake_submit)
    ms.agent.fn(request="-5%で計算して") if hasattr(ms.agent, 'fn') else ms.agent(request="-5%で計算して")
    line = seen['line']
    assert "--request-file" in line and "-5%" not in line
    path = line.split("--request-file ")[1].strip().strip("'\"")
    assert open(path, encoding="utf-8").read() == "-5%で計算して"
    os.remove(path)


def test_mcp_agent_timeout_follows_the_number_of_turns():
    """12: 900 秒固定で、手順書の 6 往復が超えると報告ごと消えていた。"""
    import vba_mcp_server as ms
    seen = {}

    def fake_submit(line, timeout=600):
        seen['timeout'] = timeout
        return "ok"

    old = ms._submit
    ms._submit = fake_submit
    try:
        fn = ms.agent.fn if hasattr(ms.agent, 'fn') else ms.agent
        fn(request="直して")
        assert seen['timeout'] == 300 + 4 * 300
        fn(request="直して", max_turns=8)
        assert seen['timeout'] == 300 + 8 * 300
        fn(request="直して", max_turns=100)
        assert seen['timeout'] == 3600            # 上限で頭打ち
    finally:
        ms._submit = old


# ================================================================
# 仕上げ検査・自己採点・点数の履歴・曖昧な依頼の弾（2026-09-04）
# ================================================================
class _AuditRange:
    """audit_table 用の作り物のセル範囲（COM の形だけ真似る）。"""

    def __init__(self, rows, header_bold=True, borders='あり', fmts=None, aligns=None,
                 texts=None, header_fill=True):
        self._rows = rows
        self._header_bold = header_bold
        self._borders = borders
        self._fmts = fmts or {}
        self._aligns = aligns or {}
        self._texts = texts or {}
        self._header_fill = header_fill

    Row = 1
    Column = 1

    @property
    def Rows(self):
        return type("R", (), {"Count": len(self._rows)})()

    @property
    def Columns(self):
        # Count（範囲の列数）と Columns(j).ColumnWidth（_hash_scan が列幅を見る）の両方を真似る。
        # 幅は 0＝数値のセルは全部「怪しい」扱いで .Text を読みに行く（texts の ### が効く）
        n = len(self._rows[0])
        return type("C", (), {"Count": n, "__call__": lambda s, j: type("W", (), {"ColumnWidth": 0.0})()})()

    @property
    def Value(self):
        return tuple(tuple(r) for r in self._rows)             # COM と同じ行タプルの組

    @property
    def Address(self):
        return "$A$1:$" + chr(64 + len(self._rows[0])) + "$" + str(len(self._rows))

    def Cells(self, i, j):
        outer = self

        class _C:
            Value = outer._rows[i - 1][j - 1]
            Address = "$" + chr(64 + j) + "$" + str(i)
            Text = outer._texts.get((i, j), str(outer._rows[i - 1][j - 1]))

            class Font:
                Bold = outer._header_bold if i == 1 else False

            class Interior:
                ColorIndex = (1 if (i == 1 and outer._header_fill) else -4142)
        return _C()

    def Borders(self, idx):
        st = self._borders
        return type("B", (), {"LineStyle": (1 if st == 'あり' else -4142 if st == 'なし' else None)})()


class _AuditWS:
    def __init__(self, rng):
        self._rng = rng

    def Range(self, a, b=None):
        outer = self._rng
        j = ord(str(a.Address).replace('$', '')[0]) - 64
        col = [r[j - 1] for r in outer._rows[1:]]

        class _Col:
            Value = [(v,) for v in col] if len(col) != 1 else col[0]   # COM は行タプルで返す
            NumberFormat = outer._fmts.get(j, 'G/標準')
            HorizontalAlignment = outer._aligns.get(j, 1)
        return _Col()


def test_audit_finds_the_look_problems_and_passes_a_tidy_table():
    """仕上げ検査: 見出しの体裁・罫線の継ぎはぎ・桁区切り・番号列の寄せ・### を機械で見つける。"""
    from vbam_edit import audit_table
    rows = [["会員番号", "氏名", "金額"], [1, "青木", 1250000], [2, "石川", 980000]]
    good = _AuditRange(rows, header_bold=True, borders='あり',
                      fmts={3: '#,##0'}, aligns={1: -4131})
    assert audit_table(_AuditWS(good), good) == []
    bad = _AuditRange(rows, header_bold=False, header_fill=False, borders='なし')
    notes = audit_table(_AuditWS(bad), bad)
    joined = " / ".join(notes)
    assert "見出し行が本文と同じ体裁" in joined
    assert "罫線がありません" in joined
    assert "桁区切りがありません" in joined
    assert "番号列" in joined and "右に寄っています" in joined
    patch = _AuditRange(rows, borders='まちまち', fmts={3: '#,##0'}, aligns={1: -4131})
    assert any("継ぎはぎ" in n for n in audit_table(_AuditWS(patch), patch))
    hashed = _AuditRange(rows, fmts={3: '#,##0'}, aligns={1: -4131}, texts={(2, 3): "######"})
    assert any("###" in n for n in audit_table(_AuditWS(hashed), hashed))


def test_plain_format_evens_out_mixed_font_size_and_color_20260911(monkeypatch):
    """AI が format で font・size・color のどれかを落とすたびに 1 往復失った。plain の範囲がまだらなら道具がそろえる。"""
    import types
    import vbam_edit as ve
    font = types.SimpleNamespace(Name=None, Size=None, Color=None, Bold=True, Italic=True,
                                 Underline=2, Strikethrough=True)
    rng = types.SimpleNamespace(Font=font, Interior=types.SimpleNamespace(Pattern=1, Color=0),
                                Address="$A$6:$I$30")
    xl = types.SimpleNamespace(StandardFont="游ゴシック", StandardFontSize=11.0)
    import vbam_core as vcore

    def _no_excel(*a, **k):
        raise RuntimeError("test: no Excel")          # protect_safe の包み紙も本物の Excel に触らせない
    monkeypatch.setattr(vcore, "get_workbook", _no_excel)
    monkeypatch.setattr(ve, "get_workbook", lambda *a, **k: (xl, object()))
    monkeypatch.setattr(ve, "_whole_sheet_spec", lambda *a, **k: None)
    monkeypatch.setattr(ve, "_resolve_range", lambda *a, **k: (types.SimpleNamespace(Name="顧客"), rng))
    args = types.SimpleNamespace(posargs=["A6:I30"], plain=True, unbold=True, bg="none",
                                 font=None, size=None, color=None, bold=False, italic=False,
                                 number_format=None, align=None, valign=None, wrap=False, border=None,
                                 col_width=None, row_height=None, merge=False, unmerge=False,
                                 autofit=False, lock=False, unlock=False, sheet_opt=None,
                                 whole_sheet=False, show=False)
    ve.cmd_format_range(args)
    assert (font.Name, font.Size, font.Color) == ("游ゴシック", 11.0, 0)
    assert font.Italic is False and font.Underline == -4142 and font.Strikethrough is False
    # そろっている範囲（None でない）は触らない
    font2 = types.SimpleNamespace(Name="Meiryo UI", Size=9.0, Color=255, Bold=False, Italic=False,
                                  Underline=-4142, Strikethrough=False)
    rng.Font = font2
    ve.cmd_format_range(args)
    assert (font2.Name, font2.Size, font2.Color) == ("Meiryo UI", 9.0, 255)


def test_hash_scan_reads_the_range_once_and_asks_only_the_tight_cells_20260913():
    """### の検査は範囲を 1 回で読み、列幅に届きそうな数値のセルだけ .Text を読む（Gemini が
    一括読みに替えて ### が落ちた分を戻した・2026-09-13）。文字のセルは聞かない。"""
    from vbam_view import _hash_scan
    rows = [["品名", "金額", "備考"], ["ペン", 1250000, "###ではない文字"], ["机", 12, ""]]
    asked = []

    class _R(_AuditRange):
        @property
        def Columns(self):
            n = len(self._rows[0])
            widths = {1: 12.0, 2: 8.0, 3: 30.0}                  # 金額の列だけ狭い（12 は入る・1250000 は届く）
            return type("C", (), {"Count": n,
                                  "__call__": lambda s, j: type("W", (), {"ColumnWidth": widths[j]})()})()

        def Cells(self, i, j):
            asked.append((i, j))
            return super().Cells(i, j)

    rng = _R(rows, texts={(2, 2): "######"})
    n, bad, note = _hash_scan(rng)
    assert (n, bad, note) == (1, ["B2"], "")
    assert asked == [(2, 2)]                                     # 1250000 だけ聞いた。12 も文字も聞かない


def test_tidy_evens_out_a_mixed_number_format_20260911():
    """#,##0 と標準がまだらの金額列を、tidy は「既に書式あり」と見て素通りし、仕上げ検査も見逃した。"""
    from vbam_edit import audit_table, _tidy_column_styles
    rows = [["会員番号", "氏名", "金額"], [1, "青木", 1250000], [2, "石川", 980000]]
    mixed = _AuditRange(rows, fmts={3: None}, aligns={1: -4131})
    assert any("表示形式がセルごとに違います" in n for n in audit_table(_AuditWS(mixed), mixed))

    class _WS(_AuditWS):
        def __init__(self, rng):
            super().__init__(rng)
            self.cols = {}

        def Range(self, a, b=None):
            j = ord(str(a.Address).replace('$', '')[0]) - 64
            if j not in self.cols:
                self.cols[j] = super().Range(a, b)
            return self.cols[j]
    ws = _WS(mixed)
    _tidy_column_styles(ws, mixed)
    assert ws.cols[3].NumberFormat == '#,##0'
    decided = _AuditRange(rows, fmts={3: '0.0%'}, aligns={1: -4131})   # 1 つに決まった書式は触らない
    ws = _WS(decided)
    _tidy_column_styles(ws, decided)
    assert ws.cols[3].NumberFormat == '0.0%'
    # 日付だけの列で表示形式がまだら（令和と西暦）なら yyyy/m/d に（2026-09-11 14:03 の差し戻し）
    import datetime as _dt
    drows = [["会員番号", "氏名", "入会日"], [1, "青木", _dt.datetime(2026, 3, 3)], [2, "石川", _dt.datetime(2026, 4, 1)]]
    dmixed = _AuditRange(drows, fmts={3: None}, aligns={1: -4131})
    assert any("日付列「入会日」(C) の表示形式がセルごとに違います" in n for n in audit_table(_AuditWS(dmixed), dmixed))
    ws = _WS(dmixed)
    _tidy_column_styles(ws, dmixed)
    assert ws.cols[3].NumberFormat == 'yyyy/m/d'
    dset = _AuditRange(drows, fmts={3: 'yyyy/mm/dd'}, aligns={1: -4131})
    assert not any("入会日" in n for n in audit_table(_AuditWS(dset), dset))
    ws = _WS(dset)
    _tidy_column_styles(ws, dset)
    assert ws.cols[3].NumberFormat == 'yyyy/mm/dd'


@_needs_srv
def test_mcp_agent_passes_backups_undo_to_and_drop_case(monkeypatch):
    """MCP の agent ツールにも同じ口（backups／undo_to／drop_case）。"""
    import vba_mcp_server as ms
    seen = []
    monkeypatch.setattr(ms, "_submit", lambda line, timeout=600: seen.append(line) or "ok")
    fn = ms.agent.fn if hasattr(ms.agent, 'fn') else ms.agent
    fn(backups=True)
    fn(undo_to="2", force=True)
    fn(drop_case="観光")
    assert seen[0].endswith("--backups")
    assert "--undo 2 --force" in seen[1]
    assert "--drop-case" in seen[2] and "観光" in seen[2]


# ================================================================
# 2026-09-17: check の 3 規則（VBM012〜014）・compile の行名指し・vba の --bg
# ================================================================

_LABEL_SRC = '''
Sub ラベル無し()
    On Error GoTo Errorgo
    x = 1
End Sub

Sub ラベルあり()
    On Error GoTo errorgo
    GoSub sub1
    x = 1
    Exit Sub
ErrorGo:
    Resume Next
sub1:
    Return
End Sub

Sub 行番号()
10  On Error GoTo 20
    x = 1
20  Resume Next
End Sub

Sub 文字列の中()
    s = "GoTo Nowhere"
    On Error GoTo 0
    On Error GoTo -1
    ' GoTo コメントの中
    Resume Next
End Sub

Sub 打ち間違い()
    On Err GoTo koko
koko:
End Sub

Sub 正しい()
    On Error GoTo koko
    If x Then Else: y = 1
    Select Case x
        Case 1: y = 2
    End Select
    handler: Resume Next
    On Error GoTo handler
    Application.GoTo Reference:="Print_Area"
    Application.Goto ws.Range("A1"), True
koko:
End Sub
'''


def _label_lines():
    return _LABEL_SRC.replace('\r\n', '\n').split('\n')


def test_vbm012_flags_only_the_missing_label():
    lines = _label_lines()
    found = vv._diag_missing_labels(lines, vv._split_procedures(lines))
    assert [(f[0], f[2], f[3]) for f in found] == [('ラベル無し', 'Errorgo', 'On Error GoTo Errorgo')]
    assert found[0][1] == lines.index('    On Error GoTo Errorgo') + 1


def test_vbm012_silent_for_case_line_numbers_strings_keywords_and_same_line_labels():
    lines = _label_lines()
    names = {f[0] for f in vv._diag_missing_labels(lines, vv._split_procedures(lines))}
    for ok in ('ラベルあり', '行番号', '文字列の中', '正しい'):
        assert ok not in names, ok


def test_vbm013_flags_on_err_typo_only():
    found = vv._diag_on_err_typo(_label_lines())
    assert [(f[1], f[2]) for f in found] == [('Err', 'On Err GoTo koko')]
    assert vv._diag_on_err_typo(['On Error GoTo x', 'On Error Resume Next', '    On Erorr GoTo y',
                                 'On n GoTo 10, 20', 's = "On Err GoTo z"',
                                 'On ErrKind GoTo L1, L2']) == [(3, 'Erorr', 'On Erorr GoTo y')]   # 計算型 GoTo は除く


def test_check_vbm014_counts_call_but_not_run(monkeypatch, capsys):
    """VBM014 は Call／裸呼びの未解決だけ error。Application.Run "名前" はアドイン等を実行時に探すので数えない。"""
    import argparse

    class _CM:
        def __init__(self, code):
            self._code = code
            self.CountOfLines = code.count('\r\n') + 1

        def Lines(self, a, n):
            return self._code

    class _Comp:
        def __init__(self, name, code):
            self.Name, self.Type, self.CodeModule = name, 1, _CM(code)

    code = ('Sub 呼ぶ()\r\n    On Error GoTo 0\r\n    Call 無いマクロ\r\n'
            '    Application.Run "アドインのマクロ"\r\n    Application.Run "\'秀.xlam\'!別の本"\r\nEnd Sub')
    wb = type("WB", (), {"Name": "t.xlsm",
                         "VBProject": type("P", (), {"VBComponents": [_Comp("M", code)]})()})()
    monkeypatch.setattr(vv, "get_workbook", lambda *a, **k: (None, wb))
    ok = vv.cmd_check(argparse.Namespace(posargs=[], json=True))
    doc = json.loads(capsys.readouterr().out)
    assert ok is False and doc["summary"]["errors"] == 1
    assert [u["name"] for u in doc["unresolved"]] == ["無いマクロ"]


def test_check_parser_has_all_warnings_and_rules_footer_mentions_it(capsys):
    import argparse
    ns = vm.build_parser().parse_args(['check', '--all-warnings'])
    assert ns.all_warnings is True
    vv.cmd_rules(argparse.Namespace(posargs=[], json=False))
    assert "--all-warnings" in capsys.readouterr().out


def test_compile_error_detail_names_the_line_and_hints_the_fix():
    loc = {"module": "写真貼り付け", "proc": "写真の貼り付け", "line": 82, "text": "On Error GoTo Errorgo"}
    assert (vv._compile_error_detail(loc, "コンパイル エラー:\n行ラベルが定義されていません。")
            == "[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― コンパイル エラー: 行ラベルが定義されていません。")
    assert vv._compile_error_detail(None, " 本文 \n 2 行目 ") == "本文 2 行目"
    assert vv._compile_error_detail(None, "") == "コンパイル エラー（本文を取得できませんでした）"
    # ダイアログ監視の本文は「タイトル「…」 本文: …」の形＝タイトルは落とす（2026-09-17 実射）
    assert vv._compile_error_detail(loc, "タイトル「Microsoft Visual Basic for Applications」 本文: コンパイル エラー: 行ラベルが定義されていません。") \
        == "[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― コンパイル エラー: 行ラベルが定義されていません。"
    hint = vv._compile_fix_hint(loc, "行ラベルが定義されていません。")
    assert "replace-procedure -y --module 写真貼り付け" in hint and "get 写真貼り付け 写真の貼り付け" in hint
    assert vv._compile_fix_hint(loc, "Sub または Function が定義されていません。") == \
        '直す: code-replace "On Error GoTo Errorgo" "新しい行" --module 写真貼り付け -y'
    assert "VBE" in vv._compile_fix_hint(None)


@_needs_srv
def test_mcp_vba_bg_strips_the_flag_and_runs_without_waiting(monkeypatch):
    """長い手は行末の --bg で待たずに job id を返す（無言消し・2026-09-17）。関所の拒否文はそのまま返す。"""
    seen = {}

    def fake_async(line):
        seen['line'] = line
        return "job1758000000"

    def fake_sync(line, timeout=600):
        raise AssertionError("--bg なのに同期で撃った: " + line)

    monkeypatch.setattr(_srv, "_submit_async", fake_async)
    monkeypatch.setattr(_srv, "_submit", fake_sync)
    fn = _srv.vba.fn if hasattr(_srv.vba, 'fn') else _srv.vba
    out = fn("checkup --form --bg")
    assert seen['line'] == "checkup --form"
    assert "job1758000000" in out and "agent_status" in out and "status" in out
    monkeypatch.setattr(_srv, "_submit_async", lambda line: "この操作（export-all）はユーザーが確認画面で却下しました。")
    assert fn("export-all --bg").startswith("この操作")
    # --bg が無ければ今までどおり同期
    monkeypatch.setattr(_srv, "_submit", lambda line, timeout=600: "sync:" + line)
    assert fn("list --bgx") == "sync:list --bgx"


def test_the_money_column_gets_commas_even_with_not_applicable_text():
    """突き合わせで「申込なし」が混じる金額列にも桁区切りを当てる（2026-09-06 の実射）。

    同じ依頼を 2 回撃って、1 回目だけ AI が format を足し、2 回目は付けず、どちらも採点は
    「満たしていない点なし」で通った＝見た目が揺れた。tidy の側で決まるようにする。
    """
    import vbam_edit as ve

    assert ve._column_style('金額', [12000, '申込なし', 7000, '申込なし']) == 'int'
    assert ve._column_style('単価', [1200.5, '-', 980.0]) == 'dec'
    assert ve._column_style('金額', [12000, '申込なし']) is None        # 数値 1 つでは決めない
    assert ve._column_style('備考', [12000, '申込なし', 7000]) is None  # 金額の見出しでなければ従来どおり
    assert ve._column_style('会員番号', [1001, 1002]) == 'id'
    assert ve._column_style('年度', [2024, 2025]) is None


def test_katakana_words_with_key_or_code_are_not_id_columns():
    """2026-09-19: 商品名「キーボード」の売上の列が番号列と判定され、左寄せ・カンマなしになった。
    「キー」「コード」は、カタカナの語の一部（キーボード・キーワード・レコード・コードレス）なら番号列にしない。"""
    import vbam_edit as ve
    import vbam_hands as vh

    assert ve._column_style('キーボード', [184000, 201000, 176000]) == 'int'
    assert ve._column_style('レコード数', [12, 30, 45]) == 'int'
    for h in ('キーワード', 'コードレス掃除機', 'ノートPC'):
        assert not ve._is_id_header(h), h
    for h in ('キー', '照合キー', '商品コード', 'バーコード', '会員番号', 'No', '社員ID'):
        assert ve._is_id_header(h), h
    assert vh._id_columns(['キーボード', '商品コード', '金額']) == {1}
    assert not ve._looks_like_key('キーボード') and ve._looks_like_key('照合キー')


def test_tidy_keeps_money_whole_and_skips_blank_rows_20260913():
    """お試し版テスト用1: 平均 10,162.5 が 1 つあるだけで金額 20 行が 7,200.00 に、空の行と合計欄の左にも罫線が付いた。"""
    import vbam_edit as ve
    assert ve._column_style('金額', [7200.0] * 20 + [203250.0, 10162.5]) == 'int'   # 小数が 1 つだけ＝整数の列
    assert ve._column_style('単価', [12.5, 3.0]) == 'dec'                            # 半分が小数＝小数の列（従来どおり）
    head = ["日付", "商品", "分類", "担当", "数量", "単価", "金額"]
    body = [["2026/8/1", "えんぴつ", "筆記具", "佐藤", 120, 60, 7200]] * 20
    rows = [head] + body + [[None] * 7, [None] * 5 + ["合計", 203250], [None] * 5 + ["平均", 10162.5]]
    blocks, note = ve._border_blocks(rows)
    assert blocks == [(1, 1, 21, 7), (23, 6, 24, 7)] and note     # 明細は全幅・合計と平均は F:G だけ・空の行は無し
    gap = [head, body[0], [None] * 7, body[0]]                     # 表の途中の空行（抜けた 1 件）は表の続き
    assert ve._border_blocks(gap) == ([(1, 1, 4, 7)], '')
    assert ve._border_blocks([head, body[0]]) == ([(1, 1, 2, 7)], '')

    # 整数の列の中の小数のセルだけ #,##0.0#（列は #,##0）。人が決めていた形は戻す
    class _Cell:
        def __init__(self, fmt='G/標準'):
            self.NumberFormat = fmt

    class _Col:
        def __init__(self, vals, fmts):
            self.Value = [(v,) for v in vals]
            self.NumberFormat = 'G/標準'
            self.HorizontalAlignment = 1
            self.cells = [_Cell(f) for f in fmts]

        def Cells(self, i, j):
            return self.cells[i - 1]

    class _Hdr:
        def __init__(self, v, j):
            self.Value, self.Address = v, "$" + chr(64 + j) + "$1"

    vals = [7200.0] * 5 + [203250.0, 10162.5]
    col = _Col(vals, ['G/標準'] * 7)

    class _Rng:
        Rows = type("R", (), {"Count": len(vals) + 1})()
        Columns = type("C", (), {"Count": 1})()

        def Cells(self, i, j):
            return _Hdr("金額" if i == 1 else None, j)

    class _WS:
        def Range(self, a, b=None):
            return col

    notes = ve._tidy_column_styles(_WS(), _Rng())
    assert col.NumberFormat == '#,##0'
    assert col.cells[6].NumberFormat == '#,##0.0#' and col.cells[5].NumberFormat == 'G/標準'
    assert any("小数のセルだけ" in n for n in notes)
    col = _Col(vals, ['G/標準'] * 6 + ['#,##0.0'])                  # 人が #,##0.0 と決めていた平均はそのまま
    ve._tidy_column_styles(_WS(), _Rng())
    assert col.cells[6].NumberFormat == '#,##0.0'


def test_seiri_one_command_names_the_error_cause_and_takes_the_screen_book_20260913(monkeypatch):
    """shu「まとめてみろ」「エラーが出てるんだから気付かなきゃおかしい」: マクロを撃って残りだけ出す 1 手。

    あわせて、対象のブックは画面の窓から取る（更新登録の後 ActiveWorkbook が消えた 秀コンボ.xlsm を指したまま残った）。
    """
    import vbam_view as vw
    import vbam_core as vcore
    import vbam_vba as vvba
    assert vm.cmd_seiri is vw.cmd_seiri
    assert "E100" in vw.error_hint('#DIV/0!', '=G27/E100', ['E100'])
    assert vw.error_hint('#DIV/0!', '=G27/SUM(E6:E25)', []) == ''
    assert vw.error_hint('#N/A', '', []).startswith('検索')
    refs = vw._CELL_REF_RE.findall('=G27/SUM(E6:E25)+Sheet1!B2+$C$3')
    assert refs == [('G', '27'), ('E', '6'), ('E', '25'), ('C', '3')]      # 別シートの B2 と関数名は拾わない

    class _Win:
        Visible = True
        Parent = "お試し版 Excelコンボ.xlsm"

    class _App:
        ActiveWindow = _Win()
        ActiveWorkbook = "秀コンボ.xlsm"                                      # 更新登録で画面から消えたブック

    assert vcore._screen_book(_App()) == "お試し版 Excelコンボ.xlsm"

    class _App2:
        ActiveWindow = None
        ActiveWorkbook = "Book1"

    assert vcore._screen_book(_App2()) == "Book1"

    class _CM:
        def __init__(self, text):
            self.text = text
            self.CountOfLines = text.count("\n") + 1

        def Lines(self, a, b):
            return self.text

    class _Proj:
        def __init__(self, name, mods):
            self.name, self.mods = name, mods

        def VBComponents(self, m):
            return type("C", (), {"CodeModule": _CM(self.mods[m])})()

    projects = [_Proj("PERSONAL.XLSB", {}), _Proj("秀コンボ.xlam", {"表の整理": "Sub 表を整える()\nEnd Sub\nSub 重複行を消す()"})]
    xl = type("X", (), {"VBE": type("V", (), {"VBProjects": projects})()})()
    monkeypatch.setattr(vvba, "_project_book_name", lambda _xl, p: p.name)
    assert vw._macro_book(xl, "表の整理", "表を整える") == "秀コンボ.xlam"
    assert vw._macro_book(xl, "表の整理", "無いマクロ") is None


def test_content_findings_catch_errors_sum_mismatch_and_broken_formula_column():
    """中身の検査（done を止める側）: エラー値・「合計」行が上の和と合わない・数式の列に直値が混ざる。"""
    from vbam_edit import _content_findings
    vals = [["品名", "単価", "数量", "金額"],
            ["りんご", 100, 3, 300],
            ["みかん", 80, 5, 400],
            ["ぶどう", -2146826265, 2, 600],      # #REF! が単価に
            ["なし", 120, 4, 480],
            ["合計", None, 14, 1700]]             # 金額の和は 1780
    fmls = [["品名", "単価", "数量", "金額"],
            ["りんご", 100, 3, "=B2*C2"],
            ["みかん", 80, 5, "=B3*C3"],
            ["ぶどう", "=#REF!", 2, 600],         # 金額が直値
            ["なし", 120, 4, "=B5*C5"],
            ["合計", None, "=SUM(C2:C5)", 1700]]
    block, seen = _content_findings(vals, fmls, 1, 1)
    joined = " / ".join(block)
    assert "エラー値" in joined and "B4(#REF!)" in joined
    assert "「合計」行の D6" in joined and "1,780" in joined
    assert "数式の列「金額」(D) に直値が 1 セル" in joined and "D4" in joined
    assert not any("合計」行の C6" in b for b in block)     # 数量の合計 14 は合っている
    assert seen == []


def test_content_findings_notice_blanks_text_numbers_and_duplicate_ids_without_blocking():
    """中身の検査（気づき側）: 表の中の空欄・数値列の文字の数字・番号列の重複は止めずに報告へ。"""
    from vbam_edit import _content_findings
    vals = [["会員番号", "氏名", "フリガナ", "金額", "備考"],
            ["1001", "青木", "アオキ", 1200, None],
            ["1002", "石川", "イシカワ", "１，５００", None],
            ["1003", "伊藤", None, 900, "退会予定"],       # C4 が空欄（初実射の C10 と同じ）
            ["1002", "上田", "ウエダ", 700, None],          # 会員番号の重複
            ["1005", "遠藤", "エンドウ", 1100, None]]
    block, seen = _content_findings(vals, None, 7, 1)          # 表は A7 から
    assert block == []
    joined = " / ".join(seen)
    assert "表の中の空欄が 1 セル（C10）" in joined
    assert "備考" not in joined                                # まばらな列は数えない
    assert "数値列「金額」(D) に文字の数字が 1 セル" in joined and "D9" in joined
    assert "番号列「会員番号」(A) に重複が 1 組" in joined and "A9 A11" in joined


def test_content_findings_ignore_legit_shapes():
    """止めてはいけない形: 合計行が 1 本だけ数式（本文は直値）／小計ごとの和／空の見出し列／2 行の表。"""
    from vbam_edit import _content_findings
    vals = [["部署", "4月", "5月"],
            ["総務", 10, 20], ["経理", 30, 40], ["小計", 40, 60],
            ["営業1", 5, 6], ["営業2", 7, 8], ["小計", 12, 14],
            ["合計", 52, 74]]
    fmls = [["部署", "4月", "5月"],
            ["総務", 10, 20], ["経理", 30, 40], ["小計", "=SUM(B2:B3)", "=SUM(C2:C3)"],
            ["営業1", 5, 6], ["営業2", 7, 8], ["小計", "=SUM(B5:B6)", "=SUM(C5:C6)"],
            ["合計", "=B4+B7", "=C4+C7"]]
    block, seen = _content_findings(vals, fmls, 1, 1)
    assert block == [] and seen == []
    assert _content_findings([["a", "b"]], None, 1, 1) == ([], [])
    assert _content_findings([["a", "b"], [1, 2]], None, 1, 1) == ([], [])

def test_content_findings_catch_total_column_short_formula_and_dead_lookup():
    """中身の検査の追加分: 「合計」列が左の和と合わない／式が途中で切れる／突き合わせが全滅。"""
    from vbam_edit import _content_findings
    vals = [["支店", "4月", "5月", "合計"],
            ["秋田", 10, 20, 30],
            ["能代", 30, 40, 71],          # 左の和は 70
            ["大館", 5, 6, 11]]
    fmls = [["支店", "4月", "5月", "合計"],
            ["秋田", 10, 20, "=SUM(B2:C2)"],
            ["能代", 30, 40, 71],
            ["大館", 5, 6, "=SUM(B4:C4)"]]
    block, _seen = _content_findings(vals, fmls, 1, 1)
    joined = " / ".join(block)
    assert "「合計」列の D3 が左の和と合いません" in joined and "70" in joined
    assert "直値が 1 セル" in joined and "D3" in joined

    # 式が最後の本文行まで届いていない
    vals2 = [["品名", "数量", "税込"],
             ["a", 1, 108], ["b", 2, 216], ["c", 3, None], ["d", 4, None]]
    fmls2 = [["品名", "数量", "税込"],
             ["a", 1, "=B2*108"], ["b", 2, "=B3*108"], ["c", 3, None], ["d", 4, None]]
    block2, _ = _content_findings(vals2, fmls2, 1, 1)
    assert any("式が C3 で切れています" in b and "本文が 2 行" in b for b in block2)

    # 突き合わせが全滅（キーの型違い）
    head = ["会員番号", "氏名", "申込"]
    body = [[f"100{i}", f"氏名{i}", "申込なし"] for i in range(1, 7)]
    fml = [head] + [[f"100{i}", f"氏名{i}", f'=IFERROR(VLOOKUP(A{i + 1},T,2,FALSE),"申込なし")'] for i in range(1, 7)]
    block3, _ = _content_findings([head] + body, fml, 1, 1)
    assert any("突き合わせの列「申込」(C) が 6/6 行とも「申込なし」" in b for b in block3)


def test_content_findings_do_not_cry_on_legit_lookup_and_totals():
    """止めてはいけない形: 突き合わせの結果がばらけている／合計列が合っている／本文が 4 行未満。"""
    from vbam_edit import _content_findings
    head = ["会員番号", "氏名", "申込"]
    rows = [head] + [[f"100{i}", f"氏名{i}", ("済" if i % 2 else "申込なし")] for i in range(1, 7)]
    fml = [head] + [[f"100{i}", f"氏名{i}", f'=VLOOKUP(A{i + 1},T,2,FALSE)'] for i in range(1, 7)]
    block, _ = _content_findings(rows, fml, 1, 1)
    assert block == []
    ok = [["支店", "4月", "5月", "合計"], ["秋田", 10, 20, 30], ["能代", 30, 40, 70]]
    assert _content_findings(ok, None, 1, 1)[0] == []
    short = [["会員番号", "氏名", "申込"], ["1", "a", "なし"], ["2", "b", "なし"], ["3", "c", "なし"]]
    sf = [["会員番号", "氏名", "申込"], ["1", "a", "=X"], ["2", "b", "=X"], ["3", "c", "=X"]]
    assert not any("突き合わせ" in b for b in _content_findings(short, sf, 1, 1)[0])

def test_total_column_ignores_id_columns_and_stays_quiet_when_the_layout_does_not_fit():
    """合計列の検査: 番号列・年は足さない／1 行も合わないときは黙る（見当違いの咎めを出さない）。"""
    from vbam_edit import _content_findings
    rows = [["会員番号", "4月", "5月", "合計"],
            [1001, 10, 20, 30],
            [1002, 30, 40, 71],          # 左の和は 70（ここだけ違う）
            [1003, 5, 6, 11]]
    block, _ = _content_findings(rows, None, 1, 1)
    joined = " / ".join(block)
    assert "「合計」列の D3 が左の和と合いません" in joined and "70" in joined
    assert "D2" not in joined and "D4" not in joined      # 会員番号を足していない

    odd = [["支店", "4月", "5月", "合計"],           # 合計が左の和と無関係な列（税抜だけ等）
           ["秋田", 10, 20, 8], ["能代", 30, 40, 9], ["大館", 5, 6, 7]]
    assert not any("左の和" in b for b in _content_findings(odd, None, 1, 1)[0])

    y = [["年", "4月", "合計"], [2024, 10, 10], [2025, 20, 20]]
    assert not any("左の和" in b for b in _content_findings(y, None, 1, 1)[0])


def test_formula_lens_reads_magic_numbers_and_aggregate_ranges():
    """数式の目の部品（純 Python）: 直書きの数値と、集計の範囲を式から取り出す（2026-09-09）。"""
    from vbam_view import fml_magic_numbers, fml_agg_ranges
    # 掛け算・割り算・比較に隣り合う数だけが「直書き」
    assert fml_magic_numbers("=B7*C7*1.05") == ["1.05"]
    assert fml_magic_numbers("=E5*0.21+A1") == ["0.21"]
    assert fml_magic_numbers('=IF(D2>=60,"合格","不合格")') == ["60"]
    # 引数の位置の数は直書きではない（VLOOKUP の列番号・ROUND の桁・SUBTOTAL の関数番号・DATE の年）
    assert fml_magic_numbers('=IFERROR(VLOOKUP(A2,$I$2:$K$60,3,FALSE),"申込なし")') == []
    assert fml_magic_numbers("=ROUND(B2*C2,2)") == []
    assert fml_magic_numbers("=SUBTOTAL(9,C2:C9)") == []
    assert fml_magic_numbers("=DATE(2026,1,1)") == []
    # 0・1 と、数式でないもの
    assert fml_magic_numbers("=A2/0") == [] and fml_magic_numbers('="a"+1') == []
    assert fml_magic_numbers(123) == [] and fml_magic_numbers("合計") == []
    assert fml_magic_numbers('=IF(A1="1.05","x","y")') == []      # 文字列の中の数は見ない
    # 集計の範囲（同じシートのものだけ。入れ子と別シートは黙る）
    assert fml_agg_ranges("=SUM(D2:D7)") == [("SUM", [(2, 4, 7, 4)])]
    assert fml_agg_ranges("=SUM(B2:B4,B6:B9)") == [("SUM", [(2, 2, 4, 2), (6, 2, 9, 2)])]
    assert fml_agg_ranges("=SUBTOTAL(9,C2:C9)") == [("SUBTOTAL", [(2, 3, 9, 3)])]
    assert fml_agg_ranges("=SUM('集計'!B2:B9)") == []
    assert fml_agg_ranges("=SUM(IF(A2:A9>0,B2:B9))") == []
    assert fml_agg_ranges("=B4+B7") == [] and fml_agg_ranges("=LOG10(A1)") == []


def test_formula_lens_finds_the_odd_formula_and_leaves_legit_shapes_alone():
    """数式の目の本体: 列の中で 1 つだけ形の違う式・集計の起点行の食い違い。合っている形は咎めない。"""
    from vbam_view import formula_notes
    # D7 だけ *1.05 が付いている（手順書「表と数式の点検」の練習台と同じ形）
    cells = {(r, 4): (f"=B{r}*C{r}", "=RC[-2]*RC[-1]") for r in (2, 3, 4, 6, 8)}
    cells[(7, 4)] = ("=B7*C7*1.05", "=RC[-2]*RC[-1]*1.05")
    cells[(9, 4)] = ("=SUM(D2:D8)", "=SUM(R[-7]C:R[-1]C)")      # 合計行は「形が違う」に数えない
    text = " / ".join(formula_notes(cells))
    assert "列の中で式の形が違うセル: D7（=B7*C7*1.05／他 5 行は" in text
    assert "1.05（D7・1セル）" in text
    # 累計（先頭行だけ形が違う）・全部同じ・数が少ない列では黙る
    run = {(2, 5): ("=D2", "=RC[-1]")}
    run.update({(r, 5): (f"=E{r - 1}+D{r}", "=R[-1]C+RC[-1]") for r in range(3, 9)})
    assert "先頭行だけ＝累計の書き始めならこのままでよい" in " / ".join(formula_notes(run))
    same = {(r, 4): (f"=B{r}*C{r}", "=RC[-2]*RC[-1]") for r in range(2, 12)}
    assert formula_notes(same) == []
    # 合計行の起点行がそろっていない（B は 2 行目から・C だけ 3 行目から）
    tot = {(10, c): (f"=SUM({chr(64 + c)}2:{chr(64 + c)}9)", None) for c in (2, 4, 5)}
    tot[(10, 3)] = ("=SUM(C3:C9)", None)
    text2 = " / ".join(formula_notes(tot))
    assert "10 行目の集計の起点行がそろっていません" in text2 and "C10 だけ違う" in text2
    # 合計列の下の総計（J2:J4 が行合計・J5 がその総計）は「形が違う」に数えない（実射で出た雑音）
    gt = {(r, 10): (f"=SUM(G{r}:I{r})", "=SUM(RC[-3]:RC[-1])") for r in (2, 3, 4)}
    gt[(5, 10)] = ("=SUM(J2:J4)", "=SUM(R[-3]C:R[-1]C)")
    assert not any("形が違う" in n for n in formula_notes(gt))
    assert formula_notes({}) == [] and formula_notes({(1, 1): ("見出し", None)}) == []


def test_formula_lens_spots_a_summary_that_cannot_be_copied_across():
    """横に並ぶ同じ関数の集計が列ごとに違う式＝月を式に埋めた形（2026-09-09 の実射で出た）。

    値は合うので値の検査は通る。月を 1 つ足した瞬間に壊れるので、道具が言葉にする。
    """
    from vbam_view import formula_notes, formula_copy_blockers
    bad = {}
    for j, mon in enumerate((1, 2, 3)):
        for r in (2, 3, 4):
            bad[(r, 7 + j)] = (f'=SUMPRODUCT((MONTH($A$2:$A$19)={mon})*($B$2:$B$19=$F{r})*$D$2:$D$19)',
                               f'=SUMPRODUCT((MONTH(R2C1:R19C1)={mon})*(R2C2:R19C2=RC6)*R2C4:R19C4)')
    for r in (2, 3, 4):
        bad[(r, 10)] = (f"=SUM(G{r}:I{r})", "=SUM(RC[-3]:RC[-1])")      # 右端の行合計は咎めない
    text = " / ".join(formula_copy_blockers(bad))          # 気づきではなく done を止める側
    assert "2〜4 行目の SUMPRODUCT が列ごとに違う式です（G2:I4）＝横にコピーできません" in text
    assert "J2" not in text and "J4" not in text
    assert not any("列ごとに違う" in n for n in formula_notes(bad))
    # 横にコピーできる形（見出しを参照する SUMIFS）は黙る
    ok = {(r, 7 + j): ('=SUMIFS($D:$D,$B:$B,$F2,$A:$A,">="&G$1)', '=SUMIFS(C4,C2,RC6,C1,">="&R1C)')
          for j in range(3) for r in (2, 3, 4)}
    assert formula_copy_blockers(ok) == []
    # 件数・合計・平均が横に並ぶ表も黙る（関数名が違う＝式が違って当たり前）
    mix = {}
    for r in (2, 3, 4):
        mix[(r, 2)] = ("=COUNTA(B2:B9)", "=COUNTA(R2C2:R9C2)")
        mix[(r, 3)] = ("=SUM(C2:C9)", "=SUM(R2C3:R9C3)")
        mix[(r, 4)] = ("=AVERAGE(D2:D9)", "=AVERAGE(R2C4:R9C4)")
    assert formula_copy_blockers(mix) == []


def test_content_findings_catch_a_total_whose_range_misses_body_rows():
    """集計の範囲が本文行を取りこぼしている（値がたまたま合っていても止める・2026-09-09）。

    2. の値の突き合わせは「今の数」しか見ないので、取りこぼした行がまだ空なら気づけない
    （人がその行を埋めた瞬間に合計が狂う）。範囲そのものを見る目をここで持つ。
    """
    from vbam_edit import _content_findings
    vals = [["品名", "数量", "金額"],
            ["a", 1, 100], ["b", 2, 200], ["c", 3, None], ["合計", 6, 300]]
    fmls = [["品名", "数量", "金額"],
            ["a", 1, 100], ["b", 2, 200], ["c", 3, None], ["合計", "=SUM(B2:B4)", "=SUM(C2:C3)"]]
    block, _seen = _content_findings(vals, fmls, 1, 1)
    joined = " / ".join(block)
    assert "「合計」行の C5 の SUM が本文 1 行を範囲に入れていません（C4）" in joined
    assert "範囲を C2:C4 に直す" in joined
    assert not any("B5" in b for b in block)              # 数量は取りこぼしていない
    # 見出し行まで範囲に入れているのは気づき（止めない）
    f2 = [r[:] for r in fmls]
    f2[4][2] = "=SUM(C1:C4)"
    block2, seen2 = _content_findings(vals, f2, 1, 1)
    assert not any("範囲に入れていません" in b for b in block2)
    assert any("見出し行（C1）まで範囲に入れています" in s for s in seen2)
    # 小計を足し上げる合計（=SUM(B4,B7)）は本文を直に足す形ではないので咎めない
    v3 = [["部署", "4月"], ["総務", 10], ["経理", 30], ["小計", 40],
          ["営業1", 5], ["営業2", 7], ["小計", 12], ["合計", 52]]
    f3 = [["部署", "4月"], ["総務", 10], ["経理", 30], ["小計", "=SUM(B2:B3)"],
          ["営業1", 5], ["営業2", 7], ["小計", "=SUM(B5:B6)"], ["合計", "=SUM(B4,B7)"]]
    assert _content_findings(v3, f3, 1, 1)[0] == []
    # 文字として入った式（Excel は計算していない）は気づきで出す＝止めない
    v4 = [["品目", "金額"], ["a", "=B2*2"], ["b", 200], ["c", 300]]
    b4, s4 = _content_findings(v4, None, 1, 1)
    assert b4 == [] and any("式が文字として入っているセルが 1 個" in x and "B2" in x for x in s4)


# ================================================================
# 対象ブックの固定（2026-09-08）
# 道具の内側から cmd_* を呼ぶ所は、対象ブックを固定してから呼ぶ。
# 付け忘れると get_workbook(None) の「アクティブブック自動検出」に落ち、
# 人が開いているブックを掴む（--fire の 1 本目が実際にここで落ちていた）。
# ================================================================

def test_pinned_workbook_holds_and_restores():
    import vbam_core as vc

    class _App:
        def __init__(self, active):
            self.ActiveWorkbook = active

    class _WB:
        def __init__(self, name):
            self.Name = name
            self.Application = _App(self)

    mine, other = _WB("Book1"), _WB("人のブック.xlsm")
    mine.Application.ActiveWorkbook = other      # 人が Excel 側で別のブックを選んでいる状態
    prev = vc._wb_cache.get("__active__")
    try:
        vc._wb_cache.pop("__active__", None)
        with vc.pinned_workbook(mine):
            _xl, wb = vc.get_workbook(None)
            assert wb is mine                    # 固定中は「切替に追従」しない
            assert vc._wb_pin_hard
        # 抜けたら元に戻る（常駐の MCP に固定が残らない）
        assert not vc._wb_pin_hard
        assert "__active__" not in vc._wb_cache
    finally:
        if prev is None:
            vc._wb_cache.pop("__active__", None)
        else:
            vc._wb_cache["__active__"] = prev


import datetime as _dt                                              # noqa: E402


# ----------------------------------------------------------------
# 2026-09-11: 動画のお題（BCG vs AI・顧客リストの整理）を 51 行で撃ったら、同じ会社の 2 行を両方消し・
# 繰り上がった行に電話番号を書き・「受注済みみ」を作り・それでも「合格」と出た。その穴を塞いだ分。
# ----------------------------------------------------------------
def _messy_rows_20260911():
    return [["会社名", "担当者名", "電話番号", "メール", "金額", "状態"],
            ["株式会社サンライズ", "田中 太郎", "03-1234-5678", "tanaka@example.co.jp", 1500000, "相談中"],
            ["株式会社スカイネット", "加藤 健", "０３１１１１９９９９", "KATO@EXAMPLE.JP", "¥3,500,000", "Active"],
            ["株式会社サンライズ ", "田中 太郎", "0312345678", "TANAKA@EXAMPLE.CO.JP", "1,500,000", "相談中"],
            ["株式会社スカイネット", "加藤 健", "03-1111-9999", "kato@example.jp", 3500000, "Active"],
            ["有限会社みらい工芸", "鈴木 健二", "045-111-2222", "suzuki@example.jp", 500000, "受注済み"],
            ["株式会社ハートビート", "木村 拓", "03-8888-2222", "kimura@example.jp", 300000, "相談中"]]


def test_duplicates_are_matched_through_spelling_and_lost_rows_are_seen_20260911():
    import vbam_hands as vh
    rows = _messy_rows_20260911()
    # 空白・全角・大文字・ハイフン・¥ とカンマの違いは同じ相手（値が丸ごと同じ行だけ数えていた＝26 組中 3 組）
    assert sorted((i, k) for i, k, _ in vh._dup_pairs(rows, 0)) == [(3, 1), (4, 2)]
    # 2 行とも消えたスカイネットは「消えた」、1 行残っていれば消えていない
    assert vh._lost_rows(rows, [rows[0], rows[1], rows[5], rows[6]], 0, 0) == [2, 4]
    assert vh._lost_rows(rows, [rows[0], rows[1], rows[2], rows[5], rows[6]], 0, 0) == []
    # 消す行のうち、残る行に相手のいない行（みらい工芸）を拾う
    assert vh._no_twin(rows, 0, {3, 5}) == [5]
    # 材料の「重複行」も同じ照合で数える
    import vbam_view as vv
    notes = "\n".join(vv.dirt_notes(rows, r0=5, c0=1, header_idx=0))
    assert "行8 = 行6" in notes and "行9 = 行7" in notes
    # 重複は全列一致が基本（2026-09-12・shu）: 状態が 1 列違えば重複ではない＝消さず「重複の疑い」に並べる
    rows_b = [list(r) for r in rows]
    rows_b[3][5] = "進行中"
    assert [(i, k) for i, k, _ in vh._dup_pairs(rows_b, 0)] == [(4, 2)]
    assert vh._near_dup_pairs(rows_b, 0) == [(3, 1, [5])]
    assert vh._no_twin(rows_b, 0, {3}) == [3]                         # 状態の違う行は消させない
    notes_b = "\n".join(vv.dirt_notes(rows_b, r0=5, c0=1, header_idx=0))
    assert "重複の疑い" in notes_b and "行8 と 行6（状態が違う）" in notes_b


def test_normalize_reads_yen_man_english_dates_and_lowercases_20260911():
    import vbam_hands as vh
    num = [("number", None)]
    assert vh._normalize_value("¥3,000,000", num)[0] == 3000000
    assert vh._normalize_value("95万円", num)[0] == 950000
    assert vh._normalize_value("500万", num)[0] == 5000000
    assert vh._normalize_value("0312345678", num)[0] == "0312345678"      # 先頭ゼロの番号は文字のまま
    d = vh._normalize_value("March 15, 2026", [("date", None)])[0]
    assert (d.year, d.month, d.day) == (2026, 3, 15)
    d = vh._normalize_value("April 2, 2026", [("date", None)])[0]
    assert (d.month, d.day) == (4, 2)
    assert vh._normalize_value(" KATO@EXAMPLE.JP", [("trim", None), ("lower", None)])[0] == "kato@example.jp"
    assert "lower" in vh._NORMALIZE_RULES


def test_a_name_column_is_not_an_id_column_after_dedupe_20260911():
    """重複を消した後は会社名が全部違う値になる。それを番号の列と見なすと、書き方の違う同じ会社を見逃す。"""
    import vbam_hands as vh
    rows = [["会社名", "担当者名", "電話番号", "メール", "金額"],
            ["株式会社アクアテック", "渡辺 隆", "06-1234-5678", "watanabe@example.jp", 750000],
            ["株式会社サンライズ", "田中 太郎", "03-1234-5678", "tanaka@example.co.jp", 1500000],
            ["（株） アクアテック", "渡辺 隆", "0612345678", "watanabe@example.jp", 750000],
            ["有限会社みらい工芸", "鈴木 健二", "045-111-2222", "suzuki@example.jp", 500000]]
    assert [(i, k) for i, k, _ in vh._dup_pairs(rows, 0)] == [(3, 1)]
    # 見出しが番号の列で値が違えば別の相手（伝票番号の違う同じ取引先・同じ品）
    rows2 = [["伝票番号", "取引先", "品名", "数量"],
             ["A-001", "サンライズ", "用紙", 10], ["A-002", "サンライズ", "用紙", 10], ["A-003", "みらい", "封筒", 5]]
    assert vh._dup_pairs(rows2, 0) == []


def test_phone_rule_splits_only_what_it_can_know_20260911():
    """名古屋の 052-333-4444 を AI の正規表現が 05-2333-4444 にした。区切りは道具が決め、決まらなければ残す。"""
    import vbam_hands as vh
    col = ["0523334444", "03-1234-5678", "0120555666", "+81-3-6666-5555", "+81-45-1111-2222",
           "０３１１１１９９９９", "09012345678", "0612345678", "0182-32-1234", "0182325678", "0187123456"]
    rules = [("trim", None), ("phone", None)]
    out, counts, _ex = vh._normalize_column(col, rules)
    assert out == ["052-333-4444", "03-1234-5678", "0120-555-666", "03-6666-5555", "045-1111-2222",
                   "03-1111-9999", "090-1234-5678", "06-1234-5678", "0182-32-1234", "0182-32-5678", "0187123456"]
    assert counts["phone"] == 8
    assert vh._phone_format("0920123456") == "0920123456"      # 0920（壱岐）を 092（福岡）で切らない
    assert "phone" in vh._NORMALIZE_RULES


def test_company_designators_do_not_split_the_same_company_20260911():
    """「株式会社アクアテック」と「アクアテック（株）」を別の会社として残し、採点係の差し戻しで 1 往復増えた。"""
    import vbam_hands as vh
    k = vh._loose_key("株式会社アクアテック")
    assert k == vh._loose_key("（株）アクアテック") == vh._loose_key("㈱アクアテック") == vh._loose_key(" 株式会社 アクアテック ")
    assert vh._loose_key("アクアテック（株）") == vh._loose_key("アクアテック株式会社") == vh._loose_key("アクアテック㈱")
    assert vh._loose_key("有限会社みらい工芸") == vh._loose_key("（有）みらい工芸")
    assert vh._loose_key("株式会社") == "株式会社"               # 種類の表記だけの値は消さない
    assert vh._loose_key("株式会社サンライズ") != vh._loose_key("株式会社スカイネット")
    # 前株・後株の位置は区別する（2026-09-12・全列一致が基本＝株式会社A と A株式会社 を同じにしない）
    assert vh._loose_key("株式会社アクアテック") != vh._loose_key("アクアテック株式会社")


# ---- 2026-09-16: 速さの改修 3 つ（materials の一括読み・呼び出し台帳 stats・再接続の軽量化）----
def test_formula_patterns_counts_from_bulk_grids():
    """数式の型は一括で読んだ 2 つの格子から数える（数式セル 1 つにつき COM 3 回＝5,000 セルで 15,000 往復を 0 に）。"""
    fa = [["品目", "金額", "=B5*2"], ["a", 100.0, "=B6*2"], ["b", 200.0, "=SUM(B5:B6)"], ["c", None, None]]
    fr = [["品目", "金額", "=RC[-1]*2"], ["a", 100.0, "=RC[-1]*2"],
          ["b", 200.0, "=SUM(R[-2]C[-1]:R[-1]C[-1])"], ["c", None, None]]
    total, pats = vm._formula_patterns(fa, fr, 5, 2)          # 左上が B5
    assert total == 3 and len(pats) == 2
    assert pats["=RC[-1]*2"] == [2, "D5", "=B5*2", "D6"]       # 個数・代表の番地・代表の A1 式・最後の番地
    assert pats["=SUM(R[-2]C[-1]:R[-1]C[-1])"] == [1, "D7", "=SUM(B5:B6)", "D7"]
    # 上限を超えたら数えるのをやめる（本数は上限+1＝呼び手が '+' を付ける）
    big = [["=A1"] * 10] * 10
    total, pats = vm._formula_patterns(big, big, 1, 1, cap=5)
    assert total == 6 and pats["=A1"][0] == 5
    assert vm._formula_patterns([[1.0, "x"]], [[1.0, "x"]], 1, 1) == (0, {})
    src = inspect.getsource(vm.cmd_materials)
    assert "_formula_patterns(" in src and "_hidden_count(" in src


def test_hidden_count_asks_whole_first_and_counts_only_when_mixed():
    calls = []

    class _Whole:
        def __init__(self, h):
            self.Hidden = h
    one = lambda i: calls.append(i) or (i % 2 == 0)           # noqa: E731
    assert vm._hidden_count(_Whole(False), 100, one) == 0 and calls == []
    assert vm._hidden_count(_Whole(True), 100, one) == 100 and calls == []
    assert vm._hidden_count(_Whole(None), 6, one) == 3 and calls == [1, 2, 3, 4, 5, 6]
    assert vm._hidden_count(_Whole(None), 10, one, max_count=5) is None     # 大きい混在は数えない


def test_merged_areas_skips_rows_without_merge():
    """結合セルの走査は行ごとに 1 回聞き、無い行はセルを見ない（8,000 セルで 8,000 往復だった）。"""
    asked = []

    row_asked = []

    class _Row:
        def __init__(self, i):
            row_asked.append(i)
            if isinstance(i, str):                             # "1:16" のブロック（2026-09-17）
                a, b = (int(x) for x in i.split(':'))
                self.MergeCells = None if a <= 3 <= b else False
            else:
                self.MergeCells = (i == 3)                     # 3 行目だけ結合がある

    class _Cell:
        def __init__(self, r, c):
            asked.append((r, c))
            self.MergeCells = (r == 3 and c <= 2)
            self.MergeArea = self                              # Address が無い＝古い道（Row/Column）へ落ちる
            self.Row, self.Column = 3, 1
            self.Rows = type("R", (), {"Count": 1})()
            self.Columns = type("C", (), {"Count": 2})()

    class _WS:
        def Rows(self, i):
            return _Row(i)

        def Columns(self, c):
            return type("Col", (), {"MergeCells": (c <= 2)})()  # 3 列目に結合は無い（2026-09-17: 列も先に飛ばす）

        def Cells(self, r, c):
            return _Cell(r, c)

    class _Rng:
        MergeCells = None
        Cells = type("Cs", (), {"Count": 12})()
        Worksheet = _WS()
        Row, Column = 1, 1
        Rows = type("R", (), {"Count": 40})()
        Columns = type("C", (), {"Count": 3})()
    areas, skipped = vm._merged_areas_in_range(_Rng())
    assert areas == ["A3:B3"] and skipped is None
    assert asked == [(3, 1)], "3 行目 × 結合のある列の交点だけ聞き、結合矩形の残りは飛ばす"
    assert row_asked == ["1:16", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, "17:32", "33:40"], \
        "結合の無いブロックは行ごとに聞かない"


def test_merged_areas_reads_the_rectangle_from_address_in_one_call():
    """結合の矩形は MergeArea.Address 1 回で取る（Row・Column・Rows.Count・Columns.Count の 4 往復を 1 に・2026-09-17）。"""
    class _Area:
        Address = "$C$13:$E$14"                             # win32com ではプロパティ（$ 付き）。呼び出しではない

        def __getattr__(self, name):                        # Row 等に触ったら落とす＝Address だけで済んだことを見る
            raise AssertionError(f"MergeArea.{name} に触った")

    class _Cell:
        MergeArea = _Area()

        @property
        def MergeCells(self):
            raise AssertionError("Address で済むのに MergeCells に触った")

    class _WS:
        def Rows(self, i):
            return type("Row", (), {"MergeCells": None})()

        def Columns(self, c):
            return type("Col", (), {"MergeCells": None})()

        def Cells(self, r, c):
            return _Cell()

    class _Rng:
        MergeCells = None
        Cells = type("Cs", (), {"Count": 6})()
        Worksheet = _WS()
        Row, Column = 13, 3
        Rows = type("R", (), {"Count": 2})()
        Columns = type("C", (), {"Count": 3})()
    areas, skipped = vm._merged_areas_in_range(_Rng())
    assert areas == ["C13:E14"] and skipped is None
    import vbam_view as _vw
    assert _vw._col_num_of("A") == 1 and _vw._col_num_of("Z") == 26 and _vw._col_num_of("AA") == 27


def test_any_excel_process_uses_last_pid_before_tasklist(monkeypatch):
    """前回掴んだ Excel の PID が生きていれば tasklist（36〜52 ミリ秒）を撃たない（2026-09-16）。"""
    import vbam_core as vc

    def _no_tasklist(*a, **k):
        raise AssertionError("PID が生きているのに tasklist を撃った")
    monkeypatch.setattr("subprocess.run", _no_tasklist)
    monkeypatch.setattr(vc, "_last_excel_pid", 4321)
    monkeypatch.setattr(vc, "_pid_is_excel", lambda pid: pid == 4321)
    assert vc._any_excel_process() is True
    # PID が死んでいれば従来どおり tasklist に聞く
    monkeypatch.setattr(vc, "_pid_is_excel", lambda pid: False)
    monkeypatch.setattr("subprocess.run",
                        lambda *a, **k: _FakeCompleted("情報: 条件に一致する実行中のタスクはありません。"))
    assert vc._any_excel_process() is False


def test_call_log_write_and_stats(monkeypatch, tmp_path, capsys):
    """呼び出し台帳: 1 行ずつ足し、stats がコマンド別・遅かった呼び出し・呼び出しの間を出す。"""
    import json
    import vbam_core as vc
    log = tmp_path / "calls.jsonl"
    monkeypatch.setattr(vc, "_CALL_LOG_FILE", str(log))
    monkeypatch.setattr(vc, "_last_book_name", "Book1.xlsx")
    vc.call_log_write("materials", _ap.Namespace(posargs=["テスト用1"]), 0.83, True, via="mcp")
    vc.call_log_write("write-cells", _ap.Namespace(posargs=["C7", "x" * 200, "D7"]), 0.12, True, via="mcp")
    vc.call_log_write("tidy", _ap.Namespace(posargs=["A1:D9"]), 6.2, False, via="mcp")
    rows = [json.loads(ln) for ln in log.read_text(encoding="utf-8").splitlines()]
    assert [r["cmd"] for r in rows] == ["materials", "write-cells", "tidy"]
    assert rows[0]["book"] == "Book1.xlsx" and rows[0]["via"] == "mcp" and rows[0]["sec"] == 0.83
    assert len(rows[1]["args"]) <= vc._CALL_LOG_ARGS_MAX + 1          # 引数は頭だけ
    assert rows[2]["ok"] is False
    assert vc.cmd_stats(_ap.Namespace(posargs=[], days=7, top=None, slow=None, json=False)) is True
    out = capsys.readouterr().out
    assert "3 回・失敗 1 回" in out and "materials" in out and "遅かった呼び出し" in out and "tidy A1:D9" in out
    assert vc.cmd_stats(_ap.Namespace(posargs=[], days=7, top=2, slow=1, json=True)) is True
    j = json.loads(capsys.readouterr().out)
    assert j["count"] == 3 and len(j["commands"]) == 2 and j["slowest"][0]["cmd"] == "tidy"
    # 記録が無ければその旨（失敗にしない）
    monkeypatch.setattr(vc, "_CALL_LOG_FILE", str(tmp_path / "none.jsonl"))
    assert vc.cmd_stats(_ap.Namespace(posargs=[], days=7, top=None, slow=None, json=False)) is True
    assert "記録がありません" in capsys.readouterr().out


def test_command_table_logs_outer_call_only(monkeypatch, tmp_path):
    """表の実装は台帳の包みつき。外側の 1 回だけ記録し、入れ子（batch／agent の中の手）は数えない。"""
    import json
    import pytest
    import vbam_core as vc
    log = tmp_path / "calls.jsonl"
    monkeypatch.setattr(vc, "_CALL_LOG_FILE", str(log))
    table = vm._command_table()
    assert vm.raw_command(table["rules"]) is vv.cmd_rules
    assert table["gate"] is table["関所"]                          # 別名は同じ包み
    ns = _ap.Namespace(posargs=[], json=False, command="rules")
    assert table["rules"](ns) is True
    rows = [json.loads(ln) for ln in log.read_text(encoding="utf-8").splitlines()]
    assert len(rows) == 1 and rows[0]["cmd"] == "rules" and rows[0]["ok"] is True and rows[0]["book"] is None

    # 入れ子: 包みの中から表を通して呼んでも 1 行のまま
    def _outer(args):
        return table["rules"](ns)
    vm._logged(_outer)(ns)
    rows = [json.loads(ln) for ln in log.read_text(encoding="utf-8").splitlines()]
    assert len(rows) == 2 and rows[1]["cmd"] == "rules"           # 外側の 1 回（台帳の名前は ns.command）

    # 失敗（例外）も ok=False で残る
    def _boom(args):
        raise RuntimeError("x")
    with pytest.raises(RuntimeError):
        vm._logged(_boom)(_ap.Namespace(posargs=[], command="boom"))
    rows = [json.loads(ln) for ln in log.read_text(encoding="utf-8").splitlines()]
    assert rows[-1]["cmd"] == "boom" and rows[-1]["ok"] is False


def test_parser_accepts_stats():
    p = vm.build_parser()
    ns, unknown = p.parse_known_args(["stats", "--days", "3", "--top", "5", "--slow", "2", "--json"])
    assert ns.command == "stats" and ns.days == 3 and ns.top == 5 and ns.slow == 2 and ns.json and not unknown
    assert "stats" in vm._command_table()


def test_next_block_is_continuation_only_when_it_covers_the_table():
    # 2026-09-17: 途中の空行 1 行の下は表の続き（同じ幅以上）・細い塊（合計欄・注記）は別の塊
    import vbam_edit as ve
    assert ve._next_block_is_continuation(7, 1, 7, 1, 7) is True
    assert ve._next_block_is_continuation(8, 1, 8, 1, 7) is True
    assert ve._next_block_is_continuation(2, 6, 7, 1, 7) is False
    assert ve._next_block_is_continuation(1, 1, 1, 1, 1) is False


def test_loose_key_levels_postal_mark_8digit_date_units_and_keeps_sign():
    # 2026-09-17: 重複を消した後の「消えた行」の照合が、書き方だけ違う値を別の値と読んで誤検知した
    import datetime as _dt
    from vbam_hands import _loose_key as k
    assert k('〒240-7550') == k('240-7550') == '2407550'
    assert k('20250526') == k(_dt.date(2025, 5, 26)) == k('2025/5/26')
    assert k('4422000円') == k(4422000) == k('4,422,000') == k('¥4,422,000')
    assert k('12個') == k(12)
    assert k('△1,000') == k('▲1000') == k('(1,000)') == k('-1,000') == k(-1000) == '-1000'
    assert k('△1,000') != k(1000)
    assert k('5%') == k(0.05) and k('100%') == k(1)
    assert k('03-1234-5678') == '0312345678' and k('99999999') == '99999999'   # 日付にならない 8 桁はそのまま


def test_lost_rows_reads_serial_dates_against_date_column():
    import datetime as _dt
    from vbam_hands import _lost_rows
    before = [['会社', '電話', '登録日', '取引額'],
              ['a社', '0187634853', '20250526', '4422000円'],
              ['b社', '0185 62 4228', 43795, 4229000],
              ['c社', '03-1111-2222', _dt.date(2020, 1, 2), 100]]
    after = [['会社', '電話', '登録日', '取引額'],
             ['a社', '0187-63-4853', _dt.datetime(2025, 5, 26), 4422000],
             ['b社', '0185-62-4228', _dt.date(1899, 12, 30) + _dt.timedelta(days=43795), 4229000],
             ['c社', '03-1111-2222', _dt.date(2020, 1, 2), 100]]
    assert _lost_rows(before, after, 0, 0) == []
    assert _lost_rows(before, after[:2] + after[3:], 0, 0) == [2]


def test_content_findings_read_group_named_subtotal_rows_as_subtotals():
    """小計の行は「小計」だけとは限らない（「一般会計 小計」）。区分名つきでも小計＝下の合計は二重に数えない。

    2026-09-18: B 集計の「小計行」を鍛えたとき、区分名つきの小計を本文の行と見なし、
    正解の表でも「合計が上の和と合わない」で止まった（登録しても毎回 AI に回る形）。
    """
    from vbam_edit import _content_findings
    vals = [["区分", "課", "金額"],
            ["一般会計", "総務課", 100],
            ["一般会計", "財政課", 200],
            ["一般会計 小計", None, 300],
            ["特別会計", "税務課", 400],
            ["特別会計 小計", None, 400],
            ["合計", None, 700]]
    block, _seen = _content_findings(vals, None, 1, 1)
    assert not [b for b in block if "上の和と合いません" in b]
    # 小計そのものが合わなければ、今までどおり止める
    vals[3][2] = 999
    block2, _s2 = _content_findings(vals, None, 1, 1)
    assert [b for b in block2 if "「小計」行の C4" in b]


def test_fired_ledger_counts_and_summary(tmp_path, monkeypatch):
    """2026-09-18 夜: 先撃ちで撃ったマクロの名前を残し、回数と「一度も撃たれていない本」を出す（引退の判断に使う）。"""
    import vbam_ledger as vl
    monkeypatch.setattr(vl, '_FIRED_FILE', str(tmp_path / 'f.jsonl'))
    vl.fired_append('r1', 'a.xlsx', 'S', '合計行を足して', ['表を整える', '合計行を足す'], ['秀コンボ.xlam'])
    vl.fired_append('r2', 'b.xlsx', 'S', '合計行', ['合計行を足す'])
    c = vl.fired_counts()
    assert c['合計行を足す'][0] == 2 and c['表を整える'][0] == 1
    txt = vl.fired_summary(['合計行を足す', '散布図を作る'])
    assert '2 回  合計行を足す' in txt and '一度も撃たれていない' in txt and '散布図を作る' in txt
    monkeypatch.setattr(vl, '_FIRED_FILE', str(tmp_path / 'none.jsonl'))
    assert '記録がありません' in vl.fired_summary()


# ----------------------------------------------------------------
# 2026-09-19 Gemini の実射の記録から: open が「開いたのに 30 秒待って失敗」と誤報した
# （「作業ファイル/test.xlsx」のように / で渡されたパスを、/ が混ざったまま Excel の FullName と比べていた）
# ----------------------------------------------------------------

def test_same_path_ignores_slash_and_case(tmp_path):
    import vbam_core
    f = tmp_path / 'sub' / 'Book.xlsx'
    f.parent.mkdir()
    f.write_bytes(b'x')
    win = str(f)
    assert vbam_core.same_path(win, win.replace('\\', '/'))
    assert vbam_core.same_path(win.upper(), win.lower())
    assert vbam_core.same_path(str(tmp_path / 'sub' / '..' / 'sub' / 'Book.xlsx'), win)
    assert not vbam_core.same_path(win, str(tmp_path / 'sub' / 'Other.xlsx'))
    assert not vbam_core.same_path('', win) and not vbam_core.same_path(win, None)


def test_smart_path_resolve_returns_backslash_path(tmp_path, monkeypatch):
    import vbam_core
    (tmp_path / 'work').mkdir()
    (tmp_path / 'work' / 'test_pivot.xlsx').write_bytes(b'x')
    (tmp_path / 'a' / 'b').mkdir(parents=True)
    # 今いる場所からは見つからず、SCRIPT_DIR の 2 つ上（＝tmp_path）で見つかる形にする＝直した枝を通す
    monkeypatch.setattr(vbam_core, 'SCRIPT_DIR', str(tmp_path / 'a' / 'b'))
    monkeypatch.chdir(tmp_path / 'a')
    got = vbam_core.smart_path_resolve('work/test_pivot.xlsx')
    assert got and '/' not in got, got
    assert vbam_core.same_path(got, str(tmp_path / 'work' / 'test_pivot.xlsx'))


def test_note_if_macro_free_book(capsys):
    import vbam_core

    class _WB:
        def __init__(self, name):
            self.Name = name

    assert vbam_core.note_if_macro_free_book(_WB('test_pivot.xlsx')) is True
    assert 'マクロを持てない形式' in capsys.readouterr().out and True
    assert vbam_core.note_if_macro_free_book(_WB('book.xlsm')) is False
    assert vbam_core.note_if_macro_free_book(_WB('addin.XLAM')) is False
    assert capsys.readouterr().out == ''


def test_new_core_helpers_are_exported():
    """vbam_core は __all__ を持つ＝足した関数を一覧に入れ忘れると、import * で使う側から見えない。
    2026-09-19: same_path を入れ忘れ、open の中の NameError が except に黙って飲まれて「既に開いています」を素通りした。"""
    import vbam_core
    import vbam_edit
    import vbam_vba
    for name in ('same_path', 'note_if_macro_free_book', 'smart_path_resolve'):
        assert name in vbam_core.__all__, name
    assert vbam_edit.same_path is vbam_core.same_path
    assert vbam_vba.note_if_macro_free_book is vbam_core.note_if_macro_free_book


def test_every_public_function_in_core_is_exported():
    """vbam_core の中で def された、下線で始まらない関数は全部 __all__ に入っていること（入れ忘れの見張り）。"""
    import inspect
    import vbam_core
    missing = [n for n, f in vars(vbam_core).items()
               if inspect.isfunction(f) and f.__module__ == 'vbam_core' and not n.startswith('_') and n not in vbam_core.__all__]
    assert not missing, missing
