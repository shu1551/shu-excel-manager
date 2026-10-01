# -*- coding: utf-8 -*-
"""structure（値を出さずにブックの構造だけ）と値なしの切り替え・check VBM020/021 の純 Python の部品（2026-10-01）。"""
import datetime

import vbam_structure as vs
import vbam_vba as vv


def test_fml_functions_ignores_string_literals_and_prefixes():
    f = '=IF(A1="SUM(x)",_xlfn.XLOOKUP(B1,C:C,D:D),VLOOKUP(A1,顧客!A1:E4,2,FALSE))'
    assert vs.fml_functions(f) == {'IF', 'XLOOKUP', 'VLOOKUP'}


def test_fml_sheet_refs_splits_sheets_and_books():
    r = vs.fml_sheet_refs("=顧客!A1+'月 次'!B2+[ext.xlsx]Sheet1!$A$1+\"x!y\"")
    assert r['sheets'] == {'顧客', '月 次'}
    assert r['books'] == {'[ext.xlsx]Sheet1'}


def test_value_kind_gives_type_not_value():
    assert vs.value_kind(98765432.0) == '数'
    assert vs.value_kind('機密太郎') == '文字'
    assert vs.value_kind(True) == '論理'
    assert vs.value_kind(None) == '空'
    assert vs.value_kind(datetime.datetime(2026, 1, 5)) == '日付'
    assert vs.value_kind(-2146826281) == '#DIV/0!'
    assert vs.value_kind(-2146826273) == '#VALUE!'
    assert vs.value_kind(-2146826252) == '#NUM!'


def test_numfmts_from_styles_xml_no_duplicates_and_unescaped():
    xml = ('<numFmts count="2"><numFmt numFmtId="176" formatCode="&quot;機密&quot;#,##0&quot;円&quot;"/>'
           '<numFmt formatCode="@&quot;様&quot;" numFmtId="177"/></numFmts>')
    assert vs.numfmts_from_styles_xml(xml) == [(176, '"機密"#,##0"円"'), (177, '@"様"')]


def test_mask_connection_hides_password():
    s = 'Provider=SQLOLEDB;Data Source=srv;User ID=a;Password=secret;Pwd="x;y"'
    m = vs.mask_connection(s)
    assert 'secret' not in m and 'Password=***' in m


def test_vba_procs_checklist_hits_and_comments_ignored():
    code = '\n'.join([
        'Option Explicit',
        'Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
        'Sub Auto_Open()',
        'End Sub',
        'Private Sub 古い()',
        '    Range("A" & i).Select',
        '    Selection.Value = 1',
        "    ' Selection はコメントの中なので数えない",
        '    On Error GoTo 終わり',
        '    If x Then GoTo 終わり',
        '    n = Cells.Count',
        '    s = "CreateObject("',
        '終わり:',
        'End Sub',
    ])
    head, procs = vs.vba_procs(code)
    assert head['option_explicit'] and head['declares'] == ['GetTickCount']
    assert [p['name'] for p in procs] == ['Auto_Open', '古い']
    assert procs[1]['scope'] == 'Private'
    labels = [h[1] for h in procs[1]['hits']]
    assert labels.count('Selection') == 1                     # コメントの中は数えない
    assert 'Range の中で文字列結合' in labels
    assert labels.count('GoTo（On Error を除く）') == 1        # On Error GoTo は数えない
    assert any(lb.startswith('Cells.Count') for lb in labels)
    assert 'CreateObject' not in labels                       # 文字列の中は数えない
    assert procs[0]['hits'][0][1].startswith('古い自動実行')


def _d(**kw):
    d = {'sheets': [], 'names': [], 'cf': [], 'dv': [], 'shapes': [], 'shape_counts': {}, 'conns': [], 'vba': [],
         'fml': [], 'ext_links': [], 'totals': {'formulas': 0, 'ext_cells': 0}}
    d.update(kw)
    return d


def test_points_carry_reason_and_where_to_check():
    d = _d(sheets=[{'name': '裏', 'vis': '完全に非表示（VBA からしか戻せない）', 'protect': False, 'unlocked': 0, 'hidden_fml': 0},
                   {'name': '入力', 'vis': '表示', 'protect': True, 'unlocked': 0, 'hidden_fml': 0}],
           names=[{'name': 'N壊れ', 'ref': '=#REF!', 'visible': True, 'scope': '(ブック)', 'error': True, 'system': False},
                  {'name': '_FilterDatabase', 'ref': '=A1', 'visible': False, 'scope': 'S', 'error': False, 'system': True}],
           dv=[{'sheet': 'S', 'addr': '1:1048576', 'cells': 17179869184, 'used': 10, 'whole': True, 'empty': True}],
           fml=['| 集計 | E1 | 1 | `=TODAY()` | 日付 | 揮発関数 TODAY・NOW |'])
    pts = vs.points(d)
    whats = [p[0] for p in pts]
    assert any('完全に非表示' in w for w in whats)
    assert any('ロック解除のセルが 0' in w for w in whats)
    assert any('N壊れ' in w for w in whats)
    assert not any('非表示の名前' in w for w in whats)          # Excel が作る _FilterDatabase は数えない
    assert any('中身の無い入力規則' in w for w in whats)
    assert not any('入力規則が列まるごと' in w for w in whats)  # 中身の無い規則は「残骸」の方で 1 回だけ言う
    assert any(w == '揮発関数（NOW・TODAY）' for w in whats)
    assert all(why and where for _w, why, where in pts)      # どの所見にも「なぜ」と「確かめる所」


def test_points_merge_lines_of_same_rule():
    proc = {'name': 'P', 'kind': 'Sub', 'scope': 'Public', 'lines': 5,
            'hits': [(2, '古いブック形式で保存（Excel 95/97 形式）', 3), (2, '古いブック形式で保存（Excel 95/97 形式）', 4)]}
    pts = vs.points(_d(vba=[{'module': 'M', 'type': '標準モジュール', 'lines': 9, 'head': {}, 'procs': [proc]}]))
    assert len(pts) == 1 and pts[0][2] == 'M の 3・4 行目'


def test_no_values_mode_env_overrides(monkeypatch, tmp_path):
    monkeypatch.setattr(vs, '_NV_FILE', str(tmp_path / 'no_values.flag'))
    monkeypatch.delenv('EXCEL_MANAGER_NO_VALUES', raising=False)
    assert vs.no_values_mode() is False
    (tmp_path / 'no_values.flag').write_text('on')
    assert vs.no_values_mode() is True
    monkeypatch.setenv('EXCEL_MANAGER_NO_VALUES', '0')
    assert vs.no_values_mode() is False
    monkeypatch.setenv('EXCEL_MANAGER_NO_VALUES', '1')
    assert vs.no_values_refusal('read-range')
    assert vs.no_values_refusal('materials') is None          # materials・structure・seiri は値なしで動く
    assert vs.no_values_refusal('structure') is None


def test_changes_table_hides_values_in_no_values_mode(monkeypatch):
    import vbam_undo
    rows = [{'sheet': 'S', 'addr': 'A2', 'before': '機密太郎', 'after': '機密 太郎'}]
    monkeypatch.setenv('EXCEL_MANAGER_NO_VALUES', '1')
    out = '\n'.join(vbam_undo._changes_table(rows))
    assert 'A2' in out and '機密' not in out


def test_check_vbm020_legacy_rules():
    lines = ['n = Cells.Count', 'm = rng.Cells.Count', 'k = ActiveSheet.Cells.CountLarge',
             'Set f = Application.FileSearch', 'x.SaveAs "a", FileFormat:=43', 'x.SaveAs "a", FileFormat:=56',
             "' FileSearch はコメント", 's = "FileSearch"', 'c = ActiveSheet.Cells.Count']
    got = [ln for ln, _why, _t in vv._diag_legacy(lines)]
    assert got == [1, 4, 5, 9]


def test_check_vbm021_auto_open():
    assert vv._diag_auto_open([{'name': 'Auto_Open'}, {'name': 'auto_close'}, {'name': 'Main'}]) == ['Auto_Open', 'auto_close']


def test_vbm020_021_in_rules_table():
    ids = [r[0] for r in vv._CHECK_RULES]
    assert 'VBM020' in ids and 'VBM021' in ids
