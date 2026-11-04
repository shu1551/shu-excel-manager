"""vbam_dock（点検の部品を 1 回で撃つ入口）の純 Python 部分のテスト。Excel は使わない。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vbam_dock as vd


def _data(nv=False, **kw):
    d = {'book': 'T.xlsx', 'format': 'xlsx', 'size': 1, 'props': [], 'totals': {'formulas': 3},
         'sheets': [{'name': 'S1', 'type': 'ワークシート', 'vis': '表示', 'protect': False, 'used': 'A1:C5'}],
         'names': [], 'tables': [], 'pivots': [], 'cf': [], 'dv': [], 'shape_counts': {}, 'links': [],
         'fml': [], 'funcs': {}, 'refs': set(), 'shapes': [], 'comments': [], 'queries': [], 'conns': [],
         'vba': [], 'ext_links': [], 'styles': [], 'numfmts': None}
    data = {'d': d, 'issues': [], 'score': 100, 'fragile': [], 'approx': [], 'causes': [], 'traces': [], 'skipped': []}
    data.update(kw)
    return data


def test_lookup_findings():
    f = [('S1', 'A1', '=VLOOKUP(B1,D:F,3,FALSE)'), ('S1', 'A2', '=VLOOKUP(B2,D:F,2)'), ('S2', 'A1', '=SUM(A1:A3)')]
    fragile, approx = vd.lookup_findings(f)
    assert {x[1] for x in fragile} == {'A1', 'A2'}
    assert [x[1] for x in approx] == ['A2']
    assert fragile[0][2] == 'VLOOKUP' and '3' in fragile[0][3]


def test_pick_totals_biggest_first():
    f = [('S', 'A1', '=SUM(B1:B3)'), ('S', 'A2', '=SUM(B1:B100)'), ('S', 'A3', '=B1+B2'),
         ('S', 'A4', '=SUBTOTAL(9,B1:B50)'), ('S', 'A5', '=SUM(B1:B5)')]
    got = vd.pick_totals(f, top=2)
    assert [x[1] for x in got] == ['A2', 'A4']


def test_summarize_prefixes_sheet_and_scores():
    iss = {'S1': [{'cell': 'A1', 'severity': 'critical', 'type': 'formula_error', 'msg': 'x'}], 'S2': []}
    allv, score = vd.summarize(iss)
    assert allv[0]['cell'] == 'S1!A1'
    assert score == 85


def test_markdown_headings_and_score():
    md = vd.render_markdown(_data())
    for h in ('## 概要と点数', '## 骨', '## 病巣', '## 壊れやすい式', '## エラーの原因', '## 式のつながり'):
        assert h in md
    assert '点数: 100 / 100' in md


def test_markdown_no_values_hides_messages():
    iss = [{'cell': 'S1!A1', 'severity': 'critical', 'type': 'omitted_sum_range', 'msg': 'B5 (123456) が漏れ'}]
    shown = vd.render_markdown(_data(issues=iss), nv=False)
    hidden = vd.render_markdown(_data(issues=iss), nv=True)
    assert '123456' in shown
    assert '123456' not in hidden and '集計の範囲漏れ' in hidden


def test_markdown_sections_filled():
    md = vd.render_markdown(_data(
        fragile=[('S1', 'A1', 'VLOOKUP', '列番号 3 を直書き')], approx=[('S1', 'A2', 'VLOOKUP', '検索方法を省略')],
        causes=[('S1', ['C3 が #REF!'])], traces=[('S1!A9  =SUM(A1:A8)', ['S1!A9  =SUM(A1:A8)'])]))
    assert 'S1!A1  VLOOKUP' in md and '近似一致' in md and '[S1] C3 が #REF!' in md and '### S1!A9' in md


def test_markdown_does_not_count_approx_twice_20261105():
    """近似一致は「壊れやすい式」の節に出し、病巣の表には重ねない（Enron の公開ブック 60 本で 895 件と 1,458 件の二重）。
    点数は diagnose と同じ（全部の issue で数えた data['score'] のまま）。"""
    iss = [{'cell': 'S1!A2', 'severity': 'warning', 'type': 'approx_match', 'msg': 'VLOOKUP が近似一致'},
           {'cell': 'S1!B3', 'severity': 'warning', 'type': 'hardcoded_constant', 'msg': '固定値'}]
    md = vd.render_markdown(_data(issues=iss, score=90, approx=[('S1', 'A2', 'VLOOKUP', '検索方法を省略')]))
    table = md.split('## 病巣')[1].split('## 壊れやすい式')[0]
    assert 'S1!A2' not in table and 'S1!B3' in table and '近似一致の 1 セルは下の' in table
    assert '警告 1 件' in md and '点数: 90 / 100' in md
    assert 'S1!A2  VLOOKUP  検索方法を省略' in md
    # 壊れやすい式の節が空なら（近似一致を lookup が拾わなかったとき）、病巣の表に残す
    md2 = vd.render_markdown(_data(issues=iss[:1]))
    assert 'S1!A2' in md2.split('## 病巣')[1]


def test_groups_copies_of_the_same_formula_20261105():
    """同じ形の式のコピー（R1C1 が同じ）は 1 行にまとめ、件数と囲む範囲を見せる。値の種類は同じ列ごと。
    点数は変えない（Enron の公開ブック 60 本: 指摘の 7 割が同じ式のコピーの数え直しだった）。"""
    iss = [{'cell': f'S1!N{r}', 'severity': 'warning', 'type': 'hardcoded_constant',
            'msg': "数式内に固定値 '*16' が直書きされています", 'formula': f'=L{r}*16'} for r in range(10, 15)]
    iss += [{'cell': 'S1!B3', 'severity': 'warning', 'type': 'hardcoded_constant',
             'msg': "数式内に固定値 '*1.001' が直書きされています", 'formula': '=B2*1.001'}]
    iss += [{'cell': f'S1!C{r}', 'severity': 'info', 'type': 'hidden_whitespace',
             'msg': f"値 'NAME{r} ' に末尾空白が含まれています", 'value': 'x '} for r in range(2, 6)]
    r1c1 = {f'S1!N{r}': '=RC[-2]*16' for r in range(10, 15)}
    r1c1['S1!B3'] = '=R[-1]C*1.001'
    gs = vd.group_issues(iss, r1c1)
    assert [len(g['cells']) for g in gs] == [5, 1, 4]
    md = vd.render_markdown(_data(issues=iss, score=88, r1c1=r1c1))
    assert '10 件を、同じ形の式・同じ列ごとに 3 行にまとめています' in md
    assert 'S1!N10 ほか 4 か所（S1!N10:N14） | 5 |' in md
    assert 'S1!C2 ほか 3 か所（S1!C2:C5） | 4 |' in md
    assert '点数: 88 / 100' in md
    # 壊れやすい式も同じ形ごとに
    fr = [('S1', f'A{r}', 'HLOOKUP', '行番号 11 を直書き') for r in range(1, 4)]
    md2 = vd.render_markdown(_data(fragile=fr, r1c1={f'S1!A{r}': '=HLOOKUP(RC2,R3C1:R22C44,11)' for r in range(1, 4)}))
    assert 'S1!A1 ほか 2 か所（S1!A1:A3）  HLOOKUP  行番号 11 を直書き' in md2


def test_html_adds_dock_section():
    iss = [{'cell': 'S1!A1', 'severity': 'warning', 'type': 'hardcoded_constant', 'msg': 'm', 'formula': '=A1*3'}]
    data = _data(issues=iss, score=95)
    html = vd.render_html(data, vd.render_markdown(data))
    assert 'dock のまとめ' in html and '</body>' in html


def test_fix_of_and_dock_column_20261105():
    """診断の 1 件ごとに「直す手」を添える（棚で直せる物は shelf-run の名前と選ぶ範囲、人の判断は理由）。
    dock と seiri は同じ対応表（vbam_audit.fix_of）を引く。"""
    import vbam_audit as va
    assert va.fix_of({'type': 'inconsistent_formula'})['name'] == va.SHELF_FILL_FORMULA
    assert va.fix_of({'type': 'omitted_sum_range'}, total_row=True)['name'] == va.SHELF_ADD_TOTAL
    assert va.fix_of({'type': 'omitted_sum_range'}, total_row=False)['kind'] == 'human'
    assert va.fix_of({'type': 'hidden_whitespace'})['name'] == 'diagnose --fix'
    for t in ('formula_error', 'hardcoded_constant', 'approx_match', 'outlier_value'):
        assert va.fix_of({'type': t})['kind'] == 'human'
    iss = [{'cell': f'S1!G{r}', 'severity': 'warning', 'type': 'inconsistent_formula', 'msg': '同列の標準数式パターンと異なります',
            'formula': '=E4+F4'} for r in (4, 9)]
    iss += [{'cell': 'S1!B3', 'severity': 'warning', 'type': 'hardcoded_constant', 'msg': "固定値 '*1.001'", 'formula': '=B2*1.001'}]
    md = vd.render_markdown(_data(issues=iss, score=85))
    assert '| 直す手 |' in md
    assert 'shelf-run 選んだ列の途切れた式を戻す --select G4:G9 --sheet S1' in md
    assert '人: 変わり得る数なら条件のセルに出して参照する' in md
    assert '直す手: 棚で直せる 2 件・道具で直せる 0 件・人の判断 1 件。' in md


def test_dock_and_seiri_name_the_same_shelf_20261105():
    """同じ表で、dock の対応表が棚と言う物を seiri の先撃ちも同じ棚で撃つ（言い方を 2 か所で変えない）。"""
    import vbam_audit as va
    import vbam_view as vv
    fa = [["品名", "数量", "単価", "金額"], ["a", 2, 100, "=B2*C2"], ["b", 3, 200, "=B3*C3"], ["c", 1, 300, 999],
          ["d", 4, 100, "=B5*C5"], ["e", 5, 100, "=B6*C6"], ["合計", "", "", "=SUM(D2:D6)"]]
    fr = [["品名", "数量", "単価", "金額"]] + [[r[0], r[1], r[2], "=RC[-2]*RC[-1]" if str(r[3]).startswith("=") else r[3]]
                                         for r in fa[1:6]] + [["合計", "", "", "=SUM(R[-5]C:R[-1]C)"]]
    vals = [["品名", "数量", "単価", "金額"], ["a", 2, 100, 200], ["b", 3, 200, 600], ["c", 1, 300, 999],
            ["d", 4, 100, 400], ["e", 5, 100, 500], ["合計", "", "", 2699]]
    issues = va.check_formula_linter(fa, fr, vals, 0, 0)
    dock_shelves = {va.fix_of(i)['name'] for i in issues if va.fix_of(i)['kind'] == 'shelf'}
    seiri_shelves = {h[1] for h in vv.seiri_formula_hints(fa, fr, vals, 1, 1, 0)}
    assert dock_shelves and dock_shelves <= seiri_shelves
