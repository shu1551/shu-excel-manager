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


def test_html_adds_dock_section():
    iss = [{'cell': 'S1!A1', 'severity': 'warning', 'type': 'hardcoded_constant', 'msg': 'm', 'formula': '=A1*3'}]
    data = _data(issues=iss, score=95)
    html = vd.render_html(data, vd.render_markdown(data))
    assert 'dock のまとめ' in html and '</body>' in html
