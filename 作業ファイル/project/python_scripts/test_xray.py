# -*- coding: utf-8 -*-
"""test_xray.py — vbam_xray の純 Python 判定の試験（Excel 非依存）。"""
import datetime
import time

import vbam_xray as vx


def mk(rows, r0=1, c0=1, formulas=None, merged=None, hidden_rows=None, hidden_cols=None, **kw):
    n = max(len(r) for r in rows)
    vals = [list(r) + [None] * (n - len(r)) for r in rows]
    sh = {'name': 't', 'r0': r0, 'c0': c0, 'nr': len(vals), 'nc': n, 'values': vals,
          'formulas': formulas or {}, 'merged': merged or [], 'hidden_rows': hidden_rows or [],
          'hidden_cols': hidden_cols or [], 'unknown': {}, 'objects': {}}
    sh.update(kw)
    return sh


def roles(m, b=0):
    out = {}
    for run in m['blocks'][b]['rows']:
        a, _, z = run['rows'].partition('-')
        for r in range(int(a), int(z or a) + 1):
            out[r] = run['role']
    return out


def sym(m, code):
    return [s for s in m['symptoms'] if s['code'] == code]


BASIC = [
    ['番号', '氏名', '日付', '金額'],
    [1.0, '山田 太郎', datetime.datetime(2026, 4, 1), 1000.0],
    [2.0, '佐藤 花子', datetime.datetime(2026, 4, 2), 2500.0],
    [3.0, '鈴木 一郎', datetime.datetime(2026, 4, 3), 300.0],
    ['合計', None, None, 3800.0],
]


def test_basic_table_roles_and_faces():
    m = vx.analyze(mk(BASIC))
    assert len(m['blocks']) == 1
    r = roles(m)
    assert r[1] == 'H' and r[2] == r[3] == r[4] == 'D' and r[5] == 'T'
    faces = {c['col']: c['face'] for c in m['blocks'][0]['columns']}
    assert faces == {'A': '番号', 'B': '氏名', 'C': '日付', 'D': '金額'}


def test_title_and_note_are_not_detail():
    rows = [['2026年度 売上表'], [], ['商品', '数量', '金額'], ['A', 1.0, 100.0], ['B', 2.0, 200.0], ['C', 3.0, 300.0], [], ['注 税込み']]
    m = vx.analyze(mk(rows))
    kinds = [b['kind'] for b in m['blocks']]
    assert kinds == ['memo', 'table', 'memo']
    assert set(roles(m, 0).values()) == {'N'}
    assert set(roles(m, 2).values()) == {'N'}


def test_total_in_the_middle_is_not_detail():
    rows = [['区', '件数'], ['A', 1.0], ['B', 2.0], ['小計', 3.0], ['C', 4.0], ['合計', 7.0]]
    m = vx.analyze(mk(rows))
    r = roles(m)
    assert r[4] == 'T' and r[6] == 'T' and r[2] == r[3] == r[5] == 'D'
    assert m['blocks'][0]['safety']['subtotal_inside'] is True


def test_two_tier_header_with_merge():
    rows = [['区分', '売上', None, '費用', None], [None, '前期', '当期', '前期', '当期'],
            ['A', 1.0, 2.0, 3.0, 4.0], ['B', 5.0, 6.0, 7.0, 8.0], ['C', 1.0, 1.0, 1.0, 1.0]]
    m = vx.analyze(mk(rows, merged=[(1, 2, 1, 3), (1, 4, 1, 5), (1, 1, 2, 1)]))
    r = roles(m)
    assert r[1] == r[2] == 'H' and r[3] == 'D'
    assert any(s['code'] == 'S16' for s in m['symptoms'])


def test_text_number_and_text_date_symptoms():
    rows = [['日付', '金額']] + [[f'2026/4/{d}', f'{d},000'] for d in range(1, 6)] + [[datetime.datetime(2026, 4, 9), 5000.0]]
    m = vx.analyze(mk(rows))
    codes = {(s['code'], s.get('col')) for s in m['symptoms']}
    assert ('S01', 'A') in codes and ('S02', 'B') in codes


def test_clean_table_has_no_symptoms():
    m = vx.analyze(mk(BASIC))
    assert not [s for s in m['symptoms'] if s['code'] in ('S01', 'S02', 'S03', 'S04', 'S05', 'S17', 'S18')]


def test_paren_inner_count_needs_a_note():
    rows = [['区分', '件数'], ['A', 10.0], ['（内数）', None], ['B', '(1)'], ['C', '(3)'], ['D', 5.0]]
    m = vx.analyze(mk(rows))
    pm = m['blocks'][0]['safety'].get('paren')
    assert pm and pm[0]['meaning'] == 'inner_count'
    rows2 = [r[:] for r in rows]
    rows2[2] = ['E', 6.0]
    m2 = vx.analyze(mk(rows2))
    pm2 = m2['blocks'][0]['safety'].get('paren')
    assert pm2 and pm2[0]['meaning'] == 'unknown' and pm2[0]['reason']


def test_indent_hierarchy_is_not_whitespace_symptom():
    rows = [['科目', '金額'], ['収入', 100.0], ['　税収', 60.0], ['　　市税', 50.0], ['　その他', 40.0], ['支出', 90.0]]
    m = vx.analyze(mk(rows))
    assert m['blocks'][0]['safety'].get('indent')
    assert not [s for s in m['symptoms'] if s['code'] == 'S03']


def test_hidden_column_is_marked():
    m = vx.analyze(mk(BASIC, hidden_cols=[(3, 3)]))
    assert m['blocks'][0]['safety']['hidden_cols'] == ['C']
    assert sym(m, 'S35')


def test_formula_column_break_and_sum_range():
    rows = [['品', '数', '単価', '金額'], ['a', 1.0, 10.0, 10.0], ['b', 2.0, 10.0, 20.0], ['c', 3.0, 10.0, 99.0], ['d', 4.0, 10.0, 40.0], ['合計', None, None, 70.0]]
    fm = {(2, 4): '=B2*C2', (3, 4): '=B3*C3', (5, 4): '=B5*C5', (6, 4): '=SUM(D2:D4)'}
    m = vx.analyze(mk(rows, formulas=fm))
    assert [s for s in sym(m, 'S17') if s.get('col') == 'D']
    assert [s for s in sym(m, 'S19') if s.get('col') == 'D']


def test_no_header_is_unknown_with_reason():
    rows = [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]
    m = vx.analyze(mk(rows))
    assert m['blocks'][0]['kind'] == 'unknown'
    assert m['unknown'] and all(u.get('reason') for u in m['unknown'])


def test_no_values_mode_hides_header_text():
    m = vx.analyze(mk(BASIC), nv=True)
    assert all('header' not in c for c in m['blocks'][0]['columns'])
    assert '山田' not in repr(m)


def test_speed_ordinary_sheet_under_100ms():
    rows = [['番号', '氏名', '日付', '金額', '備考']] + [[float(i), f'名{i}', datetime.datetime(2026, 4, 1 + i % 28), 100.0 * i, None] for i in range(1, 400)]
    t = time.perf_counter()
    vx.analyze(mk(rows))
    assert time.perf_counter() - t < 0.1


def test_neighbors_master_and_comparable():
    me = mk(BASIC)
    me['neighbors'] = [
        {'name': '対応表', 'nr': 5, 'nc': 2, 'complete': True,
         'values': [['コード', '名称'], ['A', 'あ'], ['B', 'い'], ['C', 'う'], ['D', 'え']]},
        {'name': '前月', 'nr': 5, 'nc': 4, 'complete': True,
         'values': [['番号', '氏名', '日付', '金額'], [1.0, 'x', None, 1.0], [2.0, 'y', None, 2.0], [3.0, 'z', None, 3.0], [4.0, 'w', None, 4.0]]},
        {'name': 'メモ', 'nr': 1, 'nc': 1, 'complete': True, 'values': [['ただの文章']]},
    ]
    m = vx.analyze(me)
    kinds = {n['sheet']: n['kind'] for n in m['neighbors']}
    assert kinds['対応表'] == 'master' and kinds['前月'] == 'comparable' and kinds['メモ'] == 'other'
    assert m['not_read']


def test_s18_only_for_formulas_error_values_are_listed_separately():
    rows = [['品', '率'], ['a', -2146826281], ['b', 1.0], ['c', 2.0], ['d', 3.0]]
    m = vx.analyze(mk(rows))
    assert not sym(m, 'S18') and m['error_cells']['#DIV/0!']['n'] == 1
    m2 = vx.analyze(mk(rows, formulas={(2, 2): '=1/0'}))
    assert sym(m2, 'S18')


def test_group_label_row_inside_table_is_G_and_blank_rows_between_details():
    rows = [['年', '件数', '金額'], [None, None, None], ['卸売業', None, None], [None, None, None],
            ['平成11年', 1.0, 10.0], ['平成12年', 2.0, 20.0], ['平成13年', 3.0, 30.0], [None, None, None],
            ['小売業', None, None], [None, None, None], ['平成11年', 4.0, 40.0], ['平成12年', 5.0, 50.0], ['平成13年', 6.0, 60.0]]
    m = vx.analyze(mk(rows))
    r = roles(m)
    assert r[5] == r[6] == 'D' and r[11] == 'D'
    assert r.get(3) in ('G', 'N', 'B', 'U') and r[3] != 'D'


def test_long_sheet_shortcut_is_fast_and_marks_sampling():
    rows = [['番号', '名称', '日付', '金額'] + [f'列{k}' for k in range(8)]] + [[float(i), f'名{i}', datetime.datetime(2026, 4, 1), 100.0] + [1.0] * 8 for i in range(1, 12000)]
    t = time.perf_counter()
    m = vx.analyze(mk(rows))
    assert time.perf_counter() - t < 0.5
    assert m['stats'].get('long') and any(u['item'] == 'symptoms' for u in m['unknown'])
    assert roles(m)[12000] == 'D'


def test_svg_renders_from_map():
    import vbam_xray_svg as sv
    sh = mk(BASIC, formulas={(5, 4): '=SUM(D2:D4)'})
    svg = sv.render_svg(vx.analyze(sh), sh)
    assert svg.startswith('<svg') and 'f4f7fb' in svg and '22c55e' in svg and '山田' not in svg


def test_total_label_variants():
    rows = [['区', '件数'], ['A区', 1.0], ['B区', 2.0], ['区部計', 3.0], ['一般会計', 9.0], ['市部計', 4.0], ['合計（１～２）', 7.0]]
    r = roles(vx.analyze(mk(rows)))
    assert r[4] == 'T' and r[6] == 'T' and r[7] == 'T' and r[5] == 'D'


def test_text_dates_in_all_text_row_do_not_split_table():
    rows = [['題名', '内容', '日付']] + [[f'記事{i}', f'内容{i}', datetime.datetime(2026, 4, i)] for i in range(1, 9)]
    rows[5][2] = '2026/04/05'                       # 文字の日付が入った、全部が文字の行
    m = vx.analyze(mk(rows))
    tables = [b for b in m['blocks'] if b['kind'] == 'table']
    assert len(tables) == 1 and tables[0]['range'] == 'A1:C9'
