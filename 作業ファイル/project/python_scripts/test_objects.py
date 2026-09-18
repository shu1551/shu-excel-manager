# -*- coding: utf-8 -*-
"""グラフ・ピボット・テーブルの姿の比べ方（vbam_objects.compare_objects・純 Python）。COM で読む部分は fixtures の実射で確かめる。"""
import copy

import vbam_forge as vf
import vbam_objects as vo


def _chart(**kw):
    c = {'sheet': '課別支出', 'chartsheet': False, 'type': 51, 'title': '課別支出額', 'legend': 'なし', 'font': 'Meiryo UI',
         'gap': 80, 'axes': {'数値軸': {'title': '', 'format': '#,##0', 'gridlines': False}},
         'series': [{'name': '支出額', 'values': [100, 200], 'categories': ['総務課', '経理課'], 'labels': '値', 'color': '#1F4E79'}]}
    c.update(kw)
    return c


def test_same_objects_pass_and_old_ledger_is_skipped():
    e = {'charts': [_chart()], 'pivots': [], 'tables': []}
    assert vo.compare_objects(e, copy.deepcopy(e)) == []
    assert vo.compare_objects(None, e) == []


def test_chart_differences_are_named():
    e = {'charts': [_chart()], 'pivots': []}
    g = copy.deepcopy(e)
    g['charts'][0].update({'type': 4, 'legend': '右'})
    g['charts'][0]['series'][0].update({'values': [100, 200, 300], 'color': '#4472C4'})
    out = vo.compare_objects(e, g)
    assert any('種類 期待 集合縦棒 ／ マクロ後 折れ線' in x for x in out)
    assert any("凡例 期待 'なし'" in x for x in out)
    assert any('値: 期待 [100, 200]' in x for x in out) and any('塗りの色' in x for x in out)
    out = vo.compare_objects(e, {'charts': [], 'pivots': []})
    assert out and out[0].startswith('グラフの数: 期待 1 ／ マクロ後 0（グラフを作っていない')


def test_pivot_and_table_differences():
    p = {'sheet': 'ピボット', 'rows': ['課名'], 'cols': ['科目'], 'pages': [], 'row_grand': True, 'col_grand': True,
         'data': [{'source': '支出額', 'func': '合計', 'format': '#,##0'}], 'numbers': [1, 2, 3], 'labels': ['総務課']}
    t = {'sheet': '名簿', 'name': 'T名簿', 'range': 'A1:C9', 'style': 'TableStyleMedium2', 'totals': True, 'filter': True,
         'stripes': True, 'columns': [{'name': '数量', 'formula': '', 'total': '合計'},
                                      {'name': '金額', 'formula': '=[@数量]*[@単価]', 'total': '合計'}]}
    e = {'charts': [], 'pivots': [p], 'tables': [t]}
    g = copy.deepcopy(e)
    g['pivots'][0].update({'rows': ['科目'], 'numbers': [1, 2]})
    g['tables'][0]['columns'][1]['formula'] = '=C2*D2'
    g['tables'][0]['name'] = 'テーブル1'
    out = vo.compare_objects(e, g)
    assert any('行のフィールド' in x for x in out) and any('出てきた数が違う' in x and '[3]' in x for x in out)
    assert any("名前 期待 'T名簿'" in x for x in out) and any('計算列の式' in x for x in out)


def test_forge_compare_includes_objects_only_when_expect_has_them():
    base = {'row': 1, 'col': 1, 'values': [['a']]}
    e = dict(base, objects={'charts': [_chart()], 'pivots': [], 'tables': []})
    assert vf._compare(e, dict(base, objects={'charts': [], 'pivots': [], 'tables': []}))
    assert vf._compare(base, dict(base, objects={'charts': [_chart()], 'pivots': [], 'tables': []})) == []


def test_bgr_to_rgb():
    assert vo._bgr_to_rgb(0x794E1F) == 0x1F4E79


def test_gap_width_only_for_bar_charts():
    pie = _chart(type=5, gap=100)
    got = copy.deepcopy(pie)
    got['gap'] = 150
    assert vo.compare_objects({'charts': [pie]}, {'charts': [got]}) == []
    bar_got = _chart(gap=150)
    assert any('GapWidth' in x for x in vo.compare_objects({'charts': [_chart()]}, {'charts': [bar_got]}))


def test_bar_colors_compare_by_look():
    e = _chart()
    e['series'][0].update({'values': [3, 2, 1], 'color': '#1F4E79', 'point_colors': ['#C00000', '#1F4E79', '#1F4E79']})
    g = copy.deepcopy(e)
    g['series'][0].update({'color': '#4F81BD', 'point_colors': ['#C00000', '#1F4E79', '#1F4E79']})
    assert not any('色' in x for x in vo.compare_objects({'charts': [e]}, {'charts': [g]}))
    g['series'][0]['point_colors'] = []
    assert any('棒の色' in x for x in vo.compare_objects({'charts': [e]}, {'charts': [g]}))


def test_pivot_design_differences():
    """2026-09-18: いつもの月別ピボット＝スタイル・表形式・ラベルの繰り返し・小計・空白に 0。古い台帳（キー無し）は比べない。"""
    base = {'sheet': '課別月別', 'rows': ['課名', '科目'], 'cols': ['月 (支出日)'], 'pages': [], 'data': [], 'row_grand': True,
            'col_grand': True, 'numbers': [1], 'labels': ['4月']}
    rf = [{'field': '課名', 'layout': '表形式', 'repeat': True, 'subtotal': False},
          {'field': '科目', 'layout': '表形式', 'repeat': False, 'subtotal': True}]
    e = {'charts': [], 'pivots': [dict(base, style='PivotStyleMedium9', null='0', row_fields=rf)]}
    g = {'charts': [], 'pivots': [dict(base, style='PivotStyleLight16', null='',
                                       row_fields=[dict(rf[0], layout='コンパクト', repeat=False, subtotal=True),
                                                   dict(rf[1], repeat=True, subtotal=False, layout='コンパクト')])]}
    out = vo.compare_objects(e, g)
    assert any('スタイル' in x for x in out) and any('空白のセル' in x for x in out)
    assert any('「課名」のレイアウト' in x for x in out) and any('「課名」のラベルの繰り返し' in x for x in out)
    assert any('「課名」の小計' in x for x in out)
    assert not any('「科目」の小計' in x or '「科目」のラベル' in x for x in out)      # いちばん内側は見た目に出ない
    old = {'charts': [], 'pivots': [base]}
    assert vo.compare_objects(old, g) == []

def test_sparklines_compare_per_cell_and_old_ledger_skips():
    """2026-09-18 第二期: スパークラインの列。1 本ずつ（置いたセル・元の範囲・種類・印・色）で比べる。キーの無い古い台帳は比べない。"""
    sp = {'sheet': '月別', 'type': '折れ線', 'high': True, 'low': False, 'markers': False, 'color': '#1F4E79'}
    e = {'charts': [], 'pivots': [], 'sparklines': [dict(sp, cell='O2', source='B2:M2'), dict(sp, cell='O3', source='B3:M3')]}
    assert vo.compare_objects(e, copy.deepcopy(e)) == []
    none = {'charts': [], 'pivots': []}
    out = vo.compare_objects(e, none)
    assert any('スパークラインの数: 期待 2 本' in x and 'SparklineGroups.Add' in x for x in out)
    g = copy.deepcopy(e)
    g['sparklines'][1]['source'] = 'B3:N3'
    g['sparklines'][0]['high'] = False
    out = vo.compare_objects(e, g)
    assert any('O3: 元の範囲' in x for x in out) and any('O2: 最高点の印' in x for x in out)
    assert vo.compare_objects({'charts': [], 'pivots': []}, g) == []


def test_addr_drops_sheet_and_dollars():
    assert vo._addr("'月別'!$B$2:$M$2") == 'B2:M2'
    assert vo._addr('$O$2') == 'O2'


def test_bar_fill_hidden_is_compared_and_its_color_is_not():
    """2026-09-18 第二期: 増減の説明（積み上げ縦棒の土台は塗りなし）。塗りの有無は比べ、塗りの無い棒の色は比べない。"""
    base = {'sheet': 'S', 'chartsheet': False, 'type': 52, 'title': '', 'legend': 'なし', 'font': '', 'axes': {}}
    s0 = {'name': '土台', 'values': [0, 5], 'categories': ['a', 'b'], 'labels': 'なし', 'color': '#4472C4',
          'point_colors': [], 'type': 52, 'axis_group': 1, 'fill': False}
    e = {'charts': [dict(base, series=[s0])]}
    g = {'charts': [dict(base, series=[dict(s0, color='#FF0000')])]}
    assert vo.compare_objects(e, g) == []
    g2 = {'charts': [dict(base, series=[dict(s0, fill=True)])]}
    assert any('塗り 期待 なし' in x for x in vo.compare_objects(e, g2))


def test_uncomma_makes_formatted_numbers_plain():
    """2026-09-18 第二期: 散布図の XValues がセルの表示形式つき（'5,954,000'）で返った。数の文字だけ直し、ほかの文字は触らない。"""
    assert vo._s(vo._num(vo._uncomma('5,954,000'))) == '5954000'
    assert vo._uncomma('1,000円') == '1,000円' and vo._uncomma('総務課') == '総務課' and vo._uncomma(12) == 12


def test_pivots_compared_in_place_order():
    """2026-09-18 第二期: 同じシートのピボット 2 つが、触った後で PivotTables の並びが入れ替わっても取り違えない（置いた場所の順）。"""
    p1 = {'sheet': 'S', 'at': [1, 1], 'rows': ['課'], 'cols': [], 'pages': [], 'data': [], 'row_grand': True, 'col_grand': True,
          'numbers': [1], 'labels': ['a']}
    p2 = dict(p1, at=[1, 7], rows=['課'], cols=['科目'], numbers=[2], labels=['b'])
    assert vo.compare_objects({'charts': [], 'pivots': [p1, p2]}, {'charts': [], 'pivots': [p2, p1]}) == []
    old = {k: v for k, v in p1.items() if k != 'at'}
    assert vo.compare_objects({'charts': [], 'pivots': [old]}, {'charts': [], 'pivots': [p1]}) == []
