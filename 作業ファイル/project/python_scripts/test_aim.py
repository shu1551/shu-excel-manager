# -*- coding: utf-8 -*-
"""test_aim.py — vbam_aim（地図で棚を当てる）の試験。Excel 非依存。"""
import datetime

import vbam_xray as vx
import vbam_aim as va
from test_xray import mk, BASIC

D = datetime.datetime


def _m(rows, **kw):
    return vx.analyze(mk(rows, **kw))


def test_no_symptom_skips():
    m = _m(BASIC)
    r = va.aim(m, '選んだ列の文字の日付を日付にする')
    assert r['verdict'] == 'skip'


def test_text_date_fires_on_that_column_only():
    rows = [['番号', '氏名', '日付', '金額'],
            [1.0, '山田', '2026/04/01', 1000.0],
            [2.0, '佐藤', '2026/04/02', 2500.0],
            [3.0, '鈴木', '2026/04/03', 300.0],
            ['合計', None, None, 3800.0]]
    r = va.aim(_m(rows), '選んだ列の文字の日付を日付にする')
    assert r['verdict'] == 'fire'
    assert r['select'] == ['C2:C4']              # 見出しと合計の行は外す


def test_unknown_shelf_is_silent():
    assert va.aim(_m(BASIC), '存在しない棚') is None
    assert not va.has_aim('存在しない棚')


def test_subtotal_inside_refuses_row_delete():
    rows = [['区分', '名称', '金額'],
            ['A', 'あ', 100.0],
            [None, None, None],
            ['A計', None, 100.0],
            ['B', 'い', 200.0],
            [None, None, None],
            ['B計', None, 200.0]]
    m = _m(rows)
    r = va.aim(m, '表の中の空行を削除して詰める')
    assert r['verdict'] in ('refuse', 'skip', 'unknown')      # 小計が途中にある表で空行を詰めて撃つことはない
    assert r['verdict'] != 'fire' or not m['blocks'][0]['safety'].get('subtotal_inside')


def test_verify_detects_lost_total_and_header():
    before = _m(BASIC)
    after = _m(BASIC[:-1])                                       # 合計行が消えた
    assert '合計・小計の行が減った' in va.verify(before, after, '何かの棚')
    assert va.verify(before, before, '何かの棚') == []


def test_all_aim_rows_have_known_kind():
    for name, a in va._load().items():
        assert a['kind'] in ('直す', '作る', '操作', '調べる'), name
        assert a['sym'], name


def test_blank_row_inside_table_is_found_not_split():
    rows = [['番号', '氏名', '日付', '金額', '倍額'],
            [1.0, 'あ', None, 200.0, 400.0],
            [2.0, 'い', D(2026, 4, 2), 300.0, 600.0],
            [3.0, 'う', D(2026, 4, 3), 400.0, 800.0],
            [4.0, 'え', None, 500.0, 1000.0],
            [None] * 5,
            [6.0, 'お', None, 700.0, 1400.0],
            ['合計', None, None, 2100.0, None]]
    m = _m(rows)
    assert len([b for b in m['blocks'] if b['kind'] == 'table']) == 1
    assert va.aim(m, '表の中の空行を削除して詰める')['verdict'] == 'fire'


def test_near_unknown_block_is_not_skip():
    rows = [['番号', '氏名', '金額'], [1.0, 'あ', 10.0], [2.0, 'い', 20.0], [3.0, 'う', 30.0], [None] * 3, [None, 'メモ', None]]
    m = _m(rows)
    r = va.aim(m, '選んだ列の文字の日付を日付にする')
    assert r['verdict'] in ('skip', 'unknown')
