# -*- coding: utf-8 -*-
"""vbam_csvrisk（CSV を Excel で開くと壊れる列）のテスト（2026-10-01）。純 Python・COM なし。"""
import pytest

import vbam_csvrisk as cr


@pytest.mark.parametrize('v,kind', [
    ('007', 'leading_zero'), ('0123456', 'leading_zero'), ('0', None), ('0.5', None), ('10', None), ('100', None),
    ('1234567890123456', 'long_digits'), ('12345678901234567890', 'long_digits'),
    ('123456789012', 'sci_display'), ('12345678901', None),                       # 11 桁までは標準で数字のまま表示される
    ('1-2', 'date_like'), ('10/3', 'date_like'), ('JAN1', 'date_like'), ('Mar-3', 'date_like'), ('sept 1', 'date_like'),
    ('2026-01-02', None), ('2026/1/2', None), ('03-1234-5678', None),               # 本物の日付・電話番号は壊れの対象にしない
    ('12E5', 'sci_notation'), ('1e3', 'sci_notation'), ('E5', None),
    ('+81', 'plus'), ('+81.5', 'plus'), ('=SUM(A1)', 'formula'), ('', None), ('  ', None), ('りんご', None), ('A001', None),
])
def test_damage_kind(v, kind):
    assert cr.damage_kind(v) == kind


def _write(tmp_path, text, enc='utf-8'):
    p = tmp_path / 'a.csv'
    p.write_bytes(text.encode(enc))
    return str(p)


def test_scan_finds_only_risky_columns_and_counts_examples(tmp_path):
    p = _write(tmp_path, "コード,品名,数量,型番,ID\n007,りんご,10,1-2,1234567890123456\n008,みかん,20,3-4,1234567890123457\n10,かぶ,30,AB,5\n")
    s = cr.scan_csv(p)
    assert sorted(s['cols']) == [1, 4, 5]
    assert s['cols'][1]['kinds'] == {'leading_zero': 2} and s['cols'][1]['examples'] == ['007', '008']
    assert s['cols'][4]['kinds'] == {'date_like': 2} and s['cols'][5]['kinds'] == {'long_digits': 2}
    assert s['rows'] == 3 and s['ncols'] == 5 and s['delimiter'] == ','
    assert cr.risky_field_info(s) == [(1, 2), (2, 1), (3, 1), (4, 2), (5, 2)]


def test_header_row_is_not_counted_and_clean_file_has_no_risk(tmp_path):
    p = _write(tmp_path, "007,1-2\nあ,い\n")                      # 1 行目（見出し）に危ない形があっても数えない
    assert cr.scan_csv(p)['cols'] == {}
    assert cr.describe_risks(cr.scan_csv(p)) == []


def test_encoding_and_delimiter_detection(tmp_path):
    s = cr.scan_csv(_write(tmp_path, "コード\t名前\n007\tあ\n", 'cp932'))
    assert s['encoding'] == 'cp932' and s['origin'] == 932 and s['delimiter'] == '\t' and sorted(s['cols']) == [1]
    s = cr.scan_csv(_write(tmp_path, "コード;名前\n007;あ\n", 'utf-8-sig'))
    assert s['encoding'] == 'utf-8-sig' and s['origin'] == 65001 and s['delimiter'] == ';'
    s = cr.scan_csv(_write(tmp_path, "コード,名前\n007,あ\n", 'utf-8'))
    assert s['origin'] == 65001


def test_quoted_commas_do_not_shift_columns(tmp_path):
    p = _write(tmp_path, 'a,b,c\n"x,y",007,1-2\n')
    assert sorted(cr.scan_csv(p)['cols']) == [2, 3]


def test_describe_risks_names_the_columns_and_how_to_opt_out(tmp_path):
    p = _write(tmp_path, "コード,型番\n007,1-2\n")
    text = '\n'.join(cr.describe_risks(cr.scan_csv(p)))
    assert '2 列を文字列として取り込みます' in text and '1 列目「コード」' in text and '先頭の 0 が消える' in text
    assert '2 列目「型番」' in text and '--excel-default' in text
