# -*- coding: utf-8 -*-
"""vbam_cause（式と値から原因を決める）の試験。題はオフィス田中「ExcelのAI活用」で Copilot が外した例（2026-10-02）。"""
import datetime

import vbam_cause as vc

NA, DIV0, VALUE, REF, NAME = -2146826246, -2146826281, -2146826273, -2146826265, -2146826259


def _one(lines, addr):
    return [x for x in lines if x.startswith(addr + ':') or x.startswith(addr + ' ')]


def test_lookup_number_vs_text_number():
    # A2=102（数値）を D2:D4（文字の '101'…）で探して #N/A（田中さん 05「型が一致せず」）
    fa = [['', '=VLOOKUP(A2,$D$2:$E$4,2,FALSE)', '', '', ''] for _ in range(1)]
    fa = [['番号', '名前', '', '番号', '名前'], ['', '=VLOOKUP(A2,$D$2:$E$4,2,FALSE)', '', '', ''],
          ['', '', '', '', ''], ['', '', '', '', '']]
    fv = [['番号', '名前', None, '番号', '名前'], [102.0, NA, None, '101', '田中'], [None, None, None, '102', '小原'],
          [None, None, None, '103', '佐倉']]
    out = vc.cause_notes(fa, fv, 1, 1, {}, set())
    assert any('B2' in x and '数値' in x and '文字の数字' in x for x in out)


def test_lookup_text_date_vs_serial():
    fa = [['', '=VLOOKUP(TEXT(A1,"yyyy/m/d"),$D$1:$E$2,2,FALSE)', '', '', ''], ['', '', '', '', '']]
    fv = [['2025/8/1(金)', NA, None, datetime.datetime(2025, 8, 1), '朝礼'], [None, None, None, datetime.datetime(2025, 8, 2), '会議']]
    out = vc.cause_notes(fa, fv, 1, 1, {}, set())
    assert any('TEXT' in x and '日付' in x for x in out)


def test_match_and_rank_argument_order_and_map_lambda():
    fa = [['=INDEX(G2:G5,MATCH(H2:H5,J1,0))', '=RANK($G$2:$G$5,G2)', '=MAP(G2:G5,LAMBDA(a,b,a+b))']]
    fv = [[NA, NA, VALUE]]
    out = vc.cause_notes(fa, fv, 2, 10, {}, set())
    assert any('MATCH の引数の順が逆' in x for x in out)
    assert any('RANK の引数の順が逆' in x for x in out)
    assert any('LAMBDA の引数は 2 つ' in x for x in out)


def test_ref_and_missing_name_and_error_inside_range():
    fa = [['=VLOOKUP(P2,#REF!,2,FALSE)'], ['=VLOOKUP(P2,リスト,2,FALSE)'], ['=MIN(A1:A3)']]
    fv = [[REF], [NAME], [DIV0]]
    out = vc.cause_notes(fa, fv, 1, 2, {'Data': '=Sheet1!$A$1'}, set())
    assert any('#REF!' in x and '消した' in x for x in out)
    assert any('名前「リスト」が定義されていない' in x for x in out)
    fa2 = [[''], [''], [''], ['=MIN(A1:A3)']]
    fv2 = [[10.0], [DIV0], [30.0], [DIV0]]
    out2 = vc.cause_notes(fa2, fv2, 1, 1, {}, set())
    assert any('A2（#DIV/0!）がそのまま伝わっている' in x for x in out2)


def test_sum_of_text_numbers_is_zero():
    fa = [[''], [''], [''], ['=SUM(A1:A3)']]
    fv = [['100'], ['200'], ['300'], [0.0]]
    out = vc.cause_notes(fa, fv, 1, 1, {}, set())
    assert any('文字の数字 3 個' in x for x in out)


def test_approx_lookup_on_unsorted_table_returns_another_row():
    # 18: 並んでいない表を FALSE 忘れの VLOOKUP で探す＝田中は A のはずが F、久保は表にあるのに #N/A
    fa = [['', '', '', ''], ['', '', '', ''], ['', '', '', '=VLOOKUP(C2,$A$1:$B$3,2)'], ['', '', '', '=VLOOKUP(C3,$A$1:$B$3,2)']]
    fa = [['', '', '', ''], ['', '', '', '=VLOOKUP(C2,$A$1:$B$3,2)'], ['', '', '', '=VLOOKUP(C3,$A$1:$B$3,2)']]
    fv = [['田中', 'A', None, None], ['小原', 'B', '田中', 'C'], ['久保', 'F', '久保', NA]]
    out = vc.cause_notes(fa, fv, 1, 1, {}, set())
    assert any(x.startswith('D2') and '本当は「A」' in x for x in out)
    assert any(x.startswith('D3') and 'あるのに見つからない' in x for x in out)


def test_trailing_space_breaks_filter_comparison():
    fa = [['', '=FILTER(A1:B3,A1:A3="田中")'], ['', ''], ['', '']]
    fv = [['田中', '田中'], ['田中 ', None], ['小原', None]]
    out = vc.cause_notes(fa, fv, 1, 1, {}, set())
    assert any('A2 の値は前後に空白' in x for x in out)


def test_names_and_udfs_are_explained():
    out = vc.cause_notes([['=TEST(CHECK(J1))']], [['佐倉']], 1, 11, {'TEST': '=LAMBDA(a,1)'}, {'CHECK'})
    assert any('TEST は名前' in x for x in out) and any('CHECK は VBA の自作関数' in x for x in out)


def test_no_values_mode_hides_values():
    fa = [['番号', '名前', '', '番号', '名前'], ['', '=VLOOKUP(A2,$D$2:$E$3,2,FALSE)', '', '', ''], ['', '', '', '', '']]
    fv = [['番号', '名前', None, '番号', '名前'], [102.0, NA, None, '機密番号', '機密太郎'], [None, None, None, '103', '佐倉']]
    out = '\n'.join(vc.cause_notes(fa, fv, 1, 1, {}, set(), nv=True))
    assert 'B2' in out and '102' not in out and '機密' not in out


def test_spelling_variants_ascii_only():
    fv = [['地域'], ['Tokyo'], ['Tokyoo'], ['Tokyoo'], ['Osaka'], ['osaka'], ['斉藤'], ['齋藤']]
    out = vc.spelling_notes(fv, 1, 1)
    assert any('Tokyo' in x and 'Tokyoo' in x for x in out)
    assert any('osaka' in x and '大文字小文字だけ' in x for x in out)
    assert not any('斉藤' in x for x in out)               # 日本語の名前は揺れと言い切れない（田中さん）
