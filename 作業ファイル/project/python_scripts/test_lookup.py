# -*- coding: utf-8 -*-
"""vbam_lookup（列・行の挿入／削除で VLOOKUP などの直書きの番号を直す）のテスト（2026-10-01）。純 Python・COM なし。"""
import vbam_lookup as vl


def _plan(f, op, start, count, axis='col', own='売上', target='売上'):
    return vl.plan_formula(f, own, target, axis, op, start, count)


def test_scan_vlookup_hlookup_index_literals():
    items = vl.scan_lookups('=VLOOKUP(A2,$I$2:$K$60,3,FALSE)+HLOOKUP(B1,A1:F3,2,0)+INDEX(A1:D9,5,2)')
    assert [(i['fn'], i['kind'], i['val']) for i in items] == [('VLOOKUP', 'col', 3), ('HLOOKUP', 'row', 2),
                                                               ('INDEX', 'row', 5), ('INDEX', 'col', 2)]
    assert (items[0]['c1'], items[0]['c2'], items[0]['r1'], items[0]['r2']) == (9, 11, 2, 60)


def test_scan_ignores_non_literal_and_unreadable():
    # 番号がセル・MATCH・名前・INDIRECT の式は直書きではない／範囲が読めない
    assert vl.scan_lookups('=VLOOKUP(A2,$I$2:$K$60,MATCH("金額",$I$1:$K$1,0),0)') == []
    assert vl.scan_lookups('=VLOOKUP(A2,マスタ,3,0)') == []
    assert vl.scan_lookups('=VLOOKUP(A2,INDIRECT("A1:C9"),3,0)') == []
    assert vl.scan_lookups('=INDEX(A1:D9,0,2)')[0]['kind'] == 'col'                 # 0（まるごと）は数えない
    assert vl.scan_lookups('=IF(A1="VLOOKUP(x,A:B,2,0)",1,2)') == []                # 文字列の中は式ではない


def test_scan_sheet_prefix_and_quotes_and_whole_columns():
    it = vl.scan_lookups("=VLOOKUP(A2,'売上 2026'!$A:$D,4,FALSE)")[0]
    assert it['sheet'] == '売上 2026' and (it['c1'], it['c2'], it['r1']) == (1, 4, None)
    it = vl.scan_lookups("=VLOOKUP(A2,マスタ!A2:C9,2,0)")[0]
    assert it['sheet'] == 'マスタ'
    # シート名の中のコンマ・かっこで引数が切れない
    it = vl.scan_lookups("=VLOOKUP(A2,'a,(b)'!A1:C9,2,0)")[0]
    assert it['sheet'] == 'a,(b)' and it['val'] == 2


def test_index_with_two_args_on_a_single_row_is_a_column():
    it = vl.scan_lookups('=INDEX(A1:F1,3)')[0]
    assert it['kind'] == 'col'
    it = vl.scan_lookups('=INDEX(A1:A9,3)')[0]
    assert it['kind'] == 'row'


def test_insert_column_inside_range_before_target_bumps_the_number():
    f = '=VLOOKUP(A2,$I$2:$K$60,3,FALSE)'            # 表は I〜K 列（9〜11）・3 番目＝K 列
    assert _plan(f, 'insert', 10, 1)[0] == {0: (3, 4)}   # J 列に挿入 → K は L に動くので 4
    assert _plan(f, 'insert', 11, 1)[0] == {0: (3, 4)}   # K 列に挿入（K は L へ）
    assert _plan(f, 'insert', 12, 1)[0] == {}             # L 列に挿入（表の右の外）→ そのまま
    assert _plan(f, 'insert', 9, 1)[0] == {}              # I 列に挿入 → 表ごと右へずれて番号はそのまま
    assert _plan(f, 'insert', 3, 2)[0] == {}              # 表より左 → 表ごとずれる
    assert _plan(f, 'insert', 10, 2)[0] == {0: (3, 5)}    # 2 列まとめて


def test_insert_after_target_column_inside_range_keeps_the_number():
    f = '=VLOOKUP(A2,$I$2:$L$60,2,FALSE)'            # 指すのは J 列（10）。K 列（11）に入れても J は動かない
    assert _plan(f, 'insert', 11, 1)[0] == {}
    assert _plan(f, 'insert', 10, 1)[0] == {0: (2, 3)}    # J に入れると J は K へ


def test_delete_column_before_target_reduces_the_number_and_target_deleted_breaks():
    f = '=VLOOKUP(A2,$I$2:$L$60,3,FALSE)'            # I〜L・3 番目＝K（11）
    assert _plan(f, 'delete', 9, 1) == ({0: (3, 2)}, 0)   # I を消すと K は J へ
    assert _plan(f, 'delete', 12, 1) == ({}, 0)            # L を消す（指す列より右）
    assert _plan(f, 'delete', 11, 1) == ({}, 1)            # K そのものを消す → 直せない
    assert _plan(f, 'delete', 9, 2) == ({0: (3, 1)}, 0)   # I・J を消す


def test_row_axis_for_hlookup_and_index():
    f = '=HLOOKUP(B1,$A$1:$F$5,3,0)'                 # 1〜5 行・3 番目＝3 行目
    assert _plan(f, 'insert', 2, 1, axis='row')[0] == {0: (3, 4)}
    assert _plan(f, 'insert', 4, 1, axis='row')[0] == {}
    f = '=INDEX(A1:D9,5,2)'
    assert _plan(f, 'insert', 3, 1, axis='row')[0] == {0: (5, 6)}                      # 行番号 5
    assert _plan(f, 'insert', 3, 1, axis='col')[0] == {} and _plan(f, 'insert', 2, 1, axis='col')[0] == {1: (2, 3)}


def test_other_sheet_is_matched_by_prefix_not_by_own_sheet():
    f = '=VLOOKUP(A2,マスタ!$A$2:$D$9,3,0)'
    assert vl.plan_formula(f, '売上', 'マスタ', 'col', 'insert', 2, 1)[0] == {0: (3, 4)}
    assert vl.plan_formula(f, '売上', '売上', 'col', 'insert', 2, 1)[0] == {}          # 売上に挿入しても影響しない
    f2 = '=VLOOKUP(A2,$A$2:$D$9,3,0)'                                                   # シート名なし＝その式のあるシート
    assert vl.plan_formula(f2, '売上', '売上', 'col', 'insert', 2, 1)[0] == {0: (3, 4)}
    assert vl.plan_formula(f2, '集計', '売上', 'col', 'insert', 2, 1)[0] == {}


def test_apply_values_rewrites_only_the_number_and_handles_many_calls():
    f = '=IFERROR(VLOOKUP(A2,$I$2:$K$60,3,FALSE),"なし")&VLOOKUP(B2,$I$2:$K$60, 2 ,0)&"VLOOKUP(x,A:B,2,0)"'
    changes, _ = _plan(f, 'insert', 10, 1)
    assert changes == {0: (3, 4), 1: (2, 3)}
    new = vl.apply_values(f, {k: v[1] for k, v in changes.items()})
    assert new == '=IFERROR(VLOOKUP(A2,$I$2:$K$60,4,FALSE),"なし")&VLOOKUP(B2,$I$2:$K$60, 3 ,0)&"VLOOKUP(x,A:B,2,0)"'


def test_nested_lookups_and_insert_into_the_expanded_range():
    # Excel は挿入で範囲を $I$2:$L$60 に伸ばす。そのあとの式でも番号の位置を取り違えない
    f = '=VLOOKUP(VLOOKUP(A2,$I$2:$K$60,2,0),$I$2:$K$60,3,0)'
    changes, _ = _plan(f, 'insert', 10, 1)
    assert changes == {0: (2, 3), 1: (3, 4)}                                           # 番号の位置順（内側が先）
    new = vl.apply_values(f, {k: v[1] for k, v in changes.items()})
    assert new == '=VLOOKUP(VLOOKUP(A2,$I$2:$K$60,3,0),$I$2:$K$60,4,0)'


def test_report_lines_wording():
    out = vl.report_lines([('売上', 'F2', '=VLOOKUP(A2,I:K,3,0)', '=VLOOKUP(A2,I:L,4,0)')], [('売上', 'G2', 'x')], 'J列の挿入')
    text = '\n'.join(out)
    assert '1 式直しました' in text and '--keep-lookup' in text and '売上!F2' in text and '直せなかった式 1 件' in text


class _Cell:
    def __init__(self, formula):
        self.Formula2 = formula
        self.HasArray = False


def test_apply_plan_uses_moved_cell_position_and_marks_vanished_ones():
    # 売上 シートの J 列（10）に 1 列挿入: F2 の式はそのまま・K2（11）の式は L2（12）に動いている
    class _WS:
        def __init__(self):
            self.cells = {(2, 6): _Cell('=VLOOKUP(A2,$I$2:$L$60,3,0)'), (2, 12): _Cell('=VLOOKUP(A2,$I$2:$L$60,3,0)')}

        def Cells(self, r, c):
            return self.cells[(r, c)]

    class _WB:
        def __init__(self):
            self.ws = _WS()

        def Worksheets(self, name):
            return self.ws

    wb = _WB()
    plan = [{'sheet': '売上', 'row': 2, 'col': 6, 'changes': {0: (3, 4)}, 'broken': 0, 'formula': ''},
            {'sheet': '売上', 'row': 2, 'col': 11, 'changes': {0: (3, 4)}, 'broken': 0, 'formula': ''}]
    fixed, skipped = vl.apply_plan(wb, plan, '売上', 'col', 'insert', 10, 1)
    assert [(a) for _s, a, _o, _n in fixed] == ['F2', 'L2']
    assert wb.ws.cells[(2, 6)].Formula2.endswith(',4,0)') and wb.ws.cells[(2, 12)].Formula2.endswith(',4,0)')
    assert skipped == []
