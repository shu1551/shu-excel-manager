# -*- coding: utf-8 -*-
"""書込後の知らせ（入力規則に合わない値・#SPILL! の原因）のテスト（2026-10-01）。COM なし（差し替えの偽物で）。"""
import vbam_edit as ve


class _List(list):
    @property
    def CountLarge(self):
        return len(self)


class _Val:
    def __init__(self, type_, f1='', f2='', ok=True):
        self.Type, self.Formula1, self.Formula2, self.Value = type_, f1, f2, ok


class _Cell:
    def __init__(self, addr, text='', value=None, val=None):
        self.Address = '$' + addr[0] + '$' + addr[1:]
        self.Text, self.Value, self._val = text, value, val
        self.Row, self.Column = int(addr[1:]), ord(addr[0]) - 64
        self.CountLarge = 1

    @property
    def Validation(self):
        if self._val is None:
            raise RuntimeError('規則なし')
        return self._val

    @property
    def Cells(self):
        return _List([self])


class _Many:
    """複数セルの範囲（SpecialCells(-4174) で規則のあるセルだけ返す）。"""
    def __init__(self, cells):
        self._cells = cells
        self.CountLarge = len(cells)

    @property
    def Cells(self):
        return _List(self._cells)

    def SpecialCells(self, kind, *a):
        assert kind == -4174
        got = [c for c in self._cells if c._val is not None]
        if not got:
            raise RuntimeError('該当なし')
        return _Many(got)


def test_rule_text_for_list_and_range():
    assert ve._validation_rule_text(_Val(3, '赤,青')) == 'リスト: 赤,青'
    assert ve._validation_rule_text(_Val(1, '1', '10')) == '整数: 1〜10'
    assert ve._validation_rule_text(_Val(6, '5')) == '文字数: 5'


def test_validation_violations_lists_only_cells_that_break_the_rule():
    cells = [_Cell('C2', '赤', val=_Val(3, '赤,青', ok=True)),
             _Cell('C3', '緑', val=_Val(3, '赤,青', ok=False)),
             _Cell('D3', '99', val=_Val(1, '1', '10', ok=False)),
             _Cell('E3', 'x')]                                              # 規則なし
    bad = ve._validation_violations(_Many(cells))
    assert bad == [('C3', '緑', 'リスト: 赤,青'), ('D3', '99', '整数: 1〜10')]


def test_validation_violations_single_cell_path():
    assert ve._validation_violations(_Cell('C3', '緑', val=_Val(3, '赤,青', ok=False))) == [('C3', '緑', 'リスト: 赤,青')]
    assert ve._validation_violations(_Cell('C3', '赤', val=_Val(3, '赤,青', ok=True))) == []
    assert ve._validation_violations(_Cell('C3', 'x')) == []                 # 規則が無い


class _WS:
    def __init__(self, filled):
        self.filled = filled

    def Cells(self, r, c):
        a = chr(64 + c) + str(r)
        return _Cell(a, value=self.filled.get(a))


def test_spill_blockers_finds_the_value_below_or_right():
    ws = _WS({'F2': '人の値', 'G1': 3})
    assert ve._spill_blockers(_Cell('F1'), ws) == ['F2', 'G1']
    assert ve._spill_blockers(_Cell('F1'), _WS({})) == []


def test_report_prints_spill_and_validation_without_stopping(capsys):
    ws = _WS({'F2': '人の値'})
    bad = _Cell('C3', '緑', val=_Val(3, '赤,青', ok=False))
    ve._report_spill_and_validation(ws, [bad], [_Cell('F1')])
    out = capsys.readouterr().out
    assert '【スピル】F1 は #SPILL!' in out and 'F2' in out and '人の値を消すのは人の判断' in out
    assert '【入力規則】' in out and 'C3=緑（リスト: 赤,青）' in out and '止まりません' in out


def test_report_is_silent_when_nothing_is_wrong(capsys):
    ve._report_spill_and_validation(_WS({}), [_Cell('C2', '赤', val=_Val(3, '赤,青', ok=True))], [])
    assert capsys.readouterr().out == ''
