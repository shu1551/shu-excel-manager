# -*- coding: utf-8 -*-
"""修理の試験（vbam_mend・2026-09-17 夜）: Excel にも AI にも触らない部分。

壊し方の候補（字面の書き換え）／撃った結果から症状を決める／人の言い方の依頼文／行の差の数え方。
本物（写しで撃つ・macro モードに直させる）は agent --mend で。
"""
import json as _json

import pytest

import vbam_macro as vmac
import vbam_mend as vm

_CODE = '''Sub 集計する()
    ' 依頼の語: 集計
    Dim ws As Worksheet
    Dim i As Long, total As Double
    Set ws = ActiveSheet
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" And ws.Cells(i, 2).Value > 0 Then
            total = total + ws.Cells(i, 2).Value
        End If
        If ws.Cells(i, 1).Value = "対応表" Then GoTo NextI
        ws.Cells(i, 3).Value = "済"
NextI:
    Next i
    ws.Cells(lastRow + 1, 2).Value = total
End Sub'''


def _ops(code=_CODE):
    return vm.mutants(code)


def test_mutants_cover_each_kind_of_break():
    ops = {m['op'] for m in _ops()}
    assert {'drop_end_if', 'drop_next', 'label_typo', 'member_typo', 'set_drop', 'literal_typo', 'var_typo',
            'loop_start', 'loop_end', 'cmp_flip', 'offset', 'drop_stmt'} <= ops


def test_mutants_change_exactly_one_line_and_never_the_comment_or_sub_line():
    base = _CODE.split('\n')
    for m in _ops():
        got = m['code'].split('\n')
        assert got[0] == base[0]                    # Sub の行は触らない
        assert "    ' 依頼の語: 集計" in got         # 注記の行は触らない
        assert got[-1] == 'End Sub'
        if m['op'] == 'drop_end_if' or m['op'] in ('drop_next', 'drop_stmt'):
            assert len(got) == len(base) - 1
        else:
            assert len(got) == len(base)
            assert sum(1 for a, b in zip(base, got) if a != b) == 1


def test_mutants_are_unique():
    codes = [m['code'] for m in _ops()]
    assert len(codes) == len(set(codes))


def test_label_typo_renames_only_a_goto_target():
    m = [x for x in _ops() if x['op'] == 'label_typo']
    assert len(m) == 1 and m[0]['new'] == 'NextI2:'


def test_var_typo_keeps_strings_and_keywords():
    for m in (x for x in _ops() if x['op'] == 'var_typo'):
        assert '"' not in m['old'] or m['old'].count('"') == m['new'].count('"')
        assert 'ws.Cells' in m['old'] or 'total' in m['old'] or 'lastRow' in m['old']
    typos = {x['new'] for x in _ops() if x['op'] == 'var_typo'}
    assert 'total = totall + ws.Cells(i, 2).Value' in typos


def test_literal_typo_only_touches_japanese_literals_and_not_createobject():
    code = _CODE.replace('Set ws = ActiveSheet', 'Set ws = ActiveSheet\n    Set d = CreateObject("Scripting.Dictionary")')
    lits = [x for x in vm.mutants(code) if x['op'] == 'literal_typo']
    assert lits and all('CreateObject' not in x['old'] for x in lits)
    assert any('"対応"' in x['new'] for x in lits)


def test_cmp_flip_does_not_touch_strings():
    code = '''Sub a()
    If s = "a<>b" Then x = 1
End Sub'''
    flips = [x['new'] for x in vm.mutants(code) if x['op'] == 'cmp_flip']
    assert flips == ['If s <> "a<>b" Then x = 1']


def test_loop_bounds():
    news = {x['op']: x['new'] for x in _ops() if x['op'] in ('loop_start', 'loop_end')}
    assert news['loop_start'] == 'For i = 2 + 1 To lastRow'
    assert news['loop_end'] == 'For i = 2 To lastRow - 1'


def test_no_mutants_without_sub():
    assert vm.mutants("' 注記だけ") == []


def test_classify():
    assert vm.classify([]) is None
    assert vm.classify(["コンパイルエラーで撃てません: [表の整理] a:3: Next i ― Next に対応する For がありません。"]) == 'compile'
    assert vm.classify(["マクロ a が実行時エラーで止まりました: 438 オブジェクトは…"]) == 'runtime'
    assert vm.classify(["マクロが 60 秒で終わらず、止めました（無限ループの疑い）"]) == 'hang'
    assert vm.classify(["マクロは表に何も書かずに終わりました（…）", "E2: 期待 'x' ／ マクロ後 ''"]) == 'nothing'
    assert vm.classify(["E2: 期待 'x' ／ マクロ後 ''"]) == 'wrong'


def test_symptom_compile_strips_tool_words_and_line():
    s = vm.symptom('compile', ["コンパイルエラーで撃てません: [表の整理] 集計する:12: Next i ― コンパイル エラー: "
                               "Next に対応する For がありません。（呼んでいる Function・Sub が無い／変数名・定数の打ち間違い）"],
                   '集計する', '明細')
    assert s == ("シート「明細」でマクロ「集計する」を撃つと「コンパイル エラー: Next に対応する For がありません。」と出て"
                 "動きません。前は動いていました。直してください。")


def test_symptom_runtime_and_wrong():
    s = vm.symptom('runtime', ["マクロ 集計する が実行時エラーで止まりました: 91 オブジェクト変数または With ブロック変数が"
                               "設定されていません。（ThisWorkbook でなく…）"], '集計する', '明細')
    assert "「実行時エラー '91': オブジェクト変数または With ブロック変数が設定されていません。」で止まります" in s
    assert 'ThisWorkbook' not in s
    w = vm.symptom('wrong', ["表の大きさ／位置が違う: 期待 5×3（左上 A1）・マクロ後 4×3（左上 A1）",
                             "C5: 期待 '120' ／ マクロ後 ''", "C4: 期待 '' ／ マクロ後 '9'"], '集計する', '明細')
    assert '表の行数が違います（5 行のはずが 4 行）' in w and 'C5 は「120」のはずが「空」' in w
    assert '期待' not in w and 'マクロ後' not in w


def test_changed_lines_ignores_indent_blank_and_attribute():
    a = 'Attribute VB_Name = "表の整理"\nSub a()\n    x = 1\n\nEnd Sub'
    b = 'Sub a()\nx = 1\nEnd Sub\n'
    assert vm.changed_lines(a, b) == 0
    assert vm.changed_lines(a, 'Sub a()\n    x = 2\nEnd Sub') == 1
    assert vm.changed_lines(a, 'Sub a()\n    x = 1\n    y = 2\nEnd Sub') == 1


# ----------------------------------------------------------------
# macro モードの直し（2026-09-17 夜・この試験で見つけた穴）
# ----------------------------------------------------------------

_MOD = '''Option Base 0

Sub 前の処理()
    x = 1
End Sub

Sub 振り直す()
    For dc = 1 To 3
        If hv = "" Then
            y = 2
    Next dc
End Sub
'''


def test_block_replace_inserts_end_if_and_returns_only_that_procedure():
    name, proc = vmac._block_replace_code(_MOD, "            y = 2\n    Next dc",
                                          "            y = 2\n        End If\n    Next dc")
    assert name == '振り直す'
    assert proc.split('\n')[0] == 'Sub 振り直す()' and proc.split('\n')[-1] == 'End Sub'
    assert '        End If\n    Next dc' in proc and '前の処理' not in proc


def test_block_replace_tolerates_indent_but_refuses_ambiguous_and_missing():
    name, proc = vmac._block_replace_code(_MOD, "y = 2\nNext dc", "y = 3\nNext dc")
    assert 'y = 3' in proc
    with pytest.raises(ValueError, match='並びがそのまま'):
        vmac._block_replace_code(_MOD, "y = 9\nNext dc", "x")
    two = _MOD + "\nSub 別()\n    y = 2\n    Next dc\nEnd Sub\n"
    with pytest.raises(ValueError, match='2 か所'):
        vmac._block_replace_code(two, "y = 2\nNext dc", "y = 3\nNext dc")


def test_block_replace_refuses_crossing_declarations_and_losing_end_sub():
    with pytest.raises(ValueError, match='またいで'):
        vmac._block_replace_code(_MOD, "    x = 1\nEnd Sub\n\nSub 振り直す()", "    x = 1")
    with pytest.raises(ValueError, match='宣言部'):
        vmac._block_replace_code(_MOD, "Option Base 0\n", "Option Base 1\n")


def test_replace_loss_note_counts_dropped_comments_but_not_indent():
    old = "Sub a()\n    ' 見出しを書く\n    x = 1\n    ' データを書く\n    y = 2\nEnd Sub"
    assert vmac._replace_loss_note(old, "Sub a()\n' 見出しを書く\nx = 1\n\n' データを書く\ny = 2\nEnd Sub") == ''
    note = vmac._replace_loss_note(old, "Sub a()\n    x = 1\n    y = 2\n    z = 3\nEnd Sub")
    assert '2 行消えて（うち注記 2 行）' in note and '1 行増えました' in note
    assert vmac._replace_loss_note(old, old.replace('y = 2', 'y = 3')) == ''


def test_repair_shows_diff_against_forged_canon(tmp_path, monkeypatch):
    import vbam_devtools as vd
    canon = "Sub 集計する()\n    For dc = amtCol + 1 To lastCol\n        x = 1\n    Next dc\nEnd Sub"
    (tmp_path / '_agent_forge.json').write_text(_json.dumps({'振り直し': {'sub': '集計する', 'code': canon}},
                                                          ensure_ascii=False), encoding='utf-8')
    monkeypatch.setattr(vd, 'SCRIPT_DIR', str(tmp_path))
    monkeypatch.delenv('VBAM_REPAIR_NO_CANON', raising=False)
    now = canon.replace('amtCol + 1 To', 'amtCol + 1 + 1 To')
    cn = vd._forge_canon_diff('集計する', now)
    assert cn and cn['minus'] == 1 and cn['plus'] == 1 and '「振り直し」' in cn['label']
    parts = {'module': '表の整理', 'proc': '集計する', 'code': now, 'canon': cn}
    text = "\n".join(vd._repair_report(parts))
    assert '前は動いていた' in text and 'amtCol + 1 + 1' in text
    assert vd._forge_canon_diff('ほかのマクロ', now) is None
    monkeypatch.setenv('VBAM_REPAIR_NO_CANON', '1')
    assert vd._forge_canon_diff('集計する', now) is None                  # --blind


def test_gosub_input_gap_finds_the_one_call_without_its_input():
    import vbam_devtools as vd
    code = '''Sub 突合()
    For ri = 2 To 9
        normInput = ws.Cells(ri, 1).Value
        GoSub NormTkt
        keyA = normOutput
        tktVal = ws.Cells(ri, 2).Value
        GoSub NormTkt
        keyB = normOutput
        normInput = ws.Cells(ri, 3).Value
        GoSub NormTkt
    Next ri
    Exit Sub
NormTkt:
    normOutput = Trim(normInput)
    Return
End Sub'''.split('\n')
    gaps = vd._gosub_input_gaps(code)
    assert gaps == [(6, 'NormTkt', 'normInput', 2)]
    fixed = code[:6] + ['        normInput = tktVal'] + code[6:]
    assert vd._gosub_input_gaps(fixed) == []
    assert vd._gosub_input_gaps(['Sub a()', '    x = 1', 'End Sub']) == []


class _FakeCM:
    def __init__(self, text):
        self._lines = text.split('\n')
        self.CountOfLines = len(self._lines)

    def Lines(self, start, count):
        return '\r\n'.join(self._lines[start - 1:start - 1 + count])


class _FakeWB:
    def __init__(self, text):
        cm = _FakeCM(text)

        class _Comp:
            CodeModule = cm

        class _Proj:
            @staticmethod
            def VBComponents(name):
                if name != '表の整理':
                    raise KeyError(name)
                return _Comp

        self.VBProject = _Proj


_BIG = "\n".join(["Sub 表を整える()", "    GoSub 数を読む", "    Exit Sub", "数を読む:", "    t = Left(t, Len(t) - 2)",
                  "    Return", "End Sub"])


def test_get_on_a_gosub_label_returns_the_block_and_lines_reads_a_range():
    wb = _FakeWB(_BIG)
    lab, ok, out = vmac._get_label_act({'op': 'get', 'name': '数を読む'}, wb, '表の整理')
    assert ok and 'ラベル' in lab and '4: 数を読む:' in out and '6:     Return' in out and 'Exit Sub' not in out
    assert vmac._get_label_act({'op': 'get', 'name': '無いラベル'}, wb, '表の整理') is None
    lab, ok, out = vmac._read_lines_act({'op': 'lines', 'from': 2, 'to': 3}, wb, '表の整理')
    assert ok and lab == 'lines 2-3' and out.startswith('   2:     GoSub 数を読む') and '4 行目から先' in out
    assert vmac._read_lines_act({'op': 'lines', 'from': 'x'}, wb, '表の整理')[1] is False
    res = vmac._execute_macro([{'op': 'get', 'name': '数を読む'}, {'op': 'lines', 'from': 5, 'to': 5}], False, wb,
                              existing=['表を整える'], default_module='表の整理')
    assert [r[1] for r in res] == [True, True] and 'Len(t) - 2' in res[1][2]


def test_read_before_assign_finds_a_dropped_assignment_but_not_loop_carried_or_gosub_outputs():
    import vbam_devtools as vd
    base = '''Sub 集計()
    Dim total As Double, prev As String, key As String, outV As String
    total = 0
    For r = 2 To 9
        key = ws.Cells(r, 1).Value
        If key <> prev Then cnt = cnt + 1
        prev = key
        GoSub 読む
        total = total + Len(outV)
    Next r
    ws.Range("A1").Value = total
    Exit Sub
読む:
    outV = Trim(key)
    Return
End Sub'''.split('\n')
    assert vd._read_before_assign(base) == []            # prev は前の周回・outV は GoSub が入れる
    dropped = [l for l in base if l.strip() != 'total = 0']
    assert vd._read_before_assign(dropped) == []         # total は同じループで入れている（前の周回）＝拾わない
    dropped = [l for l in base if l.strip() != 'key = ws.Cells(r, 1).Value']
    assert vd._read_before_assign(dropped) == [(4, 'key')]


def test_needs_run_and_expectations():
    wrong = "シート「支出明細」でマクロ「集計」を撃つとエラーは出ませんが結果が合いません。例えば F8 は「対応なし」のはずが「空」・G8 は「72000」のはずが「空」・H2 は空のはずが「3」。"
    assert vmac._needs_run(wrong)
    assert vmac._request_expectations(wrong) == [('F8', '対応なし'), ('G8', '72000'), ('H2', '')]
    assert vmac._needs_run("撃つとエラーは出ませんが、表に何も書かれません。")
    assert not vmac._needs_run("撃つと「コンパイル エラー: Next に対応する For がありません。」と出て動きません。")
    assert not vmac._needs_run("撃つと「実行時エラー '438': オブジェクトは、このプロパティまたはメソッドをサポートしていません。」で止まります。")


def test_size_expectations_are_read_and_checked():
    req = "エラーは出ませんが結果が合いません。例えば 表の列数が違います（7 列のはずが 8 列）・G1 は「支出額」のはずが「空」。"
    exp = vmac._request_expectations(req)
    assert exp == [('G1', '支出額'), ('#列', '7')]
    ok, lines = vmac._expectation_check(exp, {'G1': '支出額', '#列': '8'})
    assert not ok and '表の列数' in lines[1] and lines[1].endswith('×')
    assert vmac._expectation_check(exp, {'G1': '支出額', '#列': '7'})[0]


def test_minimize_fix_reverts_only_the_unneeded_hunk(monkeypatch):
    start = "Sub a()\n    outRow = hdrRow + 2 + j\n    x = 1\n    uRow = hdrRow + 1 + n\nEnd Sub"
    now = "Sub a()\n    outRow = hdrRow + 1 + j\n    x = 1\n    uRow = hdrRow + n\nEnd Sub"
    state = {'code': now}
    monkeypatch.setattr(vmac, '_proc_text', lambda wb, name, module=None: state['code'])

    def fake_apply(wb, text, module=None):
        state['code'] = text
        return True
    monkeypatch.setattr(vmac, '_apply_proc', fake_apply)
    monkeypatch.setattr(vmac.va, '_run_cmd', lambda toks, wb=None: (True, 'ok'))

    def fake_verify(results, verified, expectations, sheet_hint=None, wb=None, score=None):
        good = 'outRow = hdrRow + 1 + j' in state['code']        # 要るのは outRow の直しだけ
        return results, good, None
    monkeypatch.setattr(vmac, '_verify_results', fake_verify)
    notes = vmac._minimize_fix(None, 'a', '表の整理', start, [('F2', '人件費')])
    assert len(notes) == 1 and 'uRow = hdrRow + n' in notes[0]
    assert state['code'].split('\n')[1].strip() == 'outRow = hdrRow + 1 + j'
    assert state['code'].split('\n')[3].strip() == 'uRow = hdrRow + 1 + n'


def test_expectation_check_numbers_and_blank():
    ok, lines = vmac._expectation_check([('G8', '72000'), ('H2', '')], {'G8': '72,000.0'})
    assert ok and all(l.endswith('○') for l in lines)
    ok, lines = vmac._expectation_check([('F8', '対応なし')], {'F8': '合計'})
    assert not ok and '×' in lines[0]


def test_verify_results_tracks_writes_and_rehearse(tmp_path):
    after = tmp_path / 'x.xlsm.after.json'
    after.write_text(_json.dumps({'active': '支出明細', 'sheets': {'支出明細': {'cells': [
        {'r': 8, 'c': {'F': '対応なし', 'G': '72000'}}]}}}, ensure_ascii=False), encoding='utf-8')
    reh = f"マクロ実行: 成功\n  実行前/後 snapshot: {tmp_path / 'x.xlsm.before.json'} / {after}\n"
    exp = [('F8', '対応なし'), ('G8', '72000')]
    res, v, s = vmac._verify_results([('code_replace x', True, 'ok'), ('compile（書き換え後・道具が自動）', True, 'ok')],
                                     True, exp, '支出明細')
    assert v is None                                         # 書き換えたら確かめ直し
    res, v, s = vmac._verify_results([('rehearse 集計', True, reh)], v, exp, '支出明細')
    assert v is True and '依頼の期待との突き合わせ' in res[0][2] and '全部合いました' in res[0][2] and s == (2, 2)
    res, v, s = vmac._verify_results([('rehearse 集計', True, reh)], None, [('F8', '合計')], '支出明細')
    assert v is False and '×' in res[0][2]
    res, v, s = vmac._verify_results([('rehearse 集計', False, 'マクロ実行: 失敗')], None, exp, None)
    assert v is False
    # 直しの後に ○ が増えない／減った＝undo を勧める（前の score と比べる・書き換えていない撃ち直しでは言わない）
    bad = [('F8', '対応なし'), ('G8', '1')]
    res, v, s = vmac._verify_results([('code_replace y', True, 'ok'), ('rehearse 集計', True, reh)], None, bad, '支出明細',
                                     score=(1, 2))
    assert s == (1, 2) and 'のまま' in res[1][2] and '"op":"undo"' in res[1][2]
    res, v, s = vmac._verify_results([('code_replace y', True, 'ok'), ('rehearse 集計', True, reh)], None, bad, '支出明細',
                                     score=(2, 2))
    assert '減りました' in res[1][2]
    res, v, s = vmac._verify_results([('rehearse 集計', True, reh)], None, bad, '支出明細', score=(1, 2))
    assert 'undo' not in res[0][2]


def test_execute_macro_pushes_one_undo_snapshot_per_successful_fix(monkeypatch):
    snaps = iter(['姿1', '姿2', '姿3'])
    monkeypatch.setattr(vmac, '_proc_text', lambda wb, name, module=None: next(snaps))
    ran = []

    def fake_run(toks, wb=None):
        ran.append(list(toks))
        if toks[0] == 'code-replace' and toks[2] == 'x':
            return True, '「x」にマッチする行はありません'
        return True, 'ok'
    monkeypatch.setattr(vmac.va, '_run_cmd', fake_run)
    stack = []
    res = vmac._execute_macro([{'op': 'code_replace', 'search': 'a', 'replace': 'b', 'module': '表の整理'},
                               {'op': 'code_replace', 'search': 'c', 'replace': 'd', 'module': '表の整理'},
                               {'op': 'grep', 'text': 'm'}], True, None, default_module='表の整理', undo=(stack, 'm'))
    assert stack == ['姿1', '姿2'] and res[-1][0].startswith('compile')
    stack = []
    res = vmac._execute_macro([{'op': 'code_replace', 'search': 'x', 'replace': 'b', 'module': '表の整理'},
                               {'op': 'grep', 'text': 'm'}], True, None, default_module='表の整理', undo=(stack, 'm'))
    assert stack == [] and '止めました' in res[1][2]


def test_undo_step_restores_previous_procedure(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(vmac.va, 'LAST_PROC_FILE', str(tmp_path / '_last_proc.vba'))
    monkeypatch.setattr(vmac.va, '_run_cmd', lambda toks, wb=None: (calls.append(list(toks)) or (True, 'ok')))
    assert vmac._undo_step([], None)[0][1] is False
    stack = ['Sub a()\n    x = 1\nEnd Sub']
    res = vmac._undo_step(stack, None, '表の整理')
    assert res[0][1] and stack == [] and calls[0] == ['replace-procedure', '-y', '--module', '表の整理']
    assert calls[1] == ['compile'] and (tmp_path / '_last_proc.vba').read_text(encoding='utf-8').startswith('Sub a()')
    with pytest.raises(ValueError, match='undo'):
        vmac._macro_action_to_tokens({'op': 'undo'})
