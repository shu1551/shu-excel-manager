# -*- coding: utf-8 -*-
"""vbam_flow（マクロの流れ図）のテスト（2026-10-01）。純 Python・COM なし。"""
import glob
import os
import re

import pytest

import vbam_flow as vf


def _edges(mer):
    """辺を [(元, 先, 言葉, 点線か)] に。"""
    out = []
    for m in re.finditer(r'^\s+(\w+) (-->|-\.->)(?:\|"([^"]*)"\|)? (\w+)$', mer, re.M):
        out.append((m.group(1), m.group(4), m.group(3) or '', m.group(2) == '-.->'))
    return out


def _labels(mer):
    """節点 id → 文字。"""
    return {m.group(1): m.group(2) for m in re.finditer(r'^\s+(\w+)(?:\[\[|\(\[|\{\{|\[|\{)"(.*)"(?:\]\]|\]\)|\}\}|\]|\})$', mer, re.M)}


def _find(mer, text):
    labs = _labels(mer)
    hits = [k for k, v in labs.items() if text in v]
    assert hits, f'{text!r} の節点がありません: {labs}'
    return hits[0]


def _has_edge(mer, a_text, b_text, label=None, dotted=None):
    a, b = _find(mer, a_text), _find(mer, b_text)
    for s, d, lab, dot in _edges(mer):
        if s == a and d == b and (label is None or lab == label) and (dotted is None or dot == dotted):
            return True
    return False


def test_if_else_and_exit_sub():
    code = ("Sub 試験()\n"
            "    Dim x As Long\n"
            "    ' 読み込む\n"
            "    x = 1\n"
            "    If x = 0 Then\n"
            "        Exit Sub\n"
            "    ElseIf x = 1 Then\n"
            "        x = 10\n"
            "    Else\n"
            "        x = 20\n"
            "    End If\n"
            "    x = x + 1\n"
            "End Sub\n")
    mer, st = vf.flow_to_mermaid(code)
    assert st['decisions'] == 2 and st['loops'] == 0
    assert _has_edge(mer, 'x = 0', 'Exit Sub', 'はい')
    assert _has_edge(mer, 'x = 0', 'x = 1 ?', 'いいえ')
    assert _has_edge(mer, 'x = 1 ?', 'x = 10', 'はい')
    assert _has_edge(mer, 'x = 1 ?', 'x = 20', 'いいえ')
    assert _has_edge(mer, 'x = 10', 'x = x + 1') and _has_edge(mer, 'x = 20', 'x = x + 1')
    assert _has_edge(mer, 'Exit Sub', '終了')
    assert '読み込む' in mer                               # 直前のコメントが節点の言葉になる
    assert 'Dim x' not in mer                             # 宣言は図に載せない


def test_for_loop_back_edge_and_exit_for():
    code = ("Sub 試験()\n"
            "    For i = 1 To 5\n"
            "        If a(i) = 0 Then Exit For\n"
            "        n = n + 1\n"
            "    Next i\n"
            "    Done = True\n"
            "End Sub\n")
    mer, st = vf.flow_to_mermaid(code)
    assert st['loops'] == 1
    assert _has_edge(mer, 'For i = 1 To 5', 'a(i) = 0', '続ける')
    assert _has_edge(mer, 'n = n + 1', 'For i = 1 To 5')                     # 戻る辺
    assert _has_edge(mer, 'For i = 1 To 5', 'Done = True', '終わり')
    assert _has_edge(mer, 'Exit For', 'Done = True', '抜ける')


def test_do_loop_while_post_test():
    code = ("Sub 試験()\n"
            "    Do\n"
            "        n = n + 1\n"
            "    Loop While n < 3\n"
            "End Sub\n")
    mer, _ = vf.flow_to_mermaid(code)
    assert _has_edge(mer, 'n = n + 1', 'Loop While n', None)
    assert _has_edge(mer, 'Loop While n', 'n = n + 1', '続ける')


def test_select_case_and_single_line_if_goto():
    code = ("Sub 試験()\n"
            "    Select Case k\n"
            "        Case 1\n"
            "            a = 1\n"
            "        Case 2, 3\n"
            "            a = 2\n"
            "        Case Else\n"
            "            a = 9\n"
            "    End Select\n"
            "    If a = 9 Then GoTo 後始末\n"
            "    b = 1\n"
            "後始末:\n"
            "    c = 1\n"
            "End Sub\n")
    mer, st = vf.flow_to_mermaid(code)
    assert _has_edge(mer, 'Select Case k', 'a = 1', 'Case 1')
    assert _has_edge(mer, 'Select Case k', 'a = 2', 'Case 2, 3')
    assert _has_edge(mer, 'Select Case k', 'a = 9', 'それ以外')
    assert _has_edge(mer, 'GoTo 後始末', '後始末:')                          # GoTo は同じ枠のラベルへ
    assert _has_edge(mer, 'b = 1', '後始末:')


def test_gosub_section_and_error_handler_are_separate_frames():
    code = ("Sub 試験()\n"
            "    On Error GoTo 失敗\n"
            "    s = \"x\"\n"
            "    GoSub 整える\n"
            "    Exit Sub\n"
            "整える:\n"
            "    s = Trim(s)\n"
            "    Return\n"
            "失敗:\n"
            "    MsgBox \"x\"\n"
            "End Sub\n")
    mer, st = vf.flow_to_mermaid(code)
    assert st['gosubs'] == 1
    assert mer.count('subgraph') == 2 and mer.count('\n    end') == 2        # 整える と 失敗 が別の枠
    assert _has_edge(mer, 'GoSub 整える', '整える:', '呼ぶ', dotted=True)
    assert _has_edge(mer, 'エラーが起きたら 失敗', '失敗:', 'エラー時', dotted=True)
    assert _has_edge(mer, 'GoSub 整える', 'Exit Sub')                      # 戻ってきて次へ進む
    assert _find(mer, 'Return')


def test_escaping_quotes_and_angle_brackets():
    code = "Sub 試験()\n    If a <> \"x\" And b > 1 Then\n        c = \"<b>\" & d\n    End If\nEnd Sub\n"
    mer, _ = vf.flow_to_mermaid(code)
    assert '#quot;' in mer and '#lt;' in mer and '#gt;' in mer
    assert '"x"' not in mer                                              # 生の " が節点の言葉に残らない


def test_collapse_deep_nesting_when_too_many_nodes():
    body = ''.join(f"        If a{i} = 1 Then\n            For j = 1 To 3\n                x{i} = j\n            Next j\n        End If\n"
                   for i in range(40))
    code = "Sub 大きい()\n" + "    For k = 1 To 9\n" + body + "    Next k\nEnd Sub\n"
    full, st_full = vf.flow_to_mermaid(code, max_nodes=10000)
    small, st_small = vf.flow_to_mermaid(code, max_nodes=60)
    assert st_full['collapsed'] is None and st_full['nodes'] > 100
    assert st_small['collapsed'] is not None and st_small['nodes'] < st_full['nodes']
    assert '省略' in small


def test_decision_keeps_condition_and_adds_comment():
    """分岐のひし形は条件を必ず出し、直前のコメントは添える（コメントだけにすると何を聞いているか消える）。"""
    code = "Sub a()\n    ' 関数が入っていれば参照ではない\n    If InStr(s, \"(\") > 0 Then\n        x = 1\n    End If\nEnd Sub\n"
    mer, _ = vf.flow_to_mermaid(code)
    lab = _labels(mer)[_find(mer, 'InStr(s')]
    assert 'InStr(s' in lab and '関数が入っていれば参照ではない' in lab and lab.endswith('）')


def test_no_procedure_raises():
    with pytest.raises(ValueError):
        vf.flow_to_mermaid("x = 1\n")


def test_markdown_wrapper():
    mer, st = vf.flow_to_mermaid("Sub a()\n    x = 1\nEnd Sub\n")
    md = vf.mermaid_markdown('a', mer, st)
    assert md.startswith('# a の流れ図') and '```mermaid' in md and 'AI は使っていません' in md


_BAS = sorted(glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), '*.bas')) +
              glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'macros', '*.bas')))


@pytest.mark.skipif(not _BAS, reason='.bas が無い')
def test_every_real_procedure_makes_a_consistent_diagram():
    """実在の全プロシージャで、例外が出ず、辺の端点がすべて定義された節点になる。"""
    n = 0
    for fn in _BAS:
        if os.path.basename(fn).startswith('_'):
            continue
        t = open(fn, encoding='cp932', errors='replace').read().replace('\r\n', '\n')
        for m in re.finditer(r'^(?:Public |Private |Friend )?(?:Sub|Function) (\S+?)\(.*?^End (?:Sub|Function)', t, re.S | re.M):
            mer, st = vf.flow_to_mermaid(m.group(0), m.group(1))
            defined = set(re.findall(r'^\s*(\w+)[\[\(\{]', mer, re.M)) | {'END', 'START'}
            used = {x for e in _edges(mer) for x in e[:2]}
            assert used <= defined, (fn, m.group(1), sorted(used - defined)[:5])
            n += 1
    assert n > 0
