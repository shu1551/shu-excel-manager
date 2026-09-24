# -*- coding: utf-8 -*-
"""vbam_mend.py — マクロ修理の試験（agent --mend・2026-09-17 夜）

鍛えて登録したマクロ（_agent_forge.json＝前の表・正解の表・別の表を持つ）を Python が種ごとに壊し、
壊れたマクロを練習台のブック（前の表の写し .xlsm）に植えて、人の言い方の症状だけを添えて agent の macro モードに直させる。
直ったマクロは写しで前の表と別の表に撃ち、正解と値で突き合わせる（AI の done・コンパイルの通過とは別に数える）。

  agent --mend [--seed N] [--only 名,名] [--kinds compile,runtime,nothing,wrong] [--max-turns N] [--dry-run]

なぜ: 表の仕事は鍛える回路と試験で種ごとに回したが、マクロの修理は手書きの 10 問（B2:B4 の合計ほどの 5 行のマクロ）しか
無かった。本物の大きさ（150〜400 行）のマクロで、症状（コンパイル・実行時エラー・何も書かない・数が合わない）ごとに測る。

壊し方（_OPS）は字面の書き換えだけ（End If／Next を抜く・ラベル名・メンバー名・Set 抜け・文字の打ち間違い・変数名の打ち間違い・
For の始まりと終わりの 1 ずれ・比べ方の反転・列のずれ・文を 1 行抜く）。壊したマクロは写しで撃ち、前の表の正解と違う物だけを
お題にする（撃っても同じ結果の壊し方＝お題にならない）。症状の分類も撃った結果から決める（壊し方の名前からは決めない）。

状態ファイル（台帳・控え・覚書・記録）はこの回の置き場（_agent_mend/日時/_state）へ逃がす＝本番の台帳に混ぜない。
点数は _agent_mend.jsonl に 1 回 1 行。
"""
import os
import re
import sys
import json
import time
import random
import difflib
import hashlib
import shutil
import subprocess
from vbam_core import SCRIPT_DIR

_MEND_DIR = os.path.join(SCRIPT_DIR, '_agent_mend')
_MEND_SCORE_FILE = os.path.join(SCRIPT_DIR, '_agent_mend.jsonl')
_MODULE = '表の整理'
_KINDS = ('compile', 'runtime', 'nothing', 'wrong')
_HARD_KINDS = ('edge', 'double')     # ある表でだけ合わない／壊れが 2 つ（コンパイル＋数が合わない）
_KIND_NAME = {'compile': 'コンパイル', 'runtime': '実行時エラー', 'nothing': '何も書かない', 'wrong': '数が合わない',
              'hang': '終わらない', 'edge': 'ある表でだけ', 'double': '壊れが2つ'}
_TRIES_PER_CASE = 16          # 1 本のマクロで壊し方を撃って試す上限（1 回 3〜6 秒）

# ----------------------------------------------------------------
# 壊す（純 Python）
# ----------------------------------------------------------------

_SUB_HEAD_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(', re.I)
_END_SUB_RE = re.compile(r'^\s*End\s+Sub\s*$', re.I)
_KEYWORD_HEAD_RE = re.compile(
    r'^\s*(Dim|ReDim|Static|Const|If|ElseIf|Else|End|For|Next|Do|Loop|While|Wend|Select|Case|With|On|Exit|GoTo|'
    r'GoSub|Return|Set|Sub|Function|Private|Public|Call|Resume|Option|Attribute)\b', re.I)
_LABEL_RE = re.compile(r'^\s*([A-Za-z_぀-ヿ一-鿿][\w぀-ヿ一-鿿]*):\s*$')
_IDENT_RE = re.compile(r'(?<![\w\.぀-ヿ一-鿿])([A-Za-z_][A-Za-z0-9_]*)(?![\w぀-ヿ一-鿿(])')
_MEMBER_TYPO = {'Value': 'Valeu', 'Row': 'Rwo', 'Column': 'Colum', 'Count': 'Cuont', 'Formula': 'Fromula',
                'Cells': 'Cels', 'Rows': 'Rosw'}
_VB_WORDS = {w.lower() for w in (
    'And Or Not Xor Mod Is Like New Nothing True False Empty Null Me To Step Then Else ElseIf End If For Each In Next '
    'Do Loop While Until Wend With Dim ReDim Preserve As Set Let Call Exit Sub Function GoTo On Error Resume Select Case '
    'Long Integer String Double Boolean Variant Object Range Worksheet Workbook Date Single Currency Byte '
    'Trim LTrim RTrim Len Left Right Mid InStr Replace Split Join UCase LCase CStr CLng CDbl CInt CDate IsNumeric IsDate '
    'IsEmpty IsError Val Format Abs Int Round Array UBound LBound Chr Asc Year Month Day DateSerial Now StrConv '
    'ActiveSheet ActiveWorkbook Application Worksheets Sheets Cells Range Rows Columns CreateObject Debug Print '
    'vbNullString vbCrLf vbLf xlUp xlDown xlToLeft xlToRight WorksheetFunction').split()}


def _mask(line):
    """文字列の中身と行末の注記を同じ長さの記号で伏せる（位置はそのまま）。演算子や名前を字面の中から拾わないため。"""
    out, in_str = [], False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif in_str:
            out.append('#')
        elif ch == "'":
            out.append('#' * (len(line) - i))
            break
        else:
            out.append(ch)
    return ''.join(out)


def _assign_split(m):
    """伏せた行の代入の = の位置（かっこの外の最初の =・<> <= >= は除く）。代入でなければ None。"""
    if re.match(r'^\s*(Dim|ReDim|Const|Static|If|ElseIf|For|Do|While|Loop|Select|Case|With|Exit|On|GoTo|Call)\b', m, re.I):
        return None
    depth = 0
    for i, ch in enumerate(m):
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == '=' and depth == 0:
            prev = m[i - 1] if i else ''
            nxt = m[i + 1] if i + 1 < len(m) else ''
            if prev in '<>' or nxt in '<>':
                continue
            return i if m[:i].strip() else None
    return None


def _sub_span(lines):
    """(Sub の行, End Sub の行) の添字。見つからなければ (None, None)。"""
    start = next((i for i, l in enumerate(lines) if _SUB_HEAD_RE.match(l)), None)
    if start is None:
        return None, None
    end = next((i for i in range(start + 1, len(lines)) if _END_SUB_RE.match(lines[i])), None)
    return start, end


def _variables(lines, lo, hi):
    """本文で左辺に出る名前（＝変数）。Dim の名前も入れる。"""
    names = set()
    for l in lines[lo:hi]:
        m = _mask(l)
        md = re.match(r'^\s*(?:Dim|ReDim(?:\s+Preserve)?|Static)\s+(.+)$', m, re.I)
        if md:
            for part in md.group(1).split(','):
                w = re.match(r'\s*([A-Za-z_]\w*)', part)
                if w:
                    names.add(w.group(1))
            continue
        ma = re.match(r'^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=', m, re.I)
        if ma and not _KEYWORD_HEAD_RE.match(m.replace('Set ', '', 1) if m.strip().lower().startswith('set ') else m):
            names.add(ma.group(1))
        mf = re.match(r'^\s*For\s+(?:Each\s+)?([A-Za-z_]\w*)', m, re.I)
        if mf:
            names.add(mf.group(1))
    return {n for n in names if n.lower() not in _VB_WORDS}


def mutants(code):
    """壊し方の候補を全部並べる（純 Python・順序は決まっている）。→ [{'op','line','old','new','code'}]"""
    code = code.replace('\r\n', '\n')
    lines = code.split('\n')
    s, e = _sub_span(lines)
    if s is None or e is None:
        return []
    out = []
    seen = set()

    def put(op, i, new=None):
        nl = list(lines)
        if new is None:
            del nl[i]
        else:
            if new == lines[i]:
                return
            nl[i] = new
        c = '\n'.join(nl)
        h = hashlib.md5(c.encode('utf-8')).hexdigest()
        if h in seen:
            return
        seen.add(h)
        out.append({'op': op, 'line': i + 1, 'old': lines[i].strip(), 'new': '' if new is None else new.strip(),
                    'code': c})

    body = [i for i in range(s + 1, e) if lines[i].strip() and not lines[i].strip().startswith("'")]
    variables = _variables(lines, s + 1, e)
    labels = {m.group(1) for i in body for m in [_LABEL_RE.match(lines[i])] if m}
    goto_targets = {g.rstrip(':') for i in body for g in re.findall(r'\bGo(?:To|Sub)\s+([^\s:]+)', _mask(lines[i]), re.I)}
    for i in body:
        raw, m = lines[i], _mask(lines[i])
        st = m.strip()
        # コンパイル: ブロックの終わりを抜く
        if re.match(r'^End\s+If$', st, re.I):
            put('drop_end_if', i)
        if re.match(r'^Next\b', st, re.I):
            put('drop_next', i)
        # コンパイル: 飛び先のラベルの名前違い
        lm = _LABEL_RE.match(raw)
        if lm and lm.group(1) in goto_targets:
            put('label_typo', i, raw.replace(lm.group(1), lm.group(1) + '2', 1))
        # コンパイル／実行時: メンバー名の打ち間違い
        for mm in re.finditer(r'\.(' + '|'.join(_MEMBER_TYPO) + r')\b', m):
            a, b = mm.span(1)
            put('member_typo', i, raw[:a] + _MEMBER_TYPO[mm.group(1)] + raw[b:])
        # 実行時: Set の付け忘れ
        sm = re.match(r'^(\s*)Set\s+', m, re.I)
        if sm:
            put('set_drop', i, raw[:sm.start(0)] + sm.group(1) + raw[sm.end(0):])
        # 何も書かない／数が合わない: 探す語の打ち間違い（最後の 1 字を落とす）
        if 'createobject' not in m.lower():
            for lm2 in re.finditer(r'"([^"]{2,})"', raw):
                lit = lm2.group(1)
                if re.search(r'[^\x00-\x7f]', lit):
                    a, b = lm2.span(1)
                    put('literal_typo', i, raw[:a] + lit[:-1] + raw[b:])
        # 数が合わない: 右辺の変数名の打ち間違い（Option Explicit が無い＝コンパイルは通る）
        eq = _assign_split(m)
        if eq is not None and ':' not in m:
            for vm in _IDENT_RE.finditer(m[eq + 1:]):
                name = vm.group(1)
                if name in variables and len(name) >= 3:
                    a, b = eq + 1 + vm.start(1), eq + 1 + vm.end(1)
                    put('var_typo', i, raw[:a] + name + name[-1] + raw[b:])
        # 数が合わない: For の始まり・終わりの 1 ずれ
        fm = re.match(r'^(\s*For\s+[A-Za-z_]\w*\s*=\s*)(.+?)(\s+To\s+)(.+?)(\s+Step\s+.+)?\s*$', m, re.I)
        if fm:
            a, b = fm.span(2)
            put('loop_start', i, raw[:a] + raw[a:b] + ' + 1' + raw[b:])
            a, b = fm.span(4)
            put('loop_end', i, raw[:a] + raw[a:b] + ' - 1' + raw[b:])
        # 何も書かない／数が合わない: 比べ方の反転
        cm = re.match(r'^\s*(?:ElseIf|If)\b(.*?)\bThen\b', m, re.I)
        if cm:
            a0 = cm.start(1)
            cond = cm.group(1)
            for om in re.finditer(r'<>|>=|<=|(?<![<>])=(?![<>])|(?<![<])>(?!=)|<(?![>=])|\bAnd\b|\bOr\b', cond):
                a, b = a0 + om.start(), a0 + om.end()
                rep = {'<>': '=', '>=': '>', '<=': '<', '=': '<>', '>': '>=', '<': '<=',
                       'And': 'Or', 'Or': 'And'}.get(om.group(0), om.group(0))
                if rep.lower() in ('and', 'or'):
                    rep = rep if om.group(0)[0].isupper() else rep.lower()
                put('cmp_flip', i, raw[:a] + rep + raw[b:])
        # 数が合わない: 列・行のずれ（代入・If の中の + n／- n を 1 増やす。For の範囲は loop_* が持つ）
        if eq is not None or re.match(r'^\s*(?:ElseIf|If)\b', m, re.I):
            for nm in re.finditer(r'([+\-])\s*(\d+)\b', m):
                a, b = nm.start(2), nm.end(2)
                put('offset', i, raw[:a] + str(int(nm.group(2)) + 1) + raw[b:])
        # 何でも: 文を 1 行抜く（代入・メソッド呼び）
        if not _KEYWORD_HEAD_RE.match(m) and not _LABEL_RE.match(raw) and ('=' in m or '.' in m) and ':' not in m:
            put('drop_stmt', i)
    return out


# ----------------------------------------------------------------
# 撃って分類する（写し・自分の Excel・AI なし）
# ----------------------------------------------------------------

_COMPILE_PREFIX = 'コンパイルエラーで撃てません: '
_RUNTIME_RE = re.compile(r'実行時エラーで止まりました: (\S+) (.*?)(?:（.*）)?$')
_CELL_RE = re.compile(r"^\s*([A-Z]{1,3}\d+): 期待 '(.*)' ／ マクロ後 '(.*)'$")
_KIND_CELL_RE = re.compile(r"^\s*([A-Z]{1,3}\d+): 型が違う 期待 (\S+) '(.*)' ／ マクロ後 (\S+) '(.*)'$")
_SIZE_RE = re.compile(r'期待 (\d+)×(\d+).*?マクロ後 (\d+)×(\d+)')


def classify(lines):
    """_run_on_copies の 1 枚ぶんの不一致の行 → 症状（None＝正解どおり）。"""
    if not lines:
        return None
    first = str(lines[0])
    if first.startswith(_COMPILE_PREFIX):
        return 'compile'
    if '実行時エラーで止まりました' in first or first.startswith('マクロの実行で止まりました'):
        return 'runtime'
    if '秒で終わらず' in first:
        return 'hang'
    if first.startswith('マクロを読み込めませんでした'):
        return 'load'
    if first.startswith('マクロは表に何も書かずに'):
        return 'nothing'
    return 'wrong'


def symptom(kind, lines, sub, sheet, edge=False):
    """人が気づく形の依頼文（直し方・壊した場所は書かない）。edge＝ほかの表では合う（この表でだけ出る）。"""
    head = (f"マクロ「{sub}」はいつもの表では合っているのに、シート「{sheet}」のこの表で撃つと" if edge
            else f"シート「{sheet}」でマクロ「{sub}」を撃つと")
    if kind == 'compile':
        detail = str(lines[0])[len(_COMPILE_PREFIX):]
        detail = re.sub(r'（呼んでいる Function・Sub が無い／変数名・定数の打ち間違い）$', '', detail)
        msg = detail.rsplit(' ― ', 1)[-1].strip() if ' ― ' in detail else detail.strip()
        msg = re.sub(r'^(コンパイル\s*エラー\s*[:：]\s*)+', '', msg)
        return f"{head}「コンパイル エラー: {msg}」と出て動きません。前は動いていました。直してください。"
    if kind == 'runtime':
        m = _RUNTIME_RE.search(str(lines[0]))
        msg = f"実行時エラー '{m.group(1)}': {m.group(2).strip()}" if m else str(lines[0])
        return f"{head}「{msg}」で止まります。前は動いていました。直してください。"
    if kind == 'nothing':
        return f"{head}エラーは出ませんが、表に何も書かれません。前はちゃんと書けていました。直してください。"
    if kind == 'hang':
        return f"{head}終わらずに固まります。前はすぐ終わっていました。直してください。"
    ex = []
    for l in lines:
        c = _CELL_RE.match(str(l))
        if c and c.group(2) != '':
            ex.append(f"{c.group(1)} は「{c.group(2)}」のはずが「{c.group(3) or '空'}」")
        k = _KIND_CELL_RE.match(str(l))
        if k:
            ex.append(f"{k.group(1)} は{k.group(2)}の「{k.group(3)}」のはずが{k.group(4)}になっています")
        if len(ex) >= 3:
            break
    size = next((_SIZE_RE.search(str(l)) for l in lines if _SIZE_RE.search(str(l))), None)
    if size:
        er, ec, gr, gc = (int(x) for x in size.groups())
        if er != gr:
            ex.insert(0, f"表の行数が違います（{er} 行のはずが {gr} 行）")
        elif ec != gc:
            ex.insert(0, f"表の列数が違います（{ec} 列のはずが {gc} 列）")
    if not ex:
        for l in lines:
            c = _CELL_RE.match(str(l))
            if c:
                ex.append(f"{c.group(1)} は空のはずが「{c.group(3)}」")
            if len(ex) >= 2:
                break
    return (f"{head}エラーは出ませんが結果が合いません。前は合っていました。例えば "
            + "・".join(ex[:3]) + "。直してください。")


def _fire(case, code, work):
    """壊した（か直した）マクロ 1 本を写しで撃つ。→ ([前の表の不一致], 秒)。"""
    import vbam_forge as fg
    bas = os.path.join(work, f"_mend_{int(time.time() * 1000)}.bas")
    err = fg._write_bas(bas, _MODULE, code)
    if err:
        return [f"マクロを読み込めませんでした: {err}"], 0.0
    try:
        res, sec = fg._run_on_copies(dict(case, prefire=[]), bas, case['sub'], [(case['before'], case['expect'])])
    finally:
        try:
            os.remove(bas)
        except OSError:
            pass
    return res[0], sec


def judge(case, code, work):
    """直ったマクロを前の表と別の表で撃つ。→ {'main': [...], 'tests': [(名, [...])], 'sec'}"""
    import vbam_forge as fg
    bas = os.path.join(work, f"_mend_fixed_{int(time.time() * 1000)}.bas")
    err = fg._write_bas(bas, _MODULE, code)
    if err:
        return {'main': [f"マクロを読み込めませんでした: {err}"], 'tests': [], 'sec': 0.0}
    targets = [(case['before'], case['expect'])] + [(t['before'], t['expect'], t.get('sheet'))
                                                    for t in (case.get('tests') or [])]
    try:
        res, sec = fg._run_on_copies(dict(case, prefire=[]), bas, case['sub'], targets)
        # 道具の側の止まり（写しを開けない・COM の例外）はマクロの外れではない＝1 回だけ撃ち直す（2026-09-17 夜・
        # 元どおりに直ったマクロが、並べて撃った別の回とぶつかって 1 枚だけ「シートが無い」で落ち、嘘の完了に数えられた）
        if any(r and str(r[0]).startswith(_INFRA_PREFIX) for r in res):
            res, sec2 = fg._run_on_copies(dict(case, prefire=[]), bas, case['sub'], targets)
            sec += sec2
    finally:
        try:
            os.remove(bas)
        except OSError:
            pass
    tests = [(t['label'], m) for t, m in zip(case.get('tests') or [], res[1:])]
    return {'main': res[0], 'tests': tests, 'sec': sec}


_INFRA_PREFIX = ('マクロの実行で止まりました: (-2147', 'マクロを読み込めませんでした', '前の表でマクロが止まったので')


def _fire_all(case, code, work):
    """壊したマクロを前の表と別の表の全部で撃つ。→ [(表の名前, before, sheet, 不一致の行)]、秒。"""
    import vbam_forge as fg
    bas = os.path.join(work, f"_mend_all_{int(time.time() * 1000)}.bas")
    err = fg._write_bas(bas, _MODULE, code)
    tables = [('前の表', case['before'], case['sheet'], case['expect'])] + [
        (t['label'], t['before'], t.get('sheet') or case['sheet'], t['expect']) for t in (case.get('tests') or [])]
    if err:
        return [(n, b, s, [f"マクロを読み込めませんでした: {err}"]) for n, b, s, _e in tables], 0.0
    try:
        res, sec = fg._run_on_copies(dict(case, prefire=[]), bas, case['sub'], [(b, e, s) for _n, b, s, e in tables])
    finally:
        try:
            os.remove(bas)
        except OSError:
            pass
    return [(n, b, s, r) for (n, b, s, _e), r in zip(tables, res)], sec


def pick_edge(case, rng, work, log=print, budget=12):
    """「いつもの表では合うのに、この表でだけ合わない」壊し方を 1 つ（前の表は正解どおり・別の表のどれかで外れる）。"""
    cands = [c for c in mutants(case['code']) if c['op'] in ('literal_typo', 'cmp_flip', 'drop_stmt', 'var_typo',
                                                              'offset', 'loop_end', 'loop_start')]
    rng.shuffle(cands)
    for k, c in enumerate(cands[:budget], 1):
        main, sec = _fire(case, c['code'], work)
        if main:
            log(f"    端の試し {k:>2}: {c['op']:<12} {c['line']:>3} 行 → 前の表でもう外れる（{sec:.1f} 秒）")
            continue
        rows, sec2 = _fire_all(case, c['code'], work)
        bad = [(n, b, s, r) for n, b, s, r in rows[1:] if r and classify(r) in ('wrong', 'nothing', 'runtime')]
        log(f"    端の試し {k:>2}: {c['op']:<12} {c['line']:>3} 行 → "
            + (f"別の表 {len(bad)} 枚で外れる" if bad else "どの表でも正解どおり") + f"（{sec + sec2:.1f} 秒）")
        if bad:
            n, b, s, r = bad[0]
            return dict({k2: v for k2, v in c.items()}, kind='edge', seen=classify(r), lines=r, before=b, sheet=s,
                        table=n, bad_tables=len(bad))
    return None


def pick_double(case, rng, work, log=print, budget=12):
    """壊れが 2 つ: 数が合わない壊し方＋その上にコンパイルエラー（人はまずコンパイルエラーに気づく）。"""
    first = pick(case, rng, ['wrong'], work, log=log, budget=budget)
    if not first:
        return None
    w = first[0]
    comp = [c for c in mutants(w['code']) if c['op'] in ('drop_end_if', 'drop_next', 'label_typo', 'member_typo')
            and c['line'] != w['line']]
    rng.shuffle(comp)
    for k, c in enumerate(comp[:budget], 1):
        lines, sec = _fire(case, c['code'], work)
        kind = classify(lines)
        log(f"    二つ目 {k:>2}: {c['op']:<12} {c['line']:>3} 行 → {_KIND_NAME.get(kind, kind or '正解どおり')}（{sec:.1f} 秒）")
        if kind == 'compile':
            return dict(c, kind='double', seen='compile', lines=lines, wrong_lines=w['lines'],
                        op=f"{w['op']}+{c['op']}", line=f"{w['line']}+{c['line']}",
                        old=f"{w['old']} ／ {c['old']}", new=f"{w['new'] or '（行を抜いた）'} ／ {c['new'] or '（行を抜いた）'}")
    return None


def pick(case, rng, kinds, work, log=print, budget=_TRIES_PER_CASE):
    """症状ごとに 1 つ、撃って確かめた壊し方を選ぶ。→ [{'kind','op','line','old','new','code','lines'}]"""
    cands = mutants(case['code'])
    rng.shuffle(cands)
    # 症状が出やすい壊し方を前に（足りない症状の分だけ・決まった順）
    likely = {'compile': ('drop_end_if', 'drop_next', 'label_typo', 'member_typo'),
              'runtime': ('set_drop', 'member_typo'),
              'nothing': ('literal_typo', 'cmp_flip'),
              'wrong': ('var_typo', 'loop_start', 'loop_end', 'offset', 'cmp_flip', 'drop_stmt', 'literal_typo')}
    got = {}
    tried = 0
    while tried < budget and any(k not in got for k in kinds):
        want = next(k for k in kinds if k not in got)
        pool = [c for c in cands if c['op'] in likely[want] and not c.get('_tried')]
        if not pool:
            pool = [c for c in cands if not c.get('_tried')]
        if not pool:
            break
        c = pool[0]
        c['_tried'] = True
        tried += 1
        lines, sec = _fire(case, c['code'], work)
        kind = classify(lines)
        log(f"    試し {tried:>2}: {c['op']:<12} {c['line']:>3} 行 → {_KIND_NAME.get(kind, kind or '正解どおり（お題にならない）')}"
            f"（{sec:.1f} 秒）")
        if kind in kinds and kind not in got:
            got[kind] = dict({k: v for k, v in c.items() if k != '_tried'}, kind=kind, lines=lines)
    return [got[k] for k in kinds if k in got]


# ----------------------------------------------------------------
# 直させる（別プロセス・自分の Excel・状態は置き場へ逃がす）
# ----------------------------------------------------------------

def _probe_source():
    """直させる子プロセスの台本。前の表の写し（この回の置き場）に、鍛えた台帳のマクロを壊したコードを入れる（注入経路。
    Sub 名・識別子は鍛えたときに規則と check-bas を通ったまま＝壊し方は字面の 1〜2 行だけ。人のブックには書かない）。"""
    return _PROBE.replace('__SCRIPTS__', SCRIPT_DIR)


_PROBE = r'''# -*- coding: utf-8 -*-
import io, json, os, sys, time, shutil, contextlib, traceback
sys.path.insert(0, r"__SCRIPTS__")
import vbam_core
vbam_core.setup_encoding()
import vba_manager, vbam_agent as va, vbam_shake as vs, vbam_macro as vmac

plan_path = sys.argv[1]
with open(plan_path, encoding="utf-8") as f:
    plan = json.load(f)
if plan.get("blind"):
    os.environ["VBAM_REPAIR_NO_CANON"] = "1"      # 正本を見せない＝台帳に無い手書きのマクロと同じ条件
else:
    _src = os.path.join(r"__SCRIPTS__", "_agent_forge.json")   # repair が正本との差分を出せるよう、逃がす先へ写す
    if os.path.isfile(_src):
        shutil.copy2(_src, os.path.join(plan["state"], "_agent_forge.json"))
vs._divert_state(plan["state"])
import pythoncom, win32com.client.dynamic
pythoncom.CoInitialize()
xl = win32com.client.dynamic.Dispatch(pythoncom.CoCreateInstanceEx(
    "Excel.Application", None, pythoncom.CLSCTX_SERVER, None, (pythoncom.IID_IDispatch,))[0])
try:
    import win32process
    _pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
except Exception:
    _pid = None
vbam_core._created_instances.append({"xl": xl, "pid": _pid})
xl.Visible = False
xl.DisplayAlerts = False

for item in plan["items"]:
    row = {"id": item["id"]}
    print("__START__" + item["id"], flush=True)
    buf = io.StringIO()
    t0 = time.time()
    wb = None
    try:
        shutil.copyfile(item["before"], item["work_xlsx"])
        wb = xl.Workbooks.Open(item["work_xlsx"], UpdateLinks=0)
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = item["module"]
        comp.CodeModule.AddFromString(item["code"])
        ws = wb.Worksheets(item["sheet"])
        ws.Activate()
        wb.SaveAs(item["book"], 52)
        vbam_core._wb_cache.clear()
        vbam_core._wb_cache[os.path.abspath(item["book"]).lower()] = (xl, wb, False)
        r, err = None, None
        with vbam_core.pinned_workbook(wb), contextlib.redirect_stdout(buf):
            try:
                r = vmac.run_macro_agent(item["request"], xl, wb, item["sub"], plan["ai"], plan.get("model"),
                                         max_turns=plan["max_turns"], rehearse=plan["rehearse"], ledger=False)
            except Exception as ex:
                err = "%s: %s" % (type(ex).__name__, ex)
                buf.write(traceback.format_exc())
        row["err"] = err
        if r is not None:
            row.update({"ok": bool(r.get("ok")), "done": bool(r.get("done")), "turns": r.get("turns"),
                        "gates": {k: v for k, v in (r.get("gates") or {}).items() if v},
                        "ai_sec": float((r.get("usage") or {}).get("sec") or 0),
                        "inv": r.get("inv") or [], "report": (r.get("report") or "")[:600]})
        cm = wb.VBProject.VBComponents(item["module"]).CodeModule
        n = int(cm.CountOfLines)
        row["code"] = str(cm.Lines(1, n)) if n else ""
        try:
            if os.path.isfile(va._LAST_AGENT_LOG_FILE):
                shutil.copy2(va._LAST_AGENT_LOG_FILE, item["log"])
        except Exception:
            pass
    except Exception as ex:
        row["err"] = row.get("err") or "%s: %s" % (type(ex).__name__, ex)
        row["trace"] = traceback.format_exc()[-1500:]
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        wb = None
    with open(item["out"], "w", encoding="utf-8") as f:
        f.write(buf.getvalue())
    row["sec"] = round(time.time() - t0, 1)
    print("__ROW__" + json.dumps(row, ensure_ascii=False, default=str), flush=True)

try:
    vbam_core.release_created_instances()
except Exception:
    pass
xl = None
print("__END__", flush=True)
'''


def _norm_lines(code):
    return [re.sub(r'\s+', ' ', l.strip()) for l in code.replace('\r\n', '\n').split('\n')
            if l.strip() and not l.strip().lower().startswith('attribute vb_')]


def changed_lines(before, after):
    """行の差の数（空行・字下げ・Attribute 行は数えない）。"""
    a, b = _norm_lines(before), _norm_lines(after)
    n = 0
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag != 'equal':
            n += max(i2 - i1, j2 - j1)
    return n


# ----------------------------------------------------------------
# 表の書き方と罫線と列幅をそろえる（表の整理.bas・約 700 行）も載せる（2026-09-17 夜）
#
# agent の先撃ちで毎回撃たれるいちばん大きいマクロ。鍛えた台帳には無いので、試験（vbam_exam）の 5 種の表を種で作り、
# 元のマクロで撃った後の姿を正解にする（元のマクロは試験で 50 題合格済み＝正解として使える）。
# ----------------------------------------------------------------

_TIDY_SUB = '表の書き方と罫線と列幅をそろえる'
_TIDY_SEED = 7


def _proc_of(text, name):
    """モジュールの本文から Sub name 〜 End Sub を抜く（純 Python）。無ければ None。"""
    lines = text.replace('\r\n', '\n').split('\n')
    s = next((i for i, l in enumerate(lines) if re.match(r'^\s*(?:Public\s+|Private\s+)?Sub\s+' + re.escape(name) + r'\s*\(', l)),
             None)
    if s is None:
        return None
    e = next((i for i in range(s + 1, len(lines)) if _END_SUB_RE.match(lines[i])), None)
    return None if e is None else '\n'.join(lines[s:e + 1])


def _capture(code, sub, books):
    """元のマクロを写しで撃ち、撃った後のシートの姿を読む（正解づくり）。books＝[(名, パス, シート)] → [snap]。"""
    import tempfile
    import contextlib
    import vbam_forge as fg
    bas = os.path.join(tempfile.gettempdir(), f"_mend_capture_{int(time.time() * 1000)}.bas")
    err = fg._write_bas(bas, _MODULE, code)
    if err:
        raise RuntimeError(err)
    xl, pid = fg._start_own()
    mwb = None
    macro_file = None
    out = []
    try:
        mwb = xl.Workbooks.Add()
        mwb.VBProject.VBComponents.Import(bas)
        h = mwb.VBProject.VBComponents.Add(1)
        h.Name = fg._HARNESS
        h.CodeModule.AddFromString(fg._harness_code([sub]))
        comp, macro_file = fg._compile_scratch(xl, mwb)
        if comp:
            raise RuntimeError(comp)
        for label, path, sheet in books:
            tmp = os.path.join(tempfile.gettempdir(), f"_mend_capture_{int(time.time() * 1000)}{os.path.splitext(path)[1]}")
            shutil.copyfile(path, tmp)
            wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
            try:
                ws = wb.Sheets(sheet)
                wb.Activate()
                ws.Activate()
                got = str(xl.Run(f"'{mwb.Name}'!{fg._HARNESS}.鍛冶_撃つ") or '')
                if got.startswith('ERR|'):
                    raise RuntimeError(f"{label}: 元のマクロが止まりました {got}")
                out.append(fg._snapshot_texts(ws))
            finally:
                with contextlib.suppress(Exception):
                    wb.Close(SaveChanges=False)
                with contextlib.suppress(OSError):
                    os.remove(tmp)
        return out
    finally:
        with contextlib.suppress(Exception):
            if mwb is not None:
                mwb.Close(SaveChanges=False)
        xl = None
        fg._quit(pid)
        for p in (macro_file, bas):
            if p:
                with contextlib.suppress(OSError):
                    os.remove(p)


def _tidy_case(seed=_TIDY_SEED, log=print):
    """表の書き方と罫線と列幅をそろえる を修理の試験に載せる台（置き場 _agent_mend\\_tidy_case_種 に作り、次からは読むだけ）。"""
    import vbam_exam as ve
    import vbam_forge as fg
    cdir = os.path.join(_MEND_DIR, f"_tidy_case_{seed}")
    cjson = os.path.join(cdir, 'case.json')
    if os.path.isfile(cjson):
        with open(cjson, encoding='utf-8') as f:
            return json.load(f)
    if not os.path.isfile(fg._TIDY_BAS):
        return None
    with open(fg._TIDY_BAS, encoding='cp932') as f:
        code = _proc_of(f.read(), _TIDY_SUB)
    if not code:
        return None
    os.makedirs(cdir, exist_ok=True)
    books = []
    for ex in ve.make_exams(seed):
        path = os.path.join(cdir, f"{ex['name']}.xlsx")
        ve.write_book(ex, path)
        books.append((ex['name'], path, ex['sheet']))
    log(f"  表の書き方と罫線と列幅をそろえる の台を作ります（試験の表 {len(books)} 枚・種 {seed}・元のマクロで撃った姿を正解にする）")
    snaps = _capture(code, _TIDY_SUB, books)
    (n0, p0, s0), e0 = books[0], snaps[0]
    case = {'name': _TIDY_SUB, 'sub': _TIDY_SUB, 'code': code, 'passed': True, 'prefire': [],
            'sheet': s0, 'before': p0, 'expect': e0,
            'tests': [{'label': n, 'before': p, 'sheet': s, 'expect': e} for (n, p, s), e in zip(books[1:], snaps[1:])]}
    with open(cjson, 'w', encoding='utf-8') as f:
        json.dump(case, f, ensure_ascii=False)
    return case


def _load_cases(only=None):
    from vbam_forge import _forge_load
    d = _forge_load()
    names = [n.strip() for n in (only or '').replace('、', ',').split(',') if n.strip()]
    out = []
    if _TIDY_SUB in names:                   # 表の書き方と罫線と列幅をそろえる は名指ししたときだけ（台づくりに Excel を 1 回起こす）
        tc = _tidy_case()
        if tc:
            out.append((_TIDY_SUB, tc))
    for name, case in d.items():
        if names and name not in names:
            continue
        if not case.get('passed') or not case.get('code') or not case.get('sub'):
            continue
        if case.get('prefire'):
            continue              # 先撃ち（表の書き方と罫線と列幅をそろえる）と組の弾は、植えるブックに 表の整理 全部が要る＝今回は外す
        if not os.path.isfile(case.get('before') or ''):
            continue
        out.append((name, case))
    return out


def _redo_items(src, only):
    """前の回の置き場（items.json）からお題をそのまま読む（--from）。only は題の id か マクロの名前（, 区切り）。"""
    path = src if src.lower().endswith('.json') else os.path.join(src, 'items.json')
    if not os.path.isfile(path):
        cand = os.path.join(_MEND_DIR, src, 'items.json')
        path = cand if os.path.isfile(cand) else path
    with open(path, encoding='utf-8') as f:
        items = json.load(f)
    names = [n.strip() for n in (only or '').replace('、', ',').split(',') if n.strip()]
    return [it for it in items if not names or it['id'] in names or it['name'] in names]


def mend_exam(seed=None, only=None, kinds=None, dry_run=False, ai=None, model=None, max_turns=4, rehearse=True,
              timeout=5400, redo=None, blind=False):
    """agent --mend の本体。戻り値は「全部のお題が直ったか」。redo＝前の回の置き場（同じお題で撃ち直す）。
    blind＝repair に鍛えた正本との差分を出さない（台帳に無い手書きのマクロと同じ条件で測る）。"""
    import vbam_agent as va
    if redo:
        return _mend_run(_redo_items(redo, only), None, dry_run, ai, model, max_turns, rehearse, timeout,
                         note=f"前の回のお題で撃ち直し（{redo}）", blind=blind)
    seed = int(seed) if seed not in (None, '') else random.randrange(1, 100000)
    kinds = [k.strip() for k in (kinds or ','.join(_KINDS)).replace('、', ',').split(',') if k.strip()]
    bad = [k for k in kinds if k not in _KINDS + _HARD_KINDS]
    if bad:
        print("エラー: 無い症状: " + "・".join(bad) + "（" + " / ".join(_KINDS + _HARD_KINDS) + "）")
        return False
    cases = _load_cases(only)
    if not cases:
        print("エラー: 撃てる鍛えたマクロがありません（agent --forged で一覧）")
        return False
    stamp = time.strftime('%Y%m%d_%H%M%S')
    vdir = os.path.join(_MEND_DIR, stamp)
    os.makedirs(vdir, exist_ok=True)
    print(f"修理の試験: 種 {seed}（同じ壊し方で撃ち直すなら agent --mend --seed {seed}）　置き場 {vdir}")
    rng = random.Random(seed)
    items = []
    t_pick = time.time()
    for name, case in cases:
        n_lines = len(_norm_lines(case['code']))
        print(f"\n■ {name}（{case['sub']}・{n_lines} 行・別の表 {len(case.get('tests') or [])} 枚）壊し方を撃って選びます")
        crng = random.Random(rng.randrange(1 << 30))
        basic = [k for k in kinds if k in _KINDS]
        picked = pick(case, crng, basic, vdir) if basic else []
        if 'edge' in kinds:
            e = pick_edge(case, crng, vdir)
            if e:
                picked.append(e)
        if 'double' in kinds:
            d = pick_double(case, crng, vdir)
            if d:
                picked.append(d)
        missing = [k for k in kinds if k not in {p['kind'] for p in picked}]
        if missing:
            print("  （見つからなかった症状: " + "・".join(_KIND_NAME[k] for k in missing) + "）")
        for p in picked:
            iid = f"{name}_{p['kind']}"
            sheet = p.get('sheet') or case['sheet']
            if p['kind'] == 'edge':
                req = symptom(p['seen'], p['lines'], case['sub'], sheet, edge=True)
            elif p['kind'] == 'double':
                more = symptom('wrong', p['wrong_lines'], case['sub'], sheet)
                more = more[more.find('例えば'):].replace('。直してください。', '')
                req = (symptom('compile', p['lines'], case['sub'], sheet).replace('直してください。', '')
                       + f"それと、エラーが出る前から結果も合っていませんでした。{more}。両方直してください。")
            else:
                req = symptom(p['kind'], p['lines'], case['sub'], sheet)
            print(f"  {_KIND_NAME[p['kind']]}: {p['op']} {p['line']} 行目「{p['old']}」→「{p['new'] or '（行を抜いた）'}」"
                  + (f"（{p['table']}・外れる表 {p['bad_tables']} 枚）" if p['kind'] == 'edge' else ''))
            print(f"    依頼: {req}")
            items.append({'id': iid, 'name': name, 'kind': p['kind'], 'op': p['op'], 'line': p['line'],
                          'old': p['old'], 'new': p['new'], 'code': p['code'], 'request': req,
                          'sub': case['sub'], 'sheet': sheet, 'module': _MODULE,
                          'before': p.get('before') or case['before'],
                          'work_xlsx': os.path.join(vdir, f"{iid}_前.xlsx"),
                          'book': os.path.join(vdir, f"{iid}.xlsm"),
                          'log': os.path.join(vdir, f"{iid}.log.jsonl"),
                          'out': os.path.join(vdir, f"{iid}.out.txt")})
    print(f"\n壊し方を選び終えました: {len(items)} 題（{time.time() - t_pick:.0f} 秒・AI なし）")
    with open(os.path.join(vdir, 'items.json'), 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=1)
    if dry_run or not items:
        print("（--dry-run: 壊し方を選んで写しで撃っただけ。AI には触っていません）" if dry_run else "エラー: お題が 0 題です")
        return bool(items)
    return _mend_run(items, seed, False, ai, model, max_turns, rehearse, timeout, vdir=vdir, blind=blind)


def _mend_run(items, seed, dry_run, ai, model, max_turns, rehearse, timeout, note='', vdir=None, blind=False):
    """選んだお題を子プロセスで直させ、写しで判定して台帳に残す。vdir が無ければ新しい置き場を作って道を付け替える。"""
    import vbam_agent as va
    if vdir is None:
        vdir = os.path.join(_MEND_DIR, time.strftime('%Y%m%d_%H%M%S'))
        os.makedirs(vdir, exist_ok=True)
        for it in items:
            it.update({'work_xlsx': os.path.join(vdir, f"{it['id']}_前.xlsx"), 'book': os.path.join(vdir, f"{it['id']}.xlsm"),
                       'log': os.path.join(vdir, f"{it['id']}.log.jsonl"), 'out': os.path.join(vdir, f"{it['id']}.out.txt")})
        with open(os.path.join(vdir, 'items.json'), 'w', encoding='utf-8') as f:
            json.dump(items, f, ensure_ascii=False, indent=1)
        print(f"修理の試験: {note}　{len(items)} 題　置き場 {vdir}")
        for it in items:
            print(f"  {it['id']}: {it['op']} {it['line']} 行目\n    依頼: {it['request']}")
    if not items:
        print("エラー: お題が 0 題です")
        return False
    if dry_run:
        print("（--dry-run: お題を並べただけ。AI には触っていません）")
        return True
    ai = ai or va._CC_AI
    state = os.path.join(vdir, '_state')
    os.makedirs(state, exist_ok=True)
    plan = {'state': state, 'ai': ai, 'model': model, 'max_turns': int(max_turns), 'rehearse': bool(rehearse),
            'blind': bool(blind), 'items': items}
    plan_path = os.path.join(vdir, '_plan.json')
    with open(plan_path, 'w', encoding='utf-8') as f:
        json.dump(plan, f, ensure_ascii=False, indent=1)
    probe = os.path.join(vdir, '_probe.py')
    with open(probe, 'w', encoding='utf-8') as f:
        f.write(_probe_source())
    print(f"直させます: {len(items)} 題・頭 {ai}{(' ' + model) if model else ''}（run_macro_agent・往復 {max_turns}"
          f"・rehearse {'あり' if rehearse else 'なし'}・正本の差分 {'見せない' if blind else '見せる'}）。Excel は自分の台（非表示）。台帳・控え・覚書は {state} へ")
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    proc = subprocess.Popen([sys.executable, probe, plan_path], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding='utf-8', errors='replace', env=env)
    by_id = {it['id']: it for it in items}
    cases_d = dict(_load_cases(','.join(sorted({it['name'] for it in items}))))
    rows, t0 = [], time.time()
    try:
        for line in proc.stdout:
            line = line.rstrip('\n')
            if line.startswith('__START__'):
                it = by_id.get(line[9:], {})
                print(f"\n  {line[9:]} を直させています…（{it.get('op')} {it.get('line')} 行目）", flush=True)
            elif line.startswith('__ROW__'):
                row = json.loads(line[7:])
                it = by_id[row['id']]
                case = cases_d[it['name']]
                fixed = row.get('code') or ''
                if fixed:
                    j = judge(case, fixed, vdir)
                else:
                    j = {'main': ['直したコードを読めませんでした'], 'tests': [], 'sec': 0.0}
                main_ok = not j['main']
                tests_ok = sum(1 for _n, m in j['tests'] if not m)
                row.update({'name': it['name'], 'kind': it['kind'], 'op': it['op'], 'line': it['line'],
                            'main_ok': main_ok, 'tests_ok': tests_ok, 'tests': len(j['tests']),
                            'pass': main_ok and tests_ok == len(j['tests']) and not row.get('err'),
                            'changed_vs_original': changed_lines(case['code'], fixed) if fixed else None,
                            'changed_vs_broken': changed_lines(it['code'], fixed) if fixed else None,
                            'main_why': j['main'][:4],
                            'tests_why': [(n, m[:2]) for n, m in j['tests'] if m][:3]})
                rows.append(row)
                print(_line(row), flush=True)
                if row.get('err'):
                    print(f"      エラー: {row['err']}")
                for w in row['main_why'][:3]:
                    print(f"      - 前の表: {w}")
                for n, ws_ in row['tests_why']:
                    print(f"      - {n}: {' / '.join(ws_)}")
            elif line.startswith('__END__'):
                pass
            elif line.strip():
                print("    | " + line, flush=True)
            if time.time() - t0 > timeout:
                proc.kill()
                print(f"エラー: {timeout} 秒を超えたので止めました")
                break
    finally:
        try:
            proc.wait(timeout=60)
        except Exception:
            proc.kill()
    if not rows:
        print("エラー: 1 題も返ってきませんでした（出力は上の行）")
        return False
    passed = [r for r in rows if r['pass']]
    lies = [r for r in rows if r.get('ok') and not r['pass']]
    secs = [r.get('sec', 0) for r in rows]
    print(f"\nまとめ（種 {seed}）: 直った {len(passed)}/{len(rows)} 題　嘘の完了 {len(lies)}"
          f"（AI が done・道具が ok と言ったのに値が合わない）　計 {sum(secs):.0f} 秒（1 題 平均 {sum(secs) / len(secs):.0f} 秒）")
    by_kind = {}
    for r in rows:
        by_kind.setdefault(r['kind'], []).append(r['pass'])
    print("  症状ごと: " + "　".join(f"{_KIND_NAME[k]} {sum(v)}/{len(v)}" for k, v in by_kind.items()))
    print(f"記録: {vdir}\\<題>.out.txt（往復の表示）・<題>.log.jsonl（往復の全文）")
    rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'seed': seed, 'ai': ai, 'model': model, 'dir': vdir,
           'max_turns': int(max_turns), 'rehearse': bool(rehearse), 'blind': bool(blind),
           'pass': len(passed), 'of': len(rows),
           'lies': len(lies), 'sec': round(sum(secs), 1),
           'items': [{k: r.get(k) for k in ('id', 'kind', 'op', 'line', 'pass', 'done', 'ok', 'turns', 'sec', 'ai_sec',
                                            'tests_ok', 'tests', 'changed_vs_original', 'changed_vs_broken', 'err',
                                            'main_why', 'gates')} for r in rows]}
    with open(os.path.join(vdir, 'report.json'), 'w', encoding='utf-8') as f:
        json.dump(rec, f, ensure_ascii=False, indent=1, default=str)
    with open(_MEND_SCORE_FILE, 'a', encoding='utf-8') as f:
        f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    return len(passed) == len(rows)


def _line(r):
    verdict = '直った' if r['pass'] else ('落ちた' if r.get('err') else '直っていない')
    lie = '  ← 嘘の完了' if (r.get('ok') and not r['pass']) else ''
    chg = r.get('changed_vs_original')
    chg_t = '-' if chg is None else ('元どおり' if chg == 0 else f"元と {chg} 行違う")
    return (f"  {r['id']:<18} {verdict:<6} 往復 {r.get('turns', '-')}  {r.get('sec', 0):5.0f} 秒  "
            f"done {'○' if r.get('done') else '×'}  別の表 {r.get('tests_ok', 0)}/{r.get('tests', 0)}  {chg_t}{lie}")


__all__ = ['mend_exam', 'mutants', 'classify', 'symptom', 'changed_lines', 'pick', 'judge']
