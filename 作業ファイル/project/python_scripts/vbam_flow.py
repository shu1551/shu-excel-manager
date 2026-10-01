# -*- coding: utf-8 -*-
"""vbam_flow.py — マクロの中の流れ（分岐・ループ・GoSub・エラー処理）を Mermaid の流れ図にする（2026-10-01）。

オフィス田中「VBAのコードから簡単にフロー図を作る」（属人化の解消）の取り入れ。動画では AI にコードを渡して図を書かせるが、
**図の元になる流れはコードから機械的に取れる**ので、道具が AI なしで書く（読み違い・作り話が混ざらない・何度でも同じ図）。

  flow [excel_file] <マクロ名> [--module 名] [--out 出力.md] [--max-nodes 150]

純 Python（COM なし）の本体は flow_to_mermaid(code)。If／ElseIf／Else・For／For Each・Do／Loop・While・Select Case・With・
GoTo・GoSub／Return・On Error GoTo・Exit Sub／For／Do・ラベル（GoSub の飛び先・エラー処理の節は別の枠）を扱う。
図の言葉は、各ブロックの直前のコメント（なければコードの先頭の数行）。節（Dim など宣言）は図に載せない。
"""
import os
import re
import sys

_ID = r'[^\W\d]\w*'
_LABEL_RE = re.compile(rf'^({_ID}):(?!=)\s*(.*)$')
_NOT_LABELS = {'else', 'case', 'default'}


# ----------------------------------------------------------------
# 1. 行を「文」に切る
# ----------------------------------------------------------------

def _strip_comment(line):
    """文字列リテラルの外の ' 以降をコメントとして切る → (コード, コメント)。"""
    inq = False
    for i, ch in enumerate(line):
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            return line[:i], line[i + 1:].strip()
    return line, ''


def _split_colon(code):
    """文字列の外・:= の外のコロンで 1 行の複数文を切る。"""
    parts, cur, inq, i = [], [], False, 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            inq = not inq
        if ch == ':' and not inq and not (i + 1 < len(code) and code[i + 1] == '='):
            parts.append(''.join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    parts.append(''.join(cur))
    return [p.strip() for p in parts if p.strip()]


def _items(code):
    """プロシージャのコードを [{'t': 文, 'note': 直前のコメント, 'line': 行}] に（継続行は畳む）。"""
    raw = code.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    items, pend, i = [], [], 0
    while i < len(raw):
        s = raw[i].strip()
        lineno = i + 1
        if not s:
            pend = []
            i += 1
            continue
        if s.startswith("'") or re.match(r'(?i)rem(\s|$)', s):
            pend.append(re.sub(r"^('|(?i:rem))\s*", '', s).strip())
            i += 1
            continue
        body, cmt = _strip_comment(raw[i])
        body = body.strip()
        while (body.endswith(' _') or body == '_') and i + 1 < len(raw):
            i += 1
            nxt, c2 = _strip_comment(raw[i])
            body = body[:-1].rstrip() + ' ' + nxt.strip()
            cmt = cmt or c2
        note = next((c for c in pend if c), '') or cmt
        pend = []
        stmts = []
        # 1 行の If … Then 文 … はコロンで切らない（Then の後ろが全部その If の中身）
        if re.match(r'(?i)^if\s+.+?\s+then\s+\S', body):
            stmts = [body]
        else:
            lm = _LABEL_RE.match(body)
            if lm and lm.group(1).lower() not in _NOT_LABELS:
                stmts.append(lm.group(1) + ':')
                body = lm.group(2)
            stmts += _split_colon(body)
        for k, st in enumerate(stmts):
            items.append({'t': st, 'note': note if k == 0 else '', 'line': lineno})
        i += 1
    return items


# ----------------------------------------------------------------
# 2. 文の種類
# ----------------------------------------------------------------

def _classify(t):
    s = t.strip()
    low = s.lower()
    if re.match(r'^(public\s+|private\s+|friend\s+)?(static\s+)?(sub|function|property\s+(get|let|set))\s', low):
        return 'proc_start', s
    if re.match(r'^end\s+(sub|function|property)\b', low):
        return 'proc_end', s
    m = re.match(r'(?i)^if\s+(.+?)\s+then$', s)
    if m:
        return 'if_block', m.group(1)
    m = re.match(r'(?i)^elseif\s+(.+?)\s+then$', s)
    if m:
        return 'elseif', m.group(1)
    if low == 'else':
        return 'else', ''
    if re.match(r'^end\s+if$', low):
        return 'endif', ''
    m = re.match(r'(?i)^if\s+(.+?)\s+then\s+(\S.*)$', s)
    if m:
        return 'single_if', (m.group(1), m.group(2))
    if re.match(r'(?i)^for\s', s):
        return 'for', s
    if re.match(r'(?i)^next\b', s):
        return 'next', ''
    m = re.match(r'(?i)^do(\s+(while|until)\s+(.+))?$', s)
    if m:
        return 'do', (m.group(2) or '', m.group(3) or '')
    m = re.match(r'(?i)^loop(\s+(while|until)\s+(.+))?$', s)
    if m:
        return 'loop', (m.group(2) or '', m.group(3) or '')
    m = re.match(r'(?i)^while\s+(.+)$', s)
    if m:
        return 'while', m.group(1)
    if low == 'wend':
        return 'wend', ''
    m = re.match(r'(?i)^select\s+case\s+(.+)$', s)
    if m:
        return 'select', m.group(1)
    m = re.match(r'(?i)^case\s+(.+)$', s)
    if m:
        return 'case', m.group(1)
    if re.match(r'^end\s+select$', low):
        return 'endselect', ''
    if re.match(r'(?i)^with\s', s):
        return 'with', s
    if re.match(r'^end\s+with$', low):
        return 'endwith', ''
    m = re.match(r'(?i)^exit\s+(sub|function|property)$', s)
    if m:
        return 'exit_proc', s
    m = re.match(r'(?i)^exit\s+(for|do)$', s)
    if m:
        return 'exit_loop', s
    m = re.match(r'(?i)^gosub\s+(\S+)$', s)
    if m:
        return 'gosub', m.group(1)
    m = re.match(r'(?i)^goto\s+(\S+)$', s)
    if m:
        return 'goto', m.group(1)
    if low == 'return':
        return 'return', ''
    m = re.match(r'(?i)^on\s+error\s+goto\s+(\S+)$', s)
    if m:
        return ('simple_skip', s) if m.group(1) in ('0', '-1') else ('onerror', m.group(1))
    m = re.match(rf'^({_ID}):$', s)
    if m and m.group(1).lower() not in _NOT_LABELS:
        return 'label', m.group(1)
    if re.match(r'(?i)^(dim|const|static|option|redim|private\s+const|public\s+const)\b', s):
        return 'decl', s
    return 'simple', s


# ----------------------------------------------------------------
# 3. 構文木
# ----------------------------------------------------------------

class _Parse:
    def __init__(self, items):
        self.items = items
        self.i = 0

    def peek(self):
        return _classify(self.items[self.i]['t']) if self.i < len(self.items) else (None, None)


def _block(p, stop):
    """stop の種類の文に出会うまでを 1 つの並びにする（出会った文は消費しない）。"""
    body = []
    while p.i < len(p.items):
        kind, data = p.peek()
        if kind in stop:
            break
        it = p.items[p.i]
        p.i += 1
        note = it['note']
        if kind in ('proc_start', 'proc_end', 'decl', 'simple_skip', 'endif', 'next', 'loop', 'wend', 'endselect', 'endwith',
                    'elseif', 'else', 'case'):
            continue                                    # 構造の終わりの余り・宣言は図に載せない
        if kind == 'if_block':
            branches = [(data, note, _block(p, {'elseif', 'else', 'endif'}))]
            els = None
            while True:
                k2, d2 = p.peek()
                if k2 == 'elseif':
                    n2 = p.items[p.i]['note']
                    p.i += 1
                    branches.append((d2, n2, _block(p, {'elseif', 'else', 'endif'})))
                elif k2 == 'else':
                    p.i += 1
                    els = _block(p, {'endif'})
                else:
                    break
            if p.peek()[0] == 'endif':
                p.i += 1
            body.append({'k': 'if', 'branches': branches, 'else': els, 'line': it['line']})
        elif kind == 'single_if':
            cond, act = data
            then_s, else_s = act, None
            m = re.match(r'(?is)^(.*?)\s+else\s+(.*)$', act)
            if m and '"' not in m.group(1):
                then_s, else_s = m.group(1), m.group(2)
            sub = _Parse([{'t': x, 'note': '', 'line': it['line']} for x in _split_colon(then_s)])
            sub_else = _Parse([{'t': x, 'note': '', 'line': it['line']} for x in _split_colon(else_s)]) if else_s else None
            body.append({'k': 'if', 'branches': [(cond, note, _block(sub, set()))],
                         'else': _block(sub_else, set()) if sub_else else None, 'line': it['line']})
        elif kind == 'for':
            inner = _block(p, {'next'})
            if p.peek()[0] == 'next':
                p.i += 1
            body.append({'k': 'loop', 'head': data, 'note': note, 'body': inner, 'post': False, 'line': it['line']})
        elif kind == 'while':
            inner = _block(p, {'wend'})
            if p.peek()[0] == 'wend':
                p.i += 1
            body.append({'k': 'loop', 'head': 'While ' + data, 'note': note, 'body': inner, 'post': False, 'line': it['line']})
        elif kind == 'do':
            inner = _block(p, {'loop'})
            post = ('', '')
            if p.peek()[0] == 'loop':
                post = p.peek()[1]
                p.i += 1
            if data[0]:
                head, post_flag = f'Do {data[0]} {data[1]}', False
            elif post[0]:
                head, post_flag = f'Loop {post[0]} {post[1]}', True
            else:
                head, post_flag = 'Do … Loop（Exit Do まで）', False
            body.append({'k': 'loop', 'head': head, 'note': note, 'body': inner, 'post': post_flag, 'line': it['line']})
        elif kind == 'select':
            cases = []
            while p.peek()[0] not in ('endselect', None):
                k2, d2 = p.peek()
                if k2 == 'case':
                    n2 = p.items[p.i]['note']
                    p.i += 1
                    cases.append((d2, n2, _block(p, {'case', 'endselect'})))
                else:
                    p.i += 1                           # 最初の Case の前の余り
            if p.peek()[0] == 'endselect':
                p.i += 1
            body.append({'k': 'select', 'expr': data, 'note': note, 'cases': cases, 'line': it['line']})
        elif kind == 'with':
            inner = _block(p, {'endwith'})
            if p.peek()[0] == 'endwith':
                p.i += 1
            if note and inner and inner[0].get('k') == 'simple':
                inner[0]['note'] = inner[0].get('note') or note
            body.extend(inner)
        elif kind in ('exit_proc', 'exit_loop', 'gosub', 'goto', 'return', 'onerror', 'label'):
            body.append({'k': kind, 'data': data, 'note': note, 'line': it['line']})
        else:                                           # simple
            if body and body[-1]['k'] == 'simple' and not note:
                body[-1]['lines'].append(data)
            else:
                body.append({'k': 'simple', 'lines': [data], 'note': note, 'line': it['line']})
    return body


def _size(nodes):
    n = 0
    for x in nodes:
        k = x['k']
        if k == 'simple':
            n += len(x['lines'])
        elif k == 'if':
            n += 1 + sum(_size(b[2]) for b in x['branches']) + (_size(x['else']) if x['else'] else 0)
        elif k == 'loop':
            n += 1 + _size(x['body'])
        elif k == 'select':
            n += 1 + sum(_size(c[2]) for c in x['cases'])
        else:
            n += 1
    return n


def _walk(nodes):
    for x in nodes:
        yield x
        k = x['k']
        if k == 'if':
            for b in x['branches']:
                yield from _walk(b[2])
            if x['else']:
                yield from _walk(x['else'])
        elif k == 'loop':
            yield from _walk(x['body'])
        elif k == 'select':
            for c in x['cases']:
                yield from _walk(c[2])


# ----------------------------------------------------------------
# 4. Mermaid
# ----------------------------------------------------------------

def _esc(s):
    s = str(s).replace('#', '#35;').replace('"', '#quot;').replace('<', '#lt;').replace('>', '#gt;').replace('&', '#amp;')
    return s.replace('|', '#124;').replace('`', "'")


def _clip(s, n=36):
    s = ' '.join(str(s).split())
    return s if len(s) <= n else s[:n - 1] + '…'


class _Emit:
    def __init__(self, collapse):
        self.n = 0
        self.defs = []                  # 今の枠（本体か節）の節点の定義
        self.edges = []
        self.labels = {}                # ラベル名 → 節点
        self.pending = []               # (節点, ラベル名, 種類)
        self.loops = []                 # Exit For/Do の集め先
        self.collapse = collapse
        self.end_used = False
        self.stats = {'nodes': 0, 'decisions': 0, 'loops': 0, 'gosubs': 0}

    def node(self, shape, text):
        self.n += 1
        nid = f'n{self.n}'
        self.stats['nodes'] += 1
        t = _esc(text) if '<br/>' not in text else '<br/>'.join(_esc(x) for x in text.split('<br/>'))
        o, c = {'box': ('["', '"]'), 'dec': ('{"', '"}'), 'loop': ('{{"', '"}}'), 'sub': ('[["', '"]]'),
                'stad': ('(["', '"])')}[shape]
        self.defs.append(f'    {nid}{o}{t}{c}')
        return nid

    def edge(self, a, b, label='', dotted=False):
        arrow = '-.->' if dotted else '-->'
        self.edges.append(f'    {a} {arrow}|"{_esc(label)}"| {b}' if label else f'    {a} {arrow} {b}')

    def connect(self, incoming, nid):
        for src, lab in incoming:
            self.edge(src, nid, lab)

    def end_node(self):
        self.end_used = True
        return 'END'


def _simple_text(x):
    lines = [_clip(l, 34) for l in x['lines']]
    if x.get('note'):
        t = _clip(x['note'], 38)
        return t + (f'<br/>（{len(lines)} 行）' if len(lines) > 1 else '')
    t = '<br/>'.join(lines[:3])
    return t + (f'<br/>…他 {len(lines) - 3} 行' if len(lines) > 3 else '')


def _emit(em, nodes, incoming, depth):
    """nodes を図にして、つながっていない出口 [(節点, 辺の言葉)] を返す。"""
    outs = list(incoming)
    for x in nodes:
        k = x['k']
        if k in ('if', 'loop', 'select') and em.collapse is not None and depth >= em.collapse:
            nid = em.node('box', _clip(x.get('note') or _head_of(x), 30) + f'<br/>（中身 {_size([x])} 行は省略）')
            em.connect(outs, nid)
            outs = [(nid, '')]
            continue
        if k == 'simple':
            nid = em.node('box', _simple_text(x))
            em.connect(outs, nid)
            outs = [(nid, '')]
        elif k == 'if':
            outs = _emit_if(em, x, outs, depth)
        elif k == 'loop':
            outs = _emit_loop(em, x, outs, depth)
        elif k == 'select':
            sid = em.node('dec', _clip('Select Case ' + x['expr'], 40) + (f'<br/>（{_clip(x["note"], 26)}）' if x['note'] else ''))
            em.stats['decisions'] += 1
            em.connect(outs, sid)
            new_outs, has_else = [], False
            for cond, _n, body in x['cases']:
                is_else = cond.strip().lower() == 'else'
                has_else = has_else or is_else
                lab = 'それ以外' if is_else else _clip('Case ' + cond, 24)
                new_outs += _emit(em, body, [(sid, lab)], depth + 1)
            if not has_else:
                new_outs.append((sid, 'どれでもない'))
            outs = new_outs
        elif k == 'exit_proc':
            nid = em.node('box', _clip(x['data'], 20))
            em.connect(outs, nid)
            em.edge(nid, em.end_node())
            outs = []
        elif k == 'exit_loop':
            nid = em.node('box', _clip(x['data'], 20))
            em.connect(outs, nid)
            if em.loops:
                em.loops[-1].append((nid, '抜ける'))
            outs = []
        elif k == 'goto':
            nid = em.node('box', 'GoTo ' + x['data'])
            em.connect(outs, nid)
            em.pending.append((nid, x['data'], 'goto'))
            outs = []
        elif k == 'gosub':
            nid = em.node('sub', 'GoSub ' + x['data'] + (('<br/>' + _clip(x['note'], 30)) if x.get('note') else ''))
            em.stats['gosubs'] += 1
            em.connect(outs, nid)
            em.pending.append((nid, x['data'], 'gosub'))
            outs = [(nid, '')]
        elif k == 'return':
            nid = em.node('stad', 'Return（呼び元へ戻る）')
            em.connect(outs, nid)
            outs = []
        elif k == 'onerror':
            nid = em.node('box', 'エラーが起きたら ' + x['data'] + ' へ')
            em.connect(outs, nid)
            em.pending.append((nid, x['data'], 'onerror'))
            outs = [(nid, '')]
        elif k == 'label':
            nid = em.node('stad', x['data'] + ':')
            em.labels[x['data']] = nid
            em.connect(outs, nid)
            outs = [(nid, '')]
    return outs


def _head_of(x):
    if x['k'] == 'if':
        return 'If ' + x['branches'][0][0]
    if x['k'] == 'loop':
        return x['head']
    return 'Select Case ' + x['expr']


def _emit_if(em, x, outs, depth):
    all_outs = []
    cur = outs
    for idx, (cond, note, body) in enumerate(x['branches']):
        did = em.node('dec', _clip(cond, 40) + ' ?' + (f'<br/>（{_clip(note, 26)}）' if note else ''))
        em.stats['decisions'] += 1
        em.connect(cur, did)
        all_outs += _emit(em, body, [(did, 'はい')], depth + 1)
        cur = [(did, 'いいえ')]
    if x['else'] is not None:
        all_outs += _emit(em, x['else'], cur, depth + 1)
    else:
        all_outs += cur
    return all_outs


def _emit_loop(em, x, outs, depth):
    lid = em.node('loop', _clip(x['head'], 40) + (f'<br/>（{_clip(x["note"], 26)}）' if x.get('note') else ''))
    em.stats['loops'] += 1
    em.loops.append([])
    if x['post']:                                       # Do … Loop While 条件（先に中身・あとで判定）
        before = len(em.defs)
        body_outs = _emit(em, x['body'], outs, depth + 1)
        first = em.defs[before].split('[')[0].split('{')[0].split('(')[0].strip() if len(em.defs) > before else None
        em.connect(body_outs, lid)
        if first:
            em.edge(lid, first, '続ける')
    else:
        em.connect(outs, lid)
        body_outs = _emit(em, x['body'], [(lid, '続ける')], depth + 1)
        for src, lab in body_outs:
            em.edge(src, lid, lab)
    exits = em.loops.pop()
    return [(lid, '終わり')] + exits


def _sections(top):
    """トップの並びを、本体と「節」（GoSub の飛び先・エラー処理の飛び先）に分ける → (本体, [(ラベル, 並び)])。"""
    gosub_targets = {x['data'] for x in _walk(top) if x['k'] == 'gosub'}
    main, secs, cur, prev_terminal = [], [], None, False
    for x in top:
        if x['k'] == 'label' and (prev_terminal or x['data'] in gosub_targets):
            cur = (x['data'], [])
            secs.append(cur)
            prev_terminal = False
            continue
        (cur[1] if cur is not None else main).append(x)
        prev_terminal = x['k'] in ('exit_proc', 'return', 'goto')
    return main, secs


def flow_to_mermaid(code, title='', max_nodes=150):
    """プロシージャのコード → (Mermaid の本文, 数え {nodes, decisions, loops, gosubs, collapsed})。

    節点が max_nodes を超えるときは、深い入れ子を「（中身 n 行は省略）」の 1 つにたたんで収める。"""
    items = _items(code)
    if not any(_classify(i['t'])[0] == 'proc_start' for i in items):
        raise ValueError('Sub / Function の宣言が見つかりません')
    head = next(i for i in items if _classify(i['t'])[0] == 'proc_start')['t']
    name = title or (re.search(r'(?i)(?:sub|function|get|let|set)\s+([^\s(]+)', head) or [None, 'マクロ'])[1]
    p = _Parse(items)
    top = _block(p, set())
    main, secs = _sections(top)
    last = None
    for collapse in [None] + list(range(6, 0, -1)):
        em = _Emit(collapse)
        em.defs.append(f'    START(["開始: {_esc(name)}"])')
        main_outs = _emit(em, main, [('START', '')], 0)
        main_defs = em.defs
        sec_blocks = []
        for lab, body in secs:
            em.defs = []
            lid = em.node('stad', lab + ':')
            em.labels[lab] = lid
            s_outs = _emit(em, body, [(lid, '')], 0)
            sec_blocks.append((lab, em.defs, s_outs))
        for src, lab in main_outs:
            em.edge(src, em.end_node(), lab)
        for lab, defs, s_outs in sec_blocks:
            for src, l2 in s_outs:
                em.edge(src, em.end_node(), l2)
        for nid, target, kind in em.pending:
            dst = em.labels.get(target)
            if dst:
                em.edge(nid, dst, {'goto': '', 'gosub': '呼ぶ', 'onerror': 'エラー時'}[kind], dotted=(kind != 'goto'))
        last = (em, main_defs, sec_blocks, collapse)
        if em.stats['nodes'] <= max_nodes:
            break
    em, main_defs, sec_blocks, collapse = last
    out = ['flowchart TD']
    out += main_defs
    for k, (lab, defs, _o) in enumerate(sec_blocks, 1):
        out.append(f'    subgraph sec{k}["{_esc(lab)}: の節"]')
        out += ['    ' + d for d in defs]
        out.append('    end')
    if em.end_used:
        out.append('    END(["終了"])')
    out += em.edges
    stats = dict(em.stats)
    stats['collapsed'] = collapse
    return '\n'.join(out) + '\n', stats


def mermaid_markdown(name, mermaid, stats, source=''):
    note = ''
    if stats.get('collapsed'):
        note = f'節点が多いので、入れ子 {stats["collapsed"]} 段より深い所は「（中身 n 行は省略）」にたたみました。\n\n'
    return (f'# {name} の流れ図\n\n{source}節点 {stats["nodes"]}・分岐 {stats["decisions"]}・繰り返し {stats["loops"]}・GoSub {stats["gosubs"]}'
            f'（コードから機械的に作った図。AI は使っていません）\n\n{note}```mermaid\n{mermaid}```\n')


# ----------------------------------------------------------------
# 5. コマンド
# ----------------------------------------------------------------

def cmd_flow(args):
    """マクロの流れ図: flow [excel_file] <マクロ名> [--module 名] [--out 出力.md] [--max-nodes N]

    コードから分岐・ループ・GoSub・エラー処理の流れを機械的に読み、Mermaid の流れ図を出す（AI なし）。
    GitHub・Qiita・Mermaid Live（mermaid.live）にそのまま貼れる。ラベルは各ブロックの直前のコメント。"""
    import vbam_core as vc
    import vbam_vba as vv
    target_file, rest = vc.parse_target_and_rest(args.posargs)
    if not rest:
        print('使い方: flow [excel_file] <マクロ名> [--module 名] [--out 出力.md] [--max-nodes N]')
        return False
    macro = rest[0]
    module = getattr(args, 'module_opt', None)
    if len(rest) >= 2:
        module, macro = module or rest[0], rest[1]
    xl, wb = vc.get_workbook(target_file, readonly=True)
    try:
        comp, code = vv._extract_proc(wb, module, macro)
    except ValueError as multiple:
        print(f"エラー: '{macro}' が複数のモジュールにあります: {multiple}\n  --module で指定してください。")
        return False
    if not comp or not code:
        print(f"エラー: マクロ '{macro}' が見つかりません")
        return False
    try:
        mer, stats = flow_to_mermaid(code, macro, int(getattr(args, 'max_nodes', None) or 150))
    except ValueError as e:
        print(f'エラー: {e}')
        return False
    md = mermaid_markdown(macro, mer, stats, source=f'場所: {comp}.{macro}／')
    out = getattr(args, 'out', None) or os.path.join(os.path.dirname(os.path.abspath(__file__)), '_last_flow.md')
    out = os.path.abspath(out)
    with open(out, 'w', encoding='utf-8') as f:
        f.write(md)
    print(md)
    print(f'保存: {out}')
    return True


__all__ = ['flow_to_mermaid', 'mermaid_markdown', 'cmd_flow']
