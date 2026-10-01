# -*- coding: utf-8 -*-
"""vbam_lookup.py — 列・行を挿入／削除したとき、VLOOKUP・HLOOKUP・INDEX の「直書きの番号」をずれないように直す（2026-10-01）。

Excel は列を挿入すると、式の中の**範囲（A1:D9）**は自動で伸ばすが、VLOOKUP の列番号（`,3,`）のような**直書きの数字**は
そのままにする。表の途中に列を入れると、番号が指す列が 1 つずれて、黙って別の値を返す。オフィス田中「Excel のエージェントモード」は
これを直せず、「長年の悩み」だった。道具が挿入・削除の前に影響を受ける式を読み、あとで同じデータを指すように番号を直す。

純 Python の部分（式を読む・番号を計算する）と、COM の部分（ブックの全式を読む・直す）に分かれる。
  scan_lookups(formula)            → 直書きの番号を持つ VLOOKUP/HLOOKUP/INDEX の一覧
  plan_adjust(item, own, target, axis, op, start, count) → (旧, 新, 状態)
  apply_values(formula, {k: 新})   → 番号だけ書き換えた式
  plan_for_book / apply_plan       → ブック全体（cmd_col・cmd_row が呼ぶ）
"""
import re

_FN_RE = re.compile(r'(?<![A-Za-z0-9_.])(VLOOKUP|HLOOKUP|INDEX)\s*\(', re.IGNORECASE)
_CAND_RE = re.compile(r'VLOOKUP|HLOOKUP|INDEX', re.IGNORECASE)

_FULL = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$')
_COLS = re.compile(r'^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$')
_ROWS = re.compile(r'^\$?(\d+):\$?(\d+)$')
_CELL = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')


def col_num(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _mask_strings(f):
    """"…" の中身を空白にする（コンマ・かっこを数えないため。長さは変えない）。"""
    out, inq = [], False
    for ch in f:
        if ch == '"':
            inq = not inq
            out.append(ch)
        else:
            out.append(' ' if inq else ch)
    return ''.join(out)


def _split_args(f, masked, open_pos):
    """f[open_pos] の '(' から対応する ')' までの引数を [(文字, 開始, 終了)] で（かっこ・波かっこ・'…' の中のコンマは切らない）。"""
    depth, i, start, args, n = 0, open_pos, open_pos + 1, [], len(f)
    while i < n:
        ch = masked[i]
        if ch == "'":                                    # シート名の '…'（'' は 1 つの '）
            i += 1
            while i < n:
                if masked[i] == "'":
                    if i + 1 < n and masked[i + 1] == "'":
                        i += 2
                        continue
                    break
                i += 1
        elif ch in '({':
            depth += 1
        elif ch in ')}':
            depth -= 1
            if depth == 0:
                args.append((f[start:i], start, i))
                return args
        elif ch == ',' and depth == 1:
            args.append((f[start:i], start, i))
            start = i + 1
        i += 1
    return None                                          # かっこが閉じない（式として壊れている）


def parse_ref(text):
    """範囲の文字 → (シート名 or None, c1, c2, r1, r2)。読めなければ None。列だけ・行だけの参照は片方が None。"""
    t = text.strip()
    sheet = None
    m = re.match(r"^(?:'((?:[^']|'')+)'|([^'!\s(),:]+))!(.*)$", t)
    if m:
        sheet = (m.group(1) or '').replace("''", "'") or m.group(2)
        t = m.group(3).strip()
    mm = _FULL.match(t)
    if mm:
        return sheet, col_num(mm.group(1)), col_num(mm.group(3)), int(mm.group(2)), int(mm.group(4))
    mm = _COLS.match(t)
    if mm:
        return sheet, col_num(mm.group(1)), col_num(mm.group(2)), None, None
    mm = _ROWS.match(t)
    if mm:
        return sheet, None, None, int(mm.group(1)), int(mm.group(2))
    mm = _CELL.match(t)
    if mm:
        c, r = col_num(mm.group(1)), int(mm.group(2))
        return sheet, c, c, r, r
    return None


def _lit(arg_text):
    t = arg_text.strip()
    return int(t) if re.fullmatch(r'\d+', t) else None


def scan_lookups(formula):
    """式の中の、直書きの番号を持つ VLOOKUP（列番号）・HLOOKUP（行番号）・INDEX（行・列番号）を位置順に返す。

    [{'fn', 'kind': 'col'|'row', 'sheet', 'c1','c2','r1','r2', 'val', 'pos': (開始, 終了)}]。範囲が読めない（名前・INDIRECT・表の構造化参照）は返さない。"""
    s = formula or ''
    masked = _mask_strings(s)
    out = []
    for m in _FN_RE.finditer(masked):
        fn = m.group(1).upper()
        args = _split_args(s, masked, m.end() - 1)
        if not args:
            continue
        if fn == 'VLOOKUP' and len(args) >= 3:
            ref, idx, kind = parse_ref(args[1][0]), args[2], 'col'
            lit = _lit(idx[0])
            if ref and lit and lit >= 1:
                out.append({'fn': fn, 'kind': kind, 'sheet': ref[0], 'c1': ref[1], 'c2': ref[2], 'r1': ref[3], 'r2': ref[4],
                            'val': lit, 'pos': _lit_span(s, idx)})
        elif fn == 'HLOOKUP' and len(args) >= 3:
            ref, idx, kind = parse_ref(args[1][0]), args[2], 'row'
            lit = _lit(idx[0])
            if ref and lit and lit >= 1:
                out.append({'fn': fn, 'kind': kind, 'sheet': ref[0], 'c1': ref[1], 'c2': ref[2], 'r1': ref[3], 'r2': ref[4],
                            'val': lit, 'pos': _lit_span(s, idx)})
        elif fn == 'INDEX' and len(args) >= 2:
            ref = parse_ref(args[0][0])
            if not ref:
                continue
            rows_only_one = ref[3] is not None and ref[3] == ref[4] and ref[1] != ref[2]
            if len(args) == 2:                            # INDEX(横 1 行の範囲, n) の n は列
                kinds = [('col' if rows_only_one else 'row', args[1])]
            else:
                kinds = [('row', args[1]), ('col', args[2])]
            for kind, a in kinds:
                lit = _lit(a[0])
                if lit and lit >= 1:                      # 0（行・列まるごと）は直さない
                    out.append({'fn': fn, 'kind': kind, 'sheet': ref[0], 'c1': ref[1], 'c2': ref[2], 'r1': ref[3], 'r2': ref[4],
                                'val': lit, 'pos': _lit_span(s, a)})
    out.sort(key=lambda it: it['pos'][0])
    return out


def _lit_span(s, arg):
    """引数の文字（前後の空白を除いた数字）の [開始, 終了) を式の中の位置で。"""
    text, start, _end = arg
    lead = len(text) - len(text.lstrip())
    return (start + lead, start + lead + len(text.strip()))


def _same_sheet(item_sheet, own_sheet, target_sheet):
    return (item_sheet if item_sheet is not None else own_sheet).lower() == str(target_sheet).lower()


def plan_adjust(item, own_sheet, target_sheet, axis, op, start, count):
    """1 つの式の番号が、(axis 'col'|'row') の挿入／削除（op 'insert'|'delete'・start 番目から count 個）でどうなるか。

    → None（影響なし）／(旧, 新, 'ok')／(旧, None, 'broken')＝指していた列・行そのものが消える。
    Excel は範囲を自動で伸ばす・縮める。範囲の中（先頭の列・行より後ろ・末尾まで）に入れたとき、指していた列・行より左・上なら 1 つずれる。"""
    if item['kind'] != axis or not _same_sheet(item['sheet'], own_sheet, target_sheet):
        return None
    lo, hi = (item['c1'], item['c2']) if axis == 'col' else (item['r1'], item['r2'])
    if lo is None or hi is None:
        return None
    val = item['val']
    target = lo + val - 1                                 # いま指している列・行
    if op == 'insert':
        if lo < start <= hi and start <= target:
            return (val, val + count, 'ok')
        return None
    end = start + count - 1                               # delete
    if end < lo or start > hi:
        return None
    if start <= target <= end:
        return (val, None, 'broken')
    before = max(0, min(end, target - 1) - max(start, lo) + 1)
    return (val, val - before, 'ok') if before > 0 else None


def apply_values(formula, new_by_k):
    """scan_lookups の k 番目の番号を new_by_k[k] に書き換えた式（位置は後ろから直す）。"""
    items = scan_lookups(formula)
    out = formula
    for k in sorted(new_by_k, reverse=True):
        if k >= len(items):
            continue
        a, b = items[k]['pos']
        out = out[:a] + str(new_by_k[k]) + out[b:]
    return out


def plan_formula(formula, own_sheet, target_sheet, axis, op, start, count):
    """1 つの式 → ({k: (旧, 新)}, 壊れる数)。影響が無ければ ({}, 0)。"""
    changes, broken = {}, 0
    for k, it in enumerate(scan_lookups(formula)):
        r = plan_adjust(it, own_sheet, target_sheet, axis, op, start, count)
        if not r:
            continue
        if r[2] == 'broken':
            broken += 1
        else:
            changes[k] = (r[0], r[1])
    return changes, broken


# ----------------------------------------------------------------
# 近似一致の検索（第 4 引数の省略・TRUE）
# ----------------------------------------------------------------
# オフィス田中「VLOOKUP の第 4 引数に毎回 FALSE を指定する意味は？」: 第 4 引数は省略すると TRUE（近似一致）になる。
# 近似一致は「並べ替えた表から、探す値以下で最大のものを返す」検索で、完全一致のつもりで省略すると、見つからない値に
# 黙って別の行の値を返す（エラーにならない）。MATCH の第 3 引数も同じ（省略は 1＝昇順の近似一致）。
_APPROX_FN_RE = re.compile(r'(?<![A-Za-z0-9_.])(VLOOKUP|HLOOKUP|MATCH)\s*\(', re.IGNORECASE)


def approx_calls(formula):
    """式の中の、近似一致になっている VLOOKUP・HLOOKUP・MATCH → [(関数名, 'omitted'|'TRUE'|'1'|'-1', 位置)]（純 Python）。

    VLOOKUP・HLOOKUP は第 4 引数が省略か TRUE・1。MATCH は第 3 引数が省略か 1・-1（0 だけが完全一致）。
    引数が空（`VLOOKUP(a,b,2,)`）は 0＝FALSE と同じなので完全一致。文字列の中は見ない。"""
    s = formula or ''
    masked = _mask_strings(s)
    out = []
    for m in _APPROX_FN_RE.finditer(masked):
        fn = m.group(1).upper()
        args = _split_args(s, masked, m.end() - 1)
        if not args:
            continue
        need = 4 if fn in ('VLOOKUP', 'HLOOKUP') else 3
        if len(args) < need:
            if len(args) >= 2:
                out.append((fn, 'omitted', m.start()))
            continue
        flag = args[need - 1][0].strip().upper()
        if fn == 'MATCH':
            if flag in ('1', '-1'):
                out.append((fn, flag, m.start()))
        elif flag in ('TRUE', '1'):
            out.append((fn, flag, m.start()))
    return out


def approx_note(cells):
    """[(番地, 書いた式)] → 近似一致の検索を書いたときの注意（無ければ ''・純 Python）。write-range・write-cells が使う。"""
    hit = []
    for a, v in cells:
        if isinstance(v, str) and v.startswith('=') and approx_calls(v):
            hit.append(a)
    if not hit:
        return ''
    return (f"⚠ 近似一致の検索を書きました（VLOOKUP/HLOOKUP の第 4 引数が省略か TRUE・MATCH の第 3 引数が省略か 1・-1）: {' '.join(hit[:8])}"
            + ("…" if len(hit) > 8 else "")
            + "。探す値が表に無いとき、エラーにならず別の行の値を返します（表が昇順に並んでいる前提の「以下で最大」の検索）。"
            "完全一致のつもりなら、第 4 引数に FALSE（MATCH は第 3 引数に 0）を書いてください")


# ----------------------------------------------------------------
# COM（ブック全体）
# ----------------------------------------------------------------

def _formulas_of(ws):
    """[(行, 列, 式)]。ws の式のセルを大きなまとまりで読む（Formula2 があればそれ＝動的配列の式を壊さない）。"""
    try:
        rng = ws.UsedRange.SpecialCells(-4123)
    except Exception:
        return []
    out = []
    for area in rng.Areas:
        try:
            vals = area.Formula2
        except Exception:
            vals = area.Formula
        r0, c0 = area.Row, area.Column
        if isinstance(vals, str):
            vals = ((vals,),)
        for i, row in enumerate(vals):
            for j, f in enumerate(row):
                if isinstance(f, str) and f.startswith('=') and _CAND_RE.search(f):
                    out.append((r0 + i, c0 + j, f))
    return out


def plan_for_book(wb, target_sheet, axis, op, start, count):
    """ブックの全式を読み、操作の影響を受ける式の計画を返す → [{'sheet','row','col','changes','broken'}]。"""
    plan = []
    for ws in wb.Worksheets:
        try:
            own = str(ws.Name)
            cells = _formulas_of(ws)
        except Exception:
            continue
        for r, c, f in cells:
            changes, broken = plan_formula(f, own, target_sheet, axis, op, start, count)
            if changes or broken:
                plan.append({'sheet': own, 'row': r, 'col': c, 'changes': changes, 'broken': broken, 'formula': f})
    return plan


def _moved(entry, target_sheet, axis, op, start, count):
    """操作のあとの、式のセルの (行, 列)。消えるなら None。"""
    r, c = entry['row'], entry['col']
    if entry['sheet'].lower() != str(target_sheet).lower():
        return r, c
    pos = c if axis == 'col' else r
    end = start + count - 1
    if op == 'insert':
        pos = pos + count if pos >= start else pos
    else:
        if start <= pos <= end:
            return None
        pos = pos - count if pos > end else pos
    return (r, pos) if axis == 'col' else (pos, c)


def apply_plan(wb, plan, target_sheet, axis, op, start, count):
    """操作のあとで、計画の式の番号を直す → (直した [(シート, 番地, 旧式, 新式)], 直せなかった [(シート, 番地, 理由)])。"""
    fixed, skipped = [], []
    for e in plan:
        if not e['changes']:
            if e['broken']:
                pos = _moved(e, target_sheet, axis, op, start, count)
                if pos:
                    skipped.append((e['sheet'], _addr(*pos), '指していた列・行そのものが消えたので直せません'))
            continue
        pos = _moved(e, target_sheet, axis, op, start, count)
        if pos is None:
            continue
        try:
            ws = wb.Worksheets(e['sheet'])
            cell = ws.Cells(pos[0], pos[1])
            if getattr(cell, 'HasArray', False):
                skipped.append((e['sheet'], _addr(*pos), '配列数式（Ctrl+Shift+Enter）は触りません'))
                continue
            try:
                now = cell.Formula2
                prop = 'Formula2'
            except Exception:
                now = cell.Formula
                prop = 'Formula'
            new = apply_values(now, {k: v[1] for k, v in e['changes'].items()})
            if new != now:
                setattr(cell, prop, new)
                fixed.append((e['sheet'], _addr(*pos), now, new))
            if e['broken']:
                skipped.append((e['sheet'], _addr(*pos), '一部の検索が指していた列・行そのものが消えたので直せません'))
        except Exception as ex:
            skipped.append((e['sheet'], _addr(*pos), f'直せませんでした（{ex}）'))
    return fixed, skipped


def _addr(r, c):
    s, n = '', c
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return f'{s}{r}'


def report_lines(fixed, skipped, what, limit=12):
    """cmd_col・cmd_row が出す文。"""
    out = []
    if fixed:
        out.append(f'番号の自動調整: {what}で、同じデータを指すように VLOOKUP/HLOOKUP/INDEX の直書きの番号を {len(fixed)} 式直しました'
                   '（--keep-lookup を付けると直さない）')
        for sh, addr, old, new in fixed[:limit]:
            out.append(f'  {sh}!{addr}: {old}  →  {new}')
        if len(fixed) > limit:
            out.append(f'  … 他 {len(fixed) - limit} 式')
    if skipped:
        out.append(f'⚠ 直せなかった式 {len(skipped)} 件（人が見てください）:')
        for sh, addr, why in skipped[:limit]:
            out.append(f'  {sh}!{addr}: {why}')
    return out


__all__ = ['scan_lookups', 'plan_adjust', 'plan_formula', 'apply_values', 'plan_for_book', 'apply_plan', 'report_lines', 'parse_ref',
           'approx_calls', 'approx_note']
