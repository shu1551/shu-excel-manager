# -*- coding: utf-8 -*-
"""vbam_dock.py — dock: 点検の部品を 1 回で撃って 1 枚にまとめる入口（2026-11-05）。

新しい検査はしない。既にある部品を呼んで並べるだけ:
  骨・気をつける所 … vbam_structure.collect / points
  病巣と点数       … vbam_audit.check_formula_linter / check_data_cleaner / audit_score
  壊れやすい式     … vbam_lookup.scan_lookups / approx_calls（ブックの全式）
  エラーの原因     … vbam_cause.cause_notes
  式のつながり     … vbam_view._trace_walk（合計・小計らしい式の上位数個・深さ 2）
純 Python（まとめ・Markdown/HTML の組み立て）と COM（ブックを読む）を分けてある。
"""
import os
import re

MAX_CELLS_PER_SHEET = 200000
TRACE_TOP = 3
_TOTAL_RE = re.compile(r'^=\s*(?:SUM|SUBTOTAL|AGGREGATE)\s*\(', re.IGNORECASE)

_TYPE_LABEL = {
    'circular_reference': '循環参照', 'formula_error': '数式エラー', 'hardcoded_constant': '式に直書きの数',
    'approx_match': '近似一致', 'omitted_sum_range': '集計の範囲漏れ', 'inconsistent_formula': '隣と違う式',
    'value_in_formula_column': '式の列に値', 'hidden_whitespace': '隠れ空白', 'text_number': '文字列の数字',
    'outlier_value': '外れ値', 'outlier': '外れ値',
    'text_date': '文字の日付', 'duplicate_cf': '重複した条件付き書式', 'orphan_formula': '孤立した式',
    'merged_cells': '縦の結合セル', 'hidden_content': '隠した中身', 'three_d_reference': '3-D 参照',
    'external_link': '外部ブック参照', 'long_formula': '長い式',
}
_SEV = {'critical': '重大', 'warning': '警告', 'info': '気づき'}


def _md(s):
    return str(s).replace('|', '｜').replace('\r', '').replace('\n', '⏎')


# ----------------------------------------------------------------
# 純 Python
# ----------------------------------------------------------------

def lookup_findings(formulas):
    """[(シート, 番地, 式)] → ([ずれやすい番号の検索], [近似一致の検索])（どちらも [(シート, 番地, 関数, 内容)]）。"""
    import vbam_lookup as vl
    fragile, approx = [], []
    for sheet, addr, f in formulas:
        if not isinstance(f, str) or not f.startswith('='):
            continue
        for it in vl.scan_lookups(f):
            kind = '列番号' if it.get('kind') == 'col' else '行番号'
            fragile.append((sheet, addr, it['fn'], f"{kind} {it.get('val')} を直書き"))
        for fn, flag, _pos in vl.approx_calls(f):
            approx.append((sheet, addr, fn, '検索方法を省略' if flag == 'omitted' else f'検索方法が {flag}'))
    return fragile, approx


def _span(f):
    import vbam_lookup as vl
    n = 0
    for m in re.finditer(r'\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)', f.upper()):
        n += (abs(vl.col_num(m.group(3)) - vl.col_num(m.group(1))) + 1) * (abs(int(m.group(4)) - int(m.group(2))) + 1)
    return n


def pick_totals(formulas, top=TRACE_TOP):
    """合計・小計らしい式のセル [(シート, 番地, 式)] → 参照する範囲が大きい順の上位 top 個。"""
    cand = [(s, a, f) for s, a, f in formulas if isinstance(f, str) and _TOTAL_RE.match(f)]
    cand.sort(key=lambda x: -_span(x[2]))
    return cand[:top]


def summarize(sheet_issues):
    """{シート: [issue]} → 全体の issue 一覧（cell を シート!番地 にしたもの）と点数。"""
    import vbam_audit as va
    allv = []
    for sh, lst in sheet_issues.items():
        for i in lst:
            j = dict(i)
            j['cell'] = f"{sh}!{i.get('cell', '')}"
            allv.append(j)
    return allv, va.audit_score(allv)


# 値を見る種類（列ごとにまとめる）。ほかは式を見る種類（同じ形＝R1C1 の式ごとにまとめる）
_VALUE_TYPES = {'hidden_whitespace', 'text_number', 'outlier_value', 'outlier', 'value_in_formula_column'}


def _split_cell(cell):
    """'シート!A12' → (シート, 'A', 12)。"""
    sh, _, a = str(cell).rpartition('!')
    m = re.match(r'([A-Z]+)(\d+)$', a)
    return (sh, m.group(1), int(m.group(2))) if m else (sh, a, 0)


def _box(cells):
    """番地の並び → 囲む範囲 'A1:C9'（1 つなら番地そのもの）。"""
    import vbam_lookup as vl
    parts = [_split_cell(c) for c in cells]
    cols = [vl.col_num(p[1]) for p in parts if p[2]]
    rows = [p[2] for p in parts if p[2]]
    if not rows:
        return ''
    import vbam_audit as va
    a = va._rowcol_to_a1(min(rows) - 1, min(cols) - 1)
    b = va._rowcol_to_a1(max(rows) - 1, max(cols) - 1)
    return a if a == b else f"{a}:{b}"


def _shape_of_msg(msg):
    """内容の文の中の値（数・引用の文字）を伏せた形（同じ種類の文をまとめる鍵）。"""
    s = re.sub(r"'[^']*'", "'…'", str(msg or ''))
    return re.sub(r'-?\d[\d,.]*(?:e[+-]?\d+)?', '#', s)


def group_issues(issues, r1c1=None):
    """病巣の issue を「同じシート・同じ種類・同じ形の式（R1C1）」ごと（値の種類は同じ列ごと）にまとめる（純 Python）。
    2026-11-05 の実測: Enron の公開ブック 60 本の指摘の 7 割が、同じ式のコピーを 1 件ずつ数えたもの（*16 を 3,530 回など）。
    返すのは [{'severity', 'type', 'cells', 'msg', 'formula'}]（重い順・件数の多い順）。点数には使わない（数は変えない）。"""
    r1c1 = r1c1 or {}
    groups = {}
    for i in issues:
        sh, col, _row = _split_cell(i.get('cell', ''))
        if i.get('type') in _VALUE_TYPES:
            key = (i.get('severity'), i.get('type'), sh, 'col:' + col, _shape_of_msg(i.get('msg')))
        else:
            shape = r1c1.get(i.get('cell')) or _shape_of_msg(i.get('formula') or i.get('msg'))
            key = (i.get('severity'), i.get('type'), sh, shape, _shape_of_msg(i.get('msg')))
        g = groups.setdefault(key, {'severity': i.get('severity'), 'type': i.get('type'), 'cells': [],
                                    'msg': i.get('msg', ''), 'formula': i.get('formula', '')})
        g['cells'].append(i.get('cell', ''))
    order = {'critical': 0, 'warning': 1, 'info': 2}
    return sorted(groups.values(), key=lambda g: (order.get(g['severity'], 3), -len(g['cells'])))


def group_lookups(items, r1c1=None):
    """壊れやすい式の [(シート, 番地, 関数, 内容)] を同じ形の式ごとにまとめる → [(シート, [番地], 関数, 内容)]。"""
    r1c1 = r1c1 or {}
    groups = {}
    for sh, a, fn, what in items:
        key = (sh, fn, what, r1c1.get(f"{sh}!{a}", a))
        groups.setdefault(key, [sh, [], fn, what])[1].append(a)
    return sorted((tuple(g) for g in groups.values()), key=lambda g: -len(g[1]))


def fix_hint(g):
    """まとめた 1 行の「直す手」（vbam_audit.fix_of と同じ対応表）。棚の物は選ぶ範囲まで書く。"""
    import vbam_audit as va
    fx = va.fix_of(g)
    if fx['kind'] == 'shelf' and fx['name'] == va.SHELF_FILL_FORMULA:
        sh = g['cells'][0].rpartition('!')[0]
        return f"shelf-run {fx['name']} --select {_box(g['cells'])}" + (f" --sheet {sh}" if sh else '')
    return va.fix_label(fx)


def fix_summary(groups):
    """直す手のまとめの 1 行: 棚・道具で直せる件数と、人の判断の件数。次の手は seiri（dock は書かない）。"""
    import vbam_audit as va
    n = {'shelf': 0, 'tool': 0, 'human': 0}
    for g in groups:
        n[va.fix_of(g)['kind']] += len(g['cells'])
    s = f"直す手: 棚で直せる {n['shelf']} 件・道具で直せる {n['tool']} 件・人の判断 {n['human']} 件。"
    if n['shelf']:
        s += "棚の物は、表を選んで `seiri` を 1 回撃てば同じ対応表で直します（dock は読むだけで書きません）。"
    return s


def _place(cells):
    """まとめた番地の見せ方: 1 つなら番地、2 つ以上は「先頭 ほか N か所（囲む範囲）」。"""
    if len(cells) == 1:
        return cells[0]
    sh = cells[0].rpartition('!')[0]
    box = _box(cells)
    return f"{cells[0]} ほか {len(cells) - 1} か所（{sh + '!' if sh else ''}{box}）"


def grade_of(score):
    return 'A（極めて健全）' if score >= 90 else 'B（概ね良好）' if score >= 75 else 'C（要点検）' if score >= 50 else 'D（危険）'


def render_markdown(data, nv=False):
    """data（構造 d・issues・score・fragile・approx・causes・traces・skipped）→ Markdown 1 枚。"""
    import vbam_structure as vs
    d = data['d']
    L = []
    add = L.append
    add(f"# ブックの点検（dock）: {d['book']}")
    add("（既存の点検 structure・diagnose・lookup・cause・trace を 1 回で撃ってまとめたもの。" + ("値は出していません）" if nv else "）"))
    add("")
    add("## 概要と点数")
    # 近似一致は「壊れやすい式」の節に 1 か所ずつ出すので、病巣の表には重ねない（2026-11-05 Enron の公開ブック 60 本の実測:
    # 同じ VLOOKUP を病巣 895 件・壊れやすい式 1,458 件と 2 回数えていた）。点数は diagnose と同じ数にするため全部で数える
    dup_approx = [i for i in data['issues'] if i.get('type') == 'approx_match'] if data['approx'] else []
    issues = [i for i in data['issues'] if not (dup_approx and i.get('type') == 'approx_match')]
    nc = sum(1 for i in issues if i['severity'] == 'critical')
    nw = sum(1 for i in issues if i['severity'] == 'warning')
    ni = sum(1 for i in issues if i['severity'] == 'info')
    add(f"- 点数: {data['score']} / 100  {grade_of(data['score'])}")
    add(f"- 重大 {nc} 件・警告 {nw} 件・気づき {ni} 件")
    add(f"- シート {len(d['sheets'])} 枚・式 {d['totals'].get('formulas', 0)} 個・形式 {d.get('format', '?')}")
    add(f"- 壊れやすい式: 列の挿入でずれる番号 {len(data['fragile'])} 件・近似一致 {len(data['approx'])} 件")
    for s in data.get('skipped') or []:
        add(f"- 大きすぎて点検を省いたシート: {_md(s)}")
    add("")
    add("## 骨")
    add("| シート | 種類 | 表示 | 保護 | 使用範囲 |")
    add("|---|---|---|---|---|")
    for s in d['sheets']:
        add(f"| {_md(s['name'])} | {s['type']} | {s['vis']} | {'あり' if s['protect'] else '-'} | {s.get('used', '')} |")
    shp = sum(d['shape_counts'].values()) if d.get('shape_counts') else 0
    add(f"- 名前 {len(d['names'])}・テーブル {len(d['tables'])}・ピボット {len(d['pivots'])}・条件付き書式 {len(d['cf'])}・"
        f"入力規則 {len(d['dv'])}・図形 {shp}・外部リンク {len(d['links'])}")
    pts = vs.points(d)
    if pts:
        add("")
        add("気をつける所:")
        for what, why, where in pts:
            add(f"- {_md(what)} — {_md(why)}（確かめる所: {_md(where)}）")
    add("")
    add("## 病巣")
    if not issues:
        add("なし")
    else:
        groups = group_issues(issues, data.get('r1c1'))
        if len(groups) < len(issues):
            add(f"（{len(issues)} 件を、同じ形の式・同じ列ごとに {len(groups)} 行にまとめています）")
        add("| 重さ | 種類 | 場所 | 件数 | 内容 | 直す手 |")
        add("|---|---|---|---:|---|---|")
        for g in groups[:60]:
            msg = '（値は出しません）' if nv else _md(g['msg'])
            add(f"| {_SEV.get(g['severity'], g['severity'])} | {_TYPE_LABEL.get(g['type'], g['type'])} | "
                f"{_md(_place(g['cells']))} | {len(g['cells'])} | {msg} | {_md(fix_hint(g))} |")
        if len(groups) > 60:
            add(f"（ほか {len(groups) - 60} 行・{sum(len(g['cells']) for g in groups[60:])} 件）")
        add("")
        add(fix_summary(groups))
    if dup_approx:
        add(f"（近似一致の {len(dup_approx)} セルは下の「壊れやすい式」に出します。ここには重ねません）")
    add("")
    add("## 壊れやすい式")
    if not data['fragile'] and not data['approx']:
        add("なし")
    else:
        for items, head in ((data['fragile'], "列を挿入・削除すると黙ってずれる（番号の直書き）:"),
                            (data['approx'], "近似一致（探す値が無いとき、エラーにならず別の行を返す）:")):
            if not items:
                continue
            add(head)
            gs = group_lookups(items, data.get('r1c1'))
            for sh, cells, fn, what in gs[:30]:
                where = f"{_md(sh)}!{cells[0]}" if len(cells) == 1 else \
                    f"{_md(sh)}!{cells[0]} ほか {len(cells) - 1} か所（{_md(sh)}!{_box([sh + '!' + c for c in cells])}）"
                add(f"- {where}  {fn}  {what}")
            if len(gs) > 30:
                add(f"- …ほか {len(gs) - 30} 行・{sum(len(g[1]) for g in gs[30:])} 件")
    add("")
    add("## エラーの原因")
    if not data['causes']:
        add("エラーのセルなし")
    else:
        for sh, notes in data['causes']:
            for n in notes:
                add(f"- [{_md(sh)}] {_md(n)}")
    add("")
    add("## 式のつながり")
    if not data['traces']:
        add("合計・小計らしい式なし")
    else:
        for head, lines in data['traces']:
            add(f"### {_md(head)}")
            add("```")
            L.extend(lines)
            add("```")
    return '\n'.join(L) + '\n'


def render_html(data, md):
    """diagnose の HTML 報告（generate_audit_html）に、dock のまとめを足して 1 枚に。"""
    import html
    import vbam_audit as va
    base = va.generate_audit_html(data['d']['book'], '全シート', data['issues'])
    extra = ("<section style=\"max-width:1100px;margin:24px auto;padding:0 16px;font-family:sans-serif\">"
             "<h2>dock のまとめ</h2><pre style=\"white-space:pre-wrap\">" + html.escape(md) + "</pre></section>")
    return base.replace('</body>', extra + '</body>', 1) if '</body>' in base else base + extra



# ----------------------------------------------------------------
# 新しい指摘（2026-11-06）: 重複した条件付き書式・孤立した式
# ----------------------------------------------------------------

def _first_cell(addr):
    m = re.match(r"\$?([A-Z]{1,3})\$?(\d+)", str(addr))
    return f"{m.group(1)}{m.group(2)}" if m else str(addr)


def cf_duplicate_issues(cf_rows):
    """条件付き書式の行（vbam_structure._cf_rows の形）→ 重複・完全に隠れた規則の issue（純 Python）。
    重複＝同じシート・同じ範囲・同じ条件・同じ書式の規則が 2 つ以上。
    隠れ＝同じ範囲・同じ条件の規則が先にあり、それが「真のとき停止」で、あとの規則に順番が回らない。"""
    seen, out = {}, []
    for r in cf_rows or []:
        sig = r.get('sig')
        if sig is None:
            continue
        cond = (r['sheet'], r['addr'], r['type'], r['op'], r['f1'], r['f2'])
        prior = seen.setdefault(cond, [])
        label = f"{r['type']} {r['op']} {r['f1']} {r['f2']}".strip()
        if any(p == sig for p, _ in prior):
            out.append({'sheet': r['sheet'], 'cell': _first_cell(r['addr']), 'severity': 'info', 'type': 'duplicate_cf',
                        'msg': f"条件付き書式の規則が、同じ範囲 {r['addr']} に同じ条件・同じ書式で重なっています（{label[:40]}）＝ルールの管理で 1 つを消す"})
        elif any(st for _, st in prior):
            out.append({'sheet': r['sheet'], 'cell': _first_cell(r['addr']), 'severity': 'info', 'type': 'duplicate_cf',
                        'msg': f"条件付き書式の規則が、同じ範囲 {r['addr']} の同じ条件の先の規則（真のとき停止）に隠れて働きません（{label[:40]}）"})
        prior.append((sig, bool(r.get('stop'))))
    return out


_REF_RE = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|([^\s'!=+\-*/^&<>(),;:\[\]{}\"]+))!)?"
    r"(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![\w(!.])")


def _col_n(s):
    import vbam_lookup as vl
    return vl.col_num(s)


def refs_of(formula, here):
    """式の中のセル参照 → [(シート, r1, c1, r2, c2)]。3-D 参照があれば None（読み切れない）。"""
    f = re.sub(r'"(?:[^"]|"")*"', '""', formula)
    if re.search(r"[A-Za-z0-9_']:[A-Za-z0-9_']+!", f):
        return None
    out = []
    for m in _REF_RE.finditer(f):
        sh = (m.group(1) or m.group(2) or here)
        if m.group(1):
            sh = sh.replace("''", "'")
        if sh.startswith('[') or ']' in sh:
            continue
        a = m.group(3).replace('$', '')
        parts = a.split(':')
        try:
            if re.fullmatch(r'\d+', parts[0]):
                r1, r2, c1, c2 = int(parts[0]), int(parts[-1]), 1, 16384
            elif re.fullmatch(r'[A-Z]+', parts[0]):
                r1, r2, c1, c2 = 1, 1048576, _col_n(parts[0]), _col_n(parts[-1])
            else:
                x = [re.match(r'([A-Z]+)(\d+)', p) for p in parts]
                r1, c1 = int(x[0].group(2)), _col_n(x[0].group(1))
                r2, c2 = int(x[-1].group(2)), _col_n(x[-1].group(1))
        except Exception:
            continue
        out.append((sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
    return out


_DYNAMIC_RE = re.compile(r'\b(?:INDIRECT|OFFSET)\s*\(|\[', re.I)


ORPHAN_EMIT = False


def find_orphan_formulas(formulas, grids, extra_text=(), tables=(), skip=False):
    """孤立した式（純 Python）。
    formulas: [(シート, 番地, 式)]・grids: {シート: (r0, c0, 式の格子, 値の格子)}・extra_text: [(シート, 式の文字)]
    （名前・条件付き書式・入力規則・グラフの参照先）。
    孤立＝ほかのどの式（同じシート・他シート）・名前・書式・グラフからも参照されず、合計の範囲にも入っていない、
    参照を持つ数値の式。最終の合計・出力欄（下にも右にも数・式が無いセル）は言わない。
    動的な参照（INDIRECT・OFFSET）・テーブルの構造化参照・3-D 参照のあるブックは読み切れないので言わない。"""
    if skip or not formulas:
        return []
    for _sh, _a, f in formulas:
        if _DYNAMIC_RE.search(re.sub(r'"(?:[^"]|"")*"', '""', f)):
            return []
    by_col = {}      # (シート, 列) -> [(r1, r2)]

    def add(sh, r1, c1, r2, c2):
        for c in range(c1, min(c2, c1 + 600) + 1):
            by_col.setdefault((sh, c), []).append((r1, r2))
    for sh, a, f in formulas:
        rs = refs_of(f, sh)
        if rs is None:
            return []
        for x in rs:
            add(*x)
    for sh, txt in extra_text:
        for x in refs_of(txt, sh) or []:
            add(*x)
    for t in tables:
        for x in refs_of(f"{t['sheet']}!{t['addr']}", t['sheet']) or []:
            add(*x)

    def used(sh, r, c):
        return any(r1 <= r <= r2 for r1, r2 in by_col.get((sh, c), ()))

    def num(x):
        return isinstance(x, (int, float)) and not isinstance(x, bool)

    def live(x):
        return num(x) or (isinstance(x, str) and x.startswith('='))

    pat_n = {}
    for sh_, g_ in grids.items():
        if len(g_) > 4 and g_[4]:
            for row_ in g_[4]:
                for x_ in row_:
                    if isinstance(x_, str) and x_.startswith('='):
                        pat_n[(sh_, x_)] = pat_n.get((sh_, x_), 0) + 1
    out = []
    for sh, a, f in formulas:
        g = grids.get(sh)
        m = re.match(r'([A-Z]+)(\d+)$', a)
        if not g or not m:
            continue
        r0, c0, fa, fv = g[:4]
        c, r = _col_n(m.group(1)), int(m.group(2))
        i, j = r - r0, c - c0
        if len(g) > 4 and g[4]:                 # 同じ形（R1C1）の式がシートに 2 つ以上ある＝列・行に並んだ計算は
            if pat_n.get((sh, g[4][i][j]), 0) > 1:     # 出力欄（比率・差・小計）が多く、使われないのが普通。孤立は「1 つだけ浮いた式」に絞る
                continue
        if i < 0 or j < 0 or i >= len(fv) or j >= len(fv[i]) or not num(fv[i][j]):
            continue
        if not refs_of(f, sh) or used(sh, r, c):
            continue
        below = any(j < len(fa[k]) and (live(fa[k][j]) or num(fv[k][j])) for k in range(i + 1, len(fa)))
        right = any(live(fa[i][k]) or num(fv[i][k]) for k in range(j + 1, len(fa[i])))
        if not (below and right):
            continue
        out.append({'sheet': sh, 'cell': a, 'severity': 'info', 'type': 'orphan_formula', 'formula': f,
                    'msg': "この式の結果は、ほかのどの式・名前・グラフにも使われていません（合計にも入っていません）"})
    return out


# ----------------------------------------------------------------
# 「39 の原則」から足した型（2026-11-06）。式の文字・格子・結合/非表示の事実だけで決まるもの（純 Python）
# ----------------------------------------------------------------

_STR_RE = re.compile(r'"(?:[^"]|"")*"')
# 3-D 参照: Sheet1:Sheet3!A1 / 'Jan:Dec'!B2（シート名の範囲）。A1:B2 の範囲や Sheet1!A1:B2 は含まない
_3D_RE = re.compile(r"(?<![\w!.'])(?:'[^'!\[\]]+:[^'!\[\]]+'|[A-Za-z_][\w.]*:[A-Za-z_][\w.]*)!(?=\$?[A-Z]{1,3}\$?\d|\$?[A-Z]{1,3}:|\$?\d+:)")
# 外部ブック参照: [Book.xlsx]Sheet!A1 / 'C:\dir\[Book.xls]Sheet'!A1 / [1]Sheet!A1（構造化参照 Tbl[列] は語の直後なので当たらない）
_EXT_RE = re.compile(r"\[[^\]\[]+\.(?:xls[xmb]?|xlam|xla)\]|(?<![\w\]\)])\[\d+\](?=[A-Za-z_'])", re.IGNORECASE)
LONG_FORMULA_CHARS = 300
LONG_FORMULA_DEPTH = 8


def _formula_depth(f):
    d = m = 0
    for ch in _STR_RE.sub('""', f):
        if ch == '(':
            d += 1
            m = max(m, d)
        elif ch == ')':
            d -= 1
    return m


def formula_principle_issues(formulas):
    """formulas: [(シート, 番地, 式)] → 原則 13（読める式）・26（単純な範囲＝3-D 参照を使わない）・31（外部参照は控えめに）。"""
    out = []
    for sh, a, f in formulas:
        if not isinstance(f, str) or not f.startswith('='):
            continue
        t = _STR_RE.sub('""', f)
        if _3D_RE.search(t):
            out.append({'sheet': sh, 'cell': a, 'severity': 'info', 'type': 'three_d_reference', 'formula': f,
                        'msg': "複数シートをまたぐ範囲（3-D 参照）です＝間にシートを挿入・移動すると集計対象が黙って変わります"})
        if _EXT_RE.search(t):
            out.append({'sheet': sh, 'cell': a, 'severity': 'info', 'type': 'external_link', 'formula': f,
                        'msg': "ほかのブックを参照する式です＝リンク先が移動・改名・未更新だと古い値のまま、または壊れます"})
        n, dp = len(f), _formula_depth(f)
        if n >= LONG_FORMULA_CHARS or dp >= LONG_FORMULA_DEPTH:
            why = f"{n} 文字" if n >= LONG_FORMULA_CHARS else f"かっこが {dp} 重"
            out.append({'sheet': sh, 'cell': a, 'severity': 'info', 'type': 'long_formula', 'formula': f,
                        'msg': f"式が長く複雑です（{why}）＝途中の計算を別のセルに分けると読めて検算できます"})
    return out


def structure_principle_issues(sheet, vis, merged, hidden_cols, hidden_rows, grid_vals, r0, c0):
    """原則 16（結合しない）・17（隠すより保護）。純 Python（COM の事実を引数で受ける）。
    vis: -1 表示・0 非表示・2 VeryHidden。merged: [(r1, c1, r2, c2)]（シートの絶対行列・1 始まり）。
    hidden_cols / hidden_rows: 値か式の入った非表示の列番号・行番号（フィルタ・グループ化の折りたたみは除いて渡す）。"""
    import vbam_audit as va
    out = []
    if vis in (0, 2):
        out.append({'sheet': sheet, 'cell': 'A1', 'severity': 'info', 'type': 'hidden_content',
                    'msg': ("シートが VeryHidden（通常の操作では再表示できない）です" if vis == 2 else "シートが非表示です")
                           + "＝隠すより、数式のセルをロックしてシートを保護する"})
    for c in hidden_cols:
        out.append({'sheet': sheet, 'cell': va._rowcol_to_a1(max(r0 - 1, 0), c - 1), 'severity': 'info',
                    'type': 'hidden_content', 'msg': "値・式の入った列が非表示です＝隠すより、数式のセルをロックしてシートを保護する"})
    for r in hidden_rows:
        out.append({'sheet': sheet, 'cell': va._rowcol_to_a1(r - 1, max(c0 - 1, 0)), 'severity': 'info',
                    'type': 'hidden_content', 'msg': "値・式の入った行が非表示です＝隠すより、数式のセルをロックしてシートを保護する"})

    def filled(r, c):
        i, j = r - r0, c - c0
        return 0 <= i < len(grid_vals) and 0 <= j < len(grid_vals[i]) and grid_vals[i][j] not in (None, '')
    for r1, c1, r2, c2 in merged:
        if r2 <= r1:                       # 縦に結合していない（見出し・タイトルの横結合は言わない）
            continue
        inside = {(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
        col_n = sum(1 for r in range(r0, r0 + len(grid_vals)) if (r, c1) not in inside and filled(r, c1))
        if col_n < 5:                      # 同じ列に表の明細が無い＝帳票の入力欄・ラベル。並べ替え・フィルタの対象ではない
            continue
        out.append({'sheet': sheet, 'cell': va._rowcol_to_a1(r1 - 1, c1 - 1), 'severity': 'info', 'type': 'merged_cells',
                    'msg': "表の列の途中で縦にセルを結合しています＝並べ替え・絞り込み・コピーが効かなくなります（「選択範囲の中央」で代用できます）"})
    return out


def _structure_facts(ws, r0, c0, nr, nc, fa, fv):
    """COM で結合・非表示の事実を集める → (vis, merged, hidden_cols, hidden_rows)。読めなければ空。"""
    vis, merged, hc, hr = -1, [], [], []
    try:
        vis = int(ws.Visible)
    except Exception:
        pass

    def live(x):
        return x not in (None, '')
    try:
        for j in range(nc):
            col = ws.Columns(c0 + j)
            if col.Hidden and int(col.OutlineLevel) <= 1 and any(live(row[j]) for row in fv if j < len(row)):
                hc.append(c0 + j)
    except Exception:
        pass
    try:
        if not ws.AutoFilterMode and nr <= 3000:
            for i in range(nr):
                if any(live(x) for x in fv[i]):
                    rw = ws.Rows(r0 + i)
                    if rw.Hidden and int(rw.OutlineLevel) <= 1:
                        hr.append(r0 + i)
    except Exception:
        pass
    try:
        import vbam_view as vv
        areas, skipped = vv._merged_areas_in_range(ws.UsedRange)
        for ad in areas or []:
            m = re.match(r'\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$', ad)
            if m and m.group(3):
                merged.append((int(m.group(2)), _col_n(m.group(1)), int(m.group(4)), _col_n(m.group(3))))
    except Exception:
        pass
    return vis, merged, hc, hr


# ----------------------------------------------------------------
# COM
# ----------------------------------------------------------------

def _grids(ws):
    """(r0, c0, 式の格子, R1C1 の格子, 値の格子) か None（大きすぎ）。"""
    import vbam_audit as va
    rng = ws.UsedRange
    n_rows, n_cols = int(rng.Rows.Count), int(rng.Columns.Count)
    if n_rows * n_cols > MAX_CELLS_PER_SHEET:
        return None
    return (int(rng.Row), int(rng.Column), va.to_2d_list(rng.Formula, n_rows, n_cols),
            va.to_2d_list(rng.FormulaR1C1, n_rows, n_cols), va.to_2d_list(rng.Value2, n_rows, n_cols))


def _has_error(fv):
    import vbam_cause as vc
    return any(vc._err(x) for row in fv for x in row)


def _err_text(fv):
    """Value2 はエラーを整数で返す。linter は '#DIV/0!' の文字列で見るので文字にして渡す（格子は写し）。"""
    import vbam_cause as vc
    return [[vc._err(x) or x for x in row] for row in fv]



def _has_macros(wb):
    try:
        return bool(wb.HasVBProject)
    except Exception:
        return False


def _chart_refs(wb):
    """グラフの系列の参照先 [(シート, '=式')]（ワークシート上のグラフ・グラフシート）。"""
    out = []
    try:
        for ws in wb.Worksheets:
            for co in ws.ChartObjects():
                for se in co.Chart.SeriesCollection():
                    out.append((ws.Name, str(se.Formula)))
    except Exception:
        pass
    try:
        for ch in wb.Charts:
            for se in ch.SeriesCollection():
                out.append(('', str(se.Formula)))
    except Exception:
        pass
    return out


def collect_all(wb, only_sheet=None):
    """ブックを読んで render_markdown に渡す data を作る。"""
    import vbam_structure as vs
    import vbam_audit as va
    import vbam_cause as vc
    d = vs.collect(wb, only_sheet)
    nv = vs.no_values_mode()
    sheet_issues, formulas, causes, skipped, r1c1 = {}, [], [], [], {}
    names, udfs = vc.workbook_names_udfs(wb)
    # 同じ番地・同じ値の比較用に、ほかのシートの値を先に集める（1 シートだけの点検では空）
    peer_all, grid_cache = {}, {}
    if not only_sheet:
        for ws_ in wb.Worksheets:
            try:
                g_ = _grids(ws_)
            except Exception:
                continue
            grid_cache[ws_.Name] = g_
            if g_ is not None:
                peer_all[ws_.Name] = (g_[0] - 1, g_[1] - 1, g_[4])
    for ws in wb.Worksheets:
        if only_sheet and ws.Name != only_sheet:
            continue
        try:
            g = grid_cache[ws.Name] if ws.Name in grid_cache else _grids(ws)
        except Exception as e:
            skipped.append(f"{ws.Name}（読めません: {e}）")
            continue
        if g is None:
            skipped.append(f"{ws.Name}（{MAX_CELLS_PER_SHEET:,} セル超）")
            continue
        r0, c0, fa, f1, fv = g
        fvs = _err_text(fv)
        peers = [v_ for n_, v_ in peer_all.items() if n_ != ws.Name]
        iss = (va.check_formula_linter(fa, f1, fvs, r0 - 1, c0 - 1)
               + va.check_data_cleaner(fvs, r0 - 1, c0 - 1, fa, peers))
        fml = set()
        for i, row in enumerate(fa):
            for j, x in enumerate(row):
                if isinstance(x, str) and x.startswith('='):
                    a = va._rowcol_to_a1(r0 - 1 + i, c0 - 1 + j)
                    fml.add(a)
                    formulas.append((ws.Name, a, x))
                    try:
                        r1c1[f"{ws.Name}!{a}"] = str(f1[i][j])   # 同じ形の式をまとめて見せる鍵
                    except Exception:
                        pass
        sheet_issues[ws.Name] = [i for i in iss if not (i.get('type') == 'outlier_value' and i.get('cell') in fml)]
        try:
            _vis, _mg, _hc, _hr = _structure_facts(ws, r0, c0, len(fa), len(fa[0]) if fa else 0, fa, fvs)
            sheet_issues[ws.Name] += structure_principle_issues(ws.Name, _vis, _mg, _hc, _hr, fvs, r0, c0)
        except Exception:
            pass
        if _has_error(fv):
            try:
                notes = vc.cause_notes(fa, fv, r0, c0, names, udfs, nv)
                if notes:
                    causes.append((ws.Name, notes))
            except Exception:
                pass
    try:
        for it in formula_principle_issues(formulas):
            sheet_issues.setdefault(it.pop('sheet'), []).append(it)
    except Exception:
        pass
    try:
        for it in cf_duplicate_issues(d.get('cf')):
            sh_ = it.pop('sheet')
            if not only_sheet or sh_ == only_sheet:
                sheet_issues.setdefault(sh_, []).append(it)
        if not only_sheet:
            extra = [(r['sheet'], f"={r['f1']}") for r in d.get('cf') or [] if r.get('f1')]
            extra += [(r['sheet'], f"={r['f2']}") for r in d.get('cf') or [] if r.get('f2')]
            extra += [(r['sheet'], str(r.get('formula') or '')) for r in d.get('dv') or []]
            extra += [('', v) for v in names.values()]
            extra += _chart_refs(wb)
            gr = {n_: (g_[0], g_[1], g_[2], g_[4], g_[3]) for n_, g_ in grid_cache.items() if g_ is not None}
            skip = len(gr) < len(grid_cache) or len(formulas) > 150000 or _has_macros(wb)
            # 孤立した式は、実物 27 件（290 冊）が全部出力欄（比率・合計・見出しの写し）で本物 0。拾える形（単体の補助セル）の実例が無く、出すと誤爆になるため既定では出さない（2026-11-06）
            for it in (find_orphan_formulas(formulas, gr, extra, d.get('tables') or [], skip) if ORPHAN_EMIT else []):
                sheet_issues.setdefault(it.pop('sheet'), []).append(it)
    except Exception:
        pass
    issues, score = summarize(sheet_issues)
    fragile, approx = lookup_findings(formulas)
    traces = []
    try:
        import vbam_view as vv
        tnames, ttables = vv._trace_book_words(wb)
        for sh, a, f in pick_totals(formulas):
            st = {'lines': [], 'nodes': 0, 'names': tnames, 'tables': ttables}
            vv._trace_walk(wb, sh, a, 2, frozenset(), st, indent=0)
            lines = st['lines']
            if nv:
                lines = [re.sub(r'\s+→ .*$', '', x) for x in lines if not x.lstrip().startswith('内訳')]
            traces.append((f"{sh}!{a}  {f}", lines))
    except Exception:
        pass
    return {'d': d, 'issues': issues, 'score': score, 'fragile': fragile, 'approx': approx,
            'causes': causes, 'traces': traces, 'skipped': skipped, 'r1c1': r1c1}


def cmd_dock(args):
    """dock [excel_file] [--sheet 名] [--out 出力.md] [--html]: 骨・病巣・壊れやすい式・エラーの原因・式のつながりを 1 枚に。"""
    import vbam_core as vc
    import vbam_structure as vs
    target_file, rest = vc.parse_target_and_rest(args.posargs)
    xl, wb = vc.get_workbook(target_file, readonly=True)
    only = getattr(args, 'sheet_opt', None) or (rest[0] if rest else None)
    try:
        vc.cmd_progress_note(f"{wb.Name}: 点検の部品をまとめて撃っています")
    except Exception:
        pass
    nv = vs.no_values_mode()
    data = collect_all(wb, only)
    md = render_markdown(data, nv)
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.abspath(getattr(args, 'out', None) or os.path.join(here, '_last_dock.md'))
    with open(out, 'w', encoding='utf-8') as f:
        f.write(md)
    print(md)
    msg = f"（同じ内容を {out} に保存しました" + ("。値は入っていません）" if nv else "）")
    if getattr(args, 'html', False):
        hp = os.path.splitext(out)[0] + '.html'
        with open(hp, 'w', encoding='utf-8') as f:
            f.write(render_html(data, md))
        msg += f"\n🌐 HTML: {hp}"
    print(msg)
    return True
