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
    issues = data['issues']
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
        add("| 重さ | 種類 | 場所 | 内容 |")
        add("|---|---|---|---|")
        order = {'critical': 0, 'warning': 1, 'info': 2}
        for i in sorted(issues, key=lambda x: order.get(x['severity'], 3))[:60]:
            msg = '（値は出しません）' if nv else _md(i.get('msg', ''))
            add(f"| {_SEV.get(i['severity'], i['severity'])} | {_TYPE_LABEL.get(i['type'], i['type'])} | {_md(i['cell'])} | {msg} |")
        if len(issues) > 60:
            add(f"（ほか {len(issues) - 60} 件）")
    add("")
    add("## 壊れやすい式")
    if not data['fragile'] and not data['approx']:
        add("なし")
    else:
        if data['fragile']:
            add("列を挿入・削除すると黙ってずれる（番号の直書き）:")
            for sh, a, fn, what in data['fragile'][:30]:
                add(f"- {_md(sh)}!{a}  {fn}  {what}")
            if len(data['fragile']) > 30:
                add(f"- …ほか {len(data['fragile']) - 30} 件")
        if data['approx']:
            add("近似一致（探す値が無いとき、エラーにならず別の行を返す）:")
            for sh, a, fn, what in data['approx'][:30]:
                add(f"- {_md(sh)}!{a}  {fn}  {what}")
            if len(data['approx']) > 30:
                add(f"- …ほか {len(data['approx']) - 30} 件")
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


def collect_all(wb, only_sheet=None):
    """ブックを読んで render_markdown に渡す data を作る。"""
    import vbam_structure as vs
    import vbam_audit as va
    import vbam_cause as vc
    d = vs.collect(wb, only_sheet)
    nv = vs.no_values_mode()
    sheet_issues, formulas, causes, skipped = {}, [], [], []
    names, udfs = vc.workbook_names_udfs(wb)
    for ws in wb.Worksheets:
        if only_sheet and ws.Name != only_sheet:
            continue
        try:
            g = _grids(ws)
        except Exception as e:
            skipped.append(f"{ws.Name}（読めません: {e}）")
            continue
        if g is None:
            skipped.append(f"{ws.Name}（{MAX_CELLS_PER_SHEET:,} セル超）")
            continue
        r0, c0, fa, f1, fv = g
        fvs = _err_text(fv)
        iss = va.check_formula_linter(fa, f1, fvs, r0 - 1, c0 - 1) + va.check_data_cleaner(fvs, r0 - 1, c0 - 1)
        fml = set()
        for i, row in enumerate(fa):
            for j, x in enumerate(row):
                if isinstance(x, str) and x.startswith('='):
                    a = va._rowcol_to_a1(r0 - 1 + i, c0 - 1 + j)
                    fml.add(a)
                    formulas.append((ws.Name, a, x))
        sheet_issues[ws.Name] = [i for i in iss if not (i.get('type') == 'outlier_value' and i.get('cell') in fml)]
        if _has_error(fv):
            try:
                notes = vc.cause_notes(fa, fv, r0, c0, names, udfs, nv)
                if notes:
                    causes.append((ws.Name, notes))
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
            'causes': causes, 'traces': traces, 'skipped': skipped}


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
