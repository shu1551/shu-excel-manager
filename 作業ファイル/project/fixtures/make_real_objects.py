# -*- coding: utf-8 -*-
"""グラフ・ピボット・テーブル・パワークエリの実物らしい表（2026-09-18 未明）: 鍛えた 6 本に、最初のお題に無かった癖を当てる。

  課別棒グラフ: 表がテーブル（ListObject）／課名の前後に空白（系列の項目は元のセルのまま）／課が 1 つだけ
  推移折れ線: 課が 1 つだけ／種違い
  構成比円: 区分の表がテーブル（合計行なし）／0 でない区分が 1 つだけ
  課別科目別ピボット: 明細がテーブル（T支出明細）／1 万行
  職員名簿テーブル: 既存のテーブルの範囲が狭い（最後の 3 人が外れている＝範囲を広げて直す）
  PQ課別集計: T支出 で始まるテーブルが 1 つだけ／3 か月
  見やすい棒グラフ: テーブル／課が 1 つ／表題と単位の行／単位の行だけ／16 課／表題と合計行の 8 課

各型の make_pairs.py の write_before / 正解の作り方をそのまま使い、置き場は <型>\\未見_実物（種 1）／未見_実物種N。
  py make_real_objects.py [種]
"""
import importlib.util
import os
import random
import sys

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', 'python_scripts')))
SEED = 1


def _mod(kind):
    spec = importlib.util.spec_from_file_location(f"mp_{kind}", os.path.join(HERE, kind, 'make_pairs.py'))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def _out(kind):
    p = os.path.join(HERE, kind, '未見_実物' if SEED == 1 else f'未見_実物種{SEED}')
    os.makedirs(p, exist_ok=True)
    for f in os.listdir(p):
        if f.endswith('.xlsx'):
            os.remove(os.path.join(p, f))
    return p


def as_table(path, sheet, ref, name):
    wb = load_workbook(path)
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
    wb[sheet].add_table(t)
    wb.save(path)


def bar(xl, seed):
    G = _mod('グラフ')
    out = _out('グラフ')
    rng = random.Random(seed)
    jobs = []
    rows = G.build(seed * 100 + 1, rng.randint(4, 9))
    jobs.append(('実物01_テーブル', rows, {}, lambda p, lay, n=len(rows): as_table(p, G.SHEET, f"A1:B{1 + n}", 'T課別支出')))
    rows = [(f" {k}" if i % 2 else f"{k}　", a) for i, (k, a) in enumerate(G.build(seed * 100 + 2, 6))]
    jobs.append(('実物02_課名に空白', rows, {}, None))
    jobs.append(('実物03_課が1つ', G.build(seed * 100 + 3, 1), {}, None))
    for stem, rows, kw, post in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = G.write_before(before, rows, **kw)
        if post:
            post(before, lay)
        G.build_truth(xl, before, truth, lay)


def trend(xl, seed):
    T = _mod('グラフ推移')
    out = _out('グラフ推移')
    for stem, rows in (('実物01_課が1つ', T.build(seed * 100 + 1, 1)), ('実物02_標準の種違い', T.build(seed * 100 + 2, 3))):
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = T.write_before(before, rows)
        T.build_truth(xl, before, truth, lay)


def pie(xl, seed):
    C = _mod('グラフ構成比')
    out = _out('グラフ構成比')
    rows = C.build(seed * 100 + 1, 6)
    before, truth = os.path.join(out, "実物01_テーブル_前.xlsx"), os.path.join(out, "実物01_テーブル_正解.xlsx")
    lay = C.write_before(before, rows, total=False)
    as_table(before, C.SHEET, f"A1:B{1 + len(rows)}", 'T性質別')
    C.build_truth(xl, before, truth, lay)
    rows = [(k, 0) for k, _a in C.build(seed * 100 + 2, 5)]
    rows[2] = (rows[2][0], 1234000)
    before, truth = os.path.join(out, "実物02_0でない区分が1つ_前.xlsx"), os.path.join(out, "実物02_0でない区分が1つ_正解.xlsx")
    lay = C.write_before(before, rows)
    C.build_truth(xl, before, truth, lay)


def pivot(xl, seed):
    V = _mod('ピボット')
    out = _out('ピボット')
    jobs = [('実物01_明細がテーブル', V.build(seed * 100 + 1, 80), 'table'), ('実物02_1万行', V.build(seed * 100 + 2, 10000, n_ka=10), None)]
    for stem, rows, kind in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = V.write_before(before, rows)
        if kind == 'table':
            as_table(before, V.SHEET, f"A1:E{lay['last']}", 'T支出明細')
        wb = xl.Workbooks.Open(before)
        try:
            V.build_pivot(wb, lay)
            wb.Worksheets(V.SHEET).Activate()
            wb.SaveAs(truth, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)


def table(xl, seed):
    B = _mod('テーブル')
    out = _out('テーブル')
    rows = B.build(seed * 100 + 1, 15)
    stem = '実物01_既存の範囲が狭い'
    before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
    lay = B.write_before(before, rows)
    as_table(before, B.SHEET, f"A1:E{lay['last'] - 3}", 'テーブル1')
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Sheets(B.SHEET)
        lo = ws.ListObjects(1)
        lo.Resize(ws.Range(ws.Cells(lay['h'], 1), ws.Cells(lay['last'], lay['ncols'])))   # 正解は範囲を広げて直す
        B.build_table(ws, lay)
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def mieru(xl, seed):
    """見やすい棒グラフ（2026-09-18）: テーブル／課が 1 つ／表題と単位の行／単位の行だけ（表題ではない）／16 課／表題・合計行つき 8 課。"""
    M = _mod('グラフ見やすい')
    out = _out('グラフ見やすい')
    from openpyxl import load_workbook as _lw

    def unit_at(cell, drop_title=False):
        def post(p, lay):
            wb = _lw(p)
            ws = wb[M.SHEET]
            if drop_title:
                ws['A1'] = None
            ws[cell] = '（単位：円）'
            wb.save(p)
            if drop_title:
                lay['title'] = None
        return post

    rng = random.Random(seed)
    rows = M.build(seed * 100 + 1, rng.randint(5, 12))
    jobs = [('実物01_テーブル', rows, {}, lambda p, lay, n=len(rows): as_table(p, M.SHEET, f"A1:B{1 + n}", 'T課別支出')),
            ('実物02_課が1つ', M.build(seed * 100 + 2, 1), {}, None),
            ('実物03_表題と単位の行', M.build(seed * 100 + 3, 10), {'title': M.TITLES[2]}, unit_at('B2')),
            ('実物04_単位の行だけ', M.build(seed * 100 + 4, 7), {'title': 'x'}, unit_at('A1', drop_title=True)),
            ('実物05_16課', M.build(seed * 100 + 5, 16), {}, None),
            ('実物06_表題と合計行の8課', M.build(seed * 100 + 6, 8), {'title': M.TITLES[0], 'total': True}, None)]
    for stem, rows, kw, post in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = M.write_before(before, rows, **kw)
        if post:
            post(before, lay)
        M.build_truth(xl, before, truth, lay)


def mpivot(xl, seed):
    """課別月別ピボット（2026-09-18）: 明細がテーブル／1 万行／年度をまたぐ（4 月〜翌 3 月）。"""
    import datetime
    V = _mod('ピボット月別')
    out = _out('ピボット月別')
    fiscal = V.build(seed * 100 + 3, 90)
    rng = random.Random(seed)
    for r in fiscal:
        m = rng.randint(4, 15)
        r['支出日'] = datetime.datetime(2026 + (m > 12), (m - 1) % 12 + 1, r['支出日'].day)
    jobs = [('実物01_明細がテーブル', V.build(seed * 100 + 1, 80), 'table'), ('実物02_1万行', V.build(seed * 100 + 2, 10000, n_ka=10), None),
            ('実物03_年度をまたぐ', fiscal, None)]
    for stem, rows, kind in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = V.write_before(before, rows)
        if kind == 'table':
            as_table(before, V.SHEET, f"A1:E{lay['last']}", 'T支出明細')
        wb = xl.Workbooks.Open(before)
        try:
            V.build_pivot(wb, lay)
            wb.Worksheets(V.SHEET).Activate()
            wb.SaveAs(truth, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)


def pqlong(xl, seed):
    """PQ縦持ち（2026-09-18）: 下半期だけ（10 月〜3 月）／テーブルの上に表題（2 行下げる）／月の見出しが「4月分」／課が 1 つ。"""
    Q = _mod('パワークエリ縦持ち')
    out = _out('パワークエリ縦持ち')
    rows, months = Q.build(seed * 100 + 1, 6, 12)
    half = months[6:]
    jobs = [('実物01_下半期だけ', ([{k: v for k, v in r.items() if k in ('課コード', '課名') or k in half} for r in rows], half), None),
            ('実物02_表の上に表題', Q.build(seed * 100 + 2, 5, 6), 'title'),
            ('実物03_月の見出しが分つき', Q.build(seed * 100 + 3, 4, 6), 'bun'),
            ('実物04_課が1つ', Q.build(seed * 100 + 4, 1, 12), None)]
    for stem, data, kind in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        Q.write_before(before, data)
        if kind:
            wb = xl.Workbooks.Open(before)
            try:
                ws = wb.Worksheets(Q.SHEET)
                lo = ws.ListObjects(1)
                if kind == 'title':
                    ws.Rows("1:2").Insert()
                    ws.Range("A1").Value = '令和8年度 月別支出（単位：円）'
                else:
                    for j in range(1, int(lo.ListColumns.Count) + 1):
                        nm = str(lo.ListColumns(j).Name)
                        if nm.endswith('月'):
                            lo.HeaderRowRange.Cells(1, j).Value = nm + '分'
                wb.Save()
            finally:
                wb.Close(SaveChanges=False)
        wb = xl.Workbooks.Open(before)
        try:
            Q.load_query(wb)
            wb.Worksheets(1).Activate()
            wb.SaveAs(truth, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)


def dash(xl, seed):
    """課別ダッシュボード（2026-09-18）: 明細がテーブル／1 万行／前のダッシュボードが 2 つのスライサーつきで残っている。"""
    V = _mod('ピボットグラフ')
    out = _out('ピボットグラフ')
    jobs = [('実物01_明細がテーブル', V.P.build(seed * 100 + 1, 80), 'table'), ('実物02_1万行', V.P.build(seed * 100 + 2, 10000, n_ka=10), None),
            ('実物03_前のスライサー2つ', V.P.build(seed * 100 + 3, 60), 'stale2')]
    for stem, rows, kind in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = V.P.write_before(before, rows)
        if kind == 'table':
            as_table(before, V.SHEET, f"A1:E{lay['last']}", 'T支出明細')
        if kind == 'stale2':
            wb = xl.Workbooks.Open(before)
            try:
                pt = V.build_dashboard(wb, lay, slicer_field='課名', title='旧')
                ws = wb.Worksheets(V.OUT_SHEET)
                wb.SlicerCaches.Add2(pt, '摘要').Slicers.Add(SlicerDestination=ws, Caption='摘要', Top=300, Left=600, Width=144, Height=200)
                wb.Worksheets(V.SHEET).Activate()
                wb.Save()
            finally:
                wb.Close(SaveChanges=False)
        wb = xl.Workbooks.Open(before)
        try:
            V.build_dashboard(wb, lay)
            wb.Worksheets(V.SHEET).Activate()
            wb.SaveAs(truth, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)


def combo(xl, seed):
    """予算執行の複合グラフ（2026-09-18）: 表がテーブル／課が 1 つ／見出しに単位「予算額（千円）」／15 課で執行率が値。"""
    C = _mod('グラフ複合')
    out = _out('グラフ複合')
    from openpyxl import load_workbook as _lw

    def units(p, lay):
        wb = _lw(p)
        ws = wb[C.SHEET]
        for j, c in enumerate(lay['cols'], 1):
            if c in ('予算額', '執行額'):
                ws.cell(lay['h'], j).value = c + '（千円）'
        wb.save(p)
        lay['cols'] = [c + '（千円）' if c in ('予算額', '執行額') else c for c in lay['cols']]

    rows = C.build(seed * 100 + 1, 6)
    jobs = [('実物01_テーブル', rows, {}, lambda p, lay, n=len(rows): as_table(p, C.SHEET, f"A1:D{1 + n}", 'T予算執行')),
            ('実物02_課が1つ', C.build(seed * 100 + 2, 1), {}, None),
            ('実物03_見出しに単位', C.build(seed * 100 + 3, 7), {}, units),
            ('実物04_15課で執行率が値', C.build(seed * 100 + 4, 15), {'formula': False}, None)]
    for stem, rows, kw, post in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = C.write_before(before, rows, **kw)
        if post:
            post(before, lay)
        if stem == '実物03_見出しに単位':
            cols = lay['cols']
            lay = dict(lay, cols=[c.replace('（千円）', '') for c in cols])      # 正解の作り方は列の位置で探す（名前は元のまま）
        C.build_truth(xl, before, truth, lay)


def stacked(xl, seed):
    """積み上げ縦棒（2026-09-18）: ピボットを値で貼り付けた表（行ラベル・総計）／合計の列が「計」／課が 1 つ。"""
    S = _mod('グラフ積み上げ')
    out = _out('グラフ積み上げ')
    from openpyxl import load_workbook as _lw

    def pivot_paste(p, lay):
        wb = _lw(p)
        ws = wb[S.SHEET]
        ws.insert_rows(1, 2)
        ws['A1'] = '合計 / 支出額'
        ws['B1'] = '列ラベル'
        h = lay['h'] + 2
        ws.cell(h, 1).value = '行ラベル'
        for c in range(1, lay['n_km'] + 3):
            if ws.cell(h, c).value == '合計':
                ws.cell(h, c).value = '総計'
            if ws.cell(h + lay['n_ka'] + 1, c).value == '合計':
                ws.cell(h + lay['n_ka'] + 1, c).value = '総計'
        for r in range(h + 1, h + lay['n_ka'] + 2):            # 式は値に（貼り付けた表）
            for c in range(2, lay['n_km'] + 3):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith('='):
                    ws.cell(r, c).value = None
        wb.save(p)
        lay['h'] = h

    def kei(p, lay):
        wb = _lw(p)
        ws = wb[S.SHEET]
        ws.cell(lay['h'], lay['n_km'] + 2).value = '計'
        ws.cell(lay['h'] + lay['n_ka'] + 1, 1).value = '計'
        wb.save(p)

    jobs = [('実物01_ピボットの貼り付け', S.build(seed * 100 + 1, 6, 4), {}, pivot_paste),
            ('実物02_合計が計', S.build(seed * 100 + 2, 5, 3), {}, kei),
            ('実物03_課が1つ', S.build(seed * 100 + 3, 1, 5), {}, None)]
    for stem, data, kw, post in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        lay = S.write_before(before, data, **kw)
        if post:
            post(before, lay)
        S.build_truth(xl, before, truth, lay)


def pq(xl, seed):
    Q = _mod('パワークエリ')
    out = _out('パワークエリ')
    for stem, months in (('実物01_1か月だけ', [4]), ('実物02_種違い', [4, 5, 6])):
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        Q.write_before(before, Q.build(seed * 100 + len(months), months, 25))
        wb = xl.Workbooks.Open(before)
        try:
            Q.load_query(wb)
            wb.Worksheets(1).Activate()
            wb.SaveAs(truth, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)


if __name__ == '__main__':
    seed = SEED = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else 1
    only = sys.argv[2] if len(sys.argv) > 2 else ''
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for name, fn in (('課別棒グラフ', bar), ('推移折れ線', trend), ('構成比円', pie), ('ピボット', pivot), ('テーブル', table),
                         ('パワークエリ', pq), ('見やすい棒グラフ', mieru), ('課別月別ピボット', mpivot), ('PQ縦持ち', pqlong), ('課別ダッシュボード', dash), ('予算執行複合グラフ', combo), ('課別科目別積み上げ', stacked)):
            if only and only != name:
                continue
            fn(xl, seed)
            print(f"{name}: 作りました")
    finally:
        xl = None
        vf._quit(pid)
