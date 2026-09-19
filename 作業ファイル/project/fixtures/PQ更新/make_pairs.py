# -*- coding: utf-8 -*-
"""ブックのパワークエリの読み込み先とピボットを、まとめて最新にする（G パワークエリ 62・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い。アクティブなブック全体。
正解の決まり（依頼文に全部書く）:
  - クエリの読み込み先のテーブル（ListObject の QueryTable）を全部、待って更新（BackgroundQuery:=False）してから、
    ピボットテーブルを全部更新（PivotCache.Refresh）する。元のテーブルは変えない。
組: 本番 クエリ 1 本／A クエリ 2 本／B クエリとピボット／C 元のテーブルに行を足した／D ピボットだけ／E 別のシートにある元／
  F 撃った後にもう一度／G 3 本・シートの並びが違う／H クエリもピボットも無い
"""
import importlib.util
import os
import random
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('tate', os.path.join(HERE, '..', '縦持ち', 'make_pairs.py'))
tate = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(tate)
import detail   # noqa: E402

REQUEST = ('このブックのパワークエリの読み込み先と、ピボットテーブルを、まとめて最新にしてください。'
           'すべてのシートの、クエリを読み込んだテーブル（ListObject の QueryTable があるもの）を全部、終わるのを待って更新し（QueryTable.Refresh BackgroundQuery:=False）、'
           'その後で、すべてのシートのピボットテーブルを全部更新する（PivotCache.Refresh）。元のテーブルや、ほかのセルは変えない。何も無ければ何もしない。')
PHRASES = ['クエリを全部更新して', 'パワークエリとピボットを最新にして', '読み込んだクエリを更新してください',
           'すべて更新をお願い', 'クエリの結果を最新の状態に', 'ピボットとクエリをまとめて更新']


def build(seed, n, queries=1, pivot=False, grow=False, src_other=False, **kw):
    t = detail.design(seed, n, **kw)
    t.update({'queries': queries, 'pivot': pivot, 'grow': grow, 'src_other': src_other,
              'table': random.Random(seed).choice(['T支出', 'T明細', 'Tデータ']) + str(seed % 5)})
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = 'はじめに' if t['src_other'] else t['sheet']
    if t['src_other']:
        ws['A1'], ws['A2'] = '集計の手順', '明細を貼って更新する'
        ws2 = wb.create_sheet(t['sheet'])
    else:
        ws2 = ws
    lay = detail.write_into(ws2, t)
    if t['queries'] or t['pivot']:
        tb = Table(displayName=t['table'], ref=f"A1:{get_column_letter(lay['ncols'])}{lay['last']}")
        tb.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws2.add_table(tb)
    wb.save(path)
    return lay


def group_m(t, key, name):
    amt = detail.name_of(t, 'num')
    return f'''let
    Source = Excel.CurrentWorkbook(){{[Name="{t['table']}"]}}[Content],
    Grouped = Table.Group(Source, {{"{key}"}}, {{{{"{name}", each List.Sum([#"{amt}"]), type number}}}})
in
    Grouped'''


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = write_before(before, t)
            wb = xl.Workbooks.Open(before)
            try:
                keys = [detail.name_of(t, 'item'), detail.name_of(t, 'kamoku'), detail.name_of(t, 'date')]
                for q in range(t['queries']):
                    nm = f"{t['table']}_集計{q + 1}"
                    tate.load_query(wb, nm, group_m(t, keys[q], '合計'), nm)
                if t['pivot']:
                    src = wb.Worksheets(t['sheet'])
                    ps = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                    ps.Name = 'ピボット'
                    pc = wb.PivotCaches().Create(1, t['table'])
                    pt = pc.CreatePivotTable("'ピボット'!R3C1", 'P1')
                    pt.PivotFields(keys[0]).Orientation = 1
                    pt.AddDataField(pt.PivotFields(detail.name_of(t, 'num')), '合計 / 金額', -4157)
                # 元のテーブルを書き換える（更新しないで保存＝読み込み先とピボットは古いまま）
                src = wb.Worksheets(t['sheet'])
                lo = src.ListObjects(1) if int(src.ListObjects.Count) else None
                nc = detail.col_of(t, 'num')
                rng = random.Random(t['seed'] + 3)
                for r in range(2, lay['last'] + 1, 3):
                    if src.Cells(r, nc).Value is not None:
                        src.Cells(r, nc).Value = rng.randrange(10, 900) * 100
                if t['grow'] and lo is not None:
                    last = lay['last']
                    for k in range(1, 4):
                        for c in range(1, lay['ncols'] + 1):
                            src.Cells(last + k, c).Value = src.Cells(2 + k, c).Value
                    lo.Resize(src.Range(src.Cells(1, 1), src.Cells(last + 3, lay['ncols'])))
                wb.Worksheets(1).Activate()
                wb.Save()
            finally:
                wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                for sh in wb.Worksheets:
                    for i in range(1, int(sh.ListObjects.Count) + 1):
                        lo = sh.ListObjects(i)
                        try:
                            qt = lo.QueryTable
                        except Exception:
                            continue
                        qt.Refresh(False)
                for sh in wb.Worksheets:
                    for i in range(1, int(sh.PivotTables().Count) + 1):
                        sh.PivotTables(i).PivotCache().Refresh()
                wb.Worksheets(1).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem, t['sheet'], t['table'], 'クエリ', t['queries'], 'ピボット' if t['pivot'] else '')
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(15, 40), queries=rng.randint(1, 2))),
          ('未見2_ピボットつき', build(s(2), 20, queries=1, pivot=True)),
          ('未見3_行を足した', build(s(3), 18, queries=2, grow=True)),
          ('未見4_元が別のシート', build(s(4), 16, queries=1, pivot=True, src_other=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('更新_本番', build(3901, 20)),
          ('更新_試験A', build(3902, 30, queries=2)),
          ('更新_試験B', build(3903, 20, queries=1, pivot=True)),
          ('更新_試験C', build(3904, 20, queries=1, grow=True)),
          ('更新_試験D', build(3905, 20, queries=0, pivot=True)),
          ('更新_試験E', build(3906, 20, queries=1, src_other=True)),
          ('更新_試験G', build(3907, 25, queries=3)),
          ('更新_試験H', build(3908, 15, queries=0))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '更新_本番_正解.xlsx'), os.path.join(HERE, f'更新_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
