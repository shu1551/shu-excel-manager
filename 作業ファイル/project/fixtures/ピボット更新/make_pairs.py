# -*- coding: utf-8 -*-
"""明細を足した後に、ピボットの元の範囲を広げて更新する（鍛える回路の題材・ピボット 5 本目・直す仕事・2026-09-18）。

仕事: 「支出明細」を元にしたピボット（ブックの中の全部）の元の範囲を、明細の見出しから最後の明細の行までに広げて更新する。
正解の決まり（依頼文に全部書く）:
  - 明細の下の合計の行は入れない。ピボットの行・列・値・置き場所・見た目は変えない（元の範囲と中身だけ新しくする）。
  - 支出明細を元にしないピボットは触らない。
組: 本番 60 行に 20 行足した／A 1,500 行に 300 行／B 表題・見出しが 4 行目／C 明細の下に合計の行／D 同じ明細のピボットが 2 つ（2 シート）／
  E もう最新（何も変わらない）／F 撃った後にもう一度／G 別の表のピボットもある（触らない）
"""
import importlib.util
import os
import random
import shutil
import sys

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('mp_pivot', os.path.join(HERE, '..', 'ピボット', 'make_pairs.py'))
P = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(P)
SHEET = P.SHEET
REQUEST = ('明細を足したので、支出明細を元にしたピボット（ブックの中の全部）の元の範囲を、明細の見出しから最後の明細の行まで広げて更新してください。'
           '明細の下の合計の行は入れない。ピボットの行・列・値・置き場所・見た目は変えない。支出明細を元にしないピボットは触らない')


def pivot_on(wb, lay, last, sheet, rows_field='課名', cols_field='科目'):
    src = wb.Sheets(SHEET)
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(last, lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = sheet
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{SHEET}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{sheet}'!R3C1", TableName='P' + sheet)
    pt.PivotFields(rows_field).Orientation = 1
    if cols_field:
        pt.PivotFields(cols_field).Orientation = 2
    df = pt.AddDataField(pt.PivotFields('支出額'), '合計 / 支出額', -4157)
    df.NumberFormat = '#,##0'
    return pt


def other_table_pivot(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = '予算'
    ws.Range('A1:B1').Value = ('課名', '予算額')
    for i, (k, v) in enumerate([('総務課', 100000), ('財政課', 200000), ('企画課', 150000)], 2):
        ws.Cells(i, 1).Value, ws.Cells(i, 2).Value = k, v
    pc = wb.PivotCaches().Create(SourceType=1, SourceData="'予算'!R1C1:R3C2")         # わざと 1 行足りない（触らない）
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = '予算集計'
    pt = pc.CreatePivotTable(TableDestination="'予算集計'!R3C1", TableName='P予算集計')
    pt.PivotFields('課名').Orientation = 1
    pt.AddDataField(pt.PivotFields('予算額'), '合計 / 予算額', -4157)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, rows, added, kw in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            two = kw.pop('two', False)
            other = kw.pop('other', False)
            lay = P.write_before(before, rows, **kw)
            wb = xl.Workbooks.Open(before)
            try:
                old_last = lay['last'] - added
                pivot_on(wb, lay, old_last, '課別集計')
                if two:
                    pivot_on(wb, lay, old_last, '月別集計', rows_field='支出日', cols_field=None)
                if other:
                    other_table_pivot(wb)
                wb.Worksheets(SHEET).Activate()
                wb.Save()
                src = wb.Sheets(SHEET)
                new = f"'{SHEET}'!{src.Range(src.Cells(lay['h'], 1), src.Cells(lay['last'], lay['ncols'])).GetAddress(True, True, -4150)}"
                for sh in wb.Worksheets:
                    for i in range(1, int(sh.PivotTables().Count) + 1):
                        pt = sh.PivotTables(i)
                        if SHEET in str(pt.SourceData):
                            pt.ChangePivotCache(wb.PivotCaches().Create(SourceType=1, SourceData=new))
                            pt.RefreshTable()
                wb.Worksheets(SHEET).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', P.build(s(1), 90), rng.randint(5, 40), {}),
          ('未見2_表題と合計行と2つ', P.build(s(2), 80), 15, {'title': True, 'total': True, 'two': True}),
          ('未見3_並び違いと別の表', P.build(s(3), 50), 10, {'cols': P.SWAP, 'other': True}),
          ('未見4_大量', P.build(s(4), 3000, n_ka=10), 700, {})], out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    make([('更新_本番', P.build(3, 80), 20, {}),
          ('更新_試験A', P.build(13, 1800, n_ka=10), 300, {}),
          ('更新_試験B', P.build(23, 60), 12, {'title': True}),
          ('更新_試験C', P.build(33, 50), 8, {'total': True}),
          ('更新_試験D', P.build(43, 70), 25, {'two': True}),
          ('更新_試験E', P.build(53, 40), 0, {}),
          ('更新_試験G', P.build(63, 45), 9, {'other': True})], HERE)
    shutil.copyfile(os.path.join(HERE, '更新_本番_正解.xlsx'), os.path.join(HERE, '更新_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '更新_本番_正解.xlsx'), os.path.join(HERE, '更新_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    lines = ['明細を追加したのでピボットの範囲を広げて', 'ピボットのデータソースを最後の行まで変更して更新', '足した明細がピボットに入っていないので直して',
             'ピボットの元データの範囲を更新して', '支出明細の増えた行までピボットに反映して', 'ピボットテーブルのソース範囲を広げて最新に']
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('ok')
