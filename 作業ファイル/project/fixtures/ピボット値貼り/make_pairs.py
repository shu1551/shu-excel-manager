# -*- coding: utf-8 -*-
"""アクティブなシートのピボットテーブルを、同じ場所に値で貼り付けて普通の表にする（F ピボット 49・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い。アクティブなシートのピボット全部が対象（ほかのシートのピボットは触らない）。
正解の決まり（依頼文に全部書く）:
  - ピボットごとに、フィルタの欄も含めた全体（TableRange2）の値を、同じ場所に値で書き戻し、ピボットは無くす。
  - 見た目の文字（行ラベル・総計・「合計 / …」）もそのまま値で残す。元の明細・ほかのシートは変えない。ピボットが無ければ何もしない。
組: 本番 行＝課／A 行＝課・列＝科目／B 同じシートに 2 つ／C フィルタつき／D 表題の下（A3）／E 行 2 つ（課・科目）／
  F 撃った後にもう一度／G ほかのシートにもピボット（それは残す）／H 明細と同じシートの右
"""
import os
import random
import shutil
import sys

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('アクティブなシートにあるピボットテーブルを全部、同じ場所に値で貼り付けて、普通の表にしてください。'
           'ピボットごとに、フィルタの欄も含めた全体（PivotTable.TableRange2）の値を、同じセルに値で書き戻し、ピボットは無くす'
           '（値を控えてから TableRange2.Clear し、同じ範囲に値を書く）。行ラベル・総計・「合計 / …」などの見た目の文字も、そのまま値で残す。'
           'ほかのシートのピボットと、元の明細は変えない。ピボットが無ければ何もしない。')
PHRASES = ['ピボットを値で貼り付けて', 'ピボットテーブルを普通の表にして', 'ピボットを値だけの表に変換',
           'ピボットを外して集計の結果だけ残して', 'ピボットを静的な表にして', 'このピボットを値貼りして']


def write_before(path, t, psheet, same):
    wb = Workbook()
    if same:
        ws = wb.active
        ws.title = t['sheet']
        detail.write_into(ws, t)
    else:
        wb.active.title = psheet                        # ピボットのシートを先頭に（道具は先頭のシートを撃つ）
        ws = wb.create_sheet(t['sheet'])
        detail.write_into(ws, t)
    wb.save(path)


def add_pivots(xl, path, t, psheet, kinds, same, other):
    wb = xl.Workbooks.Open(path)
    try:
        src = wb.Worksheets(t['sheet'])
        nc = len(t['cols'])
        last = int(src.Cells(src.Rows.Count, 1).End(-4162).Row)
        item, kam = detail.name_of(t, 'item'), detail.name_of(t, 'kamoku')
        amt = detail.name_of(t, 'num')
        dest_sheet = t['sheet'] if same else psheet
        for k, kind in enumerate(kinds):
            pc = wb.PivotCaches().Create(1, f"'{t['sheet']}'!R1C1:R{last}C{nc}")
            if same:
                dest = f"'{dest_sheet}'!R1C{nc + 2 + k * 6}"
            elif kind == 'title':
                wb.Worksheets(psheet).Cells(1, 1).Value = '令和8年度 支出の集計'
                dest = f"'{dest_sheet}'!R3C1"
            elif kind == 'page':
                dest = f"'{dest_sheet}'!R3C{1 + k * 6}"
            else:
                dest = f"'{dest_sheet}'!R1C{1 + k * 6}"
            pt = pc.CreatePivotTable(dest, f"P{k + 1}")
            pt.PivotFields(item).Orientation = 1
            if kind == 'rows2':
                pt.PivotFields(kam).Orientation = 1
            if kind == 'cols':
                pt.PivotFields(kam).Orientation = 2
            if kind == 'page':
                pt.PivotFields(kam).Orientation = 3
            pt.AddDataField(pt.PivotFields(amt), f"合計 / {amt}", -4157)
        if other:
            os_ = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
            os_.Name = '別の集計'
            pc = wb.PivotCaches().Create(1, f"'{t['sheet']}'!R1C1:R{last}C{nc}")
            pt = pc.CreatePivotTable("'別の集計'!R1C1", 'P9')
            pt.PivotFields(kam).Orientation = 1
            pt.AddDataField(pt.PivotFields(amt), f"合計 / {amt}", -4157)
        wb.Worksheets(1).Activate()
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def build_truth(xl, before, truth):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(1)
        for i in range(int(ws.PivotTables().Count), 0, -1):
            pt = ws.PivotTables(i)
            rng = pt.TableRange2
            addr, vals = rng.Address, rng.Value
            rng.Clear()
            ws.Range(addr).Value = vals
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kinds, same, other in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            psheet = random.Random(t['seed']).choice(['ピボット', '集計', '課別集計', 'Sheet2'])
            write_before(before, t, psheet, same)
            if kinds:
                add_pivots(xl, before, t, psheet, kinds, same, other)
            build_truth(xl, before, truth)
            print('  ', stem, t['sheet'], kinds, '同じシート' if same else psheet, 'ほかのシートにも' if other else '')
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', detail.design(s(1), rng.randint(20, 60)), ['rows'], False, False),
          ('未見2_列とフィルタ', detail.design(s(2), 30), ['cols', 'page'], False, False),
          ('未見3_表題と別シート', detail.design(s(3), 25), ['title'], False, True),
          ('未見4_同じシート', detail.design(s(4), 28), ['rows2'], True, False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('値貼り_本番', detail.design(3601, 30), ['rows'], False, False),
          ('値貼り_試験A', detail.design(3602, 200, n_items=10), ['cols'], False, False),
          ('値貼り_試験B', detail.design(3603, 25), ['rows', 'cols'], False, False),
          ('値貼り_試験C', detail.design(3604, 24), ['page'], False, False),
          ('値貼り_試験D', detail.design(3605, 26), ['title'], False, False),
          ('値貼り_試験E', detail.design(3606, 28), ['rows2'], False, False),
          ('値貼り_試験G', detail.design(3607, 22), ['rows'], False, True),
          ('値貼り_試験H', detail.design(3608, 20), ['rows'], True, False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '値貼り_本番_正解.xlsx'), os.path.join(HERE, f'値貼り_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
