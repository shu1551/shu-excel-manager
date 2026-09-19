# -*- coding: utf-8 -*-
"""明細を足した後に、ピボットの元の範囲を広げて更新する（どの表でも動く形・2026-09-18 作り直し。前の版はピボット範囲更新）。

仕事: アクティブなシートを元にしたピボット（ブックの中の全部）の元の範囲を、そのシートの見出しから最後の行までに広げて更新する。
見出しの語・シート名・項目の名前は表ごとに違う（vocab.py・detail.py）。
正解の決まり:
  - 明細の下の合計の行は入れない。ピボットの行・列・値・置き場所・見た目は変えない（元の範囲と中身だけ新しくする）。
  - アクティブなシートを元にしないピボットは触らない。
組: 本番 80 行に 20 行足した／A 1,800 行に 300 行／B 表題・単位の行／C 明細の下に合計の行／D 同じ明細のピボットが 2 つ／
  E もう最新（何も変わらない）／F 撃った後にもう一度／G 別のシートのピボットもある（触らない）／H 列の並び違い・番号の列
"""
import os
import random
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('明細を足したので、アクティブなシートを元にしたピボット（ブックの中の全部）の元の範囲を、'
           'そのシートの見出しの行から最後の明細の行まで広げて更新してください。明細の下の合計の行は入れない。'
           'ピボットの行・列・値・置き場所・見た目は変えない。アクティブなシートを元にしないピボットは触らない')
PHRASES = ['明細を追加したのでピボットの範囲を広げて', 'ピボットのデータソースを最後の行まで変更して更新',
           '足した明細がピボットに入っていないので直して', 'ピボットの元データの範囲を更新して',
           '増えた行までピボットに反映して', 'ピボットテーブルのソース範囲を広げて最新に']


def pivot_on(wb, t, lay, last, out_name, rows_role='item', cols_role='kamoku'):
    src = wb.Worksheets(t['sheet'])
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(last, lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = out_name
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{t['sheet']}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{out_name}'!R3C1", TableName='P' + out_name)
    pt.PivotFields(detail.name_of(t, rows_role)).Orientation = 1
    if cols_role:
        pt.PivotFields(detail.name_of(t, cols_role)).Orientation = 2
    df = pt.AddDataField(pt.PivotFields(detail.name_of(t, 'num')), '合計 / ' + detail.name_of(t, 'num'), -4157)
    df.NumberFormat = '#,##0'
    return pt


def other_sheet_pivot(wb, t):
    """別のシートの表を元にしたピボット（触らない）。わざと 1 行足りない範囲で作る。"""
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = '別表'
    ws.Range('A1:B1').Value = (t['ka'], '予算額')
    for i, k in enumerate(t['items'][:3], 2):
        ws.Cells(i, 1).Value, ws.Cells(i, 2).Value = k, 100000 * i
    pc = wb.PivotCaches().Create(SourceType=1, SourceData="'別表'!R1C1:R3C2")
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = '別表集計'
    pt = pc.CreatePivotTable(TableDestination="'別表集計'!R3C1", TableName='P別表集計')
    pt.PivotFields(t['ka']).Orientation = 1
    pt.AddDataField(pt.PivotFields('予算額'), '合計 / 予算額', -4157)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, added, kw in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = detail.write_before(before, t)
            wb = xl.Workbooks.Open(before)
            try:
                old_last = lay['last'] - added
                pivot_on(wb, t, lay, old_last, '集計1')
                if kw.get('two'):
                    pivot_on(wb, t, lay, old_last, '集計2', rows_role='date', cols_role=None)
                if kw.get('other'):
                    other_sheet_pivot(wb, t)
                wb.Worksheets(t['sheet']).Activate()
                wb.Save()
                src = wb.Worksheets(t['sheet'])
                new = (f"'{t['sheet']}'!"
                       + src.Range(src.Cells(lay['h'], 1), src.Cells(lay['last'], lay['ncols'])).GetAddress(True, True, -4150))
                for sh in wb.Worksheets:
                    for i in range(1, int(sh.PivotTables().Count) + 1):
                        pt = sh.PivotTables(i)
                        if t['sheet'] in str(pt.SourceData):
                            pt.ChangePivotCache(wb.PivotCaches().Create(SourceType=1, SourceData=new))
                            pt.RefreshTable()
                wb.Worksheets(t['sheet']).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem, t['sheet'], [c for c, _r in t['cols']])
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', detail.design(s(1), 90), rng.randint(5, 40), {}),
          ('未見2_表題と合計行と2つ', detail.design(s(2), 80, title=True, unit=True, total=True), 15, {'two': True}),
          ('未見3_並び違いと別のシート', detail.design(s(3), 50, swap=True, no_col=True), 10, {'other': True}),
          ('未見4_大量', detail.design(s(4), 1200), 400, {})], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('更新_本番', detail.design(3, 80), 20, {}),
          ('更新_試験A', detail.design(13, 1800), 300, {}),
          ('更新_試験B', detail.design(23, 60, title=True, unit=True), 12, {}),
          ('更新_試験C', detail.design(33, 50, total=True), 8, {}),
          ('更新_試験D', detail.design(43, 70), 25, {'two': True}),
          ('更新_試験E', detail.design(53, 40), 0, {}),
          ('更新_試験G', detail.design(63, 45), 9, {'other': True}),
          ('更新_試験H', detail.design(73, 55, swap=True, no_col=True), 14, {})], HERE)
    shutil.copyfile(os.path.join(HERE, '更新_本番_正解.xlsx'), os.path.join(HERE, '更新_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '更新_本番_正解.xlsx'), os.path.join(HERE, '更新_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
