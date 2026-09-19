# -*- coding: utf-8 -*-
"""選んでいる列で「いつもの月別ピボット」を作る（どの表でも動く形・2026-09-18 作り直し。前の版は課別月別ピボット）。

列の決まり: 選んでいる 3 列（1 つ目＝行の外側、2 つ目＝行の内側、3 つ目＝値）と、明細の日付の列（月でグループ化して列に置く）。
依頼は短い（見せ方は書かない）＝正解のピボットが課で決めている型を持ち、AI は正解との差から学ぶ。
見せ方の型（正解の作り方）:
  - 新しいシート（名前は 行の外側の見出し＋「別」＋「月別」）の A3。同じ名前のシートがあれば消して作り直す。
  - 行＝1 つ目 → 2 つ目（表形式・ラベルを繰り返す・小計なし）、列＝日付の列を月でグループ化、値＝3 つ目の合計 #,##0。
  - 行と列の総計あり・空白のセルに 0 を表示・スタイル PivotStyleMedium9。
  - 元の範囲は明細の見出しの行から最後の明細の行まで（下の合計の行は入れない）。
組: 本番 60 行／A 1,500 行／B 表題と単位の行／C 前の結果のシート／D 明細の下に合計の行／E 列の並びが違う／
  F 撃った後にもう一度／G 1 か月だけ／H 空行と番号の列
"""
import os
import random
import shutil
import sys

from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('選んでいる 3 列で、いつもの月別のピボットにしてください。前に作ったものがあれば作り直す'
           '（1 つ目の列を行の外側、2 つ目を行の内側、3 つ目を値にする）')
PHRASES = ['選んだ列でいつもの月別ピボットに', '月別のピボット表をいつもの形で', 'いつもの月別のピボットを作って',
           '月ごとのピボットをいつもの見せ方で', '月別ピボットにして', '選んでいる列で月別のピボット']


def out_name(t):
    return f"{detail.name_of(t, 'item')}別月別"


def build_pivot(wb, t, lay, last=None):
    xl = wb.Application
    name = out_name(t)
    for sh in list(wb.Worksheets):
        if str(sh.Name) == name:
            xl.DisplayAlerts = False
            sh.Delete()
    src = wb.Worksheets(t['sheet'])
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(last or lay['last'], lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = name
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{t['sheet']}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{name}'!R3C1", TableName='P' + name)
    f1 = pt.PivotFields(detail.name_of(t, 'item'))
    f2 = pt.PivotFields(detail.name_of(t, 'kamoku'))
    f1.Orientation = 1
    f1.Position = 1
    f2.Orientation = 1
    f2.Position = 2
    for f in (f1, f2):
        f.LayoutForm = 0                      # 表形式
        f.RepeatLabels = True
        f.Subtotals = [False] * 12
    d = pt.PivotFields(detail.name_of(t, 'date'))
    d.Orientation = 2
    d.DataRange.Cells(1).Group(True, True, 1, [False, False, False, False, True, False, False])   # 月でグループ化
    df = pt.AddDataField(pt.PivotFields(detail.name_of(t, 'num')), '合計 / ' + detail.name_of(t, 'num'), -4157)
    df.NumberFormat = '#,##0'
    pt.RowGrand = True
    pt.ColumnGrand = True
    pt.NullString = '0'
    pt.DisplayNullString = True
    pt.TableStyle2 = 'PivotStyleMedium9'
    return pt


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kw in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            lay = detail.write_before(before, t)
            sel = ",".join(f"{get_column_letter(detail.col_of(t, role))}{lay['h']}"
                           for role in ('item', 'kamoku', 'num'))
            if kw.get('stale'):
                wb = xl.Workbooks.Open(before)
                try:
                    build_pivot(wb, t, lay, last=lay['h'] + 5)
                    wb.Worksheets(t['sheet']).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                build_pivot(wb, t, lay)
                wb.Worksheets(t['sheet']).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], out_name(t), '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', detail.design(s(1), rng.randint(30, 200)), {}),
          ('未見2_表題と合計行', detail.design(s(2), 80, title=True, unit=True, total=True), {}),
          ('未見3_並び違いと前の結果', detail.design(s(3), 50, swap=True), {'stale': True}),
          ('未見4_1か月と空行', detail.design(s(4), 30, gap_months=(4, 5, 6, 7, 8, 9, 10, 11, 1, 2, 3), blanks=2), {})], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('月別_本番', detail.design(3, 60), {}),
          ('月別_試験A', detail.design(13, 1500, n_items=10), {}),
          ('月別_試験B', detail.design(23, 60, title=True, unit=True), {}),
          ('月別_試験C', detail.design(33, 50), {'stale': True}),
          ('月別_試験D', detail.design(43, 40, total=True), {}),
          ('月別_試験E', detail.design(53, 45, swap=True), {}),
          ('月別_試験G', detail.design(63, 25, gap_months=(5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)), {}),
          ('月別_試験H', detail.design(73, 35, blanks=2, no_col=True), {})], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '月別_本番_正解.xlsx'), os.path.join(HERE, f'月別_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '月別_本番_選択.txt'), os.path.join(HERE, '月別_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
