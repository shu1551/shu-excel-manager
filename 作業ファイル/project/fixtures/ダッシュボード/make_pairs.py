# -*- coding: utf-8 -*-
"""選んでいる列でダッシュボード（ピボット＋ピボットグラフ＋スライサー）を作る
（どの表でも動く形・2026-09-18 作り直し。前の版は課別ダッシュボード）。

列の決まり: 選んでいる 3 列（1 つ目＝行、2 つ目＝値、3 つ目＝スライサーの列）。選んでいる列は <名前>_選択.txt に書く。
正解の決まり:
  - 新しいシート（名前は 行の列の見出し＋「別ダッシュボード」）。同じ名前のシートがあれば消して作り直す。
  - A3 にピボット（行＝1 つ目の列、値＝2 つ目の列の合計・表示形式 #,##0）。
  - その右にピボットグラフ（集合縦棒・タイトルは 行の見出し＋「別」＋値の見出し・凡例なし）。
  - グラフの右に 3 つ目の列のスライサー。前のスライサーは消す。
  - 元の範囲は明細の見出しの行から最後の明細の行まで（下の合計の行は入れない）。
組: 本番 60 行／A 1,500 行／B 表題と単位の行／C 前の結果のシートとスライサー／D 明細の下に合計の行／
  E 列の並びが違う／F 撃った後にもう一度／G 空行と番号の列
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

REQUEST = ('選んでいる 3 列（1 つ目を行、2 つ目を値、3 つ目をスライサーの列）で、新しいシートにダッシュボードを'
           '作ってください。新しいシートの名前は 行の列の見出しと「別ダッシュボード」をつないだ名前。'
           'A3 にピボット（行は 1 つ目の列、値は 2 つ目の列の合計・表示形式 #,##0）、'
           'その右にピボットグラフ（集合縦棒・タイトルは 行の見出しと「別」と値の見出しをつないだ文字・凡例なし）、'
           'グラフの右に 3 つ目の列のスライサーを置く。前のスライサーは消す。'
           '元の範囲は明細の見出しの行から最後の明細の行まで（下の合計の行は入れない）。'
           '同じ名前のシートが既にあれば消して作り直す')
PHRASES = ['選んだ列でダッシュボードを作って', 'ピボットとグラフとスライサーを1枚に', 'ダッシュボードにして',
           'スライサー付きのピボットとグラフを', '1枚にまとめたダッシュボードを', '選んでいる列でダッシュボード']


def out_name(t):
    return f"{detail.name_of(t, 'item')}別ダッシュボード"


def build_dash(wb, t, lay, last=None):
    xl = wb.Application
    name = out_name(t)
    for sh in list(wb.Worksheets):
        if str(sh.Name) == name:
            xl.DisplayAlerts = False
            sh.Delete()
    for i in range(int(wb.SlicerCaches.Count), 0, -1):
        wb.SlicerCaches.Item(i).Delete()
    src = wb.Worksheets(t['sheet'])
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(last or lay['last'], lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = name
    item, num, sl = (detail.name_of(t, 'item'), detail.name_of(t, 'num'), detail.name_of(t, 'kamoku'))
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{t['sheet']}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{name}'!R3C1", TableName='PT' + name)
    pt.PivotFields(item).Orientation = 1
    df = pt.AddDataField(pt.PivotFields(num), '合計 / ' + num, -4157)
    df.NumberFormat = '#,##0'
    co = out.ChartObjects().Add(out.Cells(3, 4).Left, out.Cells(3, 4).Top, 420, 260)
    ch = co.Chart
    ch.SetSourceData(pt.TableRange1)
    ch.ChartType = 51
    ch.HasTitle = True
    ch.ChartTitle.Text = f"{item}別{num}"
    ch.HasLegend = False
    scache = wb.SlicerCaches.Add2(pt, sl)
    slicer = scache.Slicers.Add(out)
    slicer.Top = out.Cells(3, 12).Top
    slicer.Left = out.Cells(3, 12).Left
    slicer.Width = 144
    slicer.Height = 200
    return pt


def make(jobs, out=HERE):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kw in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            lay = detail.write_before(before, t)
            sel = ",".join(f"{get_column_letter(detail.col_of(t, role))}{lay['h']}"
                           for role in ('item', 'num', 'kamoku'))
            if kw.get('stale'):
                wb = xl.Workbooks.Open(before)
                try:
                    build_dash(wb, t, lay, last=lay['h'] + 5)
                    wb.Worksheets(t['sheet']).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                build_dash(wb, t, lay)
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
          ('未見4_空行と番号の列', detail.design(s(4), 40, blanks=2, no_col=True), {})], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('ダッシュ_本番', detail.design(3, 60), {}),
          ('ダッシュ_試験A', detail.design(13, 1200, n_items=10), {}),
          ('ダッシュ_試験B', detail.design(23, 60, title=True, unit=True), {}),
          ('ダッシュ_試験C', detail.design(33, 50), {'stale': True}),
          ('ダッシュ_試験D', detail.design(43, 40, total=True), {}),
          ('ダッシュ_試験E', detail.design(53, 45, swap=True), {}),
          ('ダッシュ_試験G', detail.design(63, 35, blanks=2, no_col=True), {})])
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, 'ダッシュ_本番_正解.xlsx'), os.path.join(HERE, f'ダッシュ_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, 'ダッシュ_本番_選択.txt'), os.path.join(HERE, 'ダッシュ_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
