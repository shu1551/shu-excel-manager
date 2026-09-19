# -*- coding: utf-8 -*-
"""支出明細から、課別のピボット＋ピボットグラフ＋科目のスライサーを 1 枚のシートに作る（鍛える回路の題材・ピボット 3 本目・2026-09-18）。

仕事: 新しいシート「課別ダッシュボード」の A3 にピボット（行＝課名・値＝支出額の合計 #,##0・行の総計あり）、その右に
      ピボットグラフ（集合縦棒・タイトル「課別支出額」・凡例なし）、グラフの右に科目のスライサーを置く。
正解の決まり（依頼文に全部書く）:
  - 元の範囲は明細の見出しから最後の明細の行まで（下の合計の行は入れない）。見出しの語で列を探す。
  - 「課別ダッシュボード」シートが既にあれば、スライサーも含めて消して作り直す（2 回撃ってもピボット・グラフ・スライサーは 1 つずつ）。
組: 本番 60 行／A 1,500 行／B 表題・見出しが 4 行目／C 前のダッシュボード（スライサーつき）が残っている／D 明細の下に合計の行／
  E 列の並びが違う／F 撃った後にもう一度
"""
import importlib.util
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('mp_pivot', os.path.join(HERE, '..', 'ピボット', 'make_pairs.py'))
P = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(P)
SHEET = P.SHEET
OUT_SHEET = '課別ダッシュボード'
REQUEST = ('支出明細から、新しいシート「課別ダッシュボード」を作り、A3 にピボットテーブル（行は課名、値は支出額の合計（表示形式 #,##0））、'
           'その右にピボットグラフ（集合縦棒・タイトル「課別支出額」・凡例なし）、グラフの右に科目のスライサーを置いてください。'
           '元の範囲は明細の見出しから最後の明細の行まで（下の合計の行は入れない）。シートが既にあればスライサーも含めて消して作り直す')


def build_dashboard(wb, lay, slicer_field='科目', title='課別支出額'):
    xl = wb.Application
    for sh in list(wb.Worksheets):
        if str(sh.Name) == OUT_SHEET:
            xl.DisplayAlerts = False
            sh.Delete()
    for i in range(int(wb.SlicerCaches.Count), 0, -1):          # 持ち主のピボットが消えたスライサーの残り
        sc = wb.SlicerCaches.Item(i)
        if int(sc.PivotTables.Count) == 0:
            sc.Delete()
    src = wb.Sheets(SHEET)
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(lay['last'], lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = OUT_SHEET
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{SHEET}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{OUT_SHEET}'!R3C1", TableName='P課別ダッシュボード')
    pt.PivotFields('課名').Orientation = 1
    df = pt.AddDataField(pt.PivotFields('支出額'), '合計 / 支出額', -4157)
    df.NumberFormat = '#,##0'
    pt.RowGrand = True
    left = out.Range('D3').Left
    co = out.ChartObjects().Add(left, out.Range('D3').Top, 400, 250)
    ch = co.Chart
    ch.SetSourceData(pt.TableRange1)
    ch.ChartType = 51
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.HasLegend = False
    sc = wb.SlicerCaches.Add2(pt, slicer_field)
    sc.Slicers.Add(SlicerDestination=out, Caption=slicer_field, Top=out.Range('D3').Top, Left=left + 420, Width=144, Height=200)
    return pt


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, rows, kw in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            stale = kw.pop('stale', False)
            lay = P.write_before(before, rows, **kw)
            if stale:                                         # 前のダッシュボード（スライサーは課名・タイトル違い）
                wb = xl.Workbooks.Open(before)
                try:
                    build_dashboard(wb, lay, slicer_field='課名', title='旧')
                    wb.Worksheets(SHEET).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                build_dashboard(wb, lay)
                wb.Worksheets(SHEET).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    import random
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', P.build(s(1), rng.randint(30, 120), n_ka=rng.randint(3, 8)), {}),
          ('未見2_表題と合計行', P.build(s(2), 80), {'title': True, 'total': True}),
          ('未見3_並び違いと前の結果', P.build(s(3), 50), {'cols': P.SWAP, 'stale': True}),
          ('未見4_大量', P.build(s(4), 3000, n_ka=10), {})], out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    make([('ダッシュ_本番', P.build(3, 60), {}),
          ('ダッシュ_試験A', P.build(13, 1500, n_ka=10), {}),
          ('ダッシュ_試験B', P.build(23, 50), {'title': True}),
          ('ダッシュ_試験C', P.build(33, 40), {'stale': True}),
          ('ダッシュ_試験D', P.build(43, 45), {'total': True}),
          ('ダッシュ_試験E', P.build(53, 55), {'cols': P.SWAP})], HERE)
    shutil.copyfile(os.path.join(HERE, 'ダッシュ_本番_正解.xlsx'), os.path.join(HERE, 'ダッシュ_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, 'ダッシュ_本番_正解.xlsx'), os.path.join(HERE, 'ダッシュ_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    lines = ['支出明細から課別のダッシュボードを作って', 'ピボットとピボットグラフとスライサーを 1 枚に', '課別ダッシュボードのシートをお願い',
             'ピボットグラフに科目のスライサーを付けて', '課別の集計をピボットグラフで見られるように', '支出明細のダッシュボード（ピボット・グラフ・スライサー）']
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('ok')
