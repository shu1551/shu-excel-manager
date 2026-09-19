# -*- coding: utf-8 -*-
"""選んでいる 2 列（1 つ目＝横軸・2 つ目＝縦軸）で散布図を作る（E グラフ 41・2026-09-18 第二期）。

列の決まり: 選んでいる 2 列。<名前>_選択.txt に書く。表は集計表（shukei）。
正解の決まり（依頼文に全部書く）:
  - 本文＝見出しの次の行から最後の行まで。2 列とも数の行だけ（空行と、行のどこかに「合計」「計」「総計」とある行は入れない）。
  - 散布図（マーカーのみ・xlXYScatter）・系列 1 本（名前＝縦軸の列の見出し・横の値＝1 つ目の列・縦の値＝2 つ目の列・表の上から順）。
  - タイトル「（横の見出し）と（縦の見出し）の関係」14pt 太字。横軸・縦軸のタイトルにそれぞれの見出し。両軸 #,##0・目盛線なし。
  - 凡例なし・グラフの文字 Meiryo UI。前に作ったグラフ（このシートのグラフ全部）は消して作り直す。
組: 本番／A 40 行／B 表題と単位／C 合計の行／D 途中の空行／E 数の列が左（並び違い）／F 撃った後にもう一度／G 番号の列と摘要／H 前年度・今年度
"""
import os
import random
import shutil
import sys

from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import shukei   # noqa: E402

REQUEST = ('選んでいる 2 列（1 つ目＝横軸、2 つ目＝縦軸）で散布図を作ってください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '本文は見出しの次の行から最後の行まで。2 列とも数が入っている行だけを使い、空の行と、行のどこかのセル（ふつうは左の項目の列）に'
           '「合計」「計」「総計」とある行は入れない（合計の行も 2 列に数が入っているので、数だけでは見分けられない）。'
           '種類は散布図（マーカーのみ・xlXYScatter）。系列は 1 本で、名前＝縦軸の列の見出し、XValues＝1 つ目の列の値、Values＝2 つ目の列の値（表の上から順）。'
           'タイトルは「（1 つ目の見出し）と（2 つ目の見出し）の関係」で 14pt・太字。横軸（Axes(xlCategory)）のタイトルに 1 つ目の見出し、'
           '縦軸（Axes(xlValue)）のタイトルに 2 つ目の見出し。両方の軸の表示形式は #,##0、目盛線なし。凡例なし。'
           'グラフ全体の文字は Meiryo UI（全体に当ててからタイトルの大きさ・太字を当てる）。前に作ったグラフ（このシートのグラフ全部）は消して作り直す。置き場所は表の右。')
PHRASES = ['選んだ2列で散布図を作って', '散布図にしてください', '2つの列の関係を散布図で見たい',
           '予算と執行の関係を点で見せて', '相関が分かるグラフにして', '選んでいる列を散布図に']


def build_truth(xl, before, truth, t, lay, nx, ny):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        h, c0 = lay['h'], lay['c0']
        cx = c0 + shukei.col_of(t, 'num', nx) - 1
        cy = c0 + shukei.col_of(t, 'num', ny) - 1
        xs, ys = [], []
        for r in range(h + 1, lay['last'] + 1):
            a, b = ws.Cells(r, cx).Value, ws.Cells(r, cy).Value
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                xs.append(a)
                ys.append(b)
        hx, hy = ws.Cells(h, cx).Value, ws.Cells(h, cy).Value
        co = ws.ChartObjects().Add(ws.Cells(h, c0 + lay['ncols'] + 1).Left, ws.Cells(h, 1).Top, 480, 300)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = -4169
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        s = ch.SeriesCollection().NewSeries()
        s.Name = hy
        s.XValues = xs
        s.Values = ys
        ch.HasTitle = True
        ch.ChartTitle.Text = f"{hx}と{hy}の関係"
        ch.HasLegend = False
        for idx, txt in ((1, hx), (2, hy)):
            ax = ch.Axes(idx)
            ax.HasTitle = True
            ax.AxisTitle.Text = txt
            ax.TickLabels.NumberFormat = '#,##0'
            ax.HasMajorGridlines = False
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = 'Meiryo UI'
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def old_chart(xl, path, t, lay):
    wb = xl.Workbooks.Open(path)
    try:
        ws = wb.Worksheets(t['sheet'])
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], 8).Left, ws.Cells(lay['h'], 8).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 51
        co.Chart.SeriesCollection().NewSeries().Values = [1, 2, 3]
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, old in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = shukei.write_before(before, t)
            if old:
                old_chart(xl, before, t, lay)
            build_truth(xl, before, truth, t, lay, 0, 1)
            sel = ",".join(f"{get_column_letter(shukei.col_of(t, 'num', k))}{lay['h']}" for k in (0, 1))
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def d(seed, n, **kw):
    kw.setdefault('nums', 2)
    return shukei.design(seed, n, **kw)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', d(s(1), rng.randint(6, 14)), False),
          ('未見2_表題と合計行', d(s(2), 9, title=True, unit=True, total=True), True),
          ('未見3_並び違い', d(s(3), 8, swap=True, prev=True), False),
          ('未見4_空行と番号列', d(s(4), 10, blanks=2, no_col=True), False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('散布_本番', d(2701, 10), False),
          ('散布_試験A', d(2702, 40), False),
          ('散布_試験B', d(2703, 8, title=True, unit=True), False),
          ('散布_試験C', d(2704, 10, total=True), False),
          ('散布_試験D', d(2705, 11, blanks=2), False),
          ('散布_試験E', d(2706, 9, swap=True), False),
          ('散布_試験G', d(2707, 7, no_col=True, text=True), True),
          ('散布_試験H', d(2708, 8, prev=True), False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '散布_本番_正解.xlsx'), os.path.join(HERE, f'散布_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '散布_本番_選択.txt'), os.path.join(HERE, '散布_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
