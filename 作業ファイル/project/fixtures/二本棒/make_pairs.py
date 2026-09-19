# -*- coding: utf-8 -*-
"""選んでいる 2 列（1 つ目＝前・2 つ目＝後）を並べた棒グラフにする（E グラフ 43・2026-09-18 第二期）。

列の決まり: 選んでいる 2 列。<名前>_選択.txt に書く。項目の列＝表の左端の文字の列。表は集計表（shukei）。
正解の決まり（依頼文に全部書く）:
  - 本文＝見出しの次の行から最後の行まで。2 列とも数の行だけ（空行と「合計」「計」「総計」の行は入れない）。表の上から順。
  - 項目が 8 以下なら集合縦棒、9 以上なら集合横棒で上から表の順（項目軸を反転）。
  - 系列 1＝1 つ目の列（名前＝見出し・色 #A6A6A6）、系列 2＝2 つ目の列（色 #1F4E79）。項目＝項目の列の値。
  - タイトル: 表の上に表題があればその文字（「単位」の行は表題にしない）、無ければ「（1 つ目の見出し）と（2 つ目の見出し）の比較」14pt 太字。
  - 凡例は下・数値軸 #,##0・目盛線なし・棒の間隔 60・データラベルなし・グラフの文字 Meiryo UI。前に作ったグラフは消して作り直す。
組: 本番／A 12 行（横棒）／B 表題と単位／C 合計の行／D 途中の空行／E 数の列が左（並び違い）／F 撃った後にもう一度／G 番号の列と前のグラフ／H 前年度・今年度
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

REQUEST = ('選んでいる 2 列（1 つ目＝前、2 つ目＝後）を、項目ごとに 2 本並べた棒グラフにしてください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '項目の列＝表の左端の文字の列（数でも日付でもない列）。本文は見出しの次の行から最後の行まで。2 列とも数が入っている行だけを表の上から順に使い、'
           '空の行と、行のどこかのセル（ふつうは項目の列）に「合計」「計」「総計」とある行は入れない（合計の行も数が入っているので、数だけでは見分けられない）。'
           '項目が 8 個以下なら集合縦棒、9 個以上なら集合横棒にして項目軸を反転（上から表の順）。'
           '系列 1＝1 つ目の列（名前は見出し・塗りの色 #A6A6A6）、系列 2＝2 つ目の列（名前は見出し・塗りの色 #1F4E79）。項目（XValues）＝項目の列の値。'
           'タイトル: 表の上に表題があればその文字（「単位」の行は表題にしない）、無ければ「（1 つ目の見出し）と（2 つ目の見出し）の比較」。14pt・太字。'
           '凡例は下。数値軸（Axes(xlValue)）の表示形式 #,##0・目盛線なし。棒の間隔（ChartGroups(1).GapWidth）60。データラベルなし。'
           'グラフ全体の文字は Meiryo UI（全体に当ててからタイトルの大きさ・太字を当てる）。前に作ったグラフ（このシートのグラフ全部）は消して作り直す。置き場所は表の右。')
PHRASES = ['前年度と今年度を並べた棒グラフに', '選んだ2列を並べた棒グラフにして', '2本並びの棒グラフを作って',
           '予算と執行を2本の棒で並べて', '前と後を並べて棒で比べたい', '2列を横に並べた棒グラフで']
GRAY, BLUE = (0xA6, 0xA6, 0xA6), (0x1F, 0x4E, 0x79)


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        h, c0 = lay['h'], lay['c0']
        ca, cb = (c0 + shukei.col_of(t, 'num', k) - 1 for k in (0, 1))
        ci = c0 + shukei.col_of(t, 'item') - 1
        items, va, vb = [], [], []
        for r in range(h + 1, lay['last'] + 1):
            a, b = ws.Cells(r, ca).Value, ws.Cells(r, cb).Value
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                items.append(ws.Cells(r, ci).Value)
                va.append(a)
                vb.append(b)
        ha, hb = ws.Cells(h, ca).Value, ws.Cells(h, cb).Value
        co = ws.ChartObjects().Add(ws.Cells(h, c0 + lay['ncols'] + 1).Left, ws.Cells(h, 1).Top, 520, 300)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        horizontal = len(items) >= 9
        ch.ChartType = 57 if horizontal else 51
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        for name, vals, col in ((ha, va, GRAY), (hb, vb, BLUE)):
            s = ch.SeriesCollection().NewSeries()
            s.Name = name
            s.Values = vals
            s.XValues = items
            s.Format.Fill.ForeColor.RGB = rgb(col)
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{ha}と{hb}の比較"
        ch.HasLegend = True
        ch.Legend.Position = -4107
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = '#,##0'
        ax.HasMajorGridlines = False
        if horizontal:
            ch.Axes(1).ReversePlotOrder = True
        ch.ChartGroups(1).GapWidth = 60
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
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], 9).Left, ws.Cells(lay['h'], 9).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 4
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
            build_truth(xl, before, truth, t, lay)
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
    make([('未見1_標準', d(s(1), rng.randint(4, 12)), False),
          ('未見2_表題と合計行', d(s(2), 7, title=True, unit=True, total=True), True),
          ('未見3_並び違い', d(s(3), 9, swap=True, prev=True), False),
          ('未見4_空行と番号列', d(s(4), 6, blanks=1, no_col=True), False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('二本棒_本番', d(2901, 6), False),
          ('二本棒_試験A', d(2902, 12), False),
          ('二本棒_試験B', d(2903, 7, title=True, unit=True), False),
          ('二本棒_試験C', d(2904, 8, total=True), False),
          ('二本棒_試験D', d(2905, 8, blanks=2), False),
          ('二本棒_試験E', d(2906, 5, swap=True), False),
          ('二本棒_試験G', d(2907, 6, no_col=True, text=True), True),
          ('二本棒_試験H', d(2908, 7, prev=True), False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '二本棒_本番_正解.xlsx'), os.path.join(HERE, f'二本棒_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '二本棒_本番_選択.txt'), os.path.join(HERE, '二本棒_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
