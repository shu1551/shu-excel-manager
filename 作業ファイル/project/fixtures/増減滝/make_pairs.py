# -*- coding: utf-8 -*-
"""選んでいる 2 列（前・後）から、合計の増減を項目ごとに説明するグラフ（ウォーターフォール）を積み上げ縦棒で作る
（E グラフ 44・2026-09-18 第二期）。

列の決まり: 選んでいる 2 列（1 つ目＝前・2 つ目＝後）。項目の列＝表の左端の文字の列。表は集計表（shukei）。
正解の決まり（依頼文に全部書く）:
  - 本文の行＝2 列とも数の行（空行と「合計」「計」「総計」の行は入れない）。表の上から順。
  - 項目（横軸）＝［1 つ目の見出し］＋本文の項目の名前＋［2 つ目の見出し］。
  - 前の合計 S0＝1 つ目の列の和。項目ごとの増減 d＝後−前。走る合計 S は S0 から d を足していく。
  - 積み上げ縦棒・系列は 4 本（この順・値は項目の数だけ・当てはまらない所は 0）:
      土台（塗りなし）＝［0］＋（d≥0 なら足す前の S、d<0 なら足した後の S）＋［0］
      合計（#7F7F7F）＝［S0］＋ 0 … ＋［後の列の和］／増加（#1F4E79）＝［0］＋ max(d,0) ＋［0］／減少（#C00000）＝［0］＋ max(−d,0) ＋［0］
  - タイトル: 表題があればその文字、無ければ「（1 つ目の見出し）から（2 つ目の見出し）への増減」14pt 太字。凡例なし・数値軸 #,##0・目盛線なし・
    棒の間隔 50・グラフの文字 Meiryo UI。前に作ったグラフは消して作り直す。
組: 本番／A 10 行／B 表題と単位／C 合計の行／D 途中の空行／E 数の列が左（並び違い）／F 撃った後にもう一度／G 番号の列と前のグラフ／H 前年度・今年度
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

REQUEST = ('選んでいる 2 列（1 つ目＝前、2 つ目＝後）から、合計の増減を項目ごとに説明するグラフ（ウォーターフォール）を、棒を重ねる形で作ってください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '項目の列＝表の左端の文字の列（数でも日付でもない列）。本文は見出しの次の行から最後の行まで。2 列とも数が入っている行だけを表の上から順に使い、'
           '空の行と、行のどこかのセル（ふつうは項目の列）に「合計」「計」「総計」とある行は入れない（合計の行も数が入っているので、数だけでは見分けられない）。'
           '項目（XValues）＝1 つ目の見出し、本文の項目の名前（上から順）、2 つ目の見出し。'
           'S0＝1 つ目の列の本文の和。項目ごとの増減 d＝2 つ目の値−1 つ目の値。走る合計 S は S0 から、上の行から順に d を足していく。'
           '種類は xlColumnStacked（棒を上に重ねる縦棒）。系列は次の 4 本をこの順に作り、値は項目の数だけ並べる（当てはまらない所は 0）: '
           '1「土台」＝先頭 0、各項目は d が 0 以上なら足す前の S・d が負なら足した後の S、末尾 0（塗りなし＝Format.Fill.Visible = msoFalse）／'
           '2「合計」＝先頭 S0、各項目 0、末尾は 2 つ目の列の本文の和（塗り #7F7F7F）／'
           '3「増加」＝先頭 0、各項目 d が正ならその値・それ以外 0、末尾 0（塗り #1F4E79）／'
           '4「減少」＝先頭 0、各項目 d が負なら −d・それ以外 0、末尾 0（塗り #C00000）。'
           'タイトル: 表の上に表題があればその文字（「単位」の行は表題にしない）、無ければ「（1 つ目の見出し）から（2 つ目の見出し）への増減」。14pt・太字。'
           '凡例なし。数値軸（Axes(xlValue)）の表示形式 #,##0・目盛線なし。棒の間隔（ChartGroups(1).GapWidth）50。データラベルなし。'
           'グラフ全体の文字は Meiryo UI（全体に当ててからタイトルの大きさ・太字を当てる）。前に作ったグラフ（このシートのグラフ全部）は消して作り直す。置き場所は表の右。')
PHRASES = ['増減の内訳をウォーターフォールで', 'ウォーターフォール図を作って', '前年度から今年度への増減を滝グラフで',
           '増減の要因を積み上げで説明するグラフ', '滝グラフにしてください', '合計の増減を項目別に階段状のグラフで']
GRAY, BLUE, RED = (0x7F, 0x7F, 0x7F), (0x1F, 0x4E, 0x79), (0xC0, 0x00, 0x00)


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def series_of(items, va, vb, ha, hb):
    """4 本の系列の値（純 Python）。"""
    s0 = sum(va)
    s = s0
    base, inc, dec = [0], [0], [0]
    for a, b in zip(va, vb):
        d = b - a
        base.append(s if d >= 0 else s + d)
        inc.append(d if d > 0 else 0)
        dec.append(-d if d < 0 else 0)
        s += d
    base.append(0)
    inc.append(0)
    dec.append(0)
    total = [s0] + [0] * len(items) + [sum(vb)]
    return [ha] + list(items) + [hb], base, total, inc, dec


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
        cats, base, total, inc, dec = series_of(items, va, vb, ha, hb)
        co = ws.ChartObjects().Add(ws.Cells(h, c0 + lay['ncols'] + 1).Left, ws.Cells(h, 1).Top, 560, 320)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 52
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        for name, vals, col in (('土台', base, None), ('合計', total, GRAY), ('増加', inc, BLUE), ('減少', dec, RED)):
            s = ch.SeriesCollection().NewSeries()
            s.Name = name
            s.Values = vals
            s.XValues = cats
            if col is None:
                s.Format.Fill.Visible = 0
            else:
                s.Format.Fill.ForeColor.RGB = rgb(col)
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{ha}から{hb}への増減"
        ch.HasLegend = False
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = '#,##0'
        ax.HasMajorGridlines = False
        ch.ChartGroups(1).GapWidth = 50
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
    make([('未見1_標準', d(s(1), rng.randint(4, 9), prev=True), False),
          ('未見2_表題と合計行', d(s(2), 6, title=True, unit=True, total=True), True),
          ('未見3_並び違い', d(s(3), 5, swap=True), False),
          ('未見4_空行と番号列', d(s(4), 6, blanks=1, no_col=True, prev=True), False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('滝_本番', d(3001, 5, prev=True), False),
          ('滝_試験A', d(3002, 10), False),
          ('滝_試験B', d(3003, 6, title=True, unit=True, prev=True), False),
          ('滝_試験C', d(3004, 6, total=True), False),
          ('滝_試験D', d(3005, 7, blanks=2, prev=True), False),
          ('滝_試験E', d(3006, 5, swap=True), False),
          ('滝_試験G', d(3007, 5, no_col=True, text=True), True),
          ('滝_試験H', d(3008, 4, prev=True), False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '滝_本番_正解.xlsx'), os.path.join(HERE, f'滝_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '滝_本番_選択.txt'), os.path.join(HERE, '滝_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
