# -*- coding: utf-8 -*-
"""表を積み上げ縦棒グラフにする（どの表でも動く形・2026-09-18 作り直し。前の版は課別科目別積み上げ）。

列の決まり: 左端の文字の列が項目（横軸）、その右の数の列（合計・計・総計の列は除く）が系列（見出しが系列の名前）。
系列の見出しは科目・費目・月など表ごとに違う。見出しの語・シート名・項目の名前も違う（vocab.py）。
正解の決まり:
  - 見出しの行＝値の入ったセルが 2 つ以上ある最初の行（ピボットを貼った表は「行ラベル」の行）。
    合計・計・総計の行と空の行は入れない。
  - 積み上げ縦棒（ChartType 52）。系列は左から順（名前＝その列の見出し・項目＝左端の文字の列）。
  - 数値軸 #,##0・凡例は右。タイトル: 表の上に表題があればその文字、無ければ「（項目の列の見出し）別の内訳」。
  - 前に作ったグラフ（このシートのグラフ全部）は消して作り直す（置き場所は比べない）。
組: 本番 5 件 x 4 系列・合計の行と列／A 8 件 x 3／B 表題つき・合計なし／C 系列が月／D 前のグラフ・番号の列／
  E ピボットを貼った表（行ラベル・総計）／F 撃った後にもう一度／G 途中に空行・備考の列
"""
import os
import random
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
from vocab import Vocab   # noqa: E402

REQUEST = ('表から積み上げ縦棒グラフを作ってください。見出しの行は、数が並ぶ本文のすぐ上の行（ピボットを貼り付けた表では'
           '「行ラベル」の行。その上の「合計 / …」「列ラベル」の行は見出しではない）。'
           '左端の文字の列を項目、その右の数の列（合計・計・総計の列は除く）を系列にする。'
           '合計・計・総計の行は入れない。数値軸は #,##0、凡例は右。'
           'タイトルは、表の上に表題があればその文字（「合計 / …」「列ラベル」「行ラベル」の行は表題ではない）、'
           '無ければ「（項目の列の見出し）別の内訳」。前に作ったグラフは消して作り直す')
PHRASES = ['積み上げ縦棒グラフにして', '積み上げ棒グラフを作って', '内訳を積み上げの棒グラフに',
           '積み上げグラフをお願いします', '積み上げの縦棒で見せて', '構成を積み上げ棒で']
TITLES = ['令和8年度 {k}別・費目別支出', '{k}別の内訳（9月末）', '令和8年度上半期 {k}ごとの内訳']


def build(seed, n, m=4, series='kamoku', **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    if series == 'month':
        heads = [f"{mm}月" for mm in (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)][:m]
    else:
        heads = v.kamokus(m, rng)
    items = v.items(n, rng)
    rows = [(k, [rng.randrange(10, 900) * 1000 for _ in heads]) for k in items]
    t = {'sheet': v.sheet(rng.choice(['shukei', 'getsuji'])), 'ka': v.word('ka'), 'heads': heads, 'rows': rows,
         'series': series, 'title': None, 'totals': True, 'pivot_style': False, 'old_chart': False,
         'no_col': False, 'note_col': False, 'blank_at': None, 'no': v.word('no'), 'biko': v.word('tekiyo'),
         'amount': v.word('amount')}
    if kw.pop('title', False):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update(kw)
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    h = 1
    if t['title']:
        ws.cell(1, 1, t['title']).font = Font(bold=True, size=14)
        h = 3
    if t['pivot_style']:
        ws.cell(h, 1, f"合計 / {t['amount']}")
        ws.cell(h, 2, '列ラベル')
        h += 1
    c = 1
    if t['no_col']:
        ws.cell(h, c, t['no']).font = Font(bold=True)
        c += 1
    ic = c
    ws.cell(h, ic, '行ラベル' if t['pivot_style'] else t['ka']).font = Font(bold=True)
    for j, hd in enumerate(t['heads'], 1):
        ws.cell(h, ic + j, hd).font = Font(bold=True)
    last_c = ic + len(t['heads'])
    if t['totals']:
        last_c += 1
        ws.cell(h, last_c, '総計' if t['pivot_style'] else '合計').font = Font(bold=True)
    if t['note_col']:
        ws.cell(h, last_c + 1, t['biko']).font = Font(bold=True)
    r = h
    data_rows = []
    for i, (k, vals) in enumerate(t['rows']):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        r += 1
        if t['no_col']:
            ws.cell(r, 1, i + 1)
        ws.cell(r, ic, k)
        for j, val in enumerate(vals, 1):
            ws.cell(r, ic + j, val).number_format = '#,##0'
        if t['totals']:
            ws.cell(r, ic + len(vals) + 1, sum(vals)).number_format = '#,##0'
        if t['note_col']:
            ws.cell(r, last_c + 1, random.Random(i).choice(['', '確定', '見込']))
        data_rows.append(r)
    if t['totals']:
        r += 1
        ws.cell(r, ic, '総計' if t['pivot_style'] else '合計').font = Font(bold=True)
        for j in range(len(t['heads']) + 1):
            ws.cell(r, ic + 1 + j, sum((vals + [sum(vals)])[j] for _k, vals in t['rows'])).number_format = '#,##0'
    wb.save(path)
    return {'h': h, 'ic': ic, 'rows': data_rows, 'last': r}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        h, ic, rows = lay['h'], lay['ic'], lay['rows']
        cats = ws.Range(",".join(ws.Cells(r, ic).Address for r in rows))
        co = ws.ChartObjects().Add(ws.Cells(h, ic + len(t['heads']) + 3).Left, ws.Cells(h, 1).Top, 560, 320)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 52
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        for j in range(1, len(t['heads']) + 1):
            s = ch.SeriesCollection().NewSeries()
            s.Name = f"='{t['sheet']}'!{ws.Cells(h, ic + j).GetAddress(True, True, 1)}"
            s.Values = ws.Range(",".join(ws.Cells(r, ic + j).Address for r in rows))
            s.XValues = cats
        ch.Axes(2).TickLabels.NumberFormat = '#,##0'
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{'行ラベル' if t['pivot_style'] else t['ka']}別の内訳"
        ch.HasLegend = True
        ch.Legend.Position = -4152
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def add_old_chart(xl, path, t, lay):
    wb = xl.Workbooks.Open(path)
    try:
        ws = wb.Worksheets(t['sheet'])
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], lay['ic'] + 8).Left, ws.Cells(lay['h'], 1).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 4
        s = co.Chart.SeriesCollection().NewSeries()
        s.Values = ws.Range(ws.Cells(lay['rows'][0], lay['ic'] + 1), ws.Cells(lay['rows'][-1], lay['ic'] + 1))
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = write_before(before, t)
            if t['old_chart']:
                add_old_chart(xl, before, t, lay)
            build_truth(xl, before, truth, t, lay)
            print('  ', stem, t['sheet'], t['ka'], t['heads'][:2], len(t['rows']))
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(3, 8), rng.randint(2, 5))),
          ('未見2_表題と月の系列', build(s(2), 6, 6, 'month', title=True)),
          ('未見3_合計なしと前のグラフ', build(s(3), 7, 3, totals=False, old_chart=True, no_col=True)),
          ('未見4_ピボット貼りと空行', build(s(4), 5, 4, pivot_style=True, blank_at=2))], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('積み上げ_本番', build(3, 5, 4)),
          ('積み上げ_試験A', build(13, 8, 3)),
          ('積み上げ_試験B', build(23, 6, 4, title=True, totals=False)),
          ('積み上げ_試験C', build(33, 5, 6, 'month')),
          ('積み上げ_試験D', build(43, 6, 3, old_chart=True, no_col=True)),
          ('積み上げ_試験E', build(53, 5, 4, pivot_style=True)),
          ('積み上げ_試験G', build(63, 7, 5, blank_at=3, note_col=True))], HERE)
    shutil.copyfile(os.path.join(HERE, '積み上げ_本番_正解.xlsx'), os.path.join(HERE, '積み上げ_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '積み上げ_本番_正解.xlsx'), os.path.join(HERE, '積み上げ_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
