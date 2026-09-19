# -*- coding: utf-8 -*-
"""評価の表をレーダーチャートにする（E グラフ 42・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い（表全体）。表＝左端の文字の列（課・施設など）＋右に指標の数の列 4〜6 本。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝空でないセルが 2 つ以上並ぶ最初の行。項目の列＝左端の文字の列。
  - 指標＝項目の列より右の数の列で、見出しが「合計」「計」「平均」でないもの（左の番号・コードの列、右の文字の列は入れない）。
  - 系列＝本文の行（表の上から順・6 本まで）。合計・計・平均・総計の行と空の行は入れない。系列の名前＝項目の列の値、項目＝指標の見出し。
  - マーカー付きレーダー（xlRadarMarkers）。線の色は順に #1F4E79・#C00000・#548235・#BF8F00・#7030A0・#7F7F7F、太さ 2.25、マーカーの大きさ 6。
  - タイトル: 表の上に表題があればその文字、無ければ「（項目の列の見出し）別の比較」14pt 太字。凡例は下。軸は既定のまま。
    グラフの文字 Meiryo UI。前に作ったグラフ（このシートのグラフ全部）は消して作り直す。
組: 本番 4 行 5 指標／A 8 行（6 本まで）／B 表題／C 平均の列と平均の行／D 前のグラフ・コードの列／E 3 行 6 指標／F 撃った後にもう一度／
  G 途中の空行・備考の列／H 合計の列と行
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

REQUEST = ('評価の表をレーダーチャートにしてください。'
           '見出しの行＝空でないセルが 2 つ以上並ぶ最初の行（文字でも数でもよい）。項目の列＝左端の文字の列。'
           '指標＝項目の列より右にある数の列で、見出しが「合計」「計」「平均」でないもの（項目の列より左の番号・コードの列と、右の文字の列は入れない）。'
           '系列は本文の行ごと（表の上から順・6 本まで）。合計・計・平均・総計の行と、行がまるごと空の行は入れない。'
           '系列の名前＝その行の項目の列の値、項目（XValues）＝指標の見出し、値＝その行の指標の数。'
           '種類はマーカー付きレーダー（xlRadarMarkers）。線の色は系列の順に #1F4E79・#C00000・#548235・#BF8F00・#7030A0・#7F7F7F、'
           '線の太さ 2.25、マーカーの大きさ 6。タイトル: 表の上に表題があればその文字、無ければ「（項目の列の見出し）別の比較」。14pt・太字。'
           '凡例は下。軸は既定のまま。グラフ全体の文字は Meiryo UI（全体に当ててからタイトルの大きさ・太字を当てる）。'
           '前に作ったグラフ（このシートのグラフ全部）は消して作り直す。置き場所は表の右。')
PHRASES = ['レーダーチャートにして', '評価をレーダーチャートで比べて', 'クモの巣グラフを作って',
           'この表をレーダーチャートに', '指標ごとのバランスをレーダーで見せて', '各課の評価をレーダーグラフに']
METRICS = [['正確さ', '速さ', '丁寧さ', '協調性', '積極性', '企画力'],
           ['執行率', '達成率', '満足度', '充足率', '改善率', '参加率'],
           ['安全', '環境', '福祉', '教育', '産業', '防災']]
TITLES = ['令和8年度 {k}別の評価', '{k}ごとの達成状況（9月末）']
COLORS = [(0x1F, 0x4E, 0x79), (0xC0, 0x00, 0x00), (0x54, 0x82, 0x35), (0xBF, 0x8F, 0x00), (0x70, 0x30, 0xA0), (0x7F, 0x7F, 0x7F)]


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def build(seed, n, m=5, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    t = {'sheet': rng.choice(['評価', '比較', 'Sheet1', '指標', '達成状況']), 'ka': v.word('ka'),
         'metrics': rng.choice(METRICS)[:m], 'items': v.items(n, rng), 'title': None, 'avg_col': False, 'avg_row': False,
         'sum_col': False, 'sum_row': False, 'code_col': False, 'note_col': False, 'blank_at': None, 'old_chart': False,
         'code': v.word('code'), 'note': v.word('tekiyo')}
    t['vals'] = [[rng.randint(40, 100) for _ in t['metrics']] for _ in t['items']]
    if kw.pop('title', False):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update(kw)
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    h = 3 if t['title'] else 1
    if t['title']:
        ws.cell(1, 1, t['title']).font = Font(bold=True, size=14)
    heads = ([t['code']] if t['code_col'] else []) + [t['ka']] + list(t['metrics']) \
        + (['平均'] if t['avg_col'] else []) + (['合計'] if t['sum_col'] else []) + ([t['note']] if t['note_col'] else [])
    for j, nm in enumerate(heads, 1):
        ws.cell(h, j, nm).font = Font(bold=True)
    ic = 2 if t['code_col'] else 1
    r = h
    rows = []
    for i, (it, vals) in enumerate(zip(t['items'], t['vals'])):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        r += 1
        if t['code_col']:
            ws.cell(r, 1, 100 + i * 10)
        ws.cell(r, ic, it)
        for j, x in enumerate(vals, 1):
            ws.cell(r, ic + j, x)
        c = ic + len(vals) + 1
        if t['avg_col']:
            ws.cell(r, c, round(sum(vals) / len(vals), 1))
            c += 1
        if t['sum_col']:
            ws.cell(r, c, sum(vals))
            c += 1
        if t['note_col']:
            ws.cell(r, c, random.Random(i).choice(['', '重点', '継続']))
        rows.append(r)
    for lab, on in (('平均', t['avg_row']), ('合計', t['sum_row'])):
        if on:
            r += 1
            ws.cell(r, ic, lab).font = Font(bold=True)
            for j in range(len(t['metrics'])):
                col = [vals[j] for vals in t['vals']]
                ws.cell(r, ic + 1 + j, round(sum(col) / len(col), 1) if lab == '平均' else sum(col))
    wb.save(path)
    return {'h': h, 'ic': ic, 'rows': rows, 'ncols': len(heads)}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], lay['ncols'] + 2).Left, ws.Cells(lay['h'], 1).Top, 420, 320)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 81
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        for k, (it, vals) in enumerate(list(zip(t['items'], t['vals']))[:6]):
            s = ch.SeriesCollection().NewSeries()
            s.Name = it
            s.Values = vals
            s.XValues = list(t['metrics'])
            s.Format.Line.ForeColor.RGB = rgb(COLORS[k])
            s.Format.Line.Weight = 2.25
            s.MarkerSize = 6
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{t['ka']}別の比較"
        ch.HasLegend = True
        ch.Legend.Position = -4107
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
        for stem, t in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = write_before(before, t)
            if t['old_chart']:
                old_chart(xl, before, t, lay)
            build_truth(xl, before, truth, t, lay)
            print('  ', stem, t['sheet'], t['ka'], t['metrics'], len(t['items']))
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(3, 5), m=rng.randint(4, 6))),
          ('未見2_表題と平均', build(s(2), 4, title=True, avg_col=True, avg_row=True)),
          ('未見3_コードと前のグラフ', build(s(3), 5, code_col=True, old_chart=True)),
          ('未見4_空行と合計', build(s(4), 4, blank_at=2, sum_col=True, sum_row=True, note_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('レーダー_本番', build(2801, 4)),
          ('レーダー_試験A', build(2802, 8)),
          ('レーダー_試験B', build(2803, 4, title=True)),
          ('レーダー_試験C', build(2804, 5, avg_col=True, avg_row=True)),
          ('レーダー_試験D', build(2805, 4, code_col=True, old_chart=True)),
          ('レーダー_試験E', build(2806, 3, m=6)),
          ('レーダー_試験G', build(2807, 5, blank_at=2, note_col=True)),
          ('レーダー_試験H', build(2808, 4, m=4, sum_col=True, sum_row=True))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, 'レーダー_本番_正解.xlsx'), os.path.join(HERE, f'レーダー_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
