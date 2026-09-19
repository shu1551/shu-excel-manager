# -*- coding: utf-8 -*-
"""月次集計の表を「いつもの見やすい推移グラフ」にする（鍛える回路の題材・グラフのデザイン 2 本目・2026-09-18）。

依頼は短い（見せ方の決まりは書かない）。正解のグラフが課で決めている見せ方の型を持つ＝AI は正解との違いから型を学ぶ。
見せ方の型（正解の作り方）:
  - マーカー付き折れ線。系列は課ごと。年間の合計の大きい課から 5 つまで（5 課以下なら全部）、大きい順に並べる。
  - 合計の行と列は入れない。項目は 4月〜3月。
  - 線の色は順に #1F4E79・#C00000・#548235・#BF8F00・#7030A0、線の太さ 2.25、マーカーの大きさ 6。
  - タイトル: 表の上に表題があればその文字、無ければ「課別月別支出の推移」。14pt・太字。
  - 凡例は右・数値軸 #,##0・目盛線あり・グラフの文字 Meiryo UI。前に作ったグラフは消して作り直す（置き場所は比べない）。
組: 本番 4 課／A 8 課（上位 5）／B 表題つき 5 課／C 合計の行と列が無い 7 課／D 前のグラフ／E 3 課／F 撃った後にもう一度／G 表題つき 12 課
"""
import importlib.util
import os
import random
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('mp_trend', os.path.join(HERE, '..', 'グラフ推移', 'make_pairs.py'))
R = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(R)
SHEET = R.SHEET
REQUEST = '月次集計の表を、いつもの見やすい推移グラフにしてください。前に作ったグラフが残っていたら消して作り直す'
COLORS = [(0x1F, 0x4E, 0x79), (0xC0, 0x00, 0x00), (0x54, 0x82, 0x35), (0xBF, 0x8F, 0x00), (0x70, 0x30, 0xA0)]
TITLE_TEXT = '令和8年度 月次集計（一般会計）'


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def build(seed, n):
    rng = random.Random(seed)
    kas = rng.sample(R.KA, n)
    return [(k, [rng.randrange(0, 900) * 1000 for _ in R.MONTHS]) for k in kas]      # 合計が同じ課は作らない（ほぼ起きない）


def build_truth(xl, before, truth, lay, rows, title):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Sheets(SHEET)
        h, c0 = lay['h'], lay['c0']
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        order = sorted(range(len(rows)), key=lambda i: -sum(rows[i][1]))[:5]
        co = ws.ChartObjects().Add(ws.Cells(lay['last'] + 2, c0).Left, ws.Cells(lay['last'] + 2, c0).Top, 640, 320)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 65
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        cats = ws.Range(ws.Cells(h, c0 + 1), ws.Cells(h, c0 + 12))
        for k, i in enumerate(order):
            s = ch.SeriesCollection().NewSeries()
            s.Name = f"='{SHEET}'!{ws.Cells(h + 1 + i, c0).GetAddress(True, True, 1)}"
            s.Values = ws.Range(ws.Cells(h + 1 + i, c0 + 1), ws.Cells(h + 1 + i, c0 + 12))
            s.XValues = cats
            s.Format.Line.ForeColor.RGB = rgb(COLORS[k])
            s.Format.Line.Weight = 2.25
            s.MarkerSize = 6
        ch.HasTitle = True
        ch.ChartTitle.Text = title or '課別月別支出の推移'
        ch.HasLegend = True
        ch.Legend.Position = -4152
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = '#,##0'
        ax.HasMajorGridlines = True
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = 'Meiryo UI'
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, rows, kw in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            old = kw.pop('old_chart', False)
            lay = R.write_before(before, rows, **kw)
            if old:
                R.add_old_chart(xl, before, lay)
            build_truth(xl, before, truth, lay, rows, TITLE_TEXT if kw.get('title') else None)
            print('  ', stem)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(2, 5)), {}),
          ('未見2_表題と多い課', build(s(2), rng.randint(6, 12)), {'title': True}),
          ('未見3_合計なしと前のグラフ', build(s(3), 6), {'totals': False, 'old_chart': True}),
          ('未見4_ちょうど5課', build(s(4), 5), {})], out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    make([('推移見やすい_本番', build(3, 4), {}),
          ('推移見やすい_試験A', build(13, 8), {}),
          ('推移見やすい_試験B', build(23, 5), {'title': True}),
          ('推移見やすい_試験C', build(33, 7), {'totals': False}),
          ('推移見やすい_試験D', build(43, 4), {'old_chart': True}),
          ('推移見やすい_試験E', build(53, 3), {}),
          ('推移見やすい_試験G', build(63, 12), {'title': True})], HERE)
    shutil.copyfile(os.path.join(HERE, '推移見やすい_本番_正解.xlsx'), os.path.join(HERE, '推移見やすい_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '推移見やすい_本番_正解.xlsx'), os.path.join(HERE, '推移見やすい_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    lines = ['月次集計をいつもの推移グラフに', '見やすい推移グラフを月次集計から', '月次集計の推移をいつもの見せ方の折れ線で',
             'いつもの形で月別推移のグラフを', '月次集計を見栄えのいい推移グラフに', '月別の推移を見やすいグラフでお願い']
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('ok')
