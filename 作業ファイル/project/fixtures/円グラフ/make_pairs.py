# -*- coding: utf-8 -*-
"""表の構成比を円グラフにする（どの表でも動く形・2026-09-18 作り直し。前の版は性質別の構成比円）。

列の決まり: 左端の文字の列が項目、右端の数の列（表示形式が % の列は除く）が値。見出しの語・シート名・項目の名前は表ごとに違う。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝値の入ったセルが 2 つ以上ある最初の行。合計・計の行、空の行、値が 0 の項目は入れない。
  - 円グラフ・系列 1 本（名前＝値の列の見出し）。タイトル: 表の上に表題があればその文字、無ければ「（値の列の見出し）の構成比」
    （見出しの改行は詰める）。凡例は右・データラベルはパーセント（値は出さない）・グラフの文字 Meiryo UI。
  - 前に作ったグラフ（このシートのグラフ全部）は消して作り直す（置き場所は比べない）。
組: 本番 6 件・合計行／A 9 件・%の列が右に／B 0 が 2 件／C 表題 B3 から／D 合計行なし・予算の列が左に／E 前のグラフ・番号の列／
  F 撃った後にもう一度／G 途中に空行・備考の列
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

KUBUN = ['人件費', '物件費', '維持補修費', '扶助費', '補助費等', '普通建設事業費', '災害復旧事業費', '公債費', '積立金', '繰出金']
REQUEST = ('表の構成比を円グラフにしてください。左端の文字の列を項目、右端の数の列（% の列は除く）を値にする。'
           '合計の行と値が 0 の項目はグラフに入れない。前に作ったグラフが残っていたら消して作り直す')
PHRASES = ['この表の構成比を円グラフにして', '割合を円グラフで表示して', '構成比の円グラフを作ってください',
           '項目ごとの割合を円グラフに', '円グラフで構成比をお願いします', '内訳を円グラフにして']
TITLES = ['令和7年度 決算 性質別歳出', '令和8年度 {k}別の内訳', '{k}ごとの割合']


def build(seed, n, zeros=0, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    use_kubun = rng.random() < 0.3
    items = rng.sample(KUBUN, min(n, len(KUBUN))) if use_kubun else v.items(n, rng)
    rows = [(k, rng.randrange(100, 90000) * 1000) for k in items]
    for i in rng.sample(range(len(rows)), zeros):
        rows[i] = (rows[i][0], 0)
    t = {'sheet': v.sheet(rng.choice(['shukei', 'yosan'])), 'ka': '区分' if use_kubun else v.word('ka'),
         'amount': v.word(rng.choice(['amount', 'shikko'])), 'yosan': v.word('yosan'), 'no': v.word('no'),
         'biko': v.word('tekiyo'), 'rows': rows, 'title': None, 'r0': 1, 'c0': 1, 'total': True, 'rate_col': False,
         'left_num': False, 'no_col': False, 'blank_at': None, 'old_chart': False, 'note_col': False}
    if kw.pop('title', False):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update(kw)
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    r0, c0 = t['r0'], t['c0']
    h = r0
    if t['title']:
        ws.cell(r0, c0, t['title']).font = Font(bold=True, size=14)
        h = r0 + 2
    heads = ([t['no']] if t['no_col'] else []) + [t['ka']] + ([t['yosan']] if t['left_num'] else []) + [t['amount']] \
        + (['構成比'] if t['rate_col'] else []) + ([t['biko']] if t['note_col'] else [])
    for j, name in enumerate(heads):
        ws.cell(h, c0 + j, name).font = Font(bold=True)
    ic, vc = c0 + heads.index(t['ka']), c0 + heads.index(t['amount'])
    total = sum(a for _k, a in t['rows'])
    r = h + 1
    data = []
    for i, (k, a) in enumerate(t['rows']):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        if t['no_col']:
            ws.cell(r, c0, i + 1)
        ws.cell(r, ic, k)
        if t['left_num']:
            ws.cell(r, ic + 1, int(a * 1.2) + 1000).number_format = '#,##0'
        ws.cell(r, vc, a).number_format = '#,##0'
        if t['rate_col']:
            ws.cell(r, vc + 1, a / total).number_format = '0.0%'
        if t['note_col']:
            ws.cell(r, c0 + len(heads) - 1, random.Random(i).choice(['', '前年並み', '増']))
        data.append((r, a))
        r += 1
    if t['total']:
        ws.cell(r, ic, '合計').font = Font(bold=True)
        ws.cell(r, vc, total).number_format = '#,##0'
        if t['rate_col']:
            ws.cell(r, vc + 1, 1).number_format = '0.0%'
    wb.save(path)
    return {'h': h, 'ic': ic, 'vc': vc, 'rows': [rr for rr, a in data if a != 0], 'all': [rr for rr, _a in data]}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        h, ic, vc, rows = lay['h'], lay['ic'], lay['vc'], lay['rows']
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        co = ws.ChartObjects().Add(ws.Cells(h, vc + 3).Left, ws.Cells(h, vc + 3).Top, 420, 300)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 5
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        s = ch.SeriesCollection().NewSeries()
        s.Name = f"='{t['sheet']}'!{ws.Cells(h, vc).Address}"
        s.Values = ws.Range(",".join(ws.Cells(r, vc).Address for r in rows))
        s.XValues = ws.Range(",".join(ws.Cells(r, ic).Address for r in rows))
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{t['amount']}の構成比"
        ch.HasLegend = True
        ch.Legend.Position = -4152
        s.HasDataLabels = True
        s.DataLabels().ShowPercentage = True
        s.DataLabels().ShowValue = False
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = 'Meiryo UI'
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def add_old_chart(xl, path, t, lay):
    wb = xl.Workbooks.Open(path)
    try:
        ws = wb.Worksheets(t['sheet'])
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], lay['vc'] + 3).Left, ws.Cells(lay['h'], lay['vc'] + 3).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 51
        s = co.Chart.SeriesCollection().NewSeries()
        s.Values = ws.Range(ws.Cells(lay['all'][0], lay['vc']), ws.Cells(lay['all'][-1], lay['vc']))
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
            print('  ', stem, t['sheet'], t['ka'], t['amount'], len(t['rows']))
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(3, 10), rate_col=True)),
          ('未見2_0と表題', build(s(2), 8, zeros=3, title=True, r0=2, c0=2, no_col=True)),
          ('未見3_合計なしと前のグラフ', build(s(3), 7, zeros=1, total=False, old_chart=True, left_num=True)),
          ('未見4_空行と備考', build(s(4), 6, blank_at=2, note_col=True))], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('円_本番', build(3, 6)),
          ('円_試験A', build(13, 9, rate_col=True)),
          ('円_試験B', build(23, 8, zeros=2)),
          ('円_試験C', build(33, 6, title=True, r0=3, c0=2)),
          ('円_試験D', build(43, 7, total=False, left_num=True)),
          ('円_試験E', build(53, 6, old_chart=True, no_col=True)),
          ('円_試験G', build(63, 7, blank_at=3, note_col=True))], HERE)
    shutil.copyfile(os.path.join(HERE, '円_本番_正解.xlsx'), os.path.join(HERE, '円_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '円_本番_正解.xlsx'), os.path.join(HERE, '円_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
