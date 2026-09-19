# -*- coding: utf-8 -*-
"""表を棒と折れ線の複合グラフにする（どの表でも動く形・2026-09-18 作り直し。前の版は予算執行複合グラフ）。

列の決まり: 左端の文字の列が項目、数の列（率でない）が集合縦棒、率の列（表示形式に %）がマーカー付き折れ線で第 2 軸。
見出しの語・シート名・項目の名前は表ごとに違う（vocab.py）。
正解の決まり:
  - 見出しの行＝値の入ったセルが 2 つ以上ある最初の行。合計・計の行と空の行は入れない。
  - 系列は、数の列を表の左から順に、そのあと率の列（名前＝それぞれの列の見出し・項目＝左端の文字の列）。
  - 数値軸 #,##0、第 2 軸は 0%・最小 0・最大 1。タイトル: 表の上に表題があればその文字、無ければ「（項目の列の見出し）別の状況」。凡例は下。
  - 前に作ったグラフ（このシートのグラフ全部）は消して作り直す（置き場所は比べない）。
組: 本番 6 件・率は式／A 15 件／B 表題つき／C 合計の行／D 前のグラフ・率は値／E 率の列が 2 列目／F 撃った後にもう一度／
  G 数の列が 3 つ・備考の列／H 番号の列・途中に空行
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

REQUEST = ('表から複合グラフを作ってください。左端の文字の列を項目、率でない数の列は集合縦棒（数値軸 #,##0）、'
           '率の列（表示形式に %）はマーカー付き折れ線で第2軸（表示形式 0%・最小 0・最大 1）。'
           '折れ線の色は #C00000。見出しに 番号・No・コード・年度 を含む列と合計の行は入れない。前に作ったグラフは消して作り直す')
PHRASES = ['棒と折れ線の複合グラフにして', '複合グラフを作ってください', '率を第2軸にした複合グラフに',
           '金額は棒・率は折れ線のグラフを', '2軸のグラフにしてください', '棒に率の折れ線を重ねて']
TITLES = ['令和8年度 予算執行状況（9月末）', '{k}別の執行状況', '令和8年度上半期 {k}ごとの実績']


def build(seed, n, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    rows = []
    for k in v.items(n, rng):
        y = rng.randrange(100, 5000) * 1000
        rows.append((k, y, int(y * rng.uniform(0.3, 1.0) / 1000) * 1000))
    t = {'sheet': v.sheet('yosan'), 'ka': v.word('ka'), 'yosan': v.word('yosan'), 'shikko': v.word('shikko'),
         'rate': v.word('rate'), 'zan': '残額', 'biko': v.word('tekiyo'), 'no': v.word('no'), 'rows': rows,
         'title': None, 'total': False, 'formula': True, 'old_chart': False, 'rate_second': False,
         'three_num': False, 'note_col': False, 'no_col': False, 'blank_at': None, 'r0': 1}
    if kw.pop('title', False):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update(kw)
    return t


def heads_of(t):
    """見出しの並び（左から）と、それぞれの役。"""
    out = ([(t['no'], 'no')] if t['no_col'] else []) + [(t['ka'], 'item')]
    if t['rate_second']:
        out.append((t['rate'], 'rate'))
    out += [(t['yosan'], 'num'), (t['shikko'], 'num')]
    if t['three_num']:
        out.append((t['zan'], 'num'))
    if not t['rate_second']:
        out.append((t['rate'], 'rate'))
    if t['note_col']:
        out.append((t['biko'], 'text'))
    return out


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    h = t['r0']
    if t['title']:
        ws.cell(h, 1, t['title']).font = Font(bold=True, size=14)
        h += 2
    heads = heads_of(t)
    for j, (name, _role) in enumerate(heads, 1):
        ws.cell(h, j, name).font = Font(bold=True)
    col = {role_name: j for j, (role_name, _r) in enumerate(heads, 1)}
    L = lambda name: chr(64 + col[name])      # noqa: E731
    body = list(t['rows']) + ([('合計', sum(r[1] for r in t['rows']), sum(r[2] for r in t['rows']))] if t['total'] else [])
    r = h
    data_rows = []
    for i, (k, y, s) in enumerate(body):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        r += 1
        if t['no_col']:
            ws.cell(r, 1, i + 1)
        ws.cell(r, col[t['ka']], k)
        ws.cell(r, col[t['yosan']], y).number_format = '#,##0'
        ws.cell(r, col[t['shikko']], s).number_format = '#,##0'
        if t['three_num']:
            ws.cell(r, col[t['zan']], y - s).number_format = '#,##0'
        ws.cell(r, col[t['rate']], f"={L(t['shikko'])}{r}/{L(t['yosan'])}{r}" if t['formula'] else round(s / y, 4)
                ).number_format = '0.0%'
        if t['note_col']:
            ws.cell(r, col[t['biko']], random.Random(i).choice(['', '執行中', '完了']))
        if i < len(t['rows']):
            data_rows.append(r)
    wb.save(path)
    return {'h': h, 'rows': data_rows, 'col': col}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        h, col, rows = lay['h'], lay['col'], lay['rows']
        heads = heads_of(t)
        rng_of = lambda name: ws.Range(",".join(ws.Cells(r, col[name]).Address for r in rows))      # noqa: E731
        co = ws.ChartObjects().Add(ws.Cells(h, len(heads) + 2).Left, ws.Cells(h, 1).Top, 520, 300)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 51
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        nums = [nm for nm, role in heads if role == 'num']
        for name in nums + [t['rate']]:
            s = ch.SeriesCollection().NewSeries()
            s.Name = f"='{t['sheet']}'!{ws.Cells(h, col[name]).GetAddress(True, True, 1)}"
            s.Values = rng_of(name)
            s.XValues = rng_of(t['ka'])
        sr = ch.SeriesCollection(len(nums) + 1)
        sr.ChartType = 65
        sr.AxisGroup = 2
        sr.Format.Line.ForeColor.RGB = 0xC0                  # #C00000（既定の色は系列の数で変わるので決めておく）
        ch.Axes(2, 1).TickLabels.NumberFormat = '#,##0'
        ax2 = ch.Axes(2, 2)
        ax2.TickLabels.NumberFormat = '0%'
        ax2.MinimumScale = 0
        ax2.MaximumScale = 1
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or f"{t['ka']}別の状況"
        ch.HasLegend = True
        ch.Legend.Position = -4107
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def add_old_chart(xl, path, t, lay):
    wb = xl.Workbooks.Open(path)
    try:
        ws = wb.Worksheets(t['sheet'])
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], 8).Left, ws.Cells(lay['h'], 8).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 5
        s = co.Chart.SeriesCollection().NewSeries()
        s.Values = ws.Range(ws.Cells(lay['rows'][0], lay['col'][t['yosan']]), ws.Cells(lay['rows'][-1], lay['col'][t['yosan']]))
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
            print('  ', stem, t['sheet'], t['ka'], t['rate'], len(t['rows']))
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(4, 12))),
          ('未見2_表題と合計', build(s(2), 8, title=True, total=True, r0=2)),
          ('未見3_率が2列目と前のグラフ', build(s(3), 6, rate_second=True, old_chart=True, formula=False)),
          ('未見4_3つの数と空行', build(s(4), 7, three_num=True, blank_at=3, note_col=True))], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('複合_本番', build(3, 6)),
          ('複合_試験A', build(13, 15)),
          ('複合_試験B', build(23, 6, title=True, r0=2)),
          ('複合_試験C', build(33, 7, total=True)),
          ('複合_試験D', build(43, 5, old_chart=True, formula=False)),
          ('複合_試験E', build(53, 6, rate_second=True)),
          ('複合_試験G', build(63, 8, three_num=True, note_col=True)),
          ('複合_試験H', build(73, 6, no_col=True, blank_at=2))], HERE)
    shutil.copyfile(os.path.join(HERE, '複合_本番_正解.xlsx'), os.path.join(HERE, '複合_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '複合_本番_正解.xlsx'), os.path.join(HERE, '複合_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
