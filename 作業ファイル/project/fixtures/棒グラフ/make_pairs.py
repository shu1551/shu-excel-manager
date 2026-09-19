# -*- coding: utf-8 -*-
"""表を「いつもの見やすい棒グラフ」にする（どの表でも動く形・2026-09-18 作り直し）。
課別棒グラフ（デザインを依頼に書いた版）と見やすい棒グラフ（型を差分から学んだ版）を 1 本にまとめる。

列の決まり: 左端の文字の列が項目、右端の数の列が値。見出しの語・シート名・項目の名前は表ごとに違う（vocab.py）。
見せ方の型（正解の作り方）:
  - 見出しの行＝値の入ったセルが 2 つ以上ある最初の行。本文は見出しの次の行から表の最後まで。合計・計の行と空の行は入れない。
  - 値の大きい順に並べる（表の並びは変えない・グラフの中だけ）。項目が 8 以下なら集合縦棒、9 以上なら集合横棒で大きい順を上から（項目軸を反転）。
  - 系列の名前＝値の列の見出し。タイトル: 表の上に表題があればその文字（「（単位：…）」の行は表題にしない）、無ければ値の列の見出し（改行は詰める）。14pt・太字。
  - 凡例なし・棒に値のデータラベル・数値軸 #,##0・目盛線なし・棒の間隔 50・グラフの文字 Meiryo UI。
  - 棒の色 #1F4E79、いちばん大きい棒だけ #C00000。
  - 前に作ったグラフ（このシートのグラフ全部）は消して作り直す。置き場所は表の右（位置は比べない）。
組: 本番 5 件／A 7 件・予算の列が左に／B 12 件（横棒）・備考の列が右に／C 表題 6 件／D 表題と単位の行 10 件（横棒）／
  E 合計行 9 件／F 前のグラフ・番号の列／G 途中に空行 8 件・合計行／H 撃った後にもう一度／I 見出しの改行と単位
"""
import os
import random
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
from vocab import Vocab   # noqa: E402

REQUEST = ('表を、いつもの見やすい棒グラフにしてください。左端の文字の列を項目、右端の数の列を値にする。'
           '前に作ったグラフが残っていたら消して作り直す')
PHRASES = ['この表を見やすい棒グラフにして', 'いつもの型で棒グラフを作って', '見栄えのいい棒グラフにしてください',
           'いつもの見せ方で棒グラフに', '見やすい棒グラフをお願い', '表をいつもの棒グラフの形に']
BLUE, RED = (0x1F, 0x4E, 0x79), (0xC0, 0x00, 0x00)
TITLES = ['令和8年度 {k}別の支出（9月末）', '令和8年度上半期 {k}別の状況', '{k}ごとの実績']


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def build(seed, n, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    items = v.items(n, rng)
    amounts = [a * 1000 for a in rng.sample(range(50, 5000), n)]       # 同じ額を作らない（大きい順が一つに決まる）
    t = {'sheet': v.sheet('shukei'), 'ka': v.word('ka'), 'amount': kw.get('head') or v.word(rng.choice(['amount', 'shikko'])),
         'rows': list(zip(items, amounts)), 'title': None, 'unit': False, 'total': False, 'blank_at': None,
         'left_num': False, 'right_text': False, 'no_col': False, 'old_chart': False}
    if kw.get('title'):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update({k: val for k, val in kw.items() if k not in ('title', 'head')})
    t['yosan'] = v.word('yosan')
    t['biko'] = v.word('tekiyo')
    t['no'] = v.word('no')
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    h = 1
    if t['title']:
        ws.cell(1, 1, t['title']).font = Font(bold=True, size=14)
        h = 3
    if t['unit']:
        ws.cell(h - 1 if t['title'] else 1, 4, '（単位：円）')
        h = max(h, 2)
    heads = ([t['no']] if t['no_col'] else []) + [t['ka']] + ([t['yosan']] if t['left_num'] else []) + [t['amount']] \
        + ([t['biko']] if t['right_text'] else [])
    for j, name in enumerate(heads, 1):
        c = ws.cell(h, j, name)
        c.font = Font(bold=True)
        if '\n' in name:
            c.alignment = Alignment(wrap_text=True)
    item_c = heads.index(t['ka']) + 1
    val_c = heads.index(t['amount']) + 1
    r = h + 1
    data_rows = []
    for i, (k, a) in enumerate(t['rows']):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        if t['no_col']:
            ws.cell(r, 1, i + 1)
        ws.cell(r, item_c, k)
        if t['left_num']:
            ws.cell(r, item_c + 1, int(a * 1.3)).number_format = '#,##0'
        ws.cell(r, val_c, a).number_format = '#,##0'
        if t['right_text']:
            ws.cell(r, val_c + 1, random.Random(i).choice(['', '継続', '新規']))
        data_rows.append(r)
        r += 1
    if t['total']:
        ws.cell(r, item_c, '合計').font = Font(bold=True)
        ws.cell(r, val_c, sum(a for _k, a in t['rows'])).number_format = '#,##0'
    wb.save(path)
    return {'h': h, 'val_c': val_c}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        for i in range(int(ws.ChartObjects().Count), 0, -1):
            ws.ChartObjects(i).Delete()
        data = sorted(t['rows'], key=lambda x: -x[1])
        horizontal = len(data) >= 9
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], 8).Left, ws.Cells(lay['h'], 8).Top, 520, 300)
        co.Name = '表の整理_グラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        ch = co.Chart
        ch.ChartType = 57 if horizontal else 51
        for i in range(int(ch.SeriesCollection().Count), 0, -1):
            ch.SeriesCollection(i).Delete()
        s = ch.SeriesCollection().NewSeries()
        s.Name = f"='{t['sheet']}'!{ws.Cells(lay['h'], lay['val_c']).Address}"
        s.Values = [a for _k, a in data]
        s.XValues = [k for k, _a in data]
        ch.HasTitle = True
        ch.ChartTitle.Text = t['title'] or t['amount'].replace('\n', '')
        ch.HasLegend = False
        s.HasDataLabels = True
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = '#,##0'
        ax.HasMajorGridlines = False
        if horizontal:
            ch.Axes(1).ReversePlotOrder = True
        ch.ChartGroups(1).GapWidth = 50
        s.Format.Fill.ForeColor.RGB = rgb(BLUE)
        s.Points(1).Format.Fill.ForeColor.RGB = rgb(RED)
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = 'Meiryo UI'
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14      # 文字の書式を全体に当てた後で
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def add_old_chart(xl, path, t, lay):
    wb = xl.Workbooks.Open(path)
    try:
        ws = wb.Worksheets(t['sheet'])
        co = ws.ChartObjects().Add(ws.Cells(lay['h'], 8).Left, ws.Cells(lay['h'], 8).Top, 300, 200)
        co.Name = '表の整理_前のグラフ'      # マクロは「表の整理_」で始まる名前のグラフ（自分が作った物）だけを消す
        co.Chart.ChartType = 4
        s = co.Chart.SeriesCollection().NewSeries()
        s.Values = ws.Range(ws.Cells(lay['h'] + 1, lay['val_c']), ws.Cells(lay['h'] + len(t['rows']), lay['val_c']))
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
    make([('未見1_縦棒', build(s(1), rng.randint(3, 8), left_num=True)),
          ('未見2_横棒と表題', build(s(2), rng.randint(9, 16), title=True, unit=True, right_text=True)),
          ('未見3_合計行と前のグラフ', build(s(3), 8, total=True, old_chart=True, no_col=True)),
          ('未見4_ちょうど9件と空行', build(s(4), 9, blank_at=4)),
          ('未見5_見出しの改行', build(s(5), 6, head='支出額\n（円）', total=True))], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('棒_本番', build(3, 5)),
          ('棒_試験A', build(13, 7, left_num=True)),
          ('棒_試験B', build(23, 12, right_text=True)),
          ('棒_試験C', build(33, 6, title=True)),
          ('棒_試験D', build(43, 10, title=True, unit=True)),
          ('棒_試験E', build(53, 9, total=True)),
          ('棒_試験F', build(63, 6, old_chart=True, no_col=True)),
          ('棒_試験G', build(73, 8, total=True, blank_at=3)),
          ('棒_試験I', build(83, 7, head='金額\n（千円）'))], HERE)
    shutil.copyfile(os.path.join(HERE, '棒_本番_正解.xlsx'), os.path.join(HERE, '棒_試験H_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '棒_本番_正解.xlsx'), os.path.join(HERE, '棒_試験H_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
