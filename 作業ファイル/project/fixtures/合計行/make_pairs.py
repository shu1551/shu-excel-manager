# -*- coding: utf-8 -*-
"""合計の行を足す（どの表でも動く形・2026-09-18 作り直し）。

仕事: アクティブなシートの表の下に合計の行を足す。見出しの語・シート名・項目の名前は表ごとに違う（vocab.py）。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝値の入ったセルが 2 つ以上ある最初の行（上に表題の行があることもある）。本文は見出しの次の行から、
    表の最後の行まで（途中の空行 1 行は続き）。表の下の「※」で始まる注記の行は本文に入れない。
  - 合計の行は本文のすぐ下（すぐ下に注記の行があれば行を挿入して注記の上）。最後の行が既に合計の行（「合計」「計」）なら、その行に書き直す。
  - 「合計」は左端の文字の列に書く（左端が番号の列なら、その右の文字の列）。
  - 数の列は =SUM(本文の範囲)。番号の列（見出しに 番号・No・コード を含む列）・年度の列（見出しに 年度 を含む列）・
    率の列（表示形式に %）は足さない（空のまま）。文字の列・日付の列も空のまま。
組: 本番 予算執行（率の列）／A 番号の列／B 表題つき B3 から／C 前の合計行（計）が残る／D 途中に空行／E 表の下に※注記／
  F 日付の列・年度の列／G コードの列が数／H 撃った後にもう一度
  py make_pairs.py            → *_前.xlsx と *_正解.xlsx
  py make_pairs.py --unseen N → 未見_種N
"""
import os
import random
import shutil
import sys
import datetime

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
from vocab import Vocab   # noqa: E402

REQUEST = ('表の下に合計の行を足してください。「合計」は左端の文字の列に書き、数の列は =SUM の式で本文を合計する。'
           '見出しに 番号・No・コード・年度 を含む列（見出しのこの語で見分ける）と、率（表示形式が %）の列、文字や日付の列は足さない。'
           '見出しの行は値が 2 つ以上並ぶ最初の行、'
           '途中の空行も本文に含める。表のすぐ下に※の注記の行があれば、行を挿入して注記の上に合計の行を入れる。'
           '最後の行が既に合計（合計・計）の行ならその行を書き直す（行を増やさない）')
PHRASES = ['表のいちばん下に合計行を付けて', '数の列の合計を表の下に出して', '合計行を追加してください', '最終行の下に列ごとの合計を入れて',
           '下に合計の行をお願いします', '縦計を表の下に足して']


def build(seed, kind):
    """種と型 → 表の設計 {'sheet','title','r0','c0','heads':[(見出し, 型)], 'rows':[[値]], 'blank_at','note','total_label'}。"""
    rng = random.Random(seed)
    v = Vocab(seed)
    n = rng.randint(4, 9)
    items = v.items(n, rng)
    cols = []
    if kind in ('no', 'code'):
        cols.append((v.word('no') if kind == 'no' else v.word('code'), 'no'))
    cols.append((v.word('ka'), 'text'))
    if kind == 'date':
        cols.append((v.word('date'), 'date'))
        cols.append(('年度', 'year'))
    cols.append((v.word('yosan'), 'num'))
    cols.append((v.word('shikko'), 'num'))
    if kind in ('rate', 'plain'):
        cols.append((v.word('rate'), 'rate'))
    if kind in ('title', 'note'):
        cols.append((v.word('tekiyo'), 'text2'))
    rows = []
    for i, it in enumerate(items):
        y = rng.randrange(100, 9000) * 1000
        s = int(y * rng.uniform(0.2, 1.0) / 1000) * 1000
        row = []
        for _h, t in cols:
            row.append({'no': (i + 1) if kind == 'no' else 100 + (i + 1) * 10, 'text': it,
                        'date': datetime.date(2026, rng.randint(4, 12), rng.randint(1, 28)), 'year': 2026,
                        'num': None, 'rate': None, 'text2': rng.choice(['継続', '新規', '', '一部']) }[t])
        nums = [k for k, (_h, t) in enumerate(cols) if t == 'num']
        row[nums[0]], row[nums[1]] = y, s
        if kind in ('rate', 'plain'):
            row[[k for k, (_h, t) in enumerate(cols) if t == 'rate'][0]] = s / y
        rows.append(row)
    return {'sheet': v.sheet(rng.choice(['shukei', 'yosan'])), 'cols': cols, 'rows': rows,
            'title': f"令和8年度 {v.word('ka')}別の状況" if kind == 'title' else None,
            'r0': 3 if kind == 'title' else 1, 'c0': 2 if kind == 'title' else 1,
            'blank_at': rng.randint(1, n - 2) if kind == 'blank' else None,
            'note': '※金額は税込み' if kind == 'note' else None,
            'total_label': '計' if kind == 'total' else None}


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    r0, c0 = t['r0'], t['c0']
    if t['title']:
        ws.cell(1, c0, t['title']).font = Font(bold=True, size=14)
    for j, (h, _k) in enumerate(t['cols']):
        ws.cell(r0, c0 + j, h).font = Font(bold=True)
    r = r0 + 1
    first = r
    for i, row in enumerate(t['rows']):
        if t['blank_at'] is not None and i == t['blank_at']:
            r += 1
        for j, v in enumerate(row):
            c = ws.cell(r, c0 + j, v)
            kind = t['cols'][j][1]
            if kind == 'num':
                c.number_format = '#,##0'
            elif kind == 'rate':
                c.number_format = '0.0%'
            elif kind == 'date':
                c.number_format = 'yyyy/m/d'
        r += 1
    last = r - 1
    if t['total_label']:
        text_col = next(j for j, (_h, k) in enumerate(t['cols']) if k == 'text')
        ws.cell(r, c0 + text_col, t['total_label'])
        for j, (_h, k) in enumerate(t['cols']):
            if k == 'num':
                ws.cell(r, c0 + j, 12345)                    # 前の合計（古い値）
        r += 1
    if t['note']:
        ws.cell(r, c0, t['note'])                             # 表のすぐ下に注記（行を挿入しないと合計を入れられない）
    wb.save(path)
    return {'first': first, 'last': last, 'total_at': (last + 1) if t['total_label'] else None,
            'note_at': r if t['note'] else None}


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0 = t['c0']
        at = lay['total_at'] or (lay['last'] + 1)
        if lay['note_at'] and not lay['total_at']:
            ws.Rows(at).Insert()
        text_col = next(j for j, (_h, k) in enumerate(t['cols']) if k == 'text')
        ws.Cells(at, c0 + text_col).Value = '合計'
        for j, (_h, k) in enumerate(t['cols']):
            if k == 'num':
                a = ws.Cells(lay['first'], c0 + j).GetAddress(False, False)
                b = ws.Cells(lay['last'], c0 + j).GetAddress(False, False)
                ws.Cells(at, c0 + j).Formula = f"=SUM({a}:{b})"
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            lay = write_before(before, t)
            build_truth(xl, before, truth, t, lay)
            print('  ', stem, t['sheet'], [h for h, _k in t['cols']])
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    kinds = ['plain', 'no', 'title', 'total', 'blank', 'note', 'date', 'code', 'rate']
    rng = random.Random(seed)
    make([(f"未見{k + 1}_{kind}", build(seed * 100 + k, kind)) for k, kind in enumerate(rng.sample(kinds, 6))], out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('合計行_本番', build(1, 'rate')),
          ('合計行_試験A', build(2, 'no')),
          ('合計行_試験B', build(3, 'title')),
          ('合計行_試験C', build(4, 'total')),
          ('合計行_試験D', build(5, 'blank')),
          ('合計行_試験E', build(6, 'note')),
          ('合計行_試験F', build(7, 'date')),
          ('合計行_試験G', build(8, 'code'))], HERE)
    shutil.copyfile(os.path.join(HERE, '合計行_本番_正解.xlsx'), os.path.join(HERE, '合計行_試験H_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '合計行_本番_正解.xlsx'), os.path.join(HERE, '合計行_試験H_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
