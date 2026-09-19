# -*- coding: utf-8 -*-
"""選んでいる列でテーブルを並べ替えて絞り込む（どの表でも動く形・2026-09-18 作り直し。前の版は職員名簿の絞り込み）。

列の決まり: 選んでいる 2〜3 列（1 つ目＝昇順の並べ替え、2 つ目＝降順の並べ替え、3 つ目があれば その列が 0 の行を除く）。
選んでいる列は <名前>_選択.txt に書く。見出しの語・シート名・テーブル名・列の並びは表ごとに違う（vocab.py）。
正解の決まり:
  - 前の絞り込みは解除してから（ほかの列の条件は残さない）。テーブルの名前・範囲・列は変えない。
  - 並べ替えは 1 つ目の列で昇順、同じ値は 2 つ目の列で降順。3 つ目の列があれば その列が 0 の行を隠す。
組: 本番 2 列＋0 を除く／A 並べ替えだけ（2 列）／B 前の絞り込みが別の列に付いている／C 前の並べ替えが違う／
  D 集計行つき／E 大量 500 行／F 撃った後にもう一度／G 番号の列・備考の列
"""
import datetime
import os
import random
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
from vocab import Vocab   # noqa: E402

SEI = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤']
MEI = ['一郎', '花子', '誠', '恵', '学', '悟', '由美', '健太', '真理', '修']
REQUEST = ('選んでいる列で、アクティブなシートのテーブルを並べ替えて絞り込んでください。'
           '1 つ目の列で昇順、同じ値は 2 つ目の列で降順に並べ替える。'
           '3 つ目の列が選ばれていれば、その列が 0 の行を除いて表示（絞り込み）。'
           '前の絞り込みは解除してから（ほかの列の条件は残さない）。テーブルの名前・範囲・列は変えない')
PHRASES = ['選んだ列で並べ替えて0を除いて', 'テーブルを並べ替えて絞り込んで', '昇順と降順で並べ替えて0の行は隠して',
           '並べ替えと絞り込みで整えて', 'テーブルを整えて（並べ替えと絞り込み）', '選んでいる列で並べ替え']


def build(seed, n, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    items = v.items(rng.randint(3, 6), rng)
    rows = []
    for i in range(n):
        rows.append({'item': rng.choice(items), 'name': rng.choice(SEI) + ' ' + rng.choice(MEI),
                     'date': datetime.datetime(rng.randint(1995, 2025), rng.randint(1, 12), 1),
                     'kihon': rng.randrange(180, 450) * 1000,
                     'teate': 0 if rng.random() < 0.25 else rng.randrange(1, 90) * 1000,
                     'memo': rng.choice(['', '再任用', '育休'])})
    t = {'rows': rows, 'sheet': v.sheet('meibo'), 'table': v.table('meibo'), 'item_head': v.word('ka'),
         'name_head': v.word('name'), 'date_head': '採用日', 'kihon_head': v.word('kihon'),
         'teate_head': v.word('teate'), 'memo_head': v.word('tekiyo'), 'no_head': v.word('no'),
         'sort_only': False, 'old_filter': False, 'old_sort': False, 'totals': False, 'no_col': False,
         'memo_col': False}
    t.update(kw)
    return t


def cols_of(t):
    out = ([(t['no_head'], 'no')] if t['no_col'] else []) + [
        (t['item_head'], 'item'), (t['name_head'], 'name'), (t['date_head'], 'date'),
        (t['kihon_head'], 'kihon'), (t['teate_head'], 'teate')]
    if t['memo_col']:
        out.append((t['memo_head'], 'memo'))
    return out


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    cols = cols_of(t)
    for j, (name, _role) in enumerate(cols, 1):
        ws.cell(1, j, name).font = Font(bold=True)
    for i, row in enumerate(t['rows'], 2):
        for j, (_name, role) in enumerate(cols, 1):
            v = (i - 1) if role == 'no' else row[role]
            c = ws.cell(i, j, v)
            if role == 'date':
                c.number_format = 'yyyy/m/d'
            elif role in ('kihon', 'teate'):
                c.number_format = '#,##0'
    tb = Table(displayName=t['table'], ref=f"A1:{get_column_letter(len(cols))}{len(t['rows']) + 1}")
    tb.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(tb)
    wb.save(path)
    return {'cols': cols}


def apply_rules(wb, t, truth=True):
    ws = wb.Worksheets(t['sheet'])
    lo = ws.ListObjects(t['table'])
    cols = cols_of(t)
    idx = {role: j for j, (_n, role) in enumerate(cols, 1)}
    with_sort = lo.Sort
    with_sort.SortFields.Clear()
    if truth:
        lo.AutoFilter.ShowAllData() if lo.AutoFilter is not None else None
        with_sort.SortFields.Add2(lo.ListColumns(idx['item']).Range, 0, 1)      # 昇順
        with_sort.SortFields.Add2(lo.ListColumns(idx['kihon']).Range, 0, 2)     # 降順
        with_sort.Header = 1
        with_sort.Apply()
        if not t['sort_only']:
            lo.Range.AutoFilter(Field=idx['teate'], Criteria1='<>0')
    else:
        if t['old_sort']:
            with_sort.SortFields.Add2(lo.ListColumns(idx['name']).Range, 0, 1)
            with_sort.Header = 1
            with_sort.Apply()
        if t['old_filter']:
            lo.Range.AutoFilter(Field=idx['item'], Criteria1='=' + str(t['rows'][0]['item']))
    if t['totals']:
        lo.ShowTotals = True
        for role in ('kihon', 'teate'):
            lo.ListColumns(idx[role]).TotalsCalculation = 1


def make(jobs, out=HERE):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            lay = write_before(before, t)
            wb = xl.Workbooks.Open(before)
            try:
                apply_rules(wb, t, truth=False)
                wb.Save()
            finally:
                wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                apply_rules(wb, t, truth=True)
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            idx = {role: j for j, (_n, role) in enumerate(lay['cols'], 1)}
            roles = ('item', 'kihon') if t['sort_only'] else ('item', 'kihon', 'teate')
            sel = ",".join(f"{get_column_letter(idx[r])}1" for r in roles)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], t['table'], [n for n, _r in lay['cols']], '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def kw(t, **over):
    t.update(over)
    return t


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(10, 60))),
          ('未見2_並べ替えだけ', build(s(2), 20, sort_only=True)),
          ('未見3_前の絞り込みと並べ替え', build(s(3), 25, old_filter=True, old_sort=True)),
          ('未見4_集計行と番号', build(s(4), 18, totals=True, no_col=True, memo_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('絞り込み_本番', build(3, 20)),
          ('絞り込み_試験A', build(13, 18, sort_only=True)),
          ('絞り込み_試験B', build(23, 22, old_filter=True)),
          ('絞り込み_試験C', build(33, 16, old_sort=True)),
          ('絞り込み_試験D', build(43, 20, totals=True)),
          ('絞り込み_試験E', build(53, 300)),
          ('絞り込み_試験G', build(63, 15, no_col=True, memo_col=True))])
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '絞り込み_本番_正解.xlsx'), os.path.join(HERE, f'絞り込み_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '絞り込み_本番_選択.txt'), os.path.join(HERE, '絞り込み_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
