# -*- coding: utf-8 -*-
"""別のテーブルから引く計算列を足す（どの表でも動く形・2026-09-18 作り直し。前の版は課名を引く）。

形で決める: アクティブなシートのテーブルと、別のテーブル（共通の見出しの列＝キーを持つ 2 列以上の表）。
キーの列のすぐ右に、相手のもう 1 つの列を引く計算列を足す。見出しの語・シート名・テーブル名は表ごとに違う（vocab.py）。
正解の決まり:
  - 足す列の名前＝相手の引く列の見出し。構造化参照（INDEX/MATCH・VLOOKUP・XLOOKUP どれでもよい）で引く。
  - キーが相手に無いときは「未登録」。
  - 既に同じ名前の列があれば、その列を計算列に入れ直す（列を増やさない）。
組: 本番／A 既に名前の列がある（値が手で入っている）／B マスタが同じシートにある／C マスタの列の並びが違う（名前・コード）／
  D マスタに無いコードがある／E 大量 1,000 行／F 撃った後にもう一度／G テーブルの列が多い・番号の列
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

REQUEST = ('アクティブなシートのテーブルに、別のテーブル（共通の見出しの列＝キーを持つ表）から引く計算列を足してください。'
           '足す場所はキーの列のすぐ右、列の名前は相手の引く列の見出し。構造化参照で引き、'
           'キーが相手に無いときは「未登録」。既に同じ名前の列があれば、その列を計算列に入れ直す（列を増やさない）')
PHRASES = ['マスタから名前を引く列を足して', 'コードから名称を引いて計算列に', '別のテーブルを参照する列を入れて',
           'マスタ参照の列を足してください', 'コードの右に名前の列を', '対応する名前を引く計算列を作って']
MASTER_SHEETS = ['課マスタ', 'マスタ', 'コード表', '所属一覧', '一覧']


def build(seed, n, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    items = v.items(rng.randint(4, 8), rng)
    codes = [f"{101 + i * 3}" for i in range(len(items))]
    master = list(zip(codes, items))
    rows = []
    for i in range(n):
        c = rng.choice(codes)
        rows.append({'code': c, 'no': f"{1000 + i}", 'date': datetime.datetime(2026, rng.randint(4, 12), rng.randint(1, 28)),
                     'amount': rng.randrange(1, 900) * 1000})
    t = {'master': master, 'rows': rows, 'sheet': v.sheet('meisai'), 'table': v.table('meisai'),
         'master_sheet': rng.choice(MASTER_SHEETS), 'master_table': v.prefix + 'マスタ' if v.prefix else 'マスタ表',
         'code_head': v.word('code'), 'name_head': v.word('ka'), 'no_head': v.word('no'), 'date_head': v.word('date'),
         'amount_head': v.word('amount'), 'has_name_col': False, 'same_sheet': False, 'master_swapped': False,
         'missing': 0, 'extra_cols': False}
    t.update(kw)
    if t['missing']:
        for i in rng.sample(range(len(rows)), min(t['missing'], len(rows))):
            rows[i]['code'] = '999'
    return t


def write_before(stem, t, out=HERE):
    mm = dict(t['master'])
    for truth in (False,):
        wb = Workbook()
        ws = wb.active
        ws.title = t['sheet']
        cols = [(t['no_head'], 'no'), (t['date_head'], 'date'), (t['code_head'], 'code')]
        if t['has_name_col']:
            cols.append((t['name_head'], 'name'))
        cols.append((t['amount_head'], 'amount'))
        if t['extra_cols']:
            cols.append(('備考', 'memo'))
        for j, (name, _role) in enumerate(cols, 1):
            ws.cell(1, j, name).font = Font(bold=True)
        for i, row in enumerate(t['rows'], 2):
            for j, (name, role) in enumerate(cols, 1):
                if role == 'name':
                    ws.cell(i, j, '（手入力）' + mm.get(row['code'], '？'))
                elif role == 'memo':
                    ws.cell(i, j, random.Random(i).choice(['', '定期', '臨時']))
                else:
                    c = ws.cell(i, j, row[role])
                    if role == 'date':
                        c.number_format = 'yyyy/m/d'
                    elif role == 'amount':
                        c.number_format = '#,##0'
        tb = Table(displayName=t['table'], ref=f"A1:{get_column_letter(len(cols))}{len(t['rows']) + 1}")
        tb.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(tb)
        # マスタ
        ms = ws if t['same_sheet'] else wb.create_sheet(t['master_sheet'])
        c0 = len(cols) + 3 if t['same_sheet'] else 1
        mcols = [t['name_head'], t['code_head']] if t['master_swapped'] else [t['code_head'], t['name_head']]
        for j, name in enumerate(mcols, c0):
            ms.cell(1, j, name).font = Font(bold=True)
        for i, (c, nm) in enumerate(t['master'], 2):
            for j, name in enumerate(mcols, c0):
                ms.cell(i, j, c if name == t['code_head'] else nm)
        mt = Table(displayName=t['master_table'],
                   ref=f"{get_column_letter(c0)}1:{get_column_letter(c0 + 1)}{len(t['master']) + 1}")
        mt.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
        ms.add_table(mt)
        wb.save(os.path.join(out, f"{stem}_前.xlsx"))
    return None


def build_truth(xl, before, truth, t):
    """計算列を COM で入れる（テーブルの計算列は COM で入れないと「計算列」にならない）。"""
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        lo = ws.ListObjects(t['table'])
        names = [str(lo.ListColumns(i).Name) for i in range(1, int(lo.ListColumns.Count) + 1)]
        if t['name_head'] in names:
            lc = lo.ListColumns(t['name_head'])
        else:
            pos = names.index(t['code_head']) + 2
            lc = lo.ListColumns.Add(pos)
            lc.Name = t['name_head']
        lc.DataBodyRange.Formula = (f"=IFERROR(INDEX({t['master_table']}[{t['name_head']}],"
                                    f"MATCH([@{t['code_head']}],{t['master_table']}[{t['code_head']}],0)),\"未登録\")")
        wb.Worksheets(t['sheet']).Activate()
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out=HERE):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            write_before(stem, t, out)
            build_truth(xl, before, truth, t)
            print('  ', stem, t['sheet'], t['table'], t['master_table'], t['code_head'], t['name_head'], len(t['rows']))
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
          ('未見2_名前の列あり', build(s(2), 20, has_name_col=True)),
          ('未見3_同じシートのマスタ', build(s(3), 15, same_sheet=True, master_swapped=True)),
          ('未見4_無いコードと備考', build(s(4), 25, missing=3, extra_cols=True)),
          ('未見5_大量', build(s(5), 300))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('参照_本番', build(3, 20)),
          ('参照_試験A', build(13, 18, has_name_col=True)),
          ('参照_試験B', build(23, 15, same_sheet=True)),
          ('参照_試験C', build(33, 16, master_swapped=True)),
          ('参照_試験D', build(43, 20, missing=2)),
          ('参照_試験E', build(53, 600)),
          ('参照_試験G', build(63, 14, extra_cols=True))])
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '参照_本番_正解.xlsx'), os.path.join(HERE, f'参照_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
