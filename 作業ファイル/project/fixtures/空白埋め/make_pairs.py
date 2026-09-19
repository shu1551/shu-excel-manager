# -*- coding: utf-8 -*-
"""選んでいる列の空白セルを、上の値で埋める（A 掃除 3・2026-09-18）。

列の決まり: 選んでいる 1 列。選んでいる列は <名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 選んだ列の本文の空のセルに、その上にある一番近い値を入れる（上書きしてよい）。
  - 本文の先頭より上には書かない。見出しの行・合計の行・行がまるごと空の行は触らない。
  - ほかの列は触らない。2 回撃っても結果が変わらない。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 空行／E 列の並び違い／F 撃った後にもう一度／G 番号の列／H 空白なし
"""
import os
import random
import shutil
import sys

from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('選んでいる列の本文の空のセルを、その上にある一番近い値で埋めてください（上書きしてよい）。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '本文の先頭より上には書かない。見出しの行・表の下の合計（合計・計）の行・行がまるごと空の行は触らない。'
           'ほかの列は触らない。2 回撃っても結果が変わらない。')
PHRASES = ['選んだ列の空白を上の値で埋めて', '空いているセルに上と同じ値を入れて', '省略されている課名を上から埋めて',
           '選んでいる列の空欄を上の値でうめて', '同じ値の繰り返しを省いてある列を埋めて', '上のセルの値をコピーして空白を埋めて']


def gaps(seed, n, holes=None, **kw):
    """項目の列で、同じ値が続く 2 つ目以降を空にする（省略して書いた表）。"""
    t = detail.design(seed, n, **kw)
    name = detail.name_of(t, 'item')
    rng = random.Random(seed + 33)
    t['rows'].sort(key=lambda r: str(r[name]))              # 同じ項目が続く並び（省略して書く表の形）
    if holes == 0:
        return t
    prev = None
    for row in t['rows']:
        v = row[name]
        if v == prev and rng.random() < 0.9:
            row[name] = None
        else:
            prev = v
    return t


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0, h, last = lay['c0'], lay['h'], lay['last']
        ic = c0 + detail.col_of(t, 'item') - 1
        above = None
        for r in range(h + 1, last + 1):
            row_empty = all(ws.Cells(r, c).Value is None for c in range(c0, c0 + lay['ncols']))
            if row_empty:
                continue
            v = ws.Cells(r, ic).Value
            if v is None or str(v).strip() == '':
                if above is not None:
                    ws.Cells(r, ic).Value = above
            else:
                above = v
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
            lay = detail.write_before(before, t)
            sel = f"{get_column_letter(detail.col_of(t, 'item'))}{lay['h']}"
            build_truth(xl, before, truth, t, lay)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', gaps(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計行', gaps(s(2), 30, title=True, unit=True, total=True)),
          ('未見3_並び違い', gaps(s(3), 25, swap=True)),
          ('未見4_空行と番号列', gaps(s(4), 28, blanks=2, no_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('空白_本番', gaps(1301, 30)),
          ('空白_試験A', gaps(1302, 800, n_items=10)),
          ('空白_試験B', gaps(1303, 25, title=True, unit=True)),
          ('空白_試験C', gaps(1304, 24, total=True)),
          ('空白_試験D', gaps(1305, 26, blanks=2)),
          ('空白_試験E', gaps(1306, 28, swap=True)),
          ('空白_試験G', gaps(1307, 22, no_col=True)),
          ('空白_試験H', gaps(1308, 20, holes=0)),
          ('空白_試験I', gaps(1309, 24, no_date=True))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '空白_本番_正解.xlsx'), os.path.join(HERE, f'空白_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '空白_本番_選択.txt'), os.path.join(HERE, '空白_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
