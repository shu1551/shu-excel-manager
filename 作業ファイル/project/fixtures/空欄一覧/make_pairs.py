# -*- coding: utf-8 -*-
"""選んでいる列（必須の列）の空欄の行を右に一覧にする（I 検査 73・2026-09-18）。

列の決まり: 選んでいる 1 列＝必須の列。選んでいる列は <名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 置き場所＝表の右に 1 列空けた所。見出しの行は選んでいるセルの行。
  - 見出しは「空欄の行」の 1 列だけ。空欄（空・空白だけ）のセルがある行の行番号を上から順に 1 行ずつ。
  - 途中の空行（その行がまるごと空）は数えない。表の下の合計の行も数えない。
  - 空欄が 1 つも無ければ見出しだけ書く。既に同じ表が右にあれば作り直す。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 空行／E 列の並び違い／F 撃った後にもう一度／G 空白だけの文字／H 空欄なし
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

REQUEST = ('選んでいる列（必須の列）に空欄がある行を、表の右に 1 列空けて一覧にしてください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '見出しは「空欄の行」の 1 列だけ。空欄（空のセル、または空白だけの文字）のセルがある行の行番号を、上から順に'
           '1 行ずつ書く。行がまるごと空の行（表の途中の空行）は数えない。表の下に合計（合計・計）の行があればそれも数えない。'
           '空欄が 1 つも無ければ見出しだけ書く。既に同じ表が右にあれば作り直す。')
PHRASES = ['選んだ列の空欄がある行を一覧にして', '必須の列の空欄チェックをして', '入力漏れの行を右に出して',
           '選んでいる列で未記入の行を洗い出して', '空欄になっている行番号を出して', '記入漏れの一覧を作って']


def hole(seed, n, spaces=False, holes=4, **kw):
    t = detail.design(seed, n, **kw)
    rng = random.Random(seed + 55)
    name = detail.name_of(t, 'text')
    idx = sorted(rng.sample(range(n), min(holes, n)))
    t['holes'] = idx
    for i in idx:
        t['rows'][i][name] = ('   ' if spaces and i % 2 else None)
    return t


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0, h, last = lay['c0'], lay['h'], lay['last']
        tc = c0 + detail.col_of(t, 'text') - 1
        out = c0 + lay['ncols'] + 1
        ws.Cells(h, out).Value = '空欄の行'
        k = 0
        for r in range(h + 1, last + 1):
            row_empty = all(ws.Cells(r, c).Value is None for c in range(c0, c0 + lay['ncols']))
            if row_empty:
                continue
            v = ws.Cells(r, tc).Value
            if v is None or str(v).strip() == '':
                k += 1
                ws.Cells(h + k, out).Value = r
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
            sel = f"{get_column_letter(detail.col_of(t, 'text'))}{lay['h']}"
            build_truth(xl, before, truth, t, lay)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '選ぶ列', sel, '空欄', len(t.get('holes') or []))
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', hole(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計行', hole(s(2), 30, title=True, unit=True, total=True)),
          ('未見3_並び違いと空白文字', hole(s(3), 25, spaces=True, swap=True)),
          ('未見4_空行', hole(s(4), 28, blanks=2))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('空欄_本番', hole(1101, 30)),
          ('空欄_試験A', hole(1102, 800, holes=12, n_items=10)),
          ('空欄_試験B', hole(1103, 25, title=True, unit=True)),
          ('空欄_試験C', hole(1104, 24, total=True)),
          ('空欄_試験D', hole(1105, 26, blanks=2)),
          ('空欄_試験E', hole(1106, 28, swap=True)),
          ('空欄_試験G', hole(1107, 22, spaces=True)),
          ('空欄_試験H', hole(1108, 20, holes=0)),
          ('空欄_試験I', hole(1109, 24, no_date=True))], HERE)   # 日付の無い表（2026-09-18 夜）
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '空欄_本番_正解.xlsx'), os.path.join(HERE, f'空欄_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '空欄_本番_選択.txt'), os.path.join(HERE, '空欄_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
