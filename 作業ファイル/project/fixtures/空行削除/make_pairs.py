# -*- coding: utf-8 -*-
"""表の中の空行を消して詰める（A 掃除 6・2026-09-18）。

列の決まり: 選ぶ列は無い（表全体）。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝空でないセルが 2 つ以上並ぶ最初の行。見出しの行から表の最後の行（合計の行を含む）までの間で、
    表の列（見出しの左端〜右端）が全部空の行を、行ごと削除して詰める。
  - 見出しの行より上（表題・単位・その下の空行）は触らない。一部だけ空の行は消さない。合計の行は残す。
  - 空行が無ければ何もしない（2 回撃っても結果が変わらない）。
作り方: 正解＝空行なしの明細（detail）を書き、そこへ空行を差し込んで「前」にする。
組: 本番／A 800 行／B 表題と単位（表題の下の空行は残す）／C 合計の行（合計の直前にも空行）／D 空行が 3 行続く／
  E 列の並び違い＋一部だけ空の行／F 撃った後にもう一度／G 番号の列／H 空行なし
"""
import os
import random
import shutil
import sys

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('アクティブなシートの表の中にある空行を、行ごと削除して詰めてください。'
           '見出しの行＝空でないセルが 2 つ以上並ぶ最初の行（文字でも数でもよい）。見出しの行から表の最後の行（下の合計の行を含む）までの間で、'
           '表の列（見出しの左端の列〜右端の列）が全部空の行だけを消す。'
           '見出しの行より上（表題・単位の行・その下の空行）は触らない。一部だけ空の行は消さない。合計の行は残す。'
           '値・書式は変えない。空行が無ければ何もしない（2 回撃っても結果が変わらない）。')
PHRASES = ['表の中の空行を消して詰めて', '途中の空白行を削除して', '明細の間の空いている行を詰めて',
           '空行を取り除いて表をつなげて', '行の抜けを詰めて', '空白の行を全部削除して']


def spec(seed, n, gaps=3, run=1, before_total=False, holes=0, **kw):
    t = detail.design(seed, n, **kw)
    t['blank_rows'] = []
    rng = random.Random(seed + 71)
    body = list(range(1, n))                               # i 行目（0 始まり）の前に空行を入れる（先頭の前には入れない）
    t['gaps'] = sorted(rng.sample(body, min(gaps, len(body)))) if gaps else []
    t['run'] = run
    t['before_total'] = before_total and t['total']
    if holes:
        # 一部だけ空の行（摘要・項目の列のセルだけ空）＝消さない
        text = next(nm for nm, r in t['cols'] if r in ('text', 'kamoku'))
        for i in rng.sample(range(n), holes):
            t['rows'][i][text] = None
    return t


def make(jobs, out):
    for stem, t in jobs:
        before = os.path.join(out, f"{stem}_前.xlsx")
        truth = os.path.join(out, f"{stem}_正解.xlsx")
        lay = detail.write_before(truth, t)
        wb = load_workbook(truth)
        ws = wb.worksheets[0]
        first = lay['h'] + 1
        ins = []
        if t['before_total']:
            ins.append(lay['last'] + 1)                    # 合計の行の直前
        ins += [first + i for i in t['gaps']]
        for r in sorted(ins, reverse=True):                # 下から差し込めば上の位置はずれない
            ws.insert_rows(r, t['run'])
        wb.save(before)
        print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '空行', len(ins) * t['run'])


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', spec(s(1), rng.randint(20, 60), gaps=rng.randint(2, 5))),
          ('未見2_表題と合計行', spec(s(2), 30, title=True, unit=True, total=True, before_total=True)),
          ('未見3_並び違い', spec(s(3), 25, gaps=2, run=2, swap=True, holes=2)),
          ('未見4_番号列', spec(s(4), 28, gaps=3, no_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('空行_本番', spec(1601, 30)),
          ('空行_試験A', spec(1602, 800, gaps=25, n_items=10)),
          ('空行_試験B', spec(1603, 25, gaps=2, title=True, unit=True)),
          ('空行_試験C', spec(1604, 24, gaps=2, total=True, before_total=True)),
          ('空行_試験D', spec(1605, 26, gaps=2, run=3)),
          ('空行_試験E', spec(1606, 28, gaps=2, swap=True, holes=3)),
          ('空行_試験G', spec(1607, 22, gaps=3, no_col=True)),
          ('空行_試験H', spec(1608, 20, gaps=0, title=True)),
          ('空行_試験I', spec(1609, 24, gaps=2, no_date=True))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '空行_本番_正解.xlsx'), os.path.join(HERE, f'空行_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
