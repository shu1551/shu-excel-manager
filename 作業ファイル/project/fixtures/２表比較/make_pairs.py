# -*- coding: utf-8 -*-
"""同じ形の 2 つのシートを比べる表を作る（どの表でも動く形・2026-09-18 作り直し。前の版は前年度比較）。

形で決める: アクティブなシート（こちら）と、同じ見出しの別のシート（相手）。項目の列＝左端の文字の列、金額の列＝数の列。
比較表の見出しは 項目の列の見出し・このシートの名前・相手のシートの名前・増減額・増減率。
見出しの語・シート名・項目の名前・位置は表ごとに違う（vocab.py）。
正解の決まり:
  - 左上＝こちらの見出しの行・表の右端の 2 つ右（1 列空ける）。行の順＝こちらの順 → 相手にだけある項目（相手の順）。
  - 片方に無い額は 0。増減額＝こちら−相手。増減率＝増減額÷相手（相手が 0 なら「皆増」、こちらが 0 なら「皆減」）。
  - 最後に「合計」の行。項目の前後の空白は無視。項目が空の行・合計や計の行は入れない。
  - 金額はカンマ・全角・前後の空白・「円」つきの文字でも数で読む。前の比較表が残っていたら消してから書く。
組: 本番 12 項目／A 25 項目／B 相手に無い項目／C こちらに無い項目／D 金額が文字／E 表題と位置／F 両方に合計行／
  G 項目の空白の揺れ／H 前の比較表が残る／I 撃った後にもう一度／J シート名が年度／K 前の比較表が今回より長い
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
from vocab import Vocab, KAMOKUS   # noqa: E402

_ZEN = str.maketrans('0123456789,', '０１２３４５６７８９，')
SHEET_PAIRS = [('今年度', '前年度'), ('令和8年度', '令和7年度'), ('R8当初', 'R7当初'), ('当年', '前年'), ('新計画', '旧計画')]
REQUEST = ('アクティブなシートの表と、同じ見出しの別のシートの表を、項目の列（左端の文字の列）で突き合わせて、'
           'アクティブなシートの表の右に 1 列空けて比較表を作ってください。'
           '比較表の左上は、こちらの表の見出しの行（値が並ぶ本文のすぐ上の行。表題の行は見出しではない）・表の右端の 2 つ右。'
           '見出しは 1 行で 項目の列の見出し・このシートの名前・相手のシートの名前・増減額・増減率。'
           '行はこのシートの順、そのあと相手にだけある項目（相手の順）。片方に無い額は 0。増減額＝こちら−相手。'
           '増減率＝増減額÷相手（相手が 0 なら文字「皆増」、こちらが 0 なら文字「皆減」）。最後に「合計」の行。'
           '項目の前後の空白は無視して同じ項目にし、項目が空の行と 合計・計 の行は入れない。'
           '金額はカンマ・全角・前後の空白・「円」つきの文字でも数で読む。前の比較表が残っていたら消してから書く')
PHRASES = ['前年度と比べる表を作って', '2つのシートを比べて増減を出して', '相手のシートとの差を出して',
           '増減額と増減率の表を作って', '比較表を右に作ってください', 'こちらと相手を突き合わせて増減を']


def build(seed, n, new=0, gone=0):
    rng = random.Random(seed)
    v = Vocab(seed)
    pool = KAMOKUS[v.kamoku_kind]
    names = rng.sample(pool, n) if n <= len(pool) else [f"{k}{i}" for i, k in enumerate(rng.choices(pool, k=n))]
    this = [(k, rng.randrange(1, 900) * 1000) for k in names]
    prev = [(k, max(0, a + rng.randrange(-200, 200) * 1000) or 1000) for k, a in this]
    rng.shuffle(prev)
    if new:
        drop = {k for k, _a in rng.sample(this, new)}
        prev = [(k, a) for k, a in prev if k not in drop]
    if gone:
        extra = [f"{k}（旧）" for k in rng.sample(pool, min(gone, len(pool)))]
        prev += [(k, rng.randrange(1, 500) * 1000) for k in extra]
    t = {'this': this, 'prev': prev, 'item': v.word('kamoku'), 'amount': v.word('yosan'),
         'sheets': rng.choice(SHEET_PAIRS), 'title': False, 'r0': 1, 'c0': 1, 'text_amount': False,
         'total': False, 'spaces': False, 'stale': 0}
    return t


def compare(t):
    pm = {k: a for k, a in t['prev']}
    tm = {k: a for k, a in t['this']}
    rows = [(k, a, pm.get(k, 0)) for k, a in t['this']] + [(k, 0, a) for k, a in t['prev'] if k not in tm]
    out = []
    for k, x, p in rows + [('合計', sum(r[1] for r in rows), sum(r[2] for r in rows))]:
        d = x - p
        rate = '皆増' if p == 0 else ('皆減' if x == 0 else d / p)
        out.append([k, x, p, d, rate])
    return out


def write_pair(stem, t, out=HERE):
    head = [t['item'], t['sheets'][0], t['sheets'][1], '増減額', '増減率']
    for truth in (False, True):
        wb = Workbook()
        ws = wb.active
        ws.title = t['sheets'][0]
        h, c0 = t['r0'], t['c0']
        if t['title']:
            ws.cell(t['r0'], c0, '令和8年度 当初予算（一般会計・歳出）').font = Font(bold=True, size=14)
            h = t['r0'] + 2

        def put(sh, hr, cc, items, key):
            r3 = random.Random(key)
            sh.cell(hr, cc, t['item']).font = Font(bold=True)
            sh.cell(hr, cc + 1, t['amount']).font = Font(bold=True)
            for i, (k, a) in enumerate(items, 1):
                name = k
                if t['spaces'] and r3.random() < 0.3:
                    name = r3.choice((' ', '　')) + k if r3.random() < 0.5 else k + r3.choice((' ', '　'))
                sh.cell(hr + i, cc, name)
                v = a
                if t['text_amount'] and r3.random() < 0.5:
                    x = r3.random()
                    v = f"{a:,}".translate(_ZEN) if x < 0.3 else (f" {a:,} " if x < 0.6 else f"{a:,}円")
                sh.cell(hr + i, cc + 1, v).number_format = '#,##0'
            last = hr + len(items)
            if t['total']:
                sh.cell(last + 1, cc, '合計').font = Font(bold=True)
                sh.cell(last + 1, cc + 1, sum(a for _k, a in items)).number_format = '#,##0'
        put(ws, h, c0, t['this'], stem + 't')
        lc = c0 + 3
        if truth:
            for j, v in enumerate(head):
                ws.cell(h, lc + j, v).font = Font(bold=True)
            for i, row in enumerate(compare(t), 1):
                for j, v in enumerate(row):
                    ws.cell(h + i, lc + j, v)
        elif t['stale']:
            for j, v in enumerate(head):
                ws.cell(h, lc + j, v)
            for i in range(1, t['stale'] + 1):
                for j, v in enumerate([f'旧項目{i}', 1000 * i, 900 * i, 100 * i, 0.1]):
                    ws.cell(h + i, lc + j, v)
        pv = wb.create_sheet(t['sheets'][1])
        put(pv, 1, 1, t['prev'], stem + 'p')
        wb.save(os.path.join(out, f"{stem}_{'正解' if truth else '前'}.xlsx"))
    print('  ', stem, t['sheets'], t['item'], t['amount'], len(t['this']))


def kw(t, **over):
    t.update(over)
    return t


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    write_pair('未見1_標準', build(s(1), rng.randint(6, 20)), out)
    write_pair('未見2_新規と廃止', build(s(2), 14, new=3, gone=2), out)
    write_pair('未見3_文字の金額と空白', kw(build(s(3), 12, new=1), text_amount=True, spaces=True), out)
    write_pair('未見4_表題と位置と合計行', kw(build(s(4), 10, gone=1), title=True, r0=2, c0=2, total=True), out)
    write_pair('未見5_前の比較表', kw(build(s(5), 8), stale=15), out)
    write_pair('未見6_全部入り', kw(build(s(6), 18, new=2, gone=3), title=True, r0=3, c0=2, text_amount=True,
                                total=True, spaces=True, stale=25), out)
    write_pair('未見7_大量', build(s(7), 200, new=20, gone=15), out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    write_pair('比較_本番', build(3, 12))
    write_pair('比較_試験A', build(13, 25))
    write_pair('比較_試験B', build(23, 12, new=3))
    write_pair('比較_試験C', build(33, 10, gone=2))
    write_pair('比較_試験D', kw(build(43, 11), text_amount=True))
    write_pair('比較_試験E', kw(build(53, 9, new=1), title=True, r0=3, c0=2))
    write_pair('比較_試験F', kw(build(63, 10, gone=1), total=True))
    write_pair('比較_試験G', kw(build(73, 12), spaces=True))
    write_pair('比較_試験H', kw(build(83, 8), stale=12))
    write_pair('比較_試験K', kw(build(103, 6, gone=1), stale=20))
    shutil.copyfile(os.path.join(HERE, '比較_本番_正解.xlsx'), os.path.join(HERE, '比較_試験I_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '比較_本番_正解.xlsx'), os.path.join(HERE, '比較_試験I_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
