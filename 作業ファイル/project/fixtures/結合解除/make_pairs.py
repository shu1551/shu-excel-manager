# -*- coding: utf-8 -*-
"""表の結合セルを解除して、結合していた範囲に左上の値を埋める（A 掃除 4・2026-09-18）。

列の決まり: 選ぶ列は無い（表全体）。
正解の決まり（依頼文に全部書く）:
  - 表の中の結合セルを全部解除し、解除した範囲のセル全部に、結合の左上にあった値を入れる。
  - 値・見出しの語は変えない。行や列は増やさない。書式（罫線・色・幅）は触らない。
  - 結合が 1 つも無ければ何もしない（2 回撃っても結果が変わらない）。
組: 本番 縦の結合／A 2 段見出し／B 表題の結合つき／C 縦と横の両方／D 800 行／E 結合なし／F 撃った後にもう一度／
  G 番号の列と合計行／H 結合が 3 行まとめ
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

REQUEST = ('アクティブなシートの表の結合セルを全部解除して、解除した範囲のセル全部に、結合の左上にあった値を入れてください。'
           '値と見出しの語は変えない。行や列は増やさない。罫線・色・列幅などの書式は触らない。'
           '結合が 1 つも無ければ何もしない（2 回撃っても結果が変わらない）。')
PHRASES = ['結合セルを解除して値を埋めて', 'セルの結合をはずして同じ値を入れて', '結合を解いて空いた所を埋めて',
           '結合セルをばらして値を入れて', '結合をぜんぶ解除して', 'マージを解除して値を埋めてほしい']


def build(seed, n, kind='tate'):
    rng = random.Random(seed)
    v = Vocab(seed)
    t = {'sheet': v.sheet('shukei'), 'seed': seed, 'n': n, 'kind': kind, 'vocab': v,
         'ka': v.word('ka'), 'kamoku': v.word('kamoku'), 'yosan': v.word('yosan'), 'shikko': v.word('shikko'),
         'no': v.word('no'), 'title': None, 'total': False, 'r0': 1, 'group': max(2, n // 4)}
    if kind in ('title', 'both'):
        t['title'] = f"令和8年度 {t['ka']}別の状況"
        t['r0'] = 2
    if kind == 'sokei':
        t['total'] = True
    items = v.items(max(2, n // t['group'] + 1), rng)
    t['rows'] = [(items[i // t['group']], rng.choice(v.kamokus(5, rng)),
                  rng.randrange(100, 900) * 1000, rng.randrange(100, 900) * 1000) for i in range(n)]
    return t


def write_before(path, t, merged=True):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    r = t['r0']
    two = t['kind'] in ('nidan', 'both')
    no = t['kind'] == 'sokei'
    c0 = 1
    if t['title']:
        ws.cell(1, 1, t['title']).font = Font(bold=True, size=14)
        if merged:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4 + (1 if no else 0))
    head = r
    cols = ([t['no']] if no else []) + [t['ka'], t['kamoku'], t['yosan'], t['shikko']]
    if two:
        # 2 段見出し: 上段「予算の状況」を 2 列結合、下段に 予算額／執行額
        ws.cell(head, c0 + len(cols) - 2, '予算の状況').font = Font(bold=True)
        if merged:
            ws.merge_cells(start_row=head, start_column=c0 + len(cols) - 2, end_row=head,
                           end_column=c0 + len(cols) - 1)
        else:
            ws.cell(head, c0 + len(cols) - 1, '予算の状況').font = Font(bold=True)
        for j, nm in enumerate(cols[:-2], c0):
            ws.cell(head, j, nm).font = Font(bold=True)
            if merged:
                ws.merge_cells(start_row=head, start_column=j, end_row=head + 1, end_column=j)
            else:
                ws.cell(head + 1, j, nm).font = Font(bold=True)
        ws.cell(head + 1, c0 + len(cols) - 2, cols[-2]).font = Font(bold=True)
        ws.cell(head + 1, c0 + len(cols) - 1, cols[-1]).font = Font(bold=True)
        r = head + 2
    else:
        for j, nm in enumerate(cols, c0):
            ws.cell(head, j, nm).font = Font(bold=True)
        r = head + 1
    first = r
    kc = c0 + (1 if no else 0)
    for i, (ka, km, y, s) in enumerate(t['rows']):
        j = c0
        if no:
            ws.cell(r, j, 1000 + i)
            j += 1
        ws.cell(r, j, ka)
        ws.cell(r, j + 1, km)
        ws.cell(r, j + 2, y).number_format = '#,##0'
        ws.cell(r, j + 3, s).number_format = '#,##0'
        r += 1
    last = r - 1
    if t['total']:
        ws.cell(r, kc, '合計').font = Font(bold=True)
        ws.cell(r, kc + 2, sum(x[2] for x in t['rows'])).number_format = '#,##0'
        ws.cell(r, kc + 3, sum(x[3] for x in t['rows'])).number_format = '#,##0'
    if t['kind'] != 'none':
        # 項目の列は同じ値が続く所を縦に結合（結合の左上にだけ値が残る）
        i = 0
        while i < len(t['rows']):
            k = i
            while k + 1 < len(t['rows']) and t['rows'][k + 1][0] == t['rows'][i][0]:
                k += 1
            if k > i:
                if merged:
                    ws.merge_cells(start_row=first + i, start_column=kc, end_row=first + k, end_column=kc)
                # 解除後は全部に値が入る＝正解では merged=False でそのまま書いてある
            i = k + 1
    wb.save(path)
    return {'h': head, 'first': first, 'last': last}


def make(jobs, out):
    for stem, t in jobs:
        write_before(os.path.join(out, f"{stem}_前.xlsx"), t, merged=True)
        write_before(os.path.join(out, f"{stem}_正解.xlsx"), t, merged=False)
        print('  ', stem, t['sheet'], t['kind'], len(t['rows']))


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_縦の結合', build(s(1), rng.randint(12, 40))),
          ('未見2_2段見出し', build(s(2), 20, 'nidan')),
          ('未見3_表題と2段', build(s(3), 16, 'both')),
          ('未見4_番号と合計行', build(s(4), 18, 'sokei'))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('結合_本番', build(1201, 20)),
          ('結合_試験A', build(1202, 24, 'nidan')),
          ('結合_試験B', build(1203, 18, 'title')),
          ('結合_試験C', build(1204, 20, 'both')),
          ('結合_試験D', build(1205, 800)),
          ('結合_試験E', build(1206, 16, 'none')),
          ('結合_試験G', build(1207, 18, 'sokei')),
          ('結合_試験H', build(1208, 30))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '結合_本番_正解.xlsx'), os.path.join(HERE, f'結合_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
