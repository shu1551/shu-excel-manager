# -*- coding: utf-8 -*-
"""別のシートの表と突き合わせる（どの表でも動く形・2026-09-18 作り直し。前の版は突合）。

シートの形で決める: アクティブなシートの表と、別のシートの表（キーの列と同じ値が並ぶ列を持つ表）。
キーの列＝両方の表に同じ値が並ぶ列。金額の列＝それぞれの表の数の列。見出しの語・シート名・列の並びは表ごとに違う。
正解の決まり:
  - 表の右の空いた列に「相手の金額」「突合」を足す（既にその見出しがあればそこを使う＝2 回撃っても同じ）。
    突合＝一致／金額違い／相手に無し。相手に無しの行の「相手の金額」は空。
  - キーは先頭の 0・数値と文字の違い・前後の空白を無視して比べる。金額はカンマ・全角・前後の空白を無視して数で比べる。
  - キーが空の行（途中の空行・最後の合計行）には書かず、数えない。
  - 「突合」の列の右に 1 列空けて件数の表。見出しは表の見出しと同じ行に「突合」「件数」。
    行は 一致・金額違い・相手に無し・こちらに無し の順で 0 件も出す（こちらに無し＝相手にだけあるキーの数）。
組: 本番／A 相手のキーが数値（こちらは 0 詰めの文字）／B 相手に表題の行と列の並び違い・日付の列／C 摘要の列／
  D B3 から・途中に空行／E 最後に合計行／F 相手の金額が文字（カンマ・全角）／G シート名が違う／H 撃った後にもう一度
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

_ZEN = str.maketrans('0123456789,', '０１２３４５６７８９，')
KEY_HEADS = ['伝票番号', '整理番号', '管理番号', '受付番号', 'No']
OLD_SHEETS = ['旧システム', '移行前', 'R7データ', '旧台帳', '現行']
NEW_SHEETS = ['新システム', '移行後', 'R8データ', '新台帳', '新しい方']
REQUEST = ('アクティブなシートの表を、別のシートの表（キーの列と同じ値が並ぶ列を持つシート）と突き合わせてください。'
           'キーの列は両方の表に同じ値が並ぶ列（値を突き合わせて決める。文字の列とは限らない・先頭 0 の番号は数の列に見える）、'
           '金額の列はそれぞれの表の数の列。'
           '相手の表の見出しの行は、値が並ぶ本文のすぐ上の行（表題や「出力日 2027/01/15」のような行は見出しではない）。'
           'キーの列は先頭 0 の文字のこともあるので、数の列を金額と決めるときはキーの列を外す。'
           '表の右の空いた列に「相手の金額」「突合」の 2 列を足す（既にその見出しがあればそこを使う）。'
           '突合は 一致／金額違い／相手に無し。相手に無しの行の「相手の金額」は空。'
           'キーは先頭の 0・数値と文字の違い・前後の空白を無視して比べ、金額はカンマ・全角・空白を無視して数で比べる。'
           'キーが空の行（空行・合計の行）には書かず、数えない。'
           '「突合」の列の右に 1 列空けて件数の表を作る（見出しは表の見出しと同じ行に「突合」「件数」）。'
           '行は 一致・金額違い・相手に無し・こちらに無し の順で 0 件も出す（こちらに無し＝相手にだけあるキーの数）')
PHRASES = ['別のシートと突き合わせて件数も出して', '相手の表と照合して差異を出して', '新旧のデータを突合して',
           '金額が合っているか照合して', '移行前と移行後を突き合わせて', '相手に無い伝票を突合で出して']


def build(seed, n, diff, miss, extra):
    rng = random.Random(seed)
    v = Vocab(seed)
    start = rng.randint(100, 5000)
    keys = sorted(rng.sample(range(start, start + n * 3), n + extra))
    names = v.kamokus(6, rng)
    old = [(k, rng.choice(names), rng.randrange(1, 900) * 1000) for k in keys[:n]]
    new = {k: a for k, _m, a in old}
    for k, _m, _a in rng.sample(old, miss):
        del new[k]
    for k in rng.sample([k for k in new], diff):
        new[k] = new[k] + rng.choice([-1, 1]) * rng.randrange(1, 50) * 100
    for k in keys[n:]:
        new[k] = rng.randrange(1, 900) * 1000
    t = {'old': old, 'new': new, 'key_head': rng.choice(KEY_HEADS), 'name_head': v.word('kamoku'),
         'amount': v.word('amount'), 'amount2': v.word('shikko'), 'memo_head': v.word('tekiyo'),
         'date_head': v.word('date'), 'sheets': (rng.choice(OLD_SHEETS), rng.choice(NEW_SHEETS)),
         'key_num': False, 'new_title': False, 'new_swapped': False, 'memo': False, 'r0': 1, 'c0': 1,
         'blanks': (), 'total': False, 'amount_text': False}
    return t


def write_pair(stem, t, out=HERE):
    rng = random.Random(stem)
    old, new = t['old'], t['new']
    for truth in (False, True):
        wb = Workbook()
        ws = wb.active
        ws.title = t['sheets'][0]
        r0, c0 = t['r0'], t['c0']
        h = r0
        base = [t['key_head'], t['name_head'], t['amount']] + ([t['memo_head']] if t['memo'] else [])
        add = c0 + len(base)
        heads = base + (['相手の金額', '突合'] if truth else [])
        for j, v in enumerate(heads):
            ws.cell(h, c0 + j, v).font = Font(bold=True)
        counts = {'一致': 0, '金額違い': 0, '相手に無し': 0,
                  'こちらに無し': len([k for k in new if k not in {o[0] for o in old}])}
        r = h + 1
        for i, (k, m, a) in enumerate(old):
            if i in t['blanks']:
                r += 1
            ws.cell(r, c0, f"{k:06d}").number_format = '@'
            ws.cell(r, c0 + 1, m)
            ws.cell(r, c0 + 2, a).number_format = '#,##0'
            if t['memo']:
                ws.cell(r, c0 + 3, f"{m}の支払")
            if truth:
                if k in new:
                    ws.cell(r, add, new[k]).number_format = '#,##0'
                    res = '一致' if new[k] == a else '金額違い'
                else:
                    res = '相手に無し'
                ws.cell(r, add + 1, res)
                counts[res] += 1
            r += 1
        if t['total']:
            ws.cell(r, c0 + 1, '合計').font = Font(bold=True)
            ws.cell(r, c0 + 2, sum(a for _k, _m, a in old)).number_format = '#,##0'
        if truth:
            sc = add + 3
            ws.cell(h, sc, '突合').font = Font(bold=True)
            ws.cell(h, sc + 1, '件数').font = Font(bold=True)
            for i, lab in enumerate(['一致', '金額違い', '相手に無し', 'こちらに無し'], h + 1):
                ws.cell(i, sc, lab)
                ws.cell(i, sc + 1, counts[lab])
        ns = wb.create_sheet(t['sheets'][1])
        nh = 1
        if t['new_title']:
            ns.cell(1, 1, '新しいシステムの出力一覧').font = Font(bold=True)
            ns.cell(2, 1, '出力日 2027/01/15')
            nh = 4
        cols = [t['amount2'], t['date_head'], t['key_head']] if t['new_swapped'] else [t['key_head'], t['amount2']]
        for j, v in enumerate(cols, 1):
            ns.cell(nh, j, v).font = Font(bold=True)
        items = list(new.items())
        rng.shuffle(items)
        for i, (k, a) in enumerate(items, nh + 1):
            vals = {t['key_head']: k if t['key_num'] else f"{k:06d}", t['amount2']: a,
                    t['date_head']: f"2027/1/{(k % 28) + 1}"}
            if t['amount_text']:
                vals[t['amount2']] = f"{a:,}".translate(_ZEN) if rng.random() < 0.5 else f" {a:,} "
            for j, v in enumerate(cols, 1):
                c = ns.cell(i, j, vals[v])
                if v == t['key_head'] and not t['key_num']:
                    c.number_format = '@'
        wb.save(os.path.join(out, f"{stem}_{'正解' if truth else '前'}.xlsx"))
    print('  ', stem, t['sheets'], t['key_head'], t['amount'], '/', t['amount2'], len(old))


def kw(t, **over):
    t.update(over)
    return t


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    b = lambda k: build(s(k), rng.randint(15, 50), rng.randint(0, 5), rng.randint(0, 4), rng.randint(0, 4))  # noqa: E731
    write_pair('未見1_標準', b(1), out)
    write_pair('未見2_番号が数値', kw(b(2), key_num=True), out)
    write_pair('未見3_相手に表題と並び違い', kw(b(3), new_title=True, new_swapped=True), out)
    write_pair('未見4_摘要と空行', kw(b(4), memo=True, blanks=(3, 9)), out)
    write_pair('未見5_合計行と位置', kw(b(5), total=True, r0=3, c0=2), out)
    write_pair('未見6_文字の金額', kw(b(6), amount_text=True), out)
    write_pair('未見7_全部入り', kw(b(7), key_num=True, new_title=True, new_swapped=True, memo=True, r0=2, c0=2,
                                blanks=(5,), total=True, amount_text=True), out)
    write_pair('未見8_大量', build(s(8), 800, 40, 30, 25), out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    write_pair('突合_本番', build(11, 25, 3, 2, 2))
    write_pair('突合_試験A', kw(build(23, 30, 4, 3, 3), key_num=True))
    write_pair('突合_試験B', kw(build(31, 20, 2, 1, 2), new_title=True, new_swapped=True))
    write_pair('突合_試験C', kw(build(43, 22, 2, 2, 1), memo=True))
    write_pair('突合_試験D', kw(build(51, 26, 3, 2, 2), r0=3, c0=2, blanks=(6, 13)))
    write_pair('突合_試験E', kw(build(67, 18, 1, 1, 1), total=True))
    write_pair('突合_試験F', kw(build(73, 24, 3, 2, 2), amount_text=True))
    write_pair('突合_試験G', kw(build(89, 20, 2, 1, 1), new_title=True, memo=True, total=True))
    shutil.copyfile(os.path.join(HERE, '突合_本番_正解.xlsx'), os.path.join(HERE, '突合_試験H_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '突合_本番_正解.xlsx'), os.path.join(HERE, '突合_試験H_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
