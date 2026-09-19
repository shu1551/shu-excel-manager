# -*- coding: utf-8 -*-
"""コードを対応表で振り直して集計する（どの表でも動く形・2026-09-18 作り直し。前の版は振り直し）。

シートの形で決める: 明細＝アクティブなシート。対応表＝別のシートで、明細のコードの列と同じ値が並ぶ列を持つ表
（その隣の列が区分）。見出しの語・シート名・区分の名前・列の並びは表ごとに違う（vocab.py）。
正解の決まり:
  - 明細の右の空いた列に「対応表の区分の列の見出し」を書き、行ごとに引いた区分を書く（対応表に無いコードは「対応なし」）。
    既に同じ見出しの列があればそこを使う（2 回撃っても同じ）。人の列は上書きしない。
  - コードの前後の空白は無視して引く（明細のコードのセルは書き換えない）。
  - コードが空の行（途中の空行・合計の行）には区分を書かず、集計にも入れない。
  - 足した列の右に 1 列空けて集計表。見出しは明細の見出しと同じ行に「区分の見出し」「金額の列の見出し」。
  - 区分は対応表に出てくる順（重複は 1 回）。明細に 1 件も無い区分も 0 で並べる。
    その下に「対応なし」（無くても 0 で出す）、最後に「合計」（明細の行の合計。明細の合計行は入れない）。
組: 本番 30 行／A 45 行・区分の順が違う／B 表題つき・0 円の区分／C 摘要の列／D 位置と空行／E 合計行／
  F 対応表の並びが違う（区分・備考・コード）／G コードの前後に空白／H 撃った後にもう一度／I 対応表が 3 列・別の区分の組
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

# 区分の組（表ごとに変える）
KUBUN_SETS = [
    ['人件費', '物件費', '維持補修費', '扶助費', '補助費等', '普通建設事業費'],
    ['義務的経費', '投資的経費', 'その他の経費'],
    ['一般管理', '施設管理', '事業費', '補助事業'],
    ['A分類', 'B分類', 'C分類', 'D分類'],
]
KUBUN_HEADS = ['決算統計区分', '統計区分', '区分', '分類', '性質別区分']
CODE_HEADS = ['予算科目コード', '科目コード', 'コード', '細節コード', '費目コード']
NAME_HEADS = ['科目名', '節の名称', '費目名', '名称']
MAP_SHEETS = ['対応表', '区分表', 'コード対応', 'マスタ', '読替表']
REQUEST = ('アクティブなシートの明細のコードの列を、別のシートの対応表（明細のコードと同じ値が並ぶ列と、その隣の区分の列）で'
           '振り直して集計してください。明細の右の空いた列に対応表の区分の列の見出しを書き、行ごとに引いた区分を書く'
           '（対応表に無いコードは「対応なし」）。既に同じ見出しの列があればそこを使う。コードの前後の空白は無視して引く'
           '（明細のコードのセルは書き換えない）。コードが空の行（空行・合計の行）には書かず、集計にも入れない。'
           '集計表は、区分の列の 2 つ右から必ず書く（そこから右に何か残っていれば消してから書く）。'
           '集計表の見出しは明細の見出しと同じ行に 区分の見出し・金額の列の見出し。'
           '区分は対応表に出てくる順（重複は 1 回）、明細に無い区分も 0 で並べ、その下に「対応なし」（無くても 0）、'
           '最後に「合計」（明細の金額の合計）。金額の列は明細の数の列')
PHRASES = ['対応表で区分を振り直して集計して', 'コードから区分を引いて区分ごとに集計', '対応表を見て区分を付けて集計表も',
           '区分に振り分けて合計を出して', 'コードを読み替えて分類ごとに集計', '対応表で置き換えて区分別の集計を']


def build(seed, rows, missing, zero_kubun=False, kubun_set=None):
    rng = random.Random(seed)
    v = Vocab(seed)
    kubun = list(kubun_set or rng.choice(KUBUN_SETS))
    rng.shuffle(kubun)
    names = v.kamokus(6, rng)
    codes = []
    for i, k in enumerate(kubun):
        for j in range(rng.randint(2, 4)):
            codes.append((f"{10 + i * 2}-{j + 1:02d}", rng.choice(names), k))
    table = [(c, k) for c, _n, k in codes]
    drop = kubun[-1] if zero_kubun else None
    usable = [p for p in codes if p[2] != drop]
    body = []
    for _ in range(rows - missing):
        c, name, _k = rng.choice(usable)
        body.append([c, name, rng.randrange(3, 900) * 1000])
    for _ in range(missing):
        body.append([f"99-{rng.randint(1, 9):02d}", rng.choice(names), rng.randrange(3, 90) * 1000])
    rng.shuffle(body)
    t = {'table': table, 'body': body, 'sheet': v.sheet('meisai'), 'map_sheet': rng.choice(MAP_SHEETS),
         'code_head': rng.choice(CODE_HEADS), 'name_head': rng.choice(NAME_HEADS), 'amount': v.word('amount'),
         'kubun_head': rng.choice(KUBUN_HEADS), 'memo_head': v.word('tekiyo'),
         'title': None, 'memo': False, 'r0': 1, 'c0': 1, 'blanks': (), 'total': False, 'map_layout': 'std',
         'spaces': False, 'map_extra': False}
    return t


def write_pair(stem, t, out=HERE):
    kmap = dict(t['table'])
    kubun = []
    for _c, k in t['table']:
        if k not in kubun:
            kubun.append(k)
    rng = random.Random(stem)
    shown = []
    for c, _n, _a in t['body']:
        shown.append((' ' + c if rng.random() < 0.5 else c + '  ') if t['spaces'] and rng.random() < 0.6 else c)
    for truth in (False, True):
        wb = Workbook()
        ws = wb.active
        ws.title = t['sheet']
        r0, c0 = t['r0'], t['c0']
        h = r0
        if t['title']:
            ws.cell(r0, c0, t['title']).font = Font(bold=True, size=14)
            h = r0 + 2
        base = [t['code_head'], t['name_head'], t['amount']] + ([t['memo_head']] if t['memo'] else [])
        kcol = c0 + len(base)
        heads = base + ([t['kubun_head']] if truth else [])
        for j, v in enumerate(heads):
            ws.cell(h, c0 + j, v).font = Font(bold=True)
        r = h + 1
        for i, ((c, name, amt), code) in enumerate(zip(t['body'], shown)):
            if i in t['blanks']:
                r += 1
            ws.cell(r, c0, code).number_format = '@'
            ws.cell(r, c0 + 1, name)
            ws.cell(r, c0 + 2, amt).number_format = '#,##0'
            if t['memo']:
                ws.cell(r, c0 + 3, f"{name[:2]}の支払 {i + 1}")
            if truth:
                ws.cell(r, kcol, kmap.get(c, '対応なし'))
            r += 1
        if t['total']:
            ws.cell(r, c0 + 1, '合計').font = Font(bold=True)
            ws.cell(r, c0 + 2, sum(a for _c, _n, a in t['body'])).number_format = '#,##0'
        if truth:
            sums = {k: 0 for k in kubun}
            none = 0
            for c, _n, amt in t['body']:
                if c in kmap:
                    sums[kmap[c]] += amt
                else:
                    none += amt
            sc = kcol + 2
            ws.cell(h, sc, t['kubun_head']).font = Font(bold=True)
            ws.cell(h, sc + 1, t['amount']).font = Font(bold=True)
            rr = h + 1
            for k in kubun + ['対応なし']:
                ws.cell(rr, sc, k)
                ws.cell(rr, sc + 1, sums.get(k, none if k == '対応なし' else 0)).number_format = '#,##0'
                rr += 1
            ws.cell(rr, sc, '合計')
            ws.cell(rr, sc + 1, sum(a for _c, _n, a in t['body'])).number_format = '#,##0'
        mt = wb.create_sheet(t['map_sheet'])
        if t['map_layout'] == 'swapped':
            cols = [t['kubun_head'], '備考', t['code_head']]
        elif t['map_extra']:
            cols = [t['code_head'], t['kubun_head'], '備考']
        else:
            cols = [t['code_head'], t['kubun_head']]
        for j, v in enumerate(cols, 1):
            mt.cell(1, j, v).font = Font(bold=True)
        for i, (c, k) in enumerate(t['table'], 2):
            row = {t['code_head']: c, t['kubun_head']: k, '備考': '' if i % 3 else '令和7年度から'}
            for j, v in enumerate(cols, 1):
                cell = mt.cell(i, j, row[v] or None)
                if v == t['code_head']:
                    cell.number_format = '@'
        wb.save(os.path.join(out, f"{stem}_{'正解' if truth else '前'}.xlsx"))
    print('  ', stem, t['sheet'], t['map_sheet'], t['code_head'], t['kubun_head'], len(t['body']))


def kw(t, **over):
    t.update(over)
    return t


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    write_pair('未見1_標準', build(s(1), rng.randint(15, 60), rng.randint(0, 4)), out)
    write_pair('未見2_表題と0の区分', kw(build(s(2), 25, 2, zero_kubun=True), title='令和8年度 支出明細（建設課）'), out)
    write_pair('未見3_摘要', kw(build(s(3), 30, 1), memo=True), out)
    write_pair('未見4_位置と空行', kw(build(s(4), 30, 2), r0=rng.randint(2, 5), c0=rng.randint(2, 4),
                                blanks=tuple(sorted(rng.sample(range(2, 28), 3)))), out)
    write_pair('未見5_合計行と空白', kw(build(s(5), 25, 1), total=True, spaces=True), out)
    write_pair('未見6_対応表の並び', kw(build(s(6), 25, 2), map_layout='swapped'), out)
    write_pair('未見7_全部入り', kw(build(s(7), 35, 3), title='令和8年度 支出の明細', memo=True, r0=2, c0=3,
                                blanks=(4, 19), total=True, map_layout='swapped', spaces=True), out)
    write_pair('未見8_大量', kw(build(s(8), 1500, 25), map_extra=True), out)
    return out


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    write_pair('振り直し_本番', build(17, 30, 2))
    write_pair('振り直し_試験A', build(29, 45, 3))
    write_pair('振り直し_試験B', kw(build(41, 18, 0, zero_kubun=True), title='令和7年度 支出明細（総務課）'))
    write_pair('振り直し_試験C', kw(build(53, 24, 1), memo=True))
    write_pair('振り直し_試験D', kw(build(67, 26, 2), r0=3, c0=2, blanks=(7, 15)))
    write_pair('振り直し_試験E', kw(build(79, 22, 1), total=True))
    write_pair('振り直し_試験F', kw(build(83, 28, 2), map_layout='swapped'))
    write_pair('振り直し_試験G', kw(build(97, 25, 2), spaces=True))
    write_pair('振り直し_試験I', kw(build(101, 20, 1, kubun_set=KUBUN_SETS[1]), map_extra=True))
    shutil.copyfile(os.path.join(HERE, '振り直し_本番_正解.xlsx'), os.path.join(HERE, '振り直し_試験H_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '振り直し_本番_正解.xlsx'), os.path.join(HERE, '振り直し_試験H_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
