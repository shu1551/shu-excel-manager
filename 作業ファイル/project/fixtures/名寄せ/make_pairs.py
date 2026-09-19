# -*- coding: utf-8 -*-
"""選んでいる列で、書き方の違う同じ名前（株式会社の有無・空白・全角半角）の組を、表の右に一覧にする（C 突合 27・2026-09-18 第二期）。

列の決まり: 選んでいる 1 列＝名前の列（相手先など）。<名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 比べるための形＝「株式会社」「有限会社」「(株)」「（株）」「㈱」「(有)」「（有）」「㈲」を取り、半角・全角の空白を全部取り、StrConv vbNarrow で半角にした文字。
  - 比べる形が同じで、元の書き方が 2 通り以上ある組を、表に初めて出た順に 1 行ずつ。
  - 表の右に 1 列空けて「まとめる名前」（組で最初に出た書き方）「書き方の数」「書き方」（出た順に「／」でつなぐ）。無ければ見出しの行だけ。
  - 本文＝見出しの次の行から最後の行まで（空の行と「合計」「計」「総計」の行は見ない）。表の値は触らない。既に一覧が右にあれば消して作り直す。
組: 本番／A 400 行／B 表題と単位／C 合計の行／D 全角の英字（ＡＢＣ）／E 列の並び違い／F 撃った後にもう一度／G 空白だけの違い／H 書き方の揺れなし
"""
import datetime
import os
import random
import shutil
import sys
import unicodedata

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))

REQUEST = ('選んでいる列で、字面が違うだけの同じ名前（株式会社の有無・空白・全角と半角）の組を、表の右に 1 列空けて一覧にしてください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '表の列＝見出しの行で、表のいちばん左の列から右へ見て空のセルに当たる手前の列まで。本文は見出しの次の行から最後の行まで'
           '（行がまるごと空の行と、行のどこかに「合計」「計」「総計」とある行は見ない）。'
           '比べるための形＝値から「株式会社」「有限会社」「(株)」「（株）」「㈱」「(有)」「（有）」「㈲」を取り除き、半角と全角の空白を全部取り除いてから、'
           'StrConv(…, vbNarrow) で半角にした文字。比べる形が同じで、元の文字が 2 通り以上ある組を、表に初めて出た順に 1 行ずつ書く。'
           '表の右端の 2 つ右の列から、見出しの行に「まとめる名前」「種類の数」「元の名前」と書き、その下に: まとめる名前＝その組で表に最初に出た名前・'
           '種類の数＝違う名前の数・元の名前＝違う名前を表に出た順に「／」でつないだ文字。組が無ければ見出しの行だけ書く。'
           '表の値は触らない。既に一覧が右にあれば消して作り直す。')
PHRASES = ['選んだ列の名寄せの候補を出して', '株式会社の有無で分かれている相手先を探して', '名前の揺れを一覧に',
           '同じ会社なのに名前が違うものを洗い出して', '相手先の名寄せをしたい', '選んでいる列で同じ相手の別名を出して']
BASES = ['秋田商事', '大館建設', '横手運輸', '能代印刷', '湯沢電気', '男鹿水産', 'ABC設備', '北都リース']
FORMS = ['{b}', '株式会社{b}', '{b}株式会社', '(株){b}', '㈱{b}', '{b}(株)', '{b}（株）', '有限会社{b}', '(有){b}', '{b} ', '{b}　']
TOKENS = ['株式会社', '有限会社', '(株)', '（株）', '㈱', '(有)', '（有）', '㈲']


def key_of(v):
    s = str(v)
    for tk in TOKENS:
        s = s.replace(tk, '')
    s = s.replace(' ', '').replace('　', '')
    return unicodedata.normalize('NFKC', s)


def build(seed, n, **kw):
    rng = random.Random(seed)
    bases = rng.sample(BASES, 5)
    t = {'sheet': rng.choice(['支払一覧', '契約一覧', 'Sheet1', '相手先別']), 'name': rng.choice(['相手先', '支払先', '業者名', '契約相手']),
         'title': False, 'total': False, 'swap': False, 'zen': False, 'spaces_only': False, 'clean': False, 'no_date': False}
    t.update(kw)
    rows = []
    for i in range(n):
        b = rng.choice(bases)
        if t['zen'] and b == 'ABC設備':
            b = rng.choice(['ABC設備', 'ＡＢＣ設備'])
        if t['clean']:
            f = '{b}'
        elif t['spaces_only']:
            f = rng.choice(['{b}', '{b}', '{b} ', '{b}　'])
        else:
            f = rng.choice(FORMS[:3]) if rng.random() < 0.6 else rng.choice(FORMS)
        rows.append((1000 + i, datetime.datetime(2026, rng.randint(4, 12), rng.randint(1, 28)), f.format(b=b),
                     rng.randrange(10, 900) * 1000))
    if t['zen']:
        rows[0] = (rows[0][0], rows[0][1], 'ABC設備', rows[0][3])
        rows[1] = (rows[1][0], rows[1][1], 'ＡＢＣ設備', rows[1][3])
    t['rows'] = rows
    return t


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    h = 1
    if t['title']:
        ws['A1'] = '令和8年度 支払一覧'
        ws['A1'].font = Font(bold=True, size=14)
        ws['A2'] = '単位：円'
        h = 4
    heads = ['番号', '支払日', t['name'], '金額']
    order = [0, 3, 1, 2] if t['swap'] else [0, 1, 2, 3]
    if t['no_date']:
        order = [k for k in order if k != 1]                # 日付の無い表（2026-09-18 夜）
    for j, k in enumerate(order, 1):
        ws.cell(h, j, heads[k]).font = Font(bold=True)
    r = h
    for row in t['rows']:
        r += 1
        for j, k in enumerate(order, 1):
            c = ws.cell(r, j, row[k])
            if k == 1:
                c.number_format = 'yyyy/m/d'
            elif k == 3:
                c.number_format = '#,##0'
    if t['total']:
        ws.cell(r + 1, order.index(2) + 1, '合計')
        ws.cell(r + 1, order.index(3) + 1, sum(x[3] for x in t['rows'])).number_format = '#,##0'
    # 正解（組の一覧）
    groups, order_k = {}, []
    for row in t['rows']:
        k = key_of(row[2])
        if k not in groups:
            groups[k] = []
            order_k.append(k)
        if row[2] not in groups[k]:
            groups[k].append(row[2])
    wb.save(path)
    return h, order.index(2) + 1, len(order), [(groups[k][0], len(groups[k]), '／'.join(groups[k])) for k in order_k if len(groups[k]) >= 2]


def make(jobs, out):
    from openpyxl import load_workbook
    for stem, t in jobs:
        before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
        h, nc, ncols, groups = write_before(before, t)
        wb = load_workbook(before)
        ws = wb.worksheets[0]
        o = ncols + 2
        for j, nm in enumerate(('まとめる名前', '種類の数', '元の名前')):
            ws.cell(h, o + j, nm)
        for i, g in enumerate(groups, h + 1):
            for j, x in enumerate(g):
                ws.cell(i, o + j, x)
        wb.save(truth)
        open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(f"{get_column_letter(nc)}{h}\n")
        print('  ', stem, t['sheet'], t['name'], '組', len(groups))


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', build(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計', build(s(2), 30, title=True, total=True)),
          ('未見3_並び違いと全角', build(s(3), 25, swap=True, zen=True)),
          ('未見4_空白だけ', build(s(4), 28, spaces_only=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('名寄せ_本番', build(4801, 30)),
          ('名寄せ_試験A', build(4802, 400)),
          ('名寄せ_試験B', build(4803, 25, title=True)),
          ('名寄せ_試験C', build(4804, 24, total=True)),
          ('名寄せ_試験D', build(4805, 26, zen=True)),
          ('名寄せ_試験E', build(4806, 28, swap=True)),
          ('名寄せ_試験G', build(4807, 22, spaces_only=True)),
          ('名寄せ_試験H', build(4808, 20, clean=True)),
          ('名寄せ_試験I', build(4809, 24, no_date=True))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '名寄せ_本番_正解.xlsx'), os.path.join(HERE, f'名寄せ_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '名寄せ_本番_選択.txt'), os.path.join(HERE, '名寄せ_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
