# -*- coding: utf-8 -*-
"""前年度比較（鍛える回路の題材・2026-09-17 夜。予算要求・決算の説明資料で毎年作る表）。

仕事: 「今年度」と「前年度」のシート（科目・予算額）を科目で突き合わせ、今年度の表の右に 1 列空けて比較表を作る。
正解の決まり:
  - 比較表の見出しは 1 行: 科目・今年度・前年度・増減額・増減率。左上＝今年度の見出しの行・表の右端の 2 つ右。
  - 行の順: 今年度の表の順 → そのあと前年度にだけある科目（前年度の表の順）。
  - 片方に無い科目の額は 0。増減額＝今年度−前年度。
  - 増減率＝増減額÷前年度（丸めない）。前年度が 0 なら文字「皆増」、今年度が 0（前年度は 0 でない）なら文字「皆減」。
  - 最後に「合計」の行（今年度・前年度・増減額の合計、増減率も同じ決まり）。
  - 科目の前後の空白（半角・全角）は無視して同じ科目にする。科目が空の行・「合計」「計」の行は入れない。
  - 金額はカンマ・全角・前後の空白・「円」があっても数で読む。見出しの語（科目・予算額）で列を探す。
  - 比較表の場所に前の比較表が残っていたら消してから書く（2 回撃っても同じ）。
組（依頼文は同じ）: 本番 12 科目／A 30 科目／B 新規の科目（前年度に無い）／C 廃止の科目（今年度に無い）／D 金額が文字／
  E 今年度に表題・B3 から／F 両方に合計行／G 科目の空白の揺れ／H 前の比較表が残っている／I 撃った後にもう一度／
  J シート名が「令和8年度」「令和7年度」（年度の大きい方が今年度）／K 前の比較表が今回より長い
  py make_pairs.py            → *_前.xlsx と *_正解.xlsx
  py make_pairs.py --unseen N → 未見_種N
"""
import os
import random
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
KAMOKU = ['報酬', '給料', '職員手当等', '共済費', '報償費', '旅費', '交際費', '需用費', '役務費', '委託料', '使用料及び賃借料',
          '工事請負費', '原材料費', '公有財産購入費', '備品購入費', '負担金補助及び交付金', '扶助費', '貸付金', '補償補填及び賠償金',
          '償還金利子及び割引料', '投資及び出資金', '積立金', '寄附金', '公課費', '繰出金', '予備費']
_ZEN = str.maketrans('0123456789,', '０１２３４５６７８９，')
HEAD = ['科目', '今年度', '前年度', '増減額', '増減率']


def build(seed, n, new=0, gone=0):
    rng = random.Random(seed)
    names = rng.sample(KAMOKU, n) if n <= len(KAMOKU) else [f"{k}{i}" for i, k in enumerate(rng.choices(KAMOKU, k=n))]
    this = [(k, rng.randrange(1, 900) * 1000) for k in names]
    prev = [(k, max(0, a + rng.randrange(-200, 200) * 1000) or 1000) for k, a in this]
    rng.shuffle(prev)
    if new:
        drop = {k for k, _a in rng.sample(this, new)}
        prev = [(k, a) for k, a in prev if k not in drop]           # 今年度の新規＝前年度に無い
    if gone:
        extra = [f"{k}（旧）" for k in rng.sample(KAMOKU, gone)]
        prev += [(k, rng.randrange(1, 500) * 1000) for k in extra]  # 前年度にだけある（廃止）
    return this, prev


def compare(this, prev):
    pm = {k: a for k, a in prev}
    tm = {k: a for k, a in this}
    rows = [(k, a, pm.get(k, 0)) for k, a in this] + [(k, 0, a) for k, a in prev if k not in tm]
    out = []
    for k, t, p in rows + [('合計', sum(r[1] for r in rows), sum(r[2] for r in rows))]:
        d = t - p
        rate = '皆増' if p == 0 else ('皆減' if t == 0 else d / p)
        out.append([k, t, p, d, rate])
    return out


def write_pair(stem, this, prev, out=HERE, title=False, r0=1, c0=1, text_amount=False, total=False, spaces=False,
               stale=0, sheets=('今年度', '前年度')):
    rng = random.Random(stem)
    for truth in (False, True):
        wb = Workbook()
        ws = wb.active
        ws.title = sheets[0]
        h = r0
        if title:
            ws.cell(r0, c0, '令和8年度 当初予算（一般会計・歳出）').font = Font(bold=True, size=14)
            h = r0 + 2

        def put(sh, hr, cc, items, key):
            r3 = random.Random(key)
            sh.cell(hr, cc, '科目').font = Font(bold=True)
            sh.cell(hr, cc + 1, '予算額').font = Font(bold=True)
            for i, (k, a) in enumerate(items, 1):
                name = k
                if spaces and r3.random() < 0.3:
                    name = r3.choice((' ', '　')) + k if r3.random() < 0.5 else k + r3.choice((' ', '　'))
                sh.cell(hr + i, cc, name)
                v = a
                if text_amount and r3.random() < 0.5:
                    x = r3.random()
                    v = f"{a:,}".translate(_ZEN) if x < 0.3 else (f" {a:,} " if x < 0.6 else f"{a:,}円")
                sh.cell(hr + i, cc + 1, v).number_format = '#,##0'
            last = hr + len(items)
            if total:
                sh.cell(last + 1, cc, '合計').font = Font(bold=True)
                sh.cell(last + 1, cc + 1, sum(a for _k, a in items)).number_format = '#,##0'
        put(ws, h, c0, this, stem + 't')
        lc = c0 + 3
        if truth:
            for j, v in enumerate(HEAD):
                ws.cell(h, lc + j, v)
            for i, row in enumerate(compare(this, prev), 1):
                for j, v in enumerate(row):
                    ws.cell(h + i, lc + j, v)
        elif stale:
            for j, v in enumerate(HEAD):
                ws.cell(h, lc + j, v)
            for i in range(1, stale + 1):
                for j, v in enumerate([f'旧科目{i}', 1000 * i, 900 * i, 100 * i, 0.1]):
                    ws.cell(h + i, lc + j, v)
        pv = wb.create_sheet(sheets[1])
        put(pv, 1, 1, prev, stem + 'p')
        wb.save(os.path.join(out, f"{stem}_{'正解' if truth else '前'}.xlsx"))


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    write_pair('未見1_標準', *build(s(1), rng.randint(6, 20)), out=out)
    write_pair('未見2_新規と廃止', *build(s(2), 14, new=3, gone=2), out=out)
    write_pair('未見3_文字の金額と空白', *build(s(3), 12, new=1), text_amount=True, spaces=True, out=out)
    write_pair('未見4_表題と位置と合計行', *build(s(4), 10, gone=1), title=True, r0=2, c0=2, total=True, out=out)
    write_pair('未見5_前の比較表', *build(s(5), 8), stale=15, out=out)
    write_pair('未見6_全部入り', *build(s(6), 18, new=2, gone=3), title=True, r0=3, c0=2, text_amount=True, total=True,
               spaces=True, stale=25, out=out)
    write_pair('未見7_シート名', *build(s(7), 9, new=1), sheets=('R8当初', 'R7当初'), out=out)
    write_pair('未見8_大量', *build(s(8), 400, new=20, gone=15), out=out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    write_pair('比較_本番', *build(3, 12))
    write_pair('比較_試験A', *build(13, 25))
    write_pair('比較_試験B', *build(23, 12, new=3))
    write_pair('比較_試験C', *build(33, 10, gone=2))
    write_pair('比較_試験D', *build(43, 11), text_amount=True)
    write_pair('比較_試験E', *build(53, 9, new=1), title=True, r0=3, c0=2)
    write_pair('比較_試験F', *build(63, 10, gone=1), total=True)
    write_pair('比較_試験G', *build(73, 12), spaces=True)
    write_pair('比較_試験H', *build(83, 8), stale=12)
    write_pair('比較_試験J', *build(93, 9, new=1), sheets=('令和8年度', '令和7年度'))   # 年度の大きい方が今年度
    write_pair('比較_試験K', *build(103, 6, gone=1), stale=20)                         # 前の比較表が今回より 13 行多い
    shutil.copyfile(os.path.join(HERE, '比較_本番_正解.xlsx'), os.path.join(HERE, '比較_試験I_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '比較_本番_正解.xlsx'), os.path.join(HERE, '比較_試験I_正解.xlsx'))
    print('ok')
