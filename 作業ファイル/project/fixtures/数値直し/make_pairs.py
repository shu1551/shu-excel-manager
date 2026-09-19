# -*- coding: utf-8 -*-
"""選んでいる列に文字で入っている数字を、数値に直す（A 掃除 2・2026-09-18）。

列の決まり: 選んでいる 1 列＝数の列。選んでいる列は <名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 「１２，３００」「 12,300 」「12,300円」「△5,000」のように文字で入っている数を数値に直す（△・▲・- はマイナス）。
  - 表示形式は #,##0。もう数値のセル・空のセルは触らない。
  - 0 で始まる 2 桁以上の数字だけの文字（0120 など）は番号として残す。数として読めない文字もそのまま残す。
  - ほかの列は触らない。2 回撃っても結果が変わらない。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 空行／E 列の並び違い／F 撃った後にもう一度／G 番号の列／H 全部が文字（△・円・0 始まり・読めない文字）
"""
import os
import random
import re
import shutil
import sys

from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('選んでいる列に文字で入っている数字を、数値に直してください（上書きしてよい）。'
           '「１２，３００」「 12,300 」「12,300円」「△5,000」のような文字を数値にする。全角の数字・全角のカンマ・前後の空白・'
           '円は取り除き、頭の △・▲・- はマイナスにする。表示形式は #,##0 にする。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           'もう数値になっているセルは触らない。空のセルも触らない。0 で始まる 2 桁以上の数字だけの文字（0120 など）は番号なのでそのまま残す。'
           '数として読めない文字もそのまま残す。ほかの列は触らない。2 回撃っても結果が変わらない。')
PHRASES = ['選んだ列の文字の数字を数値に直して', '金額が文字列になっているので数値にして', '全角で入った金額を数値に変換して',
           '選んでいる列を計算できる数値にして', '文字で入っている金額を数に直して', 'カンマ付きの文字を数値に変えて']

_ZEN = str.maketrans('０１２３４５６７８９，．－', '0123456789,.-')


def num_or_none(v):
    """依頼文の決まりで文字を数にする（純 Python）。数にしないものは None。"""
    s = str(v).replace('　', ' ').strip().translate(_ZEN)
    neg = False
    if s[:1] in ('△', '▲', '-'):
        neg, s = True, s[1:].strip()
    s = s.replace(',', '').replace('円', '').strip()
    if not re.fullmatch(r'\d+(\.\d+)?', s):
        return None
    if len(s) >= 2 and s[0] == '0' and '.' not in s and not neg:
        return None                                  # 0120 などの番号
    x = float(s) if '.' in s else int(s)
    return -x if neg else x


def yen(seed, n, **kw):
    """文字の金額に「円」をつけたものも混ぜる（合計の行を持つ表では混ぜない＝detail の合計が読めない）。"""
    kw.setdefault('text_amount', True)
    t = detail.design(seed, n, **kw)
    if t['total']:
        return t
    name = detail.name_of(t, 'num')
    rng = random.Random(seed + 51)
    for row in t['rows']:
        if isinstance(row[name], str) and rng.random() < 0.3:
            row[name] = row[name].strip() + '円'
    return t


def all_text(seed, n, **kw):
    """数の列を全部文字にする（全角・空白・円・△）＋0 始まりの番号と読めない文字を 1 つずつ。"""
    t = detail.design(seed, n, **kw)
    name = detail.name_of(t, 'num')
    rng = random.Random(seed + 61)
    zen = str.maketrans('0123456789,', '０１２３４５６７８９，')
    for i, row in enumerate(t['rows']):
        v = row[name]
        if isinstance(v, str):
            continue
        k = i % 5
        row[name] = (f"{v:,}".translate(zen) if k == 0 else f" {v:,} " if k == 1 else f"{v:,}円" if k == 2
                     else f"△{v:,}" if k == 3 else f"▲{v}")
    a, b = rng.sample(range(len(t['rows'])), 2)
    t['rows'][a][name] = '0120'
    t['rows'][b][name] = rng.choice(['未定', '－'])
    return t


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0, h, last = lay['c0'], lay['h'], lay['last']
        nc = c0 + detail.col_of(t, 'num') - 1
        for r in range(h + 1, last + 1):
            v = ws.Cells(r, nc).Value
            if v is None or not isinstance(v, str) or not v.strip():
                continue
            x = num_or_none(v)
            if x is None:
                continue
            ws.Cells(r, nc).NumberFormat = '#,##0'
            ws.Cells(r, nc).Value2 = x
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
            sel = f"{get_column_letter(detail.col_of(t, 'num'))}{lay['h']}"
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
    make([('未見1_標準', yen(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計行', yen(s(2), 30, title=True, unit=True, total=True)),
          ('未見3_並び違い', yen(s(3), 25, swap=True)),
          ('未見4_全部が文字と番号列', all_text(s(4), 28, blanks=2, no_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('数値_本番', yen(1501, 30)),
          ('数値_試験A', yen(1502, 800, n_items=10)),
          ('数値_試験B', yen(1503, 25, title=True, unit=True)),
          ('数値_試験C', yen(1504, 24, total=True)),
          ('数値_試験D', yen(1505, 26, blanks=2)),
          ('数値_試験E', yen(1506, 28, swap=True)),
          ('数値_試験G', yen(1507, 22, no_col=True)),
          ('数値_試験H', all_text(1508, 20)),
          ('数値_試験I', yen(1509, 24, no_date=True))], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '数値_本番_正解.xlsx'), os.path.join(HERE, f'数値_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '数値_本番_選択.txt'), os.path.join(HERE, '数値_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
