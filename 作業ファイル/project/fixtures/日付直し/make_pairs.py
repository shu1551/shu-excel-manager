# -*- coding: utf-8 -*-
"""選んでいる列の文字で入っている日付を、日付の値に直す（A 掃除 1・2026-09-18）。

列の決まり: 選んでいる 1 列＝日付の列。選んでいる列は <名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 「2026/04/15」「2026年4月15日」のように文字で入っている日付を、日付の値（シリアル値）に直す。
  - 表示形式は yyyy/m/d。もう日付の値になっているセルは触らない。空のセルも触らない。
  - 日付として読めない文字はそのまま残す。ほかの列は触らない。2 回撃っても結果が変わらない。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 空行／E 列の並び違い／F 撃った後にもう一度／G 番号の列／H 全部が文字
"""
import datetime
import os
import random
import shutil
import sys

from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('選んでいる列に文字で入っている日付を、日付の値に直してください（上書きしてよい）。'
           '「2026/04/15」「2026年4月15日」のような文字を日付の値にして、表示形式は yyyy/m/d にする。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           'もう日付の値になっているセルは触らない。空のセルも触らない。日付として読めない文字はそのまま残す。'
           'ほかの列は触らない。2 回撃っても結果が変わらない。')
PHRASES = ['選んだ列の文字の日付を日付に直して', '日付が文字列になっているので直して', '受付日の書き方をそろえて日付にして',
           '選んでいる列を日付の値にして', '文字で入っている日付を日付型に', '日付列の文字列を日付に変換して']


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0, h, last = lay['c0'], lay['h'], lay['last']
        dc = c0 + detail.col_of(t, 'date') - 1
        for r in range(h + 1, last + 1):
            v = ws.Cells(r, dc).Value
            if v is None or not isinstance(v, str) or not str(v).strip():
                continue
            d = detail.date_of(v)
            if d is None:
                continue
            ws.Cells(r, dc).NumberFormat = 'yyyy/m/d'
            # datetime を COM で渡すと時差で前日の 15 時になる＝シリアル値で書く
            ws.Cells(r, dc).Value2 = (d - datetime.datetime(1899, 12, 30)).days
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
            sel = f"{get_column_letter(detail.col_of(t, 'date'))}{lay['h']}"
            build_truth(xl, before, truth, t, lay)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def d(seed, n, **kw):
    kw.setdefault('text_date', True)
    return detail.design(seed, n, **kw)


def all_text(seed, n, **kw):
    """日付の列を全部文字にする（/・年月日・- の揺れ）＋読めない文字（未定・不明）を 2 つ。"""
    t = detail.design(seed, n, **kw)
    name = detail.name_of(t, 'date')
    rng = random.Random(seed + 41)
    for i, row in enumerate(t['rows']):
        v = row[name]
        if isinstance(v, str):
            continue
        k = i % 3
        row[name] = (f"{v.year}/{v.month:02d}/{v.day:02d}" if k == 0 else
                     f"{v.year}年{v.month}月{v.day}日" if k == 1 else f"{v.year}-{v.month}-{v.day}")
    for i in rng.sample(range(len(t['rows'])), 2):
        t['rows'][i][name] = rng.choice(['未定', '不明'])
    return t


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', d(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計行', d(s(2), 30, title=True, unit=True, total=True)),
          ('未見3_並び違い', d(s(3), 25, swap=True)),
          ('未見4_空行と番号列', d(s(4), 28, blanks=2, no_col=True))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('日付_本番', d(1401, 30)),
          ('日付_試験A', d(1402, 800, n_items=10)),
          ('日付_試験B', d(1403, 25, title=True, unit=True)),
          ('日付_試験C', d(1404, 24, total=True)),
          ('日付_試験D', d(1405, 26, blanks=2)),
          ('日付_試験E', d(1406, 28, swap=True)),
          ('日付_試験G', d(1407, 22, no_col=True)),
          ('日付_試験H', all_text(1408, 20)),
          ('日付_試験I', d(1409, 22, no_num=True))], HERE)   # 数の無い表（受付日だけの一覧）
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '日付_本番_正解.xlsx'), os.path.join(HERE, f'日付_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '日付_本番_選択.txt'), os.path.join(HERE, '日付_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
