# -*- coding: utf-8 -*-
"""選んでいる列の重複しているキーを右に一覧にする（C 突合 24・2026-09-18）。

列の決まり: 選んでいる 1 列＝キーの列。選んでいる列は <名前>_選択.txt に書く。
正解の決まり（依頼文に全部書く）:
  - 置き場所＝表の右に 1 列空けた所。見出しの行は選んでいるセルの行。
  - 見出しは［キーの列の見出し］「件数」「行」。
  - 2 回以上出るキーだけを、初めて出た順に 1 行ずつ。件数＝出た数。行＝出た行番号を「, 」でつないだ文字。
  - キーが空の行は数えない。既に同じ表が右にあれば作り直す。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 空行／E 列の並び違い／F 撃った後にもう一度／G キーが数値／H 重複なし
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

REQUEST = ('選んでいる列（キーの列）で、2 回以上出てくる値を右に一覧にしてください。'
           '置き場所は表の右に 1 列空けた所。見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row を'
           'そのまま使う（自分で見出しの行を探さない）。見出しはキーの列の見出し・「件数」・「行」の 3 つ。'
           '2 回以上出る値だけを、初めて出た順に 1 行ずつ並べる。件数は出た数、行は出た行番号を「, 」（読点と半角空白）で'
           'つないだ文字。キーが空の行は数えない。本文は見出しの次の行から最後の本文の行まで'
           '（表の下に合計（合計・計）の行があればそれは入れない）。既に同じ表が右にあれば作り直す。')
PHRASES = ['選んだ列で重複している番号の一覧を右に', 'ダブっている伝票番号を洗い出して', '重複キーの一覧を作って',
           '同じ番号が何件あるか右に出して', '選んでいる列の重複を一覧にして', '重複している値と行番号を出して']


def dup_keys(t, seed):
    """キーの列を作り直して重複を仕込む（純 Python）。"""
    rng = random.Random(seed + 91)
    name = detail.name_of(t, 'no')
    n = len(t['rows'])
    base = [1000 + i for i in range(max(3, n - n // 4))]
    keys = [rng.choice(base) for _ in range(n)]
    for i, row in enumerate(t['rows']):
        row[name] = keys[i] if t.get('key_num') else f"{keys[i]:06d}"
    return t


def build_truth(xl, before, truth, t, lay):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0, h, last = lay['c0'], lay['h'], lay['last']
        kc = c0 + detail.col_of(t, 'no') - 1
        order, rows_of = [], {}
        for r in range(h + 1, last + 1):
            v = ws.Cells(r, kc).Value
            if v is None or str(v).strip() == '':
                continue
            k = str(v).strip()
            if k not in rows_of:
                order.append(k)
                rows_of[k] = []
            rows_of[k].append(r)
        out = c0 + lay['ncols'] + 1
        ws.Cells(h, out).Value = detail.name_of(t, 'no')
        ws.Cells(h, out + 1).Value = '件数'
        ws.Cells(h, out + 2).Value = '行'
        k = 0
        for key in order:
            rs = rows_of[key]
            if len(rs) < 2:
                continue
            k += 1
            cell = ws.Cells(h + k, out)
            cell.NumberFormat = ws.Cells(h + 1, kc).NumberFormat
            cell.Value = ws.Cells(rs[0], kc).Value
            ws.Cells(h + k, out + 1).Value = len(rs)
            ws.Cells(h + k, out + 2).Value = ", ".join(str(x) for x in rs)
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
            sel = f"{get_column_letter(detail.col_of(t, 'no'))}{lay['h']}"
            build_truth(xl, before, truth, t, lay)
            open(os.path.join(out, f"{stem}_選択.txt"), 'w', encoding='utf-8').write(sel + '\n')
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '選ぶ列', sel)
    finally:
        xl = None
        vf._quit(pid)


def one(seed, n, **kw):
    key_num = kw.pop('key_num', False)
    t = detail.design(seed, n, no_col=True, **kw)
    t['key_num'] = key_num
    return dup_keys(t, seed)


def nodup(seed, n):
    t = detail.design(seed, n, no_col=True)
    t['key_num'] = False
    name = detail.name_of(t, 'no')
    for i, row in enumerate(t['rows']):
        row[name] = f"{2000 + i:06d}"
    return t


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', one(s(1), rng.randint(20, 60))),
          ('未見2_表題と合計行', one(s(2), 30, title=True, unit=True, total=True)),
          ('未見3_並び違いと数値キー', one(s(3), 25, swap=True, key_num=True)),
          ('未見4_空行', one(s(4), 28, blanks=2))], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('重複_本番', one(901, 30)),
          ('重複_試験A', one(902, 800, n_items=10)),
          ('重複_試験B', one(903, 25, title=True, unit=True)),
          ('重複_試験C', one(904, 24, total=True)),
          ('重複_試験D', one(905, 26, blanks=2)),
          ('重複_試験E', one(906, 28, swap=True)),
          ('重複_試験G', one(907, 22, key_num=True)),
          ('重複_試験H', nodup(908, 12)),
          ('重複_試験I', one(909, 24, no_date=True))], HERE)   # 日付の無い表（2026-09-18 夜）
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '重複_本番_正解.xlsx'), os.path.join(HERE, f'重複_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '重複_本番_選択.txt'), os.path.join(HERE, '重複_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
