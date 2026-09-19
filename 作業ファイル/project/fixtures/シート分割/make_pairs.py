# -*- coding: utf-8 -*-
"""選んでいる列の値ごとに、表をシートに分ける（C 突合 30・2026-09-18 第二期）。

列の決まり: 選んでいる 1 列＝分ける列。<名前>_選択.txt に書く。表は明細（detail）。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝選んでいるセルの行。本文＝見出しの次の行から最後の行まで（空の行と「合計」「計」「総計」の行は入れない）。
  - 分ける列の値ごとに（表に初めて出た順）、最後のシートの後ろへ新しいシートを足す。シートの名前＝その値（前後の空白は落とす）。
    同じ名前のシートが既にあれば消して作り直す。
  - 新しいシートの 1 行目＝見出しの行（表の左端〜右端）、2 行目から＝その値の行を上から順に（値と表示形式を写す）。
  - 元の表は変えない。
組: 本番 課／A 400 行・10 項目／B 表題と単位／C 合計の行／D 途中の空行／E 列の並び違い／F 撃った後にもう一度／G 番号の列・科目で分ける／H 項目 1 つ
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

REQUEST = ('選んでいる列の値ごとに、表をシートに分けてください。'
           '見出しの行は必ず「選んでいるセルの行」＝Selection.Cells(1,1).Row をそのまま使う（自分で見出しの行を探さない）。'
           '表の列＝見出しの行で、表のいちばん左の列から右へ見て空のセルに当たる手前の列まで。本文は見出しの次の行から最後の行まで。'
           '行がまるごと空の行と、行のどこかに「合計」「計」「総計」とある行は入れない。'
           '選んだ列の値ごとに（表に初めて出た順に）、最後のシートの後ろへ新しいシートを足し、シートの名前はその値にする（前後の空白は落とす）。'
           '同じ名前のシートが既にあれば、消してから作り直す。'
           '新しいシートの 1 行目に見出しの行（表の列の分）、2 行目からその値の行を表の上から順に写す（値と表示形式）。元の表は変えない。')
PHRASES = ['選んだ列の値ごとにシートを分けて', '課ごとにシートを分割して', 'この表を項目別のシートに振り分けて',
           '担当ごとに別々のシートへ分けて', 'シートを値ごとに分割', '選んでいる列でシート分けして']


def build_truth(xl, before, truth, t, lay, role):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        h, last, nc = lay['h'], lay['last'], lay['ncols']
        kc = detail.col_of(t, role)
        order, rows = [], {}
        for r in range(h + 1, last + 1):
            vals = [ws.Cells(r, c).Value for c in range(1, nc + 1)]
            if all(v is None for v in vals) or any(str(v).strip() in ('合計', '計', '総計') for v in vals if v is not None):
                continue
            k = str(vals[kc - 1]).strip()
            if k not in rows:
                order.append(k)
                rows[k] = []
            rows[k].append(r)
        for k in order:
            ns = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
            ns.Name = k
            ws.Range(ws.Cells(h, 1), ws.Cells(h, nc)).Copy(ns.Cells(1, 1))
            for i, r in enumerate(rows[k], 2):
                ws.Range(ws.Cells(r, 1), ws.Cells(r, nc)).Copy(ns.Cells(i, 1))
        ws.Activate()
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, role in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = detail.write_before(before, t)
            build_truth(xl, before, truth, t, lay, role)
            sel = f"{get_column_letter(detail.col_of(t, role))}{lay['h']}"
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
    make([('未見1_標準', detail.design(s(1), rng.randint(20, 50)), 'item'),
          ('未見2_表題と合計行', detail.design(s(2), 30, title=True, unit=True, total=True), 'item'),
          ('未見3_並び違いで科目', detail.design(s(3), 25, swap=True), 'kamoku'),
          ('未見4_空行と番号列', detail.design(s(4), 28, blanks=2, no_col=True, item_space=True), 'item')], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('分割_本番', detail.design(3501, 30), 'item'),
          ('分割_試験A', detail.design(3502, 400, n_items=10), 'item'),
          ('分割_試験B', detail.design(3503, 25, title=True, unit=True), 'item'),
          ('分割_試験C', detail.design(3504, 24, total=True), 'item'),
          ('分割_試験D', detail.design(3505, 26, blanks=2, item_space=True), 'item'),
          ('分割_試験E', detail.design(3506, 28, swap=True), 'item'),
          ('分割_試験G', detail.design(3507, 22, no_col=True), 'kamoku'),
          ('分割_試験H', detail.design(3508, 12, n_items=1), 'item'),
          ('分割_試験I', detail.design(3509, 24, no_date=True), 'item')], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '分割_本番_正解.xlsx'), os.path.join(HERE, f'分割_試験F_{x}.xlsx'))
    shutil.copyfile(os.path.join(HERE, '分割_本番_選択.txt'), os.path.join(HERE, '分割_試験F_選択.txt'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
