# -*- coding: utf-8 -*-
"""アクティブなシートの表の見出しの行に、オートフィルタ（絞り込みのボタン）を付ける（2026-09-18 第二期）。

列の決まり: 選ぶ列は無い（表全体）。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝空でないセルが 2 つ以上並ぶ最初の行。表の列＝見出しの行で左端から右へ空に当たる手前まで。
  - フィルタの範囲＝見出しの行から、表の最後の行まで（下の合計の行も入る＝Excel も自動でそこまで広げる）。途中の空行は越えて入れる。
  - 前のフィルタ（ほかの範囲・絞り込み中）は外してから付け直す。値は変えない。
組: 本番／A 800 行／B 表題と単位／C 合計の行／D 途中の空行／E 番号の列／F 撃った後にもう一度／G 別の範囲に前のフィルタ／H 表が B 列から
"""
import os
import random
import shutil
import sys

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import detail   # noqa: E402

REQUEST = ('アクティブなシートの表の見出しの行に、オートフィルタ（絞り込みのボタン）を付けてください。'
           '見出しの行＝空でないセルが 2 つ以上並ぶ最初の行（文字でも数でもよい）。表の列＝見出しの行で、最初に値のある列から右へ見て空のセルに当たる手前の列まで。'
           'フィルタの範囲＝見出しの行から、表の最後の行まで（表の列のどれかに値がある最後の行。下の合計の行も入る）。途中の空行は越えて入れる。'
           'シートに前のフィルタがあれば（ほかの範囲・絞り込み中でも）外してから付け直す（ws.AutoFilterMode = False → Range.AutoFilter）。値は変えない。')
PHRASES = ['見出しにフィルタを付けて', 'オートフィルタを設定して', '絞り込みのボタンを付けて',
           'この表でフィルタを使えるようにして', 'フィルタをかけられるようにして', '見出しの行にオートフィルタを']


def shift_right(path):
    wb = load_workbook(path)
    wb.worksheets[0].insert_cols(1)
    wb.save(path)


def build(xl, before, truth, t, lay, old, shift):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(t['sheet'])
        c0 = 2 if shift else 1
        h, nc = lay['h'], lay['ncols']
        if old:
            ws.Range(ws.Cells(h, c0), ws.Cells(h + 3, c0 + 1)).AutoFilter(1, '<>')    # 前のフィルタ（狭い範囲・絞り込み中）
            wb.Save()
        last = max(int(ws.Cells(ws.Rows.Count, c).End(-4162).Row) for c in range(c0, c0 + nc))
        ws.AutoFilterMode = False
        ws.Range(ws.Cells(h, c0), ws.Cells(last, c0 + nc - 1)).AutoFilter(1)      # 引数なしは pywin32 で失敗する＝1 列目に条件なし
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, old, shift in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            lay = detail.write_before(before, t)
            if shift:
                shift_right(before)
            build(xl, before, truth, t, lay, old, shift)
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']], '前のフィルタ' if old else '', 'B 列から' if shift else '')
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', detail.design(s(1), rng.randint(20, 60)), False, False),
          ('未見2_表題と合計行', detail.design(s(2), 30, title=True, unit=True, total=True), True, False),
          ('未見3_番号と空行', detail.design(s(3), 25, no_col=True, blanks=2), False, False),
          ('未見4_B列から', detail.design(s(4), 28), False, True)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('フィルタ_本番', detail.design(4401, 30), False, False),
          ('フィルタ_試験A', detail.design(4402, 800, n_items=10), False, False),
          ('フィルタ_試験B', detail.design(4403, 25, title=True, unit=True), False, False),
          ('フィルタ_試験C', detail.design(4404, 24, total=True), False, False),
          ('フィルタ_試験D', detail.design(4405, 26, blanks=2), False, False),
          ('フィルタ_試験E', detail.design(4406, 28, no_col=True), False, False),
          ('フィルタ_試験G', detail.design(4407, 22), True, False),
          ('フィルタ_試験H', detail.design(4408, 20), False, True)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, 'フィルタ_本番_正解.xlsx'), os.path.join(HERE, f'フィルタ_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
