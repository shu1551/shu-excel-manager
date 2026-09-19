# -*- coding: utf-8 -*-
"""アクティブなシートの表を、いつもの形で印刷できるように設定する（A9＋D33・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い（表全体）。
正解の決まり（依頼文に全部書く）:
  - 見出しの行＝空でないセルが 2 つ以上並ぶ最初の行。表の左端＝見出しの行の最初の値の列・右端＝そこから右へ空に当たる手前。
    表の最後の行＝表の列のどれかに値がある最後の行（下の合計の行も入れる）。
  - 用紙 A4。表の列が 6 以上なら横・5 以下なら縦。拡大縮小は「横 1 ページ・縦は何ページでも」（Zoom=False・FitToPagesWide=1・FitToPagesTall=False）。
  - 見出しの行を毎ページに（PrintTitleRows）。印刷範囲＝1 行目（表題も入れる）の表の左端〜表の最後の行の右端。
  - 水平の中央。フッターの中央に「&P / &N」。余白: 左右 0.25 インチ・上下 0.75 インチ。
組: 本番 5 列（縦）／A 800 行／B 表題と単位／C 合計の行／D 番号の列で 6 列（横）／E 集計表 2 列／F 撃った後にもう一度／
  G 7 列（番号・横）と空行／H 表が B 列から
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
import shukei   # noqa: E402

REQUEST = ('アクティブなシートの表を、いつもの形で印刷できるように印刷の設定をしてください。値と書式は変えない。'
           '見出しの行＝空でないセルが 2 つ以上並ぶ最初の行（文字でも数でもよい）。表の左端＝見出しの行で最初に値のある列、右端＝そこから右へ見て空のセルに当たる手前の列。'
           '表の最後の行＝表の列のどれかに値がある最後の行（下の合計の行も入れる）。'
           '用紙は A4（PaperSize = xlPaperA4）。表の列が 6 以上なら横（xlLandscape）、5 以下なら縦（xlPortrait）。'
           '拡大縮小は横 1 ページ・縦は何ページでも（Zoom = False、FitToPagesWide = 1、FitToPagesTall = False）。'
           '見出しの行を毎ページに印刷（PrintTitleRows ＝ 見出しの行）。印刷範囲（PrintArea）＝1 行目（表題も入れる）の表の左端の列から、表の最後の行の右端の列まで。'
           '水平の中央（CenterHorizontally = True）。フッターの中央に「&P / &N」。'
           '余白は左右 0.25 インチ・上下 0.75 インチ（Application.InchesToPoints）。')
PHRASES = ['印刷できるように設定して', 'A4で横1ページに収まるように印刷設定を', 'いつもの印刷の設定にして',
           '見出しを毎ページに出して印刷できるように', '横幅を1ページに収めて印刷', '印刷の体裁を整えてページ番号も']


def build_truth(xl, before, truth, sheet):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(sheet)
        ur = ws.UsedRange
        h = next(r for r in range(1, int(ur.Row) + int(ur.Rows.Count))
                 if sum(1 for c in range(1, int(ur.Column) + int(ur.Columns.Count)) if ws.Cells(r, c).Value is not None) >= 2)
        left = next(c for c in range(1, 60) if ws.Cells(h, c).Value is not None)
        right = left
        while ws.Cells(h, right + 1).Value is not None:
            right += 1
        last = max(int(ws.Cells(ws.Rows.Count, c).End(-4162).Row) for c in range(left, right + 1))
        ncol = right - left + 1
        ps = ws.PageSetup
        xl.PrintCommunication = False
        ps.PaperSize = 9
        ps.Orientation = 2 if ncol >= 6 else 1
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        ps.PrintTitleRows = f"${h}:${h}"
        ps.PrintArea = ws.Range(ws.Cells(1, left), ws.Cells(last, right)).Address
        ps.CenterHorizontally = True
        ps.CenterFooter = '&P / &N'
        ps.LeftMargin = xl.InchesToPoints(0.25)
        ps.RightMargin = xl.InchesToPoints(0.25)
        ps.TopMargin = xl.InchesToPoints(0.75)
        ps.BottomMargin = xl.InchesToPoints(0.75)
        xl.PrintCommunication = True
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def shift_right(path):
    """表を B 列から始まるように 1 列右へずらす（openpyxl）。"""
    wb = load_workbook(path)
    wb.worksheets[0].insert_cols(1)
    wb.save(path)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, kind, t, shift in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            (detail if kind == 'd' else shukei).write_before(before, t)
            if shift:
                shift_right(before)
            build_truth(xl, before, truth, t['sheet'])
            print('  ', stem, t['sheet'], [n for n, _r in t['cols']])
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', 'd', detail.design(s(1), rng.randint(20, 60)), False),
          ('未見2_表題と合計行', 'd', detail.design(s(2), 30, title=True, unit=True, total=True, no_col=True), False),
          ('未見3_集計表', 's', shukei.design(s(3), 8, nums=2, title=True), False),
          ('未見4_B列から', 'd', detail.design(s(4), 25, blanks=1), True)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('印刷_本番', 'd', detail.design(3301, 30), False),
          ('印刷_試験A', 'd', detail.design(3302, 800, n_items=10), False),
          ('印刷_試験B', 'd', detail.design(3303, 25, title=True, unit=True), False),
          ('印刷_試験C', 'd', detail.design(3304, 24, total=True), False),
          ('印刷_試験D', 'd', detail.design(3305, 26, no_col=True), False),
          ('印刷_試験E', 's', shukei.design(3306, 8), False),
          ('印刷_試験G', 's', shukei.design(3307, 9, nums=2, rate=True, text=True, no_col=True, code_col=True, blanks=1), False),
          ('印刷_試験H', 'd', detail.design(3308, 20), True)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '印刷_本番_正解.xlsx'), os.path.join(HERE, f'印刷_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
