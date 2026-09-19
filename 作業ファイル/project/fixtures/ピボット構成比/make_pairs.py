# -*- coding: utf-8 -*-
"""科目別の支出額と構成比のピボット（鍛える回路の題材・ピボット 4 本目・値の表示方法と並べ替え・2026-09-18）。

仕事: 支出明細から、新しいシート「科目別構成比」の A3 にピボット（行＝科目、値＝支出額の合計 #,##0 と、支出額の列集計に対する比率 0.0%）。
正解の決まり（依頼文に全部書く）:
  - 科目は支出額の合計の大きい順。総計あり。
  - 元の範囲は明細の見出しから最後の明細の行まで（下の合計の行は入れない）。見出しの語で列を探す。
  - 「科目別構成比」シートが既にあれば消して作り直す。
組: 本番 60 行／A 1,500 行／B 表題・見出しが 4 行目／C 前の結果のシート／D 明細の下に合計の行／E 列の並びが違う／F 撃った後にもう一度
"""
import importlib.util
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('mp_pivot', os.path.join(HERE, '..', 'ピボット', 'make_pairs.py'))
P = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(P)
SHEET = P.SHEET
OUT_SHEET = '科目別構成比'
REQUEST = ('支出明細から、新しいシート「科目別構成比」の A3 にピボットテーブルを作ってください。行は科目、値は支出額の合計（表示形式 #,##0）と、'
           '支出額の列集計に対する比率（表示形式 0.0%）の 2 つ。科目は支出額の合計の大きい順に並べ、総計あり。'
           '元の範囲は明細の見出しから最後の明細の行まで（下の合計の行は入れない）。シートが既にあれば消して作り直す')


def build_pivot(wb, lay, stale=False):
    xl = wb.Application
    for sh in list(wb.Worksheets):
        if str(sh.Name) == OUT_SHEET:
            xl.DisplayAlerts = False
            sh.Delete()
    src = wb.Sheets(SHEET)
    rng = src.Range(src.Cells(lay['h'], 1), src.Cells(lay['last'], lay['ncols']))
    out = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    out.Name = OUT_SHEET
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{SHEET}'!{rng.GetAddress(True, True, -4150)}")
    pt = pc.CreatePivotTable(TableDestination=f"'{OUT_SHEET}'!R3C1", TableName='P科目別構成比')
    pt.PivotFields('科目').Orientation = 1
    d1 = pt.AddDataField(pt.PivotFields('支出額'), '合計 / 支出額', -4157)
    d1.NumberFormat = '#,##0'
    if stale:                                            # 前の結果（比率なし・科目の名前順）
        return pt
    d2 = pt.AddDataField(pt.PivotFields('支出額'), '構成比', -4157)
    d2.Calculation = 7
    d2.NumberFormat = '0.0%'
    pt.PivotFields('科目').AutoSort(2, '合計 / 支出額')
    return pt


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, rows, kw in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            stale = kw.pop('stale', False)
            lay = P.write_before(before, rows, **kw)
            if stale:
                wb = xl.Workbooks.Open(before)
                try:
                    build_pivot(wb, lay, stale=True)
                    wb.Worksheets(SHEET).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                build_pivot(wb, lay)
                wb.Worksheets(SHEET).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    import random
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', P.build(s(1), rng.randint(30, 120), n_ka=rng.randint(3, 8)), {}),
          ('未見2_表題と合計行', P.build(s(2), 80), {'title': True, 'total': True}),
          ('未見3_並び違いと前の結果', P.build(s(3), 50), {'cols': P.SWAP, 'stale': True}),
          ('未見4_大量', P.build(s(4), 3000, n_ka=10), {})], out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    make([('構成比P_本番', P.build(3, 60), {}),
          ('構成比P_試験A', P.build(13, 1500, n_ka=10), {}),
          ('構成比P_試験B', P.build(23, 50), {'title': True}),
          ('構成比P_試験C', P.build(33, 40), {'stale': True}),
          ('構成比P_試験D', P.build(43, 45), {'total': True}),
          ('構成比P_試験E', P.build(53, 55), {'cols': P.SWAP})], HERE)
    shutil.copyfile(os.path.join(HERE, '構成比P_本番_正解.xlsx'), os.path.join(HERE, '構成比P_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '構成比P_本番_正解.xlsx'), os.path.join(HERE, '構成比P_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    lines = ['科目別の支出額と構成比をピボットで', '科目ごとの割合が分かるピボットを大きい順に', '支出明細を科目別に集計して構成比も',
             '科目別構成比のピボットテーブルを作って', 'ピボットで科目の合計と比率を並べて', '科目の支出の多い順に構成比つきで']
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('ok')
