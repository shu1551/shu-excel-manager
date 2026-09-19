# -*- coding: utf-8 -*-
"""アクティブなシートのピボットに、累計の値のフィールドを足す（F ピボット 52・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い。アクティブなシートのピボット全部。
正解の決まり（依頼文に全部書く）:
  - ピボットごとに、1 つ目の値のフィールドと同じ元の列を、名前「累計」・合計・計算の種類＝累計（xlRunningTotal）・
    基準＝いちばん外側の行のフィールド・表示形式 #,##0 で足す。
  - 既に「累計」の値のフィールドがあれば足さない。ほかのシートのピボットは触らない。ピボットが無ければ何もしない。
組: 本番 行＝課／A 行＝科目／B ピボット 2 つ／C 列＝科目つき／D 表題の下（A3）／E 行 2 つ／F 撃った後にもう一度／G 値のフィールドが 2 つ／H ピボットなし
"""
import importlib.util
import os
import random
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('vp', os.path.join(HERE, '..', 'ピボット値貼り', 'make_pairs.py'))
pv = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pv)
import detail   # noqa: E402

REQUEST = ('アクティブなシートのピボットテーブルに、累計の値のフィールドを足してください。'
           'ピボットごとに、1 つ目の値のフィールド（DataFields(1)）と同じ元の列を、名前「累計」・集計は合計（xlSum）で AddDataField し、'
           '計算の種類を累計（.Calculation = xlRunningTotal）、基準のフィールド（.BaseField）をいちばん外側の行のフィールド、表示形式を #,##0 にする。'
           '既に「累計」という値のフィールドがあれば足さない。ほかのシートのピボットは触らない。ピボットが無ければ何もしない。')
PHRASES = ['ピボットに累計を足して', 'ピボットで累計も出して', '累計の列をピボットに追加',
           'ピボットテーブルに累計の値を入れて', 'ピボットの合計の横に累計を', '累計をピボットで表示して']


def add_count(xl, path, t):
    """値のフィールドを 2 つにする（合計の後ろに個数）。"""
    wb = xl.Workbooks.Open(path)
    try:
        pt = wb.Worksheets(1).PivotTables(1)
        pt.AddDataField(pt.PivotFields(detail.name_of(t, 'num')), '件数', -4112)
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def build_truth(xl, before, truth):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(1)
        for i in range(1, int(ws.PivotTables().Count) + 1):
            pt = ws.PivotTables(i)
            if any(str(pt.DataFields.Item(k).Name) == '累計' for k in range(1, int(pt.DataFields.Count) + 1)):
                continue
            src = str(pt.DataFields.Item(1).SourceName)
            base = str(pt.RowFields.Item(1).Name)
            df = pt.AddDataField(pt.PivotFields(src), '累計', -4157)
            df.Calculation = 5
            df.BaseField = base
            df.NumberFormat = '#,##0'
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kinds, count in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            psheet = random.Random(t['seed']).choice(['ピボット', '集計', '課別集計', 'Sheet2'])
            pv.write_before(before, t, psheet, False)
            if kinds:
                pv.add_pivots(xl, before, t, psheet, kinds, False, False)
            if count:
                add_count(xl, before, t)
            build_truth(xl, before, truth)
            print('  ', stem, t['sheet'], kinds, '件数も' if count else '')
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', detail.design(s(1), rng.randint(20, 60)), ['rows'], False),
          ('未見2_列とフィルタ', detail.design(s(2), 30), ['cols', 'page'], False),
          ('未見3_表題と件数', detail.design(s(3), 25), ['title'], True),
          ('未見4_行2つ', detail.design(s(4), 28), ['rows2'], False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('累計P_本番', detail.design(4901, 30), ['rows'], False),
          ('累計P_試験A', detail.design(4902, 200, n_items=10), ['rows'], False),
          ('累計P_試験B', detail.design(4903, 25), ['rows', 'cols'], False),
          ('累計P_試験C', detail.design(4904, 24), ['cols'], False),
          ('累計P_試験D', detail.design(4905, 26), ['title'], False),
          ('累計P_試験E', detail.design(4906, 28), ['rows2'], False),
          ('累計P_試験G', detail.design(4907, 22), ['rows'], True),
          ('累計P_試験H', detail.design(4908, 20), [], False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '累計P_本番_正解.xlsx'), os.path.join(HERE, f'累計P_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
