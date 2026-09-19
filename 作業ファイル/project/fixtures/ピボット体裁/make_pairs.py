# -*- coding: utf-8 -*-
"""アクティブなシートのピボットを、いつもの見せ方にそろえる（F ピボット 53・2026-09-18 第二期）。

列の決まり: 選ぶ列は無い。アクティブなシートのピボット全部。フィールドの置き方は変えない。
正解の決まり（依頼文に全部書く）:
  - スタイル PivotStyleMedium9（TableStyle2）・表形式（RowAxisLayout xlTabularRow）・ラベルの繰り返し（RepeatAllLabels xlRepeatLabels）・
    行のフィールドの小計なし（Subtotals を全部 False）・空白のセルに 0（NullString "0"・DisplayNullString True）・値のフィールドの表示形式 #,##0。
  - ほかのシートのピボットは触らない。ピボットが無ければ何もしない。
組: 本番 行 2 つ／A 列＝科目／B ピボット 2 つ／C フィルタつき／D 表題の下（A3）／E 行 1 つ／F 撃った後にもう一度／G ほかのシートにも（それは残す）／H ピボットなし
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

REQUEST = ('アクティブなシートのピボットテーブルを全部、いつもの見せ方にそろえてください。フィールドの置き方（行・列・値・フィルタ）は変えない。'
           'スタイルは PivotStyleMedium9（pt.TableStyle2）。表形式（pt.RowAxisLayout xlTabularRow）。ラベルを繰り返す（pt.RepeatAllLabels xlRepeatLabels）。'
           '行のフィールドは全部小計なし（PivotField.Subtotals の 12 個を全部 False）。空白のセルに 0 を出す（pt.NullString = "0"・pt.DisplayNullString = True）。'
           '値のフィールドの表示形式は全部 #,##0（DataField.NumberFormat）。ほかのシートのピボットは触らない。ピボットが無ければ何もしない。')
PHRASES = ['ピボットをいつもの見た目にして', 'ピボットの体裁を整えて', 'ピボットを表形式にして小計を消して',
           'ピボットテーブルの見せ方をそろえて', 'ピボットのデザインをいつもの形に', 'ピボットを見やすい表の形に']


def build_truth(xl, before, truth):
    wb = xl.Workbooks.Open(before)
    try:
        ws = wb.Worksheets(1)
        for i in range(1, int(ws.PivotTables().Count) + 1):
            pt = ws.PivotTables(i)
            pt.TableStyle2 = 'PivotStyleMedium9'
            pt.RowAxisLayout(1)
            pt.RepeatAllLabels(2)
            for k in range(1, int(pt.RowFields.Count) + 1):
                pf = pt.RowFields.Item(k)
                pf.Subtotals = tuple([False] * 12)
            pt.NullString = '0'
            pt.DisplayNullString = True
            for k in range(1, int(pt.DataFields.Count) + 1):
                pt.DataFields.Item(k).NumberFormat = '#,##0'
        wb.SaveAs(truth, FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kinds, other in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            psheet = random.Random(t['seed']).choice(['ピボット', '集計', '課別集計', 'Sheet2'])
            pv.write_before(before, t, psheet, False)
            if kinds:
                pv.add_pivots(xl, before, t, psheet, kinds, False, other)
            build_truth(xl, before, truth)
            print('  ', stem, t['sheet'], kinds, 'ほかのシートにも' if other else '')
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_行2つ', detail.design(s(1), rng.randint(20, 60)), ['rows2'], False),
          ('未見2_列とフィルタ', detail.design(s(2), 30), ['cols', 'page'], False),
          ('未見3_表題とほかのシート', detail.design(s(3), 25), ['title'], True),
          ('未見4_行1つ', detail.design(s(4), 28), ['rows'], False)], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('体裁P_本番', detail.design(5001, 30), ['rows2'], False),
          ('体裁P_試験A', detail.design(5002, 200, n_items=10), ['cols'], False),
          ('体裁P_試験B', detail.design(5003, 25), ['rows2', 'cols'], False),
          ('体裁P_試験C', detail.design(5004, 24), ['page'], False),
          ('体裁P_試験D', detail.design(5005, 26), ['title'], False),
          ('体裁P_試験E', detail.design(5006, 28), ['rows'], False),
          ('体裁P_試験G', detail.design(5007, 22), ['rows2'], True),
          ('体裁P_試験H', detail.design(5008, 20), [], False)], HERE)
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '体裁P_本番_正解.xlsx'), os.path.join(HERE, f'体裁P_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
