# -*- coding: utf-8 -*-
"""パワークエリで別のテーブルを結合して名前を付ける（どの表でも動く形・2026-09-18 作り直し。前の版は PQ支出課名）。

形で決める: アクティブなシートのテーブルと、別のテーブル（共通の見出しの列＝キーを持つ 2 列の表）。
作る物の名前は元のテーブル名から（「（元のテーブル名）_結合」）。見出しの語・シート名・テーブル名は表ごとに違う。
正解の決まり:
  - 元のテーブルに相手のテーブルを キーの列 で左外部結合し、相手のもう 1 つの列を付ける。相手に無いキーは「未登録」。
  - 列は元のテーブルの列の順＋相手の列。キーの昇順。
  - 新しいシート（名前は 元のテーブル名 と _結合 をつないだ名前）の A1 に同じ名前のテーブルとして読み込む。クエリ名も同じ。
  - 同じ名前のクエリやシートが既にあれば消して作り直す。
組: 本番／A 大量 800 行／B マスタが同じシート／C マスタの列の並びが違う／D マスタに無いキー／E 前のクエリとシート／
  F 撃った後にもう一度／G 列が多い
"""
import os
import random
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
import importlib.util   # noqa: E402

_spec = importlib.util.spec_from_file_location('mp_ref', os.path.join(HERE, '..', 'マスタ参照', 'make_pairs.py'))
R = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(R)
_spec2 = importlib.util.spec_from_file_location('mp_tate', os.path.join(HERE, '..', '縦持ち', 'make_pairs.py'))
T = importlib.util.module_from_spec(_spec2)
_spec2.loader.exec_module(T)

REQUEST = ('パワークエリで、アクティブなシートのテーブルに、別のテーブル（共通の見出しの列＝キーを持つ表）を'
           'キーの列で左外部結合（マージ）して、相手のもう 1 つの列を付けてください。相手に無いキーは「未登録」。'
           '列は元のテーブルの列の順＋相手の列、キーの昇順。'
           '新しいシート（名前は 元のテーブル名 と _結合 をつないだ名前）の A1 に同じ名前のテーブルとして読み込む。'
           'クエリ名も同じ。同じ名前のクエリやシートが既にあれば消して作り直す')
PHRASES = ['パワークエリでマスタを結合して', 'クエリでマージして名前を付けて', 'マスタを左外部結合したクエリを作って',
           'パワークエリで名前付きの一覧を作って', 'クエリでマスタを付けて', 'キーで結合したテーブルをクエリで']


def m_formula(t):
    names = [t['no_head'], t['date_head'], t['code_head']] + ([t['name_head']] if t['has_name_col'] else []) \
        + [t['amount_head']] + (['備考'] if t['extra_cols'] else [])
    keep = [n for n in names if n != t['name_head']] + [t['name_head']]
    keep_txt = ", ".join(f'"{n}"' for n in keep)
    return f'''let
    Base = Excel.CurrentWorkbook(){{[Name="{t['table']}"]}}[Content],
    Master = Excel.CurrentWorkbook(){{[Name="{t['master_table']}"]}}[Content],
    Dropped = Table.RemoveColumns(Base, {{"{t['name_head']}"}}, MissingField.Ignore),
    Joined = Table.NestedJoin(Dropped, {{"{t['code_head']}"}}, Master, {{"{t['code_head']}"}}, "M", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "M", {{"{t['name_head']}"}}, {{"{t['name_head']}"}}),
    Filled = Table.ReplaceValue(Expanded, null, "未登録", Replacer.ReplaceValue, {{"{t['name_head']}"}}),
    Cols = Table.SelectColumns(Filled, {{{keep_txt}}}),
    Sorted = Table.Sort(Cols, {{{{"{t['code_head']}", Order.Ascending}}}}),
    Typed = Table.TransformColumnTypes(Sorted, {{{{"{t['date_head']}", type date}}, {{"{t['amount_head']}", Int64.Type}}, {{"{t['name_head']}", type text}}}})
in
    Typed'''


def make(jobs, out=HERE):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, t, kw in jobs:
            before = os.path.join(out, f"{stem}_前.xlsx")
            truth = os.path.join(out, f"{stem}_正解.xlsx")
            R.write_before(stem, t, out)
            name = t['table'] + '_結合'
            if kw.get('stale'):
                wb = xl.Workbooks.Open(before)
                try:
                    T.load_query(wb, name, m_formula(t).replace('Order.Ascending', 'Order.Descending'), name)
                    wb.Worksheets(t['sheet']).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                T.load_query(wb, name, m_formula(t), name)
                wb.Worksheets(t['sheet']).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem, t['sheet'], t['table'], t['master_table'], name)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    o = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(o, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', R.build(s(1), rng.randint(10, 50)), {}),
          ('未見2_同じシートのマスタ', R.build(s(2), 20, same_sheet=True, master_swapped=True), {}),
          ('未見3_無いキーと前のクエリ', R.build(s(3), 25, missing=3), {'stale': True}),
          ('未見4_列が多い', R.build(s(4), 18, extra_cols=True), {})], o)
    return o


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        return
    make([('結合_本番', R.build(3, 20), {}),
          ('結合_試験A', R.build(13, 400), {}),
          ('結合_試験B', R.build(23, 15, same_sheet=True), {}),
          ('結合_試験C', R.build(33, 16, master_swapped=True), {}),
          ('結合_試験D', R.build(43, 20, missing=2), {}),
          ('結合_試験E', R.build(53, 14), {'stale': True}),
          ('結合_試験G', R.build(63, 14, extra_cols=True), {})])
    for x in ('前', '正解'):
        shutil.copyfile(os.path.join(HERE, '結合_本番_正解.xlsx'), os.path.join(HERE, f'結合_試験F_{x}.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(PHRASES) + '\n')
    print('ok')


if __name__ == '__main__':
    main()
