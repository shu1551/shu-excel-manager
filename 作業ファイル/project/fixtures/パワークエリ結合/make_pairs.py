# -*- coding: utf-8 -*-
"""支出明細と課マスタをパワークエリでマージする（鍛える回路の題材・パワークエリ 3 本目・2026-09-18）。

仕事: テーブル T支出明細 に、テーブル T課マスタ を課コードで左外部結合して課名を付け、新しいシート「PQ支出課名」の A1 に
      テーブル「Q支出課名」として読み込む。クエリ名も Q支出課名。
正解の決まり（依頼文に全部書く）:
  - 列は 伝票番号・支出日・課コード・課名・支出額 の順。伝票番号の昇順。マスタに無いコードの課名は「未登録」。
  - 同じ名前のクエリやシートが既にあれば消して作り直す。
組: 本番 30 行／A 3,000 行／B マスタに無いコード／C 列の並びが違う（課コードが右端）／D 表題の下にテーブル／
  E 前のクエリとシート／F 撃った後にもう一度／G マスタが同じシート
"""
import importlib.util
import os
import random
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', '..', 'python_scripts')))
_spec = importlib.util.spec_from_file_location('mp_ref', os.path.join(HERE, '..', 'テーブル参照', 'make_pairs.py'))
R = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(R)
QNAME = 'Q支出課名'
OUT_SHEET = 'PQ支出課名'
M = '''let
    Meisai = Excel.CurrentWorkbook(){[Name="T支出明細"]}[Content],
    Master = Excel.CurrentWorkbook(){[Name="T課マスタ"]}[Content],
    Joined = Table.NestedJoin(Meisai, {"課コード"}, Master, {"課コード"}, "M", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "M", {"課名"}, {"課名"}),
    Filled = Table.ReplaceValue(Expanded, null, "未登録", Replacer.ReplaceValue, {"課名"}),
    Cols = Table.SelectColumns(Filled, {"伝票番号", "支出日", "課コード", "課名", "支出額"}),
    Sorted = Table.Sort(Cols, {{"伝票番号", Order.Ascending}}),
    Typed = Table.TransformColumnTypes(Sorted, {{"伝票番号", Int64.Type}, {"支出日", type date}, {"課コード", Int64.Type},
        {"課名", type text}, {"支出額", Int64.Type}})
in
    Typed'''
REQUEST = ('パワークエリで、テーブル T支出明細 にテーブル T課マスタ を課コードで左外部結合（マージ）して課名を付け、新しいシート「PQ支出課名」の A1 に'
           'テーブル「Q支出課名」として読み込んでください。クエリ名も Q支出課名。列は 伝票番号・支出日・課コード・課名・支出額 の順で、'
           '伝票番号の昇順。マスタに無いコードの課名は「未登録」。同じ名前のクエリやシートが既にあれば消して作り直す')


def load_query(wb, formula=M):
    xl = wb.Application
    for i in range(int(wb.Queries.Count), 0, -1):
        if str(wb.Queries.Item(i).Name) == QNAME:
            wb.Queries.Item(i).Delete()
    for sh in list(wb.Worksheets):
        if str(sh.Name) == OUT_SHEET:
            xl.DisplayAlerts = False
            sh.Delete()
    wb.Queries.Add(QNAME, formula)
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = OUT_SHEET
    conn = f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={QNAME};Extended Properties=\"\""
    lo = ws.ListObjects.Add(SourceType=0, Source=conn, Destination=ws.Range("$A$1"))
    lo.QueryTable.CommandType = 2
    lo.QueryTable.CommandText = f"SELECT * FROM [{QNAME}]"
    lo.QueryTable.Refresh(False)
    lo.Name = QNAME
    return lo


def make(jobs, out):
    import vbam_forge as vf
    xl, pid = vf._start_own()
    try:
        xl.DisplayAlerts = False
        for stem, data, kw in jobs:
            before, truth = os.path.join(out, f"{stem}_前.xlsx"), os.path.join(out, f"{stem}_正解.xlsx")
            stale = kw.pop('stale', False)
            rows, kas = data
            rng = random.Random(len(rows))
            rows = sorted(rows, key=lambda r: rng.random())             # 伝票番号の順に並んでいない明細
            R.write_before(before, (rows, kas), **kw)
            if stale:                                                   # 前のクエリ（課名なし・降順）とシート
                wb = xl.Workbooks.Open(before)
                try:
                    load_query(wb, M.replace('Order.Ascending', 'Order.Descending'))
                    wb.Worksheets(1).Activate()
                    wb.Save()
                finally:
                    wb.Close(SaveChanges=False)
            wb = xl.Workbooks.Open(before)
            try:
                load_query(wb)
                wb.Worksheets(1).Activate()
                wb.SaveAs(truth, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
            print('  ', stem)
    finally:
        xl = None
        vf._quit(pid)


def unseen(seed):
    out = os.path.join(HERE, f"未見_種{seed}")
    os.makedirs(out, exist_ok=True)
    rng = random.Random(seed)
    s = lambda k: seed * 10 + k      # noqa: E731
    make([('未見1_標準', R.build(s(1), rng.randint(10, 80)), {}),
          ('未見2_無いコードと表題', R.build(s(2), 40, unknown=True), {'title': True}),
          ('未見3_並び違いと前のクエリ', R.build(s(3), 25), {'cols': R.SWAP, 'stale': True}),
          ('未見4_マスタが同じシート', R.build(s(4), 30, unknown=True), {'master_same_sheet': True})], out)
    return out


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--unseen':
        print(unseen(int(sys.argv[2])))
        sys.exit(0)
    make([('結合_本番', R.build(3, 30), {}),
          ('結合_試験A', R.build(13, 3000), {}),
          ('結合_試験B', R.build(23, 40, unknown=True), {}),
          ('結合_試験C', R.build(33, 25), {'cols': R.SWAP}),
          ('結合_試験D', R.build(43, 22), {'title': True}),
          ('結合_試験E', R.build(53, 20), {'stale': True}),
          ('結合_試験G', R.build(63, 18), {'master_same_sheet': True})], HERE)
    shutil.copyfile(os.path.join(HERE, '結合_本番_正解.xlsx'), os.path.join(HERE, '結合_試験F_前.xlsx'))
    shutil.copyfile(os.path.join(HERE, '結合_本番_正解.xlsx'), os.path.join(HERE, '結合_試験F_正解.xlsx'))
    open(os.path.join(HERE, 'request.txt'), 'w', encoding='utf-8').write(REQUEST + '\n')
    lines = ['パワークエリで支出明細と課マスタをマージして', '課マスタを結合して課名付きの明細をクエリで', 'T支出明細とT課マスタをパワークエリで結合',
             'クエリのマージで課名を付けた明細を作って', '明細にマスタの課名をパワークエリで付けて', 'PQ支出課名のシートにマージ結果を読み込んで']
    open(os.path.join(HERE, 'phrases.txt'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print('ok')
