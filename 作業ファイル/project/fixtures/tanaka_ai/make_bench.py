# -*- coding: utf-8 -*-
"""オフィス田中「ExcelのAI活用」の題を、試しブックに再現する（2026-10-02・第 2 段の 1＝道具の目の試験）。

  py make_bench.py            → bench_tanaka.xlsm（題ごとに 1 シート）

田中さんが Copilot に見せた「エラー」「エラーにならないのにおかしい結果」「汚れ」「ばらばらの式」「条件付き書式の罠」を
シートごとに作る。値は田中さんのページの例に合わせる（名前は架空）。別の Excel（DispatchEx）で作り、自分の PID だけ畳む。
"""
import datetime
import gc
import os
import time

import win32com.client
import win32process

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'bench_tanaka.xlsm')

UDF = '''Function CHECK(s As String) As Long
    CHECK = Val(Left(s, 1))
End Function
'''


def sheet(wb, name):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    return ws


def build(xl):
    wb = xl.Workbooks.Add()
    first = wb.Worksheets(1)

    # 03 データのクリーニング: 表記ゆれ・余分な空白・文字の数字
    ws = sheet(wb, '03清掃')
    ws.Range('A1:C1').Value = ('地域', '名前', '数値')
    rows = [('Tokyo', '佐倉', 100), ('Tokyoo', '佐倉 ', 200), ('Tokyoo', '小原　', 300), ('Osaka', ' 田中', 400),
            ('osaka', '佐 倉', 500), ('Fukuoka', '小原  ', 600), ('Hukuoka', '田中', 700)]
    for i, r in enumerate(rows, start=2):
        ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value = r
    ws.Range('C3').Value = "'200"
    ws.Range('C5').NumberFormat = '@'
    ws.Range('C5').Value = '400'

    # 04 説明: 近似一致の VLOOKUP（FALSE 忘れ）・変な数式・名前の LAMBDA と UDF
    ws = sheet(wb, '04説明')
    ws.Range('A1:B1').Value = ('記号', '名前')
    for i, v in enumerate(['A-102', 'A-101', 'A-104', 'A-103', 'A-105'], start=2):
        ws.Cells(i, 1).Value = v
    ws.Range('D1:E1').Value = ('記号', '名前')
    for i, (k, n) in enumerate([('A-101', '田中'), ('A-102', '小原'), ('A-103', '佐倉')], start=2):
        ws.Cells(i, 4).Value, ws.Cells(i, 5).Value = k, n
    for r in range(2, 7):
        ws.Cells(r, 2).Formula = f'=IFERROR(VLOOKUP(A{r},$D$2:$E$4,2),"")'
    ws.Range('G1:G4').Value = (('100円',), ('200円',), ('300円',), ('400円',))
    ws.Range('H1').FormulaArray = '=FIXED(SUM(VALUE((SUBSTITUTE(G1:G4,"円","")))),FALSE)&"円"'
    wb.Names.Add(Name='TEST', RefersTo='=LAMBDA(a,XLOOKUP(a,{1,2,3},{"田中","小原","佐倉"},"---"))')
    ws.Range('J1').Value = '3-ABC'
    ws.Range('K1').Formula2 = '=TEST(CHECK(J1))'

    # 05 エラー: 型の違い・日付の文字・引数の順・MAP・エラーを含む集計・消えた列／名前
    ws = sheet(wb, '05エラー')
    ws.Range('A1:B1').Value = ('番号', '名前')
    ws.Range('A2').Value = 102
    ws.Range('D1:E1').Value = ('番号', '名前')
    for i, (k, n) in enumerate([('101', '田中'), ('102', '小原'), ('103', '佐倉')], start=2):
        ws.Cells(i, 4).NumberFormat = '@'
        ws.Cells(i, 4).Value, ws.Cells(i, 5).Value = k, n
    ws.Range('B2').Formula = '=VLOOKUP(A2,$D$2:$E$4,2,FALSE)'                       # 数値と文字の番号
    ws.Range('A6').Value = '2025/8/1(金)'
    for i, d in enumerate([datetime.datetime(2025, 8, 1), datetime.datetime(2025, 8, 2)], start=6):
        ws.Cells(i, 4).Value = d
        ws.Cells(i, 5).Value = ('朝礼', '会議')[i - 6]
    ws.Range('B6').Formula = '=VLOOKUP(TEXT(A6,"yyyy/m/d"),$D$6:$E$7,2,FALSE)'      # 文字の日付と日付
    ws.Range('G1:H1').Value = ('数値', '名前')
    for i, (v, n) in enumerate([(10, '田中'), (20, '佐倉'), (30, '小原'), (40, '花澤')], start=2):
        ws.Cells(i, 7).Value, ws.Cells(i, 8).Value = v, n
    ws.Range('J1').Value = '佐倉'
    ws.Range('J2').Formula2 = '=INDEX(G2:G5,MATCH(H2:H5,J1,0))'                       # MATCH の引数の順
    ws.Range('K2').Formula = '=RANK($G$2:$G$5,G2)'                                   # RANK の引数の順
    ws.Range('L2').Formula2 = '=MAP(G2:G5,LAMBDA(a,b,a+b))'                          # MAP と LAMBDA の引数の数
    ws.Range('N2').Value, ws.Range('N3').Value = 10, 20
    ws.Range('N4').Formula = '=1/0'
    ws.Range('N5').Value = 30
    ws.Range('N6').Formula = '=MIN(N2:N5)'                                           # エラーを含む集計
    ws.Range('P2').Value = 101
    ws.Range('Q2').Formula = '=VLOOKUP(P2,#REF!,2,FALSE)'                            # 列を消して #REF!
    ws.Range('Q3').Formula = '=VLOOKUP(P2,リスト,2,FALSE)'                            # 名前を消して #NAME?

    # 17 条件付き書式: 行全体を塗る（アクティブセルの罠は cond-format の道具で確かめる）
    ws = sheet(wb, '17条件')
    ws.Range('A1:D1').Value = ('日付', '名前', '記号', '数値')
    for i, (n, v) in enumerate([('田中', 300), ('小原', 600), ('佐倉', 450), ('田中', 700), ('花澤', 120),
                                ('小原', 520), ('佐倉', 90), ('田中', 501)], start=2):
        ws.Cells(i, 1).Value = datetime.datetime(2025, 8, i)
        ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value = n, 'A', v

    # 18 エラーにならないのにおかしい: 文字の数値の合計・近似一致・末尾の空白
    ws = sheet(wb, '18おかしい')
    ws.Range('A1:B1').Value = ('名前', '数値')
    for i, (n, v) in enumerate([('田中', 100), ('小原', 200), ('佐倉', 300), ('花澤', 400), ('雨宮', 500)], start=2):
        ws.Cells(i, 1).Value = n
        ws.Cells(i, 2).NumberFormat = '@'
        ws.Cells(i, 2).Value = str(v)
        ws.Cells(i, 2).HorizontalAlignment = -4152
    ws.Range('B7').Formula = '=SUM(B2:B6)'
    ws.Range('D1:E1').Value = ('名前', '記号')
    for i, (n, k) in enumerate([('田中', 'A'), ('小原', 'B'), ('佐倉', 'C'), ('花澤', 'D'), ('雨宮', 'E'), ('久保', 'F')],
                               start=2):
        ws.Cells(i, 4).Value, ws.Cells(i, 5).Value = n, k
    ws.Range('G1:H1').Value = ('名前', '記号')
    for i, n in enumerate(['佐倉', '久保', '田中', '雨宮'], start=2):
        ws.Cells(i, 7).Value = n
        ws.Cells(i, 8).Formula = f'=VLOOKUP(G{i},$D$2:$E$7,2)'
    ws.Range('J1:L1').Value = ('名前', '地域', '数値')
    for i, (n, a, v) in enumerate([('田中', '東京', 1), ('小原', '大阪', 2), ('田中', '福岡', 3), ('田中 ', '東京', 4),
                                   ('佐倉', '大阪', 5)], start=2):
        ws.Cells(i, 10).Value, ws.Cells(i, 11).Value, ws.Cells(i, 12).Value = n, a, v
    ws.Range('N2').Formula2 = '=FILTER(J2:L6,J2:J6="田中")'

    # 21 ばらばらの式: 合計の列に参照のずれ
    ws = sheet(wb, '21式のずれ')
    ws.Range('A1:E1').Value = ('名前', '国語', '数学', '英語', '合計')
    for r in range(2, 11):
        ws.Cells(r, 1).Value = f'生徒{r - 1}'
        for c in range(2, 5):
            ws.Cells(r, c).Value = 50 + (r * 7 + c * 13) % 50
        ws.Cells(r, 5).Formula = f'=SUM(B{r}:D{r})'
    ws.Range('E4').Formula = '=SUM(B4:C4)'
    ws.Range('E8').Formula = '=SUM(C8:D8)'

    first.Name = '目次'
    first.Range('A1').Value = 'オフィス田中「ExcelのAI活用」の題の再現（各シート名の数字＝題の番号）'

    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = 'Mod題'
    comp.CodeModule.AddFromString(UDF)

    # 05 の #REF!・#NAME? は「消した後」の式を直接書いたので、名前「リスト」は作らない（消えた状態）
    if os.path.exists(OUT):
        os.remove(OUT)
    wb.SaveAs(OUT, FileFormat=52)
    print('保存:', OUT)
    wb.Close(SaveChanges=False)


def main():
    xl = win32com.client.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        build(xl)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        time.sleep(2)
        os.system(f'taskkill /PID {pid} /F >nul 2>&1')


if __name__ == '__main__':
    main()
