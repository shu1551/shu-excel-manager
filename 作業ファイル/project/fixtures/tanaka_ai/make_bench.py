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

# 第 2 段（AI の道）: 田中さんの「止まるマクロ」「遅いマクロ」
SLOW = """Sub CSV読込12()
    Dim buf As String, n As Long, tmp
    Open ThisWorkbook.Path & "\\tanaka12.csv" For Input As #1
    n = 1
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")
        Worksheets("12止まる").Cells(n, 1).Value = tmp(0)
        Worksheets("12止まる").Cells(n, 2).Value = tmp(1)
        Worksheets("12止まる").Cells(n, 3).Value = tmp(2)
        n = n + 1
    Loop
    Close #1
End Sub

Sub 田中削除25()
    Dim i As Long
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If Cells(i, 1).Value = "田中" Then Rows(i).Delete
    Next i
End Sub

Sub 値貼付26()
    Dim i As Long, j As Long
    With Worksheets("26値貼付")
        For i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            For j = 2 To Worksheets("26マスタ").Cells(Rows.Count, 1).End(xlUp).Row
                If .Cells(i, 1).Value = Worksheets("26マスタ").Cells(j, 1).Value Then
                    Worksheets("26マスタ").Cells(j, 2).Copy
                    .Cells(i, 2).PasteSpecial xlPasteValues
                    Exit For
                End If
            Next j
        Next i
    End With
    Application.CutCopyMode = False
End Sub
"""

NAMES = ['田中', '小原', '佐倉', '花澤', '雨宮']


def csv_files():
    """08（読み込み）と 12（空行で止まる）の CSV を Shift-JIS で置く。"""
    rows = ['日付,名前,数値']
    for i in range(12):
        rows.append(f'2025/9/{i + 1},{NAMES[i % 5] if i % 3 else "田中"},{(i + 1) * 1250}')
    with open(os.path.join(HERE, 'tanaka08.csv'), 'w', encoding='cp932', newline='\r\n') as f:
        f.write('\n'.join(rows) + '\n')
    rows = ['日付,名前,数値', '2025/9/1,田中,100', '2025/9/2,小原,200', '', '2025/9/3,佐倉,300']
    with open(os.path.join(HERE, 'tanaka12.csv'), 'w', encoding='cp932', newline='\r\n') as f:
        f.write('\n'.join(rows) + '\n')


def fill_slow(wb):
    """25・26 の遅いマクロのデータ（撃つたびに作り直す: 試しの台本からも呼ぶ）。"""
    ws = wb.Worksheets('25削除')
    ws.Cells.Clear()
    data = [('名前', '数値')] + [(NAMES[(i * 7) % 5] if i % 4 else '田中', i) for i in range(1, 10001)]
    ws.Range(f'A1:B{len(data)}').Value = data
    ws = wb.Worksheets('26マスタ')
    ws.Cells.Clear()
    data = [('記号', '単価')] + [(f'K{i:04d}', i * 10) for i in range(1, 1001)]
    ws.Range(f'A1:B{len(data)}').Value = data
    ws = wb.Worksheets('26値貼付')
    ws.Cells.Clear()
    data = [('記号', '単価')] + [(f'K{(i * 37) % 1000 + 1:04d}', None) for i in range(1, 1001)]
    ws.Range(f'A1:B{len(data)}').Value = data


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

    build_stage2(wb)

    first.Name = '目次'
    first.Range('A1').Value = 'オフィス田中「ExcelのAI活用」の題の再現（各シート名の数字＝題の番号）'

    comp = wb.VBProject.VBComponents.Add(1)
    comp.Name = 'Mod題'
    comp.CodeModule.AddFromString(UDF)
    comp.CodeModule.AddFromString(SLOW)
    wb.Worksheets('12止まる').Protect()
    csv_files()

    # 05 の #REF!・#NAME? は「消した後」の式を直接書いたので、名前「リスト」は作らない（消えた状態）
    if os.path.exists(OUT):
        os.remove(OUT)
    wb.SaveAs(OUT, FileFormat=52)
    print('保存:', OUT)
    wb.Close(SaveChanges=False)


def build_stage2(wb):
    """第 2 段（AI の道で作る・直す）の題のシート。"""
    # 06 田中だったら 2 倍（13 行＝10 行固定の書き方は外れる）
    ws = sheet(wb, '06田中2倍')
    ws.Range('A1:C1').Value = ('名前', '数値', '結果')
    for i in range(2, 15):
        ws.Cells(i, 1).Value = NAMES[(i * 3) % 5]
        ws.Cells(i, 2).Value = i * 10
    # 06b C 列が A の行を F 列の既存データの下へコピー（書式も）
    ws = sheet(wb, '06bコピー')
    ws.Range('A1:D1').Value = ('日付', '名前', '区分', '数値')
    ws.Range('F1:I1').Value = ('日付', '名前', '区分', '数値')
    for i in range(2, 14):
        ws.Cells(i, 1).Value = datetime.datetime(2025, 9, i)
        ws.Cells(i, 2).Value = NAMES[i % 5]
        ws.Cells(i, 3).Value = 'A' if i % 3 else 'B'
        ws.Cells(i, 4).Value = i * 1234
    for i in range(2, 5):
        ws.Cells(i, 6).Value = datetime.datetime(2025, 8, i)
        ws.Cells(i, 7).Value = '既存'
        ws.Cells(i, 8).Value = 'A'
        ws.Cells(i, 9).Value = i
    ws.Range('A2:A13').NumberFormat = 'yyyy/m/d'
    ws.Range('F2:F4').NumberFormat = 'yyyy/m/d'
    ws.Range('D2:D13').NumberFormat = '#,##0'
    ws.Range('D2:D13').Interior.Color = 0xCCFFFF
    # 07 F1 の見出しの列を出す（行数は可変）
    ws = sheet(wb, '07見出し')
    ws.Range('A1:D1').Value = ('日付', '名前', '地域', '数値')
    for i in range(2, 10):
        ws.Cells(i, 1).Value = datetime.datetime(2025, 9, i)
        ws.Cells(i, 2).Value = NAMES[i % 5]
        ws.Cells(i, 3).Value = ('東京', '大阪', '福岡')[i % 3]
        ws.Cells(i, 4).Value = i * 100
    ws.Range('F1').Value = '名前'
    # 08 CSV を A1 から（空のシート）
    sheet(wb, '08CSV')
    # 09 9 月と 10 月を結合して田中だけ新しいブックへ
    for m in (9, 10):
        ws = sheet(wb, f'{m}月')
        ws.Range('A1:C1').Value = ('日付', '名前', '数値')
        for i in range(2, 9):
            ws.Cells(i, 1).Value = datetime.datetime(2025, m, i)
            ws.Cells(i, 2).Value = NAMES[(i + m) % 5]
            ws.Cells(i, 3).Value = i * m
    # 10 全角半角・半角スペース・結合セル
    ws = sheet(wb, '10全半角')
    ws.Range('A1').Value = '文字'
    for i, v in enumerate(['ＡＢＣ１２３', 'Ｅｘｃｅｌ　２０２５', '田中 太郎', '東京－ＳＨＩＮＪＵＫＵ（新宿）',
                           'テスト１', '小原 花子 ＶＢＡ', 'Ｎｏ．１０'], start=2):
        ws.Cells(i, 1).Value = v
    ws = sheet(wb, '10結合')
    ws.Range('A1:B1').Value = ('名前', '数値')
    for i in range(2, 9):
        ws.Cells(i, 2).Value = i * 10
    ws.Range('A2').Value = '田中'
    ws.Range('A5').Value = '小原'
    ws.Range('A7').Value = '佐倉'
    ws.Range('A2:A4').Merge()
    ws.Range('A5:A6').Merge()
    ws.Range('A7:A8').Merge()
    # 11 数式で
    ws = sheet(wb, '11数式')
    ws.Range('A1').Value = '文字'
    for i, v in enumerate(['ＡＢＣ１２３', 'Ｅｘｃｅｌ　２０２５', 'ﾃｽﾄ１', '東京\nＯＳＡＫＡ\n福岡'], start=2):
        ws.Cells(i, 1).Value = v
    # 12 止まるマクロ（保護と CSV の空行）
    sheet(wb, '12止まる')
    # 13 名前別・月別を 1 つの式で
    ws = sheet(wb, '13集計')
    ws.Range('A1:C1').Value = ('日付', '名前', '金額')
    for i in range(2, 20):
        ws.Cells(i, 1).Value = datetime.datetime(2025, 8 + (i % 3), (i * 3) % 27 + 1)
        ws.Cells(i, 2).Value = NAMES[(i * 2) % 4]
        ws.Cells(i, 3).Value = i * 500
    # 20 数式をメモで
    ws = sheet(wb, '20メモ')
    ws.Range('A1:A4').Value = ((10,), (20,), (30,), (40,))
    ws.Range('C2').Formula = '=IF(SUM(A1:A4)>50,ROUND(AVERAGE(A1:A4),0),MAX(A1:A4))'
    # 23 名前別合計を D1 に 1 つの式で（テーブルでない表）
    ws = sheet(wb, '23名前別')
    ws.Range('A1:B1').Value = ('名前', '金額')
    for i in range(2, 16):
        ws.Cells(i, 1).Value = NAMES[(i * 3) % 5]
        ws.Cells(i, 2).Value = i * 300
    # 25・26 遅いマクロ
    for n in ('25削除', '26値貼付', '26マスタ'):
        sheet(wb, n)
    fill_slow(wb)


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
