Attribute VB_Name = "表の整理_調べる"
' 表の整理_調べる - 棚（言葉で頼んで撃つマクロ）の調べる（調べる・一覧・点検）（2026-09-23 表の整理から分けた）

Sub 選んだ列の重複する値を一覧にする()
    ' 選んでいる列の中で2回以上出現する重複データを検出し、重複件数とともに表の右側に一覧表示する。
    ' 依頼の語: 重複キー|番号の重複|番号が二重|同じ番号|列の重複|伝票番号を洗|列で重複|重複している値|重複している番号|重複を一覧|番号の一覧を右|何件あるか右|回以上出|重複チェック
    ' 依頼の組: 重複,かぶ,二重,同じ+番号,キー,値,伝票+調べ,探,一覧,洗い出,チェック,確認,出,ないか+-消,削除,除,マクロ
    ' 扱う: 重複キー一覧 重複洗い出し 件数 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Dim keyCol As Long, hdrRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim outCol As Long
    Dim keyVal As String, normKey As String
    Dim dict As Object
    Dim orderList() As String, orderCount As Long
    Dim i As Long, cnt As Long, pos As Long
    Dim bodyLastRow As Long, cv As String

    Set ws = ActiveSheet
    keyCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    Dim ur As Range
    Set ur = ws.UsedRange
    ' UsedRangeの右端は本体列のみ（出力済み列を除くため、keyCol基準で元表の右端を探す）
    ' 元表の右端: hdrRowの左からkeyColを含む連続した列を見る
    ' 簡易に: UsedRangeの右端を使うが、出力列があれば出力列の左の空列の左が元表右端
    lastRow = ur.Row + ur.rows.Count - 1

    ' 元表の右端列を求める: hdrRow行で、左から連続して値がある列の末尾
    ' ただし既存出力列（空列＋見出し）があるかもしれないので、
    ' UsedRangeの右端から逆に空列を探す
    Dim urLastCol As Long
    urLastCol = ur.Column + ur.Columns.Count - 1

    ' 元表の右端 = hdrRow行で最初の空列の手前（ただし出力済みの表が右にある場合を考慮）
    ' 方針: hdrRow行を左から走査し、最初に空になった列の1つ前を元表右端とする
    Dim tableLastCol As Long
    tableLastCol = 0
    For c = ur.Column To urLastCol
        If Trim(セルの字(ws.Cells(hdrRow, c).Value)) = "" Then
            tableLastCol = c - 1
            Exit For
        End If
    Next c
    If tableLastCol = 0 Then tableLastCol = urLastCol

    lastCol = tableLastCol

    ' 合計行を除く本文最終行
    bodyLastRow = lastRow
    For r = lastRow To hdrRow + 1 Step -1
        cv = Trim(CStr(ws.Cells(r, keyCol).Value))
        If 集計の語か(cv) Then
            bodyLastRow = r - 1
        Else
            Exit For
        End If
    Next r

    ' 出力先列 = lastCol + 2
    outCol = 右の出し先(ws, hdrRow, lastCol + 2, 3, "*|件数|行")   ' 右に別の物があれば避ける（2026-09-23）

    ' 既存の出力表を消去（outCol以降のhdrRow行に値があれば）
    Dim clrLast As Long
    clrLast = hdrRow
    For r = hdrRow To lastRow + 50
        If Trim(セルの字(ws.Cells(r, outCol).Value)) <> "" Or _
           Trim(CStr(ws.Cells(r, outCol + 1).Value)) <> "" Or _
           Trim(CStr(ws.Cells(r, outCol + 2).Value)) <> "" Then
            clrLast = r
        End If
    Next r
    If clrLast >= hdrRow Then
        ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(clrLast, outCol + 2)).ClearContents
    End If

    ' キー列集計
    Set dict = CreateObject("Scripting.Dictionary")
    orderCount = 0
    ReDim orderList(1 To bodyLastRow - hdrRow + 1)

    For r = hdrRow + 1 To bodyLastRow
        keyVal = Trim(CStr(ws.Cells(r, keyCol).Value))
        If セルの字(keyVal) = "" Then GoTo NextRow
        normKey = 数のキー(keyVal)      ' CLng で 13 桁の番号が止まり、1.5 と 2 を同じにしていた（2026-09-24 総点検）
        If dict.exists(normKey) Then
            dict(normKey) = dict(normKey) & ", " & r
        Else
            dict.Add normKey, CStr(r)
            orderCount = orderCount + 1
            orderList(orderCount) = normKey
        End If
NextRow:
    Next r

    ' 見出し
    ws.Cells(hdrRow, outCol).Value = ws.Cells(hdrRow, keyCol).Value
    ws.Cells(hdrRow, outCol + 1).Value = "件数"
    ws.Cells(hdrRow, outCol + 2).Value = "行"

    ' 重複のみ出力
    Dim outRow As Long
    outRow = hdrRow + 1
    Dim rowsStr As String

    For i = 1 To orderCount
        rowsStr = dict(orderList(i))
        cnt = 1
        pos = 1
        Do
            pos = InStr(pos, rowsStr, ",")
            If pos = 0 Then Exit Do
            cnt = cnt + 1
            pos = pos + 1
        Loop
        If cnt >= 2 Then
            ws.Cells(outRow, outCol).Value = orderList(i)
            ws.Cells(outRow, outCol + 1).Value = cnt
            ws.Cells(outRow, outCol + 2).Value = rowsStr
            outRow = outRow + 1
        End If
    Next i
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 2))
        Call 表の仕上げ
    End If

End Sub

Sub 選んだ列が空欄の行を一覧にする()
    ' 選んでいる列で必須入力が漏れている（空欄になっている）行番号と該当データを調査シートに一覧抽出する。
    ' 依頼の語: 空欄チェック|空いているセルがある|入力漏れ|記入漏れ|未記入|空欄がある行
    ' 依頼の組: 空欄,未入力,入力されていない,未記入,記入漏れ,入力漏れ,空いているセル,空白のセル+行,一覧,探,調べ,チェック+-埋め,上の値,削除,消,詰
    ' 扱う: 空欄チェック 入力漏れ 記入漏れ 未記入 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Dim selCol As Long
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim outCol As Long
    Dim r As Long
    Dim outRow As Long
    Dim cellVal As String
    Dim isAllBlank As Boolean
    Dim isTotalRow As Boolean
    Dim ur As Range
    Dim c As Long
    Dim cv As String

    Set ws = ActiveSheet
    selCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    Set ur = ws.UsedRange
    lastRow = 表の本文の最終行(ws)      ' 表の下のメモの行を空欄と言わない（2026-09-24 総点検）
    lastCol = ur.Column + ur.Columns.Count - 1

    ' 既に「空欄の行」見出しの列を探す（表の右側）
    outCol = 0
    Dim checkC As Long
    For checkC = lastCol To lastCol + 20
        If Trim$(セルの字(ws.Cells(hdrRow, checkC).Value)) = "空欄の行" Then
            outCol = checkC
            Exit For
        End If
    Next checkC

    If outCol = 0 Then
        ' 表の右に1列空けて配置: lastCol+2
        ' ただし lastCol+1 列が既に使われている場合は lastCol+2 がずれないよう確認
        outCol = 右の出し先(ws, hdrRow, lastCol + 2, 1, "空欄の行", ur.Column)   ' 前の一覧より右に物があっても見つけて書き直す（2026-09-23）
        If セルの字(ws.Cells(hdrRow, outCol).Value) = "空欄の行" Then ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(lastRow, outCol)).ClearContents
    Else
        ' 既存列をクリア
        ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(lastRow, outCol)).ClearContents
    End If

    ws.Cells(hdrRow, outCol).Value = "空欄の行"
    outRow = hdrRow + 1

    For r = hdrRow + 1 To lastRow
        ' 行がまるごと空かチェック（出力列を除く）
        isAllBlank = True
        For c = ur.Column To lastCol
            If Trim$(セルの字(ws.Cells(r, c).Value)) <> "" Then
                isAllBlank = False
                Exit For
            End If
        Next c
        If isAllBlank Then GoTo NextRow

        ' 合計行チェック
        isTotalRow = False
        For c = ur.Column To lastCol
            cv = Trim$(CStr(ws.Cells(r, c).Value))
            If 集計の語か(cv) Then
                isTotalRow = True
                Exit For
            End If
        Next c
        If isTotalRow Then GoTo NextRow

        cellVal = Trim$(CStr(ws.Cells(r, selCol).Value))
        If セルの字(cellVal) = "" Then
            ws.Cells(outRow, outCol).Value = CStr(r)
            outRow = outRow + 1
        End If

NextRow:
    Next r
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol))
        Call 表の仕上げ
    End If

End Sub

Sub 合計行が明細の和と合うか確かめる()
    ' 表の合計行の数値が、上の明細行の合計値と一致しているか検算し、差異があれば報告する。
    ' 依頼の語: 検算|合計が合って|合計の行を検算|縦計の検算|合計欄が正しいか|合計と内訳の和|足し算
    ' 依頼の組: 合計,縦計+正しい,合って,合う,検算,確かめ,チェック,確認+-数式
    ' 扱う: 検算 合計検算
    ' 見出し: なし
    ' 形: 数の列

    Dim ws As Object
    Set ws = ActiveSheet

    Dim ur As Object
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    ' 見出しの行を探す: 空でないセルが2つ以上並ぶ最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    For r = urTop To urBottom
        Dim cnt As Long
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    ' 表の右端: 見出し行でいちばん左の列から右へ見て空のセルに当たる手前の列
    Dim tblRight As Long
    ' 左端＝見出しの行で最初に値のある列（表題が A1・表が C3 からのとき A3 が空で右端を取り違え、何もせずに終わった。2026-09-22）
    Do While urLeft < urRight And セルの字(ws.Cells(hdrRow, urLeft).Value) = ""
        urLeft = urLeft + 1
    Loop
    tblRight = urLeft - 1
    For c = urLeft To urRight
        If セルの字(ws.Cells(hdrRow, c).Value) = "" Then Exit For
        tblRight = c
    Next c
    If tblRight < urLeft Then Exit Sub

    ' 合計の行: 見出し行より下で「合計」「計」「総計」のセルがある最後の行
    Dim totalRow As Long
    totalRow = 0
    For r = hdrRow + 1 To urBottom
        For c = urLeft To tblRight
            Dim cv As String
            cv = CStr(ws.Cells(r, c).Value)
            If cv = "合計" Or cv = "計" Or cv = "総計" Then
                totalRow = r
                Exit For
            End If
        Next c
    Next r
    If totalRow = 0 Then Exit Sub

    ' 検算する列: 合計行に数が入っている列
    Dim numCols() As Long
    Dim numColCount As Long
    numColCount = 0
    ReDim numCols(1 To tblRight - urLeft + 1)
    For c = urLeft To tblRight
        If IsNumeric(ws.Cells(totalRow, c).Value) And セルの字(ws.Cells(totalRow, c).Value) <> "" Then
            numColCount = numColCount + 1
            numCols(numColCount) = c
        End If
    Next c
    If numColCount = 0 Then Exit Sub

    ' 明細範囲: 見出し行の次の行から合計行の1つ上
    Dim detailFirst As Long, detailLast As Long
    detailFirst = hdrRow + 1
    detailLast = totalRow - 1

    ' 出力先: 表の右端の2つ右
    Dim outCol As Long
    outCol = 右の出し先(ws, hdrRow, tblRight + 2, 5, "検算の列|合計の行|明細の和|差|判定")   ' 右に別の物があれば避ける（2026-09-23）

    ' 既に同じ表が右にあれば消す
    If ws.Cells(hdrRow, outCol).Value = "検算の列" Then
        ' 前の結果は見出しの行から下の 5 列だけ消す（列まるごと消すと、同じ列の表題・注記まで消える）
        ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(hdrRow + tblRight - urLeft + 2, outCol + 4)).ClearContents
    End If

    ' 見出し
    ws.Cells(hdrRow, outCol).Value = "検算の列"
    ws.Cells(hdrRow, outCol + 1).Value = "合計の行"
    ws.Cells(hdrRow, outCol + 2).Value = "明細の和"
    ws.Cells(hdrRow, outCol + 3).Value = "差"
    ws.Cells(hdrRow, outCol + 4).Value = "判定"

    ' 各検算列の行を書く
    Dim i As Long
    For i = 1 To numColCount
        Dim tc As Long
        tc = numCols(i)
        Dim dataRow As Long
        dataRow = hdrRow + i

        ' 検算の列（見出し値）
        ws.Cells(dataRow, outCol).Value = CStr(ws.Cells(hdrRow, tc).Value)

        ' 合計の行（式で参照）
        ws.Cells(dataRow, outCol + 1).Formula = "=" & ws.Cells(totalRow, tc).Address(True, True)

        ' 明細の和（SUM 絶対参照）
        Dim sumAddr As String
        sumAddr = ws.Cells(detailFirst, tc).Address(True, True) & ":" & ws.Cells(detailLast, tc).Address(True, True)
        ' 明細のあいだの小計の行は明細の和に入れない（小計まで足して合計の 2 倍と比べていた・2026-09-24 棚の試験 F9）
        Dim subMinus As String, sr As Long, scc As Long
        subMinus = ""
        For sr = detailFirst To detailLast
            For scc = urLeft To tblRight
                If VarType(ws.Cells(sr, scc).Value) = vbString Then
                    If 集計の語か(ws.Cells(sr, scc).Value) Then
                        subMinus = subMinus & "-" & ws.Cells(sr, tc).Address(True, True)
                        Exit For
                    End If
                End If
            Next scc
        Next sr
        ws.Cells(dataRow, outCol + 2).Formula = "=SUM(" & sumAddr & ")" & subMinus

        ' 差
        Dim totAddr As String
        Dim sumCellAddr As String
        totAddr = ws.Cells(dataRow, outCol + 1).Address(False, False)
        sumCellAddr = ws.Cells(dataRow, outCol + 2).Address(False, False)
        ws.Cells(dataRow, outCol + 3).Formula = "=" & totAddr & "-" & sumCellAddr

        ' 判定
        Dim diffAddr As String
        diffAddr = ws.Cells(dataRow, outCol + 3).Address(False, False)
        ws.Cells(dataRow, outCol + 4).Formula = "=IF(" & diffAddr & "=0,""一致"",""不一致"")"

        ' 数値列の表示形式
        ws.Cells(dataRow, outCol + 1).NumberFormat = "#,##0"
        ws.Cells(dataRow, outCol + 2).NumberFormat = "#,##0"
        ws.Cells(dataRow, outCol + 3).NumberFormat = "#,##0"
    Next i
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If numColCount > 0 Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(hdrRow + numColCount, outCol + 4))
        Call 表の仕上げ
    End If

End Sub

Sub 同日同額の二重計上を一覧にする()
    ' 同じ日付かつ同じ金額で計上されているデータ（二重払いや二重登録の疑いがある行）を洗い出して一覧にする。
    ' 依頼の語: 二重計上|二重払い|同日同額|同じ日付|同じ金額|同じ日に同じ金額|二重に支払|同じ組
    ' 依頼の組: 同じ日,同日+同じ額,同じ金額,同額
    ' 扱う: 二重計上 二重払い 同日同額 合計
    ' 見出し: なし
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)

    ' 見出し行を探す: 空でないセルが2つ以上並ぶ最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    For r = urTop To urBottom
        Dim nonEmpty As Long
        nonEmpty = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub
    hdrRow = 見出しの下の段(ws, hdrRow, urLeft, urRight)   ' 二段見出しは下の段（2026-09-24）

    ' 表の右端列: 見出し行で左端から右へ見て空のセルに当たる手前の列（結合の右・下は結合の字で見る）
    Dim tblLeft As Long, tblRight As Long
    ' 左端＝見出しの行で最初に値のある列（表題が A1・表が C3 からのとき A3 が空で右端を取り違え、何もしなかった。2026-09-22）
    tblLeft = urLeft
    Do While tblLeft < urRight And 見出しの字(ws.Cells(hdrRow, tblLeft)) = ""
        tblLeft = tblLeft + 1
    Loop
    tblRight = tblLeft - 1
    For c = tblLeft To urRight
        If 見出しの字(ws.Cells(hdrRow, c)) <> "" Then
            tblRight = c
        Else
            Exit For
        End If
    Next c
    If tblRight < tblLeft Then Exit Sub

    Dim dataTop As Long
    dataTop = hdrRow + 1

    ' 日付列を探す: 本文の8割以上が日付の列（いちばん左）
    Dim dateCol As Long
    dateCol = 0
    Dim totalRows As Long
    totalRows = urBottom - dataTop + 1
    If totalRows <= 0 Then Exit Sub
    For c = tblLeft To tblRight
        Dim dateCnt As Long
        dateCnt = 0: neC = 0
        For r = dataTop To urBottom
            Dim v As Variant
            v = ws.Cells(r, c).Value
            If Not IsError(v) Then
                If CStr(v) <> "" Then neC = neC + 1
                If IsDate(v) And CStr(v) <> "" Then dateCnt = dateCnt + 1
            End If
        Next r
        ' 値の入っているセルの 6 割が日付なら日付の列（空行・合計行まで分母に入れて 8 割に届かず、下に合計行のある表で何もしなかった。2026-09-22）
        If neC > 0 And dateCnt >= neC * 0.6 Then
            dateCol = c
            Exit For
        End If
    Next c
    If dateCol = 0 Then Exit Sub

    ' 金額列を探す: 本文の8割以上が数の列のうちいちばん右
    Dim amtCol As Long
    amtCol = 0
    For c = tblRight To tblLeft Step -1
        Dim numCnt As Long
        numCnt = 0
        For r = dataTop To urBottom
            v = ws.Cells(r, c).Value
            If IsNumeric(v) And セルの字(v) <> "" And Not IsDate(v) Then numCnt = numCnt + 1
        Next r
        If numCnt >= totalRows * 0.8 Then
            amtCol = c
            Exit For
        End If
    Next c
    If amtCol = 0 Then Exit Sub

    ' 出力列: 表の右（前に作った一覧があればそこへ書き直す・右に人のメモや別の表があれば避ける。2026-09-24 総点検:
    ' 表の右から使用範囲の右端までを全部消していて、右に置いた人のメモが消えた）
    Dim outStartCol As Long
    outStartCol = 右の出し先(ws, hdrRow, tblRight + 2, tblRight - tblLeft + 2, "元の行|" & 見出しの字(ws.Cells(hdrRow, tblLeft)))
    If セルの字(ws.Cells(hdrRow, outStartCol).Value) = "元の行" Then 前の出力を消す ws, hdrRow, outStartCol, tblRight - tblLeft + 2

    ' 日付と金額のキーでグループ化
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim orderList() As String
    Dim orderCount As Long
    orderCount = 0
    ReDim orderList(0)

    For r = dataTop To urBottom
        Dim dv As Variant
        Dim av As Variant
        dv = ws.Cells(r, dateCol).Value
        av = ws.Cells(r, amtCol).Value
        If セルの字(dv) = "" Or セルの字(av) = "" Then GoTo NextRow
        If Not IsDate(dv) Then GoTo NextRow
        If Not IsNumeric(av) Then GoTo NextRow

        Dim key As String
        key = CStr(CDbl(CDate(dv))) & "_" & CStr(av)

        If dict.exists(key) Then
            dict(key) = dict(key) & "|" & CStr(r)
        Else
            dict.Add key, CStr(r)
            ReDim Preserve orderList(orderCount)
            orderList(orderCount) = key
            orderCount = orderCount + 1
        End If
NextRow:
    Next r

    ' 見出し行を出力列に書く
    Dim colCount As Long
    colCount = tblRight - tblLeft + 1
    ws.Cells(hdrRow, outStartCol).Value = "元の行"
    Dim ci As Long
    For ci = 0 To colCount - 1
        ws.Cells(hdrRow, outStartCol + 1 + ci).Value = 見出しの字(ws.Cells(hdrRow, tblLeft + ci))
    Next ci

    ' 重複組を出力
    Dim outRow As Long
    outRow = hdrRow + 1
    Dim ki As Long
    For ki = 0 To orderCount - 1
        Dim k As String
        k = orderList(ki)
        Dim rows() As String
        rows = Split(dict(k), "|")
        If UBound(rows) >= 1 Then
            Dim ri As Long
            For ri = 0 To UBound(rows)
                Dim srcRow As Long
                srcRow = CLng(rows(ri))
                ws.Cells(outRow, outStartCol).Value = srcRow
                For ci = 0 To colCount - 1
                    Dim srcVal As Variant
                    srcVal = ws.Cells(srcRow, tblLeft + ci).Value
                    ws.Cells(outRow, outStartCol + 1 + ci).Value = srcVal
                Next ci
                ' 日付の表示形式
                ws.Cells(outRow, outStartCol + 1 + (dateCol - tblLeft)).NumberFormat = "yyyy/m/d"
                ' 金額の表示形式
                ws.Cells(outRow, outStartCol + 1 + (amtCol - tblLeft)).NumberFormat = "#,##0"
                outRow = outRow + 1
            Next ri
        End If
    Next ki
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outStartCol), ws.Cells(outRow - 1, outStartCol + 1 + (tblRight - tblLeft)))
        Call 表の仕上げ
    End If

End Sub

Sub おかしな日付の行を一覧にする()
    ' 未来の日付、会計年度外の日付、不正なフォーマットなど、妥当性のない異常な日付データを検出して一覧にする。
    ' 依頼の語: おかしい日付|ありえない日付|あり得ない日付|年度外の日付|日付の入力ミス|未来の日付や空|日付の妥当性|日付がおかしい|列の日付
    ' 依頼の組: 日付+間違い,ミス,おかしい,おかしな,変な,異常,ありえない,あり得ない,妥当,未来,年度外
    ' 扱う: 日付検査 日付異常 年度外 未来の日付 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column
    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urLastCol As Long
    urLastCol = ur.Column + ur.Columns.Count - 1
    Dim urLastRow As Long
    urLastRow = 表の本文の最終行(ws)

    ' 見出し行で左から空セルに当たる手前の列
    Dim tableLastCol As Long
    tableLastCol = 0
    Dim c As Long
    For c = 1 To urLastCol
        If セルの字(ws.Cells(hdrRow, c).Value) = "" Then
            tableLastCol = c - 1
            Exit For
        End If
    Next c
    If tableLastCol = 0 Then tableLastCol = urLastCol

    Dim outCol As Long
    outCol = 右の出し先(ws, hdrRow, tableLastCol + 2, 3, "元の行|*|理由")   ' 右に別の物があれば避ける（2026-09-23）

    Dim lastRow As Long
    lastRow = urLastRow

    ' 既存の一覧を消す（outCol以降をhdrRowから最終行まで）
    Dim clearRow As Long
    For clearRow = hdrRow To lastRow
        ws.Cells(clearRow, outCol).Value = ""
        ws.Cells(clearRow, outCol + 1).Value = ""
        ws.Cells(clearRow, outCol + 2).Value = ""
    Next clearRow

    ' 日付の列見出し
    Dim datHdr As String
    datHdr = CStr(ws.Cells(hdrRow, selCol).Value)

    ' 一覧見出し
    ws.Cells(hdrRow, outCol).Value = "元の行"
    ws.Cells(hdrRow, outCol + 1).Value = datHdr
    ws.Cells(hdrRow, outCol + 2).Value = "理由"

    ' 行まるごと空・合計行判定用ヘルパー
    Dim r As Long
    Dim rowEmpty As Boolean
    Dim isSumRow As Boolean
    Dim cv As String

    ' 表の年度を決める
    Dim yrCounts As Object
    Set yrCounts = CreateObject("Scripting.Dictionary")
    For r = hdrRow + 1 To lastRow
        rowEmpty = True
        For c = 1 To urLastCol
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If rowEmpty Then GoTo NextR1

        isSumRow = False
        For c = 1 To urLastCol
            cv = CStr(ws.Cells(r, c).Value)
            If cv = "合計" Or cv = "計" Or cv = "総計" Then
                isSumRow = True
                Exit For
            End If
        Next c
        If isSumRow Then GoTo NextR1

        Dim dv As Variant
        dv = ws.Cells(r, selCol).Value
        If Not isEmpty(dv) And セルの字(dv) <> "" And IsDate(dv) Then
            Dim d As Date
            d = CDate(dv)
            Dim mn As Integer
            mn = Month(d)
            Dim yr As Integer
            If mn >= 4 Then
                yr = Year(d)
            Else
                yr = Year(d) - 1
            End If
            If yrCounts.exists(yr) Then
                yrCounts(yr) = yrCounts(yr) + 1
            Else
                yrCounts.Add yr, 1
            End If
        End If
NextR1:
    Next r

    Dim mainYear As Integer
    mainYear = 0
    Dim maxCnt As Long
    maxCnt = 0
    Dim k As Variant
    For Each k In yrCounts.keys
        If yrCounts(k) > maxCnt Then
            maxCnt = yrCounts(k)
            mainYear = CInt(k)
        End If
    Next k

    Dim fyStart As Date
    Dim fyEnd As Date
    If mainYear > 0 Then
        fyStart = DateSerial(mainYear, 4, 1)
        fyEnd = DateSerial(mainYear + 1, 3, 31)
    End If

    ' 一覧書き出し
    Dim outRow As Long
    outRow = hdrRow + 1

    For r = hdrRow + 1 To lastRow
        rowEmpty = True
        For c = 1 To urLastCol
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If rowEmpty Then GoTo NextR2

        isSumRow = False
        For c = 1 To urLastCol
            cv = CStr(ws.Cells(r, c).Value)
            If cv = "合計" Or cv = "計" Or cv = "総計" Then
                isSumRow = True
                Exit For
            End If
        Next c
        If isSumRow Then GoTo NextR2

        Dim cellVal As Variant
        cellVal = ws.Cells(r, selCol).Value

        Dim reason As String
        reason = ""
        Dim dispVal As Variant
        dispVal = ""

        If isEmpty(cellVal) Or セルの字(cellVal) = "" Then
            reason = "空"
            dispVal = ""
        ElseIf Not IsDate(cellVal) Then
            reason = "日付でない"
            dispVal = cellVal
        Else
            Dim dDate As Date
            dDate = CDate(cellVal)
            dispVal = Format(dDate, "yyyy/m/d")
            If dDate > Date Then
                reason = "未来の日付"
            ElseIf mainYear > 0 Then
                If dDate < fyStart Or dDate > fyEnd Then
                    reason = "年度外"
                End If
            End If
        End If

        If セルの字(reason) <> "" Then
            ws.Cells(outRow, outCol).Value = r
            ws.Cells(outRow, outCol + 1).Value = dispVal
            ws.Cells(outRow, outCol + 2).Value = reason
            outRow = outRow + 1
        End If
NextR2:
    Next r
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 2))
        Call 表の仕上げ
    End If

End Sub

Sub エラー値のセルを一覧にする()
    ' シート内の #N/A、#VALUE!、#DIV/0! などの数式エラーセルをすべて検出し、セル番地とエラー内容を一覧にする。
    ' 依頼の語: エラー値|エラーのセル|エラー値を探|DIV/0|#N/A|#VALUE|計算エラー|エラーになって|エラー値のある
    ' 依頼の組: エラー,#N/A,#DIV,#VALUE+一覧,探,洗い出
    ' 扱う: エラー値 エラー一覧 合計
    ' 見出し: なし
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urRows As Long, urCols As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRows = ur.rows.Count
    urCols = ur.Columns.Count

    ' 見出し行を探す: 空でないセルが2つ以上並ぶ最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    Dim cnt As Long
    For r = urTop To urTop + urRows - 1
        cnt = 0
        For c = urLeft To urLeft + urCols - 1
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    ' 表の左端列・右端列を見出し行で決める
    Dim tblLeft As Long, tblRight As Long
    tblLeft = 0
    For c = urLeft To urLeft + urCols - 1
        If セルの字(ws.Cells(hdrRow, c).Value) <> "" Then
            If tblLeft = 0 Then tblLeft = c
            tblRight = c
        Else
            If tblLeft <> 0 Then Exit For
        End If
    Next c
    If tblLeft = 0 Then Exit Sub

    ' 表の最後の行
    Dim tblLastRow As Long
    tblLastRow = hdrRow
    For r = urTop + urRows - 1 To hdrRow + 1 Step -1
        Dim hasVal As Boolean
        hasVal = False
        For c = tblLeft To tblRight
            If セルの字(ws.Cells(r, c).Value) <> "" Or IsError(ws.Cells(r, c).Value) Then
                hasVal = True
                Exit For
            End If
        Next c
        If hasVal Then
            tblLastRow = r
            Exit For
        End If
    Next r

    ' 表の左端の文字列列を探す
    Dim textCol As Long
    textCol = 0
    For c = tblLeft To tblRight
        Dim isTextCol As Boolean
        isTextCol = False
        For r = hdrRow + 1 To tblLastRow
            If Not IsError(ws.Cells(r, c).Value) Then
                If VarType(ws.Cells(r, c).Value) = vbString And セルの字(ws.Cells(r, c).Value) <> "" Then
                    isTextCol = True
                    Exit For
                End If
            End If
        Next r
        If isTextCol Then
            textCol = c
            Exit For
        End If
    Next c

    ' 出力列: 表の右端の2つ右
    Dim outCol As Long
    outCol = 右の出し先(ws, hdrRow, tblRight + 2, 3, "番地|見出し|項目")   ' 右に別の物があれば避ける（2026-09-23）

    ' 既存の一覧を消す（見出し行から下）
    Dim lastOutRow As Long
    lastOutRow = hdrRow
    For r = hdrRow To urTop + urRows - 1
        If セルの字(ws.Cells(r, outCol).Value) <> "" Or セルの字(ws.Cells(r, outCol + 1).Value) <> "" Or セルの字(ws.Cells(r, outCol + 2).Value) <> "" Then
            lastOutRow = r
        End If
    Next r
    If lastOutRow >= hdrRow Then
        ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(lastOutRow, outCol + 2)).ClearContents
    End If

    ' 見出しを書く
    ws.Cells(hdrRow, outCol).Value = "番地"
    ws.Cells(hdrRow, outCol + 1).Value = "見出し"
    ws.Cells(hdrRow, outCol + 2).Value = "項目"

    ' エラーセルを探して書く
    Dim outRow As Long
    outRow = hdrRow + 1
    For r = hdrRow + 1 To tblLastRow
        For c = tblLeft To tblRight
            If IsError(ws.Cells(r, c).Value) Then
                ' 番地
                ws.Cells(outRow, outCol).Value = ws.Cells(r, c).Address(False, False)
                ' 見出し
                ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
                ' 項目
                If textCol > 0 Then
                    ws.Cells(outRow, outCol + 2).Value = ws.Cells(r, textCol).Value
                End If
                outRow = outRow + 1
            End If
        Next c
    Next r
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 2))
        Call 表の仕上げ
    End If

End Sub

Sub 選んだ列の名前の揺れを一覧にする()
    ' 「株式会社」の有無や全角半角、スペースの違いによる取引先名や氏名の表記揺れをグループ化して一覧にする。
    ' 依頼の語: 名寄せ|表記揺れを一覧|表記ゆれを一覧|株式会社の有無|名寄せの候補|相手先を探|まとめる名前|種類の数|元の名前|名前の揺|同じ会社|列で同じ相手
    ' 依頼の組: 名前,取引先,相手先,会社名,業者,氏名,名の+揺れ,ゆれ,ばらつき,名寄せ,違う+-直,そろえ,揃え,統一,整え
    ' 扱う: 名寄せ 名前の揺れ 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim tblLeft As Long, tblRight As Long
    ' 左端＝見出しの行で最初に値のある列（A 列固定だと、表が C3 から始まる表で右端を取り違えた。2026-09-22）
    tblLeft = ws.UsedRange.Column
    Do While tblLeft < ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 And セルの字(ws.Cells(hdrRow, tblLeft).Value) = ""
        tblLeft = tblLeft + 1
    Loop
    tblRight = tblLeft
    Dim c As Long
    For c = tblLeft To ws.Columns.Count
        If セルの字(ws.Cells(hdrRow, c).Value) = "" Then
            tblRight = c - 1
            Exit For
        End If
        tblRight = c
    Next c

    Dim urLast As Long
    urLast = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    Dim lastRow As Long
    lastRow = hdrRow
    Dim r As Long
    For r = hdrRow + 1 To urLast
        Dim rowEmpty As Boolean
        rowEmpty = True
        For c = tblLeft To tblRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If Not rowEmpty Then lastRow = r
    Next r

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column

    Dim outCol As Long
    outCol = 右の出し先(ws, hdrRow, tblRight + 2, 3, "まとめる名前|種類の数|元の名前")   ' 右に別の物があれば避ける（2026-09-23）

    ' 前の結果は、見出しの行から下の 3 列だけ消す（列まるごと消して、同じ列の表題・注記まで消し、表題の結合セルでは 1004 で止まった。2026-09-22）
    Dim clrLast As Long
    clrLast = urLast
    If ws.Cells(ws.rows.Count, outCol).End(xlUp).Row > clrLast Then clrLast = ws.Cells(ws.rows.Count, outCol).End(xlUp).Row
    ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(clrLast, outCol + 2)).ClearContents

    ws.Cells(hdrRow, outCol).Value = "まとめる名前"
    ws.Cells(hdrRow, outCol + 1).Value = "種類の数"
    ws.Cells(hdrRow, outCol + 2).Value = "元の名前"

    Dim s As String

    Dim dictVariants As Object
    Dim dictFirst As Object
    Dim orderList() As String
    Dim orderCount As Long
    orderCount = 0
    ReDim orderList(0)
    Set dictVariants = CreateObject("Scripting.Dictionary")
    Set dictFirst = CreateObject("Scripting.Dictionary")

    For r = hdrRow + 1 To lastRow
        Dim rEmpty As Boolean
        rEmpty = True
        For c = tblLeft To tblRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                rEmpty = False
                Exit For
            End If
        Next c
        If rEmpty Then GoTo NextRow

        Dim hasSum As Boolean
        hasSum = False
        Dim cv As String
        For c = tblLeft To tblRight
            cv = CStr(ws.Cells(r, c).Value)
            If 集計の語か(cv) Then
                hasSum = True
                Exit For
            End If
        Next c
        If hasSum Then GoTo NextRow

        Dim rawVal As String
        rawVal = CStr(ws.Cells(r, selCol).Value)
        If セルの字(rawVal) = "" Then GoTo NextRow

        s = rawVal
        GoSub 正規化キー作成
        Dim nKey As String
        nKey = s

        If Not dictVariants.exists(nKey) Then
            Set dictVariants(nKey) = CreateObject("Scripting.Dictionary")
            dictFirst(nKey) = rawVal
            ReDim Preserve orderList(orderCount)
            orderList(orderCount) = nKey
            orderCount = orderCount + 1
        End If

        If Not dictVariants(nKey).exists(rawVal) Then
            dictVariants(nKey)(rawVal) = r
        End If

NextRow:
    Next r

    Dim outRow As Long
    outRow = hdrRow + 1
    Dim i As Long
    For i = 0 To orderCount - 1
        Dim k As String
        k = orderList(i)
        If dictVariants(k).Count >= 2 Then
            Dim varKeys() As Variant
            varKeys = dictVariants(k).keys
            Dim varVals() As Variant
            varVals = dictVariants(k).items
            Dim vCount As Long
            vCount = dictVariants(k).Count
            Dim j As Long, m As Long
            For j = 0 To vCount - 2
                For m = j + 1 To vCount - 1
                    If varVals(m) < varVals(j) Then
                        Dim tmpK As Variant, tmpV As Variant
                        tmpK = varKeys(j): tmpV = varVals(j)
                        varKeys(j) = varKeys(m): varVals(j) = varVals(m)
                        varKeys(m) = tmpK: varVals(m) = tmpV
                    End If
                Next m
            Next j
            Dim varStr As String
            varStr = ""
            For j = 0 To vCount - 1
                If j = 0 Then
                    varStr = CStr(varKeys(j))
                Else
                    varStr = varStr & "／" & CStr(varKeys(j))
                End If
            Next j
            ws.Cells(outRow, outCol).Value = CStr(dictFirst(k))
            ws.Cells(outRow, outCol + 1).Value = CStr(vCount)
            ws.Cells(outRow, outCol + 2).Value = varStr
            outRow = outRow + 1
        End If
    Next i
    ' 見た目は棚の共通の下請けに任せる（2026-09-23）
    If outRow > hdrRow Then
        Set m仕上げ範囲 = ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 2))
        Call 表の仕上げ
    End If

    Exit Sub

正規化キー作成:
    s = Trim$(s)
    s = Replace(s, "株式会社", "")
    s = Replace(s, "有限会社", "")
    s = Replace(s, "(株)", "")
    s = Replace(s, "（株）", "")
    s = Replace(s, "㈱", "")
    s = Replace(s, "(有)", "")
    s = Replace(s, "（有）", "")
    s = Replace(s, "㈲", "")
    s = Replace(s, Chr(32), "")
    s = Replace(s, ChrW(12288), "")
    s = StrConv(s, vbNarrow)
    s = Trim$(s)
    Return

End Sub

Sub 表のおかしい所を右に報告する()
    ' 依頼の語: 表の点検|おかしい所が|おかしいところ|不備がない|変な所がない|変なところ|問題がないか|表を点検|中身を確認|おかしい所を|点検して|問題が無いか|表の問題を|不備が無いか|おかしな所|変な所が無いか|入力ミスが起き|ミスが起きやすい|間違いが起きやすい
    ' 依頼の組: おかしい,変な,問題,不備,ミス,間違い+ないか,無いか,所,ところ,点検,チェック,見て,調べ+-数式,式,日付,マクロ,合計,グラフ
    ' 扱う: 点検 報告 空白 重複 エラー
    ' 見出し: なし
    ' 形: 数の列
    ' 表を調べて、おかしい所を「種類・場所・件数・内容」の一覧で右に報告する。直さない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim hdrRow As Long, tblLeft As Long, tblRight As Long, tblLast As Long
    Dim r As Long, c As Long, cnt As Long, outCol As Long, outRow As Long
    Dim s As String, t As String, v As Variant
    Dim n空 As Long, n数 As Long, n日 As Long, n乱 As Long, n文字 As Long, n値 As Long
    Dim nエラー As Long, n結合 As Long, n重複 As Long
    Dim キー As String, 前キー As String, 並び() As String, i As Long, j As Long

    Set ws = ActiveSheet
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If Not IsError(ws.Cells(r, c).Value) Then
                If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
            End If
        Next c
        If cnt >= 2 Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Then Exit Sub

    tblLeft = 0
    For c = urLeft To urRight
        If セルの字(ws.Cells(hdrRow, c).Value) <> "" Then
            If tblLeft = 0 Then tblLeft = c
            tblRight = c
        Else
            If tblLeft <> 0 Then Exit For
        End If
    Next c
    If tblLeft = 0 Then Exit Sub

    tblLast = hdrRow
    For r = urBottom To hdrRow + 1 Step -1
        For c = tblLeft To tblRight
            If IsError(ws.Cells(r, c).Value) Then
                tblLast = r: Exit For
            ElseIf セルの字(ws.Cells(r, c).Value) <> "" Then
                tblLast = r: Exit For
            End If
        Next c
        If tblLast = r Then Exit For
    Next r

    ' 前の点検の報告があればそこへ書き直し、右に別の物（別のマクロが作った表など）があれば避ける
    ' （2026-09-23: 項目別件数表を右に作った後に撃つと、件数表を消して報告を書いた）
    outCol = 右の出し先(ws, hdrRow, tblRight + 2, 4, "種類|場所|件数|内容")
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(urBottom + 200, outCol + 3)).ClearContents
    ws.Cells(hdrRow, outCol).Value = "種類"
    ws.Cells(hdrRow, outCol + 1).Value = "場所"
    ws.Cells(hdrRow, outCol + 2).Value = "件数"
    ws.Cells(hdrRow, outCol + 3).Value = "内容"
    ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(hdrRow, outCol + 3)).Font.Bold = True
    outRow = hdrRow + 1

    ws.Cells(outRow, outCol).Value = "表の大きさ"
    ws.Cells(outRow, outCol + 1).Value = ws.Range(ws.Cells(hdrRow, tblLeft), ws.Cells(tblLast, tblRight)).Address(0, 0)
    ws.Cells(outRow, outCol + 2).Value = tblLast - hdrRow
    ws.Cells(outRow, outCol + 3).Value = "見出し " & hdrRow & " 行目／データ " & (tblLast - hdrRow) & " 行 × " & (tblRight - tblLeft + 1) & " 列"
    outRow = outRow + 1

    ' 列ごとに調べる
    For c = tblLeft To tblRight
        n空 = 0: n数 = 0: n日 = 0: n乱 = 0: n文字 = 0: n値 = 0
        For r = hdrRow + 1 To tblLast
            v = ws.Cells(r, c).Value
            If IsError(v) Then
                nエラー = nエラー + 1
            ElseIf セルの字(v) = "" Then
                n空 = n空 + 1
            Else
                n値 = n値 + 1
                If VarType(v) = vbString Then
                    s = CStr(v): n文字 = n文字 + 1
                    t = s
                    t = Replace(t, vbCr, ""): t = Replace(t, vbLf, "")
                    t = Replace(t, vbTab, " "): t = Replace(t, ChrW(160), " ")
                    Do While Len(t) > 0
                        If Left$(t, 1) = " " Or Left$(t, 1) = ChrW(&H3000) Then t = Mid$(t, 2) Else Exit Do
                    Loop
                    Do While Len(t) > 0
                        If Right$(t, 1) = " " Or Right$(t, 1) = ChrW(&H3000) Then t = Left$(t, Len(t) - 1) Else Exit Do
                    Loop
                    If Len(t) > 0 And IsNumeric(t) And Not (Left$(t, 1) = "0" And Len(t) > 1 And Left$(t, 2) <> "0.") Then
                        n数 = n数 + 1
                    ElseIf 日付の字か(t) Then   ' 年・月・日のそろう形だけ（枝番 1-2・郵便番号・電話は日付でない・2026-09-24 総点検）
                        n日 = n日 + 1
                    ElseIf t <> s Then
                        n乱 = n乱 + 1
                    End If
                End If
            End If
        Next r

        If n空 > 0 Then
            ws.Cells(outRow, outCol).Value = "空白"
            ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
            ws.Cells(outRow, outCol + 2).Value = n空
            ws.Cells(outRow, outCol + 3).Value = "値が入っていない行がある"
            outRow = outRow + 1
        End If
        If n数 > 0 Then
            ws.Cells(outRow, outCol).Value = "文字の数字"
            ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
            ws.Cells(outRow, outCol + 2).Value = n数
            ws.Cells(outRow, outCol + 3).Value = "見た目は数字だが文字。合計できない"
            outRow = outRow + 1
        End If
        If n日 > 0 Then
            ws.Cells(outRow, outCol).Value = "文字の日付"
            ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
            ws.Cells(outRow, outCol + 2).Value = n日
            ws.Cells(outRow, outCol + 3).Value = "日付の値になっていない。並べ替えがずれる"
            outRow = outRow + 1
        End If
        If n乱 > 0 Then
            ws.Cells(outRow, outCol).Value = "隠れ空白"
            ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
            ws.Cells(outRow, outCol + 2).Value = n乱
            ws.Cells(outRow, outCol + 3).Value = "前後の空白・改行・見えない空白が混じる"
            outRow = outRow + 1
        End If
        If n文字 > 0 And n値 > n文字 And n文字 * 2 < n値 Then
            ws.Cells(outRow, outCol).Value = "型の混在"
            ws.Cells(outRow, outCol + 1).Value = ws.Cells(hdrRow, c).Value
            ws.Cells(outRow, outCol + 2).Value = n文字
            ws.Cells(outRow, outCol + 3).Value = "数値の列に文字が混じっている"
            outRow = outRow + 1
        End If
    Next c

    ' 重複行（全列一致）
    ReDim 並び(1 To tblLast - hdrRow + 1)
    j = 0
    For r = hdrRow + 1 To tblLast
        キー = ""
        For c = tblLeft To tblRight
            If IsError(ws.Cells(r, c).Value) Then
                キー = キー & vbTab & "#エラー"
            Else
                キー = キー & vbTab & CStr(ws.Cells(r, c).text)
            End If
        Next c
        j = j + 1: 並び(j) = キー
    Next r
    For i = 1 To j
        For r = i + 1 To j
            If セルの字(並び(i)) <> "" And 並び(i) = 並び(r) Then
                n重複 = n重複 + 1
                並び(r) = ""
            End If
        Next r
    Next i
    If n重複 > 0 Then
        ws.Cells(outRow, outCol).Value = "重複行"
        ws.Cells(outRow, outCol + 1).Value = "表の全体"
        ws.Cells(outRow, outCol + 2).Value = n重複
        ws.Cells(outRow, outCol + 3).Value = "すべての列が同じ行がある"
        outRow = outRow + 1
    End If

    If nエラー > 0 Then
        ws.Cells(outRow, outCol).Value = "エラー値"
        ws.Cells(outRow, outCol + 1).Value = "表の全体"
        ws.Cells(outRow, outCol + 2).Value = nエラー
        ws.Cells(outRow, outCol + 3).Value = "#N/A や #DIV/0! などのセル"
        outRow = outRow + 1
    End If

    For r = hdrRow To tblLast
        For c = tblLeft To tblRight
            If ws.Cells(r, c).MergeCells Then n結合 = n結合 + 1
        Next c
    Next r
    If n結合 > 0 Then
        ws.Cells(outRow, outCol).Value = "結合セル"
        ws.Cells(outRow, outCol + 1).Value = "表の全体"
        ws.Cells(outRow, outCol + 2).Value = n結合
        ws.Cells(outRow, outCol + 3).Value = "並べ替え・集計ができなくなる"
        outRow = outRow + 1
    End If

    If outRow = hdrRow + 2 Then
        ws.Cells(outRow, outCol).Value = "問題なし"
        ws.Cells(outRow, outCol + 3).Value = "目立つ不備は見つかりませんでした"
        outRow = outRow + 1
    End If

    ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 3)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(outRow - 1, outCol + 3)).Columns.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = "表の点検: " & (outRow - hdrRow - 2) & " 件の指摘を " & ws.Cells(hdrRow, outCol).Address(0, 0) & " から右に出しました（直してはいません）。"
End Sub

Sub 数式の危ない所を右に報告する()
    ' 依頼の語: 数式の監査|直書き|数式を監査|式が合っているか|数式が不安|計算が合わない|式のチェック|数式を調べ|式が正しいか|数式の点検|数式を確かめ
    ' 依頼の組: 数式,式+間違い,監査,チェック,正しい,合って,調べ,点検,確かめ+-一覧,説明,途切れ,消えて,抜け,崩れ,パターン,何を,株式,形式,様式,書式,正式
    ' 扱う: 数式 監査 エラー 直書き 報告
    ' 見出し: なし
    ' 形: 数の列
    ' 数式を調べて、危ない所を「種類・番地・内容」の一覧で右に報告する。直さない。
    '   エラー値／式の中の直書きの数値／隣とならびの違う式／縦計と内訳の食い違い
    Dim ws As Worksheet, ur As Range, cel As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim r As Long, c As Long, i As Long, k As Long
    Dim outCol As Long, outRow As Long
    Dim s As String, u As String, p1 As Long, p2 As Long
    Dim 式() As String, 番地() As String, 件 As Long
    Dim 多数 As String, 回数 As Long, 最多 As Long, 数式数 As Long
    Dim tot As Double, dif As Double, 正規 As Object, 欠片 As Variant

    Set ws = ActiveSheet
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    ' 正規表現: VBScript が外された環境では VBA 標準の RegExp に切り替える（2027 年対応）
    On Error Resume Next
    Set 正規 = CreateObject("VBScript.RegExp")
    If 正規 Is Nothing Then Set 正規 = 正規表現_標準版を作る()
    On Error GoTo 0
    If 正規 Is Nothing Then
        Application.StatusBar = "数式の監査: 正規表現が使えないので調べられませんでした（VBScript が無効で、この Excel の VBA にも標準の RegExp がありません）"
        Exit Sub
    End If
    正規.Global = True
    urTop = ur.Row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    ' 報告の置き場: 前の報告があればそこへ書き直す（撃ち直すたびに使用範囲の右へ報告が増え、前の報告まで調べていた・2026-09-24 総点検）。
    ' 調べるのは報告より左だけ
    outCol = 右の出し先(ws, urTop, urRight + 2, 3, "種類|番地|内容", urLeft)
    If outCol <= urRight Then urRight = outCol - 2
    If urRight < urLeft Then urRight = urLeft
    Application.ScreenUpdating = False
    If セルの字(ws.Cells(urTop, outCol).Value) = "種類" Then 前の出力を消す ws, urTop, outCol, 3
    ws.Cells(urTop, outCol).Value = "種類"
    ws.Cells(urTop, outCol + 1).Value = "番地"
    ws.Cells(urTop, outCol + 2).Value = "内容"
    ws.Range(ws.Cells(urTop, outCol), ws.Cells(urTop, outCol + 2)).Font.Bold = True
    outRow = urTop + 1

    ' 1) エラー値
    For r = urTop To urBottom
        For c = urLeft To urRight
            If IsError(ws.Cells(r, c).Value) Then
                ws.Cells(outRow, outCol).Value = "エラー値"
                ws.Cells(outRow, outCol + 1).Value = ws.Cells(r, c).Address(0, 0)
                ws.Cells(outRow, outCol + 2).Value = "'" & ws.Cells(r, c).text & "　" & ws.Cells(r, c).Formula
                outRow = outRow + 1
            End If
        Next c
    Next r

    ' 2) 式の中の直書きの数値
    For r = urTop To urBottom
        For c = urLeft To urRight
            If ws.Cells(r, c).HasFormula Then
                数式数 = 数式数 + 1
                ' 式（A1 形式）から文字列・シート名・番地を消し、演算子とかっこで切った欠片のうち数だけのものを直書きと見る。
                ' 0・1・100・1000（単位の換算）は数えない（2026-09-24 総点検: 数字を 1 字ずつ拾っていて、$B$2・ROUND(,2) を
                ' 直書きと言い、×1.1 の税率を見逃していた）
                s = ws.Cells(r, c).Formula
                GoSub 直書きの数
                If Len(u) > 0 Then
                    ws.Cells(outRow, outCol).Value = "直書きの数値"
                    ws.Cells(outRow, outCol + 1).Value = ws.Cells(r, c).Address(0, 0)
                    ws.Cells(outRow, outCol + 2).Value = "'" & ws.Cells(r, c).Formula & "　（" & Mid$(u, 2) & " を式に直接書いている。変わったとき直し漏れる）"
                    outRow = outRow + 1
                End If
            End If
        Next c
    Next r

    ' 3) 隣とならびの違う式（列ごとに多数派と違うもの）
    For c = urLeft To urRight
        件 = 0
        ReDim 式(1 To urBottom - urTop + 1): ReDim 番地(1 To urBottom - urTop + 1)
        For r = urTop To urBottom
            If ws.Cells(r, c).HasFormula Then
                件 = 件 + 1
                式(件) = ws.Cells(r, c).FormulaR1C1
                番地(件) = ws.Cells(r, c).Address(0, 0)
            End If
        Next r
        If 件 >= 3 Then
            多数 = "": 最多 = 0
            For i = 1 To 件
                回数 = 0
                For k = 1 To 件
                    If 式(k) = 式(i) Then 回数 = 回数 + 1
                Next k
                If 回数 > 最多 Then 最多 = 回数: 多数 = 式(i)
            Next i
            If 最多 * 2 > 件 Then
                For i = 1 To 件
                    ' 合計の式は形が違って当たり前なので除く
                    If 式(i) <> 多数 And InStr(UCase$(式(i)), "SUM(") = 0 And InStr(UCase$(式(i)), "SUBTOTAL(") = 0 Then
                        ws.Cells(outRow, outCol).Value = "ならびの違う式"
                        ws.Cells(outRow, outCol + 1).Value = 番地(i)
                        ws.Cells(outRow, outCol + 2).Value = "'" & ws.Range(番地(i)).Formula & "　（この列の他の行と形が違う）"
                        outRow = outRow + 1
                    End If
                Next i
            End If
        End If
    Next c

    ' 4) 縦計と内訳の食い違い
    For r = urTop To urBottom
        For c = urLeft To urRight
            Set cel = ws.Cells(r, c)
            If cel.HasFormula Then
                If InStr(UCase$(cel.FormulaR1C1), "SUM(") > 0 And InStr(cel.FormulaR1C1, "R[-") > 0 Then
                    If Not IsError(cel.Value) Then
                        tot = 0: k = r - 1
                        Do While k >= urTop
                            If IsError(ws.Cells(k, c).Value) Then Exit Do
                            If セルの字(ws.Cells(k, c).Value) = "" Then Exit Do
                            If Not IsNumeric(ws.Cells(k, c).Value) Then Exit Do
                            If ws.Cells(k, c).HasFormula Then
                                If InStr(UCase$(ws.Cells(k, c).FormulaR1C1), "SUM(") > 0 Then Exit Do
                            End If
                            tot = tot + CDbl(ws.Cells(k, c).Value)
                            k = k - 1
                        Loop
                        dif = CDbl(cel.Value) - tot
                        ' すぐ上が空行（合計の上の区切り）で 1 行も足せなかったときは比べない（「上の 0 行の和 0」と言っていた・2026-09-24）
                        If Abs(dif) > 0.005 And k < r - 1 Then
                            ws.Cells(outRow, outCol).Value = "合計のずれ"
                            ws.Cells(outRow, outCol + 1).Value = cel.Address(0, 0)
                            s = Format$(tot, "#,##0.##"): If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
                            u = Format$(cel.Value, "#,##0.##"): If Right$(u, 1) = "." Then u = Left$(u, Len(u) - 1)
                            多数 = Format$(dif, "#,##0.##"): If Right$(多数, 1) = "." Then 多数 = Left$(多数, Len(多数) - 1)
                            ws.Cells(outRow, outCol + 2).Value = "上の " & (r - 1 - k) & " 行の和 " & s & " と " & u & " が " & 多数 & " 違う（範囲の取りこぼし）"
                            outRow = outRow + 1
                        End If
                    End If
                End If
            End If
        Next c
    Next r

    If outRow = urTop + 1 Then
        ws.Cells(outRow, outCol).Value = "問題なし"
        ws.Cells(outRow, outCol + 2).Value = "数式 " & 数式数 & " 本を調べましたが、危ない所は見つかりませんでした"
        outRow = outRow + 1
    End If

    ws.Range(ws.Cells(urTop, outCol), ws.Cells(outRow - 1, outCol + 2)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(urTop, outCol), ws.Cells(outRow - 1, outCol + 2)).Columns.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = "数式の監査: 数式 " & 数式数 & " 本のうち " & (outRow - urTop - 1) & " 件を " & ws.Cells(urTop, outCol).Address(0, 0) & " から右に出しました（直してはいません）。"
    Exit Sub

直書きの数:
    ' s（A1 形式の式）の中の直書きの数を u に「、1.1、25」の形で（無ければ ""）
    u = ""
    正規.Pattern = """[^""]*"""
    s = 正規.Replace(s, " ")
    正規.Pattern = "'[^']*'!"
    s = 正規.Replace(s, " ")
    正規.Pattern = "[^\s=+\-*/^&,;:<>()!{}%']+!"
    s = 正規.Replace(s, " ")
    正規.Pattern = "\$?[A-Za-z]{1,3}\$?\d+"
    s = 正規.Replace(s, " ")
    正規.Pattern = "\$?\d+:\$?\d+"
    s = 正規.Replace(s, " ")
    正規.Pattern = "[=+\-*/^&,;:<>(){}%]"
    s = 正規.Replace(s, " ")
    For Each 欠片 In Split(s, " ")
        If 欠片 Like "#*" Or 欠片 Like ".#*" Then
            If IsNumeric(欠片) Then
                If CDbl(欠片) <> 0 And CDbl(欠片) <> 1 And CDbl(欠片) <> 100 And CDbl(欠片) <> 1000 Then u = u & "、" & 欠片
            End If
        End If
    Next 欠片
    Return
End Sub

Sub 引き継ぎの説明シートを作る()
    ' 依頼の語: 引き継ぎ|引継ぎ|人に渡す前|初めての人|使い方をまとめ|ブックの説明|渡す準備|どこに入力するか|説明書を作
    ' 依頼の組: 引き継,引継,後任,初めての人+説明,資料,まとめ,作
    ' 扱う: 引き継ぎ 説明 入力欄 数式
    ' 見出し: なし
    ' 形: 数の列
    ' シートごとに「どこに入力してどこが自動計算か」「触ってはいけない所」を調べ、
    ' 新しいシート「このブックの説明」にまとめる。使い方の欄は空けておくので人が書き足す。
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, ur As Range, 式域 As Range
    Dim i As Long, r As Long, c As Long, cnt As Long, outRow As Long, 名 As String
    Dim hdrRow As Long, n式 As Long, n手 As Long, 入力列 As String, 計算列 As String
    Dim 式範囲 As String, 列名 As String, リンク As Variant

    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False

    名 = "このブックの説明"
    i = 1
    Do While i < 100
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(名)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        名 = "このブックの説明" & i: i = i + 1
    Loop
    Set sh = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    sh.Name = 名

    sh.Cells(1, 1).Value = "このブックの説明（" & wb.Name & "　" & Format$(Now, "yyyy/m/d") & " 作成）"
    sh.Cells(1, 1).Font.Bold = True
    sh.Cells(1, 1).Font.Size = 14
    sh.Cells(3, 1).Value = "シート"
    sh.Cells(3, 2).Value = "表の範囲"
    sh.Cells(3, 3).Value = "行数"
    sh.Cells(3, 4).Value = "入力する列（手で入れる）"
    sh.Cells(3, 5).Value = "自動で計算する列（触らない）"
    sh.Cells(3, 6).Value = "数式のある場所"
    sh.Range(sh.Cells(3, 1), sh.Cells(3, 6)).Font.Bold = True
    outRow = 4

    For Each ws In wb.Worksheets
        If ws.Name <> 名 Then
            Set ur = Nothing
            On Error Resume Next
            Set ur = ws.UsedRange
            On Error GoTo 0
            If ur Is Nothing Then GoTo 次のシート

            hdrRow = 0
            For r = ur.Row To ur.Row + ur.rows.Count - 1
                cnt = 0
                For c = ur.Column To ur.Column + ur.Columns.Count - 1
                    If Not IsError(ws.Cells(r, c).Value) Then
                        If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
                    End If
                Next c
                If cnt >= 2 Then hdrRow = r: Exit For
            Next r

            入力列 = "": 計算列 = ""
            If hdrRow > 0 And hdrRow < ur.Row + ur.rows.Count - 1 Then
                For c = ur.Column To ur.Column + ur.Columns.Count - 1
                    n式 = 0: n手 = 0
                    For r = hdrRow + 1 To ur.Row + ur.rows.Count - 1
                        If ws.Cells(r, c).HasFormula Then
                            n式 = n式 + 1
                        ElseIf Not IsError(ws.Cells(r, c).Value) Then
                            If セルの字(ws.Cells(r, c).Value) <> "" Then n手 = n手 + 1
                        End If
                    Next r
                    列名 = CStr(ws.Cells(hdrRow, c).Value)
                    If セルの字(列名) = "" Then 列名 = "第" & (c - ur.Column + 1) & "列"
                    If n式 > 0 And n式 >= n手 Then
                        計算列 = 計算列 & "／" & 列名
                    ElseIf n手 > 0 Then
                        入力列 = 入力列 & "／" & 列名
                    End If
                Next c
            End If

            式範囲 = ""
            Set 式域 = Nothing
            On Error Resume Next
            Set 式域 = ur.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not 式域 Is Nothing Then
                式範囲 = 式域.Address(0, 0)
                If Len(式範囲) > 120 Then 式範囲 = Left$(式範囲, 120) & " …"
            Else
                式範囲 = "なし"
            End If

            sh.Cells(outRow, 1).Value = ws.Name
            sh.Cells(outRow, 2).Value = ur.Address(0, 0)
            sh.Cells(outRow, 3).Value = IIf(hdrRow > 0, ur.Row + ur.rows.Count - 1 - hdrRow, 0)
            sh.Cells(outRow, 4).Value = IIf(入力列 = "", "（なし）", Mid$(入力列, 2))
            sh.Cells(outRow, 5).Value = IIf(計算列 = "", "（なし）", Mid$(計算列, 2))
            sh.Cells(outRow, 6).Value = "'" & 式範囲
            outRow = outRow + 1
        End If
次のシート:
    Next ws

    sh.Range(sh.Cells(3, 1), sh.Cells(outRow - 1, 6)).Borders.LineStyle = xlContinuous
    outRow = outRow + 1

    sh.Cells(outRow, 1).Value = "気をつけること"
    sh.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "・「自動で計算する列」と「数式のある場所」は、手で数字を打ち込まないでください（式が消えます）"
    outRow = outRow + 1
    リンク = wb.LinkSources(xlExcelLinks)
    If isEmpty(リンク) Then
        sh.Cells(outRow, 1).Value = "・外部のブックへのリンクはありません"
    Else
        sh.Cells(outRow, 1).Value = "・外部のブックへのリンクが " & (UBound(リンク) - LBound(リンク) + 1) & " 本あります（リンク元が動くと数字が変わります）"
    End If
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "・名前の定義が " & wb.names.Count & " 本あります"
    outRow = outRow + 2

    sh.Cells(outRow, 1).Value = "使い方（ここは人が書いてください）"
    sh.Cells(outRow, 1).Font.Bold = True
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "1. 毎月いつ・何をするか："
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "2. どこから数字をもらうか："
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "3. できたものを誰に渡すか："
    outRow = outRow + 1
    sh.Cells(outRow, 1).Value = "4. 困ったときの連絡先："

    sh.Columns("A:F").AutoFit
    For c = 1 To 6
        If sh.Columns(c).ColumnWidth > 45 Then sh.Columns(c).ColumnWidth = 45
    Next c
    sh.Activate
    sh.Cells(1, 1).Select
    Application.ScreenUpdating = True
    Application.StatusBar = "引き継ぎの説明: シート「" & 名 & "」を作りました。使い方の4つの欄は人が書き足してください（置き場所も人に確認を）。"
End Sub

Sub マクロの呼び出し関係を一覧にする()
    ' 依頼の語: 呼び出し関係|呼び出し元|どこから呼ばれ|影響範囲|どこまで影響|直したときの影響|呼んでいるマクロ|呼ばれているマクロ
    ' 依頼の組: 呼び出,呼ばれ,影響+マクロ,元+-使われていない
    ' 扱う: マクロ 呼び出し 影響
    ' 見出し: なし
    ' 形: なし
    ' Excelコンボの［マクロ］タブで選んだマクロ（選んでいなければ名前を聞く）について、そのマクロが呼んでいるマクロ・そのマクロを呼んでいる所（マクロ・ボタン）を、
    ' 新しいシート「調査_呼び出し関係」に一覧にする。直すとどこまで影響するかの見当に使う。何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object, comp As Object, cm As Object, shp As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim 対象 As String, n As Long, jj As Long, proc As String, nxt As Long, k As Long, i As Long, tgt As Long
    Dim モジュ() As String, 名() As String, 始() As Long, 数() As Long, 本文() As String, 本数 As Long, 種() As Long
    Dim 行 As Variant, ln As String, p As Long, 前 As String, 後 As String, 見つけた As Boolean, act As String, nc1 As Long, nc2 As Long
    Dim 境 As String
    On Error GoTo 失敗
    境 = " ()=,.:""'&+-*/<>[]{}!?" & vbTab
    Set wb = ActiveWorkbook
    対象 = ""
    On Error Resume Next
    対象 = Application.Run("'" & ThisWorkbook.Name & "'!コンボ道具.選択マクロを返す")
    On Error GoTo 失敗
    If Len(対象) = 0 Then 対象 = Trim$(InputBox("調べるマクロの名前を入れてください。" & vbCrLf & "（Excelコンボの［マクロ］タブで選んでおくと、入力は要りません）", "マクロの呼び出し関係"))
    If Len(対象) = 0 Then Exit Sub
    n = 0
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "マクロの呼び出し関係: VBA プロジェクトを読めません（マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    ReDim モジュ(0 To 3000): ReDim 名(0 To 3000): ReDim 始(0 To 3000): ReDim 数(0 To 3000): ReDim 本文(0 To 3000): ReDim 種(0 To 3000)
    本数 = 0
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        jj = 1
        Do While jj <= n And 本数 < 3000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                始(本数) = cm.ProcStartLine(proc, 0)
                数(本数) = cm.ProcCountLines(proc, 0)
                モジュ(本数) = comp.Name
                名(本数) = proc
                種(本数) = comp.Type
                本文(本数) = cm.lines(始(本数), 数(本数))
                nxt = 始(本数) + 数(本数)
                本数 = 本数 + 1
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    tgt = -1
    For i = 0 To 本数 - 1
        If LCase$(名(i)) = LCase$(対象) Then tgt = i: Exit For
    Next i
    If tgt < 0 Then
        Application.StatusBar = "マクロの呼び出し関係: 「" & 対象 & "」というマクロは見つかりませんでした。"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    nm0 = "調査_呼び出し関係": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：マクロの呼び出し関係　" & 名(tgt) & "（" & wb.Name & "）"
    out.Range("A2").Value = "「呼んでいる」＝このマクロの中で使われている他のマクロ／「呼ばれている」＝このマクロ名を使っている所。コメント行は数えません。"
    hr = 4: NC = 5
    out.Range("A4:E4").Value = Array("種類", "モジュール", "マクロ", "行", "その行")
    r = hr
    行 = Split(本文(tgt), vbCrLf)
    For i = 0 To 本数 - 1
        If i <> tgt And Len(名(i)) >= 3 Then
            見つけた = False
            For k = 0 To UBound(行)
                ln = Trim$(CStr(行(k)))
                If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
                    p = InStr(1, ln, 名(i), 1)
                    Do While p > 0
                        前 = " "
                        If p > 1 Then 前 = Mid$(ln, p - 1, 1)
                        後 = Mid$(ln, p + Len(名(i)), 1)
                        If Len(後) = 0 Then 後 = " "
                        If InStr(境, 前) > 0 And InStr(境, 後) > 0 Then
                            見つけた = True
                            Exit Do
                        End If
                        p = InStr(p + 1, ln, 名(i), 1)
                    Loop
                    If 見つけた Then Exit For
                End If
            Next k
            If 見つけた Then
                r = r + 1: nc1 = nc1 + 1
                out.Cells(r, 1).Value = "呼んでいる"
                out.Cells(r, 2).Value = "'" & モジュ(i)
                out.Cells(r, 3).Value = "'" & 名(i)
                out.Cells(r, 4).Value = 始(tgt) + k
                out.Cells(r, 5).Value = "'" & Left$(ln, 100)
            End If
        End If
    Next i
    For i = 0 To 本数 - 1
        If i <> tgt Then
            行 = Split(本文(i), vbCrLf)
            見つけた = False
            For k = 0 To UBound(行)
                ln = Trim$(CStr(行(k)))
                If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
                    p = InStr(1, ln, 名(tgt), 1)
                    Do While p > 0
                        前 = " "
                        If p > 1 Then 前 = Mid$(ln, p - 1, 1)
                        後 = Mid$(ln, p + Len(名(tgt)), 1)
                        If Len(後) = 0 Then 後 = " "
                        If InStr(境, 前) > 0 And InStr(境, 後) > 0 Then
                            見つけた = True
                            Exit Do
                        End If
                        p = InStr(p + 1, ln, 名(tgt), 1)
                    Loop
                    If 見つけた Then Exit For
                End If
            Next k
            If 見つけた Then
                r = r + 1: nc2 = nc2 + 1
                out.Cells(r, 1).Value = "呼ばれている"
                out.Cells(r, 2).Value = "'" & モジュ(i)
                out.Cells(r, 3).Value = "'" & 名(i)
                out.Cells(r, 4).Value = 始(i) + k
                out.Cells(r, 5).Value = "'" & Left$(ln, 100)
            End If
        End If
    Next i
    For Each sh In wb.Worksheets
        For Each shp In sh.Shapes
            act = ""
            On Error Resume Next
            act = shp.OnAction
            On Error GoTo 失敗
            If Len(act) > 0 Then
                If InStr(1, act, 名(tgt), 1) > 0 Then
                    r = r + 1: nc2 = nc2 + 1
                    out.Cells(r, 1).Value = "図形に割り当て"
                    out.Cells(r, 2).Value = "'" & sh.Name
                    out.Cells(r, 3).Value = "'" & shp.Name
                    out.Cells(r, 5).Value = "'" & act
                End If
            End If
        Next shp
    Next sh
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "他のマクロとのつながりは見つかりませんでした（ボタン・ショートカット・リボンから直接使われているだけかもしれません）"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_呼び出し関係: 「" & 名(tgt) & "」が呼んでいる " & nc1 & " 本・呼ばれている " & nc2 & " か所。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "マクロの呼び出し関係: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub マクロが何を変えるかを一覧にする()
    ' 依頼の語: マクロが何を|マクロは何を|何が変わる|実行すると何が|マクロの中身|マクロの中を|先に教えて|実行したら何が
    ' 依頼の組: マクロ+何を,何が,中身+-一覧
    ' 扱う: マクロ 中身 変更
    ' 見出し: なし
    ' 形: なし
    ' Excelコンボの［マクロ］タブで選んだマクロ（選んでいなければ名前を聞く）の、頭のコメント・呼んでいるマクロ・触っているシート・出すメッセージ・
    ' 変わる可能性のある命令（削除・書き込み・保存など）を、行番号つきで新しいシート「調査_マクロの中身」に一覧にする。実行前の見当に使う。何も変えない・実行もしない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object, comp As Object, cm As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim 対象 As String, n As Long, jj As Long, proc As String, nxt As Long, k As Long, i As Long, tgt As Long
    Dim モジュ() As String, 名() As String, 始() As Long, 数() As Long, 本文() As String, 本数 As Long, 種() As Long
    Dim 行 As Variant, ln As String, p As Long, 前 As String, 後 As String, 見つけた As Boolean, w As Variant, ラベル As Variant, m As Long
    Dim 境 As String, t As String, 数変 As Long, 命令 As Variant, 名前 As String
    On Error GoTo 失敗
    境 = " ()=,.:""'&+-*/<>[]{}!?" & vbTab
    Set wb = ActiveWorkbook
    対象 = ""
    On Error Resume Next
    対象 = Application.Run("'" & ThisWorkbook.Name & "'!コンボ道具.選択マクロを返す")
    On Error GoTo 失敗
    If Len(対象) = 0 Then 対象 = Trim$(InputBox("調べるマクロの名前を入れてください。" & vbCrLf & "（Excelコンボの［マクロ］タブで選んでおくと、入力は要りません）", "マクロの中身"))
    If Len(対象) = 0 Then Exit Sub
    n = 0
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "マクロの中身: VBA プロジェクトを読めません（マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    ReDim モジュ(0 To 3000): ReDim 名(0 To 3000): ReDim 始(0 To 3000): ReDim 数(0 To 3000): ReDim 本文(0 To 3000): ReDim 種(0 To 3000)
    本数 = 0
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        jj = 1
        Do While jj <= n And 本数 < 3000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                始(本数) = cm.ProcStartLine(proc, 0)
                数(本数) = cm.ProcCountLines(proc, 0)
                モジュ(本数) = comp.Name
                名(本数) = proc
                種(本数) = comp.Type
                本文(本数) = cm.lines(始(本数), 数(本数))
                nxt = 始(本数) + 数(本数)
                本数 = 本数 + 1
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    tgt = -1
    For i = 0 To 本数 - 1
        If LCase$(名(i)) = LCase$(対象) Then tgt = i: Exit For
    Next i
    If tgt < 0 Then
        Application.StatusBar = "マクロの中身: 「" & 対象 & "」というマクロは見つかりませんでした。"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    nm0 = "調査_マクロの中身": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：マクロの中身　" & 名(tgt) & "（" & wb.Name & "）"
    out.Range("A2").Value = "コードを読んで拾った目安です（動かしていません）。変わる可能性のある命令は、実際に動くかどうかまでは分かりません。"
    hr = 4: NC = 4
    out.Range("A4:D4").Value = Array("項目", "行", "内容", "メモ")
    r = hr
    行 = Split(本文(tgt), vbCrLf)
    r = r + 1: out.Cells(r, 1).Value = "モジュール": out.Cells(r, 3).Value = "'" & モジュ(tgt)
    r = r + 1: out.Cells(r, 1).Value = "行数": out.Cells(r, 3).Value = 数(tgt) & " 行（" & 始(tgt) & "?" & (始(tgt) + 数(tgt) - 1) & " 行目）"
    For k = 0 To UBound(行)
        t = Trim$(CStr(行(k)))
        If Left$(t, 1) <> "'" And Len(t) > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "宣言"
            out.Cells(r, 2).Value = 始(tgt) + k
            out.Cells(r, 3).Value = "'" & Left$(t, 120)
            Exit For
        End If
    Next k
    m = 0
    For k = 0 To UBound(行)
        t = Trim$(CStr(行(k)))
        If Left$(t, 1) = "'" Then
            t = Trim$(Mid$(t, 2))
            If Left$(t, 5) <> "依頼の語:" And Left$(t, 3) <> "扱う:" And Left$(t, 4) <> "見出し:" And Left$(t, 2) <> "形:" And Left$(t, 4) <> "選ぶ列:" And Len(t) > 0 Then
                m = m + 1
                If m <= 12 Then
                    r = r + 1
                    out.Cells(r, 1).Value = "コメント"
                    out.Cells(r, 2).Value = 始(tgt) + k
                    out.Cells(r, 3).Value = "'" & Left$(t, 120)
                End If
            End If
        End If
    Next k
    For i = 0 To 本数 - 1
        If i <> tgt And Len(名(i)) >= 3 Then
            見つけた = False
            For k = 0 To UBound(行)
                ln = Trim$(CStr(行(k)))
                If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
                    p = InStr(1, ln, 名(i), 1)
                    Do While p > 0
                        前 = " "
                        If p > 1 Then 前 = Mid$(ln, p - 1, 1)
                        後 = Mid$(ln, p + Len(名(i)), 1)
                        If Len(後) = 0 Then 後 = " "
                        If InStr(境, 前) > 0 And InStr(境, 後) > 0 Then
                            見つけた = True
                            Exit Do
                        End If
                        p = InStr(p + 1, ln, 名(i), 1)
                    Loop
                    If 見つけた Then Exit For
                End If
            Next k
            If 見つけた Then
                r = r + 1
                out.Cells(r, 1).Value = "呼んでいるマクロ"
                out.Cells(r, 2).Value = 始(tgt) + k
                out.Cells(r, 3).Value = "'" & 名(i)
                out.Cells(r, 4).Value = "'" & モジュ(i)
            End If
        End If
    Next i
    For Each sh In wb.Worksheets
        見つけた = False
        For k = 0 To UBound(行)
            ln = Trim$(CStr(行(k)))
            If Left$(ln, 1) <> "'" Then
                If InStr(1, ln, """" & sh.Name & """", 1) > 0 Then
                    見つけた = True
                    Exit For
                End If
            End If
        Next k
        If 見つけた Then
            r = r + 1
            out.Cells(r, 1).Value = "触っているシート"
            out.Cells(r, 2).Value = 始(tgt) + k
            out.Cells(r, 3).Value = "'" & sh.Name
        End If
    Next sh
    For k = 0 To UBound(行)
        ln = Trim$(CStr(行(k)))
        If Left$(ln, 1) <> "'" Then
            If InStr(1, ln, "MsgBox", 1) > 0 Or InStr(1, ln, "InputBox", 1) > 0 Then
                r = r + 1
                out.Cells(r, 1).Value = "画面に出すメッセージ"
                out.Cells(r, 2).Value = 始(tgt) + k
                out.Cells(r, 3).Value = "'" & Left$(ln, 120)
            End If
        End If
    Next k
    命令 = Array(".Delete", "削除・消去", ".Clear", "削除・消去", ".ClearContents", "削除・消去", ".Insert", "挿入", ".Value =", "書き込み", ".Value2 =", "書き込み", ".Formula =", "書き込み", ".FormulaR1C1 =", "書き込み", ".Copy", "コピー", ".PasteSpecial", "貼り付け", ".Cut", "切り取り", ".Sort", "並べ替え", ".Merge", "結合", ".UnMerge", "結合解除", ".Hidden =", "表示・非表示", ".Visible =", "表示・非表示", ".Save", "保存", ".SaveAs", "保存", ".Close", "閉じる", "Kill ", "ファイル削除", "Shell", "外部プログラム", "SendKeys", "キー送信", "Application.Quit", "Excel終了", "DisplayAlerts = False", "警告を出さない")
    数変 = 0
    For k = 0 To UBound(行)
        ln = Trim$(CStr(行(k)))
        If Left$(ln, 1) <> "'" And Len(ln) > 0 And 数変 < 300 Then
            For j = 0 To UBound(命令) Step 2
                If InStr(1, ln, 命令(j), 1) > 0 Then
                    r = r + 1: 数変 = 数変 + 1
                    out.Cells(r, 1).Value = "変わる可能性"
                    out.Cells(r, 2).Value = 始(tgt) + k
                    out.Cells(r, 3).Value = "'" & Left$(ln, 120)
                    out.Cells(r, 4).Value = 命令(j + 1)
                    Exit For
                End If
            Next j
        End If
    Next k
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_マクロの中身: 「" & 名(tgt) & "」を調べました（変わる可能性のある命令 " & 数変 & " か所）。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "マクロの中身: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 使われていないマクロを一覧にする()
    ' 依頼の語: 使われていないマクロ|使っていないマクロ|重複しているマクロ|重複マクロ|不要なマクロ|どこからも呼ばれ|使われていない
    ' 依頼の組: マクロ+使われていない,使っていない,いらない,不要,重複
    ' 扱う: マクロ 未使用 重複
    ' 見出し: なし
    ' 形: なし
    ' 開いているブックのマクロのうち、他のマクロからもボタンからも呼ばれていないもの（使われていない候補）と、中身が同じマクロを、新しいシート「調査_使われていないマクロ」に一覧にする。
    ' ショートカットキーやリボンから直接使っているものは見分けられないので「候補」。何も変えない・消さない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object, comp As Object, cm As Object, shp As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim n As Long, jj As Long, proc As String, nxt As Long, k As Long, i As Long, a As Long
    Dim モジュ() As String, 名() As String, 始() As Long, 数() As Long, 本文() As String, 本数 As Long, 種() As Long
    Dim 行 As Variant, ln As String, p As Long, 前 As String, 後 As String, 総 As Long, 自 As Long, 全 As String, 自本 As String
    Dim 境 As String, act As String, 図 As String, 鍵 As Object, 正 As String, kv As Variant, 候補 As Long, 重複 As Long, decl As String, 組 As Variant
    On Error GoTo 失敗
    境 = " ()=,.:""'&+-*/<>[]{}!?" & vbTab & vbLf & vbCr
    Set wb = ActiveWorkbook
    n = 0
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "使われていないマクロ: VBA プロジェクトを読めません（マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    ReDim モジュ(0 To 3000): ReDim 名(0 To 3000): ReDim 始(0 To 3000): ReDim 数(0 To 3000): ReDim 本文(0 To 3000): ReDim 種(0 To 3000)
    本数 = 0
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        jj = 1
        Do While jj <= n And 本数 < 3000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                始(本数) = cm.ProcStartLine(proc, 0)
                数(本数) = cm.ProcCountLines(proc, 0)
                モジュ(本数) = comp.Name
                名(本数) = proc
                種(本数) = comp.Type
                本文(本数) = cm.lines(始(本数), 数(本数))
                nxt = 始(本数) + 数(本数)
                本数 = 本数 + 1
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    ' 全部のコード（コメント行を除く・小文字）を 1 本の文字にする。1 行ずつ & でつなぐと大きなブックで極端に遅いので、配列にして Join する
    Dim 片() As String
    ReDim 片(0 To 本数)
    For i = 0 To 本数 - 1
        行 = Split(本文(i), vbCrLf)
        For k = 0 To UBound(行)
            ln = Trim$(CStr(行(k)))
            If Len(ln) > 0 And Left$(ln, 1) <> "'" Then 行(k) = LCase$(ln) Else 行(k) = ""
        Next k
        片(i) = Join(行, vbLf)
    Next i
    全 = Join(片, vbLf)
    For Each sh In wb.Worksheets
        For Each shp In sh.Shapes
            act = ""
            On Error Resume Next
            act = shp.OnAction
            On Error GoTo 失敗
            If Len(act) > 0 Then 図 = 図 & LCase$(act) & vbLf
        Next shp
    Next sh
    Application.ScreenUpdating = False
    nm0 = "調査_使われていないマクロ": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：使われていないマクロ・中身が同じマクロ　" & wb.Name
    out.Range("A2").Value = "「使われていない候補」は、他のマクロ・図形のボタンから呼ばれていないもの。ショートカットキー・リボン・Alt+F8 で直接使っているものは見分けられません。"
    hr = 4: NC = 6
    out.Range("A4:F4").Value = Array("種類", "モジュール", "マクロ", "行数", "公開", "メモ")
    r = hr
    Set 鍵 = CreateObject("Scripting.Dictionary")
    For i = 0 To 本数 - 1
        行 = Split(本文(i), vbCrLf)
        自本 = ""
        正 = ""
        decl = ""
        For k = 0 To UBound(行)
            ln = Trim$(CStr(行(k)))
            If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
                自本 = 自本 & LCase$(ln) & vbLf
                If Len(decl) = 0 Then
                    decl = ln
                Else
                    正 = 正 & LCase$(ln) & vbLf
                End If
            End If
        Next k
        If Len(正) > 40 Then
            If 鍵.exists(正) Then
                鍵(正) = 鍵(正) & "|" & i
            Else
                鍵.Add 正, CStr(i)
            End If
        End If
        総 = 0: 自 = 0
        p = InStr(1, 全, LCase$(名(i)), 0)
        Do While p > 0
            前 = vbLf
            If p > 1 Then 前 = Mid$(全, p - 1, 1)
            後 = Mid$(全, p + Len(名(i)), 1)
            If Len(後) = 0 Then 後 = vbLf
            If InStr(境, 前) > 0 And InStr(境, 後) > 0 Then 総 = 総 + 1
            p = InStr(p + 1, 全, LCase$(名(i)), 0)
        Loop
        p = InStr(1, 自本, LCase$(名(i)), 0)
        Do While p > 0
            前 = vbLf
            If p > 1 Then 前 = Mid$(自本, p - 1, 1)
            後 = Mid$(自本, p + Len(名(i)), 1)
            If Len(後) = 0 Then 後 = vbLf
            If InStr(境, 前) > 0 And InStr(境, 後) > 0 Then 自 = 自 + 1
            p = InStr(p + 1, 自本, LCase$(名(i)), 0)
        Loop
        If 総 - 自 <= 0 And InStr(図, LCase$(名(i))) = 0 Then
            If Not (InStr(名(i), "_") > 0 And 種(i) <> 1) Then
                r = r + 1: 候補 = 候補 + 1
                out.Cells(r, 1).Value = "使われていない候補"
                out.Cells(r, 2).Value = "'" & モジュ(i)
                out.Cells(r, 3).Value = "'" & 名(i)
                out.Cells(r, 4).Value = 数(i)
                If LCase$(Left$(decl, 7)) = "private" Then out.Cells(r, 5).Value = "非公開" Else out.Cells(r, 5).Value = "公開"
                out.Cells(r, 6).Value = "どこからも呼ばれていない"
            End If
        End If
    Next i
    For Each kv In 鍵.keys
        If InStr(鍵(kv), "|") > 0 Then
            組 = Split(鍵(kv), "|")
            For a = 0 To UBound(組)
                r = r + 1: 重複 = 重複 + 1
                out.Cells(r, 1).Value = "中身が同じ"
                out.Cells(r, 2).Value = "'" & モジュ(CLng(組(a)))
                out.Cells(r, 3).Value = "'" & 名(CLng(組(a)))
                out.Cells(r, 4).Value = 数(CLng(組(a)))
                out.Cells(r, 6).Value = "同じ中身のマクロ " & (UBound(組) + 1) & " 本のうちの 1 本"
            Next a
        End If
    Next kv
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "使われていない候補も、中身が同じマクロも見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_使われていないマクロ: 使われていない候補 " & 候補 & " 本・中身が同じ " & 重複 & " 本。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "使われていないマクロ: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub このシートを使うマクロの一覧()
    ' 依頼の語: このシートを使っているマクロ|シートを使っているマクロ|シートを参照しているマクロ|このシートに触れる|シートを触るマクロ|このシートのマクロ
    ' 依頼の組: シート+マクロ+使って,参照,触+-一覧,対応
    ' 扱う: マクロ シート 参照
    ' 見出し: なし
    ' 形: なし
    ' 今のシートの名前（と、VBA 上のオブジェクト名）を使っているマクロを、行数の多い順に新しいシート「調査_シートを使うマクロ」に一覧にする。何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object, comp As Object, cm As Object, ws As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim n As Long, jj As Long, proc As String, nxt As Long, k As Long, i As Long
    Dim モジュ() As String, 名() As String, 始() As Long, 数() As Long, 本文() As String, 本数 As Long, 種() As Long
    Dim 行 As Variant, ln As String, p As Long, 前 As String, 後 As String, 境 As String, シート名 As String, コード名 As String
    Dim 件 As Long, 初 As Long, 初行 As String, 合計 As Long, ヒット As Boolean
    On Error GoTo 失敗
    境 = " ()=,.:""'&+-*/<>[]{}!?" & vbTab
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    シート名 = ws.Name
    コード名 = ""
    On Error Resume Next
    コード名 = ws.CodeName
    On Error GoTo 失敗
    n = 0
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "このシートを使っているマクロ: VBA プロジェクトを読めません（マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    ReDim モジュ(0 To 3000): ReDim 名(0 To 3000): ReDim 始(0 To 3000): ReDim 数(0 To 3000): ReDim 本文(0 To 3000): ReDim 種(0 To 3000)
    本数 = 0
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        jj = 1
        Do While jj <= n And 本数 < 3000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                始(本数) = cm.ProcStartLine(proc, 0)
                数(本数) = cm.ProcCountLines(proc, 0)
                モジュ(本数) = comp.Name
                名(本数) = proc
                種(本数) = comp.Type
                本文(本数) = cm.lines(始(本数), 数(本数))
                nxt = 始(本数) + 数(本数)
                本数 = 本数 + 1
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    Application.ScreenUpdating = False
    nm0 = "調査_シートを使うマクロ": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：シート「" & シート名 & "」を使っているマクロ　" & wb.Name
    out.Range("A2").Value = "シート名を文字（「""" & シート名 & """」）で書いている所と、オブジェクト名（" & コード名 & "）を使っている所を探しました。コメント行は数えません。"
    hr = 4: NC = 5
    out.Range("A4:E4").Value = Array("モジュール", "マクロ", "該当の行数", "最初の行", "その行")
    r = hr
    For i = 0 To 本数 - 1
        行 = Split(本文(i), vbCrLf)
        件 = 0: 初 = 0: 初行 = ""
        For k = 0 To UBound(行)
            ln = Trim$(CStr(行(k)))
            If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
                ヒット = (InStr(1, ln, """" & シート名 & """", 1) > 0)
                If Not ヒット And Len(コード名) > 0 Then
                    p = InStr(1, ln, コード名 & ".", 1)
                    Do While p > 0
                        前 = " "
                        If p > 1 Then 前 = Mid$(ln, p - 1, 1)
                        If InStr(境, 前) > 0 Then
                            ヒット = True
                            Exit Do
                        End If
                        p = InStr(p + 1, ln, コード名 & ".", 1)
                    Loop
                End If
                If ヒット Then
                    件 = 件 + 1
                    If 初 = 0 Then
                        初 = 始(i) + k
                        初行 = ln
                    End If
                End If
            End If
        Next k
        If 件 > 0 Then
            r = r + 1: 合計 = 合計 + 1
            out.Cells(r, 1).Value = "'" & モジュ(i)
            out.Cells(r, 2).Value = "'" & 名(i)
            out.Cells(r, 3).Value = 件
            out.Cells(r, 4).Value = 初
            out.Cells(r, 5).Value = "'" & Left$(初行, 100)
        End If
    Next i
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, NC)).Sort Key1:=out.Cells(hr + 1, 3), Order1:=2, Header:=2
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "このシートを使っているマクロは見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_シートを使うマクロ: 「" & シート名 & "」を使っているマクロ " & 合計 & " 本。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "このシートを使っているマクロ: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 言葉でマクロのコードを探す()
    ' 依頼の語: 直したい箇所|言葉を検索|どのマクロのどこ|マクロのどこ|コードから探|コードを検索|マクロの中を探|コード検索|コードの中を探
    ' 依頼の組: コード,マクロの中+検索,探+-使われ
    ' 扱う: マクロ コード 検索
    ' 見出し: なし
    ' 形: なし
    ' 探す言葉を聞いて、開いているブックのマクロのコードの中から、その言葉を含む行をモジュール・マクロ名・行番号つきで新しいシート「調査_コード検索」に一覧にする（500 行まで）。何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object, comp As Object, cm As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim n As Long, jj As Long, proc As String, nxt As Long, k As Long, i As Long, 探す As String
    Dim モジュ() As String, 名() As String, 始() As Long, 数() As Long, 本文() As String, 本数 As Long, 種() As Long
    Dim 行 As Variant, ln As String, 件 As Long, 打切 As Boolean
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    探す = Trim$(InputBox("マクロのコードの中から探す言葉を入れてください。" & vbCrLf & "（例: シート名・関数名・メッセージの一部）", "言葉でマクロのコードを探す"))
    If Len(探す) = 0 Then Exit Sub
    n = 0
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "言葉でマクロのコードを探す: VBA プロジェクトを読めません（マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    ReDim モジュ(0 To 3000): ReDim 名(0 To 3000): ReDim 始(0 To 3000): ReDim 数(0 To 3000): ReDim 本文(0 To 3000): ReDim 種(0 To 3000)
    本数 = 0
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        jj = 1
        Do While jj <= n And 本数 < 3000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                始(本数) = cm.ProcStartLine(proc, 0)
                数(本数) = cm.ProcCountLines(proc, 0)
                モジュ(本数) = comp.Name
                名(本数) = proc
                種(本数) = comp.Type
                本文(本数) = cm.lines(始(本数), 数(本数))
                nxt = 始(本数) + 数(本数)
                本数 = 本数 + 1
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    Application.ScreenUpdating = False
    nm0 = "調査_コード検索": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：言葉でマクロのコードを探す「" & 探す & "」　" & wb.Name
    hr = 4: NC = 4
    out.Range("A4:D4").Value = Array("モジュール", "マクロ", "行", "その行")
    r = hr
    For i = 0 To 本数 - 1
        行 = Split(本文(i), vbCrLf)
        For k = 0 To UBound(行)
            If InStr(1, CStr(行(k)), 探す, 1) > 0 Then
                If 件 >= 500 Then
                    打切 = True
                    Exit For
                End If
                r = r + 1: 件 = 件 + 1
                out.Cells(r, 1).Value = "'" & モジュ(i)
                out.Cells(r, 2).Value = "'" & 名(i)
                out.Cells(r, 3).Value = 始(i) + k
                out.Cells(r, 4).Value = "'" & Left$(Trim$(CStr(行(k))), 120)
            End If
        Next k
        If 打切 Then Exit For
    Next i
    out.Range("A2").Value = 件 & " 行が見つかりました" & IIf(打切, "（500 行で打ち切り）", "")
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "「" & 探す & "」を含む行は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_コード検索: 「" & 探す & "」を含む行 " & 件 & " 行。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "言葉でマクロのコードを探す: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub ブックの全シートを一覧にする()
    ' 依頼の語: 全体像|シート構成|ブックの中身|どんなシートがある|シートの一覧|ブックの構成|このブックの概要|ブックをざっと|データの中身をざっと
    ' 依頼の組: ブック+全体,構成,シート,概要,中身+教え,調べ,一覧+-重い,軽く,マクロ
    ' 扱う: 全体像 シート構成 一覧
    ' 見出し: なし
    ' 形: なし
    ' 開いているブックの全シートを、状態・使用範囲・行列数・データのセル数・数式のセル数・テーブル・グラフ・図形の数で一覧にする。
    ' 新しいシート「調査_全体像」に出す（前に作った同名の調査シートは作り直す）。元のシートは何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim ur As Range, fr As Range, lk As Variant, st As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    nm0 = "調査_全体像": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：ブックの全体像　" & wb.Name
    out.Range("A2").Value = "調べた日時 " & Format(Now, "yyyy/m/d hh:nn")
    hr = 4: NC = 12
    out.Range("A4:L4").Value = Array("シート名", "種類", "表示", "使用範囲", "行数", "列数", "データのセル", "数式のセル", "テーブル", "グラフ", "図形", "保護")
    r = hr
    For Each sh In wb.Sheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        r = r + 1
        out.Cells(r, 1).Value = "'" & sh.Name
        Select Case sh.Visible
            Case -1: st = "表示"
            Case 0: st = "非表示"
            Case Else: st = "深く隠れている"
        End Select
        out.Cells(r, 3).Value = st
        If TypeName(sh) = "Worksheet" Then
            out.Cells(r, 2).Value = "ワークシート"
            Set ur = sh.UsedRange
            out.Cells(r, 4).Value = ur.Address(False, False)
            out.Cells(r, 5).Value = ur.rows.Count
            out.Cells(r, 6).Value = ur.Columns.Count
            out.Cells(r, 7).Value = Application.WorksheetFunction.CountA(ur)
            Set fr = Nothing
            On Error Resume Next
            Set fr = ur.SpecialCells(-4123)
            On Error GoTo 失敗
            If fr Is Nothing Then
                out.Cells(r, 8).Value = 0
            Else
                out.Cells(r, 8).Value = fr.CountLarge
            End If
            out.Cells(r, 9).Value = sh.ListObjects.Count
            out.Cells(r, 10).Value = sh.ChartObjects.Count
            out.Cells(r, 11).Value = sh.Shapes.Count
            If sh.ProtectContents Then out.Cells(r, 12).Value = "保護あり"
        Else
            out.Cells(r, 2).Value = "グラフシート"
            out.Cells(r, 11).Value = sh.Shapes.Count
        End If
次のシート1:
    Next sh
    rt = r
    r = r + 2
    out.Cells(r, 1).Value = "名前定義の数"
    out.Cells(r, 2).Value = wb.names.Count
    lk = Empty
    On Error Resume Next
    lk = wb.LinkSources(1)
    On Error GoTo 失敗
    r = r + 1
    out.Cells(r, 1).Value = "外部リンクの数"
    If IsArray(lk) Then out.Cells(r, 2).Value = UBound(lk) - LBound(lk) + 1 Else out.Cells(r, 2).Value = 0
    r = r + 1
    out.Cells(r, 1).Value = "ファイルの大きさ"
    If セルの字(wb.path) <> "" Then
        out.Cells(r, 2).Value = Format(FileLen(wb.FullName), "#,##0") & " バイト（最後に保存した時点）"
    Else
        out.Cells(r, 2).Value = "（まだ保存していません）"
    End If
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_全体像: " & wb.Sheets.Count & " シートを一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "ブックの全体像: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 表の見出しと列の型を一覧にする()
    ' 依頼の語: 表の作り|列の意味|見出しと件数|表の構造|どんな表|列の型|表の中身を教|表の様子|列の内容
    ' 依頼の組: 表+作り,構成,構造,列の意味,列の型+教え,調べ+-点検
    ' 扱う: 表の作り 見出し 件数 列の型
    ' 見出し: あり
    ' 形: なし
    ' 今のシートの表（選んでいるセルを含むひとまとまり。空のセルなら使用範囲）について、見出し・データの件数・列ごとの型／空欄／種類の数／最小・最大／例を、
    ' 新しいシート「調査_表の作り」に一覧にする。見出し行は範囲の先頭行とみなす。列の「意味」までは決められないので、見出しと例を並べる。元の表は何も変えない。
    Dim wb As Workbook, ws As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim tb As Range, v As Variant, nr As Long, nc2 As Long, x As Variant, s As String, 注 As String
    Dim d As Object, n数 As Long, n日 As Long, n文 As Long, n他 As Long, n空 As Long, n誤 As Long, n文数 As Long
    Dim mn As Variant, mx As Variant, 例 As String, 例数 As Long, 型 As String, n型数 As Long, 部 As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    If TypeName(ws) <> "Worksheet" Then Exit Sub
    Set tb = ActiveCell.CurrentRegion
    If Application.WorksheetFunction.CountA(tb) = 0 Then Set tb = ws.UsedRange
    nr = tb.rows.Count: nc2 = tb.Columns.Count
    If nr > 100000 Then
        nr = 100000
        Set tb = tb.Resize(nr, nc2)
        注 = "（先頭 100,000 行まで）"
    End If
    If nr * nc2 = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = tb.Value
    Else
        v = tb.Value
    End If
    Application.ScreenUpdating = False
    nm0 = "調査_表の作り": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：表の作り　" & ws.Name
    out.Range("A2").Value = "範囲 " & tb.Address(False, False) & "　見出し行 " & tb.Row & "　データ " & Format(nr - 1, "#,##0") & " 行　" & nc2 & " 列 " & 注
    out.Range("A3").Value = "見出し行は範囲の先頭行とみなしています。列の意味は決められないので、見出しと例を並べています。"
    hr = 5: NC = 11
    out.Range("A5:K5").Value = Array("列", "見出し", "入力あり", "空欄", "型", "種類の数", "最小", "最大", "文字の数字", "エラー", "例")
    r = hr
    For j = 1 To nc2
        n数 = 0: n日 = 0: n文 = 0: n他 = 0: n空 = 0: n誤 = 0: n文数 = 0
        mn = Empty: mx = Empty: 例 = "": 例数 = 0
        Set d = CreateObject("Scripting.Dictionary")
        For i = 2 To nr
            x = v(i, j)
            If IsError(x) Then
                n誤 = n誤 + 1
            ElseIf isEmpty(x) Then
                n空 = n空 + 1
            ElseIf VarType(x) = 8 And Len(Trim$(CStr(x))) = 0 Then
                n空 = n空 + 1
            Else
                Select Case VarType(x)
                    Case 7
                        n日 = n日 + 1
                    Case 8
                        n文 = n文 + 1
                        If IsNumeric(x) Then n文数 = n文数 + 1
                    Case 11
                        n他 = n他 + 1
                    Case Else
                        n数 = n数 + 1
                End Select
                If VarType(x) <> 8 And VarType(x) <> 11 Then
                    If isEmpty(mn) Then
                        mn = x: mx = x
                    Else
                        If x < mn Then mn = x
                        If x > mx Then mx = x
                    End If
                End If
                s = CStr(x)
                If Not d.exists(s) Then
                    If d.Count < 50000 Then d.Add s, 1
                    If 例数 < 3 Then
                        If 例数 > 0 Then 例 = 例 & " / "
                        例 = 例 & Left$(s, 20)
                        例数 = 例数 + 1
                    End If
                End If
            End If
        Next i
        n型数 = (IIf(n数 > 0, 1, 0) + IIf(n日 > 0, 1, 0) + IIf(n文 > 0, 1, 0) + IIf(n他 > 0, 1, 0))
        If n型数 = 0 Then
            型 = "（データなし）"
        ElseIf n型数 = 1 Then
            If n数 > 0 Then
                型 = "数値"
            ElseIf n日 > 0 Then
                型 = "日付"
            ElseIf n文 > 0 Then
                型 = "文字"
            Else
                型 = "真偽"
            End If
        Else
            部 = ""
            If n数 > 0 Then 部 = 部 & "・数値 " & n数
            If n日 > 0 Then 部 = 部 & "・日付 " & n日
            If n文 > 0 Then 部 = 部 & "・文字 " & n文
            If n他 > 0 Then 部 = 部 & "・真偽 " & n他
            型 = "混在（" & Mid$(部, 2) & "）"
        End If
        r = r + 1
        out.Cells(r, 1).Value = Split(tb.Cells(1, j).Address(True, False), "$")(0)
        s = ""
        If Not IsError(v(1, j)) Then s = CStr(v(1, j))
        If Len(s) = 0 Then
            out.Cells(r, 2).Value = "（見出しなし）"
        Else
            out.Cells(r, 2).Value = "'" & s
        End If
        out.Cells(r, 3).Value = n数 + n日 + n文 + n他
        out.Cells(r, 4).Value = n空
        out.Cells(r, 5).Value = 型
        out.Cells(r, 6).Value = d.Count
        If Not isEmpty(mn) Then
            out.Cells(r, 7).Value = mn
            out.Cells(r, 8).Value = mx
            If VarType(mn) = 7 Then
                out.Cells(r, 7).NumberFormat = "yyyy/m/d"
                out.Cells(r, 8).NumberFormat = "yyyy/m/d"
            End If
        End If
        out.Cells(r, 9).Value = n文数
        If n文数 > 0 Then out.Cells(r, 9).Font.Color = RGB(192, 0, 0)
        out.Cells(r, 10).Value = n誤
        out.Cells(r, 11).Value = "'" & 例
    Next j
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_表の作り: " & nc2 & " 列・データ " & Format(nr - 1, "#,##0") & " 行を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "表の作り: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 入力規則などの仕掛けを一覧にする()
    ' 依頼の語: 仕掛け|入力規則|条件付き書式|ドロップダウン|プルダウン|結合しているセル|結合セルを調べ|結合セルの一覧
    ' 依頼の組: 入力規則,条件付き書式,プルダウン,ドロップダウン+調べ,どこ,一覧,探+-付け,設定し,入れて,作って,追加し
    ' 扱う: 入力規則 条件付き書式 結合セル
    ' 見出し: なし
    ' 形: なし
    ' 今のシートにある入力規則（ドロップダウンなど）・条件付き書式・結合セルを、範囲と内容つきで新しいシート「調査_仕掛け」に一覧にする。元のシートは何も変えない。
    Dim wb As Workbook, ws As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim dv As Range, ar As Range, fc As Object, c As Range, ur As Range, s As String, s2 As String, s3 As String
    Dim n1 As Long, n2 As Long, n3 As Long, n As Long, first As String, op As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    If TypeName(ws) <> "Worksheet" Then Exit Sub
    Application.ScreenUpdating = False
    nm0 = "調査_仕掛け": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：仕掛け（入力規則・条件付き書式・結合セル）　" & ws.Name
    hr = 4: NC = 5
    out.Range("A4:E4").Value = Array("種類", "範囲", "内容1", "内容2", "内容3")
    r = hr
    Set dv = Nothing
    On Error Resume Next
    Set dv = ws.Cells.SpecialCells(-4174)
    On Error GoTo 失敗
    If Not dv Is Nothing Then
        For Each ar In dv.Areas
            r = r + 1: n1 = n1 + 1
            out.Cells(r, 1).Value = "入力規則"
            out.Cells(r, 2).Value = ar.Address(False, False)
            s = "": s2 = "": s3 = "": op = ""
            On Error Resume Next
            Select Case ar.Cells(1, 1).Validation.Type
                Case 1: s = "整数"
                Case 2: s = "小数"
                Case 3: s = "リスト"
                Case 4: s = "日付"
                Case 5: s = "時刻"
                Case 6: s = "文字数"
                Case 7: s = "ユーザー設定"
                Case Else: s = "入力メッセージのみ"
            End Select
            Select Case ar.Cells(1, 1).Validation.Operator
                Case 1: op = "間"
                Case 2: op = "間以外"
                Case 3: op = "等しい"
                Case 4: op = "等しくない"
                Case 5: op = "より大きい"
                Case 6: op = "より小さい"
                Case 7: op = "以上"
                Case 8: op = "以下"
            End Select
            s2 = ar.Cells(1, 1).Validation.Formula1
            s3 = ar.Cells(1, 1).Validation.Formula2
            On Error GoTo 失敗
            If Len(s3) > 0 Then s2 = s2 & " ? " & s3
            If Len(op) > 0 Then s2 = op & " " & s2
            If Left$(s2, 1) = "=" Or Left$(s2, 1) = "+" Or Left$(s2, 1) = "-" Then s2 = "'" & s2
            out.Cells(r, 3).Value = s
            out.Cells(r, 4).Value = s2
            s3 = ""
            On Error Resume Next
            s3 = ar.Cells(1, 1).Validation.ErrorMessage
            On Error GoTo 失敗
            out.Cells(r, 5).Value = "'" & s3
        Next ar
    End If
    n = 0
    On Error Resume Next
    n = ws.Cells.FormatConditions.Count
    On Error GoTo 失敗
    If n > 2000 Then n = 2000
    For i = 1 To n
        Set fc = ws.Cells.FormatConditions(i)
        r = r + 1: n2 = n2 + 1
        out.Cells(r, 1).Value = "条件付き書式"
        s = "": s2 = ""
        On Error Resume Next
        out.Cells(r, 2).Value = fc.AppliesTo.Address(False, False)
        Select Case fc.Type
            Case 1: s = "セルの値"
            Case 2: s = "数式"
            Case 3: s = "色スケール"
            Case 4: s = "データバー"
            Case 5: s = "上位／下位"
            Case 6: s = "アイコンセット"
            Case 8: s = "重複／一意"
            Case 9: s = "文字列"
            Case 10: s = "空白"
            Case 11: s = "日付"
            Case 12: s = "平均"
            Case 13: s = "空白以外"
            Case 16: s = "エラー"
            Case 17: s = "エラー以外"
            Case Else: s = "種類 " & fc.Type
        End Select
        s2 = fc.Formula1
        s3 = "優先順位 " & fc.Priority
        On Error GoTo 失敗
        If Left$(s2, 1) = "=" Or Left$(s2, 1) = "+" Or Left$(s2, 1) = "-" Then s2 = "'" & s2
        out.Cells(r, 3).Value = s
        out.Cells(r, 4).Value = s2
        out.Cells(r, 5).Value = s3
    Next i
    ' 結合セル: 使用範囲のセルを順に見る（書式で探す Find の続き FindNext は書式の条件を引き継がず、
    ' 表題の結合セルがあると空のセルを巡り続けて固まった。2026-09-22）。20 万セル・1000 件で打ち切る
    Set ur = ws.UsedRange
    If Not IsNull(ur.MergeCells) Then
        If ur.MergeCells = False Then GoTo 結合おわり
    End If
    If ur.Cells.CountLarge > 200000 Then GoTo 結合おわり
    For Each c In ur.Cells
        If c.MergeCells Then
            If c.Address = c.MergeArea.Cells(1, 1).Address Then
                r = r + 1: n3 = n3 + 1
                out.Cells(r, 1).Value = "結合セル"
                out.Cells(r, 2).Value = c.MergeArea.Address(False, False)
                s = ""
                On Error Resume Next
                s = CStr(c.text)
                On Error GoTo 失敗
                out.Cells(r, 4).Value = "'" & Left$(s, 30)
                If n3 >= 1000 Then Exit For
            End If
        End If
    Next c
結合おわり:
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "仕掛けは見つかりませんでした"
    End If
    rt = r
    out.Range("A2").Value = "入力規則 " & n1 & " 件・条件付き書式 " & n2 & " 件・結合セル " & n3 & " 件"
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_仕掛け: 入力規則 " & n1 & "・条件付き書式 " & n2 & "・結合セル " & n3 & " 件を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.FindFormat.Clear
    Application.StatusBar = "仕掛け: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 印刷ページ数と切れ目を一覧にする()
    ' 依頼の語: 何ページ|ページが切れ|ページ数|どこで切れ|改ページ|ページに収ま|印刷すると|印刷したら|ページになる
    ' 依頼の組: 印刷,ページ+何枚,何ページ,切れ,調べ+-設定
    ' 扱う: ページ数 改ページ 印刷
    ' 見出し: なし
    ' 形: なし
    ' 開いているブックの表示中のシートごとに、印刷範囲・向き・用紙・拡大縮小・総ページ数・行と列の切れ目・繰り返す見出し行を、新しいシート「調査_印刷ページ」に一覧にする。
    ' プリンターが未設定だとページ数が出ないことがある。元のシートは何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim pg As Long, hb As String, vb2 As String, s As String, zm As Variant, col As Long
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    nm0 = "調査_印刷ページ": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：印刷ページ　" & wb.Name
    out.Range("A2").Value = "切れ目の「N行目」は、その行から次のページが始まる位置です。"
    hr = 4: NC = 9
    out.Range("A4:I4").Value = Array("シート", "印刷範囲", "向き", "用紙", "拡大縮小", "総ページ数", "行の切れ目", "列の切れ目", "繰り返す見出し行")
    r = hr
    For Each sh In wb.Worksheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        If sh.Visible = -1 Then
            r = r + 1
            out.Cells(r, 1).Value = "'" & sh.Name
            pg = -1: hb = "": vb2 = ""
            On Error Resume Next
            With sh.PageSetup
                s = .PrintArea
                If Len(s) = 0 Then s = "（シート全体）"
                out.Cells(r, 2).Value = "'" & s
                out.Cells(r, 3).Value = IIf(.Orientation = 2, "横", "縦")
                Select Case .PaperSize
                    Case 8: out.Cells(r, 4).Value = "A3"
                    Case 9: out.Cells(r, 4).Value = "A4"
                    Case 11: out.Cells(r, 4).Value = "A5"
                    Case 12: out.Cells(r, 4).Value = "B4"
                    Case 13: out.Cells(r, 4).Value = "B5"
                    Case Else: out.Cells(r, 4).Value = "用紙 " & .PaperSize
                End Select
                zm = .Zoom
                If zm = False Then
                    out.Cells(r, 5).Value = "横 " & .FitToPagesWide & " ×縦 " & .FitToPagesTall & " ページに収める"
                Else
                    out.Cells(r, 5).Value = zm & " %"
                End If
                out.Cells(r, 9).Value = "'" & .PrintTitleRows
                pg = .Pages.Count
            End With
            For i = 1 To sh.HPageBreaks.Count
                If i > 30 Then hb = hb & " …": Exit For
                If Len(hb) > 0 Then hb = hb & ", "
                hb = hb & sh.HPageBreaks(i).Location.Row & "行目"
            Next i
            For i = 1 To sh.VPageBreaks.Count
                If i > 30 Then vb2 = vb2 & " …": Exit For
                col = sh.VPageBreaks(i).Location.Column
                If Len(vb2) > 0 Then vb2 = vb2 & ", "
                vb2 = vb2 & Split(sh.Cells(1, col).Address(True, False), "$")(0) & "列"
            Next i
            On Error GoTo 失敗
            If pg >= 0 Then
                out.Cells(r, 6).Value = pg
            Else
                out.Cells(r, 6).Value = "取得できません（プリンター未設定？）"
            End If
            out.Cells(r, 7).Value = "'" & hb
            out.Cells(r, 8).Value = "'" & vb2
        End If
次のシート1:
    Next sh
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_印刷ページ: " & (rt - hr) & " シートのページ数を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "印刷ページ数: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 非表示のシートと行列を一覧にする()
    ' 依頼の語: 非表示の|非表示が|非表示を調べ|隠れている|隠れた行|隠している行|見えない行|見えない列|隠しシート|隠れたシート
    ' 依頼の組: 非表示,隠れ,隠し,見えない+-全部表示,表示して,表示に戻,元に戻,解除,再表示,見せて
    ' 扱う: 非表示 行 列 シート 名前
    ' 見出し: なし
    ' 形: なし
    ' ブックの中の非表示のシート・行・列・名前と、絞り込みで隠れている行を、新しいシート「調査_非表示」に一覧にする。何も表示し直さず、元のシートは変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim ur As Range, x As Variant, a As Long, lastR As Long, lastC As Long, nn As Object, 何 As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    nm0 = "調査_非表示": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：非表示のもの　" & wb.Name
    hr = 4: NC = 4
    out.Range("A4:D4").Value = Array("種類", "シート", "場所", "備考")
    r = hr
    For Each sh In wb.Sheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        If sh.Visible <> -1 Then
            r = r + 1
            If sh.Visible = 0 Then 何 = "非表示のシート" Else 何 = "深く隠れたシート"
            out.Cells(r, 1).Value = 何
            out.Cells(r, 2).Value = "'" & sh.Name
            If sh.Visible = 0 Then out.Cells(r, 4).Value = "シート見出しの右クリック［再表示］で出せる" Else out.Cells(r, 4).Value = "VBE のプロパティでしか出せない"
        End If
        If TypeName(sh) = "Worksheet" Then
            Set ur = sh.UsedRange
            If sh.FilterMode Then
                r = r + 1
                out.Cells(r, 1).Value = "絞り込み中"
                out.Cells(r, 2).Value = "'" & sh.Name
                out.Cells(r, 4).Value = "フィルターで行が隠れている"
            End If
            x = Null
            On Error Resume Next
            x = Null      ' 混在でも False が返るので、必ず 1 行ずつ調べる
            On Error GoTo 失敗
            If IsNull(x) Then
                lastR = ur.Row + ur.rows.Count - 1
                If lastR > 50000 Then lastR = 50000
                i = ur.Row
                Do While i <= lastR
                    If sh.rows(i).Hidden Then
                        a = i
                        Do While i < lastR
                            If Not sh.rows(i + 1).Hidden Then Exit Do
                            i = i + 1
                        Loop
                        r = r + 1
                        out.Cells(r, 1).Value = "非表示の行"
                        out.Cells(r, 2).Value = "'" & sh.Name
                        If i > a Then out.Cells(r, 3).Value = a & "?" & i & " 行目" Else out.Cells(r, 3).Value = a & " 行目"
                    End If
                    i = i + 1
                Loop
            ElseIf x = True Then
                r = r + 1
                out.Cells(r, 1).Value = "非表示の行"
                out.Cells(r, 2).Value = "'" & sh.Name
                out.Cells(r, 3).Value = "使用範囲の行がすべて非表示"
            End If
            x = Null
            On Error Resume Next
            x = Null      ' 混在でも False が返るので、必ず 1 行ずつ調べる
            On Error GoTo 失敗
            If IsNull(x) Then
                lastC = ur.Column + ur.Columns.Count - 1
                If lastC > 2000 Then lastC = 2000
                i = ur.Column
                Do While i <= lastC
                    If sh.Columns(i).Hidden Then
                        a = i
                        Do While i < lastC
                            If Not sh.Columns(i + 1).Hidden Then Exit Do
                            i = i + 1
                        Loop
                        r = r + 1
                        out.Cells(r, 1).Value = "非表示の列"
                        out.Cells(r, 2).Value = "'" & sh.Name
                        If i > a Then
                            out.Cells(r, 3).Value = Split(sh.Cells(1, a).Address(True, False), "$")(0) & "?" & Split(sh.Cells(1, i).Address(True, False), "$")(0) & " 列"
                        Else
                            out.Cells(r, 3).Value = Split(sh.Cells(1, a).Address(True, False), "$")(0) & " 列"
                        End If
                    End If
                    i = i + 1
                Loop
            ElseIf x = True Then
                r = r + 1
                out.Cells(r, 1).Value = "非表示の列"
                out.Cells(r, 2).Value = "'" & sh.Name
                out.Cells(r, 3).Value = "使用範囲の列がすべて非表示"
            End If
        End If
次のシート1:
    Next sh
    For Each nn In wb.names
        If Not nn.Visible Then
            If InStr(nn.Name, "_xlnm") = 0 And InStr(nn.Name, "_FilterDatabase") = 0 And InStr(nn.Name, "_xlfn") = 0 Then
                r = r + 1
                out.Cells(r, 1).Value = "非表示の名前"
                out.Cells(r, 3).Value = "'" & nn.Name
                out.Cells(r, 4).Value = "'" & nn.RefersTo
            End If
        End If
    Next nn
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "非表示のものは見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_非表示: " & (rt - hr) & " 件を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "非表示: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 壊れた参照と名前を一覧にする()
    ' 依頼の語: 壊れている参照|壊れた参照|参照が壊れ|参照エラー|#REF|リンク切れ|参照先が消え|参照や名前が残|名前が壊れ
    ' 依頼の組: 参照+壊れ,REF,切れ,消え
    ' 扱う: 壊れた参照 REF 名前
    ' 見出し: なし
    ' 形: なし
    ' ブックの中の #REF!・#NAME? のセルと、式の中に #REF! を持つセル、参照先が壊れた名前を、新しいシート「調査_壊れた参照」に一覧にする。何も直さない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim fr As Range, c As Range, f As Range, s As String, n As Long, first As String, seen As Object, nn As Object, 他 As Long, 何 As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set seen = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    nm0 = "調査_壊れた参照": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：壊れた参照　" & wb.Name
    hr = 4: NC = 5
    out.Range("A4:E4").Value = Array("種類", "シート", "セル／名前", "内容（式）", "備考")
    r = hr
    For Each sh In wb.Worksheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        Set fr = Nothing
        On Error Resume Next
        Set fr = sh.UsedRange.SpecialCells(-4123, 16)
        On Error GoTo 失敗
        If Not fr Is Nothing Then
            n = 0
            For Each c In fr.Cells
                n = n + 1
                If n > 2000 Then Exit For
                s = CStr(c.Value)
                If s = "Error 2023" Or s = "Error 2029" Then
                    r = r + 1
                    If s = "Error 2023" Then 何 = "#REF!（参照が壊れている）" Else 何 = "#NAME?（名前が不明）"
                    out.Cells(r, 1).Value = 何
                    out.Cells(r, 2).Value = "'" & sh.Name
                    out.Cells(r, 3).Value = c.Address(False, False)
                    out.Cells(r, 4).Value = "'" & c.Formula
                    seen(sh.Name & "!" & c.Address) = 1
                Else
                    他 = 他 + 1
                End If
            Next c
        End If
        Set f = Nothing
        On Error Resume Next
        Set f = sh.Cells.Find("#REF!", LookIn:=-4123, LookAt:=2, SearchOrder:=1, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
        On Error GoTo 失敗
        If Not f Is Nothing Then
            first = f.Address
            n = 0
            Do
                n = n + 1
                If Not seen.exists(sh.Name & "!" & f.Address) Then
                    r = r + 1
                    out.Cells(r, 1).Value = "式や文字に #REF! を含む"
                    out.Cells(r, 2).Value = "'" & sh.Name
                    out.Cells(r, 3).Value = f.Address(False, False)
                    out.Cells(r, 4).Value = "'" & f.Formula
                    out.Cells(r, 5).Value = "今は別の値に見えているが、式の中に壊れた参照がある"
                End If
                Set f = sh.Cells.FindNext(f)
                If f Is Nothing Then Exit Do
            Loop While f.Address <> first And n < 2000
        End If
次のシート1:
    Next sh
    Dim 外s As String, 外p1 As Long, 外p2 As Long, 外t As String, 外前 As String
    For Each nn In wb.names
        外s = nn.RefersTo
        If InStr(外s, "#REF!") > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "壊れた名前"
            out.Cells(r, 3).Value = "'" & nn.Name
            out.Cells(r, 4).Value = "'" & 外s
        Else
            ' 他のブックを指す名前（別のブックへシートやセルをコピーすると、使われている名前だけが持ち込まれる）。
            ' 2026 年の仕様変更で、参照先が消えていても名前の管理でエラーにならないので、#REF! の検査では見つからない
            外p1 = InStr(外s, "[")
            If 外p1 > 0 Then
                外p2 = InStr(外p1, 外s, "]")
                If 外p2 > 外p1 Then
                    外t = LCase$(Mid$(外s, 外p1 + 1, 外p2 - 外p1 - 1))
                    外前 = "="
                    If 外p1 > 1 Then 外前 = Mid$(外s, 外p1 - 1, 1)
                    If 外t Like "*.xl*" Or (IsNumeric(外t) And InStr("=(,+-*/&' ", 外前) > 0) Then
                        r = r + 1
                        out.Cells(r, 1).Value = "他のブックを指す名前"
                        out.Cells(r, 3).Value = "'" & nn.Name
                        out.Cells(r, 4).Value = "'" & 外s
                        out.Cells(r, 5).Value = "別のブックから持ち込まれた名前。開くたびに更新確認が出る。参照先が消えていてもエラーにならない" & _
                            IIf(InStr(LCase$(外s), "\users\") > 0 Or InStr(LCase$(外s), "onedrive") > 0 Or InStr(LCase$(外s), "sharepoint") > 0 Or InStr(LCase$(外s), "http") > 0, _
                                "／参照先のフルパスに個人名・OneDrive・SharePoint が含まれる", "")
                    End If
                End If
            End If
        End If
    Next nn
    If 他 > 0 Then out.Range("A2").Value = "このほかに、#REF!・#NAME? 以外のエラー値のセルが " & 他 & " 件あります（一覧は「エラー値のセルを一覧にする」）。"
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "壊れた参照は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_壊れた参照: " & (rt - hr) & " 件を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "壊れた参照: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 参照とリンクを一覧にする()
    ' 依頼の語: シート間の参照|参照とリンク|リンクを調べ|外部リンクを調べ|外部リンクの一覧|外部ファイルへのリンク|参照関係|どのシートを参照|参照しているシート
    ' 依頼の組: 参照,リンク+調べ,一覧+-壊れ,REF,切,解消
    ' 扱う: 参照 リンク シート間 外部
    ' 見出し: なし
    ' 形: なし
    ' 数式が他のシートや他のブックを参照している関係と、ブックのリンク元ファイルを、新しいシート「調査_参照とリンク」に一覧にする。何も変えない（リンクの解消は「外部リンクを値に変えて切る」）。
    Dim wb As Workbook, sh As Object, t As Object, out As Worksheet, old As Object
    Dim kv As Variant
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim fr As Range, ar As Range, v As Variant, a As Long, b As Long, f As String, key As String, total As Long
    Dim cnt As Object, ex As Object, cnt2 As Object, ex2 As Object, lk As Variant, p1 As Long, p2 As Long, bk As String, 打切 As Boolean
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set cnt = CreateObject("Scripting.Dictionary"): Set ex = CreateObject("Scripting.Dictionary")
    Set cnt2 = CreateObject("Scripting.Dictionary"): Set ex2 = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    nm0 = "調査_参照とリンク": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：参照とリンク　" & wb.Name
    hr = 4: NC = 5
    out.Range("A4:E4").Value = Array("種類", "参照元", "参照先", "件数", "最初の例")
    r = hr
    For Each sh In wb.Worksheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        Set fr = Nothing
        On Error Resume Next
        Set fr = sh.UsedRange.SpecialCells(-4123)
        On Error GoTo 失敗
        If Not fr Is Nothing Then
            For Each ar In fr.Areas
                If ar.CountLarge = 1 Then
                    ReDim v(1 To 1, 1 To 1)
                    v(1, 1) = ar.Formula
                Else
                    v = ar.Formula
                End If
                For a = 1 To UBound(v, 1)
                    For b = 1 To UBound(v, 2)
                        total = total + 1
                        If total > 200000 Then
                            打切 = True
                            GoTo 集計おわり
                        End If
                        f = CStr(v(a, b))
                        If InStr(f, "!") > 0 Then
                            For Each t In wb.Worksheets
                                If t.Name <> sh.Name Then
                                    If InStr(f, "'" & t.Name & "'!") > 0 Or InStr(f, t.Name & "!") > 0 Then
                                        key = sh.Name & "|" & t.Name
                                        If cnt.exists(key) Then
                                            cnt(key) = cnt(key) + 1
                                        Else
                                            cnt.Add key, 1
                                            ex.Add key, ar.Cells(a, b).Address(False, False) & "  " & f
                                        End If
                                    End If
                                End If
                            Next t
                            p1 = InStr(f, "[")
                            If p1 > 0 Then
                                p2 = InStr(p1, f, "]")
                                If p2 > p1 Then
                                    bk = Mid$(f, p1 + 1, p2 - p1 - 1)
                                    key = sh.Name & "|" & bk
                                    If cnt2.exists(key) Then
                                        cnt2(key) = cnt2(key) + 1
                                    Else
                                        cnt2.Add key, 1
                                        ex2.Add key, ar.Cells(a, b).Address(False, False) & "  " & f
                                    End If
                                End If
                            End If
                        End If
                    Next b
                Next a
            Next ar
        End If
次のシート1:
    Next sh
集計おわり:
    For Each kv In cnt.keys
        key = CStr(kv)
        r = r + 1
        out.Cells(r, 1).Value = "シート間の参照"
        out.Cells(r, 2).Value = "'" & Split(key, "|")(0)
        out.Cells(r, 3).Value = "'" & Split(key, "|")(1)
        out.Cells(r, 4).Value = cnt(key)
        out.Cells(r, 5).Value = "'" & ex(key)
    Next kv
    For Each kv In cnt2.keys
        key = CStr(kv)
        r = r + 1
        out.Cells(r, 1).Value = "他のブックの参照（式の中）"
        out.Cells(r, 2).Value = "'" & Split(key, "|")(0)
        out.Cells(r, 3).Value = "'" & Split(key, "|")(1)
        out.Cells(r, 4).Value = cnt2(key)
        out.Cells(r, 5).Value = "'" & ex2(key)
    Next kv
    lk = Empty
    On Error Resume Next
    lk = wb.LinkSources(1)
    On Error GoTo 失敗
    If IsArray(lk) Then
        For i = LBound(lk) To UBound(lk)
            r = r + 1
            out.Cells(r, 1).Value = "リンク元ファイル"
            out.Cells(r, 3).Value = "'" & lk(i)
            ' フルパスに Windows のログイン名・OneDrive・SharePoint のサーバー名が残る（ブックを人に渡すと見える）
            If InStr(LCase$(lk(i)), "\users\") > 0 Or InStr(LCase$(lk(i)), "onedrive") > 0 Or InStr(LCase$(lk(i)), "sharepoint") > 0 Or InStr(LCase$(lk(i)), "http") > 0 Then
                out.Cells(r, 5).Value = "フルパスに個人名・OneDrive・SharePoint が含まれる"
            End If
        Next i
    End If
    ' 名前が他のブックを指している（別のブックへのコピーで持ち込まれた名前。式の中には外部参照が見えないので上の一覧に出ない）
    Dim 外nn As Object, 外s As String, 外p1 As Long, 外p2 As Long, 外t As String, 外前 As String
    For Each 外nn In wb.names
        外s = ""
        On Error Resume Next
        外s = 外nn.RefersTo
        On Error GoTo 失敗
        外p1 = InStr(外s, "[")
        If 外p1 > 0 Then
            外p2 = InStr(外p1, 外s, "]")
            If 外p2 > 外p1 Then
                外t = LCase$(Mid$(外s, 外p1 + 1, 外p2 - 外p1 - 1))
                外前 = "="
                If 外p1 > 1 Then 外前 = Mid$(外s, 外p1 - 1, 1)
                If 外t Like "*.xl*" Or (IsNumeric(外t) And InStr("=(,+-*/&' ", 外前) > 0) Then
                    r = r + 1
                    out.Cells(r, 1).Value = "他のブックを指す名前"
                    out.Cells(r, 2).Value = "'" & 外nn.Name
                    out.Cells(r, 3).Value = "'" & 外s
                    out.Cells(r, 4).Value = 1
                    If InStr(LCase$(外s), "\users\") > 0 Or InStr(LCase$(外s), "onedrive") > 0 Or InStr(LCase$(外s), "sharepoint") > 0 Or InStr(LCase$(外s), "http") > 0 Then
                        out.Cells(r, 5).Value = "フルパスに個人名・OneDrive・SharePoint が含まれる"
                    End If
                End If
            End If
        End If
    Next 外nn
    If 打切 Then out.Range("A2").Value = "数式が多いので、先頭の 200,000 個までを調べました。"
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "他のシート・他のブックへの参照は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_参照とリンク: " & (rt - hr) & " 件を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "参照とリンク: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub ブックが重い原因を一覧にする()
    ' 依頼の語: 重い原因|原因を探|重くなった|なぜ重い|ブックが重|遅い原因|重い理由|動作が重
    ' 依頼の組: 重い,遅い+原因,理由,なぜ
    ' 扱う: 重い 原因 サイズ 使用範囲 図形
    ' 見出し: なし
    ' 形: なし
    ' ブックが重くなる代表的な原因（余計に広がった使用範囲・図形と画像・条件付き書式・スタイル・名前・外部リンク・再計算の多い関数など）の状況を、
    ' 新しいシート「調査_重い原因」に一覧にして、気になるものに印を付ける。何も変えない（直すのは「余分な行列を削除し軽くする」）。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim ur As Range, lc As Range, fr As Range, endR As Long, endC As Long, lastR As Long, lastC As Long, gR As Long, gC As Long
    Dim shp As Object, np As Long, ns As Long, fcN As Long, lk As Variant, f As Range, first As String, n As Long, w As Variant, 数式計 As Long, s As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    nm0 = "調査_重い原因": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：ブックが重い原因　" & wb.Name
    out.Range("A2").Value = "「要確認」は目安です。直すなら「余分な行列を削除し軽くする」（使用範囲の余計な書式を落とす）から。"
    hr = 4: NC = 4
    out.Range("A4:D4").Value = Array("項目", "状況", "判定", "説明")
    r = hr
    r = r + 1
    out.Cells(r, 1).Value = "ファイルの大きさ"
    If セルの字(wb.path) <> "" Then
        out.Cells(r, 2).Value = Format(FileLen(wb.FullName), "#,##0") & " バイト"
        If FileLen(wb.FullName) > 20000000 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "20 MB を超えています"
    Else
        out.Cells(r, 2).Value = "（まだ保存していません）"
    End If
    For Each sh In wb.Worksheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        Set ur = sh.UsedRange
        endR = ur.Row + ur.rows.Count - 1
        endC = ur.Column + ur.Columns.Count - 1
        lastR = 0: lastC = 0
        Set lc = Nothing
        On Error Resume Next
        Set lc = sh.Cells.Find("*", SearchOrder:=1, SearchDirection:=2, LookIn:=-4123, LookAt:=2, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
        If Not lc Is Nothing Then lastR = lc.Row
        Set lc = Nothing
        Set lc = sh.Cells.Find("*", SearchOrder:=2, SearchDirection:=2, LookIn:=-4123, LookAt:=2, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
        If Not lc Is Nothing Then lastC = lc.Column
        On Error GoTo 失敗
        gR = endR - lastR: gC = endC - lastC
        r = r + 1
        out.Cells(r, 1).Value = "使用範囲：" & sh.Name
        out.Cells(r, 2).Value = "使用範囲 " & ur.Address(False, False) & "／データの最後 " & lastR & " 行・" & lastC & " 列"
        If gR > 1000 Or (gR > 100 And gC > 50) Then
            out.Cells(r, 3).Value = "要確認"
            out.Cells(r, 4).Value = "使用範囲がデータより " & gR & " 行・" & gC & " 列ぶん広い（消したはずの書式が残っている可能性）"
        End If
        np = 0
        Dim 幽 As Long, 幅 As Double, 高 As Double
        幽 = 0
        For Each shp In sh.Shapes
            If shp.Type = 13 Then np = np + 1
            ' 幅か高さが 0 の図形（幽霊）。線・コネクタ・コメントは幅か高さが 0 でも正常なので数えない
            幅 = 1: 高 = 1
            On Error Resume Next
            幅 = shp.Width: 高 = shp.Height
            If 幅 < 1 Or 高 < 1 Then
                If shp.Type <> 9 And shp.Type <> 4 And shp.Connector = False Then 幽 = 幽 + 1
            End If
            On Error GoTo 失敗
        Next shp
        ns = sh.Shapes.Count
        If ns > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "図形・画像：" & sh.Name
            out.Cells(r, 2).Value = "図形 " & ns & " 個（うち画像 " & np & " 個）"
            If ns > 200 Or np > 30 Then
                out.Cells(r, 3).Value = "要確認"
                out.Cells(r, 4).Value = "図形や画像が多い（同じ図形が重なっていないか）"
            End If
        End If
        If 幽 > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "幽霊の図形：" & sh.Name
            out.Cells(r, 2).Value = "幅か高さが 0 の図形 " & 幽 & " 個"
            If 幽 >= 10 Then out.Cells(r, 3).Value = "要確認"
            out.Cells(r, 4).Value = "見えないのに残っている図形（コピー元の貼り付け残り）。ブックを重くする。「シートの図形と画像を全部削除する」「図形を一覧にし位置をそろえる」で確かめる"
        End If
        fcN = 0
        On Error Resume Next
        fcN = sh.Cells.FormatConditions.Count
        On Error GoTo 失敗
        If fcN > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "条件付き書式：" & sh.Name
            out.Cells(r, 2).Value = fcN & " 件"
            If fcN > 200 Then
                out.Cells(r, 3).Value = "要確認"
                out.Cells(r, 4).Value = "条件付き書式が多い（コピーのたびに増えることがある）"
            End If
        End If
        Set fr = Nothing
        On Error Resume Next
        Set fr = ur.SpecialCells(-4123)
        On Error GoTo 失敗
        If Not fr Is Nothing Then 数式計 = 数式計 + fr.CountLarge
次のシート1:
    Next sh
    r = r + 1
    out.Cells(r, 1).Value = "数式のセル（全シート）"
    out.Cells(r, 2).Value = Format(数式計, "#,##0") & " 個"
    If 数式計 > 200000 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "数式が多い（再計算に時間がかかる）"
    r = r + 1
    out.Cells(r, 1).Value = "スタイルの数"
    out.Cells(r, 2).Value = wb.Styles.Count & " 個"
    If wb.Styles.Count > 500 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "他のブックから貼り付けるたびに増える。500 を超えると重くなりやすい"
    ' ユーザー定義のスタイル（組み込みでないもの。他のブックから貼るたびに増え、表示形式もスタイルごとに持つ）
    Dim 型, 自 As Long
    自 = 0
    On Error Resume Next
    For Each 型 In wb.Styles
        If 型.BuiltIn = False Then 自 = 自 + 1
    Next 型
    On Error GoTo 失敗
    r = r + 1
    out.Cells(r, 1).Value = "ユーザー定義のスタイル"
    out.Cells(r, 2).Value = 自 & " 個"
    If 自 > 200 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "組み込みでないスタイルが多い（他のブックから貼るたびに増える。表示形式もここに溜まる）。直すのは「余分な行列を削除し軽くする」ではなく、不要なスタイルを消す"
    r = r + 1
    out.Cells(r, 1).Value = "名前定義の数"
    out.Cells(r, 2).Value = wb.names.Count & " 個"
    If wb.names.Count > 1000 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "名前が多い（使っていない名前が溜まっていないか）"
    ' 別のブックから持ち込まれた名前（外部リンクの元。参照先が消えていても名前の管理にエラーが出ない）
    Dim nn As Object, 外s As String, 外p1 As Long, 外p2 As Long, 外t As String, 外前 As String, n外 As Long
    For Each nn In wb.names
        外s = ""
        On Error Resume Next
        外s = nn.RefersTo
        On Error GoTo 失敗
        外p1 = InStr(外s, "[")
        If 外p1 > 0 Then
            外p2 = InStr(外p1, 外s, "]")
            If 外p2 > 外p1 Then
                外t = LCase$(Mid$(外s, 外p1 + 1, 外p2 - 外p1 - 1))
                外前 = "="
                If 外p1 > 1 Then 外前 = Mid$(外s, 外p1 - 1, 1)
                If 外t Like "*.xl*" Or (IsNumeric(外t) And InStr("=(,+-*/&' ", 外前) > 0) Then n外 = n外 + 1
            End If
        End If
    Next nn
    If n外 > 0 Then
        r = r + 1
        out.Cells(r, 1).Value = "他のブックを指す名前"
        out.Cells(r, 2).Value = n外 & " 本"
        out.Cells(r, 3).Value = "要確認"
        out.Cells(r, 4).Value = "別のブックから持ち込まれた名前（外部リンクの元・開くたびに更新確認が出る）。一覧は「参照とリンクを一覧にする」"
    End If
    lk = Empty
    On Error Resume Next
    lk = wb.LinkSources(1)
    On Error GoTo 失敗
    r = r + 1
    out.Cells(r, 1).Value = "外部リンク"
    If IsArray(lk) Then
        out.Cells(r, 2).Value = (UBound(lk) - LBound(lk) + 1) & " 件"
        out.Cells(r, 3).Value = "要確認"
        out.Cells(r, 4).Value = "開くたびに更新確認が出て遅くなることがある"
    Else
        out.Cells(r, 2).Value = "なし"
    End If
    n = 0
    On Error Resume Next
    n = wb.PivotCaches.Count
    On Error GoTo 失敗
    r = r + 1
    out.Cells(r, 1).Value = "ピボットのキャッシュ"
    out.Cells(r, 2).Value = n & " 個"
    If n > 10 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "ピボットごとにデータの写しを持つので、多いと大きくなる"
    For Each w In Array("INDIRECT(", "OFFSET(", "TODAY(", "NOW(", "RAND(", "RANDBETWEEN(")
        n = 0
        For Each sh In wb.Worksheets
            If TypeName(sh) = "Worksheet" Then
                If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート2
            End If
            Set f = Nothing
            On Error Resume Next
            Set f = sh.Cells.Find(w, LookIn:=-4123, LookAt:=2, SearchOrder:=1, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
            On Error GoTo 失敗
            If Not f Is Nothing Then
                first = f.Address
                Do
                    n = n + 1
                    Set f = sh.Cells.FindNext(f)
                    If f Is Nothing Then Exit Do
                Loop While f.Address <> first And n < 5000
            End If
            If n >= 5000 Then Exit For
次のシート2:
        Next sh
        If n > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "再計算の多い関数：" & Left$(w, Len(w) - 1)
            out.Cells(r, 2).Value = IIf(n >= 5000, "5,000 個以上", n & " 個")
            If n > 500 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "何かを変えるたびに再計算される関数が多い"
        End If
    Next w
    r = r + 1
    out.Cells(r, 1).Value = "計算方法"
    Select Case Application.Calculation
        Case -4105: s = "自動"
        Case -4135: s = "手動"
        Case Else: s = "テーブル以外は自動"
    End Select
    out.Cells(r, 2).Value = s
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    For i = hr + 1 To rt
        If out.Cells(i, 3).Value = "要確認" Then out.Range(out.Cells(i, 1), out.Cells(i, NC)).Interior.Color = RGB(255, 242, 204)
    Next i
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_重い原因: " & (rt - hr) & " 項目を一覧にしました（要確認は色付き）。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "重い原因: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 同じ形の表のシートを一覧にする()
    ' 依頼の語: 同じ形の表|似た表|似た形の表|同じ見出しの表|他のシートにないか|同じ構造の表|同じ形の他の表
    ' 依頼の組: 同じ形,似た,同じ構造,同じ見出し+表+-クエリ,集計
    ' 扱う: 同じ形 見出し 比較 シート
    ' 見出し: あり
    ' 形: なし
    ' ブックの各シートの見出し行（先頭 30 行のうち最初に 3 つ以上の文字が並ぶ行）を比べ、見出しが 6 割以上そろっているシートの組を、新しいシート「調査_同じ形の表」に一覧にする。
    ' 集計で足し合わせられる表を見つけるのに使う。何も変えない。
    Dim wb As Workbook, sh As Object, out As Worksheet, old As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim ur As Range, hrow As Long, v As Variant, s As String, n As Long, a As Long, b As Long, common As Long, sim As Double
    Dim nms() As String, dic() As Object, rowsN() As Long, k As Variant, 差A As String, 差B As String, cA As Long, cB As Long
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    For Each sh In wb.Worksheets
        If TypeName(sh) = "Worksheet" Then
            If Left$(セルの字(sh.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート1
        End If
        If sh.Visible = -1 Then
            Set ur = sh.UsedRange
            If Application.WorksheetFunction.CountA(ur) > 0 Then
                hrow = 0
                For i = 1 To IIf(ur.rows.Count > 30, 30, ur.rows.Count)
                    If Application.WorksheetFunction.CountA(ur.rows(i)) >= 3 Then
                        hrow = i
                        Exit For
                    End If
                Next i
                If hrow > 0 Then
                    n = n + 1
                    ReDim Preserve nms(1 To n)
                    ReDim Preserve dic(1 To n)
                    ReDim Preserve rowsN(1 To n)
                    nms(n) = sh.Name
                    Set dic(n) = CreateObject("Scripting.Dictionary")
                    v = ur.rows(hrow).Value
                    For j = 1 To UBound(v, 2)
                        s = ""
                        If Not IsError(v(1, j)) Then s = Trim$(CStr(v(1, j)))
                        If Len(s) > 0 Then
                            If Not dic(n).exists(s) Then dic(n).Add s, 1
                        End If
                    Next j
                    rowsN(n) = ur.rows.Count - hrow
                End If
            End If
        End If
次のシート1:
    Next sh
    Application.ScreenUpdating = False
    nm0 = "調査_同じ形の表": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：同じ形の表　" & wb.Name
    out.Range("A2").Value = "各シートの見出し行を比べ、見出しが 6 割以上そろっている組を出しています（類似度が高い順）。"
    hr = 4: NC = 8
    out.Range("A4:H4").Value = Array("シートA", "シートB", "そろっている見出し", "類似度", "Aだけの見出し", "Bだけの見出し", "Aのデータ行", "Bのデータ行")
    r = hr
    For a = 1 To n - 1
        For b = a + 1 To n
            common = 0
            For Each k In dic(a).keys
                If dic(b).exists(k) Then common = common + 1
            Next k
            If dic(a).Count + dic(b).Count - common > 0 Then sim = common / (dic(a).Count + dic(b).Count - common) Else sim = 0
            If common >= 3 And sim >= 0.6 Then
                差A = "": 差B = "": cA = 0: cB = 0
                For Each k In dic(a).keys
                    If Not dic(b).exists(k) Then
                        cA = cA + 1
                        If cA <= 6 Then 差A = 差A & IIf(Len(差A) > 0, "、", "") & k
                    End If
                Next k
                For Each k In dic(b).keys
                    If Not dic(a).exists(k) Then
                        cB = cB + 1
                        If cB <= 6 Then 差B = 差B & IIf(Len(差B) > 0, "、", "") & k
                    End If
                Next k
                r = r + 1
                out.Cells(r, 1).Value = "'" & nms(a)
                out.Cells(r, 2).Value = "'" & nms(b)
                out.Cells(r, 3).Value = common
                out.Cells(r, 4).Value = sim
                out.Cells(r, 4).NumberFormat = "0%"
                out.Cells(r, 5).Value = "'" & 差A
                out.Cells(r, 6).Value = "'" & 差B
                out.Cells(r, 7).Value = rowsN(a)
                out.Cells(r, 8).Value = rowsN(b)
            End If
        Next b
    Next a
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, NC)).Sort Key1:=out.Cells(hr + 1, 4), Order1:=2, Header:=2
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "見出しが似ているシートの組は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_同じ形の表: " & n & " シートを比べました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "同じ形の表: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub マクロを役割つきで一覧にする()
    ' 依頼の語: マクロ一覧|マクロの一覧|マクロを一覧|役割つき|どんなマクロがある|マクロの役割|マクロの数
    ' 依頼の組: マクロ+一覧,どんな+-使われ,呼び
    ' 扱う: マクロ 一覧 役割
    ' 見出し: なし
    ' 形: なし
    ' 開いているブックの VBA のマクロ（Sub・Function）を、モジュール名・種類・公開／非公開・行数・役割（頭のコメントの 1 行）つきで、新しいシート「調査_マクロ一覧」に一覧にする。
    ' VBA プロジェクトを読めない設定のときは、その旨を出して止まる。何も変えない。
    Dim wb As Workbook, out As Worksheet, old As Object, comp As Object, cm As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim proc As String, nxt As Long, jj As Long, n As Long, bl As Long, decl As String, role As String, t As String, k As Long, 型名 As String, 種 As String, 公 As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 失敗
        Application.StatusBar = "マクロ一覧: VBA プロジェクトを読めません（［開発］→マクロのセキュリティで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を入れてください）"
        Exit Sub
    End If
    On Error GoTo 失敗
    Application.ScreenUpdating = False
    nm0 = "調査_マクロ一覧": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：マクロ一覧　" & wb.Name
    hr = 4: NC = 6
    out.Range("A4:F4").Value = Array("モジュール", "モジュールの種類", "マクロ名", "公開", "行数", "役割（頭のコメント）")
    r = hr
    For Each comp In wb.VBProject.VBComponents
        Set cm = comp.CodeModule
        n = cm.CountOfLines
        Select Case comp.Type
            Case 1: 型名 = "標準モジュール"
            Case 2: 型名 = "クラス"
            Case 3: 型名 = "フォーム"
            Case Else: 型名 = "シート／ThisWorkbook"
        End Select
        jj = 1
        Do While jj <= n And r < 5000
            proc = ""
            On Error Resume Next
            proc = cm.ProcOfLine(jj, 0)
            On Error GoTo 失敗
            If Len(proc) = 0 Then
                jj = jj + 1
            Else
                bl = cm.ProcBodyLine(proc, 0)
                decl = Trim$(cm.lines(bl, 1))
                If LCase$(Left$(decl, 7)) = "private" Then 公 = "非公開" Else 公 = "公開"
                If InStr(LCase$(decl), "function ") > 0 Then 種 = "Function" Else 種 = "Sub"
                role = ""
                k = bl + 1
                Do While k <= bl + 12 And k <= n
                    t = Trim$(cm.lines(k, 1))
                    If Left$(t, 1) <> "'" Then Exit Do
                    t = Trim$(Mid$(t, 2))
                    If Left$(t, 5) <> "依頼の語:" And Left$(t, 3) <> "扱う:" And Left$(t, 4) <> "見出し:" And Left$(t, 2) <> "形:" And Left$(t, 4) <> "選ぶ列:" And Len(t) > 0 Then
                        role = t
                        Exit Do
                    End If
                    k = k + 1
                Loop
                If Len(role) = 0 Then
                    k = cm.ProcStartLine(proc, 0)
                    Do While k < bl
                        t = Trim$(cm.lines(k, 1))
                        If Left$(t, 1) = "'" Then
                            t = Trim$(Mid$(t, 2))
                            If Len(t) > 0 Then
                                role = t
                                Exit Do
                            End If
                        End If
                        k = k + 1
                    Loop
                End If
                r = r + 1
                out.Cells(r, 1).Value = "'" & comp.Name
                out.Cells(r, 2).Value = 型名
                out.Cells(r, 3).Value = "'" & proc & IIf(種 = "Function", "（Function）", "")
                out.Cells(r, 4).Value = 公
                out.Cells(r, 5).Value = cm.ProcCountLines(proc, 0)
                out.Cells(r, 6).Value = "'" & Left$(role, 120)
                nxt = cm.ProcStartLine(proc, 0) + cm.ProcCountLines(proc, 0)
                If nxt <= jj Then nxt = jj + 1
                jj = nxt
            End If
        Loop
    Next comp
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "マクロは見つかりませんでした"
    End If
    rt = r
    out.Range("A2").Value = "マクロ " & (rt - hr) & " 本" & IIf(r >= 5000, "（5,000 本で打ち切り）", "")
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_マクロ一覧: " & (rt - hr) & " 本を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "マクロ一覧: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub ボタンとマクロの対応を一覧にする()
    ' 依頼の語: マクロの対応|ボタンとマクロ|図形とマクロ|割り当てているマクロ|割り当てられたマクロ|どのマクロが割り当て|ボタンのマクロ|ボタンや図形とマクロ
    ' 依頼の組: ボタン,図形+マクロ+付い,割り当,対応
    ' 扱う: ボタン 図形 マクロ 割り当て
    ' 見出し: なし
    ' 形: なし
    ' ブックのシートにあるボタン・図形・画像のうち、マクロが割り当てられているものを、場所・表示文字・割り当てマクロ・マクロが実在するかつきで、新しいシート「調査_ボタンとマクロ」に一覧にする。
    ' マクロの実在は、VBA プロジェクトを読める場合だけ確かめる。何も変えない（図形の位置をそろえるのは「図形を一覧にし位置をそろえる」）。
    Dim wb As Workbook, sh As Object, shp As Object, out As Worksheet, old As Object, comp As Object, cm As Object
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, NC As Long
    Dim procs As Object, okVba As Boolean, n As Long, jj As Long, proc As String, nxt As Long
    Dim act As String, maC As String, p As Long, t As String, 種 As String, 確 As String, 他 As Boolean, loc As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set procs = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number = 0 Then okVba = True
    Err.Clear
    On Error GoTo 失敗
    If okVba Then
        For Each comp In wb.VBProject.VBComponents
            Set cm = comp.CodeModule
            n = cm.CountOfLines
            jj = 1
            Do While jj <= n
                proc = ""
                On Error Resume Next
                proc = cm.ProcOfLine(jj, 0)
                On Error GoTo 失敗
                If Len(proc) = 0 Then
                    jj = jj + 1
                Else
                    procs(LCase$(proc)) = 1
                    nxt = cm.ProcStartLine(proc, 0) + cm.ProcCountLines(proc, 0)
                    If nxt <= jj Then nxt = jj + 1
                    jj = nxt
                End If
            Loop
        Next comp
    End If
    Application.ScreenUpdating = False
    nm0 = "調査_ボタンとマクロ": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：ボタン・図形とマクロの対応　" & wb.Name
    hr = 4: NC = 7
    out.Range("A4:G4").Value = Array("シート", "図形の名前", "種類", "場所", "表示文字", "割り当てマクロ", "マクロの有無")
    r = hr
    For Each sh In wb.Worksheets
        For Each shp In sh.Shapes
            act = ""
            On Error Resume Next
            act = shp.OnAction
            On Error GoTo 失敗
            If Len(act) > 0 Then
                r = r + 1
                Select Case shp.Type
                    Case 1: 種 = "図形"
                    Case 8: 種 = "フォームコントロール"
                    Case 13: 種 = "画像"
                    Case 17: 種 = "テキストボックス"
                    Case Else: 種 = "種類 " & shp.Type
                End Select
                loc = ""
                t = ""
                On Error Resume Next
                loc = shp.TopLeftCell.Address(False, False)
                t = shp.TextFrame2.TextRange.text
                If Len(t) = 0 Then t = shp.DrawingObject.Caption
                On Error GoTo 失敗
                out.Cells(r, 1).Value = "'" & sh.Name
                out.Cells(r, 2).Value = "'" & shp.Name
                out.Cells(r, 3).Value = 種
                out.Cells(r, 4).Value = loc
                out.Cells(r, 5).Value = "'" & Left$(t, 30)
                out.Cells(r, 6).Value = "'" & act
                maC = act
                p = InStrRev(maC, "!")
                If p > 0 Then maC = Mid$(maC, p + 1)
                p = InStrRev(maC, ".")
                If p > 0 Then maC = Mid$(maC, p + 1)
                maC = Replace(maC, "'", "")
                他 = (InStr(act, "!") > 0 And InStr(LCase$(act), LCase$(wb.Name)) = 0)
                If Not okVba Then
                    確 = "確認できません（VBA に触れない設定）"
                ElseIf 他 Then
                    確 = "他のブックのマクロ（未確認）"
                ElseIf procs.exists(LCase$(maC)) Then
                    確 = "あり"
                Else
                    確 = "見つからない"
                    out.Range(out.Cells(r, 1), out.Cells(r, NC)).Interior.Color = RGB(255, 199, 206)
                End If
                out.Cells(r, 7).Value = 確
            End If
        Next shp
    Next sh
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "マクロを割り当てたボタン・図形は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_ボタンとマクロ: " & (rt - hr) & " 件を一覧にしました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "ボタンとマクロ: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub シートの数式を形ごとに一覧にする()
    ' 依頼の語: 数式の一覧|数式を一覧|数式のパターン|数式が何を|何をしている数式|数式の中身|数式を説明|数式の説明
    ' 依頼の組: 数式+一覧,パターン,説明,何をし+-監査,間違い
    ' 扱う: 数式 パターン 関数 一覧
    ' 見出し: なし
    ' 形: なし
    ' 今のシートの数式を、同じ形（R1C1 の式）ごとにまとめ、セル数・最初のセル・例・使っている関数・参照しているシートを、新しいシート「調査_数式の一覧」に一覧にする。
    ' 式が「何をしているか」の言葉での説明は AI の役目なので、形と例を並べる。何も変えない。
    Dim wb As Workbook, ws As Object, out As Worksheet, old As Object
    Dim kv As Variant
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, NC As Long
    Dim fr As Range, ar As Range, v As Variant, v2 As Variant, a As Long, b As Long, key As String, total As Long, 打切 As Boolean
    Dim cnt As Object, first As Object, ex As Object, re As Object, ms As Object, m As Object, fn As Object, sm As Object, s As String, f As String, tk As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    If TypeName(ws) <> "Worksheet" Then Exit Sub
    Set cnt = CreateObject("Scripting.Dictionary"): Set first = CreateObject("Scripting.Dictionary"): Set ex = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    nm0 = "調査_数式の一覧": nM = nm0: kk = 1
    Do
        Set old = Nothing
        On Error Resume Next
        Set old = wb.Sheets(nM)
        On Error GoTo 失敗
        If old Is Nothing Then Exit Do
        If TypeName(old) = "Worksheet" Then
            If Left$(セルの字(old.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                old.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        kk = kk + 1
        nM = nm0 & "_" & kk
    Loop
    Set out = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    out.Name = nM
    out.Range("A1").Value = "調査：数式の一覧　" & ws.Name
    hr = 4: NC = 6
    out.Range("A4:F4").Value = Array("数式の形（R1C1）", "セル数", "最初のセル", "例（A1形式）", "使っている関数", "参照しているシート")
    r = hr
    Set fr = Nothing
    On Error Resume Next
    Set fr = ws.UsedRange.SpecialCells(-4123)
    On Error GoTo 失敗
    If Not fr Is Nothing Then
        For Each ar In fr.Areas
            If ar.CountLarge = 1 Then
                ReDim v(1 To 1, 1 To 1): ReDim v2(1 To 1, 1 To 1)
                v(1, 1) = ar.FormulaR1C1
                v2(1, 1) = ar.Formula
            Else
                v = ar.FormulaR1C1
                v2 = ar.Formula
            End If
            For a = 1 To UBound(v, 1)
                For b = 1 To UBound(v, 2)
                    total = total + 1
                    If total > 100000 Then
                        打切 = True
                        GoTo 集計おわり
                    End If
                    key = CStr(v(a, b))
                    If cnt.exists(key) Then
                        cnt(key) = cnt(key) + 1
                    Else
                        cnt.Add key, 1
                        first.Add key, ar.Cells(a, b).Address(False, False)
                        ex.Add key, CStr(v2(a, b))
                    End If
                Next b
            Next a
        Next ar
    End If
集計おわり:
    ' 正規表現: VBScript が外された環境では VBA 標準の RegExp に切り替える（2027 年対応）
    Set re = Nothing
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    If re Is Nothing Then Set re = 正規表現_標準版を作る()
    On Error GoTo 失敗
    If re Is Nothing Then Err.Raise 429, , "正規表現が使えません。VBScript が無効で、この Excel の VBA にも標準の RegExp がありません"
    re.Global = True
    For Each kv In cnt.keys
        key = CStr(kv)
        If r - hr >= 500 Then Exit For
        r = r + 1
        out.Cells(r, 1).Value = "'" & key
        out.Cells(r, 2).Value = cnt(key)
        out.Cells(r, 3).Value = first(key)
        out.Cells(r, 4).Value = "'" & ex(key)
        f = ex(key)
        Set fn = CreateObject("Scripting.Dictionary")
        re.Pattern = "([A-Z][A-Z0-9\._]*)\("
        Set ms = re.Execute(f)
        s = ""
        For Each m In ms
            tk = m.SubMatches(0)
            If Not fn.exists(tk) Then
                fn.Add tk, 1
                s = s & IIf(Len(s) > 0, "、", "") & tk
            End If
        Next m
        out.Cells(r, 5).Value = "'" & s
        Set sm = CreateObject("Scripting.Dictionary")
        re.Pattern = "('[^']+'|[^ =+\-*/(),;:<>&^!]+)!"
        Set ms = re.Execute(f)
        s = ""
        For Each m In ms
            tk = Replace(m.SubMatches(0), "'", "")
            If tk <> "REF" And Not sm.exists(tk) Then
                sm.Add tk, 1
                s = s & IIf(Len(s) > 0, "、", "") & tk
            End If
        Next m
        out.Cells(r, 6).Value = "'" & s
    Next kv
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, NC)).Sort Key1:=out.Cells(hr + 1, 2), Order1:=2, Header:=2
    If 打切 Then out.Range("A2").Value = "数式が多いので、先頭の 100,000 個までを調べました。"
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "このシートに数式はありません"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, NC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, NC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To NC
            If .Columns(j).ColumnWidth > 60 Then
                .Columns(j).ColumnWidth = 60
                .Columns(j).WrapText = True
            End If
            If .Columns(j).ColumnWidth < 6 Then .Columns(j).ColumnWidth = 6
        Next j
    End With
    out.Activate
    ActiveWindow.FreezePanes = False
    out.Cells(hr + 1, 1).Select
    ActiveWindow.FreezePanes = True
    out.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_数式の一覧: 数式 " & Format(total, "#,##0") & " 個を " & (rt - hr) & " 通りの形にまとめました。"
    Exit Sub
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "数式の一覧: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub 列番号を直書きした式を一覧にする()
    ' 依頼の語: 列を挿入すると壊れる式|列を挿入するとずれる式|列を追加すると壊れる式|列を入れると壊れる式|列番号の直書き|列番号を直書き|列番号が直書き|VLOOKUPの列番号|列を挿入するとずれ|列を挿入したら壊れ|列を足すとずれ|列の追加で壊れ|列番号がずれ|番号の直書き|行番号の直書き|列を入れると壊れ|列を増やすとずれ
    ' 依頼の組: 列番号,行番号,番号+直書き,ずれ,壊れ,危な,調べ,一覧,探+-削除,直して,書き換え,置換
    ' 扱う: 数式 VLOOKUP 列番号 直書き
    ' 見出し: なし
    ' 形: なし
    ' VLOOKUP・HLOOKUP・INDEX の「列番号・行番号が数字で直書きされた式」を、新しいシート「調査_列番号の直書き」に一覧にする。何も変えない。
    ' 範囲は自動で伸びるのに、この数字は列や行を挿入・削除しても直らず、黙って別の列を返す。番号が指している見出しの字も出す。
    ' （道具の col・row の insert/delete は、この番号を同じデータを指すように直す。直すのは MATCH で見出しから探す形にするのが根本）
    Dim wb As Workbook, 元 As Object, 出 As Worksheet, 古 As Object
    Dim 名 As String, 名0 As String, 回 As Long, 行 As Long, 末 As Long, 項 As Long
    Dim 域 As Range, 面 As Range, 値 As Variant, a As Long, b As Long, 式 As String, 大 As String
    Dim 関 As Variant, 位 As Long, 前 As String, 深 As Long, 引 As Long, 始 As Long, 文字 As String, 引用 As Boolean, 引用S As Boolean
    Dim 引数(1 To 9) As String, 引数数 As Long, 番号 As String, 表 As String, 種 As String, 図 As Range
    Dim 参照シート As Object, 区 As Long, 参照名 As String, 見出1 As String, 見出2 As String, 件 As Long, 打切 As Boolean
    Dim 番 As Long, 検 As Long

    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    名0 = "調査_列番号の直書き": 名 = 名0: 回 = 1
    Do
        Set 古 = Nothing
        On Error Resume Next
        Set 古 = wb.Sheets(名)
        On Error GoTo 失敗
        If 古 Is Nothing Then Exit Do
        If TypeName(古) = "Worksheet" Then
            If Left$(セルの字(古.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                古.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        回 = 回 + 1
        名 = 名0 & "_" & 回
    Loop
    Set 出 = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    出.Name = 名
    出.Range("A1").Value = "調査：列番号・行番号が直書きの式　" & wb.Name
    出.Range("A4:H4").Value = Array("シート", "セル", "関数", "数字の種類", "直書きの番号", "表の範囲", "番号が指す見出しの字", "式")
    行 = 4

    For Each 元 In wb.Worksheets
        If Left$(セルの字(元.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート
        Set 域 = Nothing
        On Error Resume Next
        Set 域 = 元.UsedRange.SpecialCells(-4123)
        On Error GoTo 失敗
        If 域 Is Nothing Then GoTo 次のシート
        For Each 面 In 域.Areas
            If 面.CountLarge = 1 Then
                ReDim 値(1 To 1, 1 To 1)
                値(1, 1) = 面.Formula
            Else
                値 = 面.Formula
            End If
            For a = 1 To UBound(値, 1)
                For b = 1 To UBound(値, 2)
                    式 = CStr(値(a, b))
                    大 = UCase$(式)
                    If InStr(大, "VLOOKUP(") > 0 Or InStr(大, "HLOOKUP(") > 0 Or InStr(大, "INDEX(") > 0 Then
                        ' "…" の中（文字として書いた VLOOKUP( など）は式ではないので、空白に伏せて探す（位置はそのまま）
                        If InStr(式, """") > 0 Then
                            引用 = False: 大 = ""
                            For 区 = 1 To Len(式)
                                文字 = Mid$(式, 区, 1)
                                If 文字 = """" Then
                                    引用 = Not 引用
                                    大 = 大 & """"
                                ElseIf 引用 Then
                                    大 = 大 & " "
                                Else
                                    大 = 大 & UCase$(文字)
                                End If
                            Next 区
                        End If
                        For Each 関 In Array("VLOOKUP", "HLOOKUP", "INDEX")
                            位 = InStr(大, 関 & "(")
                            Do While 位 > 0
                                前 = ""
                                If 位 > 1 Then 前 = Mid$(大, 位 - 1, 1)
                                If 前 Like "[A-Z0-9_.]" Then GoTo 次の呼び出し
                                始 = 位 + Len(関) + 1
                                GoSub 引数を切る
                                If 引数数 >= 2 Then
                                    表 = Trim$(引数(IIf(関 = "INDEX", 1, 2)))
                                    For 検 = 1 To 2
                                        番号 = "": 種 = ""
                                        If 関 = "VLOOKUP" And 検 = 1 And 引数数 >= 3 Then 番号 = Trim$(引数(3)): 種 = "列番号"
                                        If 関 = "HLOOKUP" And 検 = 1 And 引数数 >= 3 Then 番号 = Trim$(引数(3)): 種 = "行番号"
                                        If 関 = "INDEX" Then
                                            If 検 = 1 And 引数数 >= 2 Then 番号 = Trim$(引数(2)): 種 = "行番号"
                                            If 検 = 2 And 引数数 >= 3 Then 番号 = Trim$(引数(3)): 種 = "列番号"
                                            If 検 = 1 And 引数数 = 2 Then 種 = "行番号"
                                        End If
                                        If 番号 <> "" And Len(番号) <= 6 And Not (番号 Like "*[!0-9]*") Then
                                            If CLng(番号) >= 1 Then
                                                GoSub 見出しを引く
                                                項 = 項 + 1
                                                If 行 < 5000 Then
                                                    行 = 行 + 1
                                                    出.Cells(行, 1).Value = "'" & 元.Name
                                                    出.Cells(行, 2).Value = 面.Cells(a, b).Address(False, False)
                                                    出.Cells(行, 3).Value = 関
                                                    出.Cells(行, 4).Value = 種
                                                    出.Cells(行, 5).Value = CLng(番号)
                                                    出.Cells(行, 6).Value = "'" & 表
                                                    出.Cells(行, 7).Value = "'" & 見出1 & IIf(見出2 <> "", "（1 つ上の行: " & 見出2 & "）", "")
                                                    出.Cells(行, 8).Value = "'" & 式
                                                Else
                                                    打切 = True
                                                End If
                                            End If
                                        End If
                                    Next 検
                                End If
次の呼び出し:
                                位 = InStr(位 + 1, 大, 関 & "(")
                            Loop
                        Next 関
                    End If
                Next b
            Next a
        Next 面
次のシート:
    Next 元

    If 打切 Then 出.Range("A2").Value = "件数が多いので、先頭の 5,000 件までを出しました。"
    If 行 = 4 Then
        行 = 5
        出.Cells(行, 1).Value = "列番号・行番号が数字で直書きされた VLOOKUP・HLOOKUP・INDEX は見つかりませんでした"
    Else
        出.Range("A3").Value = "列や行を挿入・削除すると、この数字は自動では直りません（範囲だけが伸びる）。道具の col・row は直しますが、Excel で手で挿入する前に確かめてください。根本は MATCH で見出しから探す形：=VLOOKUP(キー, 表, MATCH(""見出し"", 見出しの行, 0), FALSE)"
    End If
    末 = 行
    With 出
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(4, 1), .Cells(4, 8))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(4, 1), .Cells(末, 8))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For 区 = 1 To 8
            If .Columns(区).ColumnWidth > 60 Then
                .Columns(区).ColumnWidth = 60
                .Columns(区).WrapText = True
            End If
        Next 区
    End With
    出.Activate
    出.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_列番号の直書き: " & 項 & " 件を一覧にしました（何も変えていません）。"
    Exit Sub

引数を切る:
    ' 式 の 始（"(" の次）から、対応する ")" までの引数を 引数(1..) に・数を 引数数 に（かっこと "…" と '…' の中のコンマでは切らない）
    引数数 = 0
    For 区 = 1 To 9: 引数(区) = "": Next 区
    深 = 1: 引 = 始: 引用 = False: 引用S = False
    For 区 = 始 To Len(式)
        文字 = Mid$(式, 区, 1)
        If 引用 Then
            If 文字 = """" Then 引用 = False
        ElseIf 引用S Then
            If 文字 = "'" Then 引用S = False
        ElseIf 文字 = """" Then
            引用 = True
        ElseIf 文字 = "'" Then
            引用S = True
        ElseIf 文字 = "(" Or 文字 = "{" Then
            深 = 深 + 1
        ElseIf 文字 = ")" Or 文字 = "}" Then
            深 = 深 - 1
            If 深 = 0 Then
                引数数 = 引数数 + 1
                If 引数数 <= 9 Then 引数(引数数) = Mid$(式, 引, 区 - 引)
                Exit For
            End If
        ElseIf 文字 = "," And 深 = 1 Then
            引数数 = 引数数 + 1
            If 引数数 <= 9 Then 引数(引数数) = Mid$(式, 引, 区 - 引)
            引 = 区 + 1
        End If
    Next 区
    If 深 > 0 Then 引数数 = 0
    If 引数数 > 9 Then 引数数 = 9
    Return

見出しを引く:
    ' 表（範囲の文字）の 番号 番目の列・行の見出しの字を 見出1（範囲の 1 つ目）・見出2（1 つ上・1 つ左）に。読めなければ ""
    見出1 = "": 見出2 = ""
    Set 参照シート = 元
    参照名 = 表
    区 = InStrRev(参照名, "!")
    If 区 > 0 Then
        Set 参照シート = Nothing
        On Error Resume Next
        Set 参照シート = wb.Worksheets(Replace(Left$(参照名, 区 - 1), "'", ""))
        On Error GoTo 失敗
        参照名 = Mid$(参照名, 区 + 1)
    End If
    If Not 参照シート Is Nothing Then
        Set 図 = Nothing
        On Error Resume Next
        Set 図 = 参照シート.Range(参照名)
        On Error GoTo 失敗
        If Not 図 Is Nothing Then
            番 = CLng(番号)
            On Error Resume Next
            If 種 = "列番号" Then
                If 番 <= 図.Columns.Count Then
                    見出1 = CStr(図.Cells(1, 番).text)
                    If 図.Row > 1 Then 見出2 = CStr(参照シート.Cells(図.Row - 1, 図.Column + 番 - 1).text)
                End If
            Else
                If 番 <= 図.rows.Count Then
                    見出1 = CStr(図.Cells(番, 1).text)
                    If 図.Column > 1 Then 見出2 = CStr(参照シート.Cells(図.Row + 番 - 1, 図.Column - 1).text)
                End If
            End If
            On Error GoTo 失敗
        End If
    End If
    Return
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "列番号の直書き: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub ブックに隠れた物を一覧にする()
    ' 依頼の語: 隠れた物を|隠れた物は|隠れているもの|隠れている物|隠してある物|隠してあるもの|ブックの診断|ブックを診断|隠れた情報|何が隠れ|見えない物を|見えないものを|ブックに隠れ
    ' 依頼の組: 隠れ,隠し,見えない+調べ,一覧,探,診断,洗い出,見つけ+-削除,消して,表示に戻,再表示,直して,行と列,シートと行
    ' 扱う: 隠れた物 診断 名前 スタイル プロパティ 引き継ぎ
    ' 見出し: なし
    ' 形: なし
    ' 画面を見ても分からない物を、新しいシート「調査_隠れた物」に一覧にする。何も変えない（セルの値は出さず、場所と種類を出す）。
    '   非表示・極秘のシート／オブジェクト名が既定でないシート／セルに付いていない名前（名前ボックスに出ない）・非表示の名前／
    '   ユーザー定義のセルスタイルとテーブルスタイル／ブックのプロパティ（作成者・ユーザー設定など）／白い文字のセル／
    '   セル内改行で 2 行目が見えないセル／入力規則のメッセージ（空のセルにも付く）／文字を別の字に見せる表示形式／非表示の図形／離れた所にぽつんとある最後のセル
    ' （オフィス田中「ブックの中に 10 種類の動物が隠れています」の隠し場所を、引き継ぎ・点検で漏れなく洗う）
    Dim wb As Workbook, 出 As Worksheet, 古 As Object, 元 As Object
    Dim 名 As String, 名0 As String, 回 As Long, 行 As Long, 件 As Long, 打切 As Boolean
    Dim 種 As String, 場 As String, 容 As String, 法 As String
    Dim 数 As Long, 区 As Long
    Dim 域 As Range, 面 As Range, 値 As Variant, 末 As Range, 見 As Range, 組 As Variant
    Dim 名前 As Object, 型 As Object, 図 As Object, 参照 As String, 短 As String
    Dim 辞 As Object, 書 As String, 部() As String, 検 As Long, 検行 As Long, 改先 As String, 改容 As String

    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    名0 = "調査_隠れた物": 名 = 名0: 回 = 1
    Do
        Set 古 = Nothing
        On Error Resume Next
        Set 古 = wb.Sheets(名)
        On Error GoTo 失敗
        If 古 Is Nothing Then Exit Do
        If TypeName(古) = "Worksheet" Then
            If Left$(セルの字(古.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                古.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        回 = 回 + 1
        名 = 名0 & "_" & 回
    Loop
    Set 出 = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    出.Name = 名
    出.Range("A1").Value = "調査：隠れた物　" & wb.Name
    出.Range("A4:D4").Value = Array("種類", "場所", "内容（値は出さず、形と場所）", "見つけ方")
    行 = 4

    ' --- 1. シート（非表示・極秘・オブジェクト名）
    For Each 元 In wb.Sheets
        If 元.Visible <> -1 Then
            If 元.Visible = 2 Then
                種 = "極秘のシート（VeryHidden）"
                法 = "［再表示］の一覧にも出ない。VBE のプロパティの Visible か、［校閲］→［ブックの統計情報］の枚数で気づく"
            Else
                種 = "非表示のシート"
                法 = "シート見出しを右クリック→［再表示］（非表示が 1 枚でもあれば押せる）"
            End If
            場 = 元.Name
            容 = TypeName(元)
            If TypeName(元) = "Worksheet" Then 容 = "使用範囲 " & 元.UsedRange.Address(False, False)
            GoSub 行を足す
        End If
        If TypeName(元) = "Worksheet" Then
            If 元.CodeName <> "" And Left$(元.Name, 3) <> "調査_" And Not (元.CodeName Like "Sheet#*" Or 元.CodeName Like "シート#*") Then
                種 = "既定でないオブジェクト名"
                場 = 元.Name
                容 = "オブジェクト名 " & 元.CodeName & "（表示名とは別。VBE の (オブジェクト名) プロパティ）"
                法 = "VBE（Alt+F11）のプロジェクトエクスプローラーで、シート名の左側の名前"
                GoSub 行を足す
            End If
        End If
    Next 元

    ' --- 2. 名前（セルに付いていない・非表示）
    For Each 名前 In wb.names
        参照 = "": 短 = ""
        On Error Resume Next
        参照 = 名前.RefersTo
        短 = 名前.Name
        On Error GoTo 失敗
        If InStr(短, "!") > 0 Then 短 = Mid$(短, InStr(短, "!") + 1)
        If Not (短 Like "_xlnm*" Or 短 Like "Print_*" Or 短 Like "_FilterDatabase" Or 短 Like "_xlfn*" Or 短 Like "_xlpm*" Or 短 Like "Consolidate_Area" Or 短 Like "Sheet_Title") Then
            If 名前.Visible = False Then
                種 = "非表示の名前"
                場 = 名前.Name
                容 = 参照
                法 = "［名前の管理］にも出ない（VBA で Visible=False にした名前）。VBA の Names か『ワークシート診断』で"
                GoSub 行を足す
            ElseIf Left$(参照, 1) = "=" And InStr(参照, "!") = 0 And InStr(参照, "#REF") = 0 Then
                種 = "セルに付いていない名前（定数・式）"
                場 = 名前.Name
                容 = 参照
                法 = "名前ボックス（左上）の一覧には出ない。［数式］→［名前の管理］で見る"
                GoSub 行を足す
            End If
        End If
    Next 名前

    ' --- 3. ユーザー定義のスタイル（セルスタイル・テーブルスタイル）
    数 = 0
    On Error Resume Next
    For Each 型 In wb.Styles
        ' 貼り付けで自動で増える名前（標準 2・桁区切り 2・スタイル 1 など）は隠し物ではないので数えない（実際の仕事のブック 15 本で 10 本に出た）
        If 型.BuiltIn = False And Not (型.Name Like "標準*" Or 型.Name Like "桁区切り*" Or 型.Name Like "スタイル #*" Or 型.Name Like "ハイパーリンク*" Or 型.Name Like "Excel Built-in*" Or 型.Name Like "Normal*" Or 型.Name Like "通貨*" Or 型.Name Like "パーセント*") Then
            数 = 数 + 1
            If 数 <= 40 Then
                種 = "ユーザー定義のセルスタイル"
                場 = "ブック全体"
                容 = 型.Name
                法 = "［ホーム］→［セルのスタイル］の「ユーザー設定」"
                GoSub 行を足す
            End If
        End If
    Next 型
    If 数 > 40 Then
        種 = "ユーザー定義のセルスタイル（続き）": 場 = "ブック全体": 容 = "ほかに " & (数 - 40) & " 個": 法 = "「ブックが重い原因を一覧にする」でも数が分かる"
        GoSub 行を足す
    End If
    For Each 型 In wb.TableStyles
        If 型.BuiltIn = False Then
            種 = "ユーザー定義のテーブルスタイル"
            場 = "ブック全体"
            容 = 型.Name
            法 = "［テーブルとして書式設定］の一覧の下のほうの「ユーザー設定」"
            GoSub 行を足す
        End If
    Next 型
    On Error GoTo 失敗

    ' --- 4. ブックのプロパティ
    On Error Resume Next
    ' 作成者・最終更新者は、ほぼ全部のブックに入っていて見つけ物にならないので出さない（個人名を消すなら［ドキュメント検査］）
    For Each 組 In Array("Title", "Subject", "Keywords", "Comments", "Company", "Manager", "Category", "Hyperlink base")
        書 = ""
        書 = CStr(wb.BuiltinDocumentProperties(組).Value)
        If 書 <> "" Then
            種 = "ブックのプロパティ"
            場 = 組
            容 = "（" & Len(書) & " 字）" & Left$(書, 30)
            法 = "［ファイル］→［情報］→［プロパティ］→［詳細プロパティ］。エクスプローラーのポップアップにも出る"
            GoSub 行を足す
        End If
    Next 組
    For 区 = 1 To wb.CustomDocumentProperties.Count
        種 = "ブックのプロパティ（ユーザー設定）"
        場 = wb.CustomDocumentProperties(区).Name
        容 = Left$(CStr(wb.CustomDocumentProperties(区).Value), 30)
        法 = "［詳細プロパティ］の［ユーザー設定］タブ（エクスプローラーには出ない）"
        GoSub 行を足す
    Next 区
    On Error GoTo 失敗

    ' --- 5. シートごとの隠し場所
    For Each 元 In wb.Worksheets
        If Left$(セルの字(元.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート
        Set 域 = Nothing
        On Error Resume Next
        Set 域 = 元.UsedRange
        Set 末 = 元.Cells.SpecialCells(11)           ' xlCellTypeLastCell（Ctrl+End の行き先）
        On Error GoTo 失敗
        If 域 Is Nothing Then GoTo 次のシート

        ' 5a. 離れた所にぽつんとある最後のセル
        If Not 末 Is Nothing Then
            If セルの字(末.Value) <> "" Then
                If Application.WorksheetFunction.CountA(元.rows(末.Row)) = 1 And Application.WorksheetFunction.CountA(元.Columns(末.Column)) = 1 And (末.Row > 20 Or 末.Column > 10) Then
                    種 = "離れた所にぽつんとあるセル"
                    場 = 元.Name & "!" & 末.Address(False, False)
                    容 = "使用範囲の最後（Ctrl+End の行き先）。その行にも列にも、ほかの値が無い（" & Len(CStr(末.Value)) & " 字）"
                    法 = "Ctrl+End を押す"
                    GoSub 行を足す
                End If
            End If
        End If

        ' 5b. 白い文字・文字を別の字に見せる表示形式（先頭 3000 セルまで）・セル内改行で 2 行目が見えない
        Set 辞 = CreateObject("Scripting.Dictionary")
        数 = 0: 検 = 0: 検行 = 0
        On Error Resume Next
        For Each 見 In 域.Cells
            数 = 数 + 1
            If 数 > 3000 Then Exit For
            値 = 見.Value
            If Not IsError(値) Then
                If セルの字(値) <> "" Then
                    If 見.Font.Color = 16777215 And 見.Interior.ColorIndex = -4142 Then
                        If 検 < 15 Then
                            種 = "白い文字のセル（背景なし）"
                            場 = 元.Name & "!" & 見.Address(False, False)
                            容 = "値あり（" & Len(CStr(値)) & " 字）。文字の色が白で、塗りつぶしが無い"
                            法 = "Ctrl+End・［検索と置換］・数式バーで見る"
                            GoSub 行を足す
                        End If
                        検 = 検 + 1
                    End If
                    If VarType(値) = vbString Then
                        If InStr(値, vbLf) > 0 Then
                            If 見.WrapText = False Or 見.RowHeight < (UBound(Split(値, vbLf)) + 1) * 見.Font.Size * 1.15 Then
                                ' 1 シートに何十もあることがある（データの中の複数行の文）ので、行は 1 つにまとめて出す
                                If 検行 = 0 Then 改先 = 見.Address(False, False): 改容 = Replace(Left$(CStr(値), 30), vbLf, "[改行]")
                                検行 = 検行 + 1
                            End If
                        End If
                    End If
                End If
            End If
            書 = 見.NumberFormat
            If InStr(書, ";") > 0 Then
                If Not 辞.exists(書) Then
                    辞.Add 書, 見.Address(False, False)
                    部 = Split(書, ";")
                    If UBound(部) >= 3 Then
                        If 部(3) <> "" And InStr(部(3), "@") = 0 Then
                            種 = "文字を別の字に見せる表示形式"
                            場 = 元.Name & "!" & 見.Address(False, False)
                            容 = "表示形式 " & 書 & "（文字の欄に @ が無い＝何を入れても同じ字に見える）"
                            法 = "［セルの書式設定］→［ユーザー定義］"
                            GoSub 行を足す
                        End If
                    End If
                    If 書 = ";;;" Or 書 Like ";;;*" Then
                        種 = "値を隠す表示形式（;;;）"
                        場 = 元.Name & "!" & 見.Address(False, False)
                        容 = "表示形式 " & 書
                        法 = "［セルの書式設定］→［ユーザー定義］。数式バーには値が出る"
                        GoSub 行を足す
                    End If
                End If
            End If
        Next 見
        On Error GoTo 失敗
        If 検行 > 0 Then
            種 = "セル内改行で見えない行があるセル"
            場 = 元.Name & "!" & 改先
            容 = 検行 & " セル（先頭 " & 改先 & "「" & 改容 & "」）"
            法 = "数式バーを広げる（Ctrl+Shift+U）か、行の高さを広げる。LEN で字数がずれる"
            GoSub 行を足す
        End If

        ' 5c. 入力規則のメッセージ（空のセルにも付く）
        Set 面 = Nothing
        On Error Resume Next
        Set 面 = 元.Cells.SpecialCells(-4174)       ' xlCellTypeAllValidation
        On Error GoTo 失敗
        If Not 面 Is Nothing Then
            数 = 0
            For Each 域 In 面.Areas
                数 = 数 + 1
                If 数 > 300 Then Exit For
                容 = ""
                On Error Resume Next
                If CStr(域.Cells(1, 1).Validation.InputTitle & 域.Cells(1, 1).Validation.InputMessage) <> "" Then 容 = "入力時メッセージ「" & Left$(域.Cells(1, 1).Validation.InputTitle & " " & 域.Cells(1, 1).Validation.InputMessage, 40) & "」"
                If CStr(域.Cells(1, 1).Validation.ErrorTitle & 域.Cells(1, 1).Validation.ErrorMessage) <> "" Then 容 = 容 & " エラーメッセージ「" & Left$(域.Cells(1, 1).Validation.ErrorTitle & " " & 域.Cells(1, 1).Validation.ErrorMessage, 40) & "」"
                On Error GoTo 失敗
                If 容 <> "" Then
                    種 = "入力規則のメッセージ"
                    場 = 元.Name & "!" & 域.Address(False, False) & IIf(Application.WorksheetFunction.CountA(域) = 0, "（セルは空）", "")
                    法 = "［データの入力規則］の［入力時メッセージ］［エラーメッセージ］タブ。［検索と選択］→［条件を選択してジャンプ］→［入力規則］"
                    GoSub 行を足す
                End If
            Next 域
        End If

        ' 5d. 非表示の図形
        For Each 図 In 元.Shapes
            If 図.Visible = False Then
                種 = "非表示の図形"
                場 = 元.Name & "!" & 図.Name
                容 = "種類 " & 図.Type
                法 = "［ホーム］→［検索と選択］→［選択ウィンドウ］の目のマーク"
                GoSub 行を足す
            End If
        Next 図
次のシート:
    Next 元

    If 打切 Then 出.Range("A2").Value = "件数が多いので、先頭の 3,000 件までを出しました。"
    出.Range("A3").Value = "白い文字・表示形式・セル内改行の検査は、各シートの先頭 3,000 セルまでです。セルの値は出さず（場所と字数だけ）、ブックのプロパティ・名前・スタイルは中身の先頭を出しています。"
    If 行 = 4 Then
        行 = 5
        出.Cells(行, 1).Value = "隠れた物は見つかりませんでした"
    End If
    With 出
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(4, 1), .Cells(4, 4))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(4, 1), .Cells(行, 4))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For 区 = 1 To 4
            If .Columns(区).ColumnWidth > 60 Then
                .Columns(区).ColumnWidth = 60
                .Columns(区).WrapText = True
            End If
        Next 区
    End With
    出.Activate
    出.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_隠れた物: " & 件 & " 件を一覧にしました（何も変えていません）。"
    Exit Sub

行を足す:
    If 行 < 3000 Then
        行 = 行 + 1
        出.Cells(行, 1).Value = 種
        出.Cells(行, 2).Value = "'" & 場
        出.Cells(行, 3).Value = "'" & 容
        出.Cells(行, 4).Value = 法
    Else
        打切 = True
    End If
    件 = 件 + 1
    Return
失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "隠れた物: 調べられませんでした（" & Err.Description & "）"
End Sub

Sub ブックの構造を一覧にする()
    ' 依頼の語: ブックの構造|ブック全体の構造|ブックの仕組み|ブックの作り|ワークシート診断|診断ツール|中身を出さず|値を出さずに|AIに渡せる|ＡＩに渡せる|構造だけ
    ' 依頼の組: ブック+構造,仕組み,作り,中身を出さず,値を出さず+調べ,一覧,教え,まとめ,出して,見せ,診断,洗い出+-削除,消して,直して
    ' 扱う: 構造 診断 数式 名前 条件付き書式 入力規則 テーブル ピボット 図形 クエリ 接続 リンク VBA 引き継ぎ
    ' 見出し: なし
    ' 形: なし
    ' ブックの仕組みを、セルの値を出さずに新しいシート「調査_構造」に 1 枚の表で一覧にする。何も変えない。
    '   概要（件数）／気をつける所（なぜ・Excel のどこで確かめるか）／シート／ユーザー定義の表示形式とスタイル／名前／
    '   数式（同じ形の式は 1 行にまとめ、結果は型だけ）／関数／参照／条件付き書式／入力規則／テーブル／ピボット／
    '   図形・グラフ・コメント（文字は字数だけ）／クエリ（M）／接続（パスワードは伏せる）／ハイパーリンク／VBA の点検
    '   式の中の文字・入力規則のリスト・テーブルの列名は表の形として出す。このシートをコピーして AI に貼れば、中身を出さずに相談できる
    ' （オフィス田中・田中亨さんのワークシート診断ツールの項目に学んだ。田中さんのツールも Excel だけで動く）
    Dim wb As Workbook, 出 As Worksheet, 古 As Object, 元 As Object, ws As Worksheet
    Dim 詳 As Collection, 点 As Collection, 要 As Collection, 列 As Collection
    Dim 区 As String, 項 As String, 場 As String, 容 As String, 法 As String
    Dim 名 As String, 名0 As String, 回 As Long, 行 As Long, 全 As Long, 打切 As Boolean
    Dim 域 As Range, 面 As Range, 交 As Range, 見 As Range, 柱 As Range
    Dim fa As Variant, fr As Variant, fv As Variant, v As Variant, 組 As Variant, 並 As Variant
    Dim i As Long, j As Long, k As Long, r0 As Long, c0 As Long, nr As Long, NC As Long
    Dim 一括 As Boolean, 式 As String, 式R As String, 種 As String, 鍵 As Variant
    Dim 型数 As Object, 型先 As Object, 型末 As Object, 型式 As Object, 型種 As Object, 辞 As Object
    Dim 関辞 As Object, 関全 As Object, 参辞 As Object, 名辞 As Object
    Dim 字 As String, 語 As String, 括 As Boolean, p As Long, q As Long, 前 As String, 印 As String
    Dim 式数 As Long, 外数 As Long, 隠式 As Long, 読めず As Long, 解除 As Double, 結合 As Long, 規数 As Double, 規全 As Double
    Dim シ数 As Long, シ隠 As Long, グ数 As Long, 保数 As Long, 揮発 As String
    Dim 条数 As Long, 条全 As Long, 図 As Object, 親 As String, 図種 As Long, 図類 As String, 文字数 As Long
    Dim 図表 As Long, 図絵 As Long, 図箱 As Long, 図注 As Long, 図AX As Long, 図他 As Long, 図上 As Long, 幽 As String, 幽数 As Long
    Dim 注数 As Long, 表数 As Long, 回転 As Long, リ数 As Long, ク数 As Long, 接数 As Long, 名数 As Long, 名本 As Long
    Dim 壊名 As String, 隠名 As String, 残規 As String, 残規数 As Long
    Dim 名前 As Object, 型 As Object, lo As Object, pt As Object, 欄 As Object, 繋 As Object, 問 As Object, hl As Object, 系 As Object
    Dim 部品 As Object, 注 As Object, 参照 As String, 短 As String, 範 As String
    Dim 殻 As Object, 箱 As Object, 中 As Object, 一時 As String, 書式 As String, 書式数 As Long, 流 As Object, xml As String
    Dim cm As Object, 行数 As Long, 手 As String, 手種 As Long, 始 As Long, 長 As Long, 本文 As String, 宣言 As String
    Dim 当 As String, 当辞 As Object, マ数 As Long, 標外 As String, 標外数 As Long
    Dim 手名 As String, 頭 As String, 規計 As Double, 時 As Single, 書辞 As Object, 位 As Long
    Dim 出力() As Variant

    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    Set 詳 = New Collection: Set 点 = New Collection: Set 要 = New Collection
    Set 関全 = CreateObject("Scripting.Dictionary")

    ' --- 出す先のシート（前の「調査：」のシートは作り直す）
    名0 = "調査_構造": 名 = 名0: 回 = 1
    Do
        Set 古 = Nothing
        On Error Resume Next
        Set 古 = wb.Sheets(名)
        On Error GoTo 失敗
        If 古 Is Nothing Then Exit Do
        If TypeName(古) = "Worksheet" Then
            If Left$(CStr(古.Range("A1").Value), 3) = "調査：" Then
                Application.DisplayAlerts = False
                古.Delete
                Application.DisplayAlerts = True
                Exit Do
            End If
        End If
        回 = 回 + 1
        名 = 名0 & "_" & 回
    Loop

    ' === 1. シートごと（式・結合・ロック・条件付き書式・入力規則・テーブル・ピボット・図形・コメント・リンク）
    For Each 元 In wb.Sheets
        If TypeName(元) <> "Worksheet" Then
            グ数 = グ数 + 1
            区 = "シート": 項 = "グラフシート": 場 = 元.Name: 容 = IIf(元.Visible = -1, "表示", IIf(元.Visible = 2, "完全に非表示", "非表示")): 法 = ""
            GoSub 詳を足す
            GoTo 次のシート
        End If
        Set ws = 元
        If Left$(CStr(ws.Range("A1").Value), 3) = "調査：" Then GoTo 次のシート
        シ数 = シ数 + 1
        If ws.Visible <> -1 Then シ隠 = シ隠 + 1
        If ws.Visible = 2 Then
            項 = "シート「" & ws.Name & "」が完全に非表示"
            容 = "シート見出しを右クリックした［再表示］の一覧にも出ない＝あることに誰も気づかない"
            法 = "Alt+F11 → プロジェクトのシート → プロパティの Visible"
            GoSub 点を足す
        End If
        If ws.ProtectContents Then 保数 = 保数 + 1
        Set 域 = ws.UsedRange
        nr = 域.rows.Count: NC = 域.Columns.Count
        r0 = 域.Row: c0 = 域.Column

        位 = 詳.Count + 1
        ' 1a. 式（一括で A1・R1C1・値を読む。保護で隠した式があると一括は拒まれる＝1 セルずつ）
        Set 型数 = CreateObject("Scripting.Dictionary"): Set 型先 = CreateObject("Scripting.Dictionary")
        Set 型末 = CreateObject("Scripting.Dictionary"): Set 型式 = CreateObject("Scripting.Dictionary")
        Set 型種 = CreateObject("Scripting.Dictionary"): Set 関辞 = CreateObject("Scripting.Dictionary")
        隠式 = 0: 読めず = 0: 一括 = False
        If CDbl(nr) * NC <= 1000000# Then
            On Error Resume Next
            If nr * NC = 1 Then
                ReDim fa(1 To 1, 1 To 1): ReDim fr(1 To 1, 1 To 1): ReDim fv(1 To 1, 1 To 1)
                fa(1, 1) = 域.Formula: fr(1, 1) = 域.FormulaR1C1: fv(1, 1) = 域.Value
            Else
                fa = 域.Formula: fr = 域.FormulaR1C1: fv = 域.Value
            End If
            一括 = (Err.Number = 0)
            Err.Clear
            On Error GoTo 失敗
        End If
        If 一括 Then
            For i = 1 To nr
                For j = 1 To NC
                    If VarType(fa(i, j)) = vbString Then
                        If Left$(fa(i, j), 1) = "=" Then
                            式 = fa(i, j): 式R = fr(i, j): v = fv(i, j)
                            Set 見 = ws.Cells(r0 + i - 1, c0 + j - 1)
                            GoSub 式を数える
                        End If
                    End If
                Next j
            Next i
        ElseIf CDbl(nr) * NC <= 200000# Then
            For Each 見 In 域.Cells
                On Error Resume Next
                If 見.HasFormula Then
                    If ws.ProtectContents And 見.FormulaHidden Then 隠式 = 隠式 + 1
                    式 = "": 式 = 見.Formula
                    If 式 = "" Then
                        読めず = 読めず + 1
                    Else
                        式R = 見.FormulaR1C1: v = 見.Value
                        On Error GoTo 失敗
                        GoSub 式を数える
                    End If
                End If
                On Error GoTo 失敗
            Next 見
        End If
        式数 = 式数 + 読めず
        For Each 鍵 In 型数.keys
            印 = ""
            Set 見 = ws.Range(型先(鍵))
            On Error Resume Next
            If 見.HasArray Then
                If 見.HasSpill Then
                    印 = "スピル " & 見.SpillingToRange.Address(False, False)
                Else
                    印 = "配列数式（CSE）"
                End If
            ElseIf 見.HasSpill Then
                印 = "スピル " & 見.SpillingToRange.Address(False, False)
            End If
            On Error GoTo 失敗
            式 = 型式(鍵): GoSub 関数を拾う
            並 = Array("NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO", "RANDARRAY")
            For k = 0 To UBound(並)
                If InStr("|" & 当 & "|", "|" & 並(k) & "|") > 0 Then
                    印 = 印 & IIf(印 = "", "", "・") & "揮発関数 " & 並(k)
                    If InStr("・" & 揮発 & "・", "・" & 並(k) & "・") = 0 Then 揮発 = 揮発 & IIf(揮発 = "", "", "・") & 並(k)
                End If
            Next k
            容 = ""
            For Each 組 In 型種(鍵).keys
                容 = 容 & IIf(容 = "", "", "・") & 組 & IIf(型数(鍵) > 1, " " & 型種(鍵)(組), "")
            Next 組
            区 = "数式": 項 = 型数(鍵) & " セル"
            場 = ws.Name & "!" & 型先(鍵) & IIf(型数(鍵) > 1, "～" & 型末(鍵), "")
            法 = "結果の型: " & 容 & IIf(印 = "", "", "／" & 印)
            容 = 型式(鍵)
            GoSub 詳を足す
        Next 鍵
        If 読めず > 0 Then
            区 = "数式": 項 = 読めず & " セル": 場 = ws.Name: 容 = "（保護で隠した式＝中身は読めません。保護は外していません）": 法 = ""
            GoSub 詳を足す
        End If
        If 関辞.Count > 0 Then
            区 = "関数": 項 = 関辞.Count & " 種": 場 = ws.Name: 容 = Join(関辞.keys, "・"): 法 = ""
            GoSub 詳を足す
        End If

        ' 1b. ロック解除・結合
        解除 = 0
        On Error Resume Next
        v = 域.Locked
        If IsNull(v) Then
            For Each 柱 In 域.Columns
                v = 柱.Locked
                If IsNull(v) Then
                    For Each 見 In 柱.Cells
                        If 見.Locked = False Then 解除 = 解除 + 1
                    Next 見
                ElseIf v = False Then
                    解除 = 解除 + 柱.Cells.CountLarge
                End If
            Next 柱
        ElseIf v = False Then
            解除 = 域.CountLarge
        End If
        結合 = 0
        If IsNull(域.MergeCells) Or 域.MergeCells = True Then
            If CDbl(nr) * NC <= 200000# Then
                For Each 見 In 域.Cells
                    If 見.MergeCells Then
                        If 見.Address = 見.MergeArea.Cells(1, 1).Address Then 結合 = 結合 + 1
                    End If
                Next 見
            End If
        End If
        On Error GoTo 失敗
        If ws.ProtectContents And 解除 = 0 Then
            項 = "シート「" & ws.Name & "」は保護されていて、ロック解除のセルが 0"
            容 = "入力欄を空けずに保護している＝どのセルにも入力できない"
            法 = "［ホーム］→［書式］→［セルのロック］／［校閲］→［シート保護の解除］"
            GoSub 点を足す
        End If
        If 隠式 > 0 Then
            項 = "シート「" & ws.Name & "」に保護で隠した式 " & 隠式 & " セル"
            容 = "数式バーに式が出ない＝計算の中身を確かめられない（引き継ぎの時に困る）"
            法 = "［セルの書式設定］→［保護］→［表示しない］"
            GoSub 点を足す
        End If

        ' 1c. 条件付き書式
        条数 = 0
        On Error Resume Next
        k = ws.Cells.FormatConditions.Count
        On Error GoTo 失敗
        For i = 1 To k
            条数 = 条数 + 1
            On Error Resume Next
            Set 部品 = ws.Cells.FormatConditions(i)
            範 = "": 範 = 部品.AppliesTo.Address(False, False)
            種 = "": 種 = CStr(部品.Type)
            Select Case 種
                Case "1": 種 = "セルの値"
                Case "2": 種 = "数式"
                Case "3": 種 = "カラースケール"
                Case "4": 種 = "データバー"
                Case "5": 種 = "上位/下位"
                Case "6": 種 = "アイコン"
                Case "8": 種 = "一意/重複"
                Case "9": 種 = "文字列"
                Case "10": 種 = "空白"
                Case "11": 種 = "日付"
                Case "12": 種 = "平均より上/下"
                Case "13": 種 = "空白以外"
                Case "16": 種 = "エラー"
                Case "17": 種 = "エラー以外"
            End Select
            容 = "": 容 = CStr(部品.Formula1)
            字 = "": 字 = CStr(部品.Formula2)
            If 字 <> "" Then 容 = 容 & " ～ " & 字
            前 = ""
            前 = Choose(部品.Operator, "間", "間以外", "等しい", "等しくない", "より大きい", "より小さい", "以上", "以下")
            印 = ""
            If CStr(部品.NumberFormat) <> "" Then 印 = 印 & "・表示形式"
            If Not IsNull(部品.Font.Bold) Or Not IsNull(部品.Font.Italic) Then 印 = 印 & "・フォント"
            If 部品.Interior.ColorIndex <> -4142 And 部品.Interior.ColorIndex <> -4105 And Not IsNull(部品.Interior.ColorIndex) Then 印 = 印 & "・塗り"
            For j = 1 To 4
                If 部品.Borders(j).LineStyle <> -4142 And Not IsNull(部品.Borders(j).LineStyle) Then 印 = 印 & "・罫線": Exit For
            Next j
            括 = False
            For Each 柱 In 部品.AppliesTo.Areas
                If 柱.rows.Count >= ws.rows.Count Or 柱.Columns.Count >= ws.Columns.Count Then 括 = True
            Next 柱
            Err.Clear
            On Error GoTo 失敗
            区 = "条件付き書式": 項 = 種 & IIf(前 = "", "", "（" & 前 & "）"): 場 = ws.Name & "!" & 範 & IIf(括, " ★まるごと", "")
            法 = "書式: " & Mid$(印, 2)
            GoSub 詳を足す
            If 括 Then
                項 = "条件付き書式が列・行まるごと（" & ws.Name & "!" & 範 & "）"
                容 = "100 万行ぶん判定する＝スクロールや入力が重くなる元"
                法 = "［ホーム］→［条件付き書式］→［ルールの管理］"
                GoSub 点を足す
            End If
        Next i
        条全 = 条全 + 条数
        If 条数 >= 50 Then
            項 = "シート「" & ws.Name & "」の条件付き書式が " & 条数 & " ルール"
            容 = "コピーや行の挿入で分かれて増えたもの＝重さの元"
            法 = "［ホーム］→［条件付き書式］→［ルールの管理］（このワークシート）"
            GoSub 点を足す
        End If

        ' 1d. 入力規則（件数は使用範囲の中で数える）
        規数 = 0
        Set 面 = Nothing
        On Error Resume Next
        Set 面 = ws.Cells.SpecialCells(-4174)
        On Error GoTo 失敗
        If Not 面 Is Nothing Then
            For Each 柱 In 面.Areas
                Set 交 = Nothing
                On Error Resume Next
                Set 交 = Application.Intersect(柱, 域)
                On Error GoTo 失敗
                If 交 Is Nothing Then v = 0 Else v = 交.CountLarge
                規数 = 規数 + v: 規全 = 規全 + 柱.CountLarge
                On Error Resume Next
                With 柱.Cells(1, 1).Validation
                    j = -1: j = .Type
                    種 = "?"
                    種 = Choose(j + 1, "すべての値", "整数", "小数", "リスト", "日付", "時刻", "文字列（長さ）", "ユーザー設定")
                    容 = "": 容 = CStr(.Formula1)
                    字 = "": 字 = CStr(.Formula2)
                    If 字 <> "" Then 容 = 容 & " ～ " & 字
                    字 = "": 字 = .InputTitle & "／" & .InputMessage
                    前 = "": 前 = .ErrorTitle & "／" & .ErrorMessage
                    p = 0: p = .AlertStyle
                    q = 0: q = .IMEMode
                End With
                Err.Clear
                On Error GoTo 失敗
                括 = (柱.rows.Count >= ws.rows.Count Or 柱.Columns.Count >= ws.Columns.Count)
                If j = 0 And 字 = "／" And 前 = "／" And q = 0 Then
                    残規数 = 残規数 + 1
                    If 残規数 <= 4 Then 残規 = 残規 & IIf(残規 = "", "", "・") & ws.Name & "!" & 柱.Address(False, False)
                    区 = "入力規則": 項 = "中身の無い規則（残骸）": 場 = ws.Name & "!" & 柱.Address(False, False): 容 = "": 法 = "使用範囲の中 " & Format$(v, "#,##0") & " セル"
                    GoSub 詳を足す
                Else
                    区 = "入力規則": 項 = 種: 場 = ws.Name & "!" & 柱.Address(False, False) & IIf(括, " ★まるごと", "")
                    法 = "使用範囲の中 " & Format$(v, "#,##0") & " セル・入力時 " & 字 & "・エラー時（" & Choose(IIf(p < 1 Or p > 3, 1, p), "停止", "注意", "情報") & "）" & 前 & IIf(q > 0, "・日本語入力 " & Choose(q, "オン", "オフ", "無効", "ひらがな", "全角カタカナ", "半角カタカナ", "全角英数", "半角英数"), "")
                    GoSub 詳を足す
                    If 括 Then
                        項 = "入力規則が列・行まるごと（" & ws.Name & "!" & 柱.Address(False, False) & "）"
                        容 = "使っていない行まで規則が付く＝ファイルが重くなり、どこまでが表か分からなくなる"
                        法 = "［データ］→［データの入力規則］"
                        GoSub 点を足す
                    End If
                End If
            Next 柱
        End If

        ' 1e. シートの行
        区 = "シート": 項 = IIf(ws.Visible = -1, "表示", IIf(ws.Visible = 2, "完全に非表示", "非表示")) & IIf(ws.ProtectContents, "・保護", "")
        場 = ws.Name
        容 = "使用範囲 " & 域.Address(False, False) & "（" & nr & " 行×" & NC & " 列）"
        法 = "ロック解除 " & Format$(解除, "#,##0") & "・非表示の式 " & 隠式 & "・結合 " & 結合 & "・条件付き書式 " & 条数 & "・入力規則 " & Format$(規数, "#,##0") & " セル"
        規計 = 規計 + 規数
        If 位 <= 詳.Count Then 詳.Add Array(区, 項, 場, 容, 法), Before:=位 Else GoSub 詳を足す

        ' 1f. テーブル
        For Each lo In ws.ListObjects
            表数 = 表数 + 1
            容 = ""
            For Each 欄 In lo.ListColumns
                容 = 容 & IIf(容 = "", "", "・") & 欄.Name
            Next 欄
            字 = ""
            On Error Resume Next
            字 = lo.QueryTable.WorkbookConnection.Name
            On Error GoTo 失敗
            区 = "テーブル": 項 = lo.Name: 場 = ws.Name & "!" & lo.Range.Address(False, False)
            法 = IIf(lo.ShowTotals, "集計行あり", "") & IIf(字 = "", "", IIf(lo.ShowTotals, "・", "") & "接続 " & 字)
            容 = "列 " & 容
            GoSub 詳を足す
        Next lo

        ' 1g. ピボット
        For Each pt In ws.PivotTables
            回転 = 回転 + 1
            字 = "": 前 = ""
            On Error Resume Next
            字 = pt.SourceData
            前 = pt.PivotCache.WorkbookConnection.Name
            容 = ""
            For Each 組 In Array("RowFields", "ColumnFields", "DataFields", "PageFields")
                並 = ""
                For Each 欄 In CallByName(pt, CStr(組), VbGet)
                    並 = 並 & IIf(並 = "", "", "・") & 欄.Name
                Next 欄
                If 並 <> "" Then 容 = 容 & IIf(容 = "", "", "／") & Choose(InStr("RCDP", Left$(組, 1)), "行 ", "列 ", "値 ", "フィルター ") & 並
            Next 組
            p = 0: p = pt.PivotCache.RefreshOnFileOpen
            On Error GoTo 失敗
            区 = "ピボット": 項 = pt.Name: 場 = ws.Name & "!" & pt.TableRange2.Address(False, False)
            法 = "元 " & IIf(字 = "", 前, 字) & IIf(p, "・開くとき更新", "")
            GoSub 詳を足す
        Next pt

        ' 1h. 図形・グラフ（グループの中もたどる。文字は字数だけ）
        Set 列 = New Collection
        For Each 図 In ws.Shapes
            列.Add Array(図, "")
        Next 図
        Do While 列.Count > 0
            組 = 列(1): 列.Remove 1
            Set 図 = 組(0): 親 = 組(1)
            If 親 = "" Then 図上 = 図上 + 1
            On Error Resume Next
            図種 = -1: 図種 = 図.Type
            Select Case 図種
                Case 1: 図類 = "オートシェイプ"
                Case 3: 図類 = "グラフ": 図表 = 図表 + 1
                Case 4: 図類 = "コメント": 図注 = 図注 + 1
                Case 6: 図類 = "グループ"
                Case 8: 図類 = "フォームコントロール"
                Case 9: 図類 = "線"
                Case 12: 図類 = "ActiveX コントロール（" & 図.OLEFormat.progID & "）": 図AX = 図AX + 1
                Case 13: 図類 = "図": 図絵 = 図絵 + 1
                Case 17: 図類 = "テキストボックス": 図箱 = 図箱 + 1
                Case Else: 図類 = "種類 " & 図種
            End Select
            If 図種 <> 3 And 図種 <> 4 And 図種 <> 12 And 図種 <> 13 And 図種 <> 17 Then 図他 = 図他 + 1
            文字数 = 0
            文字数 = Len(図.TextFrame2.TextRange.text)
            If 文字数 = 0 Then 文字数 = Len(図.TextFrame.Characters.text)
            範 = "": 範 = 図.TopLeftCell.Address(False, False)
            字 = "": 字 = 図.OnAction
            前 = "": 前 = 図.Hyperlink.Address
            容 = ""
            If 図種 = 3 Then
                For Each 系 In 図.Chart.SeriesCollection
                    容 = 容 & IIf(容 = "", "", "　") & 系.Formula
                Next 系
            End If
            If 図種 <> 4 And (図.Visible = False Or 図.Height <= 1 Or 図.Width <= 1) Then
                幽数 = 幽数 + 1
                If 幽数 <= 5 Then 幽 = 幽 & IIf(幽 = "", "", "・") & ws.Name & "!" & 図.Name
            End If
            区 = "図形": 項 = 図類: 場 = ws.Name & "!" & 範 & "　" & 図.Name & IIf(親 = "", "", "（グループ " & 親 & "）")
            法 = Round(図.Height) & "×" & Round(図.Width) & IIf(字 = "", "", "・マクロ " & 字) & IIf(前 = "", "", "・リンク " & 前) & IIf(文字数 > 0, "・文字 " & 文字数 & " 字", "") & IIf(図.Visible, "", "・非表示")
            If 図種 = 6 Then
                For Each 部品 In 図.GroupItems
                    列.Add Array(部品, 図.Name)
                Next 部品
            End If
            Err.Clear
            On Error GoTo 失敗
            If 図種 <> 4 Then GoSub 詳を足す
        Loop
        On Error Resume Next
        For Each 注 In ws.Comments
            注数 = 注数 + 1
            区 = "コメント": 項 = "メモ": 場 = ws.Name & "!" & 注.Parent.Address(False, False): 容 = "": 法 = "文字 " & Len(注.text) & " 字"
            GoSub 詳を足す
        Next 注
        For Each 注 In ws.CommentsThreaded
            注数 = 注数 + 1
            区 = "コメント": 項 = "スレッド コメント": 場 = ws.Name & "!" & 注.Parent.Address(False, False): 容 = "": 法 = ""
            GoSub 詳を足す
        Next 注
        ' 1i. ハイパーリンク（表示の文字はセルの値なので出さない）
        For Each hl In ws.Hyperlinks
            リ数 = リ数 + 1
            範 = "(図形)": 範 = hl.Range.Address(False, False)
            区 = "ハイパーリンク": 項 = "リンク": 場 = ws.Name & "!" & 範
            容 = hl.Address & IIf(hl.SubAddress = "", "", "#" & hl.SubAddress): 法 = ""
            GoSub 詳を足す
        Next hl
        Err.Clear
        On Error GoTo 失敗
次のシート:
    Next 元

    ' === 2. 名前
    For Each 名前 In wb.names
        参照 = "": 短 = "": 範 = "(ブック)"
        On Error Resume Next
        参照 = 名前.RefersTo
        短 = 名前.Name
        If 名前.Parent.Name <> wb.Name Then 範 = 名前.Parent.Name
        On Error GoTo 失敗
        名数 = 名数 + 1
        字 = 短
        If InStr(字, "!") > 0 Then 字 = Mid$(字, InStr(字, "!") + 1)
        括 = (字 Like "_xl*" Or 字 = "_FilterDatabase" Or 字 Like "ExternalData_#*")
        If Not 括 Then 名本 = 名本 + 1
        区 = "名前": 項 = IIf(括, "Excel が作る名前", IIf(名前.Visible, "名前", "非表示の名前")): 場 = 短 & "（範囲 " & 範 & "）": 容 = 参照
        法 = IIf((InStr(参照, "#REF!") > 0 Or InStr(参照, "#NAME?") > 0) And Not 括, "★参照先が壊れている", "")
        GoSub 詳を足す
        If Not 括 Then
            If InStr(参照, "#REF!") > 0 Or InStr(参照, "#NAME?") > 0 Then
                項 = "名前「" & 短 & "」の参照先が壊れている（" & 参照 & "）"
                容 = "この名前を使う式・入力規則・印刷範囲が #REF! になる"
                法 = "［数式］→［名前の管理］"
                GoSub 点を足す
            ElseIf 名前.Visible = False Then
                隠名 = 隠名 & IIf(隠名 = "", "", "・") & 短
            End If
        End If
    Next 名前
    If 隠名 <> "" Then
        項 = "非表示の名前（" & Left$(隠名, 80) & "）"
        容 = "［名前の管理］にも出ない。他のブックから写ってきた古い名前が多い"
        法 = "VBA の Names（Visible）でしか見えない"
        GoSub 点を足す
    End If

    ' === 3. スタイル・ユーザー定義の表示形式（表示形式は VBA で一覧が取れないので、保存済みのファイルの styles.xml から読む）
    On Error Resume Next
    For Each 型 In wb.Styles
        If 型.BuiltIn = False Then
            区 = "スタイル": 項 = "ユーザー定義のスタイル": 場 = "ブック全体": 容 = 型.NameLocal: 法 = ""
            GoSub 詳を足す
        End If
    Next 型
    Err.Clear
    書式数 = -1
    If wb.path <> "" And (LCase$(Right$(wb.Name, 5)) = ".xlsx" Or LCase$(Right$(wb.Name, 5)) = ".xlsm") Then
        一時 = Environ$("TEMP") & "\構造_" & Format$(Now, "hhnnss") & Int(Rnd * 10000)
        Set 箱 = CreateObject("Scripting.FileSystemObject")
        箱.CreateFolder 一時
        箱.CopyFile wb.FullName, 一時 & "\b.zip"
        Set 殻 = CreateObject("Shell.Application")
        Set 中 = 殻.Namespace(CVar(一時 & "\b.zip")).ParseName("xl").GetFolder
        殻.Namespace(CVar(一時)).CopyHere 中.ParseName("styles.xml"), 4 + 16 + 1024
        時 = Timer
        Do While Timer - 時 < 5 And Timer >= 時
            If 箱.FileExists(一時 & "\styles.xml") Then
                If FileLen(一時 & "\styles.xml") > 0 Then Exit Do
            End If
            DoEvents
        Loop
        If 箱.FileExists(一時 & "\styles.xml") Then
            Set 流 = CreateObject("ADODB.Stream")
            流.Charset = "utf-8"
            流.Open
            流.LoadFromFile 一時 & "\styles.xml"
            xml = 流.ReadText
            流.Close
            書式数 = 0
            Set 書辞 = CreateObject("Scripting.Dictionary")
            p = InStr(xml, "<numFmt ")
            Do While p > 0
                q = InStr(p, xml, ">")
                字 = Mid$(xml, p, q - p)
                i = InStr(字, "formatCode=""")
                j = InStr(字, "numFmtId=""")
                If i > 0 And j > 0 Then
                    書式 = Mid$(字, i + 12, InStr(i + 12, 字, """") - i - 12)
                    書式 = Replace(Replace(Replace(Replace(Replace(書式, "&quot;", """"), "&lt;", "<"), "&gt;", ">"), "&apos;", "'"), "&amp;", "&")
                    If val(Mid$(字, j + 10)) >= 164 And Not 書辞.exists(書式) Then
                        書辞.Add 書式, 1
                        書式数 = 書式数 + 1
                        区 = "表示形式": 項 = "ユーザー定義の表示形式": 場 = "ブック全体": 容 = 書式: 法 = ""
                        GoSub 詳を足す
                    End If
                End If
                p = InStr(q, xml, "<numFmt ")
            Loop
        End If
        箱.DeleteFolder 一時, True
    End If
    Err.Clear
    On Error GoTo 失敗

    ' === 4. クエリ・接続・他ブックへのリンク
    On Error Resume Next
    For Each 問 In wb.Queries
        ク数 = ク数 + 1
        字 = ""
        For Each 元 In wb.Worksheets
            For Each lo In 元.ListObjects
                前 = "": 前 = lo.QueryTable.WorkbookConnection.Name
                If 前 = "クエリ - " & 問.Name Or 前 = "Query - " & 問.Name Then 字 = 字 & IIf(字 = "", "", "・") & 元.Name & "!" & lo.Range.Address(False, False)
            Next lo
        Next 元
        区 = "クエリ": 項 = 問.Name: 場 = IIf(字 = "", "（接続のみ・データモデル）", 字): 容 = Left$(問.Formula, 32000): 法 = "Power Query の M 言語"
        GoSub 詳を足す
    Next 問
    For Each 繋 In wb.Connections
        接数 = 接数 + 1
        字 = "": 前 = "": p = 0
        字 = 繋.OLEDBConnection.Connection
        If 字 = "" Then 字 = 繋.ODBCConnection.Connection
        組 = Empty: 組 = 繋.OLEDBConnection.CommandText
        If IsArray(組) Then 前 = Join(組, " ") Else 前 = CStr(組)
        p = 繋.OLEDBConnection.RefreshOnFileOpen
        ' パスワードは伏せる
        For Each 鍵 In Array("Password=", "Pwd=")
            i = InStr(1, 字, 鍵, vbTextCompare)
            If i > 0 Then
                j = InStr(i, 字, ";")
                If j = 0 Then j = Len(字) + 1
                字 = Left$(字, i + Len(鍵) - 1) & "***" & Mid$(字, j)
            End If
        Next 鍵
        区 = "接続": 項 = 繋.Name: 場 = "種類 " & 繋.Type: 容 = 字: 法 = IIf(前 = "", "", "コマンド " & 前) & IIf(p, "・開くとき更新", "")
        GoSub 詳を足す
        If p And InStr(字, "Microsoft.Mashup") = 0 Then
            項 = "開くときに更新する接続「" & 繋.Name & "」"
            容 = "接続先に届かない PC では開くたびに待たされ、エラーが出る"
            法 = "［データ］→［クエリと接続］→［プロパティ］"
            GoSub 点を足す
        End If
    Next 繋
    組 = Empty: 組 = wb.LinkSources(1)
    If IsArray(組) Then
        For i = LBound(組) To UBound(組)
            外数 = 外数 + 1
            区 = "参照": 項 = "リンク元のファイル": 場 = "ブック全体": 容 = 組(i): 法 = ""
            GoSub 詳を足す
        Next i
        項 = "他のブックを参照している（" & (UBound(組) - LBound(組) + 1) & " ファイル）"
        容 = "相手のファイルが動く・名前が変わると #REF! になり、開くたびに更新を聞かれる"
        法 = "［データ］→［リンクの編集］"
        GoSub 点を足す
    End If
    Err.Clear
    On Error GoTo 失敗

    ' === 5. VBA（モジュール・手続き・田中さんの VBA CheckList の点検）
    On Error Resume Next
    k = -1: k = wb.VBProject.VBComponents.Count
    On Error GoTo 失敗
    If k < 0 Then
        区 = "VBA": 項 = "読めません": 場 = "": 容 = "［VBA プロジェクト オブジェクト モデルへのアクセスを信頼する］が要ります": 法 = ""
        GoSub 詳を足す
    Else
        For Each 部品 In wb.VBProject.VBComponents
            Set cm = 部品.CodeModule
            行数 = cm.CountOfLines
            If 行数 > 0 Then
                宣言 = ""
                If cm.CountOfDeclarationLines > 0 Then 宣言 = cm.lines(1, cm.CountOfDeclarationLines)
                字 = ""
                For Each 並 In Split(宣言, vbCrLf)
                    If 並 Like "*Declare *" And Left$(Trim$(並), 1) <> "'" Then
                        語 = Trim$(Split(Split(Trim$(Mid$(並, InStr(並, "Declare ") + 8)), "(")(0), " Lib")(0))
                        字 = 字 & IIf(字 = "", "", "・") & Mid$(語, InStrRev(語, " ") + 1)
                    End If
                Next 並
                手 = ""
                i = cm.CountOfDeclarationLines + 1
                Do While i <= 行数
                    手種 = 0
                    手名 = cm.ProcOfLine(i, 手種)
                    If 手名 = "" Then
                        i = i + 1
                    Else
                        If 部品.Type = 1 Then マ数 = マ数 + 1
                        始 = cm.ProcStartLine(手名, 手種): 長 = cm.ProcCountLines(手名, 手種)
                        頭 = cm.lines(cm.ProcBodyLine(手名, 手種), 1)
                        本文 = cm.lines(始, 長)
                        Set 当辞 = CreateObject("Scripting.Dictionary")
                        For Each 並 In Split(本文, vbCrLf)
                            ' コメントを外し、文字列を潰す
                            語 = "": 括 = False
                            For p = 1 To Len(並)
                                前 = Mid$(並, p, 1)
                                If 前 = """" Then
                                    括 = Not 括
                                ElseIf 前 = "'" And Not 括 Then
                                    Exit For
                                ElseIf Not 括 Then
                                    語 = 語 & 前
                                End If
                            Next p
                            語 = " " & 語 & " "
                            If 語 Like "*CreateObject(*" Then 当辞("留意: CreateObject") = 1
                            If 語 Like "*[!.A-Za-z0-9_]Selection[!A-Za-z0-9_]*" Then 当辞("好ましくない: Selection") = 1
                            If 語 Like "*Range(*&*" Then 当辞("好ましくない: Range の中で文字列結合") = 1
                            If 語 Like "*[!A-Za-z0-9_]GoTo *" And Not 語 Like "*On Error*" Then 当辞("好ましくない: GoTo") = 1
                            If 語 Like "*FileSearch*" Then 当辞("互換性: FileSearch（2007 で消えた）") = 1
                            If 語 Like "*Cells.Count*" And Not 語 Like "*Cells.CountLarge*" Then 当辞("互換性: Cells.Count（CountLarge に）") = 1
                            If 語 Like "*SaveAs*" And (語 Like "*xlExcel9795*" Or 語 Like "*FileFormat:=43*") Then 当辞("互換性: Excel 95/97 形式で保存") = 1
                            If 語 Like "*ChartObjects*" Then 当辞("互換性: ChartObjects") = 1
                            If 語 Like "*[!A-Za-z0-9_]Shapes[!A-Za-z0-9_]*" Then 当辞("互換性: Shapes") = 1
                            If 語 Like "*CommandBars*" Then 当辞("互換性: CommandBars") = 1
                            If 語 Like "*PivotCaches.Add*" Or 語 Like "*PivotCaches().Add*" Then 当辞("互換性: PivotCaches.Add") = 1
                        Next 並
                        If LCase$(手名) = "auto_open" Or LCase$(手名) = "auto_close" Then 当辞("留意: 古い自動実行") = 1
                        当 = Join(当辞.keys, "／")
                        手 = 手 & IIf(手 = "", "", "・") & 手名 & "（" & 長 & "行）"
                        区 = "VBA": 項 = IIf(InStr(頭, "Private ") > 0, "Private ", "") & IIf(手種 > 0, Choose(手種, "Property Let", "Property Set", "Property Get"), IIf(頭 Like "*Function *", "Function", "Sub"))
                        場 = 部品.Name & "." & 手名: 容 = 長 & " 行": 法 = 当
                        GoSub 詳を足す
                        For Each 鍵 In 当辞.keys
                            If 鍵 Like "*FileSearch*" Or 鍵 Like "*Cells.Count*" Or 鍵 Like "*95/97*" Then
                                項 = 部品.Name & "." & 手名 & ": " & Mid$(鍵, 6)
                                容 = "今の Excel では止まるか、意図と違う形で動く"
                                法 = "Alt+F11 で " & 部品.Name & " の " & 手名
                                GoSub 点を足す
                            ElseIf 鍵 = "留意: 古い自動実行" Then
                                項 = 部品.Name & "." & 手名 & " が古い自動実行"
                                容 = "Auto_Open／Auto_Close は VBA から開いたとき（Workbooks.Open）は動かない"
                                法 = "ThisWorkbook の Workbook_Open／Workbook_BeforeClose に移す"
                                GoSub 点を足す
                            End If
                        Next 鍵
                        i = 始 + 長
                    End If
                Loop
                区 = "VBA": 項 = Choose(IIf(部品.Type = 100, 4, IIf(部品.Type > 3, 4, 部品.Type)), "標準モジュール", "クラス", "フォーム", "シート/ブック")
                場 = 部品.Name: 容 = 行数 & " 行・Option Explicit " & IIf(宣言 Like "*Option Explicit*", "あり", "なし") & IIf(字 = "", "", "・API 宣言 " & 字)
                法 = 手
                GoSub 詳を足す
            End If
        Next 部品
        On Error Resume Next
        For Each 部品 In wb.VBProject.References
            If Not 部品.BuiltIn Then
                字 = UCase$(部品.GUID)
                If 字 <> "{00020430-0000-0000-C000-000000000046}" And 字 <> "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}" And 字 <> "{0D452EE1-E08F-101A-852E-02608C4D0BB4}" Then
                    標外数 = 標外数 + 1
                    標外 = 標外 & IIf(標外 = "", "", "・") & 部品.Description & IIf(部品.IsBroken, "（参照不可）", "")
                End If
            End If
        Next 部品
        Err.Clear
        On Error GoTo 失敗
        If 標外数 > 0 Then
            区 = "VBA": 項 = "標準でない参照設定": 場 = "ブック全体": 容 = 標外: 法 = ""
            GoSub 詳を足す
            項 = "標準でない参照設定 " & 標外数 & " 本"
            容 = "その部品が入っていない PC では、関係ないマクロまでコンパイルエラーで止まる"
            法 = "Alt+F11 → ［ツール］→［参照設定］"
            GoSub 点を足す
        End If
    End If

    ' === 6. まとめの気づき
    If シ隠 > 0 Then
        項 = "非表示のシート " & シ隠 & " 枚"
        容 = "見えない所の式や値が、見えている表の結果を左右していることがある"
        法 = "シート見出しを右クリック → ［再表示］"
        GoSub 点を足す
    End If
    If 揮発 <> "" Then
        項 = "揮発関数（" & 揮発 & "）"
        容 = "開くたびに再計算される＝何も変えていなくても閉じるときに保存を聞かれ、大きい表では遅くなる"
        法 = "［数式］→［数式の表示］"
        GoSub 点を足す
    End If
    If 幽数 > 0 Then
        項 = "見えないオブジェクト " & 幽数 & " 個（" & 幽 & IIf(幽数 > 5, "…", "") & "）"
        容 = "非表示か大きさ 0 の図形。行のコピーで増え続け、ファイルを重くする"
        法 = "［ホーム］→［検索と選択］→［オブジェクトの選択と表示］"
        GoSub 点を足す
    End If
    If 図AX > 0 Then
        項 = "ActiveX コントロール " & 図AX & " 個"
        容 = "Office の更新や別の PC で動かなくなることがある"
        法 = "［開発］→［デザインモード］（フォームコントロールに替える）"
        GoSub 点を足す
    End If
    If 残規数 > 0 Then
        項 = "中身の無い入力規則 " & 残規数 & " か所（" & 残規 & IIf(残規数 > 4, "…", "") & "）"
        容 = "何も制限しない規則がコピーや貼り付けで残っている＝ファイルを重くするだけ"
        法 = "範囲を選んで ［データ］→［データの入力規則］→［すべてクリア］"
        GoSub 点を足す
    End If

    ' === 7. 概要（田中さんの診断ツールの［概要］と同じ項目）
    区 = "概要": 場 = "": 法 = ""
    項 = "ブック": 容 = wb.Name & IIf(wb.path = "", "（未保存）", "・" & Format$(FileLen(wb.FullName), "#,##0") & " バイト"): GoSub 要を足す
    字 = ""
    On Error Resume Next
    For Each 組 In Array("Author", "Last Author", "Company", "Title")
        前 = "": 前 = CStr(wb.BuiltinDocumentProperties(組).Value)
        If 前 <> "" Then 字 = 字 & IIf(字 = "", "", "・") & Choose(1 + Abs(組 = "Last Author") + 2 * Abs(組 = "Company") + 3 * Abs(組 = "Title"), "作成者", "最終更新者", "会社", "タイトル")
    Next 組
    On Error GoTo 失敗
    If 字 <> "" Then 項 = "プロパティに書かれているもの": 容 = 字 & "（中身は出しません）": GoSub 要を足す
    項 = "シート": 容 = シ数 & " 枚（非表示 " & シ隠 & "）・グラフシート " & グ数 & "・保護 " & 保数: GoSub 要を足す
    項 = "数式": 容 = 式数 & " セル（他ブックを参照 " & 外数 & " ファイル）": GoSub 要を足す
    項 = "名前": 容 = 名数 & " 個（Excel が作る名前を除くと " & 名本 & "）": GoSub 要を足す
    項 = "条件付き書式・入力規則": 容 = 条全 & " ルール・入力規則 " & Format$(規計, "#,##0") & " セル（使用範囲の中・範囲の全体では " & Format$(規全, "#,##0") & "）": GoSub 要を足す
    項 = "ユーザー定義の表示形式": 容 = IIf(書式数 < 0, "（保存済みの xlsx/xlsm から読みます）", 書式数 & " 個"): GoSub 要を足す
    項 = "テーブル・ピボット・クエリ・接続": 容 = 表数 & "・" & 回転 & "・" & ク数 & "・" & 接数: GoSub 要を足す
    項 = "オブジェクト": 容 = 図上 & " 個（グラフ " & 図表 & "・図 " & 図絵 & "・テキストボックス " & 図箱 & "・ActiveX " & 図AX & "・コメント " & 注数 & " セル）": GoSub 要を足す
    項 = "ハイパーリンク": 容 = リ数 & " 個": GoSub 要を足す
    項 = "マクロ": 容 = IIf(マ数 > 0, "あり（標準モジュールの手続き " & マ数 & "）", "なし") & IIf(標外数 > 0, "・標準でない参照設定 " & 標外数, ""): GoSub 要を足す

    ' === 8. 書き出す（概要 → 気をつける所 → 詳しく）
    Set 出 = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    出.Name = 名
    全 = 要.Count + 点.Count + 詳.Count
    If 全 > 60000 Then 全 = 60000: 打切 = True
    ReDim 出力(1 To 全, 1 To 5)
    行 = 0
    For Each 組 In 要
        行 = 行 + 1
        For j = 0 To 4: 出力(行, j + 1) = 組(j): Next j
    Next 組
    For Each 組 In 点
        行 = 行 + 1
        For j = 0 To 4: 出力(行, j + 1) = 組(j): Next j
    Next 組
    For Each 組 In 詳
        If 行 >= 全 Then Exit For
        行 = 行 + 1
        For j = 0 To 4: 出力(行, j + 1) = 組(j): Next j
    Next 組
    With 出
        .Range("A1:E" & 全 + 4).NumberFormat = "@"
        .Range("A1").Value = "調査：ブックの構造　" & wb.Name
        .Range("A2").Value = "セルの値は出していません（式・名前・書式・規則の形だけ）。このシートをコピーして AI に貼れば、中身を出さずに相談できます。"
        .Range("A3").Value = IIf(打切, "多いので先頭の 60,000 行までを出しました。", "") & "気をつける所は「見つけたこと → なぜ → Excel のどこで確かめるか」。オフィス田中のワークシート診断ツールの項目に学んでいます。"
        .Range("A4:E4").Value = Array("区分", "項目", "場所", "内容（値は出さない）", "なぜ・確かめる所・補足")
        .Range("A5").Resize(全, 5).Value = 出力
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range("A4:E4")
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range("A4").Resize(全 + 1, 5)
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
        End With
        .Range("A1:A3").WrapText = False
        .Range("A4:C" & 全 + 4).Columns.AutoFit
        .Columns("D").ColumnWidth = 60
        .Columns("E").ColumnWidth = 60
        .Range("D5:E" & 全 + 4).WrapText = True
        For j = 1 To 3
            If .Columns(j).ColumnWidth > 40 Then .Columns(j).ColumnWidth = 40: .Columns(j).WrapText = True
        Next j
    End With
    出.Activate
    出.Range("A1").Select
    Application.ScreenUpdating = True
    Application.StatusBar = "調査_構造: 気をつける所 " & 点.Count & " 件・全 " & 全 & " 行（値は出していません）"
    Exit Sub

式を数える:
    式数 = 式数 + 1
    If 型数.exists(式R) Then
        型数(式R) = 型数(式R) + 1
        型末(式R) = 見.Address(False, False)
    Else
        型数.Add 式R, 1
        型先.Add 式R, 見.Address(False, False)
        型末.Add 式R, 見.Address(False, False)
        型式.Add 式R, 式
        型種.Add 式R, CreateObject("Scripting.Dictionary")
        GoSub 関数を拾う
        For Each 組 In Split(当, "|")
            If 組 <> "" Then 関辞(組) = 1: 関全(組) = 1
        Next 組
    End If
    If IsError(v) Then
        Select Case CLng(v)
            Case 2000: 種 = "#NULL!"
            Case 2007: 種 = "#DIV/0!"
            Case 2015: 種 = "#VALUE!"
            Case 2023: 種 = "#REF!"
            Case 2029: 種 = "#NAME?"
            Case 2036: 種 = "#NUM!"
            Case 2042: 種 = "#N/A"
            Case 2045: 種 = "#SPILL!"
            Case 2050: 種 = "#CALC!"
            Case Else: 種 = "#エラー"
        End Select
    ElseIf isEmpty(v) Then
        種 = "空"
    ElseIf VarType(v) = vbBoolean Then
        種 = "論理"
    ElseIf VarType(v) = vbDate Then
        種 = "日付"
    ElseIf VarType(v) = vbString Then
        種 = IIf(v = "", "空", "文字")
    Else
        種 = "数"
    End If
    Set 辞 = 型種(式R)
    辞(種) = 辞(種) + 1
    Return

関数を拾う:
    ' 式の中の関数名を拾う（文字列の中は数えない・_xlfn. などを外す）→ 当 = "|SUM|IF|"
    当 = "|": 語 = "": 括 = False
    For p = 1 To Len(式)
        字 = Mid$(式, p, 1)
        If 字 = """" Then
            括 = Not 括: 語 = ""
        ElseIf Not 括 Then
            If 字 Like "[A-Za-z0-9._]" Then
                語 = 語 & 字
            ElseIf 字 = "(" And 語 <> "" Then
                語 = UCase$(Replace(Replace(Replace(語, "_xlfn.", "", , , vbTextCompare), "_xlws.", "", , , vbTextCompare), "_xludf.", "", , , vbTextCompare))
                If 語 Like "[A-Z]*" And InStr(当, "|" & 語 & "|") = 0 Then 当 = 当 & 語 & "|"
                語 = ""
            Else
                語 = ""
            End If
        End If
    Next p
    Return

詳を足す:
    詳.Add Array(区, 項, 場, 容, 法)
    Return

点を足す:
    点.Add Array("気をつける所", 項, "", 容, 法)
    Return

要を足す:
    要.Add Array("概要", 項, "", 容, "")
    Return

失敗:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "ブックの構造: 調べられませんでした（" & Err.Description & "）"
End Sub
