Attribute VB_Name = "表の整理_調べる"
' 表の整理_調べる - 棚（言葉で頼んで撃つマクロ）の調べる（調べる・一覧・点検）（2026-09-23 表の整理から分けた）

Sub 選んだ列の重複する値を一覧にする()
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
        If dict.Exists(normKey) Then
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

        If dict.Exists(key) Then
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
            If yrCounts.Exists(yr) Then
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

        If Not dictVariants.Exists(nKey) Then
            Set dictVariants(nKey) = CreateObject("Scripting.Dictionary")
            dictFirst(nKey) = rawVal
            ReDim Preserve orderList(orderCount)
            orderList(orderCount) = nKey
            orderCount = orderCount + 1
        End If

        If Not dictVariants(nKey).Exists(rawVal) Then
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
    If 正規 Is Nothing Then
        Set 正規 = CreateObject("VBScript.RegExp")
        正規.Global = True
    End If
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 5
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 4
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 6
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
            If 鍵.Exists(正) Then
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 5
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
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, nC)).Sort Key1:=out.Cells(hr + 1, 3), Order1:=2, Header:=2
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "このシートを使っているマクロは見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 4
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 12
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
    Dim tb As Range, v As Variant, nR As Long, nc2 As Long, x As Variant, s As String, 注 As String
    Dim d As Object, n数 As Long, n日 As Long, n文 As Long, n他 As Long, n空 As Long, n誤 As Long, n文数 As Long
    Dim mn As Variant, mx As Variant, 例 As String, 例数 As Long, 型 As String, n型数 As Long, 部 As String
    On Error GoTo 失敗
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    If TypeName(ws) <> "Worksheet" Then Exit Sub
    Set tb = ActiveCell.CurrentRegion
    If Application.WorksheetFunction.CountA(tb) = 0 Then Set tb = ws.UsedRange
    nR = tb.rows.Count: nc2 = tb.Columns.Count
    If nR > 100000 Then
        nR = 100000
        Set tb = tb.Resize(nR, nc2)
        注 = "（先頭 100,000 行まで）"
    End If
    If nR * nc2 = 1 Then
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
    out.Range("A2").Value = "範囲 " & tb.Address(False, False) & "　見出し行 " & tb.Row & "　データ " & Format(nR - 1, "#,##0") & " 行　" & nc2 & " 列 " & 注
    out.Range("A3").Value = "見出し行は範囲の先頭行とみなしています。列の意味は決められないので、見出しと例を並べています。"
    hr = 5: nC = 11
    out.Range("A5:K5").Value = Array("列", "見出し", "入力あり", "空欄", "型", "種類の数", "最小", "最大", "文字の数字", "エラー", "例")
    r = hr
    For j = 1 To nc2
        n数 = 0: n日 = 0: n文 = 0: n他 = 0: n空 = 0: n誤 = 0: n文数 = 0
        mn = Empty: mx = Empty: 例 = "": 例数 = 0
        Set d = CreateObject("Scripting.Dictionary")
        For i = 2 To nR
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
                If Not d.Exists(s) Then
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Application.StatusBar = "調査_表の作り: " & nc2 & " 列・データ " & Format(nR - 1, "#,##0") & " 行を一覧にしました。"
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 5
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 9
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 4
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 5
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
        Set f = sh.Cells.Find("#REF!", LookIn:=-4123, LookAt:=2)
        On Error GoTo 失敗
        If Not f Is Nothing Then
            first = f.Address
            n = 0
            Do
                n = n + 1
                If Not seen.Exists(sh.Name & "!" & f.Address) Then
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
    For Each nn In wb.names
        If InStr(nn.RefersTo, "#REF!") > 0 Then
            r = r + 1
            out.Cells(r, 1).Value = "壊れた名前"
            out.Cells(r, 3).Value = "'" & nn.Name
            out.Cells(r, 4).Value = "'" & nn.RefersTo
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 5
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
                                        If cnt.Exists(key) Then
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
                                    If cnt2.Exists(key) Then
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
        Next i
    End If
    If 打切 Then out.Range("A2").Value = "数式が多いので、先頭の 200,000 個までを調べました。"
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "他のシート・他のブックへの参照は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 4
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
        Set lc = sh.Cells.Find("*", SearchOrder:=1, SearchDirection:=2, LookIn:=-4123)
        If Not lc Is Nothing Then lastR = lc.Row
        Set lc = Nothing
        Set lc = sh.Cells.Find("*", SearchOrder:=2, SearchDirection:=2, LookIn:=-4123)
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
        For Each shp In sh.Shapes
            If shp.Type = 13 Then np = np + 1
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
    r = r + 1
    out.Cells(r, 1).Value = "名前定義の数"
    out.Cells(r, 2).Value = wb.names.Count & " 個"
    If wb.names.Count > 1000 Then out.Cells(r, 3).Value = "要確認": out.Cells(r, 4).Value = "名前が多い（使っていない名前が溜まっていないか）"
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
            Set f = sh.Cells.Find(w, LookIn:=-4123, LookAt:=2)
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
        If out.Cells(i, 3).Value = "要確認" Then out.Range(out.Cells(i, 1), out.Cells(i, nC)).Interior.Color = RGB(255, 242, 204)
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
                            If Not dic(n).Exists(s) Then dic(n).Add s, 1
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
    hr = 4: nC = 8
    out.Range("A4:H4").Value = Array("シートA", "シートB", "そろっている見出し", "類似度", "Aだけの見出し", "Bだけの見出し", "Aのデータ行", "Bのデータ行")
    r = hr
    For a = 1 To n - 1
        For b = a + 1 To n
            common = 0
            For Each k In dic(a).keys
                If dic(b).Exists(k) Then common = common + 1
            Next k
            If dic(a).Count + dic(b).Count - common > 0 Then sim = common / (dic(a).Count + dic(b).Count - common) Else sim = 0
            If common >= 3 And sim >= 0.6 Then
                差A = "": 差B = "": cA = 0: cB = 0
                For Each k In dic(a).keys
                    If Not dic(b).Exists(k) Then
                        cA = cA + 1
                        If cA <= 6 Then 差A = 差A & IIf(Len(差A) > 0, "、", "") & k
                    End If
                Next k
                For Each k In dic(b).keys
                    If Not dic(a).Exists(k) Then
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
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, nC)).Sort Key1:=out.Cells(hr + 1, 4), Order1:=2, Header:=2
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "見出しが似ているシートの組は見つかりませんでした"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
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
    hr = 4: nC = 6
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, hr As Long, nC As Long
    Dim procs As Object, okVba As Boolean, n As Long, jj As Long, proc As String, nxt As Long
    Dim act As String, mac As String, p As Long, t As String, 種 As String, 確 As String, 他 As Boolean, loc As String
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
    hr = 4: nC = 7
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
                mac = act
                p = InStrRev(mac, "!")
                If p > 0 Then mac = Mid$(mac, p + 1)
                p = InStrRev(mac, ".")
                If p > 0 Then mac = Mid$(mac, p + 1)
                mac = Replace(mac, "'", "")
                他 = (InStr(act, "!") > 0 And InStr(LCase$(act), LCase$(wb.Name)) = 0)
                If Not okVba Then
                    確 = "確認できません（VBA に触れない設定）"
                ElseIf 他 Then
                    確 = "他のブックのマクロ（未確認）"
                ElseIf procs.Exists(LCase$(mac)) Then
                    確 = "あり"
                Else
                    確 = "見つからない"
                    out.Range(out.Cells(r, 1), out.Cells(r, nC)).Interior.Color = RGB(255, 199, 206)
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
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
    Dim nM As String, nm0 As String, kk As Long, r As Long, rt As Long, j As Long, i As Long, hr As Long, nC As Long
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
    hr = 4: nC = 6
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
                    If cnt.Exists(key) Then
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
    Set re = CreateObject("VBScript.RegExp")
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
            If Not fn.Exists(tk) Then
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
            If tk <> "REF" And Not sm.Exists(tk) Then
                sm.Add tk, 1
                s = s & IIf(Len(s) > 0, "、", "") & tk
            End If
        Next m
        out.Cells(r, 6).Value = "'" & s
    Next kv
    If r > hr + 1 Then out.Range(out.Cells(hr + 1, 1), out.Cells(r, nC)).Sort Key1:=out.Cells(hr + 1, 2), Order1:=2, Header:=2
    If 打切 Then out.Range("A2").Value = "数式が多いので、先頭の 100,000 個までを調べました。"
    If r = hr Then
        r = r + 1
        out.Cells(r, 1).Value = "このシートに数式はありません"
    End If
    rt = r
    With out
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12
        With .Range(.Cells(hr, 1), .Cells(hr, nC))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        With .Range(.Cells(hr, 1), .Cells(rt, nC))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4160
            .Columns.AutoFit
        End With
        For j = 1 To nC
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
