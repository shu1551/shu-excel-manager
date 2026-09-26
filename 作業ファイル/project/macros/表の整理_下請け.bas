Attribute VB_Name = "表の整理_下請け"
Option Private Module
' 棚の共通の下請け（2026-09-23 表の整理から分けた）。Option Private Module＝このブックの中からだけ呼べる・
' マクロの一覧・Alt+F8 には出ない。受け渡しの変数と下請けは、棚の各モジュールから呼ぶので公開にしている。

' 仕上げ … 右や下に新しく作った表の見た目は、共通の下請け「表の仕上げ」で付ける（2026-09-23 shu の指示）。
'          呼ぶ側は m仕上げ範囲 に範囲（1 行目が見出し）を入れて Call 表の仕上げ。
' 足した列の仕上げは元の表の見出しをまねる。m仕上げ手本 に元の見出しのセルを入れてから呼ぶ（青い塗りは付けない）
Public m仕上げ手本 As Range
Public m仕上げ範囲 As Range

' 表の本文の最後の行（2026-09-23）。
' 使用範囲の最終行を表の終わりにすると、表の下のメモや計算の行（「1個あたり平均単価 →」と D31 の値）まで
' 明細に数えた（項目に数が並び、SUMIF が D6:D31 になり、エラー値の行では型不一致で止まった）。
' 行ごとに値の入ったセルを数え、いちばん多い行の数を表の幅とする。幅の 3 分の 1 以上（2 つ以上）ある最初の行
' （見出し）から下へたどり、本文に続いている行は値が少なくても本文に数える（入力が途中の行で切らない）。
' 合計・小計・平均などの行は本文に数えずに越える。空行か合計の行を越えた後は、幅の半分以上ある行だけ本文に戻し
' （表の中の空行・小計の後の明細は越える）、それより少ない行（メモの行）か、空行が 2 つ目になったら止める。
' 本文が見つからなければ使用範囲の最終行を返す（前と同じ）。
Function 表の本文の最終行(ws As Object) As Long
    Dim ur As Object, v As Variant, i As Long, j As Long, n As Long, w As Long
    Dim need As Long, full As Long, lastData As Long, blanks As Long, started As Boolean, after As Boolean
    Dim cnt() As Long, sumRow() As Boolean, s As String
    Set ur = ws.UsedRange
    表の本文の最終行 = ur.Row + ur.rows.Count - 1
    v = ur.Value
    If Not IsArray(v) Then Exit Function
    ReDim cnt(1 To UBound(v, 1))
    ReDim sumRow(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        n = 0
        For j = 1 To UBound(v, 2)
            If VarType(v(i, j)) = vbString Then
                s = Trim$(v(i, j))
                If Len(s) > 0 Then
                    n = n + 1
                    If InStr(s, "合計") > 0 Or InStr(s, "小計") > 0 Or InStr(s, "総計") > 0 Or s = "計" _
                       Or s = "平均" Or s = "最大" Or s = "最小" Or s = "件数" Then sumRow(i) = True
                End If
            ElseIf Not isEmpty(v(i, j)) Then
                n = n + 1
            End If
        Next j
        cnt(i) = n
        If n > w Then w = n
    Next i
    If w < 2 Then Exit Function
    need = (w + 2) \ 3
    If need < 2 Then need = 2
    full = (w + 1) \ 2
    If full < need Then full = need
    For i = 1 To UBound(v, 1)
        If Not started Then
            If cnt(i) >= need And Not sumRow(i) Then
                started = True
                lastData = i
            End If
        ElseIf cnt(i) = 0 Then
            after = True
            blanks = blanks + 1
            If blanks >= 2 Then Exit For
        ElseIf sumRow(i) Then
            after = True                 ' 合計・平均などの行は本文に数えずに越える
        ElseIf Not after Or cnt(i) >= full Then
            lastData = i
            after = False
            blanks = 0
        Else
            Exit For                     ' 空行・合計の行の後の、値の少ない行＝表の下のメモ
        End If
    Next i
    If lastData > 0 Then 表の本文の最終行 = ur.Row + lastData - 1
End Function

' 表の本文のすぐ下（空行 1 つまで）にある合計の行（無ければ 0）。表の下に合計行を足す が、表の下に既にある合計の行を
' 見つけられずに、メモの行のさらに下へ 2 本目を足していた（2026-09-23）。
Function 表の下の合計行(ws As Object, lastRow As Long) As Long
    Dim ur As Object, r As Long, c As Long, n As Long, s As String, bottom As Long
    Set ur = ws.UsedRange
    bottom = ur.Row + ur.rows.Count - 1
    For r = lastRow + 1 To lastRow + 2
        If r > bottom Then Exit Function
        n = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            If VarType(ws.Cells(r, c).Value) = vbString Then
                s = Replace(Trim$(ws.Cells(r, c).Value), " ", "")
                If s = "合計" Or s = "総計" Or s = "計" Then
                    表の下の合計行 = r
                    Exit Function
                End If
                If Len(s) > 0 Then n = n + 1
            ElseIf Not isEmpty(ws.Cells(r, c).Value) Then
                n = n + 1
            End If
        Next c
        If n > 0 Then Exit Function      ' 合計でない行が挟まったら探さない
    Next r
End Function

Sub 表の仕上げ()
    ' 棚の共通の仕上げ（2026-09-23）。m仕上げ範囲（1 行目が見出し）の見た目を付ける:
    '   見出し＝水色の塗り・太字・中央／番号とコードの列＝中央／日付＝左で yyyy/m/d／
    '   数＝右で #,##0（率・年・% の列は触らない）／そのほかの文字＝左／細い格子の罫線・上下中央・列幅 6～60
    Dim r As Range, c As Long, rr As Long, hname As String
    Dim nNum As Long, nDate As Long, nAll As Long, v As Variant
    If m仕上げ範囲 Is Nothing Then Exit Sub
    Set r = m仕上げ範囲
    If r.rows.Count < 1 Then Set m仕上げ範囲 = Nothing: Exit Sub
    With r
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
    End With
    If m仕上げ手本 Is Nothing Then
        With r.rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(221, 235, 247)
        End With
    Else
        ' 元の表に足した列は、その表の見出しの書式をまねる（別の色にしない。2026-09-23）。
        ' クリップボード（Copy/PasteSpecial）は使わない＝1004 で止まる・人の切り取りを消すため
        With r.rows(1)
            .Font.Name = m仕上げ手本.Font.Name
            .Font.Size = m仕上げ手本.Font.Size
            .Font.Bold = m仕上げ手本.Font.Bold
            .Font.Color = m仕上げ手本.Font.Color
            .HorizontalAlignment = m仕上げ手本.HorizontalAlignment
            If m仕上げ手本.Interior.ColorIndex = xlNone Then
                .Interior.ColorIndex = xlNone
            Else
                .Interior.Color = m仕上げ手本.Interior.Color
            End If
        End With
    End If
    For c = r.Column To r.Column + r.Columns.Count - 1
        hname = CStr(r.Worksheet.Cells(r.Row, c).Value)
        nNum = 0: nDate = 0: nAll = 0
        For rr = r.Row + 1 To r.Row + r.rows.Count - 1
            v = r.Worksheet.Cells(rr, c).Value
            If Not isEmpty(v) And Not IsError(v) Then
                nAll = nAll + 1
                If VarType(v) = vbDate Then
                    nDate = nDate + 1
                ElseIf VarType(v) = vbDouble Or VarType(v) = vbCurrency Then
                    nNum = nNum + 1
                End If
            End If
        Next rr
        With r.Worksheet.Range(r.Worksheet.Cells(r.Row + 1, c), r.Worksheet.Cells(r.Row + r.rows.Count - 1, c))
            If InStr(hname, "番号") > 0 Or InStr(hname, "No") > 0 Or InStr(hname, "ＮＯ") > 0 Or hname = "ID" Or InStr(hname, "コード") > 0 Then
                .HorizontalAlignment = xlCenter
            ElseIf nAll > 0 And nDate >= nAll * 0.6 Then
                .HorizontalAlignment = xlLeft
                .NumberFormat = "yyyy/m/d"
            ElseIf nAll > 0 And nNum >= nAll * 0.6 Then
                .HorizontalAlignment = xlRight
                If InStr(hname, "率") = 0 And InStr(hname, "年") = 0 And InStr(hname, "%") = 0 And InStr(hname, "順位") = 0 Then .NumberFormat = "#,##0"
            Else
                .HorizontalAlignment = xlLeft
            End If
        End With
    Next c
    r.Columns.AutoFit
    For c = r.Column To r.Column + r.Columns.Count - 1
        If r.Worksheet.Columns(c).ColumnWidth > 60 Then r.Worksheet.Columns(c).ColumnWidth = 60
        If r.Worksheet.Columns(c).ColumnWidth < 6 Then r.Worksheet.Columns(c).ColumnWidth = 6
    Next c
    Set m仕上げ範囲 = Nothing
    Set m仕上げ手本 = Nothing
End Sub

Function セルの字(ByVal v As Variant) As String
    ' セルの値を文字にする（エラー値は "#エラー"）。If v <> "" を エラー値の入った表でも落ちずに比べるための下請け（2026-09-23）
    ' セル（Range）が渡ってきたら値を読む（If c <> "" は c の値と比べる＝同じ意味にする）
    Dim x As Variant
    If IsObject(v) Then
        If v Is Nothing Then Exit Function
        If TypeName(v) <> "Range" Then
            On Error Resume Next
            セルの字 = CStr(v)
            Exit Function
        End If
        x = v.Cells(1, 1).Value
    Else
        x = v
    End If
    If IsError(x) Then
        セルの字 = "#エラー"
    ElseIf IsNull(x) Or IsArray(x) Then
        セルの字 = ""
    Else
        セルの字 = CStr(x)
    End If
End Function

' 表の右に出す一覧・表の置き場所（2026-09-23）。既定列＝表の右端の 1 列空けた先。
'   ・その行の右（探し始め?使用範囲の右端）に自分の前の出力（見出しが 目印＝「|」区切り・* は空でなければ何でも）があればそこ
'     ＝撃ち直すと同じ場所に書き直す
'   ・既定列から 幅 列が使用範囲の中で空ならそこ
'   ・ほかの物（別のマクロが右に作った表・人のメモ）があれば、使用範囲の右端の 1 列空けた先＝人の物を消さない
'   （表のおかしい所を右に報告する が右の件数表を消して報告を書いた。同じ形が 9 本あった）
Function 右の出し先(ws As Object, ByVal hdrRow As Long, ByVal 既定列 As Long, ByVal 幅 As Long, ByVal 目印 As String, Optional ByVal 探し始め As Long = 0) As Long
    Dim ur As Object, urRight As Long, c As Long, k As Long, 印 As Variant, 合う As Boolean
    Set ur = ws.UsedRange
    urRight = ur.Column + ur.Columns.Count - 1
    印 = Split(目印, "|")
    If 探し始め <= 0 Then 探し始め = 既定列
    For c = 探し始め To urRight
        合う = True
        For k = 0 To UBound(印)
            If 印(k) = "*" Then
                If セルの字(ws.Cells(hdrRow, c + k).Value) = "" Then 合う = False
            ElseIf セルの字(ws.Cells(hdrRow, c + k).Value) <> 印(k) Then
                合う = False
            End If
            If Not 合う Then Exit For
        Next k
        If 合う Then
            右の出し先 = c
            Exit Function
        End If
    Next c
    If 既定列 > urRight Then
        右の出し先 = 既定列
    ElseIf Application.WorksheetFunction.CountA(ws.Range(ws.Cells(ur.Row, 既定列), ws.Cells(ur.Row + ur.rows.Count - 1, 既定列 + 幅 - 1))) = 0 Then
        右の出し先 = 既定列
    Else
        右の出し先 = urRight + 2
    End If
End Function

' 集計の行の目印の語か（合計・小計・総計・総合計・計・平均。「総務課 小計」「合計（税込）」のように前に語・後ろに括弧書きが付いたものも）。
' 棚の判定が「小計」とぴったり同じ字だけを見ていて、「総務課 小計」の行に番号を振り、太字と塗りを消していた（2026-09-24 通しの実測 4）。
' 「合計請求書」のように語で始まる文は拾わない（末尾で見る）。
Function 集計の語か(ByVal v As Variant) As Boolean
    Dim s As String, k As Long, w As Variant
    s = Replace(Replace(Trim$(セルの字(v)), " ", ""), "　", "")
    k = InStr(s, "（")
    If k = 0 Then k = InStr(s, "(")
    If k > 1 Then s = Left$(s, k - 1)
    If Len(s) = 0 Or Len(s) > 12 Then Exit Function
    If s = "計" Or s = "平均" Or s = "合計" Then 集計の語か = True: Exit Function
    For Each w In Array("小計", "合計", "総計")
        If Right$(s, Len(w)) = w Then 集計の語か = True: Exit Function
    Next w
End Function

' 見出しのセルの字。結合セル（二段見出しの縦の結合・「実績」の横の結合）は結合の左上の字を返す（2026-09-24 棚の試験 F10:
' 結合の右側・下側を空と見て表の右端を取り違え、備考の列を一覧の書き出し先にして消していた）。
Function 見出しの字(ByVal c As Object) As String
    If c.MergeCells Then
        見出しの字 = Trim$(セルの字(c.MergeArea.Cells(1, 1).Value))
    Else
        見出しの字 = Trim$(セルの字(c.Value))
    End If
End Function

' 二段見出しなら下の段の行、そうでなければ hdrRow のまま（2026-09-24 棚の試験 F10・通しの実測 5）。
' 二段＝下の段に、上の段と縦に結合した列（コード・品目）があり、下の段の残りが 2 つ以上の文字（4月・5月／数量・金額）で数・日付が無い。
Function 見出しの下の段(ByVal ws As Object, ByVal hdrRow As Long, ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long, vMerge As Boolean, txt As Long, v As Variant
    見出しの下の段 = hdrRow
    For c = c1 To c2
        If ws.Cells(hdrRow + 1, c).MergeCells Then
            If ws.Cells(hdrRow + 1, c).MergeArea.Row = hdrRow Then vMerge = True
        Else
            v = ws.Cells(hdrRow + 1, c).Value
            If セルの字(v) <> "" Then
                If IsNumeric(v) Or IsDate(v) Then Exit Function
                txt = txt + 1
            End If
        End If
    Next c
    If vMerge And txt >= 2 Then 見出しの下の段 = hdrRow + 1
End Function

' ピボットの行・列の項目から、元の表の集計の行（「総務課 小計」「合計」）を外す（2026-09-24 棚の試験 F9: 明細のあいだの小計の行も
' 項目に数えて総計が 2 倍になっていた）。集計の語の項目を隠し、明細の行でそのキーの列が空のものが 1 つも無ければ「(空白)」も隠す
' （小計の字が別の列にあってキーの列が空の形＝実測4）。隠した項目は総計にも入らない。
Sub ピボットの集計の行を外す(ByVal pt As Object, ByVal src As Object)
    Dim pf As Object, pi As Object, coll As Variant, c As Long, r As Long, k As Long
    Dim isSum() As Boolean, blankDetail As Boolean, hdr As String
    ReDim isSum(2 To src.rows.Count)
    For r = 2 To src.rows.Count
        For c = 1 To src.Columns.Count
            If VarType(src.Cells(r, c).Value) = vbString Then
                If 集計の語か(src.Cells(r, c).Value) Then isSum(r) = True: Exit For
            End If
        Next c
    Next r
    On Error Resume Next
    For Each coll In Array(pt.RowFields, pt.ColumnFields)
        For Each pf In coll
            k = 0
            For c = 1 To src.Columns.Count
                If 見出しの字(src.Cells(1, c)) = pf.SourceName Then k = c: Exit For
            Next c
            blankDetail = False
            If k > 0 Then
                For r = 2 To src.rows.Count
                    If Not isSum(r) And セルの字(src.Cells(r, k).Value) = "" Then
                        If Application.WorksheetFunction.CountA(src.rows(r)) > 0 Then blankDetail = True: Exit For
                    End If
                Next r
            End If
            For Each pi In pf.PivotItems
                If 集計の語か(pi.Name) Then
                    pi.Visible = False
                ElseIf (pi.Name = "(空白)" Or pi.Name = "(blank)") And Not blankDetail Then
                    pi.Visible = False
                End If
            Next pi
        Next pf
    Next coll
    On Error GoTo 0
End Sub

' 表の行 r が集計の行（「合計」「平均」「総務課 小計」など、c1?c2 のどこかに集計の語）か（2026-09-24 棚の試験 F9:
' グラフのマクロが項目の列の「合計」だけを見て「企画課 小計」の行をグラフの項目に入れていた）。
Function 集計の行か(ByVal ws As Object, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long
    For c = c1 To c2
        If VarType(ws.Cells(r, c).Value) = vbString Then
            If 集計の語か(ws.Cells(r, c).Value) Then 集計の行か = True: Exit Function
        End If
    Next c
End Function

' テーブルにする明細の表の範囲（見出しの行?最後の明細の行）。題の行・下の合計・平均・メモは入れない（2026-09-24 棚の試験:
' 使用範囲まるごとをテーブルにして、題「令和8年度 支出一覧」を見出しに・本当の見出しを明細の 1 行目にしていた）。
' 二段見出しの表と、明細のあいだに小計の行がある表は、テーブルにすると見出しが崩れる・並べ替えで小計が混ざるので
' Nothing を返し、why に理由を入れる（呼び手はステータスバーで言って終わる）。
Function 明細の表の範囲(ByVal ws As Object, ByRef why As String) As Object
    Dim ur As Object, r As Long, c As Long, hdr As Long, c1 As Long, c2 As Long, lastR As Long
    Dim filled As Long, txt As Long, v As Variant
    why = ""
    Set ur = ws.UsedRange
    For r = ur.Row To ur.Row + ur.rows.Count - 1
        filled = 0: txt = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            v = ws.Cells(r, c).Value
            If セルの字(v) <> "" Then
                filled = filled + 1
                If VarType(v) = vbString And Not IsNumeric(v) Then txt = txt + 1
            End If
        Next c
        If filled >= 2 And txt * 2 > filled Then hdr = r: Exit For
    Next r
    If hdr = 0 Then why = "見出しの行が見つかりません": Exit Function
    If 見出しの下の段(ws, hdr, ur.Column, ur.Column + ur.Columns.Count - 1) <> hdr Then
        why = "二段見出しの表はテーブルにできません（見出しが 1 行でないと崩れる）。先に「二段の見出しを一行に畳む」を撃ってください"
        Exit Function
    End If
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        If 見出しの字(ws.Cells(hdr, c)) <> "" Then
            If c1 = 0 Then c1 = c
            c2 = c
        ElseIf c1 > 0 Then
            Exit For
        End If
    Next c
    lastR = 表の本文の最終行(ws)
    If lastR <= hdr Then why = "明細の行が見つかりません": Exit Function
    For r = hdr + 1 To lastR
        If 集計の行か(ws, r, c1, c2) Then
            why = "明細のあいだに集計の行（" & r & " 行目）がある表はテーブルにできません（並べ替え・絞り込みで明細と混ざる）"
            Exit Function
        End If
    Next r
    Set 明細の表の範囲 = ws.Range(ws.Cells(hdr, c1), ws.Cells(lastR, c2))
End Function

' 数に見えるキーを比べやすい字にする（2026-09-24 総点検: CLng で 21 億を超える番号・金額が「オーバーフロー」で止まり、
' 1.5 と 2 を同じキーにしていた）。数字だけの字は桁をそのまま（先頭の 0 は落とす＝0012 と 12 は同じ）、
' 小数・符号つきは数に直した字。数でなければ前後の空白を取った字のまま
Function 数のキー(ByVal s As String) As String
    Dim d As Double
    s = Trim$(s)
    数のキー = s
    If Len(s) = 0 Then Exit Function
    If Not s Like "*[!0-9]*" Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
        数のキー = s
        Exit Function
    End If
    If s Like "*[!0-9.,+-]*" Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    On Error Resume Next
    d = CDbl(s)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    If d = Int(d) And Abs(d) < 1E+15 Then
        数のキー = Format$(d, "0")
    Else
        数のキー = CStr(d)
    End If
End Function

' 文字が日付として読めるか（2026-09-24 総点検: IsDate は「1-2」「3/4」のような枝番・分数も日付にし、貼ったCSVを使える形に整える が
' 区画「1-2」を 1 月 2 日に化かしていた）。年・月・日のそろう形（2026/4/1・2026-04-01・4/1/2026・令和8年4月1日・Mar 1, 2026）だけ。
' 郵便番号（010-0001）・電話・先頭 0 のコードは日付でない
Function 日付の字か(ByVal t As String) As Boolean
    Dim n As String, sep As Variant
    t = Trim$(t)
    If Len(t) < 6 Then Exit Function
    If IsNumeric(t) Then Exit Function
    If Not IsDate(t) Then Exit Function
    n = StrConv(t, vbNarrow)
    If Not n Like "*[!0-9-]*" Then
        If Left$(n, 1) = "0" Or n Like "###-####" Then Exit Function
    End If
    If InStr(n, "年") > 0 Or n Like "*[A-Za-z][A-Za-z][A-Za-z]*" Then 日付の字か = True: Exit Function
    For Each sep In Array("/", "-", ".")
        If Len(n) - Len(Replace(n, sep, "")) = 2 Then 日付の字か = True: Exit Function
    Next sep
End Function

' 棚が作ったグラフの印＝名前の頭「棚_」。作り直すときは棚の印のグラフだけ消し、人が作ったグラフは残す（2026-09-24 総点検:
' グラフを作る 7 本が、シートのグラフを全部消してから作っていた＝人が作ったグラフまで消えた）
Sub 棚のグラフを消す(ByVal ws As Object)
    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        If Left$(ws.ChartObjects(i).Name, 2) = "棚_" Then ws.ChartObjects(i).Delete
    Next i
End Sub

' 棚が作ったシートの印（シートの中の名前「棚_作成」・見えない）。撃ち直しで消してよいのは印のあるシートだけ（2026-09-24 総点検:
' 選んだ列の値ごとにシートを分ける・一覧を1件1枚でひな形に差し込む が、値と同じ名前の人のシート（「総務課」など）を消していた）
Sub 棚のシートの印(ByVal sh As Object)
    On Error Resume Next
    sh.names.Add Name:="棚_作成", RefersTo:="=TRUE", Visible:=False
End Sub

Function 棚のシートか(ByVal sh As Object) As Boolean
    Dim nM As Object
    On Error Resume Next
    For Each nM In sh.names
        If nM.Name Like "*!棚_作成" Or nM.Name = "棚_作成" Then 棚のシートか = True: Exit Function
    Next nM
End Function

' シート名に使えない字（: \ / ? * [ ]）を _ に、31 字まで、空なら「無題」（2026-09-24 総点検: 日付・「A/B」の値で
' 分けると名前を付けられずに実行時エラーで止まり、名前の無いシートが残った）
Function シート名にできる字(ByVal s As String) As String
    Dim ch As Variant
    s = Trim$(Replace(Replace(s, vbCr, " "), vbLf, " "))
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        s = Replace(s, ch, "_")
    Next ch
    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "無題"
    If s = "履歴" Then s = "履歴_"
    シート名にできる字 = s
End Function

' 棚が作るシートの名前。同じ名前が棚の印のシート（今回の撃ちで作ったものを除く）ならそれを消して同じ名前で作り直し、
' 人のシートなら (2)(3)… を付けて別に作る。今回＝この撃ちで作った名前の辞書（同じ名前にならした 2 つの値で上書きしない）
Function 棚のシート名(ByVal wb As Object, ByVal 名 As String, Optional ByVal 今回 As Object = Nothing) As String
    Dim base As String, cand As String, k As Long, sh As Object, 消す As Boolean
    base = シート名にできる字(名)
    cand = base: k = 1
    Do
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Sheets(cand)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        消す = 棚のシートか(sh)
        If 消す And Not 今回 Is Nothing Then 消す = Not 今回.exists(LCase$(cand))
        If 消す Then
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = True
            Exit Do
        End If
        k = k + 1
        cand = Left$(base, 31 - Len("(" & k & ")")) & "(" & k & ")"
    Loop
    If Not 今回 Is Nothing Then 今回(LCase$(cand)) = True
    棚のシート名 = cand
End Function

' 右の出し先 が前の出力（見出しの目印が合う所）を返したとき、その出力のかたまりだけを消す（2026-09-24 総点検: 表の右から
' 使用範囲の右端までを全部消すマクロが 6 本あり、右に置いた人のメモが消えた）。見出しの行から、空行が 2 つ続く手前まで
Sub 前の出力を消す(ByVal ws As Object, ByVal hdrRow As Long, ByVal c1 As Long, ByVal 幅 As Long)
    Dim c2 As Long, lastR As Long, r As Long, blank As Long, bottom As Long
    c2 = c1 + 幅 - 1
    bottom = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
    lastR = hdrRow
    r = hdrRow + 1
    Do While blank < 2 And r <= bottom
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))) > 0 Then
            lastR = r: blank = 0
        Else
            blank = blank + 1
        End If
        r = r + 1
    Loop
    ws.Range(ws.Cells(hdrRow, c1), ws.Cells(lastR, c2)).Clear
End Sub
