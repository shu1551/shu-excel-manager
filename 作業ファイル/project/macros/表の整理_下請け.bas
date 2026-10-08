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

' 頼みの文と見出しの行から、ピボットの 行・列・値 の列を決める（2026-10-09 ピボットを頼みの文だけで）。
'   見出しそのものが頼みにあればそれを先に当てて頼みから伏せ、残りで言い換え（部署／部門／課、金額／売上／売上高 など）を当てる
'   （「実績金額」と書いた中の「金額」で「予算金額」まで当てない・段 2 の P03）。値＝数の列。
'   行と列＝頼みに出てくる順。「行は部門」「列に区分」「区分を列」と言われた見出しはそこへ（何度目に出てきた語でも・段 2 の Q1-4）。
'   「年度別」のように 別・ごと・単位 が続く数の列は行か列にする。値を言っていなければ、合計できる数の列が 1 本だけのときそれを使う。
'   決まった＝True（列が無いときは 列 = 0）。決まらない＝False（候補 に見出しの並びと理由。勝手な列で作らない）
Function 頼みからピボットの列(ByVal ws As Worksheet, ByVal 見出し行 As Long, ByVal 左 As Long, ByVal 右 As Long, _
        ByVal 最終行 As Long, ByVal 文 As String, ByRef 行の列 As Long, ByRef 列の列 As Long, ByRef 値の列 As Long, _
        ByRef 候補 As String, Optional ByVal 数えるだけ As Boolean = False) As Boolean
    Dim 組 As Variant, 語々 As Variant, 頼 As String, 伏 As String, 見 As String, 元 As String, 一覧 As String
    Dim c As Long, r As Long, g As Long, k As Long, j As Long, 位置 As Long, 最初 As Long, 語 As String, 当て語 As String
    Dim 見々() As String, 当て々() As String, 位置々() As Long, 数か() As Boolean, 区切りか() As Boolean, 済み() As Boolean
    Dim 札々 As String, 値々 As String, 数 As Long, 字 As Long, v As Variant, 後 As String, 長 As Long, 選 As Long
    Dim 札(1 To 3) As Long, 札数 As Long, 値候補 As Long, 値候補数 As Long, 入替 As Long, 置1 As Long, 置2 As Long
    頼みからピボットの列 = False
    行の列 = 0: 列の列 = 0: 値の列 = 0: 候補 = ""
    頼 = ピボットの語ならし(文)
    If Len(頼) = 0 Or 右 < 左 Then Exit Function
    ' 言い換えの組（見出しの末尾がその組の語なら、組のどの語で頼まれても当てる。1 字の語は後ろに 別・ごと 等が続くときだけ）
    組 = Array("部署|部門|課|所属|部局|部課", "金額|売上|売上高|売り上げ|額|支出額|請求額|支払額|販売額", _
              "区分|種別|分類|カテゴリ|カテゴリー|種類", "担当|担当者|営業担当|社員", _
              "商品|品名|品目|製品", "取引先|得意先|顧客", "地域|エリア|地区|支店|店舗|拠点", "数量|個数|台数", _
              "費目|科目|勘定科目")
    ReDim 見々(左 To 右): ReDim 当て々(左 To 右): ReDim 位置々(左 To 右)
    ReDim 数か(左 To 右): ReDim 区切りか(左 To 右): ReDim 済み(左 To 右)
    For c = 左 To 右
        元 = 見出しの字(ws.Cells(見出し行, c))
        見々(c) = ピボットの語ならし(元)
        If Len(見々(c)) > 0 Then 一覧 = 一覧 & IIf(Len(一覧) > 0, "・", "") & 元
        ' 型: 本文の 8 割以上が数（日付は除く）なら数の列
        数 = 0: 字 = 0
        For r = 見出し行 + 1 To 最終行
            v = ws.Cells(r, c).Value
            If Not IsError(v) And Not isEmpty(v) Then
                If Len(Trim$(CStr(v))) > 0 Then
                    字 = 字 + 1
                    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then 数 = 数 + 1
                End If
            End If
        Next r
        数か(c) = (字 > 0 And 数 >= 字 * 0.8)
    Next c
    ' 1 回目: 見出しそのもの（長い見出しから）。当たった語は頼みから伏せる（同じ字数の〓に。位置は変わらない）
    伏 = 頼
    Do
        選 = 0: 長 = 0
        For c = 左 To 右
            If Not 済み(c) And Len(見々(c)) > 長 Then 長 = Len(見々(c)): 選 = c
        Next c
        If 選 = 0 Then Exit Do
        済み(選) = True
        位置 = InStr(1, 伏, 見々(選), vbBinaryCompare)
        If 位置 > 0 Then
            位置々(選) = 位置: 当て々(選) = 見々(選)
            伏 = Replace(伏, 見々(選), String$(Len(見々(選)), "〓"))
        End If
    Loop
    ' 2 回目: 見出しが頼みに無い列だけ、言い換えの組で（伏せた頼みで探す）
    For c = 左 To 右
        If 位置々(c) = 0 And Len(見々(c)) > 0 Then
            最初 = 0: 当て語 = ""
            ' 組に入るかは見出しの末尾で見る（「売上日」は日付＝金額の組に入れない。「部署名」「担当者名」の 名 は外して見る）
            元 = 見々(c)
            If Len(元) > 2 And Right$(元, 1) = "名" Then 元 = Left$(元, Len(元) - 1)
            For g = 0 To UBound(組)
                語々 = Split(組(g), "|")
                For k = 0 To UBound(語々)
                    If Right$(元, Len(語々(k))) = 語々(k) And (Len(語々(k)) >= 2 Or 元 = 語々(k)) Then
                        For j = 0 To UBound(語々)
                            語 = 語々(j)
                            位置 = InStr(1, 伏, 語, vbBinaryCompare)
                            Do While 位置 > 0
                                後 = Mid$(伏, 位置 + Len(語), 1)
                                If Len(語) >= 2 Or Len(後) = 0 Or InStr("別ご毎単でをとの・,、にが", 後) > 0 Then Exit Do
                                位置 = InStr(位置 + 1, 伏, 語, vbBinaryCompare)
                            Loop
                            If 位置 > 0 And (最初 = 0 Or 位置 < 最初) Then 最初 = 位置: 当て語 = 語
                        Next j
                        Exit For
                    End If
                Next k
            Next g
            位置々(c) = 最初: 当て々(c) = 当て語
        End If
    Next c
    For c = 左 To 右
        If 位置々(c) > 0 Then
            後 = Mid$(頼, 位置々(c) + Len(当て々(c)), 2)
            区切りか(c) = (Left$(後, 1) = "別" Or 後 = "ごと" Or Left$(後, 1) = "毎" Or 後 = "単位")
        End If
    Next c
    ' 「月別」＝日付の列を月でまとめる行（か列）にする（2026-10-09）。日付の列の見出しを書いていなくても「月」の位置で当てる
    Dim 月列 As Long, 月位置 As Long
    月列 = 頼みの月の列(ws, 見出し行, 左, 右, 最終行, 頼, 月位置)
    If 月列 > 0 Then
        If 位置々(月列) = 0 Then 位置々(月列) = 月位置: 当て々(月列) = "月"
        区切りか(月列) = True
    End If
    ' 同じ位置の語に 2 つの見出しが当たったら（「金額」とだけ書いて、実績金額と予算金額）、どちらか決まらない＝作らない
    For c = 左 To 右
        If 位置々(c) > 0 Then
            For k = 左 To c - 1
                If 位置々(k) = 位置々(c) Then
                    候補 = "「" & 当て々(c) & "」が " & 見出しの字(ws.Cells(見出し行, k)) & "・" & 見出しの字(ws.Cells(見出し行, c)) & _
                           " のどちらか決まりません（見出しのとおりに書いてください）。見出し: " & 一覧
                    Exit Function
                End If
            Next k
        End If
    Next c
    ' 行・列の札（数でない列か、別・ごと の続く数の列）と値（数の列）に分ける
    For c = 左 To 右
        If 位置々(c) > 0 Then
            If 数か(c) And Not 区切りか(c) Then
                値々 = 値々 & "・" & 見出しの字(ws.Cells(見出し行, c))
                If 値の列 = 0 Then 値の列 = c Else 値の列 = -1
            Else
                札数 = 札数 + 1
                If 札数 <= 3 Then 札(札数) = c
                札々 = 札々 & "・" & 見出しの字(ws.Cells(見出し行, c))
            End If
        End If
    Next c
    候補 = "見出し: " & 一覧
    If 値の列 = -1 Then
        候補 = "値にする数の列が 2 つ以上当たりました（" & Mid$(値々, 2) & "）。" & 候補
        値の列 = 0
        Exit Function
    End If
    If 札数 = 0 Then
        候補 = "行にする見出しが頼みの文にありません。" & 候補
        Exit Function
    End If
    If 札数 > 2 Then
        候補 = "行・列にする見出しが 3 つ以上当たりました（" & Mid$(札々, 2) & "）。" & 候補
        Exit Function
    End If
    If 値の列 = 0 And 数えるだけ Then 値の列 = 札(1)       ' 「件数」は行の項目そのものを数える（2026-10-09）
    If 値の列 = 0 Then
        ' 値を言っていない＝合計できる数の列（番号・コード・年・月・率の列を除く）が 1 本だけならそれ
        For c = 左 To 右
            If 数か(c) And c <> 札(1) And c <> 札(2) Then
                見 = 見出しの字(ws.Cells(見出し行, c))
                If Not (見 Like "*番号*" Or 見 Like "*No*" Or 見 Like "*NO*" Or 見 Like "*№*" Or 見 Like "*コード*" _
                        Or 見 Like "*ID*" Or 見 Like "*年*" Or 見 Like "*月*" Or 見 Like "*日*" Or 見 Like "*率*") Then
                    値候補数 = 値候補数 + 1
                    値候補 = c
                End If
            End If
        Next c
        If 値候補数 <> 1 Then
            候補 = "値にする数の列が決まりません（「金額を」のように頼んでください）。" & 候補
            Exit Function
        End If
        値の列 = 値候補
    End If
    ' 行と列: 頼みに出てくる順。「行は部門」「列に区分」と言われた見出しはそこへ
    If 札数 = 2 Then
        If 位置々(札(2)) < 位置々(札(1)) Then 入替 = 札(1): 札(1) = 札(2): 札(2) = 入替
        置1 = ピボットの置き場(頼, 当て々(札(1)))
        置2 = ピボットの置き場(頼, 当て々(札(2)))
        If 置1 = 2 Or 置2 = 1 Then
            If Not (置1 = 1 Or 置2 = 2) Then 入替 = 札(1): 札(1) = 札(2): 札(2) = 入替
        End If
        列の列 = 札(2)
    End If
    行の列 = 札(1)
    候補 = ""
    頼みからピボットの列 = True
End Function


' 見出しと頼みの文を比べる前のならし（2026-10-09）: 全角の英数・記号を半角に・空白を落とす・括弧の中（（円）（税抜）など）を落とす。
'   かなは触らない（StrConv の vbNarrow はカタカナまで半角にして、言い換えの組の語と合わなくなる）
Function ピボットの語ならし(ByVal 文 As String) As String
    Dim s As String, i As Long, w As Long, 開 As Long, 閉 As Long
    For i = 1 To Len(文)
        w = AscW(Mid$(文, i, 1))
        If w >= &HFF01 And w <= &HFF5E Then
            s = s & ChrW(w - &HFEE0)
        ElseIf w <> 32 And w <> &H3000 And w <> 10 And w <> 13 And w <> 9 Then
            s = s & Mid$(文, i, 1)
        End If
    Next i
    Do
        開 = InStr(s, "(")
        If 開 = 0 Then Exit Do
        閉 = InStr(開, s, ")")
        If 閉 = 0 Then Exit Do
        s = Left$(s, 開 - 1) & Mid$(s, 閉 + 1)
    Loop
    ピボットの語ならし = s
End Function



' 先撃ち（Excelコンボ・seiri）が渡した頼みの文（2026-10-09）。棚から人が直に撃ったとき・コンボ道具が無いブックでは ""
Function 先撃ちの頼みの文() As String
    On Error Resume Next
    先撃ちの頼みの文 = CStr(Application.Run("'" & ThisWorkbook.Name & "'!コンボ道具.頼みの文を返す"))
End Function

' 棚のマクロが撃てなかった理由を、ステータスバーと先撃ち（コンボ道具の断りの口）に出す（2026-10-09）。戻り値は使わない
Function 先撃ちに断る(ByVal 文 As String) As Boolean
    Application.StatusBar = 文
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!コンボ道具.棚の断りを入れる", 文
End Function

' 頼みの文で、見出しの語がピボットのどこに置くよう言われているか（2026-10-09）。1＝行・2＝列・0＝言われていない。
'   語が何度出てきても全部見る（「商品区分別に行は部門、列に商品区分で」の 2 度目の「列に商品区分」・段 2 の Q1-4）。
'   行: 「行に部門」「行は部門」「縦に」「部門を行」　列: 「列に区分」「列は区分」「横に」「区分を列」「区分は横」
Function ピボットの置き場(ByVal 頼 As String, ByVal 語 As String) As Long
    Dim 位置 As Long, 前 As String, 後 As String
    If Len(語) = 0 Then Exit Function
    位置 = InStr(1, 頼, 語, vbBinaryCompare)
    Do While 位置 > 0
        前 = Mid$(頼, IIf(位置 > 4, 位置 - 4, 1), IIf(位置 > 4, 4, 位置 - 1))
        後 = Left$(Mid$(頼, 位置 + Len(語), 4), 2)
        If Right$(前, 2) = "列に" Or Right$(前, 2) = "列は" Or Right$(前, 2) = "横に" Or Right$(前, 3) = "列方向" Or _
           Right$(前, 3) = "横軸に" Or 後 = "を列" Or 後 = "は列" Or 後 = "が列" Or 後 = "も列" Or 後 = "を横" Or 後 = "は横" Or 後 = "も横" Then
            ピボットの置き場 = 2
            Exit Function
        End If
        If Right$(前, 2) = "行に" Or Right$(前, 2) = "行は" Or Right$(前, 2) = "縦に" Or Right$(前, 3) = "行方向" Or _
           Right$(前, 3) = "縦軸に" Or 後 = "を行" Or 後 = "は行" Or 後 = "が行" Or 後 = "も行" Or 後 = "を縦" Or 後 = "は縦" Or 後 = "も縦" Then
            ピボットの置き場 = 1
            Exit Function
        End If
        位置 = InStr(位置 + 1, 頼, 語, vbBinaryCompare)
    Loop
End Function


' 頼みの文でグラフを作る（2026-10-09 段 4・棒グラフと円グラフから呼ぶ）。種類＝"棒" か "円"。
'   戻り値 0＝頼みに見出しが無い（呼び元は今までどおり左端の文字の列・右端の数の列で作る）／1＝作った・断った（呼び元は終わる）
'   ・ピボットのシートなら、そのピボットからピボットグラフ
'   ・頼みの文で項目と値が決まれば、項目ごとに合計して描く（明細の表でも部署 1 つに棒 1 本）
'   ・項目が 2 つ（「部署別・区分ごとに」）なら、ピボットを作ってそこからピボットグラフ（円グラフは断る）
'   ・決まらなければ作らずに候補の見出しを返す（勝手な列で作らない）
'   ・件数・平均・最大・最小（頼みの集計）と月別（日付の列を年月でまとめ、月の順に並べる）も（2026-10-09）
Function 頼みでグラフを作る(ByVal ws As Worksheet, ByVal 頼 As String, ByVal 種類 As String) As Long
    Dim ur As Range, 左 As Long, 右 As Long, 上 As Long, 見出し行 As Long, 最終行 As Long
    Dim r As Long, c As Long, 数 As Long, 字 As Long, v As Variant, 集計行 As Boolean
    Dim 項目列 As Long, 列の列 As Long, 値列 As Long, 候補 As String, 見たか As Boolean
    Dim 名() As String, 値() As Double, n As Long, i As Long, j As Long, 名1 As String, 値1 As Double
    Dim 系列 As String, 題 As String, co As ChartObject, ch As Chart, pt As PivotTable
    Dim 集計 As Long, 月列 As Long, 月位置 As Long, 件() As Long, 大() As Double, 小() As Double, 鍵() As Double, 鍵1 As Double

    頼みでグラフを作る = 0
    If ws.PivotTables.Count > 0 Then
        Set pt = ws.PivotTables(1)
        If 種類 = "円" And pt.ColumnFields.Count > 0 Then
            先撃ちに断る "列のあるピボット（" & pt.ColumnFields(1).Name & "）は円グラフにできません。棒グラフで頼んでください。"
            頼みでグラフを作る = 1
            Exit Function
        End If
        ピボットからグラフを作る pt, 種類
        頼みでグラフを作る = 1
        Exit Function
    End If

    Set ur = ws.UsedRange
    左 = ur.Column: 右 = ur.Column + ur.Columns.Count - 1: 上 = ur.Row
    For r = 上 To 上 + ur.rows.Count - 1
        数 = 0: 字 = 0
        For c = 左 To 右
            v = ws.Cells(r, c).Value
            If IsNumeric(v) And Trim(セルの字(v)) <> "" Then
                数 = 数 + 1
            ElseIf Trim(セルの字(v)) <> "" Then
                字 = 字 + 1
            End If
        Next c
        If 字 >= 2 And 数 = 0 Then 見出し行 = r: Exit For
    Next r
    If 見出し行 = 0 Then Exit Function
    最終行 = 表の本文の最終行(ws)
    集計 = 頼みの集計(頼)
    If 頼みからピボットの列(ws, 見出し行, 左, 右, 最終行, 頼, 項目列, 列の列, 値列, 候補, 集計 = xlCount) Then
        If 列の列 > 0 Then
            If 種類 = "円" Then
                先撃ちに断る "円グラフの項目は 1 つです（" & 見出しの字(ws.Cells(見出し行, 項目列)) & "・" & _
                    見出しの字(ws.Cells(見出し行, 列の列)) & " の 2 つが当たりました）。棒グラフで頼むか、項目を 1 つにしてください。"
                頼みでグラフを作る = 1
                Exit Function
            End If
            ' 項目が 2 つ＝ピボットを作り（頼みの文で列を決める）、そこからピボットグラフ
            選んだ3列でピボットを作る
            If ActiveSheet.PivotTables.Count > 0 Then ピボットからグラフを作る ActiveSheet.PivotTables(1), 種類
            頼みでグラフを作る = 1
            Exit Function
        End If
    Else
        ' 見出しが 1 つも当たらない頼み（「棒グラフにして」）は今までどおり。当たったのに決まらないときは断る
        ' ただし文字の列が 2 本以上ある表（明細）は、どれを項目にするか決まらない＝作らずに断る（段 4 の G1-5「グラフ作って」で
        ' 伝票番号の 12 行がそのまま棒になった）。文字の列が 1 本だけの集計表は今までどおり
        If Left$(候補, Len("行にする見出しが")) = "行にする見出しが" And 値列 = 0 Then
            字 = 0
            For c = 左 To 右
                v = ws.Cells(見出し行 + 1, c).Value
                If Not IsError(v) Then
                    If VarType(v) = vbString And Len(Trim$(セルの字(v))) > 0 Then
                        If Not IsNumeric(v) Then 字 = 字 + 1
                    End If
                End If
            Next c
            If 字 < 2 Then Exit Function
            先撃ちに断る "グラフの項目にする列が決まらないので作っていません（文字の列が " & 字 & " 本あります）。" & _
                Mid$(候補, InStr(候補, "見出し:")) & "（頼み方の例: 「部署別に金額を棒グラフで」）"
            頼みでグラフを作る = 1
            Exit Function
        End If
        先撃ちに断る "グラフの項目と値が決まらないので作っていません。" & 候補 & "（頼み方の例: 「部署別に金額を棒グラフで」）"
        頼みでグラフを作る = 1
        Exit Function
    End If

    If 集計 = xlCount And InStr(見出しの字(ws.Cells(見出し行, 値列)), "件数") > 0 And 値列 <> 項目列 Then 集計 = xlSum
    月列 = 頼みの月の列(ws, 見出し行, 左, 右, 最終行, ピボットの語ならし(頼), 月位置)
    ' 項目ごとに集計（合計・小計の行は外す。月の列は年月でまとめる。円グラフは 0 以下を外す）
    ReDim 名(1 To 最終行 - 見出し行 + 1): ReDim 値(1 To 最終行 - 見出し行 + 1): ReDim 件(1 To 最終行 - 見出し行 + 1)
    ReDim 大(1 To 最終行 - 見出し行 + 1): ReDim 小(1 To 最終行 - 見出し行 + 1): ReDim 鍵(1 To 最終行 - 見出し行 + 1)
    For r = 見出し行 + 1 To 最終行
        集計行 = False
        For c = 左 To 右
            If 集計の語か(Trim$(セルの字(ws.Cells(r, c).Value))) Then 集計行 = True: Exit For
        Next c
        名1 = Trim$(セルの字(ws.Cells(r, 項目列).Value))
        鍵1 = 0
        If 項目列 = 月列 And VarType(ws.Cells(r, 項目列).Value) = vbDate Then
            鍵1 = Year(ws.Cells(r, 項目列).Value) * 100 + Month(ws.Cells(r, 項目列).Value)
            名1 = Year(ws.Cells(r, 項目列).Value) & "年" & Month(ws.Cells(r, 項目列).Value) & "月"
        End If
        v = ws.Cells(r, 値列).Value
        If Not 集計行 And Len(名1) > 0 And Not IsError(v) Then
            If Len(Trim$(セルの字(v))) > 0 And (集計 = xlCount Or IsNumeric(v)) Then
                For i = 1 To n
                    If 名(i) = 名1 Then Exit For
                Next i
                If i > n Then n = n + 1: 名(n) = 名1: 鍵(n) = 鍵1
                件(i) = 件(i) + 1
                If 集計 <> xlCount Then
                    If 件(i) = 1 Then 大(i) = CDbl(v): 小(i) = CDbl(v)
                    値(i) = 値(i) + CDbl(v)
                    If CDbl(v) > 大(i) Then 大(i) = CDbl(v)
                    If CDbl(v) < 小(i) Then 小(i) = CDbl(v)
                End If
            End If
        End If
    Next r
    For i = 1 To n
        Select Case 集計
            Case xlCount: 値(i) = 件(i)
            Case xlAverage: 値(i) = 値(i) / 件(i)
            Case xlMax: 値(i) = 大(i)
            Case xlMin: 値(i) = 小(i)
        End Select
    Next i
    If 種類 = "円" Then
        j = 0
        For i = 1 To n
            If 値(i) > 0 Then j = j + 1: 名(j) = 名(i): 値(j) = 値(i): 鍵(j) = 鍵(i)
        Next i
        n = j
    End If
    If n = 0 Then
        先撃ちに断る "グラフにする値がありません（" & 見出しの字(ws.Cells(見出し行, 値列)) & "）。"
        頼みでグラフを作る = 1
        Exit Function
    End If
    ' 大きい順（月でまとめたときは月の順）
    For i = 1 To n - 1
        For j = 1 To n - i
            If IIf(項目列 = 月列, 鍵(j) > 鍵(j + 1), 値(j) < 値(j + 1)) Then
                値1 = 値(j): 値(j) = 値(j + 1): 値(j + 1) = 値1
                名1 = 名(j): 名(j) = 名(j + 1): 名(j + 1) = 名1
                鍵1 = 鍵(j): 鍵(j) = 鍵(j + 1): 鍵(j + 1) = 鍵1
            End If
        Next j
    Next i
    Dim 名々() As String, 値々() As Double
    ReDim 名々(1 To n): ReDim 値々(1 To n)
    For i = 1 To n
        名々(i) = 名(i): 値々(i) = 値(i)
    Next i

    系列 = 見出しの字(ws.Cells(見出し行, 値列))
    Select Case 集計
        Case xlCount: 系列 = "件数"
        Case xlAverage: 系列 = 系列 & "の平均"
        Case xlMax: 系列 = 系列 & "の最大"
        Case xlMin: 系列 = 系列 & "の最小"
    End Select
    題 = IIf(項目列 = 月列, "月", 見出しの字(ws.Cells(見出し行, 項目列))) & "別の" & 系列 & IIf(種類 = "円", "の構成比", "")
    ' 同じ種類の棚のグラフだけ作り直し、別の種類の棚のグラフがあればその下に置く（人のグラフは残す）
    名1 = IIf(種類 = "円", "棚_円グラフ", "棚_棒グラフ")
    Set co = ws.ChartObjects.Add(ws.Cells(見出し行, 右 + 2).Left, _
                                 棚のグラフの置き場(ws, 名1, ws.Cells(見出し行, 右 + 2).Top), IIf(種類 = "円", 360, 440), 300)
    co.Name = 名1
    Set ch = co.Chart
    If 種類 = "円" Then
        ch.ChartType = xlPie
    ElseIf n >= 9 Then
        ch.ChartType = xlBarClustered
    Else
        ch.ChartType = xlColumnClustered
    End If
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.SeriesCollection.NewSeries
    ch.SeriesCollection(1).Values = 値々
    ch.SeriesCollection(1).XValues = 名々
    ch.SeriesCollection(1).Name = 系列
    グラフの体裁 ch, 種類, 題, n
    頼みでグラフを作る = 1
End Function


' ピボットからピボットグラフを作る（2026-10-09 段 4）。種類＝"棒" か "円"。ピボットの右に置き、棚が前に作ったグラフだけ作り直す。
'   ピボットグラフなので、ピボットの並べ替え・絞り込み・更新にグラフがついてくる
Sub ピボットからグラフを作る(ByVal pt As PivotTable, ByVal 種類 As String)
    Dim sh As Worksheet, co As ChartObject, ch As Chart, 右 As Range, 題 As String, n As Long
    Set sh = pt.Parent
    Set 右 = pt.TableRange1
    題 = IIf(種類 = "円", "棚_円グラフ", "棚_ピボットグラフ")
    Set co = sh.ChartObjects.Add(右.Offset(0, 右.Columns.Count + 1).Left, 棚のグラフの置き場(sh, 題, 右.Top), 480, 300)
    co.Name = 題
    Set ch = co.Chart
    ch.SetSourceData pt.TableRange1
    ch.ChartType = IIf(種類 = "円", xlPie, xlColumnClustered)
    On Error Resume Next
    題 = pt.DataFields(1).SourceName
    If InStr(pt.DataFields(1).Caption, " / ") = 0 Then 題 = pt.DataFields(1).Caption     ' 件数・平均は見出しの名で
    If pt.RowFields.Count > 0 Then 題 = pt.RowFields(1).Name & "別の" & 題
    If pt.ColumnFields.Count > 0 Then 題 = 題 & "（" & pt.ColumnFields(1).Name & "ごと）"
    n = pt.RowFields(1).PivotItems.Count
    On Error GoTo 0
    グラフの体裁 ch, 種類, 題 & IIf(種類 = "円", "の構成比", ""), n
    ' ピボットグラフの中のボタン（フィールドの絞り込み）は残す＝人がピボットと同じように絞れる
    If pt.ColumnFields.Count > 0 And 種類 <> "円" Then
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionRight
    End If
End Sub

' 棚のグラフの見た目をそろえる（2026-10-09 段 4・棒グラフの今までの形に合わせる）。種類＝"棒" か "円"、n＝項目の数
'   棒: 系列 1 本なら紺で一番大きい棒だけ赤・値を表示・凡例なし・数の軸は #,##0・目盛線なし／円: 割合を表示・凡例は右
Sub グラフの体裁(ByVal ch As Chart, ByVal 種類 As String, ByVal 題 As String, ByVal n As Long)
    Dim sr As Series, 値 As Variant, i As Long, 最大 As Long
    On Error Resume Next
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    ch.HasTitle = True
    ch.chartTitle.text = 題
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Size = 14
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Bold = True
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    If 種類 = "円" Then
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionRight
        Set sr = ch.SeriesCollection(1)
        sr.HasDataLabels = True
        sr.DataLabels.ShowPercentage = True
        sr.DataLabels.ShowValue = False
        sr.DataLabels.ShowCategoryName = False
        sr.DataLabels.ShowSeriesName = False
        Exit Sub
    End If
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlValue).HasMajorGridlines = False
    ch.ChartGroups(1).GapWidth = 50
    If ch.ChartType = xlBarClustered Then ch.Axes(xlCategory).ReversePlotOrder = True
    If ch.SeriesCollection.Count = 1 Then
        ch.HasLegend = False
        Set sr = ch.SeriesCollection(1)
        sr.Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        値 = sr.Values
        最大 = LBound(値)
        For i = LBound(値) To UBound(値)
            If 値(i) > 値(最大) Then 最大 = i
        Next i
        sr.Points(最大 - LBound(値) + 1).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
        sr.HasDataLabels = True
        sr.DataLabels.NumberFormat = "#,##0"
    End If
End Sub

' 棚のグラフを置く上の位置（2026-10-09 段 4）: 同じ名前の棚のグラフは消して作り直し、別の種類の棚のグラフがあればその下に置く
'   （棒グラフの後に円グラフを頼んだら棒グラフが消えていた）。人のグラフは触らない
Function 棚のグラフの置き場(ByVal ws As Object, ByVal 名 As String, ByVal 既定の上 As Double) As Double
    Dim i As Long, 下 As Double
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = 名 Then ws.ChartObjects(i).Delete
    Next i
    下 = 既定の上
    For i = 1 To ws.ChartObjects.Count
        If Left$(ws.ChartObjects(i).Name, 2) = "棚_" Then
            If ws.ChartObjects(i).Top + ws.ChartObjects(i).Height + 10 > 下 Then 下 = ws.ChartObjects(i).Top + ws.ChartObjects(i).Height + 10
        End If
    Next i
    棚のグラフの置き場 = 下
End Function

' 頼みの文の集計の種類（2026-10-09）: 件数・何件・カウント＝xlCount／平均＝xlAverage／最大・最高＝xlMax／最小・最低＝xlMin／それ以外＝xlSum
Function 頼みの集計(ByVal 頼 As String) As Long
    頼みの集計 = xlSum
    If InStr(頼, "件数") > 0 Or InStr(頼, "何件") > 0 Or InStr(頼, "カウント") > 0 Or InStr(頼, "数を数え") > 0 Or InStr(頼, "人数") > 0 Then
        頼みの集計 = xlCount
    ElseIf InStr(頼, "平均") > 0 Then
        頼みの集計 = xlAverage
    ElseIf InStr(頼, "最大") > 0 Or InStr(頼, "最高") > 0 Then
        頼みの集計 = xlMax
    ElseIf InStr(頼, "最小") > 0 Or InStr(頼, "最低") > 0 Then
        頼みの集計 = xlMin
    End If
End Function

' 頼みの文が「月別」（月ごと・毎月・月単位・月次・各月）を求めていれば、月でまとめる日付の列を返す（2026-10-09）。無ければ 0。
'   日付の列が 1 本ならそれ。2 本以上なら頼みに見出しが書いてある日付の列（「受注日の月別」）。決まらなければ 0。月の位置＝頼みの中の「月」の位置
Function 頼みの月の列(ByVal ws As Worksheet, ByVal 見出し行 As Long, ByVal 左 As Long, ByVal 右 As Long, _
        ByVal 最終行 As Long, ByVal 頼 As String, ByRef 月の位置 As Long) As Long
    Dim 語 As Variant, p As Long, c As Long, r As Long, 日 As Long, 字 As Long, v As Variant, 候補数 As Long, 候補 As Long, 書いた As Long
    月の位置 = 0
    For Each 語 In Array("月別", "月ごと", "毎月", "月単位", "月次", "各月", "月ベース")
        p = InStr(1, 頼, 語, vbBinaryCompare)
        If p > 0 And (月の位置 = 0 Or p < 月の位置) Then 月の位置 = p
    Next 語
    If 月の位置 = 0 Then Exit Function
    For c = 左 To 右
        日 = 0: 字 = 0
        For r = 見出し行 + 1 To 最終行
            v = ws.Cells(r, c).Value
            If Not IsError(v) And Not isEmpty(v) Then
                字 = 字 + 1
                If VarType(v) = vbDate Then 日 = 日 + 1
            End If
        Next r
        If 字 > 0 And 日 >= 字 * 0.8 Then
            候補数 = 候補数 + 1
            候補 = c
            If InStr(1, 頼, ピボットの語ならし(見出しの字(ws.Cells(見出し行, c))), vbBinaryCompare) > 0 Then 書いた = c
        End If
    Next c
    If 候補数 = 1 Then
        頼みの月の列 = 候補
    ElseIf 書いた > 0 Then
        頼みの月の列 = 書いた
    End If
End Function


' 頼みの文の中で、見出し（ピボットのフィールド名）を指す語の位置（2026-10-09 ピボットの手直し）。見出しそのもの・言い換えの組の語。無ければ 0
'   組は 頼みからピボットの列 と同じ（部署／部門／課、金額／売上／売上高 …）。1 字の語は後ろに 別・ごと・を・も 等が続くときだけ
Function 見出しの語の位置(ByVal 頼 As String, ByVal 見出し As String, Optional ByRef 当て語 As String) As Long
    Dim 組 As Variant, 語々 As Variant, 見 As String, 元 As String, g As Long, k As Long, j As Long, p As Long, 後 As String
    頼 = ピボットの語ならし(頼)
    見 = ピボットの語ならし(見出し)
    当て語 = ""
    If Len(見) = 0 Then Exit Function
    p = InStr(1, 頼, 見, vbBinaryCompare)
    If p > 0 Then 見出しの語の位置 = p: 当て語 = 見: Exit Function
    組 = Array("部署|部門|課|所属|部局|部課", "金額|売上|売上高|売り上げ|額|支出額|請求額|支払額|販売額", _
              "区分|種別|分類|カテゴリ|カテゴリー|種類", "担当|担当者|営業担当|社員", _
              "商品|品名|品目|製品", "取引先|得意先|顧客", "地域|エリア|地区|支店|店舗|拠点", "数量|個数|台数", _
              "費目|科目|勘定科目")
    元 = 見
    If Len(元) > 2 And Right$(元, 1) = "名" Then 元 = Left$(元, Len(元) - 1)
    For g = 0 To UBound(組)
        語々 = Split(組(g), "|")
        For k = 0 To UBound(語々)
            If Right$(元, Len(語々(k))) = 語々(k) And (Len(語々(k)) >= 2 Or 元 = 語々(k)) Then
                For j = 0 To UBound(語々)
                    p = InStr(1, 頼, 語々(j), vbBinaryCompare)
                    Do While p > 0
                        後 = Mid$(頼, p + Len(語々(j)), 1)
                        If Len(語々(j)) >= 2 Or Len(後) = 0 Or InStr("別ご毎単でをとの・,、にがもは", 後) > 0 Then Exit Do
                        p = InStr(p + 1, 頼, 語々(j), vbBinaryCompare)
                    Loop
                    If p > 0 And (見出しの語の位置 = 0 Or p < 見出しの語の位置) Then 見出しの語の位置 = p: 当て語 = 語々(j)
                Next j
                Exit For
            End If
        Next k
    Next g
End Function


' 文に「語|語|語」のどれかが入っているか（2026-10-09 ピボットの手直し・グラフの種類）
Function 含む語(ByVal 文 As String, ByVal 語々 As String) As Boolean
    Dim 語 As Variant
    For Each 語 In Split(語々, "|")
        If Len(語) > 0 And InStr(1, 文, 語, vbBinaryCompare) > 0 Then 含む語 = True: Exit Function
    Next 語
End Function
