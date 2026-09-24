Attribute VB_Name = "表の整理"
' 表の整理 - アクティブシートの一覧表（見出し 1 行＋本文）を直す 2 本（2026-09-11 夜・09-12 深夜に強化）
'   表の書き方と罫線と列幅をそろえる   … 値の意味は変えずに書き方をそろえる。行は消さない。数式のセルには触らない。
'                   前後の空白・全角の英数字→半角・半角カナ→全角・文字の間の空白（列の多い方の 1 つに）・
'                   会社名の（株）㈱→株式会社・電話 03-1234-5678（+81 も）・郵便 000-0000・メールは小文字・
'                   文字の日付／和暦／英語／8 桁／数値のままの日付→日付（yyyy/m/d）・文字の金額／「12個」／△▲(　)の負の数→数値（#,##0）・
'                   5%・５％→率（0%）・セルの中の改行→空白・フリガナのひらがな→カタカナ・カタカナの後ろの - → ー（長音）・
'                   状態の列の言い換え（送り仮名の有無・ひらがな書き → 漢字の多い形）・
'                   本文の文字の書式（見出しの書体と大きさ・太字なし・色は自動）・塗りのばらつき・罫線・列幅
'   全列が同じ重複行を削除する … 全部の列が同じ行（書き方の違いはならして比べる）の 2 行目以降を下から消す。
'                   1 列でも中身が違えば別の行として残す。会社名の前株・後株は区別する。途中の空行は消さない。表の外（右の列）には触らない。
'   表の範囲 … 見出しの下から、空行が 2 行続く手前まで（途中の空行 1 行は表の続き）



Sub 表の書き方と罫線と列幅をそろえる()
    ' 依頼の語: 表を整え|表記を統一|全角と半角|書き方をそろえ|書き方を揃え|入力の書き方|見やすく整理|表を整理|表をきれいに|表記をそろえ|表記を揃え|表記ゆれ|表記揺れ|表記の揺れ|全角半角|表の書式をそろえ|表の書式を整え|表の見た目を整え|表の体裁を整え|表全体を整え|整ってない|整っていない
    ' 依頼の組: 表記,書き方,全角,半角,ばらばら,バラバラ,ごちゃごちゃ,ひらがな,カタカナ,混ざ,(株),株式会社,形式,書式,罫線+そろえ,揃え,統一,整え,整理,直,引き直+-住所,番地,郵便,グラフ,シート,名前の揺,列幅,ピボット,データ,洗い出,一覧,探,調べ
    ' 値の意味は変えずに書き方をそろえる（行は消さない・数式のセルは触らない）: 前後の空白・全角英数→半角・半角カナ→全角・
    ' 文字の間の空白（列の多い方の 1 つ）・（株）→株式会社・電話・郵便番号・メール・文字の日付と和暦→日付・文字の金額→数値・率・罫線・列幅
    Dim ws As Worksheet, ur As Range, tb As Range, c As Range, dv As Object
    Dim r0 As Long, c0 As Long, nR As Long, nC As Long, scanTo As Long
    Dim hr As Long, hc1 As Long, hc2 As Long, lastR As Long
    Dim i As Long, j As Long, k As Long, n As Long, best As Long, ln As Long, a As Long, b As Long, g As Long
    Dim v As Variant, nv As Variant, x As Variant, parts As Variant, ks As Variant
    Dim s As String, t As String, u As String, ch As String, hd As String, kind As String, sp As String
    Dim digits As String, kanaRun As String, mon As String
    Dim cnt As Long, dn As Long, nn As Long, cn As Long, zs As Long, hs As Long, an As Long, pn As Long
    Dim ok As Boolean, bad As Boolean, dt As Date, dbl As Double
    Dim era As Long, y As Long, m As Long, d As Long, cd As Long
    Dim ar(1 To 200) As String, al(1 To 200) As Long, bl(1 To 200) As Long, nA As Long
    Dim grp() As Long, k1() As String, k2() As String, rep() As String, hira() As Long, cntOf() As Long
    Dim hasKana() As Boolean, hasKanji() As Boolean
    Dim mainEnd As Long, sumR1 As Long, sumC1 As Long, sumC2 As Long, gapR As Long, rr As Long
    Dim colW() As Double

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    r0 = ur.Row: c0 = ur.Column: nR = ur.rows.Count: nC = ur.Columns.Count
    If nR < 2 Then Exit Sub
    GoSub 見出しを探す
    If hr = 0 Then Exit Sub
    ' 見出しの行を「列まるごと空の列」と「式の見出し（=COUNTA などの集計）」で区切り、見出しが 2 つ以上ある塊ごとに
    ' 1 つの表として直す（2026-09-23）。左右に並んだ 2 つの表と右のメモの列を 1 つの表と見て、間の空き列・右の表の下・
    ' メモの列にまで罫線を引き、メモの全角かっこを半角にしていた（お試し版テスト用4）。見出しの無い列も中に値があれば
    ' 表の列、2024 のような数の見出しも区切りにしない（年を並べた横持ちの表を割らない）
    Dim blk1(1 To 50) As Long, blk2(1 To 50) As Long, nBlk As Long, bi As Long, inBlk As Boolean
    nBlk = 0: inBlk = False
    For j = c0 To c0 + nC
        ok = False
        If j <= c0 + nC - 1 Then
            If ws.Cells(hr, j).HasFormula Then
                ok = False
            ElseIf Len(Trim(ws.Cells(hr, j).text)) > 0 Then
                ok = True
            Else
                ok = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(hr, j), ws.Cells(r0 + nR - 1, j))) > 0
            End If
        End If
        If ok Then
            If Not inBlk And nBlk < 50 Then
                nBlk = nBlk + 1
                blk1(nBlk) = j
                inBlk = True
            End If
            blk2(nBlk) = j
        Else
            inBlk = False
        End If
    Next
    For bi = 1 To nBlk
        hc1 = blk1(bi): hc2 = blk2(bi)
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(hr, hc1), ws.Cells(hr, hc2))) >= 2 Then
            GoSub 本文の終わり
            If lastR > hr Then
                GoSub 集計欄を分ける      ' 空行の下の細い塊（合計・平均）で表を終える。その先の行（注記など）は触らない（2026-09-13）
                GoSub 一つの表を直す
            End If
        End If
    Next
    Exit Sub

一つの表を直す:
    For j = hc1 To hc2
        hd = UCase(StrConv(ws.Cells(hr, j).text, vbNarrow))
        GoSub 列を調べる
        For i = hr + 1 To lastR
            Set c = ws.Cells(i, j)
            If Not c.HasFormula Then
                v = c.Value
                If Not isEmpty(v) And Not IsError(v) Then
                    nv = v
                    GoSub 値を直す
                    If VarType(nv) <> VarType(v) Then
                        GoSub 書く
                    ElseIf nv <> v Then
                        GoSub 書く
                    End If
                End If
            End If
        Next
        If kind = "text" Then GoSub 言い換えをそろえる
        If kind = "date" Then ws.Range(ws.Cells(hr + 1, j), ws.Cells(lastR, j)).NumberFormat = "yyyy/m/d"
        If kind = "pct" Then ws.Range(ws.Cells(hr + 1, j), ws.Cells(lastR, j)).NumberFormat = "0%"
        If kind = "num" And InStr(hd, "年") = 0 And InStr(hd, "率") = 0 And InStr(hd, "%") = 0 Then
            ws.Range(ws.Cells(hr + 1, j), ws.Cells(lastR, j)).NumberFormat = "#,##0"
            GoSub 小数のセル
        End If
        ' 番号・コードの列は左寄せ（数の 1, 2, 3 が右に寄って金額に見えた・tidy と同じ決まり・2026-09-23）
        If kind = "code" Then ws.Range(ws.Cells(hr + 1, j), ws.Cells(lastR, j)).HorizontalAlignment = xlLeft
    Next

    ' 本文の文字の書式（見出しの書体と大きさ・太字／斜体／下線／取り消し線なし・文字色は自動）・塗りのばらつき・罫線・列幅
    Set tb = ws.Range(ws.Cells(hr + 1, hc1), ws.Cells(lastR, hc2))
    With tb.Font
        .Name = ws.Cells(hr, hc1).Font.Name
        .Size = ws.Cells(hr, hc1).Font.Size
    End With
    ' 太字・塗りをそろえるのは明細の行だけ。小計・合計の行の強調は情報なので残す（2026-09-24 通しの実測 4 で小計行の塗りと太字が消えた）
    ' 集計の行と行のあいだを塊でまとめる（行ごとに Union を重ねると大きな表で遅い）
    Dim keepSumR As Long, keepSumC As Long, keepBody As Range, keepIsSum As Boolean, keepFrom As Long
    keepFrom = hr + 1
    For keepSumR = hr + 1 To lastR + 1
        keepIsSum = (keepSumR > lastR)
        If Not keepIsSum Then
            For keepSumC = hc1 To hc2
                If 集計の語か(ws.Cells(keepSumR, keepSumC).Value) Then keepIsSum = True: Exit For
            Next keepSumC
        End If
        If keepIsSum Then
            If keepSumR > keepFrom Then
                If keepBody Is Nothing Then
                    Set keepBody = ws.Range(ws.Cells(keepFrom, hc1), ws.Cells(keepSumR - 1, hc2))
                Else
                    Set keepBody = Union(keepBody, ws.Range(ws.Cells(keepFrom, hc1), ws.Cells(keepSumR - 1, hc2)))
                End If
            End If
            keepFrom = keepSumR + 1
        End If
    Next keepSumR
    If Not keepBody Is Nothing Then
        With keepBody.Font
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .ColorIndex = xlColorIndexAutomatic
        End With
        If IsNull(keepBody.Interior.ColorIndex) Then keepBody.Interior.ColorIndex = xlNone
    End If
    ' 引き直す前に二重線の場所を覚える（下で戻す）
    Dim dblTop() As Boolean, dblBot() As Boolean, dblI As Long, dblJ As Long
    ReDim dblTop(1 To lastR - hr + 1, 1 To hc2 - hc1 + 1)
    ReDim dblBot(1 To lastR - hr + 1, 1 To hc2 - hc1 + 1)
    For dblI = hr To lastR
        For dblJ = hc1 To hc2
            dblTop(dblI - hr + 1, dblJ - hc1 + 1) = (ws.Cells(dblI, dblJ).Borders(xlEdgeTop).LineStyle = xlDouble)
            dblBot(dblI - hr + 1, dblJ - hc1 + 1) = (ws.Cells(dblI, dblJ).Borders(xlEdgeBottom).LineStyle = xlDouble)
        Next dblJ
    Next dblI
    ' 罫線: 空の行の下の、表より細い塊（合計・平均など）は表と分けて、埋まっている列だけを囲む（2026-09-13）
    If sumR1 > 0 Then
        ' 前に引いた罫線が空の行と集計欄の左に残っていれば外す（外してから表と集計欄を引き直す）
        ws.Range(ws.Cells(mainEnd + 1, hc1), ws.Cells(lastR, hc2)).Borders.LineStyle = xlNone
    End If
    Set tb = ws.Range(ws.Cells(hr, hc1), ws.Cells(mainEnd, hc2))
    tb.Borders.LineStyle = xlContinuous
    tb.Borders.Weight = xlThin
    If sumR1 > 0 Then
        With ws.Range(ws.Cells(sumR1, sumC1), ws.Cells(lastR, sumC2)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
    ' 合計の上・下の二重線は引き直しで細線にしない（2026-09-24 通しの実測 4 で合計行の上の二重線が消えた）
    Dim dblR As Long, dblC As Long
    For dblR = hr To lastR
        For dblC = hc1 To hc2
            If dblTop(dblR - hr + 1, dblC - hc1 + 1) Then ws.Cells(dblR, dblC).Borders(xlEdgeTop).LineStyle = xlDouble
            If dblBot(dblR - hr + 1, dblC - hc1 + 1) Then ws.Cells(dblR, dblC).Borders(xlEdgeBottom).LineStyle = xlDouble
        Next dblC
    Next dblR
    ' 列幅: 表の中で合わせる。細くした列で表の外のセル（下の注記・D31 など）が ### になったら元の幅に戻す（2026-09-13）
    ReDim colW(hc1 To hc2)
    For k = hc1 To hc2
        colW(k) = ws.Columns(k).ColumnWidth
    Next
    ws.Range(ws.Cells(hr, hc1), ws.Cells(lastR, hc2)).Columns.AutoFit
    For k = hc1 To hc2
        If ws.Columns(k).ColumnWidth < colW(k) Then
            For rr = r0 To r0 + nR - 1
                If rr < hr Or rr > lastR Then
                    t = ws.Cells(rr, k).text
                    If Len(t) > 0 And Replace(t, "#", "") = "" Then
                        ws.Columns(k).ColumnWidth = colW(k)
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    Return

見出しを探す:
    ' 先頭 20 行で、値のある列が最も多い行の 6 割以上を最初に満たす行＝見出し。本文は見出しの列が全部空の行の手前まで
    scanTo = r0 + IIf(nR > 20, 20, nR) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nC - 1)))
        If n > best Then best = n
    Next
    hr = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nC - 1)))
        If hr = 0 And n >= 2 And n >= best * 0.6 Then hr = i
    Next
    Return

本文の終わり:
    ' 本文の終わり: 途中の空行 1 行は飛ばして続ける。空行が 2 行続いたら表の終わり
    lastR = hr
    n = 0
    For i = hr + 1 To r0 + nR - 1
        If n < 2 Then
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2))) = 0 Then
                n = n + 1
            Else
                n = 0
                lastR = i
            End If
        End If
    Next
    Return

集計欄を分ける:
    ' 空の行の下の塊が表より細い（合計・平均など）なら、表はその手前まで・塊は埋まっている列だけ。
    ' 表と同じ幅の塊は表の続き（抜けた 1 件の空行）＝空行ごと表に入れる（2026-09-13 お試し版テスト用1）
    mainEnd = lastR: sumR1 = 0: sumC1 = 0: sumC2 = 0
    i = hr + 1
    Do While i <= lastR
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2))) = 0 Then
            gapR = i
            Do While i <= lastR
                If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2))) > 0 Then Exit Do
                i = i + 1
            Loop
            sumC1 = 0: sumC2 = 0
            rr = i
            Do While rr <= lastR
                If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(rr, hc1), ws.Cells(rr, hc2))) = 0 Then Exit Do
                For k = hc1 To hc2
                    If Not isEmpty(ws.Cells(rr, k).Value) Then
                        If sumC1 = 0 Or k < sumC1 Then sumC1 = k
                        If k > sumC2 Then sumC2 = k
                    End If
                Next
                rr = rr + 1
            Loop
            If sumC1 > hc1 Or sumC2 < hc2 Then
                mainEnd = gapR - 1: sumR1 = i: lastR = rr - 1     ' その塊で表を終える（次の空行の先は表の外）
                For rr = sumR1 To lastR
                    For k = hc1 To hc2
                        If Not isEmpty(ws.Cells(rr, k).Value) Then
                            If k < sumC1 Then sumC1 = k
                            If k > sumC2 Then sumC2 = k
                        End If
                    Next
                Next
                Return
            End If
            i = rr
        Else
            i = i + 1
        End If
    Loop
    sumC1 = 0: sumC2 = 0
    Return

小数のセル:
    ' 数の列を #,##0 にすると、小数のセル（金額の下の「平均 10,162.5」など）が 10,163 と丸めて見える。
    ' そのセルだけ小数の見える形にする（2026-09-13）
    For i = hr + 1 To lastR
        v = ws.Cells(i, j).Value2
        If VarType(v) = vbDouble Then
            ' 計算の端数（12000*1.1＝13200.000000000002）は小数に数えない（「13,200.0」と「16,500」が混ざった・2026-09-23）
            If Abs(v - Round(v, 0)) > 0.000001 Then ws.Cells(i, j).NumberFormat = "#,##0.0#"
        End If
    Next
    Return

列を調べる:
    ' 列の種類: 見出しの語（電話・郵便・カナ・番号・メール）→ それ以外は値の割合（日付 6 割・数 6 割・@ 6 割・会社 3 割）
    kind = ""
    If InStr(hd, "電話") > 0 Or InStr(hd, "TEL") > 0 Or InStr(hd, "携帯") > 0 Or InStr(hd, "FAX") > 0 Then
        kind = "phone"
    ElseIf InStr(hd, "郵便") > 0 Or InStr(hd, "〒") > 0 Then
        kind = "postal"
    ElseIf InStr(hd, "ｶﾅ") > 0 Or InStr(hd, "ﾌﾘｶﾞﾅ") > 0 Or InStr(hd, "かな") > 0 Or InStr(hd, "ふりがな") > 0 Then
        kind = "kana"
    ElseIf InStr(hd, "番号") > 0 Or InStr(hd, "ｺｰﾄﾞ") > 0 Or InStr(hd, "品番") > 0 Or hd = "ID" Or hd = "NO" Or hd = "NO." Then
        kind = "code"
    ElseIf InStr(hd, "ﾒｰﾙ") > 0 Or InStr(hd, "MAIL") > 0 Then
        kind = "mail"
    End If
    cnt = 0: dn = 0: nn = 0: cn = 0: zs = 0: hs = 0: nA = 0: an = 0: pn = 0
    For i = hr + 1 To lastR
        v = ws.Cells(i, j).Value
        If Not isEmpty(v) And Not IsError(v) Then
            cnt = cnt + 1
            GoSub 日付を読む
            If ok Then
                dn = dn + 1
            ElseIf InStr(hd, "日") > 0 And VarType(v) = vbDouble Then
                If v > 20000 And v < 80000 And v = Int(v) Then dn = dn + 1     ' 日付の列で数値のままの日付
            End If
            GoSub 数を読む
            If ok Then nn = nn + 1
            If VarType(v) = vbString Then
                s = v
                t = Trim(s)
                If Right(t, 1) = "%" Or Right(t, 1) = "％" Then pn = pn + 1
                If InStr(s, "@") > 0 Or InStr(s, "＠") > 0 Then an = an + 1
                If InStr(s, "株式会社") > 0 Or InStr(s, "有限会社") > 0 Or InStr(s, "(株)") > 0 Or InStr(s, "（株）") > 0 _
                   Or InStr(s, "㈱") > 0 Or InStr(s, "(有)") > 0 Or InStr(s, "（有）") > 0 Or InStr(s, "㈲") > 0 Then cn = cn + 1
                t = s
                Do While Len(t) > 0
                    If Left(t, 1) = " " Or Left(t, 1) = "　" Or Left(t, 1) = vbTab Then t = Mid(t, 2) Else Exit Do
                Loop
                Do While Len(t) > 0
                    If Right(t, 1) = " " Or Right(t, 1) = "　" Or Right(t, 1) = vbTab Then t = Left(t, Len(t) - 1) Else Exit Do
                Loop
                If InStr(t, "　") > 0 Then zs = zs + 1
                If InStr(t, " ") > 0 Then hs = hs + 1
                If kind = "phone" Then
                    ' 同じ列のハイフン付きの例から、市外局番の区切りを覚える
                    t = StrConv(t, vbNarrow)
                    parts = Split(t, "-")
                    If UBound(parts) = 2 Then
                        u = parts(0) & parts(1) & parts(2)
                        If Left(t, 1) = "0" And Len(u) = 10 And Not u Like "*[!0-9]*" _
                           And Left(t, 4) <> "0120" And Left(t, 4) <> "0800" And Left(t, 4) <> "0570" Then
                            ok = False
                            For k = 1 To nA
                                If ar(k) = parts(0) Then ok = True
                            Next
                            If Not ok And nA < 200 Then
                                nA = nA + 1: ar(nA) = parts(0): al(nA) = Len(parts(1)): bl(nA) = Len(parts(2))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    If セルの字(kind) = "" Then
        If cnt > 0 And dn >= cnt * 0.6 Then
            kind = "date"
        ElseIf cnt > 0 And (pn >= cnt * 0.3 Or (InStr(hd, "率") > 0 And nn + pn >= cnt * 0.6)) Then
            kind = "pct"
        ElseIf cnt > 0 And nn >= cnt * 0.6 Then
            kind = "num"
        ElseIf cnt > 0 And an >= cnt * 0.6 Then
            kind = "mail"
        ElseIf cnt > 0 And cn >= cnt * 0.3 Then
            kind = "corp"
        Else
            kind = "text"
        End If
    End If
    If zs > hs Then sp = "　" Else sp = " "
    Return

値を直す:
    Select Case kind
        Case "date"
            GoSub 日付を読む
            If Not ok And VarType(v) = vbDouble Then
                ' 数値のままの日付（表示形式が標準に戻ったもの）
                If v > 20000 And v < 80000 And v = Int(v) Then dt = DateSerial(1899, 12, 30) + v: ok = True
            End If
            If ok Then
                nv = dt
            ElseIf VarType(v) = vbString Then
                s = v: GoSub 文字を直す: nv = s
            End If
        Case "num"
            GoSub 数を読む
            If ok Then
                nv = dbl
            ElseIf VarType(v) = vbString Then
                s = v: GoSub 文字を直す: nv = s
            End If
        Case "phone"
            If VarType(v) = vbString Then
                s = v: GoSub 電話を直す: nv = s
            End If
        Case "postal"
            If VarType(v) = vbString Then
                t = StrConv(Trim(Replace(v, "　", " ")), vbNarrow)
                digits = t
                For Each x In Array("〒", " ", "-", "ｰ", ChrW(&H2010), ChrW(&H2212))
                    digits = Replace(digits, x, "")
                Next
                If Len(digits) = 7 And Not digits Like "*[!0-9]*" Then
                    nv = Left(digits, 3) & "-" & Right(digits, 4)
                Else
                    nv = t
                End If
            ElseIf VarType(v) = vbDouble Then
                ' 数になった郵便番号（先頭の 0 が消えた）は 7 桁に戻す
                If v >= 0 And v < 10000000 And v = Int(v) Then
                    digits = Format(v, "0000000")
                    nv = Left(digits, 3) & "-" & Right(digits, 4)
                End If
            End If
        Case "pct"
            If VarType(v) = vbString Then
                t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
                If Right(t, 1) = "%" Then
                    t = Left(t, Len(t) - 1)
                    If IsNumeric(t) And Not t Like "*[!0-9.+-]*" Then nv = CDbl(t) / 100
                ElseIf IsNumeric(t) And Not t Like "*[!0-9.+-]*" And Len(t) > 0 Then
                    nv = CDbl(t)
                End If
            End If
        Case "mail"
            If VarType(v) = vbString Then
                nv = LCase(Replace(Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", ""), "＠", "@"))
            End If
        Case Else
            If VarType(v) = vbString Then
                s = v: GoSub 文字を直す: nv = s
            End If
    End Select
    Return

書く:
    If VarType(nv) = vbString Then
        If kind = "phone" Or kind = "postal" Or kind = "code" Or IsNumeric(nv) Or IsDate(nv) Then c.NumberFormat = "@"
    ElseIf c.NumberFormat = "@" Then
        c.NumberFormat = "General"
    End If
    c.Value = nv
    Return

文字を直す:
    ' セルの中の改行は空白に。全角の英数字・記号 → 半角（全角空白はあとで）／半角カナ → 全角（濁点も合わせる）
    s = Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " ")
    u = "": kanaRun = ""
    For k = 1 To Len(s)
        ch = Mid(s, k, 1)
        cd = AscW(ch): If cd < 0 Then cd = cd + 65536
        If cd >= &HFF61& And cd <= &HFF9F& Then
            kanaRun = kanaRun & ch
        Else
            If Len(kanaRun) > 0 Then u = u & StrConv(kanaRun, vbWide): kanaRun = ""
            If cd >= &HFF01& And cd <= &HFF5E& Then u = u & ChrW(cd - &HFEE0&) Else u = u & ch
        End If
    Next
    If Len(kanaRun) > 0 Then u = u & StrConv(kanaRun, vbWide)
    s = u
    If kind = "kana" Then s = StrConv(s, vbKatakana)          ' フリガナのひらがな → カタカナ
    ' カタカナの後ろの - ― ‐ － は長音「ー」（コピ-用紙・テ-プ・トナ‐TN。後ろが数字なら区切りのまま＝ルーム-2）
    u = ""
    For k = 1 To Len(s)
        ch = Mid(s, k, 1)
        If (ch = "-" Or ch = ChrW(&H2015) Or ch = ChrW(&H2010) Or ch = ChrW(&H2212)) And k > 1 Then
            cd = AscW(Mid(s, k - 1, 1)): If cd < 0 Then cd = cd + 65536
            If cd >= &H30A1& And cd <= &H30FA& Then
                If k = Len(s) Then
                    ch = "ー"
                ElseIf Not Mid(s, k + 1, 1) Like "#" Then
                    ch = "ー"
                End If
            End If
        End If
        u = u & ch
    Next
    s = u
    ' 前後の空白を取る
    Do While Len(s) > 0
        If Left(s, 1) = " " Or Left(s, 1) = "　" Or Left(s, 1) = vbTab Then s = Mid(s, 2) Else Exit Do
    Loop
    Do While Len(s) > 0
        If Right(s, 1) = " " Or Right(s, 1) = "　" Or Right(s, 1) = vbTab Then s = Left(s, Len(s) - 1) Else Exit Do
    Loop
    ' 文字の間の空白は、列の中で多い方の 1 つに
    s = Replace(s, "　", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If kind = "corp" Then
        ' 会社の種類の略は正式の名前に・種類の前後の空白は取る（前株・後株の位置は変えない）
        s = Replace(Replace(Replace(Replace(s, "(株)", "株式会社"), "(有)", "有限会社"), "㈱", "株式会社"), "㈲", "有限会社")
        s = Replace(Replace(s, "株式会社 ", "株式会社"), " 株式会社", "株式会社")
        s = Replace(Replace(s, "有限会社 ", "有限会社"), " 有限会社", "有限会社")
    End If
    If sp = "　" Then s = Replace(s, " ", "　")
    Return

言い換えをそろえる:
    ' 状態・区分のような列（同じ値が並ぶ・種類 15 以下）で、送り仮名の有無（受注済／受注済み）と
    ' ひらがな書き（しょうだん中／商談中＝Excel のふりがなが同じ）を、ひらがなの少ない形 → 多く使われている形にそろえる。
    ' 英語の状態語（Active 等）は、決まった対応表で、列に既にある日本語（進行中 等）へそろえる（無ければ触らない）
    Set dv = CreateObject("Scripting.Dictionary")
    cnt = 0
    For i = hr + 1 To lastR
        v = ws.Cells(i, j).Value
        If VarType(v) = vbString Then
            If Len(v) > 0 Then
                cnt = cnt + 1
                dv(v) = dv(v) + 1
            End If
        End If
    Next
    If cnt < 5 Or dv.Count < 2 Or dv.Count > 15 Or dv.Count * 2 > cnt Then Return
    ks = dv.keys
    n = dv.Count
    ReDim grp(0 To n - 1): ReDim k1(0 To n - 1): ReDim k2(0 To n - 1): ReDim rep(0 To n - 1)
    ReDim hira(0 To n - 1): ReDim cntOf(0 To n - 1): ReDim hasKana(0 To n - 1): ReDim hasKanji(0 To n - 1)
    For a = 0 To n - 1
        s = ks(a)
        grp(a) = a
        cntOf(a) = dv(s)
        hasKana(a) = False: hasKanji(a) = False
        For k = 1 To Len(s)
            cd = AscW(Mid(s, k, 1)): If cd < 0 Then cd = cd + 65536
            If cd >= &H3041& And cd <= &H30FF& Then hasKana(a) = True
            If cd >= &H4E00& And cd <= &H9FFF& Then hasKanji(a) = True
        Next
        ' 送り仮名を落とした形（末尾のひらがな）と、その中のひらがなの数
        t = s
        Do While Len(t) > 1
            cd = AscW(Right(t, 1)): If cd < 0 Then cd = cd + 65536
            If cd >= &H3041& And cd <= &H309F& Then t = Left(t, Len(t) - 1) Else Exit Do
        Loop
        hira(a) = 0
        For k = 1 To Len(t)
            cd = AscW(Mid(t, k, 1)): If cd < 0 Then cd = cd + 65536
            If cd >= &H3041& And cd <= &H309F& Then hira(a) = hira(a) + 1
        Next
        If hasKanji(a) Then k1(a) = t Else k1(a) = ""
        ' 読み（Excel のふりがな）
        k2(a) = ""
        If hasKana(a) Or hasKanji(a) Then
            On Error Resume Next
            k2(a) = StrConv(Application.GetPhonetic(s), vbKatakana Or vbWide)
            On Error GoTo 0
        End If
    Next
    For a = 0 To n - 1
        For b = a + 1 To n - 1
            ok = False
            If Len(k1(a)) > 0 And k1(a) = k1(b) Then ok = True
            If Len(k2(a)) > 0 And k2(a) = k2(b) And (hasKana(a) Or hasKana(b)) And (hasKanji(a) Or hasKanji(b)) Then ok = True
            If ok And grp(a) <> grp(b) Then
                g = grp(b)
                For k = 0 To n - 1
                    If grp(k) = g Then grp(k) = grp(a)
                Next
            End If
        Next
    Next
    For a = 0 To n - 1
        best = -1
        For b = 0 To n - 1
            If grp(b) = grp(a) Then
                If best = -1 Then
                    best = b
                ElseIf hira(b) < hira(best) Then
                    best = b
                ElseIf hira(b) = hira(best) Then
                    If cntOf(b) > cntOf(best) Or (cntOf(b) = cntOf(best) And Len(ks(b)) > Len(ks(best))) Then best = b
                End If
            End If
        Next
        rep(a) = ks(best)
    Next
    ' 英語の状態語 → 列に既にある日本語（その日本語が列に無ければ触らない）
    For a = 0 To n - 1
        s = LCase(Trim(ks(a)))
        If Len(s) > 0 And Not s Like "*[!a-z ]*" Then
            t = ""
            Select Case s
                Case "active", "in progress", "ongoing", "open", "wip": t = "進行中|対応中"
                Case "pending", "on hold", "hold": t = "保留|保留中"
                Case "won", "closed won", "ordered", "order": t = "受注済み|受注済|受注"
                Case "lost", "closed lost": t = "失注"
                Case "done", "completed", "complete", "closed", "finished": t = "完了|完了済み|済"
                Case "negotiating", "negotiation", "in negotiation": t = "商談中|交渉中"
                Case "consulting", "inquiry", "considering": t = "相談中|検討中"
                Case "new": t = "新規"
                Case "cancel", "cancelled", "canceled": t = "キャンセル|取消"
            End Select
            If Len(t) > 0 Then
                For Each x In Split(t, "|")
                    For b = 0 To n - 1
                        If rep(a) = ks(a) And rep(b) = x Then rep(a) = x
                    Next
                Next
            End If
        End If
    Next
    For i = hr + 1 To lastR
        Set c = ws.Cells(i, j)
        v = c.Value
        If VarType(v) = vbString And Not c.HasFormula Then
            For a = 0 To n - 1
                If ks(a) = v Then
                    If rep(a) <> v Then c.Value = rep(a)
                    Exit For
                End If
            Next
        End If
    Next
    Return

電話を直す:
    ' 03-1234-5678 の形に。市外局番の区切りは、+81 の区切り → 同じ列の例 → 携帯・0120 → 03・06・011・045・052・075・078・092
    ' の順で決める。決まらない番号は区切らない（当て推量で切らない）
    t = StrConv(Trim(Replace(s, "　", " ")), vbNarrow)
    s = t
    If Left(t, 3) = "+81" Then
        u = Replace(Replace(Replace(Replace(Mid(t, 4), "(", " "), ")", " "), "-", " "), ".", " ")
        u = Trim(u)
        Do While InStr(u, "  ") > 0
            u = Replace(u, "  ", " ")
        Loop
        parts = Split(u, " ")
        If UBound(parts) = 2 Then
            If Not (parts(0) & parts(1) & parts(2)) Like "*[!0-9]*" And Len(parts(0)) > 0 And Len(parts(1)) > 0 And Len(parts(2)) > 0 Then
                s = "0" & parts(0) & "-" & parts(1) & "-" & parts(2)
                Return
            End If
        End If
    End If
    digits = t
    For Each x In Array(" ", "-", "(", ")", "ｰ", ChrW(&H2010), ChrW(&H2212), ".")
        digits = Replace(digits, x, "")
    Next
    If Left(digits, 3) = "+81" Then digits = "0" & Mid(digits, 4)
    If Len(digits) = 0 Or digits Like "*[!0-9]*" Or Left(digits, 1) <> "0" Then Return
    If Len(digits) = 11 Then
        Select Case Left(digits, 3)
            Case "090", "080", "070", "060", "050"
                s = Left(digits, 3) & "-" & Mid(digits, 4, 4) & "-" & Right(digits, 4)
        End Select
        Return
    End If
    If Len(digits) <> 10 Then Return
    Select Case Left(digits, 4)
        Case "0120", "0800", "0570", "0990"
            s = Left(digits, 4) & "-" & Mid(digits, 5, 3) & "-" & Right(digits, 3)
            Return
    End Select
    ok = False
    For ln = 5 To 2 Step -1
        For k = 1 To nA
            If Not ok And Len(ar(k)) = ln And Left(digits, ln) = ar(k) And ln + al(k) + bl(k) = 10 Then
                s = Left(digits, ln) & "-" & Mid(digits, ln + 1, al(k)) & "-" & Right(digits, bl(k))
                ok = True
            End If
        Next
    Next
    If ok Then Return
    If (Left(digits, 2) = "03" Or Left(digits, 2) = "06") And Mid(digits, 3, 1) <> "0" Then
        s = Left(digits, 2) & "-" & Mid(digits, 3, 4) & "-" & Right(digits, 4)
        Return
    End If
    Select Case Left(digits, 3)
        Case "011", "045", "052", "075", "078", "092"
            If Mid(digits, 4, 1) <> "0" Then s = Left(digits, 3) & "-" & Mid(digits, 4, 3) & "-" & Right(digits, 4)
    End Select
    Return

日付を読む:
    ' 日付の値、または 2026/9/1・2026-09-01・2026年9月1日・令和8年9月1日・R8.9.1・9/1/2026・March 15, 2026・15 Mar 2026 → dt
    ok = False
    If VarType(v) = vbDate Then
        dt = v: ok = True
        Return
    End If
    If VarType(v) <> vbString Then Return
    ' 英語の日付
    u = LCase(Trim(Replace(Replace(StrConv(v, vbNarrow), ",", " "), ".", " ")))
    Do While InStr(u, "  ") > 0
        u = Replace(u, "  ", " ")
    Loop
    parts = Split(u, " ")
    If UBound(parts) = 2 Then
        mon = "": t = ""
        If parts(0) Like "[a-z][a-z][a-z]*" Then
            mon = parts(0): t = parts(1)
        ElseIf parts(1) Like "[a-z][a-z][a-z]*" Then
            mon = parts(1): t = parts(0)
        End If
        If Len(mon) > 0 Then
            k = InStr("janfebmaraprmayjunjulaugsepoctnovdec", Left(mon, 3))
            Do While Len(t) > 0
                If Right(t, 1) Like "#" Then Exit Do Else t = Left(t, Len(t) - 1)
            Loop
            If k > 0 And (k Mod 3) = 1 And Len(t) > 0 And Len(t) <= 2 And Not t Like "*[!0-9]*" _
               And Len(parts(2)) = 4 And Not parts(2) Like "*[!0-9]*" Then
                m = (k + 2) \ 3: d = CLng(t): y = CLng(parts(2))
                If d >= 1 And d <= 31 Then
                    dt = DateSerial(y, m, d)
                    If Month(dt) = m And Day(dt) = d Then ok = True
                End If
            End If
            Return
        End If
    End If
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    If Len(t) = 8 And Not t Like "*[!0-9]*" And (Left(t, 2) = "19" Or Left(t, 2) = "20") Then
        y = CLng(Left(t, 4)): m = CLng(Mid(t, 5, 2)): d = CLng(Right(t, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            dt = DateSerial(y, m, d)
            If Month(dt) = m And Day(dt) = d Then ok = True
        End If
        Return
    End If
    era = 0
    If Left(t, 2) = "令和" Then
        era = 2018: t = Mid(t, 3)
    ElseIf Left(t, 2) = "平成" Then
        era = 1988: t = Mid(t, 3)
    ElseIf Left(t, 2) = "昭和" Then
        era = 1925: t = Mid(t, 3)
    ElseIf t Like "[RHSrhs]#*" Then
        Select Case UCase(Left(t, 1))
            Case "R": era = 2018
            Case "H": era = 1988
            Case "S": era = 1925
        End Select
        t = Mid(t, 2)
    End If
    If era > 0 And Left(t, 1) = "元" Then t = "1" & Mid(t, 2)
    t = Replace(Replace(Replace(Replace(Replace(t, "年", "/"), "月", "/"), "日", ""), "-", "/"), ".", "/")
    parts = Split(t, "/")
    If UBound(parts) <> 2 Then Return
    bad = False
    For k = 0 To 2
        If Len(parts(k)) = 0 Or parts(k) Like "*[!0-9]*" Or Len(parts(k)) > 4 Then bad = True
    Next
    If bad Then Return
    If era > 0 Then
        y = era + CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(0)) = 4 Then
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        y = CLng(parts(2)): m = CLng(parts(0)): d = CLng(parts(1))
    Else
        Return
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or y < 1900 Or y > 9999 Then Return
    dt = DateSerial(y, m, d)
    If Month(dt) = m And Day(dt) = d Then ok = True
    Return

数を読む:
    ' 数の値、または ￥12,000・12000円・１２０００・12個・95万 の文字 → dbl（先頭が 0 の番号は数にしない）
    ok = False
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            dbl = CDbl(v): ok = True
            Return
    End Select
    If VarType(v) <> vbString Then Return
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    For Each x In Array(ChrW(&HA5), "\", "$", ",")
        t = Replace(t, x, "")
    Next
    ' 会計の負の数（△1,000・▲1,000・(1,000)）
    bad = False
    If Left(t, 1) = "△" Or Left(t, 1) = "▲" Then
        bad = True: t = Mid(t, 2)
    ElseIf Left(t, 1) = "(" And Right(t, 1) = ")" And Len(t) > 2 Then
        bad = True: t = Mid(t, 2, Len(t) - 2)
    End If
    Do While Len(t) > 0
        If InStr("円個本枚台冊箱点件", Right(t, 1)) > 0 Then t = Left(t, Len(t) - 1) Else Exit Do
    Loop
    u = ""
    If Right(t, 1) = "万" Then u = "万": t = Left(t, Len(t) - 1)
    If Len(t) = 0 Then Return
    If t Like "*[!0-9.+-]*" Then Return
    If InStr(2, t, "-") > 0 Or InStr(2, t, "+") > 0 Then Return
    If Len(t) > 1 And Left(t, 1) = "0" And Mid(t, 2, 1) <> "." Then Return
    If Len(Replace(Replace(t, ".", ""), "-", "")) > 15 Then Return
    If Not IsNumeric(t) Then Return
    dbl = CDbl(t)
    If u = "万" Then dbl = dbl * 10000
    If bad Then dbl = -dbl
    ok = True
    Return
End Sub

Sub 全列が同じ重複行を削除する()
    ' 依頼の語: 重複行を消|同じ行が二つ|同じ行が2つ|ダブっている行|ダブった行|重複データ|重複した行を1つ|重複した行をひとつ|二重に入っている行|重複行を削除|重複を消|重複を削除|重複を取り除|重複を除|重複データを消|重複をなくし|重複を無くし|ダブりを消|同じ行を消|同じ行を削除|重複している行を消|重複している行を削除|重複した行を消|重複した行を削除|重複行を取
    ' 依頼の組: 重複,ダブ,二重,かぶ,同じデータ,同じ行+消,削除,除,取り,なくし,無くし,1つに,ひとつに+-番号,キー,伝票,一覧にし,洗い出,マクロ,二重計上,二重払
    Dim ws As Worksheet, ur As Range, seen As Object
    Dim r0 As Long, c0 As Long, nR As Long, nC As Long, scanTo As Long
    Dim hr As Long, hc1 As Long, hc2 As Long, lastR As Long
    Dim i As Long, j As Long, k As Long, n As Long, best As Long
    Dim v As Variant, x As Variant, parts As Variant, dup() As Boolean
    Dim s As String, t As String, u As String, key As String, mon As String
    Dim ok As Boolean, bad As Boolean, dt As Date, dbl As Double
    Dim era As Long, y As Long, m As Long, d As Long

    ' 重複＝全部の列が同じ行（書き方の違い＝空白・全角半角・大文字小文字・区切り・（株）・文字の日付と金額はならして比べる）。
    ' 1 列でも中身が違えば別の行として残す（備考だけ違う行も消さない）。先に出てくる行を残し、後の行を下から消す
    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    r0 = ur.Row: c0 = ur.Column: nR = ur.rows.Count: nC = ur.Columns.Count
    If nR < 3 Then Exit Sub
    GoSub 見出しを探す
    If hr = 0 Or lastR <= hr + 1 Then Exit Sub
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim dup(hr + 1 To lastR)
    For i = hr + 1 To lastR
        key = ""
        For j = hc1 To hc2
            v = ws.Cells(i, j).Value
            GoSub 照合の値
            key = key & Chr(1) & s
        Next
        If Len(Replace(key, Chr(1), "")) = 0 Then
            ' 途中の空行は重複と見ない（消さない）
        ElseIf seen.Exists(key) Then
            dup(i) = True
        Else
            seen.Add key, i
        End If
    Next
    ' 行ごと消す（隣の列も一緒に上がる＝行とずれない）。表の外の同じ行に値があるときは、
    ' 消すとその値を巻き込み、表の中だけ詰めると隣の列が行とずれる＝その行は消さずに残す
    For i = lastR To hr + 1 Step -1
        If dup(i) Then
            n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nC - 1))) _
                - Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2)))
            If n = 0 Then ws.rows(i).Delete
        End If
    Next
    Exit Sub

見出しを探す:
    scanTo = r0 + IIf(nR > 20, 20, nR) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nC - 1)))
        If n > best Then best = n
    Next
    hr = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nC - 1)))
        If hr = 0 And n >= 2 And n >= best * 0.6 Then hr = i
    Next
    If hr = 0 Then Return
    hc1 = 0: hc2 = 0
    For j = c0 To c0 + nC - 1
        ' 二段見出しの下の段では、コード・品目のように上の段と縦に結合した列も表の列（2026-09-24 通しの実測 5:
        ' 下の段＝月の列だけを表と見て、コード・年計の列を「表の外」と数え、重複の行を残していた）
        ok = Len(Trim(ws.Cells(hr, j).text)) > 0
        If Not ok And ws.Cells(hr, j).MergeCells Then ok = (ws.Cells(hr, j).MergeArea.Row < hr)
        If ok Then
            If hc1 = 0 Then hc1 = j
            hc2 = j
        End If
    Next
    ' 本文の終わり: 途中の空行 1 行は飛ばして続ける。空行が 2 行続いたら表の終わり
    lastR = hr
    n = 0
    For i = hr + 1 To r0 + nR - 1
        If n < 2 Then
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2))) = 0 Then
                n = n + 1
            Else
                n = 0
                lastR = i
            End If
        End If
    Next
    Return

照合の値:
    ' 書き方の違いを無視した値（日付は年月日・数は数・文字は半角小文字で、空白・改行・区切りを落とし、会社の種類の略は正式の名前に）
    s = ""
    If isEmpty(v) Or IsError(v) Then Return
    If VarType(v) = vbDate Then
        s = "D" & Format(v, "yyyymmdd")
        Return
    End If
    GoSub 数を読む
    If ok Then
        s = CStr(dbl)
        Return
    End If
    If VarType(v) <> vbString Then
        s = CStr(v)
        Return
    End If
    GoSub 日付を読む
    If ok Then
        s = "D" & Format(dt, "yyyymmdd")
        Return
    End If
    ' 会社の種類の略は正式の名前に戻して比べる（前株・後株の位置は区別する＝株式会社A と A株式会社 は別）
    t = LCase(StrConv(StrConv(v, vbKatakana), vbNarrow))
    t = Replace(Replace(Replace(Replace(t, "(株)", "株式会社"), "(有)", "有限会社"), "㈱", "株式会社"), "㈲", "有限会社")
    For Each x In Array(" ", "　", vbTab, vbCr, vbLf, "-", "ｰ", "ー", ChrW(&H2010), ChrW(&H2015), ChrW(&H2212), ",", "(", ")", "〒", ChrW(&HA5), "\", "円", "個")
        t = Replace(t, x, "")
    Next
    If Left(t, 3) = "+81" Then t = "0" & Mid(t, 4)
    s = t
    Return

日付を読む:
    ok = False
    If VarType(v) = vbDate Then
        dt = v: ok = True
        Return
    End If
    If VarType(v) <> vbString Then Return
    u = LCase(Trim(Replace(Replace(StrConv(v, vbNarrow), ",", " "), ".", " ")))
    Do While InStr(u, "  ") > 0
        u = Replace(u, "  ", " ")
    Loop
    parts = Split(u, " ")
    If UBound(parts) = 2 Then
        mon = "": t = ""
        If parts(0) Like "[a-z][a-z][a-z]*" Then
            mon = parts(0): t = parts(1)
        ElseIf parts(1) Like "[a-z][a-z][a-z]*" Then
            mon = parts(1): t = parts(0)
        End If
        If Len(mon) > 0 Then
            k = InStr("janfebmaraprmayjunjulaugsepoctnovdec", Left(mon, 3))
            Do While Len(t) > 0
                If Right(t, 1) Like "#" Then Exit Do Else t = Left(t, Len(t) - 1)
            Loop
            If k > 0 And (k Mod 3) = 1 And Len(t) > 0 And Len(t) <= 2 And Not t Like "*[!0-9]*" _
               And Len(parts(2)) = 4 And Not parts(2) Like "*[!0-9]*" Then
                m = (k + 2) \ 3: d = CLng(t): y = CLng(parts(2))
                If d >= 1 And d <= 31 Then
                    dt = DateSerial(y, m, d)
                    If Month(dt) = m And Day(dt) = d Then ok = True
                End If
            End If
            Return
        End If
    End If
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    If Len(t) = 8 And Not t Like "*[!0-9]*" And (Left(t, 2) = "19" Or Left(t, 2) = "20") Then
        y = CLng(Left(t, 4)): m = CLng(Mid(t, 5, 2)): d = CLng(Right(t, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            dt = DateSerial(y, m, d)
            If Month(dt) = m And Day(dt) = d Then ok = True
        End If
        Return
    End If
    era = 0
    If Left(t, 2) = "令和" Then
        era = 2018: t = Mid(t, 3)
    ElseIf Left(t, 2) = "平成" Then
        era = 1988: t = Mid(t, 3)
    ElseIf Left(t, 2) = "昭和" Then
        era = 1925: t = Mid(t, 3)
    ElseIf t Like "[RHSrhs]#*" Then
        Select Case UCase(Left(t, 1))
            Case "R": era = 2018
            Case "H": era = 1988
            Case "S": era = 1925
        End Select
        t = Mid(t, 2)
    End If
    If era > 0 And Left(t, 1) = "元" Then t = "1" & Mid(t, 2)
    t = Replace(Replace(Replace(Replace(Replace(t, "年", "/"), "月", "/"), "日", ""), "-", "/"), ".", "/")
    parts = Split(t, "/")
    If UBound(parts) <> 2 Then Return
    bad = False
    For k = 0 To 2
        If Len(parts(k)) = 0 Or parts(k) Like "*[!0-9]*" Or Len(parts(k)) > 4 Then bad = True
    Next
    If bad Then Return
    If era > 0 Then
        y = era + CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(0)) = 4 Then
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        y = CLng(parts(2)): m = CLng(parts(0)): d = CLng(parts(1))
    Else
        Return
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or y < 1900 Or y > 9999 Then Return
    dt = DateSerial(y, m, d)
    If Month(dt) = m And Day(dt) = d Then ok = True
    Return

数を読む:
    ok = False
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            dbl = CDbl(v): ok = True
            Return
    End Select
    If VarType(v) <> vbString Then Return
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    For Each x In Array(ChrW(&HA5), "\", "$", ",")
        t = Replace(t, x, "")
    Next
    ' 会計の負の数（△1,000・▲1,000・(1,000)）
    bad = False
    If Left(t, 1) = "△" Or Left(t, 1) = "▲" Then
        bad = True: t = Mid(t, 2)
    ElseIf Left(t, 1) = "(" And Right(t, 1) = ")" And Len(t) > 2 Then
        bad = True: t = Mid(t, 2, Len(t) - 2)
    End If
    Do While Len(t) > 0
        If InStr("円個本枚台冊箱点件", Right(t, 1)) > 0 Then t = Left(t, Len(t) - 1) Else Exit Do
    Loop
    u = ""
    If Right(t, 1) = "万" Then u = "万": t = Left(t, Len(t) - 1)
    If Len(t) = 0 Then Return
    If t Like "*[!0-9.+-]*" Then Return
    If InStr(2, t, "-") > 0 Or InStr(2, t, "+") > 0 Then Return
    If Len(t) > 1 And Left(t, 1) = "0" And Mid(t, 2, 1) <> "." Then Return
    If Len(Replace(Replace(t, ".", ""), "-", "")) > 15 Then Return
    If Not IsNumeric(t) Then Return
    dbl = CDbl(t)
    If u = "万" Then dbl = dbl * 10000
    If bad Then dbl = -dbl
    ok = True
    Return
End Sub

Sub 表の下に合計行を足す()
    ' 依頼の語: 合計行を足|下に合計|合計の行を追加|最終行に合計|合計の行を足|合計行を入|合計行を作|合計行を付|合計行を追加|列ごとの合計|縦計を|合計の行を入|合計の行をお願|最終行の下に列|下に出
    ' 依頼の組: 合計,縦計,総計+下,最終行,最後,末尾,縦の,行,欄+足,入れ,出,追加,付け,作,お願,ほし,欲し+-検算,正しい,合って,合うか,小計,区分,項目,月,ピボット,テーブル,右,した表,できない
    ' 扱う: 合計 集計
    ' 見出し: なし
    ' 形: 数の列
    ' 表の下に合計行を追加する。数値列（番号・No・コード・年度・率を除く）にSUM式、左端文字列列に「合計」と書く。

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim urTop As Long, urLeft As Long, urBottom As Long, urRight As Long
    urTop = ur.Row
    urLeft = ur.Column
    urBottom = 表の本文の最終行(ws)
    urRight = ur.Column + ur.Columns.Count - 1

    Dim headerRow As Long
    headerRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    ' 二段見出し（上期・下期の下に 4月・5月…、コード・品目は縦に結合）は下の段を見出しにする
    ' （2026-09-24 通しの実測 5: 上の段を見出しと見て、下の段の月の名前から足していた）
    If headerRow + 1 <= urBottom Then
        Dim h2Merge As Boolean, h2Num As Boolean, h2Cnt As Long
        For c = urLeft To urRight
            If ws.Cells(headerRow + 1, c).MergeCells Then
                If ws.Cells(headerRow + 1, c).MergeArea.Row = headerRow Then h2Merge = True
            ElseIf セルの字(ws.Cells(headerRow + 1, c).Value) <> "" Then
                h2Cnt = h2Cnt + 1
                If IsNumeric(ws.Cells(headerRow + 1, c).Value) Or IsDate(ws.Cells(headerRow + 1, c).Value) Then h2Num = True
            End If
        Next c
        If h2Merge And h2Cnt >= 2 And Not h2Num Then headerRow = headerRow + 1
    End If

    Dim dataStart As Long
    dataStart = headerRow + 1
    If dataStart > urBottom Then Exit Sub

    Dim noteStartRow As Long
    noteStartRow = 0
    Dim rr As Long
    For rr = urBottom To dataStart Step -1
        Dim rv As String
        rv = ""
        Dim cc As Long
        For cc = urLeft To urRight
            If セルの字(ws.Cells(rr, cc).Value) <> "" Then
                rv = CStr(ws.Cells(rr, cc).Value)
                Exit For
            End If
        Next cc
        If Left(rv, 1) = "※" Then
            noteStartRow = rr
        Else
            Exit For
        End If
    Next rr

    Dim effectiveBottom As Long
    If noteStartRow > 0 Then
        effectiveBottom = noteStartRow - 1
    Else
        effectiveBottom = urBottom
    End If

    Dim sumRow As Long
    sumRow = 0
    Dim lastDataRow As Long
    lastDataRow = effectiveBottom

    Dim checkVal As String
    checkVal = ""
    For c = urLeft To urRight
        If セルの字(ws.Cells(effectiveBottom, c).Value) <> "" Then
            checkVal = CStr(ws.Cells(effectiveBottom, c).Value)
            Exit For
        End If
    Next c
    Dim s As String
    s = checkVal: GoSub 正規化: checkVal = s
    ' 最後の行が「財政課 小計」のような小計なら合計の行ではない（上書きして小計を消していた・2026-09-24 棚の試験 F9）
    If 集計の語か(checkVal) And Right$(Replace(checkVal, " ", ""), 2) <> "小計" Then
        sumRow = effectiveBottom
        lastDataRow = effectiveBottom - 1
    End If
    ' 本文の最終行のすぐ下に続く小計の行（最後の課の小計）も明細の側に入れる。入れないと、その小計の行を
    ' 「合計の行が無い所」と見て上書きしていた（2026-09-24 棚の試験 F9・実測4 の G19）
    Dim subRows As String, sr As Long, scn As Long, isSubR As Boolean
    Do While sumRow = 0 And lastDataRow + 1 <= ur.Row + ur.rows.Count - 1
        isSubR = False
        For scn = urLeft To urRight
            If VarType(ws.Cells(lastDataRow + 1, scn).Value) = vbString Then
                If 集計の語か(ws.Cells(lastDataRow + 1, scn).Value) And Right$(Replace(ws.Cells(lastDataRow + 1, scn).Value, " ", ""), 2) = "小計" Then isSubR = True: Exit For
            End If
        Next scn
        If Not isSubR Then Exit Do
        lastDataRow = lastDataRow + 1
    Loop
    ' 明細のあいだの小計の行（合計はこれらを足す＝明細と小計を両方足すと 2 倍になる）
    For sr = dataStart To lastDataRow
        For scn = urLeft To urRight
            If VarType(ws.Cells(sr, scn).Value) = vbString Then
                If 集計の語か(ws.Cells(sr, scn).Value) And Right$(Replace(ws.Cells(sr, scn).Value, " ", ""), 2) = "小計" Then
                    subRows = subRows & "," & sr
                    Exit For
                End If
            End If
        Next scn
    Next sr

    If sumRow = 0 Then sumRow = 表の下の合計行(ws, lastDataRow)   ' 表のすぐ下の合計の行は書き直す（2026-09-23）
    If sumRow = 0 Then
        If noteStartRow > 0 Then
            ws.rows(noteStartRow).Insert Shift:=xlDown
            sumRow = noteStartRow
        Else
            sumRow = lastDataRow + 1
        End If
    End If

    ' 見出しのキーワードで除外列を判定（全角のまま比較）
    Dim skipByHeader() As Boolean
    ReDim skipByHeader(urLeft To urRight)
    Dim hv As String
    For c = urLeft To urRight
        s = CStr(ws.Cells(headerRow, c).MergeArea.Cells(1, 1).Value)   ' 縦に結合した見出し（コード・前年比）は上のセルの名前
        ' 全角英字を半角にして比較するが、カタカナはそのまま
        Dim hs As String
        hs = s
        ' 全角英数のみ半角化（個別Replace）
        hs = Replace(hs, "Ａ", "A"): hs = Replace(hs, "Ｂ", "B"): hs = Replace(hs, "Ｃ", "C")
        hs = Replace(hs, "Ｎ", "N"): hs = Replace(hs, "Ｏ", "O"): hs = Replace(hs, "ｏ", "o")
        hs = Replace(hs, "ａ", "a"): hs = Replace(hs, "ｂ", "b"): hs = Replace(hs, "ｃ", "c")
        hs = Trim(hs)
        If InStr(hs, "番号") > 0 Or InStr(hs, "NO") > 0 Or InStr(hs, "No") > 0 Or _
           InStr(hs, "コード") > 0 Or InStr(hs, "CODE") > 0 Or InStr(hs, "年度") > 0 Or InStr(hs, "単価") > 0 Or _
           Right$(hs, 1) = "比" Or Right$(hs, 1) = "率" Or InStr(hs, "割合") > 0 Then   ' 比・率は足しても意味が無い（前年比を合計していた・2026-09-24）
            skipByHeader(c) = True
        End If
    Next c

    Dim colType() As String
    ReDim colType(urLeft To urRight)

    For c = urLeft To urRight
        If skipByHeader(c) Then
            colType(c) = "skip"
        Else
            Dim fmt As String
            fmt = ws.Cells(dataStart, c).NumberFormat
            If InStr(fmt, "%") > 0 Then
                colType(c) = "rate"
            Else
                Dim numCount As Long, dateCount As Long, totalCount As Long
                numCount = 0: dateCount = 0: totalCount = 0
                For r = dataStart To lastDataRow
                    Dim cv As Variant
                    cv = ws.Cells(r, c).Value
                    If セルの字(cv) <> "" Then
                        totalCount = totalCount + 1
                        If IsDate(cv) And Not IsNumeric(cv) Then
                            dateCount = dateCount + 1
                        ElseIf IsNumeric(cv) Then
                            numCount = numCount + 1
                        End If
                    End If
                Next r
                If totalCount = 0 Then
                    colType(c) = "other"
                ElseIf dateCount > 0 And dateCount >= totalCount \ 2 Then
                    colType(c) = "date"
                ElseIf numCount > 0 And numCount >= totalCount \ 2 Then
                    Dim fmtD As String
                    fmtD = ws.Cells(dataStart, c).NumberFormat
                    If InStr(fmtD, "%") > 0 Then
                        colType(c) = "rate"
                    Else
                        Dim fracCount As Long
                        fracCount = 0
                        Dim fNumCount As Long
                        fNumCount = 0
                        For r = dataStart To lastDataRow
                            If セルの字(ws.Cells(r, c).Value) <> "" And IsNumeric(ws.Cells(r, c).Value) Then
                                Dim dv As Double
                                dv = CDbl(ws.Cells(r, c).Value)
                                fNumCount = fNumCount + 1
                                If dv >= 0 And dv <= 1 Then fracCount = fracCount + 1
                            End If
                        Next r
                        If fNumCount > 0 And fracCount = fNumCount Then
                            colType(c) = "rate"
                        Else
                            colType(c) = "num"
                        End If
                    End If
                Else
                    colType(c) = "other"
                End If
            End If
        End If
    Next c

    ' 番号列・連番列・年度列をskipに
    For c = urLeft To urRight
        If colType(c) = "num" Then
            Dim allInt As Boolean, allYear As Boolean
            allInt = True: allYear = True
            Dim minV As Double, maxV As Double, firstSet As Boolean
            firstSet = False
            For r = dataStart To lastDataRow
                If セルの字(ws.Cells(r, c).Value) <> "" Then
                    If Not IsNumeric(ws.Cells(r, c).Value) Then
                        allInt = False: allYear = False
                    Else
                        Dim nv As Double
                        nv = CDbl(ws.Cells(r, c).Value)
                        If nv <> Int(nv) Then allInt = False: allYear = False
                        If nv < 1900 Or nv > 2100 Then allYear = False
                        If Not firstSet Then
                            minV = nv: maxV = nv: firstSet = True
                        Else
                            If nv < minV Then minV = nv
                            If nv > maxV Then maxV = nv
                        End If
                    End If
                End If
            Next r
            If allYear And firstSet Then
                colType(c) = "other"
            ElseIf allInt And firstSet Then
                Dim rowSpan As Long
                rowSpan = lastDataRow - dataStart + 1
                ' 連番と見るのは上から順に増えていく列だけ（数量 10,5,8,3… を連番と見て合計から外した・2026-09-23）
                Dim upOK As Boolean, prevN As Double, gotPrev As Boolean
                upOK = True: gotPrev = False
                For r = dataStart To lastDataRow
                    If IsNumeric(ws.Cells(r, c).Value) And セルの字(ws.Cells(r, c).Value) <> "" Then
                        If gotPrev Then
                            If CDbl(ws.Cells(r, c).Value) < prevN Then upOK = False
                        End If
                        prevN = CDbl(ws.Cells(r, c).Value)
                        gotPrev = True
                    End If
                Next r
                If upOK And (maxV - minV) <= rowSpan * 2 And maxV - minV >= 0 Then
                    colType(c) = "seq"
                End If
            End If
        End If
    Next c

    ' 「合計」を書く列を決定
    ' skip列を飛ばして最初の列を確認、seqならその次のother列、それ以外は最初の非skip列
    Dim goukeCol As Long
    goukeCol = 0
    Dim firstNonSkip As Long
    firstNonSkip = 0
    For c = urLeft To urRight
        If colType(c) <> "skip" Then
            firstNonSkip = c
            Exit For
        End If
    Next c

    If firstNonSkip > 0 Then
        If colType(firstNonSkip) = "seq" Then
            For c = firstNonSkip + 1 To urRight
                If colType(c) = "other" Then
                    goukeCol = c
                    Exit For
                End If
            Next c
        ElseIf colType(firstNonSkip) = "other" Then
            goukeCol = firstNonSkip
        Else
            ' numやrateが先頭の場合は次のother列を探す
            For c = firstNonSkip To urRight
                If colType(c) = "other" Then
                    goukeCol = c
                    Exit For
                End If
            Next c
        End If
    End If

    ' 前からある合計の行は、「合計」の字の位置をそのまま使う（F 列の「合計」を C 列へ動かしていた・2026-09-24）
    For c = urLeft To urRight
        If VarType(ws.Cells(sumRow, c).Value) = vbString Then
            If 集計の語か(ws.Cells(sumRow, c).Value) And colType(c) <> "num" Then goukeCol = c: Exit For   ' 数の列の「合計」は式で消える
        End If
    Next c

    ' 合計行をクリア
    For c = urLeft To urRight
        ws.Cells(sumRow, c).ClearContents
    Next c

    If goukeCol > 0 Then
        ws.Cells(sumRow, goukeCol).Value = "合計"
    End If

    For c = urLeft To urRight
        If colType(c) = "num" Then
            If Len(subRows) > 0 Then
                ' 小計のある表は小計どうしを足す
                Dim subF As String, sv As Variant
                subF = ""
                For Each sv In Split(Mid$(subRows, 2), ",")
                    subF = subF & "+" & ws.Cells(CLng(sv), c).Address(False, False)
                Next sv
                ws.Cells(sumRow, c).Formula = "=" & Mid$(subF, 2)
            Else
                ws.Cells(sumRow, c).Formula = "=SUM(" & ws.Cells(dataStart, c).Address & ":" & ws.Cells(lastDataRow, c).Address & ")"
            End If
            ' 合計は明細より桁が増える＝列幅に収まらず ##### のまま渡さない（2026-09-23 の通しの実測）
            If Left$(ws.Cells(sumRow, c).text, 1) = "#" Then ws.Columns(c).AutoFit
        End If
    Next c

    Exit Sub

正規化:
    s = Trim$(s)
    Return

End Sub

Sub 選んだコード列を名称で上書きする()
    ' 依頼の語: 列のコード|コード列|コードを名称|コードを名前|対応表の名前|コードのまま|名称に置|替えて名称
    ' 依頼の組: コード+名称,名前+置き換,置換,変換,直,にして
    ' 扱う: コード名称置き換え 突き合
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet
    Dim selCol As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim s As String
    Dim cellVal As String
    Dim normalKey As String

    Set ws = ActiveSheet
    selCol = Selection.Cells(1, 1).Column
    headerRow = Selection.Cells(1, 1).Row
    lastRow = 表の本文の最終行(ws)

    ' サンプル値収集（見出し行の次から）
    Dim sampleVals As Object
    Set sampleVals = CreateObject("Scripting.Dictionary")
    sampleVals.CompareMode = 1
    Dim sampleCount As Long
    sampleCount = 0
    For i = headerRow + 1 To lastRow
        If セルの字(ws.Cells(i, selCol).Value) <> "" Then
            s = CStr(ws.Cells(i, selCol).Value): GoSub 正規化
            If Not sampleVals.Exists(s) Then
                sampleVals.Add s, 1
                sampleCount = sampleCount + 1
                If sampleCount >= 10 Then Exit For
            End If
        End If
    Next i

    ' 対応表シートを探す
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim bestSheet As Worksheet
    Dim bestKeyCol As Long
    Dim bestValCol As Long
    Dim bestMatch As Long
    bestMatch = 0
    Dim targetSheet As Worksheet

    For Each targetSheet In wb.Sheets
        If targetSheet.Name = ws.Name Then GoTo NextSheet
        Dim tUR As Range
        Set tUR = targetSheet.UsedRange
        If tUR.Columns.Count < 2 Then GoTo NextSheet
        Dim tFirstRow As Long, tLastRow As Long, tFirstCol As Long, tLastCol As Long
        tFirstRow = tUR.Row
        tLastRow = tUR.Row + tUR.rows.Count - 1
        tFirstCol = tUR.Column
        tLastCol = tUR.Column + tUR.Columns.Count - 1
        ' 対応表は2列のみ対象（左がコード・右が名称）
        If tLastCol - tFirstCol <> 1 Then GoTo NextSheet
        Dim c As Long
        c = tFirstCol
        Dim matchCount As Long
        matchCount = 0
        For i = tFirstRow To tLastRow
            s = CStr(targetSheet.Cells(i, c).Value): GoSub 正規化
            If sampleVals.Exists(s) Then matchCount = matchCount + 1
        Next i
        If matchCount > bestMatch Then
            bestMatch = matchCount
            Set bestSheet = targetSheet
            bestKeyCol = c
            bestValCol = c + 1
        End If
NextSheet:
    Next targetSheet

    If bestMatch = 0 Or bestSheet Is Nothing Then Exit Sub

    ' 対応表から辞書構築
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim tUR2 As Range
    Set tUR2 = bestSheet.UsedRange
    Dim tFR As Long, tLR As Long
    tFR = tUR2.Row
    tLR = tUR2.Row + tUR2.rows.Count - 1
    For i = tFR To tLR
        Dim kv As String, vv As String
        kv = CStr(bestSheet.Cells(i, bestKeyCol).Value)
        vv = CStr(bestSheet.Cells(i, bestValCol).Value)
        If セルの字(kv) <> "" And セルの字(vv) <> "" Then
            s = kv: GoSub 正規化
            If Not dict.Exists(s) Then dict.Add s, vv
        End If
    Next i

    ' 名称の逆引きセット（既に名称になっている値を判定）
    Dim nameSet As Object
    Set nameSet = CreateObject("Scripting.Dictionary")
    nameSet.CompareMode = 1
    Dim dk As Variant
    For Each dk In dict.keys
        Dim nv As String
        nv = CStr(dict(dk))
        If Not nameSet.Exists(nv) Then nameSet.Add nv, 1
    Next dk

    ' 置き換え（見出し行の次から）
    For i = headerRow + 1 To lastRow
        cellVal = CStr(ws.Cells(i, selCol).Value)
        If セルの字(cellVal) = "" Then GoTo NextCell
        If nameSet.Exists(cellVal) Then GoTo NextCell
        s = cellVal: GoSub 正規化
        normalKey = s
        If dict.Exists(normalKey) Then
            ws.Cells(i, selCol).Value = dict(normalKey)
        End If
NextCell:
    Next i

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, ",", "")
    Dim tmp As String
    tmp = s
    Do While Left$(tmp, 1) = "0" And Len(tmp) > 1
        tmp = Mid$(tmp, 2)
    Loop
    s = 数のキー(tmp)      ' CLng で 13 桁のコードが「オーバーフロー」で止まった（2026-09-24 総点検）
    Return

End Sub

Sub 結合セルをほどいて値を埋める()
    ' 依頼の語: 結合セルを解除|結合を外|結合をはずして|結合を解いて|結合をぜんぶ解除|マージを解除|結合セルをばら|全部解除
    ' 依頼の組: 結合+解除,外,解い,やめ,ばら+-一覧,調べ,クエリ,マスタ,キー
    ' 扱う: 結合解除 セル結合 マージ解除
    ' 見出し: なし
    ' 形: 数の列
    Dim ws As Worksheet
    Dim cell As Range
    Dim i As Long
    Dim cnt As Long
    Dim ur As Range

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    cnt = 0
    For Each cell In ur
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                cnt = cnt + 1
            End If
        End If
    Next cell

    If cnt = 0 Then Exit Sub

    Dim mergedAreas() As String
    Dim mergedValues() As Variant
    ReDim mergedAreas(1 To cnt)
    ReDim mergedValues(1 To cnt)

    i = 0
    For Each cell In ur
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                i = i + 1
                Dim ma As Range
                Set ma = cell.MergeArea
                ' 結合エリアが UsedRange の列範囲をまたいでいる場合は列 A のみに限定しない
                ' ただし表題行（1 行・複数列に広がる横結合）は値を左上セルのみに保持し他は空のまま
                ' 表題行の判定: 結合エリアが行全体にわたる（列数 > 1 かつ行数 = 1 かつ結合エリアの左端列 = UsedRange 左端列）
                ' → 表題行は解除後も左上のみ値を持ち、他は空にする
                Dim isTitle As Boolean
                isTitle = False
                If ma.rows.Count = 1 And ma.Columns.Count > 1 Then
                    If ma.Cells(1, 1).Column = ur.Cells(1, 1).Column Then
                        isTitle = True
                    End If
                End If
                mergedAreas(i) = ma.Address
                mergedValues(i) = cell.Value
                ' isTitle フラグを値配列に埋め込む（負の値で区別）
                If isTitle Then
                    mergedValues(i) = Array(cell.Value, True)
                Else
                    mergedValues(i) = Array(cell.Value, False)
                End If
            End If
        End If
    Next cell

    Dim mc As Range
    Dim val As Variant
    Dim fillAll As Boolean
    For i = 1 To cnt
        Set mc = ws.Range(mergedAreas(i))
        val = mergedValues(i)(0)
        fillAll = Not CBool(mergedValues(i)(1))
        mc.UnMerge
        Dim c As Range
        If fillAll Then
            For Each c In mc
                c.Value = val
            Next c
        Else
            mc.Cells(1, 1).Value = val
        End If
    Next i
End Sub

Sub 選んだ列の空白を上の値で埋める()
    ' 依頼の語: 空白を上|上の値|空欄を上|上から埋め|値で埋|上と同|値の繰|セルの値|空白と欠損の処理|欠損の処理
    ' 依頼の組: 空白,空欄,空い,省略+上+-行を,行が,列,結合
    ' 扱う: 空白埋め 上の値で埋める 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Dim sel As Range
    Dim area As Range
    Dim ur As Range
    Dim headerRow As Long
    Dim lastRow As Long
    Dim startRow As Long
    Dim r As Long
    Dim c As Long
    Dim aCol As Long
    Dim cc As Long
    Dim lastVal As Variant
    Dim cellVal As String
    Dim rowEmpty As Boolean
    Dim isTotal As Boolean
    Dim i As Integer
    Dim totalWords As Variant

    Set ws = ActiveSheet
    Set sel = Selection
    Set ur = ws.UsedRange
    ' 表の下のメモの行は埋めない（使用範囲の最終行まで見ていた・2026-09-24 総点検）
    lastRow = 表の本文の最終行(ws)
    headerRow = sel.Cells(1, 1).Row
    startRow = headerRow + 1

    For Each area In sel.Areas
        For aCol = 1 To area.Columns.Count
            c = area.Columns(aCol).Column
            lastVal = Empty

            For r = startRow To lastRow
                rowEmpty = True
                For cc = ur.Column To ur.Column + ur.Columns.Count - 1
                    If セルの字(ws.Cells(r, cc).Value) <> "" Then
                        rowEmpty = False
                        Exit For
                    End If
                Next cc
                If rowEmpty Then GoTo NextRow

                ' 合計・小計の行は飛ばす。「計」を含むだけの字（会計課・計画・計算機）は明細（2026-09-24 総点検: 行のどこかに
                ' 「計」があると合計の行と見て、会計課の行を埋めずに飛ばしていた）
                If 集計の行か(ws, r, ur.Column, ur.Column + ur.Columns.Count - 1) Then GoTo NextRow

                ' 上の値がエラー値なら埋めない（比べると「型が一致しません」で止まった・2026-09-24 総点検）
                If セルの字(ws.Cells(r, c).Value) = "" Then
                    If Not isEmpty(lastVal) Then
                        If Not IsError(lastVal) Then
                            If セルの字(lastVal) <> "" Then ws.Cells(r, c).Value = lastVal
                        End If
                    End If
                Else
                    lastVal = ws.Cells(r, c).Value
                End If

NextRow:
            Next r
        Next aCol
    Next area

End Sub

Sub 選んだ列の文字の日付を日付にする()
    ' 依頼の語: 文字の日付を|和暦の文字|日付として認識|受付日の書|入っている日付|日付列|日付が文字列|列を日付の値
    ' 依頼の組: 日付+文字,和暦,認識,日付型,日付値,日付に+-異常,おかしい,間違い,ミス,妥当,年度,四半期,未来,ありえない,あり得ない
    ' 扱う: 文字日付変換
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる列の文字列日付を日付値に変換し表示形式をyyyy/m/dにする

    Dim ws As Worksheet
    Dim cell As Range
    Dim s As String
    Dim parsed As Date
    Dim parseErr As Boolean
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim area As Range
    Dim col As Long
    Dim colIdx As Long
    Dim r As Long

    Set ws = ActiveSheet
    hdrRow = Selection.Cells(1, 1).Row
    lastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    For Each area In Selection.Areas
        For col = 1 To area.Columns.Count
            colIdx = area.Column + col - 1
            For r = hdrRow + 1 To lastRow
                Set cell = ws.Cells(r, colIdx)
                If isEmpty(cell.Value) Then GoTo NextCell
                If VarType(cell.Value) = vbDate Then GoTo NextCell
                If VarType(cell.Value) = vbDouble And InStr(cell.NumberFormat, "y") > 0 Then GoTo NextCell
                If VarType(cell.Value) <> vbString Then GoTo NextCell
                s = Trim(CStr(cell.Value))
                If セルの字(s) = "" Then GoTo NextCell
                s = StrConv(s, vbNarrow)
                s = Replace(s, "年", "/")
                s = Replace(s, "月", "/")
                s = Replace(s, "日", "")
                s = Replace(s, ".", "/")
                s = Trim(s)
                parseErr = True
                ' 20260407 の 8 桁
                If Len(s) = 8 And Not s Like "*[!0-9]*" Then
                    If CLng(Mid(s, 5, 2)) >= 1 And CLng(Mid(s, 5, 2)) <= 12 And CLng(Right(s, 2)) >= 1 And CLng(Right(s, 2)) <= 31 Then
                        parsed = DateSerial(CLng(Left(s, 4)), CLng(Mid(s, 5, 2)), CLng(Right(s, 2)))
                        parseErr = (Day(parsed) <> CLng(Right(s, 2)))
                    End If
                ' R8/4/6・H31/4/1・S64/1/7 の略した和暦
                ElseIf UCase(s) Like "[RHS]#*/*#/*#" Then
                    Dim 和 As Variant
                    和 = Split(Mid(s, 2), "/")
                    If UBound(和) = 2 Then
                        If Not (和(0) & 和(1) & 和(2)) Like "*[!0-9]*" Then
                            parsed = DateSerial(CLng(和(0)) + Choose(InStr("RHS", UCase(Left(s, 1))), 2018, 1988, 1925), CLng(和(1)), CLng(和(2)))
                            parseErr = (Day(parsed) <> CLng(和(2)))
                        End If
                    End If
                Else
                    On Error Resume Next
                    parsed = CDate(s)
                    parseErr = (Err.Number <> 0)
                    Err.Clear
                    On Error GoTo 0
                End If
                If Not parseErr Then
                    ' 文字の書式（@）のまま値を入れると文字の日付に戻るので、書式を先に変える
                    cell.NumberFormat = "yyyy/m/d"
                    cell.Value = parsed
                End If
NextCell:
            Next r
        Next col
    Next area
End Sub

Sub 選んだ列の文字の数字を数にする()
    ' 依頼の語: 数値に直|文字列の数字|数値に変換|金額が文字列|文字の数字を数値|計算できる数値|数に直|数値に変|文字を数値|カンマ付
    ' 依頼の組: 数字,数値,金額+文字+-日付,CSV
    ' 扱う: 文字数字変換 数値変換
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし

    Dim ws As Worksheet
    Dim sel As Range
    Dim area As Range
    Dim colIdx As Long
    Dim r As Long
    Dim lastR As Long
    Dim headerRow As Long
    Dim cell As Range
    Dim s As String
    Dim CHK As String
    Dim v As Double
    Dim neg As Boolean
    Dim allDigit As Boolean
    Dim ci As Integer
    Dim ch As String
    Dim fc As String

    Set ws = ActiveSheet
    Set sel = Selection
    headerRow = sel.Cells(1, 1).Row

    For Each area In sel.Areas
      For colIdx = area.Column To area.Column + area.Columns.Count - 1
        lastR = ws.Cells(ws.rows.Count, colIdx).End(xlUp).Row
        For r = headerRow + 1 To lastR
            Set cell = ws.Cells(r, colIdx)
            If isEmpty(cell.Value) Then GoTo NextCell
            Select Case VarType(cell.Value)
                Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal
                    GoTo NextCell
            End Select
            If VarType(cell.Value) <> vbString Then GoTo NextCell

            s = CStr(cell.Value): GoSub 正規化

            If セルの字(s) = "" Then GoTo NextCell
            If s = "-" Then GoTo NextCell

            CHK = s
            If Left(CHK, 1) = "-" Then CHK = Mid(CHK, 2)
            If Len(CHK) >= 2 And Left(CHK, 1) = "0" Then
                allDigit = True
                For ci = 1 To Len(CHK)
                    If Mid(CHK, ci, 1) < "0" Or Mid(CHK, ci, 1) > "9" Then
                        allDigit = False: Exit For
                    End If
                Next ci
                If allDigit Then GoTo NextCell
            End If

            CHK = s
            If Left(CHK, 1) = "-" Then CHK = Mid(CHK, 2)
            If セルの字(CHK) = "" Then GoTo NextCell
            allDigit = True
            For ci = 1 To Len(CHK)
                ch = Mid(CHK, ci, 1)
                If (ch < "0" Or ch > "9") And ch <> "." Then
                    allDigit = False: Exit For
                End If
            Next ci
            If Not allDigit Then GoTo NextCell
            ' 「1.2.3」のような章番号・「.」だけは数でない（CDbl で「型が一致しません」で止まった・2026-09-24 総点検）
            If Len(CHK) - Len(Replace(CHK, ".", "")) > 1 Or Not CHK Like "*#*" Then GoTo NextCell
            If Not IsNumeric(s) Then GoTo NextCell

            v = CDbl(s)
            ' 小数は小数の見える形に（#,##0 だと 1.5 が「2」に見えた・2026-09-24 総点検）。文字の書式（@）のまま値を入れると
            ' 文字に戻るので、書式を先に変える
            If v = Int(v) Then cell.NumberFormat = "#,##0" Else cell.NumberFormat = "#,##0.0##"
            cell.Value = v

NextCell:
        Next r
      Next colIdx
    Next area

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    s = Replace(s, "円", "")
    s = Replace(s, "\", "")
    s = Replace(s, ChrW(&HA5), "")
    s = Replace(s, ChrW(&HFFE5), "")
    s = Replace(s, "$", "")
    s = Trim$(s)
    neg = False
    If Len(s) > 0 Then
        fc = Left(s, 1)
        If fc = "△" Or fc = "▲" Or fc = "-" Then
            neg = True
            s = Mid(s, 2)
            s = Trim$(s)
        End If
    End If
    If Len(s) > 0 Then
        fc = Left(s, 1)
        If fc = "△" Or fc = "▲" Or fc = "-" Then
            neg = True
            s = Mid(s, 2)
            s = Trim$(s)
        End If
    End If
    If neg And セルの字(s) <> "" Then s = "-" & s
    Return

End Sub

Sub 表の中の空行を削除して詰める()
    ' 依頼の語: 空行を消|空白行を消|空白行を削除|空白の行を|明細の間の空|行の抜けを詰め|空行を取り除|空行が無ければ|空行を詰|空白行を詰|空いた行を詰|空行を削除|空白の行を詰
    ' 依頼の組: 空行,空白行,空白の行,空いた行,空の行+削除,消,詰,除+-列
    ' 扱う: 空行削除 空行詰め 合計
    ' 見出し: なし
    ' 形: 数の列
    ' 表の中の完全な空行（見出し列範囲がすべて空）を上から順に行削除して詰める
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)

    ' 見出し行を探す: 値が2つ以上あり、数値だけでない最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
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

    ' 見出し行の左端・右端
    Dim colLeft As Long, colRight As Long
    colLeft = urLeft
    colRight = urRight

    ' 下から上に空行を行削除
    For r = urBottom To hdrRow + 1 Step -1
        Dim allEmpty As Boolean
        allEmpty = True
        For c = colLeft To colRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                allEmpty = False
                Exit For
            End If
        Next c
        If allEmpty Then
            ws.rows(r).Delete Shift:=xlUp
        End If
    Next r
End Sub

Sub 二段の見出しを一行に畳む()
    ' 依頼の語: 2段の見出しを|上下2行|上下二行|二段になっている|見出しの上下2行|見出しを1行に畳|上段と下段|見出しを平|二段の見出し|２段の見出し|2段見出し|二段見出し|見出しを一段|見出しを1段|見出しを１段|見出しを一行|見出しを1行
    ' 依頼の組: 見出し+2段,二段,２段,上下,1段,一段,1行,一行
    ' 扱う: 二段見出し 見出し畳む
    ' 見出し: なし
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    Dim hrow As Long
    hrow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            Dim ma0 As Range
            Set ma0 = ws.Cells(r, c).MergeArea
            If ma0.Cells(1, 1).Row = r And ma0.Cells(1, 1).Column = c Then
                If Trim(セルの字(ws.Cells(r, c).Value)) <> "" Then cnt = cnt + 1
            End If
        Next c
        If cnt >= 2 Then
            hrow = r
            Exit For
        End If
    Next r

    If hrow = 0 Then Exit Sub
    If hrow + 1 > urBottom Then Exit Sub

    Dim row1 As Long, row2 As Long
    row1 = hrow
    row2 = hrow + 1

    Dim hasVal As Boolean
    hasVal = False
    Dim v As Variant
    For c = urLeft To urRight
        v = ws.Cells(row2, c).Value
        If IsNumeric(v) And Trim(CStr(v)) <> "" Then hasVal = True
        If IsDate(v) And Trim(CStr(v)) <> "" Then hasVal = True
    Next c
    If hasVal Then Exit Sub

    Dim hasEmptyOrMerge As Boolean
    hasEmptyOrMerge = False
    For c = urLeft To urRight
        If ws.Cells(row1, c).MergeCells Then
            hasEmptyOrMerge = True
            Exit For
        End If
        If Trim(セルの字(ws.Cells(row1, c).Value)) = "" Then
            hasEmptyOrMerge = True
            Exit For
        End If
    Next c
    If Not hasEmptyOrMerge Then Exit Sub

    Dim topLabel() As String
    ReDim topLabel(urLeft To urRight)
    Dim lastTop As String
    lastTop = ""
    For c = urLeft To urRight
        Dim mac As Range
        Set mac = ws.Cells(row1, c).MergeArea
        Dim topVal As String
        topVal = Trim(CStr(mac.Cells(1, 1).Value))
        If セルの字(topVal) <> "" Then lastTop = topVal
        topLabel(c) = lastTop
    Next c

    Dim botLabel() As String
    ReDim botLabel(urLeft To urRight)
    For c = urLeft To urRight
        botLabel(c) = Trim(CStr(ws.Cells(row2, c).Value))
    Next c

    On Error Resume Next
    ws.rows(row1).UnMerge
    ws.rows(row2).UnMerge
    On Error GoTo 0

    Dim t As String, b As String, combined As String
    For c = urLeft To urRight
        t = topLabel(c)
        b = botLabel(c)
        If セルの字(t) = "" And セルの字(b) = "" Then
            combined = ""
        ElseIf セルの字(t) = "" Then
            combined = b
        ElseIf セルの字(b) = "" Then
            combined = t
        ElseIf t = b Then
            combined = t
        Else
            combined = t & b
        End If
        ws.Cells(row2, c).Value = combined
    Next c

    ws.rows(row1).Delete Shift:=xlUp

End Sub

Sub 選んだ列を選んだ順に左へ移す()
    ' 依頼の語: 選んだ列を先頭に|列の並びを変|列の順番|並べ直|選んだ順|列を左|入れ替
    ' 依頼の組: 列+順番,並び,並べ替え,入れ替,並べ直+-テーブル,グラフ,棒
    ' 扱う: 列並べ替え 列順変更 合計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列
    ' 選んだ列を、選んだ順に表の左へ寄せる。列は「切り取って挿入」で動かす＝Excel が式の参照を付け替える。
    ' 見出し?本文?すぐ下の合計・平均の行を一緒に動かし（合計は自分の列の下に付いていく）、表の下のメモは動かさない。列の幅も一緒に。
    ' （2026-09-24。前は値と式を 1 セルずつ写し直していて、表の下の合計の式が別の列を指し、表の中の式（金額＝数量×単価）も
    '   相対の形のまま運んで違う列を指した）。テーブル・結合セルのある表は切り取りができないので、前と同じ写し直し。

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selRow As Long
    selRow = Selection.Cells(1, 1).Row

    ' 二段見出し（番号・品名は上下 2 行を縦に結合、数量・金額の上に「実績」）は上の段から見る。下の段から選ばれると
    ' 表を数量・金額の 2 列と取り違え、切り取りで右の「備考」の縦の結合がほどけた（2026-09-24 F10）
    Dim urL As Long, urR As Long, hdrTop As Long, hdrBot As Long
    urL = ws.UsedRange.Column
    urR = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    hdrTop = selRow
    If selRow > 1 Then
        If 見出しの下の段(ws, selRow - 1, urL, urR) = selRow Then hdrTop = selRow - 1
    End If
    hdrBot = 見出しの下の段(ws, hdrTop, urL, urR)

    Dim leftCol As Long, rightCol As Long
    ' 左端＝見出しの行で最初に値のある列（表題が A1・表が C3 からのとき A 列から数えて列数 1 と取り違え、エラー 9 で止まった。2026-09-22）
    leftCol = urL
    Do While leftCol < urR And 見出しの字(ws.Cells(hdrTop, leftCol)) = ""
        leftCol = leftCol + 1
    Loop
    rightCol = leftCol
    Do While 見出しの字(ws.Cells(hdrTop, rightCol + 1)) <> ""
        rightCol = rightCol + 1
    Loop
    Dim nCols As Long
    nCols = rightCol - leftCol + 1

    Dim lastRow As Long
    lastRow = 表の本文の最終行(ws)

    Dim totalRows As Long
    totalRows = lastRow - selRow + 1

    Dim nSel As Long
    nSel = Selection.Areas.Count
    Dim selCols() As Long
    ReDim selCols(1 To nSel)
    Dim i As Long
    For i = 1 To nSel
        selCols(i) = Selection.Areas(i).Cells(1, 1).Column
        If selCols(i) < leftCol Or selCols(i) > rightCol Then Exit Sub      ' 表の外の列を選んでいる
    Next i

    Dim newOrder() As Long
    ReDim newOrder(1 To nCols)
    Dim k As Long
    k = 0
    For i = 1 To nSel
        k = k + 1
        newOrder(k) = selCols(i)
    Next i

    Dim j As Long
    For j = leftCol To rightCol
        Dim isSelected As Boolean
        isSelected = False
        For i = 1 To nSel
            If selCols(i) = j Then
                isSelected = True
                Exit For
            End If
        Next i
        If Not isSelected Then
            k = k + 1
            newOrder(k) = j
        End If
    Next j

    ' 見出しを横に結合した列の組（「実績」の下の数量・金額）は離さない＝組の最初の列が来た所に、組の列をまとめて置く
    Dim grpL() As Long, grpR() As Long, ord2() As Long, placed() As Boolean, n2 As Long, u As Long, r3 As Long
    ReDim grpL(leftCol To rightCol): ReDim grpR(leftCol To rightCol)
    For j = leftCol To rightCol
        grpL(j) = j: grpR(j) = j
        For r3 = hdrTop To hdrBot
            If ws.Cells(r3, j).MergeCells Then
                If ws.Cells(r3, j).MergeArea.Columns.Count > 1 Then
                    grpL(j) = ws.Cells(r3, j).MergeArea.Column
                    grpR(j) = grpL(j) + ws.Cells(r3, j).MergeArea.Columns.Count - 1
                End If
            End If
        Next r3
    Next j
    ReDim ord2(1 To nCols): ReDim placed(leftCol To rightCol)
    For i = 1 To nCols
        If Not placed(newOrder(i)) Then
            For u = i To nCols
                If newOrder(u) >= grpL(newOrder(i)) And newOrder(u) <= grpR(newOrder(i)) And Not placed(newOrder(u)) Then
                    n2 = n2 + 1: ord2(n2) = newOrder(u): placed(newOrder(u)) = True
                End If
            Next u
        End If
    Next i
    newOrder = ord2

    Dim alreadyOK As Boolean
    alreadyOK = True
    For i = 1 To nCols
        If newOrder(i) <> leftCol + i - 1 Then
            alreadyOK = False
            Exit For
        End If
    Next i
    If alreadyOK Then Exit Sub

    ' ---- 動かす行: 見出し?本文?すぐ下の合計・平均の行 ----
    Dim botRow As Long, tr As Long, c As Long, s As String, 続く As Boolean
    botRow = lastRow
    tr = 表の下の合計行(ws, lastRow)
    If tr > 0 Then
        botRow = tr
        Do
            続く = False
            For c = leftCol To rightCol
                s = Replace(Replace(セルの字(ws.Cells(botRow + 1, c).Value), " ", ""), "　", "")
                If s = "平均" Or s = "最大" Or s = "最小" Or s = "件数" Or s = "小計" Then
                    続く = True
                    Exit For
                End If
            Next c
            If 続く Then botRow = botRow + 1
        Loop While 続く
    End If

    ' 見出しの結合（縦の「番号」・横の「実績」）は、いったんほどいて覚え、動かした後に新しい位置で結合し直す（2026-09-24 F10）
    Dim mT() As Long, mB() As Long, mL() As Long, mR() As Long, mV() As Variant, nM As Long, ma As Range, mm As Long
    ReDim mT(1 To 1): ReDim mB(1 To 1): ReDim mL(1 To 1): ReDim mR(1 To 1): ReDim mV(1 To 1)
    For r3 = hdrTop To hdrBot
        For j = leftCol To rightCol
            If ws.Cells(r3, j).MergeCells Then
                Set ma = ws.Cells(r3, j).MergeArea
                If ma.Row = r3 And ma.Column = j And ma.Row + ma.rows.Count - 1 <= hdrBot _
                   And ma.Column + ma.Columns.Count - 1 <= rightCol Then
                    nM = nM + 1
                    ReDim Preserve mT(1 To nM): ReDim Preserve mB(1 To nM): ReDim Preserve mL(1 To nM)
                    ReDim Preserve mR(1 To nM): ReDim Preserve mV(1 To nM)
                    mT(nM) = ma.Row: mB(nM) = ma.Row + ma.rows.Count - 1
                    mL(nM) = ma.Column: mR(nM) = ma.Column + ma.Columns.Count - 1
                    mV(nM) = ma.Cells(1, 1).Formula
                    ma.UnMerge
                End If
            End If
        Next j
    Next r3

    ' テーブル・結合セルがあれば切り取りはできない → 前の写し直し
    Dim 写し直し As Boolean, lo As ListObject, blk As Range, mc As Variant
    Set blk = ws.Range(ws.Cells(hdrTop, leftCol), ws.Cells(botRow, rightCol))
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, blk) Is Nothing Then 写し直し = True
    Next lo
    mc = blk.MergeCells
    If IsNull(mc) Then
        写し直し = True
    ElseIf mc Then
        写し直し = True
    End If
    If 写し直し And nM > 0 Then                 ' 写し直しに回るなら、見出しの結合は元の位置に戻す
        Application.DisplayAlerts = False
        For mm = 1 To nM
            ws.Range(ws.Cells(mT(mm), mL(mm)), ws.Cells(mB(mm), mR(mm))).Merge
            ws.Cells(mT(mm), mL(mm)).Formula = mV(mm)
        Next mm
        Application.DisplayAlerts = True
    End If

    If Not 写し直し Then
        Dim 幅() As Double, cur() As Long, t As Long, p As Long, q As Long, tmp As Long
        Dim 語元列() As Long, r2 As Long
        If tr > 0 Then
            ReDim 語元列(1 To botRow - tr + 1)
            For r2 = tr To botRow                ' 合計・平均の語がもともと何列目にあったか
                For c = leftCol To rightCol
                    s = Replace(Replace(セルの字(ws.Cells(r2, c).Value), " ", ""), "　", "")
                    If s = "合計" Or s = "総計" Or s = "計" Or s = "平均" Or s = "最大" Or s = "最小" Or s = "件数" Or s = "小計" Then
                        語元列(r2 - tr + 1) = c
                        Exit For
                    End If
                Next c
            Next r2
        End If
        ReDim 幅(leftCol To rightCol)
        For j = leftCol To rightCol
            幅(j) = ws.Columns(j).ColumnWidth
        Next j
        ReDim cur(1 To nCols)
        For t = 1 To nCols
            cur(t) = leftCol + t - 1          ' 今その位置にある元の列
        Next t
        Application.ScreenUpdating = False
        On Error GoTo 動かせない
        For t = 1 To nCols
            For p = t To nCols
                If cur(p) = newOrder(t) Then Exit For
            Next p
            If p <= nCols And p <> t Then
                ws.Range(ws.Cells(hdrTop, leftCol + p - 1), ws.Cells(botRow, leftCol + p - 1)).Cut
                ws.Range(ws.Cells(hdrTop, leftCol + t - 1), ws.Cells(botRow, leftCol + t - 1)).Insert Shift:=xlToRight
                tmp = cur(p)
                For q = p To t + 1 Step -1
                    cur(q) = cur(q - 1)
                Next q
                cur(t) = tmp
            End If
        Next t
        On Error GoTo 0
        For t = 1 To nCols
            ws.Columns(leftCol + t - 1).ColumnWidth = 幅(newOrder(t))
        Next t
        Application.CutCopyMode = False
        If nM > 0 Then                          ' 見出しの結合を、列の新しい位置で結合し直す（組の列は隣り合っている）
            Dim posOf() As Long, a1 As Long, a2 As Long
            ReDim posOf(leftCol To rightCol)
            For t = 1 To nCols
                posOf(newOrder(t)) = leftCol + t - 1
            Next t
            Application.DisplayAlerts = False
            For mm = 1 To nM
                a1 = rightCol: a2 = leftCol
                For j = mL(mm) To mR(mm)
                    If posOf(j) < a1 Then a1 = posOf(j)
                    If posOf(j) > a2 Then a2 = posOf(j)
                Next j
                ws.Range(ws.Cells(mT(mm), a1), ws.Cells(mB(mm), a2)).Merge
                ws.Cells(mT(mm), a1).Formula = mV(mm)
            Next mm
            Application.DisplayAlerts = True
        End If
        ' 合計・平均の行の語（合計・平均…）は数字のそばに置き直す。語の書いてあった列ごと動くと、数字から離れた（2026-09-24
        '   テスト用1: 金額を先頭にしたら「合計」が単価の列と一緒に右端へ行った）。もとが表の 1 列目なら 1 列目に戻す
        If tr > 0 Then
            Dim rr As Long, 語今 As Long, 値左 As Long, 置く As Long
            For rr = tr To botRow
                語今 = 0: 値左 = 0
                For c = leftCol To rightCol
                    s = Replace(Replace(セルの字(ws.Cells(rr, c).Value), " ", ""), "　", "")
                    If 語今 = 0 And (s = "合計" Or s = "総計" Or s = "計" Or s = "平均" Or s = "最大" Or s = "最小" Or s = "件数" Or s = "小計") Then 語今 = c
                    If 値左 = 0 And (ws.Cells(rr, c).HasFormula Or (Not IsError(ws.Cells(rr, c).Value) And VarType(ws.Cells(rr, c).Value) = vbDouble)) Then 値左 = c
                Next c
                If 語今 > 0 And 値左 > 0 Then
                    置く = 0
                    If 語元列(rr - tr + 1) = leftCol And 値左 > leftCol Then
                        置く = leftCol
                    ElseIf 値左 - 1 >= leftCol Then
                        置く = 値左 - 1
                    End If
                    If 置く > 0 And 置く <> 語今 Then
                        If セルの字(ws.Cells(rr, 置く).Value) <> "" Then 置く = 0
                    End If
                    If 置く = 0 And 値左 + 1 <= rightCol Then
                        If セルの字(ws.Cells(rr, 値左 + 1).Value) = "" Or 値左 + 1 = 語今 Then 置く = 値左 + 1
                    End If
                    If 置く > 0 And 置く <> 語今 Then ws.Cells(rr, 語今).Cut ws.Cells(rr, 置く)
                End If
            Next rr
            Application.CutCopyMode = False
        End If
        Application.ScreenUpdating = True
        Application.StatusBar = "列の並べ替え: 選んだ " & nSel & " 列を左へ寄せました（式の参照・表の下の合計・列の幅も一緒に動かしました）。"
        Exit Sub
    End If

    ' ---- テーブル・結合セルのある表: 値と式を写し直す（前の作り）----
    Dim bufVal() As Variant
    Dim bufFmt() As String
    Dim bufBold() As Boolean
    ReDim bufVal(1 To totalRows, 1 To nCols)
    ReDim bufFmt(1 To totalRows, 1 To nCols)
    ReDim bufBold(1 To totalRows, 1 To nCols)

    Dim cel As Range
    For j = 1 To nCols
        For i = 1 To totalRows
            Set cel = ws.Cells(selRow + i - 1, leftCol + j - 1)
            ' 式は相対の形（R1C1）のまま運ぶ（値で写して合計行の SUM が数字になった。2026-09-22）
            If cel.HasFormula Then bufVal(i, j) = Chr(1) & cel.FormulaR1C1 Else bufVal(i, j) = cel.Value
            bufFmt(i, j) = cel.NumberFormat
            bufBold(i, j) = cel.Font.Bold
        Next i
    Next j

    Dim srcCol As Long
    For j = 1 To nCols
        srcCol = newOrder(j) - leftCol + 1
        For i = 1 To totalRows
            Set cel = ws.Cells(selRow + i - 1, leftCol + j - 1)
            cel.NumberFormat = bufFmt(i, srcCol)
            If VarType(bufVal(i, srcCol)) = vbString And Left$(CStr(bufVal(i, srcCol)), 1) = Chr(1) Then
                cel.FormulaR1C1 = Mid$(CStr(bufVal(i, srcCol)), 2)
            Else
                cel.Value = bufVal(i, srcCol)
            End If
            cel.Font.Bold = bufBold(i, srcCol)
        Next i
    Next j
    Exit Sub

動かせない:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.StatusBar = "列の並べ替え: 列を動かせませんでした（" & Err.Description & "）。控えから戻せます。"
End Sub

Sub 選んだ列の途切れた式を戻す()
    ' 依頼の語: 式が途切れ|式の途切れ|値で上書き|計算式が崩れ|上下と同じ式に戻|式が抜|列の数式
    ' 依頼の組: 式,数式+消えて,途切れ,抜け,崩れ,そろえ,揃え,戻+-一覧,監査,調べ,説明,株式,形式,様式,書式,正式
    ' 扱う: 式揃え 数式補完 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim col As Long
    Dim area As Range
    For Each area In Selection.Areas
        Dim c As Long
        For c = area.Column To area.Column + area.Columns.Count - 1
            col = c
            ' 本文の範囲
            Dim lastRow As Long
            lastRow = 表の本文の最終行(ws)

            ' 式の集計
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")

            Dim r As Long
            For r = hdrRow + 1 To lastRow
                ' 行まるごと空チェック
                Dim rowRange As Range
                Set rowRange = ws.rows(r)
                If WorksheetFunction.CountA(rowRange) = 0 Then GoTo NextRow1

                ' 合計行チェック
                Dim rowUsed As Range
                Set rowUsed = ws.UsedRange.rows(r - ws.UsedRange.Row + 1)
                Dim found As Boolean
                ' 合計・小計の行（「総務課 小計」も）は式をそろえない（小計の SUM を明細の式で上書きしていた・2026-09-24 総点検）
                found = 集計の行か(ws, r, ws.UsedRange.Column, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)
                If found Then GoTo NextRow1

                ' 式を記録
                Dim cellVal As Range
                Set cellVal = ws.Cells(r, col)
                If cellVal.HasFormula Then
                    Dim f As String
                    f = cellVal.FormulaR1C1
                    If dict.Exists(f) Then
                        dict(f) = dict(f) + 1
                    Else
                        dict.Add f, 1
                    End If
                End If
NextRow1:
            Next r

            ' 式が1つもなければスキップ
            If dict.Count = 0 Then GoTo NextCol

            ' 最多の式を探す
            Dim bestFormula As String
            Dim bestCount As Long
            bestFormula = ""
            bestCount = 0
            Dim key As Variant
            For Each key In dict.keys
                If dict(key) > bestCount Then
                    bestCount = dict(key)
                    bestFormula = CStr(key)
                End If
            Next key

            ' 式でないセルに適用
            For r = hdrRow + 1 To lastRow
                Set rowRange = ws.rows(r)
                If WorksheetFunction.CountA(rowRange) = 0 Then GoTo NextRow2

                Set rowUsed = ws.UsedRange.rows(r - ws.UsedRange.Row + 1)
                found = 集計の行か(ws, r, ws.UsedRange.Column, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)
                If found Then GoTo NextRow2

                Set cellVal = ws.Cells(r, col)
                ' 空セルは触らない
                If セルの字(cellVal.Value) = "" And Not cellVal.HasFormula Then GoTo NextRow2

                If Not cellVal.HasFormula Then
                    cellVal.FormulaR1C1 = bestFormula
                ElseIf cellVal.FormulaR1C1 <> bestFormula Then
                    cellVal.FormulaR1C1 = bestFormula
                End If
NextRow2:
            Next r
NextCol:
        Next c
    Next area
End Sub

Sub テーブルの範囲を下と右へ広げる()
    ' 依頼の語: 拡張|テーブルに新しい行|下の行|テーブルの範囲|足した行|追加した列
    ' 依頼の組: テーブル+範囲,広げ,拡張,含め+-並べ
    ' 扱う: テーブル拡張
    ' 見出し: なし
    ' 形: 数の列 テーブル
    ' テーブルのすぐ下の行・すぐ右の列まで範囲をListObject.Resizeで広げる
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastCol As Long, lastRow As Long
    Dim hdrRow As Long, hdrCol As Long, tblLastCol As Long, tblLastRow As Long
    Dim c As Long, r As Long
    Dim newLastCol As Long, newLastRow As Long
    Dim hasVal As Boolean

    Set ws = ActiveSheet

    Dim i As Integer
    For i = 1 To ws.ListObjects.Count
        Set lo = ws.ListObjects(i)

        ' --- 右に広げる列を決める ---
        hdrRow = lo.HeaderRowRange.Row
        tblLastCol = lo.Range.Column + lo.Range.Columns.Count - 1
        newLastCol = tblLastCol
        c = tblLastCol + 1
        Do While セルの字(ws.Cells(hdrRow, c).Value) <> ""
            newLastCol = c
            c = c + 1
        Loop

        ' --- 下に広げる行を決める ---
        tblLastRow = lo.Range.Row + lo.Range.rows.Count - 1
        newLastRow = tblLastRow
        r = tblLastRow + 1
        Do
            hasVal = False
            Dim cc As Long
            For cc = lo.Range.Column To newLastCol
                If セルの字(ws.Cells(r, cc).Value) <> "" Then
                    hasVal = True
                    Exit For
                End If
            Next cc
            If Not hasVal Then Exit Do
            newLastRow = r
            r = r + 1
        Loop

        ' --- 変化がなければスキップ ---
        If newLastCol = tblLastCol And newLastRow = tblLastRow Then GoTo NextTable

        ' --- Resize ---
        Dim newRange As Range
        Set newRange = ws.Range(lo.Range.Cells(1, 1), ws.Cells(newLastRow, newLastCol))
        lo.Resize newRange

NextTable:
    Next i
End Sub

Sub 表の中の空の列を削除して詰める()
    ' 依頼の語: 空の列を削除|空白の列を消|空列を削除|空列を消|空白の列を削除|空列を取り除|間の空いた列を詰|空の列を消|最後に値|入っていない列
    ' 依頼の組: 空の列,空列,空白の列,空いた列,入っていない列+削除,消,詰,除
    ' 扱う: 空列削除
    ' 見出し: なし
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    ' 見出し行を探す: 値が2つ以上ある最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    Dim cnt As Long
    For r = ur.Row To ur.Row + ur.rows.Count - 1
        cnt = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
            If cnt >= 2 Then Exit For
        Next c
        If cnt >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    ' 見出し行で最初と最後に値のある列
    Dim firstCol As Long, lastCol As Long
    firstCol = 0
    lastCol = 0
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        If セルの字(ws.Cells(hdrRow, c).Value) <> "" Then
            If firstCol = 0 Then firstCol = c
            lastCol = c
        End If
    Next c
    If firstCol = 0 Then Exit Sub

    ' 表の最後の行: 見出し行から下で、firstCol～lastColのいずれかに値がある最後の行
    Dim lastRow As Long
    lastRow = hdrRow
    For r = ur.Row + ur.rows.Count - 1 To hdrRow Step -1
        Dim found As Boolean
        found = False
        For c = firstCol To lastCol
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                found = True
                Exit For
            End If
        Next c
        If found Then
            lastRow = r
            Exit For
        End If
    Next r

    ' 右から順に、見出しも本文も全部空の列を削除
    For c = lastCol To firstCol Step -1
        Dim isEmpty As Boolean
        isEmpty = True
        For r = hdrRow To lastRow
            If セルの字(ws.Cells(r, c).Value) <> "" Then
                isEmpty = False
                Exit For
            End If
        Next r
        ' 表の外（表題・下のメモ）に字がある列は消さない（列ごと消してメモまで消していた・2026-09-24 総点検）
        If isEmpty Then
            If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Cells(hdrRow, c).EntireColumn.Delete
        End If
    Next c

End Sub

Sub 乱れた列を掃除し右に足す()
    ' 依頼の語: データの掃除|クレンジング|データを掃除|取り込んだデータが汚|表記がバラバラ|データをきれいに|汚いデータ|文字列の乱れ|表記ゆれを直
    ' 依頼の組: データ+掃除,汚,きたな,クレンジング,きれい+-CSV
    ' 扱う: 空白除去 文字数字 文字日付 作業列
    ' 見出し: なし
    ' 形: 数の列
    ' 乱れのある列の右隣に作業列を作り、直した値を「数式」で入れる。元の列は触らない。
    ' 直し方は列ごとに選ぶ（文字の数字→数値／文字の日付→日付／それ以外→空白とゴミの除去）。
    ' 置き換えるかは人が選べる＝作業列を値で貼り付けて元の列を消すだけ。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim hdrRow As Long, r As Long, c As Long, cnt As Long
    Dim s As String, t As String
    Dim n数 As Long, n日 As Long, n乱 As Long, n文字 As Long, 作列数 As Long
    Dim 種 As String, 列名 As String, 名 As String
    Dim adr As String

    Set ws = ActiveSheet
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub

    urTop = ur.Row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    ' 見出し行＝値が2つ以上ある最初の行
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Or hdrRow >= urBottom Then Exit Sub
    ' 二段見出しは下の段が見出し（下の段の「数量」を明細と見て作業列に =TRIM(E4) を入れていた・2026-09-24 棚の試験 F10）
    hdrRow = 見出しの下の段(ws, hdrRow, urLeft, urRight)

    Application.ScreenUpdating = False

    ' 右の列から左へ（挿入で列がずれないように）
    For c = urRight To urLeft Step -1
        n数 = 0: n日 = 0: n乱 = 0: n文字 = 0
        For r = hdrRow + 1 To urBottom
            If Not ws.Cells(r, c).HasFormula Then
                If VarType(ws.Cells(r, c).Value) = vbString Then
                    s = CStr(ws.Cells(r, c).Value)
                    If Len(s) > 0 Then
                        n文字 = n文字 + 1
                        t = s
                        t = Replace(t, vbCr, "")
                        t = Replace(t, vbLf, "")
                        t = Replace(t, vbTab, " ")
                        t = Replace(t, ChrW(160), " ")
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
            End If
        Next r

        種 = ""
        If n文字 > 0 Then
            If n数 * 2 >= n文字 And n数 >= n日 Then
                種 = "数"
            ElseIf n日 * 2 >= n文字 And n日 >= n数 Then
                種 = "日"
            ElseIf n乱 >= 1 Then
                種 = "乱"
            End If
        End If
        If セルの字(種) = "" Then GoTo 次の列

        ' 右に作業列を作る
        ws.Columns(c + 1).Insert Shift:=xlToRight
        ' 挿入した列は左の書式を継ぐ＝文字（@）のセルでは式が計算されず字のまま残る（2026-09-24 改名後の撃ち試し F3）
        ws.Range(ws.Cells(hdrRow + 1, c + 1), ws.Cells(urBottom, c + 1)).NumberFormat = "General"
        作列数 = 作列数 + 1
        名 = 見出しの字(ws.Cells(hdrRow, c))
        If セルの字(名) = "" Then 名 = "列" & c
        ws.Cells(hdrRow, c + 1).Value = 名 & "（直し）"
        ws.Cells(hdrRow, c + 1).Font.Bold = ws.Cells(hdrRow, c).Font.Bold

        For r = hdrRow + 1 To urBottom
            adr = ws.Cells(r, c).Address(0, 0)
            If 種 = "数" Then
                ws.Cells(r, c + 1).Formula = "=IF(" & adr & "="""","""",IFERROR(VALUE(TRIM(" & adr & ")),TRIM(" & adr & ")))"
            ElseIf 種 = "日" Then
                ws.Cells(r, c + 1).Formula = "=IF(" & adr & "="""","""",IFERROR(DATEVALUE(TRIM(" & adr & ")),TRIM(" & adr & ")))"
            Else
                ws.Cells(r, c + 1).Formula = "=IF(" & adr & "="""","""",TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & adr & ",CHAR(10),""""),CHAR(9),"" ""),CHAR(160),"" "")))"
            End If
        Next r

        If 種 = "日" Then
            ws.Range(ws.Cells(hdrRow + 1, c + 1), ws.Cells(urBottom, c + 1)).NumberFormatLocal = "yyyy/m/d"
        ElseIf 種 = "数" Then
            ws.Range(ws.Cells(hdrRow + 1, c + 1), ws.Cells(urBottom, c + 1)).NumberFormatLocal = "#,##0"
        End If
        ws.Columns(c + 1).AutoFit
        列名 = 列名 & "／" & 名
次の列:
    Next c

    Application.ScreenUpdating = True

    If 作列数 = 0 Then
        Application.StatusBar = "データの掃除: 乱れている列は見つかりませんでした。"
    Else
        Application.StatusBar = "データの掃除: " & 作列数 & " 列の右に「（直し）」列を数式で作りました（" & Mid$(列名, 2) & "）。よければ値で貼り付けて元の列を消してください。"
    End If
End Sub

Sub 貼ったCSVを使える形に整える()
    ' 依頼の語: CSV取り込み|CSVを取り込ん|CSVを貼り付け|システムの出力を貼|取り込んだ直後|CSVの整形|CSVを整え|貼り付けた表を整え|取り込み後の整形|CSVの後始末
    ' 依頼の組: CSV,貼り付け,取り込,貼った+整,後,使える形+-書き出,に出力,で出力,保存
    ' 扱う: 文字数字 文字日付 見出し 枠固定 列幅
    ' 見出し: なし
    ' 形: 数の列
    ' CSV やシステム出力を貼った直後の表を、そのまま使える形にする。
    '   文字で入った数値・日付を値に直す（先頭ゼロの列は触らない）→ 見出しを太字 → 枠固定 → 列幅
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim hdrRow As Long, r As Long, c As Long, cnt As Long
    Dim s As String, t As String
    Dim n数 As Long, n日 As Long, n文字 As Long, n先頭ゼロ As Long
    Dim 直数 As Long, 直日 As Long, 直列 As Long, 直乱 As Long

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
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Or hdrRow >= urBottom Then Exit Sub

    Application.ScreenUpdating = False

    For c = urLeft To urRight
        n数 = 0: n日 = 0: n文字 = 0: n先頭ゼロ = 0
        For r = hdrRow + 1 To urBottom
            If Not ws.Cells(r, c).HasFormula Then
                If VarType(ws.Cells(r, c).Value) = vbString Then
                    s = CStr(ws.Cells(r, c).Value)
                    If Len(s) > 0 Then
                        n文字 = n文字 + 1
                        t = Replace(Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, " "), ChrW(160), " ")
                        Do While Len(t) > 0
                            If Left$(t, 1) = " " Or Left$(t, 1) = ChrW(&H3000) Then t = Mid$(t, 2) Else Exit Do
                        Loop
                        Do While Len(t) > 0
                            If Right$(t, 1) = " " Or Right$(t, 1) = ChrW(&H3000) Then t = Left$(t, Len(t) - 1) Else Exit Do
                        Loop
                        If Len(t) > 1 And Left$(t, 1) = "0" And Left$(t, 2) <> "0." Then
                            n先頭ゼロ = n先頭ゼロ + 1
                        ElseIf Len(t) > 0 And IsNumeric(t) Then
                            n数 = n数 + 1
                        ElseIf 日付の字か(t) Then   ' 年・月・日のそろう形だけ（枝番 1-2・郵便番号は日付でない・2026-09-24 総点検）
                            n日 = n日 + 1
                        End If
                    End If
                End If
            End If
        Next r

        ' 先頭ゼロ（伝票番号・郵便番号など）が1つでもある列は文字のまま残す
        If n先頭ゼロ > 0 Then GoTo 次の列
        If n文字 = 0 Then GoTo 次の列

        If n数 * 2 >= n文字 And n数 >= n日 Then
            For r = hdrRow + 1 To urBottom
                If Not ws.Cells(r, c).HasFormula Then
                    If VarType(ws.Cells(r, c).Value) = vbString Then
                        t = Trim$(Replace(Replace(CStr(ws.Cells(r, c).Value), ChrW(160), " "), ChrW(&H3000), " "))
                        If Len(t) > 0 And IsNumeric(t) Then
                            ws.Cells(r, c).NumberFormatLocal = "#,##0"
                            ws.Cells(r, c).Value = CDbl(t)
                            ws.Cells(r, c).HorizontalAlignment = xlGeneral
                            直数 = 直数 + 1
                        End If
                    End If
                End If
            Next r
            直列 = 直列 + 1
        ElseIf n日 * 2 >= n文字 And n日 > n数 Then
            For r = hdrRow + 1 To urBottom
                If Not ws.Cells(r, c).HasFormula Then
                    If VarType(ws.Cells(r, c).Value) = vbString Then
                        t = Trim$(Replace(Replace(CStr(ws.Cells(r, c).Value), ChrW(160), " "), ChrW(&H3000), " "))
                        If 日付の字か(t) Then
                            ws.Cells(r, c).NumberFormatLocal = "yyyy/m/d"
                            ws.Cells(r, c).Value = CDate(t)
                            ws.Cells(r, c).HorizontalAlignment = xlGeneral
                            直日 = 直日 + 1
                        End If
                    End If
                End If
            Next r
            直列 = 直列 + 1
        Else
            ' 文字の列は、見えないゴミ（改行・タブ・ノーブレークスペース・前後の空白）だけ落とす
            For r = hdrRow + 1 To urBottom
                If Not ws.Cells(r, c).HasFormula Then
                    If VarType(ws.Cells(r, c).Value) = vbString Then
                        s = CStr(ws.Cells(r, c).Value)
                        t = Replace(Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, " "), ChrW(160), " ")
                        Do While Len(t) > 0
                            If Left$(t, 1) = " " Or Left$(t, 1) = ChrW(&H3000) Then t = Mid$(t, 2) Else Exit Do
                        Loop
                        Do While Len(t) > 0
                            If Right$(t, 1) = " " Or Right$(t, 1) = ChrW(&H3000) Then t = Left$(t, Len(t) - 1) Else Exit Do
                        Loop
                        If t <> s Then
                            ws.Cells(r, c).Value = t
                            直乱 = 直乱 + 1
                        End If
                    End If
                End If
            Next r
        End If
次の列:
    Next c

    ' 見出しを太字にして枠で固定、列幅を整える
    ws.Range(ws.Cells(hdrRow, urLeft), ws.Cells(hdrRow, urRight)).Font.Bold = True
    ws.Range(ws.Cells(urTop, urLeft), ws.Cells(urBottom, urRight)).Columns.AutoFit
    For c = urLeft To urRight
        If ws.Columns(c).ColumnWidth > 40 Then ws.Columns(c).ColumnWidth = 40
    Next c

    If ws.Parent Is ActiveWorkbook Then
        If ws.Name = ActiveSheet.Name Then
            ActiveWindow.FreezePanes = False
            ws.Cells(hdrRow + 1, urLeft).Select
            ActiveWindow.FreezePanes = True
            ws.Cells(hdrRow, urLeft).Select
        End If
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "CSVの整形: " & 直列 & " 列を値に直しました（数値 " & 直数 & " ／日付 " & 直日 & " ／見えないゴミ " & 直乱 & "）。見出しを太字にして " & (hdrRow + 1) & " 行目で枠を固定、列幅を整えました。"
End Sub

Sub 列幅と行の高さを全シートに写す()
    ' 依頼の語: 体裁を揃え|体裁をそろえ|見た目をそろえ|列幅をそろえ|行の高さをそろえ|全シートの見た目|シートごとにバラバラ|同じ形のシート|表示倍率をそろえ|どのシートも同じ
    ' 依頼の組: シート+列幅,倍率,見た目,体裁,そろえ,揃え,合わせ+-グラフ,ピボット,分け,集約,まとめ
    ' 扱う: 列幅 行高 倍率 全シート
    ' 見出し: なし
    ' 形: 数の列
    ' 今のシートを手本にして、同じブックの他のシートの列幅・行の高さ・表示倍率・先頭の位置をそろえる。
    ' 最後に文字が切れていないか（###）を確かめ、切れていたらその列を全シートで広げる。
    Dim 手本 As Worksheet, sh As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim 幅() As Double, 高さ() As Double
    Dim r As Long, c As Long, 倍率 As Long, 先頭行 As Long, 先頭列 As Long
    Dim n枚 As Long, n切れ As Long, 広げた As Long

    Set 手本 = ActiveSheet
    On Error Resume Next
    Set ur = 手本.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.Row + ur.rows.Count - 1

    ReDim 幅(urLeft To urRight)
    For c = urLeft To urRight
        幅(c) = 手本.Columns(c).ColumnWidth
    Next c
    ReDim 高さ(urTop To urBottom)
    For r = urTop To urBottom
        高さ(r) = 手本.rows(r).RowHeight
    Next r
    倍率 = ActiveWindow.Zoom
    先頭行 = ActiveWindow.ScrollRow
    先頭列 = ActiveWindow.ScrollColumn

    Application.ScreenUpdating = False
    For Each sh In 手本.Parent.Worksheets
        If sh.Name <> 手本.Name And sh.Visible = xlSheetVisible Then
            For c = urLeft To urRight
                sh.Columns(c).ColumnWidth = 幅(c)
            Next c
            For r = urTop To urBottom
                sh.rows(r).RowHeight = 高さ(r)
            Next r
            sh.Activate
            ActiveWindow.Zoom = 倍率
            ActiveWindow.ScrollRow = 先頭行
            ActiveWindow.ScrollColumn = 先頭列
            n枚 = n枚 + 1
        End If
    Next sh
    手本.Activate

    ' 文字が切れていないか（###）を確かめ、切れていたら全シートで同じだけ広げる
    For c = urLeft To urRight
        n切れ = 0
        For Each sh In 手本.Parent.Worksheets
            If sh.Visible = xlSheetVisible Then
                For r = urTop To urBottom
                    If Not IsError(sh.Cells(r, c).Value) Then
                        If InStr(sh.Cells(r, c).text, "###") > 0 Then n切れ = n切れ + 1
                    End If
                Next r
            End If
        Next sh
        If n切れ > 0 Then
            手本.Columns(c).AutoFit
            If 手本.Columns(c).ColumnWidth < 幅(c) Then 手本.Columns(c).ColumnWidth = 幅(c)
            If 手本.Columns(c).ColumnWidth > 40 Then 手本.Columns(c).ColumnWidth = 40
            For Each sh In 手本.Parent.Worksheets
                If sh.Visible = xlSheetVisible Then sh.Columns(c).ColumnWidth = 手本.Columns(c).ColumnWidth
            Next sh
            広げた = 広げた + 1
        End If
    Next c

    Application.ScreenUpdating = True
    Application.StatusBar = "体裁をそろえる: 手本「" & 手本.Name & "」に合わせて " & n枚 & " 枚のシートの列幅・行の高さ・表示倍率（" & 倍率 & "%）・先頭の位置をそろえました。文字が切れていた " & 広げた & " 列は広げました。"
End Sub

Sub 外部リンクを値に変えて切る()
    ' 依頼の語: 外部リンク|リンクを解消|リンクを切|リンクを更新しますか|外部参照|他のブックを参照|リンク元|別ブックへの参照|リンクが残
    ' 依頼の組: リンク,外部参照,外部ブック+切,解消,外,消,値に+-調べ,一覧,探
    ' 扱う: 外部リンク 解消 控え 記録
    ' 見出し: なし
    ' 形: 数の列
    ' 外部ブックへのリンクを一覧に残してから値に変える。切る前にブックの控えを1部作る。
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, cel As Range, 式域 As Range
    Dim リンク As Variant, i As Long, r As Long
    Dim 名 As String, 控え As String, 元エラー As Long, 後エラー As Long
    Dim outRow As Long, n参照 As Long, f As String

    Set wb = ActiveWorkbook
    リンク = wb.LinkSources(xlExcelLinks)
    If isEmpty(リンク) Then
        Application.StatusBar = "外部リンクの解消: このブックに外部ブックへのリンクはありません。"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 記録のシートを作る
    名 = "リンク解消の記録"
    i = 1
    Do While i < 100
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(名)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        名 = "リンク解消の記録" & i: i = i + 1
    Loop
    Set sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = 名

    sh.Cells(1, 1).Value = "外部リンクの解消（" & Format$(Now, "yyyy/m/d h:nn") & "）"
    sh.Cells(1, 1).Font.Bold = True
    sh.Cells(2, 1).Value = "種類"
    sh.Cells(2, 2).Value = "場所"
    sh.Cells(2, 3).Value = "内容"
    sh.Range(sh.Cells(2, 1), sh.Cells(2, 3)).Font.Bold = True
    outRow = 3

    For i = LBound(リンク) To UBound(リンク)
        sh.Cells(outRow, 1).Value = "リンク元"
        sh.Cells(outRow, 2).Value = "ブック全体"
        sh.Cells(outRow, 3).Value = リンク(i)
        outRow = outRow + 1
    Next i

    ' 外部参照を含む数式と、今のエラー数を数える
    For Each ws In wb.Worksheets
        If ws.Name <> 名 Then
            Set 式域 = Nothing
            On Error Resume Next
            Set 式域 = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not 式域 Is Nothing Then
                For Each cel In 式域
                    f = cel.Formula
                    If InStr(f, "[") > 0 And InStr(f, "]") > 0 Then
                        n参照 = n参照 + 1
                        If outRow < 5000 Then
                            sh.Cells(outRow, 1).Value = "外部参照の式"
                            sh.Cells(outRow, 2).Value = ws.Name & "!" & cel.Address(0, 0)
                            sh.Cells(outRow, 3).Value = "'" & f
                            outRow = outRow + 1
                        End If
                    End If
                    If IsError(cel.Value) Then 元エラー = 元エラー + 1
                Next cel
            End If
        End If
    Next ws

    ' 控えを1部作ってから切る
    控え = ""
    If セルの字(wb.path) <> "" Then
        控え = wb.path & "\" & Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & _
               "_リンク解消前_" & Format$(Now, "yyyymmdd_hhnnss") & Mid$(wb.Name, InStrRev(wb.Name, "."))
        On Error Resume Next
        wb.SaveCopyAs 控え
        If Err.Number <> 0 Then 控え = ""
        Err.Clear
        On Error GoTo 0
    End If

    For i = LBound(リンク) To UBound(リンク)
        On Error Resume Next
        wb.BreakLink Name:=リンク(i), Type:=xlLinkTypeExcelLinks
        On Error GoTo 0
    Next i

    For Each ws In wb.Worksheets
        If ws.Name <> 名 Then
            For Each cel In ws.UsedRange
                If IsError(cel.Value) Then 後エラー = 後エラー + 1
            Next cel
        End If
    Next ws

    sh.Cells(outRow, 1).Value = "結果"
    sh.Cells(outRow, 2).Value = "ブック全体"
    sh.Cells(outRow, 3).Value = "リンク " & (UBound(リンク) - LBound(リンク) + 1) & " 本・外部参照の式 " & n参照 & " か所を値に変えました。" & _
                               "エラーのセルは " & 元エラー & " → " & 後エラー & " 件。" & _
                               IIf(控え = "", "（控えは作れませんでした＝ブックが未保存）", "控え: " & 控え)
    sh.Range(sh.Cells(2, 1), sh.Cells(outRow, 3)).Borders.LineStyle = xlContinuous
    sh.Columns("A:C").AutoFit
    If sh.Columns(3).ColumnWidth > 80 Then sh.Columns(3).ColumnWidth = 80

    Application.ScreenUpdating = True
    Application.StatusBar = "外部リンクの解消: " & (UBound(リンク) - LBound(リンク) + 1) & " 本を値に変えました（式 " & n参照 & " か所・エラー " & 元エラー & "→" & 後エラー & " 件）。" & _
                           IIf(控え = "", "控えは作れませんでした（未保存のブック）。", "控えを作りました: " & 控え)
End Sub

Sub 名前を一覧にし壊れた名前を削除()
    ' 依頼の語: 名前の定義|名前付き範囲|名前を整理|名前の一覧|REFの名前|名前が増えすぎ|いらない名前|名前定義|定義された名前
    ' 依頼の組: 名前+定義,範囲+整理,消,削除
    ' 扱う: 名前 整理 壊れた名前 一覧
    ' 見出し: なし
    ' 形: 数の列
    ' 名前の定義を一覧にして、#REF! を含む壊れた名前だけ消す。
    ' どこからも使われていない名前は「消してよい候補」として印を付けるだけで消さない。
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, 式域 As Range, cel As Range
    Dim nM As Name, 参照 As String, 名前 As String, 短名 As String
    Dim 全式 As String, i As Long, p As Long, 使用 As Boolean
    Dim outRow As Long, 名 As String, n壊れ As Long, n未使用 As Long, n全 As Long, 消した As Long
    Dim 前後 As String, j As Long

    Set wb = ActiveWorkbook
    If wb.names.Count = 0 Then
        Application.StatusBar = "名前の整理: このブックに名前の定義はありません。"
        Exit Sub
    End If
    Application.ScreenUpdating = False

    ' 全シートの数式を1本の文字列に集める（使われているかを見るため）
    For Each ws In wb.Worksheets
        Set 式域 = Nothing
        On Error Resume Next
        Set 式域 = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not 式域 Is Nothing Then
            For Each cel In 式域
                If Len(全式) < 900000 Then 全式 = 全式 & " " & cel.Formula
            Next cel
        End If
    Next ws

    ' ほかの名前の中身からの参照も「使われている」に数える（税率→税込 のように名前から名前を使う形を
    ' 「どこからも使われていない」と言っていた・2026-09-24 総点検）
    For Each nM In wb.names
        On Error Resume Next
        If Len(全式) < 900000 Then 全式 = 全式 & " " & nM.RefersTo
        On Error GoTo 0
    Next nM

    名 = "名前の一覧"
    i = 1
    Do While i < 100
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(名)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        名 = "名前の一覧" & i: i = i + 1
    Loop
    Set sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = 名

    sh.Cells(1, 1).Value = "名前の定義の一覧（" & Format$(Now, "yyyy/m/d h:nn") & "）"
    sh.Cells(1, 1).Font.Bold = True
    sh.Cells(2, 1).Value = "名前"
    sh.Cells(2, 2).Value = "参照先"
    sh.Cells(2, 3).Value = "状態"
    sh.Cells(2, 4).Value = "どうしたか"
    sh.Range(sh.Cells(2, 1), sh.Cells(2, 4)).Font.Bold = True
    outRow = 3

    For i = wb.names.Count To 1 Step -1
        Set nM = wb.names(i)
        名前 = "": 参照 = ""
        On Error Resume Next
        名前 = nM.Name
        参照 = nM.RefersTo
        On Error GoTo 0
        If セルの字(名前) = "" Then GoTo 次の名前
        n全 = n全 + 1

        短名 = 名前
        p = InStr(短名, "!")
        If p > 0 Then 短名 = Mid$(短名, p + 1)
        ' Excel が使う名前（印刷範囲・印刷タイトル・絞り込み・新しい関数の印）は使われている扱い（2026-09-24 総点検）
        Dim 組込 As Boolean
        組込 = (短名 = "Print_Area" Or 短名 = "Print_Titles" Or 短名 = "_FilterDatabase" Or Left$(短名, 6) = "_xlfn." _
                Or Left$(短名, 6) = "_xlpm." Or 短名 = "Consolidate_Area" Or 短名 = "Criteria" Or 短名 = "Extract" _
                Or 短名 = "Auto_Open" Or 短名 = "Auto_Close" Or 短名 = "Sheet_Title")

        使用 = False
        p = 1
        Do
            p = InStr(p, 全式, 短名, vbTextCompare)
            If p = 0 Then Exit Do
            前後 = ""
            If p > 1 Then 前後 = Mid$(全式, p - 1, 1)
            j = p + Len(短名)
            If セルの字(前後) = "" Or Not ((前後 >= "A" And 前後 <= "z") Or (前後 >= "0" And 前後 <= "9") Or 前後 = "_") Then
                前後 = ""
                If j <= Len(全式) Then 前後 = Mid$(全式, j, 1)
                If セルの字(前後) = "" Or Not ((前後 >= "A" And 前後 <= "z") Or (前後 >= "0" And 前後 <= "9") Or 前後 = "_") Then
                    使用 = True
                    Exit Do
                End If
            End If
            p = p + 1
        Loop

        sh.Cells(outRow, 1).Value = "'" & 名前
        sh.Cells(outRow, 2).Value = "'" & 参照
        If InStr(参照, "#REF!") > 0 Then
            sh.Cells(outRow, 3).Value = "壊れている（#REF!）"
            n壊れ = n壊れ + 1
            On Error Resume Next
            nM.Delete
            If Err.Number = 0 Then
                sh.Cells(outRow, 4).Value = "消しました"
                消した = 消した + 1
            Else
                sh.Cells(outRow, 4).Value = "消せませんでした"
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf 組込 Then
            sh.Cells(outRow, 3).Value = "Excel が使う名前（印刷範囲・絞り込みなど）"
            sh.Cells(outRow, 4).Value = "そのまま"
        ElseIf Not 使用 Then
            sh.Cells(outRow, 3).Value = "どこからも使われていない"
            sh.Cells(outRow, 4).Value = "消してよい候補（消していません）"
            n未使用 = n未使用 + 1
        Else
            sh.Cells(outRow, 3).Value = "使われている"
            sh.Cells(outRow, 4).Value = "そのまま"
        End If
        outRow = outRow + 1
次の名前:
    Next i

    sh.Range(sh.Cells(2, 1), sh.Cells(outRow - 1, 4)).Borders.LineStyle = xlContinuous
    sh.Columns("A:D").AutoFit
    If sh.Columns(2).ColumnWidth > 60 Then sh.Columns(2).ColumnWidth = 60
    sh.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = "名前の整理: 名前 " & n全 & " 本のうち、壊れた名前 " & n壊れ & " 本を消しました（" & 消した & " 本）。使われていない候補が " & n未使用 & " 本（消していません）。一覧はシート「" & 名 & "」。"
End Sub

Sub 余分な行列を削除し軽くする()
    ' 依頼の語: ブックの掃除|軽量化|ファイルが重|開くのに時間|ブックを軽く|重いブック|容量を小さく|ファイルサイズを|動きが遅いブック
    ' 依頼の組: ファイル,ブック,容量+軽く,軽量,小さく+-原因,理由,なぜ
    ' 扱う: 軽量化 ゴースト 図形 条件付き書式
    ' 見出し: なし
    ' 形: 数の列
    ' 重い原因（使われていない末尾の行と列・図形の数・条件付き書式の本数）を数え、
    ' 値も書式も図形も無いゴーストの行と列だけ消す。書式が残っている所は消さずに報告する。
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, ur As Range, CHK As Range
    Dim i As Long, r As Long, c As Long, outRow As Long, 名 As String
    Dim 値行 As Long, 値列 As Long, 図行 As Long, 図列 As Long, 実行 As Long, 実列 As Long
    Dim ur行 As Long, ur列 As Long, ゴ行 As Long, ゴ列 As Long
    Dim shp As Shape, 書式あり As Boolean, v As Variant
    Dim 消行 As Long, 消列 As Long, n図 As Long, n条件 As Long, サイズ As Double

    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False

    名 = "軽量化の診断"
    i = 1
    Do While i < 100
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(名)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        名 = "軽量化の診断" & i: i = i + 1
    Loop
    Set sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = 名

    サイズ = 0
    On Error Resume Next
    If セルの字(wb.path) <> "" Then サイズ = FileLen(wb.FullName) / 1024
    On Error GoTo 0

    sh.Cells(1, 1).Value = "余分な行列を削除し軽くする（" & Format$(Now, "yyyy/m/d h:nn") & "）　掃除前のファイル " & _
                           IIf(サイズ = 0, "未保存", Format$(サイズ, "#,##0") & " KB")
    sh.Cells(1, 1).Font.Bold = True
    sh.Cells(2, 1).Value = "シート"
    sh.Cells(2, 2).Value = "今の範囲"
    sh.Cells(2, 3).Value = "本当の範囲"
    sh.Cells(2, 4).Value = "余分"
    sh.Cells(2, 5).Value = "図形"
    sh.Cells(2, 6).Value = "条件付き書式"
    sh.Cells(2, 7).Value = "どうしたか"
    sh.Range(sh.Cells(2, 1), sh.Cells(2, 7)).Font.Bold = True
    outRow = 3

    For Each ws In wb.Worksheets
        If ws.Name <> 名 Then
            Set ur = Nothing
            On Error Resume Next
            Set ur = ws.UsedRange
            On Error GoTo 0
            If ur Is Nothing Then GoTo 次のシート
            ur行 = ur.Row + ur.rows.Count - 1
            ur列 = ur.Column + ur.Columns.Count - 1

            ' 値と式の最終行・最終列（空の文字を返す式も「使っている」に数える。見た目が空の式の行を余分と見て、
            ' 用意しておいた式ごと消していた・2026-09-24 総点検）
            値行 = 0: 値列 = 0
            Dim 最後 As Range
            Set 最後 = Nothing
            On Error Resume Next
            Set 最後 = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not 最後 Is Nothing Then 値行 = 最後.Row
            Set 最後 = Nothing
            Set 最後 = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            If Not 最後 Is Nothing Then 値列 = 最後.Column
            On Error GoTo 0

            ' 図形の錨
            図行 = 0: 図列 = 0
            n図 = ws.Shapes.Count
            For Each shp In ws.Shapes
                On Error Resume Next
                If shp.BottomRightCell.Row > 図行 Then 図行 = shp.BottomRightCell.Row
                If shp.BottomRightCell.Column > 図列 Then 図列 = shp.BottomRightCell.Column
                On Error GoTo 0
            Next shp

            実行 = 値行: If 図行 > 実行 Then 実行 = 図行
            実列 = 値列: If 図列 > 実列 Then 実列 = 図列
            If 実行 < 1 Then 実行 = 1
            If 実列 < 1 Then 実列 = 1

            n条件 = 0
            On Error Resume Next
            n条件 = ur.FormatConditions.Count
            On Error GoTo 0

            ゴ行 = ur行 - 実行: If ゴ行 < 0 Then ゴ行 = 0
            ゴ列 = ur列 - 実列: If ゴ列 < 0 Then ゴ列 = 0

            sh.Cells(outRow, 1).Value = ws.Name
            sh.Cells(outRow, 2).Value = ur.Address(0, 0)
            sh.Cells(outRow, 3).Value = ws.Range(ws.Cells(1, 1), ws.Cells(実行, 実列)).Address(0, 0)
            sh.Cells(outRow, 4).Value = ゴ行 & " 行 + " & ゴ列 & " 列"
            sh.Cells(outRow, 5).Value = n図
            sh.Cells(outRow, 6).Value = n条件

            If ゴ行 <= 5 And ゴ列 <= 2 Then
                sh.Cells(outRow, 7).Value = "掃除の必要なし"
            Else
                ' 余分な所に書式（塗り・罫線・結合）が残っていないか、まとめて調べる
                書式あり = False
                If ゴ行 > 0 Then
                    Set CHK = ws.Range(ws.Cells(実行 + 1, 1), ws.Cells(ur行, ur列))
                    GoSub 書式を見る
                End If
                If ゴ列 > 0 And Not 書式あり Then
                    Set CHK = ws.Range(ws.Cells(1, 実列 + 1), ws.Cells(ur行, ur列))
                    GoSub 書式を見る
                End If

                If 書式あり Then
                    sh.Cells(outRow, 7).Value = "余分な所に塗り・罫線・結合・入力規則などが残っているので消しませんでした（手で確かめてください）"
                Else
                    On Error Resume Next
                    If ゴ行 > 0 Then
                        ws.Range(ws.rows(実行 + 1), ws.rows(ws.rows.Count)).Clear
                        ws.Range(ws.rows(実行 + 1), ws.rows(ws.rows.Count)).Delete
                        消行 = 消行 + ゴ行
                    End If
                    If ゴ列 > 0 Then
                        ws.Range(ws.Columns(実列 + 1), ws.Columns(ws.Columns.Count)).Clear
                        ws.Range(ws.Columns(実列 + 1), ws.Columns(ws.Columns.Count)).Delete
                        消列 = 消列 + ゴ列
                    End If
                    On Error GoTo 0
                    sh.Cells(outRow, 7).Value = "余分な " & ゴ行 & " 行と " & ゴ列 & " 列を消しました"
                End If
            End If
            outRow = outRow + 1
        End If
次のシート:
    Next ws

    sh.Range(sh.Cells(2, 1), sh.Cells(outRow - 1, 7)).Borders.LineStyle = xlContinuous
    sh.Columns("A:G").AutoFit
    If sh.Columns(7).ColumnWidth > 60 Then sh.Columns(7).ColumnWidth = 60
    sh.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = "余分な行列を削除し軽くする: 余分な " & 消行 & " 行と " & 消列 & " 列を消しました。掃除前のファイルは " & _
                           IIf(サイズ = 0, "未保存", Format$(サイズ, "#,##0") & " KB") & "。上書き保存すると小さくなります（診断はシート「" & 名 & "」）。"
    Exit Sub

書式を見る:
    ' 余分に見える所に、塗り・結合・罫線・入力規則・条件付き書式が残っていれば「書式あり」＝消さない（罫線だけの入力欄を
    ' 余分と見て消していた・2026-09-24 総点検）
    Dim 辺 As Variant, 規則 As Range
    On Error Resume Next
    v = xlNone
    v = CHK.Interior.Pattern
    If IsNull(v) Then
        書式あり = True
    ElseIf v <> xlNone Then
        書式あり = True
    End If
    v = False
    v = CHK.MergeCells
    If IsNull(v) Then
        書式あり = True
    ElseIf v = True Then
        書式あり = True
    End If
    For Each 辺 In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
        v = xlNone
        v = CHK.Borders(辺).LineStyle
        If IsNull(v) Then
            書式あり = True
        ElseIf v <> xlNone Then
            書式あり = True
        End If
    Next 辺
    Set 規則 = Nothing
    Set 規則 = Intersect(CHK, ws.Cells.SpecialCells(xlCellTypeAllValidation))
    If Not 規則 Is Nothing Then 書式あり = True
    If CHK.FormatConditions.Count > 0 Then 書式あり = True
    On Error GoTo 0
    Return
End Sub

Sub 住所と郵便番号の形をそろえる()
    ' 依頼の語: 住所の整形|住所の表記|番地の書き方|住所の書き方|郵便番号を|住所を整え|住所がバラバラ|郵便番号を整え|郵便番号の形|住所と郵便番号|番地の表記|丁目の書き方|住所の全角半角
    ' 依頼の組: 住所,郵便,番地+-調べ,分割,分け
    ' 扱う: 住所 郵便番号 全角半角
    ' 見出し: なし
    ' 形: 選ぶ列
    ' 住所の列をその場で整える（英数字とハイフンを半角に、空白を1つに）。郵便番号の列は 999-9999 の形にそろえる。
    ' 形が読めない郵便番号・住所も郵便番号も空の行は、値に印を書かずに行番号を状態バーで知らせる。
    ' （2026-09-24 shu「住所は元の列のまま」。前は右に「（整えた）」列を丸ごと挿入していて、同じシートの右の表とメモが
    '   1 列ずれ、名前の範囲が広がって名前を使った式の答えが変わった）
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    Dim hdrRow As Long, r As Long, c As Long, cnt As Long, i As Long
    Dim 住所列 As Long, 郵便列 As Long, 名 As String
    Dim s As String, t As String, ch As String, cd As Long, 前 As String
    Dim n住所 As Long, n郵便 As Long, 要確認 As String, n要確認 As Long

    Set ws = ActiveSheet
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)

    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Or hdrRow >= urBottom Then Exit Sub

    ' 見出しから住所の列・郵便番号の列を探す（無ければ選んでいる列を住所とみなす）
    For c = urLeft To urRight
        名 = CStr(ws.Cells(hdrRow, c).Value)
        If 住所列 = 0 And InStr(名, "住所") > 0 Then 住所列 = c
        If 郵便列 = 0 And (InStr(名, "郵便") > 0 Or InStr(名, "〒") > 0) Then 郵便列 = c
    Next c
    If 住所列 = 0 Then
        住所列 = ActiveCell.Column
        If 住所列 < urLeft Or 住所列 > urRight Then 住所列 = urLeft
    End If

    Application.ScreenUpdating = False

    ' 郵便番号の形をそろえる（列そのものを直す。数字7桁だけ 999-9999 に）
    If 郵便列 > 0 Then
        For r = hdrRow + 1 To urBottom
            If Not ws.Cells(r, 郵便列).HasFormula Then
                s = Trim$(CStr(ws.Cells(r, 郵便列).text))
                ' 数になって先頭の 0 が消えた郵便番号（100001）は 7 桁に戻す（2026-09-24 総点検）
                If VarType(ws.Cells(r, 郵便列).Value) = vbDouble Then
                    If ws.Cells(r, 郵便列).Value >= 0 And ws.Cells(r, 郵便列).Value < 10000000 Then s = Format$(ws.Cells(r, 郵便列).Value, "0000000")
                End If
                t = ""
                For i = 1 To Len(s)
                    ch = Mid$(s, i, 1): cd = AscW(ch)
                    If cd >= &HFF10 And cd <= &HFF19 Then
                        t = t & ChrW(cd - &HFF10 + 48)
                    ElseIf ch >= "0" And ch <= "9" Then
                        t = t & ch
                    End If
                Next i
                If Len(t) = 7 Then
                    If CStr(ws.Cells(r, 郵便列).text) <> Left$(t, 3) & "-" & Mid$(t, 4) Then
                        ws.Cells(r, 郵便列).NumberFormatLocal = "@"
                        ws.Cells(r, 郵便列).Value = Left$(t, 3) & "-" & Mid$(t, 4)
                        n郵便 = n郵便 + 1
                    End If
                End If
            End If
        Next r
    End If

    ' 住所をその場で整える（変わるセルだけ書く・式のセルは触らない）
    For r = hdrRow + 1 To urBottom
        If ws.Cells(r, 住所列).HasFormula Then GoTo 次の行
        s = CStr(ws.Cells(r, 住所列).text)
        If Len(Trim$(s)) = 0 Then
            If 郵便列 > 0 Then
                If Trim$(CStr(ws.Cells(r, 郵便列).text)) = "" And セルの字(ws.Cells(r, urLeft).Value) <> "" Then
                    要確認 = 要確認 & "・" & r
                    n要確認 = n要確認 + 1
                End If
            End If
            GoTo 次の行
        End If
        t = ""
        前 = ""
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1): cd = AscW(ch)
            If cd >= &HFF10 And cd <= &HFF19 Then
                ch = ChrW(cd - &HFF10 + 48)
            ElseIf cd >= &HFF21 And cd <= &HFF3A Then
                ch = ChrW(cd - &HFF21 + 65)
            ElseIf cd >= &HFF41 And cd <= &HFF5A Then
                ch = ChrW(cd - &HFF41 + 97)
            ElseIf cd = &HFF0D Or cd = &H2010 Or cd = &H2011 Or cd = &H2012 Or cd = &H2013 _
                Or cd = &H2014 Or cd = &H2015 Or cd = &H2212 Then
                ch = "-"
            ElseIf cd = &H30FC Then
                ' 長音「ー」は数字と数字のあいだ（1ー2）だけ - に。センター・ハイツの「ー」は残す（2026-09-24 総点検）
                If Right$(t, 1) Like "#" And (Mid$(s, i + 1, 1) Like "#" Or Mid$(s, i + 1, 1) Like "[０-９]") Then ch = "-"
            ElseIf cd = &H3000 Or cd = 160 Or cd = 9 Or cd = 10 Or cd = 13 Then
                ch = " "
            End If
            ' 空白は続けない
            If Not (ch = " " And 前 = " ") Then t = t & ch
            前 = ch
        Next i
        t = Trim$(t)
        If t <> s Then
            ws.Cells(r, 住所列).Value = t
            n住所 = n住所 + 1
        End If
        If 郵便列 > 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 郵便列).text))) <> 8 Then
                要確認 = 要確認 & "・" & r
                n要確認 = n要確認 + 1
            End If
        End If
次の行:
    Next r

    Application.ScreenUpdating = True
    Application.StatusBar = "住所と郵便番号: 住所 " & n住所 & " 件を元の列で整えました。" & _
                           IIf(郵便列 > 0, "郵便番号 " & n郵便 & " 件を 999-9999 にそろえました。", "郵便番号の列は見つかりませんでした。") & _
                           IIf(n要確認 > 0, "要確認の行: " & Mid$(要確認, 2) & "（郵便番号の形が読めない・住所も郵便番号も空）", "")
End Sub

Sub 図形を一覧にし位置をそろえる()
    ' 依頼の語: 図形の整理|図形の位置|図形を整理|図やボタンが散|画像の整理|図形の一覧|図がずれ|ボタンがバラバラ|図形をそろえ|図形が散らか
    ' 依頼の組: 図形,画像,ボタン+整理,そろえ,揃え,位置,一覧+-マクロ
    ' 扱う: 図形 画像 一覧 位置そろえ
    ' 見出し: なし
    ' 形: 数の列
    ' ブック中の図形・画像・ボタンを一覧にする（種類・場所・大きさ・マクロの割り当て）。消さない。
    ' 図形を2つ以上えらんでから撃つと、左の位置と大きさを1つ目にそろえる。
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, shp As Shape
    Dim i As Long, outRow As Long, 名 As String, n図 As Long, n画像 As Long, nマクロ As Long
    Dim 基準左 As Double, 基準幅 As Double, 基準高 As Double, そろえた As Long
    Dim sel As Object, 種類 As String, 割当 As String, 文字 As String, 位置 As String
    Dim pos As Long

    Set wb = ActiveWorkbook

    ' 2つ以上の図形をえらんでいたら、位置と大きさをそろえる
    On Error Resume Next
    Set sel = Selection
    On Error GoTo 0
    If Not sel Is Nothing Then
        If TypeName(sel) = "DrawingObjects" Or TypeName(sel) = "GroupObject" Then
            On Error Resume Next
            If sel.Count >= 2 Then
                基準左 = sel(1).Left: 基準幅 = sel(1).Width: 基準高 = sel(1).Height
                For i = 2 To sel.Count
                    sel(i).Left = 基準左
                    sel(i).Width = 基準幅
                    sel(i).Height = 基準高
                    そろえた = そろえた + 1
                Next i
            End If
            On Error GoTo 0
        End If
    End If

    Application.ScreenUpdating = False
    名 = "図形の一覧"
    i = 1
    Do While i < 100
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(名)
        On Error GoTo 0
        If sh Is Nothing Then Exit Do
        名 = "図形の一覧" & i: i = i + 1
    Loop
    Set sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = 名

    sh.Cells(1, 1).Value = "図形と画像の一覧（" & Format$(Now, "yyyy/m/d h:nn") & "）"
    sh.Cells(1, 1).Font.Bold = True
    sh.Cells(2, 1).Value = "シート"
    sh.Cells(2, 2).Value = "名前"
    sh.Cells(2, 3).Value = "種類"
    sh.Cells(2, 4).Value = "場所"
    sh.Cells(2, 5).Value = "幅×高さ"
    sh.Cells(2, 6).Value = "マクロ"
    sh.Cells(2, 7).Value = "文字"
    sh.Range(sh.Cells(2, 1), sh.Cells(2, 7)).Font.Bold = True
    outRow = 3

    For Each ws In wb.Worksheets
        If ws.Name <> 名 Then
            For Each shp In ws.Shapes
                n図 = n図 + 1
                種類 = "図形"
                On Error Resume Next
                Select Case shp.Type
                    Case 13: 種類 = "画像": n画像 = n画像 + 1
                    Case 8: 種類 = "ボタンなどの部品"
                    Case 3: 種類 = "グラフ"
                    Case 6: 種類 = "グループ"
                    Case 17: 種類 = "文字の箱"
                End Select
                割当 = ""
                割当 = shp.OnAction
                文字 = ""
                If shp.TextFrame2.hasText Then 文字 = shp.TextFrame2.TextRange.text
                位置 = shp.TopLeftCell.Address(0, 0)
                On Error GoTo 0
                If セルの字(割当) <> "" Then
                    nマクロ = nマクロ + 1
                    pos = InStrRev(割当, "!")
                    If pos > 0 Then 割当 = Mid$(割当, pos + 1)
                End If

                sh.Cells(outRow, 1).Value = ws.Name
                sh.Cells(outRow, 2).Value = shp.Name
                sh.Cells(outRow, 3).Value = 種類
                sh.Cells(outRow, 4).Value = 位置
                sh.Cells(outRow, 5).Value = Int(shp.Width) & " × " & Int(shp.Height)
                sh.Cells(outRow, 6).Value = IIf(割当 = "", "", "'" & 割当 & "（消さないこと）")
                sh.Cells(outRow, 7).Value = "'" & Left$(文字, 40)
                outRow = outRow + 1
            Next shp
        End If
    Next ws

    If outRow = 3 Then
        sh.Cells(3, 1).Value = "図形はありません"
        outRow = 4
    End If
    sh.Range(sh.Cells(2, 1), sh.Cells(outRow - 1, 7)).Borders.LineStyle = xlContinuous
    sh.Columns("A:G").AutoFit
    If sh.Columns(7).ColumnWidth > 40 Then sh.Columns(7).ColumnWidth = 40
    sh.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = "図形と画像の整理: " & n図 & " 個（うち画像 " & n画像 & " 個・マクロ付き " & nマクロ & " 個）を一覧にしました。" & _
                           IIf(そろえた > 0, "えらんでいた図形 " & そろえた & " 個の位置と大きさを1つ目にそろえました。", "消していません（消す図は人がえらんでください）。")
End Sub

Sub 隠れた行と列を全部表示する()
    ' 依頼の語: 非表示を全部表示|隠れている行と列を|隠した行と列を|全部表示|再表示|非表示を解除|隠れた行を表示|隠れた列を表示|非表示の行と列を
    ' 依頼の組: 非表示,隠れ,隠し,見えない+全部表示,表示して,表示に戻,元に戻,出して,戻して,解除,再表示,見せて+-一覧,ないか,シートが,どこに
    ' 扱う: 非表示 再表示
    ' 見出し: なし
    ' 形: なし
    ' 今のシートの非表示の行と列を、すべて表示に戻す（shu003「非表示の行列を全表示」を棚へ。2026-09-22）
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    On Error Resume Next
    ActiveWindow.View = xlNormalView
    ActiveSheet.Cells.EntireColumn.Hidden = False
    ActiveSheet.Cells.EntireRow.Hidden = False
End Sub

Sub 選んだ列を非表示にする()
    ' 依頼の語: 列を非表示|列を隠|この列を隠|選んだ列を非表示|列を見えなく
    ' 依頼の組: 列+非表示に,隠して,隠す,見えなく+-行と列,全部,表示に戻,解除,調べ,探
    ' 扱う: 非表示 列
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる列（離れた列もまとめて）を非表示にする。シートの列を全部選んでいるときは何もしない（shu003「選択列の非表示」を棚へ）
    Dim ar As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Columns.Count >= ActiveSheet.Columns.Count Then Exit Sub
    For Each ar In Selection.Areas
        ar.EntireColumn.Hidden = True
    Next ar
End Sub

Sub 文字の先頭の空白を削除する()
    ' 依頼の語: 先頭の空白|頭の空白|前の空白|先頭のスペース|行頭の空白|文字の前の空白|左の空白
    ' 依頼の組: 先頭,頭の,文字の前,左の+空白,スペース+消,削除,取,除
    ' 扱う: 空白除去
    ' 見出し: なし
    ' 形: なし
    ' 文字のセルの先頭の空白（半角・全角・タブ）を取る。数式のセルは触らない。
    ' 2 つ以上のセルを選んでいればその範囲、1 つだけなら今のシートの使っている範囲（shu003「選択範囲の先頭空白を削除」を棚へ）
    Dim rng As Range, c As Range, s As String, t As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) = "Range" And Selection.CountLarge > 1 Then
        Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If VarType(c.Value) = vbString Then
                s = c.Value
                t = s
                Do While Len(t) > 0
                    If Left$(t, 1) = " " Or Left$(t, 1) = "　" Or Left$(t, 1) = vbTab Then t = Mid$(t, 2) Else Exit Do
                Loop
                If t <> s Then c.Value = t
            End If
        End If
    Next c
End Sub

Sub 空白だけのセルを空にする()
    ' 依頼の語: 見えない空白|空に見えるセル|空白だけのセル|スペースだけのセル|空白じゃない空白|空文字|中身が空白だけ
    ' 依頼の組: 見えない,空に見える,スペースだけ,空白だけ,空白じゃない+空白,セル,スペース+-行,列,詰,シート
    ' 扱う: 空白除去 空セル
    ' 見出し: なし
    ' 形: なし
    ' 空に見えるのに中身がある（空白だけ・空の文字）セルを本当の空にする。書式と数式は残す。
    ' 2 つ以上のセルを選んでいればその範囲、1 つだけなら今のシートの使っている範囲（shu003「選択範囲の空白じゃない空白をクリア」を棚へ）
    Dim rng As Range, c As Range, s As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) = "Range" And Selection.CountLarge > 1 Then
        Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If Not c.HasFormula And Not isEmpty(c.Value) Then
            If VarType(c.Value) = vbString Then
                s = Replace(Replace(Replace(c.Value, " ", ""), "　", ""), vbTab, "")
                If セルの字(s) = "" Then c.ClearContents
            End If
        End If
    Next c
End Sub

Sub シートの図形と画像を全部削除する()
    ' 依頼の語: 図形を全部消|図形を全部削除|図を全部消|図を全部削除|画像を全部消|全部の図形を消|図形を消して|図形を削除|図形をすべて消
    ' 依頼の組: 図形,図,画像,写真+全部,すべて,全て+消,削除+-行,セル,選んだ,一覧,整理,そろえ,グラフ
    ' 扱う: 図形 削除
    ' 見出し: なし
    ' 形: なし
    ' 今のシートの図形・画像を消す。マクロの付いたボタンとグラフは残す（shu003「ワークシートの全図削除」を棚へ）
    Dim i As Long, sh As Shape, act As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        Set sh = ActiveSheet.Shapes(i)
        act = ""
        On Error Resume Next
        act = sh.OnAction
        On Error GoTo 0
        If Len(act) = 0 And sh.Type <> msoChart And sh.Type <> msoComment Then sh.Delete
    Next i
End Sub

Sub 選んだ行を図形ごと削除する()
    ' 依頼の語: 図ごと|図形ごと|画像ごと|図と行を削除|図と一緒に行|行を図ごと|写真ごと
    ' 依頼の組: 図,図形,画像,写真+ごと,一緒+行+-列
    ' 扱う: 図形 行削除
    ' 見出し: なし
    ' 形: なし
    ' 選んでいる行を、その行に掛かっている図形・画像ごと消す（マクロの付いたボタンは残す。shu003「選択範囲の図と行を削除」を棚へ）
    Dim i As Long, sh As Shape, rws As Range, act As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.rows.CountLarge > 5000 Then Exit Sub        ' 列ごと選んだまま撃ったときに、シートの全部の行を消さない
    Set rws = Selection.EntireRow
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        Set sh = ActiveSheet.Shapes(i)
        act = ""
        On Error Resume Next
        act = sh.OnAction
        On Error GoTo 0
        If Len(act) = 0 Then
            If Not Intersect(ActiveSheet.Range(sh.TopLeftCell, sh.BottomRightCell), rws) Is Nothing Then sh.Delete
        End If
    Next i
    ' 行は重なりなしに集めて下から消す（2 列を別々に選ぶと同じ行が 2 重になり「重複する選択範囲」で止まった）。
    ' テーブルの見出しの行は消せないので残す（2026-09-22）
    Dim d As Object, ar As Range, rr As Long, lo As ListObject, keep As Boolean, ks As Variant, j As Long, k As Long, tmp As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each ar In Selection.Areas
        For rr = ar.Row To ar.Row + ar.rows.Count - 1
            d(rr) = True
        Next rr
    Next ar
    ks = d.keys
    For j = 0 To UBound(ks) - 1              ' 大きい行番号から並べる
        For k = j + 1 To UBound(ks)
            If ks(k) > ks(j) Then tmp = ks(j): ks(j) = ks(k): ks(k) = tmp
        Next k
    Next j
    For j = 0 To UBound(ks)
        keep = False
        For Each lo In ActiveSheet.ListObjects
            If lo.ShowHeaders Then
                If lo.HeaderRowRange.Row = ks(j) Then keep = True
            End If
        Next lo
        If Not keep Then ActiveSheet.rows(ks(j)).Delete
    Next j
End Sub

Sub 選んだセルの図形を削除する()
    ' 依頼の語: セルの図を消|セルの図を削除|選んだ所の図|範囲の図を消|範囲の図形を消|セルにある図|セルの上の図|ここの図
    ' 依頼の組: 図,図形,画像,写真+選んだ,セル,範囲,ここ+消,削除+-行,全部,すべて,一覧
    ' 扱う: 図形 削除
    ' 見出し: なし
    ' 形: なし
    ' 選んでいるセルに掛かっている図形・画像を消す。セルの中身は触らない（マクロの付いたボタンは残す。shu003「選択セル図削除」を棚へ）
    Dim i As Long, sh As Shape, act As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        Set sh = ActiveSheet.Shapes(i)
        act = ""
        On Error Resume Next
        act = sh.OnAction
        On Error GoTo 0
        If Len(act) = 0 Then
            If Not Intersect(ActiveSheet.Range(sh.TopLeftCell, sh.BottomRightCell), Selection) Is Nothing Then sh.Delete
        End If
    Next i
End Sub

Sub 選んだ文字を1セルにまとめ行削除()
    ' 依頼の語: 1つのセルにまとめ|一つのセルにまとめ|１つのセルにまとめ|つなげて1つに|文字をつなげ|結合して余った行|余った行を削除|改行された行をまとめ
    ' 依頼の組: 1つのセル,一つのセル,１つのセル,つなげ,連結+まとめ,1つに,一つに,削除,余った
    ' 扱う: 連結 行削除
    ' 見出し: なし
    ' 形: なし
    ' 選んでいる範囲の文字を左上のセルに 1 つにつなげ、2 行目から下の行を消す。
    ' 選んでいない列に値がある行は消さずに、選んだセルの中身だけ空ける（隣の列の値を巻き込まない。shu003「選択範囲の結合と行削除」を棚へ）
    Dim sel As Range, arr As Variant, r As Long, c As Long, s As String, rr As Long, cc As Long, other As Boolean
    Dim ur As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection.Areas(1)
    If sel.rows.Count < 2 Or sel.rows.Count > 5000 Then Exit Sub
    If sel.CountLarge = 1 Then Exit Sub
    ' 結合したセル（二段見出しなど）を含むと中身を消せずに実行時エラー 1004 で止まった（2026-09-24 棚の試験 F10）
    Dim mixedMerge As Variant
    mixedMerge = sel.MergeCells                    ' 一部だけ結合なら Null
    If IsNull(mixedMerge) Then mixedMerge = True
    If mixedMerge Then
        Application.StatusBar = "結合したセルを含む範囲はまとめられません（結合のない範囲を選んでください）"
        Exit Sub
    End If
    arr = sel.Value
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If Not IsError(arr(r, c)) Then s = s & CStr(arr(r, c))
        Next c
    Next r
    Set ur = ActiveSheet.UsedRange
    other = False
    For rr = sel.Row + 1 To sel.Row + sel.rows.Count - 1
        For cc = ur.Column To ur.Column + ur.Columns.Count - 1
            If cc < sel.Column Or cc > sel.Column + sel.Columns.Count - 1 Then
                If Not isEmpty(ActiveSheet.Cells(rr, cc).Value) Then other = True
            End If
        Next cc
    Next rr
    sel.Cells(1, 1).Value = s
    If other Then
        sel.Offset(1, 0).Resize(sel.rows.Count - 1).ClearContents
        If sel.Columns.Count > 1 Then sel.rows(1).Offset(0, 1).Resize(1, sel.Columns.Count - 1).ClearContents
    Else
        If sel.Columns.Count > 1 Then sel.rows(1).Offset(0, 1).Resize(1, sel.Columns.Count - 1).ClearContents
        sel.Offset(1, 0).Resize(sel.rows.Count - 1).EntireRow.Delete
    End If
End Sub

Sub 選んだ行を1ページの高さにする()
    ' 依頼の語: 行の高さをそろえ|行の高さを揃え|行高をそろえ|行高を揃え|1ページに収まる高さ|行の高さを同じ|行の高さを均等
    ' 依頼の組: 行の高さ,行高+そろえ,揃え,同じ,合わせ,均等+-前のシート,1枚目,列幅,全シート,シートの
    ' 扱う: 行の高さ
    ' 見出し: なし
    ' 形: なし
    ' 選んでいる行（見えている行だけ）を、1 ページ（805pt）に収まる同じ高さにそろえる（shu003「選択範囲の可視行高さ調整」の既定の高さで、聞かずに撃つ）
    Dim vis As Range, ar As Range, n As Long, h As Double
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.rows.CountLarge < 2 Or Selection.rows.CountLarge > 5000 Then Exit Sub
    On Error Resume Next
    Set vis = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then Exit Sub
    For Each ar In vis.Areas
        n = n + ar.rows.Count
    Next ar
    If n = 0 Then Exit Sub
    h = 805 / n
    If h < 5 Then h = 5
    If h > 409 Then h = 409
    vis.RowHeight = h
End Sub

Sub 前のシートの行高と列幅を写す()
    ' 依頼の語: 前のシートに合わせ|前のシートと同じ|一つ前のシート|前のシートに揃え|前のシートの列幅|左のシートに合わせ
    ' 依頼の組: 前のシート,一つ前,左のシート+合わせ,同じ,揃え,そろえ
    ' 扱う: 列幅 行の高さ
    ' 見出し: なし
    ' 形: なし
    ' 一つ前（左）のシートの行の高さと列幅を、今のシートへ写す。1 枚目のシートでは何もしない（shu003「行高と列幅を一つ前のシートに揃える」を棚へ）
    Dim src As Worksheet, r As Long, c As Long
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Index = 1 Then Exit Sub
    If TypeName(ActiveWorkbook.Sheets(ActiveSheet.Index - 1)) <> "Worksheet" Then Exit Sub
    Set src = ActiveWorkbook.Sheets(ActiveSheet.Index - 1)
    Application.ScreenUpdating = False
    For r = 1 To src.UsedRange.Row + src.UsedRange.rows.Count - 1
        ActiveSheet.rows(r).RowHeight = src.rows(r).RowHeight
    Next r
    For c = 1 To src.UsedRange.Column + src.UsedRange.Columns.Count - 1
        ActiveSheet.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c
    Application.ScreenUpdating = True
End Sub

Sub 一枚目のシートの行高と列幅を写す()
    ' 依頼の語: 1枚目のシートに合わせ|最初のシートに合わせ|1枚目と同じ|先頭のシートに合わせ|一枚目のシート|１枚目のシート|1枚目のシートと同じ
    ' 依頼の組: 1枚目,一枚目,１枚目,最初のシート,先頭のシート+合わせ,同じ,揃え,そろえ
    ' 扱う: 列幅 行の高さ
    ' 見出し: なし
    ' 形: なし
    ' 1 枚目のシートの行の高さと列幅を、今のシートへ写す。1 枚目のシートでは何もしない（shu003「行高と列幅を1枚目のシートに揃える」を棚へ）
    Dim src As Worksheet, r As Long, c As Long
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Index = 1 Then Exit Sub
    If TypeName(ActiveWorkbook.Sheets(1)) <> "Worksheet" Then Exit Sub
    Set src = ActiveWorkbook.Sheets(1)
    Application.ScreenUpdating = False
    For r = 1 To src.UsedRange.Row + src.UsedRange.rows.Count - 1
        ActiveSheet.rows(r).RowHeight = src.rows(r).RowHeight
    Next r
    For c = 1 To src.UsedRange.Column + src.UsedRange.Columns.Count - 1
        ActiveSheet.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c
    Application.ScreenUpdating = True
End Sub

Sub 見出しの語で列に名前を付ける()
    ' 依頼の語: 見出しで名前|見出しの名前で|列に名前|見出しを名前に|範囲に名前を付|名前の定義を作|名前を付けて
    ' 依頼の組: 名前+付け,定義を作,作って+見出し,列,範囲+-整理,消,削除,揺れ,一覧,シート名,ファイル名
    ' 扱う: 名前 名前の定義
    ' 見出し: なし
    ' 形: なし
    ' 表の見出しの語を名前にして、その下の列の範囲に名前を付ける（ブック全体の名前）。名前に使えない字は _ に替える。
    ' 付けられない見出しは飛ばす（shu003「アクティブワークシートの一行目の値で名前をつける」を棚へ。見出しは 1 行目に限らず表の見出しの行）
    Dim ws As Worksheet, ur As Range, hr As Long, r As Long, c As Long, c1 As Long, c2 As Long, n As Long, best As Long
    Dim lastR As Long, nM As String, k As Long, ch As String, t As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur.rows.Count < 2 Then Exit Sub
    For r = ur.Row To ur.Row + IIf(ur.rows.Count > 20, 20, ur.rows.Count) - 1
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, ur.Column), ws.Cells(r, ur.Column + ur.Columns.Count - 1)))
        If n >= 2 Then hr = r: Exit For
    Next r
    If hr = 0 Then Exit Sub
    c1 = ur.Column
    Do While c1 < ur.Column + ur.Columns.Count - 1 And セルの字(ws.Cells(hr, c1).Value) = ""
        c1 = c1 + 1
    Loop
    c2 = c1
    Do While c2 < ur.Column + ur.Columns.Count - 1 And セルの字(ws.Cells(hr, c2 + 1).Value) <> ""
        c2 = c2 + 1
    Loop
    For c = c1 To c2
        t = Trim$(CStr(ws.Cells(hr, c).Value))
        If セルの字(t) <> "" Then
            nM = ""
            For k = 1 To Len(t)
                ch = Mid$(t, k, 1)
                If InStr(" !""#$%&'()*+,-/:;<=>?@[]^`{|}~　", ch) > 0 Then nM = nM & "_" Else nM = nM & ch
            Next k
            If nM Like "#*" Or UCase$(nM) Like "[A-Z]#*" Or UCase$(nM) Like "[A-Z][A-Z]#*" Or UCase$(nM) Like "[A-Z][A-Z][A-Z]#*" Or UCase$(nM) = "R" Or UCase$(nM) = "C" Then nM = "_" & nM
            ' 名前の範囲は表の本文だけ（列の最後の値までにすると、下の合計の行・メモまで入り、SUM(金額) が 2 倍になった・2026-09-24 総点検）
            lastR = 表の本文の最終行(ws)
            Do While lastR > hr And 集計の行か(ws, lastR, c1, c2)
                lastR = lastR - 1
            Loop
            If ws.Cells(ws.rows.Count, c).End(xlUp).Row < lastR Then lastR = ws.Cells(ws.rows.Count, c).End(xlUp).Row
            If lastR > hr Then
                On Error Resume Next
                ws.Parent.names.Add Name:=nM, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Range(ws.Cells(hr + 1, c), ws.Cells(lastR, c)).Address
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next c
End Sub

Sub 選んだ数の分だけ下に行を挿入する()
    ' 依頼の語: 数の分だけ行|数だけ行を挿入|数字の分だけ行|件数分の行|数量分の行|値の数だけ行|数の分だけ行を
    ' 依頼の組: 分だけ,数だけ,件数分,数量分+行+挿入,足,入れ,増や
    ' 扱う: 行挿入
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいるセルから下へ、セルの数（例 3）の分だけ行があるように、下に空の行（例 2 行）を足す。空のセル・数でないセルで止まる。
    ' 1 つのセルで 1000 行まで・全部で 20000 行まで（shu003「空白までセルの値だけ行を挿入」を棚へ）
    Dim c As Range, v As Variant, n As Long, total As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set c = Selection.Cells(1, 1)
    Do
        v = c.Value
        If isEmpty(v) Or IsError(v) Then Exit Do
        If Not IsNumeric(v) Or VarType(v) = vbString Then Exit Do
        If Abs(v) > 1000000 Then Exit Do          ' 金額の列を選んで撃つと CLng が「オーバーフロー」で止まった（2026-09-24 総点検）
        n = CLng(v)
        If n > 1 And n <= 1000 Then
            If total + n - 1 > 20000 Then Exit Do
            c.Offset(1, 0).Resize(n - 1).EntireRow.Insert
            total = total + n - 1
            Set c = c.Offset(n, 0)
        Else
            Set c = c.Offset(1, 0)
        End If
    Loop
End Sub

Sub 選んだ住所の左に郵便番号を書く()
    ' 依頼の語: 郵便番号を引|郵便番号を入れ|住所から郵便番号|郵便番号を調べて入|郵便番号を付け|郵便番号を書き
    ' 依頼の組: 郵便番号,〒+住所から,入れ,引,付け,調べて入,書き+-整え,形,そろえ,揃え,999
    ' 扱う: 郵便番号 住所
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる住所のセルから全国郵便データで郵便番号を引き、それぞれ左隣のセルに書く。
    ' 中身は shu003「選択セルの左側に郵便番号を入力」（全国郵便.xlsm を読み、初回だけ読み込んで以後は瞬時）を呼ぶ。棚にはこの入口だけを置く
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.Run "'" & ThisWorkbook.Name & "'!選択セルの左側に郵便番号を入力"
End Sub

Sub 選んだ列が空欄の行を削除する()
    ' 依頼の語: 空欄の行を削除|空欄の行を消|が空の行を|空欄の行を取り除|未入力の行を削除|未入力の行を消|空いている行を消
    ' 依頼の組: 空欄,空白,未入力,空い+行+削除,消,除+-一覧,探,調べ,チェック,詰,空行,空白行,空白の行
    ' 扱う: 行削除 空欄
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる列が空欄の行を、表（選んだセルの周りのひとまとまり）の中で消す。見出しの行と合計の行は残す（shu003「選択範囲の空白の行を削除」を棚へ）
    Dim ws As Worksheet, reg As Range, col As Long, r As Long, R1 As Long, r2 As Long, s As String, k As Long, sumRow As Boolean
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    col = Selection.Cells(1, 1).Column
    Set reg = Selection.Cells(1, 1).CurrentRegion
    If reg.rows.Count < 2 Then Exit Sub
    If reg.rows.Count > 50000 Then Exit Sub
    ' 二段見出しの下の段も見出し。縦の結合の下側は空に見えるが空欄ではない（見出しの行ごと消していた・2026-09-24 棚の試験 F10）
    R1 = 見出しの下の段(ws, reg.Row, reg.Column, reg.Column + reg.Columns.Count - 1) + 1
    r2 = reg.Row + reg.rows.Count - 1
    For r = r2 To R1 Step -1
        If Trim$(CStr(ws.Cells(r, col).text)) = "" And Not ws.Cells(r, col).MergeCells Then
            sumRow = False
            For k = reg.Column To reg.Column + reg.Columns.Count - 1
                s = Replace(Trim$(CStr(ws.Cells(r, k).text)), " ", "")
                If 集計の語か(s) Then sumRow = True
            Next k
            If Not sumRow Then ws.rows(r).Delete
        End If
    Next r
End Sub

Sub 他ブック参照を自ブックに直す()
    ' 依頼の語: リンク元を|参照先を自分のブック|このブックの参照に|ブック名を消|外部参照をこのブック|コピーしたシートのリンク|リンク元修正|リンク元を修正|リンク先をこのブック
    ' 依頼の組: リンク,参照,外部+このブック,自分のブック,同じブック,付け替,修正,直+-値に,切,解消,調べ,一覧,探,壊れ
    ' 扱う: 外部リンク 参照
    ' 見出し: なし
    ' 形: なし
    ' 今のシートの式から [別ブック.xlsx] の部分を取り、同じ名前の自分のブックのシートを指すように直す（コピーしてきたシートのリンク元直し。
    ' 値にはしない＝値にするのは「外部リンクを値に変えて切る」。shu003「アクティブなワークシートでリンクを解消」を棚へ）
    ' 閉じたブックの参照（'C:\…\[元.xlsx]集計'!B2）もシート名だけに直す。直した先のシートがこのブックに無い式は直さずに数えて言う
    ' （[元.xlsx] だけ消すとパスが残り、Excel がファイルを探す窓を出して止まった・2026-09-24 総点検）
    Dim fc As Range, c As Range, re1 As Object, re2 As Object, reS As Object, f As String, g As String, calc As Long
    Dim m As Object, nM As String, ok As Boolean, sh As Object, n直 As Long, n残 As Long
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    On Error Resume Next
    Set fc = ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fc Is Nothing Then Exit Sub
    Set re1 = CreateObject("VBScript.RegExp"): re1.Global = True: re1.IgnoreCase = True
    re1.Pattern = "'[^'\[]*\[[^\[\]]+\.xl[a-z]{1,2}\]([^']+)'!"
    Set re2 = CreateObject("VBScript.RegExp"): re2.Global = True: re2.IgnoreCase = True
    re2.Pattern = "\[[^\[\]]+\.xl[a-z]{1,2}\]"
    Set reS = CreateObject("VBScript.RegExp"): reS.Global = True
    reS.Pattern = "(?:'((?:[^']|'')+)'|([^\s=+\-*/^&,;:<>()!'""{}]+))!"
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo 戻す
    For Each c In fc.Cells
        f = c.Formula
        If InStr(f, "[") > 0 Then
            g = re2.Replace(re1.Replace(f, "'$1'!"), "")
            If g <> f Then
                ' 直した先のシートがこのブックにあるか（無いのに式を入れると、ファイルを探す窓が出る）
                ok = True
                For Each m In reS.Execute(g)
                    nM = m.SubMatches(0)
                    If Len(nM) = 0 Then nM = m.SubMatches(1)
                    nM = Replace(nM, "''", "'")
                    Set sh = Nothing
                    On Error Resume Next
                    Set sh = ActiveWorkbook.Sheets(nM)
                    On Error GoTo 戻す
                    If sh Is Nothing Then ok = False: Exit For
                Next m
                If ok Then
                    On Error Resume Next
                    c.Formula = g
                    If Err.Number = 0 Then n直 = n直 + 1 Else n残 = n残 + 1
                    On Error GoTo 戻す
                Else
                    n残 = n残 + 1
                End If
            End If
        End If
    Next c
戻す:
    Application.Calculation = calc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Application.StatusBar = "外部参照を直す: 途中で止まりました（" & Err.Description & "）。直した式 " & n直 & " か所"
    Else
        Application.StatusBar = "外部参照を直す: " & n直 & " か所をこのブックのシートの参照に直しました。" & _
            IIf(n残 > 0, "同じ名前のシートがこのブックに無い式 " & n残 & " か所は直していません。", "")
    End If
End Sub

Sub 縦横1ページに収めて印刷設定()
    ' 依頼の語: 1ページに収め|１ページに収め|一ページに収め|1枚に収め|1枚で印刷|1ページで印刷|1ページに印刷|1枚に印刷
    ' 依頼の組: 1ページ,１ページ,一ページ,1枚,１枚,一枚+収め,収まる,印刷,入れ+-調べ,何ページ,何枚,横,縦,幅,見出し
    ' 扱う: 印刷設定
    ' 見出し: なし
    ' 形: なし
    ' 今のシートを、縦も横も 1 ページに収めて印刷する設定にする（印刷はしない。shu003「印刷を1ページに変更」を棚へ）
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub

Sub 番号の列を連番に振り直す()
    ' 依頼の語: 番号を振り直|番号を振りなお|連番を振り直|連番を振りなお|連番にし|番号を付け直|番号をつけ直|番号が飛|番号の抜け|番号が抜け|番号が重複|通し番号を|Noを振り直|No.を振り直|項番を振り直
    ' 依頼の組: 番号,連番,No,通し番号,項番+振り直,振りなお,付け直,つけ直,振って,飛,抜け,詰め,重複,ずれ,1から+-一覧,洗い出,チェック,調べ,探,確認,何件,右,伝票,会員,郵便
    ' 扱う: 番号 連番
    ' 見出し: なし
    ' 形: 番号の列
    ' 見出しが No・番号・連番・通し番号・項番 の列（無ければ左端の、1 から始まる整数の列）を、表の本文の上から 1, 2, 3… に振り直す。
    ' 伝票番号・会員番号のような「名前の付いた番号」は ID なので触らない。式の入った番号列・合計の行・空の行も触らない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, lastRow As Long
    Dim headerRow As Long, dataStart As Long, noCol As Long
    Dim r As Long, c As Long, cnt As Long, n As Long
    Dim h As String, v As Variant, rowHasValue As Boolean, isTotal As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    lastRow = 表の本文の最終行(ws)

    ' 見出しの行＝値が 2 つ以上ある最初の行
    For r = urTop To lastRow
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    dataStart = headerRow + 1
    If dataStart > lastRow Then Exit Sub

    ' 番号の列＝見出しがちょうど番号の語（伝票番号・会員番号は ID なので外す）
    For c = urLeft To urRight
        h = StrConv(Trim$(Replace(CStr(ws.Cells(headerRow, c).Value), "　", "")), vbNarrow)
        h = UCase$(Replace(h, ".", ""))
        If h = "NO" Or h = "№" Or h = "番号" Or h = "連番" Or h = "通し番号" Or h = "項番" Or h = "#" Then
            noCol = c
            Exit For
        End If
    Next c
    ' 見出しで決まらなければ、左端の列が 1 から始まる整数の並びのときだけ
    If noCol = 0 Then
        v = ws.Cells(dataStart, urLeft).Value
        If IsNumeric(v) And Not isEmpty(v) Then
            If CDbl(v) = 1 Then noCol = urLeft
        End If
    End If
    If noCol = 0 Then Exit Sub

    ' 式で番号を出している列（=ROW()-1 など）は直さなくてよい
    For r = dataStart To lastRow
        If ws.Cells(r, noCol).HasFormula Then Exit Sub
    Next r

    ' 今の番号が「001」のような先頭ゼロ付きの文字なら、その桁の形のまま振り直す（数の 1, 2, 3 にしてゼロを消さない・2026-09-23）
    Dim padL As Long, zeroN As Long, filledN As Long, sv As String
    For r = dataStart To lastRow
        sv = Trim$(セルの字(ws.Cells(r, noCol).Value))
        If sv <> "" Then
            filledN = filledN + 1
            If VarType(ws.Cells(r, noCol).Value) = vbString And Len(sv) >= 2 And Left$(sv, 1) = "0" And sv Like String(Len(sv), "#") Then
                zeroN = zeroN + 1
                If Len(sv) > padL Then padL = Len(sv)
            End If
        End If
    Next r
    If zeroN < 2 Or zeroN < filledN / 2 Then padL = 0

    n = 0
    For r = dataStart To lastRow
        rowHasValue = False
        isTotal = False
        For c = urLeft To urRight
            If c <> noCol Then
                h = Trim$(セルの字(ws.Cells(r, c).Value))
                If h <> "" Then rowHasValue = True
                If 集計の語か(h) Then isTotal = True
            End If
        Next c
        If rowHasValue And Not isTotal Then
            n = n + 1
            If padL > 0 Then
                sv = Right$(String(padL, "0") & CStr(n), IIf(Len(CStr(n)) > padL, Len(CStr(n)), padL))
                If セルの字(ws.Cells(r, noCol).Value) <> sv Then
                    ws.Cells(r, noCol).NumberFormat = "@"
                    ws.Cells(r, noCol).Value = sv
                End If
            ElseIf セルの字(ws.Cells(r, noCol).Value) <> CStr(n) Then
                ws.Cells(r, noCol).Value = n
            End If
        End If
    Next r
End Sub

Sub 番号に先頭のゼロを付けてそろえる()
    ' 依頼の語: 先頭のゼロ|先頭の0|頭のゼロ|ゼロが消え|0が消え|ゼロが落ち|0が落ち|桁をそろえ|桁を揃え|ゼロ埋め|0埋め|ゼロを付け|0を付け|桁数をそろえ|桁数を揃え
    ' 依頼の組: ゼロ,0,０,桁+消え,落ち,抜け,そろえ,揃え,埋め,付け,戻+-小数,四捨五入,切り捨て,丸め,円,金額,合計
    ' 扱う: 番号 桁 先頭ゼロ
    ' 見出し: なし
    ' 形: 番号の列
    ' 「0001」のように先頭にゼロの付いた番号が並ぶ列（同じ桁のゼロ付きが 2 つ以上・列の 6 割以上）で、ゼロが消えて「4」になったセルを
    ' 列の桁にそろえて文字の「0004」に戻す。2 つ以上のセルを選んでいればその列だけ、1 つだけなら表の全部の列を見る。式のセルは触らない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, lastRow As Long
    Dim headerRow As Long, dataStart As Long
    Dim r As Long, c As Long, cnt As Long, L As Long, bestL As Long, bestN As Long, filled As Long
    Dim s As String, v As Variant
    Dim lens As Object, useCol As Object, key As Variant

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    lastRow = 表の本文の最終行(ws)

    For r = urTop To lastRow
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    dataStart = headerRow + 1
    If dataStart > lastRow Then Exit Sub

    ' 見る列（2 つ以上のセルを選んでいればその列だけ）
    Set useCol = CreateObject("Scripting.Dictionary")
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.Count > 1 And Selection.Worksheet.Name = ws.Name Then
            Dim a As Range, cc As Range
            For Each a In Selection.Areas
                For Each cc In a.Columns
                    useCol(cc.Column) = True
                Next cc
            Next a
        End If
    End If

    For c = urLeft To urRight
        If useCol.Count = 0 Or useCol.Exists(c) Then
            ' 列の中のゼロ付きの番号の桁を数える
            Set lens = CreateObject("Scripting.Dictionary")
            filled = 0
            For r = dataStart To lastRow
                v = ws.Cells(r, c).Value
                s = Trim$(セルの字(v))
                If s <> "" And Not ws.Cells(r, c).HasFormula Then
                    filled = filled + 1
                    If VarType(v) = vbString And Len(s) >= 2 And Left$(s, 1) = "0" And s Like String(Len(s), "#") Then
                        lens(Len(s)) = lens(Len(s)) + 1
                    End If
                End If
            Next r
            bestL = 0: bestN = 0
            For Each key In lens.keys
                If lens(key) > bestN Then
                    bestN = lens(key)
                    bestL = key
                End If
            Next key
            ' ゼロ付きの番号の列＝同じ桁のゼロ付きが 2 つ以上・列の 6 割以上がその桁の数字
            If bestN >= 2 Then
                cnt = 0
                For r = dataStart To lastRow
                    s = Trim$(セルの字(ws.Cells(r, c).Value))
                    If Len(s) = bestL And s Like String(bestL, "#") Then cnt = cnt + 1
                Next r
                If cnt >= filled * 0.6 Then
                    For r = dataStart To lastRow
                        If Not ws.Cells(r, c).HasFormula Then
                            v = ws.Cells(r, c).Value
                            s = Trim$(セルの字(v))
                            If s <> "" And Len(s) < bestL And s Like String(Len(s), "#") Then
                                ws.Cells(r, c).NumberFormat = "@"
                                ws.Cells(r, c).Value = Right$(String(bestL, "0") & s, bestL)
                            End If
                        End If
                    Next r
                End If
            End If
        End If
    Next c
End Sub

Sub 空白の有無を多い方にそろえる()
    ' 依頼の語: 空白の揺れ|スペースの揺れ|空白の有無|スペースの有無|空白がある|スペースがある|姓と名の間|名前の揺れをそろえ|名前の書き方をそろえ|担当者名をそろえ|空白をそろえ|スペースをそろえ
    ' 依頼の組: 空白,スペース,すきま,姓と名+揺れ,ゆれ,有無,ある,ない,入って,そろえ,揃え,統一,多い方+-消し,削除,取っ,取り除,先頭,末尾,空に,見えない,一覧,洗い出,全角,半角,1つ,一つ,１つ
    ' 扱う: 表記ゆれ 空白 名寄せ
    ' 見出し: なし
    ' 形: 文字の列
    ' 空白（半角・全角）の有無と数だけが違う書き方（「佐藤 一郎」「佐藤一郎」「佐藤　一郎」）を、列ごとに多い方の書き方にそろえる。
    ' 同じ数なら空白の無い方。2 つ以上のセルを選んでいればその列だけ、1 つだけなら表の全部の列を見る。式・数・日付のセルは触らない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, lastRow As Long
    Dim headerRow As Long, dataStart As Long
    Dim r As Long, c As Long, cnt As Long
    Dim s As String, k As String, best As String, bestN As Long
    Dim forms As Object, useCol As Object, f As Object, key As Variant, key2 As Variant

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.Row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    lastRow = 表の本文の最終行(ws)

    For r = urTop To lastRow
        cnt = 0
        For c = urLeft To urRight
            If セルの字(ws.Cells(r, c).Value) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    dataStart = headerRow + 1
    If dataStart > lastRow Then Exit Sub

    Set useCol = CreateObject("Scripting.Dictionary")
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.Count > 1 And Selection.Worksheet.Name = ws.Name Then
            Dim a As Range, cc As Range
            For Each a In Selection.Areas
                For Each cc In a.Columns
                    useCol(cc.Column) = True
                Next cc
            Next a
        End If
    End If

    For c = urLeft To urRight
        If useCol.Count = 0 Or useCol.Exists(c) Then
            ' 空白を抜いた字 → {書き方: 件数}
            Set forms = CreateObject("Scripting.Dictionary")
            For r = dataStart To lastRow
                If Not ws.Cells(r, c).HasFormula And VarType(ws.Cells(r, c).Value) = vbString Then
                    s = ws.Cells(r, c).Value
                    k = Replace(Replace(s, " ", ""), "　", "")
                    If k <> "" Then
                        If Not forms.Exists(k) Then forms.Add k, CreateObject("Scripting.Dictionary")
                        Set f = forms(k)
                        f(s) = f(s) + 1
                    End If
                End If
            Next r
            For Each key In forms.keys
                Set f = forms(key)
                If f.Count >= 2 Then
                    best = "": bestN = 0
                    For Each key2 In f.keys
                        If f(key2) > bestN Or (f(key2) = bestN And Len(key2) < Len(best)) Then
                            bestN = f(key2)
                            best = key2
                        End If
                    Next key2
                    For r = dataStart To lastRow
                        If Not ws.Cells(r, c).HasFormula And VarType(ws.Cells(r, c).Value) = vbString Then
                            s = ws.Cells(r, c).Value
                            If s <> best And Replace(Replace(s, " ", ""), "　", "") = key Then
                                ' そろえた字が数・日付に見えるなら文字のまま入れる（「1 2」→「12」が数に化けない）
                                If IsNumeric(best) Or IsDate(best) Then ws.Cells(r, c).NumberFormat = "@"
                                ws.Cells(r, c).Value = best
                            End If
                        End If
                    Next r
                End If
            Next key
        End If
    Next c
End Sub
