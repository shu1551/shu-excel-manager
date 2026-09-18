Attribute VB_Name = "表の整理"
' 表の整理 - アクティブシートの一覧表（見出し 1 行＋本文）を直す 2 本（2026-09-11 夜・09-12 深夜に強化）
'   表を整える   … 値の意味は変えずに書き方をそろえる。行は消さない。数式のセルには触らない。
'                   前後の空白・全角の英数字→半角・半角カナ→全角・文字の間の空白（列の多い方の 1 つに）・
'                   会社名の（株）㈱→株式会社・電話 03-1234-5678（+81 も）・郵便 000-0000・メールは小文字・
'                   文字の日付／和暦／英語／8 桁／数値のままの日付→日付（yyyy/m/d）・文字の金額／「12個」／△▲(　)の負の数→数値（#,##0）・
'                   5%・５％→率（0%）・セルの中の改行→空白・フリガナのひらがな→カタカナ・カタカナの後ろの - → ー（長音）・
'                   状態の列の言い換え（送り仮名の有無・ひらがな書き → 漢字の多い形）・
'                   本文の文字の書式（見出しの書体と大きさ・太字なし・色は自動）・塗りのばらつき・罫線・列幅
'   重複行を消す … 全部の列が同じ行（書き方の違いはならして比べる）の 2 行目以降を下から消す。
'                   1 列でも中身が違えば別の行として残す。会社名の前株・後株は区別する。途中の空行は消さない。表の外（右の列）には触らない。
'   表の範囲 … 見出しの下から、空行が 2 行続く手前まで（途中の空行 1 行は表の続き）

Sub 表を整える()
    Dim ws As Worksheet, ur As Range, tb As Range, c As Range, dv As Object
    Dim r0 As Long, c0 As Long, nr As Long, nc As Long, scanTo As Long
    Dim hr As Long, hc1 As Long, hc2 As Long, lastR As Long
    Dim i As Long, j As Long, k As Long, n As Long, best As Long, ln As Long, a As Long, b As Long, g As Long
    Dim v As Variant, nv As Variant, x As Variant, parts As Variant, ks As Variant
    Dim s As String, t As String, u As String, ch As String, hd As String, kind As String, sp As String
    Dim digits As String, kanaRun As String, mon As String
    Dim cnt As Long, dn As Long, nn As Long, cn As Long, zs As Long, hs As Long, an As Long, pn As Long
    Dim ok As Boolean, bad As Boolean, dt As Date, dbl As Double
    Dim era As Long, y As Long, m As Long, d As Long, cd As Long
    Dim ar(1 To 200) As String, al(1 To 200) As Long, bl(1 To 200) As Long, na As Long
    Dim grp() As Long, k1() As String, k2() As String, rep() As String, hira() As Long, cntOf() As Long
    Dim hasKana() As Boolean, hasKanji() As Boolean
    Dim mainEnd As Long, sumR1 As Long, sumC1 As Long, sumC2 As Long, gapR As Long, rr As Long
    Dim colW() As Double

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    r0 = ur.Row: c0 = ur.Column: nr = ur.rows.Count: nc = ur.Columns.Count
    If nr < 2 Then Exit Sub
    GoSub 見出しを探す
    If hr = 0 Or lastR <= hr Then Exit Sub
    GoSub 集計欄を分ける      ' 空行の下の細い塊（合計・平均）で表を終える。その先の行（注記など）は触らない（2026-09-13）

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
    Next

    ' 本文の文字の書式（見出しの書体と大きさ・太字／斜体／下線／取り消し線なし・文字色は自動）・塗りのばらつき・罫線・列幅
    Set tb = ws.Range(ws.Cells(hr + 1, hc1), ws.Cells(lastR, hc2))
    With tb.Font
        .Name = ws.Cells(hr, hc1).Font.Name
        .Size = ws.Cells(hr, hc1).Font.Size
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .ColorIndex = xlColorIndexAutomatic
    End With
    If IsNull(tb.Interior.ColorIndex) Then tb.Interior.ColorIndex = xlNone
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
    ' 列幅: 表の中で合わせる。細くした列で表の外のセル（下の注記・D31 など）が ### になったら元の幅に戻す（2026-09-13）
    ReDim colW(hc1 To hc2)
    For k = hc1 To hc2
        colW(k) = ws.Columns(k).ColumnWidth
    Next
    ws.Range(ws.Cells(hr, hc1), ws.Cells(lastR, hc2)).Columns.AutoFit
    For k = hc1 To hc2
        If ws.Columns(k).ColumnWidth < colW(k) Then
            For rr = r0 To r0 + nr - 1
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
    Exit Sub

見出しを探す:
    ' 先頭 20 行で、値のある列が最も多い行の 6 割以上を最初に満たす行＝見出し。本文は見出しの列が全部空の行の手前まで
    scanTo = r0 + IIf(nr > 20, 20, nr) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If n > best Then best = n
    Next
    hr = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If hr = 0 And n >= 2 And n >= best * 0.6 Then hr = i
    Next
    If hr = 0 Then Return
    hc1 = 0: hc2 = 0
    For j = c0 To c0 + nc - 1
        If Len(Trim(ws.Cells(hr, j).text)) > 0 Then
            If hc1 = 0 Then hc1 = j
            hc2 = j
        End If
    Next
    ' 本文の終わり: 途中の空行 1 行は飛ばして続ける。空行が 2 行続いたら表の終わり
    lastR = hr
    n = 0
    For i = hr + 1 To r0 + nr - 1
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
            If v <> Int(v) Then ws.Cells(i, j).NumberFormat = "#,##0.0#"
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
    cnt = 0: dn = 0: nn = 0: cn = 0: zs = 0: hs = 0: na = 0: an = 0: pn = 0
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
                            For k = 1 To na
                                If ar(k) = parts(0) Then ok = True
                            Next
                            If Not ok And na < 200 Then
                                na = na + 1: ar(na) = parts(0): al(na) = Len(parts(1)): bl(na) = Len(parts(2))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    If kind = "" Then
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
        For k = 1 To na
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

Sub 重複行を消す()
    Dim ws As Worksheet, ur As Range, seen As Object
    Dim r0 As Long, c0 As Long, nr As Long, nc As Long, scanTo As Long
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
    r0 = ur.Row: c0 = ur.Column: nr = ur.rows.Count: nc = ur.Columns.Count
    If nr < 3 Then Exit Sub
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
    For i = lastR To hr + 1 Step -1
        If dup(i) Then ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2)).Delete Shift:=xlUp
    Next
    Exit Sub

見出しを探す:
    scanTo = r0 + IIf(nr > 20, 20, nr) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If n > best Then best = n
    Next
    hr = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If hr = 0 And n >= 2 And n >= best * 0.6 Then hr = i
    Next
    If hr = 0 Then Return
    hc1 = 0: hc2 = 0
    For j = c0 To c0 + nc - 1
        If Len(Trim(ws.Cells(hr, j).text)) > 0 Then
            If hc1 = 0 Then hc1 = j
            hc2 = j
        End If
    Next
    ' 本文の終わり: 途中の空行 1 行は飛ばして続ける。空行が 2 行続いたら表の終わり
    lastR = hr
    n = 0
    For i = hr + 1 To r0 + nr - 1
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

Sub 合計行を足す()
    ' 依頼の語: 合計行を付|合計行を追加|列ごとの合計|縦計を|合計の行を入|合計の行をお願|最終行の下に列|下に出
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
    urBottom = ur.Row + ur.rows.Count - 1
    urRight = ur.Column + ur.Columns.Count - 1

    Dim headerRow As Long
    headerRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub

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
            If ws.Cells(rr, cc).Value <> "" Then
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
        If ws.Cells(effectiveBottom, c).Value <> "" Then
            checkVal = CStr(ws.Cells(effectiveBottom, c).Value)
            Exit For
        End If
    Next c
    Dim s As String
    s = checkVal: GoSub 正規化: checkVal = s
    If checkVal = "合計" Or checkVal = "計" Or checkVal = "総計" Or checkVal = "小計" Then
        sumRow = effectiveBottom
        lastDataRow = effectiveBottom - 1
    End If

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
        s = CStr(ws.Cells(headerRow, c).Value)
        ' 全角英字を半角にして比較するが、カタカナはそのまま
        Dim hs As String
        hs = s
        ' 全角英数のみ半角化（個別Replace）
        hs = Replace(hs, "Ａ", "A"): hs = Replace(hs, "Ｂ", "B"): hs = Replace(hs, "Ｃ", "C")
        hs = Replace(hs, "Ｎ", "N"): hs = Replace(hs, "Ｏ", "O"): hs = Replace(hs, "ｏ", "o")
        hs = Replace(hs, "ａ", "a"): hs = Replace(hs, "ｂ", "b"): hs = Replace(hs, "ｃ", "c")
        hs = Trim(hs)
        If InStr(hs, "番号") > 0 Or InStr(hs, "NO") > 0 Or InStr(hs, "No") > 0 Or _
           InStr(hs, "コード") > 0 Or InStr(hs, "CODE") > 0 Or InStr(hs, "年度") > 0 Then
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
                    If cv <> "" Then
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
                            If ws.Cells(r, c).Value <> "" And IsNumeric(ws.Cells(r, c).Value) Then
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
                If ws.Cells(r, c).Value <> "" Then
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
                If (maxV - minV) <= rowSpan * 2 And maxV - minV >= 0 Then
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

    ' 合計行をクリア
    For c = urLeft To urRight
        ws.Cells(sumRow, c).ClearContents
    Next c

    If goukeCol > 0 Then
        ws.Cells(sumRow, goukeCol).Value = "合計"
    End If

    For c = urLeft To urRight
        If colType(c) = "num" Then
            ws.Cells(sumRow, c).Formula = "=SUM(" & ws.Cells(dataStart, c).Address & ":" & ws.Cells(lastDataRow, c).Address & ")"
        End If
    Next c

    Exit Sub

正規化:
    s = Trim$(s)
    Return

End Sub

Sub 帳票を一覧に直す()
    ' 依頼の語: 帳票を一覧|結合セルの帳票|受付簿をリスト|様式の帳票を一覧|1件複数行の帳票|帳票から一覧|一覧に直|受付簿|一覧化|帳票形式
    ' 扱う: 帳票変換 一覧化
    ' 見出し: なし
    ' 形: 数の列 日付の列 結合の見出し
    ' 結合セル・1件複数行の帳票を読み取り、表の右に1列空けて1行1件の一覧表を出力する

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.Row
    urLeft = ur.Column
    urBottom = ur.Row + ur.rows.Count - 1
    urRight = ur.Column + ur.Columns.Count - 1

    Dim r As Long, c As Long

    ' --- 見出し行を探す ---
    Dim hdrRow As Long
    hdrRow = 0

    Dim r2 As Long
    For r2 = urTop To Application.Min(urTop + 20, urBottom)
        Dim ne As Long
        ne = 0
        Dim numCnt As Long
        numCnt = 0
        For c = urLeft To urRight
            Dim cv As String
            cv = Trim(CStr(ws.Cells(r2, c).Value))
            If cv <> "" Then
                ne = ne + 1
                If IsNumeric(cv) Or IsDate(cv) Then numCnt = numCnt + 1
            End If
        Next c
        If ne >= 2 Then
            If numCnt >= ne \ 2 + 1 Then
                If r2 - 1 >= urTop Then hdrRow = r2 - 1
            Else
                hdrRow = r2
            End If
            Exit For
        End If
    Next r2
    If hdrRow = 0 Then Exit Sub

    ' --- 帳票右端列を探す ---
    ' 2 段見出しは結合セル・枠線で組まれている。前に作った一覧の見出しは 1 行・結合も枠線も無い＝帳票ではない
    Dim frmRight As Long
    frmRight = urLeft - 1
    Dim bi As Long
    For c = urLeft To urRight
        Dim isFrm As Boolean
        isFrm = False
        If ws.Cells(hdrRow, c).MergeCells Then isFrm = True
        If Not isFrm Then
            Dim bdCnt As Long
            bdCnt = 0
            For bi = 7 To 10
                If ws.Cells(hdrRow, c).Borders(bi).LineStyle <> -4142 Then bdCnt = bdCnt + 1
            Next bi
            If bdCnt >= 3 Then isFrm = True
        End If
        If isFrm Then
            frmRight = c
        ElseIf frmRight >= urLeft Then
            Exit For
        End If
    Next c
    If frmRight < urLeft Then
        ' 結合も枠線も無い帳票: 見出し行の値が続く所までを帳票とする
        Dim gapCnt As Long
        gapCnt = 0
        For c = urLeft To urRight
            If Trim(CStr(ws.Cells(hdrRow, c).Value)) <> "" Then
                If gapCnt >= 1 And frmRight >= urLeft Then Exit For
                gapCnt = 0
                frmRight = c
            ElseIf frmRight >= urLeft Then
                gapCnt = gapCnt + 1
                If gapCnt >= 2 Then Exit For
            End If
        Next c
    End If
    If frmRight < urLeft Then frmRight = urLeft

    ' --- 下段見出し行があるか（帳票の列だけで見る。右に残った前の一覧の数字に惑わされない）---
    Dim sub1Row As Long
    sub1Row = 0
    If hdrRow + 1 <= urBottom Then
        Dim subNE As Long, mainNE As Long
        subNE = 0: mainNE = 0
        Dim subNumCnt As Long
        subNumCnt = 0
        For c = urLeft To frmRight
            Dim hvM As String
            hvM = Trim(CStr(ws.Cells(hdrRow, c).Value))
            If hvM <> "" Then mainNE = mainNE + 1
            Dim hvS As String
            hvS = Trim(CStr(ws.Cells(hdrRow + 1, c).Value))
            If hvS <> "" Then
                subNE = subNE + 1
                If IsNumeric(hvS) Or IsDate(hvS) Then subNumCnt = subNumCnt + 1
            End If
        Next c
        If subNE > 0 And subNE < mainNE And subNumCnt = 0 Then
            sub1Row = hdrRow + 1
        End If
    End If

    Dim dataStartRow As Long
    If sub1Row > 0 Then
        dataStartRow = sub1Row + 1
    Else
        dataStartRow = hdrRow + 1
    End If

    ' --- 列ヘッダを組み立てる ---
    Dim colHeaders() As String
    ReDim colHeaders(urLeft To frmRight)
    Dim lastHdrVal As String
    lastHdrVal = ""
    For c = urLeft To frmRight
        Dim hv3 As String
        hv3 = Trim(CStr(ws.Cells(hdrRow, c).Value))
        If hv3 <> "" Then lastHdrVal = hv3
        colHeaders(c) = lastHdrVal
    Next c
    If sub1Row > 0 Then
        For c = urLeft To frmRight
            Dim sv4 As String
            sv4 = Trim(CStr(ws.Cells(sub1Row, c).Value))
            If sv4 <> "" Then colHeaders(c) = sv4
        Next c
    End If

    ' --- 列名リスト（重複除去・順序保持）---
    Dim listColNames() As String
    ReDim listColNames(0)
    Dim listColDict As Object
    Set listColDict = CreateObject("Scripting.Dictionary")
    Dim listColCnt As Long
    listColCnt = 0
    For c = urLeft To frmRight
        Dim hn As String
        hn = colHeaders(c)
        If hn <> "" Then
            If Not listColDict.Exists(hn) Then
                listColDict.Add hn, listColCnt
                ReDim Preserve listColNames(listColCnt)
                listColNames(listColCnt) = hn
                listColCnt = listColCnt + 1
            End If
        End If
    Next c

    ' --- 見出し語辞書（繰り返し見出し行スキップ用）---
    Dim hdrKwDict As Object
    Set hdrKwDict = CreateObject("Scripting.Dictionary")
    For c = urLeft To frmRight
        Dim cv2 As String
        cv2 = Trim(CStr(ws.Cells(hdrRow, c).Value))
        If cv2 <> "" And Not hdrKwDict.Exists(cv2) Then hdrKwDict.Add cv2, 1
    Next c
    If sub1Row > 0 Then
        For c = urLeft To frmRight
            cv2 = Trim(CStr(ws.Cells(sub1Row, c).Value))
            If cv2 <> "" And Not hdrKwDict.Exists(cv2) Then hdrKwDict.Add cv2, 1
        Next c
    End If

    ' --- 表題テキスト辞書 ---
    Dim titleDict As Object
    Set titleDict = CreateObject("Scripting.Dictionary")
    Dim ttr As Long
    For ttr = urTop To hdrRow - 1
        Dim tNE As Long
        tNE = 0
        Dim tval As String
        tval = ""
        For c = urLeft To urRight
            Dim tv As String
            tv = Trim(CStr(ws.Cells(ttr, c).Value))
            If tv <> "" Then tNE = tNE + 1: tval = tv
        Next c
        If tNE = 1 And tval <> "" And Not titleDict.Exists(tval) Then titleDict.Add tval, 1
    Next ttr

    ' --- 主キー列（帳票左端）---
    Dim mainCol As Long
    mainCol = urLeft

    ' --- レコード収集 ---
    Dim recVals() As Object
    ReDim recVals(0)
    Dim curIdx As Long
    curIdx = 0
    Dim hasCur As Boolean
    hasCur = False

    For r = dataStartRow To urBottom
        Dim rowEmpty As Boolean
        rowEmpty = True
        For c = urLeft To frmRight
            If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then rowEmpty = False: Exit For
        Next c
        If rowEmpty Then GoTo NextRow

        Dim rowNE As Long
        rowNE = 0
        Dim rowVals() As String
        ReDim rowVals(frmRight - urLeft)
        For c = urLeft To frmRight
            Dim cv3 As String
            cv3 = Trim(CStr(ws.Cells(r, c).Value))
            rowVals(c - urLeft) = cv3
            If cv3 <> "" Then rowNE = rowNE + 1
        Next c

        ' ※注記行スキップ
        Dim firstNonEmpty As String
        firstNonEmpty = ""
        For c = urLeft To frmRight
            firstNonEmpty = rowVals(c - urLeft)
            If firstNonEmpty <> "" Then Exit For
        Next c
        If Len(firstNonEmpty) > 0 Then
            If AscW(Left(firstNonEmpty, 1)) = 8251 Then GoTo NextRow
        End If

        ' 繰り返し見出し行スキップ（全非空値が見出し語）
        If rowNE > 0 Then
            Dim allKw As Boolean
            allKw = True
            For c = urLeft To frmRight
                Dim cv4 As String
                cv4 = rowVals(c - urLeft)
                If cv4 <> "" Then
                    If Not hdrKwDict.Exists(cv4) Then allKw = False: Exit For
                End If
            Next c
            If allKw And rowNE >= 2 Then GoTo NextRow
        End If

        ' 繰り返し表題スキップ
        If rowNE = 1 Then
            Dim oneVal As String
            oneVal = ""
            For c = urLeft To frmRight
                If rowVals(c - urLeft) <> "" Then oneVal = rowVals(c - urLeft): Exit For
            Next c
            If titleDict.Exists(oneVal) Then GoTo NextRow
        End If

        ' 合計行スキップ
        Dim isSumRow As Boolean
        isSumRow = False
        For c = urLeft To frmRight
            Dim cvS As String
            cvS = rowVals(c - urLeft)
            cvS = Replace(Replace(cvS, " ", ""), "　", "")
            If cvS = "合計" Or cvS = "小計" Or cvS = "総計" Or cvS = "計" Then
                isSumRow = True: Exit For
            End If
        Next c
        If isSumRow Then GoTo NextRow

        Dim leftVal As String
        leftVal = rowVals(mainCol - urLeft)

        If leftVal <> "" Then
            curIdx = curIdx + 1
            ReDim Preserve recVals(curIdx)
            Set recVals(curIdx) = CreateObject("Scripting.Dictionary")
            hasCur = True
            For c = urLeft To frmRight
                Dim hN2 As String
                hN2 = colHeaders(c)
                If hN2 <> "" Then
                    If Not recVals(curIdx).Exists(hN2) Then
                        recVals(curIdx).Add hN2, ws.Cells(r, c).Value
                    End If
                End If
            Next c
        Else
            If hasCur Then
                Dim sc As Long
                sc = urLeft
                Do While sc <= frmRight
                    Dim scv As String
                    scv = rowVals(sc - urLeft)
                    If scv <> "" And Not IsNumeric(scv) And Not IsDate(scv) Then
                        Dim nsc As Long
                        nsc = sc + 1
                        Dim foundVal As Boolean
                        foundVal = False
                        Do While nsc <= frmRight
                            Dim nvvS As String
                            nvvS = Trim(CStr(ws.Cells(r, nsc).Value))
                            If nvvS <> "" Then
                                foundVal = True
                                Exit Do
                            End If
                            nsc = nsc + 1
                        Loop
                        If foundVal Then
                            Dim lname As String
                            lname = scv
                            If Not listColDict.Exists(lname) Then
                                listColDict.Add lname, listColCnt
                                ReDim Preserve listColNames(listColCnt)
                                listColNames(listColCnt) = lname
                                listColCnt = listColCnt + 1
                            End If
                            If Not recVals(curIdx).Exists(lname) Then
                                recVals(curIdx).Add lname, ws.Cells(r, nsc).Value
                            End If
                            sc = nsc + 1
                        Else
                            sc = sc + 1
                        End If
                    Else
                        sc = sc + 1
                    End If
                Loop
            End If
        End If
NextRow:
    Next r

    If curIdx = 0 Then Exit Sub

    ' --- 出力先: 帳票右端+2列, hdrRow行から ---
    Dim listStartCol As Long
    listStartCol = frmRight + 2

    On Error Resume Next
    Dim wbUR As Range
    Set wbUR = ws.UsedRange
    Dim clearRight As Long
    clearRight = wbUR.Column + wbUR.Columns.Count - 1
    If listStartCol <= clearRight Then
        Dim clearBottom As Long
        clearBottom = wbUR.Row + wbUR.rows.Count - 1
        Dim clearR As Range
        Set clearR = ws.Range(ws.Cells(urTop, listStartCol), ws.Cells(clearBottom, clearRight))
        clearR.MergeCells = False
        clearR.ClearContents
        Set wbUR = ws.UsedRange
        clearRight = wbUR.Column + wbUR.Columns.Count - 1
        If clearRight >= listStartCol Then
            clearBottom = wbUR.Row + wbUR.rows.Count - 1
            ws.Range(ws.Cells(urTop, listStartCol), ws.Cells(clearBottom, clearRight)).MergeCells = False
            ws.Range(ws.Cells(urTop, listStartCol), ws.Cells(clearBottom, clearRight)).ClearContents
        End If
    End If
    On Error GoTo 0

    ' 見出し書き出し
    Dim i As Long
    For i = 0 To listColCnt - 1
        ws.Cells(hdrRow, listStartCol + i).NumberFormat = "@"
        ws.Cells(hdrRow, listStartCol + i).Value = listColNames(i)
    Next i

    ' データ書き出し
    Dim outRow As Long
    outRow = hdrRow + 1
    For i = 1 To curIdx
        For c = 0 To listColCnt - 1
            Dim colName As String
            colName = listColNames(c)
            Dim wv As Variant
            wv = ""
            If recVals(i).Exists(colName) Then wv = recVals(i)(colName)
            Dim tgt As Range
            Set tgt = ws.Cells(outRow, listStartCol + c)
            If isEmpty(wv) Or CStr(wv) = "" Then
                tgt.NumberFormat = "General"
                tgt.Value = ""
            ElseIf VarType(wv) = vbDate Then
                tgt.NumberFormat = "General"
                tgt.Value = wv
            ElseIf VarType(wv) = vbDouble Or VarType(wv) = vbLong Or VarType(wv) = vbInteger Or VarType(wv) = vbSingle Or VarType(wv) = vbCurrency Then
                tgt.NumberFormat = "General"
                tgt.Value = wv
            Else
                Dim svw As String
                svw = CStr(wv)
                If IsNumeric(svw) And Left(svw, 1) <> "0" Then
                    tgt.NumberFormat = "General"
                    tgt.Value = CDbl(svw)
                Else
                    tgt.NumberFormat = "@"
                    tgt.Value = svw
                End If
            End If
        Next c
        outRow = outRow + 1
    Next i

End Sub

Sub 月次集計表を作る()
    ' 依頼の語: 月次集計表|月ごとの集計表|執行状況|項目ごと月|明細から月別|月別の集計表|月別集計|月別の支出|月次集計|月別・課別|課ごとの月|集計を右
    ' 扱う: 月次集計 月別集計 合計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: なし
    ' 選んでいる3列（項目列・日付列・金額列）で明細右に1列空けて月次集計表を作る

    Dim ws As Object
    Dim ur As Object
    Dim i As Long, r As Long
    Dim hRow As Long
    Dim colItem As Long, colDate As Long, colAmt As Long
    Dim lastRow As Long
    Dim startOutCol As Long
    Dim s As String
    Dim moVal As Integer
    Dim moOrder(1 To 12) As Integer
    Dim dict As Object
    Dim itemOrder() As String
    Dim itemCount As Long

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    moOrder(1) = 4:  moOrder(2) = 5:  moOrder(3) = 6:  moOrder(4) = 7
    moOrder(5) = 8:  moOrder(6) = 9:  moOrder(7) = 10: moOrder(8) = 11
    moOrder(9) = 12: moOrder(10) = 1: moOrder(11) = 2: moOrder(12) = 3

    Dim sel As Object
    Set sel = Selection

    Dim areaCount As Long
    areaCount = sel.Areas.Count

    If areaCount >= 3 Then
        colItem = sel.Areas(1).Column
        colDate = sel.Areas(2).Column
        colAmt = sel.Areas(3).Column
    ElseIf areaCount = 1 And sel.Columns.Count >= 3 Then
        colItem = sel.Column
        colDate = sel.Column + 1
        colAmt = sel.Column + 2
    Else
        Set ur = ws.UsedRange
        Dim urR1x As Long, urC1x As Long, urRowsx As Long, urColsx As Long
        urR1x = ur.Row: urC1x = ur.Column
        urRowsx = ur.rows.Count: urColsx = ur.Columns.Count

        Dim hRowX As Long: hRowX = 0
        Dim rrx As Long
        For rrx = urR1x To urR1x + urRowsx - 1
            Dim nonEmpty As Long: nonEmpty = 0
            Dim ccx As Long
            For ccx = urC1x To urC1x + urColsx - 1
                If ws.Cells(rrx, ccx).Value <> "" Then nonEmpty = nonEmpty + 1
            Next ccx
            If nonEmpty >= 3 Then
                Dim nxtx As Variant
                nxtx = ws.Cells(rrx + 1, urC1x).Value
                If nxtx <> "" Then
                    hRowX = rrx
                    Exit For
                End If
            End If
        Next rrx
        If hRowX = 0 Then Exit Sub

        colDate = 0: colAmt = 0: colItem = 0
        Dim ccx2 As Long
        For ccx2 = urC1x To urC1x + urColsx - 1
            Dim testV As Variant
            testV = ws.Cells(hRowX + 1, ccx2).Value
            If IsDate(testV) Then
                If colDate = 0 Then colDate = ccx2
            End If
        Next ccx2
        If colDate = 0 Then
            For ccx2 = urC1x To urC1x + urColsx - 1
                Dim tv2 As Variant
                tv2 = ws.Cells(hRowX + 1, ccx2).Value
                Dim sv2 As String: sv2 = Trim(CStr(tv2))
                If sv2 <> "" Then
                    If InStr(sv2, "/") > 0 Or InStr(sv2, "年") > 0 Or InStr(sv2, "-") > 0 Then
                        On Error Resume Next
                        Dim dt2 As Date: dt2 = CDate(sv2)
                        If Err.Number = 0 Then
                            If colDate = 0 Then colDate = ccx2
                        End If
                        Err.Clear
                        On Error GoTo 0
                    End If
                End If
            Next ccx2
        End If
        If colDate = 0 Then Exit Sub
        For ccx2 = urC1x To urC1x + urColsx - 1
            If ccx2 = colDate Then GoTo SkipAmt
            Dim tv3 As Variant: tv3 = ws.Cells(hRowX + 1, ccx2).Value
            Dim sv3 As String: sv3 = Trim(StrConv(CStr(tv3), vbNarrow))
            sv3 = Replace(sv3, ",", ""): sv3 = Replace(sv3, " ", "")
            If IsNumeric(sv3) And sv3 <> "" Then
                If colAmt = 0 Then colAmt = ccx2
            End If
SkipAmt:
        Next ccx2
        If colAmt = 0 Then Exit Sub
        For ccx2 = urC1x To urC1x + urColsx - 1
            If ccx2 = colDate Or ccx2 = colAmt Then GoTo SkipItem
            Dim tv4 As Variant: tv4 = ws.Cells(hRowX + 1, ccx2).Value
            If tv4 <> "" Then
                Dim sv4 As String: sv4 = Trim(CStr(tv4))
                If Not IsNumeric(sv4) And Not IsDate(sv4) Then
                    If colItem = 0 Then colItem = ccx2
                End If
            End If
SkipItem:
        Next ccx2
        If colItem = 0 Then Exit Sub
    End If

    Set ur = ws.UsedRange
    Dim urR1 As Long, urC1 As Long, urRows As Long, urCols As Long
    urR1 = ur.Row: urC1 = ur.Column
    urRows = ur.rows.Count: urCols = ur.Columns.Count

    ' 見出し行を探す
    ' まず colItem・colDate 両方が非空で非数値テキスト（見出し）の行を探す
    ' ただし colDate が数値のみの列（伝票番号など）の場合も考慮
    hRow = 0
    Dim rr As Long
    For rr = urR1 To urR1 + urRows - 1
        Dim vItem As Variant, vDate As Variant
        vItem = ws.Cells(rr, colItem).Value
        vDate = ws.Cells(rr, colDate).Value
        If vItem <> "" And vDate <> "" Then
            If Not IsNumeric(vItem) Then
                Dim nxtVal As Variant
                nxtVal = ws.Cells(rr + 1, colDate).Value
                Dim nxtOk As Boolean: nxtOk = False
                If IsDate(nxtVal) Then
                    nxtOk = True
                Else
                    Dim nxtS As String: nxtS = Trim(CStr(nxtVal))
                    If InStr(nxtS, "/") > 0 Or InStr(nxtS, "年") > 0 Or InStr(nxtS, "-") > 0 Then
                        On Error Resume Next
                        Dim nxtD As Date: nxtD = CDate(nxtS)
                        If Err.Number = 0 Then nxtOk = True
                        Err.Clear
                        On Error GoTo 0
                    End If
                    If Not nxtOk Then
                        Dim nxtItem As Variant: nxtItem = ws.Cells(rr + 1, colItem).Value
                        If Trim(CStr(nxtItem)) <> "" And Not IsNumeric(nxtItem) Then nxtOk = True
                    End If
                    ' colDate が数値列（伝票番号）の場合：次行に数値があれば見出し行とみなす
                    If Not nxtOk Then
                        If IsNumeric(nxtS) And nxtS <> "" Then nxtOk = True
                    End If
                End If
                If nxtOk Then
                    hRow = rr
                    Exit For
                End If
            End If
        End If
    Next rr
    If hRow = 0 Then
        For rr = urR1 To urR1 + urRows - 1
            Dim vd2 As Variant: vd2 = ws.Cells(rr, colDate).Value
            If IsDate(vd2) Then
                hRow = rr - 1
                If hRow < urR1 Then hRow = urR1
                Exit For
            End If
        Next rr
    End If
    If hRow = 0 Then Exit Sub

    ' 明細右端
    Dim detailRightEdge As Long
    detailRightEdge = 0
    Dim cc As Long
    For cc = urC1 To urC1 + urCols - 1
        If ws.Cells(hRow, cc).Value <> "" Then
            detailRightEdge = cc
        Else
            If detailRightEdge > 0 Then Exit For
        End If
    Next cc
    If detailRightEdge = 0 Then detailRightEdge = urC1 + urCols - 1
    If colItem > detailRightEdge Then detailRightEdge = colItem
    If colDate > detailRightEdge Then detailRightEdge = colDate
    If colAmt > detailRightEdge Then detailRightEdge = colAmt

    startOutCol = detailRightEdge + 2

    ' 明細最終行（空行・合計行除く）
    ' colDate が日付を持つ列かどうかで判断できない場合（伝票番号列選択時）に対応
    ' → colItem が非空・非合計語 かつ (colDate が日付的 or colAmt が数値的) の行を明細行とする
    lastRow = hRow
    For r = hRow + 1 To urR1 + urRows - 1
        Dim chkK As String
        chkK = Trim(CStr(ws.Cells(r, colItem).Value))
        s = chkK: GoSub 正規化
        If s = "合計" Or s = "計" Or s = "小計" Or s = "総合計" Or s = "総計" Then GoTo ChkNext
        If chkK = "" Then GoTo ChkNext
        ' 日付列または金額列に何か値があれば明細行
        Dim hasData As Boolean: hasData = False
        Dim dvChk As Variant: dvChk = ws.Cells(r, colDate).Value
        If dvChk <> "" Then hasData = True
        If Not hasData Then
            Dim avChk As Variant: avChk = ws.Cells(r, colAmt).Value
            Dim avS As String: avS = Trim(StrConv(CStr(avChk), vbNarrow))
            avS = Replace(avS, ",", ""): avS = Replace(avS, " ", "")
            If IsNumeric(avS) And avS <> "" Then hasData = True
        End If
        If hasData Then lastRow = r
ChkNext:
    Next r

    ' 既存の集計表をクリア
    Dim clearEndCol As Long
    clearEndCol = startOutCol + 14
    Dim urLastCol As Long
    urLastCol = urC1 + urCols - 1
    If urLastCol > clearEndCol Then clearEndCol = urLastCol
    Dim clearEndRow As Long
    clearEndRow = urR1 + urRows - 1
    Dim cr As Long, ccl As Long
    For cr = hRow To clearEndRow
        For ccl = startOutCol To clearEndCol
            ws.Cells(cr, ccl).Value = ""
        Next ccl
    Next cr

    ReDim itemOrder(0)
    itemCount = 0

    For r = hRow + 1 To lastRow
        Dim rawItem As String
        rawItem = Trim(CStr(ws.Cells(r, colItem).Value))
        Dim itemKey As String
        s = rawItem: GoSub 正規化
        itemKey = s
        If itemKey = "" Then GoTo NextRow
        If itemKey = "合計" Or itemKey = "計" Or itemKey = "小計" Or itemKey = "総合計" Or itemKey = "総計" Then GoTo NextRow

        ' 日付→月
        Dim dv As Variant
        dv = ws.Cells(r, colDate).Value
        moVal = 0
        If IsDate(dv) Then
            moVal = Month(CDate(dv))
        Else
            s = CStr(dv): GoSub 正規化
            If s = "" Then GoTo NextRow
            Dim mm As String
            mm = s
            If Len(mm) = 8 And IsNumeric(mm) Then
                On Error Resume Next
                Dim dtFrom8 As Date
                dtFrom8 = DateSerial(CInt(Left(mm, 4)), CInt(Mid(mm, 5, 2)), CInt(Right(mm, 2)))
                If Err.Number = 0 Then moVal = Month(dtFrom8)
                Err.Clear
                On Error GoTo 0
            End If
            If moVal = 0 And InStr(mm, "令和") > 0 And InStr(mm, "年") > 0 Then
                Dim p1r As Long, p2r As Long
                p1r = InStr(mm, "年"): p2r = InStr(mm, "月")
                If p1r > 0 And p2r > p1r Then
                    On Error Resume Next: moVal = CInt(Mid(mm, p1r + 1, p2r - p1r - 1)): On Error GoTo 0
                End If
            End If
            If moVal = 0 Then
                Dim sep As String
                sep = ""
                If InStr(mm, "/") > 0 Then sep = "/"
                If InStr(mm, "-") > 0 And sep = "" Then sep = "-"
                If sep <> "" Then
                    Dim pts() As String
                    pts = Split(mm, sep)
                    If UBound(pts) >= 1 Then
                        On Error Resume Next: moVal = CInt(pts(1)): On Error GoTo 0
                    End If
                End If
            End If
            If moVal = 0 And InStr(mm, "年") > 0 And InStr(mm, "月") > 0 Then
                Dim pn1 As Long, pn2 As Long
                pn1 = InStr(mm, "年"): pn2 = InStr(mm, "月")
                If pn2 > pn1 Then
                    On Error Resume Next: moVal = CInt(Mid(mm, pn1 + 1, pn2 - pn1 - 1)): On Error GoTo 0
                End If
            End If
            If moVal = 0 Then
                On Error Resume Next
                Dim dtTry As Date
                dtTry = CDate(mm)
                If Err.Number = 0 Then moVal = Month(dtTry)
                Err.Clear
                On Error GoTo 0
            End If
        End If
        If moVal = 0 Then GoTo NextRow

        ' 金額
        s = CStr(ws.Cells(r, colAmt).Value): GoSub 金額正規化
        Dim amtVal As Long
        amtVal = 0
        If s <> "" Then On Error Resume Next: amtVal = CLng(s): On Error GoTo 0

        If Not dict.Exists(itemKey) Then
            dict.Add itemKey, Array(CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0))
            ReDim Preserve itemOrder(itemCount)
            itemOrder(itemCount) = itemKey
            itemCount = itemCount + 1
        End If

        Dim mIdx As Integer
        mIdx = 0
        Dim mi As Integer
        For mi = 1 To 12
            If moOrder(mi) = moVal Then mIdx = mi: Exit For
        Next mi
        If mIdx = 0 Then GoTo NextRow

        Dim arr As Variant
        arr = dict(itemKey)
        arr(mIdx - 1) = arr(mIdx - 1) + amtVal
        dict(itemKey) = arr

NextRow:
    Next r

    If itemCount = 0 Then Exit Sub

    Dim dictLabel As Object
    Set dictLabel = CreateObject("Scripting.Dictionary")
    For r = hRow + 1 To lastRow
        Dim rawI2 As String
        rawI2 = Trim(CStr(ws.Cells(r, colItem).Value))
        s = rawI2: GoSub 正規化
        Dim ik2 As String: ik2 = s
        If ik2 = "" Then GoTo LabelNext
        If ik2 = "合計" Or ik2 = "計" Or ik2 = "小計" Or ik2 = "総合計" Or ik2 = "総計" Then GoTo LabelNext
        If Not dictLabel.Exists(ik2) Then
            dictLabel.Add ik2, rawI2
        End If
LabelNext:
    Next r

    Dim hItemLabel As String
    hItemLabel = CStr(ws.Cells(hRow, colItem).Value)
    If Trim(hItemLabel) = "" Then hItemLabel = "項目"

    Dim oCol As Long
    oCol = startOutCol

    ws.Cells(hRow, oCol).Value = hItemLabel
    Dim mi2 As Integer
    For mi2 = 1 To 12
        ws.Cells(hRow, oCol + mi2).Value = moOrder(mi2) & "月"
    Next mi2
    ws.Cells(hRow, oCol + 13).Value = "合計"

    Dim outRow As Long
    outRow = hRow + 1
    For i = 0 To itemCount - 1
        Dim ik As String: ik = itemOrder(i)
        Dim dispName As String
        If dictLabel.Exists(ik) Then
            dispName = dictLabel(ik)
        Else
            dispName = ik
        End If
        ws.Cells(outRow, oCol).Value = dispName
        Dim arr2 As Variant
        arr2 = dict(ik)
        For mi2 = 1 To 12
            ws.Cells(outRow, oCol + mi2).Value = arr2(mi2 - 1)
        Next mi2
        Dim c1a As String, c2a As String
        c1a = ws.Cells(outRow, oCol + 1).Address(False, False)
        c2a = ws.Cells(outRow, oCol + 12).Address(False, False)
        ws.Cells(outRow, oCol + 13).Formula = "=SUM(" & c1a & ":" & c2a & ")"
        outRow = outRow + 1
    Next i

    ws.Cells(outRow, oCol).Value = "合計"
    For mi2 = 1 To 12
        Dim r1a As String, r2a As String
        r1a = ws.Cells(hRow + 1, oCol + mi2).Address(False, False)
        r2a = ws.Cells(outRow - 1, oCol + mi2).Address(False, False)
        ws.Cells(outRow, oCol + mi2).Formula = "=SUM(" & r1a & ":" & r2a & ")"
    Next mi2
    Dim sc1 As String, sc2 As String
    sc1 = ws.Cells(outRow, oCol + 1).Address(False, False)
    sc2 = ws.Cells(outRow, oCol + 12).Address(False, False)
    ws.Cells(outRow, oCol + 13).Formula = "=SUM(" & sc1 & ":" & sc2 & ")"

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    Return

金額正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "円", "")
    Dim isNeg As Boolean
    isNeg = False
    If InStr(s, "△") > 0 Or InStr(s, "▲") > 0 Then
        isNeg = True
        s = Replace(s, "△", "")
        s = Replace(s, "▲", "")
    End If
    s = Trim$(s)
    If isNeg And s <> "" Then
        If Left(s, 1) <> "-" Then s = "-" & s
    End If
    Return

End Sub

Sub 複合グラフを作る()
    ' 依頼の語: 複合グラフ|棒と折れ線|率を第2軸|棒に率|2軸のグラフ|率の折れ線|金額は棒
    ' 扱う: 複合グラフ グラフ 合計
    ' 見出し: なし
    ' 形: 数の列
    ' 左端文字列列を項目、率でない数値列を集合縦棒(第1軸)、率列をマーカー付き折れ線(第2軸)の複合グラフを作る。合計行除外、既存グラフ削除。

    Dim ws As Object
    Set ws = ActiveSheet

    Dim co As Object
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim ur As Object
    Set ur = ws.UsedRange
    Dim rStart As Long, cStart As Long, lastR As Long, lastC As Long
    rStart = ur.Row: cStart = ur.Column
    lastR = rStart + ur.rows.Count - 1
    lastC = cStart + ur.Columns.Count - 1

    Dim titleStr As String: titleStr = ""
    Dim hdrRow As Long: hdrRow = 0
    Dim r As Long, c As Long, cc As Long, rr As Long
    For r = rStart To lastR
        Dim txtCount As Long: txtCount = 0
        For cc = cStart To lastC
            Dim cv As Variant: cv = ws.Cells(r, cc).Value
            If VarType(cv) = vbString And Trim(cv) <> "" Then txtCount = txtCount + 1
        Next cc
        If txtCount >= 2 Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Then Exit Sub

    If hdrRow > rStart Then
        Dim tr As Long
        For tr = rStart To hdrRow - 1
            Dim tcell As String: tcell = Trim(CStr(ws.Cells(tr, cStart).Value))
            If tcell <> "" Then titleStr = tcell: Exit For
        Next tr
    End If

    Dim colLabel As Long: colLabel = 0
    Dim colPct() As Long: ReDim colPct(1 To lastC - cStart + 1)
    Dim pctCount As Long: pctCount = 0
    Dim colNum() As Long: ReDim colNum(1 To lastC - cStart + 1)
    Dim numCount As Long: numCount = 0
    Dim colHdr() As String: ReDim colHdr(cStart To lastC)

    For cc = cStart To lastC
        colHdr(cc) = CStr(ws.Cells(hdrRow, cc).Value)
    Next cc

    Dim skipCols() As Boolean: ReDim skipCols(cStart To lastC)
    Dim kArr As Variant: kArr = Array("番号", "No", "no", "NO", "コード", "年度")
    For cc = cStart To lastC
        Dim hv As String: hv = Trim(colHdr(cc))
        Dim doSkip As Boolean: doSkip = False
        Dim ki As Long
        For ki = 0 To UBound(kArr)
            If InStr(hv, CStr(kArr(ki))) > 0 Then doSkip = True: Exit For
        Next ki
        skipCols(cc) = doSkip
    Next cc

    For c = cStart To lastC
        If skipCols(c) Then GoTo nextLabelCol
        Dim isTextCol As Boolean: isTextCol = True
        For rr = hdrRow + 1 To lastR
            Dim tv As Variant: tv = ws.Cells(rr, c).Value
            If tv <> "" And IsNumeric(tv) Then isTextCol = False: Exit For
        Next rr
        If isTextCol Then
            Dim hasText As Boolean: hasText = False
            For rr = hdrRow + 1 To lastR
                If Trim(CStr(ws.Cells(rr, c).Value)) <> "" Then hasText = True: Exit For
            Next rr
            If hasText Then colLabel = c: Exit For
        End If
nextLabelCol:
    Next c
    If colLabel = 0 Then colLabel = cStart

    For c = cStart To lastC
        If c = colLabel Then GoTo NextCol
        If skipCols(c) Then GoTo NextCol
        Dim fmt As String: fmt = ""
        On Error Resume Next
        fmt = ws.Cells(hdrRow + 1, c).NumberFormat
        On Error GoTo 0
        Dim isPct As Boolean: isPct = False
        If InStr(fmt, "%") > 0 Then
            isPct = True
        Else
            Dim hasVal As Boolean: hasVal = False
            Dim allFrac As Boolean: allFrac = True
            For rr = hdrRow + 1 To lastR
                Dim dv As Variant: dv = ws.Cells(rr, c).Value
                If IsNumeric(dv) And CStr(dv) <> "" Then
                    hasVal = True
                    If CDbl(dv) > 1 Or CDbl(dv) < 0 Then allFrac = False
                End If
            Next rr
            If hasVal And allFrac Then isPct = True
        End If

        If isPct Then
            pctCount = pctCount + 1
            colPct(pctCount) = c
        Else
            Dim isNum As Boolean: isNum = False
            For rr = hdrRow + 1 To lastR
                Dim nv As Variant: nv = ws.Cells(rr, c).Value
                If IsNumeric(nv) And CStr(nv) <> "" Then isNum = True: Exit For
            Next rr
            If isNum Then
                numCount = numCount + 1
                colNum(numCount) = c
            End If
        End If
NextCol:
    Next c

    If numCount = 0 And pctCount = 0 Then Exit Sub

    Dim dataRows() As Long: ReDim dataRows(1 To lastR - hdrRow)
    Dim dataCount As Long: dataCount = 0
    For r = hdrRow + 1 To lastR
        Dim lv As Variant: lv = ws.Cells(r, colLabel).Value
        If Trim(CStr(lv)) = "" Then GoTo skipRow
        Dim lstr As String: lstr = Trim(CStr(lv))
        If InStr(",合計,計,小計,総計,総合計,", "," & lstr & ",") > 0 Then GoTo skipRow
        Dim hasN As Boolean: hasN = False
        For c = cStart To lastC
            If c = colLabel Then GoTo nextChk
            If IsNumeric(ws.Cells(r, c).Value) And CStr(ws.Cells(r, c).Value) <> "" Then hasN = True: Exit For
nextChk:
        Next c
        If Not hasN Then GoTo skipRow
        dataCount = dataCount + 1
        dataRows(dataCount) = r
skipRow:
    Next r
    If dataCount = 0 Then Exit Sub

    Dim catVal() As String: ReDim catVal(1 To dataCount)
    Dim i As Long
    For i = 1 To dataCount
        catVal(i) = CStr(ws.Cells(dataRows(i), colLabel).Value)
    Next i

    If titleStr = "" Then titleStr = colHdr(colLabel) & "別の状況"

    Dim leftPos As Double: leftPos = ws.Cells(hdrRow, lastC + 2).Left
    Dim topPos As Double: topPos = ws.Cells(hdrRow, lastC + 2).Top
    Dim chObj As Object
    Set chObj = ws.ChartObjects.Add(leftPos, topPos, 420, 280)
    Dim ch As Object
    Set ch = chObj.Chart
    ch.ChartType = 51
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim n As Long
    For n = 1 To numCount
        Dim ser As Object
        Set ser = ch.SeriesCollection.NewSeries
        ser.Name = colHdr(colNum(n))
        Dim vals() As Double: ReDim vals(1 To dataCount)
        For i = 1 To dataCount
            vals(i) = CDbl(ws.Cells(dataRows(i), colNum(n)).Value)
        Next i
        ser.Values = vals
        ser.XValues = catVal
        ser.ChartType = 51
        ser.AxisGroup = 1
    Next n

    Dim p As Long
    For p = 1 To pctCount
        Dim serP As Object
        Set serP = ch.SeriesCollection.NewSeries
        serP.Name = colHdr(colPct(p))
        Dim pvals() As Double: ReDim pvals(1 To dataCount)
        For i = 1 To dataCount
            pvals(i) = CDbl(ws.Cells(dataRows(i), colPct(p)).Value)
        Next i
        serP.Values = pvals
        serP.XValues = catVal
        serP.ChartType = 65
        serP.AxisGroup = 2
        serP.Format.Line.Weight = 2.25
        serP.MarkerSize = 7
        serP.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    Next p

    On Error Resume Next
    ch.Axes(2, 1).TickLabels.NumberFormat = "#,##0"
    ch.Axes(1, 2).TickLabels.NumberFormat = "0%"
    ch.Axes(1, 2).MinimumScale = 0
    ch.Axes(1, 2).MaximumScale = 1
    ch.Axes(2, 2).TickLabels.NumberFormat = "0%"
    ch.Axes(2, 2).MinimumScale = 0
    ch.Axes(2, 2).MaximumScale = 1
    On Error GoTo 0

    ch.HasTitle = True
    ch.chartTitle.text = titleStr
    ch.chartTitle.Characters.Font.Size = 18
    ch.chartTitle.Characters.Font.Bold = True

    ch.HasLegend = True
    ch.Legend.Position = -4107

End Sub

Sub ピボット元範囲を更新する()
    ' 依頼の語: ソース範囲|明細がピボット|ピボットに反映|範囲を更新|ピボットの範囲|行まで変更|明細を足したので|足した明細
    ' 扱う: ピボット範囲更新 ピボットソース更新 合計
    ' 見出し: なし
    ' 形: 数の列 日付の列 別のシート
    ' アクティブシートを元にしたピボットのソース範囲を見出し行～最後の明細行（合計行除く）に広げてRefresh

    Dim wb As Workbook
    Set wb = ActiveSheet.Parent

    Dim meiSht As Worksheet
    Set meiSht = ActiveSheet

    Dim ur As Range
    Set ur = meiSht.UsedRange

    ' 見出し行を探す: 空でないセルが最も多い行
    Dim hdRow As Long
    hdRow = 0
    Dim r As Long, c As Long
    Dim bestCount As Long
    bestCount = 0

    Dim urR1 As Long: urR1 = ur.Row
    Dim urC1 As Long: urC1 = ur.Column
    Dim urRN As Long: urRN = ur.rows.Count
    Dim urCN As Long: urCN = ur.Columns.Count

    For r = urR1 To urR1 + urRN - 1
        Dim cnt As Long: cnt = 0
        For c = urC1 To urC1 + urCN - 1
            If Trim(CStr(meiSht.Cells(r, c).Value)) <> "" Then cnt = cnt + 1
        Next c
        If cnt > bestCount Then
            bestCount = cnt
            hdRow = r
        End If
        ' 複数行が同数の場合は最初の満行を使う（Exit For しない）
    Next r

    ' より確実に: 日付や数値が混在する行より文字ばかりの行を見出しとする
    ' 上から走査して「数値が0・文字が多い」行を見出しとする
    Dim hdRowAlt As Long: hdRowAlt = 0
    For r = urR1 To urR1 + urRN - 1
        Dim txtCnt As Long: txtCnt = 0
        Dim numCnt As Long: numCnt = 0
        For c = urC1 To urC1 + urCN - 1
            Dim cv As String: cv = Trim(CStr(meiSht.Cells(r, c).Value))
            If cv <> "" Then
                If IsNumeric(cv) Or IsDate(meiSht.Cells(r, c).Value) Then
                    numCnt = numCnt + 1
                Else
                    txtCnt = txtCnt + 1
                End If
            End If
        Next c
        If txtCnt >= 2 And numCnt = 0 Then
            hdRowAlt = r
            Exit For
        End If
    Next r
    If hdRowAlt > 0 Then hdRow = hdRowAlt

    If hdRow = 0 Then Exit Sub

    ' UsedRangeの列範囲
    Dim firstCol As Long: firstCol = urC1
    Dim lastCol As Long: lastCol = urC1 + urCN - 1

    ' 最後の明細行を探す（合計行除く）
    Dim lastDataRow As Long
    lastDataRow = hdRow
    Dim maxRow As Long: maxRow = urR1 + urRN - 1

    For r = hdRow + 1 To maxRow
        ' 行の値を結合して空行チェック
        Dim rowEmpty As Boolean: rowEmpty = True
        Dim rowStr As String: rowStr = ""
        For c = firstCol To lastCol
            Dim cellV As String: cellV = Trim(CStr(meiSht.Cells(r, c).Value))
            If cellV <> "" Then
                rowEmpty = False
                rowStr = rowStr & cellV
            End If
        Next c
        If rowEmpty Then Exit For

        ' 合計行チェック: いずれかのセルに合計・計・小計・総計を含む
        Dim isSumRow As Boolean: isSumRow = False
        For c = firstCol To lastCol
            Dim sv As String: sv = Trim(CStr(meiSht.Cells(r, c).Value))
            If InStr(sv, "合計") > 0 Or InStr(sv, "小計") > 0 Or InStr(sv, "総計") > 0 Then
                ' 数値列以外のセルに合計語があれば合計行
                If Not IsNumeric(sv) Then
                    isSumRow = True
                    Exit For
                End If
            End If
        Next c
        If Not isSumRow Then lastDataRow = r
    Next r

    If lastDataRow = hdRow Then Exit Sub

    ' 新しいソース範囲アドレス（A1形式）
    Dim newSrcAddr As String
    newSrcAddr = "'" & meiSht.Name & "'!" & _
        meiSht.Cells(hdRow, firstCol).Address(True, True) & ":" & _
        meiSht.Cells(lastDataRow, lastCol).Address(True, True)

    ' ブック内の全ピボットを走査
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            Dim srcStr As String: srcStr = ""
            On Error Resume Next
            srcStr = CStr(pt.PivotCache.SourceData)
            On Error GoTo 0
            If srcStr = "" Then GoTo nextPT

            ' ソースシート名を取得
            Dim bangPos As Long: bangPos = InStr(srcStr, "!")
            If bangPos = 0 Then GoTo nextPT
            Dim srcShtName As String
            srcShtName = Replace(Left(srcStr, bangPos - 1), "'", "")
            If srcShtName <> meiSht.Name Then GoTo nextPT

            ' キャッシュを作り直してChangePivotCache
            Dim newCache As PivotCache
            Set newCache = Nothing
            On Error Resume Next
            Set newCache = wb.PivotCaches.Create(xlDatabase, newSrcAddr)
            On Error GoTo 0
            If newCache Is Nothing Then GoTo nextPT

            On Error Resume Next
            pt.ChangePivotCache newCache
            pt.RefreshTable
            On Error GoTo 0
            Set newCache = Nothing
nextPT:
        Next pt
    Next ws

End Sub

Sub 対応表で区分を振り直して集計する()
    ' 依頼の語: 振り直|区分に振|対応表を見て区分|置き換えて区分|コードから区分|対応表で区分|区分に振り分|対応表で置き換|管理会計の区分|振り替|管理会計用|引いて集計|替えて分類|性質別に振
    ' 扱う: 振り直し 区分 集計 合計 対応なし
    ' 見出し: なし
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet, wsMap As Worksheet, s As Worksheet
    Dim wbk As Workbook
    Dim ur As Range, mapUR As Range
    Dim hdrRow As Long, dataStart As Long, lastRow As Long
    Dim codeCol As Long, amtCol As Long, divCol As Long
    Dim mapHdrRow As Long, mapCodeCol As Long, mapDivCol As Long, mapLastRow As Long
    Dim i As Long, j As Long, ri As Long, dc As Long, dr As Long
    Dim urLastCol As Long, urLastRow As Long
    Dim code As String, kCode As String, kDiv As String
    Dim divOrder() As String
    Dim divCount As Long
    Dim dict As Object
    Dim fcol As Long, gCol As Long
    Dim outRow As Long, uRow As Long, tRow As Long, clearRow As Long
    Dim ss As String
    Dim colHasData As Boolean
    Dim checkR As Long
    Dim vCode As String
    Dim amtResult As Double
    Dim as2 As String
    Dim ai As Long, ach As String, acp As Long, ares As String
    Dim rawAmt As String
    Dim divVal As String
    Dim divSum As Double
    Dim grandTotal As Double
    Dim unmatched As Double
    Dim sumDict As Object
    Dim found As Boolean
    Dim negFlag As Boolean
    Dim divHdrStr As String
    Dim amtHdrStr As String
    Dim fc As Long
    Dim rCheck As Long, cc As Long
    Dim numCnt As Long, totalCnt As Long, dashCnt As Long
    Dim rawCode As String
    Dim maxClear As Long
    Dim mapURLastCol As Long, mapURLastRow As Long
    Dim bestScore As Long, sc As Long
    Dim tMC As Long, tMD As Long, tHR As Long
    Dim tryR As Long, tryC As Long
    Dim hasDash As Long, hasText As Long
    Dim chkR2 As Long
    Dim cv As String, cv2 As String
    Dim hv As String, fhv As String
    Dim nc2 As String, ci2 As Long, ch2 As String, cp2 As Long
    Dim divHdrNorm As String, divHdrNorm2 As String
    Dim isAggRow As Boolean
    Dim r2 As Long
    Dim cellTxt As String
    Dim divRaw As String
    Dim allNum As Boolean
    Dim topRow As Long, tr As Long
    Dim tmpCode As Long, tmpAmt As Long
    Dim nonEmpty As Long
    Dim tMC2 As Long, tMD2 As Long
    Dim neededLastRow As Long
    Dim extraCol As Long

    Set ws = ActiveSheet
    Set wbk = ActiveWorkbook

    ' --- 対応表シートを形で探す ---
    bestScore = -1
    For Each s In wbk.Worksheets
        If s.Name = ws.Name Then GoTo NextSheet
        Set mapUR = s.UsedRange
        If mapUR Is Nothing Then GoTo NextSheet
        mapURLastCol = mapUR.Column + mapUR.Columns.Count - 1
        mapURLastRow = mapUR.Row + mapUR.rows.Count - 1
        If mapUR.Columns.Count < 2 Then GoTo NextSheet
        tMC = 0: tMD = 0: tHR = 0
        For tryR = mapUR.Row To Application.Min(mapUR.Row + 5, mapURLastRow)
            tMC = 0: tMD = 0
            For tryC = mapUR.Column To mapURLastCol
                hasDash = 0
                For chkR2 = tryR + 1 To Application.Min(tryR + 5, mapURLastRow)
                    cv = CStr(s.Cells(chkR2, tryC).Value)
                    ss = cv: GoSub NormalizeCode2: cv = ss
                    If InStr(cv, "-") > 0 Then hasDash = hasDash + 1
                Next chkR2
                If hasDash >= 1 Then
                    If tMC = 0 Then tMC = tryC
                ElseIf tMC > 0 And tMD = 0 And tryC > tMC Then
                    hasText = 0
                    For chkR2 = tryR + 1 To Application.Min(tryR + 5, mapURLastRow)
                        If Trim(CStr(s.Cells(chkR2, tryC).Value)) <> "" Then hasText = hasText + 1
                    Next chkR2
                    If hasText >= 1 Then tMD = tryC
                End If
            Next tryC
            If tMC = 0 Or tMD = 0 Then
                tMC2 = 0: tMD2 = 0
                For tryC = mapURLastCol To mapUR.Column Step -1
                    hasDash = 0
                    For chkR2 = tryR + 1 To Application.Min(tryR + 5, mapURLastRow)
                        cv = CStr(s.Cells(chkR2, tryC).Value)
                        ss = cv: GoSub NormalizeCode2: cv = ss
                        If InStr(cv, "-") > 0 Then hasDash = hasDash + 1
                    Next chkR2
                    If hasDash >= 1 Then
                        If tMC2 = 0 Then tMC2 = tryC
                    End If
                Next tryC
                If tMC2 > 0 Then
                    For tryC = mapUR.Column To mapURLastCol
                        If tryC = tMC2 Then GoTo SkipThisCol
                        hasText = 0
                        For chkR2 = tryR + 1 To Application.Min(tryR + 5, mapURLastRow)
                            If Trim(CStr(s.Cells(chkR2, tryC).Value)) <> "" Then hasText = hasText + 1
                        Next chkR2
                        If hasText >= 1 Then
                            allNum = True
                            For chkR2 = tryR + 1 To Application.Min(tryR + 5, mapURLastRow)
                                cv = Trim(CStr(s.Cells(chkR2, tryC).Value))
                                If cv <> "" And Not IsNumeric(cv) Then allNum = False
                            Next chkR2
                            If Not allNum Then
                                If tMD2 = 0 Then tMD2 = tryC
                            End If
                        End If
SkipThisCol:
                    Next tryC
                    If tMC2 > 0 And tMD2 > 0 Then
                        tMC = tMC2: tMD = tMD2
                    End If
                End If
            End If
            If tMC > 0 And tMD > 0 Then
                tHR = tryR
                Exit For
            End If
        Next tryR
        If tMC > 0 And tMD > 0 Then
            sc = mapURLastRow - mapUR.Row
            If sc > bestScore Then
                bestScore = sc
                Set wsMap = s
                mapCodeCol = tMC
                mapDivCol = tMD
                mapHdrRow = tHR
                mapURLastRow = mapUR.Row + mapUR.rows.Count - 1
            End If
        End If
NextSheet:
    Next s
    If wsMap Is Nothing Then Exit Sub

    ' --- 対応表を読む ---
    Set dict = CreateObject("Scripting.Dictionary")
    ReDim divOrder(0)
    divCount = 0
    mapLastRow = wsMap.UsedRange.Row + wsMap.UsedRange.rows.Count - 1

    For i = mapHdrRow + 1 To mapLastRow
        rawCode = CStr(wsMap.Cells(i, mapCodeCol).Value)
        ss = rawCode: GoSub NormalizeCode2: kCode = ss
        divRaw = CStr(wsMap.Cells(i, mapDivCol).Value)
        ss = divRaw: GoSub NormalizePlain: kDiv = ss
        If kCode <> "" And kDiv <> "" Then
            If Not dict.Exists(kCode) Then dict(kCode) = kDiv
        End If
        If kDiv <> "" Then
            found = False
            For j = 0 To divCount - 1
                If divOrder(j) = kDiv Then found = True: Exit For
            Next j
            If Not found Then
                ReDim Preserve divOrder(divCount)
                divOrder(divCount) = kDiv
                divCount = divCount + 1
            End If
        End If
    Next i

    ' --- 明細シートの見出し行を探す ---
    Set ur = ws.UsedRange
    urLastCol = ur.Column + ur.Columns.Count - 1
    urLastRow = ur.Row + ur.rows.Count - 1
    hdrRow = 0
    codeCol = 0: amtCol = 0

    For ri = ur.Row To Application.Min(ur.Row + 10, urLastRow)
        nonEmpty = 0
        For cc = ur.Column To urLastCol
            If Trim(CStr(ws.Cells(ri, cc).Value)) <> "" Then nonEmpty = nonEmpty + 1
        Next cc
        If nonEmpty < 2 Then GoTo NextHdrRow

        tmpCode = 0: tmpAmt = 0
        For cc = ur.Column To urLastCol
            If tmpCode = 0 Then
                dashCnt = 0
                For rCheck = ri + 1 To Application.Min(ri + 8, urLastRow)
                    ss = CStr(ws.Cells(rCheck, cc).Value): GoSub NormalizeCode2
                    If InStr(ss, "-") > 0 Then dashCnt = dashCnt + 1
                Next rCheck
                If dashCnt >= 2 Then tmpCode = cc
            End If
        Next cc
        If tmpCode > 0 Then
            For cc = tmpCode + 1 To urLastCol
                numCnt = 0: totalCnt = 0
                For rCheck = ri + 1 To Application.Min(ri + 8, urLastRow)
                    cv2 = CStr(ws.Cells(rCheck, cc).Value)
                    If cv2 <> "" Then
                        totalCnt = totalCnt + 1
                        If IsNumeric(cv2) Then numCnt = numCnt + 1
                    End If
                Next rCheck
                If totalCnt > 0 And numCnt >= totalCnt * 0.5 Then
                    tmpAmt = cc
                    Exit For
                End If
            Next cc
        End If
        If tmpCode > 0 And tmpAmt > 0 Then
            hdrRow = ri
            codeCol = tmpCode
            amtCol = tmpAmt
            Exit For
        End If
NextHdrRow:
    Next ri
    If hdrRow = 0 Then Exit Sub

    divHdrStr = CStr(wsMap.Cells(mapHdrRow, mapDivCol).Value)
    If Trim(divHdrStr) = "" Then divHdrStr = "区分"
    amtHdrStr = CStr(ws.Cells(hdrRow, amtCol).Value)
    If Trim(amtHdrStr) = "" Then amtHdrStr = "金額"

    dataStart = hdrRow + 1
    topRow = ur.Row

    ' lastRow: コード列に "-" を含む最後の行（合計行除く）
    lastRow = 0
    For checkR = urLastRow To dataStart Step -1
        isAggRow = False
        For r2 = ur.Column To amtCol
            cellTxt = Trim(CStr(ws.Cells(checkR, r2).Value))
            ss = cellTxt: GoSub NormalizePlain
            If ss = "合計" Or ss = "計" Or ss = "小計" Or ss = "総計" Then
                isAggRow = True: Exit For
            End If
        Next r2
        If isAggRow Then GoTo NextCheckR
        ss = CStr(ws.Cells(checkR, codeCol).Value): GoSub NormalizeCode2: vCode = ss
        If InStr(vCode, "-") > 0 Then
            lastRow = checkR
            Exit For
        End If
NextCheckR:
    Next checkR
    If lastRow = 0 Then lastRow = dataStart

    ' --- 区分列を決める ---
    ' amtCol+1 から既存の同じ見出し列を探す
    ss = divHdrStr: GoSub NormalizePlain: divHdrNorm = ss
    divCol = 0

    For dc = amtCol + 1 To urLastCol
        hv = CStr(ws.Cells(hdrRow, dc).Value)
        ss = hv: GoSub NormalizePlain
        If ss = divHdrNorm And ss <> "" Then
            divCol = dc
            Exit For
        End If
    Next dc

    If divCol = 0 Then
        ' amtCol+1 から空の列を探す
        For dc = amtCol + 1 To amtCol + 20
            colHasData = False
            For dr = hdrRow To lastRow
                If Trim(CStr(ws.Cells(dr, dc).Value)) <> "" Then
                    colHasData = True: Exit For
                End If
            Next dr
            If Not colHasData Then
                divCol = dc: Exit For
            End If
        Next dc
    End If
    If divCol = 0 Then divCol = amtCol + 1

    ws.Cells(hdrRow, divCol).Value = divHdrStr

    ' 見出し行より上はクリア
    If hdrRow > topRow Then
        For tr = topRow To hdrRow - 1
            ws.Cells(tr, divCol).Value = ""
        Next tr
    End If

    ' --- 各行に区分を書く ---
    For i = dataStart To lastRow
        rawCode = CStr(ws.Cells(i, codeCol).Value)
        ss = rawCode: GoSub NormalizeCode2: code = ss

        isAggRow = False
        For r2 = ur.Column To amtCol
            cellTxt = Trim(CStr(ws.Cells(i, r2).Value))
            ss = cellTxt: GoSub NormalizePlain
            If ss = "合計" Or ss = "計" Or ss = "小計" Or ss = "総計" Then
                isAggRow = True: Exit For
            End If
        Next r2

        If code = "" Or isAggRow Then
            ws.Cells(i, divCol).Value = ""
        ElseIf InStr(code, "-") > 0 Then
            If dict.Exists(code) Then
                ws.Cells(i, divCol).Value = dict(code)
            Else
                ws.Cells(i, divCol).Value = "対応なし"
            End If
        Else
            ws.Cells(i, divCol).Value = ""
        End If
    Next i

    ' lastRow より下の divCol はクリア
    For clearRow = lastRow + 1 To urLastRow
        If Trim(CStr(ws.Cells(clearRow, divCol).Value)) <> "" Then
            ws.Cells(clearRow, divCol).Value = ""
        End If
    Next clearRow

    ' --- 集計表の列を決める: divCol+2 から ---
    ' divCol+2 の位置に既存の同じ見出し列があればそこを使う（それ以外は作らない）
    ss = divHdrStr: GoSub NormalizePlain: divHdrNorm2 = ss
    fcol = divCol + 2
    gCol = divCol + 3

    ' divCol+2 の見出しを確認（既に書かれていれば再利用、違う見出しがあれば上書き）
    ' divCol+1 は空列として扱う（触らない）
    ' fCol より右（gCol+1以降）の余計なデータをクリア
    For extraCol = gCol + 1 To Application.Max(urLastCol, gCol + 10)
        Dim anyVal As Boolean
        anyVal = False
        For clearRow = topRow To urLastRow
            If Trim(CStr(ws.Cells(clearRow, extraCol).Value)) <> "" Then
                anyVal = True: Exit For
            End If
        Next clearRow
        If anyVal Then
            For clearRow = topRow To urLastRow
                ws.Cells(clearRow, extraCol).Value = ""
            Next clearRow
        End If
    Next extraCol

    ' 集計表書き込み範囲をクリア
    neededLastRow = hdrRow + divCount + 2
    maxClear = urLastRow
    If neededLastRow > maxClear Then maxClear = neededLastRow
    For clearRow = topRow To maxClear
        ws.Cells(clearRow, fcol).Value = ""
        ws.Cells(clearRow, gCol).Value = ""
    Next clearRow

    ws.Cells(hdrRow, fcol).Value = divHdrStr
    ws.Cells(hdrRow, gCol).Value = amtHdrStr

    ' --- 集計 ---
    Set sumDict = CreateObject("Scripting.Dictionary")
    unmatched = 0

    For i = dataStart To lastRow
        divVal = Trim(CStr(ws.Cells(i, divCol).Value))
        If divVal = "" Then GoTo NextDataRow
        rawAmt = CStr(ws.Cells(i, amtCol).Value): GoSub NormalizeAmount
        If divVal = "対応なし" Then
            unmatched = unmatched + amtResult
        Else
            If sumDict.Exists(divVal) Then
                sumDict(divVal) = sumDict(divVal) + amtResult
            Else
                sumDict(divVal) = amtResult
            End If
        End If
NextDataRow:
    Next i

    grandTotal = 0

    For j = 0 To divCount - 1
        outRow = hdrRow + 1 + j
        ws.Cells(outRow, fcol).Value = divOrder(j)
        If sumDict.Exists(divOrder(j)) Then
            divSum = sumDict(divOrder(j))
        Else
            divSum = 0
        End If
        ws.Cells(outRow, gCol).Value = divSum
        grandTotal = grandTotal + divSum
    Next j

    uRow = hdrRow + 1 + divCount
    ws.Cells(uRow, fcol).Value = "対応なし"
    ws.Cells(uRow, gCol).Value = unmatched
    grandTotal = grandTotal + unmatched

    tRow = uRow + 1
    ws.Cells(tRow, fcol).Value = "合計"
    ws.Cells(tRow, gCol).Value = grandTotal

    Exit Sub

NormalizePlain:
    ss = Replace(ss, ChrW(160), "")
    ss = Trim$(ss)
    ss = Replace(ss, "　", "")
    ss = Replace(ss, " ", "")
    Return

NormalizeCode2:
    ss = Replace(ss, ChrW(160), "")
    ss = Trim$(ss)
    ss = Replace(ss, "　", "")
    ss = Replace(ss, " ", "")
    nc2 = ""
    For ci2 = 1 To Len(ss)
        ch2 = Mid$(ss, ci2, 1)
        cp2 = AscW(ch2)
        If cp2 >= &HFF21 And cp2 <= &HFF3A Then
            nc2 = nc2 & Chr(cp2 - &HFF21 + 65)
        ElseIf cp2 >= &HFF41 And cp2 <= &HFF5A Then
            nc2 = nc2 & Chr(cp2 - &HFF41 + 97)
        ElseIf cp2 >= &HFF10 And cp2 <= &HFF19 Then
            nc2 = nc2 & Chr(cp2 - &HFF10 + 48)
        ElseIf cp2 = &HFF0D Then
            nc2 = nc2 & "-"
        ElseIf cp2 = &HFF0E Then
            nc2 = nc2 & "."
        Else
            nc2 = nc2 & ch2
        End If
    Next ci2
    ss = nc2
    ss = Replace(ss, ",", "")
    Return

NormalizeAmount:
    ares = ""
    negFlag = False
    For ai = 1 To Len(rawAmt)
        ach = Mid$(rawAmt, ai, 1)
        acp = AscW(ach)
        If acp >= &HFF10 And acp <= &HFF19 Then
            ares = ares & Chr(acp - &HFF10 + 48)
        ElseIf acp >= 48 And acp <= 57 Then
            ares = ares & ach
        ElseIf acp = &HFF0D Or acp = &H2212 Or ach = "△" Or ach = "▲" Then
            negFlag = True
        ElseIf ach = "-" And Len(ares) = 0 Then
            negFlag = True
        End If
    Next ai
    as2 = Trim$(ares)
    If as2 = "" Then
        amtResult = 0
    Else
        On Error Resume Next
        amtResult = CDbl(as2)
        If Err.Number <> 0 Then amtResult = 0
        On Error GoTo 0
        If negFlag Then amtResult = -amtResult
    End If
    Return

End Sub

Sub 別シートと突合する()
    ' 依頼の語: 突合|照合|差異を出|新旧のデータ|移行前と移行後|相手の表と|相手に無い伝票|件数も出|新システム|旧と新|伝票が合|旧システム|データ移行|移行後
    ' 扱う: 突合 照合 合計・突き合
    ' 見出し: なし
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim wbk As Workbook
    Set wbk = ActiveSheet.Parent

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim uR1 As Long, uC1 As Long, ur2 As Long, uC2 As Long
    uR1 = ur.Row: uC1 = ur.Column
    ur2 = uR1 + ur.rows.Count - 1: uC2 = uC1 + ur.Columns.Count - 1

    Dim hRow As Long: hRow = 0
    Dim tryR As Long, ci As Long, ri As Long, s As String

    ' 見出し行: 数値が2つ以上並ぶ行の直前行
    Dim dataStartR As Long: dataStartR = 0
    For tryR = uR1 To Application.Min(uR1 + 15, ur2)
        Dim numCnt2 As Long: numCnt2 = 0
        Dim ci2 As Long
        For ci2 = uC1 To uC2
            Dim cv As Variant: cv = ws.Cells(tryR, ci2).Value
            If IsNumeric(cv) And CStr(cv) <> "" Then numCnt2 = numCnt2 + 1
        Next ci2
        If numCnt2 >= 2 Then
            dataStartR = tryR
            If tryR > uR1 Then hRow = tryR - 1 Else hRow = tryR
            Exit For
        End If
    Next tryR
    If hRow = 0 Then
        For tryR = uR1 To Application.Min(uR1 + 5, ur2)
            Dim sc2 As Long: sc2 = 0
            For ci2 = uC1 To uC2
                If CStr(ws.Cells(tryR, ci2).Value) <> "" Then sc2 = sc2 + 1
            Next ci2
            If sc2 >= 2 Then hRow = tryR: dataStartR = tryR + 1: Exit For
        Next tryR
    End If
    If hRow = 0 Then Exit Sub

    Dim dR1 As Long: dR1 = hRow + 1
    Dim dR2 As Long: dR2 = ur2

    ' 既存の「相手の金額」「突合」列を探す（除外用）
    Dim exPartnerAmt As Long, exMatchCol As Long
    exPartnerAmt = 0: exMatchCol = 0
    For ci = uC1 To uC2
        s = CStr(ws.Cells(hRow, ci).Value): GoSub NormHdr
        If s = "相手の金額" And exPartnerAmt = 0 Then exPartnerAmt = ci
        If s = "突合" And exMatchCol = 0 And ci > uC1 Then exMatchCol = ci
    Next ci

    Dim totChk As Long
    Dim rowIsTotal As Boolean
    ' キー列: 非空ユニーク値が最多の列（除外列を除く）
    Dim colKey As Long, colBestKeyScore As Long
    colKey = 0: colBestKeyScore = -1
    For ci = uC1 To uC2
        If ci = exPartnerAmt Or ci = exMatchCol Then GoTo SkipKeyCol2
        Dim uniqD As Object
        Set uniqD = CreateObject("Scripting.Dictionary")
        For ri = dR1 To dR2
            rowIsTotal = False
            For totChk = uC1 To uC2
                s = Trim(CStr(ws.Cells(ri, totChk).Value))
                If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then rowIsTotal = True
            Next totChk
            If rowIsTotal Then GoTo SkipKeyRow
            s = CStr(ws.Cells(ri, ci).Value): GoSub NormKeySub
            If s <> "" Then uniqD(s) = 1
SkipKeyRow:
        Next ri
        Dim uCnt As Long: uCnt = uniqD.Count
        If uCnt > colBestKeyScore Then colBestKeyScore = uCnt: colKey = ci
SkipKeyCol2:
    Next ci

    ' 金額列: キー列・除外列を除いた列で数値が最多の列
    Dim colAmt As Long, maxNum As Long
    colAmt = 0: maxNum = -1
    Dim colCount As Long: colCount = uC2 - uC1 + 1
    ReDim cntNum(1 To colCount) As Long
    For ri = dR1 To dR2
        For ci = uC1 To uC2
            If ci = exPartnerAmt Or ci = exMatchCol Or ci = colKey Then GoTo SkipCountCol
            Dim ix As Long: ix = ci - uC1 + 1
            s = CStr(ws.Cells(ri, ci).Value): GoSub NormAmtSub
            If IsNumeric(s) And s <> "" Then cntNum(ix) = cntNum(ix) + 1
SkipCountCol:
        Next ci
    Next ri
    For ci = uC1 To uC2
        If ci = exPartnerAmt Or ci = exMatchCol Or ci = colKey Then GoTo SkipAmtCol
        ix = ci - uC1 + 1
        If cntNum(ix) > maxNum Then maxNum = cntNum(ix): colAmt = ci
SkipAmtCol:
    Next ci

    If colKey = 0 Or colAmt = 0 Then Exit Sub

    ' アクティブシートのキー値セット（合計行・空行を除く）
    Dim dicMyKeys As Object
    Set dicMyKeys = CreateObject("Scripting.Dictionary")
    For ri = dR1 To dR2
        s = CStr(ws.Cells(ri, colKey).Value): GoSub NormKeySub
        Dim nkTmp As String: nkTmp = s
        If nkTmp = "" Then GoTo SkipMyKey
        Dim kDispTmp As String: kDispTmp = Trim(CStr(ws.Cells(ri, colKey).Value))
        If kDispTmp = "合計" Or kDispTmp = "計" Or kDispTmp = "小計" Or kDispTmp = "総計" Then GoTo SkipMyKey
        dicMyKeys(nkTmp) = 1
SkipMyKey:
    Next ri

    ' 相手シートを探す（キー値が最も重なるシート）
    Dim wsP As Worksheet
    Dim bestScore As Long: bestScore = -1
    Dim sh As Worksheet
    For Each sh In wbk.Sheets
        If sh.Name = ws.Name Then GoTo NextSheet
        Dim shUr As Range
        On Error Resume Next: Set shUr = sh.UsedRange: On Error GoTo 0
        If shUr Is Nothing Then GoTo NextSheet

        Dim shR1 As Long, shC1 As Long, shR2 As Long, shC2 As Long
        shR1 = shUr.Row: shC1 = shUr.Column
        shR2 = shR1 + shUr.rows.Count - 1: shC2 = shC1 + shUr.Columns.Count - 1

        ' 相手シートの見出し行: 本文(数値2つ以上)の直前行
        Dim shHRow As Long: shHRow = 0
        Dim shDataR As Long: shDataR = 0
        For tryR = shR1 To Application.Min(shR1 + 15, shR2)
            Dim shnc As Long: shnc = 0
            For ci = shC1 To shC2
                Dim shv As Variant: shv = sh.Cells(tryR, ci).Value
                If IsNumeric(shv) And CStr(shv) <> "" Then shnc = shnc + 1
            Next ci
            If shnc >= 2 Then
                shDataR = tryR
                If tryR > shR1 Then shHRow = tryR - 1 Else shHRow = tryR
                Exit For
            End If
        Next tryR
        If shHRow = 0 Then
            For tryR = shR1 To Application.Min(shR1 + 8, shR2)
                Dim shsc As Long: shsc = 0
                For ci = shC1 To shC2
                    If CStr(sh.Cells(tryR, ci).Value) <> "" Then shsc = shsc + 1
                Next ci
                If shsc >= 2 Then shHRow = tryR: shDataR = tryR + 1: Exit For
            Next tryR
        End If
        If shHRow = 0 Then GoTo NextSheet
        Dim shDR1x As Long: shDR1x = shHRow + 1

        ' 相手シートの各列とdicMyKeysとの一致数の最大を求める
        Dim bestShScore As Long: bestShScore = 0
        Dim shci As Long
        For shci = shC1 To shC2
            Dim shHdrChk As String: shHdrChk = CStr(sh.Cells(shHRow, shci).Value)
            s = shHdrChk: GoSub NormHdr
            If s = "相手の金額" Or s = "突合" Or s = "件数" Then GoTo NextShCI
            Dim hit As Long: hit = 0
            For ri = shDR1x To shR2
                s = CStr(sh.Cells(ri, shci).Value): GoSub NormKeySub
                If s <> "" And dicMyKeys.Exists(s) Then hit = hit + 1
            Next ri
            If hit > bestShScore Then bestShScore = hit
NextShCI:
        Next shci

        If bestShScore > bestScore Then
            bestScore = bestShScore
            Set wsP = sh
        End If
NextSheet:
    Next sh
    If wsP Is Nothing Then Exit Sub

    ' 相手シートの見出し行・データ範囲
    Dim pur As Range
    Set pur = wsP.UsedRange
    Dim pR1 As Long, pC1 As Long, pR2 As Long, pC2 As Long
    pR1 = pur.Row: pC1 = pur.Column
    pR2 = pR1 + pur.rows.Count - 1: pC2 = pC1 + pur.Columns.Count - 1

    ' 相手シートの見出し行: 本文(数値2つ以上)の直前行
    Dim pHRow As Long: pHRow = 0
    For tryR = pR1 To Application.Min(pR1 + 15, pR2)
        Dim pnc2 As Long: pnc2 = 0
        For ci = pC1 To pC2
            Dim pv2 As Variant: pv2 = wsP.Cells(tryR, ci).Value
            If IsNumeric(pv2) And CStr(pv2) <> "" Then pnc2 = pnc2 + 1
        Next ci
        If pnc2 >= 2 Then
            If tryR > pR1 Then pHRow = tryR - 1 Else pHRow = tryR
            Exit For
        End If
    Next tryR
    If pHRow = 0 Then
        For tryR = pR1 To Application.Min(pR1 + 8, pR2)
            Dim psc As Long: psc = 0
            For ci = pC1 To pC2
                If CStr(wsP.Cells(tryR, ci).Value) <> "" Then psc = psc + 1
            Next ci
            If psc >= 2 Then pHRow = tryR: Exit For
        Next tryR
    End If
    If pHRow = 0 Then Exit Sub
    Dim pDR1 As Long: pDR1 = pHRow + 1

    ' 相手シートのキー列: dicMyKeysとの一致が最多の列（除外列を除く）
    Dim pColKey As Long, pBestHit As Long
    pColKey = 0: pBestHit = 0
    Dim shci2 As Long
    For shci2 = pC1 To pC2
        Dim phdrS As String: phdrS = CStr(wsP.Cells(pHRow, shci2).Value)
        s = phdrS: GoSub NormHdr
        If s = "相手の金額" Or s = "突合" Or s = "件数" Then GoTo SkipPKey
        Dim phit As Long: phit = 0
        For ri = pDR1 To pR2
            s = CStr(wsP.Cells(ri, shci2).Value): GoSub NormKeySub
            If s <> "" And dicMyKeys.Exists(s) Then phit = phit + 1
        Next ri
        If phit > pBestHit Then pBestHit = phit: pColKey = shci2
SkipPKey:
    Next shci2
    If pColKey = 0 Then Exit Sub

    ' 相手シートの金額列: キー列・除外列を除いた数値最多の列
    Dim pColAmt As Long, pMaxNum As Long
    pColAmt = 0: pMaxNum = -1
    For shci2 = pC1 To pC2
        If shci2 = pColKey Then GoTo SkipPAmt
        Dim pHdrS2 As String: pHdrS2 = CStr(wsP.Cells(pHRow, shci2).Value)
        s = pHdrS2: GoSub NormHdr
        If s = "相手の金額" Or s = "突合" Or s = "件数" Then GoTo SkipPAmt
        Dim pnnc As Long: pnnc = 0
        Dim pstrc As Long: pstrc = 0
        For ri = pDR1 To pR2
            s = CStr(wsP.Cells(ri, pColKey).Value): GoSub NormKeySub
            If s = "" Then GoTo SkipPAmtRow
            Dim pCellVal As String: pCellVal = CStr(wsP.Cells(ri, shci2).Value)
            If pCellVal = "" Then GoTo SkipPAmtRow
            s = pCellVal: GoSub NormAmtSub
            If IsNumeric(s) And s <> "" Then
                pnnc = pnnc + 1
            Else
                pstrc = pstrc + 1
            End If
SkipPAmtRow:
        Next ri
        If pstrc > pnnc Then GoTo SkipPAmt
        If pnnc > pMaxNum Then pMaxNum = pnnc: pColAmt = shci2
SkipPAmt:
    Next shci2
    If pColAmt = 0 Then Exit Sub

    ' 相手データをDictionaryに読む
    Dim dicP As Object
    Set dicP = CreateObject("Scripting.Dictionary")
    For ri = pDR1 To pR2
        s = CStr(wsP.Cells(ri, pColKey).Value): GoSub NormKeySub
        Dim pkey As String: pkey = s
        If pkey = "" Then GoTo NextPRow
        ' 合計行除外
        Dim pkDisp As String: pkDisp = Trim(CStr(wsP.Cells(ri, pColKey).Value))
        If pkDisp = "合計" Or pkDisp = "計" Or pkDisp = "小計" Or pkDisp = "総計" Then GoTo NextPRow
        Dim pamtRaw As String: pamtRaw = CStr(wsP.Cells(ri, pColAmt).Value)
        s = pamtRaw: GoSub NormAmtSub
        If IsNumeric(s) And s <> "" Then
            If Not dicP.Exists(pkey) Then dicP(pkey) = CDbl(s)
        End If
NextPRow:
    Next ri

    ' 出力列を確定
    Dim scanEnd As Long
    scanEnd = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
    If scanEnd < uC2 Then scanEnd = uC2

    Dim colPartnerAmt As Long, colMatch As Long
    colPartnerAmt = 0: colMatch = 0
    For ci = uC1 To scanEnd
        s = CStr(ws.Cells(hRow, ci).Value): GoSub NormHdr
        If s = "相手の金額" And colPartnerAmt = 0 Then colPartnerAmt = ci
        If s = "突合" And ci > colAmt And colMatch = 0 Then colMatch = ci
    Next ci

    If colPartnerAmt = 0 Then
        colPartnerAmt = scanEnd + 1
        ws.Cells(hRow, colPartnerAmt).Value = "相手の金額"
        scanEnd = colPartnerAmt
    End If
    If colMatch = 0 Then
        If colPartnerAmt > scanEnd Then scanEnd = colPartnerAmt
        colMatch = scanEnd + 1
        ws.Cells(hRow, colMatch).Value = "突合"
        scanEnd = colMatch
    End If

    ' 突合処理
    Dim cntMatch As Long, cntDiff As Long, cntNoP As Long
    cntMatch = 0: cntDiff = 0: cntNoP = 0

    Dim dicMyUsed As Object
    Set dicMyUsed = CreateObject("Scripting.Dictionary")

    For ri = dR1 To dR2
        Dim keyRaw As String: keyRaw = CStr(ws.Cells(ri, colKey).Value)
        s = keyRaw: GoSub NormKeySub
        Dim myKey As String: myKey = s
        If myKey = "" Then GoTo SkipDataRow
        Dim keyDisp2 As String: keyDisp2 = Trim(CStr(ws.Cells(ri, colKey).Value))
        If keyDisp2 = "合計" Or keyDisp2 = "計" Or keyDisp2 = "小計" Or keyDisp2 = "総計" Then GoTo SkipDataRow

        dicMyUsed(myKey) = 1

        Dim myAmtRaw As String: myAmtRaw = CStr(ws.Cells(ri, colAmt).Value)
        s = myAmtRaw: GoSub NormAmtSub
        Dim myAmt As String: myAmt = s

        If dicP.Exists(myKey) Then
            Dim pAmtDbl As Double: pAmtDbl = CDbl(dicP(myKey))
            ws.Cells(ri, colPartnerAmt).Value = pAmtDbl
            If IsNumeric(myAmt) And myAmt <> "" Then
                If CDbl(myAmt) = pAmtDbl Then
                    ws.Cells(ri, colMatch).Value = "一致"
                    cntMatch = cntMatch + 1
                Else
                    ws.Cells(ri, colMatch).Value = "金額違い"
                    cntDiff = cntDiff + 1
                End If
            Else
                ws.Cells(ri, colMatch).Value = "金額違い"
                cntDiff = cntDiff + 1
            End If
        Else
            ws.Cells(ri, colPartnerAmt).Value = ""
            ws.Cells(ri, colMatch).Value = "相手に無し"
            cntNoP = cntNoP + 1
        End If
SkipDataRow:
    Next ri

    Dim cntNoMe As Long: cntNoMe = 0
    Dim kk As Variant
    For Each kk In dicP.keys
        If Not dicMyUsed.Exists(kk) Then cntNoMe = cntNoMe + 1
    Next kk

    ' 件数表列を確定（突合列の2列右）
    Dim colSumLbl As Long, colSumCnt As Long
    colSumLbl = 0: colSumCnt = 0

    Dim scanEnd2 As Long
    scanEnd2 = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
    If scanEnd2 < colMatch Then scanEnd2 = colMatch

    For ci = colMatch + 1 To scanEnd2 + 2
        s = CStr(ws.Cells(hRow, ci).Value): GoSub NormHdr
        If s = "突合" Then
            colSumLbl = ci
            Dim nci As Long: nci = ci + 1
            s = CStr(ws.Cells(hRow, nci).Value): GoSub NormHdr
            colSumCnt = nci
            If s <> "件数" Then ws.Cells(hRow, colSumCnt).Value = "件数"
            Exit For
        End If
    Next ci

    If colSumLbl = 0 Then
        colSumLbl = colMatch + 2
        colSumCnt = colSumLbl + 1
        ws.Cells(hRow, colSumLbl).Value = "突合"
        ws.Cells(hRow, colSumCnt).Value = "件数"
    End If

    ws.Cells(hRow + 1, colSumLbl).Value = "一致"
    ws.Cells(hRow + 1, colSumCnt).Value = cntMatch
    ws.Cells(hRow + 2, colSumLbl).Value = "金額違い"
    ws.Cells(hRow + 2, colSumCnt).Value = cntDiff
    ws.Cells(hRow + 3, colSumLbl).Value = "相手に無し"
    ws.Cells(hRow + 3, colSumCnt).Value = cntNoP
    ws.Cells(hRow + 4, colSumLbl).Value = "こちらに無し"
    ws.Cells(hRow + 4, colSumCnt).Value = cntNoMe

    Dim clr As Long
    For clr = hRow + 5 To hRow + 10
        If ws.Cells(clr, colSumLbl).Value <> "" Or ws.Cells(clr, colSumCnt).Value <> "" Then
            ws.Cells(clr, colSumLbl).Value = ""
            ws.Cells(clr, colSumCnt).Value = ""
        End If
    Next clr

    Exit Sub

NormHdr:
    Dim nh As String: nh = s
    nh = Replace(nh, Chr(10), "")
    nh = Replace(nh, Chr(13), "")
    nh = Replace(nh, ChrW(12288), "")
    nh = Replace(nh, ChrW(160), "")
    nh = Replace(nh, " ", "")
    nh = Trim(nh)
    Dim nhI As Long, nhC As String, nhCW As Long, nhOut As String
    nhOut = ""
    For nhI = 1 To Len(nh)
        nhC = Mid(nh, nhI, 1)
        nhCW = AscW(nhC)
        If nhCW >= &HFF01 And nhCW <= &HFF5E Then nhC = Chr(nhCW - &HFF01 + &H21)
        nhOut = nhOut & nhC
    Next nhI
    s = Trim(nhOut)
    Return

NormKeySub:
    Dim nk As String: nk = s
    nk = Replace(nk, ChrW(160), "")
    nk = Replace(nk, ChrW(12288), "")
    nk = Trim(nk)
    Dim nkI As Long, nkC As String, nkCW As Long, nkOut As String
    nkOut = ""
    For nkI = 1 To Len(nk)
        nkC = Mid(nk, nkI, 1)
        nkCW = AscW(nkC)
        If nkCW >= &HFF10 And nkCW <= &HFF19 Then nkC = Chr(Asc("0") + nkCW - &HFF10)
        nkOut = nkOut & nkC
    Next nkI
    nkOut = Trim(nkOut)
    If IsNumeric(nkOut) Then
        s = CStr(CLng(nkOut))
    Else
        s = nkOut
    End If
    Return

NormAmtSub:
    Dim na As String: na = s
    na = Replace(na, ChrW(160), "")
    na = Replace(na, ChrW(12288), "")
    na = Trim(na)
    Dim naI As Long, naC As String, naCW As Long, naOut As String
    naOut = ""
    For naI = 1 To Len(na)
        naC = Mid(na, naI, 1)
        naCW = AscW(naC)
        If naCW >= &HFF10 And naCW <= &HFF19 Then naC = Chr(Asc("0") + naCW - &HFF10)
        If naC = "," Or naCW = &HFF0C Then GoTo SkipNaC
        If naC = " " Then GoTo SkipNaC
        naOut = naOut & naC
SkipNaC:
    Next naI
    naOut = Trim(naOut)
    If IsNumeric(naOut) Then
        s = CStr(CLng(naOut))
    Else
        s = naOut
    End If
    Return

End Sub

Sub 複数シートを集約する()
    ' 依頼の語: 集約|シートの回答を集|全シートを1つ|各シートから集|各シート|シートを集|回答を取|事業を集|各課から来
    ' 扱う: 集約 まとめ 合計
    ' 見出し: なし
    ' 形: 別のシート 共通の見出しのシート
    ' 集める先シートの見出し行に従い他シートの表を転記し最後に合計行を追加する

    Dim destWs As Worksheet
    Dim ws As Worksheet
    Dim ur As Range, srcUr As Range
    Dim r As Long, c As Long, ri As Long, ci As Long
    Dim s As String, key As String
    Dim hdrRow As Long, srcHdrRow As Long
    Dim destRow As Long, srcLastRow As Long
    Dim colFirst As Long, colLast As Long
    Dim numCols() As Long
    Dim numColCount As Long
    Dim rateCols() As Long
    Dim rateColCount As Long
    Dim rateLeftNum() As Long
    Dim rateRightNum() As Long
    Dim destHdrs() As String
    Dim destHdrCount As Long
    Dim i As Long, j As Long
    Dim srcColMap() As Long
    Dim cv As Variant
    Dim sN As String, snN As String
    Dim rawStr As String
    Dim firstDataRow As Long, lastDataRow As Long
    Dim cAddr As String, dAddr As String
    Dim nonEmpty As Long
    Dim fmt As String
    Dim colIsRate() As Boolean
    Dim tryRow As Long
    Dim allFound As Boolean
    Dim matchHdrs() As String
    Dim matchColsInDest() As Long
    Dim matchCount As Long
    Dim keyN As String
    Dim lCol As Long, rCol As Long
    Dim lAddr As String, rAddr2 As String
    Dim dc As Long
    Dim lc2 As Long, rc2 As Long
    Dim la2 As String, ra2 As String
    Dim prevNum1 As Long, prevNum2 As Long
    Dim srcColFirst As Long, srcColLast As Long
    Dim rateDestCol As Long

    Set destWs = ActiveSheet

    hdrRow = 0
    Set ur = destWs.UsedRange
    For r = ur.Row To ur.Row + ur.rows.Count - 1
        nonEmpty = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            If Trim(CStr(destWs.Cells(r, c).Value)) <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    colFirst = ur.Column
    colLast = ur.Column + ur.Columns.Count - 1

    destHdrCount = colLast - colFirst + 1
    ReDim destHdrs(1 To destHdrCount)
    For i = 1 To destHdrCount
        destHdrs(i) = CStr(destWs.Cells(hdrRow, colFirst + i - 1).Value)
    Next i

    ReDim numCols(1 To destHdrCount)
    ReDim rateCols(1 To destHdrCount)
    ReDim rateLeftNum(1 To destHdrCount)
    ReDim rateRightNum(1 To destHdrCount)
    ReDim colIsRate(1 To destHdrCount)
    numColCount = 0
    rateColCount = 0

    Dim wb As Workbook
    Set wb = ActiveSheet.Parent

    ' 率列判定: 見出し行の書式 → 見出し行の下のデータ行の書式 → 他シートのデータ行の書式 の順で確認
    For i = 1 To destHdrCount
        fmt = destWs.Cells(hdrRow, colFirst + i - 1).NumberFormat
        If InStr(fmt, "%") > 0 Then
            colIsRate(i) = True
        Else
            colIsRate(i) = False
        End If
    Next i

    Dim dataCheckRow As Long
    dataCheckRow = hdrRow + 1
    If dataCheckRow <= ur.Row + ur.rows.Count - 1 Then
        For i = 1 To destHdrCount
            If Not colIsRate(i) Then
                fmt = destWs.Cells(dataCheckRow, colFirst + i - 1).NumberFormat
                If InStr(fmt, "%") > 0 Then
                    colIsRate(i) = True
                End If
            End If
        Next i
    End If

    ' 率列がまだ不明なら他シートを走査して書式確認
    Dim needRateCheck As Boolean
    needRateCheck = False
    For i = 2 To destHdrCount
        If Not colIsRate(i) Then needRateCheck = True
    Next i

    If needRateCheck Then
        Dim checkWs As Worksheet
        Dim cur As Range
        Dim foundRate As Boolean
        foundRate = False
        For Each checkWs In wb.Worksheets
            If foundRate Then Exit For
            If checkWs.Name = destWs.Name Then GoTo NextCheckSheet
            sN = checkWs.Name
            s = sN: GoSub 正規化: snN = s
            If InStr(snN, "記入例") > 0 Then GoTo NextCheckSheet
            If InStr(snN, "例") > 0 Then GoTo NextCheckSheet
            If InStr(snN, "説明") > 0 Then GoTo NextCheckSheet
            If InStr(snN, "様式") > 0 Then GoTo NextCheckSheet
            On Error Resume Next
            Set cur = Nothing
            Set cur = checkWs.UsedRange
            On Error GoTo 0
            If cur Is Nothing Then GoTo NextCheckSheet
            If cur.rows.Count < 2 Then GoTo NextCheckSheet
            ' 見出し行を探す
            Dim chkHdrRow As Long
            chkHdrRow = 0
            Dim chkR As Long
            For chkR = cur.Row To cur.Row + cur.rows.Count - 1
                Dim chkNE As Long
                chkNE = 0
                For c = cur.Column To cur.Column + cur.Columns.Count - 1
                    If Trim(CStr(checkWs.Cells(chkR, c).Value)) <> "" Then chkNE = chkNE + 1
                Next c
                If chkNE >= 2 Then
                    chkHdrRow = chkR
                    Exit For
                End If
            Next chkR
            If chkHdrRow = 0 Then GoTo NextCheckSheet
            Dim chkDataRow As Long
            chkDataRow = chkHdrRow + 1
            If chkDataRow > cur.Row + cur.rows.Count - 1 Then GoTo NextCheckSheet
            ' 各列の書式を確認
            For i = 1 To destHdrCount
                If Not colIsRate(i) Then
                    ' この列に対応する他シートの列を見出し名で探す
                    s = destHdrs(i): GoSub 正規化: key = s
                    For ci = cur.Column To cur.Column + cur.Columns.Count - 1
                        Dim tmpS As String
                        tmpS = CStr(checkWs.Cells(chkHdrRow, ci).Value): GoSub 正規化_tmpS
                        If tmpS = key Then
                            fmt = checkWs.Cells(chkDataRow, ci).NumberFormat
                            If InStr(fmt, "%") > 0 Then
                                colIsRate(i) = True
                                foundRate = True
                            End If
                            Exit For
                        End If
                    Next ci
                End If
            Next i
NextCheckSheet:
        Next checkWs
    End If

    numColCount = 0
    For i = 2 To destHdrCount
        If Not colIsRate(i) Then
            numColCount = numColCount + 1
            numCols(numColCount) = i
        End If
    Next i

    rateColCount = 0
    For i = 2 To destHdrCount
        If colIsRate(i) Then
            rateColCount = rateColCount + 1
            rateCols(rateColCount) = i
            prevNum1 = 0: prevNum2 = 0
            For j = i - 1 To 2 Step -1
                If Not colIsRate(j) Then
                    If prevNum1 = 0 Then
                        prevNum1 = j
                    ElseIf prevNum2 = 0 Then
                        prevNum2 = j
                        Exit For
                    End If
                End If
            Next j
            rateRightNum(rateColCount) = prevNum1
            rateLeftNum(rateColCount) = prevNum2
        End If
    Next i

    ' 既存データ消去
    Dim lastR As Long
    lastR = destWs.UsedRange.Row + destWs.UsedRange.rows.Count - 1
    If lastR > hdrRow Then
        destWs.rows((hdrRow + 1) & ":" & lastR).Delete Shift:=xlUp
    End If

    destRow = hdrRow + 1

    matchCount = 0
    ReDim matchHdrs(1 To destHdrCount)
    ReDim matchColsInDest(1 To destHdrCount)
    For i = 2 To destHdrCount
        If Not colIsRate(i) Then
            matchCount = matchCount + 1
            s = destHdrs(i): GoSub 正規化: key = s
            matchHdrs(matchCount) = key
            matchColsInDest(matchCount) = i
        End If
    Next i

    For Each ws In wb.Worksheets
        If ws.Name = destWs.Name Then GoTo NextSheet

        sN = ws.Name
        s = sN: GoSub 正規化: snN = s
        If InStr(snN, "記入例") > 0 Then GoTo NextSheet
        If InStr(snN, "例") > 0 Then GoTo NextSheet
        If InStr(snN, "説明") > 0 Then GoTo NextSheet
        If InStr(snN, "様式") > 0 Then GoTo NextSheet

        On Error Resume Next
        Set srcUr = Nothing
        Set srcUr = ws.UsedRange
        On Error GoTo 0
        If srcUr Is Nothing Then GoTo NextSheet

        srcHdrRow = 0
        srcColFirst = srcUr.Column
        srcColLast = srcUr.Column + srcUr.Columns.Count - 1

        For tryRow = srcUr.Row To srcUr.Row + srcUr.rows.Count - 1
            ReDim srcColMap(1 To matchCount)
            allFound = True
            For i = 1 To matchCount
                srcColMap(i) = 0
            Next i
            For ci = srcColFirst To srcColLast
                cv = ws.Cells(tryRow, ci).Value
                s = CStr(cv): GoSub 正規化: key = s
                For i = 1 To matchCount
                    If srcColMap(i) = 0 And key = matchHdrs(i) Then
                        srcColMap(i) = ci
                        Exit For
                    End If
                Next i
            Next ci
            For i = 1 To matchCount
                If srcColMap(i) = 0 Then
                    allFound = False
                    Exit For
                End If
            Next i
            If allFound Then
                srcHdrRow = tryRow
                Exit For
            End If
        Next tryRow

        If srcHdrRow = 0 Then GoTo NextSheet

        srcLastRow = srcUr.Row + srcUr.rows.Count - 1

        For ri = srcHdrRow + 1 To srcLastRow
            cv = ws.Cells(ri, srcColMap(1)).Value
            If IsError(cv) Then GoTo NextSrcRow
            s = CStr(cv): GoSub 正規化: key = s
            If key = "" Then GoTo NextSrcRow
            keyN = Trim(StrConv(Replace(Replace(Replace(CStr(cv), " ", ""), Chr(12288), ""), Chr(9), ""), vbNarrow))
            If keyN = "合計" Or keyN = "計" Or keyN = "小計" Or keyN = "総計" Then GoTo NextSrcRow
            If Len(keyN) >= 2 And Right(keyN, 1) = "計" Then GoTo NextSrcRow

            destWs.Cells(destRow, colFirst).Value = ws.Name

            For i = 1 To matchCount
                cv = ws.Cells(ri, srcColMap(i)).Value
                If IsError(cv) Then
                    If i = 1 Then
                        destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = ""
                    Else
                        destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = 0
                    End If
                Else
                    If i > 1 Then
                        If IsNumeric(cv) Then
                            destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = CDbl(cv)
                        Else
                            rawStr = CStr(cv)
                            rawStr = Replace(rawStr, ",", "")
                            rawStr = Replace(rawStr, Chr(65292), "")
                            rawStr = Replace(rawStr, "円", "")
                            rawStr = Replace(rawStr, " ", "")
                            rawStr = Replace(rawStr, Chr(12288), "")
                            rawStr = Trim(StrConv(rawStr, vbNarrow))
                            If rawStr = "" Then
                                destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = 0
                            ElseIf IsNumeric(rawStr) Then
                                destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = CDbl(rawStr)
                            Else
                                destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = cv
                            End If
                        End If
                    Else
                        destWs.Cells(destRow, colFirst + matchColsInDest(i) - 1).Value = cv
                    End If
                End If
            Next i

            For i = 1 To rateColCount
                lCol = colFirst + rateLeftNum(i) - 1
                rCol = colFirst + rateRightNum(i) - 1
                lAddr = destWs.Cells(destRow, lCol).Address(False, False)
                rAddr2 = destWs.Cells(destRow, rCol).Address(False, False)
                rateDestCol = colFirst + rateCols(i) - 1
                destWs.Cells(destRow, rateDestCol).NumberFormat = "General"
                destWs.Cells(destRow, rateDestCol).Formula = _
                    "=IF(" & lAddr & "=0,0," & rAddr2 & "/" & lAddr & ")"
            Next i

            destRow = destRow + 1
NextSrcRow:
        Next ri
NextSheet:
    Next ws

    ' 合計行
    If destRow > hdrRow + 1 Then
        firstDataRow = hdrRow + 1
        lastDataRow = destRow - 1
        destWs.Cells(destRow, colFirst).Value = "合計"
        For i = 2 To matchCount
            dc = matchColsInDest(i)
            cAddr = destWs.Cells(firstDataRow, colFirst + dc - 1).Address(False, False)
            dAddr = destWs.Cells(lastDataRow, colFirst + dc - 1).Address(False, False)
            destWs.Cells(destRow, colFirst + dc - 1).Formula = "=SUM(" & cAddr & ":" & dAddr & ")"
        Next i
        For i = 1 To rateColCount
            lc2 = colFirst + rateLeftNum(i) - 1
            rc2 = colFirst + rateRightNum(i) - 1
            la2 = destWs.Cells(destRow, lc2).Address(False, False)
            ra2 = destWs.Cells(destRow, rc2).Address(False, False)
            rateDestCol = colFirst + rateCols(i) - 1
            destWs.Cells(destRow, rateDestCol).NumberFormat = "General"
            destWs.Cells(destRow, rateDestCol).Formula = _
                "=IF(" & la2 & "=0,0," & ra2 & "/" & la2 & ")"
        Next i
    End If

    Exit Sub
正規化:
    s = Trim$(StrConv(s, vbNarrow)): s = Replace(s, ",", "")
    Return
正規化_tmpS:
    tmpS = Trim$(StrConv(tmpS, vbNarrow)): tmpS = Replace(tmpS, ",", "")
    Return
End Sub

Sub 金額を按分表にする()
    ' 依頼の語: 按分|人数で割り振|人数比で配分|基準の数で按分|人数割り|共通経費を人数|共通経費|割り振
    ' 扱う: 按分 配分
    ' 見出し: なし
    ' 形: 別のシート
    ' 金額列を別シートの基準表（項目・数の列）の比率で按分した表を右に作成。端数は基準最大項目に加算。合計行付き。

    Dim ws As Object, wsP As Object, sh As Object
    Dim wb As Object
    Set wb = ActiveSheet.Parent
    Set ws = ActiveSheet

    Dim s As String
    Dim ri As Long, ci As Long, i As Long

    ' ── 基準シートを形で探す
    Set wsP = Nothing
    Dim bestScore As Long: bestScore = -1
    For Each sh In wb.Sheets
        If sh.Name = ws.Name Then GoTo nextSh
        Dim pur As Object
        Set pur = sh.UsedRange
        Dim purRows As Long, purCols As Long
        purRows = pur.rows.Count: purCols = pur.Columns.Count
        If purCols < 2 Or purRows < 2 Then GoTo nextSh
        Dim purR2 As Long, purC2 As Long
        purR2 = pur.Row: purC2 = pur.Column
        Dim hasStr As Boolean, hasNum As Boolean
        hasStr = False: hasNum = False
        For ri = purR2 + 1 To purR2 + purRows - 1
            Dim v1 As String, v2 As String
            v1 = Trim(CStr(sh.Cells(ri, purC2).Value))
            v2 = Trim(CStr(sh.Cells(ri, purC2 + 1).Value))
            If v1 <> "" Then hasStr = True
            If IsNumeric(v2) And v2 <> "" Then hasNum = True
        Next ri
        Dim sc As Long: sc = 0
        If hasStr And hasNum Then sc = 10
        sc = sc + (10 - purCols)
        If sc > bestScore Then
            bestScore = sc
            Set wsP = sh
        End If
nextSh:
    Next sh
    If wsP Is Nothing Then Exit Sub

    ' ── 基準表の項目列・数値列を形で探す
    Dim pur3 As Object
    Set pur3 = wsP.UsedRange
    Dim pR As Long, pc As Long, pRows As Long, pCols As Long
    pR = pur3.Row: pc = pur3.Column
    pRows = pur3.rows.Count: pCols = pur3.Columns.Count

    Dim pHRow As Long, pNameCol As Long, pNumCol As Long
    pHRow = pR: pNameCol = pc: pNumCol = 0
    Dim bestNumCol As Long, bestNumCnt As Long
    bestNumCol = 0: bestNumCnt = 0
    For ci = pc To pc + pCols - 1
        Dim numCnt As Long: numCnt = 0
        For ri = pR + 1 To pR + pRows - 1
            Dim sv As String
            sv = Trim(CStr(wsP.Cells(ri, ci).Value))
            s = sv: GoSub 正規化num
            If s <> "" And IsNumeric(s) Then numCnt = numCnt + 1
        Next ri
        If numCnt > bestNumCnt Then
            bestNumCnt = numCnt
            bestNumCol = ci
        End If
    Next ci
    If bestNumCol = 0 Then Exit Sub
    pNumCol = bestNumCol
    Dim bestStrCol As Long, bestStrCnt As Long
    bestStrCol = 0: bestStrCnt = 0
    For ci = pc To pNumCol - 1
        Dim strCnt As Long: strCnt = 0
        For ri = pR + 1 To pR + pRows - 1
            Dim sv2 As String
            sv2 = Trim(CStr(wsP.Cells(ri, ci).Value))
            If sv2 <> "" And Not IsNumeric(sv2) Then strCnt = strCnt + 1
        Next ri
        If strCnt > bestStrCnt Then
            bestStrCnt = strCnt
            bestStrCol = ci
        End If
    Next ci
    If bestStrCol = 0 Then pNameCol = pc Else pNameCol = bestStrCol

    ' ── 基準表のデータ読み込み（空行・合計行・注記行を除く）
    Dim depNames() As String
    Dim depNums() As Double
    Dim depCount As Long: depCount = 0
    Dim totalPeople As Double: totalPeople = 0
    For ri = pHRow + 1 To pR + pRows - 1
        Dim pn As String
        pn = Trim(CStr(wsP.Cells(ri, pNameCol).Value))
        If pn = "" Then GoTo nextPR
        s = pn: GoSub 正規化hdr
        If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then GoTo nextPR
        If Left(s, 1) = "※" Then GoTo nextPR
        sv = Trim(CStr(wsP.Cells(ri, pNumCol).Value))
        s = sv: GoSub 正規化num
        If s = "" Or Not IsNumeric(s) Then GoTo nextPR
        Dim pv As Double: pv = 0
        On Error Resume Next: pv = CDbl(s): On Error GoTo 0
        depCount = depCount + 1
        ReDim Preserve depNames(1 To depCount)
        ReDim Preserve depNums(1 To depCount)
        depNames(depCount) = pn
        depNums(depCount) = pv
        totalPeople = totalPeople + pv
nextPR:
    Next ri
    If depCount = 0 Then Exit Sub

    Dim maxIdx As Long: maxIdx = 1
    For i = 2 To depCount
        If depNums(i) > depNums(maxIdx) Then maxIdx = i
    Next i

    ' ── メインシートの表を形で探す（見出し行＋データ行）
    Dim ur As Object
    Set ur = ws.UsedRange
    Dim urR As Long, urC As Long, urRows As Long, urCols As Long
    urR = ur.Row: urC = ur.Column
    urRows = ur.rows.Count: urCols = ur.Columns.Count

    Dim hRow As Long, feeCol As Long, amtCol As Long
    hRow = 0: feeCol = 0: amtCol = 0

    Dim tryRow As Long
    For tryRow = urR To urR + urRows - 2
        Dim rStrCol As Long, rNumCol As Long
        rStrCol = 0: rNumCol = 0
        For ci = urC To urC + urCols - 1
            Dim cv As String
            cv = Trim(CStr(ws.Cells(tryRow + 1, ci).Value))
            If cv = "" Then GoTo nextCI
            s = cv: GoSub 正規化num
            If IsNumeric(s) And s <> "" Then
                If rNumCol = 0 Then rNumCol = ci
            Else
                If rStrCol = 0 Then rStrCol = ci
            End If
nextCI:
        Next ci
        If rStrCol > 0 And rNumCol > 0 Then
            hRow = tryRow: feeCol = rStrCol: amtCol = rNumCol
            Exit For
        End If
    Next tryRow
    If hRow = 0 Then Exit Sub

    ' ── メインシートの最右端（本表の実際の右端列）を求める
    ' 見出し行のうち空でないセルの最大列（ただし出力列は除外するため、まず本表範囲の右端を確定）
    ' 本表の右端 = feeCol～amtCol の範囲（amtCol が最右端と仮定）
    ' ただし表に3列以上ある場合（例: 項目・金額・内容）はamtColより右に列があるので確認
    Dim mainRightCol As Long
    mainRightCol = amtCol
    ' 見出し行でamtColより右に連続した非空列があれば本表に含む（ただし数字でない見出しのみ）
    Dim checkC As Long
    For checkC = amtCol + 1 To urC + urCols - 1
        Dim hdrV As String
        hdrV = Trim(CStr(ws.Cells(hRow, checkC).Value))
        If hdrV = "" Then
            Exit For
        End If
        ' この列のデータが数値ならamtColを更新しない（別表の可能性）
        ' 文字列見出しで数値データでなければ本表の一部とみなす
        Dim isNumCol As Boolean: isNumCol = True
        Dim checkR As Long
        For checkR = hRow + 1 To urR + urRows - 1
            Dim chkV As String
            chkV = Trim(CStr(ws.Cells(checkR, checkC).Value))
            If chkV <> "" Then
                s = chkV: GoSub 正規化num
                If Not IsNumeric(s) Or s = "" Then
                    isNumCol = False
                    Exit For
                End If
            End If
        Next checkR
        If Not isNumCol Then
            mainRightCol = checkC
        End If
    Next checkC

    ' ── メインシートのデータ行収集（空行・合計行を除く）
    Dim dataRows() As Long
    Dim dataCount As Long: dataCount = 0
    For ri = hRow + 1 To urR + urRows - 1
        Dim fv As String
        fv = Trim(CStr(ws.Cells(ri, feeCol).Value))
        If fv = "" Then GoTo nextScan
        s = fv: GoSub 正規化hdr
        If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then GoTo nextScan
        dataCount = dataCount + 1
        ReDim Preserve dataRows(1 To dataCount)
        dataRows(dataCount) = ri
nextScan:
    Next ri
    If dataCount = 0 Then Exit Sub

    ' ── 出力開始列の決定（既存の按分表を探して消す）
    ' 出力開始列 = mainRightCol + 2（1列空けて）
    Dim outStartCol As Long
    ' 既存按分表の探索: mainRightCol+2 から始まる領域に値があれば消す
    Dim existOut As Long: existOut = 0
    Dim sc2 As Long
    For sc2 = mainRightCol + 2 To mainRightCol + depCount + 20
        Dim hv As String
        hv = Trim(CStr(ws.Cells(hRow, sc2).Value))
        If hv <> "" Then
            existOut = sc2
            Exit For
        End If
    Next sc2

    If existOut > 0 Then
        ws.Range(ws.Cells(hRow, existOut), ws.Cells(hRow + dataCount + 5, existOut + depCount + 10)).ClearContents
        outStartCol = existOut
    Else
        outStartCol = mainRightCol + 2
    End If

    ' ── 見出し行を書く
    ws.Cells(hRow, outStartCol).Value = ws.Cells(hRow, feeCol).Value
    For i = 1 To depCount
        ws.Cells(hRow, outStartCol + i).Value = depNames(i)
    Next i
    ws.Cells(hRow, outStartCol + depCount + 1).Value = "合計"

    ' ── データ行を書く
    Dim rowOut As Long: rowOut = hRow + 1
    Dim di As Long
    For di = 1 To dataCount
        Dim dr As Long: dr = dataRows(di)
        fv = Trim(CStr(ws.Cells(dr, feeCol).Value))
        Dim amtStr As String
        amtStr = Trim(CStr(ws.Cells(dr, amtCol).Value))
        s = amtStr: GoSub 正規化num
        Dim amt As Double: amt = 0
        If s <> "" And IsNumeric(s) Then
            On Error Resume Next: amt = CDbl(s): On Error GoTo 0
        End If

        ws.Cells(rowOut, outStartCol).Value = fv

        Dim allocated() As Long
        ReDim allocated(1 To depCount)
        Dim sumAlloc As Double: sumAlloc = 0
        For i = 1 To depCount
            If totalPeople = 0 Then
                allocated(i) = 0
            Else
                allocated(i) = CLng(Int(amt * depNums(i) / totalPeople))
            End If
            sumAlloc = sumAlloc + allocated(i)
        Next i
        Dim remainder As Long
        If totalPeople = 0 Then
            remainder = 0
        Else
            remainder = CLng(Int(amt)) - CLng(sumAlloc)
        End If
        allocated(maxIdx) = allocated(maxIdx) + remainder

        For i = 1 To depCount
            ws.Cells(rowOut, outStartCol + i).Value = allocated(i)
        Next i

        Dim c1 As String, c2 As String
        c1 = ws.Cells(rowOut, outStartCol + 1).Address(False, False)
        c2 = ws.Cells(rowOut, outStartCol + depCount).Address(False, False)
        ws.Cells(rowOut, outStartCol + depCount + 1).Formula = "=SUM(" & c1 & ":" & c2 & ")"

        rowOut = rowOut + 1
    Next di

    ' ── 合計行
    ws.Cells(rowOut, outStartCol).Value = "合計"
    Dim ra1 As String, ra2 As String
    For i = 1 To depCount + 1
        Dim colAddr As Long: colAddr = outStartCol + i
        ra1 = ws.Cells(hRow + 1, colAddr).Address(False, False)
        ra2 = ws.Cells(rowOut - 1, colAddr).Address(False, False)
        ws.Cells(rowOut, colAddr).Formula = "=SUM(" & ra1 & ":" & ra2 & ")"
    Next i

    Exit Sub

正規化num:
    s = Trim(s)
    s = StrConv(s, vbNarrow)
    s = Replace(s, ",", "")
    s = Replace(s, ChrW(165), "")
    s = Replace(s, "円", "")
    s = Replace(s, "人", "")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, Chr(9), "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, "　", "")
    Dim negFlag As Boolean: negFlag = False
    If Left(s, 1) = "△" Or Left(s, 1) = "▲" Then
        negFlag = True: s = Mid(s, 2)
    End If
    If negFlag And s <> "" Then s = "-" & s
    Return

正規化hdr:
    s = Trim(s)
    s = StrConv(s, vbNarrow)
    s = Replace(s, Chr(10), "")
    s = Replace(s, Chr(13), "")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, Chr(9), "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, "　", "")
    Dim parenPos As Long: parenPos = InStr(s, "(")
    If parenPos > 1 Then
        Dim sfx As String: sfx = Mid(s, parenPos)
        If sfx = "(円)" Or sfx = "(千円)" Or sfx = "(人)" Or sfx = "(計)" Then
            s = Left(s, parenPos - 1)
        End If
    End If
    Return
End Sub

Sub シート間比較表を作る()
    ' 依頼の語: 比較表を右|増減額と増減率|増減を出して|前年度と比べ|増減率の表|こちらと相手|比較表を作って|相手のシート|前年度比|前年度との比較|対前年度|今年度と前年度
    ' 扱う: 比較表 増減額 増減率 合計 突き合
    ' 見出し: なし
    ' 形: 別のシート 共通の見出しのシート
    ' アクティブシートと同じ見出し構造の別シートを項目列で突き合わせ右に比較表を作る

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim wbk As Workbook
    Set wbk = ws.Parent

    Dim s As String
    Dim nkIn As String, nkOut As String

    ' --- 見出し行・項目列・金額列を探す（アクティブシート）---
    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urR As Long, urC As Long, urRows As Long, urCols As Long
    urR = ur.Row: urC = ur.Column
    urRows = ur.rows.Count: urCols = ur.Columns.Count

    Dim hdrRow As Long, nameCol As Long, amtCol As Long
    Dim nlIn As String, numLike As Boolean
    hdrRow = 0: nameCol = 0: amtCol = 0
    Dim r As Long, c As Long

    Dim rr As Long
    For rr = urR To urR + urRows - 1
        Dim txtCount As Long, numCount As Long
        txtCount = 0: numCount = 0
        Dim firstTxtCol As Long
        firstTxtCol = 0
        For c = urC To urC + urCols - 1
            Dim cv As String
            cv = CStr(ws.Cells(rr, c).Value)
            If cv <> "" Then
                s = cv: GoSub 正規化数値チェック
                If IsNumeric(s) Then
                    numCount = numCount + 1
                Else
                    txtCount = txtCount + 1
                    If firstTxtCol = 0 Then firstTxtCol = c
                End If
            End If
        Next c
        If txtCount >= 2 And numCount = 0 And firstTxtCol > 0 Then
            ' 次の行にデータがあるか確認
            Dim NextDataRow As Long
            NextDataRow = 0
            Dim rr2 As Long
            For rr2 = rr + 1 To urR + urRows - 1
                Dim hasNum As Boolean, hasTxt As Boolean
                hasNum = False: hasTxt = False
                For c = urC To urC + urCols - 1
                    Dim nv As String
                    nv = CStr(ws.Cells(rr2, c).Value)
                    If nv <> "" Then
                        s = nv: GoSub 正規化数値チェック
                        If IsNumeric(s) Then
                            hasNum = True
                        Else
                            hasTxt = True
                        End If
                    End If
                Next c
                If hasTxt Or hasNum Then
                    If hasNum Then
                        NextDataRow = rr2
                    End If
                    Exit For
                End If
            Next rr2
            If NextDataRow > 0 Then
                hdrRow = rr
                nameCol = firstTxtCol
                Exit For
            End If
        End If
    Next rr

    If hdrRow = 0 Then
        hdrRow = urR
        nameCol = urC
    End If

    ' 見出しの行の直し（表題や単位の行を見出しにしない）: 数のある最初の行のすぐ上を見出しにする
    Dim fixR As Long, firstNumR As Long, fcH As Long, cCnt As Long, vtH As Long
    firstNumR = 0
    For fixR = hdrRow To urR + urRows - 1
        For fcH = urC To urC + urCols - 1
            vtH = VarType(ws.Cells(fixR, fcH).Value)
            If vtH = vbDouble Or vtH = vbSingle Or vtH = vbInteger Or vtH = vbLong Or vtH = vbCurrency Or vtH = vbDate Then
                firstNumR = fixR
                Exit For
            ElseIf vtH = vbString Then
                ' 金額が「425,000円」「８９３，０００」のような文字でも、その行は明細＝見出しではない
                nlIn = CStr(ws.Cells(fixR, fcH).Value): GoSub 数の文字か
                If numLike Then
                    firstNumR = fixR
                    Exit For
                End If
            End If
        Next fcH
        If firstNumR > 0 Then Exit For
    Next fixR
    If firstNumR > hdrRow + 1 Then
        cCnt = 0
        For fcH = urC To urC + urCols - 1
            If CStr(ws.Cells(firstNumR - 1, fcH).Value) <> "" Then cCnt = cCnt + 1
        Next fcH
        If cCnt >= 2 Then
            hdrRow = firstNumR - 1
            For fcH = urC To urC + urCols - 1
                If CStr(ws.Cells(hdrRow, fcH).Value) <> "" Then
                    nameCol = fcH
                    Exit For
                End If
            Next fcH
        End If
    End If

    For c = nameCol + 1 To urC + urCols - 1
        Dim nv2 As String
        nv2 = CStr(ws.Cells(hdrRow + 1, c).Value): s = nv2: GoSub 正規化数値チェック
        If IsNumeric(s) Then
            amtCol = c
            Exit For
        End If
    Next c
    If amtCol = 0 Then amtCol = nameCol + 1

    ' --- 相手シートを探す ---
    Dim wsPrev As Worksheet
    Dim bestScore As Long
    bestScore = -1
    Dim sh As Worksheet
    For Each sh In wbk.Worksheets
        If sh.Name = ws.Name Then GoTo NextSheet
        Dim ur2 As Range
        Set ur2 = sh.UsedRange
        Dim ur2R As Long, ur2C As Long, ur2Rows As Long, ur2Cols As Long
        ur2R = ur2.Row: ur2C = ur2.Column
        ur2Rows = ur2.rows.Count: ur2Cols = ur2.Columns.Count

        Dim pvHdr2 As Long
        pvHdr2 = 0
        Dim rrS As Long
        For rrS = ur2R To ur2R + ur2Rows - 1
            Dim tc2 As Long, nc2 As Long, fc2 As Long
            tc2 = 0: nc2 = 0: fc2 = 0
            For c = ur2C To ur2C + ur2Cols - 1
                Dim cv2 As String
                cv2 = CStr(sh.Cells(rrS, c).Value)
                If cv2 <> "" Then
                    s = cv2: GoSub 正規化数値チェック
                    If IsNumeric(s) Then
                        nc2 = nc2 + 1
                    Else
                        tc2 = tc2 + 1
                        If fc2 = 0 Then fc2 = c
                    End If
                End If
            Next c
            If tc2 >= 2 And nc2 = 0 And fc2 > 0 Then
                Dim nextDR As Long
                nextDR = 0
                Dim rrS2 As Long
                For rrS2 = rrS + 1 To ur2R + ur2Rows - 1
                    Dim hN2 As Boolean, hT2 As Boolean
                    hN2 = False: hT2 = False
                    For c = ur2C To ur2C + ur2Cols - 1
                        Dim nv3 As String
                        nv3 = CStr(sh.Cells(rrS2, c).Value)
                        If nv3 <> "" Then
                            s = nv3: GoSub 正規化数値チェック
                            If IsNumeric(s) Then hN2 = True Else hT2 = True
                        End If
                    Next c
                    If hT2 Or hN2 Then
                        If hN2 Then nextDR = rrS2
                        Exit For
                    End If
                Next rrS2
                If nextDR > 0 Then
                    pvHdr2 = rrS
                    Exit For
                End If
            End If
        Next rrS

        If pvHdr2 > 0 Then
            Dim score As Long
            score = 1
            If ur2Cols = urCols Then score = score + 2
            If Abs(ur2Rows - urRows) <= 5 Then score = score + 1
            If score > bestScore Then
                bestScore = score
                Set wsPrev = sh
            End If
        End If
NextSheet:
    Next sh

    If wsPrev Is Nothing Then
        For Each sh In wbk.Worksheets
            If sh.Name <> ws.Name Then
                Set wsPrev = sh
                Exit For
            End If
        Next sh
    End If
    If wsPrev Is Nothing Then Exit Sub

    ' --- 相手シートの構造 ---
    Set ur2 = wsPrev.UsedRange
    ur2R = ur2.Row: ur2C = ur2.Column
    ur2Rows = ur2.rows.Count: ur2Cols = ur2.Columns.Count

    Dim pvHdrRow As Long, pvNameCol As Long, pvAmtCol As Long
    pvHdrRow = 0: pvNameCol = 0: pvAmtCol = 0

    For rr = ur2R To ur2R + ur2Rows - 1
        Dim tc3 As Long, nc3 As Long, fc3 As Long
        tc3 = 0: nc3 = 0: fc3 = 0
        For c = ur2C To ur2C + ur2Cols - 1
            Dim cv3 As String
            cv3 = CStr(wsPrev.Cells(rr, c).Value)
            If cv3 <> "" Then
                s = cv3: GoSub 正規化数値チェック
                If IsNumeric(s) Then
                    nc3 = nc3 + 1
                Else
                    tc3 = tc3 + 1
                    If fc3 = 0 Then fc3 = c
                End If
            End If
        Next c
        If tc3 >= 2 And nc3 = 0 And fc3 > 0 Then
            Dim nextDR3 As Long
            nextDR3 = 0
            Dim rr3 As Long
            For rr3 = rr + 1 To ur2R + ur2Rows - 1
                Dim hN3 As Boolean, hT3 As Boolean
                hN3 = False: hT3 = False
                For c = ur2C To ur2C + ur2Cols - 1
                    Dim nv4 As String
                    nv4 = CStr(wsPrev.Cells(rr3, c).Value)
                    If nv4 <> "" Then
                        s = nv4: GoSub 正規化数値チェック
                        If IsNumeric(s) Then hN3 = True Else hT3 = True
                    End If
                Next c
                If hT3 Or hN3 Then
                    If hN3 Then nextDR3 = rr3
                    Exit For
                End If
            Next rr3
            If nextDR3 > 0 Then
                pvHdrRow = rr
                pvNameCol = fc3
                Exit For
            End If
        End If
    Next rr

    If pvHdrRow = 0 Then
        pvHdrRow = ur2R
        pvNameCol = ur2C
    End If

    For c = pvNameCol + 1 To ur2C + ur2Cols - 1
        Dim nv5 As String
        nv5 = CStr(wsPrev.Cells(pvHdrRow + 1, c).Value): s = nv5: GoSub 正規化数値チェック
        If IsNumeric(s) Then
            pvAmtCol = c
            Exit For
        End If
    Next c
    If pvAmtCol = 0 Then pvAmtCol = pvNameCol + 1

    ' --- アクティブシートのデータ読み込み ---
    Dim dictCurr As Object
    Set dictCurr = CreateObject("Scripting.Dictionary")
    Dim currRawKeys() As String
    Dim currCount As Long
    currCount = 0

    r = hdrRow + 1
    Dim blankCount As Long
    blankCount = 0
    Do While r <= urR + urRows - 1
        Dim rawName As String
        rawName = CStr(ws.Cells(r, nameCol).Value)
        nkIn = rawName: GoSub normKey
        Dim cellN As String
        cellN = nkOut
        If cellN = "" Then
            blankCount = blankCount + 1
            If blankCount >= 3 Then Exit Do
        Else
            ' 合計・計・小計・総計の行はスキップ
            Dim isTotalRow As Boolean
            isTotalRow = (cellN = "合計" Or cellN = "計" Or cellN = "小計" Or cellN = "総計")
            If Not isTotalRow Then
                blankCount = 0
                If Not dictCurr.Exists(cellN) Then
                    s = CStr(ws.Cells(r, amtCol).Value): GoSub 正規化数値チェック
                    dictCurr(cellN) = CDbl(s)
                    currCount = currCount + 1
                    ReDim Preserve currRawKeys(1 To currCount)
                    currRawKeys(currCount) = rawName
                End If
            End If
        End If
        r = r + 1
    Loop

    ' --- 相手シートのデータ読み込み ---
    Dim dictPrev As Object
    Set dictPrev = CreateObject("Scripting.Dictionary")
    Dim prevNormKeys() As String
    Dim prevRawKeys() As String
    Dim prevKeyCount As Long
    prevKeyCount = 0

    r = pvHdrRow + 1
    Dim blankCount2 As Long
    blankCount2 = 0
    Do While r <= ur2R + ur2Rows - 1
        Dim rawName2 As String
        rawName2 = CStr(wsPrev.Cells(r, pvNameCol).Value)
        nkIn = rawName2: GoSub normKey
        Dim pvCellN As String
        pvCellN = nkOut
        If pvCellN = "" Then
            blankCount2 = blankCount2 + 1
            If blankCount2 >= 3 Then Exit Do
        Else
            Dim isTotalRow2 As Boolean
            isTotalRow2 = (pvCellN = "合計" Or pvCellN = "計" Or pvCellN = "小計" Or pvCellN = "総計")
            If Not isTotalRow2 Then
                blankCount2 = 0
                If Not dictPrev.Exists(pvCellN) Then
                    s = CStr(wsPrev.Cells(r, pvAmtCol).Value): GoSub 正規化数値チェック
                    dictPrev(pvCellN) = CDbl(s)
                    prevKeyCount = prevKeyCount + 1
                    ReDim Preserve prevNormKeys(1 To prevKeyCount)
                    ReDim Preserve prevRawKeys(1 To prevKeyCount)
                    prevNormKeys(prevKeyCount) = pvCellN
                    prevRawKeys(prevKeyCount) = rawName2
                End If
            End If
        End If
        r = r + 1
    Loop

    ' 相手にだけある項目
    Dim prevOnlyNorm() As String
    Dim prevOnlyRaw() As String
    Dim prevOnlyCount As Long
    prevOnlyCount = 0
    Dim ki As Long
    For ki = 1 To prevKeyCount
        If Not dictCurr.Exists(prevNormKeys(ki)) Then
            prevOnlyCount = prevOnlyCount + 1
            ReDim Preserve prevOnlyNorm(1 To prevOnlyCount)
            ReDim Preserve prevOnlyRaw(1 To prevOnlyCount)
            prevOnlyNorm(prevOnlyCount) = prevNormKeys(ki)
            prevOnlyRaw(prevOnlyCount) = prevRawKeys(ki)
        End If
    Next ki

    ' --- 出力列 ---
    Dim tableRightCol As Long
    tableRightCol = amtCol
    If nameCol > tableRightCol Then tableRightCol = nameCol
    Dim outCol As Long
    outCol = tableRightCol + 2

    ' --- 既存の比較表を消す ---
    Dim clearEndCol As Long
    clearEndCol = outCol + 4
    Dim anyExist As Boolean
    anyExist = False
    Dim cClr As Long
    For cClr = outCol To clearEndCol
        If CStr(ws.Cells(hdrRow, cClr).Value) <> "" Then anyExist = True
    Next cClr
    If anyExist Then
        Dim clearEndRow As Long
        clearEndRow = hdrRow
        Dim rClr As Long
        For rClr = hdrRow + 1 To hdrRow + 5000
            Dim anyVal As Boolean
            anyVal = False
            For cClr = outCol To clearEndCol
                If CStr(ws.Cells(rClr, cClr).Value) <> "" Then
                    anyVal = True
                    Exit For
                End If
            Next cClr
            If anyVal Then
                clearEndRow = rClr
            Else
                Exit For
            End If
        Next rClr
        ws.Range(ws.Cells(hdrRow, outCol), ws.Cells(clearEndRow, clearEndCol)).ClearContents
    End If

    ' --- 見出し行を書く ---
    Dim colHdrName As String
    colHdrName = CStr(ws.Cells(hdrRow, nameCol).Value)
    If Trim$(colHdrName) = "" Then colHdrName = CStr(wsPrev.Cells(pvHdrRow, pvNameCol).Value)

    ws.Cells(hdrRow, outCol).Value = colHdrName
    ws.Cells(hdrRow, outCol + 1).Value = ws.Name
    ws.Cells(hdrRow, outCol + 2).Value = wsPrev.Name
    ws.Cells(hdrRow, outCol + 3).Value = "増減額"
    ws.Cells(hdrRow, outCol + 4).Value = "増減率"

    ' --- データ行を書く ---
    Dim dataStartRow As Long
    dataStartRow = hdrRow + 1
    Dim rowIdx As Long
    rowIdx = 0

    Dim i As Long
    For i = 1 To currCount
        nkIn = currRawKeys(i): GoSub normKey
        Dim kName As String
        kName = nkOut
        rowIdx = rowIdx + 1
        Dim rOut As Long
        rOut = dataStartRow + rowIdx - 1
        Dim vC As Double, vP As Double
        vC = dictCurr(kName)
        If dictPrev.Exists(kName) Then vP = dictPrev(kName) Else vP = 0
        Dim diff As Double
        diff = vC - vP
        ws.Cells(rOut, outCol).Value = currRawKeys(i)
        ws.Cells(rOut, outCol + 1).Value = vC
        ws.Cells(rOut, outCol + 2).Value = vP
        ws.Cells(rOut, outCol + 3).Value = diff
        If vP = 0 Then
            If vC = 0 Then
                ws.Cells(rOut, outCol + 4).Value = 0
            Else
                ws.Cells(rOut, outCol + 4).Value = "皆増"
            End If
        ElseIf vC = 0 Then
            ws.Cells(rOut, outCol + 4).Value = "皆減"
        Else
            ws.Cells(rOut, outCol + 4).Value = diff / vP
        End If
    Next i

    For i = 1 To prevOnlyCount
        rowIdx = rowIdx + 1
        rOut = dataStartRow + rowIdx - 1
        Dim kNameP As String
        kNameP = prevOnlyNorm(i)
        vC = 0
        vP = dictPrev(kNameP)
        diff = vC - vP
        ws.Cells(rOut, outCol).Value = prevOnlyRaw(i)
        ws.Cells(rOut, outCol + 1).Value = 0
        ws.Cells(rOut, outCol + 2).Value = vP
        ws.Cells(rOut, outCol + 3).Value = diff
        ws.Cells(rOut, outCol + 4).Value = "皆減"
    Next i

    ' --- 合計行 ---
    Dim totalRow As Long
    totalRow = dataStartRow + rowIdx
    Dim r1 As Long, r2 As Long
    r1 = dataStartRow: r2 = dataStartRow + rowIdx - 1
    ws.Cells(totalRow, outCol).Value = "合計"
    ws.Cells(totalRow, outCol + 1).Value = "=SUM(" & ws.Cells(r1, outCol + 1).Address & ":" & ws.Cells(r2, outCol + 1).Address & ")"
    ws.Cells(totalRow, outCol + 2).Value = "=SUM(" & ws.Cells(r1, outCol + 2).Address & ":" & ws.Cells(r2, outCol + 2).Address & ")"
    ws.Cells(totalRow, outCol + 3).Value = "=SUM(" & ws.Cells(r1, outCol + 3).Address & ":" & ws.Cells(r2, outCol + 3).Address & ")"

    Dim sumPrev As Double
    Dim sumDiff As Double
    sumPrev = ws.Cells(totalRow, outCol + 2).Value
    sumDiff = ws.Cells(totalRow, outCol + 3).Value
    If sumPrev = 0 Then
        If sumDiff > 0 Then
            ws.Cells(totalRow, outCol + 4).Value = "皆増"
        ElseIf sumDiff < 0 Then
            ws.Cells(totalRow, outCol + 4).Value = "皆減"
        Else
            ws.Cells(totalRow, outCol + 4).Value = 0
        End If
    Else
        ws.Cells(totalRow, outCol + 4).Value = sumDiff / sumPrev
    End If

    Exit Sub

数の文字か:
    ' 数として読める文字かどうかだけを返す（正規化数値チェックは数でない文字を "0" にするので使えない）
    numLike = False
    Dim nlS As String
    nlS = Trim$(nlIn)
    nlS = Replace(nlS, Chr(160), "")
    Dim nlW As String
    nlW = ""
    Dim nlCi As Long
    For nlCi = 1 To Len(nlS)
        Dim nlCh As String
        nlCh = Mid$(nlS, nlCi, 1)
        Dim nlCode As Long
        nlCode = AscW(nlCh)
        If nlCode >= &HFF01 And nlCode <= &HFF5E Then
            nlW = nlW & Chr(nlCode - &HFF00 + &H20)
        Else
            nlW = nlW & nlCh
        End If
    Next nlCi
    nlS = nlW
    nlS = Replace(nlS, ",", "")
    nlS = Replace(nlS, "円", "")
    nlS = Replace(nlS, " ", "")
    nlS = Replace(nlS, "　", "")
    If nlS <> "" Then numLike = IsNumeric(nlS)
    Return

正規化数値チェック:
    s = Trim$(s)
    s = Replace(s, Chr(160), "")
    Dim sW As String
    sW = ""
    Dim ci As Long
    For ci = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, ci, 1)
        Dim code As Long
        code = AscW(ch)
        If code >= &HFF01 And code <= &HFF5E Then
            sW = sW & Chr(code - &HFF00 + &H20)
        Else
            sW = sW & ch
        End If
    Next ci
    s = sW
    s = Replace(s, ",", "")
    s = Replace(s, "円", "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    If Not IsNumeric(s) Then s = "0"
    Return

normKey:
    nkOut = Trim$(nkIn)
    Dim nkW As String
    nkW = ""
    Dim nkCi As Long
    For nkCi = 1 To Len(nkOut)
        Dim nkCh As String
        nkCh = Mid$(nkOut, nkCi, 1)
        Dim nkCode As Long
        nkCode = AscW(nkCh)
        If nkCode >= &HFF01 And nkCode <= &HFF5E Then
            nkW = nkW & Chr(nkCode - &HFF00 + &H20)
        Else
            nkW = nkW & nkCh
        End If
    Next nkCi
    nkOut = nkW
    nkOut = Replace(nkOut, ",", "")
    nkOut = Replace(nkOut, "円", "")
    nkOut = Replace(nkOut, " ", "")
    nkOut = Replace(nkOut, "　", "")
    nkOut = Trim$(nkOut)
    Return

End Sub

Sub 円グラフを作る()
    ' 依頼の語: 円グラフ|内訳を円|割合を円|構成比を円|構成比の円|構成比のグラフ
    ' 扱う: 円グラフ 構成比 合計
    ' 見出し: なし
    ' 形: 数の列
    ' 左端の文字列列を項目、右端の数値列（%列除く）を値にした円グラフを作る。合計行・0値行は除外し既存グラフは削除して作り直す

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange

    Dim rr As Long, cc As Long
    Dim hdrRow As Long
    hdrRow = 0

    Dim firstDataRow As Long
    firstDataRow = ur.Row
    Dim foundHdr As Boolean
    foundHdr = False
    For rr = ur.Row To ur.Row + ur.rows.Count - 1
        Dim hasNum As Boolean
        hasNum = False
        For cc = ur.Column To ur.Column + ur.Columns.Count - 1
            If IsNumeric(ws.Cells(rr, cc).Value) And CStr(ws.Cells(rr, cc).Value) <> "" Then
                hasNum = True
                Exit For
            End If
        Next cc
        If hasNum Then
            hdrRow = rr - 1
            If hdrRow < ur.Row Then hdrRow = ur.Row
            firstDataRow = rr
            foundHdr = True
            Exit For
        End If
    Next rr
    If Not foundHdr Then Exit Sub

    ' 表題行を探す（見出し行より上で空でない単独セル）
    Dim titleText As String
    titleText = ""
    Dim tr As Long
    For tr = ur.Row To hdrRow - 1
        Dim tval As String
        tval = Trim(CStr(ws.Cells(tr, ur.Column).Value))
        If tval <> "" Then
            titleText = tval
        End If
    Next tr

    Dim colLabel As Long
    colLabel = 0
    For cc = ur.Column To ur.Column + ur.Columns.Count - 1
        If Not IsNumeric(ws.Cells(firstDataRow, cc).Value) Or CStr(ws.Cells(firstDataRow, cc).Value) = "" Then
            If CStr(ws.Cells(firstDataRow, cc).Value) <> "" Then
                colLabel = cc
                Exit For
            End If
        End If
    Next cc
    If colLabel = 0 Then Exit Sub

    Dim colVal As Long
    colVal = 0
    Dim fmt As String
    For cc = ur.Column + ur.Columns.Count - 1 To ur.Column Step -1
        fmt = ws.Cells(hdrRow, cc).NumberFormat
        If InStr(fmt, "%") > 0 Then GoTo NextCol
        fmt = ws.Cells(firstDataRow, cc).NumberFormat
        If InStr(fmt, "%") > 0 Then GoTo NextCol
        If IsNumeric(ws.Cells(firstDataRow, cc).Value) And CStr(ws.Cells(firstDataRow, cc).Value) <> "" Then
            colVal = cc
            Exit For
        End If
NextCol:
    Next cc
    If colVal = 0 Then Exit Sub

    Dim seriesName As String
    seriesName = CStr(ws.Cells(hdrRow, colVal).Value)
    If seriesName = "" Then seriesName = "値"

    Dim lastRow As Long
    lastRow = firstDataRow
    For rr = firstDataRow To ur.Row + ur.rows.Count - 1
        If CStr(ws.Cells(rr, colLabel).Value) = "" And CStr(ws.Cells(rr, colVal).Value) = "" Then
            ' 空行でも続きがあるか確認
            Dim hasMore As Boolean
            hasMore = False
            Dim rr2 As Long
            For rr2 = rr + 1 To ur.Row + ur.rows.Count - 1
                Dim lbl2 As String
                lbl2 = Trim(CStr(ws.Cells(rr2, colLabel).Value))
                Dim v2 As String
                v2 = Trim(CStr(ws.Cells(rr2, colVal).Value))
                If lbl2 <> "" Or v2 <> "" Then
                    hasMore = True
                    Exit For
                End If
            Next rr2
            If Not hasMore Then Exit For
        Else
            lastRow = rr
        End If
    Next rr

    Dim labels() As String
    Dim vals() As Double
    Dim cnt As Long
    cnt = 0
    ReDim labels(1 To lastRow - firstDataRow + 1)
    ReDim vals(1 To lastRow - firstDataRow + 1)

    Dim lbl As String
    Dim v As Double
    Dim s As String
    For rr = firstDataRow To lastRow
        lbl = Trim(CStr(ws.Cells(rr, colLabel).Value))
        If lbl = "" Then GoTo skipRow
        s = lbl
        GoSub 正規化
        If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then GoTo skipRow
        If Not IsNumeric(ws.Cells(rr, colVal).Value) Then GoTo skipRow
        v = CDbl(ws.Cells(rr, colVal).Value)
        If v = 0 Then GoTo skipRow
        cnt = cnt + 1
        labels(cnt) = lbl
        vals(cnt) = v
skipRow:
    Next rr
    If cnt = 0 Then Exit Sub

    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim rightMost As Long
    rightMost = ur.Column + ur.Columns.Count - 1
    Dim chartLeft As Double
    chartLeft = ws.Cells(firstDataRow, rightMost + 2).Left
    Dim chartTop As Double
    chartTop = ws.Cells(hdrRow, rightMost + 2).Top

    Set co = ws.ChartObjects.Add(chartLeft, chartTop, 300, 250)
    Dim ch As Chart
    Set ch = co.Chart
    ch.ChartType = xlPie

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.SeriesCollection.NewSeries
    Dim sr As Series
    Set sr = ch.SeriesCollection(1)

    Dim arrV() As Double
    Dim arrL() As String
    ReDim arrV(1 To cnt)
    ReDim arrL(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        arrV(i) = vals(i)
        arrL(i) = labels(i)
    Next i
    sr.Values = arrV
    sr.XValues = arrL
    sr.Name = seriesName

    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"

    ch.HasTitle = True
    If titleText <> "" Then
        ch.chartTitle.text = titleText
    Else
        ch.chartTitle.text = seriesName & "の構成比"
    End If
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Size = 18
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Bold = True

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight

    sr.HasDataLabels = True
    Dim dl As DataLabels
    Set dl = sr.DataLabels
    dl.ShowPercentage = True
    dl.ShowValue = False
    dl.ShowCategoryName = False
    dl.ShowSeriesName = False

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    Return

End Sub

Sub 選択列でピボットを作る()
    ' 依頼の語: ピボットで集計|行と列|列でピボット|課と科目
    ' 扱う: ピボット クロス集計 合計 テーブルに 集計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列 日付の列
    ' 選んだ3列(1:行,2:列,3:値)で新シートにピボットテーブルを作る。合計行を除いた明細範囲を使う。

    Dim ws As Worksheet
    Dim ur As Range
    Dim colRow As Long, colCol As Long, colVal As Long
    Dim hRow As Long, lastRow As Long
    Dim rr As Long, cc As Long, fcol As Long, frow As Long, lastC As Long
    Dim s As String
    Dim wb As Workbook
    Dim pvSh As Worksheet
    Dim sh As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim fldRow As String, fldCol As String, fldVal As String
    Dim shName As String
    Dim dataAddr As String
    Dim numCount As Long, strCount As Long
    Dim v As Variant, cv As String
    Dim isTotRow As Boolean
    Dim maxRow As Long

    Set ws = ActiveSheet
    Set wb = ActiveWorkbook
    Set ur = ws.UsedRange
    fcol = ur.Column
    frow = ur.Row

    If Selection Is Nothing Then Exit Sub
    If Selection.Areas.Count < 3 Then Exit Sub
    colRow = Selection.Areas(1).Column
    colCol = Selection.Areas(2).Column
    colVal = Selection.Areas(3).Column

    lastC = fcol + ur.Columns.Count - 1

    hRow = 0
    For rr = frow To frow + ur.rows.Count - 1
        numCount = 0
        strCount = 0
        For cc = fcol To lastC
            v = ws.Cells(rr, cc).Value
            If IsNumeric(v) And Trim(CStr(v)) <> "" Then
                numCount = numCount + 1
            ElseIf Trim(CStr(v)) <> "" Then
                strCount = strCount + 1
            End If
        Next cc
        If strCount >= 2 And numCount = 0 Then
            hRow = rr
            Exit For
        End If
    Next rr
    If hRow = 0 Then Exit Sub

    fldRow = Trim(CStr(ws.Cells(hRow, colRow).Value))
    fldCol = Trim(CStr(ws.Cells(hRow, colCol).Value))
    fldVal = Trim(CStr(ws.Cells(hRow, colVal).Value))

    lastRow = hRow
    maxRow = frow + ur.rows.Count - 1
    For rr = hRow + 1 To maxRow
        s = Trim(CStr(ws.Cells(rr, colRow).Value))
        isTotRow = False
        For cc = fcol To lastC
            cv = Trim(CStr(ws.Cells(rr, cc).Value))
            If cv = "合計" Or cv = "計" Or cv = "小計" Or cv = "総計" Then
                isTotRow = True
                Exit For
            End If
        Next cc
        If Not isTotRow And s <> "" Then
            lastRow = rr
        End If
    Next rr
    If lastRow <= hRow Then Exit Sub

    dataAddr = "'" & ws.Name & "'!" & ws.Range(ws.Cells(hRow, fcol), ws.Cells(lastRow, lastC)).Address(True, True)

    shName = fldRow & "別" & fldCol & "別"

    Application.DisplayAlerts = False
    For Each sh In wb.Sheets
        If sh.Name = shName Then
            sh.Delete
            Exit For
        End If
    Next sh
    Application.DisplayAlerts = True

    Set pvSh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    pvSh.Name = shName

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataAddr)
    Set pt = pc.CreatePivotTable(TableDestination:=pvSh.Cells(3, 1), TableName:=shName & "PT")

    With pt
        .ManualUpdate = True
        .AddFields RowFields:=fldRow, ColumnFields:=fldCol
        With .PivotFields(fldVal)
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0"
        End With
        .RowGrand = True
        .ColumnGrand = True
        .ManualUpdate = False
    End With

End Sub

Sub 表をテーブル化して集計行を出す()
    ' 依頼の語: いつものテーブル|集計行も|集計行つき|集計行を出|集計行|職員名簿
    ' 扱う: テーブル化 集計行 フィルタ 合計 計算 テーブルに
    ' 見出し: なし
    ' 形: 数の列 日付の列
    ' 表をテーブルに変換し集計行を追加する（数値列は合計、番号・コード・年度・%列は集計なし）

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange

    Dim urR As Long, urC As Long, urRows As Long, urCols As Long
    urR = ur.Row: urC = ur.Column
    urRows = ur.rows.Count: urCols = ur.Columns.Count

    ' --- 見出し行を探す（文字列セルが2つ以上ある最初の行）
    Dim hdrRow As Long, r As Long, c As Long
    hdrRow = 0
    For r = urR To urR + urRows - 1
        Dim nonNumCount As Long
        nonNumCount = 0
        For c = urC To urC + urCols - 1
            Dim cv As String
            cv = Trim(CStr(ws.Cells(r, c).Value))
            If cv <> "" Then
                If Not IsNumeric(cv) And Not IsDate(cv) Then nonNumCount = nonNumCount + 1
            End If
        Next c
        If nonNumCount >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    Dim firstCol As Long, lastCol As Long
    firstCol = urC
    lastCol = urC + urCols - 1

    ' --- 見出し行の右端（空でない最後の列）
    Dim hdrLastCol As Long
    hdrLastCol = firstCol
    For c = firstCol To lastCol
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) <> "" Then hdrLastCol = c
    Next c

    ' 見出しの値を配列に保持（見出し行の繰り返し検出用）
    Dim hdrVals() As String
    ReDim hdrVals(firstCol To hdrLastCol)
    For c = firstCol To hdrLastCol
        hdrVals(c) = Trim(CStr(ws.Cells(hdrRow, c).Value))
    Next c

    ' --- データ最終行を探す
    ' 空行・※行・合計行・見出し繰り返し行の手前まで
    ' ただし途中の空行（前後にデータあり）はスキップしてデータとして含む
    ' → テーブル範囲はデータ最終行まで（空行も含む）
    ' 方針: 末尾から遡って最後のデータ行を探す（途中の空行は無視）
    '       ただし ※行・合計行・見出し繰り返し行は除外（その行より前まで）

    ' まず末尾方向に合計行・※行・見出し繰り返し行の位置を記録
    Dim cutRow As Long
    cutRow = urR + urRows ' この行以降は含めない（初期値は範囲外）

    For r = hdrRow + 1 To urR + urRows - 1
        ' ※行チェック
        Dim isNote As Boolean
        isNote = False
        For c = firstCol To hdrLastCol
            Dim cs As String
            cs = Trim(CStr(ws.Cells(r, c).Value))
            If cs <> "" Then
                If Left(cs, 1) = "※" Then isNote = True
                Exit For
            End If
        Next c
        If isNote Then
            If r < cutRow Then cutRow = r
            Exit For
        End If

        ' 見出し繰り返し行チェック（見出しの列と同じ値が並んでいる）
        Dim isHdrRepeat As Boolean
        isHdrRepeat = True
        Dim hasAny As Boolean
        hasAny = False
        For c = firstCol To hdrLastCol
            Dim rv As String
            rv = Trim(CStr(ws.Cells(r, c).Value))
            If rv <> "" Then hasAny = True
            If rv <> hdrVals(c) Then isHdrRepeat = False
        Next c
        If isHdrRepeat And hasAny Then
            If r < cutRow Then cutRow = r
            Exit For
        End If

        ' 合計行チェック（行の最初の非空セルが合計系語）
        Dim firstCellStr As String
        firstCellStr = ""
        For c = firstCol To hdrLastCol
            firstCellStr = Trim(CStr(ws.Cells(r, c).Value))
            If firstCellStr <> "" Then Exit For
        Next c
        If firstCellStr = "合計" Or firstCellStr = "計" Or firstCellStr = "小計" Or firstCellStr = "総計" Then
            If r < cutRow Then cutRow = r
            Exit For
        End If
    Next r

    ' cutRow-1 までの範囲でデータ最終行（末尾から逆順に最初の非空行）
    Dim lastDataRow As Long
    lastDataRow = hdrRow
    Dim searchEnd As Long
    searchEnd = cutRow - 1
    If searchEnd < hdrRow Then searchEnd = hdrRow

    For r = searchEnd To hdrRow + 1 Step -1
        Dim rowHasData As Boolean
        rowHasData = False
        For c = firstCol To hdrLastCol
            If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then
                rowHasData = True
                Exit For
            End If
        Next c
        If rowHasData Then
            lastDataRow = r
            Exit For
        End If
    Next r

    ' テーブル範囲: 見出し行からlastDataRowまで（途中の空行も含む）
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(hdrRow, firstCol), ws.Cells(lastDataRow, hdrLastCol))

    ' --- 既存テーブル検索
    Dim lo As ListObject
    Dim existLO As ListObject
    Set existLO = Nothing
    For Each lo In ws.ListObjects
        Dim inter As Range
        On Error Resume Next
        Set inter = Intersect(lo.Range, tblRange)
        On Error GoTo 0
        If Not inter Is Nothing Then
            Set existLO = lo
            Exit For
        End If
    Next lo

    If existLO Is Nothing Then
        On Error Resume Next
        Set existLO = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        On Error GoTo 0
        If existLO Is Nothing Then Exit Sub
    Else
        existLO.ShowTotals = False
        On Error Resume Next
        existLO.Resize tblRange
        On Error GoTo 0
    End If

    ' --- テーブル名: T + シート名
    Dim newName As String
    newName = "T" & ws.Name
    On Error Resume Next
    existLO.Name = newName
    On Error GoTo 0

    existLO.TableStyle = "TableStyleMedium2"
    existLO.ShowTableStyleRowStripes = True
    existLO.ShowAutoFilter = True

    ' --- 集計行
    existLO.ShowTotals = True

    ' 各列の集計設定
    Dim lc As ListColumn
    For Each lc In existLO.ListColumns
        Dim hdrVal As String
        hdrVal = lc.Name

        ' 番号・No・コード・年度を含む見出しはID列
        Dim isIdCol As Boolean
        isIdCol = False
        If InStr(hdrVal, "番号") > 0 Or InStr(hdrVal, "No") > 0 Or InStr(hdrVal, "NO") > 0 Or _
           InStr(hdrVal, "no") > 0 Or InStr(hdrVal, "コード") > 0 Or InStr(hdrVal, "年度") > 0 Then
            isIdCol = True
        End If

        If isIdCol Then
            lc.TotalsCalculation = xlTotalsCalculationNone
        Else
            ' %列チェック（書式のみ）
            Dim isPct As Boolean
            isPct = False
            If Not lc.DataBodyRange Is Nothing Then
                Dim fmtStr As String
                fmtStr = lc.DataBodyRange.Cells(1, 1).NumberFormat
                If InStr(fmtStr, "%") > 0 Then isPct = True
            End If

            If isPct Then
                lc.TotalsCalculation = xlTotalsCalculationNone
            Else
                ' 数値列・日付列・文字列列を判定（空でない最初のセルで判断）
                Dim isNumCol As Boolean
                isNumCol = False
                If Not lc.DataBodyRange Is Nothing Then
                    Dim firstCell As Range
                    Set firstCell = Nothing
                    Dim rc3 As Range
                    For Each rc3 In lc.DataBodyRange.Cells
                        If Trim(CStr(rc3.Value)) <> "" Then
                            Set firstCell = rc3
                            Exit For
                        End If
                    Next rc3
                    If Not firstCell Is Nothing Then
                        If IsNumeric(firstCell.Value) Then
                            Dim fmt2 As String
                            fmt2 = firstCell.NumberFormat
                            If InStr(fmt2, "y") > 0 Or InStr(fmt2, "m") > 0 Or InStr(fmt2, "d") > 0 Then
                                isNumCol = False
                            Else
                                isNumCol = True
                            End If
                        End If
                    End If
                End If

                If isNumCol Then
                    lc.TotalsCalculation = xlTotalsCalculationSum
                Else
                    lc.TotalsCalculation = xlTotalsCalculationNone
                End If
            End If
        End If
    Next lc

End Sub

Sub 棒グラフを作る()
    ' 依頼の語: 見やすい棒|いつもの棒|見栄えのいい棒|いつもの型で棒|方で棒グラフ|課別支出|課別の支出|課ごとの支出額
    ' 扱う: 棒グラフ作成
    ' 見出し: なし
    ' 形: 数の列
    ' 左端の文字列列を項目・右端の数値列を値にした棒グラフを作る。既存グラフは削除して作り直す。

    Dim ws As Object
    Dim ur As Object
    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    Dim rr As Long, cc As Long
    Dim headerRow As Long, dataFirstRow As Long
    Dim nameCol As Long, valCol As Long

    Dim urR As Long, urC As Long, urRows As Long, urCols As Long
    urR = ur.Row: urC = ur.Column
    urRows = ur.rows.Count: urCols = ur.Columns.Count

    Dim r As Long
    dataFirstRow = 0
    For r = urR To urR + urRows - 1
        Dim hasNum As Boolean: hasNum = False
        For cc = urC To urC + urCols - 1
            If IsNumeric(ws.Cells(r, cc).Value) And Trim$(CStr(ws.Cells(r, cc).Value)) <> "" Then
                hasNum = True: Exit For
            End If
        Next cc
        If hasNum Then
            If r > urR Then headerRow = r - 1 Else headerRow = r
            dataFirstRow = r
            Exit For
        End If
    Next r
    If dataFirstRow = 0 Then Exit Sub

    nameCol = 0: valCol = 0
    Dim testR As Long: testR = dataFirstRow
    For cc = urC To urC + urCols - 1
        Dim cv As String: cv = Trim$(CStr(ws.Cells(testR, cc).Value))
        If cv <> "" And Not IsNumeric(ws.Cells(testR, cc).Value) Then
            If nameCol = 0 Then nameCol = cc
        End If
    Next cc
    For cc = urC + urCols - 1 To urC Step -1
        If IsNumeric(ws.Cells(testR, cc).Value) And Trim$(CStr(ws.Cells(testR, cc).Value)) <> "" Then
            valCol = cc: Exit For
        End If
    Next cc
    If nameCol = 0 Or valCol = 0 Then Exit Sub

    ' 系列名：見出し行のvalColセルの値をそのまま使う（改行含む）
    Dim seriesName As String: seriesName = ""
    If headerRow >= urR Then
        seriesName = CStr(ws.Cells(headerRow, valCol).Value)
    End If

    ' データ行収集：空白行をスキップし、合計行は除外
    Dim allNames() As String
    Dim allAmounts() As Double
    Dim n As Long: n = 0
    ReDim allNames(1 To urRows): ReDim allAmounts(1 To urRows)
    For r = dataFirstRow To urR + urRows - 1
        Dim nv As String: nv = Trim$(CStr(ws.Cells(r, nameCol).Value))
        If nv = "合計" Or nv = "計" Or nv = "小計" Or nv = "総計" Then
        ElseIf nv = "" Then
        ElseIf IsNumeric(ws.Cells(r, valCol).Value) And Trim$(CStr(ws.Cells(r, valCol).Value)) <> "" Then
            n = n + 1
            allNames(n) = nv
            allAmounts(n) = CDbl(ws.Cells(r, valCol).Value)
        End If
    Next r
    If n = 0 Then Exit Sub

    Dim names() As String
    Dim amounts() As Double
    ReDim names(1 To n): ReDim amounts(1 To n)
    Dim i As Long
    For i = 1 To n
        names(i) = allNames(i)
        amounts(i) = allAmounts(i)
    Next i

    ' 降順ソート
    Dim ti As Long, tn As String, td As Double
    For i = 1 To n - 1
        For ti = 1 To n - i
            If amounts(ti) < amounts(ti + 1) Then
                td = amounts(ti): amounts(ti) = amounts(ti + 1): amounts(ti + 1) = td
                tn = names(ti): names(ti) = names(ti + 1): names(ti + 1) = tn
            End If
        Next ti
    Next i

    ' タイトル：表題行があれば使う、なければ系列名（改行除去）
    Dim titleText As String: titleText = ""
    If headerRow > urR Then
        Dim tr As Long
        For tr = urR To headerRow - 1
            Dim tv As String: tv = Trim$(CStr(ws.Cells(tr, urC).Value))
            If tv <> "" And InStr(tv, "単位") = 0 And Left$(tv, 1) <> "（" And Left$(tv, 1) <> "(" Then
                titleText = tv: Exit For
            End If
        Next tr
    End If
    If titleText = "" Then
        Dim snFlat As String: snFlat = Replace(Replace(seriesName, Chr(10), ""), Chr(13), "")
        titleText = Trim$(snFlat)
    End If

    Dim j As Long
    For j = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(j).Delete
    Next j

    Dim leftPos As Double, topPos As Double
    Dim NextCol As Long: NextCol = valCol + 1
    leftPos = ws.Cells(headerRow, NextCol).Left + 5
    topPos = ws.Cells(headerRow, NextCol).Top

    Dim useHoriz As Boolean: useHoriz = (n >= 9)

    Dim newCO As Object
    Set newCO = ws.ChartObjects.Add(leftPos, topPos, 420, 300)
    Dim ch As Object: Set ch = newCO.Chart

    If useHoriz Then
        ch.ChartType = 57
    Else
        ch.ChartType = 51
    End If

    On Error Resume Next
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    On Error GoTo 0

    ch.SeriesCollection.NewSeries
    Dim sr As Object: Set sr = ch.SeriesCollection(1)
    sr.Values = amounts
    sr.XValues = names
    If seriesName <> "" Then sr.Name = seriesName

    ch.HasLegend = False

    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 50
    If useHoriz Then
        ch.Axes(1).ReversePlotOrder = True
    End If
    ch.Axes(2).TickLabels.NumberFormat = "#,##0"
    ch.Axes(2).HasMajorGridlines = False
    On Error GoTo 0

    If titleText <> "" Then
        ch.HasTitle = True
        ch.chartTitle.text = titleText
        On Error Resume Next
        ch.chartTitle.Characters.Font.Size = 14
        ch.chartTitle.Characters.Font.Bold = True
        On Error GoTo 0
    Else
        ch.HasTitle = False
    End If

    On Error Resume Next
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Dim pi As Long
    For pi = 1 To n
        If pi = 1 Then
            ch.SeriesCollection(1).Points(pi).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
        Else
            ch.SeriesCollection(1).Points(pi).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        End If
    Next pi
    ch.SeriesCollection(1).HasDataLabels = True
    On Error GoTo 0

End Sub

Sub 月別ピボットを作る()
    ' 依頼の語: 月別ピボット|いつもの月別|月ごとのピボット|月別のピボット
    ' 扱う: 月別ピボット
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列 日付の列
    ' 選んでいる3列（外側行・内側行・値）を使い、日付列を月でグループ化した月別ピボットをシートに作る

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim ur As Range
    Dim hRow As Long, lastRow As Long
    Dim rr As Long, cc As Long
    Dim colOuter As Long, colInner As Long, colVal As Long
    Dim urRow As Long, urCol As Long
    Dim ptSheet As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ptName As String
    Dim ptSheetName As String
    Dim sht As Worksheet
    Dim dataRange As String
    Dim fldOuter As String, fldInner As String, fldVal As String
    Dim filled As Long, txtCnt As Long
    Dim cv As Variant

    Set ws = ActiveSheet
    Set wb = ActiveWorkbook

    If Selection Is Nothing Then Exit Sub
    If Selection.Areas.Count < 3 Then Exit Sub
    colOuter = Selection.Areas(1).Column
    colInner = Selection.Areas(2).Column
    colVal = Selection.Areas(3).Column

    Set ur = ws.UsedRange
    urRow = ur.Row
    urCol = ur.Column
    hRow = 0

    For rr = urRow To urRow + ur.rows.Count - 1
        filled = 0
        For cc = urCol To urCol + ur.Columns.Count - 1
            If Trim(CStr(ws.Cells(rr, cc).Value)) <> "" Then filled = filled + 1
        Next cc
        If filled >= 3 Then
            txtCnt = 0
            For cc = urCol To urCol + ur.Columns.Count - 1
                cv = ws.Cells(rr, cc).Value
                If Trim(CStr(cv)) <> "" Then
                    If Not IsNumeric(cv) And Not IsDate(cv) Then txtCnt = txtCnt + 1
                End If
            Next cc
            If txtCnt >= 2 Then
                hRow = rr
                Exit For
            End If
        End If
    Next rr
    If hRow = 0 Then Exit Sub

    fldOuter = CStr(ws.Cells(hRow, colOuter).Value)
    fldInner = CStr(ws.Cells(hRow, colInner).Value)
    fldVal = CStr(ws.Cells(hRow, colVal).Value)
    If Trim(fldOuter) = "" Or Trim(fldInner) = "" Or Trim(fldVal) = "" Then Exit Sub

    lastRow = hRow
    Dim anyVal As Boolean, isTot As Boolean, tv As String
    For rr = hRow + 1 To urRow + ur.rows.Count - 1
        anyVal = False
        isTot = False
        For cc = urCol To urCol + ur.Columns.Count - 1
            tv = Trim(CStr(ws.Cells(rr, cc).Value))
            If tv <> "" Then
                anyVal = True
                ' 下の合計の行は明細ではない＝元の範囲に入れない
                If tv = "合計" Or tv = "計" Or tv = "小計" Or tv = "総計" Or tv = "累計" Then isTot = True
            End If
        Next cc
        If isTot Then Exit For
        If anyVal Then lastRow = rr
    Next rr
    If lastRow = hRow Then Exit Sub

    ' 日付列を特定（選んだ3列以外でIsDateな値がある列）
    Dim colDate As Long: colDate = 0
    Dim fldDate As String: fldDate = ""
    Dim dv As Variant
    For cc = urCol To urCol + ur.Columns.Count - 1
        If cc <> colOuter And cc <> colInner And cc <> colVal Then
            For rr = hRow + 1 To lastRow
                dv = ws.Cells(rr, cc).Value
                If Trim(CStr(dv)) <> "" Then
                    If IsDate(dv) Then
                        colDate = cc
                        fldDate = CStr(ws.Cells(hRow, cc).Value)
                    End If
                    Exit For
                End If
            Next rr
        End If
        If colDate > 0 Then Exit For
    Next cc

    dataRange = "'" & ws.Name & "'!" & _
        ws.Cells(hRow, urCol).Address(True, True) & ":" & _
        ws.Cells(lastRow, urCol + ur.Columns.Count - 1).Address(True, True)

    ptSheetName = Left(fldOuter & "別月別", 31)
    ptName = "月別ピボット"

    Application.DisplayAlerts = False
    For Each sht In wb.Worksheets
        If sht.Name = ptSheetName Then
            sht.Delete
            Exit For
        End If
    Next sht
    Application.DisplayAlerts = True

    Set ptSheet = wb.Worksheets.Add(After:=ws)
    ptSheet.Name = ptSheetName

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = pc.CreatePivotTable(TableDestination:=ptSheet.Cells(3, 1), TableName:=ptName)

    pt.TableStyle2 = "PivotStyleMedium9"

    With pt.PivotFields(fldOuter)
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields(fldInner)
        .Orientation = xlRowField
        .Position = 2
    End With

    ' 日付フィールドを列に配置してグループ化（月のみ）
    If fldDate <> "" Then
        On Error Resume Next
        With pt.PivotFields(fldDate)
            .Orientation = xlColumnField
            .Position = 1
        End With
        On Error GoTo 0

        On Error Resume Next
        pt.PivotFields(fldDate).dataRange.Cells(1).Group _
            Start:=True, End:=True, _
            Periods:=Array(False, False, False, False, True, False, False)
        On Error GoTo 0

        ' グループ化後：元の日付フィールドを隠し、月フィールドだけ列に残す
        On Error Resume Next
        Dim pfLoop As PivotField
        For Each pfLoop In pt.PivotFields
            If pfLoop.Name = fldDate And pfLoop.Orientation = xlColumnField Then
                pfLoop.Orientation = xlHidden
            End If
        Next pfLoop
        On Error GoTo 0

        ' "月 (日付フィールド名)" 形式のフィールドを列に配置
        On Error Resume Next
        Dim monthFldName As String
        Dim pfCheck As PivotField
        For Each pfCheck In pt.PivotFields
            If pfCheck.Orientation = xlHidden Or pfCheck.Orientation = xlColumnField Then
                If InStr(pfCheck.Name, "月") > 0 And pfCheck.Name <> fldDate Then
                    pfCheck.Orientation = xlColumnField
                    pfCheck.Position = 1
                End If
            End If
        Next pfCheck
        On Error GoTo 0
    End If

    ' 行の見出しは列ごとに表形式にする（pt.RowAxisLayout を撃つと見出しが列の名前に変わってしまう）
    On Error Resume Next
    With pt.PivotFields(fldOuter)
        .LayoutForm = xlTabular
        .RepeatLabels = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With pt.PivotFields(fldInner)
        .LayoutForm = xlTabular
        .RepeatLabels = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    On Error GoTo 0

    Dim df As Object
    Set df = pt.AddDataField(pt.PivotFields(fldVal), "合計 / " & fldVal, xlSum)
    df.NumberFormat = "#,##0"

    ' 空白アイテムを非表示
    On Error Resume Next
    Dim pf2 As PivotField
    Dim pi As PivotItem
    For Each pf2 In pt.PivotFields
        If pf2.Orientation = xlColumnField Then
            For Each pi In pf2.PivotItems
                If Trim(CStr(pi.Name)) = "(空白)" Or Trim(CStr(pi.Name)) = "" Then
                    pi.Visible = False
                End If
            Next pi
        End If
    Next pf2
    Dim pi2 As PivotItem
    For Each pi2 In pt.PivotFields(fldOuter).PivotItems
        If Trim(CStr(pi2.Name)) = "(空白)" Or Trim(CStr(pi2.Name)) = "" Then
            pi2.Visible = False
        End If
    Next pi2
    On Error GoTo 0

    pt.NullString = "0"
    pt.DisplayNullString = True

End Sub

Sub 横持ちテーブルを縦持ちに変換する()
    ' 依頼の語: 縦持ち|ピボット解除|縦に並|列のピボット解除|クエリで縦
    ' 扱う: 縦持ち変換 ピボット解除 パワークエリ 合計 並べ替
    ' 見出し: なし
    ' 形: 数の列 テーブル
    ' アクティブシートのテーブルをパワークエリで縦持ちに変換し、新シートに同名_縦持ちのテーブルとして読み込む

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim tblName As String
    Dim qName As String
    Dim sheetName As String
    Dim mCode As String
    Dim q As Object
    Dim newLo As ListObject
    Dim destWs As Worksheet

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    ' アクティブシートのテーブルを探す
    If ws.ListObjects.Count = 0 Then
        Dim usedRng As Range
        Set usedRng = ws.UsedRange
        Set tbl = ws.ListObjects.Add(1, usedRng, , 1)
        tbl.Name = "Table_" & ws.Name
    Else
        Set tbl = ws.ListObjects(1)
    End If

    If tbl Is Nothing Then Exit Sub

    tblName = tbl.Name
    qName = tblName & "_縦持ち"
    sheetName = tblName & "_縦持ち"

    ' キー列を特定（番号・No・コードを含まない最初の文字列列）
    Dim keyColName As String
    Dim lc As ListColumn
    Dim r As Long
    Dim c As Long

    ' テーブルのヘッダ行から条件に合う最初の列を探す
    Dim firstTextColName As String
    firstTextColName = ""
    For Each lc In tbl.ListColumns
        Dim h As String
        h = lc.Name
        ' 番号・No・コードを含む列は除外
        If InStr(h, "番号") > 0 Or InStr(h, "No") > 0 Or InStr(h, "no") > 0 Or _
           InStr(h, "コード") > 0 Or InStr(h, "合計") > 0 Or InStr(h, "総計") > 0 Or _
           h = "計" Or (Len(h) > 0 And Right(h, 1) = "計") Then
        Else
            firstTextColName = h
            Exit For
        End If
    Next lc
    If firstTextColName = "" Then Exit Sub
    keyColName = firstTextColName

    ' 数値列（UnpivotOtherColumns の対象）を特定
    ' キー列の右にある数値列のみ（文字列列は除く）
    Dim numColNames() As String
    Dim numColCount As Integer
    numColCount = 0
    Dim foundKey As Boolean
    foundKey = False
    For Each lc In tbl.ListColumns
        h = lc.Name
        If Not foundKey Then
            If h = keyColName Then foundKey = True
        Else
            ' 番号・No・コード・合計・計・総計を含む列は除外
            If InStr(h, "番号") > 0 Or InStr(h, "No") > 0 Or InStr(h, "no") > 0 Or _
               InStr(h, "コード") > 0 Or InStr(h, "合計") > 0 Or InStr(h, "総計") > 0 Or _
               h = "計" Or (Len(h) > 0 And Right(h, 1) = "計") Then
                ' 除外
            Else
                ' 列のデータが数値かどうか確認（データ行の最初の非空値で判断）
                Dim isNumCol As Boolean
                isNumCol = True
                Dim hasAnyVal As Boolean
                hasAnyVal = False
                Dim dr As Long
                For dr = 1 To lc.DataBodyRange.rows.Count
                    Dim cv As Variant
                    cv = lc.DataBodyRange.Cells(dr, 1).Value
                    If cv <> "" And Not isEmpty(cv) Then
                        hasAnyVal = True
                        If Not IsNumeric(cv) Then
                            isNumCol = False
                            Exit For
                        End If
                    End If
                Next dr
                ' 数値列のみ追加（値が全く無い列も数値列として扱う）
                If isNumCol Then
                    ReDim Preserve numColNames(numColCount)
                    numColNames(numColCount) = h
                    numColCount = numColCount + 1
                End If
            End If
        End If
    Next lc

    ' UnpivotOtherColumns ではなく、数値列のみを明示的に Unpivot する
    ' M コードで数値列リストを直接指定する方式に変更
    Dim numColList As String
    numColList = ""
    If numColCount > 0 Then
        Dim i As Integer
        For i = 0 To numColCount - 1
            If numColList <> "" Then numColList = numColList & ", "
            numColList = numColList & """" & numColNames(i) & """"
        Next i
    End If

    If numColList = "" Then Exit Sub

    mCode = "let" & vbLf & _
        "    Source = Excel.CurrentWorkbook(){[Name=""" & tblName & """]}[Content]," & vbLf & _
        "    NumCols = {" & numColList & "}," & vbLf & _
        "    KeyCol = """ & keyColName & """," & vbLf & _
        "    IsKeyTotal = (v) =>" & vbLf & _
        "        Text.Contains(Text.From(v ?? """"), ""合計"") or" & vbLf & _
        "        Text.Contains(Text.From(v ?? """"), ""総計"") or" & vbLf & _
        "        Text.From(v ?? """") = ""計"" or" & vbLf & _
        "        Text.EndsWith(Text.From(v ?? """"), ""計"")," & vbLf & _
        "    FilteredRows = Table.SelectRows(Source, each not IsKeyTotal(Record.Field(_, KeyCol)))," & vbLf & _
        "    Unpivoted = Table.Unpivot(FilteredRows, NumCols, ""項目"", ""値"")," & vbLf & _
        "    FilteredEmpty = Table.SelectRows(Unpivoted, each [値] <> null and [値] <> """")," & vbLf & _
        "    Reordered = Table.SelectColumns(FilteredEmpty, {KeyCol, ""項目"", ""値""})" & vbLf & _
        "in" & vbLf & _
        "    Reordered"

    ' 既存クエリを削除
    On Error Resume Next
    For Each q In wb.Queries
        If q.Name = qName Then
            q.Delete
            Exit For
        End If
    Next q
    On Error GoTo 0

    ' クエリ追加
    wb.Queries.Add qName, mCode

    ' 既存シートを削除
    Application.DisplayAlerts = False
    For Each ws2 In wb.Sheets
        If ws2.Name = sheetName Then
            ws2.Delete
            Exit For
        End If
    Next ws2
    Application.DisplayAlerts = True

    ' 新シート追加
    Set destWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    destWs.Name = sheetName

    ' ListObject でクエリを読み込む
    Set newLo = destWs.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName, _
        Destination:=destWs.Range("A1"))
    newLo.QueryTable.CommandType = 2
    newLo.QueryTable.CommandText = "SELECT * FROM [" & qName & "]"
    newLo.QueryTable.Refresh False
    newLo.Name = qName

End Sub

Sub マスタ参照列を計算列に入れる()
    ' 依頼の語: 計算列を足|マスタ参照|引く列|参照する列|コードの右|名前の列|構造化参照|名称を引|名前を引
    ' 扱う: マスタ参照列追加 計算列 構造化参照 テーブルに
    ' 見出し: なし
    ' 形: 数の列 日付の列 テーブル
    ' アクティブシートのテーブルのキー列右に、別テーブルから構造化参照で引く計算列を挿入する。未登録は「未登録」

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ListObjects.Count = 0 Then Exit Sub
    Dim lo As ListObject
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim mainColCount As Long
    mainColCount = lo.ListColumns.Count
    Dim ri As Long

    Dim bestScore As Long: bestScore = -1
    Dim keyColIdxMain As Long
    Dim masterKeyIdxBest As Long
    Dim valColName As String
    Dim finalMlo As ListObject

    Dim s As String

    Dim mws As Worksheet
    Dim mLO As ListObject
    For Each mws In wb.Worksheets
        For Each mLO In mws.ListObjects
            If mws.Name = ws.Name And mLO.Name = lo.Name Then GoTo NextMlo
            If mLO.ListColumns.Count < 2 Then GoTo NextMlo
            If mLO.DataBodyRange Is Nothing Then GoTo NextMlo

            Dim mc As Long
            For mc = 1 To mLO.ListColumns.Count
                Dim dicM As Object
                Set dicM = CreateObject("Scripting.Dictionary")
                Dim mr As Long
                For mr = 1 To mLO.DataBodyRange.rows.Count
                    s = CStr(mLO.DataBodyRange.Cells(mr, mc).Value): GoSub normKey
                    If s <> "" Then dicM(s) = 1
                Next mr
                If dicM.Count = 0 Then GoTo NextMc

                Dim mainC As Long
                For mainC = 1 To mainColCount
                    Dim score As Long: score = 0
                    For ri = 1 To lo.DataBodyRange.rows.Count
                        s = CStr(lo.DataBodyRange.Cells(ri, mainC).Value): GoSub normKey
                        If s <> "" And dicM.Exists(s) Then score = score + 1
                    Next ri

                    If score > bestScore Then
                        Dim mc2 As Long
                        Dim cVal As String: cVal = ""
                        For mc2 = 1 To mLO.ListColumns.Count
                            If mc2 <> mc Then
                                cVal = mLO.ListColumns(mc2).Name
                                Exit For
                            End If
                        Next mc2
                        If cVal <> "" Then
                            bestScore = score
                            keyColIdxMain = mainC
                            masterKeyIdxBest = mc
                            valColName = cVal
                            Set finalMlo = mLO
                        End If
                    End If
                Next mainC
NextMc:
            Next mc
NextMlo:
        Next mLO
    Next mws

    If finalMlo Is Nothing Then Exit Sub
    If valColName = "" Then Exit Sub

    Dim keyColName As String
    keyColName = lo.ListColumns(keyColIdxMain).Name

    Dim nameCol As ListColumn
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If lc.Name = valColName Then
            Set nameCol = lc
            Exit For
        End If
    Next lc

    If nameCol Is Nothing Then
        Dim insertPos As Long
        insertPos = keyColIdxMain + 1
        Set nameCol = lo.ListColumns.Add(insertPos)
        nameCol.Name = valColName
    End If

    Dim mloName As String
    mloName = finalMlo.Name
    Dim keyRef As String
    keyRef = "[@" & keyColName & "]"
    Dim masterKeyColName As String
    masterKeyColName = finalMlo.ListColumns(masterKeyIdxBest).Name

    Dim formulaStr As String
    formulaStr = "=IFERROR(INDEX(" & mloName & "[" & valColName & "]," & _
        "MATCH(" & keyRef & "," & mloName & "[" & masterKeyColName & "],0)),""未登録"")"

    On Error Resume Next
    nameCol.DataBodyRange.Formula = formulaStr
    On Error GoTo 0

    Exit Sub

normKey:
    Dim rawV As Variant
    rawV = s
    If isEmpty(rawV) Or IsNull(rawV) Or rawV = "" Then
        s = ""
    ElseIf IsNumeric(rawV) Then
        s = CStr(CLng(CDbl(rawV)))
    Else
        s = Trim(CStr(rawV))
    End If
    Return

End Sub

Sub ダッシュボードを新規シートに作る()
    ' 依頼の語: ダッシュボード|スライサー付き|ピボットとグラフ
    ' 扱う: ダッシュボード ピボット ピボットグラフ スライサー グラフ 集計 合計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列 日付の列
    ' 選んでいる3列（1=行フィールド,2=値フィールド,3=スライサーフィールド）で新シートにピボット・集合縦棒グラフ・スライサーを作る

    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim dashWs As Worksheet
    Dim ur As Range
    Dim colRow As Long, colVal As Long, colSlicer As Long
    Dim hdrRow As Long, lastRow As Long
    Dim r As Long
    Dim colFirst As Long, colLast As Long
    Dim cv As String
    Dim rangeAddr As String
    Dim dashName As String
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataField As PivotField
    Dim co As ChartObject
    Dim ch As Chart
    Dim slCache As SlicerCache
    Dim slObj As Slicer
    Dim slc As SlicerCache
    Dim graphLeft As Double, graphTop As Double, slLeft As Double
    Dim ptRight As Long
    Dim ptName As String
    Dim rowHdr As String, valHdr As String, slicerHdr As String
    Dim urFirst As Long, urLast As Long
    Dim v1 As String, v2 As String

    Set wb = ActiveWorkbook
    Set srcWs = ActiveSheet
    Set ur = srcWs.UsedRange

    If Selection.Areas.Count >= 3 Then
        colRow = Selection.Areas(1).Column
        colVal = Selection.Areas(2).Column
        colSlicer = Selection.Areas(3).Column
    ElseIf Selection.Areas.Count = 1 And Selection.Areas(1).Columns.Count >= 3 Then
        colRow = Selection.Areas(1).Columns(1).Column
        colVal = Selection.Areas(1).Columns(2).Column
        colSlicer = Selection.Areas(1).Columns(3).Column
    Else
        Exit Sub
    End If

    urFirst = ur.Row
    urLast = ur.Row + ur.rows.Count - 1
    colFirst = ur.Column
    colLast = ur.Column + ur.Columns.Count - 1

    ' 見出し行を探す：値列が非数値で行フィールド列が空でない最初の行
    hdrRow = 0
    For r = urFirst To urLast
        v1 = Trim(CStr(srcWs.Cells(r, colRow).Value))
        v2 = Trim(CStr(srcWs.Cells(r, colVal).Value))
        If v1 <> "" And v2 <> "" Then
            If Not IsNumeric(srcWs.Cells(r, colVal).Value) Or isEmpty(srcWs.Cells(r, colVal).Value) Then
                hdrRow = r
                Exit For
            End If
        End If
    Next r

    If hdrRow = 0 Then
        For r = urFirst To urLast
            If Trim(CStr(srcWs.Cells(r, colRow).Value)) <> "" Then
                hdrRow = r
                Exit For
            End If
        Next r
    End If
    If hdrRow = 0 Then Exit Sub

    rowHdr = Trim(CStr(srcWs.Cells(hdrRow, colRow).Value))
    valHdr = Trim(CStr(srcWs.Cells(hdrRow, colVal).Value))
    slicerHdr = Trim(CStr(srcWs.Cells(hdrRow, colSlicer).Value))

    ' 最終明細行：合計行の手前まで（空白行を超えて走査しない。ただし空白でも最終行まで含める）
    lastRow = hdrRow
    For r = hdrRow + 1 To urLast
        cv = Trim(CStr(srcWs.Cells(r, colRow).Value))
        If cv = "合計" Or cv = "計" Or cv = "小計" Or cv = "総計" Then Exit For
        If InStr(cv, "合計") > 0 And cv <> "" Then Exit For
        lastRow = r
    Next r

    If lastRow <= hdrRow Then Exit Sub

    ' 元の範囲は見出し行から最終明細行まで（列は表全体）
    rangeAddr = "'" & srcWs.Name & "'!" & _
        srcWs.Range(srcWs.Cells(hdrRow, colFirst), srcWs.Cells(lastRow, colLast)).Address(True, True, xlR1C1)

    dashName = rowHdr & "別ダッシュボード"

    ' 既存スライサーのうちこのシートの既存ダッシュボードシートに乗っているものを先に消す
    Application.DisplayAlerts = False
    Dim scDel As SlicerCache
    Dim slDel As Slicer
    For Each scDel In wb.SlicerCaches
        On Error Resume Next
        For Each slDel In scDel.Slicers
            If slDel.Parent.Name = dashName Then
                scDel.Delete
                Exit For
            End If
        Next slDel
        On Error GoTo 0
    Next scDel

    ' 同名シートを削除
    Dim existSh As Worksheet
    For Each existSh In wb.Worksheets
        If existSh.Name = dashName Then
            existSh.Delete
            Exit For
        End If
    Next existSh
    Application.DisplayAlerts = True

    ' 元シートに紐づく既存スライサーキャッシュを削除（前回マクロの残骸）
    Dim scSrc As SlicerCache
    Dim ptCheck As PivotTable
    For Each scSrc In wb.SlicerCaches
        On Error Resume Next
        Set ptCheck = Nothing
        Set ptCheck = scSrc.SourcePivotTable
        If Not ptCheck Is Nothing Then
            If ptCheck.Parent.Name = srcWs.Name Then
                scSrc.Delete
            End If
        End If
        On Error GoTo 0
    Next scSrc

    Set dashWs = wb.Worksheets.Add(After:=srcWs)
    dashWs.Name = dashName

    ptName = "PT_" & Left(rowHdr, 8)
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rangeAddr)
    Set pt = pc.CreatePivotTable(TableDestination:=dashWs.Range("A3"), TableName:=ptName)

    With pt.PivotFields(rowHdr)
        .Orientation = xlRowField
        .Position = 1
    End With

    Set dataField = pt.AddDataField(pt.PivotFields(valHdr), "合計 / " & valHdr, xlSum)
    dataField.NumberFormat = "#,##0"

    ptRight = pt.TableRange2.Column + pt.TableRange2.Columns.Count - 1
    graphLeft = dashWs.Cells(3, ptRight + 2).Left
    graphTop = dashWs.Cells(3, 1).Top

    Set co = dashWs.ChartObjects.Add(Left:=graphLeft, Top:=graphTop, Width:=300, Height:=200)
    Set ch = co.Chart
    ch.SetSourceData Source:=pt.TableRange1
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.chartTitle.text = rowHdr & "別" & valHdr
    ch.HasLegend = False

    Set slCache = wb.SlicerCaches.Add2(pt, slicerHdr)
    slLeft = graphLeft + 310
    Set slObj = slCache.Slicers.Add(dashWs, , slicerHdr & "スライサー", slicerHdr, graphTop, slLeft, 150, 200)

End Sub

Sub パワークエリで左外部結合する()
    ' 依頼の語: 左外部結合|マスタを結合|クエリでマージ|マスタを付け|キーで結合|クエリでマスタ|名前付
    ' 扱う: パワークエリ マージ 左外部結合 クエリ作成 テーブルに
    ' 見出し: なし
    ' 形: 数の列 日付の列 テーブル
    ' 何をするか: アクティブシートのテーブルと共通列（キー）を持つ別テーブルを左外部結合し、相手の残り1列を付けてパワークエリで新シートに読み込む

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsMaster As Worksheet
    Dim wsDest As Worksheet
    Dim srcLO As ListObject
    Dim mLO As ListObject
    Dim destLO As ListObject
    Dim qExist As Object
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long, j As Long, k As Long
    Dim srcTblName As String
    Dim mstTblName As String
    Dim destSheetName As String
    Dim qryName As String
    Dim keyColSrc As String
    Dim keyColMst As String
    Dim extraCol As String
    Dim srcCols() As String
    Dim nl As String
    Dim mFormula As String
    Dim connStr As String
    Dim srcHdrCount As Long
    Dim mstHdrCount As Long
    Dim ur As Range
    Dim found As Boolean
    Dim tmpLO As ListObject
    Dim mh As String
    Dim cell As Range
    Dim typeStr As String
    Dim srcTypeList As String
    Dim keyTypeStr As String
    Dim srcColList As String
    Dim commonCols() As String
    Dim commonCount As Long
    Dim wsExist As Worksheet
    Dim mstTypeList As String
    Dim dateColNames() As String
    Dim dateColCount As Long

    Set wb = ActiveWorkbook
    Set wsSource = ActiveSheet

    If wsSource.ListObjects.Count > 0 Then
        Set srcLO = wsSource.ListObjects(1)
    Else
        Set ur = wsSource.UsedRange
        If ur Is Nothing Then Exit Sub
        On Error Resume Next
        Set srcLO = wsSource.ListObjects.Add(xlSrcRange, ur, , xlYes)
        On Error GoTo 0
        If srcLO Is Nothing Then Exit Sub
    End If
    srcTblName = srcLO.Name

    srcHdrCount = srcLO.HeaderRowRange.Columns.Count
    ReDim srcCols(1 To srcHdrCount)
    For i = 1 To srcHdrCount
        srcCols(i) = CStr(srcLO.HeaderRowRange.Cells(1, i).Value)
    Next i

    Set mLO = Nothing
    Set wsMaster = Nothing
    keyColSrc = ""
    keyColMst = ""
    extraCol = ""

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = srcLO.Name Then GoTo NextLO
            If lo.DataBodyRange Is Nothing Then GoTo NextLO
            commonCount = 0
            mstHdrCount = lo.HeaderRowRange.Columns.Count
            ReDim commonCols(1 To mstHdrCount)
            For j = 1 To mstHdrCount
                mh = CStr(lo.HeaderRowRange.Cells(1, j).Value)
                For k = 1 To srcHdrCount
                    If mh = srcCols(k) Then
                        commonCount = commonCount + 1
                        commonCols(commonCount) = mh
                        Exit For
                    End If
                Next k
            Next j
            If commonCount >= 1 And mstHdrCount - commonCount >= 1 Then
                Dim eCol As String: eCol = ""
                For j = 1 To mstHdrCount
                    mh = CStr(lo.HeaderRowRange.Cells(1, j).Value)
                    found = False
                    For k = 1 To srcHdrCount
                        If mh = srcCols(k) Then found = True: Exit For
                    Next k
                    If Not found And mh <> "" Then
                        eCol = mh
                        Exit For
                    End If
                Next j
                If eCol <> "" Then
                    keyColSrc = commonCols(1)
                    keyColMst = commonCols(1)
                    extraCol = eCol
                    Set mLO = lo
                    Set wsMaster = ws
                    GoTo FoundMaster
                End If
            End If
NextLO:
        Next lo
    Next ws

    For Each ws In wb.Worksheets
        If ws.Name = wsSource.Name Then GoTo NextSheet2
        If ws.ListObjects.Count = 0 Then
            Dim mUR2 As Range
            Set mUR2 = ws.UsedRange
            If mUR2 Is Nothing Then GoTo NextSheet2
            If mUR2.rows.Count < 2 Then GoTo NextSheet2
            Dim mstHC2 As Long
            mstHC2 = mUR2.Columns.Count
            Dim cc2 As Long: cc2 = 0
            Dim kc2 As String: kc2 = ""
            Dim ec2 As String: ec2 = ""
            For j = 1 To mstHC2
                mh = CStr(mUR2.Cells(1, j).Value)
                found = False
                For k = 1 To srcHdrCount
                    If mh = srcCols(k) Then found = True: kc2 = mh: cc2 = cc2 + 1: Exit For
                Next k
                If Not found And mh <> "" Then ec2 = mh
            Next j
            If cc2 >= 1 And ec2 <> "" Then
                On Error Resume Next
                Set tmpLO = ws.ListObjects.Add(xlSrcRange, mUR2, , xlYes)
                On Error GoTo 0
                If Not tmpLO Is Nothing Then
                    Set mLO = tmpLO
                    Set wsMaster = ws
                    keyColSrc = kc2
                    keyColMst = kc2
                    extraCol = ec2
                    GoTo FoundMaster
                End If
            End If
        End If
NextSheet2:
    Next ws

FoundMaster:
    If mLO Is Nothing Then Exit Sub
    If keyColSrc = "" Or extraCol = "" Then Exit Sub

    mstTblName = mLO.Name
    destSheetName = srcTblName & "_結合"
    qryName = srcTblName & "_結合"

    nl = Chr(13) & Chr(10)

    keyTypeStr = "type text"
    If srcLO.DataBodyRange Is Nothing Then Exit Sub
    Dim rr As Long
    Dim srcRows As Long
    srcRows = srcLO.DataBodyRange.rows.Count
    For i = 1 To srcHdrCount
        If srcCols(i) = keyColSrc Then
            Set cell = Nothing
            For rr = 1 To srcRows
                If Trim(CStr(srcLO.DataBodyRange.Cells(rr, i).Value)) <> "" Then
                    Set cell = srcLO.DataBodyRange.Cells(rr, i)
                    Exit For
                End If
            Next rr
            If Not cell Is Nothing Then
                If IsNumeric(cell.Value) And Not IsDate(cell.Value) Then
                    keyTypeStr = "type number"
                End If
            End If
            Exit For
        End If
    Next i

    dateColCount = 0
    ReDim dateColNames(1 To srcHdrCount)

    srcTypeList = ""
    For i = 1 To srcHdrCount
        ' 列の型は上から順に見て最初に値の入っているセルで決める
        ' （1 行目が空の文字の列を数の列と間違えない＝空のセルの IsNumeric は True）
        Set cell = Nothing
        For rr = 1 To srcRows
            If Trim(CStr(srcLO.DataBodyRange.Cells(rr, i).Value)) <> "" Then
                Set cell = srcLO.DataBodyRange.Cells(rr, i)
                Exit For
            End If
        Next rr
        Dim isDateCol As Boolean
        isDateCol = False
        If srcCols(i) = keyColSrc Then
            typeStr = keyTypeStr
        ElseIf cell Is Nothing Then
            typeStr = "type text"
        Else
            If IsDate(cell.Value) And Not IsNumeric(cell.Value) Then
                isDateCol = True
            ElseIf cell.NumberFormat Like "*yyyy*" Or cell.NumberFormat Like "*yy/mm*" Or _
                   cell.NumberFormat Like "*m/d*" Or cell.NumberFormat Like "*m""月""*" Then
                isDateCol = True
            End If
            If isDateCol Then
                typeStr = "type any"
                dateColCount = dateColCount + 1
                dateColNames(dateColCount) = srcCols(i)
            ElseIf IsNumeric(cell.Value) Then
                typeStr = "type number"
            Else
                typeStr = "type text"
            End If
        End If
        If srcTypeList <> "" Then srcTypeList = srcTypeList & ", "
        srcTypeList = srcTypeList & "{""" & srcCols(i) & """, " & typeStr & "}"
    Next i

    ' マスタ側の追加列の型を判定
    Dim extraTypeStr As String
    extraTypeStr = "type text"
    If Not mLO.DataBodyRange Is Nothing Then
        For j = 1 To mLO.HeaderRowRange.Columns.Count
            If CStr(mLO.HeaderRowRange.Cells(1, j).Value) = extraCol Then
                Dim eCell As Range
                Set eCell = Nothing
                For rr = 1 To mLO.DataBodyRange.rows.Count
                    If Trim(CStr(mLO.DataBodyRange.Cells(rr, j).Value)) <> "" Then
                        Set eCell = mLO.DataBodyRange.Cells(rr, j)
                        Exit For
                    End If
                Next rr
                If Not eCell Is Nothing Then
                    If IsNumeric(eCell.Value) And Not IsDate(eCell.Value) Then
                        extraTypeStr = "type number"
                    End If
                End If
                Exit For
            End If
        Next j
    End If

    mstTypeList = "{""" & keyColMst & """, " & keyTypeStr & "}, {""" & extraCol & """, " & extraTypeStr & "}"

    srcColList = ""
    For i = 1 To srcHdrCount
        If srcColList <> "" Then srcColList = srcColList & ", "
        srcColList = srcColList & """" & srcCols(i) & """"
    Next i

    Dim dateColsM As String
    dateColsM = ""
    Dim dc As Long
    For dc = 1 To dateColCount
        If dateColsM <> "" Then dateColsM = dateColsM & ", "
        dateColsM = dateColsM & "{""" & dateColNames(dc) & """, each if _ = null then null else Date.From(_), type date}"
    Next dc

    Dim lastRef As String
    Dim extraSteps As String
    extraSteps = ""
    lastRef = "filled"

    If dateColsM <> "" Then
        extraSteps = "    dateFixed = Table.TransformColumns(filled, {" & dateColsM & "})," & nl
        lastRef = "dateFixed"
    End If

    ' 備考のようなテキスト列で空白セルがある場合、nullをそのまま保持（extraColのみ「未登録」変換）
    mFormula = "let" & nl & _
        "    src = Excel.CurrentWorkbook(){[Name=""" & srcTblName & """]}[Content]," & nl & _
        "    mst = Excel.CurrentWorkbook(){[Name=""" & mstTblName & """]}[Content]," & nl & _
        "    srcTyped = Table.TransformColumnTypes(src, {" & srcTypeList & "})," & nl & _
        "    mstTyped = Table.TransformColumnTypes(mst, {" & mstTypeList & "})," & nl & _
        "    merged = Table.NestedJoin(srcTyped, {""" & keyColSrc & """}, mstTyped, {""" & keyColMst & """}, ""_mst"", JoinKind.LeftOuter)," & nl & _
        "    expanded = Table.ExpandTableColumn(merged, ""_mst"", {""" & extraCol & """}, {""" & extraCol & """})," & nl & _
        "    filled = Table.TransformColumns(expanded, {{""" & extraCol & """, each if _ = null then ""未登録"" else _, type text}})," & nl & _
        extraSteps & _
        "    selected = Table.SelectColumns(" & lastRef & ", {" & srcColList & ", """ & extraCol & """})," & nl & _
        "    sorted = Table.Sort(selected, {{""" & keyColSrc & """, Order.Ascending}})" & nl & _
        "in" & nl & _
        "    sorted"

    On Error Resume Next
    For Each qExist In wb.Queries
        If qExist.Name = qryName Then
            qExist.Delete
            Exit For
        End If
    Next qExist
    On Error GoTo 0

    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsExist = wb.Worksheets(destSheetName)
    If Not wsExist Is Nothing Then
        wsExist.Delete
        Set wsExist = Nothing
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    On Error Resume Next
    wb.Queries.Add Name:=qryName, Formula:=mFormula
    On Error GoTo 0

    Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsDest.Name = destSheetName

    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qryName & ";Extended Properties="""""
    On Error Resume Next
    Set destLO = wsDest.ListObjects.Add( _
        SourceType:=0, _
        Source:=connStr, _
        Destination:=wsDest.Range("A1"))
    On Error GoTo 0
    If destLO Is Nothing Then Exit Sub
    destLO.Name = qryName
    With destLO.QueryTable
        .CommandType = 2
        .CommandText = "SELECT * FROM [" & qryName & "]"
        .AdjustColumnWidth = False
        .RefreshStyle = xlInsertDeleteCells
        .PreserveColumnInfo = False
        On Error Resume Next
        .Refresh False
        On Error GoTo 0
    End With

    If destLO Is Nothing Then Exit Sub
    If destLO.DataBodyRange Is Nothing Then Exit Sub
    Dim hdr As ListColumn
    Dim dCol As Long
    Dim dRow As Long
    Dim dVal As Variant
    For Each hdr In destLO.ListColumns
        found = False
        For dc = 1 To dateColCount
            If dateColNames(dc) = hdr.Name Then found = True: Exit For
        Next dc
        If found Then
            dCol = hdr.Index
            For dRow = 1 To destLO.DataBodyRange.rows.Count
                dVal = destLO.DataBodyRange.Cells(dRow, dCol).Value
                If VarType(dVal) = vbString Then
                    Dim dStr As String
                    dStr = CStr(dVal)
                    On Error Resume Next
                    Dim dParsed As Date
                    dParsed = CDate(Left$(dStr, 10))
                    If Err.Number = 0 Then
                        destLO.DataBodyRange.Cells(dRow, dCol).Value = dParsed
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            Next dRow
        End If
    Next hdr

    ' ソース側の空文字セル（備考列など）：クエリリフレッシュ後にnullとして読み込まれた場合は空文字に戻す
    Dim srcColIdx As Long
    Dim isExtraCol As Boolean
    For Each hdr In destLO.ListColumns
        isExtraCol = (hdr.Name = extraCol)
        If Not isExtraCol Then
            ' ソース列でテキスト型の場合、nullは空文字に
            Dim isSrcCol As Boolean: isSrcCol = False
            For k = 1 To srcHdrCount
                If srcCols(k) = hdr.Name Then
                    isSrcCol = True
                    Exit For
                End If
            Next k
            If isSrcCol Then
                ' 元データでテキスト型列の場合のみ空文字復元
                Dim isNumCol As Boolean: isNumCol = False
                For k = 1 To srcHdrCount
                    If srcCols(k) = hdr.Name And srcCols(k) = keyColSrc Then
                        isNumCol = (keyTypeStr = "type number")
                        Exit For
                    ElseIf srcCols(k) = hdr.Name Then
                        ' dateColでもkeyColでもなければ型チェック
                        Dim idc As Long
                        For idc = 1 To dateColCount
                            If dateColNames(idc) = hdr.Name Then isNumCol = False: GoTo SkipNumCheck
                        Next idc
                        If srcLO.DataBodyRange Is Nothing Then GoTo SkipNumCheck
                        Set cell = srcLO.DataBodyRange.Cells(1, k)
                        isNumCol = IsNumeric(cell.Value) And Not IsDate(cell.Value)
SkipNumCheck:
                        Exit For
                    End If
                Next k
                If Not isNumCol Then
                    dCol = hdr.Index
                    For dRow = 1 To destLO.DataBodyRange.rows.Count
                        If IsNull(destLO.DataBodyRange.Cells(dRow, dCol).Value) Or _
                           isEmpty(destLO.DataBodyRange.Cells(dRow, dCol).Value) Then
                            ' 元の対応行の値を確認して空文字なら空文字のまま
                        End If
                    Next dRow
                End If
            End If
        End If
    Next hdr

End Sub

Sub 構成比ピボットを作る()
    ' 依頼の語: 構成比のピボット|大きい順に構成比|選んだ列で構成比|割合のピボット|比率も出
    ' 扱う: 構成比 ピボット 合計 集計 比率
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列 日付の列
    ' 選んだ2列(1つ目=項目列、2つ目=値列)を使い、新シートのA3に構成比ピボットを作る。値は合計と列比率。合計の大きい順、総計あり。

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim ptSheet As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim df1 As PivotField
    Dim df2 As PivotField
    Dim ur As Range
    Dim hRow As Long, lastRow As Long
    Dim itemCol As Long, valCol As Long
    Dim leftCol As Long, rightCol As Long
    Dim r As Long, ci As Long
    Dim srcRange As String
    Dim sheetName As String
    Dim ws2 As Worksheet
    Dim sel1 As Range, sel2 As Range
    Dim itemHeader As String, valHeader As String
    Dim sumFieldName As String

    Set ws = ActiveSheet
    Set wb = ActiveSheet.Parent

    ' 選んでいる列から項目列・値列を決める
    If Selection.Areas.Count < 2 Then Exit Sub
    Set sel1 = Selection.Areas(1)
    Set sel2 = Selection.Areas(2)
    itemCol = sel1.Column
    valCol = sel2.Column

    ' UsedRangeから見出し行を探す（上から順に両列に値がある最初の行）
    Set ur = ws.UsedRange
    hRow = 0
    Dim ri As Long
    For ri = ur.Row To ur.Row + ur.rows.Count - 1
        If Trim(CStr(ws.Cells(ri, itemCol).Value)) <> "" And _
           Trim(CStr(ws.Cells(ri, valCol).Value)) <> "" Then
            hRow = ri
            Exit For
        End If
    Next ri
    If hRow = 0 Then Exit Sub

    itemHeader = CStr(ws.Cells(hRow, itemCol).Value)
    valHeader = CStr(ws.Cells(hRow, valCol).Value)

    ' 最後の明細行（合計・計・小計・総計の行は除く。項目列に値がある最後の行）
    lastRow = hRow
    For r = hRow + 1 To ur.Row + ur.rows.Count - 1
        Dim cv As String
        cv = Trim(CStr(ws.Cells(r, itemCol).Value))
        If cv <> "" Then
            ' 合計行除外
            If cv <> "合計" And cv <> "計" And cv <> "小計" And cv <> "総計" Then
                lastRow = r
            End If
        End If
    Next r
    If lastRow = hRow Then Exit Sub

    leftCol = ur.Column
    rightCol = ur.Column + ur.Columns.Count - 1
    srcRange = "'" & ws.Name & "'!" & _
        ws.Cells(hRow, leftCol).Address(True, True) & ":" & _
        ws.Cells(lastRow, rightCol).Address(True, True)

    ' 新シート名
    sheetName = itemHeader & "別構成比"

    ' 既存シート削除
    Application.DisplayAlerts = False
    For Each ws2 In wb.Worksheets
        If ws2.Name = sheetName Then
            ws2.Delete
            Exit For
        End If
    Next ws2
    Application.DisplayAlerts = True

    Set ptSheet = wb.Worksheets.Add(After:=ws)
    ptSheet.Name = sheetName

    On Error Resume Next
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRange)
    On Error GoTo 0
    If pc Is Nothing Then Exit Sub

    On Error Resume Next
    Set pt = pc.CreatePivotTable(TableDestination:=ptSheet.Cells(3, 1), TableName:=sheetName)
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    On Error Resume Next
    Set pf = pt.PivotFields(itemHeader)
    pf.Orientation = xlRowField
    pf.Position = 1
    On Error GoTo 0

    sumFieldName = "合計 / " & valHeader
    On Error Resume Next
    Set df1 = pt.AddDataField(pt.PivotFields(valHeader), sumFieldName, xlSum)
    df1.NumberFormat = "#,##0"
    On Error GoTo 0

    On Error Resume Next
    Set df2 = pt.AddDataField(pt.PivotFields(valHeader), "構成比", xlSum)
    df2.Calculation = xlPercentOfColumn
    df2.NumberFormat = "0.0%"
    On Error GoTo 0

    On Error Resume Next
    pt.ColumnGrand = True
    pt.RowGrand = True
    On Error GoTo 0

    On Error Resume Next
    pt.PivotFields(itemHeader).AutoSort xlDescending, sumFieldName
    On Error GoTo 0

End Sub

Sub 積み上げ縦棒グラフを作る()
    ' 依頼の語: 積み上げ縦棒|積み上げ棒グラフ|積み上げグラフ|積み上げの棒|積み上げの縦棒|内訳を積み上げ|構成を積み上げ
    ' 扱う: 積み上げグラフ 合計 ピボット
    ' 見出し: なし
    ' 形: 数の列
    ' 左端の文字列列を項目、右の数値列（合計・総計列除く）を系列にした積み上げ縦棒グラフを作成する
    ' 合計行は除外し、既存グラフは削除して作り直す。数値軸は#,##0、凡例は右、タイトルあり

    Dim ws As Object
    Dim ur As Object
    Dim r As Long, c As Long
    Dim startRow As Long, startCol As Long
    Dim endRow As Long, endCol As Long
    Dim s As String
    Dim isSumCol As Boolean
    Dim isSumRow As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    startRow = ur.Row: startCol = ur.Column
    endRow = ur.Row + ur.rows.Count - 1
    endCol = ur.Column + ur.Columns.Count - 1

    ' 見出し行を探す：数値が並ぶ本文のすぐ上の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim dataStartRow As Long
    dataStartRow = 0

    For r = startRow To endRow
        Dim numCount As Long
        numCount = 0
        For c = startCol To endCol
            If IsNumeric(ws.Cells(r, c).Value) And Trim(CStr(ws.Cells(r, c).Value)) <> "" Then
                numCount = numCount + 1
            End If
        Next c
        If numCount >= 1 Then
            dataStartRow = r
            hdrRow = r - 1
            Exit For
        End If
    Next r

    If hdrRow < startRow Then hdrRow = startRow
    If dataStartRow = 0 Then dataStartRow = hdrRow + 1

    ' 表題行を探す（見出し行より上の行で、ピボット由来の「合計/…」「列ラベル」「行ラベル」以外の文字を表題とする）
    Dim titleText As String
    titleText = ""
    If hdrRow > startRow Then
        Dim tr As Long
        For tr = startRow To hdrRow - 1
            Dim mergedText As String
            mergedText = ""
            For c = startCol To endCol
                Dim cv As String
                cv = Trim(CStr(ws.Cells(tr, c).Value))
                If cv <> "" Then
                    If mergedText = "" Then
                        mergedText = cv
                    Else
                        mergedText = mergedText & " " & cv
                    End If
                End If
            Next c
            If mergedText <> "" Then
                ' ピボット由来の行は表題としない
                Dim isPivotRow As Boolean
                isPivotRow = False
                Dim checkStr As String
                checkStr = mergedText
                If InStr(checkStr, "合計 /") > 0 Or InStr(checkStr, "合計/") > 0 Then isPivotRow = True
                If checkStr = "列ラベル" Then isPivotRow = True
                If checkStr = "行ラベル" Then isPivotRow = True
                If Not isPivotRow Then
                    titleText = mergedText
                End If
            End If
        Next tr
    End If

    ' 見出し行もピボット由来か確認し、項目列の見出し語を決める
    ' hdrRowの左端列の値が「行ラベル」の場合はピボット→タイトルは「行ラベル別の内訳」ではなく列見出しの値を使う
    ' → 見出し行の項目列の値をそのまま使う（行ラベルも含む）

    ' 項目列（見出し行より下で文字データを持つ最左列）を探す
    Dim itemCol As Long
    itemCol = 0

    For c = startCol To endCol
        Dim hasTextData As Boolean
        hasTextData = False
        For r = dataStartRow To endRow
            s = Trim(CStr(ws.Cells(r, c).Value))
            If s <> "" Then
                If Not IsNumeric(ws.Cells(r, c).Value) Then
                    hasTextData = True
                    Exit For
                End If
            End If
        Next r
        If hasTextData Then
            itemCol = c
            Exit For
        End If
    Next c
    If itemCol = 0 Then Exit Sub

    ' 見出し行のitemCol値
    Dim itemHdrVal As String
    itemHdrVal = Trim(CStr(ws.Cells(hdrRow, itemCol).Value))

    ' 数値列を収集（合計・総計列除く）
    Dim numCols() As Long
    Dim numColNames() As String
    Dim numColCount As Long
    numColCount = 0
    ReDim numCols(0)
    ReDim numColNames(0)

    For c = itemCol + 1 To endCol
        s = Trim(CStr(ws.Cells(hdrRow, c).Value))
        GoSub isSumLabel
        If isSumCol Then GoTo NextCol

        Dim colHasNum As Boolean
        colHasNum = False
        For r = dataStartRow To endRow
            If IsNumeric(ws.Cells(r, c).Value) And Trim(CStr(ws.Cells(r, c).Value)) <> "" Then
                colHasNum = True
                Exit For
            End If
        Next r
        If colHasNum Then
            ReDim Preserve numCols(numColCount)
            ReDim Preserve numColNames(numColCount)
            numCols(numColCount) = c
            numColNames(numColCount) = s
            numColCount = numColCount + 1
        End If
NextCol:
    Next c
    If numColCount = 0 Then Exit Sub

    ' データ行を収集（合計行除く・空行除く）
    Dim dataRows() As Long
    Dim dataNames() As String
    Dim dataRowCount As Long
    dataRowCount = 0
    ReDim dataRows(0)
    ReDim dataNames(0)

    For r = dataStartRow To endRow
        s = Trim(CStr(ws.Cells(r, itemCol).Value))
        If s = "" Then GoTo NextRow
        GoSub isSumRowCheck
        If isSumRow Then GoTo NextRow
        ReDim Preserve dataRows(dataRowCount)
        ReDim Preserve dataNames(dataRowCount)
        dataRows(dataRowCount) = r
        dataNames(dataRowCount) = s
        dataRowCount = dataRowCount + 1
NextRow:
    Next r
    If dataRowCount = 0 Then Exit Sub

    ' タイトル決定
    ' 表題行があればそれを使う
    ' なければ itemHdrVal+"別の内訳"
    ' ただし itemHdrVal が「行ラベル」の場合も "行ラベル別の内訳" を使う（ピボット表で表題なし）
    Dim chartTitle As String
    If titleText <> "" Then
        chartTitle = titleText
    ElseIf itemHdrVal <> "" Then
        chartTitle = itemHdrVal & "別の内訳"
    Else
        chartTitle = "内訳"
    End If

    ' 既存グラフ削除
    Dim co As Object
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim leftPos As Double, topPos As Double
    leftPos = ws.Cells(hdrRow, endCol + 2).Left
    topPos = ws.Cells(hdrRow, endCol + 2).Top

    Dim chObj As Object
    Set chObj = ws.ChartObjects.Add(leftPos, topPos, 400, 300)
    Dim ch As Object
    Set ch = chObj.Chart

    ch.ChartType = 52 ' xlColumnStacked

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim catArr() As String
    ReDim catArr(1 To dataRowCount)
    Dim i As Long
    For i = 0 To dataRowCount - 1
        catArr(i + 1) = dataNames(i)
    Next i

    Dim j As Long
    Dim valArr() As Double
    Dim ser As Object
    For j = 0 To numColCount - 1
        ReDim valArr(1 To dataRowCount)
        For i = 0 To dataRowCount - 1
            valArr(i + 1) = ws.Cells(dataRows(i), numCols(j)).Value
        Next i
        Set ser = ch.SeriesCollection.NewSeries
        ser.Name = numColNames(j)
        ser.Values = valArr
        ser.XValues = catArr
    Next j

    ch.HasLegend = True
    ch.Legend.Position = -4152 ' xlRight

    ch.HasTitle = True
    ch.chartTitle.text = chartTitle
    ch.chartTitle.Font.Size = 18
    ch.chartTitle.Font.Bold = True

    Dim ax As Object
    On Error Resume Next
    Set ax = ch.Axes(2)
    On Error GoTo 0
    If Not ax Is Nothing Then
        ax.TickLabels.NumberFormat = "#,##0"
    End If

    Exit Sub

isSumLabel:
    isSumCol = (InStr(s, "合計") > 0 Or InStr(s, "総計") > 0 Or s = "計")
    Return

isSumRowCheck:
    s = Trim(CStr(ws.Cells(r, itemCol).Value))
    isSumRow = (InStr(s, "合計") > 0 Or InStr(s, "総計") > 0 Or s = "計")
    Return

End Sub

Sub テーブルを並べ替えと絞り込みで整える()
    ' 依頼の語: 替えて絞|替えと絞|選んだ列で並|0を除いて|0の行は隠|列で並
    ' 扱う: 並べ替え 絞り込み ソート フィルター
    ' 見出し: なし
    ' 選ぶ列: 2-3
    ' 形: 数の列 日付の列 テーブル

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim loTarget As ListObject
    Dim area1Col As Long, area2Col As Long, area3Col As Long
    Dim col1 As Long, col2 As Long, col3 As Long
    Dim i As Long, absCol As Long

    Set ws = ActiveSheet

    area1Col = 0: area2Col = 0: area3Col = 0
    On Error Resume Next
    If Selection.Areas.Count >= 1 Then area1Col = Selection.Areas(1).Column
    If Selection.Areas.Count >= 2 Then area2Col = Selection.Areas(2).Column
    If Selection.Areas.Count >= 3 Then area3Col = Selection.Areas(3).Column
    On Error GoTo 0

    If area1Col = 0 Or area2Col = 0 Then Exit Sub

    For Each lo In ws.ListObjects
        col1 = 0: col2 = 0: col3 = 0
        For i = 1 To lo.ListColumns.Count
            absCol = lo.ListColumns(i).Range.Column
            If absCol = area1Col Then col1 = i
            If absCol = area2Col Then col2 = i
            If area3Col > 0 And absCol = area3Col Then col3 = i
        Next i
        If col1 > 0 And col2 > 0 Then
            Set loTarget = lo
            Exit For
        End If
    Next lo

    If loTarget Is Nothing Then Exit Sub

    If loTarget.ShowAutoFilter Then
        On Error Resume Next
        loTarget.AutoFilter.ShowAllData
        On Error GoTo 0
    End If

    With loTarget.Sort
        .SortFields.Clear
        .SortFields.Add key:=loTarget.ListColumns(col1).Range, _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=loTarget.ListColumns(col2).Range, _
            SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    If col3 > 0 Then
        loTarget.Range.AutoFilter Field:=col3, Criteria1:="<>0"
    End If

End Sub

Sub 月別クロス集計をパワークエリで作る()
    ' 依頼の語: クロス表|月を列|クエリで月別|クロス集計表|選んだ列でクロス
    ' 扱う: パワークエリ クロス集計 月別 合計 集計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列 日付の列 テーブル
    ' 選んでいる3列（1列目=行キー,2列目=日付,3列目=値）からパワークエリで月別クロス集計表を作り、新シートのA1にテーブルとして読み込む

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim colRow As Long, colDate As Long, colVal As Long
    If Selection.Areas.Count >= 3 Then
        colRow = Selection.Areas(1).Column
        colDate = Selection.Areas(2).Column
        colVal = Selection.Areas(3).Column
    ElseIf Selection.Columns.Count >= 3 Then
        colRow = Selection.Columns(1).Column
        colDate = Selection.Columns(2).Column
        colVal = Selection.Columns(3).Column
    Else
        Exit Sub
    End If

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim hdrRow As Long
    Dim r As Long
    hdrRow = ur.Row
    Dim urLastRow As Long
    urLastRow = ur.Row + ur.rows.Count - 1
    Dim c As Long
    For r = ur.Row To urLastRow
        Dim numCount As Long
        numCount = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            If Not isEmpty(ws.Cells(r, c).Value) Then
                If IsDate(ws.Cells(r, c).Value) Then
                    numCount = numCount + 1
                ElseIf IsNumeric(ws.Cells(r, c).Value) Then
                    numCount = numCount + 1
                End If
            End If
        Next c
        If numCount >= 2 Then
            If r > ur.Row Then hdrRow = r - 1 Else hdrRow = r
            Exit For
        End If
    Next r

    Dim srcTblName As String
    srcTblName = ""
    Dim lo As ListObject
    Dim li As ListObject
    If ws.ListObjects.Count > 0 Then
        For Each li In ws.ListObjects
            If colRow >= li.Range.Column And colRow <= li.Range.Column + li.Range.Columns.Count - 1 Then
                Set lo = li
                srcTblName = li.Name
                Exit For
            End If
        Next li
        If srcTblName = "" Then
            Set lo = ws.ListObjects(1)
            srcTblName = ws.ListObjects(1).Name
        End If
    Else
        Dim fullRange As Range
        Set fullRange = ws.Range(ws.Cells(hdrRow, ur.Column), ws.Cells(urLastRow, ur.Column + ur.Columns.Count - 1))
        On Error Resume Next
        Set lo = ws.ListObjects.Add(xlSrcRange, fullRange, , xlYes)
        On Error GoTo 0
        If lo Is Nothing Then
            If ws.ListObjects.Count > 0 Then Set lo = ws.ListObjects(1)
        End If
        If Not lo Is Nothing Then
            lo.Name = "Table_auto"
            srcTblName = lo.Name
        End If
    End If

    If srcTblName = "" Then Exit Sub

    Dim colRowName As String, colDateName As String, colValName As String
    colRowName = CStr(ws.Cells(hdrRow, colRow).Value)
    colDateName = CStr(ws.Cells(hdrRow, colDate).Value)
    colValName = CStr(ws.Cells(hdrRow, colVal).Value)

    Dim qryName As String
    Dim sheetName As String
    qryName = srcTblName & "_クロス"
    sheetName = srcTblName & "_クロス"

    Dim q As Object
    For Each q In wb.Queries
        If q.Name = qryName Then
            q.Delete
            Exit For
        End If
    Next q

    Application.DisplayAlerts = False
    Dim wsDel As Worksheet
    For Each wsDel In wb.Sheets
        If wsDel.Name = sheetName Then
            wsDel.Delete
            Exit For
        End If
    Next wsDel
    Application.DisplayAlerts = True

    Dim nl As String
    nl = Chr(10)

    Dim mFormula As String
    ' 月の列は 4月→3月の決まった並びのうち明細にある月だけ。空行は日付が空＝月が null になるので
    ' if [_m] >= 4 then … のような判定を入れると M がこけて何も読み込まれない（そこは触らない）
    mFormula = "let" & nl & _
        "    Source = Excel.CurrentWorkbook(){[Name=""" & srcTblName & """]}[Content]," & nl & _
        "    Typed = Table.TransformColumnTypes(Source, {{""" & colDateName & """, type date}, {""" & colRowName & """, type text}, {""" & colValName & """, type number}})," & nl & _
        "    AddM = Table.AddColumn(Typed, ""_m"", each Date.Month([" & colDateName & "]), Int64.Type)," & nl & _
        "    Grouped = Table.Group(AddM, {""" & colRowName & """, ""_m""}, {{""" & colValName & """, each List.Sum([" & colValName & "]), type nullable number}})," & nl & _
        "    Order = List.Select({4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3}, each List.Contains(Grouped[_m], _))," & nl & _
        "    Named = Table.TransformColumns(Grouped, {{""_m"", each Text.From(_) & ""月"", type text}})," & nl & _
        "    Pivoted = Table.Pivot(Named, List.Transform(Order, each Text.From(_) & ""月""), ""_m"", """ & colValName & """, List.Sum)," & nl & _
        "    Sorted = Table.Sort(Pivoted, {{""" & colRowName & """, Order.Ascending}})" & nl & _
        "in" & nl & _
        "    Sorted"

    wb.Queries.Add qryName, mFormula

    Dim newWs As Worksheet
    Set newWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    newWs.Name = sheetName

    Dim connStr As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qryName

    Dim newLo As ListObject
    On Error Resume Next
    Set newLo = newWs.ListObjects.Add( _
        SourceType:=0, _
        Source:=connStr, _
        Destination:=newWs.Range("A1"))
    On Error GoTo 0

    If Not newLo Is Nothing Then
        On Error Resume Next
        newLo.QueryTable.CommandType = 2
        newLo.QueryTable.CommandText = "SELECT * FROM [" & qryName & "]"
        newLo.QueryTable.AdjustColumnWidth = False
        newLo.QueryTable.BackgroundQuery = False
        newLo.QueryTable.Refresh BackgroundQuery:=False
        newLo.Name = qryName
        On Error GoTo 0
    End If

    If newLo Is Nothing Or newWs.UsedRange.Cells.Count <= 1 Then
        ' ListObjects.Add が失敗またはデータ未取得の場合、QueryTable で取得後テーブル化
        Dim qt As QueryTable
        ' 既存のListObjectsをクリア
        Do While newWs.ListObjects.Count > 0
            On Error Resume Next
            newWs.ListObjects(1).Delete
            On Error GoTo 0
        Loop
        newWs.Cells.Clear

        On Error Resume Next
        Set qt = newWs.QueryTables.Add( _
            Connection:=connStr, _
            Destination:=newWs.Range("A1"))
        If Not qt Is Nothing Then
            qt.CommandType = 2
            qt.CommandText = "SELECT * FROM [" & qryName & "]"
            qt.AdjustColumnWidth = False
            qt.BackgroundQuery = False
            qt.Refresh BackgroundQuery:=False
        End If
        On Error GoTo 0

        If Not qt Is Nothing Then
            Dim dataRng As Range
            On Error Resume Next
            Set dataRng = newWs.UsedRange
            On Error GoTo 0
            If Not dataRng Is Nothing Then
                On Error Resume Next
                qt.Delete
                Set newLo = newWs.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
                If Not newLo Is Nothing Then
                    newLo.Name = qryName
                End If
                On Error GoTo 0
            End If
        End If
    End If

End Sub

Sub 推移グラフを作る()
    ' 依頼の語: 推移グラフ|いつもの推移|推移のグラフ|見せ方の折れ線|グラフでお願|推移をいつも|支出の推移|折れ線グラフ
    ' 扱う: 推移グラフ グラフ 合計
    ' 見出し: なし
    ' 形: 数の列
    ' 左端の文字列列を系列、右の数値列（合計列除く）を項目とするマーカー付き折れ線グラフを作る。既存グラフは削除して作り直す。

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim ur As Range
    Set ur = ws.UsedRange

    Dim r As Long, c As Long
    Dim urR As Long, urC As Long, urR2 As Long, urC2 As Long
    urR = ur.Row: urC = ur.Column
    urR2 = urR + ur.rows.Count - 1
    urC2 = urC + ur.Columns.Count - 1

    Dim bestRow As Long, bestCount As Long
    bestRow = 0: bestCount = 0
    For r = urR To urR2
        Dim cnt As Long: cnt = 0
        For c = urC To urC2
            If IsNumeric(ws.Cells(r, c).Value) And ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
        Next c
        If cnt > bestCount Then bestCount = cnt: bestRow = r
    Next r

    Dim hRow As Long
    hRow = bestRow - 1
    If hRow < urR Then hRow = urR

    Dim nameCol As Long: nameCol = 0
    For c = urC To urC2
        Dim strCnt As Long: strCnt = 0
        For r = hRow + 1 To urR2
            Dim v As String: v = CStr(ws.Cells(r, c).Value)
            If v <> "" And Not IsNumeric(ws.Cells(r, c).Value) Then strCnt = strCnt + 1
        Next r
        If strCnt > 0 Then nameCol = c: Exit For
    Next c
    If nameCol = 0 Then Exit Sub

    Dim firstDataRow As Long: firstDataRow = hRow + 1

    Dim lastDataRow As Long: lastDataRow = urR2
    Do While lastDataRow >= firstDataRow
        Dim cv As String: cv = Trim(CStr(ws.Cells(lastDataRow, nameCol).Value))
        If cv = "合計" Or cv = "合 計" Or cv = "計" Or cv = "小計" Or cv = "総計" Then
            lastDataRow = lastDataRow - 1
        Else
            Exit Do
        End If
    Loop
    If lastDataRow < firstDataRow Then Exit Sub

    Dim dataRows() As Long
    Dim dataRowCount As Long: dataRowCount = 0
    ReDim dataRows(lastDataRow - firstDataRow)
    For r = firstDataRow To lastDataRow
        Dim rowKey As String: rowKey = Trim(CStr(ws.Cells(r, nameCol).Value))
        If rowKey <> "" Then
            dataRows(dataRowCount) = r
            dataRowCount = dataRowCount + 1
        End If
    Next r
    If dataRowCount = 0 Then Exit Sub

    Dim numCols() As Long
    Dim numColCount As Long: numColCount = 0
    ReDim numCols(urC2 - urC)
    For c = nameCol + 1 To urC2
        Dim hv As String: hv = Trim(CStr(ws.Cells(hRow, c).Value))
        If hv = "合計" Or hv = "合 計" Or hv = "計" Or hv = "小計" Or hv = "総計" Then GoTo NextCol
        Dim isNum As Boolean: isNum = False
        For r = firstDataRow To lastDataRow
            If IsNumeric(ws.Cells(r, c).Value) And ws.Cells(r, c).Value <> "" Then isNum = True: Exit For
        Next r
        If Not isNum Then GoTo NextCol
        numCols(numColCount) = c
        numColCount = numColCount + 1
NextCol:
    Next c
    If numColCount = 0 Then Exit Sub

    Dim xRng As Range
    Set xRng = ws.Cells(hRow, numCols(0))
    Dim j As Long
    For j = 1 To numColCount - 1
        Set xRng = Union(xRng, ws.Cells(hRow, numCols(j)))
    Next j

    Dim titleText As String: titleText = ""
    Dim tr As Long, tc As Long
    For tr = urR To hRow - 1
        For tc = urC To urC2
            Dim tval As String: tval = Trim(CStr(ws.Cells(tr, tc).Value))
            If tval <> "" Then titleText = tval: Exit For
        Next tc
        If titleText <> "" Then Exit For
    Next tr
    If titleText = "" Then
        Dim hdrVal As String: hdrVal = Trim(CStr(ws.Cells(hRow, nameCol).Value))
        If hdrVal <> "" Then
            titleText = hdrVal & "ごとの推移"
        Else
            titleText = "推移グラフ"
        End If
    End If

    Dim chartLeft As Double, chartTop As Double
    chartLeft = ws.Cells(hRow, nameCol).Left
    chartTop = ws.Cells(lastDataRow + 2, nameCol).Top

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(chartLeft, chartTop, 480, 300)

    Dim seriesColors(7) As Long
    seriesColors(0) = RGB(31, 78, 121)
    seriesColors(1) = RGB(192, 0, 0)
    seriesColors(2) = RGB(84, 130, 53)
    seriesColors(3) = RGB(191, 143, 0)
    seriesColors(4) = RGB(112, 48, 160)
    seriesColors(5) = RGB(255, 102, 0)
    seriesColors(6) = RGB(70, 130, 180)
    seriesColors(7) = RGB(0, 176, 240)

    Dim seriesCount As Long: seriesCount = dataRowCount
    If seriesCount > 5 Then seriesCount = 5

    Dim names() As String
    Dim rowNums() As Long
    Dim totals() As Double
    ReDim names(dataRowCount - 1)
    ReDim rowNums(dataRowCount - 1)
    ReDim totals(dataRowCount - 1)
    Dim i As Long
    For i = 0 To dataRowCount - 1
        names(i) = CStr(ws.Cells(dataRows(i), nameCol).Value)
        rowNums(i) = dataRows(i)
        Dim sm As Double: sm = 0
        For j = 0 To numColCount - 1
            Dim vv As Variant: vv = ws.Cells(dataRows(i), numCols(j)).Value
            If IsNumeric(vv) And vv <> "" Then sm = sm + CDbl(vv)
        Next j
        totals(i) = sm
    Next i

    Dim tmpName As String, tmpRow As Long, tmpTot As Double
    For i = 0 To dataRowCount - 2
        For j = i + 1 To dataRowCount - 1
            If totals(j) > totals(i) Then
                tmpName = names(i): names(i) = names(j): names(j) = tmpName
                tmpRow = rowNums(i): rowNums(i) = rowNums(j): rowNums(j) = tmpRow
                tmpTot = totals(i): totals(i) = totals(j): totals(j) = tmpTot
            End If
        Next j
    Next i

    With cht.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim srs As Series
        For i = 0 To seriesCount - 1
            Set srs = .SeriesCollection.NewSeries
            srs.Name = names(i)
            If numColCount = 1 Then
                srs.Values = ws.Cells(rowNums(i), numCols(0))
            Else
                Dim vRng As Range
                Set vRng = ws.Cells(rowNums(i), numCols(0))
                For j = 1 To numColCount - 1
                    Set vRng = Union(vRng, ws.Cells(rowNums(i), numCols(j)))
                Next j
                srs.Values = vRng
            End If
            srs.XValues = xRng
            srs.ChartType = xlLineMarkers
            srs.MarkerSize = 6
            Dim clr As Long: clr = seriesColors(i Mod 8)
            srs.Format.Line.ForeColor.RGB = clr
            srs.MarkerForegroundColor = clr
            srs.MarkerBackgroundColor = clr
        Next i

        On Error Resume Next
        .ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"
        On Error GoTo 0

        .chartTitle.text = titleText
        With .chartTitle.Characters.Font
            .Name = "Meiryo UI"
            .Size = 14
            .Bold = True
        End With
    End With

End Sub

Sub パワークエリで追加して集計する()
    ' 依頼の語: 全部足|同じ形|全部追加|複数の表|月ごとの表|追加して集計|月の明細|合計を読|追加して合計
    ' 扱う: パワークエリ追加集計 合計 追加
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列 テーブル
    ' アクティブシートのテーブルと同じ見出しのテーブルを全ブックから集め、選んだ2列（文字=行、数=値）で集計し新シートに読み込む

    Dim wb As Object
    Set wb = ActiveSheet.Parent

    Dim wsBase As Object
    Set wsBase = ActiveSheet

    ' アクティブシートのテーブルを探す
    Dim loBase As Object
    If wsBase.ListObjects.Count = 0 Then Exit Sub
    Set loBase = wsBase.ListObjects(1)

    ' 選んでいる列から文字列列・数値列のヘッダ名を取得
    Dim selAreas As Object
    Set selAreas = Selection.Areas
    If selAreas.Count < 2 Then Exit Sub

    Dim colTxt As Long, colNum As Long
    colTxt = selAreas(1).Column
    colNum = selAreas(2).Column

    ' 見出し行を特定
    Dim hdrRow As Long
    hdrRow = loBase.HeaderRowRange.Row
    Dim hdrTxtName As String, hdrNumName As String
    hdrTxtName = CStr(wsBase.Cells(hdrRow, colTxt).Value)
    hdrNumName = CStr(wsBase.Cells(hdrRow, colNum).Value)

    ' ベーステーブルの全見出しを収集
    Dim baseHeaders() As String
    Dim baseHdrCount As Integer
    baseHdrCount = loBase.HeaderRowRange.Columns.Count
    ReDim baseHeaders(1 To baseHdrCount)
    Dim hc As Integer
    For hc = 1 To baseHdrCount
        baseHeaders(hc) = CStr(loBase.HeaderRowRange.Cells(1, hc).Value)
    Next hc

    ' クエリ名・シート名・テーブル名
    Dim baseTblName As String
    baseTblName = loBase.Name
    Dim qName As String
    qName = baseTblName & "_集計"
    Dim wsNewName As String
    wsNewName = qName

    ' 既存クエリ削除
    Dim q As Object
    Application.DisplayAlerts = False
    On Error Resume Next
    For Each q In wb.Queries
        If q.Name = qName Then
            q.Delete
            Exit For
        End If
    Next q
    On Error GoTo 0

    ' 既存シート削除
    Dim ws As Object
    For Each ws In wb.Sheets
        If ws.Name = wsNewName Then
            ws.Delete
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True

    ' 同じ見出しのテーブルを全シートから収集
    Dim tableNames() As String
    Dim tCount As Integer
    tCount = 0
    Dim s As Object
    For Each s In wb.Sheets
        Dim lo As Object
        For Each lo In s.ListObjects
            ' 見出し数チェック
            If lo.HeaderRowRange.Columns.Count <> baseHdrCount Then GoTo NextLO
            ' 全見出し語チェック（順不同）
            Dim hMatch As Boolean
            hMatch = True
            Dim hi As Integer
            For hi = 1 To baseHdrCount
                Dim hv As String
                hv = CStr(lo.HeaderRowRange.Cells(1, hi).Value)
                Dim found As Boolean
                found = False
                Dim bhi As Integer
                For bhi = 1 To baseHdrCount
                    If hv = baseHeaders(bhi) Then
                        found = True
                        Exit For
                    End If
                Next bhi
                If Not found Then
                    hMatch = False
                    Exit For
                End If
            Next hi
            If Not hMatch Then GoTo NextLO
            tCount = tCount + 1
            ReDim Preserve tableNames(1 To tCount)
            tableNames(tCount) = lo.Name
NextLO:
        Next lo
    Next s

    If tCount = 0 Then Exit Sub

    ' M式構築
    Dim nl As String
    nl = Chr(10)
    Dim m As String
    m = "let" & nl

    Dim i As Integer
    For i = 1 To tCount
        m = m & "    src" & i & " = Excel.CurrentWorkbook(){[Name=""" & tableNames(i) & """]}[Content]," & nl
    Next i
    For i = 1 To tCount
        m = m & "    sel" & i & " = Table.SelectColumns(src" & i & ", {""" & hdrTxtName & """, """ & hdrNumName & """})," & nl
        m = m & "    cast" & i & " = Table.TransformColumnTypes(sel" & i & ", {{""" & hdrTxtName & """, type text}, {""" & hdrNumName & """, type number}})," & nl
    Next i

    Dim combined As String
    If tCount = 1 Then
        combined = "cast1"
    Else
        combined = "Table.Combine({cast1"
        For i = 2 To tCount
            combined = combined & ", cast" & i
        Next i
        combined = combined & "})"
    End If
    m = m & "    結合 = " & combined & "," & nl
    m = m & "    集計 = Table.Group(結合, {""" & hdrTxtName & """}, {{""" & hdrNumName & """, each List.Sum([" & hdrNumName & "]), type number}})," & nl
    m = m & "    並替 = Table.Sort(集計, {{""" & hdrTxtName & """, Order.Ascending}})" & nl
    m = m & "in" & nl
    m = m & "    並替"

    wb.Queries.Add qName, m

    Dim wsNew As Worksheet
    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsNew.Name = wsNewName

    Dim loNew As ListObject
    Set loNew = wsNew.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties=""""", _
        Destination:=wsNew.Range("A1"))

    loNew.Name = qName
    loNew.QueryTable.CommandType = 2
    loNew.QueryTable.CommandText = Array("SELECT * FROM [" & qName & "]")
    loNew.QueryTable.Refresh False

End Sub

Sub 項目別合計表を右に作る()
    ' 依頼の語: 項目別合計|SUMIFで項目|項目ごとの合計表|項目別の合計表|合計した表|合計の表
    ' 扱う: 項目別合計 SUMIF合計表 行を足
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Dim colKey As Long, colVal As Long
    Dim ur As Range
    Dim hdrRow As Long, firstDataRow As Long, lastRow As Long
    Dim outCol As Long
    Dim dict As Object
    Dim keys() As String
    Dim keyCount As Long
    Dim r As Long, i As Long
    Dim keyColAddr As String, valColAddr As String
    Dim cellVal As Variant
    Dim s As String
    Dim urFirstRow As Long, urLastRow As Long, urRight As Long

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    If Selection.Areas.Count < 2 Then Exit Sub
    colKey = Selection.Areas(1).Column
    colVal = Selection.Areas(2).Column

    Set ur = ws.UsedRange
    urFirstRow = ur.Row
    urLastRow = ur.Row + ur.rows.Count - 1
    urRight = ur.Column + ur.Columns.Count - 1

    Dim testRow As Long, testC As Long, numCount As Long
    firstDataRow = 0
    For testRow = urFirstRow To urLastRow
        numCount = 0
        For testC = ur.Column To urRight
            cellVal = ws.Cells(testRow, testC).Value
            If Not isEmpty(cellVal) And cellVal <> "" Then
                If IsNumeric(cellVal) Or IsDate(cellVal) Then
                    numCount = numCount + 1
                End If
            End If
        Next testC
        If numCount >= 2 Then
            firstDataRow = testRow
            Exit For
        End If
    Next testRow
    If firstDataRow = 0 Then Exit Sub
    hdrRow = firstDataRow - 1
    If hdrRow < urFirstRow Then hdrRow = urFirstRow

    lastRow = urLastRow
    Dim checkStr As String
    Dim cc As Long
    Do While lastRow >= firstDataRow
        checkStr = ""
        For cc = ur.Column To urRight
            cellVal = ws.Cells(lastRow, cc).Value
            If Not isEmpty(cellVal) And cellVal <> "" Then
                checkStr = CStr(cellVal)
                Exit For
            End If
        Next cc
        If InStr(checkStr, "合計") > 0 Or InStr(checkStr, "総計") > 0 Or checkStr = "計" Then
            lastRow = lastRow - 1
        Else
            Exit Do
        End If
    Loop

    Dim hdrKeyName As String, hdrValName As String
    hdrKeyName = CStr(ws.Cells(hdrRow, colKey).Value)
    hdrValName = CStr(ws.Cells(hdrRow, colVal).Value)

    Dim existCol As Long
    existCol = 0
    Dim checkCol As Long
    For checkCol = ur.Column To urRight - 1
        If checkCol <> colKey Then
            If CStr(ws.Cells(hdrRow, checkCol).Value) = hdrKeyName And _
               CStr(ws.Cells(hdrRow, checkCol + 1).Value) = hdrValName Then
                existCol = checkCol
                Exit For
            End If
        End If
    Next checkCol

    If existCol > 0 Then
        outCol = existCol
        Dim clearR As Long
        For clearR = hdrRow To urLastRow + 30
            If clearR > hdrRow + 1 Then
                If isEmpty(ws.Cells(clearR, outCol).Value) And isEmpty(ws.Cells(clearR, outCol + 1).Value) Then
                    Exit For
                End If
            End If
            ws.Cells(clearR, outCol).ClearContents
            ws.Cells(clearR, outCol + 1).ClearContents
        Next clearR
    Else
        outCol = urRight + 2
    End If

    keyCount = 0
    ReDim keys(0)
    For r = firstDataRow To lastRow
        cellVal = ws.Cells(r, colKey).Value
        If Not isEmpty(cellVal) And cellVal <> "" Then
            s = CStr(cellVal)
            If Not dict.Exists(s) Then
                dict.Add s, 1
                ReDim Preserve keys(keyCount)
                keys(keyCount) = s
                keyCount = keyCount + 1
            End If
        End If
    Next r

    ws.Cells(hdrRow, outCol).Value = hdrKeyName
    ws.Cells(hdrRow, outCol + 1).Value = hdrValName

    Dim colLetterKey As String, colLetterVal As String
    colLetterKey = Mid(ws.Cells(firstDataRow, colKey).Address(True, True), 2)
    colLetterKey = Left(colLetterKey, InStr(colLetterKey, "$") - 1)
    colLetterVal = Mid(ws.Cells(firstDataRow, colVal).Address(True, True), 2)
    colLetterVal = Left(colLetterVal, InStr(colLetterVal, "$") - 1)

    keyColAddr = "$" & colLetterKey & "$" & firstDataRow & ":$" & colLetterKey & "$" & lastRow
    valColAddr = "$" & colLetterVal & "$" & firstDataRow & ":$" & colLetterVal & "$" & lastRow

    Dim outRow As Long
    outRow = hdrRow + 1
    For i = 0 To keyCount - 1
        ws.Cells(outRow, outCol).Value = keys(i)
        Dim criteriaAddr As String
        criteriaAddr = ws.Cells(outRow, outCol).Address(True, True)
        ws.Cells(outRow, outCol + 1).Formula = "=SUMIF(" & keyColAddr & "," & criteriaAddr & "," & valColAddr & ")"
        outRow = outRow + 1
    Next i

    ws.Cells(outRow, outCol).Value = "合計"
    If outRow > hdrRow + 1 Then
        Dim sumStart As String, sumEnd As String
        sumStart = ws.Cells(hdrRow + 1, outCol + 1).Address(True, True)
        sumEnd = ws.Cells(outRow - 1, outCol + 1).Address(True, True)
        ws.Cells(outRow, outCol + 1).Formula = "=SUM(" & sumStart & ":" & sumEnd & ")"
    End If

End Sub

Sub 項目別件数表を右に作る()
    ' 依頼の語: 項目別件数|項目ごとに何件|項目ごとの件数|選んだ列で項目|項目別の件数|件数表を右|列で件数
    ' 扱う: 項目別件数 COUNTIF件数表 合計 行を足
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

    ' 選んでいる列の項目ごとの件数表を明細の右に1列空けて作る

    Dim ws As Worksheet
    Dim selCol As Long
    Dim ur As Range
    Dim hdrRow As Long, dataFirst As Long, dataLast As Long
    Dim r As Long, c As Long
    Dim outCol As Long
    Dim dict As Object
    Dim keys() As String
    Dim keyCount As Long
    Dim cellVal As String
    Dim s As String
    Dim i As Long
    Dim colLetResult As String
    Dim colLetInput As Long

    Set ws = ActiveSheet
    selCol = Selection.Cells(1, 1).Column

    Set ur = ws.UsedRange
    Dim urFirstRow As Long, urLastRow As Long, urFirstCol As Long, urLastCol As Long
    urFirstRow = ur.Row
    urLastRow = ur.Row + ur.rows.Count - 1
    urFirstCol = ur.Column
    urLastCol = ur.Column + ur.Columns.Count - 1

    hdrRow = 0
    Dim rr As Long
    For rr = urFirstRow To urLastRow
        Dim numCount As Long
        numCount = 0
        Dim cv As Variant
        For c = urFirstCol To urLastCol
            cv = ws.Cells(rr, c).Value
            If Not isEmpty(cv) Then
                If IsNumeric(cv) Or IsDate(cv) Then
                    numCount = numCount + 1
                End If
            End If
        Next c
        If numCount >= 2 Then
            If rr > urFirstRow Then
                hdrRow = rr - 1
            Else
                hdrRow = rr
            End If
            dataFirst = rr
            Exit For
        End If
    Next rr

    If hdrRow = 0 Then Exit Sub

    dataLast = dataFirst
    Dim rowLabel As String
    For rr = urLastRow To dataFirst Step -1
        rowLabel = ""
        For c = urFirstCol To urLastCol
            Dim rv As String
            rv = Trim(CStr(ws.Cells(rr, c).Value))
            If rv <> "" Then
                rowLabel = rv
                Exit For
            End If
        Next c
        If rowLabel = "合計" Or rowLabel = "計" Or rowLabel = "小計" Or rowLabel = "総計" Then
            ' skip
        Else
            dataLast = rr
            Exit For
        End If
    Next rr

    Dim selHdr As String
    selHdr = CStr(ws.Cells(hdrRow, selCol).Value)

    ' 既存の出力表を探す（見出し行で selHdr が出てくる列を右側から探す）
    outCol = 0
    Dim realLastCol As Long
    realLastCol = urFirstCol - 1
    For c = urFirstCol To urFirstCol + 200
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) <> "" Then
            realLastCol = c
        End If
    Next c

    Dim srcLastCol As Long
    ' 明細の最終列を求める（selCol が属する明細ブロック）
    srcLastCol = urFirstCol - 1
    For c = urFirstCol To selCol + 50
        Dim hv As String
        hv = Trim(CStr(ws.Cells(hdrRow, c).Value))
        If hv <> "" Then
            srcLastCol = c
        End If
        ' 空白列の後を出力域とするため selCol より右で最初の空列を探す
    Next c
    ' 明細の右の空列の次が出力先候補 → selHdr が既にあればそこ
    Dim checkFrom As Long
    checkFrom = srcLastCol + 2
    ' 明細最終列は selCol 以降の連続した値ある列の末尾
    Dim mLastCol As Long
    mLastCol = selCol
    For c = selCol To selCol + 100
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) <> "" Then
            mLastCol = c
        Else
            Exit For
        End If
    Next c
    checkFrom = mLastCol + 2

    ' 既存の出力表を探す
    For c = checkFrom To realLastCol
        Dim hval As String
        hval = Trim(CStr(ws.Cells(hdrRow, c).Value))
        ' 正規化して比較
        Dim hs As String
        hs = hval: GoSub 正規化H: hval = hs
        Dim sh2 As String
        sh2 = selHdr: GoSub 正規化H: sh2 = hs
        If hval = sh2 Then
            outCol = c
            Exit For
        End If
    Next c

    If outCol = 0 Then
        ' 明細の右に1列空けた列
        outCol = mLastCol + 2
    End If

    ' 既存出力をクリア
    Dim clearRow As Long
    For clearRow = hdrRow To hdrRow + 500
        Dim cv1 As String, cv2 As String
        cv1 = Trim(CStr(ws.Cells(clearRow, outCol).Value))
        cv2 = Trim(CStr(ws.Cells(clearRow, outCol + 1).Value))
        If clearRow > hdrRow And cv1 = "" And cv2 = "" Then Exit For
        ws.Cells(clearRow, outCol).ClearContents
        ws.Cells(clearRow, outCol + 1).ClearContents
    Next clearRow

    ws.Cells(hdrRow, outCol).Value = selHdr
    ws.Cells(hdrRow, outCol + 1).Value = "件数"

    ' 項目を初出順に収集（セル値をそのまま使う・正規化はキー比較のみ）
    Set dict = CreateObject("Scripting.Dictionary")
    keyCount = 0
    ReDim keys(0)
    Dim rawKeys() As String
    ReDim rawKeys(0)

    For rr = dataFirst To dataLast
        Dim rawVal As String
        rawVal = CStr(ws.Cells(rr, selCol).Value)
        s = rawVal: GoSub 正規化: cellVal = s
        If cellVal = "" Then GoTo NextRow
        If Not dict.Exists(cellVal) Then
            dict.Add cellVal, 1
            ReDim Preserve keys(keyCount)
            ReDim Preserve rawKeys(keyCount)
            keys(keyCount) = cellVal
            rawKeys(keyCount) = rawVal
            keyCount = keyCount + 1
        End If
NextRow:
    Next rr

    colLetInput = selCol: GoSub ColLetSub
    Dim selColLet As String
    selColLet = colLetResult

    Dim absRange As String
    absRange = "$" & selColLet & "$" & dataFirst & ":$" & selColLet & "$" & dataLast

    colLetInput = outCol: GoSub ColLetSub
    Dim outColLet As String
    outColLet = colLetResult

    colLetInput = outCol + 1: GoSub ColLetSub
    Dim outCol1Let As String
    outCol1Let = colLetResult

    Dim outRow As Long
    For i = 0 To keyCount - 1
        outRow = hdrRow + 1 + i
        ws.Cells(outRow, outCol).Value = rawKeys(i)
        ws.Cells(outRow, outCol + 1).Formula = "=COUNTIF(" & absRange & "," & outColLet & outRow & ")"
    Next i

    Dim sumRow As Long
    sumRow = hdrRow + 1 + keyCount
    ws.Cells(sumRow, outCol).Value = "合計"
    Dim sumFirst As Long, sumLast As Long
    sumFirst = hdrRow + 1
    sumLast = sumRow - 1
    ws.Cells(sumRow, outCol + 1).Formula = "=SUM(" & outCol1Let & sumFirst & ":" & outCol1Let & sumLast & ")"

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbWide)): s = Replace(s, ",", "")
    Return

正規化H:
    hs = Trim$(StrConv(hs, vbWide)): hs = Replace(hs, ",", "")
    Return

ColLetSub:
    Dim tmpCol As Long
    Dim tmpS As String
    tmpCol = colLetInput
    tmpS = ""
    Do While tmpCol > 0
        Dim m As Long
        m = (tmpCol - 1) Mod 26
        tmpS = Chr(65 + m) & tmpS
        tmpCol = (tmpCol - 1) \ 26
    Loop
    colLetResult = tmpS
    Return

End Sub

Sub 累計列を挿入する()
    ' 依頼の語: 累計列|累計を列|累計を出|選んだ列の累計|積み上げの累計|上から順の累計|累計の列を右|先頭の本文
    ' 扱う: 累計列 累計挿入 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    ' 選んでいる数の列のすぐ右に累計列を挿入し SUM式を入れる（既存の累計列なら書き直す）

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urTop As Long, urLeft As Long, urBottom As Long, urRight As Long
    urTop = ur.Row
    urLeft = ur.Column
    urBottom = ur.Row + ur.rows.Count - 1
    urRight = ur.Column + ur.Columns.Count - 1

    ' 見出し行を探す
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            Dim hasNum As Boolean
            hasNum = False
            Dim rr As Long
            For rr = r + 1 To urBottom
                Dim lineEmpty As Boolean
                lineEmpty = True
                For c = urLeft To urRight
                    If ws.Cells(rr, c).Value <> "" Then lineEmpty = False: Exit For
                Next c
                If Not lineEmpty Then
                    For c = urLeft To urRight
                        Dim nv As Variant
                        nv = ws.Cells(rr, c).Value
                        If IsNumeric(nv) And nv <> "" Then hasNum = True
                    Next c
                    Exit For
                End If
            Next rr
            If hasNum Then
                hdrRow = r
                Exit For
            End If
        End If
    Next r
    If hdrRow = 0 Then hdrRow = urTop

    Dim dataStart As Long
    dataStart = hdrRow + 1

    ' 最後の本文行（合計行・空行を除く）
    Dim lastDataRow As Long
    lastDataRow = 0
    For r = urBottom To dataStart Step -1
        Dim isSumR As Boolean
        isSumR = False
        For c = urLeft To urRight
            Dim v As String
            v = CStr(ws.Cells(r, c).Value)
            If v = "合計" Or v = "計" Or v = "小計" Or v = "総計" Then isSumR = True: Exit For
        Next c
        If Not isSumR Then
            Dim rowEmpty As Boolean
            rowEmpty = True
            For c = urLeft To urRight
                If ws.Cells(r, c).Value <> "" Then rowEmpty = False: Exit For
            Next c
            If Not rowEmpty Then
                lastDataRow = r
                Exit For
            End If
        End If
    Next r
    If lastDataRow = 0 Then lastDataRow = urBottom

    Dim insCol As Long
    insCol = selCol + 1

    ' 既存累計列チェック
    Dim alreadyExists As Boolean
    alreadyExists = False
    If insCol <= urRight Then
        If CStr(ws.Cells(hdrRow, insCol).Value) = "累計" Then
            alreadyExists = True
        End If
    End If

    ' 元列の表示形式
    Dim srcFmt As String
    srcFmt = "#,##0"
    Dim fmtStr As String
    For r = dataStart To lastDataRow
        Dim fv As Variant
        fv = ws.Cells(r, selCol).Value
        If IsNumeric(fv) And fv <> "" Then
            fmtStr = ws.Cells(r, selCol).NumberFormat
            If fmtStr <> "" And fmtStr <> "General" And fmtStr <> "標準" Then srcFmt = fmtStr
            Exit For
        End If
    Next r

    If Not alreadyExists Then
        ws.Columns(insCol).Insert Shift:=xlToRight
        urRight = urRight + 1
    End If

    ws.Cells(hdrRow, insCol).Value = "累計"

    ' 列名を取得
    Dim colLetter As String
    Dim tmpAddr As String
    tmpAddr = ws.Cells(1, selCol).Address(False, False)
    ' tmpAddr は "B1" など。列部分を取り出す
    Dim i As Long
    colLetter = ""
    For i = 1 To Len(tmpAddr)
        Dim ch As String
        ch = Mid(tmpAddr, i, 1)
        If ch >= "A" And ch <= "Z" Then
            colLetter = colLetter & ch
        Else
            Exit For
        End If
    Next i

    ' 先頭本文行（空行・合計行をスキップ）
    Dim firstDataRow As Long
    firstDataRow = 0
    For r = dataStart To lastDataRow
        Dim fe As Boolean
        fe = True
        For c = urLeft To urRight
            If c = insCol Then GoTo SkipFE
            If ws.Cells(r, c).Value <> "" Then fe = False: Exit For
SkipFE:
        Next c
        If fe Then GoTo NextFirst
        Dim isSumF As Boolean
        isSumF = False
        For c = urLeft To urRight
            If c = insCol Then GoTo SkipFS
            Dim vs As String
            vs = CStr(ws.Cells(r, c).Value)
            If vs = "合計" Or vs = "計" Or vs = "小計" Or vs = "総計" Then isSumF = True: Exit For
SkipFS:
        Next c
        If Not isSumF Then firstDataRow = r: Exit For
NextFirst:
    Next r
    If firstDataRow = 0 Then firstDataRow = dataStart

    ' 先頭セルの絶対参照 例: $B$2
    Dim firstAddr As String
    firstAddr = "$" & colLetter & "$" & firstDataRow

    For r = dataStart To lastDataRow
        ' 空行チェック
        Dim re As Boolean
        re = True
        For c = urLeft To urRight
            If c = insCol Then GoTo SkipC1
            If ws.Cells(r, c).Value <> "" Then re = False: Exit For
SkipC1:
        Next c
        If re Then GoTo NextRow

        ' 合計行チェック
        Dim isSumRow As Boolean
        isSumRow = False
        For c = urLeft To urRight
            If c = insCol Then GoTo SkipC2
            Dim vv As String
            vv = CStr(ws.Cells(r, c).Value)
            If vv = "合計" Or vv = "計" Or vv = "小計" Or vv = "総計" Then isSumRow = True: Exit For
SkipC2:
        Next c
        If isSumRow Then GoTo NextRow

        ' 現在行のセルは列のみ絶対参照: $B3
        Dim curAddr As String
        curAddr = "$" & colLetter & r

        ws.Cells(r, insCol).Formula = "=SUM(" & firstAddr & ":" & curAddr & ")"
        ws.Cells(r, insCol).NumberFormat = srcFmt

NextRow:
    Next r

End Sub

Sub 構成比列を挿入する()
    ' 依頼の語: 構成比の列|構成比を列|構成比を出|割合の列|全体に占める割合|金額の構成比
    ' 扱う: 構成比 割合 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange

    Dim urFirstRow As Long, urFirstCol As Long, urLastRow As Long, urLastCol As Long
    urFirstRow = ur.Row
    urFirstCol = ur.Column
    urLastRow = ur.Row + ur.rows.Count - 1
    urLastCol = ur.Column + ur.Columns.Count - 1

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim bodyFirst As Long
    bodyFirst = hdrRow + 1

    Dim insCol As Long
    insCol = selCol + 1

    Dim alreadyExists As Boolean
    alreadyExists = False
    Dim s As String
    Dim v As Variant

    If insCol <= urLastCol Then
        s = CStr(ws.Cells(hdrRow, insCol).Value): GoSub 正規化
        If s = "構成比" Then alreadyExists = True
    End If

    If Not alreadyExists Then
        ws.Columns(insCol).Insert Shift:=xlToRight
        urLastCol = urLastCol + 1
    End If

    ws.Cells(hdrRow, insCol).Value = "構成比"

    Dim sumKeywords As Variant
    sumKeywords = Array("合計", "計", "小計", "総計")
    Dim isSumRow As Boolean
    Dim ki As Variant
    Dim r As Long, c As Long

    ' 本文最終行（空行・合計行を除く）
    Dim dataLast As Long
    dataLast = bodyFirst - 1
    Dim rowEmpty As Boolean

    For r = bodyFirst To urLastRow
        rowEmpty = True
        For c = urFirstCol To urLastCol
            If c = insCol Then GoTo skipE0
            If Not isEmpty(ws.Cells(r, c)) And CStr(ws.Cells(r, c).Value) <> "" Then rowEmpty = False: Exit For
skipE0:
        Next c
        If rowEmpty Then GoTo nextR0

        isSumRow = False
        For Each ki In sumKeywords
            For c = urFirstCol To urLastCol
                If c = insCol Then GoTo skipS0
                s = CStr(ws.Cells(r, c).Value): GoSub 正規化
                If s <> "" And InStr(s, CStr(ki)) > 0 Then isSumRow = True
skipS0:
            Next c
            If isSumRow Then Exit For
        Next ki
        If Not isSumRow Then dataLast = r
nextR0:
    Next r

    If dataLast < bodyFirst Then GoTo exitSub

    Dim clResult As String
    Dim clCol As Long
    clCol = selCol: GoSub colLetter
    Dim selColLetter As String
    selColLetter = clResult

    Dim sumRange As String
    sumRange = "$" & selColLetter & "$" & bodyFirst & ":$" & selColLetter & "$" & dataLast

    For r = bodyFirst To urLastRow
        rowEmpty = True
        For c = urFirstCol To urLastCol
            If c = insCol Then GoTo skipE2
            If Not isEmpty(ws.Cells(r, c)) And CStr(ws.Cells(r, c).Value) <> "" Then rowEmpty = False: Exit For
skipE2:
        Next c

        isSumRow = False
        For Each ki In sumKeywords
            For c = urFirstCol To urLastCol
                If c = insCol Then GoTo skipS2
                s = CStr(ws.Cells(r, c).Value): GoSub 正規化
                If s <> "" And InStr(s, CStr(ki)) > 0 Then isSumRow = True
skipS2:
            Next c
            If isSumRow Then Exit For
        Next ki

        If rowEmpty Or isSumRow Then
            ws.Cells(r, insCol).ClearContents
        Else
            ws.Cells(r, insCol).Formula = "=" & selColLetter & r & "/SUM(" & sumRange & ")"
            ws.Cells(r, insCol).NumberFormat = "0.0%"
        End If
    Next r

exitSub:
    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow)): s = Replace(s, ",", "")
    Return

colLetter:
    clResult = ""
    Dim tmpCol As Long
    tmpCol = clCol
    Do While tmpCol > 0
        Dim m As Long
        m = (tmpCol - 1) Mod 26
        clResult = Chr(65 + m) & clResult
        tmpCol = (tmpCol - 1 - m) \ 26
    Loop
    Return

End Sub

Sub 増減額と増減率の列を足す()
    ' 依頼の語: 増減の列|差額と伸び率|前年と今年の増減|選んだ2列で増減|差と伸び率の列|増減額の列
    ' 扱う: 増減額 増減率 合計
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列

    Dim ws As Worksheet
    Dim sel1 As Range, sel2 As Range
    Dim hdrRow As Long, lastRow As Long, lastCol As Long
    Dim colD As Long, colE As Long
    Dim r As Long
    Dim c As Long
    Dim ws2 As Worksheet

    Set ws = ActiveSheet

    If Selection.Areas.Count < 2 Then Exit Sub
    Set sel1 = Selection.Areas(1)
    Set sel2 = Selection.Areas(2)

    hdrRow = sel1.Cells(1, 1).Row

    ' 表の最終列を UsedRange から求める
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    ' 本文の最終行（見出し行の次から UsedRange の末尾まで）
    Dim urLastRow As Long
    urLastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
    lastRow = urLastRow

    ' 既に「増減額」「増減率」の列があるか探す
    colD = 0: colE = 0
    Dim searchCol As Long
    For searchCol = 1 To lastCol
        Dim hv As String
        hv = CStr(ws.Cells(hdrRow, searchCol).Value)
        If hv = "増減額" Then colD = searchCol
        If hv = "増減率" Then colE = searchCol
    Next searchCol

    ' なければ右に追加
    If colD = 0 Then
        colD = lastCol + 1
        ws.Cells(hdrRow, colD).Value = "増減額"
    End If
    If colE = 0 Then
        ' colE は colD の右
        If colD > lastCol Then
            colE = colD + 1
        Else
            ' colD が既存列なら colE は lastCol+1 か colD+1
            colE = colD + 1
            ' ただし colE に別の見出しがあれば lastCol+1 へ
            Dim hvE As String
            hvE = CStr(ws.Cells(hdrRow, colE).Value)
            If hvE <> "" And hvE <> "増減率" Then colE = lastCol + 1
        End If
        ws.Cells(hdrRow, colE).Value = "増減率"
    End If

    ' 合計行キーワード
    Dim sumWords(3) As String
    sumWords(0) = "合計": sumWords(1) = "計": sumWords(2) = "小計": sumWords(3) = "総計"

    ' 本文行に数式を書く
    For r = hdrRow + 1 To lastRow
        ' 空行チェック：行全体が空なら skip
        Dim rowEmpty As Boolean
        rowEmpty = True
        For c = ws.UsedRange.Column To lastCol
            If ws.Cells(r, c).Value <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If rowEmpty Then GoTo NextRow

        ' 合計行チェック：最初の列の値が合計系語なら skip
        Dim firstVal As String
        firstVal = CStr(ws.Cells(r, ws.UsedRange.Column).Value)
        Dim isSumRow As Boolean
        isSumRow = False
        Dim si As Integer
        For si = 0 To 3
            If firstVal = sumWords(si) Then isSumRow = True
        Next si
        ' 他の列も確認
        If Not isSumRow Then
            For c = ws.UsedRange.Column To lastCol
                Dim cv As String
                cv = CStr(ws.Cells(r, c).Value)
                For si = 0 To 3
                    If cv = sumWords(si) Then isSumRow = True
                Next si
            Next c
        End If
        If isSumRow Then GoTo NextRow

        ' sel1・sel2 の同じ行のセルアドレスを使って数式を組む
        Dim c1Addr As String, c2Addr As String
        c1Addr = ws.Cells(r, sel1.Column).Address(False, True)
        c2Addr = ws.Cells(r, sel2.Column).Address(False, True)

        ' 増減額
        ws.Cells(r, colD).Formula = "=" & c2Addr & "-" & c1Addr
        ws.Cells(r, colD).NumberFormat = "#,##0"

        ' 増減率
        ws.Cells(r, colE).Formula = "=(" & c2Addr & "-" & c1Addr & ")/" & c1Addr
        ws.Cells(r, colE).NumberFormat = "0.0%"

NextRow:
    Next r

End Sub

Sub 日付から年度月四半期列を足す()
    ' 依頼の語: 四半期の列|年度と月と四半期|年度・月・四半期|日付から年度|日付を年度|年度と四半期の列|年度月四半期の列
    ' 扱う: 年度 月 四半期 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列
    ' 選んでいる日付の列を基に、年度・月・四半期の3列を表の右端の右に追加する

    Dim ws As Worksheet
    Dim selCol As Long
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colNendo As Long, colMonth As Long, colQuarter As Long
    Dim r As Long
    Dim v As Variant
    Dim s As String
    Dim ur As Range
    Dim c As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    selCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    lastCol = ur.Column + ur.Columns.Count - 1
    lastRow = ur.Row + ur.rows.Count - 1

    colNendo = 0: colMonth = 0: colQuarter = 0
    For c = 1 To lastCol
        s = Trim(CStr(ws.Cells(hdrRow, c).Value))
        If s = "年度" Then colNendo = c
        If s = "月" Then colMonth = c
        If s = "四半期" Then colQuarter = c
    Next c

    If colNendo = 0 Then
        colNendo = lastCol + 1
        ws.Cells(hdrRow, colNendo).Value = "年度"
    End If
    If colMonth = 0 Then
        colMonth = colNendo + 1
        ws.Cells(hdrRow, colMonth).Value = "月"
    End If
    If colQuarter = 0 Then
        colQuarter = colMonth + 1
        ws.Cells(hdrRow, colQuarter).Value = "四半期"
    End If

    ws.Cells(hdrRow, colNendo).NumberFormatLocal = "G/標準"
    ws.Cells(hdrRow, colMonth).NumberFormatLocal = "G/標準"
    ws.Cells(hdrRow, colQuarter).NumberFormatLocal = "G/標準"

    Dim dateAddr As String
    Dim rowEmpty As Boolean
    Dim isTotal As Boolean
    Dim cc As Long

    For r = hdrRow + 1 To lastRow
        rowEmpty = True
        For cc = ur.Column To lastCol
            If ws.Cells(r, cc).Value <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next cc
        If rowEmpty Then GoTo NextRow

        isTotal = False
        For cc = ur.Column To lastCol
            v = ws.Cells(r, cc).Value
            If VarType(v) = vbString Then
                s = Trim(CStr(v))
                If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then
                    isTotal = True
                    Exit For
                End If
            End If
        Next cc
        If isTotal Then GoTo NextRow

        dateAddr = ws.Cells(r, selCol).Address(True, True)

        ws.Cells(r, colNendo).NumberFormatLocal = "G/標準"
        ws.Cells(r, colNendo).Formula = "=YEAR(" & dateAddr & ")-IF(MONTH(" & dateAddr & ")<4,1,0)"

        ws.Cells(r, colMonth).NumberFormatLocal = "G/標準"
        ws.Cells(r, colMonth).Formula = "=MONTH(" & dateAddr & ")"

        ws.Cells(r, colQuarter).NumberFormatLocal = "G/標準"
        ws.Cells(r, colQuarter).Formula = "=INT(MOD(MONTH(" & dateAddr & ")-4,12)/3)+1"

NextRow:
    Next r

End Sub

Sub 上位件数を右に抜き出す()
    ' 依頼の語: 上位5件|上位 5 件|トップ5|ベスト5|上位5|大きい順に5件
    ' 扱う: 上位抜き出し 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Object
    Dim selCol As Long
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim itemCol As Long
    Dim i As Long, k As Long
    Dim topCount As Long
    Dim outStartCol As Long
    Dim ur As Object
    Dim c As Long
    Dim isText As Boolean
    Dim hasVal As Boolean
    Dim s As String
    Dim colLetterResult As String
    Dim colLetterInput As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    selCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    lastRow = hdrRow
    For i = hdrRow + 1 To ur.Row + ur.rows.Count - 1
        s = Trim(CStr(ws.Cells(i, selCol).Value))
        If s <> "" Then
            Dim isSum As Boolean
            isSum = False
            Dim checkC As Long
            For checkC = ur.Column To ur.Column + ur.Columns.Count - 1
                Dim sv As String
                sv = Trim(CStr(ws.Cells(i, checkC).Value))
                If sv = "合計" Or sv = "計" Or sv = "小計" Or sv = "総計" Then
                    isSum = True
                    Exit For
                End If
            Next checkC
            If Not isSum Then lastRow = i
        End If
    Next i

    itemCol = -1
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        isText = True
        hasVal = False
        For i = hdrRow + 1 To hdrRow + 3
            If i > lastRow Then Exit For
            Dim v As Variant
            v = ws.Cells(i, c).Value
            If v <> "" Then
                hasVal = True
                If IsNumeric(v) And Not IsDate(v) Then isText = False
                If IsDate(v) Then isText = False
            End If
        Next i
        If isText And hasVal Then
            itemCol = c
            Exit For
        End If
    Next c
    If itemCol = -1 Then itemCol = ur.Column

    Dim bodyCount As Long
    bodyCount = 0
    For i = hdrRow + 1 To lastRow
        If ws.Cells(i, selCol).Value <> "" Then bodyCount = bodyCount + 1
    Next i
    topCount = bodyCount
    If topCount > 5 Then topCount = 5

    Dim urEnd As Long
    urEnd = ur.Column + ur.Columns.Count - 1

    Dim selHdr As String
    selHdr = CStr(ws.Cells(hdrRow, selCol).Value)
    Dim itemHdr As String
    itemHdr = CStr(ws.Cells(hdrRow, itemCol).Value)

    Dim candidateCol As Long
    candidateCol = -1
    For c = urEnd To ur.Column Step -1
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) = selHdr Then
            If c > ur.Column Then
                If Trim(CStr(ws.Cells(hdrRow, c - 1).Value)) = itemHdr Then
                    If c <> selCol Then
                        candidateCol = c - 1
                        Exit For
                    End If
                End If
            End If
        End If
    Next c

    If candidateCol = -1 Then
        outStartCol = urEnd + 2
    Else
        outStartCol = candidateCol
    End If

    Dim r1 As Long, r2 As Long
    r1 = hdrRow + 1
    r2 = lastRow

    Dim numAddr As String
    Dim itemAddr As String

    colLetterInput = selCol: GoSub colLetter: numAddr = "$" & colLetterResult & "$" & r1 & ":$" & colLetterResult & "$" & r2
    colLetterInput = itemCol: GoSub colLetter: itemAddr = "$" & colLetterResult & "$" & r1 & ":$" & colLetterResult & "$" & r2

    Dim outItemCol As Long
    Dim outNumCol As Long
    outItemCol = outStartCol
    outNumCol = outStartCol + 1

    ws.Cells(hdrRow, outItemCol).Value = itemHdr
    ws.Cells(hdrRow, outNumCol).Value = selHdr

    For k = 1 To topCount
        Dim outR As Long
        outR = hdrRow + k
        ws.Cells(outR, outNumCol).Formula = "=LARGE(" & numAddr & "," & k & ")"
        ws.Cells(outR, outNumCol).NumberFormat = "#,##0"
        ws.Cells(outR, outItemCol).Formula = "=INDEX(" & itemAddr & ",MATCH(LARGE(" & numAddr & "," & k & ")," & numAddr & ",0))"
    Next k

    Dim clearR As Long
    For clearR = hdrRow + topCount + 1 To hdrRow + 10
        If ws.Cells(clearR, outItemCol).Value = "" And ws.Cells(clearR, outNumCol).Value = "" Then Exit For
        ws.Cells(clearR, outItemCol).ClearContents
        ws.Cells(clearR, outNumCol).ClearContents
    Next clearR

    Exit Sub

colLetter:
    Dim cl As String
    Dim cc As Long
    cl = ""
    cc = colLetterInput
    Do While cc > 0
        Dim m As Long
        m = (cc - 1) Mod 26
        cl = Chr(65 + m) & cl
        cc = (cc - 1) \ 26
    Loop
    colLetterResult = cl
    Return

End Sub

Sub 区分ごとに小計行を挿入する()
    ' 依頼の語: 小計の行|切れ目|区分ごとに小計|小計行を作|小計を入|区分が変|小計の行を区切|切れ目に小計|小計行
    ' 扱う: 小計行挿入 区分小計 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Dim sel As Range
    Dim hdrRow As Long, firstDataRow As Long, lastRow As Long
    Dim grpCol As Long
    Dim i As Long, j As Long, c As Long
    Dim nCols As Long
    Dim usedLeft As Long
    Dim sumRanges() As Boolean
    Dim totRow As Long
    Dim s As String
    Dim lastUsedRow As Long
    Dim cc As Long
    Dim skipCol As Boolean
    Dim isNumCol As Boolean
    Dim colIdx As Long
    Dim cv As Variant
    Dim fmt As String
    Dim curGrp As String
    Dim boundaries() As Long
    Dim bCount As Long
    Dim bStart As Long
    Dim insertAfter As Long
    Dim r1 As String, r2 As String
    Dim hdrS As String
    Dim grpVal As String
    Dim k As Integer
    Dim skipKeywords(3) As String

    Set ws = ActiveSheet
    Set sel = Selection

    hdrRow = sel.Cells(1, 1).Row
    grpCol = sel.Cells(1, 1).Column
    firstDataRow = hdrRow + 1

    usedLeft = ws.UsedRange.Column
    nCols = ws.UsedRange.Columns.Count
    lastUsedRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    skipKeywords(0) = "番号"
    skipKeywords(1) = "No"
    skipKeywords(2) = "コード"
    skipKeywords(3) = "年度"

    ' 合計行を探す（firstDataRow以降で「合計」「計」「総計」を含むセル）
    totRow = 0
    For i = lastUsedRow To firstDataRow Step -1
        For cc = usedLeft To usedLeft + nCols - 1
            s = CStr(ws.Cells(i, cc).Value): GoSub 正規化
            If s = "合計" Or s = "計" Or s = "総計" Then
                totRow = i
                Exit For
            End If
            ' 「合計」が先頭にある場合も考慮（例：「合計行」など）
            If Left$(s, 2) = "合計" Or Left$(s, 1) = "計" Then
                totRow = i
                Exit For
            End If
        Next cc
        If totRow > 0 Then Exit For
    Next i

    If totRow > 0 Then
        lastRow = totRow - 1
    Else
        lastRow = lastUsedRow
    End If

    ' 既存の小計行を削除（後ろから）
    i = lastRow
    Do While i >= firstDataRow
        s = CStr(ws.Cells(i, grpCol).Value)
        Dim sN As String
        sN = Trim$(s)
        If Len(sN) >= 3 And Right$(sN, 2) = "小計" Then
            ws.rows(i).Delete
            lastRow = lastRow - 1
            If totRow > 0 Then totRow = totRow - 1
        Else
            i = i - 1
        End If
    Loop

    ' 数の列を特定
    ReDim sumRanges(1 To nCols)
    For c = 1 To nCols
        colIdx = usedLeft + c - 1
        hdrS = CStr(ws.Cells(hdrRow, colIdx).Value)
        skipCol = False
        Dim hdrU As String
        hdrU = UCase(hdrS)
        For k = 0 To 3
            Dim kw As String
            kw = skipKeywords(k)
            If InStr(hdrS, kw) > 0 Then skipCol = True
            If InStr(hdrU, UCase(kw)) > 0 Then skipCol = True
            Dim hdrHalf As String
            hdrHalf = hdrS
            Dim ci As Integer
            Dim tmp As String
            tmp = ""
            Dim ch As String
            For ci = 1 To Len(hdrHalf)
                ch = Mid$(hdrHalf, ci, 1)
                Dim code As Long
                code = AscW(ch)
                If code >= &HFF21 And code <= &HFF3A Then
                    tmp = tmp & Chr(code - &HFF21 + 65)
                ElseIf code >= &HFF41 And code <= &HFF5A Then
                    tmp = tmp & Chr(code - &HFF41 + 97)
                ElseIf code >= &HFF10 And code <= &HFF19 Then
                    tmp = tmp & Chr(code - &HFF10 + 48)
                Else
                    tmp = tmp & ch
                End If
            Next ci
            If InStr(UCase(tmp), UCase(kw)) > 0 Then skipCol = True
        Next k
        If Not skipCol Then
            For j = firstDataRow To lastRow
                cv = ws.Cells(j, colIdx).Value
                If CStr(cv) <> "" Then
                    fmt = ws.Cells(j, colIdx).NumberFormat
                    If InStr(fmt, "%") > 0 Then skipCol = True
                    Exit For
                End If
            Next j
        End If
        isNumCol = False
        If Not skipCol Then
            For j = firstDataRow To lastRow
                cv = ws.Cells(j, colIdx).Value
                If IsNumeric(cv) And CStr(cv) <> "" Then
                    isNumCol = True
                    Exit For
                End If
            Next j
        End If
        sumRanges(c) = isNumCol And Not skipCol
    Next c

    ' グループ境界を収集
    bCount = 0
    ReDim boundaries(1 To lastRow - firstDataRow + 2)
    curGrp = CStr(ws.Cells(firstDataRow, grpCol).Value)
    For i = firstDataRow + 1 To lastRow
        s = CStr(ws.Cells(i, grpCol).Value)
        If s <> curGrp Then
            bCount = bCount + 1
            boundaries(bCount) = i - 1
            curGrp = s
        End If
    Next i
    bCount = bCount + 1
    boundaries(bCount) = lastRow

    ' 後ろから小計行を挿入
    For j = bCount To 1 Step -1
        insertAfter = boundaries(j)
        If j = 1 Then
            bStart = firstDataRow
        Else
            bStart = boundaries(j - 1) + 1
        End If
        grpVal = CStr(ws.Cells(bStart, grpCol).Value)

        ws.rows(insertAfter + 1).Insert Shift:=xlDown

        ws.Cells(insertAfter + 1, grpCol).Value = grpVal & " 小計"

        For c = 1 To nCols
            colIdx = usedLeft + c - 1
            If sumRanges(c) And colIdx <> grpCol Then
                r1 = ws.Cells(bStart, colIdx).Address(False, True)
                r2 = ws.Cells(insertAfter, colIdx).Address(False, True)
                ws.Cells(insertAfter + 1, colIdx).Formula = "=SUM(" & r1 & ":" & r2 & ")"
            End If
        Next c
    Next j

    ' 合計行があれば数式を更新（小計行を除いた元の値だけ合計）
    If totRow > 0 Then
        Dim newTotRow As Long
        newTotRow = totRow + bCount
        For c = 1 To nCols
            colIdx = usedLeft + c - 1
            If sumRanges(c) And colIdx <> grpCol Then
                ' 小計行のSUMを合計する（小計行だけ足す）
                Dim sumFormula As String
                sumFormula = "=SUM("
                Dim isFirst As Boolean
                isFirst = True
                Dim si As Long
                For si = 1 To bCount
                    Dim subtotalRow As Long
                    ' 小計行の位置を計算（後ろから挿入したので前から数える）
                    ' 小計行は各グループの末尾に挿入済み
                    ' bStart of group si
                    Dim gStart As Long, gEnd As Long
                    If si = 1 Then
                        gStart = firstDataRow
                    Else
                        gStart = boundaries(si - 1) + 1 + (si - 1)
                    End If
                    ' 小計行は gStart + (元のグループ行数) 行目
                    ' boundaries(si) は挿入前の行番号なので、挿入後は boundaries(si) + (si-1) + 1
                    subtotalRow = boundaries(si) + si
                    If Not isFirst Then sumFormula = sumFormula & ","
                    sumFormula = sumFormula & ws.Cells(subtotalRow, colIdx).Address(False, True)
                    isFirst = False
                Next si
                sumFormula = sumFormula & ")"
                ws.Cells(newTotRow, colIdx).Formula = sumFormula
            End If
        Next c
    End If

    Exit Sub

正規化:
    s = Trim$(s)
    s = Replace(s, ",", "")
    Return
End Sub

Sub 重複キーを一覧にする()
    ' 依頼の語: 重複キー|同じ番号|列の重複|伝票番号を洗|列で重複|重複している値|重複している番号|重複を一覧|番号の一覧を右|何件あるか右|回以上出
    ' 扱う: 重複キー一覧 重複洗い出し 件数 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

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
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) = "" Then
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
        If cv = "合計" Or cv = "計" Or cv = "小計" Or cv = "総計" Then
            bodyLastRow = r - 1
        Else
            Exit For
        End If
    Next r

    ' 出力先列 = lastCol + 2
    outCol = lastCol + 2

    ' 既存の出力表を消去（outCol以降のhdrRow行に値があれば）
    Dim clrLast As Long
    clrLast = hdrRow
    For r = hdrRow To lastRow + 50
        If Trim(CStr(ws.Cells(r, outCol).Value)) <> "" Or _
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
        If keyVal = "" Then GoTo NextRow
        If IsNumeric(keyVal) Then
            normKey = CStr(CLng(keyVal))
        Else
            normKey = keyVal
        End If
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

End Sub

Sub コード列を名称に置き換える()
    ' 依頼の語: 列のコード|コード列|コードを名称|コードを名前|対応表の名前|コードのまま|名称に置|替えて名称
    ' 扱う: コード名称置き換え 突き合
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列 別のシート 共通の見出しのシート

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
    lastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    ' サンプル値収集（見出し行の次から）
    Dim sampleVals As Object
    Set sampleVals = CreateObject("Scripting.Dictionary")
    sampleVals.CompareMode = 1
    Dim sampleCount As Long
    sampleCount = 0
    For i = headerRow + 1 To lastRow
        If ws.Cells(i, selCol).Value <> "" Then
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
        If kv <> "" And vv <> "" Then
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
        If cellVal = "" Then GoTo NextCell
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
    If IsNumeric(tmp) Then
        s = CStr(CLng(tmp))
    Else
        s = tmp
    End If
    Return

End Sub

Sub 空欄行を一覧にする()
    ' 依頼の語: 空欄チェック|入力漏れ|記入漏れ|未記入|空欄がある行
    ' 扱う: 空欄チェック 入力漏れ 記入漏れ 未記入 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

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
    lastRow = ur.Row + ur.rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1

    ' 既に「空欄の行」見出しの列を探す（表の右側）
    outCol = 0
    Dim checkC As Long
    For checkC = lastCol To lastCol + 20
        If Trim$(CStr(ws.Cells(hdrRow, checkC).Value)) = "空欄の行" Then
            outCol = checkC
            Exit For
        End If
    Next checkC

    If outCol = 0 Then
        ' 表の右に1列空けて配置: lastCol+2
        ' ただし lastCol+1 列が既に使われている場合は lastCol+2 がずれないよう確認
        outCol = lastCol + 2
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
            If Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then
                isAllBlank = False
                Exit For
            End If
        Next c
        If isAllBlank Then GoTo NextRow

        ' 合計行チェック
        isTotalRow = False
        For c = ur.Column To lastCol
            cv = Trim$(CStr(ws.Cells(r, c).Value))
            If cv = "合計" Or cv = "計" Or cv = "小計" Or cv = "総計" Then
                isTotalRow = True
                Exit For
            End If
        Next c
        If isTotalRow Then GoTo NextRow

        cellVal = Trim$(CStr(ws.Cells(r, selCol).Value))
        If cellVal = "" Then
            ws.Cells(outRow, outCol).Value = CStr(r)
            outRow = outRow + 1
        End If

NextRow:
    Next r

End Sub

Sub 結合セルを解除して値を埋める()
    ' 依頼の語: 結合セルを解除|結合をはずして|結合を解いて|結合をぜんぶ解除|マージを解除|結合セルをばら|全部解除
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

Sub 空白セルを上の値で埋める()
    ' 依頼の語: 空白を上|空欄を上|上から埋め|値で埋|上と同|値の繰|セルの値
    ' 扱う: 空白埋め 上の値で埋める 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

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

    totalWords = Array("合計", "計", "小計", "総計")

    Set ws = ActiveSheet
    Set sel = Selection
    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.rows.Count - 1
    headerRow = sel.Cells(1, 1).Row
    startRow = headerRow + 1

    For Each area In sel.Areas
        For aCol = 1 To area.Columns.Count
            c = area.Columns(aCol).Column
            lastVal = Empty

            For r = startRow To lastRow
                rowEmpty = True
                For cc = ur.Column To ur.Column + ur.Columns.Count - 1
                    If ws.Cells(r, cc).Value <> "" Then
                        rowEmpty = False
                        Exit For
                    End If
                Next cc
                If rowEmpty Then GoTo NextRow

                isTotal = False
                For i = 0 To UBound(totalWords)
                    For cc = ur.Column To ur.Column + ur.Columns.Count - 1
                        If InStr(CStr(ws.Cells(r, cc).Value), totalWords(i)) > 0 Then
                            isTotal = True
                            Exit For
                        End If
                    Next cc
                    If isTotal Then Exit For
                Next i
                If isTotal Then GoTo NextRow

                cellVal = CStr(ws.Cells(r, c).Value)
                If cellVal = "" Then
                    If lastVal <> Empty And lastVal <> "" Then
                        ws.Cells(r, c).Value = lastVal
                    End If
                Else
                    lastVal = ws.Cells(r, c).Value
                End If

NextRow:
            Next r
        Next aCol
    Next area

End Sub

Sub 文字日付を日付値に直す()
    ' 依頼の語: 文字の日付を|受付日の書|入っている日付|日付列|日付が文字列|列を日付の値
    ' 扱う: 文字日付変換
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列
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
                If s = "" Then GoTo NextCell
                s = Replace(s, "年", "/")
                s = Replace(s, "月", "/")
                s = Replace(s, "日", "")
                s = Trim(s)
                On Error Resume Next
                parsed = CDate(s)
                parseErr = (Err.Number <> 0)
                Err.Clear
                On Error GoTo 0
                If Not parseErr Then
                    cell.Value = parsed
                    cell.NumberFormat = "yyyy/m/d"
                End If
NextCell:
            Next r
        Next col
    Next area
End Sub

Sub 文字数字を数値に直す()
    ' 依頼の語: 数値に直|数値に変換|金額が文字列|文字の数字を数値|計算できる数値|数に直|数値に変|文字を数値|カンマ付
    ' 扱う: 文字数字変換 数値変換
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 日付の列

    Dim ws As Worksheet
    Dim sel As Range
    Dim area As Range
    Dim colIdx As Long
    Dim r As Long
    Dim lastR As Long
    Dim headerRow As Long
    Dim cell As Range
    Dim s As String
    Dim chk As String
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
        colIdx = area.Column
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

            If s = "" Then GoTo NextCell
            If s = "-" Then GoTo NextCell

            chk = s
            If Left(chk, 1) = "-" Then chk = Mid(chk, 2)
            If Len(chk) >= 2 And Left(chk, 1) = "0" Then
                allDigit = True
                For ci = 1 To Len(chk)
                    If Mid(chk, ci, 1) < "0" Or Mid(chk, ci, 1) > "9" Then
                        allDigit = False: Exit For
                    End If
                Next ci
                If allDigit Then GoTo NextCell
            End If

            chk = s
            If Left(chk, 1) = "-" Then chk = Mid(chk, 2)
            If chk = "" Then GoTo NextCell
            allDigit = True
            For ci = 1 To Len(chk)
                ch = Mid(chk, ci, 1)
                If (ch < "0" Or ch > "9") And ch <> "." Then
                    allDigit = False: Exit For
                End If
            Next ci
            If Not allDigit Then GoTo NextCell

            v = CDbl(s)
            cell.Value = v
            cell.NumberFormat = "#,##0"

NextCell:
        Next r
    Next area

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    s = Replace(s, "円", "")
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
    If neg And s <> "" Then s = "-" & s
    Return

End Sub

Sub 空行を削除して詰める()
    ' 依頼の語: 空行を消|空白行を削除|空白の行を|明細の間の空|行の抜けを詰め|空行を取り除|空行が無ければ
    ' 扱う: 空行削除 空行詰め 合計
    ' 見出し: なし
    ' 形: 数の列 日付の列
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
    urBottom = ur.Row + ur.rows.Count - 1

    ' 見出し行を探す: 値が2つ以上あり、数値だけでない最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
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
            If ws.Cells(r, c).Value <> "" Then
                allEmpty = False
                Exit For
            End If
        Next c
        If allEmpty Then
            ws.rows(r).Delete Shift:=xlUp
        End If
    Next r
End Sub

Sub 二段見出しを畳む()
    ' 依頼の語: 2段の見出しを|二段になっている|見出しの上下2行|見出しを1行に畳|上段と下段|見出しを平
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

    Dim hRow As Long
    hRow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            Dim ma0 As Range
            Set ma0 = ws.Cells(r, c).MergeArea
            If ma0.Cells(1, 1).Row = r And ma0.Cells(1, 1).Column = c Then
                If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then cnt = cnt + 1
            End If
        Next c
        If cnt >= 2 Then
            hRow = r
            Exit For
        End If
    Next r

    If hRow = 0 Then Exit Sub
    If hRow + 1 > urBottom Then Exit Sub

    Dim row1 As Long, row2 As Long
    row1 = hRow
    row2 = hRow + 1

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
        If Trim(CStr(ws.Cells(row1, c).Value)) = "" Then
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
        Dim maC As Range
        Set maC = ws.Cells(row1, c).MergeArea
        Dim topVal As String
        topVal = Trim(CStr(maC.Cells(1, 1).Value))
        If topVal <> "" Then lastTop = topVal
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
        If t = "" And b = "" Then
            combined = ""
        ElseIf t = "" Then
            combined = b
        ElseIf b = "" Then
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

Sub 順位列を挿入する()
    ' 依頼の語: 順位|ランキング|何位|番付|RANK
    ' 扱う: 順位 ランキング 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    ' 選んでいる数の列のすぐ右に順位列を挿入し、RANK.EQ式を書く

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim rankCol As Long
    rankCol = selCol + 1

    Dim needInsert As Boolean
    needInsert = True
    If ws.Cells(hdrRow, rankCol).Value = "順位" Then
        needInsert = False
    End If

    If needInsert Then
        ws.Columns(rankCol).Insert Shift:=xlToRight
    End If

    ws.Cells(hdrRow, rankCol).Value = "順位"

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    Dim firstDataRow As Long
    firstDataRow = hdrRow + 1

    Dim colLetter As String
    colLetter = Split(ws.Cells(1, selCol).Address(True, False), "$")(0)

    ' 本文終端（合計行を除いた最後のデータ行）
    Dim endDataRow As Long
    endDataRow = lastRow
    Dim s As String
    Dim c2 As Long
    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = ur.Column
    lastCol = firstCol + ur.Columns.Count - 1

    ' 末尾から合計行をスキップして終端を決める
    Dim r As Long
    For r = lastRow To firstDataRow Step -1
        Dim isSum As Boolean
        isSum = False
        For c2 = firstCol To lastCol
            s = CStr(ws.Cells(r, c2).Value)
            GoSub 正規化
            If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then
                isSum = True
                Exit For
            End If
        Next c2
        If Not isSum Then
            endDataRow = r
            Exit For
        End If
    Next r

    Dim rangeAddr As String
    rangeAddr = "$" & colLetter & "$" & firstDataRow & ":$" & colLetter & "$" & endDataRow

    For r = firstDataRow To lastRow
        If ws.Cells(r, selCol).Value = "" Then GoTo NextRow

        isSum = False
        For c2 = firstCol To lastCol
            s = CStr(ws.Cells(r, c2).Value)
            GoSub 正規化
            If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then
                isSum = True
                Exit For
            End If
        Next c2
        If isSum Then GoTo NextRow

        Dim cellAddr As String
        cellAddr = "$" & colLetter & "$" & r
        ws.Cells(r, rankCol).Formula = "=RANK.EQ(" & cellAddr & "," & rangeAddr & ",0)"

NextRow:
    Next r

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    Return

End Sub

Sub 平均最大最小の表を右に作る()
    ' 依頼の語: 平均・最大・最小|平均と最大|最大値と最小値|最大最小平均|平均と最高|多い所
    ' 扱う: 平均 最大 最小 合計
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列

    Dim ws As Object
    Dim selCol1 As Long, selCol2 As Long
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim c As Long
    Dim outCol As Long
    Dim numHdr As String, itemHdr As String
    Dim numCol As Long, itemCol As Long
    Dim dataStart As Long, dataEnd As Long
    Dim rng1 As Object, rng2 As Object
    Dim r As Long
    Dim cellVal As Variant

    Set ws = ActiveSheet

    ' 選んでいる2列を取得
    If Selection.Areas.Count < 2 Then Exit Sub
    Set rng1 = Selection.Areas(1)
    Set rng2 = Selection.Areas(2)

    ' 見出し行は Selection.Cells(1,1).Row
    hdrRow = Selection.Cells(1, 1).Row

    ' 1つ目=項目列、2つ目=数列
    itemCol = rng1.Column
    numCol = rng2.Column

    ' 見出し
    itemHdr = CStr(ws.Cells(hdrRow, itemCol).Value)
    numHdr = CStr(ws.Cells(hdrRow, numCol).Value)

    ' 本文の最終行を求める（合計行を除く）
    ' UsedRangeの最終行
    Dim urLastRow As Long
    urLastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    dataStart = hdrRow + 1
    dataEnd = 0

    ' 数列で値がある最後の行を探す（合計行除外）
    Dim s As String
    For r = urLastRow To dataStart Step -1
        cellVal = ws.Cells(r, numCol).Value
        If cellVal <> "" Then
            ' 合計行チェック（項目列）
            s = CStr(ws.Cells(r, itemCol).Value)
            s = Trim$(StrConv(s, vbNarrow))
            If s = "合計" Or s = "計" Or s = "小計" Or s = "総計" Then
                ' skip
            Else
                dataEnd = r
                Exit For
            End If
        End If
    Next r

    If dataEnd < dataStart Then Exit Sub

    ' 表の右端列を求める（見出し行で左端から右へ空セルに当たる手前）
    Dim leftCol As Long
    leftCol = ws.UsedRange.Column
    Dim rightEndCol As Long
    rightEndCol = leftCol
    For c = leftCol To leftCol + ws.UsedRange.Columns.Count + 10
        If ws.Cells(hdrRow, c).Value = "" Then
            rightEndCol = c - 1
            Exit For
        End If
    Next c

    ' 出力列 = 右端の2つ右
    outCol = rightEndCol + 2

    ' 既に同じ表があれば消す（見出し行のoutCol="指標"かチェック）
    If CStr(ws.Cells(hdrRow, outCol).Value) = "指標" Then
        ws.Cells(hdrRow, outCol).Resize(4, 3).ClearContents
    End If

    ' 絶対参照の範囲文字列
    Dim numRng As String
    numRng = "$" & Split(ws.Cells(dataStart, numCol).Address(True, True), "$")(1) & _
             "$" & dataStart & ":$" & _
             Split(ws.Cells(dataEnd, numCol).Address(True, True), "$")(1) & _
             "$" & dataEnd

    Dim itemRng As String
    itemRng = "$" & Split(ws.Cells(dataStart, itemCol).Address(True, True), "$")(1) & _
              "$" & dataStart & ":$" & _
              Split(ws.Cells(dataEnd, itemCol).Address(True, True), "$")(1) & _
              "$" & dataEnd

    ' 見出し行
    ws.Cells(hdrRow, outCol).Value = "指標"
    ws.Cells(hdrRow, outCol + 1).Value = numHdr
    ws.Cells(hdrRow, outCol + 2).Value = itemHdr

    ' 平均行
    ws.Cells(hdrRow + 1, outCol).Value = "平均"
    ws.Cells(hdrRow + 1, outCol + 1).Formula = "=AVERAGE(" & numRng & ")"
    ws.Cells(hdrRow + 1, outCol + 1).NumberFormat = "#,##0"
    ws.Cells(hdrRow + 1, outCol + 2).Value = ""

    ' 最大行
    ws.Cells(hdrRow + 2, outCol).Value = "最大"
    ws.Cells(hdrRow + 2, outCol + 1).Formula = "=MAX(" & numRng & ")"
    ws.Cells(hdrRow + 2, outCol + 1).NumberFormat = "#,##0"
    ws.Cells(hdrRow + 2, outCol + 2).Formula = "=INDEX(" & itemRng & ",MATCH(MAX(" & numRng & ")," & numRng & ",0))"

    ' 最小行
    ws.Cells(hdrRow + 3, outCol).Value = "最小"
    ws.Cells(hdrRow + 3, outCol + 1).Formula = "=MIN(" & numRng & ")"
    ws.Cells(hdrRow + 3, outCol + 1).NumberFormat = "#,##0"
    ws.Cells(hdrRow + 3, outCol + 2).Formula = "=INDEX(" & itemRng & ",MATCH(MIN(" & numRng & ")," & numRng & ",0))"

End Sub

Sub 度数分布表を右に作る()
    ' 依頼の語: 度数分布|金額の区間|金額帯|階級に分けて|区間ごとの件数|金額の分布|価格帯別
    ' 扱う: 度数分布 区間件数 金額帯集計 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selCol As Long
    Dim hdrRow As Long
    selCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    ' 合計行を除いた本文最終行
    ' 合計行判定：その行のどこかのセルに「合計」「計」「小計」「総計」がある
    Dim dataLastRow As Long
    dataLastRow = hdrRow
    Dim r As Long
    For r = hdrRow + 1 To lastRow
        Dim isTotalRow As Boolean
        isTotalRow = False
        Dim cc2 As Long
        For cc2 = ur.Column To ur.Column + ur.Columns.Count - 1
            Dim cv2 As String
            cv2 = Trim(CStr(ws.Cells(r, cc2).Value))
            If cv2 = "合計" Or cv2 = "計" Or cv2 = "小計" Or cv2 = "総計" Then
                isTotalRow = True
                Exit For
            End If
        Next cc2
        If Not isTotalRow Then
            ' 数値セルがあれば本文行とみなす
            Dim hasNum As Boolean
            hasNum = False
            Dim cc3 As Long
            For cc3 = ur.Column To ur.Column + ur.Columns.Count - 1
                If IsNumeric(ws.Cells(r, cc3).Value) And CStr(ws.Cells(r, cc3).Value) <> "" Then
                    hasNum = True
                    Exit For
                End If
            Next cc3
            ' 空行でなければ更新
            Dim anyVal As Boolean
            anyVal = False
            For cc3 = ur.Column To ur.Column + ur.Columns.Count - 1
                If Trim(CStr(ws.Cells(r, cc3).Value)) <> "" Then
                    anyVal = True
                    Exit For
                End If
            Next cc3
            If anyVal Then dataLastRow = r
        End If
    Next r

    ' 最大値（合計行を除く本文から）
    Dim maxVal As Double
    maxVal = 0
    For r = hdrRow + 1 To dataLastRow
        ' この行が合計行か再チェック
        Dim isTR As Boolean
        isTR = False
        Dim cx As Long
        For cx = ur.Column To ur.Column + ur.Columns.Count - 1
            Dim cvx As String
            cvx = Trim(CStr(ws.Cells(r, cx).Value))
            If cvx = "合計" Or cvx = "計" Or cvx = "小計" Or cvx = "総計" Then
                isTR = True
                Exit For
            End If
        Next cx
        If Not isTR Then
            If IsNumeric(ws.Cells(r, selCol).Value) And CStr(ws.Cells(r, selCol).Value) <> "" Then
                Dim v As Double
                v = CDbl(ws.Cells(r, selCol).Value)
                If v > maxVal Then maxVal = v
            End If
        End If
    Next r

    ' 区間幅
    Dim rawWidth As Double
    rawWidth = maxVal / 10
    Dim magnitude As Double
    magnitude = 1
    If rawWidth > 0 Then
        Do While magnitude * 10 <= rawWidth
            magnitude = magnitude * 10
        Loop
    End If
    Dim stepWidth As Double
    If rawWidth <= magnitude Then
        stepWidth = magnitude
    ElseIf rawWidth <= magnitude * 2 Then
        stepWidth = magnitude * 2
    ElseIf rawWidth <= magnitude * 5 Then
        stepWidth = magnitude * 5
    Else
        stepWidth = magnitude * 10
    End If

    ' 区間数
    Dim numBins As Long
    numBins = 0
    Dim cur As Double
    cur = 0
    Do
        numBins = numBins + 1
        cur = cur + stepWidth
    Loop While cur <= maxVal

    ' 表の右端列（見出し行で左から空セルに当たる手前）
    Dim leftCol As Long
    leftCol = ur.Column
    Dim rightEnd As Long
    rightEnd = leftCol
    Dim c As Long
    For c = leftCol To leftCol + ur.Columns.Count + 10
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) = "" Then
            rightEnd = c - 1
            Exit For
        End If
    Next c

    Dim outCol As Long
    outCol = rightEnd + 2

    ' 既存の同じ表があればクリア
    Dim cc As Long
    For cc = outCol To outCol + 50
        If CStr(ws.Cells(hdrRow, cc).Value) = "以上" And _
           CStr(ws.Cells(hdrRow, cc + 1).Value) = "未満" And _
           CStr(ws.Cells(hdrRow, cc + 2).Value) = "件数" Then
            ws.Range(ws.Cells(hdrRow, cc), ws.Cells(hdrRow + numBins + 20, cc + 2)).ClearContents
            outCol = cc
            Exit For
        End If
    Next cc

    ' 見出し
    ws.Cells(hdrRow, outCol).Value = "以上"
    ws.Cells(hdrRow, outCol + 1).Value = "未満"
    ws.Cells(hdrRow, outCol + 2).Value = "件数"

    ' 本文範囲アドレス（絶対参照）
    Dim colLetter As String
    colLetter = Split(ws.Cells(1, selCol).Address(True, False), "$")(0)
    Dim dataRangeAddr As String
    dataRangeAddr = "$" & colLetter & "$" & (hdrRow + 1) & ":$" & colLetter & "$" & dataLastRow

    ' 区間行を書く
    Dim i As Long
    For i = 0 To numBins - 1
        Dim lowerVal As Double
        Dim upperVal As Double
        lowerVal = i * stepWidth
        upperVal = (i + 1) * stepWidth
        Dim outRow As Long
        outRow = hdrRow + 1 + i

        ws.Cells(outRow, outCol).Value = lowerVal
        ws.Cells(outRow, outCol).NumberFormat = "#,##0"
        ws.Cells(outRow, outCol + 1).Value = upperVal
        ws.Cells(outRow, outCol + 1).NumberFormat = "#,##0"

        Dim lowerAddr As String
        Dim upperAddr As String
        lowerAddr = ws.Cells(outRow, outCol).Address(True, True)
        upperAddr = ws.Cells(outRow, outCol + 1).Address(True, True)
        ws.Cells(outRow, outCol + 2).Formula = "=COUNTIFS(" & dataRangeAddr & ","">=""&" & lowerAddr & "," & dataRangeAddr & ",""<""&" & upperAddr & ")"
    Next i

End Sub

Sub 選んだ順に列を並べ替える()
    ' 依頼の語: 選んだ列を先頭に|並べ直|選んだ順|列を左|入れ替
    ' 扱う: 列並べ替え 列順変更 合計
    ' 見出し: なし
    ' 選ぶ列: 3
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selRow As Long
    selRow = Selection.Cells(1, 1).Row

    Dim leftCol As Long, rightCol As Long
    leftCol = ws.UsedRange.Column
    rightCol = leftCol
    Do While ws.Cells(selRow, rightCol + 1).Value <> ""
        rightCol = rightCol + 1
    Loop
    Dim nCols As Long
    nCols = rightCol - leftCol + 1

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    Dim totalRows As Long
    totalRows = lastRow - selRow + 1

    Dim nSel As Long
    nSel = Selection.Areas.Count
    Dim selCols() As Long
    ReDim selCols(1 To nSel)
    Dim i As Long
    For i = 1 To nSel
        selCols(i) = Selection.Areas(i).Cells(1, 1).Column
    Next i

    Dim alreadyOK As Boolean
    alreadyOK = True
    For i = 1 To nSel
        If selCols(i) <> leftCol + i - 1 Then
            alreadyOK = False
            Exit For
        End If
    Next i
    If alreadyOK Then Exit Sub

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
            bufVal(i, j) = cel.Value
            bufFmt(i, j) = cel.NumberFormat
            bufBold(i, j) = cel.Font.Bold
        Next i
    Next j

    For j = 1 To nCols
        Dim srcCol As Long
        srcCol = newOrder(j) - leftCol + 1
        For i = 1 To totalRows
            Set cel = ws.Cells(selRow + i - 1, leftCol + j - 1)
            cel.Value = bufVal(i, srcCol)
            cel.NumberFormat = bufFmt(i, srcCol)
            cel.Font.Bold = bufBold(i, srcCol)
        Next i
    Next j

End Sub

Sub 合計行を検算する()
    ' 依頼の語: 検算|合計の行を検算|縦計の検算|合計欄が正しいか|合計と内訳の和|足し算
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
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    ' 表の右端: 見出し行でいちばん左の列から右へ見て空のセルに当たる手前の列
    Dim tblRight As Long
    tblRight = urLeft - 1
    For c = urLeft To urRight
        If ws.Cells(hdrRow, c).Value = "" Then Exit For
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
        If IsNumeric(ws.Cells(totalRow, c).Value) And ws.Cells(totalRow, c).Value <> "" Then
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
    outCol = tblRight + 2

    ' 既に同じ表が右にあれば消す
    If ws.Cells(hdrRow, outCol).Value = "検算の列" Then
        Dim clearCol As Long
        For clearCol = outCol To outCol + 4
            ws.Columns(clearCol).ClearContents
        Next clearCol
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
        ws.Cells(dataRow, outCol + 2).Formula = "=SUM(" & sumAddr & ")"

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

End Sub

Sub 同日同額の二重計上を一覧にする()
    ' 依頼の語: 二重計上|二重払い|同日同額|同じ日付|同じ金額|同じ日に同じ金額|二重に支払|同じ組
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
    urBottom = ur.Row + ur.rows.Count - 1

    ' 見出し行を探す: 空でないセルが2つ以上並ぶ最初の行
    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    For r = urTop To urBottom
        Dim nonEmpty As Long
        nonEmpty = 0
        For c = urLeft To urRight
            If ws.Cells(r, c).Value <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    ' 表の右端列: 見出し行で左端から右へ見て空のセルに当たる手前の列
    Dim tblLeft As Long, tblRight As Long
    tblLeft = urLeft
    tblRight = tblLeft - 1
    For c = tblLeft To urRight
        If ws.Cells(hdrRow, c).Value <> "" Then
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
        dateCnt = 0
        For r = dataTop To urBottom
            Dim v As Variant
            v = ws.Cells(r, c).Value
            If IsDate(v) And v <> "" Then dateCnt = dateCnt + 1
        Next r
        If dateCnt >= totalRows * 0.8 Then
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
            If IsNumeric(v) And v <> "" And Not IsDate(v) Then numCnt = numCnt + 1
        Next r
        If numCnt >= totalRows * 0.8 Then
            amtCol = c
            Exit For
        End If
    Next c
    If amtCol = 0 Then Exit Sub

    ' 出力列: 表の右端の2つ右
    Dim outStartCol As Long
    outStartCol = tblRight + 2

    ' 既存の一覧を消す
    Dim lastUsedCol As Long
    lastUsedCol = ur.Column + ur.Columns.Count - 1
    If outStartCol <= lastUsedCol Then
        ws.Range(ws.Cells(hdrRow, outStartCol), ws.Cells(urBottom, lastUsedCol)).ClearContents
    End If

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
        If dv = "" Or av = "" Then GoTo NextRow
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
        ws.Cells(hdrRow, outStartCol + 1 + ci).Value = ws.Cells(hdrRow, tblLeft + ci).Value
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

End Sub

Sub 日付異常行を一覧にする()
    ' 依頼の語: おかしい日付|年度外の日付|日付の入力ミス|未来の日付や空|日付の妥当性|日付がおかしい|列の日付
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
    urLastRow = ur.Row + ur.rows.Count - 1

    ' 見出し行で左から空セルに当たる手前の列
    Dim tableLastCol As Long
    tableLastCol = 0
    Dim c As Long
    For c = 1 To urLastCol
        If ws.Cells(hdrRow, c).Value = "" Then
            tableLastCol = c - 1
            Exit For
        End If
    Next c
    If tableLastCol = 0 Then tableLastCol = urLastCol

    Dim outCol As Long
    outCol = tableLastCol + 2

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
            If ws.Cells(r, c).Value <> "" Then
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
        If Not isEmpty(dv) And dv <> "" And IsDate(dv) Then
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
            If ws.Cells(r, c).Value <> "" Then
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

        If isEmpty(cellVal) Or cellVal = "" Then
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

        If reason <> "" Then
            ws.Cells(outRow, outCol).Value = r
            ws.Cells(outRow, outCol + 1).Value = dispVal
            ws.Cells(outRow, outCol + 2).Value = reason
            outRow = outRow + 1
        End If
NextR2:
    Next r

End Sub

Sub 年齢と勤続年数を足す()
    ' 依頼の語: 年齢と勤続年数|生年月日から年齢|年齢・勤続|年齢と勤続|勤続年数を出して|基準日時点の年齢|職員の年齢|名簿に年齢|採用日から勤続
    ' 扱う: 年齢 勤続年数
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 選んでいる2列を取得
    Dim colDOB As Long, colHire As Long
    If Selection.Areas.Count < 2 Then Exit Sub
    colDOB = Selection.Areas(1).Column
    colHire = Selection.Areas(2).Column

    ' 見出し行 = Selection.Cells(1,1).Row
    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    ' 基準日セルを探す（見出し行より上で「基準日」で始まるセルのすぐ右）
    Dim baseRef As String
    baseRef = ""
    Dim r As Long, c As Long
    Dim ur As Range
    Set ur = ws.UsedRange
    Dim urC As Long, urR As Long
    urC = ur.Column
    urR = ur.Row

    Dim i As Long, j As Long
    For i = urR To hdrRow - 1
        For j = urC To urC + ur.Columns.Count - 1
            Dim cv As String
            cv = CStr(ws.Cells(i, j).Value)
            If Left(cv, 3) = "基準日" Then
                baseRef = ws.Cells(i, j + 1).Address(True, True)
                GoTo FoundBase
            End If
        Next j
    Next i
FoundBase:

    ' 表の右端列を探す（見出し行で左端から右へ空セルに当たる手前）
    Dim leftCol As Long
    leftCol = ur.Column
    Dim rightCol As Long
    rightCol = leftCol
    For j = leftCol To leftCol + ur.Columns.Count - 1
        If ws.Cells(hdrRow, j).Value = "" Then
            rightCol = j - 1
            Exit For
        End If
        rightCol = j
    Next j

    ' 年齢・勤続年数の列を決める（既存チェック）
    Dim colAge As Long, colSvc As Long
    colAge = 0
    colSvc = 0
    For j = leftCol To leftCol + ur.Columns.Count - 1
        Dim hv As String
        hv = CStr(ws.Cells(hdrRow, j).Value)
        If hv = "年齢" Then colAge = j
        If hv = "勤続年数" Then colSvc = j
    Next j
    If colAge = 0 Then colAge = rightCol + 1
    If colSvc = 0 Then colSvc = rightCol + 2
    If colAge = rightCol + 1 And colSvc = 0 Then colSvc = colAge + 1
    If colSvc = 0 Then colSvc = colAge + 1

    ' 見出しを書く
    ws.Cells(hdrRow, colAge).Value = "年齢"
    ws.Cells(hdrRow, colSvc).Value = "勤続年数"

    ' 本文の最終行を求める
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    ' 式を書く
    Dim basePart As String
    If baseRef = "" Then
        basePart = "TODAY()"
    Else
        basePart = baseRef
    End If

    For i = hdrRow + 1 To lastRow
        ' 行がまるごと空ならスキップ
        Dim rowEmpty As Boolean
        rowEmpty = True
        For j = leftCol To rightCol
            If ws.Cells(i, j).Value <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next j
        If rowEmpty Then
            ws.Cells(i, colAge).Value = ""
            ws.Cells(i, colSvc).Value = ""
        Else
            ' 年齢
            If ws.Cells(i, colDOB).Value = "" Then
                ws.Cells(i, colAge).Value = ""
            Else
                ws.Cells(i, colAge).Formula = "=DATEDIF(" & ws.Cells(i, colDOB).Address(False, False) & "," & basePart & ",""Y"")"
            End If
            ' 勤続年数
            If ws.Cells(i, colHire).Value = "" Then
                ws.Cells(i, colSvc).Value = ""
            Else
                ws.Cells(i, colSvc).Formula = "=DATEDIF(" & ws.Cells(i, colHire).Address(False, False) & "," & basePart & ",""Y"")"
            End If
        End If
    Next i

End Sub

Sub スパークライン折れ線を足す()
    ' 依頼の語: スパークライン|ミニグラフ|推移をスパーク|セル内
    ' 扱う: スパークライン 列を足 合計 グラフ
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

    Dim hdrRow As Long
    hdrRow = 0
    Dim r As Long, c As Long
    For r = urTop To urBottom
        Dim nonEmpty As Long
        nonEmpty = 0
        For c = urLeft To urRight
            If ws.Cells(r, c).Value <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    Dim tableRight As Long
    tableRight = urLeft - 1
    For c = urLeft To urRight
        If ws.Cells(hdrRow, c).Value = "" Then Exit For
        tableRight = c
    Next c
    If tableRight < urLeft Then Exit Sub

    Dim itemCol As Long
    itemCol = 0
    For c = urLeft To tableRight
        Dim allText As Boolean
        allText = True
        Dim hasVal As Boolean
        hasVal = False
        For r = hdrRow + 1 To urBottom
            If ws.Cells(r, c).Value <> "" Then
                hasVal = True
                If IsNumeric(ws.Cells(r, c).Value) And Not IsDate(ws.Cells(r, c).Value) Then
                    allText = False
                    Exit For
                End If
            End If
        Next r
        If allText And hasVal Then
            itemCol = c
            Exit For
        End If
    Next c
    If itemCol = 0 Then itemCol = urLeft

    Dim dataFirst As Long, dataLast As Long
    dataFirst = 0
    dataLast = 0
    For c = itemCol + 1 To tableRight
        Dim hdrVal As String
        hdrVal = Trim(CStr(ws.Cells(hdrRow, c).Value))
        Dim isTotCol As Boolean
        isTotCol = (hdrVal = "合計" Or hdrVal = "計" Or hdrVal = "小計" Or hdrVal = "総計")
        Dim hasNum As Boolean
        hasNum = False
        For r = hdrRow + 1 To urBottom
            Dim cv As Variant
            cv = ws.Cells(r, c).Value
            If IsNumeric(cv) And cv <> "" Then
                hasNum = True
                Exit For
            End If
        Next r
        If hasNum And Not isTotCol Then
            If dataFirst = 0 Then dataFirst = c
            dataLast = c
        End If
    Next c
    If dataFirst = 0 Or dataLast = 0 Then Exit Sub

    Dim sparkCol As Long
    sparkCol = 0
    For c = urLeft To urRight
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) = "推移" Then
            sparkCol = c
            Exit For
        End If
    Next c
    If sparkCol = 0 Then
        sparkCol = tableRight + 1
        ws.Cells(hdrRow, sparkCol).Value = "推移"
    Else
        For r = hdrRow + 1 To urBottom
            On Error Resume Next
            ws.Cells(r, sparkCol).SparklineGroups.Clear
            On Error GoTo 0
        Next r
    End If

    Dim totWords As Variant
    totWords = Array("合計", "計", "小計", "総計")

    For r = hdrRow + 1 To urBottom
        Dim rowEmpty As Boolean
        rowEmpty = True
        For c = urLeft To tableRight
            If ws.Cells(r, c).Value <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If rowEmpty Then GoTo NextRow

        Dim isTotRow As Boolean
        isTotRow = False
        Dim itemVal As String
        itemVal = Trim(CStr(ws.Cells(r, itemCol).Value))
        Dim ti As Variant
        For Each ti In totWords
            If itemVal = CStr(ti) Then
                isTotRow = True
                Exit For
            End If
        Next ti
        If isTotRow Then GoTo NextRow

        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(r, dataFirst), ws.Cells(r, dataLast))
        Dim destCell As Range
        Set destCell = ws.Cells(r, sparkCol)

        On Error Resume Next
        destCell.SparklineGroups.Clear
        On Error GoTo 0

        Dim sg As SparklineGroup
        Set sg = Nothing
        On Error Resume Next
        Set sg = destCell.SparklineGroups.Add(xlSparkLine, dataRange.Address)
        On Error GoTo 0

        If Not sg Is Nothing Then
            On Error Resume Next
            sg.Points.Highpoint.Visible = True
            sg.SeriesColor.Color = RGB(31, 78, 121)
            On Error GoTo 0
        End If
        Set sg = Nothing

NextRow:
    Next r

End Sub

Sub 二本棒グラフを作る()
    ' 依頼の語: 2本並び|並べた棒グラフ|前と後を並|予算と執行を2本|2列を並べた棒
    ' 扱う: 二本棒グラフ 集合縦棒 集合横棒 合計 塗り 比較
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列

    ' 選んでいる2列（前・後）から集合棒グラフを作る

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng1 As Range, rng2 As Range
    If Selection.Areas.Count >= 2 Then
        Set rng1 = Selection.Areas(1)
        Set rng2 = Selection.Areas(2)
    ElseIf Selection.Columns.Count >= 2 Then
        Set rng1 = Selection.Columns(1)
        Set rng2 = Selection.Columns(2)
    Else
        Exit Sub
    End If

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim hdr1 As String, hdr2 As String
    hdr1 = CStr(ws.Cells(hdrRow, rng1.Column).Value)
    hdr2 = CStr(ws.Cells(hdrRow, rng2.Column).Value)

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    Dim catCol As Long
    catCol = 0
    Dim c As Long
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        Dim testVal As Variant
        testVal = ws.Cells(hdrRow + 1, c).Value
        If Not IsNumeric(testVal) And Not IsDate(testVal) And CStr(testVal) <> "" Then
            catCol = c
            Exit For
        End If
    Next c
    If catCol = 0 Then catCol = ur.Column

    Dim dataRows() As Long
    Dim dataCount As Long
    dataCount = 0
    ReDim dataRows(1 To lastRow)

    Dim r As Long
    For r = hdrRow + 1 To lastRow
        Dim v1 As Variant, v2 As Variant
        v1 = ws.Cells(r, rng1.Column).Value
        v2 = ws.Cells(r, rng2.Column).Value
        If v1 = "" And v2 = "" Then GoTo NextRow
        Dim isTotal As Boolean
        isTotal = False
        Dim cc As Long
        For cc = ur.Column To ur.Column + ur.Columns.Count - 1
            Dim cellStr As String
            cellStr = CStr(ws.Cells(r, cc).Value)
            If cellStr = "合計" Or cellStr = "計" Or cellStr = "総計" Or cellStr = "小計" Then
                isTotal = True
                Exit For
            End If
        Next cc
        If isTotal Then GoTo NextRow
        If IsNumeric(v1) And IsNumeric(v2) And v1 <> "" And v2 <> "" Then
            dataCount = dataCount + 1
            dataRows(dataCount) = r
        End If
NextRow:
    Next r

    If dataCount = 0 Then Exit Sub

    Dim chartTitle As String
    chartTitle = ""
    Dim tr As Long
    For tr = ur.Row To hdrRow - 1
        Dim rowText As String
        rowText = ""
        Dim tc As Long
        For tc = ur.Column To ur.Column + ur.Columns.Count - 1
            rowText = rowText & CStr(ws.Cells(tr, tc).Value)
        Next tc
        rowText = Trim(rowText)
        If rowText <> "" And InStr(rowText, "単位") = 0 Then
            chartTitle = rowText
        End If
    Next tr
    If chartTitle = "" Then
        chartTitle = hdr1 & "と" & hdr2 & "の比較"
    End If

    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    Dim placeCol As Long
    placeCol = ur.Column + ur.Columns.Count
    Dim leftPos As Double
    leftPos = ws.Cells(1, placeCol).Left + 10
    Dim topPos As Double
    topPos = ws.Cells(ur.Row, 1).Top + 10

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(leftPos, topPos, 400, 300)

    Dim ch As Chart
    Set ch = cht.Chart

    Dim xlChartTypeVal As Long
    If dataCount <= 8 Then
        xlChartTypeVal = xlColumnClustered
    Else
        xlChartTypeVal = xlBarClustered
    End If
    ch.ChartType = xlChartTypeVal

    ch.SeriesCollection.NewSeries
    ch.SeriesCollection.NewSeries

    Dim cats() As String
    Dim vals1() As Double
    Dim vals2() As Double
    ReDim cats(1 To dataCount)
    ReDim vals1(1 To dataCount)
    ReDim vals2(1 To dataCount)
    Dim i As Long
    For i = 1 To dataCount
        cats(i) = CStr(ws.Cells(dataRows(i), catCol).Value)
        vals1(i) = CDbl(ws.Cells(dataRows(i), rng1.Column).Value)
        vals2(i) = CDbl(ws.Cells(dataRows(i), rng2.Column).Value)
    Next i

    Dim s1 As Series
    Set s1 = ch.SeriesCollection(1)
    s1.Name = hdr1
    s1.Values = vals1
    s1.XValues = cats
    s1.Interior.Color = RGB(166, 166, 166)

    Dim S2 As Series
    Set S2 = ch.SeriesCollection(2)
    S2.Name = hdr2
    S2.Values = vals2
    S2.XValues = cats
    S2.Interior.Color = RGB(31, 78, 121)

    If dataCount >= 9 Then
        ch.Axes(xlCategory).ReversePlotOrder = True
    End If

    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"

    ch.HasTitle = True
    ch.chartTitle.text = chartTitle
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Size = 14
    ch.chartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    Dim axVal As Axis
    Set axVal = ch.Axes(xlValue)
    axVal.TickLabels.NumberFormat = "#,##0"
    On Error Resume Next
    axVal.MajorGridlines.Delete
    On Error GoTo 0
    axVal.HasMajorGridlines = False

    ch.ChartGroups(1).GapWidth = 60

    Dim sAll As Series
    For Each sAll In ch.SeriesCollection
        sAll.HasDataLabels = False
    Next sAll

End Sub

Sub ウォーターフォールグラフを作る()
    ' 依頼の語: 滝グラフ|階段状|要因を積|説明するグラフ|増減の要因|増減の内訳|ウォーターフォール
    ' 扱う: ウォーターフォール 滝グラフ 増減グラフ 塗り
    ' 見出し: なし
    ' 選ぶ列: 2
    ' 形: 数の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Selection.Areas.Count < 2 Then Exit Sub
    Dim col1 As Long, col2 As Long
    col1 = Selection.Areas(1).Column
    col2 = Selection.Areas(2).Column
    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim ur As Range
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.rows.Count - 1

    Dim hdr1 As String, hdr2 As String
    hdr1 = CStr(ws.Cells(hdrRow, col1).Value)
    hdr2 = CStr(ws.Cells(hdrRow, col2).Value)

    ' 項目列を探す（選んだ2列以外で文字の列を探す）
    Dim itemCol As Long
    itemCol = 0
    Dim tc As Long
    For tc = ur.Column To ur.Column + ur.Columns.Count - 1
        If tc = col1 Or tc = col2 Then GoTo NextCol
        Dim testVal As Variant
        testVal = ws.Cells(hdrRow + 1, tc).Value
        If Not IsNumeric(testVal) And Not IsDate(testVal) And Trim(CStr(testVal)) <> "" Then
            itemCol = tc
            Exit For
        End If
NextCol:
    Next tc
    ' 見つからなければUsedRangeの左端から探す
    If itemCol = 0 Then
        For tc = ur.Column To ur.Column + ur.Columns.Count - 1
            testVal = ws.Cells(hdrRow + 1, tc).Value
            If Not IsNumeric(testVal) And Not IsDate(testVal) And Trim(CStr(testVal)) <> "" Then
                itemCol = tc
                Exit For
            End If
        Next tc
    End If
    If itemCol = 0 Then itemCol = ur.Column

    Dim titleStr As String
    titleStr = ""
    Dim tr As Long
    For tr = ur.Row To hdrRow - 1
        Dim rowVals As Long
        rowVals = 0
        Dim rv As String
        rv = ""
        Dim cc As Long
        For cc = ur.Column To ur.Column + ur.Columns.Count - 1
            Dim cv As String
            cv = Trim(CStr(ws.Cells(tr, cc).Value))
            If cv <> "" Then
                rowVals = rowVals + 1
                rv = cv
            End If
        Next cc
        If rowVals = 1 Then
            If InStr(rv, "単位") = 0 Then
                titleStr = rv
            End If
        End If
    Next tr
    If titleStr = "" Then
        titleStr = hdr1 & "から" & hdr2 & "への増減"
    End If

    Dim items() As String
    Dim vals1() As Double
    Dim vals2() As Double
    Dim n As Long
    n = 0
    ReDim items(0)
    ReDim vals1(0)
    ReDim vals2(0)

    Dim r As Long
    For r = hdrRow + 1 To lastRow
        Dim v1 As Variant, v2 As Variant
        v1 = ws.Cells(r, col1).Value
        v2 = ws.Cells(r, col2).Value
        If isEmpty(v1) Or isEmpty(v2) Then GoTo NextRow
        If Not IsNumeric(v1) Or Not IsNumeric(v2) Then GoTo NextRow
        If Trim(CStr(v1)) = "" Or Trim(CStr(v2)) = "" Then GoTo NextRow
        Dim isSumRow As Boolean
        isSumRow = False
        For cc = ur.Column To ur.Column + ur.Columns.Count - 1
            Dim cellStr As String
            cellStr = Trim(CStr(ws.Cells(r, cc).Value))
            If cellStr = "合計" Or cellStr = "計" Or cellStr = "総計" Or _
               InStr(cellStr, "合計") > 0 Or InStr(cellStr, "総計") > 0 Then
                isSumRow = True
                Exit For
            End If
        Next cc
        If isSumRow Then GoTo NextRow
        ReDim Preserve items(n)
        ReDim Preserve vals1(n)
        ReDim Preserve vals2(n)
        items(n) = Trim(CStr(ws.Cells(r, itemCol).Value))
        vals1(n) = CDbl(v1)
        vals2(n) = CDbl(v2)
        n = n + 1
NextRow:
    Next r

    If n = 0 Then Exit Sub

    Dim S0 As Double, S2 As Double
    S0 = 0: S2 = 0
    Dim i As Long
    For i = 0 To n - 1
        S0 = S0 + vals1(i)
        S2 = S2 + vals2(i)
    Next i

    Dim d() As Double
    ReDim d(n - 1)
    For i = 0 To n - 1
        d(i) = vals2(i) - vals1(i)
    Next i

    Dim s() As Double
    ReDim s(n - 1)
    Dim runS As Double
    runS = S0
    For i = 0 To n - 1
        s(i) = runS
        runS = runS + d(i)
    Next i

    Dim nPts As Long
    nPts = n + 2

    Dim base() As Double
    ReDim base(nPts - 1)
    base(0) = 0
    For i = 0 To n - 1
        If d(i) >= 0 Then
            base(i + 1) = s(i)
        Else
            base(i + 1) = s(i) + d(i)
        End If
    Next i
    base(nPts - 1) = 0

    Dim totSer() As Double
    ReDim totSer(nPts - 1)
    totSer(0) = S0
    For i = 1 To n
        totSer(i) = 0
    Next i
    totSer(nPts - 1) = S2

    Dim incSer() As Double
    ReDim incSer(nPts - 1)
    incSer(0) = 0
    For i = 0 To n - 1
        If d(i) > 0 Then incSer(i + 1) = d(i) Else incSer(i + 1) = 0
    Next i
    incSer(nPts - 1) = 0

    Dim decSer() As Double
    ReDim decSer(nPts - 1)
    decSer(0) = 0
    For i = 0 To n - 1
        If d(i) < 0 Then decSer(i + 1) = -d(i) Else decSer(i + 1) = 0
    Next i
    decSer(nPts - 1) = 0

    Dim xv() As String
    ReDim xv(nPts - 1)
    xv(0) = hdr1
    For i = 0 To n - 1
        xv(i + 1) = items(i)
    Next i
    xv(nPts - 1) = hdr2

    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    Dim urRight As Double
    urRight = ws.Cells(ur.Row, ur.Column + ur.Columns.Count - 1).Left + _
              ws.Cells(ur.Row, ur.Column + ur.Columns.Count - 1).Width + 10

    Dim coNew As ChartObject
    Set coNew = ws.ChartObjects.Add(urRight, ws.Cells(ur.Row, ur.Column).Top, 480, 300)

    Dim ch As Chart
    Set ch = coNew.Chart
    ch.ChartType = xlColumnStacked

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim sr1 As Series
    Set sr1 = ch.SeriesCollection.NewSeries
    sr1.Name = "土台"
    sr1.Values = base
    sr1.XValues = xv

    Dim sr2 As Series
    Set sr2 = ch.SeriesCollection.NewSeries
    sr2.Name = "合計"
    sr2.Values = totSer

    Dim sr3 As Series
    Set sr3 = ch.SeriesCollection.NewSeries
    sr3.Name = "増加"
    sr3.Values = incSer

    Dim sr4 As Series
    Set sr4 = ch.SeriesCollection.NewSeries
    sr4.Name = "減少"
    sr4.Values = decSer

    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"

    ch.HasTitle = True
    ch.chartTitle.text = titleStr
    With ch.chartTitle.Format.TextFrame2.TextRange.Font
        .Name = "Meiryo UI"
        .Size = 14
        .Bold = msoTrue
    End With

    ch.HasLegend = False

    Dim ax As Axis
    Set ax = ch.Axes(xlValue)
    ax.TickLabels.NumberFormat = "#,##0"
    ax.HasMajorGridlines = False
    ax.HasMinorGridlines = False

    ch.ChartGroups(1).GapWidth = 50

    sr1.Format.Fill.Visible = msoFalse
    sr1.Border.LineStyle = xlNone

    sr2.Format.Fill.Visible = msoTrue
    sr2.Format.Fill.ForeColor.RGB = RGB(127, 127, 127)

    sr3.Format.Fill.Visible = msoTrue
    sr3.Format.Fill.ForeColor.RGB = RGB(31, 78, 121)

    sr4.Format.Fill.Visible = msoTrue
    sr4.Format.Fill.ForeColor.RGB = RGB(192, 0, 0)

    sr1.HasDataLabels = False
    sr2.HasDataLabels = False
    sr3.HasDataLabels = False
    sr4.HasDataLabels = False

End Sub

Sub グラフの見た目をそろえる()
    ' 依頼の語: グラフの見せ方|グラフの見|グラフを全部|既存のグラフ|グラフのフォント|グラフの体裁|作ってあるグラフ|グラフを全部同じ
    ' 扱う: グラフ体裁 グラフフォント グラフ凡例
    ' 見出し: なし
    ' 形: 数の列 グラフ
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim cht As Chart
    Dim serCount As Integer
    Dim chtType As Long
    Dim isPie As Boolean
    Dim fmtStr As String
    Dim valAx As Object

    Set ws = ActiveSheet

    If ws.ChartObjects.Count = 0 Then Exit Sub

    For Each co In ws.ChartObjects
        Set cht = co.Chart

        cht.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"

        If cht.HasTitle Then
            cht.chartTitle.Format.TextFrame2.TextRange.Font.Name = "Meiryo UI"
            cht.chartTitle.Format.TextFrame2.TextRange.Font.Size = 14
            cht.chartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        End If

        serCount = cht.SeriesCollection.Count

        isPie = False
        If serCount > 0 Then
            chtType = cht.ChartType
            If chtType = xlPie Or chtType = xlPieExploded Or _
               chtType = xl3DPie Or chtType = xl3DPieExploded Or _
               chtType = xlDoughnut Or chtType = xlDoughnutExploded Then
                isPie = True
            End If
        End If

        If isPie Then
            cht.HasLegend = True
            cht.Legend.Position = xlLegendPositionRight
        ElseIf serCount <= 1 Then
            cht.HasLegend = False
        Else
            cht.HasLegend = True
            cht.Legend.Position = xlLegendPositionBottom
        End If

        If Not isPie Then
            On Error Resume Next
            Set valAx = Nothing
            Set valAx = cht.Axes(xlValue)
            On Error GoTo 0
            If Not valAx Is Nothing Then
                fmtStr = valAx.TickLabels.NumberFormat
                If InStr(fmtStr, "%") = 0 Then
                    valAx.TickLabels.NumberFormat = "#,##0"
                Else
                    Dim baseFmt As String
                    baseFmt = fmtStr
                    If InStr(baseFmt, ".") = 0 Then
                        baseFmt = "0.0%"
                    End If
                    valAx.TickLabels.NumberFormat = baseFmt
                End If
                On Error Resume Next
                cht.SetElement msoElementPrimaryValueGridLinesNone
                On Error GoTo 0
            End If
        End If
    Next co
End Sub

Sub 一覧からひな形を差し込む()
    ' 依頼の語: 差し込|件1枚|1件1枚|通知書を作|流し込|ひな形に|帳票を作|人ごとのシート|行ずつシート
    ' 扱う: 差し込み 帳票 ひな形 一覧シート 合計 印刷
    ' 見出し: なし
    ' 形: 数の列 日付の列 別のシート

    Dim wb As Workbook
    Dim listWs As Worksheet
    Dim tmplWs As Worksheet
    Dim ws As Worksheet
    Dim ur As Range
    Dim hdrRow As Long
    Dim c As Long, r As Long
    Dim lastCol As Long, lastRow As Long
    Dim nonEmpty As Long

    Set wb = ActiveWorkbook
    Set listWs = ActiveSheet

    ' ひな形シートを探す（アクティブシート以外で {見出し} 形式のプレースホルダを含むシート）
    Set tmplWs = Nothing
    Dim cel As Range
    Dim found As Boolean
    For Each ws In wb.Worksheets
        If ws.Name <> listWs.Name Then
            found = False
            For Each cel In ws.UsedRange
                If CStr(cel.Value) Like "*{*}*" Then
                    found = True
                    Exit For
                End If
            Next cel
            If found Then
                Set tmplWs = ws
                Exit For
            End If
        End If
    Next ws

    ' ひな形が見つからない場合、一覧シート以外の全シートで {…} を再探索（ひな形シート自体がアクティブな場合も考慮）
    If tmplWs Is Nothing Then
        ' アクティブシートがひな形の可能性があるので、一覧シートを別途特定する
        ' → この依頼ではアクティブ=一覧なので終了
        Exit Sub
    End If

    ' 見出し行を探す
    Set ur = listWs.UsedRange
    hdrRow = 0
    Dim firstRow As Long, firstCol As Long
    firstRow = ur.Row
    firstCol = ur.Column
    lastRow = ur.Row + ur.rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1

    For r = firstRow To lastRow
        nonEmpty = 0
        For c = firstCol To lastCol
            If listWs.Cells(r, c).Value <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    Dim headers() As String
    ReDim headers(firstCol To lastCol)
    For c = firstCol To lastCol
        headers(c) = CStr(listWs.Cells(hdrRow, c).Value)
    Next c

    Dim sheetName As String
    Dim newWs As Worksheet
    Dim skipRow As Boolean
    Dim existWs As Worksheet

    For r = hdrRow + 1 To lastRow
        nonEmpty = 0
        For c = firstCol To lastCol
            If listWs.Cells(r, c).Value <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty = 0 Then GoTo NextRow

        skipRow = False
        Dim cv As String
        For c = firstCol To lastCol
            cv = CStr(listWs.Cells(r, c).Value)
            If cv = "合計" Or cv = "計" Or cv = "総計" Then
                skipRow = True
                Exit For
            End If
        Next c
        If skipRow Then GoTo NextRow

        sheetName = CStr(listWs.Cells(r, firstCol).Value)
        If sheetName = "" Then GoTo NextRow

        On Error Resume Next
        Set existWs = Nothing
        Set existWs = wb.Worksheets(sheetName)
        On Error GoTo 0
        If Not existWs Is Nothing Then
            Application.DisplayAlerts = False
            existWs.Delete
            Application.DisplayAlerts = True
            Set existWs = Nothing
        End If

        tmplWs.Copy After:=wb.Sheets(wb.Sheets.Count)
        Set newWs = wb.Sheets(wb.Sheets.Count)
        On Error Resume Next
        newWs.Name = sheetName
        On Error GoTo 0

        Dim newCel As Range
        For Each newCel In newWs.UsedRange
            Dim cellVal As String
            cellVal = CStr(newCel.Value)
            If cellVal = "" Then GoTo NextCell
            If Not (cellVal Like "*{*}*") Then GoTo NextCell

            Dim exactMatch As Boolean
            exactMatch = False
            For c = firstCol To lastCol
                If cellVal = "{" & headers(c) & "}" Then
                    newCel.Value = listWs.Cells(r, c).Value
                    exactMatch = True
                    Exit For
                End If
            Next c

            If Not exactMatch Then
                Dim newVal As String
                newVal = cellVal
                Dim changed As Boolean
                changed = False
                For c = firstCol To lastCol
                    Dim placeholder As String
                    placeholder = "{" & headers(c) & "}"
                    If InStr(newVal, placeholder) > 0 Then
                        newVal = Replace(newVal, placeholder, listWs.Cells(r, c).text)
                        changed = True
                    End If
                Next c
                If changed Then newCel.Value = newVal
            End If
NextCell:
        Next newCel

NextRow:
    Next r

End Sub

Sub 印刷設定をする()
    ' 依頼の語: 印刷設定|印刷できるように|印刷の設定|印刷の体裁|横1ページに収|見出しを毎ページ|横幅を1ページ|ページ番号も
    ' 扱う: 印刷設定 印刷体裁
    ' 見出し: なし
    ' 形: 数の列

    Dim ws As Worksheet
    Dim ur As Range
    Dim hdrRow As Long, firstCol As Long, lastCol As Long, lastRow As Long
    Dim i As Long, j As Long, cnt As Long
    Dim numCols As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    Dim urFirstRow As Long, urFirstCol As Long, urLastRow As Long, urLastCol As Long
    urFirstRow = ur.Row
    urFirstCol = ur.Column
    urLastRow = ur.Row + ur.rows.Count - 1
    urLastCol = ur.Column + ur.Columns.Count - 1

    hdrRow = 0
    For i = urFirstRow To urLastRow
        cnt = 0
        For j = urFirstCol To urLastCol
            If ws.Cells(i, j).Value <> "" Then cnt = cnt + 1
        Next j
        If cnt >= 2 Then
            hdrRow = i
            Exit For
        End If
    Next i
    If hdrRow = 0 Then Exit Sub

    firstCol = 0
    For j = urFirstCol To urLastCol
        If ws.Cells(hdrRow, j).Value <> "" Then
            firstCol = j
            Exit For
        End If
    Next j
    If firstCol = 0 Then Exit Sub

    lastCol = firstCol
    For j = firstCol To urLastCol
        If ws.Cells(hdrRow, j).Value <> "" Then
            lastCol = j
        Else
            Exit For
        End If
    Next j

    lastRow = hdrRow
    Dim hasVal As Boolean
    For i = urLastRow To hdrRow + 1 Step -1
        hasVal = False
        For j = firstCol To lastCol
            If ws.Cells(i, j).Value <> "" Then
                hasVal = True
                Exit For
            End If
        Next j
        If hasVal Then
            lastRow = i
            Exit For
        End If
    Next i

    numCols = lastCol - firstCol + 1

    Dim printRange As String
    printRange = ws.Cells(1, firstCol).Address(False, False) & ":" & ws.Cells(lastRow, lastCol).Address(False, False)

    With ws.PageSetup
        .PaperSize = xlPaperA4
        If numCols >= 6 Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ws.rows(hdrRow).Address(False, False)
        .PrintArea = printRange
        .CenterHorizontally = True
        .CenterVertically = False
        .CenterFooter = "&P / &N"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
    End With
End Sub

Sub 様式に転記する()
    ' 依頼の語: 様式に転記|様式に写|様式シートへ転記|様式に書き込|様式に数字|管理会計の様式|報告様式
    ' 扱う: 様式転記 合計 集計
    ' 見出し: なし
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim ur As Range
    Dim hdrRow As Long, dataStart As Long, dataEnd As Long
    Dim c As Long, r As Long
    Dim itemCol As Long
    Dim i As Long, j As Long
    Dim numCols() As Long
    Dim numCount As Long
    Dim s As String

    Set wsData = ActiveSheet

    ' 様式シートを探す: アクティブ以外で集計表の項目名が2つ以上含まれるシート
    ' まず集計表の情報を取得
    Set ur = wsData.UsedRange

    ' 見出し行を探す: 空でないセルが2つ以上並ぶ最初の行
    hdrRow = 0
    Dim urTop As Long, urLeft As Long, urRows As Long, urCols As Long
    urTop = ur.Row
    urLeft = ur.Column
    urRows = ur.rows.Count
    urCols = ur.Columns.Count

    Dim nonEmpty As Long
    For r = urTop To urTop + urRows - 1
        nonEmpty = 0
        For c = urLeft To urLeft + urCols - 1
            If wsData.Cells(r, c).Value <> "" Then nonEmpty = nonEmpty + 1
        Next c
        If nonEmpty >= 2 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Sub

    dataStart = hdrRow + 1
    dataEnd = urTop + urRows - 1

    ' 項目列 = 左端の文字の列
    itemCol = 0
    For c = urLeft To urLeft + urCols - 1
        Dim hasText As Boolean
        hasText = False
        For r = dataStart To dataEnd
            If wsData.Cells(r, c).Value <> "" Then
                If Not IsNumeric(wsData.Cells(r, c).Value) Then
                    hasText = True
                    Exit For
                End If
            End If
        Next r
        If hasText Then
            itemCol = c
            Exit For
        End If
    Next c
    If itemCol = 0 Then Exit Sub

    ' 数の列を探す
    numCount = 0
    ReDim numCols(1 To urCols)
    For c = urLeft To urLeft + urCols - 1
        If c = itemCol Then GoTo NextCol
        If wsData.Cells(hdrRow, c).Value = "" Then GoTo NextCol
        Dim hasNum As Boolean
        hasNum = False
        For r = dataStart To dataEnd
            If IsNumeric(wsData.Cells(r, c).Value) And wsData.Cells(r, c).Value <> "" Then
                hasNum = True
                Exit For
            End If
        Next r
        If hasNum Then
            numCount = numCount + 1
            numCols(numCount) = c
        End If
NextCol:
    Next c

    ' 集計表の項目リストを作成
    Dim dictItems As Object
    Set dictItems = CreateObject("Scripting.Dictionary")
    For r = dataStart To dataEnd
        ' 行が全て空かチェック
        Dim allEmpty As Boolean
        allEmpty = True
        For c = urLeft To urLeft + urCols - 1
            If wsData.Cells(r, c).Value <> "" Then allEmpty = False: Exit For
        Next c
        If allEmpty Then GoTo NextDataRow

        ' 合計行チェック
        Dim isSumRow As Boolean
        isSumRow = False
        For c = urLeft To urLeft + urCols - 1
            s = CStr(wsData.Cells(r, c).Value): GoSub 正規化
            If s = "合計" Or s = "計" Or s = "総計" Then isSumRow = True: Exit For
        Next c
        If isSumRow Then GoTo NextDataRow

        s = CStr(wsData.Cells(r, itemCol).Value): GoSub 正規化
        If s <> "" Then dictItems(s) = r
NextDataRow:
    Next r

    ' 様式シートを探す
    Set wsForm = Nothing
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If sh.Name = wsData.Name Then GoTo NextSheet
        ' 集計表の項目名が2つ以上含まれているか確認
        Dim matchCount As Long
        matchCount = 0
        Dim shUr As Range
        Set shUr = sh.UsedRange
        Dim key As Variant
        For Each key In dictItems.keys
            Dim fr As Long, fc As Long
            For fr = shUr.Row To shUr.Row + shUr.rows.Count - 1
                For fc = shUr.Column To shUr.Column + shUr.Columns.Count - 1
                    s = CStr(sh.Cells(fr, fc).Value): GoSub 正規化
                    If s = key Then matchCount = matchCount + 1: GoTo NextKey
                Next fc
            Next fr
NextKey:
        Next key
        If matchCount >= 2 Then
            Set wsForm = sh
            Exit For
        End If
NextSheet:
    Next sh
    If wsForm Is Nothing Then Exit Sub

    ' 様式の見出し行を探す: 集計表の数の列の見出しが入っている行
    Dim shUr2 As Range
    Set shUr2 = wsForm.UsedRange
    Dim fHdrRow As Long
    fHdrRow = 0

    ' 数の列の見出しリスト
    Dim numHeaders() As String
    ReDim numHeaders(1 To numCount)
    For i = 1 To numCount
        numHeaders(i) = CStr(wsData.Cells(hdrRow, numCols(i)).Value)
        s = numHeaders(i): GoSub 正規化: numHeaders(i) = s
    Next i

    Dim fTop As Long, fLeft As Long, fRows As Long, fCols As Long
    fTop = shUr2.Row
    fLeft = shUr2.Column
    fRows = shUr2.rows.Count
    fCols = shUr2.Columns.Count

    For r = fTop To fTop + fRows - 1
        For i = 1 To numCount
            For c = fLeft To fLeft + fCols - 1
                s = CStr(wsForm.Cells(r, c).Value): GoSub 正規化
                If s = numHeaders(i) Then
                    fHdrRow = r
                    GoTo FoundFHdr
                End If
            Next c
        Next i
    Next r
FoundFHdr:
    If fHdrRow = 0 Then Exit Sub

    ' 様式の数の列を見出しから探す（見出しが一致する列）
    Dim fNumCols() As Long
    ReDim fNumCols(1 To numCount)
    For i = 1 To numCount
        fNumCols(i) = 0
        For c = fLeft To fLeft + fCols - 1
            s = CStr(wsForm.Cells(fHdrRow, c).Value): GoSub 正規化
            If s = numHeaders(i) Then
                fNumCols(i) = c
                Exit For
            End If
        Next c
    Next i

    ' 様式の項目列を探す（見出し行の左側で文字が並ぶ列）
    ' 様式の各行を走査して項目名を探す
    ' 様式の行ごとに項目名があるセルを探す（全セルを走査）
    ' まず様式の各行の全セルから項目名と一致するものを探す

    ' 転記
    Dim fDataStart As Long, fDataEnd As Long
    fDataStart = fHdrRow + 1
    fDataEnd = fTop + fRows - 1

    For r = fDataStart To fDataEnd
        ' 様式の行から項目名を探す（全列）
        For c = fLeft To fLeft + fCols - 1
            ' 数の列見出しと一致する列はスキップ
            Dim isNumHdrCol As Boolean
            isNumHdrCol = False
            For i = 1 To numCount
                If fNumCols(i) = c Then isNumHdrCol = True: Exit For
            Next i
            If isNumHdrCol Then GoTo NextFCol

            s = CStr(wsForm.Cells(r, c).Value): GoSub 正規化
            If s = "" Then GoTo NextFCol
            If Not dictItems.Exists(s) Then GoTo NextFCol

            ' 一致した: 集計表の対応行を取得
            Dim srcRow As Long
            srcRow = dictItems(s)

            ' 各数の列を転記
            For i = 1 To numCount
                If fNumCols(i) = 0 Then GoTo NextNumCol
                ' 式が入っているセルには書かない
                If wsForm.Cells(r, fNumCols(i)).HasFormula Then GoTo NextNumCol
                wsForm.Cells(r, fNumCols(i)).Value = wsData.Cells(srcRow, numCols(i)).Value
NextNumCol:
            Next i
            Exit For
NextFCol:
        Next c
    Next r

    Exit Sub

正規化:
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(s, " ", "")
    s = Replace(s, Chr(12288), "")
    Return

End Sub

Sub 列の値でシートに分割する()
    ' 依頼の語: シートに分けて|シートを分割|シートに振り分け|列でシート分け|シートを分|シートへ分|シートを値
    ' 扱う: シート分割 シート分け 値ごとに分割 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Dim selCol As Long
    Dim hdrRow As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim keyVal As String
    Dim keyTrim As String
    Dim dict As Object
    Dim keys() As String
    Dim keyCount As Long
    Dim destWs As Worksheet
    Dim destRow As Long
    Dim wb As Workbook
    Dim isBlankRow As Boolean
    Dim cellVal As String
    Dim isSumRow As Boolean
    Dim fmt As String
    Dim v As Variant

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    ' 選んでいる列と見出し行
    selCol = Selection.Cells(1, 1).Column
    hdrRow = Selection.Cells(1, 1).Row

    ' 表の列範囲: 見出し行の左端から右へ空セルに当たる手前まで
    ' 左端は UsedRange の左端
    Dim urLeft As Long
    urLeft = ws.UsedRange.Column
    lastCol = urLeft - 1
    For c = urLeft To urLeft + ws.UsedRange.Columns.Count - 1
        If Trim(CStr(ws.Cells(hdrRow, c).Value)) = "" Then
            Exit For
        End If
        lastCol = c
    Next c

    ' 本文の最終行
    lastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

    ' キーを出た順に収集
    Set dict = CreateObject("Scripting.Dictionary")
    keyCount = 0
    ReDim keys(0)

    For r = hdrRow + 1 To lastRow
        ' 行全体が空か
        isBlankRow = True
        For c = urLeft To lastCol
            If ws.Cells(r, c).Value <> "" Then
                isBlankRow = False
                Exit For
            End If
        Next c
        If isBlankRow Then GoTo NextRow

        ' 合計行か
        isSumRow = False
        For c = urLeft To lastCol
            cellVal = Trim(CStr(ws.Cells(r, c).Value))
            If cellVal = "合計" Or cellVal = "計" Or cellVal = "小計" Or cellVal = "総計" Then
                isSumRow = True
                Exit For
            End If
        Next c
        If isSumRow Then GoTo NextRow

        keyVal = CStr(ws.Cells(r, selCol).Value)
        keyTrim = Trim(keyVal)
        If keyTrim = "" Then GoTo NextRow
        If Not dict.Exists(keyTrim) Then
            dict.Add keyTrim, 1
            ReDim Preserve keys(keyCount)
            keys(keyCount) = keyTrim
            keyCount = keyCount + 1
        End If
NextRow:
    Next r

    ' 各キーのシートを作成
    Dim ki As Long
    For ki = 0 To keyCount - 1
        keyTrim = keys(ki)

        ' 既存シートを削除
        Dim existWs As Worksheet
        Set existWs = Nothing
        On Error Resume Next
        Set existWs = wb.Worksheets(keyTrim)
        On Error GoTo 0
        If Not existWs Is Nothing Then
            Application.DisplayAlerts = False
            existWs.Delete
            Application.DisplayAlerts = True
        End If

        ' 新シートを末尾に追加
        Set destWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        destWs.Name = keyTrim

        ' 見出し行をコピー
        For c = urLeft To lastCol
            destWs.Cells(1, c - urLeft + 1).Value = ws.Cells(hdrRow, c).Value
            destWs.Cells(1, c - urLeft + 1).NumberFormat = ws.Cells(hdrRow, c).NumberFormat
        Next c

        ' 本文をコピー
        destRow = 2
        For r = hdrRow + 1 To lastRow
            isBlankRow = True
            For c = urLeft To lastCol
                If ws.Cells(r, c).Value <> "" Then
                    isBlankRow = False
                    Exit For
                End If
            Next c
            If isBlankRow Then GoTo NextRow2

            isSumRow = False
            For c = urLeft To lastCol
                cellVal = Trim(CStr(ws.Cells(r, c).Value))
                If cellVal = "合計" Or cellVal = "計" Or cellVal = "小計" Or cellVal = "総計" Then
                    isSumRow = True
                    Exit For
                End If
            Next c
            If isSumRow Then GoTo NextRow2

            If Trim(CStr(ws.Cells(r, selCol).Value)) = keyTrim Then
                For c = urLeft To lastCol
                    destWs.Cells(destRow, c - urLeft + 1).Value = ws.Cells(r, c).Value
                    destWs.Cells(destRow, c - urLeft + 1).NumberFormat = ws.Cells(r, c).NumberFormat
                Next c
                destRow = destRow + 1
            End If
NextRow2:
        Next r
    Next ki

End Sub

Sub ピボットを値の表に変換する()
    ' 依頼の語: ピボットを値貼り|ピボットを値で|静的な表|値で貼|集計の結果|ピボットを値
    ' 扱う: ピボット変換 ピボット値貼り フィルタ 合計 ピボット
    ' 見出し: なし
    ' 形: 数の列 ピボット

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim arr As Variant
    Dim ptList() As String
    Dim ptCount As Integer

    Set ws = ActiveSheet

    If ws.PivotTables.Count = 0 Then Exit Sub

    ptCount = ws.PivotTables.Count
    ReDim ptList(1 To ptCount)
    Dim i As Integer
    For i = 1 To ptCount
        ptList(i) = ws.PivotTables(i).Name
    Next i

    For i = 1 To ptCount
        Dim pt2 As PivotTable
        On Error Resume Next
        Set pt2 = ws.PivotTables(ptList(i))
        On Error GoTo 0
        If pt2 Is Nothing Then GoTo NextPivot

        Set rng = pt2.TableRange2

        arr = rng.Value

        rng.Clear

        Dim destCell As Range
        Set destCell = ws.Cells(rng.Row, rng.Column)
        destCell.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

        Set pt2 = Nothing
NextPivot:
    Next i

End Sub

Sub 列をクエリで分割する()
    ' 依頼の語: 列を区切|列を分|課と係|姓と名
    ' 扱う: パワークエリ列分割
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 テーブル

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim selCol As Long
    Dim colName As String
    Dim tblName As String
    Dim qryName As String
    Dim newSheetName As String
    Dim newTblName As String
    Dim i As Long, r As Long
    Dim cellVal As String
    Dim delimiters(5) As String
    Dim counts(5) As Long
    Dim bestDelim As String
    Dim bestCount As Long
    Dim maxParts As Long
    Dim parts() As String
    Dim partCount As Long
    Dim wb As Workbook
    Dim newWs As Worksheet
    Dim existWs As Worksheet
    Dim mFormula As String
    Dim colNames As String
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim selArea As Range

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    Set selArea = Selection

    ' テーブル取得
    On Error Resume Next
    Set lo = selArea.Cells(1, 1).ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    tblName = lo.Name
    selCol = selArea.Cells(1, 1).Column - lo.Range.Column + 1
    colName = lo.HeaderRowRange.Cells(1, selCol).Value

    ' 区切り文字候補
    delimiters(0) = "・"
    delimiters(1) = "／"
    delimiters(2) = "/"
    delimiters(3) = "、"
    delimiters(4) = "-"
    delimiters(5) = " "

    ' データ範囲
    Dim dataRange As Range
    Set dataRange = lo.DataBodyRange
    If dataRange Is Nothing Then Exit Sub

    ' 各区切り文字の出現行数カウント
    Dim d As Long
    For d = 0 To 5
        counts(d) = 0
        For r = 1 To dataRange.rows.Count
            cellVal = CStr(dataRange.Cells(r, selCol).Value)
            If InStr(cellVal, delimiters(d)) > 0 Then
                counts(d) = counts(d) + 1
            End If
        Next r
    Next d

    ' 最多区切り文字を選ぶ（同数なら先のもの）
    bestDelim = delimiters(0)
    bestCount = counts(0)
    For d = 1 To 5
        If counts(d) > bestCount Then
            bestCount = counts(d)
            bestDelim = delimiters(d)
        End If
    Next d

    ' 最大分割数を求める
    maxParts = 1
    For r = 1 To dataRange.rows.Count
        cellVal = CStr(dataRange.Cells(r, selCol).Value)
        parts = Split(cellVal, bestDelim)
        partCount = UBound(parts) + 1
        If partCount > maxParts Then maxParts = partCount
    Next r

    ' 新しい列名リスト（M 式用）
    Dim splitColNames As String
    splitColNames = ""
    For i = 1 To maxParts
        If i > 1 Then splitColNames = splitColNames & ", "
        splitColNames = splitColNames & """" & colName & CStr(i) & """"
    Next i

    ' 全列名リストを組み立て（分割列の位置に分割後列を入れる）
    Dim allCols As String
    allCols = ""
    Dim totalCols As Long
    totalCols = lo.ListColumns.Count
    Dim c As Long
    For c = 1 To totalCols
        Dim hdr As String
        hdr = lo.HeaderRowRange.Cells(1, c).Value
        If c = selCol Then
            For i = 1 To maxParts
                If allCols <> "" Then allCols = allCols & ", "
                allCols = allCols & """" & colName & CStr(i) & """"
            Next i
        Else
            If allCols <> "" Then allCols = allCols & ", "
            allCols = allCols & """" & hdr & """"
        End If
    Next c

    ' クエリ名・シート名・テーブル名
    qryName = tblName & "_分割"
    newSheetName = tblName & "_分割"
    newTblName = tblName & "_分割"

    ' 既存クエリ削除
    Dim q As Object
    For Each q In wb.Queries
        If q.Name = qryName Then
            q.Delete
            Exit For
        End If
    Next q

    ' 既存シート削除
    Application.DisplayAlerts = False
    For Each existWs In wb.Sheets
        If existWs.Name = newSheetName Then
            existWs.Delete
            Exit For
        End If
    Next existWs
    Application.DisplayAlerts = True

    ' 区切り文字エスケープ（M 文字列用）
    Dim delimEsc As String
    delimEsc = bestDelim
    ' M 式組み立て
    ' シート名とテーブル名
    Dim srcSheet As String
    srcSheet = ws.Name
    ' M 式
    mFormula = "let" & Chr(13) & Chr(10)
    mFormula = mFormula & "    Source = Excel.CurrentWorkbook(){[Name=" & Chr(34) & tblName & Chr(34) & "]}[Content]," & Chr(13) & Chr(10)
    mFormula = mFormula & "    Split = Table.SplitColumn(Source, " & Chr(34) & colName & Chr(34) & ", Splitter.SplitTextByDelimiter(" & Chr(34) & delimEsc & Chr(34) & ", QuoteStyle.None), {" & splitColNames & "})," & Chr(13) & Chr(10)
    ' 全列を型変換（分割列はtext、他は元の型を維持するためtextで統一後は気にしない）
    ' 分割列をTextに型変換
    Dim typeLines As String
    typeLines = "    Typed = Table.TransformColumnTypes(Split, {"
    For i = 1 To maxParts
        If i > 1 Then typeLines = typeLines & ", "
        typeLines = typeLines & "{" & Chr(34) & colName & CStr(i) & Chr(34) & ", type text}"
    Next i
    typeLines = typeLines & "})," & Chr(13) & Chr(10)
    mFormula = mFormula & typeLines
    ' 列順を元の列順に並べ替え
    mFormula = mFormula & "    Reordered = Table.ReorderColumns(Typed, {" & allCols & "})" & Chr(13) & Chr(10)
    mFormula = mFormula & "in" & Chr(13) & Chr(10)
    mFormula = mFormula & "    Reordered"

    ' クエリ追加
    wb.Queries.Add qryName, mFormula

    ' 新シート追加
    Set newWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    newWs.Name = newSheetName

    ' ListObject でクエリ読み込み
    Dim newLo As ListObject
    Set newLo = newWs.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qryName, _
        Destination:=newWs.Range("A1"))
    newLo.QueryTable.CommandType = 2
    newLo.QueryTable.CommandText = Array("SELECT * FROM [" & qryName & "]")
    newLo.QueryTable.Refresh False
    newLo.Name = newTblName

End Sub

Sub 選択列の式を揃える()
    ' 依頼の語: 式が途切れ|式の途切れ|値で上書き|計算式が崩れ|上下と同じ式に戻|式が抜|列の数式
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
            lastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

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
                found = False
                Dim cc As Long
                For cc = 1 To rowUsed.Columns.Count
                    Dim cv As String
                    cv = CStr(rowUsed.Cells(1, cc).Value)
                    If cv = "合計" Or cv = "計" Or cv = "総計" Then
                        found = True
                        Exit For
                    End If
                Next cc
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
                found = False
                For cc = 1 To rowUsed.Columns.Count
                    cv = CStr(rowUsed.Cells(1, cc).Value)
                    If cv = "合計" Or cv = "計" Or cv = "総計" Then
                        found = True
                        Exit For
                    End If
                Next cc
                If found Then GoTo NextRow2

                Set cellVal = ws.Cells(r, col)
                ' 空セルは触らない
                If cellVal.Value = "" And Not cellVal.HasFormula Then GoTo NextRow2

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

Sub エラー値を一覧にする()
    ' 依頼の語: エラー値|エラーのセル|エラー値を探|DIV/0|#N/A|#VALUE|計算エラー|エラーになって|エラー値のある
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
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
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
        If ws.Cells(hdrRow, c).Value <> "" Then
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
            If ws.Cells(r, c).Value <> "" Or IsError(ws.Cells(r, c).Value) Then
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
                If VarType(ws.Cells(r, c).Value) = vbString And ws.Cells(r, c).Value <> "" Then
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
    outCol = tblRight + 2

    ' 既存の一覧を消す（見出し行から下）
    Dim lastOutRow As Long
    lastOutRow = hdrRow
    For r = hdrRow To urTop + urRows - 1
        If ws.Cells(r, outCol).Value <> "" Or ws.Cells(r, outCol + 1).Value <> "" Or ws.Cells(r, outCol + 2).Value <> "" Then
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

End Sub

Sub テーブルの範囲を拡張する()
    ' 依頼の語: 拡張|下の行|テーブルの範囲|足した行|追加した列
    ' 扱う: テーブル拡張
    ' 見出し: なし
    ' 形: 数の列 日付の列 テーブル
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
        Do While ws.Cells(hdrRow, c).Value <> ""
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
                If ws.Cells(r, cc).Value <> "" Then
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

Sub 空の列を削除する()
    ' 依頼の語: 空の列を削除|空白の列を削除|空列を取り除|間の空いた列を詰|空の列を消|最後に値|入っていない列
    ' 扱う: 空列削除
    ' 見出し: なし
    ' 形: 数の列 日付の列

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
            If ws.Cells(r, c).Value <> "" Then cnt = cnt + 1
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
        If ws.Cells(hdrRow, c).Value <> "" Then
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
            If ws.Cells(r, c).Value <> "" Then
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
            If ws.Cells(r, c).Value <> "" Then
                isEmpty = False
                Exit For
            End If
        Next r
        If isEmpty Then
            ws.Cells(hdrRow, c).EntireColumn.Delete
        End If
    Next c

End Sub

Sub 名前の揺れを一覧にする()
    ' 依頼の語: 名寄せ|株式会社の有無|名寄せの候補|相手先を探|まとめる名前|種類の数|元の名前|名前の揺|同じ会社|列で同じ相手
    ' 扱う: 名寄せ 名前の揺れ 合計
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 日付の列

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).Row

    Dim tblLeft As Long, tblRight As Long
    tblLeft = 1
    tblRight = tblLeft
    Dim c As Long
    For c = tblLeft To ws.Columns.Count
        If ws.Cells(hdrRow, c).Value = "" Then
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
            If ws.Cells(r, c).Value <> "" Then
                rowEmpty = False
                Exit For
            End If
        Next c
        If Not rowEmpty Then lastRow = r
    Next r

    Dim selCol As Long
    selCol = Selection.Cells(1, 1).Column

    Dim outCol As Long
    outCol = tblRight + 2

    Dim clearCol As Long
    For clearCol = outCol To outCol + 2
        ws.Columns(clearCol).ClearContents
    Next clearCol

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
            If ws.Cells(r, c).Value <> "" Then
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
            If cv = "合計" Or cv = "計" Or cv = "総計" Or cv = "小計" Then
                hasSum = True
                Exit For
            End If
        Next c
        If hasSum Then GoTo NextRow

        Dim rawVal As String
        rawVal = CStr(ws.Cells(r, selCol).Value)
        If rawVal = "" Then GoTo NextRow

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

Sub ピボットに累計フィールドを足す()
    ' 依頼の語: ピボットに累計|ピボットで累計|累計をピボット|累計の値|合計の横に累計|列をピボット
    ' 扱う: 累計 ピボット
    ' 見出し: なし
    ' 形: 別のシート ピボット

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pt_idx As Integer
    Dim df_idx As Integer
    Dim rf_idx As Integer

    Set ws = ActiveSheet
    If ws.PivotTables.Count = 0 Then Exit Sub

    For pt_idx = 1 To ws.PivotTables.Count
        Set pt = ws.PivotTables(pt_idx)

        Dim alreadyExists As Boolean
        alreadyExists = False
        For df_idx = 1 To pt.DataFields.Count
            If pt.DataFields(df_idx).Caption = "累計" Then
                alreadyExists = True
                Exit For
            End If
        Next df_idx
        If alreadyExists Then GoTo NextPivot

        If pt.DataFields.Count = 0 Then GoTo NextPivot

        Dim srcName As String
        srcName = pt.DataFields(1).SourceName

        Dim srcField As PivotField
        On Error Resume Next
        Set srcField = Nothing
        Set srcField = pt.PivotFields(srcName)
        On Error GoTo 0
        If srcField Is Nothing Then GoTo NextPivot

        ' 列フィールド名を辞書に集める
        Dim colFieldNames As Object
        Set colFieldNames = CreateObject("Scripting.Dictionary")
        Dim cf_idx As Integer
        For cf_idx = 1 To pt.ColumnFields.Count
            Dim cfName As String
            cfName = pt.ColumnFields(cf_idx).Name
            If cfName <> "Values" Then
                colFieldNames(cfName) = 1
            End If
        Next cf_idx

        ' いちばん外側の行フィールドを Position 最小で探す
        Dim outerRowField As PivotField
        Set outerRowField = Nothing
        Dim minPos As Integer
        minPos = 32767
        For rf_idx = 1 To pt.RowFields.Count
            Dim tmpPF As PivotField
            Set tmpPF = pt.RowFields(rf_idx)
            If tmpPF.Name <> "Values" And Not colFieldNames.Exists(tmpPF.Name) Then
                If tmpPF.Position < minPos Then
                    minPos = tmpPF.Position
                    Set outerRowField = tmpPF
                End If
            End If
        Next rf_idx
        If outerRowField Is Nothing Then GoTo NextPivot

        Dim newDF As PivotField
        On Error Resume Next
        Set newDF = Nothing
        Set newDF = pt.AddDataField(srcField, "累計", xlSum)
        On Error GoTo 0
        If newDF Is Nothing Then GoTo NextPivot

        On Error Resume Next
        newDF.Calculation = xlRunningTotal
        newDF.BaseField = outerRowField.Name
        newDF.NumberFormat = "#,##0"
        On Error GoTo 0

NextPivot:
    Next pt_idx
End Sub

Sub ピボットを統一形式にそろえる()
    ' 依頼の語: ピボットを表形式|ピボットの体裁|デザインをいつも|全部小計|ピボットを見
    ' 扱う: ピボット体裁 ピボット表形式 ピボット小計 ピボットスタイル フィルタ ピボット
    ' 見出し: なし
    ' 形: 別のシート ピボット
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim df As PivotField
    Dim i As Integer
    Dim subs(11) As Boolean

    Set ws = ActiveSheet

    If ws.PivotTables.Count = 0 Then Exit Sub

    For i = 0 To 11
        subs(i) = False
    Next i

    Dim pt_idx As Integer
    For pt_idx = 1 To ws.PivotTables.Count
        Set pt = ws.PivotTables(pt_idx)

        On Error Resume Next
        pt.TableStyle2 = "PivotStyleMedium9"
        pt.RowAxisLayout xlTabularRow
        pt.RepeatAllLabels xlRepeatLabels
        pt.NullString = "0"
        pt.DisplayNullString = True
        On Error GoTo 0

        For Each pf In pt.RowFields
            On Error Resume Next
            pf.Subtotals = subs
            On Error GoTo 0
        Next pf

        For Each df In pt.DataFields
            On Error Resume Next
            df.NumberFormat = "#,##0"
            On Error GoTo 0
        Next df
    Next pt_idx
End Sub
