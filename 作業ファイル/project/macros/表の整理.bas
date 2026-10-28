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
    ' 変えてよい: 文字日付 文字数字
    ' 依頼の語: 表を修正|表の修正|表を修復|表の修復|表を直して|表を直す|表の直し|表を整え|表記を統一|全角と半角|書き方をそろえ|書き方を揃え|入力の書き方|見やすく整理|表を整理|表をきれいに|表記をそろえ|表記を揃え|表記ゆれ|表記揺れ|表記の揺れ|全角半角|表の書式をそろえ|表の書式を整え|表の見た目を整え|表の体裁を整え|表全体を整え|整ってない|整っていない|表記ゆれを直して|表記ゆれを直す|お任せしますので良いように|おまかせしますので良いように|お任せしますのでよいように|おまかせしますのでよいように
    ' 階級: 基本
    ' 依頼の組: 表,表の,表記,書き方,全角,半角,ばらばら,バラバラ,ごちゃごちゃ,ひらがな,カタカナ,混ざ,(株),株式会社,形式,書式,罫線+修正,修復,そろえ,揃え,統一,整え,整理,直,引き直+-住所,番地,郵便,グラフ,シート,名前の揺,列幅,ピボット,データ,洗い出,一覧,探,調べ
    ' 値の意味は変えずに書き方をそろえる（行は消さない・数式のセルは触らない）: 前後の空白・全角英数→半角・半角カナ→全角・
    ' 文字の間の空白（列の多い方の 1 つ）・（株）→株式会社・電話・郵便番号・メール・文字の日付と和暦→日付・文字の金額→数値・率・罫線・列幅
    Dim ws As Worksheet, ur As Range, tb As Range, c As Range, dv As Object
    Dim r0 As Long, c0 As Long, nr As Long, nc As Long, scanTo As Long
    Dim hR As Long, hc1 As Long, hc2 As Long, lastR As Long
    Dim i As Long, j As Long, k As Long, n As Long, best As Long, ln As Long, a As Long, b As Long, g As Long
    Dim v As Variant, nv As Variant, x As Variant, parts As Variant, ks As Variant
    Dim s As String, t As String, u As String, ch As String, hd As String, kind As String, sp As String
    Dim digits As String, kanaRun As String, mon As String
    Dim cnt As Long, dn As Long, nn As Long, cn As Long, zs As Long, hs As Long, an As Long, pn As Long
    Dim ok As Boolean, bad As Boolean, dt As Date, dbl As Double
    Dim era As Long, Y As Long, m As Long, d As Long, cd As Long
    Dim ar(1 To 200) As String, al(1 To 200) As Long, bl(1 To 200) As Long, nA As Long
    Dim grp() As Long, k1() As String, k2() As String, rep() As String, hira() As Long, cntOf() As Long
    Dim hasKana() As Boolean, hasKanji() As Boolean
    Dim mainEnd As Long, sumR1 As Long, sumC1 As Long, sumC2 As Long, gapR As Long, rr As Long
    Dim colW() As Double
    Dim kanaHira As Boolean, bigN As Long
    Dim tv As String, tbm As String, tv0 As String, tb0 As String, rgC As Range, arr As Variant, 集計行() As Boolean
    Dim colV As Variant, colF As Variant, tmpA() As Variant, 半カナ列 As Boolean
    Dim want As String, wantAl As Long, alNum As Long, alTxt As Long, 罫線様式 As Boolean, 自由記述 As Boolean
    Dim zp As Long, hp As Long, 全角括弧 As Boolean, 欄 As Boolean, 下の表 As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    r0 = ur.row: c0 = ur.Column: nr = ur.rows.Count: nc = ur.Columns.Count
    If nr < 2 Then Exit Sub
    GoSub 見出しを探す
    If hR = 0 Then Exit Sub
    Dim 最後 As Long
表の始まり:
    最後 = hR
    ' 見出しの帯が 2 行（縦に結合した見出し＋「数量」の下の「実棚」「帳簿」）なら、下の行まで見出し（下の段を明細と見て寄せを
    ' 変えた・2026-10-08 Sonnet の採点 S7 C4・D4）。下の行が文字だけで、見出しの行に 2 行にまたがる結合があるときだけ
    ok = False
    For j = c0 To c0 + nc - 1
        If ws.Cells(hR, j).MergeCells Then
            If ws.Cells(hR, j).MergeArea.rows.Count > 1 And ws.Cells(hR, j).MergeArea.row = hR Then ok = True: Exit For
        End If
    Next
    If ok And hR + 1 <= r0 + nr - 1 Then
        For j = c0 To c0 + nc - 1
            v = ws.Cells(hR + 1, j).Value
            If ws.Cells(hR + 1, j).HasFormula Then ok = False: Exit For
            If Not isEmpty(v) And VarType(v) <> vbString Then ok = False: Exit For
        Next
        If ok Then hR = hR + 1
    End If
    ' 見出しの行を「列まるごと空の列」と「式の見出し（=COUNTA などの集計）」で区切り、見出しが 2 つ以上ある塊ごとに
    ' 1 つの表として直す（2026-09-23）。左右に並んだ 2 つの表と右のメモの列を 1 つの表と見て、間の空き列・右の表の下・
    ' メモの列にまで罫線を引き、メモの全角かっこを半角にしていた（お試し版テスト用4）。見出しの無い列も中に値があれば
    ' 表の列、2024 のような数の見出しも区切りにしない（年を並べた横持ちの表を割らない）
    Dim blk1(1 To 50) As Long, blk2(1 To 50) As Long, nBlk As Long, bi As Long, inBlk As Boolean
    nBlk = 0: inBlk = False
    For j = c0 To c0 + nc
        ok = False
        If j <= c0 + nc - 1 Then
            If ws.Cells(hR, j).HasFormula Then
                ok = False
            ElseIf Len(Trim(ws.Cells(hR, j).text)) > 0 Then
                ok = True
            Else
                ok = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(hR, j), ws.Cells(r0 + nr - 1, j))) > 0
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
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(hR, hc1), ws.Cells(hR, hc2))) >= 2 Then
            GoSub 本文の終わり
            If lastR > hR Then
                GoSub 集計欄を分ける      ' 空行の下の細い塊（合計・平均）で表を終える。その先の行（注記など）は触らない（2026-09-13）
                GoSub 一つの表を直す
                If lastR > 最後 Then 最後 = lastR
            End If
        End If
    Next
    ' 表の下に空行をはさんで別の表（見出しの行＝値が 2 つ以上）があれば、それも直す（2 つ目の表の「２」が文字のまま残り、
    ' 合計が合わなかった・2026-10-08 Sonnet の採点 S9 B14）。注記（1 行に 1 セル）は表ではない
    For i = 最後 + 2 To r0 + nr - 1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i - 1, c0), ws.Cells(i - 1, c0 + nc - 1))) = 0 Then
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1))) >= 2 Then
                If i + 1 <= r0 + nr - 1 Then
                    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i + 1, c0), ws.Cells(i + 1, c0 + nc - 1))) >= 1 Then
                        hR = i: 下の表 = True
                        GoTo 表の始まり
                    End If
                End If
            End If
        End If
    Next
    Exit Sub

一つの表を直す:
    ' 合計・小計の行は表示形式と寄せを決めるときに数えず、そろえもしない（「合計」の結合セルの右寄せが混ざって、No の列を
    ' 中央にそろえ直していた・2026-10-08 Sonnet の出題 Q5）
    ReDim 集計行(hR + 1 To lastR)
    arr = ws.Range(ws.Cells(hR + 1, hc1), ws.Cells(lastR, hc2)).Value2
    If IsArray(arr) Then
        For i = 1 To UBound(arr, 1)
            For k = 1 To UBound(arr, 2)
                If 集計の語か(arr(i, k)) Then 集計行(hR + i) = True: Exit For
            Next k
        Next i
    End If
    For j = hc1 To hc2
        hd = UCase(StrConv(ws.Cells(hR, j).MergeArea.Cells(1, 1).text, vbNarrow))     ' 縦に結合した見出しは上の行の字
        ' 列の値と式は最初にまとめて読む（セルを 1 つずつ 2 回読み、1 万行で遅かった・2026-10-08）
        colV = ws.Range(ws.Cells(hR + 1, j), ws.Cells(lastR, j)).Value
        colF = ws.Range(ws.Cells(hR + 1, j), ws.Cells(lastR, j)).Formula
        If Not IsArray(colV) Then
            ReDim tmpA(1 To 1, 1 To 1): tmpA(1, 1) = colV: colV = tmpA
            ReDim tmpA(1 To 1, 1 To 1): tmpA(1, 1) = colF: colF = tmpA
        End If
        GoSub 列を調べる
        For i = hR + 1 To lastR
            If Left$(CStr(colF(i - hR, 1)), 1) <> "=" Then
                v = colV(i - hR, 1)
                If Not isEmpty(v) And Not IsError(v) Then
                    nv = v
                    GoSub 値を直す
                    If VarType(nv) <> VarType(v) Then
                        Set c = ws.Cells(i, j)
                        GoSub 書く
                    ElseIf nv <> v Then
                        Set c = ws.Cells(i, j)
                        GoSub 書く
                    End If
                End If
            End If
        Next
        If kind = "text" Then GoSub 言い換えをそろえる
        ' 表示形式と寄せは、列の中で混ざっているとき・標準のままのときだけそろえる。列の全部が同じ形なら
        ' それが表を作った人の様式＝触らない（2026-10-08 shu「完成した表まで強制的に直している」）
        want = ""
        If kind = "date" Then
            want = "yyyy/m/d": wantAl = xlLeft
        ElseIf kind = "pct" Then
            want = "0%": wantAl = xlRight
        ElseIf kind = "num" Then
            If InStr(hd, "年") = 0 And InStr(hd, "率") = 0 And InStr(hd, "%") = 0 Then want = "#,##0"
            wantAl = xlRight
        ElseIf kind = "code" Then
            wantAl = xlLeft     ' 番号・コードの列は左寄せ（tidy と同じ決まり。中央にしていた・2026-10-09 shu「会員番号がなんで右寄せなんだよ」）
        Else
            wantAl = xlLeft
        End If
        If Len(want) > 0 Then GoSub 表示形式をそろえる
        GoSub 寄せをそろえる
    Next

    ' 本文の文字の書式（見出しの書体と大きさ・太字／斜体／下線／取り消し線なし・文字色は自動）・塗りのばらつき・罫線・列幅
    ' 書体と大きさは本文の中で混ざっているときだけ見出しにそろえる。太字・斜体・下線・取り消し線・文字の色・塗りは
    ' 人が付けた印（赤字のマイナス・要注意の塗り）なので触らない（2026-10-08 完成した表まで直していた）
    Set tb = ws.Range(ws.Cells(hR + 1, hc1), ws.Cells(lastR, hc2))
    If IsNull(tb.Font.Name) Or IsNull(tb.Font.Size) Then
        With tb.Font
            .Name = ws.Cells(hR, hc1).Font.Name
            .Size = ws.Cells(hR, hc1).Font.Size
        End With
    End If
    GoSub 罫線は様式か
    If Not 罫線様式 Then
    ' 罫線は、無い・途切れている（足した行に罫線が無い）ときだけ引き直す。一様に引いてあれば人の決めた罫線
    ' 引き直す前に二重線の場所を覚える（下で戻す）
    Dim dblTop() As Boolean, dblBot() As Boolean, dblI As Long, dblJ As Long
    ReDim dblTop(1 To lastR - hR + 1, 1 To hc2 - hc1 + 1)
    ReDim dblBot(1 To lastR - hR + 1, 1 To hc2 - hc1 + 1)
    Dim dblBottomStart As Long
    dblBottomStart = IIf(lastR - hR > 200, IIf(sumR1 > 0 And sumR1 - 2 < lastR - 30, sumR1 - 2, lastR - 30), hR)
    If dblBottomStart < hR Then dblBottomStart = hR
    For dblI = hR To lastR
        If lastR - hR <= 200 Or dblI <= hR + 5 Or dblI >= dblBottomStart Then
            For dblJ = hc1 To hc2
                dblTop(dblI - hR + 1, dblJ - hc1 + 1) = (ws.Cells(dblI, dblJ).Borders(xlEdgeTop).LineStyle = xlDouble)
                dblBot(dblI - hR + 1, dblJ - hc1 + 1) = (ws.Cells(dblI, dblJ).Borders(xlEdgeBottom).LineStyle = xlDouble)
            Next dblJ
        End If
    Next dblI
    ' 罫線: 空の行の下の、表より細い塊（合計・平均など）は表と分けて、埋まっている列だけを囲む（2026-09-13）
    If sumR1 > 0 Then
        ' 前に引いた罫線が空の行と集計欄の左に残っていれば外す（外してから表と集計欄を引き直す）
        ws.Range(ws.Cells(mainEnd + 1, hc1), ws.Cells(lastR, hc2)).Borders.LineStyle = xlNone
    End If
    Set tb = ws.Range(ws.Cells(hR, hc1), ws.Cells(mainEnd, hc2))
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
    For dblR = hR To lastR
        If lastR - hR <= 200 Or dblR <= hR + 5 Or dblR >= dblBottomStart Then
            For dblC = hc1 To hc2
                If dblTop(dblR - hR + 1, dblC - hc1 + 1) Then ws.Cells(dblR, dblC).Borders(xlEdgeTop).LineStyle = xlDouble
                If dblBot(dblR - hR + 1, dblC - hc1 + 1) Then ws.Cells(dblR, dblC).Borders(xlEdgeBottom).LineStyle = xlDouble
            Next dblC
        End If
    Next dblR
    End If
    ' 列幅は広げるだけ（### と字の切れを直す）。人が広めに取った幅は狭めない・折り返しの列は触らない・広げても 60 まで
    ' （2026-10-08 AutoFit で印刷用に決めた幅まで変えていた）
    ReDim colW(hc1 To hc2)
    For k = hc1 To hc2
        colW(k) = ws.Columns(k).ColumnWidth
    Next
    ws.Range(ws.Cells(hR, hc1), ws.Cells(lastR, hc2)).Columns.AutoFit
    For k = hc1 To hc2
        v = ws.Range(ws.Cells(hR, k), ws.Cells(lastR, k)).WrapText
        If IsNull(v) Then v = True
        If v = True Or ws.Columns(k).ColumnWidth < colW(k) Then
            ws.Columns(k).ColumnWidth = colW(k)
        ElseIf ws.Columns(k).ColumnWidth > 60 Then
            ws.Columns(k).ColumnWidth = IIf(colW(k) > 60, colW(k), 60)
        End If
    Next
    Return

表示形式をそろえる:
    ' 表示形式は、列の中で混ざっているときだけ、値のあるセルの多い方（標準も 1 つの形として数える）にそろえる。
    ' 列まるごと同じ形なら標準のままでもその表の決まり＝触らない（2026-10-08 Haiku の出題: そろった標準の件数の列まで #,##0 にしていた）。
    ' 列まるごと同じかは 1 回で分かる（NumberFormat は混ざると Null を返す）。セルを 1 つずつ読むのは混ざったときだけ（1 万行で遅くなった）
    ' ただし標準のままの数値列（金額等）や日付列には基本書式を適用する（2026-10-17 shu）
    Set rgC = ws.Range(ws.Cells(hR + 1, j), ws.Cells(lastR, j))
    If Not IsNull(rgC.NumberFormat) Then
        ' 標準のままの数値列（金額・数量等）には3桁区切り、日付列には日付書式を適用する（率は維持）
        If (rgC.NumberFormat = "General" Or rgC.NumberFormat = "G/標準") And Len(want) > 0 Then
            If kind = "num" Or kind = "date" Then
                u = want
                ' 列はまるごと標準＝範囲に 1 回で付ける（1 セルずつだと 1 万行で 4 秒かかった・2026-10-21 総点検）。
                ' 日付の列の集計行だけは標準に戻す（前の 1 セルずつの版と同じ結果）
                rgC.NumberFormat = u
                If kind = "date" Then
                    For i = hR + 1 To lastR
                        If 集計行(i) Then ws.Cells(i, j).NumberFormat = "General"
                    Next
                End If
                If kind = "num" And InStr(u, ".") = 0 And u <> "General" And u <> "G/標準" Then GoSub 小数のセル
            End If
        End If
        Return
    End If
    Set dv = CreateObject("Scripting.Dictionary")
    For i = hR + 1 To lastR
        If Not 集計行(i) Then
            Set c = ws.Cells(i, j)
            If Not isEmpty(c.Value) And Not c.MergeCells Then dv(c.NumberFormat) = dv(c.NumberFormat) + 1
        End If
    Next
    If dv.Count <= 1 Then Return     ' 違うのは空のセル・合計の行・結合だけ
    u = "": n = 0
    For Each x In dv.keys
        ' 同じ数なら標準でない方（人が付けた形）に（△ の列で 1 対 1 になり #,##0;[赤]△#,##0 を標準に戻した・2026-10-09 Gemini の出題 S12 Q7）
        If dv(x) > n Or (dv(x) = n And (u = "General" Or u = "G/標準")) Then n = dv(x): u = x
    Next
    If kind = "date" And Len(want) > 0 Then u = want
    For i = hR + 1 To lastR
        Set c = ws.Cells(i, j)
        If Not isEmpty(c.Value) And Not c.MergeCells Then
            If Not 集計行(i) Then
                If c.NumberFormat <> u Then c.NumberFormat = u
            ElseIf kind = "num" Then
                If c.NumberFormat <> u Then c.NumberFormat = u
            End If
        End If
    Next
    If kind = "num" And InStr(u, ".") = 0 And u <> "General" And u <> "G/標準" Then GoSub 小数のセル
    Return

寄せをそろえる:
    ' 値のあるセルの寄せが 1 種類なら触らない（標準のままでも、数と文字が混ざって左右に割れて見えるときだけそろえる）。
    ' 混ざっていれば多い方（多いのが標準なら列の種類の既定）でそろえる
    ' 合計・小計の行と結合セルは数えず、そろえもしない
    Set rgC = ws.Range(ws.Cells(hR + 1, j), ws.Cells(lastR, j))
    v = rgC.HorizontalAlignment
    If Not IsNull(v) Then
        If v <> xlGeneral Then Return
        alNum = 0: alTxt = 0
        arr = rgC.Value2
        If IsArray(arr) Then
            For i = 1 To UBound(arr, 1)
                If Not isEmpty(arr(i, 1)) And Not 集計行(hR + i) Then
                    If VarType(arr(i, 1)) = vbString Then alTxt = alTxt + 1 Else alNum = alNum + 1
                End If
            Next
        End If
        If (alNum = 0 Or alTxt = 0) And kind <> "code" And kind <> "date" Then Return     ' 番号・日付の列は標準のままだと数が右に寄る＝左にそろえる（2026-10-17 shu）
        u = CStr(wantAl)
    Else
        Set dv = CreateObject("Scripting.Dictionary")
        alNum = 0: alTxt = 0
        For i = hR + 1 To lastR
            If Not 集計行(i) Then
                Set c = ws.Cells(i, j)
                If Not isEmpty(c.Value) And Not c.MergeCells Then
                    dv(CStr(c.HorizontalAlignment)) = dv(CStr(c.HorizontalAlignment)) + 1
                    If VarType(c.Value) = vbString Then alTxt = alTxt + 1 Else alNum = alNum + 1
                End If
            End If
        Next
        If dv.Count = 0 Then Return
        If dv.Count = 1 Then
            If dv.keys()(0) <> CStr(xlGeneral) Then Return
            If (alNum = 0 Or alTxt = 0) And kind <> "code" And kind <> "date" Then Return     ' 番号・日付の列は標準のままだと数が右に寄る＝左にそろえる（2026-10-17 shu）
            u = CStr(wantAl)
        Else
            u = "": n = 0
            For Each x In dv.keys
                If dv(x) > n Then n = dv(x): u = x
            Next
            If u = CStr(xlGeneral) Then u = CStr(wantAl)
        End If
    End If
    If rgC.MergeCells = False Then
        rgC.HorizontalAlignment = CLng(u)
    Else
        For i = hR + 1 To lastR
            If Not 集計行(i) Then
                Set c = ws.Cells(i, j)
                If Not c.MergeCells Then
                    If c.HorizontalAlignment <> CLng(u) Then c.HorizontalAlignment = CLng(u)
                End If
            End If
        Next
    End If
    Return

罫線は様式か:
    ' 6 方向（上下左右・内側の横と縦）のどれも混ざっておらず、どれかは引いてある＝人が決めた罫線（外枠だけ・横罫だけも含む）。
    ' 下の合計の行（上の二重線）は数えない
    罫線様式 = False
    rr = mainEnd
    Do While rr > hR + 1
        ok = False
        For k = hc1 To hc2
            If 集計の語か(ws.Cells(rr, k).Value) Then ok = True: Exit For
        Next
        If ok Then rr = rr - 1 Else Exit Do
    Loop
    ' 範囲の Borders(辺) は混ざっていても最初のセルの値を返す（Null にならない・2026-10-08 実測）＝明細の行ごとにセルの罫線を比べる。
    ' 明細の 1 行目に罫線が無ければ引く。どの行も 1 行目と同じ形（縦の線・行の下の線。最後の行の下は外枠なので比べない）なら人の罫線
    If rr < hR + 1 Then Return
    i = hR + 1: GoSub 行の罫線
    tv0 = tv: tb0 = tbm
    ' 明細の 1 行目に罫線が無い本体の表（列が 3 つ以上）＝罫線が出来上がっていない表＝引く（2026-10-09 shu「なんで罫線とかちゃんと
    ' 全部引かれないんだよ」。10/8 は「表のどこにも罫線が無ければ、罫線を引かない様式」として引かず、テスト用4 の左の表が罫線なしの
    ' まま・右に足した突き合わせの列だけ罫線が付いた）。表の下に見つけた 2 つ目の塊・2 列の「項目｜値」の組（請求書の振込先）は
    ' 表ではない＝引かない
    If InStr(tv0 & tb0, "1") = 0 Then
        If 下の表 Or hc2 - hc1 + 1 < 3 Then 罫線様式 = True
        Return
    End If
    ' 明細の 1 行目で、下の線がある列と無い列が混ざっている＝一部の列にだけ罫線（足した列だけ格子・2026-10-09 テスト用4 の
    ' 突き合わせの F・G 列）＝継ぎはぎ＝引く。外枠だけ（下の線は全部無し）・横罫だけ・格子（全部有り）は様式のまま
    If InStr(tb0, "0") > 0 And InStr(tb0, "1") > 0 Then Return
    ' 200 行を超える表は、等間隔の約 100 行と最後の 20 行だけ比べる（2 回目に撃つと引いた罫線を 1 万行ぶん 1 セルずつ読み、
    ' 1 回目の倍近くかかった・2026-10-08 Sonnet の採点。足した行の罫線抜けは表の下に出るので下の 20 行は全部見る）
    n = 1
    If rr - hR > 200 Then n = (rr - hR) \ 100
    For i = hR + 2 To rr
        If n = 1 Or (i - hR) Mod n = 0 Or i > rr - 20 Then
            GoSub 行の罫線
            If tv <> tv0 Then Return
            If i < rr And tbm <> tb0 Then Return
        End If
    Next
    罫線様式 = True
    Return

行の罫線:
    ' 行 i のセルごとの左の線・下の線と、右端の線（結合の中は m＝比べない）
    tv = "": tbm = ""
    For k = hc1 To hc2
        If ws.Cells(i, k).MergeCells Then
            tv = tv & "m": tbm = tbm & "m"
        Else
            tv = tv & IIf(ws.Cells(i, k).Borders(xlEdgeLeft).LineStyle = xlNone, "0", "1")
            tbm = tbm & IIf(ws.Cells(i, k).Borders(xlEdgeBottom).LineStyle = xlNone, "0", "1")
        End If
    Next
    tv = tv & IIf(ws.Cells(i, hc2).Borders(xlEdgeRight).LineStyle = xlNone, "0", "1")
    Return

前後の空白を取る:
    Do While Len(s) > 0
        If Left(s, 1) = " " Or Left(s, 1) = "　" Or Left(s, 1) = vbTab Then s = Mid(s, 2) Else Exit Do
    Loop
    Do While Len(s) > 0
        If Right(s, 1) = " " Or Right(s, 1) = "　" Or Right(s, 1) = vbTab Then s = Left(s, Len(s) - 1) Else Exit Do
    Loop
    Return

見出しを探す:
    ' 先頭 20 行で、値のある列が最も多い行の 6 割以上を最初に満たす行＝見出し。本文は見出しの列が全部空の行の手前まで
    scanTo = r0 + IIf(nr > 20, 20, nr) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If n > best Then best = n
    Next
    hR = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If hR = 0 And n >= 2 And n >= best * 0.6 Then hR = i
    Next
    Return

本文の終わり:
    ' 本文の終わり: 途中の空行 1 行は飛ばして続ける。空行が 2 行続いたら表の終わり
    lastR = hR
    n = 0
    For i = hR + 1 To r0 + nr - 1
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
    i = hR + 1
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
                ' 数・式・集計の語が 1 つも無い塊は注記（「※ 2026年度版」など）＝表の外。罫線も書き方の直しも当てない
                ' （表の下の結合した注記に枠を付け、全角を半角にしていた・2026-10-08 Sonnet の採点 S2 A18）
                ok = False
                For rr = i To rr - 1
                    For k = hc1 To hc2
                        v = ws.Cells(rr, k).Value
                        ' 集計欄と見なすのは式か「合計」などの語がある塊だけ（数だけの塊＝請求書の下の振込先の組を集計欄として
                        ' 罫線と寄せを当てた・2026-10-08 Sonnet の採点 S10）
                        If ws.Cells(rr, k).HasFormula Or 集計の語か(v) Then ok = True
                    Next
                    If ok Then Exit For
                Next
                If Not ok Then
                    lastR = gapR - 1: mainEnd = lastR: sumR1 = 0: sumC1 = 0: sumC2 = 0
                    Return
                End If
                rr = i
                Do While rr <= lastR
                    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(rr, hc1), ws.Cells(rr, hc2))) = 0 Then Exit Do
                    rr = rr + 1
                Loop
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
    For i = hR + 1 To lastR
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
        kanaHira = (InStr(ws.Cells(hR, j).text, "かな") > 0 Or InStr(ws.Cells(hR, j).text, "がな") > 0)     ' hd は半角にした後（ひらがなも半角カナになる）なので元の見出しで見る
    ElseIf InStr(hd, "番号") > 0 Or InStr(hd, "ｺｰﾄﾞ") > 0 Or InStr(hd, "品番") > 0 Or hd = "ID" Or hd = "NO" Or hd = "NO." Then
        kind = "code"
    ElseIf InStr(hd, "ﾒｰﾙ") > 0 Or InStr(hd, "MAIL") > 0 Then
        kind = "mail"
    End If
    cnt = 0: dn = 0: nn = 0: cn = 0: zs = 0: hs = 0: nA = 0: an = 0: pn = 0: bigN = 0: zp = 0: hp = 0
    For i = hR + 1 To lastR
        v = colV(i - hR, 1)
        If Not isEmpty(v) And Not IsError(v) Then
            cnt = cnt + 1
            GoSub 日付を読む
            If ok Then
                dn = dn + 1
            ElseIf InStr(hd, "日") > 0 And VarType(v) = vbDouble Then
                If v > 20000 And v < 80000 And v = Int(v) Then dn = dn + 1     ' 日付の列で数値のままの日付
            End If
            GoSub 数を読む
            If ok Then
                nn = nn + 1
                If Abs(dbl) > 1.5 Then bigN = bigN + 1     ' 割合（0.12）でない数（12.5・120）
            End If
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
                ' チェック欄（□ 消耗品   ■ 備品）の空白は並びの間隔＝空白の多い方の数に入れない（2026-10-09 Gemini の出題 S13 C10）
                欄 = (InStr(t, ChrW(&H25A1)) + InStr(t, ChrW(&H25A0)) + InStr(t, ChrW(&H2610)) + InStr(t, ChrW(&H2611)) + InStr(t, ChrW(&H2612)) > 0)
                If Not 欄 And InStr(t, "　") > 0 Then zs = zs + 1
                If Not 欄 And InStr(t, " ") > 0 Then hs = hs + 1
                If InStr(t, ChrW(&HFF08)) > 0 Or InStr(t, ChrW(&HFF09)) > 0 Then zp = zp + 1
                If InStr(t, "(") > 0 Or InStr(t, ")") > 0 Then hp = hp + 1
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
        ' 見出しの「率」だけで % にするのは、数が割合（1.5 以下）のときだけ（120 が 12000% に見えた・2026-10-08）
        ElseIf cnt > 0 And (pn >= cnt * 0.3 Or (InStr(hd, "率") > 0 And nn + pn >= cnt * 0.6 And bigN = 0)) Then
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
    ' かっこも列の多い方に（全角だけの列の「ノート（B5）」を半角にしていた・2026-10-09 Gemini の出題 S14 C4）
    全角括弧 = (zp > hp)
    ' カナが全部半角の列（全銀の口座名義・半角でそろえたフリガナ）は半角のまま（そろった書き方は変えない・2026-10-08）
    a = 0: b = 0
    For i = hR + 1 To lastR
        v = colV(i - hR, 1)
        If VarType(v) = vbString Then
            If v Like "*[" & ChrW(&HFF66) & "-" & ChrW(&HFF9F) & "]*" Then a = a + 1
            If v Like "*[" & ChrW(&H3041) & "-" & ChrW(&H30FA) & "]*" Then b = b + 1
        End If
    Next
    半カナ列 = (a > 0 And b = 0)
    ' 備考・メモのような自由に書く列は、前後の空白を取るだけ（改行・かっこ・全角の英数は書いた人の書き方・2026-10-08）
    自由記述 = False
    For Each x In Array("備考", "ﾒﾓ", "摘要", "ｺﾒﾝﾄ", "内容", "説明", "特記", "注記", "所見", "理由", "経過", "連絡事項", "意見", "感想", "要望", "回答", "記述", "質問")
        If InStr(hd, x) > 0 Then 自由記述 = True
    Next
    ' 見出しの語が無くても、セルの中に改行がある列は文章（改行を空白にし全角を半角にしていた・2026-10-08 Sonnet の採点 S8 H 列）
    ' 住所・氏名・電話・番号・日付・金額のような形の決まった列は、改行があっても文章ではない（紛れ込んだ改行は汚れ）
    ok = False
    For Each x In Array("住所", "所在地", "氏名", "名前", "電話", "TEL", "郵便", "ﾒｰﾙ", "MAIL", "番号", "ｺｰﾄﾞ", "ID", "日", "金額", "数", "額")
        If InStr(hd, x) > 0 Then ok = True
    Next
    If Not 自由記述 And Not ok Then
        For i = hR + 1 To lastR
            If VarType(colV(i - hR, 1)) = vbString Then
                If InStr(colV(i - hR, 1), vbLf) > 0 Then 自由記述 = True: Exit For
            End If
        Next
    End If
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
            ElseIf VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
                If v >= 10000000 And v < 10000000000# And v = Int(v) Then
                    s = Format(v, "0"): GoSub 電話を直す: nv = s
                End If
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
                ' 大文字小文字は変えない（2026-10-08 Sonnet の採点 S8 D7）
                nv = Replace(Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", ""), "＠", "@")
            End If
        Case Else
            If VarType(v) = vbString Then
                s = v
                If 自由記述 Then GoSub 前後の空白を取る Else GoSub 文字を直す
                nv = s
            End If
    End Select
    Return

書く:
    If VarType(nv) = vbString Then
        If kind = "phone" Or kind = "postal" Or kind = "code" Or IsNumeric(nv) Or IsDate(nv) Then c.NumberFormat = "@"
    ElseIf c.NumberFormat = "@" Then
        c.NumberFormat = "General"
    End If
    ' 改行を含む文字を書くと Excel が折り返しを付ける＝書く前の設定に戻す（2026-10-08 Sonnet の採点 S9 D4）
    v = c.WrapText
    dbl = c.RowHeight
    c.Value = nv
    If c.WrapText <> v Then c.WrapText = v
    If c.RowHeight <> dbl Then c.EntireRow.RowHeight = dbl
    Return

文字を直す:
    ' セルの中の改行は残す（改行つきの住所は書いた人の書き方・2026-10-08 Sonnet の採点 S9 D4）。紛れ込んだ CR だけ LF にそろえる。
    ' 全角の英数字・記号 → 半角（全角空白はあとで）／半角カナ → 全角（濁点も合わせる）
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    u = "": kanaRun = ""
    For k = 1 To Len(s)
        ch = Mid(s, k, 1)
        cd = AscW(ch): If cd < 0 Then cd = cd + 65536
        If cd >= &HFF61& And cd <= &HFF9F& And Not 半カナ列 Then
            kanaRun = kanaRun & ch
        Else
            If Len(kanaRun) > 0 Then u = u & StrConv(kanaRun, vbWide): kanaRun = ""
            ' 全角の「～」は期間の区切り（9/1～9/3）として書く字＝半角の「~」にしない（2026-10-08 Sonnet の採点 S3 G 列）
            If cd >= &HFF01& And cd <= &HFF5D& Then u = u & ChrW(cd - &HFEE0&) Else u = u & ch
        End If
    Next
    If Len(kanaRun) > 0 Then u = u & StrConv(kanaRun, vbWide)
    s = u
    ' フリガナは見出しの書き方にそろえる（「フリガナ」ならカタカナ・「ふりがな」ならひらがな。2026-10-08 ふりがなの列をカタカナにしていた）
    If kind = "kana" Then s = StrConv(s, IIf(kanaHira, vbHiragana, vbKatakana))
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
    ' 文字の間の空白は、列の中で多い方の 1 つに。チェック欄（□ 消耗品   ■ 備品）の空白は並びの間隔なので残す（2026-10-09 S13 C10）
    欄 = (InStr(s, ChrW(&H25A1)) + InStr(s, ChrW(&H25A0)) + InStr(s, ChrW(&H2610)) + InStr(s, ChrW(&H2611)) + InStr(s, ChrW(&H2612)) > 0)
    If Not 欄 Then
        s = Replace(s, "　", " ")
        Do While InStr(s, "  ") > 0
            s = Replace(s, "  ", " ")
        Loop
    End If
    If kind = "corp" Then
        ' 会社の種類の略は正式の名前に・種類の前後の空白は取る（前株・後株の位置は変えない）
        s = Replace(Replace(Replace(Replace(s, "(株)", "株式会社"), "(有)", "有限会社"), "㈱", "株式会社"), "㈲", "有限会社")
        s = Replace(Replace(s, "株式会社 ", "株式会社"), " 株式会社", "株式会社")
        s = Replace(Replace(s, "有限会社 ", "有限会社"), " 有限会社", "有限会社")
    End If
    If 全角括弧 Then s = Replace(Replace(s, "(", ChrW(&HFF08)), ")", ChrW(&HFF09))
    If sp = "　" And Not 欄 Then s = Replace(s, " ", "　")
    Return

言い換えをそろえる:
    ' 状態・区分のような列（同じ値が並ぶ・種類 15 以下）で、送り仮名の有無（受注済／受注済み）と
    ' ひらがな書き（しょうだん中／商談中＝Excel のふりがなが同じ）を、ひらがなの少ない形 → 多く使われている形にそろえる。
    ' 英語の状態語（Active 等）は、決まった対応表で、列に既にある日本語（進行中 等）へそろえる（無ければ触らない）
    Set dv = CreateObject("Scripting.Dictionary")
    cnt = 0
    For i = hR + 1 To lastR
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
    For i = hR + 1 To lastR
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
    If Len(digits) = 0 Or digits Like "*[!0-9]*" Then Return
    If Len(digits) = 9 And Left(digits, 1) <> "0" Then digits = "0" & digits
    If Left(digits, 1) <> "0" Then Return
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
                m = (k + 2) \ 3: d = CLng(t): Y = CLng(parts(2))
                If d >= 1 And d <= 31 Then
                    dt = DateSerial(Y, m, d)
                    If Month(dt) = m And Day(dt) = d Then ok = True
                End If
            End If
            Return
        End If
    End If
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    If Len(t) = 8 And Not t Like "*[!0-9]*" And (Left(t, 2) = "19" Or Left(t, 2) = "20") Then
        Y = CLng(Left(t, 4)): m = CLng(Mid(t, 5, 2)): d = CLng(Right(t, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            dt = DateSerial(Y, m, d)
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
        Y = era + CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(0)) = 4 Then
        Y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        Y = CLng(parts(2)): m = CLng(parts(0)): d = CLng(parts(1))
    Else
        Return
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or Y < 1900 Or Y > 9999 Then Return
    dt = DateSerial(Y, m, d)
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
        If InStr("円個本枚台冊箱点件分歳", Right(t, 1)) > 0 Then t = Left(t, Len(t) - 1) Else Exit Do     ' 分＝休憩の「60分」・歳＝年齢の「35歳」（2026-10-08 Sonnet の採点 S3・S8）
    Loop
    u = ""
    If Right(t, 1) = "万" Then u = "万": t = Left(t, Len(t) - 1)
    If Len(t) = 0 Then Return
    If t Like "*[!0-9.+-]*" Then Return
    If InStr(2, t, "-") > 0 Or InStr(2, t, "+") > 0 Then Return
    If Len(t) > 1 And Left(t, 1) = "0" And Mid(t, 2, 1) <> "." Then Return
    If Len(Replace(Replace(t, ".", ""), "-", "")) > 15 Then Return
    ' 12 桁以上の数字だけの文字は番号（JAN 13 桁・法人番号・カード番号）＝数にしない（「4.90123E+12」になった・2026-10-08 Sonnet の採点 S7 H 列）
    If Len(t) >= 12 And Not t Like "*[!0-9]*" Then Return
    If Not IsNumeric(t) Then Return
    dbl = CDbl(t)
    If u = "万" Then dbl = dbl * 10000
    If bad Then dbl = -dbl
    ok = True
    Return
End Sub

Sub 全列が同じ重複行を削除する()
    ' 変えてよい: 行移動
    ' 依頼の語: 重複行を消|同じ行が二つ|同じ行が2つ|ダブっている行|ダブった行|重複データ|重複した行を1つ|重複した行をひとつ|二重に入っている行|重複行を削除|重複を消|重複を削除|重複を取り除|重複を除|重複データを消|重複をなくし|重複を無くし|ダブりを消|同じ行を消|同じ行を削除|重複している行を消|重複している行を削除|重複した行を消|重複した行を削除|重複行を取
    ' 依頼の組: 重複,ダブ,二重,かぶ,同じデータ,同じ行+消,削除,除,取り,なくし,無くし,1つに,ひとつに+-番号,キー,伝票,一覧にし,洗い出,マクロ,二重計上,二重払
    Dim ws As Worksheet, ur As Range, seen As Object
    Dim r0 As Long, c0 As Long, nr As Long, nc As Long, scanTo As Long
    Dim hR As Long, hc1 As Long, hc2 As Long, lastR As Long
    Dim i As Long, j As Long, k As Long, n As Long, best As Long
    Dim v As Variant, x As Variant, parts As Variant, dup() As Boolean
    Dim s As String, t As String, u As String, key As String, mon As String
    Dim ok As Boolean, bad As Boolean, dt As Date, dbl As Double
    Dim era As Long, Y As Long, m As Long, d As Long
    Dim hasDup As Boolean
    Dim oldScreenUpdating As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    Dim delRange As Range, batchStr As String
    Dim tblRange As Range, tblVal As Variant
    Dim nDataRows As Long, nCols As Long

    ' 重複＝全部の列が同じ行（書き方の違い＝空白・全角半角・大文字小文字・区切り・（株）・文字の日付と金額はならして比べる）。
    ' 1 列でも中身が違えば別の行として残す（備考だけ違う行も消さない）。先に出てくる行を残し、後の行を下から消す
    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    r0 = ur.row: c0 = ur.Column: nr = ur.rows.Count: nc = ur.Columns.Count
    If nr < 3 Then Exit Sub
    GoSub 見出しを探す
    If hR = 0 Or lastR <= hR + 1 Then Exit Sub

    nDataRows = lastR - hR
    nCols = hc2 - hc1 + 1
    Set tblRange = ws.Range(ws.Cells(hR + 1, hc1), ws.Cells(lastR, hc2))
    If nDataRows = 1 Then
        ReDim tblVal(1 To 1, 1 To nCols)
        Dim cIdx As Long
        For cIdx = 1 To nCols: tblVal(1, cIdx) = tblRange.Cells(1, cIdx).Value: Next cIdx
    Else
        tblVal = tblRange.Value2
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    ReDim dup(hR + 1 To lastR)
    hasDup = False

    For i = 1 To nDataRows
        key = ""
        For j = 1 To nCols
            v = tblVal(i, j)
            GoSub 照合の値
            key = key & Chr(1) & s
        Next
        If Len(Replace(key, Chr(1), "")) = 0 Then
            ' 途中の空行は重複と見ない（消さない）
        ElseIf seen.exists(key) Then
            dup(hR + i) = True
            hasDup = True
        Else
            seen.Add key, hR + i
        End If
    Next

    ' 事前高速判定: 重複が0件なら即終了（0秒判定）
    If Not hasDup Then Exit Sub

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 行ごと消す（アドレスバッチ結合による一撃削除）
    batchStr = ""
    For i = lastR To hR + 1 Step -1
        If dup(i) Then
            n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1))) _
                - Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, hc1), ws.Cells(i, hc2)))
            If n = 0 Then
                If Len(batchStr) > 200 Then
                    If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
                    batchStr = ""
                End If
                batchStr = batchStr & IIf(Len(batchStr) > 0, ",", "") & i & ":" & i
            End If
        End If
    Next i
    If Len(batchStr) > 0 Then
        If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
    End If

    If Not delRange Is Nothing Then delRange.Delete Shift:=xlUp

CleanUp:
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
    Exit Sub

見出しを探す:
    scanTo = r0 + IIf(nr > 20, 20, nr) - 1
    best = 0
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If n > best Then best = n
    Next
    hR = 0
    If best < 2 Then Return
    For i = r0 To scanTo
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, c0), ws.Cells(i, c0 + nc - 1)))
        If hR = 0 And n >= 2 And n >= best * 0.6 Then hR = i
    Next
    If hR = 0 Then Return
    hc1 = 0: hc2 = 0
    For j = c0 To c0 + nc - 1
        ok = Len(Trim(ws.Cells(hR, j).text)) > 0
        If Not ok And ws.Cells(hR, j).MergeCells Then ok = (ws.Cells(hR, j).MergeArea.row < hR)
        If ok Then
            If hc1 = 0 Then hc1 = j
            hc2 = j
        End If
    Next
    lastR = hR
    n = 0
    For i = hR + 1 To r0 + nr - 1
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
                m = (k + 2) \ 3: d = CLng(t): Y = CLng(parts(2))
                If d >= 1 And d <= 31 Then
                    dt = DateSerial(Y, m, d)
                    If Month(dt) = m And Day(dt) = d Then ok = True
                End If
            End If
            Return
        End If
    End If
    t = Replace(StrConv(Trim(Replace(v, "　", " ")), vbNarrow), " ", "")
    If Len(t) = 8 And Not t Like "*[!0-9]*" And (Left(t, 2) = "19" Or Left(t, 2) = "20") Then
        Y = CLng(Left(t, 4)): m = CLng(Mid(t, 5, 2)): d = CLng(Right(t, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            dt = DateSerial(Y, m, d)
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
        Y = era + CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(0)) = 4 Then
        Y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        Y = CLng(parts(2)): m = CLng(parts(0)): d = CLng(parts(1))
    Else
        Return
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or Y < 1900 Or Y > 9999 Then Return
    dt = DateSerial(Y, m, d)
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
    ' 変えてよい: 空欄
    ' 依頼の語: 合計行を足|下に合計|合計の行を追加|最終行に合計|合計の行を足|合計行を入|合計行を作|合計行を付|合計行を追加|列ごとの合計|縦計を|合計の行を入|合計の行をお願|最終行の下に列|下に出
    ' 依頼の組: 合計,縦計,総計+下,最終行,最後,末尾,縦の,行,欄+足,入れ,出,追加,付け,作,お願,ほし,欲し+-検算,正しい,合って,合うか,小計,区分,項目,月,ピボット,テーブル,右,した表,できない,凡例,グラフ
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
    urTop = ur.row
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
                If ws.Cells(headerRow + 1, c).MergeArea.row = headerRow Then h2Merge = True
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
    Do While sumRow = 0 And lastDataRow + 1 <= ur.row + ur.rows.Count - 1
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
           InStr(hs, "内線") > 0 Or InStr(hs, "電話") > 0 Or InStr(hs, "コード") > 0 Or _
           InStr(hs, "CODE") > 0 Or InStr(UCase$(hs), "ID") > 0 Or InStr(hs, "郵便") > 0 Or _
           hs = "年" Or hs = "月" Or hs = "日" Or InStr(hs, "日付") > 0 Or InStr(hs, "年月") > 0 Or InStr(hs, "生年") > 0 Or _
           InStr(hs, "年齢") > 0 Or InStr(hs, "年度") > 0 Or InStr(hs, "単価") > 0 Or _
           Right$(hs, 1) = "比" Or Right$(hs, 1) = "率" Or InStr(hs, "割合") > 0 Then   ' 2026-10-15 指標列の除外（レシピの試験）
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
    ' 選んでいるコード列の値を、対応表（マスタ）を参照して名称に置き換える。対応表にないコードはそのまま残す。
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
    Dim 変えた() As Boolean

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If ws.UsedRange Is Nothing Then Exit Sub

    selCol = Selection.Cells(1, 1).Column
    headerRow = Selection.Cells(1, 1).row
    Dim maxRow As Long
    maxRow = 表の本文の最終行(ws)
    If Selection.rows.Count > 1 Then ' 2026-10-15 選んだ範囲だけ（レシピの試験）
        lastRow = Selection.row + Selection.rows.Count - 1
        If lastRow > maxRow Then lastRow = maxRow
    Else
        lastRow = maxRow
    End If
    If lastRow <= headerRow Then Exit Sub

    Dim oldScreen As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim nData As Long
    nData = lastRow - headerRow

    Dim srcRange As Range
    Set srcRange = ws.Range(ws.Cells(headerRow + 1, selCol), ws.Cells(lastRow, selCol))
    Dim arrCode As Variant
    If nData = 1 Then
        ReDim arrCode(1 To 1, 1 To 1)
        arrCode(1, 1) = srcRange.Value2
    Else
        arrCode = srcRange.Value2
    End If

    ' サンプル値収集（コード列先頭10件）
    Dim sampleVals As Object
    Set sampleVals = CreateObject("Scripting.Dictionary")
    sampleVals.CompareMode = 1
    Dim sampleCount As Long
    sampleCount = 0
    For i = 1 To nData
        If セルの字(arrCode(i, 1)) <> "" Then
            s = CStr(arrCode(i, 1)): GoSub 正規化
            If Not sampleVals.exists(s) Then
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
        If tUR Is Nothing Then GoTo NextSheet
        If tUR.Columns.Count < 2 Then GoTo NextSheet
        Dim tFirstCol As Long, tLastCol As Long, tRows As Long
        tFirstCol = tUR.Column
        tLastCol = tUR.Column + tUR.Columns.Count - 1
        tRows = tUR.rows.Count
        ' 対応表は2列のみ対象（左がコード・右が名称）
        If tLastCol - tFirstCol <> 1 Then GoTo NextSheet

        Dim tArr As Variant
        If tRows = 1 Then
            ReDim tArr(1 To 1, 1 To 2)
            tArr(1, 1) = tUR.Cells(1, 1).Value2
            tArr(1, 2) = tUR.Cells(1, 2).Value2
        Else
            tArr = tUR.Value2
        End If

        Dim matchCount As Long
        matchCount = 0
        For i = 1 To tRows
            s = CStr(tArr(i, 1)): GoSub 正規化
            If sampleVals.exists(s) Then matchCount = matchCount + 1
        Next i
        If matchCount > bestMatch Then
            bestMatch = matchCount
            Set bestSheet = targetSheet
            bestKeyCol = tFirstCol
            bestValCol = tFirstCol + 1
        End If
NextSheet:
    Next targetSheet

    If bestMatch = 0 Or bestSheet Is Nothing Then
        Application.ScreenUpdating = True
        Set bestSheet = シート選び直しフォーム.選択("コード列に対応する名称のシート（マスタ）を選んでください", wb, ws.Name)
        Unload シート選び直しフォーム
        Application.ScreenUpdating = False
        If bestSheet Is Nothing Then
            Call 先撃ちに断る("対応表シートが選ばれませんでした")
            GoTo CleanUp
        End If
        bestKeyCol = bestSheet.UsedRange.Column
        bestValCol = bestKeyCol + 1
    End If

    ' 対応表から辞書構築
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim tUR2 As Range
    Set tUR2 = bestSheet.UsedRange
    Dim mRows As Long, mCols As Long
    mRows = tUR2.rows.Count
    mCols = tUR2.Columns.Count

    Dim mArr As Variant
    If mRows = 1 And mCols = 1 Then
        ReDim mArr(1 To 1, 1 To 1)
        mArr(1, 1) = tUR2.Value2
    Else
        mArr = tUR2.Value2
    End If

    Dim relKeyCol As Long, relValCol As Long
    relKeyCol = bestKeyCol - tUR2.Column + 1
    relValCol = bestValCol - tUR2.Column + 1

    For i = 1 To mRows
        Dim kv As String, vv As String
        kv = CStr(mArr(i, relKeyCol))
        vv = CStr(mArr(i, relValCol))
        If セルの字(kv) <> "" And セルの字(vv) <> "" Then
            s = kv: GoSub 正規化
            If Not dict.exists(s) Then dict.Add s, vv
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
        If Not nameSet.exists(nv) Then nameSet.Add nv, 1
    Next dk

    ' 置き換え（メモリ配列上で一括置換）
    Dim hasChange As Boolean
    hasChange = False
    ReDim 変えた(1 To nData, 1 To 1)
    For i = 1 To nData
        cellVal = CStr(arrCode(i, 1))
        If セルの字(cellVal) = "" Then GoTo NextCell
        If nameSet.exists(cellVal) Then GoTo NextCell
        s = cellVal: GoSub 正規化
        normalKey = s
        If dict.exists(normalKey) Then
            ' 数式のあるセルは触らない（保護）
            If Not ws.Cells(headerRow + i, selCol).HasFormula Then
                arrCode(i, 1) = dict(normalKey)
                変えた(i, 1) = True
                hasChange = True
            End If
        End If
NextCell:
    Next i

    ' 変えたセルだけ書く（未置換のコード・文字の数字 "00123"・数式セルを壊さない・2026-10-22）
    If hasChange Then
        Call 表の整理_下請け.変えたセルだけ書く(srcRange.Cells(1, 1), arrCode, 変えた, False)
    End If

CleanUp:
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
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
    ' 変えてよい: 空欄
    ' 選択範囲（または表全体）のセル結合を解除し、結合されていた全セルに元の左上の値を埋めて補完する。
    ' 依頼の語: 結合セルを解除|結合を外|結合をはずして|結合を解いて|結合をぜんぶ解除|マージを解除|結合セルをばら|全部解除
    ' 依頼の組: 結合+解除,外,解い,やめ,ばら+-一覧,調べ,クエリ,マスタ,キー
    ' 扱う: 結合解除 セル結合 マージ解除
    ' 見出し: なし
    ' 形: 数の列
    Dim ws As Worksheet
    Dim ur As Range
    Dim cell As Range
    Dim i As Long
    Dim cnt As Long
    Dim oldScreenUpdating As Boolean, oldEvents As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    ' 超高速事前判定: シート内に結合セルが1つも無ければ0秒で即終了（掟第10条）
    If ur.MergeCells = False Then Exit Sub

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim mergedAreas() As String
    Dim mergedValues() As Variant
    Dim isTitles() As Boolean
    ReDim mergedAreas(1 To 128)
    ReDim mergedValues(1 To 128)
    ReDim isTitles(1 To 128)

    cnt = 0
    For Each cell In ur
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                cnt = cnt + 1
                If cnt > UBound(mergedAreas) Then
                    ReDim Preserve mergedAreas(1 To cnt * 2)
                    ReDim Preserve mergedValues(1 To cnt * 2)
                    ReDim Preserve isTitles(1 To cnt * 2)
                End If
                Dim ma As Range
                Set ma = cell.MergeArea
                Dim isTitle As Boolean
                isTitle = False
                If ma.rows.Count = 1 And ma.Columns.Count > 1 Then
                    If ma.Cells(1, 1).Column = ur.Cells(1, 1).Column Then
                        isTitle = True
                    End If
                End If
                mergedAreas(cnt) = ma.Address
                mergedValues(cnt) = cell.Value
                isTitles(cnt) = isTitle
            End If
        End If
    Next cell

    If cnt = 0 Then GoTo CleanUp

    Dim mc As Range
    Dim val As Variant
    For i = 1 To cnt
        Set mc = ws.Range(mergedAreas(i))
        val = mergedValues(i)
        mc.UnMerge
        If Not isTitles(i) Then
            ' 文字の数字や日付が化けないよう文字のまま書く（2026-10-22）
            If VarType(val) = vbString Then
                mc.Value = 表の整理_下請け.文字のまま書く字(CStr(val))
            Else
                mc.Value = val
            End If
        Else
            If VarType(val) = vbString Then
                mc.Cells(1, 1).Value = 表の整理_下請け.文字のまま書く字(CStr(val))
            Else
                mc.Cells(1, 1).Value = val
            End If
        End If
    Next i

CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 選んだ列の空白を上の値で埋める()
    ' 変えてよい: 空欄
    ' 選んでいる列の空欄セルに、直上の有効なセルの値をコピーして埋める。伝票や一覧表の欠損補完に使う。
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
    Dim urTop As Long, urLeft As Long, urRight As Long, nCols As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim startRow As Long
    Dim r As Long
    Dim c As Long
    Dim aCol As Long
    Dim cc As Long
    Dim lastVal As Variant
    Dim maxRow As Long
    Dim 表値 As Variant, 行スキップ() As Boolean
    Dim 列値 As Variant, 列式 As Variant, colRange As Range
    Dim nRows As Long, idx As Long, tblR As Long, 変えた() As Boolean
    Dim rowEmpty As Boolean, isTotal As Boolean, colChanged As Boolean, 式あり As Boolean

    Set ws = ActiveSheet
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    nCols = urRight - urLeft + 1
    maxRow = 表の本文の最終行(ws)
    If maxRow < urTop Then Exit Sub

    ' 表全体の値を一括取得し、各行の空行・集計行判定を事前計算
    If maxRow = urTop And nCols = 1 Then
        ReDim 表値(1 To 1, 1 To 1)
        表値(1, 1) = ws.Cells(urTop, urLeft).Value
    Else
        表値 = ws.Range(ws.Cells(urTop, urLeft), ws.Cells(maxRow, urRight)).Value
    End If

    ReDim 行スキップ(urTop To maxRow)
    For r = urTop To maxRow
        tblR = r - urTop + 1
        rowEmpty = True
        isTotal = False
        For cc = 1 To nCols
            If セルの字(表値(tblR, cc)) <> "" Then
                rowEmpty = False
            End If
            If VarType(表値(tblR, cc)) = vbString Then
                If 集計の語か(CStr(表値(tblR, cc))) Then
                    isTotal = True
                End If
            End If
        Next cc
        行スキップ(r) = (rowEmpty Or isTotal)
    Next r

    Application.ScreenUpdating = False

    For Each area In sel.Areas
        headerRow = area.row
        If area.rows.Count > 1 Then ' 2026-10-15 選んだ範囲だけ（レシピの試験）
            lastRow = area.row + area.rows.Count - 1
            If lastRow > maxRow Then lastRow = maxRow
        Else
            lastRow = maxRow
        End If
        startRow = headerRow + 1
        If startRow <= lastRow Then
            nRows = lastRow - startRow + 1
            For aCol = 1 To area.Columns.Count
                c = area.Columns(aCol).Column
                lastVal = Empty
                Set colRange = ws.Range(ws.Cells(startRow, c), ws.Cells(lastRow, c))
                If nRows > 1 Then
                    列値 = colRange.Value
                    列式 = colRange.Formula
                Else
                    ReDim 列値(1 To 1, 1 To 1): ReDim 列式(1 To 1, 1 To 1)
                    列値(1, 1) = colRange.Value: 列式(1, 1) = colRange.Formula
                End If

                colChanged = False
                式あり = False
                ReDim 変えた(1 To nRows, 1 To 1)
                For idx = 1 To nRows
                    If Left$(CStr(列式(idx, 1)), 1) = "=" Then 式あり = True
                    r = startRow + idx - 1
                    If r >= urTop And r <= maxRow Then
                        If Not 行スキップ(r) Then
                            ' 上の値がエラー値なら埋めない（比べると「型が一致しません」で止まった・2026-09-24 総点検）
                            If セルの字(列値(idx, 1)) = "" And Not ws.Cells(r, c).MergeCells Then
                                If Not isEmpty(lastVal) Then
                                    If Not IsError(lastVal) Then
                                        If セルの字(lastVal) <> "" Then
                                            列値(idx, 1) = lastVal      ' 文字の数字・式に見える文字は書くときに文字のまま書く（変えたセルだけ書く）
                                            変えた(idx, 1) = True
                                            colChanged = True
                                        End If
                                    End If
                                End If
                            Else
                                lastVal = 列値(idx, 1)
                            End If
                        End If
                    End If
                Next idx

                ' 埋めたセルだけ書く（列ごと書き戻すと、文字の数字 "00123" が数に化け、式が値に潰れる・2026-10-22）
                If colChanged Then 変えたセルだけ書く colRange.Cells(1, 1), 列値, 変えた
            Next aCol
        End If
    Next area

    Application.ScreenUpdating = True
End Sub

Sub 選んだ列の文字の日付を日付にする()
    ' 変えてよい: 文字日付
    ' 依頼の語: 文字の日付を|和暦の文字|日付として認識|受付日の書|入っている日付|日付列|日付が文字列|列を日付の値|文字列の日付|日付データに変換|日付データに
    ' 依頼の組: 日付+文字,和暦,認識,日付型,日付値,日付に+-異常,おかしい,間違い,ミス,妥当,年度,四半期,未来,ありえない,あり得ない
    ' 扱う: 文字日付変換
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる列の文字列日付を日付値に変換し表示形式をyyyy/m/dにする

    Dim ws As Worksheet
    Dim s As String
    Dim parsed As Date
    Dim parseErr As Boolean
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim area As Range
    Dim col As Long
    Dim colIdx As Long
    Dim r As Long
    Dim okNum As Boolean, v数 As Double
    Dim maxRow As Long
    Dim oldScreenUpdating As Boolean, oldEvents As Boolean
    Dim colRange As Range, arr As Variant, nRows As Long
    Dim colChanged As Boolean, changed() As Boolean

    Set ws = ActiveSheet
    If ws.UsedRange Is Nothing Then Exit Sub
    maxRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each area In Selection.Areas
        hdrRow = area.row
        If area.rows.Count > 1 Then ' 2026-10-15 選んだ範囲だけ（レシピの試験）
            lastRow = area.row + area.rows.Count - 1
            If lastRow > maxRow Then lastRow = maxRow
        Else
            lastRow = maxRow
        End If
        If lastRow > hdrRow Then
            nRows = lastRow - hdrRow
            For col = 1 To area.Columns.Count
                colIdx = area.Column + col - 1
                Set colRange = ws.Range(ws.Cells(hdrRow + 1, colIdx), ws.Cells(lastRow, colIdx))
                If nRows = 1 Then
                    ReDim arr(1 To 1, 1 To 1)
                    arr(1, 1) = colRange.Value2
                Else
                    arr = colRange.Value2
                End If
                ReDim changed(1 To nRows, 1 To 1)

                ' 8 桁の「数値」（20260615 が数のもの。基幹システムの出力に多い）は、列の数値が全部 8 桁の日付の形のときだけ日付にする
                ' （金額の列を日付にしない）。数値が無い列・日付の形でない数値のある列は、数値には触らない
                okNum = True: v数 = 0
                For r = 1 To nRows
                    If VarType(arr(r, 1)) = vbDouble Then
                        v数 = arr(r, 1)
                        If v数 <> Int(v数) Or v数 < 19000101 Or v数 > 21001231 Then
                            okNum = False: Exit For
                        End If
                    End If
                Next r

                colChanged = False
                For r = 1 To nRows
                    If isEmpty(arr(r, 1)) Then GoTo NextRow
                    If VarType(arr(r, 1)) = vbDate Then GoTo NextRow
                    If VarType(arr(r, 1)) = vbDouble Then
                        If Not okNum Then GoTo NextRow
                    ElseIf VarType(arr(r, 1)) <> vbString Then
                        GoTo NextRow
                    End If

                    s = Trim(CStr(arr(r, 1)))
                    If セルの字(s) = "" Then GoTo NextRow
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
                        ' 数式のあるセルは触らない（保護）
                        If Not ws.Cells(hdrRow + r, colIdx).HasFormula Then
                            arr(r, 1) = parsed
                            changed(r, 1) = True
                            colChanged = True
                        End If
                    End If
NextRow:
                Next r

                ' 変更があったセルだけ書き戻す（元から日付・日時のセルの書式は触らない・2026-10-22）
                If colChanged Then
                    Call 表の整理_下請け.変えたセルだけ書く(colRange.Cells(1, 1), arr, changed, False, "yyyy/m/d")
                End If
            Next col
        End If
    Next area

CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 選んだ列の文字の数字を数にする()
    ' 変えてよい: 文字数字
    ' 選んでいる列の文字列として保存されている数字（カンマや空白付き含む）を計算可能な数値に変換する。先頭ゼロのコード列は触らない。
    ' 依頼の語: 数値に直|文字列の数字|数値に変換|金額が文字列|文字の数字を数値|計算できる数値|数に直|数値に変|文字を数値|カンマ付|数値をカンマ区切り|数値のカンマ区切り
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
    Dim s As String
    Dim CHK As String
    Dim v As Double
    Dim neg As Boolean
    Dim allDigit As Boolean
    Dim ci As Integer
    Dim ch As String
    Dim fc As String
    Dim colRange As Range, arr As Variant, nRows As Long
    Dim colChanged As Boolean
    Dim oldScreenUpdating As Boolean, oldEvents As Boolean
    Dim 頭文字 As String
    Dim 変えた() As Boolean

    Set ws = ActiveSheet
    Set sel = Selection
    If ws.UsedRange Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    headerRow = sel.Cells(1, 1).row

    For Each area In sel.Areas
      For colIdx = area.Column To area.Column + area.Columns.Count - 1
        Dim maxR As Long
        maxR = ws.Cells(ws.rows.Count, colIdx).End(xlUp).row
        If area.rows.Count > 1 Then ' 2026-10-15 選んだ範囲だけ（レシピの試験）
            lastR = area.row + area.rows.Count - 1
            headerRow = area.row
            If lastR > maxR Then lastR = maxR
        Else
            lastR = maxR
        End If

        If lastR > headerRow Then
            Dim headerText As String
            headerText = CStr(ws.Cells(headerRow, colIdx).Value)
            If InStr(headerText, "コード") > 0 Or InStr(headerText, "番号") > 0 Or _
               InStr(headerText, "TEL") > 0 Or InStr(headerText, "tel") > 0 Or _
               InStr(headerText, "電話") > 0 Or InStr(headerText, "ZIP") > 0 Or _
               InStr(headerText, "郵便") > 0 Then
                ' コード列・電話番号列は数値化しない（先頭ゼロ保護・055）
                GoTo NextCol
            End If

            nRows = lastR - headerRow
            Set colRange = ws.Range(ws.Cells(headerRow + 1, colIdx), ws.Cells(lastR, colIdx))
            If nRows = 1 Then
                ReDim arr(1 To 1, 1 To 1)
                arr(1, 1) = colRange.Value2
            Else
                arr = colRange.Value2
            End If

            ' 代表表示形式の事前取得（COM走査は最小限・見つかれば即終了）
            fc = "General"
            Dim sampleRow As Long
            For sampleRow = headerRow + 1 To lastR
                If VarType(arr(sampleRow - headerRow, 1)) = vbDouble Then
                    fc = ws.Cells(sampleRow, colIdx).NumberFormat
                    If fc <> "@" And fc <> "" Then Exit For
                End If
            Next sampleRow
            If fc = "@" Or fc = "" Then fc = "General"

            colChanged = False
            ReDim 変えた(1 To nRows, 1 To 1)
            For r = 1 To nRows
                If isEmpty(arr(r, 1)) Then GoTo NextRow
                Select Case VarType(arr(r, 1))
                    Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal
                        GoTo NextRow
                End Select
                If VarType(arr(r, 1)) <> vbString Then GoTo NextRow

                s = CStr(arr(r, 1)): GoSub 正規化

                If セルの字(s) = "" Then GoTo NextRow
                If s = "-" Then GoTo NextRow

                CHK = s
                If Left(CHK, 1) = "-" Then CHK = Mid(CHK, 2)
                If Len(CHK) >= 2 And Left(CHK, 1) = "0" Then
                    allDigit = True
                    For ci = 1 To Len(CHK)
                        If Mid(CHK, ci, 1) < "0" Or Mid(CHK, ci, 1) > "9" Then
                            allDigit = False: Exit For
                        End If
                    Next ci
                    If allDigit Then GoTo NextRow
                End If

                CHK = s
                If Left(CHK, 1) = "-" Then CHK = Mid(CHK, 2)
                If セルの字(CHK) = "" Then GoTo NextRow
                allDigit = True
                For ci = 1 To Len(CHK)
                    ch = Mid(CHK, ci, 1)
                    If (ch < "0" Or ch > "9") And ch <> "." Then
                        allDigit = False: Exit For
                    End If
                Next ci
                If Not allDigit Then GoTo NextRow
                If Len(CHK) - Len(Replace(CHK, ".", "")) > 1 Or Not CHK Like "*#*" Then GoTo NextRow
                If Not IsNumeric(s) Then GoTo NextRow

                v = CDbl(s)
                If Not ws.Cells(headerRow + r, colIdx).HasFormula Then
                    arr(r, 1) = v
                    変えた(r, 1) = True
                    colChanged = True
                End If
NextRow:
            Next r

            ' 数にしたセルだけ書く（列ごと書き戻すと、数式が値に潰れ、先頭ゼロの文字 "00123" が数に化けた・2026-10-22）
            If colChanged Then 変えたセルだけ書く colRange.Cells(1, 1), arr, 変えた, False, fc
        End If
NextCol:
      Next colIdx
    Next area

CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
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
        頭文字 = Left(s, 1)
        If 頭文字 = "△" Or 頭文字 = "▲" Or 頭文字 = "-" Then
            neg = True
            s = Mid(s, 2)
            s = Trim$(s)
        End If
    End If
    If Len(s) > 0 Then
        頭文字 = Left(s, 1)
        If 頭文字 = "△" Or 頭文字 = "▲" Or 頭文字 = "-" Then
            neg = True
            s = Mid(s, 2)
            s = Trim$(s)
        End If
    End If
    If neg And セルの字(s) <> "" Then s = "-" & s
    Return

End Sub

Sub 表の中の空行を削除して詰める()
    ' 変えてよい: 行移動
    ' 依頼の語: 空行を消|空白行を消|空白行を削除|空白の行を|明細の間の空|行の抜けを詰め|空行を取り除|空行が無ければ|空行を詰|空白行を詰|空いた行を詰|空行を削除|空白の行を詰|詰めて|詰める
    ' 依頼の組: 空行,空白行,空白の行,空いた行,空の行+削除,消,詰,除+-列,番号,連番
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
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)
    Dim maC As Long
    For maC = urLeft To urRight
        If ws.Cells(urBottom, maC).MergeCells Then
            Dim maBottom As Long
            maBottom = ws.Cells(urBottom, maC).MergeArea.row + ws.Cells(urBottom, maC).MergeArea.rows.Count - 1
            If maBottom > urBottom Then urBottom = maBottom
        End If
    Next maC
    Dim maxR As Long
    maxR = ur.row + ur.rows.Count - 1
    Do While urBottom < maxR
        Dim nextEmpty As Boolean
        nextEmpty = True
        For maC = urLeft To urRight
            If セルの字(ws.Cells(urBottom + 1, maC).Value) <> "" Then
                nextEmpty = False
                Exit For
            End If
        Next maC
        If nextEmpty Then
            urBottom = urBottom + 1
        Else
            Exit Do
        End If
    Loop

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

    ' 表の下の注記（空行の下の、文字だけで 1 行に 1 セルの行＝「※ 単価は税抜」など）は表の外。注記と、その手前の空行は消さない
    Dim 注r0 As Long, 注r1 As Long, 注ok As Boolean, 注n As Long
    Do
        注r1 = urBottom
        Do While 注r1 > hdrRow And Application.WorksheetFunction.CountA(ws.Range(ws.Cells(注r1, urLeft), ws.Cells(注r1, urRight))) = 0
            注r1 = 注r1 - 1
        Loop
        注r0 = 注r1
        Do While 注r0 - 1 > hdrRow And Application.WorksheetFunction.CountA(ws.Range(ws.Cells(注r0 - 1, urLeft), ws.Cells(注r0 - 1, urRight))) > 0
            注r0 = 注r0 - 1
        Loop
        If 注r0 - 1 <= hdrRow + 1 Then Exit Do
        注ok = True
        For r = 注r0 To 注r1
            注n = 0
            For c = urLeft To urRight
                If セルの字(ws.Cells(r, c).Value) <> "" Then
                    注n = 注n + 1
                    If ws.Cells(r, c).HasFormula Or VarType(ws.Cells(r, c).Value) <> vbString Or 集計の語か(ws.Cells(r, c).Value) Then 注ok = False
                End If
            Next c
            If 注n > 1 Then 注ok = False
        Next r
        If Not 注ok Then Exit Do
        urBottom = 注r0 - 1
        Do While urBottom > hdrRow And Application.WorksheetFunction.CountA(ws.Range(ws.Cells(urBottom, urLeft), ws.Cells(urBottom, urRight))) = 0
            urBottom = urBottom - 1
        Loop
    Loop
    ' 表の最後の行の下に続く空行は「表の中の空行」ではない
    Do While urBottom > hdrRow And Application.WorksheetFunction.CountA(ws.Range(ws.Cells(urBottom, urLeft), ws.Cells(urBottom, urRight))) = 0
        urBottom = urBottom - 1
    Loop

    If urBottom <= hdrRow Then Exit Sub

    ' 見出し行の左端・右端
    Dim colLeft As Long, colRight As Long
    colLeft = urLeft
    colRight = urRight

    ' 二次元配列でメモリ走査し空行を判定
    Dim tblRange As Range, tblVal As Variant
    Set tblRange = ws.Range(ws.Cells(hdrRow + 1, colLeft), ws.Cells(urBottom, colRight))

    ' 057: 帳票構造（データ行の縦結合）の検知・安全停止ガード
    Dim hasVerticalMerge As Boolean, chkR As Long, chkC As Long
    hasVerticalMerge = False
    For chkR = hdrRow + 1 To urBottom
        For chkC = colLeft To colRight
            If ws.Cells(chkR, chkC).MergeCells Then
                If ws.Cells(chkR, chkC).MergeArea.rows.Count > 1 Then
                    hasVerticalMerge = True
                    Exit For
                End If
            End If
        Next chkC
        If hasVerticalMerge Then Exit For
    Next chkR

    If hasVerticalMerge Then
        Call 表の整理_下請け.帳票安全停止フラグ(True)
        Application.StatusBar = "表の中の空行を削除して詰める: 帳票構造（縦結合）を検知したため、空行削除を安全停止しました。"
        Exit Sub
    Else
        Call 表の整理_下請け.帳票安全停止フラグ(False)
    End If
    Dim nRows As Long, nCols As Long
    nRows = urBottom - hdrRow
    nCols = colRight - colLeft + 1
    If nRows = 1 Then
        ReDim tblVal(1 To 1, 1 To nCols)
        Dim cIdx As Long
        For cIdx = 1 To nCols: tblVal(1, cIdx) = tblRange.Cells(1, cIdx).Value: Next cIdx
    Else
        tblVal = tblRange.Value2
    End If

    Dim delFlags() As Boolean
    ReDim delFlags(1 To nRows)
    Dim hasEmptyRow As Boolean
    hasEmptyRow = False

    For r = 1 To nRows
        Dim allEmpty As Boolean
        allEmpty = True
        For c = 1 To nCols
            If セルの字(tblVal(r, c)) <> "" Then
                allEmpty = False
                Exit For
            End If
        Next c
        If allEmpty Then
            delFlags(r) = True
            hasEmptyRow = True
        End If
    Next r

    ' 事前高速判定: 空行が0件なら即復帰（0秒判定）
    If Not hasEmptyRow Then Exit Sub

    Dim oldScreenUpdating As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' アドレスバッチ結合による一撃削除
    Dim delRange As Range, batchStr As String
    batchStr = ""
    For r = 1 To nRows
        If delFlags(r) Then
            Dim realRow As Long
            realRow = hdrRow + r
            ' 結合セルがあれば結合解除
            For maC = colLeft To colRight
                If ws.Cells(realRow, maC).MergeCells Then
                    ws.Cells(realRow, maC).MergeArea.UnMerge
                End If
            Next maC
            If Len(batchStr) > 200 Then
                If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
                batchStr = ""
            End If
            batchStr = batchStr & IIf(Len(batchStr) > 0, ",", "") & realRow & ":" & realRow
        End If
    Next r
    If Len(batchStr) > 0 Then
        If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
    End If

    ' たった1回の一撃削除
    If Not delRange Is Nothing Then delRange.Delete Shift:=xlUp

CleanUp:
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 二段の見出しを一行に畳む()
    ' 上下2行に分かれている二段見出しを、上段と下段の文字を連結して1行の見出しに統合する。
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
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.row + ur.rows.Count - 1

    Dim hrow As Long
    hrow = 0
    Dim r As Long, c As Long, cnt As Long
    For r = urTop To urBottom
        cnt = 0
        For c = urLeft To urRight
            Dim ma0 As Range
            Set ma0 = ws.Cells(r, c).MergeArea
            If ma0.Cells(1, 1).row = r And ma0.Cells(1, 1).Column = c Then
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
    Dim 縦結合() As Boolean
    ReDim 縦結合(urLeft To urRight)
    Dim lastTop As String
    lastTop = ""
    For c = urLeft To urRight
        Dim maC As Range
        Set maC = ws.Cells(row1, c).MergeArea
        ' 上下 2 段に縦に結合した見出し（整理番号・備考）は、上の段の書式を下の段へ写す（下の段は塗りも太字も無いので、
        ' 上の段を消すと見出しの色が抜けた・2026-10-09 shu「表題の色とかついてないところある」テスト用5 A7・F7）
        If maC.rows.Count >= 2 And maC.row = row1 Then 縦結合(c) = True
        Dim topVal As String
        topVal = Trim(CStr(maC.Cells(1, 1).Value))
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
    For c = urLeft To urRight
        If 縦結合(c) Then ws.Cells(row1, c).Copy ws.Cells(row2, c)     ' 書式ごと写す（字は下で書き直す）
    Next c

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
    selRow = Selection.Cells(1, 1).row

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
    Dim mT() As Long, mB() As Long, mL() As Long, mR() As Long, mV() As Variant, nm As Long, ma As Range, mm As Long
    ReDim mT(1 To 1): ReDim mB(1 To 1): ReDim mL(1 To 1): ReDim mR(1 To 1): ReDim mV(1 To 1)
    For r3 = hdrTop To hdrBot
        For j = leftCol To rightCol
            If ws.Cells(r3, j).MergeCells Then
                Set ma = ws.Cells(r3, j).MergeArea
                If ma.row = r3 And ma.Column = j And ma.row + ma.rows.Count - 1 <= hdrBot _
                   And ma.Column + ma.Columns.Count - 1 <= rightCol Then
                    nm = nm + 1
                    ReDim Preserve mT(1 To nm): ReDim Preserve mB(1 To nm): ReDim Preserve mL(1 To nm)
                    ReDim Preserve mR(1 To nm): ReDim Preserve mV(1 To nm)
                    mT(nm) = ma.row: mB(nm) = ma.row + ma.rows.Count - 1
                    mL(nm) = ma.Column: mR(nm) = ma.Column + ma.Columns.Count - 1
                    mV(nm) = ma.Cells(1, 1).Formula
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
    If 写し直し And nm > 0 Then                 ' 写し直しに回るなら、見出しの結合は元の位置に戻す
        Application.DisplayAlerts = False
        For mm = 1 To nm
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
        If nm > 0 Then                          ' 見出しの結合を、列の新しい位置で結合し直す（組の列は隣り合っている）
            Dim posOf() As Long, a1 As Long, a2 As Long
            ReDim posOf(leftCol To rightCol)
            For t = 1 To nCols
                posOf(newOrder(t)) = leftCol + t - 1
            Next t
            Application.DisplayAlerts = False
            For mm = 1 To nm
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
    ' 選んでいる列で途中のセルが値で上書きされたり消えている数式を、上下のセルと同じ計算式で復元・統一する。
    ' 依頼の語: 式が途切れ|式の途切れ|値で上書き|計算式が崩れ|上下と同じ式に戻|式が抜|列の数式|計算式がエラーになって|数式がエラーになって|計算式のエラーを修正|数式のエラーを修正|計算式のエラーを直|数式のエラーを直|式がエラーになる原因|エラーになっている数式|計算式を修正|計算式がエラー出て|数式がエラー出て
    ' 依頼の組: 式,数式,計算式+消えて,途切れ,抜け,崩れ,そろえ,揃え,戻+-一覧,監査,調べ,説明,株式,形式,様式,書式,正式
    ' 扱う: 式揃え 数式補完 合計
    ' 変えてよい: 空欄 数式
    ' 見出し: なし
    ' 選ぶ列: -
    ' 形: なし

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim hdrRow As Long
    hdrRow = Selection.Cells(1, 1).row

    Dim maxRow As Long
    maxRow = 表の本文の最終行(ws)

    ' 頼みの文の取得（2026-10-24 指示書069）
    Dim 頼み As String
    On Error Resume Next
    頼み = コンボ道具.頼みの文を返す()
    On Error GoTo 0

    ' 対象列の自動走査フォールバック判定
    Dim needFullScan As Boolean
    needFullScan = False

    If Selection.Cells.CountLarge = 1 Then
        Dim selColHasFormula As Boolean
        selColHasFormula = False
        Dim chkR As Long
        For chkR = hdrRow + 1 To maxRow
            If ws.Cells(chkR, Selection.Column).HasFormula Then
                selColHasFormula = True
                Exit For
            End If
        Next chkR

        If Not selColHasFormula Then
            needFullScan = True
        ElseIf Len(頼み) > 0 Then
            If InStr(頼み, "エラー") > 0 Or InStr(頼み, "原因") > 0 Then
                Dim selColHasError As Boolean
                selColHasError = False
                For chkR = hdrRow + 1 To maxRow
                    If IsError(ws.Cells(chkR, Selection.Column).Value) Then
                        selColHasError = True
                        Exit For
                    End If
                Next chkR
                If Not selColHasError Then needFullScan = True
            End If
        End If
    End If

    Dim targetCols() As Long, numCols As Long
    Dim area As Range, c As Long

    Dim skipAddrs As String: skipAddrs = ""
    If needFullScan Then
        hdrRow = ws.UsedRange.row
        Dim uColStart As Long, uColEnd As Long
        uColStart = ws.UsedRange.Column
        uColEnd = uColStart + ws.UsedRange.Columns.Count - 1
        ReDim targetCols(1 To ws.UsedRange.Columns.Count)
        numCols = 0
        For c = uColStart To uColEnd
            numCols = numCols + 1
            targetCols(numCols) = c
        Next c
    Else
        ReDim targetCols(1 To Selection.CountLarge)
        numCols = 0
        For Each area In Selection.Areas
            For c = area.Column To area.Column + area.Columns.Count - 1
                numCols = numCols + 1
                targetCols(numCols) = c
            Next c
        Next area
    End If

    Dim colIdx As Long, col As Long
    For colIdx = 1 To numCols
        col = targetCols(colIdx)

        ' 本文の範囲
        Dim lastRow As Long
        If Not needFullScan And Selection.Areas.Count = 1 And Selection.rows.Count > 1 Then
            lastRow = Selection.row + Selection.rows.Count - 1
            hdrRow = Selection.row
            If lastRow > maxRow Then lastRow = maxRow
        Else
            lastRow = maxRow
        End If

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
            Dim found As Boolean
            found = 集計の行か(ws, r, ws.UsedRange.Column, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)
            If found Then GoTo NextRow1

            ' 式を記録（IsError のセルは多数派手本に数えない・指示書069）
            Dim cellVal As Range
            Set cellVal = ws.Cells(r, col)
            If cellVal.HasFormula And Not IsError(cellVal.Value) Then
                Dim f As String
                f = cellVal.FormulaR1C1
                If dict.exists(f) Then
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

        ' 式でないセル・エラーセル・異なる式に適用
        For r = hdrRow + 1 To lastRow
            Set rowRange = ws.rows(r)
            If WorksheetFunction.CountA(rowRange) = 0 Then GoTo NextRow2

            found = 集計の行か(ws, r, ws.UsedRange.Column, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)
            If found Then GoTo NextRow2

            Set cellVal = ws.Cells(r, col)

            ' エラーセルは手本式で修復
            If IsError(cellVal.Value) Then
                cellVal.FormulaR1C1 = bestFormula
                GoTo NextRow2
            End If

            ' 空セルは触らない
                If セルの字(cellVal.Value) = "" And Not cellVal.HasFormula Then GoTo NextRow2

                If Not cellVal.HasFormula Then
                    If セルの字(cellVal.Value) <> "" And Not IsNumeric(cellVal.Value) And Not IsDate(cellVal.Value) Then GoTo NextRow2
                    Dim origCellVal As Variant: origCellVal = cellVal.Value
                    cellVal.FormulaR1C1 = bestFormula
                    Dim sameAns As Boolean: sameAns = False
                    If Not IsError(cellVal.Value) Then
                        If CStr(cellVal.Value) = CStr(origCellVal) Then sameAns = True
                    End If
                    If Not sameAns Then
                        ws.Cells(r, col).Value = origCellVal
                        If Len(skipAddrs) < 80 Then skipAddrs = skipAddrs & IIf(Len(skipAddrs) > 0, "・", "") & cellVal.Address(False, False)
                    End If
                ElseIf cellVal.FormulaR1C1 <> bestFormula Then
                    If Len(skipAddrs) < 80 Then skipAddrs = skipAddrs & IIf(Len(skipAddrs) > 0, "・", "") & cellVal.Address(False, False)
                End If
NextRow2:
        Next r
NextCol:
    Next colIdx

    If Len(skipAddrs) > 0 Then Application.StatusBar = "手本と違う値・式はそのまま残しました: " & skipAddrs
    ' 列の手本で直らないエラーの式はそのまま残す（seiri の残りに出る＝AI と人が直す）。見出しの語から式を作る・
    ' 税率を決め打ちする・IFERROR で隠すは入れない（2026-10-27 番頭の点検で外した。業務の式を推し量らない＝2026-10-08 と同じ決まり）
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
    Dim hdrRow As Long, tblLastCol As Long, tblLastRow As Long
    Dim c As Long, r As Long
    Dim newLastCol As Long, newLastRow As Long
    Dim hasVal As Boolean

    Set ws = ActiveSheet

    Dim i As Integer
    For i = 1 To ws.ListObjects.Count
        Set lo = ws.ListObjects(i)

        ' --- 右に広げる列を決める ---
        hdrRow = lo.HeaderRowRange.row
        tblLastCol = lo.Range.Column + lo.Range.Columns.Count - 1
        newLastCol = tblLastCol
        c = tblLastCol + 1
        Do While セルの字(ws.Cells(hdrRow, c).Value) <> ""
            newLastCol = c
            c = c + 1
        Loop

        ' --- 下に広げる行を決める ---
        tblLastRow = lo.Range.row + lo.Range.rows.Count - 1
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
    ' 表の中でデータが1件も入っていない完全な空列を検出し、削除して左へ詰める。
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
    For r = ur.row To ur.row + ur.rows.Count - 1
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
    For r = ur.row + ur.rows.Count - 1 To hdrRow Step -1
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

    ' 表範囲の値を二次元配列で一括取得
    Dim tblRange As Range, tblVal As Variant
    Dim nRows As Long, nCols As Long
    nRows = lastRow - hdrRow + 1
    nCols = lastCol - firstCol + 1
    Set tblRange = ws.Range(ws.Cells(hdrRow, firstCol), ws.Cells(lastRow, lastCol))
    If nRows = 1 Then
        ReDim tblVal(1 To 1, 1 To nCols)
        Dim cIdx As Long
        For cIdx = 1 To nCols: tblVal(1, cIdx) = tblRange.Cells(1, cIdx).Value: Next cIdx
    Else
        tblVal = tblRange.Value2
    End If

    Dim delFlags() As Boolean, hasDel As Boolean
    ReDim delFlags(firstCol To lastCol)
    hasDel = False

    ' 見出しも本文も全部空の列を判定
    For c = lastCol To firstCol Step -1
        Dim relC As Long
        relC = c - firstCol + 1
        Dim isEmpty As Boolean
        isEmpty = True
        For r = 1 To nRows
            If セルの字(tblVal(r, relC)) <> "" Then
                isEmpty = False
                Exit For
            End If
        Next r
        ' 表の外（表題・下のメモ）に字がある列は消さない
        If isEmpty Then
            If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then
                delFlags(c) = True
                hasDel = True
            End If
        End If
    Next c

    ' 事前高速判定: 空列が0件なら即終了（0秒判定）
    If Not hasDel Then Exit Sub

    Dim oldScreenUpdating As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim delColRange As Range
    For c = lastCol To firstCol Step -1
        If delFlags(c) Then
            If delColRange Is Nothing Then
                Set delColRange = ws.Columns(c)
            Else
                Set delColRange = Union(delColRange, ws.Columns(c))
            End If
        End If
    Next c

    If Not delColRange Is Nothing Then delColRange.EntireColumn.Delete

CleanUp:
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 乱れた列を掃除し右に足す()
    ' 依頼の語: データの掃除|クレンジング|データを掃除|取り込んだデータが汚|表記がバラバラ|データをきれいに|汚いデータ|文字列の乱れ|表記ゆれを直
    ' 依頼の組: データ+掃除,汚,きたな,クレンジング,きれい+-CSV
    ' 階級: 基本
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

    urTop = ur.row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.row + ur.rows.Count - 1

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
    ' 変えてよい: 文字日付 文字数字
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
    urTop = ur.row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.row + ur.rows.Count - 1

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
    ' 依頼の語: 体裁を揃え|体裁をそろえ|見た目をそろえ|列幅をそろえ|全シートの見た目|シートごとにバラバラ|同じ形のシート|表示倍率をそろえ|どのシートも同じ
    ' 依頼の組: 全シート,全部のシート,シートごと,どのシート,他のシート,ほかのシート,各シート+列幅,倍率,見た目,体裁,そろえ,揃え,合わせ+-グラフ,ピボット,分け,集約,まとめ
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
    urTop = ur.row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = ur.row + ur.rows.Count - 1

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
        If Err.Number <> 0 Then
            ' 保存してあるブックで控えが作れないときは切らない（式が値に変わって戻せない・ブックの式を全部値にすると同じ・2026-10-28 点検）
            Err.Clear
            On Error GoTo 0
            sh.Cells(outRow, 1).Value = "結果"
            sh.Cells(outRow, 2).Value = "ブック全体"
            sh.Cells(outRow, 3).Value = "控えが作れなかったので、リンクは切っていません（" & 控え & "）"
            Application.ScreenUpdating = True
            Application.StatusBar = "外部リンクの解消: 控えが作れなかったので止めました（" & 控え & "）。"
            Exit Sub
        End If
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
    Dim nm As Name, 参照 As String, 名前 As String, 短名 As String
    Dim 全式 As String, i As Long, p As Long, 使用 As Boolean
    Dim outRow As Long, 名 As String, n壊れ As Long, n未使用 As Long, n全 As Long, 消した As Long
    Dim 前後 As String, j As Long, n外部 As Long

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
    For Each nm In wb.names
        On Error Resume Next
        If Len(全式) < 900000 Then 全式 = 全式 & " " & nm.RefersTo
        On Error GoTo 0
    Next nm

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
        Set nm = wb.names(i)
        名前 = "": 参照 = ""
        On Error Resume Next
        名前 = nm.Name
        参照 = nm.RefersTo
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

        ' 他のブックを指す名前か（[book.xlsx]Sheet!A1 の形、または [1]Sheet!A1 の外部リンクの番号の形。表の構造化参照 Table1[列] は除く）
        Dim 外部 As Boolean, 外p1 As Long, 外p2 As Long, 外t As String, 外前 As String
        外部 = False
        外p1 = InStr(参照, "[")
        If 外p1 > 0 Then
            外p2 = InStr(外p1, 参照, "]")
            If 外p2 > 外p1 Then
                外t = LCase$(Mid$(参照, 外p1 + 1, 外p2 - 外p1 - 1))
                外前 = "="
                If 外p1 > 1 Then 外前 = Mid$(参照, 外p1 - 1, 1)
                If 外t Like "*.xl*" Or (IsNumeric(外t) And InStr("=(,+-*/&' ", 外前) > 0) Then 外部 = True
            End If
        End If

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
            nm.Delete
            If Err.Number = 0 Then
                sh.Cells(outRow, 4).Value = "消しました"
                消した = 消した + 1
            Else
                sh.Cells(outRow, 4).Value = "消せませんでした"
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf 外部 Then
            ' 別のブックから持ち込まれた名前（参照先が消えていても、2026 年の仕様変更で名前の管理にエラーが出ない）。消さない
            n外部 = n外部 + 1
            sh.Cells(outRow, 3).Value = "他のブックを指している（持ち込まれた名前）"
            If 使用 Then
                sh.Cells(outRow, 4).Value = "そのまま（式で使われている。外部リンクの元。解消は「他ブック参照を自ブックに直す」「外部リンクを値に変えて切る」）"
            Else
                sh.Cells(outRow, 4).Value = "消してよい候補（消していません・どこからも使われていない）"
            End If
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
    Application.StatusBar = "名前の整理: 名前 " & n全 & " 本のうち、壊れた名前 " & n壊れ & " 本を消しました（" & 消した & " 本）。使われていない候補が " & n未使用 & " 本（消していません）。" & _
        IIf(n外部 > 0, "他のブックを指す名前が " & n外部 & " 本あります（消していません）。", "") & "一覧はシート「" & 名 & "」。"
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
            ur行 = ur.row + ur.rows.Count - 1
            ur列 = ur.Column + ur.Columns.Count - 1

            ' 値と式の最終行・最終列（空の文字を返す式も「使っている」に数える。見た目が空の式の行を余分と見て、
            ' 用意しておいた式ごと消していた・2026-09-24 総点検）
            値行 = 0: 値列 = 0
            Dim 最後 As Range
            Set 最後 = Nothing
            On Error Resume Next
            Set 最後 = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlPart, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
            If Not 最後 Is Nothing Then 値行 = 最後.row
            Set 最後 = Nothing
            Set 最後 = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookAt:=xlPart, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
            If Not 最後 Is Nothing Then 値列 = 最後.Column
            On Error GoTo 0

            ' 図形の錨
            図行 = 0: 図列 = 0
            n図 = ws.Shapes.Count
            For Each shp In ws.Shapes
                On Error Resume Next
                If shp.BottomRightCell.row > 図行 Then 図行 = shp.BottomRightCell.row
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
    ' 並びの代わり: 表の書き方と罫線と列幅をそろえる
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
    Dim oldScreenUpdating As Boolean, oldEvents As Boolean
    Dim nRows As Long, urVal As Variant
    Dim arrZip As Variant, arrAddr As Variant
    Dim colZipRange As Range, colAddrRange As Range
    Dim zipChanged As Boolean, addrChanged As Boolean
    Dim 変えた住所() As Boolean

    Set ws = ActiveSheet
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    urTop = ur.row: urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)
    If urBottom <= urTop Then Exit Sub

    urVal = ur.Value2

    For r = 1 To UBound(urVal, 1)
        cnt = 0
        For c = 1 To UBound(urVal, 2)
            If セルの字(urVal(r, c)) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            hdrRow = urTop + r - 1
            Exit For
        End If
    Next r
    If hdrRow = 0 Or hdrRow >= urBottom Then Exit Sub

    ' 見出しから住所の列・郵便番号の列を探す（無ければ選んでいる列を住所とみなす）
    Dim hdrRelRow As Long
    hdrRelRow = hdrRow - urTop + 1
    For c = 1 To UBound(urVal, 2)
        名 = CStr(urVal(hdrRelRow, c))
        If 住所列 = 0 And InStr(名, "住所") > 0 Then 住所列 = urLeft + c - 1
        If 郵便列 = 0 And (InStr(名, "郵便") > 0 Or InStr(名, "〒") > 0) Then 郵便列 = urLeft + c - 1
    Next c
    If 住所列 = 0 Then
        住所列 = ActiveCell.Column
        If 住所列 < urLeft Or 住所列 > urRight Then 住所列 = urLeft
    End If

    nRows = urBottom - hdrRow
    If nRows < 1 Then Exit Sub

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 郵便番号列の取得
    If 郵便列 > 0 Then
        Set colZipRange = ws.Range(ws.Cells(hdrRow + 1, 郵便列), ws.Cells(urBottom, 郵便列))
        If nRows = 1 Then
            ReDim arrZip(1 To 1, 1 To 1)
            arrZip(1, 1) = colZipRange.Value2
        Else
            arrZip = colZipRange.Value2
        End If
        zipChanged = False
        Dim 変えた郵便() As Boolean
        ReDim 変えた郵便(1 To nRows, 1 To 1)
        For r = 1 To nRows
            If Not ws.Cells(hdrRow + r, 郵便列).HasFormula Then
                s = Trim$(CStr(arrZip(r, 1)))
                If VarType(arrZip(r, 1)) = vbDouble Then
                    If arrZip(r, 1) >= 0 And arrZip(r, 1) < 10000000 Then s = Format$(arrZip(r, 1), "0000000")
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
                    Dim zipFormatted As String
                    zipFormatted = Left$(t, 3) & "-" & Mid$(t, 4)
                    If CStr(arrZip(r, 1)) <> zipFormatted Then
                        arrZip(r, 1) = zipFormatted
                        n郵便 = n郵便 + 1
                        変えた郵便(r, 1) = True
                        zipChanged = True
                    End If
                End If
            End If
        Next r
        If zipChanged Then
            Call 表の整理_下請け.変えたセルだけ書く(colZipRange.Cells(1, 1), arrZip, 変えた郵便, 文字書式にする:=True)
        End If
    End If

    ' 住所列の取得
    Set colAddrRange = ws.Range(ws.Cells(hdrRow + 1, 住所列), ws.Cells(urBottom, 住所列))
    If nRows = 1 Then
        ReDim arrAddr(1 To 1, 1 To 1)
        arrAddr(1, 1) = colAddrRange.Value2
    Else
        arrAddr = colAddrRange.Value2
    End If
    addrChanged = False
    ReDim 変えた住所(1 To nRows, 1 To 1)

    For r = 1 To nRows
        If ws.Cells(hdrRow + r, 住所列).HasFormula Then GoTo 次の行
        s = CStr(arrAddr(r, 1))
        If Len(Trim$(s)) = 0 Then
            If 郵便列 > 0 Then
                If Trim$(CStr(arrZip(r, 1))) = "" And セルの字(ws.Cells(hdrRow + r, urLeft).Value) <> "" Then
                    要確認 = 要確認 & "・" & (hdrRow + r)
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
                If Right$(t, 1) Like "#" And (Mid$(s, i + 1, 1) Like "#" Or Mid$(s, i + 1, 1) Like "[０-９]") Then ch = "-"
            ElseIf cd = &H3000 Or cd = 160 Or cd = 9 Or cd = 10 Or cd = 13 Then
                ch = " "
            End If
            If Not (ch = " " And 前 = " ") Then t = t & ch
            前 = ch
        Next i
        t = Trim$(t)
        If t <> s Then
            arrAddr(r, 1) = t
            変えた住所(r, 1) = True
            n住所 = n住所 + 1
            addrChanged = True
        End If
        If 郵便列 > 0 Then
            If Len(Trim$(CStr(arrZip(r, 1)))) <> 8 Then
                要確認 = 要確認 & "・" & (hdrRow + r)
                n要確認 = n要確認 + 1
            End If
        End If
次の行:
    Next r

    ' 変えたセルだけ書く（数式セル・空欄・番地 "1-2-3" の日付化を防ぐ・2026-10-22）
    If addrChanged Then
        Call 表の整理_下請け.変えたセルだけ書く(colAddrRange.Cells(1, 1), arrAddr, 変えた住所, False)
    End If

CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.StatusBar = "住所と郵便番号: 住所 " & n住所 & " 件を元の列で整えました。" & _
                           IIf(郵便列 > 0, "郵便番号 " & n郵便 & " 件を 999-9999 にそろえました。", "郵便番号の列は見つかりませんでした。") & _
                           IIf(n要確認 > 0, "要確認の行: " & Mid$(要確認, 2) & "（郵便番号の形が読めない・住所も郵便番号も空）", "")
    If Err.Number <> 0 Then Err.Raise Err.Number
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
    ' 依頼の語: 列を非表示|列を隠|この列を隠|選んだ列を非表示|選んだ列を隠|列を見えなく
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
    ' ※ 021吸収合併: 共通本体（セル空白トリム_本体）へ委譲し高速化
    Call セル空白トリム_本体(1)
End Sub

Sub 空白だけのセルを空にする()
    ' 依頼の語: 見えない空白|空に見えるセル|空白だけのセル|スペースだけのセル|空白じゃない空白|空文字|中身が空白だけ
    ' 依頼の組: 見えない,空に見える,スペースだけ,空白だけ,空白じゃない+空白,セル,スペース+-行,列,詰,シート
    ' 扱う: 空白除去 空セル
    ' 見出し: なし
    ' 形: なし
    ' 空に見えるのに中身がある（空白だけ・空の文字）セルを本当の空にする。書式と数式は残す。
    ' 2 つ以上のセルを選んでいればその範囲、1 つだけなら今のシートの使っている範囲（shu003「選択範囲の空白じゃない空白をクリア」を棚へ）
    ' ※ 021吸収合併: 共通本体（セル空白トリム_本体）へ委譲し高速化
    Call セル空白トリム_本体(2)
End Sub

Sub シートの図形と画像を全部削除する()
    ' 依頼の語: 図形を全部消|図形を全部削除|図を全部消|図を全部削除|画像を全部消|全部の図形を消|図形を消して|図形を削除|図形をすべて消|要らない図形|不要な図形
    ' 依頼の組: 図形,図,画像,写真+全部,すべて,全て,要らない,不要+消,削除+-行,セル,選んだ,一覧,整理,そろえ,グラフ
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
        For rr = ar.row To ar.row + ar.rows.Count - 1
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
                If lo.HeaderRowRange.row = ks(j) Then keep = True
            End If
        Next lo
        If Not keep Then ActiveSheet.rows(ks(j)).Delete
    Next j
End Sub

Sub 選んだセルの図形を削除する()
    ' 依頼の語: セルの図を消|セルの図を削除|選んだ所の図|範囲の図を消|範囲の図形を消|セルにある図|セルの上の図|ここの図
    ' 依頼の組: 図,図形,画像,写真+選んだ,セル,範囲,ここ+消,削除+-行,全部,すべて,一覧
    ' 階級: 基本
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
    For rr = sel.row + 1 To sel.row + sel.rows.Count - 1
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
    Call シートの行高と列幅を写す_本体(ActiveSheet.Index - 1)
End Sub

Sub 一枚目のシートの行高と列幅を写す()
    ' 依頼の語: 1枚目のシートに合わせ|最初のシートに合わせ|1枚目と同じ|先頭のシートに合わせ|一枚目のシート|１枚目のシート|1枚目のシートと同じ
    ' 依頼の組: 1枚目,一枚目,１枚目,最初のシート,先頭のシート+合わせ,同じ,揃え,そろえ
    ' 扱う: 列幅 行の高さ
    ' 見出し: なし
    ' 形: なし
    ' 1 枚目のシートの行の高さと列幅を、今のシートへ写す。1 枚目のシートでは何もしない（shu003「行高と列幅を1枚目のシートに揃える」を棚へ）
    Call シートの行高と列幅を写す_本体(1)
End Sub

Sub 見出しの語で列に名前を付ける()
    ' 依頼の語: 見出しで名前|見出しの名前で|列に名前|見出しを名前に|範囲に名前を付|名前の定義を作|名前を付けて
    ' 依頼の組: 名前+付け,定義を作,作って+見出し,列,範囲+-整理,消,削除,揺れ,一覧,シート名,ファイル名
    ' 扱う: 名前 名前の定義
    ' 見出し: なし
    ' 形: なし
    ' 表の見出しの語を名前にして、その下の列の範囲に名前を付ける（ブック全体の名前）。名前に使えない字は _ に替える。
    ' 付けられない見出しは飛ばす（shu003「アクティブワークシートの一行目の値で名前をつける」を棚へ。見出しは 1 行目に限らず表の見出しの行）
    Dim ws As Worksheet, ur As Range, hR As Long, r As Long, c As Long, c1 As Long, c2 As Long, n As Long, best As Long
    Dim lastR As Long, nm As String, k As Long, ch As String, t As String
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur.rows.Count < 2 Then Exit Sub
    For r = ur.row To ur.row + IIf(ur.rows.Count > 20, 20, ur.rows.Count) - 1
        n = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, ur.Column), ws.Cells(r, ur.Column + ur.Columns.Count - 1)))
        If n >= 2 Then hR = r: Exit For
    Next r
    If hR = 0 Then Exit Sub
    c1 = ur.Column
    Do While c1 < ur.Column + ur.Columns.Count - 1 And セルの字(ws.Cells(hR, c1).Value) = ""
        c1 = c1 + 1
    Loop
    c2 = c1
    Do While c2 < ur.Column + ur.Columns.Count - 1 And セルの字(ws.Cells(hR, c2 + 1).Value) <> ""
        c2 = c2 + 1
    Loop
    For c = c1 To c2
        t = Trim$(CStr(ws.Cells(hR, c).Value))
        If セルの字(t) <> "" Then
            nm = ""
            For k = 1 To Len(t)
                ch = Mid$(t, k, 1)
                If InStr(" !""#$%&'()*+,-/:;<=>?@[]^`{|}~　", ch) > 0 Then nm = nm & "_" Else nm = nm & ch
            Next k
            If nm Like "#*" Or UCase$(nm) Like "[A-Z]#*" Or UCase$(nm) Like "[A-Z][A-Z]#*" Or UCase$(nm) Like "[A-Z][A-Z][A-Z]#*" Or UCase$(nm) = "R" Or UCase$(nm) = "C" Then nm = "_" & nm
            ' 名前の範囲は表の本文だけ（列の最後の値までにすると、下の合計の行・メモまで入り、SUM(金額) が 2 倍になった・2026-09-24 総点検）
            lastR = 表の本文の最終行(ws)
            Do While lastR > hR And 集計の行か(ws, lastR, c1, c2)
                lastR = lastR - 1
            Loop
            If ws.Cells(ws.rows.Count, c).End(xlUp).row < lastR Then lastR = ws.Cells(ws.rows.Count, c).End(xlUp).row
            If lastR > hR Then
                On Error Resume Next
                ws.Parent.names.Add Name:=nm, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Range(ws.Cells(hR + 1, c), ws.Cells(lastR, c)).Address
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
    ' 変えてよい: 行移動
    ' 依頼の語: 空欄の行を削除|空欄の行を消|が空の行を|空欄の行を取り除|未入力の行を削除|未入力の行を消|空いている行を消
    ' 依頼の組: 空欄,空白,未入力,空い+行+削除,消,除+-一覧,探,調べ,チェック,詰,空行,空白行,空白の行
    ' 扱う: 行削除 空欄
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる列が空欄の行を、表（選んだセルの周りのひとまとまり）の中で消す。見出しの行と合計の行は残す（shu003「選択範囲の空白の行を削除」を棚へ）
    Dim ws As Worksheet, reg As Range, col As Long, r As Long, R1 As Long, r2 As Long, s As String, k As Long, sumRow As Boolean
    Dim regTop As Long, regLeft As Long, regCols As Long, relCol As Long, nRows As Long
    Dim regVal As Variant
    Dim delFlags() As Boolean, hasDel As Boolean
    Dim oldScreenUpdating As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    Dim delRange As Range, batchStr As String

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    col = Selection.Cells(1, 1).Column
    Set reg = Selection.Cells(1, 1).CurrentRegion
    If reg.rows.Count < 2 Then Exit Sub

    regTop = reg.row
    regLeft = reg.Column
    regCols = reg.Columns.Count
    relCol = col - regLeft + 1
    If relCol < 1 Or relCol > regCols Then Exit Sub

    ' 二段見出しの下の段も見出し。縦の結合の下側は空に見えるが空欄ではない
    R1 = 見出しの下の段(ws, reg.row, reg.Column, reg.Column + reg.Columns.Count - 1) + 1
    r2 = reg.row + reg.rows.Count - 1
    If r2 < R1 Then Exit Sub

    nRows = reg.rows.Count
    If nRows = 1 Then
        ReDim regVal(1 To 1, 1 To regCols)
        For k = 1 To regCols: regVal(1, k) = reg.Cells(1, k).Value: Next k
    Else
        regVal = reg.Value2
    End If

    ReDim delFlags(R1 To r2)
    hasDel = False

    For r = r2 To R1 Step -1
        Dim relRow As Long
        relRow = r - regTop + 1
        Dim isMergedChild As Boolean
        isMergedChild = False
        For k = regLeft To regLeft + regCols - 1
            If ws.Cells(r, k).MergeCells Then
                If ws.Cells(r, k).MergeArea.rows.Count > 1 Then
                    isMergedChild = True
                    Exit For
                End If
            End If
        Next k

        If Trim$(CStr(regVal(relRow, relCol))) = "" And Not ws.Cells(r, col).MergeCells And Not isMergedChild Then
            sumRow = False
            For k = 1 To regCols
                s = Replace(Trim$(CStr(regVal(relRow, k))), " ", "")
                If 集計の語か(s) Then
                    sumRow = True
                    Exit For
                End If
            Next k
            If Not sumRow Then
                delFlags(r) = True
                hasDel = True
            End If
        End If
    Next r

    ' 事前高速判定: 削除対象が0件なら即復帰（0秒判定）
    If Not hasDel Then Exit Sub

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' アドレスバッチ結合による一撃削除
    batchStr = ""
    For r = r2 To R1 Step -1
        If delFlags(r) Then
            If Len(batchStr) > 200 Then
                If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
                batchStr = ""
            End If
            batchStr = batchStr & IIf(Len(batchStr) > 0, ",", "") & r & ":" & r
        End If
    Next r
    If Len(batchStr) > 0 Then
        If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
    End If

    If Not delRange Is Nothing Then delRange.Delete Shift:=xlUp

CleanUp:
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
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
    Dim fc As Range, c As Range, re1 As Object, re2 As Object, res As Object, f As String, g As String, calc As Long
    Dim m As Object, nm As String, ok As Boolean, sh As Object, n直 As Long, n残 As Long
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    On Error Resume Next
    Set fc = ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fc Is Nothing Then Exit Sub
    ' 正規表現: VBScript が外された環境では VBA 標準の RegExp に切り替える（2027 年対応）
    On Error Resume Next
    Set re1 = CreateObject("VBScript.RegExp"): If re1 Is Nothing Then Set re1 = 正規表現_標準版を作る()
    Set re2 = CreateObject("VBScript.RegExp"): If re2 Is Nothing Then Set re2 = 正規表現_標準版を作る()
    Set res = CreateObject("VBScript.RegExp"): If res Is Nothing Then Set res = 正規表現_標準版を作る()
    On Error GoTo 0
    If re1 Is Nothing Or re2 Is Nothing Or res Is Nothing Then
        Application.StatusBar = "外部参照を直す: 正規表現が使えないので直せませんでした（VBScript が無効で、この Excel の VBA にも標準の RegExp がありません）"
        Exit Sub
    End If
    re1.Global = True: re1.IgnoreCase = True
    re1.Pattern = "'[^'\[]*\[[^\[\]]+\.xl[a-z]{1,2}\]([^']+)'!"
    re2.Global = True: re2.IgnoreCase = True
    re2.Pattern = "\[[^\[\]]+\.xl[a-z]{1,2}\]"
    res.Global = True
    res.Pattern = "(?:'((?:[^']|'')+)'|([^\s=+\-*/^&,;:<>()!'""{}]+))!"
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
                For Each m In res.Execute(g)
                    nm = m.SubMatches(0)
                    If Len(nm) = 0 Then nm = m.SubMatches(1)
                    nm = Replace(nm, "''", "'")
                    Set sh = Nothing
                    On Error Resume Next
                    Set sh = ActiveWorkbook.Sheets(nm)
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
    ' 依頼の語: 1ページに収め|１ページに収め|一ページに収め|1枚に収め|1枚で印刷|1ページで印刷|1ページに印刷|1枚に印刷|1ページに収めて印刷できるように
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
    ' 変えてよい: 文字数字
    ' 依頼の語: 番号を振り直|番号を振りなお|連番を振り直|連番を振りなお|連番にし|番号を付け直|番号をつけ直|番号が飛|番号の抜け|番号が抜け|番号が重複|通し番号を|Noを振り直|No.を振り直|項番を振り直
    ' 依頼の組: 番号,連番,No,通し番号,項番+振り直,振りなお,付け直,つけ直,振って,飛,抜け,詰め,重複,ずれ,1から+-一覧,洗い出,チェック,調べ,探,確認,何件,右,伝票,会員,郵便
    ' 扱う: 番号 連番
    ' 見出し: なし
    ' 形: 番号の列
    ' 見出しが No・番号・連番・通し番号・項番 の列（無ければ左端の、1 から始まる整数の列）を、表の本文の上から 1, 2, 3… に振り直す。
    ' 伝票番号・会員番号のような「名前の付いた番号」は ID なので触らない。式の入った番号列・合計の行・空の行も触らない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, lastRow As Long, nCols As Long, nRows As Long
    Dim headerRow As Long, dataStart As Long, noCol As Long, relNoCol As Long
    Dim r As Long, c As Long, cnt As Long, n As Long, idx As Long, tblR As Long
    Dim h As String, v As Variant, rowHasValue As Boolean, isTotal As Boolean
    Dim 表値 As Variant, 列式 As Variant, colRange As Range, colChanged As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    nCols = urRight - urLeft + 1
    lastRow = 表の本文の最終行(ws)
    If lastRow < urTop Then Exit Sub

    ' 表全体の値を一括取得
    If lastRow = urTop And nCols = 1 Then
        ReDim 表値(1 To 1, 1 To 1)
        表値(1, 1) = ws.Cells(urTop, urLeft).Value
    Else
        表値 = ws.Range(ws.Cells(urTop, urLeft), ws.Cells(lastRow, urRight)).Value
    End If

    ' 見出しの行＝値が 2 つ以上ある最初の行
    For r = urTop To lastRow
        tblR = r - urTop + 1
        cnt = 0
        For c = 1 To nCols
            If セルの字(表値(tblR, c)) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    dataStart = headerRow + 1
    If dataStart > lastRow Then Exit Sub
    nRows = lastRow - dataStart + 1

    ' 番号の列＝見出しがちょうど番号の語（伝票番号・会員番号は ID なので外す）
    tblR = headerRow - urTop + 1
    For c = 1 To nCols
        h = StrConv(Trim$(Replace(セルの字(表値(tblR, c)), "　", "")), vbNarrow)
        h = UCase$(Replace(h, ".", ""))
        If h = "NO" Or h = "№" Or h = "番号" Or h = "連番" Or h = "通し番号" Or h = "項番" Or h = "#" Then
            noCol = urLeft + c - 1
            Exit For
        End If
    Next c
    ' 見出しで決まらなければ、左端の列が 1 から始まる整数の並びのときだけ
    If noCol = 0 Then
        v = 表値(dataStart - urTop + 1, 1)
        ' 文字の「0001」（銀行コード・顧客コード）は番号の列ではない＝数で入った 1 のときだけ（1,2,3… に振り直すおそれ・2026-10-08）
        If IsNumeric(v) And Not isEmpty(v) And VarType(v) <> vbString Then
            If CDbl(v) = 1 Then noCol = urLeft
        End If
    End If
    If noCol = 0 Then Exit Sub
    relNoCol = noCol - urLeft + 1

    ' 番号列の式を一括取得し、式で番号を出している列（=ROW()-1 など）は触らない
    Set colRange = ws.Range(ws.Cells(dataStart, noCol), ws.Cells(lastRow, noCol))
    If nRows > 1 Then
        列式 = colRange.Formula
    Else
        ReDim 列式(1 To 1, 1 To 1)
        列式(1, 1) = colRange.Formula
    End If
    For idx = 1 To nRows
        If Left$(CStr(列式(idx, 1)), 1) = "=" Then Exit Sub
    Next idx

    ' 今の番号が「001」のような先頭ゼロ付きの文字なら、その桁の形のまま振り直す（数の 1, 2, 3 にしてゼロを消さない・2026-09-23）
    Dim padL As Long, zeroN As Long, filledN As Long, sv As String
    For r = dataStart To lastRow
        v = 表値(r - urTop + 1, relNoCol)
        sv = Trim$(セルの字(v))
        If sv <> "" Then
            filledN = filledN + 1
            If VarType(v) = vbString And Len(sv) >= 2 And Left$(sv, 1) = "0" And sv Like String(Len(sv), "#") Then
                zeroN = zeroN + 1
                If Len(sv) > padL Then padL = Len(sv)
            End If
        End If
    Next r
    If zeroN < 2 Or zeroN < filledN / 2 Then padL = 0

    n = 0
    colChanged = False
    Dim 変えた連番() As Boolean
    ReDim 変えた連番(1 To lastRow - dataStart + 1, 1 To 1)
    For r = dataStart To lastRow
        tblR = r - urTop + 1
        idx = r - dataStart + 1
        rowHasValue = False
        isTotal = False
        For c = 1 To nCols
            h = Trim$(セルの字(表値(tblR, c)))
            If 集計の語か(h) Then isTotal = True
            If c <> relNoCol And h <> "" Then rowHasValue = True
        Next c
        If rowHasValue And Not isTotal Then
            n = n + 1
            If padL > 0 Then
                sv = Right$(String(padL, "0") & CStr(n), IIf(Len(CStr(n)) > padL, Len(CStr(n)), padL))
                If セルの字(表値(tblR, relNoCol)) <> sv Then
                    列式(idx, 1) = sv
                    変えた連番(idx, 1) = True
                    colChanged = True
                End If
            ElseIf セルの字(表値(tblR, relNoCol)) <> CStr(n) Then
                列式(idx, 1) = n
                変えた連番(idx, 1) = True
                colChanged = True
            End If
        End If
    Next r

    If colChanged Then
        Application.ScreenUpdating = False
        Call 表の整理_下請け.変えたセルだけ書く(colRange.Cells(1, 1), 列式, 変えた連番, 文字書式にする:=(padL > 0))
        Application.ScreenUpdating = True
    End If
End Sub

Sub 番号に先頭のゼロを付けてそろえる()
    ' 変えてよい: 文字数字
    ' 依頼の語: 先頭のゼロ|先頭の0|頭のゼロ|ゼロが消え|0が消え|ゼロが落ち|0が落ち|桁をそろえ|桁を揃え|ゼロ埋め|0埋め|ゼロを付け|0を付け|桁数をそろえ|桁数を揃え
    ' 依頼の組: ゼロ,0,０,桁+消え,落ち,抜け,そろえ,揃え,埋め,付け,戻+-小数,四捨五入,切り捨て,丸め,円,金額,合計,グラフ,縦軸,横軸,目盛り
    ' 扱う: 番号 桁 先頭ゼロ
    ' 見出し: なし
    ' 形: 番号の列
    ' 「0001」のように先頭にゼロの付いた番号が並ぶ列（同じ桁のゼロ付きが 2 つ以上・列の 6 割以上）で、ゼロが消えて「4」になったセルを
    ' 列の桁にそろえて文字の「0004」に戻す。2 つ以上のセルを選んでいればその列だけ、1 つだけなら表の全部の列を見る。式のセルは触らない。
    Dim ws As Worksheet, ur As Range
    Dim urTop As Long, urLeft As Long, urRight As Long, lastRow As Long
    Dim headerRow As Long, dataStart As Long
    Dim r As Long, c As Long, cnt As Long, bestL As Long, bestN As Long, filled As Long
    Dim s As String, v As Variant
    Dim lens As Object, useCol As Object, key As Variant
    Dim colChanged As Boolean, colRange As Range

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.row
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

    Application.ScreenUpdating = False

    For c = urLeft To urRight
        If useCol.Count = 0 Or useCol.exists(c) Then
            ' 列の中のゼロ付きの番号の桁を数える
            Set lens = CreateObject("Scripting.Dictionary")
            filled = 0
            ' 列の値と式は最初にまとめて読む
            Dim 列値 As Variant, 列式 As Variant
            If lastRow > dataStart Then
                列値 = ws.Range(ws.Cells(dataStart, c), ws.Cells(lastRow, c)).Value
                列式 = ws.Range(ws.Cells(dataStart, c), ws.Cells(lastRow, c)).Formula
            Else
                ReDim 列値(1 To 1, 1 To 1): ReDim 列式(1 To 1, 1 To 1)
                列値(1, 1) = ws.Cells(dataStart, c).Value: 列式(1, 1) = ws.Cells(dataStart, c).Formula
            End If
            For r = dataStart To lastRow
                v = 列値(r - dataStart + 1, 1)
                s = Trim$(セルの字(v))
                If s <> "" And Left$(CStr(列式(r - dataStart + 1, 1)), 1) <> "=" Then
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
                    s = Trim$(セルの字(列値(r - dataStart + 1, 1)))
                    If Len(s) = bestL And s Like String(bestL, "#") Then cnt = cnt + 1
                Next r
                If cnt >= filled * 0.6 Then
                    colChanged = False
                    Dim 変えた() As Boolean
                    ReDim 変えた(1 To lastRow - dataStart + 1, 1 To 1)
                    For r = dataStart To lastRow
                        If Left$(CStr(列式(r - dataStart + 1, 1)), 1) <> "=" Then
                            v = 列値(r - dataStart + 1, 1)
                            s = Trim$(セルの字(v))
                            If s <> "" And Len(s) < bestL And s Like String(Len(s), "#") Then
                                列式(r - dataStart + 1, 1) = Right$(String(bestL, "0") & s, bestL)
                                変えた(r - dataStart + 1, 1) = True
                                colChanged = True
                            End If
                        End If
                    Next r
                    If colChanged Then
                        Call 表の整理_下請け.変えたセルだけ書く(ws.Cells(dataStart, c), 列式, 変えた, 文字書式にする:=True)
                    End If
                End If
            End If
        End If
    Next c

    Application.ScreenUpdating = True
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
    Dim headerRow As Long, dataStart As Long, nRows As Long
    Dim r As Long, c As Long, cnt As Long, idx As Long
    Dim s As String, k As String, best As String, bestN As Long
    Dim forms As Object, bestMap As Object, useCol As Object, f As Object, key As Variant, key2 As Variant
    Dim 列値 As Variant, 列式 As Variant, colRange As Range
    Dim colChanged As Boolean, 式あり As Boolean, v As Variant
    Dim 変えた() As Boolean
    Dim oldScreenUpdating As Boolean, oldEvents As Boolean
    Dim urVal As Variant

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    lastRow = 表の本文の最終行(ws)
    If lastRow < urTop Then Exit Sub

    urVal = ur.Value2

    For r = 1 To UBound(urVal, 1)
        cnt = 0
        For c = 1 To UBound(urVal, 2)
            If セルの字(urVal(r, c)) <> "" Then cnt = cnt + 1
        Next c
        If cnt >= 2 Then
            headerRow = urTop + r - 1
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub
    dataStart = headerRow + 1
    If dataStart > lastRow Then Exit Sub
    nRows = lastRow - dataStart + 1

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

    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For c = urLeft To urRight
        If useCol.Count = 0 Or useCol.exists(c) Then
            Set colRange = ws.Range(ws.Cells(dataStart, c), ws.Cells(lastRow, c))
            If nRows > 1 Then
                列値 = colRange.Value
                列式 = colRange.Formula
            Else
                ReDim 列値(1 To 1, 1 To 1): ReDim 列式(1 To 1, 1 To 1)
                列値(1, 1) = colRange.Value: 列式(1, 1) = colRange.Formula
            End If

            ' 空白を抜いた字 → {書き方: 件数}
            Set forms = CreateObject("Scripting.Dictionary")
            式あり = False
            For idx = 1 To nRows
                If Left$(CStr(列式(idx, 1)), 1) = "=" Then
                    式あり = True
                Else
                    v = 列値(idx, 1)
                    If VarType(v) = vbString Then
                        s = CStr(v)
                        k = Replace(Replace(s, " ", ""), "　", "")
                        If k <> "" Then
                            If Not forms.exists(k) Then forms.Add k, CreateObject("Scripting.Dictionary")
                            Set f = forms(k)
                            f(s) = f(s) + 1
                        End If
                    End If
                End If
            Next idx

            ' 揺れ（2種類以上）があるキーの採用表記を決定
            Set bestMap = CreateObject("Scripting.Dictionary")
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
                    bestMap(key) = best
                End If
            Next key

            If bestMap.Count > 0 Then
                colChanged = False
                ReDim 変えた(1 To nRows, 1 To 1)
                For idx = 1 To nRows
                    If Left$(CStr(列式(idx, 1)), 1) <> "=" And VarType(列値(idx, 1)) = vbString Then
                        s = CStr(列値(idx, 1))
                        k = Replace(Replace(s, " ", ""), "　", "")
                        If bestMap.exists(k) Then
                            best = bestMap(k)
                            If s <> best Then
                                列値(idx, 1) = best
                                変えた(idx, 1) = True
                                colChanged = True
                            End If
                        End If
                    End If
                Next idx
                ' 変えたセルだけ書く（数式セル・文字の数字・空欄を壊さない・2026-10-22）
                If colChanged Then
                    Call 表の整理_下請け.変えたセルだけ書く(colRange.Cells(1, 1), 列値, 変えた, False)
                End If
            End If
        End If
    Next c

CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 選んだ文字で入力規則リストを作る()
    ' 依頼の語: 入力規則のリスト|プルダウンを作|ドロップダウンを作|プルダウンにし|ドロップダウンにし|プルダウンリスト|ドロップダウンリスト|リストから選べ|選べるようにし|選択肢を作|入力規則を設定|入力規則にし|入力規則のプルダウン|入力規則を付け
    ' 依頼の組: プルダウン,ドロップダウン,入力規則,選択肢+作,設定,付け,つけ,にし,選べ,入れ+-一覧,調べ,探,点検,削除,解除,消
    ' 扱う: 入力規則 プルダウン リスト
    ' 見出し: なし
    ' 形: なし
    ' 今のセル（アクティブセル）に書いた「項目をスペースかカンマで区切った文字」を、選んでいるセルすべての入力規則のリスト（プルダウン）にする。
    ' 区切りは、半角・全角のスペース・カンマ・読点・タブ・改行（カンマがあるときはカンマと改行だけ＝項目の中にスペースを使える）。重複は除く。
    ' 項目の代わりに =A2:A10 や =Sheet2!$A$1:$A$9 のような範囲参照（文字として入れた形でも、式として入れた形でも）を書くと、その範囲がリストになる。
    ' 元の文字を書いたセルは設定のあと空にする。項目が 1 つだけのとき・255 字を超えるとき・範囲参照が使えないときは、何も変えずに知らせる（普通のデータを消さない）。
    ' 2026-10-01 オフィス田中「アドイン解説 入力規則のリスト」の型（Selection に一括・既存の規則は入れ替え）に、安全のための確認を足した
    Dim 元 As String, 参照 As String, 項目 As Variant, 結果 As String, 区切り As String
    Dim 見 As Object, 一つ As String, 設定 As String
    Dim 域 As Range, 先頭 As Range

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set 先頭 = ActiveCell
    If 先頭 Is Nothing Then Exit Sub
    If 先頭.Cells.CountLarge > 1 Then Set 先頭 = 先頭.Cells(1, 1)

    ' 元の文字（式なら式の字）。空なら何もしない
    参照 = ""
    If 先頭.HasFormula Then
        元 = 先頭.Formula
    Else
        元 = 先頭.Value
    End If
    元 = Trim$(元)
    If 元 = "" Then Exit Sub

    If Left$(元, 1) = "=" Then
        ' 範囲参照（=A2:A10・=Sheet2!$A$1:$A$9・=部署リスト）。関数が入っていれば参照ではない
        If InStr(元, "(") > 0 Then
            Application.StatusBar = "入力規則のリスト: 式が入っています。項目をスペースかカンマで区切って書くか、=A2:A10 のような範囲参照だけを書いてください。"
            Exit Sub
        End If
        参照 = 元
        設定 = 参照
    Else
        ' 区切りをカンマにそろえる（改行・読点・全角のカンマは常に区切り。カンマが無ければスペースとタブも区切り）
        区切り = Replace(Replace(Replace(元, vbCrLf, ","), vbLf, ","), vbCr, ",")
        区切り = Replace(Replace(区切り, "、", ","), "，", ",")
        If InStr(区切り, ",") = 0 Then
            区切り = Replace(Replace(区切り, vbTab, " "), "　", " ")
            区切り = Replace(区切り, " ", ",")
        End If
        Set 見 = CreateObject("Scripting.Dictionary")
        結果 = ""
        For Each 項目 In Split(区切り, ",")
            一つ = CStr(項目)
            Do While Len(一つ) > 0
                If Left$(一つ, 1) = " " Or Left$(一つ, 1) = "　" Or Left$(一つ, 1) = vbTab Then 一つ = Mid$(一つ, 2) Else Exit Do
            Loop
            Do While Len(一つ) > 0
                If Right$(一つ, 1) = " " Or Right$(一つ, 1) = "　" Or Right$(一つ, 1) = vbTab Then 一つ = Left$(一つ, Len(一つ) - 1) Else Exit Do
            Loop
            If 一つ <> "" Then
                If Not 見.exists(一つ) Then
                    見.Add 一つ, 1
                    結果 = 結果 & IIf(結果 = "", "", ",") & 一つ
                End If
            End If
        Next 項目
        If 見.Count < 2 Then
            Application.StatusBar = "入力規則のリスト: 項目が 1 つしか見つかりません（" & 見.Count & " 個）。スペースかカンマで 2 つ以上を区切って書いてください。何も変えていません。"
            Exit Sub
        End If
        If Len(結果) > 255 Then
            Application.StatusBar = "入力規則のリスト: 項目が長すぎます（" & Len(結果) & " 字。直接書けるのは 255 字まで）。項目をどこかのセルに並べて =A2:A10 の形で書いてください。何も変えていません。"
            Exit Sub
        End If
        設定 = 結果
    End If

    ' 先頭のセルで試す（範囲参照が使えない形・シートの保護はここで分かる。元の文字はまだ消さない）
    On Error Resume Next
    先頭.Validation.Delete
    Err.Clear
    先頭.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=設定
    If Err.Number <> 0 Then
        Application.StatusBar = "入力規則のリスト: 設定できませんでした（" & Err.Description & "）。範囲参照の書き方か、シートの保護を確かめてください。元の文字はそのままです。"
        Exit Sub
    End If
    On Error GoTo 0

    ' 選んでいるセルすべて（離れた範囲もまとめて）に設定する。既にある入力規則は入れ替える
    For Each 域 In Selection.Areas
        On Error Resume Next
        域.Validation.Delete
        Err.Clear
        域.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=設定
        If Err.Number <> 0 Then
            Application.StatusBar = "入力規則のリスト: " & 域.Address(False, False) & " に設定できませんでした（" & Err.Description & "）。"
            Err.Clear
        Else
            域.Validation.IgnoreBlank = True
            域.Validation.InCellDropdown = True
        End If
        On Error GoTo 0
    Next 域

    ' 元の文字は残さない（残すとリストの中に見えないまま邪魔になる）
    先頭.ClearContents
End Sub

Public Sub マクロ撃ち(Optional ByVal rng As Range = Nothing)
    ' 本体は modMacroUchi50.マクロ撃ち50_実行 に一本化（旧版の守り・クレンジングは移植済み。旧本文は backups に控えあり）
    Dim ws As Worksheet
    If rng Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = rng.Worksheet
    End If
    If ws Is Nothing Then Exit Sub
    Call modMacroUchi50.マクロ撃ち50_実行(ws)
End Sub


Public Sub 表の自律修復(Optional ByVal rng As Range = Nothing)
    ' 互換性のためのエイリアス
    マクロ撃ち rng
End Sub

Public Function 原本シート作成(ByVal ws As Worksheet) As Worksheet
    On Error GoTo CatchErr
    If ws Is Nothing Then Exit Function
    If Right$(ws.Name, 2) = "_元" Then Exit Function ' 原本シート自身はコピーしない
    
    Dim wb As Workbook: Set wb = ws.Parent
    ' テスト用シート（関所検査用一時など）の場合は原本を作らない（本体残存を物理防止）
    If InStr(ws.Name, "関所") > 0 Or InStr(ws.Name, "UT_") > 0 Or InStr(ws.Name, "Test") > 0 Then
        Set 原本シート作成 = Nothing
        Exit Function
    End If
    
    Dim bkName As String: bkName = ws.Name & "_元"
    
    Dim prevSU As Boolean, prevDA As Boolean
    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 既存の同名原本シートがあれば確実に削除（非表示であっても確実に消す）
    On Error Resume Next
    Dim oldWs As Worksheet
    Set oldWs = wb.Worksheets(bkName)
    If Not oldWs Is Nothing Then
        oldWs.Visible = xlSheetVisible
        oldWs.Delete
    End If
    ' もし既存の (2) シートなどの残骸があれば削除
    Dim dupWs As Worksheet
    Set dupWs = wb.Worksheets(ws.Name & " (2)")
    If Not dupWs Is Nothing Then
        dupWs.Delete
    End If
    On Error GoTo 0
    
    ' シートをコピー
    ws.Copy after:=ws
    Dim newWs As Worksheet
    Set newWs = wb.Worksheets(ws.Index + 1)
    
    On Error Resume Next
    newWs.Name = bkName
    newWs.Tab.Color = RGB(192, 192, 192) ' 原本はグレーのタブ
    newWs.Visible = xlSheetVisible ' 控えシートを表示
    On Error GoTo 0
    
    ws.Activate
    Set 原本シート作成 = newWs
    
    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU
    Exit Function
CatchErr:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Function

Public Sub 原本から元に戻す(Optional ByVal ws As Worksheet = Nothing, Optional ByVal isQuiet As Boolean = True)
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    
    Dim wb As Workbook: Set wb = ws.Parent
    Dim targetWs As Worksheet, bkWs As Worksheet
    
    If Right$(ws.Name, 2) = "_元" Then
        Set bkWs = ws
        Dim orgName As String: orgName = Left$(ws.Name, Len(ws.Name) - 2)
        On Error Resume Next
        Set targetWs = wb.Worksheets(orgName)
        On Error GoTo 0
    Else
        Set targetWs = ws
        On Error Resume Next
        Set bkWs = wb.Worksheets(ws.Name & "_元")
        On Error GoTo 0
    End If
    
    If bkWs Is Nothing Then Exit Sub
    If targetWs Is Nothing Then Exit Sub
    
    ' 呼んだ側の警告の設定は元に戻す（True に決め打ちで戻すと、続けて消すシートで確認の窓が出る・2026-10-08）
    Dim 前の警告 As Boolean: 前の警告 = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    targetWs.Cells.Clear
    bkWs.Cells.Copy targetWs.Range("A1")
    
    bkWs.Cells.Copy
    targetWs.Range("A1").PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False
    
    targetWs.Activate
    targetWs.Range("A1").Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = 前の警告
End Sub

Public Function 直した場所を着色(Optional ByVal ws As Worksheet = Nothing, Optional ByVal 付与コメント As Boolean = False) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Then Exit Function
    
    Dim wb As Workbook: Set wb = ws.Parent
    Dim bkWs As Worksheet
    On Error Resume Next
    Set bkWs = wb.Worksheets(ws.Name & "_元")
    On Error GoTo 0
    
    If bkWs Is Nothing Then Exit Function
    
    Application.ScreenUpdating = False
    
    Dim rCur As Range: Set rCur = ws.UsedRange
    Dim rBk As Range: Set rBk = bkWs.UsedRange
    
    Dim maxR As Long, maxC As Long
    maxR = rCur.rows.Count: If rBk.rows.Count > maxR Then maxR = rBk.rows.Count
    maxC = rCur.Columns.Count: If rBk.Columns.Count > maxC Then maxC = rBk.Columns.Count
    
    If maxR > 10000 Then maxR = 10000
    If maxC > 100 Then maxC = 100
    
    Dim vCur As Variant, vBk As Variant
    vCur = ws.Range(ws.Cells(1, 1), ws.Cells(maxR, maxC)).Value2
    vBk = bkWs.Range(bkWs.Cells(1, 1), bkWs.Cells(maxR, maxC)).Value2
    
    Dim fCur As Variant, fBk As Variant
    fCur = ws.Range(ws.Cells(1, 1), ws.Cells(maxR, maxC)).Formula
    fBk = bkWs.Range(bkWs.Cells(1, 1), bkWs.Cells(maxR, maxC)).Formula
    
    Dim diffCount As Long: diffCount = 0
    Dim i As Long, j As Long
    For i = 1 To maxR
        For j = 1 To maxC
            Dim isDiff As Boolean: isDiff = False
            Dim sFCur As String: sFCur = "": If Not IsError(fCur(i, j)) Then sFCur = CStr(fCur(i, j))
            Dim sFBk As String: sFBk = "": If Not IsError(fBk(i, j)) Then sFBk = CStr(fBk(i, j))
            Dim sVCur As String: sVCur = "": If Not IsError(vCur(i, j)) Then sVCur = CStr(vCur(i, j))
            Dim sVBk As String: sVBk = "": If Not IsError(vBk(i, j)) Then sVBk = CStr(vBk(i, j))
            
            If (Left$(sFCur, 1) = "=" Or Left$(sFBk, 1) = "=") And sFCur <> sFBk Then
                isDiff = True
                ' 数式復元・修復（薄緑: RGB(220, 245, 220)）
                ws.Cells(i, j).Interior.Color = RGB(220, 245, 220)
                If 付与コメント Then
                    On Error Resume Next
                    ws.Cells(i, j).ClearComments
                    ws.Cells(i, j).AddComment "【数式復元】" & vbCrLf & "元: " & sFBk & vbCrLf & "新: " & sFCur
                    On Error GoTo 0
                End If
            ElseIf sVCur <> sVBk Then
                isDiff = True
                ' 値・型クレンジング（薄黄: RGB(255, 255, 180)）
                ws.Cells(i, j).Interior.Color = RGB(255, 255, 180)
                If 付与コメント Then
                    On Error Resume Next
                    ws.Cells(i, j).ClearComments
                    ws.Cells(i, j).AddComment "【マクロ撃ち修復】" & vbCrLf & "元: " & sVBk & vbCrLf & "新: " & sVCur
                    On Error GoTo 0
                End If
            End If
            If isDiff Then diffCount = diffCount + 1
        Next j
    Next i
    
    Application.ScreenUpdating = True
    直した場所を着色 = diffCount
End Function

Public Sub 原本シート削除(Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    
    Dim wb As Workbook: Set wb = ws.Parent
    Dim targetBk As Worksheet
    
    If Right$(ws.Name, 2) = "_元" Then
        Set targetBk = ws
    Else
        On Error Resume Next
        Set targetBk = wb.Worksheets(ws.Name & "_元")
        On Error GoTo 0
    End If
    
    If targetBk Is Nothing Then Exit Sub
    
    Application.DisplayAlerts = False
    targetBk.Delete
    Application.DisplayAlerts = True
End Sub

Public Function マクロ撃ち診断(Optional ByVal ws As Worksheet = Nothing) As String
    ' 本体は modMacroUchi50.マクロ撃ち50_診断 に一本化（本番エンジンと同じ判定基準で高速診断）
    マクロ撃ち診断 = modMacroUchi50.マクロ撃ち50_診断(ws)
End Function

Sub 選んだ列の日付を和暦で表示する()
    ' 変えてよい: 文字日付
    ' 依頼の語: 日付を和暦|日付の和暦|和暦で表示|和暦表示|和暦にして|元号で表示|日付を元号|日付を令和|和暦に直|和暦に変
    ' 依頼の組: 日付,年月日+和暦,元号,令和,平成,表示
    ' 扱う: 和暦表示 和暦 元号 日付
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 選んでいる範囲の日付（シリアル値）の表示形式を和暦にする。値は変えない。文字の日付は触らずスキップする。

    Dim ws As Worksheet
    Dim area As Range
    Dim cell As Range
    Dim hdrRow As Long
    Dim lastRow As Long
    Dim col As Long, colIdx As Long, r As Long
    Dim maxRow As Long
    Dim isDateVal As Boolean
    Dim startR As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    maxRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1

    For Each area In Selection.Areas
        hdrRow = area.row
        If area.rows.Count > 1 Then
            lastRow = area.row + area.rows.Count - 1
            If lastRow > maxRow Then lastRow = maxRow
        Else
            lastRow = maxRow
        End If

        For col = 1 To area.Columns.Count
            colIdx = area.Column + col - 1
            If area.rows.Count = 1 Then
                startR = hdrRow + 1
            Else
                If VarType(ws.Cells(hdrRow, colIdx).Value) = vbString And Not IsDate(ws.Cells(hdrRow, colIdx).Value) Then
                    startR = hdrRow + 1
                Else
                    startR = hdrRow
                End If
            End If

            For r = startR To lastRow
                Set cell = ws.Cells(r, colIdx)
                If Not isEmpty(cell.Value) Then
                    isDateVal = False
                    If VarType(cell.Value) = vbDate Then
                        isDateVal = True
                    ElseIf VarType(cell.Value) = vbDouble Then
                        If (cell.Value >= 1 And cell.Value <= 73050) Or InStr(cell.NumberFormat, "y") > 0 Or InStr(cell.NumberFormat, "m") > 0 Or InStr(cell.NumberFormat, "d") > 0 Then
                            isDateVal = True
                        End If
                    End If

                    If isDateVal And Not cell.HasFormula Then
                        On Error Resume Next
                        cell.NumberFormatLocal = "ggge年m月d日"
                        If Err.Number <> 0 Then
                            Err.Clear
                            cell.NumberFormat = "[$-411]ggge""年""m""月""d""日"
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            cell.NumberFormatLocal = "ggge""年""m""月""d""日"
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next r
        Next col
    Next area
End Sub

Sub 表の合計と小計の行を消す()
    ' 変えてよい: 行移動
    ' 依頼の語: 合計行を消|小計行を消|合計と小計を消|小計を削除|合計行を削除|合計の行を消|小計の行を消|合計行と小計行を|小計行を削除|合計小計を削除|集計行を消|集計行を削除
    ' 依頼の組: 合計,小計,総計,総合計,集計+行+消,削除,取,クリア+-足,挿入,作
    ' 扱う: 合計行削除 小計行削除 合計
    ' 見出し: なし
    ' 選ぶ列: -
    ' 形: 数の列
    ' 表の中の、合計・総計・計・小計の行（見出しの列のどこかにその字がある行）を削除して、明細だけにする。前の月の表を使い回して集計し直すときの下ごしらえ

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim urTop As Long, urLeft As Long, urRight As Long, urBottom As Long
    urTop = ur.row
    urLeft = ur.Column
    urRight = ur.Column + ur.Columns.Count - 1
    urBottom = 表の本文の最終行(ws)

    ' 見出し行を探す
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

    If urBottom <= hdrRow Then Exit Sub

    ' 本体の表の列＝見出しの行で、最初の見出しから空の見出しの手前まで（2026-10-28 点検）。
    ' 前は使用範囲の全列で「合計」を探していたので、表の右の集計表の「合計」の行が本体の明細の行と同じ行だと、明細の行ごと消していた
    Dim hc1 As Long, hc2 As Long, 外の値 As Long, 残した As Long
    For c = urLeft To urRight
        If 表の整理_下請け.見出しの字(ws.Cells(hdrRow, c)) <> "" Then
            If hc1 = 0 Then hc1 = c
            hc2 = c
        ElseIf hc1 > 0 Then
            Exit For
        End If
    Next c
    If hc1 = 0 Then Exit Sub

    Dim nRows As Long
    nRows = urBottom - hdrRow
    Dim delFlags() As Boolean, hasDel As Boolean
    ReDim delFlags(hdrRow + 1 To urBottom)
    hasDel = False

    For r = urBottom To hdrRow + 1 Step -1
        If 集計の行か(ws, r, hc1, hc2) Then
            ' 行ごと消すので、同じ行の表の外（右・左の別の表やメモ）に値があれば消さずに数える（重複行の棚と同じ守り）
            外の値 = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, urLeft), ws.Cells(r, urRight))) _
                   - Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, hc1), ws.Cells(r, hc2)))
            If 外の値 = 0 Then
                delFlags(r) = True
                hasDel = True
            Else
                残した = 残した + 1
            End If
        End If
    Next r
    If 残した > 0 Then
        Application.StatusBar = "表の合計と小計の行を消す: 表の外の同じ行に値がある " & 残した & " 行は、巻き込まないよう消していません"
    End If

    ' 事前高速判定: 集計行が0件なら即終了（0秒判定）
    If Not hasDel Then Exit Sub

    Dim oldScreenUpdating As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    On Error GoTo CleanUp
    oldScreenUpdating = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim delRange As Range, batchStr As String
    batchStr = ""
    For r = urBottom To hdrRow + 1 Step -1
        If delFlags(r) Then
            ' 結合セルがあれば結合解除
            Dim maC As Long
            For maC = urLeft To urRight
                If ws.Cells(r, maC).MergeCells Then
                    ws.Cells(r, maC).MergeArea.UnMerge
                End If
            Next maC
            If Len(batchStr) > 200 Then
                If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
                batchStr = ""
            End If
            batchStr = batchStr & IIf(Len(batchStr) > 0, ",", "") & r & ":" & r
        End If
    Next r
    If Len(batchStr) > 0 Then
        If delRange Is Nothing Then Set delRange = ws.Range(batchStr) Else Set delRange = Union(delRange, ws.Range(batchStr))
    End If

    If Not delRange Is Nothing Then delRange.Delete Shift:=xlUp

CleanUp:
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 選んだ列で表を並べ替える()
    ' 変えてよい: 行移動
    ' 選んだ1列をキーに表全体（見出しの下から合計行の上まで）を昇順に並べ替える。頼みに「降順・大きい順・新しい順」があれば降順。合計・小計の行は動かさない。
    ' 階級: 基本
    ' 依頼の語: 表を並べ替え|表をソート|昇順に並べ替え|降順に並べ替え|表を昇順に|表を降順に|日付順に表を|新しい順に表を|古い順に表を|日付の古い順|古い順に並べ|金額の大きい順|大きい順に並べ|コード順に並べ替え|コード順に並べ|順に並べて
    ' 依頼の組: 並べ替え,ソート,並べて+表,古い順,新しい順,大きい順,コード順+-列を,列の,選んだ順,左へ,0の行,ゼロ,テーブル,何位,順位,ピボット,ピポット
    ' 扱う: 並べ替え ソート 昇順 降順 表並べ替え
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: なし
    ' 列は頼みから: あり

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If ws.UsedRange Is Nothing Then Exit Sub

    Dim hdrRow As Long, keyCol As Long
    If Selection Is Nothing Then
        hdrRow = ws.UsedRange.row
        keyCol = ws.UsedRange.Column
    Else
        hdrRow = Selection.Cells(1, 1).row
        keyCol = Selection.Cells(1, 1).Column
    End If

    Dim usedLeft As Long, usedRight As Long
    usedLeft = ws.UsedRange.Column
    usedRight = usedLeft + ws.UsedRange.Columns.Count - 1

    Dim firstDataRow As Long, lastRow As Long
    firstDataRow = hdrRow + 1
    Dim lastUsedRow As Long
    lastUsedRow = 表の本文の最終行(ws)

    ' 合計行を探す（firstDataRow以降で集計の語を含むセル）
    Dim totRow As Long, i As Long, cc As Long, s As String
    totRow = 0
    Dim scanFrom As Long
    scanFrom = lastUsedRow + 3
    If scanFrom > ws.UsedRange.row + ws.UsedRange.rows.Count - 1 Then scanFrom = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    For i = scanFrom To firstDataRow Step -1
        For cc = usedLeft To usedRight
            s = CStr(ws.Cells(i, cc).Value)
            If 集計の語か(s) Then
                totRow = i
                Exit For
            End If
        Next cc
        If totRow > 0 Then Exit For
    Next i

    If totRow > 0 Then
        lastRow = totRow - 1
        Do While lastRow > firstDataRow And Application.WorksheetFunction.CountA(ws.Range(ws.Cells(lastRow, usedLeft), ws.Cells(lastRow, usedRight))) = 0
            lastRow = lastRow - 1
        Loop
    Else
        lastRow = lastUsedRow
    End If

    If lastRow <= firstDataRow Then Exit Sub

    ' 頼みの文の判定（降順判定）
    Dim isDesc As Boolean
    isDesc = False
    Dim req As String
    req = ""
    On Error Resume Next
    req = コンボ道具.m頼みの文
    On Error GoTo 0
    If InStr(req, "降順") > 0 Or InStr(req, "大きい順") > 0 Or InStr(req, "新しい順") > 0 Or InStr(req, "高い順") > 0 Or InStr(req, "多い順") > 0 Then
        isDesc = True
    End If

    Dim sortOrder As XlSortOrder
    If isDesc Then sortOrder = xlDescending Else sortOrder = xlAscending

    ' 並べ替え実行
    Dim targetRng As Range
    Set targetRng = ws.Range(ws.Cells(firstDataRow, usedLeft), ws.Cells(lastRow, usedRight))
    targetRng.Sort Key1:=ws.Cells(firstDataRow, keyCol), Order1:=sortOrder, Header:=xlNo, Orientation:=xlTopToBottom
End Sub

Sub 見出しで枠を固定しA1に戻す()
    ' 全シートで、表の見出しの下で枠を固定し、選択を A1、表示を左上へ。値は変えない。
    ' 階級: 基本
    ' 依頼の語: 枠を固定|枠固定|A1に戻|A1選択|全シートの枠|全シートA1|見出しで枠固定|ウィンドウ枠固定|枠の固定|A1にして保存|A1にして|見出しが見えるように|スクロールしても見出し
    ' 依頼の組: 枠,固定,A1+シート,見出し,戻,そろえ+-CSV,文字数字,印刷,ピボット
    ' 扱う: 枠固定 A1 全シート
    ' 見出し: なし
    ' 選ぶ列: なし
    ' 形: なし

    Dim ws As Worksheet, actWs As Worksheet
    Dim origScreen As Boolean
    Set actWs = ActiveSheet
    origScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim sh As Worksheet
    For Each sh In actWs.Parent.Worksheets
        If sh.Visible = xlSheetVisible Then
            sh.Activate
            ActiveWindow.FreezePanes = False

            ' 見出し行の判定
            Dim hdrRow As Long
            hdrRow = 1
            If Not sh.UsedRange Is Nothing Then
                hdrRow = sh.UsedRange.row
            End If

            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1

            sh.Cells(hdrRow + 1, 1).Select
            ActiveWindow.FreezePanes = True

            sh.Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next sh

    If Not actWs Is Nothing Then
        If actWs.Visible = xlSheetVisible Then actWs.Activate
    End If
    Application.ScreenUpdating = origScreen
End Sub

Sub 選んだ列の式を値にする()
    ' 変えてよい: 数式
    ' 選んだ列（選ばなければ表全体）の式を、今の値に置き換える。書式は変えない。エラーの式は値にせず数を報告。
    ' 階級: 基本
    ' 依頼の語: 式を値に|数式を値に|式を値にする|数式を値にする|式を値に変換|数式を値に変換|計算式を値に|式を固定|値貼り付け|値貼付|数式を消して値|関数を消して値|値だけにして|式を消して値|関数を値に
    ' 依頼の組: 式,数式,計算式,関数+値,貼り付け,貼付,固定,変換,値だけ+-外部リンク,別ブック,参照
    ' 扱う: 式を値 数式値化 値貼り付け
    ' 見出し: なし
    ' 選ぶ列: 0-1
    ' 形: なし

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If ws.UsedRange Is Nothing Then Exit Sub

    Dim targetRng As Range
    If Selection Is Nothing Then
        Set targetRng = ws.UsedRange
    ElseIf Selection.CountLarge = 1 And Selection.Address = ws.Cells(1, 1).Address Then
        Set targetRng = ws.UsedRange
    Else
        Set targetRng = Selection
    End If

    ' 数式を含むセルを探す
    Dim formulaRng As Range
    On Error Resume Next
    Set formulaRng = targetRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaRng Is Nothing Then Exit Sub

    Dim cell As Range, errCount As Long, convertedCount As Long
    Dim spillRng As Range
    errCount = 0
    convertedCount = 0

        For Each cell In formulaRng
            If IsError(cell.Value) Then
                errCount = errCount + 1
            Else
                Set spillRng = 表の整理_下請け.スピルの元なら範囲(cell)
                If Not spillRng Is Nothing Then
                    ' 動的配列の式（スピル）は広がった範囲ごと値にする。元のセルだけ値にすると、広がった側の値が消えていた（2026-10-28 点検）
                    表の整理_下請け.値を書き戻す spillRng, spillRng.Value
                ElseIf VarType(cell.Value) = vbString Then
                    cell.Value = 表の整理_下請け.文字のまま書く字(CStr(cell.Value))
                Else
                    cell.Value = cell.Value
                End If
                convertedCount = convertedCount + 1
            End If
        Next cell

    If errCount > 0 Then
        ' エラーの式は値にせず数を報告
        Call 先撃ちに断る("エラーの式 " & errCount & " 件は値にせず残しました（" & convertedCount & " 件を値に変換）")
    End If
End Sub

Private Sub シートの行高と列幅を写す_本体(ByVal 参照先Index As Long)
    Dim src As Worksheet, r As Long, c As Long
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Index = 1 Then Exit Sub
    If 参照先Index < 1 Or 参照先Index >= ActiveSheet.Index Then Exit Sub
    If TypeName(ActiveWorkbook.Sheets(参照先Index)) <> "Worksheet" Then Exit Sub
    Set src = ActiveWorkbook.Sheets(参照先Index)
    Application.ScreenUpdating = False
    For r = 1 To src.UsedRange.row + src.UsedRange.rows.Count - 1
        ActiveSheet.rows(r).RowHeight = src.rows(r).RowHeight
    Next r
    For c = 1 To src.UsedRange.Column + src.UsedRange.Columns.Count - 1
        ActiveSheet.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c
    Application.ScreenUpdating = True
End Sub

Sub 選んだコード列の右に名称列を足す()
    ' 変えてよい: 空欄 列移動
    ' 選んでいるコード列の右隣に1列挿入し、対応表（マスタ）を参照して名称を入力する。元のコード列は残す。
    ' 依頼の語: コードの右に名称|コードの右に名前|名称列を挿入|名称列を足|名前の列を挿入|名前の列を足|マスタから名称を足|対応表から名称を足|コードから名称列
    ' 依頼の組: コード+名称,名前+右,足す,挿入,追加,列
    ' 扱う: コード名称追加 マスタ参照 列挿入
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet, selCol As Long, headerRow As Long, lastRow As Long, maxRow As Long
    Dim i As Long, s As String
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If ws.UsedRange Is Nothing Then Exit Sub

    selCol = Selection.Cells(1, 1).Column
    headerRow = Selection.Cells(1, 1).row
    maxRow = 表の本文の最終行(ws)
    If Selection.rows.Count > 1 Then
        lastRow = Selection.row + Selection.rows.Count - 1
        If lastRow > maxRow Then lastRow = maxRow
    Else
        lastRow = maxRow
    End If
    If lastRow <= headerRow Then Exit Sub

    Dim oldScreen As Boolean, oldCalc As XlCalculation, oldEvents As Boolean
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim cnt As Long: cnt = lastRow - headerRow

    Dim srcRange As Range
    Set srcRange = ws.Range(ws.Cells(headerRow + 1, selCol), ws.Cells(lastRow, selCol))
    Dim vCode As Variant
    If cnt = 1 Then
        ReDim vCode(1 To 1, 1 To 1)
        vCode(1, 1) = srcRange.Value2
    Else
        vCode = srcRange.Value2
    End If

    ' 1. サンプル値収集（コード列先頭10件）
    Dim sampleVals As Object
    Set sampleVals = CreateObject("Scripting.Dictionary")
    sampleVals.CompareMode = 1
    Dim sampleCount As Long: sampleCount = 0
    For i = 1 To cnt
        If セルの字(vCode(i, 1)) <> "" Then
            s = CStr(vCode(i, 1)): GoSub 正規化
            If Not sampleVals.exists(s) Then
                sampleVals.Add s, 1
                sampleCount = sampleCount + 1
                If sampleCount >= 10 Then Exit For
            End If
        End If
    Next i

    ' 2. 対応表シートを探す（メモリ配列走査）
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim bestSheet As Worksheet, bestKeyCol As Long, bestValCol As Long, bestMatch As Long: bestMatch = 0
    Dim targetSheet As Worksheet

    For Each targetSheet In wb.Sheets
        If targetSheet.Name = ws.Name Then GoTo NextSheet
        Dim tUR As Range: Set tUR = targetSheet.UsedRange
        If tUR Is Nothing Then GoTo NextSheet
        If tUR.Columns.Count < 2 Then GoTo NextSheet
        Dim tFirstCol As Long, tLastCol As Long, tRows As Long
        tFirstCol = tUR.Column: tLastCol = tUR.Column + tUR.Columns.Count - 1
        tRows = tUR.rows.Count
        If tLastCol - tFirstCol <> 1 Then GoTo NextSheet

        Dim tArr As Variant
        If tRows = 1 Then
            ReDim tArr(1 To 1, 1 To 2)
            tArr(1, 1) = tUR.Cells(1, 1).Value2
            tArr(1, 2) = tUR.Cells(1, 2).Value2
        Else
            tArr = tUR.Value2
        End If

        Dim matchCount As Long: matchCount = 0
        For i = 1 To tRows
            s = CStr(tArr(i, 1)): GoSub 正規化
            If sampleVals.exists(s) Then matchCount = matchCount + 1
        Next i
        If matchCount > bestMatch Then
            bestMatch = matchCount: Set bestSheet = targetSheet: bestKeyCol = tFirstCol: bestValCol = tFirstCol + 1
        End If
NextSheet:
    Next targetSheet

    ' 自動検出で見つからなければ小窓
    If bestMatch = 0 Or bestSheet Is Nothing Then
        Application.ScreenUpdating = True
        Set bestSheet = シート選び直しフォーム.選択("コード列に対応する名称のシート（マスタ）を選んでください", wb, ws.Name)
        Unload シート選び直しフォーム
        Application.ScreenUpdating = False
        If bestSheet Is Nothing Then
            Call 先撃ちに断る("対応表シートが選ばれませんでした")
            GoTo CleanUp
        End If
        bestKeyCol = bestSheet.UsedRange.Column
        bestValCol = bestKeyCol + 1
    End If

    ' 3. 対応表から辞書構築（メモリ配列走査）
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim mUR As Range: Set mUR = bestSheet.UsedRange
    Dim mRows As Long, mCols As Long
    mRows = mUR.rows.Count
    mCols = mUR.Columns.Count

    Dim mArr As Variant
    If mRows = 1 And mCols = 1 Then
        ReDim mArr(1 To 1, 1 To 1)
        mArr(1, 1) = mUR.Value2
    Else
        mArr = mUR.Value2
    End If

    Dim relKeyCol As Long, relValCol As Long
    relKeyCol = bestKeyCol - mUR.Column + 1
    relValCol = bestValCol - mUR.Column + 1

    Dim masterHeader As String: masterHeader = ""
    If mRows >= 1 Then
        masterHeader = Trim$(CStr(mArr(1, relValCol)))
    End If
    If Len(masterHeader) = 0 Then masterHeader = "名称"

    For i = 1 To mRows
        s = CStr(mArr(i, relKeyCol)): GoSub 正規化
        If Len(s) > 0 And Not dict.exists(s) Then
            dict.Add s, CStr(mArr(i, relValCol))
        End If
    Next i

    ' 4. 列挿入と一括代入（最速・掟10）
    ws.Columns(selCol + 1).Insert Shift:=xlToRight
    ws.Cells(headerRow, selCol + 1).Value = masterHeader
    ws.Cells(headerRow, selCol).Copy
    ws.Cells(headerRow, selCol + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    Dim vName() As Variant: ReDim vName(1 To cnt, 1 To 1)

    For i = 1 To cnt
        s = CStr(vCode(i, 1)): GoSub 正規化
        If dict.exists(s) Then
            vName(i, 1) = dict(s)
        Else
            vName(i, 1) = ""
        End If
    Next i

    表の整理_下請け.文字を守って書く ws.Range(ws.Cells(headerRow + 1, selCol + 1), ws.Cells(lastRow, selCol + 1)), vName ' 書き戻し: 新しい所
    ws.Columns(selCol + 1).AutoFit

CleanUp:
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Exit Sub

正規化:
    s = Trim$(s)
    s = StrConv(s, vbNarrow)
    Return
End Sub

Sub 別シートと突合して不一致と新規を色付けする()
    ' キー列を基準に別シートと突合し、相手にない新規行を薄緑、値が異なるセルを薄黄でハイライトする。
    ' 依頼の語: 突合して色付け|突合して色|差分を色付け|差分をハイライト|違いを色付け|違いをハイライト|変更箇所を色|変わった所を色|前月と突き合わせて色|前月との違いを色|新旧の差を色
    ' 依頼の組: 突合,突き合,突合し,照合,差分,違い,変更+色,ハイライト,塗り,着色
    ' 扱う: 突合 差分 ハイライト 色付け
    ' 見出し: なし
    ' 選ぶ列: -
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet, wb As Workbook
    Set ws = ActiveSheet
    Set wb = ws.Parent
    If TypeName(ws) <> "Worksheet" Then Exit Sub

    Dim ur As Range: Set ur = ws.UsedRange
    Dim 上行 As Long, 左列 As Long, 下行 As Long, 右列 As Long
    上行 = ur.row: 左列 = ur.Column
    下行 = 上行 + ur.rows.Count - 1: 右列 = 左列 + ur.Columns.Count - 1
    If 下行 <= 上行 Then Exit Sub

    ' 1. 見出し行を探す
    Dim hrow As Long: hrow = 上行
    Dim dR1 As Long: dR1 = hrow + 1
    Dim dR2 As Long: dR2 = 下行
    Dim maxRow As Long: maxRow = 表の本文の最終行(ws)
    If maxRow > hrow And maxRow < dR2 Then dR2 = maxRow

    ' 2. キー列を決める（選択範囲が1列ならそれ、無ければ1列目）
    Dim colKey As Long: colKey = 左列
    If Selection.Columns.Count = 1 And Selection.Cells(1, 1).Column >= 左列 And Selection.Cells(1, 1).Column <= 右列 Then
        colKey = Selection.Cells(1, 1).Column
    End If

    ' 3. 相手シートを探す（010の型を活用）
    Dim targetSheet As Worksheet, oppSheet As Worksheet
    Dim bestScore As Long: bestScore = 0
    For Each targetSheet In wb.Worksheets
        If targetSheet.Name <> ws.Name And targetSheet.Visible = xlSheetVisible Then
            Dim tUR As Range: Set tUR = targetSheet.UsedRange
            ' 見出しの一致数を簡易スコアリング
            Dim matchHdr As Long, c As Long
            matchHdr = 0
            For c = 左列 To Application.Min(右列, tUR.Column + tUR.Columns.Count - 1)
                If Trim$(CStr(ws.Cells(hrow, c).Value)) <> "" And Trim$(CStr(ws.Cells(hrow, c).Value)) = Trim$(CStr(targetSheet.Cells(hrow, c).Value)) Then
                    matchHdr = matchHdr + 1
                End If
            Next c
            If matchHdr > bestScore Then
                bestScore = matchHdr
                Set oppSheet = targetSheet
            End If
        End If
    Next targetSheet

    If oppSheet Is Nothing Or bestScore = 0 Then
        Set oppSheet = シート選び直しフォーム.選択("突き合わせる相手のシート（前月データなど）を選んでください", wb, ws.Name)
        Unload シート選び直しフォーム
        If oppSheet Is Nothing Then
            Call 先撃ちに断る("突合相手のシートが選ばれませんでした")
            Exit Sub
        End If
    End If

    ' 4. 相手シートのデータをメモリ辞書へ格納（キー -> 行番号）
    Dim oppUR As Range: Set oppUR = oppSheet.UsedRange
    Dim oppLastRow As Long: oppLastRow = 表の本文の最終行(oppSheet)
    If oppLastRow <= hrow Then Exit Sub

    Dim dictOppRow As Object: Set dictOppRow = CreateObject("Scripting.Dictionary")
    dictOppRow.CompareMode = 1
    Dim r As Long, sKey As String
    For r = hrow + 1 To oppLastRow
        sKey = Trim$(CStr(oppSheet.Cells(r, colKey).Value))
        If Len(sKey) > 0 And Not dictOppRow.exists(sKey) Then
            dictOppRow.Add sKey, r
        End If
    Next r

    ' 5. メモリ配列で高速突合・色付け範囲を Union
    Dim vCur As Variant, vOpp As Variant
    vCur = ws.Range(ws.Cells(hrow + 1, 左列), ws.Cells(dR2, 右列)).Value2
    vOpp = oppSheet.Range(oppSheet.Cells(hrow + 1, 左列), oppSheet.Cells(oppLastRow, 右列)).Value2

    Dim rngNew As Range, rngDiff As Range
    Dim curRowIdx As Long, oppR As Long, cIdx As Long
    Dim v1 As Variant, v2 As Variant

    For curRowIdx = 1 To dR2 - hrow
        sKey = Trim$(CStr(vCur(curRowIdx, colKey - 左列 + 1)))
        If Len(sKey) > 0 Then
            If Not dictOppRow.exists(sKey) Then
                ' 新規行: 行全体を薄緑
                If rngNew Is Nothing Then
                    Set rngNew = ws.Range(ws.Cells(hrow + curRowIdx, 左列), ws.Cells(hrow + curRowIdx, 右列))
                Else
                    Set rngNew = Union(rngNew, ws.Range(ws.Cells(hrow + curRowIdx, 左列), ws.Cells(hrow + curRowIdx, 右列)))
                End If
            Else
                ' 既存行: 各セルを比較して不一致セルを薄黄
                oppR = dictOppRow(sKey) - hrow
                For cIdx = 1 To 右列 - 左列 + 1
                    v1 = vCur(curRowIdx, cIdx)
                    v2 = vOpp(oppR, cIdx)
                    If CStr(v1) <> CStr(v2) Then
                        If rngDiff Is Nothing Then
                            Set rngDiff = ws.Cells(hrow + curRowIdx, 左列 + cIdx - 1)
                        Else
                            Set rngDiff = Union(rngDiff, ws.Cells(hrow + curRowIdx, 左列 + cIdx - 1))
                        End If
                    End If
                Next cIdx
            End If
        End If
    Next curRowIdx

    ' 6. 一括着色（最速・掟10）
    Application.ScreenUpdating = False
    If Not rngNew Is Nothing Then
        rngNew.Interior.Color = RGB(220, 245, 220) ' 薄い緑（新規）
    End If
    If Not rngDiff Is Nothing Then
        rngDiff.Interior.Color = RGB(255, 235, 200) ' 薄い黄（変更）
    End If
    Application.ScreenUpdating = True
End Sub

Sub 横持ちマトリクス表を一撃で縦持ちデータベースにする()
    ' 横持ちのクロス集計表を行見出し・列見出し・値の縦持ち正規化テーブルに変換し、新シートに出力する。
    ' 依頼の語: アンピボット|クロス表を縦持ち|マトリクスを縦持ち|クロス集計を縦持ち|一撃で縦持ち|マトリクス表を縦持ち|マトリクスをデータベースに|クロス表を正規化
    ' 依頼の組: マトリクス,クロス集計,クロス表+縦持ち,アンピボット,正規化
    ' 扱う: アンピボット マトリクス 縦持ち変換 正規化
    ' 見出し: なし
    ' 選ぶ列: -
    ' 形: 数の列 別のシート

    Dim ws As Worksheet, wb As Workbook
    Set ws = ActiveSheet
    Set wb = ws.Parent
    If TypeName(ws) <> "Worksheet" Then Exit Sub

    Dim ur As Range: Set ur = ws.UsedRange
    Dim 上行 As Long, 左列 As Long, 下行 As Long, 右列 As Long
    上行 = ur.row: 左列 = ur.Column
    下行 = 上行 + ur.rows.Count - 1: 右列 = 左列 + ur.Columns.Count - 1
    If 下行 <= 上行 Or 右列 <= 左列 Then Exit Sub

    ' 1. 見出し行とデータ最終行
    Dim hrow As Long: hrow = 上行
    Dim lastRow As Long: lastRow = 表の本文の最終行(ws)
    If lastRow <= hrow Then lastRow = 下行

    ' 2. 左端のデータ数値列（展開対象の先頭列）を自動判定
    ' 各列で hrow+1 以降に数値が多く入っている最初の列を探す
    Dim firstDataCol As Long: firstDataCol = 0
    Dim c As Long, r As Long, numCount As Long, checkRows As Long
    checkRows = Application.Min(lastRow, hrow + 10)
    For c = 左列 To 右列
        numCount = 0
        For r = hrow + 1 To checkRows
            If IsNumeric(ws.Cells(r, c).Value) And Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then
                numCount = numCount + 1
            End If
        Next r
        If numCount >= 2 Then
            firstDataCol = c
            Exit For
        End If
    Next c

    ' 数値列が見つからなければ、左端1列を除いた2列目を既定とする
    If firstDataCol <= 左列 Then firstDataCol = 左列 + 1
    If firstDataCol > 右列 Then Exit Sub

    Dim nFixedCols As Long: nFixedCols = firstDataCol - 左列
    Dim nUnpivotCols As Long: nUnpivotCols = 右列 - firstDataCol + 1
    Dim nDataRows As Long: nDataRows = lastRow - hrow
    Dim maxRecords As Long: maxRecords = nDataRows * nUnpivotCols
    If maxRecords <= 0 Then Exit Sub

    ' 3. メモリ配列で元データを一括取得（掟10）
    Dim vSrc As Variant
    vSrc = ws.Range(ws.Cells(hrow, 左列), ws.Cells(lastRow, 右列)).Value2

    ' 4. 結果格納配列を一括確保
    Dim vDest() As Variant
    ReDim vDest(1 To maxRecords, 1 To nFixedCols + 2)

    Dim recIdx As Long: recIdx = 0
    Dim rIdx As Long, cIdx As Long, f As Long
    Dim valCell As Variant

    For rIdx = 2 To nDataRows + 1
        For cIdx = nFixedCols + 1 To nFixedCols + nUnpivotCols
            valCell = vSrc(rIdx, cIdx)
            ' 空白セルはスキップして実データのみをスマートに縦持ち化
            If Len(Trim$(CStr(valCell))) > 0 Then
                recIdx = recIdx + 1
                ' 固定項目列（部署・品名など）を転写
                For f = 1 To nFixedCols
                    vDest(recIdx, f) = vSrc(rIdx, f)
                Next f
                ' 展開項目名（列ヘッダー文字: 4月、5月など）
                vDest(recIdx, nFixedCols + 1) = vSrc(1, cIdx)
                ' 値
                vDest(recIdx, nFixedCols + 2) = valCell
            End If
        Next cIdx
    Next rIdx

    If recIdx = 0 Then Exit Sub

    ' 5. 新シートを作成して一撃出力（掟10）
    Application.ScreenUpdating = False
    Dim newWs As Worksheet, sheetBaseName As String, newName As String, seq As Long
    Dim wsCheck As Worksheet, exists As Boolean
    sheetBaseName = Left$(ws.Name, 25) & "_縦持ち"
    newName = sheetBaseName
    seq = 1
    Do
        exists = False
        For Each wsCheck In wb.Worksheets
            If wsCheck.Name = newName Then
                exists = True
                Exit For
            End If
        Next wsCheck
        If Not exists Then Exit Do
        seq = seq + 1
        newName = Left$(ws.Name, 23) & "_縦持ち" & seq
    Loop

    Set newWs = wb.Worksheets.Add(after:=ws)
    newWs.Name = newName

    ' 新シートの見出し行
    For f = 1 To nFixedCols
        newWs.Cells(1, f).Value = vSrc(1, f)
    Next f
    newWs.Cells(1, nFixedCols + 1).Value = "項目"
    newWs.Cells(1, nFixedCols + 2).Value = "値"

    ' 書式設定（見出し太字・背景色・枠線）
    With newWs.Range(newWs.Cells(1, 1), newWs.Cells(1, nFixedCols + 2))
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

    ' データ一撃代入
    表の整理_下請け.文字を守って書く newWs.Range(newWs.Cells(2, 1), newWs.Cells(recIdx + 1, nFixedCols + 2)), vDest ' 書き戻し: 新しい所

    ' 数値列の書式と列幅調整
    newWs.Columns(nFixedCols + 2).NumberFormatLocal = "#,##0"
    newWs.Columns.AutoFit
    newWs.Activate
    newWs.Range("A2").Select
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
End Sub

Sub 表記ゆれをあいまい判定してマスタを引き当てる()
    ' 表記ゆれやあいまいな名称を2-gram類似度で判定し、最も近いマスタの値を右隣に挿入・引き当てる。
    ' 依頼の語: 表記ゆれを引き当て|あいまい検索でマスタ|あいまい突合|ファジー検索|ファジーVLOOKUP|ゆれを吸収してマスタ|名寄せしてマスタ|似た名前で引き当て|あいまい引き当て
    ' 依頼の組: 表記ゆれ,あいまい,ファジー,名寄せ+マスタ,引き当て,引く,参照,突合
    ' 扱う: 表記ゆれ あいまい引き当て ファジー マスタ参照
    ' 見出し: なし
    ' 選ぶ列: 1
    ' 形: 数の列 別のシート 共通の見出しのシート

    Dim ws As Worksheet, wb As Workbook
    Set ws = ActiveSheet
    Set wb = ws.Parent
    If TypeName(ws) <> "Worksheet" Then Exit Sub

    Dim selCol As Long, headerRow As Long, lastRow As Long, maxRow As Long
    selCol = Selection.Cells(1, 1).Column
    headerRow = Selection.Cells(1, 1).row
    maxRow = 表の本文の最終行(ws)
    If Selection.rows.Count > 1 Then
        lastRow = Selection.row + Selection.rows.Count - 1
        If lastRow > maxRow Then lastRow = maxRow
    Else
        lastRow = maxRow
    End If
    If lastRow <= headerRow Then Exit Sub

    ' 1. サンプル値収集
    Dim sampleVals As Object: Set sampleVals = CreateObject("Scripting.Dictionary")
    sampleVals.CompareMode = 1
    Dim i As Long, s As String, sampleCount As Long: sampleCount = 0
    For i = headerRow + 1 To lastRow
        s = CStr(ws.Cells(i, selCol).Value): GoSub 基本正規化
        If Len(s) > 0 And Not sampleVals.exists(s) Then
            sampleVals.Add s, 1
            sampleCount = sampleCount + 1
            If sampleCount >= 10 Then Exit For
        End If
    Next i

    ' 2. 対応表シートを探す（010/012と同様）
    Dim bestSheet As Worksheet, bestKeyCol As Long, bestValCol As Long, bestMatch As Long: bestMatch = 0
    Dim targetSheet As Worksheet
    For Each targetSheet In wb.Sheets
        If targetSheet.Name <> ws.Name Then
            Dim tUR As Range: Set tUR = targetSheet.UsedRange
            If tUR.Columns.Count >= 2 Then
                Dim tFR As Long: tFR = tUR.row
                Dim tLR As Long: tLR = tUR.row + tUR.rows.Count - 1
                Dim tFC As Long: tFC = tUR.Column
                Dim tLC As Long: tLC = tUR.Column + tUR.Columns.Count - 1
                If tLC - tFC = 1 Then
                    Dim c As Long: c = tFC
                    Dim matchCount As Long: matchCount = 0
                    For i = tFR To tLR
                        s = CStr(targetSheet.Cells(i, c).Value): GoSub 基本正規化
                        If sampleVals.exists(s) Then matchCount = matchCount + 1
                    Next i
                    If matchCount > bestMatch Then
                        bestMatch = matchCount: Set bestSheet = targetSheet: bestKeyCol = c: bestValCol = c + 1
                    End If
                End If
            End If
        End If
    Next targetSheet

    If bestMatch = 0 Or bestSheet Is Nothing Then
        Set bestSheet = シート選び直しフォーム.選択("照合するマスタのシートを選んでください", wb, ws.Name)
        Unload シート選び直しフォーム
        If bestSheet Is Nothing Then
            Call 先撃ちに断る("マスタシートが選ばれませんでした")
            Exit Sub
        End If
        bestKeyCol = bestSheet.UsedRange.Column
        bestValCol = bestKeyCol + 1
    End If

    ' 3. マスタの全件取得とインデックス化
    Dim mUR As Range: Set mUR = bestSheet.UsedRange
    Dim mFR As Long: mFR = mUR.row
    Dim mLR As Long: mLR = mUR.row + mUR.rows.Count - 1
    Dim masterHeader As String: masterHeader = ""
    If mFR <= mLR Then masterHeader = Trim$(CStr(bestSheet.Cells(mFR, bestValCol).Value))
    If Len(masterHeader) = 0 Then masterHeader = "名称（照合）"

    Dim mCount As Long: mCount = mLR - mFR + 1
    Dim mKeys() As String, mVals() As String, mNorms() As String
    ReDim mKeys(1 To mCount), mVals(1 To mCount), mNorms(1 To mCount)
    Dim dictExact As Object: Set dictExact = CreateObject("Scripting.Dictionary")
    dictExact.CompareMode = 1

    Dim kIdx As Long: kIdx = 0
    Dim rawK As String, rawV As String
    For i = mFR To mLR
        rawK = CStr(bestSheet.Cells(i, bestKeyCol).Value)
        rawV = CStr(bestSheet.Cells(i, bestValCol).Value)
        If Len(Trim$(rawK)) > 0 Then
            kIdx = kIdx + 1
            mKeys(kIdx) = rawK
            mVals(kIdx) = rawV
            s = rawK: GoSub 強力正規化
            mNorms(kIdx) = s
            If Not dictExact.exists(s) Then dictExact.Add s, rawV
        End If
    Next i
    mCount = kIdx
    If mCount = 0 Then Exit Sub

    ' 4. 列挿入と一括判定（最速・掟10）
    Application.ScreenUpdating = False
    ws.Columns(selCol + 1).Insert Shift:=xlToRight
    ws.Cells(headerRow, selCol + 1).Value = masterHeader
    ws.Cells(headerRow, selCol).Copy
    ws.Cells(headerRow, selCol + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    Dim cnt As Long: cnt = lastRow - headerRow
    Dim vCode As Variant
    vCode = ws.Range(ws.Cells(headerRow + 1, selCol), ws.Cells(lastRow, selCol)).Value2
    Dim vName() As Variant: ReDim vName(1 To cnt, 1 To 1)
    Dim fuzzyFlags() As Boolean: ReDim fuzzyFlags(1 To cnt)
    Dim fuzzyNotes() As String: ReDim fuzzyNotes(1 To cnt)

    Dim queryNorm As String, bestScore As Double, currentScore As Double, bestTargetIdx As Long
    Dim qLen As Long, mLen As Long, mJ As Long
    For i = 1 To cnt
        s = CStr(vCode(i, 1)): GoSub 強力正規化
        queryNorm = s
        If Len(queryNorm) = 0 Then
            vName(i, 1) = ""
        ElseIf dictExact.exists(queryNorm) Then
            ' 完全一致（正規化含む）
            vName(i, 1) = dictExact(queryNorm)
        Else
            ' 2-gram 類似度によるファジー検索
            bestScore = 0: bestTargetIdx = 0
            qLen = Len(queryNorm)
            For mJ = 1 To mCount
                mLen = Len(mNorms(mJ))
                If InStr(mNorms(mJ), queryNorm) > 0 Or InStr(queryNorm, mNorms(mJ)) > 0 Then
                    currentScore = 0.85
                Else
                    ' 2-gram Jaccard類似度
                    Dim nGramMatches As Long: nGramMatches = 0
                    Dim totalPairs As Long: totalPairs = 0
                    Dim p As Long, gram As String
                    If qLen >= 2 And mLen >= 2 Then
                        totalPairs = (qLen - 1) + (mLen - 1)
                        For p = 1 To qLen - 1
                            gram = Mid$(queryNorm, p, 2)
                            If InStr(mNorms(mJ), gram) > 0 Then nGramMatches = nGramMatches + 1
                        Next p
                        If totalPairs > 0 Then
                            currentScore = (nGramMatches * 2#) / CDbl(totalPairs)
                        Else
                            currentScore = 0#
                        End If
                    Else
                        currentScore = 0#
                    End If
                End If
                If currentScore > bestScore Then
                    bestScore = currentScore
                    bestTargetIdx = mJ
                End If
            Next mJ

            ' 類似度 0.4 (40%) 以上なら採用し、薄オレンジ印
            If bestScore >= 0.4 And bestTargetIdx > 0 Then
                vName(i, 1) = mVals(bestTargetIdx)
                fuzzyFlags(i) = True
                fuzzyNotes(i) = "類似度: " & Format$(bestScore * 100, "0") & "% (候補: " & mKeys(bestTargetIdx) & ")"
            Else
                vName(i, 1) = ""
            End If
        End If
    Next i

    ' 5. データ代入とハイライト（最速・Union）
    表の整理_下請け.文字を守って書く ws.Range(ws.Cells(headerRow + 1, selCol + 1), ws.Cells(lastRow, selCol + 1)), vName ' 書き戻し: 新しい所

    Dim rngFuzzy As Range
    For i = 1 To cnt
        If fuzzyFlags(i) Then
            Dim cellTarget As Range: Set cellTarget = ws.Cells(headerRow + i, selCol + 1)
            If rngFuzzy Is Nothing Then Set rngFuzzy = cellTarget Else Set rngFuzzy = Union(rngFuzzy, cellTarget)
            On Error Resume Next
            cellTarget.ClearComments
            cellTarget.AddComment fuzzyNotes(i)
            On Error GoTo 0
        End If
    Next i

    If Not rngFuzzy Is Nothing Then
        rngFuzzy.Interior.Color = RGB(255, 240, 210) ' 薄いオレンジ（あいまい一致要確認）
    End If

    ws.Columns(selCol + 1).AutoFit
    Application.ScreenUpdating = True
    Exit Sub

基本正規化:
    s = Trim$(s)
    s = StrConv(s, vbNarrow)
    Return

強力正規化:
    s = Trim$(s)
    s = StrConv(s, vbNarrow)
    s = UCase$(s)
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "-", "")
    s = Replace(s, "ー", "")
    s = Replace(s, "株式会社", "")
    s = Replace(s, "(株)", "")
    s = Replace(s, "（株）", "")
    s = Replace(s, "有限会社", "")
    s = Replace(s, "(有)", "")
    s = Replace(s, "（有）", "")
    Return
End Sub

Sub 選んだ列の値ごとにシートを分割する()
    ' 依頼の語: 部署ごとに分割|支店ごとに分割|値ごとにシート|列の値でシート|シート別にする|グループごとにシート|カテゴリごとにシート
    ' 依頼の組: ワークシート+分割,分ける,振り分け,別+-行,列,結合
    ' 扱う: シート分割 部署別 振り分け
    ' 見出し: 選択列または表の見出し
    ' 形: 表全体
    ' 選んだ列（またはアクティブセルがある列）のユニークな値ごとに新シートを作成し、
    ' 元の表の見出し行＋その値に一致する行データを一括転記して振り分ける。
    ' 1万行でも破綻しないようメモリ配列で一撃抽出し、元シートの後ろにシートを並べる。
    ' 中身は 選んだ列の値ごとにシートを分割する_本体（引数のある手伝い）。引数で目録とマクロ一覧から隠していたのをやめ、入口を引数なしにした（2026-10-29）
    Call 選んだ列の値ごとにシートを分割する_本体(True)
End Sub

Sub 選んだ列の数値を指定桁数でゼロ埋め文字列にする()
    ' 変えてよい: 文字数字
    ' 依頼の語: ゼロ埋め文字列|指定桁数でゼロ|0埋め文字列|ゼロパディング|前ゼロ復元|先頭ゼロを復元|0埋めにする|ゼロ埋めにする|桁数をそろえてゼロ埋め
    ' 依頼の組: ゼロ,0+埋め,パディング,復元+文字列,文字,桁+-行,小数
    ' 扱う: ゼロ埋め パディング コード復元
    ' 見出し: 選択列
    ' 形: 数値の列
    ' 選んだ列（または選択範囲）の数値を、指定された桁数のゼロ埋めテキスト（例: 4桁なら 12 → "0012"）に変換する。
    ' 1万行でも一瞬で完了するよう、セルの表示形式をテキスト「@」にし、メモリ配列で一撃代入する。
    ' 中身は 選んだ列の数値を指定桁数でゼロ埋め文字列にする_本体（引数のある手伝い）。引数で目録とマクロ一覧から隠していたのをやめ、入口を引数なしにした（2026-10-29）
    Call 選んだ列の数値を指定桁数でゼロ埋め文字列にする_本体(0)
End Sub

Private Sub セル空白トリム_本体(ByVal モード As Long)
    ' モード 1: 文字の先頭の空白（半角・全角・タブ）を取る
    ' モード 2: 空白・空文字だけのセルを完全にクリアする（Emptyにする）
    ' モード 3: 前後の空白（Trim）を取る
    ' 変えたセルだけを書き戻す（範囲ごと書くと文字の数字・文字の日付が数・日付に化け、数式セルが潰れた・2026-10-22）
    Dim rng As Range, area As Range
    Dim dataArr As Variant, formArr As Variant, r As Long, c As Long
    Dim v As Variant, fStr As String, s As String, t As String, changed As Boolean
    Dim chg() As Boolean

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) = "Range" And Selection.CountLarge > 1 Then
        Set rng = Intersect(Selection, ActiveSheet.UsedRange)
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    For Each area In rng.Areas
        If area.Cells.CountLarge = 1 Then
            v = area.Value2
            If VarType(v) = vbString And Not area.HasFormula Then
                s = CStr(v)
                If モード = 1 Then
                    t = s
                    Do While Len(t) > 0
                        If Left$(t, 1) = " " Or Left$(t, 1) = "　" Or Left$(t, 1) = vbTab Then t = Mid$(t, 2) Else Exit Do
                    Loop
                    If t <> s Then
                        If (Left$(t, 1) = "0" And Len(t) >= 2) Or IsDate(t) Then
                            If area.NumberFormat <> "@" Then area.NumberFormat = "@"
                        End If
                        area.Cells(1, 1).Value = IIf(area.NumberFormat = "@", t, 文字のまま書く字(t))
                    End If
                ElseIf モード = 2 Then
                    t = Replace(Replace(Replace(s, " ", ""), "　", ""), vbTab, "")
                    If Len(t) = 0 Then area.ClearContents
                ElseIf モード = 3 Then
                    t = Trim$(Replace(s, "　", " "))
                    If t <> s Then
                        If (Left$(t, 1) = "0" And Len(t) >= 2) Or IsDate(t) Then
                            If area.NumberFormat <> "@" Then area.NumberFormat = "@"
                        End If
                        area.Cells(1, 1).Value = IIf(area.NumberFormat = "@", t, 文字のまま書く字(t))
                    End If
                End If
            End If
        Else
            dataArr = area.Value2
            formArr = area.Formula
            ReDim chg(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))
            changed = False
            For r = 1 To UBound(dataArr, 1)
                For c = 1 To UBound(dataArr, 2)
                    v = dataArr(r, c)
                    fStr = CStr(formArr(r, c))
                    ' 数式セル（先頭が "="）は絶対に破壊しない
                    If VarType(v) = vbString And Left$(fStr, 1) <> "=" Then
                        s = CStr(v)
                        If モード = 1 Then
                            t = s
                            Do While Len(t) > 0
                                If Left$(t, 1) = " " Or Left$(t, 1) = "　" Or Left$(t, 1) = vbTab Then t = Mid$(t, 2) Else Exit Do
                            Loop
                            If t <> s Then
                                dataArr(r, c) = t
                                chg(r, c) = True
                                changed = True
                            End If
                        ElseIf モード = 2 Then
                            t = Replace(Replace(Replace(s, " ", ""), "　", ""), vbTab, "")
                            If Len(t) = 0 Then
                                dataArr(r, c) = Empty
                                chg(r, c) = True
                                changed = True
                            End If
                        ElseIf モード = 3 Then
                            t = Trim$(Replace(s, "　", " "))
                            If t <> s Then
                                dataArr(r, c) = t
                                chg(r, c) = True
                                changed = True
                            End If
                        End If
                    End If
                Next c
            Next r
            If changed Then 変えたセルだけ書く area.Cells(1, 1), dataArr, chg
        End If
    Next area

    Application.ScreenUpdating = True
End Sub

Public Sub スピルエラーを検出し邪魔なセルをクリアする()
    ' 依頼の語: スピルエラーを直|SPILLエラーを直|スピルエラーを解消|SPILLエラーを解消|スピルの邪魔を消|SPILLの邪魔を消
    ' 依頼の組: スピル,SPILL,spill+エラー,邪魔,解消
    ' 階級: 基本
    Dim ws As Worksheet, cell As Range, spillRng As Range, blkCell As Range
    Dim formCells As Range, solvedCount As Long, clearedCount As Long

    Set ws = ActiveSheet
    solvedCount = 0: clearedCount = 0

    On Error Resume Next
    Set formCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formCells Is Nothing Then
        Application.StatusBar = "スピルエラーを検出し邪魔なセルをクリアする: シート内に数式セルが見つかりません。"   ' 窓を出さない＝先撃ちで止まらない（2026-10-21）
        Exit Sub
    End If

    For Each cell In formCells
        ' 日本語版の表示は「#スピル!」＝文字でなくエラーの種類（2045＝#SPILL!）で見る（2026-10-21）
        Dim スピルか As Boolean: スピルか = False
        If IsError(cell.Value) Then スピルか = (cell.Value = CVErr(2045))
        If スピルか Then
            ' #SPILL! の間は SpillingToRange が空（広がれていない）＝式を評価して広がる大きさを求める（2026-10-21 総点検・前は一度も見つけられなかった）
            Dim 答え As Variant, 行数 As Long, 列数 As Long
            Set spillRng = Nothing
            On Error Resume Next
            答え = ws.Evaluate(Mid$(cell.Formula2, 2))
            If IsArray(答え) Then
                行数 = UBound(答え, 1) - LBound(答え, 1) + 1
                列数 = 1
                列数 = UBound(答え, 2) - LBound(答え, 2) + 1
                Set spillRng = cell.Resize(行数, 列数)
            End If
            On Error GoTo 0

            If Not spillRng Is Nothing Then
                ' 先に邪魔なセルを全部見つけてから消す（1 つ消すと式が広がり、広がった先まで数えていた・2026-10-21）
                Dim 邪魔 As Range: Set 邪魔 = Nothing
                For Each blkCell In spillRng
                    If blkCell.Address <> cell.Address Then
                        If blkCell.MergeCells Or Not isEmpty(blkCell.Value) Or Len(blkCell.Formula) > 0 Then
                            If 邪魔 Is Nothing Then Set 邪魔 = blkCell Else Set 邪魔 = Union(邪魔, blkCell)
                        End If
                    End If
                Next blkCell
                If Not 邪魔 Is Nothing Then
                    clearedCount = clearedCount + 邪魔.Cells.Count
                    For Each blkCell In 邪魔.Cells
                        If blkCell.MergeCells Then blkCell.MergeArea.UnMerge
                    Next blkCell
                    邪魔.ClearContents
                End If
                cell.Calculate
                If Not IsError(cell.Value) Then solvedCount = solvedCount + 1
            End If
        End If
    Next cell

    If clearedCount > 0 Then
        Application.StatusBar = "スピルエラーを検出し邪魔なセルをクリアする: " & clearedCount & " 個の障害セルをクリアし、スピルエラーを解消しました。"
    Else
        Application.StatusBar = "スピルエラーを検出し邪魔なセルをクリアする: #SPILL! エラーまたはクリア可能な障害セルは見つかりませんでした。"
    End If
End Sub

Public Sub 英数を半角にカナを全角にする()
    ' 依頼の語: 全角英数を半角|全角の英数を半角|英数を半角|英数字を半角|全角を半角に|半角カナを全角|半角カタカナを全角|カナを全角|全角半角を整え|全角と半角を整え|全角半角をそろえ|全角と半角をそろえ|全角半角の混在|全角と半角が混ざ|文字種を統一|全角半角を整流
    ' 依頼の組: -フリガナ,ふりがな,住所,郵便,数字を全角,英数を全角
    ' 扱う: 全角 半角 カタカナ 英数
    ' 選んだ範囲の文字のセルで、全角の英数字を半角に、半角カナを全角にする。数式のセルは触らない。
    On Error GoTo Catch
    Dim 選択範囲 As Range
    Set 選択範囲 = Selection
    If 選択範囲 Is Nothing Then Exit Sub

    Dim データ As Variant
    データ = 選択範囲.Value
    If Not IsArray(データ) Then Exit Sub

    Dim 行総数 As Long, 列総数 As Long
    行総数 = UBound(データ, 1)
    列総数 = UBound(データ, 2)
    If 行総数 < 1 Or 列総数 < 1 Then Exit Sub

    Dim 行位置 As Long, 列位置 As Long
    Dim 文字列値 As String
    Dim 文字位置 As Long, 文字長 As Long
    Dim 判定文字 As String
    Dim コード値 As Long
    Dim 変換後文字列 As String
    Dim 変更件数 As Long
    Dim 変えた() As Boolean

    ReDim 変えた(1 To 行総数, 1 To 列総数)
    For 行位置 = 1 To 行総数
        For 列位置 = 1 To 列総数
            If VarType(データ(行位置, 列位置)) = vbString Then
                文字列値 = CStr(データ(行位置, 列位置))
                文字長 = Len(文字列値)
                If 文字長 > 0 Then
                    変換後文字列 = ""
                    ' 半角カナは続いている分をまとめて全角にする（1 字ずつだと「ｶﾞ」が「カ゛」に割れた・2026-10-21）
                    Dim カナの続き As String: カナの続き = ""
                    For 文字位置 = 1 To 文字長
                        判定文字 = Mid(文字列値, 文字位置, 1)
                        コード値 = AscW(判定文字) And &HFFFF&
                        If コード値 >= &HFF61& And コード値 <= &HFF9F& Then
                            カナの続き = カナの続き & 判定文字
                        Else
                            If Len(カナの続き) > 0 Then 変換後文字列 = 変換後文字列 & StrConv(カナの続き, vbWide): カナの続き = ""
                            If (コード値 >= &HFF10& And コード値 <= &HFF19&) Or _
                               (コード値 >= &HFF21& And コード値 <= &HFF3A&) Or _
                               (コード値 >= &HFF41& And コード値 <= &HFF5A&) Then
                                変換後文字列 = 変換後文字列 & StrConv(判定文字, vbNarrow)
                            Else
                                変換後文字列 = 変換後文字列 & 判定文字
                            End If
                        End If
                    Next 文字位置
                    If Len(カナの続き) > 0 Then 変換後文字列 = 変換後文字列 & StrConv(カナの続き, vbWide)
                    If 変換後文字列 <> 文字列値 Then
                        If Not 選択範囲.Cells(行位置, 列位置).HasFormula Then
                            データ(行位置, 列位置) = 変換後文字列
                            変えた(行位置, 列位置) = True
                            変更件数 = 変更件数 + 1
                        End If
                    End If
                End If
            End If
        Next 列位置
    Next 行位置

    ' 変えたセルだけ書く（文字の数字 "00123" が数に化けるのを防ぐ・2026-10-22）
    If 変更件数 > 0 Then
        Call 表の整理_下請け.変えたセルだけ書く(選択範囲.Cells(1, 1), データ, 変えた, False)
    End If

    Application.StatusBar = "英数を半角にカナを全角にする: " & 変更件数 & " 個のセルを直しました。"
    Exit Sub
Catch:
    Application.StatusBar = "英数を半角にカナを全角にする: 途中で止まりました（" & Err.Description & "）。"
End Sub


Public Sub セルの見えない空白と改行を取る()
    ' 依頼の語: 前後の見えない|前後の空白|前後のスペース|前後の改行|文字の前後|余分な空白|余計な空白|トリム|TRIM|ノーブレーク|空白を取|スペースを消|スペースを取
    ' 依頼の組: 空白,スペース,改行+前後,余分,余計,紛れ+消,取,削,除去,落と+-空白の行,空白行,空行,先頭,だけのセル,空のセル,空に,そろえ,揃え,列
    ' 扱う: 空白 スペース 改行 トリム
    ' 選んだ範囲の文字のセルの前後にある空白・全角空白・タブ・改行・ノーブレークスペースを取る。数式のセルは触らない。
    On Error GoTo Catch
    Dim 選択範囲 As Range
    Set 選択範囲 = Selection
    If 選択範囲 Is Nothing Then Exit Sub

    Dim データ As Variant
    データ = 選択範囲.Value
    If Not IsArray(データ) Then Exit Sub

    Dim 行総数 As Long, 列総数 As Long
    行総数 = UBound(データ, 1)
    列総数 = UBound(データ, 2)
    If 行総数 < 1 Or 列総数 < 1 Then Exit Sub

    Dim 行位置 As Long, 列位置 As Long
    Dim 文字列値 As String, 新文字 As String
    Dim 変更件数 As Long
    Dim 変えた() As Boolean
    Dim 文字長 As Long, p1 As Long, p2 As Long, cCode As Long

    ReDim 変えた(1 To 行総数, 1 To 列総数)

    For 行位置 = 1 To 行総数
        For 列位置 = 1 To 列総数
            If VarType(データ(行位置, 列位置)) = vbString Then
                文字列値 = CStr(データ(行位置, 列位置))
                文字長 = Len(文字列値)
                If 文字長 > 0 Then
                    ' 前後の見えない空白・改行（半角・全角空白、タブ、NBSP、vbCr、vbLf、ゼロ幅空白）だけを取る（途中は触らない・2026-10-22 組長決定）
                    p1 = 1
                    Do While p1 <= 文字長
                        cCode = AscW(Mid$(文字列値, p1, 1)) And &HFFFF&
                        If cCode = 32 Or cCode = 12288 Or cCode = 9 Or cCode = 160 Or cCode = 10 Or cCode = 13 Or cCode = 8203 Then
                            p1 = p1 + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    p2 = 文字長
                    Do While p2 >= p1
                        cCode = AscW(Mid$(文字列値, p2, 1)) And &HFFFF&
                        If cCode = 32 Or cCode = 12288 Or cCode = 9 Or cCode = 160 Or cCode = 10 Or cCode = 13 Or cCode = 8203 Then
                            p2 = p2 - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    If p1 > p2 Then
                        新文字 = ""
                    ElseIf p1 > 1 Or p2 < 文字長 Then
                        新文字 = Mid$(文字列値, p1, p2 - p1 + 1)
                    Else
                        新文字 = 文字列値
                    End If

                    If 新文字 <> 文字列値 Then
                        If Not 選択範囲.Cells(行位置, 列位置).HasFormula Then
                            データ(行位置, 列位置) = 新文字
                            変えた(行位置, 列位置) = True
                            変更件数 = 変更件数 + 1
                        End If
                    End If
                End If
            End If
        Next 列位置
    Next 行位置

    ' 変えたセルだけ書く（文字の数字 "00123" や数式セルを壊さない・2026-10-22）
    If 変更件数 > 0 Then
        Call 表の整理_下請け.変えたセルだけ書く(選択範囲.Cells(1, 1), データ, 変えた, False)
        Application.StatusBar = "セルの見えない空白と改行を取る: " & 変更件数 & " 個のセルの前後の空白・改行を取りました。"
    Else
        Application.StatusBar = "セルの見えない空白と改行を取る: 前後の空白・改行は見つかりませんでした。"
    End If
    Exit Sub
Catch:
    Application.StatusBar = "セルの見えない空白と改行を取る: 途中で止まりました（" & Err.Description & "）。"
End Sub

Public Sub ブックの式を全部値にする()
    ' 依頼の語: 全シートの計算式を値|全シートの数式を値|全シートの式を値|全部のシートの式を値|すべてのシートの数式を値|すべてのシートの式を値|ブック全体の数式を値|ブック全体の計算式を値|ブック全体の式を値|ブックの式を全部値|ブックの数式を全部値|ブック内の数式を値|ブック内の計算式を値
    ' 依頼の組: 全シート,全部のシート,すべてのシート,ブック全体,ブック内+式,数式,計算式+値+-列,選んだ,外部リンク,調べ,一覧
    ' 扱う: 計算式 数式 値 全シート ブック全体 控え
    ' ブックの全シートの式を、今の値に置き換える。消す前にブックの控えを1部作る（未保存のブックは控えが作れないので止める）。
    Dim wb As Workbook, 対象シート As Worksheet, 式域 As Range
    Dim 控え As String, 処理シート数 As Long, 式の数 As Long
    Dim 一区 As Range, 一つ As Range, 文字の式 As Collection, 組 As Variant
    Dim スピル As Collection, 広がり As Range
    Dim 前の計算 As XlCalculation

    Set wb = ActiveWorkbook
    If wb.path = "" Then
        Application.StatusBar = "ブックの式を全部値にする: 未保存のブックなので止めました（控えが作れないため。一度保存してから撃ってください）。"
        Exit Sub
    End If
    控え = wb.path & "\" & Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & _
           "_値にする前_" & Format$(Now, "yyyymmdd_hhnnss") & Mid$(wb.Name, InStrRev(wb.Name, "."))
    On Error Resume Next
    wb.SaveCopyAs 控え
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Application.StatusBar = "ブックの式を全部値にする: 控えが作れなかったので止めました（" & 控え & "）。"
        Exit Sub
    End If
    On Error GoTo Catch

    Application.ScreenUpdating = False
    前の計算 = Application.Calculation           ' 終わりに前の値へ戻す（手動計算にしていた人を自動に変えない・2026-10-28 点検）
    Application.Calculation = xlCalculationManual
    For Each 対象シート In wb.Worksheets
        Set 式域 = Nothing
        On Error Resume Next
        Set 式域 = 対象シート.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Catch
        If Not 式域 Is Nothing Then
            式の数 = 式の数 + 式域.Count
            ' 式のセルだけを値にする（使用範囲を丸ごと書き戻すと、式でない 00123 や文字の日付まで数・日付に化けた・2026-10-22）。
            ' 式の結果が文字のセルは、数や日付に化けないよう文字のまま書き直す
            Set 文字の式 = New Collection
            Set スピル = New Collection
            For Each 一つ In 式域.Cells
                If VarType(一つ.Value) = vbString Then 文字の式.Add Array(一つ, CStr(一つ.Value))
                ' 動的配列の式（スピル）は広がった範囲と値を先に控える。式のセルだけ値にすると、広がった側の値が消えていた（2026-10-28 点検）
                Set 広がり = 表の整理_下請け.スピルの元なら範囲(一つ)
                If Not 広がり Is Nothing Then スピル.Add Array(広がり, 広がり.Value)
            Next 一つ
            For Each 一区 In 式域.Areas
                一区.Value = 一区.Value ' 書き戻し: 新しい所（式のセルだけ）
            Next 一区
            For Each 組 In 文字の式
                組(0).Value = "'" & 組(1)
            Next 組
            For Each 組 In スピル
                表の整理_下請け.値を書き戻す 組(0), 組(1)
            Next 組
            処理シート数 = 処理シート数 + 1
        End If
    Next 対象シート
    Application.Calculation = 前の計算
    Application.ScreenUpdating = True
    Application.StatusBar = "ブックの式を全部値にする: " & 処理シート数 & " 枚のシートの式 " & 式の数 & " か所を値にしました。控え: " & 控え
    Exit Sub
Catch:
    If 前の計算 = 0 Then 前の計算 = xlCalculationAutomatic
    Application.Calculation = 前の計算
    Application.ScreenUpdating = True
    Application.StatusBar = "ブックの式を全部値にする: 途中で止まりました（" & Err.Description & "）。控え: " & 控え
End Sub


Public Sub 同じ形のシートを縦につなぐ()
    ' 依頼の語: 縦につな|縦に結合|縦に連結|縦に積|シートを縦に|全シートを縦|シートを連結|シートの表を連結|同じ形のシートをつな|同じレイアウトのシート|同じ形の表をつな
    ' 依頼の組: シート,表+縦,連結,つなぎ合わ,つなげ+-横,分け,分割,突合,照合,グラフ,ピボット,縦持ち,縦棒,集約
    ' 扱う: シート 縦 連結 つなぐ
    ' 見えている全シートの表（使っている範囲）を、先頭のシートの見出しを 1 回だけ使って新しいシートに縦につなぐ。元のシートは触らない。
    On Error GoTo Catch
    If ActiveWorkbook.Worksheets.Count < 2 Then
        Application.StatusBar = "同じ形のシートを縦につなぐ: シートが 2 枚以上無いのでつなげません。"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim 統合シート As Worksheet
    Set 統合シート = Worksheets.Add(Before:=Worksheets(1))
    統合シート.Name = "全表統合_" & Format(Now, "hhnnss")

    Dim 対象シート As Worksheet
    Dim 見出しコピー済 As Boolean
    見出しコピー済 = False
    Dim 書き込み行 As Long
    書き込み行 = 1

    Dim シートデータ As Variant
    Dim データ行数 As Long, データ列数 As Long
    Dim コピー行数 As Long
    Dim 行オフセット As Long, 列オフセット As Long
    Dim 貼付範囲 As Range
    Dim 部分データ() As Variant

    For Each 対象シート In ActiveWorkbook.Worksheets
        If 対象シート.Name <> 統合シート.Name And 対象シート.Visible = xlSheetVisible Then
            シートデータ = 対象シート.UsedRange.Value
            If IsArray(シートデータ) Then
                データ行数 = UBound(シートデータ, 1)
                データ列数 = UBound(シートデータ, 2)
                If Not 見出しコピー済 Then
                    ' 文字の 00123・1-2・16 桁の番号が数・日付に化けないよう、文字は文字のまま書く（2026-10-28 点検）
                    統合シート.Range(統合シート.Cells(1, 1), 統合シート.Cells(データ行数, データ列数)).Value = 表の整理_下請け.文字を守った配列(シートデータ) ' 書き戻し: 新しい所
                    書き込み行 = データ行数 + 1
                    見出しコピー済 = True
                ElseIf データ行数 >= 2 Then
                    コピー行数 = データ行数 - 1
                    ReDim 部分データ(1 To コピー行数, 1 To データ列数)
                    For 行オフセット = 1 To コピー行数
                        For 列オフセット = 1 To データ列数
                            部分データ(行オフセット, 列オフセット) = シートデータ(行オフセット + 1, 列オフセット)
                        Next 列オフセット
                    Next 行オフセット
                    Set 貼付範囲 = 統合シート.Range(統合シート.Cells(書き込み行, 1), 統合シート.Cells(書き込み行 + コピー行数 - 1, データ列数))
                    貼付範囲.Value = 表の整理_下請け.文字を守った配列(部分データ) ' 書き戻し: 新しい所
                    書き込み行 = 書き込み行 + コピー行数
                End If
            End If
        End If
    Next 対象シート

    Application.ScreenUpdating = True
    Application.StatusBar = "同じ形のシートを縦につなぐ: シート「" & 統合シート.Name & "」に " & (書き込み行 - 1) & " 行（見出しを含む）をつなぎました。"
    Exit Sub
Catch:
    Application.ScreenUpdating = True
    Application.StatusBar = "同じ形のシートを縦につなぐ: 途中で止まりました（" & Err.Description & "）。"
End Sub

Public Sub 二つの表の片方だけの行を抜く()
    ' 依頼の語: 片方にしか|片方だけ|片方にだけ|一方にしか|一方だけにある|差分の行を抜|差分行を抜|差分を抜き出|差分の行だけ|無い行を抜|ない行を抜|載っていない行
    ' 依頼の組: 片方,一方,差分,にない,に無い,載っていない+抜き,抜い,抜く,抽出,取り出+-色,ハイライト,塗り,着色,差異を書,差額
    ' 扱う: 差分 片方 抜き出す
    ' 表A と表B の範囲を聞き、表A にあって表B に無い行（全部の列が同じ行が無いもの）を新しいシートに抜き出す。元の表は触らない。
    On Error GoTo Catch
    Dim 表A範囲 As Range, 表B範囲 As Range
    On Error Resume Next
    Set 表A範囲 = Application.InputBox("基準となる表（表A）のデータ範囲を選択してください", "表A選択", Type:=8)
    If 表A範囲 Is Nothing Then Exit Sub
    Set 表B範囲 = Application.InputBox("比較対象の表（表B）のデータ範囲を選択してください", "表B選択", Type:=8)
    If 表B範囲 Is Nothing Then Exit Sub
    On Error GoTo Catch

    Dim データA As Variant, データB As Variant
    データA = 表A範囲.Value
    データB = 表B範囲.Value
    If Not IsArray(データA) Or Not IsArray(データB) Then Exit Sub

    Dim 行数A As Long, 列数A As Long
    行数A = UBound(データA, 1): 列数A = UBound(データA, 2)
    Dim 行数B As Long, 列数B As Long
    行数B = UBound(データB, 1): 列数B = UBound(データB, 2)

    Dim 辞書B As Object
    Set 辞書B = CreateObject("Scripting.Dictionary")

    Dim 行番号 As Long, 列番号 As Long
    Dim キー文字列 As String

    For 行番号 = 1 To 行数B
        キー文字列 = ""
        For 列番号 = 1 To 列数B
            キー文字列 = キー文字列 & CStr(データB(行番号, 列番号)) & vbTab
        Next 列番号
        If Not 辞書B.exists(キー文字列) Then
            辞書B.Add キー文字列, True
        End If
    Next 行番号

    Dim 差分データ() As Variant
    ReDim 差分データ(1 To 行数A, 1 To 列数A)
    Dim 差分行数 As Long
    差分行数 = 0

    For 行番号 = 1 To 行数A
        キー文字列 = ""
        For 列番号 = 1 To 列数A
            キー文字列 = キー文字列 & CStr(データA(行番号, 列番号)) & vbTab
        Next 列番号
        If Not 辞書B.exists(キー文字列) Then
            差分行数 = 差分行数 + 1
            For 列番号 = 1 To 列数A
                差分データ(差分行数, 列番号) = データA(行番号, 列番号)
            Next 列番号
        End If
    Next 行番号

    If 差分行数 = 0 Then
        Application.StatusBar = "二つの表の片方だけの行を抜く: 表A の行はすべて表B にあります（差分なし）。"
        Exit Sub
    End If

    Dim 出力シート As Worksheet
    Set 出力シート = Worksheets.Add(, ActiveSheet)
    出力シート.Name = "表Aのみ差分_" & Format(Now, "hhnnss")

    表の整理_下請け.文字を守って書く 出力シート.Range(出力シート.Cells(1, 1), 出力シート.Cells(差分行数, 列数A)), 差分データ ' 書き戻し: 新しい所
    Application.StatusBar = "二つの表の片方だけの行を抜く: 表A にしか無い行 " & 差分行数 & " 件をシート「" & 出力シート.Name & "」に出しました。"
    Exit Sub
Catch:
    Application.StatusBar = "二つの表の片方だけの行を抜く: 途中で止まりました（" & Err.Description & "）。"
End Sub

Sub ふりがなの空欄を埋める()
    ' 見出しが ふりがな・フリガナ・よみ・カナ の列の空欄だけを、同じ行の氏名（見出しが 氏名・名前・名称 の列）の読みで埋める（Application.GetPhonetic）。見出しやすでに入っている字のひらがな／カタカナに合わせる。埋まっている所と氏名が空の行は触らない（2026-10-22）
    ' 依頼の語: ふりがなの空欄|ふりがなが空欄|ふりがなは空欄|ふりがなの空白|ふりがなが空白|ふりがなを埋|フリガナの空欄|フリガナが空欄|フリガナの空白|フリガナを埋|よみがなの空欄|読みがなの空欄|よみがなを埋|読みがなを埋|ふりがなを修正|フリガナを修正
    ' 依頼の組: ふりがな,フリガナ,ﾌﾘｶﾞﾅ,よみがな,ヨミガナ,読みがな+空欄,空白,空き,埋め,埋ま,未入力,抜け,入っていない,修正+-消,削除,並べ,並び,ソート,ひらがなに,カタカナに,探,調べ,一覧,チェック,数え,件数,色,赤,塗,強調,目立,抽出,絞り込
    ' 扱う: ふりがな フリガナ 読み 空欄 氏名
    ' 見出し: なし
    ' 変えてよい: 空欄

    Dim ws As Worksheet, 表 As Range, why As String
    Dim hdr As Long, c1 As Long, c2 As Long, 最終 As Long, c As Long, k As Long
    Dim 読み列() As Long, 名前列() As Long, 読み数 As Long, 名数 As Long
    Dim 一覧 As String, 名列 As Long, 差 As Long, 最小 As Long, 型 As String
    Dim 名前々 As Variant, 読々 As Variant, 変() As Boolean, r As Long, n As Long, 行数 As Long
    Dim 辞書 As Object, 名 As String, 読 As String, 埋めた As Long, 読めない As Long, 空き As Long
    Dim 対 As String, 状況 As String
    Dim oldScreen As Boolean, oldEvents As Boolean

    Set ws = ActiveSheet
    Set 表 = 明細の表の範囲(ws, why)
    If 表 Is Nothing Then Call 先撃ちに断る(why): Exit Sub
    hdr = 表.row
    c1 = 表.Column
    c2 = 表.Column + 表.Columns.Count - 1
    最終 = 表.row + 表.rows.Count - 1
    If 最終 <= hdr Then Call 先撃ちに断る("見出しの下に明細の行がありません。"): Exit Sub
    行数 = 最終 - hdr

    ReDim 読み列(1 To 表.Columns.Count)
    ReDim 名前列(1 To 表.Columns.Count)
    For c = c1 To c2
        If ふりがなの見出しか(見出しの字(ws.Cells(hdr, c))) Then
            読み数 = 読み数 + 1: 読み列(読み数) = c
        ElseIf 氏名の見出しか(見出しの字(ws.Cells(hdr, c))) Then
            名数 = 名数 + 1: 名前列(名数) = c
        End If
        一覧 = 一覧 & IIf(Len(一覧) > 0, "・", "") & 見出しの字(ws.Cells(hdr, c))
    Next c
    If 読み数 = 0 Then
        Call 先撃ちに断る("見出しが ふりがな・フリガナ・よみ・カナ の列がありません（見出し: " & 一覧 & "）。ふりがなを入れる列の見出しを付けてください。")
        Exit Sub
    End If
    If 名数 = 0 Then
        Call 先撃ちに断る("読みのもとになる氏名の列（見出しが 氏名・名前・名称）がありません（見出し: " & 一覧 & "）。")
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo 後始末
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set 辞書 = CreateObject("Scripting.Dictionary")

    For k = 1 To 読み数
        ' この読みの列に対応する氏名の列: いちばん近い氏名の列（同じ近さなら左）
        名列 = 0: 最小 = 100000
        For c = 1 To 名数
            差 = Abs(名前列(c) - 読み列(k))
            If 名前列(c) > 読み列(k) Then 差 = 差 * 2 - 1
            If 差 < 最小 Then 最小 = 差: 名列 = 名前列(c)
        Next c
        型 = ふりがなの型(ws, 読み列(k), hdr + 1, 最終, 見出しの字(ws.Cells(hdr, 読み列(k))))
        名前々 = 二次元にする(ws.Range(ws.Cells(hdr + 1, 名列), ws.Cells(最終, 名列)).Value2)
        読々 = 二次元にする(ws.Range(ws.Cells(hdr + 1, 読み列(k)), ws.Cells(最終, 読み列(k))).Value2)
        ReDim 変(1 To 行数, 1 To 1)
        For r = 1 To 行数
            If VarType(名前々(r, 1)) = vbString Then
                名 = Trim$(CStr(名前々(r, 1)))
            Else
                名 = ""
            End If
            If Len(名) > 0 Then
                ' 空欄: 空のセル・空白だけの字（式のセルは触らない）
                If isEmpty(読々(r, 1)) Then
                    空き = 1
                ElseIf VarType(読々(r, 1)) = vbString Then
                    If Trim$(Replace(CStr(読々(r, 1)), ChrW(&H3000), " ")) = "" And Not ws.Cells(hdr + r, 読み列(k)).HasFormula Then 空き = 1 Else 空き = 0
                Else
                    空き = 0
                End If
                If 空き = 1 Then
                    If 埋めた Mod 200 = 199 Then Application.StatusBar = "ふりがなを調べています… " & r & " / " & 行数 & " 行"
                    If 辞書.exists(名 & "|" & 型) Then
                        読 = 辞書(名 & "|" & 型)
                    Else
                        読 = 読みの字(名, 型)
                        辞書(名 & "|" & 型) = 読
                    End If
                    If Len(読) > 0 Then
                        読々(r, 1) = 読
                        変(r, 1) = True
                        埋めた = 埋めた + 1
                    Else
                        読めない = 読めない + 1
                    End If
                End If
            End If
        Next r
        If 埋めた > 0 Then Call 表の整理_下請け.変えたセルだけ書く(ws.Cells(hdr + 1, 読み列(k)), 読々, 変, False, "")
        対 = 対 & IIf(Len(対) > 0, "・", "") & Split(ws.Cells(1, 読み列(k)).Address(True, False), "$")(0) & "列←" & Split(ws.Cells(1, 名列).Address(True, False), "$")(0) & "列"
    Next k

    If 埋めた = 0 Then
        状況 = "ふりがなの空欄はありませんでした"
        If 読めない > 0 Then 状況 = "読みを取れた氏名がありませんでした（" & 読めない & " 件）"
        Call 先撃ちに断る(状況 & "（" & 対 & "）。")
    Else
        状況 = "ふりがなの空欄 " & 埋めた & " 件を氏名の読みで埋めました（" & 対 & "）。IME の読みなので人名は確かめてください"
        If 読めない > 0 Then 状況 = 状況 & "。読みを取れなかった " & 読めない & " 件は空欄のままです"
        Application.StatusBar = 状況
    End If

後始末:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

' 見出しがふりがなの列か（ふりがな・フリガナ・よみ・カナ）。短い語（よみ・カナ）は見出しの頭か終わりにあるときだけ（「カナダ」を除く）
Function ふりがなの見出しか(ByVal 見 As String) As Boolean
    Dim s As String, w As Variant
    s = ピボットの語ならし(見)
    If Len(s) = 0 Then Exit Function
    For Each w In Array("ふりがな", "フリガナ", "ﾌﾘｶﾞﾅ", "よみがな", "ヨミガナ", "読みがな", "読み仮名", "ふりかな", "フリカナ")
        If InStr(s, CStr(w)) > 0 Then ふりがなの見出しか = True: Exit Function
    Next w
    For Each w In Array("よみ", "ヨミ", "読み", "カナ", "ｶﾅ")
        If s = CStr(w) Or Left$(s, Len(w)) = CStr(w) Or Right$(s, Len(w)) = CStr(w) Then ふりがなの見出しか = True: Exit Function
    Next w
End Function

' 見出しが氏名の列か（氏名・名前・名称。ふりがなの列は除く）
Function 氏名の見出しか(ByVal 見 As String) As Boolean
    Dim s As String, w As Variant
    If ふりがなの見出しか(見) Then Exit Function
    s = ピボットの語ならし(見)
    For Each w In Array("氏名", "名前", "名称")
        If InStr(s, CStr(w)) > 0 Then 氏名の見出しか = True: Exit Function
    Next w
End Function

' ふりがなの列を埋める字の型。すでに入っている最初の字、無ければ見出しの字のひらがな／カタカナに合わせる。どちらも決まらなければカタカナ
Function ふりがなの型(ByVal ws As Worksheet, ByVal 列 As Long, ByVal 開始行 As Long, ByVal 終行 As Long, ByVal 見 As String) As String
    Dim 値 As Variant, r As Long, 型 As String
    値 = 二次元にする(ws.Range(ws.Cells(開始行, 列), ws.Cells(終行, 列)).Value2)
    For r = 1 To UBound(値, 1)
        If VarType(値(r, 1)) = vbString Then
            型 = かなの型(CStr(値(r, 1)))
            If Len(型) > 0 Then ふりがなの型 = 型: Exit Function
        End If
    Next r
    型 = かなの型(ピボットの語ならし(見))
    If Len(型) > 0 Then ふりがなの型 = 型 Else ふりがなの型 = "カナ"
End Function

' 字の最初のかなが ひらがな なら "ひら"、カタカナ（半角も）なら "カナ"、かなが無ければ ""
Function かなの型(ByVal s As String) As String
    Dim i As Long, w As Long
    For i = 1 To Len(s)
        w = AscW(Mid$(s, i, 1))
        If w < 0 Then w = w + 65536
        If w >= &H3041 And w <= &H3096 Then かなの型 = "ひら": Exit Function
        If (w >= &H30A1 And w <= &H30FA) Or (w >= &HFF66& And w <= &HFF9F&) Then かなの型 = "カナ": Exit Function
    Next i
End Function

' 氏名の読み（IME）。型に合わせたかなの字にする。取れなければ ""
Function 読みの字(ByVal 名 As String, ByVal 型 As String) As String
    Dim s As String
    On Error Resume Next
    s = Application.GetPhonetic(名)
    If Err.Number <> 0 Then s = "": Err.Clear
    On Error GoTo 0
    If Len(s) = 0 Then Exit Function
    If Len(かなの型(s)) = 0 Then Exit Function         ' かなが 1 字も無い（英字の氏名など）は読みにしない
    ' 氏名の区切りが半角の空白なら、読みの区切りも半角に合わせる
    If InStr(名, " ") > 0 And InStr(名, ChrW(&H3000)) = 0 Then s = Replace(s, ChrW(&H3000), " ")
    読みの字 = カナの字をそろえる(s, 型)
End Function

Sub 選んだ列のカナをひらがなにする()
    ' 選んだ列（頼みに見出しの語があればその列・選ばず頼みも無ければ表全体）の文字の、全角カタカナと半角カナをひらがなにする。カナ以外の字（漢字・英数・記号）と数式のセルは触らない（2026-10-22）
    ' 依頼の語: カタカナをひらがな|カナをひらがな|カタカナからひらがな|カナからひらがな|カタカナを平仮名|片仮名をひらがな|片仮名を平仮名|ｶﾅをひらがな
    ' 依頼の組: カタカナ,カナ,片仮名,ｶﾅ+ひらがなに,ひらがなへ,平仮名に,ひらがな化,ひらがなで
    ' 扱う: カタカナ ひらがな 半角カナ 変換
    ' 見出し: なし
    ' 列は頼みから: 列

    Dim ws As Worksheet, 表 As Range, why As String, 頼 As String
    Dim 列々() As Long, 上々() As Long, 下々() As Long, 本数 As Long
    Dim hdr As Long, c1 As Long, c2 As Long, 最終 As Long, 最大行 As Long
    Dim c As Long, k As Long, r As Long, 見 As String, 領 As Range, 全 As Variant
    Dim 値 As Variant, 変() As Boolean, 行数 As Long, 変えた As Long, 変えた列 As Long, 列で As Boolean
    Dim s As String, 新 As String, 名々 As String, 頼の語 As Variant, 取れた As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, 選単 As Boolean

    Set ws = ActiveSheet
    最大行 = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    頼 = 先撃ちの頼みの文()
    選単 = True
    If TypeName(Selection) = "Range" Then 選単 = (Selection.CountLarge = 1)
    Set 表 = 明細の表の範囲(ws, why)
    ReDim 列々(1 To 1): ReDim 上々(1 To 1): ReDim 下々(1 To 1)

    If Len(頼) > 0 Or 選単 Then
        ' 頼みの文に出てくる見出しの列。無ければ表全体。選んでいる 1 セルは「選んでいない」と見る
        If 表 Is Nothing Then Call 先撃ちに断る(why): Exit Sub
        hdr = 表.row
        c1 = 表.Column
        c2 = 表.Column + 表.Columns.Count - 1
        最終 = 表.row + 表.rows.Count - 1
        If Len(頼) > 0 Then
            For Each 頼の語 In Array("半角カタカナ", "全角カタカナ", "半角カナ", "全角カナ", "カタカナ", "片仮名", "ひらがな", "平仮名", "半角", "全角", "変換")
                頼 = Replace(頼, CStr(頼の語), "")
            Next 頼の語
            For c = c1 To c2
                見 = 見出しの字(ws.Cells(hdr, c))
                If Len(見) >= 2 Then
                    If 見出しの語の位置(頼, 見) > 0 Then
                        取れた = True
                        本数 = 本数 + 1
                        ReDim Preserve 列々(1 To 本数): ReDim Preserve 上々(1 To 本数): ReDim Preserve 下々(1 To 本数)
                        列々(本数) = c: 上々(本数) = hdr + 1: 下々(本数) = 最終
                    End If
                End If
            Next c
        End If
        If Not 取れた Then
            For c = c1 To c2
                本数 = 本数 + 1
                ReDim Preserve 列々(1 To 本数): ReDim Preserve 上々(1 To 本数): ReDim Preserve 下々(1 To 本数)
                列々(本数) = c: 上々(本数) = hdr + 1: 下々(本数) = 最終
            Next c
        End If
    Else
        ' 選んでいる範囲（先頭の行は見出し）
        最終 = 表の本文の最終行(ws)
        For Each 領 In Selection.Areas
            If 領.rows.Count > 100000 And Not 表 Is Nothing Then
                hdr = 表.row
            Else
                hdr = 領.row
            End If
            If 領.rows.Count > 1 And 領.rows.Count <= 100000 Then
                k = 領.row + 領.rows.Count - 1
                If k > 最大行 Then k = 最大行
            Else
                k = 最終
            End If
            If k > hdr Then
                For c = 領.Column To 領.Column + 領.Columns.Count - 1
                    If c <= ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 Then
                        本数 = 本数 + 1
                        ReDim Preserve 列々(1 To 本数): ReDim Preserve 上々(1 To 本数): ReDim Preserve 下々(1 To 本数)
                        列々(本数) = c: 上々(本数) = hdr + 1: 下々(本数) = k
                    End If
                Next c
            End If
        Next 領
    End If
    If 本数 = 0 Then Call 先撃ちに断る("ひらがなにする列が決まりません。列を選ぶか、頼みに見出しの語を書いてください。"): Exit Sub

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo 後始末
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For k = 1 To 本数
        行数 = 下々(k) - 上々(k) + 1
        If 行数 >= 1 Then
            値 = 二次元にする(ws.Range(ws.Cells(上々(k), 列々(k)), ws.Cells(下々(k), 列々(k))).Value2)
            ReDim 変(1 To 行数, 1 To 1)
            列で = False
            For r = 1 To 行数
                If VarType(値(r, 1)) = vbString Then
                    s = CStr(値(r, 1))
                    If カナを含むか(s) Then
                        ' 数式のセルは触らない
                        If Not ws.Cells(上々(k) + r - 1, 列々(k)).HasFormula Then
                            新 = カナの字をそろえる(s, "ひら")
                            If 新 <> s Then
                                値(r, 1) = 新
                                変(r, 1) = True
                                変えた = 変えた + 1
                                列で = True
                            End If
                        End If
                    End If
                End If
            Next r
            If 列で Then
                Call 表の整理_下請け.変えたセルだけ書く(ws.Cells(上々(k), 列々(k)), 値, 変, False, "")
                変えた列 = 変えた列 + 1
                If 変えた列 <= 4 Then 名々 = 名々 & IIf(Len(名々) > 0, "・", "") & Split(ws.Cells(1, 列々(k)).Address(True, False), "$")(0) & "列"
            End If
        End If
    Next k
    If 変えた = 0 Then
        Call 先撃ちに断る("ひらがなにするカナはありませんでした。")
    Else
        Application.StatusBar = "カナをひらがなにしました: " & 変えた & " セル（" & 名々 & IIf(変えた列 > 4, " ほか", "") & "）。漢字・英数・数式は触っていません"
    End If

後始末:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub

Sub 選んだ列の値ごとにシートを分割する_本体(ByVal 知らせる As Boolean)
    ' 選んだ列の値ごとにシートを分割する の中身。知らせる＝終わりに窓で知らせるか（選んだ列の値ごとにシートを分ける からは False）
    Dim wsSrc As Worksheet, wb As Workbook
    Dim ur As Range, headerRow As Long, lastRow As Long, firstCol As Long, lastCol As Long, nCols As Long
    Dim targetCol As Long, relCol As Long, r As Long, c As Long
    Dim key As Variant, valStr As String, cellStr As String, sheetTitle As String
    Dim dict As Object, colList As Collection, 今回 As Object
    Dim dataArr As Variant, headerArr As Variant
    Dim colFmts() As String, hdrFmts() As String
    Dim wsNew As Worksheet, matchCount As Long
    Dim outArr As Variant, outR As Long, srcR As Long
    Dim isBlankRow As Boolean, isSumRow As Boolean

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set wsSrc = ActiveSheet
    Set wb = wsSrc.Parent

    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub
    firstCol = ur.Column
    lastCol = ur.Column + ur.Columns.Count - 1
    lastRow = 表の本文の最終行(wsSrc)

    ' 分割キー列の特定
    If TypeName(Selection) = "Range" Then
        targetCol = Selection.Cells(1, 1).Column
    Else
        targetCol = firstCol
    End If
    If targetCol < firstCol Or targetCol > lastCol Then targetCol = firstCol

    ' 見出し行の特定（途中のデータ行を選んでいても表の上端から見出し行を正しく特定する）
    headerRow = 0
    For r = ur.row To lastRow
        If Len(セルの字(wsSrc.Cells(r, targetCol).Value)) > 0 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then
        For r = ur.row To lastRow
            For c = firstCol To lastCol
                If Len(セルの字(wsSrc.Cells(r, c).Value)) > 0 Then
                    headerRow = r
                    Exit For
                End If
            Next c
            If headerRow > 0 Then Exit For
        Next r
    End If
    If headerRow = 0 Or headerRow >= lastRow Then Exit Sub

    ' 表の左右端を見出し行の非空セル範囲に合わせる（A1が表題で表がC3からのケースに対応）
    Do While firstCol < lastCol And Trim$(セルの字(wsSrc.Cells(headerRow, firstCol).Value)) = ""
        firstCol = firstCol + 1
    Loop
    Dim scanLastCol As Long: scanLastCol = firstCol
    For c = firstCol To lastCol
        If Trim$(セルの字(wsSrc.Cells(headerRow, c).Value)) = "" Then Exit For
        scanLastCol = c
    Next c
    lastCol = scanLastCol
    nCols = lastCol - firstCol + 1
    If nCols <= 0 Then Exit Sub

    If targetCol < firstCol Or targetCol > lastCol Then targetCol = firstCol
    relCol = targetCol - firstCol + 1

    ' 見出しと1行目データの表示形式（NumberFormat）をキャッシュ
    ReDim hdrFmts(1 To nCols)
    ReDim colFmts(1 To nCols)
    For c = 1 To nCols
        hdrFmts(c) = CStr(wsSrc.Cells(headerRow, firstCol + c - 1).NumberFormat)
        colFmts(c) = CStr(wsSrc.Cells(headerRow + 1, firstCol + c - 1).NumberFormat)
    Next c

    ' ヘッダーとデータの一括取得（最速化）
    If nCols = 1 Then
        ReDim headerArr(1 To 1, 1 To 1)
        headerArr(1, 1) = wsSrc.Cells(headerRow, firstCol).Value
    Else
        headerArr = wsSrc.Range(wsSrc.Cells(headerRow, firstCol), wsSrc.Cells(headerRow, lastCol)).Value2
    End If

    If headerRow + 1 = lastRow And nCols = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = wsSrc.Cells(headerRow + 1, firstCol).Value
    Else
        dataArr = wsSrc.Range(wsSrc.Cells(headerRow + 1, firstCol), wsSrc.Cells(lastRow, lastCol)).Value2
    End If

    ' ユニーク値ごとの行インデックス収集（空行・集計行は除外）
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(dataArr, 1)
        isBlankRow = True
        isSumRow = False
        For c = 1 To nCols
            cellStr = Trim$(セルの字(dataArr(r, c)))
            If Len(cellStr) > 0 Then
                isBlankRow = False
                If 集計の語か(cellStr) Then
                    isSumRow = True
                    Exit For
                End If
            End If
        Next c
        If Not isBlankRow And Not isSumRow Then
            valStr = Trim$(セルの字(dataArr(r, relCol)))
            If Len(valStr) > 0 Then
                If Not dict.exists(valStr) Then
                    Set dict(valStr) = New Collection
                End If
                dict(valStr).Add r
            End If
        End If
    Next r

    If dict.Count = 0 Then Exit Sub

    Dim oldScreen As Boolean, oldAlerts As Boolean
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' シート生成と一括転記（人のシートは消さず(2)を付け、棚が作ったシートは作り直す）
    Set 今回 = CreateObject("Scripting.Dictionary")
    Dim createdCount As Long: createdCount = 0
    For Each key In dict.keys
        valStr = CStr(key)
        sheetTitle = 棚のシート名(wb, valStr, 今回)
        Set wsNew = wb.Worksheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        wsNew.Name = sheetTitle
        棚のシートの印 wsNew

        ' 見出し転記と書式復元
        表の整理_下請け.文字を守って書く wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(1, nCols)), headerArr ' 書き戻し: 新しい所
        For c = 1 To nCols
            If Len(hdrFmts(c)) > 0 Then wsNew.Cells(1, c).NumberFormat = hdrFmts(c)
        Next c

        ' データ抽出と一括代入
        Set colList = dict(key)
        matchCount = colList.Count
        If matchCount > 0 Then
            ReDim outArr(1 To matchCount, 1 To nCols)
            For outR = 1 To matchCount
                srcR = colList(outR)
                For c = 1 To nCols
                    outArr(outR, c) = dataArr(srcR, c)
                Next c
            Next outR
            表の整理_下請け.文字を守って書く wsNew.Range(wsNew.Cells(2, 1), wsNew.Cells(matchCount + 1, nCols)), outArr ' 書き戻し: 新しい所
            For c = 1 To nCols
                If Len(colFmts(c)) > 0 Then
                    wsNew.Range(wsNew.Cells(2, c), wsNew.Cells(matchCount + 1, c)).NumberFormat = colFmts(c)
                End If
            Next c
        End If

        ' 枠固定
        On Error Resume Next
        wsNew.Activate
        wsNew.rows(2).Select
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.FreezePanes = True
        End If
        wsNew.Range("A1").Select
        On Error GoTo 0

        createdCount = createdCount + 1
    Next key

    コンボ道具.gRecipeContext.NewSheetsCreated = コンボ道具.gRecipeContext.NewSheetsCreated + createdCount

    On Error Resume Next
    wsSrc.Activate
    On Error GoTo 0

    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen

    If 知らせる And Not コンボ道具.mレシピ中 Then
        Application.StatusBar = "選んだ列の値ごとにシートを分割する: " & createdCount & " 個のシートに振り分けました。"   ' 窓を出さない＝先撃ちで止まらない（2026-10-21）
    End If
End Sub

Sub 選んだ列の数値を指定桁数でゼロ埋め文字列にする_本体(ByVal 指定桁数 As Long)
    ' 選んだ列の数値を指定桁数でゼロ埋め文字列にする の中身。指定桁数が 0 以下なら桁数を聞く（レシピの中では 4）
    Dim ws As Worksheet, rng As Range, targetCol As Range, ur As Range
    Dim inVal As String, padLen As Long
    Dim dataArr As Variant, formArr As Variant, r As Long
    Dim v As Variant, fStr As String, numVal As Double, fmtStr As String
    Dim startRow As Long, endRow As Long, headerRow As Long, maxRow As Long
    Dim colRng As Range
    Dim chg() As Boolean, changed As Boolean

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection

    ' 桁数の決定（引数が0なら対話形式で聞く。レシピ実行時・非対話時は安全に自動決定）
    If 指定桁数 <= 0 Then
        If コンボ道具.mレシピ中 Then
            inVal = "4"
        Else
            inVal = InputBox("ゼロ埋めする桁数を入力してください（例: 4, 5, 8）", "ゼロ埋めパディング", "4")
        End If
        If Len(Trim$(inVal)) = 0 Or Not IsNumeric(inVal) Then Exit Sub
        padLen = CLng(inVal)
    Else
        padLen = 指定桁数
    End If
    If padLen <= 0 Or padLen > 30 Then Exit Sub

    fmtStr = String$(padLen, "0")
    headerRow = ur.row
    maxRow = 表の本文の最終行(ws)
    If maxRow < headerRow Then maxRow = ur.row + ur.rows.Count - 1

    Application.ScreenUpdating = False

    ' 選択されている各列を処理（見出し行と合計行・数式セルは保護）
    For Each targetCol In rng.Columns
        Set colRng = Intersect(targetCol.EntireColumn, ur)
        If Not colRng Is Nothing Then
            startRow = colRng.row
            endRow = colRng.row + colRng.rows.Count - 1
            If endRow > maxRow Then endRow = maxRow

            ' 複数行ある表で先頭行が見出し（非数値または1行目）ならデータ開始行を+1する
            If startRow = headerRow And endRow > headerRow Then
                If Not IsNumeric(ws.Cells(headerRow, targetCol.Column).Value2) Or Len(Trim$(CStr(ws.Cells(headerRow, targetCol.Column).Value2))) = 0 Then
                    startRow = headerRow + 1
                End If
            End If

            If startRow <= endRow Then
                Set colRng = ws.Range(ws.Cells(startRow, targetCol.Column), ws.Cells(endRow, targetCol.Column))
                If colRng.Cells.CountLarge = 1 Then
                    v = colRng.Value2
                    If Not colRng.HasFormula And IsNumeric(v) And Len(CStr(v)) > 0 And VarType(v) <> vbDate Then
                        numVal = CDbl(v)
                        If numVal >= 0 And numVal = Fix(numVal) Then
                            colRng.NumberFormatLocal = "@"
                            colRng.Cells(1, 1).Value = Format$(numVal, fmtStr)
                        End If
                    End If
                Else
                    ' 変えるセル（0 以上の整数）だけ文字にして書く。数式・負の数・小数・空欄は触らない（列ごと "@" にして書き戻すと
                    ' 数式が文字になり、空欄が空文字に化けて =C10*2 がエラーになった・2026-10-22）
                    dataArr = colRng.Value2
                    formArr = colRng.Formula
                    ReDim chg(1 To UBound(dataArr, 1), 1 To 1)
                    changed = False
                    For r = 1 To UBound(dataArr, 1)
                        v = dataArr(r, 1)
                        fStr = CStr(formArr(r, 1))
                        If Left$(fStr, 1) <> "=" And IsNumeric(v) And Len(CStr(v)) > 0 And VarType(v) <> vbDate Then
                            numVal = CDbl(v)
                            If numVal >= 0 And numVal = Fix(numVal) Then
                                dataArr(r, 1) = Format$(numVal, fmtStr)
                                chg(r, 1) = True
                                changed = True
                            End If
                        End If
                    Next r
                    If changed Then 変えたセルだけ書く colRng.Cells(1, 1), dataArr, chg, True
                End If
            End If
        End If
    Next targetCol

    Application.ScreenUpdating = True
End Sub
