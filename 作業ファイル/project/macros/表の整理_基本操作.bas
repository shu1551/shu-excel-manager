Attribute VB_Name = "表の整理_基本操作"

' 表の整理_基本操作 - 棚（言葉で頼んで撃つマクロ）の基本の操作（2026-10-28 リボンを目録にして棚にそろえた）
'   Excel のリボンの全コマンド（Microsoft の Office Fluent UI Command Identifiers・fixtures\ribbon_catalog.tsv）を
'   棚にする／もう棚にある／棚にしない の 3 つに分け、棚にする物を分類ごとの「どこを × どうする」の 6 本にまとめた。
'   どこを＝見出し・合計行・選んだ範囲・見出しの名前で指した列（「金額の列」）。どうする＝頼みの文から読む（大きく・赤・中央・
'   太字・3 桁区切り・幅をそろえる など）。頼む回数の一番多い基本の書式（「見出しの文字を大きくして赤にして」）が棚に当たらず
'   AI に回っていた（10/25）。どれも値は書かない（書式・幅・絞り込み・設定・シートの名前だけを変える）
'   1 操作 1 本の棚（表の整理_書式 の 中央揃え・3 桁区切り など）は選んだ範囲の頼みのまま残し、場所を言った頼みはこちらへ回す
Sub 頼みの文で指した所の書式を変える()
    ' 依頼の語: 見出しの文字|合計行の文字|見出しの背景|合計行の背景|見出しを中央|見出しを左寄せ|見出しを右寄せ|列の文字を
    ' 依頼の組: 見出し,ヘッダー,項目名,項目の行,合計行,合計の行,総計の行,小計の行,選んだ範囲,選んだセル,選んだ所,選択範囲,選択した,この範囲,このセル,この列,この行,の列+大きく,小さく,文字サイズ,フォントサイズ,文字の大きさ,ポイント,太字,ボールド,細字,斜体,イタリック,下線,取り消し線,赤,青,緑,黄,黒,白,灰色,グレー,オレンジ,紫,紺,ピンク,水色,茶色,塗,背景,網掛,中央,真ん中,センター,左寄せ,左揃え,左詰め,右寄せ,右揃え,右詰め,上揃え,下揃え,上下中央,上に寄せ,下に寄せ,縦書き,横書き,折り返,セル内で改行,縮小して,字下げ,インデント,明朝,ゴシック,メイリオ,書体,フォント,罫線,二重線,太線,細線,線を引+-一覧,調べ,探して,どこ,確認,チェック,点検,条件付き,グラフ,ピボット,スパーク,マイナス,負の,以上,以下,未満,超え,下回,上回,より大き,より小さ,場合,なら,だけ,ごとに,重複,空欄,空白,エラー,表示形式,桁区切り,カンマ,通貨,円マーク,パーセント,小数,和暦,幅,高さ,隠,非表示,並べ,絞,フィルタ,印刷,シート見出し,シートの見出し,タブ,美化,デザイン,テーマ,ダッシュボード,KPI,行を足,列を足,挿入,色分け,交互,縞,教え
    ' 階級: 基本
    ' 扱う: 文字の大きさ 文字の色 塗りつぶし 太字 斜体 下線 配置 折り返し 書体 罫線
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文の「どこを」（見出し・合計行・選んだ範囲・「金額の列」のように見出しの名前で指した列）を、「どうする」のとおりに変える
    ' （2026-10-28 リボンの目録の ホーム＞フォント・配置 から）: 大きく・小さく・N ポイント・文字の色・背景の色・太字・斜体・下線・
    ' 取り消し線・中央／左／右・上下中央・縦書き・折り返し・縮小して全体を表示・字下げ・書体・罫線（下に二重線・外枠・格子）。値は触らない
    Dim 頼 As String, ws As Worksheet, rng As Range, 何 As String, なぜ As String, 節 As Variant, 済 As String
    頼 = 基_頼み()
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「見出しの文字を大きくして赤にして」「金額の列を右寄せにして」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    If ws.ProtectContents Then 先撃ちに断る ws.Name & " は保護されているので変えられません（「シートの保護を解除して」と頼むか、校閲＞シート保護の解除）": Exit Sub
    Set rng = 基_指した所(ws, 頼, True, 何, なぜ)
    If rng Is Nothing Then 先撃ちに断る なぜ: Exit Sub
    Application.ScreenUpdating = False
    For Each 節 In 基_節に分ける(頼)
        If Len(節) > 0 Then 済 = 済 & 基_文字の書式(rng, CStr(節))
    Next 節
    Application.ScreenUpdating = True
    If Len(済) = 0 Then
        先撃ちに断る "どう変えるか分かりませんでした（例: 大きく・14ポイント・赤・背景を黄色・太字・中央・折り返し・下に二重線）"
        Exit Sub
    End If
    Application.StatusBar = 何 & " を変えました: " & Mid$(済, 2)
End Sub

Private Function 基_文字の書式(ByVal rng As Range, ByVal s As String) As String
    ' 1 つの節（「見出しの文字を大きく」「背景を黄色に」）を rng に当てる。→ 変えた物（「・大きく・赤」）。何も読めなければ ""
    Dim out As String, n As Double, 名 As String, 塗り As Boolean, t As String, ar As Range, c As Range, f As String
    If 含む語(s, "罫線|二重線|太線|細線|点線|破線|線で囲|に線|線を消|線をなくし|線を引") And Not 含む語(s, "下線を引|取り消し線") Then
        基_文字の書式 = 基_罫線(rng, s)
        Exit Function
    End If
    ' 文字の大きさ
    n = -1
    If 含む語(s, "ポイント|pt|サイズ|大きさ") Then n = 基_数(s)
    If n >= 6 And n <= 409 And InStr(s, "倍") = 0 Then
        rng.Font.Size = n: out = out & "・" & n & "ポイント"
    ElseIf 含む語(s, "大きく|小さく") Then
        n = 2
        If 含む語(s, "もっと|かなり|すごく|うんと|ぐっと") Then n = 4
        If InStr(s, "小さく") > 0 Then n = -n
        For Each ar In rng.Areas
            If IsNull(ar.Font.Size) Then
                For Each c In ar.Cells
                    c.Font.Size = 基_範囲内(c.Font.Size + n, 6, 409)
                Next c
            Else
                ar.Font.Size = 基_範囲内(ar.Font.Size + n, 6, 409)
            End If
        Next ar
        out = out & IIf(n > 0, "・大きく", "・小さく")
    End If
    ' 書体
    f = 基_書体(s)
    If Len(f) > 0 Then rng.Font.Name = f: out = out & "・" & f
    ' 太字・斜体・下線・取り消し線
    If 基_やめか(s, "太字") Or 基_やめか(s, "ボールド") Or InStr(s, "細字") > 0 Then
        rng.Font.Bold = False: out = out & "・太字なし"
    ElseIf 含む語(s, "太字|ボールド") Then
        rng.Font.Bold = True: out = out & "・太字"
    End If
    If 基_やめか(s, "斜体") Or 基_やめか(s, "イタリック") Then
        rng.Font.Italic = False: out = out & "・斜体なし"
    ElseIf 含む語(s, "斜体|イタリック") Then
        rng.Font.Italic = True: out = out & "・斜体"
    End If
    If 基_やめか(s, "取り消し線") Then
        rng.Font.Strikethrough = False: out = out & "・取り消し線なし"
    ElseIf InStr(s, "取り消し線") > 0 Then
        rng.Font.Strikethrough = True: out = out & "・取り消し線"
    ElseIf 基_やめか(s, "下線") Then
        rng.Font.Underline = xlUnderlineStyleNone: out = out & "・下線なし"
    ElseIf InStr(s, "二重下線") > 0 Then
        rng.Font.Underline = xlUnderlineStyleDouble: out = out & "・二重下線"
    ElseIf 含む語(s, "下線|アンダーライン") Then
        rng.Font.Underline = xlUnderlineStyleSingle: out = out & "・下線"
    End If
    ' 横の寄せ（「上下中央」は縦の寄せなので外して見る）
    t = Replace(Replace(Replace(s, "上下中央", ""), "縦の中央", ""), "縦中央", "")
    If 含む語(t, "中央|真ん中|センター|センタリング") Then
        rng.HorizontalAlignment = xlCenter: out = out & "・中央揃え"
    ElseIf 含む語(s, "左寄|左揃|左詰|左に寄|左ぞろえ") Then
        rng.HorizontalAlignment = xlLeft: out = out & "・左寄せ"
    ElseIf 含む語(s, "右寄|右揃|右詰|右に寄|右ぞろえ") Then
        rng.HorizontalAlignment = xlRight: out = out & "・右寄せ"
    ElseIf InStr(s, "均等割") > 0 Then
        rng.HorizontalAlignment = xlDistributed: out = out & "・均等割り付け"
    End If
    ' 縦の寄せ
    If 含む語(s, "上下中央|縦の中央|縦中央|上下の中央") Then
        rng.VerticalAlignment = xlCenter: out = out & "・上下中央"
    ElseIf 含む語(s, "上揃|上に寄|上詰|上ぞろえ") Then
        rng.VerticalAlignment = xlTop: out = out & "・上揃え"
    ElseIf 含む語(s, "下揃|下に寄|下詰|下ぞろえ") Then
        rng.VerticalAlignment = xlBottom: out = out & "・下揃え"
    End If
    ' 縦書き・横書き
    If InStr(s, "縦書き") > 0 Then
        rng.Orientation = xlVertical: out = out & "・縦書き"
    ElseIf InStr(s, "横書き") > 0 Then
        rng.Orientation = xlHorizontal: out = out & "・横書き"
    End If
    ' 折り返し・縮小して全体を表示・字下げ
    If 基_やめか(s, "折り返し") Or 含む語(s, "折り返さな") Then
        rng.WrapText = False: out = out & "・折り返しなし"
    ElseIf InStr(s, "折り返") > 0 Or 含む語(s, "セル内で改行|セルの中で改行") Then
        rng.ShrinkToFit = False: rng.WrapText = True: out = out & "・折り返し"
    ElseIf 含む語(s, "縮小して|縮小表示|縮めて全体") Then
        rng.WrapText = False: rng.ShrinkToFit = True: out = out & "・縮小して全体を表示"
    End If
    If 基_やめか(s, "字下げ") Or 基_やめか(s, "インデント") Then
        rng.IndentLevel = 0: out = out & "・字下げなし"
    ElseIf 含む語(s, "字下げ|インデント") Then
        For Each ar In rng.Areas
            If IsNull(ar.IndentLevel) Then ar.IndentLevel = 1 Else ar.IndentLevel = 基_範囲内(ar.IndentLevel + 1, 0, 15)
        Next ar
        out = out & "・字下げ"
    End If
    ' 色（節に 背景・塗り があれば背景、無ければ文字の色）
    塗り = 含む語(s, "背景|塗|網掛|セルの色|地の色|地色")
    名 = 基_色の名(s)
    If Len(名) > 0 Then
        If 塗り Then
            rng.Interior.Color = 基_色(名, True): out = out & "・背景を" & 名
        Else
            rng.Font.Color = 基_色(名, False): out = out & "・文字を" & 名
        End If
    ElseIf 基_やめか(s, "背景") Or 基_やめか(s, "塗り") Or 基_やめか(s, "塗りつぶし") Or 基_やめか(s, "色") Then
        rng.Interior.Pattern = xlNone: out = out & "・背景なし"
    End If
    基_文字の書式 = out
End Function

Private Function 基_範囲内(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then v = lo
    If v > hi Then v = hi
    基_範囲内 = v
End Function

Private Function 基_書体(ByVal s As String) As String
    ' 頼みの文の書体の名前（ならした後の字で見る：「ＭＳ Ｐゴシック」→ MSPゴシック）。無ければ ""
    If InStr(s, "MSPゴシック") > 0 Then
        基_書体 = "ＭＳ Ｐゴシック"
    ElseIf InStr(s, "MSゴシック") > 0 Then
        基_書体 = "ＭＳ ゴシック"
    ElseIf InStr(s, "MSP明朝") > 0 Then
        基_書体 = "ＭＳ Ｐ明朝"
    ElseIf InStr(s, "MS明朝") > 0 Then
        基_書体 = "ＭＳ 明朝"
    ElseIf InStr(s, "BIZUDPゴシック") > 0 Then
        基_書体 = "BIZ UDPゴシック"
    ElseIf InStr(s, "BIZUDゴシック") > 0 Then
        基_書体 = "BIZ UDゴシック"
    ElseIf InStr(s, "BIZUDP明朝") > 0 Then
        基_書体 = "BIZ UDP明朝"
    ElseIf InStr(s, "BIZUD明朝") > 0 Then
        基_書体 = "BIZ UD明朝"
    ElseIf InStr(s, "MeiryoUI") > 0 Then
        基_書体 = "Meiryo UI"
    ElseIf InStr(s, "メイリオ") > 0 Then
        基_書体 = "メイリオ"
    ElseIf InStr(s, "明朝") > 0 Then
        基_書体 = "游明朝"
    ElseIf InStr(s, "ゴシック") > 0 Then
        基_書体 = "游ゴシック"
    End If
End Function

Private Function 基_罫線(ByVal rng As Range, ByVal s As String) As String
    ' 節の罫線の頼み（下に二重線・上に罫線・外枠を太線・格子・罫線を消す）を当てる。→ 変えた物
    Dim 辺 As Variant, E As Variant, 種 As Long, 太 As Long, 名 As String, 何 As String, ar As Range, 消す As Boolean
    消す = 基_やめか(s, "罫線") Or 基_やめか(s, "線")
    種 = xlContinuous: 太 = xlThin: 何 = "罫線"
    If 含む語(s, "二重線|二重の線") Then
        種 = xlDouble: 太 = xlThick: 何 = "二重線"
    ElseIf InStr(s, "極太") > 0 Then
        太 = xlThick: 何 = "極太線"
    ElseIf 含む語(s, "太線|太い線|太く") Then
        太 = xlMedium: 何 = "太線"
    ElseIf 含む語(s, "点線") Then
        種 = xlDot: 何 = "点線"
    ElseIf 含む語(s, "破線") Then
        種 = xlDash: 何 = "破線"
    End If
    If 含む語(s, "外枠|周り|まわり|囲") Then
        辺 = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight): 何 = "外枠を" & 何
    ElseIf 含む語(s, "の下|下に|下側|下の線|下辺") Then
        辺 = Array(xlEdgeBottom): 何 = "下に" & 何
    ElseIf 含む語(s, "の上|上に|上側|上の線|上辺") Then
        辺 = Array(xlEdgeTop): 何 = "上に" & 何
    ElseIf 含む語(s, "の左|左に|左側") Then
        辺 = Array(xlEdgeLeft): 何 = "左に" & 何
    ElseIf 含む語(s, "の右|右に|右側") Then
        辺 = Array(xlEdgeRight): 何 = "右に" & 何
    Else
        辺 = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
        If Not 消す Then 何 = "格子の" & 何
    End If
    名 = 基_色の名(s)
    On Error Resume Next                 ' 1 行・1 列の範囲に内側の線は無い
    For Each ar In rng.Areas
        For Each E In 辺
            With ar.Borders(E)
                If 消す Then
                    .LineStyle = xlNone
                Else
                    .LineStyle = 種
                    If 種 <> xlDouble Then .Weight = 太
                    If Len(名) > 0 Then .Color = 基_色(名, False) Else .ColorIndex = xlAutomatic
                End If
            End With
        Next E
    Next ar
    On Error GoTo 0
    基_罫線 = "・" & IIf(消す, "罫線なし", 何 & IIf(Len(名) > 0, "（" & 名 & "）", ""))
End Function

Sub 頼みの文で指した列の表示形式を変える()
    ' 依頼の語: 列の表示形式|曜日も表示|曜日を表示|曜日付き|列を3桁区切り|列を桁区切り|列をカンマ区切り|列をパーセント|列を円マーク|列を通貨|列の小数|合計行を3桁区切り|合計行の表示形式
    ' 依頼の組: 合計行,合計の行,総計の行,選んだ範囲,選んだセル,選択範囲,この範囲,このセル,この列,の列+桁区切り,3桁,３桁,カンマ,円マーク,￥,通貨,パーセント,％,小数,千円単位,千円で,百万円,万円単位,日付の形,日付の表示,年月日,曜日,時刻,文字列,表示形式,マイナスを赤,▲,ゼロを表示しない,0を表示しない,ゼロを隠,0を隠,整数+-一覧,調べ,探して,どこ,確認,チェック,条件付き,グラフ,ピボット,以上,以下,未満,超え,下回,上回,場合,なら,だけ,行を足,列を足,足して,追加,作って,作る,集計,ゼロ埋め,0埋め,先頭のゼロ,ゼロを付,和暦,元号,を数にして,数値に,日付にして,日付に直,文字の数字,文字の日付,行の,で付け,でつけ,で足,構成比,クエリ,区切って,分けて,分割,きれい,掃除,整え,教え
    ' 階級: 基本
    ' 扱う: 表示形式 桁区切り 通貨 パーセント 小数 日付の形
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文で指した列（見出しの名前・この列・選んだ範囲・合計行）の表示形式を変える（2026-10-28 リボンの目録の ホーム＞数値 から）:
    ' 3 桁区切り・円マーク・パーセント・小数の桁・千円単位・マイナスを赤／▲・ゼロを表示しない・日付の形・曜日・時刻・文字列・標準。
    ' 表示形式だけを変え、値は書かない（文字の「00123」や文字の日付は文字のまま＝数や日付に化けない）
    Dim 頼 As String, ws As Worksheet, rng As Range, 何 As String, なぜ As String, 形 As String, 名 As String, ar As Range
    頼 = 基_頼み()
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「金額の列を3桁区切りにして」「率の列を小数1桁のパーセントにして」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    If ws.ProtectContents Then 先撃ちに断る ws.Name & " は保護されているので変えられません（「シートの保護を解除して」と頼むか、校閲＞シート保護の解除）": Exit Sub
    形 = 基_表示形式(頼, 名)
    If Len(形) = 0 Then
        先撃ちに断る "どの表示形式にするか分かりませんでした（例: 3桁区切り・円マーク・パーセント・小数1桁・千円単位・yyyy/mm/dd・文字列）"
        Exit Sub
    End If
    Set rng = 基_指した所(ws, 頼, False, 何, なぜ)
    If rng Is Nothing Then 先撃ちに断る なぜ: Exit Sub
    For Each ar In rng.Areas
        ar.NumberFormatLocal = 形
    Next ar
    Application.StatusBar = 何 & " の表示形式を " & 名 & "（" & 形 & "）にしました"
End Sub

Private Function 基_表示形式(ByVal 頼 As String, ByRef 名 As String) As String
    ' 頼みの文 → 表示形式の字（NumberFormatLocal）と、知らせに出す名前。読めなければ ""
    Dim 小 As Long, 点 As String, 本 As String, 負 As String, 零 As Boolean, n As Double, 区切る As Boolean
    名 = ""
    If InStr(頼, "文字列") > 0 Then 名 = "文字列": 基_表示形式 = "@": Exit Function
    If 含む語(頼, "表示形式を標準|標準の表示形式|標準に戻") Then 名 = "標準": 基_表示形式 = "G/標準": Exit Function
    If 含む語(頼, "時刻|時分") Then
        名 = "時刻": 基_表示形式 = IIf(InStr(頼, "秒") > 0, "h:mm:ss", "h:mm")
        Exit Function
    End If
    If 含む語(頼, "日付の形|日付の表示|年月日|月日|年月|曜日|yyyy|yy/|mm/dd|m/d") Then
        名 = "日付"
        Select Case True
            Case InStr(LCase$(頼), "yyyy/mm/dd") > 0: 本 = "yyyy/mm/dd"
            Case InStr(LCase$(頼), "yyyy-mm-dd") > 0: 本 = "yyyy-mm-dd"
            Case InStr(LCase$(頼), "yy/mm/dd") > 0: 本 = "yy/mm/dd"
            Case InStr(LCase$(頼), "mm/dd") > 0: 本 = "mm/dd"
            Case InStr(頼, "年月日") > 0: 本 = "yyyy""年""m""月""d""日"""
            Case InStr(頼, "月日") > 0: 本 = "m""月""d""日"""
            Case InStr(頼, "年月") > 0: 本 = "yyyy""年""m""月"""
            Case InStr(LCase$(頼), "m/d") > 0 And InStr(LCase$(頼), "yyyy") = 0: 本 = "m/d"
            Case Else: 本 = "yyyy/m/d"
        End Select
        If InStr(頼, "曜日") > 0 Then 本 = 本 & "(aaa)": 名 = "日付と曜日"
        基_表示形式 = 本
        Exit Function
    End If
    ' 数
    小 = -1
    If InStr(頼, "小数") > 0 Then
        n = 基_数(頼, "小数")
        If n >= 0 And n <= 10 Then 小 = n
        If 含む語(頼, "小数なし|小数を消|小数を表示しない|小数はいらな") Then 小 = 0
    End If
    If 含む語(頼, "整数") Then 小 = 0
    点 = ""
    If 小 > 0 Then 点 = "." & String$(小, "0")
    区切る = 含む語(頼, "桁区切り|3桁|カンマ") Or 小 >= 0
    If 含む語(頼, "パーセント|%") Then
        本 = "0" & 点 & "%": 名 = "パーセント"
    ElseIf 含む語(頼, "円マーク|￥|\|通貨|円表示|円を付|円をつけ|円にして|円で表示") Then
        本 = "\#,##0" & 点: 名 = "円マーク"
    ElseIf 含む語(頼, "百万円") Then
        本 = "#,##0" & 点 & ",,": 名 = "百万円単位"
    ElseIf 含む語(頼, "千円") Then
        本 = "#,##0" & 点 & ",": 名 = "千円単位"
    ElseIf 区切る Or 含む語(頼, "マイナス|負の|▲|△|ゼロを|0を") Then
        本 = "#,##0" & 点: 名 = IIf(小 >= 0 And Not 含む語(頼, "桁区切り|3桁|カンマ"), "小数" & 小 & "桁", "3桁区切り")
    Else
        Exit Function
    End If
    負 = ""
    If 含む語(頼, "マイナスを赤|負の数を赤|マイナスは赤|マイナスの数を赤|赤字") Then
        負 = "[赤]-" & 本: 名 = 名 & "・マイナスは赤"
    ElseIf InStr(頼, "▲") > 0 Then
        負 = """▲""" & 本: 名 = 名 & "・マイナスは▲"
    ElseIf InStr(頼, "△") > 0 Then
        負 = """△""" & 本: 名 = 名 & "・マイナスは△"
    End If
    零 = 含む語(頼, "ゼロを表示しない|0を表示しない|ゼロを隠|0を隠|ゼロは表示しない|0は表示しない|ゼロを非表示|0を非表示|ゼロは空|0は空")
    If 零 Then 名 = 名 & "・ゼロは表示しない"
    If Len(負) > 0 Or 零 Then
        If Len(負) = 0 Then 負 = "-" & 本
        本 = 本 & ";" & 負 & IIf(零, ";", "")
    End If
    基_表示形式 = 本
End Function

Sub 頼みの文で指した列と行の幅と表示を変える()
    ' 依頼の語: 列の幅|列幅を広|列幅を狭|列幅を自動|行の高さを高|行の高さを低|の列を隠|列をグループ|行をグループ
    ' 依頼の組: の列,列の,列を,列は,この列,選んだ列,見出しの行,見出し行,合計行,合計の行,この行,選んだ行,行の,行を,全部の列,すべての列,全ての列,表の列+幅,高さ,広げ,広く,狭く,狭め,高く,低く,自動調整,自動で合わせ,ぴったり,文字に合わせ,幅をそろえ,幅を揃え,高さをそろえ,高さを揃え,同じ幅,同じ高さ,隠して,隠す,非表示,グループ化,グループを解除,グループ解除,折りたた+-一覧,調べ,探して,どこ,確認,印刷,ページ,全シート,どのシート,シートごと,他のシート,前のシート,1枚目,一枚目,グラフ,ピボット,再表示,全部表示,表示に戻,見えるように,シートを,倍率,教え
    ' 階級: 基本
    ' 扱う: 列幅 行の高さ 非表示 グループ化
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文で指した列（見出しの名前・この列・選んだ列・全部の列）と行（見出しの行・合計行・この行）の幅と高さと表示を変える
    ' （2026-10-28 リボンの目録の ホーム＞セル＞書式・データ＞アウトライン から）: 幅を広げる／狭める／N にする／自動で合わせる／そろえる・
    ' 高さを高く／低く／N に／自動・隠す・グループ化／グループ解除。隠した列と行を戻すのは「隠れた行と列を全部表示する」
    Dim 頼 As String, ws As Worksheet, 見r As Long, 左 As Long, 右 As Long, 末r As Long, 合r As Long, 下 As Long, 表あり As Boolean
    Dim 行の As Boolean, 対 As Range, 何 As String, 済 As String, cols As Collection, c As Variant, 最小 As Long, 最大 As Long
    Dim 選 As Boolean, 全部 As Boolean, n As Double, 足す As Double, 大 As Double, ar As Range, k As Long
    頼 = 基_頼み()
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「金額の列の幅を広げて」「見出しの行の高さを高くして」「備考の列を隠して」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    If ws.ProtectContents Then 先撃ちに断る ws.Name & " は保護されているので変えられません（「シートの保護を解除して」と頼むか、校閲＞シート保護の解除）": Exit Sub
    表あり = 基_表を読む(ws, 見r, 左, 右, 末r, 合r)
    下 = 末r
    If 合r > 下 Then 下 = 合r
    選 = 含む語(頼, "選んだ|選択|この列|この行") And TypeName(Selection) = "Range"
    行の = 含む語(頼, "高さ|高く|低く|行高|この行|選んだ行|見出しの行|見出し行|合計行|合計の行|全部の行|すべての行|表の行|行を")
    If 含む語(頼, "列の幅|列幅|の列|列を|この列|選んだ列") Then 行の = 含む語(頼, "高さ|高く|低く|行高")
    全部 = 含む語(頼, "全部の列|すべての列|全ての列|表の列|列を全部|全列|全部の行|すべての行|表の行|行を全部")
    If 行の Then
        If 全部 And 表あり Then Set 対 = ws.Range(ws.rows(見r), ws.rows(下)): 何 = "表の行"
        If 含む語(頼, "見出しの行|見出し行|見出し") And 表あり And Not 全部 Then Set 対 = 基_足す(対, ws.rows(見r)): 何 = 何 & "・見出しの行"
        If 含む語(頼, "合計行|合計の行") And 合r > 0 Then Set 対 = 基_足す(対, ws.rows(合r)): 何 = 何 & "・合計行"
        If 選 And 対 Is Nothing Then Set 対 = Selection.EntireRow: 何 = "選んだ行"
    Else
        If 全部 And 表あり Then Set 対 = ws.Range(ws.Columns(左), ws.Columns(右)): 何 = "表の列"
        If Not 全部 And 表あり Then
            Set cols = 基_頼みの列(ws, 頼, 見r, 左, 右)
            If cols.Count > 0 Then
                最小 = 右: 最大 = 左
                For Each c In cols
                    If c < 最小 Then 最小 = c
                    If c > 最大 Then 最大 = c
                Next c
                If InStr(頼, "から") > 0 And cols.Count >= 2 Then
                    Set 対 = ws.Range(ws.Columns(最小), ws.Columns(最大))
                    何 = 見出しの字(ws.Cells(見r, 最小)) & "から" & 見出しの字(ws.Cells(見r, 最大)) & "の列"
                Else
                    For Each c In cols
                        Set 対 = 基_足す(対, ws.Columns(c))
                        何 = 何 & "・" & 見出しの字(ws.Cells(見r, c))
                    Next c
                    何 = Mid$(何, 2) & "の列"
                End If
            End If
        End If
        If 選 And 対 Is Nothing Then Set 対 = Selection.EntireColumn: 何 = "選んだ列"
    End If
    If Left$(何, 1) = "・" Then 何 = Mid$(何, 2)
    If 対 Is Nothing Then
        先撃ちに断る "どの" & IIf(行の, "行", "列") & "か分かりませんでした（「金額の列」のように見出しの名前で、または「見出しの行」「合計行」「選んだ列」「全部の列」で頼んでください）"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    If 含む語(頼, "グループを解除|グループ解除|グループ化を解除|グループをやめ|グループを外") Then
        On Error Resume Next
        For Each ar In 対.Areas
            ar.Ungroup
        Next ar
        On Error GoTo 0
        済 = " のグループを解除しました"
    ElseIf 含む語(頼, "グループ|折りたた") Then
        For Each ar In 対.Areas
            ar.Group
        Next ar
        済 = " をグループにまとめました"
    ElseIf 含む語(頼, "隠|非表示") Then
        For Each ar In 対.Areas
            If 行の Then ar.EntireRow.Hidden = True Else ar.EntireColumn.Hidden = True
        Next ar
        済 = " を隠しました"
    ElseIf 含む語(頼, "自動|ぴったり|文字に合わせ|中身に合わせ|内容に合わせ|長さに合わせ") Then
        For Each ar In 対.Areas
            If 行の Then
                ar.EntireRow.AutoFit
            ElseIf 表あり Then
                Intersect(ar.EntireColumn, ws.Range(ws.rows(見r), ws.rows(下))).Columns.AutoFit    ' 表の上の題で広がらないように表の行で合わせる
            Else
                ar.EntireColumn.AutoFit
            End If
        Next ar
        済 = " の" & IIf(行の, "高さ", "幅") & "を自動で合わせました"
    Else
        n = 基_数(頼, IIf(行の, "高さ", "幅"))
        If 含む語(頼, "そろえ|揃え|同じ|合わせ") Then
            大 = 0
            For Each ar In 対.Areas
                For k = 1 To IIf(行の, ar.rows.Count, ar.Columns.Count)
                    If 行の Then
                        If ar.rows(k).RowHeight > 大 Then 大 = ar.rows(k).RowHeight
                    Else
                        If ar.Columns(k).ColumnWidth > 大 Then 大 = ar.Columns(k).ColumnWidth
                    End If
                Next k
            Next ar
            If n > 0 Then 大 = n
            For Each ar In 対.Areas
                If 行の Then ar.RowHeight = 基_範囲内(大, 1, 409) Else ar.ColumnWidth = 基_範囲内(大, 1, 255)
            Next ar
            済 = " の" & IIf(行の, "高さ", "幅") & "を " & 大 & " にそろえました"
        ElseIf n > 0 And 含む語(頼, "幅を|高さを|幅は|高さは") Then
            For Each ar In 対.Areas
                If 行の Then ar.RowHeight = 基_範囲内(n, 1, 409) Else ar.ColumnWidth = 基_範囲内(n, 1, 255)
            Next ar
            済 = " の" & IIf(行の, "高さ", "幅") & "を " & n & " にしました"
        ElseIf 含む語(頼, "広げ|広く|大きく|高く|狭く|狭め|小さく|低く|縮め") Then
            If 行の Then 足す = 6 Else 足す = 4
            If 含む語(頼, "もっと|かなり|すごく|うんと|ぐっと") Then 足す = 足す + 足す
            If 含む語(頼, "狭く|狭め|小さく|低く|縮め") Then 足す = -足す
            For Each ar In 対.Areas
                For k = 1 To IIf(行の, ar.rows.Count, ar.Columns.Count)
                    If 行の Then
                        ar.rows(k).RowHeight = 基_範囲内(ar.rows(k).RowHeight + 足す, 6, 409)
                    Else
                        ar.Columns(k).ColumnWidth = 基_範囲内(ar.Columns(k).ColumnWidth + 足す, 1, 255)
                    End If
                Next k
            Next ar
            If 行の Then
                済 = " の高さを" & IIf(足す > 0, "高く", "低く") & "しました"
            Else
                済 = " の幅を" & IIf(足す > 0, "広げ", "狭め") & "ました"
            End If
        End If
    End If
    Application.ScreenUpdating = True
    If Len(済) = 0 Then
        先撃ちに断る "どう変えるか分かりませんでした（例: 幅を広げる・幅を15に・幅を自動で合わせる・幅をそろえる・高さを高く・隠す・グループ化）"
        Exit Sub
    End If
    Application.StatusBar = 何 & 済
End Sub

Sub 頼みの文で指した列で表を絞り込む()
    ' 依頼の語: フィルターを付け|フィルターをつけ|フィルターを外|フィルターを解除|フィルターをかけ|フィルターを掛け|フィルターをクリア|表の絞り込み|絞り込みをクリア|絞り込みを外|絞り込みをやめ|行だけ表示|行だけにして|以外の行を表示|以外の行だけ|以外の行を見せ|以外に絞|以外を表示|以外のレコード|だけを画面に|だけ画面に
    ' 依頼の組: 絞り込み,フィルター,の列,列で,列が,列を,列の,見出し,表に,表の,表で,行だけ,の行を表示,の行を見せ+フィルター,フィルタ,オートフィルタ,絞り込,絞って,で絞,に絞,だけ表示,だけにして,だけ見せ,だけを表示,だけ見え,以外の行,以外に絞,以外を表示,だけを画面,画面に出+-一覧,調べ,確認,抜き出,抽出,別シート,シートに,コピー,貼,書き出,印刷,グラフ,ピボット,ピポット,スライサー,並べ,ソート,上位,重複,隠,テーブルを並べ,教え,探
    ' 階級: 基本
    ' 扱う: フィルター 絞り込み
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文で表のフィルターを付ける・外す・絞り込む・絞り込みを解除する（2026-10-28 リボンの目録の データ＞並べ替えとフィルター から）:
    ' 「区分が東京の行だけ表示して」「金額の列で10000以上に絞り込んで」「担当が空欄の行だけ」「フィルターを外して」「絞り込みを解除して」。
    ' 行は消さず、見せる行を変えるだけ（合計の行は絞り込みの範囲に入れない）。並べ替えは「選んだ列で表を並べ替える」
    Dim 頼 As String, ws As Worksheet, 見r As Long, 左 As Long, 右 As Long, 末r As Long, 合r As Long, lo As ListObject
    Dim rng As Range, cols As Collection, c As Variant, 値 As String, 演 As String, 条件 As String, 済 As String, n As Long
    Dim ボタン As Boolean, 外す As Boolean, 解く As Boolean
    頼 = 基_頼み()
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「区分が東京の行だけ表示して」「フィルターを外して」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    If ws.ProtectContents Then 先撃ちに断る ws.Name & " は保護されているので変えられません（「シートの保護を解除して」と頼むか、校閲＞シート保護の解除）": Exit Sub
    If Not 基_表を読む(ws, 見r, 左, 右, 末r, 合r) Then 先撃ちに断る "表の見出しの行が見つかりません（見出しが 1 行ある表で頼んでください）": Exit Sub
    Set lo = 基_表のテーブル(ws)
    If Not lo Is Nothing Then
        Set rng = lo.Range                 ' テーブルの絞り込みは集計行を入れない
    ElseIf ws.AutoFilterMode Then
        Set rng = ws.AutoFilter.Range
    Else
        Set rng = ws.Range(ws.Cells(見r, 左), ws.Cells(末r, 右))
    End If
    ボタン = 含む語(頼, "フィルター|フィルタ|オートフィルタ|絞り込みボタン|絞り込みのボタン")
    外す = ボタン And 含む語(頼, "外|はず|消|取って|取り|なくし|無くし|いらな|不要") And Not 含む語(頼, "解除|クリア")
    解く = 含む語(頼, "解除|クリア|全部表示|すべて表示|全て表示|全部見せ|元に戻|戻して|絞り込みをやめ|絞り込みを外")

    If 外す Then
        If Not lo Is Nothing Then lo.ShowAutoFilter = False Else ws.AutoFilterMode = False
        Application.StatusBar = "フィルターを外しました（見えていなかった行も全部出ています）"
        Exit Sub
    End If
    If 解く Then
        If lo Is Nothing And Not ws.AutoFilterMode Then 先撃ちに断る ws.Name & " にはフィルターが付いていません（解除する絞り込みがありません）": Exit Sub
        On Error Resume Next
        If Not lo Is Nothing Then
            If lo.ShowAutoFilter Then lo.AutoFilter.ShowAllData
        ElseIf ws.FilterMode Then
            ws.ShowAllData
        End If
        On Error GoTo 0
        Application.StatusBar = "絞り込みを解除しました（フィルターのボタンは残しています）"
        Exit Sub
    End If

    Set cols = 基_頼みの列(ws, 頼, 見r, 左, 右)
    For Each c In cols
        値 = 基_絞る値(頼, 見出しの字(ws.Cells(見r, c)), 演)
        If c < rng.Column Or c > rng.Column + rng.Columns.Count - 1 Then
            先撃ちに断る "「" & 見出しの字(ws.Cells(見r, c)) & "」の列は今のフィルターの範囲（" & rng.Address(False, False) & "）の外です。先に「フィルターを外して」と頼んでください"
            Exit Sub
        End If
        If Len(演) > 0 Then
            If 演 = "含む" Then
                条件 = "=*" & 値 & "*"
            Else
                条件 = 演 & 値
            End If
            rng.AutoFilter Field:=c - rng.Column + 1, Criteria1:=条件
            済 = 済 & "・" & 見出しの字(ws.Cells(見r, c)) & " " & 条件
        End If
    Next c
    If Len(済) > 0 Then
        ' Excel は絞り込みの範囲を下に続く行まで広げる（A1:E6 を渡しても A1:E7 になり、合計の行が隠れた。2026-10-29 実測）
        If 合r > 0 And lo Is Nothing Then ws.rows(合r).Hidden = False
        On Error Resume Next
        n = rng.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
        On Error GoTo 0
        Application.StatusBar = Mid$(済, 2) & " で絞り込みました（見えている行 " & n & "）"
        Exit Sub
    End If
    If ボタン Then
        If Not lo Is Nothing Then
            lo.ShowAutoFilter = True
        ElseIf Not ws.AutoFilterMode Then
            rng.AutoFilter
        End If
        Application.StatusBar = "見出しにフィルターを付けました（" & rng.Address(False, False) & "）"
        Exit Sub
    End If
    If cols.Count > 0 Then
        先撃ちに断る "何で絞り込むか分かりませんでした（例:「区分が東京の行だけ表示して」「金額の列で10000以上に絞り込んで」）"
    Else
        先撃ちに断る "どの列で絞り込むか分かりませんでした（見出しの名前で「区分が東京の行だけ表示して」のように頼んでください）"
    End If
End Sub

Private Function 基_絞る値(ByVal 頼 As String, ByVal 見出し As String, ByRef 演 As String) As String
    ' 頼みの文の見出しの名前の後ろから、絞り込む値と比べ方を読む。演＝"=" "<>" ">=" "<=" ">" "<" "含む"。
    ' 「区分が東京の行だけ」→ =東京／「金額の列で10000以上」→ >=10000／「担当が空欄の行」→ =（空欄）／「区分が東京以外」→ <>東京
    Dim p As Long, s As String, 後 As String, E As Long, k As Long, w As Variant, 当 As String, h As String
    演 = ""
    h = ピボットの語ならし(見出し)
    p = InStr(1, 頼, h, vbBinaryCompare)
    当 = h
    If p = 0 Then p = 見出しの語の位置(頼, 見出し, 当)
    If p = 0 Or Len(当) = 0 Then Exit Function
    s = Mid$(頼, p + Len(当))
    For Each w In Array("の列で", "の列が", "の列は", "の列を", "の列の", "列で", "列が", "の欄が", "の欄で", "が", "は", "で", "を", "に", "=")
        If Left$(s, Len(w)) = w Then s = Mid$(s, Len(w) + 1): Exit For
    Next w
    E = Len(s) + 1
    For Each w In Array("の行", "だけ", "で絞", "に絞", "を絞", "で抽", "以上", "以下", "未満", "以降", "以前", "より", "以外", "を超", "超え", _
                        "を含", "が入", "を表示", "の物", "のもの", "を見", "で見", "にして", "にな", "で", "かつ", "、", "。", "|")
        k = InStr(1, s, w, vbBinaryCompare)
        If k > 0 And k < E Then E = k
    Next w
    後 = Mid$(s, E)
    s = Left$(s, E - 1)
    If 含む語(s, "空欄|空白|ブランク|未入力") Then
        基_絞る値 = ""
        演 = IIf(含む語(後, "以外|でない|じゃない"), "<>", "=")
        Exit Function
    End If
    If Len(s) = 0 Then Exit Function
    Select Case True
        Case Left$(後, 2) = "以上", Left$(後, 2) = "以降": 演 = ">="
        Case Left$(後, 2) = "以下", Left$(後, 2) = "以前": 演 = "<="
        Case Left$(後, 2) = "未満": 演 = "<"
        Case Left$(後, 2) = "を超", Left$(後, 2) = "超え": 演 = ">"
        Case Left$(後, 2) = "より"
            If 含む語(Mid$(後, 3, 3), "大き|多|後|上|高") Then 演 = ">" Else 演 = "<"
        Case Left$(後, 2) = "以外": 演 = "<>"
        Case Left$(後, 2) = "を含", Left$(後, 2) = "が入": 演 = "含む"
        Case Else: 演 = "="
    End Select
    If 演 <> "含む" And 演 <> "=" And 演 <> "<>" Then
        s = Replace(Replace(s, ",", ""), "円", "")
        For Each w In Array("個", "件", "人", "本", "枚", "台", "点", "回", "冊", "箱")   ' 「50個以上」の単位（2026-10-29）
            If Right$(s, Len(w)) = w And IsNumeric(Left$(s, Len(s) - Len(w))) Then s = Left$(s, Len(s) - Len(w))
        Next w
        If Right$(s, 1) = "万" And IsNumeric(Left$(s, Len(s) - 1)) Then s = CStr(CDbl(Left$(s, Len(s) - 1)) * 10000)
    End If
    基_絞る値 = s
End Function


Sub 頼みの文でシートの表示と印刷を変える()
    ' 依頼の語: 枠の固定を解除|固定を解除|固定を外|目盛線を消|目盛線を表示|目盛り線を消|枠線を消|改ページプレビュー|標準表示|ページレイアウト表示|印刷の向き|用紙の向き|印刷範囲|余白を|用紙をA|用紙サイズ|ページ番号を|横向きに印刷|縦向きに印刷|横向きで印刷|縦向きで印刷|設定を横向き|設定を縦向き|表示倍率を|マス目を非表示|マス目を消|方眼を消|方眼を非表示
    ' 依頼の組: 印刷,用紙,余白,ページ,ヘッダー,フッター,目盛線,目盛り線,枠線,グリッド線,改ページ,表示倍率,倍率,ズーム,枠の固定,固定を,画面,行番号,列番号,数式バー,表示を+横向き,縦向き,横に,縦に,A4,A3,B4,B5,Ａ４,Ａ３,狭く,広く,標準,中央に,真ん中に,ページ番号,日付,シート名,ファイル名,100%,１００％,拡大,縮小,合わせ,解除,外して,外す,消して,隠して,非表示,表示して,出して,白黒,プレビュー,範囲を,横幅,縦を+-一覧,調べ,何ページ,何枚,どこで,どこ,確認,PDF,ＰＤＦ,保存,グラフ,ピボット,全シート,すべてのシート,全部のシート,空欄,空白,1ページに収,１ページに収,1枚に収,毎ページ,各ページ,全ページ,文字を,セル,罫線,教え,以上,以下,未満,超え,だけを,レコード,収ま
    ' 階級: 基本
    ' 扱う: 枠の固定 目盛線 表示倍率 印刷の向き 用紙 余白 印刷範囲 ヘッダー フッター
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文で前のシートの表示と印刷の設定を変える（2026-10-28 リボンの目録の 表示・ページレイアウト から）:
    ' 枠の固定を解除・目盛線を消す／出す・行列番号・表示倍率・標準／改ページプレビュー／ページレイアウト・印刷の向き・用紙・余白・
    ' ページの中央に印刷・印刷範囲（表だけ／選んだ範囲／解除）・ヘッダーとフッター（ページ番号・日付・シート名・ファイル名）・白黒・目盛線の印刷。
    ' 印刷はしない。1 ページに収める・見出しを毎ページは「縦横1ページに収めて印刷設定」「横で見出し毎ページの印刷設定」
    Dim 頼 As String, ws As Worksheet, 済 As String, n As Double, 見r As Long, 左 As Long, 右 As Long, 末r As Long, 合r As Long
    Dim 消す As Boolean, 出す As Boolean, 用 As Long, 字 As String, 注 As String
    頼 = 基_頼み()
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「目盛線を消して」「印刷を横向きにして」「フッターにページ番号を入れて」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    消す = 含む語(頼, "消|隠|非表示|なくし|無くし|外|オフ|いらな")
    出す = 含む語(頼, "表示して|出して|見せて|戻して|オン|表示する|表示に")

    ' 画面の表示
    If 含む語(頼, "枠の固定|固定を") And 含む語(頼, "解除|外|やめ|なくし") Then
        ActiveWindow.FreezePanes = False: 済 = 済 & "・枠の固定を解除"
    End If
    If 含む語(頼, "目盛線|目盛り線|枠線|グリッド線|マス目|方眼") And Not 含む語(頼, "印刷") Then
        If 消す Then
            ActiveWindow.DisplayGridlines = False: 済 = 済 & "・目盛線を消す"
        ElseIf 出す Then
            ActiveWindow.DisplayGridlines = True: 済 = 済 & "・目盛線を出す"
        End If
    End If
    If 含む語(頼, "行番号|列番号|行列番号|行と列の番号|行列見出し") Then
        If 消す Then
            ActiveWindow.DisplayHeadings = False: 済 = 済 & "・行列番号を消す"
        ElseIf 出す Then
            ActiveWindow.DisplayHeadings = True: 済 = 済 & "・行列番号を出す"
        End If
    End If
    If InStr(頼, "数式バー") > 0 Then
        If 消す Then
            Application.DisplayFormulaBar = False: 済 = 済 & "・数式バーを消す"
        ElseIf 出す Then
            Application.DisplayFormulaBar = True: 済 = 済 & "・数式バーを出す"
        End If
    End If
    If 含む語(頼, "表示倍率|倍率|ズーム") Then
        n = 基_数(頼)
        If n >= 10 And n <= 400 Then
            ActiveWindow.Zoom = n: 済 = 済 & "・表示倍率を " & n & "%"
        ElseIf 含む語(頼, "選んだ範囲に合わせ|選択範囲に合わせ") And TypeName(Selection) = "Range" Then
            ActiveWindow.Zoom = True: 済 = 済 & "・表示倍率を選んだ範囲に合わせる"
        ElseIf 含む語(頼, "拡大|大きく") Then
            ActiveWindow.Zoom = 基_範囲内(ActiveWindow.Zoom + 25, 10, 400): 済 = 済 & "・表示倍率を " & ActiveWindow.Zoom & "%"
        ElseIf 含む語(頼, "縮小|小さく") Then
            ActiveWindow.Zoom = 基_範囲内(ActiveWindow.Zoom - 25, 10, 400): 済 = 済 & "・表示倍率を " & ActiveWindow.Zoom & "%"
        End If
    End If
    If InStr(頼, "改ページプレビュー") > 0 Then
        ActiveWindow.View = xlPageBreakPreview: 済 = 済 & "・改ページプレビュー"
    ElseIf InStr(頼, "ページレイアウト") > 0 Then
        ActiveWindow.View = xlPageLayoutView: 済 = 済 & "・ページレイアウト表示"
    ElseIf 含む語(頼, "標準表示|標準の表示|表示を標準|標準に戻") Then
        ActiveWindow.View = xlNormalView: 済 = 済 & "・標準の表示"
    End If

    ' 印刷の設定（印刷はしない）。プリンターが無いと PageSetup が読めないので、読めなければ画面の表示だけにする
    If Not 含む語(頼, "印刷|用紙|余白|ページ番号|ヘッダー|フッター|横向き|縦向き|横で|縦で|横長|縦長|白黒|A3|A4|A5|B4|B5|シート名|ファイル名|ブック名|日付") Then GoTo 知らせ
    On Error Resume Next
    用 = ws.PageSetup.Orientation
    If Err.Number <> 0 Then
        On Error GoTo 0
        注 = "（印刷の設定はプリンターが無く読めないので変えていません）"
        GoTo 知らせ
    End If
    Application.PrintCommunication = False
    On Error GoTo 0
    With ws.PageSetup
        If 含む語(頼, "横向き|横で印刷|を横に|横長") Then
            .Orientation = xlLandscape: 済 = 済 & "・横向き"
        ElseIf 含む語(頼, "縦向き|縦で印刷|を縦に|縦長") Then
            .Orientation = xlPortrait: 済 = 済 & "・縦向き"
        End If
        用 = 0
        If InStr(頼, "A3") > 0 Then 用 = xlPaperA3: 字 = "A3"
        If InStr(頼, "A4") > 0 Then 用 = xlPaperA4: 字 = "A4"
        If InStr(頼, "A5") > 0 Then 用 = xlPaperA5: 字 = "A5"
        If InStr(頼, "B4") > 0 Then 用 = xlPaperB4: 字 = "B4"
        If InStr(頼, "B5") > 0 Then 用 = xlPaperB5: 字 = "B5"
        If 用 > 0 Then
            On Error Resume Next             ' 使っているプリンターに無い用紙は選べない
            .PaperSize = 用
            If Err.Number = 0 Then 済 = 済 & "・用紙 " & 字 Else 済 = 済 & "・用紙 " & 字 & "（プリンターに無いので変えられませんでした）"
            On Error GoTo 0
        End If
        If InStr(頼, "余白") > 0 Then
            If 含む語(頼, "狭く|狭い|狭め|小さく|少なく") Then
                基_余白 ws, 1.91, 1.91, 0.64, 0.64, 0.76: 済 = 済 & "・余白を狭く"
            ElseIf 含む語(頼, "広く|広い|広め|大きく|多く") Then
                基_余白 ws, 2.54, 2.54, 2.54, 2.54, 1.27: 済 = 済 & "・余白を広く"
            ElseIf 含む語(頼, "標準|元に戻|普通") Then
                基_余白 ws, 1.91, 1.91, 1.78, 1.78, 0.76: 済 = 済 & "・余白を標準"
            End If
        End If
        If 含む語(頼, "ページの中央|中央に印刷|真ん中に印刷|中央に寄せて印刷|左右の中央") Then
            .CenterHorizontally = True: 済 = 済 & "・左右の中央に印刷"
            If 含む語(頼, "上下も|上下の中央|縦も") Then .CenterVertically = True: 済 = 済 & "・上下の中央に印刷"
        End If
        If InStr(頼, "印刷範囲") > 0 Then
            If 含む語(頼, "解除|クリア|なくし|消|外") Then
                .PrintArea = "": 済 = 済 & "・印刷範囲を解除"
            ElseIf 含む語(頼, "選んだ範囲|選択範囲|選んだ所|ここだけ") And TypeName(Selection) = "Range" Then
                .PrintArea = Selection.Address: 済 = 済 & "・印刷範囲を " & Selection.Address(False, False)
            ElseIf InStr(頼, "表") > 0 And 基_表を読む(ws, 見r, 左, 右, 末r, 合r) Then
                If 合r > 末r Then 末r = 合r
                .PrintArea = ws.Range(ws.Cells(見r, 左), ws.Cells(末r, 右)).Address
                済 = 済 & "・印刷範囲を " & ws.Range(ws.Cells(見r, 左), ws.Cells(末r, 右)).Address(False, False)
            End If
        End If
        If 含む語(頼, "ヘッダー|フッター") And 消す And Not 含む語(頼, "ページ番号|日付|シート名|ファイル名|ブック名") Then
            .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
            .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
            済 = 済 & "・ヘッダーとフッターを消す"
        Else
            If InStr(頼, "ページ番号") > 0 Then
                If InStr(頼, "総ページ") > 0 Or InStr(頼, "分の") > 0 Then 字 = "&P / &N" Else 字 = "&P"
                基_頭と足 ws, 頼, "フッター", "中央", 字: 済 = 済 & "・ページ番号"
            End If
            If 含む語(頼, "日付|今日") Then 基_頭と足 ws, 頼, "ヘッダー", "右", "&D": 済 = 済 & "・印刷した日付"
            If InStr(頼, "シート名") > 0 Then 基_頭と足 ws, 頼, "ヘッダー", "中央", "&A": 済 = 済 & "・シート名"
            If 含む語(頼, "ファイル名|ブック名") Then 基_頭と足 ws, 頼, "ヘッダー", "左", "&F": 済 = 済 & "・ファイル名"
        End If
        If InStr(頼, "白黒") > 0 Then .BlackAndWhite = Not 消す: 済 = 済 & IIf(消す, "・白黒印刷をやめる", "・白黒印刷")
        If 含む語(頼, "目盛線|目盛り線|枠線|グリッド線") And InStr(頼, "印刷") > 0 Then
            .PrintGridlines = Not 消す: 済 = 済 & IIf(消す, "・目盛線を印刷しない", "・目盛線も印刷")
        End If
    End With
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
知らせ:
    If Len(済) = 0 Then
        If Len(注) > 0 Then 先撃ちに断る 注: Exit Sub
        先撃ちに断る "どの表示・印刷の設定を変えるか分かりませんでした（例: 枠の固定を解除・目盛線を消す・表示倍率を100%・印刷を横向き・用紙をA3・余白を狭く・印刷範囲を表だけ・フッターにページ番号）"
        Exit Sub
    End If
    Application.StatusBar = ws.Name & " の表示と印刷の設定を変えました: " & Mid$(済, 2) & 注
End Sub

Private Function 基_余白(ByVal ws As Worksheet, ByVal 上 As Double, ByVal 下 As Double, ByVal 左 As Double, ByVal 右 As Double, ByVal 頭 As Double) As Boolean
    ' 余白（cm）。リボンの「余白」の 狭い・標準・広い と同じ値
    With ws.PageSetup
        .TopMargin = Application.CentimetersToPoints(上)
        .BottomMargin = Application.CentimetersToPoints(下)
        .LeftMargin = Application.CentimetersToPoints(左)
        .RightMargin = Application.CentimetersToPoints(右)
        .HeaderMargin = Application.CentimetersToPoints(頭)
        .FooterMargin = Application.CentimetersToPoints(頭)
    End With
End Function

Private Function 基_頭と足(ByVal ws As Worksheet, ByVal 頼 As String, ByVal 既定の所 As String, ByVal 既定の側 As String, ByVal 字 As String) As Boolean
    ' ヘッダー・フッターの左・中央・右のどこかに字を入れる。頼みの文に ヘッダー／フッター・右／左／中央 があればそこ
    Dim 所 As String, 側 As String
    所 = 既定の所: 側 = 既定の側
    If 含む語(頼, "ヘッダー|上に") And Not 含む語(頼, "フッター") Then 所 = "ヘッダー"
    If 含む語(頼, "フッター|下に") And Not 含む語(頼, "ヘッダー") Then 所 = "フッター"
    If 含む語(頼, "右に|右側|右上|右下") Then 側 = "右"
    If 含む語(頼, "左に|左側|左上|左下") Then 側 = "左"
    If 含む語(頼, "中央に|真ん中に") Then 側 = "中央"
    With ws.PageSetup
        Select Case 所 & 側
            Case "ヘッダー左": .LeftHeader = 字
            Case "ヘッダー中央": .CenterHeader = 字
            Case "ヘッダー右": .RightHeader = 字
            Case "フッター左": .LeftFooter = 字
            Case "フッター中央": .CenterFooter = 字
            Case "フッター右": .RightFooter = 字
        End Select
    End With
End Function

Sub 頼みの文でシートの名前と色と並びを変える()
    ' 依頼の語: シート名を|シートの名前を|シート見出しの色|シートの見出しの色|タブの色|シートを複製|シートをコピー|シートを名前順|シートを並べ替え|シートを保護|シートの保護|隠したシートを|非表示のシートを|シートを再表示|新しいシートを|ワークシートを全部再表示|ワークシートを再表示|あいうえお順|五十音順|タブを並べ
    ' 依頼の組: シート名,シートの名前,シート見出し,シートの見出し,シートタブ,シートのタブ,タブ,シートを,シートの並び,シートの順+変え,変更,名前を,赤,青,緑,黄,紫,オレンジ,黒,灰色,グレー,ピンク,水色,色を,隠して,隠す,非表示,再表示,表示して,コピー,複製,先頭に,末尾に,最後に,一番前,一番後ろ,一番左,一番右,名前順,並べ,保護,追加,新しいシート+-一覧,調べ,確認,PDF,ＰＤＦ,CSV,保存,別ブック,印刷,分け,分割,集約,まとめ,つな,整え,きれい,比較,突合,グラフ,ピボット,表を,値に,式を,削除,教え
    ' 階級: 基本
    ' 扱う: シート名 シート見出しの色 シートの非表示 シートの複製 シートの並び シートの保護
    ' 見出し: なし
    ' 形: なし
    ' 列は頼みから: あり
    ' 頼みの文で前のシートの名前・シート見出しの色・表示・並び・保護を変える（2026-10-28 リボンの目録の ホーム＞セル＞書式・校閲＞保護 から）:
    ' 「シート名を売上にして」「シート見出しを赤にして」「このシートを隠して」「隠したシートを全部表示して」「このシートをコピーして」
    ' 「シートを名前順に並べて」「このシートを先頭に移して」「シートを保護して」「新しいシートを追加して」。シートは消さない（消すのは人が押す）
    Dim 頼 As String, 元 As String, ws As Worksheet, wb As Workbook, 済 As String, 名 As String, 色 As String, sh As Object
    Dim i As Long, j As Long, n As Long, 新 As Worksheet, 見える As Long
    元 = Trim$(先撃ちの頼みの文())
    頼 = ピボットの語ならし(元)
    If Len(頼) = 0 Then 先撃ちに断る "この棚は頼みの文で動きます（例:「シート名を売上にして」「シート見出しを赤にして」「このシートを隠して」）": Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then 先撃ちに断る "ワークシートを前に出してから頼んでください": Exit Sub
    Set ws = ActiveSheet
    Set wb = ws.Parent

    ' 隠したシートを全部表示する（「非表示のシートを表示して」）
    If 含む語(頼, "隠したシート|隠れたシート|隠れているシート|非表示のシート|シートを再表示|シートを全部表示|ワークシートを全部再表示|ワークシートを再表示|になっているワークシート|になっているシート") And _
       含む語(頼, "表示|出して|戻して|見せて") Then
        For Each sh In wb.Sheets
            If sh.Visible = xlSheetHidden Then sh.Visible = xlSheetVisible: n = n + 1
        Next sh
        Application.StatusBar = "隠していたシートを " & n & " 枚表示しました（VBA で隠したシートはそのまま）"
        Exit Sub
    End If

    ' 保護
    If InStr(頼, "保護") > 0 Then
        If 含む語(頼, "解除|外|はず|やめ|とって|取って") Then
            On Error Resume Next
            ws.Unprotect
            If Err.Number <> 0 Then
                On Error GoTo 0
                先撃ちに断る ws.Name & " はパスワード付きで保護されているので外せません（校閲＞シート保護の解除 で人がパスワードを入れてください）"
                Exit Sub
            End If
            On Error GoTo 0
            済 = 済 & "・保護を解除"
        Else
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, AllowSorting:=True
            済 = 済 & "・保護（パスワードなし・並べ替えと絞り込みはできる）"
        End If
    End If

    ' 新しいシートを足す・複製する（名前の頼みがあれば、できたシートの名前にする）
    If 含む語(頼, "新しいシート|シートを追加|シートを足|シートを1枚|シートを一枚") Then
        Set 新 = wb.Worksheets.Add(after:=ws)
        Set ws = 新
        済 = 済 & "・新しいシートを追加"
    ElseIf 含む語(頼, "コピー|複製") Then
        ws.Copy after:=ws
        Set ws = ActiveSheet
        済 = 済 & "・複製（" & ws.Name & "）"
    End If

    ' 名前
    If 含む語(頼, "シート名|シートの名前|名前を|名前は|という名前|の名前で") Then
        名 = 基_シートの新しい名(元)
        If Len(名) > 0 Then
            名 = シート名にできる字(名)
            Set sh = Nothing
            On Error Resume Next
            Set sh = wb.Sheets(名)
            On Error GoTo 0
            If Not sh Is Nothing Then
                If Not sh Is ws Then 先撃ちに断る "「" & 名 & "」という名前のシートがもうあります（別の名前で頼んでください）": Exit Sub
            Else
                ws.Name = 名
                済 = 済 & "・名前を「" & 名 & "」"
            End If
        End If
    End If

    ' シート見出しの色（「シート見出し」「タブ」と言ったときだけ。「このシートを青に」は表の色かもしれないので変えない）
    色 = 基_色の名(頼)
    If 含む語(頼, "シート見出し|シートの見出し|タブ") Then
        If Len(色) > 0 Then
            ws.Tab.Color = 基_色(色, False): 済 = 済 & "・シート見出しを" & 色
        ElseIf 基_やめか(頼, "色") Then
            ws.Tab.ColorIndex = xlColorIndexNone: 済 = 済 & "・シート見出しの色なし"
        End If
    ElseIf Len(色) > 0 And Len(済) = 0 Then
        先撃ちに断る "シート見出しの色のことなら「シート見出しを" & 色 & "にして」と頼んでください（表の色なら「見出しを" & 色 & "にして」）"
        Exit Sub
    End If

    ' 並び
    If 含む語(頼, "名前順|あいうえお順|五十音順|アルファベット順") Or 含む語(頼, "シートを並べ替え|シートを並び替え|シートの順を|タブを並べ|シートタブを並") Then
        For i = 1 To wb.Sheets.Count - 1
            For j = i + 1 To wb.Sheets.Count
                If (StrComp(wb.Sheets(j).Name, wb.Sheets(i).Name, vbTextCompare) < 0) Xor 含む語(頼, "逆順|降順|逆に") Then
                    wb.Sheets(j).Move Before:=wb.Sheets(i)
                End If
            Next j
        Next i
        ws.Activate
        済 = 済 & "・シートを名前の" & IIf(含む語(頼, "逆順|降順|逆に"), "逆順", "順") & "に並べ"
    ElseIf 含む語(頼, "先頭|一番前|一番左|最初に|はじめに") Then
        ws.Move Before:=wb.Sheets(1): 済 = 済 & "・先頭へ"
    ElseIf 含む語(頼, "末尾|一番後|一番右|最後に|最後へ|終わりに") Then
        ws.Move after:=wb.Sheets(wb.Sheets.Count): 済 = 済 & "・末尾へ"
    End If

    ' 隠す（最後に。見えるシートが 1 枚だけなら隠せない）
    If 含む語(頼, "隠|非表示") And Not 含む語(頼, "隠した|隠れた|隠れている|非表示の|再表示") Then
        For Each sh In wb.Sheets
            If sh.Visible = xlSheetVisible Then 見える = 見える + 1
        Next sh
        If 見える <= 1 Then 先撃ちに断る "見えるシートが 1 枚だけなので隠せません": Exit Sub
        ws.Visible = xlSheetHidden
        済 = 済 & "・隠す"
    End If

    If Len(済) = 0 Then
        先撃ちに断る "シートをどう変えるか分かりませんでした（例: シート名を売上にして・シート見出しを赤に・隠す・隠したシートを表示・コピー・名前順に並べる・先頭へ・保護）"
        Exit Sub
    End If
    Application.StatusBar = "シート「" & ws.Name & "」を変えました: " & Mid$(済, 2)
End Sub


Private Function 基_シートの新しい名(ByVal 元 As String) As String
    ' 頼みの文からシートの新しい名前を読む（「」『』“”の中を先に。無ければ「名前を〇〇に」「シート名を〇〇にして」の〇〇）。読めなければ ""
    Dim 開 As Variant, 閉 As Variant, i As Long, p As Long, q As Long, s As String, w As Variant, E As Long
    開 = Array("「", "『", ChrW(&H201C), """")    ' “” は VBA の中では引用符に読まれるので ChrW で
    閉 = Array("」", "』", ChrW(&H201D), """")
    For i = 0 To UBound(開)
        p = InStr(1, 元, 開(i), vbBinaryCompare)
        If p > 0 Then
            q = InStr(p + 1, 元, 閉(i), vbBinaryCompare)
            If q > p + 1 Then 基_シートの新しい名 = Mid$(元, p + 1, q - p - 1): Exit Function
        End If
    Next i
    For Each w In Array("シート名を", "シートの名前を", "名前を", "シート名は", "名前は")
        p = InStr(1, 元, w, vbBinaryCompare)
        If p > 0 Then s = Mid$(元, p + Len(w)): Exit For
    Next w
    If Len(s) = 0 Then Exit Function
    E = Len(s) + 1
    For Each w In Array("に変え", "に変更", "にして", "にする", "にし", "へ変え", "と付け", "とつけ", "に直", "にかえ", "に", "、", "。")
        q = InStr(1, s, w, vbBinaryCompare)
        If q > 1 And q < E Then E = q
    Next w
    基_シートの新しい名 = Trim$(Left$(s, E - 1))
End Function

' ---- 6 本の共通（Private＝棚の目録に出ない） ----
Private Function 基_頼み() As String
    ' 頼みの文（全角の英数を半角に、空白と括弧の中を落とす＝見出しの語と同じならし）
    基_頼み = ピボットの語ならし(先撃ちの頼みの文())
End Function

Private Function 基_表を読む(ByVal ws As Worksheet, ByRef 見r As Long, ByRef 左 As Long, ByRef 右 As Long, _
                            ByRef 末r As Long, ByRef 合r As Long) As Boolean
    ' 前のシートの表の 見出しの行・左右の列・本文の最後の行・合計の行（無ければ 0）。テーブルならテーブルの見出しと集計行
    Dim lo As ListObject, ur As Range, r As Long, c As Long, filled As Long, txt As Long, v As Variant
    見r = 0: 左 = 0: 右 = 0: 末r = 0: 合r = 0
    Set lo = 基_表のテーブル(ws)
    If Not lo Is Nothing Then
        If lo.ShowHeaders Then
            見r = lo.HeaderRowRange.row
            左 = lo.Range.Column
            右 = 左 + lo.Range.Columns.Count - 1
            If lo.DataBodyRange Is Nothing Then
                末r = 見r
            Else
                末r = lo.DataBodyRange.row + lo.DataBodyRange.rows.Count - 1
            End If
            If lo.ShowTotals Then 合r = lo.TotalsRowRange.row
            基_表を読む = True
            Exit Function
        End If
    End If
    Set ur = ws.UsedRange
    For r = ur.row To ur.row + ur.rows.Count - 1
        filled = 0: txt = 0
        For c = ur.Column To ur.Column + ur.Columns.Count - 1
            v = ws.Cells(r, c).Value
            If セルの字(v) <> "" Then
                filled = filled + 1
                If VarType(v) = vbString And Not IsNumeric(v) Then txt = txt + 1
            End If
        Next c
        If filled >= 2 And txt * 2 > filled Then 見r = r: Exit For
    Next r
    If 見r = 0 Then Exit Function
    For c = ur.Column To ur.Column + ur.Columns.Count - 1
        If 見出しの字(ws.Cells(見r, c)) <> "" Then
            If 左 = 0 Then 左 = c
            右 = c
        ElseIf 左 > 0 Then
            Exit For
        End If
    Next c
    末r = 表の本文の最終行(ws)
    If 末r < 見r Then 末r = 見r
    合r = 表の下の合計行(ws, 末r)
    If 合r = 0 Then
        For r = 末r + 1 To 末r + 2
            If 集計の行か(ws, r, 左, 右) Then 合r = r: Exit For
        Next r
    End If
    基_表を読む = True
End Function


Private Function 基_表のテーブル(ByVal ws As Worksheet) As ListObject
    ' 選んでいるセルのテーブル。選んでいる所がテーブルの外で、シートにテーブルが 1 つだけならそれ
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing And ws.ListObjects.Count = 1 Then Set lo = ws.ListObjects(1)
    Set 基_表のテーブル = lo
End Function

Private Function 基_場所の語を外す(ByVal 頼 As String) As String
    ' 見出しの名前を探す前に、場所を指す語（見出し・合計行・選んだ範囲…）を外す（「合計」の列が「合計行」で当たらないように）
    Dim w As Variant
    For Each w In Array("見出しの行", "見出し行", "見出し", "ヘッダー", "項目名", "項目の行", "合計行", "合計の行", "総計の行", "小計の行", _
                        "選んだ範囲", "選んだセル", "選んだ列", "選んだ行", "選んだ所", "選択範囲", "選択した", "この範囲", "このセル", "この列", "この行", _
                        "全部の列", "すべての列", "全ての列", "表の列", "全部の行", "すべての行", "表の行")
        頼 = Replace(頼, w, "|")
    Next w
    基_場所の語を外す = 頼
End Function


Private Function 基_頼みの列(ByVal ws As Worksheet, ByVal 頼 As String, ByVal 見r As Long, ByVal 左 As Long, ByVal 右 As Long) As Collection
    ' 頼みの文に名前の出てくる見出しの列番号。字のとおりに出てくる見出しを先に見て（長い見出しに含まれる短い見出しは外す）、
    ' 1 つも無ければ言い換え（金額→売上 など・見出しの語の位置）で探す
    Dim out As New Collection, c As Long, j As Long, h As String, p As Long, n As Long, 残す As Boolean
    Dim 位() As Long, 長() As Long
    Set 基_頼みの列 = out
    If 見r = 0 Or 左 = 0 Then Exit Function
    頼 = 基_場所の語を外す(頼)
    ReDim 位(左 To 右): ReDim 長(左 To 右)
    For c = 左 To 右
        h = ピボットの語ならし(見出しの字(ws.Cells(見r, c)))
        If Len(h) >= 2 Or (Len(h) = 1 And InStr(1, 頼, h & "の列", vbBinaryCompare) > 0) Then
            p = InStr(1, 頼, h, vbBinaryCompare)
            If p > 0 Then 位(c) = p: 長(c) = Len(h): n = n + 1
        End If
    Next c
    If n > 0 Then
        For c = 左 To 右
            If 長(c) > 0 Then
                残す = True
                For j = 左 To 右
                    If j <> c And 長(j) > 長(c) Then
                        If 位(j) <= 位(c) And 位(c) + 長(c) <= 位(j) + 長(j) Then 残す = False: Exit For
                    End If
                Next j
                If 残す Then out.Add c
            End If
        Next c
    Else
        For c = 左 To 右
            h = 見出しの字(ws.Cells(見r, c))
            If Len(h) > 0 Then
                If 見出しの語の位置(頼, h) > 0 Then out.Add c
            End If
        Next c
    End If
End Function

Private Function 基_足す(ByVal a As Range, ByVal b As Range) As Range
    If a Is Nothing Then
        Set 基_足す = b
    ElseIf b Is Nothing Then
        Set 基_足す = a
    Else
        Set 基_足す = Union(a, b)
    End If
End Function

Private Function 基_指した所(ByVal ws As Worksheet, ByVal 頼 As String, ByVal 見出しも As Boolean, _
                            ByRef 何 As String, ByRef なぜ As String) As Range
    ' 頼みの文の「どこを」→ セルの範囲。見出し（見出しもが True のとき）・合計行・選んだ範囲（この列・この行）・見出しの名前で指した列。
    ' 列と見出し・合計行が両方あれば、その列の見出し・合計のセル（「金額の見出し」「合計行の金額」）。何＝知らせに出す場所の名前
    Dim 見r As Long, 左 As Long, 右 As Long, 末r As Long, 合r As Long, 下 As Long, 表あり As Boolean
    Dim cols As Collection, c As Variant, out As Range, 見出し As Boolean, 合計 As Boolean, 列の印 As Boolean, 両方 As Boolean
    何 = "": なぜ = ""
    見出し = 見出しも And 含む語(頼, "見出し|ヘッダー|項目名|項目の行") And Not 含む語(頼, "シート見出し|シートの見出し")
    合計 = 含む語(頼, "合計行|合計の行|総計の行|小計の行")
    列の印 = 含む語(頼, "の列|列を|列の|列は|列が|列で|の欄")
    両方 = 含む語(頼, "見出しと|と見出し|合計行と|と合計行")
    表あり = 基_表を読む(ws, 見r, 左, 右, 末r, 合r)
    If 表あり Then
        下 = 末r
        If 合r > 下 Then 下 = 合r
        Set cols = 基_頼みの列(ws, 頼, 見r, 左, 右)
        If 合計 And 合r = 0 Then なぜ = "表の下に合計の行が見つかりません（「合計」「総計」「計」と書いた行）": Exit Function
        If cols.Count > 0 And (見出し Or 合計) And Not 両方 Then
            For Each c In cols
                If 見出し Then Set out = 基_足す(out, ws.Cells(見r, c))
                If 合計 Then Set out = 基_足す(out, ws.Cells(合r, c))
                何 = 何 & "・" & 見出しの字(ws.Cells(見r, c))
            Next c
            何 = Mid$(何, 2) & IIf(見出し, "の見出し", "の合計")
        Else
            If 見出し Then Set out = 基_足す(out, ws.Range(ws.Cells(見r, 左), ws.Cells(見r, 右))): 何 = 何 & "・見出し"
            If 合計 Then Set out = 基_足す(out, ws.Range(ws.Cells(合r, 左), ws.Cells(合r, 右))): 何 = 何 & "・合計行"
            If cols.Count > 0 And 列の印 And 下 > 見r Then
                For Each c In cols
                    Set out = 基_足す(out, ws.Range(ws.Cells(見r + 1, c), ws.Cells(下, c)))
                    何 = 何 & "・" & 見出しの字(ws.Cells(見r, c)) & "の列"
                Next c
            End If
            何 = Mid$(何, 2)
        End If
    End If
    If 含む語(頼, "選んだ|選択|この範囲|このセル|この列|この行") And TypeName(Selection) = "Range" Then
        If InStr(頼, "この列") > 0 And 表あり Then
            Set out = 基_足す(out, Intersect(Selection.EntireColumn, ws.Range(ws.Cells(見r, 左), ws.Cells(下, 右))))
        ElseIf InStr(頼, "この行") > 0 And 表あり Then
            Set out = 基_足す(out, Intersect(Selection.EntireRow, ws.Range(ws.Cells(見r, 左), ws.Cells(下, 右))))
        Else
            Set out = 基_足す(out, Selection)
        End If
        何 = IIf(Len(何) > 0, 何 & "・", "") & "選んだ範囲"
    End If
    If out Is Nothing Then
        If Not 表あり Then
            なぜ = "表の見出しの行が見つかりません（見出しが 1 行ある表で頼むか、範囲を選んで「選んだ範囲を…」と頼んでください）"
        ElseIf 列の印 Then
            なぜ = "頼みの文にある列の名前が表の見出しに見つかりません（見出しの名前のとおりに「金額の列を…」と頼んでください）"
        Else
            なぜ = "どこを変えるか分かりませんでした（見出し・合計行・選んだ範囲・「金額の列」のように見出しの名前で頼んでください）"
        End If
    End If
    Set 基_指した所 = out
End Function


Private Function 基_節に分ける(ByVal 頼 As String) As Variant
    ' 「見出しの文字を大きくして赤にして」→ 大きく／赤に。色を文字と背景のどちらに付けるかを節ごとに読む
    Dim w As Variant
    For Each w In Array("して", "、", ",", "。", "そして", "それから")
        頼 = Replace(頼, w, vbLf)
    Next w
    基_節に分ける = Split(頼, vbLf)
End Function

Private Function 基_やめか(ByVal s As String, ByVal 語 As String) As Boolean
    ' 「太字を外して」「下線をやめて」「折り返しなし」のような、付いている物を外す頼みか
    基_やめか = 含む語(s, 語 & "を外|" & 語 & "をはず|" & 語 & "をやめ|" & 語 & "を解除|" & 語 & "を取|" & 語 & "を消|" & _
                        語 & "をなくし|" & 語 & "を無くし|" & 語 & "なし|" & 語 & "はいらな|" & 語 & "をしな")
End Function


Private Function 基_色の名(ByVal s As String) As String
    ' 節の中の色の名前（長い名前を先に見る）。無ければ ""
    Dim w As Variant
    s = Replace(Replace(s, "淡い", "薄い"), "うすい", "薄い")
    For Each w In Array("薄い赤", "薄い青", "薄い緑", "薄い黄", "薄い灰", "薄いグレー", "薄いオレンジ", "薄い紫", "薄いピンク", _
                        "濃い赤", "濃い青", "濃い緑", "水色", "ピンク", "紺", "ネイビー", "オレンジ", "紫", "茶", _
                        "赤", "青", "緑", "黄", "黒", "白", "灰", "グレー")
        If InStr(1, s, w, vbBinaryCompare) > 0 Then 基_色の名 = w: Exit Function
    Next w
End Function


Private Function 基_色(ByVal 名 As String, ByVal 塗り As Boolean) As Long
    ' 色の名前 → RGB。リボンの「標準の色」に合わせ、薄い色は背景に使う淡い色。灰色の背景は文字が読める明るさにする
    Select Case 名
        Case "薄い赤": 基_色 = RGB(255, 199, 206)
        Case "薄い青": 基_色 = RGB(221, 235, 247)
        Case "薄い緑": 基_色 = RGB(226, 239, 218)
        Case "薄い黄": 基_色 = RGB(255, 242, 204)
        Case "薄い灰", "薄いグレー": 基_色 = RGB(242, 242, 242)
        Case "薄いオレンジ": 基_色 = RGB(252, 228, 214)
        Case "薄い紫": 基_色 = RGB(228, 223, 236)
        Case "薄いピンク": 基_色 = RGB(255, 228, 240)
        Case "濃い赤": 基_色 = RGB(192, 0, 0)
        Case "濃い青", "紺", "ネイビー": 基_色 = RGB(0, 32, 96)
        Case "濃い緑": 基_色 = RGB(0, 97, 0)
        Case "水色": 基_色 = RGB(0, 176, 240)
        Case "ピンク": 基_色 = RGB(255, 153, 204)
        Case "オレンジ": 基_色 = RGB(255, 192, 0)
        Case "紫": 基_色 = RGB(112, 48, 160)
        Case "茶": 基_色 = RGB(153, 102, 51)
        Case "赤": 基_色 = RGB(255, 0, 0)
        Case "青": 基_色 = RGB(0, 112, 192)
        Case "緑": 基_色 = RGB(0, 176, 80)
        Case "黄": 基_色 = RGB(255, 255, 0)
        Case "黒": 基_色 = RGB(0, 0, 0)
        Case "白": 基_色 = RGB(255, 255, 255)
        Case "灰", "グレー": 基_色 = IIf(塗り, RGB(217, 217, 217), RGB(128, 128, 128))
    End Select
End Function

Private Function 基_数(ByVal s As String, Optional ByVal 後 As String = "") As Double
    ' 文の中の最初の数（後＝その語の後ろから探す）。無ければ -1
    Dim i As Long, t As String, ch As String, p As Long
    基_数 = -1
    If Len(後) > 0 Then
        p = InStr(1, s, 後, vbBinaryCompare)
        If p = 0 Then Exit Function
        s = Mid$(s, p + Len(後))
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Or (ch = "." And Len(t) > 0) Then
            t = t & ch
        ElseIf Len(t) > 0 Then
            Exit For
        End If
    Next i
    If Len(t) > 0 Then 基_数 = val(t)
End Function
