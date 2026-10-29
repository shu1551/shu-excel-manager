Attribute VB_Name = "表の整理_書式"

Sub 選んだ範囲の文字を中央揃えにする()
    ' 依頼の語: 中央揃え|センタリング|真ん中に寄せ|中央揃えにして|中央ぞろえ|真ん中揃え|中央寄せ
    ' 依頼の組: 中央,センタリング,真ん中+揃え,寄せ,配置+-ピボット,グラフ,縦,の列,見出し,ヘッダー,合計行,合計の行
    ' 扱う: 中央揃え センタリング 水平配置 揃え
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub

Sub 選んだ範囲の文字を左寄せにする()
    ' 依頼の語: 左寄せ|左揃え|左詰め|左に寄せ|左揃えにして|左寄せにして
    ' 依頼の組: 左+揃え,寄せ,詰め,配置+-ピボット,グラフ,の列,見出し,ヘッダー,合計行,合計の行
    ' 扱う: 左寄せ 左揃え 左詰め 水平配置
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.HorizontalAlignment = xlLeft
    Application.ScreenUpdating = True
End Sub

Sub 選んだ範囲の文字を右寄せにする()
    ' 依頼の語: 右寄せ|右揃え|右詰め|右に寄せ|右揃えにして|右寄せにして
    ' 依頼の組: 右+揃え,寄せ,詰め,配置+-ピボット,グラフ,の列,見出し,ヘッダー,合計行,合計の行
    ' 扱う: 右寄せ 右揃え 右詰め 水平配置
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.HorizontalAlignment = xlRight
    Application.ScreenUpdating = True
End Sub

Sub 選んだ列の数値を3桁カンマ区切りにする()
    ' 依頼の語: カンマ区切り|3桁区切り|桁区切り|3桁ごと|カンマを入れて|カンマをつけて|3桁カンマ|数値の桁区切り
    ' 依頼の組: カンマ,3桁,桁+区切り,ごと,入れ,つけ+-分割,クエリ,改行,スペース,の列,合計行,合計の行
    ' 扱う: 3桁カンマ カンマ区切り 桁区切り 表示形式
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.NumberFormatLocal = "#,##0"
    Application.ScreenUpdating = True
End Sub

Sub 選んだ列の数値を円マーク通貨表示にする()
    ' 依頼の語: 通貨表示|円マーク|金額表示|通貨記号|￥マーク|円表示|円をつけて
    ' 依頼の組: 通貨,円,金額,￥+表示,マーク,記号+-外貨,ドル,千円,百万円,万円,単位,順位,の列,合計行,合計の行
    ' 扱う: 通貨表示 円マーク 金額表示 表示形式
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.NumberFormatLocal = "\#,##0;-\#,##0;\0"
    Application.ScreenUpdating = True
End Sub

Sub 選んだ列の数値をパーセント表示にする()
    ' 依頼の語: パーセント表示|％表示|割合表示|パーセントにして|％にして|パーセント表記|％表記
    ' 依頼の組: パーセント,％,割合+表示,表記,にして+-グラフ,円,の列,合計行,合計の行,倍率,ズーム
    ' 扱う: パーセント表示 割合表示 ％ 表示形式
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.NumberFormatLocal = "0.0%"
    Application.ScreenUpdating = True
End Sub

Sub 見出し行を太字にして背景を薄く塗る()
    ' 依頼の語: 見出し強調|見出し行を目立|見出しを目立|太字にして背景|背景を薄く
    ' 依頼の組: 見出し,ヘッダー+強調,目立+-消,削除,pt,ポイント,大きく,小さく,フォント,書体,ゴシック,明朝,メイリオ,赤,青,紺,黄,緑,文字色,文字の色
    ' 扱う: 見出し強調 ヘッダー太字 表の見た目
    ' 見出し: なし
    On Error Resume Next
    Dim ur As Range
    Set ur = ActiveSheet.UsedRange
    If ur Is Nothing Then Exit Sub
    Dim headRng As Range
    Set headRng = ur.rows(1)
    Application.ScreenUpdating = False
    headRng.Font.Bold = True
    headRng.Interior.Color = RGB(242, 244, 247)
    Application.ScreenUpdating = True
End Sub

Sub 選んだ範囲の文字を太字にする()
    ' 依頼の語: 太字にして|ボールドにして|文字を太字|選択範囲を太字|太字適用
    ' 依頼の組: 太字,ボールド+して,適用,変え+-見出し,ヘッダー,消,削除,の列,合計行,合計の行
    ' 扱う: 太字 ボールド 文字強調
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    targetRng.Font.Bold = True
    Application.ScreenUpdating = True
End Sub

Sub 表の外枠を太線で囲み中は細線にする()
    ' 依頼の語: 罫線を細く|外枠を太く|格子罫線|罫線を整えて|格子状に罫線|表の罫線を整えて|外枠太線|外枠を太線|太線で囲
    ' 依頼の組: 罫線,格子,外枠+細く,太く,太線,整え,引いて+-消,削除,波線,引き直
    ' 扱う: 罫線調整 格子 外枠太線 罫線
    ' 見出し: なし
    On Error Resume Next
    Dim targetRng As Range
    If TypeName(Selection) = "Range" And Selection.CountLarge > 1 Then
        Set targetRng = Selection
    Else
        Set targetRng = ActiveSheet.UsedRange
    End If
    If targetRng Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    With targetRng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With targetRng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With targetRng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With targetRng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With targetRng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With targetRng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Application.ScreenUpdating = True
End Sub
