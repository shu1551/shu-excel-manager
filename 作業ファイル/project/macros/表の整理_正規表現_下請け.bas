Attribute VB_Name = "表の整理_正規表現_下請け"
Function 正規表現_標準版を作る() As Object
    ' VBScript.RegExp が使えない環境（2027 年に VBScript が外される）用の代わり。VBA 標準の RegExp（Microsoft 365・2025-08 以降）を返す。
    ' 古い版（RegExp の型が無い）はこのモジュールだけがコンパイルエラーになるので、他の棚に響かないよう独立させ、
    ' VBScript.RegExp が作れなかったときだけ呼ぶ。書き方・使い方は VBScript.RegExp と同じ（Pattern・Global・IgnoreCase・Execute・Replace・Test）。
    Set 正規表現_標準版を作る = New RegExp
End Function
