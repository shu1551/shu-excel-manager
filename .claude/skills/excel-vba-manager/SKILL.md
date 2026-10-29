---
name: excel-vba-manager
description: |
  Excel の VBA マクロ・シート・テーブル・ピボット・チャート・パワークエリ等を
  Python(vba_manager.py / form_builder.py)から操作する汎用スキル。
  対象は常に「今アクティブに開いている Excel ブック（ActiveWorkbook）」で、特定のブック・パス・
  モジュール名に依存しない。秀.xlsm はこのツールの利用例の一つにすぎない。

  以下のときに使う：
  - 「マクロを追加/修正/作って」「VBAマクロ」など Excel VBA の作成・修正
  - vba_manager.py / form_builder.py を使う作業
  - 開いている Excel ブックのプロシージャ/モジュールの取得・置換
  - シート操作・セル編集・テーブル・ピボット・チャート・パワークエリ・データモデル
  - UserForm の作成・修正、ショートカットキー(Attribute行)の操作

  ※「秀」「秀.xlsm」「アドイン」「shu001/003/005」と名指しされた秀.xlsm固有の話（パス・
    モジュール地図・フォーム一覧）は shu-addin-manager スキルを併用する。
---

# Excel VBA マネージャー（汎用）

この SKILL.md は **型と索引だけ**（約 110 行）。細部は `references/` にあり、**要るときだけ 1 ファイルを Read する**
（2026-10-29 にもう一段分けた。それまで 2 万字を発火のたびに読んでいた）。索引に無い引数は `vba_help` を引かず撃つ＝外れれば usage が返る。

| 読む場面 | ファイル |
|---|---|
| 表・シートを触る手の全文（決め直さない既定・棚の鍛え方・帳票→一覧） | `references/sheet.md` |
| コマンド索引の全部（下の「よく使うコマンド」に無い手） | `references/index.md` |
| コマンドの引数・注意（全文） | `references/commands.md` |
| agent（依頼文 1 つを道具が回す）・--recipe・--undo・先撃ち・関所 | `references/agent.md` |
| UserForm（form_inspect・form_tool・form_layout・form_builder） | `references/userform.md` |
| 現地調査・動作検証・標準作業フロー・絶対ルールの経緯・変更前チェック | `references/workflows.md` |

## 対象の原則（最重要・まずこれを当てる）

vba_manager.py / form_builder.py は **アクティブな開いているブック（ActiveWorkbook）** に対して動く。
引数なしなら今アクティブなブック、第 1 引数にパスを渡せばそのファイル。既存コードに `XLSM_PATH`・`shu001` 等の
固有名詞があっても「対象」と鵜呑みにしない。**パス・ブック名を決め打ちしない。**

## 作業の型

1. **道具は MCP が第一選択**（`mcp__excel-manager__vba("コマンド1行")` ほか）。確認を出す手は `-y`。道具の .py を直したら `reload_tools()`。
2. **点検**: list → export-all → コードを読む → 所見を数行。checkup／check は利用者が明示したときだけ。
3. **「遅い」**はまず `stats`（道具の中 vs 間・失敗の理由の多い順）。独立な調査・実行は 1 往復に並列で畳む。

## 手の型（要点）

表・シートの仕事は **往復 4 回が上限**。要点に無い既定・引数は `references/sheet.md`（要るときだけ Read）。
1. 表の修正・点検 → **`seiri 頼みの文`**（頼みの語が無ければ引数なし）。表を整えるマクロ・頼みに当たる棚・tidy まで道具が 1 回で撃つ。残りが「エラーセル: なし」なら報告して終わる（screenshot も手直しも足さない）。
2. 表を見るだけ → **`materials`**（read-range を重ねない）。式の元は `trace D31 --depth 3`、ブックの仕組みは `structure`。
3. 書く → 飛び飛びは `write-cells C7 値 C11 値 --show`、範囲・数式は `write_grid(range, tsv)`。複数の書き込みは同じ返事に並べる。文字列の化けは道具が戻す。
4. 仕上げ → `tidy 表全体`（列を足したら元の列を含む表全体）。自分の手で見た目を変えたときだけ `screenshot` を 1 回。
5. 棚（秀コンボのモジュール「表の整理」「表の整理_作る」「表の整理_調べる」「表の整理_書式」「表の整理_基本操作」）は、自分で書く前に `shelf --ask 依頼文`／`shelf --grep 語` → `shelf-run 名前 --select A1:A9,C1:C9`（選ぶ列は 1 列ずつカンマで）。
6. 10 手を超える組み立てはマクロ 1 本に書いて撃つ（下の「マクロの手の型」）。
7. 報告は番地つきで 3 行程度（5 行以内）。戻すのは `agent --undo`。保存はしない。

## マクロの手の型（要点）

1. **読む**: 長いプロシージャを丸ごと取らない。`get 名前 --around 語`（語の前後だけ）／`get 名前 --lines 40-80`。名前の定義・使う所・呼び元（間接も）は `grep 名前 --where` 1 回（語で少しずつ探さない）。当たったブックは返事の頭の「対象:」で分かる（list-open は要らない）。
2. **直す**: 数行は `patch-procedure モジュール Sub --old "旧" --new "新" -y`（MCP は patch_procedure）。新規は set_procedure_code → `add-procedure モジュール -y`。過半が変わるときだけ replace_procedure（code に全文）。
3. **確かめる**: `compile` と `run-macro 名前`（または `test 絞り込み`）を直しと同じ返事に並べる。検証を独立の往復にしない。
4. **棚・アドイン（秀コンボ）を直したら**: 確かめは写しで済ませる（`gate --changed`＝前に焼いてから変わったモジュールのテストだけ・頭の語は `shelf --words 直し.tsv --try 頼み`）。`register-addin` は全部直してから最後に 1 回（短い間の焼き直しは道具が知らせる）。返事の最後の「次の手:」に従う（止まったら直す所、通ったら公開）。
5. **「動かない」**は `repair マクロ` 1 回で材料を取り、仮説を一言で。直し方は利用者が選ぶ。

## ツールの場所

- SCRIPTS: vba_manager.py 等が置いてあるフォルダ（決め打ち禁止。不明なら Glob `**/vba_manager.py`）。`_last_proc.vba`（get の出力・UTF-8）・`backups\` もここ基準。`py` で実行。SCRIPTS は道具の場所であって操作対象ではない。

## よく使うコマンド（全部は references/index.md）

```bash
# --- VBA を読む・直す ---
list [--detail] [--module 名]                           # マクロ一覧
get <Sub名> [--around 語] [--lines 40-80] | get <モジュール> <Sub名>   # コード取得 → _last_proc.vba（--around/--lines は返す所だけ絞る）
grep "文字" [--regex] [-i] [--module 名] [--where]     # 全モジュール横断の検索（--where＝定義・使う所・呼び元を 1 回で）
patch-procedure <モジュール> <Sub名> --old "旧" --new "新" -y   # 数行の直し（変わった所の前後だけ表示）
replace-procedure -y [--module 名] [--code-file f]    # _last_proc.vba で置換（控え・自動保存）
add-procedure <モジュール> -y / delete-procedure [モジュール] <Sub名> -y
code-replace "旧" "新" [--module 名] -y                 # 全マクロ横断の一括置換
rename-procedure <旧名> <新名> [--module 名] -y        # 呼び元・OnAction・ショートカットまで
repair <マクロ>                                       # 修理の材料を 1 手で
compile / test [絞り込み] / gate [絞り込み] [--changed]   # 全体コンパイル・テスト Sub・関所（写しで check＋compile＋test）
run-macro <マクロ> [引数…] [--input-text 値]           # 実行（実行時エラーが文字で返る）
register-addin [ブック] [-y]                           # .xlam に焼き直して登録（焼く前に棚の検査と gate。最後に 1 回）
stats [--days 7] [--cmd 名]                           # 呼び出し台帳＝回数・秒・失敗の理由
# 表・シートの手は上の「手の型（要点）」、残りは references/index.md
```

## 絶対に守るべきルール（詳細と経緯は references/workflows.md）

- **識別子**: プロシージャ名・モジュール名は英字か日本語で始める（先頭 `_` はコンパイルで死ぬ。道具が拒む。--force で潰さない）。
  注入経路の台帳は test_tools.py の `test_injection_route_ledger`＝新しい注入経路を作ったらガードを配線して登録する。
- **エンコーディング**: .bas は **CP932**。**.bas に Edit／Write ツールを絶対に使うな**。直すなら Python で CP932 のまま読み書きするか、`_last_proc.vba`（UTF-8）経由の replace-procedure。取り込み前は `check-bas`。
- **モジュール適用**: AddFromString は Attribute を壊すことがある＝**replace-module（Remove+Import）**（モジュールは末尾へ移る・ショートカットは道具が再登録）。
- **勝手な変更の禁止**: 指示されていない機能を足さない・影響範囲を確認せずに変えない。
- **破壊的操作の前に現地調査**: シートに触るマクロ・行列削除・クリアの前に `snapshot`／`materials` で実勢を見る。
- **動作検証は COM 経由が第一選択**（run-macro／test／snapshot-diff）。打鍵は最終手段。
- **Excel が詰まったとき**は `unstick` で状況を見て `unstick --fix`。表示中のフォームは `close-form`。

## マクロ作成・修正時の最速原則

- **万単位の実務データ前提**: セル直接ループ・`Union`・`EntireRow.Delete`・`AddItem` 連打を書かない。`Range.Value` で配列に取り、
  メモリで判定し、`Range.Resize(...).Value = 配列` で 1 手で書く。`ScreenUpdating`・`Calculation`・`EnableEvents` を止めて戻す。
- **テスト**: 直したら `compile` と関係するテスト（`test 絞り込み`）を同じ往復で撃って PASS を確かめてから報告する。
