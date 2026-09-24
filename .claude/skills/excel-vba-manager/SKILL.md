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

この SKILL.md は **型と索引**（約 300 行）。引数の細部・経緯・agent の全文は `references/` にあり、
**要るときだけ 1 ファイルを Read する**（2026-09-16 に分割。それまで 1,225 行を発火のたびに読んでいた）。
索引に無い引数は `vba_help` を引かず撃つ＝外れれば usage が返る。

| 読む場面 | ファイル |
|---|---|
| コマンドの引数・注意（全文） | `references/commands.md` |
| agent（依頼文 1 つを道具が回す）・--fire・--shake・--recipe・--undo・関所 | `references/agent.md` |
| UserForm（form_inspect・form_tool・form_layout・form_builder） | `references/userform.md` |
| 現地調査・動作検証・標準作業フロー・絶対ルールの経緯・変更前チェック | `references/workflows.md` |

## 対象の原則（最重要・まずこれを当てる）

vba_manager.py / form_builder.py は **アクティブな開いているブック（ActiveWorkbook）** に対して動く。

- 引数なし → 今アクティブな Excel ブックを自動使用。第 1 引数にパスを渡せばそのファイル。
- だから**任意の Excel ファイルの VBA も、Excel で開いてアクティブにすれば同じ手順で修正できる**。
- 既存コードに `XLSM_PATH`・`shu001`/`shu003` 等の固有名詞があっても「対象」と鵜呑みにしない。
  **固有名詞より「対象はアクティブブック」を先に当てる。** パス・ブック名を決め打ちしない。

## 作業の型（速度既定・2026-07-12 制定）

「時間かかりすぎ」の叱責を受けて制定。**段取りをその場で考え直さず、この型どおり動く。**

1. **道具は MCP が第一選択**（`mcp__excel-manager__vba("コマンド1行")` ほか。サーバーの名前は登録しだい＝古い登録なら `mcp__vba-manager__vba`）。常駐 COM で応答 0.01〜0.2 秒。
   確認プロンプト系は必ず `-y`。CLI に落とすのは ①本体を改修中で新コードを検証するとき ②MCP 不在時 だけ
   （道具の .py を直したら `reload_tools()` で読み直す＝サーバーの再起動は要らない）。
2. **点検の型**: list → export-all → コードを読む → 所見を数行。checkup は利用者が明示したときだけ。
   対象モジュールが分かっているときの list は `--module 名`。
   On Error Resume Next の行・150 行超は `check` の VBM015・016 に含む（clean-vba は 9/17 に check へ畳んだ）。
3. **修正の型（差分の大きさで経路を選ぶ）**: 1〜数行は `code-replace "旧" "新" --module 名 -y` または `patch-procedure モジュール Sub --old "旧" --new "新" -y`（**変更行だけ送る・0.09秒で超高速**）。
   新規は add-procedure。過半が変わる大改造だけ get→replace-procedure。モジュール先頭行は replace-module。
   改名は `rename-procedure`（呼び元・OnAction・ショートカットまで一括）。
4. **修理の型（「動かない」）**: 名指しされたマクロは **`repair <マクロ>` 一発**で取り（本文・入口・呼び元呼び先・
   コンパイルの落ちた行・check の error 級・直前の控えとの差分が 1 画面。get＋call-graph＋compile＋check を撃ち分けない・
   2026-09-17）、同じ返答の中で仮説を一言で言い切る。直し方は利用者が選ぶ（最小の一手を先に）。検証は本人の実機テストが
   最速。初手で「いつから・何を変えた」を聞く。同名ブックが複数ありそうなら `versions <ブック名>` で写しを並べてから直す。
5. **検証の型**: 軽微な修正は `run-macro` 1 本。複数モジュールは影響するマクロの run-macro/test を同じ往復に相乗り。
   無人で回す停止条件は `gate`（構文＋全体コンパイル＋テスト Sub。テスト 0 本は FAIL）。全体だけなら `compile`。
   **検証を独立の往復にしない**（最後の置換と同じ往復に相乗り）。
6. **公開の型**: `py publish_check.py <公開用コピー> --scrub` → commit/push → 生ファイルを再チェック。
7. **表を直した後の報告は要点だけ 3 行程度（5 行以内）**（何をどこに直したかを番地で。経過・秒数・長い説明は書かない）。独立な調査・実行は 1 往復に並列で畳む。「遅い」はまず `stats` で内訳（道具の中 vs 間）。

## 手の型（シートを触るとき）

「値が合った＝できた」で止まらない。使う人が見るのは画面。**見る → やる → 仕上げる** の **3 往復** を
1セットにし、後ろを省かない。道具の中の処理そのものは数秒で終わる。時間を食うのは往復のあいだの
考え込みなので、型は「往復 3 回・決め直さない既定・時計は道具が持つ」で固める。

0. **表の修正・点検を頼まれたら**（「間違えているところを直して」もここから）:
   - 表の体裁修正は **`seiri`**（表を整えるマクロ＋残りだけの報告）。気づきのうち棚で直せる汚れ（重複行・空行・
     先頭のゼロ・空白の揺れ・文字の日付・式のずれ）は、**seiri が同じ呼び出しで撃って tidy まで仕上げる**
     （1 本撃つたびに数え直す・式を揃える手が行を消す手より先。2026-09-25）。AI が shelf-run を重ねない。
     残りに「棚で直せる手」が出たら、それは人の判断が要るもの（番号の重複）＝撃たずに番地で報告。
     マイナス・桁違い・式の中の数も撃たずに番地で報告。「棚を撃っても残った所」だけ write-cells で直す。
   - **頼みの文があれば `seiri` にそのまま渡す**（`seiri 郵便番号をハイフンつきに統一して 左の名簿と右の申込一覧を突き合わせて`）。
     表を整えるマクロに続けて、頼みの文に当たる棚（Excelコンボの先撃ちと同じ採点。頼みは空白・改行・「。」で分けて 1 つずつ当てる）も
     **同じ呼び出しで撃ち、書いた表を tidy まで仕上げる**＝棚撃ちが 1 回・数秒で済む（2026-09-24 夜: テスト用4 の 3 つの頼みが 1 回・3.2 秒）。
     当たらない頼みは報告の「残り」として直す。頼みの語が無い「修正して」は引数なしの `seiri`。
   - **seiri の「残り」がエラーセル: なし なら、それで仕上がり＝次の返事で 3 行程度（5 行以内）の報告を書いて終わる。**
     screenshot・気づき（文字色など）の直し・寄せの手直しを足さない（seiri が tidy と ### の検査まで済ませている。
     気づきは頼まれたときだけ直す）。Excelコンボの先撃ちと同じく道具は 1 回で終える
     （2026-09-25: seiri の後に 4 手足して、頼まれてから 27 秒かかった。道具の中は約 7 秒）。
   - 数式・データの点検は **`diagnose`**（循環参照・集計漏れ・外れ値。`--fix -y` は依頼に承認の語があるときだけ＝
     空白除去と文字列の数字の数値化だけ書き戻す。集計漏れの式は提案だけ・先頭ゼロは触らない・控えを取る）。
   - **残りは自分で書く前に棚を見る**: 依頼文から選ぶなら `shelf --ask 依頼文`（Excelコンボの先撃ちと同じ採点で 1 本。頼みが 2 つ以上で全部が棚に当たれば撃つ順の並び）、
     語で探すなら `shelf --grep 語`（棚＝秀コンボのモジュール「表の整理」「表の整理_作る」「表の整理_調べる」122 本の目録。名前・選ぶ列・窓の有無・扱う・説明）
     → 当たりがあれば **`shelf-run 名前 [--select 範囲] [--sheet 名]`** で撃つ。選ぶ列のあるマクロは列を 1 列ずつ
     カンマで分けて渡す（`--select A1:A6,B1:B6`。ひと続きの A1:B6 だと何もせずに終わる）。
     撃つ前に控えを取り、撃った後に差分（番地: 前→後・上限 30 件＋件数・書式の変化・足された行/列）と、ブックの違い
     （足されたシート・図形・グラフ・テーブル・ピボット・名前／書き換わった別シート）が出る。
     戻すのは `agent --undo`（作った物も別シートもまとめて戻る。消えたシート・名前は「戻らない」と出る）。
     **棚に無い所だけ** `write-cells`／`write_grid` で手で直す。
     目録に「窓あり」と付いたマクロは確認の窓を出す。答えるなら `--input-text 値`（無ければ安全側で閉じ、本文が報告に出る）。
   - 式のどこが元か分からないときは **`trace セル [--depth N]`**（番地・式・値の木。シートをまたぐ参照も追い、
     範囲は先頭と末尾と件数に畳んで中身の内訳（数・文字の数字・空・エラーの番地）を添える。列まるごと（A:C）は
     使っている所に絞り、名前（税率）は参照先へたどる。循環は印を付けて止める。何も書き換えない）。
1. **見る（1往復）**: **`materials`** だけ（見出しの結合の親子・二重下線の合計行・塗りの有無・循環参照・外れ値も
   materials が出す。番地の細部が要るときだけ 2 手目に `style-map`／`diagnose`）。開いているシートは **Excel に聞く。推測しない**（materials が
   アクティブシートを出す。sheet-info を前置しない）。小さい表（40行×30列まで）は全体が出る。
   **read-range を重ねない**。「続き」「もう一度」
   「この表」と言われたときも初手はこれ（会話記録・記憶・ソースを読み返さない）。
   materials が**仕事の時計を押す**（以後の write／tidy が「materials から N 秒」を出す）。
2. **やる（1往復）**: 飛び飛びのセルは **`write-cells C7 値 C11 値 …`** で 1 回。まとまった範囲・数式は
   MCP の **`write_grid(range, tsv)`** で 1 回（TSV を文字列でそのまま渡す。ファイルもパスも要らない。
   範囲は全範囲で書く）。CLI なら `write-range 範囲 --tsv ファイル --show`（パスはスラッシュ区切り）。
   複数の書き込みは**同じ往復に並べる**（順に実行される）。`--show`（write_grid は常時）で書いた直後の
   見え方（画面の文字・###）を読む。文字列が日付や数値に化けないかは**考えない**。道具が読み戻して
   化けていれば文字列で書き直す（`--raw` は要らない）。
3. **仕上げる（1往復）**: **`tidy 範囲1 範囲2 …`** で終える。表の左上が既存の見出しならその書式を
   行全体へ自動複写（`--header-from` は要らない）・罫線（細・格子）・**番号列は左寄せ・数値列は #,##0**
   （見出しと値から道具が判定。年・日付・率の列と既に書式のある列は触らない）・列幅の自動調整（6〜60）・
   ### 検査・**経過秒**を 1 回で出す。「整えて」と言われていなくても、表を新しく作った／列を足したら
   tidy まで含めて完成。範囲は**自分が書いた・広げた表の全体**（列を足したら、足した列だけでなく元の列を
   含む表全体。足した列だけに当てると罫線が継ぎはぎになる）。触っていない隣の表は含めない。
4. **完了条件**（報告の前に自分で確かめる）: 指示の項目が全部当たっている／### 0／見出し・罫線・列幅・
   寄せ・カンマが既存の表と揃っている／名前定義と既存の数式を壊していない（materials に出た名前を消さない）／
   保存はしていない。表の見た目を自分の手（write-cells・write_grid・format-range・tidy）で変えたら **`screenshot` で 1 回見てから渡す**（seiri だけで終わったときは撃たない）
   （文字の読み戻しだけでは罫線の継ぎはぎ・右寄せ・カンマ無しは見えない）。
5. **手数の多い組み立ては、マクロ 1 本に書いて撃つ。** ダッシュボード・帳票の作り直し・グラフやスライサーの配置のように
   10 手を超えそうな仕事は、一手ずつ道具を撃たない。`set_procedure_code`（Sub 全文）→ `add-procedure モジュール -y`
   （モジュールが無ければ先に `add-module Module1 -y`）→ `compile` → `run-macro 名前` → `screenshot`。直しは
   `patch_procedure` で数行だけ。実測（2026-09-19）: 一手ずつ＝52 回・298 秒で配置も崩れた／マクロ 1 本＝走るのは約 1 秒で、
   二度目からは AI も要らない。図形の位置はセルに合わせる（`Range("F1").Left`・`Range("F1:I1").Width`）。
   `.xlsx` にはマクロが残らない（道具が知らせる）。

**materials の次の返答は、道具の呼び出しで始める。** 書く前に決めてよいのは「どのセルに何を書くか」だけ。
道具の挙動・規則同士の両立・失敗の可能性は考えない。外れは `--show` に出て、撃ち直しは 1 秒で済む。

**決め直さない既定**（迷ったぶんだけ遅くなる。ここに無いものだけ考える）:
- 突き合わせ＝ `IFERROR(VLOOKUP(キー,表,列,FALSE),"〜なし")`。無いものは空欄でなく「〜なし」の文字。
- 表記の統一（全角→半角・ハイフン・桁）は値を直接書く（式にしない）。
- **掃除（空白・全角半角・カナ・重複・日付の形式）は汚れを調べない。** きれいな値を表の全部に上書きする。
  汚れの実態は materials の「気づき」と write の読み戻しで見える。汚れ方が見えなくても書く手は変わらない。
- **表の中の空欄で、隣の列から一意に決まる値（氏名→フリガナ、対応表のあるコード→名称）は埋める。**
  掃除の依頼に「空欄」の語が無くても埋め、報告に「C10 を埋めた」と書く。決まらない値（金額・日付・人の判断）
  だけ空欄のまま番地で報告。指示に無いのは「触るな」ではなく「決めて書け」。
- 行・列の手は `row delete 16`／`row insert 16`／`col delete C`（`--delete` ではない）。
- **スペースを含む名前（シート名・図形名・マクロ名）は引用符で囲む**: `shape --list --sheet "売上 2026"`。
  囲まないとスペースで別々の引数に割れ、違う対象（アクティブシートなど）に当たることがある。
- 数式は TSV 経由（インラインの引用符事故を避ける）。TSV は UTF-8・1行＝1行・タブ＝列。
- tidy の範囲＝自分が書いた・広げた表の全体（列を足したら表全体）。隣の表は指示があるときだけ。
- 文字列の化け（"1-2"→日付、"1,000"→数値）は道具が戻す。`--raw` を付けるかは考えない。
- 依存する道具の呼び出しも同じ返答に並べる（順に実行される）。TSV の置き場も考えない（write_grid）。
- **表を直した後の報告は要点だけ 3 行程度（5 行以内）**（2026-09-25 shu「直しましただけで中身が全然ないのはダメ・長い説明はいらない」「報告は 3 行程度・5 行以内」）。
  何をどこに直したかを番地で（例: C7・C11 の郵便番号をハイフンつきに／F・G 列に突き合わせの式）。経過・秒数・
  手順の説明は書かない（聞かれたら答える。長い報告は書くだけで 9 秒かかっていた）。往復は 4 回が上限。
- **表について答えるときは番地を添える**（「C15 の合計は E4:E14 の和・E9 だけ文字の数字」）。「合計が合いません」
  のように番地の無い言い方をしない。番地が分からないなら `trace` か `diagnose` で確かめてから答える。
- **`vba_help` を引かない。** 引数がうろ覚えでも help を挟まずに撃つ。外れれば usage が返る＝それが help。

**シートを触る手の引数（これだけ覚えて help を引かない）**:
- 結合の解除 `format-range 範囲 --unmerge`（同コマンドに `--merge --border thin --align left|center|right
  --number-format "#,##0" --autofit --bold --bg #DDEBF7 --row-height N --wrap --show`）
- 中身を消す `clear-range 範囲`
- 図形 `shape --list` ／ 消す `shape --delete "名前" "名前" …`（名前は materials に出る。マクロ付きボタンは
  道具が拒む＝安全。保存前なら閉じれば戻る）
- **「図形を消して」＝マクロの付いていない図形は全部消す。** materials の図形一覧で「→ マクロ名」が
  付いていないものが対象。ロゴ・飾り枠・見出しの飾りを「正当な部品だから」と自分の判断で残さない
  （線引きは道具が持っている＝マクロ付きは拒まれる。保存前なら閉じれば戻る）。
- 他: `print-setup` `freeze` `sort` `autofilter` `find-replace` `fill` `copy-range` `row` `col` `name`
- **帳票→一覧（結合セルの帳票を 1 行 1 件に）は 4 手**: `format-range 表全体 --unmerge` →
  `clear-range 表全体` → `write_grid`（見出しは 1 段に畳む・1 行 1 件）→ `tidy 表全体`。
  「不要な図形を消して」が付いたら `shape --delete` を write_grid と同じ往復に並べる。

## ツールの場所

- SCRIPTS: vba_manager.py 等が置いてあるフォルダ（絶対パス決め打ち禁止。不明なら Glob `**/vba_manager.py`）。
  `_last_proc.vba`（get の出力・UTF-8）・`backups\`（自動バックアップ）もここ基準。すべて SCRIPTS から `py` で実行（`python` ではない）。
- SCRIPTS は「ツールの場所」であって「操作対象」ではない。対象は常にアクティブブック。
- ⚠ Excel 未起動でパス指定すると Excel が自動化起動される（アドイン・PERSONAL を読まない）。作業後はその Excel を閉じる。

## コマンド索引（1 行 1 コマンド。引数の細部は references/commands.md）

```bash
# --- VBA を読む・直す ---
list [--detail] [--module 名] [--all] [--json]         # マクロ一覧（--module で絞る）
get <Sub名> | get <モジュール> <Sub名> | get A.x B.y     # コード取得 → _last_proc.vba（3 個以上は連結）
replace-procedure -y [--module 名] [--code-file f]    # _last_proc.vba で置換（控え・自動保存）
add-procedure <モジュール> -y / delete-procedure [モジュール] <Sub名> [Sub名…] -y   # 追加・削除（複数は控え 1 冊・保存 1 回）
add-module <名> [--type std|class|form] / delete-module <名> -y
replace-module <モジュール> <basファイル>                 # Remove+Import（VB_Name 照合・.frm は .frx 同伴）
code-replace "旧" "新" [--module 名] [--regex] -y       # 全マクロ横断の一括置換（変更行だけ ReplaceLine・寛容な探し直し）
rename-procedure <旧名> <新名> [--module 名] -y [--dry-run]  # 改名＝宣言・呼び元・Application.Run の文字・OnAction・ショートカット再登録（2026-09-16）
grep "文字" [--regex] [-i] [--module 名] [--max N]      # 全モジュール横断の検索
export-module <モジュール> / export-all [--dir 先] [--check] [--history]   # .bas 書き出し（--history＝_exports/<ブック>/<日時>/ と前回の差・2026-09-16）
diff-module <モジュール> [比較先.bas] [--backup [名]] [--book 他ブック]   # モジュールの差分（既定は直前の控え・2026-09-16）
format-module <モジュール> [--apply -y]                 # 整形（Dim 先頭・区切り・見出しコメント。既定は差分だけ・2026-09-16）
copy-modules <モジュール…> --to <ブック> [--overwrite] -y   # 別の開いているブックへ複製（ショートカット再登録つき・2026-09-16）
references list | add <GUID|パス> | remove <名> -y      # 参照設定（参照不可＝眠ったブックが動かない筆頭原因・2026-09-16）
list-shortcuts / set-shortcut <マクロ> <キー> -y | --clear   # ショートカット一覧・付け替え（k=Ctrl+k・K=Ctrl+Shift+K・2026-09-16）
#  同じブックでほかのマクロが同じキーを持っていれば、そちらから外して掛ける（Excel は両方に残し名前の順で先の方を動かす・9/20）。一覧は同じキーに印
reorder-macro <マクロ> <up|down|top|bottom|番号>         # メニュー表示順
list-backups [語] / restore <控え.bas> / backup-prune [--days 30] [--keep 1] [--force]   # 控えの一覧・復元・間引き（2026-09-16）
repair <マクロ> [--module 名]                          # 修理の材料を 1 手で＝本文（_last_proc.vba にも保存）・入口・呼び元呼び先・コンパイルの落ちた行・check の error 級・控えとの差分（2026-09-17）
versions <ブック名> [--dir 追加] / history <マクロ> [--book 名] [--deep]   # 同名ブックの写しを日時順に（最新との差・開いている印）・マクロ本文の歴史（同じ本文は畳む）。Excel を開かない（2026-09-17）
list-file <path.xlsm> / grep-files "文字" <フォルダ|ファイル…> [--regex] [-i] / export-file <path.xlsm> [モジュール…] [--dir 先]   # 閉じたブックを読む（oletools＝Workbook_Open もアドインも起きない。50 冊でも 1 手・2026-09-17）
check [excel_file] [--all-warnings] / check-bas <f.bas> [--fix] / rules  # 静的診断（VBM001〜014。ラベル無し・On Err・存在しない呼び先は error。VBM003 は件数だけ＝--all-warnings で行）
compile / gate [絞り込み] [--timeout 秒] / test [絞り込み]   # 全体コンパイル（落ちた行を [モジュール] Sub:行: 本文 で名指し）・関所（コピーで check＋compile＋test）・テスト Sub 一括
run-macro <マクロ> [引数…] [--input-text 値] [--raw]   # 実行（ハーネス経由＝実行時エラーが文字で返る）
rehearse <マクロ> [--timeout 秒] [--addins]             # コピーに試し撃ちして差分報告（本体は無傷）
call-graph [--macro 名] [--mermaid] / impact <マクロ> / wiring / docs   # 呼び出し関係・影響範囲・ボタン配線・取説
checkup [--form] [--note "…"] [--history]              # 健康診断（利用者が明示したときだけ）
inspect-gui [マクロ] / capabilities [コマンド] / metrics / process [--kill PID] / setup-check / diag
close-form [題名] [--list] / vbe-reset [--check]        # 表示中フォームを閉じる・中断モード解除（人の指示でだけ）
form-to-vba <フォーム> [--all] [--add モジュール] [--verify]   # UserForm を作成マクロに変換
stats [--days 7] [--top 20] [--slow 8] / status         # 呼び出し台帳の集計・いま走っているコマンドと進み具合（COM 不要）
<長い手> --bg                                          # MCP の vba() で行末に付けると待たずに job id（checkup・export-all・gate・rehearse・snapshot・docs）。段は status、結果は agent_status("job…")（2026-09-17）
batch cmds.txt / shell                                  # 1 接続で連続実行・対話

# --- 目（読むだけ）---
materials [シート] [--rows 5]     # 1 手目。使用範囲・値・結合・テーブル・名前・数式の型と気づき・図形・###・列幅を 1 回で
seiri [--dedupe]                  # 表の修正の 1 手目＝整えるマクロ＋残りだけの報告
shelf [--grep 語] [--ask 依頼文]  # 棚（表の整理・_作る・_調べる 122 本）の目録＝名前・選ぶ列・窓の有無・扱う・説明／依頼文から 1 本選ぶ
trace <セル> [--depth 3]          # 式の元を番地・式・値の木で（シートをまたぐ・範囲は畳んで内訳・列まるごと・名前・循環は止める）
read-range 範囲… [--formula] [--tsv] / read-selection / sheet-info [--preview 3] / table read <名> [--tsv]
screenshot [範囲] [--out f.png]   # 範囲を PNG に（Undo 履歴が消える）
snapshot [--sheet 名] [--no-format] / snapshot-diff <before.json> [after.json]   # 意味構造 JSON と機械的差分
audit 範囲                        # 仕上げ検査（指摘 0 件で合格）
find 文字 [--book] [--whole] [--formula]

# --- 手（開いたままのブックに直接書く。既定では保存しない）---
write-range 範囲 値 | --tsv f [--raw] [--show] [--append]   # MCP は write_grid(range, tsv)
write-cells C7 値 C11 値 … [--show] / clear-range 範囲 [--contents|--formats]
format-range 範囲 [--bold --bg --color --align --number-format --border --merge --unmerge --autofit …]
tidy 範囲… [--header-from セル] [--bg]     # 仕上げ（見出し・罫線・列の型・列幅・###・経過秒）
shelf-run <名前> [--select 範囲] [--sheet 名] [--input-text 値]   # 棚のマクロを選んで撃つ＝控え→撃つ→差分。戻すのは agent --undo
register-addin [ブック] [-y]              # 前に出ているブックを .xlam に焼き直してアドイン登録（別名 更新登録。
#   アドインの VBA を直したら**人に頼まず自分でこれを撃つ**。焼く先・空のブック・表示中フォームは道具が確かめ、
#   素の Run で撃って（ハーネスは xlam に焼き付く）、焼けた .xlam の日時と大きさを報告する（2026-09-23）
sheet add|rename|copy|delete|activate|show|hide|very-hide|visibility|tab-color
table create|list|delete|ref|column add|rename|format|filter|filter-values|filters|filter-clear|sort|sort-multi
name add|list|delete / row insert|delete 行 [本数] / col insert|delete 列 [本数]
copy-range 元 先 [--values] / fill 範囲 [--right] / sort 範囲 --key 列 [--desc] [--header] / autofilter [範囲] [--off]
find-replace 旧 新 [範囲] / open パス / close ブック --save|--no-save -y / save / save-as パス [--overwrite]
export-pdf 出力.pdf [--sheet|--range] / print-setup [--area --title-rows --landscape --fit-wide]
cond-format 範囲 --gt 値 --bg 色 / hyperlink セル URL / validation 範囲 --list "A,B" / freeze セル|off / comment セル 文
shape --list | --delete 名… | 名 --left N / calc-mode [manual|auto|recalc]
build-sheet [設計図.json] [--ask 依頼文] [--recipe 名] [--new-book] [--dry-run] [--fire]   # 設計図からシートを組む
agent "依頼文" [--sheet 名] [--mode sheet|build|macro] [--macro 名] [--dry-run] [--undo] [--runs] [--fire …]   # 全文は references/agent.md

# --- 重量級 ---
chart create 範囲 --type column --title … / chart list|delete / chart-config <set-title|set-type|legend|axis-scale|add-series|set-source|trendline …>
pivot create 元!A1:C100 --rows 部門 --cols 月 --values 売上 [--sheet 集計] / pivot list|delete
pivot-field list|add-row|add-col|add-value|remove|set-func|sort|group-date|show-as|top-n|position … / pivot-calc get-data|calc-field|layout|refresh|set-source …
slicer add|list|delete / powerquery list|refresh|add|edit|delete|load --to sheet|model / connection list|refresh|delete
datamodel list / relation add|delete / measure add|delete
```

## 絶対に守るべきルール（詳細と経緯は references/workflows.md）

- **識別子**: プロシージャ名・モジュール名は英字か日本語で始める。先頭 `_` は注入は通るがコンパイルで死ぬ
  （道具は validate_vba_code 等で機械拒否する。--force で潰さない）。注入経路の台帳は test_tools.py の
  `test_injection_route_ledger`＝新しい注入経路を作ったらガードを配線して登録する。
- **エンコーディング**: .bas は **CP932**。**.bas に Edit／Write ツールを絶対に使うな**（UTF-8 化して日本語が全滅する）。
  直すなら Python で CP932 のまま読み書きするか、`_last_proc.vba`（UTF-8）経由の replace-procedure。
  取り込み前は `check-bas`（文字コード・改行二重化・重複・識別子）。
- **モジュール適用**: AddFromString は Attribute を壊すことがある＝**replace-module（Remove+Import）**。
  副作用: モジュールが末尾へ移動（表示順が変わる）。ショートカットは Import 直後に道具が MacroOptions で再登録する。
- **InsertLines の改行**: Python の \n では分割されないことがある＝.bas 直接編集方式。
- **勝手な変更の禁止**: 指示されていない機能を足さない・コードを変えない・影響範囲を確認せずに変えない。
- **破壊的操作の前に現地調査**: シートに触るマクロ・行列削除・クリアの前に `snapshot`／`materials` で実勢（結合・見出し・
  図形・テーブル）を見る。変更を伴う作業は前後で snapshot → snapshot-diff。
- **動作検証は COM 経由が第一選択**（run-macro／test／snapshot-diff）。computer-use の打鍵は最終手段。
- **checkup／check は利用者が明示したときだけ**。裁量で回さない。
