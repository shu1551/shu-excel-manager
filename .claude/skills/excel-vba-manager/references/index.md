# コマンド索引（全部）

SKILL.md の「よく使うコマンド」に無い手を探すときに読む。引数の細部・注意は commands.md。

### コマンド索引（1 行 1 コマンド。引数の細部は references/commands.md）

```bash
# --- VBA を読む・直す ---
list [--detail] [--module 名] [--all] [--json]         # マクロ一覧（--module で絞る）
get <Sub名> | get <モジュール> <Sub名> | get A.x B.y     # コード取得 → _last_proc.vba（3 個以上は連結）
replace-procedure -y [--module 名] [--code-file f]    # _last_proc.vba で置換（控え・自動保存）
add-procedure <モジュール> -y / delete-procedure [モジュール] <Sub名> [Sub名…] -y   # 追加・削除（複数は控え 1 冊・保存 1 回）
add-module <名> [--type std|class|form] / delete-module <名> -y
replace-module <モジュール> <basファイル>                 # Remove+Import（VB_Name 照合・.frm は .frx 同伴）
import-module <ファイル…> [--book 名] [--replace] -y      # モジュール（.bas/.cls/.frm）を取り込む（同名は --replace で置換、取り込み後全体コンパイル確認）
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
flow <マクロ名> [--module 名] [--out x.md]    # マクロの中の流れ図（Mermaid・コードから機械的に・AI なし。引き継ぎ資料に）
checkup [--form] [--note "…"] [--history]              # 健康診断（利用者が明示したときだけ）
inspect-gui [マクロ] / capabilities [コマンド] / metrics / process [--kill PID] / setup-check / diag
close-form [題名] [--list] / vbe-reset [--check] / unstick [--fix] / close-vbe   # 表示中フォームを閉じる・中断モード解除・Excelの詰まり解除・VBEを閉じる
form-to-vba <フォーム> [--all] [--add モジュール] [--verify]   # UserForm を作成マクロに変換
stats [--days 7] [--top 20] [--slow 8] / status         # 呼び出し台帳の集計・いま走っているコマンドと進み具合（COM 不要）
stats --slow [--days 14]                                  # 遅い手を名指し（合計秒順・1.3倍悪化や目安超えに⚠・COM 不要・2026-10-29）
<長い手> --bg                                          # MCP の vba() で行末に付けると待たずに job id（checkup・export-all・gate・rehearse・snapshot・docs）。段は status、結果は agent_status("job…")（2026-09-17）
batch cmds.txt / shell                                  # 1 接続で連続実行・対話

# --- 目（読むだけ）---
materials [シート] [--rows 5]     # 1 手目。使用範囲・値・結合・テーブル・名前・数式の型と気づき・図形・###・列幅を 1 回で
seiri [--dedupe]                  # 表の修正の 1 手目＝整えるマクロ＋残りだけの報告
shelf [--grep 語] [--ask 依頼文]  # 棚（表の整理・_作る・_調べる・_書式・_基本操作）の目録＝名前・選ぶ列・窓の有無・扱う・説明／依頼文から 1 本選ぶ
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
#   焼く前に棚の当たり方の検査と gate（コピーでテスト Sub を全部撃つ・30 秒ほど）を回し、落ちたら焼かない
#   （テスト Sub の無いブックは gate を飛ばす。飛ばすなら --no-check / --no-test・2026-10-21）
sheet add|rename|copy|delete|activate|show|hide|very-hide|visibility|tab-color
table create|list|delete|ref|column add|rename|format|filter|filter-values|filters|filter-clear|sort|sort-multi
name add|list|delete / row insert|delete 行 [本数] / col insert|delete 列 [本数]
copy-range 元 先 [--values] / fill 範囲 [--right] / sort 範囲 --key 列 [--desc] [--header] / autofilter [範囲] [--off]
find-replace 旧 新 [範囲] / open パス / close ブック --save|--no-save -y / save / save-as パス [--overwrite]
export-pdf 出力.pdf [--sheet|--range] / print-setup [--area --title-rows --landscape --fit-wide]
cond-format 範囲 --gt 値 --bg 色 / hyperlink セル URL / validation 範囲 --list "A,B" / freeze セル|off / comment セル 文
shape --list | --delete 名… | 名 --left N / calc-mode [manual|auto|recalc]
build-sheet [設計図.json] [--ask 依頼文] [--recipe 名] [--new-book] [--dry-run] [--fire]   # 設計図からシートを組む
agent "依頼文" [--sheet 名] [--mode sheet|build|macro] [--macro 名] [--dry-run] [--undo] [--runs]   # 全文は references/agent.md

# --- 重量級 ---
chart create 範囲 --type column --title … / chart list|delete / chart-config <set-title|set-type|legend|axis-scale|add-series|set-source|trendline …>
pivot create 元!A1:C100 --rows 部門 --cols 月 --values 売上 [--sheet 集計] / pivot list|delete
pivot-field list|add-row|add-col|add-value|remove|set-func|sort|group-date|show-as|top-n|position … / pivot-calc get-data|calc-field|layout|refresh|set-source …
slicer add|list|delete / powerquery list|refresh|add|edit|delete|load --to sheet|model / connection list|refresh|delete
datamodel list / relation add|delete / measure add|delete
```
