# 変更履歴

## 0.1.1（2026-09-23）— 表を壊す不具合の直し

- `agent --undo`: 戻すときに、文字で入っていた数字を数値・日付に化かしていたのを直しました（「1,000」→1000・「0021」→21・「4-1」→日付）。
  控えの .Formula で範囲の全部を書き直していたためです。入れ直すのは、外部リンクに化けた数式のセルだけにしました。
- 書く道具（write-cells・write_grid・write-range）: 先頭が 0 の数字（伝票番号「0004」・郵便番号「007」）を数にして 0 を消していたのを直しました。
  文字のまま書きます。

## 0.1.0（2026-09-18）— 秀エクセルマネージャーとして公開

前身の「秀VBAマネージャー」（https://github.com/shu1551/shu-vba-manager ・2026-09-05 の姿で凍結）から、
手元で開発中の道具一式を、そのまま出したものです。研究中・仕様は変わります。

- MCP サーバーの名前を `vba-manager` から `excel-manager` に変えました（ファイル名と道具の名前はそのまま）
- AI エージェント（agent）を 11 本に分割。マクロの先撃ち（登録済みのマクロが AI より先に撃つ）
- 正解の表を持つ試験（agent --exam）・記録の揺らし（agent --shake）
- 鍛える回路（agent --forge）: AI がやった仕事を AI がマクロに書き、正解つきのお題で試し、合格したら登録する
- 修理の試験（agent --mend）: 合格済みのマクロを 1 行壊し、人の言い方で AI に直させ、全部の表で確かめる
- グラフ・ピボット・テーブル・クエリの姿を読んで比べる目（vbam_objects）
- 修理を 1 手に（repair）・コンパイルエラーの行を名指し・閉じたブックの VBA を読む手（versions / history / list-file / grep-files）
- 開発向けの手（rename-procedure・diff-module・references・copy-modules・set-shortcut・format-module・backup-prune・status）
- 数式とデータの診断（diagnose）・書式の地図（style-map）・呼び出しの台帳（stats）
- 表の整理のマクロ（`作業ファイル/project/macros/表の整理.bas`）
