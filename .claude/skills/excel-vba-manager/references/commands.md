# vba_manager.py コマンド一覧（全文）

> excel-vba-manager の参照。SKILL.md（型と索引）から必要なときだけ読む（2026-09-16 に分割）。

## vba_manager.py コマンド一覧

```bash
cd <SCRIPTS>   # vba_manager.py のあるフォルダへ。不明なら Glob **/vba_manager.py で特定（絶対パスを決め打ちしない）

# マクロ一覧表示
py vba_manager.py list
#  オプション: --standard（標準モジュールのみ）--personal（PERSONAL.XLSB）--addin（アドイン）
#              --all（全ブック横断）--json（機械可読出力）
#              --detail（所属モジュール・行数・先頭コメント付き）--module 名（対象モジュール限定）

# 全モジュール横断のVBAコード検索（どのマクロが何を使っているか調べる）
py vba_manager.py grep "ActiveSheet"                  # [モジュール] プロシージャ:行番号: 該当行
py vba_manager.py grep "On Error" --module shu003 -i  # --regex（正規表現）--max N --json も可

# ブックの構成ドキュメント（取説）を自動生成 → _last_docs.md
py vba_manager.py docs [--out f.md] [--preview 3]  # --preview で各シート先頭N行のMarkdown表も含める

# マクロの呼び出し関係を解析（Call/Application.Run/裸呼び）
py vba_manager.py call-graph                    # 未解決Call(存在しないマクロ呼び=一語バグ)・関係一覧・孤立
py vba_manager.py call-graph --macro 親処理      # そのマクロ起点の呼び出しツリー＋呼び元
py vba_manager.py call-graph --mermaid          # Mermaid図を _last_callgraph.md に（未解決は赤ノード）

# 対話セッション（接続を張ったままコマンドを打ち続ける。2コマンド目から再接続なし）
py vba_manager.py shell                         # exit で終了。batch のファイル版に対する対話版
#  history で履歴一覧、!番号 で再実行（!! は直前のコマンド）

# ブックの健康診断（総合判定+壊れた参照+シート検査+コード/フォーム検査+前回との比較を1枚に）
#  ※ AI はこの診断を裁量で回さない（利用者が「健康診断」「checkup」と明示したときだけ実行）
py vba_manager.py checkup [--out f.md]          # → _last_checkup.md。初見ブックの受け入れ検査にも
py vba_manager.py checkup --form                # フォーム検査込みのフル診断
#  既定はフォーム検査なし＝数秒で完了（Designer のコントロール走査が診断時間の支配項のため、
#  --form 指定時のみ実施。未実施はレポートに明記され、前回比較でも偽の解消/新規は出ない）
py vba_manager.py 健康診断                       # ↑の日本語別名（同じもの）
py vba_manager.py 健康診断 --all                 # 開いている全ブックを一括診断（PERSONAL.XLSB含む）
py vba_manager.py 健康診断 --note "直した内容"    # カルテのメモ＝今回の記録に治療内容を残す（修理直後に打つ）
py vba_manager.py 健康診断 --history             # 経過観察＝過去の診断履歴を表で表示（診断はしない）
py vba_manager.py checkup --ack-all              # 現在の所見を「確認済み」にする（以後の診断で既知として扱う）
py vba_manager.py checkup --show-ack             # 確認済みにした所見の一覧
py vba_manager.py checkup --unack "文字列"       # 文字列を含む確認済みの所見から印を外す（また未確認として出る）
py vba_manager.py 健康診断 --history --detail    # 各回の間の所見/マクロ増減も表示（メモと合わせてカルテになる）
#  終了コードは「診断完了=0」（所見があっても失敗ではない）。所見の有無で合否を取りたい
#  自動化ゲートだけ --strict を付ける（所見1件以上で終了コード1）
#  総合判定 A（所見なし）/B（所見あり）/C（壊れた参照・呼び出しあり）を機械分類でレポート先頭に表示。
#  所見（カウント対象）: 未解決Call・参照設定の破損(MISSING＝眠ったブックが動かない筆頭原因)・
#  外部リンク切れ・#REF!名前付き範囲・コード重複・フォームlint。
#  参考所見（数えない）: エラーセル数・UsedRangeゴースト（ファイル肥大候補）・破壊的操作の所在
#  （Kill/FSO/シート削除/行列削除＝起こす前の問診）・ScreenUpdating/Calculation戻し忘れ・
#  自動実行イベント・ハードコードパス・Option Explicit なし・On Error Resume Next・150行超。
#  診断のたびにブック別履歴（_checkup_history/）へ記録し、次回から「前回との比較（定期健診）」を
#  自動表示（新規/解消した所見・マクロ/シート/フォームの増減・行数の増減。行ずれでは差分を出さない）。
#  シート上のボタン/図形に登録されたマクロ（OnAction）も走査し「孤立」から除外＋一覧表示。
#  VBAパスワード保護・VBOM未信頼のブックはシート側だけの縮退診断に自動切替（判定保留・履歴に混ぜない）。
#  Declare 宣言済み Win32 API への Call は未解決扱いにしない（call-graph 共通）。
#  ※ シート検査は COM Find を使うため Excel の検索ダイアログ設定が変わる（内容は変えない）

# マクロ修正前の影響範囲予告（このマクロに手を入れるとどこまで波及するか）
py vba_manager.py impact <マクロ名>              # 別名: 影響範囲。--json も可
#  ■呼び元（直すと影響が及ぶ先・間接含む。フォームのボタン/イベントも出る）
#  ■呼び先（依存している部品）■入口（ショートカット/自動実行イベントから届く経路）

# モジュール一覧表示
py vba_manager.py list-modules

# 開いているブックの一覧（Excel生存確認にも。素の py -c ワンライナーでなくこれを使う）
py vba_manager.py list-open

# マクロを実行（ブック名!マクロ名 or マクロ名だけなら全プロジェクトから検索）
py vba_manager.py run-macro <マクロ名> [--json] [--timeout 秒] [--input-text 値] [--raw]
#  ⚠ 既定は「ハーネス経由」（2026-08-23 に既定を変更）。持ち主ブックに一時モジュール
#    （VbaManagerRunHarness）を注入し、On Error で受けてから対象を呼ぶ。実行時エラーは Excel の
#    モーダルダイアログにならず「実行時エラー 番号: 説明」の文字で返る。一時モジュールは finally で必ず撤去。
#    理由: Application.Run 越しの呼び出しはエラーが呼び元に伝播しない（実測 2026-08-23）ので、
#    ハーネスは必ずマクロの持ち主ブックに置いて直接呼ぶ。
#  --raw … ハーネスを使わず素の Application.Run で撃つ（従来の挙動へ戻す逃げ道）。
#    ⚠ write-range --raw（数値変換せず文字列で書く）とは意味がまったく違う同名フラグ。取り違えるな。
#  次の場合は指定しなくても素の Run に落ちる: 引数付き実行／持ち主ブックが特定できない／
#    マクロ名・モジュール名が VBA 識別子として妥当でない／宣言行が見つからない。

# 表示中の UserForm を閉じる（フォームを直す前に人へ「閉じてください」と頼まない＝自分で閉じる）
py vba_manager.py close-form                  # Excel が表示している全フォーム
py vba_manager.py close-form <キャプション>     # 窓の題名で絞る（完全一致→部分一致）
py vba_manager.py close-form --list           # 閉じずに一覧だけ
#  ×ボタンと同じ WM_SYSCOMMAND(SC_CLOSE) を窓（ThunderDFrame）へ送る。COM を使わないのでモーダル
#  表示中（COM が「呼び出し先が拒否しました」で弾かれる状態）でも効く。QueryClose で Cancel する
#  フォーム（処理の実行中など）は閉じられない＝その旨を報告して終わる（無理に殺さない）。
#  ⚠ 自前で閉じるコードを書くとき WM_CLOSE を直接投げるな（2026-08-23 実測）：窓だけ壊れて VBA の
#    Unload を通らず、次に開くと実行時エラー 80010108（切断）→ Excel ごと落ちる。

# VBE の中断モードを外す（実行時エラーで黄色い行のまま止まった状態の解除）
py vba_manager.py vbe-reset --check            # 押さずに「今リセットが要る状態か」だけ返す
py vba_manager.py vbe-reset                    # VBE の「実行 > リセット」を外から押す
#  症状: 中断モードだと Designer が取れない・他のマクロが「実行できません」になる・
#    form_inspect 等が「VBE が中断モードでないか確認してください」で止まる。
#  判定は VBProject.Mode（0=実行中 / 1=中断 / 2=設計）。リセットボタンの Enabled は VBE 窓が
#    隠れていると更新されず当てにならない（実測 2026-08-23）ので見ない。
#  押したあと Mode を読み直し、解けていなければ「解けなかった」と正直に返す（押せる≠解ける）。
#  ⚠ **自動では撃つな。人が「戻せ」と言ったときだけ。** 走っている VBA を全部止めるので、
#    表示中のフォームは閉じ、モジュール変数は消え、温めている常駐プロセスがあれば切れる。

# モジュール丸ごと削除（要約表示→確認→バックアップ。戻すのは restore）
py vba_manager.py delete-module <モジュール名> -y

# マクロを引数つきで実行（数値に見える引数は数値化して渡す）
py vba_manager.py run-macro 加算 3 4        # Function なら戻り値も表示

# InputBox を出すマクロに値を入れて先へ進ませる（2026-08-23・xlflow の --inputbox 相当）
py vba_manager.py run-macro 名前を聞く --input-text 太郎
py vba_manager.py run-macro 範囲を聞く --input-text 太郎 --input-text B2:C3   # 出た順に1つずつ
#  VBA の InputBox() も Excel 内蔵の Application.InputBox(Type:=8 等) も可。値を入れて OK（内蔵は Enter）で確定。
#  足りなければ最後の値を使い回す。入れた値は報告に出る（--json なら inputs_entered）。
#  --input-text 無しの InputBox は従来どおり閉じるだけ（安全解除＝キャンセル）。test / rehearse でも同じ引数。

# 予行演習run（2026-07-15）＝マクロをコピーに試し撃ちして差分報告（本体は無傷）
py vba_manager.py rehearse <マクロ名> [引数...] [--auto-dialog ok] [--input-text 値] [--addins] [--timeout 秒] [--out path] [--discard] [--max N]
py vba_manager.py 予行演習 <マクロ名>            # 日本語別名（同じもの）
#  流れ: SaveCopyAs（未保存の変更込みの「今この瞬間」）→ 別インスタンスでコピーを開く
#  （Workbook_Openは起こさない）→ 前snapshot → 実行 → 後snapshot → snapshot-diff報告。
#  マクロが途中で落ちても、そこまでの変化を差分で報告する（それも収穫）。
#  実行は run-macro と同じハーネス経由（2026-08-23）＝実行時エラーは「実行時エラー 番号: 説明」で返る（引数付きは素の Run）。
#  --timeout 秒 を過ぎたら演習用 Excel を強制終了して「時間切れ」で報告（本体は無傷・固まらない）。
#  結果コピーと前後snapshotは残す（--discard で削除）。本体に焼く判断は人がやる。
#  ⚠ 非破壊が保証できるのはブックの中だけ。ファイル出力・メール送信など外への副作用は止められない。
#  アドイン（秀.xlam等）の関数を呼ぶマクロは --addins を付ける（演習用Excelは素の環境のため）。

# VBAテストランナー（xlflowのテスト基盤から発想だけ移植・2026-07-09）
py vba_manager.py test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--input-text 値] [--timeout 秒] [--json]
py vba_manager.py テスト                     # 日本語別名（同じもの）
#  名前が「テスト」または「test」で始まる**引数なしの公開Sub**を一括実行し成否一覧。
#  失敗の知らせ方は Err.Raise だけ（assert例: If 実際 <> 期待 Then Err.Raise 5, , "説明"）。
#  補助モジュール不要＝テストSubも単体で他ブックへ移植できる自立ユニット。
#  実行はエラー捕捉ハーネス（一時モジュール注入→直接呼びラッパー→撤去）経由なので
#  実行時エラーでVBAダイアログは出ない。ブックは保存しない。
#  全部成功で終了コード0／1本でも失敗なら1（自動化ゲートに使える）。
#  ⚠ テスト中にMsgBoxを出すコードには --auto-dialog ok を付ける（付けないと止まる）。
#    InputBox を出すコードには --input-text 値（値を入れて OK で先へ進む。無ければキャンセルで閉じるだけ）。

# 関所＝停止条件（2026-08-23）。「直した。で、通ったのか？」を機械が一言で答える＝自動で回すときの終わり方
py vba_manager.py gate [excel_file] [絞り込み] [--module 名] [--addins] [--timeout 秒] [--auto-dialog ok] [--input-text 値] [--keep] [--json]
py vba_manager.py 関所                       # 日本語別名（同じもの）
#  流れ: SaveCopyAs（未保存の変更込み）→ 別インスタンスでコピーを開く（Workbook_Open は起こさない）
#  → check（構文検査）→ **全体コンパイル** → test（テスト Sub 一括）→ 判定。
#    本体には指一本触れず、演習用 Excel は畳んでコピーも消す（--keep で残す）。
#  判定: 構文エラー 0 かつ 全体コンパイル OK かつ テスト 1 本以上 かつ 全部成功 → 通過（PASS・終了コード 0）。
#    それ以外は不通過（FAIL・終了コード 1）。
#  **テスト 0 本は FAIL**（合否を言えないものは通さない）。時間切れ（既定 120 秒）は演習用 Excel を強制終了して FAIL。
#  --json は {"pass","book","check":{errors,warnings},"compile":{ok,state,detail},"tests":{…},"timeout","seconds","reasons"} の1行。
#  回し方: 書く → gate → FAIL の理由（failed の error／構文エラー／コンパイルエラー）を読んで直す → gate … PASS で止まる。
#  ⚠ テスト Sub が触る外への副作用（ファイル出力・送信）は止められない。
#  ⚠ 演習用 Excel は素の環境（アドイン無し）。アドインの関数を呼ぶテストは --addins。

# 全体コンパイル（2026-08-28）。check が見ない「呼ばれないマクロの中」まで見る唯一の手
py vba_manager.py compile [excel_file] [--json]
py vba_manager.py 全体コンパイル               # 日本語別名（同じもの）
#  VBE の「デバッグ > VBAProject のコンパイル」(CommandBars Id 578) を外から押す。通れば終了コード 0、
#  コンパイルエラーなら理由（VBE のメッセージ本文）を出して 1。VBE の窓は出さない・走っている VBA も止めない。
#  2026-09-17: 落ちた行を名指しする＝`[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― 行ラベルが定義されていません。`
#  （エラー直後の VBE の ActiveCodePane から読む。--json は module/proc/line/text も返す）＋直し方の 1 行
#  （ラベル無し→ get→replace-procedure、それ以外→ code-replace）。エラーで前に出る VBE の窓は押す前が隠れていたなら戻す。
#  **なぜ要るか**: VBA はプロシージャ単位のオンデマンド コンパイル＝存在しない処理を呼ぶ Sub は
#  「実行するまで」誰も気づかない。check は静的検査なので素通りする（2026-08-28 実測で確認）。
#  合否は Id 578 の Enabled で見る（通れば「もう押す必要がない」＝False に落ちる）。エラー時の
#  モーダル MsgBox はダイアログ監視が閉じて本文を持ち帰る。
#  押せない環境（VBOM 未信頼・Id 578 が無い）は state=unavailable で **合否を言わない**（ok 扱い）。
#  「見られなかった」を「落ちた」にしないため＝gate が全部 FAIL になるのを避ける。

# 制限時間（2026-08-23・--timeout 秒。付けないときは従来どおり無制限）
#  rehearse / gate: 演習用（自分で起こした）Excel を強制終了して即戻る。本体は無傷。無人で回す経路はこちら。
#  run-macro / test: 生きている Excel は殺せないので「待つのをやめる」だけ（CoCancelCall）。VBA が DoEvents 等で
#    応答する状態でだけ効き、`Do: Loop` の密なループには効かない（実測）。マクロ自体は止まらない＝報告にその旨が出る。
#    止めるのは人（Esc／vbe-reset）。

# ショートカットキー付きマクロの一覧 / メニュー表示順の入替
py vba_manager.py list-shortcuts
py vba_manager.py reorder-macro <マクロ名> <up|down|top|bottom|位置番号>   # top/bottom/番号は一発移動

# 導入セルフ診断（初心者が最初に打つ1コマンド。Python/pywin32/Excel/VBOM信頼設定を○×表示）
py vba_manager.py setup-check [--json]

# 環境診断 / VBA構文チェック / プリンタ
py vba_manager.py diag
py vba_manager.py check [excel_file] [--all-warnings]   # ブックの全モジュールを静的診断（.bas 単体は check-bas）
#  診断規則は VBM001〜014（`rules` で一覧）。うち VBM008〜011 は xlflow の117本から
#  輸入した4本＝「走らせても出ない」欠陥だけを選んだもの（状態の戻し忘れ・呼び先が状態を
#  変えたまま返る・モジュール変数の書き手が散る・Find/Replace の引数省略）
#  2026-09-17: VBM012（飛び先ラベルが無い）・VBM013（On Err GoTo＝Error の打ち間違い）・VBM014（存在しないマクロを呼ぶ＝
#  call-graph の未解決）を error に足した＝コンパイル・実行で止まる欠陥を check で先に拾う（ポスター.xlsm の 3 件が素通りしていた）。
#  VBM003（On Error が無い）は既定で件数だけ（ポスターで 190 件中 150 件がこれ）。行も見るなら --all-warnings。JSON は全部残る
py vba_manager.py rules                         # 診断規則の一覧（COM不要・Excelを開かない）

# 自動操縦を止めるGUI境界を、撃つ前に洗い出す（rehearse / gate が黙って固まる場所）
py vba_manager.py inspect-gui [excel_file] [マクロ名] [--module 名] [--json]
#  MsgBox / InputBox / Application.FileDialog / GetOpenFilename / GetSaveAsFilename /
#  Application.Dialogs / モーダルの .Show / SendKeys / Application.Wait を検出。
#  マクロ名を付けるとそこから到達する範囲だけに絞る（ブック全体237件→対象2件のように効く）。
#  文字列リテラルの中と .ShowAllData・.Show vbModeless は数えない

# コマンドの安全性（読むだけ / 書き換える / 消す・止める / 任意コードを走らせる）
py vba_manager.py capabilities [コマンド名] [--json]   # COM不要。破壊操作を撃つ前の照会に

# プロシージャ計量とホットスポット順位（コード行数・分岐数・ネスト深さ）
py vba_manager.py metrics [excel_file] [--top N] [--json]

# Excelプロセスの点呼（PID・表示/非表示・抱えているブック）
py vba_manager.py process [--json]
py vba_manager.py process --kill <PID>          # ブック0冊のプロセスだけ。作業中のExcelは落とさない
py vba_manager.py printer-list
py vba_manager.py printer-setup --printer <プリンタ名> [--duplex ...]
#  ↑ プリンタ名は必ず --printer で渡す（位置引数は無視され、既定プリンタの設定が変わる）。
#     OS のプリンター設定そのものを即時・不可逆に変更し、他アプリにも影響する。

# 全マクロ横断の一括置換（grepの対。diffプレビュー→確認→バックアップ→変更行だけReplaceLine）
py vba_manager.py code-replace "旧" "新" [-y] [--regex] [--module 名]
#  ReplaceLine方式なのでショートカット定義(Attribute)は壊れない。戻すのは restore
#  完全一致が0行のときだけ ①大小文字の違いを無視 → ②空白の違いも無視 の順に探し直し、どの寛容で当てたかを表示する
#  （VBA は保存時に識別子の大小文字をプロジェクト内の初出に揃え、演算子前後の空白も整える＝
#    `.Value` と書いても `.value` で保存されることがある。2026-08-23 に3周空振りした実害の吸収）

# プロシージャのコード取得 → _last_proc.vba に保存
py vba_manager.py get <Sub名>

# モジュール指定してプロシージャ取得（同名プロシージャが複数ある場合に使う）
py vba_manager.py get <モジュール名> <Sub名>
py vba_manager.py get <モジュール名>.<Sub名>   # ドット区切りも可
py vba_manager.py get 名1 名2 名3              # 3個以上は複数取得（1接続・連結保存。書き戻しは1本ずつ）
#  --out f で _last_proc.vba 以外に保存（参照用コピー）、--json も可。見つからないときは近似候補を提示

# _last_proc.vba の内容でプロシージャを置換（バックアップ自動作成）
py vba_manager.py replace-procedure -y
#  ※ -y/--yes で確認プロンプト(y/N)をスキップ。Claude経由の非対話実行では必ず -y を付ける。
#    （-y なしだと input() 待ちで止まる。PowerShellから "y" をパイプで流す方式は失敗することがある）
#    差分(Diff)は -y を付けても表示される。--module <名> で対象モジュールの明示も可。
#    --code-file <f> で _last_proc.vba 以外のコードファイルを指定。差分ゼロなら置換せずスキップ。
#    バックアップが取れないと停止する（--force で強行可）。

# プロシージャの部分行置換（ピンポイント安全修正・0.09秒で超高速・2026-09-17）
py vba_manager.py patch-procedure <モジュール名> <Sub名> --old "置換前コード" --new "置換後コード" -y
#  プロシージャ全体を書き直さず、指示された特定行だけを直接置換（MCP は patch_procedure(module, proc, old, new)）。
#  周辺コードの破壊やリファクタリング事故を防ぎ、置換後の自動構文チェックも内蔵。

# モジュール全体を Remove+Import で置換
py vba_manager.py replace-module <モジュール名> <basファイル>
#  .bas の Attribute VB_Name と指定モジュール名が一致しないと停止（対象取り違え防止）。
#  Import 失敗時はモジュールバックアップから自動復旧を試みる。

# 新規プロシージャの追加 / 削除（get→replace と対称の軽量経路）
py vba_manager.py add-procedure <モジュール名> -y      # _last_proc.vba のコードを末尾に追加（同名重複は停止）
py vba_manager.py delete-procedure [モジュール] <Sub名> [Sub名…] -y   # 削除コードを表示して確認。--module で対象明示。複数は控え 1 冊・保存 1 回

# モジュールを .bas にエクスポート
py vba_manager.py export-module <モジュール名>
py vba_manager.py export-all [--dir 出力先] [--check]  # 全モジュール一括（1接続・--check で検査つき）

# UserForm を「VBAだけで組み立て直す作成マクロ」に変換（別名: フォーム書き出し）
py vba_manager.py form-to-vba <フォーム名>                # → _<フォーム名>_作成マクロ.vba
py vba_manager.py form-to-vba <フォーム名> --add shu002    # 書き出して指定モジュールに追加まで
py vba_manager.py form-to-vba --all --add shu002          # 全 UserForm を一括
py vba_manager.py form-to-vba <フォーム名> --verify        # 捨てブックで組み立てて元と照合（元ブックは無傷）
#  .frx はバイナリでメールの添付検査に弾かれることがある。作成マクロはただのテキストなので、
#  貼って一回実行すれば同じフォームが組み上がる（持ち出しに .frx が要らない）。
#  ※ 表示中のフォームは Designer が取れない。書き出す前に close-form で閉じる。
#  ※ コードの大きいフォームは VBA の1プロシージャ 64KB 上限に収めるため
#    「〜つづき2, 3…」に自動分割される。押すのは1本目だけ（中から順に呼ぶ）。

# 呼び出し台帳の集計（2026-09-16・COM不要）: どのコマンドを何回・平均何秒・最長・失敗、遅かった呼び出し、
#  呼び出しの間（前が終わってから次を撃つまで＝道具の外＝AI が考える・読む・人の時間）。「遅い」はまずこれで内訳を見る
py vba_manager.py stats [--days 7] [--top 20] [--slow 8] [--json]
#  台帳は _calls.jsonl（コマンドを撃つたびに 1 行。batch／agent の中で呼ばれた手は数えない＝外側の 1 回だけ）

# バックアップの一覧と復元（undo導線）
py vba_manager.py list-backups [キーワード]            # 新しい順・COM不要
py vba_manager.py restore <バックアップ.bas>           # VB_Name から対象を特定し replace-module 経路で復元

# マクロ修理の材料を 1 手で（2026-09-17）: get＋call-graph＋compile＋check＋diff-module を撃ち分けない
#  （9/16 のポスター.xlsm はコンパイルエラー 3 件の修理に 9 手・5 分。台帳では道具の中は 4%＝AI の手数を減らすのが効く）
py vba_manager.py repair <マクロ名> [--module 名] [--json]
#  ① 本文（get と同じ。_last_proc.vba にも保存＝replace-procedure がそのまま使える）
#  ② 入口（ショートカット・シートのボタン・フォームのボタン・自動実行）③ 呼び元・呼び先（間接 1 段）
#  ④ 全体コンパイル（落ちた行を [モジュール] Sub:行: 本文 で名指し）⑤ check の error 級（VBM004/005/012/013/014）でこのプロシージャに当たるもの
#  ⑥ 直前の控え（backups/<ブック>_<モジュール>_<日時>.bas）との、このプロシージャの範囲だけの差分。agent の macro モードも同じ材料を読む

# 系譜（2026-09-17・COM 不要・oletools で読む＝Excel を開かない）
py vba_manager.py versions <ブック名> [--dir 追加フォルダ]   # Desktop／プロジェクト／保存／エクセル／公開リポ／backups の同名ブックを日時順に
#  1 冊 1 行＝日時・大きさ・モジュール数・プロシージャ数・最新との差（変わった／無い／この写しだけのモジュール名）・[開いている]。
#  同名が 4 冊あって「ポスター - コピー」に直しが入っていた 9/16 の取り違えを 0 にする手。VBAM_VERSION_DIRS（; 区切り）でも場所を足せる
py vba_manager.py history <マクロ名> [--book ブック名] [--deep] [--max N]
#  backups の *.bas/*.frm/*.cls と _exports/**（--deep で写しの .xlsm）からその Sub を集め、同じ本文は畳み、
#  古い順に「いつ・どのファイル・何行」と隣との差分。日時はファイル名の YYYYMMDD_HHMMSS → 無ければ mtime

# 閉じたブックを読む（2026-09-17・Excel を開かない＝Workbook_Open もアドインも起きない。50 冊でも 1 手）
py vba_manager.py list-file <path.xlsm> [--json]            # モジュール（種別・行数）とプロシージャ名。.xlsx は「マクロ無し」
py vba_manager.py grep-files "文字" <フォルダ|ファイル…> [--regex] [-i] [--max N]   # [ファイル名][モジュール] プロシージャ:行: 本文（行番号は VBE と同じ）
py vba_manager.py export-file <path.xlsm> [モジュール…] [--dir 先]   # .bas/.cls/.frm に書き出す（モジュール名を続けるとそれだけ）（既定 _exports/<ブック>/<日時>/。.frx は無い）
#  .xlsb も読める（vbaProject.bin）。フォームのコードは .frm として出るがレイアウト（.frx）は付かない

# コマンド列を1接続で連続実行（一括作業の高速化。実測: 18本export 数分→約2秒級）
py vba_manager.py batch cmds.txt                       # 1行=1コマンド。#始まりと空行は無視
printf 'get shu003 マクロA\nreplace-procedure -y\n' | py vba_manager.py batch -   # 標準入力からも可
#  途中失敗で停止（--keep-going で最後まで実行）
```

### 目コマンド（シート状態の読み取り・読み取り専用）

```bash
py vba_manager.py read-range     [excel_file] [range...]  # セル値をテキスト格子で読む（複数範囲可・1接続）
py vba_manager.py read-range     A1:D10 --formula        # 計算結果でなく数式(.Formula)を表示
py vba_manager.py read-range     "集計!A1:D50" --tsv     # _last_values.tsv に書き出し→編集→write-range で書き戻す往復
py vba_manager.py read-range     A1:D10 --width 80       # 列の表示幅を広げる（既定40、切り詰めは…付き）。--json も可
py vba_manager.py read-selection [excel_file] [--formula] # 今選択している範囲を読む（--formulaで数式）
py vba_manager.py sheet-info     [excel_file] [--preview 3] # シート構成一覧（--preview で各シート先頭N行も＝ブック俯瞰が1接続）
py vba_manager.py materials      [excel_file] [シート] [--rows 5]  # 先回り材料＝1シートの使用範囲・値・結合・テーブル・名前定義・数式の型（式は A1 形式＋代表の番地。R1C1 の相対番号を数え直さない・2026-09-08）・エラー・図形ボタン・###・列幅を1回で（手を動かす前に見る・2026-09-02）。40行×30列までの表は全体＋長文セルの全文。**気づき（数式）＝列の中で式の形が違うセル・式に直書きされた率や単価・集計の起点行の食い違いを、道具が全部の式について数えて出す（2026-09-09。式の点検はここが出発点。read で全部読み直さない）**。仕事の時計を押す（tidy／write が経過秒を出す）。式・非表示・結合は一括で読む（2026-09-16。それまで数式セル 1 つにつき COM 3 回＝数式の多い表で数十秒）
py vba_manager.py table read <テーブル名> [--tsv]        # テーブル名で直接読む（番地調べ→read-range の2段を1段に）
py vba_manager.py screenshot     [excel_file] [range] [--out f.png]  # 範囲を画像(PNG)で書き出す（省略時 _last_view.png）
py vba_manager.py snapshot       [excel_file] [--sheet 名] [--out f.json] [--max-rows N]
                                 [--format-rows N] [--no-format]
#  ブック(または1シート)を意味構造JSONに畳む（セル疎+結合+書式+図形/ボタン(OnAction)+テーブル）
#  → _last_snapshot.json。機械的事実だけ吐き意味付けはAIがやる＝このJSONを Read して
#  質問に答えれば「開いたままブックLM」になる
#  書式（四つ目の目）: フォント/太字/文字色/塗り/罫線/表示形式/横位置/列幅/行高。
#    シート全体→列→行の順に「揃っているか」で降りて採る（セル単位は1万セルで33秒＝実用外）。
#    値 None は「その範囲の中で不揃い」の意味（欠測ではない）。読めなければ format_error を残す。
#    行ごとにCOM往復が要るので既定2000行まで（--format-rows）。急ぐときは --no-format
py vba_manager.py snapshot-diff  <before.json> [after.json] [--max N]
#  2つの snapshot の機械的差分（COM不要）。after 省略時は _last_snapshot.json（直近のsnapshot）と比較。
#  セル変更/追加/削除・結合・書式・図形（文字/OnAction/位置/大きさ）・テーブルをシート別に列挙
#  ★クリーン化系マクロは値も結合も図形も残したまま書式だけを吹き飛ばす。書式の目が無いと
#    これを「差分なし（一致）」と報告する（ポスター.xlsm 破壊の教訓）。
#  シート全体で起きた変化は列・行では繰り返さない（局所的な変化だけが個別に出る）
py vba_manager.py wiring         [excel_file] [--json]    # 別名: 配線図
#  ボタン⇔マクロの配線図＝シート上の全 OnAction（グループ内も展開）をマクロ名簿と突き合わせ、
#  行き先のないボタン（存在しないマクロを指す壊れた配線）を検出。孤立マクロ側は call-graph で
```

> **screenshot の注意**: 内部で一時グラフの作成・削除を伴うため、目コマンドの中で唯一
> **ブックの Undo 履歴が消える**（内容は変えない）。実行後は選択を A1 に戻しクリップボードもクリアする。

> **数式の修正ループ**: `read-range --formula` で今の式を読む → 直した式を `write-range`
> で書き戻す。読み(.Formula)も書き(.Value)も US規約（英語関数名・カンマ区切り）でそろえてあるので往復が一致する。

### 手コマンド（シートの編集・整形・構造操作）

Excel MCP 相当の編集機能を「**開いたままのアクティブブック**」に対して COM で直接行う。
Excel MCP はファイルを閉じる必要があるが、これらは開いたまま使える代わりに
**プログラム経由の変更は Excel の Undo 履歴を消す**。
**既定では保存しない**（Excelで確認後に手動保存、または保存せず閉じれば変更を破棄できる＝Undo代わり）。

```bash
# 値・数式を書き込む（'='始まりは数式）。単一値はインライン、グリッドは TSV。
py vba_manager.py write-range A1 "こんにちは"
py vba_manager.py write-range C1 "=SUM(A1:A10)"
py vba_manager.py write-range A3 --tsv data.tsv     # 省略時は _last_values.tsv（タブ区切り）
#  MCP からは write_grid(range, tsv) ＝ TSV を文字列で直接渡す（ファイル不要・--show 常時）
py vba_manager.py write-range A1 "007" --raw        # 数値変換せず文字列で書く（先頭ゼロ保持）
#  ※ シート名だけの指定（使用範囲全域）は write-range では拒否される。セル/範囲を明示すること。
#  ※ 文字列を Excel が日付・数値に読み替えたら（"1-2"→1月2日、"1,000"→1000）道具が文字列に書き直して
#    報告する（write-range / write-cells 共通・2026-09-02）。--raw は "007" のような数値化そのものを止めたいときだけ
# 飛び飛びのセルを1回で（2026-09-02 夜）: 「セル 値」の組を並べる。'=' 始まりは数式。化けは道具が戻す（--raw 不要）
py vba_manager.py write-cells C7 000-0002 C11 000-0006 D7 "新しい住所" --show
py vba_manager.py write-cells --tsv cells.tsv       # 「セル<TAB>値」を1行1セル（値にスペース・引用符があるとき）

# 範囲をクリア（既定すべて／--contents 値のみ／--formats 書式のみ）
py vba_manager.py clear-range A1:D10 --contents
#  ※ シート名だけの指定（使用範囲全域）は --whole-sheet を付けない限り拒否される。
#    余分な位置引数もエラーになる（"Sheet1 A1:B2" と分けて渡す事故の防止）。fill / sort も同様。

# 書式・整形（複数オプション同時指定可）
py vba_manager.py format-range A1:D1 --bold --bg "#FFFF00" --color "#FF0000" --align center --border thin
#  指定可: --font 名 --size N --bold --unbold --italic --color/--bg '#RRGGBB'
#          --number-format 書式 --align left|center|right --valign top|center|bottom
#          --wrap --border thin|medium|thick|hairline|none --col-width N --row-height N
#          --merge --unmerge --autofit

# 表を整える（仕上げ・2026-09-02）: 見出し書式（--header-from で複写／省略時は表の左上が見出しの体裁なら
#  それを行全体へ複写、そうでなければ太字+#DDEBF7）・罫線（細・格子）・列の型に合わせた寄せと表示形式
#  （見出しが番号/コード/No/ID の列は左寄せ、数値だけの列は #,##0・小数は #,##0.00。年・日付・率の列と
#  既に書式のある列は触らない・同日深夜）・列幅の自動調整（6〜60）を当て、
#  最後に見え方（画面の文字・###）と materials からの経過秒を出す
#  見出しの複写は表示形式を潰さない（2026-09-08）: 先頭行に数値・日付があれば、その表示形式は複写後に戻す
#   ＝集計ブロック（左が「合計」・右が 203,250）や年を見出しにした表でカンマ・桁が消えない
#  列幅は表の外まで細くしない（2026-09-08）: 幅は列全体に効くので、細くした結果 表の上下のセルが '###' に
#   なったら、その列だけ元の幅へ戻す（戻した列は「整えました」の行に出る）
# 表の仕上げ検査（見た目＝見出し・罫線・列の型・列幅・空の見出し／中身＝エラー値・「合計」行が上の和と合わない・
#   数式の列に混ざる直値・**集計の式が本文の行を範囲に入れていない**（=SUM(C2:C7) の下に本文の C8 がある形。
#   値が今は合っていても、その行を埋めた瞬間に狂う・2026-09-09）・**横に並ぶ同じ関数の集計が列ごとに違う式**
#   （月や区分を式に埋めた集計表＝列を 1 つ足すと壊れる。値は合っているので値の検査では出ない・2026-09-09）。指摘0件で合格。
#   表の中の空欄・数値列の文字の数字・番号列の重複・列の中で式の形が違うセル・式に直書きされた数値は
#   「気づき」で出す＝合否に入れない）
py vba_manager.py audit A1:E9              # セル1つを渡すとその表（CurrentRegion）全体

# 表・シートの数式＆データ総合診断（2026-09-17）。materials／seiri が循環参照（要修正）と外れ値（気づき）を出す＝1 手目には要らない。
#  全部の一覧・採点・--fix が要るときだけ 2 手目に撃つ
py vba_manager.py diagnose [excel_file] [範囲|セル] [--sheet シート名] [--json] [--fix] [--html 出力先.html]
#  別名: データ診断 / formula-lint（MCP は diagnose(range, sheet, fix, html)）
#  わずか 0.015秒で数千セルの表を完全診断し、100点満点スコアと重要度別レポートを出力。
#  ■ 数式静的解析:
#    - 循環参照（Circular Reference）の静的検出（有向グラフ閉路解析。Excelフリーズ防止）
#    - 縦・横の SUM/AVERAGE 集計範囲漏れ（直前・直後の数値セル漏れ）
#    - ハードコード定数混入（*1.1, +500 等）
#    - 非一貫数式（列内で一部だけ計算式が異なるコピペミス・手入力上書き）
#    - 数式エラー（#REF!, #DIV/0!, #VALUE!, #N/A）
#  ■ データクレンジング＆統計:
#    - 隠れ空白（前後の全半角スペース・改行コード。VLOOKUP破壊の元）
#    - 文字列型数字（'123 等。SUMから除外されるリスク）
#    - 統計的外れ値（IQR法による桁間違い・マイナス誤入力）
#  ■ --fix（人の値を書き換える手＝依頼に承認の語があるときだけ。-y 必須・2026-09-17）:
#    書き戻すのは 隠れ空白の除去 と 文字列の数字の数値化 の 2 種類だけ。先頭ゼロ（0123）・16 桁以上・数式のセルは触らない。
#    集計漏れの式は「直し方の提案」を並べるだけで書かない（本文か別の塊かは判断＝人か AI）。撃つ前に backups へ写しを取る。
#    列内で形の違う式は warning（気づき）。agent の done を止めるのは critical（循環参照・エラー値・集計漏れ）だけ。
#  ■ HTML品質証明書 (--html):
#    SVG円形プログレスゲージ、重要度別カラーバッジ、モダンCSSを備えた自己完結型ダッシュボードを出力。
#    上司やクライアントへの品質納品物・証明書としてそのまま提出可能。

# clean-vba（9/16 夜・Gemini）は 9/17 に check へ畳んだ＝コマンドは無い。
#  On Error Resume Next の行＝VBM015・150 行超＝VBM016（warning）。戻し忘れは VBM008 が前から見る。Option Explicit は数えない（原則 7）。

# 表・シートの視覚レイアウト＆書式マップ（2026-09-17）。materials／seiri が要約（見出しの結合の親子・二重下線の合計行・
#  塗りと文字色の有無）を出すので 1 手目には要らない。セルごとの番地と色が要るときだけ 2 手目に撃つ
py vba_manager.py style-map [excel_file] [範囲|セル] [--sheet シート名] [--json]
#  別名: 書式マップ（MCP は style_map(range, sheet)）
#  セルの文字だけでなく「人間の肉眼で見えるデザイン構造」を AI に伝える（所要時間 0.2秒）。
#  ■ 複合見出し（結合セル）の階層ツリー解読（2段組・3段組の親子関係を展開）
#  ■ 合計行・集計行の二重下線（Double Bottom Border）自動検知
#  ■ 背景ハイライト色（#RRGGBB）、フォント、太字、赤字警告セルの抽出
# 並べ替えは、隠れた行を出してから並べる（2026-09-08）: Excel の並べ替えは見えている行しか動かさない＝
#  隠れた行はその場に取り残されて並びが黙って狂う。絞り込みも手で隠した行も出してから並べ、出したことを報告する
#  （並べた後に隠し直さない＝行が動いた後の行番号は別の人のデータを指す）
py vba_manager.py tidy A5:G13 I5:L11                 # 自分が書いた・広げた表の全体（列を足したら足した列だけでなく表全体）。範囲は複数可。セル1つ（tidy A5）なら CurrentRegion＝表全体
#  --header-from セル --bg '#RRGGBB' --no-header --no-border --no-col-format --no-autofit --min-width N --max-width N
# 書いた直後の見え方を読み戻す --show（write-range / write-cells / fill / copy-range / format-range 共通）
py vba_manager.py write-range F6 "=A6*2" --show

# シートを設計図から組み上げる（2026-09-03・「作る」側の手1。フォームを宣言から組むのと同じ型をシートに）
py vba_manager.py build-sheet --sample            # 設計図（JSON）の例を _last_sheet_spec.json に書き出す
py vba_manager.py build-sheet --dry-run           # 計画（番地・数式・仕上げ）だけ表示。Excel に触らない
py vba_manager.py build-sheet [設計図.json] [--overwrite]   # 省略時は _last_sheet_spec.json。同名シートに中身があれば止まる
#  設計図: sheet／title・note（1・2 行目。あれば見出しは 3 行目）／columns（name・type=text|int|number|
#  decimal|currency|percent|date・format・width・align・formula・list・default）／rows か data／
#  total（label・sum・label_col）／freeze／filter／table／names（range=table|data|full|列名|番地）／
#  cond_format（column・gt|lt|ge|le|eq|ne|between・formula・bg・color・bold）／print（area・title_rows・
#  landscape・fit_wide・fit_tall・center_h）／header_bg／overwrite
#  数式は {列名}（同じ行のそのセル）・{row}・{first}・{last} で書く（例 "=IF({数量}=\"\",\"\",{数量}*{単価})"）。
#  流れ: 列の表示形式 → 見出し → 格子（数式は実番地に展開）→ 合計行 → tidy → 幅・寄せ → 枠固定 →
#  入力規則 → テーブル/オートフィルタ → 名前定義 → 条件付き書式 → 印刷設定 → エラーセル検査（保存はしない）
#  組み上げの最後に「検査: エラーセル / ### / 数式の本数 / 画像 _last_build.png → 合格|要確認」を出す（手3）
py vba_manager.py build-sheet --ask "会員の出欠名簿。会員番号・氏名・区分（正会員/準会員）・出欠・会費。15行。会費の合計行" --new-book
#  --ask 依頼文 ＝ AI に設計図を 1 回書かせてから組み上げる（手2。ループは道具、モデルは設計図を 1 回書くだけ）。
#    送った文は _last_sheet_ask.txt、返った設計図は _last_sheet_spec.json に残る（検分・手直し→build-sheet で再組み上げ）。
#    「請求書: 送り字数 / 返り字数 / 待ち秒 / トークン 入力・出力（・思考） / モデル」を必ず出す。
#    --ai claude-code（既定・ヘッドレスの Claude Code＝claude -p を stream-json で温める・鍵なし・sonnet）
#         | gemini（GEMINI_API_KEY・gemini-3.7-flash）| claude（ANTHROPIC_API_KEY・claude-haiku-4-5-20251001）、--model で変更
#         （--ai を省いても --model が gemini で始まる名前なら gemini。claude-code は道具を持たせず材料は文で渡す＝他の頭と同じ土俵）
#  --new-book ＝ まっさらな新しいブックに組み上げる（訓練場。人のブックに触らない）。--dry-run と併用可
py vba_manager.py build-sheet --recipes                      # 手順書（帳票の型）の一覧: 名簿・備品台帳・月次集計表・請求明細・出勤簿
py vba_manager.py build-sheet --recipe 請求明細 --new-book     # 手順書の名前で組む（設計図は _last_sheet_spec.json にも置く＝手直しの起点）
py vba_manager.py build-sheet --fire [名前 ...]               # 自動実射: まっさらなブックに全部組んで検査→PASS/FAIL→保存せず閉じる（手順書を足したら必ず）
#  合計行の下に行を続ける: total.extra [{"label":"消費税（10%）","column":"金額","formula":"=ROUNDDOWN({total:金額}*0.1,0)"},
#    {"label":"税込合計","column":"金額","formula":"={total:金額}+{above}"}]。type には time（h:mm・合計は format "[h]:mm"）もある
#  AI は設計図を書くだけ。組み立てるのも確かめるのも道具。手順書は RECIPES（vbam_build.py）に足す＝足したら --fire

# 依頼文 1 つを道具が回す薄い入口 agent … 全文は references/agent.md（この段の続きは agent.md に切り出した）


# シート操作
py vba_manager.py sheet add 新シート [--before 既存 | --after 既存]
py vba_manager.py sheet rename 旧 新
py vba_manager.py sheet copy 元 [新名]
py vba_manager.py sheet delete 名 / activate 名 / show 名 / hide 名 / very-hide 名
py vba_manager.py sheet visibility 名                         # 表示状態を表示
py vba_manager.py sheet tab-color 名 "#FF8800"               # タブ色（#RRGGBB / R G B / --clear / 引数なしで取得）

# テーブル(ListObject)
py vba_manager.py table create Sheet1!A1:D10 売上表 [--no-headers]
py vba_manager.py table list
py vba_manager.py table delete 売上表     # テーブル解除（データは残る）
py vba_manager.py table ref 売上表 [列名]  # 構造化参照の書き方と列一覧を表示
# 列・フィルタ・ソート（excel-mcp の table_column 相当）
py vba_manager.py table column add 売上表 単価 [--at 3]       # add/remove/rename/format
py vba_manager.py table column rename 売上表 単価 税込単価
py vba_manager.py table column format 売上表 数量 "#,##0"     # 引数なしで取得
py vba_manager.py table filter 売上表 数量 ">8"              # 条件フィルタ
py vba_manager.py table filter-values 売上表 商品 りんご ぶどう # 値フィルタ
py vba_manager.py table filters 売上表 / filter-clear 売上表
py vba_manager.py table sort 売上表 数量 [--desc]
py vba_manager.py table sort-multi 売上表 商品:asc 数量:desc

# 名前付き範囲
py vba_manager.py name add 基準値 Sheet1!A2
py vba_manager.py name list
py vba_manager.py name delete 基準値
```

#### 手コマンド 第2弾（編集の足回り／検索置換／保存印刷／仕上げ）

```bash
# --- a. 編集の足回り ---
py vba_manager.py row insert 5 2          # 5行目に2行挿入 / row delete 5 2（--sheet 名 で対象明示可）
py vba_manager.py col insert C 1          # C列に1列挿入 / col delete C 1
py vba_manager.py copy-range A1:C1 E1     # 範囲コピー（--values で値のみ）
py vba_manager.py fill D2:D5              # 先頭セルを下にフィル（--right で右）
py vba_manager.py sort A1:C20 --key B --desc --header   # B列キーで降順、見出しあり
py vba_manager.py autofilter A1:C20       # オートフィルタ設定（--off で解除）

# --- b. 検索・置換 ---
py vba_manager.py find 田中 --book        # 全シート横断で検索（--whole/--formula/--max N）
#  結果は「シート名!$A$1: 値」形式＝そのまま次コマンドの range 引数に貼れる
py vba_manager.py find-replace 旧 新 A1:Z99   # 範囲一括置換（範囲省略で使用範囲全体）
#  ヒットセル数を事前カウントして報告、0件なら置換しない。--match-case で大小文字区別。

# --- c. ブックの開閉・保存・印刷まわり ---
py vba_manager.py open "C:\path\book.xlsm"     # ブックを開く（2026-07-15）
#  「人が開くのと同じ場所」に開く: 既に開いていれば前面化のみ／見えているExcelが居れば合流／
#  Excel未起動なら通常起動（アドイン・PERSONALが普段どおり読み込まれる）。
#  非表示の残骸Excelしか居ないときは取り込まれないよう開かずに止まる。
py vba_manager.py close <ブック名> --no-save -y   # ブックを1冊閉じる（2026-07-15）
#  鎧の三点セット: ①ブック名指し必須 ②--save/--no-save の明示必須 ③確認プロンプト（-y でスキップ）。
#  Excel本体は終了しない。PERSONAL.XLSB とアドインブックは閉じない。同名複数はフルパス名指しを要求。
py vba_manager.py save                    # 上書き保存（手コマンドはOK出たらこれで確定）
py vba_manager.py save-as "C:\path\out.xlsx"   # 別名保存（拡張子で形式判定）
#  既存ファイルへは --overwrite が無いと停止。未対応拡張子はエラー。
#  xlsm→xlsx はマクロが落ちる旨を警告。[excel_file] を第1引数に取る流儀も他コマンドと同じ。
py vba_manager.py export-pdf 出力.pdf [--sheet 名 | --range "集計!A1:H50"] [--overwrite]  # PDF出力（ブックは変更しない）
py vba_manager.py print-setup --area A1:H50 --title-rows 1:3 --landscape --fit-wide 1
#  指定可: --area --title-rows 1:3 --title-cols A:B --landscape/--portrait
#          --fit-wide N --fit-tall N --zoom N --center-h --center-v

# --- d. 仕上げ・見た目 ---
py vba_manager.py cond-format B2:B20 --gt 85 --bg "#FFC7CE"   # 85超を赤に（--clearで全削除）
#  比較: --gt --lt --ge --le --eq --ne 値 / --between v1 v2、色: --bg --color、--bold
#  数式ルール: --formula "=B2>AVERAGE($B$2:$B$20)"（相対参照は範囲左上セル基準）
py vba_manager.py hyperlink A1 "https://..." --text "リンク"   # --remove で削除
py vba_manager.py hyperlink A1                                # そのセルのリンクを表示
py vba_manager.py hyperlink --list [シート名]                  # シート内の全リンク一覧
py vba_manager.py sheet protect シート名 [--password p]       # シート保護 / unprotect で解除
py vba_manager.py format-range A1:B5 --unlock                 # 保護中も編集可のセルに（--lockで戻す）
py vba_manager.py validation C2:C20 --list "A,B,C"            # ドロップダウン（--clearで削除）
py vba_manager.py freeze B2               # B2の左上で枠固定 / freeze off
py vba_manager.py comment A1 "見出し"     # セルコメント / comment A1 --remove
```

> **--append（ログ追記）**: `write-range "ログ!A" 値 --append` で使用範囲の最終行の次に書く
> （書き込み先番地は実行結果に必ず表示される）。
> **--sheet 分離指定**: read-range / write-range / clear-range / format-range / fill / sort /
> find-replace は `--sheet シート名 A1:B2` の分離指定も可（`!` や記号を含むシート名のクォート地獄の救済）。
> row / col の `--sheet` と同趣旨。

#### 重量級コマンド

```bash
# (1) チャート  ※ 重量級は一通り実装済（PowerQuery の load 読み込み配線まで完了）
py vba_manager.py chart create A1:B5 --type column --title "月別売上" --name 売上グラフ
#  --type column|bar|line|pie|scatter|area|doughnut（既定 column）
#  --at セル（左上を合わせる／省略時はデータ範囲の右隣）--width N --height N
py vba_manager.py chart list
py vba_manager.py chart delete 売上グラフ
# (1b) グラフ詳細設定 chart-config（excel-mcp の chart_config 相当）
py vba_manager.py chart-config set-title 売上グラフ "月別売上"
py vba_manager.py chart-config set-type 売上グラフ line
py vba_manager.py chart-config set-axis-title 売上グラフ value 数量
py vba_manager.py chart-config legend 売上グラフ right            # bottom/top/right/left/corner/off
py vba_manager.py chart-config style 売上グラフ 5                 # 組込スタイル 1-48
py vba_manager.py chart-config axis-scale 売上グラフ value --min 0 --max 40 --major 10
py vba_manager.py chart-config gridlines 売上グラフ value --major on --minor off
py vba_manager.py chart-config axis-format 売上グラフ value "#,##0"   # 引数なしで取得
py vba_manager.py chart-config data-labels 売上グラフ --value --percent [--position outsideend]
py vba_manager.py chart-config placement 売上グラフ 3            # 1=移動+サイズ/2=移動のみ/3=自由
py vba_manager.py chart-config add-series 売上グラフ C2:C4 --series-name 在庫 [--category-range A2:A4]
py vba_manager.py chart-config set-source 売上グラフ A1:C10       # データ範囲の再設定
py vba_manager.py chart-config remove-series 売上グラフ 2          # index 必須（省略はエラー）
py vba_manager.py chart-config series-format 売上グラフ 1 --marker-style 2 --marker-size 8 --marker-bg "#FF0000" [--invert]
py vba_manager.py chart-config trendline add 売上グラフ 1 linear   # list/add/delete、種別 linear/exponential/logarithmic/movingaverage/polynomial/power

# (2) ピボットテーブル
py vba_manager.py pivot create 元データ!A1:C100 --rows 部門 --cols 月 --values 売上 --func sum --sheet 集計 --name 売上ピボット
#  --rows/--cols/--filter/--values はカンマ区切りで複数可（同じ列を 2 本 --values 売上,売上 も可）、--func sum|count|average|max|min（既定 sum）
#  出力先: --sheet シート名（無ければ作成。--at と組めばそのシートのそのセル）/ --at セル（元データと同じシートの空き場所）/ 省略時は新規シート「ピボット」
#  既存のデータの上や隣（右下 40 行×20 列に何かある所）には置かない＝Excel の「上書きしますか」の窓で止まる前に道具が断る
py vba_manager.py pivot create model --rows T売上.地域 --measures 平均単価 --sheet 単価 --name P単価   # データモデルから（テーブル名.列名／メジャー名。--values T売上.金額 は暗黙のメジャー）
py vba_manager.py pivot list
py vba_manager.py pivot delete 売上ピボット
# (2b) フィールド管理 pivot-field（excel-mcp の pivottable_field 相当）
py vba_manager.py pivot-field list 売上ピボット
py vba_manager.py pivot-field add-row|add-col|add-filter 売上ピボット 部門
py vba_manager.py pivot-field add-value 売上ピボット 売上 --func sum [--name 表示名]
py vba_manager.py pivot-field remove 売上ピボット 部門
py vba_manager.py pivot-field set-func 売上ピボット 売上 average     # 値フィールドの集計
py vba_manager.py pivot-field set-name 売上ピボット 部門 部署        # 表示名(Caption)
py vba_manager.py pivot-field set-format 売上ピボット 売上 "#,##0"
py vba_manager.py pivot-field set-filter 売上ピボット 部門 営業 開発  # 表示する値を限定
py vba_manager.py pivot-field sort 売上ピボット 部門 desc
py vba_manager.py pivot-field group-date 売上ピボット 日付 months    # days/months/quarters/years
py vba_manager.py pivot-field group-numeric 売上ピボット 金額 0 1000 100
py vba_manager.py pivot-field show-as 売上ピボット sum/売上2 percent_total        # 計算の種類（値フィールドの設定）。構成比・前月比は同じ列を 2 本値に置いて片方に当てる
#   種類: percent_total|percent_row|percent_col|percent_parent_row|percent_parent_col|percent_of|diff_from|percent_diff_from|
#         running_total|percent_running_total|rank_asc|rank_desc|index|none。基準は --base-field（既定＝先頭の行フィールド）--base-item（既定＝前の値）
py vba_manager.py pivot-field top-n 売上ピボット 部門 3 [--by sum/売上] [--bottom]   # 上位/下位 N 件（xlTopCount=1・31 は日付の種類＝E_INVALIDARG）
py vba_manager.py pivot-field filter-clear 売上ピボット 部門
py vba_manager.py pivot-field position 売上ピボット 部門 1                         # 同じ区画の中の順番
# (2c) 計算フィールド・レイアウト pivot-calc（excel-mcp の pivottable_calc 相当）
py vba_manager.py pivot-calc get-data 売上ピボット                   # 出力範囲の値を表示
py vba_manager.py pivot-calc calc-field create 売上ピボット 利益 "=売上-原価"   # その後 add-value で値に
py vba_manager.py pivot-calc calc-field list|delete 売上ピボット [名前]
py vba_manager.py pivot-calc layout 売上ピボット tabular            # compact/tabular/outline
py vba_manager.py pivot-calc subtotals 売上ピボット 部門 off
py vba_manager.py pivot-calc grand-totals 売上ピボット both off      # rows/cols/both
py vba_manager.py pivot-calc refresh 売上ピボット                    # 元データが変わったとき
py vba_manager.py pivot-calc style 売上ピボット PivotStyleMedium9
py vba_manager.py pivot-calc repeat-labels 売上ピボット on           # tabular と組む
py vba_manager.py pivot-calc empty-as 売上ピボット 0                 # 空白セルの表示
py vba_manager.py pivot-calc set-source 売上ピボット 元データ!A1:C200  # 行が増えたとき
py vba_manager.py chart create --pivot 売上ピボット --type column --title "部門別 売上（円）" --at H10   # ピボットグラフ
# データモデルのピボット: powerquery load --to model で載せた列は、M に型付けが無いと全部「文字」になり SUM 系のメジャーが
#   無言で消える（2026-09-04 実射）。powerquery add/edit はテーブルを読むだけの M に型付けを自動で足す。datamodel list に列の型が出る

# (3) スライサー（ピボット名 or テーブル名に紐づけ）
py vba_manager.py slicer add 売上ピボット 部門 --name 部門スライサー --at H1
py vba_manager.py slicer list
py vba_manager.py slicer delete 部門スライサー
```

#### 計算モード（大量書き込みの高速化）

```bash
py vba_manager.py calc-mode             # 現在のモードを表示
py vba_manager.py calc-mode manual      # 手動に（大量 write-range の前に）
py vba_manager.py calc-mode recalc      # 今すぐ一括再計算
py vba_manager.py calc-mode auto        # 自動に戻す（automatic も可。recalc は now/calculate も可）
```

#### PowerQuery（一覧・更新・作成・M式書換・削除・読み込み配線）

```bash
py vba_manager.py powerquery list           # クエリ（M式行数・説明）と接続の一覧
py vba_manager.py powerquery refresh         # 全クエリ/接続を更新（RefreshAll）
py vba_manager.py powerquery refresh 売上    # 指定クエリ/接続を更新
py vba_manager.py powerquery add 商品マスタ          # _last_query.m のM式から新規作成（接続のみ、--desc 説明 も可）
py vba_manager.py powerquery add 数列 --m "let S=#table(...) in S"   # インラインM式
py vba_manager.py powerquery edit 商品マスタ          # _last_query.m のM式で既存クエリを書き換え
py vba_manager.py powerquery delete 数列     # クエリ削除
py vba_manager.py powerquery load 商品マスタ --to sheet            # シートにテーブルとして読み込み（アクティブシートA1）
py vba_manager.py powerquery load 商品マスタ --to sheet --sheet 一覧 --at B2  # 出力先シート・左上セル指定
py vba_manager.py powerquery load 商品マスタ --to model            # データモデル(Power Pivot)に読み込み
```

> add/edit とも M式は `_last_query.m`(UTF-8) / `--m-file f` / `--m "..."` から取得。
> add は **接続のみ**のクエリを作る。edit は `WorkbookQuery.Formula` への代入（書込可・検証済）。
> **load（読み込み配線・実装済/検証済）**: 接続のみクエリをシートのテーブルまたはデータモデルに読み込む。
> - `--to sheet`: `ListObjects.Add(xlSrcExternal)` でシートに表として出す。作られる接続は
>   `Query - <名前>` に揃えるので `powerquery refresh <名前>` で更新できる。`--sheet`/`--at` で出力先指定。
> - `--to model`: `Connections.Add2(..., CommandText=クエリ名, lCmdtype=6=xlCmdTableCollection, CreateModelConnection=True)`。
>   この形だとモデルテーブル名がクエリ名になる（SQL/SELECT形式だと "クエリ" の汎用名になるので不可）。
>   一度モデルを使うと空でも `ThisWorkbookDataModel` 接続が残るが無害。`datamodel list` で確認。
> いずれも **保存はしない**（Excelで確認後に手動保存、破棄したいなら保存せず閉じる）。

#### コネクション / データモデル

```bash
# ブック接続（外部データ・クエリ接続）
py vba_manager.py connection list            # 種別・接続文字列つき一覧
py vba_manager.py connection refresh 顧客接続  # 更新（name省略で全件RefreshAll）
py vba_manager.py connection delete 顧客接続   # 削除

# データモデル（Power Pivot）
py vba_manager.py datamodel list             # テーブル(行数)・リレーション・メジャーを一覧
py vba_manager.py datamodel relation add    売上 商品ID 商品マスタ ID  # リレーション作成
py vba_manager.py datamodel relation delete 売上 商品ID 商品マスタ ID  # リレーション削除
py vba_manager.py datamodel measure add 売上 合計数量 --dax "SUM('売上'[数量])"  # メジャー(DAX)作成
py vba_manager.py datamodel measure add 売上 平均数量            # DAXは _last_dax.dax(UTF-8) から（--desc 説明 も可）
py vba_manager.py datamodel measure add 売上 売上額 --dax "SUM('売上'[数量])" --format currency --symbol JPY --decimals 0
py vba_manager.py datamodel measure add 売上 構成比 --dax "DIVIDE(...)" --format percent --decimals 1
py vba_manager.py datamodel measure delete 合計数量             # メジャー削除
```

> **datamodel relation add/delete（実装済/検証済）**: `ModelRelationships.Add(FK列, PK列)` で作成。
> 引数順は **FKテーブル FK列 PKテーブル PK列**（FK=多側/参照する側、PK=一側/参照される側）。
> PK側の列は値が一意であること。テーブル自体の追加は `powerquery load --to model`。
>
> **datamodel measure add/delete（実装済/検証済）**: `ModelMeasures.Add(名前, テーブル, DAX, 書式, 説明)`。
> DAX は `--dax` / `--dax-file` / `_last_dax.dax`(UTF-8) から取得。**先頭の `=` は自動で外す**（付けても可）。
> **DAX 内の日本語テーブル名はシングルクォート必須**（例: `SUM('売上'[数量])`／`売上[...]` は構文エラー）。
> 書式は `--format general|whole|decimal|currency|percent|scientific`（既定 general）。
> `--decimals N`（小数桁）、`--thousands`（桁区切り）、`--symbol`（**通貨コード** USD/JPY/EUR。`$`等のグリフは不可、無効時は既定にフォールバック）。
> 引数付き書式は `GetModelFormat*` メソッド経由（`ModelFormat*` プロパティは既定値専用で引数を渡せない）。

> **write-range のグリッド入力**: `_last_values.tsv`（UTF-8・タブ区切り）に
> 行＝改行、列＝タブで書き出してから `write-range <左上セル>` で適用する。
> `get` → `_last_proc.vba` と同じ発想。`'='` 始まりのセルは数式として書き込まれる。
