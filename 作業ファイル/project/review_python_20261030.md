# Python の道具の総点検（2026-10-30・報告だけ）

- 対象: `作業ファイル/project/python_scripts/` の `vbam_*.py`・`vba_manager.py`・`vba_mcp_server.py`（main e60ac80 時点）
- やり方: Excel の無いクラウドでコードを読んだだけ。COM は一度も呼んでいない。道具の .py は直していない
- 見た型: 例外の握りつぶしによる誤報／確認の窓（DisplayAlerts）で止まる手／名前で引いて名前が変わると壊れる所／
  COM の名前付き引数の取り違え／保存していない人の変更を消しうる手／子プロセスの stdin 引き継ぎ・終わらない呼び出し
- 重さ: **壊す**（人の中身・コード・シートが消える／書き換わる）＞ **止まる**（窓や子プロセスで戻らない）＞
  **誤報**（失敗・未実施を成功と言う、理由を取り違える）＞ **読みにくい**
- 確度の書き分け:
  - 「筋は確か」＝コードの流れはその通りに読める（Excel 側の振る舞いも Excel の決まりごととして確か）
  - 「読んだだけ」＝Excel や Windows の振る舞いを推測で補っている。実機で 1 回確かめてから直すのがよい

番号は重い順。同じ重さの中は、起きやすい順。

---

## 壊す

### 1. run-macro がマクロの書き換えを「保存済み」に戻す（閉じるときの保存の確認が出ない）
- 場所: `vbam_vba.py:8240`（注入前の `owner.Saved` を控える）・`vbam_vba.py:8334-8336`（撤去後に `Saved = True`）
- 何が起きるか: アクティブブックにあるマクロを `run-macro マクロ名` で撃つと、ハーネスはそのブック（持ち主）に入る。
  撃つ前にブックが保存済みなら、撃った後に `h["wb"].Saved = True` を立て直す。マクロがそのブックのセルを書き換えていても
  「保存済み」になるので、人が Excel を閉じると保存の確認が出ず、マクロの書き換えが黙って消える。
- 入力と結果: 保存済みの 売上.xlsm を前に出し `run-macro 集計` →（集計 が同じブックの B 列を書き換える）→ 成功と出る →
  人が×で閉じる → 確認なしで閉じ、B 列の書き換えが無くなる。
- 対比: seiri（`vbam_view.py:2886`）は `if was_saved and not has_changes:` と、変化が無いときだけ戻している。
  run-macro のハーネス（`_remove_run_harness`）にはこの条件が無い。
- 確度: 筋は確か。

### 2. run-macro が人のマクロを DisplayAlerts=False のまま走らせる
- 場所: `vbam_vba.py:7898`（撃つ前に `xl.DisplayAlerts = False`）・`vbam_vba.py:8101`（終わりに無条件で True）
- 何が起きるか: 撃たれるマクロの中の確認の窓が全部「既定の答え」で通る。マクロ作者が人に聞くつもりで書いた
  `ThisWorkbook.Close`（保存の確認）・`Sheets("x").Delete`（削除の確認）・`SaveAs`（上書きの確認）が、聞かれずに進む。
  窓の見張り（キャンセル優先）は窓が出たときに効く仕掛けなので、ここでは働かない。
- 入力と結果: `ThisWorkbook.Close` だけのマクロを run-macro で撃つ → 保存の確認が出ず、閉じる（既定の答えは「保存しない」と読んでいる）。
- あわせて: `vbam_vba.py:8097` のコメント「ツール側では切っていないが」は 7898 と食い違う（→ 30）。
- 確度: 読んだだけ（DisplayAlerts=False のときの既定の答えが、閉じる＝保存しない・削除＝消す・上書き＝上書き、という部分）。

### 3. agent --undo が、控えの後に人が足したシートも消す。しかも確かめる前に消す
- 場所: `vbam_undo.py:1383-1385`（`_restore_layout` を呼ぶ）・`vbam_undo.py:667-681`（控えに無いコード名のシートを DisplayAlerts を切って Delete）・
  `vbam_undo.py:1386-1390`（`missing` の確かめはその後）・`vbam_undo.py:1405-1500`（控えを開く try はさらに後）
- 何が起きるか:
  - 消す相手は「控えを取った時点に無かったシート」全部。agent が足したシートだけでなく、agent の後に人が手で足したシートも消える。
    控え（.xlsx）にはそのシートが無いので、--undo では戻らない。`stale` の印は道具の build／macro の走行だけが付けるので、人の手作業では止まらない。
  - 並び・名前・シートの削除を先にやり、その後で `missing`（戻す先のシートが無い）や控えを開く失敗で `return False` する。
    「エラー: 戻せませんでした」と言いながら、シートは消えて名前は戻った半端な状態が残る。
- 入力と結果: agent で表を直す → 人が「メモ」シートを足して書く → `agent --undo` → 「控えの後に増えたシートを消しました: メモ」。
- 確度: 筋は確か。

### 4. agent --undo が、貼り戻す前にシートを Clear する（貼り付けが落ちると空のシートが残る）
- 場所: `vbam_undo.py:1446-1452`（`ws.UsedRange.Clear()` → `src_ur.Copy()` → `PasteSpecial`）
- 何が起きるか: クリップボードを他のアプリが掴んでいる等で Copy／PasteSpecial が例外になると、外側の except で
  「エラー: 戻せませんでした」と返る。その時点で戻す先のシートは Clear 済み＝空。控えファイルは残っているが、--undo のやり直しは
  今の空のシートを相手にする。
- 確度: 読んだだけ（PasteSpecial がクリップボードの都合で落ちる頻度は推測）。

### 5. test を本体で回したとき、名前の変わった元のシートを「テストが作ったシート」として消す
- 場所: `vbam_vba.py:8642-8655`（`initial_sheet_names` に名前が無いシートを DisplayAlerts を切って Delete）
- 何が起きるか: 本体で回るのは `--in-place` のときと、一度も保存していないブック（Path が空）のとき（`vbam_vba.py:8452`）。
  テストが既存シートの名前を変えると（Sheet1 → 集計）、「集計」は初めの名前の集合に無いので消える。中身は元の Sheet1 そのもの。
  名前で見ているので、テストの間に人が足したシートも消える。
- 入力と結果: 未保存の Book1 で `Sheet1.Name = "集計"` をするテストを `test` → 「テストが残したシートを消しました: 集計」。
- 確度: 筋は確か。

### 6. sheet copy で、隠れたシートを写すと前に出ているシートの名前を変える
- 場所: `vbam_edit.py:2553-2556`（`sh.Copy(None, sh)` の後 `newsh = wb.ActiveSheet` に新しい名前を付ける）
- 何が起きるか: 写したシートを「アクティブシート」で拾っている。隠れたシートの写しは隠れたまま作られアクティブにならないので、
  名前が付くのは元から前に出ていた人のシート。
- 入力と結果: 「設定」（非表示）が隠れていて「売上」が前 → `sheet copy 設定 設定2` → 「売上」が「設定2」に改名される。
- 確度: 読んだだけ（隠れたシートの写しが隠れたまま・アクティブにならない、の部分）。

### 7. replace-procedure（Attribute 行なしの道）で、消した後に入れるのが落ちるとプロシージャが消えたまま
- 場所: `vbam_vba.py:3009-3011`（`DeleteLines` → `InsertLines` → `Save`。try も戻しも無い）
- 何が起きるか: `InsertLines` が com_retry を使い切って例外になると、プロシージャは削除済みのまま例外が上へ抜ける。
  Remove＋Import の道（3082-3129）には「バックアップから戻す」「保存せずに閉じれば戻る」の案内があるが、こちらには無い。
- 同じ形: `code-replace`（`vbam_vba.py:6703-6709`）と `rename-procedure`（`vbam_devtools.py:436-440`）は `ReplaceLine` を 1 行ずつ回し、
  途中で落ちると一部の行だけ書き換わったモジュールが残る（保存はされない）。
- 確度: 筋は確か（落ちる頻度は低いと見る）。

### 8. copy-modules --overwrite で、行き先のモジュールを外した後に取り込みが落ちると、行き先から消えたまま
- 場所: `vbam_devtools.py:722-735`（`Remove(old)` → `_import_module_verified`。名前衝突以外の例外は受けていない）
- 何が起きるか: 取り込みが名前衝突以外で落ちると、例外がそのまま上へ抜け、行き先のモジュールは外れたまま。
  replace-procedure のような「控えから入れ直す」戻しは無い（控え `make_module_backup` は取っている）。
- 確度: 筋は確か。

### 9. 道具が VBA を書き換えるたびに、人の未保存の変更まで黙って保存する
- 場所: `vbam_vba.py:3085`（Remove＋Import の前に `wb.Save()`）・`vbam_vba.py:3011`・`3441`・`3502`・`3586`・`3658`・`6709`・
  `vbam_devtools.py:447`・`743` ほか
- 何が起きるか: マクロ 1 本の置換で、人が保存前にいじっていたセルの変更も一緒にファイルへ書かれる。
  「保存せずに閉じて開き直せば戻る」という逃げ道が、セルの変更については無くなる。報告は「→ 保存しました」の 1 行だけで、
  セルの変更も保存したとは言わない。
- 確度: 筋は確か（設計としての選択かもしれない。少なくとも未保存の変更があったことは言っていない）。

### 10. agent --undo のブックの照合が名前だけ
- 場所: `vbam_undo.py:1378`（`if str(wb.Name) != book`）
- 何が起きるか: 別のフォルダにある同じ名前のブック（売上.xlsx の写し等）を前に出して `agent --undo` すると、そのまま通り、
  別のブックの控えでシートを置き換える。3 の「増えたシートを消す」も同じブックに当たる。
- 確度: 筋は確か。

### 11. restore が、控えがどのブックのものかを確かめずに前のブックへ入れる
- 場所: `vbam_vba.py:4385-4419`
- 何が起きるか: モジュールの控えのファイル名は「ブック名_モジュール名_日時.bas」（`vbam_core.py:1528`）なのに、
  restore は中の `VB_Name` だけを見て、前に出ているブックの同じ名前のモジュールを置き換える。
- 入力と結果: B.xlsm を前にして `restore A_Module1_20261029_120000.bas` → B の Module1 が A の中身になる（replace-module の控えは取られる）。
- 確度: 筋は確か。

### 12. 試し撃ち・自分の台の片付けで、その台に合流した人のブックを保存せずに閉じる
- 場所: `vbam_core.py:315-323`（その台の全ブックを `Close(SaveChanges=False)` → Quit → 必要なら kill）・
  `vbam_rehearse.py:115-126`（`release_instance(pid)` は `only_saved=False`）
- 何が起きるか: 試し撃ち（rehearse）は非表示の自分の台を起こす。その間に人がエクスプローラーからブックを開き、それがこの台に合流すると
  （このコードベースの記録にある「見えない Excel に合流する」事故）、片付けで保存の確認なしに閉じられる。
  MCP のアイドル解放（`vba_mcp_server.py:192`）は未保存の台を残すが、合流して保存済みのブックは窓ごと閉じる。
- 確度: 読んだだけ（合流が起きる条件は推測）。

### 13. gate の後片付けが、新しく開いた「空に見える」ブックを保存せずに閉じる
- 場所: `vbam_vba.py:8987-9013`（`_blank_book_names`・`_close_new_blank_books`）
- 何が起きるか: 「空」の判定はシート 1 枚・UsedRange が 1 セルで空、だけ。VBA モジュールを入れた新しいブック、図形やグラフだけ置いた
  新しいブックも「空」に入る。gate の間に人が作った新しいブックがこれに当たると、保存の確認なしに閉じる。
- 確度: 筋は確か（gate の間に人が新しいブックを作る場面は多くないと見る）。

---

## 止まる

### 14. マクロを持てないブック（.xlsx）を、確認の窓を切らずに保存する手がある
- 場所: `vbam_vba.py:3011`（replace-procedure の InsertLines の道）・`3440-3441`（add-procedure）・`3502`（add-module）・
  `3586`（delete-procedure）・`3658`（delete-module）・`6709`（code-replace）・`vbam_devtools.py:447`（rename-procedure）・
  `vbam_devtools.py:743`（copy-modules の行き先）
- 何が起きるか: .xlsx に VBA を入れて `wb.Save()` すると、Excel は「次の機能はマクロなしのブックに保存できません」の窓を出す。
  ここでは DisplayAlerts を切っていないので、窓が閉じられるまで COM の呼び出しが戻らない（MCP では 600 秒で時間切れ）。
  Remove＋Import の道（3082）や save（`vbam_edit.py:3753`）は窓を切っている＝窓は出ないが、マクロは黙って落ちる。
  道具は「.xlsx にマクロを足す」場面を `note_if_macro_free_book` で想定しているので、起きる前提の手。
- 確度: 読んだだけ。`vbam_core.py:112` のコメントにある 2026-09-19 の件では「保存しました」と返っているので、
  COM からの保存では窓が出ない版もあり得る。実機で 1 回確かめたい。

### 15. 試し撃ちが、ブックの写しをイベントを切らずに非表示の Excel で開く
- 場所: `vbam_rehearse.py:240`（`xl.Workbooks.Open(book_path, UpdateLinks=0)`。前後に `EnableEvents = False` が無い）
- 何が起きるか: 写しが .xlsm で Workbook_Open に MsgBox があると、見えない Excel の中で窓が出て Open が戻らない。
  時間切れの見張り（`_watchdog`）は Run の間だけで、Open の間は掛かっていない。
  同じ「写しを別の台で開く」手の `vbam_heavy.py:2683` と `vbam_vba.py:9121` は Open の間だけイベントを切っている。
- 確度: 筋は確か（COM で起こした Excel はマクロが有効のまま開く、という前提）。

### 16. clear-key が PowerShell を、MCP の通信線の stdin を引き継いだまま・時間切れなしで呼ぶ
- 場所: `vbam_keys.py:151`（`subprocess.run(['powershell', …, '-Command', script], env=env, capture_output=True, text=True)`。
  `stdin=` も `timeout=` も無い）。呼び元は `cmd_clear_key`（`vbam_keys.py:304`）。
- 何が起きるか: MCP サーバーの中で `clear-key -y` を撃つと、子の PowerShell は親の stdin（JSON-RPC の線）を受け継ぐ。
  PowerShell 5.1 は stdin がパイプだと終わり際にその終わりを待つことがあり、その場合は戻らない。線の中身を子が読めば、
  次の要求が MCP サーバーに届かない。CLI（コンソール）からなら起きない。
- 確度: 読んだだけ（PowerShell が stdin を待つ・読む、の部分）。
- あわせて見たもの: `tasklist`／`taskkill` の呼び出し（`vbam_core.py:235`・`vbam_edit.py:3992`・`vbam_vba.py:1544`・`1590`・
  `vbam_ai.py:180`・`vbam_rehearse.py:196`）も stdin を引き継ぐが、どれも stdin を読まないので害は無いと見る。
  ただし `vbam_vba.py:1544`・`1590`・`vbam_ai.py:180` には `timeout=` が無い。`vbam_edit.py:3956` は `stdin=DEVNULL`、
  `vbam_ai.py:126` は `stdin=PIPE` で問題なし。

### 17. clean-table に窓の見張りが無い（値を書き、行を消す手なのに）
- 場所: `vbam_clean.py:134-247`（`cmd_clean_table` に `@dialog_safe` も `_start_dialog_watcher` も無い）
- 何が起きるか: 書き込み（`_do_normalize`）と `ws.Rows(r).Delete()`（`vbam_clean.py:217`）は Worksheet_Change を起こす。
  そのマクロが MsgBox を出すと、閉じるまで戻らない。write-range・write-cells は見張りを置いている（`vbam_edit.py:660`・`850`）。
  `@protect_safe` も無いので、保護したシートでは列ごとに「失敗」と出して進む（→ 21）。
- 同じく見張りが無く値を書く手: `cmd_table`（結合を解いて値を書く `vbam_edit.py:2708-2711`）、`cmd_build_sheet`。
- 確度: 筋は確か（窓が出るのはイベントマクロのあるブックだけ）。

### 18. open が確認の窓を切らずに Workbooks.Open する
- 場所: `vbam_edit.py:3940`（`xl.Workbooks.Open(path)`）
- 何が起きるか: リンクの更新の確認・パスワードの入力・「同じ名前のブックは開けません」の窓が出ると、そこで戻らない。
  先の「既に開いているか」は同じパスしか見ていないので（`vbam_edit.py:3894`）、別のフォルダの同じ名前のブックは素通りして Open に当たる。
- 確度: 読んだだけ（どの窓が COM の Open で出るかは推測）。

### 19. MCP の時間切れで世代交代した後、古いワーカーが Excel を触り続ける
- 場所: `vba_mcp_server.py:438-453`（時間切れ → `_restart_worker()` → 新しいワーカーがすぐ次の仕事を取る）
- 何が起きるか: 時間切れの仕事は止められていない（スレッドは止めようがない）。窓が閉じられた瞬間に古い仕事が続きを走らせ、
  同じ時に新しいワーカーが次のコマンドを撃つ。replace-procedure の Remove と Import の間に別の get が入る、のような食い違いが起き得る。
  返事の文は「Excel 側で続いている可能性があります」と言っているが、次のコマンドを待たせはしない。
- 確度: 読んだだけ。

---

## 誤報

### 20. MCP の確認（AI作業窓）が、戻せない手の一部しか見ていない
- 場所: `vba_mcp_server.py:480`（`GATE_COMMANDS = {"write-range", "clear-range", "delete-module", "save-as", "close"}`）・
  `vba_mcp_server.py:673-679`（`_submit_async` は `_diff_confirm` を通さない）
- 何が起きるか: 線引きは「退避の無い操作だけ人の承認」と書いてあるが、`sheet delete`・`write-cells`・`clean-table --delete-dups`・
  `find-replace`・`delete-procedure`・`code-replace` は通る。`replace-procedure … --bg` は差分の確認を飛ばす。
  効くのは環境変数 `AI_WIN_LOG` のある子セッションだけ。
- 確度: 筋は確か。

### 21. clean-table が、列の直しや仕上げが落ちても成功を返す
- 場所: `vbam_clean.py:182-186`（列ごとの失敗は print して continue）・`vbam_clean.py:243`（`cmd_tidy(ns)` の戻り値を見ない）・
  `vbam_clean.py:245`（最後は無条件に `return True`）
- 入力と結果: 保護したシートで `clean-table` → 全列「失敗」と出たうえで、戻り値は成功（MCP では「（コマンドは失敗しました）」が付かない）。
- 確度: 筋は確か。

### 22. process --kill が taskkill の結果を見ずに「終了しました」と言う
- 場所: `vbam_vba.py:1590-1592`
- 入力と結果: 管理者で動いている Excel など、taskkill が「アクセスが拒否されました」で失敗しても、「PID n を終了しました」と True を返す。
- 確度: 筋は確か。

### 23. find-replace（置き換え後が探す字を含むとき）が、数のセルを黙って飛ばす
- 場所: `vbam_edit.py:3683`（`if not isinstance(s, str) …: continue`）
- 何が起きるか: 「受注済→受注済み」のように置き換え後が探す字を含むと、道具がセルごとに置き換える道に入る。この道は文字列のセル
  （と数式）しか見ないので、数の入ったセル（例 `find-replace 1 12` の 1・10・21）を置き換えない。報告は「（n セル。…飛ばしました）」で、
  数のセルを飛ばしたことは言わない。Excel の Replace（もう一方の道）なら数のセルも置き換える。
- 確度: 筋は確か。

### 24. find が MatchByte／SearchFormat を渡さない（前の検索の設定を引き継ぐ）
- 場所: `vbam_edit.py:3572`（7 個の位置引数。MatchByte・SearchFormat が無い）
- 何が起きるか: Excel は Find の MatchByte を前回の値で続ける。find-replace は MatchByte=True を渡す（`vbam_edit.py:3644`）ので、
  その後の find は半角と全角を区別し、人が［検索］の窓で区別なしにした後は区別しない。同じ `find` の件数が前の操作で変わる。
  find-replace のコメント（`vbam_edit.py:3634-3640`）はこの性質を書いているが、find には当てていない。
- 確度: 筋は確か（MatchByte が引き継がれるのは Excel の決まりごと）。COM の名前付き引数の取り違えは見つからなかった（下の「見たが問題なし」）。

### 25. datamodel relation delete が、消せなかったときに「見つかりません」と言う
- 場所: `vbam_heavy.py:2372-2383`（`r.Delete()` の例外を `continue` で飲み、最後に「該当するリレーションが見つかりません」）
- 確度: 筋は確か。

### 26. format-range の列幅・行の高さの「+n／-n」が、列ごとの失敗を飲んで「当てた」と言う
- 場所: `vbam_edit.py:1112-1118`・`1140-1146`
- 確度: 筋は確か（失敗する場面は少ないと見る）。

### 27. agent --undo の「消せなかったもの」に、3 で消し済みのシートが出る
- 場所: `vbam_undo.py:1109-1117`（`_delete_created` の sheet は `wb.Sheets(name).Delete()`）と 3 の `_restore_layout`
- 何が起きるか: 作ったシートは `_restore_layout` が先にコード名で消す。その後 `_delete_created` が同じ名前で消しにいき、無いので
  「消せなかったもの: sheet 集計: …」と出る。名前で引いているので、そのシートの名前を人が別のシートに付け直していれば、そちらを消す。
- 確度: 筋は確か（棚の目印 `created` に sheet が積まれるのは shelf-run と agent の read_file 等）。

---

## 読みにくい

### 28. run-macro のコメントが中身と食い違う
- 場所: `vbam_vba.py:8097-8098`「ツール側では切っていないが」— 実際は `vbam_vba.py:7898` で切っている（→ 2）。

### 29. raise の後に届かない旧処理が 50 行残っている
- 場所: `vbam_core.py:1127-1130`（必ず raise）の後の `vbam_core.py:1131-1183`。`_open_maybe_repair`（`vbam_core.py:786-798`）は
  ここからしか呼ばれていないので、今は使われていない。

### 30. 更新登録（register-addin）の後片付けが、人の設定を「入」に揃える
- 場所: `vbam_vba.py:7843-7848`（ScreenUpdating・EnableEvents・DisplayAlerts を無条件に True）・
  `vbam_vba.py:7799-7807`（ほかの Excel で、もともと外していたアドインも `Installed = True` にする）
- 何が起きるか: 外す側（`vbam_vba.py:7748-7766`）は「入っていたものだけ」外すが、戻す側は入っていなかったものも入れる。

### 31. 接続キャッシュの「アクティブブックの切替に追従」が名前だけで比べる
- 場所: `vbam_core.py:891`（`cur.Name != wb.Name`）
- 何が起きるか: 別のフォルダ・別の Excel にある同じ名前のブックへ人が切り替えても、追従しない（前のブックのまま撃つ）。
- 確度: 筋は確か（起きる場面は少ない）。

---

## 見たが問題なしと読んだもの（重複して点検しないための控え）

- シートの Delete はどれも DisplayAlerts を切っている: `vbam_edit.py:2531-2532`・`vbam_heavy.py:560-565`・`654-659`・`2043-2048`・`2093-2100`・
  `vbam_undo.py:672-681`・`1112-1117`・`vbam_vba.py:8645-8661`。
- COM の名前付き・位置引数は、見た範囲では Excel の引数の並びと合っていた:
  `ListObjects.Add`（`vbam_edit.py:2717`・`vbam_build.py:575`・`vbam_heavy.py:2055`）・`SortFields.Add`（`vbam_edit.py:2924`・`3406`）・
  `Range.Find`（`vbam_edit.py:3572`・`3644`・`vbam_hands.py:698-699`）・`Range.Replace`（`vbam_edit.py:3727`）・
  `PivotCaches().Create`／`CreatePivotTable`／`AddDataField`（`vbam_heavy.py:600-639`）・`SlicerCaches.Add2`・`Slicers.Add`（`1571-1573`）・
  `Shapes.AddPicture`（`vbam_agent.py:3763`）・`Validation.Add`・`FormatConditions.Add`・`Hyperlinks.Add`・`Workbooks.Open(…, UpdateLinks=0, ReadOnly=True)`。
- close は名指し・保存方針の明示・確認（-y）の三点がそろっている（`vbam_edit.py:4023-4144`）。save-as は既存の出力先を --overwrite なしでは止める。
- agent の export_csv は既存のファイルを上書きしない（`vbam_hands.py:87-97`）。
- MCP の stdin はスレッドごとに空の StringIO（`vba_mcp_server.py:285`）＝ `input()` は EOFError になり止まらない。
  set-key の `getpass` は MCP では `sys.stdin` が差し替わっているので fallback（同じく EOFError）に落ちると読んだ。
- 例外を飲んで成功を返す形を構文木で拾った（cmd_* の中で、書き換えの呼び出しを含む try を pass／continue だけで受ける所: 33 か所）。
  上の 21・25・26 以外は、後始末（一時モジュールの撤去・作りかけの片付け）か、報告に別の行が出るものだった。
- cmd_* が「エラー」を print した直後に値なしの `return`（＝ MCP・CLI では成功扱い）をする所は、構文木で見た範囲では 0 か所。

## 確かめ方の提案（実機で 1 回ずつ）

- 1: 保存済みの .xlsm で、同じブックの A1 を書き換えるマクロを run-macro → Excel の `ThisWorkbook.Saved` を見る（True なら再現）。
- 2: `ThisWorkbook.Close` だけのマクロを run-macro → 保存の確認が出るか。
- 3: agent で 1 手 → 人がシートを 1 枚足す → `agent --undo --dry-run` では何も言わないこと、`agent --undo` で消えることを見る。
- 14: 新しい .xlsx に `add-module` → `add-procedure` → 窓が出るか・何秒で戻るか。
- 15: Workbook_Open に MsgBox を置いた .xlsm で seiri の試し撃ちを通す。
- 16: MCP から `clear-key -y`（消す鍵が無いと PowerShell まで行かないので、環境変数に仮の鍵を入れてから）。
