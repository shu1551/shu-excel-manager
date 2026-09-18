# 作業フロー・現地調査・動作検証・絶対に守るべきルール（全文）

> excel-vba-manager の参照。SKILL.md（型と索引）から必要なときだけ読む（2026-09-16 に分割）。

## 現地調査（シートに触るマクロの設計・破壊的操作の前に必ず）

シートを読み書きするマクロの新規作成・改修や、破壊的操作（行列削除・クリア・並べ替え等）の前に、
`snapshot` で対象の実勢（結合セル・見出しの位置・図形/ボタン・テーブル）を先に読む。

- 自動の最終行/最終列判定や「たぶんこうなっている」の推測は手作りシートで誤る
  （実害例: 最終列の誤判定で結合セルを巻き込み削除）。snapshot の事実を見てからコードを書く。
- 変更を伴う作業は**前後で snapshot を取り snapshot-diff**（checkup を前後に挟む型のシート側版）。
  「実際に何が変わったか」をセル単位の事実で確認・報告する。「直したはず」で済ませない。
- ボタンとマクロの対応を調べるときは `wiring`（行き先のないボタンも検出）、
  マクロ側から見た地図は `call-graph` / `docs`。

## 動作検証（マクロを動かして確かめる。UI操縦は最終手段）

修正したマクロの動作確認は **COM経由を第一選択**にする。computer-use での
マウス/キー操縦は使わない（打鍵がダイアログからシートに漏れ、選択範囲に Delete を
送る事故が実際に起きた。2026-07-12・未保存だったため実害なしの結果オーライ）。

- **ダイアログを開くマクロの検証**: `py vba_manager.py run-macro <マクロ名>` 一発でよい。
  run-macro のダイアログ安全解除ガードが Excel のモーダルダイアログ(#32770)を検出し、
  **タイトル（例: タイトル「セルの書式設定」）と本文を報告してから安全に閉じる**。
  この報告テキストが「どのダイアログが開いたか」の物証になる＝スクリーンショット不要。
- **セルへの効果の検証**: 前後で `snapshot` → `snapshot-diff`（セル単位の事実で差分確認）。
  範囲の見た目は `screenshot [range]`（範囲をPNG化。モーダルダイアログは写らない点に注意）。
- **MsgBox を出すマクロ**: 応答を選びたいときだけ `--auto-dialog ok|cancel|yes|no`。
  既定でもキャンセル優先で自動解除され、内容は報告される（無言ハングしない）。
- **InputBox を出すマクロ**: `--input-text 値`（複数回で出た順）。値を入れて OK で確定し、
  マクロを先へ進ませる。VBA の InputBox() も Application.InputBox も可。入れた値は報告に出る。
- どうしても実画面の目視が要るときだけ computer-use。その場合も
  ①打鍵前にスクリーンショットでフォーカス位置を確信 ②矢印連打・BackSpace連打の
  探り打ち禁止 ③シートが背後にある状態で Delete/Enter 系を打たない。

## 標準作業フロー（マクロ修正）

この手順を必ず守ること。勝手にコードを変更しない。対象は今アクティブにしているブック。

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得
   - 同名プロシージャが複数フォームにある場合は **モジュール指定** を使う：
     `py vba_manager.py get <モジュール名> <Sub名>`
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure -y` で適用（非対話実行のため -y 必須）
6. 動作を自分で検証できるものは上の「動作検証」の COM 経由手順で確認し、
   その結果を添えてユーザーに最終確認を依頼

## 標準作業フロー（フォームの .bas を修正して適用）

フォームや複数プロシージャをまとめて修正する場合（.bas ファイルを直接編集）：

1. `py vba_manager.py export-module <モジュール名>` で最新の .bas をエクスポート
2. Python スクリプトで CP932 のまま編集する（Edit/Write ツール禁止）：
   ```python
   # 編集スクリプトの雛形
   path = r"SCRIPTS\対象フォーム.bas"
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   # 行番号ベースで修正（多行文字列マッチは使わない）
   # lines[N] = 新しい行内容
   with open(path, 'w', encoding='cp932') as f:
       f.writelines(lines)
   ```
   **行番号は inspect スクリプトで事前確認する：**
   ```python
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   for i, line in enumerate(lines, 1):
       if "対象キーワード" in line:
           print(f"L{i}: {repr(line)}")
   ```
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用

## 標準作業フロー（標準モジュール全体の適用）

.bas ファイルを編集してモジュール全体を置換する場合：

1. `py vba_manager.py export-module <モジュール名>` でエクスポート
2. .bas ファイルを Python で CP932 のまま編集
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用


## 絶対に守るべきルール

### プロシージャ名・モジュール名は VBA の識別子規則で（先頭 `_` 禁止）
- VBA の識別子は**英字か日本語で始める**。`_`・数字・記号では始められない。
- `_tmp検証` のような先頭 `_` の Sub を注入する事故が過去に繰り返された。
  AddFromString / InsertLines は構文検査をしないため**注入自体は成功報告になり**、
  モジュールがコンパイルエラーで死ぬ（成り済まし成功の典型）。
- 一時検証用の Sub にも普通の名前を使う（例: `tmp検証`・`テスト検証`。`_` を頭に付けない）。
- ツール側でも機械拒否する: validate_vba_code（replace-procedure / add-procedure）・
  replace-module・add-module・check-bas・form_builder.inject_vba・
  form_layout（起動マクロ名）・form_tool copy-form（新フォーム名）が
  識別子規則違反を検出して注入前に停止する。エラーが出たら名前を直す（--force で潰さない）。
- **注入経路の台帳**: test_tools.py の `test_injection_route_ledger` が全注入プリミティブ
  （AddFromString / InsertLines / VBComponents.Import）の所在とガード状態を機械照合する。
  新しい注入経路を実装したら、①ガードを配線 ②台帳に理由つきで登録（未登録はテストが落ちる）。
  「どの経路が塞がっているか」を調べたいときも、この台帳を読めば再調査ゼロで済む。

### 勝手な変更の禁止
- 指示されていない機能を追加しない
- 指示されていないコードを変更しない
- 影響範囲を確認せずに変更しない

### エンコーディング（最重要 - 違反厳禁）
- .bas ファイルは **CP932**（Shift-JIS）で保存
- win32com 経由の文字列は Unicode
- _last_proc.vba は UTF-8

**⚠ .bas ファイルに Edit ツール・Write ツールを絶対に使うな ⚠**
Claude の Edit / Write ツールは UTF-8 で書き込むため、CP932 の .bas ファイルが破壊される。
VBA にインポートするとモジュール名・プロシージャ名・日本語文字列が全て文字化けする。

.bas ファイルを修正するときは必ず以下のいずれかを使うこと：
1. **Python で CP932 のまま読み書き**:
   ```python
   with open(path, 'r', encoding='cp932') as f:
       content = f.read()
   # 修正処理
   with open(path, 'w', encoding='cp932') as f:
       f.write(content)
   ```
2. **_last_proc.vba 経由で replace-procedure**（プロシージャ単位の修正）

### 取り込み前の単体検査 check-bas（COM不要・バイパス時の安全網）

`.bas` を VBA に取り込む前に、機械的な事故（文字コード・改行二重化・重複）を1コマンドで検査する。
**COM接続が落ちていても動く**ので、何らかの理由でツールの通常経路を通さず手書きで `.bas` を
作った場合でも、取り込み前にこれを必ず通すこと。

```bash
py vba_manager.py check-bas <file.bas>         # 検査だけ（問題があれば終了コード1）
py vba_manager.py check-bas <file.bas> --fix   # 改行二重化(\r\r\n)だけCP932のまま自動修正
```
検査項目: ①UTF-8化/BOM（文字コード事故）②改行二重化（行数が倍に膨れる事故）
③Sub/Function 名の重複（重複挿入）④VBA識別子規則違反（先頭 `_` 等＝コンパイルエラーになる名前）
⑤連続する同一コード行（重複挿入の臭い・警告）。
③⑤は判断が要るので自動修正しない（報告のみ）。④も自動修正しない（名前を直してから取り込む）。
replace-module は取り込み直前に①②を
自動で行うが、**ツールを通さず Import する場合の最後の砦**がこの check-bas。

### モジュール適用方式
- AddFromString は Attribute 行を正しく処理しないことがある
- **必ず Remove + Import 方式（replace-module）を使うこと**
- ショートカットキーは Attribute VB_ProcData.VB_Invoke_Func で定義される

### replace-module の副作用
- Remove+Import でモジュールが VBComponents の末尾に移動する
- マクロの表示順（メニュー等）が変わる場合がある
- 影響を受けたモジュールも replace-module して順番を揃える
- ショートカットキー（Attribute VB_Invoke_Func）の**セッション登録は Remove で剥がれる**が、
  Import 直後にツールが MacroOptions で自動再登録する（2026-07-12〜）。
  「修正直後にショートカットが効かない・開き直すと直る」症状はこれで解消済み。
  再登録された鍵は実行結果に「ショートカット再登録: Ctrl+Shift+X → マクロ名」と表示される。

### InsertLines の改行問題
- Python の \n では正しく複数行に分割されないことがある
- .bas ファイル直接編集方式を使うこと

## 変更前の確認チェックリスト

コードを変更する前に必ず以下を確認：

1. 変更対象のフォーム/モジュールが他の機能から参照されていないか
2. AutoFilter の起点列（B1始まりかA1始まりか）
3. 既存のコントロール名と用途（別機能で使われていないか）
4. 変更がユーザーの指示の範囲内か
5. シートに触るマクロ・破壊的操作は、snapshot で実勢（結合・見出し・図形）を確認したか（上の「現地調査」）
