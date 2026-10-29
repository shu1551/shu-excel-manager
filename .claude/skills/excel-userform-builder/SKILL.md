---
name: excel-userform-builder
description: >
  Excel VBA UserForm の作成・修正・再構築を行うスキル。form_builder.py モジュールを使って
  Python から UserForm のコントロール配置と VBA コード注入を自動化する。
  ユーザーが「フォームを作って」「UserForm を追加して」「フォームを修正して」「ボタンを追加して」
  「Excel のフォーム」「倍率回転フォーム」「カレンダーフォーム」など、Excel UserForm に関する
  作成・修正・再構築の話をしたときに必ずこのスキルを使うこと。
  フォームのコントロール配置だけでなく、VBA イベントコードの作成も含む。
---

# Excel UserForm Builder

Excel VBA の UserForm を Python (win32com) 経由で作成・修正するためのガイド。

## 前提条件

- 対象の .xlsm を Excel で開いて**アクティブ**にしておくこと（対象は常に**アクティブな開いているブック**。特定ブック名・パスに依存しない）
- ツール（`form_builder.py` / `vba_manager.py` / `form_inspect.py`）は作業中プロジェクトの `…\作業ファイル\project\python_scripts\` にある。**絶対パスを決め打ちしない** — 場所が不明なら Glob `**/form_builder.py` で特定し、そのディレクトリで実行する
- Python は `py` コマンドを使う（`python` ではない）

## まず最短ルートを選ぶ（重要）

UserForm の作業は2種類。混同すると遠回りになる。

### A. 既存フォームの「コードだけ」直す（イベント処理の修正・バグ取り）
**form_builder で作り直さない。** 差分の大きさで経路を選ぶ：

- **1〜数行の小修正**は `py vba_manager.py code-replace "旧" "新" --module <フォーム名> -y` が最短。
  **変更行だけ送る＝プロシージャ全文を再送しない**（所要時間の支配項はツールでなく全文の打ち直し）。
- プロシージャの新規追加や大改造（過半が変わる）だけ、下のプロシージャ単位置換を使う：

1. 把握： `py form_inspect.py <フォーム名> [proc ...]`
   → **1回のCOM接続**でコントロール配置＋VBAコード（全体 or 指定プロシージャ）をまとめて出す。
     export-module と get を別々に何度も叩く（毎回接続し直し＝遅い）のを避ける入口。
   → `py form_inspect.py --list` で開いているブックの全フォーム名。
2. 取得： `py vba_manager.py get <フォーム名>.<プロシージャ>` → `_last_proc.vba`(UTF-8) に保存
3. 修正版を **UTF-8 の .vba** に書く（**.bas は絶対に Edit/Write しない**）
4. 置換： `py vba_manager.py replace-procedure --code-file fix.vba --module <フォーム名> -y`
   - **ブック保存まで自動**。Attribute 行の無い普通のプロシージャは InsertLines＝**速い経路**。
   - 同名プロシージャが複数モジュールにある時だけ `--module` が必須。

### B. コントロールの配置・追加・全面再構築
レイアウトを変える／コントロールを足す／作り直すとき**だけ**、form_builder フロー（`references/build.md`）
（get_or_create → clear_controls → 再配置 → inject_vba）を使う。

### C. フォームを別のブックへ運ぶ・テキストで残す
`.frx` はバイナリで中身が読めず差分も取れない。**`form-to-vba` で「VBA だけで組み立て直す
作成マクロ」に変換する**と、テキスト1本になるので差分が読め、別ブックへは流し込むだけで移る。

```
py vba_manager.py form-to-vba <フォーム名>            # → _<フォーム名>_作成マクロ.vba（別名: フォーム書き出し）
py vba_manager.py form-to-vba --all --add shu002     # ブック内の全フォームを書き出して指定モジュールへ追加
py vba_manager.py form-to-vba <フォーム名> --verify   # 捨てブックで実際に組み立てて元と照合（元ブックには触らない）
py vba_manager.py form-to-vba <フォーム名> --name 新名 --sub 作成マクロ名
```

- 受け取る側には **VBOM（VBA プロジェクト オブジェクト モデルへのアクセス）の信頼設定**が要る。
- 自分自身のブックで作成マクロを実行しても現存フォームを壊さないよう、生成コードは同名フォームを
  いきなり消さず一度逃がしてから差し替える（削除直後の名前は Excel を閉じるまで VBE が握るため）。

### 作業前に：表示中のフォームは Designer が取れない
フォームが画面に出ていると、そのフォームの Designer に触れない＝配置の取得も変更もできない。
**人に「閉じてください」と頼まず、自分で閉じる。**

```
py vba_manager.py close-form --list           # 閉じずに表示中フォームの一覧だけ
py vba_manager.py close-form                  # 表示中の全フォームを閉じる
py vba_manager.py close-form <キャプション>     # 窓の題名で絞る（完全一致→無ければ部分一致）
```

- ×ボタンと同じ合図（`WM_SYSCOMMAND` / `SC_CLOSE`）を送る。COM を使わないのでモーダル表示中でも効く。
- `QueryClose` で閉じるのを拒むフォーム（処理の実行中など）は閉じられない旨が返る。**そのときだけ人に頼む。**
- ⚠ 自前で閉じるコードを書くとき `WM_CLOSE` を直接投げてはいけない（実測 2026-08-23）。窓だけ壊れて
  VBA の `Unload` を通らず、次に同じフォームを開くと実行時エラー 80010108（切断）→ Excel ごと落ちる。
- VBE が中断モードだと Designer はやはり取れない。`py vba_manager.py vbe-reset --check` で状態を見て、
  解除は人の指示を得てから `vbe-reset`（走っている VBA が全部止まるため）。

## 細部は別の文書（要るときだけ Read）

| 読む場面 | ファイル |
|---|---|
| フォームを設計する・配置を直す（寸法・整列・ボタン・定番レシピ 4 つ） | `references/design.md` |
| Python で配置を組む・form_builder.py の API（接続・メソッド・コントロール追加） | `references/build.md` |

## 重要な注意事項

- **VBA コードファイルのエンコーディング**: `.vba` ファイルは必ず **UTF-8** で保存する（`inject_vba` が UTF-8 で読む）
- **.bas ファイルのエンコーディング**: 標準モジュールの `.bas` は **CP932 (Shift-JIS)**
- **⚠ .bas ファイルに Edit ツール・Write ツールを絶対に使うな ⚠**: Claude の Edit/Write は UTF-8 で書き込むため CP932 が破壊される。修正は必ず Python で `open(path, 'r', encoding='cp932')` → 修正 → `open(path, 'w', encoding='cp932')` とする
- **初回作成時のリネーム失敗**: VBComponents.Add(3) 後の Name 設定が失敗することがある → UserFormN になった場合は手動リネーム後にスクリプト再実行
- **既存フォームの修正**: `get_or_create` は既存フォームを再利用し、`clear_controls` で全コントロールを削除してから再配置する
- **コントロール名**: VBA イベントプロシージャ名と一致させること（例: `CommandButton1` → `Private Sub CommandButton1_Click()`）

## 既存の実装例

`create_*.py` / `_*_code.vba` 形式の作成スクリプトが**プロジェクト内に残っていれば**参考にする
（例：カレンダーフォーム、ノード編集フォーム）。ただし**存在を前提にしない** — 無ければ
`references/build.md` の雛形（「作業フロー」§2）から書き起こす。あるか確かめるときは Glob `**/create_*.py` で探す。
