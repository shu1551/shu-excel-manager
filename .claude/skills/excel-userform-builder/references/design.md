# デザイン原則と定番レシピ

フォームを新しく設計する・配置を直すときだけ読む（SKILL.md から分けた・2026-10-29）。

## デザイン原則（AI がフォームを設計するときの規範）

MSForms の素の見た目は古いが、**整列・余白・統一・順序**の4つが揃うだけで見違える。
座標を暗算で一発置きせず、以下を規範として設計し、必ず PNG で検証する。

### 寸法のリズム（8の倍数で刻む）
- フォーム外周の余白: **12pt**。要素間の縦間隔: **8pt**（強い関連は 6、セクション区切りは 16）
- 高さの標準: TextBox/ComboBox **22**、ボタン **24〜28**、ラベル **18**、リスト行数×12+6
- **同種のコントロールは同じ幅・高さに揃える**（特にボタン。バラバラの高さは素人臭さの最大要因）

### 整列（縦ラインを最少に）
- 基本は「ラベル列＋入力列」の**2本の縦ライン**。左端が3本以上あると散らかって見える
- ラベルは入力欄と**上下中央**を合わせる（top = 入力top + (入力H − ラベルH)/2）
- 右端も揃える（一番幅の広い入力に合わせる）

### ボタンの作法（Windows 規約）
- OK/キャンセルは**右下・同サイズ**（最低 72×24）。並びは「実行系 → キャンセルが右端」
- Default=True（Enter）を実行ボタンに、Cancel=True（Esc）をキャンセルに設定する
- 破壊的操作（削除等）のボタンは OK 群から**離して**置く（左下など）

### フォント・色（増やさない）
- 全コントロール **12pt 統一**（ツールの既定）。強調は **Bold のみ**。フォント種は増やさない
- 色は最小限: 背景は既定のまま、強調色は**1色まで**。多色は判断コストを増やすだけ

### 順序と操作
- **TabIndex は視線順**（左→右、上→下）——ただし**新規フォームのみ**。既存フォームの
  TabIndex・初期フォーカスは挙動の一部なので勝手に整列し直さない（一覧系は ListBox が先頭）
- 主要ボタンに Accelerator（&文字）を設定
- UserForm_Initialize で最初の入力欄に SetFocus

### 命名
- `btn/txt/lbl/lst/cmb/chk/opt/fra` ＋意味名（btnSave、txtName）。イベントプロシージャ名と一致させる
- **コントロール名・Sub 名とも先頭 `_` 禁止**（VBA の識別子は英字か日本語で始める。
  `_tmp検証` 級の先頭 `_` は注入は通るがコンパイルで死ぬ——form_builder.inject_vba が注入前に機械拒否する）

### 検証は規模に合わせる（レイアウトを触ったら PNG 1枚・触らなければ撮らない）
- **新規フォーム・大改造**: `py form_inspect.py <フォーム> --png` で実表示を見て、`--lint`
  （重なり・はみ出し・不揃い）まで通す。ズレは form_tool（set/align/scale）で直して再撮影。
- **既存フォームの小さな配置変更（ボタンの追加・移動・幅調整など）**: PNG 1枚の目視だけで良い
  （lint 不要）。
- **コードだけの修正**: 撮らない（上のルートAの範囲）。
「一発で置けたはず」の無検証納品は禁物だが、**検証を膨らませて遅くするのも同罪**。
所要時間の支配項は検証の往復数——検証は最後の変更と同じ往復に相乗りさせ、独立の往復にしない。

## 定番レシピ（form_layout の宣言をそのまま流用してよい）

### 1. 確認ダイアログ
```python
build_form("F_Confirm", "確認", rows=[
    row(lbl("この操作は元に戻せません。実行しますか？")),
    spacer(),
    button_bar(ok("btnYes", "実行する"), cancel("btnNo", "やめる")),
], vba_stub=True)
```

### 2. 検索フォーム
```python
build_form("F_Search", "検索", rows=[
    row(lbl("キーワード"), txt("txtKey", required=True), btn("btnFind", "検索", accel="F")),
    row(lbl("結果"), lst("lstResult", rows_visible=10)),
    spacer(),
    button_bar(btn("btnJump", "移動"), cancel("btnClose", "閉じる")),
], vba_stub=True)
```

### 3. 範囲を選んで実行（シート処理ツールの型）
```python
build_form("F_Range", "範囲処理", rows=[
    row(lbl("対象範囲"), refedit("refTarget")),   # TextBox+選択ボタン+InputBox(Type:=8) の複合部品
    row(lbl("件数"), spin_txt("txtN", value="1", min_=1, max_=999), lbl("件")),
    spacer(),
    button_bar(ok("btnRun", "実行"), cancel("btnClose", "閉じる")),
], vba_stub=True)
```

### 4. 設定画面（タブ付き）
```python
build_form("F_Setting", "設定", rows=[
    multipage("mpMain",
        page("基本", row(lbl("名前"), txt("txtName", required=True)),
                     row(lbl("区分"), combo("cmbKind", items=["A", "B"]))),
        page("詳細", row(lbl("メモ"), txt("txtMemo", height=54, multiline=True)),
                     row(chk("chkAdv", "詳細モードを有効にする")))),
    spacer(),
    button_bar(ok("btnSave", "保存", accel="S"), cancel("btnCancel", "キャンセル")),
], vba_stub=True, png=True)
```

いずれも `preview_layout(rows)` で Excel 無しの配置確認 → `build_form(..., png=True)` →
`form_inspect --lint` → `form_tool` 微調整、のループで仕上げる。
required=True はラベルに ＊ が付き、実行ボタンに空チェック雛形が入る。
