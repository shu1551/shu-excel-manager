# UserForm の作成・修正（全文）

> excel-vba-manager の参照。SKILL.md（型と索引）から必要なときだけ読む（2026-09-16 に分割）。

## UserForm の作成・修正

最短ルートの使い分け（UserForm 専門の作業は excel-userform-builder スキルも参照）:

1. **現状把握（目）**: `py form_inspect.py <フォーム名>` — 1接続でフォーム情報＋コントロール配置＋コード。
   `--font`（Font.Size列）`--json`（機械可読）`--png`（実表示してPNG撮影・見た目の確認）`--list`（フォーム一覧）
   `--png --names`（枠と名前を画像に描き込み＝名指しで直せる）`--png-all`（MultiPage の全タブを1枚ずつ撮影）
   `--lint`（重なり/はみ出し/不揃い/タブ順に加え、**孤児イベントハンドラ・Click未実装・右端の揃い忘れ**も機械検査）
2. **コードだけ直す**: form_builder で作り直さず `vba_manager.py get → replace-procedure`（速い・自動保存）
3. **幾何だけ直す（手）**: `py form_tool.py` — 1接続の機械操作CLI。
   ```
   py form_tool.py scale <フォーム> 1.15            # フォーム+全コントロール+フォント一括倍率
   py form_tool.py set <フォーム> btnOK,btnNG --top 160 --font-size 14
   py form_tool.py move <フォーム> btnA,btnB --dx 8 --dy -4      # 相対移動
   py form_tool.py align <フォーム> btnA,btnB --left 10          # 辺を揃える
   py form_tool.py size-match <フォーム> btnA,btnB --ref btnB    # 同サイズ化
   py form_tool.py distribute <フォーム> a,b,c --vertical --gap 8  # 等間隔配置
   py form_tool.py tab-order <フォーム>                          # TabIndex を視線順に自動整列
   py form_tool.py rename-control <フォーム> btnA btnB   # イベント宣言行も機械追随（件数報告）
   py form_tool.py delete-control <フォーム> btnOld      # イベントコードは残す（警告表示）
   py form_tool.py copy-form <フォーム> <新名>           # レイアウト+コード丸ごと複製
   ```
   既定では保存しない（--save で保存）。lint が指摘した内容はここで直せる。
   rename-control / delete-control は実行前に .frm/.frx を backups へ自動退避する。
   tab-order の自動整列はコンテナ（フォーム/Frame/Page）単位で行われる。
4. **新規・作り直し（推奨: 宣言的レイアウト）**: form_layout.py — 行構造を書くだけで
   ラベル列整列・8ptリズム余白・ボタンバー右寄せ・TabIndex・Default/Cancel まで機械計算。
   ```python
   from form_layout import (build_form, preview_layout, row, lbl, txt, combo, lst, chk,
                            opt_group, button_bar, ok, cancel, spacer, heading, frame)
   rows = [
       heading("顧客情報"),
       row(lbl("顧客名"), txt("txtName")),
       row(lbl("区分"), combo("cmbKind", items=["法人", "個人"])),
       frame("配送オプション",
             row(lbl("優先度"), opt_group(("optHigh", "急ぎ"), ("optNorm", "通常"))),
             row(chk("chkGift", "ギフト包装"))),
       spacer(),
       button_bar(ok("btnSave", "登録", accel="S"), cancel("btnClose", "閉じる")),
   ]
   preview_layout(rows)         # Excel 不要のワイヤーフレームPNG（設計の高速な試行錯誤）
   build_form("F_Order", "受注入力", rows, vba_stub=True, png=True, launcher="Module1")
   #  vba_stub=True: Initialize（items の AddItem・先頭入力へ SetFocus）と
   #  各ボタンの Click 雛形を機械生成して注入。png=True: 構築後に実表示PNG。
   #  launcher="モジュール名": Sub <フォーム名>を開く() を標準モジュールに自動追加
   #  （メニュー方式のブックならそのままメニューに載る）。
   #  combo/lst は rowsource="シート名!A1:A10" でシート範囲に直結も可。
   #  既存フォームは backups へ .frm/.frx 自動退避してから作り直す。
   ```
   **CLI からも使える**: `py form_layout.py preview 宣言.py`（Excel不要の配置図）／
   `py form_layout.py build 宣言.py`（実構築）。同じ宣言ファイルが両方で動く。
   幅未指定の入力は右端まで自動ストレッチ。frame は入れ子1段まで。スタイル定数は form_layout.STYLE。
   **追加部品**: `refedit("refX")`＝範囲選択欄（TextBox+選択ボタン+InputBox(Type:=8)。本物の RefEdit は
   COM挿入が信頼設定でブロックされるため複合部品で実装）、`spin_txt("txtN", min_=1, max_=99)`＝▲▼付き数値、
   `img("imgX", w, h, picture=パス)`＝画像（実行時 LoadPicture）、`txt(..., required=True)`＝ラベルに＊＋
   実行ボタンに空チェック雛形。定番の組み合わせは excel-userform-builder スキルの「定番レシピ」参照。
   **タブ付きフォーム**は `multipage("mpMain", page("基本", row(...)), page("詳細", row(...)))`。
   ページ内に frame も置ける。--to-layout の逆変換・preview_layout（1ページ目＋タブ帯）にも対応済み。
   ※ --png の実表示撮影は「表示中のタブ」しか写らない（他タブは build 後にExcelで切り替えて確認）。
   構築後は `form_inspect --lint` と `--png --names` で検証（デザイン規範は excel-userform-builder スキル参照）。
5. **既存フォームの大改修**: `py form_inspect.py <フォーム> --to-layout` で
   既存フォームを form_layout の宣言コード（たたき台）に逆変換 → 宣言を編集 → build_form。
   往復（逆変換→再構築）で見た目が保たれることは検証済み。items 等の実行時情報だけ手動補完。
6. **自由配置（カレンダー格子等）**: form_builder.py（下記）

```python
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo
from form_builder import Grid, vstack, hstack   # 機械的な座標計算ヘルパー

with FormBuilder.connect() as fb:  # アクティブブック接続
    frm = fb.get_or_create("FormName", caption="タイトル", width=300, height=200)
    f = fb.clear_controls(frm)     # ※レイアウト全消し。コードだけ直すなら使わない（上の2へ）
    add_btn(f, "BtnOK", "OK", 80, 160, 60, 20)
    add_lbl(f, "Label1", "テキスト", 10, 10, 100, 18)
    add_txt(f, "TextBox1", 10, 30, 200, 22)
    add_lst(f, "ListBox1", 10, 60, 200, 90)
    fb.inject_vba(frm)             # VBAコード注入（省略時 _last_form_code.vba / 明示指定も可）
    fb.save()
```

デフォルトフォントは全コントロール 12pt。グリッド配置は `g = Grid(10, 40, 24, 20); g.pos(row, col)`、
縦積み/横並びは `vstack(直前コントロール)` / `hstack(直前コントロール)`。

> フォームの .frm を replace-module に渡すときは **同名 .frx を同じフォルダに置く**こと
> （無いと停止する。.frx にレイアウトが入っているため）。restore も .frm/.frx ペアで復元できる。
