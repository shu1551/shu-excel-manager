# 作業フローと form_builder.py の API

Python で配置を組む・API の引数を確かめるときだけ読む（SKILL.md から分けた・2026-10-29）。

## 作業フロー

### 0. まず form_layout.py を検討する（推奨・宣言的レイアウト）

行構造で書けるフォームなら、座標計算を自前でやらず form_layout.py に任せる
（ラベル列幅の自動揃え・余白リズム・ボタンバー・TabIndex 整列が全部機械計算になる）。
使い方は excel-vba-manager スキルまたは form_layout.py の docstring を参照。
自由配置（カレンダー格子等）だけ form_builder の add_* + Grid を直接使う。

### 1. 要件を確認する

- フォーム名、キャプション、サイズ
- 必要なコントロール（ボタン、テキストボックス、リスト等）とレイアウト
- 各コントロールのイベント処理（クリック時の動作等）
- 対象ブック（秀.xlsm、ポスター.xlsm 等）

### 2. Python スクリプトを作成する

`create_<フォーム名>.py` として保存する。基本構造:

```python
"""
<フォーム名>を<ブック名>に作成するスクリプト
実行前に<ブック名>をExcelで開いておくこと。
"""
import os
import traceback
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_FILE = os.path.join(BASE_DIR, "_<フォーム名>_code.vba")


def build_controls(f):
    # コントロールを配置
    add_btn(f, "BtnOK", "OK", 100, 160, 60, 20)
    # ... 他のコントロール


def main():
    try:
        with FormBuilder.connect() as fb:   # 既定はアクティブブック（特定ブックに縛らない）
            frm = fb.get_or_create("フォーム名",
                                   caption="タイトル", width=300, height=200)
            f = fb.clear_controls(frm)
            build_controls(f)
            print(f"コントロール合計: {f.Controls.Count} 個")
            fb.inject_vba(frm, VBA_FILE)
            fb.save()
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    main()
```

### 3. VBA コードファイルを作成する

`_<フォーム名>_code.vba` として **UTF-8** で保存する。フォームのイベントプロシージャを記述:

```vba
Private Sub CommandButton1_Click()
    ' ボタンクリック時の処理
End Sub

Private Sub UserForm_Initialize()
    ' フォーム初期化処理
End Sub
```

### 4. スクリプトを実行する

```bash
# form_builder.py のあるディレクトリで実行（py コマンド・対象はアクティブブック）
py create_<フォーム名>.py
```

## form_builder.py API リファレンス

### 接続方法

```python
FormBuilder.connect(wb_path=r"C:\...\file.xlsm")   # フルパス指定
FormBuilder.connect(wb_keyword="ポスター")           # ブック名キーワード検索
FormBuilder.connect()                                # アクティブブック
```

### FormBuilder メソッド

| メソッド | 説明 |
|---------|------|
| `get_or_create(name, caption=, width=, height=)` | フォームを取得 or 新規作成 |
| `delete_form(name)` | フォームを削除 |
| `clear_controls(frm)` | 全コントロール削除、Designer を返す |
| `inject_vba(frm, vba_file)` | UTF-8 の .vba ファイルからコード注入 |
| `save()` | ブックを保存 |
| `list_forms()` | フォーム名一覧を返す |

### コントロール追加関数

全関数の共通引数: `(frm, name, left, top, width, height)`
デフォルトフォントサイズ: 12pt（`font_size=` で個別変更可）

| 関数 | 追加引数 | 説明 |
|------|---------|------|
| `add_btn(frm, name, caption, ...)` | `font_bold=False` | CommandButton |
| `add_lbl(frm, name, caption, ...)` | `align=1, fore=, back=, font_bold=` | Label (align: 1左,2中央,3右) |
| `add_txt(frm, name, ...)` | `value="", multiline=False, scrollbars=0` | TextBox |
| `add_lst(frm, name, ...)` | | ListBox |
| `add_combo(frm, name, ...)` | | ComboBox |
| `add_chk(frm, name, caption, ...)` | | CheckBox |
| `add_opt(frm, name, caption, ...)` | | OptionButton |
| `add_frame(frm, name, caption, ...)` | ※font_size なし | Frame |
| `add_img(frm, name, ...)` | ※font_size なし | Image |
| `add_spin(frm, name, ...)` | ※font_size なし | SpinButton |
| `add_scroll(frm, name, ...)` | `orientation=0` (0水平,1垂直) | ScrollBar |
