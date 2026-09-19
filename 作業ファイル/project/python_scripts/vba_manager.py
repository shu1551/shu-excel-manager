"""
VBAマネージャー (アクティブブック対応版)

【特徴】
- target_file を省略するとアクティブなExcelブックを自動使用
- get で取得したコードは _last_proc.vba に保存 → Claudeが読み取り・修正
- replace-procedure は _last_proc.vba を自動使用 (--code-file 省略時)
- replace-module は Remove+Import で Attribute を正しく処理

【コマンド一覧】
  list            [excel_file]                     マクロ一覧
  list-modules    [excel_file]                     モジュール一覧
  list-forms      [excel_file]                     フォーム一覧（コントロール数・コード行数付き）
  get             [excel_file] <macro_name>        プロシージャのコード取得
  replace-procedure [excel_file] [--code-file f]  プロシージャを置換
  replace-module  [excel_file] <module> <bas_file> モジュール全体を置換
  export-module   [excel_file] <module>            モジュールを .bas にエクスポート
  form-to-vba     [excel_file] <form> [--all]      UserForm を作成マクロ(.vba)に変換
                  [--add モジュール] [--verify]      （別名: フォーム書き出し）
  diag                                             動作確認

  reorder-macro   <macro> <up|down>                マクロの表示順を入れ替え
  list-shortcuts  [excel_file]                      ショートカットキー一覧
  rename-procedure <旧名> <新名> [--module 名] -y   マクロの改名（呼び元・OnAction・ショートカットまで一括）
  diff-module     <モジュール> [比較先.bas|--backup|--book 他]  モジュールの差分
  references      list | add <GUID|パス> | remove <名前> -y   参照設定
  copy-modules    <モジュール…> --to <ブック> [--overwrite]   別の開いているブックへ複製
  set-shortcut    <マクロ名> <キー> | --clear         ショートカットキーの付け替え・解除
  format-module   <モジュール> [--apply]              モジュールの整形（Dim 先頭・区切り・見出しコメント）
  backup-prune    [--days N] [--keep K] [--force]     backups の間引き（COM不要）
  status                                           いま走っているコマンドと進み具合（COM不要）
  repair          <マクロ名> [--module 名]            マクロ修理の材料を 1 手で（本文・入口・呼び元呼び先・コンパイルの行名指し・check・控えとの差分）
  versions        <ブック名> [--dir 追加]             同名ブックの写しを日時順に（Excel を開かない・COM不要）
  history         <マクロ名> [--book 名] [--deep]      マクロ本文の歴史（控え・_exports・写し。同じ本文は畳む・COM不要）
  list-file       <path.xlsm>                        閉じたブックのマクロ一覧（Excel を開かない）
  grep-files      "文字" <フォルダ|ファイル…> [-i]     閉じたブック横断の検索（Excel を開かない）
  export-file     <path.xlsm> [--dir 先]             閉じたブックのモジュールを .bas に書き出す

【目コマンド（シート状態の読み取り）】
  materials      [excel_file] [シート名]            シートを触る前の材料（使用範囲・全体の値・結合・数式・図形・列幅を1回で）
  read-range     [excel_file] [range] [--formula]  セル値（--formulaで数式）をテキスト格子で読む
  read-selection [excel_file] [--formula]           今選択している範囲を読む
  sheet-info     [excel_file]                       シート構成・使用範囲の一覧
  screenshot     [excel_file] [range] [--out f]    範囲を画像(PNG)で書き出す

【手コマンド（シートの編集・整形・構造操作／開いたままのブックに直接書込）】
  write-range    [excel_file] <range> [値]          値・数式を書込（グリッドは --tsv / _last_values.tsv）
  write-cells    <番地> <値> <番地> <値>… [--show]  飛び飛びのセルを1回で書込
  tidy           <範囲…> [--show]                   表の仕上げ（見出し・罫線・番号列は左寄せ・数値列は #,##0・列幅）
  audit          <範囲…>                           仕上げ検査（見出し・罫線・列の型・列幅・空の見出し。指摘0で合格）
  clear-range    [excel_file] <range>               範囲をクリア（--contents/--formats/--all）
  format-range   [excel_file] <range> [書式opt...]  フォント・色・罫線・書式・列幅等
  sheet          <add|delete|rename|copy|activate|show|hide|very-hide|visibility|tab-color>
  table          <create|list|delete|column|filter|filter-values|filter-clear|filters|sort|sort-multi|ref>
  name           [excel_file] <add|list|delete>     名前付き範囲

  -- 編集の足回り --
  row            <insert|delete> <行番号> [本数]    行の挿入・削除
  col            <insert|delete> <列文字> [本数]    列の挿入・削除
  copy-range     <src> <dst> [--values]            範囲コピー
  fill           <range> [--right]                  オートフィル（既定は下）
  sort           <range> [--key 列][--desc][--header] 並べ替え
  autofilter     [range] [--off]                    オートフィルタ
  -- 検索・置換 --
  find           <文字> [--book][--whole][--formula] セル検索（番地を返す）
  find-replace   <検索> <置換> [range] [--whole]     一括置換
  -- ブックの開閉 --
  open           <path>                             ブックを開く（見えているExcelに合流／未起動なら通常起動）
  close          <ブック名> --save|--no-save         ブックを1冊閉じる（名指し・保存方針・確認の三点セット）
  close-form     [キャプション] [--list]             表示中の UserForm を閉じる（×相当。モーダル中でも効く。引数なしで全部）
  vbe-reset      [--check]                          VBE「実行>リセット」＝中断モードの解除（走っているVBAは全部止まる。人の指示でだけ）
  -- 予行演習・関所 --
  rehearse       [excel_file] <マクロ名> [引数...]   コピーに試し撃ちして差分を報告（本体は無傷。--timeout 秒）
  gate           [excel_file] [絞り込み] [--timeout 秒]  関所＝停止条件。コピーで check＋全体コンパイル＋test して PASS/FAIL（終了コード 0/1。別名: 関所）
  compile        [excel_file]                          全体コンパイル。VBEの「VBAProjectのコンパイル」を押す＝呼ばれないマクロの中も見る（別名: 全体コンパイル）
  -- 保存・印刷 --
  save           [excel_file]                       上書き保存
  save-as        <path>                             別名保存
  print-setup    [--area R][--title-rows 1:3]...    印刷設定
  shape          --list / <名前>… --delete / <名前> --left N  図形の一覧・削除・位置と大きさ
  -- 仕上げ --
  cond-format    <range> --gt 100 --bg '#...'        条件付き書式
  hyperlink      <cell> <url> [--text t]             ハイパーリンク
  validation     <range> --list 'A,B,C'              入力規則(ドロップダウン)
  freeze         <cell> | off                        ウィンドウ枠固定
  comment        <cell> <text>                       セルコメント
  -- 重量級 --
  chart          <create|list|delete>                グラフ（column/bar/line/pie/scatter/area）
  chart-config   <set-title|set-type|legend|style|axis-scale|data-labels|add-series|trendline...>  グラフ詳細設定
  pivot          <create|list|delete>                ピボットテーブル（--rows/--cols/--filter/--values/--func。create model でデータモデルから）
  pivot-field    <list|add-row|add-col|add-value|remove|set-func|sort|group-date|group-numeric|show-as|top-n|filter-clear|position>  フィールド管理
  pivot-calc     <get-data|calc-field|layout|subtotals|grand-totals|refresh|style|repeat-labels|empty-as|set-source>  計算・レイアウト
  slicer         <add|list|delete>                   スライサー（ピボット/テーブルに紐づけ）
  calc-mode      [manual|auto|recalc]                計算モード確認・切替・再計算
  powerquery     <list|refresh|add|edit|delete|load>  PowerQueryの一覧・更新・作成・書換・削除・読込配線(--to sheet|model)
  connection     <list|refresh|delete> [name]        ブック接続の一覧・更新・削除
  datamodel      <list|relation|measure>             データモデル一覧／リレーション・メジャー(DAX)の作成削除
  -- AI に回させる（道具が材料を集め・API の AI に 1 回ずつ聞き・書き・読み戻し・仕上げ・検査まで自分で回す） --
  agent          "依頼文" [--sheet 名][--mode sheet|build|macro][--macro 名][--new-book][--dry-run]
                 [--recipes][--recipe 名前 [補足]][--continue][--undo [番号]][--backups][--changes][--rehearse][--fire ...]
                 [--runs][--keep-case 名前][--cases][--drop-case 名前][--ai gemini|claude]
  build-sheet    [--ask "仕様"]                      白紙から表を組む（agent --mode build の下回り）
  set-key        <gemini|claude> [キー]              API キーを金庫（DPAPI）に預ける／確認（--check）
  clear-key      <gemini|claude>                     預けたキーを消す
  -- 書き出し・掃除 --
  export-pdf     <出力.pdf> [--sheet|--range]        PDF 書き出し
  export-all     [出力先]                            全モジュールを .bas に書き出し
  process        [--list][--kill PID]                Excel のプロセス一覧・掃除（ブック 0 冊の残骸を消す）
"""

import sys
import os
import re
import shutil
import zlib
import argparse
import functools
import threading
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic


import vbam_core as _vc  # noqa: E402  (呼び出し台帳が対象ブック名を戻すため・2026-09-16)
from vbam_core import *  # noqa: F401,F403
from vbam_vba import *  # noqa: F401,F403
from vbam_view import *  # noqa: F401,F403
from vbam_edit import *  # noqa: F401,F403
from vbam_heavy import *  # noqa: F401,F403
from vbam_form2vba import *  # noqa: F401,F403
from vbam_build import *  # noqa: F401,F403
from vbam_devtools import *  # noqa: F401,F403  (2026-09-16 職場向け VBA ツール作りの手)
from vbam_lineage import *  # noqa: F401,F403  (2026-09-17 系譜と閉じたブック: versions・history・list-file・grep-files・export-file)
from vbam_audit import *  # noqa: F401,F403  (2026-09-17 数式・データ総合診断エンジン)
# AI の機能（agent・set-key・clear-key）は vbam_agent 一式があるときだけ読む（2026-09-12・本体だけでも動く形に）。
# 黙って飛ばすのは「ファイルが無い」ときだけ。中身に誤りがあれば今まで通り例外で止まる
import importlib.util as _ilu
if _ilu.find_spec("vbam_agent") is not None:
    from vbam_agent import *  # noqa: F401,F403
else:
    def _no_ai_feature(args=None):
        print("エラー: AI の機能（vbam_agent 一式）がこのフォルダにありません。本体のコマンドはそのまま使えます。")
    cmd_agent = cmd_set_key = cmd_clear_key = _no_ai_feature
from vbam_clean import cmd_clean_table, clean_plan  # noqa: F401  (2026-09-11 分割)

# ================================================================
# エントリポイント
# ================================================================

def build_parser():
    """argparse の構築（main と batch で共用）"""
    parser = argparse.ArgumentParser(
        description="VBAマネージャー (アクティブブック対応版)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python vba_manager.py list                                      # アクティブブックのマクロ一覧
  python vba_manager.py list 秀.xlsm                              # 指定ファイルのマクロ一覧
  python vba_manager.py list-modules                              # モジュール一覧
  python vba_manager.py get 空白行の削除                          # プロシージャ取得 → _last_proc.vba に保存
  python vba_manager.py get shu001 空白行の削除                   # モジュール指定してプロシージャ取得
  python vba_manager.py get アクティブマクロフォーム.CommandButton2_Click  # ドット区切りでも可
  python vba_manager.py replace-procedure                         # _last_proc.vba の内容で置換
  python vba_manager.py replace-procedure --code-file my.vba
  python vba_manager.py replace-module shu001 shu001_new.bas
  python vba_manager.py export-module shu001                      # shu001.bas にエクスポート
""")

    sub = parser.add_subparsers(dest="command")

    # diag
    sub.add_parser("diag")

    # setup-check（導入セルフ診断・初心者が最初に打つ1コマンド）
    p = sub.add_parser("setup-check", help="導入セルフ診断（Python/pywin32/Excel/VBOM信頼設定を○×表示）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list-open
    p = sub.add_parser("list-open")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list [excel_file]
    p = sub.add_parser("list")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--standard", action="store_true", help="標準モジュールのみを抽出")
    p.add_argument("--detail", action="store_true", help="所属モジュール・行数・先頭コメント付きで表示")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # list-modules [excel_file]
    p = sub.add_parser("list-modules")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # list-forms [excel_file]
    p = sub.add_parser("list-forms")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # get [excel_file] <macro_name> [...]
    p = sub.add_parser("get")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="保存先ファイル（省略時は _last_proc.vba。参照用コピーを残したいときに）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # replace-procedure [excel_file] [code_file] [--code-file file] [--module name]
    p = sub.add_parser("replace-procedure")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("--module", dest="module_opt", default=None,
                   help="適用先モジュール名を指定（同名プロシージャが複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップして自動で置換を実行します")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告を無視して強制適用します")

    # patch-procedure [excel_file] <macro_name> --target "..." --replacement "..."
    p = sub.add_parser("patch-procedure", help="プロシージャ内の指定コードをピンポイント置換（差分置換）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--target", dest="target", default=None, help="置換前の文字列")
    p.add_argument("--replacement", dest="replacement", default="", help="置換後の文字列")
    p.add_argument("--target-file", dest="target_file_opt", default=None, help="置換前の文字列を記述したファイル")
    p.add_argument("--replacement-file", dest="replacement_file_opt", default=None, help="置換後の文字列を記述したファイル")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュール名（同名が複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="構文エラー警告・バックアップ失敗を無視して強行")

    # add-procedure [excel_file] <module_name> [--code-file f] [-y]
    p = sub.add_parser("add-procedure", help="新規プロシージャをモジュール末尾に追加（コードは _last_proc.vba から）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告・バックアップ失敗を無視して強行")

    # add-module [excel_file] <module_name> [--type std|class|form]
    p = sub.add_parser("add-module", help="新規モジュールを追加（標準/クラス/フォーム）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", dest="type_opt", default="std",
                   help="モジュール種別 std|class|form（既定 std）")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-procedure [excel_file] <macro_name> [--module name] [-y]
    p = sub.add_parser("delete-procedure", help="プロシージャを削除（削除コードを表示して確認）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None,
                   help="対象モジュール名（同名が複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # docs [excel_file] [--out f.md]
    p = sub.add_parser("docs", help="ブックの構成ドキュメント（取説）を自動生成")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_docs.md）")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シートの先頭N行をMarkdown表で含める")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # checkup(健康診断) [excel_file] [--out f.md]
    p = sub.add_parser("checkup", aliases=["健康診断"],
                       help="ブックの健康診断レポート（総合判定+壊れた参照+シート検査+前回との比較）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_checkup.md）")
    p.add_argument("--form", action="store_true",
                   help="フォーム検査も実施（既定は速度優先でスキップ。Designer走査が重いため）")
    p.add_argument("--all", dest="all_books", action="store_true",
                   help="開いている全ブックを一括診断（PERSONAL.XLSB含む）")
    p.add_argument("--history", action="store_true",
                   help="診断せず過去の診断履歴（経過観察）を表で表示")
    p.add_argument("--detail", action="store_true",
                   help="--history で各回の間に起きた所見/マクロの増減も表示")
    p.add_argument("--note", default=None,
                   help="今回の診断にカルテのメモを添付（例: --note \"ボタン18を一語修正\"）")
    p.add_argument("--strict", action="store_true",
                   help="所見が1件でもあれば終了コード1（自動化のゲート用。既定は診断完了=0）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--ack-all", dest="ack_all", action="store_true",
                   help="今回のコード/フォーム所見を全て確認済み（意図的）として登録し、"
                        "以降の所見サマリ・総合判定から除外する"
                        "（存在しないマクロ呼び出し等の致命的所見は対象外）")
    p.add_argument("--show-ack", dest="show_ack", action="store_true",
                   help="診断はせず、確認済み（意図的）所見の一覧を表示")
    p.add_argument("--unack", default=None, metavar="文字列",
                   help="部分一致する確認済み所見を確認済みから外す（見直したくなった時）")

    # call-graph [excel_file] [--macro 名]
    p = sub.add_parser("call-graph", help="マクロの呼び出し関係を解析（未解決Call＝一語バグ検出つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--macro", dest="macro_opt", default=None,
                   help="このマクロを起点に呼び出しツリーを展開")
    p.add_argument("--mermaid", nargs="?", const="_DEFAULT_", default=None,
                   help="Mermaid図をMarkdownに出力（省略時 _last_callgraph.md。GitHub/Qiitaで描画可）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # impact(影響範囲) [excel_file] <マクロ名>
    p = sub.add_parser("impact", aliases=["影響範囲"],
                       help="マクロ修正前の影響範囲予告（呼び元/呼び先を間接まで一覧）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # grep [excel_file] <pattern>
    p = sub.add_parser("grep", help="全モジュール横断のVBAコード検索")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として検索")
    p.add_argument("-i", "--ignore-case", dest="ignore_case", action="store_true",
                   help="大文字小文字を区別しない")
    p.add_argument("--module", dest="module_opt", default=None, help="検索対象モジュールを限定")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定200）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # code-replace [excel_file] <検索> <置換>
    p = sub.add_parser("code-replace", help="全マクロ横断の一括置換（diffプレビュー・バックアップ・確認つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として置換")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # list-backups [キーワード] / restore <バックアップファイル>
    # stats [--days N] [--top N] [--slow N] [--json]（2026-09-16・呼び出し台帳の集計）
    p = sub.add_parser("stats", help="呼び出し台帳（_calls.jsonl）の集計＝どのコマンドを何回・平均何秒・失敗・"
                                     "遅かった呼び出し・呼び出しの間（道具の外の時間）。COM不要")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--days", dest="days", type=int, default=None, help="直近 N 日（既定 7）")
    p.add_argument("--top", dest="top", type=int, default=None, help="コマンド別の表示件数（既定 20）")
    p.add_argument("--slow", dest="slow", type=int, default=None, help="遅かった呼び出しの表示件数（既定 8）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    p = sub.add_parser("list-backups", help="backups のバックアップ一覧（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定30）")
    p = sub.add_parser("restore", help="モジュールバックアップ(.bas/.frm)を開いているブックへ書き戻す")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # list-shortcuts [excel_file]
    p = sub.add_parser("list-shortcuts")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # ---- 職場向け VBA ツール作りの手（2026-09-16・vbam_devtools）----
    # rename-procedure [excel_file] <旧名> <新名> [--module 名] [-y] [--dry-run]
    p = sub.add_parser("rename-procedure", help="マクロの改名（宣言・呼び元・Application.Run の文字・図形の OnAction・"
                                                "ショートカットまで一括。変更行だけ ReplaceLine）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None, help="宣言のあるモジュール（同名が複数のとき）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--dry-run", action="store_true", dest="dry_run", help="プレビューだけ（書き換えない）")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # diff-module [excel_file] <モジュール名> [比較先.bas] [--backup [名前]] [--book 他ブック] [--max N] [--json]
    p = sub.add_parser("diff-module", help="モジュールの差分（.bas ファイル／直前の控え／別の開いているブックの同名モジュール）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--backup", dest="backup_opt", nargs="?", const=True, default=None,
                   help="backups の控えと比べる（名前省略で最新）")
    p.add_argument("--book", dest="book_opt", default=None, help="別の開いているブックの同名モジュールと比べる")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="差分の表示行数の上限（既定200）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # references list | add <GUID|パス> [--major N --minor N] | remove <名前> -y
    p = sub.add_parser("references", help="参照設定の一覧・追加・削除（参照不可＝眠ったブックが動かない筆頭原因を直す）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--major", dest="major", type=int, default=None, help="add の版（GUID のとき・省略 0＝最新）")
    p.add_argument("--minor", dest="minor", type=int, default=None)
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # copy-modules <モジュール名…> --to <ブック名> [--overwrite] [-y]
    p = sub.add_parser("copy-modules", help="モジュールを別の開いているブックへ複製（同名は --overwrite・ショートカット再登録つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--to", dest="to_opt", default=None, help="行き先のブック名（開いていること）")
    p.add_argument("--overwrite", action="store_true", help="行き先の同名モジュールを控えを取ってから置き換える")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # set-shortcut <マクロ名> <キー> [--module 名] [-y] / --clear
    p = sub.add_parser("set-shortcut", help="ショートカットキーの付け替え・解除（MacroOptions＝保存で Attribute に書かれる）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None, help="モジュール（同名が複数のとき）")
    p.add_argument("--clear", action="store_true", help="ショートカットを外す")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="アドインと同名でも掛ける／バックアップ失敗時も強行")

    # format-module <モジュール名> [--apply] [-y] [--out f.bas]
    p = sub.add_parser("format-module", help="モジュールの整形（Dim を先頭へ・区切りの空行・ブロックの見出しコメント。--apply で反映）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--apply", action="store_true", help="replace-module 経路で反映する（既定は整形後の .bas と差分を出すだけ）")
    p.add_argument("--out", dest="out_opt", default=None, help="整形後 .bas の書き出し先（省略時 <モジュール名>_formatted.bas）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # backup-prune [--days N] [--keep K] [--force] [--json]
    p = sub.add_parser("backup-prune", help="backups の間引き＝N 日より古い控えを数えて見せる（--force で消す。COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--days", dest="days", type=int, default=None, help="これより古い控えが対象（既定 30）")
    p.add_argument("--keep", dest="keep", type=int, default=None, help="系列ごとに残す新しい控えの数（既定 1）")
    p.add_argument("--force", action="store_true", dest="force", help="実際に消す")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # status [--json]
    p = sub.add_parser("status", help="いま走っているコマンドと進み具合（COM不要・Excel に触らない）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # ---- 2026-09-17: 修理の材料 1 手・系譜・閉じたブック ----
    # repair [excel_file] <マクロ名> [--module 名] [--json]
    p = sub.add_parser("repair", help="マクロ修理の材料を 1 手で（本文・入口・呼び元呼び先・コンパイルの行名指し・"
                                      "check の error 級・直前の控えとの差分。読むだけ）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None, help="モジュール（同名が複数のとき）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # versions <ブック名> [--dir 追加フォルダ] [--json]
    p = sub.add_parser("versions", help="同名ブックの写しを日時順に（Desktop／プロジェクト／保存／エクセル／backups。"
                                        "oletools で読む＝Excel を開かない・COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dir", dest="dir_opt", action="append", default=None, help="探すフォルダを足す（複数可）")
    p.add_argument("--all", action="store_true", help="backups の控えも全部読む（既定は新しい 3 冊だけ）")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="backups の控えを読む冊数（既定 3）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # history <マクロ名> [--book ブック名] [--deep] [--max N] [--json]
    p = sub.add_parser("history", help="マクロ本文の歴史（backups・_exports・--deep で写しの .xlsm。同じ本文は畳んで隣との差分。COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--book", dest="book_opt", default=None, help="ブック名（拡張子なし）で絞る")
    p.add_argument("--deep", action="store_true", help="写しの .xlsm（versions と同じ場所）からも読む（--book が要る）")
    p.add_argument("--dir", dest="dir_opt", action="append", default=None, help="--deep で探すフォルダを足す")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="差分の表示行数の上限（既定 40）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list-file <path.xlsm> [--json]
    p = sub.add_parser("list-file", help="閉じたブックのマクロ一覧（oletools・Excel を開かない）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # grep-files "文字" <フォルダ|ファイル…> [--regex] [-i] [--max N] [--json]
    p = sub.add_parser("grep-files", help="閉じたブック横断のコード検索（フォルダかファイルを並べる。oletools・Excel を開かない）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として検索")
    p.add_argument("-i", "--ignore-case", dest="ignore_case", action="store_true", help="大文字小文字を区別しない")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定200）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # export-file <path.xlsm> [--dir 先] [--json]
    p = sub.add_parser("export-file", help="閉じたブックのモジュールを .bas/.cls/.frm に書き出す（既定 _exports/<ブック>/<日時>/。.frx は無い）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dir", dest="dir_opt", default=None, help="出力先フォルダ")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # run-macro [excel_file] <macro_name> [args...]
    p = sub.add_parser("run-macro", help="Excel内の指定されたマクロを実行します")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--json", action="store_true", help="実行結果をJSON形式で出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="実行中に出るMsgBox/InputBoxを自動応答 ok|cancel|yes|no（既定は応答しない）")
    p.add_argument("--input-text", dest="input_text", action="append", default=None,
                   help="InputBox にこの値を入れて OK で確定（複数回指定で出た順に1つずつ。"
                        "足りなければ最後の値を使い回す。VBA の InputBox() も Application.InputBox も可）")
    p.add_argument("--raw", action="store_true",
                   help="ハーネスを使わず素の Application.Run で撃つ（既定はハーネス経由＝実行時エラーを"
                        "ダイアログにせず番号と説明で持ち帰る。引数付きは常に素の Run）")
    p.add_argument("--timeout", dest="timeout", default=None,
                   help="制限時間（秒）。過ぎたら待つのをやめて失敗で戻る。マクロ自体は止められない"
                        "（VBA が応答する状態でだけ効く。Do: Loop の密なループには効かない）")

    # rehearse [excel_file] <macro_name> [args...]  （予行演習run＝コピーに試し撃ち）
    p = sub.add_parser("rehearse", aliases=["予行演習"],
                       help="マクロをコピーに試し撃ちして差分を報告（本体無傷の予行演習run）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="実行中に出るMsgBox/InputBoxを自動応答 ok|cancel|yes|no（既定は安全解除のみ）")
    p.add_argument("--input-text", dest="input_text", action="append", default=None,
                   help="InputBox にこの値を入れて OK で確定（複数回指定で出た順に1つずつ）")
    p.add_argument("--addins", action="store_true",
                   help="演習用Excelにアドイン・PERSONALを読み込む（アドインの関数を呼ぶマクロ用）")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="コピーの保存先パス（省略時は一時フォルダに ブック名_予行_日時）")
    p.add_argument("--discard", action="store_true",
                   help="報告後に結果コピーとsnapshotを削除する")
    p.add_argument("--max", dest="max_opt", default=None,
                   help="差分の表示件数（snapshot-diff と同じ・既定20）")
    p.add_argument("--timeout", dest="timeout", default=None,
                   help="制限時間（秒）。過ぎたら演習用 Excel を強制終了して報告（本体は無傷）")

    # test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]
    p = sub.add_parser("test", aliases=["テスト"],
                       help="テストSub（名前が「テスト」/test で始まる引数なしSub）を一括実行して成否一覧")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式でも出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="テスト中に出るMsgBox等を自動応答 ok|cancel|yes|no")
    p.add_argument("--input-text", dest="input_text", action="append", default=None,
                   help="テスト中の InputBox にこの値を入れて OK で確定（複数回指定で出た順に1つずつ）")
    p.add_argument("--timeout", dest="timeout", default=None,
                   help="テスト全体の制限時間（秒）。過ぎたら待つのをやめて残りは未実行にする"
                        "（生きている Excel のマクロは止められない。VBA が応答する状態でだけ効く）")

    # compile [excel_file] [--json]（全体コンパイル。check が見ない「呼ばれないマクロの中」まで見る）
    p = sub.add_parser("compile", aliases=["全体コンパイル"],
                       help="VBE の「VBAProject のコンパイル」を外から押す。"
                            "check が見ない“呼ばれないマクロの中”のコンパイルエラーを捕まえる（終了コード 0/1）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="判定を JSON 1行で出力（pass/state/detail）")

    # gate [excel_file] [絞り込み] [--module 名] [--addins] [--timeout 秒] [--keep] [--json]
    # （関所＝停止条件。コピーで check＋test を走らせ PASS/FAIL を終了コード 0/1 で返す）
    p = sub.add_parser("gate", aliases=["関所"],
                       help="関所＝停止条件。本体のコピーで構文検査＋テストSub一括実行して PASS/FAIL"
                            "（終了コード 0/1）。テスト0本は FAIL。時間切れは演習用 Excel を強制終了")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module", default=None, help="テストの対象モジュールを限定")
    p.add_argument("--addins", action="store_true",
                   help="演習用Excelにアドイン・PERSONALを読み込む（アドインの関数を呼ぶテスト用）")
    p.add_argument("--timeout", dest="timeout", default=120,
                   help="制限時間（秒・既定 120）。過ぎたら演習用 Excel を強制終了して FAIL")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="テスト中に出るMsgBox等を自動応答 ok|cancel|yes|no")
    p.add_argument("--input-text", dest="input_text", action="append", default=None,
                   help="テスト中の InputBox にこの値を入れて OK で確定（複数回指定で出た順に1つずつ）")
    p.add_argument("--keep", action="store_true", help="検分に使ったコピーを消さずに残す")
    p.add_argument("--json", action="store_true", help="判定を JSON 1行で出力（pass/reasons/tests…）")

    # replace-module [excel_file] <module_name> <bas_file>
    p = sub.add_parser("replace-module")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-module [excel_file] <module_name> [-y]
    p = sub.add_parser("delete-module", help="モジュール丸ごと削除（要約表示→確認→バックアップつき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # export-module [excel_file] <module_name>
    p = sub.add_parser("export-module")
    p.add_argument("posargs", nargs="+")

    # export-all [excel_file] [--dir 出力先] [--check]
    p = sub.add_parser("export-all", help="全モジュールを一括エクスポート（1接続）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dir", dest="dir_opt", default=None, help="出力先フォルダ（省略時はSCRIPTS）")
    p.add_argument("--check", action="store_true", help="書き出した各ファイルに check-bas 相当の検査をかける")
    p.add_argument("--history", action="store_true",
                   help="_exports/<ブック>/<日時>/ に書き出して前回との差（増えた・消えた・変わった本数）を出す（2026-09-16）")

    # form-to-vba [excel_file] <フォーム名> … UserForm を作成マクロ(.vba)に書き出す
    for _nm in ("form-to-vba", "フォーム書き出し"):
        p = sub.add_parser(_nm, help="UserForm を『VBAだけで組み立て直す作成マクロ』に変換（.frx を持ち歩かずに済む）")
        p.add_argument("posargs", nargs="*")
        p.add_argument("--all", action="store_true", dest="all_opt", help="ブック内の全 UserForm を一括で書き出す")
        p.add_argument("--out", dest="out_opt", default=None, help="出力先（省略時 _<フォーム名>_作成マクロ.vba）")
        p.add_argument("--add", dest="add_opt", default=None, help="書き出したうえで指定モジュールに追加する（例 shu002）")
        p.add_argument("--verify", action="store_true", help="捨てブックで実際に組み立てて元と照合する（元ブックには触らない）")
        p.add_argument("--name", dest="name_opt", default=None, help="作られるフォーム名（省略時は元と同じ）")
        p.add_argument("--sub", dest="sub_opt", default=None, help="作成マクロの名前（省略時 <フォーム名>作成）")

    # reorder-macro <macro_name> <up|down>
    p = sub.add_parser("reorder-macro")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true",
                   help="バックアップが取れなくても実行する（未保存の新規ブック等）")

    # --- 目コマンド ---
    # read-range [excel_file] [range ...] [--formula] [--tsv [f]] [--width N] [--json]
    p = sub.add_parser("read-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="TSVに書き出す（省略時 _last_values.tsv。編集して write-range で書き戻す往復用）")
    p.add_argument("--width", dest="width", default=None,
                   help="列の最大表示幅（既定40。超えた分は…付きで切り詰め）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定。'!'入り・記号入りシート名向け）")

    # read-selection [excel_file] [--formula]
    p = sub.add_parser("read-selection")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")

    # sheet-info [excel_file] [--preview N] [--fast]
    p = sub.add_parser("sheet-info")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シート使用範囲の先頭N行も表示（ブック俯瞰・1接続）")
    p.add_argument("--fast", action="store_true",
                   help="結合セル走査をスキップして高速表示（瞬時にシート構成とサイズのみ確認）")

    # materials [excel_file] [sheet] [--rows N]
    p = sub.add_parser("seiri", aliases=["表の整理"],
                       help="表を直す 1 手目（materials の代わり）：「表を整える」マクロを撃ち、残り（エラーセルと式・"
                            "数式と表の気づき・指示文・###）だけを出す。直す手は残りの分だけ（2026-09-13）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dedupe", action="store_true",
                   help="「重複行を消す」も撃つ（依頼かシートの指示文が重複行の削除を頼んでいるときだけ）")
    p = sub.add_parser("materials",
                       help="先回り材料：1シートの使用範囲・値（小さい表は全体＋長文セルの全文）・結合・テーブル・名前・"
                            "数式の型・エラー・図形・###・列幅を1回で出し、仕事の時計を押す（手を動かす前に見る）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--rows", dest="rows", default=None,
                   help="先頭に見せる行数（既定5。省略時、40行×30列までの表は全体を出す）")
    p.add_argument("--full", action="store_true",
                   help="200行×30列までの表は全体を出す（エージェント・採点係が使う。読み直しの往復を無くす）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（posargでも可。省略時はアクティブシート）")

    # snapshot [excel_file] [sheet] [--out file] [--sheet NAME] [--max-rows N]
    p = sub.add_parser("snapshot",
                       help="ブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力JSONパス（省略時は _last_snapshot.json）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（1シートだけ畳む。posargでも可）")
    p.add_argument("--max-rows", dest="max_rows", default=None,
                   help="1シートあたりのセル読み込み行上限（既定5000。超過は打ち切り注記）")
    p.add_argument("--no-format", dest="no_format", action="store_true",
                   help="書式を採らない（速いが、書式が消えても snapshot-diff で気づけない）")

    # snapshot-diff <before.json> [after.json] [--max N]
    p = sub.add_parser("snapshot-diff",
                       help="2つのsnapshot JSONを比較＝前後差分の検分（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_opt", default=None,
                   help="1分類あたりの表示件数上限（既定20。超過は件数のみ表示）")

    # wiring [excel_file] [--json]
    p = sub.add_parser("wiring", aliases=["配線図"],
                       help="ボタン⇔マクロの配線図（OnAction一覧＋行き先のないボタン検出）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true",
                   help="JSON形式で出力（機械処理用）")

    # screenshot [excel_file] [range] [--out file]
    p = sub.add_parser("screenshot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力PNGパス（省略時は _last_view.png）")

    # --- 手コマンド (シートの編集・整形・構造操作) ---
    # write-range [excel_file] <range> [値] [--tsv file]
    p = sub.add_parser("write-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--tsv", dest="tsv_opt", default=None,
                   help="グリッドを読み込むTSVファイル（省略時は _last_values.tsv）")
    p.add_argument("--raw", action="store_true",
                   help="数値変換せず文字列として書き込む（セル書式を文字列にする。'007'等の先頭ゼロ保持）")
    p.add_argument("--append", action="store_true",
                   help="使用範囲の最終行の次に書く（rangeは「シート名!列文字」。ログ追記用）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p.add_argument("--show", action="store_true",
                   help="書いた直後に、その範囲の見え方（画面の文字・###）を読み戻す")

    # write-cells [excel_file] <セル> <値> [<セル> <値> ...] [--tsv file] [--raw] [--sheet] [--show]
    p = sub.add_parser("write-cells",
                       help="飛び飛びのセルを1回で書く（セル 値 セル 値 …／--tsv は「セル<TAB>値」を1行1セル）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--tsv", dest="tsv_opt", default=None,
                   help="「セル<TAB>値」を1行1セルで並べたTSV（値にスペースや引用符があるとき）")
    p.add_argument("--raw", action="store_true",
                   help="数値変換せず文字列として書き込む（セル書式を文字列にする。'007'等の先頭ゼロ保持）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（セル番地と分離指定）")
    p.add_argument("--show", action="store_true",
                   help="書いた各セルの見え方（画面の文字・###）を読み戻す")

    # clear-range [excel_file] <range> [--contents|--formats|--all]
    p = sub.add_parser("clear-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--contents", action="store_true", help="値のみクリア")
    p.add_argument("--formats", action="store_true", help="書式のみクリア")
    p.add_argument("--all", action="store_true", help="すべてクリア（既定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # format-range [excel_file] <range> [書式オプション...]
    p = sub.add_parser("format-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--font")
    p.add_argument("--size")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--unbold", action="store_true")
    p.add_argument("--italic", action="store_true")
    p.add_argument("--plain", action="store_true", help="斜体・下線・取り消し線を外す")
    p.add_argument("--color")
    p.add_argument("--bg")
    p.add_argument("--number-format", dest="number_format")
    p.add_argument("--align", choices=['left', 'center', 'right', 'fill', 'justify'])
    p.add_argument("--valign", choices=['top', 'center', 'bottom'])
    p.add_argument("--wrap", action="store_true")
    p.add_argument("--border", choices=['thin', 'medium', 'thick', 'hairline', 'none'])
    p.add_argument("--col-width", dest="col_width")
    p.add_argument("--row-height", dest="row_height")
    p.add_argument("--merge", action="store_true")
    p.add_argument("--unmerge", action="store_true")
    p.add_argument("--autofit", action="store_true")
    p.add_argument("--lock", action="store_true", help="セルをロック（sheet protect 時に有効）")
    p.add_argument("--unlock", action="store_true", help="セルのロック解除（保護中も編集可に）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--show", action="store_true",
                   help="適用後の見え方（画面の文字・###）を読み戻す")

    # tidy [excel_file] <範囲|セル> [--header-from セル] [--bg 色] [--no-header] [--no-border] [--no-col-format] [--no-autofit]
    p = sub.add_parser("tidy",
                       help="表を整える（見出し書式・罫線・番号列は左寄せ・数値列は#,##0・列幅の自動調整）"
                            "→ 見え方と経過秒を読み戻す。範囲は複数可、セル1つなら表全体")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--header-from", dest="header_from", default=None,
                   help="見出し行の書式をこのセルから複写する（省略時は表の左上が見出しの体裁ならそれを行全体へ複写）")
    p.add_argument("--bg", default=None, help="見出しの背景色 #RRGGBB（既定 #DDEBF7）")
    p.add_argument("--no-header", dest="no_header", action="store_true", help="見出し行を触らない")
    p.add_argument("--no-border", dest="no_border", action="store_true", help="罫線を引かない")
    p.add_argument("--no-col-format", dest="no_col_format", action="store_true",
                   help="列の型に合わせた寄せ・表示形式（番号列は左寄せ・数値列は#,##0）を当てない")
    p.add_argument("--no-autofit", dest="no_autofit", action="store_true", help="列幅を触らない")
    p.add_argument("--min-width", dest="min_width", default=None, help="列幅の下限（既定6）")
    p.add_argument("--max-width", dest="max_width", default=None, help="列幅の上限（既定60）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # clean-table [excel_file] [範囲|セル]（2026-09-08・AI を使わない掃除）
    p = sub.add_parser("clean-table",
                       help="表を掃除する（AI なし・数秒）: 空白の全角半角と連続・半角カナ→全角・全角英数→半角・"
                            "文字の日付→日付・日付の表示形式・tidy。重複行は --delete-dups のときだけ消す。"
                            "空欄は埋めず番地で報告。範囲を省くと見出し行から広がる表")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="対象シート名")
    p.add_argument("--delete-dups", dest="delete_dups", action="store_true",
                   help="完全に重複している行を削除する（承認の言葉に当たる。既定は報告だけ）")
    p.add_argument("--no-tidy", dest="no_tidy", action="store_true", help="仕上げの tidy を当てない")

    # audit [excel_file] <範囲|セル> …（2026-09-04・仕上げ検査）
    p = sub.add_parser("audit",
                       help="表の見た目を検査する（見出し・罫線・列の型・列幅・空の見出し）。"
                            "値の正しさは見ない。指摘0件なら合格。範囲は複数可、セル1つなら表全体")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # diagnose [excel_file] [範囲|セル] [--sheet] [--json]
    p = sub.add_parser("diagnose", aliases=["データ診断", "formula-lint"],
                       help="表・シートの数式＆データ総合診断（集計漏れ・定数直書き・非一貫数式・隠れ空白・外れ値検知）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="対象シート名")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--fix", action="store_true", help="隠れ空白の除去と文字列の数字の数値化を書き戻す（-y 必須・先頭ゼロと数式のセルは触らない・集計漏れは提案だけ）")
    p.add_argument("-y", "--yes", action="store_true", help="--fix の承認（依頼に承認の語があるときだけ）")
    p.add_argument("--html", dest="html_out", default=None, nargs="?", const="default", help="診断結果を美麗な単一HTMLレポートとして書き出す")

    # style-map [excel_file] [範囲|セル] [--sheet] [--json]
    p = sub.add_parser("style-map", aliases=["書式マップ"],
                       help="表・シートの視覚レイアウト＆書式マップ（背景色・フォント・二重罫線・複合見出しツリー）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="対象シート名")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # build-sheet [excel_file] [設計図.json] [--overwrite] [--dry-run] [--sample]（2026-09-03・作る犬）
    p = sub.add_parser("build-sheet",
                       help="設計図（JSON）からシートを一枚組み上げる（見出し・数式・合計行・書式・枠固定・"
                            "入力規則・テーブル・名前定義・条件付き書式・印刷設定→tidy→読み戻し）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--overwrite", action="store_true",
                   help="同名シートに中身があっても消して組み直す（既定は止まる）")
    p.add_argument("--dry-run", dest="dry_run", action="store_true",
                   help="計画（番地・数式・仕上げ）だけ表示して Excel に触らない")
    p.add_argument("--sample", action="store_true",
                   help="設計図の例を _last_sheet_spec.json に書き出す")
    p.add_argument("--ask", default=None,
                   help="依頼文を渡して AI に設計図を 1 回書かせてから組み上げる（GEMINI_API_KEY / ANTHROPIC_API_KEY）")
    p.add_argument("--ai", default=None, help="設計図を書かせる先: claude-code（既定・ヘッドレスの Claude Code・鍵なし）| gemini | claude（API 鍵）")
    p.add_argument("--model", default=None, help="モデル名（既定 claude-code=sonnet・--forge でマクロを書く頭だけ既定 opus / gemini-3.7-flash / claude-haiku-4-5-20251001。gemini で始まる名前なら --ai 省略時も gemini）")
    p.add_argument("--new-book", dest="new_book", action="store_true",
                   help="まっさらな新しいブックに組み上げる（訓練場。人のブックに触らない）")
    p.add_argument("--recipes", action="store_true", help="手順書（帳票の型）の一覧")
    p.add_argument("--recipe", default=None, help="手順書の名前で組み上げる（名簿・備品台帳・月次集計表・請求明細・出勤簿）")
    p.add_argument("--fire", action="store_true",
                   help="手順書の自動実射: まっさらなブックに全部（または名指しのもの）を組んで検査し、保存せず閉じる")

    # set-key / clear-key（2026-09-04・APIキーの預かり。人に環境変数の画面をやらせない）
    p = sub.add_parser("set-key", aliases=["キー設定"],
                       help="APIキーを預かる: set-key [gemini|claude] [--days N]。画面には出さずに受け取り、"
                            "Windows のログインで暗号化して保管し（既定 30 日で自動削除）、疎通まで確かめる。"
                            "環境変数には書かない（平文・期限なしで残るため）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--days", default=None, help="保管する日数（既定 30。0 で期限なし）")
    p.add_argument("--model", default=None, help="疎通確認に使うモデル名（既定は --ai ごとの既定）")
    p.add_argument("--no-ping", dest="no_ping", action="store_true", help="疎通確認をしない（書くだけ）")
    p.add_argument("--stdin", action="store_true", help="キーを標準入力から読む（.bat やテスト用。画面に出ない）")

    p = sub.add_parser("clear-key", aliases=["キー削除"],
                       help="APIキーを消す: clear-key [gemini|claude] [-y]。金庫から消し、環境変数に残っていれば"
                            "それも消して読み直して確かめる（発行元のキーは生きているので、要るなら失効させること）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("-y", "--yes", dest="yes", action="store_true", help="確認を出さずに消す")

    # agent [excel_file] "依頼文" [--sheet 名] [--ai] [--model] [--max-turns N] [--dry-run] / --fire（2026-09-03・ループ）
    p = sub.add_parser("agent", aliases=["エージェント"],
                       help="依頼文をシート 1 枚に対して回す（materials→AI に 1 回聞く→書く→読み戻す→…→tidy→検査）。"
                            "AI は JSON で手（read/write_cells/write_grid/format/tidy＋inspect/view/page_setup/chart/shape/"
                            "export_csv＋行が多い表の normalize/fill/find_replace＋重い道具 table/pivot/pivot_field/pivot_calc/"
                            "slicer/powerquery/datamodel）を返すだけ。"
                            "往復の上限は --max-turns。手順書 22 本は --recipes / --recipe 名前 [補足]")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="対象シート（省略時はアクティブシート）")
    p.add_argument("--mode", default=None,
                   help="入口の指定: sheet（既存シートを直す）| build（白紙／新しいシートを組む）| macro（マクロを修理）"
                        "| both（シートとマクロが 1 つの依頼に混ざる）。"
                        "省略時は道具が決めて 1 行で言う")
    p.add_argument("--macro", default=None, help="修理するマクロ名（指定すると mode=macro）")
    p.add_argument("--new-book", dest="new_book", action="store_true",
                   help="まっさらな新しいブックに組み上げる（mode=build。人のブックに触らない）")
    p.add_argument("--rehearse", action="store_true",
                   help="mode=macro: 直したあと rehearse（コピーで試し撃ち）まで回して結果を AI に見せる")
    p.add_argument("--ai", default=None, help="聞く先: claude-code（既定・ヘッドレスの Claude Code・鍵なし）| gemini | claude（API 鍵）")
    p.add_argument("--model", default=None, help="モデル名（既定 claude-code=sonnet・--forge でマクロを書く頭だけ既定 opus / gemini-3.7-flash / claude-haiku-4-5-20251001。gemini で始まる名前なら --ai 省略時も gemini）")
    p.add_argument("--max-turns", dest="max_turns", default=None, help="往復の上限（既定 4）")
    p.add_argument("--dry-run", dest="dry_run", action="store_true",
                   help="1 往復目の手（JSON）を表示するだけで Excel には触らない")
    p.add_argument("--request-file", dest="request_file", default=None,
                   help="依頼文を UTF-8 のファイルで渡す（Excelコンボの python エンジンが使う。引用符事故を避ける）")
    p.add_argument("--recipes", action="store_true", help="手順書（sheet の定型依頼・22 本）の一覧")
    p.add_argument("--recipe", default=None,
                   help="手順書の名前を依頼文として使う。後ろに補足（承認の言葉「置き換えてよい」等・番地・仕様）を続けられる"
                        "（--recipes で一覧）")
    p.add_argument("--shake", nargs="?", const="", default=None, metavar="記録",
                   help="揺らし: 記録（既定は _last_agent_log.jsonl）の返事を Python で書き換えた版（消す手の形・消した後の番地・"
                        "書体落とし・表示形式の当て忘れ…22 通り）を作り、開いているブックに 1 通りずつ再生する（AI は呼ばない・"
                        "毎回、保存せずに閉じて開き直す）。一発合格しなかった版＝道具が吸収できない揺れ＝直す候補。"
                        "--dry-run で一覧だけ・--only 名前,名前 で絞る（2026-09-11 夜）")
    p.add_argument("--only", default=None, metavar="名前,名前",
                   help="--shake で作る版を名前で絞る（原本は必ず入る）。--exam ではお題を名前で絞る")
    p.add_argument("--exam", action="store_true",
                   help="試験: 正解の表を持つお題 5 種（顧客名簿・売上明細・会員名簿・経費精算・在庫表）を種ごとに違う汚れ方で作り、"
                        "本番と同じ道で撃って、終わった表を正解の表とセルの値で突き合わせる（道具の合格判定・採点係とは別）。"
                        "--seed で同じお題・--only で絞る・--dry-run はお題を作って見せるだけ。点数は _agent_exam.jsonl（2026-09-11 夜）")
    p.add_argument("--by-macro", dest="by_macro", default=None, metavar=".bas",
                   help="--exam を AI の代わりにマクロで撃つ（.bas の「表を整える」、依頼が消すことを承認していれば続けて「重複行を消す」）")
    p.add_argument("--with-macro", dest="with_macro", default=None, metavar=".bas",
                   help="--exam を、その .bas を読み込んだ Excel で agent に撃たせる（入口のマクロの先撃ちが効くか＝秀コンボと同じ姿）")
    p.add_argument("--imagine", type=int, default=0, metavar="N",
                   help="--shake に足す: AI に「別の AI ならこの返事をどう書くか」を N 通り書かせて版にする（1 往復・"
                        "既定は claude-code の haiku＝頭と癖の違う方。--ai／--model で変える）。Python の型に無い揺れを広げる")
    p.add_argument("--fire", action="store_true",
                   help="自動実射: まっさらなブックに練習台を組み、依頼を撃って現物で答え合わせ→PASS/FAIL→保存せず閉じる。"
                        "後ろに組（sheet / vague / macro / recipe）か弾の名前で絞れる。撃つたびに点数を記録し、"
                        "前回と比べて退行した弾を出す（履歴は --score）")
    p.add_argument("--score", action="store_true",
                   help="実射の点数の履歴を出す（何本中何本・いつ・落ちた弾。--fire のたびに記録される）")
    p.add_argument("--runs", action="store_true",
                   help="本番の走行台帳を出す（いつ・どのブック・done か・関所で何回止まったか・人が戻したか。"
                        "1 走行 1 行で溜まる。記録は _agent_logs に控えるので agent --replay で撃ち直せる）")
    p.add_argument("--forge", default=None, metavar="名前",
                   help="鍛える: 直前の走行（AI が残りをやった走行）を弾にし、その残りの型を AI にマクロ 1 本として書かせ、"
                        "撃つ前の写しに読み込んで AI の後の表と値で突き合わせる（合うまで --max-turns・既定 3）。"
                        "合格した .bas は _agent_forge に置く。--register で開いているブックの「表の整理」に足す（2026-09-17）")
    p.add_argument("--forged", action="store_true", help="鍛えたマクロの一覧（弾・合格・登録先）")
    p.add_argument("--mend", action="store_true",
                   help="修理の試験: 鍛えたマクロ（前の表と正解を持つ）を Python が種ごとに壊し（End If 抜け・変数名の打ち間違い・"
                        "For の 1 ずれ…）、写しで撃って症状（コンパイル・実行時エラー・何も書かない・数が合わない）を確かめてから、"
                        "人の言い方の症状だけを添えて macro モードに直させ、直ったマクロを前の表と別の表で正解と突き合わせる。"
                        "--seed・--only 名,名・--kinds・--dry-run（壊し方を選ぶだけ）。点数は _agent_mend.jsonl（2026-09-17 夜）")
    p.add_argument("--kinds", default=None, metavar="症状,症状",
                   help="--mend で出す症状を絞る（compile / runtime / nothing / wrong。難しい型: edge＝ある表でだけ合わない・"
                        "double＝コンパイルエラー＋数が合わない）")
    p.add_argument("--blind", action="store_true",
                   help="--mend で repair に鍛えた正本との差分を出さない（台帳に無い手書きのマクロと同じ条件で測る）")
    p.add_argument("--from", dest="mend_from", default=None, metavar="置き場",
                   help="--mend を前の回の置き場（_agent_mend\\日時）と同じお題で撃ち直す（--only で題の id かマクロ名に絞る）")
    p.add_argument("--register", action="store_true",
                   help="--forge と: 合格したマクロを「表を整える」を持つ開いているブック（--to で名指し）の標準モジュール"
                        "「表の整理」の末尾に足し、全体コンパイルまで")
    p.add_argument("--to", dest="register_to", default=None, metavar="ブック名",
                   help="--forge --register の登録先（開いているブック。省略時は「表を整える」を持つブックが 1 冊のときだけそこ）")
    p.add_argument("--truth", default=None, metavar="正解.xlsx",
                   help="--forge と: 人が直した正解のブック（同じシート名）。AI の後の姿の代わりにこれを正解にする")
    p.add_argument("--before", dest="forge_before", default=None, metavar="直す前.xlsx",
                   help="--forge と: 直す前のブック。--truth と 2 冊で弾を作る（agent の走行は要らない・依頼文は位置引数）")
    p.add_argument("--phrases", dest="forge_phrases", default=None, metavar="phrases.txt",
                   help="--forge と: 同じ仕事の言い換え（UTF-8・1 行 1 つ）。マクロの「依頼の語」が全部に当たるまで鍛える")
    p.add_argument("--test", dest="forge_tests", nargs="+", default=None, metavar="xlsx",
                   help="--forge と: 別の表でも試す（直す前 正解 の 2 冊ずつ並べる）。合格の条件に入る（AI なしで撃つ）")
    p.add_argument("--prompt", dest="forge_prompt", action="store_true",
                   help="--forge と: 会話している AI（Gemini・Claude など）が自分でマクロを書くとき（API の鍵は要らない）。"
                        "AI を呼ばずに、書き手への問い（決まり・依頼・表・前回の外れ）を _agent_forge\\名前_prompt.txt に書いて止まる（2026-09-19）")
    p.add_argument("--answer", dest="forge_answer", default=None, metavar="答え.txt",
                   help="--forge と: 会話している AI が書いた答え（Sub 全文・UTF-8 可）を 1 往復ぶん採点する（採点は道具）。"
                        "不合格なら _agent_forge\\名前_prompt.txt が次の問い（外れの説明と前回のマクロつき）に書き換わる")
    p.add_argument("--keep-case", dest="keep_case", default=None, metavar="名前",
                   help="直前の本番の走行を弾にする（そのブックを写し取って練習台にし、依頼文と一緒に控える）。"
                        "撃ち直すのは --fire mine。判定は写し取りと同じ＝正解の表は要らない")
    p.add_argument("--cases", action="store_true",
                   help="本番から拾った弾の一覧（--keep-case で登録したもの）")
    p.add_argument("--drop-case", dest="drop_case", default=None, metavar="名前",
                   help="本番から拾った弾を消す（台帳の項目と、弾の置き場にある練習台のファイル）")
    p.add_argument("--state", default=None,
                   help="実射の練習台に当てる状態を固定する（隠れ行／前のフィルタ／枠固定／名前定義／条件付き書式／"
                        "他シート参照／結合タイトル／書式ばらばら／なし）。無指定なら弾ごとに seed から選ぶ"
                        "＝きれいな表だけで撃たない（2026-09-05）")
    p.add_argument("--seed", default=None,
                   help="実射の状態を選ぶ種。落ちた組み合わせを同じ姿でもう一度撃つときに使う（--fire の出力に出る）")
    p.add_argument("--twice", action="store_true",
                   help="実射で合格した弾に、同じ依頼をもう 1 回撃つ（冪等性の検査）。済んだ仕事なので何も"
                        "変わらないのが正しい＝合計行の二重足し・列の二重生成を、正解を持たずに捕まえる。往復は 2 倍")
    p.add_argument("--repeat", default=None, metavar="N",
                   help="--fire で同じ弾を N 回ずつ撃ち、弾ごとの合格率を出す（1 発の合否では見えない揺れを数字にする。"
                        "--no-image・--grade-ai と組めば、その効果を同じ弾で測れる）")
    p.add_argument("--prune-days", dest="prune_days", default=None, metavar="N",
                   help="控えの置き場（backups）で N 日より古いファイルを数えて見せる（消すのは --force を付けたときだけ）")
    p.add_argument("--harvest", default=None,
                   help="本物のブックの構造だけを写し取った練習台を作る（--harvest 元のブックのパス）。"
                        "値は一目で偽物と分かるダミーに置き換え、構造（結合・絞り込み・隠れ行・数式・名前定義・"
                        "条件付き書式・入力規則・保護・印刷設定）は元のまま。元のブックには触らない。"
                        "自分の想像で組んだ練習台の外から弾の的を供給する口（2026-09-05）")
    p.add_argument("--keep-rows", dest="keep_rows", default=None,
                   help="--harvest で残す見出しの行数（既定 1）")
    p.add_argument("--bed", default=None,
                   help="--fire の的を、写し取った練習台（--harvest で作ったブック）にする。弾ごとの正解表を"
                        "持たない依頼（整える・重複に印・合計）を撃ち、仕上げ検査＋不変条件＋冪等性だけで判定する"
                        "＝どんな構造のブックにも撃てる。後ろにシート名を並べると、そのシートに撃つ")
    p.add_argument("--changes", action="store_true",
                   help="直前の仕事で変わったセルの明細（シート・番地・前・後）と、書式の変化"
                        "（列幅・表示形式・寄せ・太字・罫線）を出す。Excel には触らない。"
                        "実体は _last_agent_changes.tsv（1 行 1 件・Excel で開ける）")
    p.add_argument("--no-grade", dest="no_grade", action="store_true",
                   help="終わりの採点（依頼文と現物を照らして、満たしていない点を AI に挙げさせる 1 往復）を止める。"
                        "既定は採点する＝頼んだ半分で done にさせないため")
    p.add_argument("--grade", dest="grade", action="store_true",
                   help="一発で通った回（1 往復目で手が全部通り、仕上げ検査・中身・頼んでいない変化に指摘なし）でも"
                        "採点する。既定はその回だけ採点を省く（採点役は別プロセスで 8 秒・2026-09-11）")
    p.add_argument("--grade-ai", dest="grade_ai", default=None, choices=["claude-code", "gemini", "claude"],
                   help="採点だけ別の AI にやらせる（既定は同じ AI。採点はもともと別の会話で聞くが、"
                        "別の AI にすると自分の癖まで別になる）")
    p.add_argument("--grade-model", dest="grade_model", default=None,
                   help="採点に使うモデル名（--grade-ai と一緒に。省略すればその AI の既定）")
    p.add_argument("--image", dest="image", action="store_true",
                   help="書いた直後の見た目（PNG）を AI にも見せる（既定は見せない＝罫線・寄せ・列幅は仕上げ検査が文字で検査する。"
                        "画像は道具側 2.6 秒＋入力トークンの費用・2026-09-08 に既定を反転）")
    p.add_argument("--no-image", dest="no_image", action="store_true",
                   help="（互換のため残す。既定が見せないになったので付けても変わらない）")
    p.add_argument("--continue", dest="cont", action="store_true",
                   help="前回の続きとして回す（記録 _last_agent_log.jsonl の会話を引き継ぐ）。手順書が「一覧を報告」で止まった後に "
                        "agent --continue 置き換えてよい と撃つと、前に自分で挙げた候補のとおりに進む。後ろの文が補足になる")
    p.add_argument("--replay", nargs="?", const="", default=None,
                   help="API を呼ばず、記録（既定は _last_agent_log.jsonl。パスを続けると別の記録）の reply を"
                        "そのまま順に再生してループだけを回す。課金・待ち時間なしでループ側の直しを確かめる用（開発用）。"
                        "再生できるのは sheet の記録だけ。依頼文・シートは省略でき、記録の meta から取る")
    p.add_argument("--undo", nargs="?", const="latest", default=None, metavar="番号",
                   help="直前の agent が書き換えたシートを、書き換える前の控えで置き換える（そのシートだけ・保存はしない）。"
                        "--dry-run を付けると何を戻すかだけ表示する。番号（--backups の一覧）か走行の名札を続けると、"
                        "直前より前の控えに戻す（残っている控えはブックごとに 5 個）")
    p.add_argument("--backups", action="store_true",
                   help="残っている控えの一覧（番号・いつ・どのブックとシート・依頼）。--undo 番号 で戻す先を指す")
    p.add_argument("--force", action="store_true", dest="force",
                   help="--undo: 控えを取った後に build／macro を回していても戻す（既定は止まる。"
                        "後からやったぶんが消えるため）。--undo 番号 で、その後に走行のある控えに戻すときも要る")

    # sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...
    p = sub.add_parser("sheet")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--after")
    p.add_argument("--before")
    p.add_argument("--clear", action="store_true", help="tab-color のクリア")
    p.add_argument("--password", default=None, help="protect/unprotect のパスワード")

    # table [excel_file] <create|list|delete> ...
    p = sub.add_parser("table")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--no-headers", dest="no_headers", action="store_true")
    p.add_argument("--at", dest="at", default=None, help="column add の挿入位置(1始まり)")
    p.add_argument("--desc", dest="desc", action="store_true", help="sort を降順に")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="table read の結果をTSVに書き出す（省略時 _last_values.tsv）")

    # name [excel_file] <add|list|delete> ...
    p = sub.add_parser("name")
    p.add_argument("posargs", nargs="*")

    # --- 手コマンド 第2弾 ---
    # a. 編集の足回り
    p = sub.add_parser("row")          # row <insert|delete> <行番号> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("col")          # col <insert|delete> <列文字> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("copy-range")   # copy-range <src> <dst> [--values]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--values", action="store_true", help="値のみ貼り付け")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="コピー元シート名（srcと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="コピー元にシート名だけ（使用範囲全域）を許可する")
    p.add_argument("--show", action="store_true",
                   help="貼り付け後の見え方（画面の文字・###）を読み戻す")
    p = sub.add_parser("fill")         # fill <range> [--right]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--right", action="store_true", help="右方向にフィル（既定は下）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p.add_argument("--show", action="store_true",
                   help="フィル後の見え方（画面の文字・###）を読み戻す")
    p = sub.add_parser("sort")         # sort <range> [--key 列] [--desc] [--header|--no-header]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--key", help="並べ替えキー列（列文字）")
    p.add_argument("--desc", action="store_true", help="降順")
    p.add_argument("--header", action="store_true", help="先頭行を見出しとして扱う")
    p.add_argument("--no-header", dest="no_header", action="store_true", help="見出しなし")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p = sub.add_parser("autofilter")   # autofilter [range] [--off] [--column 見出し --equals 値]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--off", action="store_true", help="オートフィルタを解除")
    p.add_argument("--column", dest="column_opt", default=None,
                   help="絞り込む列（見出しの文字 か 範囲内の列番号）")
    p.add_argument("--equals", dest="equals_opt", action="append", default=None,
                   help="その列で残す値（複数指定できる。既にフィルタが付いていても掛け直す）")
    p.add_argument("--show-all", dest="show_all", action="store_true",
                   help="絞り込みだけ解除する（フィルタの▼は残す）")

    # b. 検索・置換
    p = sub.add_parser("find")         # find <文字> [--book] [--whole] [--formula]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--book", action="store_true", help="全シート横断で検索")
    p.add_argument("--whole", action="store_true", help="完全一致")
    p.add_argument("--formula", action="store_true", help="数式も検索対象にする")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（既定はアクティブシート）")
    p.add_argument("--max", dest="max_hits", type=int, default=None,
                   help="表示件数の上限（既定200）")
    p = sub.add_parser("find-replace") # find-replace <検索> <置換> [range] [--whole]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--whole", action="store_true", help="完全一致のみ置換")
    p.add_argument("--match-case", dest="match_case", action="store_true",
                   help="大文字小文字を区別する（既定は区別しない）")
    p.add_argument("--wildcard", action="store_true",
                   help="検索文字列の * ? をワイルドカードとして扱う（既定は文字どおり）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # c. ブックの開閉・保存・印刷まわり
    p = sub.add_parser("open",         # open <path>
                       help="ブックを開く（見えているExcelに合流。未起動なら通常起動＝アドインも読み込まれる）")
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("close",        # close <ブック名|path> (--save|--no-save) [-y]
                       help="開いているブックを1冊閉じる（名指し・保存方針・確認の三点セット。Excel本体は終了しない）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--save", action="store_true", dest="save_flag", help="保存して閉じる")
    p.add_argument("--no-save", action="store_true", dest="no_save_flag",
                   help="保存せずに閉じる（未保存の変更は破棄）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p = sub.add_parser("close-form",   # close-form [キャプション] [--list] [--wait 秒]
                       help="表示中の UserForm を閉じる（×ボタンと同じ SC_CLOSE。COM 不使用＝モーダル中でも効く。引数なしで全部）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", action="store_true", dest="list_flag", help="閉じずに表示中フォームの一覧だけ出す")
    p.add_argument("--wait", dest="wait_opt", default=None, help="閉じるのを待つ秒数（既定3）")
    p = sub.add_parser("vbe-reset",    # vbe-reset [excel_file] [--check]
                       help="VBE の「実行>リセット」を押す（中断モードの解除。走っている VBA は全部止まるので人の指示でだけ使う）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--check", action="store_true", help="押さずに、今リセットが要る状態かだけ返す")
    p = sub.add_parser("save")         # save [excel_file]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("save-as")      # save-as [excel_file] <path> [--overwrite]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("export-pdf")   # export-pdf <出力.pdf> [--sheet|--range]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="このシートだけをPDF化")
    p.add_argument("--range", dest="range_opt", default=None, help='この範囲だけをPDF化（例 "集計!A1:H50"）')
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("shape")        # shape --list / shape <名前>… --delete / shape <名前> --left N …
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", dest="list_shapes", action="store_true", help="図形一覧（名前・行き先マクロ・位置と大きさ）")
    p.add_argument("--delete", action="store_true", help="名指しした図形を消す（複数可。マクロ付きのボタンは拒む）")
    p.add_argument("--left", type=float, help="左位置")
    p.add_argument("--top", type=float, help="上位置")
    p.add_argument("--width", type=float, help="幅")
    p.add_argument("--height", type=float, help="高さ")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="対象シート名（既定はアクティブシート）")
    p = sub.add_parser("print-setup")  # print-setup [opts]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--area", help="印刷範囲（例 A1:H50）")
    p.add_argument("--title-rows", dest="title_rows", help="印刷タイトル行（例 1:3）")
    p.add_argument("--title-cols", dest="title_cols", help="印刷タイトル列（例 A:B）")
    p.add_argument("--landscape", action="store_true", help="横向き")
    p.add_argument("--portrait", action="store_true", help="縦向き")
    p.add_argument("--fit-wide", dest="fit_wide", help="横N ページに収める")
    p.add_argument("--fit-tall", dest="fit_tall", help="縦N ページに収める")
    p.add_argument("--zoom", help="拡大縮小率(%%)")
    p.add_argument("--center-h", dest="center_h", action="store_true", help="水平中央")
    p.add_argument("--center-v", dest="center_v", action="store_true", help="垂直中央")

    # d. 仕上げ・見た目
    p = sub.add_parser("cond-format")  # cond-format <range> --gt 100 --bg '#...'
    p.add_argument("posargs", nargs="*")
    p.add_argument("--gt"); p.add_argument("--lt")
    p.add_argument("--ge"); p.add_argument("--le")
    p.add_argument("--eq"); p.add_argument("--ne")
    p.add_argument("--between", nargs=2, metavar=("V1", "V2"))
    p.add_argument("--formula", dest="formula_opt", default=None,
                   help='数式ベースのルール（例 --formula "=B2>AVERAGE($B$2:$B$20)"）')
    p.add_argument("--bg"); p.add_argument("--color")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--duplicates", action="store_true",
                   help="重複する値を色分けする（同じ値が 2 回以上ある行を塗る）")
    p.add_argument("--clear", action="store_true", help="条件付き書式を全削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("hyperlink")    # hyperlink <cell> <url> [--text t] / --remove / --list
    p.add_argument("posargs", nargs="*")
    p.add_argument("--text", help="表示文字")
    p.add_argument("--remove", action="store_true", help="ハイパーリンク削除")
    p.add_argument("--list", dest="list_links", nargs="?", const="__ACTIVE__", default=None,
                   help="シート内の全ハイパーリンクを一覧（--list シート名 で対象指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="--remove でシート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("validation")   # validation <range> --list 'A,B,C' / --clear
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", help="ドロップダウン候補（カンマ区切り）")
    p.add_argument("--clear", action="store_true", help="入力規則を削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("freeze")       # freeze <cell> / freeze off
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("comment")      # comment <cell> <text> / --remove
    p.add_argument("posargs", nargs="*")
    p.add_argument("--remove", action="store_true", help="コメント削除")

    # 重量級(1) chart <create|list|delete>
    p = sub.add_parser("chart")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", help="column|stacked-column|stacked-column-100|bar|stacked-bar|stacked-bar-100|line|line-markers|pie|scatter|area|stacked-area|doughnut|radar")
    p.add_argument("--title", help="グラフタイトル")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="グラフ名")
    p.add_argument("--width", help="幅(pt)")
    p.add_argument("--height", help="高さ(pt)")
    p.add_argument("--pivot", help="ピボットグラフにする元のピボット名（data_range の代わり）")

    # 重量級(1b) chart-config <action> <chart名> ...
    p = sub.add_parser("chart-config")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--min"); p.add_argument("--max")
    p.add_argument("--major"); p.add_argument("--minor")
    p.add_argument("--value", action="store_true"); p.add_argument("--percent", action="store_true")
    p.add_argument("--category", action="store_true"); p.add_argument("--series", action="store_true")
    p.add_argument("--position")
    p.add_argument("--series-name", dest="series_name")
    p.add_argument("--category-range", dest="category_range")
    p.add_argument("--marker-style", dest="marker_style")
    p.add_argument("--marker-size", dest="marker_size")
    p.add_argument("--marker-fg", dest="marker_fg")
    p.add_argument("--marker-bg", dest="marker_bg")
    p.add_argument("--invert", action="store_true")
    p.add_argument("--name")

    # 重量級(2) pivot <create|list|delete>
    p = sub.add_parser("pivot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--rows", help="行フィールド（カンマ区切り）")
    p.add_argument("--cols", help="列フィールド（カンマ区切り）")
    p.add_argument("--values", help="値フィールド（カンマ区切り）")
    p.add_argument("--filter", help="フィルタ（ページ）フィールド（カンマ区切り）")
    p.add_argument("--func", help="集計方法 sum|count|average|max|min（既定 sum）")
    p.add_argument("--sheet", help="出力シート名（無ければ作成）")
    p.add_argument("--at", help="出力先セル（--sheet と組めばそのシートのそのセル。単独なら元データと同じシートの空き場所）")
    p.add_argument("--name", help="ピボットテーブル名")
    p.add_argument("--model", action="store_true", help="データモデルから作る（rows/cols は テーブル名.列名）")
    p.add_argument("--measures", help="データモデルのメジャー名（カンマ区切り。--model のとき）")

    # 重量級(2b) pivot-field <action> <pivot> <field> ...
    p = sub.add_parser("pivot-field")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--func", help="add-value/set-func の集計 sum|count|average|max|min")
    p.add_argument("--name", help="add-value の表示名")
    p.add_argument("--by", help="top-n の基準にする値フィールド（既定＝先頭の値フィールド）")
    p.add_argument("--bottom", action="store_true", help="top-n を下位 N 件にする")
    p.add_argument("--base-field", dest="base_field", help="show-as の基準フィールド（既定＝先頭の行フィールド）")
    p.add_argument("--base-item", dest="base_item", help="show-as の基準項目（既定＝前の値）")

    # 重量級(2c) pivot-calc <action> <pivot> ...
    p = sub.add_parser("pivot-calc")
    p.add_argument("posargs", nargs="*")

    # 重量級(3) slicer <add|list|delete>
    p = sub.add_parser("slicer")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="スライサー名")

    # calc-mode [manual|auto|recalc]
    p = sub.add_parser("calc-mode")
    p.add_argument("posargs", nargs="*")

    # 重量級(4) powerquery <list|refresh|add|delete>
    p = sub.add_parser("powerquery")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--m-file", dest="m_file", default=None, help="add 用 M式ファイル（省略時 _last_query.m）")
    p.add_argument("--m", dest="m_opt", default=None, help="add 用 M式をインライン指定")
    p.add_argument("--desc", default=None, help="クエリの説明")
    p.add_argument("--to", dest="to", default=None, help="load 用 読み込み先: sheet|model")
    p.add_argument("--sheet", dest="sheet", default=None, help="load --to sheet の出力先シート（省略時アクティブ）")
    p.add_argument("--at", dest="at", default=None, help="load --to sheet の左上セル（省略時 A1）")
    p.add_argument("--force", action="store_true",
                   help="load --to model: 既にモデルに載っていても作り直す（メジャー/リレーションは失われる）")

    # 重量級(5) connection <list|refresh|delete> / datamodel [list]
    p = sub.add_parser("connection")
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("datamodel")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dax", dest="dax", default=None, help="measure add 用 DAX式をインライン指定")
    p.add_argument("--dax-file", dest="dax_file", default=None, help="measure add 用 DAXファイル（省略時 _last_dax.dax）")
    p.add_argument("--desc", dest="desc", default=None, help="measure の説明")
    p.add_argument("--format", dest="format", default=None,
                   help="measure の書式: general|whole|decimal|currency|percent|scientific（既定 general）")
    p.add_argument("--decimals", dest="decimals", default=None, help="小数桁数（decimal/currency/percent/scientific、既定2）")
    p.add_argument("--thousands", dest="thousands", action="store_true", help="桁区切りを使う（whole/decimal/percent）")
    p.add_argument("--symbol", dest="symbol", default=None, help="通貨コード（currency、例: USD/JPY/EUR。グリフ$¥は不可。無効なら既定）")

    # printer-list
    p = sub.add_parser("printer-list")
    p.add_argument("posargs", nargs="*")

    # printer-setup
    p = sub.add_parser("printer-setup")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--printer", help="対象プリンター名（省略時はActivePrinterまたは既定プリンター）")
    p.add_argument("--duplex", choices=['simplex', 'vertical', 'horizontal'], help="両面印刷（simplex:片面, vertical:長辺, horizontal:短辺）")
    p.add_argument("--color", choices=['mono', 'color'], help="カラーモード（mono:モノクロ, color:カラー）")
    p.add_argument("--orientation", choices=['portrait', 'landscape'], help="用紙の向き（portrait:縦, landscape:横）")

    # check [excel_file]
    p = sub.add_parser("check", help="全モジュールの構文チェックと診断を実行します")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--all-warnings", action="store_true", dest="all_warnings",
                   help="VBM003（On Error が無い）の行も全部出す（既定は件数だけ・2026-09-17）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # inspect-gui [excel_file] [マクロ名] [--module M]
    p = sub.add_parser("inspect-gui",
                       help="自動操縦を止めるGUI境界（MsgBox/InputBox/FileDialog/モーダル）を撃つ前に洗い出す")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # capabilities [コマンド名] （COM不要）
    p = sub.add_parser("capabilities",
                       help="各コマンドが読むだけ/書き換える/消す/走らせるのどれかを出す。COM不要")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # rules （COM不要）
    p = sub.add_parser("rules", help="check の診断規則一覧（VBM001〜011）。COM不要")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # metrics [excel_file] [--top N]
    p = sub.add_parser("metrics", help="プロシージャ計量（行数/分岐/ネスト深さ）とホットスポット順位")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--top", type=int, default=None, help="表示件数（既定20）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # process [--kill PID]
    p = sub.add_parser("process", help="Excelプロセスの点呼（PID/表示/ブック数）。kill はPID名指しのときだけ")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--kill", type=int, default=None, help="このPIDを終了（ブック0冊のときだけ通す）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # check-bas <file.bas> [--fix]  (COM不要・取り込み前の単体検査)
    p = sub.add_parser("check-bas", help="取り込み前に .bas を単体検査（文字コード/改行二重化/重複）。COM不要・複数可")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--fix", action="store_true", help="改行二重化を CP932 のまま自動修正する")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # batch <コマンドファイル|->  （1接続・1プロセスでコマンド列を実行）
    p = sub.add_parser("batch", help="コマンド列を1回の接続で連続実行（ファイル or 標準入力 '-'）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--keep-going", dest="keep_going", action="store_true",
                   help="途中の失敗で止まらず最後まで実行する")

    # shell（対話セッション。接続を張ったままコマンドを打ち続ける）
    sub.add_parser("shell", help="対話セッション（接続維持のREPL。2コマンド目から再接続なし）")

    return parser


def cmd_batch(args):
    """コマンド列を1プロセス・1COM接続で連続実行: batch <file|->

    各行は通常のCLI引数列そのもの（例: `get shu003 空白行の削除`）。
    空行と # 始まりは無視。get_workbook の接続キャッシュにより全行が同じ
    COM接続を使い回すため、「1コマンド毎の再接続で数分」級の一括作業が
    数秒に縮む。各行の実行は既存コマンドの機械的な再生のみ（判断はしない）。
    """
    import shlex
    src = args.posargs[0] if args.posargs else None
    if not src:
        print("使い方: batch <コマンドファイル|->   （- で標準入力から読む）")
        print("  例: get shu003 マクロA")
        print("      replace-procedure -y")
        return False
    if src == '-':
        # パイプ経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する（shell と同じ対処）
        if not sys.stdin.isatty():
            try:
                sys.stdin.reconfigure(encoding='utf-8')
            except Exception:
                pass
        text = sys.stdin.read()
    else:
        path = smart_path_resolve(src)
        if not path or not os.path.exists(path):
            print(f"エラー: コマンドファイルが見つかりません: {src}")
            return False
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                text = f.read()
        except UnicodeDecodeError:
            # メモ帳等で CP932 (Shift-JIS) 保存されたコマンドファイルも受け付ける
            with open(path, 'r', encoding='cp932') as f:
                text = f.read()

    parser = build_parser()
    table = _command_table()
    keep_going = getattr(args, 'keep_going', False)
    total = ok_n = 0
    for lineno, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        total += 1
        print(f"----- [batch:{lineno}] {line} -----")
        try:
            # Windows パスの \ をエスケープ扱いしない（クォートは通常どおり効く）
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''
            # shlex 既定のコメント文字 '#' を無効化。行頭 # は上で処理済みで、
            # 行中の # を生かすと「テスト#1」「--bg #FF0000」の # 以降が黙って消える
            lex.commenters = ''
            tokens = list(lex)
        except ValueError as e:
            print(f"[batch:{lineno}] 引数の解析に失敗: {e}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit as e:
            if e.code in (0, None):
                # 行内の -h/--help はヘルプ表示済み。エラーではない
                ok_n += 1
                continue
            print(f"[batch:{lineno}] 引数エラー")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"[batch:{lineno}] 不明な引数/オプション: {' '.join(unknown)}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        if not sub_args.command or sub_args.command in ('batch', 'shell'):
            print(f"[batch:{lineno}] このコマンドは batch 内で実行できません")
            if keep_going:
                continue
            return False
        try:
            res = table[sub_args.command](sub_args)
        except SystemExit as e:
            # reorder-macro 等は sys.exit で終了コードを返すため、ここで吸収する
            # （code=None の素の sys.exit() は正常終了。shell 側の判定と揃える）
            res = (e.code in (0, None))
        except Exception as e:
            print(f"[batch:{lineno}] エラー: {e}")
            res = False
        if res is not False:
            ok_n += 1
        elif not keep_going:
            print(f"[batch] {lineno}行目で失敗したため停止（--keep-going で続行可）")
            print(f"===== batch 結果: {ok_n}/{total} 成功 =====")
            return False
    print(f"===== batch 完了: {ok_n}/{total} 成功 =====")
    return ok_n == total


def cmd_shell(args):
    """対話セッション: shell

    接続を張ったままコマンドを打ち続ける REPL。batch のファイル版に対する対話版で、
    get_workbook の接続キャッシュにより2コマンド目からは COM 再接続なしで動く
    （1コマンド約1秒 → 体感即応）。exit / quit / Ctrl+C で終了。
    """
    import shlex
    # パイプ/リダイレクト経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する
    # （対話（コンソール直打ち）は Windows のコンソールAPIが処理するので触らない）
    if not sys.stdin.isatty():
        try:
            sys.stdin.reconfigure(encoding='utf-8')
        except Exception:
            pass
    parser = build_parser()
    table = _command_table()
    hist_cmds = []
    print("===== vba_manager 対話セッション =====")
    print("  コマンドをそのまま入力（例: list / get マクロ名 / read-range A1:D10）")
    print("  help で使い方、history で履歴（!番号 で再実行・!! は直前）、exit で終了。")
    while True:
        try:
            line = input("vba> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n終了します。")
            break
        if not line or line.startswith('#'):
            continue
        if line.lower() in ('exit', 'quit', 'q'):
            print("終了します。")
            break
        if line.lower() == 'help':
            parser.print_help()
            continue
        if line.lower() == 'history':
            for i, h in enumerate(hist_cmds, 1):
                print(f"  {i}: {h}")
            if not hist_cmds:
                print("  (まだ履歴はありません)")
            continue
        if line.startswith('!'):
            ref = line[1:].strip()
            if ref == '!':
                idx = len(hist_cmds)
            else:
                try:
                    idx = int(ref)
                except ValueError:
                    idx = 0
            if not (1 <= idx <= len(hist_cmds)):
                print("履歴にありません。history で番号を確認して !番号 で再実行（!! は直前）")
                continue
            line = hist_cmds[idx - 1]
            print(f"vba> {line}")
        hist_cmds.append(line)
        try:
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''                      # Windows パスの \ をエスケープ扱いしない
            lex.commenters = ''                  # 行中の # をコメント扱いしない（#FF0000 等が消える）
            tokens = list(lex)
        except ValueError as e:
            print(f"引数の解析に失敗: {e}")
            continue
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit:
            continue                             # 引数エラーは argparse が表示済み
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"不明な引数/オプション: {' '.join(unknown)}")
            continue
        if not sub_args.command or sub_args.command in ('shell',):
            print("このコマンドはセッション内で実行できません")
            continue
        try:
            table[sub_args.command](sub_args)
        except SystemExit as e:
            if e.code not in (0, None):
                print(f"（終了コード {e.code}）")
        except KeyboardInterrupt:
            print("（中断しました）")
        except Exception as e:
            print(f"エラー: {e}")
    return True


_CALL_LOCAL = threading.local()   # 呼び出し台帳の入れ子の見張り（batch／agent の中から呼ばれた手は数えない）
_LOGGED_CACHE = {}                # 元の cmd_* → 包み（別名 gate／関所 が同じ包みを共有する）


def raw_command(fn):
    """台帳の包みを外して元の cmd_* を返す（配線の確認用。包みが無ければそのまま）。"""
    return getattr(fn, '__wrapped__', fn)


def _logged(fn):
    """cmd_* を呼び出し台帳つきで包む（2026-09-16）。

    外側の 1 回だけ記録する。batch／shell／agent の中から表を通して呼ばれた手は入れ子なので数えない。
    記録は call_log_write（失敗しても仕事を止めない）。台帳の名前は argparse の command（別名はそのまま）。
    """
    if fn in _LOGGED_CACHE:
        return _LOGGED_CACHE[fn]

    @functools.wraps(fn)
    def run(args):
        if getattr(_CALL_LOCAL, 'depth', 0):
            return fn(args)
        _CALL_LOCAL.depth = 1
        _vc._last_book_name = None
        t0 = time.perf_counter()
        ok = None
        name = getattr(args, 'command', None) or fn.__name__.replace('cmd_', '', 1).replace('_', '-')
        if name != 'status':
            cmd_progress_write('run', cmd=name, args=args)          # `status` が読む「始めた」（2026-09-16）
        try:
            ok = fn(args)
            return ok
        except SystemExit as e:
            ok = e.code in (0, None)
            raise
        except BaseException:
            ok = False
            raise
        finally:
            _CALL_LOCAL.depth = 0
            _main = getattr(sys.modules.get('__main__'), '__file__', '') or ''
            via = "mcp" if os.path.basename(_main).startswith('vba_mcp_server') else "cli"
            call_log_write(name, args, time.perf_counter() - t0, ok, via=via)
            if name != 'status':
                cmd_progress_write('done', cmd=name, args=args, ok=ok, sec=time.perf_counter() - t0)

    _LOGGED_CACHE[fn] = run
    return run


def _command_table():
    """コマンド名→実装の対応表（main と batch で共用）。各実装は呼び出し台帳つきで包んで返す"""
    return {name: _logged(fn) for name, fn in _raw_command_table().items()}


def _raw_command_table():
    """コマンド名→実装の対応表（包み無し）"""
    return {
        "check":             cmd_check,
        "check-bas":         cmd_check_bas,
        "inspect-gui":       cmd_inspect_gui,
        "capabilities":      cmd_capabilities,
        "rules":             cmd_rules,
        "metrics":           cmd_metrics,
        "process":           cmd_process,
        "diag":              cmd_diag,
        "setup-check":       cmd_setup_check,
        "list-open":         cmd_list_open,
        "list":              cmd_list,
        "list-modules":      cmd_list_modules,
        "list-forms":        cmd_list_forms,
        "get":               cmd_get,
        "replace-procedure": cmd_replace_procedure,
        "patch-procedure":   cmd_patch_procedure,
        "add-procedure":     cmd_add_procedure,
        "add-module":        cmd_add_module,
        "delete-procedure":  cmd_delete_procedure,
        "grep":              cmd_grep,
        "code-replace":      cmd_code_replace,
        "docs":              cmd_docs,
        "call-graph":        cmd_call_graph,
        "checkup":           cmd_checkup,
        "健康診断":            cmd_checkup,
        "test":              cmd_test,
        "テスト":              cmd_test,
        "impact":            cmd_impact,
        "影響範囲":            cmd_impact,
        "replace-module":    cmd_replace_module,
        "delete-module":     cmd_delete_module,
        "export-module":     cmd_export_module,
        "export-all":        cmd_export_all,
        "form-to-vba":       cmd_form_to_vba,
        "フォーム書き出し":       cmd_form_to_vba,
        "list-backups":      cmd_list_backups,
        "stats":             cmd_stats,
        "restore":           cmd_restore,
        "reorder-macro":     cmd_reorder_macro,
        "list-shortcuts":    cmd_list_shortcuts,
        "rename-procedure":  cmd_rename_procedure,
        "diff-module":       cmd_diff_module,
        "references":        cmd_references,
        "copy-modules":      cmd_copy_modules,
        "set-shortcut":      cmd_set_shortcut,
        "format-module":     cmd_format_module,
        "backup-prune":      cmd_backup_prune,
        "status":            cmd_status,
        "repair":            cmd_repair,          # 2026-09-17 修理の材料 1 手
        "versions":          cmd_versions,        # 2026-09-17 系譜（vbam_lineage）
        "history":           cmd_history,
        "list-file":         cmd_list_file,       # 2026-09-17 閉じたブック（oletools）
        "grep-files":        cmd_grep_files,
        "export-file":       cmd_export_file,
        "run-macro":         cmd_run_macro,
        "rehearse":          cmd_rehearse,
        "予行演習":            cmd_rehearse,
        "compile":           cmd_compile,
        "全体コンパイル":       cmd_compile,
        "gate":              cmd_gate,
        "関所":               cmd_gate,
        "read-range":        cmd_read_range,
        "read-selection":    cmd_read_selection,
        "sheet-info":        cmd_sheet_info,
        "seiri":             cmd_seiri,
        "表の整理":            cmd_seiri,
        "materials":         cmd_materials,
        "snapshot":          cmd_snapshot,
        "snapshot-diff":     cmd_snapshot_diff,
        "wiring":            cmd_wiring,
        "配線図":              cmd_wiring,
        "screenshot":        cmd_screenshot,
        "write-range":       cmd_write_range,
        "write-cells":       cmd_write_cells,
        "clear-range":       cmd_clear_range,
        "format-range":      cmd_format_range,
        "tidy":              cmd_tidy,
        "clean-table":       cmd_clean_table,
        "表の掃除":            cmd_clean_table,
        "audit":             cmd_audit,
        "diagnose":          cmd_diagnose,
        "style-map":         cmd_style_map,
        "build-sheet":       cmd_build_sheet,
        "agent":             cmd_agent,
        "エージェント":         cmd_agent,
        "set-key":           cmd_set_key,
        "キー設定":            cmd_set_key,
        "clear-key":         cmd_clear_key,
        "キー削除":            cmd_clear_key,
        "sheet":             cmd_sheet,
        "table":             cmd_table,
        "name":              cmd_name,
        "row":               cmd_row,
        "col":               cmd_col,
        "copy-range":        cmd_copy_range,
        "fill":              cmd_fill,
        "sort":              cmd_sort,
        "autofilter":        cmd_autofilter,
        "find":              cmd_find,
        "find-replace":      cmd_find_replace,
        "open":              cmd_open,
        "close":             cmd_close,
        "close-form":        cmd_close_form,
        "vbe-reset":         cmd_vbe_reset,
        "save":              cmd_save,
        "save-as":           cmd_save_as,
        "export-pdf":        cmd_export_pdf,
        "shape":             cmd_shape,
        "print-setup":       cmd_print_setup,
        "cond-format":       cmd_cond_format,
        "hyperlink":         cmd_hyperlink,
        "validation":        cmd_validation,
        "freeze":            cmd_freeze,
        "comment":           cmd_comment,
        "chart":             cmd_chart,
        "chart-config":      cmd_chart_config,
        "pivot-field":       cmd_pivot_field,
        "pivot-calc":        cmd_pivot_calc,
        "pivot":             cmd_pivot,
        "slicer":            cmd_slicer,
        "calc-mode":         cmd_calc_mode,
        "powerquery":        cmd_powerquery,
        "connection":        cmd_connection,
        "datamodel":         cmd_datamodel,
        "printer-list":      cmd_printer_list,
        "printer-setup":     cmd_printer_setup,
        "batch":             cmd_batch,
        "shell":             cmd_shell,
    }


def main():
    setup_encoding()
    parser = build_parser()
    args, unknown = parser.parse_known_args()

    # 未知オプションの黙殺はタイポを事故に変える（例: clear-range --content が
    # 「値のみクリア」でなく既定の全消し Clear() に化ける）。グローバルの
    # --visible/-v だけ許容し、それ以外の残留はエラーで止める。
    unknown = [u for u in unknown if u not in ("--visible", "-v")]
    if unknown:
        print(f"エラー: 不明な引数/オプションです: {' '.join(unknown)}")
        print("  タイプミスの可能性があります。--help で正しいオプションを確認してください。")
        sys.exit(1)

    cmds = _command_table()

    if args.command in cmds:
        ok = False
        try:
            try:
                ok = cmds[args.command](args)
            except SystemExit:
                raise
            except Exception as e:
                print(f"エラー: {e}")
                sys.exit(1)
        finally:
            cleanup_excel()
        # 明示的に False を返したコマンドは失敗(1)、それ以外は成功(0)
        sys.exit(0 if ok is not False else 1)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
