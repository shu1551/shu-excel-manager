# -*- coding: utf-8 -*-
"""vbam_agent.py — vba_manager 分割パート: 「ループ」の薄い入口（agent）

依頼文を 1 つ受けて、道具が往復を回す（2026-09-03）。Excelコンボと同じ形: 道具が先に材料を集め、
AI は「何をどこに書くか」だけを JSON で返す。道具が実行し、結果（書いた直後の見え方・コンパイルの可否）を
次の返事に付けるので、AI は自分の結果を見てから done を言う。往復の上限は --max-turns（既定 4）。

  agent "依頼文" [--sheet 名] [--mode sheet|build|macro] [--macro 名] [--new-book]
                 [--ai claude-code|gemini|claude] [--model 名] [--max-turns N] [--dry-run] [--image]
  agent --continue "補足"      前回の続き（記録の会話を引き継ぐ。手順書が報告で止まった後の承認に）
  agent --undo                 直前に書き換えたシートを、書き換える前の控えで置き換える（保存はしない）
  agent --fire [名前 ...]      自動実射（まっさらなブックに練習台を組み、依頼を撃ってセルの実物で答え合わせ）

入口の振り分け（--mode で指定できる。省略時は道具が決めて 1 行で言う）:
  sheet  既存シートを直す … materials → AI に聞く → write_cells/write_grid/format/tidy → 読み戻し＋見た目の画像 → … → 検査
  build  白紙／新しいシートを組む … build-sheet --ask と同じ（設計図を 1 回書かせて組み上げ→検査→落ちたら sheet のループで直す）
  macro  マクロを修理する・新しく作る … get・call-graph → AI に聞く → replace/code-replace/add → compile → … → 検査
  決め方: --mode > --macro > --new-book > 依頼文にマクロの語 > 依頼文に「新しいシート」等か対象シートが白紙 > sheet

sheet の AI に許す手は read / write_cells / write_grid / format / tidy に、read_sheet（他シートを読む・読むだけ。旧 inspect）／
view（枠固定・表示倍率）／page_setup（印刷設定）／chart（グラフ）／shape（図の位置と大きさ・削除。マクロ付きは触らない）／
export_csv（範囲を CSV ファイルへ。上書きしない）を足した 11 手（2026-09-04・手順書 20 本が回るように）。
同日夕、表の足回り 14 手（sort/autofilter/cond_format/validation/row・col の挿入/copy_range/find/export_pdf/
chart_config/hyperlink/comment/sheet_op/name_add）を足して 35 手。「並べ替えて」を黙って落とした実測から。
同日、行が多い表のために normalize（列に規則を当てる）／fill（数式を末尾まで）／find_replace（一括置換）を足した
（AI に行を触らせない。材料には 40 行超の表で「列プロファイル」が付く。write_grid は 100 行まで）。重い道具 7 本と合わせて 42 手（2026-09-09 に read_file を足して 43 手）
（2026-09-06 夜: read_sheet・eval・image・row_insert/row_delete/col_insert/col_delete・row_group/col_group を足した）。
名前・リンクを消す手は無い＝戻せない手は渡さない。図形は消せる（2026-09-04・手順書「図形と画像の整理」が
回るように。控えの図形を --undo で貼り戻すので戻せる。マクロ付きの図＝ボタンは拒む）。行・列は row_delete／col_delete で
消せる（2026-09-05・承認の言葉があるときだけ。--undo で戻せる）。
書けるのは対象シートの中だけ（番地にシート名は付けさせない）。
手順書（sheet の定型 22 本・2026-09-06 夜に 25 本から 20 本に切り直し、09-09 に read_file の 2 本を足した）と練習台は vbam_recipes.py（agent --recipes / --recipe 名前 [補足] / --fire recipe）。
保存はしない（気に入らなければ保存せずに閉じる＝元に戻す）。最初の書き込みの前に控え（SaveCopyAs）を取るので
`agent --undo` でそのシートだけ戻せる。終わりに「変わったセル」の**明細**（番地・前・後）を道具が出し、全部を
_last_agent_changes.tsv に残す（後から見るのは `agent --changes`）＝AI の報告を信じずに、戻さずに検分できる。
書いた直後の見た目（PNG）は既定では AI に見せない（2026-09-08 に反転。罫線・寄せ・列幅・### は仕上げ検査が文字で検査して
差し戻すので絵は要らず、画像は道具側 2.6 秒＋入力トークンの費用だけだった。Excelコンボの窓は画像なしで速く正確）。--image で見せる。
請求書（送り字数・返り字数・待ち秒・トークン）を往復ごとと合計で出す＝費用は道具が数える。
"""
import os
import re
import io
import difflib
import unicodedata
import sys
import json
import time
import argparse
import contextlib
import csv
import datetime
import hashlib
import threading
import subprocess
import shutil
import queue
import atexit
import urllib.request
import urllib.error

from vbam_core import (SCRIPT_DIR, BACKUP_DIR, LAST_PROC_FILE, get_workbook, parse_target_and_rest,
                       job_clock_elapsed, job_clock_get, job_clock_set, job_clock_start,
                       _LAST_VALUES_FILE, _com_is_busy, _col_letter, _get_active_excel)
from vbam_recipes import (SHEET_RECIPES, RECIPE_CASES, recipe_when, compose_recipe_request,
                          _RECIPE_MAX_TURNS)
from vbam_vba import _all_procedure_names, _suggest_similar
from vbam_build import (_AI_DEFAULT_MODEL, _AI_TIMEOUT, _extract_json, plan_sheet, _apply,
                        ask_design, _plan_summary, _verify as _verify_build, _book_materials,
                        _api_call_with_retry)
from vbam_keys import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_ai import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_hands import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_inv import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_ledger import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_undo import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_grade import *  # noqa: F401,F403  (2026-09-11 分割)
from vbam_view_ai import *  # noqa: F401,F403  (2026-09-11 分割)

_LAST_AGENT_ASK_FILE = os.path.join(SCRIPT_DIR, '_last_agent_ask.txt')     # 1 通目（検分用）
_LAST_AGENT_LOG_FILE = os.path.join(SCRIPT_DIR, '_last_agent_log.jsonl')   # 往復の全記録
_AGENT_LOCK_FILE = os.path.join(SCRIPT_DIR, '_agent_running.lock')         # 同時実行の関所（記録が混ざるのを防ぐ）

_DEFAULT_MAX_TURNS = 4
_GATE_EXTRA_TURNS = 4         # 関所の差し戻しで足せる往復（audit / inv / coverage / grade の 4 つ・各 1 回）
_RESULT_LIMIT = 6000          # 1 つの結果を AI に見せる上限（字）
_RESULT_LIMIT_BIG = 20000     # 読む手（read / inspect / normalize の結果 …）はもっと見せる（2026-09-04・大きい表）
_READ_ROW_LIMIT = 80          # read の丸読み上限（行）
_RESULT_BIG_LABELS = ('read', 'inspect', 'read_sheet', 'normalize', 'table read', 'pivot_calc get_data', 'get ')
_WRITE_GRID_MAX_ROWS = 100    # write_grid で AI が書ける行数の上限。多い行は fill / normalize（AI に行を触らせない）
# 表の足回り（2026-09-04・「並べ替えて」を黙って落とした実測から）。CLI をそのまま呼ぶ。消す向きは通さない
_TABLE_OPS = ('sort', 'autofilter', 'cond_format', 'validation', 'row', 'col', 'copy_range',
              'find', 'export_pdf', 'chart_config', 'hyperlink', 'comment', 'sheet_op', 'name_add',
              # 2026-09-06 夜: 挿入と削除を名前で割った（row の action に隠れていた「消す」を手の名前に出す）。
              # row／col（action 付き）も古い記録の replay のために受ける
              'row_insert', 'row_delete', 'col_insert', 'col_delete')
_ALLOWED_OPS = ('read', 'read_sheet', 'eval', 'write_cells', 'write_grid', 'format', 'tidy',
                'inspect', 'view', 'page_setup', 'chart', 'image', 'shape', 'export_csv', 'read_file',
                'normalize', 'fill', 'find_replace', 'clear_range', 'row_group', 'col_group', 'dedupe') + _HEAVY_OPS + _TABLE_OPS
# 道具が COM で直接やる手。read_sheet は inspect の新しい名前（read＝範囲・read_sheet＝シート・view＝表示、と
# 境を名前で言う・2026-09-06 夜）。eval＝式を計算して結果を見るだけ（セルに書かない＝自由コードの逃げ道の
# 読み取り専用版）。image＝写真を置く。row_group／col_group＝隠さずに畳む（非表示にする手は無い）
_DIRECT_OPS = ('inspect', 'read_sheet', 'eval', 'view', 'page_setup', 'chart', 'image', 'shape', 'export_csv',
               'read_file', 'normalize', 'fill', 'row_group', 'col_group', 'dedupe')
_ALIAS_OPS = ('inspect', 'row', 'col')   # 古い名前（read_sheet／row_*・col_* の前）。規則文には出さず、古い記録の --replay のために受けるだけ

# format の許す項目 → format-range のオプション（値つきは (名, True)）
_FORMAT_KEYS = {
    'bold': ('--bold', False), 'unbold': ('--unbold', False), 'italic': ('--italic', False),
    'plain': ('--plain', False),
    'color': ('--color', True), 'bg': ('--bg', True), 'font': ('--font', True), 'size': ('--size', True),
    'number_format': ('--number-format', True), 'align': ('--align', True), 'valign': ('--valign', True),
    'wrap': ('--wrap', False), 'border': ('--border', True),
    'col_width': ('--col-width', True), 'row_height': ('--row-height', True),
    # 結合の解除・結合（2026-09-04）。帳票を一覧に直す手順が、結合を外せずに元の場所では回らなかった。
    # 控え（--undo）があるので戻せる＝渡してよい手
    'unmerge': ('--unmerge', False), 'merge': ('--merge', False),
}
# format の値が決まっている項目（vba_manager.py の format-range の choices と同じ並び）。
# 外れた値（align:"middle"・border:"all"）は argparse が SystemExit(2) でループごと落としていた（2026-09-04）。
_FORMAT_CHOICES = {
    'align': ('left', 'center', 'right', 'fill', 'justify'),
    'valign': ('top', 'center', 'bottom'),
    'border': ('thin', 'medium', 'thick', 'hairline', 'none'),
}

RULES = """あなたは Excel のシートを直す係です。道具（Python）が材料を集め、あなたの返事どおりにセルへ書き、
書いた直後の見え方を次の【結果】で見せます。返事は JSON だけ（前後の文・コードフェンスは付けない）。

**必ず守る 5 つ**（他の規則より先。破ると道具が差し戻します）:
 1. **頼まれたことは全部やるか、できない分を必ず report に書くか**のどちらか。できる分だけやって黙るのは禁止。
 2. **人の値を消す・置き換える手は、依頼・補足・手順書に承認の言葉があるときだけ。** 無ければ手を出さず、
    候補を report に書いて人に聞く。承認の言葉＝「消してよい・置き換えてよい・上書きしてよい・埋めてよい・
    切ってよい・整理してよい・直してよい」など、**人が「元の値が変わってよい」と言っている言葉**。
    「見やすくして」「並べ替えて」「集計して」は、その仕事をせよという指示であって、値を消す承認ではない。
 3. **他のセルから導ける値（合計・平均・割合・前月比・突き合わせの結果）は必ず数式で書く。**
    自分で計算した数を値で置かない（人が元の行を直しても黙って古いままになる）。
 4. **率・単価・しきい値のような前提は、ラベルを付けたセルに置いて参照する。** 数式に直接埋めない（=B5*0.21 にしない）。
 5. **手は上から順に実行し、1 つ失敗すると残りは実行しません。** だから read と安全な手を先に、
    人の表を書き換える手を後に並べる。失敗したら、直した手と残りを一緒に並べ直す。

返事の形:
{"say": "いま何をするかを一言",
 "plan": [{"item": "依頼を分解した 1 つ（短く）", "state": "未"}],
 "actions": [ ... ],
 "done": false,
 "report": "終わったときの報告（done が true のとき。手と一緒に done を返すなら、全部の手が通った前提の報告）"}

**手と done は同じ返事に載せてよい**（往復が 1 回減る）: 並べた手が全部通る見込みなら "actions" と一緒に
"done": true と report を返す。全部通れば道具は次の往復を待たずに検査へ進む。1 つでも失敗すれば【結果】が返るので、
直して続ける（done は取り消される）。読んでから決める手・失敗しそうな手は、今までどおり done を付けずに結果を見る。

plan は 1 往復目に必ず書く（依頼を分解した項目を短文で並べる。state は「未」で始める）。
2 往復目からは同じ項目の state を毎回返す。state は 4 つ:
  未 … まだやっていない（1 つでも残っている done は道具が受け付けない）
  済 … やった
  不可 … 当てる手が無い・材料から番地が決められない（report に「人がやるなら何をするか」を書く）
  要判断 … **人に聞かないと決められない**（消してよいか・どちらの列が正か・重複をどうするか）。
        このときは **安全側で仮に処理してから** 要判断にする（消さずに印を付ける・元の列を残して隣に
        整形列を作る・並べ替えずに件数を数える）。その項目に "ask": "人への質問（1〜2 文）" を書く。
        道具が report の末尾に質問としてまとめ、人は agent --cont で答えます。**要判断は done を塞ぎません。**
**項目は減らさない。** 現場を見て分かった項目（文字列の数値が混ざっていた・見出しが 3 行目だった）は
足してよい。足した項目には "why": "足した理由（1 文）" を付ける（道具が report の末尾に写す）。

report は 3 段で書く（【やったこと】は書かない＝実行した手から道具が作って先頭に付ける。書いてもよい）:
 【できなかったこと】対象・理由・人がやるなら何をするか
 【決めたこと】仕様が書かれていない所を自分で決めた分（「重複の鍵は会員番号にした」など）
 【確認していないこと】読んでいない範囲・前提にしたこと
（要判断・落とし物・頼んでいない変化は、道具が report の末尾に足します。自分で書かなくてよい）

actions に並べられる手（この 44 個だけ）:
 {"op":"read","range":"A1:D10"}                       範囲の値を読む（数式で見たいときは "formula": true。他のシートは "sheet":"名前" を足す）。**材料の「全体（値）」に出ている範囲は読み直さない**（同じ格子が返るだけで往復を 1 回失う。「前にやったこと」の記録と格子が食い違っていても、格子のほうが今の現物）
 {"op":"read_sheet","sheet":"名簿"}                   シート 1 枚の材料を読み直す（他のシートも可。読むだけ。範囲を読むのは read・表示を変えるのは view）
 {"op":"eval","formula":"=SUMPRODUCT((B2:B300>0)*C2:C300)"}   式を計算して結果を見るだけ（セルには書かない。期待値の検算・件数の確認に。"formulas": [...] で 20 本まで一度に）
 {"op":"write_cells","cells":{"C7":"000-0002","C11":"=A7*2"}}   飛び飛びのセルに値・数式を書く
 {"op":"write_grid","range":"F5","rows":[["見出し","金額"],["a","=B6*2"]]}   左上を起点に表を書く（行＝配列）
   ※ この 2 つは、**書き先にもとからある値があると道具が手ごと止めます**（1 セルでもあれば実行しません）。
     空いている所へ書くなら番地を変える。置き換えてよいなら "overwrite": true を足す（承認の言葉があるときだけ）。
     同じ値の書き直しは止まりません（仕上げ検査の差し戻しで書き直すのは自由）。
 {"op":"format","range":"A5:E5","bold":true,"bg":"#DDEBF7","align":"center","number_format":"#,##0","col_width":12,"row_height":18,"wrap":true,"border":"thin","unmerge":true}   見た目を当てる（見出しの色・寄せ・桁・列幅。**フォントのばらつきをそろえるなら、表の本文全体に font・size・color と "plain": true（斜体・下線・取り消し線を外す）・"bg": "none"（塗りを消す。白で塗らない）・"unbold": true を当てる**。結合の解除は "unmerge": true、結合は "merge": true。align は left|center|right|fill|justify、valign は top|center|bottom、border は thin|medium|thick|hairline|none。この語以外は道具が断る＝"middle" や "all" は無い）。**条件で色を分けるとき（未入力・期限切れ・0 以下・しきい値超え）は format の塗りを使わず cond_format を使う**。format の塗りは条件に関係なく固定なので、値が変わっても並べ替えても、色だけがその番地に取り残される
 {"op":"tidy","ranges":["A5:F13"]}                    表の仕上げ（見出し・罫線・番号列は左寄せ・数値列は #,##0・列幅）。範囲は見出し行から材料の「データの末尾」まで（使用範囲が膨張していても、空の列・行は含めない＝空の列に罫線が付く）。**表を書いた・掃除した・行や列を消した回は、同じ返事の最後にこの手を並べ、手と一緒に "done": true を返す**（tidy だけを次の往復に分けない＝sonnet で 15 秒の無駄。「罫線を引き直して」の依頼は tidy 1 手で満たす＝format の border を別に並べない。当て忘れると仕上げ検査が「見出しが本文と同じ体裁」「罫線がありません」で差し戻す＝往復を 1 回失う）
 {"op":"clear_range","range":"A1:F30","what":"contents","overwrite":true}   範囲の中身を消す（what: contents＝値だけ（既定）| formats＝書式だけ | all＝両方）。人の表を消すので "overwrite": true が要る（依頼・補足・手順書に「消してよい・置き換えてよい」があるときだけ）。行・列そのものを消すのは row_delete／col_delete
 {"op":"view","freeze":"A2","zoom":100}               枠固定（そのセルの左上で固定。"off" で解除）と表示倍率（片方だけでも可）
 {"op":"page_setup","area":"A1:L41","title_rows":"1:1","landscape":true,"fit_wide":1,"fit_tall":0,"footer_page":true,"center_h":true}   印刷設定（fit_tall 0＝縦は伸ばす。footer_page＝フッターにページ番号）
 {"op":"chart","range":"F1:G4","type":"column","title":"品目別 金額（円）","at":"I1","width":360,"height":216}   グラフ（type: column|bar|line|pie|scatter|area）。ピボットグラフは range の代わりに "pivot":"ピボット名"（ピボットのシートに置く）
 {"op":"image","path":"C:/写真/外観.jpg","at":"G2","width":240}   写真・画像ファイルをセルの左上に置く（png/jpg/gif/bmp。width か height の片方で縦横比を保って縮小。ファイルが無ければ道具が断る＝依頼にパスが無いときは「不可」にして report で聞く。消すのは shape の delete）
 {"op":"shape","name":"案内1","left":50,"top":50,"width":150,"height":60}   図形の位置と大きさ。消すなら {"op":"shape","names":["重なった要らない図","はみ出した矢印"],"delete":true}（複数まとめて可。控えから戻せる＝agent --undo で復活する）。**「要らない図形」「マクロの付いていない図形」を消せと言われたら {"op":"shape","names":"all_without_macro","delete":true}**（名前を写さない＝写し間違いで止まらない）。マクロが割り当たっている図（ボタン）は動かすのも消すのも道具が拒む＝材料の「図形・ボタン」で → の付いた図は触らない
 {"op":"export_csv","range":"A1:D7","path":"C:/出力/一覧.csv","encoding":"utf-8-sig","date_format":"yyyy-mm-dd"}   範囲を CSV ファイルへ（encoding: utf-8-sig|utf-8|cp932。既にあるファイルには書かない）
 {"op":"read_file","path":"C:/受領/旧システム.csv","to":"取込_旧システム"}   **別のファイルを読んで、このブックの新しいシートに写す**（csv/tsv/txt・xlsx/xlsm。文字コードは道具が当てる。先頭ゼロは文字のまま守る）。元のファイルには一切書かない＝読むだけ。他のブックの表と突き合わせる依頼（旧システムの出力と照らす・対応表で振り替える）は、まずこの手で隣に置いてから、いつもの手（fill の突合の式・tidy）で処理する。"from_sheet" でブックの中のシートを選ぶ（省略＝先頭）。"to" を省くとファイル名から付ける。読む行の上限は "limit"（既定 20000）。**依頼にパスが無いときは自分で作らず、report で人に聞く**
 --- 行が多い表の手（道具が全行を処理する。AI は列と規則だけ決める） ---
 {"op":"normalize","range":"C2:C300","rules":["trim","hankaku","hyphen","number","date"],"to":"H2","header":"電話（整形）"}   列に規則を順に当てて整える。rules: trim（前後の空白を取る。空白だけのセルは空に）| hankaku（全角の数字・英字・記号→半角）| hyphen（－ ‐ − ― や数字に挟まれた ー →-）| number（数値に見える文字→数値。¥・円・カンマ・「万」・後ろの単位（個・本・枚・件）も読む＝95万円→950000・12個→12。先頭ゼロの番号「0001」は文字のまま）| date（2026/1/5・2026-01-05・2026年1月5日・令和8年3月1日・R8.3.1・March 15, 2026 の文字→日付）| {"regex":{"search":"^(\\\\d{3})(\\\\d{4})$","replace":"\\\\1-\\\\2"}}（正規表現の置換。7 桁の郵便番号→000-0000 など）| as_text（控え用に、値を文字のまま写す）| fill_down（空白を上の値で埋める）| lower（英字を小文字に。メールアドレス）| phone（電話番号を 03-1234-5678 の形に。**市外局番の区切りは道具が決める**＝同じ列にハイフン付きの例があればそれに合わせ、無ければ 03・06・011・045・052・075・078・092・携帯・0120 だけ表で決める。+81 は 0 に。決まらない番号はそのまま残して番地を返す＝report に書く。**電話番号を regex で区切らない**（052 を 05-2… と切った）)| postal（郵便番号を 000-0000 に。〒・全角・区切りなし・空白も読む）| corp（会社名の（株）・㈱→株式会社・(有)→有限会社にそろえ、会社の種類の前後の空白を取る。前株・後株の位置は変えない＝**会社名の表記ゆれは regex でなくこれ**）| kana_zenkaku（半角カナ→全角カナ。濁点・半濁点も合成）| space_zenkaku（文字のあいだの空白の連なり→全角空白 1 つ。姓と名の間）| space_hankaku（同・半角空白 1 つ）。rules が空なら写すだけ。to＝書き先の左上（右の整形列。header で見出し）。元の列に書くなら to を省いて "overwrite": true（依頼・補足・手順書に承認の言葉があるときだけ）。数式のセルは触らない。結果は件数・例・結果の列の様子で返る
 {"op":"fill","range":"F2:F300","formula":"=IFERROR(VLOOKUP(A2,$I$2:$K$60,3,FALSE),\\"申込なし\\")"}   先頭セルに数式（か "value"）を書いて末尾まで伸ばす（相対参照は行ごとにずれる。固定したい範囲は $ を付ける）。"right": true で右へ。書き先に値があれば道具が断る（"overwrite": true で上書き）
 {"op":"find_replace","range":"C2:C300","search":"－","replace":"-","overwrite":true}   範囲の一括置換（"whole": true で完全一致）。元の値を書き換えるので overwrite が要る。置き換え後が探す文字を含む置換（受注済→受注済み）は、道具がセルごとに見て、もとから置き換え後になっているセルを飛ばす（二重にならない＝whole を足さなくてよい）
 --- 表の足回り（並べ替え・絞り込み・色分け・入力規則・挿入・複写。2026-09-04 に追加） ---
 {"op":"sort","range":"A7:F19","key":"B","desc":true}   範囲を並べ替える（key＝並べ替える列の列文字。desc＝降順＝新しい順・大きい順。先頭行は見出し扱い。見出しの無い範囲だけ "header": false）
 {"op":"autofilter","range":"A7:F19"}                   オートフィルタ（絞り込みの▼）を付ける（"off": true で外す）
 {"op":"autofilter","range":"A1:N307","column":"件名","equals":"観光関係一般"}  「◯◯だけ抜き出して」＝その値の行だけ表示（equals は配列も可。既にフィルタが付いていても掛け直す。"show_all": true で絞り込みだけ解除）
 {"op":"cond_format","range":"A2:A304","duplicates":true}  重複する値（2 回以上ある値）を色分け＝「同じ行が二重に入っている」の答え。行は消さない
 {"op":"cond_format","range":"A8:F19","formula":"=$D8=\\"\\"","bg":"#FFC7CE"}   条件付き書式。条件は 1 つだけ（gt/lt/ge/le/eq/ne | "between":[下,上] | formula）＋見た目（bg/color/bold）。formula は範囲の左上の行を基準に書く（列を固定するなら $D8 のように $）。行が動いても崩れないので、色分けは format の塗りでなくこちらを使う
 {"op":"validation","range":"C8:C19","list":["佐藤","鈴木"]}   入力規則（ドロップダウン。カンマ区切りの文字列でも可）
 {"op":"row_insert","at":7,"count":1}                  行の挿入（at＝その行の上に入る）
 {"op":"row_delete","at":7,"count":2,"overwrite":true}   行の削除（人の行を消すので承認の言葉があるときだけ。--undo で戻せる）。同じ返事の番地は全部、返事を出す前のいまのシートの番地で書く（道具が消す手を書く・整える手の後ろへ回し、下の行から消す）
 {"op":"dedupe","range":"A5:I56","overwrite":true}   **重複行を消すのはこの手**（行番号を並べて row_delete しない）。range＝見出し行から表の末尾まで。**keys は省くのが既定**＝材料の「重複行」と同じ照合で消す（**重複＝全部の列が同じ行**。書き方の違い＝空白・全角半角・大文字小文字・ハイフン・+81・（株）は道具がならして比べる）。**1 列でも中身が違う行（備考だけ違う等）は重複ではない＝消さない**（材料の「重複の疑い」に出る。どちらを残すかは人の判断＝report の【人に判断してほしいこと】へ）。keys を書いても、全部の列が同じでない行は道具が消さない。道具が先に出てくる行を残して後の行を下から消す。材料の「重複行」は同じ照合で数えてある。同じ返事の書く・整える手は、消す前のいまの番地で書いてよい（道具が dedupe を最後に回す。後ろの tidy は消した後の表全体に当たる）
 {"op":"col_insert","at":"C","count":1,"header":"部署"}   列の挿入（at＝その列の左に入る。既存の列は右へずれる）。header を付けると見出し行（材料の「見出し行の推定」）に見出しを書く＝同じ名前の見出しがその行にすでにあれば道具が止める（同じ列を二重に足す事故）
 {"op":"col_delete","at":"C","count":1,"overwrite":true}   列の削除（同上）
 {"op":"row_group","from":5,"to":12}                   行をグループ化して畳めるようにする（隠す手は無い＝見せたくない明細はこれで畳む。"off": true で解除）
 {"op":"col_group","from":"C","to":"F"}                列のグループ化（同上）
 {"op":"copy_range","src":"A7:F19","dst":"H7","values":true}   範囲の複写（values＝値だけ貼る）
 {"op":"find","text":"貸出中","whole":false}            探す（読むだけ。"book": true で全シート・"formula": true で数式も）
 {"op":"export_pdf","path":"C:/出力/貸出簿.pdf","range":"A1:F19"}   PDF へ書き出す（range を省くとシート全体。既にあるファイルには書かない。印刷の形は先に page_setup で決める）
 {"op":"chart_config","action":"legend","chart":"グラフ 1","args":["bottom"]}   グラフの仕上げ（action: legend|set-title|set-axis-title|axis-scale|axis-format|gridlines|style|set-type|set-source|placement|data-labels。chart＝材料の「グラフ」に出ている名前。軸目盛は "min"/"max"/"major"/"minor"、ラベルは "value"/"percent"/"category"/"series"）
 {"op":"hyperlink","cell":"A1","url":"https://…","text":"一覧へ"}   ハイパーリンク
 {"op":"comment","cell":"D8","text":"未返却"}           セルのコメント
 {"op":"sheet_op","action":"add","name":"集計"}         シートを足す／名前を変える／複写／タブ色（action: add|rename|copy|activate|tab-color。rename・copy・tab-color は "to" に新しい名前か色 #RRGGBB。置き場所は "after"/"before"。足したシートには、同じ返事の後ろの手で "sheet":"名前" を付けて書ける）
 {"op":"name_add","name":"貸出簿","range":"A7:F19"}     名前定義を足す
 --- 重い道具（テーブル・ピボット・スライサー・パワークエリ・データモデル。名前は材料の「テーブル」「ピボット」にあるものを使う） ---
 {"op":"table","action":"create","range":"A3:E9","name":"T名簿"}   見出しつきの範囲をテーブルに（action: create|list|read|sort|filter|filter_values|filter_clear|column_add）
   read {"name":"T名簿"} ／ sort {"name":"T名簿","column":"金額","desc":true} ／ filter {"name":"T名簿","column":"金額","criteria":">10000"} ／ filter_values {"name":"T名簿","column":"コース","values":["年間","半年"]} ／ filter_clear {"name":"T名簿"} ／ column_add {"name":"T名簿","column":"備考","at":3}
 {"op":"pivot","action":"create","range":"I5:K10","rows":"申込コース","cols":"","filter":"","values":"金額","func":"sum","sheet":"集計","at":"A3","name":"P申込"}   ピボットを作る（rows/cols/filter/values は見出し名。複数はカンマ区切り。同じ列を 2 本値に置くなら "values":"金額,金額"。func: sum|count|average|max|min。置き場所は別シート "sheet"（無ければ作る。at はそのシートの中のセル）。元の表と同じシートに置くときだけ sheet を省いて空き場所を at で（表の上や隣は道具が断る）。データモデルから作るなら "model":true で range を省き、rows/cols は "テーブル名.列名"、値は "measures":"メジャー名"。action: create|list）
 {"op":"pivot_field","action":"add_value","pivot":"P申込","field":"金額","func":"count","name":"件数"}   ピボットのフィールド（action: list|add_row|add_col|add_filter|add_value|remove|set_func{func}|set_name{name}|set_format{format:"#,##0"}|set_filter{values:[…]}|sort{order:asc|desc}|group_date{by:days|months|quarters|years}|group_numeric{start,end,interval}|show_as{kind:percent_total|percent_row|percent_col|percent_parent_row|percent_parent_col|diff_from|percent_diff_from|running_total|percent_running_total|rank_desc|rank_asc|index|none, base_field, base_item}|top_n{n, by:値フィールド名, bottom}|filter_clear|position{n}）。show_as の field は値フィールド名（"sum/金額" の形か元の列名）。構成比・前月比・累計は同じ列を 2 本値に置いて片方に show_as（diff_from の基準は既定で先頭の行フィールドの前の項目）。sort の field は行フィールド名で、"by":"sum/金額" を付けると値の順（多い順は order:desc）、無ければラベル順。前の値との差・累計（位置参照）は値の並べ替え・上位 N と同時に使えない（Excel の制約。両方頼まれたら sort/top_n を先に当て、show_as は最後＝道具が並べ替えを外して言う。report にその旨を書く）
 {"op":"pivot_calc","action":"get_data","pivot":"P申込"}   ピボットの中身を読む／calc_field_create{name,formula:"=金額*1.1"}／calc_field_list／layout{layout:compact|tabular|outline}／subtotals{field,on}／grand_totals{which:rows|cols|both,on}（Excel の名前は見た目と逆向き。rows＝行ごとの総計＝**右端の列**／cols＝列ごとの総計＝**一番下の行**）／refresh／style{style:"PivotStyleMedium9"}／repeat_labels{on}／empty_as{text:"0"}／set_source{range}
 {"op":"slicer","action":"add","source":"P申込","field":"申込コース","at":"M12","name":"S申込コース"}   スライサー（source はピボット名かテーブル名。action: add|list）
 {"op":"powerquery","action":"add","name":"Q名簿","m":"let ソース = Excel.CurrentWorkbook(){[Name=\"T名簿\"]}[Content] in ソース"}   パワークエリ（action: list|refresh{name}|add{name,m}|edit{name,m}|load{name,to:sheet|model,sheet,at}）。add は接続だけ＝表に出すには load {"to":"sheet","at":"M3"} を続ける
 {"op":"datamodel","action":"measure_add","table":"T売上","name":"売上合計","dax":"SUM('T売上'[金額])","format":"whole","thousands":true}   データモデル（DAX のテーブル名は 'T売上' とシングルクォートで囲む。action: list|relation_add{fk_table,fk_column,pk_table,pk_column}|measure_add{table,name,dax,format:general|whole|decimal|currency|percent,decimals,thousands,symbol}）。モデルにテーブルを載せるのは powerquery load {"to":"model"}

言葉の決め:
 - **承認の言葉** … 人が「元の値が変わってよい」と言っている言葉（消してよい・置き換えてよい・上書きしてよい・
   埋めてよい・切ってよい・整理してよい・直してよい）。依頼文・補足・手順書のどれにあってもよい。
   並べ替える・書式を整える・列を足す・空いた所に書く・別の列に整形して写すのは、承認が要りません
   （元の値が消えないので）。承認が要るのは、人の値が消えるか書き換わる手だけ＝
   clear_range / row_delete・col_delete / find_replace / normalize・fill の overwrite / shape の delete / 非空セルへの write。
   **承認の言葉があるかは道具も依頼文で確かめる**＝無いのに "overwrite": true や delete を返しても、その返事の手は
   1 つも実行されない（候補を report に書き、plan を「要判断」にして人に聞く）。
 - **表全体** … 見出し行から最後のデータ行まで（tidy を当てる範囲）。上のタイトル行・隣の表は含めない。
   **自分が足した合計行・合計列は含める**（外すと罫線が継ぎはぎになり、仕上げ検査が差し戻す。
   規則が「下の合計行は含めない」と言い、検査が「合計行に罫線が無い」と咎めていた・2026-09-06 深夜）。
   見出し行は材料の「見出し行の推定」を見る（**1 行目とは限らない**）。
 - **見出し行** … 材料の「見出し行の推定」の行。並べ替え・テーブル化・フィルタ・tidy の範囲はこの行から始める。

規則:
 - 手の "range" に書く番地はこのシートの中だけ（"A1" や "A5:E13"。番地にシート名は書かない。列全体 "A:A" も不可）。
   **数式の中身はこの制限を受けない**（=SUM(A:A) や ='集計'!B2 と書いてよい）。
 - **別のシートに書くときは、その手に "sheet": "シート名" を足す**（例 {"op":"write_grid","sheet":"集計","range":"A1","rows":[…]}）。
   材料の「他のシート」に出ている名前だけ。無いシートには書けないので、先に {"op":"sheet_op","action":"add","name":"集計"} で作り、
   同じ返事の後ろの手から "sheet":"集計" で書く。"sheet" を書かなければ今までどおり対象シート。
   書いたシートは全部 --undo で戻せる。ピボット・パワークエリの "sheet" は「置き場」の指定なので、この読み替えとは別。
 - 数式は '=' で始める。文字列はそのまま書く。日付や数値に見える文字（2026/01/05・0001・1,000）を文字のまま書きたいときは先頭に ' を付ける（Excel と同じ。' は残らない）。**表記の統一・掃除（空白・全角半角・カナ・ハイフン・桁・日付）は、行数に関わらず normalize に規則を当てる**（例: 氏名の空白を全角 1 つに＝ {"op":"normalize","range":"B6:B20","rules":["space_zenkaku"],"overwrite":true}。AI が値を書き直さない＝返事が短く速く、写し間違いも無い。元の列に書くので "overwrite": true）。write_cells で書くのは、空欄の補填や 1 か所の誤記のように**規則で表せない個別の値**だけ。式にはしない。write_grid は 100 行まで。
 - 集計の式は**同じ式を右にも下にもコピーできる形**で書く。月別なら見出しにその月の 1 日の日付を置き、
   式は ">="&B$1 と "<"&EDATE(B$1,1) で判定する（列ごとに月の数や年を式に書き換えない。TEXT や MONTH で判定しない）。
   区分別・担当別なら見出しや左端のセル（$A2）を参照する。列ごとに違う式を並べると、人が月や区分を足したとき壊れる。
 - 突き合わせは =IFERROR(VLOOKUP(キー,表,列,FALSE),"〜なし") の形。無いものは空欄でなく「〜なし」の文字。
   ただし金額など数の列に返すときは "" にする（文字が混ざると数値列と見なされず、仕上げ検査に引っかかる）。
   その場合、未一致がどれかは隣に印の列を作るか cond_format の色で示す。
 - **式の点検は材料の「気づき（数式）」を出発点にする**（列の中で形の違う式・式に直書きされた数値・
   集計の起点行の食い違いを、道具が全部の式について数えている）。そこに出ていないものを探すときだけ
   read（"formula": true）で読む。道具が挙げた番地は必ず報告に写す（黙って落とさない）。
 - 材料に無いことを想像で書かない。見えていない範囲は read で見る。**材料の「全体（値）」に表が全部出ているとき（40 行 × 30 列まで）は read しない**（同じ見え方が返るだけで往復を 1 回失う）。空白が全角か半角か・半角カナ・全角英数の有無は材料の「列の中の文字種」に数字で出ている。
 - 行が多い表（材料に「列プロファイル」が出る表）は、列プロファイルで列ごとの型と乱れを見て、normalize / fill / find_replace で直す。結果は道具が件数と例で返すので、それを見て done にする。write_grid は 100 行までで、多い行の値を自分で書かない（道具が断る）。
 - 一度の返事に必要な actions を全部並べる（往復を減らす）。道具は実行前に全部の手を検査し、通らない手が
   1 つでもあれば **1 手も実行しない**（全部の問題を一度に返す＝直して全部を並べ直す）。実行に入ってからは
   上から順で、**1 つ失敗すると残りは実行されない**（結果に「止めました」と出る）ので、**read と安全な手を先に、
   人の表を書き換える手を後に**並べる。結果は次の【結果】で見せる（手と一緒に done を返した回は、全部通れば見せずに検査へ進む）。
 - 表を新しく書いた／列を足した／合計の行を足したときは、足した分を含む表全体（上の「言葉の決め」を見る）に tidy を当ててから done。
 - 既存の見出し・名前定義・既存の数式は壊さない。
 - 新しく作る表の型: 見出しは 1 行・結合なし・空の列や空の行を挟まない・単位は見出しに（「金額（円）」）・1 セル 1 値。
 - 行や列を隠さない（隠す手は無い）。見せたくない明細は row_group／col_group で畳む（人が開ける）。
 - 元から隠れている行・列（材料の「非表示」）と元からの絞り込みは、依頼に無ければそのまま（人がそうしておいた状態）。
   再表示するのは依頼にそう書いてあるときだけ。row_height や format で見せることも「再表示」に当たる。
 - ピボットを作ったら pivot_calc get_data で中身（合計）を読み、元の表と合っているか確かめてから done。
 - 月別・四半期別は pivot_field group_date、構成比・前月比・累計・順位は show_as、上位 N は top_n、並べ替えは sort で（ピボットの代わりに数式で作り直さない）。データモデル・パワークエリの流れは、要る依頼のときだけ道具が材料の末尾に足す。
 - 帳票（結合セットのレイアウト表）を一覧に直すときは、format の "unmerge": true で結合を外し、clear_range（"overwrite": true）で古い中身を消し、write_grid で 1 行 1 件に書き直し、tidy で仕上げる。結合を外した所は上の値で埋める（normalize の fill_down）。
 - 消せるのは、範囲の中身（clear_range）・行と列（row_delete／col_delete）・図形と画像（shape の delete）。どれも "overwrite": true か delete の指定が要り、**承認の言葉があるときだけ**（無ければ消す候補を report に書いて止まる）。行・列を消すと、他のシートからその行を参照している数式は #REF! に化ける＝消す前に材料の「他のシート」と数式を見て、壊れそうなら消さずに report で断る。名前定義とリンクを消す手は無い。
 - **頼まれたことは全部やるか、できない分を必ず report に書くかのどちらか。できる分だけやって黙るのは禁止。**
   返事を作る前に依頼を分解し、plan に並べる（例「目立たせて、日付の新しい順に並べ替えて」＝色分け＋並べ替えの 2 つ）。
   分解した 1 つずつについて、上の手のどれで当てるかを決める。当てる手が無いもの・材料から番地が決められないものは
   「不可」、人に聞かないと決められないものは（安全側で仮に処理してから）「要判断」にする。
   say にも一言入れる（黙って落とすと、頼んだ人は済んだと思ったまま気づかない）。
 - 安全側の仮処理（要判断にするときの型）:
     重複が見つかった → 消さずに cond_format の duplicates で色を付け、件数を report に。
     突き合わせで相手が見つからない → 「〜なし」を残し、その行を印か色で示す。
     数の列に文字が混ざっている → 元の列は触らず、隣に normalize で整形した列を作る。
     見出しが 2 段（結合）→ 結合は外さず、平坦な見出しの表を別の場所に作ってそこを参照する。
     表の途中に空の行がある → 消さずに印を付け、集計は空行を無視する式（SUMIF 系）で書く。
     並べ替えの鍵の列に空欄がある → 並べ替えず、空欄の件数と番地を report に（黙って末尾に回さない）。
 - **仕様が書かれていない依頼（「見やすくして」「整理して」「集計して」「分かるようにして」）は、報告で止めずに
   一つの答えを出す。** 決めきれない所は材料から素直に決めて、report に「こう決めた」と 1 行書く。既定の読み替え:
     見やすく／整えて／体裁を直して → 表全体に tidy（見出し・罫線・列の型・列幅）。値は変えない。
     ○○が分かるように／目立たせて → cond_format で色分け（固定の塗りにしない。並べ替えても付いて回るように）。
     集計して／合計を出して → 表の下に「合計」の行を作り、金額の列を =SUM で（元の表は壊さない）。
       追記が続く一覧なら表と合計の間を 1 行空ける（すぐ下に置くと、次に行を足す人が合計行を潰す）。テーブルなら集計行。
     ○○ごとに何件・いくら → ピボット（別シート）か COUNTIF/SUMIF の小さい表を、元の表の右か下の空いた所に。
     並べ替えて／多い順・新しい順 → sort（**見出し行**は動かさない。材料の「見出し行の推定」を見る）。
   決めようがないとき（表が複数あってどれか選べない・人の値を消してよいか分からない）は、
   止まるのではなく**安全側で仮に処理して plan を「要判断」にし、"ask" に質問を書く**。
 - 終わりに道具が仕上げ検査（見出しの体裁・罫線の継ぎはぎ・数値列の桁区切り・番号列の寄せ・列幅の ###・
   空の見出し）と中身の検査（エラー値・「合計」行が上の和と合わない・数式の列に直値が混ざる・
   **集計の式が本文の行を範囲に入れていない**＝ =SUM(C2:C7) の下に本文の C8 がある形・
   **横に並ぶ同じ関数の集計が列ごとに違う式**＝月や区分を式に埋めた表は列を足すと壊れるので直させる）を実物に当てる。
   引っかかっているうちは done を受け付けない。表を書いたら tidy まで当ててから done。
   表の中の空欄・数値列の文字の数字・番号列の重複・**列の中で式の形が違うセル・式に直書きされた数値**は
   道具が数えて報告の【気づいたこと】に必ず出す
   （依頼に無ければ直さない。空欄に値を作って埋めない。「埋めてよい」と言われて埋めるときは、同じ列の他の行と同じ形
   ＝空白の種類（材料の「列の中の文字種」で全角か半角か分かる）・カナの幅に合わせる。C10 だけ全角空白にして混在させたことがある）。
 - 終わりにもう一度、依頼文と現物（道具が読み直した文字。--image のときは画像も）を照らして自分の仕事を採点させる。満たしていない点を挙げたら、
   その分を直してから done になる（挙げた分は必ず直す。思いつきの追加は挙げない）。
 - 報告だけの依頼（点検・監査・整理の候補出し）は、セルに書かず actions を空にして done にし、番地・件数・案を report に書く。
 - 往復には上限があるので、早めに done にする。**手と一緒に "done": true を返すのが基本**（結果は道具が読み戻して検査し、足りなければ差し戻すので、結果を見てから done にする往復は要らない。手が無ければ actions を空にして done）。
 - 日本の表で気をつけること: 1 行目がタイトルで見出しが 3 行目のことがある（材料の「見出し行の推定」を見る）。
   郵便番号・電話・会員番号の先頭ゼロは文字のまま（' を付ける）。全角の数字・記号は normalize の hankaku で半角に。
   和暦（令和6年3月1日・R6.3.1）は normalize の date が日付に直す。数式の中に全角の , ( ) " を書かない。
"""


# ----------------------------------------------------------------
# 純 Python 部分（Excel 不要）: 依頼文の組み立て・返事の読み取り・手の変換
# ----------------------------------------------------------------

def _first_prompt(request, book, sheet, materials, rules=None):
    return (_rules_for(request, materials, rules or RULES, sheet) + "\n【依頼】\n" + request.strip()
            + f"\n\n【材料】（ブック {book} のシート「{sheet}」。道具が読んだ現物）\n" + materials.rstrip() + "\n"
            # 返事を書く直前に置く 1 行（規則文の中に書いても sonnet は「現物を確認してから」と read に 1 往復使った・2026-09-08）
            + "\n【この返事で】材料の「全体（値）」に表が全部出ているなら、read で確かめ直さず（同じ格子が返るだけ）、"
              "手を全部並べて \"done\": true を返す。\n")


# 重い道具（テーブル・ピボット・スライサー・パワークエリ・データモデル）の長い段は、要る依頼のときだけ送る
# （2026-09-06 夜: 規則文 20,000 字のうちこの段が 4 割で、今日の依頼で使った手は 6 本以下だった。毎往復
# 履歴に残って入力になる）。手の名前は残す（「手が無い」と誤答させない）。
_HEAVY_BLOCK_HEAD = " --- 重い道具（"
_ROWS_BLOCK_HEAD = " --- 行が多い表の手（"
_LEGS_BLOCK_HEAD = " --- 表の足回り（"
_ROWS_STUB = (" --- 行が多い表の手 normalize / fill / find_replace の詳しい説明は、材料に列プロファイルが出る表"
              "（40 行超）のときだけ出す。**掃除（空白・全角半角・カナ・ハイフン・桁・日付）は小さい表でも normalize**＝ "
              '{"op":"normalize","range":"B6:B20","rules":["space_zenkaku"],"overwrite":true}（rules: trim | hankaku | hyphen | '
              "number | date | kana_zenkaku | space_zenkaku | space_hankaku | fill_down。元の列に書くので overwrite）。"
              "表を書く・1〜2 セルを直すのは write_grid / write_cells ---")
_LEGS_STUB = (" --- 表の足回り sort / autofilter / cond_format / validation / row_insert / col_insert / "
              "row_delete / col_delete / copy_range / view / page_setup / chart / hyperlink / comment / "
              "sheet_op / name_add の説明は、依頼にその話があるときだけ出す（要るなら op の名前を書けば"
              "道具が引数を教えます） ---")
_LEGS_RULE_WORDS = re.compile(r'並べ替え|並び替え|ソート|sort|絞り込|フィルタ|抽出|色分け|条件付き書式|'
                              r'入力規則|プルダウン|リスト選択|挿入|行を足|列を足|行を削|列を削|複写|コピー|'
                              r'貼り付け|印刷|改ページ|枠固定|固定|ハイパーリンク|リンク|コメント|メモ|'
                              r'グラフ|チャート|シートを|名前定義|名前を付け', re.I)
_HEAVY_BLOCK_END = "\n言葉の決め:"
_HEAVY_RULE_WORDS = re.compile(r'ピボット|pivot|テーブル|table|スライサー|slicer|パワークエリ|power ?query|クエリ|'
                               r'データモデル|リレーション|メジャー|DAX|構成比|前月比|累計|上位|順位|集計', re.I)
_HEAVY_MAT_RE = re.compile(r'テーブル（ブック全体）: (\d+) 本|ピボット（ブック全体）: (\d+) 本|スライサー: (\d+) 個|'
                           r'パワークエリ: (\d+) 本|データモデル: テーブル (\d+) 本')
_HEAVY_STUB = (" --- 重い道具 table / pivot / pivot_field / pivot_calc / slicer / powerquery / datamodel の説明は、"
               "依頼か材料にテーブル・ピボット・スライサー・パワークエリ・データモデルがあるときだけ出す"
               "（この依頼には無い＝これらの手は使わない） ---")


def _fold_block(base, head, end, stub):
    """規則文の 1 段を 1 行の見出しに畳む（純 Python）。見つからなければそのまま返す。"""
    i = base.find(head)
    j = base.find(end, i + 1) if i >= 0 else -1
    if i < 0 or j < 0:
        return base
    return base[:i] + stub + base[j:]


def _heavy_needed(request, materials, sheet=None):
    """重い道具（テーブル・ピボット・スライサー・PQ・データモデル）の説明が要るか（純 Python）。

    2026-09-07 未明: 前は「ブックのどこかに 1 本でもあれば要る」だった。実射は 15 弾を 1 冊で回すので、
    3 弾目からは毎回まるごと送っていた（1 往復あたり 3,000 字）。**そのシートに在るか**で見る。
    """
    if _HEAVY_RULE_WORDS.search(request or ''):
        return True
    for line in (materials or '').split(chr(10)):
        t = line.strip()
        if t.startswith('['):
            if sheet is None or t.startswith("[" + str(sheet) + "]"):
                return True          # このシートのテーブル・ピボット
            continue
        if t.startswith(('スライサー:', 'パワークエリ:', 'データモデル:')):
            m = _HEAVY_MAT_RE.search(t)
            if m and any(int(g) > 0 for g in m.groups() if g):
                return True          # シートに紐づかない道具（このブックで使っている＝説明が要る）
    return False


def _legs_needed(request, materials):
    """表の足回り（並べ替え・絞り込み・色分け・入力規則・挿入・複写・印刷・グラフ）の説明が要るか。"""
    if _LEGS_RULE_WORDS.search(request or ''):
        return True
    mat = materials or ''
    if '条件付き書式: 0 本' not in mat or '入力規則: 0 セル' not in mat:
        return True                  # 既にあるものは触る可能性がある（材料に無いときも安全側で送る）
    return False


def _rules_for(request, materials, rules=None, sheet=None):
    """依頼と材料に合わせて規則文を絞る（純 Python）。要らない段は 1 行の見出しに畳む。

    費用の 9 割は入力で、その大半が規則文（16,333 字）。畳んでも手の名前は見出しに残すので、
    AI が「その手は無い」と誤答することはない（要るなら op を書けば道具が引数を教える）。
    位置は全部**元の文の上で先に測る**（順に置き換えると、前の畳みが次の段の終わり目印を消す）。
    """
    base = rules or RULES
    i_rows = base.find(_ROWS_BLOCK_HEAD)
    i_legs = base.find(_LEGS_BLOCK_HEAD)
    i_heavy = base.find(_HEAVY_BLOCK_HEAD)
    i_end = base.find(_HEAVY_BLOCK_END, i_heavy) if i_heavy >= 0 else -1
    if min(i_rows, i_legs, i_heavy, i_end) < 0 or not (i_rows < i_legs < i_heavy < i_end):
        return base                  # 段の並びが変わったら何も畳まない（黙って壊さない）
    rows_txt = base[i_rows:i_legs] if '列プロファイル' in (materials or '') else _ROWS_STUB + chr(10)
    legs_txt = base[i_legs:i_heavy] if _legs_needed(request, materials) else _LEGS_STUB + chr(10)
    heavy_txt = base[i_heavy:i_end] if _heavy_needed(request, materials, sheet) else _HEAVY_STUB
    return base[:i_rows] + rows_txt + legs_txt + heavy_txt + base[i_end:]


def _continue_prompt(supplement, materials, last_results=None):
    """--continue の 1 通目。規則と前回の会話は history に入っているので、ここは追加の指示と今の材料だけ。"""
    parts = []
    if last_results:
        parts.append(_result_prompt(last_results, None).rstrip())
    parts.append("【追加の指示】（前の報告を受けて人が答えました。前に自分で挙げた候補・番地のとおりに進めてください）\n"
                 + (supplement or '').strip())
    parts.append("【いまの材料】（前の往復のあと、道具が読み直した現物。ここを見てから手を返す）\n" + (materials or '').rstrip())
    return "\n\n".join(parts) + "\n"


def _pid_alive(pid):
    """その PID のプロセスが今も生きているか（Windows）。分からないときは False（＝古い錠と見なす）。"""
    try:
        pid = int(pid)
    except Exception:
        return False
    if pid == os.getpid():
        return True
    try:
        import ctypes
        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
        if h:
            ctypes.windll.kernel32.CloseHandle(h)
            return True
    except Exception:
        pass
    return False


@contextlib.contextmanager
def _agent_lock():
    """記録（_last_agent_log.jsonl）を 2 つの agent が同時に書かないための錠（2026-09-04）。

    MCP と CLI を同時に回すと往復が混ざり、--continue が別の仕事の会話を読み戻す。
    錠は PID つきで、書いたプロセスが死んでいれば無視する（異常終了で残っても次は通る）。
    """
    try:
        with open(_AGENT_LOCK_FILE, encoding='utf-8') as f:
            other = json.load(f)
    except Exception:
        other = None
    if not isinstance(other, dict):
        other = None            # 壊れた錠（途中まで書かれた・JSON だが辞書でない）は無視する。
        #                         前は other.get で AttributeError になり、agent が全部起動できなくなっていた（2026-09-04）
    # 同じ PID の別スレッドも「別の agent」（MCP はタイムアウトの後、同じプロセスの新しいスレッドで次を始める。
    # 前の世代がまだ Excel を触っていると 2 つの agent が同じブックに入っていた・2026-09-04）
    same_thread = (int(other.get('pid') or 0) == os.getpid()
                   and other.get('thread') == threading.get_ident()) if other else False
    if other and _pid_alive(other.get('pid')) and not same_thread:
        raise RuntimeError(
            f"別の agent が走っています（PID {other.get('pid')}・{other.get('at_text') or '時刻不明'}に開始）。"
            "記録（_last_agent_log.jsonl）が混ざって --continue が壊れるので、終わるのを待ってください。"
            f"（止まったままなら {_AGENT_LOCK_FILE} を消す）")
    try:
        with open(_AGENT_LOCK_FILE, 'w', encoding='utf-8') as f:
            json.dump({'pid': os.getpid(), 'thread': threading.get_ident(), 'at': time.time(),
                       'at_text': time.strftime('%H:%M:%S')}, f, ensure_ascii=False)
    except Exception:
        pass
    try:
        yield
    finally:
        try:
            os.remove(_AGENT_LOCK_FILE)
        except Exception:
            pass


def _load_resume(path=None):
    """記録（_last_agent_log.jsonl）から前回の会話を読み戻す → (history, meta, 最後の結果)。無ければ (None, None, None)。"""
    path = path or _LAST_AGENT_LOG_FILE
    try:
        with open(path, encoding='utf-8') as f:
            lines = [json.loads(l) for l in f if l.strip()]
    except (OSError, ValueError):
        return None, None, None
    if not lines:
        return None, None, None
    meta = lines[0].get('meta') if isinstance(lines[0], dict) else None
    turns = [d for d in lines if isinstance(d, dict) and 'reply' in d]
    if not turns:
        return None, meta, None
    history = []
    for d in turns:
        history.append(('user', d.get('prompt') or ''))
        history.append(('model', d.get('reply') or ''))
    last = turns[-1].get('results') or None
    last_results = [(r.get('label', '?'), bool(r.get('ok')), r.get('out', '')) for r in last] if last else None
    return history, meta, last_results


def _result_prompt(results, turns_left):
    """turns_left に None を渡すと「残りの往復」を書かない（--continue の 1 通目＝前回の結果を見せるだけ）。"""
    parts = ["【前回の結果】（前の往復で道具が実行した結果）" if turns_left is None else
             "【結果】（あなたの actions を道具が実行した結果。読んで、続きの actions か done を返す）"]
    for i, (label, ok, text) in enumerate(results, 1):
        head = f"--- {i}: {label} {'' if ok else '（失敗）'}---"
        body = text.rstrip()
        limit = _RESULT_LIMIT_BIG if str(label).startswith(_RESULT_BIG_LABELS) else _RESULT_LIMIT
        if len(body) > limit:
            body = body[:limit] + f"\n…（{len(text) - limit:,} 字を省略。範囲を分けて読む）"
        parts.append(head + "\n" + body)
    if turns_left is not None:
        parts.append(f"（残りの往復: {turns_left} 回）")
    return "\n".join(parts) + "\n"


# 要判断（2026-09-06・Claude for Excel の ask_user_question を、無人の往復に合う形へ移した）。
# 向こうは分岐が出ると選択肢を出して**止まる**。こちらは往復の途中で人に聞く口が無いので、
# 「安全側で仮処理して、report に質問としてまとめる」に読み替えた。人は agent --cont で答える。
# 済・不可と同じく「未ではない」＝done を塞がない。塞ぐと無人のループが必ず往復を使い切る。
_PLAN_STATES = ('未', '済', '不可', '要判断')


def _parse_plan(v):
    """返事の plan → [{'item': 文, 'state': 未|済|不可|要判断, 'ask': 人への質問}]。読めない要素は落とす。"""
    out = []
    for e in (v or []):
        if isinstance(e, dict) and str(e.get('item') or '').strip():
            st = str(e.get('state') or '未').strip()
            p = {'item': str(e['item']).strip(), 'state': st if st in _PLAN_STATES else '未'}
            if str(e.get('ask') or '').strip():
                p['ask'] = str(e['ask']).strip()
            if str(e.get('why') or '').strip():        # 途中で足した項目の理由（2026-09-06 夜・discovered）
                p['why'] = str(e['why']).strip()
            out.append(p)
    return out


_PLAN_ADDR_RE = re.compile(r'\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\d{1,7}\s*行目?|[A-Za-z]{1,3}\s*列')
_PLAN_STRIP_RE = re.compile(r'[\s　「」『』（）()\[\]【】・、。，．,.:：;；!！?？\-—–~〜/／\'"“”‘’*＊]+')


def _plan_norm(item):
    return _PLAN_STRIP_RE.sub('', str(item or '')).lower()


def _plan_addrs(item):
    return {re.sub(r'[\s$　]', '', m).upper() for m in _PLAN_ADDR_RE.findall(str(item or ''))}


def _plan_same(a, b):
    """項目の文が同じことを指しているか（純 Python）。AI は往復ごとに言い換えるので、完全一致では足りない。

    2026-09-06 夕（シュウさんの初実射）: 採点係が足した「A17セルの会員番号が全角数字「１０１２」のままなので
    半角「1012」に修正してください」を、AI が「A17の会員番号「１０１２」を半角「1012」に修正」と言い換えて
    要判断にした。文字が違うので別の項目として足され、元の項目は「未」のまま残って done を差し戻し、最後は
    同じ項目が不可と要判断で二重に出た（一覧 6 件・中身 3 件・往復 5 回）。
    見方: 記号と空白を落として (1) 同じ (2) 片方が片方を含む (3) 短いほうの 2 文字組の 7 割が長いほうにある。
    番地が両方にあって食い違う（A17 と A18・B 列と C 列）ときは別の項目。
    限界: 「氏名の空白を…」と「フリガナの空白を…」のように列の名前だけ違う 2 項目は同じに見える。
    _merge_plan は先に完全一致で組にしてから残りをこれで組むので、AI が片方だけ言い換えたときは正しく組める。
    """
    na, nb = _plan_norm(a), _plan_norm(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    aa, ab = _plan_addrs(a), _plan_addrs(b)
    if aa and ab and not (aa & ab):
        return False
    if min(len(na), len(nb)) >= 4 and (na in nb or nb in na):
        return True
    short, long_ = (na, nb) if len(na) <= len(nb) else (nb, na)
    grams = {short[i:i + 2] for i in range(len(short) - 1)}
    if len(grams) >= 4:
        hit = sum(1 for g in grams if g in long_)
        if hit / len(grams) >= 0.7:
            return True
    return False


def _merge_plan(old, new):
    """前の一覧を残したまま、返ってきた分の state だけ更新する（項目は消させない）。

    前は `plan = plan_now` で丸ごと差し替えていた＝AI が 2 往復目に項目を書き落とすと、
    その項目は「未」ごと消えて落とし物の検査をすり抜けた（2026-09-04）。
    新しい項目は足す（依頼の分解が細かくなるのはよい）。項目は文で照合する＝言い換えも同じ項目と見る
    （_plan_same・2026-09-06 夕）。前の文を残し、state・ask・why だけ新しいほうから取る。
    """
    if not old:
        return list(new or [])
    rest = list(new or [])
    pair = {}
    for p in old:                                   # 1 周目: 完全一致（記号を落とした文）で組む
        got = next((q for q in rest if _plan_norm(p['item']) == _plan_norm(q['item'])), None)
        if got:
            rest.remove(got)
            pair[id(p)] = got
    for p in old:                                   # 2 周目: 残りを言い換えとして組む
        if id(p) in pair:
            continue
        got = next((q for q in rest if _plan_same(p['item'], q['item'])), None)
        if got:
            rest.remove(got)
            pair[id(p)] = got
    out = []
    for p in old:
        keep = dict(p)
        got = pair.get(id(p))
        if got:
            keep['state'] = got['state']
            if got.get('ask'):
                keep['ask'] = got['ask']
            if got.get('why') and not keep.get('why'):
                keep['why'] = got['why']
        out.append(keep)
    out.extend(rest)
    return out


def _plan_added(old, new):
    """返ってきた plan のうち、前の一覧に無かった項目（途中で足した分・純 Python）。言い換えは足した分に数えない。"""
    return [p for p in (new or []) if not any(_plan_same(p['item'], q['item']) for q in (old or []))]


def _plan_pending(plan):
    return [p['item'] for p in (plan or []) if p['state'] == '未']


def _plan_mark_done(plan, items):
    """一覧の「未」を道具の側で「済」にする（2026-09-08）。

    手と done を同じ返事で受け、手が全部通った回だけに使う。その返事の一覧は手を撃つ前に
    書かれているので「未」は状態が古いだけ。ここで直しておかないと、報告の一覧に
    「未」が残ったまま合格になる（数だけ見て「落とし物がある」と読めてしまう）。
    """
    want = set(items or [])
    for p in (plan or []):
        if p.get('state') == '未' and p.get('item') in want:
            p['state'] = '済'


def _plan_asks(plan):
    """要判断のまま終わった項目 → 「項目（質問）」の一覧。report と報告に必ず出す（2026-09-06）。"""
    out = []
    for p in (plan or []):
        if p.get('state') == '要判断':
            out.append(p['item'] + (f"　→ 聞きたいこと: {p['ask']}" if p.get('ask') else ""))
    return out


def _plan_line(plan):
    if not plan:
        return ""
    n = {s: sum(1 for p in plan if p['state'] == s) for s in _PLAN_STATES}
    head = (f"やることの一覧（{len(plan)} 件）: 済 {n['済']} / 不可 {n['不可']}"
            f" / 要判断 {n['要判断']} / 未 {n['未']}")
    return head + "\n" + "\n".join(
        f"  [{p['state']}] {p['item']}" + (f"　→ {p['ask']}" if p.get('ask') else "") for p in plan)


def _parse_reply(text):
    """返事 → (say, actions, done, report, plan)。形が違えば ValueError（理由つき）。"""
    try:
        d = _extract_json(text)
    except Exception as ex:
        raise ValueError(f"JSON として読めません: {ex}")
    if not isinstance(d, dict):
        raise ValueError("JSON の最上位がオブジェクトではありません")
    actions = d.get('actions') or []
    if not isinstance(actions, list):
        raise ValueError("actions が配列ではありません")
    done = bool(d.get('done')) and not actions      # done は actions が空のときだけ
    return (str(d.get('say') or ''), actions, done, str(d.get('report') or ''), _parse_plan(d.get('plan')))


def _wanted_done(text):
    """返事が done を立てていたか（actions と同時でも拾う）。読めなければ False。

    done は actions が空のときだけ有効（_parse_reply）だが、最後の往復で「手も出すし done」と
    返されると、手は実行済みなのに「AI は done を言っていません」で失敗になっていた（2026-09-04）。
    """
    try:
        d = _extract_json(text)
        return bool(d.get('done')) if isinstance(d, dict) else False
    except Exception:
        return False








def _cell_text(v):
    """セルに書く値 → 文字列。セル内改行はそのまま残す（潰すのは TSV に並べるときだけ・2026-09-04）。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _tsv_text(v):
    """TSV の 1 マス（タブ・改行は行と列の区切りなので空白に潰す）。"""
    return _cell_text(v).replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')


# ----------------------------------------------------------------
# 重い道具の手（2026-09-04）: CLI の table / pivot / pivot-field / pivot-calc / slicer / powerquery / datamodel を
#   引数列に写すだけ（実行は _run_cmd＝同じプロセス・同じ COM 接続・対象ブック固定）。
#   消す手（delete / column remove / measure delete）は写さない＝道具に無い手のまま。
# ----------------------------------------------------------------
_HEAVY_ACTIONS = {
    'table': ('create', 'list', 'read', 'sort', 'filter', 'filter_values', 'filter_clear', 'column_add'),
    'pivot': ('create', 'list'),
    'pivot_field': ('list', 'add_row', 'add_col', 'add_filter', 'add_value', 'remove', 'set_func', 'set_name',
                    'set_format', 'set_filter', 'sort', 'group_date', 'group_numeric',
                    'show_as', 'top_n', 'filter_clear', 'position'),
    'pivot_calc': ('get_data', 'calc_field_create', 'calc_field_list', 'layout', 'subtotals', 'grand_totals',
                   'refresh', 'style', 'repeat_labels', 'empty_as', 'set_source'),
    'slicer': ('add', 'list'),
    'powerquery': ('list', 'refresh', 'add', 'edit', 'load'),
    'datamodel': ('list', 'relation_add', 'measure_add'),
}
_PIVOT_FUNCS = ('sum', 'count', 'average', 'max', 'min')
_SHOW_AS_KINDS = ('none', 'percent_total', 'percent_row', 'percent_col', 'percent_parent_row', 'percent_parent_col',
                  'percent_of', 'diff_from', 'percent_diff_from', 'running_total', 'percent_running_total',
                  'rank_asc', 'rank_desc', 'index')      # vbam_heavy._XL_SHOW_AS と同じ並び（テストで照合）
_MEASURE_FORMATS = ('general', 'whole', 'decimal', 'currency', 'percent', 'scientific')


def _sheet_range(sheet, rng):
    """'シート名'!A1:D10 の形（' は '' に。_resolve_range が剥がす）"""
    return "'" + str(sheet).replace("'", "''") + "'!" + rng


def _need(act, key, what=None):
    v = act.get(key)
    if v in (None, ''):
        raise ValueError(f"{act.get('op')} {act.get('action')} には {key}（{what or key}）が要ります")
    return str(v)


def _opt(toks, act, key, opt=None):
    v = act.get(key)
    # 0 は値（0 == False なので `not in (None, '', False)` だと "decimals": 0 を黙って落としていた・2026-09-10）
    if v is not None and v != '' and v is not False:
        toks += [opt or ('--' + key.replace('_', '-')), str(v)]


def _values_of(act, key='values'):
    vals = act.get(key)
    if not isinstance(vals, list) or not vals:
        raise ValueError(f"{act.get('op')} {act.get('action')} には {key}（配列）が要ります")
    return [str(v) for v in vals]


def _fields_csv(act, key):
    v = act.get(key)
    if isinstance(v, list):
        v = ",".join(str(x) for x in v)
    return '' if v in (None, '') else str(v)


def _pivot_func(act, key='func', default=None):
    f = act.get(key)
    if f in (None, ''):
        return default
    f = str(f).lower()
    if f not in _PIVOT_FUNCS:
        raise ValueError(f"func は {'/'.join(_PIVOT_FUNCS)}（{f!r}）")
    return f


def _heavy_action_to_tokens(act, sheet):
    """重い道具の手 1 つ → (表示名, 引数列)。action を _HEAVY_ACTIONS で絞る（消す手は通さない）。"""
    op = str(act.get('op') or '')
    action = str(act.get('action') or '').strip().lower().replace('-', '_')
    allowed = _HEAVY_ACTIONS[op]
    if action not in allowed:
        raise ValueError(f"{op} の action は {'/'.join(allowed)}（{action!r}）。消す手は無い")
    label = f"{op} {action}"
    if op == 'table':
        if action == 'create':
            toks = ['table', 'create', _sheet_range(sheet, _check_addr(act.get('range')))]
            if act.get('name'):
                toks.append(str(act['name']))
            if act.get('no_headers'):
                toks.append('--no-headers')
            return label, toks
        if action == 'list':
            return label, ['table', 'list']
        name = _need(act, 'name', 'テーブル名')
        label += f" {name}"
        if action == 'read':
            return label, ['table', 'read', name]
        if action == 'sort':
            toks = ['table', 'sort', name, _need(act, 'column', '列名')]
            if act.get('desc'):
                toks.append('--desc')
            return label, toks
        if action == 'filter':
            return label, ['table', 'filter', name, _need(act, 'column', '列名'), _need(act, 'criteria', '条件 ">100" など')]
        if action == 'filter_values':
            return label, ['table', 'filter-values', name, _need(act, 'column', '列名')] + _values_of(act)
        if action == 'filter_clear':
            return label, ['table', 'filter-clear', name]
        toks = ['table', 'column', 'add', name, _need(act, 'column', '足す列名')]
        _opt(toks, act, 'at')
        return label, toks
    if op == 'pivot':
        if action == 'list':
            return label, ['pivot', 'list']
        is_model = bool(act.get('model'))
        if is_model:
            toks = ['pivot', 'create', 'model', '--model']
        else:
            toks = ['pivot', 'create', _sheet_range(sheet, _check_addr(act.get('range')))]
        rows, cols, values = _fields_csv(act, 'rows'), _fields_csv(act, 'cols'), _fields_csv(act, 'values')
        filters, measures = _fields_csv(act, 'filter'), _fields_csv(act, 'measures')
        if not rows and not cols:
            raise ValueError("pivot create には rows か cols（見出し名）が要ります")
        if not values and not measures:
            raise ValueError("pivot create には values（集計する列の見出し名）か、データモデルなら measures（メジャー名）が要ります")
        if measures and not is_model:
            raise ValueError('measures はデータモデルのピボット（"model": true）だけ')
        if rows:
            toks += ['--rows', rows]
        if cols:
            toks += ['--cols', cols]
        if filters:
            toks += ['--filter', filters]
        if values:
            toks += ['--values', values]
        if measures:
            toks += ['--measures', measures]
        f = _pivot_func(act)
        if f:
            toks += ['--func', f]
        # sheet が先。AI は「別シートへ」と言いながら at も添えてくる（at だけを先にすると
        # 元データのシートの A3 ＝ 表の真上に置いて上書きの窓で止まる・2026-09-04 実射）。
        # 両方あれば「そのシートのそのセル」（CLI が既存データの上を断る）
        if act.get('sheet'):
            toks += ['--sheet', str(act['sheet'])]
        if act.get('at'):
            toks += ['--at', _check_addr(act['at'])]
        _opt(toks, act, 'name')
        return label, toks
    if op == 'pivot_field':
        pv = _need(act, 'pivot', 'ピボット名')
        label += f" {pv}"
        toks = ['pivot-field', action.replace('_', '-'), pv]
        if action == 'list':
            return label, toks
        if action in ('group_date', 'top_n', 'sort', 'filter_clear', 'position') and not act.get('field'):
            toks.append('*')            # 道具側で行フィールドから当てる（当てられなければ道具が言う・2026-09-06 深夜）
        else:
            toks.append(_need(act, 'field', 'フィールド名'))
        if action == 'add_value':
            f = _pivot_func(act)
            if f:
                toks += ['--func', f]
            _opt(toks, act, 'name')
        elif action == 'set_func':
            toks.append(_pivot_func(act) or _need(act, 'func', '/'.join(_PIVOT_FUNCS)))
        elif action == 'set_name':
            toks.append(_need(act, 'name', '新しい表示名'))
        elif action == 'set_format':
            toks.append(_need(act, 'format', '書式コード "#,##0" など'))
        elif action == 'set_filter':
            toks += _values_of(act)
        elif action == 'sort':
            order = str(act.get('order') or 'asc').lower()
            if order not in ('asc', 'desc'):
                raise ValueError("sort の order は asc|desc")
            toks.append(order)
            _opt(toks, act, 'by')                      # 値の順（"by":"sum/売上"）。無ければラベル順
        elif action == 'group_date':
            by = str(act.get('by') or '').lower()
            if by not in ('days', 'months', 'quarters', 'years'):
                raise ValueError("group_date の by は days|months|quarters|years")
            toks.append(by)
        elif action == 'group_numeric':
            toks += [_need(act, 'start'), _need(act, 'end'), _need(act, 'interval')]
        elif action == 'show_as':
            kind = str(act.get('kind') or act.get('as') or '').lower().replace('-', '_')
            if kind not in _SHOW_AS_KINDS:
                raise ValueError("show_as の kind は " + "|".join(_SHOW_AS_KINDS) + f"（{kind!r}）")
            toks.append(kind)
            _opt(toks, act, 'base_field')
            _opt(toks, act, 'base_item')
        elif action == 'top_n':
            try:
                n = int(act.get('n') or 0)
            except (TypeError, ValueError):
                n = 0
            if n <= 0:
                raise ValueError("top_n には n（1 以上の整数）が要ります")
            toks.append(str(n))
            _opt(toks, act, 'by')
            if act.get('bottom'):
                toks.append('--bottom')
        elif action == 'position':
            try:
                n = int(act.get('n') or act.get('position') or 0)
            except (TypeError, ValueError):
                n = 0
            if n <= 0:
                raise ValueError("position には n（1 以上の整数）が要ります")
            toks.append(str(n))
        return label, toks
    if op == 'pivot_calc':
        pv = _need(act, 'pivot', 'ピボット名')
        label += f" {pv}"
        if action.startswith('calc_field_'):
            sub = action[len('calc_field_'):]
            toks = ['pivot-calc', 'calc-field', sub, pv]
            if sub == 'create':
                toks += [_need(act, 'name', '計算フィールド名'), _need(act, 'formula', '数式 "=金額*1.1" など')]
            return label, toks
        toks = ['pivot-calc', action.replace('_', '-'), pv]
        if action == 'layout':
            lay = str(act.get('layout') or 'compact').lower()
            if lay not in ('compact', 'tabular', 'outline'):
                raise ValueError("layout は compact|tabular|outline")
            toks.append(lay)
        elif action == 'subtotals':
            toks += [_need(act, 'field', 'フィールド名'), 'on' if act.get('on', True) else 'off']
        elif action == 'grand_totals':
            which = str(act.get('which') or 'both').lower()
            if which not in ('rows', 'cols', 'both'):
                raise ValueError("grand_totals の which は rows|cols|both")
            toks += [which, 'on' if act.get('on', True) else 'off']
        elif action == 'style':
            toks.append(_need(act, 'style', 'スタイル名 PivotStyleMedium9 など'))
        elif action == 'repeat_labels':
            toks.append('on' if act.get('on', True) else 'off')
        elif action == 'empty_as':
            toks.append(str(act.get('text', '') if act.get('text') is not None else ''))
        elif action == 'set_source':
            toks.append(_sheet_range(sheet, _check_addr(act.get('range'))))
        return label, toks
    if op == 'slicer':
        if action == 'list':
            return label, ['slicer', 'list']
        toks = ['slicer', 'add', _need(act, 'source', 'ピボット名かテーブル名'), _need(act, 'field', 'フィールド名')]
        if act.get('at'):
            toks += ['--at', _check_addr(act['at'])]
        _opt(toks, act, 'name')
        return label, toks
    if op == 'powerquery':
        if action == 'list':
            return label, ['powerquery', 'list']
        if action == 'refresh':
            return label, ['powerquery', 'refresh'] + ([str(act['name'])] if act.get('name') else [])
        name = _need(act, 'name', 'クエリ名')
        label += f" {name}"
        if action in ('add', 'edit'):
            toks = ['powerquery', action, name, '--m', _need(act, 'm', 'M 式（let … in …）')]
            _opt(toks, act, 'desc')
            return label, toks
        to = str(act.get('to') or 'sheet').lower()
        if to not in ('sheet', 'model'):
            raise ValueError("load の to は sheet|model")
        toks = ['powerquery', 'load', name, '--to', to]
        if to == 'sheet':
            toks += ['--sheet', str(act.get('sheet') or sheet)]
            if act.get('at'):
                toks += ['--at', _check_addr(act['at'])]
        return label, toks
    # datamodel
    if action == 'list':
        return label, ['datamodel', 'list']
    if action == 'relation_add':
        return label, ['datamodel', 'relation', 'add', _need(act, 'fk_table', '多側のテーブル'), _need(act, 'fk_column'),
                       _need(act, 'pk_table', '一側のテーブル'), _need(act, 'pk_column')]
    toks = ['datamodel', 'measure', 'add', _need(act, 'table', 'テーブル名'), _need(act, 'name', 'メジャー名'),
            '--dax', _need(act, 'dax', 'DAX 式')]
    if act.get('format'):
        fmt = str(act['format']).lower()
        if fmt not in _MEASURE_FORMATS:
            raise ValueError(f"format は {'|'.join(_MEASURE_FORMATS)}")
        toks += ['--format', fmt]
    _opt(toks, act, 'decimals')
    if act.get('thousands'):
        toks.append('--thousands')
    _opt(toks, act, 'symbol')
    return label + f" {act['name']}", toks


_CF_KEYS = {'gt': '--gt', 'lt': '--lt', 'ge': '--ge', 'le': '--le', 'eq': '--eq', 'ne': '--ne'}
_CHART_CFG_ACTIONS = ('legend', 'set-title', 'set-axis-title', 'axis-scale', 'axis-format',
                      'gridlines', 'style', 'set-type', 'set-source', 'placement', 'data-labels')
_SHEET_OP_ACTIONS = ('add', 'rename', 'copy', 'activate', 'tab-color')


def _table_action_to_tokens(act, sheet):
    """表の足回りの手 1 つ → (表示名, 引数列)。消す向き（delete/clear/remove/off）は通さない。"""
    op = str(act.get('op') or '')
    if op == 'sort':
        rng = _check_addr(act.get('range'))
        toks = ['sort', '--sheet', sheet, rng]
        key = str(act.get('key') or '').strip()
        if key:
            if not (key.isascii() and key.isalpha()):
                raise ValueError(f"sort の key は列文字（A〜XFD）です（{key!r}）")
            toks += ['--key', key.upper()]
        if act.get('desc'):
            toks.append('--desc')
        toks.append('--no-header' if act.get('header') is False else '--header')
        return f"sort {rng}" + (f" key={key.upper()}" if key else "") + ("（降順）" if act.get('desc') else "（昇順）"), toks
    if op == 'autofilter':
        # autofilter / cond-format / validation に --sheet は無い。対象は範囲に 'シート名'!A1 の形で渡す
        # （2026-09-04 の実射: --sheet を付けて 3 本とも「引数エラー」で落ち、AI が report に「道具の制約」と書いた）
        if act.get('off'):
            return "autofilter 解除", ['autofilter', _sheet_range(sheet, 'A1'), '--off']
        rng = _check_addr(act.get('range'))
        if act.get('show_all'):
            return "autofilter 絞り込み解除", ['autofilter', _sheet_range(sheet, rng), '--show-all']
        col, eq = act.get('column'), act.get('equals')
        if col is not None or eq is not None:
            # 「◯◯だけ抜き出して」＝条件つきの絞り込み（既にフィルタが付いていても掛け直す）
            if col is None or eq is None:
                raise ValueError("autofilter の絞り込みは column（見出し）と equals（残す値）を両方書きます")
            vals = eq if isinstance(eq, (list, tuple)) else [eq]
            toks = ['autofilter', _sheet_range(sheet, rng), '--column', str(col)]
            for v in vals:
                toks += ['--equals', _cell_text(v)]
            return f"autofilter {rng} {col}={'/'.join(map(str, vals))}", toks
        return f"autofilter {rng}", ['autofilter', _sheet_range(sheet, rng)]
    if op == 'cond_format':
        if act.get('clear'):
            raise ValueError("条件付き書式を消す手は無い（消す候補は report に書く）")
        rng = _check_addr(act.get('range'))
        toks = ['cond-format', _sheet_range(sheet, rng)]
        n = 0
        if act.get('duplicates'):
            # 「同じ人・同じ行が二重に入っている」を色で示す（Excel の「重複する値」の規則）
            toks.append('--duplicates')
            n += 1
        for k, opt in _CF_KEYS.items():
            if act.get(k) not in (None, ''):
                toks += [opt, _cell_text(act[k])]
                n += 1
        if act.get('between'):
            bt = act['between']
            if not isinstance(bt, (list, tuple)) or len(bt) != 2:
                raise ValueError("cond_format の between は [下限, 上限] の配列です")
            toks += ['--between', _cell_text(bt[0]), _cell_text(bt[1])]
            n += 1
        if act.get('formula'):
            f = str(act['formula'])
            if not f.startswith('='):
                raise ValueError("cond_format の formula は '=' で始めます（範囲の左上のセルを基準に書く）")
            toks += ['--formula', f]
            n += 1
        if n == 0:
            raise ValueError("cond_format に条件がありません（gt/lt/ge/le/eq/ne/between/formula のどれか）")
        if n > 1:
            raise ValueError("cond_format の条件は 1 つだけ（複数当てるなら手を分ける）")
        m = 0
        for k, opt in (('bg', '--bg'), ('color', '--color')):
            if act.get(k):
                toks += [opt, str(act[k])]
                m += 1
        if act.get('bold'):
            toks.append('--bold')
            m += 1
        if m == 0 and not act.get('duplicates'):
            raise ValueError("cond_format に当てる見た目がありません（bg／color／bold）")
        return f"cond_format {rng}", toks
    if op == 'validation':
        if act.get('clear'):
            raise ValueError("入力規則を消す手は無い（消す候補は report に書く）")
        rng = _check_addr(act.get('range'))
        # 候補の名前は list が正。values・items・options・choices も受ける（2026-09-11: 3 回とも 1 往復を捨てた）
        items = next((act.get(k) for k in ('list', 'source', 'value', 'values', 'items', 'options', 'choices', 'formula1')
                      if act.get(k)), None)      # AI は Excel の用語どおり source と書く（5 回とも・2026-09-11）
        if not items:
            # 候補の名前は往復ごとに揺れる（list・source・value…・2026-09-11 の 7 回で 3 通り）。op・range 以外で、
            # 並びかカンマ区切りの文字を持つ項目を候補として拾う
            for k, v in act.items():
                if k in ('op', 'range', 'sheet', 'type', 'overwrite', 'clear'):
                    continue
                if (isinstance(v, (list, tuple)) and v) or (isinstance(v, str) and (',' in v or '、' in v)):
                    items = v
                    break
        if isinstance(items, (list, tuple)):
            items = ",".join(str(v) for v in items)
        items = str(items or '').strip().replace('、', ',')
        if not items:
            raise ValueError("validation には list（ドロップダウンの候補。配列かカンマ区切り）が要ります")
        return f"validation {rng}", ['validation', _sheet_range(sheet, rng), '--list', items]
    if op in ('row', 'col') or op in _ROWCOL_SPLIT:
        if op in _ROWCOL_SPLIT:
            op, action = _ROWCOL_SPLIT[op]            # row_insert 等（2026-09-06 夜・消す手を名前に出した）
        else:
            action = str(act.get('action') or 'insert').strip().lower()
        if action not in ('insert', 'delete'):
            raise ValueError(f"{op} の action は insert か delete（{action!r}）")
        if action == 'delete' and not act.get('overwrite') and not _approved():
            # 人の行・列を消す手（2026-09-05 に開けた）。控えで戻せる（--undo）が、
            # 他のシートからの参照は #REF! に化ける＝承認の言葉があるときだけ通す。
            # 依頼文に承認の言葉があるとき（_approved）は AI の旗を待たない＝承認は道具が依頼文で確かめる設計
            # （9/6 夜）なので旗は重複。旗が無いだけで 1 往復（sonnet 25 秒）を捨てていた（2026-09-08）
            raise ValueError(f'{op} の delete は "overwrite": true が要ります'
                             '（依頼・補足・手順書に承認の言葉「消してよい・削除してよい」があるときだけ）。'
                             '無ければ消す候補を report に書いて止まってください')
        at = act.get('at')
        count = int(act.get('count') or 1)
        if at is None and op == 'row':
            # AI は "rows": [16] や "row": 16 と書きがち（gemini も sonnet も 1 通目はこれ・2026-09-08）＝連番なら at と count に畳む
            rows = act.get('rows', act.get('row', act.get('range')))
            if isinstance(rows, (list, tuple)) and rows:
                try:
                    seq = sorted({int(r) for r in rows})
                except Exception:
                    raise ValueError("row の rows は行番号（数）の並びです")
                if seq != list(range(seq[0], seq[0] + len(seq))):
                    raise ValueError(f"row の rows は連番だけです（{seq}）。飛び飛びなら手を分けてください")
                at, count = seq[0], len(seq)
            elif rows is not None:
                at = rows
        if count < 1:
            raise ValueError(f"{op} の count は 1 以上です")
        if op == 'row':
            span = _row_range(at)
            if span:
                at, count = span[0], span[1] or count
            try:
                at = int(at)
            except Exception:
                raise ValueError("row の at は行番号（数）です")
            if at < 1:
                raise ValueError("row の at は 1 以上です")
        else:
            at = str(at or '').strip().upper()
            if not (at.isascii() and at.isalpha()):
                raise ValueError(f"col の at は列文字（A〜XFD）です（{at!r}）")
        return f"{op} {action} {at} x{count}", [op, action, str(at), str(count), '--sheet', sheet]
    if op == 'copy_range':
        src = _check_addr(act.get('src'))
        dst = _check_addr(act.get('dst'))
        toks = ['copy-range', '--sheet', sheet, src, _sheet_range(sheet, dst), '--show']
        if act.get('values'):
            toks.append('--values')
        return f"copy_range {src}→{dst}", toks
    if op == 'find':
        text = str(act.get('text') or '')
        if not text.strip():
            raise ValueError("find には text が要ります")
        toks = ['find']
        if not act.get('book'):
            toks += ['--sheet', sheet]      # 対象シートが非アクティブでもここを探す（前はアクティブシート）
        if act.get('whole'):
            toks.append('--whole')
        if act.get('formula'):
            toks.append('--formula')
        if act.get('book'):
            toks.append('--book')
        # 先頭が '-' の文字（「-要確認」）は '--' の後ろに置く（前は「不明な引数」で落ちた・2026-09-04）
        toks += ['--', text]
        return f"find {text[:30]!r}", toks
    if op == 'export_pdf':
        path = str(act.get('path') or '').strip()
        if not path.lower().endswith('.pdf'):
            raise ValueError("export_pdf の path は .pdf で終わるファイル名です")
        if os.path.exists(path):
            raise ValueError(f"そのファイルは既にあります（上書きしない）: {path}")
        rng = act.get('range')
        toks = ['export-pdf', path]
        if rng:
            toks += ['--range', _sheet_range(sheet, _check_addr(rng))]
        else:
            toks += ['--sheet', sheet]
        return f"export_pdf {os.path.basename(path)}", toks
    if op == 'chart_config':
        action = str(act.get('action') or '').strip().lower()
        if action not in _CHART_CFG_ACTIONS:
            raise ValueError(f"chart_config の action は {'/'.join(_CHART_CFG_ACTIONS)}（{action!r}）")
        chart = str(act.get('chart') or '').strip()
        if not chart:
            raise ValueError("chart_config には chart（グラフ名。材料の「グラフ」に出ている名前）が要ります")
        args = act.get('args') or []
        if not isinstance(args, (list, tuple)):
            args = [args]
        toks = ['chart-config', action, chart] + [str(a) for a in args]
        for k in ('min', 'max', 'major', 'minor'):
            if act.get(k) not in (None, ''):
                toks += [f'--{k}', _cell_text(act[k])]
        for k in ('value', 'percent', 'category', 'series'):
            if act.get(k):
                toks.append(f'--{k}')
        return f"chart_config {action} {chart}", toks
    if op == 'hyperlink':
        if act.get('remove'):
            raise ValueError("ハイパーリンクを消す手は無い（消す候補は report に書く）")
        cell = _check_addr(act.get('cell'))
        url = str(act.get('url') or '').strip()
        if not url:
            raise ValueError("hyperlink には cell と url が要ります")
        toks = ['hyperlink', _sheet_range(sheet, cell), url]
        if act.get('text'):
            toks.append('--text=' + str(act['text']))       # 先頭が '-' の文字でも「値が無い」で落ちない
        return f"hyperlink {cell}", toks
    if op == 'comment':
        if act.get('remove'):
            raise ValueError("コメントを消す手は無い（消す候補は report に書く）")
        cell = _check_addr(act.get('cell'))
        text = str(act.get('text') or '')
        if not text.strip():
            raise ValueError("comment には cell と text が要ります")
        return f"comment {cell}", ['comment', _sheet_range(sheet, cell), '--', text]
    if op == 'sheet_op':
        action = str(act.get('action') or '').strip().lower().replace('_', '-')
        if action not in _SHEET_OP_ACTIONS:
            raise ValueError(f"sheet_op の action は {'/'.join(_SHEET_OP_ACTIONS)}（{action!r}）。"
                             "シートを消す・隠す手は無い")
        # 名前の置き場は name。AI は sheet / to にも書いてくる（2026-09-06 深夜の実射で
        # 「sheet_op には name が要ります」が 2 弾で往復を 1 回ずつ潰した）。読み替える
        name = str(act.get('name') or act.get('sheet') or
                   (act.get('to') if action == 'add' else '') or '').strip()
        if not name:
            raise ValueError("sheet_op には name（シート名）が要ります")
        # 対象シートの名前を変えると、この後の画像・差分・検査が「シートがありません」で落ちる（2026-09-04）
        if action == 'rename' and name == sheet:
            raise ValueError(f"対象シート「{sheet}」の名前は変えられません（次の依頼で。この往復はこのシートを直す）")
        toks = ['sheet', action, name]
        if action in ('rename', 'copy', 'tab-color'):
            to = str(act.get('to') or '').strip()
            if not to:
                raise ValueError(f"sheet_op {action} には to（新しい名前／色 #RRGGBB）が要ります")
            if action == 'copy' and to == sheet:
                raise ValueError(f"複写の名前を対象シート「{sheet}」と同じにはできません")
            toks.append(to)
        if act.get('after'):
            toks += ['--after', str(act['after'])]
        if act.get('before'):
            toks += ['--before', str(act['before'])]
        return f"sheet_op {action} {name}", toks
    # name_add
    nm = str(act.get('name') or '').strip()
    rng = _check_addr(act.get('range'))
    if not nm:
        raise ValueError("name_add には name（名前）と range が要ります")
    return f"name_add {nm}", ['name', 'add', nm, _sheet_range(sheet, rng)]


def _action_to_tokens(act, sheet):
    """AI の手 1 つ → (表示名, vba_manager のコマンド引数列)。規約外は ValueError。"""
    if not isinstance(act, dict):
        raise ValueError("actions の要素がオブジェクトではありません")
    op = str(act.get('op') or '')
    if op not in _ALLOWED_OPS:
        raise ValueError(f"許していない手です: {op!r}（" + " / ".join(_ALLOWED_OPS) + "）")
    if op in _HEAVY_OPS:
        return _heavy_action_to_tokens(act, sheet)
    sheet = _act_sheet(act, sheet)      # "sheet" を書けばそのシートへ（2026-09-05・書ける先を広げた）
    if op in _TABLE_OPS:
        return _table_action_to_tokens(act, sheet)
    if op == 'read':
        rng = _check_addr(act.get('range'))
        # 大きい表を丸読みすると、材料が会話を埋めて往復が尽きる（2026-09-04・本物のブックで
        # 307 行を read して「重複を見つける」が何もできずに終わった）。行数で断り、道具へ誘導する
        _m = re.match(r'^[A-Za-z]{1,3}(\d+):[A-Za-z]{1,3}(\d+)$', rng.replace('$', ''))
        if _m and (int(_m.group(2)) - int(_m.group(1)) + 1) > _READ_ROW_LIMIT:
            raise ValueError(
                f"read は {_READ_ROW_LIMIT} 行までです（{rng} は "
                f"{int(_m.group(2)) - int(_m.group(1)) + 1} 行）。大きい表は丸読みせず、"
                "材料の列プロファイル（全行ぶんの型と乱れ）を使い、直すときは normalize（列に規則を当てる）・"
                "find_replace・fill・sort・cond_format・table を使ってください。"
                "どうしても中身を見るなら先頭 40 行など範囲を絞って read してください")
        other = str(act.get('sheet') or '').strip()
        target = other or sheet
        toks = ['read-range', '--sheet', target, rng]
        if act.get('formula'):
            # 数式は既定の 40 字で切ると外部参照やパスが読めない（控えを報告に書けない・2026-09-04 実射で FAIL）
            toks += ['--formula', '--width', '200']
        return (f"read {target}!{rng}" if other else f"read {rng}"), toks
    if op == 'write_cells':
        cells = act.get('cells')
        if not isinstance(cells, dict) or not cells:
            raise ValueError("write_cells には cells（{番地: 値}）が要ります")
        toks = ['write-cells', '--sheet', sheet, '--show', '--']   # 値の先頭が '-' でも位置引数として通す
        for a, v in cells.items():
            toks += [_check_addr(a), _cell_text(v)]
        return f"write_cells {len(cells)} セル", toks
    if op == 'write_grid':
        rng = _check_addr(act.get('range'))
        rows = act.get('rows')
        if not isinstance(rows, list) or not rows or not all(isinstance(r, list) for r in rows):
            raise ValueError("write_grid には rows（配列の配列）が要ります")
        if len(rows) > _WRITE_GRID_MAX_ROWS:
            raise ValueError(f"write_grid は {_WRITE_GRID_MAX_ROWS} 行まで（{len(rows)} 行）。数式なら fill で末尾まで伸ばす、"
                             "値の整形なら normalize で規則を当てる（AI が全行を書かない）")
        return f"write_grid {rng} {len(rows)}行", ['write-range', '--sheet', sheet, '--tsv', _LAST_VALUES_FILE,
                                                    '--show', rng]
    if op == 'clear_range':
        # 値・書式を消す手（2026-09-04）。控え（--undo）で戻せるので渡す。行・列そのものは row／col の delete。
        # 承認の言葉が要る（人の表を消す手なので、依頼・補足・手順書に「消してよい」があるときだけ）
        if not act.get('overwrite'):
            raise ValueError('clear_range は人の表を消すので "overwrite": true が要ります'
                             '（依頼・補足・手順書に承認の言葉「消してよい・置き換えてよい」があるときだけ）')
        rng = _check_addr(act.get('range'))
        what = str(act.get('what') or 'contents').strip().lower()
        if what not in ('contents', 'formats', 'all'):
            raise ValueError(f"clear_range の what は contents（値だけ・既定）| formats（書式だけ）| all（両方）（{what!r}）")
        return f"clear_range {rng} {what}", ['clear-range', '--sheet', sheet, f"--{what}", rng]
    if op == 'find_replace':
        if not act.get('overwrite'):
            raise ValueError('find_replace は元の値を書き換えるので "overwrite": true が要ります（依頼・補足・手順書に承認の言葉があるときだけ）')
        rng = _check_addr(act.get('range'))
        search, repl = act.get('search'), act.get('replace')
        if not isinstance(search, str) or not search or not isinstance(repl, str):
            raise ValueError("find_replace には range と search・replace（文字列）が要ります")
        toks = ['find-replace', '--sheet', sheet]
        if act.get('whole'):
            toks.append('--whole')
        if act.get('match_case'):
            toks.append('--match-case')
        toks += ['--', search, repl, rng]        # 先頭が '-' の文字（'－'→'-' の置換）でも通す
        return f"find_replace {rng} {search[:20]!r}→{repl[:20]!r}", toks
    if op == 'format':
        rng = _check_addr(act.get('range'))
        toks = ['format-range', '--sheet', sheet, '--show', rng]
        n = 0
        for k, (opt, has_val) in _FORMAT_KEYS.items():
            if k not in act or act[k] in (None, False, ''):
                continue
            if k in _FORMAT_CHOICES and str(act[k]).strip().lower() not in _FORMAT_CHOICES[k]:
                raise ValueError(f"format の {k} は {' | '.join(_FORMAT_CHOICES[k])} のどれか（{act[k]!r}）")
            if has_val and k in ('bg', 'color', 'font', 'number_format'):
                toks.append(f"{opt}={_cell_text(act[k])}")   # 先頭が '-' の値でも「値が無い」で落ちない
            else:
                toks.append(opt)
                if has_val:
                    toks.append(_cell_text(act[k]))
            n += 1
        if n == 0:
            raise ValueError("format に当てる項目がありません（bold/bg/color/align/number_format/col_width/wrap/border …）")
        return f"format {rng}", toks
    ranges = act.get('ranges') or ([act['range']] if act.get('range') else [])
    if not ranges:
        raise ValueError("tidy には ranges（表の範囲の配列）が要ります")
    addrs = [_check_addr(r) for r in ranges]
    return "tidy " + " ".join(addrs), ['tidy', '--sheet', sheet] + addrs


def _rows_to_tsv(rows):
    return "\n".join("\t".join(_tsv_text(v) for v in r) for r in rows) + "\n"


# ----------------------------------------------------------------
# 道具の呼び出し（同じプロセス・同じ COM 接続。出力を文字列で受ける）
# ----------------------------------------------------------------

@contextlib.contextmanager
def _capture():
    """print の出力をこの呼び出しのあいだだけ受ける。MCP（スレッド別の出力先）でも CLI でも効く。"""
    buf = io.StringIO()
    local = getattr(type(sys.stdout), '_local', None)          # MCP の _ThreadRouted
    if local is not None and getattr(local, 'out', None) is not None:
        saved = local.out
        local.out = buf
        try:
            yield buf
        finally:
            local.out = saved
    else:
        with contextlib.redirect_stdout(buf):
            yield buf


def _pin_target(wb):
    """対象ブックの名指し。保存済みのブックならフルパス（位置引数の先頭に置くと対象ブックになる）、未保存なら None。

    ループの途中で人が Excel 側で別のブックをアクティブにしても、道具はこのブックだけを触る
    （2026-09-03 実射で、開かれた別ブックに grep／get が流れた＝近い事故の記録）。
    wb.FullName は一時的な COM 混雑（VBA_E_IGNORE 等）で失敗することがある（Excelコンボの python
    エンジンが OnTime で 1 秒ごとに様子を見に来るのと重なる瞬間・2026-09-03 実射で 1 回踏んだ）。
    ここで諦めて None を返すと、呼び出し側は「アクティブブックを自動使用」にフォールバックし、
    その瞬間 Excel 側で人が見ていた別のブックへ書いてしまう危険がある＝黙って諦めない。
    """
    if wb is None:
        return None
    from vbam_core import looks_like_xl_file
    last = None
    for i in range(5):
        try:
            full = str(wb.FullName)
            return full if (looks_like_xl_file(full) and os.path.exists(full)) else None
        except Exception as ex:
            last = ex
            if not _com_is_busy(ex):
                return None             # 混雑以外（オブジェクトが本当に死んでいる等）は諦めてよい
            time.sleep(0.3)
    print(f"⚠ 対象ブックの名指しに失敗しました（COM混雑が続いています）: {last}")
    return None                        # 未保存の新規ブック（Book1 等）は Activate で固定する


def _run_cmd(tokens, wb=None):
    """vba_manager のコマンドを同じプロセスで実行し (ok, 出力) を返す。wb を渡すとそのブックに固定する。"""
    import vba_manager as vm
    parser = vm.build_parser()
    table = vm._command_table()
    # 引数の形が違うと argparse は usage を stderr に出して SystemExit(2)＝ループごと落ちる（2026-09-04）。
    # ここで受けて「失敗した手」として AI に返す（次の往復で直せる）。
    err = io.StringIO()
    try:
        with contextlib.redirect_stderr(err):
            ns, unknown = parser.parse_known_args(list(tokens))
    except SystemExit:
        msg = err.getvalue().strip().splitlines()
        return False, "引数の形が違います: " + (msg[-1] if msg else " ".join(str(t) for t in tokens))
    if unknown:
        return False, "不明な引数: " + " ".join(unknown)
    full = _pin_target(wb)
    if full and hasattr(ns, 'posargs'):
        ns.posargs = [full] + list(ns.posargs or [])    # 解析後に足す（引数列に挟むと位置引数が分断される）
    elif wb is not None:
        from vbam_core import pin_active_workbook
        pin_active_workbook(wb)      # 未保存のブックは名指しできない＝このプロセスの「アクティブ」に据える
        # 名指しできるときは Activate を撃たない
        #  ＝Excelコンボが待っている（モードレス表示＋DoEvents ループ）最中に外から Activate すると
        #    UI スレッドと取り合って固まる（2026-09-03・フォーム経由の実射が materials の直前で 15 分停止）
        try:
            wb.Activate()
        except Exception:
            pass
    with _capture() as buf:
        try:
            ok = table[ns.command](ns)
        except SystemExit as ex:
            ok = ex.code in (0, None)
        except Exception as ex:
            print(f"エラー: {ex}")
            ok = False
    return ok is not False, buf.getvalue()


def _stop_rest(results, actions, idx):
    """失敗した手（results の末尾）の後ろに残る手を、実行せずに「止めました」で積む（2026-09-05）。"""
    failed_label = results[-1][0]
    for rest in actions[idx + 1:]:
        rop = (rest or {}).get('op', '?') if isinstance(rest, dict) else '?'
        results.append((str(rop), False,
                        f"実行していません: 前の手「{failed_label}」が失敗したので止めました。"
                        "並べ直して返してください"))


# ----------------------------------------------------------------
# 事前の門: 非空セルへの書き込みを実行前に止める（2026-09-06・Claude for Excel の allow_overwrite を輸入）
#
# 不変条件（頼んでいない変化）は「書いた後」に咎める事後検査で、戻すのは人の手番になる。
# 向こうは道具の側で事前に止めていて、既存シートを直す仕事ではほぼ毎回 1 度この門で止まる。
# 誤爆の代償は「1 往復増える」だけ＝破壊よりずっと安い、というのが向こうの言い分（2026-09-06 の聞き取り）。
# fill・normalize・clear_range・find_replace には既にこの門があり、write_cells と write_grid だけ無かった。
#
# 向こうと違えた点: **同じ値の書き直しは通す**。こちらは仕上げ検査・採点の差し戻しで同じ表を
# 書き直す流れがあり、そこで止めると往復を 1 回無駄にするだけで誰も守れない。
# 判定は呼び出しごと（1 つでも人の値があれば、その手は丸ごと実行しない）＝部分適用があると
# 「どこまで書けたか」を AI が追えない。
# ----------------------------------------------------------------

_OVERWRITE_OPS = ('write_cells', 'write_grid')
_OVERWRITE_SHOW = 8            # エラー文に並べる番地の数
_OVERWRITE_MARK = 'もとからある値が'      # 門で止めた結果の目印（往復の枠を 1 回足すため）


def _grid_box(rng, n_rows, n_cols):
    """write_grid の書き先。左上だけの番地なら rows の大きさへ広げる → 'A1:C5'（純 Python）。"""
    box = _range_box(rng)
    if not box:
        return None
    r1, c1, r2, c2 = box
    r2 = max(r2, r1 + max(int(n_rows or 1), 1) - 1)
    c2 = max(c2, c1 + max(int(n_cols or 1), 1) - 1)
    return f"{_col_letter(c1)}{r1}:{_col_letter(c2)}{r2}"


def _cell_taken(cur, new_text):
    """そのセルに人の値があり、これから書く値と違うか（同じ値の書き直しは破壊ではない・純 Python）。"""
    cur = '' if cur is None else str(cur)
    if cur == '':
        return False
    new = str(new_text if new_text is not None else '')
    return cur != new and cur != new.lstrip("'")


def _overwrite_hits(act, sheet, wb):
    """write_cells / write_grid の書き先にある「人の値」の番地。空なら書いてよい（COM は 1〜2 回）。"""
    op = str((act or {}).get('op') or '')
    if op not in _OVERWRITE_OPS or not isinstance(act, dict) or act.get('overwrite') or wb is None:
        return []
    target = _act_sheet(act, sheet)
    ws = wb.Sheets(target)
    hits = []
    if op == 'write_cells':
        cells = act.get('cells')
        if not isinstance(cells, dict):
            return []
        for a, v in cells.items():
            try:
                cur = ws.Range(str(a)).Formula
            except Exception:
                continue
            if _cell_taken(cur, _cell_text(v)) and not _own_has(target, a):    # 自分が書いたセルは人の値ではない
                hits.append(str(a).replace('$', ''))
        return hits
    rows = act.get('rows')
    if not isinstance(rows, list) or not rows:
        return []
    n_cols = max((len(r) for r in rows if isinstance(r, list)), default=1)
    box = _grid_box(act.get('range'), len(rows), n_cols)
    if not box:
        return []
    try:
        cur = _rows_of(ws.Range(box).Formula)
    except Exception:
        return []
    r1, c1, _r2, _c2 = _range_box(box)
    for i, line in enumerate(rows):
        if not isinstance(line, list):
            continue
        for j, v in enumerate(line):
            try:
                f = cur[i][j]
            except (IndexError, TypeError):
                continue
            a = f"{_col_letter(c1 + j)}{r1 + i}"
            if _cell_taken(f, _tsv_text(v)) and not _own_has(target, a):
                hits.append(a)
    return hits


def _overwrite_error(op, hits):
    """門で止めたときに AI へ返す文（次の往復でそのまま直せる形にする）。"""
    show = " ".join(hits[:_OVERWRITE_SHOW]) + (" …" if len(hits) > _OVERWRITE_SHOW else "")
    return (f"{op} の書き先に、もとからある値が {len(hits)} セルあります: {show}。"
            "この手は実行していません（1 セルでもあると手ごと止めます）。"
            "置き換えてよいなら同じ手に \"overwrite\": true を足して返してください"
            "（依頼・補足・手順書に承認の言葉「置き換えてよい・上書きしてよい・直してよい・埋めてよい」が"
            "あるときだけ）。空いている所へ書くなら番地を変える、中身を見てから決めるなら read で読む。"
            "承認が無く、消してよいか分からないなら、その分を report に書いて人に聞いてください")


# ----------------------------------------------------------------
# 承認の言葉（2026-09-06 夜・Claude for Excel の査読 (h)「語の解釈を AI にさせず、道具が確かめて approved を付ける形に」）
#
# それまでは AI が自分で "overwrite": true を付け、道具は「書き先が空か」しか見ていなかった＝依頼文に承認の
# 言葉があるかは誰も確かめていなかった。ここで道具が依頼文（手順書なら補足）を読み、無ければその手を実行しない。
# 強い動詞（消す・削除・上書き・置き換え・書き換え・切る）は「〜して」の命令形でも承認と読む（元の値が消える
# ことを人が言っている）。弱い動詞（埋める・直す・整理する・揃える）は「〜てよい」の許可形だけ。
# 「〜してはいけない」「〜しないで」は承認ではない。--cont の答え（「はい」「お願いします」）は、道具が聞いた
# 問いへの返事なので承認と読む（answer_mode）。
# ----------------------------------------------------------------
_APPROVAL_STRONG = r'(消し|消去し|削除し|上書きし|置き換え|置換し|書き換え|入れ替え|切っ|捨て|取り除い|外し|除去し)'
_APPROVAL_SOFT = r'(埋め|整理し|直し|変え|変更し|揃え|そろえ|統一し|まとめ|更新し|反映し|書き戻し|書き込ん|除い|落とし|詰め)'
_APPROVAL_PERMIT = (r'(て|ても|た)?\s*(よい|いい|良い|OK|ＯＫ|オッケー|構わない|構いません|構わん|かまわない|かまいません|'
                    r'大丈夫|問題ない|問題ありません|結構です)')
_APPROVAL_NEG = re.compile(r'^(はいけ|はだめ|はダメ|は駄目|はいや|は嫌|は困|しまわ|しまう|しまい|ほしくな|欲しくな|は不可|禁止|はなら|'
                           r'はいらな|は要らな|はいけま)')
_APPROVAL_WORDS = re.compile(r'上書き(可|OK|ＯＫ)|削除(可|OK|ＯＫ)|置換(可|OK|ＯＫ)|承認します|承認済|許可します|直接書い')
_APPROVAL_ANSWER = re.compile(r'^\s*(はい|うん|ええ|OK|ＯＫ|お願い|進めて|やって|それで|いいよ|いいです|大丈夫|問題ない|了解|承知|'
                              r'どうぞ|頼む|続けて|全部)')
_APPROVAL_STANDING = 'この手順書が承認'
_APPROVAL_MARK = '承認の言葉が見つかりません'      # 承認なしで止めた結果の目印（関所として数えるため）


def _has_approval(text, answer_mode=False):
    """依頼文に承認の言葉があるか（純 Python）。"""
    s = str(text or '')
    if not s.strip():
        return False
    if _APPROVAL_STANDING in s or _APPROVAL_WORDS.search(s):
        return True
    if re.search(_APPROVAL_STRONG + _APPROVAL_PERMIT, s) or re.search(_APPROVAL_SOFT + _APPROVAL_PERMIT, s):
        return True
    for m in re.finditer(_APPROVAL_STRONG + r'(て|ちゃって)', s):
        if not _APPROVAL_NEG.match(s[m.end():m.end() + 8]):
            return True
    if answer_mode and _APPROVAL_ANSWER.match(s):
        return True
    return False


def _approval_scope(request):
    """承認の言葉を探す範囲（純 Python）。手順書の依頼文は本文に「置き換えてよい」が説明として出るので、
    補足だけを見る（手順書自身が承認している手は本文の「この手順書が承認」で通す）。"""
    s = str(request or '')
    if s.startswith('手順書「'):
        body, sep, sup = s.rpartition('\n補足: ')
        standing = _APPROVAL_STANDING if _APPROVAL_STANDING in body else ''
        return (sup if sep else '') + ' ' + standing
    return s


_APPROVAL_CTX = {'text': '', 'answer': False}        # run_agent が走行のあいだ置く（無人の往復＝1 走行 1 つ）


def _approval_set(text, answer_mode=False):
    _APPROVAL_CTX['text'] = str(text or '')
    _APPROVAL_CTX['answer'] = bool(answer_mode)


def _approved():
    return _has_approval(_APPROVAL_CTX['text'], _APPROVAL_CTX['answer'])


def _column_space_kind(ws, addr):
    """書き先の列（同じ表の中）の「文字の間の空白」が全角だけなら '　'、半角だけなら ' '、混在・無しは None（COM 3 回）。"""
    try:
        cell = ws.Range(addr)
        reg = cell.CurrentRegion
        r0, c = int(reg.Row), int(cell.Column)
        vals = ws.Range(ws.Cells(r0, c), ws.Cells(r0 + int(reg.Rows.Count) - 1, c)).Value
    except Exception:
        return None
    rows = vals if isinstance(vals, tuple) else ((vals,),)
    zen = han = 0
    for r in rows:
        v = r[0] if isinstance(r, tuple) else r
        if not isinstance(v, str):
            continue
        for w in _INNER_WS_RE.findall(v.strip(_WS_CHARS)):
            if '　' in w:
                zen += 1
            else:
                han += 1
    if zen and not han:
        return '　'
    if han and not zen:
        return ' '
    return None


def _match_column_spaces(act, sheet, wb):
    """write_cells の値の「文字の間の空白」を、書き先の列の空白の種類に合わせる → 直した番地の文の一覧（2026-09-08）。

    C 列 14 行が半角空白なのに、AI は材料の注記と規則文の 2 か所に「半角」とあっても C10 を「イトウ　サクラ」（全角）で
    書いた。列の中で種類が揃っているときだけ道具が合わせる（機械的な判断は道具が持つ）。数式・空白の無い値は触らない。
    """
    notes = []
    if wb is None or not isinstance(act, dict) or not isinstance(act.get('cells'), dict):
        return notes
    try:
        ws = wb.Sheets(_act_sheet(act, sheet))
    except Exception:
        return notes
    for a, v in list(act['cells'].items()):
        if not isinstance(v, str) or v.startswith('=') or not _INNER_WS_RE.search(v):
            continue
        kind = _column_space_kind(ws, a)
        if kind is None:
            continue
        new = _INNER_WS_RE.sub(kind, v)
        if new != v:
            act['cells'][a] = new
            label = '全角' if kind == '　' else '半角'
            notes.append(f"{a} の空白を列に合わせて{label}に")
    return notes


def _needs_approval(act):
    """人の値が消える・書き換わる手か（純 Python）。write は書き先に値があるときだけ＝別に見る。"""
    if not isinstance(act, dict):
        return False
    op = str(act.get('op') or '')
    action = str(act.get('action') or '').strip().lower()
    if op in ('clear_range', 'find_replace', 'row_delete', 'col_delete', 'dedupe'):
        return True
    if op in ('row', 'col') and action == 'delete':
        return True
    if op == 'shape' and act.get('delete'):
        return True
    if op == 'normalize' and act.get('overwrite') and _format_only_normalize(act) \
            and _FORMAT_ASK_RE.search(_APPROVAL_CTX['text']):
        # 書き方だけを変える規則を「書き方をそろえて」の依頼で当てる＝頼まれたそのもの（値の意味は変わらない）。
        # 2026-09-11 夜・試験の経費精算で、承認の言葉が無いと門が normalize を全部止め、75 秒・5 往復で何も直らなかった
        return False
    if op in ('normalize', 'fill') and act.get('overwrite'):
        return True
    return False


_FORMAT_ASK_RE = re.compile(r'そろえ|揃え|統一|整え|整理|直し|直して|修正|書き方|表記|きれいに|綺麗に|全角|半角')


def _format_only_normalize(act):
    """normalize の rules が全部「書き方だけを変える規則」か（regex・fill_down・as_text を含めば False）。"""
    from vbam_hands import _FORMAT_RULES
    rules = act.get('rules')
    rules = [rules] if isinstance(rules, str) else rules
    if not isinstance(rules, list) or not rules:
        return False
    return all(isinstance(r, str) and r.strip().lower() in _FORMAT_RULES for r in rules)


def _approval_error(op):
    return (f"{op} は人の値を消す・書き換える手ですが、依頼・補足・手順書に{_APPROVAL_MARK}"
            "（消してよい・置き換えてよい・上書きしてよい・埋めてよい・切ってよい など。道具が依頼文を確かめました）。"
            "この手は実行していません。消す・置き換える候補を report に書き、plan を「要判断」にして人に聞いてください"
            "（人は agent --cont で答えます）")


# ----------------------------------------------------------------
# この走行で道具が書いたセル（2026-09-06 夜・承認の門を入れた直後の実射で踏んだ）。
# AI が同じ走行で自分が書いた式を直そうとしたら、門が「人の値がある」と止めて往復が尽きた
# （突き合わせの弾: F4:F9 に自分で書いた VLOOKUP を、型を合わせた式に書き直せなかった）。
# 自分が書いたセルは人の値ではない＝上書きにも承認にも数えない。行・列が動く手（挿入・削除・並べ替え）の
# 後は番地がずれるので、そのシートの記憶は捨てる（安全側＝以後は人の値として扱う）。
# --cont のときは前回の「変わったセルの明細」から読み戻す。
# ----------------------------------------------------------------
_OWN_WRITES = {}                 # シート名 → 番地の集合
_OWN_MAX_CELLS = 50000           # 1 手で覚える上限（列全体の fill などは覚えきれない＝人の値として扱う側に倒す）


_WIDTH_KEEP = {}                 # シート名 → {列番号: 幅}（この走行で手が決めた列幅）
_TIDY_RANGES = {}                # シート名 → [範囲]（この走行で tidy を当てた所。採点係に現物を見せる）


_RUN_FACTS = {'normalize': []}   # この走行で normalize を当てた (シート, 列, 規則)。終わりの検査で「直らずに残った文字」を探す（2026-09-11）


def _own_reset():
    _OWN_WRITES.clear()
    _WIDTH_KEEP.clear()
    _TIDY_RANGES.clear()
    _RUN_FACTS['normalize'] = []


def _own_add(sheet, addrs):
    got = _OWN_WRITES.setdefault(str(sheet), set())
    for a in addrs:
        got.add(str(a).replace('$', '').upper())


def _own_has(sheet, addr):
    return str(addr).replace('$', '').upper() in _OWN_WRITES.get(str(sheet), ())


def _box_cells(box, limit=_OWN_MAX_CELLS):
    """'A1:C3' → その中の番地の並び（純 Python）。大きすぎれば空（覚えない）。"""
    b = _range_box(box)
    if not b:
        return []
    r1, c1, r2, c2 = b
    if (r2 - r1 + 1) * (c2 - c1 + 1) > limit:
        return []
    return [f"{_col_letter(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def _own_target_box(act):
    """fill／normalize（overwrite）が書き換える範囲（純 Python）。無ければ None。"""
    if not isinstance(act, dict):
        return None
    op = str(act.get('op') or '')
    if op in ('fill', 'normalize') and act.get('range'):
        try:
            return _check_addr(act.get('range'))
        except Exception:
            return None
    return None


def _own_record(act, sheet):
    """成功した手から、道具が書いたセルを覚える（純 Python）。"""
    if not isinstance(act, dict):
        return
    op = str(act.get('op') or '')
    target = _act_sheet(act, sheet)
    action = str(act.get('action') or '').strip().lower()
    if op in ('sort', 'dedupe') or op in _ROWCOL_SPLIT or (op in ('row', 'col') and action in ('insert', 'delete')):
        _OWN_WRITES.pop(str(target), None)           # 番地がずれる＝忘れる（以後は人の値として扱う）
        return
    try:
        if op == 'write_cells' and isinstance(act.get('cells'), dict):
            _own_add(target, act['cells'].keys())
        elif op == 'write_grid' and isinstance(act.get('rows'), list):
            n_cols = max((len(r) for r in act['rows'] if isinstance(r, list)), default=1)
            box = _grid_box(act.get('range'), len(act['rows']), n_cols)
            if box:
                _own_add(target, _box_cells(box))
        elif op == 'fill' and act.get('range'):
            _own_add(target, _box_cells(_check_addr(act['range'])))
        elif op == 'normalize' and act.get('range'):
            r1, c1, r2, _c2 = _range_box(_check_addr(act['range']))
            if act.get('to'):
                tr, tc = _a1_to_rc(_check_addr(act['to']))
                cells = [f"{_col_letter(tc)}{tr + i}" for i in range(r2 - r1 + 1)]
                if act.get('header') and tr > 1:
                    cells.append(f"{_col_letter(tc)}{tr - 1}")
                _own_add(target, cells)
            elif act.get('overwrite'):
                _own_add(target, _box_cells(_check_addr(act['range'])))
        elif op == 'copy_range' and act.get('src') and act.get('dst'):
            r1, c1, r2, c2 = _range_box(_check_addr(act['src']))
            tr, tc = _a1_to_rc(_check_addr(act['dst']).split(':')[0])
            _own_add(target, _box_cells(f"{_col_letter(tc)}{tr}:{_col_letter(tc + c2 - c1)}{tr + r2 - r1}"))
        elif op == 'clear_range' and act.get('range'):
            _own_add(target, _box_cells(_check_addr(act['range'])))
    except Exception:
        pass


def _own_load_from_changes(path=None):
    """--cont の 1 通目: 前回の「変わったセルの明細」（値の行）を、この走行の書いたセルとして読み戻す。"""
    try:
        with open(path or _LAST_AGENT_CHANGES_FILE, 'r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f, delimiter='\t'))
    except Exception:
        return 0
    n = 0
    for r in rows[1:]:
        if len(r) >= 3 and r[0] == '値' and r[1] and r[2]:
            _own_add(r[1], [r[2]])
            n += 1
    return n


def _human_cells_in(wb, sheet, box, limit=_OVERWRITE_SHOW):
    """範囲の中で「値があって、この走行で自分が書いたのではない」セル（人の値）。読めなければ空。"""
    try:
        ws = wb.Sheets(sheet)
        cur = _rows_of(ws.Range(box).Formula)
        r1, c1, _r2, _c2 = _range_box(box)
    except Exception:
        return []
    out = []
    for i, line in enumerate(cur):
        for j, f in enumerate(line):
            a = f"{_col_letter(c1 + j)}{r1 + i}"
            if f not in (None, '') and not _own_has(sheet, a):
                out.append(a)
                if len(out) >= limit:
                    return out
    return out


def _header_dup(wb, sheet, header):
    """col_insert の header と同じ見出しが、見出し行にすでにあるか → 理由の文（無ければ ''）。

    「列を足す」は見出し名で upsert（同名があれば足さない）＝同じ依頼が 2 回来て列がもう 1 本増える事故を
    道具が止める（2026-09-06 夜・輸入候補 7 の残り）。
    """
    try:
        ws = wb.Sheets(sheet)
        hr = _guess_header_row(ws)
        if not hr:
            return ''
        row = int(hr[0])
        ur = ws.UsedRange
        c0, nc = int(ur.Column), int(ur.Columns.Count)
        vals = _rows_of(ws.Range(ws.Cells(row, c0), ws.Cells(row, c0 + nc - 1)).Value)
        want = str(header).strip()
        for j, v in enumerate(vals[0] if vals else []):
            if _text_of(v).strip() == want:
                return (f"同じ名前の見出し「{want}」が {_col_letter(c0 + j)}{row} にすでにあります"
                        "（同じ列を二重に足すのを止めました。その列を使うか、別の名前にしてください）")
    except Exception:
        return ''
    return ''


def _col_insert_header(wb, sheet, act):
    """col_insert の header を、挿入した列の見出し行に書く → 結果に足す 1 行（書けなければその旨）。"""
    try:
        ws = wb.Sheets(sheet)
        hr = _guess_header_row(ws)
        row = int(hr[0]) if hr else 1
        col = str(act.get('at') or '').strip().upper()
        ws.Range(f"{col}{row}").Value = str(act['header'])
        return f"\n見出し「{act['header']}」を {col}{row} に書きました（見出し行の推定 {row} 行目）"
    except Exception as ex:
        return f"\n（見出しを書けませんでした: {ex}）"


# ----------------------------------------------------------------
# 行・列を消す手の並び（2026-09-11）。動画のお題を撃ったら、AI が元の行番号のまま row_delete を上から 8 回並べ、
# 2 回目からは繰り上がった別の行を消した。続く write_cells も消す前の番地で書き、別の会社に電話番号が入った。
# ----------------------------------------------------------------
_AFTER_DELETE_OK = ('tidy', 'read', 'read_sheet', 'inspect', 'eval', 'view', 'find')


def _is_delete_hand(act):
    op = str(act.get('op') or '')
    action = str(act.get('action') or '').strip().lower()
    return op in ('row_delete', 'col_delete', 'dedupe') or (op in ('row', 'col') and action == 'delete')


_ROW_RANGE_RE = re.compile(r'\$?[A-Za-z]{0,3}\$?(\d+)(?::\$?[A-Za-z]{0,3}\$?(\d+))?')


def _row_range(v):
    """行の範囲の書き方（"18:18"・"18:20"・"A18:I20"・"18"）を (先頭行, 行数) に。行数は範囲で書かれたときだけ
    （1 つなら None＝count に任せる）。読めなければ None（純 Python）。
    Haiku は row_delete を "range": "18:18" と書き、「row の at は行番号（数）です」ではねていた（2026-09-11 揺らし）。"""
    m = _ROW_RANGE_RE.fullmatch(v.strip()) if isinstance(v, str) else None
    if not m:
        return None
    a, b = int(m.group(1)), int(m.group(2) or m.group(1))
    return min(a, b), (abs(b - a) + 1 if m.group(2) else None)


def _row_span(act):
    """行を消す手の (先頭行, 行数)。読めなければ None（純 Python）。"""
    rows = act.get('rows', act.get('row', act.get('range')))
    span = _row_range(act.get('at') if act.get('at') is not None else rows)
    try:
        if span:
            return span[0], span[1] or int(act.get('count') or 1)
        if act.get('at') is not None:
            return int(act.get('at')), int(act.get('count') or 1)
        if isinstance(rows, (list, tuple)) and rows:
            seq = sorted({int(r) for r in rows})
            return seq[0], seq[-1] - seq[0] + 1
        if rows is not None:
            return int(rows), int(act.get('count') or 1)
    except Exception:
        return None
    return None


_CLIP_OPS = ('format', 'normalize', 'find_replace', 'cond_format', 'validation', 'fill')


def _split_row_deletes(actions):
    """飛び飛びの行番号を並べた row_delete を、連番のかたまりごとの手に分ける（2026-09-11 午後）。

    AI は "rows": [17, 18, 20, 21, 26, …] と 1 手に書き、道具が「連番だけです。手を分けてください」で
    はねて、並べ直しに 2 往復（65 秒）を捨てた（同じ日に 2 回）。分けるのは機械の仕事。
    """
    if not isinstance(actions, list):
        return False
    out, changed = [], False
    for act in actions:
        op = str(act.get('op') or '') if isinstance(act, dict) else ''
        is_row_del = op == 'row_delete' or (op == 'row' and str(act.get('action') or '').strip().lower() == 'delete')
        rows = act.get('rows', act.get('row')) if is_row_del else None
        if not is_row_del or act.get('at') is not None or not isinstance(rows, (list, tuple)) or len(rows) < 2:
            out.append(act)
            continue
        try:
            seq = sorted({int(r) for r in rows})
        except Exception:
            out.append(act)
            continue
        runs, start = [], seq[0]
        for a, b in zip(seq, seq[1:] + [None]):
            if b is None or b != a + 1:
                runs.append((start, a - start + 1))
                start = b
        if len(runs) <= 1:
            out.append(act)
            continue
        for at, count in runs:
            piece = dict(act)
            piece.pop('rows', None)
            piece.pop('row', None)
            piece['at'], piece['count'] = at, count
            out.append(piece)
        changed = True
    if changed:
        actions[:] = out
        print(f"（飛び飛びの行を消す手を、連番のかたまりごとに {len(out) - len([a for a in out if not isinstance(a, dict)])} 手に分けました）")
    return changed


def _clip_to_data(actions, sheet, wb):
    """書く・整える手の範囲が、シートの値の末尾より下へはみ出していたら末尾で切る（2026-09-11）。

    AI が A6:I300 と書き、空の 244 行に書式・入力規則・条件付き書式が付いて、使用範囲が 275 行目まで膨らんだ
    （画面の写しも印刷も空の行まで伸びる）。末尾は値か数式のある最後の行。

    逆に、同じ返事に行を消す手があるのに範囲が末尾より手前で終わっていたら、末尾まで伸ばす（2026-09-11 午後）。
    消す手は道具が最後に回すのに、AI は「消した後の表」の番地（A6:A35）で normalize を書く＝36〜56 行が
    直らないまま繰り上がり、下の 5 行に汚れが残って不合格になった（138 秒・5 往復）。
    """
    if wb is None:
        return
    last = {}
    del_sheets = {_act_sheet(a, sheet) for a in (actions or [])
                  if isinstance(a, dict) and _is_delete_hand(a)
                  and str(a.get('op') or '') in ('row_delete', 'row', 'dedupe')}
    for act in actions or []:
        if not isinstance(act, dict) or str(act.get('op') or '') not in _CLIP_OPS:
            continue
        m = re.match(r'^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$', str(act.get('range') or '').replace('$', '').strip())
        if not m:
            continue
        want = _act_sheet(act, sheet)
        if want not in last:
            # Cells.Find は道具のサーバーの中で黙って失敗し、1 行も切れていなかった（2026-09-11 実測）。値から数える
            try:
                ur = wb.Sheets(want).UsedRange
                r0 = int(ur.Row)
                vals = _rows_of(ur.Value) if int(ur.Rows.Count) * int(ur.Columns.Count) <= 2000000 else []
                idx = [i for i, row in enumerate(vals) if any(v not in (None, '') for v in row)]
                last[want] = r0 + idx[-1] if idx else 0
            except Exception:
                last[want] = 0
        end = last[want]
        r1, r2 = int(m.group(2)), int(m.group(4))
        if end and r1 <= end < r2:
            new = f"{m.group(1).upper()}{r1}:{m.group(3).upper()}{end}"
            print(f"（{act.get('op')} の範囲を値の末尾で切りました: {act.get('range')} → {new}）")
            act['range'] = new
        elif end and want in del_sheets and r1 < r2 < end:
            new = f"{m.group(1).upper()}{r1}:{m.group(3).upper()}{end}"
            print(f"（行を消す手と同じ返事なので、{act.get('op')} の範囲を値の末尾まで伸ばしました: "
                  f"{act.get('range')} → {new}。消す前の番地で当てます）")
            act['range'] = new


def _move_deletes_last(actions):
    """行・列を消す手を、書く・整える手の後ろへ回す（2026-09-11）。

    AI は同じ返事の番地を全部「返事を出す前のシート」の番地で書く（動画のお題の 2 回とも、dedupe の後ろに
    消す前の番地の normalize を並べた）。消す手を先に実行すると後ろの手が繰り上がった別の行に当たるので、
    消す手は書く・整える手の後ろに回し、行は下から、列は右から消す。消す手より後ろにあった tidy・読む手は
    消した後に回す（tidy は消した後の表全体に当てる）。
    """
    if not isinstance(actions, list):
        return False
    dels = [a for a in actions if isinstance(a, dict) and _is_delete_hand(a)]
    if not dels:
        return False
    first = next(i for i, a in enumerate(actions) if a is dels[0])
    del_ids = {id(a) for a in dels}
    tail = [a for i, a in enumerate(actions) if i > first and id(a) not in del_ids
            and isinstance(a, dict) and str(a.get('op') or '') in _AFTER_DELETE_OK]
    tail_ids = {id(a) for a in tail}
    head = [a for a in actions if id(a) not in del_ids and id(a) not in tail_ids]

    def row_key(a):
        sp = _row_span(a)
        return -sp[0] if sp else 0

    def col_key(a):
        try:
            return -_col_num(str(a.get('at') or 'A').strip().upper())
        except Exception:
            return 0
    rows = sorted([a for a in dels if str(a.get('op') or '') in ('row_delete', 'row')], key=row_key)
    dedupes = [a for a in dels if str(a.get('op') or '') == 'dedupe']
    cols = sorted([a for a in dels if str(a.get('op') or '') in ('col_delete', 'col')], key=col_key)
    new = head + rows + dedupes + cols + tail
    if [id(a) for a in new] == [id(a) for a in actions]:
        return False
    actions[:] = new
    print("（行・列を消す手を、書く・整える手の後ろへ回しました。同じ返事の番地は全部、返事を出す前のシートの番地として"
          "実行します。行は下から消します）")
    return True


def _order_after_delete(actions, sheet, wb, problems):
    """消す手の後ろの tidy を表全体に当て直し、重複を消す依頼では、残る行に相手のいない行を消す手を止める（problems に足す）。"""
    seen = set()
    row_dels = []              # (手の番号, シート, 先頭行, 行数)
    for i, act in enumerate(actions or []):
        if not isinstance(act, dict) or i in problems:
            continue
        op = str(act.get('op') or '')
        want = _act_sheet(act, sheet)
        if _is_delete_hand(act):
            if op in ('row_delete', 'row'):
                span = _row_span(act)
                if span:
                    row_dels.append((i, want, span[0], span[1]))
            seen.add(want)
            continue
        if want in seen and op == 'tidy':
            got = act.get('ranges') or ([act.get('range')] if act.get('range') else [])
            try:
                # 1 セル＝その表全体（消した後の末尾は道具が決める。消す前の番地の範囲だと空の行に罫線が付く）
                act['ranges'] = [_check_addr(x).split(':')[0] for x in got]
                act.pop('range', None)
            except Exception:
                pass
        elif want in seen and op not in _AFTER_DELETE_OK:
            problems[i] = (f"行・列を消す手の後ろに、同じシートの番地を使う手（{op}）があります。"
                           "番地がずれるので実行していません（道具の並べ替えの漏れ）。次の往復で並べ直してください")
    # 重複を消す依頼で、残る行のどれとも重複していない行を消す手は止める（同じ会社の 2 行を両方消した・2026-09-11）
    if wb is None or not row_dels or '重複' not in (_APPROVAL_CTX.get('text') or ''):
        return
    from vbam_view import _guess_header_idx
    for want in {w for _i, w, _a, _n in row_dels}:
        try:
            ur = wb.Sheets(want).UsedRange
            r0 = int(ur.Row)
            rows = [list(r) for r in _rows_of(ur.Value)]
        except Exception:
            continue
        if len(rows) > 20000:
            continue
        mine = [(i, a, n) for i, w, a, n in row_dels if w == want]
        drop = {r - r0 for _i, a, n in mine for r in range(a, a + n) if 0 <= r - r0 < len(rows)}
        missing = set(_no_twin(rows, _guess_header_idx(rows), drop))
        for i, a, n in mine:
            hit = [r for r in range(a, a + n) if (r - r0) in missing]
            if hit:
                ex = "／".join(f"行{r}（{next((str(v) for v in rows[r - r0] if v not in (None, '')), '')}）"
                              for r in hit[:6])
                problems[i] = (f"消そうとしている {ex} は、残る行のどれとも重複していません。消すとこの相手が表から"
                               "無くなります（重複を消す依頼）。行番号で消さず、dedupe の手を使ってください")


def _prevalidate(actions, sheet, wb):
    """実行前に全部の手を検査する → {番号: 理由}。空なら全部通る（2026-09-06 夜・輸入候補 12）。

    途中の手で止まると、前の手だけが当たった中途半端な姿が残る（人の表に列だけ入って見出しが無い、など）。
    手を並べ替えるのは依存（sheet_op add → その sheet へ write、col_insert → その列へ fill）を壊すのでしない。
    代わりに、実行前に分かる問題（形の違い・許していない手・承認の無い消す手・人の値がある所への書き込み・
    無いシート・同名の見出し）を全部の手について先に見て、1 つでもあれば 1 手も実行しない。
    AI は全部の問題を一度に直して並べ直せる（往復も減る）。前の手が書き換えるシートへの後ろの手は、
    番地がずれるので実行時の門に任せる。
    """
    _split_row_deletes(actions)
    _clip_to_data(actions, sheet, wb)
    _move_deletes_last(actions)
    problems = {}
    known = None
    made = set()
    dirty = set()
    for i, act in enumerate(actions or []):
        if not isinstance(act, dict):
            problems[i] = "actions の要素がオブジェクトではありません"
            continue
        op = str(act.get('op') or '')
        if op not in _ALLOWED_OPS:
            problems[i] = f"許していない手です: {op!r}（" + " / ".join(_ALLOWED_OPS) + "）"
            continue
        want = _act_sheet(act, sheet)
        if op not in _DIRECT_OPS:
            try:
                _action_to_tokens(act, sheet)
            except ValueError as ex:
                problems[i] = str(ex)
                continue
        if wb is not None and want != sheet and want not in made:
            if known is None:
                try:
                    known = [str(s.Name) for s in wb.Sheets]
                except Exception:
                    known = []
            if known and want not in known:
                problems[i] = (f"シート '{want}' がありません（ある: {', '.join(known)}）。"
                               "先に sheet_op add で作ってください")
                continue
        if _needs_approval(act) and not _approved():
            box = _own_target_box(act)
            if box and wb is not None and want not in dirty and want not in made:
                # fill／normalize の overwrite: 書き換える先が空か自分の書いたセルだけなら、人の値は消えない＝承認は要らない
                humans = _human_cells_in(wb, want, box)
                if not humans:
                    pass
                else:
                    problems[i] = _approval_error(op) + f"（書き先にある人の値: {' '.join(humans)}）"
                    continue
            else:
                problems[i] = _approval_error(op)
                continue
        if op in _OVERWRITE_OPS and wb is not None and want not in dirty and want not in made:
            try:
                probe = dict(act)
                probe.pop('overwrite', None)
                hits = _overwrite_hits(probe, sheet, wb)
            except Exception:
                hits = []
            if hits and act.get('overwrite') and not _approved():
                problems[i] = _approval_error(op) + f"（書き先にある人の値: {' '.join(hits[:_OVERWRITE_SHOW])}）"
                continue
            if hits and not act.get('overwrite'):
                problems[i] = _overwrite_error(op, hits)
                continue
        if op == 'col_insert' and act.get('header') and wb is not None and want not in dirty and want not in made:
            dup = _header_dup(wb, want, act.get('header'))
            if dup:
                problems[i] = dup
                continue
        action = str(act.get('action') or '').strip().lower().replace('-', '_')
        if op == 'sheet_op' and action in ('add', 'copy'):
            made.add(str((act.get('to') if action == 'copy' else act.get('name')) or '').strip())
        if op == 'read_file':
            # 同じ返事の中で「読み込む → その新しいシートへ書く／tidy」と並べられるように（2026-09-09）
            made.add(str(act.get('to') or '').strip())
        if _act_writes(act):
            dirty.add(want)
            dirty |= _other_sheets_in([act])
    _order_after_delete(actions, sheet, wb, problems)
    return problems


def _width_record(act, sheet, wb):
    """format の col_width で決めた列幅を覚える（2026-09-09）。

    手順書「見た目を直す」の実射で、AI は基準シートの幅（A=14 B=22 C=9 D=12 E=26）を
    format で正しく写したあと tidy を当て、tidy の列幅=自動が 5 列すべてを 8.08・6.83… に
    戻していた。報告は「写した列幅: A=14…」のまま＝現物と食い違う（黙殺 1 件）。
    tidy の既定は変えない（人が直に打つ tidy はこれまでどおり自動調整）。
    """
    if wb is None or not isinstance(act, dict) or str(act.get('op') or '') != 'format':
        return
    w = act.get('col_width')
    if w in (None, ''):
        return
    name = _act_sheet(act, sheet)
    try:
        cols = wb.Sheets(name).Range(str(act.get('range') or '')).Columns
        keep = _WIDTH_KEEP.setdefault(name, {})
        for i in range(1, int(cols.Count) + 1):
            keep[int(cols(i).Column)] = float(w)
    except Exception:
        pass


def _tidy_record(act, sheet):
    """この走行で tidy を当てた範囲を覚える（2026-09-09）。

    採点係に渡す【見た目】は見出し行から広がる表を 1 つしか読まない。突合の弾（表 A と表 B が
    横に並ぶ）で AI が `tidy A1:C7 E1:G6` と両方に当てても、2 つ目の現物が渡らず
    「表B に tidy が適用されていません」と差し戻していた（9 状態中 4 状態）。
    """
    if not isinstance(act, dict) or str(act.get('op') or '') != 'tidy':
        return
    name = _act_sheet(act, sheet)
    got = act.get('ranges') or ([act.get('range')] if act.get('range') else [])
    lst = _TIDY_RANGES.setdefault(name, [])
    for a in got:
        a = str(a or '').strip()
        if a and a not in lst:
            lst.append(a)


def _width_restore(act, sheet, wb):
    """tidy の後、この走行で手が決めた列幅を戻す。戻した分の一言を返す（無ければ空文字）。"""
    if wb is None or not isinstance(act, dict):
        return ''
    name = _act_sheet(act, sheet)
    keep = _WIDTH_KEEP.get(name) or {}
    if not keep:
        return ''
    from vbam_core import _col_letter
    back = []
    try:
        ws = wb.Sheets(name)
        cols = ws.Range(str(act.get('range') or '')).Columns
        for i in range(1, int(cols.Count) + 1):
            j = int(cols(i).Column)
            if j not in keep:
                continue
            if abs(float(ws.Columns(j).ColumnWidth) - keep[j]) > 0.01:
                ws.Columns(j).ColumnWidth = keep[j]
                back.append(f"{_col_letter(j)}={keep[j]:g}")
    except Exception:
        return ''
    if not back:
        return ''
    return ("\n（この走行で決めた列幅を戻しました: " + " ".join(back)
            + "＝手で決めた幅は tidy の自動調整より優先します）")


def _execute(actions, sheet, wb=None):
    """AI の actions を順に実行し [(表示名, ok, 出力)] を返す。規約外の手は実行せず理由を結果にする。

    1 つでも失敗したら、それ以降の手は実行せず「止めました」を積んで終わる（2026-09-05）。
    失敗した手の結果を前提にした後続の手（normalize が断られた後の fill・write_grid が断られた後の
    tidy 等）が、見当違いの場所に書き進むのを防ぐ。
    """
    results = []
    problems = _prevalidate(actions, sheet, wb)        # 実行前に全部の手を見る（2026-09-06 夜）
    if problems:
        for i, act in enumerate(actions):
            op = (act or {}).get('op', '?') if isinstance(act, dict) else '?'
            if i in problems:
                results.append((str(op), False, "実行していません: " + problems[i]))
            else:
                results.append((str(op), False,
                                "実行していません: 同じ返事の中に通らない手があるので、1 手も実行していません"
                                "（実行前に全部の手を検査しました）。通らない手を直して、全部をもう一度並べてください"))
        return results
    known = None
    for idx, act in enumerate(actions):
        op = (act or {}).get('op', '?') if isinstance(act, dict) else '?'
        want = _act_sheet(act, sheet)
        if wb is not None and want != sheet:
            # "sheet" で別のシートを指した手は、そのシートが実在するときだけ通す（2026-09-05）。
            # 書き先が無いまま CLI に渡すと分かりにくい COM エラーになる。
            # 無いときは読み直す＝同じ返事の前の手が sheet_op add で作っているかもしれない
            if known is None or want not in known:
                try:
                    known = [str(s.Name) for s in wb.Sheets]
                except Exception:
                    known = []
            if known and want not in known:
                results.append((str(op), False,
                                f"実行していません: シート '{want}' がありません"
                                f"（ある: {', '.join(known)}）。先に sheet_op add で作ってください"))
                _stop_rest(results, actions, idx)
                break
        if op in _DIRECT_OPS:
            try:
                results.append(_direct_action(act, sheet, wb))
            except ValueError as ex:
                results.append((str(op), False, f"実行していません: {ex}"))
            except Exception as ex:
                results.append((str(op), False, f"失敗: {ex}"))
            if results[-1][1]:
                _own_record(act, sheet)                # 自分が書いたセルを覚える（人の値ではない）
        else:
            if op == 'write_cells':
                for note in _match_column_spaces(act, sheet, wb):   # 列の空白の種類に合わせる（2026-09-08）
                    print(f"（{note}）")
            try:
                label, toks = _action_to_tokens(act, sheet)
            except ValueError as ex:
                results.append((str(op), False, f"実行していません: {ex}"))
            else:
                try:
                    hits = _overwrite_hits(act, sheet, wb)       # 事前の門（2026-09-06）
                except Exception:
                    hits = []                                    # 門が読めないときは通す（門で仕事を殺さない）
                if hits:
                    results.append((str(op), False, "実行していません: " + _overwrite_error(op, hits)))
                else:
                    if isinstance(act, dict) and act.get('op') == 'write_grid':
                        with open(_LAST_VALUES_FILE, 'w', encoding='utf-8') as f:
                            f.write(_rows_to_tsv(act['rows']))
                    ok, out = _run_cmd(toks, wb)
                    if ok and isinstance(act, dict) and act.get('op') == 'col_insert' and act.get('header'):
                        out += _col_insert_header(wb, _act_sheet(act, sheet), act)   # 見出しも同じ手で
                    if ok and isinstance(act, dict) and act.get('op') == 'tidy':
                        out += _width_restore(act, sheet, wb)     # tidy の自動調整に、決めた幅を消させない
                    if ok:
                        _own_record(act, sheet)        # 自分が書いたセルを覚える（人の値ではない）
                        _width_record(act, sheet, wb)  # format の col_width を覚える（後の tidy 用）
                        _tidy_record(act, sheet)       # tidy を当てた範囲を覚える（採点係に見せる）
                    results.append((label, ok, out))
        if not results[-1][1]:
            _stop_rest(results, actions, idx)
            break
    return results


# AI に聞く は vbam_ai.py へ切り出した（2026-09-11）


# 手3 相当 は vbam_undo.py へ切り出した（2026-09-11）


# ^_VIEW_MARGIN = は vbam_view_ai.py へ切り出した（2026-09-11）


# 仕上げ検査の関所と自己採点 は vbam_grade.py へ切り出した（2026-09-11）


# 本番の走行台帳 は vbam_ledger.py へ切り出した（2026-09-11）


# ----------------------------------------------------------------
# ループ本体
# ----------------------------------------------------------------

def run_agent(request, sheet, wb, ai=_CC_AI, model=None, max_turns=_DEFAULT_MAX_TURNS, dry_run=False,
              materials=None, verify=True, backup=True, resume=None, resume_results=None, show_image=True,
              base=None, grade=True, ask=None, rules=None, execute=None, mode='sheet',
              grade_ai=None, grade_model=None, approval_text=None, approval_answer=False,
              grade_always=False, macro_first=None):
    """依頼文をシート 1 枚に対して回す。wb はブック（COM）＝往復のあいだ固定。戻り値 dict（done・ok・turns…）。

    grade_always＝一発で通った回（1 往復目で手が全部通り、仕上げ検査・中身・頼んでいない変化に指摘なし）でも
    採点する（--grade）。既定はその回だけ採点を省く（採点役は別プロセスで 8 秒・2026-09-11 午後）。

    approval_text＝承認の言葉を探す文（省略時は依頼文から道具が決める。手順書なら補足だけ）。
    approval_answer＝その文が道具の問いへの答え（--cont）＝「はい」「お願いします」も承認と読む（2026-09-06 夜）。

    resume に前回の往復（[(role, text) …]）を渡すと、その続きとして回す（--continue）。
    base＝終わりの検査の基準 (エラーセル, ###)。渡さなければ「始める前の材料」から数える（元からある
    #N/A を数えないため）。build の直しループは (0, 0) を渡す＝組み上げた直後の壊れた数を基準にすると、
    AI が何もしなくても「増えていない」で合格になっていた（2026-09-04）。
    ask に関数を渡すと _ask の代わりに使う（--replay 用。API を呼ばずキーも要らない＝ループの側だけを無料で確かめる）。
    rules／execute／mode を渡すと、同じ関所（plan・仕上げ検査・頼んでいない変化・自己採点・検査）のまま
    別の規約と実行系で回せる（2026-09-05・シートとマクロをまたぐ統合ループ run_both がこれを使う）。
    backup=True のときは最初の書き込みの前にブック全体の控えも取り、done の前に「頼んでいない変化」を
    照らす（2026-09-05・それまで実射でしか見ていなかった不変条件を本番に配線した）。実射（--fire）は
    backup=False で呼び、弾ごとの inv_skip を持つ自前の検査を使う＝ここでは二重に取らない。
    """
    ai, model, api_key = (ai or 'gemini', model or 'replay', '') if ask else _ai_setup(ai, model)
    # 採点係（2026-09-06）: 既定は同じモデルだが会話は別。--grade-ai で別の AI に採点だけ任せられる
    g_ai, g_model, g_key = ((ai, model, api_key) if (ask or not (grade_ai or grade_model))
                            else _ai_setup(grade_ai or ai, grade_model))
    book = str(wb.Name)
    tool_sec = {}                      # 道具側の秒（段ごと）。AI 待ちの請求書と並べて内訳を出す（2026-09-06 夜）

    def _tick(key, t0):
        tool_sec[key] = tool_sec.get(key, 0.0) + (time.time() - t0)

    secrets = _secret_values(wb)       # 名前定義が KEY／TOKEN／PASSWORD を指すセルの値＝AI に送る文から伏せる

    def _mask(text):
        return _mask_secrets(text, secrets) if secrets else text

    # 承認の言葉は道具が依頼文で確かめる（AI の "overwrite": true を信じない・2026-09-06 夜）
    _approval_set(approval_text if approval_text is not None else _approval_scope(request), approval_answer)
    _own_reset()                                   # この走行で自分が書いたセル（人の値ではない）を数え直す
    if resume:
        n_own = _own_load_from_changes()           # --cont: 前回の走行で書いたセルも「自分の」に含める
        if n_own:
            print(f"（前回の走行で書いたセル {n_own} 個は、この続きでも承認なしで書き直せます）")
    print(f"対象ブック: {book}   シート: {sheet}（この往復のあいだ固定。別のブックには触らない）")
    if resume:
        print(f"前回の続き: 往復 {len(resume) // 2} 回ぶんの会話を引き継ぎます（報告した候補と別のことをしないため）")
    run_id = _run_id()                 # この走行の名札（控えの覚書と走行台帳を結ぶ＝undo された走行が分かる）
    # マクロの先撃ち（2026-09-11 深夜・shu と決めた形）: 書き方の整理の依頼で「表の書き方と罫線と列幅をそろえる」が開いていれば、
    # AI より先に道具が撃つ。残りが無ければここで終わる（AI は呼ばない）。残りがあれば控えを引き継いで往復へ。
    # 実射（backup=False）・--replay（ask）・--cont・統合ループ（rules／execute）では撃たない
    pre = None
    if (macro_first is not False and backup and not dry_run and ask is None and not resume
            and rules is None and execute is None and mode == 'sheet'):
        try:
            from vbam_prefire import macro_first as _prefire
            pre = _prefire(request, sheet, wb, max_turns, run_id)
        except Exception as ex:
            print(f"（マクロの先撃ちをやめて AI の往復で回します: {ex}）")
            pre = None
        if pre and pre.get('result') is not None:
            return pre['result']
    t0 = time.time()
    if materials is None:
        ok_m, materials = _run_cmd(['materials', sheet, '--full'], wb)
        if not ok_m:
            raise RuntimeError("materials が失敗しました:\n" + materials.strip())
        try:
            materials += _extra_materials(wb, wb.Sheets(sheet))
        except Exception as ex:
            materials += f"（追加の材料を読めませんでした: {ex}）\n"
        try:
            fp_now = _sheet_fingerprint(wb.Sheets(sheet))     # 同じ依頼の 2 回目を見分ける指紋（2026-09-06）
        except Exception:
            fp_now = {}
        materials += _notes_recall(book, sheet, request, fp_now)   # 前にこのシートでやったこと（あれば）
        materials += _since_last_run(book, sheet, wb)              # 前回の終わりから人の手で変わった所（2026-09-06 夜）
    if pre and pre.get('ran'):
        materials += pre['note']        # 先撃ちで済んだこと・残りの指摘（AI にやり直させない）
    _tick('材料', t0)
    if _wants_datamodel_note(request, materials):
        materials += _DATAMODEL_NOTE       # 規則文から外した重い段は、要る依頼のときだけ（2026-09-06 夜）
    materials += _peek_files_note(request)  # 依頼にパスがあれば道具が先に読む＝読むだけの往復を作らない（2026-09-09）
    materials = _mask(materials)
    print(materials.rstrip())
    if base is None:
        base = _count_err_hash(materials)      # 始める前のエラーセル・###（終わりの検査は「増えたか」で見る）
    if ai == _CC_AI and not ask and not dry_run:
        # 本体と採点係の claude -p を先に起こす（起動 2.2 秒を AI の考える時間に重ねる・2026-09-11 午後）
        try:
            cc_prewarm([model] + ([g_model] if (grade and g_ai == _CC_AI) else []))
        except Exception:
            pass
    prompt = (_continue_prompt(request, materials, resume_results) if resume
              else _first_prompt(request, book, sheet, materials, rules))
    first_img = None
    if show_image and not dry_run and _wants_first_image(materials):
        # 結合セルの帳票・重なった図形は materials の文字だけでは崩れて見える＝最初の返事の前に見せる（2026-09-05）
        t0 = time.time()
        first_img, first_addr = _shot_for_ai(sheet, wb)
        _tick('画像', t0)
        if first_img:
            prompt += _view_note(first_addr, "結合セルの境目・図形の重なりは文字では見えません。"
                                             "画像で確かめてから手を決めてください")
    with open(_LAST_AGENT_ASK_FILE, 'w', encoding='utf-8') as f:
        f.write(prompt)
    log = open(_LAST_AGENT_LOG_FILE, 'w', encoding='utf-8')
    log.write(json.dumps({'meta': {'book': book, 'sheet': sheet, 'mode': mode, 'request': request,
                                   'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'ai': ai, 'model': model,
                                   'max_turns': max_turns}},
                         ensure_ascii=False) + "\n")
    history = list(resume or [])       # --continue: 前回の往復をそのまま引き継ぐ
    # 引き継いだ会話を新しい記録にも書き戻す（記録は 'w' で開き直すので、書かないと 2 回目の --continue で
    # RULES・元の依頼・材料が消えていた・2026-09-04）
    turn_base = 0
    if resume:
        turn_base = len(resume) // 2
        for i in range(0, len(resume) - 1, 2):
            past = {'turn': i // 2 + 1, 'prompt': resume[i][1], 'reply': resume[i + 1][1], 'resumed': True}
            if i + 2 >= len(resume) and resume_results:
                past['results'] = [{'label': l, 'ok': o, 'out': t} for l, o, t in resume_results]
            log.write(json.dumps(past, ensure_ascii=False) + "\n")
    total = {'in': 0, 'out': 0, 'think': 0, 'sent': 0, 'recv': 0, 'sec': 0.0}
    done = False
    report = ''
    plan = []                          # AI が 1 往復目に書く「やることの一覧」（未が残るうちは done にさせない）
    dropped = []                       # 最後の往復まで「未」だった項目＝落とし物（done にはしない）
    discovered = []                    # 2 往復目以降に plan へ足した項目と理由（report の末尾に写す・2026-09-06 夜）
    last_actions = []
    wanted_done = False                 # 最後の返事が done を立てていたか（actions と同時でも拾う）
    touched = []                       # 書いた範囲（仕上げ検査を当てる先）
    try:
        # 走行の頭のシート名（採点に「作った別シート」を見せるため・2026-09-06 深夜）。
        # 名前は sheets_at_start（往復の中に同名の sheets_before があり、毎往復その時点の
        # 一覧で上書きされる＝採点のときには作ったシートも入っていて差分が空になっていた）
        sheets_at_start = [str(ws.Name) for ws in wb.Sheets] if wb is not None else []
    except Exception:
        sheets_at_start = []
    audited = False                    # 仕上げ検査で差し戻したか（差し戻しは 1 回だけ）
    graded = False                     # 自己採点をしたか（採点も 1 回だけ）
    grade_skipped = False              # 一発で通ったので採点を省いたか（2026-09-11 午後）
    did = []                           # 通った手の名札（【やったこと】を道具が作る・AI に書かせない・2026-09-11 午後）
    audit_notes = []                   # 最後の仕上げ検査の指摘（報告に出す）
    content_noticed = []               # 中身の検査の気づき（空欄・文字の数字・重複。報告に必ず写す・2026-09-06 夜）
    c_block = []                       # 中身の指摘（done を止めるもの）。最後の判定にも使う（2026-09-11）
    unmet_left = []                    # 採点で挙がったまま直せなかった点
    noticed_left = []                  # 採点係が「依頼には無いが気づいた」こと（報告に写すだけ・2026-09-06 夕）
    before = None                  # 最初の書き込みの前のシート（差分と --undo のため）
    others = {}                        # 別シートに書く手を出したときの、そのシートの書き込み前
    book_before = None                 # 最初の書き込みの前のブック全体（頼んでいない変化を見るため）
    wrote_sheets = {sheet}             # 意図して書いたシート（別シートに書いた分を違反と呼ばないため）
    inv_skip = set()                   # 依頼どおりに動かした結果として当然変わるもの（手から決める）
    inv_notes = []                     # 頼んでいない変化（関所と報告に出す）
    read_formula_ranges = set()        # 数式表示で読んだ範囲（読み残しの検査用・2026-09-05）
    coverage_checked = False           # 読み残しの検査は 1 回だけ
    inv_checked = False                # 不変条件で差し戻したか（差し戻しは 1 回だけ）
    over_blocked = False               # 事前の門で止めたか（往復の枠を足すのは 1 回だけ）
    appr_blocked = False               # 承認の無い消す手を止めたか（同上・2026-09-06 夜）
    inv_dirty = False                  # 最後に照らした後にまた書いたか（終わりに見直すかの判定）
    gates = dict.fromkeys(_GATE_LABELS, 0)   # どの関所で何回差し戻したか（2026-09-06・一発合格率を測るため）
    merged = 0
    merged_now = False                 # この往復で手と done を同時に受けたか（未の項目の関所で見る）                         # 手と done を同じ返事で受けて往復を省いた回数（2026-09-06 夜）
    merged_res = ''                    # その回の結果文（関所で差し戻すときは AI がまだ見ていないので先頭に付ける）
    audit_dirty = False                # 仕上げ検査の後にまた書いたか（終わりに見直すかの判定）
    grade_keep = None                  # 採点が通った直後の（材料の全文, 画像）＝終わりの検査が読み直さずに使う
    if pre and pre.get('ran'):
        # 先撃ちで取った控えを引き継ぐ（往復の最初の書き込みで撃った後の姿を控えにしない＝--undo は撃つ前へ戻る）
        before = pre['before']
        book_before = pre['book_before']
        inv_skip |= set(pre.get('inv_skip') or ())
        did += list(pre.get('did') or [])
        touched += list(pre.get('touched') or [])
        for k, v in (pre.get('tool_sec') or {}).items():
            tool_sec['先撃ち・' + k] = v
    turn = 0
    # 関所（仕上げ検査・頼んでいない変化・読み残し・採点）の差し戻しは、依頼をこなす往復とは別の仕事。
    # 前は同じ 4 回の枠を食い合い、最短でも「書く → 検査の直し → 採点の直し → done」で使い切っていた
    # （2026-09-06・規則文を Claude for Excel に査読させて出た指摘）。関所が差し戻すたびに枠を 1 回足す。
    # 足せるのは _GATE_EXTRA_TURNS まで。この 4 つの関所はどれも 1 回しか撃たないので、無限には伸びない。
    limit = max_turns
    t_start = time.time()
    progress_write('start', book=book, sheet=sheet, max_turns=max_turns, request=request[:80],
                   note='材料を読み終えました', sec=job_clock_elapsed())
    try:
        msg = prompt
        img = first_img
        while turn < limit:
            turn += 1
            merged_now = False         # 往復ごとに戻す（2026-09-08）。前は立てっぱなしで、1 度でも
            #                            手と done を同時に受けると、以後の「手ゼロの done」まで
            #                            「手は全部通った回」として扱われていた
            # 画像は最新の 1 通だけに添える。前の往復の user にも 3 つ組で残していたので、送るときに
            # 同じ _last_agent_view.png を読み直し、往復 n で「最新の 1 枚」を n-1 枚送っていた（2026-09-04）
            history = [(h[0], h[1]) for h in history]
            history.append(('user', msg, img))
            # 上限は関所の差し戻しで伸びる。前は max_turns を出していて「往復 5/4」と読めた（2026-09-06 夕）
            print(f"\n===== 往復 {turn}/{limit}"
                  + (f"（上限 {max_turns}＋関所の差し戻し {limit - max_turns}）" if limit > max_turns else "")
                  + " =====")
            progress_write('ask', book=book, sheet=sheet, turn=turn_base + turn, max_turns=limit,
                           request=request[:80], sec=job_clock_elapsed())
            text, usage, sec = (ask or _ask)(ai, model, api_key, history)
            history.append(('model', text))
            total['in'] += usage['in']
            total['out'] += usage['out']
            total['think'] += usage.get('think', 0)
            total['sent'] += len(msg)
            total['recv'] += len(text)
            total['sec'] += sec
            print(_bill_line(len(msg), text, usage, sec, model))
            entry = {'turn': turn_base + turn, 'prompt': msg, 'reply': text, 'usage': usage, 'sec': round(sec, 2)}
            try:
                say, actions, done, report, plan_now = _parse_reply(text)
                why_missing = []
                if plan_now:
                    if plan:                                # 2 往復目からの追加＝現場を見て足した項目（2026-09-06 夜）
                        added = _plan_added(plan, plan_now)
                        discovered += [(p['item'], p.get('why') or '') for p in added]
                        why_missing = [p['item'] for p in added if not p.get('why')]
                    plan = _merge_plan(plan, plan_now)      # 項目は消させない（state だけ更新）
                wanted_done = _wanted_done(text)        # actions と done を同時に返したか
            except ValueError as ex:
                print(f"AI の返事が規約に合っていません: {ex}")
                entry['error'] = str(ex)
                gates['format'] += 1
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                msg = f"【結果】\n返事が規約に合っていません: {ex}\n規約どおりの JSON で返してください。（残りの往復: {limit - turn} 回）\n"
                continue
            if say:
                print(f"AI: {say}")
            if actions:
                hands = " / ".join((a.get('op', '?') if isinstance(a, dict) else '?') for a in actions)
                print(f"手 {len(actions)} 本: " + hands)
                progress_write('run', book=book, sheet=sheet, turn=turn_base + turn, max_turns=limit,
                               say=say[:60], hands=f"手 {len(actions)} 本: {hands}", sec=job_clock_elapsed())
            if dry_run:
                print(json.dumps(actions, ensure_ascii=False, indent=1))
                print("（--dry-run: Excel には触っていません）")
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                break
            if actions:
                last_actions = actions
                if backup and before is None and _will_write(actions):
                    t0 = time.time()
                    before = _sheet_snapshot(wb.Sheets(sheet)) or {}      # 最初の書き込みの前だけ（控え＋差分の元）
                    _agent_backup(wb, wb.Sheets(sheet), sheet, request, run_id)
                    # 不変条件（頼んでいない変化）の控え。ブックが大きいと読むだけで待たせるので上限で切る
                    t_inv = time.time()
                    n_inv_cells = _book_cell_count(wb)
                    if n_inv_cells <= _INV_BOOK_MAX_CELLS:
                        book_before = _snapshot_book(wb, own={sheet})       # 触らないシートは軽い控え（2026-09-16）
                        print(f"控え（頼んでいない変化の検査用）: {len(book_before['sheets'])} シート・"
                              f"{n_inv_cells:,} セル・{time.time() - t_inv:.1f} 秒")
                    else:
                        print(f"頼んでいない変化の検査: しません（ブックが大きい＝{n_inv_cells:,} セル）")
                    _tick('控え', t0)
                inv_skip |= set(_inv_skip_for(actions))                  # 依頼どおりに動かした結果は違反にしない
                # 数式表示で読んだ範囲を覚える（「見ていない範囲を『なし』と報告する」を後で数えるため）
                read_formula_ranges |= {(_act_sheet(a, sheet), str(a.get('range') or ''))
                                        for a in actions if isinstance(a, dict)
                                        and str(a.get('op') or '') == 'read' and a.get('formula')}
                wrote_sheets |= {_act_sheet(a, sheet) for a in actions if _act_writes(a)}
                # 既存のピボットを直す手は置き場所を書かない（名前しか書かない）ので _act_sheet では拾えない。
                # 拾わないと、ピボットの出力シートが「頼んでいないシート」のまま残り、直せば必ず落ちる
                # （2026-09-09 の実射: 別シートに置いたピボットを直す弾 12 本中 8 本がこれ一つで不合格）
                wrote_sheets |= _pivot_sheets_in(actions, wb)
                # 別シートに出す手（ピボット作成・load・sheet_op）と、既存ピボットを直す手の出力シート
                for nm in (_other_sheets_in(actions) | _pivot_sheets_in(actions, wb)):
                    if nm != sheet and nm not in others:
                        try:
                            others[nm] = _sheet_snapshot(wb.Sheets(nm)) or dict(_EMPTY_SNAPSHOT)
                        except Exception:
                            others[nm] = dict(_EMPTY_SNAPSHOT)           # まだ無いシート＝これから作る
                        if backup and before is not None:
                            _undo_add_sheet(wb, nm)      # --undo で戻す対象に足す（書き換える前の今のうちに）
                try:
                    sheets_before = {str(s.Name) for s in wb.Sheets}
                except Exception:
                    sheets_before = None
                t0 = time.time()
                results = (execute or _execute)(actions, sheet, wb)
                _tick('実行', t0)
                inv_dirty = True                  # 照らした後にまた書いた＝終わりにもう一度見る
                audit_dirty = True
                grade_keep = None
                if before is not None:            # この回で控えを取っている＝覚書はこの仕事のもの
                    made = _created_from_results(results)
                    if sheets_before is not None:
                        # ピボットの置き場は道具が黙って作る（出力に「シート追加」が出ない）＝実物で拾う
                        try:
                            for nm in [str(s.Name) for s in wb.Sheets]:
                                if nm not in sheets_before and not any(
                                        c['kind'] == 'sheet' and c['name'] == nm for c in made):
                                    made.append({'kind': 'sheet', 'name': nm})
                        except Exception:
                            pass
                    _record_created(made)
                for label, ok, out in results:
                    print(f"--- {label} {'' if ok else '（失敗）'}---")
                    print(out.rstrip())
                    if ok and not str(label).startswith(('read', 'eval', 'find', 'inspect')):
                        did.append(str(label))
                if (not over_blocked) and any(_OVERWRITE_MARK in str(t) for _l, _o, t in results):
                    # 事前の門（非空セルへの書き込み）で止めた往復も関所の 1 つ。依頼をこなす往復とは
                    # 別枠にする（2026-09-06。この門を入れた日に、承認の言葉つきの弾まで 1 往復失うと
                    # 分かったため）。足すのは 1 回だけ＝同じ所で何度も止まるのは AI の側の問題
                    over_blocked = True
                    limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                    gates['overwrite'] += 1
                elif (not appr_blocked) and any(_APPROVAL_MARK in str(t) for _l, _o, t in results):
                    # 承認の無い消す手を道具が止めた＝AI は report で聞く側へ回る。枠を 1 回足す（2026-09-06 夜）
                    appr_blocked = True
                    limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                    gates['approval'] += 1
                elif any(not ok for _l, ok, _o in results):
                    gates['hand'] += 1          # 手が失敗した往復＝AI は次の往復で並べ直す（一発ではない）
                touched += _audit_targets(actions)
                entry['results'] = [{'label': l, 'ok': o, 'out': t} for l, o, t in results]
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                msg = _mask(_result_prompt(results, limit - turn))
                if any(_APPROVAL_MARK in str(t) for _l, _o, t in results):
                    # 同じ門にもう一度ぶつかる往復を無くす（2026-09-06 夕: 止められた後にまた試して往復を失った）
                    msg += ("（承認の要る手は、この走行では実行されません＝同じ手をもう一度出さないでください。"
                            "その項目は plan で「要判断」にして ask に人への質問を書き、残りの手だけを actions に"
                            "並べてください。残りが無ければ done にしてください）\n")
                if why_missing:
                    msg += ("（plan に足した項目 " + " / ".join(f"「{w}」" for w in why_missing[:3])
                            + " に why（足した理由）が付いていません。次の返事で付けてください）\n")
                # 手が全部通り、最後が tidy（仕上げ）なら done の返事を待たずに検査へ（2026-09-11: 手を全部当てた後、
                # 「完了しました」と言うだけの往復に 30 秒かかった）。足りなければ関所と採点係が差し戻す
                auto_done = (not wanted_done and bool(actions) and isinstance(actions[-1], dict)
                             and str(actions[-1].get('op') or '') == 'tidy')
                if (wanted_done or auto_done) and not dry_run and all(ok for _l, ok, _o in results):
                    # 手と done を同じ返事で受けた（2026-09-06 夜）。14 走行のうち 10 走行が「1 往復目に手・
                    # 2 往復目に actions 空の done」の形で、2 往復目は 12,000〜16,000 トークン・4 秒の儀式だった。
                    # 全部通ったので次の往復を待たずに関所へ進む。関所が差し戻すときは AI がまだ結果を見ていない
                    # ので、結果文を先頭に付けて返す（merged_res）。1 つでも失敗していれば今までどおり結果を返す
                    merged += 1
                    merged_res = _mask(_result_prompt(results, None))
                    done = True
                    merged_now = True     # この返事の plan は「実行前」に書かれている＝未の項目で突き返さない
                    print("（手と done を同じ返事で返し、全部通りました＝往復を使わず検査へ進みます）" if wanted_done
                          else "（手が全部通り、最後が tidy でした＝done の返事を待たずに検査へ進みます）")
                else:
                    t0 = time.time()
                    img, img_addr = (_shot_for_ai(sheet, wb) if (show_image and _will_write(actions))
                                     else (None, None))
                    _tick('画像', t0)
                    if img:
                        msg += _view_note(img_addr, "罫線の継ぎはぎ・寄せ・カンマ・列幅の不足（###）は"
                                                    "文字では見えません。画像で確かめてから done にしてください")
                    continue
            else:
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
            pre = (merged_res + "\n") if merged_res else ""     # 関所の差し戻し文に前置きする結果（手と done を同時に返した回）
            merged_res = ''
            if done:
                pending = _plan_pending(plan)
                if pending and merged_now:
                    # 手と done を同じ返事で受けた回＝その返事の plan は「手を撃つ前」に書かれている。
                    # ここで突き返すと、状態を書き換えるだけの往復を 1 回失う
                    # （実射「パワークエリで読み込み」が毎回これだった・2026-09-07 未明）。
                    #
                    # 2026-09-08: 「全部まだ未」の回だけ見逃していたので、**一部を「済」と書き分けた回だけが
                    # 突き返される**＝正直に書いた側が罰される形になっていた（本番の走行で 12.7 秒を捨てた）。
                    # merged_now はこの返事の手が**全部通った**ことを含む＝「done と言って何もしていない」
                    # 落とし物ではない。残りの検査（仕上げ検査・中身の検査・頼んでいない変化・採点）は
                    # そのまま通るので、ここは道具が「済」に直して先へ進む。
                    print("（手と done を同じ返事で受け、手は全部通りました＝一覧の「未」は手の前に書かれた"
                          "ものとして道具が「済」にします: " + " / ".join(pending[:3])
                          + ("…" if len(pending) > 3 else "") + "）")
                    _plan_mark_done(plan, pending)
                    pending = []
                if pending:
                    # 分解した項目が「未」のまま done＝落とし物。文章の約束でなく道具が受け付けない（2026-09-04）
                    if turn < limit:
                        print("未の項目が残っているので done を受け付けません: " + " / ".join(pending))
                        done = False
                        gates['plan'] += 1
                        img = None
                        msg = pre + ("【結果】\n未の項目が残っています: " + " / ".join(pending)
                               + "。やったなら「済」、できないなら「不可」にして report に理由を書いてください。"
                                 "まだやっていないなら、その手を actions に並べてください。"
                                 f"（残りの往復: {limit - turn} 回）\n")
                        continue
                    # 最後の往復＝続きを頼めない。ここで done を通すと落とし物が「合格」で出ていく
                    # （2026-09-04: 未のチェックが最終往復だけ素通りしていた）
                    done = False
                    dropped = pending
                if done and not dry_run and turn < limit:
                    # (1) 仕上げ検査（道具が実物を見る・AI を使わない）。指摘があるうちは done を受け付けない
                    if touched and not audited:
                        audited = True
                        t0 = time.time()
                        audit_notes = _audit_now(wb, sheet, touched)
                        c_block, content_noticed = _content_now(wb, sheet, touched)
                        r_block, r_noticed = _request_gate(wb, sheet, book_before, request, _RUN_FACTS)
                        c_block = c_block + r_block
                        content_noticed = content_noticed + r_noticed
                        audit_notes = audit_notes + c_block
                        _tick('仕上げ検査', t0)
                        audit_dirty = False
                        if audit_notes:
                            print("仕上げ検査の指摘（done を受け付けません）:")
                            for n in audit_notes:
                                print("  - " + n)
                            done = False
                            gates['audit'] += 1
                            limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                            t0 = time.time()
                            img, img_addr = (_shot_for_ai(sheet, wb) if show_image else (None, None))
                            _tick('画像', t0)
                            msg = pre + ("【結果】\n道具が実物を見た仕上げ検査で、指摘があります:\n"
                                   + "\n".join("- " + n for n in audit_notes)
                                   + "\n見た目の指摘（見出し・罫線・桁区切り・寄せ・###）は、表全体（見出しを含む元の列＋"
                                     "足した列）に tidy を当てて直してください。値は変えないこと。"
                                     "（見出しが空の指摘だけは tidy では直りません＝列の名前を書いてください）"
                                   + ("\n中身の指摘（エラー値・合計の不一致・数式の列の直値）は tidy では直りません＝"
                                      "原因のセルを直してください（式の範囲・参照先・隣と同じ式）。人の値を変えないと"
                                      "直せないものは触らず、report に「どこが・なぜ」を書いてください。"
                                      if c_block else "")
                                   + "直したら done にしてください。"
                                   + f"（残りの往復: {limit - turn} 回）\n")
                            if img:
                                msg += _view_note(img_addr)
                            continue
                    # (1.5) 不変条件＝頼んでいない変化（道具が撃つ前の控えと現物を照らす・AI を使わない）。1 回だけ
                    if book_before is not None and not inv_checked:
                        inv_checked = True
                        t0 = time.time()
                        inv_notes = _inv_violations(wb, book_before, wrote_sheets, tuple(inv_skip))
                        _tick('不変条件', t0)
                        inv_dirty = False
                        if inv_notes:
                            print("頼んでいない変化（done を受け付けません）:")
                            for n in inv_notes:
                                print("  - " + n)
                            done = False
                            gates['inv'] += 1
                            limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                            img = None
                            msg = pre + ("【結果】\n道具が「撃つ前の控え」と「いまのブック」を照らしました。"
                                   "依頼に無い変化があります:\n"
                                   + "\n".join("- " + n for n in inv_notes)
                                   + "\n戻せるもの（潰れた数式・減った条件付き書式・別のシートに書いた分）は"
                                     "元に戻してください。戻す手が無いもの・依頼を果たすために必要だったものは、"
                                     "report に「なぜそうなったか」を書いてから done にしてください。"
                                   + f"（残りの往復: {limit - turn} 回）\n")
                            continue
                        print("頼んでいない変化: なし")
                    # (1.6) 読み残し＝数式を読みに行ったのに、見ていない数式セルが残っている。
                    #       道具が数える（AI を使わない）。「エラーが無い」と「見ていない」が
                    #       同じ『なし』になるのを止める（2026-09-05・本物の 161 列の表で出た欠陥）
                    if read_formula_ranges and not coverage_checked:
                        coverage_checked = True
                        try:
                            _addrs = _sheet_formula_addrs(wb.Sheets(sheet))
                        except Exception:
                            _addrs = []
                        _uncov = _uncovered_formula_cells(
                            _addrs, [r for (s, r) in read_formula_ranges if s == sheet])
                        if _uncov:
                            print(f"読み残し（done を受け付けません）: 数式 {len(_addrs)} セルのうち "
                                  f"{len(_uncov)} セルを見ていません（例 "
                                  + " ".join(sorted(_uncov)[:5]) + "）")
                            done = False
                            gates['coverage'] += 1
                            limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                            img = None
                            msg = pre + ("【結果】\n道具が数えました。数式のあるセル "
                                   f"{len(_addrs)} 個のうち、まだ read していないセルが {len(_uncov)} 個あります"
                                   "（例: " + " ".join(sorted(_uncov)[:8]) + "）。\n"
                                   "見ていない範囲について「なし」「不整合はありません」と書かないでください。"
                                   "残りを read（\"formula\": true）で読んでから、あらためて報告してください。"
                                   f"（残りの往復: {limit - turn} 回）\n")
                            continue
                        print(f"読み残し: なし（数式 {len(_addrs)} セルを全部見ています）")
                    # (2) 採点（依頼文と現物を照らし、満たしていない点を挙げる）。1 回だけ。
                    #     2026-09-06: 同じ会話の続きではなく**別の会話**で聞く。前は自分の手順も
                    #     言い分も全部見た同じ相手が採点していた＝追認しやすい作りだった。
                    #     渡すのは「依頼文・道具が読み直した現物・見た目の画像」だけ。
                    # 一発で通った回は採点を省く（2026-09-11 午後）: 1 往復目で手が全部通り、仕上げ検査・中身の検査・
                    # 頼んでいない変化に指摘が無ければ、採点役（別プロセス・8 秒）を呼ばない。付けるなら --grade。
                    # この日の午後に採点役だけが捕まえた漏れ（日付の表示形式のまだら）は、整え直しの手が吸収した
                    clean_first = (turn == 1 and merged_now and bool(touched) and sum(gates.values()) == 0
                                   and not audit_notes and not c_block and not inv_notes)
                    if grade and not graded and clean_first and not grade_always:
                        grade_skipped = True
                        print("採点: 省きました（一往復目で手が全部通り、仕上げ検査・中身・頼んでいない変化に指摘なし"
                              "＝一発。採点まで付けるなら --grade）")
                    elif grade and not graded:
                        graded = True
                        t0 = time.time()
                        _extra = _grade_extra_sheets(wb, sheet, sheets_at_start, wrote_sheets)
                        if _extra:
                            print("採点に見せる別シート: " + "、".join(_extra))
                        gmat_full = _grade_materials(sheet, wb, _extra)
                        _tick('採点の材料', t0)
                        gmsg = _mask(_GRADE_PROMPT.format(request=request, materials=_clip_grade(gmat_full)))
                        t0 = time.time()
                        gimg, gaddr = (_shot_for_ai(sheet, wb) if (show_image and touched) else (None, None))
                        _tick('画像', t0)
                        if gimg:
                            gmsg += _view_note(gaddr)
                        print(f"\n===== 採点（別の会話で聞く{'・' + str(grade_ai) if grade_ai else ''}） =====")
                        gtext, gusage, gsec = (ask or _ask)(g_ai, g_model, g_key, [('user', gmsg, gimg)])
                        total['in'] += gusage['in']
                        total['out'] += gusage['out']
                        total['think'] += gusage.get('think', 0)
                        total['sent'] += len(gmsg)
                        total['recv'] += len(gtext)
                        total['sec'] += gsec
                        print(_bill_line(len(gmsg), gtext, gusage, gsec, g_model, head="請求書（採点）"))
                        log.write(json.dumps({'turn': turn_base + turn, 'grade': True, 'prompt': gmsg,
                                              'reply': gtext, 'usage': gusage, 'sec': round(gsec, 2),
                                              'grader': f"{g_ai}/{g_model}"},
                                             ensure_ascii=False) + chr(10))
                        unmet = _parse_unmet(gtext)
                        noticed_left = _parse_noticed(gtext)
                        # 現物と食い違う指摘は道具が落とす（印刷が既に満たされている・パスの無い写真）
                        unmet, contradicted = _unmet_drop_satisfied(unmet, wb, sheet, request)
                        if contradicted:
                            print(f"採点の指摘のうち現物と食い違うもの {len(contradicted)} 件は"
                                  "落としました（往復を使わない）:")
                            for c in contradicted:
                                print("  - " + c)
                        # 承認の言葉が無い走行で、人の値を変えないと満たせない指摘は plan に混ぜない（往復を失うだけ）
                        unmet, moved = _split_unmet_for_approval(unmet, _approved())
                        if moved:
                            print(f"採点の指摘のうち人の値を変えるもの {len(moved)} 件は、承認の言葉が無いので"
                                  "報告の気づきへ回します（往復は使わない）: " + " / ".join(moved))
                            noticed_left = noticed_left + [m for m in moved
                                                           if not any(_plan_same(m, n) for n in noticed_left)]
                        if noticed_left:
                            print("採点の気づき（依頼には無い＝往復は使わず報告に写す）: " + " / ".join(noticed_left))
                        if unmet:
                            print("採点: 満たしていない点が {} 件（done を受け付けません）:".format(len(unmet)))
                            for u in unmet:
                                print("  - " + u)
                            unmet_left = list(unmet)
                            plan = _merge_plan(plan, [{'item': u, 'state': '未'} for u in unmet])
                            done = False
                            gates['grade'] += 1
                            limit = min(limit + 1, max_turns + _GATE_EXTRA_TURNS)
                            img = None
                            msg = pre + ("【結果】\nあなたの採点で、まだ満たしていない点が挙がりました:\n"
                                   + "\n".join("- " + u for u in unmet)
                                   + "\nこの分を直す actions を返してください。直せないものは plan の state を"
                                     "「不可」にして report に理由を書いてください。"
                                   + ("" if _approved() else
                                      "依頼文に承認の言葉が無いので、人の値を消す・書き換える手（上書き・行や列の削除・"
                                      "置換・元の列への書き戻し）は道具が実行しません。そういう点は actions を出さずに、"
                                      "plan の state を「要判断」にして ask に人への質問を書いてください。")
                                   + f"（残りの往復: {limit - turn} 回）\n")
                            continue
                        print("採点: 満たしていない点はありません")
                        grade_keep = (gmat_full, gimg)     # この後は書かない＝終わりの検査はこれを使い回す
                break
            img = None
            gates['empty'] += 1
            msg = ("【結果】\nactions が空で done でもありません。続きの actions を返すか、終わりなら "
                   f"done を true にして report を書いてください。（残りの往復: {limit - turn} 回）\n")
    finally:
        log.close()
        if ai == _CC_AI and not ask:
            try:
                cc_close_used()            # 使った会話は閉じ、予備（空）は次の走行のために残す（2026-09-11 午後）
            except Exception:
                pass
    if did and '【やったこと】' not in (report or ''):
        # 【やったこと】は道具が通った手の名札から作る（AI に書かせると 300〜400 トークン＝5 秒・2026-09-11 午後）
        made_line = "【やったこと】（道具が実行した手から）\n" + "\n".join("・" + d for d in did)
        report = made_line + ("\n\n" + report.strip() if (report or '').strip() else "")
    if (not done) and wanted_done and last_actions and not dry_run:
        # 最後の往復で「手」と done を一緒に返した＝手は実行済み。前は「AI は done を言っていません」で
        # 失敗にしていたが、仕事は済んでいる。落とし物が無いときだけ done にし、確かめは道具の検査に任せる
        pending = _plan_pending(plan)
        if pending:
            dropped = pending
        else:
            done = True
            print("\n（最後の往復で手と done を一緒に返しました。手は実行済みで、AI はその結果を見ていません"
                  "＝下の検査と「変わったセル」で確かめてください）")
    print()
    if plan:
        print(_plan_line(plan))
    if done:
        print("AI の報告:\n" + (report.strip() or "（なし）"))
    elif dry_run:
        pass
    elif dropped:
        print("落とし物: 未の項目が残ったまま往復が尽きました（{} 件）: ".format(len(dropped))
              + " / ".join(dropped) + "\n  → --max-turns を増やすか、agent --continue \"続けて\" で続きを回してください。")
        print("AI の報告:\n" + (report.strip() or "（なし）"))
    else:
        print(f"往復の上限（{limit} 回）に達しました。AI は done を言っていません"
              + ("（最後の actions は実行済み）" if last_actions else "") + "。")
    ok = True
    elapsed = job_clock_elapsed()                 # 検査の materials が時計を押し直す前に読む
    formula_rate = None                # 書いた数のうち数式で書いた割合（2026-09-06・派生値の数式率）
    # 変わったセルの明細（2026-09-05）: 番地・前・後を 1 行 1 セルで出し、全部をファイルに残す。
    #   数を数えるだけでは「戻せるから大丈夫」にしかならない。**戻さずに検分できる**ようにする。
    if not dry_run:
        t0 = time.time()
        changed_rows, fmt_rows, unmeasured = [], [], False
        if before is not None:
            rows = _changes_of(before, wb.Sheets(sheet), sheet)
            if rows is None:
                unmeasured = True
            else:
                changed_rows.extend(rows)
            try:
                fmt_rows.extend(_format_changes(before.get('format'), wb.Sheets(sheet)))
            except Exception:
                pass
        for nm, snap in (others or {}).items():
            try:
                rows = _changes_of(snap, wb.Sheets(nm), nm)
            except Exception:
                continue
            if rows:
                changed_rows.extend(rows)
            try:
                fmt_rows.extend(_format_changes((snap or {}).get('format'), wb.Sheets(nm)))
            except Exception:
                pass
        if unmeasured and not changed_rows and not fmt_rows:
            print("変わったセル: 数えていません（表が大きい）")
        elif before is not None or changed_rows or fmt_rows:
            path = _write_changes_file(changed_rows, fmt_rows=fmt_rows) if (changed_rows or fmt_rows) else None
            print(f"変わったセル: {len(changed_rows)} 個" + (f"   明細: {path}" if path else ""))
            multi = len({r['sheet'] for r in changed_rows}) > 1
            for line in _changes_table(changed_rows, multi_sheet=multi):
                print(line)
            if unmeasured:
                print("  （対象シートは大きすぎて数えていません。上は別シートのぶんだけです）")
            if fmt_rows:
                print(f"書式の変化: {len(fmt_rows)} 件（値の明細には出ないぶん）")
                for line in _format_table(fmt_rows):
                    print(line)
        # 派生値の数式率（2026-09-06・Claude for Excel の助言「100% でないものは後で壊れる」）。
        # 書いた数のうち、いくつを数式で書いたか。合計・割合・突き合わせを自分で計算して値で置くと、
        # 人が後から元の行を直したときに黙って古いままになる。関所にはしない（数を置く仕事もある）
        _tick('差分', t0)
        formula_rate = _formula_rate(changed_rows)
        if formula_rate is not None:
            n_num, n_f = formula_rate
            print(f"派生値の数式率: 書いた数 {n_num} 個のうち数式 {n_f} 個"
                  + (f"（{100.0 * n_f / n_num:.0f}%）" if n_num else "")
                  + ("　← 合計・割合・突き合わせを値で置いていないか確かめてください"
                     if n_num and n_f < n_num else ""))
    if touched and not dry_run and (audit_dirty or not audited):
        # 終わりの姿をここで見直す（関所で差し戻した後に直っていれば「指摘なし」になる。
        # 前は関所で拾った古い指摘をそのまま報告に出していた＝直った後も「指摘 N 件」と言っていた）。
        # 関所の後に書いていなければ読み直さない（同じ表を 2 回検査していた・2026-09-06 夜）
        t0 = time.time()
        audit_notes = _audit_now(wb, sheet, touched)
        c_block, content_noticed = _content_now(wb, sheet, touched)
        r_block, r_noticed = _request_gate(wb, sheet, book_before, request, _RUN_FACTS)
        c_block = c_block + r_block
        content_noticed = content_noticed + r_noticed
        audit_notes = audit_notes + c_block
        _tick('仕上げ検査', t0)
    if audit_notes:
        print(f"仕上げ検査: 指摘 {len(audit_notes)} 件")
        for n in audit_notes:
            print("  - " + n)
    elif touched and not dry_run:
        print("仕上げ検査: 指摘なし（見出し・罫線・列の型・列幅・エラー値・合計・式）")
    if content_noticed:
        print("中身の気づき（依頼には無い＝報告に写す）: " + " / ".join(content_noticed))
        noticed_left = noticed_left + [c for c in content_noticed
                                       if not any(_plan_same(c, n) for n in noticed_left)]
    if book_before is not None and not dry_run:
        # 関所の後にまた書いた／関所で引っかかったときだけ照らし直す（読むだけで数秒かかるブックがある）
        if inv_dirty or inv_notes:
            t0 = time.time()
            inv_notes = _inv_violations(wb, book_before, wrote_sheets, tuple(inv_skip))
            _tick('不変条件', t0)
        if inv_notes:
            print(f"頼んでいない変化: {len(inv_notes)} 件（依頼に無い変化です。--undo で戻せます）")
            for n in inv_notes:
                print("  - " + n)
        else:
            print("頼んでいない変化: なし（他シート・名前定義・隠れ行・元の数式・条件付き書式・入力規則・"
                  "結合・テーブル・グラフ・使用範囲・シートの並び・計算モード・枠固定・グラフの参照元・循環参照）")
    # 報告に道具が足す分（2026-09-06）。Claude for Excel の報告の型は 4 段（やったこと／できなかったこと／
    # 決めたこと／確認していないこと）で、こちらは 2 段だった。3 段目・4 段目を AI の作文に任せると
    # 忘れられる（忘れたことは報告に出ない＝読む側は気づけない）ので、**道具が実物から足す**。
    report_extras = []
    asks = _plan_asks(plan)
    if noticed_left and plan:
        # 同じ 3 件を【人に判断してほしいこと】と【気づいたこと】で二重に言わない（初実射の報告＝三重・2026-09-06 夜）
        said = [p['item'] for p in plan if p.get('state') == '要判断'] + asks
        noticed_left = [n for n in noticed_left if not any(_plan_same(n, a) for a in said)]
    if asks:
        report_extras.append("【人に判断してほしいこと】（安全側で仮に処理しました。"
                             "答えるときは agent --cont \"…\" で続きを回してください）\n"
                             + "\n".join("・" + a for a in asks))
    if noticed_left:
        report_extras.append("【気づいたこと】（依頼には無いので触っていません。直すなら agent --cont \"…\" で"
                             "承認の言葉を添えて頼んでください）\n" + "\n".join("・" + n for n in noticed_left))
    if discovered:
        report_extras.append("【現場を見て足した項目】（1 往復目の plan に無く、途中で足したもの）\n"
                             + "\n".join("・" + it + (f"　→ 理由: {w}" if w else "　→ 理由の記載なし")
                                         for it, w in discovered))
    if dropped:
        report_extras.append("【できなかったこと（往復が尽きました）】\n" + "\n".join("・" + d for d in dropped))
    if inv_notes:
        report_extras.append("【頼んでいない変化（道具が控えと照らして見つけました）】\n"
                             + "\n".join("・" + n for n in inv_notes))
    if report and not dry_run and before is not None and 'formula' not in inv_skip:
        # 「書いた」と言っている番地が空でないか（消す手を使った回は見送る＝消したのなら空で正しい）
        try:
            _ws = wb.Sheets(sheet)
            _empty = [a for a in _report_addrs(report)
                      if str(_ws.Range(a).Formula or '') == '']
        except Exception:
            _empty = []
        if _empty:
            report_extras.append("【報告に出た番地のうち、いま空のもの】（道具が読み直しました。"
                                 "書いたつもりで書けていないか、番地の書き間違いかもしれません。"
                                 "問い番号などを拾う誤検知もあります）\n・" + " ".join(_empty[:12]))
    if report_extras:
        print("\n道具が報告に足した分:")
        for blk in report_extras:
            print(blk)
        report = (report.rstrip() + "\n\n" + "\n\n".join(report_extras)) if report else "\n\n".join(report_extras)
    if verify and not dry_run:
        t0 = time.time()
        if grade_keep:
            # 採点の後に書けば grade_keep は None に戻る＝ここに残っているなら材料も画像も現物のまま（実射で
            # inv_dirty を条件に入れていて、backup=False の実射では一度も使い回されなかった・2026-09-06 夜）
            ok = _verify_sheet(sheet, wb, base, materials=grade_keep[0], png_src=grade_keep[1])
        else:
            ok = _verify_sheet(sheet, wb, base)
        _tick('終わりの検査', t0)
    if inv_notes:
        ok = False                      # 頼んでいない変化が残ったまま＝合格にしない
    if c_block:
        ok = False                      # 中身の指摘（依頼が満たされていない・行が消えた）が残ったまま＝合格にしない（2026-09-11）
        print("判定: 中身の指摘が残っています＝合格にしません")
    print(_gate_line(gates) + f"　＝ {_gate_grade(ok and done, gates)}")
    print(_bill_line(total['sent'], 'x' * total['recv'], total, total['sec'], model,
                     head=f"合計（往復 {turn} 回" + (f"・手と done を同時に {merged} 回" if merged else "") + "）"))
    if tool_sec:
        # 道具側の内訳（2026-09-06 夜: 全体の 3〜4 割が道具側なのに、請求書は AI の秒だけだった）
        print("道具側の内訳: " + " / ".join(f"{k} {v:.1f}" for k, v in tool_sec.items())
              + f" 秒（合計 {sum(tool_sec.values()):.1f}・AI 待ち {total['sec']:.1f}）")
    print(f"経過: materials から {elapsed:.1f} 秒" if elapsed is not None
          else f"経過: {time.time() - t_start:.1f} 秒")
    if done and not dry_run:
        try:
            if str(wb.Path or ''):          # 保存済みのブックだけ覚える（実射の使い捨てブックは覚えない）
                _notes_remember(book, sheet, request, report, _sheet_fingerprint(wb.Sheets(sheet)))
        except Exception:
            pass
    if not dry_run:
        _after_save(book, sheet, wb, before is not None)     # 次の走行で「人の手で変わった所」を出すため
    print(f"記録: {_LAST_AGENT_LOG_FILE}（往復の全文）")
    if not dry_run:
        print("（保存はしていません。Excelで確認後に保存してください。気に入らなければ保存せずに閉じる）")
    progress_write('done' if (done or dry_run) else 'fail', book=book, sheet=sheet,
                   turn=turn_base + turn, max_turns=limit, request=request[:80],   # 上限は関所の差し戻しで伸びる（「5/4」を出さない）
                   note=_report_head(report) if report else
                        ('落とし物あり' if dropped else '往復が尽きました'),
                   sec=elapsed if elapsed is not None else time.time() - t_start)
    r = {'done': done, 'ok': ok and (done or dry_run), 'verify': ok, 'turns': turn, 'report': report,
         'plan': plan, 'dropped': dropped, 'usage': total, 'gates': gates, 'run_id': run_id,
         'audit': audit_notes, 'unmet': unmet_left, 'graded': graded, 'grade_skipped': grade_skipped,
         'inv': inv_notes,
         'asks': asks, 'formula_rate': formula_rate, 'limit': limit, 'discovered': discovered,
         'merged': merged, 'tool_sec': {k: round(v, 2) for k, v in tool_sec.items()},
         'inv_skip': tuple(sorted(inv_skip)),      # 手から決めた「当然変わるもの」（実射の検査も同じ物差しで見る）
         # 意図して書いたシート（対象シート＋別シートに出す手＋既存ピボットの出力シート）。
         # 実射側は対象シート 1 枚しか知らず、ピボットを別シートに置く弾が全部落ちていた（2026-09-09）
         'inv_own': tuple(sorted(wrote_sheets)),
         'in_tokens': total['in'], 'out_tokens': total['out']}
    # 走行台帳（本番だけ）。実射（--fire）と build の直しループは backup=False で呼ぶ＝実戦の記録に混ぜない
    if backup and not dry_run:
        _runs_record(run_id, mode, book, sheet, request, r,
                     elapsed if elapsed is not None else time.time() - t_start, path=_book_path(wb))
    return r


# 実射（練習台・弾・答え合わせ・状態・写し取り・点数）は vbam_fire.py へ切り出した（2026-09-11）
# 不変条件の検査 は vbam_inv.py へ切り出した（2026-09-11）


# ----------------------------------------------------------------
# コマンド
# ----------------------------------------------------------------

# ----------------------------------------------------------------
# 薄い入口: 依頼文 1 つ → sheet / build / macro に振り分ける
# ----------------------------------------------------------------

_MODES = ('sheet', 'build', 'macro', 'both')
_MACRO_WORDS = re.compile(r'マクロ|プロシージャ|\bSub\b|\bFunction\b|VBA|コンパイル|実行時エラー|動かない', re.I)
# マクロの新規作成（直すのでなく）。「書いて」「足して」はセルの仕事にも使う語なので外した
# （「数式が動かない。合計を書いて」がマクロの新規作成に振られていた・2026-09-04）
_CREATE_WORDS = re.compile(r'作って|作成|追加して|新しく|新規|組んで|マクロを.{0,6}(書|足)')   # 「マクロを書いて」「マクロを 1 本足して」は作る（2026-09-04 夜・外しすぎを戻した）
# 「別シートの名簿と突き合わせて」（＝読むだけ）は build に振らない。書き先としての「別シートに」だけ拾う
_BUILD_WORDS = re.compile(r'新しいシート|新規シート|新規のシート|別の?シートに|シートを(作|追加|組|新|起こ)|シートに組|一枚組')
_HEAVY_WORDS = re.compile(r'ピボット|pivot|テーブル化|スライサー|slicer|パワークエリ|power ?query|データモデル|リレーション|メジャー|グラフ|チャート|chart', re.I)
_MODE_LABEL = {'sheet': "既存シート「{sheet}」を直す", 'build': "白紙／新しいシートを組む（build-sheet と同じ）",
               'macro': "マクロを修理する",
               'both': "シート「{sheet}」とマクロの両方を 1 つのループで直す"}


def _route(request, forced=None, macro=None, new_book=False, sheet_empty=False, macro_names=None):
    """入口の振り分け（純 Python）。--mode > --macro > --new-book > マクロの語 > 「新しいシート」の語か白紙 > sheet。

    macro_names（ブックのマクロ名の一覧）を渡すと、マクロの語があっても依頼文にどのマクロ名も無ければ sheet に落とす
    （「数式が動かない」が macro に振られ「直すマクロが分かりません」で止まっていた・2026-09-04）。None なら語だけで決める。
    """
    from vbam_macro import (_find_macro_in_request)   # vbam_macro は別ファイル（2026-09-11・遅延 import）
    if forced:
        f = str(forced).lower()
        if f not in _MODES:
            raise ValueError(f"--mode は {' | '.join(_MODES)}（{forced}）")
        return f
    if macro:
        return 'macro'
    if new_book:
        return 'build'
    if _MACRO_WORDS.search(request or ''):
        # 直す＝依頼文にそのブックのマクロ名がある。作る＝「マクロを作って」等の語がある（まだ名前が無くてよい）
        if macro_names is None or _find_macro_in_request(request, macro_names) or _CREATE_WORDS.search(request or ''):
            # シートの仕事も混ざっているなら、片方を落とさずに 1 つのループでまたぐ（2026-09-05）
            return 'both' if _BOTH_WORDS.search(request or '') else 'macro'
    if _HEAVY_WORDS.search(request or ''):
        return 'sheet'      # ピボット・テーブル等の手は sheet にしか無い（build に振ると空の表を組んで「合格」・2026-09-04）
    if _BUILD_WORDS.search(request or '') or sheet_empty:
        return 'build'
    return 'sheet'


_SHEET_VERBS = re.compile(r'整え|見やす|並べ替|色分け|合計|集計|突き合わ|列を|行を|セルを')

# シートの仕事とマクロの仕事が「2 つ並んでいる」と読める言い回しだけ（統合ループへ振り分ける判定）。
# _SHEET_VERBS より狭くする＝「A 列を消すマクロを作って」は列の話がマクロの中身なので macro のまま
# （2026-09-05・緩い判定で振り分けたら、マクロだけの依頼まで統合ループに流れた）。
_BOTH_WORDS = re.compile(r'(も|、|。|てから|た後|たあと|その後|そのあと|あわせて|併せて|同時に)'
                         r'[^。]{0,12}?(整え|見やす|並べ替|色分け|合計を|集計し|反映|表に書|シートに書)')


# build の設計図では組めない物（2026-09-09）。テーブル・グラフは含めない（組む engine が扱える／別の手で足せる）
_BUILD_CANNOT_RE = re.compile(r'ピボット|pivot|スライサー|slicer|パワークエリ|power ?query|データモデル|メジャー|DAX',
                              re.I)


def _compound_notice(mode, request, macro_name):
    """入口が 1 つに決まった依頼が、実は複数の仕事にまたがっていないかの一言（無ければ ''）。

    「マクロを直してから表も整えて」は macro だけに振られ、シート側の仕事が黙って落ちていた（2026-09-05）。
    黙って半分だけやるのではなく、残りは別の回で頼んでもらう案内を出す。
    """
    if mode == 'macro' and _SHEET_VERBS.search(request or ''):
        return ("この依頼にはシートの仕事も含まれています。この回はマクロだけを直します。"
                "シート側は続けて agent \"…\" --mode sheet で頼んでください。")
    if mode == 'sheet' and macro_name:
        return (f"マクロ名「{macro_name}」が依頼にありますが、この回はシートだけを直します。"
                f"マクロは agent --macro {macro_name} で頼んでください。")
    # build は設計図（列と型と仕上げ）で組む engine で、ピボット・スライサー・クエリ・モデルを組めない。
    # 2026-09-09: 「表を作って、そこからピボットも作って」が、表だけ組んで**何も言わずに**終わっていた
    # （【できなかったこと】も出ない）。macro↔sheet と同じ形で案内を出す。
    if mode == 'build':
        hit = _BUILD_CANNOT_RE.search(request or '')
        if hit:
            return (f"この依頼には「{hit.group(0)}」の仕事も含まれています。この回は表を組むだけです"
                    "（組む engine はピボット・スライサー・クエリ・データモデルを作れません）。"
                    "組み上がったら続けて agent \"…\" --mode sheet で頼んでください。")
    return ''


def _sheet_is_empty(ws):
    try:
        ur = ws.UsedRange
        return (int(ur.Rows.Count) == 1 and int(ur.Columns.Count) == 1 and ur.Cells(1, 1).Value is None
                and int(ws.Shapes.Count) == 0)
    except Exception:
        return False


_BUILD_FIX_TURNS = 2                 # 組み上げの検査に落ちたときの直しの往復（sheet のループに引き継ぐ）


def run_build(request, xl, wb, ai=_CC_AI, model=None, new_book=False, dry_run=False, fix=True,
              max_turns=None, book_is_new=False):
    """白紙／新しいシートを組む（build-sheet --ask と同じ流れを同じプロセスで）。

    検査（エラーセル・###）に落ちたら、そのまま sheet のループに渡して直させる（2026-09-04）。
    前は一発勝負で、落ちても「要確認」と言って終わりだった。
    直しの往復は max_turns（--max-turns）で決める。渡さなければ _BUILD_FIX_TURNS（2026-09-04: build が
    --max-turns を黙って無視していた）。
    book_is_new=True は「渡された wb がいま足したばかりの白紙」＝ここでは足さない（2026-09-05）。

    仕事の時計と進み具合は入口で押し直す（2026-09-05）。build は materials を読まない道があり、
    前の仕事の時計が残ったまま「経過: materials から 1030.2 秒」と出ていた。進み具合も
    run_agent／run_macro_agent にしか無く、build の最中に agent_status を見ると前の仕事の行が出た。
    """
    book = str(getattr(wb, 'Name', '') or '')
    t_start = time.time()          # 進み具合の秒はここから（直しのループが仕事の時計を押し直すため）
    run_id = _run_id()
    job_clock_start(f"{book} build")
    progress_write('start', book=book, request=(request or '')[:80],
                   note='組み上げを始めました', sec=0.0)
    materials = None if new_book else _book_materials(wb)
    progress_write('ask', book=book, request=(request or '')[:80],
                   note='設計図を聞いています', sec=time.time() - t_start)
    spec = ask_design(request, materials, ai, model)           # 請求書はここが出す
    plan = plan_sheet(spec)
    print(_plan_summary(plan))
    if dry_run:
        print("（--dry-run: Excel には触っていません）")
        progress_write('done', book=book, note='--dry-run（Excel には触っていません）',
                       sec=time.time() - t_start)
        return True
    if new_book and not book_is_new:
        wb = xl.Workbooks.Add()
        print(f"新しいブックに組み上げます: {wb.Name}")
        book = str(getattr(wb, 'Name', '') or book)
    # 既に中身のあるシートへ組むときは、AI が overwrite と言っても人の承認が要る（2026-09-04）。
    # ここは _apply が Cells.Clear する道＝控えも無しに人の表が消えていた。
    if not new_book and plan.get('overwrite'):
        from vbam_build import _find_sheet, _sheet_has_content
        try:
            ws_t = _find_sheet(wb, plan['sheet'])
            hit = ws_t is not None and _sheet_has_content(ws_t)
        except Exception:
            ws_t, hit = None, False              # 読めないブック（試験の作り物など）は素通し
        if hit:
            print(f"エラー: AI が既存シート '{plan['sheet']}' の中身を全部消して組み直そうとしました（overwrite）。"
                  "\n  中身のあるシートを潰す判断は人がします。消してよいなら CLI の "
                  "`build-sheet --ask --overwrite`、残すなら依頼文で別のシート名を指定してください。"
                  "\n  （このシートの一部を直すだけなら agent --mode sheet で回してください）")
            progress_write('fail', book=book, sheet=plan.get('sheet'),
                           note='既存シートの中身を消す組み直しは人の承認が要ります',
                           sec=time.time() - t_start)
            return False
    # 既存ブックに組むときは、書く前に控えを取る（--undo で戻せるように。新しいブックは戻す物が無い）
    if not new_book:
        from vbam_build import _find_sheet
        try:
            ws_b = _find_sheet(wb, plan['sheet'])
        except Exception:
            ws_b = None
        if ws_b is not None:
            _agent_backup(wb, ws_b, str(ws_b.Name), request)
        else:
            _mark_undo_stale('build（新しいシートを組んだ＝戻す元が無い）', book)
    else:
        _mark_undo_stale('build --new-book（新しいブックに組んだ）', book)
    progress_write('run', book=book, sheet=plan.get('sheet'),
                   note=f"組み上げています（{plan.get('sheet')}）", sec=time.time() - t_start)
    ws = _apply(xl, wb, plan, plan['overwrite'])
    if ws is None:
        progress_write('fail', book=book, sheet=plan.get('sheet'),
                       note='組み上げに失敗しました', sec=time.time() - t_start)
        return False
    ok = _verify_build(xl, ws, plan)
    fix_r = None
    if not ok and fix:
        fix_turns = int(max_turns or _BUILD_FIX_TURNS)
        print(f"\n組み上げの検査に落ちたので、直しをシートのループに渡します（往復 {fix_turns} 回まで）")
        try:
            r = run_agent("いま組み上げたこのシートの検査で、エラーセルか ### が出ました。材料の「エラーセル」「'###' で"
                          "読めないセル」を見て直してください（### は列幅を広げる。数式のエラーは参照や引数を直す）。"
                          "組み上げた形（見出し・合計行・書式）は崩さない。元の依頼: " + (request or '').strip(),
                          str(ws.Name), wb, ai, model, max_turns=fix_turns, backup=False, base=(0, 0))
            ok = bool(r.get('verify'))
            fix_r = r
        except Exception as ex:
            print(f"⚠ 直しのループを回せませんでした: {ex}")
    # 走行台帳（2026-09-06）。組み上げは 1 発勝負なので往復は 1、直しに回ったぶんは足す
    _runs_record(run_id, 'build', book, plan.get('sheet'), request,
                 {'turns': 1 + int((fix_r or {}).get('turns') or 0), 'done': ok, 'ok': ok,
                  'gates': (fix_r or {}).get('gates') or {},
                  'usage': (fix_r or {}).get('usage') or {}},
                 time.time() - t_start, save_log=bool(fix_r), path=_book_path(wb))
    print("（保存はしていません。Excelで確認後に保存してください）")
    progress_write('done' if ok else 'fail', book=book, sheet=plan.get('sheet'),
                   note=('組み上げました' if ok else '検査に落ちたまま終わりました'),
                   sec=time.time() - t_start)
    return ok


def _agent_workbook(target_file, new_book):
    """agent の対象ブックを掴む。返り値は (xl, wb, いま足した白紙か)。掴めなければ wb が None。

    --new-book（まっさらに組む）は元のブックが要らないのに、開いているブックが 1 冊も無いと
    get_workbook の関所で断られていた（2026-09-05 の実機確認で判明＝白紙の Excel から始められない）。
    その組み合わせのときだけ、起動中の Excel に白紙を 1 冊足して続ける。
    Excel 自体が起動していないときは今までどおり断る（勝手に起こすとゾンビが残る）。
    """
    try:
        xl, wb = get_workbook(target_file)
        return xl, wb, False
    except Exception as ex:
        if not new_book or target_file:
            print(f"エラー: {ex}")
            return None, None, False
    try:
        xl = _get_active_excel()
        xl.Visible = True
        wb = xl.Workbooks.Add()
    except Exception as ex2:
        print(f"エラー: {ex2}")
        return None, None, False
    print(f"開いているブックが無いので、白紙を 1 冊足しました: {wb.Name}")
    return xl, wb, True


def cmd_agent(args):
    """agent の入口。同時実行の関所（記録が混ざらないように）を掛けてから本体へ。

    --recipes（一覧を出すだけ）と --undo（控えで置き換えるだけ）は記録を書かないので錠を取らない
    （走っている agent がいても一覧が見られる・戻せる）。--backups（控えの一覧）も同じ。
    """
    try:
        if (getattr(args, 'recipes', False) or getattr(args, 'undo', False)
                or getattr(args, 'score', False) or getattr(args, 'changes', False)
                or getattr(args, 'backups', False)
                or getattr(args, 'mend', False)):      # 修理の試験は子プロセスが状態を置き場へ逃がす＝本番の記録に触らない
            return _cmd_agent_body(args)
        with _agent_lock():
            return _cmd_agent_body(args)
    finally:
        _cc_close_all()            # 温めていた claude-code を残さない（MCP サーバーは長生き＝放すと居残る・2026-09-07）


def _cmd_agent_replay(args, target_file, rest, src):
    """agent --replay: 記録の reply をそのまま再生してループだけを回す（API 不要・課金なし・開発用）。

    run_agent は記録ファイルを 'w' で開き直す＝再生元をそのまま渡すと消える。先に控えを取ってから読む。
    """
    import shutil
    if not os.path.isfile(src):
        print(f"エラー: 再生する記録がありません（{src}）")
        return False
    _, meta, _ = _load_resume(src)
    if not meta:
        print(f"エラー: 記録に meta がありません（{src}）")
        return False
    if meta.get('mode') != 'sheet':
        print(f"エラー: 再生できるのは sheet の記録だけです（この記録は「{meta.get('mode') or '不明'}」）")
        return False
    request = " ".join(rest).strip() or str(meta.get('request') or '')
    if not request:
        print("エラー: 記録に依頼文がなく、コマンド側にも依頼文がありません")
        return False
    try:
        xl, wb = get_workbook(target_file)
    except Exception as ex:
        print(f"エラー: {ex}")
        return False
    sheet = getattr(args, 'sheet_opt', None) or meta.get('sheet') or wb.ActiveSheet.Name
    if sheet not in [sh.Name for sh in wb.Sheets]:
        print(f"エラー: シート '{sheet}' が見つかりません")
        return False
    replay_src = os.path.join(SCRIPT_DIR, '_last_agent_replay_src.jsonl')
    shutil.copyfile(src, replay_src)
    print(f"再生: {replay_src}（{src} の控え）の reply をそのまま返します。API は呼びません")
    try:
        max_turns = int(getattr(args, 'max_turns', None) or meta.get('max_turns') or _DEFAULT_MAX_TURNS)
        r = run_agent(request, sheet, wb, meta.get('ai') or 'gemini', meta.get('model'), max_turns,
                      bool(getattr(args, 'dry_run', False)),
                      show_image=bool(getattr(args, 'image', False)),
                      grade=not bool(getattr(args, 'no_grade', False)),
                      grade_always=bool(getattr(args, 'grade', False)),
                      ask=_replay_ask(replay_src))
    except (RuntimeError, ValueError) as ex:
        print(f"エラー: {ex}")
        return False
    return r['ok']


def _cmd_agent_body(args):
    """依頼文 1 つを道具が回す薄い入口:
    agent [excel_file] "依頼文" [--sheet 名] [--mode sheet|build|macro] [--macro 名] [--new-book]
                         [--ai claude-code|gemini|claude] [--model 名] [--max-turns N] [--dry-run]
    agent --fire [名前 ...]      自動実射（まっさらなブックに練習台を組み、依頼を撃って答え合わせ）
    """
    from vbam_macro import (_find_macro_in_request, run_both, run_macro_agent)   # vbam_macro は別ファイル（2026-09-11・遅延 import）
    from vbam_fire import (fire_agent, keep_case, drop_case, cases_list, score_history,   # 実射は別ファイル（2026-09-11）
                           harvest_book, _HARVEST_KEEP_ROWS)
    target_file, rest = parse_target_and_rest(args.posargs)
    ai = getattr(args, 'ai', None)
    model = getattr(args, 'model', None)
    if getattr(args, 'recipes', False):
        for n in SHEET_RECIPES:
            print(f"{n}　{recipe_when(n)}")
        print(f"（{len(SHEET_RECIPES)} 本。使い方: agent --recipe 名前 [補足]。補足＝承認の言葉（置き換えてよい・埋めてよい・"
              "切ってよい・並べてよい）や番地・仕様。自動実射: agent --fire recipe [名前 ...]）")
        return True
    if getattr(args, 'undo', False):
        which = args.undo if isinstance(args.undo, str) else None     # --undo 番号／名札（無指定＝直前）
        try:
            return undo_agent(target_file, bool(getattr(args, 'dry_run', False)),
                              bool(getattr(args, 'force', False)), which=which)
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    if getattr(args, 'prune_days', None):
        try:
            return backups_prune(int(args.prune_days), bool(getattr(args, 'force', False)))
        except (TypeError, ValueError):
            print("エラー: --prune-days は日数（整数）で指定してください")
            return False
    if getattr(args, 'backups', False):
        return backups_list()
    if getattr(args, 'drop_case', None):
        return drop_case(args.drop_case)
    if getattr(args, 'changes', False):
        return show_changes()
    if getattr(args, 'score', False):
        return score_history()
    if getattr(args, 'runs', False):
        return runs_history()
    if getattr(args, 'cases', False):
        return cases_list()
    if getattr(args, 'forged', False):
        from vbam_forge import forged_list          # 鍛える回路（2026-09-17）
        return forged_list()
    if getattr(args, 'mend', False):
        from vbam_mend import mend_exam          # 修理の試験（鍛えたマクロを種ごとに壊して直させる・2026-09-17 夜）
        try:
            return mend_exam(getattr(args, 'seed', None), getattr(args, 'only', None), getattr(args, 'kinds', None),
                             bool(getattr(args, 'dry_run', False)), ai=ai, model=model,
                             max_turns=int(getattr(args, 'max_turns', None) or 4),
                             redo=getattr(args, 'mend_from', None), blind=bool(getattr(args, 'blind', False)))
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    if getattr(args, 'forge', None):
        from vbam_forge import forge
        # --test は nargs="+"＝後ろに書いた依頼文まで飲み込む。ブックに見えないものは依頼文へ戻す
        _tests = list(getattr(args, 'forge_tests', None) or [])
        _book_re = re.compile(r'\.(xlsx|xlsm|xlsb|xls)$', re.I)
        rest = list(rest) + [t for t in _tests if not _book_re.search(str(t))]
        args.forge_tests = [t for t in _tests if _book_re.search(str(t))]
        # 会話している AI が自分でマクロを書くとき（2026-09-19・API の鍵は要らない）: --prompt で問いを書き出し、
        # --answer 答えのファイル で 1 往復ぶん採点する（採点は道具）。不合格なら問いのファイルが次の問いに書き換わる
        answer = getattr(args, 'forge_answer', None)
        extra = {}
        if answer or getattr(args, 'forge_prompt', False):
            import vbam_forge as _vf
            os.makedirs(_vf._AGENT_FORGE_DIR, exist_ok=True)
            extra['prompt_out'] = os.path.join(_vf._AGENT_FORGE_DIR, _vf._safe_name(args.forge) + '_prompt.txt')
            if answer:
                if args.forge not in _vf._forge_load():
                    print(f"エラー: 台帳に「{args.forge}」がありません。先に agent --forge {args.forge} --prompt"
                          "（--before・--truth・依頼文つき）で問いを書き出してください")
                    return False
                if not os.path.isfile(answer):
                    print(f"エラー: 答えのファイルがありません: {answer}")
                    return False
                ai, model = 'file', os.path.abspath(answer)
            else:
                extra['prompt_only'] = True
        try:
            ok = forge(args.forge, target_file, ai=ai, model=model,
                       max_turns=1 if answer else int(getattr(args, 'max_turns', None) or 3),
                       register=bool(getattr(args, 'register', False)),
                       dry_run=bool(getattr(args, 'dry_run', False)),
                       truth=getattr(args, 'truth', None), before=getattr(args, 'forge_before', None),
                       tests=getattr(args, 'forge_tests', None), sheet=getattr(args, 'sheet_opt', None),
                       request=" ".join(rest).strip() or None, to=getattr(args, 'register_to', None),
                       phrases=getattr(args, 'forge_phrases', None), **extra)
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
        if ok is None:
            print(f"問いを読んでマクロ（Sub 全文）を書いたら: agent --forge {args.forge} --answer 答えのファイル")
            return True
        if ok and extra:
            _vf._write_prompt(extra['prompt_out'], "（合格しました。次の問いはありません）\n")
        return ok
    if getattr(args, 'keep_case', None):
        try:
            return keep_case(args.keep_case, target_file, getattr(args, 'seed', None))
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    if getattr(args, 'harvest', None):
        # 出力パス（"agent --harvest 元.xlsm 出力.xlsm" の "出力.xlsm"）は posargs の先頭で、
        # .xlsm/.xlsx に見える先頭引数は parse_target_and_rest が target_file 側へ取ってしまう
        # （harvest には「触る対象ブック」という概念が無いのに、この関数だけがそれを前提にしている）。
        # rest 側が空でも target_file を出力パスとして拾う（2026-09-05）
        keep = getattr(args, 'keep_rows', None)
        seed = getattr(args, 'seed', None)
        out = rest[0] if rest else target_file
        return bool(harvest_book(args.harvest, out,
                                 int(keep) if keep else _HARVEST_KEEP_ROWS,
                                 int(seed) if seed else None))
    if getattr(args, 'shake', None) is not None:
        from vbam_shake import shake        # 揺らし（記録の返事を書き換えて AI なしで再生）は別ファイル（2026-09-11 夜）
        try:
            return shake(args.shake or None, target_file, bool(getattr(args, 'dry_run', False)),
                         getattr(args, 'only', None), imagine=int(getattr(args, 'imagine', 0) or 0),
                         ai=ai, model=model)
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    if getattr(args, 'exam', False):
        from vbam_exam import exam          # 試験（正解の表を持つお題を種ごとに作って撃つ）は別ファイル（2026-09-11 夜）
        try:
            return exam(getattr(args, 'seed', None), getattr(args, 'only', None),
                        bool(getattr(args, 'dry_run', False)), ai=ai, model=model,
                        grade_always=bool(getattr(args, 'grade', False)),
                        macro_bas=getattr(args, 'by_macro', None), with_macro=getattr(args, 'with_macro', None))
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    if getattr(args, 'fire', False):
        try:
            if target_file:
                xl, _wb = get_workbook(target_file)
            else:
                from vbam_core import get_or_start_excel
                xl = get_or_start_excel()          # 練習台は自分で作る＝人のブックは要らない（2026-09-06 深夜）
            seed = getattr(args, 'seed', None)
            rep = getattr(args, 'repeat', None)
            # 2026-09-09: 実射の既定を claude-code にした。9/7 に対話側だけ claude-code へ変え、実射は
            # 「大量反復で Max のレート制限に当たらない」を理由に gemini（従量課金）のまま残していた。
            # その 1 行が残り続けた結果、9/4〜9/9 に実射 108 回・のべ 2,488 発が従量課金へ流れ、
            # 残高切れ（HTTP 429）で 83 本が全滅した。**金のかかる頭は、名指ししたときだけ使う。**
            return fire_agent(xl, ai or _CC_AI, model, rest or None, getattr(args, 'state', None),
                              int(seed) if seed else None, getattr(args, 'twice', False),
                              getattr(args, 'bed', None), repeat=int(rep) if rep else 1,
                              show_image=bool(getattr(args, 'image', False)),
                              grade_ai=getattr(args, 'grade_ai', None), grade_model=getattr(args, 'grade_model', None))
        except Exception as ex:
            print(f"エラー: {ex}")
            return False
    replay_path = getattr(args, 'replay', None)
    if replay_path is not None:
        return _cmd_agent_replay(args, target_file, rest, replay_path or _LAST_AGENT_LOG_FILE)
    recipe = getattr(args, 'recipe', None)
    if recipe:
        if recipe not in SHEET_RECIPES:
            print(f"エラー: 無い手順書: {recipe}")
            print("  ある手順書: " + "、".join(SHEET_RECIPES))
            return False
        rest = [compose_recipe_request(recipe, " ".join(rest))]     # 依頼文の位置に書いた文は補足（承認・番地・仕様）
    resume_hist = resume_meta = resume_res = None
    approval_kw = {}
    if getattr(args, 'cont', False):
        resume_hist, resume_meta, resume_res = _load_resume()
        if not resume_hist:
            print(f"エラー: 続きにできる記録がありません（{_LAST_AGENT_LOG_FILE}）")
            return False
        mode_of = (resume_meta or {}).get('mode')
        if mode_of != 'sheet':          # meta の無い古い記録（None）も続きにしない
            print(f"エラー: 続きにできるのは sheet の往復だけです（前回は {mode_of or '不明'}）")
            return False
        rest = [" ".join(rest)]          # 後ろの文は補足＝承認の言葉（空でもよい＝「そのまま続けて」）
        # 承認の言葉は元の依頼文＋この答えで探す。答えは道具の問いへの返事なので「はい」も承認（2026-09-06 夜）
        approval_kw = {'approval_text': _approval_scope((resume_meta or {}).get('request')) + ' ' + rest[0],
                       'approval_answer': True}
    req_file = getattr(args, 'request_file', None)
    if req_file:
        # 注文をファイルで受ける（Excelコンボの python エンジンが使う。コマンドラインの引用符事故を避ける）
        if rest:
            print("エラー: --request-file と依頼文の同時指定はできません")
            return False
        try:
            with open(req_file, 'r', encoding='utf-8-sig') as f:
                rest = [f.read().strip()]
        except OSError as ex:
            print(f"エラー: 依頼文のファイルを読めません: {ex}")
            return False
        if not rest[0]:
            print("エラー: 依頼文のファイルが空です")
            return False
    if len(rest) != 1:
        print('使い方: agent "依頼文" [--sheet 名] [--mode sheet|build|macro] [--macro 名] [--new-book]')
        print('              [--ai claude-code|gemini|claude] [--model 名] [--max-turns N] [--dry-run] [--request-file f]')
        print('        agent --recipe 手順書の名前 [補足]   （手順書を依頼文にして回す。--recipes で一覧）')
        print('        agent --fire [sheet|macro|recipe|名前 ...]   （自動実射）')
        return False
    request = rest[0]
    # --continue は前回と同じ往復数を既定にする（手順書の続きが 4 回に戻っていた・2026-09-04）
    default_turns = (resume_meta or {}).get('max_turns') if resume_hist else None
    default_turns = default_turns or (_RECIPE_MAX_TURNS if recipe else _DEFAULT_MAX_TURNS)
    try:
        max_turns = int(getattr(args, 'max_turns', None) or default_turns)
    except (TypeError, ValueError):
        print("エラー: --max-turns は数値で指定してください")
        return False
    if max_turns < 1:
        print("エラー: --max-turns は 1 以上で指定してください")
        return False
    dry_run = bool(getattr(args, 'dry_run', False))
    new_book = bool(getattr(args, 'new_book', False))
    macro = getattr(args, 'macro', None)
    xl, wb, book_is_new = _agent_workbook(target_file, new_book)
    if wb is None:
        return False
    sheet = getattr(args, 'sheet_opt', None) or wb.ActiveSheet.Name
    if resume_hist:
        want_book = (resume_meta or {}).get('book')
        if want_book and str(wb.Name) != want_book:
            print(f"エラー: 前回の続きは「{want_book}」のものです（いま開いているのは「{wb.Name}」）")
            return False
        if not getattr(args, 'sheet_opt', None) and (resume_meta or {}).get('sheet'):
            sheet = resume_meta['sheet']
    if sheet not in [sh.Name for sh in wb.Sheets]:
        print(f"エラー: シート '{sheet}' が見つかりません")
        return False
    # 手順書（--recipe）は本文の言い回し（「別のシート」「マクロ」）で build／macro に振られていた
    # （2026-09-04・25 本のうち「2つの表の突合」は必ず build、「図形と画像の整理」は macro）。
    # 手順書はシートを直す手順なので sheet に固定する。
    forced = 'sheet' if (resume_hist or recipe) else getattr(args, 'mode', None)
    names = None
    if not forced and not macro and not new_book and _MACRO_WORDS.search(request):
        try:
            names = _all_procedure_names(wb)        # マクロの語があるときだけ読む（名前が依頼文に無ければ sheet に落とす）
        except Exception:
            names = None                            # 読めなければ従来どおり語だけで決める
    try:
        mode = _route(request, forced, macro, new_book, _sheet_is_empty(wb.Sheets(sheet)), names)
    except ValueError as ex:
        print(f"エラー: {ex}")
        return False
    print(f"入口: {_MODE_LABEL[mode].format(sheet=sheet)}（--mode sheet|build|macro で指定できます）")
    if names is not None and mode != 'macro':
        print("（マクロの語はありますが、このブックのマクロ名が依頼文に無いのでマクロの修理には振りません。修理なら --macro 名）")
    notice = _compound_notice(mode, request, _find_macro_in_request(request, names) if names else None)
    if notice:
        print(f"（{notice}）")
    try:
        if mode == 'build':
            return run_build(request, xl, wb, ai, model, new_book, dry_run,
                             max_turns=(max_turns if getattr(args, 'max_turns', None) else None),
                             book_is_new=book_is_new)
        if mode == 'macro':
            # dict をそのまま返すと vba_manager の `ok is not False` が真＝往復が尽きても・最終コンパイルが
            # 不通過でも終了コード 0 で「成功」になっていた（sheet は r['ok'] を返している・2026-09-04）
            rm = run_macro_agent(request, xl, wb, macro, ai, model, max_turns, dry_run,
                                 rehearse=bool(getattr(args, 'rehearse', False)))
            return rm['ok']
        if mode == 'both':
            rb = run_both(request, sheet, xl, wb, macro, ai, model, max_turns, dry_run,
                          rehearse=bool(getattr(args, 'rehearse', False)),
                          show_image=bool(getattr(args, 'image', False)),
                          grade=not bool(getattr(args, 'no_grade', False)),
                          grade_ai=getattr(args, 'grade_ai', None),
                          grade_model=getattr(args, 'grade_model', None))
            return rb['ok']
        r = run_agent(request, sheet, wb, ai, model, max_turns, dry_run,
                      resume=resume_hist, resume_results=resume_res,
                      show_image=bool(getattr(args, 'image', False)),
                      grade=not bool(getattr(args, 'no_grade', False)),
                      grade_always=bool(getattr(args, 'grade', False)),
                      grade_ai=getattr(args, 'grade_ai', None),
                      grade_model=getattr(args, 'grade_model', None), **approval_kw)
    except (RuntimeError, ValueError, urllib.error.URLError) as ex:
        print(f"エラー: {ex}")
        return False
    return r['ok']


# macro は vbam_macro.py へ切り出した（2026-09-11）


# 道具が COM で直接やる手 は vbam_hands.py へ切り出した（2026-09-11）


def _normalize_fact(sheet, act, res):
    """normalize を当てた列と規則を覚える（終わりの検査で「規則で直らずに残った文字」を探す・2026-09-11）。"""
    try:
        if not res or not res[1]:
            return
        _r1, c1, _r2, c2 = _range_box(_check_addr(act['range']))
        if act.get('to'):
            _tr, tc = _a1_to_rc(_check_addr(act['to']))
            cols = range(tc, tc + c2 - c1 + 1)
        else:
            cols = range(c1, c2 + 1)
        names = tuple(r if isinstance(r, str) else next(iter(r), '') for r in (act.get('rules') or []))
        for c in cols:
            _RUN_FACTS['normalize'].append((str(sheet), _col_letter(c), names))
    except Exception:
        pass


def _direct_action(act, sheet, wb):
    """道具が COM で直接やる手。(表示名, ok, 出力) を返す。規約外は ValueError（実行せず理由を結果に）。"""
    op = str(act.get('op') or '')
    if wb is None:
        raise ValueError("ブックが無い（--dry-run では実行しない）")
    sheet = _act_sheet(act, sheet)      # "sheet" を書けばそのシートへ（2026-09-05）
    ws = wb.Sheets(sheet)
    if op == 'normalize':
        res = _do_normalize(ws, act)
        _normalize_fact(sheet, act, res)
        return res
    if op == 'dedupe':
        return _do_dedupe(ws, act)
    if op == 'fill':
        return _do_fill(ws, act)
    if op in ('inspect', 'read_sheet'):
        name = str(act.get('sheet') or sheet).strip()
        names = [sh.Name for sh in wb.Sheets]
        if name not in names:
            raise ValueError(f"シート '{name}' が無い（ある: {', '.join(names)}）")
        # materials は仕事の時計を押す。往復の途中の inspect で押し直すと、経過秒が「inspect から」になる
        clock = job_clock_get()
        ok, out = _run_cmd(['materials', name], wb)
        job_clock_set(clock)
        try:
            out += _extra_materials(wb, wb.Sheets(name))
        except Exception as ex:
            out += f"（追加の材料を読めませんでした: {ex}）\n"
        return f"read_sheet {name}", ok, out
    if op == 'eval':
        # 式を計算して結果を見るだけ（2026-09-06 夜・自由コードの逃げ道の読み取り専用版。セルには書かない）
        forms = act.get('formulas')
        if forms is None:
            forms = [act.get('formula')]
        if not isinstance(forms, (list, tuple)) or not forms or not all(str(f or '').strip() for f in forms):
            raise ValueError("eval には formula（か formulas の配列）が要ります")
        if len(forms) > 20:
            raise ValueError(f"eval は一度に 20 本まで（{len(forms)} 本）")
        out = []
        for f in forms:
            f = str(f).strip()
            try:
                v = ws.Evaluate(f if f.startswith('=') else '=' + f)
                out.append(f"{f} → {_eval_text(v)}")
            except Exception as ex:
                out.append(f"{f} → 計算できません（{ex}）")
        return f"eval {len(forms)}", True, "\n".join(out) + "\n（計算しただけ。セルには書いていません）"
    if op == 'image':
        # 写真・画像を置く（2026-09-06 夜・「写真も入れて」に応える手が無かった）
        path = str(act.get('path') or '').strip()
        if not path:
            raise ValueError("image には path（画像ファイル）が要ります。依頼にパスが無ければ「不可」にして report で聞く")
        if os.path.splitext(path)[1].lower() not in _IMAGE_EXTS:
            raise ValueError(f"image の path は {'/'.join(_IMAGE_EXTS)} の画像ファイルです（{os.path.basename(path)}）")
        if not os.path.isfile(path):
            raise ValueError(f"画像ファイルが見つかりません: {path}（人にパスを確かめる）")
        # Excel の AddPicture はスラッシュ区切りのパスを「指定したファイルが見つかりませんでした」で
        # 弾く（Python の isfile は通るので、実在するのに見つからないと言われる・2026-09-08 の実射
        # 「写真を入れる」がこれで落ちていた）。円記号区切りの絶対パスに直してから渡す
        path = os.path.abspath(path)
        anchor = ws.Range(_check_addr(act.get('at') or 'A1'))
        shp = ws.Shapes.AddPicture(path, 0, -1, float(anchor.Left), float(anchor.Top), -1, -1)
        shp.LockAspectRatio = -1
        if act.get('width'):
            shp.Width = float(act['width'])
        elif act.get('height'):
            shp.Height = float(act['height'])
        if str(act.get('name') or '').strip():
            shp.Name = str(act['name']).strip()
        return f"image {os.path.basename(path)}", True, (
            f"画像を置きました: {shp.Name}（{os.path.basename(path)}）  位置 {str(anchor.Address).replace('$', '')} の左上"
            f"  大きさ w={float(shp.Width):.0f} h={float(shp.Height):.0f}\n（消すなら shape の delete。--undo でも戻る）")
    if op in ('row_group', 'col_group'):
        # 隠さずに畳む（2026-09-06 夜・「行列は非表示にせずグループ化」。非表示にする手は無い）
        a, b = act.get('from'), act.get('to')
        if a in (None, '') or b in (None, ''):
            raise ValueError(f"{op} には from と to（{'行番号' if op == 'row_group' else '列文字'}）が要ります")
        if op == 'row_group':
            a, b = int(a), int(b)
            if a < 1 or b < a:
                raise ValueError("row_group の from・to は 1 以上で from ≦ to")
            target = ws.Rows(f"{a}:{b}")
        else:
            a, b = str(a).strip().upper(), str(b).strip().upper()
            if not (a.isascii() and a.isalpha() and b.isascii() and b.isalpha()):
                raise ValueError("col_group の from・to は列文字（A〜XFD）です")
            target = ws.Columns(f"{a}:{b}")
        try:
            if act.get('off'):
                target.Ungroup()
            else:
                target.Group()
        except Exception as ex:
            raise ValueError(f"{op} を当てられません（{'解除する所がグループ化されていない' if act.get('off') else ex}）")
        return f"{op} {a}:{b}", True, (f"{'グループ化を解除' if act.get('off') else 'グループ化'}: {a}:{b}"
                                        "（隠してはいません。左上の [-] で畳めます）")
    if op == 'view':
        fr, zoom = act.get('freeze'), act.get('zoom')
        if not fr and not zoom:
            raise ValueError("view には freeze か zoom が要ります")
        ws.Activate()
        win = wb.Application.ActiveWindow
        out = []
        if fr:
            win.FreezePanes = False
            if str(fr).lower() != 'off':
                a = _check_addr(fr)
                ws.Range(a).Select()
                win.FreezePanes = True
                ws.Range("A1").Select()
                # 言いっぱなしにせず実物を読み戻す（固定できていないのに「固定した」と出ていた・2026-09-04）
                try:
                    sr, sc = int(win.SplitRow), int(win.SplitColumn)
                    out.append(f"枠固定: {a} の左上で固定（上 {sr} 行・左 {sc} 列を固定）"
                               if (sr or sc) else f"⚠ 枠固定: {a} を指定しましたが固定されていません")
                except Exception:
                    out.append(f"枠固定: {a} の左上で固定")
            else:
                out.append("枠固定: 解除")
        if zoom:
            win.Zoom = int(zoom)
            out.append(f"表示倍率: {int(zoom)}%")
        return "view", True, "\n".join(out)
    if op == 'page_setup':
        ps = ws.PageSetup
        done = []
        if act.get('area'):
            ps.PrintArea = str(ws.Range(_check_addr(act['area'])).Address)
            done.append('area')
        if act.get('title_rows'):
            t = str(act['title_rows'])
            a, b = (t.split(':') + [t])[:2]
            ps.PrintTitleRows = f"${a}:${b}"
            done.append('title_rows')
        if 'landscape' in act:
            ps.Orientation = 2 if act['landscape'] else 1
            done.append('landscape' if act['landscape'] else 'portrait')
        if act.get('fit_wide') is not None or act.get('fit_tall') is not None:
            ps.Zoom = False
            fw, ft = act.get('fit_wide'), act.get('fit_tall')
            ps.FitToPagesWide = int(fw) if fw else False
            ps.FitToPagesTall = int(ft) if ft else False
            done.append('fit')
        if act.get('zoom'):
            ps.Zoom = int(act['zoom'])
            done.append('zoom')
        if act.get('center_h'):
            ps.CenterHorizontally = True
            done.append('center_h')
        if act.get('footer_page'):
            ps.CenterFooter = "&P / &N"
            done.append('footer_page')
        if not done:
            raise ValueError("page_setup に当てる項目がありません（area/title_rows/landscape/fit_wide/fit_tall/zoom/center_h/footer_page）")
        return "page_setup", True, _print_summary(ws)
    if op == 'chart':
        piv = str(act.get('pivot') or '').strip()
        ws_c = ws
        if piv:
            from vbam_heavy import _find_pivot
            psh, pt = _find_pivot(wb, piv)
            if pt is None:
                raise ValueError(f"ピボット '{piv}' が見つかりません（材料の「ピボット」の名前を使う）")
            ws_c, rng = psh, pt.TableRange1      # ピボットの出力範囲を元にすると Excel がピボットグラフにする
        else:
            rng = ws.Range(_check_addr(act.get('range')))
        ctype = str(act.get('type') or 'column').lower()
        if ctype not in _CHART_TYPES:
            raise ValueError(f"type は {'/'.join(_CHART_TYPES)}（{ctype}）")
        if act.get('at'):
            anchor = ws_c.Range(_check_addr(act['at']))
            left, top = float(anchor.Left), float(anchor.Top)
        else:
            left, top = float(rng.Left) + float(rng.Width) + 10, float(rng.Top)
        w, h = float(act.get('width') or 360), float(act.get('height') or 216)
        co = ws_c.ChartObjects().Add(left, top, w, h)
        ch = co.Chart
        ch.SetSourceData(rng)
        ch.ChartType = _CHART_TYPES[ctype]
        if act.get('title'):
            ch.HasTitle = True
            ch.ChartTitle.Text = str(act['title'])
        return f"chart {ctype}", True, (f"グラフ作成: [{ws_c.Name}] {co.Name}  種別 {ctype}  データ {str(rng.Address).replace('$', '')}"
                                        + ("（ピボットグラフ）" if piv else "")
                                        + f"  表題 {act.get('title') or '（なし）'}  位置 l={left:.0f} t={top:.0f} w={w:.0f} h={h:.0f}")
    if op == 'shape':
        delete = bool(act.get('delete'))
        names = act.get('names')
        # 「マクロの付いていない図形を全部」（2026-09-06 夜）。同じ依頼を 4 回撃って 3 回、AI が材料に無い
        # 図形名を書いて 1 往復目が丸ごと止まった（名前を写す工程そのものを無くす）。
        if delete and (act.get('all') or (isinstance(names, str) and names.strip().lower()
                                          in ('all', 'all_without_macro', '*'))):
            names = _unbound_shape_names(ws)
            if not names:
                return "shape delete 0", True, "マクロの付いていない図形はありません（消すものなし）"
        if names is None:
            names = [act.get('name')]
        elif not isinstance(names, (list, tuple)):
            raise ValueError("shape の names は名前の配列です（例 [\"矢印1\",\"ロゴ枠\"]）。"
                             "マクロ無しを全部消すなら \"names\":\"all_without_macro\"")
        names = [str(n or '').strip() for n in names]
        if not names or not all(names):
            raise ValueError("shape には name（消すなら names も可）が要ります")
        if len(names) > 1 and not delete:
            raise ValueError("names で複数まとめて扱えるのは delete のときだけ（位置と大きさは name で 1 つずつ）")
        picked = []
        for nm in names:
            shp = _pick_shape(ws, nm)
            if shp is None:
                cands = _unbound_shape_names(ws)
                raise ValueError(f"図形 '{nm}' が無い（材料の「図形・ボタン」に出ている名前を使う。"
                                 + (f"いまあるマクロ無しの図形: {'、'.join(cands[:12])}" if cands
                                    else "マクロ無しの図形はありません")
                                 + "。マクロ無しを全部消すなら \"names\":\"all_without_macro\"）")
            try:
                oa = str(shp.OnAction or '')
            except Exception:
                oa = ''
            if oa:
                raise ValueError(f"'{nm}' にはマクロ（{oa}）が割り当たっている＝"
                                 + ("消さない（ボタンを壊すため）" if delete else "動かさない"))
            picked.append((nm, shp))
        if delete:
            for nm, shp in picked:
                shp.Delete()
            try:
                left = int(ws.Shapes.Count)
            except Exception:
                left = -1
            return (f"shape delete {len(names)}", True,
                    f"図形を消しました（{len(names)} 個）: " + "、".join(names)
                    + (f"\n残っている図形: {left} 個" if left >= 0 else "")
                    + "\n（消した図形は控えに入っています＝ py vba_manager.py agent --undo で戻ります）")
        nm, shp = picked[0]
        keys = [k for k in ('left', 'top', 'width', 'height') if act.get(k) is not None]
        if not keys:
            raise ValueError("shape に当てる項目がありません（left/top/width/height。消すなら \"delete\": true）")
        for k in keys:
            setattr(shp, k.capitalize(), float(act[k]))
        return f"shape {nm}", True, (f"{nm}: l={float(shp.Left):.0f} t={float(shp.Top):.0f} "
                                     f"w={float(shp.Width):.0f} h={float(shp.Height):.0f}")
    if op == 'export_csv':
        rng = ws.Range(_check_addr(act.get('range')))
        path = _csv_check_path(act.get('path'))
        enc = str(act.get('encoding') or 'utf-8-sig').lower()
        if enc not in _CSV_ENCODINGS:
            raise ValueError(f"encoding は {'|'.join(_CSV_ENCODINGS)}（{enc}）")
        fmt = _date_fmt(act.get('date_format'))
        rows = [[_csv_text(v, fmt) for v in r] for r in _rows_of(rng.Value)]
        with open(path, 'w', encoding=enc, newline='') as f:
            csv.writer(f).writerows(rows)
        head = " / ".join(",".join(r) for r in rows[:2])
        tail = ",".join(rows[-1]) if len(rows) > 2 else ''
        return "export_csv", True, (f"CSV 書き出し: {path}（{len(rows)} 行 x {len(rows[0]) if rows else 0} 列・{enc}・日付 {fmt}）\n"
                                    f"先頭: {head}" + (f"\n末尾: {tail}" if tail else ""))
    if op == 'read_file':
        path, ext = _read_file_path(act.get('path'))
        limit = max(1, min(int(act.get('limit') or _READ_FILE_MAX_ROWS), 200000))
        want_sheet = str(act.get('from_sheet') or '').strip() or None
        opened = _open_book_by_path(wb, path) if wb is not None else None
        if opened is not None:
            # 同じ Excel でその人が開いたまま（保存前の直しも見える）＝ファイルからでなくブックから読む
            sh = opened.Sheets(want_sheet) if want_sheet else opened.Sheets(1)
            rows = _rows_of(sh.UsedRange.Value)[:limit]
            src = f"開いているブック {opened.Name}!{sh.Name}"
        elif ext in ('.xlsx', '.xlsm'):
            rows, name_in, names = _read_book_rows(path, want_sheet, limit)
            src = (f"{os.path.basename(path)}!{name_in}"
                   + (f"（ほかのシート: {'・'.join(n for n in names if n != name_in)}）" if len(names) > 1 else ''))
        else:
            rows, enc = _read_csv_rows(path, limit)
            src = f"{os.path.basename(path)}（{enc}）"
        if not rows:
            raise ValueError(f"中身が空でした: {path}")
        name = _new_sheet_name(wb, act.get('to'), path)
        ws_new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_new.Name = name
        n_rows, n_cols, text_cols = _read_file_write(ws_new, rows)
        head = " / ".join(",".join('' if v is None else str(v) for v in r[:6]) for r in rows[:2])
        return "read_file", True, (
            f"シート追加: {name}\n"
            f"読み込み: {src} → {name}!A1:{_col_letter(n_cols)}{n_rows}（{n_rows} 行 x {n_cols} 列）"
            + (f"　先頭ゼロを守った列: {'・'.join(text_cols)}" if text_cols else "")
            + f"\n先頭: {head}"
            + (f"\n（{limit} 行で切りました＝続きが要るなら limit を増やす）" if len(rows) >= limit else "")
            + "\n（元のファイルには書いていません。このブックも保存していません）")
    raise ValueError(f"許していない手です: {op!r}")


# APIキーの預かり は vbam_keys.py へ切り出した（2026-09-11）


# clean-table は vbam_clean.py へ切り出した（2026-09-11）


__all__ = ['cmd_agent', 'run_agent', 'run_build', 'RULES', # fire_agent・FIRE_CASES は vbam_fire（2026-09-11）
           'SHEET_RECIPES', 'RECIPE_CASES',
           '_snapshot_book', '_inv_violations', '_grid_diff', '_cell_name',
           'cmd_set_key', 'cmd_clear_key', '_mask_key', '_key_problem', '_KEY_ENV']
