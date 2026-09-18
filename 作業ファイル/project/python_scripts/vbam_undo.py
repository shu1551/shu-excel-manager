# -*- coding: utf-8 -*-
"""vbam_undo.py — vba_manager 分割パート: 控えと戻す手（書く前の控え・差分の明細・作った物の始末・--undo）

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。手の書き先（_act_sheet）・シートの控え
（_sheet_snapshot・_changes_of）・控えの覚書・戻す手（undo_agent）・図形の名前。
vbam_agent は `from vbam_undo import *` で名前を引き継ぐ。
"""
import os
import re
import json
import time
import contextlib
import csv
from vbam_core import (BACKUP_DIR, SCRIPT_DIR, _col_letter, get_workbook)
from vbam_hands import (_addr, _rows_of, _text_of)
from vbam_inv import (_a1_to_rc)
from vbam_ledger import (_runs_load, runs_append)

_HEAVY_OPS = ('table', 'pivot', 'pivot_field', 'pivot_calc', 'slicer', 'powerquery', 'datamodel')   # 重い道具（CLI をそのまま呼ぶ）


# ----------------------------------------------------------------
# 手3 相当: 終わったときの検査（エラーセル・###・画像）
# ----------------------------------------------------------------

# ----------------------------------------------------------------
# 取り消し（2026-09-04）: COM で書くと Excel の Ctrl+Z の履歴が消える。最初の書き込みの前に控えを取り、
#   `agent --undo` でそのシートだけ戻す。あわせて「変わったセル」を道具が数えて報告する
#   （AI の report を信じずに検分できる。前は「保存せずに閉じる」しか戻す道が無く、人の未保存の作業も一緒に捨てていた）。
# ----------------------------------------------------------------
_LAST_AGENT_UNDO_FILE = os.path.join(SCRIPT_DIR, '_last_agent_undo.json')
# 変わったセルの明細（2026-09-05）: 数だけでは検分できない＝番地・前・後を 1 行 1 セルで残す
_LAST_AGENT_CHANGES_FILE = os.path.join(SCRIPT_DIR, '_last_agent_changes.tsv')
_CHANGES_MAX_ROWS = 5000               # 明細のファイルはここで切る（白紙から組むと全セルが「新しい」になる）
_AGENT_BACKUP_KEEP = 5                 # 控えは新しい 5 個だけ残す
_AGENT_BACKUP_MARK = '_agent_before_'
_AGENT_BACKUP_META_EXT = '.json'       # 控えごとの覚書（控えのファイルの隣。--undo 番号 で 5 個のどれにでも戻せる・2026-09-06）
_SNAPSHOT_MAX_CELLS = 200000           # これを超える表は差分を数えない（読むだけで遅くなる）
# 読むだけの手（控えも見た目の画像も要らない）。view／page_setup は書く手のまま（印刷設定はシートの持ち物）
_READONLY_OPS = ('read', 'inspect', 'read_sheet', 'eval', 'find', 'export_csv', 'export_pdf')
_READONLY_HEAVY_ACTIONS = ('list', 'read', 'get_data', 'calc_field_list')


def _act_sheet(act, sheet):
    """その手が触るシート（"sheet" を書けばそのシート・書かなければ対象シート）。純 Python（2026-09-05）。

    重い道具（ピボット・パワークエリ）の "sheet" は「作った物の置き場」という別の意味なので、
    ここでは読み替えない（pivot create の "sheet" はこれから作るシートのこともある）。
    """
    if not isinstance(act, dict) or str(act.get('op') or '') in _HEAVY_OPS:
        return sheet
    return str(act.get('sheet') or '').strip() or sheet


def _act_writes(act):
    """その手 1 つがシートを変えるか（純 Python）。読むだけの手は False。"""
    if not isinstance(act, dict):
        return True                                  # 形が違う＝実行時に弾くが、書くものとして扱う
    op = str(act.get('op') or '')
    if op in _READONLY_OPS:
        return False
    if op in _HEAVY_OPS and str(act.get('action') or '').strip().lower().replace('-', '_') in _READONLY_HEAVY_ACTIONS:
        return False
    return True


def _will_write(actions):
    """この手の並びに「シートを変える手」が 1 つでもあるか（控えを取るかの判定・純 Python）。"""
    return any(_act_writes(a) for a in actions)


_EMPTY_SNAPSHOT = {'addr': None, 'row': 1, 'col': 1, 'values': [[None]], 'formulas': None, 'shapes': []}


def _other_sheets_in(actions):
    """この手の並びが対象シート以外に書く先のシート名（純 Python）。

    ピボットの置き場・パワークエリの読み込み先・sheet_op で足す／複写するシート。
    「変わったセル」を対象シートだけで数えていたので、別シートに出したときは 0 個と報告していた（2026-09-04）。
    書く手が "sheet" で名指しした別シートもここに入れる（2026-09-05・書ける先を広げた）。
    """
    names = set()
    for a in actions:
        if not isinstance(a, dict):
            continue
        op = str(a.get('op') or '')
        action = str(a.get('action') or '').strip().lower().replace('-', '_')
        if op not in _HEAVY_OPS and a.get('sheet') and _act_writes(a):
            names.add(str(a['sheet']).strip())
        if op == 'pivot' and action == 'create' and a.get('sheet'):
            names.add(str(a['sheet']))
        elif op == 'powerquery' and action == 'load' and a.get('sheet'):
            names.add(str(a['sheet']))
        elif op == 'sheet_op' and action == 'add' and a.get('name'):
            names.add(str(a['name']))
        elif op == 'sheet_op' and action == 'copy' and a.get('to'):
            names.add(str(a['to']))
    return names


def _pivot_sheets_in(actions, wb):
    """pivot_field / pivot_calc / slicer が名指ししたピボットが出ているシート名（2026-09-09）。

    既存のピボットを直す手は**置き場所を書かない**（ピボット名しか書かない）ので、_other_sheets_in では
    拾えなかった。拾えないと出力シートが「頼んでいないシート」のまま残り、ピボットを直せば必ず
    「頼んでいないシート「◯◯」が変わった」で落ちる——別シートにピボットを置いている人は、
    ピボットを直すたびに毎回これに当たる（2026-09-09 の実射で 12 本中 8 本がこれ一つで落ちた）。
    名指しされたピボットのシートだけを足す（他シート全部を検査から外すのではない）。
    """
    want = set()
    for a in (actions or []):
        if not isinstance(a, dict):
            continue
        if str(a.get('op') or '') not in ('pivot_field', 'pivot_calc', 'slicer'):
            continue
        for key in ('pivot', 'source'):
            v = a.get(key)
            if v:
                want.add(str(v).strip())
    if not want:
        return set()
    names = set()
    try:
        for sh in wb.Worksheets:
            for pt in sh.PivotTables():
                try:
                    if str(pt.Name) in want:
                        names.add(str(sh.Name))
                except Exception:
                    continue
    except Exception:
        pass
    return names


def _shapes_geometry(ws):
    out = []
    try:
        for shp in ws.Shapes:
            out.append({'name': str(shp.Name), 'left': float(shp.Left), 'top': float(shp.Top),
                        'width': float(shp.Width), 'height': float(shp.Height)})
    except Exception:
        pass
    return out


def _sheet_snapshot(ws):
    """シートの「今」を記憶に写す（値と数式と図形の位置と書式の姿）。大きすぎる表は None（差分は数えない）。"""
    try:
        ur = ws.UsedRange
        n = int(ur.Cells.CountLarge)
        if n > _SNAPSHOT_MAX_CELLS:
            return {'too_big': n}
        snap = {'addr': _addr(ur), 'row': int(ur.Row), 'col': int(ur.Column),
                'values': _rows_of(ur.Value), 'shapes': _shapes_geometry(ws)}
        try:
            snap['formulas'] = _rows_of(ur.Formula)
        except Exception:
            snap['formulas'] = None
        snap['format'] = _format_snapshot(ws)
        return snap
    except Exception:
        return None


# 書式の姿（2026-09-05）: 値の明細だけでは「見た目に何をしたか」が追えない。tidy が罫線・表示形式・
#   列幅・寄せを変えても、値の明細には一行も出なかった。セル 1 つずつ書式を読むと遅いので、
#   **列ごと（幅・表示形式・寄せ・太字）と範囲ごと（罫線 6 か所）**という粗さで写す＝COM の読みは 100 回台。
_XL_BORDERS = ((7, '左'), (8, '上'), (9, '下'), (10, '右'), (11, '内・縦'), (12, '内・横'))
_XL_ALIGN = {1: '標準', -4131: '左', -4108: '中央', -4152: '右', 5: '繰り返し', -4130: '両端', 7: '選択範囲内'}
_FORMAT_COLS_MAX = 40                  # 書式を写す列の上限（横に長い表で遅くならないように）


def _mixed(v, table=None):
    """COM が「まちまち」を None で返すのを言葉にする。"""
    if v is None:
        return 'まちまち'
    if table:
        return table.get(v, str(v))
    return str(v)


def _line_text(v):
    if v is None:
        return 'まちまち'
    return 'なし' if v == -4142 else 'あり'


def _format_snapshot(ws):
    """列ごとの書式（幅・表示形式・寄せ・太字）と、範囲の罫線 6 か所を写す。読めなければ None。"""
    try:
        ur = ws.UsedRange
        c0, ncols = int(ur.Column), int(ur.Columns.Count)
        cols = {}
        for j in range(min(ncols, _FORMAT_COLS_MAX)):
            col = ur.Columns(j + 1)
            cols[_col_letter(c0 + j)] = {
                'width': f"{float(ws.Columns(c0 + j).ColumnWidth):.2f}",
                'numfmt': _mixed(col.NumberFormatLocal),
                'halign': _mixed(col.HorizontalAlignment, _XL_ALIGN),
                'bold': _mixed(col.Font.Bold, {True: '太字', False: '並'}),
            }
        borders = {}
        for idx, name in _XL_BORDERS:
            try:
                borders[name] = _line_text(ur.Borders(idx).LineStyle)
            except Exception:
                pass
        return {'addr': _addr(ur), 'cols': cols, 'borders': borders}
    except Exception:
        return None


def _format_changes(before_fmt, ws):
    """書式の変化を 1 行 1 件で返す（列幅・表示形式・寄せ・太字・罫線）。読めなければ空。"""
    after = _format_snapshot(ws)
    if not before_fmt or not after:
        return []
    name = str(getattr(ws, 'Name', '') or '')
    out = []
    labels = (('width', '列幅'), ('numfmt', '表示形式'), ('halign', '寄せ'), ('bold', '太字'))
    for col, b in (before_fmt.get('cols') or {}).items():
        a = (after.get('cols') or {}).get(col)
        if not a:
            continue
        for key, label in labels:
            if str(b.get(key)) != str(a.get(key)):
                out.append({'sheet': name, 'addr': f"{col}列", 'what': label,
                            'before': str(b.get(key)), 'after': str(a.get(key))})
    for _idx, nm in _XL_BORDERS:
        b, a = (before_fmt.get('borders') or {}).get(nm), (after.get('borders') or {}).get(nm)
        if a is not None and b is not None and str(b) != str(a):
            out.append({'sheet': name, 'addr': after.get('addr') or '', 'what': f"罫線・{nm}",
                        'before': str(b), 'after': str(a)})
    return out


def _cell_pair(snap, key, i, j):
    grid = snap.get(key)
    if not grid or i >= len(grid) or j >= len(grid[i]):
        return None
    return grid[i][j]


def _changes_of(before, ws, sheet_name=None):
    """控えを取った時点からの変化を **1 セル 1 行** で全部返す。数えられなければ None。

    2026-09-05 に足した。それまで報告に出るのは「変わったセル: N 個」と例が 8 個だけで、
    **どのセルが何から何になったかを人（と Claude Code）が確かめる道が無かった**。
    数だけでは「戻せるから大丈夫」にしかならず、**戻さずに検分する**ことができない。
    """
    if not before or before.get('too_big') or not before.get('values'):
        return None
    after = _sheet_snapshot(ws)
    if not after or after.get('too_big') or not after.get('values'):
        return None
    name = sheet_name or str(getattr(ws, 'Name', '') or '')
    r0b, c0b = before['row'], before['col']
    r0a, c0a = after['row'], after['col']
    seen = set()
    for src, (r0, c0) in ((before, (r0b, c0b)), (after, (r0a, c0a))):
        for i, row in enumerate(src.get('values') or []):
            for j, _v in enumerate(row):
                seen.add((r0 + i, c0 + j))
    rows = []
    for (r, c) in sorted(seen):
        b = _cell_pair(before, 'values', r - r0b, c - c0b) if r >= r0b and c >= c0b else None
        a = _cell_pair(after, 'values', r - r0a, c - c0a) if r >= r0a and c >= c0a else None
        bf = _cell_pair(before, 'formulas', r - r0b, c - c0b) if before.get('formulas') and r >= r0b and c >= c0b else None
        af = _cell_pair(after, 'formulas', r - r0a, c - c0a) if after.get('formulas') and r >= r0a and c >= c0a else None
        if _text_of(b) == _text_of(a) and _text_of(bf) == _text_of(af):
            continue
        rows.append({'sheet': name, 'addr': f"{_col_letter(c)}{r}",
                     'before': _text_of(b), 'after': _text_of(a),
                     'before_formula': _text_of(bf), 'after_formula': _text_of(af)})
    return rows


_NUMLIKE_RE = re.compile(r'^-?[\d,]+(?:\.\d+)?%?$')


def _formula_rate(changed_rows):
    """書いたセルのうち「数に見えるもの」と「数式で書いたもの」を数える → (数, 数式) か None（純 Python）。

    2026-09-06・Claude for Excel の助言「派生値の数式率。100% でないものは後で壊れる」を測る。
    合計・割合・突き合わせを自分で計算して値で置くと、人が元の行を直しても黙って古いままになる。
    関所にはしない（連番・区分のように数を置くのが正しい仕事もある）。数えて見せるだけ。
    """
    rows = [r for r in (changed_rows or []) if str(r.get('after') or '') != '']
    if not rows:
        return None
    n_num = n_f = 0
    for r in rows:
        after = str(r.get('after') or '').strip()
        is_f = str(r.get('after_formula') or '').startswith('=')
        if is_f or _NUMLIKE_RE.match(after):
            n_num += 1
            if is_f:
                n_f += 1
    return (n_num, n_f) if n_num else None


def _unbound_shape_names(ws):
    """マクロ（OnAction）の付いていない図形の名前の一覧（COM）。読めなければ空。"""
    out = []
    try:
        count = int(ws.Shapes.Count)
    except Exception:
        return out
    for i in range(1, count + 1):
        try:
            shp = ws.Shapes(i)
            if not str(shp.OnAction or ''):
                out.append(str(shp.Name))
        except Exception:
            continue
    return out


def _pick_shape(ws, nm):
    """図形を名前で拾う。材料の表示（名前「文字」）をそのまま渡されても当てる（2026-09-06）。

    材料の「図形・ボタン」は `ロゴ枠「総務課」` のように **名前と表示テキストを並べて**出す。
    AI がその 1 行をそのまま name に写して 2 回とも同じ所で失敗した（実測 2026-09-06・
    1 往復・約 8 秒を毎回捨てていた）。名前そのもの → 飾りを剥がした形 → 表示テキスト、の順で照合する。
    見つからなければ None（呼び側が「材料に出ている名前を使う」と言って止める）。
    """
    nm = str(nm or '').strip()
    if not nm:
        return None
    try:
        return ws.Shapes(nm)
    except Exception:
        pass
    key = re.sub(r'\s+', '', nm)
    bare = re.sub(r'\s+', '', re.sub(r'「[^「」]*」\s*$', '', nm))   # 末尾の「…」を落とす
    try:
        count = int(ws.Shapes.Count)
    except Exception:
        return None
    for i in range(1, count + 1):
        try:
            shp = ws.Shapes(i)
            sname = str(shp.Name)
        except Exception:
            continue
        try:
            text = str(shp.TextFrame.Characters().Text or '')
        except Exception:
            text = ''
        labeled = f"{sname}「{text}」" if text else sname
        for form in (sname, labeled, text):
            f = re.sub(r'\s+', '', str(form or ''))
            if f and (f == key or f == bare):
                return shp
    return None


# 番地の手前に # が来たら色コード（#FFC7CE から FFC7 を拾っていた・2026-09-06 実測）。
# 後ろに英字が続くのも番地ではない（色コードの後半・語の一部）。
_REPORT_ADDR_RE = re.compile(r'(?<![A-Za-z0-9$!.#])(\$?[A-Za-z]{1,3}\$?\d{1,7})(?![0-9(A-Za-z])')
_REPORT_RANGE_RE = re.compile(r'(\$?[A-Za-z]{1,3}\$?\d{1,7})\s*:\s*\$?[A-Za-z]{1,3}\$?\d{1,7}')


_REPORT_FORMULA_RE = re.compile(r'=[A-Za-z0-9_$!\'"()\[\]:,．.+\-*/&<>^%　-ヿ一-鿿]+')
# 「〜は空欄」と自分で言っている番地（＝空なのが分かっていて書いている）は、空でも指摘しない
_REPORT_SAID_BLANK_RE = re.compile(r'(空欄|空です|空のまま|空だっ|空である|未入力|何も入っていな)')
_REPORT_SAID_BLANK_WINDOW = 20      # 番地の後ろ何字まで見るか
# 報告の段見出し。【やったこと】だけが「書いた」と言っている段（2026-09-08）
_REPORT_DID_RE = re.compile(r'【\s*やったこと\s*】(.*?)(?=\n\s*【|\Z)', re.S)


def _report_did_section(report):
    """報告のうち【やったこと】の段だけ。段が無ければ全文（古い形の報告・自由文のとき）。"""
    hits = _REPORT_DID_RE.findall(report)
    return "\n".join(hits) if hits else report


def _report_addrs(report, limit=40):
    """報告文に出てくるセル番地（重複なし・上限つき・純 Python）。

    2026-09-06・Claude for Excel の助言「report に書いた番地・件数を道具が読み直して照合する。
    『できた』と書いて現物に無いものが一番信用を壊す」。数式の中の番地（=SUM(D2:D301)）は
    実在するので拾っても害が無い。Q17 のような問い番号を拾うのは避けられないので、
    照合は関所にせず「空だった番地」の**注意書き**として出す。
    """
    # 2026-09-08: 拾いすぎて誤検知が続いたので、見る所と見ない字を絞った。
    #  ・見るのは【やったこと】の段だけ（書いたと言っているのはここ。【確認していないこと】に出る
    #    「使用範囲が H31 まで膨張」の H31 を「空です」と言っていた）。段が無ければ今までどおり全文。
    #  ・数式の字面（=G27/E100）の中の番地は外す。**参照先**であって「書いた番地」ではない
    #    （壊れた式の原因として挙げた E100 を毎回「空です」と言っていた）。
    text = _report_did_section(str(report or ''))
    text = _REPORT_FORMULA_RE.sub(' ', text)
    # 範囲（A1:H3）は左上だけ見る。右下は表の端で、空でも正しいことが多い（2026-09-06・
    # 報告の「A1:H3 のタイトル行」から H3 を拾って「空です」と言っていた）。
    text = _REPORT_RANGE_RE.sub(lambda m: m.group(1) + ' ', text)
    out = []
    for m in _REPORT_ADDR_RE.finditer(text):
        a = m.group(1).replace('$', '').upper()
        if a in out or not _a1_to_rc(a):
            continue
        tail = text[m.end():m.end() + _REPORT_SAID_BLANK_WINDOW].split('\n')[0].split('。')[0]
        nxt = _REPORT_ADDR_RE.search(tail)
        if nxt:
            tail = tail[:nxt.start()]      # 次の番地より先は「その番地の話」＝ここで切る
        if _REPORT_SAID_BLANK_RE.search(tail):
            continue           # 「E100 が空欄のため」＝自分で空だと言っている＝指摘しても意味が無い
        out.append(a)
        if len(out) >= limit:
            break
    return out


def _change_side(row, side):
    """明細の片側を文字にする。数式のあるセルは数式の字面を出す（値だけでは検分にならない）。"""
    f = row.get(side + '_formula') or ''
    return f if f.startswith('=') else (row.get(side) or '')


def _diff_sheet(before, ws, limit=8):
    """控えを取った時点からの差分 → (変わったセルの数, 例の文の並び)。数えられなければ (None, [])。"""
    rows = _changes_of(before, ws)
    if rows is None:
        return None, []
    examples = [f"{r['addr']} {r['before'][:14]!r}→{_change_side(r, 'after')[:14]!r}"
                for r in rows[:limit]]
    return len(rows), examples


_CHANGES_HEADER = ['種類', 'シート', '番地', '前', '後', '前の数式', '後の数式']


def _write_changes_file(rows, path=_LAST_AGENT_CHANGES_FILE, fmt_rows=()):
    """変わったセルの明細を TSV で残す（1 セル 1 行）。戻り値＝書けたパス（書けなければ None）。

    「種類」は 値（セルの中身）か 書式（列幅・表示形式・寄せ・太字・罫線）。
    Excel でそのまま開けるように UTF-8（BOM 付き）。多すぎるときは _CHANGES_MAX_ROWS 行で切る。
    """
    try:
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f, delimiter='\t', lineterminator='\n')
            w.writerow(_CHANGES_HEADER)
            for r in rows[:_CHANGES_MAX_ROWS]:
                w.writerow(['値', r['sheet'], r['addr'], r['before'], r['after'],
                            r['before_formula'], r['after_formula']])
            for r in fmt_rows:
                w.writerow([f"書式・{r['what']}", r['sheet'], r['addr'], r['before'], r['after'], '', ''])
        return path
    except Exception as ex:
        print(f"⚠ 明細を書けませんでした: {ex}")
        return None


def _clip(s, n):
    s = str(s or '')
    return s if len(s) <= n else s[:n - 1] + '…'


def _changes_table(rows, show=12, multi_sheet=False):
    """明細を人が読める形に整える（頭の show 行だけ・残りは行数で言う）。戻り値＝行の並び。"""
    out = []
    head = (f"  {'シート':<10}" if multi_sheet else "") + f"  {'番地':<6}  {'前':<22}  →  後"
    out.append(head)
    for r in rows[:show]:
        line = (f"  {_clip(r['sheet'], 10):<10}" if multi_sheet else "")
        line += f"  {r['addr']:<6}  {_clip(_change_side(r, 'before'), 22):<22}  →  {_clip(_change_side(r, 'after'), 30)}"
        out.append(line)
    if len(rows) > show:
        out.append(f"  …ほか {len(rows) - show} 行（全部は明細のファイルに）")
    return out


def _format_table(fmt_rows, show=12):
    """書式の変化を人が読める形に整える（見た目に何をしたかは値の明細に出ないため）。"""
    out = [f"  {'どこ':<8}  {'何を':<8}  {'前':<14}  →  後"]
    for r in fmt_rows[:show]:
        out.append(f"  {_clip(r['addr'], 8):<8}  {_clip(r['what'], 8):<8}  "
                   f"{_clip(r['before'], 14):<14}  →  {_clip(r['after'], 20)}")
    if len(fmt_rows) > show:
        out.append(f"  …ほか {len(fmt_rows) - show} 件（全部は明細のファイルに）")
    return out


def show_changes(path=_LAST_AGENT_CHANGES_FILE):
    """`agent --changes`: 直前の仕事で変わったセルと書式の明細をそのまま出す（Excel に触らない）。"""
    if not os.path.exists(path):
        print(f"明細がありません（{path}）。agent で何か書いた後に出ます。")
        return False
    try:
        with open(path, encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f, delimiter='\t'))
    except Exception as ex:
        print(f"エラー: 明細を読めません: {ex}")
        return False
    if len(rows) <= 1:
        print("明細: 変わったところはありません")
        return True
    body, fmt = [], []
    for r in rows[1:]:
        if not r:
            continue
        kind = r[0] if len(r) > 0 else '値'
        if kind.startswith('書式'):
            fmt.append({'sheet': r[1], 'addr': r[2], 'what': kind.split('・', 1)[-1],
                        'before': r[3], 'after': r[4] if len(r) > 4 else ''})
        else:
            body.append({'sheet': r[1], 'addr': r[2], 'before': r[3], 'after': r[4] if len(r) > 4 else '',
                         'before_formula': r[5] if len(r) > 5 else '',
                         'after_formula': r[6] if len(r) > 6 else ''})
    print(f"変わったセル: {len(body)} 個   {path}")
    if body:
        multi = len({r['sheet'] for r in body}) > 1
        for line in _changes_table(body, show=len(body), multi_sheet=multi):
            print(line)
    if fmt:
        print(f"書式の変化: {len(fmt)} 件")
        for line in _format_table(fmt, show=len(fmt)):
            print(line)
    return True


def _prune_agent_backups(keep=_AGENT_BACKUP_KEEP, stem=None):
    """控えを新しい keep 個だけ残す。stem（ブック名の拡張子より前）を渡すと、そのブックの控えだけを数える。

    前はブック横断で 5 個だった＝別のブックを 5 回回すと、前のブックの控えが消えて --undo できなくなった
    （2026-09-04）。控えの名前は「ブック名_agent_before_日時.拡張子」。
    """
    head = f"{stem}{_AGENT_BACKUP_MARK}" if stem else None
    try:
        files = [os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR)
                 if (f.startswith(head) if head else _AGENT_BACKUP_MARK in f)
                 and not f.endswith(_AGENT_BACKUP_META_EXT)]        # 隣の覚書は数に入れない（控えと一緒に消す）
    except OSError:
        return
    for p in sorted(files, key=lambda p: os.path.getmtime(p), reverse=True)[keep:]:
        for q in (p, _backup_meta_path(p)):
            try:
                os.remove(q)
            except OSError:
                pass


def _agent_backup(wb, ws, sheet, request, run_id=None):
    """最初の書き込みの前に控えを取る（SaveCopyAs＝未保存の変更込み）。戻り値＝控えのパス（取れなければ None）。

    run_id はその走行の名札（走行台帳と結ぶ）。--undo で戻したとき、台帳のその走行に印が付く
    （2026-09-06・「人が戻した割合」が本番でいちばん正直な失敗率だから）。
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stem, ext = os.path.splitext(str(wb.Name))
    if not ext.lower().startswith('.xls'):
        ext = '.xlsx'
    path = os.path.join(BACKUP_DIR, f"{stem}{_AGENT_BACKUP_MARK}{time.strftime('%Y%m%d_%H%M%S')}{ext}")
    try:
        wb.SaveCopyAs(path)
    except Exception as ex:
        print(f"⚠ 控えを作れませんでした（--undo で戻せません）: {ex}")
        return None
    meta = {'book': str(wb.Name), 'sheet': sheet, 'sheets': [sheet], 'path': path,
            'request': (request or '')[:300], 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
            'run_id': run_id, 'shapes': _shapes_geometry(ws)}
    _save_undo_meta(meta)
    _prune_agent_backups(stem=stem)             # 同じブックの控えだけを数えて間引く（他のブックのは消さない）
    print(f"控え: {path}（書き換える前のブック全体。戻すなら py vba_manager.py agent --undo。"
          f"残っている控えの一覧は agent --backups）")
    return path


def _backup_meta_path(path):
    """控えのファイル → その控えの覚書（同じ名前＋.json）。"""
    return str(path) + _AGENT_BACKUP_META_EXT


def _save_undo_meta(meta):
    """控えの覚書を 2 か所に書く: 直前用（_last_agent_undo.json）と、控えのファイルの隣（控えごと・2026-09-06）。

    隣の覚書があるので、直前の 1 つだけでなく、残っている控え（ブックごとに 5 個）のどれにでも戻せる
    （agent --backups で一覧 → agent --undo 番号）。それまでは控えは 5 個あるのに戻せるのは 1 つだった。
    書けなくても仕事は止めない（False を返す）。
    """
    ok = True
    try:
        with open(_LAST_AGENT_UNDO_FILE, 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=1)
    except OSError as ex:
        print(f"⚠ 控えの覚書を書けませんでした: {ex}")
        ok = False
    # 隣の覚書は、控えのファイルが実在するときだけ書く（テストが 'p'・'x' のような相対名を渡すと、
    # カレントに p.json・x.json が残っていた＝プロジェクト直下の残骸・2026-09-06 夜）
    if isinstance(meta, dict) and meta.get('path') and os.path.isfile(str(meta['path'])):
        with contextlib.suppress(OSError, TypeError, ValueError):
            with open(_backup_meta_path(meta['path']), 'w', encoding='utf-8') as f:
                json.dump(meta, f, ensure_ascii=False, indent=1)
    return ok


def _load_undo_meta(path=None):
    """覚書を読む（既定は直前用）。読めなければ None。"""
    try:
        with open(path or _LAST_AGENT_UNDO_FILE, encoding='utf-8') as f:
            meta = json.load(f)
    except (OSError, ValueError):
        return None
    return meta if isinstance(meta, dict) else None


def _list_backups():
    """残っている控え（新しい順・全ブック）。1 件＝{'n': 番号, 'path': 控え, 'meta': 覚書 or None, 'mtime': …}。

    覚書の無い控え（この作りより前に取ったもの）も一覧には出す（戻せない印を付ける）。
    直前用の覚書がその控えのものなら、隣の覚書が無くてもそれを使う。
    """
    try:
        names = [f for f in os.listdir(BACKUP_DIR)
                 if _AGENT_BACKUP_MARK in f and not f.endswith(_AGENT_BACKUP_META_EXT)]
    except OSError:
        return []
    last = _load_undo_meta() or {}
    items = []
    for f in names:
        p = os.path.join(BACKUP_DIR, f)
        meta = _load_undo_meta(_backup_meta_path(p))
        if meta is None and last.get('path') and _same_path(last['path'], p):
            meta = last
        try:
            mtime = os.path.getmtime(p)
        except OSError:
            mtime = 0
        items.append({'path': p, 'meta': meta, 'mtime': mtime, 'n': None})
    items.sort(key=lambda d: (d['mtime'], d['path']), reverse=True)
    n = 0
    for d in items:                      # 番号は戻せる控え（覚書あり）にだけ振る（覚書の無い控えは名前で指す）
        if d['meta']:
            n += 1
            d['n'] = n
    return items


def _same_path(a, b):
    try:
        return os.path.normcase(os.path.abspath(str(a))) == os.path.normcase(os.path.abspath(str(b)))
    except (TypeError, ValueError):
        return False


def _pick_backup(which):
    """--undo の指定（一覧の番号／走行の名札 run_id／控えのファイル名かパス）→ 一覧の 1 件。無ければ None。"""
    w = str(which or '').strip()
    if not w:
        return None
    for d in _list_backups():
        m = d['meta'] or {}
        if ((w.isdigit() and d['n'] is not None and int(w) == d['n']) or (m.get('run_id') and m['run_id'] == w)
                or os.path.basename(d['path']) == w or _same_path(w, d['path'])):
            return d
    return None


def _runs_after(meta):
    """この控えより後に、同じブックであった走行（台帳から）。戻すとそのぶんが消える＝止める材料。"""
    if not isinstance(meta, dict):
        return []
    book, t0, rid = meta.get('book'), str(meta.get('time') or ''), meta.get('run_id')
    return [r for r in _runs_load()
            if r.get('book') == book and str(r.get('time') or '') > t0 and r.get('run_id') != rid]


def backups_list():
    """残っている控えの一覧（agent --backups）。番号で `agent --undo 番号` と指せる。"""
    items = _list_backups()
    if not items:
        print(f"控えはありません（{BACKUP_DIR}）。agent がシートに書く前に取ります（ブックごとに新しい {_AGENT_BACKUP_KEEP} 個）")
        return True
    last = _load_undo_meta() or {}
    old = [d for d in items if not d['meta']]
    print(f"残っている控え（新しい順・{BACKUP_DIR}）:")
    if not any(d['meta'] for d in items):
        print("  （--undo で戻せる控えはありません）")
    for d in items:
        m = d['meta']
        if not m:
            continue
        mark = "　← 直前（--undo だけで戻る）" if (last.get('path') and _same_path(last['path'], d['path'])) else ""
        if m.get('stale'):
            mark += f"　⚠ この後に {m['stale']}"
        later = _runs_after(m)
        if later:
            mark += f"　⚠ この後に同じブックで {len(later)} 走行（戻すと消える＝--force が要る）"
        print(f"  [{d['n']}] {m.get('time', '?')}  {m.get('book', '?')}!" + "・".join(_undo_sheets(m))
              + (f"  走行 {m['run_id']}" if m.get('run_id') else "") + mark)
        print(f"        依頼「{(m.get('request') or '')[:60]}」")
    if old:
        # この作り（控えごとの覚書）より前の控え。どのシートを書き換えたか分からないので --undo では戻せない
        print(f"  ほか {len(old)} 個は覚書なし（2026-09-06 より前の控え＝--undo では戻せない。"
              "戻すなら Excel で開いて手で写す。古い順に消してよい）:")
        print("    " + " / ".join(os.path.basename(d['path']) for d in old[:6])
              + (f" / …ほか {len(old) - 6} 個" if len(old) > 6 else ""))
    print("（戻す: agent --undo 番号（走行の名札でも可）。--dry-run で下見。そのブックを開いてアクティブにしてから）")
    return True


def backups_prune(days, force=False):
    """控えの置き場（BACKUP_DIR）で、days 日より古いファイルを数えて見せる。force で消す（2026-09-06 夜）。

    agent の間引きは自分の控え 5 個だけで、replace-procedure 等の控えは春から溜まって 3,615 ファイル＝
    一覧を出すだけで 120 秒を超えた。消すのは人の判断なので、既定は数と大きさを見せるだけ。
    """
    days = int(days)
    if days < 1:
        print("エラー: --prune-days は 1 以上")
        return False
    cutoff = time.time() - days * 86400
    old, total, size_old = [], 0, 0
    try:
        for f in os.listdir(BACKUP_DIR):
            p = os.path.join(BACKUP_DIR, f)
            if not os.path.isfile(p):
                continue
            total += 1
            try:
                mt = os.path.getmtime(p)
            except OSError:
                continue
            if mt < cutoff:
                old.append((mt, p))
                size_old += os.path.getsize(p)
    except OSError as ex:
        print(f"エラー: 控えの置き場を読めません: {ex}")
        return False
    old.sort()
    print(f"控えの置き場: {BACKUP_DIR}  ファイル {total:,} 個。{days} 日より古いもの {len(old):,} 個・{size_old / 1048576:.1f} MB"
          + (f"（いちばん古い {time.strftime('%Y-%m-%d', time.localtime(old[0][0]))}〜"
             f"{time.strftime('%Y-%m-%d', time.localtime(old[-1][0]))}）" if old else ""))
    if not old:
        return True
    if not force:
        print("  " + " / ".join(os.path.basename(p) for _m, p in old[:5]) + (" / …" if len(old) > 5 else ""))
        print(f"（見せただけで消していません。消すなら agent --prune-days {days} --force）")
        return True
    n_del = 0
    for _m, p in old:
        try:
            os.remove(p)
            n_del += 1
        except OSError:
            pass
    print(f"消しました: {n_del:,} 個（残り {total - n_del:,} 個）")
    return True


def _undo_sheets(meta):
    """控えの覚書 → 戻す先のシート名（順番は控えた順）。古い覚書（'sheet' だけ）も読める。"""
    if not isinstance(meta, dict):
        return []
    got = meta.get('sheets') or ([meta.get('sheet')] if meta.get('sheet') else [])
    out = []
    for s in got:
        name = str(s or '').strip()
        if name and name not in out:
            out.append(name)
    return out


def _undo_shapes(meta, primary):
    """控えの覚書 → シート名ごとの図形の位置。'shapes'（1 枚ぶん）だけの古い覚書も読める。"""
    by = dict((meta.get('shapes_by_sheet') or {}) if isinstance(meta, dict) else {})
    if primary not in by and isinstance(meta, dict) and meta.get('shapes'):
        by[primary] = meta['shapes']
    return by


def _undo_add_sheet(wb, name):
    """控えの覚書に「このシートも戻す対象」と足す（2026-09-05・別シートにも書けるようにしたため）。

    書き換える前に呼ぶ＝図形の位置も今のうちに写す。まだ無いシート（これから作る）は控えに入っていない
    ので足さない（作った物は 'created' の側が消す）。
    """
    try:
        ws = wb.Sheets(name)
    except Exception:
        return
    try:
        with open(_LAST_AGENT_UNDO_FILE, encoding='utf-8') as f:
            meta = json.load(f)
    except (OSError, ValueError):
        return
    if not isinstance(meta, dict) or name in _undo_sheets(meta):
        return
    meta['sheets'] = _undo_sheets(meta) + [name]
    by = dict(meta.get('shapes_by_sheet') or {})
    with contextlib.suppress(Exception):
        by[name] = _shapes_geometry(ws)
    meta['shapes_by_sheet'] = by
    _save_undo_meta(meta)


def _mark_undo_stale(what, book=None):
    """控えの覚書に「この後で別の仕事をした」と印を付ける（2026-09-04）。

    覚書は sheet のループでしか書き換わらない。build や macro を回した後に --undo すると、いまの仕事ではなく
    その前の sheet 仕事の控えで上書きしてしまう（build で組んだシートが黙って消える）。印があるときは
    undo_agent が既定で止まる（それでも戻すなら --force）。控えが無ければ何もしない。

    2026-09-08: book（いま組んだ・直したブックの名前）を渡すと、**同じブックの控えにだけ**印を付ける。
    渡さないと、別のブックで build を回しただけで無関係な控えに「⚠ この後に build が走った」が付き、
    一覧が嘘をつく（build を回していない走行にその注記が出た、という報告の正体）。
    """
    try:
        with open(_LAST_AGENT_UNDO_FILE, encoding='utf-8') as f:
            meta = json.load(f)
    except (OSError, ValueError):
        return
    if not isinstance(meta, dict):
        return
    if book and str(meta.get('book') or '') and str(meta['book']) != str(book):
        return                      # 別のブックの控え＝この仕事とは関係ない
    meta['stale'] = str(what)
    meta['stale_time'] = time.strftime('%Y-%m-%d %H:%M:%S')
    _save_undo_meta(meta)


# 作った「重い物」を控えの覚書に記録して --undo で消す（2026-09-04）。
#   控えはセルと図形しか戻せず、テーブル・ピボット・スライサー・グラフ・クエリ・メジャー・別シートは
#   貼り戻しても残っていた（ピボットのあるシートは Clear 自体が拒まれる）。作った順に積んで、逆順で消す。
_CREATED_PATTERNS = (
    # 名前に空白が入ることがある（Excel が自動で付ける「グラフ 1」）ので、後ろの見出しまでで切る
    ('table', re.compile(r'^テーブル作成: \[([^\]]+)\] (.+?)  範囲=', re.M), ('sheet', 'name')),
    ('pivot', re.compile(r'^ピボット作成: \[([^\]]+)\] (.+?)  ソース=', re.M), ('sheet', 'name')),
    ('chart', re.compile(r'^グラフ作成: \[([^\]]+)\] (.+?)  種別 ', re.M), ('sheet', 'name')),
    ('slicer', re.compile(r'^スライサー追加: (.+?)  ソース=', re.M), ('name',)),
    ('query', re.compile(r'^クエリ作成: (.+?)（', re.M), ('name',)),
    ('table', re.compile(r'^シートに読み込みました: .+? → ([^!]+)!\S+?（テーブル: ([^,）]+)', re.M), ('sheet', 'name')),
    ('sheet', re.compile(r'^シート追加: (.+)$', re.M), ('name',)),
    ('sheet', re.compile(r'^シート複製: .+ → (.+)$', re.M), ('name',)),
    ('name', re.compile(r'^名前付き範囲を追加: (\S+) →', re.M), ('name',)),
    ('measure', re.compile(r'^メジャー作成: (\S+)\[([^\]]+)\]', re.M), ('table', 'name')),
    ('calc_field', re.compile(r'^計算フィールド作成: (\S+?)\[([^\]]+)\]', re.M), ('pivot', 'name')),
)


def _created_from_results(results):
    """[(表示名, ok, 出力)] → この往復で作られた重い物の並び（純 Python）。"""
    out = []
    for _label, ok, text in results:
        if not ok or not text:
            continue
        for kind, rx, keys in _CREATED_PATTERNS:
            for m in rx.finditer(text):
                item = {'kind': kind}
                for k, v in zip(keys, m.groups()):
                    item[k] = str(v).strip()
                if item.get('name'):
                    out.append(item)
    return out


def _record_created(items):
    """作った重い物を控えの覚書（_last_agent_undo.json）に足す。控えが無ければ何もしない。"""
    if not items:
        return
    try:
        with open(_LAST_AGENT_UNDO_FILE, encoding='utf-8') as f:
            meta = json.load(f)
    except (OSError, ValueError):
        return
    meta.setdefault('created', []).extend(items)
    _save_undo_meta(meta)


def _find_listobject(wb, name):
    for sh in wb.Worksheets:
        try:
            for lo in sh.ListObjects:
                if str(lo.Name) == name:
                    return lo
        except Exception:
            continue
    return None


def _delete_created(xl, wb, items):
    """作った重い物を逆順で消す → (消した数, 消せなかったものの説明)。"""
    done, failed = [], []
    for it in reversed(items or []):
        kind, name, sh = it.get('kind'), it.get('name'), it.get('sheet')
        try:
            if kind == 'pivot':
                from vbam_heavy import _find_pivot
                _psh, pt = _find_pivot(wb, name)
                if pt is None:
                    continue
                pt.TableRange2.Clear()
            elif kind == 'table':
                lo = _find_listobject(wb, name)
                if lo is None:
                    continue
                lo.Delete()
            elif kind == 'slicer':
                got = None
                for sc in wb.SlicerCaches:
                    for sl in sc.Slicers:
                        if str(sl.Name) == name:
                            got = sl
                            break
                    if got is not None:
                        break
                if got is None:
                    continue
                got.Delete()
            elif kind == 'chart':
                wb.Sheets(sh).ChartObjects(name).Delete()
            elif kind == 'query':
                wb.Queries(name).Delete()
            elif kind == 'sheet':
                if int(wb.Sheets.Count) <= 1:
                    failed.append(f"{kind} {name}（最後の 1 枚は消せない）")
                    continue
                old = xl.DisplayAlerts
                xl.DisplayAlerts = False
                try:
                    wb.Sheets(name).Delete()
                finally:
                    xl.DisplayAlerts = old
            elif kind == 'name':
                wb.Names(name).Delete()
            elif kind == 'measure':
                for mm in wb.Model.ModelMeasures:
                    if str(mm.Name) == name:
                        mm.Delete()
                        break
            elif kind == 'calc_field':
                from vbam_heavy import _find_pivot
                _psh, pt = _find_pivot(wb, it.get('pivot'))
                if pt is None:
                    continue
                pt.CalculatedFields()(name).Delete()
            else:
                continue
            done.append(f"{kind} {name}")
        except Exception as ex:
            failed.append(f"{kind} {name}: {ex}")
    return done, failed


def _restore_formulas(ws, addr, snap):
    """貼り戻した範囲の数式を、控えから読んだ字面で入れ直す（外部リンク化を消す）。

    ブックをまたぐ貼り付けは、Excel が別シート参照を控えファイルへの外部リンクに書き換える
    （='明細'!B2 → ='C:\\…\\[名簿_agent_before_….xlsx]明細'!B2）。控えは 5 個で間引かれるので、
    そのままだといずれ #REF! になる（2026-09-04・別インスタンスの Excel で再現して確認）。
    控えの UsedRange から読んだ .Formula は控えブックの中での字面＝外部リンクが付いていないので、
    貼り付けの直後にそれで上書きすれば元に戻る（値だけのセルも .Formula で同じ値が入る）。
    """
    grid = (snap or {}).get('formulas')
    if not grid:
        return None                     # 控えが大きすぎて写せなかった（_SNAPSHOT_MAX_CELLS 超）
    try:
        rng = ws.Range(addr)
        if int(rng.Rows.Count) == 1 and int(rng.Columns.Count) == 1:
            rng.Formula = grid[0][0]
        else:
            rng.Formula = tuple(tuple(r) for r in grid)
        return True
    except Exception as ex:
        print(f"  ⚠ 数式を控えの字面で入れ直せませんでした: {ex}")
        return False


def _drop_backup_link(wb, path):
    """貼り戻しの後、控えファイルへの外部リンクが残っていないか見て、残っていれば外す。

    _restore_formulas が効かなかったとき（控えが大きすぎる・書き戻しに失敗）の最後の受け皿。
    このブック自身にリンク先を差し替えると、Excel が内部参照に畳み直す。
    """
    try:
        srcs = wb.LinkSources(1) or ()          # xlLinkTypeExcelLinks
    except Exception:
        return None
    want = os.path.normcase(os.path.abspath(path))
    hit = [s for s in srcs if os.path.normcase(str(s)) == want]
    if not hit:
        return True
    try:
        wb.ChangeLink(hit[0], str(wb.FullName), 1)
        print("  控えへの外部リンクが残っていたので、このブック自身に差し替えました")
        return True
    except Exception as ex:
        print(f"  ⚠ 控えへの外部リンクが残りました（{hit[0]}）。数式を確かめてください: {ex}")
        return False


def undo_agent(target_file=None, dry_run=False, force=False, which=None):
    """`agent --undo`: 直前の agent が触ったシートを、控えの同じシートで置き換える（そのシートだけ）。

    シートは消さずに中身を入れ替える（消すと他のシートからの参照が #REF! になる）。保存はしない。
    控えを取った後に build／macro を回していると（stale の印）、既定で止まる＝--force で越える。
    which＝一覧（agent --backups）の番号・走行の名札・控えのファイル名で、直前より前の控えにも戻せる
    （2026-09-06）。その控えより後に同じブックで走行があれば、そのぶんが消えるので既定で止まる（--force で越える）。
    """
    if which and str(which) != 'latest':
        hit = _pick_backup(which)
        if not hit:
            print(f"エラー: その控えがありません: {which}（一覧は agent --backups）")
            return False
        meta = hit['meta']
        if not meta:
            print(f"エラー: この控えには覚書がありません（この作りより前の控え）: {hit['path']}"
                  "\n  どのシートを書き換えたか分からないので --undo では戻せません。Excel で開いて手で写してください")
            return False
        later = _runs_after(meta)
        if later and not force:
            print(f"エラー: この控え（{meta.get('time')}）より後に、同じブックで走行があります:")
            for r in later:
                print(f"  {r.get('time')}  [{r.get('mode')}] 「{(r.get('request') or '')[:50]}」")
            print(f"  これで戻すと、後からやったぶんが消えます。中身を見るなら --undo {which} --dry-run、"
                  f"それでも戻すなら --undo {which} --force。")
            return False
    else:
        meta = _load_undo_meta()
        if meta is None:
            print("エラー: 戻せる控えがありません（この道具で書き換えた記録が無い）")
            return False
    if meta.get('stale') and not force:
        print(f"エラー: この控えは、いまの仕事のものではありません。"
              f"\n  控え: {meta.get('time')} の sheet の仕事「{(meta.get('request') or '')[:60]}」"
              f"\n  その後にやったこと: {meta['stale']}（{meta.get('stale_time')}）"
              f"\n  これで戻すと、後からやったぶんが消えます。中身を見るなら --undo --dry-run、"
              "それでも戻すなら --undo --force。")
        return False
    path, book = meta.get('path'), meta.get('book')
    sheets = _undo_sheets(meta)
    if not sheets:
        print("エラー: 控えの覚書に戻す先のシートがありません")
        return False
    sheet = sheets[0]
    if not path or not os.path.isfile(path):
        print(f"エラー: 控えのファイルがありません: {path}")
        return False
    xl, wb = get_workbook(target_file)
    if str(wb.Name) != book:
        print(f"エラー: 控えは「{book}」のものです（いま開いているのは「{wb.Name}」）。"
              "そのブックを開いてアクティブにしてから、もう一度 --undo してください")
        return False
    have_sheets = [sh.Name for sh in wb.Sheets]
    missing = [s for s in sheets if s not in have_sheets]
    if missing:
        print(f"エラー: シート {'・'.join(missing)} がこのブックにありません（名前を変えた？）")
        return False
    created = meta.get('created') or []
    print(f"戻す先: {book} のシート " + "・".join(f"「{s}」" for s in sheets)
          + f"   控え: {path}（{meta.get('time')} に取得）")
    print(f"  そのときの依頼: {meta.get('request') or '（記録なし）'}")
    if meta.get('stale'):
        print(f"  ⚠ この控えを取った後に「{meta['stale']}」をしています（{meta.get('stale_time')}）"
              "＝そのぶんは戻すと消えます")
    if created:
        print("  そのとき作った物（逆順で消します）: "
              + " / ".join(f"{c.get('kind')} {c.get('name')}" for c in created))
    if dry_run:
        print("（--dry-run: 何も戻していません。実際に戻すなら --undo だけで）")
        return True
    from vbam_edit import _alerts_off
    src_wb = None
    events = None
    n_before, back, back_to = 0, 0, []
    try:
        with _alerts_off(xl):
            try:                                  # 控えは .xlsm のことがある＝開いた瞬間に Workbook_Open が走る
                events = bool(xl.EnableEvents)
                xl.EnableEvents = False
            except Exception:
                events = None
            src_wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True)
            src_names = [sh.Name for sh in src_wb.Sheets]
            lost = [s for s in sheets if s not in src_names]
            if lost:
                print(f"エラー: 控えにシート {'・'.join(lost)} がありません")
                return False
            if created:
                # 貼り戻しの前に消す（ピボットの残った範囲は Clear が拒まれる＝先に消さないと戻せない）。
                # 戻す先のシートそのものは消さない（消したら貼る先が無い）
                gone, ng = _delete_created(xl, wb, [c for c in created
                                                    if not (c.get('kind') == 'sheet' and c.get('name') in sheets)])
                if gone:
                    print(f"  作った物を消しました: {len(gone)} 個（" + " / ".join(gone) + "）")
                if ng:
                    print("  消せなかったもの: " + " / ".join(ng))
            shapes_by = _undo_shapes(meta, sheet)
            for name in sheets:                          # 書いたシートを 1 枚ずつ控えの姿に戻す（2026-09-05）
                ws = wb.Sheets(name)
                src_ws = src_wb.Sheets(name)
                src_ur = src_ws.UsedRange
                addr = _addr(src_ur)
                src_snap = _sheet_snapshot(src_ws)      # 数式の字面もここで写す（貼り付け後に入れ直す）
                n_sheet, _ = _diff_sheet(src_snap, ws, limit=0)
                # 大きすぎて数えられないシートは None。足し算で落ちると、作った物を消した後で戻しが止まっていた（2026-09-10 夜）
                n_before += n_sheet or 0
                try:
                    ws.UsedRange.Clear()
                except Exception:
                    pass
                src_ur.Copy()
                ws.Range(addr).PasteSpecial(-4104)      # xlPasteAll
                src_ur.Copy()
                ws.Range(addr).PasteSpecial(8)          # xlPasteColumnWidths
                xl.CutCopyMode = False
                # ブックをまたぐ貼り付けは別シート参照を控えへの外部リンクに書き換える＝控えの字面で入れ直す
                if _restore_formulas(ws, addr, src_snap) is None:
                    print(f"  ⚠ 控えが大きすぎて「{name}」の数式を写せませんでした"
                          "（外部リンクが残っていないか下の行を見てください）")
                back_to.append(f"{name}!{addr}")
                try:
                    ws.Activate()             # 図形を貼り戻すには対象シートが前に出ている必要がある
                except Exception:
                    pass
                try:
                    have = {str(s.Name) for s in ws.Shapes}
                except Exception:
                    have = set()
                for s in (shapes_by.get(name) or []):
                    nm = s['name']
                    try:
                        if nm not in have:                          # agent が消した図形＝控えから貼り戻す
                            src_ws.Shapes(nm).Copy()
                            ws.Paste()
                            ws.Shapes(int(ws.Shapes.Count)).Name = nm   # 貼った図形は最前面に来る
                            back += 1
                        shp = ws.Shapes(nm)
                        shp.Left, shp.Top, shp.Width, shp.Height = s['left'], s['top'], s['width'], s['height']
                    except Exception:
                        pass
                xl.CutCopyMode = False
            _drop_backup_link(wb, path)
            if back:
                print(f"  消えていた図形を控えから戻しました: {back} 個")
            try:
                ws = wb.Sheets(sheet)
                ws.Activate()                 # Select はアクティブシートでないと失敗する（戻し自体は済んでいる）
                ws.Range("A1").Select()
            except Exception:
                pass
    except Exception as ex:
        print(f"エラー: 戻せませんでした: {ex}")
        return False
    finally:
        if src_wb is not None:
            try:
                src_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if events is not None:
            try:
                xl.EnableEvents = events           # 元の値に戻す
            except Exception:
                pass
    print("戻しました: " + "・".join(back_to) + " を控えの状態に置き換え"
          + (f"（書き換わっていたセル {n_before} 個）" if n_before else "") + "・図形の位置も戻した")
    print("（保存はしていません。これでよければ保存、違えば保存せずに閉じる）")
    if meta.get('run_id'):
        # 走行台帳にその走行の印を付ける（2026-09-06）。人が戻した割合＝本番でいちばん正直な失敗率。
        # 台帳は追記だけ（書き換えない）＝読むときに畳む
        runs_append({'undo': meta['run_id'], 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
                     'book': book, 'sheets': sheets})
        print(f"（走行台帳に「人が戻した」の印を付けました: {meta['run_id']}　一覧は agent --runs）")
    return True


__all__ = ['_AGENT_BACKUP_KEEP', '_AGENT_BACKUP_MARK', '_AGENT_BACKUP_META_EXT', '_CHANGES_HEADER', '_CHANGES_MAX_ROWS', '_CREATED_PATTERNS', '_EMPTY_SNAPSHOT', '_FORMAT_COLS_MAX', '_HEAVY_OPS', '_LAST_AGENT_CHANGES_FILE', '_LAST_AGENT_UNDO_FILE', '_NUMLIKE_RE', '_READONLY_HEAVY_ACTIONS', '_READONLY_OPS', '_REPORT_ADDR_RE', '_REPORT_DID_RE', '_REPORT_FORMULA_RE', '_REPORT_RANGE_RE', '_REPORT_SAID_BLANK_RE', '_REPORT_SAID_BLANK_WINDOW', '_SNAPSHOT_MAX_CELLS', '_XL_ALIGN', '_XL_BORDERS', '_act_sheet', '_act_writes', '_agent_backup', '_backup_meta_path', '_cell_pair', '_change_side', '_changes_of', '_changes_table', '_clip', '_created_from_results', '_delete_created', '_diff_sheet', '_drop_backup_link', '_find_listobject', '_format_changes', '_format_snapshot', '_format_table', '_formula_rate', '_line_text', '_list_backups', '_load_undo_meta', '_mark_undo_stale', '_mixed', '_other_sheets_in', '_pick_backup', '_pick_shape', '_pivot_sheets_in', '_prune_agent_backups', '_record_created', '_report_addrs', '_report_did_section', '_restore_formulas', '_runs_after', '_same_path', '_save_undo_meta', '_shapes_geometry', '_sheet_snapshot', '_unbound_shape_names', '_undo_add_sheet', '_undo_shapes', '_undo_sheets', '_will_write', '_write_changes_file', 'backups_list', 'backups_prune', 'show_changes', 'undo_agent']
