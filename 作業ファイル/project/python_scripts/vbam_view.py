# -*- coding: utf-8 -*-
"""vbam_view.py — vba_manager 分割パート: 「目」コマンド（read-range/sheet-info/screenshot/snapshot 等）

vba_manager.py から機械分割（2026-07-12）。単体で実行せず、vba_manager.py 経由で使う。
"""
import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

from vbam_core import *  # noqa: F401,F403
from vbam_vba import *  # noqa: F401,F403
# ================================================================
# 「目」コマンド (シート状態の読み取り)
# ================================================================

LAST_VIEW_FILE = os.path.join(SCRIPT_DIR, '_last_view.png')   # screenshot の出力先




def _resolve_range(xl, wb, spec, sheet_name=None):
    """
    範囲指定を (ws, rng) に解決する。
      'A1:D20'         → アクティブシートの範囲
      'Sheet1!A1:D20'  → シート指定の範囲
      'Sheet1'         → そのシートの UsedRange
      None / ''        → アクティブシートの UsedRange
    sheet_name（--sheet オプション）が来たら spec はアドレスのみとして扱う。
    「シート!範囲」一本槍だと、'!' を含むシート名（Excelでは合法）や
    記号入り日本語シート名のクォートで詰むための分離指定の口。
    """
    if sheet_name:
        ws = None
        for sh in wb.Sheets:
            if sh.Name == sheet_name:
                ws = sh
                break
        if ws is None:
            raise Exception(f"シート '{sheet_name}' が見つかりません")
        if not spec:
            return ws, ws.UsedRange
        return ws, ws.Range(spec)

    if not spec:
        ws = wb.ActiveSheet
        return ws, ws.UsedRange

    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        # Excel の数式バー表記（'月次 集計'!A1）のクォートを剥がす（'' は ' に戻す）
        if len(sheet_part) >= 2 and sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        ws = wb.Sheets(sheet_part)
        if not addr:
            return ws, ws.UsedRange
        return ws, ws.Range(addr)

    # シート名そのものなら UsedRange
    for sh in wb.Sheets:
        if sh.Name == spec:
            return sh, sh.UsedRange

    # それ以外はアクティブシートのアドレスとして扱う
    ws = wb.ActiveSheet
    return ws, ws.Range(spec)


def _whole_sheet_spec(wb, spec, sheet_name=None):
    """spec がシート全域(UsedRange)に解決される形ならシート名を返す（破壊系コマンドのガード用）。

    「シート名だけ」「末尾!」「空」の spec は _resolve_range で UsedRange 全域になる。
    読み取り系では便利だが、clear/fill/sort/write 等の破壊系では
    範囲指定ミス1つで全域破壊になるため、明示指定(--whole-sheet)なしでは拒否する。
    """
    if sheet_name:
        return sheet_name if not spec else None
    if not spec:
        return wb.ActiveSheet.Name
    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        return sheet_part if not addr else None
    for sh in wb.Sheets:
        if sh.Name == spec:
            return spec
    return None






def _disp_width(s):
    """全角文字を2幅として数えた表示幅"""
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
    return w


def _disp_truncate(s, width):
    """表示幅 width に収まるよう切り詰める。

    切れたことが分かるよう末尾に '…' を付ける（黙って切ると、欠けた値を
    全文と誤読して write で書き戻す事故の芽になる）。全文が要るときは
    read-range --width で広げるか --tsv で書き出す。
    """
    if _disp_width(s) <= width:
        return s
    lim = max(width - 2, 1)     # '…' は全角幅2として確保
    out = []
    w = 0
    for ch in s:
        cw = 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
        if w + cw > lim:
            break
        out.append(ch)
        w += cw
    return ''.join(out) + '…'


def _disp_pad(s, width, right=False):
    """表示幅基準で width までスペース埋め（right=Trueで右寄せ）"""
    pad = width - _disp_width(s)
    if pad <= 0:
        return s
    return (' ' * pad + s) if right else (s + ' ' * pad)




def _merged_areas_in_range(rng, cap=8000):
    """range 内の結合セル領域を "A3:I8" 形式のアドレス一覧で返す（重複なし）。

    戻り値は (areas, skipped_total)。
      ・結合が1つも無い範囲は rng.MergeCells が False を返すので即 ([], None)（最速パス）。
      ・cap を超えるセル数の範囲は走査せず (None, total) を返す（巨大UsedRangeの暴走防止）。
    件数・秒での打ち切りは持たない（2026-09-16 夜に 24 件・0.5 秒で切る版が入ったが 9/17 に外した。
    列×行の交点だけ聞く走査で 結合 102 個の 名前変換 も 1.3 秒＝打ち切る理由が無く、切ると結合が黙って欠ける）。
    MergeArea の矩形は Python 側の seen 集合に算術で畳み、余分な COM 呼び出しを避ける。
    神エクセルの二重構造（文字はA1・見た目はI列まで結合）を読み解くための素。
    """
    try:
        if rng.MergeCells is False:      # 範囲内に結合が皆無 → 走査不要
            return [], None
    except Exception:
        pass                              # None(混在)や例外は通常走査へ
    try:
        total = int(rng.Cells.Count)
    except Exception:
        return [], None
    if total > cap:
        return None, total
    ws = rng.Worksheet
    r0, c0 = rng.Row, rng.Column
    nr, nc = rng.Rows.Count, rng.Columns.Count
    seen = set()
    areas = []
    # 列も先に 1 回ずつ聞く（False＝この列に結合は無い）。行だけで飛ばすと、結合のある行は全列を 1 セルずつ
    # 聞いていた（ポスターの 名前変換: 102 行 × 24 列＝2,448 往復・12 秒無言・2026-09-17）。
    # 行×列の交点だけ聞けば C〜E の 3 列 × 102 行＝306 往復で済む
    cols = []
    for j in range(nc):
        try:
            if ws.Columns(c0 + j).MergeCells is False:
                continue
        except Exception:
            pass
        cols.append(j)
    if not cols:
        return [], None
    # 行は 16 行ずつのブロックでまず聞き（"13:28" の 1 往復）、結合の無いブロックは行ごとに聞かない。
    # 1 往復 3.7 ミリ秒（2026-09-17 実測）なので、130 行に結合 1 つの 図形 シートで 130 往復→9＋16 往復
    rows = []
    for b0 in range(0, nr, 16):
        b1 = min(b0 + 16, nr)
        try:
            if ws.Rows(f"{r0 + b0}:{r0 + b1 - 1}").MergeCells is False:
                continue
        except Exception:
            pass
        for i in range(b0, b1):
            try:
                # 行ごとに 1 回聞く（False＝この行に結合は無い）。セル 1 つずつ聞くと 8,000 セルで 8,000 往復
                # だった（2026-09-16）。結合は使用範囲の中にしか無いので、行全体で無ければ範囲の中にも無い
                if ws.Rows(r0 + i).MergeCells is False:
                    continue
            except Exception:
                pass
            rows.append(i)
    for i in rows:
        for j in cols:
            key = (r0 + i, c0 + j)
            if key in seen:
                continue
            try:
                cell = ws.Cells(r0 + i, c0 + j)
                # 結合の矩形は MergeArea.Address 1 回で取る（MergeCells＋Row・Column・Rows.Count・Columns.Count の
                # 6 往復 → 2 往復・2026-09-17）。結合が無ければ MergeArea はそのセル自身＝Address に ':' が無い。
                # win32com の動的ディスパッチでは Address はプロパティ（"$C$13:$E$13"）。Address(False, False) と
                # 呼ぶと 'str' object is not callable で毎回フォールバックしていた（同日実測）
                ma = cell.MergeArea
                rect = None
                try:
                    m = _A1_RECT_RE.match(str(ma.Address))
                    if m and not m.group(3):
                        seen.add(key)                  # 結合なし
                        continue
                    if m:
                        rect = (int(m.group(2)), _col_num_of(m.group(1)), int(m.group(4)), _col_num_of(m.group(3)))
                except Exception:
                    rect = None
                if rect is None:
                    if not cell.MergeCells:
                        seen.add(key)
                        continue
                    mr, mc = ma.Row, ma.Column
                    rect = (mr, mc, mr + ma.Rows.Count - 1, mc + ma.Columns.Count - 1)
                mr, mc, mr2, mc2 = rect
                areas.append(f"{_col_letter(mc)}{mr}:{_col_letter(mc2)}{mr2}")
                for ii in range(mr, mr2 + 1):            # 結合矩形を丸ごと走査済みにする
                    for jj in range(mc, mc2 + 1):
                        seen.add((ii, jj))
            except Exception:
                seen.add(key)
    return areas, None


_A1_RECT_RE = re.compile(r'^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$')


def _col_num_of(letters):
    """列文字（A, AA）→ 列番号（1 始まり）。"""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _values_to_grid(rng, use_formula=False, max_col_width=40):
    """Range の値を、列文字＋行番号つきのテキスト格子にする

    use_formula=True のときは計算結果ではなく数式(.Formula)を表示する。
    数式のないセルは定数値がそのまま入る（write-range の .Value と同じ規約）。
    max_col_width を超える列は '…' 付きで切り詰める（--width で変更可）。
    """
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return "(空の範囲です)"

    # 単一セル
    if not isinstance(raw, tuple):
        a1 = f"{_col_letter(rng.Column)}{rng.Row}"
        return f"{a1}: {_cell_str(raw)}"

    # tuple-of-tuples へ正規化
    rows = []
    for row in raw:
        rows.append(list(row) if isinstance(row, tuple) else [row])
    if not rows:
        return "(空の範囲です)"

    start_row = rng.Row
    start_col = rng.Column
    ncols = max(len(r) for r in rows)

    str_rows = [[_cell_str(v) for v in row] + [''] * (ncols - len(row)) for row in rows]
    headers = [_col_letter(start_col + j) for j in range(ncols)]

    rownum_w = len(str(start_row + len(str_rows) - 1))
    col_w = []
    for j in range(ncols):
        w = _disp_width(headers[j])
        for i in range(len(str_rows)):
            w = max(w, _disp_width(str_rows[i][j]))
        col_w.append(min(w, max_col_width))

    def fmt_row(cells, label):
        parts = [_disp_pad(label, rownum_w, right=True)]
        for j, c in enumerate(cells):
            parts.append(_disp_pad(_disp_truncate(c, col_w[j]), col_w[j]))
        return ' | '.join(parts)

    out = [fmt_row(headers, ''),
           '-' * (rownum_w + sum(col_w) + 3 * ncols)]
    for i, row in enumerate(str_rows):
        out.append(fmt_row(row, str(start_row + i)))
    return '\n'.join(out)


def _text_to_grid(rng, max_col_width=40, max_rows=100, max_cols=30):
    """Range の「画面に見えている文字」(.Text) を列文字＋行番号つきの格子にする。

    _values_to_grid が .Value（中身）を見せるのに対し、こちらは表示形式と列幅を
    通した後の文字。列幅が足りず '####' になっているセルは値を見ても分からない
    （Excelコンボ「表示結果」から輸入・2026-09-02）。
    戻り値は (格子, ### のセル数, 切り詰めの注記)。100行×30列で切る。
    """
    nrows = int(rng.Rows.Count)
    ncols = int(rng.Columns.Count)
    note = ""
    if nrows > max_rows:
        nrows = max_rows
        note += f"（先頭{max_rows}行まで）"
    if ncols > max_cols:
        ncols = max_cols
        note += f"（左{max_cols}列まで）"
    start_row = int(rng.Row)
    start_col = int(rng.Column)
    hashes = 0
    str_rows = []
    for i in range(1, nrows + 1):
        row = []
        for j in range(1, ncols + 1):
            try:
                tx = rng.Cells(i, j).Text
            except Exception:
                tx = ""
            tx = "" if tx is None else str(tx)
            # 「####」だけのセル＝列幅が足りず数が読めていない（#N/A 等は文字が残るので当たらない）
            if tx and tx.replace("#", "") == "":
                hashes += 1
            row.append(tx.replace("\r", "").replace("\n", " "))
        str_rows.append(row)
    headers = [_col_letter(start_col + j) for j in range(ncols)]
    rownum_w = len(str(start_row + nrows - 1))
    col_w = []
    for j in range(ncols):
        w = _disp_width(headers[j])
        for i in range(nrows):
            w = max(w, _disp_width(str_rows[i][j]))
        col_w.append(min(w, max_col_width))

    def fmt_row(cells, label):
        parts = [_disp_pad(label, rownum_w, right=True)]
        for j, c in enumerate(cells):
            parts.append(_disp_pad(_disp_truncate(c, col_w[j]), col_w[j]))
        return ' | '.join(parts)

    out = [fmt_row(headers, ''),
           '-' * (rownum_w + sum(col_w) + 3 * ncols)]
    for i, row in enumerate(str_rows):
        out.append(fmt_row(row, str(start_row + i)))
    return '\n'.join(out), hashes, note


def _hash_scan(rng, max_rows=100, max_cols=30):
    """列幅が足りず '###' になっているセルを、範囲を 1 回で読んでから怪しいセルだけ Excel に聞く。

    .Text は範囲ごとには読めず 1 セルずつ＝78 セルで 30 秒かかっていた（2026-09-13・Antigravity の
    Gemini が見つけて一括読みに替えた。### の検査はそのとき落ちていたので、ここで戻す）。
    ### になるのは数値と日付だけなので、.Value を 1 回で読み、表示の幅が列幅に届きそうなセルだけ
    .Text で確かめる。列幅が足りている表なら Excel への問い合わせは列の数だけで済む。
    戻り値は (### のセル数, 番地の先頭 5 個, 切り詰めの注記)。
    """
    import datetime as _dt
    nrows = int(rng.Rows.Count)
    ncols = int(rng.Columns.Count)
    note = ""
    if nrows > max_rows:
        nrows = max_rows
        note += f"（先頭{max_rows}行まで）"
    if ncols > max_cols:
        ncols = max_cols
        note += f"（左{max_cols}列まで）"
    sub = rng if (nrows == int(rng.Rows.Count) and ncols == int(rng.Columns.Count)) \
        else rng.Worksheet.Range(rng.Cells(1, 1), rng.Cells(nrows, ncols))
    raw = sub.Value
    if raw is None:
        return 0, [], note
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    rows = [list(r) if isinstance(r, tuple) else [r] for r in raw]
    widths = []
    for j in range(1, ncols + 1):
        try:
            widths.append(float(sub.Columns(j).ColumnWidth))
        except Exception:
            widths.append(8.43)
    hashes, bad = 0, []
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, bool) or v is None or isinstance(v, str):
                continue                                       # 文字は ### にならない
            if isinstance(v, _dt.datetime):
                est = 10 if (v.hour == 0 and v.minute == 0 and v.second == 0) else 16
            elif isinstance(v, (int, float)):
                digits = len(str(int(abs(v)))) if abs(v) < 1e15 else 16
                est = digits + digits // 3 + 3                 # 桁区切り・符号・小数の分を見込む
            else:
                continue
            if est + 1 < widths[j - 1]:
                continue                                       # 列幅に余裕＝聞かなくて分かる
            try:
                t = str(sub.Cells(i, j).Text or "")
            except Exception:
                continue
            if t and t.replace("#", "") == "":
                hashes += 1
                if len(bad) < 5:
                    bad.append(f"{_col_letter(int(sub.Column) + j - 1)}{int(sub.Row) + i - 1}")
    return hashes, bad, note


def _show_range(ws, rng, label="表示（画面に見えている文字）"):
    """書いた／整えた直後に、その範囲の見え方を出す（＝ループ）。

    書いた本人には、表示形式の効き目や列幅不足の ### は見えない。ここで見せないと
    「変更しました」と言い続ける（Excelコンボの「書式読み戻し」に相当・2026-09-02）。
    値は範囲を 1 回で読み、### は _hash_scan で怪しいセルだけ確かめる（1 セルずつ .Text を
    読んでいた 2026-09-02〜13 の形は 78 セルで 30 秒かかった）。
    """
    try:
        grid = _values_to_grid(rng)
    except Exception as e:
        print(f"（表示を読めませんでした: {e}）")
        return
    print(f"{label} {ws.Name}!{rng.Address.replace('$', '')}:")
    print(grid)
    try:
        hashes, bad, _note = _hash_scan(rng)
    except Exception:
        return
    if hashes:
        print(f"※ '###' で読めないセルが {hashes} 個あります（{' '.join(bad)}）"
              f"（列幅が足りないか表示形式の問題。列幅を広げるか表示形式を見直してください）")


def cmd_read_range(args):
    """シートのセル値をテキスト格子で読み取る（目・テキスト版）。

    複数範囲可（1回のCOM接続でまとめ読み）。--tsv で _last_values.tsv に
    書き出せば「読む→TSVを編集→write-range で書き戻す」の往復が
    get→_last_proc.vba→replace-procedure と同じ型になる。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    specs = rest if rest else [None]
    use_formula = getattr(args, 'formula', False)
    try:
        # `or 40` だと --width 0 が偽値で既定値に化ける（指定が無言で消える）ため is None 判定
        w_opt = getattr(args, 'width', None)
        width = 40 if w_opt is None else int(w_opt)
    except (TypeError, ValueError):
        print("エラー: --width は数値で指定してください")
        return False
    if width < 1:
        print("エラー: --width は 1 以上で指定してください")
        return False
    tsv_out = getattr(args, 'tsv_out', None)
    if tsv_out is not None and len(specs) > 1:
        print("エラー: --tsv は範囲1つのときだけ使えます")
        return False

    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    blocks = [(_resolve_range(xl, wb, spec, sheet_opt)) for spec in specs]

    if getattr(args, 'json', False):
        import json
        out = []
        for ws, rng in blocks:
            rows = _range_values_2d(rng, use_formula)
            entry = {"sheet": ws.Name, "address": rng.Address,
                     "ref": f"{ws.Name}!{rng.Address}",
                     "rows": [[_cell_str(v) for v in r] for r in rows]}
            areas, skipped = _merged_areas_in_range(rng)
            if skipped:
                entry["merged_skipped_cells"] = skipped   # cap超で未走査
            elif areas:
                entry["merged"] = areas
            out.append(entry)
        print(json.dumps({"success": True, "file": wb.Name, "ranges": out},
                         ensure_ascii=False), file=sys.stdout)
        return True

    for ws, rng in blocks:
        mode = "（数式表示）" if use_formula else ""
        # 末尾の [シート名!番地] はそのまま次コマンドの range 引数に貼れる形
        print(f"シート: {ws.Name}   範囲: {rng.Address}{mode}   [{ws.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng, use_formula, max_col_width=width))
        print("=" * 60)
        areas, skipped = _merged_areas_in_range(rng)
        if skipped:
            print(f"結合セル: 未走査（{skipped}セルはcap超・範囲を絞れば表示）")
        elif areas:
            shown = areas[:30]
            more = f"  …他{len(areas) - 30}件" if len(areas) > 30 else ""
            print(f"結合セル {len(areas)}件: " + ", ".join(shown) + more)

    if tsv_out is not None:
        path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
        ws, rng = blocks[0]
        rows = _range_values_2d(rng, use_formula)
        # セル内改行(Alt+Enter)・タブは TSV の行/列区切りと衝突し、
        # そのまま write-range で書き戻すと格子がずれて無警告のデータ破壊になる。
        # 値は変えず、該当セルを名指しで警告する
        dirty = []
        for ri, r in enumerate(rows):
            for ci, v in enumerate(r):
                if isinstance(v, str) and ('\n' in v or '\r' in v or '\t' in v):
                    dirty.append(f"{_col_letter(rng.Column + ci)}{rng.Row + ri}")
        if dirty:
            shown = ", ".join(dirty[:8]) + ("" if len(dirty) <= 8 else f" …他{len(dirty) - 8}件")
            print(f"⚠ 警告: セル内に改行/タブを含むセルが {len(dirty)}件あります: {shown}")
            print("  このTSVをそのまま write-range で書き戻すと行・列がずれます。")
            print("  該当セルは手で編集するか、write-range の対象から外してください。")
        lines = ['\t'.join(_cell_str(v) for v in r) for r in rows]
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            f.write('\n'.join(lines) + '\n')
        print(f"TSV書き出し: {path}  ({len(rows)}行 x {max(len(r) for r in rows)}列)")
        print(f"  編集後の書き戻し: py vba_manager.py write-range \"{ws.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
        print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
    return True


def cmd_read_selection(args):
    """ユーザーが今選択している範囲を読み取る"""
    target_file, _ = parse_target_and_rest(args.posargs)
    use_formula = getattr(args, 'formula', False)
    xl, wb = get_workbook(target_file)
    sel = xl.Selection
    if sel is None:
        print("選択範囲がありません。")
        return False
    try:
        wsname = sel.Worksheet.Name
    except Exception:
        wsname = "(不明)"
    try:
        addr = sel.Address
    except Exception:
        addr = "(範囲ではありません)"
    mode = "（数式表示）" if use_formula else ""
    print(f"シート: {wsname}   選択範囲: {addr}{mode}")
    print("=" * 60)
    try:
        print(_values_to_grid(sel, use_formula))
    except Exception as e:
        print(f"(値を読めませんでした: {e})")
    print("=" * 60)
    return True


def cmd_sheet_info(args):
    """ブックのシート構成・使用範囲を表示（見取り図）。

    --preview N で各シート使用範囲の先頭N行も格子表示（初見ブックの俯瞰が
    1コマンド1接続で済む。従来は sheet-info + シート毎の read-range で N+1 接続）。
    """
    target_file, _ = parse_target_and_rest(args.posargs)
    try:
        preview = int(getattr(args, 'preview', None) or 0)
    except (TypeError, ValueError):
        print("エラー: --preview は数値で指定してください")
        return False
    xl, wb = get_workbook(target_file)
    active = wb.ActiveSheet.Name
    print(f"ブック: {wb.Name}")
    print(f"シート数: {wb.Sheets.Count}   アクティブ: {active}")
    print("-" * 60)
    fast = getattr(args, 'fast', False)
    for sh in wb.Sheets:
        mark = '*' if sh.Name == active else ' '
        try:
            ur = sh.UsedRange
            dims = f"{ur.Rows.Count}行 x {ur.Columns.Count}列  ({ur.Address})"
        except Exception:
            ur = None
            dims = "(空)"
        vis = '' if sh.Visible == -1 else '  [非表示]'
        print(f"{mark} {sh.Name}: {dims}{vis}")
        if ur is not None and not fast:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                print(f"    結合: 未走査（{skipped}セル・大）" if isinstance(skipped, int) else f"    結合: {skipped}")
            elif areas:
                shown = areas[:12]
                more = f"  …他{len(areas) - 12}件" if len(areas) > 12 else ""
                print(f"    結合 {len(areas)}件: " + ", ".join(shown) + more)
        if preview > 0 and ur is not None:
            try:
                nrows = min(preview, ur.Rows.Count)
                head = ur.Worksheet.Range(ur.Cells(1, 1), ur.Cells(nrows, ur.Columns.Count))
                grid = _values_to_grid(head)
                print('    ' + grid.replace('\n', '\n    '))
            except Exception as e:
                print(f"    (先頭行を読めませんでした: {e})")
    print("-" * 60)
    return True


def _print_long_cells(rng, width=40, limit=10, max_chars=400):
    """格子で '…' に切れた長い文字セルを全文で出す。

    指示文・説明文が A3 等に長く入っている表向け（2026-09-02 夜・A3 の指示を読むために
    read-range をもう1往復していた）。_values_to_grid の切り詰め幅（既定 40）を超える
    文字列セルだけを拾う。
    """
    try:
        raw = rng.Value
    except Exception:
        return
    if raw is None:
        return
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    r0, c0 = int(rng.Row), int(rng.Column)
    hits = []
    for i, row in enumerate(raw):
        if not isinstance(row, tuple):
            row = (row,)
        for j, v in enumerate(row):
            if isinstance(v, str) and _disp_width(v) > width:
                hits.append((f"{_col_letter(c0 + j)}{r0 + i}", v))
    if not hits:
        return
    print("長文セル（格子で切れた分の全文）:")
    for addr, v in hits[:limit]:
        t = v.replace("\r", "").replace("\n", " ")
        if len(t) > max_chars:
            t = t[:max_chars] + "…"
        print(f"  {addr}: {t}")
    if len(hits) > limit:
        print(f"  …他{len(hits) - limit}件")


# ---- 気づき（表の汚れ）2026-09-07 ----
# materials の格子は、空白の幅・日付の型・半角カナを見せない。見えない分を頭で埋めようとして
# read-range と --raw を 2 往復挟み、書き始めるまで materials から 100 秒（道具は合計数秒）。
# 「直す前に壊れ方を突き止める」癖を、道具が数えて出す形で消す。掃除の手はどのみち「きれいな値で
# 上書き」なので、ここに出た番地は「直すかどうかの材料」であって、書く前の調べ物ではない。
_ZEN_DIGIT_RE = re.compile(r'[０-９]')
_HAN_KANA_RE = re.compile(r'[｡-ﾟ]')
_NUMLIKE_STR_RE = re.compile(r'^[+\-−]?[\d０-９]{1,3}([,，][\d０-９]{3})+([.．][\d０-９]+)?$'
                             r'|^[+\-−]?[\d０-９]+([.．][\d０-９]+)?$')
_DATELIKE_STR_RE = re.compile(r'^(\d{4}[/.\-年]\s?\d{1,2}[/.\-月]\s?\d{1,2}日?'
                              r'|[RHSTMrhstm]\d{1,2}[./\-]\d{1,2}[./\-]\d{1,2}'
                              r'|(令和|平成|昭和|大正|明治)\s?\d{1,2}年\s?\d{1,2}月\s?\d{1,2}日'
                              r'|\d{1,2}/\d{1,2}/\d{4})$')
_ID_HEAD_RE = re.compile(r'番号|コード|ｺｰﾄﾞ|^ID$|^No\.?$|№', re.I)


def _blank_cell(v):
    return v is None or (isinstance(v, str) and v.strip(' 　 \t\r\n') == '')


def _grid_of_value(val):
    """Range.Value（単セル＝スカラー／複数＝tuple の tuple）を 2 次元リストにする。"""
    if val is None or not isinstance(val, tuple):
        return [[val]]
    return [list(r) if isinstance(r, tuple) else [r] for r in val]


def _guess_header_idx(grid):
    """格子の見出し行（0 始まり）を推定。文字だけのセルが 2 つ以上あり、次の行にも値が 2 つ以上ある最初の行。"""
    for i in range(len(grid) - 1):
        filled = [v for v in grid[i] if not _blank_cell(v)]
        if len(filled) >= 2 and all(isinstance(v, str) for v in filled):
            if len([v for v in grid[i + 1] if not _blank_cell(v)]) >= 2:
                return i
    return None


def _date_col_formats(ws, grid, r0, c0, header_idx):
    """日付型の列ごとに表示形式の集合を集める。列まとめての NumberFormat が None（＝混在）のときだけセル単位で拾う。"""
    out = {}
    h = -1 if header_idx is None else header_idx
    width = max((len(r) for r in grid), default=0)
    for j in range(width):
        idx = [i for i in range(h + 1, len(grid))
               if j < len(grid[i]) and isinstance(grid[i][j], datetime.datetime)]
        if len(idx) < 2:
            continue
        try:
            f = ws.Range(ws.Cells(r0 + idx[0], c0 + j), ws.Cells(r0 + idx[-1], c0 + j)).NumberFormat
        except Exception:
            continue
        if f is not None:
            continue
        fmts = set()
        for i in idx[:60]:
            try:
                fmts.add(str(ws.Cells(r0 + i, c0 + j).NumberFormat))
            except Exception:
                pass
        if len(fmts) > 1:
            out[j] = fmts
    return out


def _body_look_mixed(ws, grid, r0, c0, header_idx):
    """見出しの下の本文で、セルごとに違う文字の書式 → (番地, [性質の名前])。範囲まとめての COM が None＝混在。

    材料が値しか見ておらず、文字の色・大きさ・太字・斜体・下線・取り消し線・塗りがばらばらの表を
    AI が知らずに「合格」まで行った（2026-09-11 動画のお題の撃ち直し）。
    """
    h = -1 if header_idx is None else header_idx
    rows = grid or []
    filled = [i for i in range(h + 1, len(rows)) if not all(_blank_cell(v) for v in rows[i])]
    width = max((len(r) for r in rows), default=0)
    if len(filled) < 2 or width == 0:
        return None, []
    rng = ws.Range(ws.Cells(r0 + filled[0], c0), ws.Cells(r0 + filled[-1], c0 + width - 1))
    f = rng.Font
    bad = []
    for name, v in (("フォント名", f.Name), ("大きさ", f.Size), ("文字色", f.Color), ("太字", f.Bold),
                    ("斜体", f.Italic), ("下線", f.Underline), ("取り消し線", f.Strikethrough),
                    ("塗り", rng.Interior.ColorIndex)):
        if v is None:
            bad.append(name)
    return str(rng.Address).replace('$', ''), bad


def dirt_notes(grid, r0=1, c0=1, header_idx=None, col_formats=None, limit=12, look_mixed=None):
    """表の汚れを数えて「気づき」の行にする（純 Python。COM は呼ばない）。

    grid: 値の 2 次元リスト（UsedRange.Value を _grid_of_value した形）。r0/c0: 左上のシート上の行・列番号。
    header_idx: 見出し行（grid の 0 始まり）。None なら推定。見出しの下の行だけを見る。
    col_formats: {列index: set(表示形式)} 日付列の表示形式（呼び手が _date_col_formats で集める）。
    look_mixed: (番地, [性質の名前]) 本文の文字の書式の混在（呼び手が _body_look_mixed で集める）。
    出すもの: 重複行／番号列の重複／全角数字／半角カナ／空白の乱れ（端・空白だけ・連続・列の中で全角と
    半角の区切りが混在）／文字の数字（番号列は除く）／文字の日付／日付の表示形式の混在／表の中の空欄。
    """
    rows = [list(r) for r in (grid or [])]
    if not rows:
        return []
    h = header_idx if header_idx is not None else _guess_header_idx(rows)
    if h is None:
        h = -1
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([None] * (width - len(r)))
    heads = rows[h] if h >= 0 else [None] * width
    id_cols = {j for j, v in enumerate(heads) if isinstance(v, str) and _ID_HEAD_RE.search(v.strip())}
    body = [i for i in range(h + 1, len(rows)) if not all(_blank_cell(v) for v in rows[i])]

    def addr(i, j):
        return f"{_col_letter(c0 + j)}{r0 + i}"

    def norm(v):
        if isinstance(v, str):
            return re.sub(r'\s+', ' ', unicodedata.normalize('NFKC', v)).strip()
        if isinstance(v, float) and v.is_integer():
            return int(v)
        return v

    def show(lst, n=limit):
        return " ".join(lst[:n]) + ("…" if len(lst) > n else "")

    out = []
    dup_rows, dup_row_set = [], set()
    # 重複行は「同じ相手か」で照合する（2026-09-11: 値が丸ごと同じ行だけ数えていて、51 行の名簿で 26 組中 3 組しか
    # 拾えなかった。残りを AI が自分で数えて、同じ会社の 2 行を両方消した）
    from vbam_hands import _dup_pairs
    for i, first, exact in _dup_pairs(rows, h):
        head_val = next((str(norm(v)) for v in rows[i] if not _blank_cell(v)), '')
        dup_rows.append(f"行{r0 + i} = 行{r0 + first}（{head_val}{'' if exact else '・表記ゆれあり'}）")
        dup_row_set.add(i)
    if dup_rows:
        lim = max(limit, 60)
        out.append("  重複行: " + "／".join(dup_rows[:lim]) + ("…" if len(dup_rows) > lim else "")
                   + (f"（計 {len(dup_rows)} 行。全部の列が同じ行＝空白・全角半角・大文字小文字・ハイフン・（株）の違いは"
                      "ならして照合。消すなら dedupe の手）" if len(dup_rows) > 1 else ""))
    # 重複の疑い（2026-09-12・全列一致が基本）: 同じ相手に見えるが 1〜2 列だけ違う組＝重複ではない・消さない
    try:
        from vbam_hands import _near_dup_pairs
        near = _near_dup_pairs(rows, h)
    except Exception:
        near = []
    if near:
        def _hn(j):
            v = heads[j] if j < len(heads) else None
            return str(v).strip() if not _blank_cell(v) else _col_letter(c0 + j)
        out.append("  重複の疑い（全部の列は同じでない＝重複ではない・消さない。どちらを残すか・まとめるかは人の判断）: "
                   + "／".join(f"行{r0 + i} と 行{r0 + k}（{'・'.join(_hn(j) for j in diff) or '書き方'}が違う）"
                              for i, k, diff in near[:15]) + (f" ほか {len(near) - 15} 組" if len(near) > 15 else ""))
    dup_keys = []
    for j in sorted(id_cols):
        seen = {}
        for i in body:
            v = rows[i][j]
            if _blank_cell(v):
                continue
            k = str(norm(v))
            if k in seen:
                if i not in dup_row_set:
                    dup_keys.append(f"{addr(i, j)} = {addr(seen[k], j)}（{k}）")
            else:
                seen[k] = i
    if dup_keys:
        out.append("  番号列の重複: " + show(dup_keys))
    zen, han, edge, only_ws, multi, numstr, datestr, sep = [], [], [], [], [], [], [], {}
    astext = []          # 文字として入った式（'=' で始まる文字。計算されていない・2026-09-09）
    for i in body:
        for j in range(width):
            v = rows[i][j]
            if not isinstance(v, str) or v == '':
                continue
            a = addr(i, j)
            core = v.strip(' 　 \t\r\n')
            if core == '':
                only_ws.append(a)
                continue
            if core != v:
                edge.append(a)
            if re.search(r'[ 　]{2,}', core):
                multi.append(a)
            if _ZEN_DIGIT_RE.search(core):
                zen.append(a)
            if _HAN_KANA_RE.search(core):
                han.append(a)
            if core.startswith('=') and len(core) > 1:
                astext.append(a)
            if _NUMLIKE_STR_RE.match(core):
                if j not in id_cols:
                    numstr.append(a)
            elif _DATELIKE_STR_RE.match(core):
                datestr.append(a)
            if ' ' in core or '　' in core:
                d = sep.setdefault(j, {'half': [], 'full': []})
                if ' ' in core:
                    d['half'].append(a)
                if '　' in core:
                    d['full'].append(a)
    if astext:
        out.append("  式が文字として入っているセル（計算されていません）: " + show(astext)
                   + "  → 表示形式を標準にしてから入れ直す")
    if zen:
        out.append("  全角数字: " + show(zen))
    if han:
        out.append("  半角カナ: " + show(han))
    sp = []
    if edge:
        sp.append("端に空白 " + show(edge))
    if only_ws:
        sp.append("空白だけ " + show(only_ws))
    if multi:
        sp.append("連続 " + show(multi))
    for j in sorted(sep):
        d = sep[j]
        if d['half'] and d['full']:
            minority, name = ((d['half'], '半角') if len(d['half']) <= len(d['full']) else (d['full'], '全角'))
            sp.append(f"{_col_letter(c0 + j)}列は全角と半角の区切りが混在（{name} {show(minority, 8)}）")
    if sp:
        out.append("  空白の乱れ: " + "／".join(sp))
    if numstr:
        out.append("  文字の数字（番号列は除く）: " + show(numstr))
    if datestr:
        out.append("  文字の日付: " + show(datestr))
    # 電話・郵便の列の書き方（2026-09-11 夜・試験: 空白区切り・かっこ・〒・区切りなしが気づきに出ず、AI は hankaku しか
    # 当てずに残した。規則の名前まで出す＝1 往復目で phone／postal を選べる）
    for j in range(width):
        hd = heads[j] if h >= 0 else None
        if not isinstance(hd, str):
            continue
        kind = 'phone' if re.search(r'電話|TEL|携帯|FAX', hd, re.I) else ('postal' if re.search(r'郵便|〒', hd) else None)
        if not kind:
            continue
        ok = r'0\d{1,4}-\d{1,4}-\d{3,4}' if kind == 'phone' else r'\d{3}-\d{4}'
        odd = [addr(i, j) for i in body if not _blank_cell(rows[i][j])
               and not (isinstance(rows[i][j], str) and re.fullmatch(ok, rows[i][j]))]
        if odd:
            out.append(f"  {'電話' if kind == 'phone' else '郵便番号'}の書き方がそろっていない（{_col_letter(c0 + j)}列 "
                       f"{len(odd)} セル・区切りなし・空白・かっこ・〒・全角・数値）: {show(odd)}  → normalize の {kind}")
    # 会社名の（株）と株式会社・種類の前後の空白（2026-09-11 夜・試験）
    for j in range(width):
        cells = [(i, rows[i][j]) for i in body if isinstance(rows[i][j], str)]
        full = [i for i, v in cells if re.search(r'株式会社|有限会社', v)]
        abbr = [i for i, v in cells if re.search(r'[（(]\s*[株有]\s*[）)]|[㈱㈲]', v)]
        spaced = [i for i, v in cells if re.search(r'(株式会社|有限会社)[ 　]|[ 　](株式会社|有限会社)', v.strip(' 　'))]
        if (abbr and full) or (spaced and len(spaced) < len(full)):
            bad = sorted(set(abbr) | set(spaced))
            out.append(f"  会社名の書き方がそろっていない（{_col_letter(c0 + j)}列・（株）と株式会社の混在 {len(abbr)}・"
                       f"種類の前後の空白 {len(spaced)}）: {show([addr(i, j) for i in bad])}  → normalize の corp")
    for j in sorted(col_formats or {}):
        fmts = sorted(str(f) for f in col_formats[j])
        if len(fmts) > 1:
            out.append(f"  日付の表示形式が混在: {_col_letter(c0 + j)}列（{', '.join(fmts[:4])}）")
    if look_mixed and look_mixed[1]:
        out.append(f"  文字の書式がセルごとに違う（本文 {look_mixed[0]}）: {'・'.join(look_mixed[1])}"
                   "  → format で本文全体に font・size・color と \"plain\": true・\"bg\": \"none\"・\"unbold\": true")
    blanks = []
    if h >= 0 and body:
        for j in range(width):
            if _blank_cell(heads[j]):
                continue
            cells = [(i, rows[i][j]) for i in body]
            filled = sum(1 for _i, v in cells if not _blank_cell(v))
            if filled < 3 or filled / len(cells) < 0.6:
                continue
            blanks += [addr(i, j) for i, v in cells if _blank_cell(v)]
    if blanks:
        out.append("  表の中の空欄: " + show(blanks))
    if not out:
        return ["気づき（表の汚れ）: なし（重複・全角数字・半角カナ・空白の乱れ・文字の数字/日付・空欄・"
                "文字として入った式、いずれも無し）"]
    return ["気づき（表の汚れ。直すかは指示しだい。掃除の手は「きれいな値で上書き」＝ここを見てから調べ直さない）:"] + out


# ----------------------------------------------------------------
# 数式の目（2026-09-09）。式の「中身」を機械で見る。
#
# それまで道具が見ていたのは値だけだった（audit_content の合計不一致・エラー値）。
# 「同じ列で 1 つだけ形の違う式」「式に直接書かれた率・単価」「集計の範囲が本文を取りこぼす」は
# 手順書「表と数式の点検」の ⑥⑦⑧ に文章では書いてあるが、探すのは AI の目視に任せていた
# ＝ 161 列の表では読み残しが出る（読み残しの関所を足した理由がそれ）。ここで機械の仕事にする。
# 純 Python（COM に触らない）＝ pytest で回せる。
# ----------------------------------------------------------------
_FML_STR_RE = re.compile(r'"(?:[^"]|"")*"')
# 範囲・セル参照。'集計'!B2 のようにシート名が付くものは対象外（前が ! なら拾わない）。
# LOG10( のような関数名を参照と読まないよう、後ろが ( のものは外す
_FML_RANGE_RE = re.compile(
    r'(?<![A-Za-z0-9_.$!])(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![A-Za-z0-9_(])')
_FML_NUM_RE = re.compile(r'(?<![A-Za-z0-9_.$@])(\d+(?:\.\d+)?)(?![A-Za-z0-9_.])')
# 別シート・別ブックの参照（'集計'!B2:B9）は丸ごと伏せる。前半だけ外して後半を拾う事故を防ぐ
_FML_SHEETREF_RE = re.compile(
    r"(?:'[^']*'|[^\s!(),*/+\-^&<>=:\";]+)!\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?")
_FML_OPS = set('*/+-^><=')          # この演算子に隣り合う数だけ「直書き」と数える（VLOOKUP の列番号・ROUND の桁は除く）
_FML_AGG_RE = re.compile(
    r'(?<![A-Za-z0-9_.])(SUM|SUBTOTAL|AVERAGE|COUNT|COUNTA|MAX|MIN|MEDIAN|PRODUCT)\s*\(', re.IGNORECASE)
_FML_AGG_ANY_RE = re.compile(
    r'(?<![A-Za-z0-9_.])(SUMPRODUCT|SUMIFS|SUMIF|SUBTOTAL|AVERAGEIFS|AVERAGEIF|AVERAGE|COUNTIFS|COUNTIF|COUNTA'
    r'|COUNT|AGGREGATE|MEDIAN|PRODUCT|SUM|MAX|MIN)\s*\(', re.IGNORECASE)
_FML_LENS_MAX_CELLS = 20000         # materials で式を丸ごと読む上限（大きいシートは数えない）
_FML_BULK_MAX_CELLS = 200000        # materials で式を一括で読む上限（92,228 セルの Value が 0.20 秒の実測から・2026-09-16）
_HIDDEN_COUNT_MAX = 5000            # 非表示を 1 本ずつ数える上限（混在のときだけ数える）


def _formula_patterns(fa, fr, r0, c0, cap=5000):
    """一括で読んだ式の格子（A1 形式 fa・R1C1 形式 fr）から、数式の本数と型を数える（純 Python・2026-09-16）。

    戻り値: (本数, {R1C1 の式: [個数, 代表の番地, 代表の A1 式, 最後の番地]})。cap を超えたら数えるのをやめる
    （本数は cap+1 で返る＝呼び手が '+' を付ける）。それまでは数式セル 1 つにつき COM を 3 回
    （FormulaR1C1・Address・Formula）呼んでいた＝5,000 セルで 15,000 往復。
    """
    pats, total = {}, 0
    for i, row in enumerate(fa):
        for j, v in enumerate(row):
            if not (isinstance(v, str) and v.startswith('=')):
                continue
            total += 1
            if total > cap:
                return total, pats
            rc = fr[i][j] if i < len(fr) and j < len(fr[i]) else v
            f = str(rc)
            addr = f"{_col_letter(c0 + j)}{r0 + i}"
            if f in pats:
                pats[f][0] += 1
                pats[f][3] = addr
            else:
                pats[f] = [1, addr, str(v), addr]
    return total, pats


def _hidden_count(whole, n, one, max_count=_HIDDEN_COUNT_MAX):
    """非表示の本数。まとめて 1 回聞き（False＝無し・True＝全部・None＝混在）、混在のときだけ 1 本ずつ数える。

    それまでは行数＋列数ぶん Hidden を聞いていた（2026-09-16）。混在で n が max_count を超えるときは
    数えずに None（あるが本数は不明）。
    """
    try:
        h = whole.Hidden
    except Exception:
        h = None
    if h is False:
        return 0
    if h is True:
        return n
    if n > max_count:
        return None
    cnt = 0
    for i in range(1, n + 1):
        try:
            if one(i):
                cnt += 1
        except Exception:
            pass
    return cnt


def _fml_text(f):
    """数式（'=' 始まり）なら文字列リテラルを伏せた中身を返す。数式でなければ ''。"""
    if not isinstance(f, str) or not f.startswith('='):
        return ''
    body = _FML_STR_RE.sub('""', f)[1:]
    return _FML_SHEETREF_RE.sub(lambda m: '#' * len(m.group(0)), body)


def _fml_rc(a):
    """'B12' → (12, 2)。'$B$12' も可。読めなければ None。"""
    m = re.match(r'^\$?([A-Za-z]{1,3})\$?(\d{1,7})$', str(a or '').strip())
    if not m:
        return None
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return int(m.group(2)), col


def _clip_fml(f, n=40):
    f = str(f or '').replace('\n', ' ')
    return f if len(f) <= n else f[:n] + '…'


def fml_magic_numbers(f):
    """式の中に直接書かれた数値を返す（0・1 は除く）。

    掛け算・割り算・足し引き・比較に隣り合う数だけを数える＝ VLOOKUP の列番号（,3,）や
    ROUND の桁（,2)）は「直書き」と呼ばない。=B7*C7*1.05 の 1.05、=E5*0.21 の 0.21、
    =IF(D2>=60,…) の 60 が対象。
    """
    body = _FML_RANGE_RE.sub(lambda m: '@' * len(m.group(0)), _fml_text(f))
    if not body:
        return []
    out = []
    for m in _FML_NUM_RE.finditer(body):
        try:
            v = float(m.group(1))
        except ValueError:
            continue
        if v in (0.0, 1.0):
            continue
        prev = body[:m.start()].rstrip()[-1:]
        nxt = body[m.end():].lstrip()[:1]
        if prev in _FML_OPS or nxt in _FML_OPS:
            out.append(m.group(1))
    return out


def fml_agg_ranges(f):
    """式の中の集計（SUM 等）と、その引数の範囲を返す → [(関数名, [(上,左,下,右), …]), …]。

    入れ子（=SUM(IF(…))）と別シート参照は見ない（見誤るくらいなら黙る）。
    """
    body = _fml_text(f)
    if not body:
        return []
    out = []
    for m in _FML_AGG_RE.finditer(body):
        depth, i = 1, m.end()
        while i < len(body) and depth:
            if body[i] == '(':
                depth += 1
            elif body[i] == ')':
                depth -= 1
            i += 1
        arg = body[m.end():i - 1] if depth == 0 else body[m.end():]
        if '(' in arg:
            continue
        boxes = []
        for rm in _FML_RANGE_RE.finditer(arg):
            a = _fml_rc(rm.group(1))
            b = _fml_rc(rm.group(2) or rm.group(1))
            if not a or not b:
                continue
            boxes.append((min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])))
        if boxes:
            out.append((m.group(1).upper(), boxes))
    return out


def _sums_the_others(f, col, rows):
    """その式が、同じ列の rows のセルを足し上げているか（合計列の下の総計＝形が違って当たり前）。"""
    for _func, boxes in fml_agg_ranges(f):
        for b in boxes:
            if b[1] == b[3] == col and any(b[0] <= r <= b[2] for r in rows):
                return True
    return False


def _sums_the_row(f, row, cols):
    """その式が、同じ行の cols のセルを足し上げているか（行の右端の合計＝形が違って当たり前）。"""
    for _func, boxes in fml_agg_ranges(f):
        for b in boxes:
            if b[0] == b[2] == row and any(b[1] <= c <= b[3] for c in cols):
                return True
    return False


def formula_copy_blockers(cells, limit=8):
    """横に並ぶ同じ関数の集計が、列ごとに違う式になっていないか（純 Python）。

    見つかったら **done を止める**（2026-09-09・shu 指示で気づきから関所に上げた）。
    値は合っていても、月や区分を式に埋めた表は列を 1 つ足した瞬間に壊れるため、直させる。
    関数名がそろっているときだけ見る（件数・合計・平均が横に並ぶ表は式が違って当たり前）。
    """
    # ④ 横に並ぶ同じ関数の集計が、列ごとに違う式（＝月や区分を式に埋めた＝横にコピーできない）
    #    2026-09-09 の実射で出た形: 1月 =SUMPRODUCT((MONTH(…)=1)*…)、2月 は =2、3月 は =3。
    #    値は合うが、月を 1 つ足した瞬間に壊れる。関数名がそろっているときだけ見る
    #    （件数・合計・平均が横に並ぶ表は式が違って当たり前なので咎めない）。
    row_f, hits, out = {}, {}, []
    for (r, c), fs in (cells or {}).items():
        a1 = fs[0] if isinstance(fs, (tuple, list)) else fs
        rc = fs[1] if isinstance(fs, (tuple, list)) and len(fs) > 1 else None
        if not (isinstance(a1, str) and a1.startswith('=') and isinstance(rc, str) and rc.startswith('=')):
            continue
        m = _FML_AGG_ANY_RE.search(a1)
        if not m:
            continue
        row_f.setdefault(r, []).append((c, m.group(1).upper(), a1, rc))
    for r in sorted(row_f):
        by_func = {}
        for c, func, a1, rc in row_f[r]:
            by_func.setdefault(func, []).append((c, a1, rc))
        for func, got in sorted(by_func.items()):
            cols = [c for c, _a, _rc in got]
            got = [(c, a1, rc) for c, a1, rc in got if not _sums_the_row(a1, r, [x for x in cols if x != c])]
            if len(got) < 3:
                continue
            shapes = {rc for _c, _a, rc in got}
            if len(shapes) < len(got):
                continue                      # 1 つでも同じ形があれば「列ごとに違う」ではない
            # 自分の列を指す番地を伏せると同じ式（=SUM($E$2:$E$8)・=SUM($F$2:$F$8)）＝絶対番地で書いただけで、月や区分は埋めていない
            # （2026-09-17 夜: 前年度比較の合計行を「横にコピーできない」と止め、値の正しい登録マクロの後を AI に回した）
            own = {re.sub(r'(?<![A-Z$])\$?' + _col_letter(c) + r'(?=\$?\d)', '#', a1) for c, a1, _rc in got}
            if len(own) == 1:
                continue
            # テーブルの集計行（=SUBTOTAL(109,[基本給])）は列の名前で参照する＝列ごとに違って当たり前
            if all(re.search(r'\[[^\]]+\]', a1) for _c, a1, _rc in got):
                continue
            hits.setdefault((func, tuple(sorted(c for c, _a, _rc in got))), []).append(r)
    for (func, cols), rows_hit in sorted(hits.items(), key=lambda kv: (kv[1][0], kv[0][0])):
        rr = (f"{rows_hit[0]}〜{rows_hit[-1]} 行目" if len(rows_hit) > 2 and rows_hit[-1] - rows_hit[0] == len(rows_hit) - 1
              else " ".join(str(x) for x in rows_hit) + " 行目")
        box = f"{_col_letter(cols[0])}{rows_hit[0]}:{_col_letter(cols[-1])}{rows_hit[-1]}"
        out.append(f"{rr}の {func} が列ごとに違う式です（{box}）＝横にコピーできません"
                   "（月や区分を式に埋めています）→ 同じ式を右にコピーできる形に直す。"
                   "月なら見出しのセルに**日付**を書く（2026/1/1。文字の「1月」にしない。"
                   "見え方は format の number_format で yyyy年m月 と出す）。"
                   '式はその見出しを参照して ">="&G$1 と "<"&EDATE(G$1,1) で判定する'
                   "（見出しの文字から月を切り出す式にしない）")
    return out


def formula_notes(cells, body_rows=None, limit=8):
    """数式の目（純 Python）。cells: {(行, 列): (A1 の式, R1C1 の式)} → 気づきの行のリスト。

    出すもの:
      ① 列の中で 1 つだけ形の違う式（R1C1 で比べる。合計行のように集計関数の有無が違うものは咎めない）
      ② 式に直接書かれた数値（率・単価・しきい値。前提はラベル付きのセルに置く）
      ③ 同じ行の集計で、起点行が式ごとに違う（=SUM(B2:B9) と =SUM(C3:C9) が並ぶ）
    body_rows を渡すと、その行だけを本文として見る。
    """
    items_by_col, out = {}, []
    for (r, c), fs in (cells or {}).items():
        a1 = fs[0] if isinstance(fs, (tuple, list)) else fs
        rc = fs[1] if isinstance(fs, (tuple, list)) and len(fs) > 1 else None
        if not (isinstance(a1, str) and a1.startswith('=')):
            continue
        if body_rows is not None and r not in body_rows:
            continue
        key = rc if isinstance(rc, str) and rc.startswith('=') else a1
        items_by_col.setdefault(c, []).append((r, a1, key))
    # ① 列の中で形の違う式。集計の式（合計行）と本文の式は別々に見る＝
    #    列の末尾の =SUM が「形が違う」に数えられて、本物の 1 行を薄めてしまうのを防ぐ
    odd = []
    for c in sorted(items_by_col):
        for agg in (False, True):
            items = sorted(x for x in items_by_col[c] if bool(_FML_AGG_ANY_RE.search(x[1])) is agg)
            if len(items) < 4:
                continue
            cnt = {}
            for _r, _a, k in items:
                cnt[k] = cnt.get(k, 0) + 1
            top = max(cnt, key=lambda k: cnt[k])
            if cnt[top] < 3 or (len(items) - cnt[top]) > max(1, len(items) // 5):
                continue
            top_a1 = next(a for _r, a, k in items if k == top)
            top_rows = [r for r, _a, k in items if k == top]
            diff = [(r, a1) for r, a1, k in items if k != top and not _sums_the_others(a1, c, top_rows)]
            for r, a1 in diff:
                # 先頭行だけが違うのは累計（=D2 → =E2+D3 …）の書き始めのことがある＝そう添えて出す
                tail = "・先頭行だけ＝累計の書き始めならこのままでよい" if (len(diff) == 1 and r == items[0][0]) else ""
                odd.append(f"{_col_letter(c)}{r}（{_clip_fml(a1)}／他 {cnt[top]} 行は {_clip_fml(top_a1)}{tail}）")
    if odd:
        out.append("列の中で式の形が違うセル: " + "／".join(odd[:limit]) + ("…" if len(odd) > limit else "")
                   + "  → 隣の行と同じ式に直すか、違う理由を報告に書く")
    # ② 式に直書きされた数値
    mags = {}
    for (r, c), fs in sorted((cells or {}).items()):
        a1 = fs[0] if isinstance(fs, (tuple, list)) else fs
        if body_rows is not None and r not in body_rows:
            continue
        for s in fml_magic_numbers(a1):
            mags.setdefault(s, []).append(f"{_col_letter(c)}{r}")
    if mags:
        parts = [f"{s}（{' '.join(v[:4])}{'…' if len(v) > 4 else ''}・{len(v)}セル）"
                 for s, v in sorted(mags.items(), key=lambda kv: (-len(kv[1]), kv[0]))[:4]]
        out.append("式に直接書かれた数値: " + "／".join(parts)
                   + "  → 率・単価・しきい値はラベルを付けたセルに置いて参照する")
    # ③ 同じ行の集計で起点行が違う
    by_row = {}
    for (r, c), fs in (cells or {}).items():
        a1 = fs[0] if isinstance(fs, (tuple, list)) else fs
        aggs = fml_agg_ranges(a1)
        if len(aggs) != 1:
            continue
        func, boxes = aggs[0]
        vert = [b for b in boxes if b[1] == b[3] == c and b[2] < r]
        if len(vert) != 1:
            continue
        by_row.setdefault(r, []).append((c, func, vert[0]))
    for r in sorted(by_row):
        got = by_row[r]
        if len(got) < 3:
            continue
        starts = {}
        for c, func, b in got:
            starts.setdefault((func, b[0]), []).append(c)
        if len(starts) < 2:
            continue
        top = max(starts, key=lambda k: len(starts[k]))
        odd2 = [f"{_col_letter(c)}{r}" for k, cs in starts.items() if k != top for c in sorted(cs)]
        out.append(f"{r} 行目の集計の起点行がそろっていません: 多数は {top[0]} が {top[1]} 行目から／"
                   + " ".join(odd2[:limit]) + " だけ違う  → どちらが正しいかを確かめて直す")
    return out


def cmd_materials(args):
    """先回り材料：1シートについて、手を動かす前に見るべきものを1回でまとめて出す。

    Excelコンボの「先回り材料」から輸入（2026-09-02）。手の型の初手（「見る」は これ1回）。
    使用範囲・非表示・値（小さい表は全体＋長文セルの全文）・結合・テーブル・名前定義・数式の型・
    エラーセル・図形ボタン・### のセル・列幅。推測で動く前に、ここで現物を見る。
    あわせて仕事の時計を押す（tidy／write-range／write-cells が「materials から N 秒」を出す）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    try:
        head_n = int(getattr(args, 'rows', None) or 5)
    except (TypeError, ValueError):
        print("エラー: --rows は数値で指定してください")
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    name = rest[0] if rest else sheet_opt
    ws = None
    if name:
        for sh in wb.Sheets:
            if sh.Name == name:
                ws = sh
                break
        if ws is None:
            print(f"エラー: シート '{name}' が見つかりません")
            return False
    else:
        ws = wb.ActiveSheet
    active = wb.ActiveSheet.Name
    tail = "" if ws.Name == active else f"   （アクティブは {active}）"
    job_clock_start(f"{wb.Name}!{ws.Name}")          # 仕事の時計を押す（tidy／write が経過秒を出す）
    print(f"ブック: {wb.Name}   シート: {ws.Name}{tail}")
    print("=" * 60)
    try:
        ur = ws.UsedRange
        nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
        ur_addr = ur.Address.replace('$', '')
    except Exception:
        print("使用範囲: (空)")
        print("=" * 60)
        return True
    print(f"使用範囲: {ur_addr}  {nr}行 x {nc}列")
    cmd_progress_note(f"{ws.Name}: 値・結合・テーブルを読んでいます（{nr}行 x {nc}列）")
    # 非表示の行・列（まとめて 1 回聞き、混在のときだけ数える・2026-09-16）
    try:
        hid_r = _hidden_count(ur.EntireRow, nr, lambda i: ur.Rows(i).Hidden)
        hid_c = _hidden_count(ur.EntireColumn, nc, lambda j: ur.Columns(j).Hidden)
        if hid_r or hid_c or hid_r is None or hid_c is None:
            _fmt = lambda v: "あり（本数は数えない・大）" if v is None else f"{v} 本"   # noqa: E731
            print(f"非表示: 行 {_fmt(hid_r)}・列 {_fmt(hid_c)}")
    except Exception:
        pass
    # 先頭N行（値）。小さい表（40行×30列まで・--rows 指定なし）は全体を出す＝「見る」を1往復で終える。
    # 格子で '…' に切れた長文セル（A3 の指示文など）は下に全文を出す（read-range を重ねない）
    try:
        # --full（エージェント・採点係）は 200 行まで全体を出す（2026-09-11: 56 行の表で先頭 5 行しか見せず、
        # AI が毎回 1 往復目を「表を読み直す」だけに使っていた＝7〜12 秒と 1 往復の無駄）
        whole = (getattr(args, 'rows', None) is None
                 and nc <= 30 and (nr <= 40 or (getattr(args, 'full', False) and nr <= 200)))
        n = nr if whole else min(head_n, nr)
        head = ws.Range(ur.Cells(1, 1), ur.Cells(n, nc))
        label = "全体" if whole else f"先頭{n}行"
        print(f"--- {label}（値） ---")
        print(_values_to_grid(head))
        _print_long_cells(head)
    except Exception as e:
        print(f"（先頭行を読めませんでした: {e}）")
    # 結合セル
    areas, skipped = _merged_areas_in_range(ur)
    if skipped:
        print(f"結合セル: 未走査（{skipped}セル・大）")
    elif areas:
        shown = areas[:12]
        more = f"  …他{len(areas) - 12}件" if len(areas) > 12 else ""
        print(f"結合セル {len(areas)}件: " + ", ".join(shown) + more)
    # 書式の目（見出しの結合の親子・合計行・塗りの有無。番地の細部は style-map・2026-09-17）
    try:
        for _ln in style_summary(ws, ur, nr, nc, areas):
            print(_ln)
    except Exception:
        pass
    # テーブル
    try:
        los = list(ws.ListObjects)
        if los:
            print("テーブル: " + ", ".join(
                f"{lo.Name}({lo.Range.Address.replace('$', '')})" for lo in los))
    except Exception:
        pass
    # 名前定義（このシートを指すもの）
    try:
        hits = []
        for nm in wb.Names:
            try:
                ref = str(nm.RefersTo)
            except Exception:
                continue
            if ws.Name in ref:
                hits.append(f"{nm.Name} = {ref}")
        if hits:
            print("名前定義（このシートを指す。消すと式が壊れる）:")
            for h in hits[:20]:
                print("  " + h)
            if len(hits) > 20:
                print(f"  …他{len(hits) - 20}件")
    except Exception:
        pass
    # 数式の型（同じ形は R1C1 でまとめ、見せる式は A1 形式にする）
    #
    # 2026-09-08: R1C1 のまま見せていたら、採点係が =AVERAGE(R[-22]C:R[-3]C) を「G5:G25＝見出し行を含む」と
    # 暗算し間違え、正しい式に「直せ」と差し戻して往復を 1 回捨てた（実際は G6:G25）。作業側も同じ理由で
    # D31・G27:G28 を read で読み直しており、往復 1 回ぶんが数式の読み直しに消えていた。
    # 番地と A1 形式の式をここに出せば、どちらも暗算も読み直しも要らない。
    pats, total = {}, 0                             # R1C1 → [個数, 代表の番地, 代表の A1 式, 最後の番地]
    _fa = _fr = None                                # 一括で読んだ式の格子（下の「気づき（数式）」でも使う）
    try:
        if nr * nc <= _FML_BULK_MAX_CELLS:
            # 式は一括で 2 回読む（A1 形式と R1C1 形式）。以前は数式セル 1 つにつき COM を 3 回呼んでいた＝
            # 5,000 セルで 15,000 往復・数式の多い表で数秒〜数十秒（2026-09-16）
            cmd_progress_note(f"{ws.Name}: 数式の型を読んでいます")
            _fa = _grid_of_value(ur.Formula)
            _fr = _grid_of_value(ur.FormulaR1C1)
            total, pats = _formula_patterns(_fa, _fr, int(ur.Row), int(ur.Column))
        else:
            for c in ur.SpecialCells(-4123):        # xlCellTypeFormulas（巨大シートだけ従来の 1 セルずつ）
                total += 1
                if total > 5000:
                    break
                try:
                    f = str(c.FormulaR1C1)
                except Exception:
                    continue
                addr = str(c.Address).replace('$', '')
                if f in pats:
                    pats[f][0] += 1
                    pats[f][3] = addr
                else:
                    try:
                        a1 = str(c.Formula)
                    except Exception:
                        a1 = f
                    pats[f] = [1, addr, a1, addr]
        if not total:
            print("数式: なし")
        else:
            print(f"数式: {total if total <= 5000 else 5000}{'+' if total > 5000 else ''}セル・型 {len(pats)}種"
                  + "（式は A1 形式・番地は代表の 1 つ。R1C1 を数え直さない）")
            for _f, (k, first, a1, last) in sorted(pats.items(), key=lambda kv: -kv[1][0])[:10]:
                where = first if k == 1 else f"{first}〜{last}"
                print(f"  {k:>4} × {where:<12} {a1[:70]}")
    except Exception:
        print("数式: なし")                          # 該当なしは SpecialCells が例外＝正常
    # エラーセル
    errs = []
    for typ in (-4123, 2):                          # 数式 / 定数
        try:
            for c in ur.SpecialCells(typ, 16):      # 16 = xlErrors
                errs.append(f"{c.Address.replace('$', '')}={c.Text}")
                if len(errs) >= 200:
                    break
        except Exception:
            pass
    if errs:
        more = " …" if len(errs) > 10 else ""
        print(f"⚠ 【要修正】エラーセル: {len(errs)}{'+' if len(errs) >= 200 else ''}個  " + " ".join(errs[:10]) + more)
    # 図形・ボタン
    shapes = []
    try:
        _collect_shapes(ws.Shapes, shapes)
    except Exception:
        pass
    if shapes:
        print(f"図形・ボタン: {len(shapes)}個")
        for s in shapes[:20]:
            t = s.get("text") or ""
            oa = s.get("onaction")
            geo = " ".join(f"{k}={s[k]}" for k in ("l", "t", "w", "h") if k in s)   # 位置と大きさ（並べる手の材料）
            print(f"  {s.get('name', '?')}" + (f"「{t}」" if t else "") + (f" → {oa}" if oa else "")
                  + (f"  [{geo}]" if geo else ""))
        if len(shapes) > 20:
            print(f"  …他{len(shapes) - 20}件")
    # ### で読めないセル（画面の文字で数える）
    try:
        if nr * nc <= 3000:
            hashes, _bad, _note = _hash_scan(ur, max_rows=nr, max_cols=nc)
            if hashes:
                print(f"⚠ 【表示崩れ】'###' で読めないセル: {hashes}個（{' '.join(_bad)}・列幅不足）")
        else:
            print("'###' の検査: 使用範囲が大きいため省略（範囲を絞って tidy か --show で見る）")
    except Exception:
        pass
    # 気づき（数式）。式の形の乱れ・直書きの数値・集計の起点を道具が数える（2026-09-09）。
    # それまでは式の中身を見るのが AI の目視だけで、大きい表では読み残しが出ていた
    try:
        if total and nr * nc <= _FML_LENS_MAX_CELLS:
            if _fa is None or _fr is None:          # 上で一括に読めていなければここで読む（普通は読めている）
                _fa = _grid_of_value(ur.Formula)
                _fr = _grid_of_value(ur.FormulaR1C1)
            _r0, _c0 = int(ur.Row), int(ur.Column)
            _fcells = {}
            for _i, _row in enumerate(_fa):
                for _j, _v in enumerate(_row):
                    if isinstance(_v, str) and _v.startswith('='):
                        _rc = _fr[_i][_j] if _i < len(_fr) and _j < len(_fr[_i]) else None
                        _fcells[(_r0 + _i, _c0 + _j)] = (_v, _rc)
            _notes = formula_copy_blockers(_fcells) + formula_notes(_fcells)
            if _notes:
                print("気づき（数式。値でなく式の形を道具が見ています）:")
                for _n in _notes:
                    print("  " + _n)
            else:
                print("気づき（数式）: なし（式の形の乱れ・直書きの数値・集計の起点、いずれも無し）")
    except Exception as e:
        print(f"（数式の気づきを数えられませんでした: {e}）")
    # 気づき（表の汚れ）。格子では見えない空白の幅・日付の型・半角カナ・重複を道具が数える（2026-09-07）
    try:
        if nr * nc <= 20000:
            grid = _grid_of_value(ur.Value)
            r0, c0 = int(ur.Row), int(ur.Column)
            # 診断の目（循環参照＝要修正・外れ値＝気づき。全部見るなら diagnose・2026-09-17）
            try:
                if _fa is not None:
                    _blk, _ntc = diagnose_notes(_fa, _fr, grid, r0, c0)
                    for _b in _blk:
                        print(f"⚠ 【要修正】循環参照 {_b}")
                    for _n in _ntc:
                        print(f"気づき（外れ値）: {_n}")
            except Exception:
                pass
            hidx = _guess_header_idx(grid)
            try:
                _look = _body_look_mixed(ws, grid, r0, c0, hidx)
            except Exception:
                _look = None
            for ln in dirt_notes(grid, r0, c0, hidx, _date_col_formats(ws, grid, r0, c0, hidx),
                                 look_mixed=_look):
                print(ln)
    except Exception as e:
        print(f"（気づきを数えられませんでした: {e}）")
    # 列幅
    try:
        c0 = int(ur.Column)
        widths = [f"{_col_letter(c0 + j - 1)}={float(ur.Columns(j).ColumnWidth):g}"
                  for j in range(1, min(nc, 30) + 1)]
        print("列幅: " + " ".join(widths) + ("  …" if nc > 30 else ""))
    except Exception:
        pass
    print("=" * 60)
    return True


_SEIRI_MODULE = '表の整理'
_SEIRI_TIDY = '表を整える'
_SEIRI_DEDUPE = '重複行を消す'
_CELL_REF_RE = re.compile(r"(?<![A-Za-z_!])\$?([A-Z]{1,3})\$?(\d{1,7})(?![\d(])")


def _macro_book(xl, module, sub):
    """モジュール module に Sub sub を持つブック（アドイン含む）の名前。無ければ None（2026-09-13）。"""
    from vbam_vba import _project_book_name
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return None
    for p in projects:
        try:
            cm = p.VBComponents(module).CodeModule
            n = int(cm.CountOfLines)
            if n and re.search(r'^\s*Sub\s+' + re.escape(sub) + r'\b', cm.Lines(1, n), re.M):
                name = _project_book_name(xl, p)
                if name:
                    return name
        except Exception:
            continue
    return None


def error_hint(text, formula, empty_refs):
    """エラーセルの一言（純 Python）。#DIV/0! で式が空のセルを参照していれば「分母の X が空」。"""
    t = str(text or '')
    if t == '#DIV/0!' and empty_refs:
        return f"式が空のセル {', '.join(empty_refs)} で割っている（分母が 0）"
    if t == '#REF!':
        return "消えた行・列を参照している"
    if t == '#NAME?':
        return "関数名か名前定義の綴りが違う"
    if t == '#VALUE!':
        return "文字のセルを計算に使っている"
    if t == '#N/A':
        return "検索で見つからない（IFERROR で「〜なし」に）"
    return ''


def cmd_seiri(args):
    """表を直す 1 手目: seiri [--dedupe]（2026-09-13・shu「まとめてみろ」）

    materials で表の全体を読んでから、次の往復でマクロと書き込み、の 2 往復と読む時間を畳む。
    画面のシートに「表を整える」マクロ（開いているブックかアドインのモジュール「表の整理」）を撃ち、
    残り＝エラーセル（式と一言の原因）・数式と表の気づき・指示文らしい長文セル・### だけを出す。
    直す手（式の書き直しなど判断の要るもの）はこの残りの分だけ。仕事の時計も押す。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    job_clock_start(f"{wb.Name}!{ws.Name}")
    print(f"ブック: {wb.Name}   シート: {ws.Name}")
    owner = _macro_book(xl, _SEIRI_MODULE, _SEIRI_TIDY)
    names = [_SEIRI_TIDY] + ([_SEIRI_DEDUPE] if getattr(args, 'dedupe', False) else [])
    if owner:
        t0 = time.time()
        try:
            wb.Activate()
            ws.Activate()
            q = owner.replace("'", "''")
            for nm in names:
                xl.Run(f"'{q}'!{nm}")
            print(f"マクロ: {' → '.join(names)}（{owner}・{time.time() - t0:.2f} 秒）")
        except Exception as e:
            print(f"マクロを撃てませんでした: {e}（残りだけ出します）")
    else:
        print(f"マクロ: モジュール「{_SEIRI_MODULE}」の {_SEIRI_TIDY} が開いているブック・アドインに無いので撃っていません")
    try:
        ur = ws.UsedRange
        nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
        print(f"使用範囲: {ur.Address.replace('$', '')}  {nr}行 x {nc}列")
    except Exception:
        print("使用範囲: (空)")
        return True
    print("--- 残り（ここだけ直す） ---")
    # エラーセル（いちばん上に・式と一言の原因つき）
    errs = []
    for typ in (-4123, 2):                          # 数式 / 定数
        try:
            for c in ur.SpecialCells(typ, 16):      # 16 = xlErrors
                addr = c.Address.replace('$', '')
                f = ''
                empty = []
                try:
                    if c.HasFormula:
                        f = str(c.Formula)
                        for col, row in _CELL_REF_RE.findall(f):
                            try:
                                v = ws.Range(f"{col}{row}").Value
                            except Exception:
                                continue
                            if v is None or (isinstance(v, str) and v.strip() == ''):
                                empty.append(f"{col}{row}")
                except Exception:
                    pass
                hint = error_hint(c.Text, f, empty)
                errs.append(f"  {addr} = {c.Text}" + (f"   式 {f}" if f else "") + (f"   → {hint}" if hint else ""))
                if len(errs) >= 30:
                    break
        except Exception:
            pass
    print(f"エラーセル: {len(errs)}個" if errs else "エラーセル: なし")
    for e in errs:
        print(e)
    # 指示文らしい長文セル（シートに書いてある頼みごと）
    try:
        _print_long_cells(ws.Range(ur.Cells(1, 1), ur.Cells(min(nr, 40), min(nc, 30))))
    except Exception:
        pass
    # 気づき（数式・表の汚れ）
    try:
        if nr * nc <= _FML_LENS_MAX_CELLS:
            _fa = _grid_of_value(ur.Formula)
            _fr = _grid_of_value(ur.FormulaR1C1)
            r0, c0 = int(ur.Row), int(ur.Column)
            cells = {}
            for i, row in enumerate(_fa):
                for j, v in enumerate(row):
                    if isinstance(v, str) and v.startswith('='):
                        rc = _fr[i][j] if i < len(_fr) and j < len(_fr[i]) else None
                        cells[(r0 + i, c0 + j)] = (v, rc)
            notes = (formula_copy_blockers(cells) + formula_notes(cells)) if cells else []
            if notes:
                print("気づき（数式）:")
                for n in notes:
                    print("  " + n)
            grid = _grid_of_value(ur.Value)
            try:
                _blk, _ntc = diagnose_notes(_fa, _fr, grid, r0, c0)
                for _b in _blk:
                    print(f"⚠ 【要修正】循環参照 {_b}")
                for _n in _ntc:
                    print(f"気づき（外れ値）: {_n}")
            except Exception:
                pass
            hidx = _guess_header_idx(grid)
            try:
                look = _body_look_mixed(ws, grid, r0, c0, hidx)
            except Exception:
                look = None
            for ln in dirt_notes(grid, r0, c0, hidx, _date_col_formats(ws, grid, r0, c0, hidx), look_mixed=look):
                print(ln)
            try:
                _areas, _sk = _merged_areas_in_range(ur)
                for _ln in style_summary(ws, ur, nr, nc, _areas):
                    print(_ln)
            except Exception:
                pass
    except Exception as e:
        print(f"（気づきを数えられませんでした: {e}）")
    try:
        if nr * nc <= 3000:
            hashes, _bad, _n = _hash_scan(ur, max_rows=nr, max_cols=nc)
            print(f"'###' で読めないセル: {hashes}個（{' '.join(_bad)}）" if hashes else "'###': なし")
    except Exception:
        pass
    print("（保存はしていません）")
    return True


def _shape_text(shp):
    """図形の表示文字を複数方式で拾う。

    Formsコントロール(ボタン等)は TextFrame.Characters().Text、
    AutoShape等は TextFrame2.TextRange.Text に文字が入る（相手で口が違う）。
    先に非空を返した方を採用。どちらも取れなければ None。
    """
    for getter in (
        lambda: shp.TextFrame.Characters().Text,
        lambda: shp.TextFrame2.TextRange.Text,
    ):
        try:
            t = getter()
            if t:
                return t
        except Exception:
            continue
    return None


def _collect_shapes(shapes, out, in_group=None):
    """Shapes を平坦なリストに集める。グループ(msoGroup=6)は中のボタンも展開して拾う。

    ファイル一覧等ではボタンがグループにまとめられており、展開しないと
    中の「抽出開始」「選択抽出」等が丸ごと落ちる（実測で判明）。
    """
    for shp in shapes:
        s = {}
        try:
            s["name"] = shp.Name
        except Exception:
            pass
        try:
            typ = int(shp.Type)
            s["type"] = typ
        except Exception:
            typ = None
        txt = _shape_text(shp)
        if txt:
            s["text"] = txt
        try:
            s["l"] = round(float(shp.Left))
            s["t"] = round(float(shp.Top))
            # サイズも採る。位置だけ見ていると、ボタンを 1×1pt に潰す／2倍に伸ばす
            # といった変更を snapshot-diff が「差分なし」と見逃す
            s["w"] = round(float(shp.Width))
            s["h"] = round(float(shp.Height))
        except Exception:
            pass
        try:
            oa = shp.OnAction
            if oa:
                s["onaction"] = oa.split('!')[-1].strip("'\" ")
                if '!' in oa:
                    # 「'ブック名'!マクロ名」のブック修飾。剥がして捨てると
                    # アドイン先の配線を wiring が「存在しない」と誤判定する
                    book = oa.rsplit('!', 1)[0].strip("'\" ")
                    if book:
                        s["onaction_book"] = book
        except Exception:
            pass
        if in_group:
            s["group"] = in_group
        out.append(s)
        if typ == 6:                       # msoGroup → 中身を1段展開
            try:
                _collect_shapes(shp.GroupItems, out, in_group=s.get("name"))
            except Exception:
                pass


# 書式の目が見る項目。COM のプロパティ1つ＝1往復（約1.1ms）なので、
# 数を絞りつつ「クリーン化で消えるもの」を確実に押さえる。
# 罫線は Borders.LineStyle が「全辺・全セルで同じなら値／不揃いなら None」を返す
# ため1往復で足りる（罫線が丸ごと消えたかは、これで確実に分かる）。
_FMT_PROPS = (
    ('bold',   lambda r: r.Font.Bold),
    ('size',   lambda r: r.Font.Size),
    ('font',   lambda r: r.Font.Name),
    ('italic', lambda r: r.Font.Italic),
    ('color',  lambda r: r.Font.Color),
    ('fill',   lambda r: r.Interior.Color),
    ('numfmt', lambda r: r.NumberFormat),
    ('border', lambda r: r.Borders.LineStyle),
    ('halign', lambda r: r.HorizontalAlignment),
)


def _fmt_read(rng, keys):
    """範囲の書式をまとめ読み。範囲内で不揃いな項目は None（Excel の性質）。

    値は JSON に載るように素の型へ均す（COM の float/bool/str がそのまま来る）。
    読めなかった項目は "?" にして「揃っている」と誤解させない。
    """
    out = {}
    for key, getter in _FMT_PROPS:
        if keys is not None and key not in keys:
            continue
        try:
            v = getter(rng)
        except Exception:
            out[key] = "?"          # 読めなかった（＝揃っているとは言えない）
            continue
        if v is None:
            out[key] = None         # 範囲内で不揃い
        elif isinstance(v, float) and v == int(v):
            out[key] = int(v)
        elif isinstance(v, (int, str, bool)):
            out[key] = v
        else:
            out[key] = str(v)
    return out


def _collect_format(sh, ur, nr, nc):
    """シートの書式を採る（値・結合・図形に続く「四つ目の目」）。

    Excel の COM には書式の一括読みが無い。セル単位で読むと 1万セルで33秒（実測）。
    一方 Excel は「範囲の書式が揃っていれば値／不揃いなら None」を返す性質を持つ。
    これを使って、揃っている範囲は1往復で確定させる:

      1. シート全体を1回読む。揃っている項目はここで確定（列・行を見る必要なし）
      2. 不揃い（None）だった項目だけ、列ごとに読む
      3. 行は二分割で降りる。揃っていればその塊で確定、不揃いなら半分に割る
         （見出し1行だけ太字のシートなら、1万行でも約30往復で終わる）

    行を1行ずつ読む素朴なやり方だと、見出し行が太字なだけで本文1万行を1万回
    読みに行き 38.9秒かかった。行上限で打ち切れば速いが、上限より下の書式を
    一切見ない盲点ができる（数万行のシートではそこが本体）。二分割ならどちらも要らない。

    戻り値:
      sheet / cols … 揃っていればその値、不揃いなら None（欠測ではない）
      row_blocks   … 同じ書式が続く行の塊のリスト [{r0, r1, 各項目}, ...]
      読めなかった項目は "?"。
    """
    info = {}
    sheet_fmt = _fmt_read(ur, None)
    info["sheet"] = sheet_fmt

    # 不揃いだった項目だけ下へ降りる（揃っている項目は列・行を見る必要がない）
    mixed = [k for k, v in sheet_fmt.items() if v is None]

    # 揃っている項目は、各行・各列も必ずその値になる（範囲の部分集合だから）。
    # ここで埋めておかないと、行・列の dict に「その項目のキーが無い」状態になり、
    # 前後で揃い方が変われば snapshot-diff が「不在」を「None（不揃い）への変化」
    # として誤報する（例: クリーン化で太字が揃った途端に「太字 False → None」）。
    # 「測っていない」を「不揃い」と言うのは、検分器がついてはいけない類の嘘。
    # COM 往復は増えない（シート全体の1回で確定している値を写すだけ）。
    inherited = {k: v for k, v in sheet_fmt.items() if k not in mixed}

    # 列幅・行高も同じ理屈で先にまとめて見る。ほとんどのシートは全列同じ幅・
    # 全行同じ高さなので、その場合は1回の読みで確定し、列ぶん・行ぶんの COM 往復
    # （1本あたり約1.1ms）を丸ごと省ける。不揃い（None）のときだけ下へ降りる。
    def _uniform(getter):
        try:
            v = getter()
            return round(float(v), 2) if v is not None else None
        except Exception:
            return None

    uni_w = _uniform(lambda: ur.EntireColumn.ColumnWidth)
    uni_h = _uniform(lambda: ur.EntireRow.RowHeight)
    # 非表示行の有無も先にまとめて見る。Hidden が False で揃っていれば
    # 「非表示行は1行も無い」と確定するので、行ごとの問い合わせを省ける
    try:
        rows_all_visible = (ur.EntireRow.Hidden is False)
    except Exception:
        rows_all_visible = False

    cols = {}
    for ci in range(1, nc + 1):
        cno = ur.Column + ci - 1
        col_rng = sh.Range(sh.Cells(ur.Row, cno), sh.Cells(ur.Row + nr - 1, cno))
        d = dict(inherited)
        if mixed:
            d.update(_fmt_read(col_rng, mixed))
        if uni_w is not None:
            d["w"] = uni_w                     # 全列同じ幅と確定済み（追加の往復なし）
        else:
            try:
                d["w"] = round(float(sh.Columns(cno).ColumnWidth), 2)
            except Exception:
                d["w"] = "?"
        try:
            d["hidden"] = bool(sh.Columns(cno).Hidden)
        except Exception:
            pass
        cols[_col_letter(cno)] = d
    info["cols"] = cols

    # --- 行は「1行ずつ」ではなく「同じ書式が続く塊」で採る ---
    #
    # 1行ずつ読むと、見出し行が1行太字なだけでシート全体が「不揃い」になり、
    # 本文が全部同じ書式でも1万行を1万回読みに行く（実測: 10,004行で38.9秒）。
    # 行上限で打ち切れば速くはなるが、上限より下の書式を一切見ない盲点ができる
    # （ファイル一覧.xlsm のような数万行のシートでは、そこが本体）。
    # 速度と盲点のどちらかを選ばされる時点で、やり方が間違っている。
    #
    # そこで二分割で降りる: 範囲を読んで揃っていればその塊で確定（1往復）、
    # 不揃いなら半分に割って繰り返す。見出し1行だけ違うシートなら、1万行でも
    # 約30往復で終わる。上限も要らなくなり、盲点も消える。
    #
    # ただし「列ごとの違い」で不揃いになっている項目（例: 金額列だけ表示形式が
    # #,##0）は、どの行を読んでも必ず不揃い（None）になる。それを二分割で追うと
    # 最後の1行まで割り続け、途中の読みが全部無駄になる（実測で倍近く遅くなった）。
    # そういう項目は「各列では揃っている」＝列の目で既に捉えているので、行へは
    # 降りない。降りるのは「どこかの列の中で不揃い」＝行方向の変化がある項目だけ。
    row_keys = [k for k in mixed
                if any((cols[cl].get(k) is None) for cl in cols)]
    # 降りない項目は、各行の中では必ず不揃い（列同士で値が違うから）＝ None が事実。
    # ここを空欄にすると「測っていない」と「不揃い」が混ざり、diff が誤報する
    col_only = {k: None for k in mixed if k not in row_keys}

    # 二分割には「読む回数の上限」を付ける。
    # 均一なシート（本文1万行が同じ書式＝ファイル一覧型）は数十往復で全行を覆えるが、
    # 「金額列だけカンマ書式」のように "どの行の中でも不揃い" な項目があると、
    # 二分割は最後の1行まで割り続け、途中の読みが全部無駄になる（実測で倍近く遅い）。
    # 上限を置けば、均一なシートは今までどおり全行カバー、割り続けても終わらない
    # シートは途中で止まる。止めた場合は「そこまでの粒度でしか見ていない」と必ず残す
    # （黙って粗く見て「差分なし」と言うのが、検分器として最悪の嘘）。
    budget = {"left": 600}          # 行方向の読み取り回数の上限（約1〜2秒ぶん）
    coarse = []                     # 上限に当たって細かく見られなかった範囲

    row_blocks = []

    def _row_range(r0, r1):
        return sh.Range(sh.Cells(r0, ur.Column), sh.Cells(r1, ur.Column + nc - 1))

    def _descend(r0, r1, keys):
        """[r0, r1] の書式を採る。揃っていれば1塊、不揃いなら二分割して降りる。"""
        d = dict(inherited)
        d.update(col_only)          # 列方向だけの違い＝どの行でも不揃い（None が事実）
        if keys:
            budget["left"] -= 1
            d.update(_fmt_read(_row_range(r0, r1), keys))
        # 高さ・非表示も同じ塊の単位で見る（全体で揃っていれば追加の往復ゼロ）
        if uni_h is not None:
            d["h"] = uni_h
        else:
            try:
                v = sh.Range(sh.Rows(r0), sh.Rows(r1)).RowHeight
                d["h"] = round(float(v), 2) if v is not None else None
            except Exception:
                d["h"] = "?"
        if rows_all_visible:
            hid = False
        else:
            try:
                hv = sh.Range(sh.Rows(r0), sh.Rows(r1)).Hidden
                hid = None if hv is None else bool(hv)
            except Exception:
                hid = None
        if hid:
            d["hidden"] = True

        # この塊の中で不揃いな項目（None）があるか。無ければ塊として確定。
        # col_only の項目（列方向だけの違い）は、行を割っても永遠に None のままなので
        # 「まだ不揃い」の判定から外す（外さないと最後の1行まで割り続ける）
        still_mixed = [k for k, v in d.items()
                       if v is None and k not in col_only]
        if not still_mixed or r0 >= r1:
            # 1行まで降りてなお None のものは「その行の中で不揃い」＝事実として残す
            row_blocks.append({"r0": r0, "r1": r1, **d})
            return
        if budget["left"] <= 0:
            # 上限に到達。ここから下は細かく見ない。粗い塊のまま残し、
            # 「この範囲はこの粒度でしか見ていない」を必ず報告する
            coarse.append([r0, r1])
            row_blocks.append({"r0": r0, "r1": r1, **d})
            return
        mid = (r0 + r1) // 2
        _descend(r0, mid, still_mixed)
        _descend(mid + 1, r1, still_mixed)

    r_top, r_bot = ur.Row, ur.Row + nr - 1
    _descend(r_top, r_bot, mixed)
    if coarse:
        info["rows_coarse"] = coarse

    # 隣り合う塊で中身が同じなら畳む（二分割の切れ目が残るのを均す）
    merged_blocks = []
    for b in sorted(row_blocks, key=lambda x: x["r0"]):
        if merged_blocks:
            prev = merged_blocks[-1]
            same = all(prev.get(k) == b.get(k)
                       for k in set(prev) | set(b) if k not in ("r0", "r1"))
            if same and prev["r1"] + 1 == b["r0"]:
                prev["r1"] = b["r1"]
                continue
        merged_blocks.append(dict(b))
    info["row_blocks"] = merged_blocks
    return info


def _snapshot_doc(wb, only_sheet=None, max_rows=5000, no_format=False):
    """開いているブック1冊を意味構造 dict に畳む（cmd_snapshot / rehearse の共用部品）。

    wb は既に掴んでいる Workbook をそのまま受け取る（get_workbook を通さない）。
    rehearse のように「別インスタンスで開いた非表示のコピー」を対象にするとき、
    パス指定で get_workbook を通すと ROT 走査のタイミング次第で同じファイルを
    二重に開きにいくため、ハンドル直渡しにしている。
    only_sheet が見つからないときは ValueError を送出する。
    """
    active = wb.ActiveSheet.Name

    if only_sheet:
        targets = [sh for sh in wb.Sheets if sh.Name == only_sheet]
        if not targets:
            raise ValueError(f"シート '{only_sheet}' が見つかりません")
    else:
        targets = list(wb.Sheets)

    sheets = {}
    for sh in targets:
        info = {}
        try:
            ur = sh.UsedRange
            nr, nc = ur.Rows.Count, ur.Columns.Count
            info["dims"] = f"{nr}行 x {nc}列"
            info["used"] = ur.Address
        except Exception:
            ur, nr, nc = None, 0, 0
            info["dims"] = "(空)"
        info["visible"] = (sh.Visible == -1)

        # --- セル（疎：空セル・空行は落とす。読み込みも max_rows 行までに絞る）---
        cells = []
        if ur is not None and nr > 0:
            r0, c0 = ur.Row, ur.Column
            read_rows = min(nr, max_rows)
            try:
                sub = sh.Range(ur.Cells(1, 1), ur.Cells(read_rows, nc))
                raw = sub.Value
            except Exception as e:
                # 黙って None にすると「空のシート」と見分けがつかず、snapshot-diff が
                # 数千件のセル追加/削除を誤報する（図形の shapes_error と同じ理由で
                # 「読めなかった事実」を残す）
                info["cells_error"] = str(e)
                raw = None
            if raw is None:
                grid = []
            elif not isinstance(raw, tuple):
                grid = [(raw,)]
            else:
                grid = [r if isinstance(r, tuple) else (r,) for r in raw]
            for i, row in enumerate(grid):
                cmap = {}
                for j, v in enumerate(row):
                    if v is None or v == '':
                        continue
                    cmap[_col_letter(c0 + j)] = _cell_str(v)
                if cmap:
                    cells.append({"r": r0 + i, "c": cmap})
            if nr > max_rows:
                # first_row / last_read_row は「絶対行番号」。セルは r0+i の絶対行で
                # 格納されるので、read_rows（＝読んだ行数）だけを持たせると
                # snapshot-diff 側が行数と絶対行番号を取り違える
                # （UsedRange が1行目から始まらないシートで、読み込み済みの実差分まで
                #   捨てて「差分なし」と嘘をつく。2026-07-14 実弾で確認）
                info["cells_truncated"] = {
                    "read_rows": max_rows, "total_rows": nr,
                    "first_row": r0, "last_read_row": r0 + read_rows - 1,
                }
        info["cells"] = cells

        # --- 結合（①の素を流用）---
        if ur is not None:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                info["merged_skipped_cells"] = skipped
            elif areas:
                info["merged"] = areas

        # --- 図形／ボタン（text・座標・実行マクロ。グループは1段展開）---
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception as e:
            # 黙って落とすと「図形なし」と見分けがつかず、後の snapshot-diff で
            # 図形の消失/追加を取り違える。読めなかった事実を JSON とサマリに残す
            info["shapes_error"] = str(e)
        if shapes:
            info["shapes"] = shapes

        # --- 書式（四つ目の目。値・結合・図形だけ見ていたのが従来）---
        if not no_format:
            if ur is not None and nr > 0:
                try:
                    info["format"] = _collect_format(sh, ur, nr, nc)
                except Exception as e:
                    # 読めなかったことを残す（黙って落とすと diff が
                    # 「書式が消えた」と誤報する／「一致」と嘘をつく）
                    info["format_error"] = str(e)

        # --- テーブル（正式な ListObject だけ・推定はしない）---
        tables = []
        try:
            for lo in sh.ListObjects:
                tables.append({"name": lo.Name, "address": lo.Range.Address})
        except Exception as e:
            # 握りつぶすと「テーブルなし」と区別がつかず、diff がテーブルの
            # 追加/削除を誤報する
            info["tables_error"] = str(e)
        if tables:
            info["tables"] = tables

        sheets[sh.Name] = info

    return {"success": True, "book": wb.Name, "active_sheet": active, "sheets": sheets}


def cmd_snapshot(args):
    """アクティブブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ。

    セル(疎)＋結合(merged)＋図形/ボタン(text・座標・OnAction)＋テーブル(ListObject)を
    1ファイルに束ねる。晴美さんのExStruct extract を「開いてるブック相手・その場」で焼き直したもの。
    表かどうかの"推定"はしない（機械的事実だけ吐き、意味付けは読み手のAIがやる＝この子の設計思想）。
    """
    import json
    target_file, rest = parse_target_and_rest(args.posargs)
    only_sheet = rest[0] if rest else getattr(args, 'sheet_opt', None)
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or _LAST_SNAPSHOT_FILE)
    try:
        max_rows = int(getattr(args, 'max_rows', None) or 5000)
    except (TypeError, ValueError):
        print("エラー: --max-rows は数値で指定してください")
        return False
    # 書式に行上限は要らない。行は二分割で降りて「同じ書式が続く塊」で採るため、
    # 行数が増えても往復は log 的にしか増えない（10,004行でも約30往復）。
    # 上限で打ち切ると「上限より下の書式を見ていない」盲点ができ、数万行のシートでは
    # そこが本体になる。速度と盲点のどちらかを選ばせない、が今の設計。

    xl, wb = get_workbook(target_file)
    try:
        doc = _snapshot_doc(wb, only_sheet=only_sheet, max_rows=max_rows,
                            no_format=getattr(args, 'no_format', False))
    except ValueError as e:
        print(e)
        return False
    active = doc["active_sheet"]
    sheets = doc["sheets"]
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)

    # 端末にはサマリだけ（本体JSONは大きくなり得るのでファイルへ）
    print(f"スナップショット: {doc['book']}  → {out_path}")
    print("-" * 60)
    for name, info in sheets.items():
        parts = [info.get("dims", "")]
        parts.append(f"セル{len(info.get('cells', []))}行")
        if "cells_truncated" in info:
            parts.append(f"(全{info['cells_truncated']['total_rows']}行→{info['cells_truncated']['read_rows']}打切)")
        if "merged" in info:
            parts.append(f"結合{len(info['merged'])}")
        elif "merged_skipped_cells" in info:
            parts.append("結合未走査(大)")
        if "shapes" in info:
            parts.append(f"図形{len(info['shapes'])}")
        if "shapes_error" in info:
            parts.append("図形読取不可")
        if "tables" in info:
            parts.append(f"表{len(info['tables'])}")
        if (info.get("format") or {}).get("rows_coarse"):
            parts.append(f"書式粗{len(info['format']['rows_coarse'])}範囲")
        mark = '*' if name == active else ' '
        print(f"{mark} {name}: " + "  ".join(p for p in parts if p))
    print("-" * 60)
    # 図形を列挙できなかったシートは警告として明示する（「図形0」に見せない）
    shape_err = [(n, i["shapes_error"]) for n, i in sheets.items() if "shapes_error" in i]
    if shape_err:
        print("⚠ 次のシートは図形を読み取れませんでした（このJSONでは図形なしと区別できません）:")
        for n, e in shape_err:
            print(f"   {n}: {e}")
        print("-" * 60)
    # 書式の走査が予算打ち切りで粗い粒度のまま残ったシートも明示する
    # （JSON には rows_coarse で残るが、端末で黙っていると「全部見た」に読める）
    fmt_coarse = [(n, (i.get("format") or {}).get("rows_coarse"))
                  for n, i in sheets.items()
                  if (i.get("format") or {}).get("rows_coarse")]
    if fmt_coarse:
        print("⚠ 次のシートは書式の走査が粗い粒度で打ち切られています"
              "（その範囲の行の書式は塊のままです）:")
        for n, rngs in fmt_coarse:
            head = '、'.join(f"{a}〜{b}行" for a, b in rngs[:5])
            more = len(rngs) - 5
            print(f"   {n}: {head}" + (f" 他{more}件" if more > 0 else ""))
        print("-" * 60)
    print("このJSONを read して質問すれば、開いたままブックLMになる。")
    return True


def _sheet_first_row(info):
    """シートの UsedRange 先頭行（絶対行番号）。取れなければ None。

    新しい snapshot は cells_truncated.first_row を持つ。古い JSON にはないので
    used のアドレス（"$A$200:$C$699" 等）から復元する。
    """
    t = info.get('cells_truncated') or {}
    fr = t.get('first_row')
    try:
        fr = int(fr)
        if fr > 0:
            return fr
    except (TypeError, ValueError):
        pass
    m = re.search(r'\$?[A-Z]{1,3}\$?(\d+)', str(info.get('used') or ''))
    return int(m.group(1)) if m else None


def _cell_compare_limit(old_info, new_info):
    """snapshot 2つのシート情報から「セルを比較してよい行の上限（絶対行番号）」を返す。

    snapshot は --max-rows で読み込む行を打ち切る（cells_truncated に記録）。
    打ち切りは「そこから下は読んでいない」だけで「空だった」ではないため、
    打ち切り行より下を比較すると片側だけ空＝疑似差分になる。

    ※ 返すのは「絶対行番号」であって行数ではない。セルは r0+i の絶対行で格納
    されるため、read_rows（読んだ行数）をそのまま上限にすると単位が食い違い、
    UsedRange が1行目から始まらないシートでは読み込み済みの実差分まで捨てて
    「差分なし」と嘘をつく（2026-07-14 実弾で確認）。
    上限は last_read_row（＝first_row + read_rows - 1）で判定する。
    """
    limits = []
    for info in (old_info, new_info):
        t = info.get('cells_truncated') or {}
        if not t:
            continue
        last = t.get('last_read_row')
        try:
            last = int(last)
        except (TypeError, ValueError):
            last = None
        if last is None:
            # 古い snapshot（last_read_row を持たない）は used から復元する
            fr = _sheet_first_row(info)
            try:
                rr = int(t.get('read_rows'))
            except (TypeError, ValueError):
                rr = None
            last = (fr + rr - 1) if (fr is not None and rr) else None
        if last is not None and last > 0:
            limits.append(last)
    return min(limits) if limits else None


def cmd_snapshot_diff(args):
    """2つの snapshot JSON を機械的に比較: snapshot-diff <before.json> [after.json]

    checkup を前後に挟む型の「シート側」版＝マクロや手作業が実際に何を変えたかを
    セル・結合・図形・テーブル単位の事実で示す（COM 不要の純粋処理）。
    after 省略時は _last_snapshot.json（直近の snapshot）と比較する。
    評価はしない＝差分という事実だけ並べ、意味付けは読み手がやる（snapshot と同じ思想）。
    """
    import json
    if not args.posargs:
        print("使い方: snapshot-diff <before.json> [after.json]")
        print("  after 省略時は _last_snapshot.json（直近の snapshot）と比較します")
        return False
    old_path = os.path.abspath(args.posargs[0])
    new_path = (os.path.abspath(args.posargs[1]) if len(args.posargs) >= 2
                else _LAST_SNAPSHOT_FILE)
    try:
        # `or 20` だと --max 0（件数だけ見たい）が偽値で既定に化けるため is None 判定
        m_opt = getattr(args, 'max_opt', None)
        max_show = 20 if m_opt is None else int(m_opt)
    except (TypeError, ValueError):
        print("エラー: --max は数値で指定してください")
        return False
    if max_show < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False

    docs = []
    for path in (old_path, new_path):
        if not os.path.exists(path):
            print(f"エラー: ファイルがありません: {path}")
            return False
        try:
            # utf-8-sig: PowerShell の Out-File 等が付ける BOM も受け入れる（BOM無しも可）
            with open(path, 'r', encoding='utf-8-sig') as f:
                d = json.load(f)
        except Exception as e:
            print(f"エラー: JSON を読めません: {path} ({e})")
            return False
        if not isinstance(d, dict) or 'sheets' not in d:
            print(f"エラー: snapshot 形式ではありません（'sheets' がない）: {path}")
            return False
        docs.append(d)
    old, new = docs

    def clip(v, n=40):
        s = str(v).replace('\r', ' ').replace('\n', ' ')
        return s if len(s) <= n else s[:n] + '…'

    def cell_map(info):
        out = {}
        for row in info.get('cells', ()):
            for col, v in row.get('c', {}).items():
                out[(row['r'], col)] = v
        return out

    def addr(key):
        return f"{key[1]}{key[0]}"

    def cell_order(key):
        return (key[0], len(key[1]), key[1])   # 行→列文字（桁→辞書順）で安定表示

    def show(label, items, fmt):
        print(f"  {label}: {len(items)}件")
        for it in items[:max_show]:
            print(f"    {fmt(it)}")
        if len(items) > max_show:
            print(f"    … 他 {len(items) - max_show}件（--max で表示数変更可）")

    print("===== スナップショット差分 =====")
    print(f"  旧: {old.get('book', '?')}  ({old_path})")
    print(f"  新: {new.get('book', '?')}  ({new_path})")
    if old.get('book') != new.get('book'):
        print("  ※ ブック名が異なります（別ブック同士の比較）")

    old_sheets, new_sheets = old['sheets'], new['sheets']
    diff_sheets = 0
    merged_note_hidden = []   # 結合未走査で比較できず、かつ他に差分がなく非表示のシート
    trunc_note_hidden = []    # セル打ち切りで一部しか比較できず、かつ他に差分がなく非表示のシート
    unread_note_hidden = []   # 読み取り失敗で比較を降り、かつ他に差分がなく非表示のシート
    unread_any = []           # 読み取り失敗が1件でもあったシート（最後の総括に出す）
    coarse_any = []           # 書式走査が予算打ち切りで粗い粒度だったシート（総括に出す）

    for name in [n for n in old_sheets if n not in new_sheets]:
        diff_sheets += 1
        print(f"\n- シート削除: {name}")
    for name in [n for n in new_sheets if n not in old_sheets]:
        diff_sheets += 1
        info = new_sheets[name]
        print(f"\n+ シート追加: {name}  ({info.get('dims', '?')})")

    for name in [n for n in old_sheets if n in new_sheets]:
        oi, ni = old_sheets[name], new_sheets[name]

        # 片側でも「読めなかった」snapshot は、空 vs 実データの疑似差分になる。
        # 結合（merged_skipped_cells）で比較を降りるのと同じ扱いにする。
        # 見ていないものを「消えた」と report するのが検分器として最悪の嘘
        cells_unreadable = bool(oi.get('cells_error') or ni.get('cells_error'))
        shapes_unreadable = bool(oi.get('shapes_error') or ni.get('shapes_error'))
        tables_unreadable = bool(oi.get('tables_error') or ni.get('tables_error'))
        # 書式も同じ扱いに合流させる。ここに入れないと「書式を読めていないのに、
        # 他に差分が無いから『差分なし（一致）』」という一番たちの悪い嘘になる
        # （読めなかった／片側だけ --no-format、のどちらも「比較していない」）
        _fmt_err = bool(oi.get('format_error') or ni.get('format_error'))
        _fmt_missing = bool(oi.get('format')) != bool(ni.get('format'))
        unreadable_parts = (['セル'] if cells_unreadable else []) \
            + (['図形'] if shapes_unreadable else []) \
            + (['テーブル'] if tables_unreadable else []) \
            + (['書式'] if (_fmt_err or _fmt_missing) else [])
        if unreadable_parts:
            unread_any.append(f"{name}（{'・'.join(unreadable_parts)}）")

        oc, nc_ = cell_map(oi), cell_map(ni)
        # 片側だけ --max-rows で打ち切った snapshot をそのまま比べると、打ち切り行より
        # 下は「片側だけ空」＝実際は無変更なのに丸ごとセル追加/削除に化ける。
        # 結合（merged_skipped_cells）で比較を降りるのと同じ扱いで、比較する行を
        # 両者の read_rows の小さい方までに切り詰める（切り詰めた事実は必ず表示する）。
        cell_limit = _cell_compare_limit(oi, ni)
        if cell_limit is not None:
            oc = {k: v for k, v in oc.items() if k[0] <= cell_limit}
            nc_ = {k: v for k, v in nc_.items() if k[0] <= cell_limit}
        if cells_unreadable:
            added, removed, changed = [], [], []
        else:
            added = sorted([k for k in nc_ if k not in oc], key=cell_order)
            removed = sorted([k for k in oc if k not in nc_], key=cell_order)
            changed = sorted([k for k in oc if k in nc_ and oc[k] != nc_[k]],
                             key=cell_order)

        merged_unscanned = bool(oi.get('merged_skipped_cells')
                                or ni.get('merged_skipped_cells'))
        if merged_unscanned:
            # 片側でも結合未走査なら差分は出せない（空 vs 実データの疑似差分を出さない）
            m_add, m_del = [], []
        else:
            om = set(oi.get('merged') or [])
            nm = set(ni.get('merged') or [])
            m_add, m_del = sorted(nm - om), sorted(om - nm)

        def shape_map(info):
            out, dup = {}, set()
            for s in info.get('shapes', ()):
                nm2 = s.get('name', '(無名)')
                if nm2 in out:
                    dup.add(nm2)          # Excel は図形名の重複を許す＝黙って落とさず注記する
                out.setdefault(nm2, s)
            return out, dup
        os_, odup = shape_map(oi)
        ns_, ndup = shape_map(ni)
        shape_dup = odup | ndup
        s_add, s_del, s_chg = [], [], []
        if not shapes_unreadable:
            s_add = sorted([n2 for n2 in ns_ if n2 not in os_])
            s_del = sorted([n2 for n2 in os_ if n2 not in ns_])
            for n2 in sorted(set(os_) & set(ns_)):
                a, b = os_[n2], ns_[n2]
                fields = []
                if a.get('text') != b.get('text'):
                    fields.append(f"文字 '{clip(a.get('text'))}'→'{clip(b.get('text'))}'")
                if a.get('onaction') != b.get('onaction'):
                    fields.append(f"OnAction {a.get('onaction')}→{b.get('onaction')}")
                if (a.get('l'), a.get('t')) != (b.get('l'), b.get('t')):
                    fields.append(f"位置 ({a.get('l')},{a.get('t')})→({b.get('l')},{b.get('t')})")
                # w/h を持たない旧版 snapshot と比べるときは大きさ比較を降りる。
                # 降りないと (NonexNone)→(WxH) の疑似差分が図形の数だけ出る
                if (('w' in a or 'h' in a) and ('w' in b or 'h' in b)
                        and (a.get('w'), a.get('h')) != (b.get('w'), b.get('h'))):
                    fields.append(f"大きさ ({a.get('w')}x{a.get('h')})→({b.get('w')}x{b.get('h')})")
                if fields:
                    s_chg.append((n2, fields))

        # --- 書式の差分（四つ目の目）---
        # None は「その範囲の中で不揃い」を意味する（欠測ではない）。
        # 片側でも書式を読めていなければ比較を降りる（疑似差分を出さない）。
        # 判定は上の unreadable_parts と同じものを使う（二重に持つと食い違う）
        fmt_unreadable = _fmt_err
        fmt_missing = _fmt_missing            # 片方だけ --no-format で採った
        of_, nf_ = oi.get('format') or {}, ni.get('format') or {}
        f_chg = []
        fmt_coarse_ranges = []
        fmt_unread_items = set()   # 片側でも "?"（項目単位の読取失敗）だった項目
        if not fmt_unreadable and of_ and nf_:
            def _fmt_label(k):
                return {'bold': '太字', 'size': '文字サイズ', 'font': 'フォント',
                        'italic': '斜体', 'color': '文字色', 'fill': '塗り',
                        'numfmt': '表示形式', 'border': '罫線', 'halign': '横位置',
                        'w': '列幅', 'h': '行高', 'hidden': '非表示'}.get(k, k)

            # シート全体で起きた変化（例: 全体の罫線が消えた）は、各列・各行でも
            # 同じ内容で観測される。それを列ぶん・行ぶん繰り返すと、シート全体の
            # 一撃が数十件に膨れて読めなくなる（＝検分結果として役に立たない）。
            # 「シート全体の変化と同じ(旧値→新値)」の項目は、列・行では繰り返さない。
            # 局所的な変化（その列だけ幅が変わった等）はそのまま出る。
            sheet_chg = {}

            def _diff_fmt(a, b, where, skip_same_as_sheet=False):
                for k in sorted(set(a) | set(b)):
                    va, vb = a.get(k), b.get(k)
                    if va == "?" or vb == "?":
                        # "?" は項目単位の読取失敗マーク。通常値と比較すると
                        # 「列幅 25 → ?」のような疑似差分になる（読めなかった事実は
                        # 差分ではない）。比較を降り、下でまとめて報告する
                        fmt_unread_items.add(f"{where.rstrip(': ')}の{_fmt_label(k)}")
                        continue
                    if va == vb:
                        continue
                    if skip_same_as_sheet and sheet_chg.get(k) == (va, vb):
                        continue          # シート全体の変化で説明がつく＝繰り返さない
                    f_chg.append(f"{where}{_fmt_label(k)} {va} → {vb}")

            os_f, ns_f = of_.get('sheet') or {}, nf_.get('sheet') or {}
            for k in sorted(set(os_f) | set(ns_f)):
                if os_f.get(k) != ns_f.get(k):
                    sheet_chg[k] = (os_f.get(k), ns_f.get(k))
            _diff_fmt(os_f, ns_f, 'シート全体: ')
            oc_f, nc_f = of_.get('cols') or {}, nf_.get('cols') or {}
            for k in sorted(set(oc_f) | set(nc_f), key=lambda s: (len(s), s)):
                _diff_fmt(oc_f.get(k) or {}, nc_f.get(k) or {}, f'{k}列: ',
                          skip_same_as_sheet=True)
            # --- 行の書式（同じ書式が続く塊で持っている）---
            # 前後で塊の切れ目は一致しない（見出しが太字でなくなれば塊は繋がる）。
            # そこで両者の境目を突き合わせ、重なる区間ごとに比べる。
            def _blocks(x):
                bs = x.get('row_blocks')
                if bs:
                    return [(int(b['r0']), int(b['r1']),
                             {k: v for k, v in b.items() if k not in ('r0', 'r1')})
                            for b in bs]
                # 旧形式（行ごとの dict）の snapshot とも比べられるようにする
                rows_old = x.get('rows') or {}
                return [(int(k), int(k), v) for k, v in rows_old.items() if k.isdigit()]

            # 予算打ち切りで粗い粒度のまま残った行範囲（rows_coarse）は、
            # still_mixed の None（=細かく見ていない）を抱えたままなので、
            # そのまま比べると「太字 None → False」のような疑似差分を量産する。
            # 片側でも粗い範囲に重なる区間は比較を降り、事実として報告する
            fmt_coarse_ranges = (
                [(int(a), int(b)) for a, b in (of_.get('rows_coarse') or ())]
                + [(int(a), int(b)) for a, b in (nf_.get('rows_coarse') or ())])

            def _overlaps_coarse(a0, a1):
                return any(not (a1 < c0 or c1 < a0)
                           for c0, c1 in fmt_coarse_ranges)

            ob, nb = _blocks(of_), _blocks(nf_)
            if ob and nb:
                # 両者の境目を集めて区間に割る（重なる範囲だけを比べる＝
                # 片側にしか無い行を「変わった」と誤報しない）
                lo = max(min(b[0] for b in ob), min(b[0] for b in nb))
                hi = min(max(b[1] for b in ob), max(b[1] for b in nb))
                cuts = {lo, hi + 1}
                for r0, r1, _ in ob + nb:
                    if lo <= r0 <= hi:
                        cuts.add(r0)
                    if lo <= r1 + 1 <= hi + 1:
                        cuts.add(r1 + 1)
                edges = sorted(c for c in cuts if lo <= c <= hi + 1)

                def _at(bs, r):
                    for r0, r1, d in bs:
                        if r0 <= r <= r1:
                            return d
                    return None

                for i in range(len(edges) - 1):
                    a0, a1 = edges[i], edges[i + 1] - 1
                    if a0 > a1:
                        continue
                    if _overlaps_coarse(a0, a1):
                        continue      # 粗くしか見ていない範囲＝比較しない（下で必ず報告）
                    da, db = _at(ob, a0), _at(nb, a0)
                    if da is None or db is None:
                        continue
                    where = f'{a0}行: ' if a0 == a1 else f'{a0}〜{a1}行: '
                    _diff_fmt(da, db, where, skip_same_as_sheet=True)

        def table_map(info):
            return {t['name']: t.get('address') for t in info.get('tables', ())}
        ot, nt = table_map(oi), table_map(ni)
        if tables_unreadable:
            t_add, t_del, t_chg = [], [], []
        else:
            t_add = sorted([n2 for n2 in nt if n2 not in ot])
            t_del = sorted([n2 for n2 in ot if n2 not in nt])
            t_chg = sorted([n2 for n2 in ot if n2 in nt and ot[n2] != nt[n2]])

        if fmt_coarse_ranges:
            coarse_any.append(name)
        if fmt_unread_items:
            # 項目単位の読取失敗も「比較できていない」の仲間（総括で嘘をつかない）
            unread_any.append(f"{name}（書式項目の一部）")

        has_diff = any((added, removed, changed, m_add, m_del,
                        s_add, s_del, s_chg, t_add, t_del, t_chg, f_chg,
                        oi.get('dims') != ni.get('dims')))
        if not has_diff:
            # シート自体を表示しない場合も、比較できなかった事実は落とさない
            if unreadable_parts:
                unread_note_hidden.append(f"{name}（{'・'.join(unreadable_parts)}）")
            if merged_unscanned:
                merged_note_hidden.append(name)
            if cell_limit is not None:
                trunc_note_hidden.append(f"{name}（{cell_limit}行まで）")
            continue
        diff_sheets += 1
        print(f"\n* シート: {name}")
        if oi.get('dims') != ni.get('dims'):
            print(f"  使用範囲: {oi.get('dims')} → {ni.get('dims')}")
        if cell_limit is not None:
            total = None
            for _i in (oi, ni):
                t = _i.get('cells_truncated') or {}
                if t.get('total_rows'):
                    total = max(total or 0, int(t['total_rows']))
            tail = f"（全{total}行）" if total else ""
            print(f"  ※ 打ち切られた snapshot のため、セルは {cell_limit}行までしか比較していません"
                  f"{tail}＝それ以降の行の変更は検出できません")
        if unreadable_parts:
            print(f"  ※ {'・'.join(unreadable_parts)}を読めなかった snapshot です"
                  "＝その差分は比較していません（「消えた」ではありません）")
        if merged_unscanned:
            print("  ※ 結合セル未走査の snapshot（範囲が大）＝結合の差分は比較していない")
        if shape_dup:
            print(f"  ※ 同名の図形が複数: {', '.join(sorted(shape_dup))}"
                  "（名前単位の比較のため2つ目以降は対象外）")
        if changed:
            show("セル変更", changed,
                 lambda k: f"{addr(k)}: '{clip(oc[k])}' → '{clip(nc_[k])}'")
        if added:
            show("セル追加", added, lambda k: f"{addr(k)}: '{clip(nc_[k])}'")
        if removed:
            show("セル削除", removed, lambda k: f"{addr(k)}: '{clip(oc[k])}'")
        if m_add:
            show("結合追加", m_add, lambda a: a)
        if m_del:
            show("結合解除", m_del, lambda a: a)
        if s_add:
            show("図形追加", s_add,
                 lambda n2: n2 + (f"「{clip(ns_[n2].get('text'))}」" if ns_[n2].get('text') else ""))
        if s_del:
            show("図形削除", s_del,
                 lambda n2: n2 + (f"「{clip(os_[n2].get('text'))}」" if os_[n2].get('text') else ""))
        if s_chg:
            show("図形変更", s_chg, lambda it: f"{it[0]}: " + " / ".join(it[1]))
        if t_add:
            show("テーブル追加", t_add, lambda n2: f"{n2} ({nt[n2]})")
        if t_del:
            show("テーブル削除", t_del, lambda n2: f"{n2} ({ot[n2]})")
        if t_chg:
            show("テーブル範囲変更", t_chg, lambda n2: f"{n2}: {ot[n2]} → {nt[n2]}")
        if f_chg:
            show("書式変更", f_chg, lambda s: s)
        if fmt_unreadable:
            print("  ※ 書式を読めなかった snapshot です＝書式の差分は比較していません")
        elif fmt_missing:
            print("  ※ 片方の snapshot が --no-format で採られています"
                  "＝書式の差分は比較していません")
        elif fmt_coarse_ranges:
            rng = '、'.join(f"{a}〜{b}行"
                            for a, b in sorted(set(fmt_coarse_ranges))[:5])
            more = len(set(fmt_coarse_ranges)) - 5
            print(f"  ※ 書式の走査が粗い粒度で打ち切られた行範囲があります（{rng}"
                  + (f" 他{more}件" if more > 0 else "")
                  + "）＝その範囲の行の書式差分は比較していません")
        if fmt_unread_items:
            head = '、'.join(sorted(fmt_unread_items)[:5])
            more = len(fmt_unread_items) - 5
            print(f"  ※ 片側で読めなかった書式項目があり、比較していません: {head}"
                  + (f" 他{more}件" if more > 0 else ""))

    if unread_note_hidden:
        print(f"\n※ 読み取りに失敗した snapshot のため、次のシートは比較できていません: "
              f"{', '.join(unread_note_hidden)}")
    if merged_note_hidden:
        print(f"\n※ 結合セル未走査の snapshot のため、次のシートは結合の差分を比較できていません: "
              f"{', '.join(merged_note_hidden)}")
    if trunc_note_hidden:
        print(f"\n※ 打ち切られた snapshot のため、次のシートはセルを途中までしか比較できていません: "
              f"{', '.join(trunc_note_hidden)}")
    if coarse_any:
        print(f"\n※ 書式の走査が粗い粒度で打ち切られたシートがあります"
              f"（その範囲の行の書式差分は比較していません）: {', '.join(coarse_any)}")
    print("\n" + "-" * 60)
    if diff_sheets:
        print(f"差分あり: シート{diff_sheets}枚に変更")
    elif unread_any:
        # 読めなかったシートがある状態で「一致」と言ってはいけない
        print("差分なし（ただし読み取りに失敗した箇所があり、そこは比較できていません）")
    elif coarse_any:
        # 粗くしか見ていない範囲がある状態で「一致」と言い切ってはいけない
        print("差分なし（ただし書式を粗い粒度でしか見ていない行範囲があり、"
              "そこは比較できていません）")
    else:
        # snapshot が見ているのは 値・結合・書式・図形・テーブル。
        # 書式は「シート全体／列／行」の粒度で見ており、セル1つ単位ではない
        # （セル単位は1万セルで33秒かかり実用にならない＝実測）。
        # 見ている範囲を正確に言い、それ以上を保証したように読ませない。
        print("差分なし（値・結合・書式・図形・テーブルの範囲で一致）")
        print("  ※ 書式はシート全体／列／行の単位で比較しています"
              "（同じ行・列の中だけで打ち消し合う変更は検出できません）")
    return True


def cmd_wiring(args):
    """ボタン⇔マクロの配線図: wiring [excel_file] [--json]

    シート上の図形/ボタンに登録された OnAction を全部拾い（グループは1段展開）、
    ブック内のマクロ名簿と突き合わせる。行き先のないボタン＝壊れた配線の検出器。
    call-graph（孤立マクロ）・docs（マクロ→ボタン逆引き）と対になる「ボタン側から見た地図」。
    VBA に触れないブック（保護/VBOM未信頼）でも配線一覧だけは出す（実在確認のみ縮退）。
    別ブック修飾（'秀.xlam'!マクロ 等）の配線は名簿の外＝実在確認せず「外部ブック先」として
    別枠で表示し、壊れた配線には数えない（このブックの名簿で×を付けると誤検出になる）。
    終了コード: テキスト表示は壊れた配線ありで 1（lint 型）。--json は常に 0（broken が判定を運ぶ）。
    """
    import json
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 診断は読むだけ

    # マクロ名簿（Sub/Function 名 → 正式名）
    known = {}
    vba_error = None
    proc_pat = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    try:
        for comp in wb.VBProject.VBComponents:
            if int(comp.Type) not in (1, 2, 3, 100):
                continue
            cm = comp.CodeModule
            n = cm.CountOfLines
            code = cm.Lines(1, n) if n else ""
            for m in proc_pat.finditer(code):
                known.setdefault(m.group(1).lower(), m.group(1))
    except Exception as e:
        vba_error = e

    rows = []
    unread_sheets = []          # 図形を列挙できず、配線図から丸ごと落ちたシート
    for sh in wb.Worksheets:
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception as e:
            # 黙って continue すると、そのシートのボタンが1本も出ないまま
            # 「行き先なし 0本」＝健全という誤った安心につながる
            unread_sheets.append((sh.Name, str(e)))
            continue
        for s in shapes:
            macro = s.get('onaction')
            if not macro:
                continue
            # ブック修飾つき配線。自ブック名なら普通に照合、他ブック(アドイン等)なら
            # このブックの名簿では実在確認できない＝外部ブック先として別枠にする
            book = s.get('onaction_book')
            if book and book.lower() == wb.Name.lower():
                book = None
            resolved = None
            if book is None and vba_error is None:
                hit = known.get(macro.lower())
                if hit is None and '.' in macro:
                    # 「モジュール名.マクロ名」形式（同名マクロがあると Excel が
                    # 自動でこの形式にする）は末尾名で解決する
                    hit = known.get(macro.rsplit('.', 1)[-1].lower())
                resolved = hit is not None
            rows.append({'sheet': sh.Name, 'shape': s.get('name', '(図形)'),
                         'text': s.get('text'), 'group': s.get('group'),
                         'macro': macro, 'book': book, 'resolved': resolved})

    broken = [r for r in rows if r['resolved'] is False]
    external = [r for r in rows if r['book']]

    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "book": wb.Name,
                          "vba_readable": vba_error is None,
                          "wires": rows, "broken": len(broken),
                          "external": len(external),
                          "unread_sheets": [{"sheet": n, "error": e}
                                            for n, e in unread_sheets]},
                         ensure_ascii=False))
        # JSON では broken フィールドが判定を運ぶ。ここで exit 1 にすると
        # success:true の JSON に MCP が「失敗しました」を付けて食い違う
        return True

    def label(r):
        t = f"「{str(r['text'])[:30]}」" if r['text'] else ""
        g = f"（グループ {r['group']} 内）" if r['group'] else ""
        return f"{r['shape']}{t}{g}"

    print(f"===== ボタン⇔マクロ配線図: {wb.Name} =====")
    if vba_error is not None:
        print("※ VBA プロジェクトに触れないため実在確認は未実施（配線の一覧のみ）")
        print(f"   詳細: {vba_error}")
    if unread_sheets:
        print("⚠ 次のシートは図形を読み取れず、配線図に含まれていません（未検査）:")
        for n, e in unread_sheets:
            print(f"   {n}: {e}")
    if not rows:
        print("OnAction が登録された図形/ボタンはありません"
              + ("（ただし上記シートは未検査）" if unread_sheets else ""))
        return True

    cur = None
    for r in rows:
        if r['sheet'] != cur:
            cur = r['sheet']
            print(f"\nシート: {cur}")
        if r['book']:
            mark = "（外部ブック先・このブックの名簿では確認できない）"
            dest = f"{r['book']}!{r['macro']}"
        elif r['resolved'] is True:
            mark = "○"
            dest = r['macro']
        elif r['resolved'] is False:
            mark = "×（存在しない）"
            dest = r['macro']
        else:
            mark = "（未確認）"
            dest = r['macro']
        print(f"  {label(r)} → {dest}  {mark}")

    print("\n" + "-" * 60)
    if vba_error is None:
        line = (f"配線 {len(rows)}本 / 実在 {len(rows) - len(broken) - len(external)}"
                f" / 行き先なし {len(broken)}")
        if external:
            line += f" / 外部ブック先 {len(external)}（確認対象外）"
        print(line)
        if broken:
            print("\n⚠ 行き先のないボタン（OnAction 先のマクロが見つからない）:")
            for r in broken:
                print(f"  {r['sheet']} / {label(r)} → {r['macro']}")
    else:
        print(f"配線 {len(rows)}本（実在確認なし）")
    if unread_sheets:
        # 未検査シートがある以上「行き先なし 0本」は健全の証明にならない
        print(f"※ 図形を読み取れなかったシートが {len(unread_sheets)}枚あります"
              f"（{', '.join(n for n, _ in unread_sheets)}）＝この配線図は全数ではありません")
    print("参考: マクロ側から見た地図は call-graph（孤立検出）/ docs（マクロ→ボタン逆引き）")
    return not broken


@protect_safe
def cmd_screenshot(args):
    """範囲を画像(PNG)として書き出す（目・画像版）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    spec = rest[0] if rest else None
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or LAST_VIEW_FILE)

    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)

    # 対象シートをアクティブにすると CopyPicture が安定する
    try:
        ws.Activate()
    except Exception:
        pass

    # Appearance: xlScreen=1 / xlPrinter=2,  Format: xlBitmap=2 / xlPicture(EMF)=-4147
    # ※ chart へ貼る前に cob.Activate() しないと Paste が無反応で白紙PNGになる（要・最重要）
    # ※ 成否は出力サイズではなく「貼り付け後の Shapes 数」で判定する（白紙でもファイルは生成されるため）
    attempts = [(1, 2), (1, -4147), (2, 2)]
    last_err = None
    for appearance, fmt in attempts:
        cob = None
        try:
            rng.CopyPicture(appearance, fmt)
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            cob = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
            cob.Activate()                      # ← これが無いと貼り付かない
            chart = cob.Chart
            time.sleep(0.3)
            chart.Paste()
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            pasted = chart.Shapes.Count          # 1 以上なら貼り付け成功
            if pasted >= 1:
                if os.path.exists(out_path):
                    os.remove(out_path)
                chart.Export(out_path, "PNG")
                cob.Delete()
                cob = None
                if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                    _screenshot_cleanup(xl, ws)
                    print(f"シート: {ws.Name}   範囲: {rng.Address}")
                    print(f"画像保存: {out_path}")
                    print("（注: 一時グラフの作成を伴うため、ブックの Undo 履歴は消えています）")
                    return True
                last_err = "Export に失敗しました"
            else:
                last_err = "クリップボードから貼り付けできませんでした"
        except Exception as e:
            last_err = str(e)
        finally:
            if cob is not None:
                try:
                    cob.Delete()
                except Exception:
                    pass
        time.sleep(0.4)

    _screenshot_cleanup(xl, ws)
    print(f"エラー: スクリーンショットに失敗しました ({last_err})")
    return False


def _screenshot_cleanup(xl, ws):
    """screenshot の後始末（成功・失敗の両経路で共通）。

    ChartObject 操作で選択が動くため A1 に戻し、CopyPicture で
    クリップボードに残った画像もクリアする（copy-range と対称）。
    """
    try:
        ws.Range("A1").Select()
    except Exception:
        pass
    try:
        xl.CutCopyMode = False
    except Exception:
        pass
def _excel_color_to_hex(color_val) -> Optional[str]:
    """ExcelのBGR整数カラーを #RRGGBB 形式に変換。自動/白/無色は None"""
    try:
        c = int(color_val)
        if c in (16777215, 0xFFFFFF, -4142): # 白 / 自動 / xlNone
            return None
        r = c & 0xFF
        g = (c >> 8) & 0xFF
        b = (c >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def style_summary(ws, ur, nr, nc, areas=None, max_rows=300):
    """書式の目の要約（materials／seiri 用・COM は数回〜行数回）。行のリストを返す。

    style-map（セル 1 つずつ Interior／Font を読む＝行×列×5 往復）を 1 手目に畳むと大きい表で遅いので、
    1 手目は「有るか無いか」だけを範囲 1 回で聞き、番地の要る所は style-map に回す（2026-09-17）。
    - 見出しの結合の親子: 使用範囲の上 3 行にある結合（areas から拾う）と、その直下の行の文字
    - 合計行: 先頭列の下罫線が二重（xlDouble=-4119）の行。行数ぶん聞く（max_rows まで）
    - 塗り・文字色・太字: 範囲 1 回ずつ。None＝混在＝「あり」
    """
    out = []
    try:
        r0, c0 = int(ur.Row), int(ur.Column)
    except Exception:
        return out
    # 見出しの結合の親子
    try:
        tops = []
        for a in (areas or []):
            m = _A1_RECT_RE.match(a)
            if not m or not m.group(3):
                continue
            top, left = int(m.group(2)), _col_num_of(m.group(1))
            bot, right = int(m.group(4)), _col_num_of(m.group(3))
            if top <= r0 + 2:                      # 上 3 行
                tops.append((top, left, bot, right, a))
        for top, left, bot, right, a in sorted(tops)[:8]:
            try:
                txt = str(ws.Cells(top, left).Text or '').strip()
            except Exception:
                txt = ''
            kids = []
            if bot + 1 <= r0 + nr - 1:
                for cc in range(left, right + 1):
                    try:
                        t = str(ws.Cells(bot + 1, cc).Text or '').strip()
                    except Exception:
                        t = ''
                    if t:
                        kids.append(t[:12])
            line = f"  {a}「{txt[:20]}」" if txt else f"  {a}"
            if kids:
                line += " → 下の行: " + " / ".join(kids[:8]) + (" …" if len(kids) > 8 else "")
            out.append(line)
        if out:
            out.insert(0, "見出しの結合（親 → 直下の見出し）:")
    except Exception:
        pass
    # 合計行（二重の下罫線）
    try:
        totals = []
        for i in range(1, min(nr, max_rows) + 1):
            try:
                if ur.Cells(i, 1).Borders(9).LineStyle == -4119:
                    totals.append(r0 + i - 1)
            except Exception:
                continue
        if totals:
            out.append("合計行（下罫線が二重）: 行 " + ", ".join(str(t) for t in totals[:10])
                       + (" …" if len(totals) > 10 else ""))
    except Exception:
        pass
    # 塗り・文字色・太字（範囲 1 回ずつ）
    try:
        marks = []
        ci = ur.Interior.ColorIndex
        if ci is None:
            marks.append("塗りあり（混在）")
        elif int(ci) != -4142:
            marks.append("全体に塗り")
        fc = ur.Font.Color
        if fc is None:
            marks.append("文字色の混在あり")
        fb = ur.Font.Bold
        if fb is None:
            marks.append("太字の混在あり")
        if marks:
            out.append("書式: " + "・".join(marks) + "（番地は style-map で）")
    except Exception:
        pass
    return out


def diagnose_notes(fa, fr, vals, r0, c0, limit=8):
    """診断の目のうち、数式の目（formula_notes）と表の汚れ（dirt_notes）に無い 2 つだけを 1 手目に足す。

    循環参照＝要修正（done を止める側）・外れ値＝気づき。集計漏れ・直書き・形の違い・空白・文字の数字は
    既に materials の別の段が出しているので、ここでは数えない（同じ物を 2 度言わない・2026-09-17）。
    戻り値 (要修正の行, 気づきの行)。vbam_audit が無ければ ([], [])。
    """
    try:
        from vbam_audit import check_circular_references, check_data_cleaner
    except Exception:
        return [], []
    blockers, notices = [], []
    try:
        for it in check_circular_references(fa, r0 - 1, c0 - 1):
            blockers.append(f"{it['cell']}: {it['msg']}")
    except Exception:
        pass
    try:
        for it in check_data_cleaner(vals, r0 - 1, c0 - 1):
            if it.get('type') == 'outlier_value':
                notices.append(f"{it['cell']}: {it['msg']}")
    except Exception:
        pass
    return blockers[:limit], notices[:limit]


def cmd_style_map(args):
    """表・シートの視覚レイアウト＆書式マップ（背景色・フォント・二重罫線・複合見出し）を抽出する"""
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)

    spec = rest[0] if rest else None
    if spec is None:
        ws = wb.ActiveSheet
        rng = ws.UsedRange
    else:
        whole = _whole_sheet_spec(wb, spec, sheet_opt)
        if whole is not None:
            ws = wb.Worksheets(whole)
            rng = ws.UsedRange
        else:
            ws, rng = _resolve_range(xl, wb, spec, sheet_opt)

    rng_addr = rng.Address.replace('$', '')
    n_rows = int(rng.Rows.Count)
    n_cols = int(rng.Columns.Count)
    if n_rows == 1 and n_cols == 1:
        rng = rng.CurrentRegion
        rng_addr = rng.Address.replace('$', '')
        n_rows = int(rng.Rows.Count)
        n_cols = int(rng.Columns.Count)

    # 1. 複合見出し・結合セルの解析 (最上部3行以内)
    header_tree = []
    seen_merges = set()
    scan_header_rows = min(3, n_rows)
    for r in range(1, scan_header_rows + 1):
        for c in range(1, n_cols + 1):
            cell = rng.Cells(r, c)
            try:
                if cell.MergeCells:
                    ma = cell.MergeArea
                    ma_addr = ma.Address.replace('$', '')
                    if ma_addr not in seen_merges:
                        seen_merges.add(ma_addr)
                        txt = str(cell.Text or "").strip()
                        # 直下のセルを子として紐付け
                        children = []
                        if r < n_rows:
                            sub_r = r + int(ma.Rows.Count)
                            if sub_r <= n_rows:
                                for sub_c in range(c, c + int(ma.Columns.Count)):
                                    sub_cell = rng.Cells(sub_r, sub_c)
                                    sub_txt = str(sub_cell.Text or "").strip()
                                    sub_addr = sub_cell.Address.replace('$', '')
                                    children.append(f"{sub_addr}「{sub_txt}」" if sub_txt else sub_addr)
                        header_tree.append({
                            "parent": f"{ma_addr}「{txt}」" if txt else ma_addr,
                            "children": children
                        })
            except Exception:
                pass

    # 2. 書式・スタイルのスキャン（有色セル、赤字、太字、二重下線）
    highlights = []
    total_rows = []
    scan_limit = min(n_rows, 500) # パフォーマンス上限

    for r in range(1, scan_limit + 1):
        # 行単位で二重下線（合計行）の判定
        try:
            sample_cell = rng.Cells(r, 1)
            # xlEdgeBottom = 9, xlDouble = -4119
            if sample_cell.Borders(9).LineStyle == -4119:
                total_rows.append(r)
        except Exception:
            pass

        for c in range(1, n_cols + 1):
            cell = rng.Cells(r, c)
            try:
                addr = cell.Address.replace('$', '')
                txt = str(cell.Text or "").strip()
                bg_hex = _excel_color_to_hex(cell.Interior.Color)
                f_color = cell.Font.Color
                f_hex = _excel_color_to_hex(f_color)
                is_bold = bool(cell.Font.Bold)

                # 背景色あり、または赤字・太字・警告色
                # 赤系の判定: Rが大きくG/Bが小さい
                is_red = False
                if f_hex:
                    cr = int(f_hex[1:3], 16)
                    cg = int(f_hex[3:5], 16)
                    cb = int(f_hex[5:7], 16)
                    if cr > 180 and cg < 100 and cb < 100:
                        is_red = True

                if bg_hex or is_red or (is_bold and r > scan_header_rows):
                    features = []
                    if bg_hex:
                        features.append(f"背景:{bg_hex}")
                    if is_red:
                        features.append(f"赤字:{f_hex}")
                    elif f_hex and f_hex not in ("#000000",):
                        features.append(f"文字色:{f_hex}")
                    if is_bold and r > scan_header_rows:
                        features.append("太字")
                    label = f"「{txt[:15]}」" if txt else ""
                    highlights.append(f"{addr} ({'・'.join(features)}) {label}")
            except Exception:
                pass

    # 3. レポート整形
    lines = []
    lines.append(f"╔═══════════════════════════════════════════════════════════════╗")
    lines.append(f"║ 🎨 視覚レイアウト＆書式マップ (Style Map)                     ║")
    lines.append(f"╚═══════════════════════════════════════════════════════════════╝")
    lines.append(f"対象: {ws.Name}!{rng_addr} （{n_rows}行 × {n_cols}列）\n")

    # 見出し階層ツリー
    if header_tree:
        lines.append("📌 【複合見出し・結合セル構造】")
        for node in header_tree:
            lines.append(f"  • {node['parent']}")
            if node["children"]:
                lines.append(f"      └─ " + " / ".join(node["children"]))
        lines.append("")
    else:
        lines.append("📌 【見出し構造】 単一行見出し（結合セルなし）\n")

    # 合計行・二重下線
    if total_rows:
        lines.append("📊 【集計・合計行（二重下線検知）】")
        for tr in total_rows:
            row_addr = f"{rng.Cells(tr, 1).Address.replace('$', '')}:{rng.Cells(tr, n_cols).Address.replace('$', '')}"
            lines.append(f"  • {row_addr} （二重下線＝合計/最終集計行）")
        lines.append("")

    # 強調・着色セル
    if highlights:
        lines.append(f"💡 【強調・着色・警告セル】（計 {len(highlights)} 箇所）")
        for hl in highlights[:15]:
            lines.append(f"  • {hl}")
        if len(highlights) > 15:
            lines.append(f"  … ほか {len(highlights) - 15} 箇所")
        lines.append("")
    else:
        lines.append("💡 【強調・着色セル】 特殊な背景色や警告文字色は検出されませんでした（標準書式）\n")

    if getattr(args, 'json', False):
        import json
        print(json.dumps({
            "sheet": ws.Name,
            "range": rng_addr,
            "rows": n_rows,
            "cols": n_cols,
            "header_tree": header_tree,
            "total_rows": total_rows,
            "highlights": highlights
        }, ensure_ascii=False, indent=2))
        return True

    print("\n".join(lines))
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    return True



__all__ = [
    'cmd_style_map', 'style_summary', 'diagnose_notes',
    'LAST_VIEW_FILE',
    'cmd_seiri',
    'error_hint',
    '_macro_book',
    '_CELL_REF_RE',
    '_cell_compare_limit',
    '_sheet_first_row',
    '_collect_shapes',
    '_disp_pad',
    '_disp_truncate',
    '_disp_width',
    '_merged_areas_in_range',
    '_print_long_cells',
    '_resolve_range',
    '_screenshot_cleanup',
    '_shape_text',
    '_show_range',
    '_hash_scan',
    '_snapshot_doc',
    '_text_to_grid',
    '_values_to_grid',
    '_whole_sheet_spec',
    '_grid_of_value',
    '_guess_header_idx',
    '_date_col_formats',
    '_body_look_mixed',
    'dirt_notes',
    'formula_notes',
    'formula_copy_blockers',
    'fml_agg_ranges',
    'fml_magic_numbers',
    '_fml_rc',
    '_FML_LENS_MAX_CELLS',
    '_FML_BULK_MAX_CELLS',
    '_formula_patterns',
    '_hidden_count',
    'cmd_materials',
    'cmd_read_range',
    'cmd_read_selection',
    'cmd_screenshot',
    'cmd_sheet_info',
    'cmd_snapshot',
    'cmd_snapshot_diff',
    'cmd_wiring',
]
