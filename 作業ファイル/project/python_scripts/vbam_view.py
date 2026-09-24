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
    only = getattr(args, 'sheet_opt', None)          # --sheet で 1 枚だけ（2026-09-24: 不明な引数で落ちていた）
    if only and not any(str(sh.Name) == only for sh in wb.Sheets):
        print(f"エラー: シート '{only}' が {wb.Name} にありません（シート: {', '.join(str(sh.Name) for sh in wb.Sheets)}）")
        return False
    print(f"ブック: {wb.Name}")
    print(f"シート数: {wb.Sheets.Count}   アクティブ: {active}")
    print("-" * 60)
    fast = getattr(args, 'fast', False)
    for sh in wb.Sheets:
        if only and str(sh.Name) != only:
            continue
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
    """格子の見出し行（0 始まり）を推定。文字のセルが 2 つ以上で値の過半を占め、次の行にも値が 2 つ以上ある最初の行。

    2026-09-23: 前は「値が全部文字」の行だけを見出しと見ていたので、見出しの行の端にメモ用の式（テスト用4 の M5
    =COUNTA(会員名簿)＝45）が 1 つあるだけで見出しを見失い、題・例の文と見出しの間の空行を「表の中の空行」と言った。
    """
    for i in range(len(grid) - 1):
        filled = [v for v in grid[i] if not _blank_cell(v)]
        texts = [v for v in filled if isinstance(v, str)]
        if len(texts) >= 2 and len(texts) * 2 > len(filled):
            if len([v for v in grid[i + 1] if not _blank_cell(v)]) >= 2:
                return i + 1 if _is_second_header_row(grid[i], grid[i + 1]) else i
    return None


def _is_second_header_row(top, low):
    """low が二段見出しの下の段か（純 Python・2026-09-24 通しの実測 5）。

    上の段の空き（上期・下期の結合の右側）を下の段（4月・5月…）が埋め、上の段にある所（コード・品目の縦の結合）は
    下の段が空き、下の段が文字だけ。前は上の段だけを見出しと見て、下の段を本文に数え「表の中の空欄 A4 B4…」
    「書式がセルごとに違う」と言っていた。名簿のように文字だけの明細は、見出しの下に空きが無いので当たらない。
    """
    n = max(len(top), len(low))
    tv = [top[j] if j < len(top) else None for j in range(n)]
    lv = [low[j] if j < len(low) else None for j in range(n)]
    low_filled = [v for v in lv if not _blank_cell(v)]
    if len(low_filled) < 2 or not all(isinstance(v, str) for v in low_filled):
        return False
    fills_gap = any(_blank_cell(tv[j]) and not _blank_cell(lv[j]) for j in range(n))
    leaves_gap = any(not _blank_cell(tv[j]) and _blank_cell(lv[j]) for j in range(n))
    return fills_gap and leaves_gap


_BODY_TOTAL_RE = re.compile(r'^(合計|計|総計|小計|中計|総合計|平均)$')
_TOTAL_TAILS = ('小計', '合計', '総計', '中計')


def _is_total_label(v):
    """集計の行の目印の語か（棚の VBA「集計の語か」と同じ規則・純 Python）。

    「総務課 小計」「合計（税込）」のように前に語・後ろに括弧書きが付いたものも集計の行と見る。
    ぴったり「小計」だけを見ていて、小計の行を本文に数え「表の中の空欄」「太字を外せ」と勧めていた（2026-09-24 通しの実測 4）。
    「合計請求書」のように語で始まる文は拾わない（末尾で見る）。
    """
    if not isinstance(v, str):
        return False
    s = v.strip().replace(' ', '').replace('　', '')
    cut = [k for k in (s.find('（'), s.find('(')) if k > 0]
    if cut:
        s = s[:min(cut)]
    if not s or len(s) > 12:
        return False
    return bool(_BODY_TOTAL_RE.match(s)) or s.endswith(_TOTAL_TAILS)


def _table_body_rows(rows, body, n_head):
    """見出しの下の本文の行（grid の index）を、表の終わりまでに絞る（純 Python）。

    合計・小計・平均の行は本文に数えない。空行の後は、見出しの 6 割以上に値がある行（表の中の空行の後の明細）だけ
    本文に戻し、それより値の少ない行（表の下のメモ）が来たら終わる＝棚の VBA「表の本文の最終行」と同じ考え。
    """
    need = max(2, n_head * 0.6)
    out, prev = [], None
    for i in body:
        row = rows[i]
        if any(_is_total_label(v) for v in row):
            prev = i
            continue
        filled = sum(1 for v in row if not _blank_cell(v))
        if prev is not None and (i > prev + 1 or (out and prev != out[-1])) and filled < need:
            break                          # 空行か合計の行の後の、値の少ない行＝表の下のメモ
        out.append(i)
        prev = i
    return out


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


def _num_col_format_odd(ws, grid, r0, c0, header_idx, limit=6):
    """数の列で、表示形式が 1〜少数のセルだけ違う → [(番地, そのセルの形, 多数の形, 多数の数)]。

    列まとめての NumberFormat が None（混在）の列だけセルごとに見る。集計の行は数えない。
    2026-09-24 通しの実測 5: 前年比の列で Q5 だけ 0%（他は 0.00）なのを誰も言わず、手で見つけた。
    """
    out = []
    h = -1 if header_idx is None else header_idx
    width = max((len(r) for r in grid), default=0)
    for j in range(width):
        idx = [i for i in range(h + 1, len(grid))
               if j < len(grid[i]) and isinstance(grid[i][j], (int, float)) and not isinstance(grid[i][j], bool)
               and not any(_is_total_label(v) for v in grid[i])]
        if len(idx) < 3:
            continue
        try:
            if ws.Range(ws.Cells(r0 + idx[0], c0 + j), ws.Cells(r0 + idx[-1], c0 + j)).NumberFormat is not None:
                continue
        except Exception:
            continue
        fm = {}
        for i in idx[:200]:
            try:
                fm.setdefault(str(ws.Cells(r0 + i, c0 + j).NumberFormat), []).append(i)
            except Exception:
                pass
        if len(fm) < 2:
            continue
        top = max(fm, key=lambda k: len(fm[k]))
        if len(fm[top]) < 3 or len(fm[top]) < len(idx) * 0.6:
            continue
        for f, rows in fm.items():
            if f != top:
                out += [(f"{_col_letter(c0 + j)}{r0 + i}", f, top, len(fm[top])) for i in rows]
    return out[:limit]


def _safe_fmt_odd(ws, grid, r0, c0, hidx):
    try:
        return _num_col_format_odd(ws, grid, r0, c0, hidx)
    except Exception:
        return None


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
    # 小計・合計の行の太字・塗りは強調＝汚れではない。明細の行だけを塊で見る（2026-09-24 通しの実測 4 で「太字を外せ」と勧めた）
    blocks, start = [], None
    for i in range(filled[0], filled[-1] + 2):
        is_sum = i > filled[-1] or any(_is_total_label(v) for v in rows[i])
        if is_sum:
            if start is not None:
                blocks.append(ws.Range(ws.Cells(r0 + start, c0), ws.Cells(r0 + i - 1, c0 + width - 1)))
            start = None
        elif start is None:
            start = i
    if not blocks:
        return None, []
    look = blocks[0]
    for b in blocks[1:]:
        look = ws.Application.Union(look, b)
    f = look.Font
    bad = []
    for name, v in (("フォント名", f.Name), ("大きさ", f.Size), ("文字色", f.Color), ("太字", f.Bold),
                    ("斜体", f.Italic), ("下線", f.Underline), ("取り消し線", f.Strikethrough),
                    ("塗り", look.Interior.ColorIndex)):
        if v is None:
            bad.append(name)
    return str(rng.Address).replace('$', ''), bad


def dirt_notes(grid, r0=1, c0=1, header_idx=None, col_formats=None, limit=12, look_mixed=None, hints=None, fmt_odd=None):
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
    # 郵便番号・電話は番号の列でも同じ値が当たり前（同じ町・同じ会社の代表番号）。重複とは言わない
    # （2026-09-23 通しの実測 3: 名簿の 010-0951 の 2 人を「番号列の重複」と言っていた）
    for j in sorted(j for j in id_cols if not re.search(r'郵便|〒|電話|TEL|携帯|FAX', str(heads[j]), re.I)):
        seen, col_dups, filled = {}, [], 0
        for i in body:
            v = rows[i][j]
            if _blank_cell(v):
                continue
            filled += 1
            k = str(norm(v))
            if k in seen:
                if i not in dup_row_set:
                    col_dups.append(f"{addr(i, j)} = {addr(seen[k], j)}（{k}）")
            else:
                seen[k] = i
        # 繰り返すのが当たり前の列（明細の取引先コード・課コード）は重複と言わない。番号の重複は「まれに 1〜2 件」
        # （2026-09-23 通しの実測 2: 伝票の取引先コード C001〜C004 の繰り返し 8 件を「番号列の重複」と並べていた）
        if col_dups and len(col_dups) <= max(2, filled // 5):
            dup_keys += col_dups
    if dup_keys:
        out.append("  番号列の重複: " + show(dup_keys))
    zen, han, edge, only_ws, multi, numstr, datestr, sep = [], [], [], [], [], [], [], {}
    astext = []          # 文字として入った式（'=' で始まる文字。計算されていない・2026-09-09）
    date_cols = set()    # 文字の日付のある列（棚の 選んだ列の文字の日付を日付にする に渡す）
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
            elif _DATELIKE_STR_RE.match(core) or re.fullmatch(r'\d{1,2}\s?月\s?\d{1,2}\s?日', core):
                datestr.append(a)            # 年の無い「5月16日」も文字の日付（2026-09-23 通しの実測 2）
                date_cols.add(j)
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
    if fmt_odd:
        # 数の列で 1〜少数のセルだけ表示形式が違う（呼び手が _num_col_format_odd で集める・2026-09-24）
        parts = [f"{a}（{f}／他 {n} 行は {top}）" for a, f, top, n in fmt_odd]
        out.append("  数の表示形式が 1 セルだけ違う: " + "／".join(parts)
                   + f"  → format-range {fmt_odd[0][0]} --number-format \"{fmt_odd[0][2]}\" のように多数の形にそろえる")
    if look_mixed and look_mixed[1]:
        out.append(f"  文字の書式がセルごとに違う（本文 {look_mixed[0]}）: {'・'.join(look_mixed[1])}"
                   "  → format で本文全体に font・size・color と \"plain\": true・\"bg\": \"none\"・\"unbold\": true")
    blanks = []
    if h >= 0 and body:
        # 本文の終わりまでで数える。合計・平均の行と、その後のメモの行の空きは表の中の空欄ではない
        # （2026-09-23 Gemini の試し: テスト用1 で 合計・平均・メモの行 A27 A28 B30… を「表の中の空欄」と出した）
        n_head0 = sum(1 for v in heads if not _blank_cell(v))
        core = _table_body_rows(rows, body, n_head0)
        # 左右に並んだ表は、見出しの空いた列で塊に分け、塊ごとの最後の行までで数える
        # （テスト用4: 右の申込一覧は 11 行目で終わるのに、左の名簿の 13 行目までの空きを「表の中の空欄」と出した）
        block_last, start = {}, None
        for j in range(width + 1):
            if j < width and not _blank_cell(heads[j]):
                start = j if start is None else start
                continue
            if start is not None:
                last = max((i for i in core if any(not _blank_cell(rows[i][k]) for k in range(start, j))), default=-1)
                for k in range(start, j):
                    block_last[k] = last
                start = None
        for j in range(width):
            if _blank_cell(heads[j]):
                continue
            cells = [(i, rows[i][j]) for i in core if i <= block_last.get(j, -1)]
            if not cells:
                continue
            filled = sum(1 for _i, v in cells if not _blank_cell(v))
            if filled < 3 or filled / len(cells) < 0.6:
                continue
            blanks += [addr(i, j) for i, v in cells if _blank_cell(v)]
    if blanks:
        out.append("  表の中の空欄: " + show(blanks))
    # ---- 2026-09-23 通しの実測 2 で手で見つけていた 3 つ（空行・ゼロの消えた番号・空白の有無の揺れ）----
    n_head = sum(1 for v in heads if not _blank_cell(v)) if h >= 0 else width
    blank_rows = []
    if body and h >= 0:            # 見出しが分からない表では数えない（題・例の文と見出しの間の空行を「表の中」と言った・2026-09-23 テスト用4）
        for i in range(body[0] + 1, body[-1]):
            if i in body:
                continue
            nxt = next((k for k in body if k > i), None)
            # 空行の下が本文の行（見出しの 6 割以上に値）なら表の中の空行。合計・メモの前の空行は数えない
            # （合計・平均の行は値が多いので本文と見ていた＝明細と合計の間の空行を「詰める」と勧めた・2026-09-23）
            if nxt is not None and not any(_is_total_label(v) for v in rows[nxt]) \
                    and sum(1 for v in rows[nxt] if not _blank_cell(v)) >= max(2, n_head * 0.6):
                blank_rows.append(r0 + i)
    if blank_rows:
        out.append("  表の中の空行: " + " ".join(f"行{x}" for x in blank_rows[:limit]))
    zero_lost = []
    for j in range(width):
        cells = [(i, rows[i][j]) for i in body if not _blank_cell(rows[i][j])]
        lens = {}
        for _i, v in cells:
            if isinstance(v, str) and re.fullmatch(r'0\d+', v.strip()):
                lens[len(v.strip())] = lens.get(len(v.strip()), 0) + 1
        if not lens:
            continue
        L, nL = max(lens.items(), key=lambda kv: kv[1])
        if nL < 2 or nL < len(cells) * 0.6:
            continue
        for i, v in cells:
            s = (str(int(v)) if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and v >= 0
                 else (v.strip() if isinstance(v, str) else ''))
            if s.isdigit() and len(s) < L:
                zero_lost.append(f"{addr(i, j)}（{s}・列は {L} 桁）")
    if zero_lost:
        out.append("  先頭のゼロが消えた番号: " + show(zero_lost))
    space_var = []
    for j in range(width):
        forms = {}
        for i in body:
            v = rows[i][j]
            if isinstance(v, str) and v.strip():
                forms.setdefault(re.sub(r'[ 　]', '', v), {}).setdefault(v, []).append(i)
        for _k, fs in forms.items():
            if len(fs) < 2:
                continue
            top = max(fs, key=lambda f: (len(fs[f]), -len(f)))
            space_var += [f"{addr(i, j)}（{f}→{top}）" for f, idx in fs.items() if f != top for i in idx]
    if space_var:
        out.append("  空白の有無だけ違う書き方: " + show(space_var))
    # ---- 2026-09-23 通しの実測 3（名簿）で言えなかった 4 つ（番号の抜け・未来の生年月日・電話の桁・同姓同名）----
    seq_cols = [j for j in id_cols if isinstance(heads[j], str)
                and re.fullmatch(r'(?i)no\.?|№|番号|連番|通し番号|項番', heads[j].strip())]
    seq_gaps = []
    for j in seq_cols[:1]:
        nums = []
        for i in body:
            v = rows[i][j]
            if isinstance(v, str) and v.strip().isdigit():
                v = int(v.strip())
            if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer():
                nums.append((i, int(v)))
        if len(nums) < 3 or len(nums) < len(body) * 0.8:
            continue
        ups = sum(1 for (_a, x), (_b, y) in zip(nums, nums[1:]) if y > x)
        if ups < (len(nums) - 1) * 0.7:          # 並びが上がっていない列は連番ではない
            continue
        for (_a, x), (b, y) in zip(nums, nums[1:]):
            if y != x + 1:
                seq_gaps.append(f"{addr(b, j)}（{x} の次が {y}）")
    if seq_gaps:
        out.append("  番号の抜け・飛び: " + show(seq_gaps))
    birth_future, today = [], time.localtime()[:3]
    for j in range(width):
        hd = heads[j] if h >= 0 else None
        if not (isinstance(hd, str) and re.search(r'生年月日|誕生|生まれ', hd)):
            continue
        for i in body:
            v = rows[i][j]
            if isinstance(v, datetime.datetime) and (v.year, v.month, v.day) > today:
                birth_future.append(f"{addr(i, j)}（{v.year}/{v.month}/{v.day}）")
    if birth_future:
        out.append("  生年月日が今日より後（打ち間違いか・人の確認が要る）: " + show(birth_future))
    phone_len = []
    for j in range(width):
        hd = heads[j] if h >= 0 else None
        if not (isinstance(hd, str) and re.search(r'電話|TEL|携帯|FAX', hd, re.I)):
            continue
        for i in body:
            v = rows[i][j]
            if isinstance(v, str) and v.strip():
                d = re.sub(r'\D', '', unicodedata.normalize('NFKC', v))
                if d and len(d) not in (10, 11):        # 固定・0120 は 10 桁、携帯は 11 桁
                    phone_len.append(f"{addr(i, j)}（{len(d)} 桁）")
    if phone_len:
        out.append("  電話の桁が合わない（10・11 桁でない＝抜けか打ち間違い・人の確認が要る）: " + show(phone_len))
    same_name = []
    for j in range(width):
        hd = heads[j] if h >= 0 else None
        if not (isinstance(hd, str) and re.fullmatch(r'氏名|名前|お名前|氏名（漢字）', hd.strip())):
            continue
        seen = {}
        for i in body:
            v = rows[i][j]
            if i in dup_row_set or not isinstance(v, str) or not v.strip():
                continue
            k = re.sub(r'[ 　]', '', v)
            if k in seen:
                same_name.append(f"{addr(i, j)} と {addr(seen[k], j)}（{v.strip()}）")
            else:
                seen[k] = i
    if same_name:
        out.append("  同じ氏名で中身が別の行（同姓同名か古い行か・重複ではない・消さない）: " + show(same_name))
    if hints is not None:
        last_row = r0 + (body[-1] if body else 0)
        first_row = r0 + h + 1
        if blank_rows:
            hints.append((2, "表の中の空行を削除して詰める", None, "表の中の空行 " + " ".join(f"行{x}" for x in blank_rows[:6])))
        if dup_rows:
            hints.append((3, "全列が同じ重複行を削除する", None, "重複行 " + dup_rows[0].split("（")[0]))
        if zero_lost:
            hints.append((5, "番号に先頭のゼロを付けてそろえる", None, "先頭のゼロが消えた番号 " + zero_lost[0].split("（")[0]))
        if space_var:
            hints.append((6, "空白の有無を多い方にそろえる", None, "空白の有無だけ違う書き方 " + space_var[0].split("（")[0]))
        for j in sorted(date_cols):
            col = _col_letter(c0 + j)
            hints.append((7, "選んだ列の文字の日付を日付にする", f"{col}{first_row}:{col}{last_row}", f"{col}列の文字の日付"))
        if seq_cols and (seq_gaps or any(k.startswith(_col_letter(c0 + seq_cols[0])) for k in dup_keys)):
            _dup = any(k.startswith(_col_letter(c0 + seq_cols[0])) for k in dup_keys)
            hints.append((8, "番号の列を連番に振り直す", None, f"{heads[seq_cols[0]]}の重複・抜け"
                          + ("（重複は同じ件の二重入力かもしれない＝行の中身を確かめてから撃つ）" if _dup else "")))
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
            # はずれは 4 割未満まで言う（5 行に 1 つまでだと 9 行中 2 行の年計のずれを黙った・2026-09-24 通しの実測 5）
            if cnt[top] < 3 or cnt[top] < len(items) * 0.6:
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
                        print(f"気づき（{'数式' if _n.startswith(('式の列', '集計の')) else '外れ値'}）: {_n}")
            except Exception:
                pass
            hidx = _guess_header_idx(grid)
            try:
                _look = _body_look_mixed(ws, grid, r0, c0, hidx)
            except Exception:
                _look = None
            for ln in dirt_notes(grid, r0, c0, hidx, _date_col_formats(ws, grid, r0, c0, hidx),
                                 look_mixed=_look, fmt_odd=_safe_fmt_odd(ws, grid, r0, c0, hidx)):
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
_SEIRI_TIDY = '表の書き方と罫線と列幅をそろえる'
_SEIRI_DEDUPE = '全列が同じ重複行を削除する'
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
            if module == SHELF_MODULE:           # 棚は「表の整理_〜」に分かれていても探す（2026-09-23）
                hit = shelf_module_of(p, sub) is not None
            else:
                cm = p.VBComponents(module).CodeModule
                n = int(cm.CountOfLines)
                hit = bool(n and re.search(r'^\s*Sub\s+' + re.escape(sub) + r'\b', cm.Lines(1, n), re.M))
            if hit:
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


_TOTAL_WORD_RE = re.compile(r'^(合計|計|総計|小計|総合計)$')


def seiri_formula_hints(fa, fr, grid, r0, c0, hidx):
    """式の側で棚が直せるもの → [(順, マクロ, 選択, 理由)]（純 Python・2026-09-23）。

    式のずれ・式の列の数字の直書き → 選んだ列の途切れた式を戻す（列の本文を選んで）／合計の行の循環参照・集計の取りこぼし → 表の下に合計行を足す。
    """
    try:
        from vbam_audit import check_formula_linter
        issues = check_formula_linter(fa, fr, grid, r0 - 1, c0 - 1)
    except Exception:
        return []
    rows = [list(r) for r in (grid or [])]
    first = r0 + (hidx if hidx is not None else 0) + 1
    last = r0 + len(rows) - 1

    def is_total_row(abs_row):
        i = abs_row - r0
        return 0 <= i < len(rows) and any(isinstance(v, str) and _TOTAL_WORD_RE.match(v.strip().replace(' ', ''))
                                          for v in rows[i])
    out, cols, tot = [], {}, []
    for it in issues:
        m = re.match(r'([A-Z]+)(\d+)$', it.get('cell', ''))
        if not m:
            continue
        col, row = m.group(1), int(m.group(2))
        if it['type'] in ('inconsistent_formula', 'value_in_formula_column') and not is_total_row(row):
            cols.setdefault(col, []).append(it['cell'])
        elif it['type'] in ('circular_reference', 'omitted_sum_range') and is_total_row(row):
            tot.append(it['cell'])
    for col, cells in sorted(cols.items()):
        # 選ぶのは、その列で式か数の入った明細の行だけ（見出しの 2 段目・合計の行を含めない。P4:P15 と出して
        # 合計の行まで明細の式で上書きさせるところだった・2026-09-24 通しの実測 5）
        ci = 0
        for ch in col:
            ci = ci * 26 + ord(ch) - 64
        ci -= c0
        body = [r0 + i for i, row in enumerate(fa or [])
                if first <= r0 + i <= last and not is_total_row(r0 + i) and 0 <= ci < len(row or [])
                and row[ci] not in (None, "")]
        lo, hi = (min(body), max(body)) if body else (first, last)
        out.append((1, "選んだ列の途切れた式を戻す", f"{col}{lo}:{col}{hi}", f"{col}列の式のずれ・直書き " + " ".join(cells[:4])))
    if tot:
        out.append((4, "表の下に合計行を足す", None, "合計の行の式（循環参照・集計の取りこぼし） " + " ".join(sorted(set(tot))[:4])))
    return out


def print_seiri_hints(hints):
    """棚で直せる手を、撃つ順に並べて出す（式を先に揃えてから行を消す＝消した行を指す式が #REF! にならない）。"""
    if not hints:
        return
    seen, lines = set(), []
    for _o, name, sel, why in sorted(hints, key=lambda h: h[0]):
        key = (name, sel)
        if key in seen:
            continue
        seen.add(key)
        lines.append(f"  shelf-run {name}" + (f" --select {sel}" if sel else "") + f"   ← {why}")
    print("棚で直せる手（この順に撃つ。式を先に揃えてから行を消す。人の判断が要るもの＝マイナス・桁違い・番号の重複・"
          "式の中の数は、撃たずに報告）:")
    for ln in lines:
        print(ln)


def _seiri_print_changes(ws, before_snap, backup_path, show=12):
    """seiri のマクロが直したセル（番地: 前→後）と書式の変化の件数を出す。数えられなければ黙る。"""
    if not before_snap or before_snap.get('too_big'):
        return
    try:
        from vbam_undo import _changes_of, _changes_table, _format_changes, _write_changes_file
        rows = _changes_of(before_snap, ws, str(ws.Name)) or []
        fmt_rows = _format_changes(before_snap.get('format'), ws) or []
    except Exception:
        return
    if rows:
        print(f"マクロが直したセル: {len(rows)} 個")
        for line in _changes_table(rows, show=show):
            print(line)
    else:
        print("マクロが直したセル: なし（値と式はそのまま）")
    if fmt_rows:
        print(f"書式の変化: {len(fmt_rows)} 件（明細は agent --changes）")
    try:
        _write_changes_file(rows, fmt_rows=fmt_rows)
    except Exception:
        pass
    if backup_path and (rows or fmt_rows):
        print("（戻すなら agent --undo）")


def cmd_seiri(args):
    """表を直す 1 手目: seiri [--dedupe]（2026-09-13・shu「まとめてみろ」）

    materials で表の全体を読んでから、次の往復でマクロと書き込み、の 2 往復と読む時間を畳む。
    画面のシートに「表の書き方と罫線と列幅をそろえる」マクロ（開いているブックかアドインのモジュール「表の整理」）を撃ち、
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
        # マクロが直した所を出す＋控えを取る（2026-09-23 の通しの実測: 文字の日付・全角の数字・半角カナ等 6 か所を
        # 黙って直し、報告に 1 つも出なかった＝何が変わったか分からず、戻す手も無かった）
        before_snap, backup_path = None, None
        try:
            from vbam_undo import _agent_backup, _sheet_snapshot
            before_snap = _sheet_snapshot(ws)
            backup_path = _agent_backup(wb, ws, str(ws.Name), request="seiri")
        except Exception:
            pass
        t0 = time.time()
        try:
            wb.Activate()
            ws.Activate()
            from vbam_vba import run_book_macro
            for nm in names:
                run_book_macro(xl, owner, nm)      # 実行時エラーで窓を出して止まらない（2026-09-24 総点検）
            print(f"マクロ: {' → '.join(names)}（{owner}・{time.time() - t0:.2f} 秒）")
        except Exception as e:
            print(f"マクロを撃てませんでした: {e}（残りだけ出します）")
        _seiri_print_changes(ws, before_snap, backup_path)
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
                    print(f"気づき（{'数式' if _n.startswith(('式の列', '集計の')) else '外れ値'}）: {_n}")
            except Exception:
                pass
            hidx = _guess_header_idx(grid)
            try:
                look = _body_look_mixed(ws, grid, r0, c0, hidx)
            except Exception:
                look = None
            hints = []
            for ln in dirt_notes(grid, r0, c0, hidx, _date_col_formats(ws, grid, r0, c0, hidx), look_mixed=look,
                                 hints=hints, fmt_odd=_safe_fmt_odd(ws, grid, r0, c0, hidx)):
                print(ln)
            try:
                _areas, _sk = _merged_areas_in_range(ur)
                for _ln in style_summary(ws, ur, nr, nc, _areas):
                    print(_ln)
            except Exception:
                pass
            # 棚で直せる手（2026-09-23・通しの実測で、気づきを見てから棚を探す往復が要っていた）
            try:
                hints += seiri_formula_hints(_fa, _fr, grid, r0, c0, hidx)
                print_seiri_hints(hints)
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


# ----------------------------------------------------------------
# 棚（表の整理）を自由に使う: shelf（目録）・shelf-run（選んで撃つ・差分・控え）
# 2026-09-23・作業ファイル\project\work_order_20260923_shelf_free.md
# ----------------------------------------------------------------
_SHELF_HEADER_RE = re.compile(r"\s*(?:Public\s+)?Sub\s+([^\s(]+)\s*\(\s*\)")
_SHELF_END_RE = re.compile(r"^End\s+Sub\b")
_SHELF_WINDOW_RE = re.compile(r'\b(MsgBox|InputBox)\b')
_SHELF_KEYS = ("依頼の語:", "依頼の組:", "扱う:", "見出し:", "選ぶ列:", "形:")
_SHELF_DATE_NOTE_RE = re.compile(r"（[^（）]*20\d\d-\d\d-\d\d[^（）]*）")
_SHELF_ASK_RE = re.compile(r"^依頼の語\s*[:：]\s*(.+)$")
_SHELF_FP_MAX_CELLS = 50000     # 別シートの書き換わりを見る指紋を取るセル数の上限（超えたら番地だけ比べる）


def _shelf_source(xl):
    """棚（モジュール「表の整理」）のコードを持つブックと行の並び。無ければ (None, None)。

    make_catalog.py（作業ファイル\\project\\fixtures\\shelf_exam）と同じ畳み方を道具側に持たせた。
    固定の shelf_catalog.tsv を鵜呑みにせず、開いているブック・アドインの中身から毎回組み立てる
    （棚が増減しても道具の目録が古くならない）。
    """
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return None, None
    for p in projects:
        try:
            text = shelf_text(p)            # 「表の整理_〜」に分かれていても全部（2026-09-23）
            if not text:
                continue
            from vbam_vba import _project_book_name
            name = _project_book_name(xl, p)
            if name:
                return name, text.split('\r\n')
        except Exception:
            continue
    return None, None


def _shelf_module_name(xl, owner, sub):
    """owner（ブック名）の棚で Sub sub がいるモジュール名（表示用）。分からなければ「表の整理」。"""
    try:
        from vbam_vba import _project_book_name
        for p in xl.VBE.VBProjects:
            if _project_book_name(xl, p) == owner:
                return shelf_module_of(p, sub) or SHELF_MODULE
    except Exception:
        pass
    return SHELF_MODULE


def _parse_shelf_catalog(lines):
    """棚のコード行 → [{'name','sel','kind','desc','ask','window'}, …]（引数なし Sub 単位。純 Python）。

    頭に説明の注記が無い Sub（118 本のうち 38 本・2026-09-23 実測）は、説明の代わりに「依頼の語」を出す。
    空の説明では目録から選べない（試験側の make_catalog.py は補足のファイルで埋めていたが、道具は読んでいなかった）。
    """
    rows, cur = [], None
    for ln in lines:
        m = _SHELF_HEADER_RE.match(ln)
        if m:
            cur = {"name": m.group(1), "sel": "", "kind": "", "desc": [], "ask": "", "body": [], "header": True}
            rows.append(cur)
            continue
        if cur is None:
            continue
        s = ln.strip()
        if _SHELF_END_RE.match(s):
            cur = None
            continue
        cur["body"].append(ln)
        if not cur["header"]:
            continue
        if not s.startswith("'"):
            if s and not s.startswith("Dim"):
                cur["header"] = False
            continue
        t = s[1:].strip()
        ask = _SHELF_ASK_RE.match(t)
        if t.startswith("扱う:"):
            cur["kind"] = t[3:].strip()
        elif t.startswith("選ぶ列:"):
            cur["sel"] = t[4:].strip()
        elif ask:
            cur["ask"] = ask.group(1).strip()
        elif not t.startswith(_SHELF_KEYS) and len(cur["desc"]) < 2 and not t.startswith("-"):
            cur["desc"].append(t)
    out = []
    for r in rows:
        d = _SHELF_DATE_NOTE_RE.sub('', " ".join(r["desc"])).strip()
        if not d and r["ask"]:
            d = "（依頼の語）" + "・".join(w.strip() for w in r["ask"].split("|") if w.strip())
        out.append({"name": r["name"], "sel": r["sel"] or "-", "kind": r["kind"], "desc": d[:120],
                    "ask": r["ask"], "window": bool(_SHELF_WINDOW_RE.search("\n".join(r["body"])))})
    return out


def cmd_shelf(args):
    """`shelf [--grep 語]`: 棚（表の整理）の目録を 1 手で返す（2026-09-23）。

    名前・選ぶ列・窓を出すか（本文に MsgBox/InputBox があるか）・扱う・説明。
    撃つ手は shelf-run。目録どおりに動くとは限らないマクロ側の間違いはここでは直さない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    owner, lines = _shelf_source(xl)
    if not owner:
        print(f"棚（モジュール「{_SEIRI_MODULE}」）が開いているブック・アドインに見つかりません")
        return False
    rows = _parse_shelf_catalog(lines)
    ask = (getattr(args, 'ask', None) or '').strip()
    if ask:
        return _shelf_ask(owner, lines, rows, ask)
    grep = (getattr(args, 'grep', None) or '').strip()
    if grep:
        rows = [r for r in rows if any(grep in r[k] for k in ('name', 'kind', 'desc', 'ask'))]
    print(f"棚: {owner}!{_SEIRI_MODULE}   {len(rows)} 本" + (f"（「{grep}」で絞り込み）" if grep else ""))
    for r in rows:
        win = '窓あり' if r['window'] else '-'
        print(f"  {r['name']}\t選ぶ列={r['sel']}\t{win}\t{r['kind']}\t{r['desc']}")
    print("（撃つ: shelf-run <名前> [--select 範囲] [--sheet 名]）")
    return True


def _shelf_ask(owner, lines, rows, ask):
    """`shelf --ask 依頼文`: 依頼文から撃つマクロを 1 本選ぶ（2026-09-23）。

    採点は Excelコンボの先撃ち（VBA の コンボ道具.先撃ち候補）を移した vbam_prefire.shelf_pick
    （888 問で VBA と食い違い 0）。それまで shelf_pick は試験からしか呼ばれておらず、Claude が棚を選ぶ手に
    つながっていなかった。当たらないのは「棚に無い」か「質問・コードの依頼」＝手で直すか AI の役目。
    """
    import vbam_prefire as vp
    entries = vp.shelf_entries_from_text("\n".join(lines))
    print(f"棚: {owner}!{_SEIRI_MODULE}   依頼「{ask}」")
    # 頼みが 2 つ以上で全部が棚に当たれば、撃つ順に全部を出す（2026-09-24・Excelコンボの 先撃ちの並び と同じ）
    plan = vp.shelf_plan(ask, entries)
    if plan:
        print(f"当たり: {len(plan)} 本（この順に撃つ）")
        for k, nm in enumerate(plan, 1):
            r = next((x for x in rows if x['name'] == nm), None)
            sel = f"　選ぶ列={r['sel']}" if r and r['sel'] != '-' else ""
            win = "　窓あり" if r and r['window'] else ""
            print(f"  {k}. shelf-run {nm}{sel}{win}")
        return True
    pick = vp.shelf_pick(ask, entries)
    r = next((x for x in rows if x['name'] == pick), None) if pick else None
    if not r:
        print("当たり: なし（棚に当たるマクロが無い。質問・コードの依頼も当てない）。語で探すなら shelf --grep 語")
        return True
    win = '窓あり' if r['window'] else '-'
    print(f"当たり: {r['name']}\t選ぶ列={r['sel']}\t{win}\t{r['kind']}\t{r['desc']}")
    how = f"shelf-run {r['name']}"
    if r['sel'] != '-':
        how += " --select 列1,列2…（選ぶ列=" + r['sel'] + "。列は 1 列ずつカンマで分ける）"
    if r['window']:
        how += " [--input-text 答え]"
    print(f"（撃つ: {how}）")
    return True


def _fingerprint_of(fmls, vals):
    """(書いた中身の CRC, 値の CRC)（純 Python）。

    書いた中身＝式のセルは式の字面、値のセルは字面と型（文字の「21」と数の 21 を分ける）。
    値＝計算の結果も含めた見え方。書いた中身が同じで値だけ違う＝式の再計算で変わっただけ（書き換えていない）。
    """
    fr, vr = _rows_of_value(fmls), _rows_of_value(vals)
    written = []
    for i, row in enumerate(fr):
        for j, f in enumerate(row):
            if isinstance(f, str) and f.startswith('='):
                written.append(f)
            else:
                v = vr[i][j] if i < len(vr) and j < len(vr[i]) else None
                written.append((f, type(v).__name__))
    enc = lambda x: zlib.crc32(repr(x).encode('utf-8', 'replace'))  # noqa: E731
    return enc(written), enc(vr)


def _sheet_fingerprint(ws, cap=None):
    """シートの中身の指紋 (使用範囲の番地, 書いた中身の CRC, 値の CRC)。大きすぎれば CRC は None。読めなければ None。"""
    cap = _SHELF_FP_MAX_CELLS if cap is None else cap
    try:
        ur = ws.UsedRange
        addr = str(ur.Address)
        if int(ur.Cells.CountLarge) > cap:
            return (addr, None, None)
        return (addr,) + _fingerprint_of(ur.Formula, ur.Value2)
    except Exception:
        return None


def _shelf_book_state(wb):
    """撃つ前後のブックの姿（2026-09-23）: シートの並び・シートごとの中身の指紋と図形の位置・テーブル・ピボット・
    名前（見えるものだけ）・クエリ。読めない項目は黙って飛ばす（撃つ手を止めない）。"""
    from vbam_undo import _shapes_geometry
    st = {'order': [], 'fp': {}, 'geom': {}, 'tables': set(), 'pivots': set(), 'names': {}, 'queries': set()}
    try:
        st['order'] = [str(s.Name) for s in wb.Sheets]
    except Exception:
        pass
    try:
        sheets = list(wb.Worksheets)
    except Exception:
        sheets = []
    for sh in sheets:
        try:
            n = str(sh.Name)
        except Exception:
            continue
        st['fp'][n] = _sheet_fingerprint(sh)
        st['geom'][n] = _shapes_geometry(sh)
        # 循環参照のあるシートに Excel が描く矢印（Line 10 等）は、人が足した図形ではない。数えない
        # （2026-09-23 通しの実測 2: 撃つたびに「足された図形・グラフ: 売上!Line 10」と出ていた）
        try:
            if sh.CircularReference is not None:
                st['geom'][n] = [g for g in st['geom'][n] if not re.fullmatch(r'Line \d+', g['name'])]
        except Exception:
            pass
        try:
            for lo in sh.ListObjects:
                st['tables'].add((n, str(lo.Name)))
        except Exception:
            pass
        try:
            for pt in sh.PivotTables():
                st['pivots'].add((n, str(pt.Name)))
        except Exception:
            pass
    try:
        for nm in wb.Names:
            try:
                if nm.Visible:
                    st['names'][str(nm.Name)] = str(nm.RefersTo)
            except Exception:
                continue
    except Exception:
        pass
    try:
        for q in wb.Queries:
            st['queries'].add(str(q.Name))
    except Exception:
        pass
    return st


def _shelf_book_changes(before, after, target):
    """撃つ前後のブックの姿 → 作った物・書き換わった別シート・消えた物（純 Python・2026-09-23）。

    created は undo の「作った物」の形（kind/sheet/name）。控えは対象シートの中身しか戻さなかったので、
    棚のマクロが足したシート（調査_…）・グラフ・テーブル・ピボット・クエリ・名前は、undo の後も残っていた。
    足したシートの上の物はシートごと消えるので数えない。
    """
    added = [n for n in after['order'] if n not in before['order']]
    removed = [n for n in before['order'] if n not in after['order']]
    changed, recalc = [], []
    for n, fp in after['fp'].items():
        old = before['fp'].get(n)
        if n == target or n not in before['fp'] or fp == old:
            continue
        if fp and old and fp[:2] == old[:2] and fp[1] is not None:
            recalc.append(n)                     # 書いた中身は同じ＝式の再計算で値が変わっただけ
        else:
            changed.append(n)
    created = [{'kind': 'sheet', 'name': n} for n in added]
    for n, geom in after['geom'].items():
        if n in added:
            continue
        old = {g['name'] for g in before['geom'].get(n, [])}
        created += [{'kind': 'shape', 'sheet': n, 'name': g['name']} for g in geom if g['name'] not in old]
    created += [{'kind': 'table', 'sheet': n, 'name': t}
                for n, t in sorted(after['tables'] - before['tables']) if n not in added]
    created += [{'kind': 'pivot', 'sheet': n, 'name': p}
                for n, p in sorted(after['pivots'] - before['pivots']) if n not in added]
    created += [{'kind': 'query', 'name': q} for q in sorted(after['queries'] - before['queries'])]
    created += [{'kind': 'name', 'name': nm} for nm in after['names'] if nm not in before['names']]
    lost = [nm for nm in before['names'] if nm not in after['names']]
    moved = [nm for nm in before['names'] if nm in after['names'] and after['names'][nm] != before['names'][nm]]
    return {'created': created, 'changed': changed, 'recalc': recalc, 'removed': removed,
            'lost_names': lost, 'moved_names': moved}


_SHELF_KIND_LABEL = (('sheet', '足されたシート'), ('shape', '足された図形・グラフ'), ('table', '足されたテーブル'),
                     ('pivot', '足されたピボット'), ('query', '足されたクエリ'), ('name', '足された名前'))


def _shelf_book_lines(bc):
    """撃つ前後の違い（作った物・別シート・消えた物）を報告の行に（純 Python）。"""
    out = []
    for kind, label in _SHELF_KIND_LABEL:
        items = [c for c in bc['created'] if c['kind'] == kind]
        if items:
            out.append(f"{label}: " + "・".join((f"{c['sheet']}!" if c.get('sheet') else '') + c['name'] for c in items))
    if bc['changed']:
        out.append("書き換わった別のシート: " + "・".join(bc['changed']) + "（undo で一緒に戻ります）")
    if bc.get('recalc'):
        out.append("計算で値が変わった別のシート: " + "・".join(bc['recalc'])
                   + "（式はそのまま。このシートを参照する式の結果が変わった）")
    if bc['removed']:
        out.append("⚠ 消えたシート: " + "・".join(bc['removed']) + "（agent --undo では戻りません。控えのファイルから写してください）")
    if bc['lost_names']:
        out.append("⚠ 消えた名前: " + "・".join(bc['lost_names'][:10]) + "（agent --undo では戻りません）")
    if bc['moved_names']:
        out.append("⚠ 参照先が変わった名前: " + "・".join(bc['moved_names'][:10]) + "（agent --undo では戻りません）")
    return out


def _shelf_inplace(ws, before_snap, rows_changed, sel):
    """--inplace: 右に出た結果の表で元の表を置き換える → 報告の行（2026-09-23 直し）。

    直す前は ① 結果を式ごと元の場所へ写してから元の表を消した＝元を指していた式（=COUNTIF($A$2:$A$6,…)）が
    写した自分自身を数え、件数と合計が黙って狂った（実測: 東 2 件→1 件・合計 5→4）。② 結果が元の表より
    広いと、写した後の「結果を消す」が写した先の右端まで消した。今は、式を値にしてから（元の表が消えると
    式は成り立たない）、元の表を消し、切り取り（Cut）で動かす＝重なっても Excel が正しく動かす。
    """
    vals_b = before_snap.get('values') or []
    c0_b, r0_b = int(before_snap.get('col', 1)), int(before_snap.get('row', 1))
    max_c_before = c0_b + (len(vals_b[0]) if vals_b else 0) - 1
    max_r_before = r0_b + len(vals_b) - 1
    new_cells, has_formula = [], False
    for it in rows_changed:
        if it.get('before') or not it.get('after'):
            continue
        m = re.match(r'([A-Z]+)(\d+)$', str(it.get('addr') or ''))
        if not m:
            continue
        r, c = int(m.group(2)), _col_num_of(m.group(1))
        if c > max_c_before:
            new_cells.append((r, c))
            has_formula = has_formula or str(it.get('after_formula') or '').startswith('=')
    if not new_cells:
        return ["その場置き換え: 右に出た新しい表が見つからないので、置き換えていません"]
    r_min, r_max = min(r for r, _ in new_cells), max(r for r, _ in new_cells)
    c_min, c_max = min(c for _, c in new_cells), max(c for _, c in new_cells)
    dst_r, dst_c = r_min, c0_b
    if sel:
        try:
            sr = ws.Range(sel)
            dst_r, dst_c = int(sr.Row), int(sr.Column)
        except Exception:
            pass
    if dst_c >= c_min:
        return ["その場置き換え: 置き換える先が結果の表より右にあるので、置き換えていません"]
    rng_new = ws.Range(ws.Cells(r_min, c_min), ws.Cells(r_max, c_max))
    src_addr = str(rng_new.Address).replace('$', '')
    if has_formula:                              # 値の貼り付け＝書式も文字の数字もそのまま残る
        rng_new.Copy()
        rng_new.PasteSpecial(-4163)              # xlPasteValues
        ws.Application.CutCopyMode = False
    orig = ws.Range(ws.Cells(dst_r, dst_c), ws.Cells(max(r_max, max_r_before), c_min - 1))
    try:
        orig.UnMerge()
    except Exception:
        pass
    orig.Clear()
    rng_new.Cut(ws.Cells(dst_r, dst_c))
    final = ws.Range(ws.Cells(dst_r, dst_c), ws.Cells(dst_r + r_max - r_min, dst_c + c_max - c_min))
    try:
        final.Borders.LineStyle = 1
        final.Borders.Weight = 2
    except Exception:
        pass
    out = [f"その場置き換え: {src_addr} → {str(final.Address).replace('$', '')}（元の表を消して、結果の表を移した）"]
    if has_formula:
        out.append("  結果の式は値にしました（元の表が無くなると、元を指していた式が成り立たないため）")
    return out


_EXTRA_PAGE_KEYS = (('Zoom', '拡大縮小'), ('FitToPagesWide', '横のページ数'), ('FitToPagesTall', '縦のページ数'),
                    ('Orientation', '向き'), ('PrintTitleRows', '見出しの繰り返し'), ('PrintArea', '印刷範囲'))


def _sheet_extras_state(ws, rows_max=3000):
    """セルの値・書式の差分では見えない物の姿（印刷設定・行の高さと非表示・絞り込み）。読めなければ {}（2026-09-24）。"""
    st = {}
    try:
        p = ws.PageSetup
        st['page'] = {k: str(getattr(p, k)) for k, _l in _EXTRA_PAGE_KEYS}
    except Exception:
        pass
    try:
        last = min(int(ws.UsedRange.Row) + int(ws.UsedRange.Rows.Count), rows_max)
        st['rows'] = [(round(float(ws.Rows(r).RowHeight), 2), bool(ws.Rows(r).Hidden)) for r in range(1, last + 1)]
    except Exception:
        pass
    try:
        st['filter'] = (bool(ws.AutoFilterMode), bool(ws.FilterMode))
    except Exception:
        pass
    return st


def _sheet_extras_changes(before, after):
    """_sheet_extras_state の前後 → 報告の行（純 Python）。shelf-run が印刷設定・行の高さだけを変えたマクロで
    「何も変わりませんでした」と言っていた（縦横1ページに収めて印刷設定・行の高さをそろえる・2026-09-24）。"""
    out = []
    pb, pa = (before or {}).get('page') or {}, (after or {}).get('page') or {}
    diff = [f"{lab} {pb.get(k)}→{pa.get(k)}" for k, lab in _EXTRA_PAGE_KEYS if k in pb and k in pa and pb[k] != pa[k]]
    if diff:
        out.append("印刷設定の変化: " + "・".join(diff))
    rb, ra = (before or {}).get('rows') or [], (after or {}).get('rows') or []
    n = min(len(rb), len(ra))
    hs = [i + 1 for i in range(n) if rb[i][0] != ra[i][0]]
    hid = [i + 1 for i in range(n) if rb[i][1] != ra[i][1]]
    if hs:
        out.append(f"行の高さの変化: {len(hs)} 行（{hs[0]} 行目 {rb[hs[0] - 1][0]}→{ra[hs[0] - 1][0]} ほか）")
    if hid:
        out.append(f"表示・非表示の変わった行: {len(hid)} 行（{' '.join(str(x) for x in hid[:8])}{'…' if len(hid) > 8 else ''}）")
    fb, fa = (before or {}).get('filter'), (after or {}).get('filter')
    if fb and fa and fb != fa:
        out.append(f"絞り込み: {'あり' if fb[0] else 'なし'}→{'あり' if fa[0] else 'なし'}"
                   + ("（絞り込み中）" if fa[1] else ""))
    return out


def _shelf_nothing_hint(row):
    """何も変わらなかったときの一言（純 Python）。選ぶ列のあるマクロは、選び方が前提に合わないと黙って終わる。"""
    msg = "何も変わりませんでした。"
    if row and row.get('sel') not in (None, '', '-'):
        msg += (f"このマクロは列を選んでから撃つ形です（選ぶ列={row['sel']}）。列は 1 列ずつカンマで分けて渡してください"
                "（例 --select A1:A6,C1:C6。隣り合う列も A1:A6,B1:B6 と分ける）。")
    return (msg + "すでに同じ結果になっているとき（2 回目に撃ったとき等）のほか、選んだ範囲や表の形がマクロの前提に"
            "合わないと、何も言わずに終わるマクロがあります")


_SHELF_ERR_CODES = {str(c) for c in range(-2146826288, -2146826245)}
_SHELF_ERR_WORDS = ('#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!')


def _shelf_is_error_text(t):
    """差分の片側（_text_of の文字）がエラー値か。COM はエラー値を -2146826xxx の数で返す。"""
    t = str(t or '').strip()
    return t in _SHELF_ERR_CODES or t in _SHELF_ERR_WORDS


def _font_colors_really_differ(rng, limit=4000):
    """範囲の中に見た目の違う文字色が 2 つ以上あるか（自動＝黒と明示の黒は同じと数える）。大きい範囲は True のまま（数えない）。"""
    try:
        if int(rng.Rows.Count) * int(rng.Columns.Count) > limit:
            return True
        seen = set()
        for cell in rng.Cells:
            try:
                seen.add(int(cell.Font.Color))
            except Exception:
                return True
            if len(seen) > 1:
                return True
        return False
    except Exception:
        return True


def _book_with_sheet(xl, wb, sheet):
    """--sheet の当て先。wb にあれば (wb, '')。無ければ、そのシートを持つ見えているブックが 1 冊だけなら (そのブック, 知らせ)、
    0 冊・2 冊以上なら (None, エラーの文)。"""
    def has(b):
        try:
            b.Sheets(sheet)
            return True
        except Exception:
            return False
    if has(wb):
        return wb, ''
    others = []
    try:
        for i in range(1, int(xl.Workbooks.Count) + 1):
            b = xl.Workbooks.Item(i)
            try:
                if b.IsAddin or b.Name == wb.Name or not b.Windows(1).Visible:
                    continue
            except Exception:
                continue
            if has(b):
                others.append(b)
    except Exception:
        pass
    if len(others) == 1:
        return others[0], f"（シート「{sheet}」は前に出ている {wb.Name} ではなく {others[0].Name} にありました。{others[0].Name} を対象にします）"
    try:
        names = ', '.join(str(s.Name) for s in wb.Sheets)
    except Exception:
        names = '?'
    if others:
        return None, (f"エラー: シート「{sheet}」は前に出ている {wb.Name} に無く、ほかの "
                      f"{' / '.join(b.Name for b in others)} にあります。どれか 1 冊を前に出して撃ち直してください")
    return None, f"エラー: シート「{sheet}」が {wb.Name} にも、ほかに開いているブックにもありません（{wb.Name} のシート: {names}）"


def cmd_shelf_run(args):
    """`shelf-run 名前 [--select 範囲[,範囲]] [--sheet 名] [--input-text 値]`: 棚のマクロを選んで撃つ（2026-09-23）。

    対象ブック・シートを前に出し、--select があれば選んでから棚のマクロを撃つ。撃つ前に控えを取り、
    撃った後の差分（番地: 前→後・上限 30 件＋件数・書式の変化・足された行/列）と、ブック全体の違い
    （足されたシート・図形・グラフ・テーブル・ピボット・クエリ・名前／書き換わった別シート／消えた物）を返す。
    作った物と書き換わった別シートは控えの覚書に積む＝agent --undo でまとめて戻る。
    途中で出た窓（MsgBox/InputBox）は安全側（キャンセル優先）で自動解除し、本文を報告する
    （--input-text があれば InputBox に入れて確定）。明細は agent --changes でも読める。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: shelf-run <名前> [--select 範囲[,範囲]] [--sheet 名] [--input-text 値]（一覧は shelf）")
        return False
    name = rest[0]
    xl, wb = get_workbook(target_file)
    owner, lines = _shelf_source(xl)
    if not owner:
        print(f"エラー: 棚（モジュール「{_SEIRI_MODULE}」）が開いているブック・アドインに見つかりません")
        return False
    catalog = _parse_shelf_catalog(lines)
    row = next((r for r in catalog if r['name'] == name), None)
    if row is None:
        cands = [r['name'] for r in catalog if name in r['name']][:5]
        print(f"エラー: 棚に「{name}」はありません" + (f"（似た名前: {' / '.join(cands)}）" if cands else "")
              + "（一覧は shelf --grep・依頼文からなら shelf --ask）")
        return False
    sheet_opt = getattr(args, 'sheet_opt', None)
    if sheet_opt and not target_file:
        # 前に出ているブックにそのシートが無い＝前が入れ替わっている（2026-09-24 17:25: 秀コンボ.xlsm が前に出ていて
        # Book2 の「明細」に撃てず、COM の生の例外だけが返った）。そのシートを持つブックが 1 冊だけならそちらへ移る
        wb2, msg = _book_with_sheet(xl, wb, sheet_opt)
        if wb2 is None:
            print(msg)
            return False
        if msg:
            print(msg)
        wb = wb2
    try:
        wb.Activate()
        ws = wb.Sheets(sheet_opt) if sheet_opt else wb.ActiveSheet
        ws.Activate()
    except Exception as ex:
        print(f"エラー: 対象シートを前に出せません（{wb.Name}!{sheet_opt or '今のシート'}）: {ex}")
        return False
    sel = getattr(args, 'select', None)
    if sel:
        try:
            ws.Range(sel).Select()
        except Exception as ex:
            print(f"エラー: --select の範囲を選べません（{sel}）: {ex}")
            return False
    else:
        # --select が無いときは選択を今のセル 1 つに戻す。棚のマクロは「2 つ以上選んでいればその範囲だけ」を見るので、
        # 前に撃ったときの選択（選んだ列の文字の日付を日付にする --select B2:B13 など）が残ると、表の全部でなく B 列だけを見る（2026-09-23）
        try:
            xl.ActiveCell.Select()
        except Exception:
            pass
    sheet_name = str(ws.Name)
    if job_clock_elapsed(max_age=900) is None:
        # materials が押した時計が動いていれば押し直さない（押し直すと tidy の「materials から N 秒」が
        # shelf-run からの秒になり、頼んでから終わるまでより短く出た・2026-09-23 のテスト）
        job_clock_start(f"shelf-run {name}  {wb.Name}!{sheet_name}")
    from vbam_undo import (_agent_backup, _changes_of, _changes_table, _sheet_snapshot, _format_changes,
                           _format_table, _write_changes_file, _record_created, _undo_add_sheets)
    from vbam_ledger import runs_append, _run_id
    book_before = _shelf_book_state(wb)
    before_snap = _sheet_snapshot(ws)
    extras_before = _sheet_extras_state(ws)
    run_id = _run_id()
    request = f"shelf-run {name}" + (f" --select {sel}" if sel else "")
    backup_path = _agent_backup(wb, ws, sheet_name, request=request, run_id=run_id)
    # run-macro と同じハーネスで撃つ（run_book_macro）。素の Run だと実行時エラーが「終了／デバッグ」の窓になり、
    # 窓の見張りはそれを閉じないので、人が vbe-reset するまで Excel ごと止まる（2026-09-23 選んだ列の途切れた式を戻す で 671 秒）
    watcher = _start_dialog_watcher(xl, input_texts=getattr(args, 'input_text', None))
    t0 = time.time()
    ok = True
    from vbam_vba import run_book_macro, _VbaRuntimeError
    try:
        run_book_macro(xl, owner, name)
    except _VbaRuntimeError as ex:
        ok = False
        print(f"エラー: マクロが実行時エラーで止まりました: {ex}（窓は出ていません。落ちた行までの変更は下の差分に出ます）")
    except Exception as ex:
        ok = False
        print(f"エラー: マクロを撃てませんでした: {ex}")
    finally:
        try:
            watcher.stop()
        except Exception:
            pass
        # 撃った後も対象のブックとシートを前に残す（棚の持ち主のブックが前に出たまま戻ると、次の手が
        # 持ち主のブックに当たる・2026-09-24）
        try:
            wb.Activate()
            ws.Activate()
        except Exception:
            pass
    sec = time.time() - t0
    note = _dialog_watcher_note(watcher, None)
    if note:
        print(note)
    print(f"撃った: {owner}!{_shelf_module_name(xl, owner, name)}.{name}（{wb.Name}!{sheet_name}・{sec:.2f} 秒）"
          + (f"　選択: {sel}" if sel else ""))
    bc = _shelf_book_changes(book_before, _shelf_book_state(wb), sheet_name)
    if backup_path:                              # 控えの覚書に積む＝undo で作った物を消し、別シートも戻す
        _record_created(bc['created'])
        _undo_add_sheets(bc['changed'], book_before['geom'])
    rows_changed = _changes_of(before_snap, ws, sheet_name) if before_snap else None
    inplace_lines = []
    if getattr(args, 'inplace', False) and rows_changed:
        try:
            ws.Activate()
            inplace_lines = _shelf_inplace(ws, before_snap, rows_changed, sel)
        except Exception as ex_ip:
            inplace_lines = [f"警告: --inplace の置き換えで失敗しました: {ex_ip}（戻すなら agent --undo）"]
        rows_changed = _changes_of(before_snap, ws, sheet_name)
    n_changed = None
    if rows_changed is None:
        print("差分: 数えられませんでした（表が大きすぎる、等）。控えから確かめてください")
    else:
        n_changed = len(rows_changed)
        print(f"変わったセル: {n_changed} 個")
        if rows_changed:
            for line in _changes_table(rows_changed, show=30):
                print(line)
    fmt_rows = []
    if before_snap and not before_snap.get('too_big'):
        fmt_rows = _format_changes(before_snap.get('format'), ws)
        if fmt_rows:
            print(f"書式の変化: {len(fmt_rows)} 件")
            for line in _format_table(fmt_rows, show=12):
                print(line)
    grew = False
    try:
        ur = ws.UsedRange
        nr_a, nc_a = int(ur.Rows.Count), int(ur.Columns.Count)
        if before_snap and not before_snap.get('too_big') and before_snap.get('values'):
            nr_b = len(before_snap['values'])
            nc_b = len(before_snap['values'][0]) if before_snap['values'] else 0
            if nr_a > nr_b:
                print(f"足された行: {nr_a - nr_b} 行")
                grew = True
            if nc_a > nc_b:
                print(f"足された列: {nc_a - nc_b} 列")
                grew = True
    except Exception:
        pass
    extra_lines = _sheet_extras_changes(extras_before, _sheet_extras_state(ws))
    for line in inplace_lines + extra_lines + _shelf_book_lines(bc):
        print(line)
    # 撃った後に新しくできたエラー（#REF! 等）。行を消すマクロが、消えた行を指していた式を #REF! にする
    # （2026-09-23 通しの実測 2: 空行を消したら 1 行上を指していた H10 が =F9*#REF! になり、合計まで #REF! になった）
    try:
        new_err = []
        for rw in rows_changed or []:
            af, bf = str(rw.get('after_formula') or ''), str(rw.get('before_formula') or '')
            a_err = _shelf_is_error_text(rw.get('after'))
            b_err = _shelf_is_error_text(rw.get('before'))
            if ('#REF!' in af and '#REF!' not in bf) or (a_err and not b_err):
                new_err.append(rw['addr'])
        if new_err:
            print(f"⚠ 【要確認】撃った後にエラーが新しくできたセル: {len(new_err)} 個（{' '.join(new_err[:12])}"
                  + ("…" if len(new_err) > 12 else "") + "）→ 行や列を消すマクロなら、消した所を指していた式が壊れた。"
                  "agent --undo で戻し、先に式を直してから（選んだ列の途切れた式を戻す 等）撃ち直す")
    except Exception:
        pass
    # 撃った後の ###（合計行を足すで合計が列幅に収まらず ##### のまま渡った・2026-09-23 の通しの実測）
    try:
        ur = ws.UsedRange
        nr_h, nc_h = int(ur.Rows.Count), int(ur.Columns.Count)
        if nr_h * nc_h <= 3000:
            hashes, bad_h, _n = _hash_scan(ur, max_rows=nr_h, max_cols=nc_h)
            if hashes:
                print(f"⚠ 【表示崩れ】'###' で読めないセル: {hashes}個（{' '.join(bad_h)}・列幅不足）→ tidy 表の範囲 で列幅を直す")
    except Exception:
        pass
    if rows_changed is not None:
        _write_changes_file(rows_changed, fmt_rows=fmt_rows)
    if ok and not note and n_changed == 0 and not fmt_rows and not grew and not any(bc.values()) and not extra_lines:
        print(_shelf_nothing_hint(row))
    runs_append({'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'run_id': run_id, 'mode': 'shelf-run',
                 'book': str(wb.Name), 'sheet': sheet_name, 'request': request, 'macro': name,
                 'sec': round(sec, 2), 'changed': n_changed, 'created': len(bc['created']),
                 'other_sheets': bc['changed']})
    print(f"（控え: {backup_path or 'なし'}。戻すなら agent --undo・明細は agent --changes。保存はしていません）")
    return ok


# ----------------------------------------------------------------
# 式の元をたどる trace（2026-09-23）
#   DirectPrecedents はシートをまたがない（別シート参照は取れない）ので、式の字面を自分で読んで解く。
#   「C15 の合計は 明細!E4:E14 の和・E9 だけ文字の数字」のように、番地を添えて答えるための材料。
# ----------------------------------------------------------------
_TRACE_STR_RE = re.compile(r'"[^"]*"')
_TRACE_REF_RE = re.compile(
    r"(?:\[(?P<book>[^\]]+)\])?"
    r"(?:(?P<sheet>'(?:[^']|'')+'|[^\s!,()+\-*/&<>=^%:\"'\[\]{};]+)!)?"
    r"(?:(?P<a1>\$?[A-Z]{1,3}\$?\d{1,7})(?::(?P<a2>\$?[A-Z]{1,3}\$?\d{1,7}))?"
    r"|(?P<c1>\$?[A-Z]{1,3}):(?P<c2>\$?[A-Z]{1,3})(?![\d(])"
    r"|(?P<r1>\$?\d{1,7}):(?P<r2>\$?\d{1,7}))"
)
# 名前（税率 など）とテーブルの参照（売上[金額]・[@単価]）。関数名は後ろの ( で外す
_TRACE_WORD_RE = re.compile(r"(?<![\w.$\[])([^\W\d][\w.]*)(?![\w(!\[])")
_TRACE_TABLE_RE = re.compile(r"([^\W\d][\w.]*)?(\[(?:[^\[\]]|\[[^\]]*\])*\])")
_TRACE_ERRORS = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
                 -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!'}
_TRACE_MAX_NODES = 120          # 1 回の trace で読むセルの上限（式が横に広くても止まる）
_TRACE_MAX_REFS = 12            # 1 つの式から追う参照の数の上限
_TRACE_SUMMARY_MAX_CELLS = 20000  # 範囲の中身の内訳を数えるセル数の上限
_TRACE_NUMLIKE_RE = re.compile(r"^[-+△▲]?\d[\d,]*(?:\.\d+)?%?$")


def _trace_refs(formula, default_sheet=''):
    """式の字面 → 参照している所 [{'sheet','a1','a2','book','text','whole'}]（重複なし・純 Python）。

    文字列リテラルは先に落とす。関数名（LOG10・SUMIF 等）は前後の字で外す。
    列まるごと（A:C）・行まるごと（1:3）も拾う（whole='col'/'row'。VLOOKUP(…,マスタ!A:C,…) や
    SUMIF(明細!D:D,…) を黙って落としていた＝木に元が 1 つも出ず、値だけのセルに見えた・2026-09-23）。
    別ブック（[Book.xlsx] ／ 'C:\\…\\[Book.xlsx]Sheet'）は book に入れて返す（たどらない＝開いていないことがある）。
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    body = _TRACE_STR_RE.sub(' ', formula)
    out, seen = [], set()
    for m in _TRACE_REF_RE.finditer(body):
        kind = 'a1' if m.group('a1') else ('c1' if m.group('c1') else 'r1')
        i, j = m.start(kind), m.end()
        before = body[i - 1] if i else ''
        after = body[j] if j < len(body) else ''
        if m.group('sheet') is None and m.group('book') is None and before and (before.isalpha() or before in '_.'):
            continue                                   # LOG10 のような関数名の一部
        if after and (after == '(' or after.isdigit() or after.isalpha() or after in '_['):
            continue
        book = m.group('book') or ''
        sheet = (m.group('sheet') or '').strip("'").replace("''", "'")
        if '[' in sheet and ']' in sheet:              # 'C:\…\[Book.xlsx]Sheet1'!A1（閉じた別ブック）
            book = sheet[sheet.index('[') + 1:sheet.index(']')]
            sheet = sheet[sheet.index(']') + 1:]
        sheet = sheet or default_sheet
        if kind == 'a1':
            a1 = m.group('a1').replace('$', '')
            a2 = (m.group('a2') or '').replace('$', '')
            whole = ''
        else:
            a1 = m.group(kind).replace('$', '')
            a2 = m.group('c2' if kind == 'c1' else 'r2').replace('$', '')
            whole = 'col' if kind == 'c1' else 'row'
        key = (book, sheet, a1, a2)
        if key in seen:
            continue
        seen.add(key)
        out.append({'book': book, 'sheet': sheet, 'a1': a1, 'a2': a2, 'whole': whole,
                    'text': (f"{sheet}!" if sheet else '') + (f"{a1}:{a2}" if a2 else a1)})
    return out


def _trace_words(formula, names=(), tables=()):
    """式の字面 → (使っている名前, テーブルの参照)（純 Python）。names・tables は小文字で比べる。

    名前（=B2*税率 の 税率）は _trace_refs では拾えない＝黙って落ちていた（2026-09-23）。
    テーブルの参照（売上[金額]・[@単価]）は たどらない が、あることは言う（黙って落とさない）。
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return [], []
    body = _TRACE_STR_RE.sub(' ', formula)
    tabs = []
    for m in _TRACE_TABLE_RE.finditer(body):
        head = m.group(1) or ''
        if head and head.lower() not in tables:
            continue                                   # [Book.xlsx] の類はテーブルではない
        if not head and not m.group(2).startswith('[@') and not m.group(2).startswith('[['):
            continue
        tok = head + m.group(2)
        if tok not in tabs:
            tabs.append(tok)
    body = _TRACE_TABLE_RE.sub(' ', body)
    body = _TRACE_REF_RE.sub(' ', body)
    found = []
    for m in _TRACE_WORD_RE.finditer(body):
        w = m.group(1)
        if w.lower() in names and w not in found:
            found.append(w)
    return found, tabs


def _trace_range_cells(a1, a2):
    """範囲の (件数, 先頭, 末尾)。読めなければ (None, a1, a2)（純 Python）。"""
    from vbam_inv import _range_box
    box = _range_box(f"{a1}:{a2}")
    if not box:
        return None, a1, a2
    r1, c1, r2, c2 = box
    return (r2 - r1 + 1) * (c2 - c1 + 1), f"{_col_letter(c1)}{r1}", f"{_col_letter(c2)}{r2}"


def _trace_kind_of(v):
    """セルの値の種類（純 Python）: 'blank' 'num' 'tnum'（文字の数字） 'text' 'err' 'bool'。

    COM の Value は数を float、エラーを int（-2146826281 等）で返す。文字の数字は全角・カンマ・円記号・
    △（負）をならして数に見えるもの＝SUM・AVERAGE が黙って数えない値。
    """
    if v is None or (isinstance(v, str) and v.strip() == ''):
        return 'blank'
    if isinstance(v, bool):
        return 'bool'
    if isinstance(v, int):
        return 'err'
    if isinstance(v, str):
        t = unicodedata.normalize('NFKC', v).strip().replace('¥', '').replace('\\', '').replace('円', '')
        t = t.replace(' ', '')
        return 'tnum' if _TRACE_NUMLIKE_RE.match(t) else 'text'
    return 'num'


def _trace_summary_line(vals, fmls, texts, r0, c0, limit=5):
    """範囲の中身の内訳を 1 行に（純 Python）。「C15 の合計は E4:E14 の和・E9 だけ文字の数字」の材料。"""
    counts = {'num': 0, 'tnum': 0, 'text': 0, 'blank': 0, 'err': 0, 'bool': 0}
    tnum, errs, n_f = [], [], 0
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            k = _trace_kind_of(v)
            counts[k] += 1
            f = fmls[i][j] if i < len(fmls) and j < len(fmls[i]) else None
            if isinstance(f, str) and f.startswith('='):
                n_f += 1
            addr = f"{_col_letter(c0 + j)}{r0 + i}"
            shown = texts[i][j] if texts and i < len(texts) and j < len(texts[i]) else _TRACE_ERRORS.get(v, v)
            if k == 'tnum' and len(tnum) < limit:
                tnum.append(f'{addr} "{v}"')
            elif k == 'err' and len(errs) < limit:
                errs.append(f"{addr} {shown}")
    parts = [f"数 {counts['num']}"]
    if counts['tnum']:
        parts.append(f"⚠ 文字の数字 {counts['tnum']}（{'・'.join(tnum)}）")
    if counts['text']:
        parts.append(f"文字 {counts['text']}")
    if counts['bool']:
        parts.append(f"TRUE/FALSE {counts['bool']}")
    parts.append(f"空 {counts['blank']}")
    if counts['err']:
        parts.append(f"⚠ エラー {counts['err']}（{'・'.join(errs)}）")
    line = "内訳: " + "・".join(parts) + f"（うち式 {n_f}）"
    if counts['tnum']:
        line += "　※ SUM・AVERAGE・COUNT は文字の数字を数えない"
    return line


def _trace_range(wb, r):
    """範囲の参照 → (見出しの行, 中身の内訳の行 or None, 先頭, 末尾)。列・行まるごとは使っている所に絞る。"""
    sheet = r['sheet']
    spec = f"{r['a1']}:{r['a2']}"
    head_line = f"{sheet}!{spec}"
    ws = wb.Sheets(sheet) if sheet else wb.ActiveSheet
    rng = ws.Range(spec)
    if r.get('whole'):
        eff = ws.Application.Intersect(rng, ws.UsedRange)
        if eff is None:
            return head_line + f"  {'列' if r['whole'] == 'col' else '行'}まるごと（使っている所に掛からない＝空）", None, None, None
        rng = eff
    n = int(rng.Cells.CountLarge)
    first = str(rng.Cells(1, 1).Address).replace('$', '')
    last = str(rng.Cells(int(rng.Rows.Count), int(rng.Columns.Count)).Address).replace('$', '')
    if r.get('whole'):
        head_line += (f"  {'列' if r['whole'] == 'col' else '行'}まるごと＝使っている所 {first}:{last}"
                      f"  {n:,} 件（先頭 {first} / 末尾 {last}）")
    else:
        head_line += f"  {n} 件（先頭 {first} / 末尾 {last}）"
    if n > _TRACE_SUMMARY_MAX_CELLS:
        return head_line, "内訳: 大きいので数えません", first, last
    vals, fmls = _rows_of_value(rng.Value), _rows_of_value(rng.Formula)
    try:
        texts = _rows_of_value(rng.Text) if n == 1 else None
    except Exception:
        texts = None
    return head_line, _trace_summary_line(vals, fmls, texts, int(rng.Row), int(rng.Column)), first, last


def _rows_of_value(v):
    """Range.Value（タプルのタプル／1 セルは素の値）を 2 次元リストに（純 Python）。"""
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) if isinstance(r, tuple) else [r] for r in v]


def _trace_read(wb, sheet, a1):
    """1 セルの (式 or None, 表示文字)。読めなければ (None, '?')。"""
    try:
        ws = wb.Sheets(sheet) if sheet else wb.ActiveSheet
        rng = ws.Range(a1)
        f = str(rng.Formula or '')
        return (f if f.startswith('=') else None), str(rng.Text or '')
    except Exception:
        return None, '?'


def _trace_walk(wb, sheet, a1, depth, path, state, indent=1):
    """1 セルの元をたどって state['lines'] に木を積む（深さ・件数・循環で止まる）。

    state['names']（小文字の名前 → 参照先の字面）・state['tables']（小文字のテーブル名）があれば、
    名前は参照先へたどり、テーブルの参照はあることだけ言う。
    """
    pad = '  ' * indent
    key = f"{sheet}!{a1}"
    if key in path:
        state['lines'].append(f"{pad}{key}  ⚠ 循環参照（ここで止めます）")
        return
    if state['nodes'] >= _TRACE_MAX_NODES:
        return
    state['nodes'] += 1
    formula, text = _trace_read(wb, sheet, a1)
    state['lines'].append(f"{pad}{key}  {formula or '（値）'}  → {text}")
    if formula is None or depth <= 0:
        return
    refs = _trace_refs(formula, sheet)
    names = state.get('names') or {}
    used_names, tabs = _trace_words(formula, names, state.get('tables') or ())
    for nm in used_names:
        target = names.get(nm.lower(), '')
        state['lines'].append(f"{pad}  名前 {nm} = {target.lstrip('=')}")
        refs += [dict(r, via=nm) for r in _trace_refs(target, sheet)]
    for t in tabs:
        state['lines'].append(f"{pad}  テーブルの参照 {t}（たどりません。materials のテーブル欄で列を確かめる）")
    if len(refs) > _TRACE_MAX_REFS:
        state['lines'].append(f"{pad}  …参照 {len(refs)} か所のうち先頭 {_TRACE_MAX_REFS} か所だけ追います")
        refs = refs[:_TRACE_MAX_REFS]
    for r in refs:
        if r['book']:
            state['lines'].append(f"{pad}  [{r['book']}]{r['text']}  （別ブック＝たどりません）")
            continue
        if r['a2']:
            try:
                head_line, summary, head, tail = _trace_range(wb, r)
            except Exception:
                n, head, tail = _trace_range_cells(r['a1'], r['a2']) if not r.get('whole') else (None, None, None)
                head_line, summary = (f"{r['sheet']}!{r['a1']}:{r['a2']}"
                                      + (f"  {n} 件（先頭 {head} / 末尾 {tail}）" if n else '')), None
            state['lines'].append(f"{pad}  {head_line}")
            if summary:
                state['lines'].append(f"{pad}    {summary}")
            for one in ([head, tail] if head != tail else [head]):
                if one:
                    _trace_walk(wb, r['sheet'], one, depth - 1, path | {key}, state, indent + 2)
            continue
        _trace_walk(wb, r['sheet'], r['a1'], depth - 1, path | {key}, state, indent + 1)


def _trace_book_words(wb):
    """ブックの名前（小文字 → 参照先の字面）とテーブル名（小文字）の集まり。読めなければ空。"""
    names, tables = {}, set()
    try:
        for nm in wb.Names:
            try:
                full = str(nm.Name)
                names.setdefault(full.split('!')[-1].lower(), str(nm.RefersTo))
            except Exception:
                continue
    except Exception:
        pass
    try:
        for sh in wb.Worksheets:
            try:
                for lo in sh.ListObjects:
                    tables.add(str(lo.Name).lower())
            except Exception:
                continue
    except Exception:
        pass
    return names, tables


def cmd_trace(args):
    """`trace <セル> [--depth N] [--sheet 名]`: 式の元をシートをまたいでたどる（2026-09-23）。

    「番地・式・値」の木で返す。範囲は先頭と末尾と件数に畳み、中身の内訳（数・文字の数字・空・エラーの番地）を
    添える。列まるごと（A:C）は使っている所に絞る。名前は参照先へたどる。循環は印を付けて止める。
    DirectPrecedents はシートをまたがないので、式の字面を読んで解いている。何も書き換えない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: trace <セル> [--depth N] [--sheet 名]（例: trace 集計!C15 --depth 3）")
        return False
    spec = rest[0]
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    if '!' in spec and not sheet_opt:
        sheet, a1 = spec.rsplit('!', 1)
        sheet = sheet.strip("'").replace("''", "'")
    else:
        a1 = spec
        sheet = sheet_opt or str(wb.ActiveSheet.Name)
    a1 = a1.replace('$', '').upper()
    try:
        depth = int(getattr(args, 'depth', None) or 3)
    except (TypeError, ValueError):
        depth = 3
    try:
        wb.Sheets(sheet)
    except Exception:
        print(f"エラー: シート「{sheet}」がありません（シートの一覧は sheet-info）")
        return False
    print(f"式の元: {wb.Name}   深さ {depth} まで")
    names, tables = _trace_book_words(wb)
    state = {'lines': [], 'nodes': 0, 'names': names, 'tables': tables}
    _trace_walk(wb, sheet, a1, depth, frozenset(), state, indent=0)
    for line in state['lines']:
        print(line)
    if state['nodes'] >= _TRACE_MAX_NODES:
        print(f"（読んだセルが上限 {_TRACE_MAX_NODES} 個に達したので途中で止めました）")
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
    # --sheet はほかの手（read-range・tidy・materials）と同じ口（2026-09-24: screenshot --sheet 図形 が不明な引数で 2 回落ちた）
    ws, rng = _resolve_range(xl, wb, spec, getattr(args, 'sheet_opt', None))

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
        if fc is None and _font_colors_really_differ(ur):
            # 「自動」と「黒を明示」だけの違いでも範囲の Font.Color は None になる＝見た目は同じなのに
            # 「混在あり」と言っていた（2026-09-24 通しの確かめ: 表を整えるマクロが黒を明示した後）
            marks.append("文字色の混在あり")
        fb = ur.Font.Bold
        if fb is None:
            marks.append("太字の混在あり")
        if marks:
            out.append("書式: " + "・".join(marks) + "（番地は style-map で）")
    except Exception:
        pass
    return out


def _sum_gap_and_double_notes(fa, vals, r0, c0):
    """縦の SUM の 2 つの間違い → 気づきの行（純 Python・2026-09-24 通しの実測 4）。

    ① 集計の取りこぼし: 範囲のすぐ下（上）の明細の数を外している（=SUM(G4:G9) の G10）。すぐ下・上が集計の式・
       集計の行なら取りこぼしではない（小計ごとの表）。
    ② 二重計上: 範囲の中に小計（SUM の式）と明細の数が両方入っている（=SUM(G4:G20) が小計 G11・G19 も足す）。
    どちらも式の目（formula_notes）と棚の手（seiri_formula_hints）が拾えず、黙って通していた。
    """
    grid_f, grid_v = fa or [], vals or []

    def cell(g, r, c):
        i, j = r - r0, c - c0
        if 0 <= i < len(g) and 0 <= j < len(g[i] or []):
            return g[i][j]
        return None

    def is_num(v):
        return isinstance(v, (int, float)) and not isinstance(v, bool)

    def is_agg(r, c):
        f = cell(grid_f, r, c)
        return isinstance(f, str) and bool(_FML_AGG_ANY_RE.search(f))

    def total_row(r):
        i = r - r0
        return 0 <= i < len(grid_v) and any(_is_total_label(v) for v in (grid_v[i] or []))

    out = []
    for i, row in enumerate(grid_f):
        for j, f in enumerate(row or []):
            if not (isinstance(f, str) and f.startswith('=')):
                continue
            aggs = fml_agg_ranges(f)
            if len(aggs) != 1 or aggs[0][0] != 'SUM' or len(aggs[0][1]) != 1:
                continue
            top, left, bot, right = aggs[0][1][0]
            r, c = r0 + i, c0 + j
            if left != right or left != c or not (bot < r or top > r):
                continue
            addr = f"{_col_letter(c)}{r}"
            inner = range(top, bot + 1)
            subs = [k for k in inner if is_agg(k, c)]
            plain = [k for k in inner if not is_agg(k, c) and is_num(cell(grid_v, k, c))]
            if subs and plain:
                refs = "+".join(f"{_col_letter(c)}{k}" for k in subs)
                out.append(f"集計の二重計上 {addr}: {_clip_fml(f)} は小計 {' '.join(f'{_col_letter(c)}{k}' for k in subs[:4])} と"
                           f"その明細を両方足しています  → 小計どうしを足す（={refs}）か、小計も SUBTOTAL にする")
                continue
            for k, side in ((bot + 1, "直下"), (top - 1, "直上")):
                if k == r or k < r0 or is_agg(k, c) or total_row(k):
                    continue
                v = cell(grid_v, k, c)
                if is_num(v) and plain:
                    nb = (top, k) if side == "直下" else (k, bot)
                    out.append(f"集計の取りこぼし {addr}: {_clip_fml(f)} の{side} {_col_letter(c)}{k}（{v:,.10g}）が入っていません"
                               f"  → =SUM({_col_letter(c)}{nb[0]}:{_col_letter(c)}{nb[1]})")
    return out


def diagnose_notes(fa, fr, vals, r0, c0, limit=8):
    """診断の目のうち、数式の目（formula_notes）と表の汚れ（dirt_notes）に無い 2 つだけを 1 手目に足す。

    循環参照＝要修正（done を止める側）・外れ値＝気づき。集計漏れ・直書き・形の違い・空白・文字の数字は
    既に materials の別の段が出しているので、ここでは数えない（同じ物を 2 度言わない・2026-09-17）。
    戻り値 (要修正の行, 気づきの行)。vbam_audit が無ければ ([], [])。
    """
    try:
        from vbam_audit import check_circular_references, check_data_cleaner, check_values_in_formula_columns
    except Exception:
        return [], []
    blockers, notices = [], []
    # 式の列に紛れた数字の直書き（formula_notes は式どうししか比べないので、ここで足す・2026-09-23）
    try:
        for it in check_values_in_formula_columns(fa, fr, vals, r0 - 1, c0 - 1):
            notices.append(f"式の列 {it['cell']}: {it['msg']}  → 上下と同じ式に戻す（棚なら 選んだ列の途切れた式を戻す）")
    except Exception:
        pass
    try:
        for it in check_circular_references(fa, r0 - 1, c0 - 1):
            blockers.append(f"{it['cell']}: {it['msg']}")
    except Exception:
        pass
    notices.extend(_sum_gap_and_double_notes(fa, vals, r0, c0))
    # 外れ値は手で入れた値だけ言う。式のセル（金額＝数量×単価・税込）は元の値の外れを重ねて言うだけ
    # （2026-09-23 通しの実測 2: 単価 G8 の桁違いに続けて H8・I8 も並べていた）
    fml = set()
    for i, row in enumerate(fa or []):
        for j, v in enumerate(row or []):
            if isinstance(v, str) and v.startswith('='):
                fml.add(f"{_col_letter(c0 + j)}{r0 + i}")
    try:
        for it in check_data_cleaner(vals, r0 - 1, c0 - 1):
            if it.get('type') == 'outlier_value' and it['cell'] not in fml:
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
    'cmd_shelf',
    'cmd_shelf_run',
    'cmd_trace',
    '_shelf_source',
    '_parse_shelf_catalog',
    '_trace_refs',
    '_trace_range_cells',
    '_trace_walk',
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
