# -*- coding: utf-8 -*-
"""vbam_edit.py — vba_manager 分割パート: 「手」コマンド（write/clear/format/sheet/table/検索置換/保存印刷/仕上げ）

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
from vbam_view import *  # noqa: F401,F403
# ================================================================
# 「手」コマンド (シートの編集・整形・構造操作)
#   ※ アクティブ(開いたまま)のブックに COM で直接書き込む。
#      Excel MCP と違いブックを閉じる必要がない代わりに、
#      プログラム経由の変更は Excel の Undo 履歴を消す。
#   ※ 既定では保存しない。Excelで確認後に手動保存するか、
#      保存せず閉じれば変更を破棄できる（=Undo代わりの逃げ道）。
# ================================================================

_LAST_QUERY_FILE  = os.path.join(SCRIPT_DIR, '_last_query.m')       # powerquery add のM式入力
_LAST_DAX_FILE    = os.path.join(SCRIPT_DIR, '_last_dax.dax')       # datamodel measure add のDAX入力

# 配置・罫線の定数 (xl定数の実値)
_XL_ALIGN_H = {'left': -4131, 'center': -4108, 'right': -4152,
               'fill': 5, 'justify': -4130}
_XL_ALIGN_V = {'top': -4160, 'center': -4108, 'bottom': -4107}
_XL_BORDER_WEIGHT = {'hairline': 1, 'thin': 2, 'medium': -4138, 'thick': 4}


def _unfilter_for_write(ws):
    """書く前に絞り込み（オートフィルタ）を解く。解いたら True。

    絞り込みで行が隠れているシートに範囲で書くと、隠れた行が飛ばされて値が下にずれる。
    ClearContents も隠れた行を消さない（2026-09-04・本物のブック「文書件名簿」の 2024 シートで実測。
    行 2,3,6-9 が隠れたまま 303 行を書いたら 2 行ずれて 281 セルが #N/A になった）。
    絞り込みは「見せ方」であってデータではないので、書く側が解いてよい。ただし黙って解かず報告に出す。
    """
    try:
        if not bool(ws.FilterMode):
            return False
        ws.ShowAllData()
        print(f"※ シート '{ws.Name}' の絞り込みを解いてから書きます（隠れた行があると書き込みがずれるため）")
        return True
    except Exception:
        return False


def _hex_to_excel_color(hexstr):
    """#RRGGBB / RRGGBB → Excel の BGR 整数"""
    s = hexstr.lstrip('#').strip()
    if len(s) != 6:
        raise ValueError(f"色は #RRGGBB 形式で指定してください: {hexstr}")
    r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
    return r + g * 256 + b * 65536




class _alerts_off:
    """DisplayAlerts を一時的に切り、抜けるとき「元の値」に戻す with ブロック。

    無条件に True へ戻すと、batch や MCP の 1接続セッションで呼び出し元が意図的に
    False にしていた設定まで書き換えてしまい、後続の操作で確認ダイアログが出て
    無言ハングする（xl.Run はダイアログが閉じるまで戻らない）。変更前の値を覚えて復元する。
    """

    def __init__(self, xl):
        self.xl = xl
        self.prev = True

    def __enter__(self):
        try:
            self.prev = self.xl.DisplayAlerts
        except Exception:
            self.prev = True          # 読めない場合のみ Excel 既定（True）を採用
        self.xl.DisplayAlerts = False
        return self.xl

    def __exit__(self, exc_type, exc, tb):
        try:
            self.xl.DisplayAlerts = self.prev
        except Exception as ex:
            # 戻せないと DisplayAlerts=False のままの Excel が残る＝以後の上書き確認・
            # シート削除確認が全部抑止された状態でユーザーが使い続ける（開いたまま運用では実害）。
            # 保護の再適用失敗と同じく、黙らずに報告する
            print(f"⚠ DisplayAlerts を元に戻せませんでした（{self.prev} に戻す予定でした）: {ex}",
                  file=sys.stderr)
            print("  この Excel では確認ダイアログが抑止されたままです。"
                  "Excel を開き直すか、手動で設定を確認してください。", file=sys.stderr)
        return False


def _read_tsv_grid(path, raw=False):
    """TSV(タブ区切り)をセル値の 2次元タプルに変換（行の長さは不揃いのまま返す）。

    以前は短い行を None で最大列数まで詰めていたが、None 代入は既存セルの
    クリアとして作用し「触れないつもりの右側セル」を消すため、詰め物はしない。
    矩形化の要否は書き込み側（cmd_write_range）が判断する。
    """
    with open(path, 'r', encoding='utf-8-sig') as f:
        text = f.read()
    text = text.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
    if text == '':
        return ()
    return tuple(tuple((c if raw else _coerce_cell(c)) for c in line.split('\t'))
                 for line in text.split('\n'))


def _kept_as_text(cell, value):
    """文字列を書いた直後に Excel が日付・数値に読み替えていたら、表示形式 @ で文字列に書き直す。

    Excel は .Value 代入でも入力欄と同じ解釈をする（"1-2"→1月2日、"1,000"→1000、"(1)"→-1）。
    郵便番号やコード番号がこれで化ける。--raw を付けるかどうかを人（AI）に考えさせない＝道具が守る
    （2026-09-02 深夜・--raw の要否で止まった 129 秒の記録）。
    数式（'=' 始まり）と、_coerce_cell が数値・日付に変換したもの（意図した変換）は対象外。
    真偽値（"TRUE"→True）は従来どおり Excel に任せる。戻したら True を返す。
    """
    if not isinstance(value, str) or value == '' or value.startswith('='):
        return False
    try:
        v = cell.Value
    except Exception:
        return False
    if v is None or isinstance(v, (str, bool)):
        return False
    cell.NumberFormat = "@"
    cell.Value = value
    return True


def _keep_grid_text(ws, row0, col0, grid, skip_cols=()):
    """grid（2次元）を row0/col0 起点に書いた直後に呼ぶ _kept_as_text の範囲版。

    読み戻しは範囲 1 回（セル数ぶん COM 往復を増やさない）。戻したセルの番地一覧を返す。
    skip_cols は「文字に戻さない列」の 0 起点の並び（2026-09-09）。日付・数値と型を決めた列は
    Excel の読み替えが正しく、戻すと日付が文字になってピボットの月別まとめ（group_date）ができなくなる。
    """
    skip = set(skip_cols or ())
    nrows = len(grid)
    ncols = max((len(r) for r in grid), default=0)
    if nrows == 0 or ncols == 0:
        return []
    want = [(i, j, g) for i, row in enumerate(grid) for j, g in enumerate(row)
            if isinstance(g, str) and g != '' and not g.startswith('=') and j not in skip]
    if not want:
        return []
    rng = ws.Range(ws.Cells(row0, col0), ws.Cells(row0 + nrows - 1, col0 + ncols - 1))
    rb = rng.Value
    if nrows == 1 and ncols == 1:
        rb = ((rb,),)
    fixed = []
    for i, j, g in want:
        try:
            v = rb[i][j]
        except Exception:
            continue
        if v is None or isinstance(v, (str, bool)):
            continue
        c = ws.Cells(row0 + i, col0 + j)
        c.NumberFormat = "@"
        c.Value = g
        fixed.append(c.Address.replace('$', ''))
    return fixed


_KEPT_NOTE = "文字列として保持（Excel が日付・数値に読み替えたので文字列で書き直した）: "

# 動的配列の関数。.Value（＝.Formula）で書くと暗黙の交差 @ が付いて 1 セルに潰れる（FILTER が 1 件しか出ない・
# 2026-09-04 の手順書の実射で発覚）。これらを含む数式は Formula2 で書く
_DYN_FUNC_RE = re.compile(r'\b(?:FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|RANDARRAY|XLOOKUP|XMATCH|LET|LAMBDA|TEXTSPLIT|'
                          r'VSTACK|HSTACK|TOCOL|TOROW|BYROW|BYCOL|MAP|REDUCE|SCAN|MAKEARRAY|TAKE|DROP|CHOOSEROWS|'
                          r'CHOOSECOLS|WRAPROWS|WRAPCOLS|EXPAND|GROUPBY|PIVOTBY)\s*\(', re.I)


def _is_dynamic_formula(v):
    return isinstance(v, str) and v.startswith('=') and bool(_DYN_FUNC_RE.search(v))


def _write_value(cell, v):
    """1 セル（か範囲）に値・数式を書く。動的配列の関数を含む数式だけ Formula2（無い Excel では .Value に落ちる）。"""
    if _is_dynamic_formula(v):
        try:
            cell.Formula2 = v
            return
        except Exception:
            pass
    cell.Value = v


def _untext_one(cell, v):
    """数値・日付を書く先が文字列書式（@）なら標準に戻す。

    @ のまま datetime を入れると '1/4/2026 3:00:00 PM' という文字で入る（CSV 取り込み後の列は @ が多い・
    2026-09-04 の手順書の実射で発覚）。文字列を書くときは触らない（先頭ゼロの列の @ を守る）。
    """
    # 数式（'=' 始まり）も @ のままだと文字として入る（2026-09-04）。それ以外の文字列は触らない
    # （先頭ゼロの列の @ を守る）
    if v is None or (isinstance(v, str) and not v.startswith('=')):
        return
    try:
        if cell.NumberFormat == "@":
            # 日本語 Excel は NumberFormat="General" を「設定できません」で拒む（標準は 'G/標準'・2026-09-04 実測）。
            # 触っていない末尾セルの書式＝そのシートの「標準」をそのまま写す（言語に依らない）
            for fmt in (_general_format(cell), 'General', 'G/標準'):
                try:
                    cell.NumberFormat = fmt
                    break
                except Exception:
                    continue
    except Exception:
        pass


def _general_format(cell):
    """このセルのシートで「標準」を表す書式文字列（触っていない末尾セルから読む）。読めなければ 'General'。"""
    try:
        ws = cell.Worksheet
        return str(ws.Cells(ws.Rows.Count, ws.Columns.Count).NumberFormat)
    except Exception:
        return 'General'


def _untext_grid(ws, target, row0, col0, grid):
    """_untext_one の範囲版。範囲全体が @ でない書式でそろっていれば何もしない（不揃い＝None は見に行く）。"""
    try:
        fmt = target.NumberFormat
    except Exception:
        fmt = None
    if fmt is not None and fmt != "@":
        return
    for i, row in enumerate(grid):
        for j, g in enumerate(row):
            # 数式（'=' 始まり）も @ のままだと文字として入る。_untext_one は数式を見るのに、
            # ここが「文字列は全部素通り」だったので、範囲で書くときだけ式が文字のまま残っていた
            # （2026-09-09 実射「コード対応表で振り直し」＝番号列の隣に挿した列が @ を継ぎ、
            #   VLOOKUP が 4 行とも文字で入った。仕上げ検査も採点も素通り＝黙殺 1 件）
            if g is not None and (not isinstance(g, str) or g.startswith('=')):
                _untext_one(ws.Cells(row0 + i, col0 + j), g)


def _fix_dynamic_formulas(ws, row0, col0, grid):
    """grid をまとめて .Value で書いた後、動的配列の数式だけ Formula2 で書き直す（_write_value の範囲版）。"""
    for i, row in enumerate(grid):
        for j, g in enumerate(row):
            if _is_dynamic_formula(g):
                try:
                    ws.Cells(row0 + i, col0 + j).Formula2 = g
                except Exception:
                    pass


@protect_safe
def _count_existing_cells(rng):
    """書込先範囲の非空セル数（上書き警告用）。数えられなければ 0 を返す。

    Claude for Excel の allow_overwrite（既定で非空セルへの書込を止める）から輸入
    （2026-08-28 調査）。こちらは止めずに「何セル上書きするか」を報告だけする
    （既定の動きを変えない＝情報の追加のみ）。
    """
    try:
        return int(rng.Application.WorksheetFunction.CountA(rng))
    except Exception:
        return 0


def _report_write_result(ws, rng):
    """書込後の検証読み戻し：書いた範囲のエラーセル（#REF!等）や ### 表示を数えて報告する。

    Claude for Excel の formula_results（書いた数式の評価結果を返りに同梱）から輸入。
    書き込み直後にエラーや表示崩れを検知し、AIへ即座に自己修正を促す。
    """
    n_err = 0
    err_samples = []
    for typ in (-4123, 2):          # xlCellTypeFormulas / xlCellTypeConstants
        try:
            sp = rng.SpecialCells(typ, 16)   # 16 = xlErrors
            for c in sp:
                n_err += 1
                if len(err_samples) < 5:
                    try:
                        addr = c.Address.replace('$', '')
                        txt = str(c.Text or '#ERROR!')
                        err_samples.append(f"{addr}={txt}")
                    except Exception:
                        pass
        except Exception:
            pass                     # 該当なしは SpecialCells が例外を投げる＝正常

    # 列幅不足による ### 表示崩れの検知
    n_hash = 0
    hash_samples = []
    try:
        cell_count = int(rng.Cells.CountLarge)
        if cell_count <= 300:
            for c in rng.Cells:
                try:
                    txt = str(c.Text or '')
                    if txt and len(txt) >= 2 and all(ch == '#' for ch in txt):
                        n_hash += 1
                        if len(hash_samples) < 5:
                            addr = c.Address.replace('$', '')
                            hash_samples.append(f"{addr}(幅不足)")
                except Exception:
                    pass
    except Exception:
        pass

    if n_err > 0:
        tail = "（" + " ".join(err_samples) + ("…" if n_err > len(err_samples) else "") + "）"
        print(f"⚠ 【要修正】書き込み結果に数式エラーがあります: {n_err}件 {tail}")
    if n_hash > 0:
        tail = "（" + " ".join(hash_samples) + ("…" if n_hash > len(hash_samples) else "") + "）"
        print(f"⚠ 【表示崩れ】列幅不足により '###' 表示になっているセルがあります: {n_hash}件 {tail}")
        print("  対処: 列幅を広げるか、vba(\"tidy 範囲\") で仕上げてください。")


def cmd_write_range(args):
    """セル範囲に値・数式を書き込む (read-range の対)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: write-range [excel_file] <range> [値]")
        print("  単一値はインライン、グリッドは --tsv <file> か _last_values.tsv から読み込み")
        print("  '='始まりは数式として書き込み")
        return False
    spec = rest[0]
    inline_value = rest[1] if len(rest) >= 2 else None
    tsv_opt = getattr(args, 'tsv_opt', None)
    if _reject_extra_args(rest, 2, '使い方: write-range [excel_file] <range> [値]'):
        return False

    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    append = getattr(args, 'append', False)
    if append:
        # --append: spec は「シート名!列文字」または「列文字」。使用範囲の最終行の
        # 次の行から書く（自動の最終行判定を使うため、書き込み先番地を必ず明示表示する）
        # シート名は spec の「!」指定と --sheet のどちらでも指定できる（非append経路と
        # 揃える。--sheet を無視するとアクティブシートに書いて別シートを汚す）
        sheet_part = None
        col_part = spec
        if '!' in spec:
            sheet_part, col_part = spec.split('!', 1)
            # 照合はクォートを剥がしてから（"'ログ'!A" と --sheet ログ は同じ指定。
            # 剥がす前に比べると、同じシートなのに「食い違い」と誤エラーになる）
            if sheet_opt and sheet_opt.strip("'") != sheet_part.strip("'"):
                print(f"エラー: シート指定が食い違っています（'{spec}' と --sheet {sheet_opt}）")
                return False
        elif sheet_opt:
            sheet_part = sheet_opt
        if sheet_part is not None:
            sheet_part = sheet_part.strip("'")
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == sheet_part:
                    ws = sh
                    break
            if ws is None:
                print(f"エラー: シート '{sheet_part}' が見つかりません")
                return False
        else:
            ws = wb.ActiveSheet
        if not re.fullmatch(r'[A-Za-z]{1,3}', col_part or ''):
            print("エラー: --append の range は「シート名!列文字」（例: ログ!A）で指定してください")
            return False
        ur = ws.UsedRange
        next_row = ur.Row + ur.Rows.Count if ur is not None else 1
        # 空シートでも UsedRange は $A$1 を返すため、そのままだと2行目から始まる
        try:
            if (ur is not None and ur.Rows.Count == 1 and ur.Columns.Count == 1
                    and ur.Row == 1 and ur.Column == 1
                    and ur.Cells(1, 1).Value is None):
                next_row = 1
        except Exception:
            pass
        rng = ws.Range(f"{col_part.upper()}{next_row}")
        print(f"追記位置: {ws.Name}!{rng.Address}（使用範囲の最終行の次）")
    else:
        whole = _whole_sheet_spec(wb, spec, sheet_opt)
        if whole is not None:
            print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  全セルが同じ値で上書きされる危険があるため、write-range では")
            print(f"  セル/範囲を明示してください（例: \"{whole}!A1\"）")
            return False
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
        _unfilter_for_write(ws)

    raw = getattr(args, 'raw', False)

    # 書き込みは Worksheet_Change 等のイベントマクロを同期発火させる。そのマクロが
    # MsgBox を出すと rng.Value 代入がそこでブロックし、閉じるまでコマンドが無言で
    # ハングする（2026-07-11 実害）。安全解除の監視を書き込みの間だけ常設する。
    _dlg_watcher = _start_dialog_watcher(xl)
    written_rng = None               # 書込後検証（エラーセル報告）の対象
    kept = []                        # Excel の読み替えから文字列に戻したセル（報告用）
    try:
        if inline_value is not None:
            # インライン単一値: 範囲全体に同じ値 (数式可)
            _n既存 = _count_existing_cells(rng)
            if _n既存 > 0:
                print(f"上書き: 書込先に既存データ {_n既存} セルがあります")
            if raw:
                rng.NumberFormat = "@"   # 文字列として保持（"007" 等の先頭ゼロを守る）
                rng.Value = inline_value
            else:
                cv = _coerce_cell(inline_value)
                _untext_one(rng, cv)
                _write_value(rng, cv)
                if _kept_as_text(rng.Cells(1, 1), cv):
                    if int(rng.Cells.CountLarge) > 1:
                        rng.NumberFormat = "@"
                        rng.Value = cv
                    kept = [rng.Address.replace('$', '')]
            written_rng = rng
            print(f"書き込み: {ws.Name}!{rng.Address} ← {inline_value}")
        else:
            path = (smart_path_resolve(tsv_opt) if tsv_opt else _LAST_VALUES_FILE)
            if not path or not os.path.exists(path):
                print(f"エラー: TSVが見つかりません: {tsv_opt or _LAST_VALUES_FILE}")
                print("  単一値ならインラインで: write-range A1 \"値\"")
                return False
            grid = _read_tsv_grid(path, raw=raw)
            if not grid:
                print("エラー: TSVが空です")
                return False
            nrows = len(grid)
            lens = {len(r) for r in grid}
            top = ws.Cells(rng.Row, rng.Column)
            if nrows == 1 and len(grid[0]) == 1:
                _n既存 = _count_existing_cells(top)
                if _n既存 > 0:
                    print(f"上書き: 書込先に既存データ {_n既存} セルがあります")
                if raw:
                    top.NumberFormat = "@"
                else:
                    _untext_one(top, grid[0][0])
                _write_value(top, grid[0][0])
                if not raw and _kept_as_text(top, grid[0][0]):
                    kept = [top.Address.replace('$', '')]
                written_rng = top
                print(f"書き込み: {ws.Name}!{top.Address} ← {grid[0][0]}")
            elif len(lens) == 1:
                ncols = len(grid[0])
                target = ws.Range(top, ws.Cells(rng.Row + nrows - 1,
                                                rng.Column + ncols - 1))
                _n既存 = _count_existing_cells(target)
                if _n既存 > 0:
                    print(f"上書き: 書込先に既存データ {_n既存} セルがあります")
                if raw:
                    target.NumberFormat = "@"
                else:
                    _untext_grid(ws, target, rng.Row, rng.Column, grid)
                target.Value = grid
                if not raw:
                    kept = _keep_grid_text(ws, rng.Row, rng.Column, grid)
                    _fix_dynamic_formulas(ws, rng.Row, rng.Column, grid)
                written_rng = target
                print(f"書き込み: {ws.Name}!{target.Address} ← TSV {nrows}行 x {ncols}列")
            else:
                # 行の長さが不揃い: 矩形化して None を書くと右側の既存セルが消えるため、
                # 行ごとに実際の長さぶんだけ書き込む
                print(f"⚠ TSVの行の長さが不揃いです（{min(lens)}〜{max(lens)}列）。"
                      "行ごとに書き込み、短い行の右側セルには触れません。")
                _wide = ws.Range(top, ws.Cells(rng.Row + nrows - 1,
                                               rng.Column + max(lens) - 1))
                _n既存 = _count_existing_cells(_wide)
                if _n既存 > 0:
                    print(f"上書き: 書込先に既存データ {_n既存} セルがあります（外接矩形で数えた概数）")
                if not raw:
                    _untext_grid(ws, _wide, rng.Row, rng.Column, grid)
                for i, row in enumerate(grid):
                    if not row:
                        continue
                    r_tgt = ws.Range(ws.Cells(rng.Row + i, rng.Column),
                                     ws.Cells(rng.Row + i, rng.Column + len(row) - 1))
                    if raw:
                        r_tgt.NumberFormat = "@"
                    r_tgt.Value = (row,)
                if not raw:
                    kept = _keep_grid_text(ws, rng.Row, rng.Column, grid)
                    _fix_dynamic_formulas(ws, rng.Row, rng.Column, grid)
                written_rng = _wide
                print(f"書き込み: {ws.Name}!{top.Address} 起点 ← TSV {nrows}行（不揃い）")
    finally:
        _dlg_watcher.stop()

    note = _dialog_watcher_note(_dlg_watcher, None)
    if note:
        print(note, file=sys.stderr)
    if kept:
        print(_KEPT_NOTE + " ".join(kept))
    if written_rng is not None:
        _report_write_result(ws, written_rng)
        if getattr(args, 'show', False):
            _show_range(ws, written_rng)
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


@protect_safe
def cmd_write_cells(args):
    """飛び飛びのセルを1回で書く: write-cells <セル> <値> [<セル> <値> ...] | --tsv file [--raw] [--show]

    「C7 と C11 と D7 と D11 を直す」を 4 往復でなく 1 往復にする（2026-09-02 夜・同じ仕事の
    実測 44 秒→18.5 秒で、時間の支配項は道具でなく往復のあいだの考え込みと判明した日の改修）。
    値は write-range と同じ規約（'=' 始まりは数式、数値は数値、--raw で文字列のまま）。
    --tsv は「セル<TAB>値」を 1 行 1 セルで並べたファイル（値にスペースや引用符があるとき用）。
    先に全セルを解決してから書く（途中でエラーなら 1 セルも書かない）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    tsv_opt = getattr(args, 'tsv_opt', None)
    raw = getattr(args, 'raw', False)
    usage = ("使い方: write-cells [excel_file] <セル> <値> [<セル> <値> ...] [--raw] [--show]\n"
             "        write-cells [excel_file] --tsv ファイル   （「セル<TAB>値」を1行1セル）")
    pairs = []
    if tsv_opt:
        if rest:
            print("エラー: --tsv とインラインの「セル 値」は同時に指定できません")
            return False
        path = smart_path_resolve(tsv_opt)
        if not path or not os.path.exists(path):
            print(f"エラー: TSVが見つかりません: {tsv_opt}")
            return False
        with open(path, 'r', encoding='utf-8-sig') as f:
            text = f.read()
        for ln, line in enumerate(text.replace('\r\n', '\n').replace('\r', '\n').split('\n'), 1):
            if not line.strip():
                continue
            if '\t' not in line:
                print(f"エラー: {ln}行目にタブがありません（「セル<TAB>値」で1行1セル）: {line[:40]}")
                return False
            addr, val = line.split('\t', 1)
            pairs.append((addr.strip(), val))
    else:
        if not rest or len(rest) % 2 != 0:
            print(usage)
            if rest:
                print(f"  引数が {len(rest)} 個＝「セル 値」の組になっていません")
            return False
        pairs = list(zip(rest[0::2], rest[1::2]))
    if not pairs:
        print("エラー: 書くセルがありません")
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    targets = []
    for addr, val in pairs:
        if _whole_sheet_spec(wb, addr, sheet_opt) is not None:
            print(f"エラー: '{addr}' はシートの使用範囲全域を指します。セルを明示してください")
            return False
        ws, rng = _resolve_range(xl, wb, addr, sheet_opt)
        _unfilter_for_write(ws)
        if int(rng.Cells.CountLarge) != 1:
            print(f"エラー: '{addr}' はセル1つではありません（{rng.Address}）。範囲は write-range で")
            return False
        targets.append((ws, rng, val))
    n_exist = sum(1 for ws, c, v in targets if _count_existing_cells(c) > 0)
    if n_exist:
        print(f"上書き: 書込先に既存データ {n_exist} セルがあります")
    _dlg_watcher = _start_dialog_watcher(xl)
    kept = []                        # Excel の読み替えから文字列に戻したセル（報告用）
    try:
        for ws, c, val in targets:
            if raw:
                c.NumberFormat = "@"
                c.Value = val
            else:
                cv = _coerce_cell(val)
                _untext_one(c, cv)
                _write_value(c, cv)
                if _kept_as_text(c, cv):
                    kept.append(f"{ws.Name}!{c.Address.replace('$', '')}")
            print(f"書き込み: {ws.Name}!{c.Address} ← {val}")
    finally:
        _dlg_watcher.stop()
    note = _dialog_watcher_note(_dlg_watcher, None)
    if note:
        print(note, file=sys.stderr)
    if kept:
        print(_KEPT_NOTE + " ".join(kept))
    # 書込後の検証（write-range の _report_write_result と同じ思想。単セルの SpecialCells は
    # 使用範囲全体に化けるので、セルごとに IsError で見る）
    errs = []
    for ws, c, val in targets:
        try:
            if bool(xl.WorksheetFunction.IsError(c)):
                errs.append(f"{ws.Name}!{c.Address.replace('$', '')}={c.Text}")
        except Exception:
            pass
    if errs:
        print(f"書き込み後のエラーセル: {len(errs)}個 （" + " ".join(errs[:5])
              + ("…" if len(errs) > 5 else "") + "）")
    if getattr(args, 'show', False):
        print("表示（画面に見えている文字）:")
        for ws, c, val in targets:
            try:
                t = str(c.Text)
            except Exception:
                t = ""
            flag = "   ← '###' 列幅不足" if t and t.replace("#", "") == "" else ""
            print(f"  {ws.Name}!{c.Address.replace('$', '')}: {t}{flag}")
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


@protect_safe
@dialog_safe
def cmd_clear_range(args):
    """セル範囲をクリア (既定: すべて)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: clear-range [excel_file] <range> [--contents|--formats|--all]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '範囲は「シート名!範囲」の単一引数で指定してください'
                                   '（例: clear-range "Sheet1!A1:B2" --contents）'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を本当にクリアするなら --whole-sheet を付けてください。")
        print(f"  範囲を消すつもりなら「シート名!範囲」で指定してください（例: \"{whole}!A1:B2\"）")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    _unfilter_for_write(ws)
    if getattr(args, 'contents', False):
        rng.ClearContents(); what = "値"
    elif getattr(args, 'formats', False):
        rng.ClearFormats(); what = "書式"
    else:
        rng.Clear(); what = "すべて"
    print(f"クリア({what}): {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_format_range(args):
    """セル範囲に書式・整形を適用"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: format-range [excel_file] <range> [オプション...]")
        print("  --font 名 --size N --bold --unbold --italic --plain（斜体・下線・取り消し線を外す）")
        print("  --color '#RRGGBB' --bg '#RRGGBB' --number-format 書式")
        print("  --align left|center|right --valign top|center|bottom --wrap")
        print("  --border thin|medium|thick|hairline|none")
        print("  --col-width N --row-height N --merge --unmerge --autofit")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: format-range [excel_file] <range> [オプション...]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が同じ書式で塗られ、--merge 併用なら全域結合で値が消える危険があります。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    applied = []

    if getattr(args, 'font', None):
        rng.Font.Name = args.font; applied.append(f"font={args.font}")
    if getattr(args, 'size', None):
        rng.Font.Size = float(args.size); applied.append(f"size={args.size}")
    if getattr(args, 'bold', False):
        rng.Font.Bold = True; applied.append("bold")
    if getattr(args, 'unbold', False):
        rng.Font.Bold = False; applied.append("unbold")
    if getattr(args, 'italic', False):
        rng.Font.Italic = True; applied.append("italic")
    if getattr(args, 'plain', False):
        # 飾りを外す（2026-09-11: 「斜体・下線・取り消し線のばらつきをそろえて」を外す手が無く、AI が落とした）
        rng.Font.Italic = False
        rng.Font.Underline = -4142                   # xlUnderlineStyleNone
        rng.Font.Strikethrough = False
        applied.append("plain")
        if not getattr(args, 'color', None) and rng.Font.Color is None:
            # 文字色がセルごとにまだらなら黒に（2026-09-11: AI が font・size・plain は当てて color だけ落とし、
            # 仕上げ検査の差し戻しで 1 往復失った）。1 色にそろっている範囲の色は触らない
            rng.Font.Color = 0
            applied.append("color=#000000（まだらだったので）")
        # 書体・大きさも同じ（2026-09-11 14:00: 今度は font と size を落として 1 往復失った）。
        # まだらなら Excel の標準の書体・大きさに。そろっている範囲は触らない
        if not getattr(args, 'font', None) and rng.Font.Name is None:
            try:
                std = str(xl.StandardFont)
                rng.Font.Name = std
                applied.append(f"font={std}（まだらだったので標準に）")
            except Exception:
                pass
        if not getattr(args, 'size', None) and rng.Font.Size is None:
            try:
                std = float(xl.StandardFontSize)
                rng.Font.Size = std
                applied.append(f"size={std:g}（まだらだったので標準に）")
            except Exception:
                pass
    if getattr(args, 'color', None):
        rng.Font.Color = _hex_to_excel_color(args.color); applied.append(f"color={args.color}")
    if getattr(args, 'bg', None):
        if str(args.bg).strip().lower() in ('none', 'なし'):
            rng.Interior.Pattern = -4142              # 塗りを消す（白で塗るのとは別・2026-09-11）
            applied.append("bg=none")
        else:
            rng.Interior.Color = _hex_to_excel_color(args.bg); applied.append(f"bg={args.bg}")
    if getattr(args, 'number_format', None):
        nf = str(args.number_format)
        if nf.strip().lower() in ('general', 'standard'):
            # 日本語 Excel の NumberFormatLocal は 'General' を拒む（標準は 'G/標準'・2026-09-04）
            for fmt in (_general_format(rng.Cells(1, 1)), 'G/標準', 'General'):
                try:
                    rng.NumberFormat = fmt
                    break
                except Exception:
                    continue
        else:
            rng.NumberFormatLocal = nf
        applied.append(f"numfmt={nf}")
    if getattr(args, 'align', None):
        rng.HorizontalAlignment = _XL_ALIGN_H[args.align]; applied.append(f"align={args.align}")
    if getattr(args, 'valign', None):
        rng.VerticalAlignment = _XL_ALIGN_V[args.valign]; applied.append(f"valign={args.valign}")
    if getattr(args, 'wrap', False):
        rng.WrapText = True; applied.append("wrap")
    if getattr(args, 'border', None):
        if args.border == 'none':
            rng.Borders.LineStyle = -4142            # xlNone
            applied.append("border=none")
        else:
            rng.Borders.LineStyle = 1                # xlContinuous
            rng.Borders.Weight = _XL_BORDER_WEIGHT.get(args.border, 2)
            applied.append(f"border={args.border}")
    if getattr(args, 'col_width', None) is not None:
        rng.ColumnWidth = float(args.col_width); applied.append(f"col-width={args.col_width}")
    if getattr(args, 'row_height', None) is not None:
        # 隠れている行に行高を入れると、その行が**見えるようになる**（隠す＝高さ 0 なので）。
        # 人が手で隠した行を、見た目を整えただけで表に出してしまう＝頼んでいない変化
        # （2026-09-08 の実射・手順書「見た目を直す」が「隠れていた行が変わった（4・5 行目）」で落ちた）。
        # 隠れていた行を覚えて、高さを入れた後に隠し直す。
        was_hidden = []
        try:
            r1 = int(rng.Row)
            for i in range(r1, r1 + int(rng.Rows.Count)):
                if bool(ws.Rows(i).Hidden):
                    was_hidden.append(i)
        except Exception:
            was_hidden = []
        rng.RowHeight = float(args.row_height)
        for i in was_hidden:
            try:
                ws.Rows(i).Hidden = True
            except Exception:
                continue
        applied.append(f"row-height={args.row_height}"
                       + (f"（隠れていた {len(was_hidden)} 行は隠したまま）" if was_hidden else ""))
    if getattr(args, 'merge', False):
        # 複数の値を含む範囲の Merge は Excel が確認ダイアログを出し、
        # CLI が無言で応答待ちブロックする。値の個数を先に数えて、
        # 消える場合は警告した上で DisplayAlerts を切って実行する（左上の値が残る）
        n_vals = 0
        try:
            n_vals = int(xl.WorksheetFunction.CountA(rng))
        except Exception:
            pass
        if n_vals > 1:
            print(f"⚠ 範囲に値が{n_vals}個あります。結合により左上以外の値は消えます。")
        with _alerts_off(xl):
            rng.Merge()
        applied.append("merge")
    if getattr(args, 'unmerge', False):
        rng.UnMerge(); applied.append("unmerge")
    if getattr(args, 'autofit', False):
        rng.Columns.AutoFit(); applied.append("autofit")
    if getattr(args, 'lock', False):
        rng.Locked = True; applied.append("lock（シート保護時に有効）")
    if getattr(args, 'unlock', False):
        rng.Locked = False; applied.append("unlock（シート保護時に有効）")

    if not applied:
        print("書式オプションが指定されていません。--bold --bg '#FFFF00' などを指定してください。")
        return False
    print(f"書式適用: {ws.Name}!{rng.Address}  [{', '.join(applied)}]")
    if getattr(args, 'show', False):
        _show_range(ws, rng)
    print("（保存はしていません）")
    return True


def _looks_like_header(cell):
    """セルが見出しの体裁か（太字か塗りつぶしあり）。

    tidy が --header-from 無しで「表の左上＝既存の見出し」を見つける目安。既存の表に列を足した
    ときは左上が既に見出しの体裁なので、その書式を行全体へ複写すれば新しい列も揃う
    （2026-09-02 夜・「--header-from を書くか」で迷っていた既定を道具に持たせた）。
    """
    try:
        if cell.Font.Bold is True:
            return True
    except Exception:
        pass
    try:
        return int(cell.Interior.ColorIndex) != -4142      # xlNone 以外＝塗りつぶしあり
    except Exception:
        return False


# tidy の「列の型」判定（2026-09-02 深夜・「会員番号は左寄せ・金額はカンマが入っていて当然。
# 道具でそうなるようにしろ」）。見出しの語と値の型だけで決める。迷う列は触らない。
_ID_KW_JA = ('番号', '記号', '型番', '品番')
# 「キー」「コード」は、後ろ・前にカタカナが続く語（キーボード・キーワード・レコード）を番号列にしない
# （2026-09-19・商品名「キーボード」の売上の列が番号列と判定され、左寄せ・カンマなしになった）
_ID_KW_KANA = re.compile(r'キー(?![ァ-ヶー])|(?<!レ)コード(?![ァ-ヶー])')
_ID_KW_EN = re.compile(r'(?<![a-z])(no|id|cd|code|key)(?![a-z])')


def _is_id_header(h):
    """見出し（NFKC 済みの文字）が番号列の語を持つか。道具の中で番号列の定義を 1 つにする。"""
    h = str(h or '')
    return (any(k in h for k in _ID_KW_JA) or bool(_ID_KW_KANA.search(h))
            or bool(_ID_KW_EN.search(h.lower())))
_SKIP_KW = ('年度', '年月', '日付', '期日', '期間', '時刻', '時間', '西暦', '和暦', '率', '割合', '%')
# 金額とわかる見出しの列は、欠測の文字（「申込なし」「-」など）が混じっていても数値の側で判定する
# （2026-09-06・突き合わせで「申込なし」を入れた金額列にカンマが付かず、同じ依頼の 1 回目と 2 回目で
# 見た目が食い違った）。数値が 2 つ以上あるときだけ。
_MONEY_KW = ('金額', '単価', '価格', '料金', '合計', '小計', '売上', '費用', '原価', '残高', '税込', '税抜', '円')
_SKIP_TAIL = ('年', '月', '日')
_XL_HALIGN_GENERAL = 1          # xlHAlignGeneral
_XL_HALIGN_LEFT = -4131         # xlHAlignLeft
# 「標準」の表示形式。日本語 Excel の NumberFormat は 'General' でなく 'G/標準' を返す（2026-09-02 実測・
# これを 'General' 一択で見ていて金額列にカンマが付かなかった）
_GENERAL_FMTS = {'general', 'g/標準', 'standard'}


def _numeric_formats(rng):
    """範囲の中で「数値・日付が入っていて、標準ではない表示形式を持つ」セルを覚える。

    見出しの書式を複写する前に呼ぶ（2026-09-08）。xlPasteFormats は表示形式まで複写するので、
    先頭行に数値がある表（合計ブロック・年を見出しにした表）でカンマや桁が消える。
    戻り値: [(列番号, 表示形式) …]
    """
    keep = []
    for j in range(1, int(rng.Columns.Count) + 1):
        try:
            c = rng.Cells(1, j)
            v, fmt = c.Value, c.NumberFormat
            if isinstance(v, bool) or v is None or isinstance(v, str):
                continue
            if str(fmt).lower() in _GENERAL_FMTS:
                continue
            keep.append((j, fmt))
        except Exception:
            continue
    return keep


def _restore_formats(rng, keep):
    """_numeric_formats で覚えた表示形式を書き戻す。戻り値: 戻したセル数。"""
    n = 0
    for j, fmt in keep:
        try:
            rng.Cells(1, j).NumberFormat = fmt
            n += 1
        except Exception:
            continue
    return n


def _is_whole_number(v):
    try:
        return float(v) == int(v)
    except (OverflowError, ValueError):
        return False


def _column_style(header, values):
    """列の型を見出しと値から判定し、tidy が当てる寄せ・表示形式を返す。

    戻り値: 'id'（番号・コード列＝左寄せ）／'int'（整数だけの数値列＝#,##0）／
            'dec'（小数を含む数値列＝#,##0.00）／None（触らない）。
    見出しに 番号・コード・記号・型番・品番・キー・No・ID・CD・code・key があれば番号列。
    年・年度・日付・期間・時刻・率・割合・% の見出しは触らない（2026 が 2,026 になる事故を防ぐ）。
    値が全部数値（空欄と "" は無視）で番号列でなければ数値列。文字や TRUE/FALSE が混じる列は触らない。
    """
    h = unicodedata.normalize('NFKC', str(header or '')).strip()
    if _is_id_header(h):
        return 'id'
    if any(k in h for k in _SKIP_KW) or h.endswith(_SKIP_TAIL):
        return None
    money = any(k in h for k in _MONEY_KW)
    nums, others = [], 0
    for v in values:
        if v is None or (isinstance(v, str) and v.strip() == ''):
            continue
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            if money:
                others += 1
                continue
            return None
        nums.append(v)
    if not nums or (others and len(nums) < 2):
        return None
    if all(-1.5 <= v <= 1.5 for v in nums) and any(not _is_whole_number(v) for v in nums):
        # 率の列（0.62 のような小数）は触らない＝見出しに「率」「割合」が無い言い方（執行状況・消化 など）もある
        # （2026-09-18: 集約の率の列を「数値列の表示形式がまちまち」と止めた）
        return None
    # 小数が 1 つでもあれば列ごと #,##0.00 にしていた＝金額 20 行の下の「平均 10,162.5」1 つで、7,200 が 7,200.00 に
    # なった（2026-09-13 お試し版テスト用1）。小数が 3 分の 1 に満たない列は整数の列（小数のセルだけ別に形を当てる）
    frac = sum(1 for v in nums if not _is_whole_number(v))
    return 'dec' if frac * 3 >= len(nums) else 'int'


def _all_dates(vals):
    """埋まっている値が全部 日付（datetime）か。空の列は False。"""
    filled = [v for v in vals if v not in (None, '')]
    return bool(filled) and all(isinstance(v, datetime.datetime) for v in filled)


def _tidy_column_styles(ws, rng):
    """tidy の一工程: 列の型に合わせて、番号列は左寄せ・数値列はカンマ区切り（#,##0）にする。

    既に誰かが決めた書式（表示形式が General 以外／寄せが標準以外／列内で混在）は上書きしない。
    戻り値は報告用の短い文字列の一覧（変更した列の列文字つき）。
    """
    nrows, ncols = int(rng.Rows.Count), int(rng.Columns.Count)
    if nrows < 2:
        return []
    done = {'id': [], 'int': [], 'dec': [], 'mixed': [], 'date': [], 'frac': []}
    for j in range(1, ncols + 1):
        data = ws.Range(rng.Cells(2, j), rng.Cells(nrows, j))
        vals = data.Value
        vals = [vals] if nrows == 2 else [row[0] for row in vals]   # 1セルは素の値、複数は行タプル
        style = _column_style(rng.Cells(1, j).Value, vals)
        col_letter = re.sub(r'[\d$]', '', str(rng.Cells(1, j).Address))
        if style is None and _all_dates(vals):
            # 日付だけの列で表示形式がまだら（令和と西暦が混在）なら 1 つに（2026-09-11 14:03: 文字の日付を
            # 日付にした後、AI が表示形式を当て忘れ、採点役の差し戻しで 1 往復失った）。そろっている列は触らない
            try:
                if data.NumberFormat is None:
                    data.NumberFormat = 'yyyy/m/d'
                    done['date'].append(col_letter)
            except Exception:
                pass
        try:
            if data.HorizontalAlignment is None:
                # 列の中で寄せがまちまち（2026-09-11: メールアドレスの列が左・中央・右ばらばらのまま「合格」で出た）。
                # 文字の列と番号の列は左、数・日付の列は標準（右）にそろえる
                filled = [v for v in vals if v not in (None, '')]
                texty = filled and sum(isinstance(v, str) for v in filled) * 2 >= len(filled)
                data.HorizontalAlignment = (_XL_HALIGN_LEFT if style == 'id' or (style is None and texty)
                                            else _XL_HALIGN_GENERAL)
                done['mixed'].append(col_letter)
        except Exception:
            pass
        if style is None:
            continue
        try:
            if style == 'id':
                if data.HorizontalAlignment == _XL_HALIGN_GENERAL:     # None＝混在 → 触らない
                    data.HorizontalAlignment = _XL_HALIGN_LEFT
                    done['id'].append(col_letter)
            elif data.NumberFormat is None or str(data.NumberFormat).lower() in _GENERAL_FMTS:
                # 既に 1 つの書式がある列は触らない。None＝列の中で混在（#,##0 と標準がまだら）はそろえる
                # （2026-09-11: 文字の金額を数値にした列が、元の #,##0 のセルと標準のセルのまだらで残った）
                frac_cells = []
                if style == 'int':
                    # 整数の列の中の小数（金額の下の「平均 10,162.5」など）は、列を #,##0 にすると 10,163 と
                    # 丸めて見える。先に元の形を覚えておき、そのセルだけ小数の形にする（2026-09-13）
                    for i, v in enumerate(vals):
                        if isinstance(v, (int, float)) and not isinstance(v, bool) and not _is_whole_number(v):
                            try:
                                c = data.Cells(i + 1, 1)
                                frac_cells.append((c, c.NumberFormat))
                            except Exception:
                                pass
                data.NumberFormat = '#,##0' if style == 'int' else '#,##0.00'
                done[style].append(col_letter)
                for c, prev in frac_cells:
                    try:
                        # 人が決めていた形（#,##0.0 等）はそのまま戻す。標準だったセルだけ小数の形にする
                        c.NumberFormat = ('#,##0.0#' if prev is None or str(prev).lower() in _GENERAL_FMTS
                                          else prev)
                        done['frac'].append(col_letter)
                    except Exception:
                        pass
        except Exception:
            continue
    notes = []
    if done['mixed']:
        notes.append(f"寄せのばらつきをそろえた({','.join(done['mixed'])}。文字は左・数と日付は標準)")
    if done['id']:
        notes.append(f"番号列を左寄せ({','.join(done['id'])})")
    if done['int']:
        notes.append(f"数値列を#,##0({','.join(done['int'])})")
    if done['dec']:
        notes.append(f"小数列を#,##0.00({','.join(done['dec'])})")
    if done['frac']:
        notes.append(f"整数の列の中の小数のセルだけ#,##0.0#({','.join(sorted(set(done['frac'])))}・{len(done['frac'])} セル)")
    if done['date']:
        notes.append(f"日付列の表示形式をそろえた({','.join(done['date'])}。yyyy/m/d)")
    return notes


# ----------------------------------------------------------------
# 仕上げ検査（audit・2026-09-04）: 「値は合っているのに見た目がひどい」を道具が機械的に見つける
#   ※ AI を使わない。tidy が当てるはずのもの（見出し・罫線・列の型・列幅）が当たっているかを実物で見る。
#   ※ agent はこれを done の関所に使う（検査に引っかかるうちは done を受け付けない）。
# ----------------------------------------------------------------
_XL_EDGES = {'left': 7, 'top': 8, 'bottom': 9, 'right': 10}       # xlEdgeLeft/Top/Bottom/Right
_XL_INSIDE = {'inside_v': 11, 'inside_h': 12}                     # xlInsideVertical/Horizontal
_EDGE_JA = {'left': '左', 'top': '上', 'bottom': '下', 'right': '右', 'inside_v': '縦', 'inside_h': '横'}
_XL_LINESTYLE_NONE = -4142                               # xlLineStyleNone
_AUDIT_MAX_CELLS = 4000                                  # ### を見るセル数の上限（大きい表で時間を食わない）


def _border_state(rng, idx):
    """罫線の状態 → 'あり' | 'なし' | 'まちまち'（None＝範囲内で混在）。読めなければ 'なし'。"""
    try:
        ls = rng.Borders(idx).LineStyle
    except Exception:
        return 'なし'
    if ls is None:
        return 'まちまち'
    return 'なし' if int(ls) == _XL_LINESTYLE_NONE else 'あり'


def _hash_cells(ws, rng):
    """列幅が足りずに '###' になっているセルの番地（先頭 5 個）と総数。

    範囲を 1 回で読んで怪しいセルだけ .Text で確かめる（_hash_scan）。1 セルずつ読む形は
    78 セルで 30 秒かかった（2026-09-13）。
    """
    nrows, ncols = int(rng.Rows.Count), int(rng.Columns.Count)
    if nrows * ncols > _AUDIT_MAX_CELLS:
        nrows = max(1, _AUDIT_MAX_CELLS // max(1, ncols))       # 上から見る（列幅は列で決まる＝上で足りる）
    n, bad, _note = _hash_scan(rng, max_rows=nrows, max_cols=ncols)
    return bad, n


def _outside_hash_cells(ws, rng, col_index):
    """その列の、表の範囲より上・下にある '###' セルの番地（先頭 3 個）。

    tidy の列幅は列全体に効くので、整えた表の外が読めなくなることがある。ここはその見張り。
    """
    bad = []
    try:
        used = ws.UsedRange
        top, bottom = int(used.Row), int(used.Row) + int(used.Rows.Count) - 1
        r1 = int(rng.Row)
        r2 = r1 + int(rng.Rows.Count) - 1
    except Exception:
        return bad
    for r in range(top, min(bottom, top + _AUDIT_MAX_CELLS) + 1):
        if r1 <= r <= r2:
            continue
        try:
            c = ws.Cells(r, col_index)
            t = str(c.Text)
        except Exception:
            continue
        if t and t.replace('#', '') == '':
            bad.append(str(c.Address).replace('$', ''))
            if len(bad) >= 3:
                break
    return bad


def _undo_narrowing_that_broke_outside(ws, rng, before_w):
    """幅を細くしたせいで表の外が '###' になった列を、元の幅へ戻す。戻した列文字の一覧を返す。"""
    restored = []
    for j, old in (before_w or {}).items():
        try:
            col = rng.Columns(j)
            now = float(col.ColumnWidth)
        except Exception:
            continue
        if now >= old:
            continue                                  # 広げた・変えていない列は関係ない
        try:
            col_index = int(col.Column)
        except Exception:
            continue
        if not _outside_hash_cells(ws, rng, col_index):
            continue
        try:
            col.ColumnWidth = old
            restored.append(_col_letters(col_index))
        except Exception:
            continue
    return restored


def _next_block_is_continuation(width, block_first_col, block_last_col, first_col, last_col):
    """空行 1 行の下の塊が表の続きか（純 Python）。表と同じ幅以上（表の列を全部覆う）なら続き、細ければ別の塊（合計欄・注記）。"""
    return block_first_col <= first_col and block_last_col >= last_col and width >= 2


def _over_single_blank_rows(ws, rng, limit=200):
    """CurrentRegion を、途中の空行 1 行を越えて表の終わりまで伸ばす（表を整えるマクロと同じ決まり:
    空行 1 行は続き・2 行続いたら終わり・空行の下の塊が表より細ければ別の塊）。
    2026-09-17: 途中に空行のある表で tidy が上の塊だけ整え、下の塊の番号列が左寄せにならず「寄せがまちまち」→
    マクロで合っている表を AI に回していた（試験の顧客名簿・在庫表。9/12 に試験へ空行を足してから）。"""
    try:
        wf = ws.Application.WorksheetFunction
        r0, c0 = int(rng.Row), int(rng.Column)
        r2, c2 = r0 + int(rng.Rows.Count) - 1, c0 + int(rng.Columns.Count) - 1
        for _ in range(limit):
            if wf.CountA(ws.Range(ws.Cells(r2 + 1, c0), ws.Cells(r2 + 1, c2))) != 0:
                break
            below = ws.Range(ws.Cells(r2 + 2, c0), ws.Cells(r2 + 2, c2))
            if wf.CountA(below) == 0:
                break
            cell = next((below.Cells(1, k) for k in range(1, c2 - c0 + 2)
                         if below.Cells(1, k).Value not in (None, '')), None)
            if cell is None:
                break
            cr = cell.CurrentRegion
            b1, b2 = int(cr.Column), int(cr.Column) + int(cr.Columns.Count) - 1
            if not _next_block_is_continuation(b2 - b1 + 1, b1, b2, c0, c2):
                break
            nb = int(cr.Row) + int(cr.Rows.Count) - 1
            if nb <= r2:
                break
            r2, c2 = nb, max(c2, b2)
        return ws.Range(ws.Cells(r0, c0), ws.Cells(r2, c2))
    except Exception:
        return rng


def _over_single_blank_rows_up(ws, rng, limit=200):
    """CurrentRegion を、上の空行 1 行を越えて表の頭まで伸ばす（_over_single_blank_rows の上向き）。
    伸ばすのは、いまの塊の 1 行目に数がある（＝見出しでなく本文の続き）で、上の塊がいまの塊の列を覆うときだけ。
    2026-09-18: 途中に空行のある表の下に合計の行を足すと、入口の検査が空行の下の塊だけを表と見て
    「合計が上の和と合いません」と止めた（正解の表でも止まる形＝鍛えたマクロが毎回 AI に回る）。"""
    try:
        wf = ws.Application.WorksheetFunction
        r0, c0 = int(rng.Row), int(rng.Column)
        r2, c2 = r0 + int(rng.Rows.Count) - 1, c0 + int(rng.Columns.Count) - 1
        for _ in range(limit):
            if r0 <= 2 or wf.CountA(ws.Range(ws.Cells(r0 - 1, c0), ws.Cells(r0 - 1, c2))) != 0:
                break
            above = ws.Range(ws.Cells(r0 - 2, c0), ws.Cells(r0 - 2, c2))
            if wf.CountA(above) == 0:
                break
            first = ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c2)).Value
            first = first[0] if isinstance(first, tuple) else (first,)
            if not any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in first):
                break                                    # いまの塊の頭が見出し＝別の表
            cell = next((above.Cells(1, k) for k in range(1, c2 - c0 + 2) if above.Cells(1, k).Value not in (None, '')), None)
            if cell is None:
                break
            cr = cell.CurrentRegion
            b1, b2 = int(cr.Column), int(cr.Column) + int(cr.Columns.Count) - 1
            if not _next_block_is_continuation(c2 - c0 + 1, c0, c2, b1, b2):
                break
            nr0 = int(cr.Row)
            if nr0 >= r0:
                break
            r0, c0, c2 = nr0, min(c0, b1), max(c2, b2)
        return ws.Range(ws.Cells(r0, c0), ws.Cells(r2, c2))
    except Exception:
        return rng


def _region_of_cell(ws, rng):
    """セル 1 つ → それを含む表。CurrentRegion は表題（見出しの真上の 1 セル）まで飲み込むので、材料と同じ見出しの推定
    （vbam_hands._pick_header_row）で見出しの行から始める（2026-09-17: 表題が見出しの真上にある表で tidy が表題を見出しに
    して太字にし、検査が「見出しが本文と同じ体裁」「見出しが空の列」と出て、マクロで合っている表を AI に回していた）。"""
    rng = _over_single_blank_rows(ws, _over_single_blank_rows_up(ws, rng.CurrentRegion))
    try:
        from vbam_hands import _pick_header_row, _rows_of
        n = min(int(rng.Rows.Count), 8)
        r0, c0, nc = int(rng.Row), int(rng.Column), int(rng.Columns.Count)
        grid = [list(r) for r in _rows_of(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n - 1, c0 + nc - 1)).Value)]
        got = _pick_header_row(grid)
        if got and int(got[0]) > 0:
            r1, r2 = r0 + int(got[0]), r0 + int(rng.Rows.Count) - 1
            if r2 - r1 + 1 >= 2:
                rng = ws.Range(ws.Cells(r1, c0), ws.Cells(r2, c0 + nc - 1))
    except Exception:
        pass
    return rng


def _mixed_formats_nonblank(data, vals):
    """列の表示形式が「まちまち」のとき、値の入ったセルだけで見てもまちまちか（空セルの G/標準 は数えない）。"""
    fmts = set()
    try:
        for i, v in enumerate(vals):
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            fmts.add(str(data.Cells(i + 1, 1).NumberFormat))
            if len(fmts) > 1:
                return True
    except Exception:
        return True
    return len(fmts) > 1


def audit_table(ws, rng):
    """表の見た目を機械的に検査して、直すべき点の一覧（文字列）を返す。空なら合格。

    見るもの（tidy が当てるはずのもの）:
      1. 見出し行が本文と同じ体裁（太字でも塗りでもない）
      2. 罫線が無い・まちまち（＝継ぎはぎ）
      3. 数値だけの列に桁区切りが無い／番号列が左寄せでない
      4. 列幅が足りず '###' になっている
      5. 見出しが空のまま（表の途中の列）
    「値が合っているか」は見ない（それは依頼ごとの答え合わせの仕事）。
    """
    out = []
    nrows, ncols = int(rng.Rows.Count), int(rng.Columns.Count)
    if nrows < 2:
        return out                                   # 見出し＋1 行に満たないものは表として見ない
    addr = str(rng.Address).replace('$', '')
    in_table = False
    try:
        for lo in ws.ListObjects:
            if ws.Application.Intersect(lo.Range, rng) is not None:
                in_table = True
                break
    except Exception:
        in_table = False
    try:
        # ピボットの範囲は Excel が見出し（列ラベルの空き・「合計 / …」）と書式を決める＝表の見た目の検査を当てない
        # （2026-09-18 第二期: ピボットに累計を足す鍛えたマクロが、値は正解と同じなのに「見出しが空の列」で止まった）
        for pt in ws.PivotTables():
            if ws.Application.Intersect(pt.TableRange2, rng) is not None:
                return out
    except Exception:
        pass
    if in_table:
        # テーブル（ListObject）はスタイルが見出し・罫線・縞を引く＝セルの罫線や見出しの体裁・表示形式の検査は当てない
        # （2026-09-17 深夜: テーブルを作る鍛えたマクロが「罫線がありません」「桁区切りがありません」で毎回 AI に回る形だった）
        empty_heads = [re.sub(r'[\d$]', '', str(rng.Cells(1, j).Address)) for j in range(1, ncols + 1)
                       if rng.Cells(1, j).Value is None or str(rng.Cells(1, j).Value).strip() == '']
        if empty_heads:
            out.append("見出しが空の列があります（" + ",".join(empty_heads) + "）→ 列の名前を書く")
        return out
    if not _looks_like_header(rng.Cells(1, 1)):
        out.append(f"見出し行が本文と同じ体裁です（{addr} の 1 行目が太字でも塗りでもない）→ tidy を当てる")
    empty_heads = []
    for j in range(1, ncols + 1):
        v = rng.Cells(1, j).Value
        if v is None or str(v).strip() == '':
            empty_heads.append(re.sub(r'[\d$]', '', str(rng.Cells(1, j).Address)))
    if empty_heads:
        out.append("見出しが空の列があります（" + ",".join(empty_heads) + "）→ 列の名前を書く")
    states = {}
    for key, idx in list(_XL_EDGES.items()) + list(_XL_INSIDE.items()):
        if key == 'inside_v' and ncols < 2:
            continue
        if key == 'inside_h' and nrows < 2:
            continue
        states[key] = _border_state(rng, idx)
    missing = [_EDGE_JA[k] for k, st in states.items() if st == 'なし']
    mixed = [_EDGE_JA[k] for k, st in states.items() if st == 'まちまち']
    if states and len(missing) == len(states):
        out.append(f"罫線がありません（{addr}）→ tidy を当てる")
    elif missing or mixed:
        note = ("抜け: " + "・".join(missing) if missing else "")
        note += ("　まちまち: " + "・".join(mixed) if mixed else "")
        out.append(f"罫線が継ぎはぎです（{addr}　{note.strip()}）→ 表全体に tidy を当てる")
    for j in range(1, ncols + 1):
        data = ws.Range(rng.Cells(2, j), rng.Cells(nrows, j))
        try:
            vals = data.Value
        except Exception:
            continue
        vals = [vals] if nrows == 2 else [row[0] for row in vals]
        style = _column_style(rng.Cells(1, j).Value, vals)
        head = str(rng.Cells(1, j).Value or '').strip()
        col_letter = re.sub(r'[\d$]', '', str(rng.Cells(1, j).Address))
        if style is None:
            try:
                if _all_dates(vals) and data.NumberFormat is None:
                    out.append(f"日付列「{head}」({col_letter}) の表示形式がセルごとに違います → tidy を当てる（yyyy/m/d）")
            except Exception:
                pass
            continue
        try:
            if style == 'id':
                nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if nums and data.HorizontalAlignment == _XL_HALIGN_GENERAL:
                    out.append(f"番号列「{head}」({col_letter}) が右に寄っています → tidy を当てる（左寄せ）")
            elif data.NumberFormat is None and _mixed_formats_nonblank(data, vals):
                # 空のセル（合計行の単価など）の G/標準 は見た目に出ない＝数えない（2026-09-17）
                out.append(f"数値列「{head}」({col_letter}) の表示形式がセルごとに違います → tidy を当てる（#,##0）")
            elif str(data.NumberFormat).lower() in _GENERAL_FMTS:
                out.append(f"数値列「{head}」({col_letter}) に桁区切りがありません → tidy を当てる（#,##0）")
            else:
                # 先頭行そのものが数値のとき（合計ブロック・年を見出しにした表）は、その行だけ
                # 書式が抜けていても今まで誰も見ていなかった＝見出し行はデータとして検査していない。
                # 下の行に書式があるのに先頭行だけ素なら、それは抜け（2026-09-08・G27 の 203250）
                hv = rng.Cells(1, j).Value
                if (not isinstance(hv, bool) and isinstance(hv, (int, float))
                        and str(rng.Cells(1, j).NumberFormat).lower() in _GENERAL_FMTS):
                    where = str(rng.Cells(1, j).Address).replace('$', '')
                    out.append(f"先頭行の数値 {where} だけ桁区切りがありません"
                               f"（下の行は付いている）→ tidy を当てる（#,##0）")
        except Exception:
            continue
    bad, n_hash = _hash_cells(ws, rng)
    if n_hash:
        out.append(f"列幅が足りず '###' のセルが {n_hash} 個あります（{' '.join(bad)}）→ tidy を当てる（列幅）")
    return out


# ----------------------------------------------------------------
# 中身の検査（2026-09-06 夜・「通信簿の信用度 75」の穴）。audit_table は見た目しか見ない。
# 初実射で C10 の空欄を道具も AI も採点係も言わなかった＝仕事の中身は誰も検査していなかった。
# ここは値と数式だけを見る（AI を使わない・タダ）。依頼の知識が要らない、表として壊れているものだけ:
#   done を止める（blocking）: エラー値／「合計」行が上の和と合わない／数式の列に直値が混ざる（式の断ち切れ）
#   報告の気づき（noticed）  : 表の中の空欄／数値列に混ざる文字の数字／番号列の重複
# 気づきは直させない（値を作らせない）。人が読んで決める。
# ----------------------------------------------------------------
_XL_ERROR_NAMES = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
                   -2146826273: '#NUM!', -2146826265: '#REF!', -2146826252: '#VALUE!', -2146826262: '#GETTING_DATA',
                   -2146826256: '#SPILL!', -2146826255: '#CALC!'}
_SUM_ROW_RE = re.compile(r'^(総?合\s*計|総計|小計|計|合計額|total|sum)$', re.IGNORECASE)
# 小計の行は「小計」だけとは限らない（職場の表は「一般会計 小計」「投資的経費小計」のように区分の名前を前に付ける）。
# 2026-09-18: B 集計の「小計行」を鍛えたら、区分名つきの小計を本文の行と見なし、下の合計が二重に数えられて止まった
_SUBTOTAL_RE = re.compile(r'.*小計$')
_SPARSE_HEAD_KW = ('備考', 'メモ', '注', 'コメント', '摘要', 'remarks', 'memo', 'note')
_NUMTEXT_RE = re.compile(r'^[-+]?\d[\d,]*(\.\d+)?$')
_CONTENT_MAX_EXAMPLES = 8
_SAMEVALUE_MIN_ROWS = 5          # 突き合わせの全滅を疑うのに要る行数（これ未満は偶然そろう）


def _is_blank_value(v):
    return v is None or (isinstance(v, str) and v.strip(' 　\t\r\n') == '')


def _error_name(v):
    """COM が返すエラー値（VT_ERROR＝負の整数）か '#REF!' の文字 → エラー名。違えば None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, int) and v in _XL_ERROR_NAMES:
        return _XL_ERROR_NAMES[v]
    if isinstance(v, str) and v.startswith('#') and (v.endswith('!') or v.endswith('?') or v == '#N/A'):
        return v
    return None


def _num_of(v):
    if isinstance(v, bool) or _error_name(v) is not None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def _col_letters(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _content_findings(values, formulas, r0, c0, formulas_r1c1=None):
    """中身の検査の本体（純 Python・COM に触らない）。

    values: 2 次元の値（1 行目が見出し）。formulas: 同じ形の数式の文字（無ければ None）。
    formulas_r1c1: 同じ形の R1C1 の式（無ければ None＝式の形の比べ物をしない）。
    r0, c0: 左上の行・列番号（1 起点。番地を作るため）。戻り値: (done を止める指摘, 報告の気づき)。
    """
    rows = [list(r) for r in (values or [])]
    if len(rows) < 2:
        return [], []
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([None] * (width - len(r)))
    fml = None
    if formulas:
        fml = [list(r) for r in formulas]
        for r in fml:
            r.extend([None] * (width - len(r)))

    def addr(i, j):
        return f"{_col_letters(c0 + j)}{r0 + i}"

    def head(j):
        return str(rows[0][j] or '').strip()

    def summable(j):
        """足してよい列か（2026-09-06 深夜・会員番号まで足して全行を咎めた）。

        外すのは番号列・日付や率の列・値が年号ばかりの列。「4月」「5月」のような月の見出しは
        金額が入る列なので外さない（tidy の _SKIP_TAIL をそのまま使うと、月別集計が全部黙る）。
        """
        h = unicodedata.normalize('NFKC', head(j))
        if not h:
            return False
        if _is_id_header(h):
            return False
        if any(k in h for k in _SKIP_KW):
            return False
        nums = [_num_of(rows[i][j]) for i in range(1, len(rows))]
        nums = [n for n in nums if n is not None]
        if nums and all(-1.5 <= n <= 1.5 for n in nums) and any(not float(n).is_integer() for n in nums):
            # 率の列（0.62 のような小数）は足す列ではない＝合計行は和にならない。見出しに「率」「割合」が無い
            # 言い方（執行状況・消化 など）もあるので、値の形でも外す（2026-09-18: 集約の合計行で止まった）
            return False
        return not (nums and all(1900 <= n <= 2100 and float(n).is_integer() for n in nums))

    blocking, noticed = [], []
    # 1. エラー値
    errs = [(addr(i, j), _error_name(rows[i][j])) for i in range(len(rows)) for j in range(width)
            if _error_name(rows[i][j]) is not None]
    if errs:
        ex = " ".join(f"{a}({n})" for a, n in errs[:_CONTENT_MAX_EXAMPLES])
        blocking.append(f"エラー値のセルが {len(errs)} 個あります（{ex}{'…' if len(errs) > _CONTENT_MAX_EXAMPLES else ''}）"
                        "→ 参照先・式を直す（直せない・元からのものなら report に理由を書く）")
    # 行の種類: 空行／合計行／小計行／本文
    kinds = []
    for i, r in enumerate(rows):
        if i == 0:
            kinds.append('head')
            continue
        if all(_is_blank_value(v) for v in r):
            kinds.append('blank')
            continue
        label = ''
        for v in r:
            if isinstance(v, str) and v.strip():
                label = unicodedata.normalize('NFKC', v).strip()
                break
            if _num_of(v) is not None:
                break
        if label and _SUBTOTAL_RE.match(label):
            kinds.append('sub')
        elif label and _SUM_ROW_RE.match(label):
            kinds.append('sum')
        else:
            kinds.append('body')
    body_idx = [i for i, k in enumerate(kinds) if k == 'body']
    # 2. 合計・小計の行が上の和と合わない
    for i, k in enumerate(kinds):
        if k not in ('sum', 'sub'):
            continue
        if k == 'sub':
            start = max([x for x in range(i) if kinds[x] in ('sub', 'head')] or [0]) + 1
            above = [x for x in range(start, i) if kinds[x] == 'body']
        else:
            above = [x for x in range(1, i) if kinds[x] == 'body']
        if len(above) < 2:
            continue
        for j in range(width):
            cell = _num_of(rows[i][j])
            if cell is None or not summable(j):
                continue
            nums = [_num_of(rows[x][j]) for x in above]
            nums = [n for n in nums if n is not None]
            if len(nums) < 2:
                continue
            s = sum(nums)
            if abs(s - cell) > max(0.005, abs(s) * 1e-9):
                name = '小計' if k == 'sub' else '合計'
                blocking.append(f"「{name}」行の {addr(i, j)}（{head(j)}）が上の和と合いません"
                                f"（セル {cell:,.10g}／上の和 {s:,.10g}）→ 式か値を直す（範囲の取りこぼしか、直値のまま）")
    # 2.5 「合計」列（右端など）が左の本文の和と合わない（2026-09-06 深夜・縦だけ見ていた）
    for j in range(width):
        h = unicodedata.normalize('NFKC', head(j))
        if not (h and _SUM_ROW_RE.match(h)):
            continue
        left = [x for x in range(j) if summable(x) and not _SUM_ROW_RE.match(unicodedata.normalize('NFKC', head(x)))]
        if len(left) < 2:
            continue
        hits, misses = 0, []
        for i in body_idx:
            cell = _num_of(rows[i][j])
            if cell is None:
                continue
            nums = [_num_of(rows[i][x]) for x in left]
            nums = [n for n in nums if n is not None]
            if len(nums) < 2:
                continue
            s = sum(nums)
            if abs(s - cell) > max(0.005, abs(s) * 1e-9):
                misses.append((i, cell, s))
            else:
                hits += 1
        # 1 行も合わないなら「左の全部を足す列」ではない＝黙る（見当違いの咎めを出さない）
        if misses and hits >= max(1, (hits + len(misses)) * 0.6):
            for i, cell, s in misses[:_CONTENT_MAX_EXAMPLES]:
                blocking.append(f"「{head(j)}」列の {addr(i, j)} が左の和と合いません"
                                f"（セル {cell:,.10g}／左の和 {s:,.10g}）→ 式か値を直す（範囲の取りこぼしか、直値のまま）")
    # 3. 数式の列に直値が混ざる（式の断ち切れ）／式が本文の途中で終わっている
    if fml is not None:
        for j in range(width):
            f_rows = [i for i in body_idx if isinstance(fml[i][j], str) and fml[i][j].startswith('=')]
            c_rows = [i for i in body_idx if not (isinstance(fml[i][j], str) and fml[i][j].startswith('='))
                      and not _is_blank_value(rows[i][j])]
            if len(f_rows) >= 2 and c_rows and len(c_rows) < len(f_rows) and len(c_rows) <= max(1, len(body_idx) // 5):
                ex = " ".join(addr(i, j) for i in c_rows[:_CONTENT_MAX_EXAMPLES])
                blocking.append(f"数式の列「{head(j)}」({_col_letters(c0 + j)}) に直値が {len(c_rows)} セル混ざっています（{ex}）"
                                "→ 隣と同じ式に戻す（人がわざと値にした所なら report に理由を書く）")
            # 式が最後の本文行まで届いていない（足した列が短い＝いちばん多い作り落とし）
            if len(f_rows) >= 2:
                tail = [i for i in body_idx if i > max(f_rows) and _is_blank_value(rows[i][j])]
                if tail:
                    ex = " ".join(addr(i, j) for i in tail[:_CONTENT_MAX_EXAMPLES])
                    blocking.append(f"数式の列「{head(j)}」({_col_letters(c0 + j)}) の式が {addr(max(f_rows), j)} で切れています"
                                    f"（下に本文が {len(tail)} 行あるのに空欄: {ex}）→ 最後の行まで式を入れる")
        # 3.5 突き合わせの結果がほぼ全部同じ（キーの型違い＝VLOOKUP が全部「〜なし」になる形）
        for j in range(width):
            f_rows = [i for i in body_idx if isinstance(fml[i][j], str) and fml[i][j].startswith('=')]
            if len(f_rows) < _SAMEVALUE_MIN_ROWS:
                continue
            texts = [str(rows[i][j]).strip() for i in f_rows if isinstance(rows[i][j], str) and str(rows[i][j]).strip()]
            if len(texts) < len(f_rows) * 0.9 or not texts:
                continue
            top = max(set(texts), key=texts.count)
            if len(top) <= 12 and texts.count(top) >= len(f_rows) * 0.9:
                blocking.append(f"突き合わせの列「{head(j)}」({_col_letters(c0 + j)}) が {texts.count(top)}/{len(f_rows)} 行とも"
                                f"「{top}」です → キーの型違いを疑う（片方が文字の数字・全角・前後の空白）。"
                                "本当に全部そうなら report に「元から一致が無い」と書く")
    # 3.7 集計の式が本文行を取りこぼしている（2026-09-09・数式の目）。
    #     2. の値の突き合わせは「今の数」しか見ない＝取りこぼした行がたまたま空なら気づけない
    #     （次に人がその行を埋めた瞬間に合計が狂う）。ここは範囲そのものを見る。
    if fml is not None:
        sub_rows = {r0 + x for x, kk in enumerate(kinds) if kk == 'sub'}
        for i, k in enumerate(kinds):
            if k not in ('sum', 'sub'):
                continue
            if k == 'sub':
                start = max([x for x in range(i) if kinds[x] in ('sub', 'head')] or [0]) + 1
                above = [x for x in range(start, i) if kinds[x] == 'body']
            else:
                above = [x for x in range(1, i) if kinds[x] == 'body']
            if len(above) < 2:
                continue
            for j in range(width):
                aggs = fml_agg_ranges(fml[i][j])
                if len(aggs) != 1:
                    continue
                func, boxes = aggs[0]
                col = c0 + j
                vert = [b for b in boxes if b[1] == b[3] == col]
                if not vert:
                    continue
                covered = set()
                for b in vert:
                    covered |= set(range(b[0], b[2] + 1))
                if covered & sub_rows:
                    continue                     # 小計を足し上げる合計＝本文を直に足す形ではない
                name = '小計' if k == 'sub' else '合計'
                letters = _col_letters(col)
                missing = [r0 + x for x in above if (r0 + x) not in covered]
                if missing:
                    blocking.append(
                        f"「{name}」行の {addr(i, j)} の {func} が本文 {len(missing)} 行を範囲に入れていません"
                        f"（{' '.join(letters + str(m) for m in missing[:_CONTENT_MAX_EXAMPLES])}）"
                        f"→ 範囲を {letters}{r0 + above[0]}:{letters}{r0 + above[-1]} に直す")
                elif r0 in covered:
                    noticed.append(f"「{name}」行の {addr(i, j)} の {func} が見出し行（{letters}{r0}）まで"
                                   "範囲に入れています")
    # 3.9 文字として入った式（'=' で始まる「文字」＝ Excel は計算しない。人には式に見える）。
    #     表示形式が @ の列に式を入れると起きる。道具自身も一度これをやった（_untext_grid・2026-09-09）
    astext = [addr(i, j) for i in range(len(rows)) for j in range(width)
              if isinstance(rows[i][j], str) and rows[i][j].startswith('=') and len(rows[i][j]) > 1]
    if astext:
        noticed.append(f"式が文字として入っているセルが {len(astext)} 個あります"
                       f"（{' '.join(astext[:_CONTENT_MAX_EXAMPLES])}"
                       f"{'…' if len(astext) > _CONTENT_MAX_EXAMPLES else ''}）。"
                       "Excel は計算していません → 表示形式を標準にしてから入れ直す")
    # 4. 表の中の空欄（同じ列の他の行は埋まっている。まばらな列＝備考などは見ない）
    if body_idx:
        blanks_all = []
        for j in range(width):
            h = head(j)
            if not h or any(k in h.lower() for k in _SPARSE_HEAD_KW):
                continue
            cells = [(i, rows[i][j]) for i in body_idx]
            filled = sum(1 for _i, v in cells if not _is_blank_value(v))
            if filled < 3 or filled / len(cells) < 0.6:
                continue
            blanks_all += [addr(i, j) for i, v in cells if _is_blank_value(v)]
        if blanks_all:
            noticed.append(f"表の中の空欄が {len(blanks_all)} セル（{' '.join(blanks_all[:_CONTENT_MAX_EXAMPLES])}"
                           f"{'…' if len(blanks_all) > _CONTENT_MAX_EXAMPLES else ''}）。同じ列の他の行は埋まっています")
    # 5. 数値列に混ざる文字の数字（番号列は除く）／6. 番号列の重複
    for j in range(width):
        h = head(j)
        if not h:
            continue
        hn = unicodedata.normalize('NFKC', h)
        is_id = _is_id_header(hn)
        col = [(i, rows[i][j]) for i in body_idx if not _is_blank_value(rows[i][j])]
        if not col:
            continue
        if is_id:
            seen, dups = {}, {}
            for i, v in col:
                key = unicodedata.normalize('NFKC', str(v).strip()) if isinstance(v, str) else f"{_num_of(v):g}" if _num_of(v) is not None else str(v)
                if key in seen:
                    dups.setdefault(key, [seen[key]]).append(i)
                else:
                    seen[key] = i
            if dups:
                ex = "／".join(f"{k}: " + " ".join(addr(i, j) for i in v) for k, v in list(dups.items())[:4])
                noticed.append(f"番号列「{h}」({_col_letters(c0 + j)}) に重複が {len(dups)} 組（{ex}）")
            continue
        nums = [i for i, v in col if _num_of(v) is not None]
        texts = [i for i, v in col if isinstance(v, str)
                 and _NUMTEXT_RE.match(unicodedata.normalize('NFKC', v).strip().replace(' ', ''))]
        if len(nums) >= 2 and texts:
            ex = " ".join(f"{addr(i, j)}{str(rows[i][j]).strip()!r}" for i in texts[:_CONTENT_MAX_EXAMPLES])
            noticed.append(f"数値列「{h}」({_col_letters(c0 + j)}) に文字の数字が {len(texts)} セル（{ex}）。"
                           "集計から外れます")
    # 7. 数式の目（2026-09-09）: 列の中で形の違う式・式に直書きされた数値・集計の起点行の食い違い。
    #    どれも「表として壊れている」とは言い切れないので気づき（done は止めない。人が決める）
    if fml is not None:
        fcells = {}
        for i in range(len(rows)):
            for j in range(width):
                f = fml[i][j]
                if isinstance(f, str) and f.startswith('='):
                    rc = None
                    if formulas_r1c1 is not None and i < len(formulas_r1c1) and j < len(formulas_r1c1[i]):
                        rc = formulas_r1c1[i][j]
                    fcells[(r0 + i, c0 + j)] = (f, rc)
        noticed += formula_notes(fcells)
        # 横にコピーできない集計（月や区分を式に埋めた表）は止める（2026-09-09・shu 指示）
        blocking += formula_copy_blockers(fcells)
    return blocking, noticed


def audit_content(ws, rng):
    """表の中身を検査する → (done を止める指摘, 報告の気づき)。COM は 2 回（Value と Formula）。"""
    nrows, ncols = int(rng.Rows.Count), int(rng.Columns.Count)
    if nrows < 2 or nrows * ncols > _AUDIT_MAX_CELLS * 4:
        return [], []
    vals = rng.Value
    if nrows == 1:
        vals = [list(vals)] if isinstance(vals, tuple) else [[vals]]
    elif ncols == 1:
        vals = [[r[0]] if isinstance(r, tuple) else [r] for r in vals]
    else:
        vals = [list(r) for r in vals]
    def _grid(src):
        if ncols == 1:
            return [[r[0]] if isinstance(r, tuple) else [r] for r in src]
        return [list(r) for r in src]

    fmls = None
    try:
        fmls = _grid(rng.Formula)
    except Exception:
        fmls = None
    fmls_rc = None
    try:
        fmls_rc = _grid(rng.FormulaR1C1)          # 式の形を比べるため（COM は 1 回・2026-09-09）
    except Exception:
        fmls_rc = None
    return _content_findings(vals, fmls, int(rng.Row), int(rng.Column), fmls_rc)


@protect_safe
@dialog_safe
def cmd_audit(args):
    """表の見た目と中身を検査する（仕上げ検査）: audit <範囲|セル> [<範囲|セル> ...]

    見た目（見出し・罫線・列の型・列幅・空の見出し）と、中身のうち依頼の知識が要らないもの
    （エラー値・「合計」行が上の和と合わない・数式の列に直値が混ざる・集計の式が本文の行を範囲に
    入れていない）を見る。指摘が 0 件なら合格＝終了コード 0。
    表の中の空欄・数値列の文字の数字・番号列の重複・列の中で式の形が違うセル・式に直書きされた数値は
    「気づき」として出す（合否には入れない＝人が決める）。
    agent はこれを done の関所として使う（2026-09-04・「値は合ったが見た目がひどい」を道具で止める。
    2026-09-06 夜に中身を足した＝初実射で空欄を誰も言わなかった件）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: audit [excel_file] <範囲|セル> [<範囲|セル> ...]")
        print("  セル1つなら、それを含む表（CurrentRegion）全体を検査する")
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    total = 0
    for spec in rest:
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
        if int(rng.Cells.CountLarge) == 1:
            rng = _region_of_cell(ws, rng)
        addr = str(rng.Address).replace('$', '')
        notes = audit_table(ws, rng)
        try:
            c_block, c_seen = audit_content(ws, rng)
        except Exception as ex:
            c_block, c_seen = [], [f"（中身の検査を読めませんでした: {ex}）"]
        notes = notes + c_block
        total += len(notes)
        print(f"検査: {ws.Name}!{addr}  " + (f"指摘 {len(notes)} 件" if notes else "指摘なし（合格）"))
        for n in notes:
            print(f"  - {n}")
        for n in c_seen:
            print(f"  気づき: {n}")
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    if total:
        print(f"合計 {total} 件の指摘。tidy を表全体に当ててから、もう一度 audit してください。")
    return total == 0


def _trim_trailing_empty(ws, rng):
    """範囲の右端・下端の、値が 1 つも無い列・行を外す → (範囲, 外したものの説明か '')。

    2026-09-08: AI が tidy の範囲を使用範囲（A1:F20）に合わせて A5:F19 と返し、空の F 列に格子罫線と見出しの
    書式が付いた。「データの末尾まで」と規則文に書いても sonnet は使用範囲を写す＝機械的な判断は道具が持つ。
    全部空なら触らない（別の所で「表の形をしていません」と断られる）。
    """
    try:
        vals = rng.Value
    except Exception:
        return rng, ''
    if not isinstance(vals, tuple):
        return rng, ''
    rows = [list(r) if isinstance(r, tuple) else [r] for r in vals]
    nr, nc = len(rows), max(len(r) for r in rows)

    def filled(v):
        return v is not None and not (isinstance(v, str) and v.strip() == '')

    last_r = max((i for i, r in enumerate(rows) if any(filled(v) for v in r)), default=-1)
    last_c = max((j for j in range(nc) if any(filled(r[j]) for r in rows if j < len(r))), default=-1)
    if last_r < 0 or last_c < 0 or (last_r == nr - 1 and last_c == nc - 1):
        return rng, ''
    what = []
    if last_c < nc - 1:
        what.append(f"{nc - 1 - last_c} 列")
    if last_r < nr - 1:
        what.append(f"{nr - 1 - last_r} 行")
    try:
        return ws.Range(rng.Cells(1, 1), rng.Cells(last_r + 1, last_c + 1)), "・".join(what)
    except Exception:
        return rng, ''


@protect_safe
@dialog_safe
def _header_mixed(header):
    """見出し行の書式がセルごとにばらばらか（太字・字・大きさ・色・塗りのどれかが混在＝COM は None を返す）。"""
    try:
        f = header.Font
        return (f.Bold is None or f.Name is None or f.Size is None or f.Color is None
                or f.Italic is None or header.Interior.ColorIndex is None)
    except Exception:
        return False


def _border_blocks(rows):
    """tidy の罫線を引く範囲 → ([(行1, 列1, 行2, 列2), …]（範囲の中の 1 始まり）, 説明)。

    2026-09-13（お試し版テスト用1）: 表全体に格子を引いていたので、明細と合計のあいだの空の行と、
    合計欄（F27:G28）の左の空欄にまで罫線が付いた。空の行で切れた後ろの塊が表より細い（合計・平均など）
    なら、その塊は埋まっている列だけに引き、あいだの空の行には引かない。表と同じ幅の塊は表の続き
    （抜けた 1 件の空行）とみなして、空行ごと格子に入れる。純 Python（COM を呼ばない＝テストできる）。
    """
    nr = len(rows)
    nc = max((len(r) for r in rows), default=0)
    if not nr or not nc:
        return [(1, 1, max(nr, 1), max(nc, 1))], ''

    def filled(v):
        return v is not None and not (isinstance(v, str) and v.strip() == '')

    runs, i = [], 0
    while i < nr:
        if any(filled(v) for v in rows[i]):
            j = i
            while j + 1 < nr and any(filled(v) for v in rows[j + 1]):
                j += 1
            runs.append((i, j))
            i = j + 1
        else:
            i += 1
    if not runs:
        return [(1, 1, nr, nc)], ''
    blocks = [[0, 0, runs[0][1], nc - 1]]
    for a, b in runs[1:]:
        cols = [c for c in range(nc) if any(c < len(rows[r]) and filled(rows[r][c]) for r in range(a, b + 1))]
        if cols[0] == 0 and cols[-1] == nc - 1:
            blocks[-1][2] = b                  # 表と同じ幅＝表の続き（途中の空行ごと格子に入れる）
        else:
            blocks.append([a, cols[0], b, cols[-1]])
    out = [(r1 + 1, c1 + 1, r2 + 1, c2 + 1) for r1, c1, r2, c2 in blocks]
    if out == [(1, 1, nr, nc)]:
        return out, ''
    return out, '空の行と、下の集計欄の外の空欄には引かない'


def cmd_tidy(args):
    """表を整える（仕上げ）: tidy <範囲|セル> [<範囲|セル> ...] [--header-from セル] [--bg 色]
                              [--no-header] [--no-border] [--no-col-format] [--no-autofit]
                              [--min-width N] [--max-width N]

    値を書いた後の「見た目の仕上げ」を1コマンドにまとめる（2026-09-02・Excelコンボに
    見た目で負けた日に輸入）。見出し行を既存の見出しと同じ書式に／罫線（細・格子）／
    列の型に合わせた寄せ・表示形式（番号列は左寄せ・数値列は #,##0・小数列は #,##0.00）／
    列幅の自動調整（下限・上限つき）を当て、最後に必ず見え方を読み戻す（### の検査つき）。
    セル1つを渡したときは、そのセルを含む表（CurrentRegion）全体が対象。
    範囲は複数並べられる（表が2つでも1往復）。--header-from を省いたときは、各表の左上が
    既存の見出しの体裁（太字か塗りつぶし）ならその書式を見出し行全体へ複写し、そうでなければ
    太字＋既定の背景色にする。終わりに materials からの経過秒を出す（同日夜の改修）。
    列の型（同日深夜・「番号は左寄せ・金額はカンマが当然」）は _column_style が決め、既に書式のある
    列は上書きしない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: tidy [excel_file] <範囲|セル> [<範囲|セル> ...] [--header-from セル] [--bg '#RRGGBB'] "
              "[--no-header] [--no-border] [--no-col-format] [--no-autofit] [--min-width N] [--max-width N]")
        print("  セル1つなら、それを含む表（CurrentRegion）全体を整える。終わりに見え方を読み戻す")
        print("  列の型: 見出しが番号/コード/No/ID の列は左寄せ、数値だけの列は #,##0（年・日付・率の列は触らない）")
        return False
    try:
        min_w = float(getattr(args, 'min_width', None) or 6)
        max_w = float(getattr(args, 'max_width', None) or 60)
    except (TypeError, ValueError):
        print("エラー: --min-width / --max-width は数値で指定してください")
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    # 先に全部解決してから触る（2つ目でエラーなら、1つ目だけ整った中途半端を残さない）
    targets = []
    for spec in rest:
        whole = _whole_sheet_spec(wb, spec, sheet_opt)
        if whole is not None:
            print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。"
                  f"表の範囲か、表の中のセル1つを指定してください。")
            return False
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
        if int(rng.Cells.CountLarge) == 1:
            rng = _region_of_cell(ws, rng)             # セル1つ → それを含む表（表題は含めない）
        rng, trimmed = _trim_trailing_empty(ws, rng)   # 右端・下端の空の列・行は表ではない（空の列に罫線が付く・9/8）
        if trimmed:
            print(f"範囲の右端・下端の空の {trimmed} を外しました → {ws.Name}!{rng.Address.replace('$', '')}")
        if int(rng.Rows.Count) < 2:
            print(f"エラー: {ws.Name}!{rng.Address.replace('$', '')} は表の形をしていません"
                  f"（見出し＋1行以上の範囲を指定してください）")
            return False
        targets.append((ws, rng))
    src_opt = getattr(args, 'header_from', None)
    for ws, _rng in targets:
        _unfilter_for_write(ws)                        # 絞り込みが掛かっていると書式が隠れた行を飛ばす
    for ws, rng in targets:
        applied = []
        ncols = int(rng.Columns.Count)
        header = ws.Range(rng.Cells(1, 1), rng.Cells(1, ncols))
        if not getattr(args, 'no_header', False):
            src_ws, src_cell = None, None
            if src_opt:
                src_ws, rng_s = _resolve_range(xl, wb, src_opt, sheet_opt)
                src_cell = rng_s.Cells(1, 1)
            elif _looks_like_header(rng.Cells(1, 1)) and not _header_mixed(header):
                src_ws, src_cell = ws, rng.Cells(1, 1)   # 表の左上＝既存の見出し。その書式を行全体へ
            if src_cell is not None:
                # 先頭行に数値・日付があるなら、その表示形式は複写で潰さない（2026-09-08）。
                # xlPasteFormats は表示形式まで複写するので、合計ブロック（左が「合計」・右が 203,250）に
                # 当てると #,##0 が General に戻り、カンマが消えた見た目になる。年を見出しにした表
                # （2025・2026）でも同じ。見出しらしさ（太字・塗り・罫線・寄せ）だけを複写する。
                keep_fmt = _numeric_formats(header)
                try:
                    src_cell.Copy()
                    header.PasteSpecial(-4122)        # xlPasteFormats
                finally:
                    try:
                        xl.CutCopyMode = False
                    except Exception:
                        pass
                kept = _restore_formats(header, keep_fmt)
                applied.append(f"見出し=書式を{src_ws.Name}!{src_cell.Address.replace('$', '')}から複写"
                               + (f"（数値{kept}セルの表示形式は残した）" if kept else ""))
            else:
                bg = getattr(args, 'bg', None) or '#DDEBF7'
                mixed = _header_mixed(header)
                if mixed:
                    # 見出しの書式がセルごとにばらばら（2026-09-11: 左上 1 セルの書式を写して太字でない見出しになり、
                    # 採点係に「見出しは太字に」と差し戻された）。飾りを外し、字は本文にそろえて標準の見出しにする
                    f = header.Font
                    f.Italic = False
                    f.Underline = -4142
                    f.Strikethrough = False
                    f.Color = 0
                    try:
                        body = rng.Cells(2, 1).Font
                        if body.Name:
                            f.Name = body.Name
                        if body.Size:
                            f.Size = body.Size
                    except Exception:
                        pass
                header.Font.Bold = True
                header.Interior.Color = _hex_to_excel_color(bg)
                applied.append(f"見出し=太字・背景{bg}" + ("（元の見出しの書式がセルごとにばらばらだったので標準にそろえた）"
                                                           if mixed else ""))
        if not getattr(args, 'no_border', False):
            # 空の行と、下の細い塊（合計・平均）の外の空欄には引かない（2026-09-13・_border_blocks）
            try:
                vals = rng.Value
                grid = ([list(r) if isinstance(r, tuple) else [r] for r in vals]
                        if isinstance(vals, tuple) else None)
            except Exception:
                grid = None
            blocks, bnote = _border_blocks(grid) if grid else ([], '')
            if not blocks or blocks == [(1, 1, int(rng.Rows.Count), ncols)]:
                areas = [rng]
            else:
                areas = [ws.Range(rng.Cells(r1, c1), rng.Cells(r2, c2)) for r1, c1, r2, c2 in blocks]
            for area in areas:
                area.Borders.LineStyle = 1             # xlContinuous（格子）
                area.Borders.Weight = _XL_BORDER_WEIGHT['thin']
            applied.append("罫線=細・格子" + (f"（{bnote}）" if bnote else ""))
        if not getattr(args, 'no_col_format', False):
            # 列幅の自動調整より先（12000→12,000 で幅が変わる）
            notes = _tidy_column_styles(ws, rng)
            applied.append("列の型=" + ("／".join(notes) if notes else "変更なし"))
        if not getattr(args, 'no_autofit', False):
            # 幅は列全体に効く＝表の外（上・下）のセルまで細くなる。細くしすぎて範囲外を '###' に
            # したら元の幅へ戻す（2026-09-08・前の走行の tidy が D 列を 8→6 にして、表の下の
            # D31 が ##### になっていた。ここの ### 検査は範囲の中しか見ていないので誰も気づかなかった）
            before_w = {}
            for j in range(1, ncols + 1):
                try:
                    before_w[j] = float(rng.Columns(j).ColumnWidth)
                except Exception:
                    continue
            rng.Columns.AutoFit()
            clamped = 0
            for j in range(1, ncols + 1):
                col = rng.Columns(j)
                try:
                    w = float(col.ColumnWidth)
                except Exception:
                    continue
                if w < min_w:
                    col.ColumnWidth = min_w
                    clamped += 1
                elif w > max_w:
                    col.ColumnWidth = max_w
                    clamped += 1
            restored = _undo_narrowing_that_broke_outside(ws, rng, before_w)
            applied.append("列幅=自動" + (f"（{clamped}列を{min_w:g}〜{max_w:g}に丸め）" if clamped else "")
                           + (f"（{'・'.join(restored)}は表の外が ### になるので元の幅に戻した）" if restored else ""))
        print(f"整えました: {ws.Name}!{rng.Address.replace('$', '')}  [{', '.join(applied)}]")
        _show_range(ws, rng, label="仕上がり")
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_sheet(args):
    """シート操作: add/delete/rename/copy/activate/show/hide"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_sheet(name):
        for sh in wb.Sheets:
            if sh.Name == name:
                return sh
        return None

    if action == 'add':
        new_name = rest[1] if len(rest) >= 2 else None
        before = getattr(args, 'before', None)
        after = getattr(args, 'after', None)
        # --before/--after の対象が実在しないと find_sheet が None になり、
        # 無言でアクティブシート手前に追加されてしまうため先に検証する
        if before and find_sheet(before) is None:
            print(f"エラー: --before のシート '{before}' が見つかりません"); return False
        if after and find_sheet(after) is None:
            print(f"エラー: --after のシート '{after}' が見つかりません"); return False
        # 名前重複は Add 後の rename で例外→既定名シートの残骸になるため先に検証。
        # ただし「もうある」は失敗にしない（欲しい終わりの姿＝その名前のシートがある、は既に満たされている）。
        # 失敗にしていたので、同じ返事に並んだ後ろの手が全部取り消され、往復を 1 回まるごと失っていた
        # （2026-09-06 深夜の実射・「月別」の弾）
        if new_name and find_sheet(new_name) is not None:
            print(f"シート '{new_name}' は既にあります（作らずにそのまま使います）")
            return True
        if before:
            sh = wb.Sheets.Add(find_sheet(before))
        elif after:
            sh = wb.Sheets.Add(None, find_sheet(after))
        else:
            sh = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
        if new_name:
            sh.Name = new_name
        print(f"シート追加: {sh.Name}")
    elif action == 'delete':
        if len(rest) < 2:
            print("使い方: sheet delete <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        with _alerts_off(xl):
            sh.Delete()
        print(f"シート削除: {rest[1]}")
    elif action == 'rename':
        if len(rest) < 3:
            print("使い方: sheet rename <old> <new>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Name = rest[2]
        print(f"シート名変更: {rest[1]} → {rest[2]}")
    elif action == 'copy':
        if len(rest) < 2:
            print("使い方: sheet copy <name> [newname]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        # 名前重複は Copy 後の rename で例外になり、「Sheet1 (2)」等の複製シートが
        # 残骸として残る（add と同じく先に弾く）
        if len(rest) >= 3 and find_sheet(rest[2]) is not None:
            print(f"エラー: シート '{rest[2]}' は既に存在します"); return False
        sh.Copy(None, sh)
        newsh = wb.ActiveSheet
        if len(rest) >= 3:
            newsh.Name = rest[2]
        print(f"シート複製: {rest[1]} → {newsh.Name}")
    elif action == 'activate':
        if len(rest) < 2:
            print("使い方: sheet activate <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Activate()
        print(f"アクティブ化: {rest[1]}")
    elif action in ('protect', 'unprotect'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name> [--password パスワード]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        pw = getattr(args, 'password', None)
        if action == 'protect':
            sh.Protect(Password=pw) if pw else sh.Protect()
            print(f"シート保護: {rest[1]}" + ("（パスワード付き）" if pw else ""))
            print("  ※ format-range --lock/--unlock で設定した Locked がここで効きます")
        else:
            try:
                sh.Unprotect(Password=pw) if pw else sh.Unprotect()
            except Exception as e:
                print(f"エラー: 保護解除に失敗しました（パスワード違いの可能性）: {e}")
                return False
            # 保護ガード（protect_safe）は入口で全保護シートを一時解除し、出口で
            # 記録どおり再保護する。このコマンドは「保護を外すこと自体が目的」なので、
            # 記録から落としておかないと出口で保護が戻り、成功表示のまま無言で効かない
            forget_protection(sh)
            print(f"シート保護解除: {rest[1]}")
    elif action in ('show', 'hide', 'very-hide'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        vis = {'show': -1, 'hide': 0, 'very-hide': 2}[action]   # xlVisible/-Hidden/-VeryHidden
        sh.Visible = vis
        label = {'show': '表示', 'hide': '非表示', 'very-hide': '完全非表示(VBAのみ解除可)'}[action]
        print(f"シート{label}: {rest[1]}")
    elif action == 'visibility':
        if len(rest) < 2:
            print("使い方: sheet visibility <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        label = {-1: 'visible（表示）', 0: 'hidden（非表示）',
                 2: 'veryhidden（完全非表示）'}.get(int(sh.Visible), str(sh.Visible))
        print(f"表示状態: {rest[1]} = {label}")
        return True
    elif action == 'tab-color':
        if len(rest) < 2:
            print("使い方: sheet tab-color <name> [#RRGGBB | R G B | --clear]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        if getattr(args, 'clear', False):
            sh.Tab.ColorIndex = -4142             # xlColorIndexNone
            print(f"タブ色クリア: {rest[1]}")
        elif len(rest) >= 5 and all(x.isdigit() for x in rest[2:5]):
            r, g, b = int(rest[2]), int(rest[3]), int(rest[4])
            sh.Tab.Color = r + g * 256 + b * 65536
            print(f"タブ色設定: {rest[1]} = RGB({r},{g},{b})")
        elif len(rest) >= 3:
            sh.Tab.Color = _hex_to_excel_color(rest[2])
            print(f"タブ色設定: {rest[1]} = {rest[2]}")
        else:
            if int(sh.Tab.ColorIndex) == -4142:
                print(f"タブ色: {rest[1]} = （未設定）")
            else:
                c = int(sh.Tab.Color)
                r = c & 255; g = (c >> 8) & 255; b = (c >> 16) & 255
                print(f"タブ色: {rest[1]} = #{r:02X}{g:02X}{b:02X} (R={r},G={g},B={b})")
            return True
    else:
        print(f"未知のアクション: {action}")
        return False

    print("（保存はしていません）")
    return True


@protect_safe
def cmd_table(args):
    """テーブル(ListObject)操作: create/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: table [excel_file] <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                cnt += 1
                print(f"[{sh.Name}] {lo.Name}  範囲={lo.Range.Address}")
        if cnt == 0:
            print("テーブルはありません。")
        return True

    if action == 'create':
        if len(rest) < 2:
            print("使い方: table create <range> [name] [--no-headers]"); return False
        ws, rng = _resolve_range(xl, wb, rest[1])
        has_headers = 2 if getattr(args, 'no_headers', False) else 1  # xlNo=2 / xlYes=1
        # シートにオートフィルタ（▼）が付いていると Excel はテーブルを作れない（「このタスクでは
        # ワークシート上のフィルターされた範囲が変更されます」）。ここで断ると AI は Excel の生の
        # 文言だけを見て往復を使い切る（2026-09-09 の実射「スライサー」×状態「絞り込み中」＝
        # 4 往復すべてを使って「不可」で終わった）。絞り込みは見せ方であってデータではないので、
        # 書き込みのとき（_unfilter_for_write）と同じ流儀＝外して、外した範囲を報告する。
        try:
            if bool(ws.AutoFilterMode):
                gone = str(ws.AutoFilter.Range.Address).replace('$', '')
                if ws.FilterMode:
                    ws.ShowAllData()
                ws.AutoFilterMode = False
                print(f"（オートフィルタの▼を外しました: {ws.Name}!{gone}＝Excel はフィルタのある"
                      "シートにテーブルを作れません。テーブルには自前の▼が付きます。"
                      "外した範囲に▼が要るなら autofilter で付け直してください）")
        except Exception:
            pass
        lo = ws.ListObjects.Add(1, rng, None, has_headers)             # xlSrcRange=1
        if len(rest) >= 3:
            lo.Name = rest[2]
        print(f"テーブル作成: [{ws.Name}] {lo.Name}  範囲={lo.Range.Address}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: table delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    lo.Unlist()                    # テーブル解除 (データは残す)
                    print(f"テーブル解除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: テーブル '{name}' が見つかりません")
        return False

    # ---- 以降は <table名> を rest[1] に取る列・フィルタ・ソート操作 ----
    def _find_lo(name):
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    return sh, lo
        return None, None

    def _col_field(lo, col_name):
        """テーブル内の列番号(1始まり)を名前から得る。無ければ None。"""
        for i in range(1, lo.ListColumns.Count + 1):
            if lo.ListColumns.Item(i).Name == col_name:
                return i
        return None

    if action == 'column':
        # table column <add|remove|rename|format> <table> ...
        sub = rest[1].lower() if len(rest) >= 2 else ''
        tname = rest[2] if len(rest) >= 3 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません（table list で確認）。"); return False
        if sub == 'add':
            col_name = rest[3] if len(rest) >= 4 else None
            pos = getattr(args, 'at', None)
            lc = lo.ListColumns.Add(int(pos)) if pos else lo.ListColumns.Add()
            if col_name:
                lc.Name = col_name
            print(f"列追加: {tname}[{lc.Name}]（位置 {lc.Index}）")
            print("（保存はしていません）"); return True
        if sub == 'remove':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lo.ListColumns.Item(col_name).Delete()
            print(f"列削除: {tname}[{col_name}]"); print("（保存はしていません）"); return True
        if sub == 'rename':
            if len(rest) < 5:
                print("使い方: table column rename <table> <旧列> <新列>"); return False
            old, new = rest[3], rest[4]
            if _col_field(lo, old) is None:
                print(f"エラー: 列 '{old}' が見つかりません。"); return False
            lo.ListColumns.Item(old).Name = new
            print(f"列名変更: {tname}[{old}] → [{new}]"); print("（保存はしていません）"); return True
        if sub == 'format':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lc = lo.ListColumns.Item(col_name)
            if len(rest) >= 5:                       # set
                lc.DataBodyRange.NumberFormat = rest[4]
                print(f"列書式設定: {tname}[{col_name}] = {rest[4]}")
                print("（保存はしていません）"); return True
            else:                                    # get
                try:
                    fmt = lc.DataBodyRange.Cells(1, 1).NumberFormat
                except Exception:
                    fmt = '(取得不可)'
                print(f"列書式: {tname}[{col_name}] = {fmt}"); return True
        print("使い方: table column <add|remove|rename|format> <table> ...")
        return False

    if action in ('filter', 'filter-values', 'filter-clear', 'filters'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        if action == 'filter-clear':
            try:
                if lo.AutoFilter is not None:
                    lo.AutoFilter.ShowAllData()
                print(f"フィルタ解除: {tname}")
            except Exception:
                print(f"フィルタは設定されていません: {tname}")
            print("（保存はしていません）"); return True
        if action == 'filters':
            af = lo.AutoFilter
            if af is None:
                print(f"フィルタなし: {tname}"); return True
            print(f"--- {tname} のフィルタ ---")
            any_on = False
            for i in range(1, lo.ListColumns.Count + 1):
                fl = af.Filters.Item(i)
                try:
                    on = fl.On
                except Exception:
                    on = False
                if on:
                    any_on = True
                    try:
                        c1 = fl.Criteria1
                    except Exception:
                        c1 = '(?)'
                    print(f"  {lo.ListColumns.Item(i).Name}: {c1}")
            if not any_on:
                print("  (フィルタ条件なし)")
            return True
        # filter / filter-values は列指定が必要
        col_name = rest[2] if len(rest) >= 3 else None
        field = _col_field(lo, col_name)
        if field is None:
            print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
        if action == 'filter':
            crit = rest[3] if len(rest) >= 4 else None
            if not crit:
                print("使い方: table filter <table> <列> <条件>（例: \">100\" \"=Active\"）"); return False
            lo.Range.AutoFilter(Field=field, Criteria1=crit)
            print(f"フィルタ適用: {tname}[{col_name}] {crit}")
            print("（保存はしていません）"); return True
        if action == 'filter-values':
            vals = rest[3:]
            if not vals:
                print("使い方: table filter-values <table> <列> 値1 値2 ..."); return False
            lo.Range.AutoFilter(Field=field, Criteria1=list(vals), Operator=7)   # xlFilterValues
            print(f"フィルタ適用(値): {tname}[{col_name}] {vals}")
            print("（保存はしていません）"); return True

    if action in ('sort', 'sort-multi'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        so = lo.Sort
        so.SortFields.Clear()
        if action == 'sort':
            col_name = rest[2] if len(rest) >= 3 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            order = 2 if getattr(args, 'desc', False) else 1   # xlDescending/xlAscending
            so.SortFields.Add(lo.ListColumns.Item(col_name).Range, 0, order)
            so.Apply()
            print(f"ソート: {tname} {col_name} {'降順' if order == 2 else '昇順'}")
            print("（保存はしていません）"); return True
        else:  # sort-multi  col:asc col:desc ...
            specs = rest[2:]
            if not specs:
                print("使い方: table sort-multi <table> 列:asc 列:desc ..."); return False
            applied = []
            for spec in specs:
                if ':' in spec:
                    cn, od = spec.rsplit(':', 1)
                else:
                    cn, od = spec, 'asc'
                if _col_field(lo, cn) is None:
                    print(f"エラー: 列 '{cn}' が見つかりません。"); return False
                order = 2 if od.lower().startswith('d') else 1
                so.SortFields.Add(lo.ListColumns.Item(cn).Range, 0, order)
                applied.append(f"{cn}{'↓' if order == 2 else '↑'}")
            so.Apply()
            print(f"複数ソート: {tname} {' / '.join(applied)}")
            print("（保存はしていません）"); return True

    if action == 'read':
        # テーブル名で直接読む（従来は table list で番地を得て read-range する2段＝2接続だった）
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。（table list で確認）"); return False
        rng = lo.Range
        print(f"テーブル: {tname}   [{sh.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng))
        print("=" * 60)
        tsv_out = getattr(args, 'tsv_out', None)
        if tsv_out is not None:
            path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
            rows = _range_values_2d(rng)
            # セル内改行(Alt+Enter)・タブは TSV の行/列区切りと衝突する。
            # 検査せずに書き戻しコマンドまで案内すると、改行1つで TSV の行数が増え、
            # write-range の「不揃い」経路に落ちて以降の全行が1行ずつ下へズレて
            # 上書きされる（＝無警告のデータ破壊）。read-range 側と同じ警告を出す
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
            with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                f.write('\n'.join('\t'.join(_cell_str(v) for v in r) for r in rows) + '\n')
            print(f"TSV書き出し: {path}  ({len(rows)}行)")
            print(f"  編集後の書き戻し: py vba_manager.py write-range \"{sh.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
            print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
        return True

    if action == 'ref':
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        col_name = rest[2] if len(rest) >= 3 else None
        if col_name:
            print(f"構造化参照: {tname}[{col_name}]")
        else:
            cols = [lo.ListColumns.Item(i).Name for i in range(1, lo.ListColumns.Count + 1)]
            print(f"構造化参照: {tname}[#All] / 列: {', '.join(cols)}")
        return True

    print(f"未知のアクション: {action}")
    return False


@protect_safe
def cmd_name(args):
    """名前付き範囲操作: add/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: name [excel_file] <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for nm in wb.Names:
            cnt += 1
            try:
                refers = nm.RefersTo
            except Exception:
                refers = '(?)'
            print(f"{nm.Name}  →  {refers}")
        if cnt == 0:
            print("名前付き範囲はありません。")
        return True

    if action == 'add':
        if len(rest) < 3:
            print("使い方: name add <name> <range>"); return False
        nm_name = rest[1]
        ws, rng = _resolve_range(xl, wb, rest[2])
        # rng.Address は既定で絶対参照 ($A$2)。pywin32 ではプロパティなので引数なしで使う
        refers = "='" + ws.Name.replace("'", "''") + "'!" + rng.Address
        wb.Names.Add(nm_name, refers)
        print(f"名前付き範囲を追加: {nm_name} → {refers}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: name delete <name>"); return False
        nm_name = rest[1]
        # 完全一致を優先。シートスコープ名（'Sheet1'!名前）の末尾一致は
        # 同名が複数シートにあると最初の1個を消す取り違えになるため、
        # 複数一致ならエラーで止めて候補を出す。
        exact = [nm for nm in wb.Names if nm.Name == nm_name]
        if exact:
            exact[0].Delete()
            print(f"名前付き範囲を削除: {nm_name}")
            print("（保存はしていません）")
            return True
        suffix = [nm for nm in wb.Names if nm.Name.split('!')[-1] == nm_name]
        if len(suffix) == 1:
            actual = suffix[0].Name
            suffix[0].Delete()
            print(f"名前付き範囲を削除: {actual}")
            print("（保存はしていません）")
            return True
        if len(suffix) > 1:
            print(f"エラー: 名前 '{nm_name}' はシート違いで複数あります。完全名で指定してください:")
            for nm in suffix:
                print(f"  {nm.Name}")
            return False
        print(f"エラー: 名前 '{nm_name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 「手」コマンド 第2弾 (編集の足回り / 検索置換 / 保存印刷 / 仕上げ)
#   ※ いずれもアクティブ(開いたまま)のブックに COM で直接作用。
#      save 系を除き既定では保存しない。
# ================================================================

def _col_num(s):
    """列文字(A,B,..,AA) を列番号(1始まり)に変換"""
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


# ---- a. 編集の足回り ----

def _sheet_or_active(wb, sheet_name):
    """--sheet 指定があればそのシート、なければアクティブシートを返す。
    指定シートが見つからなければエラー表示して None。"""
    if not sheet_name:
        return wb.ActiveSheet
    for sh in wb.Worksheets:
        if sh.Name == sheet_name:
            return sh
    print(f"エラー: シート '{sheet_name}' が見つかりません")
    return None


@protect_safe
@dialog_safe
def cmd_row(args):
    """行の挿入・削除: row <insert|delete> <行番号> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: row <insert|delete> <行番号> [本数]")
        return False
    action = rest[0].lower()
    start = int(rest[1])
    count = int(rest[2]) if len(rest) >= 3 else 1
    # 本数・行番号の下限検査。count=0 だと Rows("10:9") のような逆順アドレスになり、
    # Excel がこれを正規化するため delete では 9〜10 行＝隣の行まで消える
    # （バッチや計算値から 0 が渡ると無警告でデータが飛ぶ）
    if count < 1:
        print(f"エラー: 本数は1以上で指定してください: {count}")
        return False
    if start < 1:
        print(f"エラー: 行番号は1以上で指定してください: {start}")
        return False
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Rows(f"{start}:{start + count - 1}")
    if action == 'insert':
        rng.Insert()
        print(f"行挿入: {ws.Name} {start}行目に {count}行")
    elif action == 'delete':
        rng.Delete()
        print(f"行削除: {ws.Name} {start}〜{start + count - 1}行")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_col(args):
    """列の挿入・削除: col <insert|delete> <列文字> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: col <insert|delete> <列文字> [本数]")
        return False
    action = rest[0].lower()
    start = rest[1]
    # _col_num は英字以外も黙って数値化する。str.isalpha() は '名' のような日本語も
    # True になるため isascii() と併せて弾く（通すと巨大な列番号になり、
    # delete では見当違いの列が消える＝取り返しがつかない）
    if not (start.isascii() and start.isalpha()):
        print(f"エラー: 列は列文字（A〜XFD）で指定してください: '{start}'")
        return False
    count = int(rest[2]) if len(rest) >= 3 else 1
    # 本数の下限検査。count=0 だと Columns("C:B") のような逆順アドレスになり、
    # Excel がこれを正規化するため delete では B〜C 列＝隣の列まで消える
    if count < 1:
        print(f"エラー: 本数は1以上で指定してください: {count}")
        return False
    end = _col_letter(_col_num(start) + count - 1)
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Columns(f"{start}:{end}")
    if action == 'insert':
        rng.Insert()
        print(f"列挿入: {ws.Name} {start}列に {count}列")
    elif action == 'delete':
        rng.Delete()
        print(f"列削除: {ws.Name} {start}〜{end}列")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_copy_range(args):
    """範囲コピー: copy-range <src> <dst> [--values]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: copy-range <src> <dst> [--values]")
        return False
    if _reject_extra_args(rest, 2, '使い方: copy-range <src> <dst> [--values]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー元 '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が複写先を上書きする危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    # コピー先も同じガード。シート名だけ渡すと UsedRange 全域が貼り付け先になり、
    # コピー元が単一セル等だと全域タイル上書きになる（2026-07-09 再点検で発見）
    whole_d = _whole_sheet_spec(wb, rest[1])
    if whole_d is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー先 '{rest[1]}' はシート '{whole_d}' の使用範囲全域を指します。")
        print(f"  全域がタイル状に上書きされる危険があるため、貼り付け先の左上セルを明示するか --whole-sheet を付けてください。")
        return False
    ws_s, rng_s = _resolve_range(xl, wb, rest[0], sheet_opt)
    ws_d, rng_d = _resolve_range(xl, wb, rest[1])
    if getattr(args, 'values', False):
        try:
            rng_s.Copy()
            rng_d.PasteSpecial(-4163)          # xlPasteValues
        finally:
            # 失敗しても Excel にコピーの点線（CutCopyMode）を残さない
            try:
                xl.CutCopyMode = False
            except Exception:
                pass
        print(f"コピー(値のみ): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    else:
        rng_s.Copy(rng_d)
        print(f"コピー(書式・式込): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    if getattr(args, 'show', False):
        # 貼り付け先が左上セル1つでも、実際に埋まった広さ（コピー元の大きさ）を見せる
        r0, c0 = int(rng_d.Row), int(rng_d.Column)
        nr = max(int(rng_s.Rows.Count), int(rng_d.Rows.Count))
        nc = max(int(rng_s.Columns.Count), int(rng_d.Columns.Count))
        _show_range(ws_d, ws_d.Range(ws_d.Cells(r0, c0), ws_d.Cells(r0 + nr - 1, c0 + nc - 1)))
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_fill(args):
    """オートフィル: fill <range> [--right]（既定は下方向）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: fill <range> [--right]   範囲の先頭セルを残りに複写")
        return False
    if _reject_extra_args(rest, 1, '使い方: fill <range> [--right]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  先頭行/列が全域に複写される危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    _unfilter_for_write(ws)
    if getattr(args, 'right', False):
        rng.FillRight(); direction = "右"
    else:
        rng.FillDown(); direction = "下"
    print(f"フィル({direction}): {ws.Name}!{rng.Address}")
    if getattr(args, 'show', False):
        _show_range(ws, rng)
    print("（保存はしていません）")
    return True


_SORT_HIDDEN_SHOW = 8            # 報告に並べる行番号の数


def _unhide_for_sort(ws, rng):
    """並べ替える前に、範囲の中の隠れた行を出す。出したことを言う文の一覧を返す（2026-09-08）。

    Excel の並べ替えは**見えている行しか動かさない**。隠れた行はその場に取り残されるので、
    並び順が黙って狂う（実射「多い順に並べて」＝手で隠した 2 行が末尾に残ったまま合格していた）。
    絞り込みも手で隠した行も、どちらも「見せ方」であってデータではないので、書く側が出してよい。
    ただし黙って出さない＝報告に書く（_unfilter_for_write と同じ流儀）。
    並べ替えた後に隠し直すことはしない（行が動いた後の行番号は、別の人のデータを指す）。
    """
    out = []
    try:
        if bool(ws.FilterMode):
            ws.ShowAllData()
            out.append(f"※ 絞り込みを解いてから並べました（{ws.Name}）"
                       f"。絞り込んだままだと、隠れた行が動かず並びが狂います")
    except Exception:
        pass
    hidden = []
    try:
        r1 = int(rng.Row)
        for i in range(r1, r1 + int(rng.Rows.Count)):
            if bool(ws.Rows(i).Hidden):
                hidden.append(i)
    except Exception:
        return out
    for i in hidden:
        try:
            ws.Rows(i).Hidden = False
        except Exception:
            continue
    if hidden:
        where = "・".join(str(i) for i in hidden[:_SORT_HIDDEN_SHOW])
        out.append(f"※ 隠れていた行を出してから並べました（{len(hidden)} 行: {where}"
                   + ("…" if len(hidden) > _SORT_HIDDEN_SHOW else "")
                   + "）。隠れた行は並べ替えで動かず、その場に取り残されます")
    return out


@protect_safe
@dialog_safe
def cmd_sort(args):
    """並べ替え: sort <range> [--key 列文字] [--desc] [--header|--no-header]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]")
        return False
    if _reject_extra_args(rest, 1, '使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を並べ替えるなら --whole-sheet を付けてください（--header の明示も推奨）。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    keycol = getattr(args, 'key', None)
    if keycol and not (keycol.isascii() and keycol.isalpha()):
        # _col_num は英字以外も黙って数値化してしまい、"A1" が K列扱いになる。
        # なお str.isalpha() は '名' のような日本語も True になるため isascii() が必須
        # （'名' を通すと _col_num が巨大な列番号を返し、意図しない列でソートされる）
        print(f"エラー: --key は列文字（A〜XFD）で指定してください: '{keycol}'")
        return False
    key_idx = _col_num(keycol) if keycol else rng.Column
    keycell = ws.Cells(rng.Row, key_idx)
    order = 2 if getattr(args, 'desc', False) else 1       # xlDescending=2 / xlAscending=1
    if getattr(args, 'header', False):
        header = 1                                          # xlYes
    elif getattr(args, 'no_header', False):
        header = 2                                          # xlNo
    else:
        header = 0                                          # xlGuess
    # Orientation / MatchCase / OrderCustom は Excel が「前回の並べ替え設定」をシートに
    # 保存して引き継ぐ仕様。未指定だと手動の列単位ソート等が引き継がれ、行方向のつもりが
    # 列方向に並べ替わる事故になるため必ず明示する（xlTopToBottom=1）。
    # OrderCustom も同族で、UI で一度「ユーザー設定リスト」（曜日順・部署順など）を
    # 使うとその順が残り、五十音順のつもりがカスタム順で並ぶ（OrderCustom=1＝通常順）
    # 隠れている行は Excel の並べ替えの対象外＝その行だけ元の位置に取り残される（2026-09-08 の実射
    # 「多い順に並べて」が、手で隠した 7・8 行目を末尾に置き去りにしたまま「合格」で通った）。
    # 値の並びが黙って狂うので、並べる前に出す。出したことは必ず報告する（黙って見せ方を変えない）。
    shown = _unhide_for_sort(ws, rng)
    rng.Sort(Key1=keycell, Order1=order, Header=header,
             Orientation=1, MatchCase=False, OrderCustom=1)
    print(f"並べ替え: {ws.Name}!{rng.Address}  キー列={keycol or _col_letter(rng.Column)}  "
          f"{'降順' if order == 2 else '昇順'}")
    for note in shown:
        print("  " + note)
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_autofilter(args):
    """オートフィルタ: autofilter [range] [--off]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if _reject_extra_args(rest, 1, '使い方: autofilter [range] [--off]'):
        return False
    xl, wb = get_workbook(target_file)
    if getattr(args, 'off', False):
        # 位置引数（範囲/シート名）があればそのシートを解除対象にする。
        # 黙って ActiveSheet に落とすと、指定したつもりの別シートでなく
        # アクティブなシートのフィルタ（絞り込み条件ごと）が消える
        if rest:
            ws, _ = _resolve_range(xl, wb, rest[0])
        else:
            ws = wb.ActiveSheet
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            print(f"オートフィルタ解除: {ws.Name}")
            print("（保存はしていません）")
        else:
            print(f"オートフィルタは設定されていません: {ws.Name}")
        return True
    spec = rest[0] if rest else None
    ws, rng = _resolve_range(xl, wb, spec)
    if getattr(args, 'show_all', False):
        if ws.FilterMode:
            ws.ShowAllData()
            print(f"絞り込みを解除（▼は残す）: {ws.Name}")
        else:
            print(f"絞り込みは掛かっていません: {ws.Name}")
        return True
    column = getattr(args, 'column_opt', None)
    values = getattr(args, 'equals_opt', None)
    if column or values:
        # 「◯◯だけ抜き出して」＝条件つきの絞り込み。既にフィルタが付いていても掛け直す
        # （本物の帳簿はたいてい最初からフィルタ付きで、そこで諦めていた・2026-09-04 実射）
        if not values:
            print("エラー: --column と --equals は一緒に指定してください（例 --column 件名 --equals 観光関係一般）")
            return False
        field = None
        if column is not None:
            if str(column).isdigit():
                field = int(column)
            else:
                head = rng.Rows(1)
                for i in range(1, int(rng.Columns.Count) + 1):
                    if str(head.Cells(1, i).Value or '').strip() == str(column).strip():
                        field = i
                        break
                if field is None:
                    heads = [str(head.Cells(1, i).Value or '') for i in range(1, int(rng.Columns.Count) + 1)]
                    print(f"エラー: 見出し '{column}' が範囲の 1 行目にありません（見出し: {heads}）")
                    return False
        else:
            field = 1
        # 既にあるフィルタの範囲が違うと、列番号がそちらの範囲で数えられて別の列が絞られる
        # （本物の帳簿は _FilterDatabase が B1:N58 のまま残っていた・2026-09-04 実測）。
        # 範囲が違うときはいったん外して、指定の範囲で掛け直す
        try:
            cur = ws.AutoFilter.Range.Address if ws.AutoFilterMode else None
        except Exception:
            cur = None
        if cur is not None and cur != rng.Address:
            ws.AutoFilterMode = False
            print(f"（前のフィルタ範囲 {cur} を外し、{rng.Address} で掛け直します）")
        elif ws.FilterMode:
            ws.ShowAllData()                     # 前の絞り込みを解いてから掛け直す
        if len(values) == 1:
            rng.AutoFilter(field, values[0])
        else:
            rng.AutoFilter(field, values, 7)     # xlFilterValues
        shown = 0
        for r in range(int(rng.Row) + 1, int(rng.Row) + int(rng.Rows.Count)):
            if not bool(ws.Rows(r).Hidden):
                shown += 1
        print(f"絞り込み: {ws.Name}!{rng.Address} の {field} 列目 = {' / '.join(map(str, values))}"
              f"（表示 {shown} 行）")
        print("（保存はしていません）")
        return True
    if ws.AutoFilterMode:
        if ws.FilterMode:
            # 前の絞り込みが効いたまま（行が隠れている）。黙って何もしないと、**別人の絞り込みが
            # 残っているのに done が通る**（2026-09-08 の実射「鈴木さんの分だけ」＝列を書かずに
            # autofilter だけ撃ち、前のフィルタが残ったまま合格になった）。
            # かといって失敗で突き返すと、「解除したい」ときに同じ手を繰り返して往復を使い切る
            # （同日・実射「テーブル化」）。絞り込みは見せ方であってデータではないので、
            # ここは _unfilter_for_write と同じ流儀＝**解いて、解いたことを報告する**。
            ws.ShowAllData()
            print(f"前の絞り込みを解きました: {ws.Name}（隠れていた行を出しました）")
            print("（絞り込むなら --column 見出し --equals 値。ここでは何も絞り込んでいません）")
            print("（保存はしていません）")
            return True
        print(f"既にオートフィルタが設定されています: {ws.Name}"
              f"（絞り込むなら --column 見出し --equals 値、解くなら --show-all）")
    else:
        # win32com の遅延バインディングでは引数なし rng.AutoFilter() が
        # 「AutoFilter メソッドが失敗しました」で落ちる（全省略可能引数を省くとCOMが弾く）。
        # Field:=1 だけ渡すと Criteria なし＝どの列も絞り込まずに範囲全体へフィルタUIを付ける
        # ＝引数なしと同じ「オートフィルタON」になる（実弾で確認済み）。
        rng.AutoFilter(1)
        print(f"オートフィルタ設定: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
    return True


# ---- b. 検索・置換 ----

def cmd_find(args):
    """セル検索: find <文字> [--book] [--whole] [--formula] [--sheet 名]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: find <文字> [--book(全シート)] [--whole(完全一致)] [--formula(式も検索)] [--sheet 名]")
        return False
    needle = rest[0]
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    # グラフシートには UsedRange が無く例外で検索全体が落ちるため Worksheets 限定
    if getattr(args, 'book', False):
        sheets = list(wb.Worksheets)
    elif sheet_opt:
        if sheet_opt not in [sh.Name for sh in wb.Sheets]:
            print(f"エラー: シート '{sheet_opt}' が見つかりません")
            return False
        sheets = [wb.Sheets(sheet_opt)]
    else:
        sheets = [wb.ActiveSheet]
    look_in = -4123 if getattr(args, 'formula', False) else -4163  # xlFormulas / xlValues
    look_at = 1 if getattr(args, 'whole', False) else 2            # xlWhole / xlPart
    total = 0
    # `or 200` だと --max 0（件数だけ見たい）が偽値で既定に化けるため is None 判定
    mh = getattr(args, 'max_hits', None)
    try:
        max_hits = 200 if mh is None else int(mh)
    except (TypeError, ValueError):
        print("エラー: --max は数値で指定してください")
        return False
    if max_hits < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False
    for ws in sheets:
        try:
            rng = ws.UsedRange
        except Exception:
            # アクティブがグラフシート等だと UsedRange 自体が例外になる
            print(f"（'{ws.Name}' はワークシートではないためスキップ）")
            continue
        try:
            cell = rng.Find(What=needle, LookIn=look_in, LookAt=look_at, MatchCase=False)
        except Exception:
            cell = None
        first = None
        while cell is not None:
            addr = cell.Address
            if first is None:
                first = addr
            elif addr == first:
                break
            total += 1
            if total <= max_hits:
                # 「シート名!$A$1」はそのまま write-range 等の range 引数に貼れる形
                print(f"{ws.Name}!{addr}: {cell.Value}")
            cell = rng.FindNext(cell)
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
    else:
        if total > max_hits:
            print(f"…他 {total - max_hits}件（--max で上限変更可）")
        print(f"--- {total}件 ヒット ---")
    return True


@protect_safe
@dialog_safe
def cmd_find_replace(args):
    """一括置換: find-replace <検索> <置換> [range] [--whole] [--wildcard]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: find-replace <検索> <置換> [range] [--whole] [--wildcard]")
        return False
    needle, repl = rest[0], rest[1]
    spec = rest[2] if len(rest) >= 3 else None
    if _reject_extra_args(rest, 3, '使い方: find-replace <検索> <置換> [range] [--whole] [--match-case] [--wildcard]'):
        return False
    # Excel の Find/Replace は * ? を常にワイルドカード解釈する。
    # 素通しすると `find-replace "*" "×"` が全非空セルの丸ごと置換になるため、
    # 既定は ~ エスケープで文字どおりに扱い、--wildcard で明示オプトインする
    if not getattr(args, 'wildcard', False):
        escaped = re.sub(r'([~*?])', r'~\1', needle)
        if escaped != needle:
            print("（* ? ~ は文字どおりに置換します。パターンとして使うなら --wildcard）")
        needle = escaped
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    if spec or sheet_opt:
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    else:
        ws = wb.ActiveSheet
        try:
            rng = ws.UsedRange
        except Exception:
            print(f"エラー: アクティブシート '{ws.Name}' は置換できる"
                  "ワークシートではありません（グラフシート等）。")
            return False
    look_at = 1 if getattr(args, 'whole', False) else 2
    match_case = getattr(args, 'match_case', False)
    # Range.Replace は置換件数を返さないため、置換前にヒットセル数を数える
    # （LookIn は Replace と同じ数式(-4123)で揃える）
    count = 0
    first = None
    # SearchFormat / MatchByte も Sort の Orientation と同じ「省略すると前回値を
    # 引き継ぐ」族なので、引数で明示して前回値に依存しない。
    # ・SearchFormat=False で書式条件は無視される（Application.FindFormat.Clear() まで
    #   撃つのは、ユーザーが UI に仕込んだ「書式を指定して検索」の設定を消す越権）
    # ・MatchByte=True＝半角/全角を区別する（厳密側）。False だと「アイウ」の置換が
    #   "ｱｲｳ" のセルまで当たり、指示より広く書き換わる
    cell = rng.Find(What=needle, LookAt=look_at, LookIn=-4123, MatchCase=match_case,
                    SearchFormat=False, MatchByte=True)
    while cell is not None:
        addr = cell.Address
        if first is None:
            first = addr
        elif addr == first:
            break
        count += 1
        cell = rng.FindNext(cell)
    if count == 0:
        print(f"'{needle}' は {ws.Name}!{rng.Address} に見つかりませんでした（置換なし）")
        return True
    raw = rest[0]
    if (look_at == 2 and not getattr(args, 'wildcard', False) and raw and raw != repl
            and (raw in repl if match_case else raw.lower() in repl.lower())):
        # 置き換え後が探す文字を含む（受注済→受注済み）。Excel の Replace は、もとから「受注済み」のセルにも当たって
        # 「受注済みみ」にする（2026-09-11）。セルごとに見て、置き換え後の文字をもとから含むセルは飛ばす
        vals, fmls = rng.Value, rng.Formula
        vals = vals if isinstance(vals, tuple) else ((vals,),)
        fmls = fmls if isinstance(fmls, tuple) else ((fmls,),)
        r0, c0 = int(rng.Row), int(rng.Column)
        pat = re.compile(re.escape(raw), 0 if match_case else re.I)
        done = skipped = 0
        for i, row in enumerate(vals):
            for j, v in enumerate(row):
                f = fmls[i][j] if i < len(fmls) and j < len(fmls[i]) else None
                if not isinstance(v, str) or (isinstance(f, str) and f.startswith('=')) or not pat.search(v):
                    continue
                if (repl in v) if match_case else (repl.lower() in v.lower()):
                    skipped += 1
                    continue
                ws.Cells(r0 + i, c0 + j).Value = pat.sub(lambda m: repl, v)
                done += 1
        print(f"置換: {ws.Name}!{rng.Address}  '{raw}' → '{repl}'  （{done}セル。置き換え後の「{repl}」を"
              f"もとから含む {skipped} セルは、二重にならないよう飛ばしました）")
        print("（保存はしていません）")
        return True
    rng.Replace(What=needle, Replacement=repl, LookAt=look_at, MatchCase=match_case,
                SearchFormat=False, ReplaceFormat=False, MatchByte=True)
    print(f"置換: {ws.Name}!{rng.Address}  '{needle}' → '{repl}'  （{count}セルにヒット）")
    print("（保存はしていません）")
    return True


# ---- c. 保存・印刷まわり ----

@dialog_safe
def cmd_save(args):
    """上書き保存: save [excel_file]"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    with _alerts_off(xl):
        wb.Save()
    print(f"保存しました: {wb.FullName}")
    return True


def cmd_save_as(args):
    """別名保存: save-as [excel_file] <path>（省略時はアクティブブックを対象）"""
    # 他コマンドと同じ流儀: 引数2つなら第1引数を対象ブック、第2引数を出力パスとする。
    # （以前は rest[0] を無条件に出力パスにしていたため、
    #   `save-as 既存ブック.xlsx 新名.xlsx` で既存ブックが無言上書きされる罠があった）
    rest = list(args.posargs)
    if not rest:
        print("使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) >= 3:
        print("エラー: 引数が多すぎます。使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) == 2:
        target_file, out_arg = rest[0], rest[1]
    else:
        target_file, out_arg = None, rest[0]
    out = os.path.abspath(out_arg)

    FMT = {'.xlsx': 51, '.xlsm': 52, '.xlsb': 50, '.xls': 56,
           '.csv': 6, '.txt': -4158}
    ext = os.path.splitext(out)[1].lower()
    if ext not in FMT:
        # 未知拡張子を黙って xlsx にフォールバックすると「中身xlsxの .pdf」等の壊れファイルになる
        print(f"エラー: 対応していない拡張子です: '{ext or '(なし)'}'")
        print(f"  対応: {' '.join(sorted(FMT))}")
        return False
    fmt = FMT[ext]

    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False

    xl, wb = get_workbook(target_file)
    try:
        old_full = wb.FullName
    except Exception:
        old_full = None
    src_ext = os.path.splitext(wb.Name)[1].lower()
    if src_ext in ('.xlsm', '.xlsb', '.xls') and ext == '.xlsx':
        # DisplayAlerts=False で Excel の警告が出ないため、こちらで明示する
        print("⚠ 注意: マクロ付きブックを .xlsx で保存するため、VBAマクロは保存されません。")
    if ext in ('.csv', '.txt'):
        try:
            n_sheets = wb.Sheets.Count
        except Exception:
            n_sheets = 1
        if n_sheets > 1:
            # これも DisplayAlerts=False で Excel 側の警告が抑止されるため明示する
            print(f"⚠ 注意: {ext} はアクティブシート1枚しか保存されません"
                  f"（このブックは {n_sheets} シート）。")
    with _alerts_off(xl):
        wb.SaveAs(out, FileFormat=fmt)
    # batch/shell/MCP の1接続セッションでは接続キャッシュが旧パスキーのまま残り、
    # 旧名を指定した後続コマンドが改名後のブックに当たる（対象取り違え）。
    # SaveAs 成功時にキャッシュのキーを新パスへ付け替える
    if old_full:
        old_key = old_full.lower()
        if old_key != out.lower() and old_key in _wb_cache:
            _wb_cache[out.lower()] = _wb_cache.pop(old_key)
    print(f"別名保存しました: {out}")
    print("  （以後、開いているブックの保存先はこの新パスになります）")
    return True


def cmd_open(args):
    """ブックを開く: open <path>

    「人が開くのと同じ場所」に開くのが原則:
      - 既に開いていれば前面化して知らせるだけ（二重には開かない）
      - 見えている Excel が居れば、そこに合流して開く（アドインが生きている普段の環境）
      - Excel が完全に未起動なら通常起動（os.startfile）で開く
        （COM 起動と違い、アドイン・PERSONAL.XLSB が普段どおり読み込まれる）
      - 非表示の残骸 Excel しか居ないときは開かない。残骸に取り込まれると
        「開いたのに見えない・アドインが効かない」異常環境になる（2026-06-13 実害）
    """
    rest = list(args.posargs)
    if not rest:
        print("使い方: open <path>")
        return False
    if len(rest) >= 2:
        print("エラー: パスは1つだけ指定してください。使い方: open <path>")
        return False
    path = smart_path_resolve(rest[0])
    if not path or not os.path.exists(path):
        print(f"エラー: ファイルが見つかりません: {rest[0]}")
        return False

    # 既に開いていないか（全インスタンス横断）
    for wb in _running_excel_workbooks():
        try:
            if not same_path(wb.FullName, path):
                continue
            app = wb.Application
            try:
                visible = bool(app.Visible)
            except Exception:
                visible = False
            print(f"既に開いています: {wb.Name}")
            if visible:
                try:
                    wb.Activate()
                except Exception:
                    pass
            else:
                print("⚠ ただし非表示の Excel インスタンスの中にいます。"
                      "タスクマネージャで EXCEL.EXE を確認してください。")
            return True
        except Exception:
            continue

    # 合流先＝見えている Excel を探す。まずブック持ちインスタンス（ROT）、
    # 次に GetActiveObject（未保存の Book1 しか無い Excel は ROT に出ないため）
    xl = None
    excel_somewhere = False
    for wb in _running_excel_workbooks():
        try:
            app = wb.Application
            excel_somewhere = True
            if bool(app.Visible):
                xl = app
                break
        except Exception:
            continue
    if xl is None:
        try:
            cand = _get_active_excel()
            excel_somewhere = True
            if bool(cand.Visible):
                xl = cand
        except Exception:
            pass

    if xl is not None:
        wb = xl.Workbooks.Open(path)
        try:
            wb.Activate()
        except Exception:
            pass
        print(f"開きました: {wb.Name}  （起動中の Excel に合流）")
        return True

    if not excel_somewhere:
        # COM から見えなくても EXCEL.EXE のプロセスだけ残っている残骸が居る
        # （そこへ os.startfile すると非表示側に取り込まれることがある）
        try:
            import subprocess
            out = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                capture_output=True, encoding="cp932", errors="replace", timeout=30)
            excel_somewhere = "EXCEL.EXE" in (out.stdout or "")
        except Exception:
            pass

    if excel_somewhere:
        print("エラー: 非表示の EXCEL.EXE だけが残っています（残骸）。")
        print("  ここで開くと非表示側に取り込まれて見えなくなるため、開きませんでした。")
        print("  タスクマネージャで EXCEL.EXE を終了してから再実行してください。")
        return False

    # Excel 完全未起動 → 通常起動。アドイン・PERSONAL.XLSB が普段どおり読み込まれる
    os.startfile(path)
    limit = time.time() + 30.0
    while time.time() < limit:
        for wb in _running_excel_workbooks():
            try:
                if same_path(wb.FullName, path):
                    print(f"開きました: {wb.Name}  （Excel を通常起動）")
                    return True
            except Exception:
                continue
        time.sleep(0.5)
    print("⚠ 30秒待ちましたが、開いたことを確認できませんでした。")
    print("  保護ビュー・マクロ警告などのダイアログで止まっている可能性があります。"
          "画面を確認してください。")
    return False


def cmd_close(args):
    """ブックを閉じる: close <ブック名|path> (--save | --no-save) [-y]

    鎧の三点セット: ①ブックの名指し必須 ②保存方針（--save/--no-save）の明示必須
    ③確認プロンプト（-y でスキップ）。閉じるのはブック1冊だけで、Excel 本体は
    終了しない。PERSONAL.XLSB とアドインブックは閉じない（保管庫を巻き込まない）。
    """
    rest = list(args.posargs)
    if not rest:
        print("使い方: close <ブック名|path> (--save | --no-save) [-y]")
        return False
    if len(rest) >= 2:
        print("エラー: 対象は1つだけ指定してください。使い方: close <ブック名|path> (--save | --no-save)")
        return False
    save_flag = bool(getattr(args, 'save_flag', False))
    no_save_flag = bool(getattr(args, 'no_save_flag', False))
    if save_flag == no_save_flag:
        print("エラー: --save（保存して閉じる）か --no-save（保存せず閉じる＝変更破棄）の"
              "どちらか一方を必ず指定してください。")
        return False

    target = rest[0]
    resolved = smart_path_resolve(target)
    tgt_lower = target.lower()

    # 対象を全インスタンス横断で探す。同名でも別インスタンス/別パスは別物として数える
    matches = {}

    def _consider(wb):
        try:
            full = wb.FullName
            hwnd = wb.Application.Hwnd
        except Exception:
            return
        name = os.path.basename(full)
        if (name.lower() == tgt_lower or full.lower() == tgt_lower or
                (resolved and full.lower() == resolved.lower())):
            matches[(full.lower(), hwnd)] = wb

    for wb in _running_excel_workbooks():
        _consider(wb)
    try:
        # ROT に出ない未保存ブック（Book1 等）はこちらで拾う
        for wb in _get_active_excel().Workbooks:
            _consider(wb)
    except Exception:
        pass

    if not matches:
        print(f"エラー: '{target}' は開いていません。list-open で確認してください。")
        return False
    if len(matches) > 1:
        print("エラー: 該当するブックが複数開いています。フルパスで名指ししてください:")
        for wb in matches.values():
            try:
                print(f"  {wb.FullName}")
            except Exception:
                pass
        return False
    wb = next(iter(matches.values()))
    name = wb.Name

    # 番人: 保管庫は閉じない
    if name.lower() == 'personal.xlsb':
        print("エラー: PERSONAL.XLSB は閉じません（個人用マクロの保管庫。閉じると"
              "登録マクロやショートカットが使えなくなります）。")
        return False
    try:
        if bool(wb.IsAddin):
            print(f"エラー: '{name}' はアドインブックです。close では閉じません。")
            return False
    except Exception:
        pass

    try:
        full = wb.FullName
    except Exception:
        full = name
    try:
        saved = bool(wb.Saved)
    except Exception:
        saved = None

    if save_flag:
        try:
            if bool(wb.ReadOnly):
                print(f"エラー: '{name}' は読み取り専用で開かれています。--save では閉じられません。")
                print("  変更を捨ててよければ --no-save、別名で残すなら先に save-as を。")
                return False
        except Exception:
            pass

    print(f"対象ブック: {name}")
    print(f"  パス: {full}")
    print(f"  方針: {'保存して閉じる' if save_flag else '保存せずに閉じる'}")
    if not save_flag and saved is False:
        print("  ⚠ 未保存の変更があります。このまま閉じると、その変更は破棄されます。")

    if not getattr(args, 'yes', False):
        try:
            ans = input(f"ブック '{name}' を閉じますか？ (y/N): ")
        except EOFError:
            # パイプ/MCP 等の非対話環境。トレースバックでなく正常なキャンセルにする
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    app = wb.Application
    with _alerts_off(app):
        wb.Close(SaveChanges=save_flag)

    # 接続キャッシュに残った死んだ参照を掃除する（batch/shell/MCP の1接続セッションで
    # 次のコマンドが閉じたブックのキャッシュを掴まないように）
    for key in list(_wb_cache):
        try:
            _ = _wb_cache[key][1].Name
        except Exception:
            _wb_cache.pop(key, None)

    print(f"閉じました: {name}  ({'保存済み' if save_flag else '保存せず'})")
    return True


def _list_form_windows():
    """Excel が表示している UserForm の窓を [(hwnd, キャプション), ...] で返す。

    MSForms の UserForm は最上位ウィンドウ（クラス名 ThunderDFrame）として出る。
    所有プロセスが Excel（XLMAIN 窓を持つ PID）のものだけ拾う（Word/Access の
    UserForm は対象外）。COM を使わないので、モーダル表示中でも見える。
    win32gui が無い環境では [] を返す。
    """
    try:
        import win32gui
        import win32process
    except Exception:
        return []
    excel_pids = set()
    forms = []

    def _cb(hwnd, _unused):
        try:
            cls = win32gui.GetClassName(hwnd)
            if cls == 'XLMAIN':
                excel_pids.add(win32process.GetWindowThreadProcessId(hwnd)[1])
            elif cls == 'ThunderDFrame' and win32gui.IsWindowVisible(hwnd):
                forms.append((hwnd, win32process.GetWindowThreadProcessId(hwnd)[1],
                              win32gui.GetWindowText(hwnd) or ''))
        except Exception:
            pass

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return []
    return [(h, cap) for (h, pid, cap) in forms if pid in excel_pids]


def _pick_form_windows(forms, target):
    """表示中フォーム [(hwnd, キャプション)] から target に合うものを選ぶ（純粋関数）。

    target 空＝全部。完全一致（大小無視）→部分一致の順。どれにも当たらなければ []。
    """
    if not target:
        return list(forms)
    t = target.strip().lower()
    exact = [f for f in forms if f[1].strip().lower() == t]
    if exact:
        return exact
    return [f for f in forms if t in f[1].lower()]


def cmd_close_form(args):
    """表示中の UserForm を閉じる: close-form [キャプション] [--list] [--wait 秒]

    「フォームを閉じてください」と人に頼まない。フォームの窓（ThunderDFrame）に
    ×ボタンと同じ WM_SYSCOMMAND(SC_CLOSE) を送って閉じる。COM を使わないので、
    モーダル表示中（COM 呼び出しが「呼び出し先が拒否しました」で弾かれる状態）でも効く。
    ⚠ WM_CLOSE を直接投げてはいけない（2026-08-23 実測）。窓だけ壊れて VBA の Unload を
      通らず、フォームの既定インスタンスが「切断」状態で残る。次に開くと
      実行時エラー 80010108（起動されたオブジェクトはクライアントから切断されました）
      → Excel ごと落ちる。SC_CLOSE なら QueryClose→Unload→Terminate の正規の道を通る。
    引数なし＝Excel が表示している全フォーム。キャプション（窓の題名。完全一致→部分一致）
    で絞れる。--list は閉じずに一覧だけ。
    QueryClose で Cancel するフォーム（実行中の Excelコンボ等）は閉じられない。
    その事実をそのまま報告し、無理に殺さない（プロセスやスレッドには触れない）。
    """
    rest = list(getattr(args, 'posargs', []) or [])
    if len(rest) >= 2:
        print("エラー: 対象は1つだけ指定してください。使い方: close-form [キャプション] [--list]")
        return False
    target = rest[0] if rest else ''
    try:
        import win32gui
        import win32con
    except Exception as ex:
        print(f"エラー: win32gui が使えません（pywin32 不足）: {ex}")
        return False

    forms = _list_form_windows()
    if getattr(args, 'list_flag', False):
        if not forms:
            print("表示中のフォームはありません。")
        else:
            print(f"表示中のフォーム: {len(forms)}件")
            for _h, cap in forms:
                print(f"  {cap or '(無題)'}")
        return True
    if not forms:
        print("表示中のフォームはありません。")
        return True
    picked = _pick_form_windows(forms, target)
    if not picked:
        print(f"エラー: '{target}' というフォームは表示されていません。表示中のフォーム:")
        for _h, cap in forms:
            print(f"  {cap or '(無題)'}")
        return False

    try:
        wait_sec = float(getattr(args, 'wait_opt', None) or 3.0)
    except (TypeError, ValueError):
        wait_sec = 3.0
    pending = {h: cap for h, cap in picked}
    closed = []
    last_sent = {}
    deadline = time.time() + max(0.5, wait_sec)
    # PostMessage は非同期＝送っただけでは閉じたことにならない（ダイアログ監視と同じ教訓）。
    # 窓が実際に消えたのを見て初めて「閉じた」と数え、消えなければ 0.6 秒おきに再送する。
    # 送るのは WM_SYSCOMMAND+SC_CLOSE（×クリックと同一）。WM_CLOSE 直送は厳禁（docstring 参照）。
    while pending and time.time() < deadline:
        now = time.time()
        for h in list(pending):
            try:
                alive = win32gui.IsWindow(h) and win32gui.IsWindowVisible(h)
            except Exception:
                alive = False
            if not alive:
                closed.append(pending.pop(h))
                continue
            if now - last_sent.get(h, 0) >= 0.6:
                try:
                    win32gui.PostMessage(h, win32con.WM_SYSCOMMAND, win32con.SC_CLOSE, 0)
                except Exception:
                    pass
                last_sent[h] = now
        if pending:
            time.sleep(0.15)
    for cap in closed:
        print(f"閉じました: {cap or '(無題)'}")
    for cap in pending.values():
        print(f"閉じられませんでした: {cap or '(無題)'}"
              "（フォームが閉じるのを拒んでいます＝QueryClose で Cancel／処理の実行中）")
    return not pending


def cmd_vbe_reset(args):
    """VBE の「実行 > リセット」を押す: vbe-reset [excel_file] [--check]

    実行時エラーで VBE が中断モード（黄色い行で止まったまま）になると、
    フォームの Designer が取れない・マクロが「実行できません」になる・
    フォーム編集ツールが「VBE が中断モードでないか確認してください」で止まる。
    その解除を人に頼まず、標準ツールバーのリセット（Id 228）を外から押す。
    リセットは走っている VBA を全部止める（表示中のフォームも閉じる・モジュール変数も消える・
    Excelコンボが温めている claude も切れる）ので、**自動では撃たない**。人が「戻せ」と
    言ったときだけ使う。--check は押さずに「今リセットが要る状態か」だけ返す。
    状態の判定は VBProject.Mode（0=実行中 1=中断中 2=設計）。リセットボタンの Enabled は
    VBE 窓が隠れていると更新されず当てにならない（実測 2026-08-23）。
    """
    target_file, rest = parse_target_and_rest(list(getattr(args, 'posargs', []) or []))
    xl, wb = get_workbook(target_file)
    modes = []
    try:
        for p in xl.VBE.VBProjects:
            try:
                modes.append(int(p.Mode))
            except Exception:
                continue
    except Exception as ex:
        print(f"エラー: VBE に触れません（VBOM の信頼設定を確認）: {ex}", file=sys.stderr)
        return False
    busy = any(m != 2 for m in modes)
    state = "中断中（Break）" if 1 in modes else ("実行中（Run）" if 0 in modes else "設計モード")
    if not busy:
        print("VBE は実行中でも中断中でもありません（設計モード＝リセット不要）")
        return True
    if getattr(args, 'check', False):
        print(f"VBE は{state}です（リセットで解除できます）")
        return True
    ctl = None
    try:
        for bar in xl.VBE.CommandBars:
            try:
                for c in bar.Controls:
                    try:
                        if c.Id == 228:
                            ctl = c
                            break
                    except Exception:
                        continue
            except Exception:
                continue
            if ctl is not None:
                break
    except Exception as ex:
        print(f"エラー: VBE に触れません（VBOM の信頼設定を確認）: {ex}", file=sys.stderr)
        return False
    if ctl is None:
        print("エラー: VBE のリセットボタン（Id 228）が見つかりませんでした", file=sys.stderr)
        return False
    try:
        ctl.Execute()
    except Exception as ex:
        print(f"エラー: リセットを押せませんでした: {ex}", file=sys.stderr)
        return False
    # 押したあとの状態で報告する（押せても解けないことがあれば、そのまま言う）
    after = []
    try:
        for p in xl.VBE.VBProjects:
            try:
                after.append(int(p.Mode))
            except Exception:
                continue
    except Exception:
        pass
    if after and all(m == 2 for m in after):
        print(f"VBE をリセットしました（{state} → 設計モード）。走っていた VBA は止まり、表示中のフォームも閉じています")
        return True
    print(f"リセットを押しましたが、まだ{('中断中' if 1 in after else '実行中')}のままです"
          "（モーダルなダイアログが出ていないか確認してください）", file=sys.stderr)
    return False


def cmd_export_pdf(args):
    """PDF出力: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range "シート!範囲"]

    ExportAsFixedFormat による出力。既定はブック全体、--sheet で1シート、
    --range で範囲のみ。ブック自体は変更しない（保存フラグも汚さない）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range \"シート!A1:H50\"] [--overwrite]")
        return False
    out = os.path.abspath(rest[0])
    if _reject_extra_args(rest, 1, '出力パスは1つだけ指定してください'):
        return False
    if not out.lower().endswith('.pdf'):
        print(f"エラー: 出力は .pdf で指定してください: {out}")
        return False
    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False
    sheet_opt = getattr(args, 'sheet_opt', None)
    range_opt = getattr(args, 'range_opt', None)
    if sheet_opt and range_opt:
        print("エラー: --sheet と --range は同時に指定できません")
        return False

    xl, wb = get_workbook(target_file)
    if range_opt:
        ws, rng = _resolve_range(xl, wb, range_opt)
        rng.ExportAsFixedFormat(0, out)               # 0 = xlTypePDF
        scope = f"範囲 {ws.Name}!{rng.Address}"
    elif sheet_opt:
        ws = None
        for sh in wb.Worksheets:
            if sh.Name == sheet_opt:
                ws = sh
                break
        if ws is None:
            print(f"エラー: シート '{sheet_opt}' が見つかりません")
            return False
        ws.ExportAsFixedFormat(0, out)
        scope = f"シート '{ws.Name}'"
    else:
        wb.ExportAsFixedFormat(0, out)
        scope = "ブック全体"
    print(f"PDF出力: {scope} → {out}")
    return True


@protect_safe
def cmd_shape(args):
    """図形: shape --list / shape <名前> [<名前>…] --delete / shape <名前> --left N --top N --width N --height N

    名前・リンクを消す手は無い（行・列は row／col の delete で消せる・2026-09-05）。マクロが割り当たっている図（ボタン）は
    動かすのも消すのも拒む＝材料の「図形・ボタン」で → の付いた図は触らない（2026-09-04）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    ws = wb.Sheets(sheet_opt) if sheet_opt else wb.ActiveSheet

    def _pt(v):
        # 整数に丸めると 123.5 が 124 と出て、その数で当て直すと図形がずれる。端数は小数 2 桁まで出す
        return f"{float(v):.2f}".rstrip('0').rstrip('.')

    def _info(shp):
        try:
            oa = str(shp.OnAction or '')
        except Exception:
            oa = ''
        return (f"  {shp.Name}" + (f" → {oa}" if oa else "")
                + f"  [l={_pt(shp.Left)} t={_pt(shp.Top)} "
                  f"w={_pt(shp.Width)} h={_pt(shp.Height)}]")

    names = [str(n).strip() for n in rest if str(n).strip()]

    if getattr(args, 'list_shapes', False) or not names:
        print(f"シート: {ws.Name}   図形 {int(ws.Shapes.Count)} 個")
        for shp in ws.Shapes:
            print(_info(shp))
        if not names and not getattr(args, 'list_shapes', False):
            print("（対象の名前がありません。消すなら shape <名前> --delete、動かすなら --left/--top/--width/--height）")
        return True

    picked = []
    for nm in names:
        try:
            shp = ws.Shapes(nm)
        except Exception:
            print(f"エラー: 図形 '{nm}' が無い（上の一覧に出ている名前を使う）")
            return False
        try:
            oa = str(shp.OnAction or '')
        except Exception:
            oa = ''
        if oa:
            verb = "消さない（ボタンを壊すため）" if getattr(args, 'delete', False) else "動かさない"
            print(f"エラー: '{nm}' にはマクロ（{oa}）が割り当たっている＝{verb}")
            return False
        picked.append((nm, shp))

    if getattr(args, 'delete', False):
        print("消す前の図形:")
        for shp in ws.Shapes:
            print(_info(shp))
        for nm, shp in picked:
            shp.Delete()
        print(f"\n図形を消しました（{len(picked)} 個）: " + "、".join(n for n, _ in picked))
        print(f"残っている図形: {int(ws.Shapes.Count)} 個")
        for shp in ws.Shapes:
            print(_info(shp))
        print("（保存はしていません。戻すなら保存せずに閉じる）")
        return True

    if len(picked) > 1:
        print("エラー: 名前を複数まとめて扱えるのは --delete のときだけ（位置と大きさは1つずつ）")
        return False

    nm, shp = picked[0]
    keys = [k for k in ('left', 'top', 'width', 'height') if getattr(args, k, None) is not None]
    if not keys:
        print("エラー: 当てる項目がありません（--left/--top/--width/--height。消すなら --delete）")
        return False
    for k in keys:
        setattr(shp, k.capitalize(), float(getattr(args, k)))
    print(f"図形: {ws.Name}  [{', '.join(f'{k}={getattr(args, k)}' for k in keys)}]")
    print(_info(shp))
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_print_setup(args):
    """印刷設定: print-setup [--area R] [--title-rows 1:3] [--title-cols A:B] ..."""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ps = ws.PageSetup
    applied = []

    if getattr(args, 'area', None):
        ps.PrintArea = ws.Range(args.area).Address
        applied.append(f"area={args.area}")
    if getattr(args, 'title_rows', None):
        a, b = (args.title_rows.split(':') + [args.title_rows])[:2]
        ps.PrintTitleRows = f"${a}:${b}"
        applied.append(f"title-rows={args.title_rows}")
    if getattr(args, 'title_cols', None):
        a, b = (args.title_cols.split(':') + [args.title_cols])[:2]
        ps.PrintTitleColumns = f"${a}:${b}"
        applied.append(f"title-cols={args.title_cols}")
    if getattr(args, 'landscape', False):
        ps.Orientation = 2; applied.append("landscape")
    if getattr(args, 'portrait', False):
        ps.Orientation = 1; applied.append("portrait")
    if getattr(args, 'fit_wide', None) is not None:
        ps.Zoom = False; ps.FitToPagesWide = int(args.fit_wide)
        applied.append(f"fit-wide={args.fit_wide}")
    if getattr(args, 'fit_tall', None) is not None:
        ps.Zoom = False; ps.FitToPagesTall = int(args.fit_tall)
        applied.append(f"fit-tall={args.fit_tall}")
    if getattr(args, 'zoom', None) is not None:
        ps.Zoom = int(args.zoom); applied.append(f"zoom={args.zoom}")
    if getattr(args, 'center_h', False):
        ps.CenterHorizontally = True; applied.append("center-h")
    if getattr(args, 'center_v', False):
        ps.CenterVertically = True; applied.append("center-v")

    if not applied:
        print("オプションが指定されていません。--area / --title-rows / --landscape など。")
        return False
    print(f"印刷設定: {ws.Name}  [{', '.join(applied)}]")
    print("（保存はしていません）")
    return True


def cmd_printer_list(args):
    """プリンター一覧およびアクティブプリンターを取得"""
    target_file, _ = parse_target_and_rest(args.posargs)
    
    # Excelのアクティブプリンターを取得
    active_printer = None
    try:
        xl, wb = get_workbook(target_file)
        active_printer = xl.ActivePrinter
        print(f"現在のアクティブプリンター: {active_printer}")
    except Exception as e:
        print(f"警告: Excelからアクティブプリンターを取得できませんでした ({e})")

    # WMI経由でOSにインストールされているプリンター一覧を取得
    import win32com.client
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        printers = wmi.InstancesOf("Win32_Printer")
        print("\nインストールされているプリンター一覧:")
        for printer in printers:
            name = printer.Name
            status = " (選択中)" if active_printer and name in active_printer else ""
            print(f"  - {name}{status}")
    except Exception as e:
        print(f"エラー: インストールされているプリンター一覧を取得できませんでした ({e})")
        return False
    return True


_PRINTER_DUPLEX = {'simplex': 1, 'vertical': 2, 'horizontal': 3}
_PRINTER_COLOR = {'mono': 1, 'color': 2}
_PRINTER_ORIENT = {'portrait': 1, 'landscape': 2}


def cmd_printer_setup(args):
    """プリンターの詳細設定（両面印刷・カラー等）を変更・表示

    ※ 他の「手」コマンドと違い、書き込み先はブックではなく **OS のプリンター設定** そのもの
      （win32print.SetPrinter）。保存せず閉じて破棄する逃げ道が無く、即時・不可逆で
      他のアプリの印刷にも影響する。オプション無しで呼べば現在の構成の表示だけ。
    """
    # 対象プリンター名の決定
    printer_name = getattr(args, 'printer', None)
    if not printer_name:
        # Excelが起動していればそのアクティブプリンター名を使用、さもなくばデフォルトプリンター
        try:
            target_file, _ = parse_target_and_rest(args.posargs)
            xl, wb = get_workbook(target_file)
            raw_printer = xl.ActivePrinter
            if " on " in raw_printer:
                printer_name = raw_printer.split(" on ")[0]
            else:
                printer_name = raw_printer
        except Exception:
            import win32print
            printer_name = win32print.GetDefaultPrinter()

    if not printer_name:
        print("エラー: 対象プリンターが特定できません。--printer で指定してください。")
        return False

    # 辞書に無い値（--duplex bogus 等）は、以前は無言で捨てられ applied が空のまま
    # 「現在のプリンター構成」を表示して成功扱い＝変更依頼が消えていた。先に弾く。
    wants = []                       # [(名前, 設定するDevMode属性, 値)]
    for opt_name, opt_attr, table in (('--duplex', 'Duplex', _PRINTER_DUPLEX),
                                      ('--color', 'Color', _PRINTER_COLOR),
                                      ('--orientation', 'Orientation', _PRINTER_ORIENT)):
        raw = getattr(args, opt_name.lstrip('-').replace('-', '_'), None)
        if not raw:
            continue
        val = table.get(str(raw).lower())
        if val is None:
            print(f"エラー: {opt_name} に指定できない値です: '{raw}'")
            print(f"  受け付ける値: {' | '.join(table)}")
            return False
        wants.append((f"{opt_name.lstrip('-')}={raw}", opt_attr, val))

    print(f"対象プリンター: {printer_name}")
    if wants:
        # 他コマンドの「（保存はしていません）」と真逆＝逃げ道が無いことを明示する
        print("⚠ 注意: これは Windows のプリンター設定そのものを書き換えます"
              "（即時反映・元に戻す操作なし・Excel 以外の印刷にも影響します）。")

    import win32print
    try:
        # 設定変更に必要なアクセス権を指定 (PRINTER_ACCESS_ADMINISTER=4, PRINTER_ACCESS_USE=8)
        access = 4 | 8
        handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": access})
    except Exception:
        try:
            handle = win32print.OpenPrinter(printer_name)
        except Exception as e:
            print(f"エラー: プリンターを開けませんでした ({e})")
            return False

    try:
        info = win32print.GetPrinter(handle, 2)
        devmode = info["pDevMode"]
        if devmode is None:
            print("エラー: プリンターの構成情報 (DevMode) を取得できませんでした。")
            return False

        # 値の妥当性は上で検査済み（未知の値はここへ来ない＝無言で捨てない）
        applied = []
        for label, attr, val in wants:
            setattr(devmode, attr, val)
            applied.append(label)

        if applied:
            win32print.SetPrinter(handle, 2, info, 0)
            print(f"プリンター設定更新完了: [{', '.join(applied)}]")
        else:
            # 現在の設定を表示
            duplex_names = {1: '片面 (simplex)', 2: '両面/長辺綴じ (vertical)', 3: '両面/短辺綴じ (horizontal)'}
            color_names = {1: 'モノクロ (mono)', 2: 'カラー (color)'}
            orient_names = {1: '縦向き (portrait)', 2: '横向き (landscape)'}
            
            d_val = duplex_names.get(devmode.Duplex, f"不明({devmode.Duplex})")
            c_val = color_names.get(devmode.Color, f"不明({devmode.Color})")
            o_val = orient_names.get(devmode.Orientation, f"不明({devmode.Orientation})")
            
            print(f"現在のプリンター構成:")
            print(f"  - 両面印刷: {d_val}")
            print(f"  - カラー　: {c_val}")
            print(f"  - 用紙向き: {o_val}")
            
    except Exception as e:
        print(f"エラー: プリンター設定の変更に失敗しました ({e})")
        return False
    finally:
        win32print.ClosePrinter(handle)
    return True


# ---- d. 仕上げ・見た目 ----

_XL_COND_OP = {'gt': 5, 'lt': 6, 'eq': 3, 'ne': 4, 'ge': 7, 'le': 8, 'between': 1}


@protect_safe
def cmd_cond_format(args):
    """条件付き書式(セルの値): cond-format <range> --gt 100 --bg '#FFC7CE'"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: cond-format <range> [--gt|--lt|--ge|--le|--eq|--ne 値 | --between v1 v2]")
        print("         [--bg '#RRGGBB'] [--color '#RRGGBB'] [--bold] [--clear]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: cond-format [excel_file] <range> [--gt 値 ...|--clear]'):
        return False
    xl, wb = get_workbook(target_file)
    whole = _whole_sheet_spec(wb, spec)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec)

    if getattr(args, 'clear', False):
        rng.FormatConditions.Delete()
        print(f"条件付き書式を全削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True

    if getattr(args, 'duplicates', False):
        # 「同じ行が二重に入っている」を色で示す（2026-09-04・本物の帳簿の実射で手が無くて詰まった）。
        # AddUniqueValues + DupeUnique=xlDuplicate ＝ Excel の「重複する値」の規則そのもの
        fc = rng.FormatConditions.AddUniqueValues()
        fc.DupeUnique = 1                       # xlDuplicate（2 回以上ある値を塗る）
        if getattr(args, 'bg', None):
            fc.Interior.Color = _hex_to_excel_color(args.bg)
        else:
            fc.Interior.Color = _hex_to_excel_color('#FFC7CE')
        if getattr(args, 'color', None):
            fc.Font.Color = _hex_to_excel_color(args.color)
        if getattr(args, 'bold', False):
            fc.Font.Bold = True
        print(f"条件付き書式（重複する値）を追加: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True

    formula = getattr(args, 'formula_opt', None)
    if formula:
        # 数式ベースのルール（xlExpression=2）。'=' 始まりの数式が TRUE のセルに書式。
        # 相対参照は範囲の左上セル基準（Excel の条件付き書式と同じ規約）
        if not formula.startswith('='):
            formula = '=' + formula
        # xlExpression は名前付き引数だと DISP_E_PARAMNOTOPTIONAL になるため位置渡し
        # （Operator は式タイプでは無視されるがスロットとして必要。1=xlBetween をダミーに）
        fc = rng.FormatConditions.Add(2, 1, formula)
    else:
        op = None; f1 = None; f2 = None
        for name in ('gt', 'lt', 'eq', 'ne', 'ge', 'le'):
            v = getattr(args, name, None)
            if v is not None:
                op = _XL_COND_OP[name]; f1 = str(v); break
        if op is None and getattr(args, 'between', None):
            op = _XL_COND_OP['between']; f1, f2 = args.between[0], args.between[1]
        if op is None:
            print("比較条件がありません。--gt 100 か --formula \"=数式\" を指定してください。")
            return False

        if f2 is not None:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1, Formula2=f2)
        else:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1)

    if getattr(args, 'bg', None):
        fc.Interior.Color = _hex_to_excel_color(args.bg)
    if getattr(args, 'color', None):
        fc.Font.Color = _hex_to_excel_color(args.color)
    if getattr(args, 'bold', False):
        fc.Font.Bold = True
    print(f"条件付き書式を追加: {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_hyperlink(args):
    """ハイパーリンク: hyperlink <cell> <url> [--text 表示文字] / --remove / --list

    url を省略して単セルを指定すると、そのセルのリンクを表示する（取得）。
    --list でシート内の全ハイパーリンクを一覧する。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    list_opt = getattr(args, 'list_links', None)
    if list_opt is not None:
        # シート内の全リンク。シート名は --list シート名 か 位置引数、無ければアクティブ
        want = None
        if list_opt != "__ACTIVE__":
            want = list_opt
        elif rest:
            want = rest[0]
        if want:
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == want:
                    ws = sh
                    break
            if ws is None:
                # 黙ってアクティブシートに落とすと「指定シートは0件」と誤読される
                print(f"エラー: シート '{want}' が見つかりません")
                return False
        else:
            ws = wb.ActiveSheet
        cnt = ws.Hyperlinks.Count
        print(f"--- {ws.Name} のハイパーリンク（{cnt}件） ---")
        for i in range(1, cnt + 1):
            hl = ws.Hyperlinks.Item(i)
            try:
                addr = hl.Range.Address
            except Exception:
                addr = '(図形)'
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"  {ws.Name}!{addr}: {hl.Address or ''}{sub}")
        return True

    if not rest:
        print("使い方: hyperlink <cell> <url> [--text 表示文字]")
        print("       hyperlink <cell>            # そのセルのリンクを表示")
        print("       hyperlink <cell> --remove   # 削除")
        print("       hyperlink --list [シート名]  # シート内の全リンク一覧")
        return False
    if getattr(args, 'remove', False):
        # シート名だけの指定は使用範囲全域のリンク一括削除（書式も戻る）になるためガード
        whole = _whole_sheet_spec(wb, rest[0])
        if whole is not None and not getattr(args, 'whole_sheet', False):
            print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  セル/範囲を明示するか、全リンクを消すなら --whole-sheet を付けてください。")
            return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'remove', False):
        rng.Hyperlinks.Delete()
        print(f"ハイパーリンク削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 取得モード: そのセルのリンクを表示
        cell = rng.Cells(1, 1)
        if cell.Hyperlinks.Count == 0:
            print(f"{ws.Name}!{cell.Address}: ハイパーリンクはありません")
        else:
            hl = cell.Hyperlinks.Item(1)
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"{ws.Name}!{cell.Address}: {hl.Address or ''}{sub}"
                  f"  表示=「{cell.Value}」")
        return True
    url = rest[1]
    cell = rng.Cells(1, 1)
    ws.Hyperlinks.Add(Anchor=cell, Address=url)
    # TextToDisplay は環境により効かないので、表示文字は明示的にセル値で上書き
    if getattr(args, 'text', None):
        cell.Value = args.text
    print(f"ハイパーリンク追加: {ws.Name}!{cell.Address} → {url}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_validation(args):
    """入力規則(ドロップダウン): validation <range> --list 'A,B,C' / --clear"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear")
        return False
    if _reject_extra_args(rest, 1, "使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear"):
        return False
    xl, wb = get_workbook(target_file)
    # --clear も設定パス（先に既存規則を Delete する）も破壊的なので、
    # シート名だけの指定＝使用範囲全域は明示なしでは拒否する
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'clear', False):
        rng.Validation.Delete()
        print(f"入力規則を削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    lst = getattr(args, 'list', None)
    if not lst:
        print("--list 'A,B,C' を指定してください。")
        return False
    rng.Validation.Delete()
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=lst)  # xlValidateList=3
    rng.Validation.InCellDropdown = True
    print(f"入力規則(リスト)を設定: {ws.Name}!{rng.Address}  [{lst}]")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_freeze(args):
    """ウィンドウ枠固定: freeze <cell> / freeze off"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: freeze <cell>（そのセルの左上で固定）  /  freeze off")
        return False
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ws.Activate()
    if rest[0].lower() == 'off':
        xl.ActiveWindow.FreezePanes = False
        print(f"枠固定を解除: {ws.Name}")
    else:
        ws.Range(rest[0]).Select()
        xl.ActiveWindow.FreezePanes = True
        print(f"枠固定: {ws.Name} {rest[0]} の左上で固定")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_comment(args):
    """セルコメント: comment <cell> <text> / comment <cell> --remove"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: comment <cell> <text>  /  comment <cell> --remove")
        return False
    xl, wb = get_workbook(target_file)
    # シート名だけの指定は使用範囲の左上（A1とは限らない）に黙って命中するため拒否
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None:
        print(f"エラー: '{rest[0]}' はセルではなくシート '{whole}' を指します。"
              f"セルを明示してください（例: {whole}!A1）")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    cell = rng.Cells(1, 1)
    if getattr(args, 'remove', False):
        cell.ClearComments()
        print(f"コメント削除: {ws.Name}!{cell.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 引数不足のときに既存コメントを消さないよう、ClearComments はチェックの後
        print("使い方: comment <cell> <text>")
        return False
    cell.ClearComments()
    cell.AddComment(rest[1])
    print(f"コメント追加: {ws.Name}!{cell.Address}")
    print("（保存はしていません）")
    return True




__all__ = [
    '_LAST_DAX_FILE',
    '_LAST_QUERY_FILE',
    '_PRINTER_COLOR',
    '_PRINTER_DUPLEX',
    '_PRINTER_ORIENT',
    '_XL_ALIGN_H',
    '_XL_ALIGN_V',
    '_XL_BORDER_WEIGHT',
    '_XL_COND_OP',
    '_alerts_off',
    '_col_num',
    '_hex_to_excel_color',
    '_list_form_windows',
    '_pick_form_windows',
    '_read_tsv_grid',
    '_sheet_or_active',
    'cmd_autofilter',
    'cmd_clear_range',
    'cmd_close',
    'cmd_close_form',
    'cmd_vbe_reset',
    'cmd_col',
    'cmd_comment',
    'cmd_cond_format',
    'cmd_copy_range',
    'cmd_export_pdf',
    'cmd_fill',
    'cmd_find',
    'cmd_find_replace',
    'cmd_format_range',
    'cmd_freeze',
    'cmd_hyperlink',
    'cmd_name',
    'cmd_open',
    'cmd_print_setup',
    'cmd_printer_list',
    'cmd_printer_setup',
    'cmd_row',
    'cmd_save',
    'cmd_save_as',
    'cmd_shape',
    'cmd_sheet',
    'cmd_sort',
    'cmd_table',
    'cmd_tidy',
    'cmd_audit',
    'audit_table',
    'cmd_validation',
    'cmd_write_cells',
    'cmd_write_range',
]


def _cell_kind(v):
    """セル 1 つの型（'数値' / '文字' / None＝空）。日付は数値に寄せない（別扱い）。"""
    import datetime
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, bool):
        return '文字'
    if isinstance(v, (datetime.datetime, datetime.date)):
        return '日付'
    if isinstance(v, (int, float)):
        return '数値'
    return '文字'


def _looks_like_key(label):
    """突き合わせのキーになりやすい見出しか（番号・コード・ID）。"""
    s = str(label or '').strip()
    if not s or len(s) > 20:
        return False
    for w in ('番号', 'ＩＤ', 'id', 'ID', 'Id', 'No', 'no', 'ＮＯ', '記号'):
        if w in s:
            return True
    return bool(_ID_KW_KANA.search(s))


def key_type_mismatch(values):
    """同じ見出しの列が 2 か所以上にあって、下の値の型が食い違う組を返す（純 Python）。

    values は 2 次元（行 × 列）。戻り値は説明の文字列の一覧。
    VLOOKUP は文字の "2001" と数値の 2001 を別物として扱うので、これが全行「見つからない」の正体。
    """
    if not values:
        return []
    rows = len(values)
    cols = max(len(r) for r in values)
    spots = {}
    for r in range(min(rows, 200)):
        row = values[r]
        for c in range(min(len(row), 60)):
            v = row[c]
            if not isinstance(v, str) or not _looks_like_key(v):
                continue
            kinds = []
            for rr in range(r + 1, min(rows, r + 8)):
                k = _cell_kind(values[rr][c] if c < len(values[rr]) else None)
                if k:
                    kinds.append(k)
                if len(kinds) >= 5:
                    break
            if kinds:
                spots.setdefault(v.strip(), []).append((r, c, kinds[0], len(kinds)))
    out = []
    for label, hits in spots.items():
        seen = {}
        for r, c, kind, _n in hits:
            seen.setdefault(kind, []).append((r, c))
        if len(seen) < 2:
            continue
        parts = []
        for kind, places in seen.items():
            r, c = places[0]
            parts.append(f"{_a1(r + 2, c + 1)}＝{kind}")
        out.append(f"「{label}」の列で型が違います（{' / '.join(parts)}）"
                   " → VLOOKUP は文字の \"2001\" と数値の 2001 を別物として扱います。"
                   "突き合わせる前にキーの型を揃えてください（normalize の as_text／as_number か、"
                   "数式側で TEXT・VALUE を噛ませる）")
    return out


def _a1(row, col):
    """(行, 列) → A1 の番地（1 始まり）。"""
    s = ''
    while col > 0:
        col, m = divmod(col - 1, 26)
        s = chr(65 + m) + s
    return f"{s}{row}"
