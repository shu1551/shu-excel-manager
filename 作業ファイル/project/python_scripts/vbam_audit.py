"""Excel シートおよび表のインサイト・データ診断エンジン (vbam_audit.py)

プロのExcel監査役・データアナリストの視点を自動化する：
1. 数式静的解析 (Formula Linter):
   - 集計範囲漏れ (SUM/AVERAGE の直前・直後のセル漏れ)
   - ハードコード定数の直書き検知 (*1.1, +500 等)
   - 列内・行内の非一貫数式 (Inconsistent Formulas: コピペミス・手入力上書き)
   - 数式エラー (#REF!, #DIV/0!, #VALUE! 等)
2. データクレンジング診断 (Data Cleaner):
   - 隠れ空白・前後の全半角スペース・改行 (VLOOKUP破壊の原因)
   - 文字列型数字 ('123 など、SUMから除外されるリスク)
   - 統計的外れ値・桁間違い (IQR法による異常値検出)
3. エグゼクティブ健全性サマリー (Audit Summary):
   - 100点満点スコア、重要度別一覧、推奨アクション
"""

import os
import re
import sys
import math
from typing import List, Dict, Tuple, Any, Optional

from vbam_core import BACKUP_DIR
from vbam_core import (
    get_workbook,
    parse_target_and_rest,
    smart_path_resolve,
    job_clock_note,
)
from vbam_view import _resolve_range, _whole_sheet_spec


# ================================================================
# 純粋ロジック：数式解析・クレンジング・統計
# ================================================================

_RE_HARDCODED_CONST = re.compile(
    r'([*+/-])\s*([0-9]+(?:\.[0-9]+)?|\.[0-9]+)(?![A-Za-z0-9_\$:\'\"\(])'
)

_RE_SUM_FUNC = re.compile(
    r'(SUM|AVERAGE)\s*\(\s*(\$?[A-Za-z]+\$?[0-9]+)\s*:\s*(\$?[A-Za-z]+\$?[0-9]+)\s*\)',
    re.IGNORECASE
)

_RE_TEXT_NUMBER = re.compile(r'^-?\d+(?:\.\d+)?$')

_RE_RANGE_REF = re.compile(
    r'(?<![A-Za-z0-9_\$])\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})\s*:\s*\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?![A-Za-z0-9_])',
    re.IGNORECASE
)
# 2026-11-05 dock の実測（Enron の公開ブック 60 本）: 循環参照 38,153 件がほぼ全部誤爆だった。
# ='[外のブック.xls]Nominations'!I$23 や =Credit!A1 の「別シートの I23」を自分のシートの I23 と取り違えていた。
# 式の中の文字列・別シート／別ブックの番地を先に消してから、このシートの番地だけを拾う。
# 前は「(」の直後の番地（=ROUND(A1,2) の A1）も拾っていなかった＝本物の循環の見落とし。
_RE_SINGLE_CELL_REF = re.compile(
    r'(?<![A-Za-z0-9_\$:\'\"])\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})\b(?!\s*[:\(!])',
    re.IGNORECASE
)
_RE_STRING_LIT = re.compile(r'"(?:[^"]|"")*"')
_RE_QUALIFIED_REF = re.compile(
    r"(?:'(?:[^']|'')+'|\[[^\]]*\][^!\s,()+\-*/^&=<>]*|[A-Za-z0-9_.#]+)!"
    r"\$?[A-Za-z]{1,3}\$?\d+(?:\s*:\s*\$?[A-Za-z]{1,3}\$?\d+)?",
)


_TOTAL_WORDS_JA = ("小計", "計", "合計", "総計", "累計")
_RE_TOTAL_EN = re.compile(r'^(?:grand\s+|sub-?\s*)?totals?\b|\b(?:sub-?)?totals?:?$', re.I)


def _is_total_label(s) -> bool:
    """合計・小計の見出しの文字か。英語の Total・TOTAL:・Grand Total・Total Revenue・Net Total も
    （2026-11-05 dock の実測: 英語のブックの合計の行が外れ値・集計漏れの的になっていた）。"""
    if not isinstance(s, str):
        return False
    t = s.strip()
    return t in _TOTAL_WORDS_JA or bool(_RE_TOTAL_EN.search(t))


# 式に直書きしても咎めない、単位・暦の決まった数（月→年の 12・日の 365・千の 1000 など）
_UNIT_CONSTANTS = {7.0, 12.0, 24.0, 30.0, 31.0, 52.0, 60.0, 100.0, 360.0, 365.0, 366.0, 1000.0, 1440.0, 3600.0,
                   86400.0, 1000000.0}


def local_ref_text(f: str) -> str:
    """式から文字列と、別シート・別ブック（'…'!A1・Sheet2!A1:B3・[x.xls]S!A1・#REF!A1）の番地を消した残り（純 Python）。
    残りに出てくる番地は、このシートのセル。"""
    s = _RE_STRING_LIT.sub('""', str(f or ''))
    return _RE_QUALIFIED_REF.sub('0', s)


def _a1_to_rowcol(addr: str) -> Tuple[int, int]:
    """A1 形式のアドレスを 0-indexed (row, col) に変換"""
    addr = addr.replace('$', '')
    m = re.match(r'^([A-Za-z]+)([0-9]+)$', addr)
    if not m:
        return (0, 0)
    col_str, row_str = m.group(1).upper(), m.group(2)
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return (int(row_str) - 1, col - 1)


def _rowcol_to_a1(row: int, col: int) -> str:
    """0-indexed (row, col) を A1 形式のアドレスに変換"""
    col_num = col + 1
    col_str = ""
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        col_str = chr(ord('A') + rem) + col_str
    return f"{col_str}{row + 1}"


def check_circular_references(
    formulas: List[List[Any]],
    start_row: int,
    start_col: int,
) -> List[Dict[str, Any]]:
    """数式内の循環参照（自己参照および間接循環依存）を有向グラフの閉路検出により静的特定する。"""
    issues = []
    num_rows = len(formulas)
    if num_rows == 0:
        return issues
    num_cols = len(formulas[0])

    adj: Dict[str, set] = {}
    formula_map: Dict[str, str] = {}

    for r in range(num_rows):
        for c in range(num_cols):
            f = str(formulas[r][c] or "")
            if not f.startswith("="):
                continue
            addr = _rowcol_to_a1(start_row + r, start_col + c)
            formula_map[addr] = f
            deps = set()
            f = local_ref_text(f)

            # 範囲参照の展開 (例: A1:A5)
            for m in _RE_RANGE_REF.finditer(f):
                c1_str, r1_str, c2_str, r2_str = m.groups()
                cr1, cc1 = _a1_to_rowcol(f"{c1_str}{r1_str}")
                cr2, cc2 = _a1_to_rowcol(f"{c2_str}{r2_str}")
                min_r, max_r = min(cr1, cr2), max(cr1, cr2)
                min_c, max_c = min(cc1, cc2), max(cc1, cc2)
                # 探索爆発防止（最大1000セルまで展開）
                if (max_r - min_r + 1) * (max_c - min_c + 1) <= 1000:
                    for tr in range(min_r, max_r + 1):
                        for tc in range(min_c, max_c + 1):
                            deps.add(_rowcol_to_a1(tr, tc))

            # 単一セル参照の抽出
            for m in _RE_SINGLE_CELL_REF.finditer(f):
                c_str, r_str = m.groups()
                deps.add(f"{c_str.upper()}{r_str}")

            adj[addr] = deps

    visited: Dict[str, int] = {}  # 0: unvisited, 1: visiting, 2: visited
    reported_cycles = set()

    def dfs(node: str, path: List[str]):
        visited[node] = 1
        for neighbor in adj.get(node, ()):
            if neighbor in path:
                cycle_start_idx = path.index(neighbor)
                cycle_path = path[cycle_start_idx:] + [neighbor]
                cycle_key = tuple(sorted(set(cycle_path)))
                if cycle_key not in reported_cycles:
                    reported_cycles.add(cycle_key)
                    cycle_str = " -> ".join(cycle_path)
                    issues.append({
                        "cell": neighbor,
                        "severity": "critical",
                        "type": "circular_reference",
                        "msg": f"循環参照（循環依存）が検出されました: {cycle_str}",
                        "formula": formula_map.get(neighbor, "")
                    })
            elif visited.get(neighbor, 0) == 0:
                if neighbor in adj:
                    dfs(neighbor, path + [neighbor])
        visited[node] = 2

    for node in list(adj.keys()):
        if visited.get(node, 0) == 0:
            dfs(node, [node])

    return issues


def check_formula_linter(
    formulas: List[List[Any]],
    formulas_r1c1: List[List[Any]],
    values: List[List[Any]],
    start_row: int,
    start_col: int,
) -> List[Dict[str, Any]]:
    """数式の静的解析を行う純粋ロジック。
    formulas / formulas_r1c1 / values は2次元配列（0-indexed）。
    """
    issues = []
    num_rows = len(formulas)
    if num_rows == 0:
        return issues
    num_cols = len(formulas[0])

    def _is_subtotal(rr, cc):
        # 2026-09-17 夜: 小計ごとに SUM する表（小計の直上は前のグループの小計）を「直上の数値が SUM から漏れ」と誤って指摘し、
        # 入口で正しく撃てたマクロの後を AI に回していた＝隣が集計の式・集計の行なら漏れではない
        fx = str(formulas[rr][cc] or "")
        if re.match(r'^=\s*(SUM|SUBTOTAL|AGGREGATE)\s*\(', fx, re.I):
            return True
        return any(_is_total_label(values[rr][k]) for k in range(num_cols))

    def _r1(rr, cc):
        if 0 <= rr < num_rows and 0 <= cc < num_cols and rr < len(formulas_r1c1) and cc < len(formulas_r1c1[rr]):
            return str(formulas_r1c1[rr][cc] or "")
        return ""

    def _copied(rr, cc, vertical_sum):
        """同じ形（R1C1）の式が隣にある＝コピーして並べた式。縦の合計なら左右、横の合計なら上下を見る。
        並べた式がそろって同じ所を外しているのは作り（見出しの年・前の列の別の数）で、漏れではない（2026-11-05 dock の実測）。"""
        me = _r1(rr, cc)
        if not me.startswith("="):
            return False
        nb = ((rr, cc - 1), (rr, cc + 1)) if vertical_sum else ((rr - 1, cc), (rr + 1, cc))
        return any(_r1(a, b) == me for a, b in nb)

    # ほかの式の範囲に入っているセル（=SUM(J8:J10) の下の J11 は、J14=SUM(J11:J13) が足している＝漏れではない。
    # 2026-11-05 dock の実測: 塊ごとに小計を取り、その小計を足し合わせる表で、塊の境目を毎回「漏れ」と出していた）
    covered: Dict[str, set] = {}
    for rr in range(num_rows):
        for cc in range(num_cols):
            fx = str(formulas[rr][cc] or "")
            if not fx.startswith("="):
                continue
            me = _rowcol_to_a1(start_row + rr, start_col + cc)
            for m in _RE_RANGE_REF.finditer(local_ref_text(fx)):
                a_r, a_c = _a1_to_rowcol(m.group(1) + m.group(2))
                b_r, b_c = _a1_to_rowcol(m.group(3) + m.group(4))
                if (abs(b_r - a_r) + 1) * (abs(b_c - a_c) + 1) > 5000:
                    continue
                box = (min(a_r, b_r), max(a_r, b_r), min(a_c, b_c), max(a_c, b_c))
                for tr in range(box[0], box[1] + 1):
                    for tc in range(box[2], box[3] + 1):
                        covered.setdefault(_rowcol_to_a1(tr, tc), set()).add((me, box))

    above_left: set = set()

    def _skip_omitted(val, target_cell, f, me=None):
        """漏れと言わない値: 0（足しても答えが変わらない）・同じ式の中で別に足している番地（=SUM(J6:J36)+J37）・
        ほかの式の範囲が数えている数・それ自身が集計の式のセル。"""
        if val == 0:
            return True
        # 数えているのが「重ならない別の塊」の式のときだけ（累計の列 =SUM($A$2:A4) は隣の行の式と範囲が重なる＝
        # 1 行だけ範囲が短い式は今までどおり漏れとして出す）
        mine = [(a_r, b_r, a_c, b_c) for mm in _RE_RANGE_REF.finditer(local_ref_text(f))
                for a_r, a_c in [_a1_to_rowcol(mm.group(1) + mm.group(2))]
                for b_r, b_c in [_a1_to_rowcol(mm.group(3) + mm.group(4))]]
        mine = [(min(a, b), max(a, b), min(c1, c2), max(c1, c2)) for a, b, c1, c2 in mine]

        def _overlap(x, y):
            return not (x[1] < y[0] or y[1] < x[0] or x[3] < y[2] or y[3] < x[2])
        for who, box in covered.get(target_cell, set()):
            if who != me and not any(_overlap(box, mb) for mb in mine):
                return True
        tr, tc = _a1_to_rowcol(target_cell)
        tr, tc = tr - start_row, tc - start_col
        if 0 <= tr < num_rows and 0 <= tc < num_cols and re.match(
                r'^=\s*(SUM|SUBTOTAL|AGGREGATE)\s*\(', str(formulas[tr][tc] or ""), re.I):
            return True
        # 範囲の上・左の式のセル（=C12+C8 の資産計）は前の塊の計算の答え。範囲の下・右の挟まった式は今までどおり出す
        if (0 <= tr < num_rows and 0 <= tc < num_cols and str(formulas[tr][tc] or "").startswith("=")
                and target_cell in above_left):
            return True
        m = re.match(r'([A-Za-z]+)(\d+)$', target_cell)
        if not m:
            return False
        pat = r'(?<![A-Za-z0-9_$])\$?' + m.group(1) + r'\$?' + m.group(2) + r'(?!\d)'
        return re.search(pat, local_ref_text(f), re.I) is not None

    def _year_or_date(val):
        """見出しに置く年（1998）・日付のシリアル（36951＝2001/02/28）らしい整数か。"""
        return float(val).is_integer() and (1900 <= val <= 2100 or 30000 <= val <= 60000)

    # 1. 各セルの数式検証（定数混入、エラー、集計漏れ）
    for r in range(num_rows):
        for c in range(num_cols):
            f = str(formulas[r][c] or "")
            v = values[r][c]
            addr = _rowcol_to_a1(start_row + r, start_col + c)

            # A. 数式エラーの検知
            v_str = str(v or "")
            if v_str.startswith("#") and v_str.endswith("!"):
                issues.append({
                    "cell": addr,
                    "severity": "critical",
                    "type": "formula_error",
                    "msg": f"数式エラー '{v_str}' が発生しています",
                    "formula": f
                })
                continue

            if not f.startswith("="):
                continue

            # B. ハードコード定数の直書き検知
            # `=A1*1.1` や `+500` などの固定値。ただし `+0` や `*1` 以外
            # 別ブックの道のり（'C:\TEMP\[NYMEX 1990-1999.xls]1998'!G254 の -1999）・文字列の中の数は見ない。
            # 単位・暦の決まった数（/12・/365・/1000）は直書きが普通で、咎めると本物が埋もれる（2026-11-05 dock の実測）
            for op, val_str in _RE_HARDCODED_CONST.findall(local_ref_text(f)):
                try:
                    val_num = float(val_str)
                except ValueError:
                    continue
                # 0 や 1 以外の実質的な定数計算を警告
                if val_num not in (0.0, 1.0) and not (op in '*/' and val_num in _UNIT_CONSTANTS):
                    issues.append({
                        "cell": addr,
                        "severity": "warning",
                        "type": "hardcoded_constant",
                        "msg": f"数式内に固定値 '{op}{val_str}' が直書きされています（マスターセル参照推奨）",
                        "formula": f
                    })
                    break

            # B2. 近似一致の検索（VLOOKUP/HLOOKUP の第 4 引数省略・TRUE、MATCH の第 3 引数省略・1・-1）。
            # 探す値が表に無いとき、エラーにならず別の行の値を返す（2026-10-01・オフィス田中「VLOOKUP の第 4 引数」）
            try:
                import vbam_lookup as _vl
                _ap = _vl.approx_calls(f)
            except Exception:
                _ap = []
            if _ap:
                _fns = '・'.join(sorted({a[0] for a in _ap}))
                issues.append({
                    "cell": addr,
                    "severity": "warning",
                    "type": "approx_match",
                    "msg": f"{_fns} が近似一致（第 4 引数が省略か TRUE・MATCH は第 3 引数が省略か 1・-1）です。探す値が表に無いと、"
                           "エラーにならず別の行の値を返します。完全一致のつもりなら FALSE（MATCH は 0）を書く",
                    "formula": f
                })

            # C. SUM / AVERAGE 集計範囲漏れの検知
            # 例: =SUM(B2:B9) で直下や直上に数値があるのに漏れている
            m_sum = _RE_SUM_FUNC.search(f)
            if m_sum:
                fn_name = m_sum.group(1).upper()
                s_from, s_to = m_sum.group(2), m_sum.group(3)
                r1, c1 = _a1_to_rowcol(s_from)
                r2, c2 = _a1_to_rowcol(s_to)

                # 縦方向の集計の場合 (例: =SUM(B2:B9))
                if c1 == c2:
                    sum_col = c1 - start_col
                    top_r = min(r1, r2) - start_row
                    bot_r = max(r1, r2) - start_row
                    if 0 <= sum_col < num_cols:
                        # 直下のセル漏れチェック（集計セル自身でないこと）
                        # 累計の列（=SUM($B$2:B5) を C5 に置く形＝範囲の終わりが自分の行）は、直下を外すのが仕事＝漏れではない
                        # （2026-09-18: B 集計の「累計の列」を鍛えたら、値は正解と同じなのに毎回ここで止まり AI に回った）
                        running = (bot_r == r)
                        omitted_below = bot_r + 1
                        # 直下の漏れは「範囲の終わりと式の間に挟まった数」だけ（2026-11-05 dock の実測: 離れた所から
                        # =SUM(F17:F18) で塊を取る式に「F19 が漏れ」と出していた）。同じ列なら式より上、
                        # 隣の列（累計の列）なら式と同じ行か 1 つ上まで。
                        between = (omitted_below < r) if c == sum_col else (0 <= r - omitted_below <= 1)
                        if (0 <= omitted_below < num_rows and omitted_below != r and not running and between
                                and not _is_subtotal(omitted_below, sum_col)):
                            val_below = values[omitted_below][sum_col]
                            target_cell = _rowcol_to_a1(start_row + omitted_below, start_col + sum_col)
                            if (isinstance(val_below, (int, float)) and not isinstance(val_below, bool)
                                    and not _skip_omitted(val_below, target_cell, f, addr)):
                                fixed_f = f.replace(s_to, target_cell, 1)
                                issues.append({
                                    "cell": addr,
                                    "severity": "critical",
                                    "type": "omitted_sum_range",
                                    "msg": f"集計範囲直下の数値セル {target_cell} ({val_below}) が {fn_name} から漏れています",
                                    "formula": f,
                                    "fixed_formula": fixed_f
                                })
                        # 直上のセル漏れチェック
                        # 直上の漏れは、同じ列で範囲の下に置いた式だけ。見出しの年・日付と、隣にコピーした同じ形の式が
                        # そろって外している所は作り（2026-11-05 dock の実測: E2=1998 の見出し・移動平均・行ごとの合計）
                        omitted_above = top_r - 1
                        if (0 <= omitted_above < num_rows and omitted_above != r and c == sum_col and r > bot_r
                                and not _is_subtotal(omitted_above, sum_col) and not _copied(r, c, True)):
                            val_above = values[omitted_above][sum_col]
                            target_cell = _rowcol_to_a1(start_row + omitted_above, start_col + sum_col)
                            if (isinstance(val_above, (int, float)) and not isinstance(val_above, bool)
                                    and not (above_left.add(target_cell) or _skip_omitted(val_above, target_cell, f, addr)) and not _year_or_date(val_above)):
                                fixed_f = f.replace(s_from, target_cell, 1)
                                issues.append({
                                    "cell": addr,
                                    "severity": "critical",
                                    "type": "omitted_sum_range",
                                    "msg": f"集計範囲直上の数値セル {target_cell} ({val_above}) が {fn_name} から漏れています",
                                    "formula": f,
                                    "fixed_formula": fixed_f
                                })

                # 横方向の集計の場合 (例: =SUM(B2:E2))
                elif r1 == r2:
                    sum_row = r1 - start_row
                    left_c = min(c1, c2) - start_col
                    right_c = max(c1, c2) - start_col
                    if 0 <= sum_row < num_rows:
                        # 直右のセル漏れチェック
                        omitted_right = right_c + 1
                        if 0 <= omitted_right < num_cols and omitted_right != c and r == sum_row and omitted_right < c:
                            val_right = values[sum_row][omitted_right]
                            target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_right)
                            if (isinstance(val_right, (int, float)) and not isinstance(val_right, bool)
                                    and not _skip_omitted(val_right, target_cell, f, addr)):
                                fixed_f = f.replace(s_to, target_cell, 1)
                                issues.append({
                                    "cell": addr,
                                    "severity": "critical",
                                    "type": "omitted_sum_range",
                                    "msg": f"集計範囲直右の数値セル {target_cell} ({val_right}) が {fn_name} から漏れています",
                                    "formula": f,
                                    "fixed_formula": fixed_f
                                })
                        # 直左のセル漏れチェック
                        omitted_left = left_c - 1
                        if (0 <= omitted_left < num_cols and omitted_left != c and r == sum_row and c > right_c
                                and not _copied(r, c, False)):
                            val_left = values[sum_row][omitted_left]
                            target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_left)
                            if (isinstance(val_left, (int, float)) and not isinstance(val_left, bool)
                                    and not (above_left.add(target_cell) or _skip_omitted(val_left, target_cell, f, addr)) and not _year_or_date(val_left)):
                                fixed_f = f.replace(s_from, target_cell, 1)
                                issues.append({
                                    "cell": addr,
                                    "severity": "critical",
                                    "type": "omitted_sum_range",
                                    "msg": f"集計範囲直左の数値セル {target_cell} ({val_left}) が {fn_name} から漏れています",
                                    "formula": f,
                                    "fixed_formula": fixed_f
                                })

    # 2. 循環参照（Circular Reference）の静的検出
    circ_issues = check_circular_references(formulas, start_row, start_col)
    issues.extend(circ_issues)

    # 3. 隣と違う式（Inconsistent Formulas）の検知
    # 2026-11-05 dock の実測（Enron の公開ブック 60 本）: 列の多数決だけで見ると、表でない計算の格子で
    # 509 件出し Excel の判定と重なるのは 9 件だった。横に並んだ式のずれ（AA=+F6・AB=+F6・AC=+H6 の AB）は
    # 1 件も言えなかった。多数決の少数派のうち「同じ形の式に挟まれた（前後 2 つ以内）」セルだけを言い、
    # 列（縦）と行（横）の両方で見る。
    flagged_inc = set()
    for along in ("col", "row"):
        outer = num_cols if along == "col" else num_rows
        inner = num_rows if along == "col" else num_cols
        for k in range(outer):
            pats = {}
            for t in range(inner):
                r, c = (t, k) if along == "col" else (k, t)
                f_r1c1 = str(formulas_r1c1[r][c] or "") if r < len(formulas_r1c1) and c < len(formulas_r1c1[r]) else ""
                if not f_r1c1.startswith("="):
                    continue
                # 合計・平均の行（同じ列の上を集計する式）・同じ列だけ（横なら同じ行だけ）を足す式は数に入れない
                # （9 行中 2 行のずれを黙って通した 2026-09-24 通しの実測 5・節の小計を拾う総計 2026-11-05）
                if along == "col" and (_is_column_total(f_r1c1, start_row + r + 1) or _is_column_only(f_r1c1)):
                    continue
                if along == "row" and _is_row_total(f_r1c1):
                    continue
                # 番地の無い式（=3100000/2 の電卓）は形を比べない
                if not _RE_R1C1_REF.search(_RE_STRING_LIT.sub('""', f_r1c1)):
                    continue
                pats.setdefault(f_r1c1, []).append(t)
            if len(pats) < 2:
                continue
            pat_of = {t: p for p, ts in pats.items() for t in ts}
            for t, pat in sorted(pat_of.items()):
                # 前と後で一番近い式（2 つ以内）が同じ形 P で、自分の形はこの列（行）にぽつんと（2 つまで）
                # （多数決の 6 割は、別の形の式が混ざる行で AB=+F6 のずれを言えなかった・2026-11-05）
                # 横はすぐ隣だけ（値・前年差の組が横に並ぶ表で、1 つ飛ばしの式どうしを比べていた・e023）
                reach = (1, 2) if along == "col" else (1,)
                before = next((pat_of[t - d] for d in reach if t - d in pat_of), None)
                after = next((pat_of[t + d] for d in reach if t + d in pat_of), None)
                # 自分の形が 3 つ以上ある＝2 種類の式が交互に並ぶ作り（数量・金額の列の繰り返し）で、ずれではない
                # 長い列の数個のずれ（800 行の中の LN(E32/E30) 4 つ）は拾う＝周りの形の 1/5 以下までをぽつんと見る
                if not (before and before == after and before != pat and len(pats[before]) >= 2
                        and len(pats[pat]) <= max(2, len(pats[before]) // 5)):
                    continue
                r, c = (t, k) if along == "col" else (k, t)
                addr = _rowcol_to_a1(start_row + r, start_col + c)
                if addr in flagged_inc:
                    continue
                flagged_inc.add(addr)
                # warning＝気づき（小計行・最終行の式は正当に形が違うので、critical で agent の done を止めない・2026-09-17）
                issues.append({
                    "cell": addr,
                    "severity": "warning",
                    "type": "inconsistent_formula",
                    "msg": ("同列の標準数式パターンと異なります（コピペミス疑い）" if along == "col"
                            else "同じ行の左右の式と形が違います（コピペミス疑い）"),
                    "formula": str(formulas[r][c])
                })

    # 4. 式の列に数字の直書き（上下は =E2*F2 なのに G7 だけ 5000）
    issues.extend(check_values_in_formula_columns(formulas, formulas_r1c1, values, start_row, start_col))
    return issues


def check_values_in_formula_columns(formulas, formulas_r1c1, values, start_row, start_col, min_rows=3):
    """式の列の中に紛れた数字の直書き（純 Python）。

    2026-09-23 の通しの実測: 金額の列（=E2*F2 …）の G7 に 5000 が直書きされていたのを、
    materials も diagnose も言わなかった（上の 3 番は式どうしの形しか比べない）。
    列の本文の式（合計・平均の行を除く）が同じ形で min_rows 行以上あり、その式の行の最初〜最後の間に
    数が直に入ったセルがあれば出す。文字・空のセルは咎めない（備考などの書き込み）。
    """
    out = []
    num_rows = len(formulas or [])
    if not num_rows:
        return out
    num_cols = len(formulas[0])
    for c in range(num_cols):
        pats = {}
        for r in range(num_rows):
            fr = str(formulas_r1c1[r][c] or "") if r < len(formulas_r1c1) and c < len(formulas_r1c1[r]) else ""
            if fr.startswith("=") and not _is_column_total(fr, start_row + r + 1):
                pats.setdefault(fr, []).append(r)
        # 番地を含まない式（=#N/A の置き場・=2866+1200 の電卓）は「式の列」の形にしない（2026-11-05 dock の実測）
        pats = {k: v for k, v in pats.items() if _RE_R1C1_REF.search(_RE_STRING_LIT.sub('""', k))}
        if not pats:
            continue
        top, rows = max(pats.items(), key=lambda kv: len(kv[1]))
        if len(rows) < min_rows:
            continue
        first_a1 = str(formulas[rows[0]][c])
        rows_set = set(rows)
        for r in range(rows[0], rows[-1] + 1):
            f = formulas[r][c]
            if isinstance(f, str) and f.startswith("="):
                continue
            # 同じ形の式に上下から挟まれた数だけ（2 行以内）。同じ形の式が 3 つ、100 行に散っている列の間の数を
            # 「直書き」と言っていた（2026-11-05 dock の実測: 塊ごとに式と入力が入れ替わる表）
            if not (any(r - k in rows_set for k in (1, 2)) and any(r + k in rows_set for k in (1, 2))):
                continue
            v = values[r][c] if r < len(values) and c < len(values[r]) else None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append({
                    "cell": _rowcol_to_a1(start_row + r, start_col + c),
                    "severity": "warning",
                    "type": "value_in_formula_column",
                    "msg": f"式の列に数字 {v:g} が直書きされています（他 {len(rows)} 行は {first_a1} の形）",
                    "formula": str(f),
                })
    return out


# 列は相対（C）と絶対（C7。棚の「表の下に合計行を足す」は =SUM($G$4:$G$9) と書く）の両方を読む
# （2026-09-24 通しの確かめ: 絶対参照の合計を「コピペミス疑い」と言っていた）
_COL_TOTAL_RE = re.compile(r'^=(?:(?:SUM|AVERAGE|COUNTA?|MAX|MIN|MEDIAN)\(|SUBTOTAL\(\d+,)'
                           r'R(\[-\d+\]|\d+)C(?:\d+)?:R(\[-\d+\]|\d+)C(?:\d+)?\)$', re.I)


_RE_R1C1_REF = re.compile(r'(?<![A-Za-z0-9_])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])')


def _is_column_only(f_r1c1):
    """同じ列のセルだけを足し合わせる式（=R17+R33+…・=AVERAGE(H10,H14,H18)・=SUM(R[1]C:R[21]C)）か（純 Python）。
    節ごとの小計を拾う総計・見出しの下に置いた小計で、明細の式と形が違うのは当たり前
    （2026-11-05 dock の実測: 「同列の標準数式パターンと異なります」の多くがこれだった）。別シートの番地を含む式は除く。"""
    s = str(f_r1c1 or '')
    if not s.startswith('=') or '!' in s:
        return False
    s = _RE_STRING_LIT.sub('""', s)
    refs = _RE_R1C1_REF.findall(s)
    return bool(refs) and all(col == '' for _row, col in refs)


_ROW_TOTAL_RE = re.compile(r'^=(?:(?:SUM|AVERAGE|COUNTA?|MAX|MIN|MEDIAN)\(|SUBTOTAL\(\d+,)'
                           r'RC\[-\d+\]:RC\[-\d+\]\)$', re.I)


def _is_row_total(f_r1c1):
    """同じ行の左を集計する式（=SUM(RC[-12]:RC[-1]) の年計の列）か。横の比べで、左の明細と形が違うのは当たり前
    （_is_column_total の横向き）。=RC[-1]*1.1 のような横に伸ばす式は除かない（横のコピーの誤りはそこで起きる）。"""
    return bool(_ROW_TOTAL_RE.match(str(f_r1c1 or '').replace(' ', '')))


def _is_column_total(f_r1c1, abs_row):
    """同じ列の上の範囲だけを集計する式（合計・平均・小計の行）か（R1C1 の字面・純 Python）。

    2026-09-23: G27 =SUM(G6:G25)・G28 =AVERAGE(G6:G25) を「同列の標準数式パターンと異なります（コピペミス疑い）」
    と出していた。明細の =E6*F6 と形が違うのは当たり前で、表の点検を頼んだ人に誤りの印を見せていた。
    """
    m = _COL_TOTAL_RE.match(str(f_r1c1 or '').replace(' ', ''))
    if not m:
        return False

    def row_of(tok):
        return abs_row + int(tok[1:-1]) if tok.startswith('[') else int(tok)

    return row_of(m.group(1)) < abs_row and row_of(m.group(2)) < abs_row


def check_data_cleaner(
    values: List[List[Any]],
    start_row: int,
    start_col: int,
) -> List[Dict[str, Any]]:
    """データクレンジング診断（隠れ空白・文字列数字・外れ値）を行う純粋ロジック。"""
    issues = []
    num_rows = len(values)
    if num_rows == 0:
        return issues
    num_cols = len(values[0])

    def _isnum(x):
        return isinstance(x, (int, float)) and not isinstance(x, bool)

    # 列ごとの数の行（見出しの位置と、文字の数字が「番号の列」かを見るため）
    col_num_rows = [[r for r in range(num_rows) if c < len(values[r]) and _isnum(values[r][c])] for c in range(num_cols)]
    col_textnum = [sum(1 for r in range(num_rows) if c < len(values[r]) and isinstance(values[r][c], str)
                       and _RE_TEXT_NUMBER.match(values[r][c].strip())) for c in range(num_cols)]

    clean_texts: Dict[str, int] = {}
    col_texts = [0] * num_cols
    for _r in range(num_rows):
        for _c in range(min(num_cols, len(values[_r]))):
            _v = values[_r][_c]
            if isinstance(_v, str) and _v.strip():
                col_texts[_c] += 1
                if _v == _v.strip():
                    clean_texts[_v] = clean_texts.get(_v, 0) + 1

    def _is_header_cell(r, c):
        """列の数より上にある見出しのセルか（下に数が 2 つ以上・上に数が無い）。"""
        rows = col_num_rows[c]
        return len(rows) >= 2 and rows[0] > r

    for r in range(num_rows):
        for c in range(num_cols):
            v = values[r][c]
            addr = _rowcol_to_a1(start_row + r, start_col + c)

            if isinstance(v, str):
                # A. 隠れ空白・前後の空白検知
                trimmed = v.strip()
                # 咎めない空白（2026-11-05 dock の実測: 1,447 件の大半がこれ）: 空白だけのセル・字下げ（先頭だけの空白）・
                # 数の列の上の見出し（'UK Gas FIN OPT '）。照合に使うのは本文の値で、見出しの末尾の空白は表の外に響かない
                indent_only = (v.rstrip() == v and not ("\n" in v or "\r" in v)
                               and len(v) - len(v.lstrip(" 　")) >= 2)
                # 照合に響くのは「同じ語が空白なしでもある」か「文字の値が並ぶ列（3 つ以上）」の中だけ。
                # ぽつんと置いたラベル（'Total '・'DATE: '・横に並んだ見出し）の末尾の空白は誰も照合しない
                listed = clean_texts.get(trimmed, 0) > 0 or col_texts[c] >= 3
                quiet = (not trimmed) or indent_only or _is_header_cell(r, c) or not listed or _is_total_label(v)
                if v != trimmed and len(v) > 0 and not quiet:
                    space_type = []
                    if v.startswith(" ") or v.startswith("　"):
                        space_type.append("先頭空白")
                    if v.endswith(" ") or v.endswith("　"):
                        space_type.append("末尾空白")
                    if "\n" in v or "\r" in v:
                        space_type.append("改行")
                    tag = "・".join(space_type) if space_type else "余分な空白"
                    issues.append({
                        "cell": addr,
                        "severity": "info",
                        "type": "hidden_whitespace",
                        "msg": f"値 '{v[:15]}' に{tag}が含まれています（突合・照合ミスの原因）",
                        "value": v
                    })

                # B. 文字列型数字の検知
                # 数値のように見えるが型が string のもの
                # 先頭が 0 の番号（伝票番号 0001・郵便番号 007）は文字で持つのが正しい＝咎めない
                # （2026-09-23 通しの実測 2 で伝票番号 11 セルを「SUM 集計対象外リスク」と並べていた）
                # 列の数より上の見出しの年（'2000'）と、数より文字の数字が多い「番号の列」（勘定科目 '100061'）は咎めない
                # （2026-11-05 dock の実測: 1,215 件のほとんどがこの 2 つ）
                code_col = col_textnum[c] > len(col_num_rows[c])
                if (_RE_TEXT_NUMBER.match(trimmed) and len(trimmed) <= 15
                        and not re.fullmatch(r"0\d+", trimmed) and not code_col and not _is_header_cell(r, c)):
                    issues.append({
                        "cell": addr,
                        "severity": "warning",
                        "type": "text_number",
                        "msg": f"数値 '{trimmed}' が文字列型で保存されています（SUM集計対象外リスク）",
                        "value": v
                    })

    # C. 統計的外れ値・桁間違いの検知 (IQR法)
    # 合計・小計・平均の行は検査に入れない（2026-09-19・月別の表の「合計」の行が 4 列とも外れ値と出た。
    # 合計が明細と桁違いなのは当たり前で、入れると四分位も歪む）
    def _is_total_row(r: int) -> bool:
        for k in range(num_cols):
            s = values[r][k]
            if isinstance(s, str):
                t = re.sub(r"[\s　]+", "", s)
                if t in ("計", "平均", "総平均", "total", "Total", "TOTAL") or t.endswith(("合計", "小計", "総計", "累計")):
                    return True
                if _is_total_label(s) or re.match(r'^(?:net|average|avg)\b', s.strip(), re.I):
                    return True
                # 累計・残高・繰越の行（'CUMULATIVE'・'(EST) BEGINNING BALANCE'・'PRIOR MONTH'・'YTD'）も明細と桁が違って当たり前
                # （2026-11-05 dock の実測: 月のシートの 5 行目の累計を毎月「桁違い」と出していた）
                if (re.search(r'\b(?:cumulative|balance|beginning|ending|opening|closing|prior\s+month|ytd|mtd)\b', s, re.I)
                        or any(w in s for w in ("残高", "繰越", "前月", "期首", "期末"))):
                    return True
        return False

    total_rows = {r for r in range(num_rows) if _is_total_row(r)}
    for c in range(num_cols):
        num_entries = []
        nums_seen = 0
        for r in range(num_rows):
            v = values[r][c]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                nums_seen += 1
                # 列の頭（上に数が無い）の年・日付のシリアルは見出し（1998・36951）。桁違いの的にしない（2026-11-05 dock の実測）
                if nums_seen == 1 and float(v).is_integer() and (1900 <= v <= 2100 or 30000 <= v <= 60000):
                    continue
                if r in total_rows:
                    continue
                num_entries.append((r, float(v)))

        # 正の値の列に、マイナスがぽつんと 1〜2 個（符号の打ち間違い）。囲いの中に収まっていても出す
        # （2026-09-23 通しの実測 2: 数量 1〜10 の列の -5 を、IQR の囲いの中なので黙って通していた）
        head = next((values[r][c] for r in range(num_rows) if isinstance(values[r][c], str) and values[r][c].strip()), "")
        signed = (any(w in str(head) for w in ("増減", "差", "損益", "収支", "残高", "調整", "前年比", "比較"))
                  # 英語の見出し（P&L・Change・Deviation・Variance・Net・Adj）の列のマイナスは当たり前（2026-11-05 dock の実測）
                  or bool(re.search(r'p\s*&\s*l|pnl|chg|change|diff|deviation|variance|\bnet\b|adj|gain|loss|balance|\+/-',
                                    str(head), re.I)))
        # 大きい順・小さい順に並べた列（担当ごとの実績の順位表）は、両端が離れているのも端がマイナスなのも当たり前
        # （2026-11-05 dock の実測: 順位表の 1 位 1,806 万・最下位 -5.7 万を桁違い・符号の打ち間違いと出していた）
        # 並べた列では、隣の値から 10 倍以上の崖がある値だけを桁違いと見る（150 の次の 100,000 は拾う）
        seq = [v for _, v in num_entries]
        sorted_col = len(seq) >= 5 and (all(a >= b for a, b in zip(seq, seq[1:])) or all(a <= b for a, b in zip(seq, seq[1:])))
        cliff = set()
        if sorted_col:
            for i, (r, v) in enumerate(num_entries):
                nb = [num_entries[j][1] for j in (i - 1, i + 1) if 0 <= j < len(num_entries)]
                if v and nb and all(x == 0 or abs(v) >= abs(x) * 10 for x in nb) and any(x for x in nb):
                    cliff.add(r)
        neg = [(r, v) for r, v in num_entries if v < 0]
        pos = [v for _, v in num_entries if v > 0]
        flagged = set()
        # 「ぽつんと」は 2 個まで（2026-11-05 dock の実測: 410 行の値動きの列のマイナス 41 個を全部「符号の打ち間違い」と出した）
        if not signed and not sorted_col and len(pos) >= 4 and 1 <= len(neg) <= min(2, max(1, len(num_entries) // 10)):
            for r, v in neg:
                flagged.add(r)
                issues.append({
                    "cell": _rowcol_to_a1(start_row + r, start_col + c),
                    "severity": "warning",
                    "type": "outlier_value",
                    "msg": f"正の値の列にマイナス {v:,.4g} が {len(neg)} 個だけあります（符号の打ち間違い疑い）",
                    "value": v
                })
        num_entries = [(r, v) for r, v in num_entries if r not in flagged]

        # データが4件以上あれば外れ値検査
        if len(num_entries) >= 4:
            sorted_vals = sorted(v for _, v in num_entries)
            n = len(sorted_vals)
            q1 = sorted_vals[int(n * 0.25)]
            q3 = sorted_vals[int(n * 0.75)]
            iqr = q3 - q1

            if iqr > 0:
                # 著しい外れ値（通常は 1.5*IQR だが、業務Excelでは 3.0*IQR で明らかな桁間違いに絞る）
                lower_bound = q1 - 3.0 * iqr
                upper_bound = q3 + 3.0 * iqr
                median = sorted_vals[int(n * 0.5)]
                # 桁違いは「ぽつんと 1 つだけ離れた値」。囲いの外でも、残りの値のいちばん端から 5 倍以上
                # 離れていなければ出さない（2026-09-23 の通しの実測: 単価 30〜12,000 のふつうの表で
                # トナー 8,500・ホワイトボード 12,000・クリップ 100 個・金額 36,000 の 4 件を出していた）
                inside = [v for _, v in num_entries if lower_bound <= v <= upper_bound]
                hi_in = max(inside) if inside else None
                lo_in = min(inside) if inside else None
                # 桁違いは、残りの値のいちばん端から 10 倍以上（1 桁以上）ぽつんと離れた値だけ。
                # 前は 5 倍と桁（対数）の囲いで見ていて、件数の少ない列では囲いが狭く、単価 45〜150 の列の 380・
                # 個数 0〜40 の列の 200 を「著しく乖離」と出していた（2026-10-08 初見の受注の表）
                def _far(v):
                    if v > upper_bound:
                        return hi_in is not None and hi_in > 0 and v >= hi_in * 10
                    if v < 0 <= (lo_in if lo_in is not None else 0):
                        return not signed                   # 正の列にマイナス＝符号の打ち間違い（増減・差の列は当たり前）
                    return lo_in is not None and v > 0 and v * 10 <= lo_in

                far = [(r, v) for r, v in num_entries if (v < lower_bound or v > upper_bound) and _far(v)]
                # 離れた値が同じ側に 3 つ以上＝「ぽつんと 1 つ」ではなく 2 つ目の集まり（日付と値段が同じ列・損益表の
                # 売上と調整・0.005 と 0.05 の 2 種類の刻み）。桁違いの打ち間違いではない（2026-11-05 dock の実測: 残りの 476 件の大半）
                hi_far = [x for x in far if x[1] > upper_bound]
                lo_far = [x for x in far if x[1] <= upper_bound]
                far = (hi_far if len(hi_far) <= 2 else []) + (lo_far if len(lo_far) <= 2 else [])
                for r, v in far:
                    if sorted_col and r not in cliff:
                        continue
                    # 表の本体から空行 3 つ以上離れて置いた数（表の下の検算・控えの数）は別の塊
                    prev_r = max((rr for rr, _x in num_entries if rr < r), default=None)
                    if prev_r is not None and r - prev_r >= 4:
                        continue
                    # 上の数の和（表の下に見出し無しで置いた合計・検算の数）は外れ値ではない
                    above = [x for rr, x in num_entries if rr < r]
                    run, k = [], r - 1
                    while k >= 0 and isinstance(values[k][c], (int, float)) and not isinstance(values[k][c], bool):
                        run.append(float(values[k][c])); k -= 1
                    if v and any(s_ and abs(s_ - v) <= abs(v) * 0.01 for s_ in (sum(above), sum(run))):
                        continue
                    addr = _rowcol_to_a1(start_row + r, start_col + c)
                    near = hi_in if v > upper_bound else lo_in
                    msg = (f"値 {v:,.4g} は同じ列のほかの値（いちばん近い {near:,.4g}）から 10 倍以上離れています（桁違いの疑い）"
                           if v > 0 and near else
                           f"値 {v:,.4g} は同じ列のほかの値（中央値 {median:,.4g}）から離れています（符号の打ち間違いの疑い）")
                    issues.append({
                        "cell": addr,
                        "severity": "warning",
                        "type": "outlier_value",
                        "msg": msg,
                        "value": v
                    })

    return issues


def audit_score(issues: List[Dict[str, Any]]) -> int:
    """100 点満点の点数（format_audit_report と dock が同じ数を使う）。"""
    n_crit = sum(1 for i in issues if i["severity"] == "critical")
    n_warn_kinds = len({(i["type"], i["msg"]) for i in issues if i["severity"] == "warning"})
    n_info = sum(1 for i in issues if i["severity"] == "info")
    score = max(0, 100 - (n_crit * 15) - (n_warn_kinds * 5) - (n_info * 2))
    return max(score, 75) if n_crit == 0 else score


def format_audit_report(
    sheet_name: str,
    range_addr: str,
    issues: List[Dict[str, Any]]
) -> str:
    """監査結果をエグゼクティブ健全性レポートとして美しく整形する"""
    n_crit = sum(1 for i in issues if i["severity"] == "critical")
    n_warn = sum(1 for i in issues if i["severity"] == "warning")
    n_info = sum(1 for i in issues if i["severity"] == "info")

    # 100点満点のスコア算出。同じ中身の警告（1 列に同じ直書き 6 か所）は 1 つと数える＝同じ原因を 6 回減点しない。
    # 重大欠陥が 0 件なら「重大な欠陥の疑い」の C 以下にはしない（2026-09-24: 警告だけで 65 点・C と出ていた）
    score = audit_score(issues)

    lines = []
    lines.append(f"╔═══════════════════════════════════════════════════════════════╗")
    lines.append(f"║ 📊 Excel データ＆数式 総合監査レポート (Audit Report)         ║")
    lines.append(f"╚═══════════════════════════════════════════════════════════════╝")
    lines.append(f"対象: {sheet_name}!{range_addr}")
    
    # スコアと総合評価
    if score >= 90:
        grade = "A (極めて健全)"
        grade_icon = "🟢"
    elif score >= 75:
        grade = "B (概ね良好・軽微な注意あり)"
        grade_icon = "🟡"
    elif score >= 50:
        grade = "C (要点検・重大な欠陥の疑い)"
        grade_icon = "🟠"
    else:
        grade = "D (危険・直ちに修正が必要)"
        grade_icon = "🔴"

    lines.append(f"総合健全性スコア: {grade_icon} {score}点 / 100点 ── 判定: {grade}")
    lines.append(f"検出件数: 重大欠陥 {n_crit}件 ｜ 警告 {n_warn}件 ｜ クレンジング {n_info}件\n")

    if not issues:
        lines.append("✨ 欠陥や不整合は検出されませんでした。数式・データともに大変綺麗です。")
        return "\n".join(lines)

    # 1. 重大欠陥 (Critical)
    crits = [i for i in issues if i["severity"] == "critical"]
    if crits:
        lines.append("🔴 【要修正 (Critical)】直ちに修正が必要な欠陥:")
        for it in crits:
            lines.append(f"  • {it['cell']}: {it['msg']}")
            if "formula" in it:
                lines.append(f"      数式: {it['formula']}")
        lines.append("")

    # 2. 警告 (Warning)
    warns = [i for i in issues if i["severity"] == "warning"]
    if warns:
        lines.append("🟡 【要確認 (Warning)】誤集計・保守性低下の原因:")
        # 同じ中身の警告（H4〜H9 の *1.1 の直書き 6 本など）は 1 行にまとめる（2026-09-24: 6 行並べて読む量を増やしていた）
        groups = {}
        for it in warns:
            groups.setdefault((it['type'], it['msg']), []).append(it)
        for (_t, msg), its in groups.items():
            cells = [it['cell'] for it in its]
            where = cells[0] if len(cells) == 1 else f"{', '.join(cells[:6])}{' …' if len(cells) > 6 else ''}（{len(cells)} か所）"
            lines.append(f"  • {where}: {msg}")
            if "formula" in its[0]:
                lines.append(f"      数式: {its[0]['formula']}" + ("（先頭の 1 つ）" if len(its) > 1 else ""))
        lines.append("")

    # 3. クレンジング (Info)
    infos = [i for i in issues if i["severity"] == "info"]
    if infos:
        lines.append("⚪ 【クレンジング (Info)】表記ゆれ・照合エラーの原因:")
        for it in infos[:10]: # 最大10件まで表示
            lines.append(f"  • {it['cell']}: {it['msg']}")
        if len(infos) > 10:
            lines.append(f"  … ほか {len(infos) - 10} 件の空白セル")
        lines.append("")

    # 推奨アクション
    lines.append("💡 【推奨アクション】")
    acts = []
    if crits:
        acts.append("集計範囲漏れ・数式エラーを優先して数式を修正してください。")
    if any(i["type"] == "hardcoded_constant" for i in warns):
        acts.append("数式内の固定値は、別セル（設定マスタ列等）を参照するようリファクタリングを推奨します。")
    if any(i["type"] == "hidden_whitespace" for i in infos):
        acts.append("前後の空白は TRIM 関数または Python クレンジングで一括除去してください。")
    if any(i["type"] == "text_number" for i in warns):
        acts.append("文字列型数字は VALUE 関数または数値変換で統一してください。")
    if any(i["type"] in ("outlier", "outlier_value") for i in issues):
        acts.append("外れ値は入力の誤りか本当の値かを人が確かめてください（道具は直しません）。")
    for n, a in enumerate(acts, 1):
        lines.append(f"  {n}. {a}")
    # --fix が書き戻すのは空白と文字列の数字だけ（集計漏れの式は提案だけ）。前は「数式漏れも直す」と書いていた（2026-09-24）
    if any(i["type"] in ("hidden_whitespace", "text_number") for i in issues):
        lines.append("  👉 `--fix`（fix=True）で、前後の空白と文字列の数字を控えを取ってから一括で直せます（数式は直しません）。")

    return "\n".join(lines)


def _text_number_safe_to_convert(trimmed: str) -> bool:
    """文字列の数字を数値にしてよいか（純 Python）。

    先頭ゼロ（0123＝コード・口座・郵便番号）は数値にすると桁が消えるので触らない。
    "0" と "0.5" のように 0 の直後が終わりか小数点なら数。桁が 15 を超えるものも触らない（Excel の精度）。
    """
    if not _RE_TEXT_NUMBER.match(trimmed) or len(trimmed) > 15:
        return False
    digits = trimmed.lstrip('-')
    if len(digits) > 1 and digits[0] == '0' and digits[1] != '.':
        return False
    return True


def id_columns_of(values: List[List[Any]], start_col: int, head_rows: int = 5) -> Dict[str, str]:
    """見出しが番号の列（番号・コード・ID・No・品番…）→ {列文字: 見出し}（純 Python）。

    見出しは列の上から head_rows 行のうち、最初の文字のセル。番号の定義は tidy と同じ _is_id_header 1 つ。
    文字で持っている番号（社員番号 1001 など）は、照合のキーの型をそろえるためにわざと文字のことがある＝数にしない（2026-10-28 点検）。"""
    import unicodedata
    from vbam_edit import _is_id_header
    out = {}
    n_cols = max((len(r) for r in values or []), default=0)
    for c in range(n_cols):
        head = next((values[r][c] for r in range(min(head_rows, len(values)))
                     if c < len(values[r]) and isinstance(values[r][c], str) and values[r][c].strip()), None)
        if head is not None and _is_id_header(unicodedata.normalize('NFKC', head)):
            out[re.sub(r'\d+$', '', _rowcol_to_a1(0, start_col + c))] = head.strip()
    return out


def apply_audit_fixes(ws: Any, issues: List[Dict[str, Any]], id_cols: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    """診断で見つけた物のうち、機械が判断なしに直せる 2 種類だけを書き戻す。

    - 隠れ空白 (hidden_whitespace): 前後の空白・改行を除いて書き戻す
    - 文字列の数字 (text_number): 数値にして書き戻す（先頭ゼロ・16 桁以上は触らない）
    数式のセルは触らない（値が式の結果なら式を直すのが筋＝人と AI の判断）。
    集計漏れ (omitted_sum_range) は fixed_formula を **提案として並べるだけ** で書かない
    （SUM の直下の数値が本当に本文なのか、別の塊の頭なのかは判断＝Python に持たせない・2026-09-17）。
    呼び手（cmd_diagnose）は -y と控えを要求する。
    """
    fixed_counts = {
        "omitted_sum_range": 0,
        "hidden_whitespace": 0,
        "text_number": 0,
        "total": 0
    }
    details = []
    suggestions = []
    skipped = []

    # 先頭ゼロの文字の数字（007）がある列は番号の列＝同じ列の 2001 も文字のまま残す
    # （片方だけ数値にすると、同じ列に文字と数値が混ざる・2026-09-21 利用者の報告）
    def _col_of(addr):
        m = re.match(r'\$?([A-Za-z]+)\$?\d', str(addr).split('!')[-1])
        return m.group(1).upper() if m else None
    code_cols = {_col_of(it.get("cell")) for it in issues
                 if it.get("type") == "text_number" and isinstance(it.get("value"), str)
                 and not _text_number_safe_to_convert(it["value"].strip())}
    code_cols.discard(None)

    for item in issues:
        itype = item.get("type")
        cell_addr = item.get("cell")
        if not cell_addr:
            continue

        if itype == "omitted_sum_range" and "fixed_formula" in item:
            suggestions.append(f"  ・{cell_addr}: {item.get('formula', '')} → {item['fixed_formula']}（提案・書いていない）")
            continue
        if itype not in ("hidden_whitespace", "text_number"):
            continue

        try:
            cell_obj = ws.Range(cell_addr)
        except Exception:
            continue
        try:
            if getattr(cell_obj, 'HasFormula', False):
                skipped.append(f"  ・{cell_addr}: 数式のセル＝触らない")
                continue
        except Exception:
            pass

        val = item.get("value")
        if not isinstance(val, str):
            continue

        if itype == "hidden_whitespace":
            cleaned = val.strip()
            cell_obj.Value = cleaned
            fixed_counts["hidden_whitespace"] += 1
            fixed_counts["total"] += 1
            details.append(f"  ・{cell_addr} 空白除去: '{val}' → '{cleaned}'")

        elif itype == "text_number":
            trimmed = val.strip()
            if not _text_number_safe_to_convert(trimmed):
                skipped.append(f"  ・{cell_addr}: '{trimmed}' は先頭ゼロか桁が多い＝文字のまま残す")
                continue
            if _col_of(cell_addr) in code_cols:
                skipped.append(f"  ・{cell_addr}: '{trimmed}' は先頭ゼロの番号と同じ列＝文字のまま残す（列の型を揃える）")
                continue
            if id_cols and _col_of(cell_addr) in id_cols:
                skipped.append(f"  ・{cell_addr}: '{trimmed}' は番号の列（見出し「{id_cols[_col_of(cell_addr)]}」）＝文字のまま残す"
                               "（照合のキーの型をそろえるため、わざと文字のことがある）")
                continue
            try:
                num_val = float(trimmed) if "." in trimmed else int(trimmed)
            except ValueError:
                continue
            cell_obj.Value = num_val
            fixed_counts["text_number"] += 1
            fixed_counts["total"] += 1
            details.append(f"  ・{cell_addr} 数値化: '{trimmed}' → {num_val}")

    return {
        "fixed_counts": fixed_counts,
        "details": details,
        "suggestions": suggestions,
        "skipped": skipped,
    }


def generate_audit_html(
    sheet_name: str,
    range_addr: str,
    issues: List[Dict[str, Any]],
    fix_result: Optional[Dict[str, Any]] = None
) -> str:
    """監査レポートを美麗なモダンHTMLダッシュボード（単一スタンドアロン）として生成する"""
    import datetime
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n_crit = sum(1 for i in issues if i["severity"] == "critical")
    n_warn = sum(1 for i in issues if i["severity"] == "warning")
    n_info = sum(1 for i in issues if i["severity"] == "info")

    penalty = (n_crit * 15) + (n_warn * 5) + (n_info * 2)
    score = max(0, 100 - penalty)

    if score >= 90:
        grade = "A (極めて健全)"
        color = "#10B981"
        badge_bg = "rgba(16, 185, 129, 0.15)"
    elif score >= 75:
        grade = "B (概ね良好・軽微な注意)"
        color = "#F59E0B"
        badge_bg = "rgba(245, 158, 11, 0.15)"
    elif score >= 50:
        grade = "C (要点検・重大な欠陥疑い)"
        color = "#F97316"
        badge_bg = "rgba(249, 115, 22, 0.15)"
    else:
        grade = "D (危険・直ちに修正が必要)"
        color = "#EF4444"
        badge_bg = "rgba(239, 68, 68, 0.15)"

    # SVG円周: 半径45 → 円周 2 * pi * 45 ≈ 282.7
    circumference = 282.7
    dash_offset = circumference - (circumference * score / 100)

    # 項目別HTML
    crit_cards = []
    for it in issues:
        if it["severity"] == "critical":
            f_html = f'<div class="code-box">{it["formula"]}</div>' if "formula" in it else ""
            fix_html = f'<div class="fix-hint">推奨修正: <code>{it["fixed_formula"]}</code></div>' if "fixed_formula" in it else ""
            crit_cards.append(f'''
            <div class="issue-card crit">
                <div class="issue-head"><span class="badge crit-bg">{it["cell"]}</span> <span class="issue-title">{it["msg"]}</span></div>
                {f_html}
                {fix_html}
            </div>
            ''')

    warn_cards = []
    for it in issues:
        if it["severity"] == "warning":
            f_html = f'<div class="code-box">{it["formula"]}</div>' if "formula" in it else ""
            warn_cards.append(f'''
            <div class="issue-card warn">
                <div class="issue-head"><span class="badge warn-bg">{it["cell"]}</span> <span class="issue-title">{it["msg"]}</span></div>
                {f_html}
            </div>
            ''')

    info_cards = []
    for it in issues:
        if it["severity"] == "info":
            info_cards.append(f'''
            <div class="issue-card info">
                <div class="issue-head"><span class="badge info-bg">{it["cell"]}</span> <span class="issue-title">{it["msg"]}</span></div>
            </div>
            ''')

    fix_banner = ""
    if fix_result:
        fc = fix_result.get("fixed_counts", {})
        details_li = "".join(f"<li>{d}</li>" for d in fix_result.get("details", []))
        fix_banner = f'''
        <div class="fix-banner">
            <h3>🔧 自動修復 (Auto-Fix) 完了</h3>
            <p>合計 <strong>{fc.get("total", 0)}</strong> 箇所を安全に自動修復しました。</p>
            <ul>{details_li}</ul>
        </div>
        '''

    html = f'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel 健全性診断書 - {sheet_name}!{range_addr}</title>
<style>
  :root {{
    --bg: #0F172A;
    --card: #1E293B;
    --card-border: #334155;
    --text: #F8FAFC;
    --text-muted: #94A3B8;
    --score-color: {color};
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    padding: 32px 20px;
    line-height: 1.6;
  }}
  .container {{ max-width: 960px; margin: 0 auto; }}
  .header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid var(--card-border);
    padding-bottom: 20px;
    margin-bottom: 24px;
  }}
  .header h1 {{ font-size: 24px; font-weight: 700; }}
  .header .meta {{ color: var(--text-muted); font-size: 14px; }}
  .dashboard {{
    display: grid;
    grid-template-columns: 280px 1fr;
    gap: 24px;
    margin-bottom: 32px;
  }}
  .score-card {{
    background: var(--card);
    border: 1px solid var(--card-border);
    border-radius: 16px;
    padding: 24px;
    text-align: center;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
  }}
  .gauge-svg {{ width: 140px; height: 140px; transform: rotate(-90deg); }}
  .gauge-bg {{ fill: none; stroke: var(--card-border); stroke-width: 8; }}
  .gauge-bar {{
    fill: none;
    stroke: var(--score-color);
    stroke-width: 8;
    stroke-linecap: round;
    stroke-dasharray: {circumference};
    stroke-dashoffset: {dash_offset};
    transition: stroke-dashoffset 1s ease;
  }}
  .score-num {{ font-size: 40px; font-weight: 800; color: var(--score-color); margin-top: 10px; }}
  .score-label {{ font-size: 14px; color: var(--text-muted); margin-bottom: 8px; }}
  .grade-badge {{
    background: {badge_bg};
    color: var(--score-color);
    padding: 6px 14px;
    border-radius: 20px;
    font-weight: 600;
    font-size: 14px;
  }}
  .summary-card {{
    background: var(--card);
    border: 1px solid var(--card-border);
    border-radius: 16px;
    padding: 24px;
    display: flex;
    flex-direction: column;
    justify-content: space-around;
  }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 16px; }}
  .stat-box {{ background: rgba(15, 23, 42, 0.6); padding: 16px; border-radius: 12px; border: 1px solid var(--card-border); }}
  .stat-num {{ font-size: 28px; font-weight: 700; }}
  .stat-crit {{ color: #EF4444; }}
  .stat-warn {{ color: #F59E0B; }}
  .stat-info {{ color: #94A3B8; }}
  .section {{ margin-bottom: 32px; }}
  .section-title {{ font-size: 18px; font-weight: 700; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }}
  .issue-card {{
    background: var(--card);
    border-left: 4px solid var(--card-border);
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 12px;
  }}
  .issue-card.crit {{ border-left-color: #EF4444; }}
  .issue-card.warn {{ border-left-color: #F59E0B; }}
  .issue-card.info {{ border-left-color: #64748B; }}
  .badge {{ font-family: monospace; font-size: 12px; padding: 2px 8px; border-radius: 4px; font-weight: 600; }}
  .crit-bg {{ background: #EF4444; color: #FFF; }}
  .warn-bg {{ background: #F59E0B; color: #000; }}
  .info-bg {{ background: #64748B; color: #FFF; }}
  .code-box {{
    background: rgba(0, 0, 0, 0.4);
    border-radius: 6px;
    padding: 8px 12px;
    font-family: monospace;
    font-size: 13px;
    color: #38BDF8;
    margin-top: 8px;
  }}
  .fix-hint {{ margin-top: 6px; font-size: 13px; color: #10B981; }}
  .fix-banner {{
    background: rgba(16, 185, 129, 0.15);
    border: 1px solid #10B981;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 24px;
  }}
  .fix-banner h3 {{ color: #10B981; margin-bottom: 8px; }}
  .fix-banner ul {{ margin-left: 20px; margin-top: 8px; font-size: 14px; }}
  .footer {{ text-align: center; color: var(--text-muted); font-size: 12px; margin-top: 40px; }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <div>
      <h1>📊 Excel 総合健全性診断書</h1>
      <div class="meta">対象: <strong>{sheet_name}!{range_addr}</strong> ｜ 検査日時: {now_str}</div>
    </div>
    <div class="grade-badge">{grade}</div>
  </div>

  {fix_banner}

  <div class="dashboard">
    <div class="score-card">
      <svg class="gauge-svg" viewBox="0 0 100 100">
        <circle class="gauge-bg" cx="50" cy="50" r="45"></circle>
        <circle class="gauge-bar" cx="50" cy="50" r="45"></circle>
      </svg>
      <div class="score-num">{score}</div>
      <div class="score-label">健全性スコア (100点満点)</div>
    </div>
    <div class="summary-card">
      <h2>診断サマリ</h2>
      <p style="color: var(--text-muted); font-size: 14px;">プロのExcel監査役・データアナリストによる静的解析およびクレンジング結果です。</p>
      <div class="stat-grid">
        <div class="stat-box">
          <div class="stat-num stat-crit">{n_crit}</div>
          <div style="font-size: 13px; color: var(--text-muted);">要修正 (Critical)</div>
        </div>
        <div class="stat-box">
          <div class="stat-num stat-warn">{n_warn}</div>
          <div style="font-size: 13px; color: var(--text-muted);">要確認 (Warning)</div>
        </div>
        <div class="stat-box">
          <div class="stat-num stat-info">{n_info}</div>
          <div style="font-size: 13px; color: var(--text-muted);">クレンジング (Info)</div>
        </div>
      </div>
    </div>
  </div>

  {"".join(f'''<div class="section"><div class="section-title" style="color: #EF4444;">🔴 要修正の重大欠陥 ({len(crit_cards)}件)</div>{"".join(crit_cards)}</div>''' for _ in [1] if crit_cards)}
  {"".join(f'''<div class="section"><div class="section-title" style="color: #F59E0B;">🟡 要確認の警告項目 ({len(warn_cards)}件)</div>{"".join(warn_cards)}</div>''' for _ in [1] if warn_cards)}
  {"".join(f'''<div class="section"><div class="section-title" style="color: #94A3B8;">⚪ 表記ゆれ・クレンジング ({len(info_cards)}件)</div>{"".join(info_cards)}</div>''' for _ in [1] if info_cards)}

  <div class="footer">
    Generated by Antigravity VBA Manager (vbam_audit) ｜ High-Speed Excel Automation Engine
  </div>
</div>
</body>
</html>'''
    return html


def to_2d_list(raw, n_rows=1, n_cols=1):
    """Excel Range から取得した Value / Formula のCOMタプルを安全に2次元リスト [[cell, ...], ...] に変換する。"""
    if n_rows == 1 and n_cols == 1:
        return [[raw]]
    if not isinstance(raw, (tuple, list)):
        return [[raw]]
    result = []
    for row in raw:
        if isinstance(row, (tuple, list)):
            result.append(list(row))
        else:
            result.append([row])
    if n_rows == 1 and len(result) > 1 and not isinstance(raw[0], (tuple, list)):
        return [list(raw)]
    return result


# ================================================================
# CLI コマンドハンドラ: cmd_diagnose
# ================================================================

def _diagnose_backup(wb):
    """--fix の前にブックの写しを backups へ（SaveCopyAs＝本体は無変更・保存もしない）。失敗したら None。"""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        import time as _t
        stem, ext = os.path.splitext(str(wb.Name))
        path = os.path.join(BACKUP_DIR, f"{stem}_before_diagnose_fix_{_t.strftime('%Y%m%d_%H%M%S')}{ext or '.xlsx'}")
        wb.SaveCopyAs(path)
        return path
    except Exception:
        return None


def cmd_diagnose(args):
    """Excel 表・シートの数式＆データ総合監査を実行する: diagnose [excel_file] [範囲|シート名] [--json] [--fix]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)

    spec = rest[0] if rest else None
    if spec is None:
        # 省略時はアクティブシートの使用範囲
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

    # 1セルの場合
    if n_rows == 1 and n_cols == 1:
        rng = rng.CurrentRegion
        rng_addr = rng.Address.replace('$', '')
        n_rows = int(rng.Rows.Count)
        n_cols = int(rng.Columns.Count)

    start_row = int(rng.Row) - 1
    start_col = int(rng.Column) - 1

    # 高速一括読み取り（COMラウンドトリップを1回に集約）
    val_raw = rng.Value2 if hasattr(rng, 'Value2') else rng.Value
    form_raw = rng.Formula
    form_r1c1_raw = rng.FormulaR1C1

    values = to_2d_list(val_raw, n_rows, n_cols)
    formulas = to_2d_list(form_raw, n_rows, n_cols)
    formulas_r1c1 = to_2d_list(form_r1c1_raw, n_rows, n_cols)

    # 静的解析とデータクレンジングの実行
    f_issues = check_formula_linter(formulas, formulas_r1c1, values, start_row, start_col)
    d_issues = check_data_cleaner(values, start_row, start_col)
    # 外れ値は手で入れた値だけ（式のセルは元の値の外れを重ねて言うだけ・2026-09-23）
    fml_cells = {_rowcol_to_a1(start_row + i, start_col + j)
                 for i, row in enumerate(formulas or []) for j, v in enumerate(row or [])
                 if isinstance(v, str) and v.startswith("=")}
    d_issues = [i for i in d_issues if not (i.get("type") == "outlier_value" and i.get("cell") in fml_cells)]
    all_issues = f_issues + d_issues

    # --fix: 人の値を書き換える手＝ -y（依頼に承認の語があるときだけ付ける）と控えを要る（2026-09-17）
    fix_result = None
    if getattr(args, 'fix', False):
        if not getattr(args, 'yes', False):
            print("--fix は人の値を書き換えます。依頼に「直してよい・置き換えてよい」等の承認があるときだけ -y を付けて撃ってください。"
                  "今回は診断だけ出します。")
        else:
            backup = _diagnose_backup(wb)
            if not backup:
                # 「-y と控えを要る」の約束どおり、控えが無ければ書き換えない（前は控えに失敗しても書き換えていた・2026-10-28 点検）
                print("⚠ 控えを取れなかったので、--fix は撃っていません（診断だけ出します）。backups フォルダに書けるか確認してください")
            else:
                print(f"控え: {backup}")
                fix_result = apply_audit_fixes(ws, all_issues, id_columns_of(values, start_col))

    if getattr(args, 'json', False):
        import json
        payload = {
            "sheet": ws.Name,
            "range": rng_addr,
            "rows": n_rows,
            "cols": n_cols,
            "issues": all_issues
        }
        if fix_result is not None:
            payload["fix_result"] = fix_result
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return True

    report = format_audit_report(ws.Name, rng_addr, all_issues)
    print(report)

    if fix_result is not None:
        fc = fix_result["fixed_counts"]
        print(f"\n--fix: {fc['total']} 箇所を書き戻した（空白除去 {fc['hidden_whitespace']}・数値化 {fc['text_number']}）")
        for d in fix_result["details"]:
            print(d)
        if fix_result.get("skipped"):
            print("触らなかったもの:")
            for d in fix_result["skipped"]:
                print(d)
        if fix_result.get("suggestions"):
            print("集計漏れの直し方（書いていない・人か AI が決める）:")
            for d in fix_result["suggestions"]:
                print(d)

    html_out = getattr(args, 'html_out', None)
    if html_out:
        html_content = generate_audit_html(ws.Name, rng_addr, all_issues, fix_result)
        out_path = html_out if html_out != "default" else os.path.join(os.getcwd(), f"audit_{ws.Name}.html")
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"\n🌐 【HTML品質証明書を出力しました】: {out_path}")

    _cn = job_clock_note()
    if _cn:
        print(_cn)
    return True


__all__ = [
    'cmd_diagnose',
    'check_formula_linter',
    'check_circular_references',
    'check_data_cleaner',
    'apply_audit_fixes',
    'format_audit_report',
    'audit_score',
    'generate_audit_html',
    'to_2d_list',
]
