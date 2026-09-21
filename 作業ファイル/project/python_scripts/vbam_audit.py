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
_RE_SINGLE_CELL_REF = re.compile(
    r'(?<![A-Za-z0-9_\$:\'\"\(])\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})\b(?!\s*[:\(])',
    re.IGNORECASE
)


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
        return any(str(values[rr][k] or "").strip() in ("小計", "計", "合計", "総計", "累計") for k in range(num_cols))

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
            for op, val_str in _RE_HARDCODED_CONST.findall(f):
                try:
                    val_num = float(val_str)
                except ValueError:
                    continue
                # 0 や 1 以外の実質的な定数計算を警告
                if val_num not in (0.0, 1.0):
                    issues.append({
                        "cell": addr,
                        "severity": "warning",
                        "type": "hardcoded_constant",
                        "msg": f"数式内に固定値 '{op}{val_str}' が直書きされています（マスターセル参照推奨）",
                        "formula": f
                    })
                    break

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
                        if (0 <= omitted_below < num_rows and omitted_below != r and not running
                                and not _is_subtotal(omitted_below, sum_col)):
                            val_below = values[omitted_below][sum_col]
                            if isinstance(val_below, (int, float)) and not isinstance(val_below, bool):
                                target_cell = _rowcol_to_a1(start_row + omitted_below, start_col + sum_col)
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
                        omitted_above = top_r - 1
                        if 0 <= omitted_above < num_rows and omitted_above != r and not _is_subtotal(omitted_above, sum_col):
                            val_above = values[omitted_above][sum_col]
                            if isinstance(val_above, (int, float)) and not isinstance(val_above, bool):
                                target_cell = _rowcol_to_a1(start_row + omitted_above, start_col + sum_col)
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
                        if 0 <= omitted_right < num_cols and omitted_right != c:
                            val_right = values[sum_row][omitted_right]
                            if isinstance(val_right, (int, float)) and not isinstance(val_right, bool):
                                target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_right)
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
                        if 0 <= omitted_left < num_cols and omitted_left != c:
                            val_left = values[sum_row][omitted_left]
                            if isinstance(val_left, (int, float)) and not isinstance(val_left, bool):
                                target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_left)
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

    # 3. 列内の非一貫数式（Inconsistent Formulas）の検知
    # 同じ列に3個以上の数式があり、大多数が同じR1C1形式なのに、一部だけ異なるR1C1または定数の場合
    for c in range(num_cols):
        col_patterns = {}
        for r in range(num_rows):
            f_r1c1 = str(formulas_r1c1[r][c] or "")
            if f_r1c1.startswith("="):
                col_patterns.setdefault(f_r1c1, []).append(r)

        if len(col_patterns) > 1:
            total_formula_cells = sum(len(rows) for rows in col_patterns.values())
            if total_formula_cells >= 3:
                # 最多パターンを特定
                majority_pat, majority_rows = max(col_patterns.items(), key=lambda item: len(item[1]))
                if len(majority_rows) >= total_formula_cells * 0.7:
                    # 少数派のセルを非一貫性として警告
                    for pat, rows in col_patterns.items():
                        if pat != majority_pat:
                            for r in rows:
                                addr = _rowcol_to_a1(start_row + r, start_col + c)
                                # warning＝気づき。9/9 の数式の目（formula_notes）でも「列の中で形の違う式」は気づき扱い。
                                # 小計行・最終行の式は正当に形が違うので、critical で agent の done を止めない（2026-09-17）
                                issues.append({
                                    "cell": addr,
                                    "severity": "warning",
                                    "type": "inconsistent_formula",
                                    "msg": f"同列の標準数式パターンと異なります（コピペミス疑い）",
                                    "formula": str(formulas[r][c])
                                })

    return issues


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

    for r in range(num_rows):
        for c in range(num_cols):
            v = values[r][c]
            addr = _rowcol_to_a1(start_row + r, start_col + c)

            if isinstance(v, str):
                # A. 隠れ空白・前後の空白検知
                trimmed = v.strip()
                if v != trimmed and len(v) > 0:
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
                if _RE_TEXT_NUMBER.match(trimmed) and len(trimmed) <= 15:
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
        return False

    total_rows = {r for r in range(num_rows) if _is_total_row(r)}
    for c in range(num_cols):
        num_entries = []
        for r in range(num_rows):
            if r in total_rows:
                continue
            v = values[r][c]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                num_entries.append((r, float(v)))

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

                for r, v in num_entries:
                    if v < lower_bound or v > upper_bound:
                        addr = _rowcol_to_a1(start_row + r, start_col + c)
                        issues.append({
                            "cell": addr,
                            "severity": "warning",
                            "type": "outlier_value",
                            "msg": f"値 {v:,.4g} は同列の中央値 ({median:,.4g}) から著しく乖離しています（桁違い・マイナス誤入力疑い）",
                            "value": v
                        })

    return issues


def format_audit_report(
    sheet_name: str,
    range_addr: str,
    issues: List[Dict[str, Any]]
) -> str:
    """監査結果をエグゼクティブ健全性レポートとして美しく整形する"""
    n_crit = sum(1 for i in issues if i["severity"] == "critical")
    n_warn = sum(1 for i in issues if i["severity"] == "warning")
    n_info = sum(1 for i in issues if i["severity"] == "info")

    # 100点満点のスコア算出
    penalty = (n_crit * 15) + (n_warn * 5) + (n_info * 2)
    score = max(0, 100 - penalty)

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
        for it in warns:
            lines.append(f"  • {it['cell']}: {it['msg']}")
            if "formula" in it:
                lines.append(f"      数式: {it['formula']}")
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
    if crits:
        lines.append("  1. 集計範囲漏れ・数式エラーを優先して数式を修正してください。")
    if any(i["type"] == "hardcoded_constant" for i in warns):
        lines.append("  2. 数式内の固定値は、別セル（設定マスタ列等）を参照するようリファクタリングを推奨します。")
    if any(i["type"] == "hidden_whitespace" for i in infos):
        lines.append("  3. 前後の空白は TRIM 関数または Python クレンジングで一括除去してください。")
    if any(i["type"] == "text_number" for i in warns):
        lines.append("  4. 文字列型数字は VALUE 関数または数値変換で統一してください。")
    lines.append("  👉 `--fix` オプションを付与して実行すると、上記のうち数式漏れ・空白・文字列数字を安全に一括自動修復できます。")

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


def apply_audit_fixes(ws: Any, issues: List[Dict[str, Any]]) -> Dict[str, Any]:
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
    all_issues = f_issues + d_issues

    # --fix: 人の値を書き換える手＝ -y（依頼に承認の語があるときだけ付ける）と控えを要る（2026-09-17）
    fix_result = None
    if getattr(args, 'fix', False):
        if not getattr(args, 'yes', False):
            print("--fix は人の値を書き換えます。依頼に「直してよい・置き換えてよい」等の承認があるときだけ -y を付けて撃ってください。"
                  "今回は診断だけ出します。")
        else:
            backup = _diagnose_backup(wb)
            if backup:
                print(f"控え: {backup}")
            fix_result = apply_audit_fixes(ws, all_issues)

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
    'generate_audit_html',
    'to_2d_list',
]
