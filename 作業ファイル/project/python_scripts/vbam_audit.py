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

# 先頭のオペランドの直書き: 式の頭か ( , の直後の数に、演算子と、そのあとにセル・関数・括弧が続く形
_RE_HARDCODED_LEAD = re.compile(
    r'(?:^=|[(,])\s*([0-9]+(?:\.[0-9]+)?|\.[0-9]+)\s*([*+/-])\s*(?=[A-Za-z$(])'
)

# 期間の文字: '11月12日～12月25日'・'4/5-10/16'（日数の列の根拠）
_RE_PERIOD_RANGE = re.compile(
    r'(?:(\d{1,2})月\s*(\d{1,2})日?\s*[〜～~\-ー–]\s*(\d{1,2})月\s*(\d{1,2})日?)'
    r'|(?:(?<![\d/])(\d{1,2})/(\d{1,2})\s*[〜～~\-ー–]\s*(\d{1,2})/(\d{1,2})(?![\d/]))'
)

_RE_SUM_FUNC = re.compile(
    r'(SUM|AVERAGE)\s*\(\s*(\$?[A-Za-z]+\$?[0-9]+)\s*:\s*(\$?[A-Za-z]+\$?[0-9]+)\s*\)',
    re.IGNORECASE
)

# 全角のマイナス（－ U+FF0D・− U+2212）と 3 桁区切りのカンマ（半角・全角）も文字の数字（2026-11-06 日本語の注入: 全角のマイナス・カンマは 72/72 見落とし）
_RE_TEXT_NUMBER = re.compile(r'^[-－−]?(?:\d+|\d{1,3}(?:[,，]\d{3})+)(?:[.．]\d+)?$')


def _tn_value(s: str):
    """文字の数字（全角・カンマ・全角マイナス込み）を数にする（純 Python）。"""
    import unicodedata
    t = unicodedata.normalize('NFKC', s.strip()).replace('−', '-').replace(',', '')
    return float(t) if '.' in t else int(t)


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
    r"(?:'(?:[^']|'')+'|\[[^\]]*\][^!\s,()+\-*/^&=<>]*|[\w.#]+)!"
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
    if t in _TOTAL_WORDS_JA or bool(_RE_TOTAL_EN.search(t)):
        return True
    # 日本語の統計表の「大森地区計」「区部計」「総数」（2026-11-05 日本語の公開ブック 60 本: 地区計の行を桁違いと出していた）。
    # 「設計」「会計」「統計」「時計」などの言葉は合計ではない
    t2 = re.sub(r"[\s　]+", "", t)
    if t2 in ("総数", "総額", "全体", "全区", "全市", "全都"):
        return True
    return (len(t2) <= 12 and t2.endswith("計")
            and not t2.endswith(("設計", "会計", "統計", "時計", "家計", "生計", "推計", "集計", "見積計", "計画")))


# 式に直書きしても咎めない、単位・暦の決まった数（月→年の 12・日の 365・千の 1000 など）
_UNIT_CONSTANTS = {7.0, 12.0, 24.0, 30.0, 31.0, 52.0, 60.0, 100.0, 360.0, 365.0, 365.25, 366.0, 1000.0, 1440.0, 3600.0,
                   86400.0, 1000000.0, 10000.0, 100000000.0}   # 10000・1 億は万・億の単位換算
# 割る数の 2 は半分・半期・平均の決まり文句（掛ける 2 は倍率で、直書きの本物がありうるので入れない）
_UNIT_DIV_ONLY = {2.0}
# 検索・位置の関数の引数に付いた +4・+2（列番号・行番号をずらす）は直書きの定数ではなく番地のオフセット
_RE_OFFSET_FUNCS = ('VLOOKUP', 'HLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'OFFSET', 'CHOOSE')


def _enclosing_func(text: str, pos: int) -> str:
    """text[pos] を囲む一番内側の関数名（大文字）。括弧だけなら ''。"""
    depth = 0
    for i in range(pos - 1, -1, -1):
        ch = text[i]
        if ch == ')':
            depth += 1
        elif ch == '(':
            if depth == 0:
                m = re.search(r'([A-Za-z_][A-Za-z0-9_.]*)$', text[:i])
                return m.group(1).upper() if m else ''
            depth -= 1
    return ''


def _after_month_end_func(text: str, pos: int) -> bool:
    """数の直前が EOMONTH(…)・EDATE(…) の閉じ括弧＋演算子か（月末に日数を足す暦のずらし：EOMONTH(A1,0)+20）。"""
    j = pos - 1
    while j >= 0 and text[j] in ' +-':
        j -= 1
    if j < 0 or text[j] != ')':
        return False
    depth = 0
    for i in range(j, -1, -1):
        if text[i] == ')':
            depth += 1
        elif text[i] == '(':
            depth -= 1
            if depth == 0:
                return re.search(r'(?:EOMONTH|EDATE)$', text[:i], re.I) is not None
    return False


def _is_calendar_month_mix(text: str, val_num: float) -> bool:
    """28 を、30・31 の日数と並べて掛けている月別日数の式（O*31+P*28+Q*31+R*30…）か。"""
    return val_num == 28.0 and bool(re.search(r'[*/]\s*3[01](?![0-9.])', text))


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


_RE_PURE_AGG = re.compile(r'^=\s*(SUM|SUBTOTAL|AGGREGATE)\s*\((?:[^()]|\([^()]*\))*\)\s*$', re.I)
_RE_AGG_RANGE = re.compile(r'\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)')


def _is_vertical_agg(fx: str) -> bool:
    """=SUM(D6:D22)・=SUBTOTAL(9,D6:D22) のように、式の全体が集計の関数 1 つで、縦（同じ列）の範囲を足すか（純 Python）。
    =SUM(D71:E71)*0.1/12 のような計算の一部や、=SUM(C41:Q41) の行の合計（横）は、列の小計ではない
    （2026-11-05 注入テスト: 行の合計の式を『小計』と見て、その真上の集計漏れを黙らせていた）。"""
    if not _RE_PURE_AGG.match(fx or ""):
        return False
    rngs = _RE_AGG_RANGE.findall(local_ref_text(fx))
    if not rngs:
        return True
    return any(r1 != r2 for _a, r1, _c, r2 in rngs)



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
        if _is_vertical_agg(fx):
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

    # 式ごとの「1 つのセルとして指している番地」と、その逆引き（漏れと言う相手が別の式で集計のセルと一緒に使われていないかを見る）
    formula_singles: Dict[str, set] = {}
    ref_index: Dict[str, List[str]] = {}
    for rr in range(num_rows):
        for cc in range(num_cols):
            fx = str(formulas[rr][cc] or "")
            if not fx.startswith("="):
                continue
            me_a = _rowcol_to_a1(start_row + rr, start_col + cc)
            singles = {m.group(1).upper() + m.group(2) for m in _RE_SINGLE_CELL_REF.finditer(local_ref_text(fx))}
            if singles:
                formula_singles[me_a] = singles
                for t_a in singles:
                    ref_index.setdefault(t_a, []).append(me_a)

    def _sibling_shares_bound(rr, cc, side, left_c, right_c):
        """同じ列の上下 12 マス以内の、同じ行の横集計（=SUM(D6:T6)）が、同じ側の端の列を使っているか。
        並べた行がそろって同じ列から（まで）足しているのは作り（'前年比' の列を外す）で、1 行だけの漏れではない。"""
        for d in (-12, -11, -10, -9, -8, -7, -6, -5, -4, -3, -2, -1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
            k = rr + d
            if not (0 <= k < num_rows):
                continue
            m = _RE_SUM_FUNC.search(str(formulas[k][cc] or ""))
            if not m:
                continue
            q1, qc1 = _a1_to_rowcol(m.group(2))
            q2, qc2 = _a1_to_rowcol(m.group(3))
            if q1 != q2:
                continue
            sl, sr = min(qc1, qc2) - start_col, max(qc1, qc2) - start_col
            if (side == "left" and sl == left_c) or (side == "right" and sr == right_c):
                return True
        return False

    above_left: set = set()

    def _added_with(who_a1, a, b):
        """式 who_a1 が、番地 a と b を足し引きでつないでいるか（=F5+F38・=SUM(F5,F38)）。割る・掛けるは違う
        （=ROUND(B15/$B$16,3) は B15 が B16 の合計の中の 1 つであることを示す＝漏れの裏付け）。"""
        i_, j_ = _a1_to_rowcol(who_a1)
        i_, j_ = i_ - start_row, j_ - start_col
        if not (0 <= i_ < num_rows and 0 <= j_ < num_cols):
            return True
        t = local_ref_text(str(formulas[i_][j_] or "")).replace("$", "").upper()
        a, b = a.upper(), b.upper()
        for x, y in ((a, b), (b, a)):
            if re.search(r'(?<![A-Z0-9])' + x + r'(?!\d)\s*[-+,]\s*' + y + r'(?!\d)', t):
                return True
        return False

    def _skip_omitted(val, target_cell, f, me=None, orient="v"):
        """漏れと言わない値: 0（足しても答えが変わらない）・同じ式の中で別に足している番地（=SUM(J6:J36)+J37）・
        ほかの式の範囲が数えている数・それ自身が集計の式のセル。"""
        if val == 0:
            return True
        if me is not None:
            # 別の式が、漏れと言う相手とこの集計のセルを一緒に足している（=F5+F38）＝相手は集計の外で足す前の残高
            if any(who != me and me in formula_singles.get(who, ()) and _added_with(who, target_cell, me)
                   for who in ref_index.get(target_cell, ())):
                return True
            # 集計する範囲の式が、その相手を入力に使っている（AA40=AA38*AA39 を足す SUM(AA40:AA41) の AA39）＝足すと二重になる元の数
            for mm in _RE_RANGE_REF.finditer(local_ref_text(f)):
                q1, qc1 = _a1_to_rowcol(mm.group(1) + mm.group(2))
                q2, qc2 = _a1_to_rowcol(mm.group(3) + mm.group(4))
                if (abs(q2 - q1) + 1) * (abs(qc2 - qc1) + 1) > 400:
                    continue
                for tr_ in range(min(q1, q2), max(q1, q2) + 1):
                    for tc_ in range(min(qc1, qc2), max(qc1, qc2) + 1):
                        if target_cell in formula_singles.get(_rowcol_to_a1(tr_, tc_), ()):
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
            # 縦の集計の相手を数えているのが 1 行だけの範囲（=SUM(D251:E251) の行の合計）なら、列の小計を数えたものではない
            # （横の集計では 1 列だけの範囲が同じ）。2026-11-05 注入テストの見落とし 13 件
            if orient == "v" and box[0] == box[1] and box[2] != box[3]:
                continue
            if orient == "h" and box[2] == box[3] and box[0] != box[1]:
                continue
            if who != me and not any(_overlap(box, mb) for mb in mine):
                return True
        tr, tc = _a1_to_rowcol(target_cell)
        tr, tc = tr - start_row, tc - start_col
        # （縦の集計で、相手が行の合計（=SUM(D621:E621)）や計算の一部の式なら小計ではない）
        if 0 <= tr < num_rows and 0 <= tc < num_cols and (
                _is_vertical_agg(str(formulas[tr][tc] or "")) if orient == "v"
                else re.match(r'^=\s*(SUM|SUBTOTAL|AGGREGATE)\s*\(', str(formulas[tr][tc] or ""), re.I)):
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
            # ROW()-3・COLUMN()-1 は見出しの行数・列数を引く連番の決まり文句（2026-11-05 シュウさんの計算表 AM4:AM13
            # =IF(AD4<>0,ROW()-3,"") の 10 件）。数を直書きしても、変える場面が無い
            scan = re.sub(r'\b(?:ROW|COLUMN)\s*\([^()]*\)\s*[-+]\s*\d+', 'ROW()', local_ref_text(f), flags=re.I)
            # 先頭のオペランドの直書き（=856-N17・=2231+C8・=515/$E$6・IF(…,99-$N$12,…)）も同じに見る
            # （2026-11-05 注入テスト: 後ろの直書き 40/40 に対し、先頭は 0/5 で見落としていた）
            _lead = [(m_.group(2), m_.group(1), True, m_.start(1)) for m_ in _RE_HARDCODED_LEAD.finditer(scan)]
            _cands = [(m_.group(1), m_.group(2), False, m_.start(2)) for m_ in _RE_HARDCODED_CONST.finditer(scan)] + _lead
            for op, val_str, _is_lead, _pos in _cands:
                try:
                    val_num = float(val_str)
                except ValueError:
                    continue
                # 0 や 1 以外の実質的な定数計算を言う。重さは「気づき」: 誤りではなく「変わり得る数は条件のセルに出すと
                # 変えやすい」という勧め（2026-11-05 シュウさんの判断の依頼）。点数は同じ形の文・同じ列ごとに 1 回（audit_score）
                if (val_num not in (0.0, 1.0) and not (op in '*/' and val_num in _UNIT_CONSTANTS)
                        and not (op == '/' and val_num in _UNIT_DIV_ONLY)
                        and not (op in '*/' and _is_calendar_month_mix(scan, val_num))
                        # 指数の係数 ^(-2*t)（符号の - は引き算ではない）
                        and not (op == '-' and scan[max(0, _pos - 3):_pos].replace(' ', '').endswith('^(-'))
                        and not (op in '+-' and not _is_lead and float(val_num).is_integer() and val_num <= 31
                                 and _after_month_end_func(scan, _pos))
                        # 検索・位置の関数の中で参照に足し引きする小さな整数（列番号・行番号のオフセット）
                        and not (op in '+-' and not _is_lead and float(val_num).is_integer() and val_num <= 12
                                 and _enclosing_func(scan, _pos) in _RE_OFFSET_FUNCS)):
                    how = {'+': '足す', '-': '引く', '*': '掛ける', '/': '割る'}.get(op, op)
                    if _is_lead:
                        how = {'+': '足す', '-': '引かれる数', '*': '掛ける', '/': '割られる数'}.get(op, op)
                    issues.append({
                        "cell": addr,
                        "severity": "info",
                        "type": "hardcoded_constant",
                        # 数は数だけで言う（ROUNDDOWN(H4,0)-10 の引き算の - を数に含めて「-10」と言っていた）
                        "msg": f"数式内に固定値 {val_str}（{how}）が直書きされています（変わり得る数なら条件のセルを参照）",
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
                        # 合計が範囲の上にある形（G12=SUM(G13:G15)・日本語の統計表は合計を先頭の行に置く）は、範囲の下端の
                        # すぐ下の数を見る。式が範囲の真上、または式の行に「計・合計・総数」の見出しがあるときだけ
                        # （2026-11-05 注入テスト: 合計が上の形は 0 件拾えていなかった）
                        # 隣の列の同じ行の集計（=SUM(AB22:AB25)）が、同じ先頭から漏れと言う相手まで足している＝並べた集計の 1 つだけ短い
                        new_gate = False
                        if not between and c == sum_col and omitted_below != r:
                            for d_ in (-2, -1, 1, 2):
                                k_ = sum_col + d_
                                if 0 <= k_ < num_cols:
                                    m_ = _RE_SUM_FUNC.search(str(formulas[r][k_] or ""))
                                    if m_ and m_.group(1).upper() == fn_name:
                                        q1_, qc1_ = _a1_to_rowcol(m_.group(2))
                                        q2_, qc2_ = _a1_to_rowcol(m_.group(3))
                                        if (qc1_ == qc2_ and q1_ - start_row == top_r and q2_ - start_row == omitted_below):
                                            between = True
                                            new_gate = True
                                            break
                        if (not between and c == sum_col and r < top_r
                                and (top_r == r + 1 or any(_is_total_label(values[r][k_])
                                                           for k_ in range(num_cols) if k_ != sum_col))):
                            between = True
                            new_gate = True
                        # 範囲と式の間に別の合計の式（=SUM(AL5:AL27) の「計」）があれば、式は表の下の別の部分集計
                        # （「都心３区」=SUM(AL5:AL7)）で、漏れではない（2026-11-05 日本語の公開ブック 60 本）
                        if between and c == sum_col and any(
                                _is_vertical_agg(str(formulas[k][sum_col] or ""))
                                for k in range(omitted_below, r)):
                            between = False
                        if (0 <= omitted_below < num_rows and omitted_below != r and not running and between
                                and not _is_subtotal(omitted_below, sum_col)):
                            val_below = values[omitted_below][sum_col]
                            target_cell = _rowcol_to_a1(start_row + omitted_below, start_col + sum_col)
                            # 合計が上にある形・隣の列の並びから見つけた漏れは、漏れと言う相手が合計より桁違いに大きければ
                            # （=SUM(G13:G15) の下の G16 が 3 万）それ自身が別の合計。言わない
                            big_total = (new_gate and isinstance(val_below, (int, float)) and isinstance(v, (int, float))
                                         and abs(val_below) > 3 * abs(v) + 1e-9)
                            if (isinstance(val_below, (int, float)) and not isinstance(val_below, bool) and not big_total
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
                        if (0 <= omitted_right < num_cols and omitted_right != c and r == sum_row and omitted_right < c
                                and not _sibling_shares_bound(r, c, "right", left_c, right_c)):
                            val_right = values[sum_row][omitted_right]
                            target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_right)
                            if (isinstance(val_right, (int, float)) and not isinstance(val_right, bool)
                                    and not _skip_omitted(val_right, target_cell, f, addr, 'h')):
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
                                and not _copied(r, c, False)
                                and not _sibling_shares_bound(r, c, "left", left_c, right_c)):
                            val_left = values[sum_row][omitted_left]
                            target_cell = _rowcol_to_a1(start_row + sum_row, start_col + omitted_left)
                            if (isinstance(val_left, (int, float)) and not isinstance(val_left, bool)
                                    and not (above_left.add(target_cell) or _skip_omitted(val_left, target_cell, f, addr, 'h')) and not _year_or_date(val_left)):
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
                # 参照 1 つを囲んだだけの括弧 =(L44)/8977 と =L44/8977 は同じ式（2026-11-06 日本語の台帳）
                f_r1c1 = _RE_PAREN_SINGLE_REF.sub(r"\1", f_r1c1)
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
                # 違いが日数（/28・/30・/31・/365）だけの式は月ごと・年ごとの作り（2026-11-05 初見の 60 本: 月の平均を
                # 日数で割る表の 2 月の =T290/28 を「コピペミス疑い」と言っていた）
                if _only_day_counts_differ(pat, before):
                    continue
                r, c = (t, k) if along == "col" else (k, t)
                # 横の比べで、同じ列のセルだけを足す・写す式（=C17+C34・=D224）は、列ごとの集計・転記の行で、左右と形が違って当たり前
                # （自分の行を横に足す範囲 =SUM(RC[-10]:RC) を含む式は別の話で、小計の行の差としては外さない）
                # （同じ行の別の列を写す式どうしで列だけがずれた =B2 と =C2 は、横のコピーミスの本命なので外さない）
                same_row_copy = (_is_pure_ref(pat) and _is_pure_ref(before) and pat.split('C')[0] == before.split('C')[0])
                if along == "row" and not same_row_copy and (_is_column_only(pat) or (_is_column_only(before) and ":" not in pat)):
                    continue
                # 縦の比べで、同じ行に同じ列だけを足す式（=SUM(D$82,D$87,…)・=AI39+AI42+AI45）がある行は小計の行
                # （小計の行の隣の式が明細と形が違うのは当たり前。2026-11-05 dock の台帳: 小計ブロック頭の差額の式）
                if along == "col" and any(
                        cc != c and cc < len(formulas_r1c1[r]) and _is_column_only(str(formulas_r1c1[r][cc] or ""))
                        for cc in range(num_cols)):
                    continue
                # 他のシート・他のブックの決まった番地を直に指す式（=+Lonestar!F43・='[外.xls]S'!R51C37）は、行ごとに引き先を選んだ式
                # （相対の参照が 1 つも無い）。引き先の並びに決まりが無く、隣と形が違って当たり前
                if "!" in pat and not re.search(r"[RC]\[", pat):
                    continue
                # 同じシートの式が並ぶ列に、他のシート・ブックを指す式が 1 つ混じるのは、その行だけ別の所から引く作り
                if "!" in pat and "!" not in before:
                    continue
                # 他のブックを指す式（='\fs\…\[決算額調.xls]S'!R4C3・='C:\…\[一覧.xlsx]S'!…）は、引き先のブックの形に従う式で、
                # 隣と形が違って当たり前（2026-11-06 日本語の台帳: 外部ブック参照式の誤爆 11 件）
                if _only_external_refs(pat) or _only_external_refs(before):
                    continue
                # 1 つの番地を写すだけの式（=+K48）が並び、写す元の行の差がまちまち＝行ごとに写す元を選んでいる列は、形が違って当たり前
                # （写す元の行の差がそろった中の列のずれ =+F6 が =+G6 の列に混ざるのは、拾う）
                if _is_pure_ref(pat) and _is_pure_ref(before) and pat.split('C')[0] != before.split('C')[0]:
                    continue
                # SUM に飛び飛びの範囲（=SUM(L19:M19,N19:P19)）を指定した式は、足す所を選んだ式
                if re.match(r"^=\+?SUM\(", pat, re.I) and "," in pat:
                    continue
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
    # 別の式が「=+Q35」とそのセル 1 つだけを写している数は、次の期へ繰り越す初期値・入力欄で、紛れた直書きではない
    # （2026-11-05 dock の台帳: 月のブロックの頭に置いた初期値 100・52・122 を式の列の直書きと言う誤爆）。
    # =D52-D40 のように計算に使うだけの式は数えない（その D52 は明細の列に紛れた直書きの本物）
    referenced = set()
    for _row in formulas:
        for _fx in _row:
            if isinstance(_fx, str) and _fx.startswith("="):
                _m = re.fullmatch(r"=\s*\+?\s*\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})\s*", _fx)
                if _m:
                    referenced.add(_m.group(1).upper() + _m.group(2))
    issues.extend(check_values_in_formula_columns(formulas, formulas_r1c1, values, start_row, start_col,
                                                  referenced=referenced))
    return issues


def check_values_in_formula_columns(formulas, formulas_r1c1, values, start_row, start_col, min_rows=3, referenced=None):
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
                # 別のシート・ブックを行ごとに 1 行ずつ指す式（='[x.xls]S'!R254C7・次の行は R255C7）は、絶対の行番号を
                # 自分の行からの差にそろえて、同じ形にする（範囲 R5C1:R40C3 の参照表は動かさない。2026-11-05 注入テスト: 外部ブック参照が並ぶ列の途中の直書きを見落とした）
                if "!" in fr:
                    fr = re.sub(r"(!)R(\d+)(C(?:\d+|\[-?\d+\])?)(?![\d:\[])", lambda m_, ar=start_row + r + 1: "%sR[%d]%s" % (m_.group(1), int(m_.group(2)) - ar, m_.group(3)), fr)
                pats.setdefault(fr, []).append(r)
        # 番地を含まない式（=#N/A の置き場・=2866+1200 の電卓）は「式の列」の形にしない（2026-11-05 dock の実測）。
        # テーブルの計算列の構造化参照（[@unit]）と ROW()・COLUMN() の連番の式は番地を含む式と同じに数える
        pats = {k: v for k, v in pats.items()
                if _RE_R1C1_REF.search(_RE_STRING_LIT.sub('""', k)) or re.search(r"\[@|(?:ROW|COLUMN)\s*\(\s*\)", k, re.I)}
        if not pats:
            continue
        top, rows = max(pats.items(), key=lambda kv: len(kv[1]))
        if len(rows) < min_rows:
            continue
        first_a1 = str(formulas[rows[0]][c])
        rows_set = set(rows)
        # 式と式の間の半分以上に数が入っている列は「値の行・式の行」を繰り返す作り（2026-11-05 初見の 60 本:
        # 値・日数で割る式・0 の 3 行を繰り返す表の 144 件）。ぽつんと 1 つ紛れた数ではない
        gaps = [(a, b) for a, b in zip(rows, rows[1:]) if b - a > 1]
        # 式が 8 行以上並ぶ列では、間に入っているのは手入力の数（式の答えは数えない: 小計の式・別の形の式の行が間にあるだけの列を
        # 『値の行・式の行の繰り返し』と見て黙っていた＝2026-11-05 注入テスト）。短い列は今までどおり
        filled = sum(1 for a, b in gaps if any(isinstance(values[k][c], (int, float)) and not isinstance(values[k][c], bool)
                                               and (len(rows) < 8 or not str(formulas[k][c] or "").startswith("="))
                                               for k in range(a + 1, b) if k < len(values) and c < len(values[k])))
        if gaps and filled * 2 >= len(gaps) and filled >= 2:
            continue
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
                a1 = _rowcol_to_a1(start_row + r, start_col + c)
                # 0 の手入力は「無し」の印（未稼働の行）で、直書きの数字の誤りではない。別の式が指している数は入力欄
                # 式が 8 行以上並ぶ列の途中で、真上と真下がどちらも式の行（切れ目の無い式の列の 1 つ）は、0 でも、別の式が
                # 指していても、式を値で潰した貼り付けと見る。0 の手入力は『無し』の印、別の式が指す数は入力欄、は
                # 短い・途切れた式の列の話（2026-11-05 注入テスト: 値が 0 になる貼り付け 0/6・行の合計の式が指す列は黙っていた）
                solid = len(rows) >= 8 and (r - 1) in rows_set and (r + 1) in rows_set
                if not solid and (v == 0 or (referenced and a1 in referenced)):
                    continue
                out.append({
                    "cell": a1,
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


def _is_pure_ref(f_r1c1):
    """=+K48 のように、別の行の番地 1 つだけを写す式か（R1C1 の字面。同じ行の =+F6 は横に伸ばす式なので除く）。"""
    return bool(re.fullmatch(r"=\+?R(\[-?[1-9]\d*\]|\d+)C(\[-?\d+\]|\d+)?", str(f_r1c1 or "").replace(" ", "")))


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


_DAY_COUNTS = {'28', '29', '30', '31', '365', '366'}
_RE_NUM_LIT = re.compile(r'(?<![A-Za-z0-9_\[.])\d+(?:\.\d+)?(?![\d\]])')


def _only_day_counts_differ(a, b):
    """2 つの R1C1 の式が、数の直書きの日数（28〜31・365・366）だけ違うか（純 Python）。"""
    na, nb = _RE_NUM_LIT.findall(a), _RE_NUM_LIT.findall(b)
    if _RE_NUM_LIT.sub('#', a) != _RE_NUM_LIT.sub('#', b) or len(na) != len(nb):
        return False
    diff = [(x, y) for x, y in zip(na, nb) if x != y]
    return bool(diff) and all(x in _DAY_COUNTS and y in _DAY_COUNTS for x, y in diff)


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


_RE_YEN_TEXT = r"[0-9０-９][0-9０-９,.]*(?:百万|千|万|億)?円"
_RE_FREE_TEXT_HEAD = re.compile(r"内容|説明|概要|詳細|摘要")

_RE_EXT_REF_R1C1 = re.compile(r"'[^']*\[[^\]]*\.xl[^\]]*\][^']*'!(?:R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?)(?::R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?)?", re.I)


def _only_external_refs(p: str) -> bool:
    """他のブックを指す参照があり、その参照を除くと相対の参照（R[1]C[-1]）が残らない式か。
    =VLOOKUP(A20,'[外.xls]S'!…) のように自分のシートの相対参照が残る式は、行のずれを比べる本命なので外さない。"""
    if not _RE_EXTERNAL_BOOK.search(p or ""):
        return False
    return not re.search(r"R\[|C\[", _RE_EXT_REF_R1C1.sub("0", p))


_RE_PAREN_SINGLE_REF = re.compile(r"\(\s*(R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?)\s*\)")
_RE_EXTERNAL_BOOK = re.compile(r"\[[^\]\[]+\.(?:xls[xmb]?|xlam?|csv)\]", re.I)


def _ws_is_layout(values, r, c, v, trimmed) -> bool:
    """前後の空白が体裁（字割り・様式の項目名・年度表示・単位・表題・自由記述）か（純 Python・2026-11-06 日本語の台帳 129 件）。
    照合のキーの空白（同じ列の少数派・右を 3 つ以上の空白で埋めた固定長）は体裁にしない。
    1) 1 字だけの文字（'㎡ '・'… '）  2) 語の途中に改行がある折り返しの文・斜線見出し
    3) （　）で囲んだ注記  4) 同じ形（先頭・末尾・全角空白の有無）の前後の空白を持つ文字が同じ列に 2 つ以上でかつ 4 分の 1 以上（または列の 3 分の 1 以上が前後に空白を持つ）＝列全体の体裁（和暦 ' H24'・様式項目名）
       （右を 2 つ以上の空白で埋めた固定長は除く）
    5) 表の上 3 行で 3 つ以内のセルの行（表題・入力欄）  6) 見出しが内容・説明・概要の列（自由記述）
    ただし英数字だけで先頭に空白の無い文字（'Other '）は照合を邪魔する本物＝体裁にしない（4・5 を外す）"""
    if len(trimmed) <= 1:
        return True
    if "\n" in trimmed or "\r" in trimmed:
        return True
    if v != v.lstrip() and re.match(r"^[（(].*[）)]$", trimmed):
        return True
    for k in range(min(r, 15)):
        h = values[k][c] if c < len(values[k]) else None
        if isinstance(h, str) and len(h.strip()) <= 12 and _RE_FREE_TEXT_HEAD.search(h):
            return True
    # 英数字だけの文字の末尾の空白（'Other '・'Asset/Accrual Margin '）は、列に多くても照合を邪魔する本物
    # （体裁として列全体に付くのは日本語の様式の字割り・字下げ・和暦年度の先頭空白。2026-11-06 英語の台帳 3 件）
    if trimmed.isascii() and v.lstrip() == v:
        return False
    # 和暦の年度表示（' H24'・' R03'）の先頭空白は桁をそろえる字下げ
    if v != v.lstrip() and re.fullmatch(r"[HRSTMhrstm]\d{1,2}", trimmed):
        return True
    if r <= 2 and len(values) >= 8 and sum(1 for x in values[r] if x not in (None, "")) <= 3:
        return True
    padded = bool(re.search(r" {2,}$", v))
    col = [row[c] for row in values if c < len(row) and isinstance(row[c], str) and row[c].strip()]
    # 同じ列の字寄せ・字下げ・金額表記の体裁（2026-11-06 日本語のホールドアウト 7 件）
    # (a) 金額表記の文字（'20百万円　'）の末尾の全角空白: 同じ列に金額表記が他にもある（空白なしの '270百万円' が混在）
    if re.fullmatch(_RE_YEN_TEXT, trimmed) and v == trimmed + "　" * (len(v) - len(trimmed)) and v != trimmed:
        if sum(1 for x in col if x is not v and re.fullmatch(_RE_YEN_TEXT, x.strip())) >= 1:
            return True
    # (b) 見出し ' 過去５年間平均' の先頭 1 字の空白: 同じ列の和暦年度（' H25'）の字寄せと同じ
    if len(v) == len(trimmed) + 1 and v[0] in " 　" and sum(1 for x in col if re.fullmatch(r"\s[HRSTMhrstm]\d{1,2}", x)) >= 2:
        return True
    # (c) 様式項目名の字下げ（'　　内　資産割  '）: 先頭が全角空白で、列の 3 分の 1 以上が全角空白で始まる
    if v.startswith("　") and not trimmed.isascii():
        lead = sum(1 for x in col if x.startswith("　"))
        if lead >= 3 and lead * 3 >= len(col):
            return True
    # 同じ形の前後の空白（先頭が ' '・末尾が '　' など）を持つ文字が列の 4 分の 1 以上＝列全体の体裁
    def _sig(x):
        return (x != x.lstrip(), x != x.rstrip(), "　" in x)
    sig = _sig(v)
    same = sum(1 for x in col if _sig(x) == sig)
    ws = sum(1 for x in col if x != x.strip())
    if not padded and ((same >= 2 and same * 4 >= len(col)) or (ws >= 5 and ws * 3 >= len(col))):
        return True
    return False



def check_data_cleaner(
    values: List[List[Any]],
    start_row: int,
    start_col: int,
    formulas: Optional[List[List[Any]]] = None,
    peer_sheets: Optional[List[Tuple[int, int, List[List[Any]]]]] = None,
) -> List[Dict[str, Any]]:
    """データクレンジング診断（隠れ空白・文字列数字・外れ値）を行う純粋ロジック。
    peer_sheets: 同じブックのほかのシートの値 [(start_row, start_col, values), ...]（0 始まりの左上）。
    同じ番地に同じ値が別のシートにもある数は、同じ型のシートで決まって置く値で、外れ値と言わない。"""
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

    # 全角の数字の見出し（'３'・'４'）がある列は、半角の '10'・'11' も同じ見出しの番号
    col_zen = [any(c < len(values[r]) and isinstance(values[r][c], str) and re.fullmatch(r"[０-９]+", values[r][c].strip())
                   for r in range(num_rows)) for c in range(num_cols)]

    # 見出しの行: 文字が 2 つ以上・数が無く、すぐ下の行に数が 2 つ以上ある行（2026-11-05 シュウさんの計算表の
    # 3 行目 ' 区　分'。その列の下は文字の明細で、列ごとの見方では見出しと分からなかった）
    def _n_num(row):
        return sum(1 for x in row if isinstance(x, (int, float)) and not isinstance(x, bool))
    header_rows = {r for r in range(num_rows - 1)
                   if _n_num(values[r]) == 0 and sum(1 for x in values[r] if isinstance(x, str) and x.strip()) >= 2
                   and _n_num(values[r + 1]) >= 2}

    def _is_header_cell(r, c):
        """見出しのセルか（見出しの行の中・または列の数より上: 下に数が 2 つ以上・上に数が無い）。"""
        if r in header_rows:
            return True
        rows = col_num_rows[c]
        return len(rows) >= 2 and rows[0] > r

    def _cell(r, c):
        return values[r][c] if 0 <= r < num_rows and 0 <= c < len(values[r]) else None

    def _is_tn(x):
        return isinstance(x, str) and bool(_RE_TEXT_NUMBER.match(x.strip()))

    def _label_num(x):
        """並びの一員になれるセル: 文字の数字、または整数の数。"""
        if _is_tn(x):
            return float(_tn_value(x))
        return float(x) if _isnum(x) and float(x).is_integer() else None

    def _is_label_sequence(r, c):
        """文字の数字が、同じ行か列の等差の並び（年 1998〜2001・時刻 0500〜1500・日 13〜23）の一員か。
        並びは 5 つ以上、または 3 つ以上で文字の数字が 2 つ以上（数の列に 1 個だけ紛れた文字の数字は並びと見ない）。
        （2026-11-05 dock の台帳: 見出しの年・時刻・日の番号を文字の数字と言う誤爆）"""
        for dr, dc in ((1, 0), (0, 1)):
            cells = [(r, c)]
            for sgn in (-1, 1):
                k = 1
                while k <= 60:
                    pos = (r + sgn * k * dr, c + sgn * k * dc)
                    if _label_num(_cell(*pos)) is None:
                        # 1 つ飛ばし（番号の間に空の行を挟んだ年の列 25・26・27）は並びが続く
                        nxt = (r + sgn * (k + 1) * dr, c + sgn * (k + 1) * dc)
                        if _cell(*pos) in (None, "") and _label_num(_cell(*nxt)) is not None:
                            k += 1
                            continue
                        break
                    cells.append(pos)
                    k += 1
            cells.sort()
            me = cells.index((r, c))
            nums = [_label_num(_cell(a, b)) for a, b in cells]
            for i in range(0, me + 1):
                for j in range(max(me, i + 2), len(cells)):
                    step = nums[i + 1] - nums[i]
                    if step == 0 or any(nums[k + 1] - nums[k] != step for k in range(i, j)):
                        break
                    n_text = sum(1 for a, b in cells[i:j + 1] if _is_tn(_cell(a, b)))
                    # 並びの中に本物の数が混じる（連番の途中の 1 つだけが文字の数字）なら並びの見出しではなく入力の誤り。
                    # 文字の数字が 2 つ以上あり、本物の数より少なくない並びだけを見出しと見る
                    # （2026-11-05 注入テスト: 通し番号・年の列の途中の文字の数字が 9 件、等差の除外で黙っていた）
                    if n_text >= 2 and n_text >= (j - i + 1) - n_text:
                        return True
        return False

    def _is_form_field(r, c):
        """行の頭が文字の見出しで、この数の左右 2 マスに本物の数が無く、上下も文字の数字でない＝「契約番号: 26372」のような記入欄。
        （文字の数字が縦に並ぶ列は記入欄ではない＝CSV の金額の列を黙らせない）"""
        row = values[r]
        first = next((x for x in row if x not in (None, "")), None)
        return (isinstance(first, str) and bool(first.strip()) and not _is_tn(first)
                and not any(_isnum(_cell(r, c + d)) for d in (-2, -1, 1, 2))
                and not _is_tn(_cell(r - 1, c)) and not _is_tn(_cell(r + 1, c))
                # 上下が本物の数＝数の列の途中の 1 つ（数の列に混じった文字の数字）で、ぽつんと置いた記入欄ではない
                and not (_isnum(_cell(r - 1, c)) and _isnum(_cell(r + 1, c))))

    def _is_label_row(r):
        """行に文字の見出し（数でない文字）が 2 つ以上あり、本物の数が 1 つも無い＝名前・区分・番号を並べた一覧の行
        （'Atlanta'・'6484'・'PRIM'・'SCOTT'）。CSV の金額の列のように文字の数字だけが並ぶ行とは違う。"""
        row = values[r]
        return (not any(_isnum(x) for x in row)
                and sum(1 for x in row if isinstance(x, str) and x.strip() and not _is_tn(x)) >= 2)

    def _repeated_in_column(c, trimmed, r=None):
        """同じ文字の数字が列に 3 つ以上、またはすぐ上下に同じもの（'1' の区分・'6484' の番号が行ごとに繰り返す）＝区分・番号の値で金額ではない。"""
        if r is not None and any(isinstance(_cell(r + d, c), str) and _cell(r + d, c).strip() == trimmed for d in (-1, 1)):
            return True      # すぐ上か下に同じ文字の数字（PRIM・SEC の行ごとに同じ番号が並ぶ）
        return sum(1 for k in range(num_rows) if c < len(values[k]) and _is_tn(values[k][c])
                   and values[k][c].strip() == trimmed) >= 3

    def _repeated_in_row(r, trimmed):
        """同じ文字の数字が同じ行に 3 つ以上＝ブロックごとに繰り返す日の番号・年の見出し（A25・AE25・AI25 の '18'）で、金額ではない。"""
        return sum(1 for x in values[r] if _is_tn(x) and x.strip() == trimmed) >= 3

    def _in_id_column(c, trimmed):
        """列の本物の数が、桁数のそろった整数（5 桁以上・5 個以上）＝番号（部門コード・契約番号）の列で、この文字も同じ桁数。
        金額の列は桁がそろわない。"""
        nums = [values[k][c] for k in range(num_rows) if c < len(values[k]) and _isnum(values[k][c])]
        if len(nums) < 5 or not all(float(x).is_integer() and x > 0 for x in nums):
            return False
        lens = {len(str(int(x))) for x in nums}
        return len(lens) == 1 and next(iter(lens)) >= 5 and len(trimmed) == next(iter(lens))

    def _is_year_heading(r, c, trimmed):
        """年らしい 4 桁（1900〜2100）が見出しの位置か: 行に本物の数が無い（見出しだけの行）、または上下か左右が数でない文字
        （'FY' の下の 2001・'Total' の上の 1998）。数の列に紛れた年は言う。"""
        if not (len(trimmed) == 4 and trimmed.isdigit() and 1900 <= int(trimmed) <= 2100):
            return False
        if not any(_isnum(x) for x in values[r]):
            return True
        for rr, cc in ((r - 1, c), (r + 1, c)):
            x = _cell(rr, cc)
            if isinstance(x, str) and x.strip() and not _is_tn(x):
                return True
        return False

    _AMOUNT_HEAD = re.compile(r"金額|売上|価格|単価|料|円|額|合計|費|収入|支出|数量|amount|sales|price|cost|total|revenue|qty", re.I)
    _ID_HEAD = re.compile(r"番号|コード|ID|No|郵便|電話|科目|型番|code|zip|tel|phone", re.I)

    def _looks_like_id_column(c):
        """文字の数字ばかりの列が番号（勘定科目・コード）の列に見えるか。見出しが金額・売上などなら金額の列（CSV 取り込みで
        文字になった数）で、番号の列ではない。見出しが番号・コードなら番号の列。見出しで決まらなければ、桁数がそろっているか
        （8 割以上が同じ桁数で 3 桁以上）で見る。金額は 99・780・15200 と桁がまちまち。"""
        head = next((values[k][c] for k in range(num_rows) if c < len(values[k]) and isinstance(values[k][c], str)
                     and values[k][c].strip() and not _RE_TEXT_NUMBER.match(values[k][c].strip())), "")
        if _AMOUNT_HEAD.search(str(head)):
            return False
        if _ID_HEAD.search(str(head)):
            return True
        lens = [len(values[k][c].strip().lstrip("-")) for k in range(num_rows)
                if c < len(values[k]) and _is_tn(values[k][c])]
        if not lens:
            return True
        top = max(lens.count(x) for x in set(lens))
        return top * 5 >= len(lens) * 4 and max(lens, key=lens.count) >= 3

    def _stands_alone(r, c):
        """周り 8 マスに本物の数も文字の数字も無い（集計に入れる相手がいない・文字の数字が並ぶ列の一員ではない）。ぽつんと置いた番号・区分の文字の数字は SUM の漏れにならない。"""
        for rr in range(r - 1, r + 2):
            for cc in range(c - 1, c + 2):
                if (rr, cc) != (r, c) and (_isnum(_cell(rr, cc)) or _is_tn(_cell(rr, cc))):
                    return False
        return True

    for r in range(num_rows):
        for c in range(num_cols):
            v = values[r][c]
            addr = _rowcol_to_a1(start_row + r, start_col + c)

            if isinstance(v, str):
                # A. 隠れ空白・前後の空白検知
                trimmed = v.strip()
                # 咎めない空白（2026-11-05 dock の実測: 1,447 件の大半がこれ）: 空白だけのセル・字下げ（先頭だけの空白）・
                # 数の列の上の見出し（'UK Gas FIN OPT '）。照合に使うのは本文の値で、見出しの末尾の空白は表の外に響かない
                lead = v[:len(v) - len(v.lstrip(" 　"))]
                # 字下げ: 半角 2 つ以上、または全角 1 つ以上（日本語の表は「　池上五丁目」と全角 1 字で下げる・
                # 2026-11-05 日本語の公開ブック 60 本の 8,710 件）
                indent_only = (v.rstrip() == v and not ("\n" in v or "\r" in v)
                               and (len(lead) >= 2 or "　" in lead))
                # 照合に響くのは「同じ語が空白なしでもある」か「文字の値が並ぶ列（3 つ以上）」の中だけ。
                # ぽつんと置いたラベル（'Total '・'DATE: '・横に並んだ見出し）の末尾の空白は誰も照合しない
                listed = clean_texts.get(trimmed, 0) > 0 or col_texts[c] >= 3
                # 注記・説明の文（「なお、…」・小文字の英単語が 3 つ以上続く文）は照合に使わない（2026-11-05 日本語の公開ブック）。
                # 字数では決めない（右を空白で埋めた長い会社名の列 'NOBLE GAS MARKETING   ' は照合に響く本物）
                if re.search(r"[、。]", trimmed) or re.search(r"\b[a-z]{2,}\s+[a-z]{2,}\s+[a-z]{2,}\b", trimmed):
                    listed = False
                quiet = (not trimmed) or indent_only or _is_header_cell(r, c) or not listed or _is_total_label(v)
                if not quiet and v != trimmed:
                    quiet = _ws_is_layout(values, r, c, v, trimmed)
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
                code_col = col_textnum[c] > len(col_num_rows[c]) and _looks_like_id_column(c)
                # 全角の数字（'９'）は見せるための見出しの番号で、足し算に使わない（2026-11-05 日本語の公開ブック 60 本）
                zenkaku = bool(re.search(r"[０-９]", trimmed)) or col_zen[c]
                # ただし本物の数の列の途中（真上と真下が本物の数で、列に本物の数が 5 つ以上）の全角の数字の文字列は、見出しの番号ではなく
                # 足し算から漏れる入力の誤り（2026-11-05 注入テスト: 全角数字の文字列は 0/47 で全部黙っていた）
                if zenkaku and len(col_num_rows[c]) >= 5 and _isnum(_cell(r - 1, c)) and _isnum(_cell(r + 1, c)):
                    zenkaku = False
                if (_RE_TEXT_NUMBER.match(trimmed) and len(trimmed) <= (25 if "." in trimmed else 15) and not zenkaku
                        and not re.fullmatch(r"0\d+", trimmed) and not code_col and not _is_header_cell(r, c)
                        and not _is_label_sequence(r, c) and not _is_year_heading(r, c, trimmed) and not _stands_alone(r, c)
                        and not _is_form_field(r, c) and not _in_id_column(c, trimmed)
                        and not _is_label_row(r) and not _repeated_in_column(c, trimmed, r) and not _repeated_in_row(r, trimmed)
                        and not (v != v.lstrip() and v.rstrip() == v)):    # 字下げした数（'    14'）は見せるための番号の見出し
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
    date_rows = {r for r in range(num_rows)
                 if sum(1 for x in values[r] if isinstance(x, (int, float)) and not isinstance(x, bool)
                        and float(x).is_integer() and 30000 <= x <= 60000) >= 3}
    def _neg_in_columns(c):
        """この列の左右 3 列にあるマイナスの数（増減・差額の表では、どの列にもマイナスが並ぶ）。"""
        return sum(1 for k in range(max(0, c - 3), min(num_cols, c + 4)) if k != c
                   for rr in range(num_rows) if k < len(values[rr]) and _isnum(values[rr][k]) and values[rr][k] < 0)

    def _row_context_quiet(r, c, v):
        """同じ行の中でも説明のつく値: 同じ行にほかのマイナスが 2 つ以上ある（増減の行）、または行の隣が等差の連番（1・2・3・4 の番号の見出し）。"""
        row = values[r]
        if v < 0 and sum(1 for k, x in enumerate(row) if k != c and _isnum(x) and x < 0) >= 2:
            return True
        if 0 < c < len(row) - 1 and _isnum(row[c - 1]) and _isnum(row[c + 1]):
            return row[c] - row[c - 1] == row[c + 1] - row[c] != 0
        return False

    def _period_days_in_row(row, days):
        """行の文字に '11月12日～12月25日'・'4/5-10/16' の期間があり、その日数（終わり-始め、または両端を入れて +1）が days か。"""
        import datetime as _dt
        for x in row:
            if not isinstance(x, str):
                continue
            for m_ in _RE_PERIOD_RANGE.finditer(x):
                a_m, a_d, b_m, b_d = (int(g) for g in (m_.group(1) or m_.group(5), m_.group(2) or m_.group(6),
                                                       m_.group(3) or m_.group(7), m_.group(4) or m_.group(8)))
                try:
                    d1 = _dt.date(2001, a_m, a_d)
                    d2 = _dt.date(2001 if (b_m, b_d) >= (a_m, a_d) else 2002, b_m, b_d)
                except ValueError:
                    continue
                if days in ((d2 - d1).days, (d2 - d1).days + 1):
                    return True
        return False

    def _outlier_context_quiet(r, c, v, num_entries):
        """列の中で離れていても、行の並び・ラベル・列の長さ・列の桁の揺れから外れ値と言えない型を黙らせる
        （2026-11-05 outlier の新規 93 件を実物で全件判定した台帳 ledger_outlier.tsv の誤爆 87 件の型 T1〜T7）。"""
        row = values[r]
        av = abs(v)
        # T9 ほかのシートの同じ番地にも同じ値がある（月ごと・店ごとに同じ型で作ったシートの固定の値）
        if peer_sheets:
            ar_, ac_ = start_row + r, start_col + c
            for pr_, pc_, pv_ in peer_sheets:
                i_, j_ = ar_ - pr_, ac_ - pc_
                if 0 <= i_ < len(pv_) and 0 <= j_ < len(pv_[i_]) and _isnum(pv_[i_][j_]) and pv_[i_][j_] == v:
                    return True
        # T8 同じ行に日付の範囲の文字（'11月12日～12月25日'・'4/5～10/16'）があり、値がその期間の日数（両端を入れる数え方も）
        # ＝日数の列で、ほかの行より大きいのは期間が長いだけ（日本語の公開ブックの台帳: 期間日数 6 件）
        if float(v).is_integer() and v > 0 and _period_days_in_row(row, int(v)):
            return True
        # T2 同じ行に同規模（0.1〜10 倍）の数がある＝横に見れば外れていない（同値の繰り返し・同じ行の合算など）
        # （式のセルは数えない: 単価の打ち間違いが 数量×単価 の金額の列にそのまま写るため）
        def _is_formula(k):
            return bool(formulas) and r < len(formulas) and k < len(formulas[r]) and str(formulas[r][k] or "").startswith("=")
        peers = [x for k, x in enumerate(row) if _isnum(x) and k != c and x != 0 and not _is_formula(k)
                 and av / 10 <= abs(x) <= av * 10]
        # 同規模の数が 2 つ以上、または 1.5 倍以内にもう 1 つ（同規模が 1 つだけで 2 倍以上離れていると、金額の列に写った打ち間違いと区別できない）
        # ただし列のほかの値の最大の 100 倍以上で、列が 8 件以上ある（桁違いが明白）なら、横に同規模の数が 2 つあるだけでは黙らない
        # （1.5 倍以内に同値・近い値がある繰り返しは今までどおり黙る）。件数と金額など単位の違う列が横に並ぶ表で、
        # 件数の列の 2 を 2000 に打ち間違えても、金額の列の 501 が相手に見えて黙っていた（2026-11-05 注入テスト）
        if av and any(av / 1.5 <= abs(x) <= av * 1.5 for x in peers):
            return True
        if av and len(peers) >= 2:
            col_others = sorted(abs(x) for rr_, x in num_entries if rr_ != r and x != 0)
            if not (len(col_others) >= 8 and av >= 100 * col_others[-1]):
                return True
        # 同じ列の連続した数の並び（r を含む）の長さ。T3 5 個未満の短い塊は設定・控え・記入欄で、比べる相手がいない
        a = r
        while a - 1 >= 0 and c < len(values[a - 1]) and _isnum(values[a - 1][c]):
            a -= 1
        b = r
        while b + 1 < num_rows and c < len(values[b + 1]) and _isnum(values[b + 1][c]):
            b += 1
        run = b - a + 1
        if run < 5:
            return True
        others = [abs(x) for rr, x in num_entries if rr != r and x != 0]
        # T4 列の非ゼロが 3 件未満（比べる相手がいない）
        if len(others) + 1 < 3:
            return True
        # T3 行に文字のラベルがあり数が 3 つまで、列の塊も短い＝「Peak 700 2200」「Number of Rows 32」の設定の行
        n_row_nums = sum(1 for x in row if _isnum(x))
        if n_row_nums <= 3 and run <= 6 and any(isinstance(x, str) and x.strip() and not _is_tn(x) for x in row):
            return True
        # T7 行ごとに桁の違う数を積んだ列（残りの最大が最小の 1,000 倍以上。100 倍だと 30〜12,000 の単価の表の桁違いまで黙る）は、外れ検査の相手にならない
        # （両端 1 割を除いて見る: 3 つ以上の列に 1 や 2 の小さな数が 1 つ混じるだけでは、桁の違う数を積んだ報告とは言えない）
        others_s = sorted(others)
        k_t = len(others_s) // 10
        others_t = others_s[k_t:len(others_s) - k_t] if k_t else others_s
        if others_t and max(others_t) >= min(others_t) * 1000:
            return True
        # 幅のある列でも、両端を除けば幅が狭く、この値が残りの最大の 100 倍以上（明白な桁違い）なら黙らない。
        # そうでない（残りの最大に近い・小さい側）なら、今までどおり黙る
        if others and max(others) >= min(others) * 1000 and not (v > 0 and av >= 100 * max(others)):
            return True
        # T5 小さい側の外れ（列の中央値の 1/10 以下）は、列が全部同符号のときだけ。増減・調整が混じる列は小口が普通
        med = sorted(others)[len(others) // 2] if others else 0
        if av and med and av * 10 <= med and any(x < 0 for rr, x in num_entries if rr != r):
            return True
        # 符号を反転した対（-663,933 と 663,933）は相殺で、打ち間違いではない
        if any(abs(x + v) <= av * 0.001 and rr != r for rr, x in num_entries if abs(rr - r) <= 1):
            return True
        # （合計の行がすぐ下にある表の中の数を内訳として黙らせると、合計に入った本物の桁違いまで黙るので、下の合計との照合はしない）
        # T1 日付のシリアル（3 万〜6 万の整数が列にほかにも 2 つ以上）／桁数のそろった整数の列（番号）
        if float(v).is_integer():
            ints = [x for rr, x in num_entries if rr != r and float(x).is_integer()]
            if 30000 <= v <= 60000 and sum(1 for x in ints if 30000 <= x <= 60000) >= 2:
                return True
            if len(ints) >= 4 and len({len(str(int(abs(x)))) for x in ints}) == 1 and len(str(int(av))) == len(str(int(abs(ints[0])))) and av >= 10000 and v > 0:
                return True
        # T1 見出しの番号: すぐ下が文字の見出し（'Measured'）で、上が空か文字
        below = values[r + 1][c] if r + 1 < num_rows and c < len(values[r + 1]) else None
        above = values[r - 1][c] if r >= 1 and c < len(values[r - 1]) else None
        if isinstance(below, str) and below.strip() and not _is_tn(below) and (above in (None, "") or isinstance(above, str)):
            return True
        return False

    def _is_breakdown(r, c, v):
        """同じ列の連続した数の並び（r を含む）の合計が、その真上の数に合う＝合計の内訳の明細（2026-11-05 dock の台帳:
        8,294,780 の下に 447,125・1,406,895・6,440,760 と並べた内訳）。"""
        def num(k):
            x = values[k][c] if 0 <= k < num_rows and c < len(values[k]) else None
            return float(x) if _isnum(x) else None
        for a in range(max(1, r - 9), r + 1):
            if num(a - 1) is None or num(a) is None:
                continue
            tot = 0.0
            for b in range(a, min(num_rows, a + 10)):
                if num(b) is None:
                    break
                tot += num(b)
                if b >= max(r, a + 1) and abs(tot - num(a - 1)) <= abs(num(a - 1)) * 0.005 and num(a - 1) != 0:
                    return True
        return False

    def _is_labelled_side_block(r, c):
        """3 つ以下の数が上下を空き（数でないセル）に挟まれ、各行の左に文字の見出しがある小さな別の塊（'Demand Charge 0.81'・
        'ACQ = 952,000'）。表の本体とは別の記入欄・控えで、本体の数と比べても桁違いの根拠にならない。"""
        def isn(k):
            return 0 <= k < num_rows and c < len(values[k]) and _isnum(values[k][c])
        a = r
        while isn(a - 1):
            a -= 1
        b = r
        while isn(b + 1):
            b += 1
        if b - a + 1 > 3:
            return False
        for k in range(a, b + 1):
            left = values[k][c - 1] if c >= 1 and c - 1 < len(values[k]) else None
            if not (isinstance(left, str) and left.strip() and not _is_tn(left)):
                return False
        return True

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
                # 日付の見出しの行（横に日付のシリアルが 3 つ以上並ぶ行）の日付も見出し（2026-11-05 初見の 60 本:
                # 値段の表の 7 行目の月の見出し 38,000 前後を、列ごとに「桁違い」と出していた）
                if r in date_rows and float(v).is_integer() and 30000 <= v <= 60000:
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
        # 同じ値が列に 2 つ以上ある（正規取引の刻み -0.005 が 2 行）のは打ち間違いではなく決まった値。
        # 左右の列にもマイナスが並ぶ表（増減・前年比・差額）では、マイナスは符号の打ち間違いの根拠にならない
        # （2026-11-05 dock の台帳: 刻み値・日本語の増減率の表）
        dup_vals = {v for _, v in num_entries if sum(1 for _r, x in num_entries if x == v) >= 2}
        signed_table = _neg_in_columns(c) >= 3
        # 「ぽつんと」は 2 個まで（2026-11-05 dock の実測: 410 行の値動きの列のマイナス 41 個を全部「符号の打ち間違い」と出した）
        if not signed and not sorted_col and len(pos) >= 4 and 1 <= len(neg) <= min(2, max(1, len(num_entries) // 10)):
            for r, v in neg:
                flagged.add(r)
                if (v in dup_vals or signed_table or _row_context_quiet(r, c, v)
                        or _outlier_context_quiet(r, c, v, num_entries)):
                    continue
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

            # 四分位範囲が 0 の列（ほぼ全部が同じ値＝1 や 4 の列）も、8 件以上で 7 割以上が同じ 0 でない値なら見る
            # （その中の 1 つが 1000 倍に化けた 1 → 1000 を拾う。2026-11-05 注入テスト: 同じ値の列の桁違いは全部見落としていた）
            flat_col = (iqr == 0 and n >= 8 and q1 != 0 and sum(1 for x in sorted_vals if x == q1) * 10 >= n * 7)
            if iqr > 0 or flat_col:
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
                far = (hi_far if len(hi_far) <= 1 else []) + (lo_far if len(lo_far) <= 1 else [])
                # 外れ値を除いた残りだけで 3 桁（1,000 倍）以上の幅がある列は、行ごとに意味の違う数を縦に積んだ報告
                # （件数 64・損益 -3,700 万・差 18.8 万を A 列のラベルで並べる）。桁違いを比べる相手がいない
                # （2026-11-05 初見の 60 本: e026 の 72 件・前の 60 本の e059 の控えの書式）
                far_rows = {r for r, _ in far}
                rest = [abs(v) for r, v in num_entries if r not in far_rows and v != 0]
                # 幅は両端 1 割を除いて見る（件数 3 つの列に 1 や 2 の小さな数が 1 つ混じるだけの列を、桁の違う数を積んだ報告と
                # 取り違えない。2026-11-05 注入テスト: 見落とし 17 件がこの除外で黙っていた）
                rest_s = sorted(rest)
                k_trim = len(rest_s) // 10
                rest_t = rest_s[k_trim:len(rest_s) - k_trim] if k_trim else rest_s
                if rest_t and max(rest_t) >= min(rest_t) * 1000:
                    far = []
                elif rest and max(rest) >= min(rest) * 1000:
                    # 両端を除けば幅は狭い。残りの最大の 100 倍以上の正の値だけ残す（明白な桁違い）
                    far = [(r_, v_) for r_, v_ in far if v_ > 0 and v_ >= 100 * max(rest)]
                for r, v in far:
                    if sorted_col and r not in cliff:
                        continue
                    if (v in dup_vals or (v < 0 and signed_table) or _row_context_quiet(r, c, v)
                            or _outlier_context_quiet(r, c, v, num_entries)):
                        continue
                    if _is_breakdown(r, c, v) or _is_labelled_side_block(r, c):
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

    # D. 番号の列の桁数の不一致（ID・コードの桁欠け）。列の整数の 9 割以上が同じ 6 桁以上で、1 つだけ桁数が違う
    # （2026-11-05 outlier の台帳の唯一の本物: 7 桁の Org ID の列の 1853）。金額の列は桁がまちまちなので 6 桁以上・9 割以上に限る
    already = {i["cell"] for i in issues if i["type"] == "outlier_value"}
    for c in range(num_cols):
        ints = [(r, values[r][c]) for r in range(num_rows)
                if c < len(values[r]) and _isnum(values[r][c]) and float(values[r][c]).is_integer() and values[r][c] > 0
                and not (30000 <= values[r][c] <= 60000)]
        if len(ints) < 6:
            continue
        lens = [len(str(int(x))) for _r, x in ints]
        mode = max(set(lens), key=lens.count)
        if mode < 6 or lens.count(mode) * 10 < len(lens) * 9:
            continue
        # 番号は連番で割り振るので、先頭 2 桁がそろう（10xxxxx）。金額は先頭の桁がばらばら
        heads = [str(int(x))[:2] for (_r, x), ln in zip(ints, lens) if ln == mode]
        if max(heads.count(h) for h in set(heads)) * 10 < len(heads) * 9:
            continue
        for (r, x), ln in zip(ints, lens):
            addr = _rowcol_to_a1(start_row + r, start_col + c)
            if ln != mode and addr not in already:
                issues.append({
                    "cell": addr,
                    "severity": "warning",
                    "type": "outlier_value",
                    "msg": f"値 {int(x)} は桁数が列と違います（同じ列の番号は {mode} 桁がそろっています・桁欠け・入力ミスの疑い）",
                    "value": x
                })

    return issues


# ----------------------------------------------------------------
# 診断の 1 件 → 直す手（dock と seiri が同じ表を引く・2026-11-05）
# ----------------------------------------------------------------
SHELF_FILL_FORMULA = '選んだ列の途切れた式を戻す'
SHELF_ADD_TOTAL = '表の下に合計行を足す'


def fix_of(issue: Dict[str, Any], total_row: Optional[bool] = None) -> Dict[str, str]:
    """診断の 1 件に添える「直す手」（純 Python）。
    kind: 'shelf'（棚のマクロ名を name に）／'tool'（道具のコマンドを name に）／'human'（人の判断。why に確かめること）。
    total_row: その行が合計・小計の行か（seiri は表を見て分かる。dock は分からないので None）。
    棚で直すのは、直し方が 1 つに決まる物だけ（式の列のずれ・合計の行）。答えが変わり得る物は人に回す。"""
    t = issue.get('type')
    if t in ('inconsistent_formula', 'value_in_formula_column') and total_row is not True:
        return {'kind': 'shelf', 'name': SHELF_FILL_FORMULA, 'why': '列の明細を選んで撃つ（見出しと合計の行は選ばない）'}
    if t in ('circular_reference', 'omitted_sum_range'):
        if total_row:
            return {'kind': 'shelf', 'name': SHELF_ADD_TOTAL, 'why': '合計の行を足し直す'}
        if total_row is None:
            return {'kind': 'human', 'name': '',
                    'why': f'表の下の合計の行なら shelf-run {SHELF_ADD_TOTAL}。部分集計なら作りかを確かめる'}
        return {'kind': 'human', 'name': '', 'why': '範囲に入れるつもりかを確かめる（部分集計・別の塊のことがある）'}
    if t in ('hidden_whitespace', 'text_number'):
        return {'kind': 'tool', 'name': 'diagnose --fix',
                'why': '前後の空白を除く・文字の数字を数にする（控えつき・承認の語が要る・先頭ゼロの番号は触らない）'}
    if t == 'formula_error':
        return {'kind': 'human', 'name': '', 'why': '下の「エラーの原因」を見て式を直す'}
    if t == 'hardcoded_constant':
        return {'kind': 'human', 'name': '', 'why': '変わり得る数なら条件のセルに出して参照する'}
    if t == 'approx_match':
        return {'kind': 'human', 'name': '', 'why': '完全一致のつもりなら FALSE（MATCH は 0）を足す。答えが変わり得る'}
    if t in ('outlier_value', 'outlier'):
        return {'kind': 'human', 'name': '', 'why': '入力の誤りか本当の値かを確かめる'}
    return {'kind': 'human', 'name': '', 'why': ''}


def fix_label(fx: Dict[str, str]) -> str:
    """直す手の見せ方（dock の表の 1 列）。"""
    if fx['kind'] == 'shelf':
        return f"shelf-run {fx['name']}"
    if fx['kind'] == 'tool':
        return fx['name']
    return '人: ' + (fx['why'] or '確かめる')


def audit_score(issues: List[Dict[str, Any]]) -> int:
    """100 点満点の点数（format_audit_report と dock が同じ数を使う）。"""
    n_crit = sum(1 for i in issues if i["severity"] == "critical")
    n_warn_kinds = len({(i["type"], i["msg"]) for i in issues if i["severity"] == "warning"})
    # 気づきは「同じ種類・同じ形の文（値を伏せた形）・同じシート」ごとに 1 回（2026-11-05）。
    # 同じ式を 10 行コピーした直書きの数・1 列の末尾空白は 1 回の直しで済むのに、1 セルごとに 2 点引くと、
    # 警告（文ごとに 5 点）より重くなっていた。列ごとに数えると、直書きの数が何十列に並ぶ表で 0 点になった（Enron の e039）
    def _info_kind(i):
        shape = re.sub(r'-?\d[\d,.]*(?:e[+-]?\d+)?', '#', re.sub(r"'[^']*'", "'…'", str(i.get("msg", ""))))
        sh = str(i.get("cell", "")).rpartition("!")[0]
        return (i["type"], shape, sh)
    n_info = len({_info_kind(i) for i in issues if i["severity"] == "info"})
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
            lines.append(f"  … ほか {len(infos) - 10} 件")
        lines.append("")

    # 推奨アクション
    lines.append("💡 【推奨アクション】")
    acts = []
    if crits:
        acts.append("集計範囲漏れ・数式エラーを優先して数式を修正してください。")
    if any(i["type"] == "hardcoded_constant" for i in issues):
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
    import unicodedata
    digits = unicodedata.normalize('NFKC', trimmed).replace('−', '-').replace(',', '').lstrip('-')
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
                num_val = _tn_value(trimmed)
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
