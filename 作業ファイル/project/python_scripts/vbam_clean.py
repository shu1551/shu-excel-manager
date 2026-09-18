# -*- coding: utf-8 -*-
"""vbam_clean.py — vba_manager 分割パート: clean-table（名簿の掃除を AI 抜きで回す）

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。規則の決め方は clean_plan（純 Python）、
実行は cmd_clean_table。材料の読み手・normalize は `vh.` 経由で vbam_hands を直に呼ぶ。
2026-09-12: 前は vbam_agent を経由していた（AI の機能が無いと掃除も動かなかった）。持ち主の vbam_hands へ付け替えた。
"""
import re
import unicodedata
import argparse
from vbam_core import (_col_letter, parse_target_and_rest, get_workbook)

import vbam_hands as vh


# ----------------------------------------------------------------
# clean-table（2026-09-08）: 名簿の掃除を AI 抜きで回す。
#   材料の「気づき」が既に汚れを数えているので、掃除は AI に文章で考えさせなくても手が決まる
#   （実測: agent 経由は sonnet の生成に 26 秒、この手は 3 秒台）。AI が要るのは「読みを作る」ような
#   値を発明する仕事だけで、空白の幅・カナの幅・全角数字・重複・日付の形は規則で決まる。
#   列ごとの規則は clean_plan（純 Python・COM を呼ばない＝テストできる）が決め、書き込みは
#   実績のある _do_normalize に渡す（文字が数値・日付に化けたら書き直す仕掛けがそこにある）。
# ----------------------------------------------------------------
_CLEAN_DATE_FMT = 'yyyy/mm/dd'


def _clean_norm_key(v):
    """重複を見るための正規化（表記ゆれを畳む）。"""
    if isinstance(v, str):
        return re.sub(r'\s+', ' ', unicodedata.normalize('NFKC', v)).strip()
    if isinstance(v, float) and float(v).is_integer():
        return int(v)
    return v


def clean_plan(grid, hdr_idx):
    """表の格子 → 掃除の計画（純 Python）。COM を呼ばない。

    grid: 値の 2 次元リスト（見出し行を含む）。hdr_idx: 見出し行の位置（0 始まり）。
    返り値: {'rules': {列index: [規則名…]}, 'dups': [(行index, 同じ行のindex)…],
             'blanks': [(行index, 列index)…], 'date_cols': [列index…]}
    規則の順は trim → kana_zenkaku → hankaku → space_* → date（hankaku が全角空白を半角にするので空白は後）。
    """
    rows = [list(r) for r in (grid or [])]
    if not rows or hdr_idx is None or hdr_idx >= len(rows):
        return {'rules': {}, 'dups': [], 'blanks': [], 'date_cols': []}
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([None] * (width - len(r)))
    heads = rows[hdr_idx]

    def blank(v):
        return v is None or (isinstance(v, str) and v.strip(vh._WS_CHARS) == '')

    body = [i for i in range(hdr_idx + 1, len(rows)) if not all(blank(v) for v in rows[i])]
    rules, date_cols = {}, []
    for j in range(width):
        col = [rows[i][j] for i in body]
        texts = [v for v in col if isinstance(v, str) and v != '']
        names = []
        if any(v.strip(vh._WS_CHARS) != v for v in texts):
            names.append('trim')
        if any(vh._HANKANA_ANY_RE.search(v) for v in texts):
            names.append('kana_zenkaku')
        if any(vh._ZEN_ALNUM_RE.search(v) for v in texts):
            names.append('hankaku')
        zen = han = multi = 0
        for v in texts:
            runs = vh._INNER_WS_RE.findall(v.strip(vh._WS_CHARS))
            if not runs:
                continue
            if any(len(w) > 1 for w in runs):
                multi += 1
            if any('　' in w for w in runs):
                zen += 1
            if any(ch in w for w in runs for ch in '  \t'):
                han += 1
        if zen or han:
            if zen and han:                       # 混在は多数派に寄せる
                names.append('space_zenkaku' if zen >= han else 'space_hankaku')
            elif multi:                           # 種類は揃っているが 2 つ以上ある＝1 つに詰める
                names.append('space_zenkaku' if zen else 'space_hankaku')
        n_date = sum(1 for v in col if vh._is_date_value(v))
        n_datestr = sum(1 for v in texts if vh._DATELIKE_RE.match(v.strip(vh._WS_CHARS)))
        if n_datestr and (n_date or n_datestr == len(texts)):
            names.append('date')
        if n_date + n_datestr >= 2 and not [v for v in col if not blank(v) and not vh._is_date_value(v)
                                            and not (isinstance(v, str) and vh._DATELIKE_RE.match(v.strip(vh._WS_CHARS)))]:
            date_cols.append(j)
        if names:
            rules[j] = names
    seen, dups = {}, []
    for i in body:
        key = tuple(_clean_norm_key(v) for v in rows[i])
        if key in seen:
            dups.append((i, seen[key]))
        else:
            seen[key] = i
    dup_set = {i for i, _f in dups}
    blanks = []
    for j in range(width):
        if blank(heads[j]):
            continue
        cells = [(i, rows[i][j]) for i in body if i not in dup_set]
        filled = sum(1 for _i, v in cells if not blank(v))
        if filled < 3 or filled / max(len(cells), 1) < 0.6:
            continue                              # 備考のようなまばらな列は「空欄」と数えない
        blanks += [(i, j) for i, v in cells if blank(v)]
    return {'rules': rules, 'dups': dups, 'blanks': blanks, 'date_cols': date_cols}


def cmd_clean_table(args):
    """表を掃除する（AI を使わない）: clean-table [範囲|セル] [--sheet 名] [--delete-dups] [--no-tidy]

    材料の「気づき」と同じ検出を使って、列ごとに規則を決めて当てる（空白の全角半角と連続・半角カナ・
    全角英数・文字の日付）。重複行は --delete-dups があるときだけ消す（承認の言葉に当たる）。
    日付の列は表示形式を yyyy/mm/dd に揃え、最後に tidy で見た目を仕上げる。
    表の中の空欄は埋めない（読み・金額・人の判断は道具が作ってはいけない）＝番地で報告する。
    範囲を省くと、見出し行から広がる表を道具が見つける。
    """
    from vbam_edit import cmd_tidy, _trim_trailing_empty       # 遅延 import（循環を避ける）
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    if rest:
        from vbam_view import _resolve_range
        ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
        if int(rng.Cells.CountLarge) == 1:
            rng = rng.CurrentRegion
    else:
        ws = wb.Sheets(sheet_opt) if sheet_opt else wb.ActiveSheet
        got = vh._guess_header_row(ws)
        if not got:
            print("エラー: 見出し行が見つかりません。範囲を指定してください（例 clean-table A5:E19）")
            return False
        rng = ws.Cells(got[0], int(ws.UsedRange.Column)).CurrentRegion
    rng, trimmed = _trim_trailing_empty(ws, rng)
    if trimmed:
        print(f"範囲の右端・下端の空の {trimmed} を外しました")
    r0, c0 = int(rng.Row), int(rng.Column)
    nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
    if nr < 2:
        print(f"エラー: {ws.Name}!{vh._addr(rng)} は表の形をしていません（見出し＋1 行以上）")
        return False
    grid = vh._rows_of(rng.Value)
    heads = vh._pick_header_row(grid)
    hdr_idx = heads[0] if heads else 0
    plan = clean_plan(grid, hdr_idx)
    print(f"対象: {ws.Name}!{vh._addr(rng)}（{nr} 行 × {nc} 列・見出しは {r0 + hdr_idx} 行目）")
    body_top, body_bottom = r0 + hdr_idx + 1, r0 + nr - 1
    if body_top > body_bottom:
        print("データの行がありません")
        return False

    done = []
    for j in sorted(plan['rules']):
        col = _col_letter(c0 + j)
        act = {'op': 'normalize', 'range': f"{col}{body_top}:{col}{body_bottom}",
               'rules': plan['rules'][j], 'overwrite': True}
        try:
            _label, ok, out = vh._do_normalize(ws, act)
        except Exception as ex:
            print(f"  {col}列 {'+'.join(plan['rules'][j])}: 失敗（{ex}）")
            continue
        n = re.search(r'(\d+) セルを変えた', out or '')
        done.append(f"{col}列 {'+'.join(plan['rules'][j])} → {n.group(1) if n else '?'} セル")
    if done:
        print("直した列:")
        for d in done:
            print("  " + d)
    else:
        print("直す列はありませんでした（空白・カナ・全角英数・日付の乱れなし）")

    for j in plan['date_cols']:                    # 日付の表示形式（混在・標準のときだけ揃える）
        col = _col_letter(c0 + j)
        try:
            cells = ws.Range(f"{col}{body_top}:{col}{body_bottom}")
            fmt = cells.NumberFormat
            if fmt is None or str(fmt).lower() in ('general', 'g/標準', 'standard'):
                cells.NumberFormat = _CLEAN_DATE_FMT
                print(f"日付の表示形式: {col}列を {_CLEAN_DATE_FMT} に揃えました")
        except Exception:
            continue

    if plan['dups']:
        lines = [f"行{r0 + i}（行{r0 + f} と同じ）" for i, f in plan['dups']]
        if getattr(args, 'delete_dups', False):
            for i, _f in sorted(plan['dups'], reverse=True):   # 下から消す（行番号がずれない）
                ws.Rows(r0 + i).Delete()
            print(f"重複行を削除: {len(plan['dups'])} 行（" + "／".join(lines) + "）")
            body_bottom -= len(plan['dups'])
        else:
            print(f"重複行が {len(plan['dups'])} 行あります（消していません。消すなら --delete-dups）: "
                  + "／".join(lines))

    if plan['blanks']:
        addrs = [f"{_col_letter(c0 + j)}{r0 + i}" for i, j in plan['blanks'][:20]]
        print(f"表の中の空欄が {len(plan['blanks'])} セル（埋めていません＝値は人か AI が決めるもの）: "
              + " ".join(addrs) + ("…" if len(plan['blanks']) > 20 else ""))

    if not getattr(args, 'no_tidy', False):
        ns = argparse.Namespace(posargs=[f"{_col_letter(c0)}{r0 + hdr_idx}:"
                                         f"{_col_letter(c0 + nc - 1)}{body_bottom}"],
                                sheet_opt=ws.Name, header_from=None, bg=None,
                                no_header=False, no_border=False, no_col_format=False,
                                no_autofit=False, min_width=None, max_width=None)
        from vbam_core import pinned_workbook
        with pinned_workbook(wb):        # 名指しされたブックに当てる（2026-09-08・自動検出だと別ブックへ流れる）
            cmd_tidy(ns)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


__all__ = ['clean_plan', 'cmd_clean_table']
