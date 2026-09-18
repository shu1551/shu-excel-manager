# -*- coding: utf-8 -*-
"""vbam_inv.py — vba_manager 分割パート: 不変条件（頼んでいない変化を控えと照らす）

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。控え（_snapshot_book）と照合（_inv_compare）、
手から決める免除（_inv_skip_for）。vbam_agent は `from vbam_inv import *` で名前を引き継ぐ。
"""
import re
import contextlib
from vbam_core import (_col_letter)
from vbam_hands import (_addr, _rows_of, _text_of)

_ROWCOL_SPLIT = {'row_insert': ('row', 'insert'), 'row_delete': ('row', 'delete'),
                 'col_insert': ('col', 'insert'), 'col_delete': ('col', 'delete')}


# ----------------------------------------------------------------
# 不変条件の検査（2026-09-05 制定）: 弾ごとの check は「頼んだことが当たったか」しか見ていない。
# 撃つ前の姿を控えておき、「頼んでいないところが変わっていないか」を全弾に自動で当てる。
#
# 2026-09-04 に本物のブックで出た欠陥（絞り込み中の行に書くと値がずれる／既存フィルタの範囲に
# 引きずられて別の列を絞る／隠れた行が ClearContents で消えない）は、どれも弾の check では
# 捕まらなかった。check が見るのは「F 列が期待どおりか」だけなので、同じ往復で隣の列を全角に
# 変えても・隠れた行を飛ばしても・別のシートに書いても、全部 PASS になる構造だった。
# ＝「当てる目」しか無く「壊していないかを見る目」が無かった。ここがその目。
# ----------------------------------------------------------------

_INV_MAX_CELLS = 200000        # 控えを取る上限（1 シート。超えたら値だけにして数式は見ない）


def _addr_rows(addr):
    """アドレス（A1:B2,A4:B5）に入っている行番号の集合。"""
    rows = set()
    for part in (addr or '').split(','):
        got = re.findall(r'[A-Z]*(\d+)', part)
        if len(got) >= 2:
            rows.update(range(int(got[0]), int(got[1]) + 1))
        elif got:
            rows.add(int(got[0]))
    return rows


def _snap_sheet(sh):
    """1 シートの姿（値・数式・隠れ行・条件付き書式・入力規則・フィルタ）。COM は 1 シート 6〜7 回。

    隠れ行は「行番号の集合」で持つ。可視範囲の文字そのままだと、列を 1 本足しただけで
    A1:E9 → A1:F9 になって誤検知する（突き合わせの弾＝F 列を足すのが依頼、で落ちる）。
    数式は「番地 → 数式」で持つ。位置で持つと、行が 1 行増えただけで全部ずれる。
    """
    snap = {'used': '', 'values': [], 'formulas': {}, 'hidden': (), 'cf': 0, 'validation': '', 'filter': '',
            'merged': None, 'tables': {}, 'charts': 0}
    try:
        ur = sh.UsedRange
        snap['used'] = used = _addr(ur)
        n = int(ur.Rows.Count) * int(ur.Columns.Count)
        snap['values'] = [[_text_of(v) for v in row] for row in _rows_of(ur.Value)]
        if n <= _INV_MAX_CELLS:
            snap['formulas'] = {_cell_name(r, c, used): f
                                for r, row in enumerate(_rows_of(ur.Formula))
                                for c, f in enumerate(row)
                                if isinstance(f, str) and f.startswith('=')}
        with contextlib.suppress(Exception):
            snap['hidden'] = tuple(sorted(_addr_rows(used) - _addr_rows(_addr(ur.SpecialCells(12)))))
        with contextlib.suppress(Exception):
            snap['cf'] = int(ur.FormatConditions.Count)
        with contextlib.suppress(Exception):
            snap['validation'] = _addr(ur.SpecialCells(-4174))          # xlCellTypeAllValidation
        with contextlib.suppress(Exception):
            # 結合の有無だけ（True＝全部結合 / False＝1 つも無い / None＝混在）。COM 1 回で済む。
            # 領域の一覧（vbam_view._merged_areas_in_range）はセルごとに COM を撃つので、
            # 全シートぶん取ると控えだけで数十秒かかる。「結合が丸ごと消えた」が拾えれば足りる
            snap['merged'] = ur.MergeCells
    except Exception:
        pass
    with contextlib.suppress(Exception):
        if sh.AutoFilterMode:
            snap['filter'] = _addr(sh.AutoFilter.Range)
    # テーブル（ListObject）の絞り込みは AutoFilterMode に出ない（2026-09-18 朝: テーブルを絞り込んだのに
    # 「隠れていた行が変わった」で AI に回っていた）
    if not snap['filter']:
        with contextlib.suppress(Exception):
            for lo in sh.ListObjects:
                with contextlib.suppress(Exception):
                    af = lo.AutoFilter
                    if af is not None and bool(af.FilterMode):
                        snap['filter'] = _addr(lo.Range)
                        break
    with contextlib.suppress(Exception):
        snap['tables'] = {str(lo.Name): _addr(lo.Range) for lo in sh.ListObjects}
    with contextlib.suppress(Exception):
        snap['charts'] = int(sh.ChartObjects().Count)
    # 2026-09-06 夜（輸入候補 3 の残り）: グラフの参照元（系列の式）と循環参照
    with contextlib.suppress(Exception):
        src = {}
        for co in sh.ChartObjects():
            forms = []
            with contextlib.suppress(Exception):
                for k in range(1, int(co.Chart.SeriesCollection().Count) + 1):
                    forms.append(str(co.Chart.SeriesCollection(k).Formula))
            src[str(co.Name)] = forms
        snap['chart_src'] = src
    with contextlib.suppress(Exception):
        cr = sh.CircularReference
        snap['circular'] = _addr(cr) if cr is not None else ''
    return snap


def _snap_sheet_light(sh):
    """触らないシートの軽い控え（2026-09-16）: 使用範囲の番地と Formula の一括読み 1 回だけ。

    大きいブック（秀コンボ）は 1 走行で控えを 3 回取り、_snap_sheet がシートごとに COM を 10 回前後
    （値・数式・隠れ行・条件付き書式・入力規則・結合・テーブル・グラフ・系列の式・循環参照）撃っていた＝
    道具の 28 秒の主因。触らないシートに要るのは「1 セルも変わっていないか」だけなので、Formula を 1 回
    読んで、定数はその値・数式セルは式の文字で持つ（結果が動いても式が同じなら変わっていない、の物差しは
    _inv_compare (1) と同じ）。隠れ行・条件付き書式などは見ない（'light' の印。own に昇格したシートの
    (3) 隠れ行の照合はこの印で外す＝控えに無いものを「消えた」と呼ばない）。
    """
    snap = {'used': '', 'values': [], 'formulas': {}, 'hidden': (), 'cf': 0, 'validation': '', 'filter': '',
            'merged': None, 'tables': {}, 'charts': 0, 'light': True}
    try:
        ur = sh.UsedRange
        snap['used'] = used = _addr(ur)
        grid = _rows_of(ur.Formula)
        snap['values'] = [[_text_of(v) for v in row] for row in grid]
        snap['formulas'] = {_cell_name(r, c, used): f
                            for r, row in enumerate(grid) for c, f in enumerate(row)
                            if isinstance(f, str) and f.startswith('=')}
    except Exception:
        pass
    return snap


def _snapshot_book(wb, own=None):
    """撃つ前のブックの姿。シートごとの中身と、ブックの名前定義・シートの並び・計算モード。

    own（シート名か名前の集合）を渡すと、それ以外のシートは軽い控え（_snap_sheet_light）にする（2026-09-16）。
    None なら全シートを丸ごと控える（従来どおり）。
    """
    own_set = None if own is None else ({own} if isinstance(own, str) else set(own or ()))
    snap = {'sheets': {}, 'names': {}, 'order': [], 'calc': None}
    for sh in wb.Worksheets:
        with contextlib.suppress(Exception):
            name = str(sh.Name)
            full = own_set is None or name in own_set
            snap['sheets'][name] = _snap_sheet(sh) if full else _snap_sheet_light(sh)
    with contextlib.suppress(Exception):
        for nm in wb.Names:
            with contextlib.suppress(Exception):
                snap['names'][str(nm.Name)] = str(nm.RefersTo)
    with contextlib.suppress(Exception):
        snap['order'] = [str(s.Name) for s in wb.Sheets]
    with contextlib.suppress(Exception):
        # 枠固定はウィンドウの持ち物＝読めるのはアクティブなシートだけ（他のシートは前に出さないと読めない）
        app = wb.Application
        if str(app.ActiveWorkbook.Name) == str(wb.Name):
            win = app.ActiveWindow
            snap['frozen'] = {str(app.ActiveSheet.Name):
                              (int(win.SplitRow), int(win.SplitColumn)) if win.FreezePanes else (0, 0)}
    with contextlib.suppress(Exception):
        snap['calc'] = int(wb.Application.Calculation)
    return snap


def _cell_name(r, c, used):
    """控えの (行, 列) を、使われている範囲の左上からの番地に直す（A1 起点とは限らない）。"""
    m = re.match(r'^([A-Z]+)(\d+)', used or 'A1')
    base_col, base_row = m.group(1) if m else 'A', int(m.group(2)) if m else 1
    n = 0
    for ch in base_col:
        n = n * 26 + (ord(ch) - 64)
    return f"{_col_letter(n + c)}{base_row + r}"


def _grid_diff(before, after, used, kind, limit=3, skip_cells=()):
    """2 つの控え（2 次元の文字）を比べ、違う番地を並べる。行数・列数が違えばそれも言う。

    skip_cells の番地は見ない（他シートの数式セル＝こちらの表を直せば結果が動くのは正当）。
    """
    out = []
    if len(before) != len(after):
        out.append(f"{kind}の行数が {len(before)} → {len(after)}")
    for r in range(min(len(before), len(after))):
        rb, ra = before[r], after[r]
        for c in range(min(len(rb), len(ra))):
            if rb[c] != ra[c] and _cell_name(r, c, used) not in skip_cells:
                out.append(f"{_cell_name(r, c, used)}: {rb[c]!r} → {ra[c]!r}")
                if len(out) >= limit:
                    return out
    return out


def _inv_violations(wb, before, own_sheet, skip=()):
    """撃った後のブックを控えと照らす（現物を読んで _inv_compare に渡すだけ）。

    控えが丸ごとだったシートは今回も丸ごと読む（軽い控えと丸ごとの控えを混ぜて比べない・2026-09-16）。
    """
    own = {own_sheet} if isinstance(own_sheet, str) else set(own_sheet or ())
    own |= {n for n, s in (before or {}).get('sheets', {}).items() if not (s or {}).get('light')}
    return _inv_compare(before, _snapshot_book(wb, own=own), own_sheet, skip)


# ブック全体でこれを超えたら不変条件の控えを取らない（読むだけで人を待たせる）。
# 実測（2026-09-05・本物のブックの台）: 671 セル 0.02 秒／47,450 セル 0.08 秒／92,228 セル 0.20 秒。
# 1 回の仕事で最大 3 回（撃つ前・関所・終わり）取るので、200 万セル＝1 回 4 秒・合計 13 秒を上限にする。
_INV_BOOK_MAX_CELLS = 2000000


def _book_cell_count(wb):
    """ブック全体の使用セル数（控えを取る前の見積り）。読めなければ 0。"""
    n = 0
    for sh in wb.Worksheets:
        with contextlib.suppress(Exception):
            ur = sh.UsedRange
            n += int(ur.Rows.Count) * int(ur.Columns.Count)
    return n


_INV_USED_GROW_ROWS = 5000     # 使用範囲がこれ以上広がったら暴走とみなす（列全体への fill・書式）


def _shape_changed(before_snap, now_snap):
    """表の形（行数・列数）が変わったか（純 Python）。変わっていれば番地での照合は当てにならない。"""
    b, a = (before_snap or {}).get('values') or [], (now_snap or {}).get('values') or []
    if len(b) != len(a):
        return True
    return max((len(r) for r in b), default=0) != max((len(r) for r in a), default=0)


def _box_shrank(before_addr, now_addr):
    """範囲が狭くなったか（行か列のどちらかが減った。純 Python）。読めなければ False。"""
    b, a = _range_box(before_addr or ''), _range_box(now_addr or '')
    if not b or not a:
        return False
    return (a[2] - a[0]) < (b[2] - b[0]) or (a[3] - a[1]) < (b[3] - b[1])


def _inv_skip_for(actions):
    """その回の手から、不変条件の検査で外すものを決める（純 Python）。

    弾ごとに人が書く `inv_skip`（実射用）は本番には無い。依頼どおりに動かした結果として当然変わるもの
    ——行を挿入すれば番地も行番号もずれる・絞り込めば行が隠れる・消してよいと言われた範囲の数式は消える
    ——を「頼んでいない変化」と呼ばないため、手の並びから機械的に決める（2026-09-05）。
    """
    skip = set()
    for a in (actions or []):
        if not isinstance(a, dict):
            continue
        op = str(a.get('op') or '')
        action = str(a.get('action') or '').strip().lower().replace('-', '_')
        if op in ('row', 'col', 'sort', 'dedupe') or op in _ROWCOL_SPLIT:
            skip.update(('hidden', 'formula', 'tables'))   # 行・列が動くと行番号も番地もテーブルの範囲もずれる
        elif op == 'view':
            skip.add('frozen')                       # 枠固定を変えるのが依頼（2026-09-06 夜）
        elif op == 'chart_config' and action == 'set_source':
            skip.add('chart_src')                    # グラフの参照元を変えるのが依頼
        elif op == 'autofilter':
            skip.add('hidden')                       # 絞り込みは行を隠すのが仕事
        elif op == 'clear_range' and a.get('overwrite'):
            skip.update(('formula', 'used'))         # 消してよいと言われている（使用範囲も縮む）
        elif op == 'sheet_op' and action == 'rename':
            skip.update(('other_sheet', 'order'))    # 改名は「シートが消えた」「並びが変わった」に見える
        elif op == 'format' and a.get('unmerge'):
            skip.add('merged')                       # 結合を外すのが依頼（帳票→一覧）
    return tuple(sorted(skip))


def _a1_to_rc(a):
    """'BR13' → (13, 70)。'$A$1' も可。読めなければ None（純 Python）。"""
    m = re.match(r'^\$?([A-Za-z]{1,3})\$?(\d+)$', str(a or '').strip())
    if not m:
        return None
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return int(m.group(2)), col


def _range_box(rng):
    """'A1:P13' → (1, 1, 13, 16)。単一セルなら 4 つとも同じ。読めなければ None（純 Python）。"""
    parts = str(rng or '').replace('$', '').split('!')[-1].split(':')
    a = _a1_to_rc(parts[0])
    b = _a1_to_rc(parts[-1]) if len(parts) > 1 else a
    if not a or not b:
        return None
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))


def _uncovered_formula_cells(formula_addrs, read_ranges):
    """読んだ範囲に入っていない数式セルの番地を返す（純 Python＝テストできる）。

    「見ていない範囲を『なし』と報告する」を道具が捕まえるため（2026-09-05・本物の
    161 列の表に当てて出た欠陥。AI は 5 ブロックだけ read して残り約 90 列を一度も
    見ずに「不整合なし・取りこぼしなし」と断言した。エラーが無いことと、見ていない
    ことが、同じ「なし」になっていた）。
    """
    boxes = [b for b in (_range_box(r) for r in (read_ranges or [])) if b]
    out = []
    for a in (formula_addrs or []):
        rc = _a1_to_rc(a)
        if rc is None:
            continue
        if not any(b[0] <= rc[0] <= b[2] and b[1] <= rc[1] <= b[3] for b in boxes):
            out.append(a)
    return out


def _sheet_formula_addrs(ws, limit=5000):
    """シートの数式セルの番地を返す（実物を読む。上限つき）。"""
    out = []
    try:
        for c in ws.UsedRange.SpecialCells(-4123):          # xlCellTypeFormulas
            try:
                out.append(str(c.Address).replace('$', ''))
            except Exception:
                pass
            if len(out) >= limit:
                break
    except Exception:
        pass
    return out


def _is_external_formula(f):
    """別のブックを指す数式か（'…\\[Book.xlsx]Sheet'!A1 ・ [Book.xlsx]Sheet1!A1 ・ [1]Sheet1!A1）。

    テーブルの構造化参照（=SUM(表1[金額])）は別ブックではないので含めない（純 Python）。
    """
    f = str(f or '')
    if re.search(r"\[[^\[\]]*\.xl\w*\]", f, re.IGNORECASE):
        return True
    return bool(re.search(r"(?:^|[^A-Za-z0-9_.])\[\d+\]", f))


def _inv_compare(before, now, own_sheet, skip=()):
    """2 つの控えを照らし、「頼んでいない変化」を並べて返す。空なら違反なし（純 Python＝テストできる）。

    own_sheet はシート名 1 つでも、名前の集合（この回に意図して書いたシート全部）でもよい
    （2026-09-05・書ける先を広げたので、意図して書いたシートを違反と呼ばないため）。
    skip に入れた語でその検査だけ外す（弾ごとの正当な例外。'other_sheet' 'names' 'hidden'
    'formula' 'cf' 'validation'）。
    """
    bad = []
    own = {own_sheet} if isinstance(own_sheet, str) else set(own_sheet or ())
    main = own_sheet if isinstance(own_sheet, str) else None
    # こちらの表の形が変わった回（行や列を足した・減らした）は、Excel が**他シートの参照を自分で直す**。
    # それを「数式が壊れた」と呼ぶと、正しく直した仕事が落ちる（2026-09-06・実射で踏んだ。
    # 練習台の状態「他シート参照」＝ =COUNTA(実射1!$A$1:$C$5) が行の挿入で $C$6 になっただけ）。
    # 形が変わっていない回に他シートの数式が変わっていたら、それは頼んでいない変化のまま。
    shifted = any(_shape_changed(before['sheets'].get(n), now['sheets'].get(n)) for n in own)

    # (1) 練習台以外のシートは 1 セルも変わってはいけない。
    #     AI がこの回に作ったシート（ピボットの置き場など）は控えに無い＝はじめから対象外。
    if 'other_sheet' not in skip:
        for name, b in before['sheets'].items():
            if name in own:
                continue
            a = now['sheets'].get(name)
            if a is None:
                bad.append(f"頼んでいないシート「{name}」が消えた")
                continue
            # 他シートの数式セルは、こちらの表を直せば結果が動く＝それは正当。数式そのものが
            # 変わった・消えた（#REF! に化けた）ときだけ咎める。
            d = _grid_diff(b['values'], a['values'], b['used'], '値',
                           skip_cells=set(b['formulas']) | set(a['formulas']))
            broke = [f"{ad}: {f} → {a['formulas'].get(ad) or '（消えた）'}"
                     for ad, f in b['formulas'].items() if a['formulas'].get(ad) != f
                     and not (shifted and '#REF!' not in str(a['formulas'].get(ad) or '')
                              and a['formulas'].get(ad))]
            if broke:
                d = d + ["数式が壊れた " + " ".join(sorted(broke)[:2])]
            if d:
                bad.append(f"頼んでいないシート「{name}」が変わった（{'／'.join(d[:3])}）")

    # (2) 名前定義が消えていないか（材料に出た名前を消さない＝手の型の完了条件）
    if 'names' not in skip:
        lost = [n for n in before['names'] if n not in now['names']]
        if lost:
            bad.append("名前定義が消えた: " + " ".join(lost[:4]))

    # (3)〜(5) は「書いたシート」ごとに見る（2026-09-05・書ける先を広げたので 1 枚とは限らない）
    for name in sorted(own):
        b_own, a_own = before['sheets'].get(name), now['sheets'].get(name)
        if not b_own or not a_own:
            continue
        head = "" if (main is not None or len(own) == 1) else f"「{name}」の"

        # (3) 隠れていた行が、頼んでいないのに現れた／消えた。
        #     ただし絞り込み（オートフィルタ）で隠れていた場合は、道具が「解いてから書く」のが仕様
        #     （_unfilter_for_write。絞り込みは見せ方であってデータではないので書く側が解いてよい）。
        #     そこを咎めると、正しく直した仕事まで落ちる（2026-09-05・実射で踏んだ）。手で隠した行だけ見る。
        #     撃った後に絞り込みが掛かっている場合も同じ＝絞り込みで隠れた行は見せ方（2026-09-18 朝:
        #     「0 の行は隠して」と頼んで絞り込んだのに「隠れていた行が変わった」で AI に回っていた）。
        if ('hidden' not in skip and not b_own.get('light') and not b_own['filter'] and not a_own.get('filter')
                and tuple(b_own['hidden']) != tuple(a_own['hidden'])):
            b_h, a_h = list(b_own['hidden']), list(a_own['hidden'])
            bad.append(f"{head}隠れていた行が変わった（{b_h or 'なし'} → {a_h or 'なし'} 行目）")

        # (4) 数式が値に潰されていないか（増える分＝新しく作った数式は正当）。
        #     ただし別ブックを指す数式（外部参照）が値になったのは咎めない。「リンクを切って値にする」は
        #     人が普通に頼む仕事で、そこを違反と呼ぶと正しく直した仕事まで落ちる（2026-09-05・実射で踏んだ。
        #     (3) の絞り込みと同じ型）。切った控えは AI の報告に出る。
        if 'formula' not in skip and b_own['formulas']:
            if _shape_changed(b_own, a_own):
                # 表の形が変わった回（行や列が増えた・減った）は、番地で照らしても意味が無い。
                # 合計の行が 1 行下がっただけで「C5 の数式が潰された」と言っていた（2026-09-06・
                # 弾「行を足すと合計が漏れる」で踏んだ。row の手なら inv_skip で外れるが、
                # 承認をもらって write_grid で書き直すと同じ形になる）。数だけで見る
                if len(a_own['formulas']) < len(b_own['formulas']):
                    bad.append(f"{head}数式の数が減った（{len(b_own['formulas'])} → "
                               f"{len(a_own['formulas'])}）＝値に潰した可能性")
            else:
                lost = [f"{a}: {f}" for a, f in b_own['formulas'].items()
                        if a not in a_own['formulas'] and not _is_external_formula(f)]
                if lost:
                    bad.append(f"{head}元からあった数式が値に潰された: " + " ".join(sorted(lost)[:3]))

        # (5) 条件付き書式・入力規則が減っていないか（足すのは正当・消すのは頼んでいない）
        if 'cf' not in skip and a_own['cf'] < b_own['cf']:
            bad.append(f"{head}条件付き書式が減った（{b_own['cf']} → {a_own['cf']}）")
        if 'validation' not in skip and b_own['validation'] and not a_own['validation']:
            bad.append(f"{head}入力規則が消えた（{b_own['validation']}）")

    # --- ここから 2026-09-06 に足した分（Claude for Excel の助言。減る・消える方だけ咎める） ---
    for name in sorted(set(before['sheets']) & set(now['sheets'])):
        b, a = before['sheets'][name], now['sheets'][name]
        head = f"「{name}」の" if (name not in own or len(own) > 1) else ""

        # (6) 結合が丸ごと消えた（帳票を黙って崩す事故。unmerge を頼まれた回は skip される）
        if 'merged' not in skip and b.get('merged') is not False and a.get('merged') is False:
            bad.append(f"{head}結合セルが全部外れた（頼んでいれば report に書く）")

        # (7) テーブルが消えた・範囲が狭くなった（増えるのは正当＝作る手がある）
        if 'tables' not in skip:
            for t, addr in (b.get('tables') or {}).items():
                now_addr = (a.get('tables') or {}).get(t)
                if now_addr is None:
                    # 同じ所に別の名前のテーブルがあれば、消えたのでなく名前を付け替えた（2026-09-18 朝:
                    # 「テーブル1」を T職員名簿 に直すのが仕事なのに「テーブルが消えた」で AI に回っていた）
                    box = _range_box(addr or '')
                    renamed = [n for n, ad in (a.get('tables') or {}).items()
                               if n not in (b.get('tables') or {}) and _range_box(ad or '')
                               and _range_box(ad or '')[:2] == (box[:2] if box else None)]
                    if renamed:
                        continue
                    bad.append(f"{head}テーブル「{t}」が消えた（{addr}）")
                elif _box_shrank(addr, now_addr):
                    bad.append(f"{head}テーブル「{t}」の範囲が狭くなった（{addr} → {now_addr}）")

        # (8) グラフが減った（作るのは正当）
        if 'charts' not in skip and int(a.get('charts') or 0) < int(b.get('charts') or 0):
            bad.append(f"{head}グラフが減った（{b.get('charts')} → {a.get('charts')}）")

        # (9) 使用範囲が暴走した（列全体への fill・書式で 100 万行に膨らむ。開く・保存・材料が全部重くなる）
        gb, ga = _range_box(b.get('used') or ''), _range_box(a.get('used') or '')
        if 'used' not in skip and gb and ga and (ga[2] - gb[2]) > _INV_USED_GROW_ROWS:
            bad.append(f"{head}使用範囲が {ga[2] - gb[2]:,} 行ぶん広がった（{b.get('used')} → {a.get('used')}）")

    # (10) もとからあったシートの並びが入れ替わった（足すのは正当＝間に入っても順序は保たれる）
    if 'order' not in skip and before.get('order') and now.get('order'):
        keep = [s for s in now['order'] if s in set(before['order'])]
        if keep != [s for s in before['order'] if s in set(now['order'])]:
            bad.append("シートの並びが変わった（" + " ".join(before['order']) + " → " + " ".join(now['order']) + "）")

    # (11) 計算モードが変わった（手動のまま返すと、以後この人の全部の数式が更新されない）
    if 'calc' not in skip and before.get('calc') is not None and now.get('calc') != before.get('calc'):
        names = {-4105: '自動', -4135: '手動', 2: 'データテーブル以外は自動'}
        bad.append(f"計算モードが変わった（{names.get(before['calc'], before['calc'])}"
                   f" → {names.get(now.get('calc'), now.get('calc'))}）")

    # --- 2026-09-06 夜に足した分（輸入候補 3 の残り）。減る・壊れる方だけ咎める ---
    # (12) 枠固定が外れた（view の手を使った回は skip。読めるのはアクティブなシートだけ）
    if 'frozen' not in skip:
        for name, fb in (before.get('frozen') or {}).items():
            fa = (now.get('frozen') or {}).get(name)
            if fa is not None and tuple(fb) != (0, 0) and tuple(fa) == (0, 0):
                bad.append(f"「{name}」の枠固定が外れた（上 {fb[0]} 行・左 {fb[1]} 列 → なし）")
    for name in sorted(set(before['sheets']) & set(now['sheets'])):
        b, a = before['sheets'][name], now['sheets'][name]
        # (13) グラフの参照元が壊れた・系列が減った（増える・行の挿入でずれるのは正当。set-source の回は skip）
        if 'chart_src' not in skip:
            for cname, forms in (b.get('chart_src') or {}).items():
                got = (a.get('chart_src') or {}).get(cname)
                if got is None:
                    continue                                 # 消えたグラフは (8) が数える
                if len(got) < len(forms):
                    bad.append(f"グラフ「{cname}」の系列が減った（{len(forms)} → {len(got)}）")
                elif any('#REF!' in f for f in got) and not any('#REF!' in f for f in forms):
                    bad.append(f"グラフ「{cname}」の参照元が #REF! に化けた")
        # (14) 循環参照ができた（無かったシートに出たときだけ）
        if 'circular' not in skip and (a.get('circular') or '') and not (b.get('circular') or ''):
            bad.append(f"「{name}」に循環参照ができた（{a.get('circular')}）")

    return bad


__all__ = ['_INV_BOOK_MAX_CELLS', '_INV_MAX_CELLS', '_INV_USED_GROW_ROWS', '_ROWCOL_SPLIT', '_a1_to_rc', '_addr_rows', '_book_cell_count', '_box_shrank', '_cell_name', '_grid_diff', '_inv_compare', '_inv_skip_for', '_inv_violations', '_is_external_formula', '_range_box', '_shape_changed', '_sheet_formula_addrs', '_snap_sheet', '_snap_sheet_light', '_snapshot_book', '_uncovered_formula_cells']
