# -*- coding: utf-8 -*-
"""vbam_prefire.py — vba_manager 分割パート: マクロの先撃ち（agent の入口で、AI より先に「表を整える」を撃つ）

2026-09-11 深夜。試験（agent --exam）で、表の整理はマクロ（表の整理.bas・AI なし）が 45 題全問合格・1 題 0.1〜0.5 秒、
agent（Sonnet）は 1 題平均 22 秒だった。shu と決めた形: 入口は agent のまま、中の順番を「マクロが先、Sonnet は例外」に。
  1. 依頼が書き方をそろえる整理（そろえ・統一・整え・整理・表記・きれい…）で、開いているブックかアドインに
     「表を整える」があれば、道具が撃つ（依頼が重複を消すことを承認していれば「重複行を消す」も）。控えは本番と同じ。
  2. 道具の tidy で仕上げ、終わりの検査（仕上げ検査・中身・依頼の語の列の残り・消えた行・頼んでいない変化）を当てる。
  3. 残りが無く、依頼にマクロの扱わない仕事（合計・並べ替え・グラフ…）も無ければ、そこで合格（AI は呼ばない）。
     残りがあれば、その指摘を材料に足して Sonnet の往復へ（控え・関所・undo はそのまま効く）。
どちらを撃つかは依頼の語で機械的に決まる（AI にも会話の Claude にも判断させない）。

2026-09-17 鍛える回路（vbam_forge）: モジュール「表の整理」の Sub の頭 2 行（' 依頼の語: 正規表現 ／ ' 扱う: 語 語）が
登録簿。依頼に「依頼の語」が当たる Sub は「表を整える」の後に撃ち、その「扱う」の語は AI に回す仕事（_ASK_OTHER_RE）
から外す。登録簿は .bas そのもの（別の台帳を持たない＝データはファイル自身）。
"""
import contextlib
import re
import time
from vbam_core import _col_letter

import vbam_agent as va

_MACRO_TIDY = '表を整える'
_MACRO_DEDUPE = '重複行を消す'
_MACRO_MODULE = '表の整理'
_ASK_TIDY_RE = re.compile(r'そろえ|揃え|統一|整え|整理|きれいに|綺麗に|掃除|書き方|表記')
# マクロの扱わない仕事（あれば、マクロの後に Sonnet の往復へ回す）
_ASK_OTHER_RE = re.compile(r'合計|集計|小計|ピボット|グラフ|並べ替|並び替|ソート|抽出|突き合|照合|列を足|列を追加|列を挿入|'
                           r'行を足|行を追加|行を挿入|追加し|印刷|PDF|色を|色分け|色付け|網掛け|塗り|条件付き|数式|関数|VLOOKUP|テーブルに|'
                           r'フィルタ|図形|写真|画像|埋め|計算|平均|件数|分けて|分割|区分|判定|比較|差分|ふりがなを|フリガナを付')
_DUP_WORD_RE = re.compile(r'重複|二重|ダブ|同じ内容|同じ会員|同じ品番|同じ行')
# マクロの扱わない仕事のうち、表の中身の仕事と別の操作（同じ節に「扱う」の語があっても片づけない）
_STRONG_OTHER_RE = re.compile(r'グラフ|並べ替|並び替|ソート|印刷|PDF|色を|色分け|色付け|網掛け|塗り|条件付き|抽出|フィルタ|図形|写真|画像')
_NO_DELETE_RE = re.compile(r'(消|削除)さない|(消|削除)しない|残して')


_REG_ASK_RE = re.compile(r"^\s*'\s*依頼の語\s*[:：]\s*(.+?)\s*$")
_REG_HANDLE_RE = re.compile(r"^\s*'\s*扱う\s*[:：]\s*(.+?)\s*$")
_REG_HEAD_RE = re.compile(r"^\s*'\s*見出し\s*[:：]\s*(.+?)\s*$")
_REG_SELECT_RE = re.compile(r"^\s*'\s*選ぶ列\s*[:：]\s*(\d+)\s*(?:[-〜~]\s*(\d+))?\s*$")
_REG_SHAPE_RE = re.compile(r"^\s*'\s*形\s*[:：]\s*(.+?)\s*$")
# 表の形の語（登録簿の ' 形: に書ける語。入口はシートの形を読んで、そろっている仕事だけを候補にする・2026-09-18 朝）
SHAPE_WORDS = ('数の列', '日付の列', '率の列', 'テーブル', '別のシート', '同じ見出しのシート', '共通の見出しのシート', '結合の見出し', 'ピボット', 'グラフ')
# ピボット・グラフ＝アクティブなシートにそれがあるか（2026-09-18 第二期: ピボットを値で貼る・累計・見せ方、グラフの見せ方の仕事が、
# ピボットやグラフの無いシートでも依頼の語が当たると撃たれて何もせずに終わり、その依頼が AI に回らなかった）。
# 鍛える側は本番の表にあるときだけ必要条件にする（何もしない試験の表があっても落とさない・vbam_forge._job_shape）
OBJECT_SHAPE_WORDS = ('ピボット', 'グラフ')
_SHEET_WORD_ROWS = 15      # 見出しの語を探す、シートの上からの行数
_REG_SUB_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(', re.I)
_REG_HEAD_LINES = 6        # Sub の行の下、この行数までにある頭の 2 行を登録簿として読む


def registry_from_text(text, owner=None):
    """モジュールの本文 → 登録簿 [{'name','owner','ask'(正規表現の文字),'handles'([語]),'headers'([語])}]（純 Python）。
    「表を整える」「重複行を消す」は道具が持つ既定なので入れない。頭の 2 行が無い Sub も入れない。
    ' 見出し: の行（任意）＝そのマクロが要る見出しの語。シートに全部あるときだけ撃つ（2026-09-17 夕）。
    ' 見出し: なし ＝語に頼らないマクロ（2026-09-18）。headers は []。"""
    out = []
    lines = str(text or '').splitlines()
    for i, ln in enumerate(lines):
        m = _REG_SUB_RE.match(ln)
        if not m or m.group(1) in (_MACRO_TIDY, _MACRO_DEDUPE):
            continue
        ask = handles = None
        headers = []
        select = None
        shape = []
        for h in lines[i + 1:i + 1 + _REG_HEAD_LINES]:
            a = _REG_ASK_RE.match(h)
            b = _REG_HANDLE_RE.match(h)
            c = _REG_HEAD_RE.match(h)
            s = _REG_SELECT_RE.match(h)
            k = _REG_SHAPE_RE.match(h)
            if a:
                ask = a.group(1)
            elif b:
                handles = b.group(1).split()
            elif c:
                # 「なし」＝見出しの語に頼らないマクロ（表の形で列を探す・2026-09-18）＝どの表にも撃つ
                headers = [] if c.group(1).strip() == 'なし' else c.group(1).split()
            elif s:
                # ' 選ぶ列: N（や N-M）＝選んでいる列を使うマクロ。入口は依頼の語と見出しが一致する列を選んでから撃つ（2026-09-18）
                select = (int(s.group(1)), int(s.group(2) or s.group(1)))
            elif k:
                # ' 形: 語 語 ＝この仕事に要る表の形（数の列・日付の列・テーブル・同じ見出しのシート…）。「なし」は []
                shape = [] if k.group(1).strip() in ('なし', '無し') else [w for w in k.group(1).split() if w in SHAPE_WORDS]
            elif not h.strip().startswith("'"):
                break
        if not ask or not handles:
            continue
        try:
            re.compile(ask)
        except re.error:
            continue
        out.append({'name': m.group(1), 'owner': owner, 'ask': ask, 'handles': handles, 'headers': headers,
                    'select': select, 'shape': shape})
    return out


def head_key(s):
    """見出しの語のならし（純 Python）: 全角半角をそろえ、空白・改行を落とし、末尾の単位の括弧（（円）など）を落とす。
    2026-09-17 夜: 実物らしい表で「勘定科目\\nコード」「科　目　名」「支出額（円）」の見出しに登録簿のマクロを撃たず AI に回した。"""
    import unicodedata
    t = re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(s or '')))
    # 落とすのは単位の括弧だけ（「支出額（税抜）」「予算額（前年度）」は別の列＝同じ見出しにしない）
    return re.sub(r'[\(\[](?:単位[:：]?)?(?:円|千円|百万円|人|名|件|計|合計|%|％)[\)\]]$', '', t) or t


_CELL_ADDR_RE = re.compile(r'(?<![A-Za-z0-9])([A-Z]{1,3}[1-9]\d*)(?![0-9])')


def old_findings(notes, changed):
    """指摘の文のうち、書いてある番地がどれも「変わったセル」に無いもの（純 Python）。番地の無い指摘は古いと言わない。"""
    out = []
    for n in notes or ():
        cells = set(_CELL_ADDR_RE.findall(str(n)))
        if cells and not (cells & set(changed)):
            out.append(n)
    return out


def _split_old_findings(audit, c_block, before, ws, sheet):
    """→ (audit, c_block, 元の表の指摘)。変わったセルが数えられなければ何も動かさない。"""
    try:
        changed = {r['addr'] for r in (va._changes_of(before, ws, sheet) or [])}
    except Exception:
        return audit, c_block, []
    if not changed:
        return audit, c_block, []
    old = old_findings(list(audit) + list(c_block), changed)
    return ([b for b in audit if b not in old], [b for b in c_block if b not in old], old)


def sheet_words(values):
    """シートの上の行の値（2 次元）→ 見出しの語の集合（純 Python）。前後の空白を落とした語と、ならした語（head_key）の両方。"""
    words = set()
    for row in list(values or [])[:_SHEET_WORD_ROWS]:
        for v in (row or []):
            if isinstance(v, str) and v.strip():
                words.add(v.strip())
                words.add(head_key(v))
    return words


def _header_row_index(kinds):
    """見出しの行（文字が 2 つ以上並ぶ最初の行）の添字（純 Python）。無ければ 0。"""
    for i, k in enumerate(kinds or []):
        if str(k).count('s') >= 2:
            return i
    return 0


def _rate_like(col_texts):
    """数の列の値が全部 ±1.5 以内で小数を含む＝率の列（vbam_edit.summable と同じ物差し・純 Python）。"""
    nums = []
    for t in col_texts:
        try:
            nums.append(float(str(t).replace(',', '').replace('%', '')))
        except (TypeError, ValueError):
            continue
    return bool(nums) and all(-1.5 <= v <= 1.5 for v in nums) and any(not float(v).is_integer() for v in nums)


def shape_signals(values, kinds, has_table=False, other_heads=(), merged_head=False, has_pivot=False, has_chart=False):
    """表の形の語の集合（純 Python）。values＝文字の 2 次元・kinds＝行ごとの型の文字列（n 数・d 日付・s 文字・- 空）。
    other_heads＝ほかのシートの見出しの語の集合の並び。入口は登録簿の \' 形: がこの集合に全部あるときだけ候補にする。
    2026-09-18 朝: 25 本を全部「見出し: なし」にしたら語のかすり当たりで別マクロも撃たれた＝形で先に候補を絞る。"""
    out = set()
    kinds = [str(k) for k in (kinds or [])]
    values = [list(r or []) for r in (values or [])]
    h = _header_row_index(kinds)
    body = [k for k in kinds[h + 1:] if k.strip('-')]
    ncols = max((len(k) for k in kinds), default=0)
    for c in range(ncols):
        col = [k[c] for k in body if c < len(k) and k[c] != '-']
        if not col:
            continue
        n = sum(1 for x in col if x == 'n')
        d = sum(1 for x in col if x == 'd')
        if n and n >= len(col) * 0.8:
            out.add('数の列')
            texts = [r[c] for r in values[h + 1:] if c < len(r) and r[c] not in (None, '')]
            if _rate_like(texts):
                out.add('率の列')
        if d and d >= len(col) * 0.8:
            out.add('日付の列')
    if has_table:
        out.add('テーブル')
    if merged_head:
        out.add('結合の見出し')
    if has_pivot:
        out.add('ピボット')
    if has_chart:
        out.add('グラフ')
    heads = set()
    if h < len(values):
        heads = {head_key(v) for v in values[h] if isinstance(v, str) and v.strip()}
    for oh in other_heads or ():
        oh = {head_key(v) for v in (oh or ()) if isinstance(v, str) and v.strip()}
        if not oh:
            continue
        out.add('別のシート')
        if heads and oh == heads:
            out.add('同じ見出しのシート')
        if heads & oh:
            out.add('共通の見出しのシート')
    return out


def _shape_of_sheet(ws, wb, rows=60):
    """アクティブなシートの形（COM）→ shape_signals の集合。読めなければ空（＝形の要る仕事は候補にしない・安全側）。"""
    from vbam_hands import _rows_of
    try:
        from vbam_forge import _kind_of, _texts
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        nr = min(int(ur.Rows.Count), rows)
        nc = int(ur.Columns.Count)
        raw = _rows_of(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1)).Value) or []
        values = _texts(raw)
        kinds = [''.join(_kind_of(v) for v in row) for row in raw]
        has_table = False
        try:
            has_table = int(ws.ListObjects.Count) > 0
        except Exception:
            pass
        merged = False
        try:
            # 見出しの行〜+2 行までは必ず見る（見出しが下にある表でも結合の見出しを取りこぼさない・2026-09-18）
            last = min(r0 + max(4, _header_row_index(kinds) + 2), r0 + nr - 1)
            top = ws.Range(ws.Cells(r0, c0), ws.Cells(last, c0 + nc - 1))
            merged = top.MergeCells is True or top.MergeCells is None      # None＝混在（結合がある）
        except Exception:
            pass
        others = []
        try:
            for sh in list(wb.Worksheets)[:12]:
                if str(sh.Name) == str(ws.Name):
                    continue
                ur2 = sh.UsedRange
                nr2 = min(int(ur2.Rows.Count), _SHEET_WORD_ROWS)
                rows2 = _rows_of(sh.Range(sh.Cells(int(ur2.Row), int(ur2.Column)),
                                          sh.Cells(int(ur2.Row) + nr2 - 1, int(ur2.Column) + int(ur2.Columns.Count) - 1)).Value) or []
                k2 = [''.join(_kind_of(v) for v in row) for row in rows2]
                hi = _header_row_index(k2)
                if hi < len(rows2):
                    others.append([v for v in rows2[hi] if isinstance(v, str)])
        except Exception:
            pass
        has_pivot = has_chart = False
        try:
            has_pivot = int(ws.PivotTables().Count) > 0
        except Exception:
            pass
        try:
            has_chart = int(ws.ChartObjects().Count) > 0
        except Exception:
            pass
        return shape_signals(values, kinds, has_table, others, merged, has_pivot, has_chart)
    except Exception:
        return set()


def dedupe_registry(entries, prefer=()):
    """同じ名前の Sub は 1 本だけにする（純 Python）。prefer（開いているブックの名前）の持ち主を先に採る。
    2026-09-17: 秀コンボ.xlsm に登録 → 更新登録で .xlam にも入ると、登録簿に同じ Sub が 2 本並び、
    先撃ちが 2 回撃つ（合計行が二重になる）。開いている .xlsm のほうが新しい（更新登録の前）ので先に採る。"""
    prefer = set(prefer or ())
    ordered = sorted(enumerate(entries or ()), key=lambda ie: (ie[1].get('owner') not in prefer, ie[0]))
    seen, keep = set(), set()
    for i, e in ordered:
        if e['name'] in seen:
            continue
        seen.add(e['name'])
        keep.add(i)
    return [e for i, e in enumerate(entries or ()) if i in keep]


def registry_of(xl):
    """開いているブック／アドインの「表の整理」から登録簿を読む（COM。読めなければ []）。"""
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return []
    out = []
    for p in projects:
        try:
            cm = p.VBComponents(_MACRO_MODULE).CodeModule
            n = int(cm.CountOfLines)
            if not n:
                continue
            owner = _owner_name(xl, p)
            if not owner:
                continue
            out += registry_from_text(cm.Lines(1, n), owner)
        except Exception:
            continue
    books = []
    try:
        books = [str(w.Name) for w in xl.Workbooks]
    except Exception:
        pass
    return dedupe_registry(out, books)


def _top_words(ws):
    """シートの上 _SHEET_WORD_ROWS 行の見出しの語（COM）。
    2026-09-17: UsedRange.Resize(n) は win32com の動的呼び出しで通らず、例外 → 安全側で「撃たない」に倒れていた。"""
    from vbam_hands import _rows_of
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    nr = min(int(ur.Rows.Count), _SHEET_WORD_ROWS)
    nc = int(ur.Columns.Count)
    top = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1))
    return sheet_words(_rows_of(top.Value))


def head_in(h, words):
    """見出しの語 h（「課名|行ラベル」のように | で並べた言い換えもよい）が、シートの見出しの語 words にあるか（純 Python）。
    2026-09-18: ピボットを値で貼り付けた集計表は「課名」でなく「行ラベル」＝全部の表に共通の語が無く、見出しの行を書けなかった。"""
    return any(a and (a in words or head_key(a) in words) for a in str(h).split('|'))


def pick_columns(request, cells, want=None):
    """依頼文と見出しのセル（[(列番号, 見出しの文字)]）→ 依頼に出てくる順に並べた列番号（純 Python）。
    want＝(最小, 最大) を渡すと、その数でないときは None（＝入口は撃たずに AI へ回す）。
    2026-09-18: 列の選び方が要る仕事（ピボットの行・列・値など）は、人が列を選んでから撃つ決まりにした。
    入口では依頼の語と見出しが一致する列を道具が選ぶ。"""
    req = str(request or '')
    key_req = head_key(req)
    all_cells = [(int(c), str(t or '').strip()) for c, t in (cells or ()) if str(t or '').strip()]

    def _match(cs):
        hits = []
        for col, t in cs:
            at = req.find(t)
            if at < 0:
                k = head_key(t)
                at = key_req.find(k) if k else -1
            if at >= 0:
                hits.append((at, col))
        # 同じ見出しの列が 2 つ以上あるとき（前の結果の表が右にある・2026-09-18）は、依頼の語 1 つにつき左の列を採る
        by_at = {}
        for at, col in hits:
            by_at.setdefault(at, []).append(col)
        out = []
        for at in sorted(by_at):
            for col in sorted(by_at[at]):
                if col not in out:
                    out.append(col)
                    break
        return out

    cols = _match(all_cells)
    if want and len(cols) > int(want[1]):
        # 当たりすぎ＝右に前の結果の表・別の表がある。空の列で切れた左の塊だけで選び直す（2026-09-18 朝:
        # 依頼文の「4月」「3月」「合計」が、右に残った前の集計表の見出しに当たり、月次集計が AI に回っていた）
        run, prev = [], None
        for col, t in sorted(all_cells):
            if prev is not None and col > prev + 1:
                break
            run.append((col, t))
            prev = col
        narrowed = _match(run)
        if int(want[0]) <= len(narrowed) <= int(want[1]):
            cols = narrowed
    if want and not (int(want[0]) <= len(cols) <= int(want[1])):
        return None
    return cols


def _header_cells(ws):
    """シートの見出しの行のセル（COM）→ [(列番号, 見出しの文字)]。読めなければ []。"""
    from vbam_hands import _rows_of
    try:
        anchor = _table_anchor(ws)
        if not anchor:
            return []
        cell = ws.Range(anchor)
        r, c0 = int(cell.Row), int(cell.Column)
        ur = ws.UsedRange
        c2 = int(ur.Column) + int(ur.Columns.Count) - 1
        row = (_rows_of(ws.Range(ws.Cells(r, c0), ws.Cells(r, c2)).Value) or [[]])[0]
        return [(c0 + j, str(v).strip()) for j, v in enumerate(row) if isinstance(v, str) and str(v).strip()]
    except Exception:
        return []


def current_columns(ws):
    """いま選んでいるセルの列（COM）→ [列番号]（選んだ順・重複なし）。そのシートを選んでいなければ []。
    2026-09-18: 列の選び方が要る仕事は「人が選んだ列」を使う。人が選んでいればそれを使い、
    選んでいなければ入口が依頼の語と見出しから選ぶ。"""
    try:
        sel = ws.Application.Selection
        if str(sel.Parent.Name) != str(ws.Name):
            return []
        cols = []
        for a in sel.Areas:
            c0, n = int(a.Column), int(a.Columns.Count)
            if n > 3:
                continue                                 # 行全体・シート全体の選択は「列を選んだ」と見ない
            for c in range(c0, c0 + n):
                if c not in cols:
                    cols.append(c)
        return cols
    except Exception:
        return []


def select_columns(ws, cols):
    """見出しのセルを選ぶ（COM）。選んだ列がマクロの入り口（Selection）になる。"""
    if not cols:
        return False
    try:
        anchor = _table_anchor(ws)
        r = int(ws.Range(anchor).Row) if anchor else 1
        ws.Range(",".join(str(ws.Cells(r, c).Address) for c in cols)).Select()
        return True
    except Exception:
        return False


def plan_full(request, registry=(), words=None, shape=None):
    """依頼文＋登録簿（＋シートの見出しの語）→ {'fire','dedupe','other','extras','left','skipped'}（純 Python）。
    fire    … 先撃ちするか（整理の語があるか、登録簿の依頼の語が当たるか）
    dedupe  … 重複行を消すも撃つか
    extras  … 登録簿から撃つ Sub（順は登録簿の順）
    left    … 依頼にあるマクロの扱わない仕事の語のうち、撃つ Sub の「扱う」で片づかないもの
    other   … left が空でないか（＝AI の往復へ回す）
    skipped … 依頼の語は当たったが、シートに要る見出し（words）か表の形（shape）が無くて撃たない Sub
    """
    req = str(request or '')
    tidy = bool(_ASK_TIDY_RE.search(req))
    dedupe = bool(tidy and _DUP_WORD_RE.search(req) and va._has_approval(req) and not _NO_DELETE_RE.search(req))
    extras, skipped = [], []
    for e in registry or ():
        strength = _ask_strength(e['ask'], req)
        if not strength:
            continue
        # 要る見出しがシートに無い表には撃たない（「集計」の語だけで別の表に撃ち、集計を AI に回さなくなる誤爆を防ぐ）
        if words is not None and e.get('headers') and not all(head_in(h, words) for h in e['headers']):
            skipped.append(dict(e, strength=strength))
            continue
        # 表の形（' 形:）がそろわない表には撃たない（2026-09-18 朝: 語のかすり当たりで別マクロが撃たれるのを形で先に絞る）
        if shape is not None and e.get('shape') and not set(e['shape']) <= set(shape):
            skipped.append(dict(e, strength=strength, why_shape=[w for w in e['shape'] if w not in shape]))
            continue
        extras.append(dict(e, strength=strength))
    # 依頼が名指ししている仕事は、より長い語で当たった方（2026-09-17 夕: 振り直しの依頼の「集計」が集約のマクロに当たり、
    # 集約のブックでは見出しもそろうので撃たれた。依頼は「管理会計」で振り直しを名指ししていた）。
    # 見出しがそろわず撃たなかった仕事の語の方が長ければ、短い語で当たったマクロは撃たない＝AI に回す
    # 同じ長さなら、どちらを名指ししたか決められない＝撃たない（AI に回す＝安全側）
    # ただし「合計」「集計」のような作業の一般語で当たっただけの仕事は、名指しに数えない
    # （2026-09-17 夜: 「各課シートの事業を集めて合計も出して」が、見出しの無い合計行を足すの「合計」と引き分けて集約を撃たなかった）
    top_skipped = max((e['strength'] for e in skipped if not _ASK_OTHER_RE.fullmatch(_ask_best_word(e['ask'], req))),
                      default=0)
    # 同じ長さなら、見出しがそろっている方（extras）を撃つ（2026-09-18 未明: 「振り直して、区分ごとの円グラフも」で、見出しの無い
    # 円グラフの「円グラフ」と振り直しの「管理会計」が同じ 4 字で引き分け、どちらも撃たなかった）。長い語で名指しされた別の仕事は従来どおり
    shadowed = [e for e in extras if top_skipped and e['strength'] < top_skipped]
    extras = [e for e in extras if not (top_skipped and e['strength'] < top_skipped)]
    others = sorted({m.group(0) for m in _ASK_OTHER_RE.finditer(req)})
    covered = [h for e in extras for h in (e.get('handles') or [])]
    # 同じ節（。、,で区切った一区切り）に「扱う」の語があれば、その節の仕事は登録済みの Sub のもの
    # （「合計の行を足して」の「行を足」を、合計を扱う Sub と別の仕事に数えない・2026-09-17）
    clauses = [c for c in re.split(r'[。、,，\n]', req) if c]

    asks = [e['ask'] for e in extras]

    def _covered(w):
        if any(h in w or w in h for h in covered):
            return True
        if _STRONG_OTHER_RE.fullmatch(w):
            # グラフ・並べ替え・印刷・色付け・抽出は、同じ節に「扱う」の語があっても別の操作（2026-09-17 深夜:
            # 「振り直して、区分ごとの円グラフも」の「区分」でグラフを振り直しの仕事と見なし、グラフを AI に回さず終えていた）
            return False
        # 同じ節に「扱う」の語か、撃つマクロの「依頼の語」があれば、その節の仕事はそのマクロのもの
        # （2026-09-17 夕: 突合の依頼の「件数」「照合」「突き合わせ」が AI に回る仕事に数えられ、突合を撃った後も AI へ回った）
        return any(w in c and (any(h in c for h in covered) or any(_ask_strength(a, c) for a in asks)) for c in clauses)
    left = [w for w in others if not _covered(w)]
    left += [w for e in shadowed for w in (e.get('handles') or [])[:1] if w not in left]   # 撃たなかった仕事は AI に
    return {'fire': bool(tidy or extras), 'tidy': tidy, 'dedupe': dedupe, 'extras': extras,
            'left': left, 'other': bool(left), 'skipped': skipped, 'shadowed': shadowed}


def _ask_strength(ask, req):
    """依頼の語（語を | で並べたもの）が依頼文に当たった一番長い語の文字数（純 Python）。当たらなければ 0。
    正規表現の記号の入った古い登録（.* など）は、当たった部分でなく語の文字（記号を除く）の長さで数える。"""
    best = 0
    for alt in str(ask or '').split('|'):
        if not alt:
            continue
        try:
            if re.search(alt, req):
                best = max(best, len(re.sub(r'[.*+?()\[\]{}^$\\]', '', alt)))
        except re.error:
            continue
    return best


def _ask_best_word(ask, req):
    """依頼の語のうち、依頼文に当たったいちばん長い語（純 Python）。無ければ ''。"""
    best = ''
    for alt in str(ask or '').split('|'):
        try:
            m = re.search(alt, req) if alt else None
        except re.error:
            continue
        if m and len(m.group(0)) > len(best):
            best = m.group(0)
    return best


def plan_of(request, registry=()):
    """依頼文 → (先撃ちするか, 重複行を消すも撃つか, マクロの扱わない仕事があるか)（純 Python・登録簿なしなら 9/11 の形）。"""
    p = plan_full(request, registry)
    if not p['tidy'] and not p['extras']:
        return False, False, False
    return True, p['dedupe'], p['other']


def _macro_owner(xl):
    """「表を整える」を持つブックの名前（モジュール「表の整理」を先に探す＝全モジュールを読まない）。無ければ None。"""
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return None
    from vbam_vba import _project_book_name
    for p in projects:
        try:
            cm = p.VBComponents(_MACRO_MODULE).CodeModule
            n = int(cm.CountOfLines)
            if n and re.search(r'^\s*Sub\s+' + _MACRO_TIDY + r'\b', cm.Lines(1, n), re.M):
                name = _owner_name(xl, p)
                if name:
                    return name
        except Exception:
            continue
    return None


def _owner_name(xl, p):
    """プロジェクト → Run で名指しできるブック名。名前が指すブックに本当に「表の整理」があるかまで確かめる
    （2026-09-17: 未保存の Book4 が開いていると _project_book_name が別プロジェクトの名前を Book4 と取り違え、
    'Book4'!表を整える を撃って「マクロを実行できません」で先撃ちが止まった）。"""
    from vbam_vba import _project_book_name
    try:
        cm = p.VBComponents(_MACRO_MODULE).CodeModule
        n = int(cm.CountOfLines)
        sig = (n, cm.Lines(1, min(n, 5)) if n else '')
    except Exception:
        return None

    def _same(name):
        try:
            c2 = xl.Workbooks(name).VBProject.VBComponents(_MACRO_MODULE).CodeModule
            n2 = int(c2.CountOfLines)
            return (n2, c2.Lines(1, min(n2, 5)) if n2 else '') == sig
        except Exception:
            return False
    # ThisWorkbook の Name プロパティ＝そのプロジェクトのブック名（アドイン・非表示・未保存でも取れる）
    try:
        name = str(p.VBComponents('ThisWorkbook').Properties('Name').Value)
        if name and _same(name):
            return name
    except Exception:
        pass
    name = _project_book_name(xl, p)
    if name and _same(name):
        return name
    # Filename が取れない（アドイン・未保存）と _project_book_name は名前を取り違える。開いているブックと
    # 読み込まれているアドインを、同じモジュールの中身（行数と頭 5 行）で照らして探す
    cands = []
    try:
        cands += [str(w.Name) for w in xl.Workbooks]
    except Exception:
        pass
    try:
        cands += [str(a.Name) for a in xl.AddIns if a.Installed]
    except Exception:
        pass
    for c in cands:
        if _same(c):
            return c
    return None


def _table_anchor(ws):
    """表の見出しの左端のセル（tidy と検査の当て先）。読めなければ None。"""
    from vbam_hands import _rows_of, _guess_header_row
    from vbam_view import _guess_header_idx
    try:
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        rows = [list(r) for r in _rows_of(ur.Value)]
    except Exception:
        return None
    if not rows:
        return None
    # 材料の「見出し行の推定」と同じ関数で決める（2026-09-17: 表題が見出しの真上にある表で、格子の推定が
    # 表題の行を見出しに取り、tidy と検査が 1 行ずれて「見出しが空」「SUM が本文 1 行を入れていない」と出た）
    h = None
    try:
        got = _guess_header_row(ws)
        if got and r0 <= int(got[0]) < r0 + len(rows):
            h = int(got[0]) - r0
    except Exception:
        h = None
    if h is None:
        h = _guess_header_idx(rows)
    h = 0 if h is None or h < 0 else h
    j = next((k for k, v in enumerate(rows[h]) if v not in (None, '')), 0)
    return f"{_col_letter(c0 + j)}{r0 + h}"


def _changed_cells(before, now, r0, c0):
    """撃つ前の控え（{'row','col','values'}）と今の値（左上 r0,c0 の 2 次元）→ 値が入って変わったセル [(行, 列)]（純 Python）。
    消えたセル（前の結果の残りを消した所）は数えない＝表の当て先にならない。"""
    from vbam_hands import _text_of
    bv = (before or {}).get('values') or []
    br, bc = int((before or {}).get('row') or 1), int((before or {}).get('col') or 1)
    out = []
    for i, row in enumerate(now or ()):
        for j, v in enumerate(row or ()):
            t = _text_of(v)
            if t == '':
                continue
            r, c = r0 + i, c0 + j
            bi, bj = r - br, c - bc
            old = _text_of(bv[bi][bj]) if 0 <= bi < len(bv) and 0 <= bj < len(bv[bi] or ()) else ''
            if t != old:
                out.append((r, c))
    return out


def _written_anchors(ws, before, limit=6):
    """登録簿のマクロが値を書いた表ごとの見出しの左端（COM）。→ [番地]。読めなければ []。
    2026-09-17 夜: 帳票の右に一覧・明細の右に集計表を作るマクロで、tidy と検査が元の表（帳票・明細）にだけ当たり、
    作った表には書式が付かず、元の表の汚れ（明細の空白つき金額・帳票の桁区切り）を「残りの指摘」にして AI に回していた。"""
    from vbam_hands import _rows_of
    from vbam_edit import _region_of_cell
    if not before or before.get('too_big') or 'values' not in before:
        return []
    try:
        ur = ws.UsedRange
        cells = _changed_cells(before, [list(r) for r in _rows_of(ur.Value)], int(ur.Row), int(ur.Column))
    except Exception:
        return []
    regions, anchors = [], []
    for r, c in cells:
        if any(a <= r <= b and x <= c <= y for a, x, b, y in regions):
            continue
        try:
            rg = _region_of_cell(ws, ws.Cells(r, c))
            a, x = int(rg.Row), int(rg.Column)
            b, y = a + int(rg.Rows.Count) - 1, x + int(rg.Columns.Count) - 1
        except Exception:
            continue
        regions.append((min(a, r), min(x, c), max(b, r), max(y, c)))
        if _continues_above(ws, regions[-1], regions[:-1]):
            continue                                     # 上の表の続き（空行で切れた下の塊）＝別の表として仕上げ・検査しない
        if f"{_col_letter(x)}{a}" not in anchors:
            anchors.append(f"{_col_letter(x)}{a}")
        if len(anchors) >= limit:
            break
    return anchors


def _continues_above(ws, region, others, gap=3):
    """region が、上にある表（others のどれか）の続きか（COM）: 左端の列が同じで、間が空行 gap 行以内で、
    region の 1 行目に数（＝見出しでなくデータの行）がある。
    2026-09-17 夜: 空行 2 行で上下に分かれた旧システムの表で、下の塊の 1 行目（データ）を見出しとして検査し
    「見出しが空の列（D）」を AI に回した（仕上げもデータの行に見出しの書式を当てる形だった）。"""
    from vbam_hands import _rows_of
    a, x, _b, y = region
    for a1, x1, b1, _y1 in others:
        if x1 != x or not (b1 < a <= b1 + 1 + gap):
            continue
        try:
            row = (_rows_of(ws.Range(ws.Cells(a, x), ws.Cells(a, y)).Value) or [[]])[0]
        except Exception:
            return False
        if any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in row):
            return True
    return False


def _rehearse(xl, wb, sheet, plan, owner, sel_cols=None):
    """登録簿のマクロを写しで先に撃つ。→ (撃たない登録簿の行, 理由)。全部よければ ([], '')。
    撃つ順（表を整える → 重複行を消す → 登録簿）は本番と同じ。写しは wb.SaveCopyAs（未保存の変更も入る・人のブックの
    保存状態とパスは変わらない）。"""
    import os
    import tempfile
    from vbam_forge import rehearse_steps
    extras = list(plan['extras'])
    by_owner = {}
    for e in extras:
        by_owner.setdefault(e.get('owner') or owner, []).append(e)
    if len(by_owner) != 1:
        return extras, "持ち主のブックが複数で試せませんでした"
    own = next(iter(by_owner))
    try:
        cm = xl.Workbooks(own).VBProject.VBComponents(_MACRO_MODULE).CodeModule
        text = cm.Lines(1, int(cm.CountOfLines))
    except Exception as ex:
        return extras, f"マクロの本文を読めませんでした（{ex}）"
    names = ([_MACRO_TIDY] + ([_MACRO_DEDUPE] if plan['dedupe'] else []) if plan['tidy'] and owner == own else [])
    names += [e['name'] for e in extras]
    ext = os.path.splitext(str(wb.Name))[1] or '.xlsx'
    work = tempfile.mkdtemp(prefix='_prefire_copy_')
    copy = os.path.join(work, '写し' + ext)
    try:
        wb.SaveCopyAs(copy)
    except Exception as ex:
        return extras, f"写しを作れませんでした（{ex}）"
    try:
        steps = rehearse_steps(text, names, copy, sheet, _MACRO_MODULE, sel_cols or {})
    except Exception as ex:
        return extras, f"試し撃ちが止まりました（{ex}）"
    finally:
        import shutil
        shutil.rmtree(work, ignore_errors=True)
    by_name = {s['name']: s for s in steps}
    for nm in names:
        s = by_name.get(nm)
        if s is None:
            return extras, "前のマクロで止まったので試せませんでした"
        if not s['ok']:
            return extras, f"{nm} が{s['why']}"
        if nm in [e['name'] for e in extras] and not s['changed']:
            return extras, f"{nm} が表を何も変えずに終わりました（この表の形はそのマクロの想定外）"
    return [], ''


def _output_sheets(wb, sheet):
    """マクロが作る先のシート＝その回の持ち物（純 COM）。アクティブなシートを元にしたピボットのあるシートと、
    クエリで読み込んだテーブルのあるシート。前の結果が残っていて作り直したときに、それを
    「頼んでいない変化」と呼ばないため（2026-09-18 朝: 値は正解と一致したのに、前の結果のあった
    4 本（月別ピボット・行列ピボット・ダッシュボード・PQ 追加集計）が入口から AI に回っていた）。"""
    out = set()
    try:
        shs = list(wb.Worksheets)
    except Exception:
        return out
    for sh in shs:
        try:
            nm = str(sh.Name)
        except Exception:
            continue
        if nm == sheet:
            continue
        try:
            pts = sh.PivotTables()
            for i in range(1, int(pts.Count) + 1):
                try:
                    if sheet in str(pts.Item(i).SourceData or ''):
                        out.add(nm)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            los = sh.ListObjects
            for i in range(1, int(los.Count) + 1):
                try:
                    if los.Item(i).QueryTable is not None:
                        out.add(nm)
                except Exception:
                    pass
        except Exception:
            pass
    return out


def macro_first(request, sheet, wb, max_turns, run_id):
    """先撃ち 1 回。戻り値:
      None                           … 先撃ちしない（依頼が整理でない・マクロが無い）
      {'result': r}                  … マクロで合格（AI を呼ばずに終わる。r は run_agent と同じ形）
      {'result': None, 'ran': True…} … 撃ったが残りがある／ほかの仕事がある＝Sonnet の往復へ（控えを引き継ぐ）
    """
    xl = wb.Application
    registry = registry_of(xl)                       # 鍛えたマクロの登録簿（.bas の頭 2 行・2026-09-17）
    words = None
    if any(e.get('headers') for e in registry):
        try:
            words = _top_words(wb.Sheets(sheet))
        except Exception as ex:
            print(f"（シートの見出しを読めませんでした: {ex}。見出しの要る登録簿のマクロは撃ちません）")
            words = set()                            # 読めなければ、見出しの要る Sub は撃たない（安全側）
    shape = None
    if any(e.get('shape') for e in registry):
        shape = _shape_of_sheet(wb.Sheets(sheet), wb)         # 形の要る仕事は、形がそろうときだけ候補（2026-09-18 朝）
    plan = plan_full(request, registry, words, shape)
    # 列を選ぶ必要のある Sub（' 選ぶ列: N）は、依頼の語と見出しが一致する列を道具が選んでから撃つ（2026-09-18）
    sel_cols = {}
    if any(e.get('select') for e in plan['extras']):
        cells = _header_cells(wb.Sheets(sheet))
        have = current_columns(wb.Sheets(sheet))         # 人が列を選んでいれば、それを使う
        drop = []
        for e in [x for x in plan['extras'] if x.get('select')]:
            lo, hi = e['select']
            if lo <= len(have) <= hi:
                sel_cols[e['name']] = have
                print(f"（{e['name']} は人が選んでいる {len(have)} 列を使います）")
                continue
            cols = pick_columns(request, cells, e['select'])
            if not cols:
                drop.append(e['name'])
                print(f"（登録簿の {e['name']} は列を選ぶ仕事ですが、依頼の語と見出しが一致する列が "
                      f"{e['select'][0]}〜{e['select'][1]} 列になりません＝撃たずに AI に回します）")
            else:
                sel_cols[e['name']] = cols
        if drop:
            registry = [e for e in registry if e['name'] not in set(drop)]
            plan = plan_full(request, registry, words, shape)
    for e in plan.get('skipped') or []:
        if e.get('why_shape'):
            print(f"（登録簿の {e['name']} は依頼の語に当たりましたが、この表に「{'・'.join(e['why_shape'])}」が無いので撃ちません）")
        else:
            print(f"（登録簿の {e['name']} は依頼の語に当たりましたが、シートに見出し「{'・'.join(e['headers'])}」が"
                  "そろっていないので撃ちません）")
    for e in plan.get('shadowed') or []:
        print(f"（登録簿の {e['name']} も依頼の語に当たりましたが、依頼はより長い語で別の仕事を名指ししているので撃ちません）")
    if not plan['fire']:
        return None
    dedupe, other = plan['dedupe'], plan['other']
    owner = _macro_owner(xl) if plan['tidy'] else None
    if plan['tidy'] and not owner:
        # 整理の語はあるが「表を整える」が無い＝登録簿の Sub だけで撃てるならそれだけ、無ければ AI へ
        if not plan['extras']:
            return None
    if not plan['tidy'] and not plan['extras']:
        return None
    if not owner and not any(e.get('owner') for e in plan['extras']):
        return None                                   # 撃てる持ち主が 1 つも無い
    t_start = time.time()
    tool_sec = {}
    if plan['extras']:
        # 鍛えて登録したマクロは、人の Excel で撃つ前に写しで試す（実行時エラーの窓で固まらない・何もしないで抜けたら AI へ）
        t0 = time.time()
        dropped, why = _rehearse(xl, wb, sheet, plan, owner, sel_cols)
        tool_sec['試し撃ち'] = time.time() - t0
        if dropped:
            print(f"（登録簿の {'・'.join(e['name'] for e in dropped)} は写しで試すと{why}。人のブックでは撃たず、"
                  "その仕事は AI に回します）")
            gone = {x['name'] for x in dropped}
            plan = plan_full(request, [e for e in registry if e['name'] not in gone], words, shape)
            dedupe, other = plan['dedupe'], plan['other']
            if not plan['tidy'] and not plan['extras']:
                return None
            if plan['tidy'] and not owner:
                owner = _macro_owner(xl)
                if not owner:
                    return None
    ws = wb.Sheets(sheet)
    book = str(wb.Name)
    # 控え（本番の往復と同じ: シートの控え・ブックの控え＝--undo・頼んでいない変化を見るための控え）
    t0 = time.time()
    before = va._sheet_snapshot(ws) or {}
    va._agent_backup(wb, ws, sheet, request, run_id)
    book_before = None
    if va._book_cell_count(wb) <= va._INV_BOOK_MAX_CELLS:
        book_before = va._snapshot_book(wb, own={sheet})      # 触らないシートは軽い控え（2026-09-16）
    tool_sec['控え'] = time.time() - t0
    # 撃つ（マクロはアクティブシートに効く＝対象シートを前に出してから）
    t0 = time.time()
    wb.Activate()
    ws.Activate()
    shots = []                                        # [(持ち主のブック, Sub 名)] 撃つ順
    if owner:
        shots.append((owner, _MACRO_TIDY))
        if dedupe:
            shots.append((owner, _MACRO_DEDUPE))
    shots += [(e.get('owner') or owner, e['name']) for e in plan['extras'] if (e.get('owner') or owner)]
    names = [nm for _o, nm in shots]
    for o, nm in shots:
        if sel_cols.get(nm):
            select_columns(ws, sel_cols[nm])         # 列を選ぶ仕事は、選んでから撃つ（マクロは Selection を見る）
        xl.Run(f"'{o.replace(chr(39), chr(39) * 2)}'!{nm}")
    tool_sec['マクロ'] = time.time() - t0
    owner = owner or shots[0][0]
    owners_txt = '・'.join(sorted({o for o, _n in shots}))
    print(f"マクロの先撃ち: {' → '.join(names)}（{owners_txt}・{tool_sec['マクロ']:.2f} 秒・AI は呼んでいません）")
    with contextlib.suppress(Exception):
        va.fired_append(run_id, book, sheet, request, names, [o for o, _n in shots])   # 使用の記録（引退の判断に使う・2026-09-18 夜）
    # 道具の仕上げ（tidy）と終わりの検査
    anchor = _table_anchor(ws)
    targets = [anchor] if anchor else []
    if plan['extras']:
        # 登録簿のマクロが書いた表に当てる（元の表は「整えて」と頼まれたときだけ）。書いた所が読めなければ元の表
        wrote = _written_anchors(ws, before)
        if wrote:
            targets = ([anchor] if (anchor and plan['tidy']) else []) + [a for a in wrote if a != anchor or not plan['tidy']]
        elif not plan['tidy']:
            # セルに何も書かないマクロ（グラフを作る・別のシートに出す）は、元の表を仕上げも検査もしない
            # ＝元の表にもとからある指摘（見出しの空の列・罫線の継ぎはぎ）で AI に回さない（2026-09-18 朝）
            targets = []
    # 絞り込み（行を隠す）が仕事のマクロの後は仕上げを当てない＝tidy は書式が隠れた行を飛ばさないよう
    # 絞り込みを解くので、マクロの結果（隠れた行）が消える（2026-09-18 朝: 職員名簿の絞り込みが
    # 「一発合格」と出ながら隠れた行 0 だった）
    if targets and any(re.search(r'絞り込|フィルタ|隠す|抽出', (e.get('ask') or '') + ' ' + ' '.join(e.get('handles') or []))
                       for e in plan['extras']):
        print("仕上げ（tidy）は当てません: 絞り込みを扱うマクロの後は、書式を当てると絞り込みが解けます")
        targets = []
    t0 = time.time()
    if targets:
        va._run_cmd(['tidy', '--sheet', sheet] + targets, wb)
    tool_sec['仕上げ'] = time.time() - t0
    t0 = time.time()
    audit = va._audit_now(wb, sheet, targets) if targets else []
    c_block, c_noticed = va._content_now(wb, sheet, targets) if targets else ([], [])
    r_block, r_noticed = va._request_gate(wb, sheet, book_before, request, {'normalize': []})
    if plan['extras'] and not plan['tidy'] and r_block:
        # 依頼の中身の検査はシートの主な表（明細・帳票＝元の表）を見る。登録簿のマクロだけの依頼は元の表を触らない決まりで、
        # 書いた表は鍛えたときに正解と突き合わせ済み＝元の表の汚れは AI に回さず気づきに出す
        # （2026-09-17 夜: 月次集計の依頼の「空白」「支出額」で、明細の空白つき・文字の金額を指摘して AI に回していた）
        r_noticed = list(r_noticed) + [f"元の表: {b}" for b in r_block]
        r_block = []
    if plan['extras'] and not plan['tidy'] and (audit or c_block):
        # 登録簿のマクロだけの依頼で、マクロが変えていないセルの指摘（元の表の合計行が文字の金額と合わない・元の式）は
        # AI に回さず気づきに出す（2026-09-17 夜: 値は正解と一致したのに、元の表の「円」付き金額の合計行を理由に AI へ回した）
        audit, c_block, moved = _split_old_findings(audit, c_block, before, ws, sheet)
        c_noticed = list(c_noticed) + [f"元の表: {b}" for b in moved]
    skip = ('formula', 'hidden', 'tables') if dedupe else ()
    # マクロが作り直した出力シート（ピボット・クエリの読み込み先）は、控えから外して比べない＝
    # 「同じ名前のシートがあれば消して作り直す」が仕事なので、前の姿と違って当たり前
    # （2026-09-18 朝: 前の結果が残っている表で「結合セルが全部外れた」と出て AI に回っていた）
    bb = book_before
    outs = _output_sheets(wb, sheet) if book_before is not None else set()
    if outs and isinstance(bb, dict) and bb.get('sheets'):
        bb = dict(bb, sheets={n: s for n, s in bb['sheets'].items() if n not in outs},
                  order=[n for n in (bb.get('order') or []) if n not in outs])
        print("作り直した出力シートは控えと比べません: " + "・".join(sorted(outs)))
    inv = va._inv_violations(wb, bb, {sheet}, skip) if bb is not None else []
    tool_sec['検査'] = time.time() - t0
    blocks = list(audit) + list(c_block) + list(r_block)
    noticed = list(c_noticed) + list(r_noticed)
    did = [f"マクロ {nm}（{owner}）" for nm in names] + ([f"tidy {' '.join(targets)}"] if targets else [])
    if blocks or inv or other:
        why = []
        if blocks or inv:
            why.append(f"残りの指摘 {len(blocks) + len(inv)} 件")
        if other:
            why.append("依頼にマクロの扱わない仕事がある")
        print(f"先撃ちの後: {'・'.join(why)}＝Sonnet の往復へ回します（控えは先撃ちの前のものを引き継ぎます）")
        note = ("\n--- 道具が先に撃ったマクロ（AI より先に・" + f"{time.time() - t_start:.1f} 秒） ---\n"
                + f"{' → '.join(names)} を実行し、tidy で仕上げました。上の材料は撃った後の姿です"
                + "（書き方の整理" + ("・重複の削除" if dedupe else "") + "は済んでいます。やり直さない）。\n")
        if blocks or inv:
            note += "残っている指摘（ここと、依頼のうちマクロでは扱わない仕事だけを手で直してください）:\n" + "".join(
                f"  - {b}\n" for b in (blocks + list(inv))[:12])
        else:
            note += "書き方の指摘は残っていません。依頼のうちマクロでは扱わない仕事だけを手で進めてください。\n"
        return {'result': None, 'ran': True, 'before': before, 'book_before': book_before, 'inv_skip': skip,
                'did': did, 'touched': targets, 'note': note, 'tool_sec': tool_sec}
    # 合格（AI を呼ばずに終わる）
    changed_rows, fmt_rows = [], []
    try:
        changed_rows = va._changes_of(before, ws, sheet) or []
        fmt_rows = va._format_changes(before.get('format'), ws) or []
    except Exception:
        pass
    path = va._write_changes_file(changed_rows, fmt_rows=fmt_rows) if (changed_rows or fmt_rows) else None
    print(f"変わったセル: {len(changed_rows)} 個" + (f"   明細: {path}" if path else ""))
    for line in va._changes_table(changed_rows)[:14]:
        print(line)
    print("仕上げ検査: 指摘なし（見出し・罫線・列の型・列幅・エラー値・合計・式・依頼の語の列の残り）")
    print("頼んでいない変化: なし")
    report = ("【やったこと】（道具が実行した手から）\n" + "\n".join("・" + d for d in did)
              + "\n\n【できなかったこと】なし")
    if noticed:
        report += "\n\n【気づいたこと】（依頼には無いので触っていません）\n" + "\n".join("・" + n for n in noticed)
    sec = time.time() - t_start
    print("関所の差し戻し: なし（マクロで一発）　＝ 一発合格")
    print("道具側の内訳: " + " / ".join(f"{k} {v:.1f}" for k, v in tool_sec.items())
          + f" 秒（合計 {sum(tool_sec.values()):.1f}・AI 待ち 0.0）")
    print("（保存はしていません。Excelで確認後に保存してください。気に入らなければ保存せずに閉じる／agent --undo で戻す）")
    gates = dict.fromkeys(va._GATE_LABELS, 0)
    r = {'done': True, 'ok': True, 'verify': True, 'turns': 0, 'report': report, 'plan': [], 'dropped': [],
         'usage': {'in': 0, 'out': 0, 'think': 0, 'sent': 0, 'recv': 0, 'sec': 0.0}, 'gates': gates,
         'run_id': run_id, 'audit': [], 'unmet': [], 'graded': False, 'grade_skipped': True, 'inv': [],
         'asks': [], 'formula_rate': None, 'limit': max_turns, 'discovered': [], 'merged': 0,
         'tool_sec': {k: round(v, 2) for k, v in tool_sec.items()}, 'inv_skip': skip, 'inv_own': (sheet,),
         'in_tokens': 0, 'out_tokens': 0, 'macro_first': names}
    try:
        va._after_save(book, sheet, wb, True)
    except Exception:
        pass
    va._runs_record(run_id, 'sheet', book, sheet, request, r, sec, path=va._book_path(wb))
    va.progress_write('done', book=book, sheet=sheet, turn=0, max_turns=max_turns, request=str(request)[:80],
                      note='マクロで合格', sec=sec)
    print("AI の報告:\n" + report)
    return {'result': r}


__all__ = ['plan_of', 'plan_full', 'macro_first', 'registry_of', 'registry_from_text', 'dedupe_registry']
