# -*- coding: utf-8 -*-
"""vbam_cause.py — 式の答えがおかしい・エラーになる「原因」を、式と値から道具が決める（2026-10-02）。

オフィス田中「ExcelのAI活用」で Copilot が外したのは、ほとんどが理由の当て推量と、間違った直し方だった
（数値と文字の番号・文字の日付・MATCH や RANK の引数の順・FALSE 忘れの VLOOKUP・文字の数字の SUM・末尾の空白…）。
理由の多くは式の形と値の型から機械的に決まる。道具が決めて渡せば、AI は直すだけで済む。

純 Python の本体は cause_notes(fa, fv, r0, c0, names, udfs, nv) と spelling_notes(fv, r0, c0)。
fa＝式の格子（Range.Formula）、fv＝値の格子（Range.Value）、r0/c0＝格子の左上の行・列。
names＝{名前（小文字）: 参照先の字面}、udfs＝{VBA の Function 名（大文字）}。nv=True は値を言わない（値なしの切り替え）。
"""
import re

from vbam_lookup import _mask_strings, _split_args, approx_calls

XL_ERR = {-2146826288: '#NULL!', -2146826281: '#DIV/0!', -2146826273: '#VALUE!', -2146826265: '#REF!',
          -2146826259: '#NAME?', -2146826252: '#NUM!', -2146826246: '#N/A', -2146826243: '#SPILL!',
          -2146826238: '#CALC!'}
_REF_RE = re.compile(r"(?<![A-Za-z0-9_.!'\]])\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?(?![A-Za-z0-9_(!])")
_NAME_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9_.!\]\[$'])([A-Za-z_぀-ヿ一-鿿][A-Za-z0-9_.぀-ヿ一-鿿]*)(?![A-Za-z0-9_.぀-ヿ一-鿿]*\s*[(!])")
_FUNC_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:_xlfn\.|_xludf\.)?[A-Za-z_぀-ヿ一-鿿][A-Za-z0-9_.぀-ヿ一-鿿]*)\s*\(")
_NUMTEXT_RE = re.compile(r'^[\s　]*[-+]?[0-9０-９][0-9０-９,，]*(\.[0-9]+)?[\s　]*$')
_DATETEXT_RE = re.compile(r'^\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}')


def _col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def _col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _err(v):
    return XL_ERR.get(v) if isinstance(v, int) and not isinstance(v, bool) else None


def _parse_ref(text):
    """同じシートの番地（A2・$D$2:$E$4）→ (上, 左, 下, 右)。シート付き・名前・式なら None。"""
    t = (text or '').strip()
    if '!' in t:
        return None
    m = _REF_RE.fullmatch(t)
    if not m:
        return None
    c1, r1 = _col_num(m.group(1)), int(m.group(2))
    c2 = _col_num(m.group(3)) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _addr(r, c):
    return f"{_col(c)}{r}"


def _box_addr(b):
    return _addr(b[0], b[1]) if (b[0], b[1]) == (b[2], b[3]) else f"{_addr(b[0], b[1])}:{_addr(b[2], b[3])}"


class _G:
    """値の格子を番地で読む。"""

    def __init__(self, fv, r0, c0):
        self.fv, self.r0, self.c0 = fv, r0, c0

    def get(self, r, c):
        i, j = r - self.r0, c - self.c0
        if 0 <= i < len(self.fv) and 0 <= j < len(self.fv[i]):
            return self.fv[i][j]
        return None

    def cells(self, b):
        for r in range(b[0], b[2] + 1):
            for c in range(b[1], b[3] + 1):
                yield r, c, self.get(r, c)


def _calls(formula, fnames):
    """式の中の fnames の呼び出し → [(関数名, [引数の字面])]（文字列の中は見ない）。"""
    s = formula or ''
    masked = _mask_strings(s)
    rx = re.compile(r'(?<![A-Za-z0-9_.])(?:_xlfn\.)?(' + '|'.join(fnames) + r')\s*\(', re.IGNORECASE)
    out = []
    for m in rx.finditer(masked):
        args = _split_args(s, masked, m.end() - 1) or []
        out.append((m.group(1).upper(), [a[0].strip() for a in args]))
    return out


def _is_text_number(v):
    return isinstance(v, str) and bool(_NUMTEXT_RE.match(v)) and v.strip() != ''


def _num_eq_text(n, s):
    try:
        return float(str(s).replace(',', '').replace('，', '').strip().translate(str.maketrans('０１２３４５６７８９', '0123456789'))) == float(n)
    except (TypeError, ValueError):
        return False


def _q(v, nv):
    """値を言う（値なしなら伏せる）。"""
    if nv:
        return ''
    s = str(v)
    if hasattr(v, 'year'):
        s = f"{v.year}/{v.month}/{v.day}"
    elif isinstance(v, float) and v.is_integer():
        s = str(int(v))
    return f"「{s[:20]}」"


def _lookup_cause(fn, args, g, nv):
    """完全一致の検索が #N/A のとき、探す値と表の 1 列目の型を比べて理由を言う。"""
    if fn in ('VLOOKUP', 'HLOOKUP'):
        if len(args) < 2:
            return None
        key_t, tab = args[0], _parse_ref(args[1])
        if not tab:
            return None
        col = (tab[0], tab[1], tab[2], tab[1]) if fn == 'VLOOKUP' else (tab[0], tab[1], tab[0], tab[3])
    elif fn in ('XLOOKUP', 'MATCH'):
        if len(args) < 2:
            return None
        key_t, col = args[0], _parse_ref(args[1])
        if not col:
            return None
    else:
        return None
    vals = [v for _r, _c, v in g.cells(col) if v not in (None, '')]
    kb = _parse_ref(key_t)
    textish = key_t.upper().startswith('TEXT(') or (key_t.startswith('"') and key_t.endswith('"'))
    if kb and (kb[0], kb[1]) == (kb[2], kb[3]):
        kv = g.get(kb[0], kb[1])
        where = _addr(kb[0], kb[1])
        if isinstance(kv, (int, float)) and not isinstance(kv, bool) and _err(kv) is None:
            if any(isinstance(v, str) and _num_eq_text(kv, v) for v in vals):
                return (f"探す値 {where}{_q(kv, nv)} は数値、表の 1 列目（{_box_addr(col)}）は文字の数字＝型が違うので一致しない"
                        "（どちらかにそろえる。表を数に直すなら棚「選んだ列の文字の数字を数にする」）")
        if isinstance(kv, str):
            if any(isinstance(v, (int, float)) and not isinstance(v, bool) and _num_eq_text(v, kv) for v in vals):
                return f"探す値 {where}{_q(kv, nv)} は文字の数字、表の 1 列目（{_box_addr(col)}）は数値＝型が違うので一致しない"
            if kv != kv.strip(' 　') and any(isinstance(v, str) and v == kv.strip(' 　') for v in vals):
                return f"探す値 {where}{_q(kv, nv)} の前後に空白＝表の値と一致しない（空白を消すか TRIM で探す）"
            if _DATETEXT_RE.match(kv.strip()) and any(hasattr(v, 'year') for v in vals):
                return (f"探す値 {where}{_q(kv, nv)} は文字の日付、表の 1 列目（{_box_addr(col)}）は日付（シリアル値）＝型が違う"
                        "（DATEVALUE で日付にするか、探す値も日付で持つ）")
        if kv in (None, ''):
            return f"探す値 {where} が空"
        return f"探す値 {where}{_q(kv, nv)} が表の 1 列目（{_box_addr(col)}）に無い"
    if textish and any(hasattr(v, 'year') for v in vals):
        return (f"探す値を TEXT で文字にしているが、表の 1 列目（{_box_addr(col)}）は日付（シリアル値）＝文字と日付は一致しない"
                "（TEXT を外して日付どうしで探す）")
    return None


def cause_notes(fa, fv, r0, c0, names=None, udfs=None, nv=False, limit=14):
    """式と値から原因を決める → 行のリスト（純 Python）。"""
    g = _G(fv, r0, c0)
    names = {k.lower(): v for k, v in (names or {}).items()}
    udfs = {u.upper() for u in (udfs or set())}
    found = {}                          # 文 → [番地]
    explain = {}                        # 名前・自作関数の正体（最初の番地だけ）

    def add(msg, a):
        found.setdefault(msg, []).append(a)

    for i, row in enumerate(fa):
        for j, f in enumerate(row):
            if not (isinstance(f, str) and f.startswith('=')):
                continue
            r, c = r0 + i, c0 + j
            a = _addr(r, c)
            v = fv[i][j] if i < len(fv) and j < len(fv[i]) else None
            err = _err(v)
            masked = _mask_strings(f)
            # 名前と自作関数の正体（田中さん: TEST(CHECK(A1)) の TEST は名前、CHECK は VBA。説明に「名前」の語が出なかった）
            for m in _FUNC_TOKEN_RE.finditer(masked):
                tok = re.sub(r'^(_xlfn\.|_xludf\.)', '', m.group(1))
                if tok.lower() in names and tok not in explain:
                    explain[tok] = f"{tok} は名前（{a} で使用・中身 {names[tok.lower()][:60]}）＝［数式］→［名前の管理］で見る"
                elif tok.upper() in udfs and tok not in explain:
                    explain[tok] = f"{tok} は VBA の自作関数（Function {tok}・{a} で使用）＝Alt+F11 で中身を見る"
            # 1. 消えた参照
            if '#REF!' in masked:
                add("式の中の参照が #REF! になっている＝参照していた行・列・シートを消した（消す前の範囲を指し直す）", a)
                continue
            # 2. 定義されていない名前（#NAME? のときだけ）
            if err == '#NAME?' and names is not None:
                body = re.sub(r"(?:'[^']*'|[A-Za-z0-9_぀-鿿]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?", ' ', masked)
                body = _REF_RE.sub(' ', body)
                miss = [t for t in _NAME_TOKEN_RE.findall(body)
                        if t.upper() not in ('TRUE', 'FALSE') and t.lower() not in names and not re.fullmatch(r'[A-Z]{1,3}', t)]
                if miss:
                    add(f"名前「{miss[0]}」が定義されていない＝名前を消したか綴り違い（［数式］→［名前の管理］）", a)
                    continue
            # 3. 引数の順（MATCH・RANK）
            for fn, args in _calls(f, ['MATCH', 'RANK', 'RANK\\.EQ', 'RANK\\.AVG']):
                if len(args) >= 2:
                    b0, b1 = _parse_ref(args[0]), _parse_ref(args[1])
                    if b0 and b1 and (b0[0], b0[1]) != (b0[2], b0[3]) and (b1[0], b1[1]) == (b1[2], b1[3]):
                        what = "探す値、第 2 引数が探す範囲" if fn == 'MATCH' else "順位を調べる数値、第 2 引数が範囲"
                        add(f"{fn} の引数の順が逆（第 1 引数は{what}）", a)
            # 4. MAP と LAMBDA の引数の数
            for fn, args in _calls(f, ['MAP']):
                if args and args[-1].upper().startswith('LAMBDA('):
                    lam = _calls(args[-1], ['LAMBDA'])
                    if lam:
                        params = len(lam[0][1]) - 1
                        arrays = len(args) - 1
                        if params != arrays:
                            add(f"MAP に渡した配列は {arrays} つ、LAMBDA の引数は {params} つ＝数をそろえる"
                                "（配列を足すか、LAMBDA の引数を減らす）", a)
            # 5. 検索の #N/A（完全一致）の理由
            if err == '#N/A' and not approx_calls(f):      # 近似一致の #N/A は 8 で言う（表にあっても見つからない）
                for fn, args in _calls(f, ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'MATCH']):
                    why = _lookup_cause(fn, args, g, nv)
                    if why:
                        add(f"{fn} が #N/A: {why}", a)
                        break
            # 6. 集計の範囲の中のエラーが伝わっている
            if err:
                for fn, args in _calls(f, ['SUM', 'MIN', 'MAX', 'AVERAGE', 'SUMPRODUCT']):
                    for t in args:
                        b = _parse_ref(t)
                        if b:
                            bad = [(rr, cc, _err(vv)) for rr, cc, vv in g.cells(b) if _err(vv)]
                            if bad:
                                rr, cc, e = bad[0]
                                add(f"{fn} の範囲 {_box_addr(b)} の中の {_addr(rr, cc)}（{e}）がそのまま伝わっている"
                                    "＝元のセルを直すか、AGGREGATE でエラーを除いて集計する", a)
                                break
            # 7. 文字の数字が足されていない（エラーにならないのにおかしい）
            if not err:
                for fn, args in _calls(f, ['SUM', 'AVERAGE']):
                    for t in args:
                        b = _parse_ref(t)
                        if b:
                            tx = [_addr(rr, cc) for rr, cc, vv in g.cells(b) if _is_text_number(vv)]
                            if tx:
                                add(f"{fn}({_box_addr(b)}) の範囲の文字の数字 {len(tx)} 個（{' '.join(tx[:6])}{'…' if len(tx) > 6 else ''}）"
                                    "は計算に入っていない＝数に直す（棚「選んだ列の文字の数字を数にする」）", a)
            # 8. 近似一致の検索（エラーにならず別の行の値を返す）
            for fn, flag, _pos in approx_calls(f):
                why = "第 4 引数が省略" if flag == 'omitted' and fn != 'MATCH' else ("第 3 引数が省略" if flag == 'omitted' else f"照合が {flag}")
                hit = ''
                for fn2, args in _calls(f, [fn]):
                    if fn2 == 'VLOOKUP' and len(args) >= 2:
                        kb, tab = _parse_ref(args[0]), _parse_ref(args[1])
                        if kb and tab and (kb[0], kb[1]) == (kb[2], kb[3]):
                            kv = g.get(kb[0], kb[1])
                            rows = [(rr, vv) for rr, _c, vv in g.cells((tab[0], tab[1], tab[2], tab[1]))]
                            at = next((rr for rr, vv in rows if vv == kv), None)
                            try:
                                idx = int(float(args[2])) if len(args) >= 3 else None
                            except ValueError:
                                idx = None
                            if kv in (None, ''):
                                pass
                            elif at is None and not err:
                                hit = f"（{_addr(kb[0], kb[1])}{_q(kv, nv)} は表に無いのに {a}{_q(v, nv)} を返している）"
                            elif at is not None and err:
                                hit = (f"（{_addr(kb[0], kb[1])}{_q(kv, nv)} は表の {_addr(at, tab[1])} にあるのに見つからない＝"
                                       "表が昇順に並んでいないので近似一致の探し方が外れた）")
                            elif at is not None and idx:
                                want = g.get(at, tab[1] + idx - 1)
                                if want != v:
                                    hit = (f"（{_addr(kb[0], kb[1])}{_q(kv, nv)} は表の {_addr(at, tab[1])} にあり"
                                           f"{'' if nv else '本当は「' + str(want) + '」'}なのに、{a}{_q(v, nv)} を返している＝"
                                           "表が昇順に並んでいないので近似一致が別の行を拾った）")
                add(f"{fn} が近似一致（{why}）＝探す値が表に無いと、エラーにならず別の行の値を返す{hit}。"
                    "完全一致なら FALSE（MATCH は 0）を書く", a)
            # 9. 空白で一致しない比較（=FILTER(…,A2:A9="田中") / COUNTIF(A2:A9,"田中")）
            pairs = [(m.group(1), m.group(2)) for m in re.finditer(r'(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)\s*=\s*"([^"]*)"', f)]
            for fn, args in _calls(f, ['COUNTIF', 'SUMIF', 'AVERAGEIF']):
                if len(args) >= 2 and args[1].startswith('"') and args[1].endswith('"'):
                    pairs.append((args[0], args[1][1:-1]))
            for rt, lit in pairs:
                b = _parse_ref(rt)
                if b and lit:
                    sp = [_addr(rr, cc) for rr, cc, vv in g.cells(b)
                          if isinstance(vv, str) and vv != lit and vv.strip(' 　') == lit]
                    if sp:
                        add(f"{' '.join(sp[:6])} の値は前後に空白があり、式の{'' if nv else '「' + lit + '」'}と一致しない"
                            "（値の空白を消すか、TRIM で比べる）", a)

    out = []
    for msg, addrs in list(found.items())[:limit]:
        uniq = list(dict.fromkeys(addrs))
        out.append(f"{' '.join(uniq[:6])}{'…' if len(uniq) > 6 else ''}: {msg}")
    for tok, msg in explain.items():
        out.append(f"式の読み解き: {msg}")
    return out


def _lev1(a, b):
    """編集距離がちょうど 1 か（置き換え・足し・消し 1 字）。"""
    if a == b or abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    if len(a) > len(b):
        a, b = b, a
    for k in range(len(b)):
        if b[:k] + b[k + 1:] == a:
            return True
    return False


def spelling_notes(fv, r0, c0, header_row=None, nv=False, limit=8):
    """英字の表記ゆれ（大文字小文字だけ違う・1 字違い）→ 行のリスト（純 Python）。

    日本語の名前は対象にしない（田中さん: 斉藤と齋藤、田辺と田邊は揺れと言い切れない）。多い方を正しい形とみる。"""
    out = []
    ncols = max((len(r) for r in fv), default=0)
    for j in range(ncols):
        cnt, where = {}, {}
        for i, row in enumerate(fv):
            if header_row is not None and r0 + i <= header_row:
                continue
            v = row[j] if j < len(row) else None
            if isinstance(v, str) and re.fullmatch(r"[A-Za-z][A-Za-z .'\-]{2,}", v.strip()):
                s = v.strip()
                cnt[s] = cnt.get(s, 0) + 1
                where.setdefault(s, []).append(_addr(r0 + i, c0 + j))
        words = sorted(cnt, key=lambda s: (-cnt[s], s))
        seen = set()
        for k, w in enumerate(words):
            for x in words[k + 1:]:
                if x in seen:
                    continue
                if w.lower() == x.lower() or (len(w) >= 4 and _lev1(w.lower(), x.lower())):
                    seen.add(x)
                    tie = cnt[w] == cnt[x]
                    out.append(f"{' '.join(where[x][:4])}{'' if nv else '「' + x + '」'} は"
                               f"{'' if nv else '「' + w + '」'}（{where[w][0]}）の書き違い？"
                               + ("（同じ数なので、どちらが正しいかは人が決める）" if tie else "")
                               + ("　大文字小文字だけ違う" if w.lower() == x.lower() else ""))
    return out[:limit]


def workbook_names_udfs(wb):
    """COM: ブックの名前 {名前: 参照先} と VBA の Function 名の集合（読めなければ空）。"""
    names, udfs = {}, set()
    try:
        for n in wb.Names:
            try:
                names[str(n.Name).split('!')[-1]] = str(n.RefersTo)
            except Exception:
                pass
    except Exception:
        pass
    try:
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            k = int(cm.CountOfLines)
            if k:
                for m in re.finditer(r'^\s*(?:Public\s+)?Function\s+([^\W\d]\w*)', cm.Lines(1, k), re.M | re.I):
                    udfs.add(m.group(1).upper())
    except Exception:
        pass
    return names, udfs
