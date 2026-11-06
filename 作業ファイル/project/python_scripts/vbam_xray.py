# -*- coding: utf-8 -*-
"""vbam_xray.py — シートのレントゲン（地図）。2026-11-06 段1。

構想（Excelドック）: シートを一瞬で透視する。骨＝見出し、筋肉＝入力、血管＝式、病巣＝エラー・ゴミ。
棚を撃つ・AI が書く、どの手の前と後にも撮れる共通の目。値は出さず形だけ（no-values でも動く）。

作り:
  読む（COM）  read_sheet(ws)  … UsedRange.Value を 1 回・SpecialCells(式) を 1 回・あとは範囲でまとめて取れる物だけ。
                                  セルごとの COM 呼び出しはしない。
  判定（純 Python）analyze(sheet) … Excel 非依存。pytest で試せる（test_xray.py）。
  描く         render_svg(map)    … 同じ地図から SVG の写真（骨＝白・筋肉＝青・血管＝緑・病巣＝赤）。

地図が確信できない所は必ず "unknown" と理由を書く（推して埋めない）。
行の役割: H 見出し / D 明細 / T 合計・小計 / B 表の中の空行 / G 小見出し・区分の行 / N 表の外（題・注記）
"""
import re
import sys
import json
import time
import unicodedata
import datetime as _dt
from decimal import Decimal

# ----------------------------------------------------------------
# 定数・正規表現
# ----------------------------------------------------------------
_ERR_CODES = {
    -2146826281: '#DIV/0!', -2146826243: '#SPILL!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
    -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!',
}
BIG_CELLS = 120000          # これを超えたら長い表の近道を試す
LONG_ROWS = 3000
HEAD_ROWS = 400
SAMPLE_STRIDE_TARGET = 40000

_NUMTXT_RE = re.compile(r'^[+\-−△▲]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$|^[+\-−△▲]?\d+(?:\.\d+)?$')
_UNIT_RE = re.compile(r'^[+\-−]?\d[\d,]*(?:\.\d+)?\s*(?:㎡|m2|m²|㎥|km|㎞|kg|㎏|g|ha|個|台|人|件|箇所|本|枚|万|億|千|%|％)$')
_YEN_RE = re.compile(r'^[¥￥\\]?\s*[+\-−△▲]?\s*[\d,]+(?:\.\d+)?\s*円?$')
_PAREN_NUM_RE = re.compile(r'^[\(（]\s*[+\-−]?\s*[\d,]+(?:\.\d+)?\s*[\)）]$')
_TRI_NUM_RE = re.compile(r'^[△▲]\s*[\d,]+(?:\.\d+)?$')
_DATE_PATTERNS = [
    re.compile(r'^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}$'),
    re.compile(r'^\d{4}年\d{1,2}月\d{1,2}日?$'),
    re.compile(r'^(?:令和|平成|昭和|R|H|S)\s*\d{1,2}(?:元)?[/.\-年]\s*\d{1,2}[/.\-月]\s*\d{1,2}日?$'),
    re.compile(r'^\d{1,2}/\d{1,2}$'),
]
_YM_RE = re.compile(r'^\d{4}[/\-.]\d{1,2}$|^\d{4}年\d{1,2}月$')
_PERIOD_RE = re.compile(r'(?:\d{1,4}[/年.\-]\d{1,2}[/月.\-]?\d{0,2}日?\s*[～~〜]\s*\d{1,4}[/月.\-]?\d{0,2}日?|\d{1,4}年\d{1,2}月\d{1,2}日から)')
_TOTAL_WORDS = ('合計', '小計', '総計', '総合計', '累計', '中計', '計', '合 計', '小 計', '総 計', 'total', 'subtotal', 'sum', 'grand total', '全体', '全国', '総数', '総額', '合計額')
_TOTAL_EXACT_RE = re.compile(r'^[\(（]?\s*(?:総|合|小|累|中|全|通)?\s*計\s*[\)）]?$|^[\(（]?\s*(?:総合計|合計額|総数|総額|合計|小計|累計|総計)\s*[\)）]?$|^(?:sub)?total$|^grand total$|^sum$', re.IGNORECASE)
_TOTAL_SUFFIX_RE = re.compile(r'(?:合計|小計|総計|累計|中計)$|[^\W\d_]計$')
_NOTE_HEAD_RE = re.compile(r'^\s*[（(]?\s*(?:注|※|＊|\*|出典|備考|資料|（注）|\(注\)|注記|参考|付記|　?注\d|[①-⑩])')
_SUM_RE = re.compile(r'^=\s*(?:SUM|SUBTOTAL|AGGREGATE)\s*\(', re.IGNORECASE)
_A1_REF_RE = re.compile(r"(?:'[^']+'|[A-Za-z0-9_\.]+)?!?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
_CELL_RE = re.compile(r"(?<![A-Za-z0-9_!'\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")
_VLOOKUP_LIT_RE = re.compile(r'\b(?:VLOOKUP|HLOOKUP)\s*\(([^,()]+(?:\([^()]*\))?[^,()]*),([^,]+),\s*(\d+)\s*[,)]', re.IGNORECASE)
_EXT_RE = re.compile(r"\[[^\]]+\.(?:xls[xmb]?|xlsm|csv)\]|'[A-Za-z]:\\|https?://", re.IGNORECASE)

_HW_ALNUM = re.compile(r'[0-9A-Za-z]')
_FW_ALNUM = re.compile(r'[０-９Ａ-Ｚａ-ｚ]')
_HW_KANA = re.compile(r'[ｦ-ﾟ]')
_WS_EDGE = re.compile(r'^[\s　 \t\r\n]+|[\s　 \t\r\n]+$')
_WS_ONLY = re.compile(r'^[\s　 \t\r\n]+$')
_KANA_RE = re.compile(r'^[ぁ-ゟァ-ヿｦ-ﾟー\s　・ー]+$')
_KANJI_RE = re.compile(r'[一-龥]')

_HDR_PHONE = re.compile(r'電話|TEL|ＴＥＬ|携帯|FAX|ＦＡＸ|Tel', re.IGNORECASE)
_HDR_POST = re.compile(r'郵便|〒|ゆうびん|ZIP|zip')
_HDR_ADDR = re.compile(r'住所|所在地|番地|居所|都道府県|市区町村')
_HDR_NAME = re.compile(r'氏名|姓名|お名前|名前|担当者?名?$|代表者|責任者|申請者|依頼者|被保険者|職員名|生徒名|社員名|ご芳名|^名$|^氏$|起案者|作成者|決裁者|承認者|記入者|入力者')
_HDR_KANA = re.compile(r'ふりがな|フリガナ|ﾌﾘｶﾞﾅ|よみ$|ヨミ|カナ|ｶﾅ|読み')
_HDR_DATE = re.compile(r'日付|年月日|日時|生年月日|date|Date|DATE|^日$|期日|締切|納期|開始日|終了日|入社日|契約日|発行日|登録日|更新日|作成日|受付日|取引日|支払日|入金日|年月$|基準日|開催日|実施日|公開日|施行日')
_HDR_MONEY = re.compile(r'金額|円|費|料|額|収入|支出|売上|単価|価格|収益|原価|利益|残高|給与|賃金|税|予算|決算|歳入|歳出|代金|報酬|手当|負担|債|資産|資本|貸付|借入|支払|入金|出金|相当額|価額|保険料|諸経費|価|千円|百万円|万円|億円')
_HDR_QTY = re.compile(r'数量|件数|人数|個数|台数|数$|数\)|部数|回数|面積|戸数|世帯|率|割合|比|指数|人口|人$|^人|個|台|件|本数|時間|日数|年数|枚数|㎡|総数|延べ|在庫')
_HDR_ID = re.compile(r'^(?:No\.?|ＮＯ|Ｎｏ\.?|№|#|NO\.?)$|番号$|^番号|^ID$|^Id$|^ＩＤ$|連番|^通し|伝票|整理番号|管理番号|受付番号|項番|順位|^順$|^番$|会員番号|社員番号|^no$', re.IGNORECASE)
_HDR_CODE = re.compile(r'コード|ｺｰﾄﾞ|CD$|Code|code|区分コード|^CODE$|記号|品番|型番|分類番号|識別')
_HDR_NOTE = re.compile(r'備考|メモ|注記|摘要|コメント|説明|理由|内容|その他|特記|注$|^注|概要|詳細|事由|経緯')
_HDR_MAIL = re.compile(r'メール|mail|Mail|MAIL|E-mail')
_HDR_NAMEORG = re.compile(r'名称|会社|法人|施設|団体|事業所|駅|学校|機関|店|組織|部署|科目|区分|種別|名$')

_NOT_TOTAL_KEI = ('統計', '会計', '設計', '集計', '家計', '時計', '推計', '体温計', '累計表', '寒暖計', '目盛計', '気圧計', '温度計', '歩数計')
_FACES = ('日付', '金額', '数量', '番号', 'コード', '電話', '郵便', '住所', '氏名', 'カナ', '文字', '式')


# ----------------------------------------------------------------
# 小さな純 Python の部品
# ----------------------------------------------------------------
def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_num(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n


def addr_of(r1, c1, r2=None, c2=None):
    a = f'{col_letter(c1)}{r1}'
    if r2 is not None and (r2 != r1 or c2 != c1):
        a += f':{col_letter(c2)}{r2}'
    return a


def parse_addr(a):
    """'A1:B3' / '$A$1:$B$3' / 'Sheet!A1' → (r1,c1,r2,c2) か None。"""
    a = a.split('!')[-1].replace('$', '')
    m = re.match(r'^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$', a)
    if not m:
        return None
    c1, r1 = col_num(m.group(1)), int(m.group(2))
    c2 = col_num(m.group(3)) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def parse_multi_addr(s):
    """'A1:B3,D4,F1:F9' → [(r1,c1,r2,c2)]。"""
    out = []
    for part in (s or '').split(','):
        p = parse_addr(part.strip())
        if p:
            out.append(p)
    return out


def _runs(nums):
    """整数の列 → '3-5,8' 形式。"""
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        out.append(str(nums[i]) if i == j else f'{nums[i]}-{nums[j]}')
        i = j + 1
    return ','.join(out)


def kind_of(v):
    """セルの型の 1 文字。E 空 / W 空白だけ / T 文字 / N 数 / D 日付 / X エラー / B 真偽。"""
    if v is None:
        return 'E'
    t = type(v)
    if t is str:
        if v == '':
            return 'E'
        if _WS_ONLY.match(v):
            return 'W'
        return 'T'
    if t is float:
        return 'N'
    if t is bool:
        return 'B'
    if t is int:
        return 'X' if v <= -2146826200 else 'N'
    if isinstance(v, (_dt.datetime, _dt.date)):
        return 'D'
    if isinstance(v, Decimal):
        return 'N'
    if hasattr(v, 'year') and hasattr(v, 'month'):      # pywintypes.datetime
        return 'D'
    return 'T'


def nfkc(s):
    return unicodedata.normalize('NFKC', s)


def clean_text(v):
    return v.strip().replace('　', ' ').replace(' ', ' ') if isinstance(v, str) else ''


def is_text_number(s):
    """文字の数字（カンマ・空白・円付き）か。文字列 s は前後空白あり得る。"""
    t = nfkc(s).strip().replace(' ', '')
    if not t or len(t) > 24:
        return False
    if _NUMTXT_RE.match(t) or _YEN_RE.match(t):
        return True
    return False


def is_text_date(s):
    t = nfkc(s).strip()
    if not t or len(t) > 20:
        return False
    for p in _DATE_PATTERNS:
        if p.match(t):
            return True
    return False


def a1_to_r1c1_rel(formula, row, col):
    """A1 形式の式を、置かれたセルからの相対形に直す（比べるための正規形。純 Python）。"""
    def _rep(m):
        cabs, cl, rabs, rn = m.group(1), m.group(2), m.group(3), int(m.group(4))
        c = col_num(cl)
        rr = f'R{rn}' if rabs else f'R[{rn - row}]'
        cc = f'C{c}' if cabs else f'C[{c - col}]'
        return rr + cc
    # 文字列リテラルの中は触らない
    parts = re.split(r'("[^"]*")', formula)
    return ''.join(p if p.startswith('"') else _CELL_RE.sub(_rep, p) for p in parts)


def _refs_in_formula(f):
    """式の中のセル範囲 → [(r1,c1,r2,c2, 他シートか)]。"""
    out = []
    if not isinstance(f, str):
        return out
    for m in re.finditer(r"((?:'[^']+'|[A-Za-z0-9_\.]+)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", f):
        pre = m.group(1)
        c1, r1 = col_num(m.group(2)), int(m.group(3))
        if m.group(4):
            c2, r2 = col_num(m.group(4)), int(m.group(5))
        else:
            c2, r2 = c1, r1
        st = m.start()
        if st > 0 and (f[st - 1].isalnum() or f[st - 1] == '_') and not pre:
            continue
        out.append((min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2), bool(pre)))
    return out


# ----------------------------------------------------------------
# COM 側（読むだけ）。セルごとの呼び出しはしない。
# ----------------------------------------------------------------
def _grid(v, nr, nc):
    if v is None:
        return [[None] * nc for _ in range(nr)]
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) for r in v]


def _merge_areas(ws, ur, values, r0, c0, nr, nc, cap=400, budget=None, wb=None):
    """結合セルの範囲の一覧。MergeCells が False なら 1 回で終わり。
    結合の左上になれるセル（右か下が空のセル）だけを候補にし、複数領域の MergeCells が False になる群を
    まとめて捨てる（群の中に結合があれば二分して探す）。Find(書式) は結合でないセルまで返すので使わない。
    空の結合（左上も空）は値が無いので見ない。"""
    try:
        state = ur.MergeCells
    except Exception:
        return [], 'unknown: MergeCells を読めない'
    if state is False:
        return [], None
    got = merges_from_file(wb, ws.Name, ws)
    if got is not None:
        if str(getattr(wb, 'FullName', '')).lower().endswith('.xls'):
            # .xls はファイルに空の結合や範囲外の結合も入っている: COM の探し方と同じく、値のある左上だけ残す
            got = [a for a in got if 0 <= a[0] - r0 < nr and 0 <= a[1] - c0 < nc and values[a[0] - r0][a[1] - c0] not in (None, '')]
        if got:
            return list(got), None
        if str(getattr(wb, 'FullName', '')).lower().endswith('.xls'):
            return [], None
    t_end = (time.perf_counter() + budget) if budget else None

    def empty(v):
        return v is None or v == ''

    cands = []
    for i in range(nr):
        row = values[i]
        nxt = values[i + 1] if i + 1 < nr else None
        for j in range(nc):
            if empty(row[j]):
                continue
            right_e = (j + 1 < nc) and empty(row[j + 1])
            below_e = nxt is not None and empty(nxt[j])
            if right_e or below_e:
                cands.append((i, j))
    areas = []
    covered = set()
    state_box = {'timeout': False, 'calls': 0}

    def addr(c):
        return f'{col_letter(c0 + c[1])}{r0 + c[0]}'

    def group_false(cells):
        try:
            return ws.Range(','.join(addr(c) for c in cells)).MergeCells is False
        except Exception:
            return False

    def search(cells):
        cells = [c for c in cells if c not in covered]
        if not cells:
            return
        if t_end is not None and time.perf_counter() > t_end:
            state_box['timeout'] = True
            return
        if len(cells) == 1 or len(','.join(addr(c) for c in cells)) <= 8:
            for c in cells:
                if c in covered:
                    continue
                try:
                    cell = ws.Cells(r0 + c[0], c0 + c[1])
                    if cell.MergeCells is True:
                        p = parse_addr(cell.MergeArea.Address)
                        if p:
                            areas.append(p)
                            for ii in range(p[0] - r0, p[2] - r0 + 1):
                                for jj in range(p[1] - c0, p[3] - c0 + 1):
                                    covered.add((ii, jj))
                except Exception:
                    pass
            return
        if group_false(cells):
            return
        mid = len(cells) // 2
        search(cells[:mid])
        search(cells[mid:])

    # まず行ごとの範囲（"B6:K6,B7:K7,…" は 1 本 8 字ほど）で、結合の無い行をまとめて捨てる。
    # 結合のある行だけ、その行の候補セルの群へ降りて探す。
    row_set = sorted({c[0] for c in cands})
    c_first, c_last = col_letter(c0), col_letter(c0 + nc - 1)

    def row_addr(i):
        return f'{c_first}{r0 + i}:{c_last}{r0 + i}'

    def rows_false(rs):
        try:
            return ws.Range(','.join(row_addr(i) for i in rs)).MergeCells is False
        except Exception:
            return False

    by_row = {}
    for c in cands:
        by_row.setdefault(c[0], []).append(c)

    def search_rows(rs):
        if not rs or state_box['timeout'] or len(areas) >= cap:
            return
        if t_end is not None and time.perf_counter() > t_end:
            state_box['timeout'] = True
            return
        if rows_false(rs):
            return
        if len(rs) == 1:
            cells = by_row[rs[0]]
            group, glen = [], 0
            groups = []
            for c in cells:
                al = len(addr(c)) + 1
                if glen + al > 230:
                    groups.append(group)
                    group, glen = [], 0
                group.append(c)
                glen += al
            if group:
                groups.append(group)
            for g in groups:
                search(g)
            return
        mid = len(rs) // 2
        search_rows(rs[:mid])
        search_rows(rs[mid:])

    # 行の群（アドレスは 255 字まで）
    grp, glen, rgroups = [], 0, []
    for i in row_set:
        al = len(row_addr(i)) + 1
        if glen + al > 230:
            rgroups.append(grp)
            grp, glen = [], 0
        grp.append(i)
        glen += al
    if grp:
        rgroups.append(grp)
    for g in rgroups:
        search_rows(g)
        if state_box['timeout'] or len(areas) >= cap:
            break
    if len(areas) >= cap:
        return areas, f'unknown: 結合が {cap} 件を超えた（先頭 {cap} 件だけ読んだ）'
    if state_box['timeout']:
        return areas, f'unknown: 結合が多く時間切れ（{len(areas)} 件まで読んだ。残りは未読）'
    return areas, None


_FILE_MERGE_CACHE = {}


def _xls_merges(path):
    """旧形式 .xls（BIFF8）の結合セルを、ファイルの MERGEDCELLS レコードから直接読む。
    {シート名: [(行1, 列1, 行2, 列2)…]}（1 始まり）。読めなければ None（COM の探し方に戻る）。"""
    import struct
    with open(path, 'rb') as f:
        raw = f.read()
    if raw[:8] != bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]):
        return None
    ss = 1 << struct.unpack_from('<H', raw, 30)[0]
    first_dir, = struct.unpack_from('<I', raw, 48)
    cutoff, = struct.unpack_from('<I', raw, 56)
    n_difat, = struct.unpack_from('<I', raw, 72)
    first_difat, = struct.unpack_from('<I', raw, 68)
    difat = list(struct.unpack_from('<109I', raw, 76))
    sec = first_difat
    for _ in range(n_difat):
        if sec >= 0xFFFFFFFA:
            break
        off = (sec + 1) * ss
        vals = struct.unpack_from('<%dI' % (ss // 4), raw, off)
        difat += list(vals[:-1])
        sec = vals[-1]
    fat = []
    for sid in difat:
        if sid >= 0xFFFFFFFA:
            continue
        off = (sid + 1) * ss
        fat += struct.unpack_from('<%dI' % (ss // 4), raw, off)

    def chain(start):
        out, seen = [], set()
        while start < 0xFFFFFFFA and start not in seen and start < len(fat):
            seen.add(start)
            out.append(raw[(start + 1) * ss:(start + 2) * ss])
            start = fat[start]
        return b''.join(out)

    d = chain(first_dir)
    stream, ssize = None, 0
    for k in range(0, len(d), 128):
        nlen, = struct.unpack_from('<H', d, k + 64)
        nm = d[k:k + max(nlen - 2, 0)].decode('utf-16-le', 'replace')
        if nm in ('Workbook', 'Book'):
            st, = struct.unpack_from('<I', d, k + 116)
            ssize, = struct.unpack_from('<Q', d, k + 120)
            if ssize < cutoff:
                return None
            stream = chain(st)[:ssize]
            break
    if stream is None:
        return None
    n = len(stream)
    if n < 8 or struct.unpack_from('<HH', stream, 0)[0] != 0x0809 or struct.unpack_from('<H', stream, 4)[0] != 0x0600:
        return None            # BIFF8 以外は読まない
    sheets, pos = [], 0
    while pos + 4 <= n:
        rid, ln = struct.unpack_from('<HH', stream, pos)
        body = stream[pos + 4:pos + 4 + ln]
        if rid == 0x0085 and ln >= 8:
            lb, = struct.unpack_from('<I', body, 0)
            cch = body[6]
            wide = body[7] & 1
            nm = body[8:8 + cch * (2 if wide else 1)].decode('utf-16-le' if wide else 'latin-1', 'replace')
            sheets.append((nm, lb))
        elif rid == 0x000A and not sheets:
            pass
        elif rid == 0x000A:
            break
        pos += 4 + ln
    out = {}
    for nm, lb in sheets:
        pos, depth, areas = lb, 0, []
        while pos + 4 <= n:
            rid, ln = struct.unpack_from('<HH', stream, pos)
            if rid == 0x0809:
                depth += 1
            elif rid == 0x000A:
                depth -= 1
                if depth <= 0:
                    break
            elif rid == 0x00E5 and depth == 1:
                cnt, = struct.unpack_from('<H', stream, pos + 4)
                for q in range(cnt):
                    r1, r2, c1, c2 = struct.unpack_from('<4H', stream, pos + 6 + q * 8)
                    areas.append((r1 + 1, c1 + 1, r2 + 1, c2 + 1))
            pos += 4 + ln
        out[nm] = areas
    return out


def merges_from_file(wb, sheet_name, ws=None):
    """保存済みの .xlsx/.xlsm なら、ファイルの中の <mergeCell> を直接読む（COM を使わず数 ms）。
    読めなければ None（呼び出し側が COM の探し方に戻る）。未保存の変更があるブックでは使わない。"""
    import os
    import zipfile
    try:
        if wb is None:
            return None
        saved = wb.Saved is True
        path = wb.FullName
        is_xls = path.lower().endswith('.xls')
        if not (is_xls or path.lower().endswith(('.xlsx', '.xlsm'))) or not os.path.exists(path):
            return None
        mt = os.path.getmtime(path)
        key = (path, mt)
        book = _FILE_MERGE_CACHE.get(key)
        if book is None and is_xls:
            book = _xls_merges(path)
            if book is None:
                return None
            _FILE_MERGE_CACHE.clear()
            _FILE_MERGE_CACHE[key] = book
        elif book is None:
            book = {}
            with zipfile.ZipFile(path) as z:
                wbx = z.read('xl/workbook.xml').decode('utf-8', 'replace')
                rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')
                rid2t = {}
                for m in re.finditer(r'<Relationship[^>]*>', rels):
                    tag = m.group(0)
                    i = re.search(r'Id="([^"]+)"', tag)
                    t = re.search(r'Target="([^"]+)"', tag)
                    if i and t:
                        rid2t[i.group(1)] = t.group(1)
                import html as _html
                for m in re.finditer(r'<sheet[^>]*>', wbx):
                    tag = m.group(0)
                    nm = re.search(r'name="([^"]*)"', tag)
                    rid = re.search(r'r:id="([^"]+)"', tag)
                    if not nm or not rid or rid.group(1) not in rid2t:
                        continue
                    tgt = rid2t[rid.group(1)].lstrip('/')
                    if not tgt.startswith('xl/'):
                        tgt = 'xl/' + tgt
                    raw = z.read(tgt)
                    k = raw.rfind(b'<mergeCells')
                    areas = []
                    if k >= 0:
                        for mm in re.finditer(rb'<mergeCell\s+ref="([^"]+)"', raw[k:]):
                            pr = parse_addr(mm.group(1).decode())
                            if pr:
                                areas.append(pr)
                    book[_html.unescape(nm.group(1))] = areas
            _FILE_MERGE_CACHE.clear()
            _FILE_MERGE_CACHE[key] = book
        got = book.get(sheet_name)
        if got is not None and not saved and got and ws is not None:
            # 未保存の変更があるとファイルが古いかもしれない: 先頭と末尾の数件が今も結合かだけ確かめる
            for (a1, b1, a2, b2) in (got[:3] + got[-3:]):
                if ws.Range(addr_of(a1, b1, a2, b2)).MergeCells is not True:
                    return None
        return got
    except Exception:
        return None


def _visible_runs(rng):
    """SpecialCells(xlCellTypeVisible) → [(a,b)] か None（取れない）。"""
    try:
        return rng.SpecialCells(12).Address
    except Exception:
        return None


def read_sheet(ws, wb=None, max_formula_areas=60, merge_budget=0.08, merge_cap=400, neighbors=False):
    """ws から純 Python が読める dict を作る。COM 呼び出しは範囲単位だけ。"""
    t0 = time.perf_counter()
    ur = ws.UsedRange
    pa = parse_addr(ur.Address)
    r0, c0 = pa[0], pa[1]
    nr, nc = pa[2] - pa[0] + 1, pa[3] - pa[1] + 1
    raw = ur.Value
    values = _grid(raw, nr, nc)
    sheet = {'name': ws.Name, 'r0': r0, 'c0': c0, 'nr': nr, 'nc': nc, 'values': values,
             'formulas': {}, 'merged': [], 'hidden_rows': [], 'hidden_cols': [], 'unknown': {}}
    # 式（SpecialCells 1 回）
    fcells = None
    try:
        fcells = ur.SpecialCells(-4123)
    except Exception:
        fcells = None
    if fcells is not None:
        try:
            # 領域は Areas で数える（Range.Address は複数領域だと 255 字で切れて、式を取りこぼす・VBA 版との同値試験で発見）
            areas = []
            n_areas = int(fcells.Areas.Count)
            if n_areas <= max_formula_areas:
                for k in range(1, n_areas + 1):
                    ar = fcells.Areas(k)
                    areas.append((int(ar.Row), int(ar.Column), int(ar.Row) + int(ar.Rows.Count) - 1, int(ar.Column) + int(ar.Columns.Count) - 1))
            if n_areas <= max_formula_areas:
                for (a1, b1, a2, b2) in areas:
                    sub = ws.Range(ws.Cells(a1, b1), ws.Cells(a2, b2)).Formula
                    g = _grid(sub, a2 - a1 + 1, b2 - b1 + 1)
                    for i, row in enumerate(g):
                        for j, f in enumerate(row):
                            if isinstance(f, str) and f.startswith('='):
                                sheet['formulas'][(a1 + i, b1 + j)] = f
            else:
                fg = _grid(ur.Formula, nr, nc)
                for i, row in enumerate(fg):
                    for j, f in enumerate(row):
                        if isinstance(f, str) and f.startswith('='):
                            sheet['formulas'][(r0 + i, c0 + j)] = f
        except Exception as e:
            sheet['unknown']['formulas'] = f'unknown: 式を読めない（{type(e).__name__}）'
    # 結合
    areas, why = _merge_areas(ws, ur, values, r0, c0, nr, nc, cap=merge_cap, budget=merge_budget, wb=wb)
    sheet['merged'] = areas
    if why:
        sheet['unknown']['merged'] = why
    # 隠した行・列（見える範囲の一覧を 1 回ずつ取って、その補集合）
    col_a = ws.Range(f'{col_letter(c0)}{r0}:{col_letter(c0)}{r0 + nr - 1}')
    vis = _visible_runs(col_a)
    if vis is None:
        sheet['unknown']['hidden_rows'] = 'unknown: 隠した行を取れない'
    else:
        vr = parse_multi_addr(vis)
        if not (len(vr) == 1 and vr[0][0] <= r0 and vr[0][2] >= r0 + nr - 1):
            seen = set()
            for (a1, _b1, a2, _b2) in vr:
                seen.update(range(a1, a2 + 1))
            sheet['hidden_rows'] = _to_pairs([r for r in range(r0, r0 + nr) if r not in seen])
    row_a = ws.Range(f'{col_letter(c0)}{r0}:{col_letter(c0 + nc - 1)}{r0}')
    vis = _visible_runs(row_a)
    if vis is None:
        sheet['unknown']['hidden_cols'] = 'unknown: 隠した列を取れない'
    else:
        vr = parse_multi_addr(vis)
        if not (len(vr) == 1 and vr[0][1] <= c0 and vr[0][3] >= c0 + nc - 1):
            seen = set()
            for (_a1, b1, _a2, b2) in vr:
                seen.update(range(b1, b2 + 1))
            sheet['hidden_cols'] = _to_pairs([c for c in range(c0, c0 + nc) if c not in seen])
    # テーブル・ピボット・グラフ・図形（個数と範囲だけ）
    objs = {}
    try:
        lo = []
        for lob in ws.ListObjects:
            lo.append(lob.Range.Address.replace('$', ''))
        objs['listobjects'] = lo
    except Exception:
        pass
    for key, attr in (('pivots', 'PivotTables'), ('charts', 'ChartObjects'), ('shapes', 'Shapes')):
        try:
            objs[key] = int(getattr(ws, attr)().Count)
        except Exception:
            try:
                objs[key] = int(getattr(ws, attr).Count)
            except Exception:
                pass
    sheet['objects'] = objs
    # ブック全体（名前・リンク）と、ほかのシートの先頭（対応表・比べる相手を探す材料）
    if wb is not None:
        sheet['book'] = read_book_facts(wb)
        if neighbors:
            sheet['neighbors'] = read_neighbors(ws, wb)
    sheet['read_sec'] = round(time.perf_counter() - t0, 4)
    return sheet


_BOOK_CACHE = {}


def read_neighbors(ws, wb, limit=6, max_cells=3000):
    """ほかのシートの使用範囲の先頭を範囲 1 回ずつで読む（対応表・比べる相手を探す材料。最大 limit 枚）。"""
    out = []
    try:
        n = int(wb.Worksheets.Count)
    except Exception:
        return out
    if n <= 1:
        return out
    k = 0
    try:
        mine = ws.Name
        for other in wb.Worksheets:
            nm = other.Name
            if nm == mine:
                continue
            k += 1
            if k > limit:
                out.append({'skipped': n - 1 - limit})
                break
            p = parse_addr(other.UsedRange.Address)
            nr, nc = p[2] - p[0] + 1, p[3] - p[1] + 1
            rows_read = max(1, min(nr, max_cells // max(nc, 1)))
            raw = other.Range(f'{col_letter(p[1])}{p[0]}:{col_letter(p[3])}{p[0] + rows_read - 1}').Value
            out.append({'name': nm, 'nr': nr, 'nc': nc, 'values': _grid(raw, rows_read, nc), 'complete': rows_read == nr})
    except Exception:
        pass
    return out


def read_book_facts(wb):
    key = id(wb)
    got = _BOOK_CACHE.get(key)
    if got is not None:
        return got
    facts = {'names': 0, 'broken_names': [], 'external_links': 0, 'sheets': 0}
    try:
        facts['sheets'] = int(wb.Worksheets.Count)
    except Exception:
        pass
    try:
        n = int(wb.Names.Count)
        facts['names'] = n
        if n <= 300:
            for nm in wb.Names:
                try:
                    if '#REF!' in (nm.RefersTo or ''):
                        facts['broken_names'].append(nm.Name)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        ls = wb.LinkSources(1)
        facts['external_links'] = len(ls) if ls else 0
    except Exception:
        pass
    _BOOK_CACHE.clear()
    _BOOK_CACHE[key] = facts
    return facts


def _to_pairs(nums):
    out = []
    for n in nums:
        if out and out[-1][1] == n - 1:
            out[-1] = (out[-1][0], n)
        else:
            out.append((n, n))
    return out


# ================================================================
# 純 Python の判定（Excel 非依存）
# ================================================================
_KT = {type(None): 'E', float: 'N', int: 'N', bool: 'B'}


def _kind_row(row):
    out = []
    ap = out.append
    for v in row:
        k = _KT.get(type(v))
        if k is None or (k == 'N' and type(v) is int):
            k = kind_of(v)
        ap(k)
    return out


def _nz(v):
    """文字セルの前後を整えた比較用の文字（全角空白・NBSP も空白扱い）。"""
    return v.replace('\u3000', ' ').replace('\u00a0', ' ').strip() if isinstance(v, str) else ''


def _is_total_label(txt):
    t = nfkc(txt).replace(' ', '').replace('\u3000', '')
    t = t.strip('()（）[]【】')
    if not t:
        return False
    if _TOTAL_EXACT_RE.match(t):
        return True
    if re.match(r'^(?:全体|総|区部|都)?平均$', t):
        return True
    if re.match(r'^(?:総合計|合計|小計|総計)[(\[]', t):
        return True
    return False


def _is_total_suffix_label(txt):
    t = nfkc(txt).replace(' ', '').strip('()（）[]【】')
    return len(t) <= 8 and bool(_TOTAL_SUFFIX_RE.search(t)) and t[-2:] in ('合計', '小計', '総計', '累計', '中計') or (len(t) <= 6 and t.endswith('計') and len(t) >= 2 and t not in ('統計', '会計', '設計', '集計', '累計表', '家計', '時計', '体温計', '計画', '推計'))


def _is_note_text(txt):
    return bool(_NOTE_HEAD_RE.match(txt))


def _fill_count(krow):
    return sum(1 for k in krow if k not in 'EW')


def classify_columns_hint(header_text):
    return header_text


def analyze(sheet, nv=False, sample_symptoms=True):
    """sheet（read_sheet の dict）→ 地図（dict）。Excel には触らない。"""
    t0 = time.perf_counter()
    vals = sheet['values']
    r0, c0 = sheet['r0'], sheet['c0']
    nr, nc = len(vals), (len(vals[0]) if vals else 0)
    formulas = sheet.get('formulas') or {}
    merged = list(sheet.get('merged') or [])
    # 結合は、中のどこかに値がある物だけ使う（ファイルから読んだ結合には、空の結合や範囲外の結合も入る。VBA 版の読みと同じ）
    if merged and vals:
        _nr0, _nc0 = len(vals), len(vals[0])

        def _has_val(m):
            for a in range(max(m[0], r0), min(m[2], r0 + _nr0 - 1) + 1):
                row = vals[a - r0]
                for b in range(max(m[1], c0), min(m[3], c0 + _nc0 - 1) + 1):
                    x = row[b - c0]
                    if x is not None and x != '':
                        return True
            return False
        merged = [m for m in merged if _has_val(m)]
    unknown_sheet = dict(sheet.get('unknown') or {})
    out = {'sheet': sheet.get('name'), 'used': addr_of(r0, c0, r0 + nr - 1, c0 + nc - 1) if nr and nc else '',
           'blocks': [], 'sheet_symptoms': [], 'unknown': [], 'no_values': bool(nv)}
    for k, why in unknown_sheet.items():
        out['unknown'].append({'item': k, 'reason': why})
    if nr == 0 or nc == 0:
        out['stats'] = {'rows': 0, 'cols': 0}
        return out
    big = nr * nc > BIG_CELLS
    if big and nr > LONG_ROWS and not sheet.get('_no_long'):
        res = _analyze_long(sheet, nv, t0)
        if res is not None:
            return res
    # --- 型の格子（M は結合の続き）
    K = [_kind_row(r) for r in vals]
    mtl = {}                 # 結合の左上 → (r2,c2)（0 起点・相対）
    mcover = {}
    for (a1, b1, a2, b2) in merged:
        i1, j1, i2, j2 = a1 - r0, b1 - c0, a2 - r0, b2 - c0
        if i1 < 0 or j1 < 0 or i2 >= nr or j2 >= nc:
            continue
        mtl[(i1, j1)] = (i2, j2)
        for i in range(i1, i2 + 1):
            for j in range(j1, j2 + 1):
                if (i, j) != (i1, j1):
                    mcover[(i, j)] = (i1, j1)
                    if K[i][j] in 'EW':
                        K[i][j] = 'M'
    fset = {(r - r0, c - c0) for (r, c) in formulas}
    fill = [_fill_count(k) for k in K]
    content_rows = [i for i in range(nr) if fill[i] > 0 and any(x not in 'EWM' for x in K[i])]
    ctx = {'vals': vals, 'K': K, 'r0': r0, 'c0': c0, 'nr': nr, 'nc': nc, 'mtl': mtl, 'mcover': mcover,
           'fset': fset, 'formulas': formulas, 'fill': fill, 'merged': merged,
           'hidden_rows': sheet.get('hidden_rows') or [], 'hidden_cols': sheet.get('hidden_cols') or [],
           'nv': nv, 'big': big, 'unknown_sheet': unknown_sheet}
    ctx['inner_note'] = _has_inner_note(vals)
    # --- 表の塊
    blocks = _find_blocks(ctx)
    for bi, bl in enumerate(blocks, 1):
        out['blocks'].append(_analyze_block(ctx, bl, bi))
    # --- シート全体
    _sheet_level(ctx, out, sheet, blocks, content_rows)
    if sheet.get('neighbors') is not None:
        hw = set()
        tabs = [b for b in out['blocks'] if b['kind'] == 'table' and b.get('header_rows')]
        if tabs:
            main = max(tabs, key=lambda b: len(b.get('columns') or []))
            pr = parse_addr(main['range'])
            for h in main['header_rows']:
                hw |= _words([vals[h - r0][j] for j in range(pr[1] - c0, pr[3] - c0 + 1)])
        out['neighbors'] = neighbors_map(hw, sheet['neighbors'])
    out['not_read'] = [
        {'item': 'M17 書式（罫線・列幅・表示形式・寄せ）', 'reason': 'セルごとにしか取れず、範囲でまとめて取る手が無いので読まない（S31〜S34 は作らない）'},
        {'item': '空の結合セル', 'reason': '左上も空の結合は値が無く、結合の探し方（値のあるセルから）では見つからない'}]
    out['stats'] = {'rows': nr, 'cols': nc, 'formulas': len(formulas), 'merged': len(merged),
                    'blocks': len(out['blocks'])}
    out['sec'] = round(time.perf_counter() - t0, 4)
    return out


def _analyze_long(sheet, nv, t0):
    """行が多い（数千行〜）シートの近道。先頭 HEAD_ROWS 行だけを判定にかけ、残りの行は空行・合計の語・注記を
    安く見分けるだけにする。先頭の塊が最終行まで続く 1 本の表のときだけ使う（それ以外は None=全部読む）。
    先頭だけで見た列の症状は「先頭だけ」と明記する（無い症状を言い切らない）。"""
    vals = sheet['values']
    r0, c0 = sheet['r0'], sheet['c0']
    nr, nc = len(vals), len(vals[0])
    for (a1, b1, a2, b2) in sheet.get('merged') or []:
        if a2 - r0 >= HEAD_ROWS:
            return None
    head = dict(sheet)
    head['values'] = vals[:HEAD_ROWS]
    head['formulas'] = {k: v for k, v in (sheet.get('formulas') or {}).items() if k[0] - r0 < HEAD_ROWS}
    head['hidden_rows'] = [(a, min(b, r0 + HEAD_ROWS - 1)) for (a, b) in sheet.get('hidden_rows') or [] if a - r0 < HEAD_ROWS]
    head['_no_long'] = True
    m = analyze(head, nv)
    tabs = [b for b in m['blocks'] if b['kind'] == 'table']
    if len(m['blocks']) != len(tabs) + sum(1 for b in m['blocks'] if b['kind'] == 'memo' and parse_addr(b['range'])[2] - r0 < 10):
        return None
    main = [b for b in tabs if parse_addr(b['range'])[2] - r0 >= HEAD_ROWS - 2]
    if len(main) != 1 or len(tabs) != 1:
        return None
    b = main[0]
    pr = parse_addr(b['range'])
    c1, c2 = pr[1] - c0, pr[3] - c0
    # 残りの行の役割（安い判定）
    last = nr - 1
    while last >= HEAD_ROWS and all((v is None or v == '') for v in vals[last]):
        last -= 1
    roles = {}
    for run in b['rows']:
        a, _, z = run['rows'].partition('-')
        for r in range(int(a), int(z or a) + 1):
            roles[r] = run['role']
    det_last = max((r for r, rl in roles.items() if rl in ('D', 'T')), default=r0 + HEAD_ROWS - 1)
    for i in range(HEAD_ROWS, last + 1):
        row = vals[i]
        seg = row[c1:c2 + 1]
        f = sum(1 for v in seg if v is not None and v != '')
        r = r0 + i
        if f == 0:
            roles[r] = 'B'
            continue
        role = 'D'
        for v in seg[:3]:
            if type(v) is str and v.strip():
                if _is_total_label(v) or (len(v) <= 8 and v.strip().endswith(('合計', '小計', '総計')) and sum(1 for x in seg if type(x) is str and x.strip()) <= 1):
                    role = 'T'
                break
        if role == 'D' and f <= 2 and i >= nr - 200:
            tx = next((v for v in seg if type(v) is str and v.strip()), None)
            if tx and _is_note_text(tx):
                role = 'N'
        roles[r] = role
    # 末尾の空行・N の手前までを表の最終行とする
    new_r2 = max(r for r, rl in roles.items() if rl in ('D', 'T')) if roles else pr[2]
    b['range'] = addr_of(pr[0], pr[1], max(pr[2], new_r2), pr[3])
    b['rows'] = _runs_of_roles({r - r0: rl for r, rl in roles.items() if r <= max(pr[2], new_r2)}, r0)
    dets = [r for r, rl in roles.items() if rl == 'D' and r <= max(pr[2], new_r2)]
    b['detail_rows'] = _runs(dets)
    b['total_rows'] = [r for r, rl in sorted(roles.items()) if rl == 'T']
    b['sampled'] = f'先頭 {HEAD_ROWS} 行で列の顔と症状を判定（残りは行の役割だけ）'
    m['unknown'].append({'block': b['id'], 'item': 'symptoms', 'reason': f'行が多いため先頭 {HEAD_ROWS} 行だけで症状を判定。残りの行の症状は未確認（無いとは言い切らない）'})
    m['stats'] = {'rows': nr, 'cols': nc, 'formulas': len(sheet.get('formulas') or {}), 'merged': len(sheet.get('merged') or []), 'blocks': len(m['blocks']), 'long': True}
    m['used'] = addr_of(r0, c0, r0 + nr - 1, c0 + nc - 1)
    m['sec'] = round(time.perf_counter() - t0, 4)
    return m


def _words(texts):
    return {nfkc(t).strip().replace(' ', '') for t in texts if isinstance(t, str) and t.strip()}


def _row_text_share_vals(row):
    f = [v for v in row if v not in (None, '')]
    return sum(1 for v in f if isinstance(v, str)) / len(f) if f else 0.0


def neighbors_map(header_words, neighbors):
    """ほかのシートの先頭から、対応表（2〜4 列・先頭列が一意）と、見出しが同じ「比べる相手」を見分ける（純 Python）。"""
    out = []
    for nb in neighbors or []:
        if 'skipped' in nb:
            out.append({'sheet': None, 'kind': 'unknown', 'reason': f"シートが多いため {nb['skipped']} 枚は読んでいない"})
            continue
        vals = nb['values']
        if not vals or not vals[0]:
            out.append({'sheet': nb['name'], 'kind': 'empty'})
            continue
        hi = next((i for i, r in enumerate(vals) if sum(1 for v in r if v not in (None, '')) >= 2), None)
        if hi is None:
            out.append({'sheet': nb['name'], 'kind': 'other'})
            continue
        hdr = vals[hi]
        hw = _words(hdr)
        used_cols = [j for j in range(len(hdr)) if any(r[j] not in (None, '') for r in vals[hi:])]
        ent = {'sheet': nb['name'], 'rows': nb['nr'], 'cols': nb['nc'], 'kind': 'other'}
        body = vals[hi + 1:]
        j0 = used_cols[0] if used_cols else 0
        keys = [r[j0] for r in body if r[j0] not in (None, '')]
        jac = len(hw & header_words) / max(1, len(hw | header_words)) if (header_words and hw) else 0.0
        if jac >= 0.6:
            ent['kind'] = 'comparable'
            ent['shared_header'] = len(hw & header_words)
        elif (2 <= len(used_cols) <= 4 and len(keys) >= 3 and _row_text_share_vals(hdr) >= 0.6
                and len(set(map(str, keys))) >= 0.95 * len(keys)):
            ent['kind'] = 'master'
            ent['key_col'] = col_letter(1 + j0)
            if not nb.get('complete'):
                ent['reason'] = 'シートが長く、先頭だけで先頭列の一意を確かめた'
        out.append(ent)
    return out


def _has_inner_note(vals):
    for row in vals:
        for v in row:
            if type(v) is str and len(v) > 2 and ('内数' in v or 'うち' in v or '再掲' in v or '内書' in v or '内訳' in v):
                return True
    return False


# ----------------------------------------------------------------
# 表の塊
# ----------------------------------------------------------------
def _row_occ(ctx, i):
    return ctx['fill'][i] > 0


_MISSING_MARKS = ('-', '－', 'ー', '―', '…', '・', 'x', 'X', '△', '▲', '*', '＊', '―', '‐', '‑')


def _numlike_text(v):
    """数に見える文字（(1)・△5・1,200・欠測の記号）。構造の判定では数として数える。"""
    if type(v) is not str:
        return False
    t = nfkc(v).strip().replace(' ', '')
    if not t:
        return False
    if t in _MISSING_MARKS:
        return True
    return bool(_PAREN_NUM_RE.match(t) or _TRI_NUM_RE.match(t) or (len(t) <= 24 and (_NUMTXT_RE.match(t) or _YEN_RE.match(t))))


def _row_text_share(krow, j1, j2, ctx=None, i=None):
    """行の j1..j2 の文字セルの割合（空は数えない・M は数えない）。数に見える文字は数として数える。"""
    seg = krow[j1:j2 + 1]
    t = seg.count('T')
    n = sum(1 for k in seg if k in 'NDB')
    x = seg.count('X')
    if ctx is not None and i is not None and t:
        vr = ctx['vals'][i]
        nl = sum(1 for j in range(j1, j2 + 1) if krow[j] == 'T' and _numlike_text(vr[j]))
        t -= nl
        n += nl
    tot = t + n + x
    return (t / tot) if tot else 0.0


def _row_is_data_like(ctx, i, j1, j2, ratio=0.0):
    """見出しではなく明細らしい行（数・日付・式のセルがあり、文字だけではない）。
    ratio > 0 のときは、埋まったセルのうち数・日付がその割合以上のときだけ明細らしいとする。"""
    kr = ctx['K'][i]
    seg = kr[j1:j2 + 1]
    nnum = sum(1 for k in seg if k in 'ND')
    vr = ctx['vals'][i]
    nnum += sum(1 for j in range(j1, j2 + 1) if kr[j] == 'T' and _numlike_text(vr[j]))
    ftxt = any(((i, j) in ctx['fset']) for j in range(j1, j2 + 1))
    if nnum == 0 and not ftxt:
        return False
    if ratio and nnum < ratio * sum(1 for k in seg if k not in 'EWM'):
        return False
    # 数が年（見出しの 2020 2021 …）ばかりなら見出し
    ys = 0
    for j in range(j1, j2 + 1):
        v = ctx['vals'][i][j]
        if kr[j] == 'N' and isinstance(v, float) and v == int(v) and 1900 <= v <= 2100:
            ys += 1
    if ys and ys == nnum and nnum >= 2:
        return False
    return True


def _find_blocks(ctx):
    K, nr, nc = ctx['K'], ctx['nr'], ctx['nc']
    fill = ctx['fill']
    raw = []
    prev_band = []
    i = 0
    while i < nr:
        if fill[i] == 0:
            i += 1
            continue
        j = i
        while j + 1 < nr and fill[j + 1] > 0:
            j += 1
        # 列の区切り
        occ = [0] * nc
        for ii in range(i, j + 1):
            kr = K[ii]
            for jj in range(nc):
                if kr[jj] not in 'EW':
                    occ[jj] = 1
        # 空行 1〜2 行の下の続きの帯: 前の帯の表（3 行以上）の列の中に収まり、先頭が明細らしければ、空の列で割らずに 1 つの塊にする
        # （明細の一部の列が空の行だけが続くと、表の続きが「分からない」2 つの塊に割れていた）
        if prev_band:
            P = prev_band[0]
            if (len(prev_band) == 1 and P['r2'] - P['r1'] >= 2 and 1 <= i - P['r2'] - 1 <= 2
                    and P['c1'] <= min(c for c in range(nc) if occ[c]) and max(c for c in range(nc) if occ[c]) <= P['c2']
                    and sum(occ) >= 2 and ctx['fill'][i] >= 2
                    and _row_is_data_like(ctx, i, P['c1'], P['c2'], 0.3)):
                for cc in range(P['c1'], P['c2'] + 1):
                    occ[cc] = 1
        prev_band = []
        jj = 0
        while jj < nc:
            if not occ[jj]:
                jj += 1
                continue
            k = jj
            while k + 1 < nc and occ[k + 1]:
                k += 1
            # この列範囲に値のある行だけに絞る（帯の上下の余りは他の列の行）
            rows_in = [ii for ii in range(i, j + 1) if any(x not in 'EW' for x in K[ii][jj:k + 1])]
            raw.append({'r1': rows_in[0], 'r2': rows_in[-1], 'c1': jj, 'c2': k})
            prev_band.append(raw[-1])
            jj = k + 1
        i = j + 1
    # 縦の続き（表の中の空行）を結合する
    raw.sort(key=lambda b: (b['c1'], b['r1']))
    merged_blocks = []
    for b in raw:
        prev = None
        for m in reversed(merged_blocks):
            if m['r1'] == m['r2'] and (m['c2'] - m['c1']) <= 1 and ctx['fill'][m['r1']] <= 2 and _single_text_band(ctx, m):
                continue                      # 1 行だけの区分・題の帯は、表の続きを探すときは素通りする
            ov = min(m['c2'], b['c2']) - max(m['c1'], b['c1']) + 1
            nar = min(m['c2'] - m['c1'] + 1, b['c2'] - b['c1'] + 1)
            if ov > 0 and ov >= 0.7 * nar and m['r2'] < b['r1']:
                prev = m
                break
        if prev is not None:
            gap = b['r1'] - prev['r2'] - 1
            same_cols = (abs(prev['c1'] - b['c1']) <= 1 and abs(prev['c2'] - b['c2']) <= 1)
            # 前の塊に明細らしい行がまだ無い（見出しだけ）のときも、次の塊の先頭が明細らしければ続き
            prev_header_only = not any(_row_is_data_like(ctx, ii, prev['c1'], prev['c2'], 0.3) for ii in range(prev['r1'], prev['r2'] + 1))
            prev_ok = (prev['r2'] - prev['r1'] >= 2) or (prev_header_only and ctx['fill'][prev['r1']] >= 2)
            if (gap <= 6 and same_cols and prev_ok and _row_is_data_like(ctx, b['r1'], b['c1'], b['c2'])
                    and ctx['fill'][b['r1']] >= 2
                    and (_row_is_data_like(ctx, prev['r2'], prev['c1'], prev['c2']) or prev_header_only)
                    and not _band_has_other(ctx, merged_blocks, prev, b)):
                prev['r2'] = b['r2']
                prev['c1'] = min(prev['c1'], b['c1'])
                prev['c2'] = max(prev['c2'], b['c2'])
                prev['gaps'] = prev.get('gaps', []) + [(b['r1'] - gap, b['r1'] - 1)]
                continue
        merged_blocks.append(dict(b))
    out = []
    for b in merged_blocks:
        out.extend(_split_at_repeated_header(ctx, b))
    out.sort(key=lambda b: (b['c1'], b['r1']))
    # 空の列をはさんだ細い右側（URL・備考など）が同じ行の範囲なら、同じ表の続き（途中の空の列 = S13）
    changed = True
    while changed:
        changed = False
        for a in out:
            for b in out:
                if a is b:
                    continue
                gap = b['c1'] - a['c2'] - 1
                if (1 <= gap <= 4 and abs(a['r1'] - b['r1']) <= 1 and abs(a['r2'] - b['r2']) <= 1 and a['r2'] - a['r1'] >= 3
                        and ((((b['c2'] - b['c1'] + 1) <= 1) and _dense(ctx, b))
                             or (gap == 1 and ((a['c2'] - a['c1'] + 1) <= 2) and _dense(ctx, a)))):
                    a['empty_cols'] = a.get('empty_cols', []) + list(range(a['c2'] + 1, b['c1']))
                    a['c2'] = b['c2']
                    a['r1'] = min(a['r1'], b['r1'])
                    a['r2'] = max(a['r2'], b['r2'])
                    out.remove(b)
                    changed = True
                    break
            if changed:
                break
    # 別の塊の中にすっぽり入っている小さな塊（表の中の区分の行など）は、その塊の一部
    keep = []
    for b in out:
        inside = any(o is not b and o['r1'] <= b['r1'] and o['r2'] >= b['r2'] and o['c1'] <= b['c1'] and o['c2'] >= b['c2']
                     and (o['r2'] - o['r1']) > (b['r2'] - b['r1']) for o in out)
        if not inside:
            keep.append(b)
    out = keep
    out.sort(key=lambda b: (b['r1'], b['c1']))
    return out


def _single_text_band(ctx, b):
    """1 行・1〜2 セルの文字だけの帯か（数・式を含まない）。"""
    kr = ctx['K'][b['r1']][b['c1']:b['c2'] + 1]
    return all(k in 'EWMT' for k in kr) and 'T' in kr


def _dense(ctx, b):
    """塊の行のうち、値のある行が 8 割以上か。"""
    K = ctx['K']
    n = b['r2'] - b['r1'] + 1
    k = sum(1 for i in range(b['r1'], b['r2'] + 1) if any(x not in 'EWM' for x in K[i][b['c1']:b['c2'] + 1]))
    return k >= 0.8 * n


def _split_at_repeated_header(ctx, b):
    """空行なしで 2 つ目の表が続くとき（途中に見出しだけの行があり、前後が明細）は、その行で塊を割る。"""
    K = ctx['K']
    r1, r2, c1, c2 = b['r1'], b['r2'], b['c1'], b['c2']
    width = c2 - c1 + 1
    if width < 3 or r2 - r1 < 5:
        return [b]
    cuts = []
    for i in range(r1 + 2, r2 - 1):
        kr = K[i]
        f = sum(1 for k in kr[c1:c2 + 1] if k not in 'EWM')
        if f < max(3, 0.6 * width):
            continue
        if _row_text_share(kr, c1, c2, ctx, i) < 0.99 or _row_is_data_like(ctx, i, c1, c2):
            continue
        if _row_is_data_like(ctx, i - 1, c1, c2) and _row_is_data_like(ctx, i + 1, c1, c2):
            texts = [ctx['vals'][i][j] for j in range(c1, c2 + 1) if kr[j] == 'T']
            if any(_is_total_label(t) for t in texts[:2]):
                continue
            # 文字の日付・文字の数字が入っている行は、見出しではなく文字で入った明細（CSV の貼り付けなど）
            if any(isinstance(t, str) and (is_text_date(t.strip()) or is_text_number(t.strip())) for t in texts):
                continue
            cuts.append(i)
    if not cuts:
        return [b]
    out = []
    start = r1
    for i in cuts:
        out.append({'r1': start, 'r2': i - 1, 'c1': c1, 'c2': c2})
        start = i
    out.append({'r1': start, 'r2': r2, 'c1': c1, 'c2': c2})
    return out


def _band_has_other(ctx, blocks, prev, b):
    for m in blocks:
        if m is prev:
            continue
        if m['r1'] == m['r2'] and (m['c2'] - m['c1']) <= 1 and ctx['fill'][m['r1']] <= 2 and _single_text_band(ctx, m):
            continue
        if m['r1'] > prev['r2'] and m['r2'] < b['r1'] and min(m['c2'], b['c2']) >= max(m['c1'], b['c1']):
            return True
    return False


# ----------------------------------------------------------------
# 塊の中身
# ----------------------------------------------------------------
_PLACEHOLDER_RE = re.compile(r'^[\s…・．.\-－―ーxXｘＸ※*＊]+$')


def _t_cells(ctx, i, c1, c2):
    """行 i の文字セルの数（「…」「－」「X」など値なしの印は数えない）。"""
    kr, vr = ctx['K'][i], ctx['vals'][i]
    n = 0
    for j in range(c1, c2 + 1):
        if kr[j] == 'T' and not _PLACEHOLDER_RE.match(nfkc(str(vr[j]))):
            n += 1
    return n


def _first_texts(ctx, i, j1, j2, n=3):
    """行 i の左から n 個までの文字セル（相対列, 文字）。"""
    out = []
    kr, vr = ctx['K'][i], ctx['vals'][i]
    for j in range(j1, j2 + 1):
        if kr[j] == 'T':
            out.append((j, vr[j]))
            if len(out) >= n:
                break
    return out


def _row_formula_total(ctx, i, j1, j2):
    """行 i の式の数と、縦の合計式（SUM/SUBTOTAL で上の行を束ねる）の数。"""
    r_abs = ctx['r0'] + i
    nf = 0
    ok = 0
    for j in range(j1, j2 + 1):
        f = ctx['formulas'].get((r_abs, ctx['c0'] + j))
        if f is None:
            continue
        nf += 1
        if _SUM_RE.match(f):
            for (a1, b1, a2, b2, ext) in _refs_in_formula(f):
                if not ext and a2 < r_abs and b1 <= ctx['c0'] + j <= b2:
                    ok += 1
                    break
    return nf, ok


def _analyze_block(ctx, bl, bid):
    K = ctx['K']
    r1, r2, c1, c2 = bl['r1'], bl['r2'], bl['c1'], bl['c2']
    r0, c0 = ctx['r0'], ctx['c0']
    width = c2 - c1 + 1
    blk = {'id': f'T{bid}', 'range': addr_of(r0 + r1, c0 + c1, r0 + r2, c0 + c2), 'kind': 'table',
           'unknown': []}
    rows_role = {}
    gaps = set()
    for (g1, g2) in bl.get('gaps', []):
        for g in range(g1, g2 + 1):
            if all(k in 'EWM' for k in K[g][c1:c2 + 1]):      # 実際に空の行だけが空行（表の中の 1 セルの題・区分の行は空行ではない）
                gaps.add(g)
    nonblank = [ii for ii in range(r1, r2 + 1) if ii not in gaps]

    def fill_in(ii):
        return sum(1 for k in K[ii][c1:c2 + 1] if k not in 'EWM')

    # 全部が 1 セル以下の行 → メモ（題・注記）
    if all(fill_in(ii) <= 1 for ii in nonblank):
        lens = [len(v) for ii in nonblank for v in [ctx['vals'][ii][c1]] if isinstance(v, str)]
        short_list = width == 1 and len(nonblank) >= 3 and lens and (sum(lens) / len(lens)) <= 14             and not any(_is_note_text(ctx['vals'][ii][c1]) for ii in nonblank if isinstance(ctx['vals'][ii][c1], str))
        if short_list:
            for ii in range(r1, r2 + 1):
                rows_role[ii] = 'U' if ii not in gaps else 'B'
            blk['kind'] = 'unknown'
            blk['unknown'].append({'item': 'rows', 'reason': '1 列だけで、短い語の一覧か文章かを決められない'})
            return _finish_block(ctx, blk, bl, rows_role, [], [], sorted(gaps))
        blk['kind'] = 'memo'
        for ii in range(r1, r2 + 1):
            rows_role[ii] = 'N' if ii not in gaps else 'B'
        return _finish_block(ctx, blk, bl, rows_role, [], [], sorted(gaps))
    is_form = _is_form(ctx, bl)
    # 題の行（1 セルしかない上の行で、その下にまだ行がある）
    i = r1
    title_rows = []
    while i <= r2 and i not in gaps and fill_in(i) <= 1 and (i - r1) < 4 and r2 - i >= 2 and width >= 2:
        title_rows.append(i)
        i += 1
    # 見出し行を探す
    hdr = None
    scan_lim = min(r2, i + 12)
    for ii in range(i, scan_lim + 1):
        if ii in gaps:
            continue
        f = fill_in(ii)
        if f < 2 and not (width == 1 and f == 1):
            continue
        # 見出しの行は、表の幅の半分以上を埋める（結合の続きも数える）。それ未満は題・注記の行
        if width >= 4 and sum(1 for k in K[ii][c1:c2 + 1] if k not in 'EW') < 0.5 * width:
            continue
        share = _row_text_share(K[ii], c1, c2, ctx, ii)
        if share >= 0.6 and not _row_is_data_like(ctx, ii, c1, c2, 0.3):
            hdr = ii
            break
        if _row_is_data_like(ctx, ii, c1, c2, 0.3):
            break
    hdr_rows = []
    if hdr is not None:
        hdr_rows = [hdr]
        nxt = hdr + 1
        while nxt <= r2 and nxt not in gaps and len(hdr_rows) < 5:
            only_cover = all(k in 'EWM' for k in K[nxt][c1:c2 + 1]) and 'M' in K[nxt][c1:c2 + 1] \
                and any(i2 >= nxt and i <= hdr_rows[-1] for (i, j), (i2, j2) in ctx['mtl'].items() if c1 <= j <= c2)
            if only_cover:
                hdr_rows.append(nxt)
                nxt += 1
            elif fill_in(nxt) >= 2 and _row_text_share(K[nxt], c1, c2, ctx, nxt) >= 0.8 and not _row_is_data_like(ctx, nxt, c1, c2) \
                    and any(_is_second_tier(ctx, hh, nxt, c1, c2) for hh in hdr_rows):
                hdr_rows.append(nxt)
                nxt += 1
            else:
                break
        for ii in range(i, hdr):
            if ii not in title_rows and ii not in gaps:
                title_rows.append(ii)
    if hdr is None:
        for ii in range(r1, r2 + 1):
            rows_role[ii] = 'U' if ii not in gaps else 'B'
        for ii in title_rows:
            rows_role[ii] = 'N'
        blk['kind'] = 'unknown'
        blk['unknown'].append({'item': 'header', 'reason': '見出しらしい行が見つからない（文字が6割以上の行が無い）'})
        return _finish_block(ctx, blk, bl, rows_role, [], title_rows, sorted(gaps))
    # 見出しの直後に、左端の 1 セルだけの文字の行（角の見出し）があり、その下が明細らしければ見出しの続き
    nx = hdr_rows[-1] + 1
    if nx + 1 <= r2 and nx not in gaps and fill_in(nx) == 1 and 'T' in K[nx][c1:c1 + 2] and _row_is_data_like(ctx, nx + 1, c1, c2, 0.3):
        hdr_rows.append(nx)
    for ii in title_rows:
        # 見出しの真上の、2 セル以上ある行は題か上段の見出しか決められない（推さない）
        rows_role[ii] = 'U' if (ii == hdr_rows[0] - 1 and (fill_in(ii) >= 2 or (fill_in(ii) == 1 and K[ii][c1] in 'EW' and width >= 4))) else 'N'
    for ii in hdr_rows:
        rows_role[ii] = 'H'
    body_start = hdr_rows[-1] + 1
    for ii in range(body_start, r2 + 1):
        if ii in gaps or all(k in 'EW' for k in K[ii][c1:c2 + 1]):
            rows_role[ii] = 'B'
            continue
        f = fill_in(ii)
        if f == 0:
            rows_role[ii] = 'U'          # 結合の続きだけの行（上の行の一部）: 明細とも空行とも言わない
            continue
        role = 'D'
        texts = _first_texts(ctx, ii, c1, c2, 3)
        label_total = False
        for (jj, tx) in texts[:3]:
            if _is_total_label(tx):
                label_total = True
                break
            t = nfkc(tx).replace(' ', '')
            if jj - c1 <= 1 and len(t) <= 10 and re.search(r'(合計|小計|総計|累計)$', t) and _t_cells(ctx, ii, c1, c2) <= 1:
                label_total = True
                break
            if jj - c1 <= 1 and 2 <= len(t) <= 8 and t.endswith('計') and not t.endswith(_NOT_TOTAL_KEI) and _t_cells(ctx, ii, c1, c2) <= 1:
                label_total = True
                break
            if len(t) <= 8 and re.search(r'(合計|小計|総計)$', t) and len(texts) == 1 and _t_cells(ctx, ii, c1, c2) == 1:
                label_total = True         # 合計の語だけの行（ラベルが左端から離れていてもよい）
                break
        nfx, nsum = _row_formula_total(ctx, ii, c1, c2)
        has_num = any(k in 'ND' for k in K[ii][c1:c2 + 1]) or nfx > 0
        if label_total and (has_num or f == 1):
            role = 'T'
        elif nfx >= 2 and nsum >= max(2, int(0.6 * nfx)) and f >= 2:
            role = 'T'
        elif f == 1 and K[ii][c1:c2 + 1].count('T') == 1 and width >= 3:
            role = 'G'
        elif texts and f <= 2 and _is_note_text(texts[0][1]):
            role = 'N'
        elif (texts and f <= 3 and texts[0][0] == c1 and re.search(r'[：:]$', nfkc(texts[0][1]).strip())
              and not any(k in 'ND' for k in K[ii][c1:c2 + 1]) and ii - 1 >= body_start and rows_role.get(ii - 1) == 'T'):
            role = 'N'          # 合計の次に「〇〇名：」で始まる行は、次のまとまりの題（表の外）
        rows_role[ii] = role
    # 値の直書きの小計: 左端に名前があり、続く「左端が空の行」の数の合計が、その行の数と一致する行は小計
    vals = ctx['vals']
    for ii in range(body_start, r2 - 1):
        if rows_role.get(ii) != 'D' or K[ii][c1] != 'T':
            continue
        kids = []
        k = ii + 1
        while k <= r2 and rows_role.get(k) == 'D' and K[k][c1] in 'EW':
            kids.append(k)
            k += 1
        if len(kids) < 2:
            continue
        hit = tried = 0
        for j in range(c1 + 1, c2 + 1):
            a = vals[ii][j]
            if isinstance(a, bool) or not isinstance(a, (int, float)) or a == 0:
                continue
            col = [vals[x][j] for x in kids]
            if any(isinstance(y, bool) or not isinstance(y, (int, float)) for y in col if y not in (None, '')):
                continue
            tried += 1
            if abs(sum(y for y in col if y not in (None, '')) - a) <= 1e-6 * max(1.0, abs(a)):
                hit += 1
        if tried >= 2 and hit >= 2 and hit >= 0.8 * tried:
            rows_role[ii] = 'T'
    # 表の途中でもう一度出てくる見出し（ページ替わりの繰り返し）: 先頭の語が最初の見出しと同じ行は見出しの続き
    h0 = next((nfkc(str(ctx['vals'][hdr_rows[0]][j])).replace(' ', '') for j in range(c1, c2 + 1)
               if K[hdr_rows[0]][j] == 'T'), '')
    if h0 and len(h0) >= 2:
        ii = body_start
        while ii <= r2:
            if rows_role.get(ii) in ('D', 'G', 'U') and ii not in gaps:
                t0 = next((nfkc(str(ctx['vals'][ii][j])).replace(' ', '') for j in range(c1, c2 + 1) if K[ii][j] == 'T'), '')
                same = sum(1 for j in range(c1, c2 + 1) if K[ii][j] == 'T' and K[hdr_rows[0]][j] == 'T'
                           and nfkc(str(ctx['vals'][ii][j])).replace(' ', '') == nfkc(str(ctx['vals'][hdr_rows[0]][j])).replace(' ', ''))
                if (t0 == h0 or same >= 3) and fill_in(ii) >= 2 and _row_text_share(K[ii], c1, c2, ctx, ii) >= 0.8                         and not _row_is_data_like(ctx, ii, c1, c2, 0.3):
                    rows_role[ii] = 'H'
                    for back in (ii - 1, ii - 2):
                        if back >= body_start and rows_role.get(back) == 'G' and fill_in(back) <= 1:
                            rows_role[back] = 'N'
                    k = ii + 1
                    while k <= min(r2, ii + 3) and k not in gaps and rows_role.get(k) in ('D', 'G', 'U')                             and (fill_in(k) == 0 or (_row_text_share(K[k], c1, c2, ctx, k) >= 0.8
                                                     and not _row_is_data_like(ctx, k, c1, c2, 0.3))):
                        rows_role[k] = 'H'
                        k += 1
                    ii = k
                    continue
            ii += 1
    # 合計の見出しが縦結合で下の行まで続くなら、その行も合計
    for ii in [x for x in range(body_start, r2 + 1) if rows_role.get(x) == 'T']:
        for (i, j), (i2, j2) in ctx['mtl'].items():
            if i == ii and c1 <= j <= c2 and i2 > i and K[i][j] == 'T':
                for x in range(i + 1, min(i2, r2) + 1):
                    if rows_role.get(x) in ('D', 'G'):
                        rows_role[x] = 'T'
    # 数の行が続く表の中の、文字だけで幅の半分以上を埋める行（見出しの続きか注釈）は明細と言い切らない
    dl = [x for x in range(body_start, r2 + 1) if rows_role.get(x) == 'D']
    n_data_like = sum(1 for x in dl if _row_is_data_like(ctx, x, c1, c2, 0.3))
    if width >= 4 and n_data_like >= 3:
        for x in dl:
            kx = K[x][c1:c2 + 1]
            if (not any(k in 'NDX' for k in kx) and not any((x, j) in ctx['fset'] for j in range(c1, c2 + 1))
                    and _row_text_share(K[x], c1, c2, ctx, x) >= 0.99 and fill_in(x) >= 0.5 * width
                    and not _row_is_data_like(ctx, x, c1, c2)):
                rows_role[x] = 'U'
    # 数だけの行が、ラベルのある行から一定の間隔で並ぶ（ラベルは先頭の 1 行だけ）= 1 件が複数行の帳票
    if not is_form:
        dl = [x for x in range(body_start, r2 + 1) if rows_role.get(x) == 'D']
        lab = [x for x in dl if 'T' in K[x][c1:c1 + 3]]
        if len(dl) >= 12 and 5 <= len(lab) <= 0.5 * len(dl):
            pos = [dl.index(x) for x in lab]
            diffs = [b - a for a, b in zip(pos, pos[1:])]
            if diffs and 2 <= max(set(diffs), key=diffs.count) <= 4 and diffs.count(max(set(diffs), key=diffs.count)) >= 0.8 * len(diffs):
                is_form = True
    last = None
    for ii in range(r2, body_start - 1, -1):
        if rows_role.get(ii) in ('D', 'T'):
            last = ii
            break
    if last is not None:
        for ii in range(last + 1, r2 + 1):
            if rows_role.get(ii) in ('G', 'D', 'T'):
                rows_role[ii] = 'N'
    else:
        for ii in range(body_start, r2 + 1):
            if rows_role.get(ii) == 'G':
                rows_role[ii] = 'N'
    if not any(v in ('D', 'T') for v in rows_role.values()) and not is_form:
        # 見出しの下に明細が 1 行も無い: 表ではない（項目と値の 1 行など）。推して埋めない
        for ii in range(r1, r2 + 1):
            if rows_role.get(ii) not in ('N', 'B'):
                rows_role[ii] = 'U'
        blk['kind'] = 'unknown'
        blk['unknown'].append({'item': 'rows', 'reason': '見出しらしい行の下に明細が無い（項目と値の 1 行などかもしれない）'})
        return _finish_block(ctx, blk, bl, rows_role, [], title_rows, sorted(gaps))
    if is_form:
        blk['kind'] = 'form'
        for ii, rl in list(rows_role.items()):
            if rl == 'D':
                rows_role[ii] = 'F'
        blk['info'] = ['帳票（1件が複数行）: 明細の行は F とし、件数は数えない']
    return _finish_block(ctx, blk, bl, rows_role, hdr_rows, title_rows, sorted(gaps))


def _is_long_or_note(ctx, i, c1, c2):
    for j in range(c1, c2 + 1):
        v = ctx['vals'][i][j]
        if isinstance(v, str) and (len(v) > 25 or _is_note_text(v)):
            return True
    return False


def _is_form(ctx, bl):
    """1 件が複数行にまたがる帳票か。縦結合が同じ高さで 3 つ以上並び、本文の行の大半を覆う。
    高さがまちまちの縦結合（グループ見出し）は帳票ではない。"""
    r1, r2, c1, c2 = bl['r1'], bl['r2'], bl['c1'], bl['c2']
    nrow = r2 - r1 + 1
    areas = [(i, j, i2) for (i, j), (i2, j2) in ctx['mtl'].items() if r1 <= i <= r2 and c1 <= j <= c2 and i2 > i]
    if len(areas) < 3 or nrow < 8:
        return False
    from collections import Counter
    hs = Counter((i2 - i + 1) for (i, j, i2) in areas)
    h, cnt = hs.most_common(1)[0]
    if cnt < 3 or cnt < 0.8 * len(areas) and cnt < 0.6 * len(areas):
        return False
    cov = set()
    for (i, j, i2) in areas:
        if i2 - i + 1 == h:
            cov.update(range(i, i2 + 1))
    return len(cov) >= 0.5 * nrow


def _is_second_tier(ctx, hrow, nrow_i, c1, c2):
    """hrow（上段）に対して nrow_i が見出しの下段か。結合の被覆・空欄の対応を根拠にする。"""
    K = ctx['K']
    for (i, j), (i2, j2) in ctx['mtl'].items():
        if i == hrow and c1 <= j <= c2 and j2 > j and i2 == i:
            return True
        if i == hrow and c1 <= j <= c2 and i2 >= nrow_i:
            return True
    gap_match = 0
    filled_up = filled_dn = 0
    for j in range(c1, c2 + 1):
        if K[hrow][j] in 'EWM' and K[nrow_i][j] == 'T':
            gap_match += 1
        if K[hrow][j] == 'T':
            filled_up += 1
        if K[nrow_i][j] == 'T':
            filled_dn += 1
    return (gap_match >= 1 and filled_up >= 1 and filled_dn > filled_up) or (gap_match >= 2 and filled_up >= 2)


# ----------------------------------------------------------------
# 列の顔・症状・安全の印
# ----------------------------------------------------------------
def _runs_cap(nums, cap=40):
    s = _runs(nums)
    parts = s.split(',')
    if len(parts) > cap:
        return ','.join(parts[:cap]) + f',…(+{len(parts) - cap})'
    return s


def _tl_of(ctx, i, j):
    """(i,j) を覆う結合の左上。結合でなければ自分。"""
    return ctx['mcover'].get((i, j), (i, j))


def _cell_text(ctx, i, j):
    ti, tj = _tl_of(ctx, i, j)
    v = ctx['vals'][ti][tj]
    return v if isinstance(v, str) else ('' if v is None else str(v))


def _runs_of_roles(rows_role, r0):
    out = []
    keys = sorted(rows_role)
    i = 0
    while i < len(keys):
        j = i
        while j + 1 < len(keys) and keys[j + 1] == keys[j] + 1 and rows_role[keys[j + 1]] == rows_role[keys[i]]:
            j += 1
        a, b = r0 + keys[i], r0 + keys[j]
        out.append({'rows': str(a) if a == b else f'{a}-{b}', 'role': rows_role[keys[i]]})
        i = j + 1
    return out


def _finish_block(ctx, blk, bl, rows_role, hdr_rows, title_rows, gaps):
    r0, c0 = ctx['r0'], ctx['c0']
    r1, r2, c1, c2 = bl['r1'], bl['r2'], bl['c1'], bl['c2']
    for ii in range(r1, r2 + 1):
        rows_role.setdefault(ii, 'B' if ii in gaps else 'U')
    blk['rows'] = _runs_of_roles(rows_role, r0)
    blk['header_rows'] = [r0 + h for h in hdr_rows]
    blk['title_rows'] = [r0 + h for h in title_rows]
    det = [i for i in range(r1, r2 + 1) if rows_role.get(i) == 'D']
    tot = [i for i in range(r1, r2 + 1) if rows_role.get(i) == 'T']
    blk['detail_rows'] = _runs(r0 + i for i in det)
    blk['total_rows'] = [r0 + i for i in tot]
    blk['blank_rows'] = _runs(r0 + i for i in range(r1, r2 + 1) if rows_role.get(i) == 'B')
    blk['group_rows'] = [r0 + i for i in range(r1, r2 + 1) if rows_role.get(i) == 'G']
    blk['note_rows'] = [r0 + i for i in range(r1, r2 + 1) if rows_role.get(i) == 'N' and i not in title_rows]
    symptoms = []
    blk['symptoms'] = symptoms
    safety = {}
    blk['safety'] = safety
    if blk['kind'] != 'table':
        _safety_common(ctx, blk, bl, rows_role, safety)
        return blk
    if bl.get('empty_cols'):
        symptoms.append({'code': 'S13', 'scope': 'table', 'cols': [col_letter(c0 + j) for j in bl['empty_cols']]})
    # 二段見出し
    if len(hdr_rows) >= 2:
        symptoms.append({'code': 'S16', 'scope': 'table', 'rows': _runs(r0 + h for h in hdr_rows)})
    # 結合の情報が一部だけなら言い切らない
    if 'merged' in ctx['unknown_sheet']:
        blk['unknown'].append({'item': 'merged', 'reason': ctx['unknown_sheet']['merged']})
    # 明細の間の空行
    if det:
        d_first, d_last = det[0], det[-1]
        inner_blank = [i for i in range(d_first, d_last + 1) if rows_role.get(i) == 'B']
        if inner_blank:
            symptoms.append({'code': 'S12', 'scope': 'table', 'rows': _runs(r0 + i for i in inner_blank)})
    # 列
    cols = []
    blk['columns'] = cols
    last_hdr = hdr_rows[-1] if hdr_rows else None
    for j in range(c1, c2 + 1):
        col = _analyze_column(ctx, blk, j, hdr_rows, det, tot, rows_role, symptoms, safety)
        if col is not None:
            cols.append(col)
    # 合計が明細の途中にある
    if det and tot:
        safety['subtotal_inside'] = any(det[0] < t < det[-1] for t in tot)      # 明細の間にある小計（先頭の総計・末尾の合計は途中ではない）
    safety['totals'] = [{'row': r0 + t} for t in tot]
    # 重複行
    if det:
        seen = {}
        dups = []
        vals = ctx['vals']
        for i in det:
            key = tuple(('' if v is None else (v.strip() if isinstance(v, str) else v)) for v in vals[i][c1:c2 + 1])
            if sum(1 for v in key if v != '') < 2:
                continue
            if key in seen:
                dups.append(i)
                dups.append(seen[key])
            else:
                seen[key] = i
        if dups:
            symptoms.append({'code': 'S11', 'scope': 'table', 'rows': _runs(r0 + i for i in set(dups))})
    # 結合（明細・合計の中）
    mrows = []
    mcols = set()
    for (i, j), (i2, j2) in ctx['mtl'].items():
        if r1 <= i <= r2 and c1 <= j <= c2 and rows_role.get(i) in ('D', 'T', 'G'):
            mrows.append(i)
            mcols.add(j)
    if mrows:
        symptoms.append({'code': 'S15', 'scope': 'table', 'rows': _runs_cap(r0 + i for i in mrows), 'cols': [col_letter(c0 + j) for j in sorted(mcols)]})
    # 合計の行が無い（数の列があるのに）
    if det and not tot and 3 <= len(det) <= 60 and not _time_axis(ctx, det, c1):
        numcols = [c for c in cols if c.get('face') in ('金額', '数量') and _additive_header(c.get('_hdr', ''))]
        if numcols:
            symptoms.append({'code': 'S30', 'scope': 'table', 'cols': [c['col'] for c in numcols]})
    _safety_common(ctx, blk, bl, rows_role, safety)
    for c in cols:
        c.pop('_hdr', None)
    return blk


def _safety_common(ctx, blk, bl, rows_role, safety):
    r0, c0 = ctx['r0'], ctx['c0']
    r1, r2, c1, c2 = bl['r1'], bl['r2'], bl['c1'], bl['c2']
    mg = [addr_of(r0 + i, c0 + j, r0 + i2, c0 + j2) for (i, j), (i2, j2) in ctx['mtl'].items()
          if r1 <= i <= r2 and c1 <= j <= c2]
    if mg:
        safety['merged'] = mg[:60]
        safety['merged_count'] = len(mg)
        # 60 件で切っても、見出しの結合と明細の結合の別は落とさない（棚当てが使う）
        hdr_rs = {r0 + i for i, rl in rows_role.items() if rl == 'H'}
        det_rs = {r0 + i for i, rl in rows_role.items() if rl in ('D', 'F', 'T')}
        mh, md, mrows = [], 0, set()
        for (i, j), (i2, j2) in ctx['mtl'].items():
            if r1 <= i <= r2 and c1 <= j <= c2:
                rr = set(range(r0 + i, r0 + i2 + 1))
                if rr & hdr_rs and not rr & det_rs:
                    mh.append(addr_of(r0 + i, c0 + j, r0 + i2, c0 + j2))
                if rr & det_rs:
                    md += 1
                    mrows |= rr
        safety['merged_in_header'] = mh[:30]
        safety['merged_in_detail'] = md
        safety['merged_detail_rows'] = ','.join(f'{a}-{z}' if a != z else str(a) for a, z in _to_pairs(sorted(mrows))) if mrows else ''
    hr = []
    for (a, b) in ctx['hidden_rows']:
        lo, hi = max(a, r0 + r1), min(b, r0 + r2)
        if lo <= hi:
            hr.append(f'{lo}' if lo == hi else f'{lo}-{hi}')
    hc = []
    for (a, b) in ctx['hidden_cols']:
        lo, hi = max(a, c0 + c1), min(b, c0 + c2)
        if lo <= hi:
            hc.append(col_letter(lo) if lo == hi else f'{col_letter(lo)}-{col_letter(hi)}')
    if hr:
        safety['hidden_rows'] = hr
    if hc:
        safety['hidden_cols'] = hc
    if hr or hc:
        blk['symptoms'].append({'code': 'S35', 'scope': 'table', 'rows': ','.join(hr), 'cols': hc})
    if blk['note_rows']:
        safety['notes'] = blk['note_rows'][:40]
    # 他の式から参照されている範囲（この塊に重なる）
    refd = {}
    if 'refs_all' not in ctx:
        ctx['refs_all'] = [(a1, b1, a2, b2) for f in ctx['formulas'].values()
                           for (a1, b1, a2, b2, ext) in _refs_in_formula(f) if not ext]
    for (a1, b1, a2, b2) in ctx['refs_all']:
        if a1 - r0 <= r2 and a2 - r0 >= r1 and b1 - c0 <= c2 and b2 - c0 >= c1:
            key = (a1, b1, a2, b2)
            refd[key] = refd.get(key, 0) + 1
    if refd:
        top = sorted(refd.items(), key=lambda kv: -kv[1])[:12]
        safety['referenced'] = [{'range': addr_of(*k), 'by': n} for k, n in top]
        safety['referenced_total'] = len(refd)


def _hdr_of(ctx, j, hdr_rows):
    parts = []
    for h in hdr_rows:
        t = _cell_text(ctx, h, j)
        t = t.strip() if t else ''
        if t and (not parts or parts[-1] != t):
            parts.append(t)
    return parts


def _classify(ctx, hdr_parts, s):
    """列の顔を決める。s は値の集計 dict。戻り (face, reason)。"""
    ht = nfkc(' '.join(hdr_parts)) if hdr_parts else ''
    leaf = nfkc(hdr_parts[-1]) if hdr_parts else ''
    n = s['n']
    if n == 0:
        return '空', None
    fr = s['n_formula'] / n
    if fr >= 0.6:
        return '式', None
    dateish = (s['n_date'] + s['n_tdate']) / n
    numeric = (s['n_num'] + s['n_tnum'] + s['n_date'] * 0) / n
    textr = s['n_text'] / n
    if _HDR_NOTE.search(ht) and textr >= 0.5:
        return '文字', None
    if dateish >= 0.6 or (_HDR_DATE.search(leaf or ht) and dateish >= 0.3):
        return '日付', None
    # 8 桁の日付（全部 8 桁の日付になる数）
    if s['n_d8'] and s['n_d8'] / n >= 0.9 and _HDR_DATE.search(leaf or ht):
        return '日付', None
    if _HDR_POST.search(leaf or ht) and (s['n_post'] / n >= 0.5 or numeric >= 0.5):
        return '郵便', None
    if _HDR_PHONE.search(leaf or ht) and (s['n_phone'] / n >= 0.5 or numeric >= 0.5):
        return '電話', None
    if s['n_post'] / n >= 0.8 and not numeric > 0.99:
        return '郵便', None
    if s['n_phone'] / n >= 0.8:
        return '電話', None
    if _HDR_KANA.search(leaf or ht) and textr >= 0.3:
        return 'カナ', None
    if s['n_kana'] / n >= 0.8 and textr >= 0.8:
        return 'カナ', None
    if numeric >= 0.6:
        h = leaf or ht
        if s['n_year'] and s['n_year'] >= 0.9 * n and re.search(r'年|年次|年度|西暦', h):
            return '日付', '年だけ（西暦）の数'
        if (s['n_small'] + s['n_tsmall']) >= 0.75 * n and re.fullmatch(r'年|年度', re.sub(r'\s', '', h) or 'x'):
            return '日付', '年だけ（和暦の年数）の数'
        if s['n_tnum'] and s['n_tsmall'] >= 0.9 * s['n_tnum'] and s['n_tnum'] >= 0.6 * n and not _HDR_MONEY.search(h) and not _HDR_QTY.search(h):
            return '番号', '2桁以下の整数の文字（区分の番号）'
        if _HDR_ID.search(h):
            return '番号', None
        if _HDR_CODE.search(h):
            return 'コード', None
        mon = bool(_HDR_MONEY.search(h)) or bool(_HDR_MONEY.search(ht))
        qty = bool(_HDR_QTY.search(h)) or bool(_HDR_QTY.search(ht))
        if mon and not re.search(r'人口|世帯|数量|件数|人数|面積|率|割合', h):
            return '金額', None
        if qty:
            return '数量', None
        if mon:
            return '金額', None
        if s['seq']:
            return '番号', None
        if s['n_yen']:
            return '金額', None
        return '数量', '金額の語（円・額・費など）が無い数の列は数量とみなす'
    if textr >= 0.5:
        h = leaf or ht
        if _HDR_ADDR.search(h) or s['n_addr'] / n >= 0.5:
            return '住所', None
        if _HDR_NAME.search(h):
            return '氏名', None
        if _HDR_MAIL.search(h):
            return '文字', None
        if _HDR_ID.search(h):
            return '番号', None
        if _HDR_CODE.search(h):
            return 'コード', None
        return '文字', None
    return 'unknown', '文字と数が混ざる列'


_ADDR_RE = re.compile(r'(?:都|道|府|県).{1,10}(?:市|区|町|村|郡)|(?:市|区|町|村).{0,12}\d')


_YEARISH_RE = re.compile(r'^(?:令和|平成|昭和|[RHS])?\s*(?:元|\d{1,4})\s*年?度?\s*$')
_RATE_HDR = re.compile(r'率|比|指数|割合|平均|単価|増減|変化|%|％|構成|偏差|倍|順位|‰|パーミル')


def _time_axis(ctx, det, c1):
    """行の見出し（左端の列）が年・年度の並びか（時系列の表は、縦に足しても意味がない）。"""
    hits = tot = 0
    for i in det:
        for j in range(c1, min(c1 + 2, ctx['nc'])):
            v = ctx['vals'][i][j]
            if isinstance(v, str) and v.strip():
                tot += 1
                if _YEARISH_RE.match(nfkc(v).strip()):
                    hits += 1
                break
            if isinstance(v, float) and v == int(v) and 1900 <= v <= 2100:
                tot += 1
                hits += 1
                break
            if kind_of(v) == 'D':
                tot += 1
                hits += 1
                break
    return tot > 0 and hits >= 0.6 * tot


_ADD_HDR = re.compile(r'数|人|件|額|費|円|面積|人口|世帯|台|戸|量|収入|支出|税|金|料|床|延|軒')


def _additive_header(h):
    t = nfkc(h or '')
    return bool(t) and bool(_ADD_HDR.search(t)) and not _RATE_HDR.search(t) and not re.fullmatch(r'[\s\d年度月中~\-～（）()]*', t)


def _analyze_column(ctx, blk, j, hdr_rows, det, tot, rows_role, symptoms, safety):
    r0, c0 = ctx['r0'], ctx['c0']
    K, vals, fset = ctx['K'], ctx['vals'], ctx['fset']
    nv = ctx['nv']
    hdr_parts = _hdr_of(ctx, j, hdr_rows)
    letter = col_letter(c0 + j)
    s = {'n': 0, 'n_text': 0, 'n_num': 0, 'n_date': 0, 'n_tnum': 0, 'n_tdate': 0, 'n_formula': 0,
         'n_post': 0, 'n_phone': 0, 'n_kana': 0, 'n_d8': 0, 'n_addr': 0, 'n_yen': 0, 'n_year': 0, 'n_small': 0, 'n_tsmall': 0, 'seq': False}
    rows_by = {k: [] for k in ('S01', 'S02', 'S03', 'S04', 'S05', 'S09', 'S14', 'S43', 'S44', 'S17', 'S18', 'S29', 'S07', 'S08', 'S06', 'S27')}
    texts = []
    nums = []
    blank_rows = []
    errs = []
    lead_depths = {}
    lead_rows = {}
    paren = []
    tri = []
    fw_sym = 0
    n_det = len(det)
    stride = 1
    if n_det > 4000:
        stride = max(1, n_det // 4000)
    det_iter = det[::stride] if stride > 1 else det
    for i in det_iter:
        k = K[i][j]
        v = vals[i][j]
        if (i, j) in fset:
            s['n_formula'] += 1
        if k == 'T':
            s['n'] += 1
            s['n_text'] += 1
            t = v
            ts = t.strip()
            if ts != t:
                lead = len(t) - len(t.lstrip())
                if lead and (i, j) not in fset:                 # 式のセルは棚が触らない（結果の文字の空白は症状にしない）
                    lead_depths[lead] = lead_depths.get(lead, 0) + 1
                    lead_rows.setdefault(lead, []).append(i)
                if len(t) - len(t.rstrip()) and (i, j) not in fset:
                    rows_by['S03'].append(i)
            texts.append((i, t))
            tn = nfkc(ts).replace(' ', '')
            if _is_text_number_fast(tn):
                s['n_tnum'] += 1
                if len(tn) <= 2:
                    s['n_tsmall'] += 1
                s['n_text'] -= 1
                if tn.startswith(('¥', '￥')) or tn.endswith('円'):
                    s['n_yen'] += 1
                if (i, j) not in fset:
                    rows_by['S02'].append(i)
            elif is_text_date(ts):
                s['n_tdate'] += 1
                s['n_text'] -= 1
                if (i, j) not in fset:
                    rows_by['S01'].append(i)
            elif _PAREN_NUM_RE.match(ts):
                paren.append(i)
            elif _TRI_NUM_RE.match(ts):
                tri.append(i)
            if _UNIT_RE.match(tn):
                rows_by['S43'].append(i)
            if _PERIOD_RE.search(ts):
                rows_by['S44'].append(i)
            if _POST_RE.match(tn):
                s['n_post'] += 1
            if _PHONE_RE.match(tn):
                s['n_phone'] += 1
            if _KANA_RE.match(ts):
                s['n_kana'] += 1
            if _ADDR_RE.search(ts):
                s['n_addr'] += 1
            if ts in ('－', '―', '─', '‐', '〃', '…', '・', '△') or ts.startswith('－'):
                fw_sym += 1
        elif k == 'N':
            s['n'] += 1
            s['n_num'] += 1
            nums.append((i, v))
        elif k == 'D':
            s['n'] += 1
            s['n_date'] += 1
        elif k == 'B':
            s['n'] += 1
            s['n_text'] += 1
        elif k == 'X':
            s['n'] += 1
            errs.append((i, v))
            if (i, j) in fset:
                if v == -2146826281:
                    rows_by['S18'].append(i)
                elif v == -2146826243:
                    rows_by['S29'].append(i)
        elif k == 'W':
            rows_by['S04'].append(i)
            blank_rows.append(i)
        else:
            blank_rows.append(i)
    # 8 桁の数が全部日付の形（列が全部そうなら）
    if nums:
        d8 = 0
        for (i, v) in nums:
            if isinstance(v, float) and v == int(v) and 10000000 <= v <= 99999999:
                y, m, d = int(v) // 10000, (int(v) // 100) % 100, int(v) % 100
                if 1900 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31:
                    d8 += 1
        s['n_d8'] = d8
        s['n_year'] = sum(1 for (_i, v) in nums if isinstance(v, float) and v == int(v) and 1900 <= v <= 2100)
        s['n_small'] = sum(1 for (_i, v) in nums if isinstance(v, float) and v == int(v) and 1 <= v <= 99)
        ints = [int(v) for (_i, v) in nums if isinstance(v, float) and v == int(v)]
        if len(ints) == len(nums) and len(ints) >= 3:
            diffs = [b - a for a, b in zip(ints, ints[1:])]
            s['seq'] = diffs.count(1) >= 0.8 * len(diffs)
    face, why = _classify(ctx, hdr_parts, s)
    # --- 症状（面ごと）
    nn = s['n']
    ht = nfkc(' '.join(hdr_parts)) if hdr_parts else ''
    # インデント（階層）
    indent = None
    if lead_depths:
        pos_depths = sorted(lead_depths)
        n_lead = sum(lead_depths.values())
        if (len(pos_depths) >= 2 and n_lead >= 3) or (n_lead >= 3 and n_lead >= 0.3 * max(1, s['n_text']) and any(
                (vals[i][j] or '').startswith('　') for i in lead_rows.get(pos_depths[0], [])[:1])):
            indent = pos_depths
            safety.setdefault('indent', []).append({'col': letter, 'depths': pos_depths})
        else:
            for d, rs in lead_rows.items():
                rows_by['S03'].extend(rs)
    # 自由記述の列（備考・メモ・要旨）は直さない（症状は報告するが、棚は撃たない）
    if hdr_parts and _HDR_NOTE.search(ht) and s['n_text'] >= 0.5 * max(1, s['n']):
        safety.setdefault('free_text', []).append(letter)
    # 全角記号に意味がある列
    if s['n_text'] and fw_sym >= 3 and fw_sym >= 0.3 * max(1, s['n_text'] + 0):
        safety.setdefault('fullwidth_symbol', []).append({'col': letter, 'n': fw_sym})
    # S01 文字の日付
    sy = []
    if rows_by['S01'] and (s['n_date'] >= 1 or s['n_tdate'] >= 0.6 * max(1, nn) or _HDR_DATE.search(ht or '')):
        sy.append(('S01', rows_by['S01']))
    # 8 桁数は日付の形で列が全部そうなら
    if s['n_d8'] and nn and s['n_d8'] >= 0.9 * nn and _HDR_DATE.search(ht or '') and not s['n_date']:
        sy.append(('S01', [i for (i, v) in nums]))
    # S02 文字の数字
    if rows_by['S02'] and face not in ('コード', '電話', '郵便', '番号', '日付') and (s['n_num'] + s['n_tnum']) >= 0.6 * max(1, nn) \
            and not any(_LEADZERO_RE.match(nfkc(vals[i][j].strip())) for i in rows_by['S02'] if isinstance(vals[i][j], str)):
        sy.append(('S02', rows_by['S02']))
    # S03 空白
    if rows_by['S03']:
        sy.append(('S03', sorted(set(rows_by['S03']))))
    # S04 空白だけのセル
    if rows_by['S04']:
        sy.append(('S04', rows_by['S04']))
    # S05 全角半角の混在
    if texts and face not in ('式',):
        fw = hw = kana_h = other = 0
        fw_rows = []
        for (i, t) in texts:
            a = bool(_FW_ALNUM.search(t))
            b = bool(_HW_ALNUM.search(t))
            hk = bool(_HW_KANA.search(t))
            if a:
                fw += 1
                fw_rows.append(i)
            if b:
                hw += 1
            if hk:
                kana_h += 1
        mixed = fw >= 1 and hw >= 1
        hwk = kana_h >= 1 and kana_h < len(texts)
        if (mixed or hwk) and not (fw_sym >= 3):
            sy.append(('S05', [i for (i, t) in texts if _FW_ALNUM.search(t) or _HW_KANA.search(t)]))
    # S06 氏名の空白の不揃い
    if face == '氏名' and texts:
        sp_half = sp_full = sp_none = 0
        for (i, t) in texts:
            ts = t.strip()
            if len(ts) < 3 or _ADDR_RE.search(ts):
                continue
            if ' ' in ts:
                sp_half += 1
            elif '　' in ts:
                sp_full += 1
            else:
                sp_none += 1
        if sum(1 for x in (sp_half, sp_full, sp_none) if x) >= 2:
            sy.append(('S06', [i for (i, t) in texts]))
    # S07 住所・郵便の形
    if face == '郵便':
        bad = []
        for i in det:
            v = vals[i][j]
            k = K[i][j]
            if k == 'T':
                tn = nfkc(v.strip())
                if not re.match(r'^\d{3}-\d{4}$', tn):
                    bad.append(i)
            elif k == 'N':
                bad.append(i)
        if bad:
            sy.append(('S07', bad))
    elif face == '住所':
        bad = [i for (i, t) in texts if _FW_ALNUM.search(t) or re.search(r'[－ー−―]', t) and re.search(r'\d', t)]
        if bad and any(not (_FW_ALNUM.search(t) or re.search(r'[－ー−―]', t)) for (i, t) in texts):
            sy.append(('S07', bad))
    # S08 電話の書き方
    if face == '電話' and texts:
        styles = set()
        for (i, t) in texts:
            tn = t.strip()
            styles.add(('hy' if '-' in tn or '－' in tn else 'nohy', 'fw' if _FW_ALNUM.search(tn) else 'hw', 'paren' if '(' in tn or '（' in tn else ''))
        if len(styles) >= 2:
            sy.append(('S08', [i for (i, t) in texts]))
    # S09 先頭ゼロ落ち
    if face in ('コード', '番号', '郵便', '電話') or (hdr_parts and (_HDR_CODE.search(ht) or _HDR_POST.search(ht) or _HDR_PHONE.search(ht))):
        zl = [i for (i, t) in texts if _LEADZERO_RE.match(nfkc(t.strip()))]
        if zl and nums:
            sy.append(('S09', [i for (i, v) in nums]))
        elif face in ('電話',) and nums:
            sy.append(('S09', [i for (i, v) in nums]))
        elif face == '郵便' and nums:
            sy.append(('S09', [i for (i, v) in nums]))
    # S10 連番の乱れ
    if face == '番号' and nums and len(nums) >= 5 and _HDR_ID.search(nfkc(' '.join(hdr_parts[-1:]) if hdr_parts else '')):
        ints = [(i, int(v)) for (i, v) in nums if isinstance(v, float) and v == int(v)]
        diffs = [b[1] - a[1] for a, b in zip(ints, ints[1:])]
        if diffs and diffs.count(1) >= 0.7 * len(diffs) and any(d != 1 for d in diffs):
            sy.append(('S10', [b[0] for a, b, d in zip(ints, ints[1:], diffs) if d != 1]))
    # S14 上の値の繰り返しが省略
    if (face in ('文字', '氏名', 'コード', 'unknown') or not face) and n_det >= 4 and not _HDR_NOTE.search(ht):
        nb = [i for i in det if K[i][j] == 'M']
        # 結合で表している場合は S15 側
        emp = [i for i in det if K[i][j] in 'EW']
        # 空欄が末尾に固まるだけなら省略ではない（空欄のあとに値が続くこと）
        interleaved = bool(emp) and any(K[i][j] not in 'EWM' for i in det if i > emp[0])
        if emp and interleaved and K[det[0]][j] == 'T' and len(emp) >= 0.2 * n_det and len(emp) <= 0.9 * n_det and not nb:
            # 空欄の行で他の列が埋まっているか
            ok = 0
            for i in emp:
                ff = sum(1 for k in K[i][ctx_c1(blk, ctx):ctx_c2(blk, ctx) + 1] if k not in 'EWM')
                if ff >= 2:
                    ok += 1
            if ok >= 0.8 * len(emp):
                sy.append(('S14', emp))
    # S17 式の途切れ・S18・S29
    if s['n_formula'] >= 3 and s['n_formula'] >= 0.6 * max(1, len(det)):
        wd = ctx_c2(blk, ctx) - ctx_c1(blk, ctx) + 1
        brk = [i for i in det if (i, j) not in fset and K[i][j] != 'M'
               and sum(1 for k in K[i][ctx_c1(blk, ctx):ctx_c2(blk, ctx) + 1] if k not in 'EWM') >= max(2, 0.3 * wd)]
        if brk:
            sy.append(('S17', brk))
    if rows_by['S18']:
        sy.append(('S18', rows_by['S18']))
    if rows_by['S29']:
        sy.append(('S29', rows_by['S29']))
    # S27 ふりがなの空欄（左隣が氏名で埋まっている）
    if face == 'カナ' and blank_rows:
        left = j - 1
        if left >= 0:
            bad = [i for i in det if K[i][j] in 'EW' and K[i][left] == 'T']
            if bad:
                sy.append(('S27', bad))
    # S43 単位付き・S44 期間
    if rows_by['S43'] and (s['n_num'] + s['n_tnum'] + len(rows_by['S43'])) >= 0.5 * max(1, nn):
        sy.append(('S43', rows_by['S43']))
    if rows_by['S44'] and (s['n_date'] + s['n_tdate']) >= 1:
        sy.append(('S44', rows_by['S44']))
    # 式の読み（S19 S20 S21）
    _formula_symptoms(ctx, j, det, tot, sy, rows_role)
    # 括弧・△▲の意味
    if paren or tri:
        pm = {'col': letter, 'n': len(paren) + len(tri)}
        if paren:
            if ctx['inner_note']:
                pm['meaning'] = 'inner_count'
            else:
                pm['meaning'] = 'unknown'
                pm['reason'] = '括弧が内数かマイナスか、注の文言から決められない（注に「内数」等が無い）'
            pm['kind'] = 'paren'
        else:
            pm['meaning'] = 'negative_marker'
            pm['kind'] = 'tri'
        safety.setdefault('paren', []).append(pm)
    col = {'col': letter, 'face': face, '_hdr': ' '.join(hdr_parts)}
    if why:
        if face == 'unknown':
            col['unknown'] = why
        else:
            col['basis'] = why          # 決めた根拠（unknown ではない）
    if face == '式':
        vf = _value_face(ctx, hdr_parts, s)
        if vf:
            col['value_face'] = vf
    if not nv and hdr_parts:
        col['header'] = ' / '.join(hdr_parts)[:40]
    if sy:
        col['symptoms'] = []
        for code, rows in sy:
            ent = {'code': code, 'n': len(set(rows)), 'rows': _runs_cap(r0 + i for i in set(rows))}
            col['symptoms'].append(ent)
            symptoms.append({'code': code, 'scope': 'column', 'col': letter, 'n': ent['n'], 'rows': ent['rows']})
    if indent:
        col['indent'] = True
    return col


def ctx_c1(blk, ctx):
    return parse_addr(blk['range'])[1] - ctx['c0']


def ctx_c2(blk, ctx):
    return parse_addr(blk['range'])[3] - ctx['c0']


_POST_RE = re.compile(r'^〒?\d{3}-?\d{4}$')
_PHONE_RE = re.compile(r'^\(?0\d{1,4}\)?[-ー]?\d{1,4}[-ー]?\d{3,4}$')
_LEADZERO_RE = re.compile(r'^0\d+$')


def _is_text_number_fast(tn):
    if not tn or len(tn) > 24:
        return False
    c = tn[0]
    if not (c.isdigit() or c in '+-−△▲¥￥\\'):
        return False
    return bool(_NUMTXT_RE.match(tn) or _YEN_RE.match(tn))


def _value_face(ctx, hdr_parts, s):
    ht = nfkc(' '.join(hdr_parts)) if hdr_parts else ''
    if _HDR_MONEY.search(ht):
        return '金額'
    if _HDR_QTY.search(ht):
        return '数量'
    return None


def _formula_symptoms(ctx, j, det, tot, sy, rows_role):
    r0, c0 = ctx['r0'], ctx['c0']
    fm = ctx['formulas']
    if not fm:
        return
    col_abs = c0 + j
    # S21: VLOOKUP の列番号が直書き
    lit = []
    rel = []
    for i in det:
        f = fm.get((r0 + i, col_abs))
        if not f:
            continue
        if _VLOOKUP_LIT_RE.search(f):
            lit.append(i)
        m = re.search(r'(?:VLOOKUP|HLOOKUP|SUMIF|COUNTIF|INDEX|MATCH)\s*\([^,()]*,\s*(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)', f, re.IGNORECASE)
        if m:
            rng = m.group(1)
            if '$' not in rng:
                rel.append(i)
    if lit:
        sy.append(('S21', lit))
    if len(rel) >= 3:
        # 下へ写して範囲の$が抜けている（範囲が行ごとにずれている）
        starts = set()
        for i in rel:
            f = fm.get((r0 + i, col_abs))
            m = re.search(r'(?:VLOOKUP|HLOOKUP|SUMIF|COUNTIF|INDEX|MATCH)\s*\([^,()]*,\s*(\$?[A-Z]{1,3}\$?\d+):', f, re.IGNORECASE)
            if m:
                starts.add(m.group(1))
        if len(starts) >= 2:
            sy.append(('S20', rel))
    # S19: 合計の SUM が明細の最終行まで届いていない（合計の行が 1 つだけの表。区ごとの小計がある表の総計は小計を足すので、明細の最終行まで届かなくてよい）
    if tot and det and len(tot) == 1:
        short = []
        d_last_abs = r0 + det[-1]
        for t in tot:
            f = fm.get((r0 + t, col_abs))
            if f and _SUM_RE.match(f):
                refs = [(a1, b1, a2, b2) for (a1, b1, a2, b2, ext) in _refs_in_formula(f) if not ext and b1 <= col_abs <= b2]
                det_above = [i for i in det if i < t]
                if refs and det_above:
                    last_above = r0 + det_above[-1]
                    if all(a2 < last_above for (a1, b1, a2, b2) in refs):
                        short.append(t)
        if short:
            sy.append(('S19', short))


# ----------------------------------------------------------------
# シート全体
# ----------------------------------------------------------------
def _sheet_level(ctx, out, sheet, blocks, content_rows):
    r0, c0, nr, nc = ctx['r0'], ctx['c0'], ctx['nr'], ctx['nc']
    K = ctx['K']
    sy = out['sheet_symptoms']
    # 値のある範囲（S25 ゴースト）
    if content_rows:
        top, bot = content_rows[0], content_rows[-1]
        cmin, cmax = nc, -1
        for i in content_rows:
            kr = K[i]
            for j in range(nc):
                if kr[j] not in 'EWM':
                    if j < cmin:
                        cmin = j
                    break
            for j in range(nc - 1, -1, -1):
                if kr[j] not in 'EWM':
                    if j > cmax:
                        cmax = j
                    break
        gh_rows = top + (nr - 1 - bot)
        gh_cols = cmin + (nc - 1 - cmax)
        out['content'] = addr_of(r0 + top, c0 + cmin, r0 + bot, c0 + cmax)
        if gh_rows >= 1 or gh_cols >= 1:
            sy.append({'code': 'S25', 'scope': 'sheet', 'ghost_rows': gh_rows, 'ghost_cols': gh_cols})
    else:
        out['content'] = ''
        if nr * nc > 1:
            sy.append({'code': 'S25', 'scope': 'sheet', 'ghost_rows': nr, 'ghost_cols': nc})
    # 外部参照
    ext = [r for (r, c), f in ctx['formulas'].items() if _EXT_RE.search(f)]
    book = sheet.get('book') or {}
    if ext or book.get('external_links'):
        ent = {'code': 'S23', 'scope': 'book' if not ext else 'sheet', 'links': book.get('external_links', 0)}
        if ext:
            ent['rows'] = _runs_cap(ext)
        sy.append(ent)
    if book.get('broken_names'):
        sy.append({'code': 'S24', 'scope': 'book', 'n': len(book['broken_names'])})
    # エラー値（病巣）
    errs = {}
    for i in range(nr):
        kr = K[i]
        if 'X' in kr:
            for j in range(nc):
                if kr[j] == 'X':
                    code = ctx['vals'][i][j]
                    nm = _ERR_CODES.get(code, 'エラー')
                    errs.setdefault(nm, []).append(addr_of(r0 + i, c0 + j))
    if errs:
        out['error_cells'] = {k: {'n': len(v), 'first': v[:5]} for k, v in errs.items()}
    # テーブル（ListObject）の外のデータ
    for la in (sheet.get('objects') or {}).get('listobjects', []) or []:
        p = parse_addr(la)
        if not p:
            continue
        a1, b1, a2, b2 = p
        below = a2 + 1 - r0
        right = b2 + 1 - c0
        hit = False
        if 0 <= below < nr:
            hit = any(K[below][j] not in 'EWM' for j in range(max(0, b1 - c0), min(nc, b2 - c0 + 1)))
        if not hit and 0 <= right < nc:
            hit = any(K[i][right] not in 'EWM' for i in range(max(0, a1 - r0), min(nr, a2 - r0 + 1)))
        if hit:
            sy.append({'code': 'S36', 'scope': 'sheet', 'table': la})
    out['objects'] = sheet.get('objects') or {}
    if book:
        out['book'] = {k: v for k, v in book.items() if k != 'broken_names'}
    out['hidden'] = {'rows': [f'{a}-{b}' if a != b else str(a) for a, b in ctx['hidden_rows']],
                     'cols': [f'{col_letter(a)}-{col_letter(b)}' if a != b else col_letter(a) for a, b in ctx['hidden_cols']]}
    # 塊ごとの症状も平らに集める
    flat = []
    for b in out['blocks']:
        for s in b.get('symptoms', []):
            e = dict(s)
            e['block'] = b['id']
            flat.append(e)
    out['symptoms'] = flat + sy
    for b in out['blocks']:
        for u in b.get('unknown', []):
            out['unknown'].append({'block': b['id'], **u})
        for c in b.get('columns', []) or []:
            if c.get('unknown'):
                out['unknown'].append({'block': b['id'], 'item': f"face {c['col']}", 'reason': c['unknown']})


def summarize(m):
    """地図 → 人が読む短い要約（値は出さない）。"""
    lines = [f"レントゲン: {m.get('sheet')}  使用範囲 {m.get('used')}  ({m.get('sec', 0) * 1000:.0f} ms)"]
    for b in m['blocks']:
        lines.append(f"- {b['id']} {b['range']} [{b['kind']}] 見出し {b.get('header_rows') or '-'} 明細 {b.get('detail_rows') or '-'} 合計 {b.get('total_rows') or '-'}")
        faces = ' '.join(f"{c['col']}={c['face']}" for c in b.get('columns', []) or [])
        if faces:
            lines.append(f"    列の顔: {faces}")
        sy = {}
        for s in b.get('symptoms', []):
            sy.setdefault(s['code'], []).append(s.get('col') or '表')
        if sy:
            lines.append('    症状: ' + ' '.join(f"{k}({','.join(v)})" for k, v in sorted(sy.items())))
    ss = [s for s in m.get('sheet_symptoms', [])]
    if ss:
        lines.append('シート: ' + ' '.join(s['code'] for s in ss))
    nbs = [n for n in (m.get('neighbors') or []) if n.get('kind') in ('master', 'comparable')]
    if nbs:
        lines.append('ほかのシート: ' + ' '.join(f"{n['sheet']}={'対応表' if n['kind'] == 'master' else '比べる相手'}" for n in nbs))
    if m.get('unknown'):
        lines.append(f"unknown {len(m['unknown'])} 件: " + ' / '.join(f"{u.get('item')}: {u.get('reason')}" for u in m['unknown'][:5]))
    return '\n'.join(lines)


# ----------------------------------------------------------------
# 入口（vba("xray")）
# ----------------------------------------------------------------
def xray_sheet(ws, wb=None, nv=False, neighbors=False, full_merges=False):
    """ws → 地図。読む（COM 範囲単位）と判定（純 Python）の秒数を地図に添える。"""
    t0 = time.perf_counter()
    sh = read_sheet(ws, wb, neighbors=neighbors, merge_budget=None if full_merges else 0.08, merge_cap=20000 if full_merges else 400)
    t1 = time.perf_counter()
    m = analyze(sh, nv=nv)
    m['read_sec'] = round(t1 - t0, 4)
    m['total_sec'] = round(time.perf_counter() - t0, 4)
    return m


def cmd_xray(args):
    """xray [excel_file] [--sheet 名] [--json] [--svg [出力.svg]]: シートのレントゲン（地図）。値は出さない。"""
    import vbam_core as vc
    import vbam_structure as vs
    import os
    target_file, rest = vc.parse_target_and_rest(args.posargs)
    xl, wb = vc.get_workbook(target_file, readonly=True)
    only = getattr(args, 'sheet_opt', None) or (rest[0] if rest else None)
    ws = wb.Worksheets(only) if only else wb.ActiveSheet
    nv = vs.no_values_mode()
    m = xray_sheet(ws, wb, nv=nv, neighbors=bool(getattr(args, 'neighbors', False)), full_merges=bool(getattr(args, 'full_merges', False)))
    if getattr(args, 'json', False):
        print(json.dumps(m, ensure_ascii=False, default=str))
    else:
        print(summarize(m))
        print(f"（読み {m['read_sec'] * 1000:.0f} ms・判定 {m['sec'] * 1000:.0f} ms）")
    svg_out = getattr(args, 'svg', None)
    if svg_out:
        here = os.path.dirname(os.path.abspath(__file__))
        path = os.path.abspath(svg_out if isinstance(svg_out, str) and svg_out != '1' and svg_out != 'True' else os.path.join(here, '_last_xray.svg'))
        try:
            from vbam_xray_svg import render_svg
            sh = read_sheet(ws, wb)
            with open(path, 'w', encoding='utf-8') as f:
                f.write(render_svg(m, sh))
            print(f'写真: {path}')
        except Exception as e:
            print(f'写真を描けません: {type(e).__name__}: {e}')
    return True
