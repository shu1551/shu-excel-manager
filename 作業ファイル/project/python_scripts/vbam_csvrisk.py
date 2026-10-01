# -*- coding: utf-8 -*-
"""vbam_csvrisk.py — CSV を Excel で開くと壊れる列を先に見つけ、その列だけ文字列で取り込む（2026-10-01）。

Excel は CSV・テキストを開くとき、列の中身を勝手に読み替える。**一度読み替わると、保存した時点で元に戻らない**。
  先頭の 0          007 → 7（郵便番号・社員番号・電話番号・コード）
  16 桁以上の数字   1234567890123456789 → 1234567890123450000（15 桁より先が 0 になる・ID・カード番号・JAN）
  12〜15 桁の数字   123456789012 → 1.23457E+11 と表示（値は残るが見た目が壊れる）
  日付に読める短い形 1-2 → 2 月 1 日（型番・部屋番号・分類コード）、JAN1 → 1 月 1 日（遺伝子名・コード）
  指数に読める形    12E5 → 1200000（型番・ロット番号）
  先頭の +          +81 → 81
  = で始まる         式として計算される
オフィス田中の「CSV を Excel 経由で開くときの自動データ変換」。道具は open で CSV を開くとき、壊れる列だけ文字列にして開く。

  scan_csv(path)                    → 壊れる列の一覧
  risky_field_info(scan)            → Workbooks.OpenText の FieldInfo（壊れる列だけ文字列 2・ほかは標準 1）
  describe_risks(scan)              → 人に見せる文
"""
import csv
import io
import re

_MONTHS = ('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'SEPT', 'OCT', 'NOV', 'DEC')
_KINDS = {
    'leading_zero': '先頭の 0 が消える',
    'long_digits': '16 桁以上で、15 桁より先が 0 になる',
    'sci_display': '12 桁以上で 1.23E+11 のような表示になる',
    'date_like': '日付に読み替えられる',
    'sci_notation': '指数（12E5）に読み替えられる',
    'plus': '先頭の + が消える',
    'formula': '= で始まり、式として計算される',
}
# 値が失われる（元に戻せない）もの。表示だけ崩れるもの（sci_display）と区別して知らせる
_LOSSY = {'leading_zero', 'long_digits', 'date_like', 'sci_notation', 'plus', 'formula'}


def sniff_encoding(path):
    """BOM があれば UTF-8、なければ UTF-8 として読めるか→読めなければ Shift-JIS。→ (python の名前, Excel の Origin)。"""
    with open(path, 'rb') as f:
        head = f.read(200000)
    if head.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig', 65001
    try:
        head.decode('utf-8')
        if any(b > 127 for b in head):
            return 'utf-8', 65001
        return 'cp932', 932                                  # ASCII だけならどちらでも同じ。日本語 Excel の既定に合わせる
    except UnicodeDecodeError:
        return 'cp932', 932


def sniff_delimiter(line):
    counts = {',': line.count(','), '\t': line.count('\t'), ';': line.count(';')}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ','


def damage_kind(v):
    """1 つの文字 → Excel が読み替えて壊す種類（'leading_zero' など）。壊さなければ None。"""
    s = (v or '').strip()
    if not s:
        return None
    if s.startswith('='):
        return 'formula'
    if re.fullmatch(r'0\d+', s):
        return 'leading_zero'
    if re.fullmatch(r'\+\d[\d.]*', s):
        return 'plus'
    if re.fullmatch(r'\d{16,}', s):
        return 'long_digits'
    if re.fullmatch(r'\d{12,15}', s):
        return 'sci_display'
    if re.fullmatch(r'\d{1,2}[-/]\d{1,2}', s):                # 1-2・10/3（月日に読まれる）。2026-01-02 のような年月日は本物の日付なので壊れではない
        return 'date_like'
    low = s.upper()
    for m in _MONTHS:                                         # JAN1・Mar-3（月名＋数字）
        if re.fullmatch(m + r'[-/ ]?\d{1,2}', low):
            return 'date_like'
    if re.fullmatch(r'\d+[eE][+-]?\d+', s):
        return 'sci_notation'
    return None


def scan_csv(path, max_rows=20000):
    """CSV を読み、壊れる列を返す → {'encoding', 'origin', 'delimiter', 'ncols', 'rows', 'header', 'cols': {列番号(1 始まり): {'kinds': {種類: 件数}, 'examples': [...]}}}。"""
    enc, origin = sniff_encoding(path)
    with open(path, 'r', encoding=enc, errors='replace', newline='') as f:
        text = f.read(8_000_000)
    first = text.split('\n', 1)[0]
    delim = sniff_delimiter(first)
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    cols, header, ncols, n = {}, [], 0, 0
    for i, row in enumerate(reader):
        if i == 0:
            header = row
        ncols = max(ncols, len(row))
        if i == 0:
            continue                                         # 見出しの行は数えない
        n += 1
        if n > max_rows:
            break
        for j, v in enumerate(row, 1):
            k = damage_kind(v)
            if k:
                c = cols.setdefault(j, {'kinds': {}, 'examples': []})
                c['kinds'][k] = c['kinds'].get(k, 0) + 1
                if len(c['examples']) < 3 and v.strip() not in c['examples']:
                    c['examples'].append(v.strip())
    return {'encoding': enc, 'origin': origin, 'delimiter': delim, 'ncols': ncols, 'rows': n, 'header': header, 'cols': cols}


def risky_field_info(scan):
    """Workbooks.OpenText の FieldInfo。壊れる列だけ xlTextFormat（2）、ほかは xlGeneralFormat（1）。"""
    return [(j, 2 if j in scan['cols'] else 1) for j in range(1, max(scan['ncols'], 1) + 1)]


def _col_name(scan, j):
    h = scan['header']
    return (h[j - 1].strip() if j - 1 < len(h) and h[j - 1].strip() else '') or f'{j} 列目'


def describe_risks(scan, limit=8):
    """人（と AI）に見せる文の並び。壊れる列が無ければ []。"""
    cols = scan['cols']
    if not cols:
        return []
    out = [f"CSV の自動変換から守りました: {len(cols)} 列を文字列として取り込みます"
           "（Excel の既定で開くと、保存した時点で元に戻らない読み替えが起きる列。他の列は普段どおり数・日付になります）"]
    for j in sorted(cols)[:limit]:
        c = cols[j]
        what = '・'.join(_KINDS[k] for k in c['kinds'])
        ex = ' '.join(c['examples'])
        out.append(f"  {j} 列目「{_col_name(scan, j)}」: {what}（例 {ex}）")
    if len(cols) > limit:
        out.append(f"  … 他 {len(cols) - limit} 列")
    out.append("  数として使うなら、そのあとで「選んだ列の文字の数字を数にする」（先頭の 0 の列は触らない）。"
               "Excel の既定どおり開くなら open … --excel-default")
    return out


__all__ = ['scan_csv', 'risky_field_info', 'describe_risks', 'damage_kind', 'sniff_encoding']
