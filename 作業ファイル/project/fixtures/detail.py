# -*- coding: utf-8 -*-
"""お題の「明細の表」を種で作る（2026-09-18）。見出しの語・シート名・項目の名前・列の並び・表題・合計の行を種で変える。

  t = design(seed, n)                      … 明細の設計（列・行・シート名…）
  lay = write_before(path, t)              … openpyxl で 1 シートのブックを書く → {'h','last','ncols','cols'}
  ピボット・月次集計・クロス集計など「明細を元にする仕事」のお題で使う。
"""
import datetime
import os
import random
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vocab import Vocab   # noqa: E402

TITLES = ['令和8年度 {k}別支出明細（一般会計）', '令和8年度 支出の明細', '{k}ごとの支出一覧（9月末）']


def design(seed, n, **kw):
    rng = random.Random(seed)
    v = Vocab(seed)
    cols = [(v.word('date'), 'date'), (v.word('ka'), 'item'), (v.word('kamoku'), 'kamoku'),
            (v.word('tekiyo'), 'text'), (v.word('amount'), 'num')]
    t = {'sheet': v.sheet('meisai'), 'cols': cols, 'title': None, 'unit': False, 'total': False, 'blank_at': None,
         'no_col': False, 'swap': False, 'ka': v.word('ka'), 'no': v.word('no'), 'code': v.word('code'),
         'items': v.items(min(max(3, n // 8), 12), rng), 'kamokus': v.kamokus(rng.randint(3, 6), rng),
         'n': n, 'seed': seed, 'vocab': v, 'year': 2026,
         'n_items': None, 'text_date': False, 'text_amount': False, 'item_space': False, 'gap_months': (),
         'sort': True, 'blanks': 0}
    if kw.pop('title', False):
        t['title'] = rng.choice(TITLES).format(k=t['ka'])
    t.update(kw)
    if t['swap']:
        t['cols'] = [cols[1], cols[4], cols[0], cols[2], cols[3]]
    # 日付の無い表・数の無い表（2026-09-18 夜: お題が全部「日付つきの明細」だったので、道具が 28 本に「日付の列がある表だけ」の
    # 条件を導き、名簿・マスタ・集計表では撃てなかった。仕事に要らない列を抜いた表を 1 枚足して条件を導き直す）
    if t.get('no_date'):
        t['cols'] = [c for c in t['cols'] if c[1] != 'date']
    if t.get('no_num'):
        t['cols'] = [c for c in t['cols'] if c[1] != 'num']
    if t['no_col']:
        t['cols'] = [(t['no'], 'no')] + list(t['cols'])
    if t['n_items']:
        t['items'] = v.items(t['n_items'], random.Random(seed + 7))
    t['rows'] = rows_of(t, rng)
    if t['sort']:
        dn = next((n for n, r in t['cols'] if r == 'date'), None)
        if dn:
            t['rows'].sort(key=lambda r0: (date_of(r0[dn]) or datetime.datetime(2000, 1, 1)))
    if t['blanks']:
        # 空行の位置は「前」と「正解」で同じにする（お題を 2 回書くので種で決める）
        rb = random.Random(seed + 13)
        t['blank_rows'] = sorted(rb.sample(range(1, max(2, t['n'] - 1)), min(t['blanks'], max(0, t['n'] - 2))))
    else:
        t['blank_rows'] = [t['blank_at']] if t['blank_at'] is not None else []
    return t


_ZEN = str.maketrans('0123456789,', '０１２３４５６７８９，')


def rows_of(t, rng):
    out = []
    months = [m for m in (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3) if m not in (t.get('gap_months') or ())]
    for i in range(t['n']):
        m = rng.choice(months)
        row = {}
        for name, role in t['cols']:
            if role == 'no':
                row[name] = 1000 + i
            elif role == 'date':
                d = datetime.datetime(t['year'] if m >= 4 else t['year'] + 1, m, rng.randint(1, 28))
                if t.get('text_date') and rng.random() < 0.5:
                    d = (f"{d.year}/{d.month:02d}/{d.day:02d}" if rng.random() < 0.5
                         else f"{d.year}年{d.month}月{d.day}日")
                row[name] = d
            elif role == 'item':
                row[name] = rng.choice(t['items'])
                if t.get('item_space') and rng.random() < 0.3:
                    row[name] = row[name] + rng.choice((' ', '　', '  '))
            elif role == 'kamoku':
                row[name] = rng.choice(t['kamokus'])
            elif role == 'num':
                row[name] = rng.randrange(1, 900) * 100
                if t.get('text_amount') and rng.random() < 0.5:
                    row[name] = (f"{row[name]:,}".translate(_ZEN) if rng.random() < 0.5 else f" {row[name]:,} ")
            else:
                row[name] = rng.choice(['消耗品購入', '会議費用', '委託料支払', '修繕工事', '備品購入', '印刷代'])
        out.append(row)
    return out


def write_into(ws, t):
    """既にあるシートに明細を書く（openpyxl の Worksheet）。→ lay。"""
    h = 1
    c0 = 1
    if t['title']:
        ws.cell(1, c0, t['title']).font = Font(bold=True, size=14)
        h = 3
        if t['unit']:
            ws.cell(2, c0, '単位：円')
            h = 4
    for j, (name, _role) in enumerate(t['cols'], c0):
        ws.cell(h, j, name).font = Font(bold=True)
    r = h
    for i, row in enumerate(t['rows']):
        if i in (t.get('blank_rows') or []):
            r += 1
        r += 1
        for j, (name, role) in enumerate(t['cols'], c0):
            cell = ws.cell(r, j, row[name])
            if role == 'date':
                cell.number_format = 'yyyy/m/d'
            elif role == 'num':
                cell.number_format = '#,##0'
    last = r
    if t['total'] and any(role == 'num' for _n, role in t['cols']):
        text_col = next(j for j, (_n, role) in enumerate(t['cols'], c0) if role in ('item', 'text'))
        num_col = next(j for j, (_n, role) in enumerate(t['cols'], c0) if role == 'num')
        ws.cell(last + 1, text_col, '合計').font = Font(bold=True)
        ws.cell(last + 1, num_col, sum(num_of(r0[t['cols'][num_col - c0][0]]) for r0 in t['rows'])).number_format = '#,##0'
    return {'h': h, 'last': last, 'ncols': len(t['cols']), 'c0': c0}


def num_of(v):
    """文字で入れた金額（全角・カンマ・空白つき）も数にする（純 Python）。"""
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().translate(str.maketrans('０１２３４５６７８９，', '0123456789,')).replace(',', '')
    return float(s) if s else 0


def date_of(v):
    """文字で入れた日付（2026/04/15・2026年4月15日）も日付にする（純 Python）。"""
    import re as _re
    if isinstance(v, datetime.datetime):
        return v
    m = _re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', str(v))
    return datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def write_before(path, t):
    wb = Workbook()
    ws = wb.active
    ws.title = t['sheet']
    lay = write_into(ws, t)
    wb.save(path)
    return lay


def col_of(t, role):
    """役（'item','num','date','kamoku'）の列番号（1 から）。"""
    return next(j for j, (_n, r) in enumerate(t['cols'], 1) if r == role)


def name_of(t, role):
    return next(n for n, r in t['cols'] if r == role)
