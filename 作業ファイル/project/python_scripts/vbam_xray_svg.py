# -*- coding: utf-8 -*-
"""vbam_xray_svg.py — レントゲンの写真（SVG）。地図（vbam_xray.analyze の結果）から描く。

色は構想どおり:
  骨（見出し・表の背骨）＝白　／　筋肉（入力の値）＝青　／　血管（式）＝緑の線　／　病巣（症状・エラー）＝赤
写真の精度は地図で測る（地図が正しければ写真も正しい）。描けることだけ確かめる。
値は描かない（形だけ）。no-values でもそのまま使える。
"""
import re

from vbam_xray import col_letter, col_num, parse_addr, kind_of, _refs_in_formula

BG = '#0b1626'
BONE = '#f4f7fb'
MUSCLE = '#3b82f6'
MUSCLE_T = '#1d4ed8'
VESSEL = '#22c55e'
LESION = '#ef4444'
GRID = '#14243b'
CELL_W, CELL_H = 26, 12
MAX_ROWS, MAX_COLS = 140, 48


def _rows_of(s):
    out = set()
    for part in (s or '').split(','):
        part = part.strip()
        if not part or part.startswith('…'):
            continue
        a, _, z = part.partition('-')
        try:
            lo = int(a)
            hi = int(z) if z else lo
        except ValueError:
            continue
        out.update(range(lo, min(hi, lo + 5000) + 1))
    return out


def render_svg(m, sheet, title=None):
    vals = sheet['values']
    r0, c0 = sheet['r0'], sheet['c0']
    nr, nc = len(vals), (len(vals[0]) if vals else 0)
    show_r = min(nr, MAX_ROWS)
    show_c = min(nc, MAX_COLS)
    fm = sheet.get('formulas') or {}
    role = {}
    for b in m['blocks']:
        for run in b['rows']:
            a, _, z = run['rows'].partition('-')
            for r in range(int(a), int(z or a) + 1):
                role[r] = (b['id'], run['role'])
    bad = set()
    for b in m['blocks']:
        cols_idx = {}
        for s in b.get('symptoms', []):
            rows = _rows_of(s.get('rows', ''))
            if s.get('col'):
                cj = col_num(s['col'])
                for r in rows:
                    bad.add((r, cj))
            elif s['code'] in ('S12', 'S16'):
                rng = parse_addr(b['range'])
                for r in rows:
                    for cj in range(rng[1], rng[3] + 1):
                        bad.add((r, cj))
    for name, ent in (m.get('error_cells') or {}).items():
        for a in ent.get('first', []):
            p = parse_addr(a)
            if p:
                bad.add((p[0], p[1]))
    left, top = 34, 40
    W = left + show_c * CELL_W + 20
    H = top + show_r * CELL_H + 70
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}" font-family="sans-serif">',
           f'<rect width="{W}" height="{H}" fill="{BG}"/>']
    t = title or f"レントゲン {m.get('sheet', '')}  {m.get('used', '')}"
    if m.get('no_values'):
        t = f"レントゲン（値なし） {m.get('used', '')}"
    out.append(f'<text x="{left}" y="22" fill="{BONE}" font-size="13">{_esc(t)}</text>')
    # 列見出し・行見出し
    for j in range(show_c):
        out.append(f'<text x="{left + j * CELL_W + CELL_W / 2}" y="{top - 4}" fill="#5b7390" font-size="7" text-anchor="middle">{col_letter(c0 + j)}</text>')
    for i in range(0, show_r, 5):
        out.append(f'<text x="{left - 4}" y="{top + i * CELL_H + 9}" fill="#5b7390" font-size="7" text-anchor="end">{r0 + i}</text>')
    # セル
    for i in range(show_r):
        rr = r0 + i
        rl = role.get(rr, ('-', 'N'))[1]
        for j in range(show_c):
            v = vals[i][j]
            k = kind_of(v)
            x = left + j * CELL_W
            y = top + i * CELL_H
            isf = (rr, c0 + j) in fm
            if k in 'EW' and not isf:
                continue
            if (rr, c0 + j) in bad or k == 'X':
                fill, op = LESION, 0.95
            elif isf:
                fill, op = VESSEL, 0.9
            elif rl in ('H', 'T'):
                fill, op = BONE, 0.9
            elif rl in ('D', 'F', 'G'):
                fill, op = (MUSCLE if k in 'ND' else MUSCLE_T), 0.85
            else:
                fill, op = '#64748b', 0.5     # 題・注記・不明
            out.append(f'<rect x="{x + 1}" y="{y + 1}" width="{CELL_W - 2}" height="{CELL_H - 2}" fill="{fill}" fill-opacity="{op}"/>')
    # 血管（式 → 参照先の線）
    n_lines = 0
    for (fr, fc), f in fm.items():
        if n_lines > 120:
            break
        i, j = fr - r0, fc - c0
        if i >= show_r or j >= show_c:
            continue
        for (a1, b1, a2, b2, ext) in _refs_in_formula(f):
            if ext:
                continue
            ci, cj = (a1 + a2) / 2 - r0, (b1 + b2) / 2 - c0
            if ci >= show_r or cj >= show_c or ci < 0 or cj < 0:
                continue
            out.append(f'<line x1="{left + j * CELL_W + CELL_W / 2}" y1="{top + i * CELL_H + CELL_H / 2}" '
                       f'x2="{left + cj * CELL_W + CELL_W / 2}" y2="{top + ci * CELL_H + CELL_H / 2}" stroke="{VESSEL}" stroke-width="0.6" stroke-opacity="0.55"/>')
            n_lines += 1
            break
    # 表の枠（骨）
    for b in m['blocks']:
        p = parse_addr(b['range'])
        if not p:
            continue
        a1, b1, a2, b2 = p
        if a1 - r0 >= show_r or b1 - c0 >= show_c:
            continue
        x = left + (b1 - c0) * CELL_W
        y = top + (a1 - r0) * CELL_H
        w = (min(b2 - c0, show_c - 1) - (b1 - c0) + 1) * CELL_W
        h = (min(a2 - r0, show_r - 1) - (a1 - r0) + 1) * CELL_H
        col = BONE if b['kind'] == 'table' else '#64748b'
        out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="none" stroke="{col}" stroke-width="1" stroke-dasharray="{"0" if b["kind"] == "table" else "3 2"}"/>')
    # 凡例
    ly = top + show_r * CELL_H + 22
    items = [(BONE, '骨（見出し・合計）'), (MUSCLE, '筋肉（入力）'), (VESSEL, '血管（式）'), (LESION, '病巣（症状・エラー）')]
    x = left
    for c, name in items:
        out.append(f'<rect x="{x}" y="{ly - 9}" width="10" height="10" fill="{c}"/>')
        out.append(f'<text x="{x + 14}" y="{ly}" fill="{BONE}" font-size="10">{name}</text>')
        x += 14 + 11 * len(name) + 16
    note = ''
    if nr > show_r or nc > show_c:
        note = f'（先頭 {show_r} 行 × {show_c} 列まで描画。全体 {nr} 行 × {nc} 列）'
    out.append(f'<text x="{left}" y="{ly + 18}" fill="#7f95b2" font-size="9">表 {len(m["blocks"])} 個・症状 {len(m.get("symptoms", []))} 件・unknown {len(m.get("unknown", []))} 件 {note}</text>')
    out.append('</svg>')
    return '\n'.join(out)


def _esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
