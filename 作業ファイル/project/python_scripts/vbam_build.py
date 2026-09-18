# -*- coding: utf-8 -*-
"""vbam_build.py — vba_manager 分割パート: 「作る」コマンド（build-sheet）

シートの設計図（JSON）を受け取り、既存の手コマンド（値の書き込み・tidy・枠固定・入力規則・
テーブル・名前定義・条件付き書式・印刷設定）を順に当てて、帳票を一枚組み上げる。
フォームを form_layout の宣言から組み立てるのと同じ型をシートに広げたもの（2026-09-03）。
AI（または人）は設計図を書くだけ。組み立てるのは道具、確かめるのも道具。

設計図の例は `build-sheet --sample` が _last_sheet_spec.json に書き出す。
"""
import os
import sys
import json
import time
import argparse
import urllib.request
import urllib.error

from vbam_core import (SCRIPT_DIR, get_workbook, parse_target_and_rest, job_clock_note,
                       pinned_workbook)
from vbam_view import _show_range, cmd_screenshot
from vbam_edit import (cmd_tidy, _hex_to_excel_color, _report_write_result, _keep_grid_text,
                       _XL_ALIGN_H)

_LAST_SHEET_SPEC_FILE = os.path.join(SCRIPT_DIR, '_last_sheet_spec.json')
_LAST_SHEET_ASK_FILE = os.path.join(SCRIPT_DIR, '_last_sheet_ask.txt')     # AI に送った文（検分用）
_LAST_BUILD_PNG = os.path.join(SCRIPT_DIR, '_last_build.png')               # 組み上げ直後の見た目

# AI に 1 回聞く口（手2）。ループは Python が回し、モデルは設計図を 1 回書くだけ。
_AI_DEFAULT_MODEL = {'claude-code': 'sonnet', 'gemini': 'gemini-3.7-flash', 'claude': 'claude-haiku-4-5-20251001'}
#   claude-code＝ヘッドレスの Claude Code（claude -p・鍵なし・既定。2026-09-07）
_AI_TIMEOUT = 90
_API_RETRY_CODES = (429, 500, 502, 503, 504)     # 混雑・向こう側の一時エラー＝待てば通る
_API_RETRY_WAITS = (3, 8)                        # 待って撃ち直す秒（2 回まで）


def _api_call_with_retry(fn, ai='API'):
    """API を 1 回呼ぶ。429（混雑）と 5xx（向こう側の一時エラー）のときだけ、待って撃ち直す（2 回まで）。

    それ以外の HTTP エラー（400 の形違い・401 のキー違い・404 のモデル名違い）は撃ち直さない＝同じ返事が返るだけ。
    2026-09-04: flash は混む時間帯に 503 を返し、1 回で往復全体が失敗していた（agent・build-sheet --ask の両方）。
    """
    for i, wait in enumerate(_API_RETRY_WAITS + (None,)):
        try:
            return fn()
        except urllib.error.HTTPError as ex:
            if ex.code not in _API_RETRY_CODES or wait is None:
                raise
            print(f"{ai} の API が混んでいます（HTTP {ex.code}）。{wait} 秒待って撃ち直します"
                  f"（{i + 1}/{len(_API_RETRY_WAITS)}）")
            time.sleep(wait)

# 列の型 → 表示形式（"format" を明示すればそちらが勝つ）
_TYPE_FORMATS = {
    'text': '@',
    'number': '#,##0',
    'int': '0',
    'decimal': '#,##0.00',
    'currency': '#,##0',
    'percent': '0.0%',
    'date': 'yyyy/m/d',
    'time': 'h:mm',
}
_XL_OPERATORS = {'between': 1, 'not_between': 2, 'eq': 3, 'ne': 4,
                 'gt': 5, 'lt': 6, 'ge': 7, 'le': 8}

SAMPLE_SPEC = {
    "sheet": "会員名簿",
    "title": "会員名簿",
    "note": "金額は数量×単価で自動計算。区分はリストから選ぶ。",
    "columns": [
        {"name": "No", "type": "int", "width": 6},
        {"name": "会員番号", "type": "text", "width": 12},
        {"name": "氏名", "type": "text", "width": 16},
        {"name": "区分", "type": "text", "list": ["正会員", "準会員", "賛助"], "width": 10},
        {"name": "数量", "type": "int", "width": 8},
        {"name": "単価", "type": "currency", "width": 10},
        {"name": "金額", "type": "currency", "formula": "=IF({数量}=\"\",\"\",{数量}*{単価})", "width": 12},
        {"name": "入会日", "type": "date", "width": 12},
    ],
    "rows": 20,
    "total": {"label": "合計", "sum": ["数量", "金額"]},
    "freeze": True,
    "filter": True,
    "names": [{"name": "会員名簿範囲", "range": "table"}],
    "cond_format": [{"column": "金額", "gt": 100000, "bg": "#FFC7CE"}],
    "print": {"landscape": True, "fit_wide": 1, "title_rows": True},
}

# 手順書（帳票の型）。"_use" は「いつ使うか」。build-sheet --recipe 名前 で組む／--fire で全部を実射する。
# 使う人が自分の型を足すときは、ここに書き足すか、設計図 JSON を --recipe でなく直接渡す。
RECIPES = {
    "名簿": dict(SAMPLE_SPEC, _use="会員・名簿・連絡先の一覧。番号・氏名・区分・数量×単価・日付が並ぶ表"),
    "備品台帳": {
        "_use": "備品・物品の台帳。管理番号で追い、分類・取得日・取得価格・保管場所を持つ",
        "sheet": "備品台帳", "title": "備品台帳",
        "columns": [
            {"name": "管理番号", "type": "text", "width": 12},
            {"name": "品名", "type": "text", "width": 20},
            {"name": "分類", "type": "text", "list": ["机・椅子", "什器", "電子機器", "消耗品", "その他"], "width": 12},
            {"name": "数量", "type": "int", "width": 8},
            {"name": "取得日", "type": "date", "width": 12},
            {"name": "取得価格", "type": "currency", "width": 12},
            {"name": "保管場所", "type": "text", "width": 14},
            {"name": "備考", "type": "text", "width": 20},
        ],
        "rows": 30, "total": {"label": "合計", "sum": ["数量", "取得価格"]},
        "freeze": True, "filter": True,
        "print": {"landscape": True, "fit_wide": 1, "title_rows": True},
    },
    "月次集計表": {
        "_use": "項目×月の集計表。年度（4月始まり）の12か月と年計、列ごとの合計",
        "sheet": "月次集計表", "title": "月次集計表", "note": "月ごとの数値を入れると年計と合計が出ます。",
        "columns": [{"name": "項目", "type": "text", "width": 16}]
                   + [{"name": f"{m}月", "type": "number", "width": 9} for m in (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)]
                   + [{"name": "年計", "type": "number", "width": 11,
                       "formula": "=IF(COUNT({4月}:{3月})=0,\"\",SUM({4月}:{3月}))"}],
        "rows": 12,
        "total": {"label": "合計", "sum": [f"{m}月" for m in (4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3)] + ["年計"]},
        "freeze": True,
        "print": {"landscape": True, "fit_wide": 1, "title_rows": True},
    },
    "請求明細": {
        "_use": "請求書の明細。品名・数量・単価・金額、小計・消費税・税込合計",
        "sheet": "請求明細", "title": "請求明細", "note": "金額は数量×単価。消費税は10%で切り捨て。",
        "columns": [
            {"name": "No", "type": "int", "width": 5},
            {"name": "品名", "type": "text", "width": 24},
            {"name": "数量", "type": "int", "width": 8},
            {"name": "単位", "type": "text", "width": 6},
            {"name": "単価", "type": "currency", "width": 11},
            {"name": "金額", "type": "currency", "width": 12, "formula": "=IF({数量}=\"\",\"\",{数量}*{単価})"},
            {"name": "備考", "type": "text", "width": 16},
        ],
        "rows": 10,
        "total": {"label": "小計", "label_col": "品名", "sum": ["金額"],
                  "extra": [{"label": "消費税（10%）", "column": "金額", "formula": "=ROUNDDOWN({total:金額}*0.1,0)"},
                            {"label": "税込合計", "column": "金額", "formula": "={total:金額}+{above}"}]},
        "freeze": False,
        "print": {"landscape": False, "fit_wide": 1, "title_rows": True},
    },
    "出勤簿": {
        "_use": "ひと月の出勤簿。日付・曜日・出勤・退勤・休憩・実働、実働の合計",
        "sheet": "出勤簿", "title": "出勤簿", "note": "日付を入れると曜日が出ます。出勤・退勤・休憩は 9:00 のように入力。",
        "columns": [
            {"name": "日付", "type": "date", "width": 11},
            {"name": "曜日", "width": 6, "align": "center", "formula": "=IF({日付}=\"\",\"\",TEXT({日付},\"aaa\"))"},
            {"name": "出勤", "type": "time", "width": 8},
            {"name": "退勤", "type": "time", "width": 8},
            {"name": "休憩", "type": "time", "width": 8},
            {"name": "実働", "type": "time", "format": "[h]:mm", "width": 8,
             "formula": "=IF(OR({出勤}=\"\",{退勤}=\"\"),\"\",{退勤}-{出勤}-IF({休憩}=\"\",0,{休憩}))"},
            {"name": "備考", "type": "text", "width": 16},
        ],
        "rows": 31, "total": {"label": "合計", "sum": ["実働"]},
        "freeze": True,
        "print": {"landscape": False, "fit_tall": 1, "title_rows": True},
    },
}


# ----------------------------------------------------------------
# 純 Python 部分（Excel 不要）: 設計図 → 組み立て計画
# ----------------------------------------------------------------

def _letter(n):
    """1 → A, 27 → AA"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_index(s):
    """A → 1, AA → 27"""
    n = 0
    for ch in s.strip().upper():
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"列文字が不正です: {s}")
        n = n * 26 + (ord(ch) - 64)
    if n == 0:
        raise ValueError(f"列文字が不正です: {s}")
    return n


def _subst_formula(template, row, letters, first, last):
    """数式テンプレートの {row} {first} {last} {列名} を実番地に置き換える。

    {列名} は同じ行のそのセル（例: ={数量}*{単価} → =E4*F4）。未知の列名は ValueError。
    """
    out = template
    out = out.replace('{row}', str(row)).replace('{first}', str(first)).replace('{last}', str(last))
    i = 0
    while True:
        a = out.find('{', i)
        if a < 0:
            break
        b = out.find('}', a)
        if b < 0:
            raise ValueError(f"数式の波括弧が閉じていません: {template}")
        key = out[a + 1:b]
        if key not in letters:
            raise ValueError(f"数式の列名が設計図にありません: {{{key}}} （{template}）")
        out = out[:a] + f"{letters[key]}{row}" + out[b + 1:]
        i = a + 1
    return out


def plan_sheet(spec):
    """設計図（dict）→ 組み立て計画（dict）。Excel には触らない（テスト可能・--dry-run 用）。"""
    if not isinstance(spec, dict):
        raise ValueError("設計図は JSON のオブジェクト（{...}）で書いてください")
    sheet = str(spec.get('sheet') or '').strip()
    if not sheet:
        raise ValueError("設計図に sheet（シート名）がありません")
    cols = spec.get('columns') or []
    if not cols:
        raise ValueError("設計図に columns（列の一覧）がありません")
    title = spec.get('title')
    note = spec.get('note')
    col0 = _col_index(str(spec.get('start_col') or 'A'))
    header_row = spec.get('header_row')
    if header_row is None:
        header_row = 3 if (title or note) else 1
    header_row = int(header_row)
    if (title or note) and header_row < 3:
        raise ValueError("title / note があるときは header_row は 3 以上にしてください（1 行目が題名、2 行目が説明）")

    names = []
    letters = {}
    columns = []
    for j, c in enumerate(cols):
        if isinstance(c, str):
            c = {"name": c}
        nm = str(c.get('name') or '').strip()
        if not nm:
            raise ValueError(f"columns の {j + 1} 番目に name がありません")
        if nm in letters:
            raise ValueError(f"列名が重複しています: {nm}")
        letter = _letter(col0 + j)
        letters[nm] = letter
        names.append(nm)
        typ = str(c.get('type') or 'auto').lower()
        if typ not in _TYPE_FORMATS and typ != 'auto':
            raise ValueError(f"列 {nm} の type が不明です: {typ}（{', '.join(_TYPE_FORMATS)} / auto）")
        fmt = c.get('format') or _TYPE_FORMATS.get(typ)
        align = c.get('align')
        if align and align not in _XL_ALIGN_H:
            raise ValueError(f"列 {nm} の align が不明です: {align}（left/center/right）")
        columns.append({
            'name': nm, 'letter': letter, 'index': col0 + j, 'type': typ, 'format': fmt,
            'width': c.get('width'), 'align': align, 'formula': c.get('formula'),
            'list': c.get('list'), 'default': c.get('default'),
        })

    data = spec.get('data')
    if data:
        nrows = len(data)
    else:
        nrows = int(spec.get('rows') or 20)
    if nrows < 1:
        raise ValueError("rows は 1 以上にしてください")
    first = header_row + 1
    last = header_row + nrows
    ncols = len(columns)
    last_letter = _letter(col0 + ncols - 1)
    first_letter = _letter(col0)

    # データ格子（None＝空セル。数式列はテンプレートを実番地に）
    grid = []
    for i in range(nrows):
        r = first + i
        row_vals = []
        src = data[i] if data else None
        for j, c in enumerate(columns):
            if c['formula']:
                row_vals.append(_subst_formula(str(c['formula']), r, letters, first, last))
            elif src is not None and j < len(src):
                v = src[j]
                row_vals.append(None if v in ('', None) else v)
            elif c['default'] is not None:
                row_vals.append(c['default'])
            else:
                row_vals.append(None)
        grid.append(row_vals)

    total = spec.get('total')
    total_row = None
    total_cells = []
    extra_cells = []                  # 合計行の下に続く行（消費税・税込合計など）: (列文字, 値, 行)
    if total:
        total_row = last + 1
        label = total.get('label', '合計')
        label_col = total.get('label_col')
        if label_col and label_col not in letters:
            raise ValueError(f"total.label_col の列名が設計図にありません: {label_col}")
        label_letter = letters[label_col] if label_col else first_letter
        total_cells.append((label_letter, label))
        for nm in (total.get('sum') or []):
            if nm not in letters:
                raise ValueError(f"total.sum の列名が設計図にありません: {nm}")
            L = letters[nm]
            total_cells.append((L, f"=SUM({L}{first}:{L}{last})"))
        # extra: [{"label":"消費税（10%）","column":"金額","formula":"=ROUNDDOWN({total:金額}*0.1,0)"},
        #         {"label":"税込合計","column":"金額","formula":"={total:金額}+{above}"}]
        for k, ex in enumerate(total.get('extra') or []):
            row = total_row + 1 + k
            col = ex.get('column')
            if col not in letters:
                raise ValueError(f"total.extra の列名が設計図にありません: {col}")
            L = letters[col]
            tpl = str(ex.get('formula') or ex.get('value') or '')
            if tpl.startswith('='):
                tpl = tpl.replace('{above}', f"{L}{row - 1}")
                i = 0
                while True:
                    a = tpl.find('{total:', i)
                    if a < 0:
                        break
                    b = tpl.find('}', a)
                    key = tpl[a + 7:b]
                    if key not in letters:
                        raise ValueError(f"total.extra の数式の列名が設計図にありません: {{total:{key}}}")
                    tpl = tpl[:a] + f"{letters[key]}{total_row}" + tpl[b + 1:]
                    i = a + 1
                tpl = _subst_formula(tpl, row, letters, first, last)
            extra_cells.append((label_letter, ex.get('label', ''), row))
            extra_cells.append((L, tpl, row))
    end_row = (total_row + len(total.get('extra') or [])) if total_row else last

    header_addr = f"{first_letter}{header_row}:{last_letter}{header_row}"
    data_addr = f"{first_letter}{first}:{last_letter}{last}"
    table_addr = f"{first_letter}{header_row}:{last_letter}{last}"
    full_addr = f"{first_letter}{header_row}:{last_letter}{end_row}"

    def _range_of(key):
        if key in (None, 'table'):
            return table_addr
        if key == 'data':
            return data_addr
        if key == 'full':
            return full_addr
        if key in letters:
            return f"{letters[key]}{first}:{letters[key]}{last}"
        return str(key)               # 生の番地

    validations = [(f"{c['letter']}{first}:{c['letter']}{last}", [str(x) for x in c['list']])
                   for c in columns if c['list']]
    name_defs = []
    for n in (spec.get('names') or []):
        nm = str(n.get('name') or '').strip()
        if not nm:
            raise ValueError("names の要素に name がありません")
        name_defs.append((nm, _range_of(n.get('range'))))
    cond = []
    for cf in (spec.get('cond_format') or []):
        colname = cf.get('column')
        if colname and colname not in letters:
            raise ValueError(f"cond_format の列名が設計図にありません: {colname}")
        addr = _range_of(colname) if colname else _range_of(cf.get('range', 'data'))
        rule = None
        for op in _XL_OPERATORS:
            if op in cf:
                rule = (op, cf[op])
                break
        formula = cf.get('formula')
        if rule is None and not formula:
            raise ValueError("cond_format の各要素には gt/lt/ge/le/eq/ne/between か formula が要ります")
        cond.append({'addr': addr, 'rule': rule, 'formula': formula,
                     'bg': cf.get('bg'), 'color': cf.get('color'), 'bold': cf.get('bold', False)})

    pr = spec.get('print') or {}
    print_plan = None
    if pr:
        area_key = pr.get('area')
        if area_key in (None, 'full', 'table', 'data') or area_key in letters:
            area = _range_of(area_key or 'full')
            if area_key in (None, 'full') and (title or note):
                area = f"{first_letter}1:{last_letter}{end_row}"
        else:
            area = str(area_key)
        print_plan = {
            'area': area,
            'title_rows': f"{header_row}:{header_row}" if pr.get('title_rows', True) else None,
            'landscape': bool(pr.get('landscape', False)),
            'fit_wide': pr.get('fit_wide'), 'fit_tall': pr.get('fit_tall'),
            'center_h': bool(pr.get('center_h', False)),
        }

    return {
        'sheet': sheet, 'title': title, 'note': note, 'col0': col0,
        'header_row': header_row, 'first': first, 'last': last, 'nrows': nrows,
        'columns': columns, 'names_row': names, 'grid': grid,
        'total_row': total_row, 'total_cells': total_cells, 'extra_cells': extra_cells, 'end_row': end_row,
        'header_addr': header_addr, 'data_addr': data_addr, 'table_addr': table_addr,
        'full_addr': full_addr,
        'freeze': bool(spec.get('freeze', True)),
        'filter': bool(spec.get('filter', False)),
        'table': spec.get('table'),
        'validations': validations, 'names': name_defs, 'cond_format': cond,
        'print': print_plan, 'header_bg': spec.get('header_bg'),
        'overwrite': bool(spec.get('overwrite', False)),
    }


def _plan_summary(p):
    lines = [f"シート: {p['sheet']}  見出し行: {p['header_row']}  列: {len(p['columns'])}  "
             f"データ行: {p['nrows']}（{p['first']}〜{p['last']}）" + (f"  合計行: {p['total_row']}" if p['total_row'] else ""),
             "列: " + " / ".join(f"{c['letter']}={c['name']}" + (f"[{c['type']}]" if c['type'] != 'auto' else "")
                                 + ("(式)" if c['formula'] else "") + ("(選択)" if c['list'] else "")
                                 for c in p['columns'])]
    if p['grid'] and any(c['formula'] for c in p['columns']):
        ex = [v for v in p['grid'][0] if isinstance(v, str) and v.startswith('=')]
        lines.append("数式の例（1 行目）: " + "  ".join(ex))
    if p['total_cells']:
        lines.append("合計行: " + "  ".join(f"{L}{p['total_row']}={v}" for L, v in p['total_cells']))
    if p.get('extra_cells'):
        lines.append("続く行: " + "  ".join(f"{L}{r}={v}" for L, v, r in p['extra_cells'] if str(v).startswith('=')))
    extras = []
    if p['freeze']:
        extras.append(f"枠固定 {p['columns'][0]['letter']}{p['first']}")
    if p['table']:
        extras.append(f"テーブル {p['table']}")
    elif p['filter']:
        extras.append("オートフィルタ")
    if p['validations']:
        extras.append(f"入力規則 {len(p['validations'])} 列")
    if p['names']:
        extras.append("名前定義 " + ", ".join(f"{n}={a}" for n, a in p['names']))
    if p['cond_format']:
        extras.append(f"条件付き書式 {len(p['cond_format'])}")
    if p['print']:
        pp = p['print']
        extras.append("印刷 " + ("横" if pp['landscape'] else "縦")
                      + (f"・横{pp['fit_wide']}ページ" if pp['fit_wide'] else "")
                      + (f"・見出し行{pp['title_rows']}を繰り返し" if pp['title_rows'] else ""))
    if extras:
        lines.append("仕上げ: " + " / ".join(extras))
    lines.append(f"表の範囲: {p['full_addr']}")
    return "\n".join(lines)


# ----------------------------------------------------------------
# Excel に当てる
# ----------------------------------------------------------------

def _find_sheet(wb, name):
    for sh in wb.Sheets:
        if sh.Name == name:
            return sh
    return None


def _sheet_has_content(ws):
    try:
        ur = ws.UsedRange
        if ur.Rows.Count == 1 and ur.Columns.Count == 1 and ur.Row == 1 and ur.Column == 1:
            return ur.Cells(1, 1).Value is not None
        return True
    except Exception:
        return False


def _apply(xl, wb, p, overwrite):
    ws = _find_sheet(wb, p['sheet'])
    made = False
    if ws is None:
        ws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
        ws.Name = p['sheet']
        made = True
    elif _sheet_has_content(ws):
        if not overwrite:
            print(f"エラー: シート '{p['sheet']}' には既に中身があります。"
                  f"上書きするなら --overwrite（中身は全部消えます）か、別のシート名にしてください")
            return None
        ws.Cells.Clear()
        try:
            for lo in list(ws.ListObjects):
                lo.Unlist()
        except Exception:
            pass
    ws.Activate()
    col0 = p['col0']
    applied = []

    # 題名・説明
    if p['title']:
        c = ws.Cells(1, col0)
        c.Value = p['title']
        c.Font.Bold = True
        c.Font.Size = 14
        applied.append("題名")
    if p['note']:
        c = ws.Cells(2, col0)
        c.Value = p['note']
        c.Font.Size = 9
        c.Font.Color = _hex_to_excel_color('#808080')
        applied.append("説明")

    # 列の表示形式（値を書く前に。文字列列は @ にしておくと "1-2" が日付に化けない。
    # ただし数式列に @ を当てると数式が文字として入ってしまうので、数式列は @ にしない）
    for c in p['columns']:
        if c['format'] and not (c['formula'] and c['format'] == '@'):
            ws.Range(f"{c['letter']}{p['first']}:{c['letter']}{p['last']}").NumberFormat = c['format']

    # 見出し
    ws.Range(p['header_addr']).Value = (tuple(p['names_row']),)
    # データ格子（None は空。数式は '=' 始まりの文字列）
    grid = tuple(tuple(r) for r in p['grid'])
    ws.Range(p['data_addr']).Value = grid
    # 型を決めた列（日付・数値）は Excel の読み替えのほうが正しい＝文字に戻さない（2026-09-09）。
    # 戻すと設計図が [date] と言った列が文字の日付になり、ピボットの月別まとめ（group_date）ができない。
    # 守るのは文字列列（'@'）と型を決めなかった列だけ＝"0001"・"1-2" の先頭ゼロは今までどおり守る。
    typed_cols = tuple(i for i, c in enumerate(p['columns']) if c['format'] and c['format'] != '@')
    kept = _keep_grid_text(ws, p['first'], col0, grid, skip_cols=typed_cols)
    # 合計行
    if p['total_cells']:
        for L, v in p['total_cells']:
            ws.Range(f"{L}{p['total_row']}").Value = v
        ws.Range(f"{p['columns'][0]['letter']}{p['total_row']}:"
                 f"{p['columns'][-1]['letter']}{p['total_row']}").Font.Bold = True
        applied.append("合計行")
    for L, v, r in p.get('extra_cells') or []:
        ws.Range(f"{L}{r}").Value = v
        ws.Range(f"{L}{r}").Font.Bold = True
    if p.get('extra_cells'):
        # 続く行の数値列は合計行と同じ表示形式に
        for c in p['columns']:
            if c['format'] and c['format'] != '@':
                ws.Range(f"{c['letter']}{p['total_row']}:{c['letter']}{p['end_row']}").NumberFormat = c['format']
        applied.append(f"続く行{len({r for _, _, r in p['extra_cells']})}")

    # 仕上げ＝tidy（見出しの体裁・罫線・列の型・列幅・読み戻し）
    ns = argparse.Namespace(posargs=[p['full_addr']], header_from=None, bg=p['header_bg'],
                            no_header=False, no_border=False, no_col_format=False, no_autofit=False,
                            min_width=None, max_width=None, sheet_opt=ws.Name)
    # 対象ブックを固定してから呼ぶ（2026-09-08）。付けないと cmd_tidy が「アクティブブック自動検出」に落ち、
    # 組み立て先ではなく人が開いているブックを掴む＝--fire の 1 本目がここで落ちていた
    with pinned_workbook(wb):
        cmd_tidy(ns)

    # 設計図の指定が tidy の自動判定より優先（幅・寄せ）
    for c in p['columns']:
        if c['width'] is not None:
            try:
                ws.Columns(c['index']).ColumnWidth = float(c['width'])
            except Exception:
                pass
        if c['align']:
            ws.Range(f"{c['letter']}{p['first']}:{c['letter']}{p['end_row']}").HorizontalAlignment = _XL_ALIGN_H[c['align']]

    # 枠固定（見出しの下・最初の列の左上）
    if p['freeze']:
        try:
            xl.ActiveWindow.FreezePanes = False
            ws.Range(f"{p['columns'][0]['letter']}{p['first']}").Select()
            xl.ActiveWindow.FreezePanes = True
            ws.Range(f"{p['columns'][0]['letter']}{p['header_row']}").Select()
            applied.append("枠固定")
        except Exception as ex:
            print(f"⚠ 枠固定ができませんでした: {ex}")

    # 入力規則
    for addr, lst in p['validations']:
        rng = ws.Range(addr)
        rng.Validation.Delete()
        rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=",".join(lst))
        rng.Validation.InCellDropdown = True
    if p['validations']:
        applied.append(f"入力規則{len(p['validations'])}列")

    # テーブル or オートフィルタ
    if p['table']:
        lo = ws.ListObjects.Add(1, ws.Range(p['table_addr']), None, 1)   # xlSrcRange, xlYes
        try:
            lo.Name = str(p['table'])
        except Exception as ex:
            print(f"⚠ テーブル名を付けられませんでした（{p['table']}）: {ex}")
        applied.append(f"テーブル{lo.Name}")
    elif p['filter']:
        ws.Range(p['table_addr']).AutoFilter(1)
        applied.append("オートフィルタ")

    # 名前定義
    for nm, addr in p['names']:
        abs_addr = ws.Range(addr).Address                      # $A$3:$H$23
        refers = "='" + ws.Name.replace("'", "''") + "'!" + abs_addr
        wb.Names.Add(nm, refers)
    if p['names']:
        applied.append(f"名前{len(p['names'])}")

    # 条件付き書式（数値の比較は「数値であること」も条件に入れる。空文字 "" は Excel では
    # どんな数より大きいので、=IF(...,"",...) の空欄が gt で全部塗られる・2026-09-03 実測）
    for cf in p['cond_format']:
        rng = ws.Range(cf['addr'])
        head = rng.Cells(1, 1).Address.replace('$', '')                       # 範囲先頭基準（相対）。pywin32 の Address はプロパティ＝引数を渡せない
        if cf['formula']:
            fc = rng.FormatConditions.Add(2, 1, cf['formula'])                  # xlExpression は位置渡し（cmd_cond_format と同じ）
        else:
            op, val = cf['rule']
            sym = {'gt': '>', 'lt': '<', 'ge': '>=', 'le': '<=', 'eq': '=', 'ne': '<>'}
            if op in ('between', 'not_between'):
                v1, v2 = val
                inner = f"{head}>={v1},{head}<={v2}"
                f1 = (f"=AND(ISNUMBER({head}),{inner})" if op == 'between'
                      else f"=AND(ISNUMBER({head}),NOT(AND({inner})))")
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                f1 = f"=AND(ISNUMBER({head}),{head}{sym[op]}{val})"
            else:
                sval = str(val).replace('"', '""')
                f1 = f"={head}{sym[op]}\"{sval}\""
            fc = rng.FormatConditions.Add(2, 1, f1)
        if cf['bg']:
            fc.Interior.Color = _hex_to_excel_color(cf['bg'])
        if cf['color']:
            fc.Font.Color = _hex_to_excel_color(cf['color'])
        if cf['bold']:
            fc.Font.Bold = True
    if p['cond_format']:
        applied.append(f"条件付き書式{len(p['cond_format'])}")

    # 印刷設定
    if p['print']:
        pp = p['print']
        ps = ws.PageSetup
        ps.PrintArea = ws.Range(pp['area']).Address
        if pp['title_rows']:
            a, b = pp['title_rows'].split(':')
            ps.PrintTitleRows = f"${a}:${b}"
        ps.Orientation = 2 if pp['landscape'] else 1
        if pp['fit_wide'] is not None or pp['fit_tall'] is not None:
            ps.Zoom = False
            ps.FitToPagesWide = int(pp['fit_wide']) if pp['fit_wide'] is not None else False
            ps.FitToPagesTall = int(pp['fit_tall']) if pp['fit_tall'] is not None else False
        if pp['center_h']:
            ps.CenterHorizontally = True
        applied.append("印刷設定")

    # 確かめる: 書いた範囲のエラーセル
    full = ws.Range(p['full_addr'])
    _report_write_result(ws, full)
    if kept:
        print("文字列として保持: " + " ".join(kept))
    print(f"組み上げ: {ws.Name}!{p['full_addr']}  [{'新規シート' if made else '既存シート'}"
          + (", " + ", ".join(applied) if applied else "") + "]")
    return ws


# ----------------------------------------------------------------
# 手2: AI に設計図を 1 回書かせる口（Gemini / Claude の API 直・SDK 不要）
# ----------------------------------------------------------------

SPEC_RULES = """あなたは Excel の帳票を設計する係です。依頼を読んで、シート一枚の設計図を JSON で返してください。
JSON 以外の文字（説明・コードフェンス）は返さないでください。

設計図の形（キーはこの中から。要らないものは書かない）:
- sheet: シート名（短く。既存のシート名と重ねない）
- title: 1 行目に置く題名（任意） / note: 2 行目に置く短い説明（任意）
- columns: 列の一覧（左から順）。各列は
    name: 見出し（必須）
    type: text | int | number | decimal | currency | percent | date（分かるものだけ）
    width: 列幅（任意） / align: left | center | right（任意）
    formula: その列を数式にするときのテンプレート。同じ行の他の列は {列名} で書く。
             例: "=IF({数量}=\\"\\",\\"\\",{数量}*{単価})"（入力が空なら空欄にする形を基本にする）
    list: ドロップダウンの選択肢（配列。区分など決まった値の列だけ）
    default: 既定値（任意）
- rows: 空の入力行の数（既定 20。依頼に件数があればそれに合わせる）
- data: 実データを入れるなら行の配列（[[列1,列2,...],...]）。無ければ書かない
- total: 合計行 {"label":"合計","sum":["列名",...]}（金額・数量など足す意味のある列だけ）
    合計行の下に行を続けるなら "extra": [{"label":"消費税（10%）","column":"金額","formula":"=ROUNDDOWN({total:金額}*0.1,0)"},
    {"label":"税込合計","column":"金額","formula":"={total:金額}+{above}"}]（{total:列名}=その列の合計セル、{above}=一つ上のセル）
- type には time（h:mm）もある。時間の合計は format "[h]:mm"
- freeze: 見出しの下で枠固定（既定 true） / filter: オートフィルタ（一覧表なら true）
- table: テーブル化するならその名前（filter と両方は書かない）
- names: 名前定義 [{"name":"名前","range":"table"}]（他の表から参照する見込みがあるときだけ）
- cond_format: 条件付き書式 [{"column":"列名","gt":100000,"bg":"#FFC7CE"}]（gt/lt/ge/le/eq/ne/between）
- print: 印刷設定 {"landscape":true,"fit_wide":1,"title_rows":true}（印刷する帳票なら）

守ること:
- 依頼に書いてあることだけを設計図にする。頼まれていない列・仕掛けを足さない
- 計算できる列は値でなく formula にする（合計・金額・消費税・差引など）
- 既にある表を集計するシートなら、切り口の列（地域・担当・品目など）は空の入力行にせず、手元の様子の
  「多い値」に出ている実際の値を data に書く（rows は 0）。集計の列は formula で 'シート名'!範囲 を参照する
  （例 "=SUMIF('売上明細'!$B$2:$B$25,{地域},'売上明細'!$E$2:$E$25)"）。参照する範囲は手元の様子の使用範囲に合わせる
- 日付の列は type: date、金額は currency、番号や郵便番号のように先頭ゼロやハイフンを含む列は text
- 列名は依頼の言葉をそのまま使う

設計図の例:
"""


def _spec_prompt(request, materials=None):
    parts = [SPEC_RULES, json.dumps(SAMPLE_SPEC, ensure_ascii=False, indent=2)]
    if materials:
        parts.append("\n手元のブックの様子:\n" + materials)
    parts.append("\n依頼:\n" + request.strip() + "\n\n設計図（JSON）:")
    return "\n".join(parts)


def _extract_json(text):
    """返事から JSON を取り出す（コードフェンスや前後の文が混ざっていても拾う）"""
    t = text.strip()
    if t.startswith('```'):
        t = t.split('\n', 1)[1] if '\n' in t else ''
        if t.rstrip().endswith('```'):
            t = t.rstrip()[:-3]
    t = t.strip()
    try:
        return json.loads(t)
    except Exception:
        a, b = t.find('{'), t.rfind('}')
        if a >= 0:
            # 先頭の { から読める最初の JSON を取り、後ろは捨てる。sonnet（effort low・思考なし）が末尾に } を 1 つ余分に
            # 付けて返し、正しい手 7 本を「Extra data」で丸ごと捨てて 2 往復（25 秒）失った（2026-09-08）
            try:
                return json.JSONDecoder().raw_decode(t[a:])[0]
            except Exception:
                pass
        if a >= 0 and b > a:
            return json.loads(t[a:b + 1])
        raise


def _ai_gemini(prompt, model, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    body = {"contents": [{"role": "user", "parts": [{"text": prompt}]}],
            "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}}
    req = urllib.request.Request(url, data=json.dumps(body).encode('utf-8'), method='POST',
                                 headers={"Content-Type": "application/json", "x-goog-api-key": api_key})
    data = json.load(urllib.request.urlopen(req, timeout=_AI_TIMEOUT))
    cand = (data.get("candidates") or [{}])[0]
    text = "".join(p.get("text", "") for p in (cand.get("content") or {}).get("parts", []))
    u = data.get("usageMetadata") or {}
    usage = {'in': u.get('promptTokenCount', 0), 'out': u.get('candidatesTokenCount', 0),
             'think': u.get('thoughtsTokenCount', 0), 'total': u.get('totalTokenCount', 0)}
    return text, usage


def _ai_claude(prompt, model, api_key):
    url = "https://api.anthropic.com/v1/messages"
    body = {"model": model, "max_tokens": 4096,
            "messages": [{"role": "user", "content": prompt}]}
    req = urllib.request.Request(url, data=json.dumps(body).encode('utf-8'), method='POST',
                                 headers={"Content-Type": "application/json", "x-api-key": api_key,
                                          "anthropic-version": "2023-06-01"})
    data = json.load(urllib.request.urlopen(req, timeout=_AI_TIMEOUT))
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    u = data.get("usage") or {}
    usage = {'in': u.get('input_tokens', 0), 'out': u.get('output_tokens', 0), 'think': 0,
             'total': u.get('input_tokens', 0) + u.get('output_tokens', 0)}
    return text, usage


def ask_design(request, materials=None, ai=None, model=None):    # ai=None → _ai_setup が claude-code に落とす（2026-09-09）
    """依頼文 → 設計図（dict）。AI への往復は 1 回。請求書（字数・トークン・秒）を表示する。"""
    # キーの読み方は agent と同じ 1 か所（環境変数 → 金庫）。遅延 import＝vbam_ai がこのモジュールを import しているため。
    # 2026-09-04: ここだけ環境変数しか見ておらず、キーを金庫に移した後 build が「環境変数が設定されていません」で止まっていた
    # 2026-09-12: vbam_agent 経由をやめ、持ち主の vbam_ai から直に取る
    from vbam_ai import _ai_setup, _cc_oneshot
    ai, model, api_key = _ai_setup(ai, model)
    prompt = _spec_prompt(request, materials)
    with open(_LAST_SHEET_ASK_FILE, 'w', encoding='utf-8') as f:
        f.write(prompt)
    t0 = time.time()
    try:
        text, usage = _api_call_with_retry(
            lambda: (_cc_oneshot(prompt, model) if ai == 'claude-code'
                     else (_ai_gemini if ai == 'gemini' else _ai_claude)(prompt, model, api_key)), ai)
    except urllib.error.HTTPError as ex:
        detail = ''
        try:
            detail = ex.read().decode('utf-8', 'replace')[:300]
        except Exception:
            pass
        raise RuntimeError(f"{ai} の API がエラーを返しました（HTTP {ex.code}）: {detail}")
    sec = time.time() - t0
    print(f"請求書: 送り {len(prompt):,}字 / 返り {len(text):,}字 / 待ち {sec:.1f}秒 / "
          f"トークン 入力 {usage['in']:,}・出力 {usage['out']:,}"
          + (f"・思考 {usage['think']:,}" if usage['think'] else "") + f" / モデル {model}")
    try:
        spec = _extract_json(text)
    except Exception:
        print("エラー: AI の返事が設計図（JSON）になっていません。返事の先頭:")
        print(text[:400])
        raise RuntimeError("設計図を取り出せませんでした")
    with open(_LAST_SHEET_SPEC_FILE, 'w', encoding='utf-8') as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    print(f"設計図を保存: {_LAST_SHEET_SPEC_FILE}")
    return spec


def _book_materials(wb, max_sheets=20, max_cols=12):
    """AI に渡す手元の様子: シート名・使用範囲・見出し行（2026-09-04）。

    前はシート名だけだった＝新しいシートに「既存の表から集計する数式」を書かせようがなかった
    （名前の衝突を避けるためだけの材料）。見出しまで見せると、参照先の列を当てて組める。
    """
    try:
        names = [sh.Name for sh in wb.Sheets]
    except Exception:
        return None
    if not names:
        return None
    lines = ["既存のシート: " + ", ".join(names)]
    profiled, max_profiled = 0, 5
    for sh in list(wb.Sheets)[:max_sheets]:
        try:
            ur = sh.UsedRange
            addr = str(ur.Address).replace('$', '')
            nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
            if nr == 1 and nc == 1 and ur.Cells(1, 1).Value is None:
                lines.append(f"  {sh.Name}: （空）")
                continue
            head = ''
            v = ur.Rows(1).Value
            row = list(v[0]) if isinstance(v, tuple) and v and isinstance(v[0], tuple) else ([v] if not isinstance(v, tuple) else list(v))
            cells = [('' if c is None else str(c)) for c in row[:max_cols]]
            if any(cells):
                head = "  見出し: " + " | ".join(cells) + ("…" if nc > max_cols else "")
            lines.append(f"  {sh.Name}: {addr}（{nr} 行 × {nc} 列）" + head)
            # 列ごとの型と「多い値」まで見せる。見出しだけだと、集計シートの左端（地域・担当…）を
            # 埋められず、空の枠だけ組んで「合格」になる（2026-09-04 実射）。重い表は見出しだけで我慢する
            if profiled < max_profiled and nr * nc <= 5000:
                from vbam_hands import _profile_columns, _rows_of      # 持ち主から直に取る（2026-09-12・前は vbam_agent 経由）
                sub = sh.Range(sh.Cells(int(ur.Row), int(ur.Column)),
                               sh.Cells(int(ur.Row) + min(nr, 500) - 1, int(ur.Column) + min(nc, max_cols) - 1))
                lines += _profile_columns(_rows_of(sub.Value), None, int(ur.Row), int(ur.Column))
                profiled += 1
        except Exception:
            continue
    return "\n".join(lines)


# ----------------------------------------------------------------
# 手3: 確かめる係（エラーセル・###・数式の本数・画像）
# ----------------------------------------------------------------

def _verify(xl, ws, p):
    rng = ws.Range(p['full_addr'])
    n_err = 0
    for typ in (-4123, 2):
        try:
            n_err += int(rng.SpecialCells(typ, 16).Count)
        except Exception:
            pass
    n_hash = 0
    try:
        if int(rng.Cells.CountLarge) <= 2000:
            for c in rng.Cells:
                t = c.Text
                if isinstance(t, str) and t.startswith('#') and set(t) == {'#'}:
                    n_hash += 1
    except Exception:
        pass
    n_formula = sum(1 for row in p['grid'] for v in row if isinstance(v, str) and v.startswith('=')) \
        + sum(1 for _, v in p['total_cells'] if isinstance(v, str) and v.startswith('=')) \
        + sum(1 for _, v, _r in (p.get('extra_cells') or []) if isinstance(v, str) and v.startswith('='))
    png = None
    try:
        top = f"{p['columns'][0]['letter']}1" if (p['title'] or p['note']) else p['full_addr'].split(':')[0]
        shot = f"{top}:{p['full_addr'].split(':')[1]}"
        ns = argparse.Namespace(posargs=[shot], out_opt=_LAST_BUILD_PNG)
        with pinned_workbook(ws.Parent):     # 組み立て先のブックを写す（人の開いているブックではなく）
            if cmd_screenshot(ns):
                png = _LAST_BUILD_PNG
    except Exception as ex:
        print(f"⚠ 画像を書き出せませんでした: {ex}")
    ok = (n_err == 0 and n_hash == 0)
    print(f"検査: エラーセル {n_err} / ### {n_hash} / 数式 {n_formula} 本 / "
          + (f"画像 {png}" if png else "画像なし") + ("  → 合格" if ok else "  → 要確認"))
    return ok


def fire_recipes(xl, names=None):
    """手順書の自動実射: まっさらなブックに全部を組み上げて検査し、保存せず閉じる。全部合格なら True。"""
    names = list(names or RECIPES)
    bad = [n for n in names if n not in RECIPES]
    if bad:
        print(f"エラー: 無い手順書: {', '.join(bad)}（--recipes で一覧）")
        return False
    wb = xl.Workbooks.Add()
    print(f"実射用のブック: {wb.Name}（終わったら保存せず閉じます）")
    results = []
    t0 = time.time()
    try:
        for n in names:
            print(f"\n===== 実射: {n} =====")
            try:
                p = plan_sheet(RECIPES[n])
                ws = _apply(xl, wb, p, False)
                ok = bool(ws) and _verify(xl, ws, p)
            except Exception as ex:
                print(f"エラー: {ex}")
                ok = False
            results.append((n, ok))
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception as ex:
            print(f"⚠ 実射用のブックを閉じられませんでした: {ex}")
    print(f"\n実射の結果（{time.time() - t0:.1f}秒）:")
    for n, ok in results:
        print(f"  {'PASS' if ok else 'FAIL'}  {n}")
    n_ok = sum(1 for _, ok in results if ok)
    print(f"{n_ok}/{len(results)} 合格")
    return n_ok == len(results)


def cmd_build_sheet(args):
    """シートを設計図から組み上げる:
    build-sheet [excel_file] [設計図.json] [--overwrite] [--dry-run] [--sample] [--new-book]
    build-sheet --ask "依頼文" [--ai claude-code|gemini|claude] [--model 名] [--new-book] [--dry-run]
    build-sheet --recipes / --recipe 名前 [--new-book] / --fire [名前 ...]
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if getattr(args, 'recipes', False):
        print("手順書（帳票の型）:")
        for n, r in RECIPES.items():
            cols = "・".join(c['name'] if isinstance(c, dict) else str(c) for c in r['columns'])
            print(f"  {n}: {r.get('_use', '')}\n      列: {cols}")
        print("使い方: build-sheet --recipe 名前 [--new-book] / build-sheet --fire（全部を実射）")
        return True
    if getattr(args, 'fire', False):
        xl, wb = get_workbook(target_file)
        return fire_recipes(xl, rest or None)
    recipe = getattr(args, 'recipe', None)
    if recipe:
        if recipe not in RECIPES:
            print(f"エラー: 手順書 '{recipe}' はありません。--recipes で一覧")
            return False
        with open(_LAST_SHEET_SPEC_FILE, 'w', encoding='utf-8') as f:
            json.dump(RECIPES[recipe], f, ensure_ascii=False, indent=2)
        print(f"手順書 '{recipe}' の設計図を {_LAST_SHEET_SPEC_FILE} に置きました")
        rest = []
    if getattr(args, 'sample', False):
        with open(_LAST_SHEET_SPEC_FILE, 'w', encoding='utf-8') as f:
            json.dump(SAMPLE_SPEC, f, ensure_ascii=False, indent=2)
        print(json.dumps(SAMPLE_SPEC, ensure_ascii=False, indent=2))
        print(f"\n設計図の例を書き出しました: {_LAST_SHEET_SPEC_FILE}")
        print("  build-sheet で組み上げ（引数なしならこのファイル）。--dry-run で計画だけ表示")
        return True
    ask = getattr(args, 'ask', None)
    new_book = bool(getattr(args, 'new_book', False))
    dry = bool(getattr(args, 'dry_run', False))
    if len(rest) > 1 or (ask and rest):
        print("使い方: build-sheet [excel_file] [設計図.json] [--overwrite] [--dry-run] [--sample] [--new-book]")
        print("        build-sheet --ask \"依頼文\" [--ai claude-code|gemini|claude] [--model 名] [--new-book] [--dry-run]")
        return False

    xl = wb = None
    if not dry or ask:
        # 依頼のときは既存のシート名を AI に渡す（--new-book なら空のブックなので渡さない）
        if not (dry and not ask):
            xl, wb = get_workbook(target_file)
    if ask:
        materials = None if new_book else (_book_materials(wb) if wb is not None else None)
        try:
            spec = ask_design(ask, materials, getattr(args, 'ai', None), getattr(args, 'model', None))
        except (RuntimeError, ValueError, urllib.error.URLError) as ex:
            print(f"エラー: {ex}")
            return False
    else:
        path = rest[0] if rest else _LAST_SHEET_SPEC_FILE
        if not os.path.exists(path):
            print(f"エラー: 設計図が見つかりません: {path}")
            print("  build-sheet --sample で例を書き出す／--ask \"依頼文\" で AI に書かせる")
            return False
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                spec = json.load(f)
        except Exception as ex:
            print(f"エラー: 設計図（JSON）を読めません: {ex}")
            return False
    try:
        plan = plan_sheet(spec)
    except ValueError as ex:
        print(f"エラー: 設計図に問題があります: {ex}")
        return False
    overwrite = bool(getattr(args, 'overwrite', False)) or plan['overwrite']
    print(_plan_summary(plan))
    if dry:
        print("（--dry-run: Excel には触っていません）")
        return True
    if xl is None:
        xl, wb = get_workbook(target_file)
    if new_book:
        wb = xl.Workbooks.Add()                       # 訓練場＝まっさらなブック（人のブックに触らない）
        print(f"新しいブックに組み上げます: {wb.Name}")
    ws = _apply(xl, wb, plan, overwrite)
    if ws is None:
        return False
    ok = _verify(xl, ws, plan)
    _cn = job_clock_note()
    if _cn:
        print(_cn)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return ok


__all__ = ['cmd_build_sheet', 'plan_sheet', 'ask_design', 'fire_recipes', 'SAMPLE_SPEC', 'RECIPES']
