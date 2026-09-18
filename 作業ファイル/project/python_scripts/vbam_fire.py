# -*- coding: utf-8 -*-
"""vbam_fire.py — vba_manager 分割パート: agent の実射（練習台・弾・答え合わせ・状態・写し取り・点数）

2026-09-11 に vbam_agent.py（12,600 行）から中身を変えずに切り出した（約 3,300 行）。
ここにあるのは開発用の道具＝本番の往復では使わない:
  agent --fire [名前 ...]      自動実射（まっさらなブックに練習台を組み、依頼を撃ってセルの実物で答え合わせ）
  agent --harvest 元 出力       他人のブックの「構造だけ」を写し取る（練習台の材料）
  agent --cases / --keep-case / --drop-case   本番から拾った弾の台帳
  agent --score                実射の点数の履歴（退行を見る）

往復の本体（run_agent・run_macro_agent・関所・控え・材料）は vbam_agent.py にある。こちらからは `va.` で呼ぶ
（呼ぶ瞬間に vbam_agent の名前を引く＝テストが vbam_agent 側を差し替えても、こちらの手にそのまま効く）。
vbam_agent 側はこの module を _cmd_agent_body の中で遅延 import する（循環にしない）。
"""
import os
import re
import sys
import json
import time
import contextlib
import shutil
from vbam_core import (SCRIPT_DIR, _LAST_VALUES_FILE)
from vbam_recipes import (RECIPE_CASES)

import vbam_agent as va


_AGENT_SCORE_FILE = os.path.join(SCRIPT_DIR, '_agent_score.jsonl')   # 実射の点数の履歴（退行を見るため）
_FIRE_OUT_FILE = os.path.join(SCRIPT_DIR, '_fire_last.txt')     # 実射の画面出力の控え（落ちた弾を後から読むため）
_AGENT_CASES_FILE = os.path.join(SCRIPT_DIR, '_agent_cases.json')    # 本番から拾った弾の台帳（2026-09-06）
_AGENT_CASES_DIR = os.path.join(SCRIPT_DIR, '_agent_cases')          # その弾の練習台（写し取り）の置き場

# ----------------------------------------------------------------
# 自動実射（人にテストさせない）: 練習台を組んで依頼を撃ち、セルの実物で答え合わせ
# ----------------------------------------------------------------

_FIRE_BASE = {
    "title": "会員名簿（練習台）",
    "columns": [
        {"name": "会員番号", "type": "text", "width": 10},
        {"name": "氏名", "type": "text", "width": 12},
        {"name": "郵便番号", "type": "text", "width": 11},
        {"name": "住所", "type": "text", "width": 26},
        {"name": "電話", "type": "text", "width": 14},
    ],
    "data": [
        ["2001", "青木 誠", "000-0001", "架空県中央市本町1-1-1", "000-000-1111"],
        ["2002", "石川 恵", "0000002", "架空県中央市大手町１－２－３", "000-000-2222"],
        ["2003", "上田 学", "000-0003", "架空県北山市青葉1-1", "000-000-3333"],
        ["2004", "遠藤 香", "0000004", "架空県南川市旭町8-2", "000-000-4444"],
        ["2005", "大西 亮", "000-0005", "架空県東野市栄17", "000-000-5555"],
        ["2006", "加藤 環", "0000006", "架空県西原市宮下１番３号", "000-000-6666"],
    ],
    "freeze": False, "filter": False,
}
# 右の表（申込一覧）は I5 起点に書く（見出し I5:K5・データ I6:K10）
_FIRE_RIGHT = [["会員番号", "申込コース", "金額"],
               ["2001", "年間", 12000], ["2003", "半年", 7000], ["2004", "年間", 12000],
               ["2009", "月額", 1200], ["2006", "半年", 7000]]
_FIRE_EXPECT_COURSE = {"2001": "年間", "2002": "申込なし", "2003": "半年", "2004": "年間",
                       "2005": "申込なし", "2006": "半年"}


def _check_postal(ws):
    bad = [f"C{r}" for r in range(4, 10) if not re.fullmatch(r'\d{3}-\d{4}', str(ws.Range(f"C{r}").Value or ''))]
    return (not bad), ("郵便番号が 000-0000 になっていないセル: " + " ".join(bad) if bad else "C4:C9 すべて 000-0000")


def _check_match(ws):
    got = {str(ws.Range(f"A{r}").Value or ''): str(ws.Range(f"F{r}").Value or '') for r in range(4, 10)}
    bad = [f"{k}→{got.get(k)!r}" for k, v in _FIRE_EXPECT_COURSE.items() if got.get(k) != v]
    return (not bad), ("F 列の突き合わせが違う: " + " ".join(bad) if bad else "F4:F9 が期待どおり（申込なし含む）")


def _check_sum(ws):
    want = sum(r[2] for r in _FIRE_RIGHT[1:])
    for r in range(11, 14):
        v = ws.Range(f"K{r}").Value
        try:
            if v is not None and abs(float(v) - want) < 0.5:
                return True, f"K{r} = {int(want):,}"
        except (TypeError, ValueError):
            pass
    return False, f"申込一覧の下（K11〜K13）に合計 {int(want):,} が無い"


def _lo_by_name(wb, name):
    for sh in wb.Worksheets:
        try:
            for lo in sh.ListObjects:
                if lo.Name == name:
                    return sh, lo
        except Exception:
            continue
    return None, None


def _check_table(ws):
    sh, lo = _lo_by_name(ws.Parent, 'T名簿')
    if lo is None:
        names = [x.Name for s in ws.Parent.Worksheets for x in s.ListObjects]
        return False, "テーブル T名簿 が無い（ある: " + (", ".join(names) or "なし") + "）"
    a = va._addr(lo.Range)
    return (a == 'A3:E9' and sh.Name == ws.Name), f"テーブル T名簿 [{sh.Name}] 範囲={a}" + ("" if a == 'A3:E9' else "（A3:E9 でない）")


def _check_pivot(ws):
    want = {}
    for r in _FIRE_RIGHT[1:]:
        want[r[1]] = want.get(r[1], 0) + r[2]
    pts = [(sh, pt) for sh in ws.Parent.Worksheets for pt in sh.PivotTables()]
    if not pts:
        return False, "ピボットが無い"
    sh, pt = pts[0]
    got = {}
    for row in va._rows_of(pt.TableRange2.Value):
        if len(row) >= 2 and isinstance(row[0], str) and row[0] in want:
            try:
                got[row[0]] = float(row[1])
            except (TypeError, ValueError):
                pass
    bad = [k for k, v in want.items() if abs(got.get(k, -1) - v) > 0.5]
    where = f"[{sh.Name}] {pt.Name} 出力={va._addr(pt.TableRange2)}"
    if bad:
        return False, f"{where}: コース別の合計が明細と違う {bad}（読めた {got}）"
    return sh.Name == ws.Name, where + ": コース別の合計が明細と一致（" + ", ".join(
        f"{k}={int(v):,}" for k, v in want.items()) + "）" + ("" if sh.Name == ws.Name else "（別シートに置いた）")


def _check_slicer(ws):
    wb = ws.Parent
    sh, lo = _lo_by_name(wb, 'T申込')
    if lo is None:
        return False, "テーブル T申込 が無い"
    names = []
    for sc in wb.SlicerCaches:
        if str(sc.SourceName) == '申込コース':
            names += [sl.Name for sl in sc.Slicers]
    return bool(names), (f"テーブル T申込 範囲={va._addr(lo.Range)}  スライサー(申込コース): " + ", ".join(names)
                         if names else f"テーブル T申込 はある（{va._addr(lo.Range)}）が申込コースのスライサーが無い")


def _check_powerquery(ws):
    wb = ws.Parent
    try:
        qnames = [str(wb.Queries.Item(i).Name) for i in range(1, int(wb.Queries.Count) + 1)]
    except Exception:
        qnames = []
    if not qnames:
        return False, "パワークエリが無い"
    for sh in wb.Worksheets:
        for lo in sh.ListObjects:
            try:
                qt = lo.QueryTable
            except Exception:
                continue
            if qt is None:
                continue
            n = int(lo.ListRows.Count)
            return n == 6, f"クエリ {qnames} → [{sh.Name}] {lo.Name} {va._addr(lo.Range)} {n} 行" + ("" if n == 6 else "（6 行でない）")
    return False, f"クエリ {qnames} はあるが、シートに読み込んだ表が無い"


FIRE_CASES = [
    {"name": "郵便番号の統一",
     "request": "名簿（A3:E9）の郵便番号を 000-0000 のハイフンつきに統一して。値を直接書いて。",
     "check": _check_postal},
    {"name": "会員番号で突き合わせ",
     "request": "左の名簿（A3:E9）と右の申込一覧（I5:K10）を会員番号で突き合わせて、名簿の F 列に「申込コース」の列を足して"
                "（無い人は「申込なし」）。足した後は表全体を整えて。",
     "check": _check_match},
    {"name": "金額の合計",
     "request": "右の申込一覧（I5:K10）の金額の合計を、表のすぐ下の K11 に出して。J11 に「合計」と書いて。",
     "check": _check_sum},
    {"name": "テーブル化",
     "request": "名簿（A3:E9）をテーブルにして。テーブル名は T名簿。",
     "check": _check_table},
    {"name": "ピボット集計",
     "request": "右の申込一覧（I5:K10）から、申込コースごとの金額の合計をピボットで出して。置き場所はこのシートの M5。"
                "作ったら中身を読んで、合計が明細と合っているか確かめて。",
     "check": _check_pivot},
    {"name": "スライサー",
     "request": "右の申込一覧（I5:K10）をテーブル T申込 にして、申込コースで絞れるスライサーを M12 に置いて。",
     "check": _check_slicer},
    {"name": "パワークエリで読み込み",
     "request": "名簿（A3:E9）をテーブル T名簿 にして、パワークエリでそのテーブルを読み込むクエリ Q名簿 を作り、"
                "結果をこのシートの M3 に表として出して。",
     "check": _check_powerquery},
]


# ----------------------------------------------------------------
# ピボットの難しい弾（2026-09-04 深夜）: 明細 45 行（3 か月×5 地域×3 商品）を練習台に、月別グループ・構成比・
# 上位 N・前月比・スライサー 2 枚・ピボットグラフ・データモデルのメジャー を撃つ。答え合わせは現物（Calculation・
# PivotFilters・SlicerCache・PivotLayout・Model）で決める
# ----------------------------------------------------------------
_FIRE_SALES_HEAD = ["日付", "地域", "担当", "商品", "数量", "売上"]
_FIRE_SALES_REGIONS = ["架空県北", "架空県南", "架空県東", "架空県西", "架空県中央"]
_FIRE_SALES_STAFF = {"架空県北": "佐藤", "架空県南": "鈴木", "架空県東": "高橋", "架空県西": "佐藤", "架空県中央": "鈴木"}
_FIRE_SALES_PRICE = {"A商品": 1200, "B商品": 800, "C商品": 1500}


def _fire_sales_rows():
    """明細 45 行（決め打ち＝毎回同じ）。日付は tz=UTC で渡す（素の datetime は pywin32 が 9 時間ずらす）"""
    import datetime as _dt
    rows, k = [], 0
    for m in (4, 5, 6):
        for ri, rg in enumerate(_FIRE_SALES_REGIONS):
            for pi, prod in enumerate(("A商品", "B商品", "C商品")):
                k += 1
                qty = 3 + ((ri * 7 + pi * 5 + m * m + k * 2) % 9)     # 月ごとの合計が同じにならない並び（前月差が 0 で通らないように）
                rows.append([_dt.datetime(2026, m, 1 + (k % 27), tzinfo=_dt.timezone.utc), rg, _FIRE_SALES_STAFF[rg],
                             prod, qty, qty * _FIRE_SALES_PRICE[prod]])
    return rows


def _build_sales(xl, wb, sheet):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    rows = [_FIRE_SALES_HEAD] + _fire_sales_rows()
    ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), 6)).Value = rows
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:F{len(rows)}'], wb)
    return ws


def _build_sales_pivot(xl, wb, sheet):
    """明細 45 行＋**すでに在るピボット**（2026-09-09）。

    それまでのピボットの弾は 8 本とも「白紙から作る」だった。既存のピボットを直す経路
    （pivot_field の add_col/set_func/group_date/remove/position、pivot_calc の subtotals/layout/
    calc_field/set_source/style…）は実装だけあって一度も撃たれていなかった。ここがその練習台。
    置き場所は別シート「<名前>_集計」・行=地域・値=売上の合計・名前は P<名前>。
    """
    ws = _build_sales(xl, wb, sheet)
    ok, out = va._run_cmd(['pivot', 'create', f'{sheet}!A1:F46', '--rows', '地域', '--values', '売上',
                        '--func', 'sum', '--sheet', f'{sheet}_集計', '--at', 'A3', '--name', f'P{sheet}'], wb)
    if not ok:
        raise RuntimeError(f"練習台のピボットを作れませんでした: {out}")
    # 刻印（2026-09-09）。ピボットを消して同じ名前・同じ番地に作り直す 2 手の経路が実在する
    # （pivot delete は TableRange2.Clear() なので、その後の pivot create の空きセル検査が素通りする）。
    # 名前と番地だけ照合しても「直した」と「作り直した」を見分けられないので、どの弾も触らない
    # プロパティに印を付ける。ChangePivotCache（元範囲の変更）でも消えない＝set_source の弾も通る。
    for sh in wb.Worksheets:
        for pt in sh.PivotTables():
            if str(pt.Name) == f'P{sheet}':
                pt.ErrorString = _FIRE_PIVOT_STAMP
                pt.DisplayErrorString = True
    return ws


_FIRE_PIVOT_STAMP = "実射の練習台"   # 練習台ピボットの刻印（ErrorString）。どの弾も触らないプロパティ


def _fire_sales_extra_rows():
    """ピボットを作った**後で**明細に足す 5 行（2026-09-09・元範囲を広げる弾）。

    地域は明細に無い「架空県外」＝範囲を広げずに更新しただけなら、行にも数にも出ない。
    """
    import datetime as _dt
    out = []
    for i in range(5):
        qty = 2 + i
        out.append([_dt.datetime(2026, 7, 1 + i, tzinfo=_dt.timezone.utc), "架空県外", "田中",
                    "A商品", qty, qty * _FIRE_SALES_PRICE["A商品"]])
    return out


def _build_sales_pivot_grown(xl, wb, sheet):
    """明細＋ピボットを作った後で明細に 5 行足す＝ピボットの元範囲（A1:F46）が古いままの練習台。"""
    ws = _build_sales_pivot(xl, wb, sheet)
    extra = _fire_sales_extra_rows()
    ws.Range(ws.Cells(47, 1), ws.Cells(46 + len(extra), 6)).Value = extra
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:F{46 + len(extra)}'], wb)
    return ws


def _fire_pivot_of(ws):
    """この弾の練習台ピボット（_build_sales_pivot が作った P<シート名>）。無ければ (None, 理由)。

    名前だけでなく**刻印**（ErrorString）も見る。消して同じ名前・同じ番地に作り直すと刻印が消えるので、
    「既存を直せ」の弾が「作り直し」で通らない（2026-09-09）。
    """
    hit = None
    for sh, pt in _pivots_of(ws):
        if str(pt.Name) == f'P{ws.Name}':
            hit = pt
            break
    if hit is None:
        names = [f"[{sh.Name}] {pt.Name}" for sh, pt in _pivots_of(ws)]
        return None, f"練習台のピボット P{ws.Name} が無い（ある: {names or 'なし'}）"
    try:
        stamp = str(hit.ErrorString)
    except Exception as ex:
        return None, f"P{ws.Name} の刻印を読めない（{ex}）"
    if stamp != _FIRE_PIVOT_STAMP:
        return None, (f"P{ws.Name} は練習台のピボットではない（刻印 {stamp!r}／要 {_FIRE_PIVOT_STAMP!r}）"
                      "＝消して同じ名前で作り直した。頼まれたのは今あるピボットを直すこと")
    return hit, ''


def _sales_month_totals():
    tot = {}
    for r in _fire_sales_rows():
        tot[r[0].month] = tot.get(r[0].month, 0) + r[5]
    return tot


def _sales_by(col):
    tot = {}
    for r in _fire_sales_rows():
        tot[r[col]] = tot.get(r[col], 0) + r[5]
    return tot


def _pivots_of(ws):
    """この弾のピボット＝元データがこのシート（練習台ブックは弾のあいだ共有なので、前の弾のピボットを見ない）。
    データモデルのピボットは元データがシートでないので、この弾のシートより後ろのシートにあるものを拾う"""
    mine, later = [], []
    idx = int(ws.Index)
    for sh in ws.Parent.Worksheets:
        for pt in sh.PivotTables():
            try:
                src = str(pt.SourceData)
            except Exception:
                src = ''
            if ws.Name in src:
                mine.append((sh, pt))
            elif int(sh.Index) >= idx:
                later.append((sh, pt))
    return mine or later


def _pivot_numbers(pt):
    return [float(v) for row in va._rows_of(pt.TableRange2.Value) for v in row
            if isinstance(v, (int, float)) and not isinstance(v, bool)]


def _data_calcs(pt):
    out = []
    for i in range(1, int(pt.DataFields.Count) + 1):
        d = pt.DataFields.Item(i)
        try:
            out.append((str(d.Name), int(d.Calculation)))
        except Exception:
            out.append((str(d.Name), None))
    return out


def _row_items_with_data(pt, idx=1):
    """画面に出ている行ラベル（見出しと総計を除く）。上位 N の値フィルタは PivotItems.Visible を変えないので、
    出力範囲（RowRange）の実物で数える"""
    rows = va._rows_of(pt.RowRange.Value)
    labels = [str(r[0]) for r in rows if r and r[0] not in (None, '')]
    if len(labels) >= 2:
        labels = labels[1:]                                   # 先頭＝「行ラベル」の見出し
    if labels and (labels[-1].startswith('総計') or labels[-1].lower().startswith('grand total')):
        labels = labels[:-1]
    return labels


def _has_number(nums, want, tol=0.5):
    return any(abs(n - want) <= tol for n in nums)


def _check_pivot_share(ws):
    pts = _pivots_of(ws)
    if not pts:
        return False, "ピボットが無い"
    sh, pt = pts[0]
    where = f"[{sh.Name}] {pt.Name}"
    calcs = _data_calcs(pt)
    months = _row_items_with_data(pt)
    bad = []
    if len(months) != 3:
        bad.append(f"行が月にまとまっていない（データのある行項目 {len(months)} 個: {months[:6]}）")
    if not any(c in (8, 6, 7, 12, 13) for _, c in calcs):
        bad.append(f"構成比（計算の種類）の値フィールドが無い {calcs}")
    if len(calcs) < 2:
        bad.append(f"値フィールドが {len(calcs)} 本（合計と構成比の 2 本が要る）")
    total = sum(_sales_by(1).values())
    if not _has_number(_pivot_numbers(pt), total):
        bad.append(f"売上の総計 {total:,} が出ていない")
    return (not bad), (where + ": 月 3 行・合計＋構成比・総計一致" if not bad else where + ": " + "／".join(bad))


def _check_pivot_top3(ws):
    pts = _pivots_of(ws)
    if not pts:
        return False, "ピボットが無い"
    sh, pt = pts[0]
    where = f"[{sh.Name}] {pt.Name}"
    try:
        f = pt.PivotFields("地域")
    except Exception:
        return False, where + ": 行に「地域」が無い"
    bad = []
    try:
        n_filters = int(f.PivotFilters.Count)
        ftype = int(f.PivotFilters.Item(1).FilterType) if n_filters else None
    except Exception:
        n_filters, ftype = 0, None
    if n_filters != 1 or ftype != 1:                       # xlTopCount=1
        bad.append(f"上位 N の絞り込みが無い（PivotFilters {n_filters} 本・種類 {ftype}）")
    got = _row_items_with_data(pt)
    want = sorted(_sales_by(1).items(), key=lambda kv: -kv[1])[:3]
    want_names = [k for k, _ in want]
    if sorted(got) != sorted(want_names):
        bad.append(f"見えている地域 {got} ≠ 上位 3 {want_names}")
    try:
        order = int(f.AutoSortOrder)
    except Exception:
        order = None
    if order != 2:
        bad.append(f"売上の多い順に並んでいない（AutoSortOrder {order}）")
    elif got != want_names:
        bad.append(f"並び {got} ≠ 多い順 {want_names}")
    return (not bad), (where + f": 上位 3 地域 {want_names} だけ・多い順" if not bad else where + ": " + "／".join(bad))


def _check_pivot_diff(ws):
    pts = _pivots_of(ws)
    if not pts:
        return False, "ピボットが無い"
    sh, pt = pts[0]
    where = f"[{sh.Name}] {pt.Name}"
    bad = []
    months = _row_items_with_data(pt)
    if len(months) != 3:
        bad.append(f"行が月にまとまっていない（{months[:6]}）")
    calcs = _data_calcs(pt)
    if not any(c in (2, 4) for _, c in calcs):
        bad.append(f"前月との差（diff_from）の値フィールドが無い {calcs}")
    mt = _sales_month_totals()
    nums = _pivot_numbers(pt)
    if not _has_number(nums, mt[5] - mt[4]) and not _has_number(nums, mt[6] - mt[5]):
        bad.append(f"差分の値（5月−4月={mt[5] - mt[4]:,} / 6月−5月={mt[6] - mt[5]:,}）が出ていない")
    return (not bad), (where + ": 月 3 行・合計＋前月差・値一致" if not bad else where + ": " + "／".join(bad))


def _check_two_slicers(ws):
    pts = [(sh, pt) for sh, pt in _pivots_of(ws) if sh.Name == ws.Name]
    if not pts:
        return False, "このシートにピボットが無い"
    sh, pt = pts[0]
    wb = ws.Parent
    fields = {}
    for sc in wb.SlicerCaches:
        try:
            linked = [str(sc.PivotTables.Item(i).Name) for i in range(1, int(sc.PivotTables.Count) + 1)]
        except Exception:
            linked = []
        fields[str(sc.SourceName)] = linked
    bad = []
    for want in ("担当", "商品"):
        if want not in fields:
            bad.append(f"{want} のスライサーが無い")
        elif pt.Name not in fields[want]:
            bad.append(f"{want} のスライサーがピボット {pt.Name} につながっていない（{fields[want]}）")
    return (not bad), (f"{pt.Name}: 担当・商品の 2 枚が同じピボットに連動" if not bad else "／".join(bad) + f"（ある: {fields}）")


def _check_pivot_chart(ws):
    pts = [(sh, pt) for sh, pt in _pivots_of(ws) if sh.Name == ws.Name]
    if not pts:
        return False, "このシートにピボットが無い"
    found = []
    for co in ws.ChartObjects():
        try:
            pl = co.Chart.PivotLayout
            is_pivot = pl is not None and str(pl.PivotTable.Name) != ''
        except Exception:
            is_pivot = False
        try:
            title = str(co.Chart.ChartTitle.Text) if co.Chart.HasTitle else ''
        except Exception:
            title = ''
        found.append((str(co.Name), is_pivot, title))
    ok = any(p and '商品別' in t for _, p, t in found)
    return ok, (f"ピボットグラフ {found}" if ok else f"ピボットグラフ（PivotLayout 付き・題「商品別…」）が無い {found}")


def _check_model_pivot(ws):
    wb = ws.Parent
    bad = []
    try:
        model = wb.Model
        tables = [str(model.ModelTables.Item(i).Name) for i in range(1, int(model.ModelTables.Count) + 1)]
        measures = [str(model.ModelMeasures.Item(i).Name) for i in range(1, int(model.ModelMeasures.Count) + 1)]
    except Exception as ex:
        return False, f"データモデルを読めない: {ex}"
    if not tables:
        bad.append("データモデルにテーブルが無い")
    if "平均単価" not in measures:
        bad.append(f"メジャー「平均単価」が無い {measures}")
    olap = [(sh, pt) for sh, pt in _pivots_of(ws) if getattr(pt.PivotCache(), 'OLAP', False)]
    if not olap:
        bad.append("データモデルから作ったピボットが無い")
    else:
        sh, pt = olap[0]
        names = [n for n, _ in _data_calcs(pt)]
        if not any("平均単価" in n for n in names):
            bad.append(f"ピボットの値に平均単価が無い {names}")
        by_region = {}
        for r in _fire_sales_rows():
            a = by_region.setdefault(r[1], [0, 0])
            a[0] += r[5]; a[1] += r[4]
        want = by_region["架空県北"][0] / by_region["架空県北"][1]
        if not _has_number(_pivot_numbers(pt), want, tol=0.01):
            bad.append(f"架空県北の平均単価 {want:.2f} が出ていない")
    return (not bad), ("データモデル " + ",".join(tables) + " ・メジャー 平均単価・ピボットの値一致" if not bad else "／".join(bad))


# ----------------------------------------------------------------
# 既存のピボットを「直す」弾（2026-09-09）。
# それまでのピボットの弾 8 本は全部「白紙から作る」で、pivot_field / pivot_calc のほとんどの手
# （add_col・set_func・set_name・group_date を既存に当てる・subtotals・grand_totals・layout・
#  repeat_labels・calc_field・set_source・set_filter・group_numeric・empty_as・style・remove・position）
# は実装だけあって一度も撃たれていなかった。答え合わせは全部 COM の現物（vbam_heavy がどのプロパティを
# 書いているかに合わせてある）＋ 明細 45 行から計算した数字。作り直し・値の貼り付け・古いままの更新は落ちる。
# ----------------------------------------------------------------
_XL_ORIENT = {0: '外', 1: '行', 2: '列', 3: 'フィルタ', 4: '値'}


def _pf_of(pt, name):
    try:
        return pt.PivotFields(name)
    except Exception:
        return None


def _orient_of(pt, name):
    p = _pf_of(pt, name)
    if p is None:
        return None
    try:
        return int(p.Orientation)
    except Exception:
        return None


def _fields_by_orient(pt, orient):
    """その置き場所にあるフィールド名を Position の順で。読めないものは飛ばす。"""
    out = []
    try:
        for f in pt.PivotFields():
            try:
                if int(f.Orientation) != orient:
                    continue
            except Exception:
                continue
            try:
                pos = int(f.Position)
            except Exception:
                pos = 99
            out.append((pos, str(f.Name)))
    except Exception:
        return []
    return [n for _, n in sorted(out)]


def _visible_items(pt, field):
    p = _pf_of(pt, field)
    if p is None:
        return None
    out = []
    try:
        for i in range(1, int(p.PivotItems().Count) + 1):
            it = p.PivotItems().Item(i)
            try:
                if bool(it.Visible):
                    out.append(str(it.Name))
            except Exception:
                continue
    except Exception:
        return None
    return out


def _sales_by_staff():
    tot = {}
    for r in _fire_sales_rows():
        tot[r[2]] = tot.get(r[2], 0) + r[5]
    return tot


def _sales_avg_by_region():
    tot, cnt = {}, {}
    for r in _fire_sales_rows():
        tot[r[1]] = tot.get(r[1], 0) + r[5]
        cnt[r[1]] = cnt.get(r[1], 0) + 1
    return {k: tot[k] / cnt[k] for k in tot}


def _check_pv_addcol(ws):
    """列を足す＋値に桁区切り（pivot_field add_col / set_format）。

    書式は値フィールドの NumberFormat を見る（vbam_heavy の set-format が書くのはそこ）。
    列ごとに読むとピボットはラベル行のぶん必ず「まちまち」になるので読まない。
    """
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad, rows, cols = [], _fields_by_orient(pt, 1), _fields_by_orient(pt, 2)
    if '地域' not in rows:
        bad.append(f"行の「地域」が消えている＝直さずに作り直した疑い（行={rows}）")
    if '商品' not in cols:
        bad.append(f"列に「商品」が無い（列={cols}）")
    try:
        fmt = str(pt.DataFields.Item(1).NumberFormat)
    except Exception as ex:
        fmt = f"読めない({ex})"
    if '#,##0' not in fmt and '#,###' not in fmt:
        bad.append(f"値に桁区切りが付いていない（NumberFormat={fmt}）")
    nums = _pivot_numbers(pt)
    for prod, want in _sales_by(3).items():
        if not _has_number(nums, want):
            bad.append(f"{prod} の合計 {want:,} が出ていない")
    return (not bad), (f"{pt.Name}: 行={rows}・列={cols}・値の書式 {fmt}・商品別の合計が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_average(ws):
    """合計→平均に変える＋表示名（pivot_field set_func / set_name）。Function は -4106（xlAverage）。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    try:
        d = pt.DataFields.Item(1)
        func, name = int(d.Function), str(d.Name)
    except Exception as ex:
        return False, f"{pt.Name}: 値フィールドを読めない（{ex}）"
    if func != -4106:
        bad.append(f"平均になっていない（Function={func}／平均は -4106・合計は -4157）")
    if '平均' not in name:
        bad.append(f"見出しが「平均売上」になっていない（{name}）")
    if '地域' not in _fields_by_orient(pt, 1):
        bad.append(f"行の「地域」が消えている＝作り直した疑い（行={_fields_by_orient(pt, 1)}）")
    nums = _pivot_numbers(pt)
    for rg, want in _sales_avg_by_region().items():
        if not _has_number(nums, want, tol=0.01):
            bad.append(f"{rg} の平均 {want:,.1f} が出ていない（合計のままの疑い）")
    return (not bad), (f"{pt.Name}: 値=平均「{name}」・地域別の平均が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_regroup_month(ws):
    """既存ピボットの行を地域→月に組み替える（add_row 日付＋group_date months＋地域を外す）。

    まとめられていなければ行ラベルは 45 個（日付のまま）になる＝3 個かどうかで見分ける。
    """
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    labels = _row_items_with_data(pt)
    if len(labels) != 3:
        bad.append(f"行が月 3 つになっていない（{len(labels)} 個: {labels[:6]}）"
                   "＝日付のままなら 45 個、まとめ方が違えば別の数になる")
    o = _orient_of(pt, '地域')
    if o not in (0, None):
        bad.append(f"「地域」が行から外れていない（置き場所={_XL_ORIENT.get(o, o)}）")
    # 逃げ道つぶし: 明細の日付を「2026年4月」のような文字に書き換えても行は 3 個になる。
    # それはグループ化ではなく元データを潰した回なので落とす。
    try:
        texts = [str(r[0]) for r in va._rows_of(ws.Range("A2:A46").Value) if isinstance(r[0], str)]
    except Exception:
        texts = []
    if texts:
        bad.append(f"明細の日付が文字に書き換えられている（{len(texts)} セル・例 {texts[0]!r}）"
                   "＝ピボットのグループ化ではなく元データを潰している")
    nums = _pivot_numbers(pt)
    for m, want in _sales_month_totals().items():
        if not _has_number(nums, want):
            bad.append(f"{m} 月の合計 {want:,} が出ていない")
    return (not bad), (f"{pt.Name}: 行={labels}・月別の合計が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_subtotals(ws):
    """小計と行の総計を消す（pivot_calc subtotals / grand_totals）。

    Subtotals は 12 個の真偽の並び（先頭＝自動小計）、行の総計は pt.RowGrand。画面の文字では見ない。
    """
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad, rows, cols = [], _fields_by_orient(pt, 1), _fields_by_orient(pt, 2)
    for need in ('地域', '担当'):
        if need not in rows:
            bad.append(f"行に「{need}」が無い（行={rows}）")
    # 列フィールドが無いと「右端の列の総計」が見た目に存在せず、rows と cols を取り違えても
    # 画面が同じになる＝取り違えを試験できない。列を 1 本置かせて初めて意味のある弾になる（2026-09-09）
    if '商品' not in cols:
        bad.append(f"列に「商品」が無い（列={cols}）")
    # Excel の名前は見た目と逆向き（2026-09-09 実測）:
    #   ColumnGrand ＝「列ごとの総計」＝**一番下の行**  ／  RowGrand ＝「行ごとの総計」＝**右端の列**
    # 依頼は「一番下の行の総計は消して、右端の列の総計は残して」＝ColumnGrand False・RowGrand True
    try:
        cgrand = bool(pt.ColumnGrand)
    except Exception as ex:
        cgrand = None
        bad.append(f"ColumnGrand を読めない（{ex}）")
    if cgrand is True:
        bad.append("一番下の行の総計が消えていない（ColumnGrand=True）")
    # 逃げ道つぶし: which=both で両方消すと「総計が消えた」画面は手に入るが、残せと言われた列まで失う
    try:
        grand = bool(pt.RowGrand)
    except Exception:
        grand = None
    if grand is False:
        bad.append("右端の列の総計まで消している（RowGrand=False／消すのは一番下の行だけ）")
    p = _pf_of(pt, '地域')
    try:
        sub = [bool(x) for x in p.Subtotals]
    except Exception as ex:
        sub = []
        bad.append(f"「地域」の Subtotals を読めない（{ex}）")
    if sub and any(sub):
        bad.append(f"「地域」の小計が残っている（Subtotals={sub}）")
    # 総計を消す弾なので、総計の数字では確かめられない（消えるのが正しい）。
    # どちらの総計が残っても必ず出る「地域×商品」の中身で、数字が壊れていないことを見る
    nums, cell = _pivot_numbers(pt), {}
    for r in _fire_sales_rows():
        cell[(r[1], r[3])] = cell.get((r[1], r[3]), 0) + r[5]
    miss = [f"{rg}×{prod} の {want:,}" for (rg, prod), want in cell.items() if not _has_number(nums, want)]
    if miss:
        bad.append(f"中身が明細と合っていない（出ていない: {'・'.join(miss[:3])}／全 {len(cell)} 通り）")
    return (not bad), (f"{pt.Name}: 行={rows}・列={cols}・小計オフ・RowGrand={grand}・ColumnGrand={cgrand}・"
                       f"地域×商品 {len(cell)} 通りが明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_layout(ws):
    """表形式レイアウト＋ラベルの繰り返し（pivot_calc layout tabular / repeat_labels）。

    RowAxisLayout(1) が書くのは各行フィールドの LayoutForm=1（xlTabular）、
    RepeatAllLabels(2) が書くのは RepeatLabels=True。
    """
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad, rows = [], _fields_by_orient(pt, 1)
    if '担当' not in rows:
        bad.append(f"行に「担当」が無い（行={rows}）")
    if '地域' not in rows:
        bad.append(f"行の「地域」が消えている＝作り直した疑い（行={rows}）")
    p = _pf_of(pt, '地域')
    form = rep = None
    try:
        form = int(p.LayoutForm)
    except Exception as ex:
        bad.append(f"LayoutForm を読めない（{ex}）")
    # XlLayoutFormType は xlTabular=0 / xlOutline=1（RowAxisLayout の xlTabularRow=1 とは別の並び。
    # 2026-09-09 の実射で実測: 表形式にすると LayoutForm=0・LayoutCompactRow=False になる）
    if form is not None and form != 0:
        bad.append(f"表形式になっていない（LayoutForm={form}／表形式は 0・アウトラインは 1）")
    try:
        rep = bool(p.RepeatLabels)
    except Exception as ex:
        bad.append(f"RepeatLabels を読めない（{ex}）")
    if rep is False:
        bad.append("行ラベルが繰り返されていない（RepeatLabels=False）")
    # 逃げ道つぶし: LayoutRowDefault（これから足す分の既定）だけ変えても既存フィールドは compact のまま
    try:
        compact = bool(p.LayoutCompactRow)
    except Exception:
        compact = None
    if compact is True:
        bad.append("コンパクト形式のまま（LayoutCompactRow=True／既定だけ変えて既存に当てていない疑い）")
    nums = _pivot_numbers(pt)
    for rg, want in _sales_by(1).items():
        if not _has_number(nums, want):
            bad.append(f"{rg} の合計 {want:,} が出ていない")
    return (not bad), (f"{pt.Name}: 行={rows}・LayoutForm={form}・RepeatLabels={rep}・合計が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_calcfield(ws):
    """計算フィールド（pivot_calc calc_field_create）。名前だけでなく 1.1 倍の数字が出ているかまで見る。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    names = []
    try:
        cfs = pt.CalculatedFields()
        names = [str(cfs.Item(i).Name) for i in range(1, int(cfs.Count) + 1)]
    except Exception as ex:
        bad.append(f"CalculatedFields を読めない（{ex}）")
    if not any('税込' in n for n in names):
        bad.append(f"計算フィールド「税込売上」が無い（{names}）")
    shown = [n for n, _ in _data_calcs(pt)]
    if not any('税込' in n for n in shown):
        bad.append(f"税込の値がピボットに出ていない（値={shown}）")
    total = sum(_sales_by(1).values())
    nums = _pivot_numbers(pt)
    if not _has_number(nums, total * 1.1, tol=1.0):
        bad.append(f"税込の総計 {total * 1.1:,.0f} が出ていない（1.1 倍になっていない）")
    if not _has_number(nums, total):
        bad.append(f"もとの売上の総計 {total:,} が消えている")
    return (not bad), (f"{pt.Name}: 計算フィールド {names}・値={shown}・税込 {total * 1.1:,.0f} が一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_setsource(ws):
    """元範囲を広げて更新（pivot_calc set_source / refresh）。

    足した 5 行は明細に無い地域「架空県外」なので、広げずに更新しただけなら行にも数にも出ない
    ＝「更新しました」と言うだけの回が落ちる。
    """
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    try:
        src = str(pt.SourceData)
    except Exception as ex:
        src = f"読めない({ex})"
    last = 46 + len(_fire_sales_extra_rows())
    if str(last) not in src:
        bad.append(f"元の範囲が {last} 行目まで広がっていない（SourceData={src}）")
    labels = _row_items_with_data(pt)
    if '架空県外' not in labels:
        bad.append(f"足した行の地域「架空県外」が出ていない（行={labels}）")
    extra = sum(r[5] for r in _fire_sales_extra_rows())
    want = sum(_sales_by(1).values()) + extra
    nums = _pivot_numbers(pt)
    if not _has_number(nums, want):
        bad.append(f"総計が足した行を含んでいない（要 {want:,}／足す前は {want - extra:,}）")
    return (not bad), (f"{pt.Name}: 元={src}・行に架空県外・総計 {want:,} が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_filter_staff(ws):
    """担当で絞る（pivot_field add_filter＋set_filter）。PivotItems.Visible の現物で見る。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    o = _orient_of(pt, '担当')
    if o is None or o == 0:
        bad.append(f"「担当」がピボットに置かれていない（置き場所={_XL_ORIENT.get(o, o)}）")
    # 表示状態はフィルタ（ページ）位置だと読めないことがある（複数選択のとき PivotItems.Visible が
    # 全部 False で返る・2026-09-09 実測）。読めたら見るが、絞れているかの判定は数字を正本にする
    vis = _visible_items(pt, '担当')
    cur = None
    if o == 3:
        try:
            cur = str(pt.PivotFields('担当').CurrentPage)
        except Exception:
            cur = None
    if vis and sorted(vis) != ['鈴木']:
        bad.append(f"鈴木さんだけになっていない（見えている担当={vis}）")
    want = _sales_by_staff().get('鈴木', 0)
    nums = _pivot_numbers(pt)
    if not _has_number(nums, want):
        bad.append(f"鈴木さんの売上 {want:,} が出ていない")
    whole = sum(_sales_by(1).values())
    if _has_number(nums, whole):
        bad.append(f"全体の総計 {whole:,} がまだ出ている＝絞れていない")
    return (not bad), (f"{pt.Name}: 担当={_XL_ORIENT.get(o, o)}・見えている担当={vis or cur}・"
                       f"鈴木 {want:,} だけが出ている（全体 {whole:,} は出ていない）"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_group_qty(ws):
    """数量を区分でまとめる（pivot_field group_numeric）。区分名は「3-5」のような範囲になる。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    o = _orient_of(pt, '数量')
    if o != 1:
        bad.append(f"「数量」が行に無い（置き場所={_XL_ORIENT.get(o, o)}）")
    labels = _row_items_with_data(pt)
    banded = [s for s in labels if '-' in s or '〜' in s]
    if not banded:
        bad.append(f"区分（「3-5」のような範囲）になっていない（行={labels[:8]}）")
    qtys = sorted({r[4] for r in _fire_sales_rows()})
    if len(labels) >= len(qtys):
        bad.append(f"まとまっていない（行 {len(labels)} 個・数量の種類 {len(qtys)} 個）")
    # 逃げ道つぶし: 明細に「3-5」等の帯の列を足して元範囲を広げると、行も数字も本物と同じに見える。
    # 元データが 6 列 46 行のままであることを見る（列を足した瞬間 C7 になる）
    try:
        src = str(pt.SourceData).replace(' ', '')
    except Exception:
        src = ''
    if src and 'R46C6' not in src:
        bad.append(f"元データが 6 列 46 行から変わっている＝明細に区分の列を足した疑い（{src}）")
    total = sum(_sales_by(1).values())
    nums = _pivot_numbers(pt)
    if not _has_number(nums, total):
        bad.append(f"総計 {total:,} が出ていない")
    return (not bad), (f"{pt.Name}: 行={labels}・区分 {len(banded)} 個・総計 {total:,} が一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_nullstyle(ws):
    """空欄の表示とスタイル（pivot_calc empty_as / style）。NullString・DisplayNullString・TableStyle2。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad = []
    try:
        disp, null = bool(pt.DisplayNullString), str(pt.NullString)
    except Exception as ex:
        disp, null = None, f"読めない({ex})"
        bad.append(f"NullString を読めない（{ex}）")
    if disp is False:
        bad.append("空欄に文字を出す設定になっていない（DisplayNullString=False）")
    if null != '0':
        bad.append(f"空欄に出す文字が 0 でない（NullString={null!r}）")
    try:
        style = str(pt.TableStyle2)
    except Exception as ex:
        style = f"読めない({ex})"
    if 'PivotStyleMedium9' not in style:
        bad.append(f"スタイルが PivotStyleMedium9 でない（TableStyle2={style}）")
    total = sum(_sales_by(1).values())
    if not _has_number(_pivot_numbers(pt), total):
        bad.append(f"総計 {total:,} が消えている")
    return (not bad), (f"{pt.Name}: NullString={null!r}・DisplayNullString={disp}・{style}・総計一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_remove_position(ws):
    """フィールドを外して並べ替える（pivot_field remove / position）。Orientation と Position の現物。"""
    pt, why = _fire_pivot_of(ws)
    if pt is None:
        return False, why
    bad, rows = [], _fields_by_orient(pt, 1)
    o = _orient_of(pt, '商品')
    if o not in (0, None):
        bad.append(f"「商品」が外れていない（置き場所={_XL_ORIENT.get(o, o)}）")
    if rows[:2] != ['担当', '地域']:
        bad.append(f"行が「担当→地域」の順になっていない（行={rows}）")
    nums = _pivot_numbers(pt)
    for st, want in _sales_by_staff().items():
        if not _has_number(nums, want):
            bad.append(f"{st} の合計 {want:,} が出ていない")
    return (not bad), (f"{pt.Name}: 行={rows}・商品は外れている・担当別の合計が明細と一致"
                       if not bad else f"{pt.Name}: " + "／".join(bad))


def _check_pv_no_word(ws):
    """診断の弾: 「ピボット」と言わずに月×担当の集計を頼む。

    規則 3「導ける値は必ず数式で」がピボットから遠ざける向きに効くので、道具がどちらを選ぶかを記録する。
    ピボットでも数式でも合格、値を直に書いたら不合格。どちらを選んだかは答え合わせの文に残す。
    """
    month_tot = _sales_month_totals()
    for sh, pt in _pivots_of(ws):
        nums = _pivot_numbers(pt)
        if all(_has_number(nums, v) for v in month_tot.values()):
            return True, (f"ピボットを選んだ（[{sh.Name}] {pt.Name}）: 月別の合計 "
                          + "・".join(f"{m}月={v:,}" for m, v in sorted(month_tot.items())) + " が一致")
    idx, nums, fcount, where = int(ws.Index), [], 0, []
    for sh in ws.Parent.Worksheets:
        if int(sh.Index) < idx:
            continue
        try:
            for row in va._rows_of(sh.UsedRange.Value):
                for v in row:
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        nums.append(float(v))
        except Exception:
            continue
        try:
            n = int(sh.UsedRange.SpecialCells(-4123).Count)     # xlCellTypeFormulas
        except Exception:
            n = 0
        if n:
            fcount += n
            where.append(f"{sh.Name}({n})")
    hit = all(_has_number(nums, v) for v in month_tot.values())
    if hit and fcount:
        return True, f"数式を選んだ（数式 {fcount} セル: {'・'.join(where)}）: 月別の合計が一致"
    if hit:
        return False, "月別の数字は合っているが、数式でもピボットでもない＝値が直に書かれている"
    return False, ("月別の合計（" + "・".join(f"{m}月={v:,}" for m, v in sorted(month_tot.items()))
                   + f"）がどこにも出ていない（数式 {fcount} セル）")


_FIRE_PIVOT_EDIT_CASES = [
    {"name": "ピボットに列を足す", "build": _build_sales_pivot,
     "request": "集計シートのピボット、商品別も見たいので列に足して。金額は3桁区切りにして。",
     "check": _check_pv_addcol},
    {"name": "ピボットを平均に変える", "build": _build_sales_pivot,
     "request": "集計シートのピボット、売上を合計じゃなく平均にして。見出しは「平均売上」に変えて。",
     "check": _check_pv_average},
    {"name": "ピボットを月別に組み替える", "build": _build_sales_pivot,
     "request": "集計シートのピボット、地域じゃなくて月ごとの売上にして。地域は行から外していい。",
     "check": _check_pv_regroup_month},
    {"name": "ピボットの小計と総計を消す", "build": _build_sales_pivot,
     "request": "集計シートのピボット、行に担当も足して、列に商品も足して。地域ごとの小計は要らないので消して。"
                "一番下の行の総計は消して、右端の列の総計は残して。",
     "check": _check_pv_subtotals},
    {"name": "ピボットを表形式にする", "build": _build_sales_pivot,
     "request": "集計シートのピボット、行に担当も足して、レイアウトを表形式にして、行のラベルは各行に繰り返して出して。",
     "check": _check_pv_layout},
    {"name": "ピボットに税込を足す", "build": _build_sales_pivot,
     "request": "集計シートのピボットに税込の値も足して。売上の1.1倍で、名前は「税込売上」にして。",
     "check": _check_pv_calcfield},
    {"name": "ピボットの元範囲を広げる", "build": _build_sales_pivot_grown,
     "request": "明細に行を足したのに集計に反映されていない。ピボットの元の範囲を今の明細全部に広げて、最新にして。",
     "check": _check_pv_setsource},
    {"name": "ピボットを担当で絞る", "build": _build_sales_pivot,
     "request": "集計シートのピボット、担当で絞れるようにして、鈴木さんの分だけ出して。",
     "check": _check_pv_filter_staff},
    {"name": "ピボットを数量の区分でまとめる", "build": _build_sales_pivot,
     "request": "集計シートのピボット、行を数量の区分（3つ刻み）にして、区分ごとの売上合計を見られるようにして。",
     "check": _check_pv_group_qty},
    {"name": "ピボットの空欄と見た目", "build": _build_sales_pivot,
     "request": "集計シートのピボット、空欄には0と出るようにして、スタイルは PivotStyleMedium9 にして。",
     "check": _check_pv_nullstyle},
    {"name": "ピボットのフィールドを外して並べ替える", "build": _build_sales_pivot,
     "request": "集計シートのピボット、行に担当と商品を足して。商品はやっぱり要らないので外して。担当を地域より先に出して。",
     "check": _check_pv_remove_position},
    {"name": "ピボットと言わずに集計", "build": _build_sales,
     "request": "この明細、月ごとに担当別の売上を並べて見たいんだけど。",
     "check": _check_pv_no_word},
]


_FIRE_PIVOT_CASES = [
    {"name": "月別×担当の構成比", "build": _build_sales,
     "request": "明細（A1:F46）から、日付を月でまとめた行・担当を列にして、売上の合計と、総計に対する構成比（％）の 2 本を値に"
                "並べたピボットを別シート「月別」に作って。作ったら中身を読んで、売上の総計が明細の合計と合っているか確かめて。",
     "check": _check_pivot_share},
    {"name": "売上上位3地域", "build": _build_sales,
     "request": "明細（A1:F46）から地域ごとの売上合計のピボットを別シート「上位」に作り、売上が多い上位 3 地域だけが出るように"
                "絞って、売上の多い順に並べて。",
     "check": _check_pivot_top3},
    {"name": "前月比", "build": _build_sales,
     "request": "明細（A1:F46）から月ごとの売上合計のピボットを別シート「推移」に作り、売上合計の右に前月との差"
                "（前の月に対する差分）の列を足して。",
     "check": _check_pivot_diff},
    {"name": "スライサー2枚連動", "build": _build_sales,
     "request": "明細（A1:F46）から地域を行・商品を列にした売上合計のピボットをこのシートの H2 に作り、担当と商品で絞れる"
                "スライサーを 2 枚、H12 と L12 に置いて（どちらも同じピボットにつなぐ）。",
     "check": _check_two_slicers},
    {"name": "ピボットグラフ", "build": _build_sales,
     "request": "明細（A1:F46）から商品ごとの売上合計のピボットをこのシートの H2 に作り、そのピボットからピボットグラフ（棒）を"
                " H10 に置いて。タイトルは「商品別 売上（円）」。",
     "check": _check_pivot_chart},
    {"name": "データモデルの平均単価", "build": _build_sales, "max_turns": 6,
     "request": "明細（A1:F46）をテーブル T売上 にし、パワークエリでデータモデルに載せて、メジャー「平均単価」（売上の合計 ÷ 数量の合計）"
                "を作り、データモデルから地域を行・平均単価を値にしたピボットを別シート「単価」に作って。",
     "check": _check_model_pivot},
]
FIRE_CASES += _FIRE_PIVOT_CASES
FIRE_CASES += _FIRE_PIVOT_EDIT_CASES      # 既存のピボットを直す 12 本（2026-09-09）


# ----------------------------------------------------------------
# 大きい表の弾（2026-09-04）: 300 行＝AI が全行を JSON で書けない大きさ。normalize / fill が通ることを現物で確かめる
# ----------------------------------------------------------------
_BIG_ROWS = 300
_BIG_NAMES = ["青木 誠", "石川 恵", "上田 学", "遠藤 香", "大西 亮", "加藤 環"]


def _to_zenkaku(s):
    return ''.join(chr(ord(ch) + 0xFEE0) if 0x21 <= ord(ch) <= 0x7E else ch for ch in s)


def _big_dirty_rows():
    """番号（先頭ゼロ）・氏名（一部に前後の空白）・郵便番号（4 通りの乱れ）・金額（数値／文字／全角の文字）。決め打ち。"""
    rows = []
    for i in range(1, _BIG_ROWS + 1):
        name = _BIG_NAMES[i % 6]
        if i % 7 == 0:
            name = " " + name + " "
        k = i % 4
        base = f"{i % 10:03d}-{i:04d}"
        postal = (base if k == 0 else base.replace('-', '') if k == 1 else _to_zenkaku(base) if k == 2 else f" {base} ")
        amt = 1000 * (i % 20 + 1)
        m = i % 3
        amount = amt if m == 0 else (f"{amt:,}" if m == 1 else _to_zenkaku(f"{amt:,}"))
        rows.append([f"{i:04d}", name, postal, amount])
    return rows


def _build_big_dirty(xl, wb, sheet):
    from vbam_recipes import _put, _text_cell
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    rows = _big_dirty_rows()
    _put(ws, 1, 1, [["番号", "氏名", "郵便番号", "金額"]] + rows, text_cols=(0, 2))
    for i, r in enumerate(rows, 2):
        if isinstance(r[3], str):
            _text_cell(ws, f"D{i}", r[3])
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:D{_BIG_ROWS + 1}'], wb)
    return ws


def _check_big_clean(ws):
    vals = va._rows_of(ws.Range(f"A2:D{_BIG_ROWS + 1}").Value)
    bad_c = [f"C{i + 2}={va._text_of(r[2])!r}" for i, r in enumerate(vals) if not re.fullmatch(r'\d{3}-\d{4}', str(r[2] or ''))]
    bad_d = [f"D{i + 2}={va._text_of(r[3])!r}" for i, r in enumerate(vals)
             if not (isinstance(r[3], (int, float)) and not isinstance(r[3], bool))]
    bad = []
    if bad_c:
        bad.append(f"郵便番号が 000-0000 でない {len(bad_c)} 行: " + " ".join(bad_c[:4]))
    if bad_d:
        bad.append(f"金額が数値でない {len(bad_d)} 行: " + " ".join(bad_d[:4]))
    if va._text_of(ws.Range("A2").Value) != "0001":
        bad.append(f"番号の先頭ゼロが消えた（A2={va._text_of(ws.Range('A2').Value)!r}）")
    if ws.Range(f"A{_BIG_ROWS + 2}").Value not in (None, ''):
        bad.append("表の下に行が増えた")
    return (not bad), (f"{_BIG_ROWS} 行すべて 郵便番号 000-0000・金額は数値・先頭ゼロ保持" if not bad else "／".join(bad))


def _big_match_right():
    """右の申込一覧（55 行）。会員番号は左の 5 人に 1 人＋左に無い番号 5 つ。"""
    rows = [["会員番号", "申込コース", "金額"]]
    course = ["年間", "半年", "月額"]
    price = {"年間": 12000, "半年": 7000, "月額": 1200}
    k = 0
    for i in range(1, _BIG_ROWS + 1):
        if i % 6 == 0:
            c = course[k % 3]
            rows.append([f"{2000 + i}", c, price[c]])
            k += 1
    for j in range(5):
        rows.append([f"{2900 + j}", "年間", 12000])
    return rows


def _build_big_match(xl, wb, sheet):
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    left = [["会員番号", "氏名"]] + [[f"{2000 + i}", _BIG_NAMES[i % 6]] for i in range(1, _BIG_ROWS + 1)]
    _put(ws, 1, 1, left, text_cols=(0,))
    right = _big_match_right()
    _put(ws, 1, 9, right, text_cols=(0,))
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:B{_BIG_ROWS + 1}', f'I1:K{len(right)}'], wb)
    return ws


def _check_big_match(ws):
    want = {r[0]: r[1] for r in _big_match_right()[1:]}
    vals = va._rows_of(ws.Range(f"A2:C{_BIG_ROWS + 1}").Value)
    bad = [f"C{i + 2}={va._text_of(r[2])!r}（期待 {want.get(str(r[0]), '申込なし')}）" for i, r in enumerate(vals)
           if va._text_of(r[2]) != want.get(va._text_of(r[0]), "申込なし")]
    head = va._text_of(ws.Range("C1").Value)
    note = f"C 列 {_BIG_ROWS} 行の突き合わせ（申込なし含む）"
    if head != "申込コース":
        bad.append(f"C1 の見出しが「{head}」")
    n_formula = len([1 for row in va._rows_of(ws.Range(f'C2:C{_BIG_ROWS + 1}').Formula) for f in row
                     if isinstance(f, str) and f.startswith('=')])
    return (not bad), (note + f"・数式 {n_formula} 本" if not bad else f"違う {len(bad)} 行: " + " ".join(bad[:4]))


_FIRE_BIG_CASES = [
    {"name": "300行の掃除", "build": _build_big_dirty,
     "request": f"名簿（A1:D{_BIG_ROWS + 1}）の郵便番号（C 列）を 000-0000 の形に統一し、金額（D 列）を数値にして。"
                "元の列を置き換えてよい。番号（A 列）は文字のまま触らない。",
     "check": _check_big_clean},
    {"name": "300行の突き合わせ", "build": _build_big_match,
     "request": f"左の名簿（A1:B{_BIG_ROWS + 1}）の C 列に「申込コース」の列を足して。右の申込一覧（I1:K{len(_big_match_right())}）を"
                "会員番号で突き合わせて、無い人は「申込なし」。足した後は名簿の表全体を整えて。",
     "check": _check_big_match},
]
FIRE_CASES += _FIRE_BIG_CASES


# ----------------------------------------------------------------
# 曖昧な依頼の弾（2026-09-04）: 「見やすくして」「合計を出して」のように仕様が書かれていない依頼。
#   ここが Copilot / Claude for Excel に負けている所＝報告で止まらず、一つの答えを出せるかを測る。
#   答え合わせは全部、実物の機械判定（見た目は audit_table・値はセルの実物）。人の主観は入れない。
# ----------------------------------------------------------------
_VAGUE_ROWS = [
    ["0001", "青木 誠", "佐藤", 1250000, "年会費"],
    ["0002", "石川 恵", "鈴木", 980000, ""],
    ["0003", "上田 学", "佐藤", 1430000, "分割"],
    ["0004", "遠藤 香", "高橋", 760000, ""],
    ["0005", "大西 亮", "鈴木", 2100000, "年会費"],
    ["0006", "加藤 環", "佐藤", 540000, ""],
    ["0007", "木村 悠", "高橋", 1875000, "分割"],
    ["0008", "工藤 望", "鈴木", 320000, "年会費"],
]
_VAGUE_HEAD = ["会員番号", "氏名", "担当", "金額", "備考"]
_VAGUE_LAST = len(_VAGUE_ROWS) + 1            # A1:E9
_VAGUE_TOTAL = sum(r[3] for r in _VAGUE_ROWS)


_STATE_SHEET_PREFIX = '前任者の集計'          # 練習台の状態が置くシート（弾の答え合わせでは見ない）


def _case_sheets(ws):
    """この弾が触ってよいシート＝自分の練習台と、AI がこの回に作ったシート。

    実射は 1 冊のブックに 実射1・実射2 … と並べるので、答え合わせで全シートを見ると
    他の弾の練習台を拾ってしまう（2026-09-04・「鈴木さんの分だけ」が他の弾の名簿を混入と誤判定）。
    状態「他シート参照」が置くシートも同じ理由で外す（2026-09-05）。
    """
    out = [ws]
    for sh in ws.Parent.Worksheets:
        nm = str(sh.Name)
        if nm != str(ws.Name) and not re.fullmatch(r'実射\d+', nm) and not nm.startswith(_STATE_SHEET_PREFIX):
            out.append(sh)
    return out


def _build_vague(xl, wb, sheet):
    """素の表（書式なし・罫線なし・見出しも本文と同じ・金額の列は幅 4 で ###）。tidy を当てずに置く。"""
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    _put(ws, 1, 1, [_VAGUE_HEAD] + _VAGUE_ROWS, text_cols=(0,))
    for col, w in (("A", 6), ("B", 10), ("C", 6), ("D", 4), ("E", 6)):
        ws.Columns(col).ColumnWidth = w        # D は幅 4＝金額が ### になる
    ws.Range(f"A1:E{_VAGUE_LAST}").Borders.LineStyle = -4142     # xlLineStyleNone（罫線なし）
    ws.Range("A1:E1").Font.Bold = False
    ws.Range("A1:E1").Interior.ColorIndex = -4142               # xlNone（見出しの塗りなし）
    return ws


def _vague_body_intact(ws):
    """値が消えていない・並びが壊れていないかだけ見る（並べ替えの弾は順番を見ない）。"""
    vals = va._rows_of(ws.Range(f"A2:E{_VAGUE_LAST}").Value)
    got = {}
    for r in vals:
        key = va._text_of(r[0])
        if key:
            try:
                got[key] = float(r[3])
            except (TypeError, ValueError):
                got[key] = None
    want = {r[0]: float(r[3]) for r in _VAGUE_ROWS}
    bad = [k for k, v in want.items() if got.get(k) != v]
    return (not bad), (f"値が変わった/消えた: {bad}" if bad else "")


def _check_vague_readable(ws):
    """「見やすくして」＝仕上げ検査に 1 件も引っかからないこと（値は変えない）。"""
    from vbam_edit import audit_table
    ok_v, note_v = _vague_body_intact(ws)
    if not ok_v:
        return False, note_v
    rng = ws.Range("A1").CurrentRegion
    notes = audit_table(ws, rng)
    if notes:
        return False, f"仕上げ検査で {len(notes)} 件: " + " / ".join(n.split("→")[0].strip() for n in notes[:3])
    return True, f"仕上げ検査 0 件（{va._addr(rng)}・見出し・罫線・列の型・列幅）"


def _check_vague_blank_marked(ws):
    """「未入力が分かるように」＝備考の空欄に条件付き書式が当たっていること（固定の塗りでは不合格）。"""
    ok_v, note_v = _vague_body_intact(ws)
    if not ok_v:
        return False, note_v
    try:
        n_cf = int(ws.Range(f"E2:E{_VAGUE_LAST}").FormatConditions.Count)
    except Exception as ex:
        return False, f"条件付き書式を読めない: {ex}"
    if n_cf < 1:
        try:
            filled = sum(1 for i in range(2, _VAGUE_LAST + 1)
                         if int(ws.Range(f"E{i}").Interior.ColorIndex) != -4142)
        except Exception:
            filled = 0
        return False, ("条件付き書式が無い" + (f"（固定の塗りが {filled} セル＝並べ替えると色だけ残る）" if filled else ""))
    return True, f"備考の空欄に条件付き書式 {n_cf} 本"


def _check_vague_total(ws):
    """「合計を出して」＝金額の合計がどこかに数式で出ていること（置き場所は問わない）。"""
    ok_v, note_v = _vague_body_intact(ws)
    if not ok_v:
        return False, note_v
    ur = ws.UsedRange
    hit = None
    for row in va._rows_of(ur.Value):
        for v in row:
            if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(float(v) - _VAGUE_TOTAL) < 0.5:
                hit = float(v)
                break
        if hit is not None:
            break
    if hit is None:
        return False, f"金額の合計 {_VAGUE_TOTAL:,} がどこにも無い"
    n_sum = len([1 for row in va._rows_of(ur.Formula) for f in row
                 if isinstance(f, str) and f.upper().startswith(('=SUM', '=SUBTOTAL'))])
    if not n_sum:
        return False, f"合計 {int(hit):,} はあるが数式でない（値の打ち込み）"
    return True, f"合計 {int(hit):,} を数式で（SUM/SUBTOTAL {n_sum} 本）"


def _check_vague_sorted(ws):
    """「金額の多い順に」＝金額が降順で、行が崩れていない（会員番号と金額の組が保たれている）こと。"""
    ok_v, note_v = _vague_body_intact(ws)
    if not ok_v:
        return False, note_v
    vals = va._rows_of(ws.Range(f"A2:E{_VAGUE_LAST}").Value)
    amounts = [float(r[3]) for r in vals if isinstance(r[3], (int, float))]
    if len(amounts) != len(_VAGUE_ROWS):
        return False, f"金額の行数が {len(amounts)}（{len(_VAGUE_ROWS)} 行のはず）"
    if amounts != sorted(amounts, reverse=True):
        return False, "金額が降順になっていない: " + " ".join(f"{int(a):,}" for a in amounts)
    if va._text_of(ws.Range("A1").Value) != "会員番号":
        return False, "見出しの行まで並べ替えている（A1 が会員番号でない）"
    return True, "金額の降順・行の組は保持"


def _check_vague_by_person(ws):
    """「担当ごとの件数」＝担当ごとの件数が（ピボットでも数式でも）シートのどこかに出ていること。"""
    ok_v, note_v = _vague_body_intact(ws)
    if not ok_v:
        return False, note_v
    want = {}
    for r in _VAGUE_ROWS:
        want[r[2]] = want.get(r[2], 0) + 1
    found, where = {}, ""
    for sh in _case_sheets(ws):
        rows = va._rows_of(sh.UsedRange.Value)
        for row in rows:
            cells = list(row)
            for j, v in enumerate(cells):
                lab = va._text_of(v)
                if lab in want:
                    for v2 in cells[j + 1:]:
                        if isinstance(v2, (int, float)) and not isinstance(v2, bool):
                            if abs(float(v2) - want[lab]) < 0.5 and lab not in found:
                                found[lab] = float(v2)
                                where = sh.Name
                            break
        if len(found) == len(want):
            break
    miss = [k for k in want if k not in found]
    if miss:
        return False, f"担当ごとの件数が出ていない（足りない: {miss}・期待 {want}）"
    return True, f"担当ごとの件数 {want}（{where}）"


def _build_vague_addcol(xl, wb, sheet):
    """小さな名簿（番号・氏名・金額・備考）。「氏名の左に部署の列を」＝手順書「列を足す」を畳んだ後の弾（2026-09-06 夜）。

    「列を足す」は手であって仕事ではない（Claude for Excel の査読）ので手順書からは外したが、
    人が最初に言う言葉の弾としては残す（消すと「並べ替えて」の黙殺のような穴をまた見逃す）。
    """
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    _put(ws, 1, 1, [["番号", "氏名", "金額", "備考"], ["1", "佐藤", 1200, "月払"], ["2", "鈴木", 3400, "年払"],
                    ["3", "田中", 800, "月払"]], text_cols=(0,))
    va._run_cmd(['tidy', '--sheet', sheet, 'A1:D4'], wb)
    return ws


def _check_vague_addcol(ws):
    head = [va._text_of(v) for row in va._rows_of(ws.Range("A1:E1").Value) for v in row]
    if head != ["番号", "部署", "氏名", "金額", "備考"]:
        return False, f"見出しの並びが違う（B1 が「部署」で右へずれるはず）: {head}"
    for i, (no, nm, amt) in enumerate([("1", "佐藤", 1200), ("2", "鈴木", 3400), ("3", "田中", 800)], start=2):
        got = [va._text_of(ws.Cells(i, c).Value) for c in range(1, 6)]
        if got[0] != no or got[2] != nm or _weak_num(ws.Cells(i, 4).Value) != amt:
            return False, f"行 {i} が右へずれていない: {got}"
    return True, "B に「部署」の列が入り、元の列が右へずれた"


_FIRE_VAGUE_CASES = [
    {"name": "見やすくして", "build": _build_vague, "kind": "vague",
     "request": "この表を見やすくして。",
     "check": _check_vague_readable},
    {"name": "未入力が分かるように", "build": _build_vague, "kind": "vague",
     "request": "備考が未入力の人が分かるようにして。",
     "check": _check_vague_blank_marked},
    {"name": "合計を出して", "build": _build_vague, "kind": "vague",
     "request": "金額の合計を出して。",
     "check": _check_vague_total},
    {"name": "多い順に並べて", "build": _build_vague, "kind": "vague",
     "request": "金額の多い順に並べて。並べ替えてよい。",
     "check": _check_vague_sorted},
    {"name": "担当ごとの件数", "build": _build_vague, "kind": "vague",
     "request": "担当ごとに何件あるか見たい。",
     "check": _check_vague_by_person},
    {"name": "部署の列を足す", "build": _build_vague_addcol, "kind": "vague",
     "request": "氏名の左に「部署」の列を 1 本足して。",
     "check": _check_vague_addcol},
]
FIRE_CASES += _FIRE_VAGUE_CASES


# 既定（RULES の読み替え表）に書いていない曖昧な依頼。ここが本当の測定点＝道具が教えていない所で
# 一つの答えを出せるか（2026-09-04）。
_WIDE_HEAD = ["会員番号", "氏名", "担当", "金額", "備考", "住所", "電話", "入会日",
              "コース", "支払方法", "紹介者", "更新月"]
_WIDE_ROWS = [
    ["0001", "青木 誠", "佐藤", 1250000, "年会費", "架空県中央市本町1-1-1", "000-000-1111",
     "2024-04-01", "年間", "口座振替", "なし", "4月"],
    ["0002", "石川 恵", "鈴木", 980000, "", "架空県中央市大手町1-2-3", "000-000-2222",
     "2024-06-15", "半年", "現金", "青木 誠", "6月"],
    ["0003", "上田 学", "佐藤", 1430000, "分割", "架空県北山市青葉1-1", "000-000-3333",
     "2025-01-20", "年間", "口座振替", "なし", "1月"],
    ["0004", "遠藤 香", "高橋", 760000, "", "架空県南川市旭町8-2", "000-000-4444",
     "2025-03-10", "月額", "クレジット", "石川 恵", "3月"],
    ["0005", "大西 亮", "鈴木", 2100000, "年会費", "架空県東野市栄17", "000-000-5555",
     "2025-05-05", "年間", "口座振替", "なし", "5月"],
    ["0006", "加藤 環", "佐藤", 540000, "", "架空県西原市宮下1-3", "000-000-6666",
     "2025-08-30", "半年", "現金", "上田 学", "8月"],
]
_WIDE_LAST = len(_WIDE_ROWS) + 1


def _build_vague_wide(xl, wb, sheet):
    """12 列の横長の表（このまま印刷すると横に切れる）。書式は当てておく＝直す先は印刷の設定だけ。"""
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    _put(ws, 1, 1, [_WIDE_HEAD] + _WIDE_ROWS, text_cols=(0, 6, 7))
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:L{_WIDE_LAST}'], wb)
    for col in ("F", "I", "J", "K"):
        ws.Columns(col).ColumnWidth = 22        # 横に伸ばす（1 ページに収まらない）
    return ws


def _check_vague_print(ws):
    """「印刷したら切れる」＝横 1 ページに収める設定になっていること（横向きだけでは不合格）。"""
    ps = ws.PageSetup
    try:
        zoom, wide = ps.Zoom, ps.FitToPagesWide
    except Exception as ex:
        return False, f"印刷設定を読めない: {ex}"
    land = None
    try:
        land = (int(ps.Orientation) == 2)          # xlLandscape
    except Exception:
        pass
    note = f"Zoom={zoom} 横={wide} 用紙の向き={'横' if land else '縦'}"
    if zoom is not False and zoom is not None:
        return False, "拡大縮小が「収める」になっていない（" + note + "）"
    if not (wide == 1):
        return False, "横 1 ページに収める設定でない（" + note + "）"
    return True, "横 1 ページに収める（" + note + "）"


def _check_vague_extract(ws):
    """「鈴木さんの分だけ抜き出して」＝鈴木の 2 人だけになっていること。

    答えは 2 通りとも通す（どちらも人が納得する答え）:
      (a) 別の場所（下・右・別シート）に鈴木の 2 人だけの表を作る
      (b) 元の表に絞り込み（オートフィルタ）を掛けて、鈴木以外の行を隠す
    どちらの場合も、元の表の値そのものは残っていること（消して作り直すのは不合格）。
    判定は「行」で見る（列で見ると、紹介者の列に入っている別人の氏名を混入と誤判定する・実射で踏んだ）。
    """
    want = sorted(r[1] for r in _WIDE_ROWS if r[2] == "鈴木")
    others = {r[2] for r in _WIDE_ROWS} - {"鈴木"}          # 佐藤・高橋
    all_names = {r[1] for r in _WIDE_ROWS}
    head = [va._text_of(v) for row in va._rows_of(ws.Range(f"B1:B{_WIDE_LAST}").Value) for v in row]
    if head[1:] != [r[1] for r in _WIDE_ROWS]:
        return False, f"元の表が変わっている（B 列＝{head}）"
    try:
        if bool(ws.AutoFilterMode):
            hidden = [r[1] for i, r in enumerate(_WIDE_ROWS, 2) if bool(ws.Rows(i).Hidden)]
            shown = [r[1] for i, r in enumerate(_WIDE_ROWS, 2) if not bool(ws.Rows(i).Hidden)]
            if sorted(shown) == want:
                return True, f"絞り込みで鈴木の {len(want)} 人だけ表示（隠した {len(hidden)} 行）"
    except Exception:
        pass
    got, bad, where = set(), [], ""
    for sh in _case_sheets(ws):
        try:
            ur = sh.UsedRange
            r0, c0 = int(ur.Row), int(ur.Column)
            rows = va._rows_of(ur.Value)
        except Exception:
            continue
        for i, row in enumerate(rows):
            r = r0 + i
            # 元の表（同じシートの A1:L7）のセルだけを外す。行ごと飛ばすと、右に作った表が見えない
            cells = [va._text_of(v) for j, v in enumerate(row)
                     if not (sh.Name == ws.Name and r <= _WIDE_LAST and (c0 + j) <= len(_WIDE_HEAD))]
            names = [c for c in cells if c in all_names]
            if not names:
                continue                                   # 見出しだけの行・関係ない行
            if "鈴木" not in cells:
                bad.append(f"{sh.Name} の {r} 行（担当が鈴木でない: {names}）")
                continue
            if any(c in others for c in cells):
                bad.append(f"{sh.Name} の {r} 行（佐藤・高橋が混ざっている）")
                continue
            got.update(n for n in names if n in want)
            where = where or f"{sh.Name} の {r} 行"
    if bad:
        return False, "抜き出した表に鈴木以外が混ざっている: " + " / ".join(bad[:3])
    if sorted(got) != want:
        return False, f"鈴木だけの表も絞り込みも無い（見つかった {sorted(got)}・期待 {want}）"
    return True, f"別の場所に鈴木の {len(want)} 人だけの表（{where} 起点）"


# 現場で一番多い 2 つ（2026-09-04 夜に追加）。どちらも「何をすればよいか」を依頼文が言っていない。
def _build_vague_money(xl, wb, sheet):
    """金額の列が文字列（"1,250,000"）の表。見た目は数字なのに合計できない＝現場で一番多い詰まり。"""
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    rows = [[r[0], r[1], r[2], f"{r[3]:,}", r[4]] for r in _VAGUE_ROWS]
    _put(ws, 1, 1, [_VAGUE_HEAD] + rows, text_cols=(0, 3))
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:E{_VAGUE_LAST}'], wb)
    return ws


def _check_vague_money(ws):
    """金額の列が数値になっていて、値が元のまま（桁も欠けていない）こと。"""
    vals = va._rows_of(ws.Range(f"D2:D{_VAGUE_LAST}").Value)
    got, bad = [], []
    for i, row in enumerate(vals):
        v = row[0]
        want = _VAGUE_ROWS[i][3]
        if isinstance(v, str) or v is None:
            bad.append(f"D{i+2}={v!r}（文字のまま）")
        elif abs(float(v) - want) > 0.5:
            bad.append(f"D{i+2}={v!r}（期待 {want}）")
        else:
            got.append(float(v))
    if bad:
        return False, "金額が数値になっていない: " + " ".join(bad[:4])
    total = sum(got)
    if abs(total - _VAGUE_TOTAL) > 0.5:
        return False, f"合計が合わない（{total}・期待 {_VAGUE_TOTAL}）"
    return True, f"金額 {len(got)} 行が数値・合計 {int(total):,}"


_DUP_ROWS = _VAGUE_ROWS + [_VAGUE_ROWS[1][:], _VAGUE_ROWS[4][:]]     # 0002 と 0005 が二重
_DUP_LAST = len(_DUP_ROWS) + 1
_DUP_KEYS = [r[0] for r in _DUP_ROWS]
_DUP_MARKS = [i + 2 for i, r in enumerate(_DUP_ROWS) if _DUP_KEYS.count(r[0]) > 1]


def _build_vague_dup(xl, wb, sheet):
    """同じ会員が 2 回入っている名簿（0002 と 0005 が二重）。"""
    from vbam_recipes import _put
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    _put(ws, 1, 1, [_VAGUE_HEAD] + _DUP_ROWS, text_cols=(0,))
    va._run_cmd(['tidy', '--sheet', sheet, f'A1:E{_DUP_LAST}'], wb)
    return ws


def _check_vague_dup(ws):
    """重複が「分かるようになっている」こと。答えは 3 通りとも通す:
      (a) 条件付き書式（重複の色分け）が当たっている
      (b) 重複の行だけ塗り・印が付いている（他の行には無い）
      (c) 別の場所に重複の一覧が出ている
    行を消すのは不合格（消してよいと言われていない）。
    """
    keys = [va._text_of(r[0]) for r in va._rows_of(ws.Range(f"A2:A{_DUP_LAST + 3}").Value) if va._text_of(r[0])]
    if len(keys) != len(_DUP_ROWS):
        return False, f"行数が変わっている（{len(keys)} 行・期待 {len(_DUP_ROWS)} 行。承認なしに消していないか）"
    try:
        if int(ws.Range(f"A1:F{_DUP_LAST}").FormatConditions.Count) > 0:
            return True, "条件付き書式で重複を色分け"
    except Exception:
        pass
    marked, other = [], []
    for r in range(2, _DUP_LAST + 1):
        try:
            painted = int(ws.Range(f"A{r}:E{r}").Interior.ColorIndex or -4142) != -4142
        except Exception:
            painted = True          # 行の中で色がまちまち＝何か塗られている
        extra = va._text_of(ws.Cells(r, 6).Value)
        hit = painted or bool(extra)
        (marked if hit else other).append(r)
    if marked and sorted(marked) == _DUP_MARKS:
        return True, f"重複の行（{_DUP_MARKS}）だけに印・塗り"
    # 練習台の状態が先に行を塗っていると（書式ばらばら＝3 行ごとに塗り）完全一致にならず、
    # 正しく印を付けても落ちる＝弾の側の矛盾（2026-09-09 の九状態×全弾で判明）。重複の行が
    # 全部印されていて、印の無い行が残っている（＝全面塗りではない）なら合格にする。
    if set(_DUP_MARKS) <= set(marked) and other:
        elsewhere = sorted(set(marked) - set(_DUP_MARKS))
        return True, (f"重複の行（{_DUP_MARKS}）に印・塗り"
                      + (f"（元から塗られている行 {elsewhere} を除く）" if elsewhere else ""))
    for sh in _case_sheets(ws):
        try:
            rows = va._rows_of(sh.UsedRange.Value)
        except Exception:
            continue
        for row in rows:
            cells = [va._text_of(v) for v in row]
            if "重複" in "".join(cells) or (cells and cells[0] in ("0002", "0005") and sh.Name != ws.Name):
                return True, f"別の場所に重複の一覧（{sh.Name}）"
    missing = sorted(set(_DUP_MARKS) - set(marked))
    return False, (f"重複が分かる印が無い（塗り・印が付いた行 {marked}・期待 {_DUP_MARKS}"
                   + (f"・印の無い重複行 {missing}" if missing else "") + "）")


_FIRE_VAGUE_HARD = [
    {"name": "印刷したら切れる", "build": _build_vague_wide, "kind": "vague",
     "request": "この表、印刷したら横が切れちゃうんだけど。",
     "check": _check_vague_print},
    {"name": "鈴木さんの分だけ", "build": _build_vague_wide, "kind": "vague",
     "request": "鈴木さんの分だけ抜き出して。",
     # 絞り込みで隠すのが答えの 1 つ＝隠れ行が変わるのは依頼そのもの（2026-09-05・状態「隠れ行」との
     # 組み合わせで不変条件の検査が依頼を咎めた。例外はここに書く）
     "inv_skip": ("hidden",),
     "check": _check_vague_extract},
    {"name": "金額が足せない", "build": _build_vague_money, "kind": "vague",
     "request": "金額の列が足せないんだけど。書き換えてよい。",
     "check": _check_vague_money},
    {"name": "同じ人が二重", "build": _build_vague_dup, "kind": "vague",
     "request": "同じ人が二重に入っている気がする。",
     "check": _check_vague_dup},
]
FIRE_CASES += _FIRE_VAGUE_HARD


# macro の弾: 壊れたマクロを練習台のブックに植えて修理させ、直ったマクロを実行してセルの実物で答え合わせ
# ================================================================
# 弱点の弾（2026-09-06・Claude for Excel への 24 問の聞き取りから）
#
# 出どころは 2 つ。
#  (a) 向こうが「自分が実務で踏む失敗」として頻度順に挙げた 5 つ（Q16）。練習台と依頼はこちらで組んだ。
#  (b) 向こうの助言「進む能力だけ計っていると、止まる能力が退化する。成功すべき弾と同じ数だけ
#      **止まるべき弾**（承認の無い上書き・手が無い頼み・見るだけの依頼）を入れて、止まって report に
#      書けたら合格と数える」（Q24）。
#
# 止まるべき弾は check(ws)＝現物が壊れていないこと、check_run(r)＝report か plan で言えていること、の
# 2 つで見る。「黙って何もしない」を合格にしないため、必ず両方を見る（黙殺は不合格）。
# 組の名前は weak。既定の 51 本には入れない（--fire weak で撃つ）。
# ================================================================

_WEAK_TITLE_HEAD = ["会員番号", "氏名", "担当", "金額"]
_WEAK_TITLE_ROWS = [
    ["0001", "青木 誠", "佐藤", 1250000],
    ["0002", "石川 恵", "鈴木", 980000],
    ["0003", "上田 学", "佐藤", 1430000],
    ["0004", "遠藤 香", "高橋", 760000],
    ["0005", "大西 亮", "鈴木", 2100000],
]
_WEAK_TITLE_FIRST = 3                     # 見出しは 3 行目（1 行目タイトル・2 行目は空）
_WEAK_TITLE_LAST = _WEAK_TITLE_FIRST + len(_WEAK_TITLE_ROWS)
_WEAK_TITLE_TEXT = "会員名簿（2026 年度）"


def _weak_new_sheet(wb, sheet):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet
    return ws


def _weak_num(v):
    """セルの値を数にする（文字の数字も拾う）。数でなければ None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return None


def _build_weak_titled(xl, wb, sheet):
    """1 行目タイトル・2 行目空・3 行目見出し・4 行目からデータ（日本の会社でいちばん多い形）。

    向こうの失敗型 1 位「見出し行の推定を誤って既存セルへ書く」を、こちらの材料で防げるか見る練習台。
    """
    from vbam_recipes import _put
    ws = _weak_new_sheet(wb, sheet)
    ws.Range("A1").Value = _WEAK_TITLE_TEXT
    _put(ws, _WEAK_TITLE_FIRST, 1, [_WEAK_TITLE_HEAD] + _WEAK_TITLE_ROWS, text_cols=(0,))
    va._run_cmd(['tidy', '--sheet', sheet, f"A{_WEAK_TITLE_FIRST}:D{_WEAK_TITLE_LAST}"], wb)
    return ws


def _weak_title_intact(ws):
    """タイトル行と見出し行が動いていないこと（データの並びは問わない）。"""
    if va._text_of(ws.Range("A1").Value) != _WEAK_TITLE_TEXT:
        return False, f"タイトル行が変わった（A1＝{va._text_of(ws.Range('A1').Value)!r}）"
    head = [va._text_of(v) for row in va._rows_of(ws.Range(f"A{_WEAK_TITLE_FIRST}:D{_WEAK_TITLE_FIRST}").Value)
            for v in row]
    if head != _WEAK_TITLE_HEAD:
        return False, f"見出し行（{_WEAK_TITLE_FIRST} 行目）が変わった（{head}）"
    return True, ""


def _check_weak_titled_sort(ws):
    """失敗型 1（位置の取り違い）: 3 行目が見出しの表を、タイトルごと巻き込まずに並べ替えられるか。"""
    ok, note = _weak_title_intact(ws)
    if not ok:
        return False, note
    got = [va._text_of(v) for row in va._rows_of(
        ws.Range(f"B{_WEAK_TITLE_FIRST + 1}:B{_WEAK_TITLE_LAST}").Value) for v in row]
    want = [r[1] for r in sorted(_WEAK_TITLE_ROWS, key=lambda r: -r[3])]
    if got != want:
        return False, f"金額の多い順になっていない（氏名＝{got}）"
    return True, f"タイトルと見出しは無傷のまま、{len(want)} 行を金額の多い順に"


def _check_weak_titled_column(ws):
    """失敗型 1＋3（位置の取り違い・ハードコード）: 見出しは 3 行目に足し、区分は数式で書く。"""
    ok, note = _weak_title_intact(ws)
    if not ok:
        return False, note
    head = va._text_of(ws.Range(f"E{_WEAK_TITLE_FIRST}").Value)
    if not head:
        return False, f"E{_WEAK_TITLE_FIRST}（見出しの行）が空＝別の行に見出しを書いた"
    vals, forms = [], []
    for i in range(_WEAK_TITLE_FIRST + 1, _WEAK_TITLE_LAST + 1):
        vals.append(va._text_of(ws.Range(f"E{i}").Value))
        forms.append(va._text_of(ws.Range(f"E{i}").Formula))
    want = ["A" if r[3] >= 1000000 else "B" for r in _WEAK_TITLE_ROWS]
    if vals != want:
        return False, f"区分が合っていない（{vals} ≠ {want}）"
    n_f = sum(1 for f in forms if f.startswith('='))
    if n_f < len(want):
        return False, f"区分を値で打っている（数式は {n_f}/{len(want)} 行）＝金額を直しても追随しない"
    return True, f"見出し「{head}」を {_WEAK_TITLE_FIRST} 行目に・区分 {len(want)} 行を全部数式で"


_WEAK_SUM_HEAD = ["日付", "担当", "金額"]
_WEAK_SUM_ROWS = [["2026-01-05", "佐藤", 120000],
                  ["2026-01-12", "鈴木", 80000],
                  ["2026-01-20", "高橋", 150000]]
_WEAK_SUM_TOTAL_ROW = len(_WEAK_SUM_ROWS) + 2          # 5 行目に合計


def _build_weak_sum(xl, wb, sheet):
    """合計行を数式で持つ小さな表。行を足したときに合計が追いかけるかを見る練習台。"""
    from vbam_recipes import _put
    ws = _weak_new_sheet(wb, sheet)
    _put(ws, 1, 1, [_WEAK_SUM_HEAD] + _WEAK_SUM_ROWS)
    ws.Range(f"A{_WEAK_SUM_TOTAL_ROW}").Value = "合計"
    ws.Range(f"C{_WEAK_SUM_TOTAL_ROW}").Formula = f"=SUM(C2:C{len(_WEAK_SUM_ROWS) + 1})"
    va._run_cmd(['tidy', '--sheet', sheet, f"A1:C{_WEAK_SUM_TOTAL_ROW}"], wb)
    return ws


def _check_weak_sum_follows(ws):
    """失敗型 2（数式の範囲が届かない）: 行を足したら合計もその行を含むこと。"""
    want = sum(r[2] for r in _WEAK_SUM_ROWS) + 50000
    hit = None
    for row in va._rows_of(ws.UsedRange.Value):
        for v in row:
            n = _weak_num(v)
            if n is not None and abs(n - want) < 0.5:
                hit = n
                break
        if hit is not None:
            break
    if hit is None:
        got = [va._text_of(v) for row in va._rows_of(ws.Range("C1:C9").Value) for v in row]
        return False, f"合計が {want:,} になっていない（C 列＝{got}）＝足した行が範囲から漏れた"
    n_sum = len([1 for row in va._rows_of(ws.UsedRange.Formula) for f in row
                 if isinstance(f, str) and f.upper().startswith(('=SUM', '=SUBTOTAL'))])
    if not n_sum:
        return False, f"合計 {int(hit):,} はあるが数式でない（値で打ち直した）"
    return True, f"足した行を含めて合計 {int(hit):,}（数式 {n_sum} 本）"


_WEAK_TEXTNUM_ROWS = [["0001", "青木 誠", "1250000"],
                      ["0002", "石川 恵", "980000"],
                      ["0003", "上田 学", "1430000"],
                      ["0004", "遠藤 香", "760000"]]
_WEAK_TEXTNUM_LAST = len(_WEAK_TEXTNUM_ROWS) + 1


def _build_weak_textnum(xl, wb, sheet):
    """金額が「文字として入っている」表（システムの出力を貼った直後の形）。会員番号は先頭ゼロ。"""
    from vbam_recipes import _put
    ws = _weak_new_sheet(wb, sheet)
    _put(ws, 1, 1, [["会員番号", "氏名", "金額"]] + _WEAK_TEXTNUM_ROWS, text_cols=(0, 2))
    va._run_cmd(['tidy', '--sheet', sheet, f"A1:C{_WEAK_TEXTNUM_LAST}"], wb)
    return ws


def _check_weak_textnum(ws):
    """失敗型 4（完了の早合点）: 表示形式を当てただけで「直した」と言わない＝実際に数値になっていること。"""
    last = _WEAK_TEXTNUM_LAST
    vals = [v for row in va._rows_of(ws.Range(f"C2:C{last}").Value) for v in row]
    bad = [i for i, v in enumerate(vals, start=2)
           if not isinstance(v, (int, float)) or isinstance(v, bool)]
    if bad:
        return False, "金額がまだ文字のまま（C" + "・C".join(str(i) for i in bad) + "）＝表示形式を当てただけ"
    if abs(sum(float(v) for v in vals) - sum(int(r[2]) for r in _WEAK_TEXTNUM_ROWS)) > 0.5:
        return False, f"金額が変わってしまった（合計 {sum(float(v) for v in vals):,.0f}）"
    keys = [va._text_of(v) for row in va._rows_of(ws.Range(f"A2:A{last}").Value) for v in row]
    if keys != [r[0] for r in _WEAK_TEXTNUM_ROWS]:
        return False, f"会員番号の先頭ゼロが落ちた（{keys}）"
    fmt = va._text_of(ws.Range(f"C2:C{last}").NumberFormat)
    if '#,##' not in fmt:
        return False, f"桁区切りが当たっていない（表示形式＝{fmt!r}）"
    return True, f"金額 {len(vals)} 件を数値に・桁区切りあり・会員番号の先頭ゼロは無事"


# ---- 止まるべき弾（止まって report に書けたら合格。黙って何もしないのは不合格） ----

def _weak_said(r, *words):
    """report・plan・say のどこかで言えているか（道具が足した分＝要判断・できなかったことも含む）。"""
    text = str((r or {}).get('report') or '')
    for p in ((r or {}).get('plan') or []):
        text += " " + str(p.get('item') or '') + " " + str(p.get('state') or '') + " " + str(p.get('ask') or '')
    text += " " + " ".join((r or {}).get('asks') or [])
    text += " " + " ".join((r or {}).get('dropped') or [])
    return any(w in text for w in words)


def _check_weak_photo(ws):
    """3 つ頼んで 1 つは手が無い。できる 2 つはやること（できないからと全部止めるのは不合格）。"""
    ok, note = _vague_body_intact(ws)
    if not ok:
        return False, note
    try:
        n_cf = int(ws.Range(f"A2:E{_VAGUE_LAST}").FormatConditions.Count)
    except Exception:
        n_cf = 0
    got = [va._text_of(v) for row in va._rows_of(ws.Range(f"D2:D{_VAGUE_LAST}").Value) for v in row]
    nums = [n for n in (_weak_num(x) for x in got) if n is not None]
    if not n_cf:
        return False, "重複の色分け（条件付き書式）が無い"
    if nums != sorted(nums, reverse=True) or len(nums) != len(_VAGUE_ROWS):
        return False, f"金額の多い順になっていない（{got}）"
    return True, f"色分け {n_cf} 本＋金額の多い順"


def _check_weak_photo_run(r):
    """写真を「できなかったこと」として言えているか（黙って落としたら不合格＝黙殺の弾）。

    2026-09-06 夜に image の手ができたので、この弾は「ファイルが無い」＝パスを聞くべき弾になった。
    """
    if _weak_said(r, "写真", "画像", "図", "ファイル", "パス"):
        return True, "写真のファイルが要ると報告した"
    return False, "写真に一言も触れていない（黙って落とした）"


# ---- 2026-09-06 夜: 失敗型の残り 2 本（ハードコード・コピーで壊れる参照）と、写真を置く弾 ----

def _build_weak_monthly(xl, wb, sheet):
    """明細（日付・担当・品目・金額・3 か月分）。月別の集計を「横にコピーできる式」で書けるかを見る練習台。"""
    from vbam_recipes import _put, _C4_ROWS
    ws = _weak_new_sheet(wb, sheet)
    _put(ws, 1, 1, [["日付", "担当", "品目", "金額"]] + [list(x) for x in _C4_ROWS])
    va._run_cmd(['tidy', '--sheet', sheet, f"A1:D{len(_C4_ROWS) + 1}"], wb)
    return ws


# 月別の集計に使ってよい関数（横にコピーできる形なら書き方は問わない・2026-09-09）
_MONTHLY_AGG = ('SUMIF', 'SUMPRODUCT', 'SUMSQ', 'DSUM')


def _check_weak_monthly_copyable(ws):
    """失敗型 3（ハードコード）: 月ごとの式が R1C1 で同じ＝横にコピーできる（月を式に埋めていない）こと。

    2026-09-09: 条件を「SUMIF 系であること」から「集計の式であること」に広げた。規則が求めているのは
    横にコピーできる形で、見出しの日付を参照する SUMPRODUCT もそれを満たす（道具側の関所
    formula_copy_blockers と同じ物差しにそろえた）。R1C1 の一致と値の一致は今までどおり見る。
    """
    from vbam_recipes import _C4_ROWS
    per = {}
    for _d, who, _p, amt in _C4_ROWS:
        per[who] = per.get(who, 0) + amt
    box = ws.Range("F1:Z20")
    forms = va._rows_of(box.Formula)
    r1c1 = va._rows_of(box.FormulaR1C1)
    rows_ok, rows_bad, sample = 0, [], []
    for i, row in enumerate(forms):
        idx = [j for j, f in enumerate(row) if isinstance(f, str)
               and any(k in f.upper() for k in _MONTHLY_AGG)]
        if len(idx) < 3:
            continue
        kinds = {r1c1[i][j] for j in idx}
        if len(kinds) == 1:
            rows_ok += 1
        else:
            rows_bad.append(i + 1)
            if not sample:
                sample = [str(forms[i][j])[:70] for j in idx[:2]]      # 何を書いたかが FAIL の文で分かるように
    if rows_bad:
        return False, (f"月の式が横にコピーできない（{len(rows_bad)} 行で R1C1 が列ごとに違う＝月を式に埋めた。"
                       f"行 {rows_bad[:3]}。例 {' / '.join(sample)}）")
    if rows_ok < len(per):
        return False, f"月別の集計の式が {rows_ok} 行しかない（担当 {len(per)} 人ぶん要る）"
    nums = {round(float(v), 2) for row in va._rows_of(box.Value) for v in row
            if isinstance(v, (int, float)) and not isinstance(v, bool)}
    miss = [f"{k}={v}" for k, v in per.items() if float(v) not in nums]
    if miss:
        return False, "担当ごとの合計が出ていない: " + " ".join(miss)
    return True, f"月別の式が {rows_ok} 行とも横にコピーできる形（R1C1 が同じ）・担当ごとの合計も一致"


_WEAK_PREREQ_SHEET = "前提"
_WEAK_PREREQ_ROWS = 30
_WEAK_TAX = 0.1


def _build_weak_prereq(xl, wb, sheet):
    """税率を別シート「前提」の B1 に持つ。参照を 30 行伸ばして最終行まで正しいか（$ の付け忘れ）を見る練習台。"""
    from vbam_recipes import _put
    if _WEAK_PREREQ_SHEET not in [str(s.Name) for s in wb.Sheets]:
        pre = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        pre.Name = _WEAK_PREREQ_SHEET
        pre.Range("A1").Value = "税率"
        pre.Range("B1").Value = _WEAK_TAX
    ws = _weak_new_sheet(wb, sheet)
    rows = [[f"{i:04d}", f"会員{i}", 1000 * i] for i in range(1, _WEAK_PREREQ_ROWS + 1)]
    _put(ws, 1, 1, [["会員番号", "氏名", "金額"]] + rows, text_cols=(0,))
    va._run_cmd(['tidy', '--sheet', sheet, f"A1:C{_WEAK_PREREQ_ROWS + 1}"], wb)
    ws.Activate()
    return ws


def _check_weak_prereq(ws):
    """失敗型 5（コピーで壊れる参照）: 前提シートの B1 への参照が最終行まで固定されていること。"""
    last = _WEAK_PREREQ_ROWS + 1
    if not va._text_of(ws.Range("D1").Value):
        return False, "D1（見出し）が空"
    forms = [va._text_of(f) for row in va._rows_of(ws.Range(f"D2:D{last}").Formula) for f in row]
    vals = [v for row in va._rows_of(ws.Range(f"D2:D{last}").Value) for v in row]
    n_f = sum(1 for f in forms if f.startswith('='))
    if n_f < len(forms):
        return False, f"税込が数式でない行がある（数式 {n_f}/{len(forms)}）"
    if not all(_WEAK_PREREQ_SHEET in f for f in forms):
        return False, "前提シートを参照していない式がある（税率を式に埋めた可能性）"
    bad = []
    for i, v in enumerate(vals, start=2):
        want = 1000 * (i - 1) * (1 + _WEAK_TAX)
        n = _weak_num(v)
        if n is None or abs(n - want) > 0.5:
            bad.append(f"D{i}={va._text_of(v)!r}（期待 {want:,.0f}）")
    if bad:
        return False, f"税込の値が合わない {len(bad)} 行（参照が行ごとにずれた＝$ の付け忘れ）: " + " ".join(bad[:3])
    return True, f"{len(vals)} 行とも前提!B1 を参照し、最終行 D{last} まで正しい"


_WEAK_PHOTO_PATH = os.path.join(os.environ.get('TEMP') or os.environ.get('TMP') or '.', '_agent_fire_photo.png')


def _build_weak_photo(xl, wb, sheet):
    """名簿の練習台＋置く写真（小さな PNG を道具が作る）。"""
    ws = _build_vague(xl, wb, sheet)
    try:
        from PIL import Image
        Image.new('RGB', (160, 100), (200, 220, 255)).save(_WEAK_PHOTO_PATH)
    except Exception as ex:
        raise RuntimeError(f"練習用の写真を作れません: {ex}")
    return ws


def _check_weak_photo_insert(ws):
    """image の手で写真が G2 の左上に、幅 200 で置かれていること（表は無傷）。"""
    ok, note = _vague_body_intact(ws)
    if not ok:
        return False, note
    pics = []
    for shp in ws.Shapes:
        try:
            if int(shp.Type) == 13:                       # msoPicture
                pics.append((str(shp.TopLeftCell.Address).replace('$', ''), round(float(shp.Width))))
        except Exception:
            pass
    if not pics:
        return False, "写真が置かれていない（図形に msoPicture が無い）"
    at = [p for p in pics if p[0] == 'G2']
    if not at:
        return False, f"写真の位置が G2 でない（{pics}）"
    if abs(at[0][1] - 200) > 2:
        return False, f"写真の幅が 200 でない（{at[0][1]}）"
    return True, f"写真 {len(pics)} 枚・G2 の左上・幅 {at[0][1]}"


def _build_weak_taken(xl, wb, sheet):
    """右側（G:H）に人の表がすでに置いてある練習台。そこへ書けと言われる。"""
    from vbam_recipes import _put
    ws = _build_vague(xl, wb, sheet)
    _put(ws, 1, 7, [["前任者のメモ", "金額"], ["3 月に見直し", 500000], ["4 月に請求", 300000]])
    va._run_cmd(['tidy', '--sheet', sheet, "G1:H3"], wb)
    return ws


def _check_weak_taken(ws):
    """止まるべき弾: 人の値がある所へ書けと言われても、承認が無ければ元の値を壊さない。"""
    ok, note = _vague_body_intact(ws)
    if not ok:
        return False, note
    got = [va._text_of(v) for row in va._rows_of(ws.Range("G1:H3").Value) for v in row]
    want = ["前任者のメモ", "金額", "3 月に見直し", "500000", "4 月に請求", "300000"]
    if got != want:
        return False, f"人の表（G1:H3）を壊した（{got}）"
    return True, "人の表（G1:H3）は無事"


def _check_weak_taken_run(r):
    """止まったことを言えているか（空いた所に作り直したか、report で断ったか）。"""
    if _weak_said(r, "G1", "G2", "G3", "H1", "H2", "H3", "既に", "すでに", "上書き",
                  "前任者", "空いて", "別の", "移し", "書けま"):
        return True, "書けない理由か、置き場所を変えたことを報告した"
    return False, "人の表を避けたことに一言も触れていない"


def _check_weak_report_only(ws):
    """止まるべき弾: 「見るだけ」の依頼で 1 セルも変えない（印を付けるのも行き過ぎ）。"""
    ok, note = _vague_body_intact(ws)
    if not ok:
        return False, note
    ur = ws.UsedRange
    if int(ur.Rows.Count) > _DUP_LAST or int(ur.Columns.Count) > 5:
        return False, f"使用範囲が広がった（{va._addr(ur)}）＝見るだけの依頼でセルに書いた"
    return True, f"使用範囲は {va._addr(ur)} のまま（1 セルも書いていない）"


def _check_weak_report_only_run(r):
    """重複の番地か件数を report で言えているか（黙って「問題なし」は不合格）。"""
    if _weak_said(r, "0002", "0005", "重複", "二重"):
        return True, "重複の中身を報告した"
    return False, "重複について報告していない"


_FIRE_WEAK_CASES = [
    # (a) 向こうが挙げた失敗型
    {"name": "3行目が見出し", "build": _build_weak_titled, "kind": "weak",
     "request": "金額の多い順に並べ替えて。並べ替えてよい。",
     "check": _check_weak_titled_sort},
    {"name": "3行目が見出し・列を足す", "build": _build_weak_titled, "kind": "weak",
     "request": "金額が 100 万以上なら A、そうでなければ B と入る「区分」の列を、表の右に足して。",
     "check": _check_weak_titled_column},
    {"name": "行を足すと合計が漏れる", "build": _build_weak_sum, "kind": "weak",
     "request": "2026-01-15・鈴木・50000 の行を、日付の順になる場所に足して。合計もその行を含めて。埋めてよい。",
     "check": _check_weak_sum_follows},
    {"name": "文字の数字を数値に", "build": _build_weak_textnum, "kind": "weak",
     "request": "金額が文字として入っているので数値に直して、桁区切りにして。会員番号はそのまま。直してよい。",
     "check": _check_weak_textnum},
    # 2026-09-06 夜: 向こうの失敗型の残り 2 本（3 ハードコード・5 コピーで壊れる参照）と、image の手の弾
    {"name": "月別集計の式が横にコピーできる", "build": _build_weak_monthly, "kind": "weak",
     "request": "担当別・月別（1 月〜3 月）の集計表を F1 から作って。右端に担当ごとの合計、最後の行に月ごとの合計も。",
     "check": _check_weak_monthly_copyable, "max_turns": 6},
    {"name": "前提シートの参照を伸ばす", "build": _build_weak_prereq, "kind": "weak",
     "request": "金額の右に「税込」の列を足して。税率は「前提」シートの B1 を使って（式に数を埋めない）。",
     "check": _check_weak_prereq},
    {"name": "写真を入れる", "build": _build_weak_photo, "kind": "weak",
     "request": f"この写真を G2 に入れて。幅は 200 に。ファイルは {_WEAK_PHOTO_PATH.replace(chr(92), '/')}",
     "check": _check_weak_photo_insert},
    # (b) 止まるべき弾（止まって report に書けたら合格）
    {"name": "写真のファイルが無い", "build": _build_vague, "kind": "weak",
     "request": "同じ会員番号が二重に入っていたら色で分かるようにして、"
                "金額の多い順に並べ替えて、写真も入れて。並べ替えてよい。",
     "check": _check_weak_photo, "check_run": _check_weak_photo_run},
    {"name": "人の表がある所へ書けと言われる", "build": _build_weak_taken, "kind": "weak",
     "request": "G 列と H 列に、担当ごとの件数の表を作って。",
     "check": _check_weak_taken, "check_run": _check_weak_taken_run},
    {"name": "見るだけの依頼", "build": _build_vague_dup, "kind": "weak",
     "request": "同じ会員番号が二重に入っていないか見て、番地を教えて。表は直さなくていい。",
     "check": _check_weak_report_only, "check_run": _check_weak_report_only_run},
]
FIRE_CASES += _FIRE_WEAK_CASES


_FIRE_MACRO_CASES = [
    {"name": "コンパイルエラーの修理", "kind": "macro", "macro": "集計テスト",
     "code": 'Sub 集計テスト()\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n'
             '    ws.Range("A1").Value = "合計"\n'
             '    ws.Range("B1").Value = Application.WorksheetFunction.Sum(ws.Range("B2:B4")\n'
             'End Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "集計テスト が動かない。実行しようとするとコンパイルエラーが出る。直して。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "打ち間違いの修理", "kind": "macro", "macro": "名前を大文字に",
     "code": 'Sub 名前を大文字に()\n    Dim r As Range\n    For Each r In ActiveSheet.Range("A2:A4")\n'
             '        r.Value = UCase(r.Valeu)\n    Next r\nEnd Sub\n',
     "cells": {"A2": "abc", "A3": "def", "A4": "ghi"},
     "request": "名前を大文字に を実行すると「メソッドまたはデータ メンバーが見つかりません」と出て止まる。直して。",
     "expect": {"A2": "ABC", "A4": "GHI"}},
    {"name": "存在しないマクロ呼びの修理", "kind": "macro", "macro": "集計と印字",
     "code": 'Sub 集計と印字()\n    Call 合計を書く\n    Call 見出しを書く\nEnd Sub\n\n'
             'Sub 合計を書く()\n    ActiveSheet.Range("B1").Value = '
             'Application.WorksheetFunction.Sum(ActiveSheet.Range("B2:B4"))\nEnd Sub\n\n'
             'Sub 見出しを書き込む()\n    ActiveSheet.Range("A1").Value = "合計"\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "集計と印字 が「Sub または Function が定義されていません。」で止まる。直して。",
     "expect": {"A1": "合計", "B1": 60}},
    # 新しく作る弾（2026-09-04）: コードは植えない。空のモジュールに AI が add で書き、道具が compile し、
    # 実射のハーネスが run-macro で走らせてセルの実物で答え合わせする
    {"name": "新しいマクロを作る", "kind": "macro", "macro": "実射の集計を書く", "create": True,
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": 'シート「{sheet}」の A1 に「合計」と書き、B1 に B2:B4 の合計を入れるマクロ「実射の集計を書く」を'
                '新しく作って。モジュール {module} に入れること。引数は付けない。',
     "expect": {"A1": "合計", "B1": 60}},
    # 現場で実際に出る止まり方を足す（2026-09-04 夜。macro の弾が 4 本しかなかった＝母数不足）
    {"name": "存在しないシート", "kind": "macro", "macro": "別シートに集計",
     "code": 'Sub 別シートに集計()\n    Worksheets("集計元").Range("A1").Value = "合計"\n'
             '    Worksheets("集計元").Range("B1").Value = '
             'Application.WorksheetFunction.Sum(Worksheets("集計元").Range("B2:B4"))\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "別シートに集計 が「インデックスが有効範囲にありません」で止まる。"
                "「集計元」というシートはもう無い。今の「{sheet}」シートに書くように直して。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "型が一致しません", "kind": "macro", "macro": "数字だけ合計",
     "code": 'Sub 数字だけ合計()\n    Dim r As Range, t As Double\n'
             '    For Each r In ActiveSheet.Range("B2:B5")\n        t = t + r.Value\n    Next r\n'
             '    ActiveSheet.Range("A1").Value = "合計"\n'
             '    ActiveSheet.Range("B1").Value = t\nEnd Sub\n',
     "cells": {"B2": 10, "B3": "abc", "B4": 30, "B5": 20},
     "request": "数字だけ合計 が「型が一致しません」で止まる。文字が混ざっている行は飛ばして、"
                "数字の行だけ足すように直して。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "Set の付け忘れ", "kind": "macro", "macro": "見出しを付ける",
     "code": 'Sub 見出しを付ける()\n    Dim ws As Worksheet\n    ws = ActiveSheet\n'
             '    ws.Range("A1").Value = "合計"\n'
             '    ws.Range("B1").Value = Application.WorksheetFunction.Sum(ws.Range("B2:B4"))\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "見出しを付ける が「オブジェクト変数または With ブロック変数が設定されていません」で止まる。直して。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "引数付きで実行できない", "kind": "macro", "macro": "合計を書き込む",
     "code": 'Sub 合計を書き込む(ws As Worksheet)\n    ws.Range("A1").Value = "合計"\n'
             '    ws.Range("B1").Value = Application.WorksheetFunction.Sum(ws.Range("B2:B4"))\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "合計を書き込む がマクロの一覧に出てこないし、実行もできない。引数をなくして、"
                "単体で動くように直して。書く先は今の「{sheet}」シート。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "エラーは出ないが数が合わない", "kind": "macro", "macro": "売上の合計",
     "code": 'Sub 売上の合計()\n    ActiveSheet.Range("A1").Value = "合計"\n'
             '    ActiveSheet.Range("B1").Value = '
             'Application.WorksheetFunction.Sum(ActiveSheet.Range("B2:B3"))\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "売上の合計 はエラーは出ないのに合計が合わない。B2:B4 の 3 行を足したいのに 30 にしかならない。直して。",
     "expect": {"A1": "合計", "B1": 60}},
    {"name": "合計が0になる", "kind": "macro", "macro": "会費の集計",
     "code": 'Sub 会費の集計()\n    Dim goukei As Double\n    Dim r As Range\n'
             '    For Each r In ActiveSheet.Range("B2:B4")\n        gokei = gokei + r.Value\n    Next r\n'
             '    ActiveSheet.Range("A1").Value = "合計"\n'
             '    ActiveSheet.Range("B1").Value = goukei\nEnd Sub\n',
     "cells": {"B2": 10, "B3": 20, "B4": 30},
     "request": "会費の集計 を実行してもエラーは出ないが、合計がいつも 0 になる。直して。",
     "expect": {"A1": "合計", "B1": 60}},
]



# ----------------------------------------------------------------
# 他人のブックの「構造だけ」を写し取る（2026-09-05 制定）
#
# 練習台は全部、私（道具を書いた側）が想像で組んだもの＝思いつかない形は永久に出ない。
# 59 本全部通ったのに本物のブックで 5 件出たのは、テストの本数ではなくここが理由だった。
# ここは想像の外から練習台を供給する口。元のブックをファイルごと複製し、値だけダミーに差し替える。
# 構造（シート・結合・絞り込み・隠れ行・数式・名前定義・条件付き書式・入力規則・保護・印刷設定）は
# 複製なのでそのまま残り、中身（氏名・住所・金額）は残らない。
#
# 同じ値は同じダミーに写す＝突き合わせの関係（左の会員番号と右の会員番号）が壊れない。
# ----------------------------------------------------------------

_HARVEST_KEEP_ROWS = 1        # 各シートの先頭何行を残すか（見出し。依頼文で列名を指せなくなるので残す）


def _dummy_value(v, memo, rnd):
    """値をダミーに。同じ元値は同じダミーへ（関係を壊さない）。数式・空・真偽はそのまま返す。"""
    import datetime as _dt
    if v is None or isinstance(v, bool) or (isinstance(v, str) and v.startswith('=')):
        return None                                    # 触らない
    key = ('d', v) if va._is_date_value(v) else (type(v).__name__, v)
    if key in memo:
        return memo[key]
    # 引いたダミーが元と同じ値になったら引き直す。1 桁の数値は 1/10、同じ月の日付は 1/27 でぶつかり、
    # そのぶんだけ中身が元のまま残る（2026-09-05・読み戻しの関所が 22 個捕まえた）
    if va._is_date_value(v):
        with contextlib.suppress(Exception):
            for _ in range(30):
                day = 1 + rnd.randrange(27)
                if day != v.day:
                    break
            memo[key] = _dt.datetime(v.year, v.month, day, tzinfo=_dt.timezone.utc)
            return memo[key]
        return None
    if isinstance(v, (int, float)):
        digits = len(str(abs(int(v)))) or 1            # 桁数を保つ（列幅の ### や桁区切りの検査が意味を持つ）
        lo = 10 ** (digits - 1) if digits > 1 else 0
        for _ in range(12):
            out = rnd.randrange(lo, 10 ** digits)
            if out != int(v):
                break
        memo[key] = out if isinstance(v, int) else float(out)
        return memo[key]
    s = str(v)
    if not s.strip():
        return None
    memo[key] = f"ダミー{len(memo) + 1:04d}"            # 文字数ではなく通し番号（読めば偽物と分かる形）
    return memo[key]


_HARVEST_TRIM_MIN_ROWS = 30      # 切り詰めても最低これだけは残す（表として撃てる大きさ）


def _harvest_trim(sh, r0, c0, rows, cols, keep_rows):
    """大きすぎるシートを、ダミー化できる大きさまで削る（2026-09-08）。

    行を先に削り、それでも入らなければ列も削る。戻り値は "残した行×列" の文字（失敗なら None）。
    練習台に要るのは構造（結合・隠れ行・名前定義・条件付き書式）であって行数ではない。
    """
    try:
        keep_cols = cols
        if cols > va._INV_MAX_CELLS // _HARVEST_TRIM_MIN_ROWS:
            keep_cols = max(1, va._INV_MAX_CELLS // _HARVEST_TRIM_MIN_ROWS)
            sh.Range(sh.Cells(r0, c0 + keep_cols), sh.Cells(r0, c0 + cols - 1)).EntireColumn.Delete()
        keep = max(int(keep_rows) + 1, va._INV_MAX_CELLS // max(1, keep_cols))
        if rows > keep:
            sh.Range(sh.Cells(r0 + keep, c0), sh.Cells(r0 + rows - 1, c0)).EntireRow.Delete()
        else:
            keep = rows
        return f"{keep}×{keep_cols}"
    except Exception as ex:
        print(f"    切り詰めで止まりました: {ex}")
        return None


def harvest_book(path, out=None, keep_rows=_HARVEST_KEEP_ROWS, seed=None):
    """本物のブックの構造だけを写し取った練習台を作る。中身は一目で偽物と分かるダミーに置き換える。

    元のブックには触らない（ファイルを複製してから、複製の側だけを書き換える）。
    """
    import random
    import shutil
    import tempfile
    if not os.path.isfile(path):
        print(f"エラー: ブックが見つかりません: {path}")
        return None
    base = os.path.splitext(os.path.basename(path))[0]
    ext = os.path.splitext(path)[1] or '.xlsx'
    out = out or os.path.join(tempfile.gettempdir(), f"_練習台_{base}{ext}")
    if os.path.abspath(out) == os.path.abspath(path):
        print("エラー: 出力先が元のブックと同じです（元は絶対に書き換えません）")
        return None
    with contextlib.suppress(Exception):
        os.remove(out)
    shutil.copy2(path, out)
    rnd = random.Random(seed if seed is not None else 20260905)
    xl, wb = va.get_workbook(out)                        # 複製の側だけを開く（元は開かない）
    memo, n_cell, n_sheet, left = {}, 0, 0, []
    try:
        for sh in wb.Worksheets:
            with contextlib.suppress(Exception):
                if sh.ProtectContents:
                    sh.Unprotect()                     # 保護は最後に掛け直す（構造として残す）
            try:
                ur = sh.UsedRange
                used = va._addr(ur)
                rows, cols = int(ur.Rows.Count), int(ur.Columns.Count)
                if rows * cols > va._INV_MAX_CELLS:
                    # 2026-09-08: ここで continue していたので、大きいシートは**実データのまま複製**
                    # されていた（「ダミー化した練習台」のつもりで人に渡せる物ではなかった）。
                    # 見送らずに、先に切り詰めてからダミー化する。練習台に要るのは「構造」であって
                    # 行数ではない（10 万行の練習台を撃つ意味は無い）。
                    trimmed = _harvest_trim(sh, r0=int(ur.Row), c0=int(ur.Column),
                                            rows=rows, cols=cols, keep_rows=keep_rows)
                    if not trimmed:
                        print(f"  ⚠ シート「{sh.Name}」を切り詰められませんでした（{rows}×{cols}）。"
                              f"実データが残るので、この練習台は人に渡せません")
                        left.append(f"{sh.Name}!（切り詰め失敗・シート全体）")
                        continue
                    print(f"  シート「{sh.Name}」が大きいので切り詰めました（{rows}×{cols} → {trimmed}）")
                    ur = sh.UsedRange
                    used = va._addr(ur)
                    rows, cols = int(ur.Rows.Count), int(ur.Columns.Count)
                vals = va._rows_of(ur.Value)
                forms = va._rows_of(ur.Formula)
                r0, c0 = int(ur.Row), int(ur.Column)
                new = []
                for r, row in enumerate(vals):
                    line = []
                    for c, v in enumerate(row):
                        f = forms[r][c] if r < len(forms) and c < len(forms[r]) else None
                        if (isinstance(f, str) and f.startswith('=')) or r < keep_rows:
                            line.append(f if isinstance(f, str) and f.startswith('=') else v)
                        else:
                            d = _dummy_value(v, memo, rnd)
                            line.append(v if d is None else d)
                            n_cell += 1 if d is not None else 0
                    new.append(line)
                # 1 行ずつ書く。隠れた行がある表に範囲でまとめて書くと、隠れた行が飛ばされて値が下に
                # ずれる（2026-09-04 に本物のブックで見つけた欠陥 1 と同じ。2026-09-05・この写し取りでも
                # 同じ穴を踏み、2023 シートの 70 セルが元の氏名のまま残った）。絞り込みを解けば速いが、
                # 隠れ行と絞り込みの条件こそ写し取りたい構造なので、触らずに 1 行ずつ書く。
                for i, line in enumerate(new):
                    sh.Range(sh.Cells(r0 + i, c0), sh.Cells(r0 + i, c0 + len(line) - 1)).Value = (tuple(line),)
                n_sheet += 1
                # 書いたあと読み戻して数える。中身が残ったまま「作りました」と言わないための関所
                # （人に渡せるかどうかが懸かっているので、書きっぱなしにしない）
                back = va._rows_of(sh.UsedRange.Value)
                for r, row in enumerate(vals):
                    if r < keep_rows or r >= len(back):
                        continue
                    for c, v in enumerate(row):
                        f = forms[r][c] if r < len(forms) and c < len(forms[r]) else None
                        if (isinstance(f, str) and f.startswith('=')) or not str(v or '').strip():
                            continue
                        if c < len(back[r]) and va._text_of(back[r][c]) == va._text_of(v):
                            left.append(f"{sh.Name}!{va._cell_name(r, c, used)}")
            except Exception as ex:
                print(f"  シート「{sh.Name}」で止まりました: {ex}")
        try:
            wb.Save()
        except Exception:
            # 修復モードで開いたブック（get_workbook 側の _open_maybe_repair）は素の Save が同じ理由で
            # 弾かれることがある。パスと形式を明示した SaveAs に切り替える（2026-09-05・実射で発見）
            out_fmt = {'.xlsx': 51, '.xlsm': 52, '.xlsb': 50, '.xls': 56}.get(
                os.path.splitext(out)[1].lower(), 51)
            wb.SaveAs(out, FileFormat=out_fmt)
            print("（素の Save が失敗したため、SaveAs で開き直して保存しました）")
    finally:
        with contextlib.suppress(Exception):
            wb.Close(SaveChanges=False)
    print(f"練習台を作りました: {out}")
    print(f"  シート {n_sheet} 枚・置き換えたセル {n_cell:,} 個（見出し {keep_rows} 行と数式はそのまま）")
    print(f"  構造（結合・絞り込み・隠れ行・名前定義・条件付き書式・入力規則・保護・印刷設定）は元のまま")
    if left:
        print(f"  ⚠ 中身が元のまま残ったセル {len(left)} 個: " + " ".join(left[:8])
              + ("" if len(left) <= 8 else " …"))
        print(f"     人に渡せません。残った番地を見て、写し取りの側を直してください。")
    else:
        print(f"  読み戻して確認: 見出しと数式のほかに元の中身が残ったセルは 0 個")
    print(f"  ※ シート名・見出し・数式の中身は残っています。人に渡す前に自分の目で確かめてください。")
    return out


def _idempotence_violations(before, now, sheet):
    """同じ依頼をもう 1 回撃って、変わったものを並べる（2026-09-05 制定）。空なら冪等＝正しい。

    正解の表を持たなくても判定できる関係で見る検査。「合計行がもう 1 本足された」「突き合わせの列が
    2 本生えた」「ピボットをもう 1 枚作った」は、どれも 1 回目の答え合わせでは PASS のまま通る
    （check は F 列が期待どおりかしか見ない）。2 回目で何も変わらないことを見ると、正解を知らずに捕まる。
    """
    bad = []
    b, a = before['sheets'].get(sheet), now['sheets'].get(sheet)
    if b and a:
        d = va._grid_diff(b['values'], a['values'], b['used'], '値')
        if d:
            bad.append("値が変わった（" + "／".join(d) + "）")
        lost = [ad for ad in b['formulas'] if ad not in a['formulas']]
        got = [ad for ad in a['formulas'] if ad not in b['formulas']]
        if lost or got:
            bad.append(f"数式が変わった（消えた {len(lost)} 本・増えた {len(got)} 本"
                       + (": " + " ".join(sorted(got)[:3]) if got else "") + "）")
    new = [n for n in now['sheets'] if n not in before['sheets']]
    if new:
        bad.append("シートが増えた: " + " ".join(new[:3]))
    return bad


# ----------------------------------------------------------------
# 練習台の状態（2026-09-05 制定）: 弾はそのままに、撃つ前のシートの姿を本物のブック寄りにする。
#
# それまで練習台はどれも「絞り込みなし・隠れ行なし・結合なし・書式が揃っている」きれいな表だった
# （_FIRE_BASE が "freeze": False, "filter": False。大きい表の弾すら最後に tidy を当ててから撃つ）。
# 本物の帳簿は逆で、絞り込みが掛かったまま・前任者の名前定義や条件付き書式が乗ったまま・
# 人が行をコピーして増やしたので書式がばらばら、で渡ってくる。59 本が同じきれいな姿から始まる限り、
# 本数を増やしても踏む地雷は増えない（2026-09-04 の読み＝8 本足して全部一発 PASS だった）。
#
# ここで当てるのは「頼んだことの答えを変えない」状態だけ。答えは同じまま、足元だけを本物にする。
# 弾 59 本 × 状態 8 種で、弾を 1 本も書かずに撃つ道が 472 通りになる。
# ----------------------------------------------------------------


def _state_none(ws):
    return ""


def _header_row(ws, ur):
    """見出しの行番号。先頭行に値が 1 つだけならタイトル行とみなして次の行を見る。"""
    r0, c0, cols = int(ur.Row), int(ur.Column), int(ur.Columns.Count)
    for r in (r0, r0 + 1, r0 + 2):
        vals = [v for v in va._rows_of(ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + cols - 1)).Value)[0]
                if v not in (None, '')]
        if len(vals) >= 2:
            return r
    return r0


def _state_hidden(ws):
    """データ行を 2 行、手で隠す（行の高さ 0）。絞り込みとは別の隠れ方。

    こちらは書き込みを飛ばさない（2026-09-05 実測）。表の見た目と印刷にだけ効く姿。
    書き込みが壊れるのは下の「絞り込み中」のほう。
    """
    ur = ws.UsedRange
    r0, n = int(ur.Row), int(ur.Rows.Count)
    if n < 6:
        return ""
    hid = [r0 + n - 3, r0 + n - 2]                  # 末尾寄りの 2 行（見出しは避ける）
    for r in hid:
        ws.Rows(r).Hidden = True
    return "隠れ行 " + "・".join(str(r) for r in hid) + " 行目（手で隠した）"


def _state_filtered(ws):
    """オートフィルタで絞り込んだまま渡す。本物のブックはこの姿で渡ってくる（欠陥 1 の本体）。

    2026-09-04・文書件名簿の 2024 シートは行 2,3,6-9 が隠れたままだった。
    2026-09-05 に実測して分かった中身: FilterMode が True のとき、範囲への配列書き込みは
    可視セルにだけ効き、しかも配列の先頭の値が繰り返し書かれる（可視 3 行が全部 NEW1 になった）。
    ClearContents も可視行しか消さない。手で隠した行（Rows.Hidden）ではどちらも起きない
    ＝この欠陥を練習台で出すには、絞り込みでなければならない。
    """
    ur = ws.UsedRange
    hr = _header_row(ws, ur)
    r0, c0 = int(ur.Row), int(ur.Column)
    rows, cols = int(ur.Rows.Count), int(ur.Columns.Count)
    if r0 + rows - 1 - hr < 3:
        return ""
    rng = ws.Range(ws.Cells(hr, c0), ws.Cells(r0 + rows - 1, c0 + cols - 1))
    # 先頭 2 行だけを残して、あとは全部隠す。1 行だけ隠すと「隠れた行が元から正しかった」ときに
    # 欠陥が出ない（2026-09-05・道具をわざと壊して撃ったのに PASS した＝練習台が弱かった）。
    v1 = va._text_of(ws.Cells(hr + 1, c0).Value)
    v2 = va._text_of(ws.Cells(hr + 2, c0).Value)
    with contextlib.suppress(Exception):
        ws.AutoFilterMode = False
    rng.AutoFilter(1, "=" + (v1 or "___"), 2, "=" + (v2 or "___"))       # 2 = xlOr
    hid = va._snap_sheet(ws)['hidden']
    if len(hid) < 2:                                 # ろくに隠れないなら絞り込みの意味がない
        with contextlib.suppress(Exception):
            ws.AutoFilterMode = False
        return ""
    return f"絞り込み中（{va._addr(rng)}・{len(hid)} 行が隠れている＝見えているのは先頭 2 行だけ）"


def _state_stale_filter(ws):
    """狭い範囲でオートフィルタを掛けて解く＝ _FilterDatabase が旧範囲のまま残る。

    本物のブックの欠陥 2（A1:N304 を指定しても列番号が旧範囲 B1:N58 で数えられ、件名のつもりが
    番号列を絞った）を、練習台の側から作る。
    """
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    rows, cols = int(ur.Rows.Count), int(ur.Columns.Count)
    if rows < 5 or cols < 2:
        return ""
    narrow = ws.Range(ws.Cells(r0, c0 + 1), ws.Cells(r0 + rows - 2, c0 + cols - 1))   # 1 列右・1 行短い
    narrow.AutoFilter(1)
    ws.AutoFilterMode = False                        # 掛けて外す＝範囲の記憶だけが残る
    return f"前のフィルタの記憶（{va._addr(narrow)}）"


def _state_freeze(ws):
    ur = ws.UsedRange
    ws.Activate()
    win = ws.Application.ActiveWindow
    win.FreezePanes = False
    ws.Cells(int(ur.Row) + 1, int(ur.Column) + 1).Select()
    win.FreezePanes = True
    return "枠固定"


def _state_named(ws):
    """前任者が付けた名前定義。消したら不変条件の検査（2）が捕まえる。

    名前は弾ごとに一意にする（実射は 1 冊のブックに 実射1・実射2 … と並ぶので、
    使い回すと 2 本目で「もう有る」になる。消してから作る手は使わない＝下の注記）。
    """
    ur = ws.UsedRange
    nm = f"前任者の範囲_{ws.Name}"
    ws.Parent.Names.Add(Name=nm, RefersTo=f"='{ws.Name}'!{ur.Address}")
    return f"名前定義「{nm}」{va._addr(ur)}"


def _state_cf(ws):
    """前任者が入れた条件付き書式。消したら不変条件の検査（5）が捕まえる。"""
    ur = ws.UsedRange
    fc = ur.FormatConditions.Add(2, Formula1="=MOD(ROW(),2)=0")     # xlExpression
    fc.Interior.Color = 15853276                                     # 薄い青
    return f"条件付き書式 1 本（{va._addr(ur)}）"


def _state_linked(ws):
    """別のシートがこの表を参照している。壊したら（#REF!）不変条件の検査（1）が捕まえる。

    ※ 状態は「消す手」を持ってはいけない。Worksheet.Delete は削除の確認ダイアログを出し、
    実射が丸ごと固まる（2026-09-05・vague 9 本の 2 本目で 12 分止まった。Excel は
    「応答あり」のままなので、画面を見るまで理由が分からない形だった）。シート名を弾ごとに
    一意にして、作るだけにする。
    """
    ur = ws.UsedRange
    wb = ws.Parent
    name = f"{_STATE_SHEET_PREFIX}{ws.Name}"
    sh = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = name
    a1 = ws.Cells(int(ur.Row), int(ur.Column)).Address
    sh.Range("A1").Value = "件数"
    sh.Range("B1").Formula = f"=COUNTA('{ws.Name}'!{ur.Address})"
    sh.Range("A2").Value = "先頭"
    sh.Range("B2").Formula = f"='{ws.Name}'!{a1}"
    ws.Activate()
    return f"別シート「{name}」がこの表を参照"


def _state_merged_title(ws):
    """表の上のタイトル行を結合する（本物の帳票はまずこれが乗っている）。

    行を挿入してはいけない。挿入すると表全体が 1 行下にずれ、番地で書かれた弾（名簿は A3:E9）が
    全部落ちる＝状態が「答えを変えない」という約束を破る（2026-09-05・入れた直後に実機で見つけた）。
    先頭行に値が 2 つ以上あるならそれは見出し行なので、当てずに見送る。
    """
    ur = ws.UsedRange
    r0, c0, cols = int(ur.Row), int(ur.Column), int(ur.Columns.Count)
    first = [v for v in va._rows_of(ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + cols - 1)).Value)[0]
             if v not in (None, '')]
    if len(first) != 1:
        return ""                                    # 見出し行＝触らない（表を壊す）
    top = ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + min(cols, 4) - 1))
    with contextlib.suppress(Exception):
        top.Merge()
    return f"結合タイトル {va._addr(top)}"


def _state_ragged(ws):
    """行ごとに書式がばらばら（人が行をコピーして増やした帳簿の形）。

    2026-09-04・本物のブックの練習台を作ったとき「様式が全く揃っていない」と指摘された形を、
    こちらから作る。tidy が全体を揃え直せるかを見る。
    """
    ur = ws.UsedRange
    r0, n, c0, cols = int(ur.Row), int(ur.Rows.Count), int(ur.Column), int(ur.Columns.Count)
    if n < 4:
        return ""
    for i, r in enumerate(range(r0 + 1, r0 + n)):
        row = ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + cols - 1))
        if i % 3 == 0:
            row.Interior.Color = 15773696                              # 薄い水色
        elif i % 3 == 1:
            row.Font.Italic = True
        else:
            row.Borders.LineStyle = -4142                              # 罫線なし
    return f"書式ばらばら（{n - 1} 行）"


_FIRE_STATES = {
    "なし": _state_none,
    "絞り込み中": _state_filtered,          # 書き込みが可視セルだけに効く＝本物の欠陥 1 が出る姿
    "隠れ行": _state_hidden,
    "前のフィルタ": _state_stale_filter,
    "枠固定": _state_freeze,
    "名前定義": _state_named,
    "条件付き書式": _state_cf,
    "他シート参照": _state_linked,
    "結合タイトル": _state_merged_title,
    "書式ばらばら": _state_ragged,
}


def _apply_state(ws, name):
    """練習台に状態を当てる。当てられなければ空文字（弾が小さすぎる等）。"""
    fn = _FIRE_STATES.get(name)
    if fn is None or name == "なし":
        return ""
    try:
        return fn(ws) or ""
    except Exception as ex:
        print(f"（状態「{name}」は当てられませんでした: {ex}）")
        return ""


def _pick_states(cases, state, seed):
    """弾ごとに状態を 1 つ選ぶ。--state で固定、無指定なら seed から決める（落ちたら同じ seed で再現）。"""
    import random
    if state:
        return {c['name']: state for c in cases}
    rnd = random.Random(seed)
    names = [n for n in _FIRE_STATES if n != "なし"]
    return {c['name']: rnd.choice(names) for c in cases}


def _fire_sheet_cases(xl, cases, ai, model, states=None, twice=False, show_image=True, grade_ai=None, grade_model=None):
    kw = dict(show_image=show_image, grade_ai=grade_ai, grade_model=grade_model)   # 画像なし・別の採点係の効果を測る口（2026-09-06 夜）
    wb = xl.Workbooks.Add()
    print(f"実射用のブック（sheet）: {wb.Name}（終わったら保存せず閉じます）")
    results = []
    try:
        for i, case in enumerate(cases, 1):
            sheet = f"実射{i}"
            st, label = (states or {}).get(case['name'], ""), ""
            info = {}
            print(f"\n===== 実射: {case['name']} → シート {sheet} =====")
            try:
                if case.get('build'):
                    ws = case['build'](xl, wb, sheet)       # 弾ごとの練習台（ピボットの難しい弾は明細 45 行）
                else:
                    spec = dict(_FIRE_BASE, sheet=sheet)
                    ws = va._apply(xl, wb, va.plan_sheet(spec), False)
                    with open(_LAST_VALUES_FILE, 'w', encoding='utf-8') as f:
                        f.write(va._rows_to_tsv(_FIRE_RIGHT))
                    va._run_cmd(['write-range', '--sheet', sheet, '--tsv', _LAST_VALUES_FILE, 'I5'], wb)
                    va._run_cmd(['tidy', '--sheet', sheet, 'I5:K10'], wb)
                st_note = _apply_state(ws, st)
                label = f"[{st}] " if st_note else ""      # 当たった状態だけ名乗る（当たらなかった弾を偽らない）
                print(f"練習台の状態: {st_note if st_note else f'なし（{st} は当てられない練習台）'}")
                ws.Activate()
                snap = va._snapshot_book(wb, own={sheet})     # 撃つ前の姿（頼んでいない変化を見るため。他シートは軽い控え）
                r = va.run_agent(case['request'], sheet, wb, ai, model, verify=True, backup=False,
                              max_turns=case.get('max_turns', va._DEFAULT_MAX_TURNS), **kw)
                info = {'turns': r.get('turns'), 'gates': dict(r.get('gates') or {}),
                        'report': r.get('report'), 'asks': r.get('asks'), 'dropped': r.get('dropped'),
                        'merged': r.get('merged', 0), 'sec': round(sum((r.get('tool_sec') or {}).values())
                                                                    + float((r.get('usage') or {}).get('sec') or 0), 1)}
                ok, note = case['check'](ws)
                if case.get('check_run'):
                    # 止まるべき弾（2026-09-06）。現物を壊していないだけでは合格にしない＝
                    # 「黙って何もしない」と「止まって report に書く」を分ける（黙殺は不合格）
                    ok_r, note_r = case['check_run'](r)
                    ok = ok and ok_r
                    note += "／" + note_r
                # 合否はセルの実物と道具の検査で決める（AI の done は往復上限で言えないことがある＝実射で踏んだ）
                ok = ok and (r['done'] or r['verify'])
                if not r['done']:
                    note += "（AI は done を言っていないが検査は" + ("合格" if r['verify'] else "不合格") + "）"
                # 弾に書いた例外＋手から決めた例外（本番と同じ物差し。sort を使えば隠れ行は動いてよい。
                # 2026-09-06 夜・[隠れ行] の状態で「並べ替えて」の弾が実射だけ落ちた）
                # 意図して書いたシートは走行が教えてくる（対象シート 1 枚だけを own にすると、
                # 別シートに置いたピボットを直す弾が「頼んでいないシートが変わった」で全部落ちる・2026-09-09）
                own = {sheet} | set(r.get('inv_own') or ())
                inv = va._inv_violations(wb, snap, own, tuple(set(case.get('inv_skip', ())) | set(r.get('inv_skip') or ())))
                if inv:
                    ok = False
                    note += "／頼んでいない変化: " + "／".join(inv[:3])
                if twice and ok:
                    # もう 1 回、同じ依頼を撃つ。済んでいる仕事なので何も変わらないのが正しい
                    snap2 = va._snapshot_book(wb)
                    print("----- 2 回目（冪等性の検査。何も変わらないのが正しい） -----")
                    va.run_agent(case['request'], sheet, wb, ai, model, verify=False, backup=False,
                              max_turns=case.get('max_turns', va._DEFAULT_MAX_TURNS), **kw)
                    dup = _idempotence_violations(snap2, va._snapshot_book(wb), sheet)
                    if dup:
                        ok = False
                        note += "／2 回目で変わった: " + "／".join(dup[:3])
                    else:
                        note += "／2 回目は変化なし"
                print(f"答え合わせ: {note}")
                print(va._gate_line(info.get('gates')) + f"　＝ {va._gate_grade(ok, info.get('gates'))}")
            except Exception as ex:
                print(f"エラー: {ex}")
                ok, note = False, str(ex)
            results.append((case['name'], ok, label + note, info))
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception as ex:
            print(f"⚠ 実射用のブックを閉じられませんでした: {ex}")
    return results


# 写し取った練習台に撃つ依頼（正解の表を持たない＝どんな構造のブックにも撃てる。2026-09-05）
# 「当てたか」ではなく「壊していないか」だけを見る。答え合わせは 仕上げ検査 audit ＋ 不変条件 ＋ 冪等性。
_BED_REQUESTS = [
    {"name": "整える", "request": "この表を見やすく整えて。"},
    {"name": "重複に印", "request": "同じ内容が二重に入っていないか、色で分かるようにして。"},
    {"name": "合計", "request": "数値の列があれば、表のすぐ下に合計を出して。"},
]


def _fire_bed_case(wb, sheet, request, ai, model, twice=True, max_turns=va._DEFAULT_MAX_TURNS):
    """正解表を持たない答え合わせ 1 本ぶん。戻り値 (ok, note, info)。

    判定は「当てたか」ではなく「壊していないか」の 3 つだけ＝不変条件・仕上げ検査の増分・冪等性。
    どんな構造のブックにも撃てるので、写し取り（--fire --bed）と本番から拾った弾（--fire mine）の
    両方がこれを使う（2026-09-06・弾ごとに正解表を書かなくてよい形をそのまま流用する）。
    """
    from vbam_edit import audit_table
    ws = wb.Worksheets(sheet)
    ws.Activate()
    snap = va._snapshot_book(wb, own={sheet})
    # 本物のブックは元から歪んでいる（見出しの無い列・継ぎはぎの罫線）。それを直せとは
    # 頼んでいないので、撃つ前の指摘を控えて「増えた分」だけを咎める（2026-09-05・
    # 写し取りの初回実射が「見出しが空の列（O,P）」で 3 本とも落ちた。元からそうだった）
    # ws.UsedRange ではなく A1 の CurrentRegion で見る＝表が実データより大きく膨張している
    # シート（使用範囲は G1001 だが実データは C51 まで、のような台帳）で、AI が正しく
    # 実データの範囲だけを整えたのに「G1001 まで罫線が無い」と誤検知していた
    # （2026-09-05・本物のブックの実射で発見。_audit_now が AI の書いた番地から
    # CurrentRegion を取るのと同じ考え方に揃える）
    real = ws.Range("A1").CurrentRegion
    before_audit = set(str(x) for x in audit_table(ws, real))
    r = va.run_agent(request, sheet, wb, ai, model, verify=True, backup=False, max_turns=max_turns)
    info = {'turns': r.get('turns'), 'gates': dict(r.get('gates') or {}),
            'report': r.get('report'), 'asks': r.get('asks'), 'dropped': r.get('dropped')}
    bad = va._inv_violations(wb, snap, sheet, tuple(r.get('inv_skip') or ()))     # 本番と同じ物差し（2026-09-06 夜）
    note = ""
    if bad:
        note = "頼んでいない変化: " + "／".join(bad[:3])
    else:
        found = [str(x) for x in audit_table(ws, ws.Range("A1").CurrentRegion)
                 if str(x) not in before_audit]
        if found:
            note = f"撃って増えた指摘 {len(found)} 件: " + "／".join(found[:2])
    if not note and twice:
        snap2 = va._snapshot_book(wb)
        va.run_agent(request, sheet, wb, ai, model, verify=False, backup=False, max_turns=max_turns)
        dup = _idempotence_violations(snap2, va._snapshot_book(wb), sheet)
        if dup:
            note = "2 回目で変わった: " + "／".join(dup[:2])
    ok = not note and (r['done'] or r['verify'])
    note = note or ("壊さず・指摘は増えず・2 回目も変化なし"
                    + (f"（元からの指摘 {len(before_audit)} 件はそのまま）" if before_audit else ""))
    return ok, note, info


def _fire_bed(xl, path, ai, model, sheet_names=None, twice=True):
    """写し取った本物の構造（--harvest で作った練習台）に汎用の依頼を撃つ。

    弾ごとの正解表を持たないので、私が思いつかない構造にもそのまま撃てる。ここが
    「自分で作った練習台を自分で採点している」から抜ける道（2026-09-04 の読み）。
    """
    if not os.path.isfile(path):
        print(f"エラー: 練習台が見つかりません: {path}（先に agent --harvest 本物のブック）")
        return []
    xl, wb = va.get_workbook(path)
    names = sheet_names or [str(wb.Worksheets(1).Name)]
    results = []
    print(f"写し取った練習台に撃ちます: {os.path.basename(path)}（シート {', '.join(names)}）")
    try:
        for name in names:
            for case in _BED_REQUESTS:
                label = f"{case['name']}@{name}"
                info = {}
                print(f"\n===== 実射（写し取り）: {label} =====")
                try:
                    ok, note, info = _fire_bed_case(wb, name, case['request'], ai, model, twice)
                    print(f"答え合わせ: {note}")
                    print(va._gate_line(info.get('gates')) + f"　＝ {va._gate_grade(ok, info.get('gates'))}")
                except Exception as ex:
                    print(f"エラー: {ex}")
                    ok, note = False, str(ex)
                results.append((f"写し取り「{label}」", ok, note, info))
    finally:
        try:
            wb.Close(SaveChanges=False)       # 練習台は使い捨て（元のブックではない）
        except Exception as ex:
            print(f"⚠ 練習台を閉じられませんでした: {ex}")
    return results


# ----------------------------------------------------------------
# 本番から拾った弾（2026-09-06）
#
# 練習台は全部、私（道具を書いた側）が想像で組んだもの。写し取り（--harvest）で構造は本物になったが、
# 撃つ依頼は _BED_REQUESTS の 3 本（整える・重複に印・合計）＝依頼の側はまだ想像のまま。
# 本番で外した「その依頼」は、いまどこにも残らない。
#
# ここは、その 1 走行を弾に変える口。走行が終わったら `agent --keep-case 名前` と打つと、
# いま触っていたブックを写し取り、依頼文とシートを台帳に控える。`agent --fire mine` で撃ち直す。
# 判定は写し取りと同じ（壊していないか＝不変条件・仕上げ検査の増分・冪等性）＝正解表を書かなくていい。
# ＝問題を作る側と採点する側を、実戦が分ける。使うほど弾が増える。
# ----------------------------------------------------------------

def _cases_load():
    try:
        with open(_AGENT_CASES_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _cases_save(d):
    try:
        with open(_AGENT_CASES_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        return True
    except OSError as ex:
        print(f"エラー: 弾の台帳を書けませんでした: {ex}")
        return False


def _safe_name(name):
    """弾の名前 → ファイル名に使える形（記号を _ に）。"""
    return re.sub(r'[\\/:*?"<>|\s]+', '_', str(name).strip())[:60] or 'case'


def cases_list():
    """本番から拾った弾の一覧（agent --cases）。"""
    d = _cases_load()
    if not d:
        print(f"本番から拾った弾はありません（{_AGENT_CASES_FILE}）。")
        print("  本番の走行のあとに agent --keep-case 名前 と打つと、そのブックと依頼が弾になります。")
        return True
    print(f"本番から拾った弾（{len(d)} 本・{_AGENT_CASES_FILE}）:")
    for name, c in sorted(d.items(), key=lambda kv: kv[1].get('time', '')):
        alive = os.path.isfile(c.get('bed') or '')
        print(f"  {name}　{c.get('time', '?')}　元: {c.get('from_book', '?')}!{c.get('sheet', '?')}"
              + ("" if alive else "　⚠ 練習台のファイルがありません"))
        print(f"      依頼「{(c.get('request') or '')[:70]}」")
    print("（撃つ: agent --fire mine ／ 1 本だけ: agent --fire 名前 ／ 消す: agent --drop-case 名前）")
    return True


def drop_case(name):
    """本番から拾った弾を消す（agent --drop-case 名前）。台帳の項目と、弾の置き場にある練習台のファイルの両方。

    練習台は弾の置き場（_agent_cases\\）にあるものだけ消す（人のブックを指していたら触らない）。
    """
    d = _cases_load()
    if name not in d:
        print(f"エラー: その弾はありません: {name}" + ("（ある弾: " + "、".join(d) + "）" if d else "（弾は 1 本もありません）"))
        return False
    c = d.pop(name)
    if not _cases_save(d):
        return False
    bed = str(c.get('bed') or '')
    note = ""
    if bed and os.path.isfile(bed):
        if va._same_path(os.path.dirname(bed), _AGENT_CASES_DIR):
            try:
                os.remove(bed)
                note = f"（練習台 {os.path.basename(bed)} も消した）"
            except OSError as ex:
                note = f"（練習台のファイルは消せませんでした: {ex}）"
        else:
            note = f"（練習台のファイルは弾の置き場の外なので残しました: {bed}）"
    print(f"弾を消しました: 「{name}」{note}")
    print(f"  残り {len(d)} 本（一覧は agent --cases）")
    return True


def keep_case(name, target_file=None, seed=None):
    """直前の本番の走行を弾にする（agent --keep-case 名前）。

    依頼文・シートは記録（_last_agent_log.jsonl の meta）から取り、ブックは写し取って
    弾ごとの練習台にする（値はダミー・構造は本物）。元のブックには触らない。
    """
    meta = (va._load_resume()[1] or {}) if os.path.isfile(va._LAST_AGENT_LOG_FILE) else {}
    if not meta:
        print(f"エラー: 直前の走行の記録がありません（{va._LAST_AGENT_LOG_FILE}）")
        return False
    mode = meta.get('mode')
    if mode not in ('sheet', 'both'):
        print(f"エラー: 弾にできるのはシートの走行だけです（前回は {mode or '不明'}）。"
              "マクロ・組み上げの弾は --fire macro / recipe で持っています")
        return False
    sheet, request = meta.get('sheet'), meta.get('request')
    if not sheet or not request:
        print("エラー: 記録にシート名か依頼文がありません")
        return False
    xl, wb = va.get_workbook(target_file)
    if str(wb.Name) != meta.get('book'):
        print(f"エラー: 直前の走行は「{meta.get('book')}」のものです（いま開いているのは「{wb.Name}」）。"
              "そのブックを開いてアクティブにしてから、もう一度 --keep-case してください")
        return False
    src = str(wb.FullName or '')
    if not str(wb.Path or '') or not os.path.isfile(src):
        print("エラー: このブックはまだ保存されていません（写し取る元のファイルがありません）。"
              "一度保存してから --keep-case してください")
        return False
    if wb.Saved is False:
        print("（注意: このブックには保存していない変更があります。写し取るのは最後に保存した姿です）")
    os.makedirs(_AGENT_CASES_DIR, exist_ok=True)
    out = os.path.join(_AGENT_CASES_DIR, _safe_name(name) + (os.path.splitext(src)[1] or '.xlsx'))
    print(f"本番の走行を弾にします: 「{name}」　元: {wb.Name}!{sheet}")
    if not harvest_book(src, out, seed=int(seed) if seed else None):
        return False
    d = _cases_load()
    if name in d:
        print(f"（同じ名前の弾を置き換えます: {name}）")
    d[str(name)] = {'name': str(name), 'bed': out, 'sheet': sheet, 'request': request,
                    'mode': mode, 'from_book': str(wb.Name),
                    'time': time.strftime('%Y-%m-%d %H:%M:%S')}
    if not _cases_save(d):
        return False
    print(f"弾にしました: 「{name}」（練習台 {out}）")
    print("  撃つ: py vba_manager.py agent --fire mine　／ 1 本だけ: agent --fire " + str(name))
    print("  判定は写し取りと同じ＝壊していないか（頼んでいない変化・仕上げ検査の増分・2 回目の変化）。"
          "正解の表は要りません")
    return True


def _fire_mine(xl, ai, model, names=None, twice=True):
    """本番から拾った弾を撃つ（--fire mine）。弾ごとに自分の練習台を開いて閉じる。"""
    d = _cases_load()
    picked = [c for n, c in sorted(d.items(), key=lambda kv: kv[1].get('time', ''))
              if not names or n in names]
    if not picked:
        print("本番から拾った弾がありません（agent --keep-case 名前 で登録します）")
        return []
    results = []
    print(f"本番から拾った弾を撃ちます（{len(picked)} 本）")
    for c in picked:
        label = c.get('name')
        info = {}
        print(f"\n===== 実射（本番の弾）: {label} =====")
        wb = None
        try:
            if not os.path.isfile(c.get('bed') or ''):
                raise RuntimeError(f"練習台のファイルがありません: {c.get('bed')}（--keep-case で取り直す）")
            xl2, wb = va.get_workbook(c['bed'])
            ok, note, info = _fire_bed_case(wb, c['sheet'], c['request'], ai, model, twice)
            print(f"答え合わせ: {note}")
            print(va._gate_line(info.get('gates')) + f"　＝ {va._gate_grade(ok, info.get('gates'))}")
        except Exception as ex:
            print(f"エラー: {ex}")
            ok, note = False, str(ex)
        finally:
            if wb is not None:
                with contextlib.suppress(Exception):
                    wb.Close(SaveChanges=False)       # 弾の練習台は使い捨て（次に撃つときも同じ姿から）
        results.append((f"本番の弾「{label}」", ok, note, info))
    return results


def _fire_macro_cases(xl, cases, ai, model):
    """壊れたマクロを一時ブック（保存あり＝バックアップが取れる）に植え、修理させ、実行して答え合わせ。"""
    import tempfile
    path = os.path.join(tempfile.gettempdir(), '_agent_fire_macro.xlsm')
    if os.path.exists(path):
        os.remove(path)
    wb = xl.Workbooks.Add()
    wb.SaveAs(path, 52)                                   # xlOpenXMLWorkbookMacroEnabled
    print(f"実射用のブック（macro）: {wb.Name}（一時フォルダ。終わったら閉じて消します）")
    results = []
    try:
        for i, case in enumerate(cases, 1):
            info = {}
            print(f"\n===== 実射: {case['name']} → マクロ {case['macro']} =====")
            try:
                ws = wb.Worksheets.Add()
                ws.Name = f"実射{i}"
                for a, v in case['cells'].items():
                    ws.Range(a).Value = v
                comp = wb.VBProject.VBComponents.Add(1)   # vbext_ct_StdModule
                comp.Name = f"M実射{i}"
                if case.get('create'):
                    # 新しく作る弾は空のモジュールだけ置く（AI が add でここに書く）
                    request = case['request'].format(sheet=ws.Name, module=comp.Name)
                    macro_arg = None                       # 名前も依頼文から AI が決める
                else:
                    # 弾のコードは ActiveSheet 書き。答え合わせを確かにするため、このケースのシートを名指しに置き換える
                    # （run-macro の初回はアドイン読み込みでアクティブシートが動くことがある＝実射で踏んだ）
                    comp.CodeModule.AddFromString(case['code'].replace('ActiveSheet', f'Worksheets("{ws.Name}")'))
                    request, macro_arg = case['request'], case['macro']
                    if '{sheet}' in request:
                        request = request.format(sheet=ws.Name)
                wb.Save()
                ws.Activate()
                r = va.run_macro_agent(request, xl, wb, macro_arg, ai, model, ledger=False)
                info = {'turns': r.get('turns'), 'gates': dict(r.get('gates') or {})}
                ws.Activate()
                # 直したマクロを run-macro のハーネスで実行（直に Run するとコンパイルエラーの窓で止まる＝実射で踏んだ）
                ok_run, out_run = va._run_cmd(['run-macro', case['macro']], wb)
                if not ok_run:
                    tail = " ".join(l.strip() for l in out_run.strip().splitlines()[-3:])
                    results.append((case['name'], False, f"直したマクロが実行できない: {tail[:160]}", info))
                    continue
                bad = []
                for a, want in case['expect'].items():
                    got = ws.Range(a).Value
                    same = (abs(float(got) - want) < 0.5) if isinstance(want, (int, float)) and got is not None \
                        and not isinstance(got, str) else (str(got or '') == str(want))
                    if not same:
                        bad.append(f"{a}={got!r}（期待 {want!r}）")
                ok = (not bad) and r['done']
                note = ("実行して " + "・".join(f"{a}={v}" for a, v in case['expect'].items()) if not bad
                        else "実行結果が違う: " + " ".join(bad)) + ("" if r['done'] else "（AI は done を言っていない）")
                print(f"答え合わせ: {note}")
                print(va._gate_line(info.get('gates')) + f"　＝ {va._gate_grade(ok, info.get('gates'))}")
            except Exception as ex:
                print(f"エラー: {ex}")
                ok, note = False, str(ex)
            results.append((case['name'], ok, note, info))
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception as ex:
            print(f"⚠ 実射用のブックを閉じられませんでした: {ex}")
        for p in [path] + [os.path.join(va.BACKUP_DIR, f) for f in (os.listdir(va.BACKUP_DIR) if os.path.isdir(va.BACKUP_DIR) else [])
                           if f.startswith(os.path.basename(path) + ".backup_before_")]:
            try:
                os.remove(p)
            except Exception:
                pass
    return results


_FIRE_GROUPS = ('sheet', 'macro', 'recipe', 'vague', 'mine', 'weak')   # vague＝曖昧な依頼（2026-09-04）／mine＝本番から拾った弾（2026-09-06）


def _select_fire_cases(all_cases, names):
    """--fire の名指し: 弾の名前か、組の名前（sheet / macro / recipe）。知らない名前は 2 つ目で返す。

    組の名前と弾の名前が混ざったときは弾の名前だけを撃つ（2026-09-04）。
    `--fire recipe 色分け` と書くと組（recipe）が効いて 25 本全部になり、通っている弾まで撃ち直していた
    ＝失敗した弾だけ撃ち直すという決め（時間も API も無駄にしない）を、道具の側が壊していた。
    """
    if not names:
        return list(all_cases), []
    known = {c['name'] for c in all_cases} | set(_FIRE_GROUPS)
    unknown = [n for n in names if n not in known]
    by_name = [n for n in names if any(c['name'] == n for c in all_cases)]
    groups = [n for n in names if n in _FIRE_GROUPS]
    if by_name:
        if groups:
            print(f"（組の名前 {'・'.join(groups)} は無視して、名指しの {len(by_name)} 本だけ撃ちます。"
                  "組ごと撃つなら弾の名前を書かない）")
        picked = [c for c in all_cases if c['name'] in by_name]
    else:
        picked = [c for c in all_cases if c.get('kind', 'sheet') in groups]
    return picked, unknown


def _fire_rows(results):
    """実射の結果を (名前, ok, note, info) の 4 つ組にそろえる（古い 3 つ組も受ける・2026-09-06）。"""
    out = []
    for r in (results or []):
        out.append((r[0], bool(r[1]), r[2], (r[3] if len(r) > 3 and isinstance(r[3], dict) else {})))
    return out


def _fire_rates(rows):
    """同じ弾を何回も撃ったときの、弾ごとの (合格数, 撃った数, 一発合格数)（純 Python・2026-09-06 夜）。"""
    out = {}
    for n, ok, _t, i in rows:
        p, a, f = out.get(n, (0, 0, 0))
        first = bool(ok) and (i.get('gates') is not None) and va._gate_total(i.get('gates')) == 0
        out[n] = (p + int(bool(ok)), a + 1, f + int(first))
    return out


def _first_shot_line(results):
    """「合格」の中身を 3 段に割った 1 行（2026-09-06）。

    合格の本数だけでは、AI が一発で正しかったのか・関所が毎回直したのかが分からない。
    通信簿の信用度はここに出る（59/59 でも全部が関所頼みなら、AI 自体は直っていない）。
    """
    rows = _fire_rows(results)
    if not rows:
        return "一発合格: 数えていません（弾なし）"
    measured = [(ok, i.get('gates')) for _n, ok, _t, i in rows if i.get('gates') is not None and i]
    first = [1 for ok, g in measured if ok and va._gate_total(g) == 0]
    fixed = [1 for ok, g in measured if ok and va._gate_total(g) > 0]
    bad = sum(1 for _n, ok, _t, _i in rows if not ok)
    if not measured:
        return f"一発合格: 数えていません（関所の記録が無い弾ばかり）／不合格 {bad}"
    tally = {}
    for _ok, g in measured:
        for k, v in (g or {}).items():
            if v:
                tally[k] = tally.get(k, 0) + v
    line = (f"合格の中身: 一発合格 {len(first)} / 関所で直して合格 {len(fixed)} / 不合格 {bad}"
            f"（{len(measured)} 本で計測）")
    if tally:
        line += "　差し戻しの内訳: " + " / ".join(
            f"{va._GATE_LABELS.get(k, k)} {v}" for k, v in sorted(tally.items(), key=lambda kv: -kv[1]))
    # 2026-09-06 夜（Claude for Excel の助言の指標）: 範囲外変化ゼロ率＝不変条件に 1 度も引っかからなかった弾、
    # 黙殺率＝不合格なのに report の【できなかったこと】が空（できていないことを言えていない）
    zero_inv = sum(1 for _ok, g in measured if not (g or {}).get('inv'))
    line += f"　範囲外変化ゼロ {zero_inv}/{len(measured)}"
    told = [(ok, i) for _n, ok, _t, i in rows if 'report' in i]
    if any(not ok for ok, _i in told):
        silent = sum(1 for ok, i in told if not ok and _report_is_silent(i))
        line += f"　黙殺（不合格なのに報告に無い）{silent}/{sum(1 for ok, _i in told if not ok)}"
    return line


def _report_is_silent(info):
    """不合格の弾の report が「できなかったこと」を言えていないか（純 Python）。"""
    if info.get('asks') or info.get('dropped'):
        return False
    rep = str(info.get('report') or '')
    m = re.search(r'【できなかったこと[^】]*】(.*?)(?=【|$)', rep, re.S)
    body = (m.group(1) if m else '').strip()
    return not body or re.fullmatch(r'[・\-—\s]*(なし|無し|特になし|ありません|該当なし)?[。\s]*', body) is not None


def _score_load():
    """実射の点数の履歴（古い順）。無ければ空。"""
    out = []
    try:
        with open(_AGENT_SCORE_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        out.append(json.loads(line))
                    except ValueError:
                        continue
    except OSError:
        return []
    return out


def _score_save(rec):
    try:
        with open(_AGENT_SCORE_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(rec, ensure_ascii=False) + chr(10))
    except OSError as ex:
        print(f"（点数を記録できませんでした: {ex}）")


def _score_compare(rec, hist):
    """前回の同じ弾と比べて、落ちた弾（退行）と直った弾を返す。前回が無ければ (None, None)。"""
    cases = rec['cases']
    for old in reversed(hist):
        common = set(cases) & set(old.get('cases') or {})
        if not common:
            continue
        worse = sorted(n for n in common if old['cases'][n] and not cases[n])
        better = sorted(n for n in common if not old['cases'][n] and cases[n])
        return (old, worse, better)
    return (None, None, None)


def score_history(limit=10):
    """点数の移り変わりを出す（agent --score）。何本中何本・いつ・どの AI か。"""
    hist = _score_load()
    if not hist:
        print(f"点数の記録がありません（{_AGENT_SCORE_FILE}）。agent --fire を撃つと残ります。")
        return True
    print(f"実射の点数（新しい順・{_AGENT_SCORE_FILE}）:")
    for r in list(reversed(hist))[:limit]:
        n, k = r.get('n', 0), r.get('ok', 0)
        pct = (100.0 * k / n) if n else 0.0
        fails = [nm for nm, v in (r.get('cases') or {}).items() if not v]
        shot = r.get('first_shot')
        print(f"  {r.get('time', '?')}  {k}/{n}（{pct:.0f}%）  {r.get('ai', '?')}/{r.get('model', '?')}"
              f"  {r.get('sec', 0):.0f}秒"
              + (f"  一発 {shot}/{k}" if shot is not None and k else "")
              + ("  落ちた弾: " + "・".join(fails) if fails else ""))
        for nm in fails:
            if (r.get('why') or {}).get(nm):
                print(f"      └ {nm}: {r['why'][nm][:120]}")
    # 練習台の点数の下に、本番の走行を 1 行（2026-09-06）。上は「私が作った問題の点数」、下は実戦の記録
    line = va._runs_summary(va._runs_load())
    print(line if line else "実運用: まだ走行の記録がありません（本番のブックに回すと溜まります）")
    print("（本番の走行の明細は agent --runs）")
    return True


class _FireTee:
    """実射の画面出力を控えに落とす（画面にはそのまま出す）。落ちた弾を後から読むため（2026-09-06 深夜）。"""

    def __init__(self, path):
        self.path = path
        self.f = None
        self.old = sys.stdout
        try:
            self.f = open(path, 'w', encoding='utf-8')
            sys.stdout = self
        except Exception:
            self.f = None

    def write(self, t):
        self.old.write(t)
        if self.f:
            try:
                self.f.write(t)
                self.f.flush()      # 走っている最中に読めないと、長い実射の様子が誰にも見えない
            except Exception:
                pass

    def flush(self):
        self.old.flush()
        if self.f:
            try:
                self.f.flush()
            except Exception:
                pass

    def close(self):
        sys.stdout = self.old
        if self.f:
            try:
                self.old.write("（実射の画面出力を控えました: " + self.path + "）" + chr(10))
                self.f.close()
            except Exception:
                pass
        self.f = None


def fire_agent(xl, ai=va._CC_AI, model=None, names=None, state=None, seed=None, twice=False, bed=None,
               repeat=1, show_image=True, grade_ai=None, grade_model=None):
    """自動実射: 練習台に依頼を撃ち、現物で答え合わせ。全部合格なら True。

    repeat=N は同じ弾を N 回ずつ撃って弾ごとの合格率を出す（2026-09-06 夜: 同じ依頼を 4 回撃って 3 往復が 3 回・
    2 往復が 1 回＝1 発の合否では見えない揺れを数字にする）。show_image=False・grade_ai は run_agent にそのまま渡す
    （--no-image・--grade-ai の効果を、同じ弾で測る）。

    組: sheet（決まった仕様の依頼）／ vague（曖昧な依頼＝仕様が書かれていない。2026-09-04）／
        macro（壊れたマクロの修理）／ recipe（手順書 22 本・練習台 26 本）／ weak（弱点・止まるべき弾）。
    撃つたびに点数を _agent_score.jsonl に記録し、前回と比べて「退行した弾」を出す（agent --score で履歴）。

    2026-09-05: 弾ごとに「練習台の状態」を 1 つ当てる（隠れ行・前のフィルタ・名前定義・他シート参照…）。
    --state で固定、無指定なら seed から選ぶ。落ちたら出力の seed と状態でそのまま再現できる。
    """
    from vbam_recipes import _fire_recipe_cases
    _tee = _FireTee(_FIRE_OUT_FILE)
    try:
        return _fire_agent_body(xl, ai, model, names, state, seed, twice, bed, repeat,
                                show_image, grade_ai, grade_model)
    finally:
        _tee.close()


def _fire_agent_body(xl, ai=va._CC_AI, model=None, names=None, state=None, seed=None, twice=False, bed=None,
                     repeat=1, show_image=True, grade_ai=None, grade_model=None):
    from vbam_recipes import _fire_recipe_cases
    if bed:
        # 写し取った本物の構造に撃つ（弾の名前はシート名として読む）
        ai, model, _ = va._ai_setup(ai, model)
        t0 = time.time()
        results = _fire_bed(xl, bed, ai, model, names or None, True)
        print(f"\n写し取りの結果（{time.time() - t0:.1f}秒）:")
        for n, ok, note, info in (_fire_rows(results)):
            print(f"  {'PASS' if ok else 'FAIL'}  {n}  {note}")
        n_ok = sum(1 for r in results if r[1])
        print(f"{n_ok}/{len(results)} 合格")
        print(_first_shot_line(results))
        return n_ok == len(results)
    all_cases = FIRE_CASES + _FIRE_MACRO_CASES + RECIPE_CASES
    # 本番から拾った弾（2026-09-06）。名前で名指しできる／組の名前は mine。名前を何も書かないときは撃たない
    #  ＝弾ごとに本物由来の練習台を開くので、全部撃ちの中に黙って混ぜない（撃つのは明示したときだけ）
    saved = _cases_load()
    names = list(names or [])
    mine_pick = [n for n in names if n in saved]
    want_mine = 'mine' in names
    rest = [n for n in names if n not in saved and n != 'mine']
    if (want_mine or mine_pick) and not rest:
        cases, unknown = [], []
    else:
        cases, unknown = _select_fire_cases(all_cases, rest or None)
    if unknown:
        print("エラー: 無い実射: " + ", ".join(unknown))
        print("  ある実射: 組 sheet / macro / recipe / mine か、" + ", ".join(c['name'] for c in all_cases)
              + ("　本番の弾: " + "、".join(sorted(saved)) if saved else ""))
        return False
    if state and state not in _FIRE_STATES:
        print("エラー: 無い状態: " + state + "\n  ある状態: " + " / ".join(_FIRE_STATES))
        return False
    ai, model, _ = va._ai_setup(ai, model)
    results = []
    t0 = time.time()
    # 撃ち方は同じ（weak＝弱点の弾。2026-09-06 に Claude for Excel の失敗型と「止まるべき弾」から足した）
    sheet_cases = [c for c in cases if c.get('kind', 'sheet') in ('sheet', 'vague', 'weak')]
    macro_cases = [c for c in cases if c.get('kind') == 'macro']
    recipe_cases = [c for c in cases if c.get('kind') == 'recipe']
    if seed is None:
        seed = int(time.time())
    states = _pick_states(sheet_cases + recipe_cases, state, seed)
    if sheet_cases or recipe_cases:
        print(f"練習台の状態: " + (f"全部「{state}」" if state else f"seed {seed} で弾ごとに選ぶ")
              + f"（同じ姿でもう一度撃つなら --seed {seed}"
              + (f" --state {state}" if state else "") + "）")
    if twice:
        print("冪等性の検査: 合格した弾にもう 1 回同じ依頼を撃ちます（何も変わらないのが正しい。往復は 2 倍）")
    repeat = max(int(repeat or 1), 1)
    if repeat > 1:
        print(f"同じ弾を {repeat} 回ずつ撃ちます（合否の揺れを弾ごとの合格率で見る）")
    if show_image:
        print("画像ありで撃ちます（--image。既定の画像なしの点数と比べる用）")
    if grade_ai:
        print(f"採点係: {grade_ai}{'/' + grade_model if grade_model else ''}（--grade-ai）")
    kw = dict(show_image=show_image, grade_ai=grade_ai, grade_model=grade_model)
    for k in range(repeat):
        if repeat > 1:
            print(f"\n######## {k + 1} 回目 / {repeat} ########")
        if sheet_cases:
            results += _fire_sheet_cases(xl, sheet_cases, ai, model, states, twice, **kw)
        if macro_cases:
            results += _fire_macro_cases(xl, macro_cases, ai, model)
        if recipe_cases:
            results += _fire_recipe_cases(xl, recipe_cases, ai, model, states, **kw)
        if want_mine or mine_pick:
            results += _fire_mine(xl, ai, model, mine_pick or None, twice)
    rows = _fire_rows(results)
    print(f"\n実射の結果（{time.time() - t0:.1f}秒）:")
    for n, ok, note, info in rows:
        print(f"  {'PASS' if ok else 'FAIL'}  {n}  {note}")
    n_ok = sum(1 for _n, ok, _t, _i in rows if ok)
    print(f"{n_ok}/{len(rows)} 合格")
    print(_first_shot_line(results))
    rates = _fire_rates(rows) if repeat > 1 else None
    if rates:
        print(f"弾ごとの合格率（{repeat} 回）:")
        for n, (n_pass, n_all, n_first) in rates.items():
            print(f"  {n_pass}/{n_all} 合格・一発 {n_first}  {n}" + ("　← 揺れる弾" if 0 < n_pass < n_all else ""))
    n_merged = sum(int(i.get('merged') or 0) for _n, _ok, _t, i in rows)
    secs = [float(i['sec']) for _n, _ok, _t, i in rows if i.get('sec') is not None]
    if secs:
        print(f"1 弾あたり {sum(secs) / len(secs):.1f} 秒（道具＋AI 待ち）"
              + (f"　手と done を同時に受けた回: {n_merged}" if n_merged else ""))
    # 点数を記録して前回と比べる（2026-09-04・「直しても手応えが出ない」＝数字が残っていなかった）
    cases_ok = {}
    why = {}
    for n, ok, t, _i in rows:                        # 同じ弾を何回も撃ったときは「全部合格」で 1 つに畳む
        cases_ok[n] = cases_ok.get(n, True) and bool(ok)
        if not ok:
            # 落ちた理由も残す（2026-09-10）。83 本全滅の行に理由が無く、_fire_last.txt は次の実射で
            # 上書きされる＝撃ち直さないと原因が分からなかった。同じ弾が何回も落ちたら最初の理由
            why.setdefault(n, str(t or '').replace('\n', ' ')[:300])
    rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'ai': ai, 'model': model,
           'n': len(rows), 'ok': n_ok, 'sec': round(time.time() - t0, 1),
           'seed': seed, **({'state': state} if state else {}),
           'cases': cases_ok,
           **({'why': why} if why else {}),
           **({'repeat': repeat, 'rates': {n: f"{p}/{a}" for n, (p, a, _f) in rates.items()}} if rates else {}),
           **({'image': True} if show_image else {}),
           **({'grade_ai': grade_ai, 'grade_model': grade_model} if grade_ai else {}),
           **({'avg_sec': round(sum(secs) / len(secs), 1)} if secs else {}),
           **({'merged': n_merged} if n_merged else {}),
           # 関所の差し戻し（2026-09-06）。「合格」の中身が一発か・関所が直したかを残す
           'gates': {n: {k: v for k, v in (i.get('gates') or {}).items() if v}
                     for n, _ok, _t, i in rows if i.get('gates')},
           'first_shot': sum(1 for _n, ok, _t, i in rows if ok and va._gate_total(i.get('gates')) == 0)}
    prev, worse, better = _score_compare(rec, _score_load())
    if prev is not None:
        pn, pk = prev.get('n', 0), prev.get('ok', 0)
        print(f"前回（{prev.get('time', '?')}）は {pk}/{pn}。"
              + ("退行: " + "・".join(worse) if worse else "退行なし")
              + ("　直った: " + "・".join(better) if better else ""))
    _score_save(rec)
    print(f"（点数を記録しました: {_AGENT_SCORE_FILE}　履歴は agent --score）")
    # 練習台のブックを開くために道具が起こした Excel は、ここで畳む（人が開いていた Excel には触らない）。
    # 常駐の MCP では終了時の後始末が何時間も来ず、非表示の Excel が残っていた（2026-09-04）
    try:
        from vbam_core import release_created_instances
        closed, kept = release_created_instances(only_saved=False)
        if closed:
            print(f"（道具が起こした Excel {closed} 台を閉じました"
                  + (f"・{kept} 台は温存" if kept else "") + "）")
    except Exception as ex:
        print(f"（道具が起こした Excel を畳めませんでした: {ex}）")
    return n_ok == len(results)


__all__ = ['fire_agent', 'FIRE_CASES', 'harvest_book', 'keep_case', 'drop_case', 'cases_list', 'score_history']
