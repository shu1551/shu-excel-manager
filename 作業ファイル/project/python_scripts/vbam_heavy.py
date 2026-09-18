# -*- coding: utf-8 -*-
"""vbam_heavy.py — vba_manager 分割パート: 重量級コマンド（チャート/ピボット/スライサー/計算モード/PowerQuery/データモデル）

vba_manager.py から機械分割（2026-07-12）。単体で実行せず、vba_manager.py 経由で使う。
"""
import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

from vbam_core import *  # noqa: F401,F403
from vbam_vba import *  # noqa: F401,F403
from vbam_view import *  # noqa: F401,F403
from vbam_edit import *  # noqa: F401,F403
# ================================================================
# 重量級コマンド (1) チャート
# ================================================================

_XL_CHART_TYPE = {
    'column':         51,     # xlColumnClustered
    'stacked-column': 52,     # xlColumnStacked
    'stacked-column-100': 53, # xlColumnStacked100
    'bar':            57,     # xlBarClustered
    'stacked-bar':    58,     # xlBarStacked
    'stacked-bar-100': 59,    # xlBarStacked100
    'line':           4,      # xlLine
    'line-markers':   65,     # xlLineMarkers
    'pie':            5,      # xlPie
    'scatter':        -4169,  # xlXYScatter
    'area':           1,      # xlArea
    'stacked-area':   76,     # xlAreaStacked
    'doughnut':       -4120,  # xlDoughnut
    'radar':          -4151,  # xlRadar
}
_XL_CHART_TYPE_NAME = {v: k for k, v in _XL_CHART_TYPE.items()}


@protect_safe
@dialog_safe
def cmd_chart(args):
    """グラフ操作: chart <create|list|delete> ...

      chart create <data_range> [--type column|stacked-column|stacked-column-100|bar|stacked-bar|stacked-bar-100|
                                        line|line-markers|pie|scatter|area|stacked-area|doughnut|radar]
                   [--title "見出し"] [--at セル] [--name 名] [--width N --height N]
      chart create --pivot <ピボット名> [--type ...] [--title ...] [--at セル]   ピボットグラフ
      chart list
      chart delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                cnt += 1
                try:
                    t = _XL_CHART_TYPE_NAME.get(int(co.Chart.ChartType), co.Chart.ChartType)
                except Exception:
                    t = '?'
                print(f"[{sh.Name}] {co.Name}  type={t}")
        if cnt == 0:
            print("グラフはありません。")
        return True

    if action == 'create':
        piv = getattr(args, 'pivot', None)
        if piv:
            ws, pt = _find_pivot(wb, piv)
            if pt is None:
                print(f"エラー: ピボット '{piv}' が見つかりません（pivot list で確認）。"); return False
            rng = pt.TableRange1        # ピボットの出力範囲を元にすると Excel がピボットグラフにする
        else:
            if len(rest) < 2:
                print("使い方: chart create <data_range> [--type ...] [--title ...] [--at セル] [--name 名]  ／  chart create --pivot ピボット名")
                return False
            ws, rng = _resolve_range(xl, wb, rest[1])

        # 配置: --at 指定があればそのセルの左上、なければデータ範囲の右隣
        at = getattr(args, 'at', None)
        if at:
            anchor = ws.Range(at)
            left, top = anchor.Left, anchor.Top
        else:
            left, top = rng.Left + rng.Width + 10, rng.Top
        width = float(getattr(args, 'width', None) or 360)
        height = float(getattr(args, 'height', None) or 216)

        co = ws.ChartObjects().Add(left, top, width, height)
        ch = co.Chart
        ch.SetSourceData(rng)
        ctype = (getattr(args, 'type', None) or 'column').lower()
        if ctype not in _XL_CHART_TYPE:
            print(f"未知のグラフ種別: {ctype}（{'/'.join(_XL_CHART_TYPE)}）")
            co.Delete()
            return False
        ch.ChartType = _XL_CHART_TYPE[ctype]
        if getattr(args, 'title', None):
            ch.HasTitle = True
            ch.ChartTitle.Text = args.title
        if getattr(args, 'name', None):
            # 名前の重複・禁止文字で失敗すると、自動名（グラフ 1 等）のグラフだけが
            # 残って「名前を付けたはずのグラフが chart list に無い」状態になる。
            # 種別不正の経路と同じく、このコマンドが作ったものは片づけてから失敗を返す。
            try:
                co.Name = args.name
            except Exception as ex:
                print(f"エラー: グラフ名 '{args.name}' を設定できません: {ex}")
                print("  同名のグラフが既にある可能性があります（chart list で確認）。")
                try:
                    co.Delete()
                    print("  作成途中のグラフは片づけました。")
                except Exception:
                    pass
                return False
        print(f"グラフ作成: [{ws.Name}] {co.Name}  種別={ctype}  データ={rng.Address}" + ("（ピボットグラフ）" if piv else ""))
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: chart delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    co.Delete()
                    print(f"グラフ削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: グラフ '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


_XL_AXIS = {'category': 1, 'value': 2, 'series': 3, 'secondary': 2}   # xlCategory/xlValue/xlSeriesAxis
_XL_LEGEND_POS = {'bottom': -4107, 'corner': 2, 'top': -4160, 'right': -4152, 'left': -4131}
_XL_TRENDLINE = {'linear': -4132, 'exponential': 5, 'logarithmic': -4133,
                 'movingaverage': 6, 'polynomial': 3, 'power': 4}


@protect_safe
def cmd_chart_config(args):
    """グラフ詳細設定: chart-config <action> <chart名> ...

      set-source <chart> <range>                          データ範囲を再設定
      set-type <chart> <type>                             種別変更(chart create の --type と同じ語。外れれば一覧が出る)
      set-title <chart> <text>                            グラフタイトル
      set-axis-title <chart> <category|value|secondary> <text>   軸タイトル
      axis-format <chart> <axis> [format]                 軸の表示形式 get/set
      axis-scale <chart> <axis> [--min N --max N --major N --minor N]  軸目盛
      gridlines <chart> <axis> [--major on|off --minor on|off]        目盛線
      legend <chart> <bottom|top|right|left|corner|off>   凡例
      style <chart> <1-48>                                組込スタイル
      placement <chart> <1|2|3>                           1=移動+サイズ/2=移動のみ/3=自由
      data-labels <chart> [--value --percent --category --series --position 位置]
      add-series <chart> <values_range> [--series-name 名] [--category-range 範囲]
      remove-series <chart> <index>
      series-format <chart> <index> [--marker-style N --marker-size N --marker-fg #.. --marker-bg #.. --invert]
      trendline list <chart> <series_index>
      trendline add  <chart> <series_index> <type> [--name 名]
      trendline delete <chart> <series_index> <trendline_index>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart-config <action> <chart名> ...（詳細は --help）")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_chart(name):
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    return sh, co, co.Chart
        return None, None, None

    # trendline はサブ動詞が rest[1] に来る特例
    if action == 'trendline':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        cname = rest[2] if len(rest) >= 3 else None
        sh, co, ch = find_chart(cname) if cname else (None, None, None)
        if not ch:
            print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False
        try:
            sidx = int(rest[3]) if len(rest) >= 4 else 1
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        s = ch.SeriesCollection(sidx)
        if sub == 'list':
            tls = s.Trendlines()
            print(f"--- {cname} 系列{sidx} の近似曲線 ({tls.Count}) ---")
            for i in range(1, tls.Count + 1):
                tl = tls.Item(i)
                try:
                    nm = tl.Name
                except Exception:
                    nm = f"#{i}"
                print(f"  [{i}] {nm}")
            if tls.Count == 0:
                print("  (なし)")
            return True
        if sub == 'add':
            ttype = (rest[4] if len(rest) >= 5 else 'linear').lower()
            if ttype not in _XL_TRENDLINE:
                print(f"未知の種別: {ttype}（{'/'.join(_XL_TRENDLINE)}）"); return False
            tl = s.Trendlines().Add(_XL_TRENDLINE[ttype])
            if getattr(args, 'name', None):
                tl.Name = args.name
            print(f"近似曲線追加: {cname} 系列{sidx} {ttype}")
            print("（保存はしていません）"); return True
        if sub == 'delete':
            # 削除系のインデックス省略は「黙って#1が消える」事故のもと。明示必須。
            if len(rest) < 5:
                print("使い方: chart-config trendline delete <chart> <series_index> <trendline_index>")
                print("  （削除対象の trendline_index は省略できません。trendline list で確認）")
                return False
            try:
                tidx = int(rest[4])
            except ValueError:
                print("trendline_index は数値で指定してください。"); return False
            s.Trendlines().Item(tidx).Delete()
            print(f"近似曲線削除: {cname} 系列{sidx} #{tidx}")
            print("（保存はしていません）"); return True
        print("使い方: chart-config trendline <list|add|delete> <chart> <series_index> ...")
        return False

    # それ以外は rest[1] が chart 名
    cname = rest[1] if len(rest) >= 2 else None
    sh, co, ch = find_chart(cname) if cname else (None, None, None)
    if not ch:
        print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False

    def get_axis(axname):
        a = (axname or 'value').lower()
        if a not in _XL_AXIS:
            return None
        if a == 'secondary':
            return ch.Axes(2, 2)            # xlValue, xlSecondary
        return ch.Axes(_XL_AXIS[a])

    if action == 'set-source':
        if len(rest) < 3:
            print("使い方: chart-config set-source <chart> <range>"); return False
        ws, rng = _resolve_range(xl, wb, rest[2])
        ch.SetSourceData(rng)
        print(f"データ範囲再設定: {cname} ← {rng.Address}")
        print("（保存はしていません）"); return True

    if action == 'set-type':
        t = (rest[2] if len(rest) >= 3 else '').lower()
        if t not in _XL_CHART_TYPE:
            print(f"未知の種別: {t}（{'/'.join(_XL_CHART_TYPE)}）"); return False
        ch.ChartType = _XL_CHART_TYPE[t]
        print(f"種別変更: {cname} → {t}"); print("（保存はしていません）"); return True

    if action == 'set-title':
        txt = rest[2] if len(rest) >= 3 else ''
        ch.HasTitle = True
        ch.ChartTitle.Text = txt
        print(f"タイトル設定: {cname} = {txt}"); print("（保存はしていません）"); return True

    if action == 'set-axis-title':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        txt = rest[3] if len(rest) >= 4 else ''
        ax.HasTitle = True
        ax.AxisTitle.Text = txt
        print(f"軸タイトル設定: {cname} {rest[2]} = {txt}"); print("（保存はしていません）"); return True

    if action == 'axis-format':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        if len(rest) >= 4:
            ax.TickLabels.NumberFormat = rest[3]
            print(f"軸表示形式設定: {cname} {rest[2]} = {rest[3]}")
            print("（保存はしていません）"); return True
        else:
            print(f"軸表示形式: {cname} {rest[2]} = {ax.TickLabels.NumberFormat}"); return True

    if action == 'axis-scale':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        changed = []
        for opt, prop in (('min', 'MinimumScale'), ('max', 'MaximumScale'),
                          ('major', 'MajorUnit'), ('minor', 'MinorUnit')):
            v = getattr(args, opt, None)
            if v is not None:
                setattr(ax, prop, float(v)); changed.append(f"{opt}={v}")
        if changed:
            print(f"軸目盛設定: {cname} {rest[2]} [{', '.join(changed)}]")
            print("（保存はしていません）"); return True
        else:
            print(f"軸目盛: {cname} {rest[2]} min={ax.MinimumScale} max={ax.MaximumScale} "
                  f"major={ax.MajorUnit} minor={ax.MinorUnit}"); return True

    if action == 'gridlines':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        mj = getattr(args, 'major', None); mn = getattr(args, 'minor', None)
        if mj is None and mn is None:
            print(f"目盛線: {cname} {rest[2]} major={ax.HasMajorGridlines} minor={ax.HasMinorGridlines}")
            return True
        if mj is not None:
            ax.HasMajorGridlines = (mj.lower() == 'on')
        if mn is not None:
            ax.HasMinorGridlines = (mn.lower() == 'on')
        print(f"目盛線設定: {cname} {rest[2]} major={getattr(args,'major',None)} minor={getattr(args,'minor',None)}")
        print("（保存はしていません）"); return True

    if action == 'legend':
        pos = (rest[2] if len(rest) >= 3 else 'bottom').lower()
        if pos == 'off':
            ch.HasLegend = False
            print(f"凡例: {cname} = 非表示")
        else:
            if pos not in _XL_LEGEND_POS:
                print(f"位置は {'/'.join(_XL_LEGEND_POS)}|off で指定してください。"); return False
            ch.HasLegend = True
            ch.Legend.Position = _XL_LEGEND_POS[pos]
            print(f"凡例: {cname} = {pos}")
        print("（保存はしていません）"); return True

    if action == 'style':
        try:
            sid = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config style <chart> <1-48の数値>"); return False
        ch.ChartStyle = sid
        print(f"スタイル設定: {cname} = {sid}"); print("（保存はしていません）"); return True

    if action == 'placement':
        try:
            pl = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config placement <chart> <1|2|3>"); return False
        co.Placement = pl
        names = {1: '移動+サイズ', 2: '移動のみ', 3: '自由配置'}
        print(f"配置方法: {cname} = {pl}（{names.get(pl, pl)}）"); print("（保存はしていません）"); return True

    if action == 'data-labels':
        ch.ApplyDataLabels(
            ShowValue=bool(getattr(args, 'value', False)),
            ShowPercentage=bool(getattr(args, 'percent', False)),
            ShowCategoryName=bool(getattr(args, 'category', False)),
            ShowSeriesName=bool(getattr(args, 'series', False)))
        pos = getattr(args, 'position', None)
        if pos:
            posmap = {'center': -4108, 'insideend': 3, 'outsideend': 2, 'bestfit': 5, 'insidebase': 4}
            if pos.lower() not in posmap:
                print(f"⚠ 未知の位置: {pos}（{'/'.join(posmap)}）位置指定はスキップしました。")
            else:
                # 全系列に適用（以前は系列1のみで、複数系列だと部分適用のまま成功表示だった）
                pos_failed = []
                for si in range(1, ch.SeriesCollection().Count + 1):
                    try:
                        ch.SeriesCollection(si).DataLabels().Position = posmap[pos.lower()]
                    except Exception:
                        pos_failed.append(si)
                if pos_failed:
                    print(f"⚠ 位置指定が適用できなかった系列: {pos_failed}"
                          "（グラフ種別によって位置指定不可の場合があります）")
        print(f"データラベル設定: {cname}"); print("（保存はしていません）"); return True

    if action == 'add-series':
        ws, vrng = _resolve_range(xl, wb, rest[2])
        s = ch.SeriesCollection().NewSeries()
        s.Values = vrng
        if getattr(args, 'series_name', None):
            s.Name = args.series_name
        if getattr(args, 'category_range', None):
            _, crng = _resolve_range(xl, wb, args.category_range)
            s.XValues = crng
        print(f"系列追加: {cname} ← {vrng.Address}"); print("（保存はしていません）"); return True

    if action == 'remove-series':
        # 削除系のインデックス省略は「黙って系列1が消える」事故のもと。明示必須。
        if len(rest) < 3:
            print("使い方: chart-config remove-series <chart> <series_index>")
            print("  （削除対象の series_index は省略できません）")
            return False
        try:
            idx = int(rest[2])
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        ch.SeriesCollection(idx).Delete()
        print(f"系列削除: {cname} #{idx}"); print("（保存はしていません）"); return True

    if action == 'series-format':
        # 省略時は従来どおり系列1に適用する。出力に「#番号」を明示するので
        # 無言の書き換えにはならない（削除系の明示必須とは事情が違う）
        try:
            idx = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        s = ch.SeriesCollection(idx)
        ch_list = []
        if getattr(args, 'marker_style', None) is not None:
            s.MarkerStyle = int(args.marker_style); ch_list.append('style')
        if getattr(args, 'marker_size', None) is not None:
            s.MarkerSize = int(args.marker_size); ch_list.append('size')
        if getattr(args, 'marker_fg', None):
            s.MarkerForegroundColor = _hex_to_excel_color(args.marker_fg); ch_list.append('fg')
        if getattr(args, 'marker_bg', None):
            s.MarkerBackgroundColor = _hex_to_excel_color(args.marker_bg); ch_list.append('bg')
        if getattr(args, 'invert', False):
            s.InvertIfNegative = True; ch_list.append('invert')
        if not ch_list:
            # 何も指定が無いのに「系列書式: …[]」と成功表示するのは実態と食い違う
            print("エラー: 変更する書式オプションが指定されていません。")
            print("  --marker-style/--marker-size/--marker-fg/--marker-bg/--invert のいずれかを指定してください。")
            return False
        print(f"系列書式: {cname} #{idx} [{', '.join(ch_list)}]"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (2) ピボットテーブル
# ================================================================

_XL_PIVOT_FUNC = {'sum': -4157, 'count': -4112, 'average': -4106,
                  'max': -4136, 'min': -4139}


def _unique_sheet_name(wb, base):
    """重複しないシート名を返す"""
    existing = {sh.Name for sh in wb.Sheets}
    if base not in existing:
        return base
    i = 2
    while f"{base}{i}" in existing:
        i += 1
    return f"{base}{i}"


@protect_safe
@dialog_safe
def cmd_pivot(args):
    """ピボット操作: pivot <create|list|delete> ...

      pivot create <data_range> [--rows F1,F2] [--cols F1] [--filter F1] [--values F1,F2]
                   [--func sum|count|average|max|min] [--sheet 出力シート] [--at セル] [--name 名]
                   （--sheet と --at の両方なら そのシートのそのセル。既存のデータの上には置かない）
      pivot create model --rows T売上.地域 [--cols T売上.商品] [--measures 売上合計,平均単価] [--values T売上.金額]
                   データモデルから作る（テーブル名.列名／メジャー名。--values は暗黙のメジャー）
      pivot list
      pivot delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                cnt += 1
                print(f"[{sh.Name}] {pt.Name}")
        if cnt == 0:
            print("ピボットテーブルはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: pivot delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                if pt.Name == name:
                    pt.TableRange2.Clear()
                    print(f"ピボット削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: ピボット '{name}' が見つかりません")
        return False

    if action == 'create':
        if len(rest) < 2:
            print("使い方: pivot create <data_range> [--rows ..][--cols ..][--values ..][--func ..]")
            return False
        is_model = bool(getattr(args, 'model', False)) or rest[1].lower() == 'model'
        if is_model:
            ws_s, rng = xl.ActiveSheet, None
        else:
            ws_s, rng = _resolve_range(xl, wb, rest[1])

        # 出力先の決定: --sheet [--at] > --at（元データと同じシートの空き場所） > 新規シート
        sheet_opt = getattr(args, 'sheet', None)
        at = getattr(args, 'at', None)
        created_sheet = None      # このコマンドが新規に作ったシート（失敗時の後始末用）
        if sheet_opt:
            dws = None
            for sh in wb.Sheets:
                # Excel のシート名重複判定は大文字小文字を区別しない
                if sh.Name.lower() == sheet_opt.lower():
                    dws = sh; break
            if dws is None:
                dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
                created_sheet = dws
                try:
                    dws.Name = sheet_opt
                except Exception as ex:
                    # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                    print(f"エラー: シート名 '{sheet_opt}' を設定できません: {ex}")
                    xl.DisplayAlerts = False
                    try:
                        dws.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    return False
            dest = dws.Range(at) if at else dws.Range("A3")
        elif at:
            dws = ws_s
            dest = ws_s.Range(at)
        else:
            dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            dws.Name = _unique_sheet_name(wb, "ピボット")
            created_sheet = dws
            dest = dws.Range("A3")
        if created_sheet is None:
            # 既存シートに置くとき: 元データの上や隣に置くと Excel が「上書きしますか」の窓を出して止まる
            # （315 秒・2026-09-04）。pywin32 の動的ディスパッチでは Range.Resize が右下 1 セルを返し
            # Application.Intersect は None を返すので、Cells で右下 40 行×20 列の矩形を作って CountA を
            # 直に当て、窓が出る前にこちらで断る
            probe = dws.Range(dest, dws.Cells(int(dest.Row) + 39, int(dest.Column) + 19))
            if int(xl.WorksheetFunction.CountA(probe)) > 0:
                print(f"エラー: {dws.Name}!{str(dest.Address).replace('$', '')} の右下（{probe.Address}）に既存のデータがあります。"
                      "ピボットは空き場所に置くか、--sheet 出力シート名 で別シートへ出してください。")
                return False

        pt = None
        pc = None
        funcname = (getattr(args, 'func', None) or 'sum').lower()
        values = getattr(args, 'values', None)
        measures = getattr(args, 'measures', None)
        name = getattr(args, 'name', None)
        try:
            if is_model:
                try:
                    conn = wb.Connections("ThisWorkbookDataModel")
                except Exception:
                    print("エラー: データモデルにテーブルがありません（powerquery load --to model で載せる。datamodel list で確認）。")
                    return False
                pc = wb.PivotCaches().Create(2, conn, 5)    # xlExternal=2 / xlPivotTableVersion15=5
            else:
                pc = wb.PivotCaches().Create(1, rng)        # xlDatabase=1
            if name:
                pt = pc.CreatePivotTable(dest, name)
            else:
                pt = pc.CreatePivotTable(dest)

            def set_fields(spec, orient):
                if not spec:
                    return
                for f in spec.split(','):
                    f = f.strip()
                    if f:
                        _pivot_field_of(pt, f, is_model).Orientation = orient

            set_fields(getattr(args, 'rows', None), 1)   # xlRowField
            set_fields(getattr(args, 'cols', None), 2)   # xlColumnField
            set_fields(getattr(args, 'filter', None), 3)   # xlPageField

            func = _XL_PIVOT_FUNC.get(funcname, -4157)
            if values:
                for f in values.split(','):
                    f = f.strip()
                    if not f:
                        continue
                    cap = _unique_caption(pt, f"{funcname}/{f}")
                    if is_model:
                        # データモデルの列は「暗黙のメジャー」を作ってから値に置く
                        cf = pt.CubeFields.GetMeasure(_cube_name(f), func, cap)
                        pt.AddDataField(cf, cap)
                    else:
                        pt.AddDataField(pt.PivotFields(f), cap, func)
            if measures:
                if not is_model:
                    raise ValueError("--measures はデータモデルのピボット（pivot create model）だけ")
                for m in measures.split(','):
                    m = m.strip()
                    if m:
                        pt.AddDataField(pt.CubeFields(_cube_name(m, measure=True)), m)
            if is_model and (values or measures) and int(pt.TableRange2.Columns.Count) < 2:
                # メジャーが評価できない（列が「文字」型で SUM が失敗）と Excel は無言で値を落とす
                print("⚠ 値（メジャー）がピボットに出ていません。モデルの列が「文字」型だと SUM 系のメジャーは無言で消えます。"
                      " datamodel list で列の型を見て、powerquery edit（道具が型付けを足す）→ powerquery refresh → pivot-calc refresh。")
        except Exception as ex:
            # 存在しないフィールド名等で途中失敗すると、追加した新規シートと
            # 空ピボットが残骸として残り、再実行のたび「ピボット2/3…」と増殖する。
            # このコマンドが作ったものだけ片づける（既存シートのセルは絶対に消さない）
            print(f"エラー: ピボット作成に失敗しました: {ex}")
            print("  --rows/--cols/--filter/--values のフィールド名がデータ範囲の見出しと一致しているか確認してください。")
            if is_model:
                print("  データモデルのピボットは テーブル名.列名（例 T売上.地域）、値は --measures メジャー名（datamodel list で確認）。")
            try:
                if created_sheet is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_sheet.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    print("  作成途中のシート（このコマンドが作ったもの）は片づけました。")
                elif pt is not None:
                    # TableRange2 はこのピボット自身の出力範囲。作成途中の空ピボットを
                    # 残すと再実行のたび「ピボット2/3…」と残骸が増殖する（pivot delete と
                    # 同じ後始末を自動でやるだけ＝既存セルは巻き込まない）
                    pt.TableRange2.Clear()
                    print("  作成途中の空ピボットは片づけました。")
            except Exception:
                pass
            # PivotCache の後始末。参照が外れたキャッシュは保存時に破棄されるが、
            # Delete を持つバージョンでは明示的に落としておく（孤児キャッシュ対策）。
            try:
                if pc is not None:
                    pc.Delete()
            except Exception:
                pass
            pc = None
            return False

        src = "データモデル" if is_model else f"{ws_s.Name}!{rng.Address}"
        print(f"ピボット作成: [{pt.Parent.Name}] {pt.Name}  ソース={src}")
        try:
            caps = [str(pt.DataFields.Item(i).Name) for i in range(1, int(pt.DataFields.Count) + 1)]
        except Exception:
            caps = []
        print(f"  行={getattr(args,'rows',None) or '-'}  列={getattr(args,'cols',None) or '-'}  "
              f"フィルタ={getattr(args,'filter',None) or '-'}  値={values or '-'}({funcname})"
              + (f"  メジャー={measures}" if measures else "")
              + (f"  値フィールド名={', '.join(caps)}（show-as・set-format はこの名前で）" if caps else ""))
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


def _find_pivot(wb, name):
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return sh, pt
    return None, None


_XL_PIVOT_ORIENT = {'row': 1, 'col': 2, 'column': 2, 'filter': 3, 'page': 3, 'value': 4, 'data': 4, 'hidden': 0}

# 値フィールドの「計算の種類」（値フィールドの設定 → 計算の種類。XlPivotFieldCalculation）
_XL_SHOW_AS = {'none': -4143, 'percent_total': 8, 'percent_row': 6, 'percent_col': 7, 'percent_parent_row': 12,
               'percent_parent_col': 13, 'percent_of': 3, 'diff_from': 2, 'percent_diff_from': 4,
               'running_total': 5, 'percent_running_total': 14, 'rank_asc': 15, 'rank_desc': 16, 'index': 9}
_SHOW_AS_BASE_FIELD = ('percent_of', 'diff_from', 'percent_diff_from', 'running_total', 'percent_running_total',
                       'rank_asc', 'rank_desc')          # 基準フィールドが要る種類
_SHOW_AS_BASE_ITEM = ('percent_of', 'diff_from', 'percent_diff_from')   # 基準項目（前の値 など）も要る種類


def _cube_name(field, measure=False):
    """データモデルのピボットのフィールド名: 'T売上.地域' → '[T売上].[地域]'、メジャー名 → '[Measures].[名]'。[ ] 付きはそのまま"""
    f = str(field).strip()
    if f.startswith('['):
        return f
    if measure or '.' not in f:
        return f"[Measures].[{f}]"
    t, c = f.split('.', 1)
    return f"[{t}].[{c}]"


_POSITION_ITEMS = ('(previous)', '(next)', '(前の値)', '(次の値)')


def _position_calc_fields(pt):
    """位置参照（前の値との差・累計）の値フィールド名。これがあると自動並べ替え・上位 N は Excel が
    「オフにしますか」の窓を出して止まる（2026-09-04 実射・10 分停止）"""
    out = []
    try:
        for i in range(1, int(pt.DataFields.Count) + 1):
            d = pt.DataFields.Item(i)
            try:
                c = int(d.Calculation)
            except Exception:
                continue
            if c in (5, 14):
                out.append(str(d.Name))
            elif c in (2, 3, 4):
                try:
                    if str(d.BaseItem) in _POSITION_ITEMS:
                        out.append(str(d.Name))
                except Exception:
                    out.append(str(d.Name))
    except Exception:
        pass
    return out


def _turn_off_autosort_and_topn(pt):
    """行・列フィールドの自動並べ替えと上位 N を外す（位置参照の計算を当てる前に。外したものを返す）"""
    done = []
    # 行・列だけでなく全フィールドを見る（月でまとめると「月 (日付)」と元の「日付」が別のフィールドになり、
    # 並べ替えは元の側に残る＝行・列だけ見ると見落として窓が出る・2026-09-04 実射）
    for coll in (pt.PivotFields(),):
        try:
            n = int(coll.Count)
        except Exception:
            continue
        for i in range(1, n + 1):
            f = coll.Item(i)
            try:
                if int(f.AutoSortOrder) != -4135:          # xlManual
                    f.AutoSort(-4135, f.Name)
                    done.append(f"{f.Name} の並べ替え")
            except Exception:
                pass
            try:
                if int(f.PivotFilters.Count) and int(f.PivotFilters.Item(1).FilterType) in (1, 2):
                    f.ClearAllFilters()
                    done.append(f"{f.Name} の上位/下位 N")
            except Exception:
                pass
    return done


def _unique_caption(pt, base):
    """値フィールドの表示名が既にあれば 2, 3… を付ける（同じ列を 2 本値に置くと同名で落ちる・2026-09-04 実射）"""
    try:
        names = {str(pt.DataFields.Item(i).Name) for i in range(1, int(pt.DataFields.Count) + 1)}
    except Exception:
        names = set()
    if base not in names:
        return base
    k = 2
    while f"{base}{k}" in names:
        k += 1
    return f"{base}{k}"


def _pivot_is_olap(pt):
    try:
        return bool(pt.PivotCache().OLAP)
    except Exception:
        return False


def _pivot_field_of(pt, field, is_olap):
    """フィールド名 → PivotField（通常）／CubeField（データモデル）。無ければ例外"""
    if not is_olap:
        return pt.PivotFields(field)
    f = str(field).strip()
    cands = [f] if f.startswith('[') else (([_cube_name(f)] if '.' in f else []) + [_cube_name(f, measure=True)])
    last = None
    for cand in cands:
        try:
            return pt.CubeFields(cand)
        except Exception as ex:
            last = ex
    raise ValueError(f"フィールド '{field}' が見つかりません（データモデルのピボットは テーブル名.列名 かメジャー名。{last}）")


def _data_field_names(d):
    """値フィールドの呼ばれ方（表示名・元の列名・"sum/売上" の右側・DAX の [名前]）。"""
    out = []
    for attr in ('Name', 'SourceName', 'Caption'):
        try:
            v = getattr(d, attr, None)
        except Exception:
            v = None
        if v:
            v = str(v)
            out.append(v)
            out.append(v.split('/')[-1])
            if v.endswith(']') and '[' in v:
                out.append(v[v.rindex('[') + 1:-1])
    return [x.strip() for x in out if x]


def _row_fields(pt):
    """ピボットの行フィールドを並び順で返す（読めなければ空）。"""
    out = []
    try:
        for i in range(1, pt.PivotFields().Count + 1):
            f = pt.PivotFields().Item(i)
            try:
                if int(f.Orientation) == 1:
                    out.append(f)
            except Exception:
                continue
    except Exception:
        return []
    return out


def _guess_row_field(pt, action):
    """field を書かずに来た手のために、当てられる行フィールド名を返す（当てられなければ None）。

    group-date は日付の行フィールド（1 本だけのときに限る）。
    それ以外（top-n / sort / filter-clear / position）は行フィールドが 1 本だけのときに限る。
    迷う余地があるときは当てない＝黙って違うフィールドに当てるより、聞き返すほうが安い。
    """
    import datetime
    rows = _row_fields(pt)
    if not rows:
        return None
    if action == 'group-date':
        dated = []
        for f in rows:
            try:
                v = f.DataRange.Cells(1, 1).Value
            except Exception:
                v = None
            name = str(getattr(f, 'Name', ''))
            if isinstance(v, datetime.datetime) or isinstance(v, datetime.date)                     or any(w in name for w in ('日付', '年月日', '日時')) or 'date' in name.lower():
                dated.append(name)
        return dated[0] if len(dated) == 1 else None
    return str(rows[0].Name) if len(rows) == 1 else None


@protect_safe
@dialog_safe
def cmd_pivot_field(args):
    """ピボットのフィールド管理: pivot-field <action> <pivot名> <フィールド> ...

      list <pivot>
      add-row|add-col|add-filter <pivot> <field>
      add-value <pivot> <field> [--func sum|count|average|max|min] [--name 表示名]
      remove <pivot> <field>
      set-func <pivot> <field> <func>            データフィールドの集計関数
      set-name <pivot> <field> <表示名>          データフィールドの表示名
      set-format <pivot> <field> <書式コード>    数値書式
      set-filter <pivot> <field> 値1 値2 ...     表示する値を限定
      sort <pivot> <field> <asc|desc> [--by 値フィールド]   ラベル順（既定）か値の順
      group-date <pivot> <field> <days|months|quarters|years>
      group-numeric <pivot> <field> <start> <end> <interval>
      show-as <pivot> <値フィールド> <種類> [--base-field F] [--base-item I]   計算の種類（percent_total|percent_row|
                percent_col|percent_parent_row|percent_parent_col|percent_of|diff_from|percent_diff_from|running_total|
                percent_running_total|rank_asc|rank_desc|index|none。基準の既定＝先頭の行フィールド／前の値）
      top-n <pivot> <field> <N> [--by 値フィールド] [--bottom]   上位/下位 N 件に絞る
      filter-clear <pivot> <field>              絞り込みを解除
      position <pivot> <field> <N>              同じ区画の中の順番
      （データモデルのピボットはフィールドを テーブル名.列名、メジャーはメジャー名で指す）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-field <action> <pivot名> <field> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)
    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません（pivot list で確認）。"); return False

    is_olap = _pivot_is_olap(pt)

    def pf(field):
        try:
            return _pivot_field_of(pt, field, is_olap)
        except Exception:
            return None

    def data_field(name):
        """値フィールドを 表示名／元の列名／'sum/売上' の右側 で探す（無ければ None）"""
        # 名前の呼び方は往復のあいだに変わる（表示名を変えた後・同じ列を 2 本置いた 2 本目の "sum/売上2"）。
        # 厳しく照合していたので、set-format が「値フィールドが見つかりません」で往復を落としていた
        # （2026-09-06 深夜の実射・前月比の弾で 2 回）。呼び方の揺れをここで吸う。
        want = str(name or '').strip()
        fields = []
        for i in range(1, pt.DataFields.Count + 1):
            try:
                fields.append(pt.DataFields.Item(i))
            except Exception:
                continue
        for d in fields:                          # そのままの名前・元の列名・"sum/売上" の右側
            if want in _data_field_names(d):
                return d
        m = re.match(r'^(.*?)([0-9]+)$', want)    # "売上2" "sum/売上2" ＝ 同じ列の 2 本目
        if m:
            base, idx = m.group(1).strip().split('/')[-1], int(m.group(2))
            same = [d for d in fields if base in _data_field_names(d)]
            if 1 <= idx <= len(same):
                return same[idx - 1]
        return None

    if action == 'list':
        print(f"--- {pname} のフィールド ---" + ("（データモデル）" if is_olap else ""))
        orient_name = {1: '行', 2: '列', 3: 'フィルタ', 4: '値', 0: '未配置'}
        for i in range(1, pt.PivotFields().Count + 1):
            f = pt.PivotFields().Item(i)
            try:
                o = int(f.Orientation)
            except Exception:
                o = 0
            print(f"  {f.Name}: {orient_name.get(o, o)}")
        return True

    field = rest[2] if len(rest) >= 3 else None
    if (not field or field == '*') and action in ('group-date', 'top-n', 'sort', 'filter-clear', 'position'):
        # フィールド名を書かずに来たとき、当てられるなら当てる（2026-09-06 深夜の実射で、
        # group_date / top_n に field を書き落として 1 往復まるごと捨てる型が 4 弾で出た）
        guess = _guess_row_field(pt, action)
        if guess:
            field = guess
            print(f"（フィールド名が無いので行フィールド「{field}」に当てました）")
    if action in ('add-row', 'add-col', 'add-filter'):
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = {'add-row': 1, 'add-col': 2, 'add-filter': 3}[action]
        print(f"フィールド配置: {pname}[{field}] = {action[4:]}"); print("（保存はしていません）"); return True

    if action == 'add-value':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        fn = (getattr(args, 'func', None) or 'sum').lower()
        func = _XL_PIVOT_FUNC.get(fn, -4157)
        cname = _unique_caption(pt, getattr(args, 'name', None) or f"{fn}/{field}")
        if is_olap:
            try:
                is_measure = int(p.CubeFieldType) == 2      # xlMeasure
            except Exception:
                is_measure = False
            before = int(pt.TableRange2.Columns.Count)
            if is_measure:
                pt.AddDataField(p, cname)                     # メジャーはそのまま値へ
            else:
                pt.AddDataField(pt.CubeFields.GetMeasure(p.Name, func, cname), cname)   # 列は暗黙のメジャーに
            if int(pt.TableRange2.Columns.Count) <= before:
                print(f"エラー: 値フィールド {cname} がピボットに出ません。モデルの列が「文字」型だと SUM 系のメジャーは無言で消えます"
                      "（datamodel list で型を確認 → powerquery edit で型付け → powerquery refresh → pivot-calc refresh）。")
                return False
        else:
            pt.AddDataField(p, cname, func)
        print(f"値フィールド追加: {pname}[{field}] ({fn}) 表示名={cname}"); print("（保存はしていません）"); return True

    if action == 'remove':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = 0     # xlHidden
        print(f"フィールド除外: {pname}[{field}]"); print("（保存はしていません）"); return True

    if action == 'set-func':
        fn = (rest[3] if len(rest) >= 4 else 'sum').lower()
        if fn not in _XL_PIVOT_FUNC:
            print(f"未知の関数: {fn}（{'/'.join(_XL_PIVOT_FUNC)}）"); return False
        # データフィールドは表示名で参照されるため DataFields を走査
        target = None
        for i in range(1, pt.DataFields.Count + 1):
            d = pt.DataFields.Item(i)
            if d.Name == field or d.SourceName == field:
                target = d; break
        if target is None:
            print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False
        target.Function = _XL_PIVOT_FUNC[fn]
        print(f"集計関数変更: {pname}[{field}] = {fn}"); print("（保存はしていません）"); return True

    if action == 'set-name':
        newname = rest[3] if len(rest) >= 4 else None
        if not newname:
            print("使い方: pivot-field set-name <pivot> <field> <新しい表示名>")
            return False
        # 値フィールド（同じ列を 2 本置くと 'sum/売上2' の名前になる）も表示名を変えられる。
        # PivotFields だけを見ていたので、AI が 2 本目を '売上2' と呼んだ往復が丸ごと落ちた（2026-09-06 深夜）
        # 値フィールドを先に見る（「表示名」はふつう値の見出しのこと。元の列を指す名前でも、
        # それが値に置かれているならそちらを直す・2026-09-06 深夜）
        p = data_field(field) or pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Caption = newname
        print(f"表示名変更: {pname}[{field}] → {newname}"); print("（保存はしていません）"); return True

    if action == 'set-format':
        code = rest[3] if len(rest) >= 4 else None
        if not code:
            print('使い方: pivot-field set-format <pivot> <field> <書式コード（例: "#,##0"）>')
            return False
        d = data_field(field)
        if d is not None:
            d.NumberFormat = code
            print(f"値フィールド書式: {pname}[{field}] = {code}"); print("（保存はしていません）"); return True
        p = pf(field)                        # 値でなく行・列のフィールドを指していたとき（数値の行ラベル）
        if p is not None:
            try:
                p.NumberFormat = code
                print(f"フィールド書式: {pname}[{field}] = {code}"); print("（保存はしていません）"); return True
            except Exception:
                pass
        print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False

    if action == 'set-filter':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        wanted = set(rest[3:])
        if not wanted:
            print("表示する値を1つ以上指定してください。"); return False
        # 指定値が実在するか先に照合（全部タイポだと Excel が「全項目非表示」を拒否し、
        # 実状態と成功メッセージが食い違うため）
        item_names = []
        for i in range(1, p.PivotItems().Count + 1):
            item_names.append(p.PivotItems().Item(i).Name)
        missing = wanted - set(item_names)
        if missing:
            print(f"エラー: 存在しない値が指定されています: {sorted(missing)}")
            print(f"  このフィールドの値: {item_names}")
            return False
        failed = []
        # Excel は「表示ゼロ項目」を許さない。一巡ループで上から順に
        # Visible = (名前 in wanted) を代入すると、wanted の項目がまだ非表示のまま
        # 最後の表示項目を隠そうとして失敗し、その項目が表示に残る
        # （例: 項目[A,B,C]でA・B表示中に C だけを指定 → B が消せず B と C が表示）。
        # 先に「表示するもの」を出してから、「隠すもの」を隠す2パスにする。
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            if it.Name in wanted:
                try:
                    it.Visible = True
                except Exception:
                    failed.append(it.Name)
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            if it.Name not in wanted:
                try:
                    it.Visible = False
                except Exception:
                    failed.append(it.Name)
        if failed:
            print(f"⚠ 一部の項目の表示切替に失敗しました: {failed}")
            print("  （Excel の制約: 全項目非表示は不可、など。実際の表示状態を確認してください）")
        print(f"値フィルタ: {pname}[{field}] = {sorted(wanted)}"); print("（保存はしていません）"); return True

    if action == 'sort':
        order = (rest[3] if len(rest) >= 4 else 'asc').lower()
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        by = getattr(args, 'by', None)
        key = p.Name                                   # 既定はラベル順
        pos = _position_calc_fields(pt)
        if pos:
            print(f"エラー: 位置参照の計算（前の値との差・累計）の値フィールド {pos} があるあいだは並べ替えできません"
                  "（Excel の制約＝「自動並べ替えをオフにしますか」の窓で止まる）。並べ替えを先にして show-as を後にするか、"
                  "show-as none で戻してから。"); return False
        if by:
            d = data_field(by)                         # 値の順（「売上の多い順」）は値フィールドを鍵に
            if d is None:
                print(f"エラー: 値フィールド '{by}' が見つかりません（pivot-field list で確認）。"); return False
            key = d.Name
        p.AutoSort(2 if order.startswith('d') else 1, key)
        print(f"並べ替え: {pname}[{field}] = {'降順' if order.startswith('d') else '昇順'}"
              + (f"（{key} の値で）" if by else "（ラベルで）"))
        print("（保存はしていません）"); return True

    if action == 'group-date':
        interval = (rest[3] if len(rest) >= 4 else 'months').lower()
        # Periods: [秒,分,時,日,月,四半期,年]
        flags = {'days': 3, 'months': 4, 'quarters': 5, 'years': 6}
        if interval not in flags:
            print("interval は days|months|quarters|years"); return False
        periods = [False] * 7
        periods[flags[interval]] = True
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Periods=periods)
        print(f"日付グループ化: {pname}[{field}] = {interval}"); print("（保存はしていません）"); return True

    if action == 'group-numeric':
        if len(rest) < 6:
            print("使い方: pivot-field group-numeric <pivot> <field> <start> <end> <interval>"); return False
        try:
            start, end, step = float(rest[3]), float(rest[4]), float(rest[5])
        except ValueError:
            # 非数値をそのまま渡すと ValueError のトレースバックが出るだけで、
            # 何が悪かったのか分からない。使い方を出して止める。
            print("使い方: pivot-field group-numeric <pivot> <field> <start> <end> <interval>")
            print(f"  start/end/interval は数値で指定してください（指定値: {rest[3]} / {rest[4]} / {rest[5]}）")
            return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Start=start, End=end, By=step)
        print(f"数値グループ化: {pname}[{field}] = {start}〜{end} 刻み{step}"); print("（保存はしていません）"); return True

    if action == 'show-as':
        kind = (rest[3] if len(rest) >= 4 else '').lower().replace('-', '_')
        if kind not in _XL_SHOW_AS:
            print("使い方: pivot-field show-as <pivot> <値フィールド> <種類> [--base-field 行フィールド] [--base-item 項目]")
            print("  種類: " + "|".join(_XL_SHOW_AS)); return False
        d = data_field(field)
        if d is None:
            print(f"エラー: 値フィールド '{field}' が見つかりません（pivot-field list で確認）。"); return False
        base_field = getattr(args, 'base_field', None)
        base_item = getattr(args, 'base_item', None)
        if kind in _SHOW_AS_BASE_FIELD and not base_field:
            # 基準フィールドは既定で先頭の行フィールド（前月比・累計は行の並びに沿う）
            try:
                if pt.RowFields.Count >= 1:
                    base_field = pt.RowFields.Item(1).Name
            except Exception:
                pass
            if not base_field:
                print("エラー: 基準にする行フィールドがありません（--base-field で指定）。"); return False
        if kind in ('running_total', 'percent_running_total') or                 (kind in _SHOW_AS_BASE_ITEM and (not base_item or base_item in _POSITION_ITEMS)):
            # 位置参照の計算は自動並べ替え・上位 N と同時に使えない（Excel が窓を出して止まる）。先に外して言う
            off = _turn_off_autosort_and_topn(pt)
            if off:
                print("  注意: 位置参照の計算（前の値との差・累計）は値の並べ替え・上位 N と同時に使えないため外しました: "
                      + "、".join(off))
        d.Calculation = _XL_SHOW_AS[kind]
        if kind in _SHOW_AS_BASE_FIELD:
            d.BaseField = base_field
        used_item = None
        if kind in _SHOW_AS_BASE_ITEM:
            # 基準項目は既定で「前の値」。Excel の表示言語で綴りが違うので順に試す
            for c in ([base_item] if base_item else ["(previous)", "(前の値)"]):
                try:
                    d.BaseItem = c; used_item = c; break
                except Exception:
                    continue
            if used_item is None:
                print(f"エラー: 基準項目を設定できません（--base-item で {base_field} の項目名か (previous) を指定）。"); return False
        print(f"計算の種類: {pname}[{d.Name}] = {kind}"
              + (f"  基準={base_field}" if kind in _SHOW_AS_BASE_FIELD else "")
              + (f" / {used_item}" if used_item else ""))
        print("（保存はしていません）"); return True

    if action == 'top-n':
        try:
            n = int(rest[3]) if len(rest) >= 4 else 0
        except ValueError:
            n = 0
        if n <= 0:
            print("使い方: pivot-field top-n <pivot> <field> <N> [--by 値フィールド] [--bottom]"); return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        by = getattr(args, 'by', None)
        d = data_field(by) if by else (pt.DataFields.Item(1) if pt.DataFields.Count >= 1 else None)
        if d is None:
            print("エラー: 基準にする値フィールドがありません（--by で指定。pivot-field list で確認）。"); return False
        bottom = bool(getattr(args, 'bottom', False))
        pos = _position_calc_fields(pt)
        if pos:
            print(f"エラー: 位置参照の計算（前の値との差・累計）の値フィールド {pos} があるあいだは上位/下位 N にできません"
                  "（Excel の制約＝窓で止まる）。top-n を先にして show-as を後にするか、show-as none で戻してから。"); return False
        try:
            p.ClearAllFilters()
            p.PivotFilters.Add2(2 if bottom else 1, d, n)       # xlTopCount=1 / xlBottomCount=2（31/32 は日付の種類＝E_INVALIDARG・2026-09-04 実射）
        except Exception as ex:
            print(f"エラー: 上位/下位の絞り込みに失敗しました: {ex}"); return False
        print(f"{'下位' if bottom else '上位'} {n}: {pname}[{field}] を {d.Name} で絞り込み"); print("（保存はしていません）"); return True

    if action == 'filter-clear':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.ClearAllFilters()
        print(f"絞り込み解除: {pname}[{field}]"); print("（保存はしていません）"); return True

    if action == 'position':
        try:
            n = int(rest[3]) if len(rest) >= 4 else 0
        except ValueError:
            n = 0
        if n <= 0:
            print("使い方: pivot-field position <pivot> <field> <N>（同じ区画の中の順番。1 が先頭）"); return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Position = n
        print(f"並び順: {pname}[{field}] = {n} 番目"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


@protect_safe
@dialog_safe
def cmd_pivot_calc(args):
    """ピボットの計算フィールド・レイアウト: pivot-calc <action> <pivot名> ...

      get-data <pivot>                                出力範囲の値を表示
      calc-field create <pivot> <名前> <数式>         計算フィールド作成（=Revenue-Cost 等）
      calc-field list <pivot>
      calc-field delete <pivot> <名前>
      layout <pivot> <compact|tabular|outline>        レポートレイアウト
      subtotals <pivot> <field> <on|off>              小計の表示
      grand-totals <pivot> <rows|cols|both> <on|off>  総計の表示
      refresh <pivot>                                 更新（元データが変わったとき）
      style <pivot> <スタイル名>                      PivotStyleMedium9 / PivotStyleLight16 など
      repeat-labels <pivot> <on|off>                  行ラベルを全行に繰り返す（tabular と組む）
      empty-as <pivot> <文字>                         空白セルの表示（"0" など）
      set-source <pivot> <データ範囲>                 元データの範囲を変える（行が増えたとき）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-calc <action> <pivot名> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'calc-field':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        pname = rest[2] if len(rest) >= 3 else None
        sh, pt = _find_pivot(wb, pname) if pname else (None, None)
        if not pt:
            print(f"エラー: ピボット '{pname}' が見つかりません。"); return False
        if sub == 'create':
            if len(rest) < 5:
                print("使い方: pivot-calc calc-field create <pivot> <名前> <数式>"); return False
            cf_name, formula = rest[3], rest[4]
            pt.CalculatedFields().Add(cf_name, formula)
            print(f"計算フィールド作成: {pname}[{cf_name}] = {formula}")
            print("  （値に表示するには pivot-field add-value で追加）")
            print("（保存はしていません）"); return True
        if sub == 'list':
            cfs = pt.CalculatedFields()
            print(f"--- {pname} の計算フィールド ({cfs.Count}) ---")
            for i in range(1, cfs.Count + 1):
                f = cfs.Item(i)
                try:
                    formula = f.Formula
                except Exception:
                    formula = ''
                print(f"  {f.Name} = {formula}")
            if cfs.Count == 0:
                print("  (なし)")
            return True
        if sub == 'delete':
            cf_name = rest[3] if len(rest) >= 4 else None
            if not cf_name:
                print("使い方: pivot-calc calc-field delete <pivot> <計算フィールド名>")
                return False
            try:
                pt.PivotFields(cf_name).Delete()
            except Exception:
                names = []
                try:
                    cfs = pt.CalculatedFields()
                    names = [cfs.Item(i).Name for i in range(1, cfs.Count + 1)]
                except Exception:
                    pass
                print(f"エラー: 計算フィールド '{cf_name}' を削除できません。")
                print("  存在する計算フィールド: " + (", ".join(names) or "(なし)"))
                return False
            print(f"計算フィールド削除: {pname}[{cf_name}]"); print("（保存はしていません）"); return True
        print("使い方: pivot-calc calc-field <create|list|delete> <pivot> ...")
        return False

    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません。"); return False

    if action == 'get-data':
        rng = pt.TableRange2
        print(f"ピボット出力範囲: {pt.Parent.Name}!{rng.Address}")
        data = rng.Value
        if data is not None and not isinstance(data, tuple):
            data = ((data,),)      # 1 セルだけ（壊れたピボット等）は素の値で返る
        if data is not None:
            for row in data:
                cells = [('' if c is None else str(c)) for c in (row if isinstance(row, tuple) else [row])]
                print("  " + " | ".join(cells))
        return True

    if action == 'layout':
        lay = (rest[2] if len(rest) >= 3 else 'compact').lower()
        laymap = {'compact': 0, 'tabular': 1, 'outline': 2}   # xlCompactRow/xlTabularRow/xlOutlineRow
        if lay not in laymap:
            print("layout は compact|tabular|outline"); return False
        pt.RowAxisLayout(laymap[lay])
        print(f"レイアウト: {pname} = {lay}"); print("（保存はしていません）"); return True

    if action == 'subtotals':
        field = rest[2] if len(rest) >= 3 else None
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        try:
            p = pt.PivotFields(field)
        except Exception:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Subtotals = tuple([onoff == 'on'] + [False] * 11)   # 先頭=自動小計
        print(f"小計: {pname}[{field}] = {onoff}"); print("（保存はしていません）"); return True

    if action == 'grand-totals':
        which = (rest[2] if len(rest) >= 3 else 'both').lower()
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        val = (onoff == 'on')
        if which in ('rows', 'both'):
            pt.RowGrand = val
        if which in ('cols', 'both'):
            pt.ColumnGrand = val
        print(f"総計: {pname} {which} = {onoff}"); print("（保存はしていません）"); return True

    if action == 'refresh':
        pt.RefreshTable()
        print(f"更新: {pname}  出力={pt.Parent.Name}!{str(pt.TableRange2.Address).replace('$', '')}")
        print("（保存はしていません）"); return True

    if action == 'style':
        style = rest[2] if len(rest) >= 3 else None
        if not style:
            print("使い方: pivot-calc style <pivot> <スタイル名（PivotStyleMedium9 / PivotStyleLight16 など）>"); return False
        try:
            pt.TableStyle2 = style
        except Exception as ex:
            print(f"エラー: スタイル '{style}' を設定できません: {ex}"); return False
        print(f"スタイル: {pname} = {style}"); print("（保存はしていません）"); return True

    if action == 'repeat-labels':
        onoff = (rest[2] if len(rest) >= 3 else 'on').lower()
        pt.RepeatAllLabels(2 if onoff == 'on' else 1)      # xlRepeatLabels=2 / xlDoNotRepeatLabels=1
        print(f"行ラベルの繰り返し: {pname} = {onoff}"); print("（保存はしていません）"); return True

    if action == 'empty-as':
        text = rest[2] if len(rest) >= 3 else ''
        pt.NullString = text
        pt.DisplayNullString = True
        print(f"空白セルの表示: {pname} = {text!r}"); print("（保存はしていません）"); return True

    if action == 'set-source':
        spec = rest[2] if len(rest) >= 3 else None
        if not spec:
            print("使い方: pivot-calc set-source <pivot> <データ範囲（Sheet1!A1:F100 など）>"); return False
        ws_s, rng = _resolve_range(xl, wb, spec)
        pc = wb.PivotCaches().Create(1, rng)
        pt.ChangePivotCache(pc)
        pt.RefreshTable()
        print(f"元データ変更: {pname} ← {ws_s.Name}!{rng.Address}"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (3) スライサー
# ================================================================

def _find_pivot_or_table(wb, name):
    """名前からピボット or テーブル(ListObject)を探す。戻り値 (obj, kind, sheet) """
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return pt, 'pivot', sh
    for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
        for lo in sh.ListObjects:
            if lo.Name == name:
                return lo, 'table', sh
    return None, None, None


@protect_safe
@dialog_safe
def cmd_slicer(args):
    """スライサー操作: slicer <add|list|delete> ...

      slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]
      slicer list
      slicer delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: slicer <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                cnt += 1
                # Slicer.Parent は SlicerCache を返す実装があるため、シート名は Shape 経由で取る
                try:
                    sheet_name = sl.Shape.Parent.Name
                except Exception:
                    try:
                        sheet_name = sl.Parent.Name
                    except Exception:
                        sheet_name = '?'
                print(f"{sl.Name}  (フィールド={sc.SourceName}, シート={sheet_name})")
        if cnt == 0:
            print("スライサーはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: slicer delete <name>"); return False
        name = rest[1]
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                if sl.Name == name:
                    sl.Delete()
                    print(f"スライサー削除: {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: スライサー '{name}' が見つかりません")
        return False

    if action == 'add':
        if len(rest) < 3:
            print("使い方: slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]")
            return False
        src_name, field = rest[1], rest[2]
        src, kind, sh = _find_pivot_or_table(wb, src_name)
        if src is None:
            print(f"エラー: ピボット/テーブル '{src_name}' が見つかりません")
            return False

        # 配置先の座標を先に解決する（Add2 の後に --at が不正で例外になると、
        # スライサー本体のない SlicerCache だけが孤児としてブックに残るため）
        at = getattr(args, 'at', None)
        dws = sh
        if at:
            anchor = dws.Range(at)
            top, left = anchor.Top, anchor.Left
        else:
            top, left = 10.0, 400.0

        sc = wb.SlicerCaches.Add2(src, field)
        try:
            sl = sc.Slicers.Add(SlicerDestination=dws, Caption=field,
                                Top=top, Left=left, Width=144.0, Height=180.0)
        except Exception as ex:
            # Add2 は通ったのに Slicers.Add が失敗すると（保護シート・座標不正等）、
            # スライサー本体のない SlicerCache だけが孤児として残り、以後
            # slicer list にも出ないまま再実行のたびに増える。作ったものは戻す。
            print(f"エラー: スライサーの配置に失敗しました: {ex}")
            try:
                sc.Delete()
                print("  作成途中のスライサーキャッシュは片づけました。")
            except Exception:
                print("  ⚠ スライサーキャッシュの後始末に失敗しました（保存前ならブックを閉じ直すのが確実です）。")
            return False
        # Slicers.Add の Name 引数は効かないことがあるので作成後に明示セット
        req_name = getattr(args, 'name', None)
        if req_name:
            eff_name = req_name.replace(' ', '')
            if eff_name != req_name:
                print(f"⚠ スライサー名のスペースは使えないため除去しました: '{req_name}' → '{eff_name}'")
            try:
                sl.Name = eff_name
            except Exception as ex:
                print(f"⚠ 名前 '{eff_name}' を設定できませんでした（{ex}）。自動名のままです。")
        print(f"スライサー追加: {sl.Name}  ソース={src_name}({kind})  フィールド={field}  シート={dws.Name}")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（add|list|delete）")
    return False


# ================================================================
# 計算モード (大量書き込みの高速化)
# ================================================================

def cmd_calc_mode(args):
    """計算モードの確認・切替・再計算

      calc-mode                 現在のモードを表示
      calc-mode manual          手動計算に（大量書込の前に）
      calc-mode auto            自動計算に戻す
      calc-mode recalc          今すぐ再計算（手動中の一括計算）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    names = {-4105: '自動 (automatic)', -4135: '手動 (manual)', 2: '半自動 (semiautomatic)'}

    if not rest:
        m = xl.Calculation
        print(f"現在の計算モード: {names.get(m, m)}")
        return True

    sub = rest[0].lower()
    if sub in ('auto', 'automatic'):
        xl.Calculation = -4105
        print("計算モード → 自動")
    elif sub == 'manual':
        xl.Calculation = -4135
        print("計算モード → 手動（書込後は calc-mode recalc / auto で再計算）")
    elif sub in ('recalc', 'now', 'calculate'):
        xl.Calculate()
        print("再計算しました")
    else:
        print(f"未知の指定: {sub}（manual|auto|recalc）")
        return False
    return True


# ================================================================
# 重量級コマンド (4) PowerQuery （一覧・更新・作成・M式書換・削除・読み込み配線）
# ================================================================

def _refresh_connection_sync(cn):
    """接続を「完了まで待って」更新する。

    WorkbookConnection.Refresh は BackgroundQuery=True（Excel が作る Query 接続では
    しばしば既定で True）だと即座に戻り、更新はバックグラウンドで走る。そのまま
    「更新しました」と報告すると、まだ古いデータのままなのに成功と言うことになり、
    直後の read-range が旧データを読む。更新中のエラーも戻り値・例外に出ない。
    そこで一時的に BackgroundQuery=False にして同期実行させ、元に戻す。
    """
    ole = None
    prev = None
    try:
        ole = cn.OLEDBConnection
        prev = bool(ole.BackgroundQuery)
        if prev:
            ole.BackgroundQuery = False
    except Exception:
        ole = None                      # OLEDB でない接続（テキスト等）はそのまま撃つ
    try:
        cn.Refresh()
    finally:
        if ole is not None and prev:
            try:
                ole.BackgroundQuery = prev
            except Exception:
                pass


def _connection_used_by_table(wb, cn_name):
    """接続 cn_name を使っているシートのテーブルがあれば "シート名!テーブル名" を返す。

    この接続を消すとそのテーブルの更新配線が切れる（更新できないテーブルが
    シートに残る）。削除する側は撃つ前に必ずここを通す。
    """
    try:
        for ws_chk in wb.Worksheets:
            for lo_chk in ws_chk.ListObjects:
                try:
                    if lo_chk.QueryTable.WorkbookConnection.Name == cn_name:
                        return f"{ws_chk.Name}!{lo_chk.Name}"
                except Exception:
                    continue     # QueryTable を持たない普通のテーブルはここに来る
    except Exception:
        pass
    return None


def _snapshot_connection(cn):
    """接続の再作成に要る情報を控える（削除をロールバックするため）。

    取れなかった項目は None。最低限 Name と接続文字列が取れなければ復元不能なので
    None を返し、呼び出し側は「消したまま失敗した」ことを明示する。
    """
    snap = {'name': None, 'desc': '', 'conn': None, 'cmd_text': None,
            'cmd_type': None, 'in_model': None}
    try:
        snap['name'] = cn.Name
    except Exception:
        return None
    try:
        snap['desc'] = cn.Description or ''
    except Exception:
        pass                              # 説明は欠けても接続の実体は変わらない
    # 以下は復元の同一性に効く項目。1つでも取れなければ「復元できる」と言ってはいけない。
    # まとめて try に入れると、接続文字列だけ取れて CommandText/CommandType が
    # 欠けたまま初期値で復元し、別物の接続を作って「元に戻しました」と報告してしまう
    # （PowerQuery のモデル接続は CommandType=6。既定の 1 で復元すると配線が変わる）
    try:
        snap['in_model'] = bool(cn.InModel)
    except Exception:
        return None
    try:
        sub = cn.OLEDBConnection            # PowerQuery 接続は OLEDB
    except Exception:
        return None
    for key, get in (('conn', lambda: str(sub.Connection)),
                     ('cmd_text', lambda: sub.CommandText),
                     ('cmd_type', lambda: int(sub.CommandType))):
        try:
            snap[key] = get()
        except Exception:
            return None
    if not snap['conn']:
        return None
    return snap


def _restore_connection(wb, snap):
    """_snapshot_connection で控えた接続を作り直す。成功したら True"""
    try:
        wb.Connections.Add2(snap['name'], snap['desc'], snap['conn'],
                            snap['cmd_text'], snap['cmd_type'],
                            snap['in_model'], False)
        return True
    except Exception:
        return False


_M_SOURCE_RE = re.compile(r'Excel\.CurrentWorkbook\(\)\s*\{\s*\[\s*Name\s*=\s*"([^"]+)"\s*\]\s*\}\s*\[Content\]')
_MODEL_DATATYPE = {130: '文字', 20: '整数', 5: '小数', 14: '10進', 6: '通貨', 7: '日付', 11: '真偽', 3: '整数'}


def _m_source_table(formula):
    """M 式が Excel.CurrentWorkbook(){[Name="T"]}[Content] で読んでいるテーブル名（無ければ None）"""
    m = _M_SOURCE_RE.search(formula or '')
    return m.group(1) if m else None


def _m_with_types(formula, coltypes):
    """型付けの無い M 式を、Table.TransformColumnTypes で包んだ M 式にする（coltypes: [(列名, M の型), …]）。
    PQ でモデルに載せた列は型を付けないと全部「文字」になり、SUM のメジャーが無言で消える（2026-09-04 実射）"""
    if not coltypes or 'Table.TransformColumnTypes' in (formula or ''):
        return formula
    body = (formula or '').strip()
    pairs = ", ".join('{"%s", %s}' % (c.replace('"', '""'), t) for c, t in coltypes)
    return ("let __元 = (" + body + "),\n    __型付け = Table.TransformColumnTypes(__元, {" + pairs + "})\nin __型付け")


def _listobject_column_types(wb, table_name):
    """テーブル(ListObject)の列ごとに M の型を決める: 全部整数→Int64.Type／数値→type number／日付→type date／他は type text"""
    for sh in wb.Worksheets:
        try:
            lo = sh.ListObjects(table_name)
        except Exception:
            continue
        out = []
        try:
            body = lo.DataBodyRange.Value if lo.DataBodyRange is not None else None
        except Exception:
            body = None
        rows = [] if body is None else ([list(r) if isinstance(r, tuple) else [r] for r in body] if isinstance(body, tuple) else [[body]])
        for j, col in enumerate(lo.ListColumns):
            vals = [r[j] for r in rows if j < len(r) and r[j] not in (None, '')]
            if vals and all(hasattr(v, 'year') and not isinstance(v, str) for v in vals):
                t = 'type date'
            elif vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
                t = 'Int64.Type' if all(float(v).is_integer() for v in vals) else 'type number'
            else:
                t = 'type text'
            out.append((str(col.Name), t))
        return out
    return []


def _auto_typed_m(wb, formula):
    """型付けの無い M（ブックのテーブルを読むだけ）に、テーブルの実物から決めた型を足す。足したら (M, 説明) を返す"""
    if 'Table.TransformColumnTypes' in (formula or ''):
        return formula, ''
    src = _m_source_table(formula)
    if not src:
        return formula, ''
    coltypes = _listobject_column_types(wb, src)
    if not coltypes:
        return formula, ''
    typed = _m_with_types(formula, coltypes)
    note = "  型付けを足しました（付けないとモデルの列が全部「文字」になり SUM のメジャーが無言で消える）: " + \
           ", ".join(f"{c}={t.replace('type ', '').replace('.Type', '')}" for c, t in coltypes)
    return typed, note


@protect_safe
@dialog_safe
def cmd_powerquery(args):
    """PowerQuery: powerquery <list|refresh|add|edit|delete|load> ...

      powerquery list                 クエリと接続の一覧（M式の行数つき）
      powerquery refresh              全クエリ/接続を更新 (RefreshAll)
      powerquery refresh <name>       指定クエリ/接続を更新
      powerquery add <name>           M式から新規クエリ作成（接続のみ）
                                      M式は --m / --m-file / _last_query.m
      powerquery edit <name>          既存クエリのM式を書き換え（M式の指定は add と同じ）
      powerquery delete <name>        クエリを削除
      powerquery load <name> --to sheet [--sheet S] [--at A1]   シートのテーブルに読み込み
      powerquery load <name> --to model                          データモデルに読み込み
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: powerquery <list|refresh [name]|add <name>|delete <name>>")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        # クエリ一覧
        try:
            qs = wb.Queries
            qcount = qs.Count
        except Exception:
            qs = None
            qcount = 0
        if qs and qcount > 0:
            print(f"--- PowerQuery クエリ ({qcount}) ---")
            for i in range(1, qcount + 1):
                q = qs.Item(i)
                try:
                    desc = q.Description or ''
                except Exception:
                    desc = ''
                # M式の行数を補助表示
                try:
                    nlines = len(str(q.Formula).replace('\r\n', '\n').split('\n'))
                except Exception:
                    nlines = '?'
                print(f"  {q.Name}  (M式 {nlines}行)" + (f"  - {desc}" if desc else ""))
        else:
            print("PowerQuery クエリはありません。")
        # 接続一覧（更新対象の確認用）
        try:
            conns = wb.Connections
            ccount = conns.Count
        except Exception:
            ccount = 0
        if ccount > 0:
            print(f"--- 接続 ({ccount}) ---")
            for cn in conns:
                print(f"  {cn.Name}")
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            target_conn = None
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    target_conn = cn
                    break
            if target_conn:
                # 完了を待って更新する（待たないと「更新しました」と言った時点では
                # まだ古いデータのまま＝誤った成功報告になる）
                _refresh_connection_sync(target_conn)
                print(f"更新しました: {target_conn.Name}")
                return True
            # 接続が無い（読み込みなしクエリ等）
            print(f"接続 '{name}' が見つかりません。")
            print("  （読み込みなしクエリは更新対象がありません。powerquery list で名前を確認）")
            return False
        else:
            wb.RefreshAll()
            print("全クエリ/接続を更新しました (RefreshAll)")
            print("  ※ バックグラウンド更新の場合、完了まで数秒かかることがあります。")
            return True

    if action == 'add':
        if len(rest) < 2:
            print("使い方: powerquery add <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得: --m インライン > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                print("  _last_query.m にM式を書くか、--m-file / --m を指定してください。")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        # 重複チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name in existing:
            print(f"エラー: クエリ '{name}' は既に存在します（delete してから add）。")
            return False
        desc = getattr(args, 'desc', None) or ''
        formula, typed_note = _auto_typed_m(wb, formula)
        wb.Queries.Add(name, formula, desc)
        print(f"クエリ作成: {name}（接続のみ。シート/モデルへの読み込みは別途）")
        if typed_note:
            print(typed_note)
        print("（保存はしていません）")
        return True

    if action == 'edit':
        if len(rest) < 2:
            print("使い方: powerquery edit <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得（add と同じ）: --m > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            q = wb.Queries.Item(i)
            if q.Name == name:
                formula, typed_note = _auto_typed_m(wb, formula)
                q.Formula = formula                # WorkbookQuery.Formula は書込可（検証済）
                print(f"クエリ書き換え: {name}")
                if typed_note:
                    print(typed_note)
                print("（保存はしていません。反映には powerquery refresh が必要なことがあります）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: powerquery delete <name>"); return False
        name = rest[1]
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            if wb.Queries.Item(i).Name == name:
                wb.Queries.Item(i).Delete()
                print(f"クエリ削除: {name}")
                print("（保存はしていません）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'load':
        # 接続のみクエリを「シートのテーブル」または「データモデル」に読み込む配線。
        #   powerquery load <name> --to sheet  [--sheet S] [--at A1]
        #   powerquery load <name> --to model
        if len(rest) < 2:
            print('使い方: powerquery load <name> --to sheet|model [--sheet S] [--at A1]')
            return False
        name = rest[1]
        to = (getattr(args, 'to', None) or 'sheet').lower()
        # クエリ存在チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name not in existing:
            print(f"エラー: クエリ '{name}' が見つかりません（powerquery list で確認）。")
            return False

        # Power Query (Mashup) の OLEDB 接続文字列 — 記録マクロが生成する形に合わせる
        conn_str = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
                    f'Location={name};Extended Properties=""')
        cmd_text = f"SELECT * FROM [{name}]"

        if to == 'sheet':
            # 出力先シート（--sheet 省略時はアクティブシート）
            sheet_name = getattr(args, 'sheet', None)
            ws = None
            created_ws = None
            if sheet_name:
                for i in range(1, wb.Worksheets.Count + 1):
                    # Excel のシート名重複判定は大文字小文字を区別しない
                    if wb.Worksheets.Item(i).Name.lower() == sheet_name.lower():
                        ws = wb.Worksheets.Item(i); break
                if ws is None:
                    ws = wb.Worksheets.Add()
                    created_ws = ws
                    try:
                        ws.Name = sheet_name
                    except Exception as ex:
                        # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                        print(f"エラー: シート名 '{sheet_name}' を設定できません: {ex}")
                        xl.DisplayAlerts = False
                        try:
                            ws.Delete()
                        finally:
                            xl.DisplayAlerts = True
                        return False
            else:
                ws = wb.ActiveSheet
            at = getattr(args, 'at', None) or 'A1'
            dest = ws.Range(at)
            # 0 = xlSrcExternal。Source に Mashup の OLEDB 文字列を渡す
            lo = ws.ListObjects.Add(0, conn_str, None, True, dest)
            qt = lo.QueryTable
            qt.CommandType = 2                 # xlCmdSql
            qt.CommandText = cmd_text
            qt.RowNumbers = False
            qt.FillAdjacentFormulas = False
            qt.PreserveFormatting = True
            qt.RefreshOnFileOpen = False
            qt.BackgroundQuery = False
            qt.AdjustColumnWidth = True
            try:
                # QueryTable.Refresh は Boolean を返す。失敗しても例外を投げず
                # False を返す経路があり、捨てると「シートに読み込みました」と
                # 成功報告したまま空／エラー列だけのテーブルが残る
                if qt.Refresh(False) is False:  # BackgroundQuery:=False
                    raise RuntimeError("QueryTable.Refresh が False を返しました"
                                       "（M式の実行時エラーの可能性）")
            except Exception as ex:
                # M式の実行時エラー等で失敗すると、追加済みの ListObject と
                # 自動生成の接続が孤児として残り、再実行のたび「テーブル1/2…」と
                # 増殖する。このコマンドが作ったものだけ片づける
                print(f"エラー: クエリの読み込みに失敗しました: {ex}")
                print("  M式の実行時エラーの可能性があります"
                      "（powerquery list でクエリを確認、powerquery edit で書き換え）。")
                try:
                    wbconn = qt.WorkbookConnection
                except Exception:
                    wbconn = None
                try:
                    lo.Delete()
                except Exception:
                    pass
                try:
                    if wbconn is not None:
                        wbconn.Delete()
                except Exception:
                    pass
                if created_ws is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_ws.Delete()
                    except Exception:
                        pass
                    finally:
                        xl.DisplayAlerts = True
                print("  追加途中のテーブル・接続は片づけました。")
                return False
            # リネームの失敗を握りつぶすと、成功メッセージだけ出て
            # 以後 powerquery refresh <name> で引けない（名前が汎用名のまま）状態になる。
            # 読み込み自体は成功しているので中止はせず、警告として明示する。
            try:
                lo.Name = name
            except Exception as ex:
                print(f"⚠ テーブル名を '{name}' にできませんでした（{ex}）。自動名 '{lo.Name}' のままです。")
            # 既定では「接続」等の汎用名が付く。refresh <name> で引けるよう
            # Excel 標準の "Query - <name>" に揃える。
            conn_name = None
            try:
                wbconn = qt.WorkbookConnection
                if wbconn is not None:
                    wbconn.Name = f"Query - {name}"
                    conn_name = wbconn.Name
            except Exception as ex:
                print(f"⚠ 接続名を 'Query - {name}' にできませんでした（{ex}）。")
                print(f"   powerquery refresh {name} では引けません。"
                      "connection list で実際の接続名を確認し、その名前で refresh してください。")
            print(f"シートに読み込みました: {name} → {ws.Name}!{at}（テーブル: {lo.Name}"
                  + (f", 接続: {conn_name}" if conn_name else "") + "）")
            print("（保存はしていません）")
            return True

        if to == 'model':
            # データモデル（Power Pivot）へ。Queries.Add が作る "Query - name"
            # 接続が残っていると衝突するので、あれば作り直す。
            # ただしその接続がシートのテーブル（--to sheet の読み込み）に使われている
            # 場合、削除するとシート側の更新配線が壊れるため停止する。
            cn_name = f"Query - {name}"
            deleted_snap = None       # 消した既存接続の控え（Add2 失敗時のロールバック用）
            deleted_any = False       # 控えの有無と別に「実際に消したか」を持つ
            delete_failed = None      # 既存接続の Delete が失敗した事実（後段の誤診断防止）
            for cn in list(wb.Connections):
                if cn.Name == cn_name:
                    used_by = _connection_used_by_table(wb, cn_name)
                    if used_by:
                        print(f"エラー: 接続 '{cn_name}' はシートのテーブル {used_by} が使用中です。")
                        print("  削除するとテーブルの更新ができなくなるため中止しました。")
                        print("  モデルにも読み込みたい場合は、シート読み込みを解除してから実行してください。")
                        return False
                    # _connection_used_by_table はシートのテーブルしか見ない。
                    # 既にモデルに載っている接続（InModel）は「未使用」と判定されて
                    # そのまま Delete され、そのテーブルに紐づくメジャー・
                    # リレーションシップが道連れになる。しかも Add2 で作り直すので
                    # 「データモデルに読み込みました」と成功報告してしまう。
                    in_model = False
                    try:
                        in_model = bool(cn.InModel)
                    except Exception:
                        in_model = False
                    if in_model and not getattr(args, 'force', False):
                        print(f"エラー: '{cn_name}' は既にデータモデルに読み込まれています。")
                        print("  作り直すと、このテーブルに紐づくメジャーとリレーションシップが")
                        print("  失われるため中止しました（更新するだけなら powerquery refresh）。")
                        print("  承知のうえで作り直すなら --force を付けてください。")
                        return False
                    # 消す前に中身を控える。Add2 が失敗したときに元へ戻せないと、
                    # 「接続だけ消えてモデルにも載っていない＝クエリが未配線」の
                    # 一番たちの悪い状態で終わるため。
                    old_snap = _snapshot_connection(cn)
                    try:
                        cn.Delete()
                        deleted_any = True
                        deleted_snap = old_snap
                    except Exception as ex_del:
                        # 消せなかった事実は控える。ここで黙ると、後段の Add2 が
                        # 同名衝突で失敗したときに「M式のエラーかも」と誤診断する
                        delete_failed = str(ex_del)
            # Connections.Add2(Name, Description, ConnectionString, CommandText,
            #                  lCmdtype, CreateModelConnection, ImportRelationships)
            # モデル読込は記録マクロ形式に合わせる: CommandText=クエリ名,
            # lCmdtype=6 (xlCmdTableCollection)。これでモデルテーブル名が
            # クエリ名になる（SQL/SELECT形式だと "クエリ" の汎用名になる）。
            try:
                wb.Connections.Add2(cn_name, "", conn_str, name, 6, True, False)
            except Exception as ex:
                # M式の実行時エラー・モデル非対応等で失敗しうる。丸腰で呼ぶと
                # 上で消した既存接続が戻らない。控えた内容で作り直す。
                print(f"エラー: データモデルへの読み込みに失敗しました: {ex}")
                if delete_failed:
                    # 既存の同名接続が消せていない＝Add2 失敗の真因はまず名前衝突。
                    # ここで「M式のエラーかも」と言うと誤誘導になる
                    print(f"  既存の接続 '{cn_name}' を削除できていません（{delete_failed}）。")
                    print("  同名の接続が残っているため作り直せなかった可能性が高いです。")
                else:
                    print("  M式の実行時エラー、またはこのブックがデータモデル非対応の可能性があります。")
                if deleted_snap is None:
                    if deleted_any:
                        # 「何も消していない」と「消したが控えが取れず戻せない」は別物。
                        # 黙って帰ると、接続が消えた事実が一切報告されない
                        print(f"  ⚠ 既存の接続 '{cn_name}' は削除済みで、控えが取れなかったため復旧できません。")
                        print("     クエリ自体は残っています。配線をやり直すには:")
                        print(f"       powerquery load {name} --to sheet   （シートに読み込む場合）")
                        print(f"       powerquery load {name} --to model   （モデルに読み込む場合）")
                    return False
                if _restore_connection(wb, deleted_snap):
                    print(f"  既存の接続 '{cn_name}' は元に戻しました（クエリの配線は元のままです）。")
                else:
                    print(f"  ⚠ 既存の接続 '{cn_name}' を消したまま復旧できませんでした。")
                    print("     クエリ自体は残っています。配線をやり直すには:")
                    print(f"       powerquery load {name} --to sheet   （シートに読み込む場合）")
                    print(f"       powerquery load {name} --to model   （モデルに読み込む場合）")
                    print("     現状は connection list / datamodel list で確認できます。")
                return False
            print(f"データモデルに読み込みました: {name}")
            print("（保存はしていません。datamodel list で確認できます）")
            return True

        print(f"未知の読み込み先: {to}（sheet|model）")
        return False

    print(f"未知のアクション: {action}（list|refresh|add|edit|delete|load）")
    return False


# ================================================================
# 重量級コマンド (5) コネクション / データモデル （管理・読み取り）
# ================================================================

# XlConnectionType: xlConnectionTypeOLEDB=1, ODBC=2, XMLMAP=3, TEXT=4, WEB=5,
#                   DATAFEED=6, MODEL=7, WORKSHEET=8, NOSOURCE=9
_XL_CONN_TYPE = {1: 'OLEDB', 2: 'ODBC', 3: 'XMLMAP', 4: 'TEXT',
                 5: 'WEB', 6: 'DATAFEED', 7: 'MODEL', 8: 'WORKSHEET', 9: 'NOSOURCE'}


@protect_safe
@dialog_safe
def cmd_connection(args):
    """ブック接続の管理: connection <list|refresh|delete> [name]

      connection list                クエリ/外部データ接続の一覧（種別・接続文字列）
      connection refresh [name]      接続を更新（name 省略で全件 RefreshAll）
      connection delete <name>       接続を削除（使用中のテーブルがあれば注記を出す）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: connection <list|refresh|delete> [name]")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        conns = wb.Connections
        n = conns.Count
        if n == 0:
            print("接続はありません。")
            return True
        print(f"--- ブック接続 ({n}) ---")
        for cn in conns:
            try:
                t = _XL_CONN_TYPE.get(int(cn.Type), cn.Type)
            except Exception:
                t = '?'
            print(f"  {cn.Name}  [{t}]")
            try:
                if cn.Description:
                    print(f"      説明: {cn.Description}")
            except Exception:
                pass
            # 接続文字列・コマンド（OLEDB/ODBC）
            try:
                sub = None
                if int(cn.Type) == 1:
                    sub = cn.OLEDBConnection
                elif int(cn.Type) == 2:
                    sub = cn.ODBCConnection
                if sub is not None:
                    cs = str(sub.Connection)
                    print(f"      接続: {cs[:100]}{'…' if len(cs) > 100 else ''}")
            except Exception:
                pass
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    cn.Refresh()
                    print(f"更新しました: {cn.Name}")
                    return True
            print(f"エラー: 接続 '{name}' が見つかりません")
            return False
        wb.RefreshAll()
        print("全接続を更新しました (RefreshAll)")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: connection delete <name>"); return False
        name = rest[1]
        for cn in wb.Connections:
            if cn.Name == name or cn.Name == f"Query - {name}":
                actual = cn.Name           # Delete 後は参照不可になるので退避
                # 名指しされた接続はそのまま消す（シートのテーブルが使っていても、
                # シートのデータ自体は残る＝更新の配線が外れるだけ。使用中の注記は
                # 出すが、消すか消さないかの判断はツールの仕事ではない）
                used_by = _connection_used_by_table(wb, actual)
                cn.Delete()
                print(f"接続を削除: {actual}")
                if used_by:
                    print(f"  注記: シートのテーブル {used_by} がこの接続を使っていました。"
                          "以後そのテーブルは更新できません（データは残っています）。")
                print("（保存はしていません）")
                return True
        print(f"エラー: 接続 '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（list|refresh|delete）")
    return False


@protect_safe
def cmd_datamodel(args):
    """データモデル: datamodel <list|relation|measure>

      datamodel list   モデルのテーブル・リレーションシップ・メジャーを一覧
      datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション作成
      datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション削除
      datamodel measure add <テーブル> <メジャー名> --dax "式" [--format general|whole|decimal|currency|percent|scientific]
                                                                 [--decimals N] [--thousands] [--symbol JPY]   メジャー(DAX)作成
      datamodel measure delete <メジャー名>                              メジャー削除
      （※ テーブルの追加は powerquery load --to model）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    action = rest[0].lower() if rest else 'list'
    xl, wb = get_workbook(target_file)

    try:
        model = wb.Model
    except Exception:
        print("このブックはデータモデルに対応していません。")
        # list 系は「無い」を正常報告でよいが、追加/削除の要求は実行されて
        # いないので失敗として返す（batch やスクリプト連携で握りつぶさない）。
        # 'measures add' のような複数形＋動詞も add/delete は変更要求なので失敗側
        _sub = rest[1].lower() if len(rest) >= 2 else ''
        return (action in ('list', 'tables', 'measures', 'relations')
                and _sub not in ('add', 'delete'))

    # --- リレーションシップの作成・削除 ---
    #   datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>
    #   datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>
    if action in ('relation', 'rel', 'relationship'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub in ('add', 'delete'):
            if len(rest) < 6:
                print(f'使い方: datamodel relation {sub} <FKテーブル> <FK列> <PKテーブル> <PK列>')
                print('  FK=多側(参照する側) / PK=一側(参照される側)')
                return False
            fkt_name, fkc_name, pkt_name, pkc_name = rest[2], rest[3], rest[4], rest[5]
        if sub == 'add':
            try:
                fkt = model.ModelTables.Item(fkt_name)
                pkt = model.ModelTables.Item(pkt_name)
            except Exception:
                print(f"エラー: テーブルが見つかりません（{fkt_name} / {pkt_name}）。datamodel list で確認。")
                return False
            try:
                fkc = fkt.ModelTableColumns.Item(fkc_name)
                pkc = pkt.ModelTableColumns.Item(pkc_name)
            except Exception:
                print(f"エラー: 列が見つかりません（{fkt_name}[{fkc_name}] / {pkt_name}[{pkc_name}]）。")
                return False
            model.ModelRelationships.Add(fkc, pkc)
            print(f"リレーション作成: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            rels = model.ModelRelationships
            for i in range(1, rels.Count + 1):
                r = rels.Item(i)
                try:
                    if (r.ForeignKeyTable.Name == fkt_name and r.ForeignKeyColumn.Name == fkc_name
                            and r.PrimaryKeyTable.Name == pkt_name and r.PrimaryKeyColumn.Name == pkc_name):
                        r.Delete()
                        print(f"リレーション削除: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
                        print("（保存はしていません）")
                        return True
                except Exception:
                    continue
            print("エラー: 該当するリレーションが見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel relation <add|delete> <FKテーブル> <FK列> <PKテーブル> <PK列>')
        return False

    # --- メジャー(DAX)の作成・削除 ---
    #   datamodel measure add <テーブル> <メジャー名> [--dax "式" | --dax-file f | _last_dax.dax]
    #   datamodel measure delete <メジャー名>
    if action in ('measure', 'measures'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub == 'add':
            if len(rest) < 4:
                print('使い方: datamodel measure add <テーブル> <メジャー名> --dax "DAX式"')
                print('  DAX は --dax / --dax-file / _last_dax.dax(UTF-8) から取得。')
                print('  ※ 先頭の = は不要。日本語テーブル名は DAX 内でシングルクォート: SUM(\'売上\'[数量])')
                return False
            tbl_name, measure_name = rest[2], rest[3]
            # DAX の取得: --dax インライン > --dax-file > _last_dax.dax
            dax = getattr(args, 'dax', None)
            if not dax:
                df = getattr(args, 'dax_file', None)
                path = smart_path_resolve(df) if df else _LAST_DAX_FILE
                if not path or not os.path.exists(path):
                    print(f"エラー: DAXファイルが見つかりません: {df or _LAST_DAX_FILE}")
                    print("  _last_dax.dax に式を書くか、--dax / --dax-file を指定してください。")
                    return False
                dax = read_code_file(path)
            if not dax or not dax.strip():
                print("エラー: DAX式が空です。")
                return False
            dax = dax.strip()
            if dax.startswith('='):            # Excel数式の癖で = を付けても通るように
                dax = dax[1:].strip()
            try:
                tbl = model.ModelTables.Item(tbl_name)
            except Exception:
                print(f"エラー: テーブル '{tbl_name}' が見つかりません。datamodel list で確認。")
                return False
            # 数値書式（既定 general）。引数付き書式は GetModelFormat* メソッドで取得
            #   （ModelFormat* プロパティは既定値専用で引数を渡せないため）。
            fmt_name = (getattr(args, 'format', None) or 'general').lower()
            dec_arg = getattr(args, 'decimals', None)
            try:
                decimals = int(dec_arg) if dec_arg is not None else 2
            except (TypeError, ValueError):
                print(f"エラー: --decimals は数値で指定してください: '{dec_arg}'")
                return False
            thousands = bool(getattr(args, 'thousands', False))
            symbol = getattr(args, 'symbol', None) or ''
            try:
                if fmt_name == 'general':
                    fmt = model.ModelFormatGeneral
                elif fmt_name in ('whole', 'wholenumber'):
                    fmt = model.GetModelFormatWholeNumber(thousands)
                elif fmt_name in ('decimal', 'decimalnumber'):
                    fmt = model.GetModelFormatDecimalNumber(thousands, decimals)
                elif fmt_name == 'currency':
                    # Symbol は通貨コード（USD/JPY/EUR 等）。グリフ（$ 等）は無効だが
                    # GetModelFormatCurrency では落ちず Add 時に例外になるため、
                    # フォールバックは Add 側で行う。
                    fmt = model.GetModelFormatCurrency(symbol, decimals)
                elif fmt_name in ('percent', 'percentage'):
                    fmt = model.GetModelFormatPercentageNumber(thousands, decimals)
                elif fmt_name in ('scientific', 'sci'):
                    fmt = model.GetModelFormatScientificNumber(decimals)
                else:
                    print(f"エラー: 未知の書式 '{fmt_name}'（general|whole|decimal|currency|percent|scientific）")
                    return False
            except Exception as e:
                print(f"エラー: 書式オブジェクトの取得に失敗: {str(e)[:120]}")
                return False
            desc = getattr(args, 'desc', None) or ''
            try:
                model.ModelMeasures.Add(measure_name, tbl, dax, fmt, desc)
            except Exception as e:
                # currency でグリフ等の無効な通貨コードだと Add 時に例外。
                # 既定の通貨記号で 1 回だけ再試行する。
                if fmt_name == 'currency' and symbol:
                    try:
                        model.ModelMeasures.Add(measure_name, tbl, dax,
                                                model.GetModelFormatCurrency('', decimals), desc)
                        print(f"  注意: 通貨コード '{symbol}' は無効。既定の通貨記号で作成しました（有効例: USD, JPY, EUR）。")
                        print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式=currency)")
                        print("（保存はしていません）")
                        return True
                    except Exception as e2:
                        e = e2
                print(f"エラー: メジャー作成に失敗しました: {str(e)[:200]}")
                print("  DAX 構文・テーブル/列名・シングルクォートを確認してください。")
                return False
            print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式={fmt_name})")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            if len(rest) < 3:
                print('使い方: datamodel measure delete <メジャー名>')
                return False
            measure_name = rest[2]
            ms = model.ModelMeasures
            for i in range(1, ms.Count + 1):
                if ms.Item(i).Name == measure_name:
                    ms.Item(i).Delete()
                    print(f"メジャー削除: {measure_name}")
                    print("（保存はしていません）")
                    return True
            print(f"エラー: メジャー '{measure_name}' が見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel measure <add|delete> ...')
        return False

    if action != 'list':
        print(f"未知のアクション: {action}（list|relation|measure）")
        return False

    # テーブル
    try:
        mts = model.ModelTables
        tn = mts.Count
    except Exception:
        mts = None
        tn = 0
    print(f"--- データモデル: テーブル ({tn}) ---")
    for i in range(1, tn + 1):
        mt = mts.Item(i)
        try:
            rc = mt.RecordCount
        except Exception:
            rc = '?'
        try:
            cols = [f"{mt.ModelTableColumns.Item(j).Name}({_MODEL_DATATYPE.get(int(mt.ModelTableColumns.Item(j).DataType), '?')})"
                    for j in range(1, int(mt.ModelTableColumns.Count) + 1)]
        except Exception:
            cols = []
        print(f"  {mt.Name}  ({rc}行)" + ("  列: " + " ".join(cols) if cols else ""))
        if any('(文字)' in c for c in cols) and any(k in c for c in cols for k in ('金額', '数量', '売上', '単価', '件数', '額')):
            print("    ⚠ 数値らしい列が「文字」型です。SUM のメジャーが無言で消えます（powerquery edit で型付け＝道具が自動で足す）")
    if tn == 0:
        print("  (なし)")

    # リレーションシップ
    try:
        rels = model.ModelRelationships
        rn = rels.Count
    except Exception:
        rels = None
        rn = 0
    print(f"--- リレーションシップ ({rn}) ---")
    for i in range(1, rn + 1):
        r = rels.Item(i)
        try:
            fkt = r.ForeignKeyTable.Name
            fkc = r.ForeignKeyColumn.Name
            pkt = r.PrimaryKeyTable.Name
            pkc = r.PrimaryKeyColumn.Name
            active = ''
            try:
                active = '' if r.Active else '  (無効)'
            except Exception:
                pass
            print(f"  {fkt}[{fkc}] → {pkt}[{pkc}]{active}")
        except Exception:
            print(f"  (リレーション {i}: 読み取り不可)")
    if rn == 0:
        print("  (なし)")

    # メジャー（対応バージョンのみ）
    try:
        ms = model.ModelMeasures
        mn = ms.Count
        print(f"--- メジャー ({mn}) ---")
        for i in range(1, mn + 1):
            me = ms.Item(i)
            try:
                tbl = me.AssociatedTable.Name
            except Exception:
                tbl = '?'
            print(f"  {me.Name}  (所属={tbl})")
        if mn == 0:
            print("  (なし)")
    except Exception:
        pass

    return True




def cmd_rehearse(args):
    """予行演習run: rehearse [excel_file] <マクロ名> [マクロ引数...]

    本体には指一本触れずにマクロの効果を検分する:
      1. 対象ブックの「今この瞬間」のコピーを作る（SaveCopyAs＝未保存の変更込み）
      2. 別インスタンス（DispatchEx・非表示）でコピーを開く
         （健診モードと同じく、開く瞬間の Workbook_Open は起こさない。
           実行中のイベント（Worksheet_Change 等）は本番同様に生かす）
      3. 実行前 snapshot → マクロ実行（ダイアログ安全解除は run-macro と同じく常設）
         → 実行後 snapshot
      4. snapshot-diff で「何が変わるか」を事実で報告し、結果コピーを保存して畳む

    マクロが途中で落ちても、そこまでの変化を diff で報告する（それも予行演習の収穫）。
    本体に焼くかどうかは、この報告を見て人が決める。
    ※ 非破壊が保証できるのはブックの中だけ。ファイル出力・メール送信など
      「外への副作用」を持つマクロまでは止められない（そこは人が読んで判断する）。
    """
    import gc
    import json
    import tempfile
    import vbam_core

    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: rehearse [excel_file] <マクロ名> [マクロ引数...]")
        print("  --auto-dialog ok|yes|no|cancel  実行中のダイアログをこのボタンで自動応答")
        print("  --input-text 値 [--input-text 値2 ...]  InputBox にこの値を入れて OK（出た順に1つずつ）")
        print("  --addins                        演習用Excelにアドイン・PERSONALを読み込む")
        print("  --timeout 秒                    制限時間。過ぎたら演習用 Excel を強制終了して報告（本体は無傷）")
        print("  --out <path>                    コピーの保存先（省略時は一時フォルダ）")
        print("  --discard                       報告後に結果コピーとsnapshotを削除")
        return False
    macro_name = rest[0]
    run_args = []
    for a in rest[1:]:
        v = _coerce_cell(a)
        run_args.append(a if v is None else v)

    # 本体は健診モード（読み取り専用・イベント無効）で掴む。閉じているブックを
    # ここで自動オープンする場合に Workbook_Open を起こすと、その痕跡ごと
    # SaveCopyAs されて「素の状態のコピー」でなくなるため（既に開いている
    # ブックには影響しない＝ユーザーの Excel はそのまま）
    xl_src, wb_src = get_workbook(target_file, readonly=True)
    src_name = wb_src.Name
    stem, ext = os.path.splitext(src_name)
    if not ext:
        ext = '.xlsx'
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out_opt = getattr(args, 'out_opt', None)
    copy_path = (os.path.abspath(out_opt) if out_opt
                 else os.path.join(tempfile.gettempdir(), f"{stem}_予行_{stamp}{ext}"))
    if os.path.exists(copy_path):
        print(f"エラー: コピー先が既に存在します: {copy_path}")
        return False
    try:
        wb_src.SaveCopyAs(copy_path)
    except Exception as e:
        print(f"エラー: コピーを作れませんでした: {e}")
        print("  （一度も保存していない新規ブックは、先に save-as で実体を作ってください）")
        return False
    print(f"予行演習: {src_name} のコピーで '{macro_name}' を試し撃ちします（本体は無傷）")
    print(f"  コピー: {copy_path}")

    before_path = copy_path + ".before.json"
    after_path = copy_path + ".after.json"

    # 演習用の Excel は必ず別インスタンス（DispatchEx）。ユーザーの Excel には触れない
    xl2 = win32com.client.DispatchEx("Excel.Application")
    pid = None
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(xl2.Hwnd)
    except Exception:
        pass
    inst = {"xl": xl2, "pid": pid}
    # 後始末の安全網として登録する（本線は下の finally で自前で畳む。
    # 万一そこへ辿り着けなくても cleanup_excel / アイドル解放が拾う）
    vbam_core._created_instances.append(inst)

    run_ok = True
    run_err = None
    result = None
    dlg_note = ""
    timed_out = False
    wb2 = None
    try:
        xl2.Visible = "--visible" in sys.argv or "-v" in sys.argv
        xl2.DisplayAlerts = False
        try:
            xl2.EnableEvents = False       # 開く瞬間の Workbook_Open は起こさない
        except Exception:
            pass
        if getattr(args, 'addins', False):
            load_excel_addins_and_personal(xl2)
        cmd_progress_note("コピーを別の Excel で開いています")
        wb2 = xl2.Workbooks.Open(copy_path, 0)   # UpdateLinks=0
        try:
            xl2.EnableEvents = True        # 実行中のイベントは本番同様に生かす
        except Exception:
            pass

        cmd_progress_note("実行前の snapshot を取っています")
        doc_before = _snapshot_doc(wb2)
        with open(before_path, 'w', encoding='utf-8') as f:
            json.dump(doc_before, f, ensure_ascii=False, indent=2)
        # 撃つ前に全体コンパイル（2026-09-17 夜）。コンパイルエラーのまま Run すると見えない Excel に窓が出て
        # --timeout まで固まる（修理の試験で 120 秒を丸ごと失った）。通らなければ撃たずにすぐ返す
        comp_doc = {}
        try:
            from vbam_vba import _compile_vbproject
            comp_doc = _compile_vbproject(xl2, wb2) or {}
        except Exception:
            comp_doc = {}
        if comp_doc.get('ok') is False:
            run_err = f"コンパイルエラーで撃てません（直してから撃つ）: {comp_doc.get('detail') or ''}".rstrip()
            raise RuntimeError(run_err)
        cmd_progress_note(f"'{macro_name}' をコピーで実行しています")

        # ブック名の ' は Excel 規約どおり '' に重ねる（cmd_run_macro と同じ流儀）
        quoted = wb2.Name.replace("'", "''")
        full_macro = f"'{quoted}'!{macro_name}"
        _auto = getattr(args, 'auto_dialog', None)
        # run-macro と同じハーネス（実行時エラーをダイアログにせず「実行時エラー 番号: 説明」で
        # 受ける）。引数付きは素の Run。置けないときも素の Run に落ちる（2026-08-23）
        harness = None
        if not run_args:
            harness = _prepare_run_harness(xl2, full_macro, macro_name)
        # --timeout: 演習用 Excel を強制終了して戻る（本体は無傷。固まった呼び出しは即戻る）
        _watchdog = _start_call_watchdog(getattr(args, 'timeout', None), owned_pid=pid)
        _watcher = _start_dialog_watcher(xl2, _auto, getattr(args, 'input_text', None))
        try:
            if harness is not None:
                ret = xl2.Application.Run(harness["entry"])
                ret = "" if ret is None else str(ret)
                if ret.startswith("OK|"):
                    result = ret[3:] or None
                elif ret.startswith("ERR|"):
                    parts = ret.split("|", 2)
                    run_ok = False
                    run_err = (f"実行時エラー {parts[1] if len(parts) > 1 else '?'}: "
                               f"{parts[2] if len(parts) > 2 else ''}")
                else:
                    result = ret
            else:
                result = xl2.Application.Run(full_macro, *run_args)
        except Exception as e:
            run_ok = False
            run_err = _com_error_text(e)
        finally:
            _watchdog.stop()
            if _watchdog.fired:
                timed_out = True
                run_ok = False
                run_err = f"時間切れ（{_watchdog.seconds:g}秒）: 演習用 Excel を強制終了しました"
            if _watcher is not None:
                _watcher.stop()
                dlg_note = _dialog_watcher_note(_watcher, _auto)
            if harness is not None and not timed_out:
                _remove_run_harness(harness)

        if timed_out:
            print("（時間切れのため、実行後 snapshot と結果コピーの保存はありません。コピーは実行前の状態のままです）")
        else:
            doc_after = _snapshot_doc(wb2)
            with open(after_path, 'w', encoding='utf-8') as f:
                json.dump(doc_after, f, ensure_ascii=False, indent=2)

            try:
                wb2.Save()                     # 結果コピーは検分できる形で残す
            except Exception as e:
                print(f"⚠ 結果コピーを保存できませんでした: {e}")
            wb2.Close(SaveChanges=False)
            wb2 = None
    except Exception as e:
        print(f"エラー: 予行演習を続けられませんでした: {e}")
        run_ok = False
        if run_err is None:
            run_err = str(e)
    finally:
        if wb2 is not None:
            try:
                wb2.Close(SaveChanges=False)
            except Exception:
                pass
            wb2 = None
        try:
            xl2.Quit()
        except Exception:
            pass
        xl2 = None
        # 自前で畳めたので安全網から外す。外せたときだけ参照も手放す
        # （外せなかった場合は cleanup_excel 側が参照ごと始末する）
        try:
            vbam_core._created_instances.remove(inst)
            inst["xl"] = None
        except ValueError:
            pass
        gc.collect()

    print("-" * 60)
    if run_ok:
        print(f"マクロ実行: 成功  戻り値: {result}")
    else:
        print(f"マクロ実行: 失敗  {run_err}")
        print("  （落ちるまでに変えたものがあれば、下の差分に出ます）")
    if dlg_note:
        print(dlg_note)

    if os.path.exists(before_path) and os.path.exists(after_path):
        ns = argparse.Namespace(posargs=[before_path, after_path],
                                max_opt=getattr(args, 'max_opt', None))
        cmd_snapshot_diff(ns)
    else:
        print("⚠ snapshot が揃わなかったため、差分報告はありません。")

    if getattr(args, 'discard', False):
        for p in (copy_path, before_path, after_path):
            try:
                os.remove(p)
            except OSError:
                pass
        print("（--discard 指定により、結果コピーと snapshot を削除しました）")
    else:
        print(f"結果コピー: {copy_path}")
        print(f"  実行前/後 snapshot: {before_path} / {after_path}")
        print("  本体に焼いてよければ run-macro で本番実行を。コピーは不要になったら削除してください。")
    return run_ok


__all__ = [
    '_XL_AXIS',
    '_XL_CHART_TYPE',
    '_XL_CHART_TYPE_NAME',
    '_XL_CONN_TYPE',
    '_XL_LEGEND_POS',
    '_XL_PIVOT_FUNC',
    '_XL_PIVOT_ORIENT',
    '_XL_TRENDLINE',
    '_connection_used_by_table',
    '_find_pivot',
    '_find_pivot_or_table',
    '_restore_connection',
    '_snapshot_connection',
    '_unique_sheet_name',
    'cmd_calc_mode',
    'cmd_chart',
    'cmd_chart_config',
    'cmd_connection',
    'cmd_datamodel',
    'cmd_pivot',
    'cmd_pivot_calc',
    'cmd_pivot_field',
    'cmd_powerquery',
    'cmd_rehearse',
    'cmd_slicer',
]
