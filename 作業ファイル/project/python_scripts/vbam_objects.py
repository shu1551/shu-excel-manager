# -*- coding: utf-8 -*-
"""vbam_objects.py — ブックのグラフとピボットの姿を読み、正解と比べる（鍛える回路の目・2026-09-17 深夜）。

それまで鍛える回路（vbam_forge）はシートのセルの値しか比べていなかった＝グラフやピボットを作るマクロは鍛えられなかった。
shu「グラフやピボットは実際に使うケースも多い・徹底的に・デザインも考えて作れるようになれれば最高」。

読む（COM）: snapshot_objects(wb) → {'charts': [...], 'pivots': [...]}（無ければどちらも []）
  グラフ: 置いたシート／グラフシートか／種類／タイトル／凡例（有無・位置）／軸のタイトルと表示形式／目盛線／
          系列ごとの名前・値・項目・データラベル（値・パーセント）・塗りの色／棒の間隔／グラフのフォント
  ピボット: 置いたシート／行・列・フィルタのフィールド（並び順）／値のフィールド（元の列・集計の方法・表示形式）／
          総計（行・列）／出てきた表の値（見出しの文字は Excel の言語で変わるので、数と項目の名前だけ）
比べる（純 Python）: compare_objects(expect, got) → 不一致の行（空なら合格）。位置・名前は比べない。
"""
import re

CHART_TYPES = {-4111: '複合', 51: '集合縦棒', 52: '積み上げ縦棒', 57: '集合横棒', 58: '積み上げ横棒', 4: '折れ線', 65: 'マーカー付き折れ線',
               5: '円', -4120: 'ドーナツ', 1: '面', 76: '積み上げ面', -4169: '散布図', 15: 'バブル', -4151: 'レーダー',
               81: 'マーカー付きレーダー', 82: '塗りつぶしレーダー', 74: '散布図（直線とマーカー）', 75: '散布図（直線）',
               59: '100% 積み上げ縦棒', 53: '100% 積み上げ縦棒'}
_BAR_TYPES = {51, 52, 53, 57, 58, 59, 60}
LEGEND_POS = {-4107: '下', -4131: '左', -4152: '右', -4160: '上', 2: '右上'}
_PIVOT_HEAD_RE = re.compile(r'^(行ラベル|列ラベル|総計|合計 / |個数 / |平均 / |最大 / |最小 / |Row Labels|Column Labels|Grand Total|Sum of |Count of )')
PIVOT_CALCS = {-4143: '標準', 2: '基準値との差分', 3: '基準値に対する比率', 4: '基準値との差分の比率', 5: '累計', 6: '行集計に対する比率',
               7: '列集計に対する比率', 8: '総計に対する比率', 9: '指数'}
PIVOT_FUNCS = {-4157: '合計', -4112: '個数', -4106: '平均', -4136: '最大', -4139: '最小', -4113: '数値の個数'}


def _s(v):
    return '' if v is None else str(v)


def _num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 6)
    except (TypeError, ValueError):
        return _s(v)


def _uncomma(v):
    """「5,954,000」のようにカンマで区切った数の文字は数にする（2026-09-18 第二期: 散布図の横軸をセルの範囲で渡すと、
    XValues が表示形式つきの文字で返り、点の位置は同じなのに外れた）。"""
    if isinstance(v, str) and re.fullmatch(r'-?\d{1,3}(,\d{3})+(\.\d+)?', v.strip()):
        return v.strip().replace(',', '')
    return v


def _try(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def _list(v):
    if v is None:
        return []
    if isinstance(v, (list, tuple)):
        return list(v)
    return [v]


def _chart_info(ch, sheet, chartsheet=False):
    info = {'sheet': sheet, 'chartsheet': chartsheet, 'type': _try(lambda: int(ch.ChartType)),
            'title': _try(lambda: _s(ch.ChartTitle.Text) if ch.HasTitle else '', ''),
            'legend': _try(lambda: LEGEND_POS.get(int(ch.Legend.Position), int(ch.Legend.Position)) if ch.HasLegend else 'なし', 'なし'),
            'font': _try(lambda: _s(ch.ChartArea.Format.TextFrame2.TextRange.Font.Name), ''),
            'series': [], 'axes': {}}
    for key, idx in (('項目軸', 1), ('数値軸', 2)):
        ax = _try(lambda idx=idx: ch.Axes(idx))
        if ax is None:
            continue
        info['axes'][key] = {'title': _try(lambda ax=ax: _s(ax.AxisTitle.Text) if ax.HasTitle else '', ''),
                             'format': _try(lambda ax=ax: _s(ax.TickLabels.NumberFormat), ''),
                             'gridlines': _try(lambda ax=ax: bool(ax.HasMajorGridlines), False),
                             'reverse': _try(lambda ax=ax: bool(ax.ReversePlotOrder), False)}
    # 見せ方（2026-09-18 未明: 「見やすいグラフ」を型で鍛える＝タイトルの文字の大きさ・太さ）
    info['title_size'] = _try(lambda: round(float(ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size), 1) if ch.HasTitle else None)
    info['title_bold'] = _try(lambda: bool(ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold) if ch.HasTitle else None)
    info['gap'] = _try(lambda: int(ch.ChartGroups(1).GapWidth))
    info['pivot_chart'] = _try(lambda: ch.PivotLayout is not None, False)       # ピボットグラフ（2026-09-18）
    n = _try(lambda: int(ch.SeriesCollection().Count), 0)
    for i in range(1, n + 1):
        s = ch.SeriesCollection(i)
        labels = _try(lambda s=s: bool(s.HasDataLabels), False)
        info['series'].append({
            'name': _try(lambda s=s: _s(s.Name), ''),
            'values': [_num(x) for x in _list(_try(lambda s=s: s.Values, []))],
            'categories': [_s(_num(_uncomma(x))) for x in _list(_try(lambda s=s: s.XValues, []))],
            'labels': ('なし' if not labels else
                       ('パーセント' if _try(lambda s=s: bool(s.DataLabels().ShowPercentage), False) else '値')),
            'color': _try(lambda s=s: '#%06X' % _bgr_to_rgb(int(s.Format.Fill.ForeColor.RGB)), ''),
            'point_colors': _point_colors(s),
            # 複合グラフ（2026-09-18）: 系列ごとの種類と、どちらの数値軸か（1＝主軸・2＝第 2 軸）
            'type': _try(lambda s=s: int(s.ChartType)), 'axis_group': _try(lambda s=s: int(s.AxisGroup), 1),
        })
        if info['series'][-1]['type'] in _BAR_TYPES:
            # 棒の塗りの有無（2026-09-18 第二期: 増減の説明＝積み上げ縦棒の「土台」を塗りなしで見えなくする）
            info['series'][-1]['fill'] = _try(lambda s=s: bool(s.Format.Fill.Visible), True)
        if info['series'][-1]['type'] in (-4169, 72, 73, 74, 75):
            # 散布図のマーカーの塗りは作り方で既定の色が変わる（#000000 と読める）＝依頼に書かない色は比べない（2026-09-18 第二期）
            info['series'][-1]['color'] = None
        if info['series'][-1]['type'] in (4, 65, 63, 64, 66, 67, -4151, 81):
            # 折れ線の見せ方（2026-09-18: 見やすい推移グラフ）: 線の太さ・マーカーの大きさ・線の色（レーダーも線で比べる・第二期）
            info['series'][-1].update({'line_weight': _try(lambda s=s: round(float(s.Format.Line.Weight), 2)),
                                       'marker_size': _try(lambda s=s: int(s.MarkerSize)) if info['series'][-1]['type'] not in (4, -4151) else None,
                                       'line_color': _try(lambda s=s: '#%06X' % _bgr_to_rgb(int(s.Format.Line.ForeColor.RGB)), '')})
    ax2 = _try(lambda: ch.Axes(2, 2) if ch.HasAxis(2, 2) else None)
    if ax2 is not None:
        info['axes']['第2数値軸'] = {'title': _try(lambda: _s(ax2.AxisTitle.Text) if ax2.HasTitle else '', ''),
                                   'format': _try(lambda: _s(ax2.TickLabels.NumberFormat), ''),
                                   'gridlines': _try(lambda: bool(ax2.HasMajorGridlines), False),
                                   'max': _try(lambda: None if ax2.MaximumScaleIsAuto else _num(ax2.MaximumScale))}
    return info


def _point_colors(s, limit=60):
    """系列の点（棒）ごとの塗りの色。系列の色と全部同じなら []（強調の色分けが無い）。円グラフは点ごとに色が変わるのが既定＝読まない。"""
    try:
        if int(s.Parent.Parent.ChartType) in (5, -4120, 69, 70):
            return []
        base = int(s.Format.Fill.ForeColor.RGB)
        n = int(s.Points().Count)
        cols = ['#%06X' % _bgr_to_rgb(int(s.Points(i).Format.Fill.ForeColor.RGB)) for i in range(1, min(n, limit) + 1)]
    except Exception:
        return []
    return [] if all(c == '#%06X' % _bgr_to_rgb(base) for c in cols) else cols


def _bgr_to_rgb(v):
    """Excel の色の数（R + G*256 + B*65536）→ 0xRRGGBB。"""
    return ((v & 0xFF) << 16) | (((v >> 8) & 0xFF) << 8) | ((v >> 16) & 0xFF)


def _field_names(fields):
    out = []
    n = _try(lambda: int(fields.Count), 0)
    items = []
    for i in range(1, n + 1):
        f = fields.Item(i)
        # 値が 2 つ以上のときの「Σ 値」は SourceName が数（エラーの値 -2146826246）で返る＝名前で見る（2026-09-18）
        src = _try(lambda f=f: f.SourceName)
        items.append((_try(lambda f=f: int(f.Position), i), src if isinstance(src, str) and src else _try(lambda f=f: _s(f.Name), '')))
    for _p, name in sorted(items):
        if name and name not in ('Values', '値', 'Σ 値', 'データ', '∑ 値', 'Σ Values'):
            out.append(name)
    return out


def _pivot_info(pt, sheet):
    data = []
    n = _try(lambda: int(pt.DataFields.Count), 0)
    for i in range(1, n + 1):
        d = pt.DataFields.Item(i)
        data.append({'source': _try(lambda d=d: _s(d.SourceName), ''),
                     'func': PIVOT_FUNCS.get(_try(lambda d=d: int(d.Function)), _try(lambda d=d: int(d.Function))),
                     'format': _try(lambda d=d: _s(d.NumberFormat), ''),
                     # 値の表示方法（2026-09-18: 構成比のピボット）: 計算の種類（標準・列集計に対する比率など）
                     'calc': PIVOT_CALCS.get(_try(lambda d=d: int(d.Calculation), -4143), _try(lambda d=d: int(d.Calculation)))})
    body = []
    # 値のフィールドの見出し（「合計 / 支出額」「支出額合計」など）は書き方の違いで、項目ではない（2026-09-18: ダッシュボードで
    # 「支出額合計」と名付けたマクロが、ほかは全部そろっているのに外れ 1 で止まった）
    captions = {_try(lambda i=i: _s(pt.DataFields.Item(i).Name).strip(), '') for i in range(1, n + 1)}
    for row in _list(_try(lambda: pt.TableRange1.Value, [])):
        cells = []
        for v in _list(row):
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cells.append(_num(v))
            elif isinstance(v, str) and v.strip() and not _PIVOT_HEAD_RE.match(v.strip()) and v.strip() not in captions:
                cells.append(v.strip())                    # 行ラベル・総計・「合計 / 支出額」などの見出しの文字は比べない
        body.append(cells)
    # 置いた場所（行・列）＝比べるときの並びにだけ使う（2026-09-18 第二期: 同じシートにピボットが 2 つあると、ピボットを触った後で
    # PivotTables の並びが入れ替わり、中身は同じなのに 1 つ目と 2 つ目を取り違えて外れにしていた）
    at = _try(lambda: [int(pt.TableRange1.Row), int(pt.TableRange1.Column)])
    return {'sheet': sheet, 'at': at, 'rows': _field_names(_try(lambda: pt.RowFields, None)) if _try(lambda: pt.RowFields) else [],
            'cols': _field_names(pt.ColumnFields) if _try(lambda: pt.ColumnFields) else [],
            'pages': _field_names(pt.PageFields) if _try(lambda: pt.PageFields) else [],
            'data': data, 'row_grand': _try(lambda: bool(pt.RowGrand)), 'col_grand': _try(lambda: bool(pt.ColumnGrand)),
            'numbers': sorted(c for r in body for c in r if isinstance(c, (int, float))),
            'labels': sorted({c for r in body for c in r if isinstance(c, str)}),
            **_pivot_design(pt)}


PIVOT_LAYOUTS = {(0, False): '表形式', (1, False): 'アウトライン', (1, True): 'コンパクト', (0, True): '表形式'}


def _pivot_design(pt):
    """ピボットの見せ方（2026-09-18: いつもの月別ピボット）: スタイル・行のレイアウト・ラベルの繰り返し・小計・空白セルの表示。"""
    rows = []
    rf = _try(lambda: pt.RowFields, None)
    n = _try(lambda: int(rf.Count), 0) if rf is not None else 0
    for i in range(1, n + 1):
        f = rf.Item(i)
        subs = _try(lambda f=f: f.Subtotals, None)
        subs = list(subs) if isinstance(subs, (list, tuple)) else []
        rows.append((_try(lambda f=f: int(f.Position), i),
                     {'field': _try(lambda f=f: _s(f.SourceName) or _s(f.Name), ''),
                      'layout': PIVOT_LAYOUTS.get((_try(lambda f=f: int(f.LayoutForm), 1), _try(lambda f=f: bool(f.LayoutCompactRow), True)), '?'),
                      'repeat': _try(lambda f=f: bool(f.RepeatLabels), False),
                      'subtotal': bool(any(subs)) if subs else None}))
    null = _try(lambda: _s(pt.NullString), '') if _try(lambda: bool(pt.DisplayNullString), False) else ''
    order = []
    if n == 1:                                            # 行のフィールドが 1 つのときの並び（2026-09-18: 大きい順に並べるピボット）
        vals = _list(_try(lambda: pt.RowRange.Value, []))
        for row in vals[1:]:
            v = _list(row)[0] if _list(row) else None
            if isinstance(v, str) and v.strip() and not _PIVOT_HEAD_RE.match(v.strip()):
                order.append(v.strip())
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                order.append(_s(_num(v)))
    return {'style': _try(lambda: _s(pt.TableStyle2), ''), 'row_fields': [r for _p, r in sorted(rows, key=lambda x: x[0])],
            'null': null, 'row_order': order,
            # 元の範囲（2026-09-18: 明細を足した後のピボットの更新）。R1C1 の文字（シート名つき）
            'source': re.sub(r"^'?([^'!]+)'?!", r"\1!", _try(lambda: _s(pt.SourceData), ''))}


TOTALS = {0: 'なし', 1: '合計', 2: '平均', 3: '個数', 4: '数値の個数', 5: '最大', 6: '最小', 7: '標本標準偏差', 8: '標本分散', 9: 'ユーザー設定'}


def _table_info(lo, sheet):
    cols = []
    n = _try(lambda: int(lo.ListColumns.Count), 0)
    for i in range(1, n + 1):
        lc = lo.ListColumns(i)
        formula = ''
        body = _try(lambda lc=lc: lc.DataBodyRange)
        if body is not None and _try(lambda body=body: bool(body.Cells(1, 1).HasFormula), False):
            formula = _try(lambda body=body: _s(body.Cells(1, 1).Formula), '')
        cols.append({'name': _try(lambda lc=lc: _s(lc.Name), ''), 'formula': formula,
                     'total': TOTALS.get(_try(lambda lc=lc: int(lc.TotalsCalculation), 0), '')
                     if _try(lambda: bool(lo.ShowTotals), False) else ''})
    info = {'sheet': sheet, 'name': _try(lambda: _s(lo.Name), ''),
            'range': _try(lambda: _s(lo.Range.Address).replace('$', ''), ''),
            'style': _try(lambda: _s(lo.TableStyle.Name), ''),
            'totals': _try(lambda: bool(lo.ShowTotals), False), 'filter': _try(lambda: bool(lo.ShowAutoFilter), False),
            'stripes': _try(lambda: bool(lo.ShowTableStyleRowStripes), False), 'columns': cols,
            # 絞り込みで隠れた行（本体の何行目か・1 始まり）。条件の書き方（<>0・>0）でなく見え方で比べる（2026-09-18）
            'hidden': _hidden_rows(lo)}
    # パワークエリの読み込み先（外部データの表）は、読み込まれた値も比べる（シートの値はケースのシートしか比べないため）
    query = _try(lambda: _s(lo.QueryTable.WorkbookConnection.Name), '')
    if query:
        # 接続の名前は作り方で変わる（「接続」「クエリ - Q課別集計」）＝接続文字列の Location（クエリ名）で比べる
        conn = _try(lambda: _s(lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection), '')
        m = re.search(r'Location=("?)([^;"]+)\1', conn)
        info['query'] = m.group(2) if m else re.sub(r'^クエリ - |^Query - ', '', query)
        info['values'] = [[_s(_num(v)) if isinstance(v, (int, float)) else _s(v) for v in _list(row)]
                          for row in _list(_try(lambda: lo.Range.Value, []))]
    return info


def _hidden_rows(lo):
    """絞り込み中のテーブルで隠れた本体の行（1 始まり）。絞り込んでいなければ []（COM は見える範囲の塊を 1 回で読む）。"""
    if not _try(lambda: bool(lo.AutoFilter.FilterMode), False):
        return []
    body = _try(lambda: lo.DataBodyRange)
    if body is None:
        return []
    try:
        top, n = int(body.Row), int(body.Rows.Count)
        vis = body.Columns(1).SpecialCells(12)                      # xlCellTypeVisible
        shown = set()
        for k in range(1, int(vis.Areas.Count) + 1):
            a = vis.Areas(k)
            shown.update(range(int(a.Row) - top + 1, int(a.Row) - top + 1 + int(a.Rows.Count)))
        return [i for i in range(1, n + 1) if i not in shown]
    except Exception:
        return list(range(1, _try(lambda: int(body.Rows.Count), 0) + 1))   # 見える行が 1 つも無い（SpecialCells がエラー）


def _query_info(q):
    return {'name': _try(lambda: _s(q.Name), ''), 'formula': _try(lambda: _s(q.Formula), '')}


SPARK_TYPES = {1: '折れ線', 2: '縦棒', 3: '勝敗'}


def _addr(a):
    """番地の文字からシート名と $ を落とす（'月別'!$B$2:$M$2 → B2:M2）。"""
    return re.sub(r"^.*!", '', _s(a)).replace('$', '')


def _sparklines(sh, name):
    """シートのスパークライン（2026-09-18 第二期: スパークラインの列のお題）。1 本ずつ＝置いたセル・元の範囲・種類・
    最高点／最低点／マーカーの印・色。グループの分け方（1 本ずつ作る／まとめて作る）は見た目が同じなので比べない。"""
    out = []
    groups = _try(lambda: sh.Cells.SparklineGroups, None)
    for i in range(1, (_try(lambda: int(groups.Count), 0) if groups is not None else 0) + 1):
        sg = groups.Item(i)
        common = {'sheet': name, 'type': SPARK_TYPES.get(_try(lambda sg=sg: int(sg.Type)), _try(lambda sg=sg: int(sg.Type))),
                  'high': _try(lambda sg=sg: bool(sg.Points.Highpoint.Visible), False),
                  'low': _try(lambda sg=sg: bool(sg.Points.Lowpoint.Visible), False),
                  'markers': _try(lambda sg=sg: bool(sg.Points.Markers.Visible), False),
                  'color': _try(lambda sg=sg: '#%06X' % _bgr_to_rgb(int(sg.SeriesColor.Color)), '')}
        for j in range(1, _try(lambda sg=sg: int(sg.Count), 0) + 1):
            it = sg.Item(j)
            out.append(dict(common, cell=_addr(_try(lambda it=it: it.Location.Address, '')),
                            source=_addr(_try(lambda it=it: it.SourceData, ''))))
    return out


def snapshot_objects(wb):
    """ブックのグラフ（シートに置いたもの・グラフシート）とピボットとテーブル。読めなければ空。"""
    charts, pivots, tables, sparks = [], [], [], []
    for sh in _list(_try(lambda: list(wb.Worksheets), [])):
        name = _try(lambda sh=sh: _s(sh.Name), '')
        sparks += _sparklines(sh, name)
        los = _try(lambda sh=sh: sh.ListObjects, None)
        for i in range(1, (_try(lambda los=los: int(los.Count), 0) if los is not None else 0) + 1):
            tables.append(_table_info(los.Item(i), name))
        cos = _try(lambda sh=sh: sh.ChartObjects(), None)
        for i in range(1, (_try(lambda cos=cos: int(cos.Count), 0) if cos is not None else 0) + 1):
            co = cos.Item(i)
            charts.append((_try(lambda co=co: (float(co.Top), float(co.Left)), (0, 0)), _chart_info(co.Chart, name)))
        pts = _try(lambda sh=sh: sh.PivotTables(), None)
        for i in range(1, (_try(lambda pts=pts: int(pts.Count), 0) if pts is not None else 0) + 1):
            pivots.append(_pivot_info(pts.Item(i), name))
    for ch in _list(_try(lambda: list(wb.Charts), [])):
        charts.append(((0, 0), _chart_info(ch, _try(lambda ch=ch: _s(ch.Name), ''), chartsheet=True)))
    queries = []
    nq = _try(lambda: int(wb.Queries.Count), 0)
    for i in range(1, nq + 1):
        queries.append(_query_info(wb.Queries.Item(i)))
    return {'charts': [c for _pos, c in sorted(charts, key=lambda x: (x[1]['sheet'], x[0]))], 'pivots': pivots,
            'tables': sorted(tables, key=lambda t: (t['sheet'], t['range'])),
            'queries': sorted(queries, key=lambda q: q['name']), 'slicers': _slicers(wb),
            'sparklines': sorted(sparks, key=lambda s: (s['sheet'], s['cell']))}


def _slicers(wb):
    """スライサー（2026-09-18: ピボットのダッシュボードのお題）: 絞り込む列・置いたシート・元（ピボット/テーブル）。名前と位置は比べない。"""
    out = []
    caches = _try(lambda: wb.SlicerCaches, None)
    for i in range(1, (_try(lambda: int(caches.Count), 0) if caches is not None else 0) + 1):
        sc = caches.Item(i)
        kind = 'ピボット' if _try(lambda sc=sc: int(sc.PivotTables.Count), 0) else 'テーブル'
        n = _try(lambda sc=sc: int(sc.Slicers.Count), 0)
        for j in range(1, n + 1):
            sl = sc.Slicers.Item(j)
            out.append({'field': _try(lambda sc=sc: _s(sc.SourceName), ''), 'kind': kind,
                        'sheet': _try(lambda sl=sl: _s(sl.Shape.Parent.Name), ''), 'caption': _try(lambda sl=sl: _s(sl.Caption), '')})
    return sorted(out, key=lambda x: (x['sheet'], x['field']))


def has_objects(wb):
    """グラフ・ピボット・テーブルが 1 つでもあるか（読む手間を省く）。"""
    try:
        if int(wb.Charts.Count):
            return True
        if _try(lambda: int(wb.SlicerCaches.Count), 0):
            return True
        if _try(lambda: int(wb.Queries.Count), 0):
            return True
        for sh in wb.Worksheets:
            if int(sh.ChartObjects().Count) or int(sh.PivotTables().Count) or int(sh.ListObjects.Count):
                return True
            if _try(lambda sh=sh: int(sh.Cells.SparklineGroups.Count), 0):
                return True
    except Exception:
        return False
    return False


def _cmp_list(label, e, g, out, limit=6):
    if e != g:
        out.append(f"{label}: 期待 {e[:limit] if isinstance(e, list) else e!r} ／ マクロ後 "
                   f"{g[:limit] if isinstance(g, list) else g!r}")


def _items_note(e):
    """期待のグラフの項目の数（2026-09-18 未明: 9 課から横棒の境目を、表ごとの項目の数が見えず 10 と当て推量した）。"""
    es = e.get('series') or []
    n = len(es[0].get('categories') or es[0].get('values') or []) if es else 0
    return f"（期待のグラフの項目 {n} 個）" if n else ''


def compare_objects(expect, got):
    """グラフとピボットの姿を比べる（純 Python）→ 不一致の行。expect が None なら比べない（古い台帳）。"""
    if expect is None:
        return []
    got = got or {'charts': [], 'pivots': []}
    out = []
    ec, gc = expect.get('charts') or [], got.get('charts') or []
    if len(ec) != len(gc):
        out.append(f"グラフの数: 期待 {len(ec)} ／ マクロ後 {len(gc)}"
                   + ("（前のグラフが残っている・作り直しで増えた）" if len(gc) > len(ec)
                      else "（グラフを作っていない。ChartObjects.Add で本物のグラフを作る）"))
    for k, (e, g) in enumerate(zip(ec, gc), 1):
        lab = f"グラフ {k}"
        if e.get('sheet') != g.get('sheet') or e.get('chartsheet') != g.get('chartsheet'):
            out.append(f"{lab}: 置き場所 期待「{e.get('sheet')}」{'（グラフシート）' if e.get('chartsheet') else ''} ／ "
                       f"マクロ後「{g.get('sheet')}」{'（グラフシート）' if g.get('chartsheet') else ''}")
        if e.get('pivot_chart') and not g.get('pivot_chart'):
            out.append(f"{lab}: ピボットグラフ 期待 あり ／ マクロ後 ふつうのグラフ（ピボットの TableRange1 を SetSourceData すると"
                       "ピボットグラフになる。セルの範囲を写して作らない）")
        if e.get('type') != g.get('type'):
            out.append(f"{lab}: 種類 期待 {CHART_TYPES.get(e.get('type'), e.get('type'))} ／ マクロ後 {CHART_TYPES.get(g.get('type'), g.get('type'))}"
                       f"（ChartType {e.get('type')}／{g.get('type')}）{_items_note(e)}")
        for key, name in (('title', 'タイトル'), ('legend', '凡例'), ('font', 'フォント'), ('gap', '棒の間隔（GapWidth）')):
            if key == 'gap' and e.get('type') not in _BAR_TYPES:
                continue                                    # 棒の間隔は棒グラフだけ（円・折れ線は作り方で既定値が変わるだけ）
            if key in e and e.get(key) != g.get(key):
                # 2026-09-18: 見やすい棒グラフで AI がタイトル・軸・ラベルに別々に Meiryo UI を当て、全体の文字が空のまま 2 往復
                note = ("（グラフ全体の文字＝ch.ChartArea.Format.TextFrame2.TextRange.Font.Name。タイトル・軸・ラベルに別々に"
                        "当てても全体の文字にはならない。全体に当てた後でタイトルの大きさ・太字を当てる）" if key == 'font' else '')
                out.append(f"{lab}: {name} 期待 {e.get(key)!r} ／ マクロ後 {g.get(key)!r}{note}")
        for key, name in (('title_size', 'タイトルの文字の大きさ'), ('title_bold', 'タイトルの太字')):
            if e.get(key) is not None and e.get(key) != g.get(key):
                out.append(f"{lab}: {name} 期待 {e.get(key)!r} ／ マクロ後 {g.get(key)!r}")
        ea_all, ga_all = e.get('axes') or {}, g.get('axes') or {}
        if ((ea_all.get('項目軸') or {}).get('format') != (ga_all.get('項目軸') or {}).get('format')
                and (ga_all.get('項目軸') or {}).get('format') == (ea_all.get('数値軸') or {}).get('format')):
            # 2026-09-18 未明: 見やすいグラフの横棒で、横に見える軸を項目軸だと思って表示形式と目盛線を当てた
            # 2026-09-18: 複合グラフでも Axes(1) に #,##0 を当てて 2 往復くり返した（棒グラフだけの知らせだった）
            out.append(f"{lab}: 項目軸（Axes(xlCategory)＝課名などの軸）に数値軸の表示形式を当てています。"
                       "金額の軸は Axes(xlValue, xlPrimary)、第 2 軸は Axes(xlValue, xlSecondary)。"
                       "横棒グラフでも項目軸が縦・数値軸が横に見えるだけで、軸の役は入れ替わりません")
        for ax, ea in (e.get('axes') or {}).items():
            ga = (g.get('axes') or {}).get(ax) or {}
            for key, name in (('title', 'タイトル'), ('format', '表示形式'), ('gridlines', '目盛線'), ('reverse', '並びの向き（ReversePlotOrder）'),
                              ('max', '最大値（MaximumScale）')):
                if key in ('reverse', 'max') and key not in ea:
                    continue
                if ea.get(key) != ga.get(key):
                    out.append(f"{lab}: {ax}の{name} 期待 {ea.get(key)!r} ／ マクロ後 {ga.get(key)!r}"
                               + (f"（期待の種類 {CHART_TYPES.get(e.get('type'), e.get('type'))}）{_items_note(e)}" if key == 'reverse' else ''))
        es, gs = e.get('series') or [], g.get('series') or []
        en, gn = [x.get('name') for x in es], [x.get('name') for x in gs]
        if en != gn and en and set(en) <= set(gn):
            # 系列の選び方・並びの手がかり（2026-09-18: 見やすい推移グラフで「年間の合計の大きい 5 課を大きい順」が読めず 7 往復）
            def _total(x):
                return sum(v for v in (x.get('values') or []) if isinstance(v, (int, float)))
            by_sum = [x.get('name') for x in sorted(gs, key=lambda x: -_total(x))]
            if en == by_sum[:len(en)]:
                out.append(f"{lab}: 期待の系列は、値の合計の大きい順に並んでいます" + (f"（上位 {len(en)} 本だけ・全部で {len(gn)} 本ある）"
                                                                    if len(en) < len(gn) else "") + f": {en[:8]}")
            elif en == by_sum[::-1][:len(en)]:
                out.append(f"{lab}: 期待の系列は、値の合計の小さい順に並んでいます: {en[:8]}")
        if len(es) != len(gs):
            out.append(f"{lab}: 系列の数 期待 {len(es)}（{[x.get('name') for x in es]}）／ マクロ後 {len(gs)}（{[x.get('name') for x in gs]}）"
                       "（行と列の向き・合計の行や列を入れた・見出しを系列にした）")
        for j, (a, b) in enumerate(zip(es, gs), 1):
            sl = f"{lab} 系列 {j}「{a.get('name')}」"
            if a.get('name') != b.get('name'):
                out.append(f"{sl}: 名前 期待 {a.get('name')!r} ／ マクロ後 {b.get('name')!r}")
            _cmp_list(f"{sl}: 値", a.get('values'), b.get('values'), out)
            _cmp_list(f"{sl}: 項目", a.get('categories'), b.get('categories'), out)
            for key, name in (('line_weight', '線の太さ（Format.Line.Weight）'), ('marker_size', 'マーカーの大きさ（MarkerSize）'),
                              ('line_color', '線の色（Format.Line.ForeColor.RGB）')):
                if key in a and a.get(key) is not None and a.get(key) != b.get(key):
                    out.append(f"{sl}: {name} 期待 {a.get(key)!r} ／ マクロ後 {b.get(key)!r}")
            if 'type' in a and a.get('type') != b.get('type'):
                out.append(f"{sl}: 系列の種類 期待 {CHART_TYPES.get(a.get('type'), a.get('type'))} ／ マクロ後 {CHART_TYPES.get(b.get('type'), b.get('type'))}"
                           f"（Series.ChartType {a.get('type')}／{b.get('type')}）")
            if 'axis_group' in a and a.get('axis_group') != b.get('axis_group'):
                out.append(f"{sl}: 数値軸 期待 {'第 2 軸' if a.get('axis_group') == 2 else '主軸'} ／ マクロ後 {'第 2 軸' if b.get('axis_group') == 2 else '主軸'}"
                           "（Series.AxisGroup = 2 で第 2 軸。ChartType を変えた後に当てる）")
            if a.get('labels') != b.get('labels'):
                out.append(f"{sl}: データラベル 期待 {a.get('labels')!r} ／ マクロ後 {b.get('labels')!r}")
            if 'fill' in a and a.get('fill') != b.get('fill'):
                out.append(f"{sl}: 塗り 期待 {'あり' if a.get('fill') else 'なし'} ／ マクロ後 {'あり' if b.get('fill') else 'なし'}"
                           "（見えない棒は Series.Format.Fill.Visible = msoFalse）")
            if a.get('fill') is False:
                continue                                    # 塗りの無い棒の色は見えない＝比べない
            if 'point_colors' in a:
                # 見た目の色で比べる＝棒ごとの色（無ければ系列の色が全部の棒の色）。系列の色を当てて 1 本だけ塗るのも、
                # 1 本ずつ塗るのも同じ見た目（2026-09-18 未明: 棒を 1 本ずつ塗ったマクロを「系列の色が既定」で外していた）
                n = len(a.get('values') or [])
                ea = a.get('point_colors') or [a.get('color')] * n
                gb = b.get('point_colors') or [b.get('color')] * len(b.get('values') or [])
                if ea != gb:
                    out.append(f"{sl}: 棒の色 期待 {ea[:8]} ／ マクロ後 {gb[:8]}")
            elif a.get('color') != b.get('color') and 'line_color' not in a:          # 折れ線は線の色で比べる
                out.append(f"{sl}: 塗りの色 期待 {a.get('color')!r} ／ マクロ後 {b.get('color')!r}")
    ep, gp = expect.get('pivots') or [], got.get('pivots') or []
    if ep and all(p.get('at') for p in ep) and all(p.get('at') for p in gp):
        key = lambda p: (p.get('sheet') or '', list(p.get('at') or []))      # noqa: E731  置いた場所の順（位置そのものは比べない）
        ep, gp = sorted(ep, key=key), sorted(gp, key=key)
    if len(ep) != len(gp):
        # 2026-09-17 深夜: AI が Dictionary で数えた集計表をセルに書いてピボットの代わりにし、2 往復とも「0 個」で止まった
        out.append(f"ピボットの数: 期待 {len(ep)} ／ マクロ後 {len(gp)}"
                   + ("（ピボットテーブルを作っていない。セルに数えた表を書くのでなく、PivotCaches.Create → CreatePivotTable で"
                      "本物のピボットを作る）" if len(gp) < len(ep) else "（前のピボットが残っている）"))
    for k, (e, g) in enumerate(zip(ep, gp), 1):
        lab = f"ピボット {k}"
        if e.get('sheet') != g.get('sheet'):
            out.append(f"{lab}: 置いたシート 期待「{e.get('sheet')}」／ マクロ後「{g.get('sheet')}」")
        for key, name in (('rows', '行のフィールド'), ('cols', '列のフィールド'), ('pages', 'フィルタのフィールド')):
            _cmp_list(f"{lab}: {name}", e.get(key), g.get(key), out)
        # 値のフィールドは期待が持つ鍵だけで比べる（古い台帳の正解には calc が無い）
        if [{k: x.get(k) for k in x} for x in e.get('data') or []] != [{k: y.get(k) for k in x} for x, y in zip(e.get('data') or [], g.get('data') or [])] \
                or len(e.get('data') or []) != len(g.get('data') or []):
            out.append(f"{lab}: 値のフィールド 期待 {e.get('data')} ／ マクロ後 {g.get('data')}")
        for key, name in (('row_grand', '行の総計'), ('col_grand', '列の総計')):
            if e.get(key) != g.get(key):
                out.append(f"{lab}: {name} 期待 {e.get(key)} ／ マクロ後 {g.get(key)}")
        if 'style' in e and e.get('style') != g.get('style'):
            out.append(f"{lab}: スタイル（TableStyle2） 期待 {e.get('style')!r} ／ マクロ後 {g.get('style')!r}")
        if e.get('row_order') and g.get('row_order') and e.get('row_order') != g.get('row_order') \
                and sorted(e.get('row_order')) == sorted(g.get('row_order')):
            out.append(f"{lab}: 行の並び 期待 {e.get('row_order')[:8]} ／ マクロ後 {g.get('row_order')[:8]}"
                       "（PivotField.AutoSort xlDescending, \"値のフィールドの名前\" で値の大きい順）")
        if e.get('source') and e.get('source') != g.get('source'):
            out.append(f"{lab}: 元の範囲 期待 {e.get('source')!r} ／ マクロ後 {g.get('source')!r}"
                       "（pt.ChangePivotCache wb.PivotCaches.Create(xlDatabase, \"'シート'!R1C1:R80C5\") → pt.RefreshTable）")
        if 'null' in e and e.get('null') != g.get('null'):
            out.append(f"{lab}: 空白のセルに表示する文字（NullString・DisplayNullString） 期待 {e.get('null')!r} ／ マクロ後 {g.get('null')!r}")
        erf, grf = e.get('row_fields') or [], {r.get('field'): r for r in g.get('row_fields') or []}
        for idx, er in enumerate(erf):
            gr = grf.get(er.get('field'))
            if gr is None:
                continue                                   # 行のフィールドの違いは上で言っている
            fl = f"{lab}: 行のフィールド「{er.get('field')}」"
            if er.get('layout') != gr.get('layout'):
                out.append(f"{fl}のレイアウト 期待 {er.get('layout')} ／ マクロ後 {gr.get('layout')}（RowAxisLayout・LayoutForm）")
            if idx < len(erf) - 1:                          # いちばん内側の行のフィールドの小計・繰り返しは見た目に出ない
                if er.get('repeat') != gr.get('repeat'):
                    out.append(f"{fl}のラベルの繰り返し 期待 {er.get('repeat')} ／ マクロ後 {gr.get('repeat')}（RepeatAllLabels・RepeatLabels）")
                if er.get('subtotal') is not None and er.get('subtotal') != gr.get('subtotal'):
                    out.append(f"{fl}の小計 期待 {'あり' if er.get('subtotal') else 'なし'} ／ マクロ後 {'あり' if gr.get('subtotal') else 'なし'}"
                               "（PivotField.Subtotals）")
        if e.get('numbers') != g.get('numbers'):
            miss = [x for x in e.get('numbers') or [] if x not in (g.get('numbers') or [])]
            out.append(f"{lab}: 出てきた数が違う（期待にあってマクロ後に無い数 {miss[:6]}）")
        if e.get('labels') != g.get('labels'):
            miss = sorted(set(e.get('labels') or []) - set(g.get('labels') or []))
            extra = sorted(set(g.get('labels') or []) - set(e.get('labels') or []))
            out.append(f"{lab}: 出てきた項目が違う（無い {miss[:6]}・余計 {extra[:6]}）"
                       + ("（「(空白)」や「<日付」は、元の範囲に明細でない行＝下の合計の行・空の行が入ったしるし。"
                          "範囲の最後は、見出しの列が埋まっている最後の明細の行にする）"
                          if any(x in ('(空白)', '(blank)') or str(x).startswith('<') for x in extra) else ''))
    et, gt = expect.get('tables'), got.get('tables') or []
    if et is not None:
        if len(et) != len(gt):
            out.append(f"テーブルの数: 期待 {len(et)}（{[t.get('name') for t in et]}）／ マクロ後 {len(gt)}（{[t.get('name') for t in gt]}）")
        for k, (e, g) in enumerate(zip(et, gt), 1):
            lab = f"テーブル {k}「{e.get('name')}」"
            for key, name in (('sheet', 'シート'), ('name', '名前'), ('range', '範囲'), ('style', 'スタイル'),
                              ('totals', '集計行'), ('filter', 'フィルタのボタン'), ('stripes', '縞模様（行）')):
                if e.get(key) != g.get(key):
                    out.append(f"{lab}: {name} 期待 {e.get(key)!r} ／ マクロ後 {g.get(key)!r}")
            if 'hidden' in e and e.get('hidden') != g.get('hidden'):
                eh, gh = e.get('hidden') or [], g.get('hidden') or []
                out.append(f"{lab}: 絞り込みで隠れた行 期待 {len(eh)} 行{eh[:8]} ／ マクロ後 {len(gh)} 行{gh[:8]}"
                           "（ListObject.Range.AutoFilter Field:=列の番号, Criteria1:=条件。前の絞り込みは AutoFilter.ShowAllData で解除）")
            en = [c.get('name') for c in e.get('columns') or []]
            gn = [c.get('name') for c in g.get('columns') or []]
            _cmp_list(f"{lab}: 列の名前", en, gn, out, limit=12)
            gcols = {c.get('name'): c for c in g.get('columns') or []}
            for c in e.get('columns') or []:
                d = gcols.get(c.get('name')) or {}
                ef, gf = c.get('formula') or '', d.get('formula', '') or ''
                # 式の書き方は色々（INDEX/MATCH・VLOOKUP・XLOOKUP）。セルの値は別に比べている＝どちらも構造化参照の計算列なら通す
                # （2026-09-18: マスタから課名を引く計算列のお題）。A1 形式の参照・値の直書き・式が無いのは言う
                if ef != gf and not (ef and gf and '[' in ef and '[' in gf):
                    out.append(f"{lab} 列「{c.get('name')}」: 計算列の式 期待 {c.get('formula')!r} ／ マクロ後 {d.get('formula', '')!r}")
                if c.get('total') != d.get('total', ''):
                    out.append(f"{lab} 列「{c.get('name')}」: 集計行 期待 {c.get('total')!r} ／ マクロ後 {d.get('total', '')!r}")
            if 'query' in e or 'query' in g:
                if e.get('query') != g.get('query'):
                    out.append(f"{lab}: 読み込んだクエリ 期待 {e.get('query')!r} ／ マクロ後 {g.get('query')!r}")
                ev, gv = e.get('values') or [], g.get('values') or []
                if ev != gv:
                    diff = next(((i, a, b) for i, (a, b) in enumerate(zip(ev, gv)) if a != b), None)
                    out.append(f"{lab}: 読み込まれた値が違う（期待 {len(ev)} 行・マクロ後 {len(gv)} 行"
                               + (f"・{diff[0] + 1} 行目 期待 {diff[1][:6]} ／ マクロ後 {diff[2][:6]}" if diff else "") + "）")
    if 'slicers' in expect:
        es = [(s.get('field'), s.get('kind'), s.get('sheet')) for s in expect.get('slicers') or []]
        gs = [(s.get('field'), s.get('kind'), s.get('sheet')) for s in got.get('slicers') or []]
        if es != gs:
            out.append(f"スライサー（絞り込む列・元・置いたシート）: 期待 {es} ／ マクロ後 {gs}"
                       + ("（前のスライサーが残っている＝SlicerCaches から消してから作る）" if len(gs) > len(es) else
                          "（wb.SlicerCaches.Add2(ピボット, \"列名\").Slicers.Add(シート, , 名前, 見出し, 上, 左, 幅, 高さ)）" if len(gs) < len(es) else ""))
    if 'sparklines' in expect:
        esp, gsp = expect.get('sparklines') or [], got.get('sparklines') or []
        if len(esp) != len(gsp):
            out.append(f"スパークラインの数: 期待 {len(esp)} 本 ／ マクロ後 {len(gsp)} 本"
                       + ("（前のスパークラインが残っている＝その列の Range.SparklineGroups.Clear で消してから作る）" if len(gsp) > len(esp) else
                          "（Range(置く列の本文).SparklineGroups.Add(Type:=xlSparkLine, SourceData:=\"元の範囲\") で作る。"
                          "置く範囲と元の範囲は行の数をそろえる＝1 行に 1 本）" if len(gsp) < len(esp) else ""))
        gmap = {(s.get('sheet'), s.get('cell')): s for s in gsp}
        bad = 0
        for e in esp:
            g = gmap.get((e.get('sheet'), e.get('cell')))
            if g is None:
                if bad < 4:
                    out.append(f"スパークライン {e.get('cell')}: 期待 あり（元 {e.get('source')}）／ マクロ後 そのセルに無い")
                bad += 1
                continue
            for key, name in (('source', '元の範囲'), ('type', '種類'), ('high', '最高点の印（Points.Highpoint.Visible）'),
                              ('low', '最低点の印（Points.Lowpoint.Visible）'), ('markers', 'マーカー（Points.Markers.Visible）'),
                              ('color', '線の色（SeriesColor.Color）')):
                if e.get(key) != g.get(key):
                    if bad < 4:
                        out.append(f"スパークライン {e.get('cell')}: {name} 期待 {e.get(key)!r} ／ マクロ後 {g.get(key)!r}")
                    bad += 1
        if bad > 4:
            out.append(f"（スパークラインの外れ ほか {bad - 4} 件）")
    eq, gq = expect.get('queries'), got.get('queries') or []
    if eq is not None:
        en, gn = [q.get('name') for q in eq], [q.get('name') for q in gq]
        if en != gn:
            out.append(f"パワークエリのクエリ: 期待 {en} ／ マクロ後 {gn}（名前・数。前のクエリが残っている／作っていない）")
    return out
