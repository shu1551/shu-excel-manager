# -*- coding: utf-8 -*-
"""vbam_forge.py — vba_manager 分割パート: 鍛える回路（AI がマクロを鍛えて増やす）

2026-09-17。shu「残りの後始末の後に、AI がマクロを鍛えて増やす回路を作りましょう」。
回路: 依頼 → 登録済みのマクロが先撃ち（vbam_prefire・AI なし）→ 残りだけ AI（run_agent の往復）
     → **その残りの型を AI がマクロに書く（ここ）** → 写しで試す（値で突き合わせ・AI なし）
     → 合格したら開いているブック（秀コンボ）の標準モジュール「表の整理」に登録
     → 次からは prefire がそのマクロの「依頼の語」で撃つ＝AI の出番は「新しい型に出会った一度だけ」。

  agent --forge 名前 [--register] [--ai …] [--model …] [--max-turns N]
  agent --forged                    鍛えたマクロの一覧（弾・合格・登録先）
  agent --forge 名前 --register     合格したマクロを開いているブックの「表の整理」に足す（compile まで）

弾（1 本）= 直前の走行の「撃つ前の控え」（_agent_backup の写し）＋「終わりの姿」（いまのシートの値）＋依頼文＋
AI が使った手＋先に撃ったマクロの名前。判断は AI（マクロを書く）だけ。拾う・試す・登録は Python が機械的に回す。
人の値には触らない（試すのは写し・自分の Excel）。

マクロの頭の 2 行が「登録簿」（データはファイル自身＝別の台帳を持たない）:
    Sub 合計行を足す()
        ' 依頼の語: 合計|集計
        ' 扱う: 合計 集計
vbam_prefire はこの 2 行を読んで、依頼に「依頼の語」が当たれば撃ち、「扱う」の語は AI に回す仕事から外す。
"""
import os
import re
import json
import time
import shutil
import contextlib

from vbam_core import SCRIPT_DIR, LAST_PROC_FILE, _col_letter, get_or_start_excel
from vbam_hands import _rows_of, _text_of
from vbam_undo import _load_undo_meta

import vbam_agent as va

_AGENT_FORGE_FILE = os.path.join(SCRIPT_DIR, '_agent_forge.json')   # 鍛えた台帳（名前 → 弾・.bas・合格・登録先）
_AGENT_FORGE_DIR = os.path.join(SCRIPT_DIR, '_agent_forge')          # 弾の写し（撃つ前のブック）と鍛えた .bas
_TIDY_BAS = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '表の整理.bas'))   # 先撃ちの本体（写しで再現するとき読み込む）
_MODULE = '表の整理'           # 登録先のモジュール（prefire が探す所）
_SCRATCH_MODULE = '鍛冶_試し'   # 写しで試すときのモジュール名
_MAX_EXAMPLES = 6              # 列ごとの前→後の例の数（AI に渡す）
_MAX_MISMATCH = 30             # 不一致の報告の上限（AI に渡す）
_MAX_ROWS_SHOWN = 16           # AI に見せる表の行数
_MAX_OTHER_SHEETS = 4          # AI に見せる、同じブックのほかのシートの数（頭 8 行ずつ）
_MAX_TEST_MISMATCH = 12        # 別の表の不一致を AI に返す行数（1 枚あたり）
_READ_OPS = ('read', 'read_sheet', 'eval', 'view', 'inspect', 'list', 'get', 'grep')

_ASK_LINE_RE = re.compile(r"^\s*'\s*依頼の語\s*[:：]\s*(.+?)\s*$", re.M)
_HANDLE_LINE_RE = re.compile(r"^\s*'\s*扱う\s*[:：]\s*(.+?)\s*$", re.M)
_HEAD_LINE_RE = re.compile(r"^\s*'\s*見出し\s*[:：]\s*(.+?)\s*$", re.M)
_SELECT_LINE_RE = re.compile(r"^\s*'\s*選ぶ列\s*[:：]\s*(\d+)\s*(?:[-〜~]\s*(\d+))?\s*$", re.M)
_SHAPE_LINE_RE = re.compile(r"^\s*'\s*形\s*[:：]\s*(.+?)\s*$", re.M)   # 道具が書く（AI には書かせない・2026-09-18）
_SUB_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(\s*\)', re.M | re.I)
_FUNC_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Function\s+', re.M | re.I)
_ARGSUB_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+[^\s\(]+\s*\(\s*[^\)\s]', re.M | re.I)
_END_SUB_RE = re.compile(r'^\s*End\s+Sub\s*$', re.M | re.I)
_END_ANY_RE = re.compile(r'^\s*End\s+(?:Sub|Function)\s*$', re.M | re.I)
_DEF_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+\S', re.M | re.I)

FORGE_RULES = """あなたは Excel VBA のマクロを書く係です。AI（あなた）が会話で 1 回やった表の直しを、次からは AI なしで
同じ形の表に当てられるよう、マクロ 1 本にします。道具（Python）があなたの書いたマクロを撃つ前のブックの写しに
読み込んで実行し、終わった表を「AI が直した後の表」とセルの値で突き合わせます。合わなければ不一致を見せるので直してください。

返事は VBA のコードだけ（前後の文・説明・コードフェンスは付けない）。形:
Sub マクロ名()
    ' 依頼の語: 合計|集計            ← この依頼のような言葉（正規表現・| 区切り）。次からこの語が依頼にあれば道具がこのマクロを撃つ
    ' 扱う: 合計 集計                ← このマクロが片づける仕事の語（空白区切り）。ここにある仕事は AI に回さない
    ' 見出し: なし                   ← 見出しの語に頼らないマクロは「なし」（ふつうはこれ）。表の形で列を探すので、
                                     どんな見出しの表にも撃てる。特定の語が要る仕事だけ語を空白区切りで書く
    ' 選ぶ列: 3                      ← 人が選んでいる列（Selection）を使う仕事だけ書く（使う列の数。2-3 のように幅でもよい）。
                                     依頼に「選んでいる列」と書いてある仕事はこの行を書き、列は Selection.Areas の順（選んだ順）に使う。
                                     入口では、道具が依頼の語と見出しが一致する列を選んでから撃つ
    （' 形: の行は道具が書く。表の形＝数の列・日付の列・テーブルなどを道具が表から導いて足すので、あなたは書かなくてよい）
  依頼の語は、この仕事を名指しする具体的な語（決算統計・突合・集約など）を入れる。「集計」「まとめ」のような
  短い一般の語だけに頼らない（ほかの仕事の依頼にも出る。道具は長い語で当たった仕事を優先する）。
  語は依頼文・言い換えの文にそのまま含まれる文字列を | で並べるだけ（.* や ( ) は使わない。「帳票.*一覧」ではなく「帳票|一覧に直」）
    ' 何をするか 1〜3 行
    Dim …
    …
End Sub

規則（守らないと道具が受け付けない）:
 - Sub 1 本だけ。引数なし。Function は作らない。Option Explicit は書かない。Attribute 行は書かない。
   何度も使う処理（文字の正規化など）は Sub の中の GoSub で書く:
       s = CStr(ws.Cells(r, c).Value): GoSub 正規化: key = s
       …
       Exit Sub
   正規化:
       s = Trim$(StrConv(s, vbNarrow)): s = Replace(s, ",", "")
       Return
 - マクロ名は日本語の動詞止め（例: 合計行を足す・表を棒グラフにする）。「表を整える」「重複行を消す」は既にあるので使わない。
   マクロ名にもお題の表の語（課別・支出・職員名簿など）を入れない（どの表にも使うマクロの名前にする）。
 - 対象は ActiveSheet の表。表は「見出し 1 行＋本文」（上に表題の行があることもある）。見出しの行と列の範囲は UsedRange から
   自分で探す（番地を直書きしない）。
   **見出しの行＝値（数・日付）が並ぶ本文のすぐ上の行**。表題・「単位：円」・「出力日 …」・「合計 / …」「列ラベル」の行は
   見出しではない（値が 2 つ以上ある最初の行、では表題の行を見出しにしてしまうことがある）。相手のシート・対応表の
   見出しの行も同じ決め方で探す。
 - **マクロは職場のどんな表にも撃たれる。見出しの語・シート名・テーブル名・課や科目や人の名前を書き込まない。**
   道具は見出しの語とシート名を変えた別の表（課名→所属・部署名、支出額→金額・執行額、支出明細→明細・Sheet1 など）でも試し、
   表に出てくる語をマクロに書き込んでいたら受け付けない。列は語でなく形で決める: 文字の列・数の列・日付の列・
   率の列（表示形式に % がある）・左端／右端・見出しの並び。どの列を使うかは依頼に書いた決まりに従う。
   合計の行・列は「合計・計・小計・総計」の語で見分けてよい（どの表にも出る一般の語）。依頼文に書いてある語も使ってよい。
   別のシート・別のテーブルは形で探す（見出しの並びが同じ・同じ見出しの列がある）。行数・件数も直書きしない（表から数える）。
 - ブックは ActiveWorkbook（または ActiveSheet.Parent）。ThisWorkbook は使わない（マクロはアドインに入る＝ThisWorkbook は
   表のブックではない）。CreateObject は "Scripting.Dictionary" だけ（System.Collections.* は職場の PC で動かない）。
 - オブジェクト（ブック・シート・範囲・グラフ・ピボット）を変数に入れるときは必ず Set を付ける（Set ws = ActiveSheet）。
 - 成功しても MsgBox を出さない。失敗しても止まらない（On Error Resume Next は乱用しない・要る所だけ）。
 - 数式で出せる値（合計・平均など）は数式で書く（=SUM(...)）。値を直接書くのは、書き方をそろえる置き換えのときだけ。
 - 「先に撃ったマクロ」がやった分（書き方の整理・重複の削除）はやり直さない。あなたのマクロは「AI が使った手」の分だけ。
 - 書き方の置き換え（find_replace の 旧→新）は、その語だけでなく同じ種類の語も片づける（Active→進行中 なら Closed→完了 も）。
   ただし列に既にある日本語の語へ寄せる。表に無い語を新しく発明しない。
 - 罫線・列幅・書式は触らない（道具の tidy が後で当てる）。これはセルの書式のこと。
 - グラフ・ピボット・テーブル・パワークエリの仕事では、本物の Excel のオブジェクトを作る（セルに数えた表を書いて代わりにしない）:
   グラフ＝ws.ChartObjects.Add（系列は SeriesCollection.NewSeries で名前・値・項目を範囲で渡す）／
   ピボット＝ActiveWorkbook.PivotCaches.Create(1, 範囲の文字) → .CreatePivotTable(置き場所, 名前) → PivotFields(…).Orientation・AddDataField／
   テーブル＝ws.ListObjects.Add(1, 範囲, , 1)／パワークエリ＝ActiveWorkbook.Queries.Add(名前, M の式) →
   シートの ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=名前", Destination:=…)
   → .QueryTable.CommandType = 2・.CommandText = "SELECT * FROM [名前]"・.Refresh False。
   道具はセルの値に加えて、それらの姿（グラフの種類・タイトル・凡例・系列の名前と値と項目・ラベル・色・フォント／
   ピボットの行・列・値のフィールドと集計／テーブルの名前・範囲・スタイル・計算列・集計行／クエリの名前と読み込んだ値）も正解と突き合わせる。
   依頼に書いてあるデザイン（色・フォント・目盛線・データラベル・凡例の位置）はグラフに当てる。
 - 人の値のあるセルに書かない。足す列・表は表の右の空いた所へ。同じ見出しの列・表が既にあればそこを使う
   （同じ表に 2 回撃っても結果が変わらないように）。
"""


# ----------------------------------------------------------------
# 台帳（純 Python）
# ----------------------------------------------------------------

def _forge_load():
    try:
        with open(_AGENT_FORGE_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _forge_save(d):
    try:
        with open(_AGENT_FORGE_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        return True
    except OSError as ex:
        print(f"エラー: 鍛えた台帳を書けませんでした: {ex}")
        return False


def _safe_name(name):
    return re.sub(r'[\\/:*?"<>|\s]+', '_', str(name).strip())[:60] or 'forge'


# ----------------------------------------------------------------
# 拾う（直前の走行 → 弾）
# ----------------------------------------------------------------

def _hands_of_log(path):
    """記録（_last_agent_log.jsonl）から AI が使った手（読む手は除く）を短い行にする。(手の行, 先に撃ったマクロ名)。"""
    hands, prefire = [], []
    try:
        with open(path, encoding='utf-8') as f:
            lines = [json.loads(l) for l in f if l.strip()]
    except (OSError, ValueError):
        return hands, prefire
    for d in lines:
        if not isinstance(d, dict) or 'reply' not in d:
            continue
        if not prefire:
            m = re.search(r'道具が先に撃ったマクロ[^\n]*\n(.+?) を実行し', str(d.get('prompt') or ''))
            if m:
                prefire = [s.strip() for s in m.group(1).split('→') if s.strip()]
        r = str(d.get('reply') or '')
        try:
            j = json.loads(r[r.find('{'):r.rfind('}') + 1])
        except ValueError:
            continue
        for a in (j.get('actions') or []):
            if not isinstance(a, dict):
                continue
            op = str(a.get('op') or '')
            if not op or op in _READ_OPS:
                continue
            keys = [k for k in a if k not in ('op', 'overwrite', 'say')]
            body = ' '.join(f"{k}={_short(a[k])}" for k in keys)
            hands.append(f"{op} {body}".strip())
    return hands, prefire


def _short(v, n=90):
    s = json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else str(v)
    return s if len(s) <= n else s[:n] + '…'


def _texts(values):
    return [[_text_of(v) for v in row] for row in (values or [])]


def _addr_of(r0, c0, i, j):
    return f"{_col_letter(c0 + j)}{r0 + i}"


def _diff_summary(before, after):
    """撃つ前と後（どちらも {'row','col','values'(文字)}）の差を、AI に見せる文にする。"""
    bv, av = before.get('values') or [], after.get('values') or []
    r0, c0 = int(after.get('row') or 1), int(after.get('col') or 1)
    lines = []
    bw, aw = (len(bv[0]) if bv else 0), (len(av[0]) if av else 0)
    if (len(bv), bw) != (len(av), aw):
        why = []
        if len(bv) != len(av):
            why.append("行が増減しています＝行を足す／消す手が入っています")
        if bw != aw:
            why.append("列が増減しています＝列や表を足す／消す手が入っています")
        lines.append(f"表の大きさ: 前 {len(bv)} 行×{bw} 列 → 後 {len(av)} 行×{aw} 列（{'・'.join(why)}）")
    header = av[0] if av else []
    by_col = {}
    nrow = min(len(bv), len(av))
    for i in range(nrow):
        ncol = min(len(bv[i]), len(av[i]))
        for j in range(ncol):
            if (bv[i][j] or '') != (av[i][j] or ''):
                by_col.setdefault(j, []).append((_addr_of(r0, c0, i, j), bv[i][j], av[i][j]))
        for j in range(ncol, len(av[i])):          # 後で増えた列
            if av[i][j] not in (None, ''):
                by_col.setdefault(j, []).append((_addr_of(r0, c0, i, j), '', av[i][j]))
    for i in range(nrow, len(av)):               # 後で増えた行
        for j, v in enumerate(av[i]):
            if v not in (None, ''):
                by_col.setdefault(j, []).append((_addr_of(r0, c0, i, j), '', v))
    total = sum(len(v) for v in by_col.values())
    lines.append(f"変わったセル: {total} 個（列ごとの数と例）")
    for j in sorted(by_col):
        ex = by_col[j]
        hd = header[j] if j < len(header) else _col_letter(c0 + j)
        lines.append(f"  列 {_col_letter(c0 + j)}「{hd}」: {len(ex)} 個  "
                     + " ／ ".join(f"{a}: '{b}'→'{c}'" for a, b, c in ex[:_MAX_EXAMPLES]))
    return "\n".join(lines), total


def _table_text(snap, max_rows=_MAX_ROWS_SHOWN):
    """表の頭（見出し＋本文の先頭）を TSV で（AI が見出しの語と列の並びを知るため）。"""
    vals = snap.get('values') or []
    r0, c0 = int(snap.get('row') or 1), int(snap.get('col') or 1)
    out = [f"（左上 {_addr_of(r0, c0, 0, 0)}・{len(vals)} 行×{len(vals[0]) if vals else 0} 列）"]
    for i, row in enumerate(vals[:max_rows]):
        out.append(f"{r0 + i}\t" + "\t".join('' if v is None else str(v) for v in row))
    if len(vals) > max_rows:
        out.append(f"…（残り {len(vals) - max_rows} 行）")
    return "\n".join(out)


def _kind_of(v):
    """セルの値の型の 1 文字（純 Python）: n＝数・d＝日付・s＝文字・b＝真偽・-＝空。"""
    from vbam_hands import _is_date_value
    if v is None or v == '':
        return '-'
    if isinstance(v, bool):
        return 'b'
    if _is_date_value(v):
        return 'd'
    if isinstance(v, (int, float)):
        return 'n'
    return 's'


_MAX_BOOK_SHEETS = 40         # ほかのシートを比べる数の上限（1 件 1 枚の帳票で 30 枚ほど）
_MAX_BOOK_ROWS = 300          # ほかのシート 1 枚で比べる行の上限


def _book_of(wb, mine):
    """ブックのシートの並びと、ほかのシートの値（2026-09-18 第二期: 帳票の仕事＝1 件 1 枚・課ごとに配る・様式に転記は
    その仕事のシート以外に書く。それまでの目はその仕事のシート 1 枚の値しか比べず、作らなくても合格した）。"""
    names = [str(sh.Name) for sh in wb.Worksheets]
    sheets = {}
    for sh in list(wb.Worksheets)[:_MAX_BOOK_SHEETS + 1]:
        nm = str(sh.Name)
        if nm == mine:
            continue
        ur = sh.UsedRange
        rows = (_rows_of(ur.Value) or [])[:_MAX_BOOK_ROWS]
        sheets[nm] = {'row': int(ur.Row), 'col': int(ur.Column), 'values': _texts(rows)}
    return {'names': names, 'sheets': sheets}


_PAGE_DEFAULT = {'orientation': '縦', 'zoom': 100, 'fit_wide': None, 'fit_tall': None, 'title_rows': '', 'print_area': '',
                 'center_h': False, 'footer_c': '', 'margins': (50, 50, 54, 54)}
_PAGE_HINT = {'orientation': 'PageSetup.Orientation = xlLandscape（横）', 'zoom': 'PageSetup.Zoom = False にしてから FitToPagesWide/Tall',
              'fit_wide': 'PageSetup.FitToPagesWide = 1', 'fit_tall': 'PageSetup.FitToPagesTall = False（縦は何ページでも）',
              'title_rows': 'PageSetup.PrintTitleRows = "$3:$3"（見出しの行）', 'print_area': 'PageSetup.PrintArea = "$A$3:$F$40"',
              'center_h': 'PageSetup.CenterHorizontally = True', 'footer_c': 'PageSetup.CenterFooter = "&P / &N"',
              'margins': 'PageSetup.LeftMargin 等は Application.InchesToPoints(0.25) のように点で'}


def _page_of(ws, always=False):
    """シートの印刷設定（2026-09-18 第二期: 印刷設定の仕事）。既定のままなら None（読むのは正解の側と、正解に印刷設定があるときだけ）。
    余白だけ違うのは既定と見なす（openpyxl で書いたブックは余白が Excel の既定と違う＝全部の仕事で比べることになる）。
    always＝既定でも全部返す（マクロの後の姿を比べる側）。"""
    ps = ws.PageSetup

    def g(fn, default=None):
        try:
            return fn()
        except Exception:
            return default
    zoom = g(lambda: ps.Zoom)
    fw, ft = g(lambda: ps.FitToPagesWide), g(lambda: ps.FitToPagesTall)
    page = {'orientation': '横' if g(lambda: int(ps.Orientation)) == 2 else '縦',
            'zoom': int(zoom) if isinstance(zoom, (int, float)) and not isinstance(zoom, bool) else None,
            'fit_wide': (int(fw) if isinstance(fw, (int, float)) and not isinstance(fw, bool) and fw else None) if zoom is False else None,
            'fit_tall': (int(ft) if isinstance(ft, (int, float)) and not isinstance(ft, bool) and ft else None) if zoom is False else None,
            'title_rows': str(g(lambda: ps.PrintTitleRows, '') or '').replace('$', ''),
            'print_area': str(g(lambda: ps.PrintArea, '') or '').replace('$', ''),
            'center_h': bool(g(lambda: ps.CenterHorizontally, False)),
            'footer_c': str(g(lambda: ps.CenterFooter, '') or ''),
            'margins': tuple(int(round(float(g(lambda k=k: getattr(ps, k), 0) or 0))) for k in
                             ('LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin'))}
    core = {k: v for k, v in page.items() if k != 'margins'}
    if not always and core == {k: v for k, v in _PAGE_DEFAULT.items() if k != 'margins'}:
        return None
    return page


def _compare_page(e, g):
    """印刷設定を比べる（純 Python）→ 不一致の行。"""
    g = g or dict(_PAGE_DEFAULT)
    names = {'orientation': '向き', 'zoom': '拡大縮小（Zoom）', 'fit_wide': '横のページ数', 'fit_tall': '縦のページ数',
             'title_rows': '見出しの繰り返し（行）', 'print_area': '印刷範囲', 'center_h': '水平の中央', 'footer_c': 'フッターの中央',
             'margins': '余白（左・右・上・下の点）'}
    return [f"印刷設定の{names[k]}: 期待 {e.get(k)!r} ／ マクロ後 {g.get(k)!r}（{_PAGE_HINT[k]}）"
            for k in names if k in e and e.get(k) != g.get(k)]


def _snapshot_texts(ws, book=True):
    """シートの値を文字で（{'row','col','values','kinds'}）。kinds は行ごとの型の文字列（_kind_of）。
    2026-09-17 夜: 値の文字だけで比べていたので、申請額・受付日を NumberFormat "@" の文字で書いた帳票のマクロが合格した
    （見た目は同じ・集計できない一覧）。数・日付・文字の違いも比べる。book＝ほかのシートの並びと値も読む（_book_of）。"""
    ur = ws.UsedRange
    rows = _rows_of(ur.Value)
    snap = {'row': int(ur.Row), 'col': int(ur.Column), 'values': _texts(rows),
            'kinds': [''.join(_kind_of(v) for v in row) for row in (rows or [])]}
    # グラフ・ピボット・テーブル（2026-09-17 深夜: 値しか比べず、グラフやピボットを作るマクロを鍛えられなかった）
    import vbam_objects as vo
    with contextlib.suppress(Exception):
        wb = ws.Parent
        snap['objects'] = vo.snapshot_objects(wb) if vo.has_objects(wb) else {'charts': [], 'pivots': [], 'tables': []}
    if book:
        with contextlib.suppress(Exception):
            snap['book'] = _book_of(ws.Parent, str(ws.Name))
        # 式のセル（2026-09-18 第二期: 値しか比べず、式が値に化けても・値が式に変わっても合格した＝「式を値に」「式の途切れを直す」を鍛えられない）
        with contextlib.suppress(Exception):
            frows = _rows_of(ur.Formula) or []
            snap['fx'] = [_addr_of(int(ur.Row), int(ur.Column), i, j) for i, row in enumerate(frows) for j, v in enumerate(row or [])
                          if isinstance(v, str) and v.startswith('=')]
        # オートフィルタ（シートの絞り込みのボタンの範囲。テーブルの絞り込みは objects の tables が持つ）
        with contextlib.suppress(Exception):
            snap['filter'] = str(ws.AutoFilter.Range.Address).replace('$', '') if ws.AutoFilterMode else ''
    return snap


_SHAPE_CACHE = {}              # (パス, 更新時刻, シート名) → 表の形の語の集合（同じブックは 1 回だけ読む）


def _objects_of_file(path, sheet=None):
    """xlsx の中身（シートの関係ファイル）から、そのシートにピボット・グラフがあるか → (ピボット, グラフ)（純 Python・Excel を起こさない）。
    openpyxl は既存のグラフを読み込まない＝zip の中の関係ファイルを直接見る。sheet が無ければ先頭のシート。読めなければ (False, False)。"""
    import posixpath
    import zipfile
    try:
        with zipfile.ZipFile(str(path)) as z:
            names = set(z.namelist())

            def rels(part):
                d, f = posixpath.split(part)
                rp = posixpath.join(d, '_rels', f + '.rels')
                if rp not in names:
                    return []
                txt = z.read(rp).decode('utf-8', 'replace')
                out = []
                for m in re.finditer(r'<Relationship\b[^>]*>', txt):
                    tag = m.group(0)
                    ty = re.search(r'Type="([^"]+)"', tag)
                    tg = re.search(r'Target="([^"]+)"', tag)
                    ri = re.search(r'Id="([^"]+)"', tag)
                    if ty and tg:
                        target = tg.group(1)
                        full = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(d, target))
                        out.append((ri.group(1) if ri else '', ty.group(1).rsplit('/', 1)[-1], full))
                return out
            wbx = z.read('xl/workbook.xml').decode('utf-8', 'replace')
            sheets = re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]+)"', wbx)
            if not sheets:
                return False, False
            import html
            want = next((rid for nm, rid in sheets if sheet and html.unescape(nm) == str(sheet)), sheets[0][1])
            part = next((full for rid, _ty, full in rels('xl/workbook.xml') if rid == want), None)
            if not part:
                return False, False
            srels = rels(part)
            has_pivot = any(ty == 'pivotTable' for _r, ty, _f in srels)
            has_chart = any(ty == 'chart' for _r, ty, _f in
                            (x for _r2, ty2, f2 in srels if ty2 == 'drawing' for x in rels(f2)))
            return has_pivot, has_chart
    except Exception:
        return False, False


def shape_of_file(path, sheet=None, rows=60):
    """ブックのファイル → 表の形の語の集合（openpyxl・純 Python・Excel を起こさない）。
    入口（COM の vbam_prefire._shape_of_sheet）と同じ物差し（shape_signals）で、鍛えるときに表の形を導く。
    読めなければ空（＝形の行は「なし」になる＝どの表にも撃つ・安全側）。"""
    import vbam_prefire as vp
    try:
        key = (os.path.abspath(str(path)), os.path.getmtime(str(path)), str(sheet or ''))
    except OSError:
        return set()
    if key in _SHAPE_CACHE:
        return _SHAPE_CACHE[key]
    out = set()
    try:
        from openpyxl import load_workbook
        wb = load_workbook(str(path), data_only=True)
        try:
            ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
            r0, c0 = int(ws.min_row or 1), int(ws.min_column or 1)
            raw = [list(r) for r in ws.iter_rows(min_row=r0, max_row=min(int(ws.max_row or r0), r0 + rows - 1),
                                                 min_col=c0, max_col=int(ws.max_column or c0), values_only=True)]
            kinds = [''.join(_kind_of(v) for v in row) for row in raw]
            h = vp._header_row_index(kinds)
            # 結合の見出し＝見出しの行〜+2 行に結合がある（表題の行だけの結合は数えない）
            merged = any(int(m.min_row) <= r0 + h + 2 and int(m.max_row) >= r0 + h for m in ws.merged_cells.ranges)
            others = []
            for other in wb.worksheets[:12]:
                if other is ws:
                    continue
                o0, oc = int(other.min_row or 1), int(other.min_column or 1)
                rows2 = [list(r) for r in other.iter_rows(min_row=o0, max_row=min(int(other.max_row or o0),
                                                                                 o0 + vp._SHEET_WORD_ROWS - 1),
                                                          min_col=oc, max_col=int(other.max_column or oc),
                                                          values_only=True)]
                hi = vp._header_row_index([''.join(_kind_of(v) for v in row) for row in rows2])
                if hi < len(rows2):
                    others.append([v for v in rows2[hi] if isinstance(v, str)])
            has_pivot, has_chart = _objects_of_file(path, ws.title)
            out = vp.shape_signals(_texts(raw), kinds, bool(getattr(ws, 'tables', None)), others, merged, has_pivot, has_chart)
        finally:
            with contextlib.suppress(Exception):
                wb.close()
    except Exception:
        out = set()
    _SHAPE_CACHE[key] = out
    return out


def harvest(name, target_file=None, truth=None):
    """直前の走行を弾にする。控え（撃つ前のブック）を写し、いまのシートの値を「正解」に、依頼文と AI の手を添える。
    truth（人が直した正解のブック）を渡すと、正解はそちらのシートの値にする（AI が間違えた姿を正解にしない）。"""
    meta = (va._load_resume()[1] or {}) if os.path.isfile(va._LAST_AGENT_LOG_FILE) else {}
    if not meta:
        print(f"エラー: 直前の走行の記録がありません（{va._LAST_AGENT_LOG_FILE}）。先に agent で 1 回、表を直してください")
        return None
    if meta.get('mode') not in ('sheet', 'both'):
        print(f"エラー: 鍛えられるのはシートの走行だけです（前回は {meta.get('mode') or '不明'}）")
        return None
    sheet, request = meta.get('sheet'), meta.get('request')
    undo = _load_undo_meta() or {}
    if not undo.get('path') or not os.path.isfile(str(undo['path'])):
        print("エラー: 直前の走行の控え（撃つ前のブック）がありません（agent --backups で確かめてください）")
        return None
    if undo.get('book') and undo.get('book') != meta.get('book'):
        print(f"エラー: 控えは「{undo.get('book')}」・記録は「{meta.get('book')}」＝別の走行のものです")
        return None
    xl, wb = va.get_workbook(target_file)
    if str(wb.Name) != meta.get('book'):
        print(f"エラー: 直前の走行は「{meta.get('book')}」のものです（いま開いているのは「{wb.Name}」）。"
              "そのブックを開いてアクティブにしてから、もう一度 --forge してください")
        return None
    hands, prefire = _hands_of_log(va._LAST_AGENT_LOG_FILE)
    if not hands:
        print("エラー: 記録に AI の手がありません（マクロだけで合格した走行は、鍛える残りがありません）")
        return None
    os.makedirs(_AGENT_FORGE_DIR, exist_ok=True)
    ext = os.path.splitext(str(undo['path']))[1] or '.xlsx'
    before_copy = os.path.join(_AGENT_FORGE_DIR, _safe_name(name) + '_before' + ext)
    shutil.copyfile(str(undo['path']), before_copy)
    expect = _snapshot_texts(wb.Sheets(sheet))
    truth_copy = None
    if truth:
        truth_copy = _copy_in(truth, _safe_name(name) + '_truth')
        if not truth_copy:
            return None
    # 撃つ前の姿は控えの覚書に無いので、写しを自分の Excel で読む（人の Excel に開かない）
    snaps = _read_sheets([(before_copy, sheet)] + ([(truth_copy, sheet)] if truth_copy else []))
    if snaps is None:
        return None
    before = snaps[0]
    if truth_copy:
        expect = snaps[1]
    case = {'name': str(name), 'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'from_book': str(wb.Name),
            'sheet': sheet, 'request': request or '', 'before': before_copy, 'expect': expect,
            'before_values': before, 'hands': hands, 'prefire': prefire, 'run_id': undo.get('run_id'),
            'truth': truth_copy, 'tests': [],
            'bas': None, 'sub': None, 'passed': False, 'mismatch': None, 'turns': 0, 'registered_to': None}
    d = _forge_load()
    if name in d:
        print(f"（同じ名前の弾を置き換えます: {name}）")
    d[str(name)] = case
    _forge_save(d)
    print(f"弾にしました: 「{name}」　元: {wb.Name}!{sheet}　AI の手 {len(hands)} 個"
          + (f"　先に撃ったマクロ: {' → '.join(prefire)}" if prefire else "")
          + (f"　正解: {os.path.basename(truth)}（人の作った表）" if truth_copy else ""))
    return case


def _prefire_names(request):
    """依頼文 → 写しで先に撃つ既定のマクロ（本番の先撃ちと同じ語の決まり・登録簿は見ない）。"""
    import vbam_prefire as vp
    p = vp.plan_full(request)
    if not p['tidy']:
        return []
    return [vp._MACRO_TIDY] + ([vp._MACRO_DEDUPE] if p['dedupe'] else [])


def _copy_in(path, stem):
    """弾の置き場へ写す（原本は触らない）。無ければ None。"""
    if not path or not os.path.isfile(str(path)):
        print(f"エラー: ブックがありません: {path}")
        return None
    os.makedirs(_AGENT_FORGE_DIR, exist_ok=True)
    dst = os.path.join(_AGENT_FORGE_DIR, stem + (os.path.splitext(str(path))[1] or '.xlsx'))
    if os.path.abspath(str(path)) != os.path.abspath(dst):
        shutil.copyfile(str(path), dst)
    return dst


def harvest_pair(name, before_path, truth_path, sheet=None, request='', tests=()):
    """人が用意した「直す前のブック」と「直した正解のブック」から弾を作る（agent の走行は要らない・2026-09-17）。
    職場の作業は、汚れた表と人が直した表の 2 冊があれば鍛えられる。AI の手の記録は無い＝差がすべて。"""
    if not str(request or '').strip():
        print("エラー: 依頼文が要ります（次からこの語の依頼でマクロが撃たれます）。例: agent --forge 名前 --before 前.xlsx --truth 正解.xlsx \"依頼文\"")
        return None
    stem = _safe_name(name)
    before_copy = _copy_in(before_path, stem + '_before')
    truth_copy = _copy_in(truth_path, stem + '_truth')
    if not before_copy or not truth_copy:
        return None
    snaps = _read_sheets([(before_copy, sheet), (truth_copy, sheet)])
    if snaps is None:
        return None
    before, expect = snaps
    sheet = before.get('sheet') or sheet
    case = {'name': str(name), 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
            'from_book': os.path.basename(str(before_path)), 'sheet': sheet, 'request': str(request),
            'before': before_copy, 'expect': expect, 'before_values': before, 'hands': [],
            'prefire': _prefire_names(request), 'run_id': None, 'truth': truth_copy, 'tests': [],
            'sel': _sel_of(before_path),
            'bas': None, 'sub': None, 'passed': False, 'mismatch': None, 'turns': 0, 'registered_to': None}
    d = _forge_load()
    if name in d:
        print(f"（同じ名前の弾を置き換えます: {name}）")
        # 前の版のマクロは作り直しの出発点（2026-09-18: 語を書き込んだ 27 本を、手順は活かして表の形に作り直す）
        ref = (d[name] or {}).get('ref_code') or (d[name] or {}).get('code')
        if ref:
            case['ref_code'] = ref
    d[str(name)] = case
    _forge_save(d)
    print(f"弾にしました: 「{name}」　直す前:{os.path.basename(str(before_path))}!{sheet}　正解: {os.path.basename(str(truth_path))}"
          + (f"　先に撃つマクロ: {' → '.join(case['prefire'])}" if case['prefire'] else ""))
    return case


def _sel_of(path):
    """お題の「選んでいる列」の覚書（<…>_選択.txt に B1,D1,F1 のような番地）→ 文字（無ければ None・純 Python）。
    2026-09-18: 列の選び方が要る仕事は「選んでいる列を使う」決まりにした＝写しで試すときも人が選んだ形を再現する。"""
    import glob as _glob
    base = re.sub(r'_(前|正解)(\.\w+)$', '', str(path or ''))
    for cand in (base + '_選択.txt', base + '_選ぶ列.txt'):
        if os.path.isfile(cand):
            with contextlib.suppress(OSError):
                with open(cand, encoding='utf-8-sig') as f:
                    txt = f.read().strip()
                if txt:
                    return txt
    del _glob
    return None


def add_tests(case, paths):
    """別の表で試す組（直す前・正解 の順に 2 冊ずつ）を弾に足す。→ 足した数（形が違えば None）。
    鍛えたマクロが 1 つの表に合わせただけ（行数・番地・科目の数を直書き）でないかを、AI なしで確かめる。"""
    paths = [p for p in (paths or []) if p]
    if len(paths) % 2:
        print("エラー: --test は「直す前 正解」の 2 冊ずつ並べてください")
        return None
    stem = _safe_name(case['name'])
    have = {t.get('label') for t in case.get('tests') or []}
    added = 0
    pairs = []
    for k in range(0, len(paths), 2):
        label = os.path.splitext(os.path.basename(paths[k]))[0]
        if label in have:
            continue
        b = _copy_in(paths[k], f"{stem}_test_{_safe_name(label)}_before")
        t = _copy_in(paths[k + 1], f"{stem}_test_{_safe_name(label)}_truth")
        if not b or not t:
            return None
        pairs.append((label, b, t, _sel_of(paths[k])))        # 選んでいる列の覚書は「元のお題」の名前で探す（写しの名前ではない）
    if pairs:
        # シート名が弾と違う表（「旧システム出力」など）も試せるように、無ければ先頭のシートを読む＝その名前で撃つ
        snaps = _read_sheets([(t, case.get('sheet')) for _l, _b, t, _s in pairs], fallback_first=True)
        if snaps is None:
            return None
        for (label, b, _t, sel), snap in zip(pairs, snaps):
            case.setdefault('tests', []).append({'label': label, 'before': b, 'expect': snap, 'sheet': snap.get('sheet'),
                                                 'sel': sel})
            added += 1
    return added


def _read_sheets(items, fallback_first=False):
    """[(ブックのパス, シート名 or None)] → [snap]（1 つの自分の Excel＝非表示で順に開いて読む）。
    シート名が None なら先頭のシート（snap['sheet'] に名前）。fallback_first なら名前が無いときも先頭。読めなければ None。"""
    xl, pid = _start_own()
    out = []
    try:
        for path, sheet in items:
            wb = None
            try:
                wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                try:
                    ws = wb.Sheets(sheet) if sheet else wb.Worksheets(1)
                except Exception:
                    if not fallback_first:
                        print(f"エラー: {os.path.basename(path)} にシート「{sheet}」がありません")
                        return None
                    ws = wb.Worksheets(1)
                snap = _snapshot_texts(ws)
                snap['sheet'] = str(ws.Name)
                with contextlib.suppress(Exception):
                    pg = _page_of(ws)
                    if pg:
                        snap['page'] = pg                  # 既定のままなら持たない＝比べない（古い仕事の比べ方は変わらない）
                snap['others'] = []
                for other in list(wb.Worksheets)[:_MAX_OTHER_SHEETS + 1]:
                    if str(other.Name) == snap['sheet'] or len(snap['others']) >= _MAX_OTHER_SHEETS:
                        continue
                    with contextlib.suppress(Exception):
                        snap['others'].append({'name': str(other.Name),
                                               'head': _table_text(_snapshot_texts(other, book=False), max_rows=8)})
                out.append(snap)
            finally:
                with contextlib.suppress(Exception):
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        return out
    finally:
        xl = wb = ws = None
        _quit(pid)


def _read_sheet_from_file(path, sheet):
    """ブックの写しを自分の Excel（非表示）で開いて、そのシートの値を文字で読む。"""
    got = _read_sheets([(path, sheet)])
    return got[0] if got else {}


def _start_own():
    """自分の台（非表示の Excel）を起こす。→ (xl, pid)。畳むのは _quit(pid) で、この台だけ。"""
    import vbam_core
    xl = get_or_start_excel(visible=False)
    return xl, vbam_core._created_xl_pid


def _quit(pid):
    """自分の台だけを畳む（ほかに道具が起こした台＝試験や確かめの Excel には触らない）。"""
    import gc
    from vbam_core import release_instance
    gc.collect()
    with contextlib.suppress(Exception):
        release_instance(pid)


# ----------------------------------------------------------------
# 書く（AI 1 回）
# ----------------------------------------------------------------

def _prompt(case, feedback=None, prev_code=None):
    parts = [FORGE_RULES, "\n--- 依頼（AI が会話で受けた文） ---\n" + str(case.get('request') or '')]
    if case.get('phrases'):
        parts.append("\n--- 同じ仕事の言い換え（' 依頼の語: はこれにも全部当たるように） ---\n"
                     + "\n".join("  " + p for p in case['phrases']))
    if case.get('prefire'):
        parts.append("\n--- 先に撃ったマクロ（この分はやり直さない） ---\n" + " → ".join(case['prefire']))
    if case.get('hands'):
        parts.append("\n--- AI が使った手（この分をマクロにする） ---\n" + "\n".join("  " + h for h in case['hands']))
    else:
        parts.append("\n--- AI が使った手 ---\n  （記録はありません。人が直した正解の表から書きます＝下の「前 → 後」の差がすべてです）")
    parts.append("\n--- 撃つ前の表（頭） ---\n" + _table_text(case.get('before_values') or {}))
    diff, _n = _diff_summary(case.get('before_values') or {}, case.get('expect') or {})
    who = "人が直した正解" if case.get('truth') else "AI が直した結果"
    parts.append(f"\n--- 前 → 後（{who}。マクロはこれと同じ値にする） ---\n" + diff)
    parts.append("\n--- 直した後の表（頭） ---\n" + _table_text(case.get('expect') or {}))
    if case.get('tests'):
        parts.append(f"\n--- 別の表 {len(case['tests'])} 枚でも試します（見出しの語・シート名・テーブル名・項目の名前・位置・"
                     "行数・余計な列が違う同じ形の表）。番地・行数・件数・見出しの語・シート名・項目の名前を書き込まず、"
                     "表の形（文字の列・数の列・日付の列・率の列・左端／右端・依頼に書いた列の決まり）で探してください。正解（頭）: ---")
        for t in case['tests']:
            if t.get('expect'):
                parts.append(f"「{t.get('label')}」（シート「{t['expect'].get('sheet') or t.get('sheet') or '?'}」）\n"
                             + _table_text(t['expect'], max_rows=12))
    others = (case.get('before_values') or {}).get('others') or []
    if others:
        parts.append("\n--- 同じブックのほかのシート（頭。シート名は表ごとに違う＝名前でなく形で探す） ---")
        for o in others:
            parts.append(f"シート「{o.get('name')}」\n{o.get('head')}")
    if case.get('ref_code') and not prev_code:
        parts.append("\n--- 前の版（表の語を書き込んでいたので、見出しの語が違う職場の表で動かなかった。オブジェクトの作り方・"
                     "API の手順・罠の避け方は使ってよい。語で探す所・語を書く所は全部、表の形に直す） ---\n" + str(case['ref_code']))
    if feedback:
        parts.append("\n--- 前回のマクロの不一致（直してください） ---\n" + feedback)
        if prev_code:
            parts.append("\n--- 前回のマクロ ---\n" + prev_code)
    return "\n".join(parts)


FORGE_DEFAULT_MODEL = 'opus'      # 2026-09-19 shu「マクロを書く頭の既定は Opus に」。表を直す agent の既定（sonnet）とは別に持つ


def _forge_model(ai, model):
    """鍛える回路でマクロを書く頭。名指しが無く、頭が Claude Code（既定）のときだけ Opus にする。

    マクロは一度合格すれば、次からは AI なしで何度でも回る＝書くときの数十秒より、1 回で正しく書けるほうが効く。
    gemini／claude（API）を名指ししたとき・--model を付けたときは、そのまま。"""
    if model:
        return model
    if not ai or ai == 'claude-code':
        return FORGE_DEFAULT_MODEL
    return model


def _read_answer(path):
    """会話している AI（Gemini・Claude など）が書いた答え（Sub 全文）をファイルから読む。UTF-8 → CP932 の順。"""
    try:
        raw = open(path, 'rb').read()
    except OSError as e:
        return f"（答えのファイルが読めません: {path}・{e}）"
    for enc in ('utf-8-sig', 'cp932'):
        with contextlib.suppress(UnicodeDecodeError):
            return raw.decode(enc).replace('\r\n', '\n')      # AI の返事と同じ \n にそろえる（.bas の改行の二重化を避ける）
    return raw.decode('utf-8', errors='replace').replace('\r\n', '\n')


def _write_prompt(path, text):
    """書き手への問いを UTF-8 で書く（path が無ければ画面に出す）。"""
    if not path:
        print(text)
        return
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def _ask_once(ai, model, prompt):
    if ai == 'file':
        # 書き手が会話している AI のとき（2026-09-19 shu・Gemini「ハーネスを作れば自分で回せる」）: 答えはファイルで受け取る。
        # 問いは forge(prompt_only=True) か、前の往復の終わりに prompt_out へ書いてある。採点は今までどおり道具がする
        return _read_answer(model)
    from vbam_ai import _ai_setup, _ask, _cc_oneshot
    ai, model, key = _ai_setup(ai, _forge_model(ai, model))
    if ai == 'claude-code':
        res = _cc_oneshot(prompt, model)             # (text, usage) の 2 つ組で返る（1 回目は tuple の repr を読んで Sub が無いと出た）
        return res[0] if isinstance(res, tuple) else res
    text, _usage, _sec = _ask(ai, model, key, [('user', prompt)])
    return text


def _log_reply(name, turn, prompt, text):
    """AI に渡した文と生の返事を弾の置き場に残す（規則に合わなかったときに何が返ったか見えるように）。"""
    with contextlib.suppress(OSError):
        os.makedirs(_AGENT_FORGE_DIR, exist_ok=True)
        with open(os.path.join(_AGENT_FORGE_DIR, f"{_safe_name(name)}_turn{turn}.txt"), 'w', encoding='utf-8') as f:
            f.write("=== prompt ===\n" + prompt + "\n\n=== reply ===\n" + str(text or ''))


def _extract_code(text):
    """返事から Sub 〜 End Sub を 1 本取り出す（コードフェンス・前後の文は捨てる）。無ければ None。"""
    s = str(text or '')
    m = re.search(r'```(?:vba|vb|basic)?\s*\n(.*?)```', s, re.S | re.I)
    if m:
        s = m.group(1)
    if not _SUB_RE.search(s) and '{' in s:
        # JSON に包んで返すことがある（{"code":"Sub …\n…"}・2026-09-17 の 2 往復目）＝中の文字を取り出す
        with contextlib.suppress(ValueError, TypeError, AttributeError):
            j = json.loads(s[s.find('{'):s.rfind('}') + 1])
            vals = [j.get('code')] if isinstance(j.get('code'), str) else [v for v in j.values() if isinstance(v, str)]
            s = next((v for v in vals if _SUB_RE.search(v)), s)
    m = _SUB_RE.search(s)
    if not m:
        return None
    e = _END_SUB_RE.search(s, m.end())
    if not e:
        return None
    # End Sub の後ろに Function／Sub が続く返事は、そこまで含めて返す（検査で落とす）。
    # 2026-09-17: 1 本目の End Sub で切っていたので、後ろの Function NormalizeTkt が黙って捨てられ、
    # 規則の検査は素通り・写しではコンパイルエラーの窓で 60 秒止まって「無限ループの疑い」と誤報していた
    end = e.end()
    if _DEF_RE.search(s, end):
        for t in _END_ANY_RE.finditer(s, end):
            end = t.end()
    return s[m.start():end].strip('\r\n') + "\n"


def _normalize_ask_line(code):
    """' 依頼の語: の正規表現の記号を語の区切りに直す（純 Python）。→ (コード, 直した行 or None)。
    2026-09-17 夜: 帳票・月次の鍛錬で AI が「帳票.*一覧」「課別.{0,3}月別」を書き続け、規則違反の往復だけで
    6 往復を使い切った（マクロを 1 回も撃てなかった）。記号で語を分けるのは判断の要らない手順＝道具が直す。"""
    m = _ASK_LINE_RE.search(code or '')
    if not m or not re.search(r'[.*+?()\[\]{}^$\\]', m.group(1)):
        return code, None
    words = []
    for w in re.split(r'[.*+?()\[\]{}^$\\|]+', m.group(1)):
        w = w.strip()
        if len(w) >= 2 and not re.fullmatch(r'[\d,\s]*', w) and w not in words:
            words.append(w)
    if not words:
        return code, None
    line = "|".join(words)
    return code[:m.start(1)] + line + code[m.end(1):], line


def _phrase_words(text):
    """言い換えの文の中の語（漢字・カタカナの 2 文字以上の並び）（純 Python）。依頼の語の候補として AI に見せる。"""
    out = []
    for w in re.findall(r'[一-鿿ァ-ヺー]{2,}', str(text or '')):
        if w not in out:
            out.append(w)
    return out


_PARTICLE_EDGE = tuple('のをにでとはがてもへや、。・')


def _ask_candidates(text, others=(), n=5):
    """言い換えの文に当てる依頼の語の候補（純 Python）: 文の 3〜8 字の部分で、漢字かカタカナを含み、助詞で始まり・終わらず、
    一般的な依頼（_GENERIC_REQUESTS）に当たらないもの。ほかの言い換え（others）にも出る語を先に、短い語を先に。
    2026-09-17 夜: 候補を「漢字・カタカナの並び」だけで出したら「課ごとの月ごとの支出額の表」に「支出額」しか出ず、
    AI が選んだ「支出額の表」は一般の依頼にも当たって 5 往復ぶつかった（「月ごとの支出」なら通る）。"""
    s = str(text or '')
    seen, cands, split = set(), [], []
    for i in range(len(s)):
        for L in range(3, _ASK_KATAKANA_MAX + 1):
            w = s[i:i + L]
            if len(w) < L or w in seen:
                continue
            if L > _ASK_WORD_MAX and not re.fullmatch(r'[ァ-ヺー]+', w):
                continue                                    # 9 字からはカタカナだけの 1 語だけ
            seen.add(w)
            # 語の切れ目: 漢字・カタカナで始まり、漢字・カタカナで終わる（「ごとの支」「出を課」のような切れ端を出さない）
            if not re.match(r'[一-鿿ァ-ヺ]', w) or not re.search(r'[一-鿿ァ-ヺー]$', w):
                continue
            nxt = s[i + L:i + L + 1]
            if re.match(r'[一-鿿ァ-ヺー]', nxt) and bool(re.match(r'[ァ-ヺー]', nxt)) == bool(re.match(r'[ァ-ヺー]', w[-1])):
                continue                                    # 漢字・カタカナの並みの途中で終わらない（「月ごとの支」を出さない）
            # ↑ カタカナと漢字の変わり目は切れ目として認める（2026-09-18 第二期: 「ウォーターフォール図」で候補が 0 になった）
            if re.search(r'[、。，．,「」『』（）()：:；;！!？?]', w):
                continue                                    # 句読点をまたぐ切れ端（「出す。表」「式を、今」）は語にしない（2026-09-18 第二期）
            if any(w in g for g in _GENERIC_REQUESTS):
                continue
            if _clash_jobs(w):
                continue                                    # ほかの鍛えた仕事の依頼にも出る語は候補にしない（出すと衝突で弾かれる）
            # 3 字の語（「行・列」「行は入」）より 4 字以上を先に（2026-09-18: 長い依頼文の決まり文句の切れ端が選ばれた）
            key = (-sum(1 for o in others if w in o), len(w) < 4, len(w), i, w)
            if i and re.match(r'[一-鿿ァ-ヺー]', s[i - 1]):
                # 並みの途中から始まる語は、候補が無いときだけ使う。前も後ろも漢字 2 字以上の切れ目だけ
                # （2026-09-18: 「課別月別のピボットをいつもの形で」に候補が無く、課別月別＝推移折れ線と衝突・「月別のピボット」なら通る）
                head = re.match(r'[一-鿿]+', w)
                before = re.search(r'[一-鿿]+$', s[:i])
                if head and before and len(head.group()) >= 2 and len(before.group()) >= 2:
                    split.append(key)
                continue
            cands.append(key)
    return [w for *_k, w in sorted(cands or split)[:n]]


def _ask_cover(miss, texts):
    """外れた言い換え（miss）を全部カバーする語の組を、候補から欲張りに選ぶ（純 Python）。選べなければ []。
    2026-09-17 深夜: パワークエリの鍛え直しで、AI が候補の先頭の「パワークエリ」を足さず、言い換えごとに別の語を足しては
    別の文を外すのを 8 往復くり返した＝全部に当たる最小の語を道具が決めて言う。"""
    left = list(miss)
    pool = []
    for p in miss:
        for w in _ask_candidates(p, [x for x in texts if x != p], n=12):
            if w not in pool:
                pool.append(w)
    chosen = []
    while left:
        best = max(pool, key=lambda w: (sum(1 for p in left if w in p), -len(w)), default=None)
        if not best or not any(best in p for p in left):
            return []
        chosen.append(best)
        left = [p for p in left if best not in p]
    return chosen


def _auto_fix_ask(case, code, ask_re):
    """頭の行の「依頼の語」「扱う」を道具が直す（純 Python）→ (直したコード, 依頼の語) か (None, None)。
    決まりに反した語（長い・一般の依頼に当たる・ほかの仕事とぶつかる）を外し、残すべき語を戻し、外れた言い換えと依頼文に
    当たる語を足す。「扱う」には AI に回されると言われた語を足す。直した後に _check_fit が通ったときだけ使う。
    2026-09-18 未明: 値は全部そろっているのに、AI が依頼の語の行を毎回まるごと書き換え、足すと落とすを 8 往復くり返した。"""
    import vbam_prefire as vp
    words = [w for w in str(ask_re or '').split('|') if w]
    texts = [str(case.get('request') or '')] + list(case.get('phrases') or [])

    keep = [w for w in str(case.get('keep_ask') or '').split('|') if w]
    ok = lambda w: (_ask_len_ok(w) and not re.search(r'[.*+?()\[\]{}^$\\、。，．,「」（）：:]', w)      # noqa: E731
                    and any(w in t for t in texts if t)                    # 依頼文・言い換えに出ない語は使わない
                    and not any(w in g for g in _GENERIC_REQUESTS) and not _clash_jobs(w, str(case.get('name') or '')))
    words = [w for w in words if ok(w)]
    words += [w for w in keep if ok(w) and w not in words and not any(x in w for x in words)]
    miss = [t for t in texts if t and not any(w in t for w in words) and (t == texts[0] or _ask_candidates(t, texts))]
    if miss:
        cover = _ask_cover(miss, texts)
        if not cover:
            return None, None
        words += [w for w in cover if w not in words]
    new_ask = '|'.join(dict.fromkeys(words))
    new_code = re.sub(r"(^\s*'\s*依頼の語\s*[:：]\s*).*$", lambda m: m.group(1) + new_ask, code, count=1, flags=re.M)
    # 扱う: この依頼（と言い換え）で AI に回されると言われる語を足す
    reg = vp.registry_from_text(new_code)
    wds = vp.sheet_words((case.get('before_values') or {}).get('values'))
    left = []
    for text in texts:
        p = vp.plan_full(text, reg, wds)
        left += [w for w in p['left'] if w not in left]
    if left:
        new_code = re.sub(r"(^\s*'\s*扱う\s*[:：]\s*)(.*)$", lambda m: m.group(1) + ' '.join(dict.fromkeys(m.group(2).split() + left)),
                          new_code, count=1, flags=re.M)
    # 見出し: 別の表に無い語を外す（1 語は残るときだけ）。2026-09-18: 下半期だけの表を足したら「課名 4月」の 4月 だけで AI を 1 往復呼んだ
    hm = _HEAD_LINE_RE.search(new_code)
    if hm and _heads_of_line(hm.group(1)):
        heads = _heads_of_line(hm.group(1))
        tables = [(case.get('before_values') or {}).get('values')] + [(t.get('expect') or {}).get('values') for t in case.get('tests') or []]
        words_of = [vp.sheet_words(v) for v in tables]
        kept = [h for h in heads if all(vp.head_in(h, w) for w in words_of)]
        if kept and kept != heads:
            new_code = re.sub(r"(^\s*'\s*見出し\s*[:：]\s*).*$", lambda m: m.group(1) + ' '.join(kept), new_code, count=1, flags=re.M)
            heads = kept
        # ほかの仕事の表にも見出しがそろうなら、全部の表に共通してその表に無い語を 1 つ足す
        # （2026-09-18: 課名を引くを足したら、突合の「伝票番号 支出額」がそろい AI を呼んだ）
        mine = str(case.get('name') or '')
        common = _heads_common(case, heads)
        for other, words in _OTHER_JOB_WORDS.items():
            if other == mine or not all(vp.head_in(h, words) for h in heads):
                continue
            extra = sorted(k for k in common if k not in words and k not in heads)
            if extra:
                heads = heads + [extra[0]]
                new_code = re.sub(r"(^\s*'\s*見出し\s*[:：]\s*).*$", lambda m: m.group(1) + ' '.join(heads), new_code, count=1, flags=re.M)
    return new_code, new_ask


def _validate_code(code):
    """規則の検査（純 Python）。戻り値＝(Sub 名, 依頼の語, 扱う語) か、(None, 理由, None)。"""
    if not code:
        return None, "Sub 〜 End Sub が返事にありません", None
    if _FUNC_RE.search(code):
        names = re.findall(r'^\s*(?:Public\s+|Private\s+)?Function\s+([^\s\(]+)', code, re.M | re.I)
        nm0 = names[0] if names else '関数'
        return None, (f"Function は作らない（Sub 1 本だけ）。{('・'.join(names[:3]) + ' を') if names else ''}"
                      f"Sub の中の GoSub に書き換える: 呼ぶ側は「s = 元の値: GoSub {nm0}: v = 戻り値の変数」、"
                      f"本体は Exit Sub の後ろに「{nm0}:」ラベルで置き、最後に Return。"
                      "引数と戻り値は Sub の中で宣言した変数でやりとりする（Dim s As String / Dim v As Double）"), None
    if _ARGSUB_RE.search(code):
        return None, "引数つきの Sub は作らない", None
    if re.search(r'^\s*Option\s+Explicit', code, re.M | re.I):
        return None, "Option Explicit は書かない", None
    subs = _SUB_RE.findall(code)
    if len(subs) != 1:
        return None, f"Sub は 1 本だけ（{len(subs)} 本あります）", None
    name = subs[0]
    if name in ('表を整える', '重複行を消す'):
        return None, f"「{name}」は既にあるマクロの名前です。別の名前にしてください", None
    ask = _ASK_LINE_RE.search(code)
    handles = _HANDLE_LINE_RE.search(code)
    if not ask or not handles:
        return None, "頭の 2 行（' 依頼の語: … と ' 扱う: …）が要ります", None
    try:
        re.compile(ask.group(1))
    except re.error as ex:
        return None, f"依頼の語が正規表現として読めません: {ex}", None
    if re.search(r'[.*+?()\[\]{}^$\\]', ask.group(1)):
        # 2026-09-17 夕: 「比較.*差異」「シート.*集」は文の端から端まで当たり、どの仕事を名指ししたかの長さ比べが壊れた
        return None, "' 依頼の語: は語を | で並べるだけ（.* ( ) [ ] などは使わない。例: 照合|比較|差異|突合）", None
    if re.search(r'MsgBox', code, re.I):
        return None, "MsgBox は出さない（成功は無言）", None
    if re.search(r'\bThisWorkbook\b', code, re.I):
        return None, "ThisWorkbook は使わない（マクロはアドインに入る＝表のブックは ActiveWorkbook）", None
    bad = [x for x in re.findall(r'CreateObject\s*\(\s*"([^"]+)"', code, re.I) if x.lower() != 'scripting.dictionary']
    if bad:
        return None, f"CreateObject は \"Scripting.Dictionary\" だけ（{bad[0]} は職場の PC で動かないことがある）", None
    why = _too_long(code)
    if why:
        return None, why, None
    return name, ask.group(1).strip(), handles.group(1).split()


_MAX_CONT_LINES = 26          # 1 文の行数の上限（VBA は継続行 25 本まで＝26 行。越えると取り込めない）
_MAX_LINE_CHARS = 1000        # 1 行の字数の上限（VBA は 1023 字）


def _too_long(code):
    """VBA に入らない長さ（純 Python）→ 理由 か None。
    2026-09-18: 縦持ちの M の式を「& _」で 27〜33 行つないだ版が「マクロを読み込めませんでした（例外）」になり、
    AI には理由が返らず 3 往復とも同じ書き方をくり返した（VBA の継続行は 25 本まで）。"""
    lines = [ln.rstrip('\r') for ln in str(code or '').split('\n')]
    over = [ln for ln in lines if len(ln) > _MAX_LINE_CHARS]
    if over:
        return (f"1 行が長すぎます（{len(over[0])} 字・VBA は 1 行 1023 字まで）: 「{over[0].strip()[:60]}…」。"
                "長い文字は s = s & \"…\" で 1 行ずつ足す")
    stmt, n = [], 0
    for ln in lines:
        code_part = re.sub(r'"[^"]*"', '""', ln).split("'", 1)[0]
        stmt.append(ln)
        if not code_part.rstrip().endswith(' _'):
            if len(stmt) > _MAX_CONT_LINES:
                head = next((x.strip() for x in stmt if x.strip()), '')
                return (f"1 つの文を {len(stmt)} 行につないでいます（VBA の行の継続は 25 本まで＝26 行）。"
                        f"「{head[:50]}…」を分ける: 長い M の式や文字は m = m & \"…\" & vbCrLf を 1 行ずつ足す"
                        "（& _ でつなぎ続けない）")
            stmt = []
        n += 1
    return None


_ASK_WORD_MAX = 8             # 依頼の語 1 つの長さの上限（文の丸暗記を防ぐ）
_ASK_KATAKANA_MAX = 12        # カタカナだけの 1 語（外来語）は 12 字まで（2026-09-18 第二期: 「ウォーターフォール」9 字が外され、
                              # 「ウォーターフォール図を作って」で撃てなくなった＝文の丸写しではなく 1 つの語）


def _ask_len_ok(w):
    """依頼の語の長さの決まり（純 Python）: 2〜8 字、またはカタカナだけの 12 字までの 1 語。"""
    return 2 <= len(w) and (len(w) <= _ASK_WORD_MAX or bool(len(w) <= _ASK_KATAKANA_MAX and re.fullmatch(r'[ァ-ヺー]+', w)))
_OTHER_JOB_TEXTS = {}         # forge() が台帳から入れる: ほかの合格した仕事の {名前: [依頼文・言い換え]}（_check_fit のぶつかりの検査）
_OTHER_JOB_WORDS = {}         # forge() が台帳から入れる: ほかの合格した仕事の {名前: 撃つ前の表の見出しの語}（見出しのぶつかり）
_OTHER_JOB_ASKS = {}          # forge() が台帳から入れる: ほかの仕事の {名前: 依頼の語}（ぶつかりは、その文でほかの仕事の語に負けるなら数えない）


def _clash_jobs(w, mine=''):
    """語 w が、ほかの仕事の依頼文・言い換えに出る → [(仕事, 文)]（純 Python）。
    入口は当たった登録を**全部**撃つ（長い語が勝つのではない）。弱く当たっただけの仕事も一緒に撃たれて表を壊すので、
    勝ち負けに関係なくぶつかりとする（2026-09-18 朝: 全部を「見出し: なし」にして見出しの一致で外れる仕組みが消え、
    比較の依頼で突合の「突き合」が撃たれ「相手の金額」の列が付いた。ほかに 構成比・縦棒グラフ・選んだ列・折れ線・合計を出 の 6 語）。"""
    out = []
    for other, texts in _OTHER_JOB_TEXTS.items():
        if other == mine:
            continue
        for text in texts:
            if w in text:
                out.append((other, text))
    return out
# 鍛えたマクロの「依頼の語」が当たってはいけない、よくある別の仕事の依頼（道具が持つ・fixtures\check_phrases.py の MISS とは別の文）
_GENERIC_REQUESTS = ('表をきれいに整えてください', '同じ行が二重になっているので消して', '日付の順に並べ替えて',
                     '月ごとの件数を数えて', '課ごとの人数を数えて',   # 項目ごとの件数（B 集計 13）とは別の仕事

                     '棒グラフを作ってください', '金額の書き方をそろえて', '備考の列を足して', 'シートを追加して',
                     '全角の空白を消して', 'わかりやすくまとめてください', '昨年度との差を出して',   # 「文字の数字を数値に直して」は数値直しそのもの（9/18 外した）
                     'フォントをそろえて', '回答の有無を確認して', '課ごとに色分けして', '一覧表の見出しを太字に', '照会文を作って', '支出額の表を見やすくして', '伝票の金額を確認して', '伝票番号の順に並べて',
                     '名簿を五十音順にして', '住所を都道府県と市町村に分けて', 'マスタの表を最新にして', '明細のテーブルを日付順に', '課別の集計表を作って',
                     'このピボットにピボットグラフも作って', 'テーブルにスライサーを付けて',
                     '執行率の列を足して', '執行額の列を太字にして', '予算額を千円単位にして', '見出しのセルをマージして中央に',
                     '科目別に色を付けて', '件数の多い順に並べて',
                     'グラフの元データの範囲を変えて', '印刷範囲を広げて', '入力規則のデータソースを変えて',
                     '新しいピボットを作って。元の範囲は見出しから最後の行まで',
                     'ブックを保存して', 'セルの範囲を選んで', 'データを貼り付けて', 'テーブルを作って', 'グラフを作って', 'マクロを直して',
                     '下の合計の行は入れない', '前に作ったものは消して作り直す',
                     '基本給の列を太字に', '名簿を所属順に並べて', '一覧を絞り込んで表示', '日付の降順で並べて', '別シートにコピーして',   # 「グラフの見せ方を変えて」はグラフ体裁そのもの（9/18 第二期で外した）
                     '前に作ったグラフが残っていたら消して作り直す', '件数の推移を見たい',
                     '受付日の順に並べて',   # 列の名前（受付日）を依頼の語にさせない（9/18 日付直しが「受付日」を選んだ）
                     # 9/18 第二期: ピボット値貼りが「普通の表」「値だけの表」「外して集計」を選んだ＝ほかの仕事の文にも出る
                     'テーブルを普通の表に戻して', '数式を消して値だけの表にして', 'フィルタを外して集計して',   # テーブル解除は 9/18 に引退＝戻した
                     '名簿の更新をお願いします',   # 9/18 第二期: PQ更新が「更新をお願」を選んだ
                     'この一覧を表の形にして', 'この表を見やすい表にして', '備考の列を消して')   # 9/18 第二期: ピボット体裁が「表の形に」「見やすい表」を選んだ


def _heads_of_line(text):
    """' 見出し: の行の中身 → 語のリスト（純 Python）。「なし」は []（見出しの語に頼らないマクロ・2026-09-18）。"""
    s = str(text or '').strip()
    return [] if s in ('なし', '無し') else s.split()


# 表ごとに違わない語（どの表にも出る一般の語）＝マクロに書いてよい
_LITERAL_OK = ('合計', '計', '小計', '総計', '累計', '行ラベル', '列ラベル', '(空白)', '（空白）', '空白')
_LITERAL_SPLIT_RE = re.compile(r'[,，|;；/／、・\s]+')
_NUMBERISH_RE = re.compile(r'^[\s\d０-９,，.．/\-:：%％円年月日△▲+()（）]*$')


def _code_literals(code):
    """VBA の文字列リテラル（コメントの中は除く）→ [文字]（純 Python）。"" は " 1 文字。"""
    out = []
    for ln in str(code or '').splitlines():
        if re.match(r'^\s*Rem\b', ln, re.I):
            continue
        buf, i, n = None, 0, len(ln)
        while i < n:
            ch = ln[i]
            if buf is None:
                if ch == "'":
                    break
                if ch == '"':
                    buf = []
            elif ch == '"':
                if i + 1 < n and ln[i + 1] == '"':
                    buf.append('"')
                    i += 1
                else:
                    out.append(''.join(buf))
                    buf = None
            else:
                buf.append(ch)
            i += 1
    return out


def _table_word_keys(snap):
    """表の姿（snap）→ その表に出る語のならし（head_key）の集合（純 Python）。
    上 15 行の文字のセル・シート名・ほかのシートの名前と頭・テーブル名・クエリ名。数や日付のセルは入れない。"""
    import vbam_prefire as vp
    keys = set()
    snap = snap or {}
    vals = snap.get('values') or []
    kinds = snap.get('kinds') or []
    for i, row in enumerate(vals[:vp._SHEET_WORD_ROWS]):
        for j, v in enumerate(row or []):
            k = kinds[i][j] if i < len(kinds) and j < len(kinds[i]) else 's'
            if k == 's' and isinstance(v, str) and v.strip() and not _NUMBERISH_RE.match(v):
                keys.add(vp.head_key(v))
    if snap.get('sheet'):
        keys.add(vp.head_key(snap['sheet']))
    for o in snap.get('others') or []:
        if o.get('name'):
            keys.add(vp.head_key(o['name']))
        for line in str(o.get('head') or '').splitlines()[1:]:
            for v in line.split('\t')[1:]:
                if v.strip() and not _NUMBERISH_RE.match(v):
                    keys.add(vp.head_key(v))
    objs = snap.get('objects') or {}
    for key in ('tables', 'queries'):
        for t in objs.get(key) or []:
            if t.get('name'):
                keys.add(vp.head_key(t['name']))
    return keys


def _literal_words(case, code):
    """マクロに書き込まれた、表ごとに違う語（純 Python）→ [(語, [出る表], [出ない表])]。
    文字列リテラル（と , | ・ などで区切った一片）のうち、弾の表・別の表のセル（上 15 行）・シート名・テーブル名に出る語で、
    全部の表には出ないもの。除外: 合計・計・小計・総計・行ラベル・列ラベル・(空白)、依頼文と言い換えに含まれる語。
    2026-09-18: 鍛えた 27 本が課名・支出額・T支出明細・課マスタを書き込み、見出しの語が違う職場の表で動かなかった
    （お題の表が全部同じ語だった＝試験で落ちなかった）。"""
    import vbam_prefire as vp
    tables = [('撃つ前の表', [case.get('before_values'), case.get('expect')])]
    tables += [(f"別の表「{t.get('label')}」", [t.get('expect')]) for t in case.get('tests') or []]
    if len(tables) < 2:
        return []                                          # 1 枚では表ごとの違いが見えない
    keysets = []
    for label, snaps in tables:
        ks = set()
        for s in snaps:
            ks |= _table_word_keys(s)
        keysets.append((label, ks))
    texts = [str(case.get('request') or '')] + list(case.get('phrases') or [])
    ok_keys = {vp.head_key(w) for w in _LITERAL_OK}
    out, seen = [], set()
    for lit in _code_literals(code):
        pieces = [lit] + [p for p in _LITERAL_SPLIT_RE.split(lit) if p and p != lit]
        for p in pieces:
            k = vp.head_key(p)
            if not k or k in seen or k in ok_keys or _NUMBERISH_RE.match(p) or not re.search(r'\w', p):
                continue
            if any(p.strip() in t or k in vp.head_key(t) for t in texts if t):
                continue
            has = [label for label, ks in keysets if k in ks]
            if has and len(has) < len(keysets):
                seen.add(k)
                out.append((p.strip(), has, [label for label, ks in keysets if k not in ks]))
    return out


def _literal_error(case, code):
    """書き込まれた語の違反 → AI に返す文（無ければ None・純 Python）。"""
    words = _literal_words(case, code)
    if not words:
        return None
    return ("表の語をマクロに書き込んでいます（職場の表は見出しの語・シート名・テーブル名・項目の名前が違う＝その表では動かない）: "
            + " ／ ".join(f"「{w}」＝{'・'.join(has[:2])}{'ほか' if len(has) > 2 else ''}にあり、{'・'.join(miss[:2])}"
                          f"{'ほか' if len(miss) > 2 else ''}には無い" for w, has, miss in words[:6])
            + "。語で探さず、表の形で探す（文字の列・数の列・日付の列・率の列＝表示形式に %・左端／右端・見出しの並び・"
            "依頼に書いた列の決まり）。シート・テーブルも名前でなく形で探す（見出しの並びが同じ・同じ見出しの列がある・"
            "アクティブなシート）。作る物の名前は依頼に書いてある名前か、元の表・テーブル・シートの名前から作る")


def _heads_common(case, heads):
    """弾の表と別の表の全部で、見出しの行（heads が全部ある行）に共通してある語（head_key の集合・純 Python）。"""
    import vbam_prefire as vp
    tables = [(case.get('before_values') or {}).get('values')] + [(t.get('expect') or {}).get('values') for t in case.get('tests') or []]
    common = None
    for values in tables:
        rows = [[str(v).strip() for v in (row or []) if isinstance(v, str) and str(v).strip()] for row in (values or [])[:15]]
        head_row = next((r for r in rows if heads and all(any(vp.head_key(a) == vp.head_key(x) for a in h.split('|') for x in r)
                                                          for h in heads)), [])
        keys = {vp.head_key(x) for x in head_row}
        common = keys if common is None else (common & keys)
    return common or set()


def _job_shape(case):
    """この仕事の ' 形:（純 Python）＝本番の表と別の表（tests）の形の積集合。
    「合格した表の全部に共通する形」だけを必要条件にする＝証明した範囲でしか撃たない（安全側・2026-09-18）。"""
    import vbam_prefire as vp
    items = [(case.get('before'), case.get('sheet'))]
    items += [(t.get('before'), t.get('sheet')) for t in case.get('tests') or []]
    common = None
    for path, sheet in items:
        if not path or not os.path.isfile(str(path)):
            continue
        got = shape_of_file(path, sheet)
        common = set(got) if common is None else (common & set(got))
    # ピボット・グラフは本番の表にあるときだけ必要条件にする（2026-09-18 第二期: 「何もしない」を確かめる試験の表（ピボットなし）が
    # 積集合から落とし、ピボットの無いシートでも撃たれていた）。撃つ範囲を狭める側にしか働かない
    main = shape_of_file(items[0][0], items[0][1]) if items[0][0] and os.path.isfile(str(items[0][0])) else set()
    return [w for w in vp.SHAPE_WORDS
            if (w in main if w in vp.OBJECT_SHAPE_WORDS else w in (common or set()))]      # 並びは SHAPE_WORDS の順


def ensure_shape_line(case, code):
    """' 形: の行を道具が入れる／導いた形に直す（純 Python）→ (code, 変えたか)。AI には書かせない。
    入れる所は ' 選ぶ列: の次（無ければ ' 見出し: の次）。積集合が空なら「なし」＝形では絞らない。"""
    import vbam_prefire as vp
    text = str(code or '')
    if not text.strip() or not _SUB_RE.search(text):
        return code, False
    want = _job_shape(case)
    line = "' 形: " + (" ".join(want) if want else "なし")
    lines = text.replace('\r\n', '\n').split('\n')
    at = next((i for i, ln in enumerate(lines) if _SHAPE_LINE_RE.match(ln)), None)
    if at is not None:
        cur = _SHAPE_LINE_RE.match(lines[at]).group(1).strip()
        now = [] if cur in ('なし', '無し') else [w for w in cur.split() if w in vp.SHAPE_WORDS]
        if now == want:
            return code, False
        lines[at] = re.match(r'\s*', lines[at]).group() + line
        return "\n".join(lines), True
    hi = next((i for i, ln in enumerate(lines) if _HEAD_LINE_RE.match(ln)), None)
    si = next((i for i, ln in enumerate(lines) if _SELECT_LINE_RE.match(ln)), None)
    at = max([i for i in (hi, si) if i is not None], default=None)
    if at is None:
        return code, False                                 # ' 見出し: が無い＝_check_fit が返す（形の行はその後）
    lines.insert(at + 1, re.match(r'\s*', lines[at]).group() + line)
    return "\n".join(lines), True


def _check_fit(case, code, ask_re):
    """登録簿の行がこの弾に合っているか（純 Python）。→ (見出しの語, None) か (None, 理由)。
    - ' 見出し: の語が、撃つ前の表にも別の表の正解にも全部ある（シートにこの語があるときだけ撃つ＝誤爆を防ぐ）
    - 依頼の語が、この依頼文に当たる（当たらなければ次から撃たれない）"""
    import vbam_prefire as vp
    m = _HEAD_LINE_RE.search(code or '')
    if not m:
        return None, "頭の 3 行目 ' 見出し: …（ふつうは「なし」＝表の形で列を探す。特定の語が要る仕事だけ語）が要ります"
    heads = _heads_of_line(m.group(1))                     # 「なし」は []＝見出しの検査もぶつかりの検査もしない
    # 列を選ぶ仕事（お題に「選んでいる列」の覚書がある）は、頭に ' 選ぶ列: N が要る（入口が列を選んでから撃つ・2026-09-18）
    if case.get('sel'):
        want = len([a for a in str(case['sel']).split(',') if a.strip()])
        got = _SELECT_LINE_RE.search(code or '')
        if not got:
            return None, (f"頭に ' 選ぶ列: {want}（このマクロが使う、選んでいる列の数）が要ります。列は Selection（選んでいる"
                          f"セルの列）から取る決まりの仕事です（このお題では {want} 列＝{case['sel']}）")
        lo, hi = int(got.group(1)), int(got.group(2) or got.group(1))
        if not lo <= want <= hi:
            return None, f"' 選ぶ列: の数（{lo}〜{hi}）が、このお題で選んでいる列の数（{want} 列＝{case['sel']}）と合いません"
    # ' 形:（入口が先に絞る表の形）。行は道具が足す＝無いことは違反にしない。あれば、全部の表にその形があることを見る
    ms = _SHAPE_LINE_RE.search(code or '')
    if ms:
        cur = ms.group(1).strip()
        want_shape = [] if cur in ('なし', '無し') else cur.split()
        bad = [w for w in want_shape if w not in vp.SHAPE_WORDS]
        if bad:
            return None, (f"' 形: に使えない語があります（{'・'.join(bad)}）。使えるのは {'・'.join(vp.SHAPE_WORDS)}"
                          "。この行は道具が書きます＝書かなくてかまいません")
        for label, path, sheet in ([('撃つ前の表', case.get('before'), case.get('sheet'))]
                                   + [(f"別の表「{t.get('label')}」", t.get('before'), t.get('sheet'))
                                      for t in case.get('tests') or []]):
            if not path or not os.path.isfile(str(path)):
                continue
            miss = [w for w in want_shape if w not in shape_of_file(path, sheet)
                    and not (w in vp.OBJECT_SHAPE_WORDS and label != '撃つ前の表')]      # ピボット・グラフは本番の表だけで決める
            if miss:
                return None, (f"' 形: の「{'・'.join(miss)}」が{label}の形にありません＝その表では撃たれません"
                              "（この行は道具が書きます＝書かなくてかまいません）")
    tables = [('撃つ前の表', (case.get('before_values') or {}).get('values'))]
    tables += [(f"別の表「{t.get('label')}」", (t.get('expect') or {}).get('values')) for t in case.get('tests') or []]
    for label, values in tables:
        words = vp.sheet_words(values)
        miss = [h for h in heads if not vp.head_in(h, words)]
        if miss:
            return None, f"' 見出し: の「{'・'.join(miss)}」が{label}の見出しにありません（表にそのまま書いてある語だけ。表によって見出しの語が違うなら「課名|行ラベル」のように | で並べる＝どれか 1 つがあればよい）"
    # ほかの鍛えた仕事の表でも見出しがそろってしまうと、その表で撃つ（2026-09-18 未明: 推移の折れ線の見出しが「課名」だけで、
    # 「月次集計表を作って、課ごとの折れ線グラフも」の支出明細の表でも撃った）
    mine = str(case.get('name') or '')
    common = _heads_common(case, heads)
    for other, words in _OTHER_JOB_WORDS.items():
        # 全部の表に共通する見出しの語に、ほかの表に無い語が 1 つも無ければ、見出しでは区別できない＝この決まりは当てない
        # （2026-09-18 未明: 月の見出しが日付の試験の表には「4月」が無く、「4月を足せ」と「表に無い」を交互に破った）
        if not [k for k in common if k not in words]:
            continue
        if other != mine and heads and all(vp.head_in(h, words) for h in heads):
            return None, (f"' 見出し: の語（{'・'.join(heads)}）が、ほかの仕事「{other}」の表にもそろっています＝その表でも撃たれます。"
                          "この表にしか無い見出しの語を足す（例: 月の列の「4月」・「決算統計区分」など）")
    # 依頼の語の決まりは全部まとめて返す（2026-09-17 夜: 1 つずつ返したら、AI が 1 つ直すたびに別の決まりを破り
    # 「一般語を外す→長すぎる→言い換えに当たらない」を 6 往復くり返した）
    probs = []
    try:
        if not re.search(ask_re, str(case.get('request') or '')):
            probs.append("この依頼文に当たりません（次からこの依頼で撃たれない）")
        texts = [str(case.get('request') or '')] + list(case.get('phrases') or [])
        own_texts = list(texts)                    # 下のぶつかりの検査で texts はほかの仕事の文に置き換わる
        # 使える語が 1 つも無い言い換え（一般の依頼の語とほかの仕事の語だけでできた文）は、語では見分けられない＝求めない
        # （2026-09-18: 「ピボットとピボットグラフとスライサーを 1 枚に」＝ピボットグラフ・スライサーを一般の依頼に足したら直せなくなった）
        miss = [p for p in case.get('phrases') or [] if not re.search(ask_re, p)
                and _ask_candidates(p, [x for x in texts if x != p])]
        if miss:
            cover = _ask_cover(miss, texts)
            probs.append("言い換えに当たりません（" + "／".join(
                f"「{p}」＝使える語の候補: {'・'.join(_ask_candidates(p, [x for x in texts if x != p])) or 'なし'}"
                for p in miss[:6]) + "）。語は文にそのまま含まれる文字列。"
                + (f"**この語を | で足せば外れた {len(miss)} 本すべてに当たります: {'・'.join(cover)}**"
                   if cover else "候補から 1 つ足す（候補は一般の依頼に当たらない語）"))
        # 依頼文・言い換えのどれにも出ない語（お題の別の表の名前など）は、人の依頼では当たらない＝外す
        # 2026-09-18: 月次集計の依頼の語に「月次_試験」（お題のファイル名の切れ端）が入った
        stray = [w for w in ask_re.split('|') if w and not any(w in t for t in texts if t)]
        if stray:
            probs.append("依頼文にも言い換えにも出ない語があります（" + "・".join(f"「{w}」" for w in stray[:4])
                         + "）。その語は外す（依頼の語は、人が書く文にそのまま含まれる語だけ）")
        broad = [(g, vp._ask_best_word(ask_re, g)) for g in _GENERIC_REQUESTS if re.search(ask_re, g)]
        if broad:
            # 2026-09-17 夜: 集約の依頼の語「一覧」「まとめ」「課ごと」が「この表をまとめて見やすく」「一覧を印刷できるように整えて」に
            # 当たり、集約のブックでは見出しもそろうので撃たれる形だった
            probs.append("ほかの仕事の一般的な依頼にも当たります（" + "／".join(
                f"「{g}」に「{w}」" for g, w in broad[:4]) + "）。その語だけ外す（ほかの語は残す）")
        # ほかの鍛えた仕事の依頼文・言い換えに当たる語（2026-09-17 夜: 前年度比較の「突き合わせ」が突合の依頼
        # 「移行後のデータと突き合わせ」と同じ長さで引き分け、突合を撃てなくした）
        mine = str(case.get('name') or '')
        clash = []
        for other_name, texts in _OTHER_JOB_TEXTS.items():
            if other_name == mine:
                continue
            for text in texts:
                w = vp._ask_best_word(ask_re, text)
                if w and (w, other_name) not in clash and _clash_jobs(w, mine):
                    clash.append((w, other_name))
        if clash:
            probs.append("ほかの鍛えた仕事の依頼にも当たる語があります（" + "・".join(f"「{w}」＝{o}" for w, o in clash[:4])
                         + "）。その語は外す（この仕事にしか出ない語にする）")
        now = set(ask_re.split('|'))
        dropped = [w for w in str(case.get('keep_ask') or '').split('|')
                   if w and w not in now and _ask_len_ok(w) and not any(w in g for g in _GENERIC_REQUESTS)
                   and not re.search(r'[.*+?()\[\]{}^$\\]', w)                  # 記号入りの語は決まり違反＝残させない
                   and not any(x and x in w for x in now)                        # 残っている短い語が含む長い語は外してよい（当たる範囲は同じ）
                   and any(w in t for t in own_texts if t)                       # 言い換えを替えて今の文に出なくなった語は外してよい
                   and not _clash_jobs(w, mine)]
        # ↑ 2026-09-18 第二期: 言い換えを替えて出なくなった「中に小」を「残せ」と言い、次の往復で「文に出ない語は外せ」と言う板挟み
        # ↑ ほかの仕事とぶつかる語は残させない（2026-09-17 夜: 「突き合わせ」を外せ／残せ の板挟みで 6 往復）
        if dropped:
            # 2026-09-17 夜: 一般語を外させたら「集約」「帳票」「年度の月別集計」まで外し、言い方の答え合わせで当たらなくなった
            probs.append("決まりに反していない前の語を外しました: " + "・".join(f"「{w}」" for w in dropped[:8])
                         + "（外すのは決まりに反した語だけ。これらは残す）")
        punct = [a for a in ask_re.split('|') if re.search(r'[、。，．,「」（）：:]', a)]
        if punct:
            probs.append("句読点をまたいだ語があります（" + "・".join(f"「{w}」" for w in punct[:4])
                         + "）。その語は外す（文の切れ端は人の依頼に当たらない）")
        long_alts = [a for a in ask_re.split('|') if len(a) > _ASK_WORD_MAX and not _ask_len_ok(a)]
        if long_alts:
            # 2026-09-17 夜: 一般語を外させたら、言い換えの文をそのまま依頼の語にした（「旧システムと新システムの伝票を照らし合わせ」）。
            # 丸暗記の語は言い方が少し変わると当たらず、「突合」「集約」の核の語まで消えていた
            probs.append(f"語は {_ASK_WORD_MAX} 字まで（文を写すと言い方が少し変わると当たらない）: "
                         + "・".join(f"「{a}」" for a in long_alts[:4]) + " を短い語に切る")
    except re.error as ex:
        return None, f"依頼の語が正規表現として読めません: {ex}"
    if probs:
        return None, ("' 依頼の語: の決まり違反（全部を同時に満たす）: " + " ／ ".join(f"{i}) {p}" for i, p in enumerate(probs, 1))
                      + f"。語は 2〜{_ASK_WORD_MAX} 字・この仕事にしか出ない語（例: 突合・新システム・集約・按分・決算統計・受付簿）を | で並べる")
    # 入口で撃ったあと、この依頼（と言い換え）の仕事が AI に回されないこと（回されると AI 0 回にならない）
    reg = vp.registry_from_text(code)
    words = vp.sheet_words((case.get('before_values') or {}).get('values'))
    for text in [str(case.get('request') or '')] + list(case.get('phrases') or []):
        p = vp.plan_full(text, reg, words)
        # グラフは例外にしない（2026-09-17 深夜: グラフを作るマクロも鍛えるようになった＝「グラフ」を扱うに入れさせる）
        # 語が当たらない言い換え（見分けられず求めなかった文）は撃たない＝AI に回るのが正しい（2026-09-18）
        if p['left'] and p['extras'] and not re.search(r'印刷|PDF', text):
            return None, (f"依頼「{text[:40]}」の「{'・'.join(p['left'])}」が、このマクロの仕事と見なされず AI に回されます。"
                          "' 扱う: にその語を足してください（この依頼の仕事なら）")
    return heads, None


def _write_bas(path, module, code):
    """cp932 の .bas に書く（Attribute 行つき）。cp932 に無い文字があれば理由を返す。"""
    body = f'Attribute VB_Name = "{module}"\r\n' + code.replace('\r\n', '\n').replace('\n', '\r\n')
    try:
        data = body.encode('cp932')
    except UnicodeEncodeError as ex:
        return f"cp932 に無い文字があります（{ex.object[ex.start:ex.end]!r}）。VBA に入れられる文字だけ使ってください"
    with open(path, 'wb') as f:
        f.write(data)
    return None


def _check_bas(path):
    """取り込む前の関所（vba_manager の check-bas＝文字コード・改行の二重化・識別子）。"""
    import vba_manager as vm
    try:
        return bool(vm._check_bas_one(path))
    except Exception as ex:
        print(f"  check-bas が撃てませんでした: {ex}")
        return False


# ----------------------------------------------------------------
# 試す（写し・自分の Excel・AI なし。.bas は _check_bas を通ったものだけ読み込む）
# ----------------------------------------------------------------

def _compare(expect, got):
    """期待（AI の後の姿）と、マクロの後の姿を値で突き合わせる → 不一致の行（空なら合格）。"""
    ev, gv = expect.get('values') or [], got.get('values') or []
    r0, c0 = int(expect.get('row') or 1), int(expect.get('col') or 1)
    out = []
    # 末尾の空の行・列は数えない（2026-09-17 夜: 折りたたんだ行の表で、書式だけの空の最終行が正解の使用範囲に入り、
    # 入口の tidy の後は使用範囲が 1 行縮んで「表の大きさが違う」と出た＝値は同じ）
    er, ec = _filled_size(ev)
    gr, gc = _filled_size(gv)
    if (er, ec) != (gr, gc) or (int(got.get('row') or 1), int(got.get('col') or 1)) != (r0, c0):
        out.append(f"表の大きさ／位置が違う: 期待 {er}×{ec}（左上 {_addr_of(r0, c0, 0, 0)}）・"
                   f"マクロ後 {gr}×{gc}（左上 {_addr_of(int(got.get('row') or 1), int(got.get('col') or 1), 0, 0)}）")
    ek, gk = expect.get('kinds'), got.get('kinds')          # 古い台帳の正解には無い＝型は比べない
    for i in range(max(er, gr)):
        for j in range(max(ec, gc)):
            e = ev[i][j] if i < er and j < len(ev[i]) else ''
            g = gv[i][j] if i < gr and j < len(gv[i]) else ''
            if _norm(e) != _norm(g):
                out.append(f"{_addr_of(r0, c0, i, j)}: 期待 '{e}' ／ マクロ後 '{g}'")
            elif ek and gk and e != '':
                a = ek[i][j] if i < len(ek) and j < len(ek[i]) else '-'
                b = gk[i][j] if i < len(gk) and j < len(gk[i]) else '-'
                if a != b and 's' in (a, b) and ({a, b} & {'n', 'd'}):
                    out.append(f"{_addr_of(r0, c0, i, j)}: 型が違う 期待 {_KIND_NAME.get(a, a)} '{e}' ／ マクロ後 {_KIND_NAME.get(b, b)} '{g}'")
    if 'fx' in expect:                                      # 古い台帳の正解には無い＝式かどうかは比べない
        ef, gf = set(expect.get('fx') or []), set(got.get('fx') or [])
        miss, extra = sorted(ef - gf, key=_addr_key), sorted(gf - ef, key=_addr_key)
        if miss:
            out.append(f"式でない（期待は式）: {'・'.join(miss[:8])}" + (f" ほか {len(miss) - 8} か所" if len(miss) > 8 else '')
                       + "（値を書かずに .Formula で式を入れる）")
        if extra:
            out.append(f"式になっている（期待は値）: {'・'.join(extra[:8])}" + (f" ほか {len(extra) - 8} か所" if len(extra) > 8 else '')
                       + "（式を値にするなら rng.Value = rng.Value）")
    if 'filter' in expect and (expect.get('filter') or '') != (got.get('filter') or ''):
        out.append(f"オートフィルタの範囲: 期待 {expect.get('filter') or 'なし'!r} ／ マクロ後 {got.get('filter') or 'なし'!r}"
                   "（ws.Range(見出しの行〜最後の行).AutoFilter で付ける。前のフィルタは ws.AutoFilterMode = False で外す）")
    if 'page' in expect:                                    # 正解に印刷設定があるときだけ（既定のままの正解は持たない）
        out += _compare_page(expect.get('page') or {}, got.get('page'))
    if 'book' in expect:                                    # 古い台帳の正解には無い＝ほかのシートは比べない
        out += _compare_book(expect.get('book') or {}, got.get('book') or {})
    if 'objects' in expect:                                 # 古い台帳の正解には無い＝グラフ・ピボット・テーブルは比べない
        import vbam_objects as vo
        objs = vo.compare_objects(expect.get('objects'), got.get('objects'))
        top = [[str(v).strip() for v in (row or []) if isinstance(v, str) and str(v).strip()] for row in (expect.get('values') or [])[:10]]
        for k, line in enumerate(objs):
            m = re.search(r": タイトル 期待 '([^']+)' ／", str(line))
            row = next((r for r in top if m and m.group(1) in r), None)
            if row is not None:
                # 2026-09-18: 見やすい推移グラフで、表の上の表題をタイトルにする型が読めなかった
                # 見出しの行（値が 2 つ以上並ぶ行）のセルなら「見出し」と言う（2026-09-18: 値の列の見出しを表題と言い違えた）
                objs[k] = (f"{line}（期待のタイトル「{m.group(1)}」は、" + (
                    "見出しの行にそのまま書いてある文字＝その列の見出しをタイトルにする）" if len(row) >= 2 else
                    "シートの表の上のセルにそのまま書いてある文字＝表題をタイトルにする）"))
        out += objs
    return out


def _addr_key(a):
    """「B12」→ (12, 2)（番地を行・列の順に並べる）。"""
    m = re.match(r'([A-Z]+)(\d+)$', str(a))
    if not m:
        return (0, 0)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return (int(m.group(2)), col)


def _compare_book(eb, gb, limit=6):
    """ブックのシートの並びと、ほかのシートの値を比べる（純 Python）→ 不一致の行。"""
    out = []
    en, gn = eb.get('names') or [], gb.get('names') or []
    if en != gn:
        miss = [n for n in en if n not in gn]
        extra = [n for n in gn if n not in en]
        out.append(f"シートの並び: 期待 {en[:12]} ／ マクロ後 {gn[:12]}"
                   + (f"（無いシート {miss[:6]}）" if miss else '') + (f"（余計なシート {extra[:6]}＝前に作ったシートが残っている）" if extra else '')
                   + ("（名前は同じで並びが違う＝Worksheets.Add After:=最後のシート で後ろへ足す）" if not miss and not extra else ''))
    gs = gb.get('sheets') or {}
    for name, e in (eb.get('sheets') or {}).items():
        g = gs.get(name)
        if g is None:
            continue                                        # 無いシートは上で言っている
        sub = _compare({'row': e.get('row'), 'col': e.get('col'), 'values': e.get('values')},
                       {'row': g.get('row'), 'col': g.get('col'), 'values': g.get('values')})
        for line in sub[:limit]:
            out.append(f"シート「{name}」 {line}")
        if len(sub) > limit:
            out.append(f"シート「{name}」 ほか {len(sub) - limit} か所")
    return out


def _filled_size(rows):
    """値の表（文字の 2 次元）の、末尾の空の行・列を除いた大きさ（純 Python）。"""
    rows = rows or []
    nr = len(rows)
    while nr and all(str(v).strip() == '' for v in (rows[nr - 1] or [])):
        nr -= 1
    nc = max((len(r) for r in rows[:nr]), default=0)
    while nc and all(str(r[nc - 1]).strip() == '' for r in rows[:nr] if len(r) >= nc):
        nc -= 1
    return nr, nc


_KIND_NAME = {'n': '数', 'd': '日付', 's': '文字', 'b': '真偽', '-': '空'}


def _norm(v):
    s = '' if v is None else str(v)
    # 数の見た目の差（1,320 と 1320・1320.0）は同じとみなす（書式は tidy が当てる）
    t = s.replace(',', '')
    try:
        f = float(t)
        return repr(round(f, 6))
    except ValueError:
        return s.strip()


_RUN_TIMEOUT = 60             # 写しで撃つ 1 回の上限（秒）。越えたら自分の台を落とす＝無限ループの疑い
_HARNESS = '鍛冶_撃つ台'


def _harness_code(names, track=False):
    """試し撃ちの台（純 Python）: 名前を直接呼び、実行時エラーをダイアログにせず文字で持ち帰る。
    2026-09-17: 素の xl.Run だと実行時エラーで見えない Excel がデバッグのダイアログを出し、6 分止まった
    （AI のマクロが CreateObject("System.Collections.ArrayList") で落ちた）。run-macro の VMR と同じ考え方。
    names は _validate_code と check-bas を通った Sub 名と、道具の既定のマクロ名だけ。
    track＝抜けた Exit Sub の行（_instrument_exits が入れた印）を "OK|行" で持ち帰る。"""
    lines = ([f"Public {_EXIT_MARK} As String", f"Public {_LINE_MARK} As Long", ""] if track else []) + [
        "Function 鍛冶_撃つ() As String", "    Dim 段 As String", "    On Error GoTo eh"]
    if track:
        lines += [f'    {_EXIT_MARK} = ""', f'    {_LINE_MARK} = 0']
    for nm in names:
        lines += [f'    段 = "{nm}"', f"    {nm}"]
    # 実行時エラーは "ERR|段|番号|説明|行"（行は track のときだけ・説明に | があっても最後の区切りで取る）
    lines += ['    鍛冶_撃つ = "OK"' + (f' & "|" & {_EXIT_MARK}' if track else ''), "    Exit Function", "eh:",
              '    鍛冶_撃つ = "ERR|" & 段 & "|" & Err.Number & "|" & Err.Description'
              + (f' & "|" & {_LINE_MARK}' if track else ''), "End Function"]
    return "\r\n".join(lines) + "\r\n"


_EXIT_MARK = '鍛冶_抜けた'
_LINE_MARK = '鍛冶_行'
# 行の印を前に置けない行（ブロックの区切り・宣言・ラベル・Case・Sub の行）
_NO_LINE_MARK_RE = re.compile(r'^\s*(Case\b|Else\b|ElseIf\b|End\b|Next\b|Loop\b|Wend\b|Dim\b|Const\b|Static\b|Private\b|Public\b|'
                              r'Sub\b|Function\b|#|\w+:\s*$|\d+\s)', re.I)


def _instrument_exits(text, sub_name):
    """試し撃ち用の写し（純 Python）: Sub sub_name の中の Exit Sub の前に「抜けた行」を残す文を入れる。
    → (本文, {Sub の行を 1 とした行番号: その行の文字})。登録する .bas は変えない（試しのブックに入れる写しだけ）。
    2026-09-17 夜: AI が StrConv(…, vbNarrow) で見出しの「コード」を「ｺｰﾄﾞ」にして見出しが見つからず、2 往復とも
    「何も書かずに終わりました」しか返らず同じ誤りをくり返して止まった＝どの Exit Sub で抜けたかを返す。"""
    out, where = [], {}
    inside, k, prev_cont, first, after_select = False, 0, False, False, False
    start_re = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+' + re.escape(sub_name) + r'\s*\(', re.I)
    for ln in text.split('\n'):
        body = ln[:-1] if ln.endswith('\r') else ln
        cr = '\r' if ln.endswith('\r') else ''
        code_part = re.sub(r'"[^"]*"', '""', body).split("'", 1)[0]
        if not inside and start_re.match(body):
            inside, k, first = True, 0, True
        if inside:
            k += 1
            if not prev_cont and re.search(r'(?:^|[\s:])Exit\s+Sub\b', code_part, re.I):
                body = re.sub(r'\bExit\s+Sub\b', f'{_EXIT_MARK} = "{k}": Exit Sub', body, count=1, flags=re.I)
                where[k] = (ln[:-1] if ln.endswith('\r') else ln).strip()
            # 実行時エラーの行（2026-09-17 深夜: グラフのマクロが 91・438 で止まり、道具は番号しか返せず AI が当てずっぽうで直した）
            if (not first and not prev_cont and code_part.strip() and not _NO_LINE_MARK_RE.match(code_part)
                    and not after_select):
                indent = re.match(r'^\s*', body).group(0)
                out.append(f'{indent}{_LINE_MARK} = {k}{cr}')
            if re.match(r'^\s*End\s+Sub\b', code_part, re.I):
                inside = False
        if inside and code_part.strip():
            after_select = bool(re.match(r'^\s*Select\s+Case\b', code_part, re.I))
        first = False
        prev_cont = code_part.rstrip().endswith(' _')
        out.append(body + cr)
    return '\n'.join(out), where


def _sub_lines(text, sub_name):
    """モジュールの本文から Sub sub_name の行（Sub の行を 1 行目）を取り出す（純 Python）。"""
    out, inside = [], False
    start_re = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+' + re.escape(sub_name) + r'\s*\(', re.I)
    for ln in str(text or '').replace('\r\n', '\n').split('\n'):
        if not inside and start_re.match(ln):
            inside = True
        if inside:
            out.append(ln)
            if re.match(r'^\s*End\s+Sub\b', ln, re.I):
                break
    return out


_TEXT_NUMBER_RE = re.compile(r'^\s*[-△▲]?\s*[\d０-９][\d０-９,，]*(\.\d+)?\s*円?\s*$')


def _similar_heads(code, start=None, expect=None):
    """見出しを「含む」で探すマクロが取り違える列（純 Python）→ [(探している語, 取り違えそうな見出し)]。
    2026-09-17 夜: 集約が「前年度予算額」を、月次集計が「支出額（税抜）」を拾い、それらしい数字のまま外れた。"""
    if not re.search(r'\bInStr\s*\(|\bLike\s+"\*', code or '', re.I):
        return []
    import vbam_prefire as vp
    lits = set(re.findall(r'"([^"\n]{2,12})"', code))
    lit_keys = {vp.head_key(x) for x in lits}
    rows = [[str(v).strip() for v in (row or []) if isinstance(v, str)] for row in ((start or {}).get('values') or [])[:10]]
    for o in (expect or {}).get('others') or []:
        rows += [[x.strip() for x in line.split('\t')[1:]] for line in str(o.get('head') or '').splitlines()[1:4]]
    heads = set()
    for row in rows:
        # 見出しの行＝マクロが探している語がそのまま入っている行だけ（2026-09-17 夜: 前の比較表のデータ「旧科目1」まで見出しと数えた）
        if any(vp.head_key(x) in lit_keys for x in row):
            heads.update(x for x in row if 1 < len(x) <= 20)
    keys = {vp.head_key(a) for h in heads for a in h.split('|')}                # マクロの文字が表の見出しそのもののときだけ（「（例」「予算」は言わない）
    out = []
    for w in sorted(lits):
        if not re.search(r'[一-龥ァ-ヶ]', w) or vp.head_key(w) not in keys:
            continue
        for h in sorted(heads):
            if h != w and w in h and vp.head_key(h) != vp.head_key(w) and (w, h) not in out:
                out.append((w, h))
    return out[:4]


def _trap_hints(code, start=None, expect=None):
    """AI のマクロの字面に出る VBA の罠（純 Python）→ 知らせる文のリスト。start＝撃つ前の表の snap（あれば表の中身も見る）。"""
    out = []
    code = code or ''
    sim = _similar_heads(code, start, expect)
    if sim:
        # 2026-09-17 夜: 月次集計は InStr(s, "（") で括弧から後ろを落として比べ「支出額（税抜）」を「支出額」にした＝含む、だけの話ではない
        out.append("罠: この表には、探している語を含む別の列があります: "
                   + "・".join(f"「{w}」と「{h}」" for w, h in sim)
                   + "。見出しを「含む」（InStr・Like）で探したり、括弧から後ろを落として比べたりすると、こちらを拾います。"
                   "落とすのは空白・改行と単位の括弧（（円）（千円）（人）（計））だけにして、語が一致する列を探す")
    dangling = [ln.strip() for ln in code.splitlines()
                if re.search(r'(&|\+|,)\s*$', ln) and not ln.lstrip().startswith("'") and not re.search(r'\s_\s*$', ln)]
    if dangling:
        # 2026-09-18: パワークエリの M の式を「"…" & Chr(13) & Chr(10) &」で行をまたいで足し、行末の「 _」が無く構文エラー
        out.append(f"罠: 行の終わりが & ／ , で終わり、行の継続の「 _」がありません（例「{dangling[0][:60]}」）。次の行へ続けるなら"
                   "行末に「 _」（空白＋アンダースコア）。長い M の式は m = m & \"…\" & vbCrLf を 1 行ずつ足す（継続は 24 行まで）")
    big_chr = [int(m.group(1)) for m in re.finditer(r'(?<![A-Za-z])Chr\s*\(\s*(\d+)\s*\)', code, re.I)
               if int(m.group(1)) > 255]
    if big_chr:
        # 2026-09-18: 按分が全角スペースを Chr(12288) と書き、日本語の Excel ではそれが「0」＝金額から 0 が全部消えた
        # （2204000 → 224）。実測: Chr(12288)="0"・ChrW(12288)="　"
        out.append(f"罠: Chr は日本語の Excel では ANSI（cp932）で解釈され、Chr({big_chr[0]}) は思った字になりません"
                   "（Chr(12288) は全角スペースではなく「0」＝Replace で値から 0 が全部消えます）。"
                   "全角の字は ChrW で書く（ChrW(12288)＝全角スペース・ChrW(160)＝NBSP）か、\"　\" とそのまま書く")
    if re.search(r'StrConv\s*\(', code, re.I) and start:
        # 2026-09-18: 比較で「補償補塡及び賠償金」が「補償補?及び賠償金」になった（StrConv は ANSI＝cp932 を通る）
        bad = None
        for vr in (start.get('values') or []):
            for v in (vr or []):
                if isinstance(v, str):
                    try:
                        v.encode('cp932')
                    except UnicodeEncodeError:
                        bad = v
                        break
            if bad:
                break
        if bad:
            out.append(f"罠: この表には cp932（Shift-JIS）に無い字があります（例「{bad}」）。StrConv（vbNarrow・vbWide）は"
                       "ANSI を通るので、その字が「?」に変わります。項目の名前や見出しを StrConv で通さない"
                       "（前後の空白だけ Trim・Replace で落とす）")
    if re.search(r'StrConv\s*\([^\n]*vbNarrow', code, re.I) and re.search(r'"[^"\n]*[ァ-ヶー][^"\n]*"', code):
        out.append("罠: StrConv(…, vbNarrow) はカタカナも半角にします（「コード」→「ｺｰﾄﾞ」）。カタカナの語と比べると一致しません。"
                   "全角の数字・英字・記号だけ半角にするなら、その字だけ Replace で直すか、比べる語の側も同じ StrConv を通す")
    # 2026-09-17 夜: 金額の半分が文字（'12000'）の表で、集計を =SUMIF の式で書き 2 往復とも少なく出た
    # グラフの API の罠（2026-09-17 深夜: 課別棒グラフが sr.GapWidth = 60 で 438 を 5 往復くり返した）
    if re.search(r'(?<!ChartGroups\(1\))(?<!ChartGroups\(1\) )\.\s*GapWidth\b', code) and not re.search(r'ChartGroups\s*\(\s*\d+\s*\)\s*\.\s*GapWidth', code):
        out.append("罠: GapWidth（棒の間隔）は系列（Series）には無く、グラフのグループの設定です: cht.ChartGroups(1).GapWidth = 60")
    no_set = [ln.strip() for ln in code.splitlines()
              if re.match(r'^\s*(?!Set\b|Let\b|Const\b|Dim\b)\w+\s*=\s*(ActiveWorkbook|ActiveSheet|Workbooks\s*\(|Worksheets\s*\(|Sheets\s*\(|'
                          r'[\w.()"]+\.(Worksheets|Sheets|ListObjects|PivotTables|ChartObjects|PivotCaches|Queries)\s*\()', ln, re.I)
              and not re.search(r'\.(Name|Count|Value|Index)\b', ln)]
    if no_set:
        # 2026-09-18: パワークエリのクロス集計で「wb = ActiveWorkbook」（Set なし）の 91
        out.append(f"罠: オブジェクト（ブック・シート・テーブル・ピボット）を変数に入れるときは Set が要ります（例「{no_set[0][:60]}」→ Set を付ける）")
    if re.search(r'SourceData', code) and re.search(r'\.Range\s*\(\s*[A-Za-z_]\w*\s*\)', code):
        # 2026-09-18: ピボットの元の範囲の文字（'支出明細'!R1C1:R61C5）を Range() に渡して Nothing になり、全部のピボットを飛ばした
        out.append("罠: PivotTable.SourceData／PivotCache.SourceData は R1C1 形式の文字（'支出明細'!R1C1:R61C5）です。Range() にそのまま"
                   "渡せません。A1 にするなら Application.ConvertFormula(文字, xlR1C1, xlA1)、範囲を変えるなら "
                   "pt.ChangePivotCache wb.PivotCaches.Create(xlDatabase, 新しい範囲の文字) → pt.RefreshTable")
    if re.search(r'PlotArea\s*\.\s*Format\s*\.\s*TextFrame2', code, re.I):
        # 2026-09-17 深夜: 推移の折れ線が ch.PlotArea.Format.TextFrame2 で -2147467259（メソッドは失敗しました）
        out.append("罠: PlotArea には文字の書式（TextFrame2）がありません。グラフ全体の文字は ch.ChartArea.Format.TextFrame2.TextRange.Font.Name")
    # With ブロックの中の「.Size = 14」も数える（2026-09-18: With .ChartTitle.Characters.Font … .Size = 14 を読めず、
    # 全体のフォントを後で当ててタイトルが 18 に戻る罠を 2 往復言えなかった）
    t_size, a_font, stack = [], [], []
    for i, ln in enumerate(code.splitlines()):
        m = re.match(r'\s*With\s+(.+?)\s*$', ln, re.I)
        ctx = ' '.join(stack)
        if re.search(r'ChartTitle', ctx + ' ' + ln, re.I) and re.search(r'(?:Font\s*\.\s*)?(?:Size|Bold)\s*=', ln):
            t_size.append(i)
        if re.search(r'ChartArea', ctx + ' ' + ln, re.I) and re.search(r'(?:Font\s*\.\s*)?Name\s*=', ln):
            a_font.append(i)
        if m:
            stack.append(m.group(1))
        elif re.match(r'\s*End\s+With', ln, re.I) and stack:
            stack.pop()
    if t_size and a_font and max(a_font) > min(t_size):
        # 2026-09-18 未明: 見やすいグラフでタイトルを 14 にした後に ChartArea の文字を Meiryo UI にし、タイトルが 18 に戻った
        out.append("罠: ChartArea（グラフ全体）の文字の書式を当てると、タイトルの文字の大きさ・太字も既定に戻ります。"
                   "タイトルの大きさと太字は ChartArea の文字の書式の後で当てる")
    if re.search(r'\bSeries\w*\.\s*(ChartTitle|HasLegend|Legend)\b', code):
        out.append("罠: ChartTitle・HasLegend・Legend はグラフ（Chart）の設定で、系列（Series）には無い")
    if re.search(r'Queries\.Add', code, re.I) and re.search(r'"\[\s*"\s*&|\(\[\s*"\s*&|"\[\{?"?\s*&\s*\w+\s*&\s*"\]', code) \
            and not re.search(r'\[#""', code):
        # 2026-09-18 第二期: 見出し「金額（円）」を M の式で [金額（円）] と書くと式が読めず「Source を認識できません」
        out.append("罠: M の式で列の名前を [" + '"' + " & 見出し & " + '"' + "] と囲むと、括弧・空白・記号の入った見出し（「金額（円）」）で式が壊れます。"
                   "[#\"見出し\"] の形で囲む（VBA では \"[#\"\"\" & 見出し & \"\"\"]\"）。列の一覧 {\"…\"} の中は \"見出し\" のままでよい")
    if re.search(r'ChrW\s*\(\s*&H339[CE]\s*\)', code, re.I):
        # 2026-09-18 第二期: 名寄せで ㈱ を ChrW(&H339E)（㎞）・㈲ を ChrW(&H339C)（㎜）と書いた
        out.append("罠: ChrW(&H339E) は「㎞」、ChrW(&H339C) は「㎜」です。㈱＝ChrW(&H3231)・㈲＝ChrW(&H3232)。どちらも cp932 にある字なので \"㈱\" \"㈲\" と字のまま書いてよい")
    if re.search(r'Sparkline', code, re.I) and re.search(r'\.\s*Line\s*\.\s*(Color|ForeColor)|\.\s*Format\s*\.\s*Line', code, re.I):
        # 2026-09-18 第二期: スパークラインの色を sg.Line.Color.RGB と書いて 438
        out.append("罠: スパークライン（SparklineGroup）に Line・Format はありません。線の色は sg.SeriesColor.Color = RGB(…)、"
                   "最高点の印は sg.Points.Highpoint.Visible = True（色は sg.Points.Highpoint.Color.Color）")
    if re.search(r'ChrW\s*\(\s*(9652|9653|&H25B4|&H25B5)\s*\)', code, re.I):
        # 2026-09-18: 数値直しで ▲ を ChrW(9652)（▴ 小さい三角）と書き、▲ の数が負にならず 2 往復外れた
        out.append("罠: ChrW(9652) は ▴（小さい三角）で ▲ ではありません。▲＝ChrW(9650)・△＝ChrW(9651)。"
                   "どちらも cp932 にある字なので \"▲\" \"△\" と字のまま書いてよい")
    if (re.search(r"'\s*選ぶ列:\s*[1-9]", code) and re.search(r'For\s+Each\s+\w+\s+In\s+[\w.]*(Selection|sel\w*|area\w*|\.Areas\b|\.Cells\b)', code, re.I)
            and not re.search(r'End\s*\(\s*xlUp|Rows\.Count|UsedRange|CurrentRegion|SpecialCells', code, re.I)):
        # 2026-09-18: 数値直しで For Each cell In Selection（見出しのセル 1 つ）だけを回し、本文に何も書かず 6 往復外れた
        out.append("罠: 選んでいるのは見出しのセル（1 列に 1 つ）だけです。For Each で Selection のセルを回すと見出しだけで終わります。"
                   "列ごとに、見出しの次の行（Selection.Cells(1,1).Row + 1）から、その列の最後の行"
                   "（ws.Cells(ws.Rows.Count, 列).End(xlUp).Row）まで 1 行ずつ回す")
    if not start:
        return out
    kinds = start.get('kinds') or []
    vals = start.get('values') or []
    # 2026-09-17 夜: Web の画面から貼ったコードの前後の NBSP を Trim が消さず、対応表に当たらなかった
    nb = next((str(v) for vr in vals for v in (vr or []) if isinstance(v, str) and ' ' in v), None)
    if nb is not None and not re.search(r'ChrW\s*\(\s*(160|&HA0)\s*\)', code, re.I):
        # 2026-09-17 夜: AI が Replace(ss, Chr(160), "") と書き 2 往復外れた。日本語の Excel（Shift-JIS）の Chr(160) は NBSP ではない
        out.append(f"罠: この表には NBSP（Web の画面から貼った空白・U+00A0）が入っています（例「{nb.replace(chr(160), '␣')}」）。"
                   "Trim・Replace(\" \") では消えません。日本語の Excel では Chr(160) も NBSP になりません。"
                   "Replace(v, ChrW(160), \"\") と ChrW で書く")
    # 先頭が 0 の文字（伝票番号・コード '000123'）は数として足す値ではない＝数えない
    texts = [str(v) for kr, vr in zip(kinds, vals) for k, v in zip(kr, vr)
             if k == 's' and _TEXT_NUMBER_RE.match(str(v)) and not re.match(r'^\s*[0０][\d０-９]', str(v))]
    if not texts:
        return out
    if re.search(r'\b(SUMIFS?|SUMPRODUCT|SUBTOTAL)\s*\(|WorksheetFunction\.Sum', code, re.I):
        out.append("罠: SUMIF・SUMIFS・SUM の式は、文字として入った数字（'12000'・'12,000円'・全角）を足しません（0 扱い）。"
                   "この表には文字の数字があります。明細を 1 行ずつ読んで数に直して（全角・カンマ・円・△）足し、値で書く")
    marks = [t for t in texts if re.search(r'[円△▲]', t)]
    # 円を Replace で消してから数にしているマクロには言わない（2026-09-18: 数値直しで外れた見立てを 4 往復くり返させた）
    if (marks and re.search(r'\b(IsNumeric|Val|CDbl|CLng|CCur)\s*\(', code, re.I)
            and not re.search(r'Replace\s*\([^\n]*"円"', code)):
        out.append(f"罠: この表の数字の文字には「{marks[0].strip()}」のように円・△が付いたものがあります。IsNumeric は False、"
                   "Val は 0 を返す＝読み飛ばしています。円・カンマ・空白を消し、△/▲ は負の数にし、全角を半角にしてから数にする")
    if re.search(r'\bVal\s*\(', code, re.I) and any(re.search(r'\d[,，]\d', t) for t in texts):
        out.append("罠: Val(\"1,200\") は 1 です（カンマで止まる）。カンマを消してから Val にするか CDbl を使う")
    return out


def _drop_sub(text, name):
    """モジュールの本文から Sub name の塊（Sub 〜 End Sub）を抜く（純 Python）。無ければそのまま。"""
    pat = re.compile(r'^[ \t]*(?:Public\s+|Private\s+)?Sub\s+' + re.escape(name) + r'\s*\(.*?^[ \t]*End\s+Sub[ \t]*\r?\n?',
                     re.M | re.S | re.I)
    return pat.sub('', text, count=1)


def _tidy_without(sub_name):
    """先撃ちの本体（表の整理.bas）から、いま鍛えている Sub と同じ名前の塊を抜いた写し（2026-09-17 夕）。
    表の整理.bas には鍛えて登録した Sub も入る＝同じ名前を鍛え直すと、試しのブックで名前が重なりコンパイルが落ちる。"""
    import tempfile
    with open(_TIDY_BAS, encoding='cp932') as f:
        text = f.read()
    cut = _drop_sub(text, sub_name)
    if cut == text:
        return _TIDY_BAS
    path = os.path.join(tempfile.gettempdir(), '_forge_tidy_without.bas')
    with open(path, 'w', encoding='cp932', newline='') as f:
        f.write(cut)
    return path


def _compile_scratch(xl, mwb):
    """試しのブックを一時の .xlsm に保存して全体コンパイル。→ (落ちたら文 or None, 保存したパス or None)。
    VBE がプロジェクトをファイル名で探すので、保存していないブックは見つからない＝先に保存する。
    保存やコンパイルが押せないときは None（撃つ側の時間切れに任せる＝見られないことを落ちたことにしない）。"""
    import tempfile
    from vbam_vba import _compile_vbproject
    path = os.path.join(tempfile.gettempdir(), f"_forge_macros_{os.getpid()}_{int(time.time() * 1000)}.xlsm")
    try:
        mwb.SaveAs(path, FileFormat=52)
    except Exception:
        return None, None
    try:
        res = _compile_vbproject(xl, mwb) or {}
    except Exception:
        return None, path
    if res.get('ok') is False:
        return (f"コンパイルエラーで撃てません: {res.get('detail')}"
                "（呼んでいる Function・Sub が無い／変数名・定数の打ち間違い）"), path
    return None, path


def _watchdog(pid, seconds):
    """seconds 後にまだ動いていれば、その Excel（自分の台）を落とすタイマー。→ (timer, 落としたかの箱)。"""
    import threading
    fired = {'hit': False}

    def _kill():
        if not pid:
            return
        fired['hit'] = True
        with contextlib.suppress(Exception):
            import subprocess
            subprocess.run(['taskkill', '/PID', str(pid), '/F'], capture_output=True, timeout=20)
    t = threading.Timer(seconds, _kill)
    t.daemon = True
    return t, fired


_ENTRY_CHECK = True           # 値が合った写しに、入口と同じ仕上げ（tidy）と検査を当てる


def _entry_notes(wb, ws, sheet, before_va):
    """登録簿のマクロを入口で撃ったあとと同じ仕上げと検査を写しに当て、AI に回す指摘を返す（COM）。無ければ []。
    2026-09-17 夜: 前年度比較は鍛えて合格したのに、入口では合計行の式が「横にコピーできない」と止められ毎回 AI に回った。
    鍛える回路の合格条件に入口の検査が入っていなかった。"""
    import io
    import vbam_prefire as vp
    try:
        targets = vp._written_anchors(ws, before_va)
        if not targets:
            return []
        with contextlib.redirect_stdout(io.StringIO()):
            va._run_cmd(['tidy', '--sheet', sheet] + targets, wb)
            audit = va._audit_now(wb, sheet, targets)
            c_block, _noticed = va._content_now(wb, sheet, targets)
        audit, c_block, _old = vp._split_old_findings(audit, c_block, before_va, ws, sheet)
        return list(audit) + list(c_block)
    except Exception:
        return []                                         # 検査が撃てないことを外れにしない


def _run_on_copies(case, bas_path, sub_name, targets):
    """写しを自分の Excel で開き、先撃ちのマクロ（あれば）→ 鍛えたマクロ の順に撃って、値を読む。
    targets＝[(撃つ前のブック, 期待の snap)]。1 つの Excel で順に撃つ。→ ([不一致の行], 秒)。
    実行時エラーは台（鍛冶_撃つ）が受け止めて文で返す。_RUN_TIMEOUT 秒を越えたら自分の台を落として打ち切る。"""
    import tempfile
    import vbam_core
    xl, pid = _start_own()
    mwb = None
    macro_file = None
    t0 = time.time()
    results = []
    names = list(case.get('prefire') or []) + [sub_name]
    exits, traps, marked, text_code = {}, [], None, ''
    with contextlib.suppress(OSError, UnicodeError):
        with open(bas_path, encoding='cp932') as f:
            text = f.read()
        text_code = text
        traps = _trap_hints(text)
        inst, exits = _instrument_exits(text, sub_name)
        if inst != text:
            marked = os.path.join(tempfile.gettempdir(), f"_forge_marked_{os.getpid()}_{int(time.time() * 1000)}.bas")
            with open(marked, 'w', encoding='cp932', newline='') as f:
                f.write(inst)
    last_exit = max(exits) if exits else 0
    try:
        try:
            mwb = xl.Workbooks.Add()
            if case.get('prefire') and os.path.isfile(_TIDY_BAS):
                mwb.VBProject.VBComponents.Import(_tidy_without(sub_name))
            mwb.VBProject.VBComponents.Import(marked or bas_path)
            harness = mwb.VBProject.VBComponents.Add(1)
            harness.Name = _HARNESS
            harness.CodeModule.AddFromString(_harness_code(names, track=bool(marked)))
        except Exception as ex:
            return [[f"マクロを読み込めませんでした: {ex}"] for _ in targets], time.time() - t0
        # 撃つ前に全体コンパイル（コンパイルエラーは Run すると見えない窓で止まる＝時間切れと区別がつかない）
        comp, macro_file = _compile_scratch(xl, mwb)
        if comp:
            return [[comp] + traps for _ in targets], time.time() - t0
        for k, target in enumerate(targets):
            before, expect = target[0], target[1]
            sheet = (target[2] if len(target) > 2 else None) or case['sheet']
            sel = target[3] if len(target) > 3 else None
            wb = None
            # プロセス番号も名前に入れる（2026-09-17 夜: 修理の試験を 2 本並べて撃ったとき、同じミリ秒の写しの名前が
            # ぶつかった疑いで判定が 1 枚だけ「シートが無い」で落ちた）
            tmp = os.path.join(tempfile.gettempdir(),
                               f"_forge_try_{os.getpid()}_{int(time.time() * 1000)}_{k}{os.path.splitext(before)[1]}")
            timer, fired = _watchdog(pid, _RUN_TIMEOUT)
            try:
                shutil.copyfile(before, tmp)
                wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
                ws = wb.Sheets(sheet)
                wb.Activate()
                ws.Activate()
                if sel:
                    with contextlib.suppress(Exception):
                        ws.Range(sel).Select()               # 列を選ぶ仕事（' 選ぶ列: N）は、人が選んだ列を再現してから撃つ
                start = _snapshot_texts(ws)
                before_va = None
                if _ENTRY_CHECK and not case.get('prefire'):
                    with contextlib.suppress(Exception):
                        before_va = va._sheet_snapshot(ws)
                timer.start()
                got = str(xl.Run(f"'{mwb.Name}'!{_HARNESS}.鍛冶_撃つ") or '')
                timer.cancel()
                if got.startswith('ERR|'):
                    _e, step, num, desc = (got.split('|', 3) + ['', '', ''])[:4]
                    err_line = ''
                    if marked and '|' in desc:
                        desc, err_line = desc.rsplit('|', 1)
                    hint = ("（CreateObject は Scripting.Dictionary だけ）" if num in ('429', '-2146232576') else
                            "（ThisWorkbook でなく ActiveWorkbook・シートや列が見つからないときの Nothing を確かめる）"
                            if num in ('91', '9') else "")
                    charts = 0
                    with contextlib.suppress(Exception):
                        charts = int(wb.Charts.Count)
                    if (num == '13' and charts and
                            re.search(r'For\s+Each\s+\w+\s+In\s+[\w.()]*\bSheets\b', text_code or '', re.I)):
                        # 2026-09-17 深夜: 課のシートの間にグラフシートがあるブックで、集約が 5 往復とも 13 で止まった
                        hint += ("（罠: このブックにはグラフシートがあります。Sheets にはグラフシート（Chart）も入るので、"
                                 "As Worksheet の変数で For Each Sheets を回すと 13 型が一致しません。Worksheets で回す）")
                    at = ''
                    if err_line.isdigit() and int(err_line):
                        src = _sub_lines(text_code, sub_name)
                        n_line = int(err_line)
                        if 0 < n_line <= len(src):
                            at = f"（Sub の行から {n_line} 行目「{src[n_line - 1].strip()[:120]}」で止まりました）"
                    # 字面の罠の知らせは実行時エラーのときにも付ける（2026-09-18 未明: 見やすいグラフで sr.GapWidth の 438 がまた出た）
                    results.append([f"マクロ {step} が実行時エラーで止まりました: {num} {desc}{at}{hint}"] + list(traps))
                    continue
                after = _snapshot_texts(ws)
                if 'page' in expect:
                    with contextlib.suppress(Exception):
                        after['page'] = _page_of(ws, always=True)
                mism = _compare(expect, after)
                exit_at = got.split('|', 1)[1] if got.startswith('OK|') else ''
                where = (f"Sub の行から {exit_at} 行目「{exits[int(exit_at)]}」の Exit Sub で抜けました"
                         if exit_at.isdigit() and int(exit_at) in exits else '')
                if mism and after.get('values') == start.get('values') and after.get('objects') == start.get('objects'):
                    # 印つきの写しで「OK|」（抜けた行が空）＝Exit Sub でなく End Sub まで走った（2026-09-18: ピボットの更新で
                    # Range(R1C1 の文字) が On Error Resume Next に飲まれ、全部のピボットを飛ばしたのに「Exit Sub した疑い」と返した）
                    ran_to_end = bool(marked) and got.startswith('OK|') and not exit_at
                    mism.insert(0, "マクロは表に何も書かずに終わりました（"
                                + (where or ("End Sub まで走りましたが何も変えていません＝条件で全部飛ばしたか、On Error Resume Next で"
                                             "失敗を飲み込んだ疑い。On Error Resume Next の範囲を 1 行に絞ると止まった行が返ります"
                                             if ran_to_end else "見出し・シート・列が見つからず Exit Sub した疑い")) + "）")
                elif mism and where and int(exit_at) != last_exit:
                    mism.insert(0, f"途中で抜けています: {where}")
                found = traps + [x for x in _trap_hints(text_code, start, expect) if x not in traps] if mism else []
                if found:
                    mism[1:1] = found
                if not mism and before_va is not None:
                    notes = _entry_notes(wb, ws, sheet, before_va)
                    if notes:
                        mism = [f"値は正解と同じですが、入口の検査が止めます（このままだと登録しても毎回 AI に回ります）: {n}"
                                for n in notes[:4]]
                results.append(mism)
            except Exception as ex:
                timer.cancel()
                if fired['hit']:
                    results.append([f"マクロが {_RUN_TIMEOUT} 秒で終わらず、止めました（無限ループの疑い）"])
                    results += [["前の表でマクロが止まったので撃てませんでした"] for _ in targets[k + 1:]]
                    mwb = None
                    return results, time.time() - t0
                results.append([f"マクロの実行で止まりました: {ex}"])
            finally:
                with contextlib.suppress(Exception):
                    if wb is not None and not fired['hit']:
                        wb.Close(SaveChanges=False)
                with contextlib.suppress(Exception):
                    os.remove(tmp)
        return results, time.time() - t0
    finally:
        with contextlib.suppress(Exception):
            if mwb is not None:
                mwb.Close(SaveChanges=False)
        xl = None
        _quit(pid)
        for p in (macro_file, marked):
            if p:
                with contextlib.suppress(OSError):
                    os.remove(p)


def rehearse_steps(module_text, names, book_path, sheet, module='表の整理', sel_cols=None):
    """人の Excel で撃つ前の試し撃ち（2026-09-17 夕）。自分の Excel（非表示）で、モジュールの本文を新しいブックに入れ、
    book_path（撃つブックの写し）の sheet に names を 1 つずつ撃つ。→ [{'name','ok','changed','why'}]（撃てた所まで）。

    なぜ: 登録したマクロが想定外の表で実行時エラーを起こすと、人の Excel にデバッグの窓が出て agent が固まる
    （Application.Run は呼んだ側の On Error で受けられない＝別ブックの台で包めない・実測）。写しで先に撃てば、
    エラー・時間切れ・何も変えずに抜けた、を人の Excel に触る前に知れる。"""
    import tempfile
    import vbam_core
    out = []
    work = tempfile.mkdtemp(prefix='_prefire_rehearse_')
    bas = os.path.join(work, module + '.bas')
    err = _write_bas(bas, module, module_text)
    if err:
        return [{'name': names[0] if names else '', 'ok': False, 'changed': False, 'why': err}]
    xl, pid = _start_own()
    mwb = wb = None
    macro_file = None
    try:
        try:
            mwb = xl.Workbooks.Add()
            mwb.VBProject.VBComponents.Import(bas)
            h = mwb.VBProject.VBComponents.Add(1)
            h.Name = _HARNESS
            h.CodeModule.AddFromString("".join(_harness_code([nm]).replace('鍛冶_撃つ()', f'鍛冶_撃つ{k}()')
                                               .replace('鍛冶_撃つ =', f'鍛冶_撃つ{k} =')
                                               for k, nm in enumerate(names, 1)))
        except Exception as ex:
            return [{'name': names[0] if names else '', 'ok': False, 'changed': False, 'why': f"読み込めませんでした: {ex}"}]
        comp, macro_file = _compile_scratch(xl, mwb)
        if comp:
            return [{'name': names[0] if names else '', 'ok': False, 'changed': False, 'why': comp}]
        wb = xl.Workbooks.Open(book_path, UpdateLinks=0)
        ws = wb.Sheets(sheet)
        wb.Activate()
        ws.Activate()
        for k, nm in enumerate(names, 1):
            before = _snapshot_texts(ws)
            with contextlib.suppress(Exception):
                before['page'] = _page_of(ws)
            if (sel_cols or {}).get(nm):
                import vbam_prefire as vp
                vp.select_columns(ws, sel_cols[nm])          # 列を選ぶ仕事は、選んでから試し撃ち（2026-09-18）
            timer, fired = _watchdog(pid, _RUN_TIMEOUT)
            timer.start()
            try:
                got = str(xl.Run(f"'{mwb.Name}'!{_HARNESS}.鍛冶_撃つ{k}") or '')
            except Exception as ex:
                timer.cancel()
                why = (f"{_RUN_TIMEOUT} 秒で終わらず止めました（無限ループの疑い）" if fired['hit'] else f"止まりました: {ex}")
                out.append({'name': nm, 'ok': False, 'changed': False, 'why': why})
                if fired['hit']:
                    mwb = wb = None
                return out
            timer.cancel()
            if got.startswith('ERR|'):
                _e, _step, num, desc = (got.split('|', 3) + ['', '', ''])[:4]
                out.append({'name': nm, 'ok': False, 'changed': False, 'why': f"実行時エラー {num} {desc}"})
                return out
            after = _snapshot_texts(ws)
            with contextlib.suppress(Exception):
                after['page'] = _page_of(ws)
            # 2026-09-18: グラフ・ピボット・クエリを作るマクロはセルの値を変えない＝「何も変えずに終わった」と撃たずに AI へ回していた
            # 2026-09-18 第二期: シートを足す（1 件 1 枚）・印刷設定だけの仕事も同じ＝ほかのシートと印刷設定の変化も見る
            out.append({'name': nm, 'ok': True, 'why': '',
                        'changed': (after.get('values') != before.get('values') or after.get('objects') != before.get('objects')
                                    or after.get('book') != before.get('book') or after.get('page') != before.get('page')
                                    or after.get('fx') != before.get('fx') or after.get('filter') != before.get('filter'))})
        return out
    finally:
        with contextlib.suppress(Exception):
            if wb is not None:
                wb.Close(SaveChanges=False)
        with contextlib.suppress(Exception):
            if mwb is not None:
                mwb.Close(SaveChanges=False)
        xl = None
        _quit(pid)
        with contextlib.suppress(Exception):
            shutil.rmtree(work, ignore_errors=True)
        if macro_file:
            with contextlib.suppress(OSError):
                os.remove(macro_file)


def _try_macro(case, bas_path, sub_name):
    """弾の表で試す。→ (不一致の行, 秒)。"""
    res, sec = _run_on_copies(case, bas_path, sub_name, [(case['before'], case['expect'], None, case.get('sel'))])
    return res[0], sec


def _try_tests(case, bas_path, sub_name):
    """別の表（--test）で試す。→ ([(名前, 不一致の行)], 秒)。無ければ ([], 0)。"""
    tests = case.get('tests') or []
    if not tests:
        return [], 0.0
    res, sec = _run_on_copies(case, bas_path, sub_name,
                              [(t['before'], t['expect'], t.get('sheet'), t.get('sel')) for t in tests])
    return [(t['label'], m) for t, m in zip(tests, res)], sec


_MAX_CONTEXT_ROWS = 4          # 外れた行の中身を AI に見せる行数


def _context_rows(expect, mism, n=_MAX_CONTEXT_ROWS):
    """外れたセルのある行を、正解の表から見出しつきで抜く（純 Python）。
    2026-09-17: 表の最後の合計行（伝票番号が空）に「新に無し」と書く外れが 2 往復続いた。AI には番地（E22）と
    表の頭 10 行しか渡っておらず、22 行目が合計行だと読めなかった。"""
    vals = (expect or {}).get('values') or []
    r0 = int((expect or {}).get('row') or 1)
    rows = []
    for line in mism or ():
        m = re.match(r'^\s*[A-Z]{1,3}(\d+):', str(line))
        if m:
            r = int(m.group(1))
            if r not in rows:
                rows.append(r)
    head_hit = r0 in rows and bool(vals)
    rows = [r for r in rows if 0 < r - r0 < len(vals)][:n]           # 頭の行は下で必ず出す（2 回並べない）
    if not rows and not head_hit:
        return ""
    out = [f"外れた行の正解（左上 {_addr_of(r0, int(expect.get('col') or 1), 0, 0)} から・見出しは {r0} 行目）:",
           f"{r0}\t" + "\t".join('' if v is None else str(v) for v in vals[0])]
    for r in rows:
        out.append(f"{r}\t" + "\t".join('' if v is None else str(v) for v in vals[r - r0]))
    return "\n".join(out)


_TOTAL_WORDS = ('合計', '総計', '小計', '計', '累計')
_SIZE_RE = re.compile(r'期待 (\d+)×(\d+).*?マクロ後 (\d+)×(\d+)')
_CELL_RE = re.compile(r"^\s*([A-Z]{1,3})(\d+): 期待 '(.*)' ／ マクロ後 '(.*)'$")
_DATE_TEXT_RE = re.compile(r'^\d{4}/\d{1,2}/\d{1,2}( \d{1,2}:\d{2}(:\d{2})?)?$')


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _hints(mism, expect=None):
    """外れ方 → 原因の型の手がかり（純 Python）。AI が番地の列だけから読み解けなかった型を言葉にする。
    2026-09-17: 集約で課のシートの「合計」の行を明細として拾い、同じ外れを 2 往復くり返した。"""
    out = []
    got_total = expect_total = head_total = got_note = False
    spaced_total = None
    moved = baked = zero = 0
    exp_at, got_at = {}, {}
    for line in mism or ():
        m = _CELL_RE.match(str(line))
        if not m:
            continue
        exp, got = m.group(3).strip(), m.group(4).strip()
        rc = (int(m.group(2)), _col_index(m.group(1)))
        exp_at[rc], got_at[rc] = exp, got
        note = got.startswith('※') and not exp.startswith('※')
        # 空白を詰めて比べる（2026-09-17 夜: 課のシートの「小　計」「合　計」を事業として拾った外れに、合計の手がかりが出なかった）
        got_t, exp_t = re.sub(r'\s', '', got) in _TOTAL_WORDS, re.sub(r'\s', '', exp) in _TOTAL_WORDS
        if note or (got_t and not exp_t):
            if note:
                got_note = True
            else:
                got_total = True
                if got != re.sub(r'\s', '', got):
                    spaced_total = got
            if expect and rc[0] == int(expect.get('row') or 1):
                head_total = True
        if exp_t and not got_t and not note:
            expect_total = True
        # 2026-09-17 夜: 帳票の受付番号「R8-001」を書くと Excel が令和 8 年 1 月の日付に読み替え、AI は 2 往復気づかなかった
        if exp and not _DATE_TEXT_RE.match(exp) and _DATE_TEXT_RE.match(got):
            baked += 1
        elif re.fullmatch(r'0\d+', exp) and got == exp.lstrip('0'):
            zero += 1
        if exp and got and exp != got:
            moved += 1
    typed = sum(1 for line in mism or () if '型が違う 期待' in str(line))
    if typed:
        out.append("・数や日付を文字で書いています（見た目は同じでも SUM・並べ替え・絞り込みが効かない）。数・日付は .Value を"
                   "そのまま書き、NumberFormat \"@\" や頭の ' は化ける文字の列（R8-001・0012 など）だけに使う。"
                   "逆に、期待が文字の所に数を書いていないか")
    if baked:
        out.append("・文字が日付に化けています（R8-001・1-2・4/1 などは Excel が日付と読む）。その列は書く前に "
                   "NumberFormat を \"@\" にするか、値の頭に ' を付けて書く（元のセルの .Value を .Text で読む・日付の列は日付のまま）")
    if zero:
        out.append("・先頭の 0 が落ちています（0012 が 12）。書く前に NumberFormat を \"@\" にするか、値の頭に ' を付けて書く")
    # 2026-09-17 夜: 月次の集計表を、前の集計が残る表で 2 列右に書いた。「行ごとにずれ」と言っても列のずれは伝わらない
    shift = {}
    for (r, c), g in got_at.items():
        if not g:
            continue
        for k in range(-30, 31):
            if k and exp_at.get((r, c - k)) == g:
                shift[k] = shift.get(k, 0) + 1
    shifted = False
    if shift:
        k, n = max(shift.items(), key=lambda kv: kv[1])
        if n >= 3:
            out.append(f"・書いた場所が {abs(k)} 列{'右' if k > 0 else '左'}にずれています（表の右端を、前の結果の残り・"
                       "別の表・同じ見出しの 2 つ目まで数えている／書き始めの列の決め方）")
            moved = 0
            shifted = True
    # 行のずれ（見出しの行の取り違え・2026-09-18: 表題の行を見出しにして 1 行目に書いた）
    shift_row = {}
    ev = (expect or {}).get('values') or []
    er0, ec0 = int((expect or {}).get('row') or 1), int((expect or {}).get('col') or 1)

    def _exp_grid(r, c):
        i, j = r - er0, c - ec0
        return str(ev[i][j]).strip() if 0 <= i < len(ev) and 0 <= j < len(ev[i] or ()) else ''
    for (r, c), g in got_at.items():
        if not g:
            continue
        for k in range(-6, 7):
            if k and (exp_at.get((r - k, c)) == g or (ev and _exp_grid(r - k, c) == g)):
                shift_row[k] = shift_row.get(k, 0) + 1
                break
    if shift_row and not shifted:
        k, n = max(shift_row.items(), key=lambda kv: kv[1])
        if n >= 2:
            out.append(f"・書いた行が {abs(k)} 行{'下' if k > 0 else '上'}にずれています（見出しの行の決め方＝表題や"
                       "「単位：円」の行を見出しにしていないか。見出しの行は値が 2 つ以上並ぶ行・本文のすぐ上の行）")
    # 値が同じ倍率でずれている（2026-09-18: 按分で「基準の合計」を別の表から取り、全部 1/10000 になった）
    ratios = []
    for rc, g in got_at.items():
        e = exp_at.get(rc)
        try:
            ge, gg = float(str(e).replace(',', '')), float(str(g).replace(',', ''))
        except ValueError:
            continue
        if gg and ge and abs(ge - gg) > 0.5:
            ratios.append(ge / gg)
    if len(ratios) >= 3:
        near = max(((r, sum(1 for x in ratios if abs(x - r) <= abs(r) * 0.02)) for r in ratios), key=lambda t: t[1])
        if near[1] >= 3 and abs(near[0] - 1) > 0.05:
            out.append(f"・値が同じ倍率でずれています（期待 ÷ マクロ後 ≒ {near[0]:.4g} が {near[1]} 個）＝"
                       "割る数・掛ける数の取り方が違います（基準の合計を別の表・別の列から取っていないか。"
                       "単位・千円・率の取り違えも見る）")
    # 引く相手が 1 件も見つかっていない（2026-09-18: 突合でキーの列を「文字の多い列」で決め、全行「相手に無し」になった）
    from collections import Counter
    same = Counter(g for rc, g in got_at.items() if g and exp_at.get(rc) and exp_at[rc] != g
                   and not re.fullmatch(r'[\d,.\-]+', g))
    if same:
        v, n = same.most_common(1)[0]
        if n >= 5:
            out.append(f"・「{v}」を {n} 行に書いています＝引く相手が 1 件も見つかっていません。キーの列・相手の表の決め方を"
                       "見直す（キーは「両方の表に同じ値が並ぶ列」＝値を突き合わせて決める。文字の列とは限らない・"
                       "先頭 0 の番号は数の列に見える）")
    # 同じ物を 2 つ作っている（2026-09-18: 2 回目に撃つと、前の結果の隣に同じ列・同じ集計表をもう 1 つ作った）
    extra_vals = [g for rc, g in got_at.items() if g and not exp_at.get(rc)]
    rep = [v for v, n in Counter(extra_vals).items() if n >= 2 and not re.fullmatch(r'[\d,.\-]+', v)]
    if rep:
        out.append(f"・同じ見出し「{rep[0]}」を 2 か所に書いています＝前の結果（同じ見出しの列・表）の隣にもう 1 つ"
                   "作っています。同じ見出しの列・表が既にあればそこに書き直す（増やさない・2 回撃っても同じ形）")
    extra_nums = [rc for rc, g in got_at.items() if not exp_at.get(rc) and re.fullmatch(r'-?[\d,.]+', g or '')]
    if extra_nums and not shifted:
        # 2026-09-18: 合計行で「コード」の列（110,120…）まで足した。依頼に「コードの列は足さない」と書いてあっても、
        # AI は表の語を書き込む決まりを恐れて見出しの語で見分けなかった
        out.append(f"・期待が空のセルに数を書いています（{len(extra_nums)} 個）＝足さない・書かない決まりの列か行に書いています。"
                   "依頼に書いた語（番号・コード・年度など）は、見出しの語で見分けてよい")
    for line in mism or ():
        s = _SIZE_RE.search(str(line))
        if s:
            er, ec, gr, gc = (int(x) for x in s.groups())
            if gr > er:
                out.append(f"・行が {gr - er} 行多い＝入れてはいけない行（合計・小計・空行・見出しのくり返し・前の結果の残り）を拾っています")
            elif gr < er:
                out.append(f"・行が {er - gr} 行足りない＝途中の空行や表題で読むのを止めた・シートや行を読み飛ばしています")
            if gc > ec:
                out.append("・列が多い＝書く場所がずれているか、要らない列を足しています")
            elif gc < ec:
                out.append("・列が足りない＝足すはずの列・表を書いていません")
    # 2026-09-17 夜: 改ページごとに表題（結合セル）をくり返した帳票で、表題を件として一覧に書き、同じ外れで 2 往復止まった
    titles = {str(v).strip() for row in ((expect or {}).get('values') or [])[:3] for v in (row or [])
              if isinstance(v, str) and len(v.strip()) >= 6 and not re.fullmatch(r'[\d,.\-/ ]+', v.strip())}
    hit = next((g for g in got_at.values() if g in titles), None)
    if hit:
        out.append(f"・シートの上の表題「{hit}」を値として書いています＝途中でくり返した表題の行（結合セル）を件や明細として拾っています。"
                   "表題と同じ文字の行・見出しのくり返しの行は入れない")
    if head_total:
        # 2026-09-17 夜: 按分で課別人数の「合計」の行を課として列に並べ、AI は共通経費の側だけ直して 3 往復外れた
        out.append("・見出しの行に「合計」や「※」の注記が出ています＝見出しに並べる一覧（課名など・ほかのシートの表）の「合計」「計」の行まで"
                   "拾っています。その一覧の合計・小計・空行・※注記も除く（明細の表だけでなく）")
    elif got_note:
        out.append("・「※」で始まる注記の行を明細として拾っています（表の下や途中の※の行は入れない）")
    if spaced_total:
        out.append(f"・「{spaced_total}」のように空白の入った合計・小計の行を明細として拾っています。空白（全角も）を詰めてから"
                   "「合計」「小計」「総計」「計」と比べる")
    if got_total and not head_total:
        out.append("・「合計」「計」などの行を明細として扱っています（元の表の合計・小計の行は明細から除く）")
    if expect_total:
        out.append("・正解にある「合計」の行が出ていないか、位置がずれています")
    if moved >= 4 and not got_total:
        out.append("・値が行ごとにずれています＝どこかで 1 行多く／少なく数えています（先頭・末尾・空行の扱い）")
    return ("原因の手がかり:\n" + "\n".join(out)) if out else ""


def _mismatch_key(mism):
    """不一致の行 → 外れたセルの番地の組（純 Python）。値は見ない＝AI が並びを変えても同じセルが外れ続けるなら同じ。
    2026-09-17: 集計の並び順が正解の表 1 枚から読めず、同じ 12 セルが 4 往復続けて外れた（AI 4 回の空回り）。"""
    addrs = set()
    for line in mism or ():
        m = re.match(r'^\s*([A-Z]{1,3}\d+):', str(line))
        if m:
            addrs.add(m.group(1))
    return tuple(sorted(addrs))


def _key_text(key):
    return ("・".join(key[:6]) + (f" ほか {len(key) - 6}" if len(key) > 6 else "")) if key else ""


def _tests_feedback(case, failed):
    """別の表の不一致 → AI に返す文（その表の頭つき）。"""
    out = []
    by = {t['label']: t for t in case.get('tests') or []}
    for label, mism in failed:
        t = by.get(label) or {}
        out.append(f"別の表「{label}」で不一致 {len(mism)} 件（弾の表に無い形がこの表にあります＝直書き・行の拾い方・範囲の決め方を見直す）:")
        out += ["  " + m for m in mism[:_MAX_TEST_MISMATCH]]
        if len(mism) > _MAX_TEST_MISMATCH:
            out.append(f"  …ほか {len(mism) - _MAX_TEST_MISMATCH} 件")
        exp = t.get('expect') or {}
        hint = _hints(mism, exp)
        if hint:
            out.append("  " + hint.replace("\n", "\n  "))
        ctx = _context_rows(exp, mism)
        if ctx:
            out.append("  " + ctx.replace("\n", "\n  "))
        if exp:
            out.append(f"  この表のシート「{exp.get('sheet') or t.get('sheet') or '?'}」の正解（頭）:\n" + _table_text(exp, max_rows=10))
        others = exp.get('others') or []
        if others:
            # 2026-09-17: 突合で「新システム」のシート名・表題の行が違う表に外れ続けた。AI は外れた表のほかのシートを
            # 見られず、同じ直しを 3 往復くり返した＝そのブックのほかのシートの名前と頭も返す
            out.append("  このブックのほかのシート: " + "・".join(f"「{o.get('name')}」" for o in others))
            for o in others:
                out.append(f"  シート「{o.get('name')}」（頭）:\n{o.get('head')}")
    return "\n".join(out)


_LITERAL_PASS_SCORE = 0.5     # 値は全部そろったが表の語を書き込んだ版（頭の行だけの違反 -1 より悪く、外れ 1 枚より良い）
_MAIN_MISS_SCORE = 10000      # 弾の表そのものが外れた版の点（別の表の外れ枚数よりいつも悪い）
_BROKEN_SCORE = 1000000       # 撃てなかった版（実行時エラー・コンパイルエラー・時間切れ・何も書かない）＝いつもいちばん悪い
_BROKEN_RE = re.compile(r'実行時エラーで止まりました|コンパイルエラーで撃てません|秒で終わらず|何も書かずに終わりました|'
                        r'読み込めませんでした|実行で止まりました')


def _main_score(mism):
    """弾の表の外れ → 版の点（純 Python）。撃てなかった版は外れの行が 1 行でも、いちばん悪い。
    2026-09-17 夜: 前年度比較で実行時エラーの版（外れの行 1 行）を「外れ 1 セル」＝いちばん良い版と数え、
    65 セル外れる版が出るたびに実行時エラーの版へ戻し、台帳にもそれが残った。"""
    if any(_BROKEN_RE.search(str(m)) for m in (mism or ())[:3]):
        return _BROKEN_SCORE
    return _MAIN_MISS_SCORE + len(mism or ())


def _keep_best(best, score, code, sub_name, feedback, labels=()):
    """いちばん外れの少ない版を覚える（純 Python）。→ (次に AI に渡すコード, 外れの説明, 戻したか)。
    外れが増えた版は捨てて、いちばん良かった版とその外れの説明を返す。
    2026-09-17 夜: 振り直しで NBSP を直させたら全角の扱いまで崩れ（合格 17→13 枚）、道具は崩れた版を「前回のマクロ」として
    次の往復に渡していた＝崩れた所から直し直す往復になった。"""
    if best.get('code') and score > best['score']:
        broke = [x for x in labels if x not in (best.get('labels') or ())]
        # 2026-09-17 夜: 集約で戻した後も同じ材料を渡し、AI が見出しの探し方を同じように崩して 4 往復くり返した
        #   ＝捨てた版がどの表を新しく外したかを言う（その表の形を崩さずに直させる）
        fb = (f"直しで外れが増えました（外れ {_score_text(best['score'])} → {_score_text(score)}）。この版は捨てます。"
              + (f"捨てた版は、前は合格していた表「{'」「'.join(broke[:4])}」を外しました（その表の形＝見出しの空白・改行・"
                 "2 段・並び・シート名 などの扱いを崩した）。" if broke else "")
              + f"下の「前回のマクロ」（外れ {_score_text(best['score'])} の版）から、次の外れだけを直してください"
              "（合っている表の扱いは変えない。書き直さず、要る行だけ足す）。\n\n" + str(best.get('feedback') or ''))
        return best['code'], fb, True
    best.update({'score': score, 'code': code, 'sub': sub_name, 'feedback': feedback, 'labels': list(labels)})
    return code, feedback, False


def _score_text(score):
    if score >= _BROKEN_SCORE:
        return "撃てない版"
    if score < 1:
        return "全部そろった版" + ("（表の語の書き込みあり）" if score > 0 else "")
    return f"{score} 枚" if score < _MAIN_MISS_SCORE else f"弾の表で {score - _MAIN_MISS_SCORE} セル"


def _tests_key(failed):
    """別の表の外れ → 比べる鍵（純 Python）。どの表が・どのセルで（番地が無ければ文の頭で）外れたか。"""
    return tuple(sorted((lab, _mismatch_key(m) or tuple(str(x)[:24] for x in m[:1])) for lab, m in failed))


# ----------------------------------------------------------------
# 回す
# ----------------------------------------------------------------

def _truth_of_test(t):
    """別の表の組の「直す前」の写しのパス → 正解の写しのパス（add_tests の名付けの逆・純 Python）。"""
    return re.sub(r'_before(\.\w+)$', r'_truth\1', str(t.get('before') or ''))


def _refresh_kinds(case):
    """型（kinds）の無い古い正解を、置き場の正解の写しから読み直す（COM）。→ 読み直したか。
    人の 2 冊から作った弾（truth あり）と別の表だけ。走行から拾った弾の正解はシートの姿なので読み直せない。"""
    items = []
    if case.get('truth') and os.path.isfile(str(case['truth'])) and 'kinds' not in (case.get('expect') or {'kinds': 1}):
        items.append(('main', case['truth'], case.get('sheet')))
    for k, t in enumerate(case.get('tests') or []):
        tp = _truth_of_test(t)
        if 'kinds' not in (t.get('expect') or {'kinds': 1}) and os.path.isfile(tp):
            items.append((k, tp, t.get('sheet')))
    if not items:
        return False
    snaps = _read_sheets([(p, s) for _k, p, s in items], fallback_first=True)
    if not snaps:
        return False
    for (k, _p, _s), snap in zip(items, snaps):
        if k == 'main':
            case['expect'] = snap
        else:
            case['tests'][k]['expect'] = snap
    return True


def _load_phrases(path):
    """言い換えのファイル（UTF-8・1 行 1 つ）→ [文]。無ければ None（形が違えば理由を出す）。"""
    try:
        with open(path, encoding='utf-8-sig') as f:
            return [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    except OSError as ex:
        print(f"エラー: 言い換えのファイルを読めません: {ex}")
        return None


def forge(name, target_file=None, ai=None, model=None, max_turns=3, register=False, dry_run=False,
          truth=None, before=None, tests=None, sheet=None, request=None, to=None, phrases=None,
          prompt_only=False, prompt_out=None):
    """鍛える 1 周: 拾う（無ければ）→ AI が書く → 写しで試す（不一致は AI へ・max_turns まで）→ 台帳へ → --register で登録。
    before＋truth＝人が用意した 2 冊から弾を作る（agent の走行は要らない）。truth だけ＝直前の走行の正解を人の表にする。
    tests＝別の表（直す前 正解 の 2 冊ずつ）でも試す。to＝登録先のブック名。
    会話している AI が書き手のとき（2026-09-19）: prompt_only＝次に書き手へ渡す問い（前回の外れつき）を prompt_out に書いて
    止まる（戻り値 None・AI は呼ばない）。ai='file'・model＝答えのファイル＝その答えを採点し、不合格なら次の問いを prompt_out に書く。"""
    d = _forge_load()
    case = d.get(str(name))
    live = {k: v for k, v in d.items() if not v.get('retired')}     # 引退した仕事はぶつかりに数えない（2026-09-18 shu「作りすぎは消す」）
    _OTHER_JOB_TEXTS.clear()
    # 合格済みでなくても登録簿に入っている仕事（依頼の語がある）は数える（2026-09-18: 一斉の撃ち直しで不合格に戻った 3 本の
    # 言い換えが、ほかの仕事の検査から消えていた）
    _OTHER_JOB_TEXTS.update({k: [str(v.get('request') or '')] + list(v.get('phrases') or [])
                             for k, v in live.items() if k != str(name) and (v.get('passed') or v.get('ask'))})
    _OTHER_JOB_ASKS.clear()
    _OTHER_JOB_ASKS.update({k: str(v.get('ask') or '') for k, v in live.items() if k != str(name) and (v.get('passed') or v.get('ask'))})
    import vbam_prefire as _vp
    _OTHER_JOB_WORDS.clear()
    _OTHER_JOB_WORDS.update({k: _vp.sheet_words((v.get('before_values') or {}).get('values'))
                             for k, v in live.items() if k != str(name) and v.get('passed') and v.get('before_values')})
    fresh = bool(before or truth)
    if fresh or case is None or not os.path.isfile(str(case.get('before') or '')):
        if before:
            if not truth:
                print("エラー: --before には --truth（人が直した正解のブック）が要ります")
                return False
            case = harvest_pair(name, before, truth, sheet, request or '')
        else:
            case = harvest(name, target_file, truth=truth)
        if case is None:
            return False
        d = _forge_load()
    if not dry_run and _refresh_kinds(case):
        print("正解の表の型（数・日付・文字）を読み直しました＝前のマクロも型まで比べて試し直します")
        if case.get('passed'):
            case['passed'] = False
        case['tests_passed'] = False
        d[str(name)] = case
        _forge_save(d)
    if case.get('passed') and case.get('ask'):
        # 合格していた版の依頼の語＝鍛え直しで決まりに反していない語を外させない（_check_fit の keep_ask）
        keep = [w for w in str(case.get('keep_ask') or '').split('|') + str(case['ask']).split('|') if w]
        case['keep_ask'] = '|'.join(dict.fromkeys(keep))
    if phrases:
        got = _load_phrases(phrases)
        if got is None:
            return False
        if got != case.get('phrases'):
            case['phrases'] = got
            ask = case.get('ask')
            if case.get('passed') and ask and any(not re.search(ask, p) for p in got):
                # 合格済みでも、依頼の語が言い換えに当たらなければ鍛え直す（値は合っている＝規則違反として返す）
                case['passed'] = False
                case['last_feedback'] = None
                print(f"言い換え {len(got)} 本のうち、依頼の語（{ask}）に当たらないものがあります＝鍛え直します")
            d[str(name)] = case
            _forge_save(d)
    if tests:
        n = add_tests(case, tests)
        if n is None:
            return False
        if n:
            print(f"別の表で試す組を {n} つ足しました（計 {len(case.get('tests') or [])}）")
            case['tests_passed'] = False               # 足した表でまだ撃っていない＝合格済みでも撃ち直す
            d[str(name)] = case
            _forge_save(d)
    feedback, code = None, None
    start_score, start_labels = None, []
    if case.get('code') and not dry_run:
        # ' 形: の行は道具が持つ＝前に鍛えた版にも足す／導き直す（AI は呼ばない・2026-09-18）
        shaped_code, shaped = ensure_shape_line(case, case['code'])
        if shaped:
            case['code'] = shaped_code
            d[str(name)] = case
            _forge_save(d)
            if case.get('passed') and case.get('bas') and os.path.isfile(str(case['bas'])):
                _write_bas(case['bas'], _MODULE, shaped_code)
            print(f"表の形の行を入れ直しました（要登録）: {_SHAPE_LINE_RE.search(shaped_code).group(0).strip()}")
    if case.get('passed') and case.get('code') and not dry_run:
        # 合格の後に規則を締めたときは、合格済みのマクロも今の規則で検査し直す（落ちたら値はそのまま・規則だけ直させる）
        nm, why, _h = _validate_code(case['code'])      # 通れば why は依頼の語（検査は _check_fit へ）
        if nm:
            _hd, why = _check_fit(case, case['code'], why)
        why = why or _literal_error(case, case['code'])
        if why:
            print(f"合格済みのマクロが今の規則に合いません: {why}＝鍛え直します")
            case['passed'] = False
            d[str(name)] = case
            _forge_save(d)
    if case.get('code') and case.get('bas') and os.path.isfile(str(case['bas'])) and not dry_run:
        # 試す物（.bas）と合格にして残す物（台帳のコード）をそろえる（2026-09-18 第二期: 2 本同時の鍛えで台帳に壊れた往復 1 の版、
        # .bas に合格した版が残り、.bas で試して通ったのに台帳の壊れた版を「合格」として .bas に書き戻した＝レーダー）
        _write_bas(case['bas'], _MODULE, case['code'])
    if case.get('passed') and case.get('bas') and os.path.isfile(case['bas']) and not dry_run:
        failed = []
        if case.get('tests') and not case.get('tests_passed'):
            res, sec = _try_tests(case, case['bas'], case['sub'])
            failed = [(lab, m) for lab, m in res if m]
            print(f"合格済みのマクロを別の表 {len(res)} 枚で試しました: 外れ {len(failed)} 枚（{sec:.1f} 秒・AI なし）")
        if not failed:
            case['tests_passed'] = bool(case.get('tests'))
            d[str(name)] = case
            _forge_save(d)
            print(f"「{name}」は合格済み（{case.get('sub')}・{case['bas']}）。撃ち直しません")
            return _register(case, target_file, to) if register else True
        case['passed'] = False
        feedback, code = _tests_feedback(case, failed), case.get('code')
        start_score, start_labels = len(failed), [lab for lab, _m in failed]
    elif (not fresh and not dry_run and not case.get('passed') and case.get('code') and case.get('sub')
          and case.get('bas') and os.path.isfile(case['bas'])):
        # 不合格のまま終わった弾の撃ち直し＝前回のマクロを今の表で撃ち直し、その外れの説明から続ける（最初から書かせない）
        mism, _s = _try_macro(case, case['bas'], case['sub'])
        if mism:
            ctx = _context_rows(case.get('expect'), mism)
            hint = _hints(mism, case.get('expect'))
            feedback = "\n".join(mism[:_MAX_MISMATCH]) + (("\n" + hint) if hint else "") + (("\n" + ctx) if ctx else "")
            start_score, start_labels = _main_score(mism), ['弾の表']
        else:
            res, _s = _try_tests(case, case['bas'], case['sub'])
            failed = [(lab, m) for lab, m in res if m]
            feedback = _tests_feedback(case, failed) if failed else None
            start_score, start_labels = len(failed), [lab for lab, _m in failed]
        if not feedback:
            nm, why, _hh = _validate_code(case['code'])
            ask_now = why if nm else None
            if nm:
                _h, why = _check_fit(case, case['code'], why)
            if why and ask_now and not dry_run:
                fixed_code, _fa = _auto_fix_ask(case, case['code'], ask_now)
                if fixed_code:
                    nm2, ask2, handles2 = _validate_code(fixed_code)
                    heads2, why2 = _check_fit(case, fixed_code, ask2) if nm2 else (None, ask2)
                    if nm2 and not why2:
                        case.update({'code': fixed_code, 'ask': ask2, 'handles': handles2, 'headers': heads2})
                        why = None
                        print(f"（値は全部そろっていて、頭の行は道具が直しました＝AI なし: {ask2}）")
            lit = _literal_error(case, case['code'])
            if lit:
                feedback = ("値は全部そろっていますが、" + lit + ((f"\n規則違反: {why}") if why else "")
                            + "\n値はそろえたまま、語で探している所・語を書いている所を表の形に直してください")
                start_score = _LITERAL_PASS_SCORE           # 外れのある直しより良く数えない（戻り先にしない）
            else:
                feedback = f"値は全部そろっています。規則違反だけ直してください: {why}" if why else None
        code = case['code']
        print("（前回の不合格のマクロを今の表で撃ち直し、その外れから続けます）" if feedback else
              "（前回のマクロは今の表で全部そろいました＝AI に聞かずに合格にします）")
        if not feedback:
            case.update({'passed': True, 'tests_passed': bool(case.get('tests')), 'mismatch': 0})
            d[str(name)] = case
            _forge_save(d)
            _write_bas(case['bas'], _MODULE, case['code'])
            return _register(case, target_file, to) if register else True
    prompt = _prompt(case, feedback, code)
    if dry_run:
        print(prompt)
        print(f"\n（--dry-run: ここまで。AI には聞いていません。弾 {case['before']}）")
        return True
    if prompt_only:
        _write_prompt(prompt_out, prompt)
        print(f"書き手への問いを書きました（{len(prompt)} 字・AI には聞いていません）: {prompt_out or '画面'}")
        return None
    best = {}                                              # いちばん外れの少なかった版（外れが増えた直しは捨てて戻す）
    if code and feedback and case.get('sub') and start_score is not None:
        best = {'score': start_score, 'code': code, 'sub': case.get('sub'), 'feedback': feedback, 'labels': start_labels}
    t_all = time.time()
    last_fb = feedback                                     # 最後に撃てたマクロの不一致（規則違反の往復でも AI に残す）
    prev_key = None                                        # 前の往復で外れたセルの組（同じなら止める＝AI を無駄に呼ばない）
    nothing_runs = 0                                       # 何も変えずに終わった同じ外れを見逃した回数
    last_reverted = None                                   # 前に捨てた版の外れの組（同じ版がまた出たら止める）
    seen_rules = {}                                        # 頭の行の規則違反ごとの回数（交互のくり返しを止める）
    for turn in range(1, int(max_turns) + 1):
        if feedback:
            prompt = _prompt(case, feedback, code)
        if ai == 'file':
            print(f"往復 {turn}/{max_turns}: 会話の AI が書いた答えを採点します（{model}）", flush=True)
        else:
            print(f"往復 {turn}/{max_turns}: AI にマクロを書かせます（{len(prompt)} 字）…", flush=True)
        t0 = time.time()
        text = _ask_once(ai, model, prompt)
        _log_reply(name, turn, prompt, text)
        new_code = _extract_code(text)
        new_code, fixed_ask = _normalize_ask_line(new_code)
        if fixed_ask:
            print(f"  依頼の語の正規表現の記号を語の区切りに直しました: {fixed_ask}")
        sub_name, ask_re, handles = _validate_code(new_code)
        rule_err = None if sub_name else f"規則違反: {ask_re}"
        heads, fit_err, lit_err = None, None, None
        if not rule_err:
            new_code, shaped = ensure_shape_line(case, new_code)     # ' 形: は道具が足す（AI の負担を増やさない）
            if shaped:
                print(f"  表の形の行を足しました: {_SHAPE_LINE_RE.search(new_code).group(0).strip()}")
        if not rule_err:
            # 頭の行（依頼の語・扱う・見出し）の違反は、マクロの中身とは別＝撃って値も一緒に返す
            # （2026-09-17 夜: 頭の行の違反で撃たずに返し、値の不一致を 1 度も見せないまま往復を使い切った）
            heads, why = _check_fit(case, new_code, ask_re)
            fit_err = f"規則違反: {why}" if why else None
            # 表の語の書き込みは本体の違反（道具は直せない）＝撃って値と一緒に返す（2026-09-18）
            lit_err = _literal_error(case, new_code)
        bas_path = os.path.join(_AGENT_FORGE_DIR, _safe_name(name) + '.bas')
        if not rule_err:
            err = _write_bas(bas_path, _SCRATCH_MODULE, new_code)
            if err:
                rule_err = err
            elif not _check_bas(bas_path):                 # 取り込む前の関所（文字コード・識別子）＝exam と同じ
                rule_err = "check-bas（文字コード・識別子の検査）に落ちました。VBA の名前の規則（先頭が _ でない・記号なし）を守ってください"
        if rule_err:
            # 前回の不一致と、いま規則に落ちたコード（無ければ前回のマクロ）を一緒に返す
            # （2026-09-17: 規則違反だけを返したら、AI が前回の不一致を忘れて 1 往復目の形に戻った）
            feedback = rule_err + (("\n\n（その前の不一致。直すのはこちらも）\n" + last_fb) if last_fb else "")
            code = new_code or code
            head = re.sub(r'\s+', ' ', str(text or ''))[:160]
            print(f"  返事が規則に合いません: {rule_err}（{time.time() - t0:.1f} 秒）　返事の頭: {head}")
            key = ('rule', rule_err)
            if key == prev_key:
                print("  同じ規則違反が 2 往復続きました。AI をこれ以上呼ばずに止めます（規則か弾の表を見直してください）")
                break
            prev_key = key
            continue
        code = new_code
        print(f"  Sub {sub_name}（依頼の語: {ask_re}／扱う: {' '.join(handles)}・AI {time.time() - t0:.1f} 秒）→ 写しで試します"
              + (f"\n  （頭の行が規則に合いません＝値と一緒に返します: {fit_err}）" if fit_err else "")
              + (f"\n  （表の語を書き込んでいます＝値と一緒に返します: {lit_err[:200]}）" if lit_err else ""))
        pre = [x for x in (lit_err, fit_err) if x]
        fit_fb = ("\n".join(pre) + "\n\n（値の外れ。直すのはこちらも）\n") if pre else ""
        mism, sec = _try_macro(case, bas_path, sub_name)
        case.update({'bas': bas_path, 'sub': sub_name, 'ask': ask_re, 'handles': handles, 'headers': heads, 'turns': turn,
                     'mismatch': len(mism), 'passed': not mism and not fit_err and not lit_err, 'code': code,
                     'tests_passed': False})
        if not mism and case.get('tests'):
            res, tsec = _try_tests(case, bas_path, sub_name)
            failed = [(lab, m) for lab, m in res if m]
            print(f"  弾の表は合格。別の表 {len(res)} 枚: 合格 {len(res) - len(failed)}・外れ {len(failed)}（{tsec:.1f} 秒・AI なし）")
            if failed:
                case.update({'passed': False, 'mismatch': sum(len(m) for _l, m in failed)})
                d[str(name)] = case
                _forge_save(d)
                for lab, m in failed:
                    print(f"    {lab}: " + " ／ ".join(m[:3]))
                last_fb = _tests_feedback(case, failed)
                feedback = fit_fb + last_fb
                case['last_feedback'] = feedback
                d[str(name)] = case
                _forge_save(d)
                key = ('tests',) + _tests_key(failed) + ((fit_err,) if fit_err else ()) + (('lit',) if lit_err else ())
                code, feedback, reverted = _keep_best(best, len(failed), code, sub_name, feedback,
                                                      [lab for lab, _m in failed])
                if reverted:
                    print(f"  直しで外れが増えました（外れ {best['score']} 枚 → {len(failed)} 枚）。この版は捨て、"
                          "いちばん良かった版から直させます")
                    if key == last_reverted:
                        print("  戻した後も同じ外れの版が 2 回続きました。AI をこれ以上呼ばずに止めます")
                        break
                    last_reverted = key
                    prev_key = key
                    continue
                if key == prev_key:
                    print(f"  同じ表が同じ所で 2 往復続けて外れました（{'・'.join(lab for lab, _m in failed)}）。AI をこれ以上呼ばずに止めます。\n"
                          "  その表の形（シート名・表題の行・列の並び）を依頼文に書くか、agent --forge を撃ち直してください"
                          "（前回のマクロと外れの説明から続けます）")
                    break
                prev_key = key
                continue
            case['tests_passed'] = True
        d[str(name)] = case
        _forge_save(d)
        if not mism and lit_err:
            # 値は全部そろったが表の語を書き込んでいる＝本体を直させる（頭の行だけの版より悪く、外れのある版より良い）
            code, _fb, reverted = _keep_best(best, _LITERAL_PASS_SCORE, code, sub_name, lit_err, [])
            if reverted:
                print("  値は全部そろいましたが表の語を書き込んだ版です。語の無い前の版に戻します")
                break
            print(f"  値は全部そろいました（写しで {sec:.1f} 秒）が、表の語を書き込んでいます＝本体を直させます")
            feedback = ("値は全部（別の表も）そろっていますが、" + lit_err
                        + (("\n" + fit_err) if fit_err else "")
                        + "\n値はそろえたまま、語で探している所・語を書いている所を表の形に直してください（書き直さず、要る所だけ）")
            last_fb = None
            key = ('lit', lit_err)
            if key == prev_key:
                print("  同じ語の書き込みが 2 往復続きました。AI をこれ以上呼ばずに止めます（依頼文に列の決まりを書き足してください）")
                break
            prev_key = key
            continue
        if not mism and fit_err:
            # 値が全部そろった版は、頭の行の規則違反があっても「いちばん良い版」（2026-09-17 深夜: 記録に入れず、止まったときに
            # ピボットを作らない前の版を台帳に戻していた）
            _keep_best(best, -1, code, sub_name, fit_err, [])
            fixed_code, fixed_ask = _auto_fix_ask(case, code, ask_re)
            if fixed_code:
                nm2, ask2, handles2 = _validate_code(fixed_code)
                heads2, why2 = _check_fit(case, fixed_code, ask2) if nm2 else (None, ask2)
                if nm2 and not why2 and not _write_bas(bas_path, _SCRATCH_MODULE, fixed_code):
                    code = fixed_code
                    case.update({'code': code, 'ask': ask2, 'handles': handles2, 'headers': heads2, 'passed': True,
                                 'tests_passed': bool(case.get('tests')), 'mismatch': 0})
                    d[str(name)] = case
                    _forge_save(d)
                    print(f"  値は全部そろいました。頭の行（依頼の語・扱う）は道具が直して合格にしました（AI なし）: {ask2}")
                    break
            print(f"  値は全部そろいました（写しで {sec:.1f} 秒）。頭の行の規則だけ直させます")
            feedback = ("値は全部（別の表も）そろっています。マクロの本体は変えずに、頭の行の規則違反だけ直してください: "
                        + fit_err)
            last_fb = None
            key = ('rule', fit_err)
            seen_rules[key] = seen_rules.get(key, 0) + 1
            if key == prev_key or seen_rules[key] >= 3:
                # 交互のくり返し（A→B→A→B）も止める（2026-09-18 未明: 見出しの 2 つの決まりを交互に破って 8 往復）
                print("  同じ規則違反がくり返されました。AI をこれ以上呼ばずに止めます（規則か弾の表を見直してください）")
                break
            prev_key = key
            continue
        if not mism:
            print(f"  合格: 値が全部そろいました（写しで {sec:.1f} 秒・AI なし）")
            break
        print(f"  不一致 {len(mism)} 件（先頭）:")
        for line in mism[:8]:
            print("    " + line)
        ctx = _context_rows(case.get('expect'), mism)
        hint = _hints(mism, case.get('expect'))
        last_fb = ("\n".join(mism[:_MAX_MISMATCH])
                   + (f"\n…ほか {len(mism) - _MAX_MISMATCH} 件" if len(mism) > _MAX_MISMATCH else "")
                   + (("\n" + hint) if hint else "") + (("\n" + ctx) if ctx else ""))
        feedback = fit_fb + last_fb
        case['last_feedback'] = feedback
        d[str(name)] = case
        _forge_save(d)
        # 番地の無い外れ（実行時エラー・グラフの数など）は文面で比べる（2026-09-17 深夜: 438 の同じ行で 5 往復くり返しても止まらなかった）
        key = _mismatch_key(mism) or tuple(str(m)[:120] for m in mism[:2])
        code, feedback, reverted = _keep_best(best, _main_score(mism), code, sub_name, feedback, ['弾の表'])
        if reverted:
            print("  直しで弾の表まで外れました。この版は捨て、いちばん良かった版から直させます")
            if key and key == last_reverted:
                # 戻した後も同じ外れの版が続く＝AI が同じ直しをくり返している（2026-09-17 夜: 4 往復同じ 65 セル）
                print("  戻した後も同じ外れの版が 2 回続きました。AI をこれ以上呼ばずに止めます")
                break
            last_reverted = key
            prev_key = key
            continue
        nothing = bool(mism) and str(mism[0]).startswith('マクロは表に何も書かずに終わりました')
        if key and key == prev_key and nothing and nothing_runs < 2:
            # 何も変えずに終わった版は「正解から読めない決まり」ではなくマクロの誤り＝もう 1 往復は直させる（2026-09-18）
            nothing_runs += 1
            prev_key = key
            continue
        if key and key == prev_key:
            print(f"  同じセルが 2 往復続けて外れました（{_key_text(key)}）。正解の表から読めない決まり（並び順・含める行・"
                  "書く場所など）がある形です。AI をこれ以上呼ばずに止めます。\n"
                  "  依頼文にその決まりを書き足して撃ち直してください（例: 「区分は対応表の順に」）")
            break
        prev_key = key
    if not case.get('passed') and best.get('code') and best['code'] != case.get('code') and case.get('bas'):
        # 不合格で終わったら、台帳と .bas はいちばん外れの少なかった版にしておく（次の撃ち直しはそこから続く）
        case.update({'code': best['code'], 'sub': best['sub'], 'last_feedback': best['feedback']})
        _write_bas(case['bas'], _SCRATCH_MODULE, best['code'])
        d[str(name)] = case
        _forge_save(d)
        print(f"  台帳のマクロは、いちばん外れの少なかった版（外れ {best['score'] if best['score'] < _MAIN_MISS_SCORE else '弾の表'}）に戻しました")
    print(f"鍛える: 「{name}」 {'合格' if case.get('passed') else '不合格'}・往復 {case.get('turns')}・{time.time() - t_all:.1f} 秒"
          + (f"　.bas: {case.get('bas')}" if case.get('bas') else ""))
    if case.get('passed'):
        # 台帳を書き換えて再度読み込む（登録の判定に使う）
        _write_bas(case['bas'], _MODULE, case['code'])       # 登録用にモジュール名を「表の整理」に
        if register:
            return _register(case, target_file, to)
        print(f"  登録するなら: agent --forge {name} --register（「表を整える」を持つ開いているブックの「{_MODULE}」に足して"
              " compile まで。--to ブック名 で名指し）")
    else:
        print("  不合格のまま。不一致を見て依頼の言い直しか、agent --forge を撃ち直してください（弾は残っています）")
    if prompt_out:
        # 会話の AI が書き手のとき: 次の問い（外れの説明と、直す元のマクロつき）を書いておく＝読んで答えを直せばよい
        if case.get('passed'):
            _write_prompt(prompt_out, "（合格しました。次の問いはありません）\n")
        else:
            _write_prompt(prompt_out, _prompt(case, feedback, code))
            print(f"  次の問い（外れの説明つき）を書きました: {prompt_out}")
    return bool(case.get('passed'))


def _pick_register_target(books, to=None):
    """登録先を決める（純 Python）。books＝[(ブック名, 「表を整える」を持つか)]（開いているブック。アドインは入れない）。
    → (ブック名, None) か (None, 理由)。
    2026-09-17: 登録先を「アクティブなブック」にしていたので、秀コンボがアクティブだと秀コンボへ、職場のブックが
    アクティブだと職場のブックへ入った（先撃ちが探すのは「表を整える」の持ち主）。名指し か 持ち主 1 冊 に限る。"""
    names = [b for b, _h in books]
    if to:
        want = str(to).strip()
        hit = [b for b in names if b == want] or [b for b in names if os.path.splitext(b)[0] == want]
        if not hit:
            return None, f"登録先「{want}」が開いていません（開いているブック: {'・'.join(names) or 'なし'}）"
        return hit[0], None
    owners = [b for b, has in books if has]
    if len(owners) == 1:
        return owners[0], None
    if not owners:
        return None, ("「表を整える」を持つブック（秀コンボ.xlsm など）が開いていません。開いてから撃つか、"
                      "--to ブック名 で登録先を名指ししてください（アドインには直接足しません＝更新登録で入れる）")
    return None, f"「表を整える」を持つブックが {len(owners)} 冊あります（{'・'.join(owners)}）。--to ブック名 で名指ししてください"


def _book_has_tidy(wb):
    with contextlib.suppress(Exception):
        cm = wb.VBProject.VBComponents(_MODULE).CodeModule
        n = int(cm.CountOfLines)
        return bool(n and re.search(r'^\s*Sub\s+表を整える\s*\(', cm.Lines(1, n), re.M))
    return False


def _register(case, target_file=None, to=None):
    """合格したマクロを「表を整える」の持ち主（か --to のブック）の標準モジュール「表の整理」の末尾に足す → compile。"""
    code = case.get('code')
    if not code:
        print("エラー: 登録するコードがありません（先に agent --forge 名前 で合格させてください）")
        return False
    if case.get('retired'):
        # 2026-09-18 夜: 引退した仕事（shu「作りすぎは消す」）を撃つと登録し直せてしまう＝印を消すまで登録しない
        print(f"引退した仕事です（{case['retired']}）。戻すなら台帳の retired の印を消してから --register")
        return False
    xl, _active = va.get_workbook(target_file)
    books = []
    with contextlib.suppress(Exception):
        books = [(str(w.Name), _book_has_tidy(w)) for w in xl.Workbooks]
    name, why = _pick_register_target(books, to)
    if not name:
        print("エラー: " + why)
        return False
    wb = xl.Workbooks(name)
    unsaved = not str(wb.Path or '')
    force = ['--force'] if unsaved else []          # 未保存のブックは控えを取れない（ファイルが無い）＝控え無しで足す
    has_module = False
    with contextlib.suppress(Exception):
        wb.VBProject.VBComponents(_MODULE)
        has_module = True
    if not has_module:
        ok, out = va._run_cmd(['add-module', _MODULE] + force, wb)
        if not ok:
            print(out.strip())
            return False
    # 同名の Sub が既にあれば replace、無ければ add（鍛え直しを重ねても二重にしない）
    exists = False
    with contextlib.suppress(Exception):
        cm = wb.VBProject.VBComponents(_MODULE).CodeModule
        n = int(cm.CountOfLines)
        exists = bool(n and re.search(r'^\s*Sub\s+' + re.escape(case['sub']) + r'\s*\(', cm.Lines(1, n), re.M))
    with open(LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(code)
    if exists:
        ok, out = va._run_cmd(['replace-procedure', '--module', _MODULE, '-y'] + force, wb)
    else:
        ok, out = va._run_cmd(['add-procedure', _MODULE, '-y'] + force, wb)
    if not ok:
        print(out.strip())
        return False
    ok_c, out_c = va._run_cmd(['compile'], wb)
    print(out_c.strip())
    if not ok_c:
        print("エラー: 登録はしましたが全体コンパイルが通りません。上の行を直してください（控えは backups にあります）")
        return False
    d = _forge_load()
    case['registered_to'] = str(wb.Name)
    case['registered_at'] = time.strftime('%Y-%m-%d %H:%M:%S')
    d[case['name']] = case
    _forge_save(d)
    print(f"登録しました: {wb.Name} の「{_MODULE}」に Sub {case['sub']}（依頼の語: {case.get('ask')}）")
    print("  次からは、依頼にその語があれば agent の入口で道具が撃ちます（AI は呼びません）。保存はしていません。"
          "アドインにも入れるなら: そのブックを保存 → アクティブにして close-form → run-macro アドインの更新登録 --raw")
    return True


def forged_list():
    d = _forge_load()
    if not d:
        print(f"鍛えたマクロはありません（{_AGENT_FORGE_FILE}）。")
        print("  agent で表を直したあと agent --forge 名前 と打つと、その残りの型をマクロにして写しで試します")
        return True
    print(f"鍛えたマクロ（{len(d)} 本・{_AGENT_FORGE_FILE}）:")
    for name, c in sorted(d.items(), key=lambda kv: kv[1].get('time', '')):
        state = '合格' if c.get('passed') else ('不合格' if c.get('bas') else '未')
        if c.get('tests'):
            state += f"（別の表 {len(c['tests'])} 枚{'も合格' if c.get('tests_passed') else 'は未合格'}）"
        if c.get('truth'):
            state += '・正解は人の表'
        reg = f"　登録: {c['registered_to']}" if c.get('registered_to') else ""
        print(f"  {name}　{c.get('time', '?')}　元: {c.get('from_book', '?')}!{c.get('sheet', '?')}　{state}"
              + (f"　Sub {c.get('sub')}（依頼の語: {c.get('ask')}）" if c.get('sub') else "") + reg)
        print(f"      依頼「{(c.get('request') or '')[:70]}」　AI の手 {len(c.get('hands') or [])} 個"
              + (f"　先撃ち: {' → '.join(c['prefire'])}" if c.get('prefire') else ""))
    print("（撃つ: agent --forge 名前 ／ 登録: agent --forge 名前 --register）")
    return True


__all__ = ['forge', 'forged_list', 'harvest', 'harvest_pair', 'add_tests', '_hands_of_log', '_diff_summary',
           '_extract_code', '_validate_code', '_compare', '_prompt', '_write_bas', '_forge_load', '_forge_save',
           '_pick_register_target', '_tests_feedback', '_AGENT_FORGE_FILE', '_AGENT_FORGE_DIR', 'FORGE_RULES']
