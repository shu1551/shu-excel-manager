# -*- coding: utf-8 -*-
"""vbam_rehearse.py — vba_manager 分割パート: 先撃ちの試し撃ち（登録簿のマクロを写しで先に撃って確かめる）

2026-10-08: 鍛える回路（vbam_forge）を外したとき、依頼の本体の先撃ち（vbam_prefire._rehearse・形の判定）が
使っていた部品だけをここへ移した（中身は vbam_forge のときと同じ）。
  rehearse_steps … 写しのブックで、登録簿のマクロを本番と同じ順に撃ち、1 本ずつ「止まらずに終わったか・表が変わったか」を返す
  _kind_of / _texts / _snapshot_texts … 表の値と型を読む
"""
import os
import re
import json
import time
import shutil
import contextlib

import vbam_core
from vbam_core import SCRIPT_DIR, LAST_PROC_FILE, _col_letter, get_or_start_excel
from vbam_hands import _rows_of, _text_of

def _texts(values):
    return [[_text_of(v) for v in row] for row in (values or [])]

def _addr_of(r0, c0, i, j):
    return f"{_col_letter(c0 + j)}{r0 + i}"

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
    # 棚が「表の整理_〜」に分かれていると module_text は [(モジュール名, 本文)]。つないで 1 本にすると
    # Option Explicit・Option Private Module が途中に来てコンパイルが通らないので、モジュールごとに入れる（2026-09-23）
    parts = [(module, module_text)] if isinstance(module_text, str) else list(module_text)
    bases = []
    for mod, text in parts:
        bas = os.path.join(work, mod + '.bas')
        err = _write_bas(bas, mod, text)
        if err:
            return [{'name': names[0] if names else '', 'ok': False, 'changed': False, 'why': err}]
        bases.append(bas)
    xl, pid = _start_own()
    mwb = wb = None
    macro_file = None
    try:
        try:
            mwb = xl.Workbooks.Add()
            for bas in bases:
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
