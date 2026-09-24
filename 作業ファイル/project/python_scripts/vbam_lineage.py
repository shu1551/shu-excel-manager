# -*- coding: utf-8 -*-
"""vbam_lineage.py — vba_manager 分割パート: 系譜と、閉じたブックを読む手（2026-09-17）

  versions <ブック名> [--dir 追加フォルダ]      同名ブックの写しを日時順に（大きさ・モジュール数・プロシージャ数・最新との差・開いている印）
  history <マクロ名> [--book 名] [--deep]      そのマクロの本文の歴史（backups の控え・_exports・--deep で写しの .xlsm）。同じ本文は畳む
  list-file <path.xlsm>                        閉じたブックのマクロ一覧
  grep-files "文字" <フォルダ|ファイル…>        閉じたブック横断のコード検索（[ファイル名][モジュール] プロシージャ:行: 本文）
  export-file <path.xlsm> [--dir 先]           閉じたブックのモジュールを .bas/.cls/.frm に書き出す（.frx は無い）

中身は oletools（olevba）で読む＝Excel を開かない。Workbook_Open もアドインも起きない・50 冊でも 1 手・COM 不要。
9/16: 同名ブックが 4 冊（Desktop・保存・プロジェクト・エクセル）あり、9/4 の直しが「ポスター - コピー」に入っていて
9/16 に同じ欠陥をもう一度直した＝「取り違え 0」のための手。
行番号は VBE の CodeModule と同じ（書き出しの頭の VERSION／Begin…End／Attribute 行を外して数える）。
"""
import os
import re
import sys
import json
import time
import logging

from vbam_core import SCRIPT_DIR, BACKUP_DIR, _reject_extra_args, smart_path_resolve

# oletools は INFO／WARNING（Opening ZIP… / VBA stomping cannot be detected…）を logger に流す。MCP サーバー（FastMCP）が
# 根の logger に rich の handler を付けているので、そのままだと出力に何十行も混ざる（2026-09-17 versions で 187,000 字）。
for _name in ('olevba', 'oletools', 'oletools.olevba', 'olefile', 'oletools.common', 'oletools.thirdparty'):
    logging.getLogger(_name).setLevel(logging.ERROR)
from vbam_vba import _split_procedures, _proc_of_index
from vbam_devtools import _module_diff, _read_bas_text, _proc_body_from_bas

_BOOK_EXTS = ('.xlsm', '.xlam', '.xls', '.xlsb', '.xla', '.xlt', '.xltm')
_MODULE_EXTS = ('.bas', '.cls', '.frm')
_ATTR_LINE = re.compile(r'^Attribute\s+\S+\s*=')
_HEADER_VERSION = re.compile(r'^\s*VERSION\b', re.IGNORECASE)
_HEADER_BEGIN = re.compile(r'^\s*Begin\b', re.IGNORECASE)
_HEADER_END = re.compile(r'^\s*End\s*$', re.IGNORECASE)
_STAMP_RE = re.compile(r'(\d{8})_(\d{6})')
_TYPE_LABEL = {'.bas': '標準', '.cls': 'クラス', '.frm': 'フォーム'}


# ================================================================
# 読み手（純 Python＋oletools）
# ================================================================

def _strip_export_header(text):
    """書き出し（.bas/.cls/.frm）の頭の VERSION／Begin…End／Attribute 行と、中の Attribute 行（ショートカット等）を外し、
    VBE の CodeModule と同じ行並びにする（純 Python）。戻り値: 行のリスト。"""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    n = len(lines)
    i = 0
    while i < n:
        ln = lines[i]
        if _HEADER_VERSION.match(ln) or _ATTR_LINE.match(ln):
            i += 1
            continue
        if _HEADER_BEGIN.match(ln):
            depth = 0
            while i < n:
                if _HEADER_BEGIN.match(lines[i]):
                    depth += 1
                elif _HEADER_END.match(lines[i]):
                    depth -= 1
                    if depth <= 0:
                        i += 1
                        break
                i += 1
            continue
        break
    body = [ln for ln in lines[i:] if not _ATTR_LINE.match(ln)]
    while body and not body[-1].strip():
        body.pop()
    return body


def _is_book_file(name):
    low = name.lower()
    return low.endswith(_BOOK_EXTS) and not name.startswith('~$')


_OLE_MODULE_CACHE = {}  # (normpath, mtime, size) -> list of (name, ext, code)


def _read_closed_modules(path):
    """閉じたブックのモジュール → [(モジュール名, 拡張子, 書き出しそのままの本文)]。マクロが無ければ []。

    oletools の olevba で vbaProject.bin を読む（.xlsm/.xlam/.xlsb/.xls）。文字コードは PROJECT の codepage を olevba が
    当てるので日本語のモジュール名・本文もそのまま（2026-09-17 にポスター.xlsm 23 本・0.03 秒で確認）。
    mtime と size をキーにキャッシュし、同一ファイルの再パースをゼロ秒にする。
    """
    try:
        norm = os.path.normcase(os.path.abspath(path))
        st = os.stat(path)
        ckey = (norm, st.st_mtime, st.st_size)
        if ckey in _OLE_MODULE_CACHE:
            return list(_OLE_MODULE_CACHE[ckey])
    except Exception:
        ckey = None

    from oletools.olevba import VBA_Parser
    vp = VBA_Parser(path)
    try:
        if not vp.detect_vba_macros():
            if ckey:
                _OLE_MODULE_CACHE[ckey] = []
            return []
        out = []
        for (_fname, _stream, vba_filename, vba_code) in vp.extract_all_macros():
            name = os.path.splitext(os.path.basename(str(vba_filename)))[0]
            ext = os.path.splitext(str(vba_filename))[1].lower() or '.bas'
            out.append((name, ext, vba_code or ''))
        if ckey:
            _OLE_MODULE_CACHE[ckey] = out
        return list(out)
    finally:
        vp.close()


def _read_closed_book(path, raw=False):
    """{モジュール名: 本文}。本文は VBE と同じ行並び（raw=True なら書き出しそのまま）。"""
    return {name: (code if raw else '\n'.join(_strip_export_header(code)))
            for name, _ext, code in _read_closed_modules(path)}


def _module_type_label(ext, raw):
    if ext == '.cls' and re.search(r'VB_Base\s*=\s*"0\{000208(?:19|20|21)', raw or '', re.IGNORECASE):
        return 'シート/ブック'
    return _TYPE_LABEL.get(ext, ext)


def _stamp_of(path):
    """ファイル名（か親フォルダ名）の YYYYMMDD_HHMMSS → 時刻。無ければ mtime。"""
    for s in (os.path.basename(path), os.path.basename(os.path.dirname(path))):
        m = _STAMP_RE.search(s)
        if m:
            try:
                return time.mktime(time.strptime(m.group(1) + m.group(2), '%Y%m%d%H%M%S'))
            except ValueError:
                pass
    return os.path.getmtime(path)


def _fmt_time(t):
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t))


def _open_book_paths():
    """いま Excel で開いているブックのフルパス（小文字）。Excel が無ければ空＝Excel を起こさない（GetActiveObject だけ）。"""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        xl = win32com.client.GetActiveObject("Excel.Application")
        return {os.path.normcase(os.path.abspath(str(wb.FullName))) for wb in xl.Workbooks}
    except Exception:
        return set()


# ================================================================
# versions（同名ブックの写しを日時順に）
# ================================================================

def _version_dirs(extra=()):
    """探す場所: Desktop／このプロジェクト／Desktop\\保存／Desktop\\エクセル／公開リポ／backups（サブフォルダ 1 段）／cwd／--dir。
    環境変数 VBAM_VERSION_DIRS（; 区切り）でも足せる。"""
    desk = os.path.join(os.path.expanduser('~'), 'Desktop')
    dirs = [desk, os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '..')),
            os.path.join(desk, '保存'), os.path.join(desk, 'エクセル'), 
            BACKUP_DIR, os.getcwd()]
    dirs += [d for d in os.environ.get('VBAM_VERSION_DIRS', '').split(';') if d.strip()]
    dirs += list(extra or ())
    out, seen = [], set()
    for d in dirs:
        key = os.path.normcase(os.path.abspath(d))
        if key in seen or not os.path.isdir(d):
            continue
        seen.add(key)
        out.append(d)
    return out


def _find_versions(stem, extra_dirs=()):
    """名前に stem を含むブックのパス一覧（重複なし）。backups だけはサブフォルダ 1 段も見る。"""
    want = stem.lower()
    found, seen = [], set()

    def _scan(d):
        try:
            names = os.listdir(d)
        except OSError:
            return
        for nm in names:
            p = os.path.join(d, nm)
            if os.path.isfile(p) and _is_book_file(nm) and want in nm.lower():
                key = os.path.normcase(os.path.abspath(p))
                if key not in seen:
                    seen.add(key)
                    found.append(p)

    for d in _version_dirs(extra_dirs):
        _scan(d)
        if os.path.normcase(os.path.abspath(d)) == os.path.normcase(os.path.abspath(BACKUP_DIR)):
            for nm in os.listdir(d):
                sub = os.path.join(d, nm)
                if os.path.isdir(sub):
                    _scan(sub)
    return found


def _versions_rows(paths):
    """各写しを読んで [{path, mtime, size, modules, procs, code{名: 本文}, error}]（新しい順）。"""
    rows = []
    for p in paths:
        row = {'path': p, 'mtime': os.path.getmtime(p), 'size': os.path.getsize(p), 'modules': 0, 'procs': 0,
               'code': {}, 'error': None}
        try:
            for name, _ext, raw in _read_closed_modules(p):
                lines = _strip_export_header(raw)
                row['code'][name] = '\n'.join(lines)
                row['modules'] += 1
                row['procs'] += len(_split_procedures(lines))
        except Exception as ex:
            row['error'] = str(ex)
        rows.append(row)
    rows.sort(key=lambda r: r['mtime'], reverse=True)
    return rows


def _diff_against_latest(rows):
    """最新（rows[0]）との差を各行に足す: changed（変わったモジュール名）・added・removed。純 Python。"""
    if not rows:
        return rows
    latest = rows[0]['code']
    for r in rows:
        if r is rows[0] or r['error']:
            r['changed'], r['added'], r['removed'] = [], [], []
            continue
        mine = r['code']
        r['changed'] = sorted(n for n in mine if n in latest and mine[n] != latest[n])
        r['added'] = sorted(n for n in mine if n not in latest)          # この写しにだけある
        r['removed'] = sorted(n for n in latest if n not in mine)        # 最新にはあるがこの写しに無い
    return rows


def cmd_versions(args):
    """同名ブックの写しを日時順に: versions <ブック名> [--dir 追加フォルダ] [--json]（COM 不要・Excel を開かない）

    名前に語幹を含む .xlsm/.xlam/.xls/.xlsb を Desktop／プロジェクト／保存／エクセル／公開リポ／backups から集め、
    1 冊 1 行＝日時・大きさ・モジュール数・プロシージャ数・最新との差（変わったモジュール名）・「開いている」印。
    """
    rest = list(args.posargs or [])
    if not rest:
        print("使い方: versions <ブック名（拡張子なし可）> [--dir 追加フォルダ] [--json]")
        return False
    if _reject_extra_args(rest, 1, 'ブック名は 1 つ'):
        return False
    given = rest[0]
    stem = os.path.splitext(os.path.basename(given))[0]
    extra = list(getattr(args, 'dir_opt', None) or [])
    if os.path.isfile(given):
        extra.append(os.path.dirname(os.path.abspath(given)))
    paths = _find_versions(stem, extra)
    # backups の控え（replace-procedure 等が撃つたびに残る <ブック>.xlsm.backup_before_…）は百冊単位にふくらむ
    # （ポスターで 127 冊・2026-09-17）。既定は新しい方から --max 冊（既定 3）だけ読み、残りは数だけ言う。--all で全部
    bak_root = os.path.normcase(os.path.abspath(BACKUP_DIR))
    is_bak = lambda p: os.path.normcase(os.path.abspath(p)).startswith(bak_root + os.sep)
    main = [p for p in paths if not is_bak(p)]
    baks = sorted((p for p in paths if is_bak(p)), key=os.path.getmtime, reverse=True)
    _m = getattr(args, 'max_hits', None)
    keep = len(baks) if getattr(args, 'all', False) else (3 if _m is None else max(0, int(_m)))
    hidden = baks[keep:]
    rows = _diff_against_latest(_versions_rows(main + baks[:keep]))
    opened = _open_book_paths()
    for r in rows:
        r['open'] = os.path.normcase(os.path.abspath(r['path'])) in opened
        r['backup'] = is_bak(r['path'])
    if getattr(args, 'json', False):
        slim = [{k: v for k, v in r.items() if k != 'code'} for r in rows]
        print(json.dumps({"success": True, "stem": stem, "count": len(rows), "hidden_backups": len(hidden),
                          "versions": slim}, ensure_ascii=False), file=sys.stdout)
        return True
    if not rows:
        print(f"「{stem}」を名前に含むブックは見つかりません（探した場所: {len(_version_dirs(extra))} フォルダ。--dir で足せます）")
        return True
    print(f"versions: 「{stem}」 {len(rows)} 冊（新しい順・最新との差は変わったモジュール名"
          + (f"・backups の控えは新しい {keep} 冊だけ" if hidden else "") + "）")
    for i, r in enumerate(rows):
        when = time.strftime('%m/%d %H:%M', time.localtime(r['mtime']))
        mark = "  [開いている]" if r['open'] else ""
        head = f"  {when} {r['size'] / 1024:>7,.0f} KB"
        if r['error']:
            print(f"{head}  読めません: {r['path']}  ({r['error']}){mark}")
            continue
        head += f" {r['modules']:>3} モジュール {r['procs']:>4} 本  {r['path']}{mark}"
        if i == 0:
            print(head + "  ← 最新")
            continue
        if r['modules'] == 0:
            print(head)
            print("      マクロ無し（VBA プロジェクトが入っていない＝差は比べない）")
            continue
        notes = []
        if r['changed']:
            notes.append("変わった: " + "・".join(r['changed'][:8]) + (f" 他{len(r['changed']) - 8}" if len(r['changed']) > 8 else ""))
        if r['removed']:
            notes.append("無い: " + "・".join(r['removed'][:5]) + (f" 他{len(r['removed']) - 5}" if len(r['removed']) > 5 else ""))
        if r['added']:
            notes.append("この写しだけ: " + "・".join(r['added'][:5]) + (f" 他{len(r['added']) - 5}" if len(r['added']) > 5 else ""))
        print(head)
        print("      " + ("最新との差: " + " ／ ".join(notes) if notes else "最新とコードは同じ"))
    if hidden:
        oldest = time.strftime('%m/%d %H:%M', time.localtime(os.path.getmtime(hidden[-1])))
        print(f"  … 他 {len(hidden)} 冊の控え（backups・{oldest} まで。読んでいない。--all で全部、--max N で冊数）")
    print("  中身を見る: list-file <パス> ／ 差分: export-file <パス> → diff-module <モジュール> <書き出した .bas>")
    return True


# ================================================================
# history（マクロ本文の歴史）
# ================================================================

def _history_files(book=None):
    """backups と _exports から (時刻, 表示名, パス) を集める（本文はまだ読まない）。"""
    items = []
    if os.path.isdir(BACKUP_DIR):
        for nm in os.listdir(BACKUP_DIR):
            if not nm.lower().endswith(_MODULE_EXTS):
                continue
            if book and not nm.lower().startswith(book.lower() + '_'):
                continue
            p = os.path.join(BACKUP_DIR, nm)
            items.append((_stamp_of(p), 'backups/' + nm, p))
    exp = os.path.join(SCRIPT_DIR, '_exports')
    if os.path.isdir(exp):
        for bk in os.listdir(exp):
            if book and bk.lower() != book.lower():
                continue
            bdir = os.path.join(exp, bk)
            if not os.path.isdir(bdir):
                continue
            for st in sorted(os.listdir(bdir)):
                sdir = os.path.join(bdir, st)
                if not os.path.isdir(sdir):
                    continue
                for nm in os.listdir(sdir):
                    if nm.lower().endswith(_MODULE_EXTS):
                        p = os.path.join(sdir, nm)
                        items.append((_stamp_of(p), f"_exports/{bk}/{st}/{nm}", p))
    return items


def _history_versions(macro, book=None, deep=False, extra_dirs=()):
    """マクロ本文の版 → [{time, label, path, body, n_lines, dup}]（古い順・同じ本文は畳む）。純 Python（--deep は oletools）。"""
    found = []                                     # (時刻, 表示名, パス, 本文)
    needle_bytes = macro.encode('cp932', errors='replace')
    for t, label, p in _history_files(book):
        try:
            with open(p, 'rb') as f:
                blob = f.read()
        except OSError:
            continue
        if needle_bytes not in blob and macro.encode('utf-8', errors='replace') not in blob:
            continue                               # 名前が字面で無いファイルは読まない（backups は千本ある）
        body = _proc_body_from_bas(blob.decode('cp932', errors='replace'), macro)
        if body is not None:
            found.append((t, label, p, body))
    if deep and book:
        for p in _find_versions(book, extra_dirs):
            try:
                mods = _read_closed_modules(p)
            except Exception:
                continue
            for name, _ext, raw in mods:
                body = _proc_body_from_bas(raw, macro)
                if body is not None:
                    found.append((os.path.getmtime(p), f"{os.path.basename(p)}!{name}", p, body))
    found.sort(key=lambda e: e[0])
    versions = []
    for t, label, p, body in found:
        if versions and versions[-1]['body'] == body:
            versions[-1]['dup'] += 1
            continue
        versions.append({'time': t, 'label': label, 'path': p, 'body': body,
                         'n_lines': body.count('\n') + 1 if body else 0, 'dup': 1})
    return versions, len(found)


def cmd_history(args):
    """マクロ本文の歴史: history <マクロ名> [--book ブック名] [--deep] [--dir 追加] [--max N] [--json]（COM 不要）

    backups の *.bas/*.frm/*.cls と _exports/**（--deep で versions と同じ場所の写しの .xlsm）からその名前の Sub/Function
    を集め、同じ本文は畳み、日時順に「いつ・どのファイル・何行」と隣との差分を出す。日時はファイル名の日時→無ければ mtime。
    """
    rest = list(args.posargs or [])
    if not rest:
        print("使い方: history <マクロ名> [--book ブック名] [--deep] [--max N] [--json]")
        return False
    macro = rest[0]
    if _reject_extra_args(rest, 1, 'マクロ名は 1 つ'):
        return False
    book = getattr(args, 'book_opt', None)
    if book:
        book = os.path.splitext(os.path.basename(book))[0]
    deep = getattr(args, 'deep', False)
    if deep and not book:
        print("エラー: --deep には --book ブック名 が要ります（写しの .xlsm をどの名前で探すか）")
        return False
    _m = getattr(args, 'max_hits', None)
    max_lines = 40 if _m is None else int(_m)
    versions, n_files = _history_versions(macro, book, deep, getattr(args, 'dir_opt', None) or ())
    if getattr(args, 'json', False):
        out = [{k: v for k, v in v_.items() if k != 'body'} | {'time_str': _fmt_time(v_['time'])} for v_ in versions]
        print(json.dumps({"success": True, "macro": macro, "files": n_files, "versions": out}, ensure_ascii=False),
              file=sys.stdout)
        return True
    if not versions:
        print(f"'{macro}' を含む控えがありません（backups・_exports"
              + ("・写しの .xlsm" if deep else "") + "）。"
              + ("" if deep else f" 写しの .xlsm も探すなら: history {macro} --book <ブック名> --deep"))
        return True
    print(f"history: {macro}  {n_files} 件のファイル → {len(versions)} 版（同じ本文は畳んだ・古い順）")
    for i, v in enumerate(versions, 1):
        dup = f"（同じ本文 他 {v['dup'] - 1} 件）" if v['dup'] > 1 else ""
        print(f"  版{i}  {_fmt_time(v['time'])}  {v['label']}  {v['n_lines']} 行{dup}")
        if i > 1:
            lines, minus, plus = _module_diff(versions[i - 2]['body'], v['body'], f"版{i - 1}", f"版{i}", max_lines=max_lines)
            print(f"    --- 版{i - 1} → 版{i}  -{minus} +{plus} 行 ---")
            for ln in lines:
                print("    " + ln)
    if not deep:
        print(f"  （写しの .xlsm も並べるなら: history {macro} --book <ブック名> --deep）")
    return True


# ================================================================
# 閉じたブックを読む（list-file・grep-files・export-file）
# ================================================================

def _resolve_book(arg):
    p = smart_path_resolve(arg) or arg
    if not os.path.isfile(p):
        print(f"エラー: ファイルが見つかりません: {arg}")
        return None
    return os.path.abspath(p)


def cmd_list_file(args):
    """閉じたブックのマクロ一覧: list-file <path.xlsm> [--json]（Excel を開かない）"""
    rest = list(args.posargs or [])
    if not rest:
        print("使い方: list-file <path.xlsm|xlam|xlsb|xls> [--json]")
        return False
    if _reject_extra_args(rest, 1, 'ファイルは 1 つ'):
        return False
    path = _resolve_book(rest[0])
    if path is None:
        return False
    try:
        mods = _read_closed_modules(path)
    except Exception as ex:
        print(f"エラー: 読めませんでした: {ex}")
        return False
    rows = []
    for name, ext, raw in mods:
        lines = _strip_export_header(raw)
        procs = _split_procedures(lines)
        rows.append({'module': name, 'type': _module_type_label(ext, raw), 'ext': ext, 'lines': len(lines),
                     'procs': [p['name'] for p in procs]})
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "file": os.path.basename(path), "path": path, "has_macros": bool(mods),
                          "modules": rows}, ensure_ascii=False), file=sys.stdout)
        return True
    if not mods:
        print(f"{os.path.basename(path)}: マクロ無し（VBA プロジェクトが入っていない）")
        return True
    n_procs = sum(len(r['procs']) for r in rows)
    print(f"閉じたブック: {os.path.basename(path)}  {len(rows)} モジュール・{n_procs} 本（Excel は開いていない）")
    for r in rows:
        names = ", ".join(r['procs'][:8]) + (f" 他{len(r['procs']) - 8}" if len(r['procs']) > 8 else "")
        print(f"  [{r['type']}] {r['module']:<20} {r['lines']:>5} 行  {len(r['procs']):>3} 本  {names}")
    print("  中身: grep-files \"文字\" <このファイル> ／ 書き出し: export-file <このファイル>")
    return True


def _expand_targets(targets):
    """フォルダはその直下のブック（.xlsm 等）に、ファイルはそのまま。存在しないものは (None, 名前) で返す。"""
    files, missing = [], []
    for t in targets:
        p = smart_path_resolve(t) or t
        if os.path.isdir(p):
            for nm in sorted(os.listdir(p)):
                fp = os.path.join(p, nm)
                if os.path.isfile(fp) and _is_book_file(nm):
                    files.append(os.path.abspath(fp))
        elif os.path.isfile(p):
            files.append(os.path.abspath(p))
        else:
            missing.append(t)
    seen, out = set(), []
    for f in files:
        k = os.path.normcase(f)
        if k not in seen:
            seen.add(k)
            out.append(f)
    return out, missing


def cmd_grep_files(args):
    """閉じたブック横断の検索: grep-files "文字" <フォルダ|ファイル…> [--regex] [-i] [--max N] [--json]

    出力は grep と同じ形 `[ファイル名][モジュール] プロシージャ:行: 本文`（行番号は VBE の CodeModule と同じ）。
    """
    rest = list(args.posargs or [])
    if not rest:
        print("使い方: grep-files \"文字\" <フォルダ|ファイル…> [--regex] [-i] [--max N] [--json]")
        return False
    if len(rest) >= 2 and os.path.exists(rest[0]) and not any(os.path.exists(t) for t in rest[1:]):
        # grep-files <ファイル> "文字" の順でも受ける（前は「文字」をファイルとして探し 0 冊を読んで黙った・2026-09-24）
        rest = rest[1:] + rest[:1]
    needle = rest[0]
    targets = rest[1:] or ['.']
    flags = re.IGNORECASE if getattr(args, 'ignore_case', False) else 0
    try:
        pat = re.compile(needle if getattr(args, 'regex', False) else re.escape(needle), flags)
    except re.error as ex:
        print(f"エラー: 正規表現が不正です: {ex}")
        return False
    _m = getattr(args, 'max_hits', None)
    max_hits = 200 if _m is None else int(_m)
    files, missing = _expand_targets(targets)
    for t in missing:
        print(f"⚠ 見つかりません: {t}")
    hits, total, unread, hit_files = [], 0, [], set()
    for fp in files:
        try:
            mods = _read_closed_modules(fp)
        except Exception as ex:
            unread.append((fp, str(ex)))
            continue
        base = os.path.basename(fp)
        for name, _ext, raw in mods:
            lines = _strip_export_header(raw)
            procs = None
            for i, ln in enumerate(lines, 1):
                if not pat.search(ln):
                    continue
                total += 1
                hit_files.add(fp)
                if len(hits) < max_hits:
                    if procs is None:
                        procs = _split_procedures(lines)
                    hits.append({'file': base, 'path': fp, 'module': name, 'proc': _proc_of_index(procs, i - 1) or '',
                                 'line': i, 'text': ln.rstrip()})
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "pattern": needle, "files": len(files), "files_hit": len(hit_files),
                          "total": total, "hits": hits, "unread": [{"path": p, "error": e} for p, e in unread]},
                         ensure_ascii=False), file=sys.stdout)
        return True
    for h in hits:
        proc_part = f" {h['proc']}" if h['proc'] else ""
        print(f"[{h['file']}][{h['module']}]{proc_part}:{h['line']}: {h['text'].strip()}")
    if total > len(hits):
        print(f"…他 {total - len(hits)} 件（--max で上限変更可）")
    for p, e in unread:
        print(f"⚠ 読めませんでした: {p} ({e})")
    print(f"--- {total} 件（{len(files)} 冊を読んだ・ヒット {len(hit_files)} 冊・Excel は開いていない）---")
    return True


def cmd_export_file(args):
    """閉じたブックのモジュールを書き出す: export-file <path.xlsm> [モジュール名…] [--dir 先] [--json]

    既定の置き場は _exports/<ブック名>/<日時>/（export-all --history と同じ並び）。.bas/.cls/.frm はコードだけ＝
    フォームの見た目（.frx）は入っていない。文字コードは CP932（Excel の Export と同じ）。
    モジュール名を続けるとそれだけを書き出す（前は「余分な引数」で断っていた＝台帳で 2 回・2026-09-24）。
    """
    rest = list(args.posargs or [])
    if not rest:
        print("使い方: export-file <path.xlsm> [モジュール名…] [--dir 出力先] [--json]")
        return False
    path = _resolve_book(rest[0])
    if path is None:
        return False
    try:
        mods = _read_closed_modules(path)
    except Exception as ex:
        print(f"エラー: 読めませんでした: {ex}")
        return False
    if not mods:
        print(f"{os.path.basename(path)}: マクロ無し（書き出す物がありません）")
        return True
    want = [str(w) for w in rest[1:]]
    # 2 つ目にパス（C:/tmp/xlam_check）＝書き出し先のつもり（台帳で 2 回「余分な引数」だった）。モジュール名に / \ : は入らない
    dirs = [w for w in want if re.search(r'[\\/:]', w)]
    if dirs and not getattr(args, 'dir_opt', None) and len(dirs) == 1:
        args.dir_opt = dirs[0]
        want = [w for w in want if w not in dirs]
    if want:
        have = {name.lower(): name for name, _e, _r in mods}
        unknown = [w for w in want if w.lower() not in have]
        if unknown:
            print(f"エラー: モジュールが見つかりません: {', '.join(unknown)}")
            print("  あるモジュール: " + ', '.join(name for name, _e, _r in mods))
            return False
        keep = {w.lower() for w in want}
        mods = [m for m in mods if m[0].lower() in keep]
    stem = os.path.splitext(os.path.basename(path))[0]
    out_dir = getattr(args, 'dir_opt', None) or os.path.join(SCRIPT_DIR, '_exports', stem, time.strftime('%Y%m%d_%H%M%S'))
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    written = []
    for name, ext, raw in mods:
        text = raw.replace('\r\n', '\n').replace('\r', '\n')
        if not text.endswith('\n'):
            text += '\n'
        fp = os.path.join(out_dir, name + ext)
        with open(fp, 'wb') as f:
            f.write(text.replace('\n', '\r\n').encode('cp932', errors='replace'))
        written.append(fp)
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "file": os.path.basename(path), "dir": out_dir, "count": len(written),
                          "files": [os.path.basename(w) for w in written]}, ensure_ascii=False), file=sys.stdout)
        return True
    print(f"書き出し: {len(written)} 本 → {out_dir}（.frx は無い＝フォームはコードだけ）")
    print("  比べる: diff-module <モジュール> \"" + os.path.join(out_dir, '<モジュール>.bas') + "\"")
    return True


__all__ = ['_strip_export_header', '_read_closed_modules', '_read_closed_book', '_module_type_label', '_stamp_of',
           '_version_dirs', '_find_versions', '_versions_rows', '_diff_against_latest', '_history_files',
           '_history_versions', '_expand_targets',
           'cmd_versions', 'cmd_history', 'cmd_list_file', 'cmd_grep_files', 'cmd_export_file']
