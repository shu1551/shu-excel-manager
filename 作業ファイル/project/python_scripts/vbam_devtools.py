# -*- coding: utf-8 -*-
"""vbam_devtools.py — vba_manager 分割パート: 職場向けの VBA ツール作りに寄せた手（2026-09-16）

9/16 の見直しで「作って職場へ持っていく」ときに最初に欲しくなる手を足した。
  rename-procedure  マクロの改名（宣言・呼び元・Application.Run の文字・図形の OnAction・ショートカットまで一括）
  diff-module       モジュールの差分（.bas ファイル／直前の控え／別の開いているブック）
  references        参照設定の一覧・追加・削除
  copy-modules      モジュールを別の開いているブックへ複製（同名は --overwrite）
  set-shortcut      ショートカットキーの付け替え・解除（Attribute を書くのは Excel＝MacroOptions）
  format-module     モジュールの整形（format_bas.py の中身を本体の手に）
  backup-prune      backups の間引き（N 日より古い控えを数えて --force で消す）
  status            いま走っているコマンドと進み具合（COM 不要）
  repair            マクロ修理の材料を 1 手で（本文・入口・呼び元呼び先・コンパイルの行名指し・check の error 級・控えとの差分。2026-09-17）
  export-all --history の「前回との差」の中身（_export_history_note）

下の層（vbam_core・vbam_vba）は from-import で名前を引き継ぐ。vbam_vba はこの module を上から
import しない（export-all の --history だけ関数の中で遅く import する＝循環しない）。
"""
import os
import re
import sys
import json
import time
import difflib
import shutil
import tempfile
import contextlib

from vbam_core import (SCRIPT_DIR, BACKUP_DIR, LAST_PROC_FILE, get_workbook, parse_target_and_rest, make_backup,
                       make_module_backup, _reject_extra_args, _remove_export_artifacts, smart_path_resolve,
                       check_vba_identifier, _import_module_verified, _save_with_retry, ModuleNameCollisionError,
                       _print_collision_guidance, validate_bas_encoding, _reregister_shortcuts_from_bas,
                       _addin_project_owning_proc, pythoncom, _CMD_PROGRESS_FILE, cmd_progress_read,
                       _AGENT_PROGRESS_PATH, _pid_alive, job_clock_note)
from vbam_vba import (_com_error_text, cmd_replace_module, _all_procedure_names, _suggest_similar,
                      _extract_proc, _collect_book_inventory, _analyze_calls, _compile_vbproject, _split_procedures,
                      _diag_missing_labels, _diag_on_err_typo, _extra_code_scans, _clean_vba_line, _parse_module_blocks)

_EXT_MAP = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}
_PROC_DECL = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+',
                        re.IGNORECASE)


# ================================================================
# 共通（純 Python）
# ================================================================

def _ident_pattern(name):
    """識別子を「語として」探す（前後に \\w が無い）。VBA は大小文字を区別しない。

    \\b は日本語の識別子で効かない（日本語は \\w に入るので \\b が語の中に落ちる）ため、
    前後を否定の先読み・後読みで囲む。文字列の中（Application.Run "名前"）も当たる＝それも改名の対象。
    """
    return re.compile(r'(?<!\w)' + re.escape(name) + r'(?!\w)', re.IGNORECASE)


def _rename_plan_lines(code_text, old, new):
    """1 モジュールの本文で old を new に改名したときの変更行 [(行番号, 旧行, 新行)]（純 Python）。"""
    pat = _ident_pattern(old)
    out = []
    for i, line in enumerate(code_text.replace('\r\n', '\n').split('\n'), 1):
        if pat.search(line):
            new_line = pat.sub(new, line)
            if new_line != line:
                out.append((i, line, new_line))
    return out


def _strip_attribute_lines(text):
    """.bas の Attribute 行（VB_Name・VB_ProcData 等）を落とす。CodeModule.Lines には無い行なので比べる前に揃える。"""
    return [ln for ln in text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
            if not re.match(r'^Attribute\s+\S+\s*=', ln)]


def _module_diff(a_text, b_text, label_a, label_b, max_lines=200):
    """2 つのモジュール本文の差分（unified）。戻り値: (差分の行リスト, 消えた行数, 増えた行数)。純 Python。"""
    a = _strip_attribute_lines(a_text)
    b = _strip_attribute_lines(b_text)
    while a and a[-1].strip() == '':
        a.pop()
    while b and b[-1].strip() == '':
        b.pop()
    lines = list(difflib.unified_diff(a, b, fromfile=label_a, tofile=label_b, lineterm='', n=2))
    minus = sum(1 for ln in lines if ln.startswith('-') and not ln.startswith('---'))
    plus = sum(1 for ln in lines if ln.startswith('+') and not ln.startswith('+++'))
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines] + [f"… 他 {len(lines) - max_lines} 行（--max で上限変更可）"]
    return lines, minus, plus


def _read_bas_text(path):
    with open(path, 'rb') as f:
        return f.read().decode('cp932', errors='replace')


def _module_code(comp):
    cm = comp.CodeModule
    n = int(cm.CountOfLines)
    return cm.Lines(1, n) if n else ''


def _find_comp(wb, name):
    for comp in wb.VBProject.VBComponents:
        if str(comp.Name).lower() == str(name).lower():
            return comp
    return None


def _find_open_book(xl, name):
    """開いているブックを名前（拡張子あり・なし・フルパス）で探す。無ければ None。"""
    want = os.path.basename(str(name)).lower()
    stem = os.path.splitext(want)[0]
    for wb in xl.Workbooks:
        nm = str(wb.Name).lower()
        if nm == want or os.path.splitext(nm)[0] == stem:
            return wb
    return None


def _export_to_temp(comp, tmpdir):
    """モジュールを一時フォルダへ書き出す（.frm は .frx も一緒に出る）。戻り値: パス。"""
    path = os.path.join(tmpdir, str(comp.Name) + _EXT_MAP.get(int(comp.Type), '.bas'))
    comp.Export(path)
    return path


def _proc_shortcut_keys(wb, comp):
    """そのモジュールの Attribute VB_Invoke_Func を読む → {プロシージャ名: キー}。書き出せなければ {}。"""
    if comp is None:
        return {}
    tmp = os.path.join(SCRIPT_DIR, f"_tmp_sc_{comp.Name}.bas")
    try:
        comp.Export(tmp)
        text = _read_bas_text(tmp)
    except Exception:
        return {}
    finally:
        _remove_export_artifacts(tmp)
    out = {}
    for m in re.finditer(r'^Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^"\r\n]+)"',
                         text, re.IGNORECASE | re.MULTILINE):
        key = m.group(2).replace('\\n', '\n').replace('\\r', '\r').split('\n')[0].split('\r')[0]
        if key.strip():
            out[m.group(1)] = key
    return out


def _shortcut_label(key):
    return f"Ctrl+Shift+{key.upper()}" if (len(key) == 1 and key.isupper()) else f"Ctrl+{key}"


_ATTR_INVOKE = re.compile(r'^Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=', re.IGNORECASE)


def _set_invoke_attribute(bas_text, proc, key):
    """.bas の本文で proc のショートカット属性を key に差し替える（純 Python）。key が None なら外す。

    VBE は `Attribute 名前.VB_ProcData.VB_Invoke_Func = "K\\n14"` を宣言行の直下に書く。大文字＝Ctrl+Shift。
    戻り値: (新しい本文, 宣言が見つかったか)。改行は元のまま（\\r\\n）。
    """
    nl = '\r\n' if '\r\n' in bas_text else '\n'
    lines = bas_text.split(nl)
    out = []
    found = False
    for ln in lines:
        m = _ATTR_INVOKE.match(ln)
        if m and m.group(1).lower() == proc.lower():
            continue                                # 古い属性行は落とす（空のキー " \n14" も含む）
        out.append(ln)
        if not found and _PROC_DECL.match(ln):
            name = re.sub(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+',
                          '', ln, flags=re.IGNORECASE)
            name = re.split(r'[\s(]', name, maxsplit=1)[0]
            if name.lower() == proc.lower():
                found = True
                if key:
                    out.append(f'Attribute {name}.VB_ProcData.VB_Invoke_Func = "{key}\\n14"')
    return nl.join(out), found


def _persist_shortcut(wb, comp, proc, key, target_file=None, force=False):
    """ショートカットをファイルに残す＝.bas の Attribute を書いて Remove+Import（replace-module 経路）。

    Application.MacroOptions はその場では効くが、保存して開き直すと消える（2026-09-16 実測・改名テスト.xlsm）。
    Excel が Attribute に書くのは「マクロ→オプション」の窓からだけらしい。道具が持っている確かな道は
    replace-module（.bas の Attribute → Import → 直後に MacroOptions で再登録）なので、それに乗せる。
    副作用: モジュールが VBComponents の末尾へ移る（メニューの並びが変わることがある）。
    """
    tmp = os.path.join(SCRIPT_DIR, f"_sc_set_{comp.Name}.bas")
    try:
        comp.Export(tmp)
        text = _read_bas_text(tmp)
        new, found = _set_invoke_attribute(text, proc, key)
        if not found:
            print(f"  ⚠ {comp.Name} の中に '{proc}' の宣言が見つからず、ショートカットを書けませんでした")
            return False
        with open(tmp, 'wb') as f:
            f.write(new.encode('cp932', errors='replace'))
        import argparse as _ap
        ns = _ap.Namespace(posargs=([target_file] if target_file else []) + [str(comp.Name), tmp], force=force)
        return bool(cmd_replace_module(ns))
    finally:
        _remove_export_artifacts(tmp)


def _parse_shortcut_key(text):
    """'k' / 'K' / 'ctrl+k' / 'Ctrl+Shift+K' / 'ctrl shift k' → MacroOptions に渡す 1 文字（大文字＝Shift つき）。

    読めなければ None（純 Python）。
    """
    s = str(text or '').strip()
    if not s:
        return None
    parts = [p for p in re.split(r'[+\s]+', s) if p]
    shift = False
    key = None
    for p in parts:
        low = p.lower()
        if low in ('ctrl', 'control'):
            continue
        if low == 'shift':
            shift = True
            continue
        if key is not None:
            return None
        key = p
    if key is None or len(key) != 1 or not (key.isascii() and key.isalnum()):
        return None
    if key.isalpha():
        return key.upper() if (shift or key.isupper()) else key.lower()
    return key


def _backup_series(name):
    """控えのファイル名から日時（と連番）を外した「系列」の名前（純 Python）。"""
    return re.sub(r'\d{8}_\d{6}(?:_\d+)?', '#', name)


def _prune_plan(entries, days, keep, now=None):
    """消してよい控えを決める（純 Python）。entries = [(名前, mtime, size)]。

    days 日より古いものが候補。ただし系列（同じブック・同じ用途）ごとに新しい方から keep 個は
    年齢にかかわらず残す＝「その系列の最後の 1 個」が消えて戻せなくなるのを防ぐ。
    戻り値: (消す [(名前, mtime, size)], 残す件数)
    """
    now = time.time() if now is None else now
    cutoff = now - float(days) * 86400
    by = {}
    for name, mtime, size in entries:
        by.setdefault(_backup_series(name), []).append((name, mtime, size))
    doomed = []
    kept = 0
    for series, items in by.items():
        items.sort(key=lambda e: e[1], reverse=True)
        for i, e in enumerate(items):
            if i < keep or e[1] >= cutoff:
                kept += 1
            else:
                doomed.append(e)
    doomed.sort(key=lambda e: e[1])
    return doomed, kept


def _export_dirs_diff(prev_dir, cur_dir):
    """2 つの書き出しフォルダ（export-all --history）を比べる（純 Python）。

    戻り値: {'added': [名], 'removed': [名], 'changed': [(名, 消えた行数, 増えた行数)], 'same': 件数}
    """
    def _files(d):
        return {f for f in os.listdir(d) if f.lower().endswith(('.bas', '.cls', '.frm'))} if os.path.isdir(d) else set()
    prev, cur = _files(prev_dir), _files(cur_dir)
    out = {'added': sorted(cur - prev), 'removed': sorted(prev - cur), 'changed': [], 'same': 0}
    for f in sorted(prev & cur):
        a = _read_bas_text(os.path.join(prev_dir, f))
        b = _read_bas_text(os.path.join(cur_dir, f))
        lines, minus, plus = _module_diff(a, b, 'prev', 'cur', max_lines=None)
        if minus or plus:
            out['changed'].append((f, minus, plus))
        else:
            out['same'] += 1
    return out


def _export_history_note(book_name, out_dir):
    """export-all --history の後始末: 前回の書き出しと比べて 1 行ずつ報告。前回と同じなら今回のフォルダを消す。

    置き場は SCRIPT_DIR/_exports/<ブック名>/<日時>/。戻り値: (前回の日時 or None, 同じだったか)
    """
    parent = os.path.dirname(out_dir)
    stamps = sorted(d for d in os.listdir(parent) if os.path.isdir(os.path.join(parent, d)))
    prevs = [d for d in stamps if d < os.path.basename(out_dir)]
    if not prevs:
        print(f"履歴: 初回の書き出し（{out_dir}）")
        return None, False
    prev = os.path.join(parent, prevs[-1])
    d = _export_dirs_diff(prev, out_dir)
    if not (d['added'] or d['removed'] or d['changed']):
        shutil.rmtree(out_dir, ignore_errors=True)
        print(f"履歴: 前回（{prevs[-1]}）と同じ＝今回のフォルダは残しません（{d['same']} 本）")
        return prevs[-1], True
    print(f"履歴: 前回（{prevs[-1]}）との差 … 増えた {len(d['added'])} / 消えた {len(d['removed'])} / "
          f"変わった {len(d['changed'])} / 同じ {d['same']}")
    for f in d['added']:
        print(f"  + {f}")
    for f in d['removed']:
        print(f"  - {f}")
    for f, minus, plus in d['changed']:
        print(f"  ~ {f}  -{minus} +{plus} 行")
    print(f"  中身を見る: diff-module <モジュール名> \"{os.path.join(prev, '<モジュール名>.bas')}\"")
    return prevs[-1], False


# ================================================================
# rename-procedure
# ================================================================

def cmd_rename_procedure(args):
    """マクロの改名: rename-procedure [excel_file] <旧名> <新名> [--module 名] [-y] [--dry-run] [--force]

    宣言行・呼び元（Call／裸呼び／Application.Run "名前" の文字）・図形ボタンの OnAction・ショートカットの
    再登録まで 1 手で。ReplaceLine 方式（変更行だけ）なので他の行は触らない。全モジュールを見るので
    「呼び元を直し忘れて実行時に『マクロが見つかりません』」が起きない。他のブックからの呼び出しは見ない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: rename-procedure [excel_file] <旧名> <新名> [--module 名] [-y] [--dry-run]")
        return False
    old, new = rest[0], rest[1]
    if _reject_extra_args(rest, 2, '名前は 2 つ（旧名 新名）'):
        return False
    bad = check_vba_identifier(new)
    if bad:
        print(f"エラー: 新しい名前が VBA の識別子として使えません: {bad}")
        return False
    if old.lower() == new.lower():
        print("エラー: 旧名と新名が同じです（大小文字だけの違いは VBA が保存時に揃えるので改名になりません）")
        return False
    mod_filter = getattr(args, 'module_opt', None)
    dry = getattr(args, 'dry_run', False)

    xl, wb = get_workbook(target_file, readonly=dry)

    # 宣言の持ち主を探す（同名が複数のモジュールにあれば --module で名指し）
    owners = []
    for comp in wb.VBProject.VBComponents:
        if mod_filter and str(comp.Name).lower() != mod_filter.lower():
            continue
        try:
            comp.CodeModule.ProcStartLine(old, 0)
            owners.append(comp)
        except Exception:
            pass
    if not owners:
        print(f"エラー: プロシージャ '{old}' が見つかりません" + (f"（モジュール {mod_filter}）" if mod_filter else ""))
        _suggest_similar(old, _all_procedure_names(wb))
        return False
    if len(owners) > 1:
        print(f"エラー: '{old}' が複数のモジュールにあります: " + ", ".join(str(c.Name) for c in owners))
        print("  --module で対象を指定してください（呼び元は全モジュールを直します）")
        return False
    owner = owners[0]
    for comp in wb.VBProject.VBComponents:
        try:
            comp.CodeModule.ProcStartLine(new, 0)
            print(f"エラー: 新しい名前 '{new}' は既にモジュール {comp.Name} にあります")
            return False
        except Exception:
            pass

    # 変更計画（ここでは書かない）
    plans = []             # (comp, [(行番号, 旧行, 新行)])
    decl_hits = 0
    for comp in wb.VBProject.VBComponents:
        code = _module_code(comp)
        if not code:
            continue
        changes = _rename_plan_lines(code, old, new)
        if changes:
            plans.append((comp, changes))
            if comp is owner or str(comp.Name) == str(owner.Name):
                decl_hits += sum(1 for _, ln, _ in changes if _PROC_DECL.match(ln))
    shape_plans = []       # (シート名, 図形名, 旧 OnAction, 新 OnAction)
    pat = _ident_pattern(old)
    with contextlib.suppress(Exception):
        for sh in wb.Worksheets:
            with contextlib.suppress(Exception):
                for shp in sh.Shapes:
                    oa = ''
                    with contextlib.suppress(Exception):
                        oa = str(shp.OnAction or '')
                    if oa and pat.search(oa):
                        shape_plans.append((str(sh.Name), str(shp.Name), oa, pat.sub(new, oa)))
    keys = _proc_shortcut_keys(wb, owner)
    key = keys.get(old) or next((v for k, v in keys.items() if k.lower() == old.lower()), None)

    total = sum(len(c) for _, c in plans)
    print(f"--- 改名プレビュー: {old} → {new}（宣言: {owner.Name}）"
          f" {len(plans)} モジュール / {total} 行・図形 {len(shape_plans)} 個"
          + (f"・ショートカット {_shortcut_label(key)}" if key else "") + " ---")
    for comp, changes in plans:
        print(f"[{comp.Name}] {len(changes)} 行:")
        for i, o, n in changes[:20]:
            print(f"  {i}: - {o.strip()}")
            print(f"  {i}: + {n.strip()}")
        if len(changes) > 20:
            print(f"  … 他 {len(changes) - 20} 行")
    for sh, name, oa, na in shape_plans:
        print(f"[図形 {sh}!{name}] OnAction: {oa} → {na}")
    if decl_hits == 0:
        print(f"エラー: 宣言行が見つかりません（'{old}' は Property かイベントの名前かもしれません）")
        return False
    print("-" * 40)
    if dry:
        print("（--dry-run: 書き換えていません）")
        return True
    if not getattr(args, 'yes', False):
        try:
            ans = input(f"{total} 行と図形 {len(shape_plans)} 個を書き換えますか？ (y/N): ")
        except EOFError:
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    if make_backup(wb.FullName, f"rename_{old}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    for comp, _ in plans:
        make_module_backup(wb, str(comp.Name))
    for comp, changes in plans:
        cm = comp.CodeModule
        for i, _, n in changes:
            cm.ReplaceLine(i, n)
        print(f"改名: [{comp.Name}] {len(changes)} 行")
    for sh, name, oa, na in shape_plans:
        try:
            wb.Worksheets(sh).Shapes(name).OnAction = na
            print(f"改名: [図形 {sh}!{name}] OnAction → {na}")
        except Exception as ex:
            print(f"  ⚠ 図形 {sh}!{name} の OnAction を直せませんでした: {_com_error_text(ex)}")
    wb.Save()
    note = ""
    if key:
        # 宣言行を書き換えると VBE は属性のキーを空にする（2026-09-16 実測）。MacroOptions では保存に残らないので、
        # .bas の Attribute を書いて Remove+Import（replace-module 経路）で残す
        print(f"ショートカット {_shortcut_label(key)} を {new} に付け直します（Attribute を書いて Remove+Import）")
        owner_name = str(owner.Name)          # Remove+Import の後は古い comp の参照が死ぬ＝名前で引き直す
        if _persist_shortcut(wb, owner, new, key, target_file, force=getattr(args, 'force', False)):
            after = _proc_shortcut_keys(wb, _find_comp(wb, owner_name))
            note = (f"（ショートカット {_shortcut_label(key)} → {new}・Attribute で確認済み）" if after.get(new) == key
                    else "（⚠ ショートカットを付け直せませんでした＝set-shortcut で掛け直してください）")
        else:
            note = "（⚠ ショートカットを付け直せませんでした＝set-shortcut で掛け直してください）"
    print(f"完了: {old} → {new}  {len(plans)} モジュール / {total} 行・図形 {len(shape_plans)} 個 を書き換えて保存しました{note}")
    print("  確かめる: compile（呼び忘れが無いか）／ 戻す: restore <backups の .bas>")
    return True


# ================================================================
# diff-module
# ================================================================

def cmd_diff_module(args):
    """モジュールの差分: diff-module [excel_file] <モジュール名> [比較先.bas] [--backup [名前]] [--book 他ブック] [--max N] [--json]

    比較先を省くと直前の控え（backups/<ブック>_<モジュール>_<日時>.bas の最新）と比べる。
    --book で別の開いているブックの同名モジュールと比べる。Attribute 行は両側から外して比べる。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: diff-module [excel_file] <モジュール名> [比較先.bas] [--backup [名前]] [--book 他ブック]")
        return False
    module_name = rest[0]
    other_file = rest[1] if len(rest) > 1 else None
    if _reject_extra_args(rest, 2, 'モジュール名と比較先ファイルの 2 つまで'):
        return False
    _m = getattr(args, 'max_hits', None)
    max_lines = 200 if _m is None else int(_m)
    other_book = getattr(args, 'book_opt', None)
    backup_opt = getattr(args, 'backup_opt', None)      # None（指定なし）/ True（最新）/ 名前

    xl, wb = get_workbook(target_file, readonly=True)
    comp = _find_comp(wb, module_name)
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が見つかりません")
        print("  存在するモジュール: " + ', '.join(str(c.Name) for c in wb.VBProject.VBComponents))
        return False
    mine = _module_code(comp)
    label_a = f"{wb.Name}!{comp.Name}（いま）"

    if other_book:
        ob = _find_open_book(xl, other_book)
        if ob is None:
            print(f"エラー: ブック '{other_book}' は開いていません（open で開いてから）")
            return False
        oc = _find_comp(ob, module_name)
        if oc is None:
            print(f"エラー: {ob.Name} にモジュール '{module_name}' がありません")
            return False
        other = _module_code(oc)
        label_b = f"{ob.Name}!{oc.Name}"
    elif other_file:
        path = smart_path_resolve(other_file) or other_file
        if not os.path.exists(path):
            print(f"エラー: 比較先が見つかりません: {other_file}")
            return False
        other = _read_bas_text(path)
        label_b = os.path.basename(path)
    else:
        if isinstance(backup_opt, str) and backup_opt:
            path = backup_opt if os.path.isabs(backup_opt) else os.path.join(BACKUP_DIR, backup_opt)
        else:
            stem = os.path.splitext(os.path.basename(str(wb.FullName)))[0]
            pat = re.compile(r'^' + re.escape(f"{stem}_{comp.Name}_") + r'\d{8}_\d{6}\.(bas|cls|frm)$', re.IGNORECASE)
            cands = [f for f in os.listdir(BACKUP_DIR)] if os.path.isdir(BACKUP_DIR) else []
            cands = sorted((f for f in cands if pat.match(f)),
                           key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f)))
            if not cands:
                print(f"控えがありません（backups/{stem}_{comp.Name}_<日時>.bas）。比較先の .bas を渡すか --book で別ブックと比べてください")
                return False
            path = os.path.join(BACKUP_DIR, cands[-1])
        if not os.path.exists(path):
            print(f"エラー: 控えが見つかりません: {path}")
            return False
        other = _read_bas_text(path)
        label_b = "backups/" + os.path.basename(path)

    lines, minus, plus = _module_diff(other, mine, label_b, label_a, max_lines=max_lines)
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "file": str(wb.Name), "module": str(comp.Name), "against": label_b,
                          "removed": minus, "added": plus, "diff": lines}, ensure_ascii=False), file=sys.stdout)
        return True
    if not (minus or plus):
        print(f"差分なし: {label_a} と {label_b} は同じです")
        return True
    print(f"--- 差分: {label_b} → {label_a}  -{minus} +{plus} 行 ---")
    for ln in lines:
        print(ln)
    return True


# ================================================================
# references
# ================================================================

def _ref_row(r):
    row = {}
    for k in ('Name', 'Description', 'Major', 'Minor', 'GUID', 'FullPath', 'IsBroken', 'BuiltIn', 'Type'):
        try:
            v = getattr(r, k)
            row[k] = bool(v) if k in ('IsBroken', 'BuiltIn') else (int(v) if k in ('Major', 'Minor', 'Type') else str(v))
        except Exception:
            row[k] = None
    return row


def cmd_references(args):
    """参照設定: references list [--json] / references add <GUID|パス> [--major N --minor N] / references remove <名前> -y

    眠っていたブックが動かない筆頭原因＝「参照不可」を VBE を開かずに見て直す。add は GUID なら
    AddFromGuid（版は --major/--minor・省略時 0＝最新）、パスなら AddFromFile。remove は組み込み（BuiltIn）を拒む。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    sub = (rest[0].lower() if rest else 'list')
    if sub not in ('list', 'add', 'remove', 'delete'):
        print("使い方: references list | add <GUID|パス> [--major N --minor N] | remove <名前> -y")
        return False
    xl, wb = get_workbook(target_file, readonly=(sub == 'list'))
    try:
        refs = wb.VBProject.References
    except Exception as ex:
        print(f"エラー: 参照設定を読めません（VBOM 未信頼か保護）: {_com_error_text(ex)}")
        return False

    if sub == 'list':
        rows = [_ref_row(r) for r in refs]
        if getattr(args, 'json', False):
            print(json.dumps({"success": True, "file": str(wb.Name), "references": rows}, ensure_ascii=False), file=sys.stdout)
            return True
        print(f"参照設定 {len(rows)} 本（{wb.Name}）")
        for r in rows:
            mark = "×参照不可" if r['IsBroken'] else ("組込" if r['BuiltIn'] else "  ")
            ver = f"{r['Major']}.{r['Minor']}" if r['Major'] is not None else "?"
            print(f"  [{mark}] {r['Name']:<24} {ver:<6} {r['Description'] or ''}")
            print(f"        {r['GUID'] or ''}  {r['FullPath'] or ''}")
        broken = [r for r in rows if r['IsBroken']]
        if broken:
            print(f"参照不可 {len(broken)} 本: references remove <名前> -y で外すか、正しい版を add で足す")
        return True

    if len(rest) < 2:
        print(f"使い方: references {sub} <{'GUID か パス' if sub == 'add' else '名前'}>")
        return False
    what = rest[1]
    if _reject_extra_args(rest, 2, '引数は 1 つ'):
        return False
    if make_backup(wb.FullName, "references") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False

    if sub == 'add':
        try:
            if re.fullmatch(r'\{[0-9A-Fa-f-]{36}\}', what):
                major = int(getattr(args, 'major', None) or 0)
                minor = int(getattr(args, 'minor', None) or 0)
                r = refs.AddFromGuid(what, major, minor)
            else:
                path = smart_path_resolve(what) or what
                if not os.path.exists(path):
                    print(f"エラー: ファイルが見つかりません: {what}（GUID なら {{…}} の形で）")
                    return False
                r = refs.AddFromFile(path)
        except Exception as ex:
            print(f"エラー: 参照を足せません: {_com_error_text(ex)}")
            return False
        row = _ref_row(r)
        _save_with_retry(wb)
        print(f"追加して保存しました: {row['Name']} {row['Major']}.{row['Minor']}  {row['Description'] or ''}")
        return True

    # remove
    hit = None
    for r in refs:
        row = _ref_row(r)
        if (row['Name'] or '').lower() == what.lower() or (row['GUID'] or '').lower() == what.lower():
            hit = (r, row)
            break
    if hit is None:
        print(f"エラー: 参照 '{what}' がありません（references list で名前を確認）")
        return False
    r, row = hit
    if row['BuiltIn']:
        print(f"エラー: {row['Name']} は組み込みの参照なので外せません")
        return False
    if not getattr(args, 'yes', False):
        try:
            ans = input(f"参照 {row['Name']}（{row['Description'] or ''}）を外しますか？ (y/N): ")
        except EOFError:
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    try:
        refs.Remove(r)
    except Exception as ex:
        print(f"エラー: 外せませんでした: {_com_error_text(ex)}")
        return False
    _save_with_retry(wb)
    print(f"外して保存しました: {row['Name']}" + ("（参照不可だったもの）" if row['IsBroken'] else ""))
    return True


# ================================================================
# copy-modules
# ================================================================

def cmd_copy_modules(args):
    """モジュールを別の開いているブックへ複製: copy-modules <モジュール名…> --to <ブック名> [--overwrite] [-y] [--force]

    元はアクティブブック（第 1 引数にパスも可）。行き先は開いているブック（名前・拡張子なしでも可）。
    行き先に同名があれば止まる（--overwrite で控えを取ってから Remove+Import）。シート・ThisWorkbook の
    モジュールは複製できない（中身は get で取って add-procedure）。ショートカットは取り込み後に再登録する。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    to = getattr(args, 'to_opt', None)
    if not rest or not to:
        print("使い方: copy-modules <モジュール名…> --to <ブック名> [--overwrite] [-y]")
        return False
    xl, wb = get_workbook(target_file, readonly=True)
    dst = _find_open_book(xl, to)
    if dst is None:
        print(f"エラー: 行き先のブック '{to}' は開いていません（open で開いてから。開いているブック: list-open）")
        return False
    if str(dst.FullName).lower() == str(wb.FullName).lower():
        print("エラー: 元と行き先が同じブックです")
        return False
    comps = []
    for name in rest:
        c = _find_comp(wb, name)
        if c is None:
            print(f"エラー: モジュール '{name}' が {wb.Name} にありません")
            return False
        if int(c.Type) not in (1, 2, 3):
            print(f"エラー: '{name}' はシート／ThisWorkbook のモジュールなので複製できません（get → add-procedure で中身を運ぶ）")
            return False
        comps.append(c)
    exists = [str(c.Name) for c in comps if _find_comp(dst, str(c.Name)) is not None]
    overwrite = getattr(args, 'overwrite', False)
    if exists and not overwrite:
        print(f"エラー: {dst.Name} に同名のモジュールがあります: {', '.join(exists)}（置き換えるなら --overwrite）")
        return False
    print(f"複製: {wb.Name} → {dst.Name}  " + ", ".join(str(c.Name) for c in comps)
          + (f"（上書き: {', '.join(exists)}）" if exists else ""))
    if not getattr(args, 'yes', False):
        try:
            ans = input("実行しますか？ (y/N): ")
        except EOFError:
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    dst_saved = os.path.isabs(str(dst.FullName)) and os.path.exists(str(dst.FullName))
    if exists and dst_saved:
        if make_backup(dst.FullName, "copy_modules") is None and not getattr(args, 'force', False):
            print("エラー: 行き先のバックアップが取れないため中止しました（--force で強行可）。")
            return False
    tmpdir = tempfile.mkdtemp(prefix="vbam_copy_")
    done = []
    try:
        for c in comps:
            name = str(c.Name)
            path = _export_to_temp(c, tmpdir)
            old = _find_comp(dst, name)
            if old is not None:
                make_module_backup(dst, name)
                xl.DisplayAlerts = False
                try:
                    dst.VBProject.VBComponents.Remove(old)
                    time.sleep(1.0)
                    pythoncom.PumpWaitingMessages()
                finally:
                    xl.DisplayAlerts = True
            try:
                _import_module_verified(dst, path, name)
            except ModuleNameCollisionError as ex:
                _print_collision_guidance(ex, name, None)
                return False
            # ショートカットの再登録は _import_module_verified の中で済む（二重に呼ぶと表示も二重・2026-09-16 実射）
            done.append(name)
            print(f"  取り込み: {name}{_EXT_MAP.get(int(c.Type), '')}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    if dst_saved:
        _save_with_retry(dst)
        print(f"完了: {len(done)} 本を {dst.Name} へ複製して保存しました")
    else:
        print(f"完了: {len(done)} 本を {dst.Name} へ複製しました（未保存の新しいブックなので保存はしていません）")
    return True


# ================================================================
# set-shortcut
# ================================================================

def cmd_set_shortcut(args):
    """ショートカットキー: set-shortcut <マクロ名> <キー> [--module 名] [-y] / set-shortcut <マクロ名> --clear

    キーは k / K / ctrl+k / Ctrl+Shift+K（大文字＝Shift つき）。Excel の MacroOptions で掛ける＝
    保存すると Attribute VB_ProcData.VB_Invoke_Func に書かれる（Excel の「マクロ→オプション」と同じ）。
    別のマクロが同じキーを持っていれば言う（Excel は黙って持ち主を替える）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    clear = getattr(args, 'clear', False)
    if not rest or (len(rest) < 2 and not clear):
        print("使い方: set-shortcut <マクロ名> <キー> [--module 名] [-y]   /   set-shortcut <マクロ名> --clear")
        return False
    name = rest[0]
    key = None
    if not clear:
        key = _parse_shortcut_key(rest[1])
        if key is None:
            print(f"エラー: キー '{rest[1]}' が読めません（例: k ／ K ／ ctrl+k ／ Ctrl+Shift+K。英数字 1 文字）")
            return False
    mod_filter = getattr(args, 'module_opt', None)
    xl, wb = get_workbook(target_file)
    owners = []
    for comp in wb.VBProject.VBComponents:
        if mod_filter and str(comp.Name).lower() != mod_filter.lower():
            continue
        try:
            comp.CodeModule.ProcStartLine(name, 0)
            owners.append(comp)
        except Exception:
            pass
    if not owners:
        print(f"エラー: マクロ '{name}' が見つかりません")
        _suggest_similar(name, _all_procedure_names(wb))
        return False
    if len(owners) > 1:
        print(f"エラー: '{name}' が複数のモジュールにあります: " + ", ".join(str(c.Name) for c in owners) + "（--module で名指し）")
        return False
    owner = owners[0]
    if int(owner.Type) != 1:
        print(f"エラー: ショートカットは標準モジュールの Sub にしか付きません（{owner.Name} は標準モジュールではない）")
        return False
    if not clear:
        addin = _addin_project_owning_proc(wb, name)
        if addin and not getattr(args, 'force', False):
            print(f"エラー: アドイン {addin} が同名のマクロを持っています。持ち主を奪うと更新登録後にキーが宙に浮くので止めました（--force で強行）")
            return False
        taken = []
        for comp in wb.VBProject.VBComponents:
            if int(comp.Type) != 1:
                continue
            for p, k in _proc_shortcut_keys(wb, comp).items():
                if k == key and p.lower() != name.lower():
                    taken.append(f"{comp.Name}.{p}")
        if taken:
            print(f"⚠ {_shortcut_label(key)} は既に {', '.join(taken)} に付いています（掛けると持ち主が {name} に替わります）")
    if not getattr(args, 'yes', False):
        try:
            ans = input(("ショートカットを外しますか？" if clear else f"{_shortcut_label(key)} を {name} に掛けますか？") + " (y/N): ")
        except EOFError:
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    # MacroOptions はその場では効くが保存に残らない（2026-09-16 実測）。.bas の Attribute を書いて Remove+Import で残す
    # （replace-module 経路＝控え・VB_Name 照合・Import 直後の MacroOptions 再登録・保存まで）
    owner_name = str(owner.Name)              # Remove+Import の後は古い comp の参照が死ぬ＝名前で引き直す
    if not _persist_shortcut(wb, owner, name, None if clear else key, target_file, force=getattr(args, 'force', False)):
        print("エラー: ショートカットを書けませんでした（上の replace-module の報告を見てください）")
        return False
    comp_now = _find_comp(wb, owner_name)
    after = _proc_shortcut_keys(wb, comp_now).get(name) if comp_now is not None else None
    if clear:
        if after:
            print(f"⚠ 外したはずの {name} に Attribute がまだ {_shortcut_label(after)} で残っています")
            return False
        # Import 直後の再登録は .bas の属性から掛けるので、外した鍵はセッションにも残らない
        print(f"外して保存しました: {name}（Attribute で確認済み。モジュールは末尾へ移りました）")
    else:
        if after != key:
            print(f"⚠ Attribute に書けていません（{name} = {after!r}）")
            return False
        print(f"掛けて保存しました: {_shortcut_label(key)} → {name}（Attribute で確認済み・開き直しても残る。モジュールは末尾へ移りました）")
    return True


# ================================================================
# format-module
# ================================================================

_ATTR_PROC = re.compile(r'^Attribute\s+(\S+)\.VB_ProcData\.', re.IGNORECASE)


def _format_bas_text(text):
    """format_bas.format_content を Attribute 行を壊さずに当てる（純 Python）。

    Sub 直下の Attribute X.VB_ProcData 行（ショートカット等）は整形の対象から外し、整形後に同じ宣言の直下へ戻す。
    戻り値: (整形後の本文, 足したセクションコメントの数)
    """
    from format_bas import format_content
    src = text.replace('\r\n', '\n').replace('\r', '\n')
    kept = {}                    # 宣言行の字面 → [Attribute 行…]
    body = []
    last_decl = None
    for ln in src.split('\n'):
        if _ATTR_PROC.match(ln) and last_decl is not None:
            kept.setdefault(last_decl, []).append(ln)
            continue
        if _PROC_DECL.match(ln):
            last_decl = ln.strip()
        body.append(ln)
    out = format_content('\n'.join(body))
    if kept:
        res = []
        for ln in out.split('\n'):
            res.append(ln)
            if ln.strip() in kept:
                res.extend(kept.pop(ln.strip()))
        out = '\n'.join(res)
    added = out.count("' --- ") - src.count("' --- ")
    return out, max(0, added)


def cmd_format_module(args):
    """モジュールの整形: format-module [excel_file] <モジュール名> [--apply] [-y] [--out f.bas]

    Dim をまとめて先頭へ・区切りの空行・ブロックの見出しコメント（入力チェック／対象の特定／
    メインループ／後処理…）を機械で当てる（format_bas.py の中身）。既定は整形後の .bas を
    <モジュール名>_formatted.bas に書いて差分を見せるだけ。--apply で replace-module 経路（控え・照合つき）で反映。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: format-module [excel_file] <モジュール名> [--apply] [-y] [--out f.bas]")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, 'モジュール名は 1 つ'):
        return False
    apply = getattr(args, 'apply', False)
    xl, wb = get_workbook(target_file, readonly=not apply)
    comp = _find_comp(wb, module_name)
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が見つかりません")
        return False
    if int(comp.Type) != 1:
        print(f"エラー: 整形できるのは標準モジュールだけです（{comp.Name} は標準モジュールではない）")
        return False
    tmp = os.path.join(SCRIPT_DIR, f"_fmt_src_{comp.Name}.bas")
    try:
        comp.Export(tmp)
        original = _read_bas_text(tmp)
    finally:
        _remove_export_artifacts(tmp)
    formatted, added = _format_bas_text(original)
    lines, minus, plus = _module_diff(original, formatted, "いま", "整形後", max_lines=60)
    if not (minus or plus):
        print(f"{comp.Name}: 整形しても変わりません")
        return True
    out_path = getattr(args, 'out_opt', None) or os.path.join(SCRIPT_DIR, f"{comp.Name}_formatted.bas")
    with open(out_path, 'w', encoding='cp932', errors='replace', newline='\r\n') as f:
        f.write(formatted.replace('\r\n', '\n'))
    print(f"--- 整形の差分: -{minus} +{plus} 行・見出しコメント +{added} ---")
    for ln in lines:
        print(ln)
    print(f"整形後: {out_path}")
    if not apply:
        print(f"反映するなら: format-module {comp.Name} --apply -y（replace-module 経路＝控え・VB_Name 照合つき）")
        return True
    if not getattr(args, 'yes', False):
        try:
            ans = input("反映しますか？ (y/N): ")
        except EOFError:
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    import argparse as _ap
    ns = _ap.Namespace(posargs=([target_file] if target_file else []) + [str(comp.Name), out_path],
                       force=getattr(args, 'force', False))
    return cmd_replace_module(ns)


# ================================================================
# backup-prune
# ================================================================

def cmd_backup_prune(args):
    """控えの間引き: backup-prune [--days N] [--keep K] [--force] [--json]（COM 不要）

    backups で N 日（既定 30）より古い控えを数えて見せる。系列（同じブック・同じ用途）ごとに新しい K 個（既定 1）は
    年齢にかかわらず残す。消すのは --force のときだけ。
    """
    days = int(getattr(args, 'days', None) or 30)
    keep = getattr(args, 'keep', None)
    keep = 1 if keep is None else int(keep)
    if not os.path.isdir(BACKUP_DIR):
        print(f"バックアップフォルダがありません: {BACKUP_DIR}")
        return True
    entries = []
    for name in os.listdir(BACKUP_DIR):
        p = os.path.join(BACKUP_DIR, name)
        if os.path.isfile(p):
            entries.append((name, os.path.getmtime(p), os.path.getsize(p)))
    doomed, kept = _prune_plan(entries, days, keep)
    size = sum(e[2] for e in doomed)
    total_size = sum(e[2] for e in entries)
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "days": days, "keep": keep, "total": len(entries), "total_bytes": total_size,
                          "prune": len(doomed), "prune_bytes": size, "deleted": bool(getattr(args, 'force', False)),
                          "files": [e[0] for e in doomed]}, ensure_ascii=False), file=sys.stdout)
        if not getattr(args, 'force', False):
            return True
    else:
        print(f"backups: {len(entries):,} 本 {total_size / 1048576:.1f} MB のうち、{days} 日より古い控え "
              f"{len(doomed):,} 本 {size / 1048576:.1f} MB（系列ごとに新しい {keep} 個は残す・残る {kept:,} 本）")
        by = {}
        for name, _, sz in doomed:
            k = _backup_series(name)
            by[k] = (by.get(k, (0, 0))[0] + 1, by.get(k, (0, 0))[1] + sz)
        for k, (n, sz) in sorted(by.items(), key=lambda kv: -kv[1][1])[:10]:
            print(f"  {n:>4} 本 {sz / 1048576:>6.1f} MB  {k}")
    if not doomed:
        return True
    if not getattr(args, 'force', False):
        print(f"消すなら: backup-prune --days {days} --keep {keep} --force")
        return True
    n_ok = 0
    for name, _, _ in doomed:
        try:
            os.remove(os.path.join(BACKUP_DIR, name))
            n_ok += 1
        except OSError as ex:
            print(f"  ⚠ 消せませんでした: {name} ({ex})")
    print(f"消しました: {n_ok:,} 本 {size / 1048576:.1f} MB")
    return True


# ================================================================
# status
# ================================================================

def cmd_status(args):
    """いま走っているコマンドと進み具合: status [--json]（COM 不要・Excel に触らない）

    コマンドの台帳の包み（_logged）が撃つたびに _cmd_progress.json へ「始めた／終わった」を書き、
    長い手（export-all・checkup・gate・rehearse・materials）は途中の段を note に足す。agent の進み具合も並べる。
    """
    rec = cmd_progress_read()
    agent = None
    with contextlib.suppress(Exception):
        with open(_AGENT_PROGRESS_PATH, encoding='utf-8') as f:
            agent = json.load(f)
    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "command": rec, "agent": agent}, ensure_ascii=False), file=sys.stdout)
        return True
    if not rec:
        print("コマンドの記録がありません（撃つと _cmd_progress.json に残ります）")
    else:
        alive = _pid_alive(int(rec.get('pid') or 0))
        head = f"{rec.get('cmd')} {rec.get('args') or ''}".strip()
        if rec.get('state') == 'run':
            try:
                t0 = time.mktime(time.strptime(rec['time'], '%Y-%m-%d %H:%M:%S'))
                el = f"{time.time() - t0:.0f} 秒経過"
            except Exception:
                el = ""
            if alive:
                print(f"走っています: {head}（{el}）" + (f"  いま: {rec['note']}" if rec.get('note') else ""))
            else:
                print(f"途中で止まったようです: {head}（{rec.get('time')} に始め、プロセス {rec.get('pid')} はもう居ない）"
                      + (f"  最後の段: {rec['note']}" if rec.get('note') else ""))
        else:
            ok = rec.get('ok')
            mark = "成功" if ok else ("失敗" if ok is False else "終了")
            print(f"最後のコマンド: {head} … {mark}・{float(rec.get('sec') or 0):.2f} 秒（{rec.get('time')}）")
    if agent and agent.get('line'):
        print(f"agent: {agent['line']}（{agent.get('time')} 時点）")
    return True


# ================================================================
# repair（マクロ修理の材料を 1 手で・2026-09-17）
# ================================================================
# 9/16 のポスター.xlsm: コンパイルエラー 3 件の修理に check・grep×2・get×2・call-graph・compile×3 の 9 手・5 分。
# 台帳では道具の中は 4%・AI が読む・考える間が 96%＝手数を減らすのが効く。get＋call-graph＋compile＋check＋diff-module を 1 手に。

_DECL_KIND_RE = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+',
                           re.IGNORECASE | re.MULTILINE)
_COMPILE_LABEL = {"compiled": "通過（コンパイルしました）", "already": "通過（すでにコンパイル済み）",
                  "error": "不通過（コンパイル エラー）", "unavailable": "判定なし"}


def _latest_module_backup(book_stem, module_name):
    """backups/<ブック>_<モジュール>_<日時>.bas|cls|frm の最新のパス（無ければ None）。ファイル一覧だけ（COM 不要）。"""
    if not os.path.isdir(BACKUP_DIR):
        return None
    pat = re.compile(r'^' + re.escape(f"{book_stem}_{module_name}_") + r'\d{8}_\d{6}\.(bas|cls|frm)$', re.IGNORECASE)
    cands = [f for f in os.listdir(BACKUP_DIR) if pat.match(f)]
    if not cands:
        return None
    cands.sort(key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f)))
    return os.path.join(BACKUP_DIR, cands[-1])


def _proc_body_from_bas(bas_text, proc_name):
    """.bas の本文からプロシージャ 1 本を抜く（Attribute 行は外す・前後の空行は落とす）。無ければ None（純 Python）。"""
    text = bas_text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    _, blocks, _ = _parse_module_blocks(text)
    for b in blocks:
        if b['name'].lower() == proc_name.lower():
            lines = [ln for ln in b['lines'] if not re.match(r'^Attribute\s+\S+\s*=', ln)]
            while lines and not lines[0].strip():
                lines.pop(0)
            while lines and not lines[-1].strip():
                lines.pop()
            return '\n'.join(lines)
    return None


def _repair_report(parts):
    """repair の報告を組む（純 Python）。戻り値: 行のリスト。本文（code）は切らない（agent の教訓 2026-09-04）。

    parts の鍵: book, module, proc, kind, n_lines, code, saved, shortcut, buttons[], form_callers[], auto,
      callers[(mod, proc)], callers2[], callees[(mod, proc)], callees2[], compile{ok,state,detail}, checks[str],
      backup{label, minus, plus, diff[], missing} or None
    """
    p = parts
    mod = p['module']
    out = [f"===== 修理の材料: [{mod}] {p['proc']}（{p.get('kind') or 'Sub'}・{p.get('n_lines', 0)} 行・{p.get('book', '')}） =====",
           f"モジュール  : {mod}",
           f"プロシージャ: {p['proc']}"]
    if p.get('saved'):
        out.append(f"保存先      : {p['saved']}")
    out.append("=" * 60)
    out.extend((p.get('code') or '').rstrip('\n').split('\n'))
    out.append("=" * 60)

    ent = []
    if p.get('shortcut'):
        ent.append(f"ショートカット {p['shortcut']}")
    buttons = p.get('buttons') or []
    if buttons:
        ent.append("シートのボタン " + "、".join(buttons[:5]) + (f" 他{len(buttons) - 5}" if len(buttons) > 5 else ""))
    forms = p.get('form_callers') or []
    if forms:
        ent.append("フォーム " + "、".join(forms[:5]) + (f" 他{len(forms) - 5}" if len(forms) > 5 else ""))
    if p.get('auto'):
        ent.append("自動実行イベント")
    out.append("■ 入口: " + (" ／ ".join(ent) if ent else "（ショートカット・ボタン・フォームからは呼ばれていない＝メニューか直実行）"))

    def _fmt(pairs, n):
        s = "、".join(f"[{m}] {name}" for m, name in pairs[:n])
        return s + (f" 他{len(pairs) - n}" if len(pairs) > n else "")

    callers, callers2 = p.get('callers') or [], p.get('callers2') or []
    out.append(f"■ 呼び元（直接 {len(callers)}・間接 {len(callers2)}）: " + (_fmt(callers, 12) if callers else "（なし）")
               + (f" ／ 間接: {_fmt(callers2, 8)}" if callers2 else ""))
    for mm, pr, name, lineno in (p.get('near_calls') or [])[:5]:
        # 名前違いで呼んでいる＝存在しない名前にこのマクロの名前が含まれる（ポスターの「長さ面積測定の基準設定ツール起動」）
        out.append(f"  名前違いで呼んでいる: [{mm}] {pr}:{lineno} → {name}（存在しない名前）")
    callees, callees2 = p.get('callees') or [], p.get('callees2') or []
    out.append(f"■ 呼び先（直接 {len(callees)}・間接 {len(callees2)}）: " + (_fmt(callees, 12) if callees else "（なし＝単体で完結）")
               + (f" ／ 間接: {_fmt(callees2, 8)}" if callees2 else ""))

    c = p.get('compile') or {}
    out.append(f"■ コンパイル: {_COMPILE_LABEL.get(c.get('state'), c.get('state') or '?')}  {c.get('detail') or ''}".rstrip())
    checks = p.get('checks') or []
    out.append(f"■ check（error 級・このプロシージャ）: {len(checks)} 件" + ("" if checks else "（なし）"))
    for s in checks[:10]:
        out.append(f"  {s}")
    if len(checks) > 10:
        out.append(f"  … 他 {len(checks) - 10} 件（check で全部）")

    b = p.get('backup')
    if not b:
        out.append("■ 直前の控えとの差分: （控えなし）")
    elif b.get('missing'):
        out.append(f"■ 直前の控えとの差分: {b['label']} にこのプロシージャは無い（控えの後に足した）")
    elif not (b.get('minus') or b.get('plus')):
        out.append(f"■ 直前の控えとの差分: {b['label']} と同じ")
    else:
        out.append(f"■ 直前の控えとの差分: {b['label']} → いま  -{b['minus']} +{b['plus']} 行")
        diff = b.get('diff') or []
        for ln in diff[:30]:
            out.append("  " + ln)
        if len(diff) > 30:
            out.append(f"  … 他 {len(diff) - 30} 行（diff-module {mod} で全部）")
    cn = p.get('canon')
    if cn:
        if not (cn.get('minus') or cn.get('plus')):
            out.append(f"■ {cn['label']}との差分: 同じ（壊れたのはコードでなく表の側の可能性）")
        else:
            out.append(f"■ {cn['label']} → いま  -{cn['minus']} +{cn['plus']} 行"
                       "（「前は動いていた・前は合っていた」なら、まずこの差分が原因の候補）")
            diff = cn.get('diff') or []
            for ln in diff[:30]:
                out.append("  " + ln)
            if len(diff) > 30:
                out.append(f"  … 他 {len(diff) - 30} 行")
    out.append(f"直す: 1〜数行なら code-replace \"旧\" \"新\" --module {mod} -y ／ 全文なら replace-procedure -y --module {mod}"
               "（本文は _last_proc.vba に保存済み）")
    return out


_GS_ASSIGN_RE = re.compile(r'^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=(?!=)', re.IGNORECASE)
_GS_LABEL_RE = re.compile(r'^\s*([A-Za-z_\u3040-\u30ff\u4e00-\u9fff][\w\u3040-\u30ff\u4e00-\u9fff]*):\s*$')
_GS_CALL_RE = re.compile(r'\bGoSub\s+([^\s:\']+)', re.IGNORECASE)
_GS_BOUNDARY_RE = re.compile(r'^\s*(?:For\b|Next\b|If\b.*\bThen\s*$|Else\b|ElseIf\b|End\s+If\b|Do\b|Loop\b|Select\b|Case\b|'
                             r'End\s+Select\b|With\b|End\s+With\b|While\b|Wend\b|(?:Public\s+|Private\s+)?Sub\b|End\s+Sub\b)',
                             re.IGNORECASE)


def _gosub_input_gaps(proc_lines):
    """GoSub の渡し忘れ（純 Python・2026-09-17 夜）→ [(行の添字, ラベル, 変数, 入れている呼び出しの数)]。

    Function 禁止の家の決まりで、値の受け渡しは「normInput = 値 → GoSub NormTkt → normOutput を読む」になる。
    修理の試験で、NormTkt の 3 か所の呼び出しのうち 1 か所の `normInput = tktVal` を抜いた壊し方が blind で直らなかった
    （前の呼び出しの値が残るので止まらず、数だけ合わない）。ラベルの中で読むが中では入れない変数を「受け取り」とし、
    ほかの呼び出しでは直前（GoSub・ブロックの境目まで）に入れているのに、この呼び出しでだけ入れていない所を並べる。"""
    clean = [_clean_vba_line(l) for l in proc_lines]
    bodies = {}
    for idx, l in enumerate(clean):
        m = _GS_LABEL_RE.match(l)
        if m:
            end = next((k for k in range(idx + 1, len(clean)) if re.match(r'^\s*Return\s*$', clean[k], re.IGNORECASE)), None)
            if end is not None:
                bodies[m.group(1)] = (idx, end)
    if not bodies:
        return []
    in_body = set()
    for s, e in bodies.values():
        in_body.update(range(s, e + 1))
    assigned_outside = {m.group(1).lower() for k, l in enumerate(clean) if k not in in_body
                        for m in [_GS_ASSIGN_RE.match(l)] if m}
    out = []
    for label, (s, e) in bodies.items():
        assigned_in = {m.group(1).lower() for l in clean[s + 1:e] for m in [_GS_ASSIGN_RE.match(l)] if m}
        spelled = {}
        for l in clean[s + 1:e]:
            for w in re.findall(r'[A-Za-z_]\w*', l):
                spelled.setdefault(w.lower(), w)
        inputs = sorted((set(spelled) & assigned_outside) - assigned_in)
        if not inputs:
            continue
        sites = [k for k, l in enumerate(clean) if k not in in_body
                 and any(g.lower() == label.lower() for g in _GS_CALL_RE.findall(l))]
        if len(sites) < 2:
            continue
        seg = {}
        for c in sites:
            got = set()
            for k in range(c - 1, -1, -1):
                l = clean[k]
                if k in in_body or _GS_CALL_RE.search(l) or _GS_BOUNDARY_RE.match(l):
                    break
                m = _GS_ASSIGN_RE.match(l)
                if m:
                    got.add(m.group(1).lower())
            seg[c] = got
        for v in inputs:
            have = [c for c in sites if v in seg[c]]
            miss = [c for c in sites if v not in seg[c]]
            if have and miss and len(have) >= len(miss):
                out.extend((c, label, spelled[v], len(have)) for c in miss)
    return sorted(out)


_NA_TARGET_RE = re.compile(r'(?:^|:|\bThen\b|\bElse\b)\s*(?:Set\s+|Let\s+)?([A-Za-z_]\w*)\s*(?:\([^=]*\))?\s*=(?!=)',
                           re.IGNORECASE)
_NA_FOR_RE = re.compile(r'\bFor\s+(?:Each\s+)?([A-Za-z_]\w*)', re.IGNORECASE)
_NA_REDIM_RE = re.compile(r'\bReDim\s+(?:Preserve\s+)?([A-Za-z_]\w*)', re.IGNORECASE)
_NA_DIM_RE = re.compile(r'^\s*(?:Dim|Static|Private|Public)\s+(.+)$', re.IGNORECASE)
_NA_INPUT_RE = re.compile(r'\b(?:Input|Get|Line\s+Input)\s+#?\w+\s*,\s*([A-Za-z_]\w*)', re.IGNORECASE)


def _never_assigned_reads(proc_lines):
    """Dim だけして一度も値を入れずに読んでいる変数（純 Python・2026-09-17 夜）→ [(行の添字, 変数)]（読む最初の行）。

    Option Explicit を使わない家の決まりでは、`u = parts(0) & …` の 1 行が抜けても u は空のまま読まれ、止まらず結果だけ違う
    （修理の試験で表を整えるの電話番号が区切られなくなった）。Dim で宣言した名前に絞る（xlUp などの定数を拾わない）。"""
    clean = [_clean_vba_line(l) for l in proc_lines]
    declared = {}
    for k, l in enumerate(clean):
        m = _NA_DIM_RE.match(l)
        if m:
            for part in m.group(1).split(','):
                w = re.match(r'\s*([A-Za-z_]\w*)', part)
                if w:
                    declared.setdefault(w.group(1).lower(), w.group(1))
    if not declared:
        return []
    assigned = set()
    for l in clean:
        for rx in (_NA_TARGET_RE, _NA_FOR_RE, _NA_REDIM_RE, _NA_INPUT_RE):
            assigned.update(m.group(1).lower() for m in rx.finditer(l))
    out = []
    for low, name in declared.items():
        if low in assigned:
            continue
        for k, l in enumerate(clean):
            if _NA_DIM_RE.match(l):
                continue
            if re.search(r'(?<![\w.])' + re.escape(name) + r'(?!\w)', l, re.IGNORECASE):
                out.append((k, name))
                break
    return sorted(out)


_RB_LOOP_OPEN_RE = re.compile(r'^\s*(?:For\b|Do\b|While\b)', re.IGNORECASE)
_RB_LOOP_CLOSE_RE = re.compile(r'^\s*(?:Next\b|Loop\b|Wend\b)', re.IGNORECASE)


def _read_before_assign(proc_lines):
    """入れる前に読んでいる変数（純 Python・2026-09-17 夜）→ [(読んだ行の添字, 変数)]。

    上から読んで、その行より前に一度も値を入れていない（GoSub の塊の中の読みは呼び元が入れる＝数えない・
    その行を囲むループの中で後から入れるのは前の周回の値＝数えない・前で呼んだ GoSub の塊が入れるのは数えない）。
    使い回しの作業用変数（u など）で、使う直前の `u = …` を 1 行抜いた壊れ方を拾う（修理の試験・表を整えるの電話番号）。
    Dim で宣言した名前だけ見る。"""
    clean = [_clean_vba_line(l) for l in proc_lines]
    declared = {}
    for l in clean:
        m = _NA_DIM_RE.match(l)
        if m:
            for part in m.group(1).split(','):
                w = re.match(r'\s*([A-Za-z_]\w*)', part)
                if w:
                    declared.setdefault(w.group(1).lower(), w.group(1))
    if not declared:
        return []
    bodies, in_body = {}, {}
    for idx, l in enumerate(clean):
        m = _GS_LABEL_RE.match(l)
        if m:
            end = next((k for k in range(idx + 1, len(clean)) if re.match(r'^\s*Return\s*$', clean[k], re.IGNORECASE)), None)
            if end is not None:
                bodies[m.group(1).lower()] = (idx, end)
                for k in range(idx, end + 1):
                    in_body[k] = m.group(1).lower()
    loops, stack = [], []
    ifs, if_stack = [], []
    for k, l in enumerate(clean):
        if _RB_LOOP_OPEN_RE.match(l) and not re.search(r'\bNext\b', l, re.IGNORECASE):
            stack.append(k)
        elif _RB_LOOP_CLOSE_RE.match(l) and stack:
            loops.append((stack.pop(), k))
        if re.match(r'^\s*If\b.*\bThen\s*$', l, re.IGNORECASE):
            if_stack.append(k)
        elif re.match(r'^\s*End\s+If\b', l, re.IGNORECASE) and if_stack:
            ifs.append((if_stack.pop(), k))

    def _carried(a, r, s):
        """ループ（s から）の中の a の代入が、次の周回の r まで必ず届くか＝a を囲む If が、ループの中では全部 r も囲む。"""
        return all(i0 < r < i1 for i0, i1 in ifs if i0 < a < i1 and i0 > s)

    def _targets(l):
        got = set()
        for rx in (_NA_TARGET_RE, _NA_FOR_RE, _NA_REDIM_RE, _NA_INPUT_RE):
            got.update(m.group(1).lower() for m in rx.finditer(l))
        return got

    assign_at = {}
    for k, l in enumerate(clean):
        for v in _targets(l):
            assign_at.setdefault(v, []).append(k)
    body_assigns = {lab: {v for k in range(s, e + 1) for v in _targets(clean[k])} for lab, (s, e) in bodies.items()}
    out = []
    for low, name in declared.items():
        tok = re.compile(r'(?<![\w.])' + re.escape(name) + r'(?!\w)', re.IGNORECASE)
        for r, l in enumerate(clean):
            if r in in_body or _NA_DIM_RE.match(l) or not tok.search(l):
                continue
            n_all = len(tok.findall(l))
            n_tgt = sum(1 for rx in (_NA_TARGET_RE, _NA_FOR_RE, _NA_REDIM_RE, _NA_INPUT_RE)
                        for m in rx.finditer(l) if m.group(1).lower() == low)
            if n_all <= n_tgt:
                continue                                   # 入れるだけの行
            reached = any(a <= r and a not in in_body for a in assign_at.get(low, []))
            if not reached:
                reached = any(s < r < e and any(s <= a <= e and _carried(a, r, s) for a in assign_at.get(low, []))
                              for s, e in loops)
            if not reached:
                for k in range(r):
                    for g in _GS_CALL_RE.findall(clean[k]):
                        if low in body_assigns.get(g.lower().rstrip(':'), set()):
                            reached = True
            if not reached:
                out.append((r, name))
            break                                          # いちばん上の読みだけ見れば足りる
    return sorted(out)


def _forge_canon_diff(proc, code):
    """鍛えて登録した時の正本（_agent_forge.json の code）と、いまの本文の差分。台帳に無ければ None（2026-09-17 夜）。
    修理の試験で、「いつもの表では合うのにこの表だけ」の依頼に AI が全部の表に効く列の位置を動かして
    この表だけ合わせた（嘘の完了）。正本を持っているのに材料に出していなかった＝壊れた 1 行は差分で見える。"""
    path = os.path.join(SCRIPT_DIR, '_agent_forge.json')
    if os.environ.get('VBAM_REPAIR_NO_CANON') == '1' or not os.path.isfile(path):
        return None
    try:
        with open(path, encoding='utf-8') as f:
            ledger = json.load(f)
    except (OSError, ValueError):
        return None
    for name, case in (ledger or {}).items():
        if str(case.get('sub') or '').lower() == str(proc).lower() and case.get('code'):
            label = f"鍛えて登録した時の正本（agent --forged の「{name}」）"
            diff, minus, plus = _module_diff(case['code'], code, label, "いま", max_lines=None)
            return {'label': label, 'minus': minus, 'plus': plus, 'diff': diff}
    return None


def cmd_repair(args):
    """マクロ修理の材料を 1 手で: repair [excel_file] <マクロ名> [--module 名] [--json]（読むだけ）

    ① 本文（get と同じ。_last_proc.vba にも保存＝replace-procedure がそのまま使える）
    ② 入口（ショートカット・シートのボタン・フォームのボタン・自動実行）③ 呼び元・呼び先（間接 1 段まで。impact と同じ材料）
    ④ 全体コンパイル（落ちた行の名指しつき）⑤ check の error 級（VBM004/005/012/013/014）でこのプロシージャに当たるもの
    ⑥ 直前の控え（backups/<ブック>_<モジュール>_<日時>.bas）との、このプロシージャの範囲だけの差分
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: repair [excel_file] <マクロ名> [--module 名] [--json]")
        return False
    macro = rest[0]
    if _reject_extra_args(rest, 1, 'マクロ名は 1 つ'):
        return False
    mod_filter = getattr(args, 'module_opt', None)
    xl, wb = get_workbook(target_file, readonly=True)

    # ① 本文
    try:
        comp_name, code = _extract_proc(wb, mod_filter, macro)
    except ValueError as ex:
        print(f"エラー: '{macro}' が複数のモジュールにあります: {', '.join(ex.args[0])}（--module で名指し）")
        return False
    if comp_name is None:
        print(f"エラー: プロシージャ '{macro}' が見つかりません" + (f"（モジュール {mod_filter}）" if mod_filter else ""))
        _suggest_similar(macro, _all_procedure_names(wb))
        return False
    with open(LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(code)
    m = _DECL_KIND_RE.search(code)
    kind = ' '.join(m.group(1).split()) if m else 'Sub'
    n_lines = code.rstrip('\n').count('\n') + 1

    # ② ③ 入口と呼び出し関係（impact・call-graph と同じ棚卸し）
    inv = _collect_book_inventory(xl, wb, include_vba=True, quiet=True)
    res = _analyze_calls(inv)
    known, edges = res['known'], res['edges']
    hit = known.get(macro.lower())
    root = hit[0] if hit else macro
    mod_of = {v[0]: v[1] for v in known.values()}
    form_mods = {mm['name'] for mm in inv['modules'] if mm['type'] == 3}
    me = (comp_name, root)
    callers = sorted({(mm, pr) for (mm, pr), cs in edges.items() if root in cs and pr != '(宣言部)'} - {me})
    caller_names = {pr for _, pr in callers}
    callers2 = sorted({(mm, pr) for (mm, pr), cs in edges.items() if pr != '(宣言部)' and (cs & caller_names)}
                      - set(callers) - {me})
    callee_names = sorted(edges.get(me, set()) - {root})
    callees = [(mod_of.get(n, '?'), n) for n in callee_names]
    callee2_names = sorted({c for n in callee_names for c in edges.get((mod_of.get(n), n), set())}
                           - set(callee_names) - {root})
    callees2 = [(mod_of.get(n, '?'), n) for n in callee2_names]
    auto_names = set()
    with contextlib.suppress(Exception):
        auto_names = {name for _, name, _ in _extra_code_scans(inv)['auto_exec']}
    near_calls = [(mm, pr or '(宣言部)', name, lineno) for mm, pr, name, lineno, _text in res['unresolved']
                  if name.lower() != root.lower() and (root.lower() in name.lower() or name.lower() in root.lower())]

    # ④ コンパイル（落ちた行の名指しつき）
    comp_doc = _compile_vbproject(xl, wb)

    # ⑤ check の error 級（このプロシージャに当たるものだけ。読んだコードで純 Python）
    mine = next((mm for mm in inv['modules'] if mm['name'] == comp_name), None)
    lines = (mine['code'] if mine else '').split('\r\n')
    procs = _split_procedures(lines)
    span = next(((p['start'], p['end']) for p in procs if p['name'].lower() == root.lower()), None)
    checks = []
    for pname, lineno, target, text in _diag_missing_labels(lines, procs):
        if pname.lower() == root.lower():
            checks.append(f"VBM012 行 {lineno}: 飛び先ラベル '{target}' がありません: {text}")
    for lineno, word, text in _diag_on_err_typo(lines):
        if span and span[0] <= lineno - 1 <= span[1]:
            checks.append(f"VBM013 行 {lineno}: 'On {word}' は On Error の打ち間違い: {text}")
    for mm, pr, name, lineno, text in res['unresolved']:
        if str(name).startswith('Run "'):
            continue                              # Run は実行時に探す（アドイン等）＝check の VBM014 と同じく error にしない
        if mm == comp_name and (pr or '').lower() == root.lower():
            checks.append(f"VBM014 行 {lineno}: 存在しないマクロ '{name}' を呼んでいます: {text}")
    if span:
        plines = lines[span[0]:span[1] + 1]
        for rel, label, var, n_have in _gosub_input_gaps(plines):
            checks.append(f"GoSub の渡し忘れ 行 {span[0] + rel + 1}: GoSub {label} の前で {var} を入れていません"
                          f"（ほかの {n_have} か所の呼び出しでは直前に入れている）: {plines[rel].strip()}")
        for rel, var in _read_before_assign(plines):
            checks.append(f"入れる前に読んでいる 行 {span[0] + rel + 1}: {var} をこの行より前で一度も入れていません"
                          f"（空のまま読む＝止まらず結果だけ違う。入れる行が抜けていないか）: {plines[rel].strip()}")
    decl_re = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+', re.IGNORECASE)
    end_re = re.compile(r'^\s*End\s+(?:Sub|Function)\b', re.IGNORECASE)
    n_decl = sum(1 for ln in lines if decl_re.match(_clean_vba_line(ln)))
    n_end = sum(1 for ln in lines for s in _clean_vba_line(ln).split(':') if end_re.match(s))
    if n_decl != n_end:
        checks.append(f"VBM004/005 モジュール {comp_name}: Sub/Function の宣言 {n_decl} に対して End が {n_end}（閉じ忘れ）")

    # ⑥ 直前の控えとの差分（このプロシージャの範囲だけ）
    backup = None
    stem = os.path.splitext(os.path.basename(str(wb.FullName)))[0]
    bpath = _latest_module_backup(stem, comp_name)
    if bpath:
        label = "backups/" + os.path.basename(bpath)
        old = _proc_body_from_bas(_read_bas_text(bpath), root)
        if old is None:
            backup = {'label': label, 'minus': 0, 'plus': 0, 'diff': [], 'missing': True}
        else:
            diff, minus, plus = _module_diff(old, code, label, "いま", max_lines=None)
            backup = {'label': label, 'minus': minus, 'plus': plus, 'diff': diff}

    parts = {'book': str(wb.Name), 'module': comp_name, 'proc': root, 'kind': kind, 'n_lines': n_lines, 'code': code,
             'saved': LAST_PROC_FILE, 'shortcut': inv['shortcuts'].get(root),
             'buttons': list(res.get('onaction', {}).get(root, [])),
             'form_callers': [f"{mm}.{pr}" for mm, pr in callers if mm in form_mods], 'auto': root in auto_names,
             'callers': callers, 'callers2': callers2, 'callees': callees, 'callees2': callees2, 'near_calls': near_calls,
             'compile': comp_doc, 'checks': checks, 'backup': backup, 'canon': _forge_canon_diff(root, code)}
    if getattr(args, 'json', False):
        doc = dict(parts)
        doc['success'] = True
        print(json.dumps(doc, ensure_ascii=False), file=sys.stdout)
        return True
    for ln in _repair_report(parts):
        print(ln)
    return True




__all__ = ['_ident_pattern', '_rename_plan_lines', '_strip_attribute_lines', '_module_diff', '_parse_shortcut_key',
           '_shortcut_label', '_backup_series', '_prune_plan', '_export_dirs_diff', '_export_history_note',
           '_format_bas_text', '_find_open_book', '_proc_shortcut_keys', '_repair_report', '_latest_module_backup',
           '_proc_body_from_bas',
           'cmd_rename_procedure', 'cmd_diff_module', 'cmd_references', 'cmd_copy_modules', 'cmd_set_shortcut',
           'cmd_format_module', 'cmd_backup_prune', 'cmd_status', 'cmd_repair']
