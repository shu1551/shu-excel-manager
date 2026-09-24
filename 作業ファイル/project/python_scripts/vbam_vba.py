# -*- coding: utf-8 -*-
"""vbam_vba.py — vba_manager 分割パート: VBA コマンド実装（list/get/replace/export/checkup/フォーム lint 等）

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
# ================================================================
# コマンド実装
# ================================================================

def _find_duplicate_procedures(norm_text):
    """.bas 内の Sub/Function 名の重複を機械的に検出（重複プロシージャ挿入の検知）。

    Property Get/Let/Set は同名が正常なので対象外（Sub/Function のみ）。
    戻り値: {名前: [行番号, ...], ...}（重複のあるものだけ）
    """
    sub_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function)\s+(?P<name>[^\s\(]+)',
        re.IGNORECASE
    )
    seen = {}       # 小文字名 -> [行番号, ...]（照合用）
    disp = {}       # 小文字名 -> 最初に現れた綴り（表示用）
    for idx, line in enumerate(norm_text.split('\n'), 1):
        m = sub_pattern.match(line)
        if m:
            # VBA の識別子は大小文字を区別しない。Calc と calc は「同名」で
            # コンパイルエラーになるため、照合は小文字化して行う
            # （区別して数えると「重複なし・取り込み可」と誤って報告する）。
            # ただし報告は元の綴りで出す（キーごと小文字に潰すと表示が化ける）
            key = m.group('name').lower()
            seen.setdefault(key, []).append(idx)
            disp.setdefault(key, m.group('name'))
    return {disp[key]: lns for key, lns in seen.items() if len(lns) > 1}


def _find_consecutive_dup_lines(norm_text):
    """連続して同一の非空コード行を検出（重複挿入の臭い。例: On Error Resume Next ×2）。

    空行・コメント行は対象外。空行を挟むとリセット（空行連続は正常）。
    入れ子で正常に連続しうるブロック終端等（End If / End With / Next / Loop / Else / Wend）は
    重複扱いしない（実モジュールでの誤検知を避ける）。
    戻り値: [(行番号, 行内容), ...]
    """
    struct = re.compile(r'^(end\b|else\b|elseif\b|next\b|loop\b|wend\b)', re.IGNORECASE)
    hits = []
    prev = None
    for idx, raw in enumerate(norm_text.split('\n'), 1):
        s = raw.strip()
        if s and not s.startswith("'") and s == prev and not struct.match(s):
            hits.append((idx, s))
        prev = s if s else None
    return hits






def cmd_check_bas(args):
    """.bas を取り込む前の単体検査（COM不要）。複数ファイル可。

    バイパス経路（vba_manager を通さず手書きスクリプトで .bas を作る）でも、取り込み前に
    この1コマンドで「文字コード事故 / 改行二重化 / プロシージャ重複 / 識別子規則違反 /
    連続重複行」を機械的に検査できる。COM接続が落ちていても動くのが要点（安全確認を不安全な手順と同じ手数にする）。
    --fix を付けると改行二重化だけ CP932 のまま自動修正する
    （重複は判断が要るので自動修正しない＝Pythonは機械的検査まで）。
    """
    if not args.posargs:
        print("使い方: py vba_manager.py check-bas <file.bas> [file2.bas ...] [--fix] [--json]")
        return False
    results = []
    ok_all = True
    for p in args.posargs:
        ok = _check_bas_one(p, fix=getattr(args, 'fix', False))
        results.append({"file": p, "ok": bool(ok)})
        if not ok:
            ok_all = False
    if len(args.posargs) > 1:
        print(f"===== 一括検査: {len(args.posargs)}本 → {'すべて取り込み可' if ok_all else '⚠ NGあり'} =====")
    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": ok_all, "files": results}, ensure_ascii=False),
              file=sys.stdout)
    return ok_all


def _check_bas_one(path, fix=False):
    """check-bas の1ファイルぶんの検査本体（True=取り込み可）"""
    if not os.path.isfile(path):
        print(f"エラー: ファイルが見つかりません: {path}")
        return False

    name = os.path.basename(path)
    print(f"===== .bas 取り込み前検査: {name} =====")
    problems = 0
    warnings = 0

    # 1. 文字コード事故（UTF-8化 / BOM）
    if not validate_bas_encoding(path):
        problems += 1
    else:
        print("  [OK] 文字コード: CP932 として安全")

    # 2. 改行二重化（\r\r\n）
    try:
        fixed_bytes, raw_bytes, was_doubled = normalize_bas_newlines(path)
    except Exception as e:
        print(f"エラー: 改行検査に失敗 ({e})")
        return False
    if was_doubled:
        before = len(re.split(r'\r\n|\r|\n', raw_bytes.decode('cp932')))
        after = len(re.split(r'\r\n|\r|\n', fixed_bytes.decode('cp932')))
        if fix:
            # in-place 書き換えなので、書く前に元バイト列を退避しておく（undo 導線）
            try:
                bak = path + f".bak_{time.strftime('%Y%m%d_%H%M%S')}"
                with open(bak, 'wb') as f:
                    f.write(raw_bytes)
                print(f"  退避: {bak}")
            except Exception as e:
                print(f"  [WARN] 退避に失敗しました（{e}）。修正は続行します")
            with open(path, 'wb') as f:
                f.write(fixed_bytes)
            print(f"  [FIXED] 改行二重化を修正しました: {before}行 → {after}行")
        else:
            print(f"  [NG] 改行二重化を検知: {before}行 → {after}行（--fix で修正可）")
            problems += 1
    else:
        print("  [OK] 改行: 正規 CRLF（二重化なし）")

    # 3/4 の検査は現在のファイル内容（--fix 後を反映）に対して行う
    with open(path, 'rb') as f:
        norm_text = re.sub(r'\r\n|\r', '\n', f.read().decode('cp932'))

    # 3. プロシージャ名の重複（重複挿入の検知・自動修正しない）
    dups = _find_duplicate_procedures(norm_text)
    if dups:
        print("  [NG] Sub/Function 名の重複を検知（重複挿入の疑い・自動修正しません）:")
        for nm, lns in dups.items():
            print(f"        {nm}  (行 {', '.join(map(str, lns))})")
        problems += 1
    else:
        print("  [OK] プロシージャ名: 重複なし")

    # 3b. プロシージャ名の識別子規則（先頭 _ 等。VBE は黙って受け入れコンパイルで死ぬ）
    bad_names = _find_invalid_procedure_names(norm_text)
    if bad_names:
        print("  [NG] VBA の識別子規則に反するプロシージャ名を検知（コンパイルエラーになります）:")
        for ln, _nm, reason in bad_names:
            print(f"        行{ln}: {reason}")
        problems += 1
    else:
        print("  [OK] プロシージャ名: 識別子規則OK")

    # 4. 連続する同一コード行（On Error Resume Next ×2 等の臭い）
    cdl = _find_consecutive_dup_lines(norm_text)
    if cdl:
        print("  [WARN] 連続する同一コード行（重複挿入の臭い・要確認）:")
        for ln, s in cdl[:20]:
            disp = s if len(s) <= 60 else s[:60] + '…'
            print(f"        行{ln}: {disp}")
        if len(cdl) > 20:
            print(f"        … 他 {len(cdl) - 20} 件")
        warnings += 1
    else:
        print("  [OK] 連続重複行: なし")

    print(f"----- 結果: 問題 {problems} / 警告 {warnings} -----")
    if problems:
        print("  ⚠ 問題があります。修正してから replace-module / replace-procedure で取り込んでください。")
        return False
    print("  取り込み可（手書きでも、取り込みは replace-module / replace-procedure 経由を推奨）。")
    return True


# ===== 追加診断4本（xlflow の規則から輸入・2026-08-28）=====
# 選んだ基準は「走らせても出ない」欠陥だけ。走らせれば分かるもの（Select/Activate・
# 命名・暗黙 Variant 等）は入れない——本人の10秒テストの方が速くて確実だから。
#   VBM008 ← xlflow VBA203  Application の状態を戻さない
#   VBM009 ← xlflow VBA221  呼んだヘルパーが状態を変えたまま返る
#   VBM010 ← xlflow VBA240  モジュール変数の書き手が散っている
#   VBM011 ← xlflow VBA215  Find/Replace の引数省略（Excel が前回の設定を覚えている）

# 変えたら戻す約束の Application プロパティ → (表示名, 「戻す側」の値)
_APP_STATE_PROPS = {
    "screenupdating":   ("ScreenUpdating",   "true"),
    "enableevents":     ("EnableEvents",     "true"),
    "displayalerts":    ("DisplayAlerts",    "true"),
    "displaystatusbar": ("DisplayStatusBar", "true"),
    "interactive":      ("Interactive",      "true"),
    "calculation":      ("Calculation",      "xlcalculationautomatic"),
    "cursor":           ("Cursor",           "xlnormal"),
    "enablecancelkey":  ("EnableCancelKey",  "xlinterrupt"),
    "asktoupdatelinks": ("AskToUpdateLinks", "true"),
}


def _clean_vba_line(raw):
    """文字列リテラルを潰してからコメントを落とす。先に ' で切ると
    文字列内のアポストロフィで行が切れる（cmd_check 本体と同じ掃除）"""
    return re.sub(r'"[^"]*"', '""', raw).split("'")[0].strip()


def _logical_line(lines, idx):
    """行継続 ( _ ) を畳んだ1文と、畳んだ行数を返す"""
    buf = _clean_vba_line(lines[idx])
    used = 1
    while buf.endswith('_') and idx + used < len(lines):
        buf = buf[:-1] + ' ' + _clean_vba_line(lines[idx + used])
        used += 1
    return buf, used


def _split_procedures(lines):
    """コードを [{name, kind, start, end}] に切り分ける（0始まりの行番号）"""
    proc_re = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s\(\)]+)',
        re.IGNORECASE)
    end_re = re.compile(r'^end\s+(sub|function|property)\b', re.IGNORECASE)
    out, cur = [], None
    for idx, raw in enumerate(lines):
        clean = _clean_vba_line(raw)
        if clean.lower().startswith('rem '):
            continue
        if cur is None:
            m = proc_re.match(clean)
            if m:
                cur = {"name": m.group(2), "kind": m.group(1).split()[0].lower(),
                       "start": idx, "end": idx}
            continue
        segs = [s.strip() for s in clean.split(':')]
        if any(end_re.match(s) for s in segs):
            cur["end"] = idx
            out.append(cur)
            cur = None
    if cur is not None:
        cur["end"] = len(lines) - 1
        out.append(cur)
    return out


def _diag_app_state(lines, procs):
    """VBM008: Application の状態を変えたまま戻さないプロシージャ
    → {プロシージャ名: [(プロパティ, 行番号), ...]}"""
    leaks = {}
    assign_re = re.compile(r'\bApplication\s*\.\s*([A-Za-z]+)\s*=\s*(.+)$',
                           re.IGNORECASE)
    for p in procs:
        seen = {}
        for idx in range(p["start"], p["end"] + 1):
            m = assign_re.search(_clean_vba_line(lines[idx]))
            if not m:
                continue
            prop = m.group(1).lower()
            if prop in _APP_STATE_PROPS:
                seen.setdefault(prop, []).append((idx, m.group(2).strip().lower()))
        for prop, hits in seen.items():
            disp, restore_val = _APP_STATE_PROPS[prop]
            if len(hits) >= 2:
                continue                      # 変えて、戻している
            idx, val = hits[0]
            if val.rstrip('_').strip() == restore_val:
                continue                      # 戻す側だけを書いた後始末用
            leaks.setdefault(p["name"], []).append((disp, idx + 1))
    return leaks


def _diag_stateful_find(lines, procs):
    """VBM011: Range.Find / Replace が LookAt を省いている
    （省くと Excel が前回の検索ダイアログ／マクロの設定を引き継ぐ）"""
    out = []
    call_re = re.compile(r'\.\s*(Find|Replace)\s*\(', re.IGNORECASE)
    for p in procs:
        idx = p["start"]
        while idx <= p["end"]:
            stmt, used = _logical_line(lines, idx)
            m = call_re.search(stmt)
            if m:
                packed = stmt.lower().replace(' ', '')
                if 'lookat:=' not in packed:
                    missing = ['LookAt']
                    for a, label in (('searchorder:=', 'SearchOrder'),
                                     ('matchcase:=', 'MatchCase')):
                        if a not in packed:
                            missing.append(label)
                    out.append((p["name"], m.group(1), idx + 1, missing))
            idx += used
    return out


def _diag_module_state(lines, procs):
    """VBM010: モジュール変数の書き手が2か所以上に散っている
    → [(変数名, 宣言行, [書き手プロシージャ...])]"""
    first = procs[0]["start"] if procs else len(lines)
    decl_re = re.compile(
        r'^\s*(?:Public|Private|Dim|Global)\s+'
        r'(?!Declare\b|Const\b|Type\b|Enum\b|Sub\b|Function\b|Property\b|WithEvents\b)(.+)$',
        re.IGNORECASE)
    mvars = []
    for idx in range(0, min(first, len(lines))):
        m = decl_re.match(_clean_vba_line(lines[idx]))
        if not m:
            continue
        for part in m.group(1).split(','):
            v = re.split(r'\s+As\s+', part.strip(), flags=re.IGNORECASE)[0]
            v = re.sub(r'\(.*\)', '', v).strip()
            v = re.sub(r'[%&\$#!@]$', '', v).strip()
            if re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', v):
                mvars.append((v, idx))
    out = []
    for v, decl_idx in mvars:
        wre = re.compile(r'(?:^|\bSet\s+)' + re.escape(v) +
                         r'\s*(?:\([^)]*\))?\s*=(?!=)', re.IGNORECASE)
        writers = []
        for p in procs:
            for idx in range(p["start"], p["end"] + 1):
                if wre.search(_clean_vba_line(lines[idx])):
                    writers.append(p["name"])
                    break
        if len(set(writers)) >= 2:
            out.append((v, decl_idx + 1, sorted(set(writers))))
    return out


def _diag_state_callers(book_procs, leaks_by_proc):
    """VBM009: 状態を戻さないプロシージャを呼んでいる箇所（手続き間）"""
    out = []
    if not leaks_by_proc:
        return out
    callee_names = {name: (mod, props)
                    for (mod, name), props in leaks_by_proc.items()}
    for (mod, pname), body in book_procs.items():
        for off, raw in enumerate(body):
            clean = _clean_vba_line(raw)
            if not clean:
                continue
            for callee, (cmod, props) in callee_names.items():
                if callee == pname:
                    continue
                if re.search(r'(?:^|\bCall\s+|[\s=(,])' + re.escape(callee) + r'\b',
                             clean, re.IGNORECASE):
                    out.append({
                        "module": mod, "caller": pname, "callee": callee,
                        "callee_module": cmod,
                        "props": [f"Application.{p}" for p, _ in props],
                    })
                    break
    return out


# VBM012（2026-09-17）: 飛び先のラベルがプロシージャ内に無い。コンパイルは「行ラベルが定義されていません」で
# 止まるが、check は素通りしていた（ポスター.xlsm の 写真の貼り付け・On Error GoTo Errorgo にラベル無し）。
# ラベル＝行頭の「名前:」（キーワードの Else: 等は除く）か、行番号（10 …）。飛び先＝On Error GoTo／GoTo／GoSub／Resume の後ろ。
_LABEL_KEYWORDS = {'else', 'case', 'end', 'loop', 'next', 'wend', 'do', 'exit', 'stop', 'resume', 'return', 'with',
                   'select', 'if', 'then', 'sub', 'function', 'property', 'private', 'public', 'dim', 'set', 'let',
                   'call', 'goto', 'gosub', 'on', 'error', 'rem', 'debug', 'msgbox', 'while', 'for', 'each', 'static',
                   'const', 'redim', 'erase', 'print', 'open', 'close', 'input', 'write', 'get', 'put', 'seek', 'lock',
                   'unlock', 'name', 'kill', 'mkdir', 'rmdir', 'chdir', 'chdrive', 'beep', 'randomize', 'date', 'time'}
_LABEL_RE = re.compile(r'^\s*([A-Za-z_-￿][\w-￿]*):(?:\s|$)')
_LINE_NUMBER_RE = re.compile(r'^\s*(\d+)(?:\s|:|$)')
# (?<![.\w]) … `Application.GoTo Reference:="Print_Area"` はメソッド呼び出し（飛び先ではない）。ポスターの ヨコサイズ変更2 で誤検出した
_JUMP_RE = re.compile(r'(?<![.\w])(?:On\s+Error\s+GoTo|GoTo|GoSub|Resume)\s+([^\s,:\'"()]+)', re.IGNORECASE)
_JUMP_IGNORE = {'0', '-1', 'next'}


def _diag_missing_labels(lines, procs):
    """VBM012: 飛び先のラベルが無い → [(プロシージャ名, 行番号(1始まり), 飛び先, 行テキスト)]（純 Python）。

    プロシージャ単位（ラベルはプロシージャの中でしか効かない）。大小文字は無視。文字列の中とコメントは見ない。
    """
    out = []
    for p in procs:
        labels = set()
        jumps = []
        for idx in range(p["start"], p["end"] + 1):
            clean = _clean_vba_line(lines[idx])
            if not clean or clean.lower().startswith('rem '):
                continue
            m = _LABEL_RE.match(clean)
            if m and m.group(1).lower() not in _LABEL_KEYWORDS:
                labels.add(m.group(1).lower())
            m = _LINE_NUMBER_RE.match(clean)
            if m:
                labels.add(m.group(1))
            for jm in _JUMP_RE.finditer(clean):
                name = jm.group(1)
                if name.lower() in _JUMP_IGNORE:
                    continue
                jumps.append((idx + 1, name))
        for lineno, name in jumps:
            if name.lower() not in labels:
                out.append((p["name"], lineno, name, lines[lineno - 1].strip()))
    return out


# VBM013（2026-09-17）: `On Err GoTo`／`On Erorr GoTo`（Error の打ち間違い）。コンパイルは通す（On 式 GoTo の形）が
# 実行時に「型が一致しません」等で落ちる。ポスター.xlsm の 名前_Click で実測。
# 飛び先に ',' があるのは計算型の `On 式 GoTo 10, 20`（ErrKind 等の変数）＝打ち間違いではない
_ON_ERR_TYPO_RE = re.compile(r'^\s*On\s+(Er\w*)\s+(?:GoTo|Resume)\b[^,]*$', re.IGNORECASE)


def _diag_on_err_typo(lines):
    """VBM013: On Error の打ち間違い → [(行番号(1始まり), 打った語, 行テキスト)]（純 Python・モジュール単位）。"""
    out = []
    for idx, raw in enumerate(lines):
        clean = _clean_vba_line(raw)
        m = _ON_ERR_TYPO_RE.match(clean)
        if m and m.group(1).lower() != 'error':
            out.append((idx + 1, m.group(1), raw.strip()))
    return out


_RESUME_NEXT_RE = re.compile(r'^\s*On\s+Error\s+Resume\s+Next\b', re.IGNORECASE)


def _diag_resume_next(lines):
    """VBM015: On Error Resume Next の行 → [(行番号(1始まり), 行テキスト)]（純 Python）。
    clean-vba（2026-09-16 夜・Gemini）の目を check に畳んだ（2026-09-17）。"""
    out = []
    for idx, raw in enumerate(lines):
        if _RESUME_NEXT_RE.match(_clean_vba_line(raw)):
            out.append((idx + 1, raw.strip()))
    return out


def _diag_long_procs(procs, limit=150):
    """VBM016: limit 行を超えるプロシージャ → [(名前, 行数)]（純 Python）。"""
    return [(p["name"], p["end"] - p["start"] + 1) for p in procs if p["end"] - p["start"] + 1 > limit]


def _proc_of_index(procs, idx):
    """0 始まりの行番号 idx を含むプロシージャ名（無ければ None）。"""
    for p in procs:
        if p["start"] <= idx <= p["end"]:
            return p["name"]
    return None


def cmd_check(args):
    """VBAコードの静的解析・診断を行う"""
    target_file, _ = parse_target_and_rest(args.posargs)
    # 診断系なので readonly（閉じているブックを通常モードで開くと
    # Workbook_Open / Auto_Open が発火して副作用が走る）
    xl, wb = get_workbook(target_file, readonly=True)
    
    is_json = getattr(args, 'json', False)
    all_warnings = getattr(args, 'all_warnings', False)   # VBM003 の行を全部出す（既定は件数だけ・2026-09-17）

    if not is_json:
        print(f"\n===== VBA診断を実行中: {wb.Name} =====")

    results = {
        "success": True,
        "file": wb.Name,
        "modules": [],
        "duplicates": [],
        "state_leaks": [],
        "unresolved": [],
        "summary": {"errors": 0, "warnings": 0}
    }

    all_procedures = {}  # proc_name -> [module_name, ...]
    book_procs = {}      # (module, proc) -> プロシージャ本体の行
    leaks_by_proc = {}   # (module, proc) -> [(プロパティ, 行), ...]  VBM009 用
    inv_modules = []     # VBM014（存在しないマクロの呼び出し）用の最小の棚卸し（読んだコードを使い回す＝COM を増やさない）

    for comp in wb.VBProject.VBComponents:
        comp_name = comp.Name
        cm = comp.CodeModule
        count_lines = cm.CountOfLines
        
        type_names = {1: '標準モジュール', 2: 'クラスモジュール',
                      3: 'フォーム', 100: 'シート/ThisWorkbook'}
        tname = type_names.get(comp.Type, f'Type={comp.Type}')
        
        mod_info = {
            "name": comp_name,
            "type": tname,
            "type_id": comp.Type,
            "warnings": [],
            "errors": [],
            "skipped": False
        }
        
        if count_lines == 0:
            mod_info["skipped"] = True
            results["modules"].append(mod_info)
            continue
            
        code = cm.Lines(1, count_lines)
        code = code.replace('\r\n', '\n').replace('\r', '\n')
        lines = code.split('\n')
        
        # 1. Option Explicit チェック
        has_option_explicit = False
        for line in lines:
            stripped = line.strip().lower()
            if not stripped:
                continue
            if stripped.startswith("'") or stripped.startswith("rem "):
                continue
            if stripped.startswith("option explicit"):
                has_option_explicit = True
                break
            if re.match(r'^(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property)\s+', stripped) or stripped.startswith("dim ") or stripped.startswith("const "):
                break
        
        if not has_option_explicit:
            mod_info["warnings"].append("Option Explicit が記述されていません。変数宣言の強制を推奨します。")
            
        # 2. Sub/Function 閉じ忘れチェック
        decl_sub = 0
        end_sub = 0
        decl_func = 0
        end_func = 0
        
        proc_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
            r'(Sub|Function)\s+([^\s\(\)]+)',
            re.IGNORECASE
        )
        
        local_variables = []  # (var_name, line_idx)
        variable_usage = {}
        
        # プロシージャごとの詳細診断用状態管理
        current_proc_name = None
        current_proc_start_idx = None
        current_proc_has_error_handler = False
        current_proc_kind = None  # 'sub' or 'function'

        for idx, line in enumerate(lines):
            # 文字列リテラルを潰してからコメントを落とす。先に ' で切ると
            # 文字列内のアポストロフィ（"Don't" 等）で行が切断され、
            # 同じ行の End Sub を見失って正常コードが閉じ忘れ ERROR になる
            clean_line = re.sub(r'"[^"]*"', '""', line).split("'")[0].strip()
            if clean_line.lower().startswith("rem "):
                continue

            m = proc_pattern.match(clean_line)
            if m:
                # 別のプロシージャの中にいる状態で新しい宣言を見つけた場合（前のプロシージャがEnd Subなしで閉じた等）、
                # 簡易クリア（閉じ忘れ警告は後続 of if decl_sub != end_sub で処理）
                kind = m.group(1).lower()
                name = m.group(2)
                current_proc_name = name
                current_proc_start_idx = idx
                current_proc_has_error_handler = False
                current_proc_kind = kind

                if name not in all_procedures:
                    all_procedures[name] = []
                all_procedures[name].append(comp_name)

                if kind == 'sub':
                    decl_sub += 1
                elif kind == 'function':
                    decl_func += 1

            # プロシージャ内における警告チェック
            if current_proc_name:
                # On Error の検出
                if "on error " in clean_line.lower():
                    current_proc_has_error_handler = True
                # SendKeys の検出
                if "sendkeys" in clean_line.lower():
                    mod_info["warnings"].append(f"プロシージャ '{current_proc_name}' 内で危険な SendKeys が使用されています (行 {idx + 1})")

            # 終了チェック。「Sub x(): End Sub」の1行書きは End が行頭に来ないため、
            # validate_vba_code と同じく ':' で文に分割してから数える
            # （文字列内の ':' で誤分割しないよう先に潰す）
            blanked = re.sub(r'"[^"]*"', '""', clean_line)
            segs = [s.strip() for s in blanked.split(':')]
            n_end_sub = sum(1 for s in segs
                            if re.match(r'^end\s+sub\b', s, re.IGNORECASE))
            n_end_func = sum(1 for s in segs
                             if re.match(r'^end\s+function\b', s, re.IGNORECASE))
            is_end_sub = n_end_sub > 0
            is_end_func = n_end_func > 0

            end_sub += n_end_sub
            end_func += n_end_func

            if current_proc_name and (
                (current_proc_kind == 'sub' and is_end_sub) or
                (current_proc_kind == 'function' and is_end_func)
            ):
                if not current_proc_has_error_handler:
                    mod_info["warnings"].append(f"プロシージャ '{current_proc_name}' にエラーハンドリング (On Error) がありません (行 {current_proc_start_idx + 1})")

                current_proc_name = None
                current_proc_start_idx = None
                current_proc_kind = None

            # Dim宣言の簡易スキャン
            dim_match = re.match(r'^\s*Dim\s+(.+)$', clean_line, re.IGNORECASE)
            if dim_match:
                dim_body = dim_match.group(1)
                parts = dim_body.split(',')
                for p in parts:
                    p = p.strip()
                    var_part = re.split(r'\s+As\s+', p, flags=re.IGNORECASE)[0].strip()
                    var_name = re.sub(r'\(.*\)', '', var_part).strip()
                    var_name = re.sub(r'[%&\$#!@]$', '', var_name).strip()
                    if var_name and re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', var_name):
                        local_variables.append((var_name, idx))
                        variable_usage[var_name] = 0

        if decl_sub != end_sub:
            mod_info["errors"].append(f"Sub の閉じ忘れがあります (宣言数: {decl_sub}, End Sub数: {end_sub})")
        if decl_func != end_func:
            mod_info["errors"].append(f"Function の閉じ忘れがあります (宣言数: {decl_func}, End Function数: {end_func})")

        # 未使用変数のカウントスキャン
        if local_variables:
            for idx, line in enumerate(lines):
                # 上と同じ理由で、文字列を潰してからコメントを落とす
                clean_line = re.sub(r'"[^"]*"', '""', line).split("'")[0]
                for var_name, decl_idx in local_variables:
                    if idx == decl_idx:
                        continue
                    if re.search(r'\b' + re.escape(var_name) + r'\b', clean_line, re.IGNORECASE):
                        variable_usage[var_name] += 1

            for var_name, decl_idx in local_variables:
                if variable_usage[var_name] == 0:
                    mod_info["warnings"].append(f"未使用変数 Dim {var_name} があります (行 {decl_idx + 1})")

        # ---- 追加診断4本（走らせても出ない欠陥） ----
        procs = _split_procedures(lines)
        for p in procs:
            book_procs[(comp_name, p["name"])] = lines[p["start"]:p["end"] + 1]

        # VBM008: Application の状態を戻していない
        leaks = _diag_app_state(lines, procs)
        for pname, props in leaks.items():
            leaks_by_proc[(comp_name, pname)] = props
            for prop, lineno in props:
                mod_info["warnings"].append(
                    f"プロシージャ '{pname}' が Application.{prop} を変更したまま戻していません (行 {lineno})")

        # VBM011: Find/Replace の引数省略
        for pname, kind, lineno, missing in _diag_stateful_find(lines, procs):
            mod_info["warnings"].append(
                f"プロシージャ '{pname}' の .{kind} が {'/'.join(missing)} を省略しています"
                f"（Excel が前回の設定を引き継ぎます）(行 {lineno})")

        # VBM010: モジュール変数の書き手が散っている
        for vname, lineno, writers in _diag_module_state(lines, procs):
            mod_info["warnings"].append(
                f"モジュール変数 {vname} の書き手が {len(writers)} か所に散っています"
                f"（{'、'.join(writers)}）(行 {lineno})")

        # VBM012: 飛び先のラベルが無い（コンパイルで止まる欠陥を check で先に拾う・2026-09-17）
        for pname, lineno, target, text in _diag_missing_labels(lines, procs):
            mod_info["errors"].append(
                f"プロシージャ '{pname}' の飛び先ラベル '{target}' がありません (行 {lineno}): {text}")

        # VBM013: On Err GoTo（On Error の打ち間違い）
        for lineno, word, text in _diag_on_err_typo(lines):
            pname = _proc_of_index(procs, lineno - 1)
            mod_info["errors"].append(
                f"プロシージャ '{pname}' の 'On {word}' は On Error の打ち間違いです (行 {lineno}): {text}")

        # VBM015: On Error Resume Next（握りつぶし。行を出す）／VBM016: 150 行超（2026-09-17・clean-vba から）
        for lineno, text in _diag_resume_next(lines):
            pname = _proc_of_index(procs, lineno - 1)
            mod_info["warnings"].append(
                f"プロシージャ '{pname}' の On Error Resume Next (行 {lineno}): {text}")
        for pname, n in _diag_long_procs(procs):
            mod_info["warnings"].append(f"プロシージャ '{pname}' が {n} 行あります（150 行超）")

        inv_modules.append({'name': comp_name, 'type': int(comp.Type),
                            'procs': [{'name': p["name"]} for p in procs], 'code': '\r\n'.join(lines)})
        results["modules"].append(mod_info)

        # 画面出力 (JSON指定でない場合のみ)
        if not is_json:
            print(f"\n📄 モジュール: {comp_name} ({tname})")
            if not has_option_explicit:
                print("  [WARNING] Option Explicit が記述されていません。変数宣言の強制を推奨します。")
            if mod_info["errors"]:
                for err in mod_info["errors"]:
                    print(f"  [ERROR] {err}")
            n_onerr = 0
            if mod_info["warnings"]:
                for warn in mod_info["warnings"]:
                    # Option Explicitの警告はすでに出力しているのでスキップ
                    if "Option Explicit" in warn:
                        continue
                    # VBM003（On Error が無い）は既定で件数だけ（ポスターで 190 件中 150 件がこれ＝本題が埋もれる・2026-09-17）
                    if not all_warnings and "エラーハンドリング (On Error) がありません" in warn:
                        n_onerr += 1
                        continue
                    print(f"  [WARNING] {warn}")
            if n_onerr:
                print(f"  [WARNING] On Error の無いプロシージャ {n_onerr} 本（行は --all-warnings で表示）")
            print("  モジュール診断完了")

    # VBM014: 存在しないマクロを呼んでいる（call-graph・checkup と同じ純 Python の解析。Declare 済み API は除外済み・2026-09-17）
    unresolved = []
    if inv_modules:
        try:
            # Application.Run "名前" は外して Call／裸呼びだけ（Run はアドインや別ブックのマクロを実行時に探す＝
            # コンパイルエラーにならない。error にすると gate がアドイン呼びのブックを全部落とす）
            unresolved = [u for u in _analyze_calls({'modules': inv_modules})['unresolved']
                          if not str(u[2]).startswith('Run "')]
        except Exception as ex:
            print(f"⚠ 呼び出し先の実在確認ができませんでした: {ex}", file=sys.stderr)
    by_name = {m["name"]: m for m in results["modules"]}
    for mod, proc, name, lineno, text in unresolved:
        results["unresolved"].append({"module": mod, "proc": proc, "name": name, "line": lineno, "text": text})
        if mod in by_name:
            by_name[mod]["errors"].append(
                f"プロシージャ '{proc or '(宣言部)'}' が存在しないマクロ '{name}' を呼んでいます (行 {lineno}): {text}")

    # VBM009: 状態を戻さないヘルパーを呼んでいる箇所（手続き間・ブック全体で1回）
    results["state_leaks"] = _diag_state_callers(book_procs, leaks_by_proc)

    # 重複チェックの集計
    for proc_name, mods in all_procedures.items():
        if len(mods) > 1:
            results["duplicates"].append({
                "procedure": proc_name,
                "modules": mods
            })

    # サマリーの集計
    err_total = sum(len(m["errors"]) for m in results["modules"])
    warn_total = (sum(len(m["warnings"]) for m in results["modules"])
                  + len(results["duplicates"]) + len(results["state_leaks"]))
    results["summary"]["errors"] = err_total
    results["summary"]["warnings"] = warn_total

    # JSON出力
    if is_json:
        import json
        print(json.dumps(results, ensure_ascii=False), file=sys.stdout)
        return err_total == 0

    # 通常出力
    print("\n===== 存在しないマクロの呼び出し（VBM014） =====")
    if results["unresolved"]:
        for u in results["unresolved"]:
            print(f"  [ERROR] [{u['module']}] {u['proc'] or '(宣言部)'} :{u['line']}  →  {u['name']}")
            print(f"      {u['text']}")
    else:
        print("  Call／Run の呼び先はすべて実在のマクロです。")

    print("\n===== ブック全体の重複診断 =====")
    if results["duplicates"]:
        for dup in results["duplicates"]:
            print(f"  [WARNING] 重複プロシージャ名 '{dup['procedure']}' が複数のモジュールに存在します:")
            for m in dup["modules"]:
                print(f"    - {m}")
    else:
        print("  プロシージャ名の重複はありません。")

    print("\n===== 呼び出し先の状態戻し忘れ（手続き間） =====")
    if results["state_leaks"]:
        for s in results["state_leaks"]:
            print(f"  [WARNING] {s['module']}.{s['caller']} が呼ぶ '{s['callee']}' は "
                  f"{'／'.join(s['props'])} を変更したまま返ります")
    else:
        print("  状態を変えたまま返るヘルパーの呼び出しはありません。")

    print(f"\n===== 診断サマリー =====")
    print(f"  エラー数  : {err_total}")
    print(f"  警告数    : {warn_total}")
    
    if err_total > 0:
        print("  [RESULT] 致命的な構文エラーがあります。修正してください。")
        return False
    elif warn_total > 0:
        print("  [RESULT] 警告項目がありますが、実行は可能です。品質向上のため修正を推奨します。")
        return True
    else:
        print("  [RESULT] すべてのチェックを通過しました。良好な状態です。")
        return True


# ================================================================
# xlflow から輸入したコマンド（2026-08-28）
#   inspect-gui   自動操縦を止める GUI 境界を、撃つ前に洗い出す
#   capabilities  各コマンドが破壊的かどうかを機械可読で出す
#   rules         check の規則一覧（Excel を開かない）
#   metrics       プロシージャ計量とホットスポット順位
#   process       Excel プロセスの点呼と後始末
# ================================================================

# 自動操縦を止める境界＝人の手が要る場所。ここで rehearse も gate も黙って固まる
_GUI_BOUNDARIES = [
    ("MsgBox", re.compile(r'(?<![.\w])MsgBox\b', re.IGNORECASE),
     "応答があるまで止まる"),
    ("InputBox", re.compile(r'(?<![.\w])InputBox\b|\bApplication\s*\.\s*InputBox\b', re.IGNORECASE),
     "入力があるまで止まる"),
    ("Application.FileDialog", re.compile(r'\bApplication\s*\.\s*FileDialog\b', re.IGNORECASE),
     "選択があるまで止まる"),
    ("GetOpenFilename", re.compile(r'\bGetOpenFilename\b', re.IGNORECASE),
     "選択があるまで止まる"),
    ("GetSaveAsFilename", re.compile(r'\bGetSaveAsFilename\b', re.IGNORECASE),
     "選択があるまで止まる"),
    ("Application.Dialogs", re.compile(r'\bApplication\s*\.\s*Dialogs\s*\(', re.IGNORECASE),
     "組み込みダイアログで止まる"),
    ("モーダルの .Show", re.compile(r'\.\s*Show\b(?!.*vbModeless)', re.IGNORECASE),
     "フォームを閉じるまで止まる"),
    ("SendKeys", re.compile(r'(?<![.\w])SendKeys\b|\bApplication\s*\.\s*SendKeys\b', re.IGNORECASE),
     "送り先が変わると別の窓へ飛ぶ"),
    ("Application.Wait", re.compile(r'\bApplication\s*\.\s*Wait\b', re.IGNORECASE),
     "指定時刻まで固まる"),
]


def _scan_gui_boundaries(lines, procs):
    """GUI 境界を [{proc, line, kind, why, code}] で返す"""
    out = []
    for p in procs:
        for idx in range(p["start"], p["end"] + 1):
            clean = _clean_vba_line(lines[idx])
            if not clean:
                continue
            for label, rx, why in _GUI_BOUNDARIES:
                if rx.search(clean):
                    out.append({"proc": p["name"], "line": idx + 1, "kind": label,
                                "why": why, "code": lines[idx].strip()})
                    break
    return out


def _reachable_procs(bodies, start_name):
    """start_name から同ブック内で到達できるプロシージャ名の集合。
    名前を1本の選択肢正規表現に畳んでから舐める（1本ずつ回すと行数×本数で潰れる）"""
    names = sorted(bodies.keys(), key=len, reverse=True)
    if not names:
        return set()
    lower = {n.lower(): n for n in names}
    big = re.compile(r'(?<![\w.])(' + '|'.join(re.escape(n) for n in names) + r')(?![\w])',
                     re.IGNORECASE)
    start = lower.get(start_name.lower())
    if not start:
        return set()
    seen, todo = {start}, [start]
    while todo:
        cur = todo.pop()
        for raw in bodies.get(cur, []):
            for m in big.finditer(_clean_vba_line(raw)):
                cand = lower.get(m.group(1).lower())
                if cand and cand not in seen:
                    seen.add(cand)
                    todo.append(cand)
    return seen


def cmd_inspect_gui(args):
    """自動操縦を止める GUI 境界を走査する（撃つ前に、固まる場所が分かる）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)
    is_json = getattr(args, 'json', False)
    only_module = getattr(args, 'module_opt', None)
    macro = rest[0] if rest else None

    hits, bodies, per_module = [], {}, {}
    for comp in wb.VBProject.VBComponents:
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n == 0:
            continue
        lines = cm.Lines(1, n).replace('\r\n', '\n').replace('\r', '\n').split('\n')
        procs = _split_procedures(lines)
        for p in procs:
            bodies.setdefault(p["name"], []).extend(lines[p["start"]:p["end"] + 1])
        per_module[comp.Name] = (lines, procs)

    for mod_name, (lines, procs) in per_module.items():
        if only_module and mod_name.lower() != only_module.lower():
            continue
        for h in _scan_gui_boundaries(lines, procs):
            h["module"] = mod_name
            hits.append(h)

    scope_note = ""
    if macro:
        reach = _reachable_procs(bodies, macro)
        if not reach:
            print(f"マクロ '{macro}' が見つかりません")
            return False
        hits = [h for h in hits if h["proc"] in reach]
        scope_note = f"（'{macro}' から到達する {len(reach)} 本に限定）"

    if is_json:
        import json
        print(json.dumps({"success": True, "file": wb.Name, "macro": macro,
                          "boundaries": hits}, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"\n===== GUI 境界の走査: {wb.Name} {scope_note} =====")
    if not hits:
        print("  自動操縦を止める境界はありません。そのまま撃てます。")
        return True
    cur = None
    for h in hits:
        key = (h["module"], h["proc"])
        if key != cur:
            cur = key
            print(f"\n📄 {h['module']}.{h['proc']}")
        print(f"  行 {h['line']:>5}  [{h['kind']}] {h['why']}")
        print(f"         {h['code']}")
    print(f"\n  合計 {len(hits)} 件。ここでマクロは人の手を待ちます。")
    return True


# 各コマンドの安全性メタデータ。撃つ前にここを見れば、何が起きるか分かる
#   read        ブックを読むだけ（書き換えない）
#   write       書き換える（バックアップまたは元に戻す手がある）
#   destructive 消す・中断する（非可逆、または走っているものを止める）
#   execute     任意の VBA を走らせる（何が起きるかは静的に読めない）
_CAPABILITIES = {
    "read": [
        ("diag", "動作確認"), ("setup-check", "導入セルフ診断"),
        ("list-open", "開いているブックの点呼"), ("list", "マクロ一覧"),
        ("list-modules", "モジュール一覧"), ("list-forms", "フォーム一覧"),
        ("get", "プロシージャ取得"), ("docs", "取説の自動生成"),
        ("checkup", "健康診断"), ("check", "静的診断"), ("check-bas", ".bas 単体検査"),
        ("compile", "全体コンパイル（押すだけ・ブックは変わらない）"),
        ("call-graph", "呼び出し関係"), ("impact", "影響範囲"),
        ("grep", "コード横断検索"), ("metrics", "プロシージャ計量"),
        ("rules", "診断規則の一覧"), ("capabilities", "この表そのもの"),
        ("audit", "表の仕上げ検査（見出し・罫線・列の型・列幅・空の見出し。読むだけ）"),
        ("inspect-gui", "GUI 境界の走査"), ("wiring", "ボタン配線図"),
        ("list-backups", "バックアップ一覧"), ("list-shortcuts", "ショートカット一覧"),
        ("stats", "呼び出し台帳の集計（コマンド別の回数・秒・失敗・呼び出しの間。COM不要）"),
        ("status", "いま走っているコマンドと進み具合（COM不要）"),
        ("diff-module", "モジュールの差分（控え・.bas・別ブック）を読むだけ"),
        ("repair", "マクロ修理の材料（本文・入口・呼び元呼び先・コンパイル・check・控えとの差分）を 1 手で。読むだけ"),
        ("versions", "同名ブックの写しを日時順に（oletools で読む・Excel を開かない・COM不要）"),
        ("history", "マクロの本文の歴史（控え・_exports・写しの .xlsm から。COM不要）"),
        ("list-file", "閉じたブックのマクロ一覧（oletools・Excel を開かない）"),
        ("grep-files", "閉じたブック横断のコード検索（oletools・Excel を開かない）"),
        ("export-file", "閉じたブックのモジュールを .bas に書き出す（ファイルは作るがブックは無傷）"),
        ("read-range", "セル読み"), ("read-selection", "選択範囲読み"),
        ("sheet-info", "シート情報"), ("materials", "先回り材料（1シートの現物を1回で見る）"),
        ("snapshot", "意味構造JSONに畳む"),
        ("snapshot-diff", "snapshot の差分"), ("screenshot", "画面撮り"),
        ("find", "検索"), ("printer-list", "プリンター一覧"),
        ("export-module", "モジュール書き出し（ファイルは作るがブックは無傷）"),
        ("export-all", "全モジュール書き出し（同上）"),
        ("form-to-vba", "フォームを作成マクロに変換（同上）"),
        ("process", "Excel プロセスの点呼"),
        ("diagnose", "表・シートの数式＆データ総合診断（集計漏れ・定数直書き・隠れ空白・外れ値検知）"),
        ("style-map", "表・シートの視覚レイアウト＆書式マップ（背景色・フォント・二重罫線・複合見出しツリー）"),
        ("shelf", "棚（表の整理）の目録（モジュールのコードを読むだけ。--grep で絞る）"),
        ("trace", "式の元をシートをまたいでたどる（番地・式・値の木。式の字面を読むだけ）"),
    ],
    "write": [
        ("replace-procedure", "プロシージャ置換（バックアップあり）"),
        ("patch-procedure", "プロシージャ差分置換（ピンポイント置換・バックアップあり）"),
        ("add-procedure", "プロシージャ追加"), ("add-module", "モジュール追加"),
        ("replace-module", "モジュール置換（バックアップあり）"),
        ("code-replace", "横断置換（diff 確認＋バックアップ）"),
        ("reorder-macro", "並べ替え"), ("restore", "バックアップから書き戻し"),
        ("rename-procedure", "改名（宣言・呼び元・OnAction を変更行だけ書き換え。控えつき・保存する）"),
        ("copy-modules", "別ブックへ複製（同名は --overwrite のときだけ控えを取って置換）"),
        ("references", "参照設定の追加・削除（list は読むだけ。控えつき・保存する）"),
        ("set-shortcut", "ショートカットの付け替え（MacroOptions。控えつき・保存する）"),
        ("format-module", "整形（既定は差分だけ。--apply で replace-module 経路）"),
        ("backup-prune", "backups の間引き（既定は数えるだけ。--force で消す。ブックには触らない）"),
        ("write-range", "セル書き込み"), ("write-cells", "飛び飛びのセル書き込み（1往復）"),
        ("format-range", "書式"),
        ("tidy", "表の書き方と罫線と列幅をそろえる（見出し・罫線・番号列左寄せ・数値カンマ・列幅→見え方の読み戻し）"),
        ("clean-table", "表を掃除する（AI なし・数秒。空白の全角半角・半角カナ・全角英数・文字の日付→規則で直し、"
                        "重複行は --delete-dups のときだけ削除、空欄は埋めず報告、最後に tidy）"),
        ("表の掃除", "clean-table と同じ"),
        ("seiri", "表を直す 1 手目（materials の代わり）：「表の書き方と罫線と列幅をそろえる」マクロを撃ち、残り（エラーセルと原因・気づき・指示文・###）"
                  "だけを出す（2026-09-13）"),
        ("表の整理", "seiri と同じ"),
        ("build-sheet", "設計図（JSON）からシートを一枚組み上げる（新規シート／--overwrite で既存を消して組み直す）"),
        ("agent", "依頼文をシート 1 枚に対して回す（AI が JSON で手を返し、道具が書いて読み戻す。保存はしない）"),
        ("set-key", "APIキーを預かる（ユーザー環境変数へ書く。ブックには触らない・ファイルにも書かない）"),
        ("sheet", "シート操作"), ("table", "テーブル"), ("name", "名前定義"),
        ("row", "行"), ("col", "列"), ("copy-range", "複写"), ("fill", "連続入力"),
        ("sort", "並べ替え"), ("autofilter", "フィルター"),
        ("find-replace", "検索と置換"), ("cond-format", "条件付き書式"),
        ("hyperlink", "ハイパーリンク"), ("validation", "入力規則"),
        ("freeze", "ウィンドウ枠固定"), ("comment", "コメント"),
        ("chart", "グラフ"), ("chart-config", "グラフ設定"),
        ("pivot", "ピボット"), ("pivot-field", "ピボットフィールド"),
        ("pivot-calc", "ピボット計算"), ("slicer", "スライサー"),
        ("calc-mode", "計算モード"), ("powerquery", "パワークエリ"),
        ("connection", "接続"), ("datamodel", "データモデル"),
        ("print-setup", "印刷設定"), ("printer-setup", "プリンター設定"),
        ("save", "保存"), ("save-as", "名前を付けて保存"),
        ("export-pdf", "PDF 書き出し"), ("open", "ブックを開く"),
    ],
    "destructive": [
        ("delete-procedure", "プロシージャ削除（確認あり）"),
        ("delete-module", "モジュール削除（要約→確認→バックアップ）"),
        ("clear-range", "セル消去"),
        ("shape", "図形の一覧・削除・位置と大きさ（--delete は消す。マクロ付きのボタンは拒む）"),
        ("close", "ブックを閉じる（保存方針しだいで編集が消える）"),
        ("close-form", "表示中のフォームを閉じる（走っている処理を打ち切る）"),
        ("vbe-reset", "実行>リセット（走っている VBA を全部止める。人の指示でだけ）"),
        ("clear-key", "APIキーを消す（ユーザー環境変数から。発行元のキーは生きているので要るなら失効させる）"),
    ],
    "execute": [
        ("run-macro", "マクロを本体で実行（何が起きるかは静的に読めない）"),
        ("rehearse", "コピーに試し撃ち（本体は無傷。ただし任意コードは走る）"),
        ("test", "テスト Sub の一括実行"),
        ("gate", "関所（コピーで構文検査＋テスト実行）"),
        ("batch", "コマンド列の連続実行（中身しだい）"),
        ("shell", "対話セッション（中身しだい）"),
        ("shelf-run", "棚（表の整理）のマクロを選んで撃つ（何が起きるかは静的に読めない。撃つ前に控えを取り差分を返す）"),
        ("register-addin", "前に出ているブックを .xlam に焼き直してアドイン登録（「アドインの更新登録」を撃つ＝xlam を作り直す）"),
        ("更新登録", "register-addin と同じ"),
    ],
}

_CAP_MARK = {"read": "読むだけ", "write": "書き換える",
             "destructive": "消す・止める", "execute": "任意コードを走らせる"}


def cmd_capabilities(args):
    """各コマンドが破壊的かどうかを機械可読で出す（Excel に触らない）"""
    is_json = getattr(args, 'json', False)
    _, rest = parse_target_and_rest(args.posargs)
    want = rest[0] if rest else None

    if is_json:
        import json
        data = {"success": True, "commands": [
            {"name": n, "safety": cat, "note": note}
            for cat, items in _CAPABILITIES.items() for n, note in items
            if not want or n == want]}
        print(json.dumps(data, ensure_ascii=False), file=sys.stdout)
        return True

    if want:
        for cat, items in _CAPABILITIES.items():
            for n, note in items:
                if n == want:
                    print(f"{n} : {_CAP_MARK[cat]}（{cat}）")
                    print(f"  {note}")
                    return True
        print(f"'{want}' は一覧にありません")
        return False

    print("\n===== コマンドの安全性 =====")
    for cat in ("read", "write", "destructive", "execute"):
        items = _CAPABILITIES[cat]
        print(f"\n[{cat}] {_CAP_MARK[cat]}  ({len(items)}件)")
        for n, note in items:
            print(f"  {n:20s} {note}")
    total = sum(len(v) for v in _CAPABILITIES.values())
    print(f"\n  合計 {total} コマンド。"
          "destructive と execute は、撃つ前に対象を確かめること。")
    return True


# check が見ている規則の一覧。VBM008〜011 は xlflow の規則から輸入（2026-08-28）
_CHECK_RULES = [
    ("VBM001", "warning", "Option Explicit が無い", "モジュール", ""),
    ("VBM002", "warning", "SendKeys の使用", "手続き内", ""),
    ("VBM003", "warning", "On Error が無い", "手続き内", ""),
    ("VBM004", "error", "Sub の閉じ忘れ", "モジュール", ""),
    ("VBM005", "error", "Function の閉じ忘れ", "モジュール", ""),
    ("VBM006", "warning", "未使用の Dim", "手続き内", ""),
    ("VBM007", "warning", "プロシージャ名の重複", "ブック全体", ""),
    ("VBM008", "warning", "Application の状態を戻していない", "手続き内", "xlflow VBA203"),
    ("VBM009", "warning", "呼んだヘルパーが状態を変えたまま返る", "手続き間", "xlflow VBA221"),
    ("VBM010", "warning", "モジュール変数の書き手が散っている", "モジュール", "xlflow VBA240"),
    ("VBM011", "warning", "Find/Replace の引数省略", "手続き内", "xlflow VBA215"),
    # 2026-09-17: コンパイルで止まる・実行時に落ちる欠陥を check で先に拾う（ポスター.xlsm の 3 件が check を素通りしていた）
    ("VBM012", "error", "飛び先のラベルが無い（On Error GoTo／GoTo／GoSub／Resume）", "手続き内", ""),
    ("VBM013", "error", "On Err GoTo（On Error の打ち間違い）", "手続き内", ""),
    ("VBM014", "error", "存在しないマクロを呼んでいる（call-graph の未解決と同じ）", "ブック全体", ""),
    # 2026-09-17: clean-vba（9/16 夜）の目を check に畳んだ。戻し忘れは VBM008 が既に見る
    ("VBM015", "warning", "On Error Resume Next（エラーを握りつぶす行）", "手続き内", ""),
    ("VBM016", "warning", "150 行を超えるプロシージャ", "手続き内", ""),
]


def cmd_rules(args):
    """check の規則一覧を出す（Excel を開かない）"""
    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "rules": [
            {"code": c, "severity": s, "name": n, "scope": sc, "origin": o}
            for c, s, n, sc, o in _CHECK_RULES]}, ensure_ascii=False), file=sys.stdout)
        return True
    print("\n===== check の診断規則 =====")
    print(f"  {'コード':8s} {'重さ':8s} {'範囲':10s} 内容")
    for c, s, n, sc, o in _CHECK_RULES:
        tail = f"  ← {o}" if o else ""
        print(f"  {c:8s} {s:8s} {sc:10s} {n}{tail}")
    print("\n  VBM008〜011 は xlflow の117本から輸入した4本。")
    print("  選んだ基準は「走らせても出ない」欠陥だけ——走れば分かるものは入れていない。")
    print("  VBM012〜014（2026-09-17）はコンパイル・実行で止まる欠陥を check で先に拾う。VBM003 は既定で件数だけ（--all-warnings で行）。")
    return True


def cmd_metrics(args):
    """プロシージャ計量とホットスポット順位（どこが重いか）"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)
    is_json = getattr(args, 'json', False)
    top = getattr(args, 'top', None) or 20

    branch_re = re.compile(
        r'^\s*(If\b|ElseIf\b|Select\s+Case\b|Case\b|For\b|Do\b|While\b|Loop\s+(While|Until)\b)',
        re.IGNORECASE)
    open_re = re.compile(r'^\s*(If\b.*\bThen\s*$|For\b|Do\b|While\b|With\b|Select\s+Case\b)',
                         re.IGNORECASE)
    close_re = re.compile(r'^\s*(End\s+If\b|Next\b|Loop\b|Wend\b|End\s+With\b|End\s+Select\b)',
                          re.IGNORECASE)

    rows = []
    for comp in wb.VBProject.VBComponents:
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n == 0:
            continue
        lines = cm.Lines(1, n).replace('\r\n', '\n').replace('\r', '\n').split('\n')
        for p in _split_procedures(lines):
            body = lines[p["start"]:p["end"] + 1]
            code_lines = branches = 0
            depth = maxdepth = 0
            for raw in body:
                clean = _clean_vba_line(raw)
                if not clean:
                    continue
                code_lines += 1
                if branch_re.match(clean):
                    branches += 1
                if close_re.match(clean):
                    depth = max(0, depth - 1)
                if open_re.match(clean):
                    depth += 1
                    maxdepth = max(maxdepth, depth)
            rows.append({"module": comp.Name, "proc": p["name"], "lines": code_lines,
                         "branches": branches, "depth": maxdepth,
                         "score": code_lines + branches * 5 + maxdepth * 10})
    rows.sort(key=lambda r: r["score"], reverse=True)

    if is_json:
        import json
        print(json.dumps({"success": True, "file": wb.Name, "procedures": rows},
                         ensure_ascii=False), file=sys.stdout)
        return True

    print(f"\n===== ホットスポット順位: {wb.Name} =====")
    print(f"  {'順':>3} {'行':>5} {'分岐':>5} {'深さ':>4}  モジュール.プロシージャ")
    for i, r in enumerate(rows[:top], 1):
        print(f"  {i:>3} {r['lines']:>5} {r['branches']:>5} {r['depth']:>4}  "
              f"{r['module']}.{r['proc']}")
    tot_l = sum(r["lines"] for r in rows)
    print(f"\n  プロシージャ {len(rows)} 本 / コード行 {tot_l} 行"
          f"（上位 {min(top, len(rows))} 本を表示。--top で件数を変えられます）")
    return True


def cmd_process(args):
    """Excel プロセスの点呼。既定は一覧だけ——kill は PID 名指しのときにしかしない"""
    import subprocess
    is_json = getattr(args, 'json', False)
    kill_pid = getattr(args, 'kill', None)

    # tasklist で EXCEL.EXE を全部拾う（ROT に出ないゾンビもここには出る）
    pids = {}
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE",
                              "/FO", "CSV", "/NH"],
                             capture_output=True, text=True,
                             encoding='cp932', errors='replace')
        for line in out.stdout.splitlines():
            parts = [x.strip().strip('"') for x in line.split('","')]
            if len(parts) >= 5 and parts[0].upper().startswith("EXCEL"):
                try:
                    pids[int(parts[1])] = {"pid": int(parts[1]), "mem": parts[4],
                                           "books": [], "visible": None}
                except ValueError:
                    pass
    except Exception as e:
        print(f"tasklist の実行に失敗しました: {e}")

    # ROT 側から、どのプロセスがどのブックを抱えているかを埋める
    try:
        import win32process
        import win32gui  # noqa: F401  (Hwnd→PID に使う)
        for wbk in _running_excel_workbooks():
            try:
                app = wbk.Application
                hwnd = int(app.Hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                info = pids.setdefault(pid, {"pid": pid, "mem": "?",
                                             "books": [], "visible": None})
                info["visible"] = bool(app.Visible)
                info["books"].append(wbk.Name)
            except Exception:
                continue
    except Exception:
        pass

    rows = sorted(pids.values(), key=lambda r: r["pid"])

    if kill_pid is not None:
        target = pids.get(int(kill_pid))
        if target is None:
            print(f"PID {kill_pid} の EXCEL.EXE は見つかりません")
            return False
        if target["books"]:
            print(f"PID {kill_pid} はブックを {len(target['books'])} 冊抱えています: "
                  f"{'、'.join(target['books'])}")
            print("  作業中の Excel を落とす恐れがあるため、kill しません。"
                  "先に close で閉じてください。")
            return False
        subprocess.run(["taskkill", "/PID", str(kill_pid), "/F"],
                       capture_output=True, text=True)
        print(f"PID {kill_pid} を終了しました（ブック0冊のプロセス）")
        return True

    if is_json:
        import json
        print(json.dumps({"success": True, "processes": rows},
                         ensure_ascii=False), file=sys.stdout)
        return True

    print("\n===== Excel プロセスの点呼 =====")
    if not rows:
        print("  EXCEL.EXE は動いていません。")
        return True
    for r in rows:
        vis = {True: "表示", False: "非表示", None: "?"}[r["visible"]]
        books = "、".join(r["books"]) if r["books"] else "（ブック0冊）"
        print(f"  PID {r['pid']:>6}  {vis:6s}  {r['mem']:>12s}  {books}")
    ghosts = [r for r in rows if not r["books"]]
    if ghosts:
        print(f"\n  ブック0冊のプロセスが {len(ghosts)} 個あります"
              f"（PID {'、'.join(str(g['pid']) for g in ghosts)}）。")
        print("  掃除するなら PID を名指しで: process --kill <PID>")
        print("  ※ 名指し以外では落としません（作業中の Excel を巻き込まないため）")
    return True


def cmd_diag(args):
    """動作確認"""
    print("Syntax OK")
    try:
        xl = _get_active_excel()
        wb = xl.ActiveWorkbook
        if wb:
            print(f"アクティブブック: {wb.Name}")
        else:
            print("アクティブブック: なし")
    except Exception:
        print("Excelは起動していません")
    return True

def cmd_list_open(args):
    """現在開いているExcelファイルを一覧表示"""
    as_json = getattr(args, 'json', False)
    books = []
    seen = set()

    def _add(wb):
        try:
            full = wb.FullName
            name = wb.Name
        except Exception:
            return
        if full.lower() in seen:
            return
        seen.add(full.lower())
        books.append({"name": name, "fullname": full})

    # ROT 全走査: GetActiveObject は ROT 先頭の1インスタンスしか返さず、
    # 非表示ゾンビ(ブック0個)を掴んで「ブックなし」と誤報することがある。
    # 点呼コマンドこそ全インスタンスの全ブックを数える必要がある
    for wb in _running_excel_workbooks():
        _add(wb)
    excel_running = bool(books)
    # 未保存ブック(Book1等)はパスを持たず ROT に載らないことがあるため、
    # GetActiveObject 側の列挙でも補完する
    try:
        xl = _get_active_excel()
        excel_running = True
        for wb in xl.Workbooks:
            _add(wb)
    except Exception:
        pass

    if as_json:
        import json
        print(json.dumps({"success": True, "excel_running": excel_running,
                          "workbooks": books}, ensure_ascii=False), file=sys.stdout)
        return True
    if not excel_running:
        print('Excelは起動していません')
        # 「起動していません」は正常な状態報告であって異常ではない。
        # ここだけ None を返すと終了コードが 1 になり他の分岐と不揃いになる
        return True
    if not books:
        print('（開いているブックはありません）')
        return True
    for b in books:
        print(b["fullname"])
    return True



def _select_addin_project(all_projects, sel):
    """--addin の対象アドインを選ぶ。sel は True（無指定）または名前の一部。

    特定のアドイン名を優先するハードコードはしない（汎用原則）。
    複数ロード時は名前指定を促す。見つからなければ None（メッセージ出力済み）。
    """
    found = []
    for p in all_projects:
        try:
            fname = os.path.basename(p.Filename).lower()
            if fname.endswith(('.xlam', '.xla')):
                found.append(p)
        except Exception:
            continue
    if not found:
        print("エラー: アドインブック (.xlam / .xla) がロードされていません。", file=sys.stderr)
        return None
    if isinstance(sel, str):
        hits = [p for p in found
                if sel.lower() in os.path.basename(p.Filename).lower()]
        if not hits:
            names = ", ".join(os.path.basename(p.Filename) for p in found)
            print(f"エラー: '{sel}' に一致するアドインがありません。ロード中: {names}",
                  file=sys.stderr)
            return None
        if len(hits) > 1:
            names = ", ".join(os.path.basename(p.Filename) for p in hits)
            print(f"エラー: '{sel}' に複数一致します: {names}", file=sys.stderr)
            return None
        return hits[0]
    if len(found) > 1:
        names = ", ".join(os.path.basename(p.Filename) for p in found)
        print("エラー: 複数のアドインがロードされています。"
              "--addin 名前 で対象を指定してください。", file=sys.stderr)
        print(f"  ロード中: {names}", file=sys.stderr)
        return None
    return found[0]


def cmd_list(args):
    """マクロ(プロシージャ)一覧"""
    load_addins = getattr(args, 'personal', False) or getattr(args, 'addin', False) or getattr(args, 'all', False)
    target_file, _ = parse_target_and_rest(args.posargs)
    # 参照するだけのコマンド。閉じたブックを自動で開く場合に Workbook_Open /
    # Auto_Open を発火させないよう読み取り専用＋イベント無効で開く
    # （既に開いているブックには影響しない）
    xl, default_wb = get_workbook(target_file, load_addins=load_addins, readonly=True)

    # 全プロジェクトをリスト化 (ビジーエラー対策としてリトライ)
    import time
    all_projects = []
    for attempt in range(5):
        try:
            all_projects = []
            for p in xl.VBE.VBProjects:
                all_projects.append(p)
            break
        except Exception as ex:
            if "800ac472" in str(ex) and attempt < 4:
                time.sleep(0.5)
                continue
            print(f"エラー: VBAプロジェクトモデルへのアクセスが拒否されました: {ex}", file=sys.stderr)
            return False

    # 対象のプロジェクトを選択
    target_projects = []
    if getattr(args, 'all', False):
        target_projects = all_projects
    elif getattr(args, 'personal', False):
        found = None
        for p in all_projects:
            try:
                fname = os.path.basename(p.Filename).lower()
                if fname in ("personal.xlsb", "personal.xls"):
                    found = p
                    break
            except Exception:
                continue
        if not found:
            print("エラー: 個人用マクロブック (PERSONAL.XLSB) がロードされていません。", file=sys.stderr)
            return False
        target_projects.append(found)
    elif getattr(args, 'addin', False):
        target_addin = _select_addin_project(all_projects, args.addin)
        if target_addin is None:
            return False
        target_projects.append(target_addin)
    else:
        found = None
        try:
            for p in all_projects:
                try:
                    if p.Filename.lower() == default_wb.FullName.lower():
                        found = p
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not found:
            try:
                found = default_wb.VBProject
            except Exception as ex:
                print(f"エラー: VBProjectの取得に失敗しました: {ex}", file=sys.stderr)
                return False
        target_projects.append(found)

    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE
    )

    def get_project_display_name(p):
        try:
            if p.Filename:
                return os.path.basename(p.Filename)
        except Exception:
            pass
        return p.Name

    is_all = getattr(args, 'all', False)
    if is_all:
        all_results = {}
        for proj in target_projects:
            macros = []
            proj_name = get_project_display_name(proj)
            try:
                for comp in proj.VBComponents:
                    if getattr(args, 'standard', False) and comp.Type != 1:
                        continue
                    cm = comp.CodeModule
                    if cm.CountOfLines == 0:
                        continue
                    for m in pattern.finditer(cm.Lines(1, cm.CountOfLines)):
                        name = m.group(1)
                        if name not in macros:
                            macros.append(name)
            except Exception as ex:
                print(f"[DEBUG] Failed to access VBComponents of {proj_name}: {ex}", file=sys.stderr)
                continue
            all_results[proj_name] = macros

        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": True, "file": "all", "macros": all_results}, ensure_ascii=False), file=sys.stdout)
            return True

        for p_name, macros in all_results.items():
            print(f"\n--- {p_name} ---")
            print(f"マクロ数: {len(macros)}")
            for name in macros:
                print(f"MACRO:{name}")
        return True

    proj = target_projects[0]
    proj_name = get_project_display_name(proj)
    mod_filter = getattr(args, 'module_opt', None)
    detail = getattr(args, 'detail', False)
    macros = []            # 従来互換の名前リスト
    details = []           # --detail / --json 用
    try:
        for comp in proj.VBComponents:
            if getattr(args, 'standard', False) and comp.Type != 1:
                continue
            if mod_filter and comp.Name.lower() != mod_filter.lower():
                continue
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            code = cm.Lines(1, cm.CountOfLines)
            lines = code.split('\r\n')
            for m in pattern.finditer(code):
                name = m.group(1)
                if name in macros:
                    continue
                macros.append(name)
                if not (detail or getattr(args, 'json', False)):
                    continue
                info = {'module': comp.Name, 'name': name}
                try:
                    info['lines'] = cm.ProcCountLines(name, 0)
                    start = cm.ProcStartLine(name, 0)
                    body_start = cm.ProcBodyLine(name, 0)
                    # 宣言の次行が先頭コメントならそれを1行だけ添える（機械的抽出）
                    if body_start < len(lines):
                        first = lines[body_start].strip()   # body_start は1始まり＝宣言行、次行は index body_start
                        if first.startswith("'"):
                            info['comment'] = first.lstrip("'").strip()
                except Exception:
                    pass
                details.append(info)
    except Exception as ex:
        print(f"エラー: VBComponentsへのアクセスに失敗しました: {ex}", file=sys.stderr)
        return False

    if getattr(args, 'json', False):
        import json
        payload = {"success": True, "file": proj_name, "macros": macros}
        if details:
            payload["details"] = details
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"対象ブック: {proj_name}")
    print(f"マクロ数: {len(macros)}")
    if detail:
        for d in details:
            extra = f", {d['lines']}行" if 'lines' in d else ""
            cmt = f"  '{d['comment']}" if 'comment' in d else ""
            print(f"MACRO:[{d['module']}] {d['name']}{extra}{cmt}")
    else:
        for name in macros:
            print(f"MACRO:{name}")
    return True


def cmd_list_modules(args):
    """モジュール一覧"""
    load_addins = getattr(args, 'personal', False) or getattr(args, 'addin', False) or getattr(args, 'all', False)
    target_file, _ = parse_target_and_rest(args.posargs)
    # 参照するだけのコマンド → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, default_wb = get_workbook(target_file, load_addins=load_addins, readonly=True)

    # 全プロジェクトをリスト化 (ビジーエラー対策としてリトライ)
    import time
    all_projects = []
    for attempt in range(5):
        try:
            all_projects = []
            for p in xl.VBE.VBProjects:
                all_projects.append(p)
            break
        except Exception as ex:
            if "800ac472" in str(ex) and attempt < 4:
                time.sleep(0.5)
                continue
            print(f"エラー: VBAプロジェクトモデルへのアクセスが拒否されました: {ex}", file=sys.stderr)
            return False

    # 対象のプロジェクトを選択
    target_projects = []
    if getattr(args, 'all', False):
        target_projects = all_projects
    elif getattr(args, 'personal', False):
        found = None
        for p in all_projects:
            try:
                fname = os.path.basename(p.Filename).lower()
                if fname in ("personal.xlsb", "personal.xls"):
                    found = p
                    break
            except Exception:
                continue
        if not found:
            print("エラー: 個人用マクロブック (PERSONAL.XLSB) がロードされていません。", file=sys.stderr)
            return False
        target_projects.append(found)
    elif getattr(args, 'addin', False):
        target_addin = _select_addin_project(all_projects, args.addin)
        if target_addin is None:
            return False
        target_projects.append(target_addin)
    else:
        found = None
        try:
            for p in all_projects:
                try:
                    if p.Filename.lower() == default_wb.FullName.lower():
                        found = p
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not found:
            try:
                found = default_wb.VBProject
            except Exception as ex:
                print(f"エラー: VBProjectの取得に失敗しました: {ex}", file=sys.stderr)
                return False
        target_projects.append(found)

    type_names = {1: '標準モジュール', 2: 'クラスモジュール',
                  3: 'フォーム', 100: 'シート/ThisWorkbook'}

    def get_project_display_name(p):
        try:
            if p.Filename:
                return os.path.basename(p.Filename)
        except Exception:
            pass
        return p.Name

    is_all = getattr(args, 'all', False)
    if is_all:
        all_results = {}
        for proj in target_projects:
            modules = []
            proj_name = get_project_display_name(proj)
            try:
                for comp in proj.VBComponents:
                    tname = type_names.get(comp.Type, f'Type={comp.Type}')
                    modules.append({"name": comp.Name, "type": tname, "type_id": comp.Type})
            except Exception as ex:
                print(f"[DEBUG] Failed to access VBComponents of {proj_name}: {ex}", file=sys.stderr)
                continue
            all_results[proj_name] = modules

        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": True, "file": "all", "modules": all_results}, ensure_ascii=False), file=sys.stdout)
            return True

        for p_name, modules in all_results.items():
            print(f"\n--- {p_name} ---")
            for m in modules:
                print(f"MODULE:{m['name']}  ({m['type']})")
        return True

    proj = target_projects[0]
    proj_name = get_project_display_name(proj)
    modules = []
    try:
        for comp in proj.VBComponents:
            tname = type_names.get(comp.Type, f'Type={comp.Type}')
            modules.append({"name": comp.Name, "type": tname, "type_id": comp.Type})
    except Exception as ex:
        print(f"エラー: VBComponentsへのアクセスに失敗しました: {ex}", file=sys.stderr)
        return False

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": proj_name, "modules": modules}, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"対象ブック: {proj_name}")
    for m in modules:
        print(f"MODULE:{m['name']}  ({m['type']})")
    return True


def cmd_list_forms(args):
    """フォーム一覧（UserForm のみ抽出。キャプション・コントロール数・コード行数付き）"""
    load_addins = getattr(args, 'personal', False) or getattr(args, 'addin', False) or getattr(args, 'all', False)
    target_file, _ = parse_target_and_rest(args.posargs)
    # 参照するだけのコマンド → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, default_wb = get_workbook(target_file, load_addins=load_addins, readonly=True)

    # 全プロジェクトをリスト化 (ビジーエラー対策としてリトライ)
    import time
    all_projects = []
    for attempt in range(5):
        try:
            all_projects = []
            for p in xl.VBE.VBProjects:
                all_projects.append(p)
            break
        except Exception as ex:
            if "800ac472" in str(ex) and attempt < 4:
                time.sleep(0.5)
                continue
            print(f"エラー: VBAプロジェクトモデルへのアクセスが拒否されました: {ex}", file=sys.stderr)
            return False

    # 対象のプロジェクトを選択（list-modules と同じ流儀）
    target_projects = []
    if getattr(args, 'all', False):
        target_projects = all_projects
    elif getattr(args, 'personal', False):
        found = None
        for p in all_projects:
            try:
                fname = os.path.basename(p.Filename).lower()
                if fname in ("personal.xlsb", "personal.xls"):
                    found = p
                    break
            except Exception:
                continue
        if not found:
            print("エラー: 個人用マクロブック (PERSONAL.XLSB) がロードされていません。", file=sys.stderr)
            return False
        target_projects.append(found)
    elif getattr(args, 'addin', False):
        target_addin = _select_addin_project(all_projects, args.addin)
        if target_addin is None:
            return False
        target_projects.append(target_addin)
    else:
        found = None
        try:
            for p in all_projects:
                try:
                    if p.Filename.lower() == default_wb.FullName.lower():
                        found = p
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not found:
            try:
                found = default_wb.VBProject
            except Exception as ex:
                print(f"エラー: VBProjectの取得に失敗しました: {ex}", file=sys.stderr)
                return False
        target_projects.append(found)

    def get_project_display_name(p):
        try:
            if p.Filename:
                return os.path.basename(p.Filename)
        except Exception:
            pass
        return p.Name

    def collect_forms(proj):
        forms = []
        for comp in proj.VBComponents:
            if comp.Type != 3:                   # 3 = MSForm (UserForm)
                continue
            info = {"name": comp.Name}
            try:
                info["code_lines"] = comp.CodeModule.CountOfLines
            except Exception:
                pass
            try:
                d = comp.Designer
                if d is not None:
                    info["caption"] = d.Caption
                    info["controls"] = d.Controls.Count
                else:
                    info["loaded"] = True        # 表示中/ロード中は Designer が取れない
            except Exception:
                info["loaded"] = True
            forms.append(info)
        return forms

    def format_form(f):
        parts = [f"FORM:{f['name']}"]
        if f.get("loaded"):
            parts.append("(ロード中: 詳細取得不可)")
        else:
            if "controls" in f:
                parts.append(f"コントロール{f['controls']}個")
            if f.get("caption") and f["caption"] != f["name"]:
                parts.append(f"[{f['caption']}]")
        if "code_lines" in f:
            parts.append(f"コード{f['code_lines']}行")
        return "  ".join(parts)

    if getattr(args, 'all', False):
        all_results = {}
        for proj in target_projects:
            proj_name = get_project_display_name(proj)
            try:
                all_results[proj_name] = collect_forms(proj)
            except Exception as ex:
                print(f"[DEBUG] Failed to access VBComponents of {proj_name}: {ex}", file=sys.stderr)
                continue

        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": True, "file": "all", "forms": all_results}, ensure_ascii=False), file=sys.stdout)
            return True

        for p_name, forms in all_results.items():
            print(f"\n--- {p_name} ---")
            if not forms:
                print("（フォームなし）")
            for f in forms:
                print(format_form(f))
        return True

    proj = target_projects[0]
    proj_name = get_project_display_name(proj)
    try:
        forms = collect_forms(proj)
    except Exception as ex:
        print(f"エラー: VBComponentsへのアクセスに失敗しました: {ex}", file=sys.stderr)
        return False

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": proj_name, "forms": forms}, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"対象ブック: {proj_name}")
    print(f"フォーム数: {len(forms)}")
    for f in forms:
        print(format_form(f))
    return True


def _all_procedure_names(wb):
    """ブック内の全プロシージャ名を列挙（did-you-mean 用の機械的リスト）"""
    pat = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    names = []
    try:
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            for m in pat.finditer(cm.Lines(1, cm.CountOfLines)):
                if m.group(1) not in names:
                    names.append(m.group(1))
    except Exception:
        pass
    return names


def _suggest_similar(name, candidates, label="もしかして", wb=None):
    """タイポ候補の提示（difflib による機械的な近似のみ・判断はしない）。

    wb を渡すと、ほかに開いているブックにその名前があるかも見て名指しする（.xlsm 等の 1 冊に決まるときは
    get / patch / replace が _book_holding_proc で自分から移るので、ここに来るのは .xlam だけ・2 冊以上のとき）。秀コンボ.xlsm を直すつもりで 実測4.xlsx が前に出ていて
    「見つかりません」だけが返り、どこにあるかを自分で探した（2026-09-24）。
    """
    import difflib
    close = difflib.get_close_matches(name, candidates, n=3, cutoff=0.6)
    if close:
        print(f"  {label}: {' / '.join(close)}")
    if wb is not None:
        for book, mod in _other_books_with_proc(wb, name):
            how = ("アドイン＝直すなら元の .xlsm を前に出して直し、更新登録する" if book.lower().endswith('.xlam')
                   else "そのブックを前に出してから撃ち直す")
            print(f"  ほかに開いているブックにあります: {book}（{mod}）→ {how}")
    print("  py vba_manager.py list で一覧を確認できます。")


def _other_books_with_proc(wb, name):
    """アクティブ以外の開いているブック（アドインを含む）で name のプロシージャを持つもの → [(ブック名, モジュール名)]。"""
    pat = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+' + re.escape(name) + r'\s*\(',
                     re.IGNORECASE | re.MULTILINE)
    out = []
    try:
        me = str(wb.Name)
        projects = wb.Application.VBE.VBProjects
    except Exception:
        return out
    for p in projects:
        try:
            fn = str(p.FileName or '')
            book = os.path.basename(fn)
            if not book or book.lower() == me.lower():
                continue
            for comp in p.VBComponents:
                cm = comp.CodeModule
                n = cm.CountOfLines
                if n and pat.search(cm.Lines(1, n)):
                    out.append((book, str(comp.Name)))
                    break
        except Exception:
            continue
    return out


def _has_proc(wb, name):
    """wb のどこかのモジュールに name のプロシージャがあるか。"""
    for comp in wb.VBProject.VBComponents:
        try:
            comp.CodeModule.ProcStartLine(name, 0)
            return True
        except Exception:
            pass
    return False


def _book_holding_proc(wb, name):
    """アクティブの wb に name が無く、ほかに開いている .xlsm 等（.xlam を除く）の 1 冊にだけあれば、そのブックを返す。

    棚のマクロを直すとき、前に出ているのは試しの表（実測4.xlsx）で、マクロは秀コンボ.xlsm にある。
    ブックを名指しする口が MCP の patch_procedure に無く、「見つかりません」→ 前に出す手が要った
    （2026-09-24 台帳: 失敗の理由の 1 位）。1 冊に決まるときだけ移る。.xlam は元の .xlsm で直して焼くので移らない。
    """
    try:
        if _has_proc(wb, name):
            return None
        hits = [b for b, _ in _other_books_with_proc(wb, name) if not b.lower().endswith('.xlam')]
        if len(hits) != 1:
            return None
        other = wb.Application.Workbooks(hits[0])
        print(f"（'{name}' は {wb.Name} ではなく {other.Name} にありました。{other.Name} を対象にします）")
        return other
    except Exception:
        return None


def _decl_span(lines, i):
    """lines[i] から始まる宣言の行（行継続 " _" で折り返した分も含む）の終わりの添字。"""
    j = i
    while j + 1 < len(lines) and re.search(r'\s_\s*$', lines[j]):
        j += 1
    return j


def _norm_decl(lines):
    return re.sub(r'\s+', ' ', ' '.join(ln.rstrip().rstrip('_') for ln in lines)).strip().lower()


def _replace_body_keeping_decl(cm, macro_name, new_code, attr_block, end_pattern):
    """宣言の行が新旧で同じなら、宣言の下から End Sub までだけを差し替える → 差し替えたら True。

    同じでない・1 行完結の Sub・形が読めない、は False（呼び元が従来の Remove＋Import に回す）。
    差し替えた後に書き出して Attribute が残っているかを確かめ、消えていれば元の本文に戻して False。
    """
    try:
        start = cm.ProcStartLine(macro_name, 0)
        count = cm.ProcCountLines(macro_name, 0)
        body = cm.ProcBodyLine(macro_name, 0)
    except Exception:
        return False
    total = cm.CountOfLines
    mod_lines = cm.Lines(1, total).split('\r\n')
    d0 = body - 1
    d1 = _decl_span(mod_lines, d0)
    decl_old = mod_lines[d0:d1 + 1]
    if re.search(r'\bEnd\s+(?:Sub|Function)\b', _strip_vba_comment(re.sub(r'"[^"]*"', '""', decl_old[-1])), re.I):
        return False                                   # 1 行完結
    end_idx = None
    for k in range(min(start + count - 2, total - 1), d1, -1):
        if end_pattern.match(mod_lines[k]):
            end_idx = k
            break
    if end_idx is None:
        return False
    new_lines = new_code.replace('\r\n', '\n').split('\n')
    decl_pat = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+'
                          + re.escape(macro_name) + r'\s*[\(\s]', re.I)
    n0 = next((i for i, ln in enumerate(new_lines) if decl_pat.match(ln)), None)
    if n0 is None:
        return False
    n1 = _decl_span(new_lines, n0)
    if _norm_decl(new_lines[n0:n1 + 1]) != _norm_decl(decl_old):
        return False
    rest = new_lines[n1 + 1:]
    while rest and not rest[-1].strip():
        rest.pop()
    if not rest or not end_pattern.match(rest[-1]):
        return False
    old_body = mod_lines[d1 + 1:end_idx + 1]
    if n0 > 0:
        print("  (宣言より上のコメント行は、Attribute のあるマクロでは既存の行を維持します)")
    cm.DeleteLines(d1 + 2, end_idx - d1)
    cm.InsertLines(d1 + 2, '\r\n'.join(rest))
    # Attribute が残ったかを書き出して確かめる（残らなければ元の本文に戻して従来の経路へ）
    try:
        import tempfile
        p = os.path.join(tempfile.gettempdir(), f"_vbam_attr_check_{os.getpid()}.bas")
        cm.Parent.Export(p)
        with open(p, 'rb') as f:
            text = f.read().decode('cp932', errors='replace')
        try:
            os.remove(p)
        except OSError:
            pass
        if all(a.strip() in text for a in attr_block):
            return True
    except Exception:
        pass
    cm.DeleteLines(d1 + 2, len(rest))
    cm.InsertLines(d1 + 2, '\r\n'.join(old_body))
    print("  (本文だけの差し替えで Attribute が残らなかったため、元に戻して replace-module 方式に回します)")
    return False


def _book_holding_module(wb, module_name, include_addins=False):
    """wb にモジュール module_name が無く、ほかに開いているブックの 1 冊にだけあれば、そのブックを返す（_book_holding_proc のモジュール版）。

    前に出ているのが試しの表で、足す先のモジュール（表の整理・コンボ道具）は秀コンボ.xlsm にある、の形で
    add-procedure が「モジュールが見つかりません」と 3 回落ちていた（2026-09-24 総点検・会話記録 30 日）。
    .xlam は元の .xlsm で直して焼くので、書く手では移らない（読むだけの手は include_addins=True）。
    """
    try:
        if any(c.Name.lower() == module_name.lower() for c in wb.VBProject.VBComponents):
            return None
        xl = wb.Application
        hits = []
        for vp in xl.VBE.VBProjects:
            try:
                bn = _project_book_name(xl, vp)
                if not bn or bn == wb.Name or (bn.lower().endswith('.xlam') and not include_addins):
                    continue
                if any(c.Name.lower() == module_name.lower() for c in vp.VBComponents):
                    hits.append(bn)
            except Exception:
                continue
        hits = list(dict.fromkeys(hits))
        if len(hits) > 1:                            # 読むだけの手で xlsm と焼いた xlam の両方にある＝元の xlsm を採る
            hits = [h for h in hits if not h.lower().endswith('.xlam')]
        if len(hits) != 1:
            return None
        other = xl.Workbooks(hits[0])
        print(f"（モジュール '{module_name}' は {wb.Name} ではなく {other.Name} にありました。{other.Name} を対象にします）")
        return other
    except Exception:
        return None


def _extract_proc(wb, module_name, macro_name):
    """1プロシージャのコードを取り出す。

    戻り値: (comp_name, clean_code) / 見つからなければ (None, None)。
    同名複数（module_name 未指定時）は例外 ValueError(候補リスト) を投げる。
    """
    # モジュール未指定時：同名プロシージャが複数モジュールにある場合はエラー
    # （違うフォームの同名イベントを黙って掴む事故を防ぐ。replace-procedure と同じ流儀）
    if not module_name:
        matched = []
        for comp in wb.VBProject.VBComponents:
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                matched.append(comp.Name)
            except Exception:
                pass
        if len(matched) > 1:
            raise ValueError(matched)

    for comp in wb.VBProject.VBComponents:
        if module_name and comp.Name.lower() != module_name.lower():
            continue
        cm = comp.CodeModule
        try:
            proc_start = cm.ProcStartLine(macro_name, 0)
            count      = cm.ProcCountLines(macro_name, 0)
            # ProcStartLine 起点の領域には宣言の上のコメントも含まれる。
            # replace-procedure と対称にし、get→replace の往復で
            # ヘッダーコメントが消えないようにする。
            code = cm.Lines(proc_start, count)
        except Exception:
            continue

        lines = code.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n').split('\n')
        # 先頭の空行を除去（プロシージャ間の区切り空行は領域に含まれるため）
        while lines and lines[0].strip() == '':
            lines.pop(0)
        # 末尾の空行と、紛れ込んだ次プロシージャの宣言行を除去
        # （「Sub B(): 処理: End Sub」のような1行完結プロシージャは正当な本体なので
        #   対象外。cmd_replace_procedure の混入除去と同じ条件に揃える）
        while len(lines) > 1:
            last = lines[-1].strip()
            if last == '':
                lines.pop()
            elif (re.match(
                    r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
                    last, re.IGNORECASE)
                  and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
                lines.pop()
            else:
                break
        return comp.Name, '\n'.join(lines) + '\n'
    if module_name:
        # 指したモジュールに無い＝モジュールを分けた・移したあとの古い名前（2026-09-24 台帳: 棚を 4 つに分けた後の
        # 「get 表の整理 選んだ2列の平均最大最小を作る」が 表の整理_作る にあるのに「見つかりません」で落ちていた）。
        # ほかの 1 か所にだけあれば、そこを取る。2 か所以上なら選ばせる
        others = []
        for comp in wb.VBProject.VBComponents:
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                others.append(comp.Name)
            except Exception:
                pass
        if len(others) > 1:
            raise ValueError(others)
        if len(others) == 1:
            print(f"（'{macro_name}' はモジュール {module_name} ではなく {others[0]} にありました。{others[0]} を使います）")
            return _extract_proc(wb, others[0], macro_name)
    return None, None


def _narrow_proc_range(cm, start, count):
    """ProcStartLine/ProcCountLines の領域を「実体だけ」に絞る。

    ProcCountLines の領域には前後の空行に加え、次プロシージャの宣言行が
    食い込むことがある（1行完結 Sub「Sub X(): Call Main: End Sub」の直後など）。
    get(_extract_proc) と replace-procedure は同じ条件でその行を落としているが、
    delete-procedure だけが生の start/count を DeleteLines に渡しており、
    隣のプロシージャの宣言行ごと消して「削除完了」と報告していた。

    戻り値: (実効start, 実効count)
    """
    lines = cm.Lines(start, count).replace('\r\n', '\n').split('\n')
    lead = 0
    while lead < len(lines) and lines[lead].strip() == '':
        lead += 1
    end = len(lines)
    while end - lead > 1:
        last = lines[end - 1].strip()
        if last == '':
            end -= 1
        elif (re.match(
                r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
                last, re.IGNORECASE)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
            end -= 1
        else:
            break
    return start + lead, end - lead


def _inline_body_after_decl(raw, decl_end):
    """宣言行に同居する本体（1行完結 Sub）を取り出す。無ければ空文字。

    「Sub X(): Call Main: End Sub」の "Call Main: End Sub" の部分を返す。
    引数リストの括弧は対応を取って読み飛ばす（既定値の文字列に ':' や ')' が
    入っていても誤らないように、文字列リテラルの中は数えない）。
    """
    s = raw[decl_end:]
    i = 0
    if i < len(s) and s[i] == '(':
        depth = 0
        in_str = False
        while i < len(s):
            c = s[i]
            if c == '"':
                in_str = not in_str
            elif not in_str:
                if c == '(':
                    depth += 1
                elif c == ')':
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
            i += 1
    rest = s[i:]
    # ':' を探す前に行末コメントを落とす。落とさないと
    # 「Sub X() ' 例: Run "整形" してから」のようなコメント内のコロンを
    # 本体の区切りと誤認し、コメント文をコードとして走査してしまう
    # （コメント内の Call/Run が偽の呼び出しとして call-graph に載る）
    rest = _strip_vba_comment(rest)
    # 引数リストの後ろに ':' があれば、それ以降が同じ行に書かれた本体
    # （'As String' のような戻り型指定はコロンの手前にある）
    if ':' not in rest:
        return ''
    return rest.split(':', 1)[1]


def cmd_get(args):
    """プロシージャのコードを取得・表示・ファイル保存

    書式:
      get <macro_name>                       全モジュールから検索
      get <module_name> <macro_name>         モジュール指定（スペース区切り）
      get <module_name>.<macro_name>         モジュール指定（ドット区切り）
      get <名1> <名2> <名3> ...              3個以上は複数取得（各要素にドット記法可）
                                             ※出力は連結。書き戻しは従来どおり1本ずつ
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: get [excel_file] <macro_name>  または  get [excel_file] <module_name> <macro_name>")
        return False

    # 取得リクエストの解析
    requests = []                     # [(module_name or None, macro_name), ...]
    if len(rest) >= 3:
        # 複数取得モード（1回のCOM接続でまとめ読み）。各要素は 名前 or モジュール.名前
        for token in rest:
            if '.' in token:
                mn, pn = token.split('.', 1)
                requests.append((mn, pn))
            else:
                requests.append((None, token))
    elif len(rest) == 2 and '.' in rest[0] and '.' in rest[1] and not looks_like_xl_file(rest[1]):
        # 両方ドット記法なら複数取得（get A.x B.y を「モジュールA.x のマクロ B.y」と
        # 誤解釈しないため）
        for token in rest:
            mn, pn = token.split('.', 1)
            requests.append((mn, pn))
    elif len(rest) == 2 and not looks_like_xl_file(rest[1]):
        requests.append((rest[0], rest[1]))          # get <module> <macro>
        print(f"モジュール指定: {rest[0]}")
    elif len(rest) == 1 and '.' in rest[0] and not looks_like_xl_file(rest[0]):
        mn, pn = rest[0].split('.', 1)
        requests.append((mn, pn))                    # get <module>.<macro>
        print(f"モジュール指定: {mn}")
    else:
        requests.append((None, rest[0]))

    # コードを読むだけ → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, wb = get_workbook(target_file, readonly=True)
    if not target_file and len(requests) == 1:
        wb = _book_holding_proc(wb, requests[0][1]) or wb

    results = []
    for module_name, macro_name in requests:
        try:
            comp_name, clean = _extract_proc(wb, module_name, macro_name)
        except ValueError as e:
            print(f"エラー: '{macro_name}' が複数のモジュールに存在します:")
            for mn in e.args[0]:
                print(f"  - {mn}")
            print(f"  モジュールを指定してください。例: py vba_manager.py get {e.args[0][0]} {macro_name}")
            return False
        if comp_name is None:
            print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
            _suggest_similar(macro_name, _all_procedure_names(wb), wb=wb)
            return False
        results.append({'module': comp_name, 'name': macro_name, 'code': clean})

    out_path = getattr(args, 'out_opt', None)
    save_path = os.path.abspath(out_path) if out_path else LAST_PROC_FILE
    joined = '\n'.join(r['code'] for r in results)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(joined)

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": wb.Name, "saved": save_path,
                          "procs": results}, ensure_ascii=False), file=sys.stdout)
        return True

    for r in results:
        print(f"モジュール  : {r['module']}")
        print(f"プロシージャ: {r['name']}")
        print(f"保存先      : {save_path}")
        print("=" * 60)
        print(r['code'])
        print("=" * 60)
    if len(results) > 1:
        print(f"（{len(results)}本を連結して保存しました。replace-procedure での書き戻しは1本ずつ）")
    return True


def cmd_replace_procedure(args):
    """プロシージャを置換 (コードファイル省略時は _last_proc.vba を使用)"""
    target_file, rest = parse_target_and_rest(args.posargs)

    # 「replace-procedure モジュール名 [プロシージャ名]」と撃たれたとき（2026-09-24 台帳: BorderFinder FindBorders で
    # 「コードファイルが見つかりません: BorderFinder」と落ち、撃ち直しが 4 回）。位置引数がファイルとして無く、
    # パスにも見えなければモジュール名と読み、コードは _last_proc.vba から取る
    if (rest and not getattr(args, 'code_file_opt', None)
            and not os.path.exists(smart_path_resolve(rest[0]) or rest[0])
            and not re.search(r'[\\/]|\.(?:vba|bas|txt|cls|frm)$', rest[0], re.IGNORECASE)):
        if not getattr(args, 'module_opt', None):
            args.module_opt = rest[0]
            print(f"（{rest[0]} をモジュール名と読みました。コードは {os.path.basename(LAST_PROC_FILE)} から）")
        rest = []

    # コードファイルの決定: --code-file > 位置引数 > _last_proc.vba
    code_file = (getattr(args, 'code_file_opt', None)
                 or (rest[0] if rest else None)
                 or LAST_PROC_FILE)

    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False

    new_code = read_code_file(resolved)

    # 簡易構文チェック・エンコード検証
    if not validate_vba_code(new_code, getattr(args, 'force', False)):
        return False

    # プロシージャ名を特定
    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE
    )
    m = pattern.search(new_code)
    if not m:
        print("エラー: コードファイルに Sub/Function 宣言が見つかりません")
        return False
    macro_name = m.group(1)

    # new_code の末尾に次のプロシージャの Sub/Function 宣言が混入していたら除去。
    # ただし「Sub B(): 処理: End Sub」のような1行完結のプロシージャは正当なコード
    # なので対象外（End を伴わない裸の宣言行だけが混入）
    code_lines = new_code.rstrip('\n').split('\n')
    while len(code_lines) > 1:
        last = code_lines[-1].strip()
        if last == '':
            code_lines.pop()
        elif (pattern.match(last)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)
              and code_lines[-1].strip() != code_lines[0].strip()):
            code_lines.pop()
        else:
            break
    new_code = '\n'.join(code_lines) + '\n'

    # コードファイルに2本以上のプロシージャが入っていたら弾く。
    # get は複数のプロシージャを1ファイルに連結して書ける（cmd_get）が、
    # replace-procedure は先頭1本の「領域」に全文を流し込むため、2本目以降は
    # モジュール内の重複定義になりコンパイルエラーになる。MCP 経由は常に -y で
    # 確認プロンプトも挟まらないので、文言の注意ではなく機械的に止める。
    decl_names = []
    for ln in code_lines:
        s = ln.strip()
        if not s or s.startswith("'") or re.match(r'^Rem\b', s, re.IGNORECASE):
            continue
        md = pattern.match(s)
        if md:
            decl_names.append(md.group(1))
    if len(decl_names) > 1:
        print(f"エラー: コードファイルに {len(decl_names)} 本のプロシージャが入っています:")
        for dn in decl_names:
            print(f"  - {dn}")
        print("  replace-procedure は1本ずつ書き戻してください")
        print("  （先頭の1本しか対象にならず、2本目以降は重複定義になります）。")
        print("  まとめて差し替えるなら replace-module を使ってください。")
        return False

    xl, wb = get_workbook(target_file)
    if not target_file:
        wb = _book_holding_proc(wb, macro_name) or wb

    # --module 未指定時：同名プロシージャが複数モジュールにある場合はエラー
    module_opt = getattr(args, 'module_opt', None)
    if not module_opt:
        matched_modules = []
        for comp in wb.VBProject.VBComponents:
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                matched_modules.append(comp.Name)
            except Exception:
                pass
        if len(matched_modules) > 1:
            print(f"エラー: '{macro_name}' が複数のモジュールに存在します:")
            for mn in matched_modules:
                print(f"  - {mn}")
            print(f"  --module オプションで対象を指定してください。")
            print(f"  例: py vba_manager.py replace-procedure --module {matched_modules[0]}")
            return False

    if make_backup(wb.FullName, macro_name) is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        print("  ※ 未保存の新規ブックはバックアップできません。一度保存してから実行してください。")
        return False

    # 対象プロシージャの確認と差分表示
    target_comp = None
    proc_start = 0
    proc_count = 0
    for comp in wb.VBProject.VBComponents:
        if module_opt and comp.Name.lower() != module_opt.lower():
            continue
        cm = comp.CodeModule
        try:
            proc_start = cm.ProcStartLine(macro_name, 0)
            proc_count = cm.ProcCountLines(macro_name, 0)
            target_comp = comp
            break
        except Exception:
            continue

    if not target_comp:
        print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
        _suggest_similar(macro_name, _all_procedure_names(wb), wb=wb)
        print("  新規追加なら add-procedure を使ってください。")
        return False

    # 変更前コードの取得と差分表示
    # ProcStartLine 起点の領域には前後の空行が含まれるため、実置換範囲は
    # 空行を除いて絞る（プロシージャ間の区切り空行を消さないため）
    old_code = target_comp.CodeModule.Lines(proc_start, proc_count)
    old_all = old_code.replace('\r\n', '\n').split('\n')
    lead = 0
    while lead < len(old_all) and old_all[lead].strip() == '':
        lead += 1
    # 末尾は空行だけでなく「次プロシージャの宣言行」も削除範囲から外す。
    # ProcCountLines の領域には次の宣言行が食い込むことがあり（1行完結 Sub の直後など）、
    # get(_extract_proc) と new_code のサニタイザは同じ条件でその行を落とす。
    # ここだけ削除範囲に含めると、消した宣言行が new_code から復元されず
    # 次のプロシージャが宣言を失って壊れる（両者を対称に保つ）。
    end = len(old_all)
    while end - lead > 1:
        last = old_all[end - 1].strip()
        if last == '':
            end -= 1
        elif (pattern.match(last)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
            end -= 1
        else:
            break
    eff_start = proc_start + lead
    eff_count = end - lead

    import difflib
    old_lines = old_all[lead:end]
    new_lines = new_code.replace('\r\n', '\n').split('\n')
    if new_lines and new_lines[-1] == '': new_lines.pop()

    diff = list(difflib.unified_diff(
        old_lines, new_lines,
        fromfile=f"Current: {target_comp.Name}.{macro_name}",
        tofile=f"New: {macro_name}",
        lineterm=""
    ))

    if diff:
        print("\n--- 変更差分 (Diff) ---")
        for line in diff:
            print(line)
        print("----------------------\n")
    else:
        # 変更ゼロなら置換しない（Attribute経路だと無変更でも Remove+Import が走り、
        # 無用なリスクを負うだけのため）
        print("変更はありません。置換をスキップしました。")
        return True

    # 確認プロンプト
    if not getattr(args, 'yes', False):
        ans = input(f"プロシージャ '{macro_name}' を置換しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    # モジュール単位のバックアップ（Attribute経路の Import 失敗時の復旧素材を兼ねる）
    module_backup = make_module_backup(wb, target_comp.Name)

    print(f"プロシージャ '{macro_name}' を置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() != target_comp.Name.lower():
            continue
        cm = comp.CodeModule

        # モジュールをエクスポートして Attribute行の有無を確認
        module_name = comp.Name
        tmp_bas = os.path.join(SCRIPT_DIR, f"_tmp_{module_name}.bas")
        comp.Export(tmp_bas)

        with open(tmp_bas, 'rb') as f:
            bas_content = f.read().decode('cp932')
        bas_lines = bas_content.split('\r\n')

        # .bas 内で対象プロシージャの Sub/Function 宣言行を探す
        proc_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
            r'(?:Sub|Function)\s+' + re.escape(macro_name) + r'\s*[\(\s]',
            re.IGNORECASE
        )
        # 末尾コメント（End Sub 'xxx）を許容しないと次のプロシージャの End Sub まで
        # スキャンが伸び、置換範囲が次のプロシージャを丸ごと巻き込む（消失事故）
        end_pattern = re.compile(
            r"^\s*End\s+(?:Sub|Function)\s*(?:'.*)?$", re.IGNORECASE
        )

        sub_line_idx = None
        proc_end_idx = None
        attr_block = []

        for idx, line in enumerate(bas_lines):
            if sub_line_idx is None and proc_pattern.match(line):
                sub_line_idx = idx
                # Sub宣言の直後の Attribute行を収集
                # （宣言が行継続 " _" で複数行の場合は継続行を読み飛ばしてから収集。
                #   読み飛ばさないと Attribute を見逃し、ショートカット定義が失われる）
                check = idx + 1
                while check < len(bas_lines) and re.search(r'\s_\s*$', bas_lines[check - 1]):
                    check += 1
                while check < len(bas_lines) and bas_lines[check].strip().startswith('Attribute '):
                    attr_block.append(bas_lines[check])
                    check += 1
                # 「Sub X(): 処理: End Sub」の1行完結は宣言行自身で閉じている。
                # 次の End Sub まで探すと後続プロシージャを巻き込んで消すため、
                # ここで終端を確定する（随伴 Attribute 行までが置換対象）。
                # 文字列リテラルだけでなくコメントも落とす。落とさないと
                # 「Public Sub 印刷実行()  ' 途中で Exit せず End Sub まで通す」のような
                # 行末コメントを1行完結Subと誤判定し、置換範囲が宣言行だけになって
                # 旧本体が残る（新旧の本体が並んで二重化・構文破壊）
                no_str = _strip_vba_comment(re.sub(r'"[^"]*"', '""', line))
                if re.search(r'\bEnd\s+(?:Sub|Function)\b', no_str, re.IGNORECASE):
                    proc_end_idx = check - 1
                    break
            elif sub_line_idx is not None and end_pattern.match(line):
                proc_end_idx = idx
                break

        if sub_line_idx is None or proc_end_idx is None:
            _remove_export_artifacts(tmp_bas)
            continue

        if not attr_block:
            # Attribute行なし → 従来の InsertLines 方式（高速・モジュール順維持）
            # InsertLines は末尾改行を余分な空行として挿入するため取り除く
            _remove_export_artifacts(tmp_bas)
            cm.DeleteLines(eff_start, eff_count)
            cm.InsertLines(eff_start, new_code.rstrip('\n'))
            wb.Save()
            print(f"置換完了: [{comp.Name}] '{macro_name}' → 保存しました")
            note_if_macro_free_book(wb)
            return True

        # Attribute行あり、でも宣言の行が変わらないなら、宣言の下（本文〜End Sub）だけを差し替える（2026-09-24）。
        # 隠れた Attribute（ショートカット）は宣言の行に付いているので残る（試しのブックで確かめた）。
        # Remove＋Import を通らない＝「実行中のコードがあって旧モジュールが消えず shu0051 として入る」
        # 名前衝突（会話記録 30 日で 3 回・手で改名するしかなかった）が起きない。モジュールの並びも動かない
        if _replace_body_keeping_decl(cm, macro_name, new_code, attr_block, end_pattern):
            _remove_export_artifacts(tmp_bas)
            wb.Save()
            print(f"置換完了: [{comp.Name}] '{macro_name}' → 保存しました（宣言の行とショートカットはそのまま・本文だけ差し替え）")
            note_if_macro_free_book(wb)
            return True

        # Attribute行あり → .bas編集 → replace-module 方式
        print(f"  (Attribute行検出 → replace-module方式で処理)")

        # new_code の行を準備（Sub宣言の直後に Attribute行を挿入）
        # この方式は .bas の「宣言行〜End Sub」だけを差し替えるため、
        # 宣言より上のコメントは .bas 側の既存行をそのまま維持し、
        # 新コード側の宣言より上の行は使わない（使うと二重になる）。
        new_lines = new_code.rstrip('\n').split('\n')
        sub_decl_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
            re.IGNORECASE
        )
        decl_idx = None
        for ni, nl in enumerate(new_lines):
            if sub_decl_pattern.match(nl):
                decl_idx = ni
                break
        if decl_idx is None:
            _remove_export_artifacts(tmp_bas)
            print("エラー: 置換コードに Sub/Function 宣言が見つかりません")
            return False
        if decl_idx > 0:
            print("  (宣言より上のコメント行は、Attribute方式では .bas 側の既存行を維持します)")
            new_lines = new_lines[decl_idx:]
        # Attribute行の挿入位置は「宣言の実体が終わった直後」。
        # 宣言が行継続 " _" で折り返している場合（Sub X(ByVal a As Long, _ ）に
        # 1行目の直後へ入れると Attribute が継続行の前に割り込んで構文破壊になる。
        # 読み取り側（上の attr_block 収集）と対称に、継続行を読み飛ばしてから入れる
        insert_at = 1
        while insert_at < len(new_lines) and re.search(r'\s_\s*$', new_lines[insert_at - 1]):
            insert_at += 1
        for ai, al in enumerate(attr_block):
            new_lines.insert(insert_at + ai, al)

        # .bas 内の対象プロシージャを置換（Sub宣言行から End Sub まで）
        bas_lines[sub_line_idx:proc_end_idx + 1] = new_lines
        new_bas = '\r\n'.join(bas_lines)

        # --force で CP932 に無い文字（絵文字・機種依存文字等）が混ざると
        # ここで UnicodeEncodeError になる。生の例外で落とすと一時 .bas が残るので
        # 握りつぶさず・分かる形で報告して片付ける
        try:
            encoded_bas = new_bas.encode('cp932')
        except UnicodeEncodeError as ex:
            bad = new_bas[ex.start:ex.end]
            _remove_export_artifacts(tmp_bas)
            print(f"エラー: 置換コードに CP932（Shift-JIS）で扱えない文字があります: {bad!r}")
            print("  VBA モジュールは CP932 で保存されます。該当文字を通常の文字に置き換えてください。")
            return False

        with open(tmp_bas, 'wb') as f:
            f.write(encoded_bas)

        # Remove + Import（例外時も DisplayAlerts を戻し、一時ファイルを残さない）
        xl.DisplayAlerts = False
        removed = False
        imported = False
        try:
            wb.Save()
            pythoncom.PumpWaitingMessages()
            wb.VBProject.VBComponents.Remove(comp)
            removed = True
            pythoncom.PumpWaitingMessages()
            # 実名検証つき Import（_wait_component_gone で動的消滅検知）
            _import_module_verified(wb, tmp_bas, module_name)
            # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない。
            # この印を立てずに復旧へ落ちると、正しく入った新モジュールにバックアップを
            # 重ね Import して、ツール自身が連番モジュールを作ってしまう（2026-07-14 発見）
            imported = True
            _save_with_retry(wb)
        except ModuleNameCollisionError as ex:
            # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
            _print_collision_guidance(ex, module_name, module_backup)
            return False
        except Exception as ex:
            print(f"エラー: 置換中に失敗しました: {ex}")
            if imported:
                # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
                _print_save_failed_guidance(module_name)
            elif removed:
                # Remove 成功後の Import 失敗＝開いているブックからモジュール消失。
                # 直前のモジュールバックアップ（置換前の内容）から自動復旧を試みる。
                print(f"⚠ モジュール '{module_name}' は Remove 済みです。バックアップから復旧を試みます...")
                try:
                    if module_backup and os.path.exists(module_backup):
                        _import_module_verified(wb, module_backup, module_name)
                        print(f"復旧成功: {module_backup} を再インポートしました（置換前の内容に戻っています）")
                    else:
                        raise RuntimeError("モジュールバックアップがありません")
                except ModuleNameCollisionError as ex2:
                    print(f"復旧の再 Import で名前衝突: {ex2}")
                    print(f"  置換前の内容は別名モジュール '{ex2.actual_name}' 側にあります。")
                    print(f"  旧 '{module_name}' が消えているのを確認してから '{ex2.actual_name}' を改名してください。")
                except Exception as ex2:
                    print(f"復旧失敗: {ex2}")
                    print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。")
                    print("  対処: ブックを『保存せずに閉じて』開き直せば置換前の状態に戻ります。")
                    if module_backup:
                        print(f"  または backups のバックアップを手動で Import: {module_backup}")
            return False
        finally:
            xl.DisplayAlerts = True
            if os.path.exists(tmp_bas):
                _remove_export_artifacts(tmp_bas)
        # ここまで来たら Import 済み・実名検証済み（黙って成功と報告しない、の実装）
        print(f"置換完了: [{module_name}] '{macro_name}' → 保存しました (Attribute保持)")
        note_if_macro_free_book(wb)
        return True

    print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
    return False


def _patch_loose(code, target, replacement):
    """字下げ・行末の空白を無視して target の行の並びを code から探し、1 か所なら置き換える → (新しいコード, 見つかった数)。

    見つからない・2 か所以上なら (None, 数)。replacement の字下げは、見つかった先頭行と target の先頭行の差だけ足す
    （target を字下げなしで渡した AI は、replacement も字下げなしで渡すため）。
    """
    t_lines = target.strip('\n').split('\n')
    t_key = [ln.strip() for ln in t_lines]
    if not any(t_key):
        return None, 0
    c_lines = code.split('\n')
    c_key = [ln.strip() for ln in c_lines]
    n = len(t_key)
    hits = [i for i in range(len(c_key) - n + 1) if c_key[i:i + n] == t_key]
    if len(hits) != 1:
        return None, len(hits)
    i = hits[0]
    first = next(k for k in range(n) if t_key[k])
    indent = lambda s: len(s) - len(s.lstrip(' \t'))
    shift = indent(c_lines[i + first]) - indent(t_lines[first])
    r_lines = replacement.strip('\n').split('\n') if replacement.strip('\n') else []
    if shift > 0:
        r_lines = [(' ' * shift + ln) if ln.strip() else ln for ln in r_lines]
    return '\n'.join(c_lines[:i] + r_lines + c_lines[i + n:]), 1


def _print_nearest_lines(code, target, comp_name):
    """target の各行にいちばん近い実物の行を、行番号つきで見せる（どこがずれているかを AI が 1 手で知れるように）。"""
    import difflib
    c_lines = code.split('\n')
    keys = [ln.strip() for ln in c_lines]
    shown = 0
    for t in [ln.strip() for ln in target.split('\n') if ln.strip()]:
        if t in keys:
            continue                               # この行は実物にある＝ずれているのは別の行
        best = difflib.get_close_matches(t, keys, n=1, cutoff=0.5)
        if best:
            k = keys.index(best[0])
            print(f"  target の行: {t}")
            print(f"  実物の近い行（{comp_name} のプロシージャ内 {k + 1} 行目）: {c_lines[k].strip()}")
        else:
            print(f"  target の行: {t}  ← 近い行がありません")
        shown += 1
        if shown >= 3:
            break
    if not shown:
        print("  各行は実物にありますが、並び（間の行・空行）が違います。get で今のコードを取り直してください。")


def cmd_patch_procedure(args):
    """プロシージャ内の指定コードをピンポイントで置換（差分置換）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: patch-procedure [excel_file] <macro_name> --target '置換前' --replacement '置換後' [--module 名] [-y]")
        return False

    macro_name = rest[0]
    module_opt = getattr(args, 'module_opt', None)
    # get と同じく「モジュール プロシージャ」「モジュール.プロシージャ」も受ける（2026-09-24。前は 1 つ目をプロシージャ名と読み、
    # SKILL.md の手本「patch-procedure モジュール Sub --old … --new …」がモジュール名を探して「見つかりません」になった）
    if len(rest) >= 2:
        module_opt = module_opt or rest[0]
        macro_name = rest[1]
    elif '.' in macro_name and not module_opt:
        module_opt, macro_name = macro_name.split('.', 1)
    target_str = getattr(args, 'target', None)
    replacement_str = getattr(args, 'replacement', None)
    target_file_opt = getattr(args, 'target_file_opt', None)
    replacement_file_opt = getattr(args, 'replacement_file_opt', None)

    if target_file_opt and os.path.exists(target_file_opt):
        with open(target_file_opt, 'r', encoding='utf-8', errors='replace') as f:
            target_str = f.read()
    if replacement_file_opt and os.path.exists(replacement_file_opt):
        with open(replacement_file_opt, 'r', encoding='utf-8', errors='replace') as f:
            replacement_str = f.read()

    if target_str is None:
        print("エラー: --target または --target-file で置換対象の文字列を指定してください")
        return False
    if replacement_str is None:
        replacement_str = ""

    # 改行コードの正規化
    target_str = target_str.replace('\r\n', '\n')
    replacement_str = replacement_str.replace('\r\n', '\n')

    xl, wb = get_workbook(target_file)
    if not target_file:
        other = _book_holding_proc(wb, macro_name)
        if other is not None:
            wb, target_file = other, str(other.FullName)   # 下の replace-procedure にも同じブックを渡す
    try:
        comp_name, current_code = _extract_proc(wb, module_opt, macro_name)
    except ValueError as matched:
        print(f"エラー: '{macro_name}' が複数のモジュールに存在します: {matched}")
        print("  --module で対象モジュールを指定してください。")
        return False

    if not comp_name or not current_code:
        print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
        _suggest_similar(macro_name, _all_procedure_names(wb), wb=wb)
        return False

    current_norm = current_code.replace('\r\n', '\n')

    # 出現回数の確認
    count = current_norm.count(target_str)
    new_code = None
    if count == 0:
        # 字下げ・行末の空白だけが違うなら、行の中身で突き合わせて当てる（2026-09-24 台帳: patch-procedure の
        # 失敗 29 回の多くが字下げ違いで、同じ手を 5 回撃ち直していた＝「見つかりません」しか返さなかったため）
        new_code, loose = _patch_loose(current_norm, target_str, replacement_str)
        if new_code is not None:
            print("（字下げ・行末の空白の違いを無視して当てました）")
        elif loose > 1:
            print(f"エラー: target 文字列が（字下げを無視すると）プロシージャ '{macro_name}' 内に {loose} 箇所あります。")
            print("  前後の行も含めて、置換対象が一意に特定できる範囲を指定してください。")
            return False
        else:
            print(f"エラー: target 文字列がプロシージャ '{macro_name}' 内に見つかりません（字下げを無視しても）。")
            _print_nearest_lines(current_norm, target_str, comp_name)
            return False
    elif count > 1:
        print(f"エラー: target 文字列がプロシージャ '{macro_name}' 内に複数（{count}箇所）見つかりました。")
        print("  前後の行も含めて、置換対象が一意に特定できる範囲を指定してください。")
        return False

    # 置換実行
    if new_code is None:
        new_code = current_norm.replace(target_str, replacement_str, 1)

    # _last_proc.vba に書き込んで、実績ある cmd_replace_procedure で適用
    with open(LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(new_code + ('\n' if not new_code.endswith('\n') else ''))

    class _DummyArgs:
        pass
    r_args = _DummyArgs()
    r_args.posargs = [target_file] if target_file else []
    r_args.code_file_opt = LAST_PROC_FILE
    r_args.module_opt = comp_name
    r_args.yes = getattr(args, 'yes', False)
    r_args.force = getattr(args, 'force', False)

    return cmd_replace_procedure(r_args)


def cmd_add_procedure(args):
    """新規プロシージャをモジュール末尾に追加: add-procedure [excel_file] <モジュール名>

    コードは _last_proc.vba（または --code-file）から。replace-procedure が
    既存置換専用なのに対し、こちらは「新しい Sub を1本足す」軽量経路
    （InsertLines のみ・Remove+Import 不要）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: add-procedure [excel_file] <モジュール名> [--code-file f] [-y]")
        print("  追加するコードは _last_proc.vba（または --code-file）に置く")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: add-procedure [excel_file] <モジュール名>'):
        return False

    code_file = getattr(args, 'code_file_opt', None) or LAST_PROC_FILE
    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False
    new_code = read_code_file(resolved)
    if not validate_vba_code(new_code, force=getattr(args, 'force', False)):
        return False
    # コードファイルに複数本入っていることがある（get は3本以上を連結して
    # 1つの _last_proc.vba に書く）。先頭1本だけ見て重複検査すると、2本目以降が
    # 既存と同名でも素通りし、同一モジュール内に二重定義されてコンパイル不能になる
    decl_names = re.findall(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(]+)', new_code,
        re.MULTILINE | re.IGNORECASE)
    if not decl_names:
        print("エラー: コードに Sub/Function 宣言が見つかりません")
        return False
    proc_name = decl_names[0]
    # コードファイル自身の中の同名重複も弾く（VBA は大小文字を区別しない）
    _lowered = [n.lower() for n in decl_names]
    for n in decl_names:
        if _lowered.count(n.lower()) > 1:
            print(f"エラー: コードファイル内に同名のプロシージャが複数あります: {n}")
            return False

    xl, wb = get_workbook(target_file)
    if not target_file:
        wb = _book_holding_module(wb, module_name) or wb
    comp = None
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            comp = c
            break
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が {wb.Name} に見つかりません（list-modules で確認・"
              f"新しく作るなら add-module {module_name} -y）")
        return False
    cm = comp.CodeModule
    # 同名の重複挿入を防止（ブック全体はマクロ実行時の曖昧さになるだけだが、
    # 同一モジュール内はコンパイルエラーになるため必ず止める）。
    # コードファイル内の全宣言について検査する（先頭1本だけでは 2本目以降が素通りする）
    exists = False
    for _n in decl_names:
        try:
            cm.ProcStartLine(_n, 0)
            proc_name = _n          # 実際に衝突した名前を報告する
            exists = True
            break
        except Exception:
            continue
    if exists:
        print(f"エラー: '{proc_name}' は [{comp.Name}] に既に存在します。修正なら replace-procedure を使ってください。")
        return False

    print(f"--- 追加するプロシージャ: [{comp.Name}] {proc_name} ---")
    print(new_code.rstrip('\n'))
    print("-" * 40)
    if not getattr(args, 'yes', False):
        ans = input(f"モジュール '{comp.Name}' の末尾に追加しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    if make_backup(wb.FullName, f"add_{proc_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False

    body = new_code.rstrip('\n')
    if cm.CountOfLines > 0:
        body = '\n' + body     # 既存コードとの区切りの空行
    cm.InsertLines(cm.CountOfLines + 1, body)
    wb.Save()
    print(f"追加完了: [{comp.Name}] '{proc_name}' → 保存しました")
    note_if_macro_free_book(wb)
    return True


def cmd_add_module(args):
    """新規モジュールを追加: add-module [excel_file] <モジュール名> [--type std|class|form]

    add-procedure/replace-module は既存モジュール前提のため、まっさらなブックには
    器を作れなかった。これはその穴を埋める純機械コマンド（VBComponents.Add）。
    std=標準モジュール / class=クラスモジュール / form=UserForm。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: add-module [excel_file] <モジュール名> [--type std|class|form]")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: add-module [excel_file] <モジュール名> [--type std|class|form]'):
        return False

    type_opt = (getattr(args, 'type_opt', None) or 'std').lower()
    type_map = {'std': 1, 'standard': 1, 'class': 2, 'cls': 2, 'form': 3, 'userform': 3}
    if type_opt not in type_map:
        print(f"エラー: 不明な --type '{type_opt}'（std|class|form のいずれか）")
        return False
    comp_type = type_map[type_opt]

    reason = check_vba_identifier(module_name)
    if reason:
        print(f"エラー: モジュール名が VBA の識別子規則に反しています: {reason}")
        return False

    xl, wb = get_workbook(target_file)
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            print(f"エラー: モジュール '{c.Name}' は既に存在します。（別名にするか delete-module で消してから）")
            return False

    if make_backup(wb.FullName, f"add_module_{module_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False

    try:
        comp = wb.VBProject.VBComponents.Add(comp_type)
    except Exception as e:
        print(f"エラー: モジュールを追加できませんでした: {e}")
        print("  VBProject へのアクセスが信頼されているか確認してください（setup-check）。")
        return False
    try:
        comp.Name = module_name
    except Exception as e:
        print(f"エラー: モジュール名 '{module_name}' を設定できませんでした: {e}")
        # 追加済みの既定名（Module1等）のまま残すと、後で保存されたときに
        # ゴミモジュールがブックに焼き付くため撤去する
        try:
            wb.VBProject.VBComponents.Remove(comp)
        except Exception:
            print("警告: 追加途中のモジュールを撤去できませんでした"
                  "（既定名のモジュールが残っていたら手で削除してください）")
        return False
    wb.Save()
    label = {1: '標準モジュール', 2: 'クラスモジュール', 3: 'ユーザーフォーム'}[comp_type]
    print(f"追加完了: {label} '{comp.Name}' → 保存しました")
    note_if_macro_free_book(wb)
    return True


def cmd_delete_procedure(args):
    """プロシージャを削除: delete-procedure [excel_file] <Sub名>

    同名が複数モジュールにある場合は --module で明示（get/replace と同じ流儀）。
    削除対象のコードを表示してから確認（-y でスキップ）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: delete-procedure [excel_file] [モジュール] <Sub名> [Sub名…] [--module 名] [-y]")
        return False
    module_opt = getattr(args, 'module_opt', None)

    xl, wb = get_workbook(target_file)

    # 並びの読み方（2026-09-24）: 「モジュール Sub名」（get と同じ・前は余分な引数で断っていた）／「モジュール.Sub名」／
    # Sub名を複数（13 本を 1 本ずつ消して控えを 13 冊取った＝まとめて 1 回の控えと保存にする）
    comp_names = {c.Name.lower() for c in wb.VBProject.VBComponents}
    if not module_opt and len(rest) == 2 and rest[0].lower() in comp_names:
        module_opt, rest = rest[0], rest[1:]
    wanted = []
    for tok in rest:
        mn, pn = (tok.split('.', 1) if '.' in tok and tok.split('.', 1)[0].lower() in comp_names else (module_opt, tok))
        wanted.append((mn, pn))

    # 対象特定（同名複数はエラーで候補列挙＝対象取り違え防止）。1 本でも決まらなければ何も消さない
    found = []
    for mn, macro_name in wanted:
        matches = []
        for comp in wb.VBProject.VBComponents:
            if mn and comp.Name.lower() != mn.lower():
                continue
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                matches.append(comp)
            except Exception:
                continue
        if not matches:
            print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
            _suggest_similar(macro_name, _all_procedure_names(wb), wb=wb)
            return False
        if len(matches) > 1:
            print(f"エラー: '{macro_name}' は複数のモジュールにあります。--module で指定してください:")
            for comp in matches:
                print(f"  {comp.Name}")
            return False
        found.append((matches[0], macro_name))

    def _range_of(comp, macro_name):
        cm = comp.CodeModule
        start = cm.ProcStartLine(macro_name, 0)
        count = cm.ProcCountLines(macro_name, 0)
        # 領域に食い込んだ「次プロシージャの宣言行」を削除範囲から外す。
        # 外さないと 1行完結 Sub の直後のプロシージャが宣言を失って壊れる
        # （get / replace-procedure は同じ絞り込みを既に持っている）
        return _narrow_proc_range(cm, start, count)

    for comp, macro_name in found:
        start, count = _range_of(comp, macro_name)
        print(f"--- 削除するプロシージャ: [{comp.Name}] {macro_name} ({count}行) ---")
        print(comp.CodeModule.Lines(start, count).rstrip())
        print("-" * 40)
    label = found[0][1] if len(found) == 1 else f"{found[0][1]} ほか {len(found) - 1} 本"
    if not getattr(args, 'yes', False):
        ans = input(f"'{label}' を削除しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    tag = found[0][1] if len(found) == 1 else f"{found[0][1]}_ほか{len(found) - 1}本"
    if make_backup(wb.FullName, f"delete_{tag}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    for mod in dict.fromkeys(comp.Name for comp, _ in found):
        make_module_backup(wb, mod)

    for comp, macro_name in found:
        start, count = _range_of(comp, macro_name)      # 前の削除で行がずれるので、消す直前に位置を取り直す
        comp.CodeModule.DeleteLines(start, count)
    wb.Save()
    for comp, macro_name in found:
        print(f"削除完了: [{comp.Name}] '{macro_name}'")
    print("→ 保存しました" + (f"（{len(found)} 本・控えは 1 冊）" if len(found) > 1 else ""))
    return True


def cmd_delete_module(args):
    """モジュール丸ごと削除: delete-module [excel_file] <モジュール名> [-y]

    削除前に中身の要約を表示して確認。ブック＋モジュールの自動バックアップつき
    （戻すのは restore）。ThisWorkbook / シートモジュールは削除できない（VBAの制約）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: delete-module [excel_file] <モジュール名> [-y]")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: delete-module [excel_file] <モジュール名>'):
        return False

    xl, wb = get_workbook(target_file)
    comp = None
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            comp = c
            break
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が見つかりません")
        print("  存在するモジュール: " + ', '.join(c.Name for c in wb.VBProject.VBComponents))
        return False
    if int(comp.Type) == 100:
        print(f"エラー: '{comp.Name}' はブック/シートのモジュールなので削除できません（VBAの制約）")
        return False

    cm = comp.CodeModule
    n = cm.CountOfLines
    kind = {1: '標準モジュール', 2: 'クラス', 3: 'フォーム'}.get(int(comp.Type), '?')
    print(f"削除対象: [{comp.Name}]（{kind}・{n}行）")
    if n > 0:
        head = cm.Lines(1, min(n, 8)).replace('\r\n', '\n')
        print("  --- 先頭8行 ---")
        for ln in head.split('\n'):
            print(f"  {ln}")
        if n > 8:
            print(f"  … 他 {n - 8}行")

    if not getattr(args, 'yes', False):
        ans = input(f"モジュール '{comp.Name}' を丸ごと削除しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    if make_backup(wb.FullName, f"delmod_{comp.Name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    backup = make_module_backup(wb, comp.Name)

    actual_name = comp.Name
    wb.VBProject.VBComponents.Remove(comp)
    # VBE の Remove は遅延完了することがある（対象モジュールのコードが実行中など）。
    # 消える前に Save すると「削除完了」と報告しながらファイルには残り、
    # あとから遅延 Remove だけが効いて未保存のまま消える——という食い違いが起きる
    pythoncom.PumpWaitingMessages()
    if not _wait_component_gone(wb, actual_name):
        print(f"エラー: モジュール '{actual_name}' の Remove が完了しませんでした"
              f"（そのモジュールのコードが実行中の可能性があります）。", file=sys.stderr)
        print("  保存はしていません。Excel 側の実行中マクロ・開いているフォームを"
              "閉じてから、もう一度実行してください。", file=sys.stderr)
        if backup:
            print(f"  モジュールのバックアップ: {backup}", file=sys.stderr)
        return False

    wb.Save()
    print(f"削除完了: モジュール '{module_name}' → 保存しました")
    if backup:
        print(f"  戻すには: py vba_manager.py restore {os.path.basename(backup)}")
    return True


def _module_name_of_file(path):
    """.bas/.cls/.frm の Attribute VB_Name（読めなければファイル名の拡張子抜き）。"""
    p = smart_path_resolve(path) or path
    for enc in ('cp932', 'utf-8'):
        try:
            with open(p, encoding=enc) as f:
                for _ in range(40):
                    ln = f.readline()
                    if not ln:
                        break
                    m = re.match(r'\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', ln)
                    if m:
                        return m.group(1)
            break
        except (OSError, UnicodeDecodeError):
            continue
    return os.path.splitext(os.path.basename(path))[0]


def cmd_replace_module(args):
    """モジュール全体を Remove+Import で置換 (Attribute を正しく処理)"""
    target_file, rest = parse_target_and_rest(args.posargs)

    # .bas のパスだけなら、モジュール名はファイルの Attribute VB_Name（無ければファイル名）から取る
    # （2026-09-24・パスだけを渡して「使い方」が返った＝台帳の失敗の理由の 2 位）
    if len(rest) == 1 and rest[0].lower().endswith(('.bas', '.cls', '.frm')):
        rest = [_module_name_of_file(rest[0]), rest[0]]
    if len(rest) < 2:
        print("使い方: replace-module [excel_file] [<module_name>] <bas_file>（名前を省くとファイルの VB_Name）")
        return False
    module_name, code_file = rest[0], rest[1]

    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False

    # UTF-8化事故（CP932で書くべき .bas をUTF-8で保存）の水際チェック
    if not validate_bas_encoding(resolved):
        return False

    # フォーム（.frm）はレイアウトを .frx に持ち、Import は .frm と同名の .frx を
    # 同じ場所に要求する。.frx 無しで Import するとレイアウトが空のフォームになるため止める。
    is_form_file = resolved.lower().endswith('.frm')
    src_frx = os.path.splitext(resolved)[0] + '.frx'
    if is_form_file and not os.path.exists(src_frx):
        print(f"エラー: フォームの相方 {os.path.basename(src_frx)} が見つかりません。")
        print("  .frm と .frx は同じフォルダにペアで置いてください（.frx が無いとレイアウトが失われます）。")
        return False

    # 改行二重化（\r\r\n 化）の水際チェック＆修正。
    # 過去の二重化事故は、外部で作られた .bas が既に倍増した状態で replace-module に
    # 渡され、それを無検査で Import したのが原因だった。ここで修正してから取り込む。
    import_path = resolved
    tmp_norm = None
    fixed_bytes, raw_bytes, was_fixed = normalize_bas_newlines(resolved)
    if was_fixed:
        # VBA は \r\r\n を「行＋空行」と解釈して行数が倍に見える。実症状に合わせて
        # \r\n / \r / \n のいずれでも行が区切られる前提で数える。
        before_lines = len(re.split(r'\r\n|\r|\n', raw_bytes.decode('cp932')))
        after_lines = len(re.split(r'\r\n|\r|\n', fixed_bytes.decode('cp932')))
        print(f"⚠ 改行の二重化を検知しました。インポート前に修正します: {before_lines}行 → {after_lines}行")
        tmp_norm = os.path.join(SCRIPT_DIR, f"_norm_{os.path.basename(resolved)}")
        with open(tmp_norm, 'wb') as f:
            f.write(fixed_bytes)
        if is_form_file:
            # 正規化後の .frm から Import する場合も .frx を随伴させる
            # （コピーしないとレイアウト無しで取り込まれる穴だった）
            shutil.copy2(src_frx, os.path.splitext(tmp_norm)[0] + '.frx')
        import_path = tmp_norm

    # .bas の VB_Name と指定モジュール名の照合（別モジュール取り違えの防止）。
    # VB_Name が無い .bas は Import 時に Module1 等の別名で入り「Xを消してYが増える」
    # 事故になるため、ここで止める。
    with open(import_path, 'rb') as f:
        bas_head = f.read().decode('cp932', errors='replace')
    m_name = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', bas_head,
                       re.MULTILINE | re.IGNORECASE)
    if not m_name:
        print(f"エラー: {os.path.basename(resolved)} に Attribute VB_Name 行がありません。")
        print("  このまま Import すると別名モジュールとして取り込まれます。")
        print("  export-module で出力した .bas をベースに編集してください。")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False
    if m_name.group(1).lower() != module_name.lower():
        print(f"エラー: .bas の VB_Name '{m_name.group(1)}' が指定モジュール名 '{module_name}' と一致しません。")
        print("  別モジュールの .bas を取り込もうとしている可能性があります（対象取り違え防止のため停止）。")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False

    # プロシージャ名の識別子チェック（先頭 _ 等は Import 自体は通るがコンパイルで死ぬ）
    bad_names = _find_invalid_procedure_names(re.sub(r'\r\n|\r', '\n', bas_head))
    if bad_names and not getattr(args, 'force', False):
        print("エラー: VBA の識別子規則に反するプロシージャ名があります（--force で強行可）:")
        for ln, _nm, reason in bad_names:
            print(f"  行{ln}: {reason}")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False

    xl, wb = get_workbook(target_file)
    if make_backup(wb.FullName, f"module_{module_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        print("  ※ 未保存の新規ブックはバックアップできません。一度保存してから実行してください。")
        return False
    # モジュール単位のバックアップ（Import 失敗時の復旧素材を兼ねる）
    module_backup = make_module_backup(wb, module_name)
    print(f"モジュール '{module_name}' を Remove+Import で置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            xl.DisplayAlerts = False
            removed = False
            imported = False
            try:
                wb.Save()
                pythoncom.PumpWaitingMessages()
                wb.VBProject.VBComponents.Remove(comp)
                removed = True
                pythoncom.PumpWaitingMessages()
                # 実名検証つき Import（_wait_component_gone で動的消滅検知）
                _import_module_verified(wb, import_path, module_name)
                # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない
                imported = True
                _save_with_retry(wb)
            except ModuleNameCollisionError as ex:
                # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
                _print_collision_guidance(ex, module_name, module_backup)
                return False
            except Exception as ex:
                print(f"エラー: 置換中に失敗しました: {ex}")
                if imported:
                    # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
                    _print_save_failed_guidance(module_name)
                elif removed:
                    # Remove だけ成功して Import に失敗＝開いているブックからモジュール消失。
                    # 直前のモジュールバックアップから自動復旧を試みる。
                    print(f"⚠ モジュール '{module_name}' は Remove 済みです。バックアップから復旧を試みます...")
                    try:
                        if module_backup and os.path.exists(module_backup):
                            _import_module_verified(wb, module_backup, module_name)
                            print(f"復旧成功: {module_backup} を再インポートしました（置換前の内容に戻っています）")
                        else:
                            raise RuntimeError("モジュールバックアップがありません")
                    except ModuleNameCollisionError as ex2:
                        print(f"復旧の再 Import で名前衝突: {ex2}")
                        print(f"  置換前の内容は別名モジュール '{ex2.actual_name}' 側にあります。")
                        print(f"  旧 '{module_name}' が消えているのを確認してから '{ex2.actual_name}' を改名してください。")
                    except Exception as ex2:
                        print(f"復旧失敗: {ex2}")
                        print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。")
                        print("  対処: ブックを『保存せずに閉じて』開き直せば置換前の状態に戻ります。")
                        if module_backup:
                            print(f"  または backups のバックアップを手動で Import: {module_backup}")
                return False
            finally:
                xl.DisplayAlerts = True
                if tmp_norm and os.path.exists(tmp_norm):
                    _remove_export_artifacts(tmp_norm)
            print(f"置換完了: モジュール '{module_name}' → 保存しました")
            note_if_macro_free_book(wb)
            return True

    if tmp_norm and os.path.exists(tmp_norm):
        _remove_export_artifacts(tmp_norm)
    print(f"エラー: モジュール '{module_name}' が見つかりません")
    return False


def _parse_module_blocks(bas_text):
    """
    .bas モジュールの本文を ヘッダー / Sub・Functionブロック群 / 末尾 に分解する。
    各ブロックは前ブロックの End Sub/Function の次行から自分の End Sub/Function 行までを所有。
    Attribute 行（ショートカット定義）は Sub 内に含まれるので自動的にブロック内に入る。
    戻り値: (header_lines, blocks, trailing_lines)
        block: {'name': str, 'kind': 'Sub'|'Function', 'lines': [str,...]}
    """
    sub_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[^\s\(]+)',
        re.IGNORECASE
    )
    # 末尾コメント（End Sub 'xxx）を許容（cmd_replace_procedure の end_pattern と同じ理由）
    end_pattern = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\s*(?:'.*)?$", re.IGNORECASE)

    lines = bas_text.split('\r\n')
    n = len(lines)

    first_sub = None
    for i, line in enumerate(lines):
        if sub_pattern.match(line):
            first_sub = i
            break

    if first_sub is None:
        return lines, [], []

    header = lines[:first_sub]

    blocks = []
    cur_start = first_sub
    i = first_sub
    while i < n:
        m = sub_pattern.match(lines[i])
        if m:
            # 「Sub X(): 処理: End Sub」の1行完結は宣言行自身で閉じている。
            # 次の End 行まで探すと後続プロシージャを同一ブロックに巻き込むため、
            # その行（＋随伴 Attribute 行）でブロックを閉じる。
            # 文字列リテラルだけでなくコメントも落とす。落とさないと行末コメントの
            # 「End Sub」で誤判定し、ブロックが宣言行だけで閉じて本体が次の
            # プロシージャに吸収され、reorder で本体が隣へ道連れになる
            no_str = _strip_vba_comment(re.sub(r'"[^"]*"', '""', lines[i]))
            if re.search(r'\bEnd\s+(?:Sub|Function|Property)\b', no_str, re.IGNORECASE):
                j = i
                while j + 1 < n and lines[j + 1].strip().startswith('Attribute '):
                    j += 1
            else:
                j = i + 1
                while j < n and not end_pattern.match(lines[j]):
                    j += 1
                if j >= n:
                    break
            blocks.append({
                'name':  m.group('name'),
                'kind':  ' '.join(m.group('kind').split()).lower(),
                'lines': lines[cur_start:j + 1],
            })
            cur_start = j + 1
            i = j + 1
        else:
            i += 1

    trailing = lines[cur_start:]
    return header, blocks, trailing


def _write_module(header, blocks, trailing):
    parts = list(header)
    for b in blocks:
        parts.extend(b['lines'])
    parts.extend(trailing)
    return '\r\n'.join(parts)


def cmd_reorder_macro(args):
    """
    マクロを同モジュール内で 1 つ上 / 下のマクロと入れ替える。

    終了コード:
        0 : 成功
        1 : 引数エラー / その他
        2 : マクロが見つからない
        3 : 既にモジュール内の最初/最後（境界）
    """
    rest = list(args.posargs)
    if len(rest) < 2:
        print("使い方: reorder-macro <macro_name> <up|down|top|bottom|位置番号>")
        sys.exit(1)
    macro_name = rest[0]
    direction  = rest[1].lower()
    if direction not in ('up', 'down', 'top', 'bottom') and not direction.isdigit():
        print("方向は up|down|top|bottom または移動先の位置番号(1始まり)を指定してください")
        sys.exit(1)

    xl, wb = get_workbook(None)  # ActiveWorkbook 自動検出

    # 対象マクロを含む標準モジュールを特定
    # （replace-procedure / delete-procedure と同様、同名複数は取り違え防止で停止）
    matched_comps = []
    for comp in wb.VBProject.VBComponents:
        if comp.Type != 1:  # 標準モジュールのみ対象
            continue
        try:
            comp.CodeModule.ProcStartLine(macro_name, 0)
            matched_comps.append(comp)
        except Exception:
            continue

    if not matched_comps:
        print(f"エラー: マクロ '{macro_name}' が標準モジュールに見つかりません")
        sys.exit(2)
    if len(matched_comps) > 1:
        names = ', '.join(c.Name for c in matched_comps)
        print(f"エラー: 同名マクロ '{macro_name}' が複数のモジュールにあります: {names}")
        print("  対象を特定できないため中止しました（重量操作の取り違え防止）。")
        sys.exit(2)
    target_comp = matched_comps[0]

    module_name = target_comp.Name

    # モジュールをエクスポートして CP932 のまま読み込む
    tmp_bas = os.path.join(SCRIPT_DIR, f"_tmp_reorder_{module_name}.bas")
    target_comp.Export(tmp_bas)
    with open(tmp_bas, 'rb') as f:
        bas_text = f.read().decode('cp932')

    header, blocks, trailing = _parse_module_blocks(bas_text)

    # 対象ブロックの index（VBA の名前は大小無視なので比較も合わせる）
    target_idx = None
    for i, b in enumerate(blocks):
        if b['name'].lower() == macro_name.lower():
            target_idx = i
            break
    if target_idx is None:
        _remove_export_artifacts(tmp_bas)
        print(f"エラー: モジュール {module_name} に Sub '{macro_name}' が見つかりません")
        sys.exit(2)

    # 並べ替えの位置は Sub 単位で数える（Function は位置の数に入れない）
    visible_indices = [
        i for i, b in enumerate(blocks) if b['kind'] == 'sub'
    ]

    if target_idx not in visible_indices:
        _remove_export_artifacts(tmp_bas)
        print(f"エラー: '{macro_name}' は Sub ではないため並べ替えの対象外です")
        sys.exit(2)

    vis_pos = visible_indices.index(target_idx)

    if direction == 'up':
        if vis_pos == 0:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最初です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos - 1]
        # ブロック単位で入れ替え（Attribute 行は各ブロック内に含まれているので一緒に動く）
        blocks[target_idx], blocks[swap_block_idx] = blocks[swap_block_idx], blocks[target_idx]
    elif direction == 'down':
        if vis_pos == len(visible_indices) - 1:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最後です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos + 1]
        blocks[target_idx], blocks[swap_block_idx] = blocks[swap_block_idx], blocks[target_idx]
    else:
        # top / bottom / 位置番号: 一発で目的位置へ（up/down を N回＝重量処理N回、の代わり）
        if direction == 'top':
            new_pos = 0
        elif direction == 'bottom':
            new_pos = len(visible_indices) - 1
        else:
            new_pos = max(0, min(int(direction) - 1, len(visible_indices) - 1))
        if new_pos == vis_pos:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 既にその位置です")
            sys.exit(3)
        # 可視ブロックの並びだけを組み替え、非表示ブロックの位置は維持する
        vis_blocks = [blocks[i] for i in visible_indices]
        blk = vis_blocks.pop(vis_pos)
        vis_blocks.insert(new_pos, blk)
        for i, b in zip(visible_indices, vis_blocks):
            blocks[i] = b

    new_text = _write_module(header, blocks, trailing)
    with open(tmp_bas, 'wb') as f:
        f.write(new_text.encode('cp932'))

    # 破壊操作（Remove+Import）なので他コマンドと同じくバックアップ必須。
    # 取れなければ停止（--force で強行）
    if make_backup(wb.FullName, f"reorder_{macro_name}") is None and not getattr(args, 'force', False):
        _remove_export_artifacts(tmp_bas)
        print("エラー: バックアップが取れなかったため中止しました（--force で強行可能）")
        return False
    # モジュール単位のバックアップ（Import 失敗時の自動復旧素材。replace-module と同型）
    module_backup = make_module_backup(wb, module_name)
    print(f"並べ替え中: [{module_name}] '{macro_name}' を {direction}")

    # replace-module と同じ安定化手順（sleep + PumpWaitingMessages）に揃える
    xl.DisplayAlerts = False
    removed = False
    imported = False
    success = False
    try:
        wb.Save()
        pythoncom.PumpWaitingMessages()
        wb.VBProject.VBComponents.Remove(target_comp)
        removed = True
        pythoncom.PumpWaitingMessages()
        # 実名検証つき Import（_wait_component_gone で動的消滅検知）
        _import_module_verified(wb, tmp_bas, module_name)
        # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない。
        # success は最後の Save の後にしか立たないので、復旧判定には使えない
        imported = True
        _save_with_retry(wb)
        success = True
    except ModuleNameCollisionError as ex:
        # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
        _print_collision_guidance(ex, module_name, module_backup, err=True)
        return False
    except Exception as ex:
        # Remove 成功後に Import が失敗するとモジュールが消えたままになる。
        # replace-module と同じく、モジュールバックアップからの自動復旧を先に試みる
        if imported:
            # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
            print(f"エラー: 並べ替え後の保存に失敗しました: {ex}", file=sys.stderr)
            _print_save_failed_guidance(module_name, err=True)
        elif removed:
            print(f"エラー: 並べ替え中に失敗しました（モジュール '{module_name}' が"
                  f"開いているブックから外れた可能性があります）: {ex}", file=sys.stderr)
            try:
                if module_backup and os.path.exists(module_backup):
                    _import_module_verified(wb, module_backup, module_name)
                    print(f"復旧成功: {module_backup} を再インポートしました"
                          f"（並べ替え前の内容に戻っています）", file=sys.stderr)
                else:
                    raise RuntimeError("モジュールバックアップがありません")
            except ModuleNameCollisionError as ex2:
                print(f"復旧の再 Import で名前衝突: {ex2}", file=sys.stderr)
                print(f"  並べ替え前の内容は別名モジュール '{ex2.actual_name}' 側にあります。", file=sys.stderr)
                print(f"  旧 '{module_name}' が消えているのを確認してから"
                      f" '{ex2.actual_name}' を改名してください。", file=sys.stderr)
            except Exception as ex2:
                print(f"復旧失敗: {ex2}", file=sys.stderr)
                print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。", file=sys.stderr)
                print("  対処: ブックを『保存せずに閉じて』開き直せば並べ替え前の状態に戻ります。", file=sys.stderr)
                if module_backup:
                    print(f"  または並べ替え前のモジュールを restore: {module_backup}", file=sys.stderr)
                print(f"  並べ替え後のコードも残してあります（restore に渡せます）: {tmp_bas}", file=sys.stderr)
        else:
            print(f"エラー: 並べ替えに失敗しました: {ex}", file=sys.stderr)
        return False
    finally:
        xl.DisplayAlerts = True
        # 「Remove 済みで Import 失敗」のときだけ復旧用に残し、それ以外は掃除する。
        # Import 済み（保存だけ失敗）はブック側に新コードが入っているので残す必要はない
        if not (removed and not imported) and os.path.exists(tmp_bas):
            _remove_export_artifacts(tmp_bas)

    print(f"完了: [{module_name}] '{macro_name}' を {direction}に移動")
    sys.exit(0)


def cmd_export_module(args):
    """モジュールを .bas ファイルにエクスポート"""
    target_file, rest = parse_target_and_rest(args.posargs)

    if not rest:
        print("使い方: export-module [excel_file] <module_name>")
        return False
    module_name = rest[0]

    # 書き出すだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)
    if not target_file:
        wb = _book_holding_module(wb, module_name, include_addins=True) or wb

    # export-all と同じ Type 別拡張子を使う。フォームを .bas で書き出すと、
    # replace-module の「.frm なのに .frx が無い」ガードをすり抜けてしまう
    ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            # 表記ゆれ（大小文字）でファイル名が実モジュール名とズレないよう comp.Name を使う
            ext = ext_map.get(int(comp.Type), '.bas')
            out_path = os.path.join(SCRIPT_DIR, f"{comp.Name}{ext}")
            if os.path.exists(out_path):
                print(f"（既存の {os.path.basename(out_path)} を上書きします）")
            comp.Export(out_path)
            print(f"エクスポート完了: {out_path}")
            return True

    print(f"エラー: モジュール '{module_name}' が {wb.Name} に見つかりません")
    print("  存在するモジュール: " + ', '.join(c.Name for c in wb.VBProject.VBComponents))
    return False


def cmd_export_all(args):
    """全モジュールを一括エクスポート: export-all [excel_file] [--dir 出力先] [--check]

    1回のCOM接続で全 VBComponents を書き出す（1コマンドずつ回すと数分かかる
    ことが実測済みの作業を1コマンド化）。--check で書き出した .bas/.frm に
    check-bas 相当の機械検査（文字コード/改行二重化/重複）をその場でかける。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    history = getattr(args, 'history', False)

    ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}

    # 書き出すだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)
    if history:
        # --history（2026-09-16）: _exports/<ブック>/<日時>/ に書き出し、前回の書き出しと比べる（--dir はその親）
        stem = os.path.splitext(os.path.basename(str(wb.Name)))[0]
        parent = os.path.join(getattr(args, 'dir_opt', None) or os.path.join(SCRIPT_DIR, '_exports'), stem)
        out_dir = os.path.abspath(os.path.join(parent, time.strftime("%Y%m%d_%H%M%S")))
    else:
        out_dir = os.path.abspath(getattr(args, 'dir_opt', None) or SCRIPT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    comps = list(wb.VBProject.VBComponents)
    exported = []
    skipped = 0
    for k, comp in enumerate(comps, 1):
        ctype = int(comp.Type)
        if ctype not in ext_map:
            skipped += 1
            continue
        if ctype == 100 and comp.CodeModule.CountOfLines == 0:
            skipped += 1               # 空の ThisWorkbook / Sheet モジュールは省く
            continue
        cmd_progress_note(f"書き出し {k}/{len(comps)} {comp.Name}")
        out_path = os.path.join(out_dir, comp.Name + ext_map[ctype])
        comp.Export(out_path)
        exported.append(out_path)
        print(f"  {os.path.basename(out_path)}")
    print(f"エクスポート完了: {len(exported)}本 → {out_dir}"
          + (f"（空モジュール等 {skipped}本はスキップ）" if skipped else ""))
    if history:
        from vbam_devtools import _export_history_note
        _export_history_note(str(wb.Name), out_dir)

    if getattr(args, 'check', False):
        print("----- 取り込み前検査 (check-bas 相当) -----")
        ng = 0
        for p in exported:
            if not p.lower().endswith(('.bas', '.frm', '.cls')):
                continue
            ok = validate_bas_encoding(p)
            _, _, doubled = normalize_bas_newlines(p)
            with open(p, 'rb') as f:
                norm_text = re.sub(r'\r\n|\r', '\n', f.read().decode('cp932', errors='replace'))
            dups = _find_duplicate_procedures(norm_text)
            if ok and not doubled and not dups:
                print(f"  [OK] {os.path.basename(p)}")
            else:
                ng += 1
                marks = []
                if not ok:
                    marks.append("文字コード")
                if doubled:
                    marks.append("改行二重化")
                if dups:
                    marks.append(f"重複({', '.join(dups)})")
                print(f"  [NG] {os.path.basename(p)}: {' / '.join(marks)}")
        print(f"----- 検査結果: NG {ng} / {len(exported)}本 -----")
        return ng == 0
    return True


def cmd_list_backups(args):
    """バックアップの一覧: list-backups [キーワード]（COM不要）

    backups フォルダの内容を新しい順に表示。restore の対象選びに使う。
    """
    kw = args.posargs[0] if args.posargs else None
    if not os.path.isdir(BACKUP_DIR):
        print(f"バックアップフォルダがありません: {BACKUP_DIR}")
        return False
    entries = []
    for name in os.listdir(BACKUP_DIR):
        path = os.path.join(BACKUP_DIR, name)
        if not os.path.isfile(path):
            continue
        if kw and kw.lower() not in name.lower():
            continue
        entries.append((os.path.getmtime(path), name, os.path.getsize(path)))
    entries.sort(reverse=True)
    if not entries:
        print("該当するバックアップはありません。" + (f"（キーワード: {kw}）" if kw else ""))
        return True
    # `or 30` だと --max 0（件数だけ見たい）が偽値で既定に化ける（is None 判定にする）
    _m = getattr(args, 'max_hits', None)
    limit = 30 if _m is None else int(_m)
    if limit < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False
    print(f"--- バックアップ一覧（新しい順・{min(limit, len(entries))}/{len(entries)}件） ---")
    for mtime, name, size in entries[:limit]:
        stamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
        low = name.lower()
        if low.endswith(('.bas', '.frm', '.cls')):
            kind = "モジュール"
        elif low.endswith(('.xlsm', '.xlsx', '.xlsb', '.xls', '.xlam')):
            kind = "ブック"
        else:
            kind = "その他"
        print(f"  {stamp}  [{kind}] {name}  ({size:,} bytes)")
    if len(entries) > limit:
        print(f"  …他 {len(entries) - limit}件（--max で上限変更可）")
    print(f"場所: {BACKUP_DIR}")
    print("戻すには: py vba_manager.py restore <ファイル名>   （モジュール .bas/.frm のみ）")
    return True


def cmd_restore(args):
    """モジュールバックアップを開いているブックに書き戻す: restore <バックアップ.bas>

    対象モジュール名はファイル内の Attribute VB_Name から機械的に取得し、
    replace-module と同じ経路（照合・ガード・自動復旧つき）で適用する。
    どの世代に戻すかの判断はユーザー/AI側（list-backups で選ぶ）。
    ブック丸ごと（.xlsm）のバックアップはこのコマンドでは扱わない
    （開いているブックへの上書きになるため。必要ならExcelを閉じて手動コピー）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: restore [excel_file] <バックアップファイル名>   （list-backups で確認）")
        return False
    name = rest[0]
    if name.lower().endswith(('.xlsm', '.xlsx', '.xlsb', '.xls')):
        print("エラー: ブック丸ごとのバックアップは restore では扱いません。")
        print("  （開いているブックそのものの上書きになるため。Excelを閉じて手動でコピーしてください）")
        return False
    path = name if os.path.isabs(name) else os.path.join(BACKUP_DIR, name)
    if not os.path.exists(path):
        print(f"エラー: バックアップが見つかりません: {path}")
        print("  list-backups で名前を確認してください。")
        return False
    with open(path, 'rb') as f:
        head = f.read().decode('cp932', errors='replace')
    m = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', head, re.MULTILINE | re.IGNORECASE)
    if not m:
        print(f"エラー: {os.path.basename(path)} に VB_Name がありません（モジュールバックアップではない可能性）")
        return False
    module_name = m.group(1)
    print(f"復元: モジュール '{module_name}' ← backups/{os.path.basename(path)}")
    import argparse as _ap
    ns = _ap.Namespace(posargs=([target_file] if target_file else []) + [module_name, path],
                       force=getattr(args, 'force', False))
    return cmd_replace_module(ns)


def cmd_list_shortcuts(args):
    """ショートカットキーが設定されているマクロの一覧表示"""
    target_file, _ = parse_target_and_rest(args.posargs)
    # Export して読むだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)

    shortcuts = []
    unread = []      # Export に失敗して中身を見られなかったモジュール

    for comp in wb.VBProject.VBComponents:
        if comp.Type not in (1, 2, 3, 100):
            continue

        tmp_file = os.path.join(SCRIPT_DIR, f"_tmp_sc_{comp.Name}.bas")
        try:
            comp.Export(tmp_file)
            with open(tmp_file, 'rb') as f:
                content = f.read().decode('cp932', errors='replace')
        except Exception as ex:
            # 黙って飛ばすと最後に「ショートカットなし」と言い切ってしまう。
            # 読めなかったモジュールは必ず報告する
            unread.append((comp.Name, _com_error_text(ex)))
            continue
        finally:
            _remove_export_artifacts(tmp_file)

        # Attribute マクロ名.VB_ProcData.VB_Invoke_Func = "キー\n14"
        # 本物の Attribute 行は必ず行頭から始まる。行頭を見ないと、この形を例示している
        # コメント行（ショートカット取得 関数の説明等）まで拾って幽霊のキーを報告する
        pattern = re.compile(
            r'^Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^"\r\n]+)"',
            re.IGNORECASE | re.MULTILINE
        )
        for m in pattern.finditer(content):
            macro_name = m.group(1)
            raw_val = m.group(2)
            # 文字列としての "\n" や "\r" を本物の改行コードに変換してから分割
            raw_val_clean = raw_val.replace('\\n', '\n').replace('\\r', '\r')
            key_char = raw_val_clean.split('\n')[0].split('\r')[0]
            if not key_char:
                continue

            if len(key_char) == 1:
                if key_char.isupper():
                    shortcut_str = f"Ctrl + Shift + {key_char}"
                else:
                    shortcut_str = f"Ctrl + {key_char}"
            else:
                shortcut_str = f"Ctrl + {key_char}"

            shortcuts.append({
                'module': comp.Name,
                'macro': macro_name,
                'shortcut': shortcut_str
            })

    # 同じキーが 2 本以上: Excel は両方に残して片方しか動かさない（名前の順で先の方・2026-09-20 読者の報告）
    same = {}
    for item in shortcuts:
        same[item['shortcut']] = same.get(item['shortcut'], 0) + 1
    for item in shortcuts:
        item['same_key'] = same[item['shortcut']]

    if getattr(args, 'json', False):
        import json
        out = {"success": True, "file": wb.Name, "shortcuts": shortcuts}
        if unread:
            out["unread_modules"] = [{"module": m, "error": e} for m, e in unread]
        print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        return True

    if unread:
        print(f"⚠ 次のモジュールは読めませんでした（一覧に反映されていません・{len(unread)}件）:")
        for m, e in unread:
            print(f"  - {m}: {e}")

    if not shortcuts:
        if unread:
            print("読めたモジュールの範囲では、ショートカットキーの設定は見つかりませんでした。")
        else:
            print("ショートカットキーが設定されているマクロはありません。")
        return True

    print(f"設定されているショートカットキー一覧 (数: {len(shortcuts)})")
    print("-" * 60)
    for item in shortcuts:
        dup = f"   ⚠ 同じキーが {item['same_key']} 本（動くのは名前の順で先の方だけ）" if item['same_key'] > 1 else ""
        print(f"[{item['module']}] {item['macro']} -> {item['shortcut']}{dup}")
    print("-" * 60)
    if any(v > 1 for v in same.values()):
        print("同じキーを 1 本にするには: set-shortcut <動かしたいマクロ> <キー> -y（ほかの持ち主からは外します）")
    return True


def cmd_setup_check(args):
    """導入セルフ診断: setup-check

    「会話するだけでマクロが直る」環境に必要なものが揃っているかを○×で表示する。
    初心者が最初に打つ1コマンド。Excel を起動していなくても動く
    （VBOM 信頼設定のチェックだけは Excel 起動中に実施）。
    """
    results = []          # (ok: bool|None, 項目, 詳細, 対処)

    # 1. Python 本体
    v = sys.version_info
    bits = 64 if sys.maxsize > 2 ** 32 else 32
    results.append((True, "Python",
                    f"{v.major}.{v.minor}.{v.micro} ({bits}bit)", None))

    # 2. pywin32
    try:
        import win32com  # noqa: F401  （先頭 import 済みだが診断として明示確認）
        try:
            from importlib.metadata import version as _ver
            pv = _ver("pywin32")
        except Exception:
            pv = "(バージョン不明)"
        results.append((True, "pywin32", f"インストール済み {pv}", None))
    except ImportError:
        results.append((False, "pywin32", "見つかりません",
                        "py -m pip install pywin32 を実行してください（AI に「入れて」でも可）"))

    # 3. Excel のインストール（レジストリ確認・起動はしない）
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as k:
            curver = winreg.QueryValueEx(k, None)[0]     # 例: Excel.Application.16
        results.append((True, "Excel", f"インストール済み ({curver})", None))
        excel_installed = True
    except Exception:
        results.append((False, "Excel", "インストールが確認できません",
                        "Excel (デスクトップ版) が必要です"))
        excel_installed = False

    # 4. Excel の起動状態と VBOM（VBAプロジェクトへのアクセス信頼）
    xl = None
    if excel_installed:
        try:
            xl = _get_active_excel()
        except Exception:
            xl = None
    if xl is None:
        results.append((None, "Excel起動", "起動していません",
                        "VBOM 設定の診断には、Excel でブックを開いてから再実行してください"))
    else:
        try:
            wb_names = [w.Name for w in xl.Workbooks]
        except Exception:
            wb_names = []
        results.append((True, "Excel起動",
                        f"起動中（開いているブック: {', '.join(wb_names) or 'なし'}）", None))
        # VBOM: VBE にアクセスできるか（ブロックされていると例外になる）
        try:
            _ = xl.VBE.VBProjects.Count
            results.append((True, "VBOM信頼設定", "有効（VBAプロジェクトにアクセス可能）", None))
        except Exception:
            results.append((False, "VBOM信頼設定", "無効（VBAプロジェクトにアクセスできません）",
                            "Excel の [ファイル > オプション > トラストセンター > トラストセンターの設定 > "
                            "マクロの設定] で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」に"
                            "チェックを入れてください"))

    # 5. gen_py キャッシュの健全性（破損すると「Excelは起動していません」と誤報する既知問題）
    if xl is not None:
        try:
            win32com.client.GetActiveObject("Excel.Application")
            results.append((True, "COMキャッシュ(gen_py)", "正常", None))
        except Exception:
            results.append((False, "COMキャッシュ(gen_py)", "破損の疑い（キャッシュ経由の接続に失敗）",
                            r"%LOCALAPPDATA%\Temp\gen_py フォルダを削除すると自動再生成されます"))

    # 6. ツール一式の存在
    missing = [f for f in ("form_builder.py", "form_inspect.py",
                           "form_layout.py", "form_tool.py")
               if not os.path.exists(os.path.join(SCRIPT_DIR, f))]
    if missing:
        results.append((None, "ツール一式", f"見つからないファイル: {', '.join(missing)}",
                        "フォーム機能を使う場合は同じフォルダに配置してください（マクロ管理だけなら不要）"))
    else:
        results.append((True, "ツール一式", "vba_manager + フォーム4ツールが揃っています", None))

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": all(r[0] is not False for r in results),
                          "checks": [{"ok": r[0], "item": r[1], "detail": r[2],
                                      "fix": r[3]} for r in results]},
                         ensure_ascii=False), file=sys.stdout)
        return all(r[0] is not False for r in results)

    print("===== 導入セルフ診断 (setup-check) =====")
    ng = 0
    for ok, item, detail, fix in results:
        mark = "OK" if ok else ("--" if ok is None else "NG")
        if ok is False:
            ng += 1
        print(f"  [{mark}] {item}: {detail}")
        if fix and ok is not True:
            print(f"       → {fix}")
    print("-" * 44)
    if ng == 0:
        print("  問題なし。`py vba_manager.py list` から始められます。")
    else:
        print(f"  NG {ng}件。上の対処を実行してから再実行してください。")
        print("  分からないところは、導入しようとしている AI にこの出力を貼って聞いてください。")
    return ng == 0


def _collect_book_inventory(xl, wb, include_vba=True, quiet=False):
    """ブックの棚卸しデータを機械収集する（docs / call-graph の共通土台）。

    include_vba=False は VBA プロジェクトに触らない縮退モード
    （パスワード保護・VBOM 未信頼のブックをシート側だけ診るために使う）。
    """
    inv = {'name': wb.Name, 'fullname': wb.FullName}

    # シート
    sheets = []
    try:
        active = wb.ActiveSheet.Name
    except Exception:
        active = None
    for sh in wb.Sheets:
        d = {'name': sh.Name, 'active': sh.Name == active}
        try:
            d['visible'] = int(sh.Visible)
        except Exception:
            d['visible'] = -1
        try:
            ur = sh.UsedRange
            d['rows'] = ur.Rows.Count
            d['cols'] = ur.Columns.Count
            d['address'] = ur.Address
        except Exception:
            d['rows'] = d['cols'] = 0
            d['address'] = None
        sheets.append(d)
    inv['sheets'] = sheets

    # モジュールとプロシージャ（コード全文も保持＝call-graph が使う）
    proc_pat = re.compile(
        r'^\s*(?P<vis>Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function)\s+(?P<name>[^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    type_names = {1: '標準', 2: 'クラス', 3: 'フォーム', 100: 'ブック/シート'}
    modules = []
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        cm = comp.CodeModule
        n = cm.CountOfLines
        code = cm.Lines(1, n) if n > 0 else ""
        lines = code.split('\r\n') if code else []
        procs = []
        for m in proc_pat.finditer(code):
            name = m.group('name')
            info = {'name': name,
                    'kind': m.group('kind').capitalize(),
                    'private': bool(m.group('vis') and 'private' in m.group('vis').lower())}
            try:
                info['lines'] = cm.ProcCountLines(name, 0)
                body = cm.ProcBodyLine(name, 0)
                if body < len(lines):
                    first = lines[body].strip()
                    if first.startswith("'"):
                        info['comment'] = first.lstrip("'").strip()
            except Exception:
                pass
            procs.append(info)
        modules.append({'name': comp.Name, 'type': int(comp.Type),
                        'type_name': type_names.get(int(comp.Type), str(comp.Type)),
                        'total_lines': n, 'procs': procs, 'code': code})
    inv['modules'] = modules

    # フォーム（コントロール数）
    forms = []
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        if int(comp.Type) != 3:
            continue
        d = {'name': comp.Name}
        try:
            d['caption'] = comp.Properties("Caption").Value
        except Exception:
            d['caption'] = None
        try:
            d['controls'] = comp.Designer.Controls.Count
        except Exception:
            d['controls'] = None
        forms.append(d)
    inv['forms'] = forms

    # ショートカット（Attribute 走査。エクスポート方式は list-shortcuts と同じ）
    shortcuts = {}
    sc_unread = []      # Export に失敗して読めなかったモジュール（黙って落とさない）
    attr_pat = re.compile(
        r'Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^"]+)"',
        re.IGNORECASE | re.DOTALL)
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        if int(comp.Type) not in (1, 2, 3, 100):
            continue
        tmp = os.path.join(SCRIPT_DIR, f"_tmp_doc_{comp.Name}.bas")
        try:
            comp.Export(tmp)
            with open(tmp, 'rb') as f:
                content = f.read().decode('cp932', errors='replace')
            for m in attr_pat.finditer(content):
                raw = m.group(2).replace('\\n', '\n').replace('\\r', '\r')
                # 鍵を外した後の属性は `" \n14"`（空白 1 つ）＝ショートカット無し。strip しないと「Ctrl+ 」が出る（2026-09-17 repair で実測）
                key = raw.split('\n')[0].split('\r')[0].strip()
                if key:
                    # VB_Invoke_Func のキーが大文字なら Ctrl+Shift+キー の割当
                    if key.isalpha() and key == key.upper():
                        shortcuts[m.group(1)] = f"Ctrl+Shift+{key}"
                    else:
                        shortcuts[m.group(1)] = f"Ctrl+{key}"
        except Exception as ex:
            # 握りつぶすと「ショートカットなし」と誤って断言することになる
            sc_unread.append((comp.Name, _com_error_text(ex)))
        finally:
            _remove_export_artifacts(tmp)
    inv['shortcuts'] = shortcuts
    inv['shortcuts_unread'] = sc_unread
    if sc_unread and not quiet:
        # MCP では stderr と stdout が同じバッファなので、--json 時にここへ書くと
        # JSON の前にゴミが混ざって呼び出し側の json.loads が落ちる。
        # --json の呼び出し元は quiet=True にし、内容は inv['shortcuts_unread'] で受け取る
        print(f"⚠ ショートカット走査で読めなかったモジュール（{len(sc_unread)}件）: "
              + '、'.join(f"{m}({e})" for m, e in sc_unread), file=sys.stderr)

    # 図形・フォームコントロールに登録されたマクロ（OnAction）＝ボタンからの実行入口。
    # VBA ロック中でも読める（Shapes はシート側の情報）
    onaction = []
    for sh in wb.Worksheets:
        try:
            shapes = list(sh.Shapes)
        except Exception:
            continue
        for shp in shapes:
            try:
                oa = shp.OnAction
            except Exception:
                continue
            if oa:
                macro = oa.split('!')[-1].strip("'\" ")
                try:
                    shp_name = shp.Name
                except Exception:
                    shp_name = '(図形)'
                onaction.append((sh.Name, shp_name, macro))
    inv['onaction'] = onaction

    # テーブル・ピボット
    tables, pivots = [], []
    for sh in wb.Worksheets:
        try:
            for lo in sh.ListObjects:
                tables.append({'sheet': sh.Name, 'name': lo.Name,
                               'address': lo.Range.Address})
        except Exception:
            pass
        try:
            for pt in sh.PivotTables():
                pivots.append({'sheet': sh.Name, 'name': pt.Name})
        except Exception:
            pass
    inv['tables'] = tables
    inv['pivots'] = pivots

    # PowerQuery・接続・名前付き範囲
    queries = []
    try:
        for q in wb.Queries:
            queries.append({'name': q.Name})
    except Exception:
        pass
    inv['queries'] = queries
    conns = []
    try:
        for cn in wb.Connections:
            conns.append({'name': cn.Name})
    except Exception:
        pass
    inv['connections'] = conns
    names = []
    try:
        for nm in wb.Names:
            try:
                names.append({'name': nm.Name, 'refers_to': nm.RefersTo})
            except Exception:
                names.append({'name': nm.Name, 'refers_to': None})
    except Exception:
        pass
    inv['names'] = names
    return inv


def _inventory_or_explain(xl, wb):
    """棚卸しを試み、VBA に触れないブック（パスワード保護/VBOM未信頼）なら
    生の COM エラーで転ばず理由を説明して None を返す（docs/call-graph/impact 用）"""
    try:
        return _collect_book_inventory(xl, wb)
    except Exception as e:
        print("エラー: VBA プロジェクトに触れません（パスワード保護または VBOM 未信頼）。")
        print("  シート側だけの診断なら checkup（健康診断）が縮退モードで実行できます。")
        print(f"  詳細: {e}")
        return None


def cmd_docs(args):
    """ブックの取扱説明書を自動生成: docs [excel_file] [--out f.md]

    シート構成・モジュール別マクロ表（行数/ショートカット/先頭コメント）・
    フォーム・テーブル/ピボット/クエリ/接続/名前付き範囲を Markdown 1枚に棚卸しする。
    「このブックに何が入っているか」を機械が書く＝ブックと会話するための自己紹介文。
    """
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False

    if getattr(args, 'json', False):
        import json
        slim = {k: v for k, v in inv.items()}
        slim['modules'] = [{k: v for k, v in m.items() if k != 'code'}
                           for m in inv['modules']]
        print(json.dumps({"success": True, **slim}, ensure_ascii=False), file=sys.stdout)
        return True

    L = []
    total_procs = sum(len(m['procs']) for m in inv['modules'])
    total_lines = sum(m['total_lines'] for m in inv['modules'])
    L.append(f"# {inv['name']} の構成ドキュメント")
    L.append("")
    L.append(f"- 生成: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager docs）")
    L.append(f"- パス: {inv['fullname']}")
    L.append(f"- シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
             f"フォーム {len(inv['forms'])} / テーブル {len(inv['tables'])} / "
             f"ピボット {len(inv['pivots'])} / クエリ {len(inv['queries'])}")
    L.append("")

    L.append("## シート")
    L.append("")
    L.append("| シート | 使用範囲 | 大きさ | 状態 |")
    L.append("|---|---|---|---|")
    vis_label = {-1: '', 0: '非表示', 2: '完全非表示'}
    for s in inv['sheets']:
        mark = '（アクティブ）' if s['active'] else ''
        L.append(f"| {s['name']}{mark} | {s['address'] or '-'} | "
                 f"{s['rows']}行×{s['cols']}列 | {vis_label.get(s['visible'], '')} |")
    L.append("")

    # --preview N: 各シートの先頭N行を Markdown 表で（初見ブックの中身の見取り）
    try:
        preview = int(getattr(args, 'preview', None) or 0)
    except (TypeError, ValueError):
        preview = 0
    if preview > 0:
        def _md_cell(v):
            return _cell_str(v).replace('|', '\\|').replace('\n', ' ')
        for sh in wb.Sheets:
            try:
                ur = sh.UsedRange
                nrows = min(preview, ur.Rows.Count)
                ncols = min(ur.Columns.Count, 12)   # 横に広すぎる表は12列で切る
                head = sh.Range(ur.Cells(1, 1), ur.Cells(nrows, ncols))
                rows_v = _range_values_2d(head)
            except Exception:
                continue
            L.append(f"### {sh.Name} の先頭{nrows}行")
            L.append("")
            start_col = ur.Column
            headers = [_col_letter(start_col + j) for j in range(ncols)]
            L.append("| 行 | " + " | ".join(headers) + " |")
            L.append("|---|" + "---|" * ncols)
            for ri, r in enumerate(rows_v):
                cells = [_md_cell(v) for v in r] + [''] * (ncols - len(r))
                L.append(f"| {ur.Row + ri} | " + " | ".join(cells[:ncols]) + " |")
            if ur.Columns.Count > ncols:
                L.append(f"（横は {ncols} 列まで表示・実際は {ur.Columns.Count} 列）")
            L.append("")

    L.append("## マクロ")
    L.append("")
    for m in inv['modules']:
        if not m['procs'] and m['total_lines'] == 0:
            continue
        L.append(f"### [{m['name']}]（{m['type_name']}・{len(m['procs'])}プロシージャ・{m['total_lines']}行）")
        if m['procs']:
            L.append("")
            L.append("| プロシージャ | 種別 | 行数 | ショートカット | 説明（先頭コメント） |")
            L.append("|---|---|---|---|---|")
            for p in m['procs']:
                sc = inv['shortcuts'].get(p['name'], '')
                priv = 'Private ' if p.get('private') else ''
                L.append(f"| {p['name']} | {priv}{p['kind']} | {p.get('lines', '')} | "
                         f"{sc} | {p.get('comment', '')} |")
        L.append("")

    if inv['forms']:
        L.append("## フォーム")
        L.append("")
        L.append("| フォーム | キャプション | コントロール数 |")
        L.append("|---|---|---|")
        for f in inv['forms']:
            L.append(f"| {f['name']} | {f.get('caption') or ''} | {f.get('controls') or ''} |")
        L.append("")

    def _simple_list(title, items, fmt):
        if not items:
            return
        L.append(f"## {title}")
        L.append("")
        for it in items:
            L.append(f"- {fmt(it)}")
        L.append("")

    _simple_list("テーブル", inv['tables'],
                 lambda t: f"{t['name']}（{t['sheet']} {t['address']}）")
    _simple_list("ピボットテーブル", inv['pivots'],
                 lambda t: f"{t['name']}（{t['sheet']}）")
    _simple_list("PowerQuery", inv['queries'], lambda t: t['name'])
    _simple_list("接続", inv['connections'], lambda t: t['name'])
    _simple_list("名前付き範囲", inv['names'],
                 lambda t: f"{t['name']} → {t.get('refers_to') or '?'}")

    out_path = getattr(args, 'out_opt', None)
    out_path = os.path.abspath(out_path) if out_path else os.path.join(SCRIPT_DIR, "_last_docs.md")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print(f"構成ドキュメントを生成: {out_path}")
    print(f"  シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
          f"フォーム {len(inv['forms'])} / テーブル {len(inv['tables'])} / "
          f"ピボット {len(inv['pivots'])} / クエリ {len(inv['queries'])}")
    return True


def _analyze_calls(inv):
    """呼び出し関係の解析本体（call-graph / checkup 共用・COM 不要の純粋処理）。

    戻り値: {'known', 'edges', 'unresolved', 'orphans'}
    """
    known = {}
    form_modules = {m['name'] for m in inv['modules'] if m['type'] == 3}
    doc_modules = {m['name'] for m in inv['modules'] if m['type'] == 100}
    for m in inv['modules']:
        for p in m['procs']:
            known.setdefault(p['name'].lower(), (p['name'], m['name']))

    defn_pat = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(\)]+)', re.IGNORECASE)
    end_pat = re.compile(r'^\s*End\s+(?:Sub|Function)\b', re.IGNORECASE)
    call_pat = re.compile(r'\bCall\s+([^\s\(\):=]+)', re.IGNORECASE)
    # Declare 宣言された外部 API（Win32 等）は実在する呼び先＝未解決扱いにしない
    decl_pat = re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(]+)', re.IGNORECASE)
    api_names = set()
    for m in inv['modules']:
        for line in m['code'].split('\r\n'):
            dm = decl_pat.match(line)
            if dm:
                api_names.add(dm.group(1).lower())
    # Run の呼び先が丸ごと1つの文字列リテラルのときだけ静的に解決する。
    # 閉じクォート直後に & が続く（"…" & 変数）は動的呼び出し＝対象外
    any_run_pat = re.compile(r'Application\s*\.\s*Run\b', re.IGNORECASE)
    run_pat = re.compile(r'Application\s*\.\s*Run\s*\(?\s*"(?:[^"!]*!)?([^"]+)"(?!\s*&)',
                         re.IGNORECASE)
    # 裸呼び用: 全既知名の一括照合（ドット直後=他オブジェクトのメソッドは除外）
    if known:
        names_alt = '|'.join(sorted((re.escape(v[0]) for v in known.values()),
                                    key=len, reverse=True))
        # VBA の識別子は大文字小文字を区別しないため IGNORECASE で照合し、
        # マッチ後に known で正式名へ正規化する
        bare_pat = re.compile(r'(?<![\w.])(' + names_alt + r')(?![\w])', re.IGNORECASE)
    else:
        bare_pat = None

    edges = {}         # (caller_mod, caller_proc) -> {callee正式名, ...}
    unresolved = []    # (mod, proc, 呼び名, 行番号, 行テキスト)
    dynamic_runs = []  # (mod, proc, 行番号, 行テキスト) 呼び先が実行時に決まる Run
    for m in inv['modules']:
        cur = None
        for i, raw in enumerate(m['code'].split('\r\n'), 1):
            close_after = False
            scan = raw            # 走査対象（表示は raw のまま＝元の行を見せる）
            dm = defn_pat.match(raw)
            if dm:
                cur = dm.group(1)
                # 1行完結 Sub「Sub X(): Call Main: End Sub」は宣言行に本体が同居する。
                # ここで無条件に continue すると Call が走査されず、呼ばれている Main が
                # 「どこからも呼ばれていない」に誤って載り、誤記（Call 印刷実効）も
                # 未解決Callとして検出されない（call-graph の検出器がこの形だけ素通り）。
                inline = _inline_body_after_decl(raw, dm.end())
                if not inline.strip():
                    continue
                scan = inline
                # 同じ行で End Sub まで来ているなら、この行を見終えた時点で閉じる
                close_after = bool(re.search(
                    r'\bEnd\s+(?:Sub|Function|Property)\b',
                    _strip_vba_comment(inline), re.IGNORECASE))
            elif end_pat.match(raw):
                cur = None
                continue
            # 文字列リテラルを潰してからコメントを落とす（' の誤爆防止）
            line = re.sub(r'"[^"]*"', '""', scan).split("'")[0]
            # Rem コメント行の Call/Run を生きた呼び出し扱いしない
            low = line.strip().lower()
            if low == 'rem' or low.startswith('rem '):
                continue
            caller = (m['name'], cur or '(宣言部)')

            for cm_ in call_pat.finditer(line):
                name = cm_.group(1)
                if '.' in name:
                    # Call obj.Method(...) はオブジェクトのメソッド。ただし
                    # Call モジュール名.マクロ名 の修飾呼びは末尾で解決を試みる
                    tail = name.rsplit('.', 1)[-1]
                    hit = known.get(tail.lower())
                    if hit:
                        edges.setdefault(caller, set()).add(hit[0])
                    continue
                hit = known.get(name.lower())
                if hit:
                    edges.setdefault(caller, set()).add(hit[0])
                elif name.lower() not in api_names:
                    unresolved.append((m['name'], cur, name, i, raw.strip()))
            ran_static = False
            # Run の名前は文字列内にあるのでコメントだけ除いた原文から拾う
            # （生 raw だとコメントアウトされた Run を誤検知する）
            for rm_ in run_pat.finditer(_strip_vba_comment(scan)):
                ran_static = True
                name = rm_.group(1)
                hit = known.get(name.lower())
                if hit is None and '.' in name:
                    # Run "モジュール名.マクロ名" のモジュール修飾は
                    # Call 側の修飾呼びと同じく末尾名で解決を試みる
                    hit = known.get(name.rsplit('.', 1)[-1].lower())
                if hit:
                    edges.setdefault(caller, set()).add(hit[0])
                else:
                    unresolved.append((m['name'], cur, f'Run "{name}"', i, raw.strip()))
            if not ran_static and any_run_pat.search(line):
                # リテラル1本で解決できない Run（"…" & 変数 / 変数のみ）＝動的呼び出し
                dynamic_runs.append((m['name'], cur or '(宣言部)', i, raw.strip()[:60]))
            if bare_pat:
                for bm in bare_pat.finditer(line):
                    hit = known.get(bm.group(1).lower())
                    name = hit[0] if hit else bm.group(1)   # 正式名へ正規化
                    if cur and name.lower() == cur.lower():
                        continue          # 自分自身の再帰は流れの把握には不要
                    edges.setdefault(caller, set()).add(name)
            if close_after:
                # 1行完結 Sub は宣言行で閉じている。ここで戻さないと、後続行の
                # 呼び出しがこの Sub の名で計上され続ける
                cur = None

    # 図形・ボタンに登録されたマクロ（OnAction）＝ボタンからの実行入口
    onaction = {}
    for sheet, shp, macro in inv.get('onaction', ()):
        hit = known.get(macro.lower())
        if hit is None and '.' in macro:
            # 「モジュール名.マクロ名」形式の登録（同名マクロがあると Excel が
            # 自動でこの形式にする）も末尾名で解決する
            hit = known.get(macro.rsplit('.', 1)[-1].lower())
        if hit:
            onaction.setdefault(hit[0], []).append(f"{sheet}/{shp}")

    # 呼ばれる側の集合と孤立
    called = set()
    for callees in edges.values():
        called |= callees
    orphans = []
    for m in inv['modules']:
        for p in m['procs']:
            if p['name'] in called:
                continue
            if p['name'] in onaction:
                continue      # シート上のボタン/図形から呼ばれている＝孤立ではない
            # フォーム/ブック/シートのイベントプロシージャ、Private はイベント・内部用が
            # 多いので孤立には数えない（機械的な絞り込み）
            if m['name'] in form_modules or m['name'] in doc_modules:
                continue
            orphans.append((m['name'], p['name']))
    return {'known': known, 'edges': edges, 'onaction': onaction,
            'unresolved': unresolved, 'orphans': orphans,
            'dynamic_runs': dynamic_runs}


def _strip_vba_comment(raw):
    """行からコメント部を落とす（文字列リテラル内の ' は誤爆させない）"""
    in_str = False
    for i, ch in enumerate(raw):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return raw[:i]
    return raw


_AUTO_EXEC_STD = {'auto_open', 'auto_close'}
_AUTO_EXEC_PREFIXES = ('workbook_', 'worksheet_', 'chart_')


def _extra_code_scans(inv):
    """checkup の参考所見スキャン（COM 不要の純粋処理・事実の列挙のみ）。

    自動実行イベント / Option Explicit なし / On Error Resume Next /
    ハードコードされたパス / 長いプロシージャ / 破壊的な操作の所在 /
    ScreenUpdating・Calculation の戻し忘れ。いずれも要不要の判断はしない。
    """
    res = {'auto_exec': [], 'no_option_explicit': [], 'error_resume': [],
           'hardcoded_paths': [], 'long_procs': [], 'destructive': [],
           'no_restore': []}
    path_pat = re.compile(r'"((?:[A-Za-z]:\\|\\\\)[^"]{2,})"')
    oern_pat = re.compile(r'\bOn\s+Error\s+Resume\s+Next\b', re.IGNORECASE)
    defn_pat = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(\)]+)', re.IGNORECASE)
    end_pat = re.compile(r'^\s*End\s+(?:Sub|Function)\b', re.IGNORECASE)
    destr_pats = [
        (re.compile(r'(?<![\w.])(?:Kill|RmDir)\b', re.IGNORECASE), 'ファイル/フォルダ削除'),
        (re.compile(r'\.(?:DeleteFile|DeleteFolder|MoveFile|MoveFolder)\b',
                    re.IGNORECASE), 'FSOのファイル操作'),
        (re.compile(r'\b(?:Worksheets|Sheets)\s*\([^)]*\)\s*\.Delete\b'
                    r'|\bActiveSheet\s*\.Delete\b', re.IGNORECASE), 'シート削除'),
        (re.compile(r'(?:\bRows\b|\bColumns\b|\.EntireRow|\.EntireColumn)'
                    r'[^\n]*\.Delete\b', re.IGNORECASE), '行/列の削除'),
    ]
    su_off = re.compile(r'\bScreenUpdating\s*=\s*False\b', re.IGNORECASE)
    su_on = re.compile(r'\bScreenUpdating\s*=\s*True\b', re.IGNORECASE)
    ca_off = re.compile(r'\bCalculation\s*=\s*xl(?:Calculation)?Manual\b', re.IGNORECASE)
    ca_on = re.compile(r'\bCalculation\s*=\s*xl(?:Calculation)?Automatic\b', re.IGNORECASE)
    ee_off = re.compile(r'\bEnableEvents\s*=\s*False\b', re.IGNORECASE)
    ee_on = re.compile(r'\bEnableEvents\s*=\s*True\b', re.IGNORECASE)

    for m in inv['modules']:
        code = m['code']
        if not code:
            continue
        lines = code.split('\r\n')
        if not any(re.match(r'\s*Option\s+Explicit\b', ln, re.IGNORECASE)
                   for ln in lines):
            res['no_option_explicit'].append(m['name'])
        for p in m['procs']:
            nl = p['name'].lower()
            if (m['type'] == 100 and nl.startswith(_AUTO_EXEC_PREFIXES)) or \
               (m['type'] == 1 and nl in _AUTO_EXEC_STD):
                res['auto_exec'].append((m['name'], p['name'], p.get('lines')))
            if p.get('lines') and p['lines'] >= 150:
                res['long_procs'].append((m['name'], p['name'], p['lines']))

        cur = None
        state = {'su_off': False, 'su_on': False, 'ca_off': False, 'ca_on': False,
                 'ee_off': False, 'ee_on': False}

        def flush(proc, _m=m, _state=state):
            # プロシージャ末尾で ScreenUpdating/Calculation/EnableEvents の戻し忘れを確定する
            if proc:
                if _state['su_off'] and not _state['su_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'ScreenUpdating を False にしたまま True に戻す行がない'))
                if _state['ca_off'] and not _state['ca_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'Calculation を手動にしたまま自動に戻す行がない'))
                if _state['ee_off'] and not _state['ee_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'EnableEvents を False にしたまま True に戻す行がない'
                                           '（イベントが死んだままになる）'))
            for k in _state:
                _state[k] = False

        for i, raw in enumerate(lines, 1):
            dm = defn_pat.match(raw)
            if dm:
                flush(cur)
                cur = dm.group(1)
            elif end_pat.match(raw):
                flush(cur)
                cur = None
            body = _strip_vba_comment(raw)
            blanked = re.sub(r'"[^"]*"', '""', body)
            for pm in path_pat.finditer(body):
                res['hardcoded_paths'].append((m['name'], cur or '(宣言部)', i,
                                               pm.group(1)))
            if oern_pat.search(blanked):
                res['error_resume'].append((m['name'], cur or '(宣言部)', i))
            for pat, label in destr_pats:
                if pat.search(blanked):
                    res['destructive'].append((m['name'], cur or '(宣言部)', i,
                                               label, body.strip()[:60]))
            if su_off.search(blanked):
                state['su_off'] = True
            if su_on.search(blanked):
                state['su_on'] = True
            if ca_off.search(blanked):
                state['ca_off'] = True
            if ca_on.search(blanked):
                state['ca_on'] = True
            if ee_off.search(blanked):
                state['ee_off'] = True
            if ee_on.search(blanked):
                state['ee_on'] = True
        flush(cur)
    res['long_procs'].sort(key=lambda t: -t[2])
    return res


CHECKUP_HISTORY_DIR = os.path.join(SCRIPT_DIR, "_checkup_history")


def _checkup_history_path(book_name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', book_name)
    return os.path.join(CHECKUP_HISTORY_DIR, safe + ".json")


def _load_checkup_history(book_name):
    """ブック別の診断履歴（新しい順でなく古い順のリスト）を読む。無ければ空"""
    import json
    try:
        with open(_checkup_history_path(book_name), 'r', encoding='utf-8') as f:
            hist = json.load(f)
        if not isinstance(hist, list):
            return []
        # 要素の型まで守る（壊れた履歴ファイルで毎回診断が落ちるのを防ぐ）
        return [h for h in hist if isinstance(h, dict)]
    except (OSError, ValueError):
        return []


def _save_checkup_history(book_name, hist, keep=12):
    import json
    os.makedirs(CHECKUP_HISTORY_DIR, exist_ok=True)
    with open(_checkup_history_path(book_name), 'w', encoding='utf-8') as f:
        json.dump(hist[-keep:], f, ensure_ascii=False, indent=1)


# --- 確認済み（意図的）所見の記録 ---
# 「フォント混在」「同じ行のボタン幅」等、目で見て意図的なデザインと判断した所見は
# ここに登録すると次回以降の「所見サマリ」の件数・総合判定から除外される
# （消すのではなく「## 確認済み（意図的・件数から除外）」セクションに移すだけ＝事実は残す）。
CHECKUP_ACK_DIR = os.path.join(SCRIPT_DIR, "_checkup_ack")


def _checkup_ack_path(book_name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', book_name)
    return os.path.join(CHECKUP_ACK_DIR, safe + ".json")


def _load_checkup_ack(book_name):
    """確認済み（意図的）として除外する所見キーの集合を読む。無ければ空集合"""
    import json
    try:
        with open(_checkup_ack_path(book_name), 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, list):
            return set()
        # 文字列以外が混ざると保存時の sorted() で TypeError になるため除外
        return set(x for x in data if isinstance(x, str))
    except (OSError, ValueError):
        return set()


def _save_checkup_ack(book_name, ack_set):
    import json
    os.makedirs(CHECKUP_ACK_DIR, exist_ok=True)
    with open(_checkup_ack_path(book_name), 'w', encoding='utf-8') as f:
        json.dump(sorted(ack_set), f, ensure_ascii=False, indent=1)


def _checkup_diff(prev, cur):
    """前回スナップショットとの差分（純粋処理）。prev が無ければ None。

    所見キーは行番号を含まない形にしてあるので、行ずれでは差分にならない。
    """
    if not prev:
        return None
    # 履歴ファイルは外部データ＝どちら側もキー欠損に耐える（cur も .get で読む）
    pk, ck = set(prev.get('keys', [])), set(cur.get('keys', []))
    d = {'prev_time': prev.get('time'),
         'new': sorted(ck - pk), 'resolved': sorted(pk - ck)}
    # 履歴 JSON は外部データ。procs の値が list 以外に汚損していても診断ごと
    # 落とさない（キー欠損への耐性と同じ方針で型汚損にも耐える）
    def _safe(d):
        return {m: set(v) for m, v in (d.get('procs') or {}).items()
                if isinstance(v, (list, tuple, set))}
    pp = _safe(prev)
    cp = _safe(cur)
    added, removed = [], []
    for m in sorted(cp):
        for p in sorted(cp[m] - pp.get(m, set())):
            added.append(f"[{m}] {p}")
    for m in sorted(pp):
        for p in sorted(pp[m] - cp.get(m, set())):
            removed.append(f"[{m}] {p}")
    d['procs_added'], d['procs_removed'] = added, removed
    cur_lines = cur.get('total_lines', 0)
    d['lines_delta'] = cur_lines - prev.get('total_lines', cur_lines)
    cur_sheets, cur_forms = set(cur.get('sheets', [])), set(cur.get('forms', []))
    d['sheets_added'] = sorted(cur_sheets - set(prev.get('sheets', [])))
    d['sheets_removed'] = sorted(set(prev.get('sheets', [])) - cur_sheets)
    d['forms_added'] = sorted(cur_forms - set(prev.get('forms', [])))
    d['forms_removed'] = sorted(set(prev.get('forms', [])) - cur_forms)
    d['changed'] = bool(d['new'] or d['resolved'] or added or removed
                        or d['lines_delta'] or d['sheets_added'] or d['sheets_removed']
                        or d['forms_added'] or d['forms_removed'])
    return d


def _vba_references(wb):
    """VBA 参照設定の一覧と破損（MISSING）の検出（読み取りのみ）。

    眠っていたブックが動かない原因の筆頭＝「参照不可: ライブラリが見つかりません」を
    VBE を開かずに検出する。
    """
    total, broken = 0, []
    try:
        refs = wb.VBProject.References
    except Exception:
        return {'total': 0, 'broken': []}
    for r in refs:
        total += 1
        try:
            name = r.Name
        except Exception:
            name = "(名前不明)"
        try:
            is_broken = bool(r.IsBroken)
        except Exception:
            is_broken = True      # IsBroken すら読めない参照は破損として報告
        if is_broken:
            try:
                path = r.FullPath
            except Exception:
                path = None
            broken.append((name, path))
    return {'total': total, 'broken': broken}


def _sheet_health(wb):
    """シート側の健診（読み取りのみ）: エラーセル数。

    ※ ゴースト（使用範囲と実データの差）検査は撤去した（2026-07-05・shuさん指示）。
    値の無い行は「意図して罫線を引いた記入欄」のことが普通にあり（例:
    ファイル一覧.xlsm の修正シート＝連続修正の記入欄100行）、機械には
    取り残しと区別できない。この種の自動判定を根拠にした肥大縮小で
    取説シートの結合を壊した実害もある（2026-06-19）。復活させないこと。
    """
    out = {'error_cells': []}
    for sh in wb.Worksheets:
        try:
            ur = sh.UsedRange
        except Exception:
            continue
        n_err = 0
        for cell_type in (-4123, 2):          # 数式(-4123) / 定数(2)
            try:
                n_err += ur.SpecialCells(cell_type, 16).Count   # 16 = xlErrors
            except Exception:
                pass                           # 該当なしは SpecialCells が例外を投げる
        if n_err:
            out['error_cells'].append((sh.Name, n_err))
    return out


def _broken_links(wb):
    """外部ブックへのリンクのうち、参照先ファイルが存在しないもの（読み取りのみ）"""
    broken = []
    try:
        links = wb.LinkSources(1)              # 1 = xlExcelLinks
    except Exception:
        links = None
    for src in (links or ()):
        s = str(src)
        if re.match(r'^[A-Za-z]:\\|^\\\\', s) and not os.path.exists(s):
            broken.append(s)
    return broken


def _checkup_rating(total_findings, n_critical):
    """総合判定（機械的な分類。判断はしない）:
    C=壊れた参照・呼び出しあり / B=その他の所見あり / A=所見なし"""
    if n_critical:
        return "C（要確認）"
    if total_findings:
        return "B（軽度所見）"
    return "A（異常なし）"


def cmd_call_graph(args):
    """マクロの呼び出し関係を解析: call-graph [excel_file] [--macro 名]

    Call 文・Application.Run・既知プロシージャ名の裸呼びを機械的に解析する。
    - 未解決 Call: **存在しないマクロを呼んでいる行**（コピペ残骸の一語バグ検出器）
    - 呼び出し関係: どのマクロがどのマクロを使っているか
    - 孤立: どこからも呼ばれていないマクロ（メニュー/イベント直実行の可能性があるため
      機械は事実だけ報告し、要不要の判断はしない）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    focus = getattr(args, 'macro_opt', None)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False
    res = _analyze_calls(inv)
    known, edges = res['known'], res['edges']
    unresolved, orphans = res['unresolved'], res['orphans']

    if getattr(args, 'json', False):
        import json
        print(json.dumps({
            "success": True, "file": inv['name'],
            "edges": [{"caller_module": k[0], "caller": k[1], "callees": sorted(v)}
                      for k, v in sorted(edges.items())],
            "unresolved": [{"module": u[0], "proc": u[1], "name": u[2],
                            "line": u[3], "text": u[4]} for u in unresolved],
            "orphans": [{"module": o[0], "name": o[1]} for o in orphans],
        }, ensure_ascii=False), file=sys.stdout)
        return not unresolved

    if getattr(args, 'mermaid', None) is not None:
        # Mermaid 図（GitHub / Qiita でそのまま描画される）。モジュール別 subgraph、
        # 未解決の呼び先は赤ノード。呼び出しのあるマクロだけを図に載せる
        out_path = (os.path.join(SCRIPT_DIR, "_last_callgraph.md")
                    if args.mermaid == '_DEFAULT_' else os.path.abspath(args.mermaid))
        node_ids = {}

        def nid(name):
            if name not in node_ids:
                node_ids[name] = f"n{len(node_ids)}"
            return node_ids[name]

        used = set()
        for (mod, proc), callees in edges.items():
            if proc == '(宣言部)' or not callees:
                continue
            used.add(proc)
            used |= callees
        # 未解決Callしか持たないマクロもノード定義に載せる（辺だけ出すと
        # Mermaid が無ラベルの自動ノードを作り、肝心の呼び元が図から読めない）
        for _, _proc, _n, _, _ in unresolved:
            if _proc and _proc != '(宣言部)':
                used.add(_proc)
        ml = [f"# {inv['name']} 呼び出し関係図", "",
              f"生成: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager call-graph --mermaid）", "",
              "```mermaid", "flowchart LR"]
        by_mod = {}
        for m in inv['modules']:
            for p in m['procs']:
                if p['name'] in used:
                    by_mod.setdefault(m['name'], []).append(p['name'])
        for mod, procs in by_mod.items():
            ml.append(f'    subgraph {mod}')
            for p in procs:
                ml.append(f'        {nid(p)}["{p}"]')
            ml.append('    end')
        for (mod, proc), callees in sorted(edges.items()):
            if proc == '(宣言部)':
                continue
            for c in sorted(callees):
                ml.append(f'    {nid(proc)} --> {nid(c)}')
        bad_ids = []
        for _, proc, name, _, _ in unresolved:
            if proc:
                bid = nid(f"？{name}")
                bad_ids.append(bid)
                ml.append(f'    {bid}["{name}（存在しない）"]')
                ml.append(f'    {nid(proc)} -.-> {bid}')
        for b in set(bad_ids):
            ml.append(f'    style {b} fill:#ffcccc,stroke:#cc0000')
        ml.append("```")
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ml) + '\n')
        print(f"Mermaid 図を生成: {out_path}  "
              f"（ノード {len(node_ids)} / 未解決 {len(unresolved)}）")
        return not unresolved

    print(f"===== 呼び出し関係の解析: {inv['name']} =====")

    # 1. 未解決（最重要＝存在しないマクロを呼んでいる）
    if unresolved:
        print(f"\n⚠ 未解決の呼び出し（{len(unresolved)}件）— 存在しないマクロを呼んでいます:")
        for mod, proc, name, ln, text in unresolved:
            print(f"  [{mod}] {proc or '(宣言部)'} :{ln}  →  {name}")
            print(f"      {text}")
    else:
        print("\n未解決の呼び出し: なし（Call/Run はすべて実在のマクロを指しています）")

    # 2. 呼び出し関係（focus 指定ならそのマクロを起点にツリー展開）
    if focus:
        hit = known.get(focus.lower())
        if not hit:
            print(f"\nエラー: マクロ '{focus}' が見つかりません")
            _suggest_similar(focus, [v[0] for v in known.values()])
            return False
        root = hit[0]
        print(f"\n--- {root} からの呼び出しツリー ---")

        # 正式名→callee集合（モジュール横断で合算）
        by_name = {}
        for (mod, proc), callees in edges.items():
            by_name.setdefault(proc, set()).update(callees)

        def walk(name, depth, seen):
            for callee in sorted(by_name.get(name, ())):
                loop = "（循環）" if callee in seen else ""
                print("  " * depth + f"└ {callee}{loop}")
                if not loop and depth < 6:
                    walk(callee, depth + 1, seen | {callee})
        walk(root, 1, {root})
        callers = sorted({f"[{k[0]}] {k[1]}" for k, v in edges.items() if root in v})
        print(f"--- {root} を呼んでいるマクロ ---")
        for c in callers or ["  (なし)"]:
            print(f"  {c}" if not c.startswith("  ") else c)
    else:
        callers_with_edges = [(k, v) for k, v in sorted(edges.items())
                              if v and k[1] != '(宣言部)']
        print(f"\n--- 呼び出し関係（{len(callers_with_edges)}マクロが他マクロを使用） ---")
        for (mod, proc), callees in callers_with_edges:
            print(f"  [{mod}] {proc} → {', '.join(sorted(callees))}")

        if orphans:
            print(f"\n--- どこからも呼ばれていないマクロ（{len(orphans)}件・"
                  "メニュー/ショートカット直実行の可能性あり） ---")
            for mod, name in orphans[:40]:
                sc = inv['shortcuts'].get(name)
                print(f"  [{mod}] {name}" + (f"  ({sc})" if sc else ""))
            if len(orphans) > 40:
                print(f"  … 他 {len(orphans) - 40}件")
    return not unresolved


def _checkup_one(xl, wb, note=None, ack_all=False, quiet=False, include_forms=True):
    """1ブックぶんの健康診断（収集・検査・履歴比較まで）。

    note はカルテのメモ＝今回のスナップショットに添付して履歴に残す
    （「何をどう直した回か」を後から思い出すための治療欄）。
    ack_all=True のとき、今回のコード/フォーム所見を全て「確認済み（意図的）」
    として登録し、以降の所見サマリ・総合判定から除外する
    （存在しないマクロ呼び出し等の致命的所見は対象外＝常に表示する）。
    戻り値: {'name','rating','total','counts','diff','lines','json'}
    レポートの書き出しと画面表示は cmd_checkup 側で行う。
    """
    # VBA プロジェクトに触れないブック（パスワード保護・VBOM 未信頼）は
    # シート側だけの縮退診断に切り替える（入口で転ばない）
    vba_error = None
    try:
        inv = _collect_book_inventory(xl, wb, quiet=quiet)
    except Exception as e:
        vba_error = str(e).strip() or "VBAプロジェクトにアクセスできません"
        inv = _collect_book_inventory(xl, wb, include_vba=False, quiet=quiet)
    calls = _analyze_calls(inv)
    extra = _extra_code_scans(inv)
    refs = _vba_references(wb) if not vba_error else {'total': 0, 'broken': []}
    sheet_h = _sheet_health(wb)
    dead_links = _broken_links(wb)
    ref_names = [n['name'] for n in inv['names']
                 if n.get('refers_to') and '#REF!' in str(n.get('refers_to'))]

    # --- コードの機械検査（モジュール単位） ---
    # (比較キー, 表示文) のペア。キーは行番号を含めない＝行ずれで差分を出さない
    code_findings = []
    for m in inv['modules']:
        if not m['code']:
            continue
        norm = m['code'].replace('\r\n', '\n')
        dups = _find_duplicate_procedures(norm)
        for nm, lns in dups.items():
            code_findings.append((f"重複プロシージャ: [{m['name']}] {nm}",
                                  f"[{m['name']}] プロシージャ名の重複: {nm}"
                                  f"（行 {', '.join(map(str, lns))}）"))
        for ln, s in _find_consecutive_dup_lines(norm)[:10]:
            code_findings.append((f"連続重複行: [{m['name']}] 「{s[:40]}」",
                                  f"[{m['name']}] 連続する同一コード行: 行{ln} 「{s[:40]}」"))

    # --- フォームの検査（form_inspect の lint を統合） ---
    # Designer を開いてコントロールを1個ずつ COM で読むため診断時間の支配項。
    # 既定はスキップし、--form 指定時のみ実施（2026-07-17・普段の健診を数秒に）。
    form_findings = {}
    forms_checked = 0
    forms_skipped = not include_forms
    if forms_skipped:
        form_findings = None
    else:
        try:
            if vba_error:
                raise ImportError     # ロック中はフォームにも触れない（未検査扱い）
            import form_inspect as _fi
            for comp in wb.VBProject.VBComponents:
                if int(comp.Type) != 3:
                    continue
                forms_checked += 1
                try:
                    info = _fi.get_form_info(comp)
                    controls = _fi.collect_controls(comp)
                    cm = comp.CodeModule
                    code_txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines > 0 else None
                    f = _fi.lint_form(comp.Name, info, controls, code=code_txt)
                    if f:
                        form_findings[comp.Name] = f
                except Exception as e:
                    form_findings[comp.Name] = [f"(検査不可: {e})"]
        except ImportError:
            form_findings = None      # form_inspect.py が無い構成（マクロ管理のみ）

    # --- バックアップ状況 ---
    # 部分一致（base in name）だと「家計」が「家計簿2026」のバックアップまで
    # 数えてしまう。make_backup / make_module_backup の命名規則に揃えて数える
    base = os.path.splitext(inv['name'])[0]
    book_prefix = inv['name'] + ".backup_before_"   # ブック丸ごと
    mod_prefix = base + "_"                         # モジュール単位 (.bas)
    bk_count, bk_latest = 0, None
    if os.path.isdir(BACKUP_DIR):
        for name in os.listdir(BACKUP_DIR):
            if name.startswith(book_prefix) or name.startswith(mod_prefix):
                bk_count += 1
                t = os.path.getmtime(os.path.join(BACKUP_DIR, name))
                if bk_latest is None or t > bk_latest:
                    bk_latest = t

    # --- 確認済み（意図的）所見の除外 ---
    # コード検査・フォーム検査の所見だけが対象（存在しないマクロ呼び出し等の
    # 致命的な所見は対象外＝常に表示する）。ack_all 指定時はここで今回の
    # 所見を全て確認済みとして登録してから、その状態でフィルタする。
    ack_set = _load_checkup_ack(inv['name'])
    if ack_all:
        new_keys = {k for k, _ in code_findings}
        for fname, findings in (form_findings or {}).items():
            new_keys |= {f"フォーム {fname}: {s}" for s in findings}
        ack_new = new_keys - ack_set
        ack_set = ack_set | new_keys
        if ack_new:
            _save_checkup_ack(inv['name'], ack_set)
    else:
        ack_new = set()

    code_findings_active = [(k, t) for k, t in code_findings if k not in ack_set]
    code_findings_acked = [(k, t) for k, t in code_findings if k in ack_set]
    form_findings_active, form_findings_acked = {}, {}
    for fname, findings in (form_findings or {}).items():
        act, ackd = [], []
        for s in findings:
            (ackd if f"フォーム {fname}: {s}" in ack_set else act).append(s)
        if act:
            form_findings_active[fname] = act
        if ackd:
            form_findings_acked[fname] = ackd
    n_ack = len(code_findings_acked) + sum(len(v) for v in form_findings_acked.values())

    unresolved = calls['unresolved']
    n_form = sum(len(v) for v in form_findings_active.values())
    n_critical = (len(unresolved) + len(refs['broken'])
                  + len(dead_links) + len(ref_names))
    total_findings = n_critical + len(code_findings_active) + n_form
    rating = _checkup_rating(total_findings, n_critical)
    if vba_error:
        rating = "C（要確認）" if n_critical else "判定保留（VBA未検査）"
    total_procs = sum(len(m['procs']) for m in inv['modules'])
    total_lines = sum(m['total_lines'] for m in inv['modules'])
    n_err_cells = sum(n for _, n in sheet_h['error_cells'])

    # --- 定期健診: 前回スナップショットと比較して履歴に積む ---
    finding_keys = []
    for mod, proc, name, ln, text in unresolved:
        finding_keys.append(f"未解決Call: [{mod}] {proc or '(宣言部)'} → {name}")
    for name, path in refs['broken']:
        finding_keys.append(f"参照切れ: {name}")
    for s in dead_links:
        finding_keys.append(f"リンク切れ: {s}")
    for nm in ref_names:
        finding_keys.append(f"#REF!名前: {nm}")
    finding_keys += [k for k, _ in code_findings]
    for fname, findings in (form_findings or {}).items():
        finding_keys += [f"フォーム {fname}: {s}" for s in findings]

    # --- 確認済みリストの自動整理 ---
    # lintのロジック変更（閾値追加・チェック撤去等）で二度と出ない古いキーが
    # 確認済みリストに残り続けないよう、今回の生の所見に無いキーは毎回自動で
    # 除く。ただし検査が不完全な回は整理しない：
    #   - vba_error 時（縮退診断＝コード/フォームを見られていない）
    #   - form_findings is None 時（form_inspect.py が無い構成。フォームを
    #     検査できていないだけなのに、フォーム系の確認済みを「古いキー」と
    #     誤判定して全消しし、次にフォーム検査できる環境で全部復活してしまう）
    n_ack_pruned = 0
    if not vba_error and form_findings is not None:
        stale_ack = ack_set - set(finding_keys)
        if stale_ack:
            _save_checkup_ack(inv['name'], ack_set - stale_ack)
            n_ack_pruned = len(stale_ack)

    snap = {
        'time': time.strftime('%Y-%m-%d %H:%M'),
        'counts': {'unresolved': len(unresolved), 'code': len(code_findings_active),
                   'form': n_form, 'orphans': len(calls['orphans'])},
        'keys': sorted(finding_keys),
        'procs': {m['name']: [p['name'] for p in m['procs']] for m in inv['modules']},
        'total_procs': total_procs, 'total_lines': total_lines,
        'total_findings': total_findings, 'rating': rating, 'ack_count': n_ack,
        'sheets': [s['name'] for s in inv['sheets']],
        'forms': [f['name'] for f in inv['forms']],
    }
    if note:
        snap['note'] = note
    if vba_error:
        diff = None       # ロック中の縮退診断は履歴に混ぜない（解除後の差分が乱れる）
    else:
        hist = _load_checkup_history(inv['name'])
        # フォーム未検査の回は前回スナップショットのフォーム所見キーを引き継ぐ。
        # 検査していないだけの所見が「解消」と誤記録され、次の --form 診断で
        # 全部「新規」に化けるのを防ぐ（フォームの真実は --form の回だけが更新する）
        if form_findings is None and hist:
            carried = [k for k in hist[-1].get('keys', [])
                       if isinstance(k, str) and k.startswith('フォーム ')]
            if carried:
                snap['keys'] = sorted(set(snap['keys']) | set(carried))
        diff = _checkup_diff(hist[-1] if hist else None, snap)
        try:
            _save_checkup_history(inv['name'], hist + [snap])
        except OSError as e:
            print(f"警告: 診断履歴を保存できませんでした（{e}）。比較機能は次回も初回扱いになります")

    json_payload = {
        "success": total_findings == 0, "file": inv['name'], "rating": rating,
        "total_findings": total_findings,
        "unresolved_calls": len(unresolved),
        "broken_references": [{"name": n, "path": p} for n, p in refs['broken']],
        "broken_links": dead_links, "ref_error_names": ref_names,
        "code_findings": [d for _, d in code_findings_active],
        "form_findings": form_findings_active, "orphans": len(calls['orphans']),
        "acknowledged_count": n_ack, "newly_acknowledged": len(ack_new),
        "backups": bk_count,
        "error_cells": [{"sheet": s, "count": n} for s, n in sheet_h['error_cells']],
        "auto_exec": len(extra['auto_exec']),
        "hardcoded_paths": len(extra['hardcoded_paths']),
        "no_option_explicit": extra['no_option_explicit'],
        "error_resume": len(extra['error_resume']),
        "long_procs": len(extra['long_procs']),
        "destructive": len(extra['destructive']),
        "no_restore": len(extra['no_restore']),
        "onaction": calls['onaction'],
        "dynamic_runs": len(calls['dynamic_runs']),
        "forms_skipped": forms_skipped,
        "vba_error": vba_error,
        "diff": diff,
    }
    L = []
    L.append(f"# {inv['name']} 健康診断レポート")
    L.append("")
    L.append(f"- 診断日時: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager checkup）")
    if vba_error:
        L.append(f"- 規模: シート {len(inv['sheets'])} / マクロ・フォームは VBA 未検査のため不明")
    else:
        L.append(f"- 規模: シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
                 f"フォーム {len(inv['forms'])}")
    L.append(f"- 総合判定: **{rating}**"
             "（A=所見なし / B=所見あり / C=壊れた参照・呼び出しあり。機械的な分類です）")
    if forms_skipped:
        L.append("- フォーム検査: 未実施（`checkup --form` で実施）")
    if vba_error:
        L.append("- ⚠ **VBA は検査できませんでした**（プロジェクトのパスワード保護、"
                 "または VBOM 未信頼）。以下はシート側だけの縮退診断です")
    L.append("")

    L.append("## 所見サマリ")
    L.append("")
    mark = "✅" if total_findings == 0 else "⚠"
    L.append(f"{mark} **検出された所見: {total_findings}件**"
             + (f"（確認済み・意図的として除外: {n_ack}件 → 末尾参照）" if n_ack else ""))
    if ack_new:
        L.append(f"　今回 {len(ack_new)}件を新たに確認済み登録しました。")
    L.append("")
    L.append(f"| 診察項目 | 所見 |")
    L.append(f"|---|---|")
    if vba_error:
        L.append(f"| 外部ブックへのリンク切れ | **{len(dead_links)}件** |")
        L.append(f"| #REF! になった名前付き範囲 | **{len(ref_names)}件** |")
        L.append(f"| エラーセル（#REF!/#NAME?等） | {n_err_cells}個/"
                 f"{len(sheet_h['error_cells'])}シート（参考） |")
        L.append(f"| ボタン/図形に登録されたマクロ | {len(inv['onaction'])}箇所（参考） |")
        L.append("| VBAの検査（コード/フォーム/参照設定） | 実施できず（保護または未信頼） |")
    else:
        L.append(f"| 存在しないマクロへの呼び出し | **{len(unresolved)}件** |")
        L.append(f"| 参照設定の破損（MISSING） | **{len(refs['broken'])}件**（全{refs['total']}参照） |")
        L.append(f"| 外部ブックへのリンク切れ | **{len(dead_links)}件** |")
        L.append(f"| #REF! になった名前付き範囲 | **{len(ref_names)}件** |")
        L.append(f"| コードの機械検査（重複等） | {len(code_findings_active)}件 |")
        if form_findings is not None:
            L.append(f"| フォームの検査（{forms_checked}フォーム） | {n_form}件 |")
        elif forms_skipped:
            L.append("| フォームの検査 | 未実施（--form 指定時のみ） |")
        if n_ack:
            L.append(f"| 確認済み（意図的・上記から除外済み） | {n_ack}件 |")
        L.append(f"| エラーセル（#REF!/#NAME?等） | {n_err_cells}個/"
                 f"{len(sheet_h['error_cells'])}シート（参考） |")
        L.append(f"| 破壊的な操作（Kill/シート削除等） | {len(extra['destructive'])}箇所（参考） |")
        L.append(f"| ScreenUpdating/Calculation/EnableEvents 戻し忘れ | {len(extra['no_restore'])}件（参考） |")
        L.append(f"| どこからも呼ばれていないマクロ | {len(calls['orphans'])}件（参考・"
                 "メニュー/直実行の可能性あり） |")
        L.append(f"| ボタン/図形に登録されたマクロ | {len(calls['onaction'])}件（参考） |")
        L.append(f"| 動的な Application.Run（実行時に呼び先決定） | "
                 f"{len(calls['dynamic_runs'])}箇所（参考） |")
        L.append(f"| 自動実行イベント（開く/変更等で起動） | {len(extra['auto_exec'])}件（参考） |")
        L.append(f"| ハードコードされたパス | {len(extra['hardcoded_paths'])}箇所（参考） |")
        L.append(f"| Option Explicit なしのモジュール | {len(extra['no_option_explicit'])}件（参考） |")
        L.append(f"| On Error Resume Next | {len(extra['error_resume'])}箇所（参考） |")
        L.append(f"| 150行を超えるプロシージャ | {len(extra['long_procs'])}件（参考） |")
    L.append(f"| バックアップ | {bk_count}件"
             + (f"（最新: {time.strftime('%Y-%m-%d %H:%M', time.localtime(bk_latest))}）"
                if bk_latest else "（なし）") + " |")
    L.append("")

    L.append("## 前回との比較（定期健診）")
    L.append("")
    if vba_error:
        L.append("VBA 未検査の縮退診断のため、履歴への記録と前回比較は行いません。")
    elif diff is None:
        L.append("初回の診断のため比較対象がありません（今回の結果を記録しました。"
                 "次回からここに前回との差分が出ます）。")
    elif not diff['changed']:
        L.append(f"前回（{diff['prev_time']}）から変化はありません。")
    else:
        L.append(f"前回の診断: {diff['prev_time']}")
        L.append("")
        for s in diff['new']:
            L.append(f"- ＋ 新しい所見: {s}")
        for s in diff['resolved']:
            L.append(f"- － 解消した所見: {s}")
        for s in diff['procs_added']:
            L.append(f"- ＋ マクロ追加: {s}")
        for s in diff['procs_removed']:
            L.append(f"- － マクロ削除: {s}")
        if diff['lines_delta']:
            sign = '+' if diff['lines_delta'] > 0 else ''
            L.append(f"- コード行数: {sign}{diff['lines_delta']}行（計{total_lines}行）")
        for s in diff['sheets_added']:
            L.append(f"- ＋ シート追加: {s}")
        for s in diff['sheets_removed']:
            L.append(f"- － シート削除: {s}")
        for s in diff['forms_added']:
            L.append(f"- ＋ フォーム追加: {s}")
        for s in diff['forms_removed']:
            L.append(f"- － フォーム削除: {s}")
    L.append("")

    if unresolved:
        L.append("## ⚠ 存在しないマクロへの呼び出し（最優先で確認）")
        L.append("")
        for mod, proc, name, ln, text in unresolved:
            L.append(f"- `[{mod}] {proc or '(宣言部)'}` の {ln} 行目 → **{name}**")
            L.append(f"  ```vba")
            L.append(f"  {text}")
            L.append(f"  ```")
        L.append("")

    if refs['broken'] or dead_links or ref_names:
        L.append("## ⚠ 壊れた参照（最優先で確認）")
        L.append("")
        for name, path in refs['broken']:
            L.append(f"- 参照設定の破損（MISSING）: **{name}**"
                     + (f" `{path}`" if path else ""))
        if refs['broken']:
            L.append("  （VBE の [ツール]→[参照設定] で「参照不可」になっている項目。"
                     "コンパイルエラーの典型原因）")
        for s in dead_links:
            L.append(f"- 外部ブックへのリンク切れ: `{s}`（参照先ファイルが存在しません）")
        for nm in ref_names:
            L.append(f"- #REF! になった名前付き範囲: **{nm}**")
        L.append("")

    if code_findings_active:
        L.append("## コードの機械検査")
        L.append("")
        for _, s in code_findings_active:
            L.append(f"- {s}")
        L.append("")

    if form_findings_active:
        L.append("## フォームの検査")
        L.append("")
        for fname, findings in form_findings_active.items():
            L.append(f"### {fname}")
            for s in findings:
                L.append(f"- {s}")
            L.append("")

    if n_ack:
        L.append("## 確認済み（意図的・件数から除外）")
        L.append("")
        L.append("目で見て意図的なデザイン等と判断し、`checkup --ack-all` で確認済み登録した"
                 "所見です。事実としては残しつつ、所見サマリ・総合判定からは除外しています。"
                 "判断を見直す場合は `checkup --unack \"文字列\"` で確認済みから外せます。")
        L.append("")
        for _, s in code_findings_acked:
            L.append(f"- {s}")
        for fname, findings in form_findings_acked.items():
            L.append(f"### {fname}")
            for s in findings:
                L.append(f"- {s}")
        L.append("")

    if sheet_h['error_cells']:
        L.append("## シートの検査（参考）")
        L.append("")
        for s, n in sheet_h['error_cells']:
            L.append(f"- {s}: エラーセル {n}個（#REF!/#NAME?/#VALUE! 等）")
        L.append("")

    if calls['orphans']:
        L.append("## どこからも呼ばれていないマクロ（参考）")
        L.append("")
        L.append("メニューやショートカットからの直実行用かもしれません（機械には判断できません）。")
        L.append("")
        for mod, name in calls['orphans'][:40]:
            sc = inv['shortcuts'].get(name)
            L.append(f"- [{mod}] {name}" + (f"（{sc}）" if sc else ""))
        if len(calls['orphans']) > 40:
            L.append(f"- … 他 {len(calls['orphans']) - 40}件")
        L.append("")

    if extra['auto_exec']:
        L.append("## 自動実行される処理（参考）")
        L.append("")
        L.append("ブックを開く・保存する・セルを変更する等で自動的に動くマクロです。"
                 "眠っていたブックを起こす前の問診に。")
        L.append("")
        for mod, name, lns in extra['auto_exec']:
            L.append(f"- [{mod}] {name}" + (f"（{lns}行）" if lns else ""))
        L.append("")

    if calls['onaction']:
        L.append("## ボタン・図形から実行されるマクロ（参考）")
        L.append("")
        L.append("シート上のボタン/図形に登録（OnAction）されているマクロです。"
                 "これらは「どこからも呼ばれていないマクロ」には数えていません。")
        L.append("")
        for name, places in sorted(calls['onaction'].items()):
            more = f" 他{len(places) - 5}箇所" if len(places) > 5 else ""
            L.append(f"- {name} ← {', '.join(places[:5])}{more}")
        L.append("")

    if calls['dynamic_runs']:
        L.append("## 動的な Application.Run（参考）")
        L.append("")
        L.append("呼び先が実行時に変数で決まる Run です（メニュー機構などの正常な作り）。"
                 "静的検査では実在確認ができないため、事実として所在だけ記します。")
        L.append("")
        for mod, proc, ln, text in calls['dynamic_runs'][:15]:
            L.append(f"- [{mod}] {proc}:{ln} `{text}`")
        if len(calls['dynamic_runs']) > 15:
            L.append(f"- … 他 {len(calls['dynamic_runs']) - 15}箇所")
        L.append("")

    if extra['hardcoded_paths']:
        L.append("## ハードコードされたパス（参考）")
        L.append("")
        L.append("フォルダ構成が変わると動かなくなる箇所の候補です（古いブックの復活時に特に確認）。")
        L.append("")
        for mod, proc, ln, path in extra['hardcoded_paths'][:20]:
            L.append(f"- [{mod}] {proc}:{ln} `{path}`")
        if len(extra['hardcoded_paths']) > 20:
            L.append(f"- … 他 {len(extra['hardcoded_paths']) - 20}箇所")
        L.append("")

    if extra['destructive']:
        L.append("## 破壊的な操作の所在（参考・問診）")
        L.append("")
        L.append("ファイル削除・シート削除などを行う箇所です（悪ではありません。"
                 "眠っていたブックを起こす前に「どこで何を消すか」を知っておくための一覧）。")
        L.append("")
        for mod, proc, ln, label, snippet in extra['destructive'][:20]:
            L.append(f"- [{mod}] {proc}:{ln} {label} `{snippet}`")
        if len(extra['destructive']) > 20:
            L.append(f"- … 他 {len(extra['destructive']) - 20}箇所")
        L.append("")

    if (extra['no_option_explicit'] or extra['error_resume'] or extra['long_procs']
            or extra['no_restore']):
        L.append("## その他の参考情報")
        L.append("")
        for mod, proc, desc in extra['no_restore']:
            L.append(f"- {desc}: [{mod}] {proc}")
        if extra['no_option_explicit']:
            L.append(f"- Option Explicit なしのモジュール: "
                     f"{', '.join(extra['no_option_explicit'])}")
        for mod, proc, ln in extra['error_resume'][:15]:
            L.append(f"- On Error Resume Next: [{mod}] {proc}:{ln}")
        if len(extra['error_resume']) > 15:
            L.append(f"- … On Error Resume Next 他 {len(extra['error_resume']) - 15}箇所")
        for mod, name, lns in extra['long_procs'][:10]:
            L.append(f"- 150行超のプロシージャ: [{mod}] {name}（{lns}行）")
        if len(extra['long_procs']) > 10:
            L.append(f"- … 150行超 他 {len(extra['long_procs']) - 10}件")
        L.append("")

    L.append("## 次の一手")
    L.append("")
    L.append("- 構成の全貌: `py vba_manager.py docs --preview 3`")
    L.append("- 呼び出し関係の図: `py vba_manager.py call-graph --mermaid`")
    L.append("- 修正前の影響確認: `py vba_manager.py impact <マクロ名>`（呼び元・呼び先を間接まで一覧）")
    L.append("- 経過観察: `py vba_manager.py 健康診断 --history`（診断履歴の一覧表）")
    if forms_skipped:
        L.append("- フォーム込みのフル診断: `py vba_manager.py checkup --form`")
    if form_findings_active:
        L.append("- フォームの修正: `py form_tool.py`（move / align / tab-order …）")
    if unresolved:
        L.append("- 呼び先の検索: `py vba_manager.py grep \"<マクロ名>\"`")
    if total_findings and not vba_error:
        L.append("- 意図的な所見の除外: `py vba_manager.py checkup --ack-all`"
                 "（今回の所見を確認済みとして次回以降の件数・判定から除外）")
    L.append("")

    return {'name': inv['name'], 'rating': rating, 'total': total_findings,
            'counts': {'unresolved': len(unresolved), 'code': len(code_findings_active),
                       'form': n_form,
                       'broken': len(refs['broken']) + len(dead_links) + len(ref_names)},
            'ack_new': len(ack_new), 'ack_total': n_ack, 'ack_pruned': n_ack_pruned,
            'forms_skipped': forms_skipped,
            'diff': diff, 'vba_error': vba_error, 'lines': L, 'json': json_payload}


def cmd_checkup(args):
    """ブックの健康診断レポート: checkup(健康診断) [excel_file] [--out f.md]

    構成（docs）＋呼び出し関係（call-graph）＋壊れた参照/リンク＋シート検査＋
    フォーム検査（lint）＋コードの機械検査＋バックアップ状況を 1枚の Markdown に
    束ね、総合判定（A/B/C）を付ける。所見は事実の列挙のみ（要不要の判断はしない）。
    診断のたびにブック別の履歴（_checkup_history/）へ結果を残し、
    次回の診断で「前回との比較（定期健診）」を自動で出す。
    フォーム検査（lint）は --form 指定時のみ（既定はスキップ＝数秒診断。
    Designer をコントロール単位で COM 走査するため診断時間の支配項だから）。
    --history=経過観察（履歴の一覧表・診断はしない）、--all=開いている全ブックを一括診断。
    --ack-all=今回の所見を確認済み（意図的）として登録し、以降の件数・判定から除外。
    --show-ack=確認済み一覧を表示、--unack 文字列=部分一致するものを確認済みから外す。
    """
    # --- 確認済み一覧の表示（診断はしない） ---
    if getattr(args, 'show_ack', False):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        ack = _load_checkup_ack(name)
        if not ack:
            print(f"{name} に確認済み（意図的）の所見はありません")
            return True
        print(f"===== {name} の確認済み（意図的・件数から除外）所見 {len(ack)}件 =====")
        for k in sorted(ack):
            print(f"  - {k}")
        return True

    # --- 確認済みの取り消し（部分一致で解除） ---
    if getattr(args, 'unack', None):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        ack = _load_checkup_ack(name)
        pat = args.unack
        matched = {k for k in ack if pat in k}
        if not matched:
            print(f"'{pat}' に一致する確認済み所見が見つかりません")
            return False
        _save_checkup_ack(name, ack - matched)
        print(f"{len(matched)}件を確認済みから外しました:")
        for k in sorted(matched):
            print(f"  - {k}")
        return True

    # --- 経過観察モード: 診断せず履歴を表で表示 ---
    if getattr(args, 'history', False):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        hist = _load_checkup_history(name)
        if not hist:
            print(f"{name} の診断履歴はまだありません（checkup を実行すると記録されます）")
            return False
        print(f"===== {name} の経過観察（診断履歴 {len(hist)}回） =====")
        print(f"{'日時':<18}{'判定':<12}{'所見':>4} {'未解決':>4} {'フォーム':>4} "
              f"{'マクロ':>4} {'行数':>7}  メモ")
        prev = None
        for s in hist:
            c = s.get('counts', {})
            tf = s.get('total_findings')
            if tf is None:      # 旧形式のスナップショット（判定・合計なし）
                tf = c.get('unresolved', 0) + c.get('code', 0) + c.get('form', 0)
            note = s.get('note', '')
            print(f"{s.get('time', ''):<18}{s.get('rating', '-'):<12}{tf:>4} "
                  f"{c.get('unresolved', 0):>6} {c.get('form', 0):>7} "
                  f"{s.get('total_procs', 0):>6} {s.get('total_lines', 0):>7}"
                  + (f"  {note}" if note else ""))
            if getattr(args, 'detail', False) and prev is not None:
                d = _checkup_diff(prev, s)
                if d and d['changed']:
                    for x in d['new']:
                        print(f"        ＋ {x}")
                    for x in d['resolved']:
                        print(f"        － {x}")
                    for x in d['procs_added']:
                        print(f"        ＋ マクロ追加 {x}")
                    for x in d['procs_removed']:
                        print(f"        － マクロ削除 {x}")
                    if d['lines_delta']:
                        sign = '+' if d['lines_delta'] > 0 else ''
                        print(f"        行数 {sign}{d['lines_delta']}")
            prev = s
        return True

    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    if getattr(args, 'all_books', False):
        books = list(xl.Workbooks)
    else:
        books = [wb]

    results = []
    for b in books:
        try:
            results.append(_checkup_one(xl, b, note=getattr(args, 'note', None),
                                        ack_all=getattr(args, 'ack_all', False),
                                        quiet=getattr(args, 'json', False),
                                        include_forms=getattr(args, 'form', False)))
        except Exception as e:
            print(f"警告: {getattr(b, 'Name', '?')} の診断に失敗しました: {e}")
    if not results:
        print("エラー: 診断できたブックがありません")
        return False

    if getattr(args, 'json', False):
        import json
        if len(results) == 1:
            payload = results[0]['json']
        else:
            payload = {"success": all(r['json']['success'] for r in results),
                       "books": [r['json'] for r in results]}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
        if getattr(args, 'strict', False):
            return all(r['total'] == 0 for r in results)
        return True

    if len(results) > 1:
        lines = [f"# 開いている全ブックの健康診断（{len(results)}冊）", "",
                 f"- 診断日時: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager checkup --all）",
                 "", "| ブック | 総合判定 | 所見 |", "|---|---|---|"]
        for r in results:
            lines.append(f"| {r['name']} | {r['rating']} | {r['total']}件 |")
        lines.append("")
        for r in results:
            lines += ["---", ""] + r['lines'] + [""]
    else:
        lines = results[0]['lines']

    out_path = getattr(args, 'out_opt', None)
    out_path = os.path.abspath(out_path) if out_path else os.path.join(SCRIPT_DIR, "_last_checkup.md")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f"健康診断レポートを生成: {out_path}")
    for r in results:
        c = r['counts']
        form_disp = "未実施(--form)" if r.get('forms_skipped') else c['form']
        print(f"[{r['name']}] 総合判定 {r['rating']} / 所見 {r['total']}件"
              f"（未解決Call {c['unresolved']} / 参照・リンク切れ {c['broken']} / "
              f"コード {c['code']} / フォーム {form_disp}）")
        if r.get('ack_new'):
            print(f"  今回 {r['ack_new']}件を確認済み（意図的）として新規登録"
                  f"（確認済み合計 {r['ack_total']}件）")
        if r.get('ack_pruned'):
            print(f"  確認済みのうち {r['ack_pruned']}件は今回の所見に存在しないため自動整理しました"
                  "（lintロジック変更で二度と出ないキー等）")
        d = r['diff']
        if r.get('vba_error'):
            print("  VBA未検査（プロジェクト保護またはVBOM未信頼）＝シート側のみの縮退診断・履歴記録なし")
        elif d is None:
            print("  前回比: 初回診断（履歴を記録しました。次回から差分が出ます）")
        elif not d['changed']:
            print(f"  前回比: 変化なし（前回 {d['prev_time']}）")
        else:
            print(f"  前回比: 新規所見 {len(d['new'])} / 解消 {len(d['resolved'])} / "
                  f"マクロ +{len(d['procs_added'])}−{len(d['procs_removed'])}"
                  f"（前回 {d['prev_time']}）")
    # 所見があっても「診断の完了」は成功（終了コード0）。所見の有無で合否を
    # 判定したい自動化（CI/batch のゲート）だけ --strict で従来挙動にする
    if getattr(args, 'strict', False):
        return all(r['total'] == 0 for r in results)
    return True


def cmd_impact(args):
    """マクロ修正前の影響範囲予告: impact(影響範囲) [excel_file] <マクロ名>

    「このマクロに手を入れると、どこまで波及するか」を修正前に一覧する。
    - 呼び元（上流・間接含む）＝動作を変えたとき影響が及ぶ先
    - 呼び先（下流・間接含む）＝このマクロが依存している部品
    - 入口（ショートカット/自動実行イベント）も注記する
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: impact [excel_file] <マクロ名> [--json]")
        return False
    focus = rest[0]
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（読むだけ）
    if not target_file:
        wb = _book_holding_proc(wb, focus) or wb      # （2026-09-24 総点検）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False
    res = _analyze_calls(inv)
    known, edges = res['known'], res['edges']
    hit = known.get(focus.lower())
    if not hit:
        print(f"エラー: マクロ '{focus}' が見つかりません")
        _suggest_similar(focus, [v[0] for v in known.values()], wb=wb)
        return False
    root = hit[0]
    mod_of = {v[0]: v[1] for v in known.values()}
    auto_names = {name for _, name, _ in _extra_code_scans(inv)['auto_exec']}
    oa = res.get('onaction', {})      # マクロ名 → ["シート/図形", ...]

    # 名前レベルの正方向・逆方向グラフ（call-graph と同じ粒度）
    fwd, rev = {}, {}
    for (mod, proc), callees in edges.items():
        fwd.setdefault(proc, set()).update(callees)
        for c in callees:
            rev.setdefault(c, set()).add(proc)

    def reach(graph, start):
        seen, stack = set(), [start]
        while stack:
            for nxt in graph.get(stack.pop(), ()):
                if nxt not in seen and nxt != start:
                    seen.add(nxt)
                    stack.append(nxt)
        return seen

    upstream, downstream = reach(rev, root), reach(fwd, root)

    def entry_note(name):
        notes = []
        sc = inv['shortcuts'].get(name)
        if sc:
            notes.append(sc)
        if name in auto_names:
            notes.append("自動実行イベント")
        if name in oa:
            notes.append("ボタン: " + ", ".join(oa[name][:3])
                         + (f" 他{len(oa[name]) - 3}" if len(oa[name]) > 3 else ""))
        return f"（{'/'.join(notes)}）" if notes else ""

    if getattr(args, 'json', False):
        import json
        print(json.dumps({
            "success": True, "file": inv['name'], "macro": root,
            "upstream": sorted(upstream), "downstream": sorted(downstream),
            "entries": {n: (inv['shortcuts'].get(n) or
                            ("ボタン: " + ", ".join(oa[n]) if n in oa else "自動実行イベント"))
                        for n in sorted(upstream | {root})
                        if n in inv['shortcuts'] or n in auto_names or n in oa},
        }, ensure_ascii=False), file=sys.stdout)
        return True

    def label(name):
        m = mod_of.get(name)
        return (f"[{m}] {name}" if m else name) + entry_note(name)

    print(f"===== 影響範囲の予告: {label(root)} =====")

    print(f"\n■ 呼び元（このマクロを直すと影響が及ぶ先・間接含む {len(upstream)}件）")
    if upstream:
        def walk_up(name, depth, seen):
            for caller in sorted(rev.get(name, ())):
                loop = "（循環）" if caller in seen else ""
                print("  " * depth + f"└ {label(caller)}{loop}")
                if not loop and depth < 6:
                    walk_up(caller, depth + 1, seen | {caller})
        walk_up(root, 1, {root})
    else:
        print("  (なし) — メニュー/ショートカット/ボタン/イベント直実行の可能性があります")

    print(f"\n■ 呼び先（このマクロが依存している部品・間接含む {len(downstream)}件）")
    if downstream:
        def walk_down(name, depth, seen):
            for callee in sorted(fwd.get(name, ())):
                loop = "（循環）" if callee in seen else ""
                print("  " * depth + f"└ {label(callee)}{loop}")
                if not loop and depth < 6:
                    walk_down(callee, depth + 1, seen | {callee})
        walk_down(root, 1, {root})
    else:
        print("  (なし) — 単体で完結しています")

    entries = [n for n in sorted(upstream | {root})
               if n in inv['shortcuts'] or n in auto_names or n in oa]
    if entries:
        print(f"\n■ 入口（ショートカット/ボタン/自動実行から届く経路）")
        for n in entries:
            print(f"  {label(n)}")
    return True


def _needle_from_file(path):
    """--file で渡された検索語（1 行目の改行だけ落とす）。指定が無ければ None・読めなければ False（2026-09-23）。

    シェルの引用符で割れる語（'Like "function *"' が 2 つに割れて検索が外れた・2026-09-22）を、
    引用符を経由せずに渡す口。ファイルの中身をそのまま検索語にする（前後の空白は落とさない＝
    行頭のインデントごと探せる）。複数行のファイルは 1 行目だけ使う（行またぎの検索はできないため）。
    """
    if not path:
        return None
    try:
        with open(str(path), encoding='utf-8-sig') as f:
            text = f.read()
    except OSError as ex:
        print(f"エラー: --file を読めません: {ex}")
        return False
    needle = text.split('\n')[0].rstrip('\r')
    if not needle:
        print(f"エラー: --file の 1 行目が空です: {path}")
        return False
    return needle


def cmd_grep(args):
    """全モジュール横断のVBAコード検索: grep [excel_file] <検索文字列>

    「どのマクロが ActiveSheet を使っているか」等を1回のCOM接続で調べる。
    出力: [モジュール] プロシージャ名:行番号: 該当行
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    needle = _needle_from_file(getattr(args, 'needle_file', None))
    if needle is False:
        return False
    if needle is None:
        if not rest:
            print("使い方: grep [excel_file] <検索文字列> [--regex] [-i] [--module 名] [--max N] [--json]"
                  "\n  引用符で割れる文字（' \" 等）を探すときは --file 検索語.txt（中身をそのまま検索語にする）")
            return False
        needle = rest[0]
        if _reject_extra_args(rest, 1, '検索文字列は1つ。スペースを含むならクォートで囲む'
                                       '（引用符で割れるなら --file 検索語.txt）'):
            return False
    flags = re.IGNORECASE if getattr(args, 'ignore_case', False) else 0
    if getattr(args, 'regex', False):
        try:
            pat = re.compile(needle, flags)
        except re.error as e:
            print(f"エラー: 正規表現が不正です: {e}")
            return False
    else:
        pat = re.compile(re.escape(needle), flags)
    mod_filter = getattr(args, 'module_opt', None)
    # `or 200` だと --max 0（件数だけ見たい）が偽値で既定に化ける（is None 判定にする）
    _m = getattr(args, 'max_hits', None)
    max_hits = 200 if _m is None else int(_m)
    if max_hits < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False

    # コードを読むだけ → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, wb = get_workbook(target_file, readonly=True)
    # --all: 開いている全部のブック・アドイン（PERSONAL・秀コンボ.xlam も）を横に探す（2026-09-24: 会話記録で
    # grep … --all が不明な引数で落ちていた。list --all と同じ口）
    projects = [(None, wb.VBProject)]
    if getattr(args, 'all', False):
        projects = []
        for vp in xl.VBE.VBProjects:
            try:
                projects.append((_project_book_name(xl, vp) or vp.Name, vp))
            except Exception:
                continue

    def _scan(p):
        hits, total, scanned = [], 0, []
        for label, vp in projects:
            try:
                comps = list(vp.VBComponents)
            except Exception:
                continue                       # 保護されたプロジェクト
            for comp in comps:
                if mod_filter and comp.Name.lower() != mod_filter.lower():
                    continue
                cm = comp.CodeModule
                n = cm.CountOfLines
                if n == 0:
                    continue
                code = cm.Lines(1, n)
                scanned.append(code)
                for i, line in enumerate(code.split('\r\n'), 1):
                    if p.search(line):
                        total += 1
                        if len(hits) < max_hits:
                            try:
                                proc = cm.ProcOfLine(i, 0) or ''
                                # dynamic Dispatch は out引数付きメソッドをタプルで返すことがある
                                if isinstance(proc, tuple):
                                    proc = proc[0] or ''
                            except Exception:
                                proc = ''
                            hits.append({'book': label, 'module': comp.Name, 'proc': proc,
                                         'line': i, 'text': line.rstrip()})
        return hits, total, scanned

    hits, total, _scanned = _scan(pat)   # _scanned は 0件のときの「寛容に数え直し」用（COM を読み直さない）
    if total == 0 and not getattr(args, 'regex', False) and '|' in needle.strip('|'):
        # 「A|B|C」は「どれか」のつもり（2026-09-24: grep "選択セル図削除|非表示の行列を全表示|…" が 0 件と答え、
        # 会話記録でも「CreateObject|GetObject|…」を --regex 無しで撃っていた）。字面で 0 件のときだけ「どれか」で探し直す
        alts = [a for a in needle.split('|') if a]
        pat = re.compile('|'.join(re.escape(a) for a in alts), flags)
        hits, total, _scanned = _scan(pat)
        if total:
            print(f"（'|' を「どれか」と読んで探しました: {' / '.join(alts)}。字面の '|' を探すなら --regex で \\| と書く）")

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": wb.Name, "pattern": needle,
                          "total": total, "hits": hits}, ensure_ascii=False), file=sys.stdout)
        return True
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
        if not getattr(args, 'regex', False) and not getattr(args, 'ignore_case', False):
            # VBA は保存時に識別子の大小文字や演算子前後の空白を揃える＝書いた字面で空振りしやすい。
            # 寛容に数え直して案内する（黙って0件にしない・2026-08-23）
            tol = _vba_tolerant_pattern(needle)
            n_tol = sum(1 for code in _scanned for ln in code.split('\r\n') if tol.search(ln))
            if n_tol:
                print(f"  （大小文字・空白の違いを無視すると {n_tol} 件あります。VBA は保存時に識別子の大小文字や"
                      "演算子前後の空白を揃えるので、-i で探すか、出てきた字面で指定してください）")
        return True
    for h in hits:
        proc_part = f" {h['proc']}" if h['proc'] else ""
        book_part = f"[{h['book']}]" if h.get('book') else ""
        print(f"{book_part}[{h['module']}]{proc_part}:{h['line']}: {h['text'].strip()}")
    if total > len(hits):
        print(f"…他 {total - len(hits)}件（--max で上限変更可）")
    print(f"--- {total}件 ヒット ---" + (f"（{len(projects)} 冊を探した）" if getattr(args, 'all', False) else ""))
    return True


def _vba_tolerant_pattern(needle):
    """VBA が保存時に揃えてしまう字面の違いに寛容な検索パターンを作る（2026-08-23）。

    VBA は識別子の大小文字をプロジェクト内の初出に揃え（`.Value` と書いても、どこかに `value` が
    あれば `.value` で保存される）、演算子の前後の空白も整える。書いた字面で探すと空振りする。
    空白の並びは「あってもなくても」(\\s*)、演算子・区切り記号の前後も \\s* で挟み、大小文字は無視。
    文字そのものは re.escape（正規表現ではなく字面として扱う）。code-replace の探し直し用。
    """
    parts = []
    for tok in re.split(r'(\s+)', needle):
        if not tok:
            continue
        if tok.isspace():
            parts.append(r'\s*')
            continue
        for ch in tok:
            if ch in '=+-*/\\&<>(),:':
                parts.append(r'\s*' + re.escape(ch) + r'\s*')
            else:
                parts.append(re.escape(ch))
    return re.compile(''.join(parts), re.IGNORECASE)


def cmd_code_replace(args):
    """全マクロ横断の一括置換: code-replace <検索> <置換>

    grep の対。差分プレビュー → 確認 → バックアップ → 変更行だけ ReplaceLine。
    行単位の置換のみ（複数行にまたがるパターンは対象外）。
    ReplaceLine 方式なので Attribute 行（ショートカット定義）は壊れない。
    完全一致が0行のときだけ、大小文字→空白の順に寛容に探し直す（VBA の保存時整形の吸収）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    needle = _needle_from_file(getattr(args, 'needle_file', None))
    repl = _needle_from_file(getattr(args, 'repl_file', None))
    if needle is False or repl is False:
        return False
    want = (1 if needle is None else 0) + (1 if repl is None else 0)
    if len(rest) < want:
        print("使い方: code-replace [excel_file] [モジュール] [マクロ] <検索> <置換> [--regex] [--module 名] [-y]"
              "\n  引用符で割れる文字（' \" 等）は --file 検索語.txt / --repl-file 置換語.txt で渡す")
        return False
    # 検索・置換は後ろから取る。前に残った語がモジュール名・マクロ名なら絞り込み（2026-09-24: 会話記録で
    # 「code-replace マクロフォーム VBEコードへジャンプ "旧" "新"」が余分な引数で 2 回落ちていた）
    lead, tail = rest[:len(rest) - want], rest[len(rest) - want:]
    idx = 0
    if needle is None:
        needle, idx = tail[idx], idx + 1
    if repl is None:
        repl, idx = tail[idx], idx + 1
    use_regex = getattr(args, 'regex', False)
    if use_regex:
        try:
            pat = re.compile(needle)
        except re.error as e:
            print(f"エラー: 正規表現が不正です: {e}")
            return False
        rep = repl
    else:
        pat = re.compile(re.escape(needle))
        rep = repl.replace('\\', '\\\\')   # 置換文字列の \ を文字通りに
    mod_filter = getattr(args, 'module_opt', None)

    xl, wb = get_workbook(target_file)
    proc_filter = None
    if lead:
        comps = {c.Name.lower(): c.Name for c in wb.VBProject.VBComponents}
        procs = {p.lower() for p in _all_procedure_names(wb)}
        for t in lead:
            if t.lower() in comps and not mod_filter:
                mod_filter = comps[t.lower()]
            elif t.lower() in procs and not proc_filter:
                proc_filter = t
            else:
                print(f"エラー: 余分な引数があります: {t}（モジュール名でもマクロ名でもありません）")
                print("  スペースを含む場合はクォートで囲んでください（引用符で割れるなら --file / --repl-file）")
                return False
        print("絞り込み: " + " / ".join(x for x in (f"モジュール {mod_filter}" if mod_filter else "",
                                                   f"マクロ {proc_filter}" if proc_filter else "") if x))

    def _in_proc(cm, i):
        if not proc_filter:
            return True
        try:
            p = cm.ProcOfLine(i, 0)
            p = p[0] if isinstance(p, tuple) else p
            return str(p or '').lower() == proc_filter.lower()
        except Exception:
            return False

    # 変更計画の作成（この段階では何も書き換えない）。戻り値 (plans, 変更行数, 一致行数)。
    # 中止すべき不正があれば (None, 0, 0)。一致しても置換後が同じ内容の行は plans に入れない
    def _make_plans(p):
        plans = []       # (comp, [(行番号, 旧行, 新行), ...])
        total = 0
        matched = 0
        for comp in wb.VBProject.VBComponents:
            if mod_filter and comp.Name.lower() != mod_filter.lower():
                continue
            cm = comp.CodeModule
            n = cm.CountOfLines
            if n == 0:
                continue
            changes = []
            for i, line in enumerate(cm.Lines(1, n).split('\r\n'), 1):
                if not p.search(line) or not _in_proc(cm, i):
                    continue
                matched += 1
                try:
                    new_line = p.sub(rep, line)
                except re.error as e:
                    # 置換文字列側の不正（存在しないグループ参照 \1 等）。計画段階なので無傷
                    print(f"エラー: 置換文字列が不正です: {e}")
                    return None, 0, 0
                if new_line != line:
                    # ReplaceLine は行単位。置換結果に改行が入ると1行が複数行になり、
                    # 同一モジュール内の後続の行番号が全部ずれて無関係な行を上書きする
                    if '\r' in new_line or '\n' in new_line:
                        print("エラー: 置換結果に改行が含まれるため中止しました。")
                        print(f"  [{comp.Name}] {i}行目: {new_line.splitlines()[0]} …")
                        print("  （code-replace は行単位置換です。複数行への展開は replace-procedure を使ってください）")
                        return None, 0, 0
                    changes.append((i, line, new_line))
            if changes:
                plans.append((comp, changes))
                total += len(changes)
        return plans, total, matched

    plans, total_lines, matched = _make_plans(pat)
    if plans is None:
        return False
    tolerance = None
    if not use_regex and not plans and not matched:
        # VBA は保存時に識別子の大小文字をプロジェクト内の初出に揃え、演算子の前後の空白も整える。
        # 書いた字面と保存された字面がずれて「マッチなし」になる（2026-08-23: `.Value` と書いたのに
        # `.value` で保存されていて3周空振り）。完全一致が0行のときだけ ①大小文字の違いを無視 →
        # ②空白の違いも無視 の順に探し直し、どの寛容で当てたかを必ず表示する（完全一致があるときは従来どおり）
        for label, p in (("大小文字の違いを無視", re.compile(re.escape(needle), re.IGNORECASE)),
                         ("大小文字と空白の違いを無視", _vba_tolerant_pattern(needle))):
            plans, total_lines, matched = _make_plans(p)
            if plans is None:
                return False
            if plans or matched:
                tolerance = label
                break

    if not plans:
        if matched:
            print(f"'{needle}' は {matched} 行に一致しましたが、置換後も同じ内容です（変更なし）")
        else:
            extra = "" if use_regex else "。大小文字・空白の違いを無視しても一致なし"
            print(f"'{needle}' にマッチする行はありません（置換なし{extra}）")
        return True
    if tolerance:
        print(f"（完全一致なし → {tolerance}して探し直し: {total_lines}行に一致。"
              "VBA は保存時に識別子の大小文字や演算子前後の空白を揃えるため、書いた字面と違うことがあります）")

    # 差分プレビュー
    print(f"--- 置換プレビュー: {len(plans)}モジュール / {total_lines}行 ---")
    for comp, changes in plans:
        print(f"[{comp.Name}] {len(changes)}行:")
        for i, old, new in changes[:20]:
            print(f"  {i}: - {old.strip()}")
            print(f"  {i}: + {new.strip()}")
        if len(changes) > 20:
            print(f"  … 他 {len(changes) - 20}行")
    print("-" * 40)

    if not getattr(args, 'yes', False):
        try:
            ans = input(f"{total_lines}行を置換しますか？ (y/N): ")
        except EOFError:
            # パイプ/MCP 等の非対話環境。トレースバックでなく正常なキャンセルにする
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    if make_backup(wb.FullName, "code_replace") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    for comp, _ in plans:
        make_module_backup(wb, comp.Name)

    # 変更行だけを書き換える
    for comp, changes in plans:
        cm = comp.CodeModule
        for i, _, new_line in changes:
            cm.ReplaceLine(i, new_line)
        print(f"置換: [{comp.Name}] {len(changes)}行")
    wb.Save()
    print(f"完了: {len(plans)}モジュール / {total_lines}行 を置換して保存しました")
    return True


def _start_dialog_watcher(xl, mode=None, input_texts=None):
    """マクロ発火中に出る MsgBox/InputBox(#32770) を検出して自動解除する監視スレッド。

    xl.Application.Run（や書き込みで走るイベントマクロ）は MsgBox が出ると閉じるまで
    ブロックし、別スレッドから閉じない限りコマンドが無言でハングする（2026-07-11 実害:
    write-range→Worksheet_Change→エラーMsgBox で操作が固まり、人が手で閉じるまで戻らず、
    改修が異常に長引いた）。そこで Excel が所有するモーダルダイアログにだけ WM_COMMAND を
    送って解除する。
    mode=None（既定・安全解除）: キャンセル→唯一ボタンの順で「閉じられるボタン」を押す。
      破壊的操作を確定させない方向に倒し、検出した事実と本文を .count / .last_text で残す。
    mode 明示: ok/enter->OK, cancel->キャンセル, yes->はい, no->いいえ。
    input_texts（--input-text）: 入力欄を持つダイアログ（VBA の InputBox() / Excel 内蔵の
      Application.InputBox）に入れる値のリスト。出た順に1枚1つ使い、足りなくなったら
      最後の値を使い回す。値を入れたら OK（内蔵 InputBox は Enter）で確定する＝
      「閉じる」でなく「答えて先へ進ませる」。mode より優先。指定が無ければ入力欄つきでも
      従来どおり閉じるだけ。入れた値は .inputs に残す（黙って入れない）。
    返り値の .stop() で終了。.count=解除した回数 / .last_text=最後に見た本文。
    """
    import threading
    try:
        import win32gui
        import win32process
        import win32con
    except Exception as ex:
        print(f"[WARNING] ダイアログ監視を開始できません（win32 不足）: {ex}", file=sys.stderr)

        class _Noop:
            def stop(self):
                pass
        return _Noop()

    excel_pid = None
    try:
        excel_pid = win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1]
    except Exception:
        pass
    if excel_pid is None:
        # PID を特定できないまま監視すると、フィルタが外れて画面上の
        # 全アプリの #32770 ダイアログにボタンを送ってしまう。
        # フェイルセーフは「撃たない」側に倒す（自動応答は諦めて警告）
        print("[WARNING] Excel の PID を特定できないため、ダイアログ自動応答を無効化します"
              "（ダイアログが出た場合は手動で閉じてください）", file=sys.stderr)

        class _Noop:
            def stop(self):
                pass
        return _Noop()

    mode_l = (mode or 'safe').lower()
    # 標準ボタンID（Win32 DialogBox の既定値）。決め打ちの「望みのID」。
    _STD_ID = {'ok': 1, 'enter': 1, 'cancel': 2, 'yes': 6, 'no': 7,
               'abort': 3, 'retry': 4, 'ignore': 5, 'safe': 2}
    # テキストで拾う第2の網（OK専用MsgBoxのID化け対策・日英両対応）
    _TEXT_HINTS = {
        'ok':     ('ok', 'はい', '確定', '了解'),
        'enter':  ('ok', 'はい', '確定', '了解'),
        'cancel': ('cancel', 'キャンセル', '中止'),
        'yes':    ('yes', 'はい'),
        'no':     ('no', 'いいえ'),
    }
    want_id = _STD_ID.get(mode_l, 1)
    hints = _TEXT_HINTS.get(mode_l, ())

    def _dialog_buttons(hwnd):
        """ダイアログ内の Button コントロールを [(ctrl_id, text)] で返す。"""
        found = []

        def _child(ch, _):
            try:
                # 押せないボタンは数えない（VBA の実行時エラーの窓は先頭が灰色の［継続］で、
                # それを押し続けて 600 秒固まった・2026-09-23）
                if win32gui.GetClassName(ch) == 'Button' and win32gui.IsWindowEnabled(ch):
                    cid = win32gui.GetDlgCtrlID(ch)
                    txt = win32gui.GetWindowText(ch).replace('&', '').strip()
                    found.append((cid, txt))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _child, None)
        except Exception:
            pass
        return found

    def _dialog_text(hwnd):
        """ダイアログのタイトル＋本文（Static コントロールの文字）を採取して報告に使う。

        タイトルは「どのダイアログが開いたか」の一番強い証拠
        （例: セルの書式設定 / 検索と置換）。MsgBox 系はタイトルが
        「Microsoft Excel」等で情報が薄いので本文と併記する。
        """
        parts = []

        def _child(ch, _):
            try:
                if win32gui.GetClassName(ch) == 'Static':
                    t = win32gui.GetWindowText(ch).strip()
                    if t and t not in parts:
                        parts.append(t)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _child, None)
        except Exception:
            pass
        title = ''
        try:
            title = win32gui.GetWindowText(hwnd).strip()
        except Exception:
            pass
        body = ' / '.join(parts)
        if title:
            return f"タイトル「{title}」 本文: {body}" if body else f"タイトル「{title}」"
        return body

    def _dialog_edit(hwnd):
        """ダイアログ内の入力欄 (hwnd, クラス名) を返す（無ければ None）。

        VBA の InputBox() は Win32 標準の Edit、Excel 内蔵の Application.InputBox は
        EDTBX（Excel 独自の入力欄）。これが有る＝値を待っているダイアログ。
        """
        found = []

        def _child(ch, _):
            try:
                cls = win32gui.GetClassName(ch)
                if cls in ('Edit', 'EDTBX'):
                    found.append((ch, cls))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _child, None)
        except Exception:
            pass
        return found[0] if found else None

    def _set_edit_text(h, val):
        """入力欄に値を入れる。WM_SETTEXT で入ったか読み返して確かめ、入らない欄
        （Excel 独自の EDTBX 等）は全選択→1文字ずつ WM_CHAR で打ち込む。"""
        import ctypes
        u32 = ctypes.WinDLL('user32')            # 共有の windll に argtypes を残さない
        u32.SendMessageW.restype = ctypes.c_ssize_t
        u32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                     ctypes.c_size_t, ctypes.c_void_p]
        WM_SETTEXT, WM_GETTEXT, WM_GETTEXTLENGTH = 0x000C, 0x000D, 0x000E
        WM_CHAR, EM_SETSEL = 0x0102, 0x00B1
        try:
            u32.SendMessageW(h, WM_SETTEXT, 0, val)
            n = u32.SendMessageW(h, WM_GETTEXTLENGTH, 0, None)
            buf = ctypes.create_unicode_buffer(int(n) + 1)
            u32.SendMessageW(h, WM_GETTEXT, int(n) + 1, buf)
            if buf.value == val:
                return
        except Exception:
            pass
        try:
            win32gui.SendMessage(h, EM_SETSEL, 0, -1)
        except Exception:
            pass
        for ch in val:
            win32gui.PostMessage(h, WM_CHAR, ord(ch), 0)

    def _resolve_button_id(hwnd, mode_override=None):
        """このダイアログで実際に押すべきボタンIDを、実在ボタンから決める。

        OK のみの MsgBox は OK ボタンの ID が 2(IDCANCEL) になる Windows の仕様があり、
        標準IDを決め打ちで送ると閉じない（フェイブルが実弾で特定）。実在ボタンの
        ID とテキストを見て決めることで、OK専用MsgBox・Yes/No・InputBox すべてに効かせる。
        mode_override: 値を入れた InputBox を OK で確定するときだけ 'ok' を渡す。
        """
        _mode = (mode_override or mode_l).lower()
        _want = _STD_ID.get(_mode, 1) if mode_override else want_id
        _hints = _TEXT_HINTS.get(_mode, ()) if mode_override else hints
        btns = _dialog_buttons(hwnd)
        if not btns:
            return _want                         # 取れなければ従来の決め打ち
        ids = [cid for cid, _ in btns]
        # 安全解除モード（既定）: 破壊確定を避けつつ「閉じられるボタン」を必ず1つ選ぶ
        if _mode == 'safe':
            if 2 in ids:                         # キャンセル(IDCANCEL) があれば最優先
                return 2
            for cid, txt in btns:                # 文字でキャンセル/いいえ系
                tl = txt.lower()
                if any(h in tl for h in ('cancel', 'キャンセル', '中止', 'いいえ', 'no')):
                    return cid
            for cid, txt in btns:                # VBA の実行時エラー＝［終了］でマクロを止める
                tl = txt.lower()
                if '終了' in tl or tl.startswith('end'):
                    return cid
            for cid, txt in btns:                # デバッグ（中断モードに入る）とヘルプは押さない
                tl = txt.lower()
                if not any(h in tl for h in ('デバッグ', 'debug', 'ヘルプ', 'help')):
                    return cid
            return btns[0][0]                    # OK専用等はその1つで閉じる
        # 1) 望む標準IDが実在すればそれ（通常の OK+キャンセル・Yes/No 等）
        if _want in ids:
            return _want
        # 2) テキストで一致するボタン（IDが化けていても文字で拾う）
        for cid, txt in btns:
            tl = txt.lower()
            if any(h in tl for h in _hints):
                return cid
        # 3) OK系でボタンが1つだけ＝OK専用MsgBox（ID2化け）→ そのボタンを押す
        if _mode in ('ok', 'enter') and len(btns) == 1:
            return btns[0][0]
        # 4) それでも決まらなければ決め打ちに戻す
        return _want

    stop_evt = threading.Event()
    # count は「窓が実際に消えたことを確認した数」。PostMessage は非同期で、
    # 送っただけでは閉じたことにならない（ボタンID解決が外れる／WM_CLOSE を
    # 無視するダイアログでは閉じない）。従来は送信した時点で数えて「解除しました」と
    # 報告し、その hwnd を二度と再送しなかったため、安全弁が仕事をせずハングし続けた。
    state = {'count': 0, 'last': '', 'unclosed': [], 'inputs': []}
    pending = {}                           # hwnd -> {tries, last_ts, text[, value]}
    _RESEND_SEC = 0.6                      # 消えなければこの間隔で再送する
    _texts = [str(t) for t in (input_texts or [])]
    _next_idx = [0]                        # 次の InputBox に入れる値の位置（尽きたら最後を使い回す）

    def _enum_targets():
        targets = []

        def _cb(hwnd, _unused):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                # #32770=MsgBox/InputBox等の標準ダイアログ、
                # bosa_sdm_XL9=Excel内蔵ダイアログ(xlDialog*/セルの書式設定等)、
                # NUIDialog=Office描画ダイアログ。いずれもExcel PID所有のモーダルとして
                # 放置するとxl.Runが無言ハングする(2026-07-12実害: xlDialogFormatNumberを
                # 開くマクロでrun-macroが40秒超ブロック、#32770限定だったため素通り)
                if win32gui.GetClassName(hwnd) not in (
                        '#32770', 'bosa_sdm_XL9', 'NUIDialog'):
                    return
                if excel_pid is not None:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid != excel_pid:
                        return
                targets.append(hwnd)
            except Exception:
                pass

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        return targets

    def _settle(targets):
        # 消えた hwnd ＝ 本当に閉じられたもの。ここで初めて「解除した」と数える
        for hwnd in [h for h in pending if h not in targets]:
            info = pending.pop(hwnd)
            state['count'] += 1
            if info.get('text'):
                state['last'] = info['text']
            if 'value' in info:
                state['inputs'].append(info['value'])

    def _loop():
        while not stop_evt.is_set():
            targets = _enum_targets()
            _settle(targets)
            now = time.time()
            for hwnd in targets:
                info = pending.get(hwnd)
                if info is not None and now - info['last_ts'] < _RESEND_SEC:
                    continue                       # 送ったばかり。反応を待つ
                try:
                    txt = _dialog_text(hwnd)
                    cls = win32gui.GetClassName(hwnd)
                    edit = _dialog_edit(hwnd) if _texts else None
                    val = None
                    if edit is not None:
                        # 入力欄つき（InputBox 系）＝値を入れて確定する。値は窓1枚に1つ
                        # （再送でも同じ値。次の値は次の窓へ）
                        if info is not None and 'value' in info:
                            val = info['value']
                        else:
                            val = _texts[min(_next_idx[0], len(_texts) - 1)]
                            _next_idx[0] += 1
                        _set_edit_text(edit[0], val)
                        if cls == '#32770':
                            cid = _resolve_button_id(hwnd, 'ok')
                            win32gui.PostMessage(hwnd, win32con.WM_COMMAND, cid, 0)
                        else:
                            # Excel内蔵 InputBox は子が Win32 Button でない → 入力欄に Enter で確定
                            win32gui.PostMessage(edit[0], win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                            win32gui.PostMessage(edit[0], win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    elif cls == '#32770':
                        cid = _resolve_button_id(hwnd)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, cid, 0)
                    else:
                        # Excel内蔵/Office描画ダイアログは子がWin32 Buttonでないため
                        # ボタンID解決が効かない。WM_CLOSE(=×ボタン)がキャンセル相当
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    if info is None:
                        pending[hwnd] = {'tries': 1, 'last_ts': now, 'text': txt}
                        if val is not None:
                            pending[hwnd]['value'] = val
                    else:
                        info['tries'] += 1
                        info['last_ts'] = now
                        if txt:
                            info['text'] = txt
                except Exception:
                    pass
            stop_evt.wait(0.15)
        # 監視終了時点：最後にもう一度見て、消えたものは「閉じた」に数える。
        # Run が戻った直後に stop() されると、直前に閉じたダイアログを未確認のまま
        # 「閉じられなかった」と誤報していた（2026-08-23 --input-text の実機で露見）
        _settle(_enum_targets())
        # それでも残っているもの＝送っても閉じなかったダイアログ。
        # 「解除しました」と混ぜず、閉じられなかった事実として別に報告する
        for info in pending.values():
            state['unclosed'].append(info.get('text') or '(本文なし)')

    th = threading.Thread(target=_loop, daemon=True)
    th.start()

    class _Watcher:
        def stop(self):
            stop_evt.set()
            # ループが抜けきるのを待つ。待たずに戻ると、直後に unclosed を読む
            # 呼び出し側が「閉じられなかったダイアログ」を取りこぼす（競合）
            try:
                th.join(timeout=1.0)
            except Exception:
                pass

        @property
        def count(self):
            return state['count']

        @property
        def last_text(self):
            return state['last']

        @property
        def unclosed(self):
            """送っても閉じなかったダイアログの本文（安全弁が効かなかった証拠）"""
            return state['unclosed']

        @property
        def inputs(self):
            """InputBox 系に入れて確定した値（出た順）。黙って入れない＝報告の素"""
            return state['inputs']

    return _Watcher()


def _dialog_watcher_note(watcher, mode):
    """ダイアログを自動解除したときの人向けの注記（無ければ空文字）。

    「黙って握りつぶした」にしないための報告。何が出て、どう閉じたかを一行で残す。
    """
    if watcher is None:
        return ""
    if not getattr(watcher, 'count', 0):
        # 1件も閉じられていなくても、閉じられずに残ったダイアログがあれば報告する
        stuck0 = list(getattr(watcher, 'unclosed', ()) or ())
        if stuck0:
            return ("⚠ 実行中にダイアログを検出しましたが、閉じられませんでした"
                    f"（{len(stuck0)}件・Excel 側に残っている可能性があります）: "
                    + " / ".join(s[:60] for s in stuck0[:3]))
        return ""
    how = "指定ボタンで応答" if mode else "安全側（キャンセル優先）で自動解除"
    body = watcher.last_text or "(本文なし)"
    n = watcher.count
    # 値を入れて確定した InputBox は「何を入れたか」を必ず出す（黙って入れない）
    entered = list(getattr(watcher, 'inputs', ()) or ())
    if entered and len(entered) >= n:
        how = "入力欄に --input-text の値を入れて確定"
    msg = (f"⚠ 実行中にダイアログを{n}件検出し、{how}しました。"
           f"マクロがメッセージを出しています → 内容: {body}")
    if entered:
        head = "入れた値（出た順）" if len(entered) >= n else f"うち {len(entered)}件は入力欄に --input-text の値を入れて確定"
        msg += f"\n  {head}: " + " / ".join(f"「{v}」" for v in entered[:5])
    # 送っても閉じなかったもの＝安全弁が効かなかった証拠。「解除しました」と
    # 混ぜず、閉じられなかった事実として必ず別に出す（黙って成功にしない）
    stuck = list(getattr(watcher, 'unclosed', ()) or ())
    if stuck:
        msg += (f"\n⚠ うち {len(stuck)}件は閉じられませんでした"
                "（Excel 側に残っている可能性があります）: "
                + " / ".join(s[:60] for s in stuck[:3]))
    return msg


def dialog_safe(cmd_func):
    """cmd_* を「実行中のダイアログを安全側で自動解除する」に変える decorator。

    セルの書き換え・クリア・行列削除・並べ替え・置換は Worksheet_Change 等の
    イベントマクロを同期発火させる。そのマクロが MsgBox を出すと COM 呼び出しが
    そこでブロックし、人が手で閉じるまでコマンドが無言でハングする（2026-07-11 実害。
    write-range と run-macro では常設済みだったが、他の書き込み系には無かった）。
    検出したダイアログは終了後に必ず報告する（黙って握りつぶさない）。
    """
    import functools

    @functools.wraps(cmd_func)
    def wrapper(args):
        try:
            target_file, _rest = parse_target_and_rest(getattr(args, 'posargs', []) or [])
            xl, _wb = get_workbook(target_file)
        except Exception:
            # ブック解決に失敗した場合は元の関数に委ね、そちらのエラーを出させる
            return cmd_func(args)
        watcher = _start_dialog_watcher(xl)
        try:
            return cmd_func(args)
        finally:
            try:
                watcher.stop()
            except Exception:
                pass
            note = _dialog_watcher_note(watcher, None)
            if note:
                print(note, file=sys.stderr)
    return wrapper


def _project_book_name(xl, p):
    """VBProject から Application.Run 修飾に使うブック名を得る。

    p.Filename は未保存ブックだと例外／空になる。その場合は開いているブックを
    走査して同じプロジェクトのブック名（Book1 等）を拾う。Run はブック名で修飾
    できるので、未保存ブックでも取り違えずに名指しできる。見つからなければ None。
    """
    try:
        fn = p.Filename
    except Exception:
        fn = None
    if fn:
        return os.path.basename(fn)
    try:
        pname = p.Name
        comps = sorted(str(c.Name) for c in p.VBComponents)
    except Exception:
        return None
    try:
        for w in xl.Workbooks:
            try:
                # 名前（VBAProject）だけで当てると、ファイルの無い抜け殻のプロジェクト（アドインの更新登録の残り）が
                # 未保存の Book20 と取り違えられ、棚を Book20 で撃って「マクロを実行できません」（2026-09-24）。モジュールの顔ぶれも見る
                if (not w.Path and w.VBProject.Name == pname
                        and sorted(str(c.Name) for c in w.VBProject.VBComponents) == comps):
                    return w.Name
            except Exception:
                continue
    except Exception:
        pass
    return None


def _find_macro_owner_books(xl, wb, macro_name):
    """macro_name を宣言しているプロジェクトを全部探す。

    戻り値: (対象ブック名 or None, 他候補[(ブック名 or None, プロジェクト名)])

    VBProjects の列挙順で先勝ちすると、アドインや PERSONAL.XLSB に同名 Sub が
    あるとき作業ブックではなくそちら側を実行してしまう。対象ブック（get_workbook が
    返した wb）に有ればそれを最優先し、他にも同名があれば警告できるよう全部返す。
    """
    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+' + re.escape(macro_name) + r'\b',
        re.IGNORECASE | re.MULTILINE
    )
    try:
        target_name = wb.Name
    except Exception:
        target_name = None

    target_hit = None
    others = []
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception as ex:
        print(f"[DEBUG] Failed to enumerate VBProjects: {ex}", file=sys.stderr)
        return None, []

    for p in projects:
        hit = False
        try:
            for comp in p.VBComponents:
                try:
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0 and pattern.search(cm.Lines(1, cm.CountOfLines)):
                        hit = True
                        break
                except Exception:
                    # 読めないモジュールは飛ばすが、探索自体は止めない
                    continue
        except Exception:
            continue
        if not hit:
            continue
        bname = _project_book_name(xl, p)
        try:
            pname = p.Name
        except Exception:
            pname = '?'
        if bname and target_name and bname.lower() == target_name.lower():
            target_hit = bname
        else:
            others.append((bname, pname))
    return target_hit, others


def _start_call_watchdog(seconds, owned_pid=None):
    """COM 呼び出し（Application.Run 等）の待ちに制限時間を付ける見張りスレッド（2026-08-23）。

    seconds 経過で:
      owned_pid あり（rehearse / gate の演習用 Excel）: その EXCEL.EXE を強制終了する。
        固まっていた呼び出しは RPC エラーで即戻る（実測 3 秒）。ユーザーの Excel には触れない。
      owned_pid なし（ユーザーの生きている Excel）: CoCancelCall で「待つのをやめる」だけ。
        マクロは止まらない（外から止める手段が無い）。VBA が DoEvents 等で応答する状態なら
        呼び出しは RPC_E_CALL_CANCELED で戻るが、`Do: Loop` の密なループは Excel が一切
        応答しないので効かない（実測）。戻った後もマクロは Excel 側で動き続けている可能性が
        ある＝呼び出し側は報告に必ずそう書く。
    seconds が None / 0 以下なら何もしない（.fired は常に False）。
    返り値: .stop()（見張りを畳む・必ず呼ぶ）/ .fired（時間切れで撃ったか）/ .seconds
    """
    import threading

    class _Noop:
        fired = False
        seconds = None

        def stop(self):
            pass

    try:
        sec = float(seconds) if seconds is not None else 0.0
    except (TypeError, ValueError):
        sec = 0.0
    if sec <= 0:
        return _Noop()

    state = {'fired': False}
    cancel = None          # (ole32, tid) … ユーザーの Excel 向け（呼び出しの取り消し）
    if owned_pid is None:
        try:
            import ctypes
            ole32 = ctypes.windll.ole32
            tid = ctypes.windll.kernel32.GetCurrentThreadId()
            if ole32.CoEnableCallCancellation(None) == 0:
                cancel = (ole32, tid)
        except Exception:
            cancel = None
    stop_evt = threading.Event()

    def _fire():
        if stop_evt.wait(sec):
            return                          # 時間内に終わった
        state['fired'] = True
        if owned_pid is not None:
            # 撃つ前に「その PID が今も EXCEL.EXE か」を確認する（PID 再利用の誤射防止）
            import signal
            import vbam_core as _vc
            if _vc._pid_is_excel(owned_pid):
                try:
                    os.kill(owned_pid, signal.SIGTERM)
                except Exception:
                    pass
        elif cancel is not None:
            try:
                cancel[0].CoCancelCall(cancel[1], 0)
            except Exception:
                pass

    th = threading.Thread(target=_fire, daemon=True)
    th.start()

    class _Watchdog:
        seconds = sec

        @property
        def fired(self):
            return state['fired']

        def stop(self):
            stop_evt.set()
            try:
                th.join(timeout=1.0)
            except Exception:
                pass
            if cancel is not None:
                try:
                    cancel[0].CoDisableCallCancellation(None)
                except Exception:
                    pass

    return _Watchdog()


_ADDIN_REGISTER_MACRO = 'アドインの更新登録'


def _find_addin(xl, file_name):
    """AddIns コレクションから .xlam を名前で引く（Item("名前.xlam") は引けずに落ちる）。"""
    try:
        for a in xl.AddIns:
            try:
                if str(a.Name).lower() == file_name.lower():
                    return a
            except Exception:
                continue
    except Exception:
        pass
    return None


def _addin_project_loaded(xl, addin_path):
    """その .xlam の VBA プロジェクトが今の Excel に載っているか。

    登録一覧の Installed は True のままでも、プロジェクトだけ外れていることがある
    （読み込み中の .xlam を上書きしたとき）。こちらが本当の可否。
    """
    want = os.path.normcase(os.path.abspath(addin_path))
    try:
        for p in xl.VBE.VBProjects:
            try:
                fn = str(p.FileName or '')
                if fn and os.path.normcase(os.path.abspath(fn)) == want:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def cmd_register_addin(args):
    """`register-addin [ブック]`: 開いているブックを .xlam に焼き直してアドインに登録し直す（2026-09-23）。

    「アドインの更新登録」は **ActiveWorkbook** を xlam にするマクロなので、撃つ前に道具が三つ確かめる:
      ・焼く先が本当にそのブックか（前に出ているブックを名指しで報告する）
      ・ほかに空のブック（保存していない Book1 など）が開いていないか＝間違ってそれを焼かない
      ・表示中のフォームが無いか（Designer を取れずマクロが MsgBox で止まる）→ 出ていれば閉じてから撃つ
    撃つのは**素の Run**（ハーネスの一時モジュールを置くと、そのまま xlam に焼き付く）。
    撃った後に .xlam の日時と大きさを読んで、本当に焼けたかを報告する。

    **焼く前にアドインを外し、焼いた後に入れ直す。** 読み込み中の .xlam を上書きすると、
    登録一覧は Installed=True のままなのに VBA プロジェクトだけ外れ、どのブックからも
    アドインのマクロが使えなくなる（2026-09-23 に実害）。最後に VBProjects を見て、
    本当に載っているかまで確かめる。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, load_addins=True)
    book = str(wb.Name)
    stem, ext = os.path.splitext(book)
    if ext.lower() not in ('.xlsm', '.xlsb'):
        print(f"エラー: {book} はマクロを持てる形式ではありません（.xlsm / .xlsb を前に出してください）")
        return False
    blanks = []
    for other in xl.Workbooks:
        try:
            if str(other.Name) != book and not str(other.Path or ''):
                blanks.append(str(other.Name))
        except Exception:
            continue
    if blanks and not getattr(args, 'yes', False):
        print(f"エラー: 保存していない空のブックが開いています: {'・'.join(blanks)}")
        print("  更新登録は前に出ているブックを xlam にします。間違って空のブックを焼かないよう、"
              "先に閉じてください（それでも撃つなら -y）")
        return False
    try:
        from vbam_edit import _list_form_windows, cmd_close_form   # 遅延 import（vbam_edit はこちらを読む側）
        shown = _list_form_windows()
    except Exception:
        shown = []
    if shown:
        # 表示中のフォームがあると Designer を取れず、マクロが MsgBox を出して止まる。人に頼まず自分で閉じる
        print(f"表示中のフォームを閉じます（{len(shown)} 件）")
        cmd_close_form(argparse.Namespace(posargs=[], list_flag=False))
    addin_file = f"{stem}.xlam"
    addin_path = os.path.join(os.environ.get('AppData', ''), 'Microsoft', 'AddIns', addin_file)
    before = os.path.getmtime(addin_path) if os.path.isfile(addin_path) else 0
    print(f"更新登録: {book} → {addin_path}")
    # 読み込み中の .xlam をそのまま上書きすると、登録一覧は Installed=True のままなのに
    # VBA プロジェクトだけ Excel から外れる＝どのブックからもアドインのマクロが使えなくなる。
    # 焼く前に外し、焼いた後に入れ直す（2026-09-23 に実害）。
    unloaded = False
    addin = _find_addin(xl, addin_file)
    if addin is not None:
        try:
            if bool(addin.Installed):
                addin.Installed = False
                unloaded = True
                print(f"  焼く前にアドインを外しました: {addin_file}")
        except Exception as ex:
            print(f"  ⚠ 外せませんでした（このまま焼きます）: {ex}")
    try:
        try:
            wb.Activate()
            q = book.replace("'", "''")
            xl.Application.Run(f"'{q}'!{_ADDIN_REGISTER_MACRO}")    # 素の Run（ハーネスを置かない）
        except Exception as ex:
            print(f"エラー: 更新登録を撃てませんでした: {ex}")
            return False
        if not os.path.isfile(addin_path):
            print(f"エラー: .xlam ができていません（{addin_path}）")
            return False
        after = os.path.getmtime(addin_path)
        when = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(after))
        if after <= before:
            print(f"⚠ .xlam の日時が変わっていません（{when}）。フォームが開いていた・別のブックが前に出ていた等で"
                  "途中で止まったかもしれません")
            return False
        print(f"焼けました: {addin_path}  {os.path.getsize(addin_path):,} バイト  {when}")
        return True
    finally:
        if unloaded:
            back = _find_addin(xl, addin_file)      # 焼き直しでコレクションが入れ替わることがある
            try:
                if back is not None:
                    back.Installed = True
                    print(f"  アドインを入れ直しました: {addin_file}")
            except Exception as ex:
                print(f"⚠ アドインを入れ直せませんでした: {ex}")
        if _addin_project_loaded(xl, addin_path):
            print(f"読み込み確認: {addin_file} は今の Excel に載っています"
                  "（このままどのブックからもショートカットが効きます）")
        else:
            print(f"⚠ {addin_file} が今の Excel に載っていません。Excel を開き直すか、"
                  "［ファイル］→［オプション］→［アドイン］でチェックを入れ直してください")
        # 抜け殻（ファイルとのつながりが切れた VBA プロジェクト）を数える（2026-09-23 のゾンビ）。
        # 古いアドインが閉じきれずに残ると、ショートカットやボタンがそちらへ向かう
        ghosts = _ghost_projects(xl)
        if ghosts:
            print(f"⚠ 抜け殻の VBA プロジェクトが {ghosts} 個残っています（ファイル名なし）。"
                  "Excel を閉じて開き直すと消えます（閉じても EXCEL.EXE が残るなら道具の常駐が参照を握っている）")
        else:
            print("抜け殻: なし（古いアドインはきれいに閉じました）")


def _ghost_projects(xl):
    """ファイルとのつながりが切れた VBA プロジェクト（閉じたのに Excel に残った抜け殻）の数。読めなければ 0。"""
    n = 0
    try:
        for p in xl.VBE.VBProjects:
            try:
                if not str(p.FileName or ''):
                    n += 1
            except Exception:
                n += 1
    except Exception:
        return 0
    try:                                  # 保存前の新しいブック（Book1 等）もファイル名が無い＝抜け殻ではない
        n -= sum(1 for b in xl.Workbooks if not str(b.Path or ''))
    except Exception:
        pass
    return max(n, 0)


def cmd_run_macro(args):
    """Excelマクロを実行する"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: run-macro [excel_file] <macro_name>", file=sys.stderr)
        return False

    macro_name = rest[0]

    # 実行前にアドインや個人用マクロをロードする
    xl, wb = get_workbook(target_file, load_addins=True)

    # 警告を非表示にする
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    full_macro_path = macro_name

    # マクロ名に "!" が含まれていない場合、どのブックにあるか検索する
    if "!" not in macro_name:
        # 対象ブック（アクティブブック or --target）を最優先で探す。
        # 列挙順の先勝ちだと、アドイン/PERSONAL 側の同名マクロを実行してしまう
        target_hit, others = _find_macro_owner_books(xl, wb, macro_name)
        found_wb = target_hit
        if target_hit:
            if others:
                dup = '、'.join(
                    (b or f'{pn}（ブック名不明）') for b, pn in others
                )
                print(f"⚠ 同名マクロが他にもあります: {dup} → 対象ブック "
                      f"'{target_hit}' 側を実行します", file=sys.stderr)
        elif others:
            named = [(b, pn) for b, pn in others if b]
            if named:
                found_wb = named[0][0]
                if len(others) > 1:
                    dup = '、'.join(
                        (b or f'{pn}（ブック名不明）') for b, pn in others
                    )
                    print(f"⚠ 同名マクロが複数のブックにあります: {dup} → "
                          f"'{found_wb}' 側を実行します。意図が違う場合は "
                          f"\"ブック名!マクロ名\" で名指ししてください", file=sys.stderr)
            else:
                # ブック名で修飾できない（プロジェクトの帰属が取れない）ときだけ
                # 修飾なしの直接実行に任せる。黙って別ブックを掴まない
                print(f"[WARNING] Macro '{macro_name}' はブック名を特定できない"
                      f"プロジェクトにあります。修飾なしで実行します。", file=sys.stderr)

        if found_wb:
            # ブック名に空白等があると Application.Run はクォート必須
            # （cmd_test と同じ流儀。' 自体は Excel 規約どおり '' に重ねる）
            quoted_wb = found_wb.replace("'", "''")
            full_macro_path = f"'{quoted_wb}'!{macro_name}"
            print(f"[DEBUG] Macro found in: {found_wb}", file=sys.stderr)
        else:
            print(f"[WARNING] Macro '{macro_name}' not found in open projects. Trying direct run.", file=sys.stderr)

    # 引数（rest[1:]）。数値に見えるものは数値化して渡す（Excel MCP の run 相当）
    run_args = []
    for a in rest[1:]:
        v = _coerce_cell(a)
        run_args.append(a if v is None else v)
    if run_args:
        print(f"マクロ実行中: {full_macro_path}  引数: {run_args}", file=sys.stderr)
    else:
        print(f"マクロ実行中: {full_macro_path}", file=sys.stderr)

    # ダイアログ対策は既定で常設。--auto-dialog 明示時はそのボタンで応答、
    # 省略時は安全解除（キャンセル優先）で「無言ハング」を必ず断ち切る。
    # --input-text は InputBox 系に値を入れて OK（閉じるのでなく答えて先へ進ませる）。
    _auto_dialog = getattr(args, 'auto_dialog', None)
    _input_texts = getattr(args, 'input_text', None)

    # ハーネス経由（既定）：マクロと同じプロジェクトに一時モジュールを置き、
    # On Error で受けて "OK|戻り値" / "ERR|番号|説明" を持ち帰る。
    # 実行時エラーが VBA のモーダルダイアログにならず、番号と説明が手元に残る。
    # Application.Run 越しの呼び出しはエラーが呼び元に伝播しない（実測 2026-08-23）ので、
    # ハーネスは必ずマクロの持ち主ブックに置いて直接呼ぶ。引数付き・--raw・持ち主不明は素の Run。
    harness = None
    if not getattr(args, 'raw', False) and not run_args:
        harness = _prepare_run_harness(xl, full_macro_path, macro_name)
        if harness is None:
            print("[DEBUG] ハーネスを置けないので素の Run で実行します", file=sys.stderr)

    _dlg_watcher = _start_dialog_watcher(xl, _auto_dialog, _input_texts)
    # --timeout: 制限時間で待つのをやめる（ユーザーの Excel は殺せないので CoCancelCall。
    # VBA が応答する状態でだけ効き、マクロ自体は止まらない＝報告に明記する）
    _watchdog = _start_call_watchdog(getattr(args, 'timeout', None))

    try:
        if harness is not None:
            ret = xl.Application.Run(harness["entry"])
            _watchdog.stop()
            ret = "" if ret is None else str(ret)
            if ret.startswith("OK|"):
                result = ret[3:] or None
            elif ret.startswith("ERR|"):
                parts = ret.split("|", 2)
                num = parts[1] if len(parts) > 1 else "?"
                desc = parts[2] if len(parts) > 2 else ""
                raise _VbaRuntimeError(num, desc)
            else:
                result = ret
        else:
            result = xl.Application.Run(full_macro_path, *run_args)
            _watchdog.stop()

        # 報告を組む前に監視を止める（stop は最後にもう一度見て「閉じた」を確定する）。
        # 止める前に読むと、Run が戻る直前に閉じた最後のダイアログが数え漏れる
        # （2026-08-23 実測: InputBox 2枚のうち2枚目が報告から落ちた）
        _dlg_watcher.stop()
        _dlg_note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
        if getattr(args, 'json', False):
            import json
            out = {"success": True, "macro": full_macro_path, "result": str(result)}
            if _dlg_watcher.count:
                out["dialogs_dismissed"] = _dlg_watcher.count
                out["dialog_text"] = _dlg_watcher.last_text
                _entered = list(getattr(_dlg_watcher, 'inputs', ()) or ())
                if _entered:
                    out["inputs_entered"] = _entered
            print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        else:
            print(f"マクロ実行成功。戻り値: {result}")
            if _dlg_note:
                print(_dlg_note, file=sys.stderr)
        return True
    except _VbaRuntimeError as e:
        err_msg = f"実行時エラー {e.number}: {e.description}"
        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": False, "macro": full_macro_path, "error": err_msg,
                              "vba_error": {"number": e.number, "description": e.description}},
                             ensure_ascii=False), file=sys.stdout)
        else:
            print(f"エラー: マクロが実行時エラーで止まりました: {err_msg}", file=sys.stderr)
            print("（ハーネスで受けたのでダイアログは出ていません。マクロは落ちた行で終わっています）",
                  file=sys.stderr)
        return False
    except Exception as e:
        err_msg = _com_error_text(e)
        _timed_out = bool(_watchdog.fired)
        if _timed_out:
            err_msg = (f"時間切れ（{_watchdog.seconds:g}秒）: 待つのをやめました。"
                       "マクロは Excel 側でまだ動いている可能性があります（Esc か vbe-reset で止めてください）")
        if getattr(args, 'json', False):
            import json
            out = {"success": False, "macro": full_macro_path, "error": err_msg}
            if _timed_out:
                out["timeout"] = True
            print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        else:
            print(f"エラー: マクロの実行に失敗しました: {err_msg}", file=sys.stderr)
        return False
    finally:
        _watchdog.stop()
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
        if harness is not None:
            _remove_run_harness(harness)
        # ツール側では切っていないが、実行したマクロが DisplayAlerts=False を立てたまま
        # 落ちている場合がある。そのままだとユーザーの Excel セッションに残り、以後の
        # 手動操作で保存確認などの警告が出なくなるため、必ず有効に戻す
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass


class _VbaRuntimeError(Exception):
    """ハーネスが受け止めた VBA の実行時エラー（番号と説明つき）"""
    def __init__(self, number, description):
        super().__init__(f"実行時エラー {number}: {description}")
        self.number = number
        self.description = description


_RUN_HARNESS = "VbaManagerRunHarness"


def _prepare_run_harness(xl, full_macro_path, macro_name):
    """run-macro 用ハーネスをマクロの持ち主ブックに注入する。

    戻り値: {"wb": 持ち主ブック, "comp": 注入モジュール, "entry": 呼び出し名, "was_saved": bool}
            置けないとき None（呼び元は素の Run に落とす）。
    full_macro_path は "'ブック名'!マクロ" か素の名前。ハーネスは同じプロジェクトに置き、
    マクロを直接（Application.Run を挟まず）呼ぶ。Sub/Function の別は宣言行を見て決める
    （Sub に `r = Sub名` はコンパイルエラー、Function を文として呼ぶと戻り値が捨たるため）。
    """
    # 持ち主ブック名とマクロ式（"Module.Proc" / "Proc"）を分ける
    book_name = None
    expr = macro_name
    if "!" in full_macro_path:
        left, right = full_macro_path.rsplit("!", 1)
        book_name = left.strip().strip("'").replace("''", "'")
        expr = right.strip()
    if not book_name:
        return None
    try:
        owner = xl.Workbooks(book_name)
    except Exception:
        return None
    proc = expr.rsplit(".", 1)[-1]
    mod_only = expr.rsplit(".", 1)[0] if "." in expr else None
    # 生成コードに埋めるのは識別子だけ（ユーザー入力をそのまま VBA にしない）。
    # 形が識別子でない／モジュール名が実在しない／宣言が見つからない → ハーネスを置かず素の Run
    if check_vba_identifier(proc) or (mod_only is not None and check_vba_identifier(mod_only)):
        return None
    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(Sub|Function)\s+' + re.escape(proc) + r'\b',
        re.IGNORECASE | re.MULTILINE
    )
    kind = None
    try:
        for comp in owner.VBProject.VBComponents:
            if mod_only and comp.Name.lower() != mod_only.lower():
                continue
            try:
                cm = comp.CodeModule
                if cm.CountOfLines == 0:
                    continue
                m = pattern.search(cm.Lines(1, cm.CountOfLines))
            except Exception:
                continue
            if m:
                kind = m.group(1).lower()
                break
    except Exception:
        return None
    if kind is None:
        return None
    call_line = f"    r = {expr}" if kind == "function" else f"    {expr}"
    code = "\r\n".join([
        "Function VMR() As String",
        "    Dim r As Variant",
        "    On Error GoTo eh",
        call_line,
        "    On Error Resume Next",
        "    VMR = \"OK|\" & CStr(r)",
        "    If Err.Number <> 0 Then VMR = \"OK|\"",
        "    Exit Function",
        "eh:",
        "    VMR = \"ERR|\" & Err.Number & \"|\" & Err.Description",
        "End Function",
        "",
    ]) + "\r\n"
    comp = None
    try:
        was_saved = bool(owner.Saved)
    except Exception:
        was_saved = False
    try:
        for c in owner.VBProject.VBComponents:
            if c.Name == _RUN_HARNESS:
                owner.VBProject.VBComponents.Remove(c)
                break
        comp = owner.VBProject.VBComponents.Add(1)
        comp.Name = _RUN_HARNESS
        comp.CodeModule.AddFromString(code)
    except Exception as ex:
        print(f"[DEBUG] ハーネス注入失敗: {ex}", file=sys.stderr)
        if comp is not None:
            try:
                owner.VBProject.VBComponents.Remove(comp)
            except Exception:
                pass
        return None
    quoted = book_name.replace("'", "''")
    return {"wb": owner, "comp": comp, "entry": f"'{quoted}'!{_RUN_HARNESS}.VMR",
            "was_saved": was_saved, "xl": xl, "book_name": book_name}


def run_book_macro(xl, owner, name):
    """owner（ブック名）の引数なし Sub name をハーネス経由で撃つ。実行時エラーは _VbaRuntimeError で上げる。

    素の xl.Run だと実行時エラーが「終了／デバッグ」の窓になり、窓の見張りは閉じないので Excel ごと止まる
    （2026-09-23 shelf-run で 671 秒）。棚を撃つ入口（shelf-run・seiri・先撃ち）はここを通す（2026-09-24 総点検）。
    ハーネスを置けない（識別子でない名前・持ち主が見えない）ときは素の Run。
    """
    q = owner.replace("'", "''")
    h = _prepare_run_harness(xl, f"'{q}'!{name}", name)
    try:
        if h is None:
            return xl.Run(f"'{q}'!{name}")
        ret = str(xl.Run(h["entry"]) or "")
        if ret.startswith("ERR|"):
            parts = ret.split("|", 2)
            raise _VbaRuntimeError(parts[1] if len(parts) > 1 else "?", parts[2] if len(parts) > 2 else "")
        return ret[3:] if ret.startswith("OK|") else ret
    finally:
        if h is not None:
            _remove_run_harness(h)


def _remove_carried_harness(h, wait=15.0):
    """入れ替わった同名のブックからハーネスを外して保存する。外せたら True。

    開き直しは外の .vbs が閉じるのを待ってから行うので、数秒遅れて現れる。wait 秒まで待つ。
    保存しないとハーネス入りのファイルが残る（マクロの引っ越しで運ばれた分はもうファイルに書かれている）。
    """
    import time as _time
    xl, name = h.get("xl"), h.get("book_name")
    if xl is None or not name:
        return False
    deadline = _time.monotonic() + wait
    while True:
        try:
            for i in range(1, int(xl.Workbooks.Count) + 1):
                wb = xl.Workbooks.Item(i)
                if str(wb.Name).lower() != str(name).lower():
                    continue
                comps = wb.VBProject.VBComponents
                for c in comps:
                    if c.Name == _RUN_HARNESS:
                        comps.Remove(c)
                        wb.Save()
                        return True
                return True        # 開き直したブックにハーネスは無い＝片付ける物が無い
        except Exception:
            pass
        if _time.monotonic() > deadline:
            return False
        _time.sleep(1)


def _remove_run_harness(h):
    """ハーネスを撤去し、注入前に未変更だったブックは未変更に戻す（PERSONAL 等の保存確認を出さない）"""
    try:
        h["wb"].VBProject.VBComponents.Remove(h["comp"])
    except Exception:
        # 撃ったマクロがブック自身を閉じて入れ替えた（マクロの引っ越し）＝ハーネスも新しいブックへ運ばれ、
        # ファイルに書かれている（2026-09-24）。同じ名前で開き直ったブックから外し、保存して消す
        if _remove_carried_harness(h):
            return
        print(f"警告: ハーネスの撤去に失敗しました（モジュール '{_RUN_HARNESS}' が残っていたら"
              "手で削除してください）", file=sys.stderr)
    if h.get("was_saved"):
        try:
            h["wb"].Saved = True
        except Exception:
            pass


def _com_error_text(e):
    """COM例外からVBAエラーの詳細（エラー番号、説明文、発生元）を分かりやすく整形して返す"""
    hresult = getattr(e, 'hresult', None)
    desc = None
    src = None
    scode = None
    try:
        info = getattr(e, 'excepinfo', None)
        if not info and hasattr(e, 'args') and len(e.args) > 2:
            info = e.args[2]
        if info:
            if len(info) > 1 and info[1]:
                src = str(info[1]).strip()
            if len(info) > 2 and info[2]:
                desc = str(info[2]).strip()
            if len(info) > 5 and info[5]:
                scode = info[5]
    except Exception:
        pass

    # scode / hresult から VBA エラー番号を抽出 (下位16ビット)
    vba_num = None
    code_val = scode or hresult
    if isinstance(code_val, int):
        if (code_val & 0xFFFF0000) in (-2146828288, 0x800A0000):
            vba_num = code_val & 0xFFFF

    parts = []
    if vba_num is not None:
        parts.append(f"実行時エラー {vba_num}")
    if desc:
        parts.append(desc)
    elif str(e):
        parts.append(str(e))
    if src:
        parts.append(f"[{src}]")

    return " - ".join(parts) if parts else str(e)


def cmd_test(args):
    """VBAテストランナー: test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]

    名前が「テスト」または「test」で始まる**引数なしの公開 Sub** を、
    開いたままのブックの中で1本ずつ実行し、成功/失敗を一覧で返す。
    テスト側の作法はただ一つ「失敗は Err.Raise で知らせる」
    （assert は VBA の  If 実際 <> 期待 Then Err.Raise 5, , "説明"  で書く）。
    補助モジュール不要＝テストも単体で他ブックに移植できる自立ユニット。
    実行はエラー捕捉ハーネス（一時モジュールを注入→Run→撤去）経由。
    テスト内の実行時エラーは VBA 側の On Error が受けるので、
    「Microsoft Visual Basic 実行時エラー」ダイアログは出ない。
    xlflow のテスト基盤とエラー割り込みの発想だけ移植し、ビルドせず相乗りのまま回す。
    全部成功なら終了コード0、1本でも失敗なら1（自動化ゲートにそのまま使える）。
    """
    import time as _time
    target_file, rest = parse_target_and_rest(args.posargs)
    keyword = rest[0] if rest else None
    if _reject_extra_args(rest, 1 if keyword else 0,
                          '使い方: test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]'):
        return False

    xl, wb = get_workbook(target_file)

    # 引数なしの公開 Sub だけを対象にする（Private/Friend/Function/引数つきは対象外）
    sub_pattern = re.compile(
        r'^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+([^\s\(\)]+)\s*\(\s*\)',
        re.IGNORECASE | re.MULTILINE
    )

    module_filter = getattr(args, 'module', None)
    tests = []          # (モジュール名, Sub名)
    scanned_modules = 0
    try:
        for comp in wb.VBProject.VBComponents:
            if comp.Type != 1:          # 標準モジュールのみ（Runで呼べる場所）
                continue
            if module_filter and comp.Name.lower() != module_filter.lower():
                continue
            scanned_modules += 1
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            code = cm.Lines(1, cm.CountOfLines)
            for m in sub_pattern.finditer(code):
                name = m.group(1)
                if not (name.startswith('テスト') or name.lower().startswith('test')):
                    continue
                if keyword and keyword.lower() not in name.lower():
                    continue
                tests.append((comp.Name, name))
    except Exception as ex:
        print(f"エラー: VBAプロジェクトの走査に失敗しました: {ex}", file=sys.stderr)
        return False

    if not tests:
        where = f"モジュール '{module_filter}'" if module_filter else f"標準モジュール {scanned_modules} 本"
        print(f"テストが見つかりません（{where} を走査）。")
        print("  名前が「テスト」または「test」で始まる引数なしの Sub がテストとして拾われます。")
        print("  例: Sub テスト加算()  /  失敗は  Err.Raise 5, , \"期待3 実際=\" & 結果  で知らせる。")
        return False

    # 名前は先に取っておく（時間切れで演習用 Excel を畳んだ後は wb に触れない）
    _book_name = wb.Name
    print(f"テスト実行: {_book_name}  （{len(tests)}本）")
    print("-" * 60)

    # ダイアログ対策は既定で常設（run-macro / write-range と同じ）。test は任意の
    # テスト Sub を Application.Run する＝最も VBA を発火させる経路で、テスト内の
    # MsgBox で無言ハングする（--auto-dialog 明示時だけ監視、では守れない）。
    # 明示時はそのボタンで応答、省略時は安全解除（キャンセル優先）。
    # --input-text は InputBox 系に値を入れて OK。
    _auto_dialog = getattr(args, 'auto_dialog', None)
    _dlg_watcher = _start_dialog_watcher(xl, _auto_dialog, getattr(args, 'input_text', None))
    # --timeout: テスト全体の制限時間。gate（演習用 Excel）からは owned_pid 付きで呼ばれ、
    # 時間切れはその Excel を強制終了＝固まった呼び出しが即戻る。生きている Excel には
    # CoCancelCall で待つのをやめるだけ（マクロは止まらない・応答しない VBA には効かない）
    _owned_pid = getattr(args, 'owned_pid', None)
    _watchdog = _start_call_watchdog(getattr(args, 'timeout', None), owned_pid=_owned_pid)

    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    # エラー捕捉ハーネスを一時モジュールとして注入する。
    # 直接 Application.Run すると、テスト内の実行時エラー（Err.Raise 含む）は
    # COM例外にならず「Microsoft Visual Basic 実行時エラー」ダイアログで停止する
    # （実弾で確認済み・618秒ハングの正体）。さらに「VBA側で Application.Run を
    # 経由して呼ぶ」形も、Run の先のエラーが呼び元の On Error に届かず同じ結果に
    # なった（231秒・実測）。だからテストごとに**直接呼び出す**ラッパー関数を
    # 機械生成して注入する＝通常の呼び出しスタックなので On Error が確実に効き、
    # ダイアログは一切出ない。注入→Run→撤去の3ステップ、ブックは保存しない。
    _HARNESS = "VbaManagerTestHarness"
    lines = []
    for i, (mod_name, sub_name) in enumerate(tests, 1):
        lines += [
            f"Function VMT_{i}() As String",
            "    On Error GoTo eh",
            f"    {mod_name}.{sub_name}",
            f"    VMT_{i} = \"OK\"",
            "    Exit Function",
            "eh:",
            f"    VMT_{i} = \"ERR|\" & Err.Number & \"|\" & Err.Description",
            "End Function",
            "",
        ]
    harness_code = "\r\n".join(lines) + "\r\n"
    harness_comp = None
    try:
        # 前回の残骸があれば先に撤去
        for c in wb.VBProject.VBComponents:
            if c.Name == _HARNESS:
                wb.VBProject.VBComponents.Remove(c)
                break
        harness_comp = wb.VBProject.VBComponents.Add(1)
        harness_comp.Name = _HARNESS
        harness_comp.CodeModule.AddFromString(harness_code)
    except Exception as ex:
        print(f"エラー: テストハーネスの注入に失敗しました: {ex}", file=sys.stderr)
        # Add 成功後に Name 代入や AddFromString で失敗すると、既定名（Module1等）の
        # ゴミモジュールが残る。名前が _HARNESS でないと次回の残骸掃除にも拾われないため
        # ここで確実に撤去する
        if harness_comp is not None:
            try:
                wb.VBProject.VBComponents.Remove(harness_comp)
            except Exception:
                print("警告: 注入途中のモジュールを撤去できませんでした"
                      "（既定名のモジュールが残っていたら手で削除してください）", file=sys.stderr)
        _watchdog.stop()
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
            # 失敗して抜ける経路でも、検出したダイアログは必ず報告する
            # （通常の finally は報告するのに、ここだけ黙って捨てていた）
            _note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
            if _note:
                print(_note)
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass
        return False

    results = []
    # ブック名の ' は Excel 規約どおり '' に重ねる（cmd_run_macro と同じ流儀）
    quoted_wb = wb.Name.replace("'", "''")
    try:
        for i, (mod_name, sub_name) in enumerate(tests, 1):
            if _watchdog.fired:
                # 時間切れの後は残りを走らせない（Excel は止めた／止まっていない）。走らせていない事実を残す
                results.append({"module": mod_name, "name": sub_name,
                                "ok": False, "seconds": 0, "error": "未実行（時間切れのため）"})
                print(f"－ {sub_name}  [{mod_name}]  未実行（時間切れ）")
                continue
            t0 = _time.time()
            try:
                ret = xl.Application.Run(f"'{quoted_wb}'!{_HARNESS}.VMT_{i}")
                sec = _time.time() - t0
                ret = str(ret) if ret is not None else ""
                if ret == "OK":
                    results.append({"module": mod_name, "name": sub_name,
                                    "ok": True, "seconds": round(sec, 2), "error": None})
                    print(f"○ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                else:
                    parts = ret.split("|", 2)
                    err = (f"実行時エラー {parts[1]}: {parts[2]}"
                           if len(parts) == 3 else (ret or "不明なエラー"))
                    results.append({"module": mod_name, "name": sub_name,
                                    "ok": False, "seconds": round(sec, 2), "error": err})
                    print(f"✗ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                    print(f"    {err}")
            except Exception as e:
                sec = _time.time() - t0
                err = _com_error_text(e)
                if _watchdog.fired:
                    err = (f"時間切れ（{_watchdog.seconds:g}秒）: "
                           + ("演習用 Excel を強制終了しました" if _owned_pid is not None
                              else "待つのをやめました（マクロは Excel 側で動いている可能性があります）"))
                results.append({"module": mod_name, "name": sub_name,
                                "ok": False, "seconds": round(sec, 2), "error": err})
                print(f"✗ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                print(f"    {err}")
    finally:
        _watchdog.stop()
        if harness_comp is not None:
            try:
                wb.VBProject.VBComponents.Remove(harness_comp)
            except Exception:
                print("警告: テストハーネスの撤去に失敗しました（モジュール "
                      f"'{_HARNESS}' が残っていたら手で削除してください）", file=sys.stderr)
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
            # 解除したダイアログがあれば必ず本文で報告する（無言で握りつぶさない）
            _note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
            if _note:
                print(_note)
        # 戻さないとユーザーの Excel セッションに DisplayAlerts=False が残る
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass

    ok_count = sum(1 for r in results if r["ok"])
    print("-" * 60)
    print(f"結果: {ok_count}/{len(results)} 成功" + ("" if ok_count == len(results) else f"  （失敗 {len(results) - ok_count}）"))
    if _watchdog.fired:
        print(f"⚠ 時間切れ（{_watchdog.seconds:g}秒）で打ち切りました")

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": ok_count == len(results), "book": _book_name,
                          "total": len(results), "passed": ok_count,
                          "tests": results, "timeout": bool(_watchdog.fired)},
                         ensure_ascii=False), file=sys.stdout)

    return ok_count == len(results)




class _JsonLineSieve:
    """sys.stdout を一時的に差し替え、1行 JSON（{…}）だけ抜き取って残りはそのまま流す。

    gate が check / test を内側から呼ぶとき、人向けの行（○✗ の進み具合）は画面へ流したまま、
    機械向けの JSON だけを受け取るため。--json 指定時は setup_encoding が情報行を stderr へ
    退避しているので、ここを通るのは JSON 本体だけになる。
    """
    def __init__(self):
        self.docs = []
        self._buf = ''
        self._real = None

    def __enter__(self):
        self._real = sys.stdout
        sys.stdout = self
        return self

    def __exit__(self, *exc):
        if self._buf:
            self._emit(self._buf)
            self._buf = ''
        sys.stdout = self._real
        return False

    def write(self, s):
        self._buf += s
        while '\n' in self._buf:
            line, self._buf = self._buf.split('\n', 1)
            self._emit(line + '\n')
        return len(s)

    def _emit(self, line):
        st = line.strip()
        if st.startswith('{') and st.endswith('}'):
            try:
                import json
                self.docs.append(json.loads(st))
                return
            except ValueError:
                pass
        self._real.write(line)

    def flush(self):
        try:
            self._real.flush()
        except Exception:
            pass

    @property
    def encoding(self):
        return getattr(self._real, 'encoding', 'utf-8')

    @property
    def last(self):
        return self.docs[-1] if self.docs else None


def _find_vbe_compile_control(vbe):
    """VBE の「VBAProject のコンパイル」(Id 578) を CommandBars から探す。

    「デバッグ」メニューの下にあるので、ポップアップ(Type=10)の中も1段だけ潜る。
    見つからなければ None（環境差で押せないことがある＝そのときは判定しない）。
    """
    for bar in vbe.CommandBars:
        try:
            ctrls = bar.Controls
            n = ctrls.Count
        except Exception:
            continue
        for i in range(1, n + 1):
            try:
                c = ctrls.Item(i)
                if c.Type == 10:                       # msoControlPopup（デバッグ メニュー等）
                    sub = c.Controls
                    for j in range(1, sub.Count + 1):
                        s = sub.Item(j)
                        if s.Id == 578:
                            return s
                elif c.Id == 578:
                    return c
            except Exception:
                pass
    return None


def _vbe_error_location(vbe):
    """コンパイルエラー直後に VBE が指している場所 → {"module","proc","line","text"}（取れなければ None）。

    2026-09-16 の下見: エラー直後は xl.VBE.ActiveCodePane が落ちたモジュールを指し、cp.GetSelection() が
    win32com の動的ディスパッチで (82, 5, 82, 26) のタプル（out 引数）で返る。cm.Parent.Name＝モジュール名、
    cm.ProcOfLine(82, 0)＝プロシージャ、cm.Lines(82, 1)＝本文。
    """
    try:
        cp = vbe.ActiveCodePane
        if cp is None:
            return None
        sel = cp.GetSelection()
        if not isinstance(sel, (tuple, list)) or not sel:
            return None
        line = int(sel[0])
        if line <= 0:
            return None
        cm = cp.CodeModule
        module = str(cm.Parent.Name)
        try:
            proc = cm.ProcOfLine(line, 0)
            if isinstance(proc, (tuple, list)):        # out 引数つき＝(名前, 種別) のタプルで返ることがある
                proc = proc[0]
            proc = str(proc or '')
        except Exception:
            proc = ''
        try:
            text = str(cm.Lines(line, 1) or '')
        except Exception:
            text = ''
        return {"module": module, "proc": proc, "line": line, "text": text.strip()}
    except Exception:
        return None


def _compile_error_detail(loc, body):
    """コンパイルエラーの本文と場所を 1 行に（純 Python）。場所が無ければ本文だけ（今までの文）。

    形: `[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― 行ラベルが定義されていません。`
    """
    msg = " ".join((body or "").split())
    # ダイアログ監視の本文は「タイトル「Microsoft Visual Basic for Applications」 本文: コンパイル エラー: …」の形。
    # 場所を名指しする行では本文だけにする（タイトルは毎回同じ）
    msg = re.sub(r'^タイトル「[^」]*」\s*本文:\s*', '', msg)
    msg = msg or "コンパイル エラー（本文を取得できませんでした）"
    if not loc:
        return msg
    return f"[{loc['module']}] {loc.get('proc') or '(宣言部)'}:{loc['line']}: {loc.get('text', '')} ― {msg}"


def _compile_fix_hint(loc, body=""):
    """直し方の案内 1 行（純 Python）。ラベル無しならラベルを足す（replace-procedure）、それ以外は code-replace。"""
    if not loc:
        return "※ どの行かは VBE を開くと選択された状態で止まっています"
    mod = loc['module']
    text = loc.get('text', '')
    if "ラベル" in (body or ""):
        return (f"直す: ラベルを足す＝get {mod} {loc.get('proc') or ''} → replace-procedure -y --module {mod}"
                f"（飛び先の名前の打ち間違いなら code-replace \"{text}\" \"新しい行\" --module {mod} -y）")
    return f"直す: code-replace \"{text}\" \"新しい行\" --module {mod} -y"


def _compile_vbproject(xl, wb):
    """VBE の「デバッグ > VBAProject のコンパイル」を外から押して合否を返す。

    **check（静的検査）との違いは「呼ばれないマクロの中まで見る」こと。**
    VBA はプロシージャ単位のオンデマンド コンパイルなので、存在しない処理を呼ぶ Sub は
    実行するまで誰も気づかない。2026-08-28 実測: check は素通り、全体コンパイルは
    「コンパイル エラー: Sub または Function が定義されていません。」で捕まえた。

    合否は Id 578 の Enabled で見る（通れば「もう押す必要がない」＝False に落ちる。
    エラーなら True のまま）。エラー時は VBE がモーダルの MsgBox を出すので、
    既存のダイアログ監視で閉じて本文を持ち帰る。**VBE の窓は出さなくてよい**（実測）。

    返り値: {"ok": bool, "state": str, "detail": str}
      state = compiled（今コンパイルして通った）/ already（すでにコンパイル済み）/
              error（コンパイルエラー）/ unavailable（押せない＝合否を言わない）
    ※ unavailable は ok=True を返す。「見られなかった」ことを「落ちた」にはしない
      （VBOM 未信頼のブックで全部 FAIL になるのを避ける。理由は detail に残る）。
    """
    try:
        vbe = xl.VBE
    except Exception as ex:
        return {"ok": True, "state": "unavailable",
                "detail": "VBE に触れません（VBOM の信頼が要ります）: %s" % _com_error_text(ex)}

    proj = None
    try:
        want = os.path.normcase(str(wb.FullName))
        for p in vbe.VBProjects:
            try:
                if os.path.normcase(str(p.FileName)) == want:
                    proj = p
                    break
            except Exception:
                pass
    except Exception:
        pass
    if proj is None:
        return {"ok": True, "state": "unavailable",
                "detail": "対象のVBAプロジェクトが見つかりません"}

    try:
        vbe.ActiveVBProject = proj
    except Exception:
        pass

    ctl = _find_vbe_compile_control(vbe)
    if ctl is None:
        return {"ok": True, "state": "unavailable",
                "detail": "VBE のコンパイル項目（Id 578）が見つかりません"}
    try:
        if not ctl.Enabled:
            return {"ok": True, "state": "already", "detail": "すでにコンパイル済み"}
    except Exception:
        pass

    vis0 = None                                # 押す前の VBE の窓（エラーが出ると True になる＝下見で判明）
    try:
        vis0 = bool(vbe.MainWindow.Visible)
    except Exception:
        pass
    watcher = _start_dialog_watcher(xl)        # 既定＝安全解除。本文は .last_text に残る
    try:
        ctl.Execute()
    except Exception as ex:
        watcher.stop()
        return {"ok": True, "state": "unavailable",
                "detail": "コンパイルを押せませんでした: %s" % _com_error_text(ex)}
    time.sleep(0.4)                            # 出るなら出きるまで待つ
    watcher.stop()

    body = (getattr(watcher, 'last_text', '') or '').strip()
    try:
        still_enabled = bool(ctl.Enabled)
    except Exception:
        still_enabled = False
    res = None
    if still_enabled or getattr(watcher, 'count', 0):
        # 落ちた行を名指しする（2026-09-17）。読めなければ今までどおり本文だけ
        loc = _vbe_error_location(vbe)
        res = {"ok": False, "state": "error", "detail": _compile_error_detail(loc, body)}
        if loc:
            res.update(loc)
    if vis0 is False:
        try:
            if bool(vbe.MainWindow.Visible):
                vbe.MainWindow.Visible = False  # 押す前は隠れていた窓を戻す（エラーで前に出る副作用）
        except Exception:
            pass
    return res or {"ok": True, "state": "compiled", "detail": "全体コンパイル OK"}


def cmd_compile(args):
    """全体コンパイル: compile [excel_file] [--json]   （別名: 全体コンパイル）

    VBE の「デバッグ > VBAProject のコンパイル」を外から押す。
    **check が見ない「呼ばれないマクロの中」まで見る唯一の手。**
    通れば終了コード 0、コンパイルエラーなら理由を出して 1。
    ※ 走っている VBA は止めない（押すだけ）。VBE の窓は出さない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if _reject_extra_args(rest, 0, '使い方: compile [excel_file] [--json]'):
        return False
    is_json = getattr(args, 'json', False)
    xl, wb = get_workbook(target_file, readonly=True)
    try:
        was_saved = bool(wb.Saved) and bool(str(wb.Path or '')) and not bool(wb.ReadOnly)
    except Exception:
        was_saved = False
    res = _compile_vbproject(xl, wb)
    # コンパイルで作られる中間コードの分だけ「未保存」になり、閉じるときに保存を聞かれていた（中身は変わっていない・
    # 2026-09-24）。撃つ前に保存済みだったブックだけ、今コンパイルして通ったときに保存済みの印を戻す。
    # 保存し直してはいけない＝中間コードがファイルに入って 1.2MB → 1.9MB に膨らむ（マクロの引っ越しで軽くした分が戻る・
    # 2026-09-24 shu）。印だけ戻せば、ファイルは軽いまま・閉じるときも聞かれない
    resaved = False
    if was_saved and res.get("state") == "compiled":
        try:
            if not bool(wb.Saved):
                wb.Saved = True
                resaved = True
        except Exception:
            pass
    if is_json:
        import json as _json
        doc = {"pass": res["ok"], "book": wb.Name, "state": res["state"], "detail": res["detail"]}
        for k in ("module", "proc", "line", "text"):          # 落ちた行の名指し（error のときだけ入る・2026-09-17）
            if k in res:
                doc[k] = res[k]
        print(_json.dumps(doc, ensure_ascii=False))
        return res["ok"]
    label = {"compiled": "通過（コンパイルしました）", "already": "通過（すでにコンパイル済み）",
             "error": "不通過（コンパイル エラー）", "unavailable": "判定なし"}
    print("全体コンパイル: %s" % wb.Name)
    print("  %s" % label.get(res["state"], res["state"]))
    print("  %s" % res["detail"])
    if resaved:
        print("  （コンパイルの前は保存済みだったので保存済みの印を戻しました＝ファイルは書かない・閉じるときに保存を聞かれません）")
    if res["state"] == "error":
        print("  " + _compile_fix_hint(res if res.get("module") else None, res["detail"]))
    return res["ok"]


def cmd_gate(args):
    """関所（停止条件）: gate [excel_file] [絞り込み] [--module 名] [--addins] [--timeout 秒]
                           [--auto-dialog ok] [--input-text 値] [--keep] [--json]

    「直した。で、通ったのか？」を機械が一言で答える＝自動で回すときの停止条件（2026-08-23）。
    流れ（本体には指一本触れない・rehearse と同じ作法）:
      1. 対象ブックの「今この瞬間」のコピーを作る（SaveCopyAs＝未保存の変更込み）
      2. 別インスタンス（DispatchEx・非表示）でコピーを開く（開く瞬間の Workbook_Open は起こさない）
      3. コピーの中で check（構文検査）→ 全体コンパイル（VBE の「VBAProject のコンパイル」を押す）
         → test（「テスト」/test 始まりの引数なし Sub を一括実行）
      4. 判定:  構文エラー 0 かつ 全体コンパイル OK かつ テスト 1 本以上 かつ 全部成功 → 通過（PASS・終了コード 0）
               それ以外は不通過（FAIL・終了コード 1）。テスト 0 本も FAIL＝合否を言えないものは通さない
      5. 演習用 Excel を畳み、コピーを消す（--keep で残す）
    時間切れ（--timeout・既定 120 秒）は演習用 Excel を強制終了して FAIL（ユーザーの Excel には触れない）。
    --json は {"pass","book","check","compile","tests","timeout","seconds","reasons"} を1行で返す。
    ※ 全体コンパイルが入ったので「呼ばれないマクロの中のコンパイルエラー」も見る（2026-08-28）。
      ファイル出力・メール送信など外への副作用はテスト側が触れた分だけ起きる。
    """
    import gc
    import json
    import signal
    import tempfile
    import vbam_core

    target_file, rest = parse_target_and_rest(args.posargs)
    keyword = rest[0] if rest else None
    usage = ('使い方: gate [excel_file] [絞り込み] [--module 名] [--addins] [--timeout 秒] '
             '[--auto-dialog ok] [--input-text 値] [--keep] [--json]')
    if _reject_extra_args(rest, 1 if keyword else 0, usage):
        return False
    is_json = getattr(args, 'json', False)
    timeout = getattr(args, 'timeout', 120)
    keep = getattr(args, 'keep', False)

    # 本体は健診モードで掴む（閉じているブックを自動で開く場合に Workbook_Open を起こさない）
    xl_src, wb_src = get_workbook(target_file, readonly=True)
    src_name = wb_src.Name
    stem, ext = os.path.splitext(src_name)
    if not ext:
        ext = '.xlsx'
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    copy_path = os.path.join(tempfile.gettempdir(), f"{stem}_関所_{stamp}{ext}")
    seq = 1
    while os.path.exists(copy_path):
        seq += 1
        copy_path = os.path.join(tempfile.gettempdir(), f"{stem}_関所_{stamp}_{seq}{ext}")
    try:
        wb_src.SaveCopyAs(copy_path)
    except Exception as e:
        print(f"エラー: コピーを作れませんでした: {e}")
        print("  （一度も保存していない新規ブックは、先に save-as で実体を作ってください）")
        return False
    print(f"関所: {src_name} のコピーで検分します（本体は無傷）")
    print(f"  コピー: {copy_path}")
    cmd_progress_note("コピーを別の Excel で開いています")

    t_start = time.time()
    # 演習用の Excel は必ず別インスタンス（DispatchEx）。ユーザーの Excel には触れない
    xl2 = win32com.client.DispatchEx("Excel.Application")
    pid = None
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(xl2.Hwnd)
    except Exception:
        pass
    if not vbam_core.is_fresh_instance(xl2):
        # 使う人の Excel に合流した（2026-09-23）。ここで撃つと人のブックを触り、後始末で人の Excel を落とす
        print(f"エラー: 演習用の Excel を起こせませんでした（使う人の Excel・PID {pid} に合流）。"
              "Excel を開き直してからやり直してください")
        return False
    inst = {"xl": xl2, "pid": pid}
    vbam_core._created_instances.append(inst)      # 安全網（本線は下の finally で畳む）
    cache_key = os.path.abspath(copy_path).lower()

    check_doc = None
    test_doc = None
    compile_doc = None
    n_err = n_warn = 0
    fatal = None
    timed_out = False
    wb2 = None
    # 開く〜構文検査の間も時間切れを見張る（テスト段階は cmd_test の見張りに残り時間ごと引き継ぐ）
    wd = _start_call_watchdog(timeout, owned_pid=pid)
    try:
        xl2.Visible = "--visible" in sys.argv or "-v" in sys.argv
        xl2.DisplayAlerts = False
        try:
            xl2.EnableEvents = False       # 開く瞬間の Workbook_Open は起こさない
        except Exception:
            pass
        if getattr(args, 'addins', False):
            load_excel_addins_and_personal(xl2)
        wb2 = xl2.Workbooks.Open(copy_path, 0)   # UpdateLinks=0
        try:
            xl2.EnableEvents = True        # テスト中のイベントは本番同様に生かす
        except Exception:
            pass
        # 以後の check / test はこの接続を使う（get_workbook の接続キャッシュに載せる＝
        # 開き直さない・本体を掴まない）
        vbam_core._wb_cache[cache_key] = (xl2, wb2, False)

        # 1) 構文検査（JSON で受けて要点だけ出す）
        cmd_progress_note("構文検査（check）")
        with _JsonLineSieve() as sieve:
            cmd_check(argparse.Namespace(posargs=[copy_path], json=True))
        check_doc = sieve.last
        if check_doc is not None:
            summ = check_doc.get('summary') or {}
            n_err = int(summ.get('errors', 0) or 0)
            n_warn = int(summ.get('warnings', 0) or 0)
            print(f"  構文: {'OK' if n_err == 0 else 'NG'}（エラー {n_err}・警告 {n_warn}）")
            for m in check_doc.get('modules') or []:
                for e in m.get('errors') or []:
                    print(f"    [ERROR] {m.get('name')}: {e}")
        wd.stop()
        if wd.fired:
            timed_out = True
        else:
            # 2) 全体コンパイル（check が見ない「呼ばれないマクロの中」まで見る）
            cmd_progress_note("全体コンパイル")
            compile_doc = _compile_vbproject(xl2, wb2)
            print("  全体コンパイル: %s（%s）" % ("OK" if compile_doc["ok"] else "NG",
                                            compile_doc["detail"]))
            # 3) テスト（残り時間を引き継ぐ。時間切れは演習用 Excel を強制終了）
            remaining = None
            if wd.seconds:
                remaining = max(1.0, wd.seconds - (time.time() - t_start))
            ns = argparse.Namespace(
                posargs=[copy_path] + ([keyword] if keyword else []),
                module=getattr(args, 'module', None), json=True,
                auto_dialog=getattr(args, 'auto_dialog', None),
                input_text=getattr(args, 'input_text', None),
                timeout=remaining, owned_pid=pid)
            cmd_progress_note("テスト Sub を一括実行（test）")
            with _JsonLineSieve() as sieve:
                cmd_test(ns)
            test_doc = sieve.last            # None＝テストが見つからない（cmd_test が案内を出す）
            if test_doc is not None and test_doc.get('timeout'):
                timed_out = True
    except Exception as e:
        fatal = _com_error_text(e)
    finally:
        wd.stop()
        if wd.fired:
            timed_out = True
        vbam_core._wb_cache.pop(cache_key, None)
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
        try:
            vbam_core._created_instances.remove(inst)
            inst["xl"] = None
        except ValueError:
            pass
        gc.collect()
        if timed_out and pid is not None and vbam_core._pid_is_excel(pid):
            # 見張りが撃ち損ねていたら、ここで確実に畳む（演習用だけ。PID の身元は確認済み）
            try:
                os.kill(pid, signal.SIGTERM)
            except Exception:
                pass

    elapsed = time.time() - t_start
    total = int(test_doc.get('total', 0) or 0) if test_doc else 0
    passed = int(test_doc.get('passed', 0) or 0) if test_doc else 0
    failed = [t for t in ((test_doc or {}).get('tests') or []) if not t.get('ok')]

    reasons = []
    if fatal:
        reasons.append(f"最後まで通せませんでした: {fatal}")
    if check_doc is None:
        if not fatal:
            reasons.append("構文検査ができませんでした")
    elif n_err:
        reasons.append(f"構文エラー {n_err} 件")
    if compile_doc is not None and not compile_doc["ok"]:
        reasons.append(f"全体コンパイル NG: {compile_doc['detail']}")
    if timed_out:
        reasons.append(f"時間切れ（{float(timeout):g}秒）＝演習用 Excel を強制終了")
        if failed:
            reasons.append(f"テスト {len(failed)}/{total} 本 NG（未実行を含む）")
    elif not fatal:
        if total == 0:
            reasons.append("テストが 0 本＝合否を言えない（「テスト」/test 始まりの引数なし Sub を書く）")
        elif failed:
            reasons.append(f"テスト {len(failed)}/{total} 本 NG")
    ok = not reasons

    print("-" * 60)
    if ok:
        _cmp = (compile_doc or {}).get("state")
        _cmp_word = "・全体コンパイル OK" if _cmp in ("compiled", "already") else (
            "・全体コンパイルは判定なし" if _cmp else "")
        print(f"判定: 通過（PASS）  構文エラー 0{_cmp_word}・テスト {passed}/{total} 成功  （{elapsed:.1f}秒）")
    else:
        print(f"判定: 不通過（FAIL）  理由: " + " / ".join(reasons) + f"  （{elapsed:.1f}秒）")
        for t in failed[:20]:
            print(f"    ✗ {t.get('name')}  [{t.get('module')}]  {t.get('error')}")

    if keep:
        print(f"コピーを残しました: {copy_path}")
    else:
        # 強制終了直後はロックが残ることがあるので少し待って消す
        removed = False
        for _ in range(6):
            try:
                os.remove(copy_path)
                removed = True
                break
            except FileNotFoundError:
                removed = True
                break
            except OSError:
                time.sleep(0.5)
        if not removed:
            print(f"⚠ コピーを削除できませんでした（手で削除してください）: {copy_path}")

    if is_json:
        out = {"pass": ok, "book": src_name, "copy": copy_path if keep else None,
               "check": ({"errors": n_err, "warnings": n_warn} if check_doc is not None else None),
               "compile": compile_doc,
               "tests": {"total": total, "passed": passed,
                         "failed": [{"module": t.get('module'), "name": t.get('name'),
                                     "error": t.get('error')} for t in failed]},
               "timeout": timed_out, "seconds": round(elapsed, 1), "reasons": reasons}
        print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
    return ok


__all__ = [
    'CHECKUP_ACK_DIR',
    'CHECKUP_HISTORY_DIR',
    '_AUTO_EXEC_PREFIXES',
    '_AUTO_EXEC_STD',
    '_VbaRuntimeError',
    '_all_procedure_names',
    '_analyze_calls',
    '_broken_links',
    '_check_bas_one',
    '_checkup_ack_path',
    '_checkup_diff',
    '_checkup_history_path',
    '_checkup_one',
    '_checkup_rating',
    '_collect_book_inventory',
    '_com_error_text',
    '_compile_error_detail',
    '_compile_fix_hint',
    '_compile_vbproject',
    '_diag_missing_labels',
    '_diag_on_err_typo',
    '_dialog_watcher_note',
    'dialog_safe',
    '_extra_code_scans',
    '_extract_proc',
    '_find_consecutive_dup_lines',
    '_find_duplicate_procedures',
    '_find_macro_owner_books',
    '_find_vbe_compile_control',
    '_inline_body_after_decl',
    '_inventory_or_explain',
    '_load_checkup_ack',
    '_load_checkup_history',
    '_narrow_proc_range',
    '_parse_module_blocks',
    '_prepare_run_harness',
    '_project_book_name',
    '_remove_run_harness',
    'run_book_macro',
    '_save_checkup_ack',
    '_save_checkup_history',
    '_select_addin_project',
    '_sheet_health',
    '_start_call_watchdog',
    '_start_dialog_watcher',
    '_strip_vba_comment',
    '_suggest_similar',
    '_vba_references',
    '_vbe_error_location',
    '_write_module',
    'cmd_add_module',
    'cmd_add_procedure',
    'cmd_call_graph',
    'cmd_capabilities',
    'cmd_check',
    'cmd_inspect_gui',
    'cmd_metrics',
    'cmd_process',
    'cmd_rules',
    'cmd_check_bas',
    'cmd_checkup',
    'cmd_code_replace',
    'cmd_compile',
    'cmd_delete_module',
    'cmd_delete_procedure',
    'cmd_diag',
    'cmd_docs',
    'cmd_export_all',
    'cmd_export_module',
    'cmd_gate',
    'cmd_get',
    'cmd_grep',
    'cmd_impact',
    'cmd_list',
    'cmd_list_backups',
    'cmd_list_forms',
    'cmd_list_modules',
    'cmd_list_open',
    'cmd_list_shortcuts',
    'cmd_reorder_macro',
    'cmd_replace_module',
    'cmd_replace_procedure',
    'cmd_patch_procedure',
    'cmd_restore',
    'cmd_register_addin',
    'cmd_run_macro',
    'cmd_setup_check',
    'cmd_test',
]
