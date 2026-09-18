# -*- coding: utf-8 -*-
"""vbam_devtools（2026-09-16 職場向け VBA ツール作りの手）と、同日の直し（軽い控え・進み具合・フック）の純ロジックのテスト。

実行: このフォルダで `py -m pytest test_devtools.py -q`。Excel には接続しない。
"""
import os
import sys
import json
import time
import argparse
import importlib.util

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vba_manager as vm
import vbam_devtools as dt
import vbam_core as vc
import vbam_inv as vi


# ================================================================
# rename-procedure（純 Python の計画）
# ================================================================

def test_rename_plan_word_boundary_and_case():
    code = ("Sub 集計()\n"
            "    Call 集計A\n"
            "    集計 1\n"
            "    Application.Run \"集計\"\n"
            "    x = 集計表\n"
            "    Debug.Print SHUUKEI\n"
            "End Sub\n")
    plan = dt._rename_plan_lines(code, "集計", "月次集計")
    lines = {i: n for i, _, n in plan}
    assert lines[1] == "Sub 月次集計()"
    assert 2 not in lines, "集計A は別の識別子＝触らない"
    assert lines[3] == "    月次集計 1"
    assert lines[4] == '    Application.Run "月次集計"'
    assert 5 not in lines, "集計表 は別の識別子"


def test_rename_plan_ascii_case_insensitive():
    code = "Sub Main()\n    Call helper\n    HELPER\nEnd Sub\nSub Helper()\nEnd Sub\n"
    plan = dt._rename_plan_lines(code, "helper", "Helper2")
    assert [n.strip() for _, _, n in plan] == ["Call Helper2", "Helper2", "Sub Helper2()"]


def test_rename_procedure_wired_and_refuses_bad_name(capsys):
    assert vm.raw_command(vm._command_table()["rename-procedure"]) is dt.cmd_rename_procedure
    ns = argparse.Namespace(posargs=["A", "_b"], module_opt=None, yes=True, dry_run=True, force=False)
    assert dt.cmd_rename_procedure(ns) is False
    assert "識別子" in capsys.readouterr().out


# ================================================================
# diff-module（Attribute を外して比べる）
# ================================================================

def test_module_diff_ignores_attribute_lines():
    a = 'Attribute VB_Name = "M"\r\nSub A()\r\nAttribute A.VB_ProcData.VB_Invoke_Func = "k\\n14"\r\n    x = 1\r\nEnd Sub\r\n'
    b = 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n'
    lines, minus, plus = dt._module_diff(a, b, "a", "b")
    assert (minus, plus) == (0, 0) and lines == []


def test_module_diff_counts_changes_and_caps():
    a = "\n".join(f"line {i}" for i in range(30))
    b = "\n".join(f"line {i}" if i % 3 else f"LINE {i}" for i in range(30))
    lines, minus, plus = dt._module_diff(a, b, "a", "b", max_lines=5)
    assert minus == 10 and plus == 10
    assert len(lines) == 6 and "他" in lines[-1]


# ================================================================
# set-shortcut（キーの読み方）
# ================================================================

@pytest.mark.parametrize("text,key", [
    ("k", "k"), ("K", "K"), ("ctrl+k", "k"), ("Ctrl+Shift+K", "K"), ("ctrl+shift+k", "K"),
    ("Ctrl + K", "K"), ("shift k", "K"), ("5", "5"), ("", None), ("kk", None), ("ctrl+", None), ("あ", None),
])
def test_parse_shortcut_key(text, key):
    assert dt._parse_shortcut_key(text) == key


def test_set_invoke_attribute_replaces_inserts_and_clears():
    bas = ('Attribute VB_Name = "M"\r\n'
           'Sub 図形一覧を表示()\r\n'
           'Attribute 図形一覧を表示.VB_ProcData.VB_Invoke_Func = " \\n14"\r\n'
           '    x = 1\r\n'
           'End Sub\r\n'
           'Public Sub 他()\r\n'
           'Attribute 他.VB_ProcData.VB_Invoke_Func = "B\\n14"\r\n'
           'End Sub\r\n')
    out, found = dt._set_invoke_attribute(bas, "図形一覧を表示", "K")
    assert found
    lines = out.split("\r\n")
    i = lines.index("Sub 図形一覧を表示()")
    assert lines[i + 1] == 'Attribute 図形一覧を表示.VB_ProcData.VB_Invoke_Func = "K\\n14"'
    assert out.count("図形一覧を表示.VB_ProcData") == 1, "空のキーの古い行は消える"
    assert 'Attribute 他.VB_ProcData.VB_Invoke_Func = "B\\n14"' in out, "他のマクロの鍵は触らない"
    cleared, found = dt._set_invoke_attribute(out, "図形一覧を表示", None)
    assert found and "図形一覧を表示.VB_ProcData" not in cleared
    _, found = dt._set_invoke_attribute(bas, "無い", "K")
    assert not found


def test_shortcut_label():
    assert dt._shortcut_label("K") == "Ctrl+Shift+K"
    assert dt._shortcut_label("k") == "Ctrl+k"


# ================================================================
# backup-prune（系列ごとに新しい K 個は残す）
# ================================================================

def test_prune_plan_keeps_newest_per_series_and_recent():
    now = 1_800_000_000
    day = 86400
    entries = [
        ("book.xlsm.backup_before_x_ab12_20260101_000000.xlsm", now - 100 * day, 10),
        ("book.xlsm.backup_before_x_ab12_20260201_000000.xlsm", now - 60 * day, 10),
        ("book.xlsm.backup_before_x_ab12_20260301_000000.xlsm", now - 40 * day, 10),   # 系列の最新＝残す
        ("book_Mod_20260101_000000.bas", now - 90 * day, 5),                            # 系列の唯一＝残す
        ("other_Mod_20260101_000000.bas", now - 90 * day, 5),
        ("other_Mod_20260901_000000.bas", now - 1 * day, 5),
    ]
    doomed, kept = dt._prune_plan(entries, days=30, keep=1, now=now)
    names = [e[0] for e in doomed]
    assert names == ["book.xlsm.backup_before_x_ab12_20260101_000000.xlsm",
                     "other_Mod_20260101_000000.bas",
                     "book.xlsm.backup_before_x_ab12_20260201_000000.xlsm"]
    assert kept == 3
    doomed2, _ = dt._prune_plan(entries, days=30, keep=3, now=now)
    assert [e[0] for e in doomed2] == []


def test_backup_series_strips_stamp_and_seq():
    assert dt._backup_series("a.xlsm.backup_before_m_0f1e_20260916_101010_2.xlsm") == "a.xlsm.backup_before_m_0f1e_#.xlsm"
    assert dt._backup_series("book_Mod_20260916_101010.bas") == "book_Mod_#.bas"


def test_backup_prune_dry_run_does_not_delete(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(dt, "BACKUP_DIR", str(tmp_path))
    old = tmp_path / "b_M_20260101_000000.bas"
    old.write_bytes(b"x")
    os.utime(old, (time.time() - 100 * 86400,) * 2)
    new = tmp_path / "b_M_20260901_000000.bas"
    new.write_bytes(b"y")
    ns = argparse.Namespace(posargs=[], days=30, keep=1, force=False, json=True)
    assert dt.cmd_backup_prune(ns) is True
    out = json.loads(capsys.readouterr().out)
    assert out["prune"] == 1 and out["deleted"] is False and old.exists()
    ns.force = True
    ns.json = False
    assert dt.cmd_backup_prune(ns) is True
    assert not old.exists() and new.exists()


# ================================================================
# export-all --history（フォルダ同士の差）
# ================================================================

def test_export_dirs_diff(tmp_path):
    prev, cur = tmp_path / "20260101_000000", tmp_path / "20260102_000000"
    prev.mkdir()
    cur.mkdir()
    (prev / "A.bas").write_bytes("Attribute VB_Name = \"A\"\r\nSub A()\r\nEnd Sub\r\n".encode("cp932"))
    (cur / "A.bas").write_bytes("Attribute VB_Name = \"A\"\r\nSub A()\r\n    x = 1\r\nEnd Sub\r\n".encode("cp932"))
    (prev / "B.bas").write_bytes(b"Sub B()\r\nEnd Sub\r\n")
    (cur / "B.bas").write_bytes(b"Sub B()\r\nEnd Sub\r\n")
    (prev / "Gone.bas").write_bytes(b"Sub G()\r\nEnd Sub\r\n")
    (cur / "New.cls").write_bytes(b"Sub N()\r\nEnd Sub\r\n")
    (cur / "notes.txt").write_bytes(b"ignored")
    d = dt._export_dirs_diff(str(prev), str(cur))
    assert d["added"] == ["New.cls"] and d["removed"] == ["Gone.bas"]
    assert d["changed"] == [("A.bas", 0, 1)] and d["same"] == 1


def test_export_history_note_removes_identical_folder(tmp_path, capsys):
    parent = tmp_path / "Book"
    prev, cur = parent / "20260101_000000", parent / "20260102_000000"
    prev.mkdir(parents=True)
    cur.mkdir()
    (prev / "A.bas").write_bytes(b"Sub A()\r\nEnd Sub\r\n")
    (cur / "A.bas").write_bytes(b"Sub A()\r\nEnd Sub\r\n")
    stamp, same = dt._export_history_note("Book.xlsm", str(cur))
    assert stamp == "20260101_000000" and same is True
    assert not cur.exists() and prev.exists()
    assert "同じ" in capsys.readouterr().out


def test_export_all_parser_has_history():
    p = vm.build_parser()
    ns = p.parse_args(["export-all", "--history"])
    assert ns.history is True


# ================================================================
# format-module（Attribute 行を壊さない）
# ================================================================

def test_format_bas_text_keeps_procdata_attribute_under_declaration():
    src = ('Attribute VB_Name = "M"\n'
           'Sub 集計()\n'
           'Attribute 集計.VB_ProcData.VB_Invoke_Func = "k\\n14"\n'
           '    Dim y As Long\n'
           '    x = 1\n'
           '    y = x\n'
           'End Sub\n')
    out, added = dt._format_bas_text(src)
    lines = out.split("\n")
    i = lines.index("Sub 集計()")
    assert lines[i + 1].startswith("Attribute 集計.VB_ProcData"), "宣言の直下に戻る"
    assert out.count("Attribute 集計.VB_ProcData") == 1
    assert "' --- 変数宣言 ---" in out and "Dim y As Long" in out


def test_format_bas_text_keeps_comments_between_subs_and_trailing():
    # 2026-09-16 実射で発覚: Sub と Sub の間の説明コメント（10 行）が整形で消えていた
    src = ("Sub A()\n    x = 1\nEnd Sub\n\n"
           "'----\n' テスト B の説明\n'----\n"
           "Public Sub B()\n    y = 2\nEnd Sub\n\n' 末尾のメモ\n")
    out, _ = dt._format_bas_text(src)
    assert "' テスト B の説明" in out and "' 末尾のメモ" in out
    lines = out.split("\n")
    assert lines[lines.index("Public Sub B()") - 1] == "'----", "説明コメントは Sub に密着"
    assert out.count("Sub A()") == 1 and out.count("Public Sub B()") == 1
    again, _ = dt._format_bas_text(out)
    assert again == out, "2 回当てても変わらない"


# ================================================================
# 進み具合（status）と台帳の包み
# ================================================================

def test_cmd_progress_write_note_read(monkeypatch):
    rec = vc.cmd_progress_write('run', cmd='export-all', args=argparse.Namespace(posargs=['a', 'b', 'c', 'd']))
    assert rec['args'] == 'a b c' and rec['pid'] == os.getpid()
    vc.cmd_progress_note("書き出し 3/10")
    got = vc.cmd_progress_read()
    assert got['state'] == 'run' and got['note'] == "書き出し 3/10"
    vc.cmd_progress_write('done', cmd='export-all', ok=True, sec=1.234)
    got = vc.cmd_progress_read()
    assert got['state'] == 'done' and got['ok'] is True and got['sec'] == 1.234
    vc.cmd_progress_note("届かない")           # 走っていないときは書かない
    assert 'note' not in vc.cmd_progress_read()


def test_logged_wrapper_writes_progress_start_and_done(monkeypatch):
    seen = []

    def fake(args):
        seen.append(vc.cmd_progress_read())
        return True
    fake.__name__ = 'cmd_fake_thing'
    run = vm._logged(fake)
    run(argparse.Namespace(posargs=['x'], command='fake-thing'))
    assert seen[0]['state'] == 'run' and seen[0]['cmd'] == 'fake-thing'
    after = vc.cmd_progress_read()
    assert after['state'] == 'done' and after['ok'] is True and after['cmd'] == 'fake-thing'


def test_status_reports_running_and_agent(monkeypatch, capsys):
    assert vm.raw_command(vm._command_table()["status"]) is dt.cmd_status
    vc.cmd_progress_write('run', cmd='gate', args=argparse.Namespace(posargs=[]), note="全体コンパイル")
    with open(dt._AGENT_PROGRESS_PATH, 'w', encoding='utf-8') as f:
        json.dump({'time': '2026-09-16 10:00:00', 'state': 'ask', 'line': 'AI に聞いています / 往復 1/4'}, f)
    assert dt.cmd_status(argparse.Namespace(posargs=[], json=False)) is True
    out = capsys.readouterr().out
    assert "走っています: gate" in out and "全体コンパイル" in out and "AI に聞いています" in out
    # status 自身は進み具合を上書きしない（_logged の包みが 'status' を外す）
    run = vm._command_table()["status"]
    run(argparse.Namespace(posargs=[], json=True, command='status'))
    assert vc.cmd_progress_read()['cmd'] == 'gate'


def test_status_detects_dead_process(monkeypatch, capsys):
    vc.cmd_progress_write('run', cmd='checkup', args=argparse.Namespace(posargs=[]))
    monkeypatch.setattr(dt, "_pid_alive", lambda pid: False)
    dt.cmd_status(argparse.Namespace(posargs=[], json=False))
    assert "途中で止まった" in capsys.readouterr().out


# ================================================================
# 不変条件の軽い控え（触らないシートは Formula の一括 1 回）
# ================================================================

class _R:
    def __init__(self, addr, formula):
        self.Address = addr
        self.Formula = formula
        self.Value = formula
        self.Rows = type("Rows", (), {"Count": len(formula)})()
        self.Columns = type("Cols", (), {"Count": len(formula[0])})()


class _Sheet:
    def __init__(self, name, formula, addr="$A$1:$B$2"):
        self.Name = name
        self.UsedRange = _R(addr, formula)
        self.AutoFilterMode = False
        self.ListObjects = []

    def ChartObjects(self):
        return type("CO", (), {"Count": 0})()


class _Book:
    def __init__(self, sheets):
        self.Worksheets = sheets
        self.Sheets = sheets
        self.Names = []
        self.Name = "Book1"
        self.Application = type("App", (), {"Calculation": -4105,
                                            "ActiveWorkbook": type("W", (), {"Name": "Book1"})()})()


def test_snap_sheet_light_reads_formula_once():
    sh = _Sheet("他", (("見出し", 10.0), ("=SUM(B1)", 3.0)))
    snap = vi._snap_sheet_light(sh)
    assert snap['light'] is True and snap['used'] == "A1:B2"
    assert snap['formulas'] == {"A2": "=SUM(B1)"}
    assert snap['values'][0] == ["見出し", "10"]
    assert snap['hidden'] == () and snap['tables'] == {} and snap['merged'] is None


def test_snapshot_book_light_for_sheets_not_own(monkeypatch):
    monkeypatch.setattr(vi, "_snap_sheet", lambda sh: {'used': 'A1:A1', 'values': [['x']], 'formulas': {}, 'hidden': (3,),
                                                       'cf': 0, 'validation': '', 'filter': '', 'merged': None,
                                                       'tables': {}, 'charts': 0})
    wb = _Book([_Sheet("自分", (("x",),)), _Sheet("他", (("y",),))])
    snap = vi._snapshot_book(wb, own={"自分"})
    assert 'light' not in snap['sheets']["自分"] and snap['sheets']["他"]['light'] is True
    full = vi._snapshot_book(wb)
    assert 'light' not in full['sheets']["他"]


def _base(values, formulas=None, **kw):
    s = {'used': 'A1:B2', 'values': values, 'formulas': formulas or {}, 'hidden': (), 'cf': 0, 'validation': '',
         'filter': '', 'merged': None, 'tables': {}, 'charts': 0}
    s.update(kw)
    return s


def test_inv_compare_light_other_sheet_detects_constant_change_but_not_recalc():
    before = {'sheets': {'own': _base([['1']]), 'other': _base([['a', '=SUM(own!A1)'], ['b', 'c']], {'B1': '=SUM(own!A1)'}, light=True)},
              'names': {}, 'order': ['own', 'other'], 'calc': None}
    now_same = {'sheets': {'own': _base([['2']]),
                           'other': _base([['a', '=SUM(own!A1)'], ['b', 'c']], {'B1': '=SUM(own!A1)'}, light=True)},
                'names': {}, 'order': ['own', 'other'], 'calc': None}
    assert vi._inv_compare(before, now_same, 'own') == []
    now_bad = {'sheets': {'own': _base([['2']]),
                          'other': _base([['a', '=SUM(own!A1)'], ['B', 'c']], {'B1': '=SUM(own!A1)'}, light=True)},
               'names': {}, 'order': ['own', 'other'], 'calc': None}
    bad = vi._inv_compare(before, now_bad, 'own')
    assert len(bad) == 1 and "other" in bad[0] and "A2" in bad[0]


def test_inv_compare_promoted_light_sheet_skips_hidden_check():
    # 控えでは軽かったシートに後から書いた（own に昇格）→ 隠れ行の照合は控えに無いので咎めない
    before = {'sheets': {'own': _base([['1']]), 'later': _base([['a']], light=True)},
              'names': {}, 'order': ['own', 'later'], 'calc': None}
    now = {'sheets': {'own': _base([['1']]), 'later': _base([['a']], hidden=(5,))},
           'names': {}, 'order': ['own', 'later'], 'calc': None}
    assert vi._inv_compare(before, now, {'own', 'later'}) == []


def test_inv_violations_reads_full_for_sheets_that_were_full(monkeypatch):
    calls = {}

    def fake_snapshot(wb, own=None):
        calls['own'] = set(own) if own is not None else None
        return {'sheets': {}, 'names': {}, 'order': [], 'calc': None}
    monkeypatch.setattr(vi, "_snapshot_book", fake_snapshot)
    before = {'sheets': {'a': _base([['1']]), 'b': _base([['1']], light=True), 'c': _base([['1']])},
              'names': {}, 'order': ['a', 'b', 'c'], 'calc': None}
    vi._inv_violations(object(), before, 'a')
    assert calls['own'] == {'a', 'c'}


# ================================================================
# フック（「表示」「発表」では発火しない）
# ================================================================

def _hook():
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))), ".claude", "scripts", "user_prompt_submit_excel_rule.py")
    if not os.path.exists(path):
        pytest.skip("フックのスクリプトがこの配置に無い（公開リポ側）")
    spec = importlib.util.spec_from_file_location("hook_excel_rule", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.mark.parametrize("prompt,want", [
    ("この表を直して", True), ("シートの D3 を見て", True), ("集計表を作って", True), ("マクロが動かない", True),
    ("表示が崩れる", False), ("発表の資料", False), ("代表に聞いて", False), ("キャンセルして", False),
    ("フェイスブックに投稿", False), ("表示の話と、ついでに表も整えて", True), ("エクセルを開いて", True),
])
def test_hook_wants_excel(prompt, want):
    assert _hook().wants_excel(prompt) is want


# ================================================================
# repair（マクロ修理の材料を 1 手で・2026-09-17）
# ================================================================

def _repair_parts(**over):
    parts = {'book': 'ポスター.xlsm', 'module': '写真貼り付け', 'proc': '写真の貼り付け', 'kind': 'Sub', 'n_lines': 3,
             'code': "Sub 写真の貼り付け()\n    x = 1\nEnd Sub\n", 'saved': '_last_proc.vba', 'shortcut': None,
             'buttons': [], 'form_callers': [], 'auto': False, 'callers': [], 'callers2': [], 'callees': [],
             'callees2': [], 'compile': {'ok': True, 'state': 'compiled', 'detail': '全体コンパイル OK'},
             'checks': [], 'backup': None}
    parts.update(over)
    return parts


def test_repair_report_order_and_untruncated_code():
    code = "Sub 写真の貼り付け()\n" + "\n".join(f"    x{i} = {i}" for i in range(300)) + "\nEnd Sub\n"
    parts = _repair_parts(
        code=code, n_lines=302, shortcut='Ctrl+Shift+K', buttons=['名前変換/Button 3'],
        form_callers=['メニューForm.CommandButton18_Click'], callers=[('メニューForm', 'CommandButton18_Click')],
        callees=[('図形測定', '基準設定ツール起動')], callees2=[('図形測定', '円描画')],
        compile={'ok': False, 'state': 'error',
                 'detail': '[写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo ― 行ラベルが定義されていません。'},
        checks=["VBM012 行 82: 飛び先ラベル 'Errorgo' がありません: On Error GoTo Errorgo"],
        backup={'label': 'backups/ポスター_写真貼り付け_20260916_205700.bas', 'minus': 1, 'plus': 2,
                'diff': ['--- a', '+++ b', '-x', '+y', '+z']})
    lines = dt._repair_report(parts)
    text = "\n".join(lines)
    heads = [ln.split('（')[0].split(':')[0] for ln in lines if ln.startswith('■ ')]
    assert heads == ['■ 入口', '■ 呼び元', '■ 呼び先', '■ コンパイル', '■ check', '■ 直前の控えとの差分'], "①〜⑥ の順"
    assert lines[1] == "モジュール  : 写真貼り付け" and lines[2] == "プロシージャ: 写真の貼り付け"
    assert "    x299 = 299" in text, "本文は切らない"
    assert "ショートカット Ctrl+Shift+K" in text and "シートのボタン 名前変換/Button 3" in text
    assert "フォーム メニューForm.CommandButton18_Click" in text
    assert "呼び元（直接 1・間接 0）: [メニューForm] CommandButton18_Click" in text
    assert "呼び先（直接 1・間接 1）: [図形測定] 基準設定ツール起動 ／ 間接: [図形測定] 円描画" in text
    assert "不通過（コンパイル エラー）  [写真貼り付け] 写真の貼り付け:82:" in text
    assert "check（error 級・このプロシージャ）: 1 件" in text and "VBM012 行 82" in text
    assert "-1 +2 行" in text and "  +y" in text
    assert lines[-1].startswith("直す:") and "code-replace" in lines[-1] and "replace-procedure -y --module 写真貼り付け" in lines[-1]
    import vbam_macro as vmac
    assert vmac._GET_MODULE_RE.search(text).group(1) == "写真貼り付け", "agent の材料がモジュール名を拾える形"


def test_repair_report_says_when_there_is_nothing():
    text = "\n".join(dt._repair_report(_repair_parts()))
    assert "（ショートカット・ボタン・フォームからは呼ばれていない" in text
    assert "呼び元（直接 0・間接 0）: （なし）" in text and "（なし＝単体で完結）" in text
    assert "通過（コンパイルしました）  全体コンパイル OK" in text
    assert "check（error 級・このプロシージャ）: 0 件（なし）" in text and "（控えなし）" in text
    missing = "\n".join(dt._repair_report(_repair_parts(backup={'label': 'backups/x.bas', 'minus': 0, 'plus': 0,
                                                                 'diff': [], 'missing': True})))
    assert "backups/x.bas にこのプロシージャは無い" in missing
    same = "\n".join(dt._repair_report(_repair_parts(backup={'label': 'backups/x.bas', 'minus': 0, 'plus': 0, 'diff': []})))
    assert "backups/x.bas と同じ" in same
    # 名前違いで呼んでいる（ポスターの ③: CommandButton18_Click が「長さ面積測定の基準設定ツール起動」を呼ぶ）
    near = "\n".join(dt._repair_report(_repair_parts(
        proc='基準設定ツール起動', near_calls=[('メニューForm', 'CommandButton18_Click', '長さ面積測定の基準設定ツール起動', 80)])))
    assert "名前違いで呼んでいる: [メニューForm] CommandButton18_Click:80 → 長さ面積測定の基準設定ツール起動（存在しない名前）" in near


def test_repair_wired_and_proc_body_from_bas():
    assert vm.raw_command(vm._command_table()["repair"]) is dt.cmd_repair
    ns = vm.build_parser().parse_args(["repair", "写真の貼り付け", "--module", "写真貼り付け"])
    assert ns.posargs == ["写真の貼り付け"] and ns.module_opt == "写真貼り付け"
    bas = ('Attribute VB_Name = "M"\r\nSub A()\r\nAttribute A.VB_ProcData.VB_Invoke_Func = "k\\n14"\r\n    x = 1\r\nEnd Sub\r\n'
           '\r\n\' 説明\r\nSub B()\r\n    y = 2\r\nEnd Sub\r\n')
    assert dt._proc_body_from_bas(bas, "a") == "Sub A()\n    x = 1\nEnd Sub"
    assert dt._proc_body_from_bas(bas, "B") == "' 説明\nSub B()\n    y = 2\nEnd Sub"
    assert dt._proc_body_from_bas(bas, "C") is None


def test_latest_module_backup_picks_newest_of_that_book_and_module(tmp_path, monkeypatch):
    monkeypatch.setattr(dt, "BACKUP_DIR", str(tmp_path))
    old = tmp_path / "ポスター_写真貼り付け_20260101_000000.bas"
    new = tmp_path / "ポスター_写真貼り付け_20260916_205700.bas"
    other = tmp_path / "ポスター_メニュー_20260917_000000.bas"
    for f in (old, new, other):
        f.write_bytes(b"x")
    os.utime(old, (time.time() - 100,) * 2)
    assert dt._latest_module_backup("ポスター", "写真貼り付け") == str(new)
    assert dt._latest_module_backup("ポスター", "無い") is None
    assert dt._latest_module_backup("別の本", "写真貼り付け") is None


# ================================================================
# patch-procedure & 書き込み検証のユニットテスト
# ================================================================

def test_patch_procedure_parser_and_wiring():
    # パーサーのパース検証
    p = vm.build_parser()
    ns = p.parse_args(["patch-procedure", "秀コンボ.xlsm", "テストSub",
                       "--target", "Dim x As Long", "--replacement", "Dim x As Double",
                       "--module", "Module1", "-y"])
    assert ns.posargs == ["秀コンボ.xlsm", "テストSub"]
    assert ns.target == "Dim x As Long"
    assert ns.replacement == "Dim x As Double"
    assert ns.module_opt == "Module1"
    assert ns.yes is True

    # コマンドテーブルと権限テーブルの配線検証
    assert "patch-procedure" in vm._command_table()
    import vbam_vba as vv
    assert vm.raw_command(vm._command_table()["patch-procedure"]) is vv.cmd_patch_procedure
    write_cmds = {name for name, _ in vv._CAPABILITIES["write"]}
    assert "patch-procedure" in write_cmds


def test_patch_procedure_logic_counts_and_validation(monkeypatch, capsys):
    import vbam_vba as vv

    # 1. target が未指定の場合
    class ArgsEmpty:
        posargs = ["テストSub"]
        module_opt = None
        target = None
        replacement = None
        target_file_opt = None
        replacement_file_opt = None
        yes = True

    assert vv.cmd_patch_procedure(ArgsEmpty()) is False
    out = capsys.readouterr().out
    assert "エラー: --target または --target-file で置換対象の文字列を指定してください" in out

    # 2. get_workbook と _extract_proc のモック
    dummy_code = "Sub テストSub()\n    Dim a As Long\n    Dim b As Long\n    a = 10\nEnd Sub\n"
    monkeypatch.setattr(vv, "get_workbook", lambda f: ("dummy_xl", "dummy_wb"))
    monkeypatch.setattr(vv, "_extract_proc", lambda wb, mod, name: ("Module1", dummy_code))

    # 3. target が見つからない (count == 0)
    class ArgsNotFound:
        posargs = ["テストSub"]
        module_opt = "Module1"
        target = "Dim c As Long"
        replacement = "Dim c As Double"
        target_file_opt = None
        replacement_file_opt = None
        yes = True

    assert vv.cmd_patch_procedure(ArgsNotFound()) is False
    out = capsys.readouterr().out
    assert "target 文字列がプロシージャ 'テストSub' 内に見つかりません" in out

    # 4. target が複数箇所見つかる (count > 1)
    class ArgsMultiple:
        posargs = ["テストSub"]
        module_opt = "Module1"
        target = "Dim "
        replacement = "Dim "
        target_file_opt = None
        replacement_file_opt = None
        yes = True

    assert vv.cmd_patch_procedure(ArgsMultiple()) is False
    out = capsys.readouterr().out
    assert "複数（2箇所）見つかりました" in out

    # 5. 一意に特定できて正常置換（replace-procedure への引き渡し確認）
    replaced_args = []
    monkeypatch.setattr(vv, "cmd_replace_procedure", lambda args: replaced_args.append(args) or True)

    class ArgsSuccess:
        posargs = ["テストSub"]
        module_opt = "Module1"
        target = "Dim a As Long"
        replacement = "Dim a As Double"
        target_file_opt = None
        replacement_file_opt = None
        yes = True

    assert vv.cmd_patch_procedure(ArgsSuccess()) is True
    assert len(replaced_args) == 1
    assert replaced_args[0].module_opt == "Module1"
    assert replaced_args[0].yes is True
    # _last_proc.vba に書き込まれた内容を確認
    with open(vv.LAST_PROC_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    assert "Dim a As Double" in content
    assert "Dim b As Long" in content


def test_report_write_result_error_and_hash_detection(capsys):
    import vbam_edit as ve

    class DummyCell:
        def __init__(self, addr, text):
            self.Address = addr
            self.Text = text

    class DummyCells:
        def __init__(self, cell_list):
            self._list = cell_list

        @property
        def CountLarge(self):
            return len(self._list)

        def __iter__(self):
            return iter(self._list)

    class DummyRange:
        def __init__(self, cells, errors=None):
            self.Cells = DummyCells(cells)
            self._errors = errors or []

        def SpecialCells(self, typ, flag):
            if typ == -4123 and self._errors:
                return self._errors
            raise Exception("No cells found")

    # 1. 正常なセル
    r_ok = DummyRange([DummyCell("$A$1", "100"), DummyCell("$B$1", "200")])
    ve._report_write_result(None, r_ok)
    out = capsys.readouterr().out
    assert "⚠" not in out

    # 2. 数式エラー検出
    err_cell = DummyCell("$A$1", "#REF!")
    r_err = DummyRange([err_cell], errors=[err_cell])
    ve._report_write_result(None, r_err)
    out = capsys.readouterr().out
    assert "⚠ 【要修正】書き込み結果に数式エラーがあります: 1件 （A1=#REF!）" in out

    # 3. 列幅不足（###）検出
    hash_cell = DummyCell("$C$1", "#####")
    r_hash = DummyRange([hash_cell])
    ve._report_write_result(None, r_hash)
    out = capsys.readouterr().out
    assert "⚠ 【表示崩れ】列幅不足により '###' 表示になっているセルがあります: 1件 （C1(幅不足)）" in out

