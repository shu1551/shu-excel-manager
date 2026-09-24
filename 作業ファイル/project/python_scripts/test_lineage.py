# -*- coding: utf-8 -*-
"""vbam_lineage（2026-09-17 系譜と閉じたブック）のテスト。Excel には接続しない。

実行: このフォルダで `py -m pytest test_lineage.py -q`。
versions／list-file／grep-files／export-file は fixtures の本物の 1 冊（ポスター_コンパイルエラー3件.xlsm）を
oletools で読む（無ければ skip）。history は tmp に置いた .bas 3 世代で「同じ本文を畳む・差分が出る」を見る。
"""
import os
import sys
import time
import json
import argparse

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vba_manager as vm
import vbam_lineage as lg

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'fixtures', 'ポスター_コンパイルエラー3件.xlsm')
_needs_fixture = pytest.mark.skipif(not os.path.exists(FIXTURE), reason="fixtures のポスター_コンパイルエラー3件.xlsm が無い")


# ================================================================
# 読み手（純 Python）
# ================================================================

def test_strip_export_header_bas_cls_frm():
    bas = 'Attribute VB_Name = "M"\r\nSub A()\r\nAttribute A.VB_ProcData.VB_Invoke_Func = "k\\n14"\r\n    x = 1\r\nEnd Sub\r\n'
    assert lg._strip_export_header(bas) == ['Sub A()', '    x = 1', 'End Sub']
    cls = ('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\nAttribute VB_Name = "Sheet1"\r\n'
           'Attribute VB_GlobalNameSpace = False\r\nOption Explicit\r\nPrivate Sub Worksheet_Change()\r\nEnd Sub\r\n')
    assert lg._strip_export_header(cls) == ['Option Explicit', 'Private Sub Worksheet_Change()', 'End Sub']
    frm = ('VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9B7B-000000000000} F \r\n   Caption         =   "x"\r\n'
           '   ClientHeight    =   100\r\nEnd\r\nAttribute VB_Name = "F"\r\nAttribute VB_Exposed = False\r\n'
           'Private Sub CommandButton1_Click()\r\n    Unload Me\r\nEnd Sub\r\n')
    assert lg._strip_export_header(frm) == ['Private Sub CommandButton1_Click()', '    Unload Me', 'End Sub']
    assert lg._strip_export_header('') == []


def test_module_type_label_detects_sheet_and_workbook_modules():
    assert lg._module_type_label('.bas', '') == '標準'
    assert lg._module_type_label('.frm', '') == 'フォーム'
    assert lg._module_type_label('.cls', 'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"') == 'シート/ブック'
    assert lg._module_type_label('.cls', 'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"') == 'シート/ブック'
    assert lg._module_type_label('.cls', 'Attribute VB_Creatable = False') == 'クラス'


def test_stamp_of_prefers_name_then_folder_then_mtime(tmp_path):
    f = tmp_path / "Book_M_20260916_205700.bas"
    f.write_bytes(b"x")
    assert time.strftime('%Y%m%d_%H%M%S', time.localtime(lg._stamp_of(str(f)))) == "20260916_205700"
    d = tmp_path / "20260101_010203"
    d.mkdir()
    g = d / "M.bas"
    g.write_bytes(b"x")
    assert time.strftime('%Y%m%d_%H%M%S', time.localtime(lg._stamp_of(str(g)))) == "20260101_010203"
    h = tmp_path / "plain.bas"
    h.write_bytes(b"x")
    assert abs(lg._stamp_of(str(h)) - os.path.getmtime(str(h))) < 1


def test_diff_against_latest_names_changed_added_removed():
    rows = [{'path': 'new', 'mtime': 3, 'error': None, 'code': {'A': 'a2', 'B': 'b', 'C': 'c'}},
            {'path': 'old', 'mtime': 2, 'error': None, 'code': {'A': 'a1', 'B': 'b', 'X': 'x'}},
            {'path': 'bad', 'mtime': 1, 'error': 'broken', 'code': {}}]
    out = lg._diff_against_latest(rows)
    assert out[0]['changed'] == [] and out[1]['changed'] == ['A']
    assert out[1]['added'] == ['X'] and out[1]['removed'] == ['C']
    assert out[2]['changed'] == []


# ================================================================
# history（tmp の .bas 3 世代 → 2 版に畳む・差分が出る）
# ================================================================

def _bas(body_lines):
    return ('Attribute VB_Name = "M"\r\nSub 他()\r\nEnd Sub\r\nSub 集計()\r\n' + ''.join(ln + '\r\n' for ln in body_lines)
            + 'End Sub\r\n').encode('cp932')


def test_history_folds_identical_bodies_and_diffs(tmp_path, monkeypatch, capsys):
    bdir = tmp_path / "backups"
    bdir.mkdir()
    monkeypatch.setattr(lg, "BACKUP_DIR", str(bdir))
    monkeypatch.setattr(lg, "SCRIPT_DIR", str(tmp_path))          # _exports は tmp の下（無い）
    (bdir / "Book_M_20260101_000000.bas").write_bytes(_bas(['    x = 1']))
    (bdir / "Book_M_20260102_000000.bas").write_bytes(_bas(['    x = 1']))          # 同じ本文
    (bdir / "Book_M_20260103_000000.bas").write_bytes(_bas(['    x = 1', '    y = 2']))
    (bdir / "Other_M_20260104_000000.bas").write_bytes(_bas(['    z = 9']))        # 別ブック
    (bdir / "Book_N_20260105_000000.bas").write_bytes('Attribute VB_Name = "N"\r\nSub 無関係()\r\nEnd Sub\r\n'.encode('cp932'))
    versions, n_files = lg._history_versions("集計", book="Book")
    assert n_files == 3 and len(versions) == 2
    assert versions[0]['dup'] == 2 and versions[0]['n_lines'] == 3
    assert versions[1]['body'].endswith("    y = 2\nEnd Sub") and versions[1]['dup'] == 1
    assert versions[0]['time'] < versions[1]['time']
    ns = argparse.Namespace(posargs=["集計"], book_opt="Book", deep=False, dir_opt=None, max_hits=None, json=False)
    assert lg.cmd_history(ns) is True
    out = capsys.readouterr().out
    assert "3 件のファイル → 2 版" in out and "同じ本文 他 1 件" in out
    assert "-0 +1 行" in out and "    +    y = 2" in out
    # 絞らなければ別ブックの分も入る（3 版）
    versions_all, n_all = lg._history_versions("集計")
    assert n_all == 4 and len(versions_all) == 3


def test_history_deep_needs_book_and_says_when_nothing_found(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(lg, "BACKUP_DIR", str(tmp_path / "nothing"))
    monkeypatch.setattr(lg, "SCRIPT_DIR", str(tmp_path))
    ns = argparse.Namespace(posargs=["無い"], book_opt=None, deep=True, dir_opt=None, max_hits=None, json=False)
    assert lg.cmd_history(ns) is False
    assert "--book" in capsys.readouterr().out
    ns.deep = False
    assert lg.cmd_history(ns) is True
    assert "控えがありません" in capsys.readouterr().out


# ================================================================
# 閉じたブック（fixtures の本物の 1 冊を oletools で読む）
# ================================================================

@_needs_fixture
def test_read_closed_book_reads_japanese_module_names():
    mods = lg._read_closed_book(FIXTURE)
    for name in ("メニュー", "写真貼り付け", "図形複写", "メニューForm"):
        assert name in mods, name
    assert "Sub 写真の貼り付け()" in mods["写真貼り付け"]
    assert not any(ln.startswith("Attribute ") for ln in mods["写真貼り付け"].split("\n"))


@_needs_fixture
def test_list_file_lists_modules(capsys):
    ns = argparse.Namespace(posargs=[FIXTURE], json=True)
    assert lg.cmd_list_file(ns) is True
    doc = json.loads(capsys.readouterr().out)
    names = {m['module']: m for m in doc['modules']}
    assert doc['has_macros'] and {"メニュー", "写真貼り付け", "図形複写"} <= set(names)
    assert names["メニューForm"]['type'] == 'フォーム' and names["写真貼り付け"]['type'] == '標準'
    assert names["ThisWorkbook"]['type'] == 'シート/ブック'
    assert "写真の貼り付け" in names["写真貼り付け"]['procs']
    ns.json = False
    assert lg.cmd_list_file(ns) is True
    out = capsys.readouterr().out
    assert "閉じたブック:" in out and "[フォーム] メニューForm" in out and "Excel は開いていない" in out


@_needs_fixture
def test_grep_files_finds_errorgo_with_vbe_line_number(capsys):
    # 練習台の欠陥①: 写真貼り付け.写真の貼り付け 82 行目 `On Error GoTo Errorgo`（行番号は VBE の CodeModule と同じ）
    ns = argparse.Namespace(posargs=["Errorgo", os.path.dirname(FIXTURE)], regex=False, ignore_case=False,
                            max_hits=None, json=False)
    assert lg.cmd_grep_files(ns) is True
    out = capsys.readouterr().out
    assert "[ポスター_コンパイルエラー3件.xlsm][写真貼り付け] 写真の貼り付け:82: On Error GoTo Errorgo" in out
    assert "--- 1 件（1 冊を読んだ・ヒット 1 冊" in out
    ns = argparse.Namespace(posargs=["errorgo", FIXTURE], regex=False, ignore_case=True, max_hits=None, json=True)
    assert lg.cmd_grep_files(ns) is True
    doc = json.loads(capsys.readouterr().out)
    assert doc['total'] == 1 and doc['hits'][0]['line'] == 82 and doc['hits'][0]['proc'] == "写真の貼り付け"


@_needs_fixture
def test_grep_files_reports_missing_targets_and_regex(capsys):
    ns = argparse.Namespace(posargs=["On\\s+Err\\s+GoTo", FIXTURE, "no_such_dir_xyz"], regex=True, ignore_case=False,
                            max_hits=None, json=False)
    assert lg.cmd_grep_files(ns) is True
    out = capsys.readouterr().out
    assert "見つかりません: no_such_dir_xyz" in out
    assert "[図形一覧form] 名前_Click:" in out and "On Err GoTo koko" in out          # 練習台の欠陥②


@_needs_fixture
def test_versions_lists_copies_newest_first_with_diff_note(tmp_path, monkeypatch, capsys):
    import shutil
    a = tmp_path / "ポスター_a.xlsm"
    b = tmp_path / "ポスター_b.xlsm"
    shutil.copy2(FIXTURE, a)                       # copy2 は fixture の日時（6/14）を引き継ぐ
    shutil.copy2(FIXTURE, b)
    os.utime(b, (time.time(),) * 2)                 # b を最新にする
    monkeypatch.setattr(lg, "_version_dirs", lambda extra=(): [str(tmp_path)])
    monkeypatch.setattr(lg, "_open_book_paths", lambda: {os.path.normcase(str(b))})
    monkeypatch.setattr(lg, "BACKUP_DIR", str(tmp_path / "no_backups"))
    ns = argparse.Namespace(posargs=["ポスター"], dir_opt=None, all=False, max_hits=None, json=True)
    assert lg.cmd_versions(ns) is True
    doc = json.loads(capsys.readouterr().out)
    assert doc['count'] == 2 and doc['hidden_backups'] == 0
    assert doc['versions'][0]['path'].endswith("ポスター_b.xlsm") and doc['versions'][0]['open'] is True
    assert doc['versions'][1]['changed'] == [] and doc['versions'][1]['modules'] >= 20
    ns.json = False
    assert lg.cmd_versions(ns) is True
    out = capsys.readouterr().out
    assert "2 冊" in out and "← 最新" in out and "[開いている]" in out and "最新とコードは同じ" in out


@_needs_fixture
def test_versions_reads_only_the_newest_backups_unless_all(tmp_path, monkeypatch, capsys):
    """backups の控えは百冊単位（ポスターで 127 冊）＝既定は新しい 3 冊だけ読んで残りは数だけ言う。"""
    import shutil
    bdir = tmp_path / "backups"
    bdir.mkdir()
    for i in range(5):
        p = bdir / f"ポスター.xlsm.backup_before_x_2026090{i + 1}_000000.xlsm"
        shutil.copy2(FIXTURE, p)
        os.utime(p, (time.time() - (5 - i) * 3600,) * 2)
    monkeypatch.setattr(lg, "BACKUP_DIR", str(bdir))
    monkeypatch.setattr(lg, "_version_dirs", lambda extra=(): [str(bdir)])
    monkeypatch.setattr(lg, "_open_book_paths", lambda: set())
    reads = []
    real = lg._read_closed_modules
    monkeypatch.setattr(lg, "_read_closed_modules", lambda p: reads.append(p) or real(p))
    ns = argparse.Namespace(posargs=["ポスター"], dir_opt=None, all=False, max_hits=None, json=False)
    assert lg.cmd_versions(ns) is True
    out = capsys.readouterr().out
    assert len(reads) == 3 and "3 冊" in out and "他 2 冊の控え" in out and "--all" in out
    reads.clear()
    ns.all = True
    assert lg.cmd_versions(ns) is True
    assert len(reads) == 5 and "5 冊" in capsys.readouterr().out
    reads.clear()
    ns.all, ns.max_hits = False, 1
    assert lg.cmd_versions(ns) is True
    assert len(reads) == 1 and "他 4 冊の控え" in capsys.readouterr().out


@_needs_fixture
def test_export_file_writes_cp932_modules(tmp_path, capsys):
    ns = argparse.Namespace(posargs=[FIXTURE], dir_opt=str(tmp_path / "out"), json=True)
    assert lg.cmd_export_file(ns) is True
    doc = json.loads(capsys.readouterr().out)
    assert doc['count'] >= 20 and "写真貼り付け.bas" in doc['files'] and "メニューForm.frm" in doc['files']
    raw = (tmp_path / "out" / "写真貼り付け.bas").read_bytes()
    assert raw.startswith("Attribute VB_Name = \"写真貼り付け\"".encode('cp932'))
    assert b"\r\n" in raw and b"\r\r\n" not in raw


def test_expand_targets_folders_files_and_missing(tmp_path):
    (tmp_path / "a.xlsm").write_bytes(b"x")
    (tmp_path / "~$a.xlsm").write_bytes(b"x")
    (tmp_path / "b.txt").write_bytes(b"x")
    files, missing = lg._expand_targets([str(tmp_path), str(tmp_path / "a.xlsm"), "nope.xlsm"])
    assert [os.path.basename(f) for f in files] == ["a.xlsm"] and missing == ["nope.xlsm"]


# ================================================================
# 配線
# ================================================================

def test_lineage_commands_are_wired():
    table = vm._command_table()
    for name, fn in (("versions", lg.cmd_versions), ("history", lg.cmd_history), ("list-file", lg.cmd_list_file),
                     ("grep-files", lg.cmd_grep_files), ("export-file", lg.cmd_export_file)):
        assert vm.raw_command(table[name]) is fn, name
    p = vm.build_parser()
    ns = p.parse_args(["history", "集計", "--book", "Book", "--deep", "--max", "10"])
    assert ns.book_opt == "Book" and ns.deep is True and ns.max_hits == 10
    ns = p.parse_args(["grep-files", "x", "a", "b", "-i", "--regex"])
    assert ns.posargs == ["x", "a", "b"] and ns.ignore_case and ns.regex
    ns = p.parse_args(["versions", "ポスター", "--dir", "a", "--dir", "b", "--all"])
    assert ns.dir_opt == ["a", "b"] and ns.all is True and ns.max_hits is None


@_needs_fixture
def test_export_file_only_named_modules(tmp_path, capsys):
    """export-file <ファイル> <モジュール> でそれだけ書き出す（前は「余分な引数」で断った・2026-09-24）。"""
    ns = argparse.Namespace(posargs=[FIXTURE, "写真貼り付け"], dir_opt=str(tmp_path / "out"), json=True)
    assert lg.cmd_export_file(ns) is True
    doc = json.loads(capsys.readouterr().out)
    assert doc['files'] == ["写真貼り付け.bas"]
    ns = argparse.Namespace(posargs=[FIXTURE, "無いモジュール"], dir_opt=str(tmp_path / "out2"), json=False)
    assert lg.cmd_export_file(ns) is False
    assert "無いモジュール" in capsys.readouterr().out
