# -*- coding: utf-8 -*-
"""関所（gate）が写しを開いても、呼び出し台帳（_calls.jsonl）の book が呼び出しの対象のままになる試験（2026-10-29）。

register-addin の焼く前の gate で写しを開くと _last_book_name が写しの名前に書き換わり、台帳の book に
秀コンボ_関所_20261029_000514.xlsm と残った。Excel は使わない（get_workbook などを差し替える）。"""
import argparse
import json
import os
import tempfile

import vbam_core
import vbam_vba


class _FakeWb:
    def __init__(self, name):
        self.Name = name

    def SaveCopyAs(self, path):
        with open(path, 'wb') as f:
            f.write(b'copy')

    def Close(self, SaveChanges=False):
        pass


class _FakeWorkbooks:
    def Open(self, path, update_links=0):
        return _FakeWb(os.path.basename(path))


class _FakeXl:
    Hwnd = 0

    def __init__(self):
        self.Workbooks = _FakeWorkbooks()

    def Quit(self):
        pass


class _FakeWatchdog:
    fired = False
    seconds = None

    def stop(self):
        pass


def _rig(monkeypatch, tmp_path, src):
    """gate の COM の所を差し替える。check／test は本物と同じく get_workbook で写しを掴む。掴んだ名前を返す。"""
    seen = []

    def fake_get_workbook(target=None, load_addins=False, readonly=False):
        if target and os.path.abspath(str(target)).lower() in vbam_core._wb_cache:
            xl, wb, _ro = vbam_core._wb_cache[os.path.abspath(str(target)).lower()]
        else:
            xl, wb = _FakeXl(), src
        vbam_core._note_book(wb)
        seen.append(str(wb.Name))
        return xl, wb

    def fake_check(args):
        fake_get_workbook(args.posargs[0])
        print(json.dumps({"summary": {"errors": 0, "warnings": 0}, "modules": []}))

    def fake_test(args):
        fake_get_workbook(args.posargs[0])
        print(json.dumps({"total": 1, "passed": 1, "tests": [{"ok": True, "name": "テスト1", "module": "m"}]}))

    monkeypatch.setattr(vbam_vba, "get_workbook", fake_get_workbook)
    monkeypatch.setattr(vbam_vba, "cmd_check", fake_check)
    monkeypatch.setattr(vbam_vba, "cmd_test", fake_test)
    monkeypatch.setattr(vbam_vba, "_compile_vbproject", lambda xl, wb: {"ok": True, "detail": "ok", "state": "compiled"})
    monkeypatch.setattr(vbam_vba, "_blank_book_names", lambda xl: [])
    monkeypatch.setattr(vbam_vba, "_close_new_blank_books", lambda xl, before: [])
    monkeypatch.setattr(vbam_vba, "_ghost_projects", lambda xl: 0)
    monkeypatch.setattr(vbam_vba, "_start_call_watchdog", lambda seconds, owned_pid=None: _FakeWatchdog())
    monkeypatch.setattr(vbam_vba.win32com.client, "DispatchEx", lambda prog: _FakeXl(), raising=False)
    monkeypatch.setattr(vbam_core, "is_fresh_instance", lambda xl: True)
    monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))      # 写しは tmp_path に作る
    return seen


def test_gate_keeps_the_target_book_name_for_the_call_log(monkeypatch, tmp_path):
    seen = _rig(monkeypatch, tmp_path, _FakeWb("秀コンボ.xlsm"))
    monkeypatch.setattr(vbam_core, "_last_book_name", None)
    ok = vbam_vba.cmd_gate(argparse.Namespace(posargs=["秀コンボ.xlsm"], json=False, timeout=120, keep=False,
                                              addins=False, module=None))
    assert ok is True
    assert any("_関所_" in n for n in seen)                 # 写しは掴まれた（前はここで台帳の名前が写しになった）
    assert vbam_core._last_book_name == "秀コンボ.xlsm"
    assert not any("_関所_" in f for f in os.listdir(tmp_path))   # 写しは消えている


def test_register_addin_test_gate_logs_the_burned_book(monkeypatch, tmp_path):
    src = _FakeWb("秀コンボ.xlsm")
    seen = _rig(monkeypatch, tmp_path, src)
    monkeypatch.setattr(vbam_vba, "_has_test_subs", lambda wb: True)
    log = tmp_path / "calls.jsonl"
    monkeypatch.setattr(vbam_core, "_CALL_LOG_FILE", str(log))
    monkeypatch.setattr(vbam_core, "_last_book_name", None)      # 呼び出しの入口で戻すのと同じ
    vbam_vba.get_workbook(None, load_addins=True)                 # register-addin が対象を掴む
    assert vbam_vba._test_gate_before_burn(src) is True
    assert any("_関所_" in n for n in seen)
    vbam_core.call_log_write("register-addin", argparse.Namespace(posargs=[]), 1.0, True, via="mcp")
    rec = json.loads(log.read_text(encoding="utf-8").splitlines()[-1])
    assert rec["book"] == "秀コンボ.xlsm"


def test_gate_restores_the_book_name_even_when_the_check_fails(monkeypatch, tmp_path):
    seen = _rig(monkeypatch, tmp_path, _FakeWb("台帳.xlsm"))

    def broken_check(args):
        vbam_vba.get_workbook(args.posargs[0])
        raise RuntimeError("check で落ちた")

    monkeypatch.setattr(vbam_vba, "cmd_check", broken_check)
    monkeypatch.setattr(vbam_core, "_last_book_name", None)
    ok = vbam_vba.cmd_gate(argparse.Namespace(posargs=["台帳.xlsm"], json=False, timeout=120, keep=False,
                                              addins=False, module=None))
    assert ok is False
    assert any("_関所_" in n for n in seen)
    assert vbam_core._last_book_name == "台帳.xlsm"


def test_gate_changed_zero_tests_skips_copy_and_dispatchex(monkeypatch, tmp_path):
    """指示書086: 関わるテスト 0 本のときは SaveCopyAs も DispatchEx も呼ばれない"""
    src = _FakeWb("秀コンボ.xlsm")
    seen = _rig(monkeypatch, tmp_path, src)

    copy_called = []
    def fail_save_copy(path):
        copy_called.append(path)
        raise AssertionError("SaveCopyAs が呼ばれました（0本なら呼ばれてはいけない）")
    src.SaveCopyAs = fail_save_copy

    dispatch_called = []
    def fail_dispatch(prog):
        dispatch_called.append(prog)
        raise AssertionError("DispatchEx が呼ばれました（0本なら呼ばれてはいけない）")
    monkeypatch.setattr(vbam_vba.win32com.client, "DispatchEx", fail_dispatch, raising=False)

    monkeypatch.setattr(vbam_vba, "burn_marks_load", lambda name: {"time": "2026-10-29 00:00:00", "modules": {"m1": "hash_old"}})
    monkeypatch.setattr(vbam_vba, "_module_codes", lambda wb: {"m1": "Sub NotATest()\nEnd Sub\n"})
    monkeypatch.setattr(vbam_vba, "changed_modules", lambda marks, codes: ["m1"])
    monkeypatch.setattr(vbam_vba, "changed_procs", lambda marks, codes: ["NotATest"])
    monkeypatch.setattr(vbam_vba, "related_tests", lambda codes, ch, tnames, procs=None: [])

    ok = vbam_vba.cmd_gate(argparse.Namespace(posargs=["秀コンボ.xlsm"], json=False, timeout=120, keep=False,
                                              addins=False, module=None, changed=True))
    assert ok is True
    assert len(copy_called) == 0
    assert len(dispatch_called) == 0
    assert not any("_関所_" in n for n in seen)


def test_gate_changed_zero_tests_fails_when_compile_fails(monkeypatch, tmp_path):
    """指示書086: 関わるテスト 0 本でもコンパイルエラーなら止まる"""
    src = _FakeWb("秀コンボ.xlsm")
    seen = _rig(monkeypatch, tmp_path, src)

    monkeypatch.setattr(vbam_vba, "burn_marks_load", lambda name: {"time": "2026-10-29 00:00:00", "modules": {"m1": "hash_old"}})
    monkeypatch.setattr(vbam_vba, "_module_codes", lambda wb: {"m1": "Sub SyntaxErr()\nEnd Sub\n"})
    monkeypatch.setattr(vbam_vba, "changed_modules", lambda marks, codes: ["m1"])
    monkeypatch.setattr(vbam_vba, "changed_procs", lambda marks, codes: ["SyntaxErr"])
    monkeypatch.setattr(vbam_vba, "related_tests", lambda codes, ch, tnames, procs=None: [])
    monkeypatch.setattr(vbam_vba, "_compile_vbproject", lambda xl, wb: {"ok": False, "detail": "コンパイル エラー: 構文エラー", "state": "error"})

    ok = vbam_vba.cmd_gate(argparse.Namespace(posargs=["秀コンボ.xlsm"], json=False, timeout=120, keep=False,
                                              addins=False, module=None, changed=True))
    assert ok is False
    assert not any("_関所_" in n for n in seen)


def test_shelf_parity_gate_skips_when_shelf_head_unchanged(monkeypatch):
    """指示書087: 棚の頭が前と同じなら shelf--check が呼ばれない"""
    import vbam_prefire
    called = []
    monkeypatch.setattr(vbam_prefire, "shelf_check", lambda *args, **kwargs: (called.append(True) or (True, [])))

    code = "Sub 棚1()\n  ' 依頼の語: テスト頼み\n  ' 依頼の組: テスト,頼み\n  x = 1\nEnd Sub\n"
    codes = {"棚_テスト": code}
    sig = vbam_vba.shelf_signature_from_codes(codes)
    esig = [["exam.tsv", 100, 200]]
    monkeypatch.setattr(vbam_vba, "exam_signature", lambda: esig)

    marks = {"shelf_sig": sig, "exam_sig": esig}

    new_code = "Sub 棚1()\n  ' 依頼の語: テスト頼み\n  ' 依頼の組: テスト,頼み\n  x = 999\nEnd Sub\n"
    new_codes = {"棚_テスト": new_code}

    wb = _FakeWb("秀コンボ.xlsm")
    wb.VBProject = argparse.Namespace(VBComponents=lambda name: True)
    monkeypatch.setattr(vbam_core, "shelf_text", lambda proj: new_code)

    ok = vbam_vba._shelf_parity_gate(None, wb, marks=marks, codes=new_codes)
    assert ok is True
    assert len(called) == 0


def test_shelf_parity_gate_runs_when_shelf_head_changed(monkeypatch):
    """指示書087: 頼みの語を 1 つ変えたら shelf--check が呼ばれる"""
    import vbam_prefire
    called = []
    monkeypatch.setattr(vbam_prefire, "shelf_check", lambda *args, **kwargs: (called.append(True) or (True, [])))

    code = "Sub 棚1()\n  ' 依頼の語: テスト頼み\n  ' 依頼の組: テスト,頼み\n  x = 1\nEnd Sub\n"
    codes = {"棚_テスト": code}
    sig = vbam_vba.shelf_signature_from_codes(codes)
    esig = [["exam.tsv", 100, 200]]
    monkeypatch.setattr(vbam_vba, "exam_signature", lambda: esig)

    marks = {"shelf_sig": sig, "exam_sig": esig}

    new_code = "Sub 棚1()\n  ' 依頼の語: 別の頼み\n  ' 依頼の組: テスト,頼み\n  x = 1\nEnd Sub\n"
    new_codes = {"棚_テスト": new_code}

    wb = _FakeWb("秀コンボ.xlsm")
    wb.VBProject = argparse.Namespace(VBComponents=lambda name: True)
    monkeypatch.setattr(vbam_core, "shelf_text", lambda proj: new_code)

    ok = vbam_vba._shelf_parity_gate(None, wb, marks=marks, codes=new_codes)
    assert ok is True
    assert len(called) == 1


def test_shelf_parity_gate_runs_when_exam_changed(monkeypatch):
    """指示書087: 問題集が変わったら shelf--check が呼ばれる"""
    import vbam_prefire
    called = []
    monkeypatch.setattr(vbam_prefire, "shelf_check", lambda *args, **kwargs: (called.append(True) or (True, [])))

    code = "Sub 棚1()\n  ' 依頼の語: テスト頼み\n  ' 依頼の組: テスト,頼み\n  x = 1\nEnd Sub\n"
    codes = {"棚_テスト": code}
    sig = vbam_vba.shelf_signature_from_codes(codes)
    esig = [["exam.tsv", 100, 200]]
    monkeypatch.setattr(vbam_vba, "exam_signature", lambda: [["exam.tsv", 200, 200]])

    marks = {"shelf_sig": sig, "exam_sig": esig}

    wb = _FakeWb("秀コンボ.xlsm")
    wb.VBProject = argparse.Namespace(VBComponents=lambda name: True)
    monkeypatch.setattr(vbam_core, "shelf_text", lambda proj: code)

    ok = vbam_vba._shelf_parity_gate(None, wb, marks=marks, codes=codes)
    assert ok is True
    assert len(called) == 1


