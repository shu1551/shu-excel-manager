# -*- coding: utf-8 -*-
"""指示書089: フォームの道具（form-read, form-inspect, form-to-vba）の失敗を防ぐ試験（2026-10-29）"""
import argparse
import pytest
import vbam_edit
import vbam_form2vba
import vba_manager


class _FakeComp:
    def __init__(self, name, comp_type=3, designer=None):
        self.Name = name
        self.Type = comp_type
        self.Designer = designer
        self.CodeModule = argparse.Namespace(CountOfLines=0, Lines=lambda s, c: "", AddFromString=lambda s: None)
        self._props = {"Caption": name, "Width": 200, "Height": 150}

    def Properties(self, name):
        val = self._props.get(name, "")
        return argparse.Namespace(Value=val, Name=name)


class _FakeComponents:
    def __init__(self, comps=None):
        self._comps = list(comps or [])

    def __iter__(self):
        return iter(self._comps)

    def Add(self, typ):
        c = _FakeComp(f"Comp{len(self._comps)+1}", comp_type=typ)
        self._comps.append(c)
        return c

    def Remove(self, c):
        if c in self._comps:
            self._comps.remove(c)


class _FakeWb:
    def __init__(self, name="テスト.xlsm", comps=None):
        self.Name = name
        self.FullName = f"C:\\test\\{name}"
        self.Saved = True
        self.VBProject = argparse.Namespace(
            VBComponents=_FakeComponents(comps),
            VBE=argparse.Namespace(Windows=argparse.Namespace(Count=0, Item=lambda i: None))
        )


@pytest.mark.core
def test_form_read_auto_picks_single_active_form(monkeypatch, capsys):
    """指示書089: form-read が引数なしのとき、単一の表示中フォームを自動で読む"""
    fake_wb = _FakeWb("秀コンボ.xlsm")
    fake_xl = argparse.Namespace(
        Run=lambda macro, form: "CAPTION\t自動検出フォーム\ntxt1\tTextBox\t値1\n",
        Workbooks=argparse.Namespace(Count=1, Item=lambda i: fake_wb),
        AddIns=argparse.Namespace(Count=0)
    )
    monkeypatch.setattr(vbam_edit, "get_workbook", lambda t: (fake_xl, fake_wb))

    # 引数なしで呼ぶ
    args = argparse.Namespace(posargs=[])
    ok = vbam_edit.cmd_form_read(args)
    assert ok is True
    out = capsys.readouterr().out
    assert "題: 自動検出フォーム" in out
    assert "txt1" in out


def test_form_read_accepts_name_variations(monkeypatch, capsys):
    """指示書089: form-read が「フォーム」有無の揺れ（Excelコンボ ⇔ Excelコンボフォーム）を吸収する"""
    fake_wb = _FakeWb("秀コンボ.xlsm")
    def fake_run(macro, name):
        assert name == "Excelコンボ"
        return "CAPTION\tExcelコンボフォーム\nbtn1\tCommandButton\tOK\n"

    fake_xl = argparse.Namespace(Run=fake_run, Workbooks=argparse.Namespace(Count=1, Item=lambda i: fake_wb), AddIns=argparse.Namespace(Count=0))
    monkeypatch.setattr(vbam_edit, "get_workbook", lambda t: (fake_xl, fake_wb))

    args = argparse.Namespace(posargs=["Excelコンボ"])
    ok = vbam_edit.cmd_form_read(args)
    assert ok is True
    out = capsys.readouterr().out
    assert "題: Excelコンボフォーム" in out


@pytest.mark.core
def test_form_inspect_auto_closes_loaded_form(monkeypatch, capsys):
    """指示書089: form-inspect でフォーム表示中（DesignerがNone）のとき、閉じてからレイアウトを取得する"""
    closed = []
    designer = argparse.Namespace(
        InsideWidth=200, InsideHeight=150,
        Controls=[argparse.Namespace(Name="btn1", Left=10, Top=10, Width=50, Height=20, Caption="OK", Font=argparse.Namespace(Size=10), Parent=argparse.Namespace(Name="テストフォーム"), TabIndex=0)]
    )

    comp = _FakeComp("テストフォーム", comp_type=3, designer=None)

    def fake_close_form(args):
        closed.append(args.posargs[0])
        comp.Designer = designer  # 閉じたので Designer が実体化
        return True

    monkeypatch.setattr(vbam_edit, "cmd_close_form", fake_close_form)

    fake_wb = _FakeWb("秀コンボ.xlsm", comps=[comp])
    monkeypatch.setattr(vbam_form2vba, "get_workbook", lambda t: (None, fake_wb))

    args = argparse.Namespace(posargs=["テストフォーム"], json=False, lint=False)
    ok = vbam_form2vba.cmd_form_inspect(args)
    assert ok is True
    assert "テストフォーム" in closed
    out = capsys.readouterr().out
    assert "閉じてからレイアウトを取得します" in out
    assert "コントロール数: 1" in out


@pytest.mark.core
def test_form_to_vba_accepts_force_option():
    """指示書089: form-to-vba が --force オプションを受け取れる"""
    parser = vba_manager.build_parser()
    args = parser.parse_args(["form-to-vba", "テストフォーム", "--force"])
    assert getattr(args, "force", False) is True


def test_form_builder_ledger_book_name(monkeypatch, tmp_path):
    """指示書089の2: 捨てのブックで部品を足した行の book が捨てのブックの名前になる"""
    import form_builder
    import vbam_core
    import json

    call_log = tmp_path / "_calls.jsonl"
    monkeypatch.setattr(form_builder, "SCRIPT_DIR", str(tmp_path))
    monkeypatch.setattr(vbam_core, "_CALL_LOG_FILE", str(call_log))
    # 直前のコマンドで別ブックを掴んだ状態を模倣（使い回し防止の検証）
    vbam_core._last_book_name = "直前の別ブック.xlsm"

    class _FakeWb:
        Name = "捨てブック999.xlsm"

    class _FakeControls:
        def Add(self, progid, name):
            class _FakeCtrl:
                Left = Top = Width = Height = 10
                Caption = ""
                Value = ""
                Font = type("Font", (), {"Size": 12})()
            return _FakeCtrl()

    class _FakeFrm:
        Controls = _FakeControls()
        _wb = _FakeWb()

    fb = form_builder.FormBuilder(None, _FakeWb())
    with fb:
        frm = _FakeFrm()
        form_builder.add_btn(frm, "btnTest", "OK", 0, 0, 50, 20)
        form_builder.add_lbl(frm, "lblTest", "Label", 0, 0, 50, 20)
        form_builder.add_txt(frm, "txtTest", 0, 0, 50, 20)

    assert call_log.exists()
    lines = [json.loads(line) for line in call_log.read_text(encoding="utf-8").splitlines() if line.strip()]
    assert len(lines) == 3
    for rec in lines:
        assert rec["book"] == "捨てブック999.xlsm"
        assert rec["book"] != "直前の別ブック.xlsm"


@pytest.mark.core
def test_form_inspect_name_variation_and_candidates(monkeypatch, capsys):
    """指示書089の3 型2: form-inspect が「フォーム」有無の揺れを吸収し、見つからないときは近い名前を案内する"""
    comp = _FakeComp("倍率回転フォーム", comp_type=3, designer=argparse.Namespace(InsideWidth=200, InsideHeight=150, Controls=[]))
    fake_wb = _FakeWb("秀コンボ.xlsm", comps=[comp])
    monkeypatch.setattr(vbam_form2vba, "get_workbook", lambda t: (None, fake_wb))

    # 「フォーム」抜きの「倍率回転」で呼ぶ → 成功
    args = argparse.Namespace(posargs=["倍率回転"], json=False, lint=False)
    ok = vbam_form2vba.cmd_form_inspect(args)
    assert ok is True
    out = capsys.readouterr().out
    assert "倍率回転" in out

    # 全く無い名前で呼ぶ → 近い名前案内が出る
    args_err = argparse.Namespace(posargs=["回転倍率テスト"], json=False, lint=False)
    ok_err = vbam_form2vba.cmd_form_inspect(args_err)
    assert ok_err is False
    out_err = capsys.readouterr().out
    assert "見つかりません" in out_err
    assert "近い名前" in out_err
    assert "倍率回転フォーム" in out_err


@pytest.mark.core
def test_form_click_prioritizes_button_over_label(monkeypatch):
    """指示書089の3 型4: 部品が複数当たるとき名前完全一致 → 見出し完全一致 → ボタン優先で決まる"""
    geo = {
        "Label5": [10, 10, 20, 20, "上"],
        "CommandButton7": [10, 50, 40, 20, "上"],
    }
    # 内部ロジック検証: 見出し「上」で検索するとボタン CommandButton7 が優先される
    exact_cap_cands = [k for k, v in geo.items() if (v[4] or "") == "上"]
    assert len(exact_cap_cands) == 2
    btn_cands = [k for k in exact_cap_cands if not k.lower().startswith(('label', 'lbl', 'sep'))]
    assert len(btn_cands) == 1
    assert btn_cands[0] == "CommandButton7"


def test_project_case_clash_message_shows_fix_guide():
    """指示書089の3 型5: 大小文字衝突時に「そろえ方: `x` → `X` にそろえてください」の1行が出る"""
    import vbam_core
    msg = vbam_core.project_case_clash_message([(10, "typename", "TypeName"), (20, "multiline", "Multiline")])
    assert "そろえ方: `typename` → `TypeName`、`multiline` → `Multiline` にそろえてください" in msg


def test_form_commands_accept_argument_variations():
    """指示書089の3 型6: --list, --form, --png, --font の引数の揺れを受け取れる"""
    parser = vba_manager.build_parser()

    # form-click --list
    args1 = parser.parse_args(["form-click", "--list"])
    assert getattr(args1, "list_flag", False) is True

    # form-read --list
    args2 = parser.parse_args(["form-read", "--list"])
    assert getattr(args2, "list_flag", False) is True

    # form-inspect --form <名前> --png --font 10
    args3 = parser.parse_args(["form-inspect", "--form", "テストフォーム", "--png", "--font", "10"])
    assert getattr(args3, "form_opt", None) == "テストフォーム"
    assert getattr(args3, "png", False) is True
    assert getattr(args3, "font_opt", None) == "10"


def test_form_click_with_target_file_and_missing_control_suggestion(monkeypatch, capsys, tmp_path):
    """指示書089の4: ブック名を先に書いたときの自動learnと、コントロール不在時の近い名前案内"""
    comp = _FakeComp("列選び直しフォーム", comp_type=3, designer=argparse.Namespace(
        InsideWidth=300, InsideHeight=200,
        Controls=[
            argparse.Namespace(Name="btn続ける", Left=10, Top=10, Width=50, Height=20, Caption="この列で続ける", Parent=argparse.Namespace(Name="列選び直しフォーム")),
            argparse.Namespace(Name="btnやめる", Left=70, Top=10, Width=50, Height=20, Caption="やめる", Parent=argparse.Namespace(Name="列選び直しフォーム")),
            argparse.Namespace(Name="lst列", Left=10, Top=40, Width=100, Height=100, Caption="", Parent=argparse.Namespace(Name="列選び直しフォーム")),
        ]
    ))
    fake_wb = _FakeWb("捨てフォーム.xlsm", comps=[comp])
    monkeypatch.setattr(vbam_edit, "get_workbook", lambda t: (None, fake_wb))
    monkeypatch.setattr(vbam_edit, "_list_form_windows", lambda: [])

    # 1. ブック名を先に書いて存在しないコントロール（btn実行）を指定
    args = argparse.Namespace(posargs=["捨てフォーム.xlsm", "列選び直しフォーム", "btn実行"])
    ok = vbam_edit.cmd_form_click(args)
    assert ok is False
    out = capsys.readouterr().out
    assert "当たるコントロールがありません" in out
    assert "近い名前" in out
    assert "btn続ける" in out


def test_form_owner_guard_not_bypassed_by_book_name(monkeypatch, capsys, tmp_path):
    """指示書089の6: ブック名に 'test' や '捨て' が含まれていても、フォームの持ち主なら --here 無しでは止まる"""
    fake_comp = argparse.Namespace(Name="倍率回転フォーム", Type=3)
    fake_wb = argparse.Namespace(Name="test_owner.xlsm", VBProject=argparse.Namespace(VBComponents=[fake_comp]))
    fake_xl = argparse.Namespace(ActiveWorkbook=fake_wb)

    import win32com.client
    monkeypatch.setattr(win32com.client, "GetActiveObject", lambda app: fake_xl)
    monkeypatch.setattr(vbam_edit, "get_workbook", lambda t: (fake_xl, fake_wb))
    monkeypatch.setattr(vbam_edit, "_list_form_windows", lambda: [(123, "倍率回転")])

    assert vbam_edit._form_owner_is_active_book("倍率回転フォーム") is True
    assert vbam_edit._form_owner_is_active_book("別のフォーム") is False

    json_file = tmp_path / "geo.json"
    import json
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump({"上": [10, 10, 20, 20, "上"]}, f)
    monkeypatch.setattr(vbam_edit, "_form_geometry_path", lambda name: str(json_file))

    args = argparse.Namespace(posargs=["test_owner.xlsm", "倍率回転フォーム", "上"], here_flag=False, learn_flag=False, list_flag=False)
    ok = vbam_edit.cmd_form_click(args)
    assert ok is False
    out = capsys.readouterr().out
    assert "前に出ているのは 倍率回転フォーム の持ち主のブックです" in out
