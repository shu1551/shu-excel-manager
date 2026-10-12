"""UserForm レイアウト逆解析 (form-inspect) の実機エッジケース試験。

実機 Excel 上にボタンとテキストボックスを配置した UserForm を生成し、
form-inspect コマンドでコントロールの座標・サイズ・キャプション・JSON構造が
完璧に吸い出せるかを検証。
"""
import os
import subprocess
import sys
import json
import gc
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBM = os.path.join(SCRIPT_DIR, "vba_manager.py")


def _run_vbm(*args, timeout=60):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run(
        [sys.executable, VBM, *args],
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        cwd=SCRIPT_DIR,
        env=env,
        timeout=timeout,
    )
    return p.returncode, (p.stdout or "") + (p.stderr or "")


@pytest.fixture
def temp_form_excel(tmp_path):
    """実機 Excel 上にテスト用 UserForm を配置したブックを提供するフィクスチャ"""
    import win32com.client
    xl = None
    wb = None
    file_path = str(tmp_path / "form_test.xlsm")
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        
        # UserForm を作成 (Type=3: vbext_ct_MSForm)
        comp = wb.VBProject.VBComponents.Add(3)
        comp.Name = "TestInputForm"
        comp.Properties("Caption").Value = "テスト用フォーム"
        comp.Properties("Width").Value = 300
        comp.Properties("Height").Value = 200
        
        # コントロールを配置 (Label, TextBox, CommandButton)
        d = comp.Designer
        lbl = d.Controls.Add("Forms.Label.1", "lblTitle")
        lbl.Caption = "お名前:"
        lbl.Left = 10
        lbl.Top = 15
        lbl.Width = 60
        lbl.Height = 18

        txt = d.Controls.Add("Forms.TextBox.1", "txtName")
        txt.Left = 80
        txt.Top = 12
        txt.Width = 150
        txt.Height = 22

        btn = d.Controls.Add("Forms.CommandButton.1", "btnSubmit")
        btn.Caption = "送信"
        btn.Left = 80
        btn.Top = 50
        btn.Width = 80
        btn.Height = 26

        wb.SaveAs(file_path, FileFormat=52)
        yield xl, wb, file_path
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        if xl:
            try:
                xl.Quit()
            except Exception:
                pass
            xl = None
        gc.collect()


def test_form_inspect_json_output(temp_form_excel):
    """【実機試験: 既存UserFormのレイアウトJSON吸い出し】
    配置された全コントロール（名前、座標、幅、高さ、キャプション）が
    JSONで完璧に取得できることを実機検証。
    """
    xl, wb, file_path = temp_form_excel

    # form-inspect --json 実行
    code, out = _run_vbm("form-inspect", file_path, "TestInputForm", "--json")
    assert code == 0, f"form-inspect の実行に失敗: {out}"

    # JSONパース（先頭の「対象ブック: ...」をスキップして { から読む）
    json_str = out[out.find("{"):]
    data = json.loads(json_str)
    assert data["form"] == "TestInputForm"
    assert data["info"]["caption"] == "テスト用フォーム"

    ctrls = {c["name"]: c for c in data["controls"]}
    assert "lblTitle" in ctrls
    assert "txtName" in ctrls
    assert "btnSubmit" in ctrls

    # 各コントロールのプロパティ検証
    assert ctrls["lblTitle"]["caption"] == "お名前:"
    assert ctrls["lblTitle"]["left"] == 10
    assert ctrls["txtName"]["left"] == 80
    assert ctrls["btnSubmit"]["caption"] == "送信"
    assert ctrls["btnSubmit"]["width"] == 80
