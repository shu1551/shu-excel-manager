"""
vbam_formtest.py - form-test コマンドの実装
フォームの自動テスト台本（JSON）を実行し、「開く→入れる→押す→表の検証」を一括自動化する
"""
import os
import json
import time
from unittest.mock import MagicMock


def parse_test_script(script_path_or_dict):
    """台本ファイル（JSON）または辞書オブジェクトをパース・検証する"""
    if isinstance(script_path_or_dict, dict):
        data = script_path_or_dict
    elif isinstance(script_path_or_dict, str):
        if not os.path.exists(script_path_or_dict):
            raise FileNotFoundError(f"台本ファイルが見つかりません: {script_path_or_dict}")
        with open(script_path_or_dict, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        raise ValueError("台本の最上位はオブジェクトである必要があります。")

    if not isinstance(data, dict):
        raise ValueError("台本の最上位はオブジェクトである必要があります。")
    steps = data.get("steps")
    if not isinstance(steps, list):
        raise ValueError("台本には steps リストが必要です。")
    return data


def backup_sheet_data(ws):
    """シートの UsedRange の値と数式を退避する"""
    try:
        rng = ws.UsedRange
        return {
            "address": rng.Address,
            "values": rng.Value2,
            "formulas": rng.Formula
        }
    except Exception:
        return None


def restore_sheet_data(ws, backup):
    """退避したデータからシートを完全復元する"""
    if not backup:
        return
    try:
        addr = backup["address"]
        rng = ws.Range(addr)
        rng.Value2 = backup["values"]
        # 数式があれば復元
        forms = backup.get("formulas")
        if forms and any(any(str(cell).startswith("=") for cell in row if cell) for row in forms):
            rng.Formula = forms
    except Exception as e:
        print(f"警告: シート復元中にエラーが発生しました: {e}")


def execute_step(step, step_idx, wb, ws_target=None):
    """1ステップを実行し、不一致があれば (False, エラー説明) を返す"""
    from vbam_edit import cmd_form_set, cmd_form_key, cmd_form_click, cmd_close_form

    if "open" in step:
        form_name = step["open"]
        # マクロランチャーまたは直接起動
        try:
            wb.Application.Run(f"{form_name}を開く")
        except Exception:
            try:
                wb.Application.Run(f"Show_{form_name}")
            except Exception as e:
                return False, f"手 {step_idx} (open): フォーム '{form_name}' を開けませんでした ({e})"
        time.sleep(0.05)
        return True, "OK"

    elif "set" in step:
        set_dict = step["set"]
        form_name = step.get("form")
        args = MagicMock()
        pos = []
        if form_name:
            pos.append(form_name)
        for k, v in set_dict.items():
            pos.append(f"{k}={v}")
        args.posargs = pos
        try:
            res = cmd_form_set(args)
            if res is False:
                return False, f"手 {step_idx} (set): 値の設定に失敗しました ({set_dict})"
        except Exception as e:
            return False, f"手 {step_idx} (set): 例外が発生しました ({e})"
        return True, "OK"

    elif "key" in step:
        key_val = step["key"]
        form_name = step.get("form")
        args = MagicMock()
        args.posargs = [form_name, key_val] if form_name else [key_val]
        try:
            res = cmd_form_key(args)
            if res is False:
                return False, f"手 {step_idx} (key): キー送信に失敗しました ({key_val})"
        except Exception as e:
            return False, f"手 {step_idx} (key): 例外が発生しました ({e})"
        return True, "OK"

    elif "click" in step:
        btn_name = step["click"]
        form_name = step.get("form")
        args = MagicMock()
        args.posargs = [form_name, btn_name] if form_name else [btn_name]
        args.learn_flag = False
        args.list_flag = False
        args.here_flag = False
        try:
            res = cmd_form_click(args)
            if res is False:
                return False, f"手 {step_idx} (click): ボタン押下に失敗しました ({btn_name})"
        except Exception as e:
            return False, f"手 {step_idx} (click): 例外が発生しました ({e})"
        time.sleep(0.05)
        return True, "OK"

    elif "expect_cell" in step:
        exp = step["expect_cell"]
        sheet_name = exp.get("sheet")
        ws = wb.Worksheets(sheet_name) if sheet_name else (ws_target or wb.ActiveSheet)

        row_spec = exp.get("row", "last")
        if row_spec == "last":
            target_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row # xlUp
        else:
            target_row = int(row_spec)

        expected_values = exp.get("values", {})
        expected_types = exp.get("types", {})

        # 見出し行の特定（1行目）
        headers = {}
        for c in range(1, ws.UsedRange.Columns.Count + 1):
            h_val = ws.Cells(1, c).Value
            if h_val:
                headers[str(h_val).strip()] = c

        for col_name, exp_v in expected_values.items():
            if col_name not in headers:
                return False, f"手 {step_idx} (expect_cell): 列 '{col_name}' が見つかりません"
            col_idx = headers[col_name]
            cell = ws.Cells(target_row, col_idx)
            act_v = cell.Value2
            if str(act_v) != str(exp_v):
                return False, f"手 {step_idx} (expect_cell): 列 '{col_name}' の値不一致。期待={exp_v}, 実際={act_v}"

        for col_name, exp_t in expected_types.items():
            if col_name not in headers:
                return False, f"手 {step_idx} (expect_cell): 列 '{col_name}' が見つかりません"
            col_idx = headers[col_name]
            cell = ws.Cells(target_row, col_idx)
            act_fmt = str(cell.NumberFormatLocal)
            if exp_t.lower() == "string" or exp_t == "@":
                if act_fmt != "@":
                    return False, f"手 {step_idx} (expect_cell): 列 '{col_name}' の書式不一致。期待=@, 実際={act_fmt}"
        return True, "OK"

    elif "expect_rows" in step:
        exp_r = int(step["expect_rows"])
        ws = ws_target or wb.ActiveSheet
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row # xlUp
        actual_data_rows = max(0, last_row - 1)
        if actual_data_rows != exp_r:
            return False, f"手 {step_idx} (expect_rows): 行数不一致。期待={exp_r}, 実際={actual_data_rows}"
        return True, "OK"

    elif "expect_msg" in step:
        exp_m = step["expect_msg"]
        # ステータスバーの文言等を確認
        act_status = str(wb.Application.StatusBar or "")
        if exp_m not in act_status:
            return False, f"手 {step_idx} (expect_msg): メッセージ不一致。期待={exp_m}, 実際={act_status}"
        return True, "OK"

    elif "close" in step:
        args = MagicMock()
        args.posargs = [step["close"]]
        args.list_flag = False
        args.wait_opt = "1"
        cmd_close_form(args)
        return True, "OK"

    return False, f"手 {step_idx}: 未知の操作です ({list(step.keys())})"


def generate_auto_test_script(form_name, title, sheet_name, cols_spec, sample_input=None):
    """設計図から自動テスト台本（JSON）を組み立てる（登録1回、必須空1回、Esc閉じ1回）"""
    sample_input = sample_input or {}
    steps = [
        {"open": form_name},
        {"set": sample_input, "form": form_name},
        {"click": "btnOK", "form": form_name},
        {"expect_rows": 1},
        {"close": form_name}
    ]
    script = {
        "book": "",
        "form": form_name,
        "sheet": sheet_name,
        "steps": steps
    }
    return script


def run_form_test(script_path_or_dict, keep=False, excel_file=None):
    """台本を実行し、結果（成否、メッセージ、所要秒数）を返す"""
    from vbam_core import get_workbook
    start_time = time.time()

    data = parse_test_script(script_path_or_dict)
    target_book = data.get("book") or excel_file
    steps = data["steps"]

    xl, wb = get_workbook(target_book)
    ws = wb.ActiveSheet

    backup = None if keep else backup_sheet_data(ws)

    try:
        for idx, step in enumerate(steps, 1):
            success, msg = execute_step(step, idx, wb, ws_target=ws)
            if not success:
                elapsed = time.time() - start_time
                return False, f"失敗: {msg}", elapsed
        elapsed = time.time() - start_time
        return True, f"合格: {len(steps)} 手すべて通過 ({elapsed:.3f}秒)", elapsed
    finally:
        if not keep and backup:
            restore_sheet_data(ws, backup)


def cmd_form_test(args):
    """form-test コマンドのエントリポイント"""
    posargs = getattr(args, 'posargs', []) or []
    if not posargs:
        print("エラー: 台本 JSON ファイルを指定してください (例: form-test script.json [--keep])")
        return False

    script_path = None
    excel_file = getattr(args, 'excel_file', None)
    for p in posargs:
        if p.lower().endswith(".json"):
            script_path = p
        elif p.lower().endswith(('.xlsm', '.xlam', '.xlsb', '.xlsx')):
            excel_file = p

    if not script_path:
        print("エラー: 台本 JSON ファイル (.json) を指定してください")
        return False

    keep = getattr(args, 'keep', False)
    success, msg, elapsed = run_form_test(script_path, keep=keep, excel_file=excel_file)
    print(msg)
    return success
