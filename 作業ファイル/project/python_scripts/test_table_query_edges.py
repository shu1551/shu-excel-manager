"""テーブル(ListObject)および PowerQuery 操作の徹底・エッジケース試験。

意地悪な実務データや極端な条件での罠（結合セル、0行テーブル、数式列、日本語列名、非同期更新等）を
徹底的に検証・担保するテストスイート。
"""
import os
import subprocess
import sys
import gc
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBM = os.path.join(SCRIPT_DIR, "vba_manager.py")


def _run_vbm(*args, timeout=60):
    """vba_manager.py を別プロセスで叩く（UTF-8入出力）"""
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
def temp_open_excel(tmp_path):
    """テスト用の Excel インスタンスと開いた一時ブックを提供するフィクスチャ。
    テスト終了後は必ず Excel を安全に終了してゾンビプロセスを残さない。
    """
    import win32com.client
    xl = None
    wb = None
    file_path = str(tmp_path / "edge_test.xlsx")
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "テストシート"
        wb.SaveAs(file_path)
        # 開いたまま yield する
        yield xl, wb, ws, file_path
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


def test_table_create_with_merged_cells(temp_open_excel):
    """【罠1: 結合セル爆弾】
    結合セルが含まれている範囲をテーブル化しようとした際、1004エラーで死なずに
    自動で結合解除し、データを保持したままテーブル化が完走することを検証。
    """
    xl, wb, ws, file_path = temp_open_excel
    # A1:B1 を結合（見出し）、A2:B2 も結合（データ）
    ws.Range("A1").Value = "見出し結合"
    ws.Range("A1:B1").Merge()
    ws.Range("C1").Value = "数量"
    ws.Range("A2").Value = "商品A"
    ws.Range("A2:B2").Merge()
    ws.Range("C2").Value = 10
    wb.Save()

    # table create 実行
    code, out = _run_vbm("table", file_path, "create", "A1:C2", "TestMergedTable")
    assert code == 0, f"テーブル作成が失敗しました: {out}"
    assert "結合セルを解除しました" in out or "テーブル作成" in out

    # 実機でテーブルが作成され、結合が解除されているか
    lo = ws.ListObjects("TestMergedTable")
    assert lo is not None
    assert lo.ListColumns.Count == 3
    # A1, A2 に値が残っているか
    assert ws.Range("A1").Value == "見出し結合"
    assert ws.Range("A2").Value == "商品A"


def test_table_column_formula_and_add_formula(temp_open_excel):
    """【罠2&3: 計算列の数式設定と自動展開】
    table column add --formula および table column formula で、
    構造化参照数式がテーブル全体に一括展開されることを検証。
    """
    xl, wb, ws, file_path = temp_open_excel
    ws.Range("A1").Value = "単価"
    ws.Range("B1").Value = "数量"
    ws.Range("A2").Value = 100
    ws.Range("B2").Value = 5
    ws.Range("A3").Value = 200
    ws.Range("B3").Value = 3
    lo = ws.ListObjects.Add(1, ws.Range("A1:B3"), None, 1)
    lo.Name = "SalesTable"
    wb.Save()

    # 1. table column add --formula で「金額」列を追加
    code, out = _run_vbm("table", file_path, "column", "add", "SalesTable", "金額", "--formula", "=[@単価]*[@数量]")
    assert code == 0, f"列追加(数式付き)に失敗: {out}"
    assert "金額" in out

    # 3列目が追加されているか
    assert lo.ListColumns.Count == 3
    # C2, C3 の値が計算されているか (100*5=500, 200*3=600)
    c2_val = ws.Range("C2").Value
    c3_val = ws.Range("C3").Value
    assert c2_val == 500, f"C2の計算値が不正: {c2_val}"
    assert c3_val == 600, f"C3の計算値が不正: {c3_val}"

    # 2. table column formula で「金額」列の数式を書き換え (10%税込)
    code, out = _run_vbm("table", file_path, "column", "formula", "SalesTable", "金額", "=[@単価]*[@数量]*1.1")
    assert code == 0, f"数式書き換えに失敗: {out}"
    # 再計算値の確認 (500*1.1=550, 600*1.1=660)
    assert ws.Range("C2").Value == 550
    assert ws.Range("C3").Value == 660


def test_table_zero_rows_safety(temp_open_excel):
    """【罠4: 0行テーブル（ヘッダーのみ）の安全性】
    DataBodyRange が None の空テーブルに対しても、列追加や書式設定がクラッシュせず安全に完了することを検証。
    """
    xl, wb, ws, file_path = temp_open_excel
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "名称"
    lo = ws.ListObjects.Add(1, ws.Range("A1:B1"), None, 1)
    lo.Name = "EmptyTable"
    wb.Save()

    # 0行テーブルに列追加
    code, out = _run_vbm("table", file_path, "column", "add", "EmptyTable", "備考")
    assert code == 0, f"0行テーブルの列追加でクラッシュ: {out}"
    assert lo.ListColumns.Count == 3

    # 0行テーブルの列書式設定
    code, out = _run_vbm("table", file_path, "column", "format", "EmptyTable", "ID", "@")
    assert code == 0, f"0行テーブルの書式設定でクラッシュ: {out}"


def test_table_special_char_escape():
    """【罠5: 構造化参照のエスケープロジック】
    日本語カッコ、記号、スペースを含む列名が正しく構造化参照エスケープされるかを検証。
    """
    from vbam_edit import _escape_table_col_ref
    assert _escape_table_col_ref("単価") == "[@単価]"
    assert _escape_table_col_ref("単価 (税込)") == "[@[単価 (税込)]]"
    assert _escape_table_col_ref("数量[個]") == "[@[数量'[個']]]"
    assert _escape_table_col_ref("#ID") == "[@[#ID]]"


def test_table_special_char_col_formula(temp_open_excel):
    """【罠6: 記号入り列名での実機計算列テスト】
    '単価 (税込)' と '数量' という記号を含む列名に対し、構造化参照
    =[@[単価 (税込)]]*[@数量] で計算列を追加し、正しく実機計算されることを検証。
    """
    xl, wb, ws, file_path = temp_open_excel
    ws.Range("A1").Value = "単価 (税込)"
    ws.Range("B1").Value = "数量"
    ws.Range("A2").Value = 110
    ws.Range("B2").Value = 2
    lo = ws.ListObjects.Add(1, ws.Range("A1:B2"), None, 1)
    lo.Name = "SpecialColTable"
    wb.Save()

    # 計算列を追加
    code, out = _run_vbm("table", file_path, "column", "add", "SpecialColTable", "支払額", "--formula", "=[@[単価 (税込)]]*[@数量]")
    assert code == 0, f"記号入り列名の計算列追加に失敗: {out}"
    assert ws.Range("C2").Value == 220


def test_table_delete_unlist_keeps_data(temp_open_excel):
    """【罠7: テーブル解除時のデータ保持】
    table delete で Unlist された後、データがセルに残っていることを検証。
    """
    xl, wb, ws, file_path = temp_open_excel
    ws.Range("A1").Value = "A"
    ws.Range("B1").Value = "B"
    ws.Range("A2").Value = 123
    ws.Range("B2").Value = 456
    lo = ws.ListObjects.Add(1, ws.Range("A1:B2"), None, 1)
    lo.Name = "UnlistTable"
    wb.Save()

    # 解除
    code, out = _run_vbm("table", file_path, "delete", "UnlistTable")
    assert code == 0, f"テーブル解除に失敗: {out}"
    assert ws.ListObjects.Count == 0
    assert ws.Range("A1").Value == "A"
    assert ws.Range("A2").Value == 123
    assert ws.Range("B2").Value == 456


def test_table_column_rename_and_remove(temp_open_excel):
    """【列名変更と列削除の整合性】"""
    xl, wb, ws, file_path = temp_open_excel
    ws.Range("A1").Value = "旧名"
    ws.Range("B1").Value = "不要列"
    ws.Range("A2").Value = 10
    ws.Range("B2").Value = 20
    lo = ws.ListObjects.Add(1, ws.Range("A1:B2"), None, 1)
    lo.Name = "ColModTable"
    wb.Save()

    # 列名変更
    code, out = _run_vbm("table", file_path, "column", "rename", "ColModTable", "旧名", "新名")
    assert code == 0, f"列名変更に失敗: {out}"
    assert lo.ListColumns(1).Name == "新名"

    # 列削除
    code, out = _run_vbm("table", file_path, "column", "remove", "ColModTable", "不要列")
    assert code == 0, f"列削除に失敗: {out}"
    assert lo.ListColumns.Count == 1


@pytest.mark.core
def test_powerquery_refresh_no_queries_clean(temp_open_excel):
    """クエリや接続がないブックでも powerquery refresh がクラッシュせず安全に終了すること"""
    xl, wb, ws, file_path = temp_open_excel
    code, out = _run_vbm("powerquery", file_path, "refresh")
    assert code == 0, f"powerquery refresh に失敗: {out}"
