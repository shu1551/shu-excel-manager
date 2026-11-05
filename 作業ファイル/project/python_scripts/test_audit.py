"""インサイト・データ診断エンジン (vbam_audit) およびスタイルマップのユニットテスト
"""

import os
import re
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vba_manager as vm
import vbam_audit as va
import vbam_view as vv
import vbam_vba as vba


def test_address_conversions():
    assert va._a1_to_rowcol("A1") == (0, 0)
    assert va._a1_to_rowcol("$B$5") == (4, 1)
    assert va._a1_to_rowcol("Z10") == (9, 25)
    assert va._a1_to_rowcol("AA1") == (0, 26)
    assert va._a1_to_rowcol("AB2") == (1, 27)

    assert va._rowcol_to_a1(0, 0) == "A1"
    assert va._rowcol_to_a1(4, 1) == "B5"
    assert va._rowcol_to_a1(9, 25) == "Z10"
    assert va._rowcol_to_a1(0, 26) == "AA1"
    assert va._rowcol_to_a1(1, 27) == "AB2"


@pytest.mark.core
def test_formula_linter_detects_errors():
    # 数式エラー (#DIV/0!)
    formulas = [["=1/0"]]
    r1c1 = [["=1/0"]]
    values = [["#DIV/0!"]]
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 0)
    assert len(issues) == 1
    assert issues[0]["type"] == "formula_error"
    assert issues[0]["severity"] == "critical"
    assert issues[0]["cell"] == "A1"


@pytest.mark.core
def test_formula_linter_detects_hardcoded_constants():
    # =A1*1.1 (消費税等)
    formulas = [["=A1*1.1"], ["=A2+500"], ["=A3*1"]]
    r1c1 = [["=RC[-1]*1.1"], ["=RC[-1]+500"], ["=RC[-1]*1"]]
    values = [[110], [600], [100]]
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 0)
    types = [i["type"] for i in issues]
    assert types.count("hardcoded_constant") == 2 # 1.1 と 500 を検知し、*1 は除外
    assert issues[0]["cell"] == "A1"
    assert "1.1" in issues[0]["msg"]
    assert issues[1]["cell"] == "A2"
    assert "500" in issues[1]["msg"]


def test_formula_linter_detects_omitted_sum_range():
    # B2:B4 の集計なのに、B5 に数値が入っている（B6 が SUM）
    # B1: 見出し
    # B2: 10
    # B3: 20
    # B4: 30
    # B5: 40 (直下セル)
    # B6: =SUM(B2:B4)
    formulas = [
        ["売上"],
        [10],
        [20],
        [30],
        [40],
        ["=SUM(B2:B4)"]
    ]
    r1c1 = [
        ["売上"],
        [10],
        [20],
        [30],
        [40],
        ["=SUM(R[-4]C:R[-2]C)"]
    ]
    values = [
        ["売上"],
        [10],
        [20],
        [30],
        [40],
        [60]
    ]
    # start_row=0, start_col=1 (B列)
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 1)
    omitted = [i for i in issues if i["type"] == "omitted_sum_range"]
    assert len(omitted) == 1
    assert omitted[0]["cell"] == "B6"
    assert "B5 (40) が SUM から漏れています" in omitted[0]["msg"]


def test_formula_linter_detects_horizontal_omitted_range():
    # A1: 10, B1: 20, C1: 30, D1: 40, E1: =AVERAGE(A1:C1) (D1が漏れている)
    formulas = [["10", "20", "30", "40", "=AVERAGE(A1:C1)"]]
    r1c1 = [["10", "20", "30", "40", "=AVERAGE(RC[-4]:RC[-2])"]]
    values = [[10, 20, 30, 40, 20]]
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 0)
    omitted = [i for i in issues if i["type"] == "omitted_sum_range"]
    assert len(omitted) == 1
    assert omitted[0]["cell"] == "E1"
    assert "D1 (40) が AVERAGE から漏れています" in omitted[0]["msg"]


def test_formula_linter_detects_inconsistent_formulas():
    # 5行のうち、4行が =RC[-2]*RC[-1] なのに、1行だけ手入力値や別の式
    formulas = [
        ["=A1*B1"],
        ["=A2*B2"],
        ["=A3*B3"],
        ["=A4+B4"],  # ここだけ異なる式
        ["=A5*B5"],
    ]
    r1c1 = [
        ["=RC[-2]*RC[-1]"],
        ["=RC[-2]*RC[-1]"],
        ["=RC[-2]*RC[-1]"],
        ["=RC[-2]+RC[-1]"],
        ["=RC[-2]*RC[-1]"],
    ]
    values = [[10], [20], [30], [25], [50]]
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 2)
    incons = [i for i in issues if i["type"] == "inconsistent_formula"]
    assert len(incons) == 1
    assert incons[0]["cell"] == "C4"


def test_data_cleaner_detects_hidden_whitespace_and_text_numbers():
    values = [
        [" 東京都 "],    # 前後空白
        ["大阪府\n"],     # 改行
        ["'12345"],      # 文字列数字
        ["999"],         # 文字列数字 (str型)
        [1000],          # 数値型
    ]
    issues = va.check_data_cleaner(values, 0, 0)
    
    ws_issues = [i for i in issues if i["type"] == "hidden_whitespace"]
    assert len(ws_issues) == 2
    assert ws_issues[0]["cell"] == "A1"
    assert "先頭空白・末尾空白" in ws_issues[0]["msg"]
    assert ws_issues[1]["cell"] == "A2"
    assert "改行" in ws_issues[1]["msg"]

    num_issues = [i for i in issues if i["type"] == "text_number"]
    assert any(i["cell"] == "A4" and "999" in i["msg"] for i in num_issues)


def test_data_cleaner_detects_outliers_iqr():
    # ほとんどが 100〜200 の中に、1つだけ桁間違い（100000）
    vals = [[100], [110], [120], [130], [140], [150], [100000]]
    issues = va.check_data_cleaner(vals, 0, 0)
    outliers = [i for i in issues if i["type"] == "outlier_value"]
    assert len(outliers) == 1
    assert outliers[0]["cell"] == "A7"
    assert "10 倍以上離れています" in outliers[0]["msg"]


def test_outliers_small_column_not_noisy_20261008():
    """件数の少ない列で、単価 45〜150 の 380・個数 0〜40 の 200 を「著しく乖離」と出していた（2026-10-08 初見の受注の表）。
    桁違い（1 桁以上離れた値）だけを出す。"""
    for col in ([120, 120, 45, 380, 150, 150], [12, 30, 200, 5, 0, 40]):
        vals = [["列"]] + [[v] for v in col]
        assert [i for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "outlier_value"] == []
    vals = [["単価"]] + [[v] for v in (120, 120, 45, 3800, 150, 150)]
    assert [i["cell"] for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "outlier_value"] == ["A5"]


def test_data_cleaner_outliers_skip_total_rows():
    """2026-09-19: 月別の表の「合計」の行が、4 列とも外れ値と出た。合計・小計の行は検査に入れない。
    明細の中の桁違いは、合計の行があっても拾う。"""
    vals = [["月", "A", "B"], ["4月", 1850, 620], ["5月", 2120, 580], ["6月", 1980, 710],
            ["7月", 2460, 845], ["8月", 2230, 790], ["9月", 2710, 930], ["合計", 13350, 4475]]
    assert [i for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "outlier_value"] == []
    vals[3][1] = 1980000                                     # 明細の桁違い
    hit = [i["cell"] for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "outlier_value"]
    assert hit == ["B4"]
    vals2 = [["課", "額"], ["総務", 100], ["財政", 120], ["小 計", 220], ["税務", 110], ["福祉", 130], ["課別合計", 460]]
    assert [i for i in va.check_data_cleaner(vals2, 0, 0) if i["type"] == "outlier_value"] == []


@pytest.mark.core
def test_audit_report_formatting():
    issues = [
        {"cell": "A1", "severity": "critical", "type": "formula_error", "msg": "エラー", "formula": "=1/0"},
        {"cell": "B2", "severity": "warning", "type": "hardcoded_constant", "msg": "定数直書き", "formula": "=C2*1.1"},
        {"cell": "C3", "severity": "info", "type": "hidden_whitespace", "msg": "末尾空白", "value": "x "}
    ]
    report = va.format_audit_report("Sheet1", "A1:C10", issues)
    assert "総合監査レポート" in report
    assert "Sheet1!A1:C10" in report
    # ペナルティ: 15 + 5 + 2 = 22点減点 → 78点 (判定 B)
    assert "78点 / 100点" in report
    assert "【要修正 (Critical)】" in report
    assert "【要確認 (Warning)】" in report
    assert "【クレンジング (Info)】" in report


@pytest.mark.core
def test_diagnose_and_style_map_parser_and_wiring():
    p = vm.build_parser()
    
    # diagnose パース
    ns_diag = p.parse_args(["diagnose", "秀コンボ.xlsm", "Sheet1!A1:D10", "--json"])
    assert ns_diag.posargs == ["秀コンボ.xlsm", "Sheet1!A1:D10"]
    assert ns_diag.json is True

    # style-map パース
    ns_sm = p.parse_args(["style-map", "秀コンボ.xlsm", "--sheet", "データ"])
    assert ns_sm.posargs == ["秀コンボ.xlsm"]
    assert ns_sm.sheet_opt == "データ"

    # コマンドテーブル配線
    table = vm._command_table()
    assert "diagnose" in table
    assert "style-map" in table
    assert vm.raw_command(table["diagnose"]) is va.cmd_diagnose
    assert vm.raw_command(table["style-map"]) is vv.cmd_style_map

    # 権限テーブル配線
    read_cmds = {name for name, _ in vba._CAPABILITIES["read"]}
    assert "diagnose" in read_cmds
    assert "style-map" in read_cmds


def test_circular_reference_detection():
    # 1. 自己参照: A1 = "=A1+1"
    formulas_self = [["=A1+1"]]
    issues = va.check_circular_references(formulas_self, 0, 0)
    assert len(issues) == 1
    assert issues[0]["type"] == "circular_reference"
    assert "A1 -> A1" in issues[0]["msg"]

    # 2. 間接循環: A1 = "=B1*2", B1 = "=A1+5"
    formulas_indirect = [["=B1*2", "=A1+5"]]
    issues = va.check_circular_references(formulas_indirect, 0, 0)
    assert len(issues) >= 1
    assert any(i["type"] == "circular_reference" for i in issues)
    msg = issues[0]["msg"]
    assert "A1" in msg and "B1" in msg

    # 3. 範囲参照を含む循環: A3 = "=SUM(A1:A3)"
    formulas_range = [
        [10],
        [20],
        ["=SUM(A1:A3)"]
    ]
    issues = va.check_circular_references(formulas_range, 0, 0)
    assert len(issues) == 1
    assert issues[0]["cell"] == "A3"
    assert issues[0]["type"] == "circular_reference"

    # 4. 正常な依存関係（非循環）: A1=10, A2=20, A3="=A1+A2"
    formulas_normal = [
        [10],
        [20],
        ["=A1+A2"]
    ]
    issues_norm = va.check_circular_references(formulas_normal, 0, 0)
    assert len(issues_norm) == 0


def test_apply_audit_fixes_logic():
    class MockRange:
        def __init__(self, addr):
            self.Address = addr
            self.Formula = ""
            self.Value = None

    class MockWorksheet:
        def __init__(self):
            self.cells = {}

        def Range(self, addr):
            if addr not in self.cells:
                self.cells[addr] = MockRange(addr)
            return self.cells[addr]

    ws = MockWorksheet()
    ws.Range("B6").Formula = "=SUM(B2:B4)"
    ws.Range("C2").Value = "  東京都渋谷区  "
    ws.Range("D3").Value = "12345"

    issues = [
        {
            "cell": "B6",
            "type": "omitted_sum_range",
            "formula": "=SUM(B2:B4)",
            "fixed_formula": "=SUM(B2:B5)"
        },
        {
            "cell": "C2",
            "type": "hidden_whitespace",
            "value": "  東京都渋谷区  "
        },
        {
            "cell": "D3",
            "type": "text_number",
            "value": "12345"
        }
    ]

    ws.Range("E4").Value = "0123"          # 先頭ゼロ＝コード。数値にしない
    ws.Range("F5").Value = " 12 "
    ws.Range("F5").HasFormula = True        # 式の結果が文字＝式を直すのが筋。触らない
    issues += [
        {"cell": "E4", "type": "text_number", "value": "0123"},
        {"cell": "F5", "type": "hidden_whitespace", "value": " 12 "},
    ]

    res = va.apply_audit_fixes(ws, issues)
    counts = res["fixed_counts"]
    # 集計漏れの式は書かない（提案だけ）・先頭ゼロと数式のセルは触らない（2026-09-17）
    assert counts["total"] == 2
    assert counts["omitted_sum_range"] == 0
    assert counts["hidden_whitespace"] == 1
    assert counts["text_number"] == 1
    assert ws.Range("B6").Formula == "=SUM(B2:B4)"
    assert any("B6" in s_ and "=SUM(B2:B5)" in s_ for s_ in res["suggestions"])
    assert ws.Range("C2").Value == "東京都渋谷区"
    assert ws.Range("D3").Value == 12345
    assert ws.Range("E4").Value == "0123"
    assert ws.Range("F5").Value == " 12 "
    assert len(res["skipped"]) == 2

    # 見出しが番号の列（社員番号など）は、0 で始まらない番号も文字のまま（照合のキーの型をそろえる・2026-10-28 点検）
    ws.Range("G2").Value = "1001"
    res2 = va.apply_audit_fixes(ws, [{"cell": "G2", "type": "text_number", "value": "1001"}], {"G": "社員番号"})
    assert ws.Range("G2").Value == "1001" and res2["fixed_counts"]["text_number"] == 0
    assert any("社員番号" in s_ for s_ in res2["skipped"])
    # 見出しの読み取り: 上から 5 行のうち最初の文字。番号の語（番号・ID・コード・No）がある列だけ
    vals = [["社員番号", "氏名", "金額", "商品ID", "キーボード"], ["1001", "佐藤", "1200", "A1", "3"]]
    assert va.id_columns_of(vals, 0) == {"A": "社員番号", "D": "商品ID"}
    assert va.id_columns_of(vals, 2) == {"C": "社員番号", "F": "商品ID"}     # 範囲が C 列から始まるとき


def test_text_number_safe_to_convert():
    assert va._text_number_safe_to_convert("12345")
    assert va._text_number_safe_to_convert("0")
    assert va._text_number_safe_to_convert("0.5")
    assert va._text_number_safe_to_convert("-7")
    assert not va._text_number_safe_to_convert("0123")       # 先頭ゼロ＝コード
    assert not va._text_number_safe_to_convert("0001234567")
    assert not va._text_number_safe_to_convert("1234567890123456")  # 16 桁


@pytest.mark.core
def test_diagnose_fix_requires_yes():
    p = vm.build_parser()
    ns = p.parse_args(["diagnose", "--fix"])
    assert ns.fix is True and ns.yes is False
    ns2 = p.parse_args(["diagnose", "--fix", "-y"])
    assert ns2.yes is True


def test_inconsistent_formula_is_a_notice_not_a_block():
    formulas = [["=A1*B1"], ["=A2*B2"], ["=A3*B3"], ["=A4+B4"], ["=A5*B5"]]
    r1c1 = [["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"], ["=RC[-2]+RC[-1]"], ["=RC[-2]*RC[-1]"]]
    issues = va.check_formula_linter(formulas, r1c1, [[1], [2], [3], [4], [5]], 0, 2)
    incons = [i for i in issues if i["type"] == "inconsistent_formula"]
    assert incons and all(i["severity"] == "warning" for i in incons)


def test_generate_audit_html():
    issues = [
        {"cell": "A1", "severity": "critical", "type": "circular_reference", "msg": "循環参照 A1 -> B1 -> A1", "formula": "=B1"},
        {"cell": "B2", "severity": "warning", "type": "hardcoded_constant", "msg": "定数直書き *1.1", "formula": "=C2*1.1"},
        {"cell": "C3", "severity": "info", "type": "hidden_whitespace", "msg": "末尾空白", "value": "x "}
    ]
    html = va.generate_audit_html("Sheet1", "A1:C10", issues)
    assert "<!DOCTYPE html>" in html
    assert "Excel 総合健全性診断書" in html
    assert "Sheet1!A1:C10" in html
    assert "循環参照" in html
    assert "gauge-svg" in html
    assert "78" in html  # 100 - (15 + 5 + 2) = 78点


def test_clean_vba_folded_into_check():
    """clean-vba（9/16 夜・Gemini）は check に畳んだ（2026-09-17）: コマンドは無く、目は VBM015・016 として残る。
    Option Explicit は数えない（原則 7）。戻し忘れは VBM008 が見る。"""
    p = vm.build_parser()
    with pytest.raises(SystemExit):
        p.parse_args(["clean-vba"])
    assert "clean-vba" not in vm._command_table()
    import vbam_devtools as vd
    assert not hasattr(vd, 'cmd_clean_vba')
    d = {r[0]: r for r in vba._CHECK_RULES}
    assert 'Resume Next' in d['VBM015'][2] and '150' in d['VBM016'][2]
    lines = ["Sub A()", "    On Error Resume Next", "    x = 1", "End Sub"] + ["Sub B()"] + ["    y = 1"] * 160 + ["End Sub"]
    procs = vba._split_procedures(lines)
    assert vba._diag_resume_next(lines) == [(2, "On Error Resume Next")]
    assert vba._diag_long_procs(procs) == [("B", 162)]
    assert vba._diag_long_procs(procs, limit=200) == []


def test_diagnose_notes_only_adds_what_other_eyes_lack():
    """materials／seiri に畳んだ診断の目は 循環参照（要修正）と 外れ値（気づき）だけ。
    集計漏れ・直書き・空白・文字の数字は数式の目と表の汚れが既に出すので重ねない。"""
    fa = [["=B1*2", "=A1+5"], ["=SUM(A1:A1)", 10]]
    fr = [["=RC[1]*2", "=RC[-1]+5"], ["=SUM(R[-1]C:R[-1]C)", 10]]
    vals = [[1, 2], [3, 10]]
    blk, ntc = vv.diagnose_notes(fa, fr, vals, 1, 1)
    assert blk and all("循環" in b or "->" in b for b in blk)
    assert ntc == []
    vals2 = [[100], [110], [120], [130], [140], [150], [100000]]
    fa2 = [[v] for v in [100, 110, 120, 130, 140, 150, 100000]]
    blk2, ntc2 = vv.diagnose_notes(fa2, fa2, vals2, 1, 1)
    assert blk2 == [] and len(ntc2) == 1 and ntc2[0].startswith("A7:")


def test_style_summary_reads_range_once_and_names_merged_headers():
    """書式の目の要約: 結合の親子・二重下線の合計行・塗りの有無を偽の COM で確かめる（純 Python）。"""
    class Cell:
        def __init__(self, text, double=False):
            self.Text = text
            self._d = double
        def Borders(self, i):
            class B: pass
            b = B(); b.LineStyle = -4119 if self._d else 1
            return b
    class _Font: Color = None; Bold = False
    class _Interior: ColorIndex = -4142
    class UR:
        Row = 1; Column = 1
        def Cells(self, i, j):
            return Cell("", double=(i == 7))
    ur = UR(); ur.Interior = _Interior(); ur.Font = _Font()
    class WS:
        def Cells(self, r, c):
            return Cell({(1, 1): "2026年度", (2, 1): "東日本", (2, 4): "西日本", (3, 1): "Q1", (3, 2): "Q2", (3, 3): "上期"}.get((r, c), ""))
    out = vv.style_summary(WS(), ur, 7, 6, ["A1:F1", "A2:C2", "D2:F2"])
    joined = "\n".join(out)
    assert "A1:F1「2026年度」" in joined and "東日本" in joined and "西日本" in joined
    assert "A2:C2「東日本」 → 下の行: Q1 / Q2 / 上期" in joined
    assert "合計行（下罫線が二重）: 行 7" in joined
    assert "文字色の混在あり" in joined and "塗り" not in joined


def test_formula_linter_subtotal_blocks_are_not_omitted():
    # 小計ごとの SUM: B4=SUM(B2:B3)（小計）・B7=SUM(B5:B6)（小計）＝B7 の直上 B4 は前の小計で、漏れではない（2026-09-17 夜）
    formulas = [["科目", "金額"], ["a", "10"], ["b", "20"], ["小計", "=SUM(B2:B3)"], ["c", "30"], ["d", "40"], ["小計", "=SUM(B5:B6)"]]
    r1c1 = [["科目", "金額"], ["a", "10"], ["b", "20"], ["小計", "=SUM(R[-2]C:R[-1]C)"], ["c", "30"], ["d", "40"], ["小計", "=SUM(R[-2]C:R[-1]C)"]]
    values = [["科目", "金額"], ["a", 10], ["b", 20], ["小計", 30], ["c", 30], ["d", 40], ["小計", 70]]
    issues = va.check_formula_linter(formulas, r1c1, values, 1, 1)
    assert not [i for i in issues if i["type"] == "omitted_sum_range"]


def test_formula_linter_does_not_flag_a_running_total_column():
    """累計の列（=SUM($B$2:B4) を C4 に置く形）は、直下を外すのが仕事＝集計漏れではない（2026-09-18）。

    B 集計の「累計の列」を鍛えたとき、値は正解と同じなのに毎回ここで止まり AI に回った。
    """
    formulas = [
        ["金額", "累計"],
        [10, "=SUM($A$2:A2)"],
        [20, "=SUM($A$2:A3)"],
        [30, "=SUM($A$2:A4)"],
        [40, "=SUM($A$2:A5)"],
    ]
    r1c1 = [["金額", "累計"]] + [[0, "=SUM(R2C1:RC[-1])"]] * 4
    values = [["金額", "累計"], [10, 10], [20, 30], [30, 60], [40, 100]]
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 0)
    assert [i for i in issues if i["type"] == "omitted_sum_range"] == []
    # 範囲の終わりが自分の行より上なら、今までどおり漏れとして咎める（A4 が漏れている）
    formulas[4][1] = "=SUM($A$2:A3)"
    issues = va.check_formula_linter(formulas, r1c1, values, 0, 0)
    assert [i["cell"] for i in issues if i["type"] == "omitted_sum_range"] == ["B5"]


def test_value_in_formula_column_is_flagged():
    """式の列（=E2*F2 …）に紛れた数字の直書きを出す（2026-09-23 の通しの実測で G7=5000 を誰も言わなかった）。
    合計の行・文字（備考）・式の行の外は咎めない。"""
    fa = [["金額"], ["=E2*F2"], ["=E3*F3"], [5000], ["=E5*F5"], ["=E6*F6"], ["=SUM(G2:G6)"], ["メモ"]]
    fr = [["金額"], ["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"], [5000], ["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"],
          ["=SUM(R[-5]C:R[-1]C)"], ["メモ"]]
    vals = [["金額"], [4800], [6000], [5000], [7500], [3000], [26300], ["メモ"]]
    hit = va.check_values_in_formula_columns(fa, fr, vals, 0, 6)
    assert [i["cell"] for i in hit] == ["G4"]
    assert any(i["type"] == "value_in_formula_column" for i in va.check_formula_linter(fa, fr, vals, 0, 6))
    blk, ntc = vv.diagnose_notes(fa, fr, vals, 1, 7)
    assert any(n.startswith("式の列 G4") for n in ntc)
    # 式が 2 行しか無い列・文字の書き込みは咎めない
    fa2 = [["=A1"], ["=A2"], ["手入力"]]
    assert va.check_values_in_formula_columns(fa2, fa2, [[1], [2], ["手入力"]], 0, 0) == []


def test_outliers_need_a_lonely_value():
    """囲いの外でも、残りの端から 5 倍以上離れた「ぽつんと 1 つ」だけを外れ値とする（2026-09-23 の通しの実測で
    単価 30〜12,000 のふつうの表に 4 件出ていた）。桁違い・マイナスは今までどおり拾う。"""
    price = [[480], [120], [300], [1500], [250], [8500], [1200], [480], [30], [90], [700], [12000], [150]]
    qty = [[10], [50], [20], [5], [12], [2], [30], [10], [100], [24], [3], [1], [12]]
    for col in (price, qty):
        assert [i for i in va.check_data_cleaner(col, 0, 0) if i["type"] == "outlier_value"] == []
    typo = [row[:] for row in price]
    typo[4][0] = 2500000                                         # 250 → 2,500,000（桁違い）
    assert [i["cell"] for i in va.check_data_cleaner(typo, 0, 0) if i["type"] == "outlier_value"] == ["A5"]
    minus = [[1200], [1300], [1250], [1400], [-1300], [1350], [1280]]
    assert [i["cell"] for i in va.check_data_cleaner(minus, 0, 0) if i["type"] == "outlier_value"] == ["A5"]


def test_measure2_noise_and_sign_rules():
    """2026-09-23 通しの実測 2 で出た気づきの直し: 先頭 0 の番号は文字の数字と咎めない／正の列のマイナス 1 つは出す／
    式のセルの外れ値は重ねて出さない（materials・seiri の目）。"""
    vals = [["伝票番号", "数量", "単価", "金額"],
            ["0001", 10, 1200, 12000], ["0002", 5, 3000, 15000], ["0003", 8, 1500, 12000],
            ["0004", 3, 2000, 6000], ["0005", -5, 1200, -6000], ["0006", 2, 120000, 240000],
            ["0007", 4, 3000, 12000], ["0008", 7, 1200, 8400], ["0009", 9, 1500, 13500], ["0010", 1, 5000, 5000]]
    issues = va.check_data_cleaner(vals, 0, 0)
    assert not [i for i in issues if i["type"] == "text_number"]
    out = {i["cell"] for i in issues if i["type"] == "outlier_value"}
    assert "B6" in out and "C7" in out
    fa = [row[:3] + [("=B%d*C%d" % (k + 1, k + 1)) if k else row[3]] for k, row in enumerate(vals)]
    blk, ntc = vv.diagnose_notes(fa, fa, vals, 1, 1)
    cells = {n.split(":")[0] for n in ntc}
    assert "C7" in cells and "D7" not in cells and "D6" not in cells
    diff = [["増減"], [100], [120], [-30], [90], [110]]
    assert not [i for i in va.check_data_cleaner(diff, 0, 0) if i["type"] == "outlier_value"]


def test_dirt_notes_does_not_call_repeating_codes_duplicates():
    """明細の取引先コードの繰り返しは「番号列の重複」と言わない。伝票番号の 1 件の重複は言う（2026-09-23）。"""
    grid = [["伝票番号", "取引先コード", "金額"]] + [
        [f"{i:04d}" if i != 11 else "0010", "C%03d" % (i % 4 + 1), i * 100] for i in range(1, 13)]
    notes = " ".join(vv.dirt_notes(grid, 1, 1, 0, {}))
    assert "A12 = A11" in notes
    assert not re.search(r"B\d+ = B\d+", notes)


def test_seiri_hints_point_to_shelf_macros():
    """seiri が気づきごとに棚の手を出す（2026-09-23 通しの実測 2 の表の縮小版）: 空行・ゼロの消えた番号・空白の揺れ・
    年の無い文字の日付・式のずれ・合計の行の循環参照。式を揃える手が行を消す手より先に並ぶ。"""
    grid = [["伝票番号", "日付", "担当", "数量", "単価", "金額"],
            ["0001", "2026/5/1", "佐藤一郎", 10, 1200, 12000],
            ["0002", "2026/5/2", "鈴木花子", 5, 3000, 15000],
            [4, "5月16日", "佐藤 一郎", 3, 2000, 6000],
            [None, None, None, None, None, None],
            ["0005", "2026/5/9", "佐藤一郎", 6, 3000, 18000],
            ["0006", "2026/5/12", "鈴木花子", 2, 1500, 3000],
            ["0007", "2026/5/13", "田中次郎", 1, 1500, 1500],
            ["0008", "2026/5/14", "田中次郎", 2, 1500, 3000],
            ["0009", "2026/5/15", "鈴木花子", 4, 1500, 6000],
            [None, None, "合計", None, None, 0]]
    fa = [r[:] for r in grid]
    body = (1, 2, 3, 5, 6, 7, 8, 9)
    for i in body:
        fa[i][5] = f"=D{i + 1}*E{i + 1}"
    fa[3][5] = "=D4*E3"
    fa[10][5] = "=SUM(F2:F11)"
    fr = [r[:] for r in fa]
    for i in body:
        fr[i][5] = "=RC[-2]*RC[-1]"
    fr[3][5] = "=RC[-2]*R[-1]C[-1]"
    fr[10][5] = "=SUM(R[-9]C:R[0]C)"
    hints = []
    notes = " ".join(vv.dirt_notes(grid, 1, 1, 0, {}, hints=hints))
    assert "表の中の空行: 行5" in notes and "A4（4・列は 4 桁）" in notes and "C4（佐藤 一郎→佐藤一郎）" in notes
    assert "B4" in notes.split("文字の日付:")[1]
    hints += vv.seiri_formula_hints(fa, fr, grid, 1, 1, 0)
    names = [h[1] for h in sorted(hints)]
    for n in ("選んだ列の途切れた式を戻す", "表の中の空行を削除して詰める", "表の下に合計行を足す", "番号に先頭のゼロを付けてそろえる",
              "空白の有無を多い方にそろえる", "選んだ列の文字の日付を日付にする"):
        assert n in names, n
    assert names.index("選んだ列の途切れた式を戻す") < names.index("表の中の空行を削除して詰める")
    sel = [h[2] for h in hints if h[1] == "選んだ列の途切れた式を戻す"][0]
    assert sel.startswith("F2:F")


def test_diagnose_fix_does_not_write_without_a_backup():
    """「--fix は -y と控えを要る」なのに、控えに失敗しても書き換えていた（2026-10-28 点検）。控えが無ければ診断だけ。"""
    import inspect
    src = inspect.getsource(va.cmd_diagnose)
    i_backup = src.index("_diagnose_backup(wb)")
    i_guard = src.index("if not backup:", i_backup)
    i_apply = src.index("apply_audit_fixes(", i_backup)
    assert i_backup < i_guard < i_apply and "else:" in src[i_guard:i_apply]


# ---- 2026-11-05 dock の実測（Enron の公開ブック 60 本）で見つけた誤爆の最小の再現 ----

def test_circular_ignores_other_sheet_and_other_book_cells_20261105():
    """=Credit!A1・='[外.xls]Nominations'!I$23 の番地は別シートのセル。自分のシートの A1・I23 と取り違えて
    循環と言っていた（1 回目 38,153 件・Excel の判定は 0 件）。"""
    assert va.check_circular_references([["=+Credit!A1"]], 0, 0) == []
    assert va.check_circular_references([["='C:/x/[Nov.xls]Nominations'!A$1"]], 0, 0) == []
    assert va.check_circular_references([["=#REF!A1"]], 0, 0) == []
    assert va.check_circular_references([['=IF("A1"="",1,2)']], 0, 0) == []
    # 「(」の直後の番地も拾う（=ROUND(A1,2) の自己参照を見落としていた）
    assert len(va.check_circular_references([["=ROUND(A1,2)"]], 0, 0)) == 1


def _omitted(fa, fr, vals, r0=0, c0=0):
    return [i["cell"] for i in va.check_formula_linter(fa, fr, vals, r0, c0) if i["type"] == "omitted_sum_range"]


def test_omitted_sum_quiet_cases_20261105():
    """集計漏れの誤爆の型: 見出しの年・0・式の中で別に足した番地・別の小計が数えている塊・離れた所から塊を取る式。"""
    # 見出しの年 1998 の下の =SUM(A2:A3)
    assert _omitted([[1998], [10], [20], ["=SUM(A2:A3)"]], [[1998], [10], [20], ["=SUM(R[-2]C:R[-1]C)"]],
                    [[1998], [10], [20], [30]]) == []
    # =SUM(A1:A2)+A3（A3 は式の中で足している）
    assert _omitted([[10], [20], [5], ["=SUM(A1:A2)+A3"]], [[10], [20], [5], ["=SUM(R[-3]C:R[-2]C)+R[-1]C"]],
                    [[10], [20], [5], [35]]) == []
    # 塊ごとの小計: A6=SUM(A4:A5) が数えている A4 を、A7=SUM(A1:A3) の漏れと言わない
    fa = [[1], [2], [3], [4], [5], ["=SUM(A4:A5)"], ["=SUM(A1:A3)"]]
    assert _omitted(fa, fa, [[1], [2], [3], [4], [5], [9], [6]]) == []
    # 0 は足しても答えが変わらない
    assert _omitted([[10], [20], [0], ["=SUM(A1:A2)"]], [[10], [20], [0], ["=SUM(R[-3]C:R[-2]C)"]],
                    [[10], [20], [0], [30]]) == []
    # 挟まった数は今までどおり出す（A3 が漏れ）
    assert _omitted([[10], [20], [5], ["=SUM(A1:A2)"]], [[10], [20], [5], ["=SUM(R[-3]C:R[-2]C)"]],
                    [[10], [20], [5], [30]]) == ["A4"]
    # 行ごとにコピーした =SUM(B?:C?) の左の日付の列（A）は漏れではない
    fa = [[36951, 1, 2, "=SUM(B1:C1)"], [36952, 3, 4, "=SUM(B2:C2)"]]
    fr = [[36951, 1, 2, "=SUM(RC[-2]:RC[-1])"], [36952, 3, 4, "=SUM(RC[-2]:RC[-1])"]]
    assert _omitted(fa, fr, [[36951, 1, 2, 3], [36952, 3, 4, 7]]) == []


def test_hardcoded_constant_skips_units_and_paths_20261105():
    """/12・/365・/1000 の単位の数と、別ブックの道のりの中の数（1990-1999）は咎めない。*1.001 は出す。"""
    def hc(f):
        return [i for i in va.check_formula_linter([[f]], [[f]], [[1]], 0, 1) if i["type"] == "hardcoded_constant"]
    assert hc("=A1/12") == [] and hc("=A1/1000") == [] and hc("=A1*365") == []
    assert hc("=AVERAGE(A1,'C:/TEMP/[NYMEX 1990-1999.xls]1998'!G254)") == []
    assert len(hc("=A1*1.001")) == 1


def test_inconsistent_skips_same_column_summaries_20261105():
    """節ごとの小計を拾う =A1+A3（同じ列だけを足す式）は、明細の =B1*C1 と形が違っても咎めない。"""
    fa = [["=B1*C1"], ["=B2*C2"], ["=B3*C3"], ["=B4*C4"], ["=A1+A3"]]
    fr = [["=RC[1]*RC[2]"]] * 4 + [["=R[-4]C+R[-2]C"]]
    assert [i for i in va.check_formula_linter(fa, fr, [[1]] * 5, 0, 0) if i["type"] == "inconsistent_formula"] == []


def test_value_in_formula_column_needs_neighbors_20261105():
    """同じ形の式に上下 2 行以内で挟まれた数だけを直書きと言う。=#N/A の置き場は式の列の形にしない。"""
    fr = [["=R[-1]C+1"]] + [[None]] * 5 + [["=R[-1]C+1"]] + [[7]] + [["=R[-1]C+1"]]
    fa = [[x[0]] for x in fr]
    vals = [[1], [None], [None], [None], [None], [None], [2], [7], [3]]
    hit = va.check_values_in_formula_columns(fa, fr, vals, 0, 0)
    assert [i["cell"] for i in hit] == ["A8"]
    na = [["=#N/A"], ["=#N/A"], [4.45], ["=#N/A"]]
    assert va.check_values_in_formula_columns(na, na, [[0], [0], [4.45], [0]], 0, 0) == []


def test_data_cleaner_quiet_cases_20261105():
    """空白だけ・字下げ・数の列の上の見出しの空白、番号の列・見出しの年の文字の数字、英語の合計の行、
    列の頭の年・日付、値動きの列の多いマイナスは咎めない。"""
    vals = [["Item", "2000", "Code", "Amount "], ["   NET MTM", 10, "100061", 1], [" ", 12, "100062", 2],
            ["a", 11, "100063", 3], ["b", 13, 5, 4], ["Total", 900, "", 10]]
    iss = va.check_data_cleaner(vals, 0, 0)
    assert [i for i in iss if i["type"] == "hidden_whitespace"] == []
    assert [i for i in iss if i["type"] == "text_number"] == []
    assert [i for i in iss if i["type"] == "outlier_value"] == []
    col = [[36951]] + [[v] for v in (0.03, 0.031, 0.029, 0.03, 0.032)]
    assert [i for i in va.check_data_cleaner(col, 0, 0) if i["type"] == "outlier_value"] == []
    moves = [[v] for v in ([1.0, 2.0, -0.5, 1.5, -0.2, 0.8, -0.1, 1.2, 0.9, 1.1] * 4)]
    assert [i for i in va.check_data_cleaner(moves, 0, 0) if i["type"] == "outlier_value"] == []


def test_inconsistent_formula_rows_and_sandwich_20261105():
    """横に並んだ式のずれ（AA=+F6・AB=+F6・AC=+H6 の AB＝G を指すはず）を言う。列の多数決の少数派でも、
    同じ形の式に挟まれていなければ言わない（表でない計算の格子で 509 件出し、Excel と重なったのは 9 件だった）。"""
    fa = [["=B2", "=B2", "=D2", "=E2", "=F2"]]
    fr = [["=R[1]C[1]", "=R[1]C", "=R[1]C[1]", "=R[1]C[1]", "=R[1]C[1]"]]
    hit = [i["cell"] for i in va.check_formula_linter(fa, fr, [[1] * 5], 0, 0) if i["type"] == "inconsistent_formula"]
    assert hit == ["B1"]
    # 行の年計 =SUM(RC[-3]:RC[-1]) は横の明細と形が違っても言わない
    fr2 = [["=R[1]C", "=R[1]C", "=R[1]C", "=SUM(RC[-3]:RC[-1])"]]
    assert [i for i in va.check_formula_linter(fr2, fr2, [[1] * 4], 0, 0) if i["type"] == "inconsistent_formula"] == []
    # 列の端の違う式（挟まれていない）は言わない
    col = [["=RC[1]*RC[2]"]] * 4 + [["=RC[1]+RC[2]"]]
    assert [i for i in va.check_formula_linter(col, col, [[1]] * 5, 0, 0) if i["type"] == "inconsistent_formula"] == []


def test_inconsistent_formula_long_column_and_constants_20261105():
    """800 行の列の中の数個のずれ（=LN(E32/E30)）は拾う。2 種類の式が交互に並ぶ作り・番地の無い式（=3100000/2）・
    1 つ飛ばしの横の式（値・前年差の組）は言わない。"""
    good, bad = "=LN(RC[-2]/R[-1]C[-2])", "=LN(RC[-2]/R[-2]C[-2])"
    col = [[good]] * 40
    col[10] = [bad]; col[25] = [bad]; col[33] = [bad]
    hit = [i["cell"] for i in va.check_formula_linter(col, col, [[1]] * 40, 0, 0) if i["type"] == "inconsistent_formula"]
    assert hit == ["A11", "A26", "A34"]
    alt = [["=R[1]C", "=R[2]C"] * 4]
    assert [i for i in va.check_formula_linter(alt, alt, [[1] * 8], 0, 0) if i["type"] == "inconsistent_formula"] == []
    const = [["=R[1]C", "=3100000/2", "=R[1]C", "=R[1]C"]]
    assert [i for i in va.check_formula_linter(const, const, [[1] * 4], 0, 0) if i["type"] == "inconsistent_formula"] == []
    pairs = [["=R[-12]C[-1]", 5, "=RC[-1]-RC[2]", "=R[-12]C[-1]"]]
    assert [i for i in va.check_formula_linter(pairs, pairs, [[1] * 4], 0, 0) if i["type"] == "inconsistent_formula"] == []


def _outl(vals):
    return [i["cell"] for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "outlier_value"]


def test_outliers_quiet_cases_20261105():
    """外れ値の誤爆の型（Enron の公開ブック 60 本: 476 件 → 本物らしい物だけ）: 2 つ目の集まり・累計の行・
    表から離れた検算の数・上の数の和・英語の損益の見出しのマイナス・大きい順に並べた順位表。"""
    # 2 つ目の集まり: 0.005 の刻みの列に 0.05 が 4 つ（ぽつんと 1 つではない）
    assert _outl([[v] for v in [0.005] * 12 + [0.05] * 4]) == []
    # 累計の行（'CUMULATIVE'）は明細と桁が違って当たり前
    assert _outl([["", "金額"], ["CUMULATIVE", 255224], ["1", 728], ["2", 3370], ["3", 905], ["4", 1200], ["5", 640]]) == []
    # 表の本体から空行 3 つ以上離れた数
    col = [[v] for v in (728, 3370, 905, 1200, 640, 2200)] + [[None]] * 4 + [[3161324]]
    assert _outl(col) == []
    # 上の数の和（見出しの無い合計）
    assert _outl([[v] for v in (100, 120, 130, 110, 140, 600)]) == []
    # 英語の損益・差の見出しの列のマイナス
    assert _outl([["Daily P&L"], [120], [340], [210], [-59.7], [450]]) == []
    # 大きい順の順位表: 1 位と最下位のマイナスは言わない
    rank = [["Total"]] + [[v] for v in (18060000, 11017547, 3982771, 2100000, 900000, 500000, 273, 232, -54615, -57505)]
    assert _outl(rank) == []
    # 並べた列でも、隣から 10 倍以上の崖は拾う（今までどおり）
    assert _outl([[v] for v in (100, 110, 120, 130, 140, 150, 100000)]) == ["A7"]


def test_heldout_quiet_cases_20261105():
    """初見の 60 本（前の 60 本と重ならない Enron の公開ブック）で出た誤爆の型: 月の日数で割る式の違い・
    値の行と式の行を繰り返す作り・横に並んだ日付の見出しの行。"""
    # =R[-1]C/30・/31・/28 の違いは月の日数
    col = [["=R[-1]C/31"], ["=R[-1]C/30"], ["=R[-1]C/31"], ["=R[-1]C/28"], ["=R[-1]C/31"], ["=R[-1]C/30"]]
    row = [[x[0] for x in col]]
    assert [i for i in va.check_formula_linter(row, row, [[1] * 6], 0, 0) if i["type"] == "inconsistent_formula"] == []
    # 値・式・0 の 3 行の繰り返しの列の値は「式の列に直書き」ではない
    fr, vals = [], []
    for v in (191, 230, 160, 179):
        fr += [[v], ["=R[-1]C/30"], [0]]
        vals += [[v], [v / 30], [0]]
    assert va.check_values_in_formula_columns(fr, fr, vals, 0, 0) == []
    # 日付の見出しの行（7 行目に月のシリアルが横に並ぶ）は列ごとの桁違いにしない
    grid = [["", "a", "b", "c", "d"]] + [["x", 1100 + k, 990 + k, 1340 + k, 1000 + k] for k in range(5)]
    grid.insert(3, ["月", 37073, 37104, 37135, 37165])
    assert [i for i in va.check_data_cleaner(grid, 0, 0) if i["type"] == "outlier_value"] == []


def test_outliers_skip_stacked_report_20261105():
    """行ごとに意味の違う数を縦に積んだ報告（件数 64・損益 -3,700 万・差 18.8 万）は、桁違いを比べる相手がいない
    （初見の 60 本の e026・前の 60 本の e059）。値段 30〜12,000 の列の 2,500,000 は今までどおり拾う。"""
    rep = [["Calls", 64.4], ["Deals", -37721947], ["Load", -49435451], ["Generation", 11798127],
           ["Total MTM", -75359271], ["Prior Day MTM", -75547524], ["Delta", 188253], ["Fuel", 0.82], ["Cargo", 119000]]
    assert _outl(rep) == []
    price = [[480], [120], [300], [1500], [2500000], [8500], [1200], [480], [30], [90], [700], [12000], [150]]
    assert _outl(price) == ["A5"]


def test_japanese_public_books_quiet_cases_20261105():
    """日本語の公開ブック（東京都と区市のオープンデータ）60 本で出た誤爆の型: 全角 1 字の字下げ・注記の文の空白・
    「〇〇地区計」「総数」の行・全角の番号の見出しの列・「計」の下の「都心３区」の部分集計。"""
    vals = [["地区", "人口"], ["大森地区計", 127467], ["　池上五丁目", 3512], ["　池上六丁目", 3400], ["　中央四丁目", 3300],
            ["　山王二丁目", 3450], [" なお、「２か所以上の所に掲記」は再掲である。", None]]
    iss = va.check_data_cleaner(vals, 0, 0)
    assert [i for i in iss if i["type"] in ("hidden_whitespace", "outlier_value")] == []
    assert va._is_total_label("大森地区計") and va._is_total_label("総 数")
    assert not va._is_total_label("設計") and not va._is_total_label("統計")
    nums = [["３", 79792], ["４", 72381], ["９", 126760], ["10", 130549], ["11", 128000], [5, 1]]
    assert [i for i in va.check_data_cleaner(nums, 0, 0) if i["type"] == "text_number"] == []
    # 「計」の行の下の「都心３区」=SUM(A1:A3) は部分集計（漏れではない）
    fa = [[571], [114], [82], [59], [28], ["=SUM(A1:A5)"], [None], ["=SUM(A1:A3)"]]
    fr = [[571], [114], [82], [59], [28], ["=SUM(R[-5]C:R[-1]C)"], [None], ["=SUM(R[-7]C:R[-5]C)"]]
    vv_ = [[571], [114], [82], [59], [28], [854], [None], [767]]
    assert [i for i in va.check_formula_linter(fa, fr, vv_, 0, 0) if i["type"] == "omitted_sum_range"] == []


def test_padded_names_stay_but_notes_go_20261105():
    """右を空白で埋めた長い会社名の列は照合に響く本物（言う）。説明の文の空白は言わない。"""
    vals = [["Customer"], ["NOBLE GAS MARKETING INC   "], ["EL PASO MERCHANT ENERGY  "], ["INFINITE ENERGY"],
            ["CITRUS TRADING"], ["the price is set by the index  "]]
    hit = [i["cell"] for i in va.check_data_cleaner(vals, 0, 0) if i["type"] == "hidden_whitespace"]
    assert hit == ["A2", "A3"]


_SHU_H = '=IF(G{r}="","",ROUNDDOWN(SUM(MAP(TEXTSPLIT(G{r},"+"),LAMBDA(t,PRODUCT(--TEXTSPLIT(t,"*"))))),2))'
_SHU_I = ('=IF(OR(E{r}="幕",E{r}="旗",E{r}="他"),"－　",IF(H{r}="","",LET(base,IFS(H{r}<1,IF(F{r}="有",$W$1,$W$2),'
          'H{r}<5,IF(F{r}="有",$X$1,$X$2),H{r}<10,IF(F{r}="有",$Y$1,$Y$2),TRUE,IF(F{r}="有",$Z$1,$Z$2)),'
          'IF(H{r}>=10,base+(ROUNDDOWN(H{r},0)-10)*100,base))))')
_SHU_H1 = '=IF(RC[-1]="","",ROUNDDOWN(SUM(MAP(TEXTSPLIT(RC[-1],"+"),LAMBDA(t,PRODUCT(--TEXTSPLIT(t,"*"))))),2))'
_SHU_I1 = ('=IF(OR(RC[-4]="幕",RC[-4]="旗",RC[-4]="他"),"－　",IF(RC[-1]="","",LET(base,IFS(RC[-1]<1,IF(RC[-3]="有",R1C23,R2C23),'
           'RC[-1]<5,IF(RC[-3]="有",R1C24,R2C24),RC[-1]<10,IF(RC[-3]="有",R1C25,R2C25),TRUE,IF(RC[-3]="有",R1C26,R2C26)),'
           'IF(RC[-1]>=10,base+(ROUNDDOWN(RC[-1],0)-10)*100,base))))')


def test_shu_keisanhyo_20261105():
    """シュウさんの計算表（写真）の式で: ROW()-3 の連番は言わない・見出しの行の ' 区　分' は言わない・
    料金の決まりの直書きは「固定値 10（引く）」と数だけで言う（気づき）・LET・MAP・LAMBDA・TEXTSPLIT の式を壊れずに読む
    （循環・エラー・隣と違う式・範囲漏れを出さない）。"""
    hdr = ["区分", " 区　分", "有無", "寸法", "面積", "金額", "番号"]       # E〜 の見出しの行（3 行目）
    fa, fr, vals = [hdr], [hdr], [hdr]
    for r in range(4, 9):
        fa.append(["野立広告板", "看板", "有", "1.5*1.26", _SHU_H.format(r=r), _SHU_I.format(r=r), f'=IF(AD{r}<>0,ROW()-3,"")'])
        fr.append(["野立広告板", "看板", "有", "1.5*1.26", _SHU_H1, _SHU_I1, '=IF(RC[-9]<>0,ROW()-3,"")'])
        vals.append(["野立広告板", "看板", "有", "1.5*1.26", 1.89, 900, r - 3])
    iss = va.check_formula_linter(fa, fr, vals, 2, 3) + va.check_data_cleaner(vals, 2, 3)
    by = {}
    for i in iss:
        by.setdefault(i["type"], []).append(i)
    assert set(by) == {"hardcoded_constant"}, by.keys()
    hc = by["hardcoded_constant"]
    assert {i["cell"][0] for i in hc} == {"I"} and len(hc) == 5
    assert all(i["severity"] == "info" and "固定値 10（引く）" in i["msg"] and "'-10'" not in i["msg"] for i in hc)
    # 点数: 同じ式のコピー 5 行の気づきは 1 回（2 点）
    assert va.audit_score(hc) == 98


def test_dockC_quiet_and_kept_cases_20261105():
    """段 C（台帳の誤爆の型）: 見出しの年・時刻・日の並び・記入欄の番号・内訳・同じ値の刻み・初期値・別に足す番地は言わない。本物は残す。"""
    # 文字の数字: 時刻の並び・見出しの年・日の並びは言わない
    vals = [["0500", 1], ["0600", 2], ["0700", 3], ["0800", 4], ["1000", 5]]
    assert not [i for i in va.check_data_cleaner([["x", "1998", "Total"], ["a", 5, 6], ["b", 7, 8]], 0, 0) if i["type"] == "text_number"]
    # 並びの途中の 1 つだけが文字の数字（通し番号・日の列）は入力の誤りとして言う（2026-11-05 段 C3: 注入テストで 9 件見落とし）
    seq = [[13.0, 5.0], [14.0, 6.0], [15.0, 7.0], [16.0, 8.0], ["17", 9.0], [18.0, 1.0]]
    assert [i["cell"] for i in va.check_data_cleaner(seq, 0, 0) if i["type"] == "text_number"] == ["A5"]
    # ブロックごとに繰り返す日の番号（同じ行に同じ '18' が 3 つ）は見出しで言わない
    rep = [[15.0, 1.0, 15.0, 1.0, 15.0], [16.0, 1.0, 16.0, 1.0, 16.0], [17.0, 1.0, 17.0, 1.0, 17.0],
           ["18", 1.0, "18", 1.0, "18"], [19.0, 1.0, 19.0, 1.0, 19.0], [20.0, 1.0, 20.0, 1.0, 20.0]]
    assert not [i for i in va.check_data_cleaner(rep, 0, 0) if i["type"] == "text_number"]
    # 本物: 数の列に 1 つだけ紛れた文字の数字は言う
    amounts = [[120.5, 3.0], [98.0, 4.0], ["77", 5.0], [310.0, 6.0], [41.0, 7.0]]
    assert [i["cell"] for i in va.check_data_cleaner(amounts, 0, 0) if i["type"] == "text_number"] == ["A3"]
    # 外れ値: 同じ値が 2 つの刻みは言わない
    ticks = [[5.1], [5.0], [4.9], [4.8], [4.9], [-0.005], [-0.005]]
    assert not [i for i in va.check_data_cleaner(ticks, 0, 0) if i["type"] == "outlier_value"]
    # 内訳: 真上の合計に合う並びは桁違いと言わない
    brk = [[8000000.0], [10.0], [7000000.0], [450000.0], [550000.0], [6000000.0]]
    assert not [i for i in va.check_data_cleaner(brk, 0, 0) if i["type"] == "outlier_value" and i["cell"] in ("A4", "A5")]
    # 集計漏れ: =F5+F38 と別の式が足している先頭の残高は言わない
    fa = [["", "=SUM(B3:B6)+B7", ""], ["", "=B1+B2", ""]]
    # 隣と違う式: 小計の行（同じ列だけを足す式が同じ行にある）の差は言わない
    fa = [["=B2", "=SUM(B$5,B$6)", "=SUM(C$5,C$6)"], ["=B3", "=B3*2", "=B3"]]
    fr = [["=R[1]C", "=SUM(R5C,R6C)", "=SUM(R5C,R6C)"], ["=R[1]C", "=R[0]C*2", "=R[0]C"]]
    assert not [i for i in va.check_formula_linter(fa, fr, [[1, 2, 3], [4, 5, 6]], 0, 0) if i["type"] == "inconsistent_formula"]


def test_text_number_speaks_for_csv_amounts_and_stays_quiet_for_codes_20261105():
    """段 C の追加確認: 文字の数字が言われるべき型（CSV の金額の列）と、言わなくてよい型（郵便番号・電話・年の見出し）。"""
    amt = ['1200', '3450', '780', '15200', '99', '5400', '23100', '410', '8800', '670']
    nums = [100.0, 250.0, 80.0, 310.0, 45.0, 60.0, 275.0, 90.0, 120.0, 33.0]

    def tn(grid):
        return [i["cell"] for i in va.check_data_cleaner(grid, 0, 0) if i["type"] == "text_number"]
    # (a) 見出しの下に金額が全部文字・隣は本物の数
    assert tn([['品名', '金額', '数量']] + [[f'商品{i}', amt[i], nums[i]] for i in range(10)]) == [f"B{i}" for i in range(2, 12)]
    # (b) 文字の数字ばかりの独立した列
    assert tn([['金額']] + [[x] for x in amt]) == [f"A{i}" for i in range(2, 12)]
    # (c) 本物の数の列の途中に 2 個
    assert tn([['金額']] + [[str(int(n)) if i in (3, 7) else n] for i, n in enumerate(nums)]) == ["A5", "A9"]
    # (d) 売上の列が全部文字の数字で下に SUM（値は 0）
    assert tn([['売上']] + [[x] for x in amt] + [[0.0]]) == [f"A{i}" for i in range(2, 12)]
    # (e) 先頭ゼロの郵便番号・電話番号は言わない
    assert tn([['郵便番号', '電話番号']] + [[f'{1000000 + i * 137:07d}', f'0{9012345000 + i * 7}'] for i in range(10)]) == []
    # (f) 年の見出し行は言わない
    assert tn([['2020', '2021', '2022', '2023', '2024']] + [[10.0 + i + j for j in range(5)] for i in range(5)]) == []
    # 番号の列（見出しがコード・桁数がそろう）は言わない／繰り返す区分の値は言わない
    assert tn([['勘定科目コード']] + [[f'{100000 + i * 7}'] for i in range(10)]) == []
    assert tn([['区分', '値']] + [['1', 0.5 + i] for i in range(6)]) == []


def test_outlier_c2_real_typos_are_said_and_row_column_context_is_quiet_20261105():
    """段 C2（outlier_value）: 本当の桁違い（1000 倍の入力ミス）は言い、同じ行の同規模・短い塊・設定の行・桁の揺れる列・
    符号のそろわない小口は言わない。番号の列の桁欠けは別の言い方で言う。"""
    def o(g, f=None):
        return [i["cell"] for i in va.check_data_cleaner(g, 0, 0, f) if i["type"] == "outlier_value"]
    prices = [120, 95, 150, 88, 130, 45000, 110, 99, 140, 105, 125, 90]
    qty = [12, 8, 15, 9, 20000, 11, 10, 14, 7, 13, 16, 9]
    sales = [320, 410, 290, 380, 350, 360000, 400, 330, 420, 310, 390, 370]
    # 言う: 単価・数量・月別売上（下に合計）・小数の列の 100 倍
    assert o([["品名", "単価"]] + [[f"P{i}", p] for i, p in enumerate(prices)]) == ["B7"]
    assert o([["品名", "数量"]] + [[f"P{i}", q] for i, q in enumerate(qty)]) == ["B6"]
    assert o([["月", "売上"]] + [[f"{i + 1}月", v] for i, v in enumerate(sales)] + [["合計", sum(sales)]]) == ["B7"]
    assert o([["比率"]] + [[v] for v in [0.52, 0.48, 0.55, 0.5, 0.6, 0.45, 0.58, 0.49, 0.51, 0.57, 0.5, 52]]) == ["A13"]
    # 単価の打ち間違いが 数量×単価 の式の金額の列に写っても言う（式のセルは同じ行の相手に数えない）
    g = [["品名", "単価", "数量", "金額"]] + [[f"P{i}", p, 10, p * 10] for i, p in enumerate(prices)]
    f = [["", "", "", ""]] + [["", "", "", f"=B{i + 2}*C{i + 2}"] for i in range(12)]
    assert [x for x in o(g, f) if x != "D7"] == ["B7"]      # D7（式のセル）は dock が重ねて出さない
    assert "B7" in o(g)        # 式の情報が無くても、同規模の相手が 1 つ（2 倍離れた金額）だけなら黙らない
    # 言わない: 同じ行に同規模の値が並ぶ（400000 の繰り返し）・設定の行（短い塊とラベル）
    row = [[None, 4e4, 4e5, 4e5, 4e5, 4e5], [None, 8e3, 4e5, 4e5, 4e5, 4e5], [None, 7e3, 4e5, 4e5, 4e5, 4e5],
           [None, 5e3, 4e5, 4e5, 4e5, 4e5], [None, 6e3, 4e5, 4e5, 4e5, 4e5], [None, 9e3, 4e5, 4e5, 4e5, 4e5]]
    assert o(row) == []
    assert o([["Number of Rows", 32], ["Peak", 700, 2200]]) == []
    # 番号の列の桁欠け（7 桁の Org ID の中の 1853）は桁数の言い方で言う
    ids = [["Org ID"]] + [[1029300 + i] for i in range(10)] + [[1853], [1029311]]
    got = [i for i in va.check_data_cleaner(ids, 0, 0) if i["type"] == "outlier_value"]
    assert [i["cell"] for i in got] == ["A12"] and "桁数が列と違います" in got[0]["msg"]
    # 金額の列（先頭の桁がばらばら）の 1 つだけ桁が少ないのは番号の桁欠けとは言わない
    amt = [["給与"]] + [[x] for x in [310000, 420000, 285000, 530000, 610000, 390000, 270000, 480000, 95000, 350000]]
    assert not [i for i in va.check_data_cleaner(amt, 0, 0) if "桁数が列と違います" in i["msg"]]


def test_dockC3_misses_found_by_injection_20261105():
    """段 C3（注入テストの見落とし）: 直書きが式の先頭・合計が範囲の上・行の合計に数えられただけの漏れ・比率の式・
    式の列に貼り付けた 0・全角数字・明白な桁違いを言う。小計・別の合計は言わない。"""
    def lint(f, v, r0=0, c0=0):
        return va.check_formula_linter(f, f, v, r0, c0)
    # hardcoded: 先頭のオペランド
    for fx in ("=856-N17", "=2231+C8", "=515/$E$6", "=IF(A1>0,99-$N$12,0)"):
        assert [i for i in lint([[fx]], [[1]]) if i["type"] == "hardcoded_constant"], fx
    for fx in ("=1+A1", "=100*A1", "=SUM(A1:A3)", "=DATE(2020,1,1)"):
        assert not [i for i in lint([[fx]], [[1]]) if i["type"] == "hardcoded_constant"], fx
    # omitted: 合計が範囲の真上（G12=SUM(G13:G15) の下の G16 が漏れ）
    f = [["=SUM(A2:A4)"], [1], [2], [3], [4]]
    v = [[6], [1], [2], [3], [4]]
    assert [i for i in lint(f, v) if i["type"] == "omitted_sum_range"]
    # 漏れた相手が桁違いに大きい（それ自身が別の合計）なら言わない
    v2 = [[6], [1], [2], [3], [4000]]
    assert not [i for i in lint([["=SUM(A2:A4)"], [1], [2], [3], [4000]], v2) if i["type"] == "omitted_sum_range"]
    # omitted: 相手が行の合計（=SUM(B3:C3)）に数えられているだけでは黙らない・比率の式（=ROUND(A4/A5)）も漏れの裏付け
    f = [[1, 2, "=SUM(A1:B1)"], [3, 4, "=SUM(A2:B2)"], [5, 6, "=SUM(A3:B3)"], ["=SUM(A1:A2)", "=SUM(B1:B3)", "=SUM(C1:C3)"]]
    v = [[1, 2, 3], [3, 4, 7], [5, 6, 11], [4, 9, 21]]
    assert [i["cell"] for i in lint(f, v) if i["type"] == "omitted_sum_range"] == ["A4"]
    # value_in_formula_column: 8 行以上の式の列の途中の 0 の貼り付けは言う。短い列の 0 は言わない
    fa12 = [[k + 1, "=A%d*2" % (k + 1)] for k in range(12)]
    f1_12 = [[k + 1, "=RC[-1]*2"] for k in range(12)]
    v12 = [[k + 1, (k + 1) * 2] for k in range(12)]
    fa12[5][1] = 0
    f1_12[5][1] = 0
    v12[5][1] = 0
    assert [i["cell"] for i in va.check_formula_linter(fa12, f1_12, v12, 0, 0) if i["type"] == "value_in_formula_column"] == ["B6"]
    fa5, f15, v5 = [r[:] for r in fa12[:5]], [r[:] for r in f1_12[:5]], [r[:] for r in v12[:5]]
    fa5[2][1], f15[2][1], v5[2][1] = 0, 0, 0
    assert not [i for i in va.check_formula_linter(fa5, f15, v5, 0, 0) if i["type"] == "value_in_formula_column"]
    # text_number: 全角数字の文字列でも、数の列の途中なら言う。見出しの全角番号（全部文字）は言わない
    col = [[120.0], [98.0], ["７７"], [310.0], [41.0], [55.0], [60.0]]
    assert [i["cell"] for i in va.check_data_cleaner(col, 0, 0) if i["type"] == "text_number"] == ["A3"]
    head = [["１"], ["２"], ["３"], ["４"], ["５"]]
    assert not [i for i in va.check_data_cleaner(head, 0, 0) if i["type"] == "text_number"]
    # outlier: 件数と金額が横に並ぶ表の、件数の列の 2 → 2000（横の金額が相手に見えても、列の最大の 100 倍以上の桁違いは言う）
    rows = [[2.0, 501.0, 9000.0], [1.0, 450.0, 8500.0], [0.0, 520.0, 7000.0], [3.0, 610.0, 9100.0], [2.0, 399.0, 8800.0],
            [2000.0, 501.0, 9200.0], [1.0, 430.0, 9400.0], [2.0, 440.0, 8700.0], [3.0, 470.0, 9900.0], [1.0, 390.0, 9500.0]]
    assert [i["cell"] for i in va.check_data_cleaner(rows, 0, 0) if i["type"] == "outlier_value"] == ["A6"]


def test_dockC3_flat_column_and_structured_refs_20261105():
    """段 C3: ほぼ全部が同じ値の列（四分位範囲 0）の 1 → 1000 を言う。テーブルの計算列（構造化参照）・ROW() の列の途中の直書きを言う。"""
    flat = [[1.0]] * 10 + [[1000.0]] + [[1.0]] * 4
    assert [i["cell"] for i in va.check_data_cleaner(flat, 0, 0) if i["type"] == "outlier_value"] == ["A11"]
    # 同じ値でない小さな揺れの列は言わない・全部 0 の列は言わない
    assert not [i for i in va.check_data_cleaner([[0.0]] * 10 + [[1000.0]] + [[0.0]] * 4, 0, 0) if i["type"] == "outlier_value"]
    # 構造化参照の計算列
    fa = [[k + 1, '=[@a]*2'] for k in range(12)]
    f1 = [[k + 1, '=[@a]*2'] for k in range(12)]
    vv = [[k + 1, (k + 1) * 2] for k in range(12)]
    fa[6][1], f1[6][1], vv[6][1] = 99, 99, 99
    assert [i["cell"] for i in va.check_formula_linter(fa, f1, vv, 0, 0) if i["type"] == "value_in_formula_column"] == ["B7"]



def test_text_number_fullwidth_minus_and_comma():
    """全角のマイナス・全角/半角のカンマの文字の数字も言う（2026-11-06 日本語の注入）。"""
    for s in ("－７２５", "１，８６７", "1,200"):
        values = [[100], [200], [s], [300], [400], [500], [600]]
        issues = va.check_data_cleaner(values, 0, 0)
        assert any(i["type"] == "text_number" and i["cell"] == "A3" for i in issues), s
    assert va._tn_value("－７２５") == -725 and va._tn_value("１，２００") == 1200


def test_circular_ignores_japanese_unquoted_sheet_ref():
    """別シート（日本語の素のシート名）の同じ番地を、自シートの循環と読まない。"""
    f = [["=データシート!A1"], ["=1"]]
    assert va.check_circular_references(f, 0, 0) == []


def test_inconsistent_formula_skips_external_book_and_paren_ref():
    f1 = [["=R[0]C[-1]/8977"]] * 3
    assert va._RE_PAREN_SINGLE_REF.sub(r"\1", "=(RC[-1])/8977") == "=RC[-1]/8977"
    assert va._RE_EXTERNAL_BOOK.search("='C:/x/[決算.xls]S'!R4C3")


def test_hidden_whitespace_layout_vs_key():
    """和暦年度の字下げ・列全体の体裁は言わず、列の中の少数派の末尾空白（照合キー）は言う。"""
    col = [["名称"]] + [[" H%02d" % n] for n in range(10, 20)]
    assert not [i for i in va.check_data_cleaner(col, 0, 0) if i["type"] == "hidden_whitespace"]
    keys = [["区名"]] + [["港区"], ["新宿区"], ["中央区"], ["千代田区"], ["品川区"], ["目黒区"], ["大田区"], ["鶴見区\u3000\u3000"]]
    keys[0:0] = [["a"], ["b"]]
    assert [i for i in va.check_data_cleaner(keys, 0, 0) if i["type"] == "hidden_whitespace"]


def test_hardcoded_constant_skips_calendar_and_offset_types_20261106():
    """英語ブックの誤爆の型: /2（半分）・/10000（万換算）・月別日数 28 の並び・^(-2*…) の指数・EOMONTH+20 の暦ずらし・
    VLOOKUP/LOOKUP の中の +4 の番地オフセットは咎めない。*2・*28 単独・+500 は出す。"""
    def hc(f):
        return [i for i in va.check_formula_linter([[f]], [[f]], [[1]], 0, 1) if i["type"] == "hardcoded_constant"]
    for fx in ("=3100000/2", "=(SUM(A1+B1/2)*(0.065/12))", "=E21*24*H21/10000",
               "=ROUND((O24*31+P24*28+Q24*31+R24*30)*H24,0)", "=(1+B1/2)^(-2*(C1-D1)/365.25)",
               "=(EOMONTH(A1,0)+20)-C1", "=VLOOKUP($A1,T,M$4+4,FALSE)", "=LOOKUP($K$15+2,A:A,B:B)"):
        assert hc(fx) == [], fx
    for fx in ("=A1*2", "=A1*28", "=A1+500", "=A1*1.001"):
        assert len(hc(fx)) == 1, fx


def test_hardcoded_constant_types_20261106_dummy():
    assert va._enclosing_func("IF(A1,VLOOKUP(B1,T,C1+4,0))", 22) == "VLOOKUP"


def test_outlier_quiet_when_value_is_period_days_in_row():
    """行に '11月12日〜12月25日' があり値が日数（44）なら、日数の列の桁違いとは言わない。期間と合わない値は出す。"""
    def run(days):
        rows = [[1, "a", "4月5日〜4月8日", 3], [2, "b", "5月1日〜5月4日", 3], [3, "c", "6月1日〜6月3日", 2],
                [4, "d", "7月1日〜7月4日", 4], [5, "e", "11月12日〜12月25日", days],
                [6, "f", "8月1日〜8月3日", 3], [7, "g", "9月1日〜9月4日", 3]]
        return [i for i in va.check_data_cleaner(rows, 0, 0, rows) if i["type"] == "outlier_value"]
    assert run(44) == []
    assert any(i["cell"] == "D5" for i in run(440))
