"""インサイト・データ診断エンジン (vbam_audit) およびスタイルマップのユニットテスト
"""

import os
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
    assert "著しく乖離しています" in outliers[0]["msg"]


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


def test_text_number_safe_to_convert():
    assert va._text_number_safe_to_convert("12345")
    assert va._text_number_safe_to_convert("0")
    assert va._text_number_safe_to_convert("0.5")
    assert va._text_number_safe_to_convert("-7")
    assert not va._text_number_safe_to_convert("0123")       # 先頭ゼロ＝コード
    assert not va._text_number_safe_to_convert("0001234567")
    assert not va._text_number_safe_to_convert("1234567890123456")  # 16 桁


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
