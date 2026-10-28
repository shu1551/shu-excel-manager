# -*- coding: utf-8 -*-
"""dirt_notes の「未記入の段」（見出しの頭の語が同じ列の組が、ある行でそろって空）の試験（2026-10-29）。純 Python・COM なし。

10/24 の試験の表（貸出簿）の語（返却・貸出中）の決め打ちを止めた。どの台帳でも当たり、当たってはいけない形では出ないこと。"""
import vbam_view as vv


def _notes(grid):
    return vv.dirt_notes(grid, r0=1, c0=1, header_idx=0)


def _stage(notes):
    return [n.strip() for n in notes if "未記入の段" in n]


def _blanks(notes):
    return next((n for n in notes if "表の中の空欄" in n), "")


def test_lending_book_return_stage():
    grid = [
        ["番号", "品名", "貸出日", "返却日", "返却担当"],
        [1, "ノートPC", "2026/01/05", "2026/01/08", "佐藤"],
        [2, "プロジェクター", "2026/01/06", None, None],
        [3, "カメラ", "2026/01/07", "2026/01/10", "高橋"],
        [4, "マイク", "2026/01/08", "2026/01/12", "鈴木"],
    ]
    notes = _notes(grid)
    assert _stage(notes) == ["未記入の段: 返却（行 3 [2]）"]
    assert _blanks(notes) == ""                       # 未記入の段の空欄は「表の中の空欄」に数えない


def test_approval_ledger_stage():
    grid = [
        ["申請番号", "申請者", "申請日", "承認日", "承認者"],
        ["A-01", "田中", "2026/02/01", "2026/02/02", "部長"],
        ["A-02", "小原", "2026/02/03", "2026/02/04", "課長"],
        ["A-03", "佐倉", "2026/02/05", None, None],
        ["A-04", "山田", "2026/02/06", "2026/02/07", "部長"],
        ["A-05", "林", "2026/02/08", None, "　"],     # 全角の空白だけも空
    ]
    notes = _notes(grid)
    assert _stage(notes) == ["未記入の段: 承認（行 4 [A-03]・6 [A-05]）"]
    assert _blanks(notes) == ""


def test_payment_ledger_stage_and_other_blank_still_counted():
    grid = [
        ["取引先", "請求額", "支払日", "支払額"],
        ["東和商事", 12000, "2026/03/10", 12000],
        ["北斗物産", 8000, "2026/03/12", 8000],
        ["南星工業", 5000, None, None],
        ["西山商店", None, "2026/03/15", 3000],
        ["中央電機", 7000, "2026/03/20", 7000],
    ]
    notes = _notes(grid)
    assert _stage(notes) == ["未記入の段: 支払（行 4 [南星工業]）"]
    blanks = _blanks(notes)
    assert "B5" in blanks                             # 組の外の空欄は今までどおり「表の中の空欄」
    assert "C4" not in blanks and "D4" not in blanks


def test_two_groups_in_one_table():
    grid = [
        ["番号", "承認日", "承認者", "支払日", "支払額"],
        [1, "4/1", "部長", "4/10", 100],
        [2, "4/2", "課長", None, None],
        [3, None, None, None, None],
        [4, "4/4", "部長", "4/12", 300],
        [5, "4/5", "課長", "4/15", 500],
    ]
    assert _stage(_notes(grid)) == ["未記入の段: 承認（行 4 [3]）", "未記入の段: 支払（行 3 [2]・4 [3]）"]


def test_no_stage_when_the_group_is_empty_throughout():
    grid = [
        ["番号", "品名", "返却日", "返却担当"],
        [1, "ノートPC", None, None],
        [2, "カメラ", None, None],
        [3, "マイク", None, None],
    ]
    assert _stage(_notes(grid)) == []


def test_no_stage_for_scattered_blanks():
    grid = [
        ["申請番号", "申請日", "承認日", "承認者"],
        ["A-01", "2/1", "2/2", "部長"],
        ["A-02", "2/3", None, "課長"],
        ["A-03", "2/5", "2/6", None],
        ["A-04", "2/7", "2/8", "部長"],
        ["A-05", "2/9", "2/10", "課長"],
    ]
    notes = _notes(grid)
    assert _stage(notes) == []
    blanks = _blanks(notes)
    assert "C3" in blanks and "D4" in blanks         # ばらばらの空欄は「表の中の空欄」のまま


def test_no_stage_when_the_head_word_has_one_column():
    grid = [
        ["番号", "品名", "貸出日", "返却日"],
        [1, "ノートPC", "1/5", "1/8"],
        [2, "カメラ", "1/6", None],
        [3, "マイク", "1/7", "1/10"],
        [4, "三脚", "1/8", "1/11"],
    ]
    notes = _notes(grid)
    assert _stage(notes) == []
    assert "D3" in _blanks(notes)


def test_no_stage_when_filled_rows_are_fewer_than_half():
    grid = [
        ["番号", "支払日", "支払額"],
        [1, "3/1", 100],
        [2, None, None],
        [3, None, None],
        [4, None, None],
        [5, "3/5", 500],
    ]
    assert _stage(_notes(grid)) == []


def test_head_prefix_groups():
    heads = ["番号", "返却日", "返却担当", "承認日", "承認者", "2024年", "2025年", "備考", None, 3]
    assert vv._head_prefix_groups(heads) == [("返却", [1, 2]), ("承認", [3, 4])]
    # 末尾の区切りは頭の語に入れない／3 列以上の組
    assert vv._head_prefix_groups(["支払・日", "支払・額", "支払・方法"]) == [("支払", [0, 1, 2])]
