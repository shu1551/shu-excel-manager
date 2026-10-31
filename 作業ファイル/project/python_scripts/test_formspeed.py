"""
099: form-speed の単体テストスイート
速度チューニング（プロパティ取得キャッシュ、VBA生成の高速性、前後の出力一貫性）の検証
"""
import pytest
import time
from unittest.mock import MagicMock
from vbam_form2vba import _F2V_DEFAULTS_CACHE, _f2v_defaults, _f2v_lit, _f2v_q
from vbam_formbuild import build_form_spec_from_columns


def test_defaults_cache_reuses():
    """既定値キャッシュが機能し、キャッシュが存在する場合は再利用される"""
    global _F2V_DEFAULTS_CACHE
    mock_wb = MagicMock()
    # キャッシュをセット
    import vbam_form2vba
    vbam_form2vba._F2V_DEFAULTS_CACHE = {"dummy": {"Left": 0}}
    res = _f2v_defaults(mock_wb)
    assert res == {"dummy": {"Left": 0}}
    # 新規フォームの追加が行われていないこと
    mock_wb.VBProject.VBComponents.Add.assert_not_called()


def test_form_to_vba_output_consistency():
    """リテラル変換・クォート処理の一貫性（前後で同一の結果を返す）"""
    assert _f2v_q('テスト"文字列"') == '"テスト""文字列"""'
    assert _f2v_lit(123) == "123"
    assert _f2v_lit(True) == "True"
    assert _f2v_lit("abc") == '"abc"'


def test_spec_build_speed_under_threshold():
    """100列の巨大表でも設計図生成が 0.05 秒以内で完了すること"""
    cols = [
        {"name": f"col_{i}", "header": f"ヘッダ_{i}", "control": "txt", "required": False, "multiline": False, "hint": None}
        for i in range(100)
    ]
    t0 = time.time()
    spec = build_form_spec_from_columns("F_100", "100列", "Sheet1", "T_100", cols, [])
    elapsed = time.time() - t0
    assert elapsed < 0.05, f"設計図生成が遅い: {elapsed:.4f}s"
    assert spec["name"] == "F_100"
    assert len(spec["rows"]) > 0
