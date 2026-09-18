"""pytest 共通設定。

テストは2本立て:
  純ロジック（既定）  py -m pytest -q
  実機E2E（明示）     py -m pytest --run-e2e -q

E2E は本物の Excel / VBE を起こすため、既定では走らせない（オプトイン）。
CI や普段の確認は純ロジックだけで 1〜2 秒で終わる。
"""
import importlib

import pytest

# sync_tools.py が退避した旧版コピー（_sync_backup/～/test_tools.py 等）を
# pytest が再帰収集すると、本物の test_tools.py と import 名が衝突して
# collection error になり、テストが1件も走らない（公開リポ側で実測）。
# 退避フォルダは収集対象から外す
collect_ignore_glob = ["_sync_backup*"]


def pytest_addoption(parser):
    parser.addoption(
        "--run-e2e", action="store_true", default=False,
        help="実機の Excel/VBE を使う E2E テストも実行する（既定はスキップ）")


def pytest_configure(config):
    config.addinivalue_line(
        "markers", "e2e: 実機の Excel/VBE を起こすテスト（--run-e2e で有効）")


# 道具が自分で書く状態ファイル → テンポラリの名前。定数を持っている module 全部に当てる
# （2026-09-11 に vbam_agent を分割したので、定数の持ち主が module ごとに違う。vbam_agent 側の
#   再 export の写しにも当てないと、テストの `va.定数` が本物のパスのまま残る）
_TMP_STATE = {
    "_AGENT_PROGRESS_FILE": "_agent_progress.json",     # 進み具合（2026-09-05）
    "_AGENT_RUNS_FILE": "_agent_runs.jsonl",            # 本番の走行台帳（2026-09-06）
    "_FIRED_FILE": "_fired.jsonl",                      # 先撃ちで撃ったマクロ（2026-09-18 夜）
    "_AGENT_LOGS_DIR": "_agent_logs",
    "_AGENT_CASES_FILE": "_agent_cases.json",           # 本番から拾った弾の台帳
    "_LAST_AGENT_UNDO_FILE": "_last_agent_undo.json",   # 控えの覚書（--undo が読む・2026-09-10 夜）
    "_JOB_CLOCK_FILE": "_job_clock.json",               # 仕事の時計
    "_CALL_LOG_FILE": "_calls.jsonl",                   # 呼び出し台帳（2026-09-16）
    "_CMD_PROGRESS_FILE": "_cmd_progress.json",         # コマンドの進み具合（2026-09-16・status が読む）
    "_AGENT_PROGRESS_PATH": "_agent_progress.json",     # 同上の agent 側の写し（vbam_core が AI を import しないための別名）
    # 2026-09-11: 分割の途中で覚書（_agent_notes.json）に偽の行が書かれた＝道具が書く物は全部ここに並べる
    "_AGENT_NOTES_FILE": "_agent_notes.json",           # ブック・シートごとの覚書
    "_AGENT_AFTER_DIR": "_agent_after",                 # 走行の終わりの姿
    "_AGENT_SCORE_FILE": "_agent_score.jsonl",          # 実射の点数
    "_AGENT_CASES_DIR": "_agent_cases",                 # 弾の練習台
    "_FIRE_OUT_FILE": "_fire_last.txt",                 # 実射の画面出力
    "_LAST_AGENT_LOG_FILE": "_last_agent_log.jsonl",    # 往復の全記録
    "_LAST_AGENT_ASK_FILE": "_last_agent_ask.txt",      # 1 通目
    "_LAST_AGENT_CHANGES_FILE": "_last_agent_changes.tsv",
    "_LAST_AGENT_PNG": "_last_agent.png",
    "_LAST_AGENT_VIEW_PNG": "_last_agent_view.png",
    "_AGENT_LOCK_FILE": "_agent_running.lock",
    "_KEY_STORE": "keys.json",                          # API キーの金庫
    "_LAST_VALUES_FILE": "_last_values.tsv",            # write_grid の受け渡し
    "_AGENT_SHAKE_DIR": "_agent_shake",                 # 揺らした版と再生の報告（2026-09-11 夜）
}
_STATE_MODULES = ("vbam_core", "vbam_agent", "vbam_keys", "vbam_ai", "vbam_hands", "vbam_inv", "vbam_undo",
                  "vbam_grade", "vbam_ledger", "vbam_fire", "vbam_macro", "vbam_clean", "vbam_shake",
                  "vbam_devtools")


@pytest.fixture(autouse=True)
def _agent_progress_to_tmp(tmp_path, monkeypatch):
    """道具が自分で書く状態ファイルは、テストから本物を汚さない（2026-09-05／2026-09-06 に台帳を追加）。

    run_agent／run_build は往復ごとに進み具合を上書きし、終わりに走行台帳へ 1 行足す。
    テストは偽のブックでそれを回すので、pytest を走らせるたびに「Book1!名簿 / 使い捨て /
    何もしていない」という偽の行が本物のファイルに残り、agent_status がそれを最新の仕事として
    返していた（実際に 9/5 に「実射の痕跡」と読み違えた）。台帳は「実運用の量」を測る数字なので、
    偽の走行が 1 行でも混ざるとその数字自体が嘘になる。全テストの既定でテンポラリへ逃がす。
    控えの覚書も同じ（run_build を回すテストが本物の覚書に stale の印を付けていた・2026-09-10 夜）。
    """
    for name in _STATE_MODULES:
        try:
            mod = importlib.import_module(name)
        except Exception:
            continue
        for attr, fname in _TMP_STATE.items():
            if hasattr(mod, attr):
                monkeypatch.setattr(mod, attr, str(tmp_path / fname), raising=False)


def _refuse_real_ai(*_a, **_k):
    raise RuntimeError("テストが本物の AI を呼ぼうとしました（_ask／_cc_oneshot を差し替えていないテスト）。"
                       "2026-09-11 の分割作業で 1 本が本物の Claude Code を呼び、使用量を無駄にした＝この見張りで止める")


@pytest.fixture(autouse=True)
def _no_real_ai_in_tests(monkeypatch):
    """テストから本物の AI（Gemini／Claude／Claude Code）を呼ばせない。

    偽の返事を使うテストは自分で `monkeypatch.setattr(va, "_ask", fake)` するので、それが上書きになる。
    差し替えを忘れたテストは、金を使う前にここで落ちる。
    """
    for name in ("vbam_agent", "vbam_ai"):
        try:
            mod = importlib.import_module(name)
        except Exception:
            continue
        for attr in ("_ask", "_cc_oneshot"):
            if hasattr(mod, attr):
                monkeypatch.setattr(mod, attr, _refuse_real_ai)


def pytest_collection_modifyitems(config, items):
    if config.getoption("--run-e2e"):
        return
    skip_e2e = pytest.mark.skip(reason="実機E2E（走らせるには --run-e2e）")
    for item in items:
        if "e2e" in item.keywords:
            item.add_marker(skip_e2e)
