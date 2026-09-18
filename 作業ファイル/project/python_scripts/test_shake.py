# -*- coding: utf-8 -*-
"""揺らし（vbam_shake・2026-09-11 夜）: 記録の返事を Python で書き換え、AI なしで再生する。

Excel には触らない（偽の記録と偽の module だけ）。本物の Excel での再生は agent --shake で撃つ。
"""
import os
import sys
import json
import types

import pytest


def _shake_log(tmp_path):
    """動画のお題（51 行・重複 26 行）の記録を小さく真似た種。往復 1 は文だけ（JSON なし）・往復 2 に手。"""
    prompt = ("使用範囲: A1:I16  16行 x 9列\n気づき:\n  重複行: 行9 = 行6（東海精機(株)・表記ゆれあり）／行10 = 行7（サクラ物流株式会社）／"
              "行13 = 行8（ブルーオーシャン(株)）／行16 = 行11（マウンテンビュー株式会社）（計 4 行。消すなら dedupe の手）\n"
              "見出し行の推定: 5 行目（会社名 担当者名 部署・役職 電話番号 メールアドレス 最終接触日 案件金額 ステータス）\n"
              "データの末尾: I16（使用範囲と一致）\n")
    reply = {"say": "一括で直します",
             "plan": [{"item": "重複行を削除", "state": "未"}, {"item": "書式をそろえる", "state": "未"}],
             "actions": [
                 {"op": "normalize", "range": "A6:A16", "rules": ["trim"], "overwrite": True},
                 {"op": "normalize", "range": "F6:F16", "rules": ["date"], "overwrite": True},
                 {"op": "format", "range": "F6:F16", "number_format": "yyyy/m/d"},
                 {"op": "format", "range": "G6:G16", "number_format": "#,##0"},
                 {"op": "format", "range": "A6:I16", "font": "游ゴシック", "size": 11, "plain": True, "bg": "none", "unbold": True},
                 {"op": "dedupe", "range": "A5:I16", "keys": ["会社名", "担当者名"]},
                 {"op": "tidy", "ranges": ["A5:I16"]}],
             "done": True, "report": "【できなかったこと】\nなし"}
    log = tmp_path / "seed.jsonl"
    rows = [{"meta": {"book": "messy.xlsx", "sheet": "顧客リスト", "mode": "sheet", "request": "重複行は削除してよい。書式をそろえて",
                      "ai": "claude-code", "model": "sonnet", "max_turns": 4}},
            {"turn": 1, "prompt": prompt, "reply": "まず方針を書きます。重複は 4 行。"},
            {"turn": 2, "prompt": prompt, "reply": json.dumps(reply, ensure_ascii=False), "sec": 20.0},
            {"turn": 2, "grade": True, "prompt": "採点", "reply": "{\"score\": 90}"}]
    with open(log, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    return log, reply


def test_shake_seed_takes_the_first_reply_with_hands_and_reads_the_numbers_20260911(tmp_path):
    """種は「最初の手のある返事」（文だけの往復 1 は飛ばす・採点係の行も飛ばす）。重複行・末尾・見出しは材料から読む。"""
    import vbam_shake as vs
    log, reply = _shake_log(tmp_path)
    seed = vs._seed_of(str(log))
    assert seed["turn"] == 2 and seed["reply"]["actions"] == reply["actions"]
    assert seed["dups"] == [(9, 6), (10, 7), (13, 8), (16, 11)]
    assert seed["data_end"] == 16 and seed["header_row"] == 5
    assert seed["headers"][:4] == ["会社名", "担当者名", "部署・役職", "電話番号"]
    assert seed["sheet"] == "顧客リスト" and "重複行は削除してよい" in seed["request"]
    bad = tmp_path / "macro.jsonl"
    bad.write_text(json.dumps({"meta": {"mode": "macro"}}) + "\n", encoding="utf-8")
    with pytest.raises(ValueError, match="sheet"):
        vs._seed_of(str(bad))


def test_shake_variants_have_the_shapes_seen_on_20260911(tmp_path):
    """揺らした版が、実射で見た形そのものになっている（原本は先頭・当てはまらない型は飛ぶ）。"""
    import vbam_shake as vs
    log, reply = _shake_log(tmp_path)
    seed = vs._seed_of(str(log))
    made = vs.make_variants(seed)
    by = {v["name"]: v for v in made}
    names = [v["name"] for v in made]
    assert names[0] == "原本" and by["原本"]["turns"][0]["actions"] == reply["actions"]
    assert len(names) == len(set(names)) and len(names) >= 20

    def ops(n):
        return [a["op"] for a in by[n]["turns"][0]["actions"]]

    def dd(n):
        return [a for a in by[n]["turns"][0]["actions"] if a["op"] == "dedupe"][0]

    def rd(n, key):
        return [a[key] for a in by[n]["turns"][0]["actions"] if a["op"] == "row_delete"]

    # 消す手の形 5 通り（dedupe が row_delete に置き換わる・rows は重複行だけ）
    assert rd("消す手が飛び飛び", "rows") == [[9, 10, 13, 16]] and "dedupe" not in ops("消す手が飛び飛び")
    assert rd("消す手が1行ずつ上から", "at") == [9, 10, 13, 16]
    assert rd("消す手が1行ずつ下から", "at") == [16, 13, 10, 9]
    assert rd("消す手が範囲の文字", "range") == ["16:16", "13:13", "10:10", "9:9"]
    assert rd("消す手が連番のかたまり", "rows") == [[16], [13], [9, 10]]
    # 消した後の番地: 末尾 16 − 重複 4 ＝ 12 で終わる（dedupe の範囲は触らない・tidy も縮む）
    pd = {a["op"] + ":" + str(a.get("range") or a.get("ranges")) for a in by["消した後の番地"]["turns"][0]["actions"]}
    assert {"normalize:A6:A12", "format:A6:I12", "dedupe:A5:I16", "tidy:['A5:I12']"} <= pd
    a138 = by["138秒の形"]["turns"][0]["actions"]
    assert a138[-2]["op"] == "row_delete" and a138[-2]["rows"] == [9, 10, 13, 16] and a138[0]["range"] == "A6:A12"
    assert [a["range"] for a in by["範囲が300行まで"]["turns"][0]["actions"] if a["op"] == "normalize"] == ["A6:A300", "F6:F300"]
    fmt = [a for a in by["書体と大きさを落とす"]["turns"][0]["actions"] if a.get("plain")][0]
    assert "font" not in fmt and "size" not in fmt and fmt["bg"] == "none"
    assert "yyyy/m/d" not in str(by["日付の表示形式を当てない"]["turns"][0]["actions"])
    assert "#,##0" not in str(by["金額の表示形式を当てない"]["turns"][0]["actions"])
    assert not any("overwrite" in a for a in by["overwriteを書かない"]["turns"][0]["actions"])
    assert dd("keysが列の字")["keys"] == ["A", "B"] and "keys" not in dd("keysなし")
    assert "keysが見出しの名前" not in by            # 種がもう見出しの名前＝当てはまらない型は飛ぶ
    assert dd("dedupeが見出しの下から")["range"] == "A6:I16"
    assert "tidy" not in ops("tidyなし") and ops("手の順が逆") == ops("原本")[::-1]
    assert ops("dedupeが先頭")[0] == "dedupe" and ops("dedupeが先頭")[1:] == [o for o in ops("原本") if o != "dedupe"]
    t1, t2 = by["doneを次の往復で"]["turns"]
    assert t1["done"] is False and t1["actions"] == reply["actions"]
    assert t2["done"] is True and t2["actions"] == [] and all(p["state"] == "済" for p in t2["plan"])
    # --only で絞っても原本は残る（並びは _VARIANTS の順）
    assert [v["name"] for v in vs.make_variants(seed, only="tidyなし,138秒の形")] == ["原本", "138秒の形", "tidyなし"]
    # 種が row_delete しか持たない記録（dedupe 無し）では、dedupe を置き換える型は全部飛ぶ
    seed2 = dict(seed, reply=dict(seed["reply"], actions=[{"op": "row_delete", "rows": [9, 10], "overwrite": True},
                                                          {"op": "tidy", "ranges": ["A5:I16"]}]))
    n2 = [v["name"] for v in vs.make_variants(seed2)]
    assert "消す手が飛び飛び" not in n2 and "keysなし" not in n2 and "消した後の番地" in n2 and "原本" in n2


def test_shake_variants_replay_through_replay_ask_and_dry_run_touches_no_excel(tmp_path, monkeypatch, capsys):
    """書き出した版は --replay と同じ読み手（_replay_ask）で順に返る。--dry-run は Excel に触らない。"""
    import vbam_agent as va
    import vbam_shake as vs
    import vbam_core
    log, reply = _shake_log(tmp_path)
    seed = vs._seed_of(str(log))
    v = [x for x in vs.make_variants(seed) if x["name"] == "doneを次の往復で"][0]
    path = tmp_path / "v.jsonl"
    path.write_text(vs._variant_jsonl(seed["meta"], v["turns"]), encoding="utf-8")
    ask = va._replay_ask(str(path))
    r1 = json.loads(ask(None, None, None, None)[0])
    r2 = json.loads(ask(None, None, None, None)[0])
    assert r1["done"] is False and r1["actions"] == reply["actions"] and r2["done"] is True
    with pytest.raises(RuntimeError, match="記録が尽きました"):
        ask(None, None, None, None)
    meta = json.loads(path.read_text(encoding="utf-8").splitlines()[0])["meta"]
    assert meta["mode"] == "sheet" and meta["model"] == "replay"      # _cmd_agent_replay と同じ形

    def boom(*a, **k):
        raise AssertionError("dry-run で Excel に触った")
    monkeypatch.setattr(vbam_core, "get_workbook", boom)
    monkeypatch.setattr(vs, "_AGENT_SHAKE_DIR", str(tmp_path / "shake"))
    assert vs.shake(str(log), None, dry_run=True) is True
    out = capsys.readouterr().out
    assert "揺らした版" in out and "138秒の形" in out and "Excel には触っていません" in out
    stamp = os.listdir(str(tmp_path / "shake"))[0]
    made = os.listdir(str(tmp_path / "shake" / stamp))
    assert any(n.startswith("00_原本") for n in made) and any(n.endswith("138秒の形.jsonl") for n in made)
    assert vs.shake(str(tmp_path / "none.jsonl"), None, dry_run=True) is False


def test_shake_diverts_state_paths_to_a_scratch_dir_20260911(tmp_path):
    """再生の子プロセスは、SCRIPT_DIR 配下のパス定数と関数の既定引数を丸ごと逃がす（本物の台帳・覚書に書かない）。"""
    import vbam_shake as vs
    from vbam_core import SCRIPT_DIR
    fake = types.ModuleType("vbam_fakestate")
    fake.BACKUP_DIR = os.path.join(SCRIPT_DIR, "..", "..", "backups")
    fake._X_FILE = os.path.join(SCRIPT_DIR, "_x.jsonl")
    fake.OTHER = "keep me"

    def f(a, path=os.path.join(SCRIPT_DIR, "_y.json"), n=3):
        return path
    fake.f = f
    sys.modules["vbam_fakestate"] = fake
    try:
        state = str(tmp_path / "state")
        backups = vs._divert_state(state, modules=["vbam_fakestate"])
        assert backups == os.path.join(state, "backups") and os.path.isdir(backups)
        assert fake.BACKUP_DIR == backups
        assert fake._X_FILE == os.path.join(state, "_x.jsonl") and fake.OTHER == "keep me"
        assert fake.f(0) == os.path.join(state, "_y.json") and fake.f.__defaults__[1] == 3
    finally:
        sys.modules.pop("vbam_fakestate", None)


def test_shake_judge_compare_and_reason_20260911():
    """判定の言葉・原本との比べ方・理由の 1 行。"""
    import vbam_shake as vs
    base = {"name": "原本", "ok": True, "turns": 1, "vals": [["a", "1"], ["b", "2"]]}
    assert vs._judge(base) == "一発合格" and vs._compare(base, base) == "同じ"
    assert vs._judge({"ok": True, "turns": 2}) == "合格（往復 2）"
    assert vs._judge({"ok": False, "turns": 1}) == "不合格"
    assert vs._judge({"err": "記録が尽きました（往復 1）＝…"}) == "差し戻し→尽きた"
    assert vs._judge({"err": "COM error"}) == "落ちた"
    assert vs._compare(base, {"vals": [["a", "1"], ["b", "3"]]}) == "1 セル違う"
    assert vs._compare(base, {"vals": [["a", "1"]]}).startswith("行 2→1")
    assert vs._compare(None, base) == "-" and vs._compare(base, {"err": "x"}) == "-"
    assert vs._reason({"back": "\n仕上げ検査で 2 件。\n・罫線が無い\n"}) == "仕上げ検査で 2 件。"
    # 「同じ返事の中に通らない手がある」は総論＝飛ばして、本当の理由（row の at は数）を拾う
    back = ("【結果】\n--- 1: normalize （失敗）---\n実行していません: 同じ返事の中に通らない手があるので、1 手も実行していません\n"
            "--- 11: row_delete （失敗）---\n実行していません: row の at は行番号（数）です\n")
    assert vs._reason({"back": back}) == "row の at は行番号（数）です"
    assert vs._reason({"back": "【前回の結果】\n--- 1: normalize ---\n8 セル\n手は通りましたが、仕上げ検査で 3 件"}).startswith("手は通りましたが、仕上げ検査で")
    assert vs._reason({"gates": {"audit": 1, "inv": 2}}).startswith("関所: 仕上げ検査×1")
    assert vs._reason({"audit": ["列幅が足りない"]}) == "列幅が足りない"


def test_shake_imagine_turns_an_ai_reply_into_variants_20260911(tmp_path, monkeypatch, capsys):
    """AI に書かせた返事（JSON の variants）を版にする。actions の無い物は捨てる・done:false は次の往復で done を足す。
    AI は 1 往復だけ（偽の _ask で数える）。dry-run でも AI の版が一覧に出る。"""
    import vbam_agent as va
    import vbam_shake as vs
    import vbam_core
    log, reply = _shake_log(tmp_path)
    seed = vs._seed_of(str(log))
    calls = []
    fake = {"variants": [
        {"why": "消す手を行番号で", "reply": {"say": "やります", "actions": [
            {"op": "row_delete", "rows": [9, 10, 13, 16], "overwrite": True}, {"op": "tidy", "ranges": ["A5:I12"]}],
            "done": True}},
        {"why": "手だけ返して done は次", "reply": {"actions": [{"op": "tidy", "ranges": ["A5:I16"]}], "done": False}},
        {"why": "手が無い", "reply": {"actions": [], "done": True}},
        {"why": "形が違う", "reply": "文字"}]}

    def ask(ai, model, key, history):
        calls.append((ai, model, key, history[-1][1]))
        return json.dumps(fake, ensure_ascii=False), {"in": 100, "out": 50}, 1.5
    monkeypatch.setattr(va, "_ask", ask)
    out, bill, raw = vs.ai_variants(seed, 4)
    assert len(calls) == 1 and calls[0][0] == "claude-code" and calls[0][1] == "haiku"
    assert "4 通り" in calls[0][3] and "重複行: 行9 = 行6" in calls[0][3] and '"dedupe"' in calls[0][3]
    assert [v["name"] for v in out] == ["AI案1", "AI案2"]
    assert out[0]["why"] == "消す手を行番号で" and out[0]["turns"][0]["done"] is True
    assert out[0]["turns"][0]["plan"] == reply["plan"]            # plan が無ければ種の plan を足す
    assert len(out[1]["turns"]) == 2 and out[1]["turns"][1]["done"] is True and out[1]["turns"][1]["actions"] == []
    assert bill.startswith("揺れ作り:") and "haiku" in bill and "100" in bill and raw.startswith("{")

    def boom(*a, **k):
        raise AssertionError("dry-run で Excel に触った")
    monkeypatch.setattr(vbam_core, "get_workbook", boom)
    monkeypatch.setattr(vs, "_AGENT_SHAKE_DIR", str(tmp_path / "shake"))
    assert vs.shake(str(log), None, dry_run=True, only="tidyなし", imagine=4) is True
    text = capsys.readouterr().out
    assert "揺れ作り:" in text and "AI案1" in text and "AI案2" in text and "AI案3" not in text
    stamp = os.listdir(str(tmp_path / "shake"))[0]
    made = os.listdir(str(tmp_path / "shake" / stamp))
    assert "_imagine_reply.txt" in made and any(n.endswith("AI案1.jsonl") for n in made)
    # AI が JSON を返さなければ版はゼロ（落ちない）
    monkeypatch.setattr(va, "_ask", lambda *a, **k: ("ごめんなさい、書けません", {}, 0.1))
    out, bill, raw = vs.ai_variants(seed, 2)
    assert out == [] and "返り" in bill


def test_shake_cli_and_mcp_wiring_20260911():
    """agent --shake [記録] --only … が読める。版の型は 22 通りで原本が先頭。"""
    import vba_manager as vm
    import vbam_shake as vs
    ns = vm.build_parser().parse_args(["agent", "--shake", "--dry-run"])
    assert ns.shake == "" and ns.only is None and ns.dry_run is True
    ns = vm.build_parser().parse_args(["agent", "--shake", r"C:\x\rec.jsonl", "--only", "tidyなし,138秒の形"])
    assert ns.shake == r"C:\x\rec.jsonl" and ns.only == "tidyなし,138秒の形"
    ns = vm.build_parser().parse_args(["agent", "直して"])
    assert ns.shake is None and ns.imagine == 0
    ns = vm.build_parser().parse_args(["agent", "--shake", "--imagine", "8", "--model", "haiku"])
    assert ns.imagine == 8 and ns.model == "haiku"
    assert [n for n, *_ in vs._VARIANTS][0] == "原本" and len(vs._VARIANTS) == 22
    assert len({n for n, *_ in vs._VARIANTS}) == 22
