# -*- coding: utf-8 -*-
"""棚を鍛える固定の形（vbam_train）: Excel にも AI にも触らない部分（2026-10-23）。

凍結・採点・仕分け・問題集へ・守り（検査を外した焼き）・前後の帳尻。決まりが道具の側で効くことを確かめる。
"""
import json
import os

import pytest

import vbam_train as vt


def _write_phrases(d, rows):
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, 'phrases.tsv'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(vt.TEMPLATE_NOTE + '\t'.join(vt.HEADER) + '\n')
        for r in rows:
            f.write('\t'.join(r) + '\n')


def _rows(n=12, held_every=3):
    """n 個の頼み。held_every 個に 1 個が見せない。型は 4 つ、最後の 2 割は AI。"""
    out = []
    for i in range(1, n + 1):
        held = '1' if i % held_every == 0 else '0'
        expect = 'AI' if i > n - 2 else '棚'
        out.append([str(i), f"型{i % 4}", f"頼み{i}をして", held, expect, ''])
    return out


# ---- 読む・凍結 ----
def test_load_phrases_reports_each_bad_row(tmp_path):
    d = str(tmp_path / 'r')
    _write_phrases(d, [['1', '型', '頼みA', '2', '棚', ''],       # 見せないが 2
                       ['2', '型', '頼みB', '0', '多分', ''],      # 期待が違う
                       ['3', '', '頼みC', '0', '棚', ''],          # 型が空
                       ['4', '型', '頼みD', '0', '棚', ''],
                       ['5', '型', '頼みD', '0', '棚', '']])       # 重複
    rows, errs = vt.load_phrases(os.path.join(d, 'phrases.tsv'))
    assert [r['text'] for r in rows] == ['頼みD']
    assert len(errs) == 4


def test_validate_needs_enough_rows_and_a_third_held(tmp_path):
    rows = [{'no': str(i), 'kind': 'k', 'text': f't{i}', 'held': False, 'expect': '棚', 'shelf': ''} for i in range(12)]
    errs, _w = vt.validate(rows)
    assert any('見せない頼み' in e for e in errs)            # 見せないが 0 個
    for r in rows[:4]:
        r['held'] = True
    errs, warns = vt.validate(rows)
    assert not errs
    assert any('期待が AI' in w for w in warns)
    assert vt.validate(rows[:5])[0][0].startswith('頼みが 5 個')


def test_freeze_then_any_edit_is_caught_but_pinning_a_shelf_is_allowed(tmp_path):
    d = str(tmp_path / 'r')
    _write_phrases(d, _rows())
    ok, out = vt.freeze(d)
    assert ok and 'run.json' in os.listdir(d)
    assert vt.frozen_ok(d) == (True, '')
    # 正解の棚の欄を埋めるのは仕分けの仕事＝許す
    rows, _e = vt.load_phrases(os.path.join(d, 'phrases.tsv'))
    rows[0]['shelf'] = '棚A'
    _write_phrases(d, [[r['no'], r['kind'], r['text'], '1' if r['held'] else '0', r['expect'], r['shelf']] for r in rows])
    assert vt.frozen_ok(d)[0]
    # 頼みの文を結果に合わせて書き換えたら止める
    rows[1]['text'] = '書き換えた頼み'
    _write_phrases(d, [[r['no'], r['kind'], r['text'], '1' if r['held'] else '0', r['expect'], r['shelf']] for r in rows])
    ok, why = vt.frozen_ok(d)
    assert not ok and '書き換わって' in why
    # 見せない・期待を入れ替えても同じ
    rows[1]['text'] = '頼み2をして'
    rows[1]['held'] = not rows[1]['held']
    _write_phrases(d, [[r['no'], r['kind'], r['text'], '1' if r['held'] else '0', r['expect'], r['shelf']] for r in rows])
    assert not vt.frozen_ok(d)[0]


def test_unfrozen_run_is_refused():
    assert not vt.frozen_ok(os.path.join(os.path.dirname(__file__), '_no_such_run_dir'))[0]


# ---- 採点 ----
@pytest.mark.parametrize("expect,pinned,got,cat", [
    ('AI', '', '', '正しくAIへ'),
    ('AI', '', '棚A', '誤爆'),
    ('棚', '棚A', '棚A', '正しく当たった'),
    ('棚', '棚A', '棚B', '誤爆'),
    ('棚', '棚A', '', '取りこぼし'),
    ('棚', '', '棚A', '未確定'),                    # 正解の棚が未記入＝当たりに数えない
    ('棚', '', '', '取りこぼし'),
    ('棚', vt.NONE_SHELF, '', '棚が無い'),
    ('棚', vt.NONE_SHELF, '棚A', '誤爆'),
])
def test_judge(expect, pinned, got, cat):
    assert vt.judge({'expect': expect, 'shelf': pinned}, got) == cat


def _answers(mapping, py=None):
    def fn(texts):
        vba = [mapping.get(t, '') for t in texts]
        return vba, [(py or {}).get(t, mapping.get(t, '')) for t in texts]
    return fn


def _res(rows_spec):
    rows = [{'no': str(i), 'kind': k, 'text': t, 'held': h, 'expect': e, 'shelf': s}
            for i, (k, t, h, e, s) in enumerate(rows_spec, 1)]
    return rows


def test_shoot_summarize_and_parity():
    rows = _res([('k', 'a', False, '棚', '棚A'), ('k', 'b', False, '棚', '棚A'), ('k', 'c', True, '棚', '棚A'),
                 ('m', 'd', True, 'AI', ''), ('m', 'e', False, 'AI', '')])
    res = vt.shoot(rows, _answers({'a': '棚A', 'c': '棚B', 'e': '棚A'}, py={'a': '棚X'}))
    cats = {r['text']: r['cat'] for r in res}
    assert cats == {'a': '正しく当たった', 'b': '取りこぼし', 'c': '誤爆', 'd': '正しくAIへ', 'e': '誤爆'}
    learn, held = vt.summarize(res, False), vt.summarize(res, True)
    assert learn['n'] == 3 and held['n'] == 2
    assert held['wrong'] == 1 and held['rate'] == 0.0
    assert vt.summarize(res)['parity'] == 1                  # a だけ VBA と写しが違う


def test_shoot_does_not_fall_back_to_python_when_vba_fails():
    def fn(_t):
        raise RuntimeError("VBA を呼べません")
    with pytest.raises(RuntimeError):
        vt.shoot(_res([('k', 'a', False, '棚', '棚A')]), fn)


def test_results_roundtrip(tmp_path):
    rows = _res([('k', 'a', False, '棚', '棚A'), ('m', 'b', True, 'AI', '')])
    res = vt.shoot(rows, _answers({'a': '棚A'}))
    p = str(tmp_path / 'before.tsv')
    vt.write_results(p, res)
    back = vt.read_results(p)
    assert [(r['text'], r['held'], r['got'], r['cat']) for r in back] == \
           [(r['text'], r['held'], r['got'], r['cat']) for r in res]
    assert vt.read_results(str(tmp_path / 'none.tsv')) is None


# ---- 仕分け ----
def test_sort_adds_words_only_for_types_missed_in_two_or_more_phrasings():
    rows = _res([('列を分ける', '列をカンマで分けて', False, '棚', '棚A'),
                 ('列を分ける', 'この列を区切って分けて', False, '棚', '棚A'),     # 同じ型 2 つの言い方で外れ
                 ('並べ替え', '大きい順に並べて', False, '棚', '棚B'),            # 1 つだけ
                 ('並べ替え', '小さい順に並べて', False, '棚', '棚B'),
                 ('見せない', '見せない頼み', True, '棚', '棚A'),                  # 見せないは材料にしない
                 ('誤爆の型', '質問です', False, 'AI', ''),
                 ('棚なし', '何かして', False, '棚', vt.NONE_SHELF)])
    res = vt.shoot(rows, _answers({'小さい順に並べて': '棚B', '大きい順に並べて': '棚B', '質問です': '棚Z'}))
    groups = {(g['kind'], g['cat']): g for g in vt.sort_rows(res)}
    assert groups[('列を分ける', '取りこぼし')]['verdict'] == '語を足す候補'
    assert ('並べ替え', '取りこぼし') not in groups                    # 当たっているので出ない
    assert all(g['kind'] != '見せない' for g in groups.values())
    assert groups[('誤爆の型', '誤爆')]['verdict'].startswith('待つ')
    assert groups[('棚なし', '棚が無い')]['verdict'].startswith('棚の候補')
    text = '\n'.join(vt.sort_text(list(groups.values())))
    assert '列をカンマで分けて' in text


def test_sort_does_not_merge_different_shelves_under_one_broad_kind():
    """型が広すぎる（別々の仕事を同じ型に入れた）とき、違う棚の外れを「2 つの言い方」と数えない（1 周目の穴）。"""
    rows = _res([('整える', '全角の数字を半角にして', False, '棚', '棚A'),
                 ('整える', '1行目の見出しを固定して', False, '棚', '棚B')])
    res = vt.shoot(rows, _answers({}))
    groups = vt.sort_rows(res)
    assert len(groups) == 2 and all(g['verdict'].startswith('待つ') for g in groups)


def test_sort_one_phrasing_is_not_enough():
    rows = _res([('型', '言い方1', False, '棚', '棚A'), ('型', '言い方2', False, '棚', '棚A')])
    res = vt.shoot(rows, _answers({'言い方2': '棚A'}))
    g = vt.sort_rows(res)[0]
    assert g['verdict'].startswith('待つ') and len(g['items']) == 1


# ---- 問題集へ ----
def test_export_only_what_is_already_right_and_says_why_for_the_rest(tmp_path):
    rows = _res([('k', '当たる頼み', False, '棚', '棚A'), ('k', '外れる頼み', False, '棚', '棚A'),
                 ('k', '未確定の頼み', False, '棚', ''), ('m', 'AIの頼み', False, 'AI', ''),
                 ('m', '見せない頼み', True, 'AI', ''), ('k', 'すでにある', False, 'AI', ''),
                 ('n', '棚が無い頼み', False, '棚', vt.NONE_SHELF)])
    res = vt.shoot(rows, _answers({'当たる頼み': '棚A', '未確定の頼み': '棚A'}))
    out = str(tmp_path / 'entry_phrases_訓練_x.tsv')
    n, skipped = vt.export_exam(res, out, existing_texts=['すでにある'])
    assert n == 2
    body = open(out, encoding='utf-8').read().splitlines()
    assert body[0].split('\t')[:2] == ['棚A', '当たる頼み']
    assert body[1].split('\t')[:2] == ['', 'AIの頼み']
    why = dict(skipped)
    assert why['外れる頼み'] == '取りこぼしのまま'
    assert why['未確定の頼み'] == '正解の棚が未確定'
    assert why['すでにある'] == '問題集にすでにある'
    assert why['棚が無い頼み'].startswith('棚が無い')
    assert '見せない頼み' not in why and '見せない頼み' not in open(out, encoding='utf-8').read()


def test_exam_questions_reads_the_exported_file(tmp_path):
    import vbam_prefire as vp
    rows = _res([('k', '当たる頼み', False, '棚', '棚A'), ('m', 'AIの頼み', False, 'AI', '')])
    res = vt.shoot(rows, _answers({'当たる頼み': '棚A'}))
    vt.export_exam(res, str(tmp_path / 'entry_phrases_訓練_t.tsv'))
    assert vp.exam_questions(str(tmp_path)) == [('棚A', '当たる頼み'), ('', 'AIの頼み')]


# ---- 守り ----
def test_no_check_burns_are_found_after_the_freeze(tmp_path):
    log = tmp_path / '_calls.jsonl'
    recs = [{'time': '2026-10-23 09:00:00', 'cmd': 'register-addin', 'args': '秀コンボ.xlsm --no-check'},   # 凍結の前
            {'time': '2026-10-23 11:00:00', 'cmd': 'register-addin', 'args': '秀コンボ.xlsm -y'},
            {'time': '2026-10-23 11:05:00', 'cmd': 'register-addin', 'args': '秀コンボ.xlsm --no-test'},
            {'time': '2026-10-23 11:10:00', 'cmd': 'shelf', 'args': '--no-check'},
            {'time': '2026-10-23 11:20:00', 'cmd': 'register-addin', 'args': '秀コンボ.xlsm --no-check -y'}]
    log.write_text('\n'.join(json.dumps(r, ensure_ascii=False) for r in recs) + '\nこわれた行\n', encoding='utf-8')
    got = vt.no_check_burns('2026-10-23 10:00:00', str(log))
    assert [t for t, _a in got] == ['2026-10-23 11:05:00', '2026-10-23 11:20:00']
    assert vt.no_check_burns('2026-10-23 10:00:00', str(tmp_path / 'none.jsonl')) == []


def test_call_log_records_no_check_flags():
    import vbam_core
    assert vbam_core._CALL_LOG_OPTS['no_check'] == 'no-check'
    assert vbam_core._CALL_LOG_OPTS['no_test'] == 'no-test'


# ---- 帳尻 ----
def _held_results(cats):
    rows = [{'no': str(i), 'kind': 'k', 'text': f't{i}', 'held': True, 'expect': '棚', 'shelf': '棚A'}
            for i in range(len(cats))]
    out = []
    for r, c in zip(rows, cats):
        got = {'正しく当たった': '棚A', '誤爆': '棚B', '取りこぼし': ''}[c]
        out.append(dict(r, got=got, py=got, cat=vt.judge(r, got)))
    return out


def test_close_stops_when_held_out_gets_worse():
    before = _held_results(['正しく当たった', '取りこぼし', '取りこぼし'])
    better = _held_results(['正しく当たった', '正しく当たった', '取りこぼし'])
    worse_wrong = _held_results(['正しく当たった', '誤爆', '取りこぼし'])
    worse_miss = _held_results(['取りこぼし', '取りこぼし', '取りこぼし'])
    assert vt.close_verdict(before, better)[0] == []
    assert any('誤爆が増え' in s for s in vt.close_verdict(before, worse_wrong)[0])
    stops = vt.close_verdict(before, worse_miss)[0]
    assert any('正しく当たった数が減' in s for s in stops) and any('取りこぼしが増' in s for s in stops)


def test_close_stops_on_parity_break():
    before = _held_results(['正しく当たった'])
    after = _held_results(['正しく当たった'])
    after[0]['py'] = '別の棚'
    assert any('食い違う' in s for s in vt.close_verdict(before, after)[0])
    # 学習用の頼みの食い違いも数える（見せない頼みだけを見ていた穴）
    after2 = _held_results(['正しく当たった'])
    after2[0].update(held=False, py='別の棚')
    assert any('食い違う' in s for s in vt.close_verdict(_held_results(['正しく当たった']), after2)[0])


def test_ledger_appends_with_header_once(tmp_path):
    root = str(tmp_path)
    res = _held_results(['正しく当たった', '誤爆'])
    s = vt.summarize(res)
    vt.ledger_append('r1', 'before', s, vt.summarize(res, True), root=root)
    vt.ledger_append('r1', 'close', s, vt.summarize(res, True), note='ok', root=root)
    lines = open(vt.ledger_path(root), encoding='utf-8').read().splitlines()
    assert len(lines) == 3 and lines[0].startswith('日時\t名前')
    assert lines[1].split('\t')[1:3] == ['r1', 'before']


# ---- 名前・入口 ----
def test_run_dir_rejects_unsafe_names(tmp_path):
    assert vt.run_dir('a1', root=str(tmp_path)).endswith('a1')
    for bad in ('', 'a b', '../x', 'a/b', 'a:b'):
        with pytest.raises(ValueError):
            vt.run_dir(bad, root=str(tmp_path))


def test_stale_note_only_when_the_baked_owner_differs(monkeypatch):
    import vbam_view
    a = "Sub 棚A()\n    ' 依頼の語: 甲|乙\n    ' 依頼の組: 結合,連結+氏名\nEnd Sub\n"
    a_other_body = "Sub 棚A()\n    ' 依頼の語: 甲|乙\n    ' 依頼の組: 結合,連結+氏名\n    Dim x As Long   ' 処理の中身は違ってよい\nEnd Sub\n"
    b = "Sub 棚A()\n    ' 依頼の語: 甲|乙|丙\n    ' 依頼の組: 結合,連結+氏名\nEnd Sub\n"
    monkeypatch.setattr(vbam_view, '_shelf_source', lambda xl: ('秀コンボ.xlam', a.split('\n')))
    assert '焼き直す' in vt.stale_note(None, '秀コンボ.xlsm', b)                # 依頼の語が違う＝届いていない
    assert vt.stale_note(None, '秀コンボ.xlsm', a_other_body) == ''            # 処理の中身・整形だけの違いでは出さない
    assert vt.stale_note(None, '秀コンボ.xlam', b) == ''                       # 持ち主そのものを見ている


def test_probe_needs_phrases(monkeypatch, tmp_path):
    class A:
        posargs = ['probe', 'lap1']
    assert vt.cmd_train(A()) is False


def test_train_without_args_prints_the_procedure(capsys):
    class A:
        posargs = []
    assert vt.cmd_train(A()) is True
    out = capsys.readouterr().out
    for stage in ('new', 'freeze', 'before', 'sort', 'export', 'after', 'close'):
        assert stage in out


def test_train_new_then_freeze_through_the_entry(tmp_path, monkeypatch):
    monkeypatch.setattr(vt, 'train_root', lambda: str(tmp_path))

    class A:
        posargs = ['new', 'lap1']
    assert vt.cmd_train(A()) is True
    assert vt.cmd_train(A()) is False                          # 2 回目は上書きしない
    _write_phrases(str(tmp_path / 'lap1'), _rows())
    A.posargs = ['freeze', 'lap1']
    assert vt.cmd_train(A()) is True
    A.posargs = ['close', 'lap1']
    assert vt.cmd_train(A()) is False                          # before / after が無ければ閉じない


def test_shelf_command_accepts_train_flag():
    import vba_manager
    p = vba_manager.build_parser() if hasattr(vba_manager, 'build_parser') else None
    if p is None:
        pytest.skip('パーサーの組み立て関数が無い')
    a = p.parse_args(['shelf', '--train', 'new', 'x'])
    assert a.train is True and a.posargs == ['new', 'x']
