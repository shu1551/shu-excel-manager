# -*- coding: utf-8 -*-
"""リボンの目録と、目録から足した基本の操作の棚 6 本（表の整理_基本操作）の試験（2026-10-28）。Excel にも AI にも触らない。

10/25: 頼む回数の一番多い基本の書式（「見出しの文字を大きくして赤にして」）が棚に当たらず AI に回っていた。
リボンの全コマンドを目録（fixtures/ribbon_catalog.tsv）にして 棚にする／もう棚にある／棚にしない に分け、
棚にする物を「どこを × どうする」の 6 本にまとめた。ここでは
  1. 目録の形（列・分け・行き先の棚が本当にある棚か）
  2. 問題集（fixtures/shelf_exam/entry_phrases_リボン.tsv）の全問が写し（shelf_pick）と seiri の入口（plan_full）で期待どおりか
  3. 6 本の頭が検査（shelf_lint）で止まらないか・一般の語 1 つで当たらないか
  4. 6 本が値を書かないか（書式・幅・絞り込み・設定・シートの名前だけ）
を見る。VBA の正本との突き合わせ（shelf --check）は Excel のある手元で。
"""
import io
import os
import re

import pytest

import vbam_prefire as vp

HERE = os.path.dirname(os.path.abspath(__file__))
CATALOG = os.path.join(HERE, '..', 'fixtures', 'ribbon_catalog.tsv')
EXAM = os.path.join(HERE, '..', 'fixtures', 'shelf_exam', 'entry_phrases_リボン.tsv')
MODULE = '表の整理_基本操作'
NEW_SHELVES = (
    '頼みの文で指した所の書式を変える',
    '頼みの文で指した列の表示形式を変える',
    '頼みの文で指した列と行の幅と表示を変える',
    '頼みの文で指した列で表を絞り込む',
    '頼みの文でシートの表示と印刷を変える',
    '頼みの文でシートの名前と色と並びを変える',
)
KINDS = ('棚にする', 'もう棚にある', '棚にしない')


def _macro_dir():
    """配っている棚の書き出し（手元は publish\\addin\\macros、公開リポは 作業ファイル\\project\\macros）で、新しい棚のある方。無ければ None"""
    for d in (os.path.join(HERE, '..', 'publish', 'addin', 'macros'), os.path.join(HERE, '..', 'macros')):
        if os.path.exists(os.path.join(d, MODULE + '.bas')):
            return d
    return None


def _shelf_text(d):
    text = ''
    for f in sorted(os.listdir(d)):
        if f.startswith('表の整理') and f.endswith('.bas'):
            with open(os.path.join(d, f), encoding='cp932', newline='') as fh:
                text += fh.read()
    return text


def _module_text(d):
    with open(os.path.join(d, MODULE + '.bas'), encoding='cp932', newline='') as fh:
        return fh.read()


def _catalog_rows():
    if not os.path.exists(CATALOG):
        pytest.skip('リボンの目録（fixtures/ribbon_catalog.tsv）が無い写し')
    with io.open(CATALOG, encoding='utf-8') as f:
        lines = [ln.rstrip('\n') for ln in f if not ln.startswith('#')]
    assert lines[0].split('\t') == ['コマンドID', 'リボンの場所', '日本語の名前', '分け', '行き先の棚']
    return [ln.split('\t') for ln in lines[1:] if ln]


@pytest.mark.core
def test_catalog_columns_and_kinds():
    """目録の各行は 5 列。分けは 3 つのどれか。棚にしない物は行き先が空、ほかは行き先がある。棚にする物の行き先は新しい 6 本"""
    rows = _catalog_rows()
    assert len(rows) > 900
    seen = set()
    for r in rows:
        assert len(r) == 5, r
        cid, place, name, kind, dest = r
        assert cid and ' > ' in place and name, r
        assert kind in KINDS, r
        assert (cid, place) not in seen, r
        seen.add((cid, place))
        if kind == '棚にしない':
            assert dest == '', r
        else:
            assert dest, r
        if kind == '棚にする':
            assert dest in NEW_SHELVES, r
    by_kind = {k: sum(1 for r in rows if r[3] == k) for k in KINDS}
    assert all(by_kind.values()), by_kind
    # 新しい 6 本のどれもが、目録の 1 行以上の行き先になっている
    assert {r[4] for r in rows if r[3] == '棚にする'} == set(NEW_SHELVES)


def test_catalog_destinations_are_real_shelves():
    """行き先の棚（／で 2 本まで）は、配っている棚に本当にある名前（書き間違いで当たらない行き先を作らない）"""
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しに 表の整理_基本操作 が無い（秀コンボに入れて書き出した後で見る）')
    names = {e['name'] for e in vp.shelf_entries_from_text(_shelf_text(d))}
    for r in _catalog_rows():
        for n in filter(None, r[4].split('／')):
            assert n in names, r


def test_ribbon_exam_matches_copy_and_seiri_entry():
    """問題集の全問が、写し（shelf_pick）と seiri の入口（plan_full）で期待どおり（期待が空＝AI に回す）"""
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しに 表の整理_基本操作 が無い')
    qa = _exam_rows()
    assert len(qa) >= 100
    text = _shelf_text(d)
    entries = vp.shelf_entries_from_text(text)
    reg = vp.registry_from_text(text)
    bad = []
    for e, q in qa:
        got = vp.shelf_pick(q, entries)
        pf = vp.plan_full(q, reg, entries=entries)
        seiri = pf['chosen'][0] if pf['chosen'] else ''
        if got != e or seiri != e:
            bad.append((q, e, got, seiri))
    assert not bad, bad[:10]
    # 新しい棚に当たる頼みも、当たってはいけない頼み（空や既存の棚）も入っている
    assert sum(1 for e, _q in qa if e in NEW_SHELVES) >= 60
    assert sum(1 for e, _q in qa if e not in NEW_SHELVES) >= 40


def _exam_rows():
    """この問題集の (期待, 頼みの文)。手元の作業場にほかの問題集があっても、この棚の期待で見る"""
    with io.open(EXAM, encoding='utf-8') as f:
        rows = [ln.rstrip('\n').split('\t') for ln in f if '\t' in ln]
    assert all(len(r) == 2 and r[1] for r in rows)
    assert len({r[1] for r in rows}) == len(rows)
    return [(r[0], r[1]) for r in rows]


def test_general_word_alone_does_not_hit_new_shelves():
    """一般の語 1 つ（赤・大きく・色・表・グラフ）や「どこを」の無い頼みでは新しい棚に当てない。「どこを」と「どうする」がそろったときだけ"""
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しに 表の整理_基本操作 が無い')
    entries = vp.shelf_entries_from_text(_shelf_text(d))
    for q in ('赤', '大きく', '色', '表', 'グラフ', '赤にして', '大きくして', '色を変えて', '文字を赤にして', '表を赤にして',
              '赤字の行を一覧にして', 'このシートの表を青にして', '見出しを消して', '見出しの色を教えて'):
        assert vp.shelf_pick(q, entries) not in NEW_SHELVES, q
    assert vp.shelf_pick('見出しの文字を大きくして赤にして', entries) == NEW_SHELVES[0]


def test_new_shelf_heads_pass_lint():
    """6 本の頭は既存の棚と同じ書き方（依頼の語・依頼の組・階級・扱う・列は頼みから）で、当たり方の検査で止まらず、注意も出ない"""
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しに 表の整理_基本操作 が無い')
    text = _module_text(d)
    ent = {e['name']: e for e in vp.shelf_entries_from_text(text)}
    assert set(NEW_SHELVES) <= set(ent)
    for n in NEW_SHELVES:
        assert ent[n]['ask'] and ent[n]['combo'] and ent[n]['tier'] == '基本', n
    assert set(NEW_SHELVES) <= vp.request_column_macros(text)
    reg = {e['name'] for e in vp.registry_from_text(text)}
    assert set(NEW_SHELVES) <= reg
    found = [x for x in vp.shelf_lint(text, base_rate_lines=[]) if x[0] in ('止める', '注意')]
    assert not found, found
    # 棚の目録に出るのは 6 本だけ（下請けは Private Function）
    assert re.findall(r'^Sub ([^\s(]+)\(', text, re.M) == list(NEW_SHELVES)
    assert 'Private Sub' not in text


def test_new_shelves_write_no_values():
    """6 本は値を書かない（既存の範囲を値で書き戻さない・文字のコードを数や日付に化けさせない）。経緯の日付と StatusBar の知らせがある"""
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しに 表の整理_基本操作 が無い')
    text = _module_text(d)
    code = '\n'.join(ln.split("'")[0] for ln in re.sub(r'"(?:[^"]|"")*"', '""', text).splitlines())
    assert not re.search(r'\.(?:Value2?|Formula2?|FormulaR1C1)\s*=', code)
    assert '2026-10-28' in text
    for m in re.finditer(r'^Sub ([^\s(]+)\(.*?^End Sub', text.replace('\r\n', '\n'), re.M | re.S):
        assert 'Application.StatusBar' in m.group(0), m.group(1)
        assert '先撃ちの頼みの文' in m.group(0) or '基_頼み()' in m.group(0), m.group(1)
    assert '\r\r\n' not in text and text.count('\r\n') == text.count('\n')
