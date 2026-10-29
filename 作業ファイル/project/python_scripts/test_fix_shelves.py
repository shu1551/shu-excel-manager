# -*- coding: utf-8 -*-
"""見つけるだけの棚が見つける物を直す棚 3 本（表の整理_作る の末尾・2026-10-28）の試験。Excel にも AI にも触らない。

数式の危ない所を右に報告する（合計のずれ・ならびの違う式）と 列番号を直書きした式を一覧にする は報告するだけだった。
同じ物を式のまま直す棚を 3 本足した:
  合計の式の範囲を表の終わりまで広げる        … =SUM(D2:D8) を明細の終わりの =SUM(D2:D10) に
  下へ写した式の絶対参照の抜けを直す          … 同じ列の多数が =D2*$G$1 なのに 1 本だけ =D10*G9 → =D10*$G$1
  VLOOKUPの列番号を見出しから引く形にする     … =VLOOKUP(A2,表,3,FALSE) の 3 を MATCH("見出し",表の見出し,0) に
ここでは
  1. 3 本の頭（依頼の語・依頼の組・扱う・見出し・形・変えてよい: 数式）が既存の棚の書き方どおりに読めるか
  2. 当たり方の検査（shelf_lint）で止める・注意が 0 か
  3. 問題集（fixtures/shelf_exam/entry_phrases_直す棚.tsv）の全問が写し（shelf_pick）と seiri の入口（plan_full）で期待どおりか
  4. 一般の語 1 つ・ほかの問題集の頼みで 3 本に当たらないか
  5. 3 本が式のセルだけに書くか（書くのは 変えた式だけ書く だけ・値は書かない）
を見る。VBA の正本との突き合わせ（shelf --check）と実際に直るかは Excel のある手元で。
"""
import io
import os
import re

import pytest

import vbam_prefire as vp

HERE = os.path.dirname(os.path.abspath(__file__))
EXAM = os.path.join(HERE, '..', 'fixtures', 'shelf_exam', 'entry_phrases_直す棚.tsv')
MODULE = '表の整理_作る'
NEW_SHELVES = (
    '合計の式の範囲を表の終わりまで広げる',
    '下へ写した式の絶対参照の抜けを直す',
    'VLOOKUPの列番号を見出しから引く形にする',
)
BLOCK_MARK = "' ---- 見つけるだけの棚が見つける物を直す棚 3 本"
WRITER = '変えた式だけ書く'


def _macro_dir():
    """配っている棚の書き出し（手元は publish\\addin\\macros、公開リポは 作業ファイル\\project\\macros）で、新しい棚のある方。無ければ None"""
    for d in (os.path.join(HERE, '..', 'publish', 'addin', 'macros'), os.path.join(HERE, '..', 'macros')):
        p = os.path.join(d, MODULE + '.bas')
        if os.path.exists(p):
            with open(p, encoding='cp932', newline='') as fh:
                if NEW_SHELVES[0] in fh.read():
                    return d
    return None


def _need_dir():
    d = _macro_dir()
    if d is None:
        pytest.skip('棚の書き出しの 表の整理_作る に直す棚 3 本が無い（秀コンボに入れて書き出した後で見る）')
    return d


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


def _block(d):
    """表の整理_作る の末尾に足した 3 本と手伝いの本文（目印の行から終わりまで）"""
    text = _module_text(d)
    p = text.find(BLOCK_MARK)
    assert p >= 0
    return text[p:]


def _procs(text):
    """本文 → {名前: プロシージャの本文}（Sub と Function）"""
    t = text.replace('\r\n', '\n')
    return {m.group(2): m.group(0) for m in re.finditer(
        r'^(?:Private\s+)?(Sub|Function)\s+([^\s(]+)\(.*?^End \1', t, re.M | re.S)}


def _exam_rows():
    """この問題集の (期待, 頼みの文)。手元の作業場にほかの問題集があっても、この棚の期待で見る"""
    with io.open(EXAM, encoding='utf-8') as f:
        rows = [ln.rstrip('\n').split('\t') for ln in f if '\t' in ln and not ln.startswith('#')]
    assert all(len(r) == 2 and r[1] for r in rows)
    assert len({r[1] for r in rows}) == len(rows)
    return [(r[0], r[1]) for r in rows]


def test_new_shelf_heads_follow_existing_shelves():
    """3 本の頭は Sub の下の 6 行に 依頼の語・依頼の組・扱う・見出し・形・変えてよい: 数式（既存の棚と同じ書き方）。登録簿にも先撃ちの並びにも入る"""
    d = _need_dir()
    text = _module_text(d)
    ent = {e['name']: e for e in vp.shelf_entries_from_text(text)}
    reg = {e['name']: e for e in vp.registry_from_text(text)}
    lines = text.replace('\r\n', '\n').split('\n')
    for n in NEW_SHELVES:
        assert n in ent and n in reg, n
        assert ent[n]['ask'] and ent[n]['combo'], n
        assert ent[n]['allow_change'] == '数式', n
        assert reg[n]['handles'] and reg[n]['headers'] == [], n
        i = lines.index(f'Sub {n}()')
        head = [ln.strip() for ln in lines[i + 1:i + 7]]
        for key in ('依頼の語:', '依頼の組:', '扱う:', '見出し:', '形:', '変えてよい: 数式'):
            assert any(h.startswith("' " + key) for h in head), (n, key)
        # 引数なしの入口（引数のある Sub は目録とマクロ一覧に出ない）
        assert re.search(rf'^Sub {re.escape(n)}\(\)\r$', text, re.M), n


def test_new_shelves_pass_lint():
    """足した本文だけを検査して止める・注意が 0。ブックの棚の全部でも、3 本の名前の出る止める・注意が無い"""
    d = _need_dir()
    found = [x for x in vp.shelf_lint(_block(d), base_rate_lines=[]) if x[0] in ('止める', '注意')]
    assert not found, found
    whole = [x for x in vp.shelf_lint(_shelf_text(d)) if x[0] in ('止める', '注意') and any(n in x[1] for n in NEW_SHELVES)]
    assert not whole, whole
    # 棚の目録に出るのは 3 本だけ（手伝いは Private Function）
    # 3 本の後ろに足した棚（082の2 の背景色の棚など）は、頭の語を持つ棚なら許す（手伝いが Public になっていないかだけ見る）
    names = re.findall(r'^(?:Public\s+)?Sub ([^\s(]+)\(', _block(d), re.M)
    assert names[:len(NEW_SHELVES)] == list(NEW_SHELVES)
    procs = _procs(_block(d))
    for extra in names[len(NEW_SHELVES):]:
        assert "' 依頼の語:" in procs.get(extra, ''), extra
    assert 'Private Sub' not in _block(d)


def test_fix_exam_matches_copy_and_seiri_entry():
    """問題集の全問が、写し（shelf_pick）と seiri の入口（plan_full）で期待どおり（期待が空＝AI に回す）"""
    d = _need_dir()
    qa = _exam_rows()
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
    # 3 本のどれにも当たる頼みが 10 以上、当たってはいけない頼み（空や既存の棚）が 10 以上
    for n in NEW_SHELVES:
        assert sum(1 for e, _q in qa if e == n) >= 10, n
    assert sum(1 for e, _q in qa if e not in NEW_SHELVES) >= 10


def test_general_word_alone_does_not_hit_new_shelves():
    """一般の語 1 つや、直す物の語の無い頼みでは 3 本に当てない"""
    d = _need_dir()
    entries = vp.shelf_entries_from_text(_shelf_text(d))
    for q in ('直して', '修正して', '合計', '合計して', '範囲', '範囲を広げて', '$', '$を付けて', 'VLOOKUP', 'VLOOKUPを直して',
              'MATCH', '列番号', '式を直して', '数式を直して', 'そろえて', 'この表を直して', '見出しを直して'):
        assert vp.shelf_pick(q, entries) not in NEW_SHELVES, q


def test_other_exams_do_not_hit_new_shelves():
    """ほかの問題集（shelf_exam の全部）で、期待が 3 本でない頼みは 3 本に当たらない（既存の棚の頼みを横取りしない）"""
    d = _need_dir()
    qa = vp.exam_questions()
    if not qa:
        pytest.skip('問題集が無い写し')
    entries = vp.shelf_entries_from_text(_shelf_text(d))
    stolen = [(e, q, got) for e, q in qa if e not in NEW_SHELVES
              for got in [vp.shelf_pick(q, entries)] if got in NEW_SHELVES]
    assert not stolen, stolen[:10]


def test_new_shelves_write_formula_cells_only():
    """3 本が書くのは式のセルだけ: 式を書く行は 変えた式だけ書く の中だけで、その前に HasFormula を見る。値（.Value・.Value2）は書かない。
    書き戻しの行き先は 3 本とも 変えた式だけ書く。.bas は CP932・CRLF（改行の二重化なし）"""
    d = _need_dir()
    block = _block(d)
    code = '\n'.join(ln.split("'")[0] for ln in re.sub(r'"(?:[^"]|"")*"', '""', block).splitlines())
    procs = _procs(code)
    write_re = re.compile(r'\.(?:Value2?|Formula2?|FormulaR1C1|Formula2R1C1|FormulaLocal|FormulaR1C1Local)\s*=(?!=)')
    for name, body in procs.items():
        for ln in body.split('\n'):
            s = ln.strip()
            # 比べる If（If cel.Formula = 新式 Then）は書かない
            if re.match(r'^(?:Else)?If\b', s) and not re.search(r'\bThen\s+\S', s):
                continue
            if write_re.search(s):
                assert name == WRITER, (name, s)
                assert not re.search(r'\.Value2?\s*=', s), (name, s)
    w = procs[WRITER]
    assert w.index('HasFormula') < w.index('.Formula =') and w.index('HasFormula') < w.index('.FormulaR1C1 =')
    for n in NEW_SHELVES:
        assert WRITER + '(' in procs[n], n
        assert 'Application.StatusBar' in procs[n] and '先撃ちに断る' in procs[n], n
    raw = open(os.path.join(d, MODULE + '.bas'), 'rb').read()
    raw.decode('cp932')
    assert b'\r\r\n' not in raw and raw.count(b'\r\n') == raw.count(b'\n')
