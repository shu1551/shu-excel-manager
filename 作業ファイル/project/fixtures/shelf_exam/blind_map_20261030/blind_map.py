"""盲点の地図（2026-10-30）: phrases.tsv の頼みを写しの判定 vbam_prefire.shelf_pick で撃つ・数える。

  py blind_map.py          phrases.tsv を撃って results.tsv を書く（正解・判定の列が既にあれば残す）
  py blind_map.py --count  results.tsv の正解・判定の列を分野ごとに数え、誤爆と取りこぼしを書き出す
  py blind_map.py --report results.tsv から report.md を書く（数は全部ここで数える）

棚の本文は 作業ファイル/project/macros/*.bas（CP932）を shelf_entries_from_text で読む。
"""
import csv
import glob
import os
import sys
from collections import Counter, OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
sys.path.insert(0, os.path.join(PROJECT, 'python_scripts'))
try:
    import pythoncom  # noqa: F401
except ImportError:
    # Excel の無い機械（Linux など）: 判定は純 Python なので、COM の部品は空の箱で import だけ通す
    import types
    for _m in ('pythoncom', 'pywintypes', 'win32com', 'win32com.client', 'win32com.client.dynamic'):
        sys.modules.setdefault(_m, types.ModuleType(_m))
    sys.modules['pywintypes'].com_error = type('com_error', (Exception,), {})
import vbam_prefire as vp  # noqa: E402

PHRASES = os.path.join(HERE, 'phrases.tsv')
RESULTS = os.path.join(HERE, 'results.tsv')
AI = 'AI'
HEAD = ['番号', '分野', '頼み', '当たった棚', '正解の棚', '判定']


def shelf_entries():
    text = ''
    for f in sorted(glob.glob(os.path.join(PROJECT, 'macros', '*.bas'))):
        with open(f, encoding='cp932', newline='') as fh:
            text += fh.read() + '\r\n'
    return vp.shelf_entries_from_text(text)


def read_tsv(path):
    with open(path, encoding='utf-8', newline='') as fh:
        return list(csv.DictReader(fh, delimiter='\t'))


def verdict(hit, right):
    """当たり＝正解の棚／正しく AI＝正解が AI で何も当たらず／取りこぼし＝正解が棚なのに当たらず／誤爆＝違う棚が当たった"""
    if not right:
        return ''
    if hit == right:
        return '正しく AI' if hit == AI else '当たり'
    if hit == AI:
        return '取りこぼし'
    return '誤爆'


def fire():
    entries = shelf_entries()
    old = {r['番号']: r for r in read_tsv(RESULTS)} if os.path.exists(RESULTS) else {}
    rows = []
    for r in read_tsv(PHRASES):
        hit = vp.shelf_pick(r['頼み'], entries) or AI
        right = (old.get(r['番号']) or {}).get('正解の棚', '')
        rows.append([r['番号'], r['分野'], r['頼み'], hit, right, verdict(hit, right)])
    with open(RESULTS, 'w', encoding='utf-8', newline='') as fh:
        w = csv.writer(fh, delimiter='\t', lineterminator='\n')
        w.writerow(HEAD)
        w.writerows(rows)
    names = [e['name'] for e in entries]
    print(f'Sub {len(names)} 本・うち依頼の語のある棚 {sum(1 for e in entries if e.get("ask"))} 本・頼み {len(rows)} 個')
    print(f'棚に当たった {sum(1 for x in rows if x[3] != AI)} 個・AI に回った {sum(1 for x in rows if x[3] == AI)} 個')


KINDS = ('当たり', '取りこぼし', '誤爆', '正しく AI')


def count():
    rows = read_tsv(RESULTS)
    by = OrderedDict()
    for r in rows:
        assert r['判定'] == verdict(r['当たった棚'], r['正解の棚']), r['番号']
        by.setdefault(r['分野'], Counter())[r['判定']] += 1
    print('| 分野 | 頼み | ' + ' | '.join(KINDS) + ' | 正解が棚 | 正解が AI |')
    print('|---|' + '---:|' * (len(KINDS) + 3))
    tot = Counter()
    for f, c in by.items():
        n = sum(c.values())
        tot.update(c)
        shelf = c['当たり'] + c['取りこぼし'] + sum(1 for r in rows if r['分野'] == f and r['判定'] == '誤爆' and r['正解の棚'] != AI)
        print(f'| {f} | {n} | ' + ' | '.join(str(c[k]) for k in KINDS) + f' | {shelf} | {n - shelf} |')
    n = sum(tot.values())
    shelf = sum(1 for r in rows if r['正解の棚'] != AI)
    print('| 計 | %d | ' % n + ' | '.join(str(tot[k]) for k in KINDS) + f' | {shelf} | {n - shelf} |')
    mis = [r for r in rows if r['判定'] == '誤爆']
    print()
    print(f'誤爆 {len(mis)} 個のうち 正解が AI {sum(1 for r in mis if r["正解の棚"] == AI)} 個・正解が別の棚 {sum(1 for r in mis if r["正解の棚"] != AI)} 個')
    print('誤爆で当たった棚（多い順）:')
    for name, k in Counter(r['当たった棚'] for r in mis).most_common():
        print(f'  {k}\t{name}')
    print('取りこぼした正解の棚（多い順）:')
    for name, k in Counter(r['正解の棚'] for r in rows if r['判定'] == '取りこぼし').most_common():
        print(f'  {k}\t{name}')


# 誤爆の危なさ＝当たった棚が何をするか（棚の目録 shelf_catalog.tsv と各 Sub の頭の説明から）。小さいほど危ない
DANGER_LABEL = {
    1: '行や列を消す・足す・動かす',
    2: '値・式・名前の定義を書き換える／消す',
    3: 'シート・ファイル・ピボット・グラフを作る／並びや形を変える',
    4: '書式・表示・印刷の設定だけ変える',
    5: '調べて一覧にするだけ',
}
DANGER = {
    '全列が同じ重複行を削除する': 1, '二段の見出しを一行に畳む': 1, '選んだ列を選んだ順に左へ移す': 1, '表をテーブルにして集計行を出す': 1,
    '表の書き方と罫線と列幅をそろえる': 2, 'セルの見えない空白と改行を取る': 2, '選んだ列の文字の日付を日付にする': 2,
    '番号の列を連番に振り直す': 2, '選んだ列の途切れた式を戻す': 2, '他のシートの表を1枚に集約する': 2, '名前を一覧にし壊れた名前を削除': 2,
    '選んだ3列でピボットを作る': 3, '選んだ3列でダッシュボードを作る': 3, '選んだ列の値ごとにシートを分ける': 3,
    '選んだ列をクエリで区切って分ける': 3, '同じ形のシートを縦につなぐ': 3, '一覧を1件1枚でひな形に差し込む': 3,
    '選んだ値に合う行を抜き出す': 3, 'シートを別ブックで保存する': 3, 'シートをPDFで保存する': 3, '別シートと突合し差異を書く': 3,
    '選んだ列で表を並べ替える': 3, '頼みの文でシートの名前と色と並びを変える': 3, '頼みの文でピボットを手直しする': 3,
    '頼みの文でグラフの種類を変える': 3, 'テーブルの範囲を下と右へ広げる': 3,
    '頼みの文でシートの表示と印刷を変える': 4, '頼みの文で指した列と行の幅と表示を変える': 4, '頼みの文で指した列で表を絞り込む': 4,
    '選んだ行を1ページの高さにする': 4, '選んだ列の数値を円マーク通貨表示にする': 4, 'グラフの書体と凡例をそろえる': 4,
    '隠れた行と列を全部表示する': 4, '列幅と行の高さを全シートに写す': 4,
    '非表示のシートと行列を一覧にする': 5, 'エラー値のセルを一覧にする': 5, '選んだ列の名前の揺れを一覧にする': 5,
    '選んだ列の重複する値を一覧にする': 5, '参照とリンクを一覧にする': 5,
}
# 取りこぼした頼みの言い方の型（番号 → 型）
MISS_TYPE = OrderedDict((
    ('仕事の言葉で比べ方を言う（前年比・差・増えた／なくなった・合わない・違うところ）', (135, 149, 203, 207, 212, 213, 214)),
    ('Excel の機能や画面の言葉で言う（VLOOKUP・2軸・更新・元データ・1ページ・枠線・先頭行を固定・コピペ）', (205, 163, 177, 184, 226, 247, 84, 16)),
    ('調べる頼みを「教えて・確認・比べて・どこにあるか」で言う（一覧と言わない）', (276, 283, 284, 299, 300)),
    ('選んだ物・指した物を短く言う（見出しの行・この列・選択した行・選んだセル）', (26, 90, 92, 113)),
    ('「〜ごとに・〜別の」で集計を頼む', (127, 133)),
    ('色の系統で見た目を頼む（グレーを基調・青系）', (308, 314)),
    ('あいまい・似ている・表記ゆれで突き合わせを頼む', (218, 225)),
    ('棚の語の言い換え（何も入ってない列・行に〜を入れて）', (13, 178)),
))


def report():
    rows = read_tsv(RESULTS)
    entries = shelf_entries()
    out = []
    w = out.append
    by = OrderedDict()
    for r in rows:
        assert r['判定'] == verdict(r['当たった棚'], r['正解の棚']), r['番号']
        by.setdefault(r['分野'], []).append(r)
    tot = Counter(r['判定'] for r in rows)
    shelf_n = sum(1 for e in entries if e.get('ask'))
    w('# 盲点の地図（2026-10-30）')
    w('')
    w('棚に、作った本人以外の言い方で頼んだときの当たり・取りこぼし・誤爆を測った。語は足していない・棚は直していない。')
    w('')
    w('- 頼み: `phrases.tsv`（13 分野×25 個＝%d 個）。棚のコードと問題集を読む前に書いて先に commit した' % len(rows))
    w('- 判定: `vbam_prefire.shelf_pick`。棚は `macros/*.bas` を `shelf_entries_from_text` で読んだ Sub %d 本・うち依頼の語のある棚 %d 本'
      % (len(entries), shelf_n))
    w('- 結果: `results.tsv`（当たった棚・正解の棚・判定）。この文書の数は全部 `py blind_map.py --report` が数えた')
    w('')
    w('## 正解の決め方')
    w('')
    w('- 正解の棚＝頼みの言い方のまま撃ったとき、頼んだとおりに動く棚（目録 `shelf_catalog.tsv`・各 Sub の頭の説明・本文で確かめた）。無ければ AI')
    w('- 「頼みの文で指した所の書式を変える」「頼みの文で指した列の表示形式を変える」「頼みの文で指した列と行の幅と表示を変える」は、'
      '頼みに「見出し・合計行・選んだ・この列・〇〇の列」などの「どこを」の語が無いと断る作り（`表の整理_基本操作.bas` の `基_指した所` と幅の棚の `先撃ちに断る "どの…か分かりませんでした"`）。'
      'そのため「どこを」の無い頼み（例: 「%s」「%s」）は正解を AI にした' % (_row(rows, 29)['頼み'], _row(rows, 52)['頼み']))
    w('- 判定: 当たり＝正解の棚／取りこぼし＝正解が棚なのに AI に回った／誤爆＝正解と違う棚が当たった（正解が AI のときも含む）／正しく AI＝正解が AI で AI に回った')
    w('')
    w('## 分野ごとの数')
    w('')
    w('| 分野 | 頼み | ' + ' | '.join(KINDS) + ' | 正解が棚 | 正解が AI |')
    w('|---|' + '---:|' * (len(KINDS) + 3))
    for f, rs in by.items():
        c = Counter(r['判定'] for r in rs)
        sh = sum(1 for r in rs if r['正解の棚'] != AI)
        w(f'| {f} | {len(rs)} | ' + ' | '.join(str(c[k]) for k in KINDS) + f' | {sh} | {len(rs) - sh} |')
    sh = sum(1 for r in rows if r['正解の棚'] != AI)
    w(f'| 計 | {len(rows)} | ' + ' | '.join(str(tot[k]) for k in KINDS) + f' | {sh} | {len(rows) - sh} |')
    w('')
    hit_shelf = sum(1 for r in rows if r['当たった棚'] != AI)
    w(f'- 棚に当たった {hit_shelf} 個のうち 当たり {tot["当たり"]} 個・誤爆 {tot["誤爆"]} 個')
    w(f'- 正解が棚の {sh} 個のうち 当たり {tot["当たり"]} 個・取りこぼし {tot["取りこぼし"]} 個・'
      f'別の棚に誤爆 {sum(1 for r in rows if r["判定"] == "誤爆" and r["正解の棚"] != AI)} 個')
    w(f'- 正解が AI の {len(rows) - sh} 個のうち 正しく AI {tot["正しく AI"]} 個・棚に誤爆 {sum(1 for r in rows if r["判定"] == "誤爆" and r["正解の棚"] == AI)} 個')
    w('')
    mis = [r for r in rows if r['判定'] == '誤爆']
    for r in mis:
        assert r['当たった棚'] in DANGER, r['当たった棚']
    mis.sort(key=lambda r: (DANGER[r['当たった棚']], int(r['番号'])))
    w('## 誤爆の一覧（危ない順）')
    w('')
    w('危なさは当たった棚が何をするか（目録の説明）で分けた。棚が形の合わないシートで断る場合も、撃たれたときにする事で数えている。')
    w('')
    dc = Counter(DANGER[r['当たった棚']] for r in mis)
    for k, lab in DANGER_LABEL.items():
        w(f'- {k}. {lab}: {dc[k]} 個')
    for k, lab in DANGER_LABEL.items():
        rs = [r for r in mis if DANGER[r['当たった棚']] == k]
        if not rs:
            continue
        w('')
        w(f'### {k}. {lab}（{len(rs)} 個）')
        w('')
        w('| 番号 | 分野 | 頼み | 当たった棚 | 正解の棚 |')
        w('|---:|---|---|---|---|')
        for r in rs:
            w(f'| {r["番号"]} | {r["分野"]} | {r["頼み"]} | {r["当たった棚"]} | {r["正解の棚"]} |')
    w('')
    w('### 誤爆で当たった棚（多い順）')
    w('')
    w('| 回数 | 棚 | 危なさ |')
    w('|---:|---|---:|')
    for name, k in Counter(r['当たった棚'] for r in mis).most_common():
        w(f'| {k} | {name} | {DANGER[name]} |')
    w('')
    miss = [r for r in rows if r['判定'] == '取りこぼし']
    typed = {n: t for t, ns in MISS_TYPE.items() for n in ns}
    assert set(typed) == {int(r['番号']) for r in miss}, set(typed) ^ {int(r['番号']) for r in miss}
    w('## 取りこぼしの多い言い方の型')
    w('')
    w(f'取りこぼし {len(miss)} 個を言い方の型で分けた（多い順）。')
    for t, ns in sorted(MISS_TYPE.items(), key=lambda x: -len(x[1])):
        w('')
        w(f'### {t}（{len(ns)} 個）')
        w('')
        w('| 番号 | 頼み | 正解の棚 |')
        w('|---:|---|---|')
        for n in sorted(ns):
            r = _row(rows, n)
            w(f'| {n} | {r["頼み"]} | {r["正解の棚"]} |')
    w('')
    with open(os.path.join(HERE, 'report.md'), 'w', encoding='utf-8', newline='\n') as fh:
        fh.write('\n'.join(out))
    print(f'report.md を書いた（誤爆 {len(mis)} 個・取りこぼし {len(miss)} 個）')


def _row(rows, n):
    return next(r for r in rows if int(r['番号']) == n)


if __name__ == '__main__':
    if '--count' in sys.argv:
        count()
    elif '--report' in sys.argv:
        report()
    else:
        fire()
