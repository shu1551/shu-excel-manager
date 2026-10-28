# -*- coding: utf-8 -*-
"""棚を鍛える固定の形（shelf --train・2026-10-23）。

これまで棚の鍛え方は、頼みを書く・撃つ・分ける・測るを、毎回指示書で組み直していた。
やり方の中身（代役の頼み・見せない頼み・語を足すのは言い方の違う 2 つ以上で外れた型だけ）は固まっていたが、
決まりを守らせるのは指示書と人の目だった（検査を外して焼く・通らない問を黙って外す・後から頼みを書き換える）。
ここでは決まりを道具の側に持たせる。純 Python の部品（読む・凍結・採点・仕分け・書き出し・守り）と、
Excel を呼ぶ薄い入口（cmd_train）に分けてあり、部品は Excel なしで試験できる。

段: new（頼みの置き場） → freeze（凍結） → before（撃つ） → sort（仕分け） → [人が棚を直す] → export（問題集へ）
    → after（もう一度撃つ） → close（帳尻・検査を外していないか）
"""
import csv
import datetime
import glob
import hashlib
import io
import json
import os
import re
import time

HEADER = ['番号', '型', '頼み', '見せない', '期待', '正解の棚']
NONE_SHELF = '(なし)'            # 正解の棚の欄: どの棚でも片付かない（棚の候補）
CATS = ['正しく当たった', '正しくAIへ', '誤爆', '取りこぼし', '棚が無い', '未確定']
MIN_ROWS = 12
HELD_MIN, HELD_MAX = 0.25, 0.5   # 見せない頼みの割合の幅（3 分の 1 が目安）

PROCEDURE = """棚を鍛える形（shelf --train）。段は 1 本の道:
  1. new <名前>      頼みの置き場 phrases.tsv を作る
                    ※頼みは棚の目録（shelf --grep）・棚のコードを見る前に全部書く。使う人の本物の言い方（外れた依頼.csv）だけ読んでよい
                    ※列: 番号・型・頼み・見せない(0/1)・期待(棚/AI)。正解の棚の欄は仕分けのとき埋める。見せない頼みは 3 分の 1
  2. freeze <名前>   凍結（頼みの文・型・見せない・期待の指紋を残す。以後 1 字でも変えると after / close が止まる）
  3. before <名前>   VBA の判定（正本）に全部撃って記録する（AI 0 回・数秒）
  4. sort <名前>     外れ・誤爆を型ごとに分ける。語を足してよいのは「言い方の違う 2 つ以上で外れた型」だけ
                    ここで初めて目録を見てよい。正解の棚の欄を埋める（棚が無い物は (なし)）
  5. （人が）棚の語・除外を直す → probe <名前> "似た頼み" …（3〜5 個・誤爆を見る）→ shelf --check
  6. export <名前>   今の判定で期待どおりの学習用の頼みだけを問題集 entry_phrases_訓練_<名前>.tsv へ（見せない頼みは入れない）
  7. after <名前>    もう一度撃つ（頼みが変わっていれば止まる）
  8. close <名前>    帳尻: 見せない頼みの前後・誤爆が増えたら足した語を外す・検査を外して焼いていないか・shelf --check
他: ledger（これまでの帳尻の一覧）"""

TEMPLATE_NOTE = (
    "# 棚を鍛える頼み。タブ区切り UTF-8。# で始まる行は読まない。\n"
    "# 型: 同じ仕事の言い方をまとめる短い名前（例: ピボットを作る・件数）。語を足してよいかの判定は型の中の言い方の数で決まる。\n"
    "# 見せない: 1 = 見せない頼み（語を足す材料にしない・問題集に入れない。前後の当たり率を測るだけ）。全体の 3 分の 1 が目安。\n"
    "# 期待: 棚 = 棚で片付くはず／AI = AI に回すのが正しい。2 割ほどは AI でないと無理な物を混ぜる。\n"
    "# 正解の棚: 書き始めは空。sort の後、目録を見て棚の名前を入れる（どの棚でも無理なら (なし)）。\n"
)


# ---- 置き場 ----
def train_root():
    import vbam_prefire as vp
    return os.path.join(vp.exam_dir(), 'train')


def run_dir(name, root=None):
    n = str(name or '').strip()
    if not n or re.search(r'[\\/:*?"<>|\s]', n):
        raise ValueError(f"名前に使えない字があります: 「{n}」（空白・\\ / : * ? \" < > | は使えません）")
    return os.path.join(root or train_root(), n)


def _p(d, fn):
    return os.path.join(d, fn)


# ---- 頼みを読む ----
def load_phrases(path):
    """phrases.tsv → (行の並び, 誤りの並び)。行 = dict(no, kind, text, held, expect, shelf)。"""
    rows, errs, seen = [], [], set()
    if not os.path.isfile(path):
        return [], [f"頼みのファイルがありません: {path}"]
    header_seen = False
    for lineno, ln in enumerate(io.open(path, encoding='utf-8', errors='replace'), 1):
        ln = ln.rstrip('\r\n')
        if not ln.strip() or ln.lstrip().startswith('#'):
            continue
        p = [x.strip() for x in ln.split('\t')]
        if not header_seen:
            if p[0] != '番号':
                errs.append(f"{lineno} 行目: 見出し行（番号・型・頼み・見せない・期待・正解の棚）がありません")
                return rows, errs
            header_seen = True
            continue
        p += [''] * (6 - len(p))
        no, kind, text, held, expect, shelf = p[:6]
        if not text:
            errs.append(f"{lineno} 行目: 頼みが空です")
            continue
        if held not in ('0', '1'):
            errs.append(f"{lineno} 行目: 見せない は 0 か 1（「{held}」）")
            continue
        if expect not in ('棚', 'AI'):
            errs.append(f"{lineno} 行目: 期待は 棚 か AI（「{expect}」）")
            continue
        if not kind:
            errs.append(f"{lineno} 行目: 型が空です（語を足してよいかの判定に使います）")
            continue
        if text in seen:
            errs.append(f"{lineno} 行目: 同じ頼みが 2 回あります（「{text[:30]}」）")
            continue
        seen.add(text)
        rows.append({'no': no or str(len(rows) + 1), 'kind': kind, 'text': text,
                     'held': held == '1', 'expect': expect, 'shelf': shelf})
    if not header_seen:
        errs.append("見出し行がありません")
    return rows, errs


def validate(rows):
    """凍結してよいか → (止める誤り, 注意)。"""
    errs, warns = [], []
    n = len(rows)
    if n < MIN_ROWS:
        errs.append(f"頼みが {n} 個しかありません（{MIN_ROWS} 個以上要ります）")
        return errs, warns
    held = sum(1 for r in rows if r['held'])
    frac = held / n
    if frac < HELD_MIN or frac > HELD_MAX:
        errs.append(f"見せない頼みが {held}/{n}（{frac:.0%}）です。{HELD_MIN:.0%}〜{HELD_MAX:.0%} にしてください（3 分の 1 が目安）")
    if not any(r['expect'] == 'AI' for r in rows):
        warns.append("期待が AI の頼みが 1 つもありません（誤爆を測れません。2 割ほど混ぜてください）")
    kinds = {}
    for r in rows:
        if not r['held']:
            kinds.setdefault(r['kind'], set()).add(r['text'])
    for k, ts in sorted(kinds.items()):
        if len(ts) < 2 and any(r['kind'] == k and r['expect'] == '棚' for r in rows):
            warns.append(f"型「{k}」は学習用の言い方が 1 つだけです＝外れても語は足せません")
    return errs, warns


def fingerprint(rows):
    """頼みの文・型・見せない・期待の指紋。正解の棚（仕分けで埋める欄）は入れない。"""
    body = '\n'.join('\t'.join([r['no'], r['kind'], r['text'], '1' if r['held'] else '0', r['expect']])
                     for r in sorted(rows, key=lambda r: r['no']))
    return hashlib.sha256(body.encode('utf-8')).hexdigest()


def freeze(d):
    """凍結する → (ok, 行の並び)。"""
    rows, errs = load_phrases(_p(d, 'phrases.tsv'))
    v_errs, warns = validate(rows) if not errs else ([], [])
    errs = errs + v_errs
    if errs:
        return False, errs + [f"注意: {w}" for w in warns]
    meta = {'frozen_at': time.strftime('%Y-%m-%d %H:%M:%S'), 'fingerprint': fingerprint(rows),
            'n': len(rows), 'held': sum(1 for r in rows if r['held'])}
    with io.open(_p(d, 'run.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)
    out = [f"凍結しました: {meta['n']} 個（学習用 {meta['n'] - meta['held']}・見せない {meta['held']}）"]
    out += [f"注意: {w}" for w in warns]
    return True, out


def read_meta(d):
    try:
        with io.open(_p(d, 'run.json'), encoding='utf-8') as f:
            m = json.load(f)
        return m if isinstance(m, dict) else None
    except (OSError, ValueError):
        return None


def frozen_ok(d):
    """凍結後に頼みが書き換わっていないか → (ok, 理由)。"""
    meta = read_meta(d)
    if not meta or not meta.get('fingerprint'):
        return False, "まだ凍結していません（freeze が先）"
    rows, errs = load_phrases(_p(d, 'phrases.tsv'))
    if errs:
        return False, "頼みのファイルに誤りがあります: " + errs[0]
    if fingerprint(rows) != meta['fingerprint']:
        return False, ("凍結のあとに頼みの文・型・見せない・期待が書き換わっています。"
                       "結果に合わせて頼みを直すのは丸暗記の元＝元に戻してください（足したい頼みは新しい名前の run で）")
    return True, ''


# ---- 採点 ----
def judge(row, got):
    """1 つの頼みの判定。got = 判定が選んだ棚の名前（空 = AI へ）。"""
    got = got or ''
    if row['expect'] == 'AI':
        return '誤爆' if got else '正しくAIへ'
    pinned = row['shelf']
    if pinned == NONE_SHELF:
        return '誤爆' if got else '棚が無い'
    if not got:
        return '取りこぼし'
    if not pinned:
        return '未確定'
    return '正しく当たった' if got == pinned else '誤爆'


def shoot(rows, answer_fn):
    """全部の頼みを判定に撃つ。answer_fn(頼みの並び) → (VBA の答え, Python の写しの答え)。
    VBA を呼べなければ answer_fn が例外を投げる（Python の写しで黙って代用しない）。"""
    texts = [r['text'] for r in rows]
    vba, py = answer_fn(texts)
    if len(vba) != len(texts):
        raise RuntimeError("判定の答えの数が頼みの数と合いません")
    out = []
    for r, a, b in zip(rows, vba, py):
        rec = dict(r)
        rec['got'] = a or ''
        rec['py'] = b or ''
        rec['cat'] = judge(r, a)
        out.append(rec)
    return out


def write_results(path, results):
    with io.open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\t'.join(HEADER + ['当たった棚', '写し', '判定']) + '\n')
        for r in results:
            f.write('\t'.join([r['no'], r['kind'], r['text'], '1' if r['held'] else '0', r['expect'],
                               r['shelf'], r['got'], r['py'], r['cat']]) + '\n')


def read_results(path):
    if not os.path.isfile(path):
        return None
    out = []
    for i, ln in enumerate(io.open(path, encoding='utf-8')):
        if i == 0:
            continue
        p = ln.rstrip('\r\n').split('\t') + [''] * 9
        out.append({'no': p[0], 'kind': p[1], 'text': p[2], 'held': p[3] == '1', 'expect': p[4],
                    'shelf': p[5], 'got': p[6], 'py': p[7], 'cat': p[8]})
    return out


def summarize(results, held=None):
    """held=None 全部／True 見せない／False 学習用 → dict(件数, 当たり率, 誤爆, 写しとの食い違い)。"""
    rs = [r for r in results if held is None or r['held'] == held]
    c = {k: 0 for k in CATS}
    for r in rs:
        c[r['cat']] = c.get(r['cat'], 0) + 1
    denom = c['正しく当たった'] + c['取りこぼし'] + sum(1 for r in rs if r['cat'] == '誤爆' and r['expect'] == '棚')
    return {'n': len(rs), 'cats': c,
            'rate': (c['正しく当たった'] / denom) if denom else None,
            'wrong': c['誤爆'],
            'parity': sum(1 for r in rs if r['got'] != r['py'])}


def _fmt_summary(label, s):
    c = s['cats']
    rate = '-' if s['rate'] is None else f"{s['rate']:.0%}"
    return (f"{label} {s['n']:3d} 個: 正しく当たった {c['正しく当たった']}・正しくAIへ {c['正しくAIへ']}・誤爆 {c['誤爆']}・"
            f"取りこぼし {c['取りこぼし']}・棚が無い {c['棚が無い']}・未確定 {c['未確定']}（棚の当たり率 {rate}）")


# ---- 仕分け ----
def sort_rows(results):
    """学習用の外れ・誤爆を型ごとに分ける（見せない頼みは材料にしない）。
    戻り値: [{'verdict', 'kind', 'shelf', 'items': [(頼み, 当たった棚, 判定)]}]"""
    groups = {}
    for r in results:
        if r['held'] or r['cat'] in ('正しく当たった', '正しくAIへ'):
            continue
        # 型が広すぎても別々の仕事を「同じ型の 2 つの言い方」と数えないよう、正解の棚でも分ける（1 周目で見つけた穴）
        groups.setdefault((r['kind'], r['cat'], r['shelf']), []).append(r)
    out = []
    for (kind, cat, _pin), rs in sorted(groups.items()):
        distinct = {r['text'] for r in rs}
        shelves = sorted({r['shelf'] for r in rs if r['shelf'] and r['shelf'] != NONE_SHELF})
        if cat == '取りこぼし':
            verdict = ('語を足す候補' if len(distinct) >= 2 else '待つ（言い方が 1 つだけ＝語は足さない）')
        elif cat == '誤爆':
            verdict = ('除外か正しい棚の語を直す候補' if len(distinct) >= 2 else '待つ（1 つだけ＝問題集に入れて様子を見る）')
        elif cat == '棚が無い':
            verdict = '棚の候補（作るかは人が決める）'
        else:
            verdict = '正解の棚を決める（phrases.tsv の正解の棚の欄）'
        out.append({'verdict': verdict, 'kind': kind, 'cat': cat, 'shelves': shelves,
                    'items': [(r['text'], r['got'], r['cat']) for r in rs]})
    return out


def sort_text(groups):
    if not groups:
        return ["学習用の頼みは全部期待どおりです（仕分ける物はありません）"]
    lines = []
    for g in groups:
        sh = f" 正解の棚: {' / '.join(g['shelves'])}" if g['shelves'] else ''
        lines.append(f"■ 型「{g['kind']}」 {g['cat']} {len(g['items'])} 個 → {g['verdict']}{sh}")
        for t, got, _c in g['items'][:8]:
            lines.append(f"    「{t[:50]}」" + (f" → 当たった: {got}" if got else ' → AIへ'))
    return lines


# ---- 問題集へ ----
def export_exam(results, out_path, existing_texts=()):
    """今の判定で期待どおりの学習用の頼みだけを問題集に書く → (書いた数, 書かなかった物[(頼み, 理由)])。
    期待と違う物を問題集に入れると shelf --check が止まる＝直してから。見せない頼みは入れない。"""
    existing = set(existing_texts)
    lines, skipped = [], []
    for r in results:
        if r['held']:
            continue
        if r['cat'] == '正しくAIへ':
            exp = ''
        elif r['cat'] == '正しく当たった':
            exp = r['shelf']
        else:
            why = {'誤爆': '誤爆のまま', '取りこぼし': '取りこぼしのまま', '棚が無い': '棚が無い（棚の候補）',
                   '未確定': '正解の棚が未確定'}.get(r['cat'], r['cat'])
            skipped.append((r['text'], why))
            continue
        if r['text'] in existing:
            skipped.append((r['text'], '問題集にすでにある'))
            continue
        lines.append(f"{exp}\t{r['text']}\ttrain:{r['kind']}")
    if lines:
        with io.open(out_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write('\n'.join(lines) + '\n')
    return len(lines), skipped


# ---- 守り ----
def no_check_burns(since, log_path=None):
    """since（'YYYY-MM-DD HH:MM:SS'）以降に register-addin を --no-check / --no-test で撃った記録 → [(時刻, 引数)]。"""
    if log_path is None:
        try:
            import vbam_core
            log_path = vbam_core._CALL_LOG_FILE
        except Exception:
            return []
    out = []
    if not os.path.isfile(log_path):
        return out
    for ln in io.open(log_path, encoding='utf-8', errors='replace'):
        try:
            rec = json.loads(ln)
        except ValueError:
            continue
        if str(rec.get('cmd')) != 'register-addin' or str(rec.get('time', '')) < since:
            continue
        a = str(rec.get('args', ''))
        if '--no-check' in a or '--no-test' in a:
            out.append((rec.get('time', ''), a))
    return out


# ---- 帳尻の台帳 ----
def ledger_path(root=None):
    return os.path.join(root or train_root(), 'ledger.tsv')


def ledger_append(run, stage, summ_all, summ_held, note='', root=None):
    p = ledger_path(root)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    new = not os.path.isfile(p)
    with io.open(p, 'a', encoding='utf-8', newline='\n') as f:
        if new:
            f.write('日時\t名前\t段\t全部\t当たり\t正しくAIへ\t誤爆\t取りこぼし\t見せない当たり率\t見せない誤爆\t一言\n')
        c = summ_all['cats']
        rate = '' if summ_held['rate'] is None else f"{summ_held['rate']:.2f}"
        f.write('\t'.join([time.strftime('%Y-%m-%d %H:%M:%S'), run, stage, str(summ_all['n']),
                           str(c['正しく当たった']), str(c['正しくAIへ']), str(c['誤爆']), str(c['取りこぼし']),
                           rate, str(summ_held['wrong']), note.replace('\t', ' ')]) + '\n')


def close_verdict(before, after):
    """見せない頼みの前後 → (止める理由の並び, 見る点の並び)。"""
    stops, notes = [], []
    b, a = summarize(before, True), summarize(after, True)
    all_a = summarize(after)
    if all_a['wrong'] > 0:
        stops.append(f"誤爆が 0 ではありません（全体で {all_a['wrong']} 個・うち見せない頼み {a['wrong']} 個）。誤爆を 0 にしてから閉じてください")
    if a['wrong'] > b['wrong']:
        stops.append(f"見せない頼みで誤爆が増えました（{b['wrong']} → {a['wrong']}）。足した語・除外を外してください")
    if a['cats']['正しく当たった'] < b['cats']['正しく当たった']:
        stops.append(f"見せない頼みで正しく当たった数が減りました（{b['cats']['正しく当たった']} → {a['cats']['正しく当たった']}）")
    if a['cats']['取りこぼし'] > b['cats']['取りこぼし']:
        stops.append(f"見せない頼みで取りこぼしが増えました（{b['cats']['取りこぼし']} → {a['cats']['取りこぼし']}）")
    parity = all_a['parity']              # 学習用も見せない頼みも（写しの食い違いは道具の穴で、どちらの頼みでも起こる）
    if parity:
        stops.append(f"VBA と Python の写しの答えが食い違う頼みが {parity} 個あります（register-addin で写しを揃える）")
    if not a['n']:
        notes.append("見せない頼みがありません")
    return stops, notes


# ---- Excel を呼ぶ薄い入口 ----
def _connect(target=None):
    """(xl, 棚の持ち主, 棚のテキスト)。前に出ているブックに コンボ道具 があればそれ（shelf --check と同じ選び方）。"""
    from vbam_core import get_workbook, shelf_text
    import vbam_view
    xl, wb = get_workbook(target)
    owner, lines = vbam_view._shelf_source(xl)
    if not owner:
        raise RuntimeError("棚（モジュール「表の整理」）が開いているブック・アドインに見つかりません")
    src, text = owner, "\n".join(lines)
    try:
        wb.VBProject.VBComponents('コンボ道具')
        t = shelf_text(wb.VBProject)
        if t:
            src, text = str(wb.Name), t
    except Exception:
        pass
    return xl, src, text


def _shelf_signature(text):
    """棚の当たり方を決める部分（名前・依頼の語・依頼の組・各種の印）だけの指紋。焼くときの整形や処理の中身の違いでは変わらない。"""
    import vbam_prefire as vp
    keys = ('name', 'ask', 'combo', 'cf_ok', 'allow_fix', 'look_only', 'plan_substitute')
    return sorted(tuple(repr(e.get(k)) for k in keys) for e in vp.shelf_entries_from_text(text) or ())


def stale_note(xl, src, text):
    """判定に使った棚（src）と、焼いた棚の持ち主（秀コンボ.xlam など）の中身が違うときの注意文。違わなければ ''。
    shelf --ask は焼いた持ち主を見る。before / after / probe / shelf --check は直している側を見る＝食い違う間は直しが使う人に届かない。"""
    try:
        import vbam_view
        owner, lines = vbam_view._shelf_source(xl)
        if owner and owner != src and _shelf_signature("\n".join(lines)) != _shelf_signature(text):
            return (f"注意: 判定は {src} の棚で撃ちました。焼いた {owner} は棚の中身が違います"
                    f"（register-addin で焼き直すまで、shelf --ask と使う人の Excel には今の直しが届きません）")
    except Exception:
        pass
    return ''


def _real_answer_fn(xl, owner, text):
    import vbam_prefire as vp
    entries = vp.shelf_entries_from_text(text)

    def fn(texts):
        vba = vp.vba_answers(xl, owner, texts)
        if vba is None:
            raise RuntimeError("VBA の判定（コンボ道具.先撃ち候補の一括）を呼べません。"
                               "Python の写しでは代用しません＝秀コンボを読み込んでから")
        return vba, [vp.shelf_pick(t, entries) for t in texts]
    return fn


def _need_frozen(d):
    ok, why = frozen_ok(d)
    if not ok:
        print(why)
    return ok


def cmd_train(args):
    """`shelf --train [new|freeze|before|sort|export|after|close|ledger] [名前]`。"""
    pos = list(getattr(args, 'posargs', None) or [])
    if not pos:
        print(PROCEDURE)
        return True
    sub = pos[0]
    if sub == 'ledger':
        p = ledger_path()
        if not os.path.isfile(p):
            print("まだ帳尻の記録がありません（close を撃つと残ります）")
            return True
        print(io.open(p, encoding='utf-8').read().rstrip())
        return True
    if len(pos) < 2:
        print(f"名前が要ります: shelf --train {sub} <名前>")
        return False
    try:
        d = run_dir(pos[1])
    except ValueError as e:
        print(str(e))
        return False

    if sub == 'new':
        if os.path.exists(_p(d, 'phrases.tsv')):
            print(f"すでにあります: {_p(d, 'phrases.tsv')}（別の名前にするか、凍結前なら直接編集）")
            return False
        os.makedirs(d, exist_ok=True)
        with io.open(_p(d, 'phrases.tsv'), 'w', encoding='utf-8', newline='\n') as f:
            f.write(TEMPLATE_NOTE + '\t'.join(HEADER) + '\n')
        print(f"頼みの置き場を作りました: {_p(d, 'phrases.tsv')}")
        print("頼みは、棚の目録・棚のコードを見る前に全部書く（使う人の本物の外れだけ読んでよい）。書けたら freeze。")
        return True

    if sub == 'freeze':
        ok, out = freeze(d)
        for ln in out:
            print(ln)
        return ok

    if sub in ('before', 'after'):
        if not _need_frozen(d):
            return False
        if sub == 'after' and not os.path.isfile(_p(d, 'before.tsv')):
            print("before がまだです（凍結のあと、棚を直す前に撃っておく）")
            return False
        rows, _e = load_phrases(_p(d, 'phrases.tsv'))
        t0 = time.time()
        try:
            xl, owner, text = _connect(getattr(args, 'excel_file', None))
            results = shoot(rows, _real_answer_fn(xl, owner, text))
        except Exception as e:
            print(f"撃てませんでした: {e}")
            return False
        write_results(_p(d, sub + '.tsv'), results)
        print(f"{sub}: {len(results)} 個を VBA の判定（{owner}）に撃ちました（{time.time() - t0:.1f} 秒・AI 0 回）→ {_p(d, sub + '.tsv')}")
        stale = stale_note(xl, owner, text)
        if stale:
            print(stale)
        print(_fmt_summary('学習用  ', summarize(results, False)))
        print(_fmt_summary('見せない', summarize(results, True)))
        sa = summarize(results)
        if sa['parity']:
            print(f"注意: VBA と Python の写しの答えが食い違う頼み {sa['parity']} 個（register-addin で揃える）")
        if sub == 'before':
            ledger_append(pos[1], 'before', sa, summarize(results, True))
        return True

    if sub == 'probe':
        # 語を足したあと、似た頼みを撃って誤爆しないか見る（before / after と同じ判定。shelf --ask は焼いた側を見るので使わない）
        probes = [x for x in pos[2:] if str(x).strip()]
        if not probes:
            print('撃つ頼みを渡してください: shelf --train probe <名前> "頼み1" "頼み2" …（3〜5 個）')
            return False
        try:
            xl, owner, text = _connect(getattr(args, 'excel_file', None))
            vba, py = _real_answer_fn(xl, owner, text)(probes)
        except Exception as e:
            print(f"撃てませんでした: {e}")
            return False
        print(f"probe（判定: {owner}）")
        for q, a, b in zip(probes, vba, py):
            print(f"  「{q}」 → {a or '(AIへ)'}" + ('' if a == b else f"   ※写しは {b or '(AIへ)'}"))
        stale = stale_note(xl, owner, text)
        if stale:
            print(stale)
        return True

    if sub == 'sort':
        results = read_results(_p(d, 'before.tsv'))
        if results is None:
            print("before がまだです")
            return False
        # 正解の棚は凍結の指紋に入らない＝仕分けで埋めてよい。最新の phrases.tsv から正解の棚だけ読み直す
        rows, _e = load_phrases(_p(d, 'phrases.tsv'))
        pinned = {r['text']: r['shelf'] for r in rows}
        for r in results:
            r['shelf'] = pinned.get(r['text'], r['shelf'])
            r['cat'] = judge(r, r['got'])
        groups = sort_rows(results)
        lines = sort_text(groups)
        for ln in lines:
            print(ln)
        with io.open(_p(d, 'sort.md'), 'w', encoding='utf-8', newline='\n') as f:
            f.write('\n'.join(lines) + '\n')
        print("決まり: 語を足すのは「語を足す候補」の型だけ。足したら shelf --check と、似た頼みを 3〜5 個 shelf --ask で撃つ。")
        return True

    if sub == 'export':
        if not _need_frozen(d):
            return False
        results = read_results(_p(d, 'after.tsv')) or read_results(_p(d, 'before.tsv'))
        if results is None:
            print("before / after がありません")
            return False
        rows, _e = load_phrases(_p(d, 'phrases.tsv'))
        pinned = {r['text']: r['shelf'] for r in rows}
        for r in results:
            r['shelf'] = pinned.get(r['text'], r['shelf'])
            r['cat'] = judge(r, r['got'])
        import vbam_prefire as vp
        existing = [q for _e2, q in vp.exam_questions()]
        out_path = os.path.join(vp.exam_dir(), f"entry_phrases_訓練_{pos[1]}.tsv")
        n, skipped = export_exam(results, out_path, existing)
        print(f"問題集に書いた: {n} 個 → {out_path}")
        if skipped:
            print(f"書かなかった: {len(skipped)} 個（理由つき・黙って外していません）")
            for t, why in skipped[:30]:
                print(f"  「{t[:40]}」 … {why}")
        print("見せない頼みは問題集に入れていません（次の回でまた測るため）。次: shelf --check")
        return True

    if sub == 'close':
        ok, why = frozen_ok(d)
        if not ok:
            print(why)
            return False
        before, after = read_results(_p(d, 'before.tsv')), read_results(_p(d, 'after.tsv'))
        if before is None or after is None:
            print("before と after の両方が要ります")
            return False
        meta = read_meta(d) or {}
        rows, _e = load_phrases(_p(d, 'phrases.tsv'))
        pinned = {r['text']: r['shelf'] for r in rows}
        for res in (before, after):
            for r in res:
                r['shelf'] = pinned.get(r['text'], r['shelf'])
                r['cat'] = judge(r, r['got'])
        stops, notes = close_verdict(before, after)
        burns = no_check_burns(meta.get('frozen_at', '1970-01-01 00:00:00'))
        for t, a in burns:
            stops.append(f"検査を外して焼いた記録があります（{t} register-addin {a}）＝検査を通してから焼き直してください")
        try:
            xl, owner, text = _connect(getattr(args, 'excel_file', None))
            import vbam_prefire as vp
            chk_ok, chk_lines = vp.shelf_check(xl, owner, text, write_snapshot=False)
            if not chk_ok:
                stops.append("shelf --check が止まります:")
                stops += ["    " + ln for ln in chk_lines[:12]]
            repeat = vp.shelf_misses(all_days=True, entries=vp.shelf_entries_from_text(text))
            repeat_n = (len(repeat[0]), len(repeat[1]))
            exam_n = len(vp.exam_questions())
            notes_stale = stale_note(xl, owner, text)
            if notes_stale:
                notes.append(notes_stale)
        except Exception as e:
            stops.append(f"shelf --check を撃てませんでした: {e}")
            repeat_n, exam_n = None, None
        print(f"帳尻（{pos[1]}）")
        print("  " + _fmt_summary('学習用  前', summarize(before, False)))
        print("  " + _fmt_summary('学習用  後', summarize(after, False)))
        print("  " + _fmt_summary('見せない前', summarize(before, True)))
        print("  " + _fmt_summary('見せない後', summarize(after, True)))
        if repeat_n:
            print(f"  同じ頼みのくり返し外れ（2 回以上・全期間）: 棚に当たらない {repeat_n[0]} 件／今は当たる {repeat_n[1]} 件")
        if exam_n is not None:
            print(f"  問題集: {exam_n} 問")
        for n in notes:
            print("  見る点: " + n)
        ledger_append(pos[1], 'close' if not stops else 'close(止まり)', summarize(after), summarize(after, True),
                      note='; '.join(s for s in stops[:1]))
        if stops:
            print("止まります:")
            for s in stops:
                print("  - " + s)
            return False
        print("通りました（この回の帳尻は ledger に残しました）")
        return True

    print(f"知らない段です: {sub}（shelf --train で手順を見る）")
    return False
