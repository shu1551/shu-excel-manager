# -*- coding: utf-8 -*-
"""vbam_shake.py — vba_manager 分割パート: 揺らし（記録の返事を Python で書き換え、AI を呼ばずに再生する）

2026-09-11 夜。同じ材料でも AI は違う手を選ぶ（dedupe と row_delete／消した後の番地／書体の落とし／
表示形式の当て忘れ）。その揺れは実射で 1 つずつ見つけていた＝1 回 20 秒〜2 分・使用量つき。
ここでは記録に残った返事（手の JSON）を Python で書き換えた版を作り、--replay と同じ仕組みで
本物の Excel に通す。AI は 1 度も呼ばない。1 通り数秒・費用ゼロ。

  agent --shake [記録のパス] [--only 名前,名前] [--dry-run]
      記録（既定は _last_agent_log.jsonl）の最初の「手のある返事」を種にして、揺らした版を作り、
      開いているブックに 1 通りずつ再生する（ブックは毎回、保存せずに閉じて開き直す＝毎回同じ汚れた姿から）。
      --dry-run は揺らした版を作って一覧を出すだけ（Excel に触らない）。

見る物: 揺らした版ごとに「一発合格／合格（往復 N）／差し戻し→尽きた／不合格」と、原本の再生と同じ表に
なったか（値の違うセルの数）。「差し戻し→尽きた」は、道具がその揺れを吸収できず AI にもう一往復
頼もうとした所＝直す候補。

書き換えの型は、落ちた回から取る（新しく落ちた回が出るたびに 1 つ足す）。知っている形の組み合わせしか
作れない＝まったく新しい揺れは AI からしか出てこない（実射は「直した所を通る弾を数本」の役目で残る）。

状態ファイル（台帳・控え・覚書・記録・時計）は全部この走行の置き場（_agent_shake/日時/_state）へ逃がす。
本物の台帳・覚書には 1 行も書かない（E2E の replay 試験と同じ逃がし方）。
"""
import os
import re
import sys
import json
import copy
import time
import types
import unicodedata
import subprocess
from vbam_core import SCRIPT_DIR

import vbam_agent as va

_AGENT_SHAKE_DIR = os.path.join(SCRIPT_DIR, '_agent_shake')      # 揺らした版・再生の出力・報告の置き場

_CLIP_OPS = ('format', 'normalize', 'find_replace', 'cond_format', 'validation', 'fill')
_RANGE_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$')
_DUP_RE = re.compile(r'行(\d+) = 行(\d+)')
_END_RE = re.compile(r'データの末尾: [A-Z]+(\d+)')
_HEAD_RE = re.compile(r'見出し行の推定: (\d+) 行目（([^）]*)）')


# ----------------------------------------------------------------
# 種（記録から取る）
# ----------------------------------------------------------------

def _load_log(path):
    """記録 → (meta, [往復の行 …])。往復の行は 'reply' を持つもの（採点係・resumed は除く）。"""
    meta, turns = None, []
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            if 'meta' in d and meta is None:
                meta = d['meta']
            elif 'reply' in d and not d.get('resumed') and not d.get('grade'):
                turns.append(d)
    return meta, turns


def _seed_of(path):
    """記録の最初の「手のある返事」と、その材料から読める数字（重複行・末尾・見出し）。"""
    meta, turns = _load_log(path)
    if not meta:
        raise ValueError(f"記録に meta がありません（{path}）")
    if meta.get('mode') != 'sheet':
        raise ValueError(f"揺らせるのは sheet の記録だけです（この記録は「{meta.get('mode') or '不明'}」）")
    seed = None
    for t in turns:
        try:
            d = va._extract_json(t.get('reply') or '')
        except Exception:
            continue
        if isinstance(d, dict) and isinstance(d.get('actions'), list) and d['actions']:
            seed = (t, d)
            break
    if seed is None:
        raise ValueError(f"記録に手のある返事がありません（{path}）")
    t, reply = seed
    prompt = str(t.get('prompt') or '')
    dup_seg = prompt.split('重複行:', 1)[1].split('（計', 1)[0] if '重複行:' in prompt else ''
    dups = [(int(a), int(b)) for a, b in _DUP_RE.findall(dup_seg)]
    m = _END_RE.search(prompt)
    end = int(m.group(1)) if m else None
    if end is None:
        m = re.search(r'使用範囲: [A-Z]+\d+:[A-Z]+(\d+)', prompt)
        end = int(m.group(1)) if m else None
    m = _HEAD_RE.search(prompt)
    header_row = int(m.group(1)) if m else None
    headers = [h for h in re.split(r'[\s　]+', m.group(2).strip()) if h] if m else []
    return {'meta': meta, 'request': str(meta.get('request') or ''), 'sheet': str(meta.get('sheet') or ''),
            'reply': reply, 'turn': int(t.get('turn') or 1), 'dups': dups, 'data_end': end,
            'header_row': header_row, 'headers': headers, 'path': path}


# ----------------------------------------------------------------
# 書き換え（1 つ 1 つ。acts は写し＝壊してよい。当てはまらなければ None）
# ----------------------------------------------------------------

def _rng_parts(s):
    m = _RANGE_RE.match(str(s or '').strip())
    return (m.group(1).upper(), int(m.group(2)), m.group(3).upper(), int(m.group(4))) if m else None


def _with_end(s, end):
    p = _rng_parts(s)
    return f"{p[0]}{p[1]}:{p[2]}{end}" if p else s


def _with_start(s, start):
    p = _rng_parts(s)
    return f"{p[0]}{start}:{p[2]}{p[3]}" if p else s


def _end_of(s):
    p = _rng_parts(s)
    return p[3] if p else None


def _is_dedupe(a):
    return isinstance(a, dict) and str(a.get('op') or '') == 'dedupe'


def _is_delete(a):
    return isinstance(a, dict) and str(a.get('op') or '') in ('dedupe', 'row_delete', 'col_delete')


def _dup_rows(ctx):
    return sorted({d for d, _k in ctx['dups']})


def _replace_dedupe(acts, hands):
    """dedupe の手を hands に置き換える（dedupe が無ければ None）。"""
    out, hit = [], False
    for a in acts:
        if _is_dedupe(a):
            if not hit:
                out.extend(hands)
            hit = True
        else:
            out.append(a)
    return out if hit else None


def _col_index(letter):
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def v_scatter(acts, ctx):
    rows = _dup_rows(ctx)
    return _replace_dedupe(acts, [{"op": "row_delete", "rows": rows, "overwrite": True}]) if rows else None


def v_each_top(acts, ctx):
    rows = _dup_rows(ctx)
    return _replace_dedupe(acts, [{"op": "row_delete", "at": r, "overwrite": True} for r in rows]) if rows else None


def v_each_bottom(acts, ctx):
    rows = _dup_rows(ctx)
    return _replace_dedupe(acts, [{"op": "row_delete", "at": r, "overwrite": True} for r in reversed(rows)]) if rows else None


def v_range_str(acts, ctx):
    rows = _dup_rows(ctx)
    return _replace_dedupe(acts, [{"op": "row_delete", "range": f"{r}:{r}", "overwrite": True}
                                  for r in reversed(rows)]) if rows else None


def v_runs(acts, ctx):
    rows = _dup_rows(ctx)
    if not rows:
        return None
    runs, cur = [], [rows[0]]
    for r in rows[1:]:
        if r == cur[-1] + 1:
            cur.append(r)
        else:
            runs.append(cur)
            cur = [r]
    runs.append(cur)
    return _replace_dedupe(acts, [{"op": "row_delete", "rows": run, "overwrite": True} for run in reversed(runs)])


def v_post_delete(acts, ctx):
    """消す手と同じ返事なのに、書く・整える手の範囲が「消した後の表」の末尾で終わっている（138 秒の回）。"""
    end, rows = ctx.get('data_end'), _dup_rows(ctx)
    if not end or not rows or not any(_is_delete(a) for a in acts):
        return None
    new = end - len(rows)
    changed = False
    for a in acts:
        if not isinstance(a, dict):
            continue
        op = str(a.get('op') or '')
        if op in _CLIP_OPS and _end_of(a.get('range')) == end:
            a['range'] = _with_end(a['range'], new)
            changed = True
        elif op == 'tidy' and isinstance(a.get('ranges'), list):
            a['ranges'] = [_with_end(r, new) if _end_of(r) == end else r for r in a['ranges']]
            changed = True
    return acts if changed else None


def v_to300(acts, ctx):
    end = ctx.get('data_end')
    changed = False
    for a in acts:
        if isinstance(a, dict) and str(a.get('op') or '') in _CLIP_OPS and end and _end_of(a.get('range')) == end:
            a['range'] = _with_end(a['range'], 300)
            changed = True
    return acts if changed else None


def v_drop_font_size(acts, ctx):
    changed = False
    for a in acts:
        if isinstance(a, dict) and a.get('op') == 'format' and a.get('plain') and ('font' in a or 'size' in a):
            a.pop('font', None)
            a.pop('size', None)
            changed = True
    return acts if changed else None


def _drop_numfmt(acts, pred):
    out = [a for a in acts if not (isinstance(a, dict) and a.get('op') == 'format'
                                   and pred(str(a.get('number_format') or '')) and len(a) <= 3)]
    return out if len(out) != len(acts) else None


def v_drop_date_numfmt(acts, ctx):
    return _drop_numfmt(acts, lambda f: 'y' in f.lower() and ('m' in f.lower() or 'd' in f.lower()))


def v_drop_money_numfmt(acts, ctx):
    return _drop_numfmt(acts, lambda f: '#,##0' in f)


def v_no_overwrite(acts, ctx):
    changed = False
    for a in acts:
        if isinstance(a, dict) and 'overwrite' in a:
            a.pop('overwrite')
            changed = True
    return acts if changed else None


def _dedupe_keys(acts, ctx, how):
    for a in acts:
        if not _is_dedupe(a):
            continue
        keys = a.get('keys')
        if how == 'omit':
            if 'keys' not in a:
                return None
            a.pop('keys')
            return acts
        if not isinstance(keys, list) or not keys:
            return None
        headers = ctx.get('headers') or []
        p = _rng_parts(a.get('range')) or ('A', 1, 'A', 1)
        base = _col_index(p[0])
        if how == 'letters':
            new = [_col_letter(base + headers.index(k)) if k in headers else k for k in keys]
        else:
            new = [headers[_col_index(k) - base] if re.fullmatch(r'[A-Za-z]{1,3}', str(k))
                   and 0 <= _col_index(k) - base < len(headers) else k for k in keys]
        if new == keys:
            return None
        a['keys'] = new
        return acts
    return None


def v_keys_letters(acts, ctx):
    return _dedupe_keys(acts, ctx, 'letters')


def v_keys_names(acts, ctx):
    return _dedupe_keys(acts, ctx, 'names')


def v_keys_omit(acts, ctx):
    return _dedupe_keys(acts, ctx, 'omit')


def v_dedupe_from_data(acts, ctx):
    h = ctx.get('header_row')
    for a in acts:
        if _is_dedupe(a) and h and (_rng_parts(a.get('range')) or (None, None))[1] == h:
            a['range'] = _with_start(a['range'], h + 1)
            return acts
    return None


def v_no_tidy(acts, ctx):
    out = [a for a in acts if not (isinstance(a, dict) and a.get('op') == 'tidy')]
    return out if len(out) != len(acts) else None


def v_reverse(acts, ctx):
    return acts[::-1] if len(acts) > 1 else None


def v_dedupe_first(acts, ctx):
    d = [a for a in acts if _is_dedupe(a)]
    if not d or _is_dedupe(acts[0]):
        return None
    return d + [a for a in acts if not _is_dedupe(a)]


def v_dedupe_last(acts, ctx):
    d = [a for a in acts if _is_dedupe(a)]
    rest = [a for a in acts if not _is_dedupe(a)]
    if not d:
        return None
    tidy = [a for a in rest if isinstance(a, dict) and a.get('op') == 'tidy']
    body = [a for a in rest if not (isinstance(a, dict) and a.get('op') == 'tidy')]
    out = body + d + tidy
    return out if out != acts else None


# (名前, 何を揺らしたか, [書き換え …], done を次の往復に分けるか)
_VARIANTS = [
    ("原本", "記録どおり（これが物差し）", [], False),
    ("消す手が飛び飛び", "dedupe を row_delete 1 手に置き換え、rows に重複行を全部並べる（13:50 の 1 往復目の形）", [v_scatter], False),
    ("消す手が1行ずつ上から", "元の行番号のまま上から順に 1 行ずつ消す（9/11 未明の 1 回目＝行ずれの元）", [v_each_top], False),
    ("消す手が1行ずつ下から", "1 行ずつ、下の行から消す", [v_each_bottom], False),
    ("消す手が範囲の文字", "row_delete を \"range\": \"18:18\" の形で（Haiku の形）", [v_range_str], False),
    ("消す手が連番のかたまり", "連番のかたまりごとに rows を並べる（13:50 の 2 往復目の形）", [v_runs], False),
    ("消した後の番地", "消す手と同じ返事で、書く・整える手の範囲が消した後の末尾で終わる", [v_post_delete], False),
    ("138秒の形", "飛び飛びの row_delete ＋ 消した後の番地（138 秒・5 往復・不合格だった回そのもの）", [v_scatter, v_post_delete], False),
    ("1回目の形", "1 行ずつ上から ＋ 消した後の番地", [v_each_top, v_post_delete], False),
    ("範囲が300行まで", "書く・整える手の範囲を 300 行目まで伸ばす（14:17 の形）", [v_to300], False),
    ("書体と大きさを落とす", "書式をそろえる手から font・size を落とす（13:59 の形）", [v_drop_font_size], False),
    ("日付の表示形式を当てない", "日付の列の表示形式の手を抜く（14:03 の形）", [v_drop_date_numfmt], False),
    ("金額の表示形式を当てない", "#,##0 の手を抜く（tidy が拾うか）", [v_drop_money_numfmt], False),
    ("overwriteを書かない", "全部の手から overwrite を落とす（承認の門が通すか）", [v_no_overwrite], False),
    ("keysが列の字", "dedupe の keys を A・B の列の字で（14:17 の形）", [v_keys_letters], False),
    ("keysが見出しの名前", "dedupe の keys を見出しの名前で（18:35 の形）", [v_keys_names], False),
    ("keysなし", "dedupe の keys を省く（材料の重複行と同じ照合になるか）", [v_keys_omit], False),
    ("dedupeが見出しの下から", "dedupe の範囲を見出しの下の行から（Haiku の形）", [v_dedupe_from_data], False),
    ("tidyなし", "仕上げの手を書かない（仕上げ検査が拾うか）", [v_no_tidy], False),
    ("手の順が逆", "手の並びを丸ごと逆に", [v_reverse], False),
    ("dedupeが先頭", "消す手を先頭に置く（道具が後ろへ回すか）", [v_dedupe_first], False),
    ("doneを次の往復で", "同じ手を done なしで返し、次の往復で done だけ返す（道具が待たずに検査へ進むか）", [], True),
]


def _second_turn(reply):
    plan = [dict(p, state='済') if isinstance(p, dict) else p for p in (reply.get('plan') or [])]
    return {"say": "終わりました", "plan": plan, "actions": [], "done": True,
            "report": str(reply.get('report') or '【できなかったこと】\nなし')}


def make_variants(seed, only=None):
    """種 → [{'name', 'why', 'turns': [返事の dict …]} …]。原本は必ず先頭。当てはまらない型は飛ばす。"""
    want = {s.strip() for s in str(only or '').split(',') if s.strip()}
    out = []
    for name, why, ops, split in _VARIANTS:
        if want and name != '原本' and name not in want:
            continue
        acts = copy.deepcopy(seed['reply'].get('actions') or [])
        ok = True
        for op in ops:
            acts = op(acts, seed)
            if acts is None:
                ok = False
                break
        if not ok:
            continue
        reply = copy.deepcopy(seed['reply'])
        reply['actions'] = acts
        if split:
            reply['done'] = False
            out.append({'name': name, 'why': why, 'turns': [reply, _second_turn(seed['reply'])]})
        else:
            reply['done'] = True
            out.append({'name': name, 'why': why, 'turns': [reply]})
    return out


# ----------------------------------------------------------------
# AI に揺れを作らせる（--imagine N・1 往復・Excel は動かさない・2026-09-11 夜）
#   Python の書き換えは「知っている形の組み合わせ」しか作れない。まだ見ていない形は、癖の違う AI に
#   「別の AI ならこの返事をどう書くか」を書かせて広げる。作る AI は表を直さない＝1 往復で N 通り。
#   作った揺れが本物の揺れと同じ分布とは限らない（想像と実際は違う）＝ときどき本物の実射で照らす。
# ----------------------------------------------------------------

_IMAGINE_PROMPT = """あなたは「別の AI がどう返事を書くか」を想像する係です。
下に、Excel の表を直す係への依頼文（規則と材料）と、ある AI が実際に返した返事（JSON）があります。
同じ依頼・同じ材料で、**別の AI が書きそうな返事**を {n} 通り書いてください。
ねらいは、道具がどんな書き方の返事でも同じ結果に着けるかを試すことです。

- 返事の形（キー名・手の名前・引数の形）は規則にある JSON のとおりにする。ただし、手の選び方・順番・範囲の書き方・
  keys の書き方・省く物・手の分け方は、実際の AI が揺れそうな範囲で変える。全部が正しい書き方とは限らない
  ＝迷った AI が書きそうな物（範囲を短く切る・表示形式を忘れる・消す手を行番号で書く・done を別の往復に回す）も入れる。
- 元の返事の写しは出さない。{n} 通りは互いに違う所を 1 つ以上持つ。
- 出力は JSON だけ（前後の文・コードフェンスは付けない）:
  {{"variants": [{{"why": "何を変えたか（20 字以内）", "reply": {{ …返事の JSON… }} }}, …]}}

===== 依頼文（規則と材料） =====
{prompt}

===== ある AI が実際に返した返事 =====
{reply}
"""


def ai_variants(seed, n, ai=None, model=None):
    """AI に n 通りの返事を書かせて、揺らした版にする → ([版 …], 請求書の 1 行, AI の返事の全文)。"""
    n = max(1, min(int(n), 20))
    ai = (ai or va._CC_AI).lower()
    if ai == va._CC_AI and not model:
        model = 'haiku'                      # 頭（sonnet）と癖の違う安い方＝見ていない形が出やすい
    ai, model, key = va._ai_setup(ai, model)
    meta, turns = _load_log(seed['path'])
    prompt = ''
    for t in turns:
        if int(t.get('turn') or 0) == seed['turn']:
            prompt = str(t.get('prompt') or '')
            break
    text = _IMAGINE_PROMPT.format(n=n, prompt=prompt or '（材料なし）',
                                  reply=json.dumps(seed['reply'], ensure_ascii=False, indent=1))
    reply, usage, sec = va._ask(ai, model, key, [('user', text)])
    bill = (f"揺れ作り: 送り {len(text):,}字 / 返り {len(reply or ''):,}字 / 待ち {sec:.1f}秒 / "
            f"トークン 入力 {usage.get('in', 0):,}・出力 {usage.get('out', 0):,} / モデル {model}")
    out = []
    try:
        d = va._extract_json(reply or '')
        items = d.get('variants') if isinstance(d, dict) else d
    except Exception as ex:
        print(f"（AI の返事が JSON として読めません: {ex}）")
        items = []
    for i, it in enumerate(items if isinstance(items, list) else [], 1):
        r = it.get('reply') if isinstance(it, dict) else None
        if not isinstance(r, dict) or not isinstance(r.get('actions'), list) or not r['actions']:
            continue
        why = str(it.get('why') or '').strip()[:60] or '（AI が理由を書いていない）'
        r = dict(r)
        r.setdefault('plan', seed['reply'].get('plan') or [])
        if r.get('done') is False:
            # 次の往復の done は「その版の plan」を済にして返す（種の plan を使うと項目が食い違い、
            # 「未の項目が残っています」で尽きる＝道具でなく版の作り方の問題・2026-09-11 の AI案8）
            out.append({'name': f"AI案{i}", 'why': why, 'turns': [r, _second_turn(r)]})
        else:
            r['done'] = True
            out.append({'name': f"AI案{i}", 'why': why, 'turns': [r]})
    return out, bill, reply or ''


def _variant_jsonl(meta, turns):
    lines = [json.dumps({"meta": dict(meta, model='replay')}, ensure_ascii=False)]
    for i, reply in enumerate(turns, 1):
        lines.append(json.dumps({"turn": i, "prompt": "", "reply": json.dumps(reply, ensure_ascii=False)},
                                ensure_ascii=False))
    return "\n".join(lines) + "\n"


def _safe_name(s):
    return re.sub(r'[\\/:*?"<>|\s]+', '_', str(s))


def _ops_of(turns):
    return " ".join(str(a.get('op') or '?') if isinstance(a, dict) else '?' for a in (turns[0].get('actions') or []))


# ----------------------------------------------------------------
# 状態ファイルを逃がす（E2E の replay 試験と同じ）
# ----------------------------------------------------------------

def _divert_state(state_dir, modules=None):
    """vbam_*／vba_manager が持つ SCRIPT_DIR 配下のパス定数と関数の既定引数を state_dir 配下へ付け替える。
    BACKUP_DIR（控え）は state_dir/backups。modules を渡すとその名前だけ（テスト用）。"""
    real = SCRIPT_DIR
    backups = os.path.join(state_dir, 'backups')
    os.makedirs(backups, exist_ok=True)
    names = modules if modules is not None else [n for n in sys.modules if n.startswith(('vbam_', 'vba_manager'))]
    for name in names:
        mod = sys.modules.get(name)
        if mod is None:
            continue
        for key, val in list(vars(mod).items()):
            if key == 'BACKUP_DIR':
                setattr(mod, key, backups)
            elif isinstance(val, str) and val.startswith(real):
                setattr(mod, key, state_dir + val[len(real):])
            elif isinstance(val, types.FunctionType) and val.__defaults__:
                val.__defaults__ = tuple(state_dir + d[len(real):] if isinstance(d, str) and d.startswith(real) else d
                                         for d in val.__defaults__)
    return backups


# 別プロセスで回す（状態の逃がしがこのプロセスに残らない。MCP サーバーは長生きなので同じプロセスで逃がさない）
_PROBE = r'''# -*- coding: utf-8 -*-
import io, json, os, sys, time, contextlib
sys.path.insert(0, r"{scripts}")
import vbam_core
vbam_core.setup_encoding()
import vba_manager, vbam_agent as va, vbam_ai as vai, vbam_shake as vs   # _run_cmd が後から読む表も先に読む
from vbam_hands import _rows_of

book, vdir, state = sys.argv[1], sys.argv[2], sys.argv[3]
vs._divert_state(state)


def _boom(*a, **k):
    raise RuntimeError("shake: 本物の AI は呼びません")


for m in (va, vai):                      # 本物の AI を拒む見張り（テストの conftest と同じ）
    for n in ("_ask", "_cc_oneshot", "_chat_gemini", "_chat_claude", "_chat_claude_code"):
        if hasattr(m, n):
            setattr(m, n, _boom)
    for n in ("cc_prewarm", "cc_close_used", "_cc_close_all"):
        if hasattr(m, n):
            setattr(m, n, lambda *a, **k: None)

xl, wb = vbam_core.get_workbook(book)
files = sorted(f for f in os.listdir(vdir) if f.endswith(".jsonl"))
for fname in files:
    path = os.path.join(vdir, fname)
    meta = None
    with open(path, encoding="utf-8") as f:
        for line in f:
            if line.strip():
                d = json.loads(line)
                if "meta" in d:
                    meta = d["meta"]
                    break
    name = fname[3:-6]
    row = {{"name": name, "file": fname}}
    t0 = time.time()
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            # 毎回、保存していない元の姿（ファイル）から開き直す。同じ Excel（同じ窓）で開き直す
            # ＝get_workbook に任せると「どこにも開いていない」と見て DispatchEx で見えない Excel を起こす
            app = wb.Application
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = app.Workbooks.Open(book)
            vbam_core._wb_cache.clear()
            xl, wb = vbam_core.get_workbook(book)     # 開いた物を掴み直す（既に開いています）
        sheet = meta.get("sheet") or wb.ActiveSheet.Name
        inner = va._replay_ask(path)
        stop = {{}}

        def ask(ai, model, key, history, _inner=inner, _stop=stop):
            try:
                return _inner(ai, model, key, history)
            except RuntimeError:
                if history:
                    last = history[-1]
                    txt = last[1] if isinstance(last, (list, tuple)) and len(last) > 1 else last
                    _stop["back"] = str(txt)[:20000]      # 差し戻しの「なぜ」は結果の後ろに付く＝末尾まで取る
                raise

        r, err = None, None
        with contextlib.redirect_stdout(buf):
            try:
                r = va.run_agent(meta.get("request") or "", sheet, wb, "gemini", "replay",
                                 int(meta.get("max_turns") or 4), False, grade=False, backup=True, ask=ask)
            except RuntimeError as ex:
                err = str(ex)
            except Exception as ex:
                err = f"{{type(ex).__name__}}: {{ex}}"
        row["err"] = err
        row["back"] = stop.get("back")
        if r is not None:
            row.update({{"ok": bool(r.get("ok")), "done": bool(r.get("done")), "turns": r.get("turns"),
                         "gates": {{k: v for k, v in (r.get("gates") or {{}}).items() if v}},
                         "audit": r.get("audit"), "unmet": r.get("unmet"), "inv": r.get("inv"),
                         "tool_sec": sum((r.get("tool_sec") or {{}}).values())}})
        ws = wb.Sheets(sheet)
        ur = ws.UsedRange
        addr = ur.Address                       # 動的 dispatch では property（文字）・makepy では method
        row["addr"] = str(addr(False, False) if callable(addr) else addr).replace("$", "")
        row["vals"] = [[("" if v is None else str(v)) for v in rr] for rr in _rows_of(ur.Value)]
        ws = ur = None
    except Exception as ex:
        import traceback
        row["err"] = row.get("err") or f"{{type(ex).__name__}}: {{ex}}"
        row["trace"] = traceback.format_exc()[-1500:]
    with open(os.path.join(vdir, fname[:-6] + ".out.txt"), "w", encoding="utf-8") as f:
        f.write(buf.getvalue())
    row["sec"] = round(time.time() - t0, 1)
    print("__ROW__" + json.dumps(row, ensure_ascii=False, default=str), flush=True)

try:
    app = wb.Application
    wb.Close(SaveChanges=False)
    wb = app.Workbooks.Open(book)             # 最後も元の姿で、同じ窓に開いておく（人が見る）
except Exception:
    pass
ws = wb = xl = None
print("__END__", flush=True)
'''


def _w(s):
    return sum(2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1 for ch in str(s))


def _pad(s, n):
    s = str(s)
    return s + ' ' * max(0, n - _w(s))


def _judge(row):
    err = row.get('err') or ''
    if '記録が尽きました' in err:
        return '差し戻し→尽きた'
    if err:
        return '落ちた'
    if row.get('ok'):
        return '一発合格' if row.get('turns') == 1 else f"合格（往復 {row.get('turns')}）"
    return '不合格'


_REASON_RE = re.compile(r'実行していません: (?!同じ返事の中に)|合格にしません|仕上げ検査|頼んでいない変化|未の項目|読み残し|'
                        r'残っています|残った|直っていません|承認の言葉|事前の門|差し戻し')


def _reason(row):
    """判定の理由を 1 行で（差し戻し文の中の「なぜ」・仕上げ検査の指摘・関所の名札）。"""
    back = row.get('back') or ''
    if back:
        lines = [ln.strip() for ln in back.splitlines() if ln.strip()]
        hit = [ln for ln in lines if _REASON_RE.search(ln)]
        pick = hit[0] if hit else (lines[-1] if lines else back)
        pick = re.sub(r'^--- \d+: \S+ （失敗）---\s*', '', pick)
        return pick.replace('実行していません: ', '')[:110]
    g = row.get('gates') or {}
    if g:
        return "関所: " + "・".join(f"{va._GATE_LABELS.get(k, k)}×{v}" for k, v in g.items())
    for key in ('audit', 'unmet', 'inv'):
        v = row.get(key)
        if v:
            first = v[0] if isinstance(v, (list, tuple)) else v
            return str(first)[:90]
    if row.get('err'):
        return str(row['err'])[:90]
    return ''


def _compare(base, row):
    """原本の再生の表と比べる → '同じ'／'N セル違う'／'行 a→b'。"""
    if base is None or 'vals' not in row or 'vals' not in base:
        return '-'
    a, b = base['vals'], row['vals']
    if len(a) != len(b) or (a and b and len(a[0]) != len(b[0])):
        return f"行 {len(a)}→{len(b)}・列 {len(a[0]) if a else 0}→{len(b[0]) if b else 0}"
    n = sum(1 for ra, rb in zip(a, b) for x, y in zip(ra, rb) if x != y)
    return '同じ' if n == 0 else f"{n} セル違う"


def shake(log_path=None, target_file=None, dry_run=False, only=None, timeout=1800, imagine=0, ai=None, model=None):
    """agent --shake の本体。戻り値は「揺らした版が全部、原本と同じ表に一発で着いたか」。
    imagine=N なら AI に N 通り書かせた版も足す（1 往復。ai・model はその AI）。"""
    log_path = log_path or va._LAST_AGENT_LOG_FILE
    if not os.path.isfile(log_path):
        print(f"エラー: 揺らす記録がありません（{log_path}）")
        return False
    try:
        seed = _seed_of(log_path)
    except ValueError as ex:
        print(f"エラー: {ex}")
        return False
    variants = make_variants(seed, only)
    stamp = time.strftime('%Y%m%d_%H%M%S')
    vdir = os.path.join(_AGENT_SHAKE_DIR, stamp)
    os.makedirs(vdir, exist_ok=True)
    if imagine:
        try:
            extra, bill, raw = ai_variants(seed, imagine, ai, model)
        except Exception as ex:
            print(f"エラー: AI に揺れを作らせられませんでした: {ex}")
            return False
        with open(os.path.join(vdir, '_imagine_reply.txt'), 'w', encoding='utf-8') as f:
            f.write(raw)
        print(bill + f" → 使える版 {len(extra)} 通り（AI の返事の全文は {os.path.join(vdir, '_imagine_reply.txt')}）")
        variants += extra
    for i, v in enumerate(variants):
        with open(os.path.join(vdir, f"{i:02d}_{_safe_name(v['name'])}.jsonl"), 'w', encoding='utf-8') as f:
            f.write(_variant_jsonl(seed['meta'], v['turns']))
    print(f"種: {log_path}（往復 {seed['turn']} の返事・手 {len(seed['reply'].get('actions') or [])} 本）")
    print(f"  ブック {seed['meta'].get('book')} / シート {seed['sheet']} / 重複 {len(seed['dups'])} 行 / "
          f"末尾 {seed['data_end']} 行目 / 見出し {seed['header_row']} 行目")
    print(f"揺らした版 {len(variants)} 通り（置き場 {vdir}）:")
    for i, v in enumerate(variants):
        print(f"  {i:02d} {_pad(v['name'], 26)} {v['why']}")
        print(f"     {'手 ' + str(len(v['turns'][0].get('actions') or [])) + ' 本: ' + _ops_of(v['turns'])}"
              + ("　＋ 次の往復で done" if len(v['turns']) > 1 else ""))
    if dry_run:
        print("（--dry-run: 揺らした版を作って一覧を出しただけ。Excel には触っていません）")
        return True
    # 再生する相手（開いているブック＝記録のブックと同じ名前であること）
    try:
        from vbam_core import get_workbook
        xl, wb = get_workbook(target_file)
        if str(wb.Name) != str(seed['meta'].get('book')):
            print(f"エラー: 開いているブック（{wb.Name}）が記録のブック（{seed['meta'].get('book')}）と違います。"
                  "記録のブックを開いてから、または名指しして撃ってください")
            return False
        book = str(wb.FullName)
        if not str(wb.Path or ''):
            print("エラー: 保存していないブックは開き直せません（揺らした版ごとに元の姿へ戻せない）")
            return False
        others = [str(w.Name) for w in xl.Workbooks if str(w.Name) != str(wb.Name)]
        if others:
            print(f"（同じ Excel に他のブックも開いています: {', '.join(others)}。触るのは {wb.Name} だけです）")
        wb = xl = None
    except Exception as ex:
        print(f"エラー: {ex}")
        return False
    state = os.path.join(vdir, '_state')
    os.makedirs(state, exist_ok=True)
    probe = os.path.join(vdir, '_probe.py')
    with open(probe, 'w', encoding='utf-8') as f:
        f.write(_PROBE.format(scripts=SCRIPT_DIR))
    print(f"再生: {book} に {len(variants)} 通りを順に（毎回、保存せずに閉じて開き直す。AI は呼ばない。"
          f"台帳・控え・覚書は {state} へ）")
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    proc = subprocess.Popen([sys.executable, probe, book, vdir, state], stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace', env=env)
    rows, t0, ended = [], time.time(), False
    try:
        for line in proc.stdout:
            line = line.rstrip('\n')
            if line.startswith('__ROW__'):
                row = json.loads(line[7:])
                rows.append(row)
                base = rows[0] if rows and rows[0]['name'] == '原本' else None
                print(f"  {_pad(row['name'], 26)} {_pad(_judge(row), 16)} {_pad(_compare(base, row), 14)} "
                      f"{row.get('sec', 0):5.1f} 秒  {_reason(row)}", flush=True)
            elif line.startswith('__END__'):
                ended = True
            elif line.strip():
                print("    | " + line, flush=True)
            if time.time() - t0 > timeout:
                proc.kill()
                print(f"エラー: {timeout} 秒を超えたので止めました")
                break
    finally:
        try:
            proc.wait(timeout=30)
        except Exception:
            proc.kill()
    if not ended and not rows:
        print("エラー: 再生の子プロセスが 1 通りも返しませんでした（出力は上の行）")
        return False
    base = rows[0] if rows and rows[0]['name'] == '原本' else None
    one = [r for r in rows if _judge(r) == '一発合格']
    same = [r for r in rows if _compare(base, r) in ('同じ', '-')]
    bad = [r for r in rows if _judge(r) != '一発合格' or _compare(base, r) not in ('同じ', '-')]
    print(f"\nまとめ: {len(rows)} 通りのうち一発合格 {len(one)}・原本と同じ表 {len(same)}"
          + (f"・直す候補 {len(bad)}（{'、'.join(r['name'] for r in bad)}）" if bad else "・直す候補なし"))
    report = {'seed': {k: v for k, v in seed.items() if k != 'reply'}, 'stamp': stamp, 'dir': vdir,
              'rows': [dict(r, judge=_judge(r), same=_compare(base, r)) for r in rows]}
    with open(os.path.join(vdir, 'report.json'), 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=1, default=str)
    print(f"報告: {os.path.join(vdir, 'report.json')}（各版の画面出力は同じ置き場の *.out.txt）")
    return not bad


__all__ = ['_AGENT_SHAKE_DIR', '_VARIANTS', '_divert_state', '_seed_of', 'ai_variants', 'make_variants', 'shake']
