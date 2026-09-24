# -*- coding: utf-8 -*-
"""vbam_ledger.py — vba_manager 分割パート: 本番の走行台帳（1 走行 1 行・戻した印・保存した印・記録の控え）

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。vbam_agent は `from vbam_ledger import *` で
名前を引き継ぐ。記録ファイル（_LAST_AGENT_LOG_FILE）の場所は呼ぶ瞬間に vbam_agent から引く。
"""
import os
import json
import time
import contextlib
from vbam_core import (SCRIPT_DIR)

_AGENT_RUNS_FILE = os.path.join(SCRIPT_DIR, '_agent_runs.jsonl')     # 本番の走行台帳（1 走行 1 行・2026-09-06）
_FIRED_FILE = os.path.join(SCRIPT_DIR, '_fired.jsonl')              # 先撃ちで撃ったマクロ（1 走行 1 行・2026-09-18 夜）
_AGENT_LOGS_DIR = os.path.join(SCRIPT_DIR, '_agent_logs')            # 走行ごとの記録の控え（--replay で撃ち直せる）
_AGENT_LOGS_KEEP = 40        # 残す記録の本数（新しい方から）


# ----------------------------------------------------------------
# 本番の走行台帳（2026-09-06）: 1 走行 1 行。実射（練習台）とは別のファイルに残す。
#   実射の点数は「私が作った問題の点数」だが、ここに溜まるのは実戦の記録。
#   done 率よりも **undo 率**（人が戻した割合）が、本番でいちばん正直な失敗率になる。
#   記録（往復の全文）も走行ごとに控えるので、後から `agent --replay 記録のパス` で撃ち直せる。
# ----------------------------------------------------------------

_LAST_RUN_ID = ['', 0]


def _run_id():
    """走行の名札＝日時。同じ秒に続けて撃つと名札がぶつかり、控えも台帳も前の走行を上書きした
    （2026-09-23 実測 3・shelf-run 2 本が 184635 に重なった）ので、2 本目からは _2・_3 を付ける。"""
    rid = time.strftime('%Y%m%d_%H%M%S')
    if rid == _LAST_RUN_ID[0]:
        _LAST_RUN_ID[1] += 1
        return f"{rid}_{_LAST_RUN_ID[1]}"
    _LAST_RUN_ID[0], _LAST_RUN_ID[1] = rid, 1
    return rid


def _save_run_log(run_id, mode):
    """その走行の記録（_last_agent_log.jsonl）を控えて、パスを返す。取れなければ None。"""
    from vbam_agent import (_LAST_AGENT_LOG_FILE)   # 分割後の遅延 import（循環にしない・2026-09-11）
    try:
        os.makedirs(_AGENT_LOGS_DIR, exist_ok=True)
        dst = os.path.join(_AGENT_LOGS_DIR, f"{run_id}_{mode}.jsonl")
        with open(_LAST_AGENT_LOG_FILE, 'r', encoding='utf-8') as src, \
                open(dst, 'w', encoding='utf-8') as out:
            out.write(src.read())
    except OSError:
        return None
    try:                                   # 古い記録から間引く（溜め込まない）
        keep = sorted(f for f in os.listdir(_AGENT_LOGS_DIR) if f.endswith('.jsonl'))
        for f in keep[:-_AGENT_LOGS_KEEP]:
            with contextlib.suppress(OSError):
                os.remove(os.path.join(_AGENT_LOGS_DIR, f))
    except OSError:
        pass
    return dst


def runs_append(rec):
    """走行台帳に 1 行足す（書けなくても仕事は止めない）。"""
    try:
        with open(_AGENT_RUNS_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(rec, ensure_ascii=False) + chr(10))
    except OSError as ex:
        print(f"（走行台帳に書けませんでした: {ex}）")


def _runs_load(path=None):
    """台帳（古い順）。'undo' の行は走行に畳んでから返す（append-only のまま印を付けられる）。"""
    rows, undone = [], {}
    try:
        with open(path or _AGENT_RUNS_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    d = json.loads(line)
                except ValueError:
                    continue
                if d.get('undo'):
                    undone[d['undo']] = d.get('time')
                else:
                    rows.append(d)
    except OSError:
        return []
    for r in rows:
        if r.get('run_id') in undone:
            r['undone'] = undone[r['run_id']]
        saved = _saved_after(r)
        if saved:
            r['saved'] = saved
    return rows


def _saved_after(rec):
    """走行の後にそのブックが保存されたか（ファイルの更新日時が走行より新しい）。分からなければ None（2026-09-06）。

    undo 率は「人が戻した」＝失敗の側の人の判断。こちらは「人が保存した」＝受け入れの側の人の判断。
    道具は保存しないので、走行より新しい更新日時は人が保存したということ。行に 'path' が無い
    （未保存のブック・macro の走行＝道具が自分で保存する）ときは数えない。
    """
    p = rec.get('path')
    if not p or not os.path.isfile(p):
        return None
    try:
        t_run = time.mktime(time.strptime(str(rec.get('time') or ''), '%Y-%m-%d %H:%M:%S'))
        mt = os.path.getmtime(p)
    except (ValueError, OSError, OverflowError):
        return None
    if mt > t_run + 1:
        return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mt))
    return None


_REWORK_MINUTES = 30          # done の後この分数のうちに同じシートへもう一度走行したら「手直し」と数える


def _rework_count(rows, minutes=_REWORK_MINUTES):
    """done の後 N 分以内に同じブック・同じシートへもう一度走行した数（人が手直しに回った目安・純 Python）。

    2026-09-06 夜（Claude for Excel の助言の指標 4 つのうち「done 後 N 分以内の手直し率」）。
    undo は「戻した」、これは「戻さずにやり直させた」＝どちらも人が結果を受け入れなかった側の数字。
    """
    def _t(r):
        try:
            return time.mktime(time.strptime(str(r.get('time') or ''), '%Y-%m-%d %H:%M:%S'))
        except (ValueError, OverflowError):
            return None
    n = 0
    for i, r in enumerate(rows):
        t0 = _t(r)
        if t0 is None or not r.get('done'):
            continue
        key = (r.get('book'), r.get('sheet'))
        for later in rows[i + 1:]:
            if (later.get('book'), later.get('sheet')) != key:
                continue
            t1 = _t(later)
            if t1 is not None and 0 <= t1 - t0 <= minutes * 60:
                n += 1
            break
    return n


def _runs_summary(rows):
    """台帳 → 実運用の数字（走行・done 率・undo 率・一発率・保存された率・聞き返し率・手直し率）。行が無ければ空文字。"""
    from vbam_agent import (_gate_total)   # 分割後の遅延 import（循環にしない・2026-09-11）
    if not rows:
        return ""
    n = len(rows)
    done = sum(1 for r in rows if r.get('done'))
    undo = sum(1 for r in rows if r.get('undone'))
    first = sum(1 for r in rows if r.get('done') and _gate_total(r.get('gates')) == 0)
    known = [r for r in rows if r.get('path')]
    saved = sum(1 for r in known if r.get('saved') and not r.get('undone'))
    asks = sum(1 for r in rows if r.get('asks'))
    rework = _rework_count(rows)
    by_mode = {}
    for r in rows:
        by_mode[r.get('mode') or '?'] = by_mode.get(r.get('mode') or '?', 0) + 1
    return (f"実運用: {n} 走行（" + " / ".join(f"{m} {c}" for m, c in sorted(by_mode.items())) + "）"
            f"　done {done}/{n}（{100.0 * done / n:.0f}%）"
            f"　人が戻した（undo）{undo}/{n}（{100.0 * undo / n:.0f}%）"
            f"　一発 {first}/{n}（{100.0 * first / n:.0f}%）"
            + (f"　人が保存した（受け入れ）{saved}/{len(known)}（{100.0 * saved / len(known):.0f}%）" if known else "")
            + f"　聞き返し {asks}/{n}（{100.0 * asks / n:.0f}%）"
            + f"　{_REWORK_MINUTES} 分以内の手直し {rework}/{n}（{100.0 * rework / n:.0f}%）")


def fired_append(run_id, book, sheet, request, names, owners=()):
    """先撃ちで撃ったマクロの名前を残す（2026-09-18 夜: shu「作りすぎは消す」→ 使用の記録で引退を決めるための台帳。
    走行台帳（_agent_runs.jsonl）はマクロだけで終わった走行しか macro_first を持たず、AI に続いた走行では名前が消えていた）。"""
    rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'run_id': run_id, 'book': book, 'sheet': sheet,
           'request': (request or '')[:200], 'macros': list(names or []), 'owners': sorted(set(owners or ()))}
    try:
        with open(_FIRED_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(rec, ensure_ascii=False) + chr(10))
    except OSError as ex:
        print(f"（先撃ちの台帳に書けませんでした: {ex}）")
    return rec


def fired_counts(path=None):
    """先撃ちの台帳 → {Sub 名: (回数, 最後に撃った時刻)}（純 Python）。無ければ {}。"""
    out = {}
    try:
        with open(path or _FIRED_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except ValueError:
                    continue
                for nm in rec.get('macros') or []:
                    n, _t = out.get(nm, (0, ''))
                    out[nm] = (n + 1, rec.get('time') or '')
    except OSError:
        pass
    return out


def fired_summary(registry_names=(), limit=12):
    """先撃ちの回数の要約（agent --runs の末尾）。registry_names＝登録簿の Sub 名（撃たれていない本を並べる）。"""
    counts = fired_counts()
    lines = []
    if counts:
        top = sorted(counts.items(), key=lambda kv: (-kv[1][0], kv[0]))[:limit]
        lines.append(f"先撃ちの回数（{_FIRED_FILE}・多い順）:")
        lines += [f"  {n:4d} 回  {nm}（最後 {t}）" for nm, (n, t) in top]
    names = [n for n in registry_names if n not in counts]
    if registry_names:
        lines.append(f"一度も撃たれていない登録簿の Sub: {len(names)} 本" + (f"（{'・'.join(names[:10])}" + ("…" if len(names) > 10 else "") + "）" if names else ""))
    return chr(10).join(lines) if lines else "先撃ちの記録がありません（本番で登録簿のマクロが撃たれると 1 行ずつ残ります）"


def runs_history(limit=15):
    """本番の走行台帳を出す（agent --runs）。"""
    from vbam_agent import (_gate_line)   # 分割後の遅延 import（循環にしない・2026-09-11）
    rows = _runs_load()
    if not rows:
        print(f"走行の記録がありません（{_AGENT_RUNS_FILE}）。agent を本番のブックに回すと 1 行ずつ残ります。")
        return True
    print(f"本番の走行（新しい順・{_AGENT_RUNS_FILE}）:")
    for r in list(reversed(rows))[:limit]:
        mark = "done" if r.get('done') else ("落とし物" if r.get('dropped') else "未完")
        gl = _gate_line(r.get('gates'), head="関所")
        print(f"  {r.get('time', '?')}  [{r.get('mode', '?')}] {r.get('book', '?')}"
              + (f"!{r['sheet']}" if r.get('sheet') else "")
              + f"  {mark}・往復 {r.get('turns', '?')}・{r.get('sec', 0):.0f}秒"
              + ("  ⚠ 人が戻した" if r.get('undone') else
                 (f"  保存された（受け入れ・{r['saved']}）" if r.get('saved') else "")))
        print(f"      依頼「{(r.get('request') or '')[:60]}」  {gl}"
              + (f"  記録: {os.path.basename(r['log'])}" if r.get('log') else ""))
    print(_runs_summary(rows))
    try:
        # 登録簿の Sub 名は台帳（鍛えた仕事）から＝Excel を起こさない。引退した仕事は数えない
        import vbam_forge as _vf
        names = [v.get('sub') for v in _vf._forge_load().values() if v.get('sub') and v.get('passed') and not v.get('retired')]
    except Exception:
        names = []
    print(fired_summary(names))
    print("（撃ち直すなら agent --replay " + os.path.join(_AGENT_LOGS_DIR, "記録の名前.jsonl") + "）")
    return True


def _book_path(wb):
    """保存済みブックのフルパス（未保存・読めなければ None）。台帳が「その後に人が保存したか」を見るために持つ。"""
    try:
        if not str(wb.Path or ''):
            return None
        return str(wb.FullName)
    except Exception:
        return None


def _runs_record(run_id, mode, book, sheet, request, r, sec, save_log=True, path=None):
    """run_agent／run_macro_agent／run_build から呼ぶ共通の記録（1 走行 1 行）。

    save_log=False は記録の控えを取らない（build が直しのループに回らなかったときは
    _last_agent_log.jsonl が前の仕事のもの＝それを控えると別の走行の記録が混ざる）。
    path＝そのブックのフルパス（保存済みのときだけ）。台帳を読むときに更新日時と照らして
    「人が保存した（受け入れ）」の印を付ける（2026-09-06）。
    """
    rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'run_id': run_id, 'mode': mode,
           # 依頼文は全文（前は 120 字で切っていて、台帳を読んでも何を頼んだか分からなかった・2026-09-06 夜）
           'book': book, 'sheet': sheet, 'request': (request or ''), 'path': path,
           'turns': r.get('turns'), 'done': bool(r.get('done')), 'ok': bool(r.get('ok')),
           'gates': {k: v for k, v in (r.get('gates') or {}).items() if v},
           'dropped': len(r.get('dropped') or []), 'audit': len(r.get('audit') or []),
           'inv': len(r.get('inv') or []), 'unmet': len(r.get('unmet') or []),
           'asks': len(r.get('asks') or []),          # 聞き返し（要判断で人に返した数・2026-09-06 夜）
           'merged': int(r.get('merged') or 0),       # 手と done を同時に受けて往復を省いた回数（2026-09-06 夜）
           'sec': round(sec, 1), 'in': (r.get('usage') or {}).get('in', 0),
           'out': (r.get('usage') or {}).get('out', 0),
           'log': _save_run_log(run_id, mode) if save_log else None}
    runs_append(rec)
    return rec


__all__ = ['_AGENT_LOGS_DIR', '_AGENT_LOGS_KEEP', '_AGENT_RUNS_FILE', '_FIRED_FILE', 'fired_append', 'fired_counts', 'fired_summary', '_REWORK_MINUTES', '_book_path', '_rework_count', '_run_id', '_runs_load', '_runs_record', '_runs_summary', '_save_run_log', '_saved_after', 'runs_append', 'runs_history']
