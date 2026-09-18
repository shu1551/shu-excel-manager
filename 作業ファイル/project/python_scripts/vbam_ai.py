# -*- coding: utf-8 -*-
"""vbam_ai.py — vba_manager 分割パート: AI に聞く（Gemini・Claude・ヘッドレス Claude Code）と進み具合・請求書

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。会話の履歴は道具が持ち、往復ごとに請求書を出す。
vbam_agent は `from vbam_ai import *` で名前を引き継ぐ。
"""
import os
import re
import json
import time
import contextlib
import threading
import subprocess
import shutil
import queue
import atexit
import urllib.error
from vbam_core import (SCRIPT_DIR)
from vbam_build import (_AI_DEFAULT_MODEL, _AI_TIMEOUT, _api_call_with_retry)
from vbam_keys import (_key_load)

_AGENT_PROGRESS_FILE = os.path.join(SCRIPT_DIR, '_agent_progress.json')    # いま何をしているか（最新の 1 状態だけ）


# ----------------------------------------------------------------
# AI に聞く（会話は道具が持つ。往復ごとに請求書）
# ----------------------------------------------------------------

def _history_parts(item):
    """history の 1 つ → (role, text, 画像のパス or None)。画像なしの 2 つ組も受ける。"""
    role, text = item[0], item[1]
    return role, text, (item[2] if len(item) > 2 else None)


def _png_b64(path):
    import base64
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode('ascii')


def _chat_gemini(history, model, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    contents = []
    for item in history:
        r, t, img = _history_parts(item)
        parts = [{"text": t}]
        if img and r == 'user':
            parts.append({"inline_data": {"mime_type": "image/png", "data": _png_b64(img)}})
        contents.append({"role": "user" if r == 'user' else "model", "parts": parts})
    body = {"contents": contents,
            "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}}
    req = urllib.request.Request(url, data=json.dumps(body).encode('utf-8'), method='POST',
                                 headers={"Content-Type": "application/json", "x-goog-api-key": api_key})
    data = json.load(urllib.request.urlopen(req, timeout=_AI_TIMEOUT))
    cand = (data.get("candidates") or [{}])[0]
    text = "".join(p.get("text", "") for p in (cand.get("content") or {}).get("parts", []))
    u = data.get("usageMetadata") or {}
    return text, {'in': u.get('promptTokenCount', 0), 'out': u.get('candidatesTokenCount', 0),
                  'think': u.get('thoughtsTokenCount', 0)}


def _chat_claude(history, model, api_key):
    url = "https://api.anthropic.com/v1/messages"
    messages = []
    for item in history:
        r, t, img = _history_parts(item)
        if img and r == 'user':
            content = [{"type": "text", "text": t},
                       {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                    "data": _png_b64(img)}}]
        else:
            content = t
        messages.append({"role": "user" if r == 'user' else "assistant", "content": content})
    body = {"model": model, "max_tokens": 8192, "system": "返事は JSON だけを返す。前後に文を付けない。",
            "messages": messages}
    req = urllib.request.Request(url, data=json.dumps(body).encode('utf-8'), method='POST',
                                 headers={"Content-Type": "application/json", "x-api-key": api_key,
                                          "anthropic-version": "2023-06-01"})
    data = json.load(urllib.request.urlopen(req, timeout=_AI_TIMEOUT))
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    u = data.get("usage") or {}
    return text, {'in': u.get('input_tokens', 0), 'out': u.get('output_tokens', 0), 'think': 0}


# ----------------------------------------------------------------
# 3 本目の頭（2026-09-07）: Claude Code をヘッドレスで使う＝ claude -p を stream-json で温めて 1 会話 1 プロセス
#   鍵は要らない（Claude Code のログイン＝Max のサブスクで追加課金なし）。道具は持たせない（--tools "" --strict-mcp-config）
#   ＝材料は文で渡す Gemini と同じ土俵。往復ごとに stdin へ user 行を 1 行（ASCII だけ・日本語は \uXXXX）、stdout に
#   増える result 行を読む。文脈はプロセスの中に溜まるので履歴を送り直さない。起動の 2.2 秒は 1 会話 1 回
#   （実測 9/7・haiku: 画像つき 1 通目 5.0 秒・温めた 2 通目 1.2 秒）。cwd は中立フォルダ（CLAUDE.md・フック・記憶を
#   読まない＝7/17 の処方箋）。BOM や壊れた行を 1 行送ると claude が黙って終わる（8/22・コンボで 3 回転んだ）。
#   返事は ```json の柵で包まれて返るので剥ぐ。実射（--fire）の既定も claude-code（2026-09-09。それまで実射だけ
#   従量課金の gemini に落ちていた＝金のかかる頭は名指ししたときだけ使う）。
#   思考は切る（9/7 夜・実測）: ヘッドレスは API 直と違って思考量を道具が指定できず、既定のまま sonnet に渡すと 1 通目に
#   思考 9,000 トークン＝ 141.7 秒で _AI_TIMEOUT の 90 秒に毎回切られた（画像は無関係＝あり 118.6 秒）。--effort low だけでは
#   15〜90 秒と揺れて足りない。--settings '{"alwaysThinkingEnabled":false}' で思考ブロックが消えて画像込み 8.4 秒・出力 191
#   トークン（設定キー＝版で意味が変わらない）。環境変数は使わない: MAX_THINKING_TOKENS=0 は 2.1.88 では思考が消えるが
#   文書上は「上限を外す」＝版で真逆に転ぶ。CLAUDE_CODE_DISABLE_EXTENDED_THINKING=1 は 2.1.88 では効かない（91.7 秒）。
# ----------------------------------------------------------------
_CC_AI = 'claude-code'
_CC_CWD = r'C:\tmp' if os.name == 'nt' else '/tmp'
_CC_SYSTEM = "Reply with JSON only. No prose and no code fences around it."
_CC_ARGS = ['-p', '--input-format', 'stream-json', '--output-format', 'stream-json', '--verbose',
            '--tools', '', '--strict-mcp-config', '--disable-slash-commands', '--setting-sources', 'user',
            '--no-session-persistence', '--effort', 'low', '--settings', '{"alwaysThinkingEnabled":false}']
_CC_EXIT_WAIT = 5                      # stdin を閉じてから自分で終わるのを待つ秒（過ぎたら子孫ごと落とす）
_cc_convs = []                         # 生きている会話。新しい会話を始めるとき・走行の終わりに全部止める


class _CCConv:
    """claude -p 1 本＝1 会話。送った本文と返った本文の並び（seen）で、同じ履歴の続きだけを受け付ける。"""

    def __init__(self, model):
        self.model = model
        self.seen = []
        self.proc = None
        self.lines = None
        self.err = []

    def start(self):
        exe = shutil.which('claude')
        if not exe:
            raise RuntimeError("claude コマンドが見つかりません（Claude Code を入れてログインしてください。"
                               "鍵で使うなら --ai gemini か --ai claude）")
        cmd = [exe] + _CC_ARGS + ['--model', self.model, '--system-prompt', _CC_SYSTEM]
        self.proc = subprocess.Popen(cmd, cwd=_CC_CWD if os.path.isdir(_CC_CWD) else None,
                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.lines = queue.Queue()
        threading.Thread(target=self._pump, args=(self.proc.stdout, self.lines), daemon=True).start()
        threading.Thread(target=self._pump, args=(self.proc.stderr, None), daemon=True).start()

    def _pump(self, stream, q):
        for raw in iter(stream.readline, b''):
            if q is None:
                self.err.append(raw.decode('utf-8', 'replace'))
            else:
                q.put(raw)
        if q is not None:
            q.put(None)

    def alive(self):
        return self.proc is not None and self.proc.poll() is None

    def send(self, content):
        line = json.dumps({"type": "user", "message": {"role": "user", "content": content}}, ensure_ascii=True)
        self.proc.stdin.write((line + "\n").encode('ascii'))
        self.proc.stdin.flush()

    def wait_result(self, timeout):
        deadline = time.time() + timeout
        while True:
            left = deadline - time.time()
            if left <= 0:
                raise RuntimeError(f"claude-code が {timeout} 秒以内に答えませんでした")
            try:
                raw = self.lines.get(timeout=left)
            except queue.Empty:
                raise RuntimeError(f"claude-code が {timeout} 秒以内に答えませんでした")
            if raw is None:
                tail = "".join(self.err)[-600:]
                raise RuntimeError("claude-code が答える前に終了しました" + ("\n" + tail if tail.strip() else ""))
            try:
                d = json.loads(raw.decode('utf-8', 'replace'))
            except ValueError:
                continue
            if d.get('type') == 'result':
                return d

    def close(self):
        p, self.proc = self.proc, None
        if p is None:
            return
        with contextlib.suppress(Exception):
            p.stdin.close()
        try:
            p.wait(timeout=_CC_EXIT_WAIT)
        except Exception:
            # claude.CMD の殻（cmd.exe）だけ殺すと中の node が残る＝子孫ごと落とす
            if os.name == 'nt':
                subprocess.run(['taskkill', '/T', '/F', '/PID', str(p.pid)], capture_output=True)
            else:
                with contextlib.suppress(Exception):
                    p.kill()


def _cc_strip_fence(text):
    t = (text or '').strip()
    m = re.match(r'^```[\w-]*[ \t]*\n(.*?)\n?[ \t]*```\s*$', t, re.S)
    return m.group(1).strip() if m else t


def _cc_result_text(d):
    """result 行 → (text, usage)。柵を剥ぎ、トークンは入力＋キャッシュを「入力」に畳む（請求書の形は他の頭と同じ）。"""
    if d.get('is_error'):
        raise RuntimeError("claude-code がエラーを返しました: " + str(d.get('result') or d.get('subtype') or '')[:300])
    u = d.get('usage') or {}
    n_in = u.get('input_tokens', 0) + u.get('cache_read_input_tokens', 0) + u.get('cache_creation_input_tokens', 0)
    n_out = u.get('output_tokens', 0)
    return _cc_strip_fence(str(d.get('result') or '')), {'in': n_in, 'out': n_out, 'think': 0, 'total': n_in + n_out}


def _cc_content(text, img):
    if img:
        return [{"type": "text", "text": text},
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": _png_b64(img)}}]
    return text


def _cc_close_all():
    while _cc_convs:
        _cc_convs.pop().close()


atexit.register(_cc_close_all)


def cc_prewarm(models):
    """走行の頭で claude -p を先に起こしておく（2026-09-11 午後）。

    起動の 2.2 秒は、これまで 1 往復目の待ち（本体）と採点役の待ちの両方に入っていた。models＝要る会話ごとの
    モデル名の並び（本体と採点係。同じモデルが 2 つなら 2 本）。空いている会話（履歴なし）は数に入れる。
    Popen は即返るので、node の起動は AI が考えている裏で進む。戻り値＝新しく起こした本数。
    """
    need = {}
    for m in models or []:
        need[m] = need.get(m, 0) + 1
    started = 0
    for m, n in need.items():
        idle = sum(1 for c in _cc_convs if c.model == m and c.alive() and not c.seen)
        for _ in range(max(0, n - idle)):
            try:
                c = _CCConv(m)
                c.start()
            except Exception:
                break
            _cc_convs.append(c)
            started += 1
    return started


def cc_close_used():
    """使った会話（履歴のあるもの・死んだもの）を閉じ、空の会話は次の走行のために残す（走行の終わりに呼ぶ）。"""
    keep = []
    while _cc_convs:
        c = _cc_convs.pop()
        if c.seen or not c.alive():
            c.close()
        else:
            keep.append(c)
    _cc_convs.extend(reversed(keep))


def _chat_claude_code(history, model, api_key=None):
    """history の未送信ぶんを温めた claude -p に送り、(text, usage) を返す（_chat_gemini／_chat_claude と同じ形）。"""
    texts = [item[1] for item in history]
    # 同じ履歴の続きを受け付ける会話のうち、いちばん長く続いているものを使う（空の会話＝先に起こした予備は
    # どの履歴にも合うので、本体の続きが予備に流れないよう長い方を採る・2026-09-11 午後）
    cands = [c for c in _cc_convs if c.model == model and c.alive()
             and len(c.seen) < len(texts) and texts[:len(c.seen)] == c.seen]
    conv = max(cands, key=lambda c: len(c.seen)) if cands else None
    if conv is None:
        _cc_close_all()                        # 続けられる会話が無い＝全部閉じて起こし直す（居残らせない）
        conv = _CCConv(model)
        conv.start()
        _cc_convs.append(conv)
    pending = history[len(conv.seen):]
    if len(pending) == 1:
        _r, t, img = _history_parts(pending[0])
        content = _cc_content(t, img)
    else:                                      # --continue で持ち込んだ前回の往復など＝1 通にまとめて送る
        parts = []
        for item in pending:
            r, t, _img = _history_parts(item)
            parts.append(("【前回のあなたの返事】\n" if r != 'user' else "") + t)
        content = "\n\n".join(parts)
    conv.send(content)
    text, usage = _cc_result_text(conv.wait_result(_AI_TIMEOUT))
    conv.seen = texts + [text]
    return text, usage


def _cc_oneshot(prompt, model):
    """1 通だけ聞いて閉じる（build の設計図）。"""
    conv = _CCConv(model)
    conv.start()
    try:
        conv.send(prompt)
        return _cc_result_text(conv.wait_result(_AI_TIMEOUT))
    finally:
        conv.close()


_CHAT_FN = {'gemini': _chat_gemini, 'claude': _chat_claude, _CC_AI: _chat_claude_code}


def _ai_setup(ai, model):
    """(ai, model, api_key)。キーが無ければ環境変数名を言って止まる（ネットには出ない）。

    --ai を省いたときの既定は claude-code（ヘッドレスの Claude Code・鍵なし・2026-09-07）。ただし --model が
    gemini で始まる名前なら gemini（Excelコンボの python エンジンが lite／pro で渡す --model がそのまま動く）。
    """
    if not ai:
        ai = 'gemini' if (model or '').lower().startswith('gemini') else _CC_AI
    ai = ai.lower()
    if ai not in _AI_DEFAULT_MODEL:
        raise ValueError(f"--ai は claude-code か gemini か claude（{ai}）")
    if ai == _CC_AI:
        return ai, (model or _AI_DEFAULT_MODEL[ai]), ''      # 鍵は要らない（Claude Code のログイン）
    env = 'GEMINI_API_KEY' if ai == 'gemini' else 'ANTHROPIC_API_KEY'
    api_key = os.environ.get(env) or _key_load(ai)      # 環境変数（昔からの人）→ 金庫（既定）
    if not api_key:
        raise RuntimeError(
            f"{ai} のキーがありません。デスクトップの「APIキーの設定」を開いて入れてください"
            f"（コマンドなら py vba_manager.py set-key {ai}。環境変数 {env} でも動きます）")
    return ai, (model or _AI_DEFAULT_MODEL[ai]), api_key


def _ask(ai, model, api_key, history):
    t0 = time.time()
    try:
        # 429・5xx は待って撃ち直す（2 回まで）。1 回の 503 で往復全体が失敗していた（2026-09-04）
        text, usage = _api_call_with_retry(
            lambda: _CHAT_FN[ai](history, model, api_key), ai)
    except urllib.error.HTTPError as ex:
        detail = ''
        try:
            detail = ex.read().decode('utf-8', 'replace')[:300]
        except Exception:
            pass
        raise RuntimeError(f"{ai} の API がエラーを返しました（HTTP {ex.code}）: {detail}")
    return text, usage, time.time() - t0


def _replay_ask(path):
    """記録（_last_agent_log.jsonl 形式）から実際に _ask を呼んだ往復だけを順に取り出し、
    そのとおりに reply を返す関数を作る（--replay 用・API を呼ばない・課金なし）。

    ループの側（plan の関所・仕上げ検査・自己採点の分岐）だけを無料・数秒で確かめる。
    'resumed': True の行は --continue で持ち込んだ「前回の会話」の再掲であって、
    その回に _ask を呼んだ記録ではないので飛ばす（含めると本数がずれて記録より早く尽きる）。
    """
    replies = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            if 'reply' in d and not d.get('resumed'):
                replies.append(d['reply'])
    i = 0

    def ask(ai, model, api_key, history):
        nonlocal i
        if i >= len(replies):
            raise RuntimeError(f"記録が尽きました（往復 {i}）＝ループがこの記録より長く回っています")
        text = replies[i]
        i += 1
        return text, {'in': 0, 'out': 0, 'think': 0}, 0.0
    return ask


def progress_write(state, **kw):
    """いま何をしているかを 1 ファイルに上書きする（2026-09-05）。書けなくても仕事は止めない。

    CLI は往復ごとに請求書を print するが、MCP 越しだとワーカーが返るまで出力が貯まる＝終わるまで無言。
    最新の 1 状態だけをここに置き、待っている側（agent_status・Excelコンボの AI 作業窓）が読む。
    """
    rec = {'time': time.strftime('%Y-%m-%d %H:%M:%S'), 'state': state, 'pid': os.getpid()}
    rec.update({k: v for k, v in kw.items() if v not in (None, '')})
    rec['line'] = progress_line(rec)
    with contextlib.suppress(Exception):
        with open(_AGENT_PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(rec, f, ensure_ascii=False, indent=1)
    return rec


def _report_head(report, n=80):
    """report の最初の中身の行（【やったこと】のような見出しだけの行は飛ばす・純 Python）。"""
    for line in str(report or '').splitlines():
        s = line.strip()
        if not s or re.fullmatch(r'【[^】]*】[:：]?', s):
            continue
        return s[:n]
    return str(report or '').strip().splitlines()[0][:n] if str(report or '').strip() else ''


def progress_line(rec):
    """進み具合の 1 行（人が読む形）。純 Python＝テストできる。"""
    if not isinstance(rec, dict):
        return ""
    head = {'start': '始めました', 'ask': 'AI に聞いています', 'run': '手を実行しています',
            'gate': '関所を通しています', 'done': '終わりました', 'fail': '失敗しました'}
    parts = [head.get(str(rec.get('state')), str(rec.get('state') or ''))]
    if rec.get('turn'):
        parts.append(f"往復 {rec['turn']}/{rec.get('max_turns', '?')}")
    where = "!".join(str(x) for x in (rec.get('book'), rec.get('sheet')) if x)
    if where:
        parts.append(where)
    for key in ('say', 'hands', 'note'):
        if rec.get(key):
            parts.append(str(rec[key]))
    if rec.get('sec') is not None:
        parts.append(f"{float(rec['sec']):.1f} 秒")
    return " / ".join(p for p in parts if p)


def progress_read(path=None):
    """進み具合のファイルを読む（無ければ None）。"""
    try:
        with open(path or _AGENT_PROGRESS_FILE, encoding='utf-8') as f:
            rec = json.load(f)
        return rec if isinstance(rec, dict) else None
    except (OSError, ValueError):
        return None


def _bill_line(sent, text, usage, sec, model, head="請求書"):
    return (f"{head}: 送り {sent:,}字 / 返り {len(text):,}字 / 待ち {sec:.1f}秒 / "
            f"トークン 入力 {usage['in']:,}・出力 {usage['out']:,}"
            + (f"・思考 {usage['think']:,}" if usage.get('think') else "") + f" / モデル {model}")


__all__ = ['_AGENT_PROGRESS_FILE', '_CCConv', '_CC_AI', '_CC_ARGS', '_CC_CWD', '_CC_EXIT_WAIT', '_CC_SYSTEM', '_CHAT_FN', '_ai_setup', '_ask', '_bill_line', '_cc_close_all', '_cc_content', '_cc_convs', '_cc_oneshot', '_cc_result_text', '_cc_strip_fence', '_chat_claude', '_chat_claude_code', '_chat_gemini', '_history_parts', '_png_b64', '_replay_ask', '_report_head', 'cc_close_used', 'cc_prewarm', 'progress_line', 'progress_read', 'progress_write']
