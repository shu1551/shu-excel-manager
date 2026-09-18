# -*- coding: utf-8 -*-
"""vbam_keys.py — vba_manager 分割パート: API キーの預かり（set-key / clear-key・金庫・期限・疎通）

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。vbam_agent は `from vbam_keys import *` で
名前を引き継ぐ（呼ぶ側のコードは変えない）。key_gui.py もここから import する。
"""
import os
import sys
import json
import time
import datetime
import subprocess
import urllib.error
from vbam_build import (_AI_DEFAULT_MODEL)


# ----------------------------------------------------------------
# APIキーの預かり（2026-09-04）: 設定・確認・削除を道具が肩代わりする
#   人に「システムの詳細設定 → 環境変数 → 新規」をやらせない。set-key 一本で、預かって・通して・言う。
#   キーは金庫（_KEY_STORE。Windows のログインで暗号化＝DPAPI・期限つき）に預かり、いまのプロセスの
#   環境変数にも入れる（窓を閉じれば消える）。古い平文のユーザー環境変数も読めるが、書くのは金庫だけ。
#   画面にも記録にも全文は出さない（先頭 4 字と末尾 4 字だけ）。
# ----------------------------------------------------------------
_KEY_ENV = {'gemini': 'GEMINI_API_KEY', 'claude': 'ANTHROPIC_API_KEY'}
_KEY_STORE = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'),
                          'vba-manager', 'keys.json')
_KEY_DAYS_DEFAULT = 30       # 既定の保管期間（切れたら道具が自分で消す）


# ---- 金庫（2026-09-04）: 環境変数に平文で置きっぱなしにしない ----
#   環境変数は「平文・永久・そのアカウントで動く全プログラムから読める」。人に「後で消してください」と
#   頼む設計は、危険を人の記憶に預けているだけ。だから Windows の DPAPI で暗号にして自分の場所に置き、
#   期限を持たせて、切れたら道具が自分で消す。他のアカウント・他のパソコンでは戻せない（鍵は Windows のログイン）。

def _store_read():
    try:
        with open(_KEY_STORE, encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (FileNotFoundError, ValueError, OSError):
        return {}


def _store_write(d):
    os.makedirs(os.path.dirname(_KEY_STORE), exist_ok=True)
    with open(_KEY_STORE, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


def _key_save(ai, key, days=_KEY_DAYS_DEFAULT):
    """キーを暗号にして保管する。days=0 なら期限なし。戻り値＝期限（文字列 or None）。"""
    import base64
    import win32crypt
    blob = win32crypt.CryptProtectData(key.encode('utf-8'), 'vba-manager api key', None, None, None, 0)
    exp = None
    if days:
        exp = (datetime.datetime.now() + datetime.timedelta(days=int(days))).strftime('%Y-%m-%d %H:%M')
    d = _store_read()
    d[ai] = {'blob': base64.b64encode(blob).decode('ascii'),
             'saved': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
             'expires': exp, 'tail': key[-4:]}
    _store_write(d)
    return exp


def _key_expired(rec):
    exp = (rec or {}).get('expires')
    if not exp:
        return False
    try:
        return datetime.datetime.strptime(exp, '%Y-%m-%d %H:%M') < datetime.datetime.now()
    except ValueError:
        return False


def _key_load(ai):
    """保管したキーを取り出す。無い・期限切れ・他のアカウントなら None（期限切れはその場で消す）。"""
    import base64
    rec = _store_read().get(ai)
    if not rec or not rec.get('blob'):
        return None
    if _key_expired(rec):
        _key_forget(ai)
        return None
    try:
        import win32crypt
        return win32crypt.CryptUnprotectData(
            base64.b64decode(rec['blob']), None, None, None, 0)[1].decode('utf-8')
    except Exception:
        return None          # 別のアカウント・別のパソコンでは戻せない＝それが狙い


def _key_forget(ai):
    d = _store_read()
    if ai in d:
        d.pop(ai)
        _store_write(d)
        return True
    return False


def _key_info(ai):
    """(ある？, 見せる文字列)。期限切れも「切れています」と言う。"""
    rec = _store_read().get(ai)
    if not rec:
        return False, None
    tail = rec.get('tail') or ''
    if _key_expired(rec):
        return False, f"期限切れ（…{tail} / {rec.get('expires')} まででした）"
    exp = rec.get('expires')
    return True, (f"保管中（…{tail} / {exp} まで）" if exp else f"保管中（…{tail} / 期限なし）")
_KEY_ISSUE_URL = {'gemini': 'https://aistudio.google.com/apikey',
                  'claude': 'https://console.anthropic.com/settings/keys'}
_KEY_MAX = 1024          # 貼り付け事故を見つけるための長さの上限（キーはこれよりずっと短い）
_PING_ASK = '1+1 は？ 数字だけ返してください。'


def _mask_key(k):
    """キーを画面に出す形（先頭 4 字と末尾 4 字だけ）。"""
    if not k:
        return '（設定されていません）'
    k = str(k)
    if len(k) <= 8:
        return '*' * len(k)
    return k[:4] + '*' * (len(k) - 8) + k[-4:]


def _user_env_get(name):
    """ユーザー環境変数の現物（レジストリ）を読む。プロセスの写しではなく、保存されている値。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment') as k:
            return winreg.QueryValueEx(k, name)[0]
    except FileNotFoundError:
        return None
    except OSError:
        return None


def _user_env_write(name, value):
    """ユーザー環境変数を書く／消す（value=None で消す）。
    setx は使わない（setx NAME "" は空の値を残すだけで、変数そのものは消えない）。
    .NET の SetEnvironmentVariable なら項目ごと消え、開いている窓にも変更が知らされる。
    キーはコマンドラインに載せない（子プロセスの環境変数で渡す＝画面にも履歴にも残らない）。"""
    import subprocess
    script = (f"[Environment]::SetEnvironmentVariable('{name}', "
              + ("$null" if value is None else "$env:_VBAM_KEY_IN") + ", 'User')")
    env = dict(os.environ)
    if value is not None:
        env['_VBAM_KEY_IN'] = value
    p = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', script],
                       env=env, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError((p.stderr or p.stdout or '').strip() or f"powershell が {p.returncode} を返しました")


def _key_problem(key):
    """貼り付け事故を見つける。問題があれば理由（文字列）、無ければ None。"""
    if not key:
        return "何も入力されていません"
    if key != key.strip():
        return "前後に空白が付いています（貼り付けが余計なものを拾っています）"
    if any(c.isspace() for c in key):
        return "途中に空白や改行が入っています（貼り付けが途中で折れています）"
    if len(key) > _KEY_MAX:
        return f"長すぎます（{len(key):,} 字。上限 {_KEY_MAX:,} 字）"
    if any(ord(c) < 33 or ord(c) > 126 for c in key):
        return "英数字と記号以外の文字が混ざっています（全角で貼られていませんか）"
    return None


def _ping_key(ai, model, api_key):
    """一番短い問い合わせを 1 回だけ投げて、キーが本物かを確かめる。(ok, 一行) を返す。"""
    from vbam_ai import (_chat_claude, _chat_gemini)      # 遅延 import（vbam_ai がこの module を import している・2026-09-12 に持ち主へ付け替え）
    t0 = time.time()
    try:
        text, usage = (_chat_gemini if ai == 'gemini' else _chat_claude)(
            [('user', _PING_ASK)], model, api_key)
    except urllib.error.HTTPError as ex:
        detail = ''
        try:
            detail = ex.read().decode('utf-8', 'replace')[:200].replace("\n", " ")
        except Exception:
            pass
        return False, f"通りませんでした（HTTP {ex.code}）: {detail}"
    except Exception as ex:
        return False, f"通りませんでした: {ex}"
    sec = time.time() - t0
    return True, (f"通りました（{model} / {sec:.1f}秒 / 入力 {usage.get('in', 0):,}・"
                  f"出力 {usage.get('out', 0):,} トークン）")


def _key_ai_arg(args, default='gemini'):
    rest = [a for a in (getattr(args, 'posargs', None) or []) if not a.startswith('-')]
    ai = (rest[0] if rest else default).lower()
    if ai not in _KEY_ENV:
        raise ValueError(f"gemini か claude（{ai}）")
    return ai


def cmd_set_key(args):
    """APIキーを預かる: set-key [gemini|claude] [--no-ping] [--stdin]
    画面には出さずに受け取り、ユーザー環境変数へ書き、いまのプロセスにも入れて、疎通まで確かめる。"""
    try:
        ai = _key_ai_arg(args)
    except ValueError as ex:
        print(f"エラー: --ai は {ex}")
        return False
    env = _KEY_ENV[ai]
    print(f"==== {ai} APIキーの設定 ====\n")
    have, info = _key_info(ai)
    if info:
        print(f"いまの保管: {info}")
    if _user_env_get(env):
        print(f"※ 環境変数 {env} にも入っています（平文・期限なし）。clear-key で消せます。")
    if not have:
        print(f"キーは無料で発行できます: {_KEY_ISSUE_URL[ai]}")
    print()
    try:
        if getattr(args, 'stdin', False):
            key = sys.stdin.readline().rstrip("\r\n")
        else:
            import getpass
            key = getpass.getpass(f"{ai} のキーを貼り付けて Enter（画面には出ません）: ")
    except (EOFError, KeyboardInterrupt):
        print("\nやめました（変更していません）。")
        return False
    if not key:
        print("何も入力されなかったので、変更していません。")
        return False
    bad = _key_problem(key)
    if bad:
        print(f"エラー: {bad}")
        print("  もう一度、キーだけを貼り付けてください。")
        return False
    days = int(getattr(args, 'days', None) or _KEY_DAYS_DEFAULT)
    # 疎通が通ってから保管する（前は先に保管していたので、壊れたキーが金庫に残っていた・2026-09-04）
    if not getattr(args, 'no_ping', False):
        model = getattr(args, 'model', None) or _AI_DEFAULT_MODEL[ai]
        print("\n疎通を確かめています...")
        ok, line = _ping_key(ai, model, key)
        print("  " + line)
        if not ok:
            print("  → 貼り付けが途中で切れていないか確かめて、もう一度実行してください。（保管していません）")
            print(f"     キーの発行・確認: {_KEY_ISSUE_URL[ai]}")
            return False
    try:
        exp = _key_save(ai, key, days)
    except Exception as ex:
        print(f"エラー: 保管できませんでした: {ex}")
        return False
    os.environ[env] = key            # いまのプロセスにだけ入れる（窓を閉じれば消える）
    print(f"\n保管しました（{_mask_key(key)}）。"
          + (f"期限: {exp} まで。切れたら道具が自分で消します。" if exp else "期限: なし。"))
    print(f"  置き場所: {_KEY_STORE}（Windows のログインで暗号化。他のアカウント・他のパソコンでは戻せません）")
    if getattr(args, 'no_ping', False):
        print("（--no-ping: 疎通は確かめていません）")
    print(f"\nこれで agent が使えます（--ai {ai}）。この窓は閉じてかまいません。")
    return True


def cmd_clear_key(args):
    """APIキーを消す: clear-key [gemini|claude] [-y]
    環境変数から消すだけ。発行元のキーは生きているので、要るなら失効させること。"""
    rest = [a for a in (getattr(args, 'posargs', None) or []) if not a.startswith('-')]
    ais = [rest[0].lower()] if rest else list(_KEY_ENV)
    for a in ais:
        if a not in _KEY_ENV:
            print(f"エラー: gemini か claude（{a}）")
            return False
    print("==== APIキーの削除 ====\n")
    print("いま持っているもの:")
    found = []
    for a in ais:
        env = _KEY_ENV[a]
        _, info = _key_info(a)
        v = _user_env_get(env)
        if info:
            print(f"  {a}: 金庫 {info}")
        if v:
            print(f"  {a}: 環境変数 {env}  {_mask_key(v)}（平文・期限なし）")
        if not info and not v:
            print(f"  {a}: ありません")
        if info or v:
            found.append(a)
    if not found:
        print("\n消すものはありません。")
        return True
    print()
    if not getattr(args, 'yes', False):
        try:
            ans = input("削除しますか？ [y/N]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            ans = ''
        if ans not in ('y', 'yes'):
            print("やめました（変更していません）。")
            return False
    for a in found:
        env = _KEY_ENV[a]
        if _key_forget(a):
            print(f"{a}: 金庫から消しました。")
        if _user_env_get(env):
            try:
                _user_env_write(env, None)
            except Exception as ex:
                print(f"エラー: {env} を消せませんでした: {ex}")
                return False
            left = _user_env_get(env)      # 消したつもりで残っていないか、読み直して確かめる
            print(f"{a}: 環境変数 {env} を"
                  + ("消しました。" if not left else f"消せていません（{_mask_key(left)}）。"))
            if left:
                return False
        os.environ.pop(env, None)
    print("\n※ 発行元のキーは、まだ生きています。人に見られた恐れがあるなら失効させてください。")
    for a in found:
        print(f"   {a}: {_KEY_ISSUE_URL[a]}")
    print("※ Claude Code（MCP）から使っていた場合は、再起動してください。")
    return True


__all__ = ['_KEY_DAYS_DEFAULT', '_KEY_ENV', '_KEY_ISSUE_URL', '_KEY_MAX', '_KEY_STORE', '_PING_ASK', '_key_ai_arg', '_key_expired', '_key_forget', '_key_info', '_key_load', '_key_problem', '_key_save', '_mask_key', '_ping_key', '_store_read', '_store_write', '_user_env_get', '_user_env_write', 'cmd_clear_key', 'cmd_set_key']
