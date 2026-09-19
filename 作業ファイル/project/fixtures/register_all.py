# -*- coding: utf-8 -*-
"""鍛え直したマクロを「表の整理」へ 1 回で登録する（2026-09-18）。

なぜ 1 回か: 9/18 未明に add-procedure＋compile を数時間で 30 回ほど続けたら、Excel が VBE7.DLL で落ちた。
モジュールの本文を Python で組み立て、replace-module 1 回＋compile 1 回にする。

  py register_all.py            … 何をするか見せるだけ（Excel には書かない）
  py register_all.py --write    … 実際に入れ替える（控えは道具が取る）
  py register_all.py --write --to 秀コンボ.xlsm
残す Sub: 表を整える・重複行を消す・台帳に無い Sub。消す Sub: 台帳の古い名前（作り直しで名前が変わったもの）。
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = os.path.abspath(os.path.join(HERE, '..', 'python_scripts'))
sys.path.insert(0, PY)
import vbam_core   # noqa: E402
vbam_core.setup_encoding()
import vbam_forge as vf   # noqa: E402

MODULE = '表の整理'
KEEP = ('表を整える', '重複行を消す')
OLD_LEDGER = os.path.join(PY, '_agent_forge.json.bak_20260918_gen')
# 1 本にまとめて要らなくなった仕事（その Sub は消す）
DROP_JOBS = ('課別棒グラフ', '推移折れ線')
_SUB_RE = re.compile(r'^[ \t]*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(', re.M)


def blocks(text):
    """モジュールの本文 → [(名前 or None, 塊の文字)]（純 Python）。名前 None は頭のコメントなど。"""
    out, name, buf = [], None, []
    for ln in text.replace('\r\n', '\n').split('\n'):
        m = _SUB_RE.match(ln)
        if m:
            if buf:
                out.append((name, '\n'.join(buf)))
            name, buf = m.group(1), [ln]
            continue
        buf.append(ln)
        if re.match(r'^[ \t]*End\s+Sub\b', ln, re.I) and name:
            out.append((name, '\n'.join(buf)))
            name, buf = None, []
    if buf:
        out.append((name, '\n'.join(buf)))
    return out


def main():
    write = '--write' in sys.argv
    to = sys.argv[sys.argv.index('--to') + 1] if '--to' in sys.argv else '秀コンボ.xlsm'
    if not os.path.sep in to:
        # 開いているブックから名前でパスを引く（CLI の対象はパスで渡す）
        import vbam_agent as va
        xl, _wb = va.get_workbook(None)
        hit = [str(w.FullName) for w in xl.Workbooks if str(w.Name) == to]
        if not hit:
            print(f"エラー: 「{to}」が開いていません（開いているブック: "
                  + "・".join(str(w.Name) for w in xl.Workbooks) + "）")
            return 1
        to = hit[0]
        print(f"対象: {to}")
    d = vf._forge_load()
    old = {}
    if os.path.isfile(OLD_LEDGER):
        import json
        with open(OLD_LEDGER, encoding='utf-8') as f:
            oldd = json.load(f)
        old = {k: (v.get('sub') or '') for k, v in oldd.items()}
        old_code = {k: (v.get('code') or '') for k, v in oldd.items()}
    new_subs, drop, not_passed, kept_old = {}, set(), [], []
    for name in DROP_JOBS:
        if old.get(name):
            drop.add(old[name])                           # 1 本にまとめた古い仕事の Sub は消す
    retired = []
    for name, c in d.items():
        if name in DROP_JOBS:
            continue
        if c.get('retired'):
            # 引退した仕事（2026-09-18 shu「作りすぎ・役に立たないものは消す」）: 台帳とお題は残し、登録簿からだけ外す
            if c.get('sub'):
                drop.add(c['sub'])
                retired.append((name, c['sub']))
            continue
        rebuilt = (c.get('code') or '') != old_code.get(name, '')
        if c.get('passed') and c.get('code') and c.get('sub') and rebuilt:
            new_subs[c['sub']] = c['code'].replace('\r\n', '\n').strip('\n')
            if old.get(name) and old[name] != c['sub']:
                drop.add(old[name])                       # 作り直しで名前が変わった＝古い Sub は消す
        elif c.get('sub') and not rebuilt:
            kept_old.append((name, c['sub']))
        elif c.get('sub'):
            not_passed.append((name, c['sub']))
    # 書き出し（開いているブックの今の本文）
    out_bas = os.path.join(PY, f"{MODULE}.bas")
    if os.path.isfile(out_bas):
        os.remove(out_bas)
    r = subprocess.run([sys.executable, os.path.join(PY, 'vba_manager.py'), 'export-module', to, MODULE],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=PY)
    if not os.path.isfile(out_bas):
        print('エラー: モジュールを書き出せませんでした:\n' + (r.stdout or '') + (r.stderr or ''))
        return 1
    with open(out_bas, encoding='cp932') as f:
        text = f.read()
    text = re.sub(r'^\s*Attribute\s+VB_Name\s*=.*\r?\n', '', text, count=1)   # 属性行は _write_bas が付ける
    bs = blocks(text)
    have = [n for n, _b in bs if n]
    keep_blocks, replaced = [], []
    for n, b in bs:
        if n is None:
            keep_blocks.append(b)
            continue
        if n in drop and n not in new_subs:
            print(f"  消す（{'引退' if any(s == n for _j, s in retired) else '名前が変わった古い Sub'}）: {n}")
            continue
        if n in new_subs:
            keep_blocks.append(new_subs.pop(n))
            replaced.append(n)
            continue
        keep_blocks.append(b)
    added = list(new_subs)
    body = "\n\n".join(x.strip('\n') for x in keep_blocks if x.strip())
    if added:
        body += "\n\n" + "\n\n".join(new_subs[n] for n in added)
    body = body.rstrip('\n') + "\n"
    print(f"今の Sub {len(have)} 本 → 置き換え {len(replaced)} 本・新しく足す {len(added)} 本・消す {len(drop & set(have))} 本")
    print("  置き換え: " + "・".join(replaced))
    print("  足す: " + "・".join(added))
    if retired:
        print("  引退（台帳とお題は残す・登録簿から外す）: " + "・".join(f"{n}={s}" for n, s in retired))
    if not_passed:
        print("  （合格していない弾は触りません: " + "・".join(f"{n}={s}" for n, s in not_passed) + "）")
    if kept_old:
        print("  （まだ作り直していない仕事はそのまま: " + "・".join(f"{n}={s}" for n, s in kept_old) + "）")
    path = os.path.join(PY, '_表の整理_新.bas')
    err = vf._write_bas(path, MODULE, body)
    if err:
        print('エラー: ' + err)
        return 1
    print(f"  新しい本文: {path}（{len(body.splitlines())} 行）")
    if not vf._check_bas(path):
        print('エラー: check-bas に落ちました')
        return 1
    if not write:
        print("（--write を付けると replace-module で入れ替えます）")
        return 0
    for args in (['replace-module', to, MODULE, path], ['compile', to]):
        r = subprocess.run([sys.executable, os.path.join(PY, 'vba_manager.py')] + args,
                           capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=PY)
        print((r.stdout or '').strip()[-1500:])
        if r.returncode:
            print('エラー: ' + (r.stderr or '')[-500:])
            return 1
    gone = [n for n in DROP_JOBS if n in d]
    if gone:
        for n in gone:
            d.pop(n, None)
        vf._forge_save(d)
        print("台帳から外しました（1 本にまとめた古い仕事）: " + "・".join(gone))
    print("登録しました（replace-module がブックを保存済み）。次は run-macro アドインの更新登録 --raw で .xlam へ")
    return 0


if __name__ == '__main__':
    sys.exit(main())
