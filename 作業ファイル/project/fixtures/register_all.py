# -*- coding: utf-8 -*-
"""鍛え直したマクロを棚（表の整理・表の整理_作る・表の整理_調べる）へ 1 回で登録する（2026-09-18・9/24 作り直し）。

なぜ 1 回か: 9/18 未明に add-procedure＋compile を数時間で 30 回ほど続けたら、Excel が VBE7.DLL で落ちた。
モジュールの本文を Python で組み立て、変わるモジュールごとに replace-module 1 回＋最後に compile 1 回にする。

  py register_all.py            … 何をするか見せるだけ（Excel にも台帳にも書かない）
  py register_all.py --write    … 実際に入れ替える（控えは道具が取る）・台帳に「登録した本文」を記す
  py register_all.py --write --to 秀コンボ.xlsm

決まり（9/24 作り直し）:
- 正はブックの今の本文。棚は 9/23 に 3 つへ分かれ、9/24 の総点検の直しはブックにだけ入っている。
  台帳の本文（code）で上書きするのは、回路が作り直した仕事＝ code が「前回登録した本文」（registered）と違うものだけ。
- registered がまだ無い仕事は、ブックの本文を台帳に取り込む（code と registered をブックの本文にする）＝上書きしない。
- 置き換えは、その Sub が今いるモジュールで行う。どこにも無い Sub は名前で置き場を決める（_調べる／_作る／表の整理）。
- 引退した仕事・名前が変わった古い Sub は、いるモジュールから消す。
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

MAIN, MAKE, LOOK = '表の整理', '表の整理_作る', '表の整理_調べる'
MODULES = (MAIN, MAKE, LOOK)
# 1 本にまとめて要らなくなった仕事（その Sub は消す）
DROP_JOBS = ('課別棒グラフ', '推移折れ線')
_SUB_RE = re.compile(r'^[ \t]*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(', re.M)
_LOOK_RE = re.compile(r'(一覧にする|確かめる|報告する|調べる|探す|検算する)$')
_MAKE_RE = re.compile(r'(作る|足す|グラフ|ピボット|クエリ|集計|転記|差し込む|分ける|展開する|挿入する|抜き出す|印刷設定)')


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


def norm(code):
    return (code or '').replace('\r\n', '\n').strip('\n')


def home_of(sub):
    """どこにも無い Sub の置き場（名前の終わり・語で決める）。"""
    if _LOOK_RE.search(sub):
        return LOOK
    if _MAKE_RE.search(sub):
        return MAKE
    return MAIN


def vm(*args):
    r = subprocess.run([sys.executable, os.path.join(PY, 'vba_manager.py')] + list(args),
                       capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=PY,
                       stdin=subprocess.DEVNULL)
    return r.returncode, (r.stdout or '') + (r.stderr or '')


def export(to, mod):
    p = os.path.join(PY, f"{mod}.bas")
    if os.path.isfile(p):
        os.remove(p)
    _c, out = vm('export-module', to, mod)
    if not os.path.isfile(p):
        raise SystemExit(f'エラー: モジュール {mod} を書き出せませんでした:\n{out}')
    with open(p, encoding='cp932') as f:
        text = f.read()
    return re.sub(r'^\s*Attribute\s+VB_Name\s*=.*\r?\n', '', text, count=1)   # 属性行は _write_bas が付ける


def main():
    write = '--write' in sys.argv
    to = sys.argv[sys.argv.index('--to') + 1] if '--to' in sys.argv else '秀コンボ.xlsm'
    if os.path.sep not in to:
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

    mod_blocks = {m: blocks(export(to, m)) for m in MODULES}
    where = {}                                   # Sub 名 → モジュール
    book_code = {}                               # Sub 名 → ブックの今の本文
    for m, bs in mod_blocks.items():
        for n, b in bs:
            if n:
                where[n] = m
                book_code[n] = norm(b)

    d = vf._forge_load()
    put = {}          # Sub 名 → 新しい本文（置き換え・足す）
    drop = set()      # 消す Sub
    adopt, same, not_passed, retired = [], [], [], []
    for name, c in d.items():
        sub = c.get('sub') or ''
        if name in DROP_JOBS or c.get('retired'):
            if sub:
                drop.add(sub)
                retired.append((name, sub))
            continue
        if not (c.get('passed') and c.get('code') and sub):
            if sub:
                not_passed.append((name, sub))
            continue
        old_sub = c.get('registered_sub') or ''
        if old_sub and old_sub != sub and old_sub in where:
            drop.add(old_sub)                    # 作り直しで名前が変わった＝古い Sub は消す
        reg = c.get('registered')
        if reg is None:
            if sub in book_code:
                adopt.append(name)               # 初回: ブックの本文を正として台帳に取り込む
            else:
                put[sub] = norm(c['code'])       # ブックに無い＝足す
            continue
        if norm(c['code']) != norm(reg):
            put[sub] = norm(c['code'])           # 回路が作り直した
        elif sub not in book_code:
            put[sub] = norm(c['code'])           # 登録済みのはずがブックに無い＝足し直す
        else:
            same.append(name)

    # モジュールごとに本文を組み直す
    new_body, report = {}, {}
    added = {sub: (where.get(sub) or home_of(sub)) for sub in put}
    for m in MODULES:
        keep, rep, gone = [], [], []
        for n, b in mod_blocks[m]:
            if n is None:
                keep.append(b)
            elif n in drop and n not in put:
                gone.append(n)
            elif n in put:
                keep.append(put[n])
                rep.append(n)
            else:
                keep.append(b)
        add = [s for s, mm in added.items() if mm == m and s not in where]
        keep += [put[s] for s in add]
        if rep or gone or add:
            new_body[m] = "\n\n".join(x.strip('\n') for x in keep if x.strip()).rstrip('\n') + "\n"
            report[m] = (rep, add, gone)

    print(f"棚の Sub {len(where)} 本（{'・'.join(f'{m} {sum(1 for v in where.values() if v == m)}' for m in MODULES)}）")
    for m, (rep, add, gone) in report.items():
        print(f"[{m}] 置き換え {len(rep)}・足す {len(add)}・消す {len(gone)}")
        for lab, xs in (("置き換え", rep), ("足す", add), ("消す", gone)):
            if xs:
                print(f"  {lab}: " + "・".join(xs))
    if not report:
        print("入れ替えるモジュールはありません（ブックと台帳の登録済みの本文が一致）")
    if adopt:
        print(f"ブックの本文を台帳に取り込む（上書きしない）: {len(adopt)} 本")
    if same:
        print(f"（登録済みと同じ: {len(same)} 本）")
    if not_passed:
        print("（合格していない弾は触りません: " + "・".join(f"{n}={s}" for n, s in not_passed) + "）")

    paths = {}
    for m, body in new_body.items():
        path = os.path.join(PY, f'_{m}_新.bas')
        err = vf._write_bas(path, m, body)
        if err:
            print('エラー: ' + err)
            return 1
        if not vf._check_bas(path):
            print(f'エラー: check-bas に落ちました: {path}')
            return 1
        paths[m] = path
        print(f"  新しい本文: {path}（{len(body.splitlines())} 行）")
    if not write:
        print("（--write を付けると replace-module で入れ替え、台帳に登録した本文を記します）")
        return 0

    for m, path in paths.items():
        code, out = vm('replace-module', to, m, path)
        print(out.strip()[-600:])
        if code:
            return 1
    if paths:
        code, out = vm('compile', to)
        print(out.strip()[-800:])
        if code:
            return 1
    # 台帳に「登録した本文」を記す（ブックに入った形＝次からはこれと比べる）
    for name, c in d.items():
        sub = c.get('sub') or ''
        if name in adopt:
            c['code'] = book_code[sub]
            c['registered'], c['registered_sub'] = book_code[sub], sub
        elif sub in put:
            c['registered'], c['registered_sub'] = put[sub], sub
        elif name in same:
            c['registered_sub'] = sub
    for n in DROP_JOBS:
        d.pop(n, None)
    vf._forge_save(d)
    print("登録しました（replace-module がブックを保存済み）・台帳に登録した本文を記しました。"
          "次は run-macro アドインの更新登録 --raw で .xlam へ")
    return 0


if __name__ == '__main__':
    sys.exit(main())
