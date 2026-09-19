# -*- coding: utf-8 -*-
"""check_phrases.py の「人の言い方」のお題を、各お題の phrases.txt と台帳の言い換えに足す（2026-09-18）。

なぜ: 作り直しでお題の依頼文を語の要らない形にしたので、道具の決まり（依頼の語は依頼文か言い換えに
そのまま含まれる語だけ）で 決算統計・受付簿・按分 のような職場の語が落ち、入口が撃たなくなった。
言い方は人が書く文であって表の語ではない＝言い換えに足してよい。
  py add_phrases.py [--write]
"""
import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', 'python_scripts')))
sys.path.insert(0, HERE)
import vbam_core   # noqa: E402
vbam_core.setup_encoding()
import vbam_forge as vf   # noqa: E402
import check_phrases as cp   # noqa: E402


def folder_of(from_book, phrases):
    """同じ名前の表が古いフォルダにも残っているので、台帳の言い換えと一致する phrases.txt を持つフォルダを選ぶ。"""
    best, score = None, -1
    for p in glob.glob(os.path.join(HERE, '*', from_book)):
        folder = os.path.dirname(p)
        f = os.path.join(folder, 'phrases.txt')
        if not os.path.isfile(f):
            continue
        now = {x.strip() for x in open(f, encoding='utf-8').read().splitlines() if x.strip()}
        s = len(now & set(phrases or []))
        if s > score:
            best, score = folder, s
    return best if score > 0 else None


def main():
    write = '--write' in sys.argv
    d = vf._forge_load()
    add = {}
    for src in (cp.HIT1, cp.HIT2):
        for job, phrases in src.items():
            add.setdefault(job, []).extend(phrases)
    for job, phrases in sorted(add.items()):
        case = d.get(job)
        if not case:
            print(f"  {job}: 台帳に無い（とばします）")
            continue
        folder = folder_of(case.get('from_book') or '', case.get('phrases'))
        if not folder:
            print(f"  {job}: お題のフォルダが分かりません（{case.get('from_book')}）")
            continue
        path = os.path.join(folder, 'phrases.txt')
        now = [x.strip() for x in open(path, encoding='utf-8').read().splitlines() if x.strip()]
        new = [p for p in phrases if p not in now]
        if not new:
            print(f"  {job}: 足すものなし")
            continue
        print(f"  {job}（{os.path.basename(folder)}）に {len(new)} 本足す: " + "・".join(new[:3]) + " …")
        if write:
            open(path, 'w', encoding='utf-8').write("\n".join(now + new) + "\n")
            case['phrases'] = now + new
            d[job] = case
    if write:
        vf._forge_save(d)
        print('台帳と phrases.txt を書きました')
    else:
        print('（--write で書きます）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
