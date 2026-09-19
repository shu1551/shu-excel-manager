# -*- coding: utf-8 -*-
"""入口（先撃ち）を、見せていない種の表で実射する薄い入口（2026-09-18）。
  py entry_check_unseen.py <種> <仕事の名前> [<仕事の名前> …]
check_entry.py に「絶対パス＝名前」の形で渡す（相対パスだと Excel が開けない）。
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from unseen_sweep import JOBS   # noqa: E402

FOLDER = dict(JOBS)


def main():
    seed = sys.argv[1]
    rest, names, skip = [], [], False
    for a in sys.argv[2:]:                       # --named・--phrases N は check_entry.py へそのまま渡す
        if skip:
            rest.append(a)
            skip = False
        elif a.startswith('--'):
            rest.append(a)
            skip = (a == '--phrases')
        else:
            names.append(a)
    args = []
    for n in names:
        f = FOLDER.get(n)
        if not f:
            print(f"{n}: フォルダが分かりません")
            return 1
        p = os.path.join(HERE, f, f"未見_種{seed}")
        if not os.path.isdir(p):
            print(f"{n}: {p} が無い（先に unseen_sweep.py で作る）")
            return 1
        args.append(f"{p}={n}")
    cmd = [sys.executable, '-u', os.path.join(HERE, 'check_entry.py')] + args + ['--self'] + rest
    r = subprocess.run(cmd, cwd=HERE, stdin=subprocess.DEVNULL)
    return r.returncode


if __name__ == '__main__':
    sys.exit(main())
