# -*- coding: utf-8 -*-
"""入口を「列の名前を言う依頼文」で 1 枚撃つ（選ぶ列の仕事の実射・2026-09-18）。
  py entry_one_named.py <前.xlsx> <鍛えた名前> <言い方のひな形>
ひな形の {列} が、その表の選ばれる列の見出しを「と」でつないだ文字に置き換わる。
  例: py entry_one_named.py PQ追加集計/未見_種4242/未見1_標準_前.xlsx PQ課別集計 "{列}で同じ見出しの表を全部足して集計して"
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from check_select import header_cells   # noqa: E402


def main():
    before = os.path.abspath(sys.argv[1])
    name, tmpl = sys.argv[2], sys.argv[3]
    sel = before[:-len('_前.xlsx')] + '_選択.txt'
    cells = dict(header_cells(before))
    cols = []
    for a in open(sel, encoding='utf-8').read().split(','):
        a = a.strip()
        if not a:
            continue
        letters = re.match(r'[A-Z]+', a).group()
        cols.append(sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(letters))))
    if tmpl == '--ledger':                       # 台帳の依頼文の頭に列の名前を付ける（check_entry --named と同じ形）
        sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', 'python_scripts')))
        import vbam_forge as vf
        tmpl = '{列}で' + vf._forge_load()[name]['request']
    req = tmpl.replace('{列}', 'と'.join(cells.get(c, '?') for c in cols))
    print('依頼:', req)
    return subprocess.run([sys.executable, '-u', os.path.join(HERE, 'entry_one.py'), before, name, req],
                          cwd=HERE, stdin=subprocess.DEVNULL).returncode


if __name__ == '__main__':
    sys.exit(main())
