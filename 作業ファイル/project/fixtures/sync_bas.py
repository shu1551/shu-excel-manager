# -*- coding: utf-8 -*-
"""台帳の code を .bas（cp932）に書き戻す（手で直したマクロを試す前に・2026-09-18）。
  py sync_bas.py <仕事の名前> [<名前> ...]
  py sync_bas.py --check            台帳の code と .bas が食い違う仕事を並べる
cp932 に無い字は近い字に寄せる（〜→～ 等）。
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LEDGER = os.path.abspath(os.path.join(HERE, '..', 'python_scripts', '_agent_forge.json'))
_FIX = {0x301C: '～', 0x2212: '－', 0x2015: '―', 0x2225: '∥', 0xFF0D: '－', 0xFFE0: '¢', 0xFFE1: '£', 0xFFE2: '¬'}


def _to_cp932(text):
    return text.translate(_FIX)


def body_of(path):
    s = open(path, encoding='cp932').read()
    return s.split('\n', 1)[1] if s.startswith('Attribute') else s


def sync(name, entry):
    bas, code = entry.get('bas'), entry.get('code')
    if not bas or not code:
        print(f"  {name}: bas か code が無い")
        return False
    head = 'Attribute VB_Name = "鍛冶_試し"\n'
    if os.path.exists(bas):
        s = open(bas, encoding='cp932').read()
        if s.startswith('Attribute'):
            head = s.split('\n', 1)[0] + '\n'
    open(bas, 'w', encoding='cp932', newline='\n').write(head + _to_cp932(code.rstrip()) + '\n')
    print(f"  {name}: {len(code.splitlines())} 行を .bas に書いた")
    return True


def main():
    d = json.load(open(LEDGER, encoding='utf-8'))
    args = sys.argv[1:]
    if not args or args[0] == '--check':
        ng = 0
        for k, e in d.items():
            if not e.get('bas') or not e.get('code') or not os.path.exists(e['bas']):
                continue
            if body_of(e['bas']).strip() != _to_cp932(e['code']).strip():
                print('違', k)
                ng += 1
        print(f"食い違い {ng} 件")
        return 0
    for name in args:
        if name not in d:
            print(f"  {name}: 台帳に無い")
            continue
        sync(name, d[name])
    json.dump(d, open(LEDGER, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    return 0


if __name__ == '__main__':
    sys.exit(main())
