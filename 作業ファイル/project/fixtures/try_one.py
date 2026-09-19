# -*- coding: utf-8 -*-
"""台帳のマクロを 1 枚の表（お題の <名前>_前.xlsx）に撃って、正解との不一致と撃った後の姿を見る（AI なし・2026-09-18）。
  py try_one.py <仕事の名前> <前.xlsx のパス> [行数]
手で直したマクロを試すとき用（forge を回すより速い）。
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf


def main():
    name, before = sys.argv[1], os.path.abspath(sys.argv[2])
    rows = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    truth = before[:-len('_前.xlsx')] + '_正解.xlsx'
    case = dict(vf._forge_load()[name])
    snaps = vf._read_sheets([(truth, None)], fallback_first=True)
    if not snaps:
        return 1
    snap = snaps[0]
    sel = vf._sel_of(before)
    c = dict(case, sheet=snap['sheet'])
    res, sec = vf._run_on_copies(c, case['bas'], case['sub'], [(before, snap, snap['sheet'], sel)])
    mism = res[0]
    print(f"{os.path.basename(before)}: {'合格' if not mism else f'外れ {len(mism)} 件'}（{sec:.1f} 秒）"
          + (f"　選ぶ列 {sel}" if sel else ""))
    for line in mism[:12]:
        print('  ' + str(line)[:300])
    print('--- 正解（頭）')
    print(vf._table_text(snap, max_rows=rows))
    return 0


if __name__ == '__main__':
    sys.exit(main())
