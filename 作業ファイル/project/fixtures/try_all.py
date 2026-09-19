# -*- coding: utf-8 -*-
"""台帳のマクロを、お題のフォルダの表ぜんぶに撃って外れを並べる（AI なし・Excel は 1 回だけ起こす・2026-09-18）。
  py try_all.py <仕事の名前> <フォルダ> [件数]
手で直したマクロを全表で確かめるとき用（forge を回すと AI を呼ぶ）。
"""
import glob
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf


def main():
    name, folder = sys.argv[1], os.path.abspath(sys.argv[2])
    show = int(sys.argv[3]) if len(sys.argv) > 3 else 6
    case = dict(vf._forge_load()[name])
    jobs = []
    for b in sorted(glob.glob(os.path.join(folder, '*_前.xlsx'))):
        t = b[:-len('_前.xlsx')] + '_正解.xlsx'
        if os.path.isfile(t):
            jobs.append((b, t))
    snaps = vf._read_sheets([(t, None) for _b, t in jobs], fallback_first=True)
    runs, names = [], []
    for (b, _t), snap in zip(jobs, snaps):
        runs.append((b, snap, snap['sheet'], vf._sel_of(b)))
        names.append(os.path.basename(b))
    res, sec = vf._run_on_copies(dict(case, sheet=snaps[0]['sheet']), case['bas'], case['sub'], runs)
    ng = 0
    for nm, mism in zip(names, res):
        if not mism:
            print(f"合格 {nm}")
            continue
        ng += 1
        print(f"外れ {len(mism)} 件 {nm}")
        for line in mism[:show]:
            print('    ' + str(line)[:400])
    print(f"--- {len(names)} 表　外れ {ng} 表（{sec:.1f} 秒）")
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main())
