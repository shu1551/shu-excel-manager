# -*- coding: utf-8 -*-
"""鍛えたマクロを、いくつものフォルダの表（*_前.xlsx と *_正解.xlsx）にまとめて当てる（AI なし・自分の Excel 1 つ）。
  py check_many.py <鍛えた名前> <フォルダ> [<フォルダ> ...] [--bas 別の.bas]
check_unseen.py は表ごとに Excel を起こす。こちらは 1 つの Excel で順に撃つ（鍛え直した後の退行の確かめ用・2026-09-17 夜）。"""
import glob
import os
import sys
import time

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf

args = sys.argv[1:]
bas = None
if '--bas' in args:
    i = args.index('--bas')
    bas = args[i + 1]
    del args[i:i + 2]
name, folders = args[0], args[1:]
case = dict(vf._forge_load()[name])
pairs = []
for folder in folders:
    for before in sorted(glob.glob(os.path.join(folder, '*_前.xlsx'))):
        pairs.append((f"{os.path.basename(folder)}/{os.path.basename(before)[:-len('_前.xlsx')]}", before,
                      before[:-len('_前.xlsx')] + '_正解.xlsx'))
t0 = time.time()
snaps = vf._read_sheets([(t, None) for _l, _b, t in pairs])
res, sec = vf._run_on_copies(case, bas or case['bas'], case['sub'], [(b, s, s['sheet']) for (_l, b, _t), s in zip(pairs, snaps)])
failed = []
by_folder = {}
for (label, before, truth), mism in zip(pairs, res):
    folder = label.split('/')[0]
    ok_n, n = by_folder.get(folder, (0, 0))
    by_folder[folder] = (ok_n + (not mism), n + 1)
    if mism:
        failed.append(label)
        print(f"  {label}: 外れ {len(mism)}　" + ' ／ '.join(mism[:3]))
for folder, (ok_n, n) in by_folder.items():
    print(f"  {folder}: {ok_n}/{n}")
print(f"{name}: {len(pairs)} 枚　合格 {len(pairs) - len(failed)}・外れ {len(failed)}（{time.time() - t0:.0f} 秒）")
