# -*- coding: utf-8 -*-
"""鍛えたマクロを、AI に見せていない表（未見_種N）に当てて値で突き合わせる（AI なし・自分の Excel）。
  py check_unseen.py <未見のフォルダ> <鍛えた名前>
外れた表は、鍛える往復の --test に足して撃ち直す（agent --forge 名前 --test 前.xlsx 正解.xlsx …）。"""
import glob
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf

folder = os.path.abspath(sys.argv[1])
name = sys.argv[2]
case = dict(vf._forge_load()[name])
pairs = []
for before in sorted(glob.glob(os.path.join(folder, '*_前.xlsx'))):
    truth = before[:-len('_前.xlsx')] + '_正解.xlsx'
    pairs.append((os.path.basename(before)[:-len('_前.xlsx')], before, truth))
snaps = vf._read_sheets([(t, None) for _l, _b, t in pairs])
failed = []
for (label, before, truth), snap in zip(pairs, snaps):
    c = dict(case, sheet=snap['sheet'])
    res, sec = vf._run_on_copies(c, case['bas'], case['sub'], [(before, snap)])
    ok = not res[0]
    print(f"  {label}: {'合格' if ok else '外れ ' + str(len(res[0]))}（{sec:.1f} 秒）" + ('' if ok else '　' + ' ／ '.join(res[0][:3])))
    if not ok:
        failed.append((before, truth))
print(f"未見 {len(pairs)} 枚: 合格 {len(pairs) - len(failed)}・外れ {len(failed)}")
for b, t in failed:
    print(f"FAILED\t{b}\t{t}")
