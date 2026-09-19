# -*- coding: utf-8 -*-
"""撃たないはずの依頼で表の値が変わった（check_entry の「誤爆（値が変わった）」）ときに、何が変わったかを出す（AI なし・自分の Excel）。
  py entry_diff.py <前.xlsx> <依頼文>"""
import contextlib
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vba_manager  # noqa: F401
import vbam_agent as va
import vbam_forge as vf
import vbam_lineage as vl
import vbam_prefire as vp
import vbam_shake as vs

before, req = sys.argv[1], sys.argv[2]
xlam = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'AddIns', '秀コンボ.xlam')
text = vl._read_closed_book(xlam)['表の整理']
work = tempfile.mkdtemp(prefix='_entry_diff_')
bas = os.path.join(work, '表の整理.bas')
with open(bas, 'w', encoding='cp932', newline='') as f:
    f.write('Attribute VB_Name = "表の整理"\r\n' + text.replace('\r\n', '\n').replace('\n', '\r\n'))
vs._divert_state(os.path.join(work, '_state'))
start = vf._read_sheets([(before, None)])[0]
xl = vbam_core.get_or_start_excel(visible=False)
try:
    mwb = xl.Workbooks.Add()
    mwb.VBProject.VBComponents.Import(bas)
    tmp = os.path.join(work, os.path.basename(before))
    shutil.copyfile(before, tmp)
    wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
    wb.Activate()
    wb.Sheets(start['sheet']).Activate()
    opened = vf._snapshot_texts(wb.Sheets(start['sheet']))
    print("開いた直後と読み取り専用で読んだ値の違い:", vf._compare(start, opened)[:5])
    with vbam_core.pinned_workbook(wb):
        pre = vp.macro_first(req, start['sheet'], wb, 1, va._run_id())
    got = vf._snapshot_texts(wb.Sheets(start['sheet']))
    print("macro_first の戻り値:", None if pre is None else {k: str(v)[:120] for k, v in pre.items()})
    print("撃つ前との違い:", vf._compare(start, got)[:8])
finally:
    with contextlib.suppress(Exception):
        wb.Close(SaveChanges=False)
    with contextlib.suppress(Exception):
        mwb.Close(SaveChanges=False)
    vbam_core.release_created_instances(only_saved=False)
