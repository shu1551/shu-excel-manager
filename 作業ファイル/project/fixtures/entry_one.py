# -*- coding: utf-8 -*-
"""入口（先撃ち）を 1 枚だけ撃って、道具の出力を全部見る（check_entry の外れの理由を読む用・AI なし・自分の Excel）。
  py entry_one.py <前.xlsx> <鍛えた名前>     （依頼＝台帳の依頼文）"""
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

before, name = sys.argv[1], sys.argv[2]
req = sys.argv[3] if len(sys.argv) > 3 else vf._forge_load()[name]['request']
xlam = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'AddIns', '秀コンボ.xlam')
text = vl._read_closed_book(xlam)['表の整理']
work = tempfile.mkdtemp(prefix='_entry_one_')
bas = os.path.join(work, '表の整理.bas')
with open(bas, 'w', encoding='cp932', newline='') as f:
    f.write('Attribute VB_Name = "表の整理"\r\n' + text.replace('\r\n', '\n').replace('\n', '\r\n'))
vs._divert_state(os.path.join(work, '_state'))
snap = vf._read_sheets([(before, None)])[0]
xl = vbam_core.get_or_start_excel(visible=False)
try:
    mwb = xl.Workbooks.Add()
    mwb.VBProject.VBComponents.Import(bas)
    tmp = os.path.join(work, os.path.basename(before))
    shutil.copyfile(before, tmp)
    wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
    wb.Activate()
    wb.Sheets(snap['sheet']).Activate()
    with vbam_core.pinned_workbook(wb):
        pre = vp.macro_first(req, snap['sheet'], wb, 1, va._run_id())
    print("=== 戻り値:", {k: (str(v)[:300] if v is not None else None) for k, v in (pre or {}).items()})
    truth = before[:-len('_前.xlsx')] + '_正解.xlsx'
    exp = vf._read_sheets([(truth, None)])[0]
    got = vf._snapshot_texts(wb.Sheets(snap['sheet']))
    print("=== 正解との不一致:", vf._compare(exp, got)[:8])
finally:
    with contextlib.suppress(Exception):
        wb.Close(SaveChanges=False)
    with contextlib.suppress(Exception):
        mwb.Close(SaveChanges=False)
    vbam_core.release_created_instances(only_saved=False)
