# -*- coding: utf-8 -*-
"""agent の入口（先撃ち）から撃ったときを、依頼 × 未見の表の組み合わせ全部で確かめる（AI なし・自分の Excel）。

  py check_entry.py <未見のフォルダ>=<鍛えた名前> [<フォルダ>=<名前> ...] [--xlam パス]

- マクロは配った物そのもの＝秀コンボ.xlam の「表の整理」を閉じたまま読み、自分の Excel の新しいブックに取り込む
  （人の Excel・秀コンボには触らない）。
- 依頼（鍛えた台帳の依頼文）ごとに、全部のフォルダの表の写しを開いて vbam_prefire.macro_first を撃つ。
    その依頼の表     → AI を呼ばずに合格して、値が正解とそろう
    ほかの依頼の表   → 何も撃たない（値が変わらない）＝誤爆しない
- 控え・台帳・覚書は一時フォルダへ逃がす（本番の台帳を汚さない）。
"""
import contextlib
import glob
import io
import os
import shutil
import sys
import tempfile
import time

SCRIPTS = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts'))
sys.path.insert(0, SCRIPTS)
import vbam_core
vbam_core.setup_encoding()
import vba_manager  # noqa: F401  _run_cmd が後から読む表を先に読む
import vbam_agent as va
import vbam_forge as vf
import vbam_lineage as vl
import vbam_prefire as vp
import vbam_shake as vs

def _lead_of(before):
    """<名前>_選択.txt があれば「見出しと見出しで」を返す（人が列を呼ぶ言い方の頭）。無ければ空。"""
    import re as _re
    sel = before[:-len('_前.xlsx')] + '_選択.txt'
    if not os.path.isfile(sel):
        return ''
    from openpyxl import load_workbook as _lw
    ws = _lw(before, data_only=True).worksheets[0]
    head = []
    for row in ws.iter_rows(min_row=1, max_row=12):
        cells = [(c.column, str(c.value).strip()) for c in row if c.value not in (None, '')]
        if len(cells) >= 2:
            head = dict(cells)
            break
    names = []
    for a in open(sel, encoding='utf-8').read().split(','):
        a = a.strip()
        if not a:
            continue
        letters = _re.match(r'[A-Z]+', a).group()
        col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(letters)))
        names.append(head.get(col, ''))
    names = [n for n in names if n]
    return ('と'.join(names) + 'で') if names else ''

XLAM = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'AddIns', '秀コンボ.xlam')
args = [a for a in sys.argv[1:] if a != '-v']
if '--xlam' in args:
    i = args.index('--xlam')
    XLAM = args[i + 1]
    del args[i:i + 2]
OWN = '--own' in args            # --own＝表ごとに「その仕事の依頼（台帳＋言い換え）」と「ほかの仕事の台帳の依頼」だけ撃つ（2026-09-17 夜）
args = [a for a in args if a != '--own']
SELF = '--self' in args          # --self＝表ごとに、その仕事の依頼（台帳＋言い換え）だけ撃つ（2026-09-18: 同じ明細から作るグラフ・ピボットは互いに撃ってよい）
args = [a for a in args if a != '--self']
NAMED = '--named' in args        # --named＝選択.txt のある表は、依頼文の頭に選ぶ列の見出しを付ける（人が列を呼ぶ言い方）
args = [a for a in args if a != '--named']
PH_MAX = None                    # --phrases N＝言い換えは先頭 N 本だけ
if '--phrases' in args:
    i = args.index('--phrases')
    PH_MAX = int(args[i + 1])
    del args[i:i + 2]
ONLY = None                      # --requests 名前＝その仕事の依頼だけ撃つ（表は全部）
if '--requests' in args:
    i = args.index('--requests')
    ONLY = args[i + 1]
    del args[i:i + 2]
pairs = [a.split('=', 1) for a in args]
ledger = vf._forge_load()
LEDGER_SUB = {k: (v.get('sub') or '') for k, v in ledger.items()}   # 仕事 → Sub の名前（登録簿と突き合わせる）
# 依頼＝台帳の依頼文＋その仕事のフォルダの phrases.txt（言い換え）。名前!＝前提の欠けた表（撃ってはいけない）
requests = {}
for folder, name in pairs:
    task = name.rstrip('!')
    if task not in ledger or (ONLY and task != ONLY) or name.endswith('!'):
        continue
    requests.setdefault((task, '台帳'), ledger[task]['request'])
    ph = os.path.join(os.path.dirname(os.path.normpath(folder)) if os.path.basename(folder).startswith('未見') else folder,
                      'phrases.txt')
    if os.path.isfile(ph):
        for k, line in enumerate(open(ph, encoding='utf-8').read().splitlines(), 1):
            if line.strip():
                if PH_MAX and k > PH_MAX:
                    break
                requests[(task, f'言い換え{k}')] = line.strip()

# 表（直す前・正解）を集める
tables = []
for folder, name in pairs:
    broken = name.endswith('!')
    for before in sorted(glob.glob(os.path.join(folder, '*_前.xlsx'))):
        label = os.path.basename(before)[:-len('_前.xlsx')]
        task = name.rstrip('!')
        if broken and not label.startswith(task):
            continue
        tables.append({'task': task, 'broken': broken, 'label': label, 'before': before,
                       'truth': before[:-len('_前.xlsx')] + '_正解.xlsx',
                       'lead': _lead_of(before) if NAMED else ''})
print(f"表 {len(tables)} 枚 × 依頼 {len(requests)} 本（" + "・".join(f"{t}/{k}" for t, k in requests) + "）")
snaps = vf._read_sheets([(t['truth'], None) for t in tables] + [(t['before'], None) for t in tables])
for k, t in enumerate(tables):
    t['expect'], t['start'] = snaps[k], snaps[len(tables) + k]
    t['sheet'] = t['expect']['sheet']

# 配った物の「表の整理」を .bas に
mods = vl._read_closed_book(XLAM)
text = mods.get('表の整理')
if not text:
    sys.exit(f"エラー: {XLAM} に「表の整理」がありません")
work = tempfile.mkdtemp(prefix='_check_entry_')
bas = os.path.join(work, '表の整理.bas')
with open(bas, 'w', encoding='cp932', newline='') as f:
    f.write('Attribute VB_Name = "表の整理"\r\n' + text.replace('\r\n', '\n').replace('\n', '\r\n'))
reg = vp.registry_from_text(text)
print("登録簿: " + " ／ ".join(f"{e['name']}（{e['ask']}・見出し {' '.join(e['headers']) or 'なし'}）" for e in reg))

vs._divert_state(os.path.join(work, '_state'))
xl = vbam_core.get_or_start_excel(visible=False)
XL_PID = vbam_core._created_xl_pid
PER_TABLE_SEC = 180          # 1 枚にこれ以上かかったら台を落として打ち切る（2026-09-18: 1 枚で 32 分固まった）
fails = []
t_all = time.time()
done_n = 0
stalled = None
try:
    mwb = xl.Workbooks.Add()
    mwb.VBProject.VBComponents.Import(bas)
    for (rtask, rkind), req in requests.items():
        rname = f"{rtask}/{rkind}"
        for t in tables:
            if OWN and rtask != t['task'] and rkind != '台帳':
                continue
            if SELF and rtask != t['task']:
                continue
            # 登録簿の依頼の語が当たらない言い方（ピボット系の共通語だけの文など＝語では見分けられない）は
            # 撃たないのが正しい＝表が変わっていないことを見る（2026-09-18 朝）
            full_req = t.get('lead', '') + req
            # 表の形（' 形:）でも絞る＝形がそろわない表では撃たないのが正しい（2026-09-18）
            extras = vp.plan_full(full_req, reg, None, vf.shape_of_file(t['before'], t['sheet']))['extras']
            routed = any(e['name'] == LEDGER_SUB.get(rtask) for e in extras)
            want_fire = (t['task'] == rtask and not t['broken'] and routed)
            tmp = os.path.join(work, f"{t['label']}_{rtask}_{rkind}{os.path.splitext(t['before'])[1]}")
            shutil.copyfile(t['before'], tmp)
            # 1 枚ごとに始めた印を出す（止まっているのか進んでいるのかを画面で見分けられるように・2026-09-18）
            done_n += 1
            print(f"  [{done_n}] {time.strftime('%H:%M:%S')} 依頼「{rname}」× {t['label']} …", flush=True)
            dog, dog_hit = vf._watchdog(XL_PID, PER_TABLE_SEC)
            dog.start()
            wb = None
            buf = io.StringIO()
            try:
                wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
                wb.Activate()
                wb.Sheets(t['sheet']).Activate()
                # 撃つ前の姿は、開いた写しから読む（2026-09-17 夜: 読み取り専用で開いた値と使用範囲が 1 行ずれ、
                # 折りたたんだ行の表で撃っていないのに「誤爆（値が変わった）」と出た）
                start_now = vf._snapshot_texts(wb.Sheets(t['sheet']))
                t0 = time.time()
                with vbam_core.pinned_workbook(wb), contextlib.redirect_stdout(buf):
                    pre = vp.macro_first(full_req, t['sheet'], wb, 1, va._run_id())
                sec = time.time() - t0
                got = vf._snapshot_texts(wb.Sheets(t['sheet']))
                if want_fire:
                    ok_ai = bool(pre and pre.get('result') is not None)
                    mism = vf._compare(t['expect'], got)
                    good = ok_ai and not mism
                    why = ("" if good else ("AI に回した（先撃ちで終わらなかった）" if not ok_ai else "")
                           + (f" 値の不一致 {len(mism)}: {' ／ '.join(mism[:3])}" if mism else ""))
                else:
                    mism = vf._compare(start_now, got)
                    fired = [ln for ln in buf.getvalue().splitlines() if ln.startswith('マクロの先撃ち')]
                    good = not mism
                    why = "" if good else f"誤爆（{fired[0] if fired else '値が変わった'}） 変化 {len(mism)}"
            except Exception as ex:
                good, sec, why = False, 0.0, f"止まった: {ex}"
            finally:
                dog.cancel()
                with contextlib.suppress(Exception):
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                with contextlib.suppress(OSError):
                    os.remove(tmp)
            mark = '合格' if good else '外れ'
            if not good or '-v' in sys.argv:
                print(f"  依頼「{rname}」× {t['task']}/{t['label']}: {mark}（{sec:.1f} 秒）{('　' + why) if why else ''}")
            if not good:
                fails.append((rname, t, why, buf.getvalue()))
            if dog_hit['hit']:
                # 1 枚で PER_TABLE_SEC 秒を越えた＝Excel が返事をしない。台は落としたので、ここで打ち切って報告する
                stalled = (rname, t['label'])
                print(f"!! 固まりました: 依頼「{rname}」× {t['label']} が {PER_TABLE_SEC} 秒を越えたので"
                      "実射の台を落として打ち切りました", flush=True)
                break
        if stalled:
            break
finally:
    with contextlib.suppress(Exception):
        mwb.Close(SaveChanges=False)
    vbam_core.release_created_instances(only_saved=False)
n = (sum(1 for (rt, rk) in requests for t in tables if (rt == t['task'] if SELF else (rt == t['task'] or rk == '台帳')))
     if OWN or SELF else len(tables) * len(requests))
if stalled:
    print(f"入口の確かめ: 打ち切り（{done_n} 通りまで・依頼「{stalled[0]}」× {stalled[1]} で固まった）"
          f"　外れ {len(fails)}　計 {time.time() - t_all:.0f} 秒")
    sys.exit(2)
print(f"入口の確かめ: {n} 通り　合格 {n - len(fails)}・外れ {len(fails)}　計 {time.time() - t_all:.0f} 秒")
from collections import Counter
by = Counter((r, t['task'] + ('（壊れ）' if t['broken'] else '')) for r, t, _w, _o in fails)
for (r, task), k in sorted(by.items()):
    print(f"  外れの内訳: 依頼「{r}」× {task} の表 {k} 枚")
for rname, t, why, out in fails[:12]:
    print(f"FAILED\t{rname}\t{t['label']}\t{why}")
    tail = [ln for ln in out.splitlines() if ln.strip()][-4:]
    for ln in tail:
        print("    | " + ln)
