# -*- coding: utf-8 -*-
"""台帳の全部の仕事の頭の行（依頼の語・扱う・見出し）を、今の規則とほかの仕事の今の語で検査する（純 Python・Excel なし・2026-09-18）。

仕事を足すと、前に合格した仕事の語が新しい仕事の言い換えに当たる（月次集計の「課ごとの月」が推移の言い換えに当たった）。
  py check_heads.py          … 違反の一覧だけ
  py check_heads.py --fix    … 道具が直せるものは台帳と .bas を直す（登録は agent --forge 名 --register で別に）
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf
import vbam_prefire as vp

FIX = '--fix' in sys.argv
d = vf._forge_load()
bad = 0
live = {k: v for k, v in d.items() if not v.get('retired')}       # 引退した仕事は検査もしない・ほかの仕事ともぶつからない（2026-09-18）
for name, case in live.items():
    if not case.get('code'):
        continue
    vf._OTHER_JOB_TEXTS.clear()
    vf._OTHER_JOB_TEXTS.update({k: [str(v.get('request') or '')] + list(v.get('phrases') or [])
                                for k, v in live.items() if k != name and (v.get('passed') or v.get('ask'))})
    vf._OTHER_JOB_ASKS.clear()
    vf._OTHER_JOB_ASKS.update({k: str(v.get('ask') or '') for k, v in live.items() if k != name and (v.get('passed') or v.get('ask'))})
    vf._OTHER_JOB_WORDS.clear()
    vf._OTHER_JOB_WORDS.update({k: vp.sheet_words((v.get('before_values') or {}).get('values'))
                                for k, v in live.items() if k != name and v.get('passed') and v.get('before_values')})
    # ' 形: の行は道具が持つ（表から導く）＝無ければ足す・導いた形と違えば直す（2026-09-18）
    shaped, changed = vf.ensure_shape_line(case, case['code'])
    if changed:
        line = vf._SHAPE_LINE_RE.search(shaped).group(0).strip()
        if FIX:
            case['code'] = shaped
            d[name] = case
            vf._forge_save(d)
            if case.get('bas'):
                vf._write_bas(case['bas'], vf._MODULE, shaped)
            print(f"{name}: 表の形の行を入れました（要登録）: {line}")
        else:
            print(f"{name}: 表の形の行がありません（--fix で入ります）: {line}")
            bad += 1
    nm, why, _h = vf._validate_code(case['code'])
    ask = why if nm else None
    if nm:
        _hd, why = vf._check_fit(case, case['code'], why)
    lit = vf._literal_error(case, case['code'])               # 表の語の書き込み（道具では直せない＝agent --forge で）
    if lit:
        bad += 1
        print(f"{name}: {lit[:260]}")
        print("  → 表の形に作り直す（agent --forge で）")
    if not why:
        continue
    bad += 0 if lit else 1
    print(f"{name}: {why[:260]}")
    if FIX and ask:
        fixed, _a = vf._auto_fix_ask(case, case['code'], ask)
        if fixed:
            nm2, ask2, handles2 = vf._validate_code(fixed)
            heads2, why2 = vf._check_fit(case, fixed, ask2) if nm2 else (None, ask2)
            if nm2 and not why2:
                case.update({'code': fixed, 'ask': ask2, 'handles': handles2, 'headers': heads2})
                d[name] = case
                vf._forge_save(d)
                vf._write_bas(case['bas'], vf._MODULE, fixed)
                print(f"  → 道具が直しました（要登録）: {ask2}")
                continue
        print("  → 道具では直せません（agent --forge で）")
print(f"頭の行の検査: {len(live)} 本　違反 {bad}" + (f"（引退 {len(d) - len(live)} 本は見ない）" if len(d) != len(live) else ''))
