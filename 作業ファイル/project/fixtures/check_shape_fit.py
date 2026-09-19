# -*- coding: utf-8 -*-
"""配った登録簿の ' 形: が、未見の表でもそろうかを読むだけで確かめる（Excel も AI も使わない・2026-09-18）。

形は「入口が語の前に絞る必要条件」なので、外れ方は 1 つだけ＝その表に形が足りず、撃てるはずの仕事が撃たれない。
実射（entry_check_unseen）で 30 分かけて見ていたのはこの 1 点なので、openpyxl で読んで数秒で済ませる。

  py check_shape_fit.py [種 …]      … 種を省くと、その仕事の 未見_種* を全部
"""
import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', 'python_scripts')))
sys.path.insert(0, HERE)
import vbam_core   # noqa: E402
vbam_core.setup_encoding()
import vbam_forge as vf      # noqa: E402
import vbam_lineage as vl    # noqa: E402
import vbam_prefire as vp    # noqa: E402
from unseen_sweep import live_jobs   # noqa: E402


def main():
    seeds = [a for a in sys.argv[1:] if not a.startswith('--')]
    xlam = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'AddIns', '秀コンボ.xlam')
    reg = vp.registry_from_text(vl._read_closed_book(xlam).get('表の整理') or '', owner=os.path.basename(xlam))
    shape_of_sub = {e['name']: e['shape'] for e in reg}
    ledger = vf._forge_load()
    n = bad = 0
    for name, folder in live_jobs():
        sub = (ledger.get(name) or {}).get('sub')
        want = shape_of_sub.get(sub)
        if want is None:
            print(f"× {name}: 配った登録簿に「{sub or '（未登録）'}」がありません")
            bad += 1
            continue
        pats = [os.path.join(HERE, folder, f"未見_種{s}", '*_前.xlsx') for s in seeds] or \
               [os.path.join(HERE, folder, '未見_種*', '*_前.xlsx')]
        books = sorted(b for p in pats for b in glob.glob(p))
        miss = []
        for b in books:
            n += 1
            got = vf.shape_of_file(b)
            lack = [w for w in want if w not in got]
            if lack:
                miss.append((os.path.basename(b), lack))
        if miss:
            bad += len(miss)
            print(f"× {name}（{sub}・形: {' '.join(want) or 'なし'}）: {len(miss)}/{len(books)} 枚で形が足りません")
            for b, lack in miss[:4]:
                print(f"    {b}: {'・'.join(lack)} が無い")
        else:
            print(f"○ {name}（形: {' '.join(want) or 'なし'}）: 未見 {len(books)} 枚すべてでそろう")
    print(f"形の当たり: 未見の表 {n} 枚　撃てなくなる組 {bad}")
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
