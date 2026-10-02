# -*- coding: utf-8 -*-
"""題 25・26（遅いマクロ）の秒を測る（2026-10-02）。開いている bench_tanaka.xlsm のデータを作り直してからマクロを撃つ。

  py time_slow.py 田中削除25 [別のマクロ名 ...]   → 名前ごとに「秒・残った行数（25）／埋まった単価の数（26）」
正しさ: 25＝田中の行が 0・ほかの行は 7,500 行のまま。26＝1,000 行すべてに単価（記号の番号×10）。
"""
import os
import sys
import time

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from make_bench import fill_slow  # noqa: E402

BOOK = os.path.join(HERE, 'bench_tanaka.xlsm')


def check(wb, name):
    if '25' in name:
        ws = wb.Worksheets('25削除')
        vals = ws.Range('A2:A10001').Value
        names = [r[0] for r in vals if r[0] is not None]
        return f"残り {len(names)} 行・田中 {names.count('田中')} 行（田中は 0 が正解・ほかは 6000 行）"
    ws = wb.Worksheets('26値貼付')
    rows = ws.Range('A2:B1001').Value
    ok = sum(1 for k, v in rows if v is not None and int(v) == int(k[1:]) * 10)
    return f"正しい単価 {ok} / 1000"


def main():
    wb = win32com.client.GetObject(BOOK)
    xl = wb.Application
    for name in sys.argv[1:]:
        fill_slow(wb)
        sheet = '25削除' if '25' in name else '26値貼付'
        wb.Worksheets(sheet).Activate()
        t = time.time()
        try:
            xl.Run(f"'{wb.Name}'!{name}")
        except Exception:
            pass                        # 長い呼び出しは COM が「システム コールに失敗」を返す＝終わるまで待って測る
        while True:
            try:
                res = check(wb, name)
                break
            except Exception:
                time.sleep(0.5)
        print(f"{name}: {time.time() - t:.2f} 秒  {res}")


if __name__ == '__main__':
    main()
