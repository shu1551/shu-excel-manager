# -*- coding: utf-8 -*-
"""見せていない種の表を作って、鍛えたマクロを当てる（AI なし・2026-09-18）。
  py unseen_sweep.py <種> [仕事の名前 …]     … 名前を省くと下の一覧ぜんぶ
選んでいる列の覚書（*_選択.txt）も見るので try_all.py を通す。
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
JOBS = [
    ('帳票一覧', '帳票一覧'), ('課別月別ピボット', 'ピボット月別'), ('PQ課別月別', 'PQクロス'),
    ('PQ支出課名', 'PQ結合'), ('突合', '突合せ'), ('比較', '２表比較'), ('振り直し', '区分振り直し'),
    ('合計行', '合計行'), ('集約', 'シート集約'), ('按分', '按分表'), ('月次集計', '月次集計'),
    ('見やすい棒グラフ', '棒グラフ'), ('見やすい推移グラフ', '推移グラフ'), ('構成比円', '円グラフ'),
    ('予算執行複合グラフ', '複合グラフ'), ('課別科目別積み上げ', '積み上げグラフ'),
    ('職員名簿テーブル', 'テーブル化'), ('PQ縦持ち', '縦持ち'), ('ピボット範囲更新', 'ピボット範囲更新'),
    ('課別科目別ピボット', 'ピボット行列'), ('科目別構成比ピボット', '構成比ピボット'),
    ('課名を引く', 'マスタ参照'), ('職員名簿の絞り込み', 'テーブル絞り込み'), ('課別ダッシュボード', 'ダッシュボード'), ('PQ課別集計', 'PQ追加集計'),
    # B 集計 8 本（2026-09-18）
    ('項目別合計', '項目別合計'), ('項目別件数', '項目別件数'), ('累計列', '累計列'), ('構成比列', '構成比列'),
    ('増減列', '増減列'), ('年度月四半期', '年度月四半期'), ('上位5件', '上位5件'), ('小計行', '小計行'),
    # C 突合 3・A 掃除 2（2026-09-18）
    ('重複一覧', '重複一覧'), ('コード名称', 'コード名称'), ('空欄一覧', '空欄一覧'),
    ('結合解除', '結合解除'), ('空白埋め', '空白埋め'),
    # 残り 12 本（2026-09-18 午後）
    ('日付直し', '日付直し'), ('数値直し', '数値直し'), ('空行削除', '空行削除'), ('二段見出し', '二段見出し'),
    ('順位列', '順位列'), ('平均最大最小', '平均最大最小'), ('区間集計', '区間集計'), ('列並べ替え', '列並べ替え'),
    ('検算', '検算'), ('二重計上', '二重計上'), ('日付妥当性', '日付妥当性'), ('年齢勤続', '年齢勤続'),
    # 第二期 E グラフ 6 本（2026-09-18 夕）
    ('スパークライン', 'スパークライン'), ('散布図', '散布図'), ('レーダー', 'レーダー'), ('二本棒', '二本棒'),
    ('増減滝', '増減滝'), ('グラフ体裁', 'グラフ体裁'),
    # 第二期 D 帳票ほか 5 本（2026-09-18 夕）
    ('差し込み', '差し込み'), ('印刷設定', '印刷設定'), ('様式転記', '様式転記'), ('シート分割', 'シート分割'),
    ('ピボット値貼り', 'ピボット値貼り'),
    # 第二期 G パワークエリ 3 本（2026-09-18 夕）
    ('PQ列分割', 'PQ列分割'), ('PQ文字掃除', 'PQ文字掃除'), ('PQ更新', 'PQ更新'),
    # 第二期 残り 11 本（2026-09-18 夜・式とフィルタの目を足して）
    ('式を値に', '式を値に'), ('式の途切れ', '式の途切れ'), ('エラー一覧', 'エラー一覧'), ('桁ずれ', '桁ずれ'),
    ('フィルタ付け', 'フィルタ付け'), ('テーブル拡張', 'テーブル拡張'), ('テーブル解除', 'テーブル解除'), ('空列削除', '空列削除'),
    ('名寄せ', '名寄せ'), ('ピボット累計', 'ピボット累計'), ('ピボット体裁', 'ピボット体裁'),
    ('一覧を複数行帳票に展開する', '複数行帳票'),
]


def run(args, cwd=HERE):
    r = subprocess.run([sys.executable, '-u'] + args, cwd=cwd, capture_output=True, text=True,
                       encoding='utf-8', errors='replace', stdin=subprocess.DEVNULL)
    return (r.stdout or '') + (r.stderr or '')


def live_jobs():
    """一覧のうち、台帳で引退していない仕事（2026-09-18: 引退した仕事はお題を残したまま撃たない）。"""
    sys.path.insert(0, os.path.abspath(os.path.join(HERE, '..', 'python_scripts')))
    import vbam_forge as vf
    d = vf._forge_load()
    return [(n, f) for n, f in JOBS if not (d.get(n) or {}).get('retired')]


def main():
    seed = sys.argv[1]
    want = sys.argv[2:]
    jobs = [(n, f) for n, f in live_jobs() if not want or n in want]
    bad = []
    for name, folder in jobs:
        out = os.path.join(HERE, folder, f"未見_種{seed}")
        if not os.path.isdir(out) or not os.listdir(out):
            log = run([os.path.join(HERE, folder, 'make_pairs.py'), '--unseen', seed], cwd=os.path.join(HERE, folder))
            if not os.path.isdir(out):
                print(f"× {name}: 未見の表を作れませんでした\n" + log[-800:])
                bad.append(name)
                continue
        log = run([os.path.join(HERE, 'try_all.py'), name, out, '3'])
        tail = [l for l in log.splitlines() if l.startswith('---')]
        ng = [l for l in log.splitlines() if l.startswith('外れ')]
        print(f"{'○' if not ng else '×'} {name}: " + (tail[-1] if tail else log[-300:]))
        for l in ng:
            print('    ' + l[:200])
        if ng or not tail:
            bad.append(name)
            for l in log.splitlines():
                if l.startswith('    '):
                    print('    ' + l.strip()[:220])
    print(f"=== 未見の種{seed}: {len(jobs)} 本中 外れ {len(bad)} 本" + ("（" + "・".join(bad) + "）" if bad else ""))
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
