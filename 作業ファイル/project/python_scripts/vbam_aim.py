# -*- coding: utf-8 -*-
"""vbam_aim.py — レントゲンの地図（vbam_xray）で「棚を撃ってよいか・どこを選ぶか」を決める（Excelドック 段2）。

人の言葉で棚が決まったあと、撃つ前に地図を見て答える:
  fire    … 症状がこの表にある。撃つ所（select）を返す
  skip    … 症状が無い（直す所が無い）。撃たない
  refuse  … 症状はあるが危ない条件に当たる。撃たずに理由を 1 行で返す
  unknown … 地図が「分からない」所に当たる。撃たずに AI へ回す
  (棚の表に無い棚は answer=None＝この仕組みは何も言わない)

撃つ前と後に地図を撮り、壊していないか（見出し・合計・隠した行列）を verify() が確かめる。
データ: _AIM_ROWS（棚 ／ 種類 ／ 症状コード ／ 危ない条件コード。棚の頭の行と本文を読んだ表から）。
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
_AIMS = {}

# 症状ごとの「撃つ範囲」の種類: col＝症状のある列の明細だけ／table＝表全体／sheet・book＝シート・ブック全体
_SCOPE = {'S01': 'col', 'S02': 'col', 'S03': 'col', 'S04': 'col', 'S05': 'col', 'S06': 'col', 'S07': 'col', 'S09': 'col',
          'S14': 'col', 'S17': 'col', 'S18': 'col', 'S20': 'col', 'S21': 'col', 'S27': 'col', 'S28': 'col', 'S40': 'col',
          'S10': 'col', 'S19': 'col'}
# 行を消す・詰める・動かす棚（行の位置が変わると他の式・合計が壊れる）
_MOVES_ROWS = re.compile(r'削除|詰める|畳む|連番|並べ|挿入|取り込|移す')
# 値を書き換える棚（式のセル・参照されるセルを潰す）
_WRITES_VALUES = re.compile(r'上書き|にする|を数に|を日付|埋める|取る|そろえる|直す|戻す|空にする|ひらがな|半角')


_AIM_ROWS = [
    ('表の書き方と罫線と列幅をそろえる', '直す', 'S01 S02 S03 S05 S06 S07 S08 S27 S31 S32 S33 S34 S45', 'D01 D02 D03 D04 D05 D06 D07 D09 D10 D13 D18 D19 D21'),
    ('全列が同じ重複行を削除する', '直す', 'S11', 'D01 D02 D03 D06 D08 D12'),
    ('表の下に合計行を足す', '作る', 'S30', 'D01 D02 D04 D09 D12'),
    ('選んだコード列を名称で上書きする', '直す', 'S40', 'D06 D08 D09 D11'),
    ('結合セルをほどいて値を埋める', '直す', 'S15', 'D02 D03 D12'),
    ('選んだ列の空白を上の値で埋める', '直す', 'S14', 'D01 D03 D11'),
    ('選んだ列の文字の日付を日付にする', '直す', 'S01', 'D09 D11 D20'),
    ('選んだ列の文字の数字を数にする', '直す', 'S02', 'D04 D09 D10 D11'),
    ('表の中の空行を削除して詰める', '直す', 'S12', 'D02 D03 D08 D12'),
    ('二段の見出しを一行に畳む', '直す', 'S16', 'D03 D08'),
    ('選んだ列の途切れた式を戻す', '直す', 'S17', 'D01 D11'),
    ('テーブルの範囲を下と右へ広げる', '直す', 'S36', 'D01 D02 D12 D14'),
    ('表の中の空の列を削除して詰める', '直す', 'S13', 'D02 D07 D08 D12'),
    ('乱れた列を掃除し右に足す', '直す', 'S01 S02 S03', 'D02 D08 D12'),
    ('貼ったCSVを使える形に整える', '直す', 'S01 S02 S03 S31', 'D09 D13'),
    ('外部リンクを値に変えて切る', '直す', 'S23', 'D16 D17 D11'),
    ('名前を一覧にし壊れた名前を削除', '直す', 'S24', 'D08 D16'),
    ('余分な行列を削除し軽くする', '直す', 'S25', 'D08 D13 D14'),
    ('住所と郵便番号の形をそろえる', '直す', 'S07', 'D06 D09'),
    ('図形を一覧にし位置をそろえる', '操作', 'S39', 'D15 D22'),
    ('隠れた行と列を全部表示する', '操作', 'S35', 'D07'),
    ('文字の先頭の空白を削除する', '直す', 'S03', 'D05 D06 D11'),
    ('空白だけのセルを空にする', '直す', 'S04', 'D11'),
    ('選んだ列が空欄の行を削除する', '直す', 'S12', 'D02 D03 D08 D12'),
    ('他ブック参照を自ブックに直す', '直す', 'S23', 'D16 D08'),
    ('番号の列を連番に振り直す', '直す', 'S10', 'D06 D08 D11 D01'),
    ('番号に先頭のゼロを付けてそろえる', '直す', 'S09', 'D09 D11'),
    ('空白の有無を多い方にそろえる', '直す', 'S06', 'D05 D06'),
    ('表の合計と小計の行を消す', '直す', 'S46', 'D01 D02 D08 D12'),
    ('見出しで枠を固定しA1に戻す', '操作', 'S47', 'D02 D13'),
    ('選んだコード列の右に名称列を足す', '作る', 'S40', 'D02 D08'),
    ('表記ゆれをあいまい判定してマスタを引き当てる', '作る', 'S26', 'D02 D06'),
    ('スピルエラーを検出し邪魔なセルをクリアする', '直す', 'S29', 'D03 D11 D17'),
    ('英数を半角にカナを全角にする', '直す', 'S05', 'D06 D10 D15'),
    ('セルの見えない空白と改行を取る', '直す', 'S03', 'D05 D06 D11'),
    ('ふりがなの空欄を埋める', '直す', 'S27', 'D06'),
    ('選んだ列のカナをひらがなにする', '直す', 'S28', 'D06 D11'),
    ('氏名の空白を全角1つにそろえる', '直す', 'S06', 'D06'),
    ('帳票を右に1行1件の一覧にする', '直す', 'S37', 'D02 D03'),
    ('ピボットの元範囲を広げて更新する', '作る', 'S38', 'D01 D14'),
    ('グラフの書体と凡例をそろえる', '操作', 'S48', 'D13 D14'),
    ('横で見出し毎ページの印刷設定', '操作', 'S50', 'D13 D21'),
    ('ピボットをいつもの形にそろえる', '操作', 'S49', 'D13 D14'),
    ('ゼロで割るエラーの式を直す', '直す', 'S18', 'D11'),
    ('合計の式の範囲を表の終わりまで広げる', '直す', 'S19', 'D01 D08'),
    ('下へ写した式の絶対参照の抜けを直す', '直す', 'S20', 'D08 D11'),
    ('VLOOKUPの列番号を見出しから引く形にする', '直す', 'S21', 'D08 D11'),
]


def _load():
    if _AIMS:
        return _AIMS
    for name, kind, sym, dz in _AIM_ROWS:
        _AIMS[name] = {'kind': kind, 'sym': sym.split(), 'danger': dz.split()}
    return _AIMS


def has_aim(shelf):
    return shelf in _load()


def _rows_of(spec):
    """'9-50' / '9' / '9-12,20' → [(a, b)]"""
    out = []
    for part in str(spec or '').split(','):
        part = part.strip()
        if not part:
            continue
        a, _, z = part.partition('-')
        try:
            out.append((int(a), int(z or a)))
        except ValueError:
            pass
    return out


def _addr_parts(addr):
    m = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?', str(addr or ''))
    if not m:
        return None
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3) or m.group(1), int(m.group(4) or m.group(2))
    return _cn(c1), r1, _cn(c2), r2


def _cn(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def _cl(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _detail_span(b):
    """ブロックの明細の最初と最後の行（無ければ None）"""
    ds = [x for run in b.get('rows', []) if run.get('role') in ('D', 'F') for x in _rows_of(run['rows'])]
    if not ds:
        return None
    return min(a for a, _ in ds), max(z for _, z in ds)


def _blocks(m):
    return {b['id']: b for b in m.get('blocks', [])}


def _targets(m, codes):
    """地図の症状のうち codes に当たる物 → [{code, scope, block, col, rows}]"""
    out = []
    for s in m.get('symptoms', []):
        if s['code'] in codes:
            out.append(s)
    return out


def _overlap_rows(a1, a2, b1, b2):
    return not (a2 < b1 or b2 < a1)


def _hidden_in(m, rng):
    p = _addr_parts(rng)
    if not p:
        return False
    c1, r1, c2, r2 = p
    hr = m.get('hidden', {})
    for a, b in _rows_of(','.join(hr.get('rows', []) if isinstance(hr.get('rows'), list) else [])):
        if _overlap_rows(a, b, r1, r2):
            return True
    for c in hr.get('cols', []) or []:
        n = _cn(re.sub(r'\d', '', str(c)))
        if c1 <= n <= c2:
            return True
    return False


def _danger(m, b, shelf, aim, target_cols, target_rows):
    """危ない条件に当たるか → 理由（無ければ ''）。地図で言える条件だけを見る（言えない条件は棚の中の守りに任せる）。"""
    dz = set(aim['danger'])
    moves = bool(_MOVES_ROWS.search(shelf)) or aim['kind'] == '作る' and 'D01' in dz
    safety = b.get('safety', {}) if b else {}
    rng = _addr_parts(b['range']) if b else None
    # D05 字下げで階層を表す列の先頭の空白は意味がある
    if 'D05' in dz and target_cols:
        for ind in safety.get('indent', []):
            if _cn(ind['col']) in target_cols:
                return f"{ind['col']}列は字下げで階層を表しているので撃ちません（先頭の空白に意味がある）"
    # D14 テーブル・ピボット
    if 'D14' in dz and (m.get('objects', {}).get('listobjects') or m.get('objects', {}).get('pivots')):
        return 'テーブルかピボットがあるので撃ちません（範囲が壊れる）'
    # D01 合計・小計が明細の途中
    if 'D01' in dz and moves and safety.get('subtotal_inside'):
        return '合計・小計の行が表の途中にあるので撃ちません（行を動かすと合計が壊れる）'
    # D03 結合セル
    if 'D03' in dz and 'ほどい' in shelf and b:
        if b.get('kind') == 'form':
            return '帳票（1件が複数行）の結合なので撃ちません（ほどくと意味が壊れる）'
        mh = safety.get('merged_in_header') or []
        if any((_addr_parts(x) or (0, 0, 0, 0))[2] > (_addr_parts(x) or (0, 0, 0, 0))[0] for x in mh):
            return '見出しの横結合があるので撃ちません（ほどくと見出しが崩れる）'
        if mh and not safety.get('merged_in_detail'):
            return '結合は見出しだけなので撃ちません（二段の見出しは「一行に畳む」棚で）'
    elif 'D03' in dz and safety.get('merged_count'):
        if re.search(r'削除|詰める|重複|並べ|挿入', shelf) and safety.get('merged_in_detail'):
            # 消す行・動く行が結合にかかるときだけ断る（結合が別の行にあるだけなら、行を詰めても結合は崩れない）
            if target_rows is None or any(_overlap_rows(x, y, target_rows[0], target_rows[1]) for x, y in _rows_of(safety.get('merged_detail_rows'))):
                return '消す行に結合セルがかかっているので撃ちません（行を動かすと結合が崩れる）'
        for mg in safety.get('merged', []):
            p = _addr_parts(mg)
            if p and target_rows and _overlap_rows(p[1], p[3], target_rows[0], target_rows[1])                     and (not target_cols or any(p[0] <= c <= p[2] for c in target_cols)):
                if b and _rows_of(b.get('detail_rows')) and any(_overlap_rows(p[1], p[3], x, y) for x, y in _rows_of(b.get('detail_rows'))):
                    return '撃つ範囲に結合セルがあるので撃ちません（結合をほどく棚を先に）'
    # D07 隠した行・列
    if 'D07' in dz and b and _hidden_in(m, b['range']):
        return '表の中に隠した行か列があるので撃ちません（意図して隠している）'
    # D08 他の式が参照している
    if 'D08' in dz and moves and safety.get('referenced_total') and '空行' not in shelf:      # 空の行を消すだけなら、範囲の式（SUM など）は Excel が詰めてくれる
        for ref in safety.get('referenced', []):
            p = _addr_parts(ref['range'])
            if p and target_rows and _overlap_rows(p[1], p[3], target_rows[0], target_rows[1]):
                return '他の式が参照している行があるので撃ちません（行を動かすと参照が壊れる）'
    # D02 隣に別の表・メモ（全面処理で巻き込む）
    if 'D02' in dz and moves and rng:
        for o in m.get('blocks', []):
            if b and o['id'] == b['id']:
                continue
            q = _addr_parts(o['range'])
            if q and _overlap_rows(q[1], q[3], rng[1], rng[3]) and o.get('kind') in ('table', 'form') \
                    and (q[0] > rng[2] and q[0] - rng[2] <= 3 or q[2] < rng[0] and rng[0] - q[2] <= 3):
                return '隣に別の表があるので撃ちません（行を動かすと隣の表が崩れる）'
    return ''


def aim(m, shelf):
    """地図 m と棚の名前 → {'verdict', 'reason', 'select': [番地…], 'blocks': […]}。棚の表に無ければ None。"""
    a = _load().get(shelf)
    if a is None:
        return None
    codes = set(a['sym'])
    bl = _blocks(m)
    sym = _targets(m, codes)
    # 作る・操作の棚は症状が「在る」ではなく「無い」ことが当たる条件（S30 合計が無い）等: 地図が言えるのは S30 だけ
    if not sym:
        # unknown の塊が表の位置にあるときは、症状が無いとは言い切れない
        if any(b.get('kind') == 'unknown' for b in m.get('blocks', [])) and not any(b.get('kind') in ('table', 'form') for b in m.get('blocks', [])):
            return {'verdict': 'unknown', 'reason': '表の形が分からない（地図が unknown）ので撃たずに AI へ', 'select': [], 'blocks': []}
        # 表のすぐ近くに「分からない」塊があるなら、症状が無いとは言い切れない（割れた表の続きかもしれない）
        for tb in m.get('blocks', []):
            if tb.get('kind') not in ('table', 'form'):
                continue
            p = _addr_parts(tb['range'])
            for ub in m.get('blocks', []):
                q = _addr_parts(ub['range']) if ub.get('kind') == 'unknown' else None
                if q and q[1] <= p[3] + 3 and q[3] >= p[1] - 3 and q[0] <= p[2] + 1 and q[2] >= p[0] - 1:
                    return {'verdict': 'unknown', 'reason': f"{ub['range']} の形が地図で分からず、症状が無いとは言い切れないので撃たずに AI へ", 'select': [], 'blocks': []}
        return {'verdict': 'skip', 'reason': '症状（' + '・'.join(sorted(codes)) + '）がこの表に無いので撃ちません', 'select': [], 'blocks': []}
    sel, used, why = [], [], ''
    by_block = {}
    for s in sym:
        by_block.setdefault(s.get('block'), []).append(s)
    for bid, ss in by_block.items():
        b = bl.get(bid)
        if b is None:                              # ブック・シート全体の症状
            used.append(None)
            continue
        if b.get('kind') == 'unknown':
            return {'verdict': 'unknown', 'reason': f"{b['range']} は表かどうか地図が分からないので撃たずに AI へ", 'select': [], 'blocks': []}
        span = _detail_span(b)
        cols = sorted({s['col'] for s in ss if s.get('col')})
        trows = span or (None)
        tcols = [_cn(c) for c in cols]
        srows = [x for s0 in ss if s0.get('scope') == 'table' for x in _rows_of(s0.get('rows'))]
        r = _danger(m, b, shelf, a, tcols, (min(x for x, _ in srows), max(y for _, y in srows)) if srows and _MOVES_ROWS.search(shelf) else trows)
        if r:
            why = why or r
            continue
        used.append(b)
        if cols and span and any(_SCOPE.get(s['code']) == 'col' for s in ss):
            for c in cols:
                sel.append(f'{c}{span[0]}:{c}{span[1]}')
        elif b.get('range'):
            sel.append(b['range'])
    if not used and why:
        return {'verdict': 'refuse', 'reason': why, 'select': [], 'blocks': []}
    return {'verdict': 'fire', 'reason': '', 'select': sel, 'blocks': [b['id'] for b in used if b]}


def verify(before, after, shelf=''):
    """撃つ前後の地図 → 壊した所の一覧（空なら無事）。見出し・合計・隠した行列だけを見る。"""
    out = []

    ok_header_change = '畳む' in shelf or '二段' in shelf
    bh = [len(b.get('header_rows') or []) for b in before.get('blocks', []) if b.get('kind') in ('table', 'form')]
    ah = [len(b.get('header_rows') or []) for b in after.get('blocks', []) if b.get('kind') in ('table', 'form')]
    if not ok_header_change and bh and ah and sum(ah) < sum(bh):
        out.append('見出しの行が減った')
    bt = sum(len(b.get('total_rows') or []) for b in before.get('blocks', []))
    at = sum(len(b.get('total_rows') or []) for b in after.get('blocks', []))
    if at < bt and not re.search(r'合計|小計', shelf):
        out.append('合計・小計の行が減った')
    bhid = (len(before.get('hidden', {}).get('rows', []) or []), len(before.get('hidden', {}).get('cols', []) or []))
    ahid = (len(after.get('hidden', {}).get('rows', []) or []), len(after.get('hidden', {}).get('cols', []) or []))
    if ahid != bhid and '表示' not in shelf:
        out.append('隠した行・列が変わった')
    return out
