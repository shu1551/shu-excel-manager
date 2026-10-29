# -*- coding: utf-8 -*-
"""vbam_prefire.py — vba_manager 分割パート: マクロの先撃ち（agent の入口で、AI より先に「表の書き方と罫線と列幅をそろえる」を撃つ）

2026-09-11 深夜。試験（agent --exam）で、表の整理はマクロ（表の整理.bas・AI なし）が 45 題全問合格・1 題 0.1〜0.5 秒、
agent（Sonnet）は 1 題平均 22 秒だった。shu と決めた形: 入口は agent のまま、中の順番を「マクロが先、Sonnet は例外」に。
  1. 依頼が書き方をそろえる整理（そろえ・統一・整え・整理・表記・きれい…）で、開いているブックかアドインに
     「表の書き方と罫線と列幅をそろえる」があれば、道具が撃つ（依頼が重複を消すことを承認していれば「全列が同じ重複行を削除する」も）。控えは本番と同じ。
  2. 道具の tidy で仕上げ、終わりの検査（仕上げ検査・中身・依頼の語の列の残り・消えた行・頼んでいない変化）を当てる。
  3. 残りが無く、依頼にマクロの扱わない仕事（合計・並べ替え・グラフ…）も無ければ、そこで合格（AI は呼ばない）。
     残りがあれば、その指摘を材料に足して Sonnet の往復へ（控え・関所・undo はそのまま効く）。
どちらを撃つかは依頼の語で機械的に決まる（AI にも会話の Claude にも判断させない）。

2026-09-17 鍛える回路（vbam_forge）: モジュール「表の整理」の Sub の頭 2 行（' 依頼の語: 正規表現 ／ ' 扱う: 語 語）が
登録簿。依頼に「依頼の語」が当たる Sub は「表の書き方と罫線と列幅をそろえる」の後に撃ち、その「扱う」の語は AI に回す仕事（_ASK_OTHER_RE）
から外す。登録簿は .bas そのもの（別の台帳を持たない＝データはファイル自身）。
"""
import contextlib
import os
import re
import tempfile
import time
from vbam_core import _col_letter, shelf_text, shelf_texts, shelf_module_of

import vbam_agent as va

_MACRO_TIDY = '表の書き方と罫線と列幅をそろえる'
_MACRO_DEDUPE = '全列が同じ重複行を削除する'
_MACRO_MODULE = '表の整理'
_ASK_TIDY_RE = re.compile(r'そろえ|揃え|統一|整え|整理|きれいに|綺麗に|掃除|書き方|表記')
# マクロの扱わない仕事（あれば、マクロの後に Sonnet の往復へ回す）
_ASK_OTHER_RE = re.compile(r'合計|集計|小計|ピボット|グラフ|並べ替|並び替|ソート|抽出|突き合|照合|列を足|列を追加|列を挿入|'
                           r'行を足|行を追加|行を挿入|追加し|印刷|PDF|色を|色分け|色付け|網掛け|塗り|条件付き|数式|関数|VLOOKUP|テーブルに|'
                           r'フィルタ|図形|写真|画像|埋め|計算|平均|件数|分けて|分割|区分|判定|比較|差分|ふりがなを|フリガナを付')
_DUP_WORD_RE = re.compile(r'重複|二重|ダブ|同じ内容|同じ会員|同じ品番|同じ行')
# マクロの扱わない仕事のうち、表の中身の仕事と別の操作（同じ節に「扱う」の語があっても片づけない）
_STRONG_OTHER_RE = re.compile(r'グラフ|並べ替|並び替|ソート|印刷|PDF|色を|色分け|色付け|網掛け|塗り|条件付き|抽出|フィルタ|図形|写真|画像')
_NO_DELETE_RE = re.compile(r'(消|削除)さない|(消|削除)しない|残して')


_REG_ASK_RE = re.compile(r"^\s*'\s*依頼の語\s*[:：]\s*(.+?)\s*$")
_REG_COMBO_RE = re.compile(r"^\s*'\s*依頼の組\s*[:：]\s*(.+?)\s*$")
_REG_HANDLE_RE = re.compile(r"^\s*'\s*扱う\s*[:：]\s*(.+?)\s*$")
_REG_HEAD_RE = re.compile(r"^\s*'\s*見出し\s*[:：]\s*(.+?)\s*$")
_REG_SELECT_RE = re.compile(r"^\s*'\s*選ぶ列\s*[:：]\s*(\d+)\s*(?:[-〜~]\s*(\d+))?\s*$")
_REG_SHAPE_RE = re.compile(r"^\s*'\s*形\s*[:：]\s*(.+?)\s*$")
# 表の形の語（登録簿の ' 形: に書ける語。入口はシートの形を読んで、そろっている仕事だけを候補にする・2026-09-18 朝）
SHAPE_WORDS = ('数の列', '日付の列', '率の列', 'テーブル', '別のシート', '同じ見出しのシート', '共通の見出しのシート', '結合の見出し', 'ピボット', 'グラフ')
# ピボット・グラフ＝アクティブなシートにそれがあるか（2026-09-18 第二期: ピボットを値で貼る・累計・見せ方、グラフの見せ方の仕事が、
# ピボットやグラフの無いシートでも依頼の語が当たると撃たれて何もせずに終わり、その依頼が AI に回らなかった）。
# 鍛える側は本番の表にあるときだけ必要条件にする（何もしない試験の表があっても落とさない・vbam_forge._job_shape）
OBJECT_SHAPE_WORDS = ('ピボット', 'グラフ')
_SHEET_WORD_ROWS = 15      # 見出しの語を探す、シートの上からの行数
_REG_SUB_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+([^\s\(]+)\s*\(', re.I)
_REG_HEAD_LINES = 6        # Sub の行の下、この行数までにある頭の 2 行を登録簿として読む


def registry_from_text(text, owner=None):
    """モジュールの本文 → 登録簿 [{'name','owner','ask'(正規表現の文字),'handles'([語]),'headers'([語])}]（純 Python）。
    「表の書き方と罫線と列幅をそろえる」「全列が同じ重複行を削除する」は道具が持つ既定なので入れない。頭の 2 行が無い Sub も入れない。
    ' 見出し: の行（任意）＝そのマクロが要る見出しの語。シートに全部あるときだけ撃つ（2026-09-17 夕）。
    ' 見出し: なし ＝語に頼らないマクロ（2026-09-18）。headers は []。"""
    out = []
    lines = str(text or '').splitlines()
    for i, ln in enumerate(lines):
        m = _REG_SUB_RE.match(ln)
        if not m or m.group(1) in (_MACRO_TIDY, _MACRO_DEDUPE):
            continue
        ask = handles = None
        combo = None
        headers = []
        select = None
        shape = []
        for h in lines[i + 1:i + 1 + _REG_HEAD_LINES]:
            a = _REG_ASK_RE.match(h)
            b = _REG_HANDLE_RE.match(h)
            c = _REG_HEAD_RE.match(h)
            s = _REG_SELECT_RE.match(h)
            k = _REG_SHAPE_RE.match(h)
            g = _REG_COMBO_RE.match(h)
            if a:
                ask = a.group(1)
            elif g:
                # ' 依頼の組: A,B+C,D-E ＝語の組（+ で組・, で言い換え・- で除外）。VBA の先撃ちだけが読んでいた（2026-09-23）
                combo = g.group(1)
            elif b:
                handles = b.group(1).split()
            elif c:
                # 「なし」＝見出しの語に頼らないマクロ（表の形で列を探す・2026-09-18）＝どの表にも撃つ
                headers = [] if c.group(1).strip() == 'なし' else c.group(1).split()
            elif s:
                # ' 選ぶ列: N（や N-M）＝選んでいる列を使うマクロ。入口は依頼の語と見出しが一致する列を選んでから撃つ（2026-09-18）
                select = (int(s.group(1)), int(s.group(2) or s.group(1)))
            elif k:
                # ' 形: 語 語 ＝この仕事に要る表の形（数の列・日付の列・テーブル・同じ見出しのシート…）。「なし」は []
                shape = [] if k.group(1).strip() in ('なし', '無し') else [w for w in k.group(1).split() if w in SHAPE_WORDS]
            elif not h.strip().startswith("'"):
                break
        if not ask or not handles:
            continue
        try:
            re.compile(ask)
        except re.error:
            continue
        out.append({'name': m.group(1), 'owner': owner, 'ask': ask, 'handles': handles, 'headers': headers,
                    'select': select, 'shape': shape, 'combo': combo})
    return out


def shelf_entries_from_text(text):
    """モジュールの本文 → 先撃ちの採点に使う登録 [{'name','ask','combo'}]（純 Python・2026-09-23）。

    registry_from_text と違って**絞り込まない**（「表の書き方と罫線と列幅をそろえる」も頭の 2 行が無い Sub も入れる）。
    VBA の コンボ道具.棚を読む と同じものを読む＝Python 側の先撃ちが VBA と同じ答えを出せるようにする。
    """
    out, name = [], None
    for ln in str(text or '').splitlines():
        m = _REG_SUB_RE.match(ln)
        if m:
            name = m.group(1)
            out.append({'name': name, 'ask': None, 'combo': None, 'tier': '', 'allow_fix': False, 'plan_substitute': '', 'allow_change': '', 'look_only': False, 'cf_ok': False})
            continue
        if not out:
            continue
        a = _REG_ASK_RE.match(ln)
        g = _REG_COMBO_RE.match(ln)
        t = _REG_TIER_RE.match(ln)
        f = _REG_ALLOW_FIX_RE.match(ln)
        s = _REG_PLAN_SUB_RE.match(ln)
        c = _REG_ALLOW_CHG_RE.match(ln)
        lk = _REG_LOOK_ONLY_TAG_RE.match(ln)
        cf = _REG_CF_OK_RE.match(ln)
        if a:
            out[-1]['ask'] = a.group(1)
        elif g:
            out[-1]['combo'] = g.group(1)
        elif t:
            out[-1]['tier'] = t.group(1)
        elif f:
            out[-1]['allow_fix'] = (f.group(1).strip() in ('はい', '当てる', 'true', 'True', '1'))
        elif s:
            out[-1]['plan_substitute'] = s.group(1).strip()
        elif c:
            out[-1]['allow_change'] = c.group(1).strip()
        elif lk:
            out[-1]['look_only'] = (lk.group(1).strip() in ('はい', 'true', 'True', '1'))
        elif cf:
            # ' 条件付き書式: はい ＝条件付き書式を付ける棚。条件付き書式を付ける頼みは AI の役目だが、この印の棚だけは当てる（2026-10-22 指示書044）
            if cf.group(1).strip() == 'はい':
                out[-1]['cf_ok'] = True
    for e in out:
        _get_entry_cache(e)
    return out


# ---- 棚の当たり方の検査（2026-10-14 段 2・work_order_20261014_shelf_system.md） ----
# 10/14 の誤爆は、どれも棚を 1 本足したときの語が日常の言い方を横取りしたもの（見出しを青にして→インディゴ・この列を太字にして→
# グラフの種類を変える・この行を消して→ピボットの手直し）。問題集に無い言い方はすり抜けるので、語そのものを形で検査する（ゴンちゃん案）。
_REG_TIER_RE = re.compile(r"^\s*'\s*階級\s*[:：]\s*(\S+)")
_REG_ALLOW_FIX_RE = re.compile(r"^\s*'\s*直す頼みにも当てる\s*[:：]\s*(\S+)")
_REG_LOOK_ONLY_TAG_RE = re.compile(r"^\s*'\s*調べるだけ\s*[:：]\s*(\S+)")
_REG_CF_OK_RE = re.compile(r"^\s*'\s*条件付き書式\s*[:：]\s*(\S+)")
_REG_PLAN_SUB_RE = re.compile(r"^\s*'\s*並びの代わり\s*[:：]\s*(\S+)")
_REG_ALLOW_CHG_RE = re.compile(r"^\s*'\s*変えてよい\s*[:：]\s*(.+)")
SHELF_TIERS = ('基本', '集計', '作り替え')
# 一般の語＝どの棚の仕事にも出てくる日常の言い方。どの棚も、これ 1 つだけで点を取れない。
# 組の全部をこれだけで満たせるのは「基本」の階級の棚の決まり文句（表＋整え・重複＋消）だけ
GENERAL_WORDS = frozenset((
    '表', '表の', 'シート', 'セル', '行', '列', 'データ', 'この', 'これ', 'それ', 'ここ', 'カード', '番号', 'グラフ', '図',
    '色', '青', '赤', '緑', '黄', '黒', '白', 'グレー', '灰色', '茶色', 'ネイビー', '紺',
    '整え', '整理', 'きれい', '綺麗', '見やすく', 'そろえ', '揃え', '統一', '直', '直し', '直して', '修正', '修復',
    '変え', '替え', '消', '消し', '消して', '削除', '足し', '足して', '追加', '作', '作って', '作る', '作成', '出して', '出す',
    'にして', 'して', 'に直', '順', '並べ', '見せ', '見', 'お願い',
))
_LINT_PARTICLES = ('にして', 'して', 'を', 'に', 'で', 'の', 'は', 'が', 'と', 'も', 'へ', 'な', 'く')


def _combo_groups(combo):
    """依頼の組 → (+ の各組の言い換えの並び, 除外の語)。"""
    pos, neg = [], []
    for g in str(combo or '').split('+'):
        g = g.strip()
        if not g:
            continue
        alts = [w.strip() for w in g.lstrip('-').split(',') if w.strip()]
        (neg if g.startswith('-') else pos).append(alts)
    return pos, neg


_ASK_REGEX_RE = re.compile(r'[()\[\]{}*+?^$\\]|\.\*')
# Sub の宣言の行（引数つき）と、その下に続くコメントの行（頭の注記）。Private Sub は目録の外の手伝いなので見ない
_ARG_SUB_HEAD_RE = re.compile(r"^[ \t]*(?:Public[ \t]+)?Sub[ \t]+([^\s(]+)[ \t]*\(([^)\n]*)\)[^\n]*\n((?:[ \t]*'[^\n]*\n)*)", re.M)
_BULK_WRITE_PAT = re.compile(
    r'^(?:(?:\s*If\b.*?\bThen\s+)?)(?!\s*If\b)'
    r'(?P<lhs>[A-Za-z0-9_一-龠ぁ-んァ-ヶ\.\(\)]+?)'
    r'\.(?P<prop>Value2?|Formula2?)\s*=\s*(?P<rhs>.+)$',
    re.IGNORECASE
)
_BULK_EXCLUDE_LHS = re.compile(
    r'(?:^|[\s\.])Cells\s*\([^()]*(?:\([^()]*\)[^()]*)*\)$'
    r'|^(?:c|cel|cell|mc|newCel|tgt)$'
    r'|^(?:組|[A-Za-z0-9_一-龠ぁ-んァ-ヶ]+)\s*\(\s*\d+\s*\)$',
    re.IGNORECASE
)
_BULK_SINGLE_RANGE = re.compile(
    r'^(?:[A-Za-z0-9_一-龠ぁ-んァ-ヶ]+\.)?Range\s*\(\s*"[A-Za-z]+[0-9]+"\s*\)$',
    re.IGNORECASE
)
_BULK_OK_MARK_RE = re.compile(r"'\s*書き戻し\s*[:：]\s*新しい所")

# 税率らしい率の直書き（*1.1・*1.08・*0.1・*0.08 等）。「列の 6 割」の *0.6 のような閾値は数えない（2026-10-28 番頭点検）
_RATE_RE = re.compile(r'\*\s*(?:1\.(?:10?|08|05)|0\.(?:10?|08|05))(?![\d.])')
_published_rate_lines_cache = None


def _rate_lines(text):
    """棚の本文 → 率の直書きのあるコードの行（コメントを外して前後の空白を詰めた物）の並び（純 Python）。"""
    out = []
    for ln in str(text or '').splitlines():
        s = ln.strip()
        if not s or s.startswith("'"):
            continue
        c = s.find("'")
        code = (s[:c] if c >= 0 else s).strip()
        if _RATE_RE.search(code):
            out.append(re.sub(r'\s+', ' ', code))
    return out


def _published_rate_lines():
    """前の公開の棚（publish\\addin\\macros、無ければ公開の写しの macros）にある率の直書きの行。見つからなければ None。"""
    global _published_rate_lines_cache
    if _published_rate_lines_cache is not None:
        return _published_rate_lines_cache
    here = os.path.dirname(os.path.abspath(__file__))
    for d in (os.path.join(here, '..', 'publish', 'addin', 'macros'), os.path.join(here, '..', 'macros')):
        if not os.path.isdir(d):
            continue
        lines = []
        for fn in sorted(os.listdir(d)):
            if fn.startswith('表の整理') and fn.endswith('.bas'):
                try:
                    with open(os.path.join(d, fn), encoding='cp932', errors='replace') as f:
                        lines += _rate_lines(f.read())
                except OSError:
                    continue
        _published_rate_lines_cache = lines
        return lines
    return None


def new_rate_lines(text, base_lines=None):
    """前の公開に無かった率の直書きの行（同じ行が増えた分も数える）。base_lines=None なら前の公開を読む（純 Python に近い）。"""
    if base_lines is None:
        base_lines = _published_rate_lines() or []
    left = {}
    for b in base_lines:
        left[b] = left.get(b, 0) + 1
    out = []
    for ln in _rate_lines(text):
        if left.get(ln, 0) > 0:
            left[ln] -= 1
        else:
            out.append(ln)
    return out


def _only_general(word):
    """語が一般の語と助詞だけでできているか（「青で美化」は 美化 が残る＝固有、「見出しを青に」も 見出し が残る）。"""
    rest = str(word)
    for g in sorted(GENERAL_WORDS | set(_LINT_PARTICLES), key=len, reverse=True):
        rest = rest.replace(g, '')
    return not rest.strip()


def shelf_lint(text, base_rate_lines=None):
    """棚の本文 → 当たり方の問題 [(重さ, 棚の名前, 中身)]（純 Python）。重さ＝'止める'（register-addin で焼かない）か '注意'。
    - 依頼の語が一般の語 1 つ（「表」「グラフ」）            → 止める
    - 依頼の組を一般の語 1 つで満たせる（グラフ＋グラフ）     → 止める
    - 依頼の組の全部の組を一般の語で満たせる（この＋にして）  → 止める。ただし ' 階級: 基本 の棚は通す（表＋整え など）
    - ' 階級: 作り替え の棚が依頼の組を持つ・依頼の語が一般の語だけ → 止める
    - 率の直書き（*1.1・*1.08・*0.1・*0.08 等）が前の公開より増えた → 止める（base_rate_lines＝比べ元の行。None なら前の公開を読む）
    - 依頼の語が 1 字（「削」「取」）・正規表現の記号入り（「(空白」「.*」） → 止める
    - 階級の綴りが違う・同じ依頼の語が 2 本以上の棚にある    → 注意"""
    out, seen = [], {}
    for e in shelf_entries_from_text(text):
        name, tier = e['name'], e.get('tier') or ''
        if tier and tier not in SHELF_TIERS:
            out.append(('注意', name, f"階級「{tier}」は {'・'.join(SHELF_TIERS)} のどれでもない"))
        asks = [w.strip() for w in str(e.get('ask') or '').split('|') if w.strip()]
        for w in asks:
            seen.setdefault(w, []).append(name)
            # 先撃ちは依頼の語を字面で探す（正規表現ではない）。1 字は「削除」「取引先」のどれにも当たり、
            # 「(空白」「改行).*(消」は記号ごと探すので当たらない（2026-10-21・Gemini が足した棚 5 本の点検）
            if _ASK_REGEX_RE.search(w):
                out.append(('止める', name, f"依頼の語「{w}」に正規表現の記号がある（先撃ちは字面で探すので当たらない。| で語を並べる）"))
            elif len(w) == 1:
                out.append(('止める', name, f"依頼の語「{w}」が 1 字だけ（どの頼みにも当たる）"))
            elif w in GENERAL_WORDS:
                out.append(('止める', name, f"依頼の語「{w}」が一般の語 1 つだけ（日常の頼みを横取りする）"))
            elif tier == '作り替え' and _only_general(w):
                out.append(('止める', name, f"作り替えの棚の依頼の語「{w}」が一般の語だけ（固有の言い方にする）"))
        if not e.get('combo'):
            continue
        if tier == '作り替え':
            out.append(('止める', name, "作り替えの棚は依頼の組を持てない（依頼の語の固有の言い方だけで当てる）"))
            continue
        pos, _neg = _combo_groups(e['combo'])
        if not pos:
            continue
        cands = sorted({w for alts in pos for w in alts}, key=len)
        one = next((w for w in cands if w in GENERAL_WORDS and all(any(a in w for a in alts) for alts in pos)), None)
        if one:
            out.append(('止める', name, f"依頼の組が「{one}」1 語で満たされる（日常の頼みを横取りする）"))
            continue
        if tier != '基本' and all(any(a in GENERAL_WORDS for a in alts) for alts in pos):
            ex = '＋'.join(next(a for a in alts if a in GENERAL_WORDS) for alts in pos)
            out.append(('止める', name, f"依頼の組を一般の語だけで満たせる（例: {ex}）。決まり文句なら ' 階級: 基本"))
    for w, names in seen.items():
        if len(set(names)) >= 2:
            out.append(('注意', '・'.join(sorted(set(names))), f"依頼の語「{w}」が {len(set(names))} 本の棚にある"))

    # 引数のある Sub に依頼の語・依頼の組がある＝先撃ちでは撃たれるのに、棚の目録にも Excel のマクロ一覧にも出ない（2026-10-29）。
    # 使わない引数「勇退」を付けて隠した 32 本のうち 19 本が撃たれ続けていた。棚は引数なしの入口にし、引数は _本体 の手伝いに持たせる
    for m in _ARG_SUB_HEAD_RE.finditer(str(text or '').replace('\r\n', '\n')):
        if m.group(2).strip() and re.search(r"^\s*'\s*依頼の(?:語|組):", m.group(3), re.M):
            out.append(('止める', m.group(1), f"引数（{m.group(2).strip()}）のある Sub に依頼の語がある（撃たれるのに目録とマクロ一覧に出ない。"
                                              f"引数なしの入口と引数のある _本体 に分ける）"))

    # 範囲への丸ごと代入行（既存の表の数式・文字の数字を潰す穴）の検査（2026-10-22 指示書032）
    curr_sub = None
    for ln in str(text or '').splitlines():
        m_sub = _REG_SUB_RE.match(ln.strip())
        if m_sub:
            curr_sub = m_sub.group(1)
            continue
        if not curr_sub:
            continue
        stripped = ln.strip()
        if not stripped or stripped.startswith("'"):
            continue

        c_idx = ln.find("'")
        code = ln[:c_idx].strip() if c_idx >= 0 else stripped
        comment = ln[c_idx:] if c_idx >= 0 else ""

        # 2026-10-27 番頭点検: 試験の表の語（返却・貸出中）の決め打ちを止める（率の直書きは下で前の公開と比べる）
        if re.search(r'"[^"]*(?:返却|貸出中)[^"]*"', code):
            out.append(('止める', curr_sub, f"試験の表の語の決め打ち「{code}」（未知の表でも動く汎用形にし、試験の表の語を本体に書かない）"))

        if _BULK_OK_MARK_RE.search(comment):
            continue

        m = _BULK_WRITE_PAT.match(code)
        if not m:
            continue
        if re.search(r'^\s*If\b.*?Then\s*$', code, re.IGNORECASE):
            continue

        lhs = m.group('lhs').strip()
        rhs = m.group('rhs').strip()

        if _BULK_EXCLUDE_LHS.search(lhs):
            continue
        if _BULK_SINGLE_RANGE.match(lhs) and not re.search(r'\bArray\s*\(', rhs, re.IGNORECASE):
            continue
        if code.startswith("スピルか ="):
            continue

        out.append(('止める', curr_sub, f"範囲への丸ごと代入「{code}」（既存の範囲なら変えたセルだけ書く、新しい所なら末尾に ' 書き戻し: 新しい所）"))

    # 率の直書き（*1.1・*1.08・*0.1・*0.08 等）が前の公開より増えたら止める（2026-10-28 番頭点検・10/27 試験の表に合わせた税 1.1 倍の式）
    for code in new_rate_lines(text, base_rate_lines):
        out.append(('止める', '率の直書き', f"前の公開に無かった率の直書き「{code}」（指示書に無い決め打ちの計算を本体に足さない）"))

    return out


def head_key(s):
    """見出しの語のならし（純 Python）: 全角半角をそろえ、空白・改行を落とし、末尾の単位の括弧（（円）など）を落とす。
    2026-09-17 夜: 実物らしい表で「予算科目\\nコード」「科　目　名」「支出額（円）」の見出しに登録簿のマクロを撃たず AI に回した。"""
    import unicodedata
    t = re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(s or '')))
    # 落とすのは単位の括弧だけ（「支出額（税抜）」「予算額（前年度）」は別の列＝同じ見出しにしない）
    return re.sub(r'[\(\[](?:単位[:：]?)?(?:円|千円|百万円|人|名|件|計|合計|%|％)[\)\]]$', '', t) or t


_CELL_ADDR_RE = re.compile(r'(?<![A-Za-z0-9])([A-Z]{1,3}[1-9]\d*)(?![0-9])')


def old_findings(notes, changed):
    """指摘の文のうち、書いてある番地がどれも「変わったセル」に無いもの（純 Python）。番地の無い指摘は古いと言わない。"""
    out = []
    for n in notes or ():
        cells = set(_CELL_ADDR_RE.findall(str(n)))
        if cells and not (cells & set(changed)):
            out.append(n)
    return out


def _split_old_findings(audit, c_block, before, ws, sheet):
    """→ (audit, c_block, 元の表の指摘)。変わったセルが数えられなければ何も動かさない。"""
    try:
        changed = {r['addr'] for r in (va._changes_of(before, ws, sheet) or [])}
    except Exception:
        return audit, c_block, []
    if not changed:
        return audit, c_block, []
    old = old_findings(list(audit) + list(c_block), changed)
    return ([b for b in audit if b not in old], [b for b in c_block if b not in old], old)


def sheet_words(values):
    """シートの上の行の値（2 次元）→ 見出しの語の集合（純 Python）。前後の空白を落とした語と、ならした語（head_key）の両方。"""
    words = set()
    for row in list(values or [])[:_SHEET_WORD_ROWS]:
        for v in (row or []):
            if isinstance(v, str) and v.strip():
                words.add(v.strip())
                words.add(head_key(v))
    return words


def _header_row_index(kinds):
    """見出しの行（文字が 2 つ以上並ぶ最初の行）の添字（純 Python）。無ければ 0。"""
    for i, k in enumerate(kinds or []):
        if str(k).count('s') >= 2:
            return i
    return 0


def _rate_like(col_texts):
    """数の列の値が全部 ±1.5 以内で小数を含む＝率の列（vbam_edit.summable と同じ物差し・純 Python）。"""
    nums = []
    for t in col_texts:
        try:
            nums.append(float(str(t).replace(',', '').replace('%', '')))
        except (TypeError, ValueError):
            continue
    return bool(nums) and all(-1.5 <= v <= 1.5 for v in nums) and any(not float(v).is_integer() for v in nums)


def shape_signals(values, kinds, has_table=False, other_heads=(), merged_head=False, has_pivot=False, has_chart=False):
    """表の形の語の集合（純 Python）。values＝文字の 2 次元・kinds＝行ごとの型の文字列（n 数・d 日付・s 文字・- 空）。
    other_heads＝ほかのシートの見出しの語の集合の並び。入口は登録簿の \' 形: がこの集合に全部あるときだけ候補にする。
    2026-09-18 朝: 25 本を全部「見出し: なし」にしたら語のかすり当たりで別マクロも撃たれた＝形で先に候補を絞る。"""
    out = set()
    kinds = [str(k) for k in (kinds or [])]
    values = [list(r or []) for r in (values or [])]
    h = _header_row_index(kinds)
    body = [k for k in kinds[h + 1:] if k.strip('-')]
    ncols = max((len(k) for k in kinds), default=0)
    for c in range(ncols):
        col = [k[c] for k in body if c < len(k) and k[c] != '-']
        if not col:
            continue
        n = sum(1 for x in col if x == 'n')
        d = sum(1 for x in col if x == 'd')
        if n and n >= len(col) * 0.8:
            out.add('数の列')
            texts = [r[c] for r in values[h + 1:] if c < len(r) and r[c] not in (None, '')]
            if _rate_like(texts):
                out.add('率の列')
        if d and d >= len(col) * 0.8:
            out.add('日付の列')
    if has_table:
        out.add('テーブル')
    if merged_head:
        out.add('結合の見出し')
    if has_pivot:
        out.add('ピボット')
    if has_chart:
        out.add('グラフ')
    heads = set()
    if h < len(values):
        heads = {head_key(v) for v in values[h] if isinstance(v, str) and v.strip()}
    for oh in other_heads or ():
        oh = {head_key(v) for v in (oh or ()) if isinstance(v, str) and v.strip()}
        if not oh:
            continue
        out.add('別のシート')
        if heads and oh == heads:
            out.add('同じ見出しのシート')
        if heads & oh:
            out.add('共通の見出しのシート')
    return out


def _shape_of_sheet(ws, wb, rows=60):
    """アクティブなシートの形（COM）→ shape_signals の集合。読めなければ空（＝形の要る仕事は候補にしない・安全側）。"""
    from vbam_hands import _rows_of
    try:
        from vbam_rehearse import _kind_of, _texts
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        nr = min(int(ur.Rows.Count), rows)
        nc = int(ur.Columns.Count)
        raw = _rows_of(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1)).Value) or []
        values = _texts(raw)
        kinds = [''.join(_kind_of(v) for v in row) for row in raw]
        has_table = False
        try:
            has_table = int(ws.ListObjects.Count) > 0
        except Exception:
            pass
        merged = False
        try:
            # 見出しの行〜+2 行までは必ず見る（見出しが下にある表でも結合の見出しを取りこぼさない・2026-09-18）
            last = min(r0 + max(4, _header_row_index(kinds) + 2), r0 + nr - 1)
            top = ws.Range(ws.Cells(r0, c0), ws.Cells(last, c0 + nc - 1))
            merged = top.MergeCells is True or top.MergeCells is None      # None＝混在（結合がある）
        except Exception:
            pass
        others = []
        try:
            for sh in list(wb.Worksheets)[:12]:
                if str(sh.Name) == str(ws.Name):
                    continue
                ur2 = sh.UsedRange
                nr2 = min(int(ur2.Rows.Count), _SHEET_WORD_ROWS)
                rows2 = _rows_of(sh.Range(sh.Cells(int(ur2.Row), int(ur2.Column)),
                                          sh.Cells(int(ur2.Row) + nr2 - 1, int(ur2.Column) + int(ur2.Columns.Count) - 1)).Value) or []
                k2 = [''.join(_kind_of(v) for v in row) for row in rows2]
                hi = _header_row_index(k2)
                if hi < len(rows2):
                    others.append([v for v in rows2[hi] if isinstance(v, str)])
        except Exception:
            pass
        has_pivot = has_chart = False
        try:
            has_pivot = int(ws.PivotTables().Count) > 0
        except Exception:
            pass
        try:
            has_chart = int(ws.ChartObjects().Count) > 0
        except Exception:
            pass
        return shape_signals(values, kinds, has_table, others, merged, has_pivot, has_chart)
    except Exception:
        return set()


def dedupe_registry(entries, prefer=()):
    """同じ名前の Sub は 1 本だけにする（純 Python）。prefer（開いているブックの名前）の持ち主を先に採る。
    2026-09-17: 秀コンボ.xlsm に登録 → 更新登録で .xlam にも入ると、登録簿に同じ Sub が 2 本並び、
    先撃ちが 2 回撃つ（合計行が二重になる）。開いている .xlsm のほうが新しい（更新登録の前）ので先に採る。"""
    prefer = set(prefer or ())
    ordered = sorted(enumerate(entries or ()), key=lambda ie: (ie[1].get('owner') not in prefer, ie[0]))
    seen, keep = set(), set()
    for i, e in ordered:
        if e['name'] in seen:
            continue
        seen.add(e['name'])
        keep.add(i)
    return [e for i, e in enumerate(entries or ()) if i in keep]


def registry_of(xl):
    """開いているブック／アドインの「表の整理」から登録簿を読む（COM。読めなければ []）。"""
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return []
    out = []
    for p in projects:
        try:
            text = shelf_text(p)            # 「表の整理_〜」に分かれていても全部（2026-09-23）
            if not text:
                continue
            owner = _owner_name(xl, p)
            if not owner:
                continue
            out += registry_from_text(text, owner)
        except Exception:
            continue
    books = []
    try:
        books = [str(w.Name) for w in xl.Workbooks]
    except Exception:
        pass
    return dedupe_registry(out, books)


def _top_words(ws):
    """シートの上 _SHEET_WORD_ROWS 行の見出しの語（COM）。
    2026-09-17: UsedRange.Resize(n) は win32com の動的呼び出しで通らず、例外 → 安全側で「撃たない」に倒れていた。"""
    from vbam_hands import _rows_of
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    nr = min(int(ur.Rows.Count), _SHEET_WORD_ROWS)
    nc = int(ur.Columns.Count)
    top = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1))
    return sheet_words(_rows_of(top.Value))


def head_in(h, words):
    """見出しの語 h（「課名|行ラベル」のように | で並べた言い換えもよい）が、シートの見出しの語 words にあるか（純 Python）。
    2026-09-18: ピボットを値で貼り付けた集計表は「課名」でなく「行ラベル」＝全部の表に共通の語が無く、見出しの行を書けなかった。"""
    return any(a and (a in words or head_key(a) in words) for a in str(h).split('|'))


def pick_columns(request, cells, want=None):
    """依頼文と見出しのセル（[(列番号, 見出しの文字)]）→ 依頼に出てくる順に並べた列番号（純 Python）。
    want＝(最小, 最大) を渡すと、その数でないときは None（＝入口は撃たずに AI へ回す）。
    2026-09-18: 列の選び方が要る仕事（ピボットの行・列・値など）は、人が列を選んでから撃つ決まりにした。
    入口では依頼の語と見出しが一致する列を道具が選ぶ。"""
    req = str(request or '')
    key_req = head_key(req)
    all_cells = [(int(c), str(t or '').strip()) for c, t in (cells or ()) if str(t or '').strip()]

    def _match(cs):
        hits = []
        for col, t in cs:
            at = req.find(t)
            if at < 0:
                k = head_key(t)
                at = key_req.find(k) if k else -1
            if at >= 0:
                hits.append((at, col))
        # 同じ見出しの列が 2 つ以上あるとき（前の結果の表が右にある・2026-09-18）は、依頼の語 1 つにつき左の列を採る
        by_at = {}
        for at, col in hits:
            by_at.setdefault(at, []).append(col)
        out = []
        for at in sorted(by_at):
            for col in sorted(by_at[at]):
                if col not in out:
                    out.append(col)
                    break
        return out

    cols = _match(all_cells)
    if want and len(cols) > int(want[1]):
        # 当たりすぎ＝右に前の結果の表・別の表がある。空の列で切れた左の塊だけで選び直す（2026-09-18 朝:
        # 依頼文の「4月」「3月」「合計」が、右に残った前の集計表の見出しに当たり、月次集計が AI に回っていた）
        run, prev = [], None
        for col, t in sorted(all_cells):
            if prev is not None and col > prev + 1:
                break
            run.append((col, t))
            prev = col
        narrowed = _match(run)
        if int(want[0]) <= len(narrowed) <= int(want[1]):
            cols = narrowed
    if want and not (int(want[0]) <= len(cols) <= int(want[1])):
        return None
    return cols


def _header_cells(ws):
    """シートの見出しの行のセル（COM）→ [(列番号, 見出しの文字)]。読めなければ []。"""
    from vbam_hands import _rows_of
    try:
        anchor = _table_anchor(ws)
        if not anchor:
            return []
        cell = ws.Range(anchor)
        r, c0 = int(cell.Row), int(cell.Column)
        ur = ws.UsedRange
        c2 = int(ur.Column) + int(ur.Columns.Count) - 1
        row = (_rows_of(ws.Range(ws.Cells(r, c0), ws.Cells(r, c2)).Value) or [[]])[0]
        return [(c0 + j, str(v).strip()) for j, v in enumerate(row) if isinstance(v, str) and str(v).strip()]
    except Exception:
        return []


def current_columns(ws):
    """いま選んでいるセルの列（COM）→ [列番号]（選んだ順・重複なし）。そのシートを選んでいなければ []。
    2026-09-18: 列の選び方が要る仕事は「人が選んだ列」を使う。人が選んでいればそれを使い、
    選んでいなければ入口が依頼の語と見出しから選ぶ。"""
    try:
        sel = ws.Application.Selection
        if str(sel.Parent.Name) != str(ws.Name):
            return []
        cols = []
        for a in sel.Areas:
            c0, n = int(a.Column), int(a.Columns.Count)
            if n > 3:
                continue                                 # 行全体・シート全体の選択は「列を選んだ」と見ない
            for c in range(c0, c0 + n):
                if c not in cols:
                    cols.append(c)
        return cols
    except Exception:
        return []


_REG_BY_REQUEST_RE = re.compile(r"^\s*'\s*列は頼みから\s*[:：]")
_SELECTION_WORDS_RE = re.compile(r"選んだ|選択した|選択範囲|この範囲|このセル|この列|この行")


def request_column_macros(text):
    """モジュールの本文 → 頭に ' 列は頼みから: の行がある Sub の名前の集合（純 Python・2026-10-09）。
    その棚は列を選ばずに撃つ（頼みの文をコンボ道具に渡せば、マクロが見出しから行・列・値を決める）。"""
    out, name, n = set(), None, 0
    for ln in str(text or '').splitlines():
        m = _REG_SUB_RE.match(ln)
        if m:
            name, n = m.group(1), 0
            continue
        if name is None:
            continue
        n += 1
        if n > 12 or not ln.strip().startswith("'"):
            name = None                                  # 頭のコメントの並びが切れたら、その Sub は見終わり
        elif _REG_BY_REQUEST_RE.match(ln):
            out.add(name)
            name = None
    return out


def fire_with_request(xl, owner, ws, name, request):
    """頼みの文をコンボ道具に渡して棚のマクロを撃つ（COM・2026-10-09）。→ 断り（撃てなかった理由と候補の見出し。撃てたら ''）。
    前の撃ちで選んだ列が残っているとマクロがそれを使うので、表の左上の 1 セルを選んでから撃つ。撃った後は頼みの文を消す
    （残すと、後で人が棚から撃ったときに古い頼みで動く）。"""
    from vbam_vba import run_book_macro
    q = str(owner).replace("'", "''")
    xl.Run(f"'{q}'!コンボ道具.頼みの文を入れる", str(request or ''))
    try:
        # 「選んだ範囲を斜体にして」「この列を…」は人の選んだ所が相手なので選び直さない（2026-10-29 072: B3:C5 を選んで撃つと 1 セルだけ変わった）
        if not _SELECTION_WORDS_RE.search(str(request or '')):
            try:
                anchor = _table_anchor(ws)
                ws.Range(anchor or 'A1').Select()
            except Exception:
                pass
        shelf_refusal(xl, owner, clear=True)       # 前の棚の断りが残っていると、撃てた棚まで断ったと読む（2026-10-14）
        note_shot(xl, owner, None)
        with contextlib.suppress(Exception):
            xl.Run(f"'{q}'!コンボ道具.姿を控える", ws.Parent, ws, str(name))
        run_book_macro(xl, owner, name)
        why = shelf_refusal(xl, owner)
        if not why:
            violation = ''
            with contextlib.suppress(Exception):
                violation = str(xl.Run(f"'{q}'!コンボ道具.約束を破ったか", str(name)) or '')
            if violation:
                with contextlib.suppress(Exception):
                    xl.Run(f"'{q}'!表の整理.原本から元に戻す", ws, True)
                return f"約束違反で戻しました: {name}（{violation}）"
            note_shot(xl, owner, name, request)
        return why
    finally:
        with contextlib.suppress(Exception):
            xl.Run(f"'{q}'!コンボ道具.頼みの文を入れる", '')
        with contextlib.suppress(Exception):
            xl.StatusBar = False


def note_shot(xl, owner, name, request=''):
    """撃った棚の控え（コンボ道具.棚を撃つ前／撃った棚を控える・棚番 段 4・2026-10-15）。name=None は撃つ前（シートと選んだ所を覚える）。
    控えは %APPDATA%\秀コンボ\撃った棚.tsv に残り、人が「今の手順をレシピに保存」で印を付けてレシピにする。古いアドインに口が無ければ何もしない。"""
    q = str(owner).replace("'", "''")
    with contextlib.suppress(Exception):
        if name is None:
            xl.Run(f"'{q}'!コンボ道具.棚を撃つ前")
        else:
            xl.Run(f"'{q}'!コンボ道具.撃った棚を控える", str(name), str(request or ''))


_DOUBT_FILE = os.path.join(tempfile.gettempdir(), 'vbam_last_doubt.json')


def remember_doubt(request, candidates, book_name):
    """最後に迷った頼み（頼みの文・候補・ブック名）を一時ファイルに覚える（2026-10-22 指示書040）。"""
    try:
        import json
        with open(_DOUBT_FILE, 'w', encoding='utf-8') as f:
            json.dump({'request': str(request or ''), 'candidates': str(candidates or ''), 'book': str(book_name or '')}, f, ensure_ascii=False)
    except Exception:
        pass


def get_last_doubt():
    """最後に迷った頼みを取得する。無ければ None。"""
    try:
        import json
        if os.path.isfile(_DOUBT_FILE):
            with open(_DOUBT_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return None


def clear_last_doubt():
    """最後に迷った頼みを消す。"""
    with contextlib.suppress(Exception):
        if os.path.isfile(_DOUBT_FILE):
            os.remove(_DOUBT_FILE)


def save_ai_choice(xl, owner, chosen_name, book_name=None):
    """最後に迷った頼みがあれば、VBA の AIの選んだ答えを残す を呼んで記録する（2026-10-22 指示書040）。"""
    d = get_last_doubt()
    if not d:
        return False
    q = str(owner).replace("'", "''")
    b = str(book_name or d.get('book') or '')
    req = str(d.get('request') or '')
    cands = str(d.get('candidates') or '')
    ans = str(chosen_name or '')
    try:
        xl.Run(f"'{q}'!コンボ道具.AIの選んだ答えを残す", b, req, cands, ans)
        clear_last_doubt()
        return True
    except Exception:
        return False


def check_doubt(request, entries=(), ratio=1.2):
    """依頼文から迷い判定を行う（2026-10-22 指示書040）。
    迷った場合: (True, 候補文字列, 案内文)
    迷わない場合: (False, '', '')
    """
    req = request_normalize(request, entries)
    if not req:
        return False, '', ''
    low = req.lower()
    cf_block = '条件付き書式' in req and not any(w in req for w in _SHELF_CF_LOOK)
    points, details, order = {}, {}, []
    for e in entries or ():
        name = e.get('name')
        if not name:
            continue
        if cf_block and not e.get('cf_ok'):
            continue                      # 条件付き書式を付ける頼みで当たるのは、その印の棚だけ（shelf_pick と同じ）
        c_pt, c_ban = combo_score(e.get('combo'), req)
        a_pt = _ask_points(e.get('ask'), req)
        if not c_ban and (c_pt + a_pt) > 0:
            if name not in points:
                order.append(name)
            points[name] = points.get(name, 0) + c_pt + a_pt
            words = []
            if c_pt > 0:
                words.append(f"組({c_pt}点)")
            if a_pt > 0:
                for w in str(e.get('ask') or '').split('|'):
                    w = w.strip()
                    if w and w in req:
                        words.append(w)
            details[name] = '・'.join(words) if words else ''

    if not points:
        return False, '', ''
    sorted_pts = sorted([(nm, points[nm]) for nm in order], key=lambda x: x[1], reverse=True)
    top_name, top_pt = sorted_pts[0]
    sec_name, sec_pt = sorted_pts[1] if len(sorted_pts) > 1 else ('', 0)

    # 例外: マクロの名前そのものを言った頼み
    if '【手順書：' in req or (len(top_name) >= 4 and top_name.lower() in low):
        return False, '', ''

    is_doubt = False
    if sec_pt > 0 and top_pt < sec_pt * ratio:
        is_doubt = True

    if not is_doubt:
        return False, '', ''

    cands = []
    for i, (nm, pt) in enumerate(sorted_pts[:3], start=1):
        dt = details.get(nm, '')
        cands.append(f"{i}. {nm}（{pt}点・{dt}）" if dt else f"{i}. {nm}（{pt}点）")
    cands_str = " ".join(cands)
    guide = f"迷ったので撃っていません: 候補 {cands_str}（上位 {len(cands)} 本）\n→ 合う棚があれば shelf-run その名前／どれも合わなければ手で直し、shelf --chose --none\n" + HEARD_HINT
    return True, cands_str, guide


def shelf_refusal(xl, owner, clear=False):
    """棚の断りの口（コンボ道具.棚の断りを入れる／返す・2026-10-09）を読む（clear=True なら空にする）。棚が撃たなかった理由。無ければ ''。"""
    q = str(owner).replace("'", "''")
    try:
        if clear:
            xl.Run(f"'{q}'!コンボ道具.棚の断りを入れる", '')
            return ''
        return str(xl.Run(f"'{q}'!コンボ道具.棚の断りを返す") or '')
    except Exception:
        return ''


def select_columns(ws, cols):
    """見出しのセルを選ぶ（COM）。選んだ列がマクロの入り口（Selection）になる。"""
    if not cols:
        return False
    try:
        anchor = _table_anchor(ws)
        r = int(ws.Range(anchor).Row) if anchor else 1
        ws.Range(",".join(str(ws.Cells(r, c).Address) for c in cols)).Select()
        return True
    except Exception:
        return False


def plan_full(request, registry=(), words=None, shape=None, pick=None, entries=None):
    """依頼文＋登録簿（＋シートの見出しの語）→ {'fire','dedupe','other','extras','left','skipped'}（純 Python）。
    2026-10-22 指示書039: 棚を選ぶのは Excelコンボと同じ判定（VBA の先撃ち判定／shelf_plan／shelf_pick）1つにする。
    見出し（words）と表の形（shape）は、選ばれた棚の「撃たない理由（skipped）」として残す。
    """
    req = str(request or '')
    reg_by_name = {e['name']: e for e in registry or ()}

    # 判定対象のエントリ（entries が無ければ registry を使う）
    ent = entries if entries is not None else registry

    # 1. Excelコンボと同じ判定で棚を選ぶ（複数なら shelf_plan、単一なら pick/shelf_pick）
    planned = shelf_plan(req, ent, pick=pick)
    if planned:
        chosen_names = planned
    else:
        one = pick(req) if pick else shelf_pick(req, ent)
        chosen_names = [one] if one else []

    # 指示書: tidy（表を整える）・dedupe（重複行）の決め方は今のまま
    tidy = bool(_ASK_TIDY_RE.search(req))
    dedupe = bool(tidy and _DUP_WORD_RE.search(req) and va._has_approval(req) and not _NO_DELETE_RE.search(req))

    extras, skipped = [], []
    for nm in chosen_names:
        if nm in (_MACRO_TIDY, _MACRO_DEDUPE):
            continue
        e = reg_by_name.get(nm)
        if not e:
            e = {'name': nm, 'ask': '', 'handles': [], 'headers': [], 'shape': []}
        # 要る見出しがシートに無い表には撃たない
        if words is not None and e.get('headers') and not all(head_in(h, words) for h in e['headers']):
            skipped.append(dict(e, strength=1))
            continue
        # 表の形（' 形:）がそろわない表には撃たない
        if shape is not None and e.get('shape') and not set(e['shape']) <= set(shape):
            skipped.append(dict(e, strength=1, why_shape=[w for w in e['shape'] if w not in shape]))
            continue
        extras.append(dict(e, strength=1))

    shadowed = []
    others = sorted({m.group(0) for m in _ASK_OTHER_RE.finditer(req)})
    covered = [h for e in extras for h in (e.get('handles') or [])]
    clauses = [c for c in re.split(r'[。、,，\n]', req) if c]
    asks = [e.get('ask') or '' for e in extras]

    def _covered(w):
        if any(h in w or w in h for h in covered):
            return True
        if _STRONG_OTHER_RE.fullmatch(w):
            return False
        return any(w in c and (any(h in c for h in covered) or any(_ask_strength(a, c) for a in asks if a)) for c in clauses)

    left = [w for w in others if not _covered(w)]
    return {'fire': bool(tidy or dedupe or extras), 'tidy': tidy, 'dedupe': dedupe, 'extras': extras,
            'left': left, 'other': bool(left), 'skipped': skipped, 'shadowed': shadowed, 'chosen': chosen_names}


def _ask_strength(ask, req):
    """依頼の語（語を | で並べたもの）が依頼文に当たった一番長い語の文字数（純 Python）。当たらなければ 0。
    正規表現の記号の入った古い登録（.* など）は、当たった部分でなく語の文字（記号を除く）の長さで数える。"""
    best = 0
    for alt in str(ask or '').split('|'):
        if not alt:
            continue
        try:
            if re.search(alt, req):
                best = max(best, len(re.sub(r'[.*+?()\[\]{}^$\\]', '', alt)))
        except re.error:
            continue
    return best


# 棚に当てない依頼（VBA の コンボ道具.先撃ち候補 と同じ止め・2026-09-23）:
#   コードや案を頼む依頼＝AI の役目／使い方・意味・考え方の質問＝作業ではない
_SHELF_NOT_WORK = ("マクロを書", "マクロを作", "書き直", "案を", "案は", "の案")
_SHELF_QUESTION = ("とは何", "とは？", "とは?", "って何", "の違いは", "違いを教え", "関数を教え", "関数は",
                   "どうやって", "どうしたら", "どうすれば", "やり方", "方法を", "方法は", "使い方を教",
                   "使い方は", "使い方が分", "使い方がわ", "相談", "コツ", "作り方", "分析して", "傾向",
                   "考え方", "何が言え", "何が分か", "した方がいい", "しない方がいい", "使わない方がいい",
                   "ばいいか", "ばよいか", "には何を",   # 「整えるには何を直せばいいか教えて」は質問（2026-10-14）
                   "判断して", "見極め", "提案して", "どうすべき", "べきか")   # 判断を頼む言い方も AI の役目（2026-10-22 指示書046）
_SHELF_QUESTION_EXCEPT = ("ABC分析", "ＡＢＣ分析", "パレート分析", "abc分析")
# 条件付き書式を付ける依頼は AI の役目（棚にあるのは一覧にするマクロだけ）。調べる・一覧の語があるときだけ棚に当てる
# （2026-10-14: seiri の「条件付き書式で赤くして」が 入力規則などの仕掛けを一覧にする に当たっていた。VBA の 先撃ち候補 と同じ決まり）
_SHELF_CF_LOOK = ("調べ", "一覧", "どこ", "探", "確認", "チェック", "点検", "洗い出", "リスト")
# エラーの原因を尋ねる依頼（「B3 が #N/A になる原因を教えて」）は、一覧を作る棚でなく AI が構造を見て答える（2026-10-02・
# オフィス田中の使い方＝エラーの相談はブックの構造を添えて AI に聞く）。エラーの印と理由を尋ねる語が両方あるときだけ止める
# （「ブックが重い原因を調べて」はエラーの印が無いので今までどおり棚に当たる）
_SHELF_ERROR_MARK = ("エラー", "#N/A", "#REF", "#NAME", "#VALUE", "#DIV", "#SPILL", "#NUM", "#NULL", "#CALC")
_SHELF_WHY = ("原因", "なぜ", "どうして", "なんで", "何で", "理由")
_SHELF_QUESTION_LOWER = tuple(w.lower() for w in _SHELF_QUESTION)
_SHELF_QUESTION_EXCEPT_LOWER = tuple(w.lower() for w in _SHELF_QUESTION_EXCEPT)
_SHELF_ERROR_MARK_LOWER = tuple(w.lower() for w in _SHELF_ERROR_MARK)
_SHELF_NOT_WORK_RE = re.compile('|'.join(re.escape(w) for w in _SHELF_NOT_WORK))
_SHELF_QUESTION_RE = re.compile('|'.join(re.escape(w.lower()) for w in _SHELF_QUESTION))
_SHELF_QUESTION_EXCEPT_RE = re.compile('|'.join(re.escape(w.lower()) for w in _SHELF_QUESTION_EXCEPT))
_SHELF_ERROR_MARK_RE = re.compile('|'.join(re.escape(w.lower()) for w in _SHELF_ERROR_MARK))
_SHELF_WHY_RE = re.compile('|'.join(re.escape(w) for w in _SHELF_WHY))
_SHELF_FIX_RE = re.compile(r'直して|直す|修正')


def _ask_points(ask, req):
    """' 依頼の語: A|B|C のうち依頼文に当たった語の長さの**合計**（VBA と同じ点・純 Python）。

    _ask_strength（いちばん長い 1 語）とは別物。VBA の先撃ちは当たった語を全部足す。
    """
    low = str(req or '').lower()
    return sum(len(w) for w in (a.strip() for a in str(ask or '').split('|')) if w and w.lower() in low)


def combo_score(combo, req):
    """' 依頼の組: A,B+C,D-E → (点, 除外か)（純 Python・2026-09-23）。

    + で区切った組の**どれにも** 1 語ずつ当たったときだけ点が付き、各組で当たったいちばん長い語の長さを足す
    （言い方が変わっても当たる。例「重複,ダブ + 消,削除」）。- で始まる組の語が 1 つでもあれば除外
    （似たマクロとの取り合いを止める）。組が 1 つも当たらなければ 0 点。
    """
    low = str(req or '').lower()
    total, ok, banned = 0, True, False
    for group in str(combo or '').split('+'):
        group = group.strip()
        if not group:
            continue
        if group.startswith('-'):
            if any(w and w.lower() in low for w in (x.strip() for x in group[1:].split(','))):
                banned = True
            continue
        best = max((len(w) for w in (x.strip() for x in group.split(',')) if w and w.lower() in low), default=0)
        if best:
            total += best
        else:
            ok = False
    return (total if ok else 0), banned


# ---- 棚の辞書（言い方の台帳・2026-10-23 指示書047） ----
# つなぎ語と聞き違い（票→表 など）は、コードに決め打ちせず、標準モジュール「棚の辞書」（コメント行だけ）に 1 行 1 件で書く。
# VBA（コンボ道具.頼みの下ごしらえ）と同じ行を読む。使う人ごとの学び（違う頼みの文で 2 回以上残った対）は
# %APPDATA%\秀コンボ\聞き違いの学び.tsv（VBA の 聞き違いを残す だけが書く）。読む順: 台帳 → 学び → 色（今のまま・048 で概念の辞書へ）。
LEDGER_FILE = '棚の辞書.bas'
# 台帳を読む元: ①いま開いているブックの「棚の辞書」（shelf_text を通ったとき vbam_core.DICT_LIVE に入る）→ ②同梱の写し（python_scripts\棚の辞書.bas・
# 公開の写しでは macros\棚の辞書.bas）→ ③下の既定（台帳と同じ中身。写しと食い違えば試験が落ちる）
_LEDGER_DEFAULT = """' つなぎ語: ええ,えー,えーと,えっと,あの,あのー,じゃあ,では,それでは,はい
' 聞き違い: 票=>表 ; 前で止める字=伝投得帳開
' 聞き違い: 産業ごと=>3桁ごと ; 後ろに要る語=気切,区切
' 聞き違い: 中央ドライ=>中央揃え ; 出どころ=聞き違いの学び ×2
' 聞き違い: 選択制度=>選択セル ; 出どころ=聞き違いの学び ×4
' 聞き違い: 見出し尿=>見出し行 ; 出どころ=聞き違いの学び ×2
' 聞き違い: 気切り=>区切り ; 出どころ=聞き違いの学び ×4
' 聞き違い: 洗濯=>選択 ; 出どころ=聞き違いの学び ×2
' 聞き違い: 不倫仮名=>ふりがな ; 出どころ=外れた依頼 2026-10-24
' 聞き違い: 不倫がな=>ふりがな ; 出どころ=外れた依頼 2026-10-24
' 聞き違い: フリマナ=>ふりがな ; 出どころ=外れた依頼 2026-10-24
"""
# seiri が棚に当てられず AI に回すときの案内の 1 行（頼みの聞き違いを学びに残してもらう）
HEARD_HINT = '→ 頼みに聞き違いがあれば（例: 産業ごと→3桁ごと）、直したうえで shelf --heard 誤 正 --q "頼みの文" で残す（違う頼みの文で 2 回以上残ると自動で直る）'
LEARNED_MIN_TEXTS = 2           # 学びが効くのに要る、違う頼みの文の数
LEARNED_MIN_LEN = 2             # 学びの誤の最小の長さ
HEARD_WINDOW = 6                # 後ろに要る語を探す字数
LEARNED_PATH_OVERRIDE = {'path': None}  # 試験が学びのファイルを差し替える（'path' に置き場を入れる）
_LEDGER_CACHE = {}


def parse_ledger(text):
    """台帳の本文 → {'fillers': [語], 'pairs': [{'wrong','right','stop','need'}]}（純 Python。VBA の コンボ道具.辞書を読む と同じ読み方）。
    読むのは「' つなぎ語:」「' 聞き違い:」の行だけ（全角の：；も可）。聞き違いは「誤=>正 ; 前で止める字=… ; 後ろに要る語=カンマ区切り ; 出どころ=…」。"""
    fillers, pairs = [], []
    for ln in str(text or '').replace('\r', '').split('\n'):
        s = ln.strip()
        if not s.startswith("'"):
            continue
        body = s[1:].strip().replace('：', ':').replace('；', ';')
        key, sep, rest = body.partition(':')
        if not sep:
            continue
        if body.startswith('つなぎ語'):
            fillers += [w.strip() for w in rest.split(',') if w.strip()]
        elif body.startswith('聞き違い'):
            fields = rest.split(';')
            head = fields[0]
            if '=>' not in head:
                continue
            wrong, right = (x.strip() for x in head.split('=>', 1))
            stop = need = ''
            for f in fields[1:]:
                k, eq, v = f.partition('=')
                if not eq:
                    continue
                k, v = k.strip(), v.strip()
                if k == '前で止める字':
                    stop = v
                elif k == '後ろに要る語':
                    need = ','.join(w.strip() for w in v.split(',') if w.strip())
            if wrong and right:
                pairs.append({'wrong': wrong, 'right': right, 'stop': stop, 'need': need})
    return {'fillers': fillers, 'pairs': pairs}


def _ledger_file_candidates():
    here = os.path.dirname(os.path.abspath(__file__))
    return [os.path.join(here, LEDGER_FILE), os.path.join(here, '..', 'macros', LEDGER_FILE)]


def ledger_text():
    """いま使う台帳の本文と出どころ → (本文, 'ブック'|'写し'|'既定')。"""
    try:
        import vbam_core
        live = vbam_core.DICT_LIVE.get('text')
    except Exception:
        live = None
    if live is not None:
        return live, 'ブック'
    for p in _ledger_file_candidates():
        try:
            st = os.stat(p)
        except OSError:
            continue
        key = (p, st.st_mtime_ns, st.st_size)
        if _LEDGER_CACHE.get('file_key') != key:
            raw = open(p, 'rb').read()
            try:
                txt = raw.decode('cp932')
            except UnicodeDecodeError:
                txt = raw.decode('utf-8', errors='replace')
            _LEDGER_CACHE['file_key'], _LEDGER_CACHE['file_text'] = key, txt
        return _LEDGER_CACHE['file_text'], '写し'
    return _LEDGER_DEFAULT, '既定'


def current_ledger():
    """いま使う台帳（parse_ledger の結果。本文が同じなら読み直さない）。"""
    text, _src = ledger_text()
    if _LEDGER_CACHE.get('parsed_text') != text:
        led = parse_ledger(text)
        _LEDGER_CACHE['parsed_text'] = text
        _LEDGER_CACHE['parsed'] = led
        if led['fillers']:
            _LEDGER_CACHE['filler_re'] = re.compile(
                r'^(?:(?:' + '|'.join(re.escape(w) for w in led['fillers']) + r')[、,\s]+)+')
        else:
            _LEDGER_CACHE['filler_re'] = None
    return _LEDGER_CACHE['parsed']


def learned_path():
    """学びのファイル（VBA だけが書く）。"""
    return LEARNED_PATH_OVERRIDE.get('path') or os.path.join(os.environ.get('APPDATA', ''), '秀コンボ', '聞き違いの学び.tsv')


def read_learned(path=None):
    """学びのファイル → {(誤, 正): [違う頼みの文の並び]}（挿入順。cp932・タブ区切り・見出し行は読まない）。無ければ {}。"""
    p = path or learned_path()
    try:
        st = os.stat(p)
    except OSError:
        return {}
    key = (p, st.st_mtime_ns, st.st_size)
    if _LEDGER_CACHE.get('learned_key') == key:
        return _LEDGER_CACHE['learned']
    import io
    out = {}
    for ln in io.open(p, encoding='cp932', errors='replace'):
        f = ln.rstrip('\r\n').split('\t')
        if len(f) < 4 or f[0] == '日時':
            continue
        wrong, right, text = f[1].strip(), f[2].strip(), ' '.join(f[3].replace('　', ' ').split())
        if not (wrong and right and text):
            continue
        texts = out.setdefault((wrong, right), [])
        if text not in texts:
            texts.append(text)
    _LEDGER_CACHE['learned_key'], _LEDGER_CACHE['learned'] = key, out
    return out


def heard_steals_shelf(wrong, entries):
    """学びの誤が、棚の名前・依頼の語・依頼の組のどれかに入っているか（横取りの検査。VBA の 誤が棚の語か と同じ）。"""
    for e in entries or ():
        for fld in (e.get('name'), e.get('ask'), e.get('combo')):
            if fld and wrong in str(fld):
                return True
    return False


def learned_pairs(entries=None):
    """効いている学びの対 [{'wrong','right','stop','need'}]。違う頼みの文で LEARNED_MIN_TEXTS 回以上・誤が 2 字以上・棚の語に入っていない物だけ。"""
    out = []
    for (wrong, right), texts in read_learned().items():
        if len(texts) >= LEARNED_MIN_TEXTS and len(wrong) >= LEARNED_MIN_LEN and wrong != right \
                and not heard_steals_shelf(wrong, entries):
            out.append({'wrong': wrong, 'right': right, 'stop': '', 'need': ''})
    return out


def _learned_row_count(path=None):
    """学びのファイルの行数（見出しを除く）。無ければ 0。"""
    import io
    p = path or learned_path()
    if not os.path.isfile(p):
        return 0
    return sum(1 for ln in io.open(p, encoding='cp932', errors='replace')
               if ln.strip() and not ln.startswith('日時\t'))


def record_heard(xl, owner, wrong, right, text):
    """聞き違いを学びに残す（VBA の コンボ道具.聞き違いを残す を呼ぶ＝ファイルは VBA だけが書く）→ (残せたか, 説明)。"""
    wrong, right, text = str(wrong or '').strip(), str(right or '').strip(), ' '.join(str(text or '').split())
    if not wrong or not right or wrong == right:
        return False, "誤と正を別々に渡してください（例: shelf --heard 産業ごと 3桁ごと --q \"頼みの文\"）"
    if not text:
        return False, "頼みの文が要ります（--q \"そのときの頼みの文\"）。違う頼みの文で 2 回以上残ると効きます"
    q = str(owner).replace("'", "''")
    before = _learned_row_count()
    try:
        xl.Run(f"'{q}'!コンボ道具.聞き違いを残す", wrong, right, text)
    except Exception as e:
        return False, f"残せませんでした（古い秀コンボで 聞き違いを残す が無い？）: {e}"
    if _learned_row_count() <= before:
        return False, "残していません（お試し・テスト・隔離のブックは書かない守りです）"
    texts = read_learned().get((wrong, right), [])
    msg = f"聞き違いを残しました: 「{wrong}」→「{right}」（この対は違う頼みの文で {len(texts)} 件目）"
    if len(texts) < LEARNED_MIN_TEXTS:
        msg += f"。違う頼みの文でもう {LEARNED_MIN_TEXTS - len(texts)} 回残ると下ごしらえが自動で直します"
    else:
        msg += "。下ごしらえが自動で直します（棚の語に入っている誤・1 字の誤は除く）"
    return True, msg


def learned_report(entries=None):
    """聞き違いの学びの集計 → 行の並び（shelf --misses の終わりに出す）。対・回数・効いているか・台帳に載せる候補（違う頼みの文で 2 回以上）。"""
    import io
    p = learned_path()
    counts = {}
    if os.path.isfile(p):
        for ln in io.open(p, encoding='cp932', errors='replace'):
            f = ln.rstrip('\r\n').split('\t')
            if len(f) >= 4 and f[0] != '日時' and f[1].strip() and f[2].strip():
                k = (f[1].strip(), f[2].strip())
                counts[k] = counts.get(k, 0) + 1
    learned = read_learned()
    ledger_pairs = {(x['wrong'], x['right']) for x in current_ledger()['pairs']}
    lines = []
    for (wrong, right), rows in counts.items():
        texts = len(learned.get((wrong, right), []))
        if (wrong, right) in ledger_pairs:
            state = "台帳に載っている"
        elif texts < LEARNED_MIN_TEXTS:
            state = f"まだ効かない（違う頼みの文があと {LEARNED_MIN_TEXTS - texts} 回）"
        elif len(wrong) < LEARNED_MIN_LEN:
            state = f"効かない（誤が {LEARNED_MIN_LEN} 字未満）"
        elif heard_steals_shelf(wrong, entries):
            state = "効かない（誤が棚の名前・依頼の語・組の語に入っている）"
        else:
            state = "効いている"
        cand = "  ← 台帳に載せる候補" if texts >= LEARNED_MIN_TEXTS and (wrong, right) not in ledger_pairs else ""
        lines.append(f"  「{wrong}」→「{right}」  {rows} 回（違う頼みの文 {texts} 件）  {state}{cand}")
    return lines


def _fix_heard(s, wrong, right, stop='', need=''):
    """聞き違い 1 件を直す（VBA の コンボ道具.聞き違いを直す と同じ答え）。
    stop: 誤の直前の字がこの中にあれば直さない／need: カンマ区切り。誤の後ろ HEARD_WINDOW 字以内にどれかがあるときだけ直す。"""
    if not wrong or wrong not in s:
        return s
    if not stop and not need:
        return s.replace(wrong, right)
    n = len(wrong)
    if not n:
        return s
    needs = [w for w in str(need or '').split(',') if w]
    out, i = [], 0
    while i < len(s):
        if s.startswith(wrong, i):
            ok = True
            if stop and i > 0 and s[i - 1] in stop:
                ok = False
            if ok and needs:
                win = s[i + n:i + n + HEARD_WINDOW]
                ok = any(w in win for w in needs)
            if ok:
                out.append(right)
                i += n
                continue
        out.append(s[i])
        i += 1
    return ''.join(out)


# 言い方の言い換え（デザインの棚の色名へ寄せる。外れた依頼: 2026-10-18〜10-19・2026-10-22 指示書035・038）
_COLOR_CONTEXT_RE = re.compile(r'表|色|デザイン|配色|リッチな')
_COLOR_MAP = {
    'リッチなワイン': 'ノーブルボルドー',
    'リッチなボルドー': 'ノーブルボルドー',
    'リッチなグリーン': 'モダンエメラルド',
    'リッチな緑': 'モダンエメラルド',
    'リッチなエメラルド': 'モダンエメラルド',
    'リッチなネイビー': 'エグゼクティブネイビー',
    'リッチな濃紺': 'エグゼクティブネイビー',
    'リッチなチャコール': 'シックチャコール',
    'リッチなグレー': 'シックチャコール',
    'リッチなアンバー': 'ウォームアンバー',
    'リッチな琥珀': 'ウォームアンバー',
    'リッチなインディゴ': 'クラシックインディゴ',
    'リッチな藍': 'クラシックインディゴ',
    'ワイン': 'ノーブルボルドー',
    'ボルドー': 'ノーブルボルドー',
    'エメラルド': 'モダンエメラルド',
    'グリーン': 'モダンエメラルド',
    'チャコール': 'シックチャコール',
    'アンバー': 'ウォームアンバー',
    'インディゴ': 'クラシックインディゴ',
}
_COLOR_PATTERN = re.compile(
    r'(' + '|'.join(re.escape(k) for k in _COLOR_MAP.keys()) + r')(?=色に|に|で|の色)'
)
_PHRASE_REPLACEMENTS = list(_COLOR_MAP.items())





def _get_entry_cache(e):
    c = e.get('_cache')
    if c is not None:
        return c
    ask_parsed = []
    for a in str(e.get('ask') or '').split('|'):
        w = a.strip()
        if w:
            ask_parsed.append((len(w), w.lower()))
    pos_groups = []
    neg_words = []
    for group in str(e.get('combo') or '').split('+'):
        group = group.strip()
        if not group:
            continue
        if group.startswith('-'):
            for x in group[1:].split(','):
                w = x.strip().lower()
                if w:
                    neg_words.append(w)
        else:
            words = []
            for x in group.split(','):
                w = x.strip()
                if w:
                    words.append((len(w), w.lower()))
            if words:
                words.sort(key=lambda x: x[0], reverse=True)
                pos_groups.append(words)
    name = str(e.get('name') or '')
    c = {
        'name': name,
        'cf_ok': bool(e.get('cf_ok')),
        'allow_fix': bool(e.get('allow_fix')),
        'look_only': bool(e.get('look_only')),
        'is_look_only_name': bool(_LOOK_ONLY_RE.search(name)),
        'ask': ask_parsed,
        'pos_groups': pos_groups,
        'neg_words': neg_words,
        'name_len': len(name),
        'name_lower': name.lower(),
    }
    e['_cache'] = c
    return c


def request_normalize(req, entries=None, meta=None):
    """頼みの文の下ごしらえ（純 Python・2026-10-22 指示書035・038／2026-10-23 指示書047で台帳へ）。
    1. 頭のつなぎ語を落とす（台帳「棚の辞書」の つなぎ語 ＋「、」か空白）
    2. 聞き違いの言い換え（台帳の聞き違いを書いた順 → 学び＝違う頼みの文で 2 回以上残った対。票→表・産業ごと→3桁ごと など）
    3. 色の言い換え（色を頼んでいる文のときだけ・指示書038。048 で概念の辞書へ）
    entries（棚の登録）を渡すと、学びの誤が棚の語に入っていないかも見る（VBA は棚を読んで持っている）。
    VBA の コンボ道具.頼みの下ごしらえ と同じ答え。
    """
    s = str(req or '').strip()
    if not s:
        return ''
    if meta is not None:
        filler_re = meta['filler_re']
        pairs = meta['pairs']
    else:
        led = current_ledger()
        filler_re = _LEDGER_CACHE.get('filler_re')
        pairs = [
            (p['wrong'], p['right'], p['stop'], p['need'])
            for p in (led['pairs'] + learned_pairs(entries))
        ]
    if filler_re is not None:
        s = filler_re.sub('', s).strip()
    for wrong, right, stop, need in pairs:
        s = _fix_heard(s, wrong, right, stop, need)
    if _COLOR_CONTEXT_RE.search(s):
        s = _COLOR_PATTERN.sub(lambda m: _COLOR_MAP[m.group(1)], s)
    return s


def _prepare_entries_meta(entries):
    if not entries:
        return {'list': [], 'name_map': {}, 'name_cand': [], 'neg_entries': [], 'cf_list': [],
                'combo_only': [], 'ask_only': [], 'both': [], 'pairs': [], 'filler_re': None}
    meta = getattr(entries, '_meta', None)
    if meta is not None:
        return meta
    cached_list = []
    name_map = {}
    name_cand = []
    neg_entries = []
    cf_list = []
    combo_only = []
    ask_only = []
    both = []
    for e in entries:
        ec = e.get('_cache')
        if ec is None:
            ec = _get_entry_cache(e)
        cached_list.append(ec)
        name_map[ec['name']] = ec
        if ec['name_len'] >= 4:
            name_cand.append(ec)
        if ec['neg_words']:
            neg_entries.append(ec)
        if ec['cf_ok']:
            cf_list.append(ec)
        if ec['pos_groups'] and not ec['ask']:
            combo_only.append(ec)
        elif ec['ask'] and not ec['pos_groups']:
            ask_only.append(ec)
        elif ec['pos_groups'] and ec['ask']:
            both.append(ec)
    led = current_ledger()
    filler_re = _LEDGER_CACHE.get('filler_re')
    pairs = [
        (p['wrong'], p['right'], p['stop'], p['need'])
        for p in (led['pairs'] + learned_pairs(entries))
    ]
    meta = {
        'list': cached_list,
        'name_map': name_map,
        'name_cand': name_cand,
        'neg_entries': neg_entries,
        'cf_list': cf_list,
        'combo_only': combo_only,
        'ask_only': ask_only,
        'both': both,
        'pairs': pairs,
        'filler_re': filler_re,
    }
    try:
        entries._meta = meta
    except Exception:
        pass
    return meta


def shelf_pick(request, entries=(), meta=None):
    """依頼文 → 棚のマクロ名（当たらなければ ''＝AI に回す）。VBA の コンボ道具.先撃ち候補 の Python 版。

    entries は shelf_entries_from_text の並び（登録簿の順＝モジュールの順。同点なら先に出てくる方が勝つ）。
    点＝依頼の語（当たった語の長さの合計）＋依頼の組（各組の最長の語の長さの和）。
    マクロ名そのものを言った依頼（4 字以上）は名前の長さの 2 倍で当てる（手順書の依頼文は除く）。
    """
    if meta is None:
        meta = _prepare_entries_meta(entries)
    req = request_normalize(request, entries, meta=meta)
    if not req:
        return ''

    cached_list = meta['list']
    name_map = meta['name_map']
    name_cand = meta['name_cand']
    neg_entries = meta['neg_entries']

    p = req.find('【手順書：')
    if p >= 0:
        m = re.match(r'【手順書：([^】]*)】', req[p:])
        name = m.group(1).strip() if m else ''
        if name and name in name_map:
            return name
        req = req[:p].strip()
        if not req:
            return ''
    # マクロの解説・書き直し・高速化は AI の役目＝棚に当てない（指示書064。棚の名前でない答えは返さない・2026-10-27）
    if 'マクロ' in req and any(w in req for w in ('説明', '書き直', '高速化', '解説')):
        return ''
    low = req.lower()
    if _SHELF_NOT_WORK_RE.search(req) or (_SHELF_QUESTION_RE.search(low) and not _SHELF_QUESTION_EXCEPT_RE.search(low)):
        return ''
    # 条件付き書式を付ける頼みは、頭に「' 条件付き書式: はい」と書いた棚だけが当たる（無ければ AI の役目・2026-10-22 指示書044）
    cf_block = '条件付き書式' in req and not any(w in req for w in _SHELF_CF_LOOK)
    if not _SHELF_FIX_RE.search(req) and _SHELF_ERROR_MARK_RE.search(low) and _SHELF_WHY_RE.search(req):
        return ''
    points, banned, order = {}, set(), []
    if neg_entries:
        for ec in neg_entries:
            if any(w in low for w in ec['neg_words']):
                banned.add(ec['name'])
    if cf_block:
        for ec in meta['cf_list']:
            name = ec['name']
            if not name:
                continue
            pt = 0
            if ec['pos_groups']:
                c_tot = 0
                for words in ec['pos_groups']:
                    for l, w in words:
                        if w in low:
                            c_tot += l
                            break
                    else:
                        break
                else:
                    pt = c_tot
            if ec['ask']:
                for l, w in ec['ask']:
                    if w in low:
                        pt += l
            if pt:
                if name in points:
                    points[name] += pt
                else:
                    order.append(name)
                    points[name] = pt
    else:
        for ec in meta['combo_only']:
            c_tot = 0
            for words in ec['pos_groups']:
                for l, w in words:
                    if w in low:
                        c_tot += l
                        break
                else:
                    break
            else:
                name = ec['name']
                if name in points:
                    points[name] += c_tot
                else:
                    order.append(name)
                    points[name] = c_tot
        for ec in meta['ask_only']:
            tot = 0
            for l, w in ec['ask']:
                if w in low:
                    tot += l
            if tot:
                name = ec['name']
                if name in points:
                    points[name] += tot
                else:
                    order.append(name)
                    points[name] = tot
        for ec in meta['both']:
            pt = 0
            c_tot = 0
            for words in ec['pos_groups']:
                for l, w in words:
                    if w in low:
                        c_tot += l
                        break
                else:
                    break
            else:
                pt = c_tot
            for l, w in ec['ask']:
                if w in low:
                    pt += l
            if pt:
                name = ec['name']
                if name in points:
                    points[name] += pt
                else:
                    order.append(name)
                    points[name] = pt
    best, top = '', 0
    for name in order:
        if name not in banned and points[name] > top:
            top, best = points[name], name
    is_name_match = False
    if '【手順書：' not in req:
        for ec in name_cand:
            if cf_block and not ec['cf_ok']:
                continue
            nl = ec['name_len']
            if ec['name_lower'] in low:
                is_name_match = True
                if nl * 2 > top:
                    top, best = nl * 2, ec['name']
    if not best:
        return ''
    best_ec = name_map.get(best)
    # 「直して」の頼みに、調べて一覧にするだけの棚は当てない（2026-09-24・エラーの原因を探して直して→ エラー値のセルを一覧にする だった）
    # （「直すとどこまで影響する」は調べる頼み＝「直して」「修正して」と頼んだときだけ）
    # 「直して」の頼みにも当てる棚（' 直す頼みにも当てる: はい）は除外しない（2026-10-22 指示書041）
    if best and best_ec and best_ec['is_look_only_name'] and any(w in req for w in ('直して', '修正して')):
        if not best_ec['allow_fix']:
            best = ''
            best_ec = None
    # 調べるだけの棚は、調べる頼みのときだけ撃つ（2026-10-22 指示書042）
    if best and best_ec:
        if best_ec['look_only'] and not any(w in req for w in ('一覧', '調べ', '確認', 'どこ', '探', 'チェック', '点検', '洗い出', 'リスト', '教えて', '何ページ', '何枚', '何件', '何行', '何列', '何個', '何人', 'いくつ', 'どれだけ', 'どのくらい', 'どれくらい')):
            best = ''
    # 迷いの判定（2026-10-22 指示書042: 倍率1.2）
    if best and '【手順書：' not in req:
        if not is_name_match:
            sec_pt = 0
            for name in order:
                if name not in banned and name != best:
                    if points[name] > sec_pt:
                        sec_pt = points[name]
            if sec_pt > 0 and top < sec_pt * 1.2:
                best = ''
    return best


_LOOK_ONLY_RE = re.compile(r'(一覧にする|調べる|点検する|監査する|探す|検算する|報告する|確かめる|の一覧)$')


def shelf_plan(request, entries=(), pick=None):
    """依頼文 → 撃つ順の棚の並び（純 Python・VBA の コンボ道具.先撃ちの並び と同じ・2026-09-24）。

    2 つ以上の空白・改行・「。」で節に分け、全部の節が棚に当たれば並びを返す（同じマクロは 1 回・表の書き方と罫線と列幅をそろえる は先頭）。
    節が 1 つ・当たらない節がある・撃つのが 1 本だけなら []＝今までどおり shelf_pick が 1 本を選ぶ。「、」では分けない。
    pick（vba_picker の返す関数）を渡すと、節ごとの判定を VBA の正本でする（2026-10-14 段 1）。
    """
    if pick is None:
        def pick(x):
            return shelf_pick(x, entries)
    s = str(request or '').replace('\r\n', '\n').replace('\r', '\n').replace('。', '\n').replace('　', ' ')
    while '  ' in s:
        s = s.replace('  ', '\n')
    s = s.replace('して ', 'して\n').replace('て ', 'て\n')
    parts = [x.strip(' ') for x in s.split('\n') if x.strip(' ')]
    if len(parts) < 2:
        # 062: 明確な順序接続表現による分割（判定の正本 VBA と完全に一致・最長一致順）
        s2 = s
        for conj in (
            'してそのあと', 'てそのあと', 'してさらに', 'てさらに',
            'してから', 'てから', 'たうえで', 'た上で',
            'の後に', 'のあと', 'たあと', 'た後',
            '、さらに', '、そして', 'して、'
        ):
            if conj in s2:
                s2 = s2.replace(conj, '\n')
        parts = [x.strip(' ') for x in s2.split('\n') if x.strip(' ')]
        if len(parts) < 2:
            return []
        # 手段補足句（「〜で」で終わり動詞を含まない、ピボットで・グラフで・テーブルで等）による分割は除外（指示書065）
        _MEANS_RE = re.compile(r'^(?:ピボット|ピボットテーブル|グラフ|棒グラフ|円グラフ|折れ線グラフ|テーブル)で$')
        if any(_MEANS_RE.match(x) for x in parts[1:]):
            return []
    out, tidy = [], False
    for x in parts:
        n = pick(x)
        if not n:
            if x.endswith('して'):
                n = pick(x[:-2])
            elif x.endswith('て'):
                n = pick(x[:-1])
        if not n:
            return []
        # 並びの中の代わりがあれば差し替える（' 並びの代わり: ...・2026-10-22 指示書041）
        for e in entries or ():
            if e.get('name') == n and e.get('plan_substitute'):
                n = e.get('plan_substitute')
                break
        if n == '表の書き方と罫線と列幅をそろえる':
            tidy = True
        elif n not in out:
            out.append(n)
    if tidy:
        out.insert(0, '表の書き方と罫線と列幅をそろえる')
    return out if len(out) >= 2 else []


# ---- 判定の正本は VBA（2026-10-14 段 1・作業ファイル\project\work_order_20261014_shelf_system.md） ----
# 先撃ちの判定は VBA（コンボ道具.先撃ち候補）と、この Python の写し（shelf_pick）の 2 か所にあり、10/12 に条件付き書式の決まりが
# 片方にだけ入って食い違った。Excel につながっているときは VBA を呼ぶ。写しは Excel なしで回す pytest のために残し、
# 食い違いは register-addin（焼く前）と make_public（配る前）で問題集を両方に撃って止める。

def _vba_book_ref(owner):
    return "'" + str(owner).replace("'", "''") + "'!"


def vba_picker(xl, owner, entries=()):
    """棚の持ち主のブックに コンボ道具 があれば、VBA の先撃ち候補を呼ぶ関数を返す。無ければ None（写しの shelf_pick で判定する）。
    呼べなかった 1 回（Excel が忙しい等）は、その頼みだけ写し（entries で shelf_pick）に落とす。"""
    try:
        xl.Workbooks(owner).VBProject.VBComponents('コンボ道具')
    except Exception:
        return None
    ref = _vba_book_ref(owner)

    def pick(req):
        # シートを見る版（2026-10-14 段 3）: ピボット・グラフを手直しする棚は、前のシートにそれが無ければ当てない。
        # 古い秀コンボ（_シートで が無い）なら素の 先撃ち候補
        for fn in ('コンボ道具.先撃ち候補_シートで', 'コンボ道具.先撃ち候補'):
            try:
                return str(xl.Run(ref + fn, str(req or '')) or '')
            except Exception:
                continue
        return shelf_pick(req, entries)
    return pick


def vba_peek(xl, owner, req):
    """VBA の 先撃ちの下見: どの棚がなぜ当たるか（1 行目「撃つ: …」・上位 3 本・除外・弾いた理由）。呼べなければ ''。"""
    try:
        return str(xl.Run(_vba_book_ref(owner) + 'コンボ道具.先撃ちの下見', str(req or '')) or '')
    except Exception:
        return ''


def exam_dir():
    import os
    return os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'fixtures', 'shelf_exam'))


def exam_questions(dirpath=None):
    """問題集 entry_phrases*.tsv（期待するマクロ名<TAB>依頼文。期待が空＝AI に回すのが正解）→ [(期待, 依頼文)]。同じ依頼文は 1 回。
    問題集は手元だけの作業場（公開の写しには入れない）＝無ければ []。"""
    import glob
    import io
    import os
    d = dirpath or exam_dir()
    out, seen = [], set()
    for f in sorted(glob.glob(os.path.join(d, 'entry_phrases*.tsv'))):
        for ln in io.open(f, encoding='utf-8', errors='replace'):
            p = ln.rstrip('\r\n').split('\t')
            if len(p) < 2:
                continue
            q = p[1].strip()
            if not q or '\n' in q or q in seen:
                continue
            seen.add(q)
            out.append((p[0].strip(), q))
    return out


def vba_answers(xl, owner, questions):
    """VBA の 先撃ち候補の一括 で問題を全部判定する（COM の往復 1 回）。呼べなければ None。"""
    try:
        got = str(xl.Run(_vba_book_ref(owner) + 'コンボ道具.先撃ち候補の一括', '\n'.join(questions)) or '')
    except Exception:
        return None
    ans = got.split('\n')
    if len(ans) != len(questions):
        return None
    return ans


# 棚の頭の語の直しをまとめて試す近道（shelf --words・2026-10-29）。1 語ずつ patch して焼いて公開して…を 4 周回して
# 小さな直しに何十分もかかった（シュウさん「近道でいいだろ作れよ」）。直しは表 1 枚に並べ、写しで問題集を全部撃ち比べてから 1 回で書く
_WORDS_HEAD_RE = re.compile(r"^\s*'\s*依頼の(?:語|組)\s*[:：]")


def read_words_edits(path):
    """直しの表（UTF-8 の TSV: 棚の名前<TAB>元の字<TAB>直した字）→ [(名前, 元, 新)]。# で始まる行と空行は読まない。
    元の字を「+語」と書くと、直した字を依頼の語の末尾に | で足す（「+語<TAB>行目の高さ|列目の幅」）。"""
    import io
    out = []
    for ln in io.open(path, encoding='utf-8-sig'):
        ln = ln.rstrip('\r\n')
        if not ln.strip() or ln.lstrip().startswith('#'):
            continue
        p = ln.split('\t')
        if len(p) < 3 or not p[0].strip() or not p[2]:
            raise ValueError(f"直しの表の行の形が違います（棚の名前<TAB>元の字<TAB>直した字）: {ln}")
        out.append((p[0].strip(), p[1], p[2]))
    return out


def _head_lines(lines, sub):
    """lines の中の Sub sub の頭の「依頼の語・依頼の組」の行の番号（0 始まり）。Sub が無ければ None。"""
    pat = re.compile(r'^\s*(?:Public\s+|Private\s+)?Sub\s+' + re.escape(sub) + r'\s*\(')
    for i, ln in enumerate(lines):
        if pat.match(ln):
            out, j = [], i + 1
            while j < len(lines) and (lines[j].strip().startswith("'") or not lines[j].strip()):
                if _WORDS_HEAD_RE.match(lines[j]):
                    out.append(j)
                j += 1
            return out
    return None


def words_apply(modules, edits):
    """[(モジュール名, 本文)] に頭の語の直しを当てる（純 Python）→ (直した [(モジュール名, 本文)], [(モジュール名, 行番号 0 始まり, 前, 後)])。
    直せるのは Sub の頭の「' 依頼の語:」「' 依頼の組:」の行だけ。元の字はその棚の頭の行でちょうど 1 か所に無ければ止める（ValueError）。"""
    mods = [(n, t.split('\r\n') if '\r\n' in t else t.split('\n'), '\r\n' if '\r\n' in t else '\n') for n, t in modules]
    changes = {}
    for sub, old, new in edits:
        hit = None
        for mi, (mname, lines, _nl) in enumerate(mods):
            heads = _head_lines(lines, sub)
            if heads is not None:
                hit = (mi, mname, lines, heads)
                break
        if hit is None:
            raise ValueError(f"棚が見つかりません: {sub}")
        mi, mname, lines, heads = hit
        if old.strip() == '+語':
            asks = [j for j in heads if re.match(r"^\s*'\s*依頼の語\s*[:：]", lines[j])]
            if len(asks) != 1:
                raise ValueError(f"{sub}: 依頼の語の行がありません（+語 は今ある行の末尾に足すだけ）")
            j = asks[0]
            add = new.strip().lstrip('|')
            m = _REG_ASK_RE.match(lines[j])
            if m and any(w in m.group(1).split('|') for w in add.split('|')):
                raise ValueError(f"{sub}: 足す語がもう依頼の語にあります: {add}")
            after = lines[j].rstrip() + '|' + add
        else:
            found = [j for j in heads if old in lines[j]]
            if len(found) != 1 or lines[found[0]].count(old) != 1:
                raise ValueError(f"{sub}: 元の字「{old}」が頭の依頼の語・依頼の組にちょうど 1 か所ありません（{sum(lines[j].count(old) for j in heads)} か所）")
            j = found[0]
            after = lines[j].replace(old, new)
        if (mi, j) not in changes:
            changes[(mi, j)] = lines[j]
        lines[j] = after
    out = [(n, nl.join(lines)) for n, lines, nl in mods]
    diff = [(mods[mi][0], j, before, mods[mi][1][j]) for (mi, j), before in sorted(changes.items())]
    return out, diff


def words_compare(old_text, new_text, questions):
    """直す前と後の棚の本文で問題集を全部撃ち比べる（純 Python）→ (前の正解数, 後の正解数, [(印, 期待, 依頼文, 前, 後)])。
    印は 良くなった／悪くなった／変わった（期待が空＝AI に回すのが正解）。"""
    old_e, new_e = shelf_entries_from_text(old_text), shelf_entries_from_text(new_text)
    ok0 = ok1 = 0
    moved = []
    for want, q in questions:
        a = shelf_pick(q, old_e) or ''
        b = shelf_pick(q, new_e) or ''
        ok0 += a == want
        ok1 += b == want
        if a != b:
            mark = '良くなった' if b == want else ('悪くなった' if a == want else '変わった')
            moved.append((mark, want, q, a, b))
    return ok0, ok1, moved


def parity_diffs(questions, answers, entries, py_answers=None):
    """VBA の答えと写し（shelf_pick）の答えが違う問題 → [(依頼文, VBA, 写し)]（純 Python）。"""
    if py_answers is not None:
        return [(q, a, p) for q, a, p in zip(questions, answers, py_answers) if a != p]
    return [(q, a, shelf_pick(q, entries)) for q, a in zip(questions, answers) if a != shelf_pick(q, entries)]


def vba_snapshot_path(dirpath=None):
    import os
    return os.path.join(dirpath or exam_dir(), 'vba_answers.tsv')


def write_vba_snapshot(questions, answers, dirpath=None, path=None):
    """register-addin が通したときの VBA の答え（依頼文<TAB>棚の名前）を残す＝make_public が配る棚で写しと突き合わせる控え。"""
    import io
    with io.open(path or vba_snapshot_path(dirpath), 'w', encoding='utf-8', newline='\n') as f:
        for q, a in zip(questions, answers):
            f.write(f"{q}\t{a}\n")


_FILLER_PREFIX_RE = re.compile(r'^(?:(?:ええ|あの|えっと|その|まあ|はい)[、,\s]*)+')


def normalize_miss_query(q):
    """頭の「ええ、」などのつなぎ語と空白を落として頼みの型にする（2026-10-22）。"""
    s = str(q or '').strip()
    s = _FILLER_PREFIX_RE.sub('', s).strip()
    return ' '.join(s.split())


def is_test_book(name):
    """ブック名が試験の撃ち（テスト・test・book・捨て・比較・dummy を含む）か判定する（2026-10-22）。"""
    if not name:
        return False
    low = str(name).lower()
    return any(k in low for k in ['テスト', 'test', 'book', '捨て', '比較', 'dummy'])


def missed_requests(path=None, include_test=False):
    """Excelコンボが棚に当てられなかった頼み（%APPDATA%\\秀コンボ\\外れた依頼.csv・CP932）→ 頼みの文の並び（同じ文は 1 回）。
    4 列目のブック名がある場合、include_test=False なら試験の撃ちを除く。3 列の古い行でも読める（2026-10-22）。
    """
    import csv
    import io
    import os
    p = path or os.path.join(os.environ.get('APPDATA', ''), '秀コンボ', '外れた依頼.csv')
    if not os.path.isfile(p):
        return []
    out, seen = [], set()
    for row in csv.reader(io.open(p, encoding='cp932', errors='replace', newline='')):
        if len(row) < 3 or row[0] == '日時':
            continue
        book = row[3].strip() if len(row) >= 4 else ""
        if not include_test and is_test_book(book):
            continue
        q = ' '.join(row[2].split())
        if q and q not in seen:
            seen.add(q)
            out.append(q)
    return out


def _parse_miss_dt(s):
    import datetime
    s = str(s or '').strip()
    for fmt in ('%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M', '%Y-%m-%d %H:%M'):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def shelf_part_scores(pattern, entries=()):
    """頼みの文を空白・「。」で分け、1 つずつ棚に当てる＝部分点（2026-10-23）。
    戻り値: [(頼みの 1 つ, 当たった棚の名前 or ''), ...]。分けても 1 つにしかならない文は空のリスト（部分点は付けない）。
    外れの記録が「全部当たるか・全部外れるか」の 2 択だと、2 つの頼みが混ざった文は
    どちらが足りないのか見えない。足りない部分だけがマクロにする候補になる。"""
    parts = [p for p in re.split(r'[\s　。]+', str(pattern or '')) if p]
    if len(parts) < 2 or not entries:
        return []
    return [(p, shelf_pick(p, entries)) for p in parts]


def shelf_misses(path=None, days=14, all_days=False, text=None, now=None, entries=None):
    """外れた依頼.csv から 2 回以上外れた頼みの型を集計する（2026-10-22）。
    戻り値: (not_hit_list, hit_list)
      各リストの要素: {'pattern': str, 'count': int, 'latest_dt': str, 'candidate': str, 'hit_shelf': str}
    """
    import csv
    import datetime
    import io
    import os
    p = path or os.path.join(os.environ.get('APPDATA', ''), '秀コンボ', '外れた依頼.csv')
    if not os.path.isfile(p):
        return [], []

    ref_now = now or datetime.datetime.now()
    cutoff = None if all_days else (ref_now - datetime.timedelta(days=days))

    # 棚エントリの準備
    if entries is None:
        if text:
            entries = shelf_entries_from_text(text)
        else:
            try:
                from vbam_core import get_workbook, shelf_text
                _xl, wb = get_workbook(None)
                if wb:
                    txt = shelf_text(wb.VBProject)
                    if txt:
                        entries = shelf_entries_from_text(txt)
            except Exception:
                entries = ()
    entries = entries or ()

    groups = {}
    for row in csv.reader(io.open(p, encoding='cp932', errors='replace', newline='')):
        if len(row) < 3 or row[0] == '日時':
            continue
        book = row[3].strip() if len(row) >= 4 else ""
        if is_test_book(book):
            continue
        dt_str = row[0].strip()
        dt = _parse_miss_dt(dt_str)
        if cutoff and dt and dt < cutoff:
            continue
        cand = row[1].strip()
        raw_q = row[2].strip()
        pat = normalize_miss_query(raw_q)
        if not pat:
            continue
        if pat not in groups:
            groups[pat] = {
                'pattern': pat,
                'count': 0,
                'latest_dt': dt_str,
                'latest_dt_obj': dt,
                'candidate': cand,
                'samples': []
            }
        g = groups[pat]
        g['count'] += 1
        if raw_q not in g['samples']:
            g['samples'].append(raw_q)
        if not g['latest_dt_obj'] or (dt and dt > g['latest_dt_obj']):
            g['latest_dt'] = dt_str
            g['latest_dt_obj'] = dt
            g['candidate'] = cand

    # 2回以上のものを抽出
    multi = [g for g in groups.values() if g['count'] >= 2]
    # 多い順、最後の日時の新しい順
    multi.sort(key=lambda x: (-x['count'], x['latest_dt']), reverse=False)

    not_hit = []
    hit = []
    for g in multi:
        hit_name = ''
        if entries:
            hit_name = shelf_pick(g['pattern'], entries)
            if not hit_name:
                for s in g['samples']:
                    hit_name = shelf_pick(s, entries)
                    if hit_name:
                        break
        g['hit_shelf'] = hit_name
        g['parts'] = shelf_part_scores(g['pattern'], entries)
        if hit_name:
            hit.append(g)
        else:
            not_hit.append(g)

    return not_hit, hit


def cmd_shelf_misses(args):
    """`shelf --misses [--days N] [--all]`: 外れた依頼の集計（2026-10-22）。"""
    days = getattr(args, 'days', 14)
    all_days = getattr(args, 'all_days', False)
    target_file = getattr(args, 'posargs', [None])[0] if getattr(args, 'posargs', None) else None

    # 棚のソース
    entries = None
    try:
        from vbam_core import get_workbook, shelf_text
        xl, wb = get_workbook(target_file)
        if wb:
            t = shelf_text(wb.VBProject)
            if t:
                entries = shelf_entries_from_text(t)
    except Exception:
        pass

    not_hit, hit = shelf_misses(days=days, all_days=all_days, entries=entries)

    scope = "全期間" if all_days else f"直近 {days} 日"
    print(f"外れた依頼の集計（{scope}・試験のブックを除く）")
    if not entries:
        # 棚を読めない（Excel が起動していない・秀コンボが開いていない）と全部が「当たらない」に見える（2026-10-23: 解決済みの 6 件が候補に出た）
        print("★ 棚を読めていません（Excel に秀コンボが開いていない）。下の「候補」は当たり判定をしていない一覧です。")
    print("-" * 60)
    print(f"マクロにする候補（2回以上・まだ棚に当たらない物）: {len(not_hit)} 件")
    if not_hit:
        for r in not_hit:
            cand = f"  [{r['candidate']}]" if r['candidate'] else ""
            print(f"  {r['count']:2d} 回  {r['latest_dt']}  {r['pattern']}{cand}")
            parts = r.get('parts') or []
            if parts:
                got = sum(1 for _p, n in parts if n)
                print(f"        部分点 {got}/{len(parts)}（頼みを分けて 1 つずつ当てた）")
                for p, n in parts:
                    print(f"          {'○' if n else '×'} {p}" + (f" → 『{n}』" if n else "  ← 棚が足りない所"))
    else:
        print("  （該当なし）")

    print("-" * 60)
    print(f"今は棚に当たる物: {len(hit)} 件")
    if hit:
        for r in hit:
            print(f"  {r['count']:2d} 回  {r['latest_dt']}  {r['pattern']}  → 『{r['hit_shelf']}』")
    else:
        print("  （該当なし）")

    # AI の選んだ答えの集計（2026-10-22 指示書040 2b）
    print("-" * 60)
    ai_tsv = os.path.join(os.environ.get('APPDATA', ''), '秀コンボ', 'AIの選んだ答え.tsv')
    ai_cands, ai_words = [], []
    if os.path.isfile(ai_tsv):
        import csv
        ai_groups = {}
        for row in csv.reader(open(ai_tsv, encoding='cp932', errors='replace'), delimiter='\t'):
            if len(row) < 5 or row[0] == '日時':
                continue
            book = row[1].strip()
            if is_test_book(book):
                continue
            req = row[2].strip()
            chosen = row[4].strip()
            k = (req, chosen)
            ai_groups[k] = ai_groups.get(k, 0) + 1
        for (req, chosen), cnt in ai_groups.items():
            if cnt >= 2:
                if not chosen:
                    ai_cands.append((req, cnt))
                else:
                    ai_words.append((req, chosen, cnt))
    print(f"AI の選んだ答え（2回以上・試験のブックを除く）: {len(ai_words) + len(ai_cands)} 件")
    if ai_words or ai_cands:
        for req, chosen, cnt in ai_words:
            print(f"  棚の語と問題集に足す候補: 「{req}」 → {chosen}（{cnt} 回）")
        for req, cnt in ai_cands:
            print(f"  棚の候補（どれでもない）: 「{req}」（{cnt} 回）")
    else:
        print("  （該当なし）")

    # 聞き違いの学び（2026-10-23 指示書047）: 違う頼みの文で 2 回以上残った対は自動で直る。台帳に載せる候補も出す
    print("-" * 60)
    heard = learned_report(entries)
    print(f"聞き違いの学び（shelf --heard 誤 正 --q \"頼みの文\" で残した物）: {len(heard)} 対")
    if heard:
        for ln in heard:
            print(ln)
    else:
        print("  （まだありません）")
    return True






def check_hardcoded_shelf_names(xl, owner, text):
    """判定コード（Pythonのshelf_pick/shelf_plan、VBAの判定関数）の中に棚名がハードコードされていないか検査（指示書041）。"""
    stops = []
    entries = shelf_entries_from_text(text)
    shelf_names = {e['name'] for e in entries if len(e['name']) >= 4}
    # 基本の整え棚（tidy先頭配置）など共通基盤で許容される棚名
    allowed_names = {'表の書き方と罫線と列幅をそろえる'}

    # 1. Python 側の shelf_pick / shelf_plan のソース検査
    import inspect
    for fn in (shelf_pick, shelf_plan):
        try:
            fn_src = inspect.getsource(fn)
            # コメント行を除外して検査
            code_lines = [ln.split('#')[0] for ln in fn_src.splitlines()]
            code_body = '\n'.join(code_lines)
            for name in shelf_names - allowed_names:
                if f"'{name}'" in code_body or f'"{name}"' in code_body:
                    stops.append(('止める', fn.__name__, f"判定コードの中に棚名「{name}」の決め打ちがあります"))
        except Exception:
            pass

    # 2. VBA 側のコンボ道具の判定関数検査
    if xl and owner:
        try:
            wb = xl.Workbooks(owner)
            cm = wb.VBProject.VBComponents('コンボ道具').CodeModule
            vba_procs = ['調べるだけの棚か', '先撃ちの並び', '棚の約束の例外', '約束を破ったか']
            for proc in vba_procs:
                try:
                    s = cm.ProcStartLine(proc, 0)
                    cnt = cm.ProcCountLines(proc, 0)
                    proc_lines = [cm.Lines(i, 1) for i in range(s, s + cnt)]
                    # コメント行を除去
                    clean_lines = []
                    for ln in proc_lines:
                        c_idx = ln.find("'")
                        clean_lines.append(ln[:c_idx] if c_idx >= 0 else ln)
                    code_body = '\n'.join(clean_lines)
                    for name in shelf_names - allowed_names:
                        if f'"{name}"' in code_body:
                            stops.append(('止める', f"コンボ道具.{proc}", f"判定コードの中に棚名「{name}」の決め打ちがあります"))
                except Exception:
                    pass
            # 2026-10-27 番頭点検: modMacroUchi 等への試験の表の語（単価・税・返却・貸出中）や 1.1/0.1 倍決め打ち検査
            for comp in wb.VBProject.VBComponents:
                cname = str(comp.Name)
                if not cname.startswith('modMacroUchi'):
                    continue
                cm_u = comp.CodeModule
                n_u = cm_u.CountOfLines
                if not n_u:
                    continue
                for ln_u in cm_u.Lines(1, n_u).splitlines():
                    s_u = ln_u.strip()
                    if not s_u or s_u.startswith("'"):
                        continue
                    c_i = ln_u.find("'")
                    code_u = (ln_u[:c_i] if c_i >= 0 else ln_u).strip()
                    if re.search(r'"[^"]*(?:返却|貸出中)[^"]*"', code_u):
                        stops.append(('止める', cname, f"試験の表の語の決め打ち「{code_u}」（本体に試験の表の語を書かない）"))
                    if re.search(r'(?:\*\s*(?:1\.1|0\.1)\b|"\s*\*\s*(?:1\.1|0\.1)")', code_u):
                        stops.append(('止める', cname, f"税率（1.1倍・0.1倍）の決め打ち「{code_u}」（指示書に無い決め打ち処理を足さない）"))
        except Exception:
            pass
    return stops


def shelf_check(xl, owner, text, write_snapshot=False, warn=None, hold_changed=False):
    """棚の当たり方の検査を 1 回で（2026-10-14 段 2）。(焼いてよいか, 報告の行) を返す。
    1. 一般の語の検査（shelf_lint）の「止める」
    2. 問題集（entry_phrases*.tsv＝ありがちな頼み集を含む）を VBA（正本）に撃ち、写し（shelf_pick）と食い違えば止める
    3. 問題集の期待と違う答え＝横取り（または当たらなくなった頼み）があれば止める
    4. 外れた依頼（Excelコンボが AI に回した頼み）の当たり先が前の検査から変わった物を並べる（止めない。新しく棚に当たるのが正しいかは人が見る）
    通ったら VBA の答えを控える（write_snapshot）＝make_public が配る棚で写しと突き合わせる・次の検査の 4 の比べ元。
    warn（並び）を渡すと 4 の「当たり先が変わった」を足す。hold_changed=True なら、変わった物がある間は控えを書かない
    （register-addin が --why 無しで止まったのに、控えだけ進んで次の検査で消えないように・2026-10-28 番頭点検）"""
    import time as _t
    t0 = _t.perf_counter()
    lines, ok = [], True
    changed = []
    # 2026-10-22 指示書041: 判定コードの中の棚名決め打ち検査
    t_start = t0
    hardcoded_stops = check_hardcoded_shelf_names(xl, owner, text)
    lint_items = shelf_lint(text)
    stops = [x for x in lint_items if x[0] == '止める'] + hardcoded_stops
    notes = [x for x in lint_items if x[0] == '注意']
    t_lint = _t.perf_counter() - t_start
    if stops:
        ok = False
        lines.append(f"語の検査: 止める {len(stops)} 件")
        lines += [f"  {n}: {m}" for _lv, n, m in stops[:20]]
    else:
        lines.append("語の検査: 止める物なし" + (f"（注意 {len(notes)} 件・shelf --check で見る）" if notes else ""))
    entries = shelf_entries_from_text(text)
    valid_shelf_names = {e['name'] for e in entries}
    qa = exam_questions()
    ans = None
    t_vba = 0.0
    t_py = 0.0
    t_diff = 0.0
    t_entry = 0.0
    if qa:
        qs = [q for _e, q in qa]
        t_v_start = _t.perf_counter()
        ans = vba_answers(xl, owner, qs)
        t_vba = _t.perf_counter() - t_v_start
        if ans is None:
            lines.append("問題集: VBA の 先撃ち候補の一括 を呼べませんでした（照合していません）")
        else:
            before_path = os.path.join(exam_dir(), 'answers_before.tsv')
            if not os.path.isfile(before_path):
                write_vba_snapshot(qs, ans, path=before_path)
            t_d_start = _t.perf_counter()
            meta = _prepare_entries_meta(entries)
            t_pk0 = _t.perf_counter()
            py_ans = [shelf_pick(q, entries, meta=meta) for q in qs]
            t_pk = _t.perf_counter() - t_pk0
            # 2026-10-27 番頭点検: 「棚の名前」か「空（＝AIへ）」以外の第3の返却値（「AIへ: …」等）を止める
            t_bv0 = _t.perf_counter()
            bad_vba = [(q, a) for q, a in zip(qs, ans) if a and a not in valid_shelf_names]
            bad_py = [(q, p) for q, p in zip(qs, py_ans) if p and p not in valid_shelf_names]
            t_bv = _t.perf_counter() - t_bv0
            t_df0 = _t.perf_counter()
            diffs = parity_diffs(qs, ans, entries, py_answers=py_ans)
            t_df = _t.perf_counter() - t_df0
            t_ms0 = _t.perf_counter()
            miss = [(q, e, a) for (e, q), a in zip(qa, ans) if e != a]
            t_ms = _t.perf_counter() - t_ms0
            if miss:
                ok = False
                lines.append(f"問題集の横取り・当たらなくなった頼み: {len(miss)} 問（{len(qs)} 問中）")
                for q, e, a in miss[:12]:
                    why = vba_peek(xl, owner, q).split('\n')[1:2]
                    lines.append(f"  「{q}」 期待={e or '(AIへ)'}  今={a or '(AIへ)'}" + (f"  ← {why[0].strip()}" if why and why[0].strip() else ''))
            t_diff = _t.perf_counter() - t_d_start
            lines.append(f"    照合内訳: 写し計算 {t_pk:.3f}秒・返却値検査 {t_bv:.3f}秒・差分比較 {t_df:.3f}秒・期待照合 {t_ms:.3f}秒")
            if not diffs and not miss and not bad_vba and not bad_py:
                t_e_start = _t.perf_counter()
                # 2026-10-22 指示書039: 入口の照合（Excelコンボ判定 vs seiri判定 plan_full）
                reg = registry_from_text(text)
                py_map = dict(zip(qs, py_ans))
                pick_fn = lambda x: py_map.get(x) if x in py_map else shelf_pick(x, entries)
                entry_diffs = []
                for q, v_ans in zip(qs, ans):
                    pf = plan_full(q, reg, entries=entries, pick=pick_fn)
                    s_ans = pf.get('chosen', [''])[0] if pf.get('chosen') else ""
                    if v_ans != s_ans:
                        entry_diffs.append((q, v_ans, s_ans))
                t_entry = _t.perf_counter() - t_e_start
                if entry_diffs:
                    ok = False
                    lines.append(f"入口の照合（Excelコンボ vs seiri）の食い違い: {len(entry_diffs)} 問（{len(qs)} 問中）")
                    lines += [f"  「{q}」 Excelコンボ={v or '(AIへ)'}  seiri={s or '(AIへ)'}" for q, v, s in entry_diffs[:12]]
                else:
                    lines.append(f"問題集: {len(qs)} 問で VBA・写し・seiri の入口がすべて一致・全部が期待どおり")
    t_m_start = _t.perf_counter()
    misses = missed_requests()
    miss_ans = vba_answers(xl, owner, misses) if misses else None
    snap_miss = vba_snapshot_path().replace('vba_answers.tsv', 'vba_answers_miss.tsv')
    if miss_ans is not None:
        bad_miss = [(q, a) for q, a in zip(misses, miss_ans) if a and a not in valid_shelf_names]
        if bad_miss:
            ok = False
            lines.append(f"外れた依頼で約束違反の返却値（棚の名前でも空でもない第3の答え）: {len(bad_miss)} 種")
            lines += [f"  「{q[:60]}」 返却値={a!r}" for q, a in bad_miss[:12]]
        prev = dict(read_vba_snapshot_file(snap_miss) or [])
        changed = [(q, prev.get(q, ''), a) for q, a in zip(misses, miss_ans) if prev and prev.get(q, '') != a]
        hit = sum(1 for a in miss_ans if a)
        lines.append(f"外れた依頼: {len(misses)} 種のうち、今は棚に当たる {hit} 種" +
                     (f"・前の検査から当たり先が変わった {len(changed)} 種（正しいか見る）" if changed else ""))
        lines += [f"  「{q[:60]}」 前={p or '(AIへ)'} → 今={a or '(AIへ)'}" for q, p, a in changed[:12]]
        if changed and not write_snapshot:
            lines.append(f"⚠ 【報告前停止】前の検査から当たり先が変わった頼みが {len(changed)} 種あります。報告の前に止まって理由を確認してください")
        if changed and warn is not None:
            warn.append(f"外れた依頼の当たり先が前の検査から変わった: {len(changed)} 種（上の「前 → 今」が正しいか見る）")
    t_miss = _t.perf_counter() - t_m_start
    t_s_start = _t.perf_counter()
    if ok and write_snapshot and not (hold_changed and changed):
        if qa and ans is not None:
            write_vba_snapshot([q for _e, q in qa], ans)
        if miss_ans is not None:
            write_vba_snapshot(misses, miss_ans, path=snap_miss)
    t_snap = _t.perf_counter() - t_s_start
    t_total = _t.perf_counter() - t0
    lines.append(f"（内訳: 語の検査 {t_lint:.2f}秒・VBAに撃つ {t_vba:.2f}秒・照合 {t_diff:.2f}秒・入口 {t_entry:.2f}秒・外れた依頼 {t_miss:.2f}秒・控え {t_snap:.2f}秒／計 {t_total:.2f}秒）")
    return ok, lines


def read_vba_snapshot_file(path):
    import io
    import os
    if not os.path.isfile(path):
        return None
    rows = []
    for ln in io.open(path, encoding='utf-8'):
        q, _, a = ln.rstrip('\n').partition('\t')
        rows.append((q, a))
    return rows


def read_vba_snapshot(dirpath=None):
    import io
    import os
    p = vba_snapshot_path(dirpath)
    if not os.path.isfile(p):
        return None
    rows = []
    for ln in io.open(p, encoding='utf-8'):
        q, _, a = ln.rstrip('\n').partition('\t')
        rows.append((q, a))
    return rows


def _ask_best_word(ask, req):
    """依頼の語のうち、依頼文に当たったいちばん長い語（純 Python）。無ければ ''。"""
    best = ''
    for alt in str(ask or '').split('|'):
        try:
            m = re.search(alt, req) if alt else None
        except re.error:
            continue
        if m and len(m.group(0)) > len(best):
            best = m.group(0)
    return best


def plan_of(request, registry=()):
    """依頼文 → (先撃ちするか, 重複行を消すも撃つか, マクロの扱わない仕事があるか)（純 Python・登録簿なしなら 9/11 の形）。"""
    p = plan_full(request, registry)
    if not p['tidy'] and not p['extras']:
        return False, False, False
    return True, p['dedupe'], p['other']


def _macro_owner(xl):
    """「表の書き方と罫線と列幅をそろえる」を持つブックの名前（モジュール「表の整理」を先に探す＝全モジュールを読まない）。無ければ None。"""
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception:
        return None
    from vbam_vba import _project_book_name
    for p in projects:
        try:
            if shelf_module_of(p, _MACRO_TIDY):
                name = _owner_name(xl, p)
                if name:
                    return name
        except Exception:
            continue
    return None


def _owner_name(xl, p):
    """プロジェクト → Run で名指しできるブック名。名前が指すブックに本当に「表の整理」があるかまで確かめる
    （2026-09-17: 未保存の Book4 が開いていると _project_book_name が別プロジェクトの名前を Book4 と取り違え、
    'Book4'!表の書き方と罫線と列幅をそろえる を撃って「マクロを実行できません」で先撃ちが止まった）。"""
    from vbam_vba import _project_book_name
    try:
        cm = p.VBComponents(_MACRO_MODULE).CodeModule
        n = int(cm.CountOfLines)
        sig = (n, cm.Lines(1, min(n, 5)) if n else '')
    except Exception:
        return None

    def _same(name):
        try:
            c2 = xl.Workbooks(name).VBProject.VBComponents(_MACRO_MODULE).CodeModule
            n2 = int(c2.CountOfLines)
            return (n2, c2.Lines(1, min(n2, 5)) if n2 else '') == sig
        except Exception:
            return False
    # ThisWorkbook の Name プロパティ＝そのプロジェクトのブック名（アドイン・非表示・未保存でも取れる）
    try:
        name = str(p.VBComponents('ThisWorkbook').Properties('Name').Value)
        if name and _same(name):
            return name
    except Exception:
        pass
    name = _project_book_name(xl, p)
    if name and _same(name):
        return name
    # Filename が取れない（アドイン・未保存）と _project_book_name は名前を取り違える。開いているブックと
    # 読み込まれているアドインを、同じモジュールの中身（行数と頭 5 行）で照らして探す
    cands = []
    try:
        cands += [str(w.Name) for w in xl.Workbooks]
    except Exception:
        pass
    try:
        cands += [str(a.Name) for a in xl.AddIns if a.Installed]
    except Exception:
        pass
    for c in cands:
        if _same(c):
            return c
    return None


def _table_anchor(ws):
    """表の見出しの左端のセル（tidy と検査の当て先）。読めなければ None。"""
    from vbam_hands import _rows_of, _guess_header_row
    from vbam_view import _guess_header_idx
    try:
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        rows = [list(r) for r in _rows_of(ur.Value)]
    except Exception:
        return None
    if not rows:
        return None
    # 材料の「見出し行の推定」と同じ関数で決める（2026-09-17: 表題が見出しの真上にある表で、格子の推定が
    # 表題の行を見出しに取り、tidy と検査が 1 行ずれて「見出しが空」「SUM が本文 1 行を入れていない」と出た）
    h = None
    try:
        got = _guess_header_row(ws)
        if got and r0 <= int(got[0]) < r0 + len(rows):
            h = int(got[0]) - r0
    except Exception:
        h = None
    if h is None:
        h = _guess_header_idx(rows)
    h = 0 if h is None or h < 0 else h
    j = next((k for k, v in enumerate(rows[h]) if v not in (None, '')), 0)
    return f"{_col_letter(c0 + j)}{r0 + h}"


def _changed_cells(before, now, r0, c0):
    """撃つ前の控え（{'row','col','values'}）と今の値（左上 r0,c0 の 2 次元）→ 値が入って変わったセル [(行, 列)]（純 Python）。
    消えたセル（前の結果の残りを消した所）は数えない＝表の当て先にならない。"""
    from vbam_hands import _text_of
    bv = (before or {}).get('values') or []
    br, bc = int((before or {}).get('row') or 1), int((before or {}).get('col') or 1)
    out = []
    for i, row in enumerate(now or ()):
        for j, v in enumerate(row or ()):
            t = _text_of(v)
            if t == '':
                continue
            r, c = r0 + i, c0 + j
            bi, bj = r - br, c - bc
            old = _text_of(bv[bi][bj]) if 0 <= bi < len(bv) and 0 <= bj < len(bv[bi] or ()) else ''
            if t != old:
                out.append((r, c))
    return out


def _written_anchors(ws, before, limit=6):
    """登録簿のマクロが値を書いた表ごとの見出しの左端（COM）。→ [番地]。読めなければ []。
    2026-09-17 夜: 帳票の右に一覧・明細の右に集計表を作るマクロで、tidy と検査が元の表（帳票・明細）にだけ当たり、
    作った表には書式が付かず、元の表の汚れ（明細の空白つき金額・帳票の桁区切り）を「残りの指摘」にして AI に回していた。"""
    from vbam_hands import _rows_of
    from vbam_edit import _region_of_cell
    if not before or before.get('too_big') or 'values' not in before:
        return []
    try:
        ur = ws.UsedRange
        cells = _changed_cells(before, [list(r) for r in _rows_of(ur.Value)], int(ur.Row), int(ur.Column))
    except Exception:
        return []
    regions, anchors = [], []
    for r, c in cells:
        if any(a <= r <= b and x <= c <= y for a, x, b, y in regions):
            continue
        try:
            rg = _region_of_cell(ws, ws.Cells(r, c))
            a, x = int(rg.Row), int(rg.Column)
            b, y = a + int(rg.Rows.Count) - 1, x + int(rg.Columns.Count) - 1
        except Exception:
            continue
        regions.append((min(a, r), min(x, c), max(b, r), max(y, c)))
        if _continues_above(ws, regions[-1], regions[:-1]):
            continue                                     # 上の表の続き（空行で切れた下の塊）＝別の表として仕上げ・検査しない
        if f"{_col_letter(x)}{a}" not in anchors:
            anchors.append(f"{_col_letter(x)}{a}")
        if len(anchors) >= limit:
            break
    return anchors


def _continues_above(ws, region, others, gap=3):
    """region が、上にある表（others のどれか）の続きか（COM）: 左端の列が同じで、間が空行 gap 行以内で、
    region の 1 行目に数（＝見出しでなくデータの行）がある。
    2026-09-17 夜: 空行 2 行で上下に分かれた旧システムの表で、下の塊の 1 行目（データ）を見出しとして検査し
    「見出しが空の列（D）」を AI に回した（仕上げもデータの行に見出しの書式を当てる形だった）。"""
    from vbam_hands import _rows_of
    a, x, _b, y = region
    for a1, x1, b1, _y1 in others:
        if x1 != x or not (b1 < a <= b1 + 1 + gap):
            continue
        try:
            row = (_rows_of(ws.Range(ws.Cells(a, x), ws.Cells(a, y)).Value) or [[]])[0]
        except Exception:
            return False
        if any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in row):
            return True
    return False


def _rehearse(xl, wb, sheet, plan, owner, sel_cols=None):
    """登録簿のマクロを写しで先に撃つ。→ (撃たない登録簿の行, 理由)。全部よければ ([], '')。
    撃つ順（表の書き方と罫線と列幅をそろえる → 全列が同じ重複行を削除する → 登録簿）は本番と同じ。写しは wb.SaveCopyAs（未保存の変更も入る・人のブックの
    保存状態とパスは変わらない）。"""
    import os
    import tempfile
    from vbam_rehearse import rehearse_steps
    extras = list(plan['extras'])
    by_owner = {}
    for e in extras:
        by_owner.setdefault(e.get('owner') or owner, []).append(e)
    if len(by_owner) != 1:
        return extras, "持ち主のブックが複数で試せませんでした"
    own = next(iter(by_owner))
    try:
        text = shelf_texts(xl.Workbooks(own).VBProject)      # 「表の整理_〜」に分かれていてもモジュールごとに（2026-09-23）
        if not text:
            raise ValueError(f"「{_MACRO_MODULE}」がありません")
    except Exception as ex:
        return extras, f"マクロの本文を読めませんでした（{ex}）"
    names = ([_MACRO_TIDY] + ([_MACRO_DEDUPE] if plan['dedupe'] else []) if plan['tidy'] and owner == own else [])
    names += [e['name'] for e in extras]
    ext = os.path.splitext(str(wb.Name))[1] or '.xlsx'
    work = tempfile.mkdtemp(prefix='_prefire_copy_')
    copy = os.path.join(work, '写し' + ext)
    try:
        wb.SaveCopyAs(copy)
    except Exception as ex:
        return extras, f"写しを作れませんでした（{ex}）"
    try:
        steps = rehearse_steps(text, names, copy, sheet, _MACRO_MODULE, sel_cols or {})
    except Exception as ex:
        return extras, f"試し撃ちが止まりました（{ex}）"
    finally:
        import shutil
        shutil.rmtree(work, ignore_errors=True)
    by_name = {s['name']: s for s in steps}
    for nm in names:
        s = by_name.get(nm)
        if s is None:
            return extras, "前のマクロで止まったので試せませんでした"
        if not s['ok']:
            return extras, f"{nm} が{s['why']}"
        if nm in [e['name'] for e in extras] and not s['changed']:
            return extras, f"{nm} が表を何も変えずに終わりました（この表の形はそのマクロの想定外）"
    return [], ''


def _output_sheets(wb, sheet):
    """マクロが作る先のシート＝その回の持ち物（純 COM）。アクティブなシートを元にしたピボットのあるシートと、
    クエリで読み込んだテーブルのあるシート。前の結果が残っていて作り直したときに、それを
    「頼んでいない変化」と呼ばないため（2026-09-18 朝: 値は正解と一致したのに、前の結果のあった
    4 本（月別ピボット・行列ピボット・ダッシュボード・PQ 追加集計）が入口から AI に回っていた）。"""
    out = set()
    try:
        shs = list(wb.Worksheets)
    except Exception:
        return out
    for sh in shs:
        try:
            nm = str(sh.Name)
        except Exception:
            continue
        if nm == sheet:
            continue
        try:
            pts = sh.PivotTables()
            for i in range(1, int(pts.Count) + 1):
                try:
                    if sheet in str(pts.Item(i).SourceData or ''):
                        out.add(nm)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            los = sh.ListObjects
            for i in range(1, int(los.Count) + 1):
                try:
                    if los.Item(i).QueryTable is not None:
                        out.add(nm)
                except Exception:
                    pass
        except Exception:
            pass
    return out


def macro_first(request, sheet, wb, max_turns, run_id):
    """先撃ち 1 回。戻り値:
      None                           … 先撃ちしない（依頼が整理でない・マクロが無い）
      {'result': r}                  … マクロで合格（AI を呼ばずに終わる。r は run_agent と同じ形）
      {'result': None, 'ran': True…} … 撃ったが残りがある／ほかの仕事がある＝Sonnet の往復へ（控えを引き継ぐ）
    """
    xl = wb.Application
    # 迷ったら撃たずに候補を AI に返す（2026-10-22 指示書040 2b）
    all_ent = None
    try:
        from vbam_core import shelf_text
        _txt = shelf_text(wb.VBProject) if hasattr(wb, 'VBProject') else ''
        if _txt:
            all_ent = shelf_entries_from_text(_txt)
    except Exception:
        pass
    registry = registry_of(xl)
    is_d, cands_str, guide = check_doubt(request, all_ent or registry)
    if is_d:
        remember_doubt(request, cands_str, wb.Name)
        print(f'（{guide}）')
        return None                       # 鍛えたマクロの登録簿（.bas の頭 2 行・2026-09-17）
    words = None
    if any(e.get('headers') for e in registry):
        try:
            words = _top_words(wb.Sheets(sheet))
        except Exception as ex:
            print(f"（シートの見出しを読めませんでした: {ex}。見出しの要る登録簿のマクロは撃ちません）")
            words = set()                            # 読めなければ、見出しの要る Sub は撃たない（安全側）
    shape = None
    if any(e.get('shape') for e in registry):
        shape = _shape_of_sheet(wb.Sheets(sheet), wb)         # 形の要る仕事は、形がそろうときだけ候補（2026-09-18 朝）
    plan = plan_full(request, registry, words, shape)
    # 列を選ぶ必要のある Sub（' 選ぶ列: N）は、依頼の語と見出しが一致する列を道具が選んでから撃つ（2026-09-18）
    sel_cols = {}
    if any(e.get('select') for e in plan['extras']):
        cells = _header_cells(wb.Sheets(sheet))
        have = current_columns(wb.Sheets(sheet))         # 人が列を選んでいれば、それを使う
        drop = []
        for e in [x for x in plan['extras'] if x.get('select')]:
            lo, hi = e['select']
            if lo <= len(have) <= hi:
                sel_cols[e['name']] = have
                print(f"（{e['name']} は人が選んでいる {len(have)} 列を使います）")
                continue
            cols = pick_columns(request, cells, e['select'])
            if not cols:
                drop.append(e['name'])
                print(f"（登録簿の {e['name']} は列を選ぶ仕事ですが、依頼の語と見出しが一致する列が "
                      f"{e['select'][0]}〜{e['select'][1]} 列になりません＝撃たずに AI に回します）")
            else:
                sel_cols[e['name']] = cols
        if drop:
            registry = [e for e in registry if e['name'] not in set(drop)]
            plan = plan_full(request, registry, words, shape)
    for e in plan.get('skipped') or []:
        if e.get('why_shape'):
            print(f"（登録簿の {e['name']} は依頼の語に当たりましたが、この表に「{'・'.join(e['why_shape'])}」が無いので撃ちません）")
        else:
            print(f"（登録簿の {e['name']} は依頼の語に当たりましたが、シートに見出し「{'・'.join(e['headers'])}」が"
                  "そろっていないので撃ちません）")
    for e in plan.get('shadowed') or []:
        print(f"（登録簿の {e['name']} も依頼の語に当たりましたが、依頼はより長い語で別の仕事を名指ししているので撃ちません）")
    if not plan['fire']:
        return None
    dedupe, other = plan['dedupe'], plan['other']
    owner = _macro_owner(xl) if plan['tidy'] else None
    if plan['tidy'] and not owner:
        # 整理の語はあるが「表の書き方と罫線と列幅をそろえる」が無い＝登録簿の Sub だけで撃てるならそれだけ、無ければ AI へ
        if not plan['extras']:
            return None
    if not plan['tidy'] and not plan['extras']:
        return None
    if not owner and not any(e.get('owner') for e in plan['extras']):
        return None                                   # 撃てる持ち主が 1 つも無い
    t_start = time.time()
    tool_sec = {}
    if plan['extras']:
        # 鍛えて登録したマクロは、人の Excel で撃つ前に写しで試す（実行時エラーの窓で固まらない・何もしないで抜けたら AI へ）
        t0 = time.time()
        dropped, why = _rehearse(xl, wb, sheet, plan, owner, sel_cols)
        tool_sec['試し撃ち'] = time.time() - t0
        if dropped:
            print(f"（登録簿の {'・'.join(e['name'] for e in dropped)} は写しで試すと{why}。人のブックでは撃たず、"
                  "その仕事は AI に回します）")
            gone = {x['name'] for x in dropped}
            plan = plan_full(request, [e for e in registry if e['name'] not in gone], words, shape)
            dedupe, other = plan['dedupe'], plan['other']
            if not plan['tidy'] and not plan['extras']:
                return None
            if plan['tidy'] and not owner:
                owner = _macro_owner(xl)
                if not owner:
                    return None
    ws = wb.Sheets(sheet)
    book = str(wb.Name)
    # 控え（本番の往復と同じ: シートの控え・ブックの控え＝--undo・頼んでいない変化を見るための控え）
    t0 = time.time()
    before = va._sheet_snapshot(ws) or {}
    va._agent_backup(wb, ws, sheet, request, run_id)
    book_before = None
    if va._book_cell_count(wb) <= va._INV_BOOK_MAX_CELLS:
        book_before = va._snapshot_book(wb, own={sheet})      # 触らないシートは軽い控え（2026-09-16）
    tool_sec['控え'] = time.time() - t0
    # 撃つ（マクロはアクティブシートに効く＝対象シートを前に出してから）
    t0 = time.time()
    wb.Activate()
    ws.Activate()
    shots = []                                        # [(持ち主のブック, Sub 名)] 撃つ順
    if owner:
        shots.append((owner, _MACRO_TIDY))
        if dedupe:
            shots.append((owner, _MACRO_DEDUPE))
    shots += [(e.get('owner') or owner, e['name']) for e in plan['extras'] if (e.get('owner') or owner)]
    names = [nm for _o, nm in shots]
    for o, nm in shots:
        if sel_cols.get(nm):
            select_columns(ws, sel_cols[nm])         # 列を選ぶ仕事は、選んでから撃つ（マクロは Selection を見る）
        from vbam_vba import run_book_macro
        run_book_macro(xl, o, nm)                     # ハーネス経由＝実行時エラーで窓を出して止まらない（2026-09-24 総点検）
    tool_sec['マクロ'] = time.time() - t0
    owner = owner or shots[0][0]
    owners_txt = '・'.join(sorted({o for o, _n in shots}))
    print(f"マクロの先撃ち: {' → '.join(names)}（{owners_txt}・{tool_sec['マクロ']:.2f} 秒・AI は呼んでいません）")
    with contextlib.suppress(Exception):
        va.fired_append(run_id, book, sheet, request, names, [o for o, _n in shots])   # 使用の記録（引退の判断に使う・2026-09-18 夜）
    # 道具の仕上げ（tidy）と終わりの検査
    anchor = _table_anchor(ws)
    targets = [anchor] if anchor else []
    if plan['extras']:
        # 登録簿のマクロが書いた表に当てる（元の表は「整えて」と頼まれたときだけ）。書いた所が読めなければ元の表
        wrote = _written_anchors(ws, before)
        if wrote:
            targets = ([anchor] if (anchor and plan['tidy']) else []) + [a for a in wrote if a != anchor or not plan['tidy']]
        elif not plan['tidy']:
            # セルに何も書かないマクロ（グラフを作る・別のシートに出す）は、元の表を仕上げも検査もしない
            # ＝元の表にもとからある指摘（見出しの空の列・罫線の継ぎはぎ）で AI に回さない（2026-09-18 朝）
            targets = []
    # 絞り込み（行を隠す）が仕事のマクロの後は仕上げを当てない＝tidy は書式が隠れた行を飛ばさないよう
    # 絞り込みを解くので、マクロの結果（隠れた行）が消える（2026-09-18 朝: 職員名簿の絞り込みが
    # 「一発合格」と出ながら隠れた行 0 だった）
    if targets and any(re.search(r'絞り込|フィルタ|隠す|抽出', (e.get('ask') or '') + ' ' + ' '.join(e.get('handles') or []))
                       for e in plan['extras']):
        print("仕上げ（tidy）は当てません: 絞り込みを扱うマクロの後は、書式を当てると絞り込みが解けます")
        targets = []
    t0 = time.time()
    if targets:
        va._run_cmd(['tidy', '--sheet', sheet] + targets, wb)
    tool_sec['仕上げ'] = time.time() - t0
    t0 = time.time()
    audit = va._audit_now(wb, sheet, targets) if targets else []
    c_block, c_noticed = va._content_now(wb, sheet, targets) if targets else ([], [])
    r_block, r_noticed = va._request_gate(wb, sheet, book_before, request, {'normalize': []})
    if plan['extras'] and not plan['tidy'] and r_block:
        # 依頼の中身の検査はシートの主な表（明細・帳票＝元の表）を見る。登録簿のマクロだけの依頼は元の表を触らない決まりで、
        # 書いた表は鍛えたときに正解と突き合わせ済み＝元の表の汚れは AI に回さず気づきに出す
        # （2026-09-17 夜: 月次集計の依頼の「空白」「支出額」で、明細の空白つき・文字の金額を指摘して AI に回していた）
        r_noticed = list(r_noticed) + [f"元の表: {b}" for b in r_block]
        r_block = []
    if plan['extras'] and not plan['tidy'] and (audit or c_block):
        # 登録簿のマクロだけの依頼で、マクロが変えていないセルの指摘（元の表の合計行が文字の金額と合わない・元の式）は
        # AI に回さず気づきに出す（2026-09-17 夜: 値は正解と一致したのに、元の表の「円」付き金額の合計行を理由に AI へ回した）
        audit, c_block, moved = _split_old_findings(audit, c_block, before, ws, sheet)
        c_noticed = list(c_noticed) + [f"元の表: {b}" for b in moved]
    skip = ('formula', 'hidden', 'tables') if dedupe else ()
    # マクロが作り直した出力シート（ピボット・クエリの読み込み先）は、控えから外して比べない＝
    # 「同じ名前のシートがあれば消して作り直す」が仕事なので、前の姿と違って当たり前
    # （2026-09-18 朝: 前の結果が残っている表で「結合セルが全部外れた」と出て AI に回っていた）
    bb = book_before
    outs = _output_sheets(wb, sheet) if book_before is not None else set()
    if outs and isinstance(bb, dict) and bb.get('sheets'):
        bb = dict(bb, sheets={n: s for n, s in bb['sheets'].items() if n not in outs},
                  order=[n for n in (bb.get('order') or []) if n not in outs])
        print("作り直した出力シートは控えと比べません: " + "・".join(sorted(outs)))
    inv = va._inv_violations(wb, bb, {sheet}, skip) if bb is not None else []
    tool_sec['検査'] = time.time() - t0
    blocks = list(audit) + list(c_block) + list(r_block)
    noticed = list(c_noticed) + list(r_noticed)
    did = [f"マクロ {nm}（{owner}）" for nm in names] + ([f"tidy {' '.join(targets)}"] if targets else [])
    if blocks or inv or other:
        why = []
        if blocks or inv:
            why.append(f"残りの指摘 {len(blocks) + len(inv)} 件")
        if other:
            why.append("依頼にマクロの扱わない仕事がある")
        print(f"先撃ちの後: {'・'.join(why)}＝Sonnet の往復へ回します（控えは先撃ちの前のものを引き継ぎます）")
        note = ("\n--- 道具が先に撃ったマクロ（AI より先に・" + f"{time.time() - t_start:.1f} 秒） ---\n"
                + f"{' → '.join(names)} を実行し、tidy で仕上げました。上の材料は撃った後の姿です"
                + "（書き方の整理" + ("・重複の削除" if dedupe else "") + "は済んでいます。やり直さない）。\n")
        if blocks or inv:
            note += "残っている指摘（ここと、依頼のうちマクロでは扱わない仕事だけを手で直してください）:\n" + "".join(
                f"  - {b}\n" for b in (blocks + list(inv))[:12])
        else:
            note += "書き方の指摘は残っていません。依頼のうちマクロでは扱わない仕事だけを手で進めてください。\n"
        return {'result': None, 'ran': True, 'before': before, 'book_before': book_before, 'inv_skip': skip,
                'did': did, 'touched': targets, 'note': note, 'tool_sec': tool_sec}
    # 合格（AI を呼ばずに終わる）
    changed_rows, fmt_rows = [], []
    try:
        changed_rows = va._changes_of(before, ws, sheet) or []
        fmt_rows = va._format_changes(before.get('format'), ws) or []
    except Exception:
        pass
    path = va._write_changes_file(changed_rows, fmt_rows=fmt_rows) if (changed_rows or fmt_rows) else None
    print(f"変わったセル: {len(changed_rows)} 個" + (f"   明細: {path}" if path else ""))
    for line in va._changes_table(changed_rows)[:14]:
        print(line)
    print("仕上げ検査: 指摘なし（見出し・罫線・列の型・列幅・エラー値・合計・式・依頼の語の列の残り）")
    print("頼んでいない変化: なし")
    report = ("【やったこと】（道具が実行した手から）\n" + "\n".join("・" + d for d in did)
              + "\n\n【できなかったこと】なし")
    if noticed:
        report += "\n\n【気づいたこと】（依頼には無いので触っていません）\n" + "\n".join("・" + n for n in noticed)
    sec = time.time() - t_start
    print("関所の差し戻し: なし（マクロで一発）　＝ 一発合格")
    print("道具側の内訳: " + " / ".join(f"{k} {v:.1f}" for k, v in tool_sec.items())
          + f" 秒（合計 {sum(tool_sec.values()):.1f}・AI 待ち 0.0）")
    print("（保存はしていません。Excelで確認後に保存してください。気に入らなければ保存せずに閉じる／agent --undo で戻す）")
    gates = dict.fromkeys(va._GATE_LABELS, 0)
    r = {'done': True, 'ok': True, 'verify': True, 'turns': 0, 'report': report, 'plan': [], 'dropped': [],
         'usage': {'in': 0, 'out': 0, 'think': 0, 'sent': 0, 'recv': 0, 'sec': 0.0}, 'gates': gates,
         'run_id': run_id, 'audit': [], 'unmet': [], 'graded': False, 'grade_skipped': True, 'inv': [],
         'asks': [], 'formula_rate': None, 'limit': max_turns, 'discovered': [], 'merged': 0,
         'tool_sec': {k: round(v, 2) for k, v in tool_sec.items()}, 'inv_skip': skip, 'inv_own': (sheet,),
         'in_tokens': 0, 'out_tokens': 0, 'macro_first': names}
    try:
        va._after_save(book, sheet, wb, True)
    except Exception:
        pass
    va._runs_record(run_id, 'sheet', book, sheet, request, r, sec, path=va._book_path(wb))
    va.progress_write('done', book=book, sheet=sheet, turn=0, max_turns=max_turns, request=str(request)[:80],
                      note='マクロで合格', sec=sec)
    print("AI の報告:\n" + report)
    return {'result': r}


__all__ = ['plan_of', 'plan_full', 'macro_first', 'registry_of', 'registry_from_text', 'dedupe_registry']
