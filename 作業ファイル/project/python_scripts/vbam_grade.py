# -*- coding: utf-8 -*-
"""vbam_grade.py — vba_manager 分割パート: 仕上げ検査・中身の検査・採点係に渡す現物・覚書・関所の通信簿

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。道具の呼び出し（_run_cmd）と tidy の範囲
（_TIDY_RANGES）は呼ぶ瞬間に vbam_agent から引く。vbam_agent は `from vbam_grade import *` で名前を引き継ぐ。
"""
import os
import re
import json
import time
import contextlib
import hashlib
from vbam_core import (SCRIPT_DIR, _col_letter, job_clock_get, job_clock_set)
from vbam_build import (_extract_json)
from vbam_hands import (_addr, _check_addr, _guess_header_row, _heavy_materials, _rows_of, _text_of)
from vbam_undo import (_change_side, _changes_of, _clip, _sheet_snapshot)

_AGENT_NOTES_FILE = os.path.join(SCRIPT_DIR, '_agent_notes.json')     # ブック・シートごとの覚書（前にこの表で何をしたか）
_AGENT_NOTES_KEEP = 3        # 1 シートにつき残す件数


# ----------------------------------------------------------------
# 仕上げ検査の関所と自己採点（2026-09-04・「値は合ったが見た目がひどい」「頼んだ半分で done」を道具で止める）
#   検査＝道具が実物を見る（AI を使わない・タダ）。採点＝AI が依頼文と画像を照らす（1 往復ぶん課金する）。
#   どちらも 1 回だけ差し戻す（無限に往復させない）。
# ----------------------------------------------------------------
_AUDIT_OPS = ('write_grid', 'write_cells', 'tidy', 'fill', 'normalize', 'copy_range', 'format')
_AUDIT_MIN_ROWS, _AUDIT_MIN_COLS = 3, 2      # これ未満は「表」と見ない（見出し 1 行＋2 行・2 列）
_AUDIT_MAX_NOTES = 12


def _notes_load():
    try:
        with open(_AGENT_NOTES_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _notes_key(book, sheet):
    return f"{book}||{sheet}"


def _request_key(request):
    """依頼文の指紋（空白と記号のゆれを畳んだうえで短く。純 Python）。"""
    s = ''.join(str(request or '').split())
    return hashlib.sha1(s.encode('utf-8')).hexdigest()[:12]


def _sheet_fingerprint(ws):
    """シートの姿の指紋（使用範囲・見出し行の並び・数式セル数）。COM は 3 回。読めなければ空。

    2026-09-06・Claude for Excel の助言。「同じ依頼が 2 回来た」を見分ける仕組みがこちらに無く、
    列がもう 1 本増える・行が二重に並ぶ事故を止められなかった。向こうは会話履歴でしか判らない。
    """
    fp = {}
    try:
        ur = ws.UsedRange
        fp['used'] = _addr(ur)
        head = _rows_of(ur.Rows(1).Value)
        fp['head'] = " ".join(_text_of(v) for v in (head[0] if head else []))[:200]
    except Exception:
        return {}
    with contextlib.suppress(Exception):
        fp['formulas'] = int(ws.UsedRange.SpecialCells(-4123).Count)      # xlCellTypeFormulas
    return fp


def _notes_recall(book, sheet, request=None, now_fp=None):
    """このシートで前にやったこと（材料に足す文）。無ければ空。

    同じ表を翌日また触るとき、前回どう決めたかを AI に渡す（2026-09-04・「会話が続かない」）。
    人が後から手で直しているかもしれないので、材料と食い違ったら材料を優先させる。
    2026-09-06: 同じ依頼を、前の走行の直後と同じ姿のシートに撃とうとしていたら先に告げる
    （二重に足す事故＝列がもう 1 本増える・行が二重に並ぶ、を実行前に止める）。
    """
    items = _notes_load().get(_notes_key(book, sheet)) or []
    if not items:
        return ""
    out = [f"【このシートで前にやったこと】（道具の覚書・{len(items)} 件。過去の話。人が後から手で直したり元に戻したりして"
           "いるかもしれないので、いまの材料と食い違ったら材料を優先する＝上の格子が今の現物。read で確かめ直しても同じ格子が返る）"]
    for it in items[-_AGENT_NOTES_KEEP:]:
        out.append(f"  {it.get('time', '?')}  依頼「{it.get('request', '')}」"
                   + (f" → {it.get('report', '')}" if it.get('report') else ""))
    if request and now_fp:
        key = _request_key(request)
        same = [it for it in items if it.get('req_key') == key and it.get('after_fp') == now_fp]
        if same:
            out.append("  ⚠ **同じ依頼を、この姿のシートに、前にも撃っています**（"
                       + str(same[-1].get('time', '?')) + "）。そのときから使用範囲・見出し・数式の数が"
                       "変わっていません＝すでに済んでいる可能性があります。"
                       "**もう一度やると二重になります**（列がもう 1 本増える・行が二重に並ぶ）。"
                       "plan の 1 番目に「すでに済んでいないか確かめる」を置き、read で現物を見てから決めてください。"
                       "済んでいたら actions を空にして、その旨を report に書いて done にしてください")
    return chr(10).join(out) + chr(10)


def _notes_remember(book, sheet, request, report, after_fp=None):
    """終わった仕事を 1 行だけ覚える（1 シートにつき新しい方から 3 件）。"""
    d = _notes_load()
    key = _notes_key(book, sheet)
    items = d.get(key) or []
    items.append({'time': time.strftime('%Y-%m-%d %H:%M'), 'request': str(request or '')[:120],
                  'report': ' '.join(str(report or '').split())[:200],
                  'req_key': _request_key(request), 'after_fp': after_fp or {}})
    d[key] = items[-_AGENT_NOTES_KEEP:]
    if len(d) > 200:                                  # 覚書が増えすぎたら古い鍵から捨てる
        for k in list(d)[:len(d) - 200]:
            d.pop(k, None)
    try:
        with open(_AGENT_NOTES_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except OSError as ex:
        print(f"（覚書を書けませんでした: {ex}）")


# ----------------------------------------------------------------
# 走行の終わりの姿（2026-09-06 夜・輸入候補 11）。Claude for Excel は毎ターン user_changes（人が手で直した番地）を
# 受け取る。こちらは無人の往復なので「走行と走行のあいだ」に人が直した所を、次の走行の材料で見せる。
# 保存済みのブック・何か書いた走行だけ残す（実射の使い捨てブックは Book1!実射1 が毎回ぶつかるので残さない）。
# ----------------------------------------------------------------
_AGENT_AFTER_DIR = os.path.join(SCRIPT_DIR, '_agent_after')


def _after_path(book, sheet):
    return os.path.join(_AGENT_AFTER_DIR,
                        hashlib.sha1(_notes_key(book, sheet).encode('utf-8')).hexdigest()[:16] + '.json')


def _after_save(book, sheet, wb, wrote):
    """走行の終わりのシートの姿を残す。書けなくても仕事は止めない。"""
    if not wrote:
        return
    try:
        if not str(wb.Path or ''):
            return
        snap = _sheet_snapshot(wb.Sheets(sheet))
        if not snap or snap.get('too_big') or not snap.get('values'):
            return
        os.makedirs(_AGENT_AFTER_DIR, exist_ok=True)
        rec = {'book': book, 'sheet': sheet, 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
               'addr': snap.get('addr'), 'row': snap.get('row'), 'col': snap.get('col'),
               'values': [[_text_of(v) for v in row] for row in snap['values']],
               'formulas': ([[_text_of(f) for f in row] for row in snap['formulas']]
                            if snap.get('formulas') else None)}
        with open(_after_path(book, sheet), 'w', encoding='utf-8') as f:
            json.dump(rec, f, ensure_ascii=False)
    except Exception:
        pass


def _since_last_run(book, sheet, wb, limit=8):
    """前回の走行の終わりから、道具の手ではなく変わった所（材料に足す文。無ければ空）。"""
    try:
        if not str(wb.Path or ''):
            return ""
        with open(_after_path(book, sheet), encoding='utf-8') as f:
            rec = json.load(f)
        rows = _changes_of(rec, wb.Sheets(sheet), sheet)
    except Exception:
        return ""
    if not rows:
        return ""
    ex = [f"{r['addr']} {_clip(r['before'], 14)!r}→{_clip(_change_side(r, 'after'), 14)!r}" for r in rows[:limit]]
    return (f"【前回の走行（{rec.get('time', '?')}）の終わりから変わった所】（道具の手ではない＝人か別の道具が直した。"
            f"{len(rows)} セル。いまの材料が正＝上の格子が今の現物そのもので、read で読み直しても同じものが返る。"
            f"前回の報告と食い違っていたら材料を信じ、読み直さずに手を返す）\n  "
            + "  ".join(ex) + chr(10))


# 規則文から外した重い段（2026-09-06 夜・査読 (l)）。毎回読ませていた最長の段を、要る依頼のときだけ材料の末尾に足す
_DATAMODEL_NOTE = ("\n--- データモデル・パワークエリの流れ（この依頼に要るので道具が足しました） ---\n"
                   "table create → powerquery add＋load{\"to\":\"model\"} → datamodel measure_add → pivot create{\"model\":true}"
                   "（材料の「データモデル」に載ったテーブル名を使う。クエリ名はテーブル名と別にする）。"
                   "M を自分で書くなら数値・日付の列に Table.TransformColumnTypes を付ける（付けないと列が全部「文字」になり "
                   "SUM のメジャーが無言で消える）。DAX のテーブル名は 'T売上' とシングルクォートで囲む。\n")
_DATAMODEL_WORDS = re.compile(r'データモデル|パワークエリ|Power ?Query|メジャー|DAX|リレーション|powerquery|datamodel', re.IGNORECASE)


def _wants_datamodel_note(request, materials):
    """重い段を足すか（依頼にその語があるか、材料にクエリ・データモデルのテーブル・メジャーが 1 本でも載っているとき・純 Python）。

    材料の行は「パワークエリ: 0 本」「データモデル: テーブル 0 本   メジャー 0 本」の形。0 本なら足さない
    （最初の実射で 0 本の行に反応して、要らない段を毎回足していた）。
    """
    if _DATAMODEL_WORDS.search(str(request or '')):
        return True
    m = str(materials or '')
    if re.search(r'パワークエリ:\s*(?!0 本)\d+ 本', m):
        return True
    dm = re.search(r'データモデル:[^\n]*', m)
    return bool(dm and re.search(r'(テーブル|メジャー)\s*[1-9]\d* 本', dm.group(0)))


def _audit_targets(actions):
    """実行した手から、仕上げ検査を当てる番地（左上のセル）を拾う。"""
    out = []
    for a in (actions or []):
        if not isinstance(a, dict):
            continue
        op = a.get('op')
        if op not in _AUDIT_OPS:
            continue
        if op == 'tidy':
            cands = list(a.get('ranges') or [])
        elif op == 'write_cells':
            cands = list((a.get('cells') or {}).keys())[:1]
        elif op == 'copy_range':
            cands = [a.get('dst')]
        elif op == 'normalize':
            cands = [a.get('to') or a.get('range')]
        else:
            cands = [a.get('range') or a.get('at')]
        for c in cands:
            try:
                out.append(_check_addr(c).split(':')[0])
            except Exception:
                continue
    return out


def _table_region(ws, a):
    """書いた番地を含む表の範囲。CurrentRegion が表題（見出しの真上の 1 セル）まで飲み込むので、材料と同じ
    見出しの推定（_pick_header_row）で見出しの行から始める（2026-09-17: 表題が見出しの真上にある表で、検査が表題を
    見出しと取り「見出しが空の列」「SUM が本文 1 行を入れていない」と出て、マクロで合っている表を AI に回していた）。"""
    try:
        from vbam_edit import _region_of_cell        # tidy と同じ決め方（1 か所に置く）
        return _region_of_cell(ws, ws.Range(a))
    except Exception:
        return ws.Range(a).CurrentRegion


def _audit_now(wb, sheet, targets):
    """書いた範囲を含む表を実物で検査して、指摘の一覧を返す（空なら合格）。"""
    try:
        from vbam_edit import audit_table
    except Exception as ex:
        print(f"（仕上げ検査を読み込めませんでした: {ex}）")
        return []
    try:
        ws = wb.Sheets(sheet)
    except Exception:
        return []
    seen, notes = set(), []
    for a in targets:
        try:
            rng = _table_region(ws, a)
            addr = str(rng.Address).replace('$', '')
        except Exception:
            continue
        if addr in seen:
            continue
        seen.add(addr)
        try:
            if int(rng.Rows.Count) < _AUDIT_MIN_ROWS or int(rng.Columns.Count) < _AUDIT_MIN_COLS:
                continue                       # 表の形をしていない（見出しだけ・1 列の覚書）は見ない
            notes += audit_table(ws, rng)
        except Exception:
            continue
    return notes[:_AUDIT_MAX_NOTES]


def _content_now(wb, sheet, targets):
    """書いた範囲を含む表の中身を実物で検査 → (done を止める指摘, 報告の気づき)。AI を使わない。

    2026-09-06 夜。仕上げ検査（audit_table）は見た目しか見ておらず、初実射で C10 の空欄を道具も AI も
    採点係も言わなかった。エラー値・合計の不一致・式の断ち切れは done を止める。空欄・文字の数字・
    番号列の重複は気づき（報告に必ず出す。直させない＝値を作らせない）。
    """
    try:
        from vbam_edit import audit_content
        ws = wb.Sheets(sheet)
    except Exception:
        return [], []
    seen, block, noticed = set(), [], []
    for a in targets:
        try:
            rng = _table_region(ws, a)
            addr = str(rng.Address).replace('$', '')
        except Exception:
            continue
        if addr in seen:
            continue
        seen.add(addr)
        try:
            if int(rng.Rows.Count) < _AUDIT_MIN_ROWS or int(rng.Columns.Count) < _AUDIT_MIN_COLS:
                continue
            b, n = audit_content(ws, rng)
            block += b
            noticed += n

            # インサイト診断エンジン（vbam_audit）による高精度関所検査（循環参照・集計漏れ・外れ値等）
            try:
                from vbam_audit import check_formula_linter, check_data_cleaner, to_2d_list
                nr = int(rng.Rows.Count)
                nc = int(rng.Columns.Count)
                start_r = int(rng.Row) - 1
                start_c = int(rng.Column) - 1

                val_raw = rng.Value2 if hasattr(rng, 'Value2') else rng.Value
                form_raw = rng.Formula
                form_r1c1_raw = rng.FormulaR1C1

                vals = to_2d_list(val_raw, nr, nc)
                forms = to_2d_list(form_raw, nr, nc)
                forms_r1c1 = to_2d_list(form_r1c1_raw, nr, nc)

                f_issues = check_formula_linter(forms, forms_r1c1, vals, start_r, start_c)
                d_issues = check_data_cleaner(vals, start_r, start_c)

                # 同じ番地を audit_content が既に言っていれば重ねない（数式の目と診断の目が同じ集計漏れを 2 行にしない・2026-09-17）
                said = ' '.join(block + noticed)
                for item in f_issues + d_issues:
                    cell = str(item.get('cell', ''))
                    if cell and re.search(r'(?<![A-Z$])' + re.escape(cell) + r'(?![0-9])', said):
                        continue
                    msg = f"{cell}: {item['msg']}"
                    if item['severity'] == 'critical':
                        if msg not in block:
                            block.append(msg)
                    else:
                        if msg not in noticed:
                            noticed.append(msg)
            except Exception:
                pass
        except Exception:
            continue
    return block[:_AUDIT_MAX_NOTES], noticed[:_AUDIT_MAX_NOTES]


# 採点は「別人」にやらせる（2026-09-06）。前は同じ会話の続きに質問を足していたので、
# 自分の手順と言い分を全部見た同じ相手が採点していた＝追認しやすい。ここでは会話を切り、
# 依頼文・道具が読み直した現物・見た目の画像だけを渡す（誰がどう書いたかは見せない）。
_GRADE_PROMPT = """【採点】あなたは検査係です。人が出した依頼文と、いまの Excel シートの現物（道具が読み直した
材料と、添えた画像）だけを見て、依頼が満たされているかを判定してください。**誰がどうやったかの説明は
渡していません。書いてあることではなく、シートの現物で判定します。**
まだ満たしていない点だけを unmet に挙げます。満たしているなら空の配列にしてください。
依頼文に無いが表を見て気づいたこと（空欄・全角の数字・重複・不揃い など）は noticed に書きます。noticed は
人が読んで決めるためのもので、直させるためのものではありません（往復を使わない）。

元の依頼文:
{request}

返事は JSON だけ:
{{"unmet": ["満たしていない点（短く。何をどうすれば満たすかまで書く）"],
  "noticed": ["依頼には無いが気づいたこと（番地つきで短く。無ければ空の配列）"], "note": "一言"}}

採点の決まり:
 - 依頼文に書かれていないことを unmet に足さない（思いついた改善・気づいた乱れは noticed に書く）。
   材料の「表の中の空欄」「数値に見える文字」に出た番地は、依頼文がそれを頼んでいなければ noticed。
 - **元からそうだったもの**（隠れている行や列・絞り込み・結合・元の書式・元の表の見た目）を「戻せ」「直せ」と
   言わない。依頼文にその指示が無ければ、それは人がそうしておいた状態であって、直す対象ではない
   （隠れ行の再表示を求めて、頼んでいない変化を作らせたことがある）。
 - 道具に手が無いこと（名前定義やリンクを消す・元に戻す）は unmet にしない（report に書けば足りる）。
 - **印刷（範囲・タイトル行・ページに収める・向き）は材料の「印刷:」の行が現物です。** そこに
   「FitToPage=ON・FitToPagesWide=1」と出ていれば、横 1 ページに収める設定は満たされています
   （「確認できる情報が材料にない」を理由に unmet へ挙げないでください・2026-09-08）。
   **印刷範囲の絞り込みは、依頼文に「どこからどこまで」と書かれていない限り unmet にしない**
   （タイトルや説明行を範囲に含めるかは人の好み＝気になるなら noticed に書く）。
 - **マクロの割り当てられた図形（ボタン）は道具が消しません**（安全側の設計＝人の仕掛けを壊さない）。
   「要らない図形を消して」の依頼でも、マクロ付きのボタンが残っているのは**正しい姿**です。
   これを unmet に挙げないでください（2026-09-08・残った 3 個のボタンを「消せ」と差し戻して往復を 1 回捨てた）。
   【このシートの構造】の図形の行に、どれがマクロ付きかが出ています。
 - 次の往復で直せるものだけを挙げる（挙げた分は必ず直す）。
 - 疑わしいだけで、渡された材料から確かめられない点は挙げない。画像は既定では添えない。罫線・見出しの体裁・列幅・
   ### ・**表示形式（桁区切り）** は【見た目（道具が書式を読んだ現物）】の欄が現物＝そこに「格子あり」
   「### なし」「#,##0」と出ていれば満たされている
   （「確認できない」を理由に unmet へ挙げて、要らない往復を 1 回作ったことがある・2026-09-08）。
 - **「書式が読めないので確認できない」を理由に noticed へ逃がさない。** 依頼された項目は、【見た目】の欄で
   判定できます。そこに出ている表示形式が依頼と食い違うなら unmet に挙げてください
   （2026-09-08・カンマの消えた合計セルを「確認不能」と書きながら合格を出した）。
 - **【このシートの構造】の欄も現物です。** 条件付き書式・入力規則・枠固定・絞り込みは、セルの値にも
   画像にも写らないことがある（12 万行の先頭に重複が無ければ色は見えない）。その欄に出ていれば設定されています。
 - **【このループで作った別のシート】の欄も現物です。**「別シートに作って」の依頼は、その欄に出ていれば
   満たされています（添えた画像は元のシートのものなので、画像に写っていないことを理由に「作られていない」と
   言わないでください）。
 - **依頼文に「整えて」「仕上げて」「体裁」があるなら、見た目も判定に入れる**（2026-09-06・同じ依頼を
   2 回撃って、1 回目だけ金額にカンマが付き、2 回とも「満たしていない点なし」で通った）。
   金額・単価・合計とわかる列に桁区切りが無い／番号列が右寄せ／### で読めない／罫線が表の一部にしか
   無い／見出しだけ書式が違う、は unmet に挙げる（番地か列文字を添える）。

【いまのシート】（道具が読み直した現物）
{materials}
"""


_GRADE_MATERIALS_LIMIT = 12000     # 採点に渡す現物の上限（字）。長い表は頭だけで足りる
# 重い物（テーブル・ピボット・スライサー・クエリ・モデル）が 1 つでもあるか。0 本だけの回は採点に足さない
_HEAVY_NONZERO_RE = re.compile(r'[1-9]\d*\s*(?:本|個)')


def _other_sheets_now(wb, main, before_names):
    """このループで新しくできたシート名（採点に見せるため）。読めなければ空。

    2026-09-06 深夜の実射: 「別シートに月別集計を作って」の弾で、AI は正しく別シートに作ったのに
    採点係には元のシートの材料しか渡していなかったため「別シートが作成されていません」と
    差し戻していた（月別・上位・推移・単価の 4 弾で再現。自己採点の差し戻し 7 件中 6 件がこれ）。
    """
    try:
        now = [str(ws.Name) for ws in wb.Sheets]
    except Exception:
        return []
    return [n for n in now if n not in (before_names or ()) and n != main]


def _grade_extra_sheets(wb, main, before_names, wrote):
    """採点係に現物を見せる別シート＝この走行で作った分＋手が名指しした分。

    2026-09-06 深夜: 「作った分」だけでは足りなかった。実射は 15 弾を 1 冊で回すので、
    2 回目・3 回目の走行では「月別」が前の走行から残っていて「新しくできたシート」に入らず、
    採点係はまた元のシートだけを見て「別シートが作成されていません」と差し戻していた。
    """
    out = []
    for n in list(_other_sheets_now(wb, main, before_names)) + list(wrote or []):
        n = str(n).strip()
        if n and n != main and n not in out:
            out.append(n)
    return out


_CF_TYPE_NAMES = {1: 'セルの値', 2: '数式', 3: 'カラースケール', 4: 'データバー', 5: '上位/下位', 6: 'アイコン',
                  8: '重複/一意', 9: '空白', 10: '空白でない', 11: 'エラー', 12: 'エラーでない',
                  16: '平均より上/下', 17: '日付', 18: '文字を含む'}


_BORDER_KINDS = ((7, '左'), (8, '上'), (9, '下'), (10, '右'), (11, '縦の内側'), (12, '横の内側'))   # xlEdge*・xlInside*


def _pivot_covering(ws, rng):
    """範囲に重なるピボット（無ければ None）。番地の当たり判定だけ＝COM の Intersect を使わない（2026-09-09）。

    ピボットは見出し行・ラベル行と本文で表示形式が違って当たり前。列ごとに読むと NumberFormat が None
    （＝まちまち）になり、採点係が「桁区切りがまちまち」と誤って指摘する（実射で確認）。
    """
    try:
        r1, c1 = int(rng.Row), int(rng.Column)
        r2, c2 = r1 + int(rng.Rows.Count) - 1, c1 + int(rng.Columns.Count) - 1
        for pt in ws.PivotTables():
            try:
                tr = pt.TableRange2
                pr1, pc1 = int(tr.Row), int(tr.Column)
                pr2, pc2 = pr1 + int(tr.Rows.Count) - 1, pc1 + int(tr.Columns.Count) - 1
                if not (r2 < pr1 or r1 > pr2 or c2 < pc1 or c1 > pc2):
                    return pt
            except Exception:
                continue
    except Exception:
        pass
    return None


def _grade_looks(ws):
    """採点係に渡す「見た目」（罫線・見出しの体裁・列幅）。見出し行から広がる表 1 つを COM 10 回ほどで読む。

    画像を添えない既定にしたら（2026-09-08）、採点係が「格子罫線が引かれているか確認できない」を unmet に挙げて
    要らない往復（sonnet 17.5 秒）を作った。仕上げ検査は罫線を見ているのに、その事実を採点係に渡していなかった。
    Range.Borders(種類).LineStyle は範囲で揃っていればその値、まちまちなら None＝6 種類とも揃って線があれば格子。
    """
    out = []
    try:
        got = _guess_header_row(ws)
        if not got:
            return out
        rng = ws.Cells(got[0], int(ws.UsedRange.Column)).CurrentRegion    # ws.Range(Cells) の 1 引数形は型付き COM で落ちる（9/8 実測）
        addr = str(rng.Address).replace('$', '')
        missing = []
        for kind, name in _BORDER_KINDS:
            try:
                ls = rng.Borders(kind).LineStyle
            except Exception:
                ls = None
            if ls is None or int(ls) == -4142:
                missing.append(name)
        out.append(f"罫線: {addr} に格子あり（外枠と内側の縦横すべて）" if not missing
                   else f"罫線: {addr} は格子になっていない（無い・まちまち: {'・'.join(missing)}）")
        c1 = int(rng.Column)
        hdr = ws.Range(ws.Cells(got[0], c1), ws.Cells(got[0], c1 + int(rng.Columns.Count) - 1))
        bold, fill = hdr.Font.Bold, hdr.Interior.ColorIndex
        out.append(f"見出し行 {str(hdr.Address).replace('$', '')}: "
                   + ("太字" if bold is True else "太字でない" if bold is False else "太字まちまち") + "・"
                   + ("塗りまちまち" if fill is None else "塗りなし" if int(fill) == -4142 else "塗りあり"))
        widths = [float(rng.Columns(i + 1).ColumnWidth or 0) for i in range(int(rng.Columns.Count))]
        out.append(f"列幅: {min(widths):.1f}〜{max(widths):.1f}（### は材料に「'###' で読めないセル」の行が無ければ 0）")
        # 表示形式（2026-09-08）。カンマの有無は材料の値には写らない＝採点係が「書式の現物が無く確認不能」と
        # 書いて noticed に逃がしていた（G27 の 203250 がカンマを失っていたのに合格が出た日）。
        # ここに出せば「桁区切りが無い」を unmet として言い切れる。
        n_rows, n_cols = int(rng.Rows.Count), int(rng.Columns.Count)
        pt_here = _pivot_covering(ws, rng)
        if pt_here is not None:
            # ピボットは列ごとに読むと必ず「まちまち」になる（ラベル行が本文と違う書式）。本文の現物を渡す
            try:
                db = pt_here.DataBodyRange
                f = db.NumberFormat
                shown = str(db.Cells(1, 1).Text or '')
                out.append(f"ここはピボット「{pt_here.Name}」です"
                           "（見出し行・ラベル行と本文で表示形式が違うのは正常＝まちまちではありません）。"
                           f"値の表示形式（本文の現物）: {'まちまち' if f is None else f}"
                           + (f"（例 {shown}）" if shown else ""))
            except Exception:
                pass
            out += _grade_outside_numbers(ws, rng)
            return out
        fmts = []
        for i in range(n_cols):
            cl = c1 + i
            head = str(ws.Cells(got[0], cl).Value or '').strip() or _col_letter(cl)
            try:
                body = ws.Range(ws.Cells(got[0] + 1, cl), ws.Cells(got[0] + n_rows - 1, cl))
                f = body.NumberFormat
                shown = str(ws.Cells(got[0] + 1, cl).Text or '')
            except Exception:
                continue
            fmts.append(f"{head}={'まちまち' if f is None else f}"
                        + (f"（例 {shown}）" if shown else ""))
        if fmts:
            out.append("表示形式（見出しの下・列ごと。カンマの有無はここが現物）: " + "／".join(fmts))
        odd = []
        for i in range(n_cols):
            cl = c1 + i
            try:
                hc = ws.Cells(got[0], cl)
                hv = hc.Value
                # 日付も入れる（2026-09-09）。見出しそのものが日付の表（G1:I1 = 2026年1月…）で、
                # ここが数値だけだったため採点係に現物が渡らず「表示形式を yyyy年m月 にしてください」を
                # 9 状態中 7 状態で差し戻していた（そのたび往復 1 回）。
                if isinstance(hv, bool) or not (isinstance(hv, (int, float)) or hasattr(hv, 'year')):
                    continue
                odd.append(f"{str(hc.Address).replace('$', '')}={hc.Text}（{hc.NumberFormat}）")
            except Exception:
                continue
        if odd:
            out.append("先頭行に入っている数値・日付（見出しが日付の表・合計ブロック等。"
                       "表示形式はここが現物）: " + "／".join(odd))
        out += _grade_outside_numbers(ws, rng)
    except Exception:
        pass
    return out


_GRADE_OUTSIDE_MAX = 12          # 表の外の数値セルを採点係に見せる数
_GRADE_OUTSIDE_CELLS = 4000      # 走査するセル数の上限（大きいシートで時間を食わない）


def _grade_outside_numbers(ws, rng):
    """表の外にある数値セル（離れた集計ブロック・注記の下の計算）を、表示形式と画面の文字つきで。

    2026-09-08: 表の中の列だけ渡していたら、採点係が「G27（合計）・G28（平均）の表示形式が確認できない」と
    書いて noticed に逃がした。集計は表の外に置かれることが多い＝そこも現物として渡す。
    """
    out = []
    try:
        used = ws.UsedRange
        top, left = int(used.Row), int(used.Column)
        nr, nc = int(used.Rows.Count), int(used.Columns.Count)
        if nr * nc > _GRADE_OUTSIDE_CELLS:
            return out
        r1, c1 = int(rng.Row), int(rng.Column)
        r2 = r1 + int(rng.Rows.Count) - 1
        c2 = c1 + int(rng.Columns.Count) - 1
    except Exception:
        return out
    found = []
    for r in range(top, top + nr):
        for c in range(left, left + nc):
            if r1 <= r <= r2 and c1 <= c <= c2:
                continue
            try:
                cell = ws.Cells(r, c)
                v = cell.Value
                if isinstance(v, bool) or not isinstance(v, (int, float)):
                    continue
                found.append(f"{str(cell.Address).replace('$', '')}={cell.Text}（{cell.NumberFormat}）")
            except Exception:
                continue
            if len(found) >= _GRADE_OUTSIDE_MAX:
                break
        if len(found) >= _GRADE_OUTSIDE_MAX:
            break
    if found:
        out.append("表の外にある数値（離れた集計ブロック等。画面の文字と表示形式）: " + "／".join(found))
    return out


def _grade_structure(ws):
    """採点係に渡す「セルに写らない構造」（条件付き書式・入力規則・枠固定・絞り込み・グラフ・図形）。

    読めない項目は黙って飛ばす。長くても数百字。
    """
    out = []
    try:
        n = int(ws.Cells.FormatConditions.Count)
        parts = []
        for i in range(1, min(n, 8) + 1):
            try:
                fc = ws.Cells.FormatConditions.Item(i)
                addr = str(fc.AppliesTo.Address).replace('$', '')
                kind = _CF_TYPE_NAMES.get(int(fc.Type), f'種類{int(fc.Type)}')
                extra = ''
                if int(fc.Type) == 8:
                    try:
                        extra = '＝重複に色' if int(fc.DupeUnique) == 1 else '＝一意に色'
                    except Exception:
                        pass
                elif int(fc.Type) in (1, 2):
                    try:
                        extra = ' ' + str(fc.Formula1)[:40]
                    except Exception:
                        pass
                parts.append(f"{addr}（{kind}{extra}）")
            except Exception:
                continue
        out.append(f"条件付き書式: {n} 本" + ("　" + " / ".join(parts) if parts else "")
                   + ("　…" if n > 8 else ""))
    except Exception:
        pass
    try:
        try:
            n_dv = int(ws.Cells.SpecialCells(-4174).Count)
        except Exception:
            n_dv = 0
        out.append(f"入力規則: {n_dv} セル")
    except Exception:
        pass
    try:
        win = ws.Parent.Application.ActiveWindow
        if win is not None and bool(win.FreezePanes):
            out.append(f"枠固定: 上 {int(win.SplitRow)} 行・左 {int(win.SplitColumn)} 列")
    except Exception:
        pass
    try:
        if bool(ws.AutoFilterMode):
            rng = str(ws.AutoFilter.Range.Address).replace('$', '')
            on = 0
            try:
                fl = ws.AutoFilter.Filters
                on = sum(1 for i in range(1, fl.Count + 1) if fl.Item(i).On)
            except Exception:
                pass
            out.append(f"オートフィルタ: {rng}" + (f"（絞り込み中 {on} 列）" if on else "（絞り込みなし）"))
    except Exception:
        pass
    try:
        n_ch = int(ws.ChartObjects().Count)
        if n_ch:
            out.append(f"グラフ: {n_ch} 個")
    except Exception:
        pass
    try:
        n_sh = int(ws.Shapes.Count)
        if n_sh:
            # マクロ付きかどうかまで出す（2026-09-08）。数だけ渡していたら、採点係が「図形が 3 個
            # 残っている＝『要らない図形を消して』が満たされていない」と unmet に挙げ、往復を 1 回
            # 捨てた。道具はマクロ付きの図形を消さない（安全側の設計）＝「残っているのが正しい」
            withmacro, plain = [], []
            for sh in ws.Shapes:
                try:
                    name = str(sh.Name)
                    act = str(sh.OnAction or '')
                except Exception:
                    continue
                (withmacro if act else plain).append(name)
            line = f"図形・ボタン: {n_sh} 個"
            if withmacro:
                line += ("　マクロ付き（道具は消さない＝残っていて正しい）: "
                         + "・".join(withmacro[:8]))
            if plain:
                line += "　マクロなし: " + "・".join(plain[:8])
            out.append(line)
    except Exception:
        pass
    return out


def _looks_of_range(ws, addr):
    """範囲 1 つの見た目（罫線・見出しの体裁）を 1 行で。読めなければ空文字（2026-09-09）。"""
    try:
        rng = ws.Range(addr)
        missing = []
        for kind, name in _BORDER_KINDS:
            try:
                ls = rng.Borders(kind).LineStyle
            except Exception:
                ls = None
            if ls is None or int(ls) == -4142:
                missing.append(name)
        hdr = rng.Rows(1)
        bold, fill = hdr.Font.Bold, hdr.Interior.ColorIndex
        return (f"{addr}: " + ("格子あり（外枠と内側の縦横すべて）" if not missing
                               else f"格子になっていない（無い・まちまち: {'・'.join(missing)}）")
                + "／見出し行 "
                + ("太字" if bold is True else "太字でない" if bold is False else "太字まちまち")
                + "・" + ("塗りまちまち" if fill is None else "塗りなし" if int(fill) == -4142 else "塗りあり"))
    except Exception:
        return ''


def _grade_materials(sheet, wb, extra=()):
    """採点係に渡す「いまのシート」（道具が読み直した現物）。読めなければ空文字。

    仕事の時計は退避して戻す（materials は時計を押す＝ここで押し直すと「経過」が嘘になる。
    build が前の仕事の時計を読んで 1030 秒と出した件と同じ型・2026-09-05）。
    """
    from vbam_agent import (_TIDY_RANGES, _run_cmd)   # 分割後の遅延 import（循環にしない・2026-09-11）
    if wb is None:
        return "（材料を読めませんでした）"
    clock = job_clock_get()
    try:
        ok, mat = _run_cmd(['materials', sheet, '--full'], wb)
    except Exception as ex:
        return f"（材料を読めませんでした: {ex}）"
    finally:
        job_clock_set(clock)
    if not ok:
        return "（材料を読めませんでした）"
    mat = mat.rstrip()
    try:                                             # セルに写らない構造（条件付き書式・入力規則・枠固定・絞り込み）
        st = _grade_structure(wb.Sheets(sheet))
        if st:
            mat += chr(10) + "【このシートの構造（セルには写らない現物）】" + chr(10) + chr(10).join(st)
    except Exception:
        pass
    try:                                             # 見た目（罫線・見出し・列幅）。画像を添えない既定で採点係が罫線を疑わないように（9/8）
        lk = _grade_looks(wb.Sheets(sheet))
        if lk:
            mat += chr(10) + "【見た目（道具が書式を読んだ現物）】" + chr(10) + chr(10).join(lk)
    except Exception:
        pass
    try:                                             # この走行で tidy を当てた範囲（上の【見た目】は表を 1 つしか読まない・9/9）
        lines = [s for s in (_looks_of_range(wb.Sheets(sheet), a)
                             for a in (_TIDY_RANGES.get(sheet) or [])[:4]) if s]
        if lines:
            mat += (chr(10) + "【この走行で仕上げた範囲（道具が当てた tidy の現物）】" + chr(10)
                    + chr(10).join(lines))
    except Exception:
        pass
    try:                                             # 重い物の現物（ピボット・テーブル・スライサー・クエリ・モデル）。
        # 2026-09-09: 採点係は materials（そのシートのセル）しか渡されておらず、正しく作ったピボットを
        # 「値のみ貼り付けの可能性あり」と疑って done を突き返していた（実射で往復が 3 回増えた）。
        # ピボットの現物はセルに写らない＝作業側と同じ一覧（名前・置き場所・行/列/値）を渡す。
        heavy = [s for s in (_heavy_materials(wb) or []) if s]
        if heavy and _HEAVY_NONZERO_RE.search(chr(10).join(heavy)):
            mat += (chr(10) + "【ブックにある重い物（道具が COM で数えた現物）】" + chr(10)
                    + chr(10).join(heavy)
                    + chr(10) + "※ ここに名前が出ていれば、それは本物のピボット／テーブルです"
                              "（値の貼り付けではありません）。名前・置き場所・行/列/値はこの一覧で確かめてください。")
    except Exception:
        pass
    for name in list(extra or ())[:3]:               # このループで作った別シート（採点が「作られていない」と誤るのを防ぐ）
        clock = job_clock_get()
        try:
            ok2, mat2 = _run_cmd(['materials', name], wb)
        except Exception:
            ok2, mat2 = False, ''
        finally:
            job_clock_set(clock)
        if ok2:
            body = mat2.rstrip()
            if len(body) > 4000:
                body = body[:4000] + chr(10) + "…（以下略）"
            mat += chr(10) * 2 + f"【このループで作った別のシート「{name}」の現物】" + chr(10) + body
    return mat          # 全文（採点に渡す前に _clip_grade で切る。全文は終わりの検査が使い回す・2026-09-06 夜）


def _clip_grade(mat):
    """採点係に渡す長さに切る（純 Python）。"""
    mat = (mat or '').rstrip()
    return mat if len(mat) <= _GRADE_MATERIALS_LIMIT else mat[:_GRADE_MATERIALS_LIMIT] + "\n…（以下略）"


# 採点の unmet のうち「人の値を変えないと満たせない」もの（2026-09-06 夜）。承認の言葉が無い走行では、
# これを plan に混ぜても承認の門で止まって往復を失うだけ（シュウさんの初実射＝往復 5 回の型）。
# 道具が noticed（気づき）へ回して、往復を使わず報告に写す。AI を使わない＝費用ゼロ。
_UNMET_VALUE_CHANGE_RE = re.compile(r'削除|消し|消す|置き換え|置換|上書き|書き換え|半角に|全角に|埋め|修正|入れ替え|統一')


_UNMET_PRINT_RE = re.compile(r'FitToPage|ページに収める|横\s*1\s*ページ|Landscape|用紙の向き|横向き')
_UNMET_PHOTO_RE = re.compile(r'(写真|画像)[^。]{0,24}(挿入|入れて|配置|貼|置いて)')
_REQ_PATH_RE = re.compile(r'[A-Za-z]:[\\/]|\.(?:jpg|jpeg|png|gif|bmp)\b', re.I)


def _page_setup_now(wb, sheet):
    """印刷設定の現物（COM）。読めなければ None。FitToPage が ON のとき Zoom は False を返す。"""
    try:
        p = wb.Sheets(sheet).PageSetup
        zoom, wide, tall = p.Zoom, p.FitToPagesWide, p.FitToPagesTall
        land = int(p.Orientation) == 2
        return {'fit': (zoom is False) and int(wide or 0) == 1, 'landscape': land,
                'text': (f"向き {'横' if land else '縦'}・FitToPage={'ON' if zoom is False else 'OFF'}"
                         f"・FitToPagesWide={wide}・FitToPagesTall={tall}")}
    except Exception:
        return None


def _unmet_drop_satisfied(unmet, wb, sheet, request):
    """採点の unmet のうち、現物と食い違うものを落とす → (残り, 落とした理由つき)。

    2026-09-09 の九状態×全弾: 差し戻し 39 件のうち 18 件がこの 2 つだった。
      - 印刷: 材料に「FitToPage=ON・FitToPagesWide=1」と出ているのに「横 1 ページに収めてください」
        （9 状態すべてで再現。採点の規約文には 9/8 に書き足してあるのに消えない＝文では消えない）
      - 写真: パスの無い依頼で「写真を挿入してください」（9 状態すべて）。道具はパスが無ければ
        置けない＝この弾の正解は「要ると報告する」。挙げられた分だけ往復を捨てる
    規約文で頼まず、道具が現物を読んで落とす（AI と Python の役割分担）。
    """
    keep, dropped, ps, dups = [], [], None, None
    for u in (unmet or []):
        s = str(u)
        if _UNMET_PRINT_RE.search(s):
            if ps is None:
                ps = _page_setup_now(wb, sheet) or {}
            if ps.get('fit') or ps.get('landscape'):
                dropped.append(f"{s}　← 現物: {ps.get('text')}")
                continue
        if _UNMET_PHOTO_RE.search(s) and not _REQ_PATH_RE.search(str(request or '')):
            dropped.append(f"{s}　← 依頼に画像ファイルのパスが無い＝道具は写真を置けない（報告で聞くのが正解）")
            continue
        if _UNMET_DUP_RE.search(s):
            # 採点係は名前だけ見て「同じ会社の別の担当者」を重複と言う（2026-09-11 夜・試験: 差し戻しで row_delete を
            # 撃たせ、道具が止め、2 往復を捨てた）。重複かは道具の照合（表記ゆれ込み・担当・電話・日付も見る）が決める
            if dups is None:
                dups = _dup_count_now(wb, sheet)
            if dups == 0:
                dropped.append(f"{s}　← 現物: 道具の照合で重複は 0 組（名前が同じでも担当・電話・日付などが違う行は別の相手）")
                continue
        keep.append(u)
    return keep, dropped


_UNMET_DUP_RE = re.compile(r'重複')


def _dup_count_now(wb, sheet):
    """いまのシートの重複行の組の数（道具の照合）。読めなければ None（COM 1 回＋純 Python）。"""
    try:
        from vbam_view import _guess_header_idx
        from vbam_hands import _dup_pairs
        rows = [list(r) for r in _rows_of(wb.Sheets(sheet).UsedRange.Value)]
        return len(_dup_pairs(rows, _guess_header_idx(rows)))
    except Exception:
        return None


def _split_unmet_for_approval(unmet, approved):
    """(plan に回す unmet, 気づきに回す unmet)。承認があれば全部 plan へ（純 Python）。"""
    if approved:
        return list(unmet or []), []
    keep, moved = [], []
    for u in (unmet or []):
        (moved if _UNMET_VALUE_CHANGE_RE.search(str(u)) else keep).append(u)
    return keep, moved


_SECRET_NAME_RE = re.compile(r'key|token|secret|password|passwd|credential|パスワード|暗証|認証', re.I)
_SECRET_MASK = '＊＊＊（伏せました）'


def _secret_values(wb):
    """名前定義が KEY／TOKEN／PASSWORD などを指すセルの値（8 字以上の文字）を集める（COM）。
    材料の名前定義に「GEMINI_API_KEY = =目次!$A$13」が出て、AI は read_sheet でその値を読めた
    （読めば鍵が API に渡る・2026-09-06 夜）。読めなければ空。"""
    out = set()
    try:
        names = list(wb.Names)
    except Exception:
        return out
    for n in names:
        try:
            if not _SECRET_NAME_RE.search(str(n.Name)):
                continue
            v = n.RefersToRange.Value
        except Exception:
            continue
        vals = [v] if not isinstance(v, tuple) else [x for row in v for x in (row if isinstance(row, tuple) else (row,))]
        for x in vals:
            if isinstance(x, str) and len(x.strip()) >= 8:
                out.add(x.strip())
    return out


def _mask_secrets(text, secrets):
    """AI に送る文から鍵の値を伏せる（純 Python）。"""
    s = str(text or '')
    for v in sorted(secrets or (), key=len, reverse=True):
        s = s.replace(v, _SECRET_MASK)
    return s


def _parse_unmet(text):
    """採点の返事 → 満たしていない点の一覧。読めなければ空（採点は関所であって、落とし穴にしない）。"""
    try:
        d = _extract_json(text)
    except Exception:
        return []
    if not isinstance(d, dict):
        return []
    out = []
    for e in (d.get('unmet') or []):
        t = str(e).strip() if not isinstance(e, dict) else str(e.get('item') or '').strip()
        if t:
            out.append(t[:120])
    return out[:8]


def _parse_noticed(text):
    """採点の返事 → 依頼には無いが気づいたこと（2026-09-06 夕）。unmet と違って往復は使わず、報告に写すだけ。

    シュウさんの初実射: 「整えて」に採点係が A17 の全角と 16 行目の重複を unmet に入れ、AI が直しに行って
    承認の門で止まり、往復 5 回。頼んでいないことは unmet でなく noticed＝人が読んで agent --cont で決める。
    """
    try:
        d = _extract_json(text)
    except Exception:
        return []
    if not isinstance(d, dict):
        return []
    out = []
    for e in (d.get('noticed') or []):
        t = str(e).strip() if not isinstance(e, dict) else str(e.get('item') or '').strip()
        if t:
            out.append(t[:120])
    return out[:8]


# ----------------------------------------------------------------
# 関所の通信簿（2026-09-06）: どの関所で何回差し戻したかを数える。
#   「59/59 合格」だけでは、AI が一発で正しかったのか・関所が毎回直したのかが分からない。
#   点数の中身を「一発合格／関所で直して合格／不合格」の 3 段に割るための材料。
# ----------------------------------------------------------------
_GATE_LABELS = {'format': '返事の形', 'hand': '手の失敗', 'plan': '未の項目', 'audit': '仕上げ検査',
                'inv': '頼んでいない変化', 'coverage': '読み残し', 'grade': '自己採点', 'empty': '空の返事',
                'overwrite': '事前の門（人の値）', 'approval': '承認の言葉', 'verify': '動かして確かめ'}


def _gate_total(gates):
    return sum(int(v) for v in (gates or {}).values())


def _gate_line(gates, head="関所の差し戻し"):
    """差し戻しの内訳を 1 行に。1 回も無ければ「なし（一発）」。"""
    hit = [f"{_GATE_LABELS.get(k, k)} {v}" for k, v in (gates or {}).items() if v]
    return f"{head}: " + (" / ".join(hit) if hit else "なし（一発）")


def _gate_grade(ok, gates):
    """3 段の評価: 一発合格 / 関所で直して合格 / 不合格。"""
    if not ok:
        return '不合格'
    return '一発合格' if _gate_total(gates) == 0 else '関所で直して合格'


_DEDUP_REQ_RE = re.compile(r'重複[^。\n]{0,30}(削除|削|消|除|取り|まとめ|一つに|1つに)|(削除|消)[^。\n]{0,10}重複')


def _validation_misfits(ws):
    """リストの入力規則が付いたセルで、候補に無い値 → 指摘の文の並び（2026-09-11: 「受注済みみ」が合格で出ていった）。"""
    try:
        areas = ws.UsedRange.SpecialCells(-4174)            # xlCellTypeAllValidation
    except Exception:
        return []
    bad, lists = [], []
    for ar in areas.Areas:
        try:
            v = ar.Cells(1, 1).Validation
            if int(v.Type) != 3:                              # xlValidateList
                continue
            f1 = str(v.Formula1 or '')
        except Exception:
            continue
        if not f1 or f1.startswith('='):
            continue                                          # 範囲を参照するリストは見ない
        allowed = {x.strip() for x in re.split(r'[,，]', f1)}
        r1, c1 = int(ar.Row), int(ar.Column)
        for i, row in enumerate(_rows_of(ar.Value)):
            for j, x in enumerate(row):
                if x in (None, ''):
                    continue
                s = x.strip() if isinstance(x, str) else _text_of(x)
                if s not in allowed:
                    bad.append(f"{_col_letter(c1 + j)}{r1 + i} '{str(s)[:20]}'")
                    if f1 not in lists:
                        lists.append(f1)
        if len(bad) > 40:
            break
    if not bad:
        return []
    return [f"入力規則のリスト（{' / '.join(lists)}）に無い値が残っています: " + " ".join(bad[:10])
            + (f" ほか {len(bad) - 10} セル" if len(bad) > 10 else "")
            + "。候補のどれかに直してください（置換の二重当てで「受注済みみ」になる、など）"]


_LOOK_WORDS = (('フォント', 'font'), ('書体', 'font'), ('大きさ', 'size'), ('サイズ', 'size'), ('斜体', 'italic'),
               ('下線', 'underline'), ('取り消し線', 'strike'), ('文字色', 'color'), ('色', 'color'),
               ('塗り', 'fill'), ('背景', 'fill'), ('太字', 'bold'),
               ('体裁', 'align'), ('綺麗', 'align'), ('きれい', 'align'), ('見やす', 'align'), ('そろえ', 'align'),
               ('揃え', 'align'), ('整え', 'align'), ('ばらつき', 'align'), ('寄せ', 'align'))


def _look_misfits(rng, req):
    """依頼に出た見た目の語について、本文の範囲がそろっているか（COM は範囲ごとに数回）→ 指摘の文の並び。"""
    wants = {k for w, k in _LOOK_WORDS if w in req}
    # 「書式のばらつき」「体裁」など全体を指す語は、文字の書式も全部見る（寄せだけ見て、色・大きさ・太字が
    # ばらばらの表に合格を出した・2026-09-11）
    if re.search(r'書式|体裁|綺麗|きれい|見やす|ばらつき', req):
        wants |= {k for _w, k in _LOOK_WORDS}
    if not wants:
        return []
    bad = []
    try:
        f = rng.Font
        if 'font' in wants and f.Name is None:
            bad.append("フォント名")
        if 'size' in wants and f.Size is None:
            bad.append("文字の大きさ")
        if 'italic' in wants and f.Italic is not False:
            bad.append("斜体")
        if 'underline' in wants and f.Underline != -4142:
            bad.append("下線")
        if 'strike' in wants and f.Strikethrough is not False:
            bad.append("取り消し線")
        if 'color' in wants and f.Color is None:
            bad.append("文字色")
        if 'bold' in wants and f.Bold is None:
            bad.append("太字")
        if 'fill' in wants and rng.Interior.ColorIndex is None:
            bad.append("塗りつぶし")
        if 'align' in wants:
            from vbam_core import _col_letter as _cl
            mixed = [_cl(int(rng.Columns(k).Column)) for k in range(1, int(rng.Columns.Count) + 1)
                     if rng.Columns(k).HorizontalAlignment is None]
            if mixed:
                bad.append(f"寄せ（{','.join(mixed)} 列）")
    except Exception:
        return []
    if not bad:
        return []
    addr = str(rng.Address).replace('$', '')
    return [f"依頼の見た目がそろっていません（本文 {addr}）: {'・'.join(bad)} がセルごとに違います。"
            "format で本文全体に font・size・color と \"plain\": true（斜体・下線・取り消し線を外す）・"
            "\"bg\": \"none\"（塗りを消す）・\"unbold\": true を当ててください（条件付き書式の色は消えません）。寄せのばらつきは tidy が列ごとにそろえます（文字は左・数と日付は標準）"]


_MIS_FULL_ALNUM = re.compile(r'[０-９Ａ-Ｚａ-ｚ]')
_MIS_HALF_KANA = re.compile(r'[ｦ-ﾟ]')
_MIS_CORP_ABBR = re.compile(r'\((?:株|有)\)')          # NFKC の後（（株）・㈱ は (株) になる）
_MIS_CORP_FULL = re.compile(r'株式会社|有限会社')
_MIS_PHONE_OK = re.compile(r'0\d{1,4}-\d{1,4}-\d{3,4}|0\d{9,10}')
_MIS_NUM_TEXT = re.compile(r'[¥￥$]?\s*[-+]?[\d０-９][\d０-９,，]*(?:\.\d+)?\s*(?:円|個|本|枚|台|冊|箱|点|件)?')


def _column_misfits(rows, h, cols, r0, c0, req):
    """依頼に出た語の列で、書き方がそろっていない所 → done を止める指摘のリスト（純 Python）。

    2026-09-11 夜・試験（正解の表との突き合わせ）で、郵便の区切りなし・〒・電話の空白区切りとかっこ・全角の英数字・
    （株）と株式会社の混在・文字のままの日付と金額が残った表に、道具が「合格」を出した（5 題中 3 題）。
    材料の気づきと同じことを、終わった表でもう一度数える。見るのは依頼に出た語の分だけ。
    """
    import unicodedata
    from vbam_hands import _is_date_value, _parse_date_text
    out = []
    heads = rows[h] if h is not None and h < len(rows) else []
    ask = {k: bool(re.search(p, req, re.I)) for k, p in (
        ('width', r'全角|半角'), ('space', r'空白|スペース|全角|半角'), ('kana', r'カナ|全角|半角'),
        ('corp', r'株式会社|（株）|\(株\)|表記ゆれ|会社名'), ('phone', r'電話|TEL|携帯|FAX'),
        ('postal', r'郵便|〒'), ('date', r'日'), ('num', r'金額|数値|数量|単価|価格|額|数'))}
    for j in cols:
        head = str(heads[j]).strip() if j < len(heads) and heads[j] is not None else ''
        vals = [(i, rows[i][j]) for i in range(h + 1, len(rows))
                if j < len(rows[i]) and rows[i][j] is not None and rows[i][j] != '']
        if not vals:
            continue
        col = _col_letter(c0 + j)
        name = f"{col} 列（{head}）" if head else f"{col} 列"

        def ex(lst):
            return " ".join(f"{col}{r0 + i} '{str(v)[:16]}'" for i, v in lst[:6]) + (
                f" ほか {len(lst) - 6} セル" if len(lst) > 6 else "")
        strs = [(i, v) for i, v in vals if isinstance(v, str)]
        if ask['phone'] and re.search(r'電話|TEL|携帯|FAX', head, re.I):
            odd = [(i, v) for i, v in vals if not (isinstance(v, str) and _MIS_PHONE_OK.fullmatch(v))]
            hy = [(i, v) for i, v in strs if '-' in v and _MIS_PHONE_OK.fullmatch(v)]
            plain = [(i, v) for i, v in strs if '-' not in v and _MIS_PHONE_OK.fullmatch(v)]
            if odd:
                out.append(f"{name}の電話番号の書き方がそろっていません（空白・かっこ・全角・数値）: {ex(odd)}。"
                           "normalize の phone で直してください（区切りが決まらない番号は write_cells で）")
            elif hy and plain:
                out.append(f"{name}の電話番号で、ハイフンのある・ないが混ざっています（ない方: {ex(plain)}）。"
                           "normalize の phone でそろえてください")
            continue
        if ask['postal'] and re.search(r'郵便|〒', head):
            odd = [(i, v) for i, v in vals if not (isinstance(v, str) and re.fullmatch(r'\d{3}-\d{4}', v))]
            if odd:
                out.append(f"{name}の郵便番号が 000-0000 になっていません: {ex(odd)}。normalize の postal で直してください")
            continue
        if ask['width']:
            fw = [(i, v) for i, v in strs if _MIS_FULL_ALNUM.search(v)]
            if fw:
                out.append(f"{name}に全角の英数字が残っています: {ex(fw)}。normalize の hankaku で直してください")
        if ask['kana']:
            hk = [(i, v) for i, v in strs if _MIS_HALF_KANA.search(v)]
            if hk:
                out.append(f"{name}に半角カナが残っています: {ex(hk)}。normalize の kana_zenkaku で直してください")
        if ask['space']:
            zs = [(i, v) for i, v in strs if re.search(r'(?<=\S)　+(?=\S)', v)]
            hs = [(i, v) for i, v in strs if re.search(r'(?<=\S) +(?=\S)', v)]
            if zs and hs:
                few = zs if len(zs) <= len(hs) else hs
                out.append(f"{name}の文字の間の空白に全角と半角が混ざっています（全角 {len(zs)}・半角 {len(hs)}。"
                           f"少ない方: {ex(few)}）。normalize の space_hankaku か space_zenkaku でそろえてください")
        if ask['corp']:
            ab = [(i, v) for i, v in strs if _MIS_CORP_ABBR.search(unicodedata.normalize('NFKC', v)) or re.search(r'[㈱㈲]', v)]
            fu = [(i, v) for i, v in strs if _MIS_CORP_FULL.search(v)]
            sp = [(i, v) for i, v in fu if re.search(r'(株式会社|有限会社)[\s　]|[\s　](株式会社|有限会社)', v.strip())]
            if ab and fu:
                out.append(f"{name}で（株）と株式会社が混ざっています（略した方 {len(ab)}: {ex(ab)}）。"
                           "normalize の corp でそろえてください")
            if sp and len(sp) < len(fu):
                out.append(f"{name}で会社の種類（株式会社など）の前後に空白のある名前が混ざっています: {ex(sp)}。"
                           "normalize の corp で取ってください")
        if ask['date']:
            real = [1 for _i, v in vals if _is_date_value(v)]
            tx = [(i, v) for i, v in strs if _parse_date_text(v.strip())]
            if real and tx:
                out.append(f"{name}に文字のままの日付が残っています: {ex(tx)}。normalize の date で直してください")
        if ask['num']:
            nums = [1 for _i, v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
            tx = [(i, v) for i, v in strs if _MIS_NUM_TEXT.fullmatch(v.strip()) and not re.match(r'^\s*[0０]\d', v)]
            if nums and tx:
                out.append(f"{name}に数値になっていない数（文字）が残っています: {ex(tx)}。normalize の number で直してください")
    return out


def _request_gate(wb, sheet, before, request, facts):
    """依頼の中身を現物で確かめる → (done を止める指摘, 気づき)。AI を使わない（2026-09-11）。

    動画のお題（重複・日付・金額・ステータスの整理）を撃ったら、仕上げ検査も中身の検査も見た目とエラー値しか
    見ておらず、重複が 7 組残り・7 社が消え・「受注済みみ」・文字の金額が残った表に「合格」を出した。
    """
    block, noticed = [], []
    try:
        ws = wb.Sheets(sheet)
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        rows = [list(r) for r in _rows_of(ur.Value)]
    except Exception:
        return block, noticed
    from vbam_view import _guess_header_idx
    from vbam_hands import _dup_pairs, _lost_rows, _loose_key, _col_num
    h = _guess_header_idx(rows)
    try:
        block += _validation_misfits(ws)
    except Exception:
        pass
    req = str(request or '')
    last_i = max((i for i, row in enumerate(rows) if any(v not in (None, '') for v in row)), default=-1)
    if h is not None and last_i > h:
        hr = rows[h]
        filled = [j for j, v in enumerate(hr) if v not in (None, '')]
        if filled:
            try:
                # 本文は値のある最後の行まで（使用範囲の末尾までにすると、書式だけの空の行を「寄せがばらばら」と数えた）
                body = ws.Range(ws.Cells(r0 + h + 1, c0 + filled[0]), ws.Cells(r0 + last_i, c0 + filled[-1]))
                block += _look_misfits(body, req)
            except Exception:
                pass
            if re.search(r'体裁|綺麗|きれい|整え|そろえ|揃え|掃除|空白|クリーン', req):
                # 値の前後の空白（2026-09-11: 会社名の末尾の空白 3 セルが残ったまま合格した）
                ws_left = []
                for i in range(h + 1, last_i + 1):
                    for j in range(filled[0], filled[-1] + 1):
                        v = rows[i][j] if j < len(rows[i]) else None
                        if isinstance(v, str) and v.strip() and v != v.strip(' \u3000\t\r\n\xa0'):
                            ws_left.append(f"{_col_letter(c0 + j)}{r0 + i} '{v[:16]}'")
                if ws_left:
                    block.append("前後に空白の残った値があります: " + " ".join(ws_left[:10])
                                 + (f" ほか {len(ws_left) - 10} セル" if len(ws_left) > 10 else "")
                                 + "。normalize の trim で消してください")
            try:
                block += _column_misfits(rows[:last_i + 1], h, range(filled[0], filled[-1] + 1), r0, c0, req)
            except Exception:
                pass
    if _DEDUP_REQ_RE.search(req) and len(rows) <= 20000:
        snap = ((before or {}).get('sheets') or {}).get(sheet) if before else None
        if snap and snap.get('values'):
            bvals = [list(r) for r in snap['values']]
            m = re.match(r'^\$?[A-Z]+\$?(\d+)', str(snap.get('used') or 'A1'))
            br0 = int(m.group(1)) if m else 1
            lost = _lost_rows(bvals, rows, _guess_header_idx(bvals), h)
            if lost:
                ex = "／".join(f"行{br0 + i}（{next((str(v) for v in bvals[i] if v not in (None, '')), '')}）"
                              for i in lost[:8])
                block.append("重複を消す依頼なのに、重複ではない行が消えています（残っている行のどれとも同じ相手になりません）: "
                             + ex + (f" ほか {len(lost) - 8} 行" if len(lost) > 8 else "")
                             + "（消す前の行番号）。write で作り直さず、report に書いてください（人が --undo で戻してやり直します）")
        pairs = _dup_pairs(rows, h)
        if pairs:
            hr = rows[h] if h is not None else []
            j0 = next((j for j, v in enumerate(hr) if v not in (None, '')), 0)

            def k0(i):
                return _loose_key(rows[i][j0]) if j0 < len(rows[i]) else ''
            strong = [(i, k) for i, k, _e in pairs if k0(i) and k0(i) == k0(k)]
            weak = [(i, k) for i, k, _e in pairs if (i, k) not in strong]
            if strong:
                block.append("重複が残っています: " + "／".join(
                    f"行{r0 + i} = 行{r0 + k}（{rows[i][j0]}）" for i, k in strong[:10])
                    + (f" ほか {len(strong) - 10} 組" if len(strong) > 10 else "")
                    + "。dedupe の手で消してください（keys＝同じ相手を決める列）")
            if weak:
                noticed.append("名前の書き方が違う重複の疑い（担当者・電話・メールなどが同じ。消すかは人が判断）: "
                               + "／".join(f"行{r0 + i} と 行{r0 + k}" for i, k in weak[:10]))
        # 重複の疑い（全部の列は同じでない＝消さない・2026-09-12・全列一致が基本）。備考だけ違う組などを人に見せる
        try:
            from vbam_hands import _near_dup_pairs
            near = _near_dup_pairs(rows, h)
        except Exception:
            near = []
        if near:
            hr2 = rows[h] if h is not None else []

            def _cn(j):
                v = hr2[j] if j < len(hr2) else None
                return str(v).strip() if v not in (None, '') else _col_letter(c0 + j)
            noticed.append("重複の疑い（全部の列は同じでない＝消していません。どちらを残すか・まとめるかは人の判断）: "
                           + "／".join(f"行{r0 + i} と 行{r0 + k}（{'・'.join(_cn(j) for j in diff) or '書き方'}が違う）"
                                      for i, k, diff in near[:15])
                           + (f" ほか {len(near) - 15} 組" if len(near) > 15 else ""))
    if h is not None:
        seen = set()
        for sh, col, rules in ((facts or {}).get('normalize') or []):
            if sh != sheet or (col, rules) in seen:
                continue
            seen.add((col, rules))
            j = _col_num(col) - c0
            if j < 0:
                continue
            cells = [(r0 + i, rows[i][j]) for i in range(h + 1, len(rows)) if j < len(rows[i])]
            left = []
            if 'date' in rules:
                left += [(r, v) for r, v in cells if isinstance(v, str) and re.search(r'\d', v)]
            if 'number' in rules:
                left += [(r, v) for r, v in cells if isinstance(v, str) and re.search(r'\d', v)
                         and not re.match(r'^\s*[0０]\d', v) and len(re.sub(r'\D', '', v)) <= 15]
            if 'lower' in rules:
                left += [(r, v) for r, v in cells if isinstance(v, str) and v != v.lower()]
            if 'phone' in rules:
                raw = [(r, v) for r, v in cells if isinstance(v, str) and re.fullmatch(r'\s*[0-9０-９+＋]{9,13}\s*', v)]
                if raw:
                    noticed.append(f"{col} 列の電話番号で、市外局番の区切りを道具が決められなかったもの（そのまま）: "
                                   + " ".join(f"{col}{r} '{v}'" for r, v in raw[:8]))
            if left:
                uniq_left = list(dict.fromkeys(left))
                block.append(f"normalize（{'・'.join(rules)}）を当てた {col} 列に、直っていない文字が残っています: "
                             + " ".join(f"{col}{r} '{str(v)[:20]}'" for r, v in uniq_left[:8])
                             + "。規則で読めない形は write_cells で正しい値を書いてください（overwrite）")
    return block, noticed


__all__ = ['_AGENT_AFTER_DIR', '_AGENT_NOTES_FILE', '_AGENT_NOTES_KEEP', '_AUDIT_MAX_NOTES', '_AUDIT_MIN_COLS', '_AUDIT_MIN_ROWS', '_AUDIT_OPS', '_BORDER_KINDS', '_CF_TYPE_NAMES', '_DATAMODEL_NOTE', '_DATAMODEL_WORDS', '_GATE_LABELS', '_GRADE_MATERIALS_LIMIT', '_GRADE_OUTSIDE_CELLS', '_GRADE_OUTSIDE_MAX', '_GRADE_PROMPT', '_HEAVY_NONZERO_RE', '_REQ_PATH_RE', '_SECRET_MASK', '_SECRET_NAME_RE', '_UNMET_PHOTO_RE', '_UNMET_PRINT_RE', '_UNMET_VALUE_CHANGE_RE', '_after_path', '_after_save', '_audit_now', '_audit_targets', '_clip_grade', '_content_now', '_request_gate', '_look_misfits', '_LOOK_WORDS', '_validation_misfits', '_DEDUP_REQ_RE', '_gate_grade', '_gate_line', '_gate_total', '_grade_extra_sheets', '_grade_looks', '_grade_materials', '_grade_outside_numbers', '_grade_structure', '_looks_of_range', '_mask_secrets', '_notes_key', '_notes_load', '_notes_recall', '_notes_remember', '_other_sheets_now', '_page_setup_now', '_parse_noticed', '_parse_unmet', '_pivot_covering', '_request_key', '_secret_values', '_sheet_fingerprint', '_since_last_run', '_split_unmet_for_approval', '_unmet_drop_satisfied', '_wants_datamodel_note']
