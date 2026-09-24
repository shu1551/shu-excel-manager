# -*- coding: utf-8 -*-
"""vbam_macro.py — vba_manager 分割パート: macro モード（マクロの修理・新規作成）と sheet＋macro の統合ループ

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。往復の本体（run_agent・関所・控え）は
vbam_agent.py にあり、こちらからは `va.` で呼ぶ（呼ぶ瞬間に vbam_agent の名前を引く）。
vbam_agent 側はこの module を関数の中で遅延 import する（循環にしない）。
"""
import re
import json
import time
import contextlib
from vbam_vba import (_suggest_similar)
from vbam_build import (_extract_json)

import vbam_agent as va


# ----------------------------------------------------------------
# macro: マクロの修理（修理の型: get・call-graph → 仮説 → 最小の一手 → compile → …）
# ----------------------------------------------------------------

_MACRO_OPS = ('get', 'grep', 'replace', 'code_replace', 'compile', 'rehearse', 'add',
              'list', 'impact', 'call_graph', 'run', 'lines')   # 2026-09-04: 読む手 3 本と run（run は rehearse と同じ関所）／09-17 lines
_PROC_DECL_RE = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+([^\s\(\)]+)',
                           re.IGNORECASE | re.MULTILINE)

MACRO_RULES = """あなたは Excel VBA のマクロを直す／作る係です。道具（Python）がコード・呼び出し関係・コンパイルの結果を集め、
あなたの返事どおりに書き換え、書き換えたら必ず全体コンパイルして結果を次の【結果】で見せます。
返事は JSON だけ（前後の文・コードフェンスは付けない）。

返事の形:
{"say": "いま何をするかを一言",
 "hypothesis": "動かない原因の仮説を一言（1 往復目に必ず書く。新しく作るときは作りの方針を一言）",
 "plan": [{"item": "依頼を分解した 1 つ（短く）", "state": "未"}],
 "actions": [ ... ],
 "done": false,
 "report": "終わったときの報告（done が true のときだけ。何をどう直したか・コンパイルが通ったこと）"}

plan は 1 往復目に必ず書く（依頼を分解した項目を短文で並べる。state は「未」で始める）。
2 往復目からは同じ項目の state だけを「済」か「不可」に変えて毎回返す（項目は増やさない・減らさない）。
**「未」が 1 つでも残っている done は道具が受け付けない**。できない分は「不可」にして report に理由を書く。

actions に並べられる手（この 13 個だけ）:
 {"op":"get","name":"マクロ名"}                        別のプロシージャのコードを読む（"module":"名" で限定できる）。GoSub のラベル名を渡すとそのラベル〜Return の塊を行番号つきで返す
 {"op":"lines","from":650,"to":720,"module":"モジュール名"}   モジュールの行を行番号つきで読む（1 回 160 行まで。行番号は grep の「マクロ名:661」と同じ）。材料の本文が途中で切れていたらこれで続きを読む
 {"op":"grep","text":"文字"}                            全モジュールを横断検索（どこで使われているか）
 {"op":"code_replace","search":"旧の行の一部","replace":"新","module":"モジュール名"}   変更行だけを置換（1〜数行の小修正はこれ）。module は書く（省くと、修理中のマクロのモジュール。それも無ければ search の行があるモジュールが 1 つだけならそこ・複数なら断られる）。そのモジュールの中で search に当たる行は**全部**変わるので、他のプロシージャにもある行（On Error Resume Next・Next i など）を search にしない＝前後を足して一意にするか replace で全文を入れ直す
   **行を足す・抜く・数行まとめて直すときは search と replace に改行（\\n）を入れてよい**（例: 抜けた End If を足す＝search に前後の 2〜3 行、replace にそれへ End If を足した行）。複数行の search はそのモジュールで**並びごと 1 か所だけ**当たること（字下げの違いは道具が吸収する）。道具がそのプロシージャ 1 本の差し替えに組み直す＝プロシージャ全文を replace で送り直さない
 {"op":"replace","name":"マクロ名","code":"Sub マクロ名()\\n …\\nEnd Sub"}   プロシージャ 1 本を丸ごと差し替え（"module":"名"）
 {"op":"add","module":"モジュール名","name":"マクロ名","code":"Sub マクロ名()\\n …\\nEnd Sub"}   新しいプロシージャをそのモジュールの末尾に足す（新規作成のときだけ）
 {"op":"compile"}                                       全体コンパイル（書き換えた後は道具が自動で行う）
 {"op":"rehearse","name":"マクロ名","args":[]}          コピーで試し撃ちして変化を見る（許されているときだけ）
 {"op":"list"}                                          このブックのマクロ一覧（どんなマクロがあるか）
 {"op":"impact","name":"マクロ名"}                      そのマクロの呼び元・呼び先（間接まで）。直す前に影響を見る
 {"op":"call_graph"}                                    呼び出し関係の全体図（"name" で 1 本に絞れる）
 {"op":"run","name":"マクロ名"}                         本物のブックで実際に動かす（許されているときだけ。実データが変わる）
 {"op":"undo"}                                          最後の直し（replace／code_replace）を 1 つ取り消して直す前の姿に戻す（直した後の試し撃ちで ○ が増えなかった・減ったとき。並べれば 2 つ前へ）

規則:
 - 直すのは最小の一手。1〜数行なら code_replace、まとまった直しだけ replace。関係ない箇所は触らない。
 - 新しく作るときは add（既にある名前と重ねない。材料の「モジュール」から標準モジュールを選ぶ）。
   引数つきの Sub や Function は作らない（引数なしの Sub 1 本で完結させる）。Option Explicit は足さない。
   対象は ActiveSheet でなく、そのマクロが使われるシートを名指しできるならそちらを使う。
 - 書き換えの前に、読んでいないコードを想像で書かない。呼び先のコードが要るなら get で読む。
 - Option Explicit を足さない。引数つきの Sub や Function を新しく作らない。宣言部や他のマクロを消さない。
 - replace の code は宣言（Sub/Function 〜）から End Sub/End Function まで 1 本だけ。名前は name と同じ。
 - 書き換えた後は道具がコンパイルする。通らなければ次の返事で直す。通ったら done にして report を書く。
 - 手は上から順に実行し、1 つ失敗すると残りは実行しない（結果に「止めました」と出る。それまでの書き換えが済んでいればコンパイルは行う）。失敗した手を直して、残りと一緒に並べ直す。
 - **直すと言った 1 本以外のコードを変えない。** 終わりに道具が撃つ前の全コードと照らし、他のプロシージャや宣言部が
   変わっていたら done を受け付けない。code_replace の search は**そのモジュールの中で当たる行を全部**書き換えるので、
   他のプロシージャにもある行（On Error Resume Next・Next i・End With など）を search にしない（前後を足して一意にする）。
 - **頼まれたことは全部やるか、できない分を必ず report に書くかのどちらか。できる分だけやって黙るのは禁止。**
   依頼を分解し、1 つずつ上の手のどれで当てるかを決める。当てる手が無いもの（フォームを作る・参照設定を変える・
   ブックを保存する・別のブックを触る）が 1 つでもあれば、report で「やったこと」と「できなかったこと
   （人がやるなら何をするか）」を分けて書く。コンパイルが通っただけで「動く」と書かない（動かして確かめたかを分けて書く）。
 - **「結果が合わない」「何も書かれない」「0 になる」の依頼は、直した後に必ず動かして確かめる**（rehearse が使えるとき）。
   コンパイルが通っても値が合うとは限らない。直した後に rehearse を撃っていない done は道具が受け付けない。
   依頼に「F8 は「対応なし」のはず」のような期待が書いてあれば、道具が試し撃ちの後の値と突き合わせて【結果】に出す。
   合わなければ仮説を立て直す（怪しく見える行を直す前に、依頼の症状と合う行かを確かめる。症状と関係ない行は直さない）。
   **直した後の試し撃ちで ○ が 1 つも増えなければ、その直しは外れ＝まず {"op":"undo"} で戻してから次の仮説へ**（外れの直しを積まない）。
   「結果が合わない」の依頼では、道具が直す前に 1 回撃った ○× を材料に付ける（直す前の姿）。
   **コンパイルエラーと数の合わなさが両方あるときは、まずコンパイルだけ直して rehearse**（○× を見てから数の直しを別の手で）。
   1 往復に当てずっぽうの直しを混ぜると、外れたときにどれを戻すか分からなくなる。
 - **「前は動いていた・前は合っていた」は、何かが変わった**。材料に「直前の控え」「鍛えて登録した時の正本」との差分があれば、
   原因の候補はまずその差分の行（差分を元に戻す）。差分に無い行を直すのは、差分を戻しても症状が残ったときだけ。
 - **「いつもの表では合うのに、この表だけ」は、この表の形（空の列・見出しの位置・余分な行）でだけ通る分岐が壊れている**。
   全部の表に効く式（列の足し算・行のずらし）を動かしてこの表だけ合わせない（いつもの表が壊れる）。
 - done は actions が空のときだけ有効。往復には上限がある。
"""


def _hands_part(rules):
    """RULES／MACRO_RULES から「手の一覧と規則」だけを切り出す（前置きの返事の形は統合側で 1 回だけ書く）。"""
    key = "actions に並べられる手"
    i = rules.find(key)
    return rules[i:] if i >= 0 else rules


# 両方またぐ依頼を 1 つのループで回すときの規約（2026-09-05）。
# 手の一覧は sheet と macro のものをそのまま連結する＝どちらかを直したら自動でここにも効く。
BOTH_RULES = ("""あなたは Excel のシートとマクロの両方を直す係です。道具（Python）が両方の材料を集め、
あなたの返事どおりに実行し、結果を次の【結果】で見せます。返事は JSON だけ（前後の文・コードフェンスは付けない）。

返事の形:
{"say": "いま何をするかを一言",
 "hypothesis": "マクロが動かない原因の仮説（マクロを直すときだけ・一言）",
 "plan": [{"item": "依頼を分解した 1 つ（短く）", "state": "未"}],
 "actions": [ ... ],
 "done": false,
 "report": "終わったときの報告（done が true のときだけ）"}

plan は 1 往復目に必ず書く。**「未」が 1 つでも残っている done は道具が受け付けない**（できない分は「不可」にして report に理由を書く）。

**シートの手とマクロの手を同じ actions に混ぜて並べてよい**（書いた順に実行される）。
どちらの手かは op の名前で道具が振り分ける。マクロを書き換えたときは、その往復の最後に道具が
全体コンパイルを 1 回だけ掛ける（手ごとには掛けない）。
順番は自分で決める: マクロを直してからその結果をシートに書くなら、先にマクロの手、後にシートの手を並べる。

--- ここからシートの手と規則 ---
""" + _hands_part(va.RULES) + """

--- ここからマクロの手と規則 ---
""" + _hands_part(MACRO_RULES) + """

--- 両方にまたがるときの規則 ---
 - 依頼のうちシートの仕事とマクロの仕事を plan で分けて書く（どちらかを黙って落とさないため）。
 - マクロを直しただけでは表の値は変わらない。**直したマクロを動かすかどうかは自分で決めない**
   （run は許されているときだけ）。動かしていないなら report に「直したが、動かして確かめてはいない」と書く。
 - シートの手が失敗したらマクロの手も止まる（その逆も同じ）。失敗した手を直して、残りと一緒に並べ直す。
 - 終わりに道具が両方を検査する（シート＝仕上げ検査・頼んでいない変化・自己採点／マクロ＝全体コンパイル）。
""")


# ----------------------------------------------------------------
# マクロ側の「頼んでいない変化」（2026-09-06）
#
# シート側には控え＋不変条件があるのに、マクロ側は無かった。code_replace は**そのモジュールの中で
# search に当たる行を全部**書き換える（On Error Resume Next・Next i のような行を指定すると、
# 隣のプロシージャまで一緒に変わる）。コンパイルは通るので、誰も気づかない。
# ここは撃つ前の全コードを控えて、直すと言った 1 本以外が変わっていないかを道具が数える口。
# ----------------------------------------------------------------

def _split_procs(code):
    """モジュールのコード → {プロシージャ名: 本文}（宣言部は '(宣言部)' に入れる）。"""
    out, name, buf = {}, '(宣言部)', []
    for line in (code or '').splitlines():
        m = _PROC_DECL_RE.match(line)
        if m:
            out[name] = chr(10).join(buf)
            name, buf = m.group(1), [line]
        else:
            buf.append(line)
    out[name] = chr(10).join(buf)
    return out


def _code_snapshot(wb):
    """ブックの全コードを {(モジュール, プロシージャ): 本文} で控える。読めなければ None。"""
    snap = {}
    try:
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ''
            for name, body in _split_procs(code).items():
                snap[(str(comp.Name), name)] = body
    except Exception:
        return None
    return snap


def _code_violations(before, now, allowed=(), allow_new=False):
    """撃つ前と今を照らして「頼んでいないコードの変化」を並べる（純 Python＝テストできる）。

    allowed＝直すと言ったマクロ（そこだけは変わってよい）。宣言部は Option や Dim が増えるだけでも
    咎める（Option Explicit を足すなの家の決まりが、ここで機械的に効く）。
    allow_new＝「作って」と言われた回（増えるのが正解なので咎めない。消えた・変わったは咎める。
    2026-09-07 未明の実射で、頼まれて足したマクロを咎めて往復を 2 回捨てていた）。
    """
    if before is None or now is None:
        return []
    ok = {str(a) for a in (allowed or ()) if a}
    bad = []
    for (mod, name), body in sorted(before.items()):
        if name in ok:
            continue
        if (mod, name) not in now:
            bad.append(f"頼んでいないマクロが消えた: {mod}.{name}")
        elif now[(mod, name)] != body:
            bad.append(f"頼んでいないマクロが変わった: {mod}.{name}")
    for (mod, name) in sorted(now):
        if allow_new:
            continue                 # 「作って」の回は増えるのが正解
        if name not in ok and (mod, name) not in before and name != '(宣言部)':
            bad.append(f"頼んでいないマクロが増えた: {mod}.{name}")
    return bad


def _find_macro_in_request(request, names):
    """依頼文の中にあるマクロ名（最長一致）。無ければ None。"""
    hits = [n for n in names if len(n) >= 2 and n in (request or '')]
    return max(hits, key=len) if hits else None


def _one_declared_name(code, what):
    """code に宣言が 1 本だけあることを確かめ、その名前を返す（replace / add 共通）。"""
    decls = _PROC_DECL_RE.findall(code or '')
    if len(decls) != 1:
        raise ValueError(f"{what} の code はプロシージャ 1 本だけ（宣言が {len(decls)} 本）")
    return decls[0]


def _modules_with_line(wb, needle):
    """needle を含む行があるモジュール名の一覧（大小文字は無視）。wb が無ければ空。"""
    if wb is None:
        return []
    low = needle.lower()
    out = []
    try:
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            n = cm.CountOfLines
            if n and low in cm.Lines(1, n).lower():
                out.append(comp.Name)
    except Exception:
        return []
    return out


# ----------------------------------------------------------------
# 複数行の code_replace（2026-09-17 夜・修理の試験で見つけた穴）
#
# 抜けた End If を 1 行足すだけの直しを、AI は 1 往復目（13 秒）で正しく書いたのに、code_replace が行単位で
# 断ったため、2 往復目で 197 行のプロシージャ全文を replace で送り直した（78 秒・出力 2,689 トークン）。
# 送り直しは遅いだけでなく、見えていない所を書き換える危険もある。並びごと 1 か所に当たる複数行の search は、
# 道具がそのプロシージャ 1 本の差し替え（replace-procedure＝控え・差分つき）に組み直す。
# ----------------------------------------------------------------

_PROC_END_RE = re.compile(r'^\s*End\s+(?:Sub|Function|Property)\s*$', re.IGNORECASE)


def _block_replace_code(module_code, search, replace):
    """純 Python: 複数行の search を module_code の中で並びごと 1 か所だけ探し、そのプロシージャの直した本文を返す。
    → (プロシージャ名, 直した本文)。当たらない・複数・宣言部・プロシージャをまたぐは ValueError。
    まず行末の空白だけ無視して探し、無ければ字下げと空白の数も無視して探す。"""
    lines = (module_code or '').replace('\r\n', '\n').split('\n')
    s_lines = search.replace('\r\n', '\n').strip('\n').split('\n')
    r_lines = replace.replace('\r\n', '\n').strip('\n').split('\n') if replace.strip('\n') else []
    if not any(l.strip() for l in s_lines):
        raise ValueError("search が空です")

    def _find(key):
        n = len(s_lines)
        want = [key(x) for x in s_lines]
        return [i for i in range(len(lines) - n + 1) if [key(x) for x in lines[i:i + n]] == want]

    hits = _find(lambda x: x.rstrip())
    if not hits:
        hits = _find(lambda x: re.sub(r'\s+', ' ', x.strip()))
    if not hits:
        raise ValueError("search の行の並びがそのままある所がありません（get で今の字面を読み直し、続いている行をそのまま書く）")
    if len(hits) > 1:
        raise ValueError(f"search の行の並びが {len(hits)} か所にあります。前後の行を足して 1 か所にしてください")
    i = hits[0]
    j = i + len(s_lines)
    starts = [k for k, l in enumerate(lines) if _PROC_DECL_RE.match(l)]
    ps = max((k for k in starts if k <= i), default=None)
    if ps is None:
        raise ValueError("search が宣言部に当たりました（直せるのはプロシージャの中だけ）")
    pe = next((k for k in range(ps + 1, len(lines)) if _PROC_END_RE.match(lines[k])), None)
    nxt = next((k for k in starts if k > ps), None)
    if pe is None or j - 1 > pe or (nxt is not None and j - 1 >= nxt):
        raise ValueError("search がプロシージャをまたいでいます（1 本の中だけにする）")
    new = lines[:i] + r_lines + lines[j:]
    new_pe = next((k for k in range(ps + 1, len(new)) if _PROC_END_RE.match(new[k])), None)
    if new_pe is None or new_pe < i + len(r_lines) - 1:
        raise ValueError("replace で End Sub／End Function が消えます（プロシージャの終わりは残す）")
    proc = '\n'.join(new[ps:new_pe + 1])
    name = _PROC_DECL_RE.match(lines[ps]).group(1)
    if len(_PROC_DECL_RE.findall(proc)) != 1:
        raise ValueError("replace に Sub／Function の宣言を入れないでください（新しく作るなら add）")
    return name, proc


def _replace_loss_note(old_proc, new_proc):
    """丸ごと差し替えの前後 → 消えた行の知らせ（無ければ ''）。純 Python。
    2026-09-17 夜: Next c 1 行の抜けを直すのに AI が 386 行を replace で送り直し、注記 31 行を黙って落とした
    （コンパイルも値も通る＝誰も気づかない）。字下げと空行の違いは数えない。"""
    import difflib
    if not old_proc or not new_proc:
        return ''
    a = [re.sub(r'\s+', ' ', l.strip()) for l in old_proc.replace('\r\n', '\n').split('\n') if l.strip()]
    b = [re.sub(r'\s+', ' ', l.strip()) for l in new_proc.replace('\r\n', '\n').split('\n') if l.strip()]
    gone, added = [], 0
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag in ('delete', 'replace'):
            gone += a[i1:i2]
        if tag in ('insert', 'replace'):
            added += j2 - j1
    if not gone:
        return ''
    notes = [l for l in gone if l.startswith("'")]
    if len(gone) <= 3 and not notes:
        return ''
    return (f"\n⚠ 丸ごと差し替えで元の行が {len(gone)} 行消えて（うち注記 {len(notes)} 行）、{added} 行増えました。"
            "直すと言った所以外の行が消えていたら元に戻してください（数行の直しは code_replace の複数行で送る）。"
            + ("\n  消えた行の例: " + " ／ ".join(gone[:4]) if gone else ""))


def _proc_text(wb, name, module=None):
    """いまのプロシージャの本文（読めなければ ''）。"""
    try:
        for comp in wb.VBProject.VBComponents:
            if module and str(comp.Name) != module:
                continue
            cm = comp.CodeModule
            n = int(cm.CountOfLines)
            procs = _split_procs(str(cm.Lines(1, n)) if n else '')
            for k, body in procs.items():
                if k.lower() == str(name).lower():
                    return body
    except Exception:
        pass
    return ''


def _block_replace(act, wb, default_module=None):
    """複数行の code_replace を実行する。→ (表示名, ok, 出力, 書き換えたか)。"""
    search, repl = act.get('search'), act.get('replace')
    module = str(act.get('module') or '').strip() or (str(default_module).strip() if default_module else '')
    label = "code_replace（複数行）"
    if not module:
        first = next((l.strip() for l in search.split('\n') if l.strip()), '')
        hits = _modules_with_line(wb, first)
        if len(hits) != 1:
            return label, False, ("実行していません: 複数行の code_replace は module を書いてください"
                                  + (f"（search の 1 行目が {len(hits)} つのモジュールにあります: {' / '.join(hits)}）" if hits else "")), False
        module = hits[0]
    try:
        cm = wb.VBProject.VBComponents(module).CodeModule
        n = int(cm.CountOfLines)
        code = str(cm.Lines(1, n)) if n else ''
    except Exception as ex:
        return label, False, f"実行していません: モジュール '{module}' を読めませんでした: {ex}", False
    try:
        name, proc = _block_replace_code(code, search, repl)
    except ValueError as ex:
        return label, False, f"実行していません: {ex}", False
    with open(va.LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(proc if proc.endswith('\n') else proc + '\n')
    ok, out = va._run_cmd(['replace-procedure', '-y', '--module', module], wb)
    return f"code_replace（複数行→{name} を差し替え）", ok, out, ok


# ----------------------------------------------------------------
# 「結果が合わない」の修理は動かして確かめる（2026-09-17 夜・修理の試験で見つけた穴）
#
# 「F8 は「対応なし」のはずが空」の依頼で、AI は壊れた行（uRow）と一緒に正しい行（fCol）まで書き換え、
# 一度も撃たずに「正しく出力されるはずです」で done にした。コンパイルは通り、前の表も別の表 8 枚も全部外れた。
# 結果の依頼は、直した後の試し撃ちが無い done を受け付けない。依頼に書かれた期待は道具が試し撃ちの値と突き合わせる。
# ----------------------------------------------------------------

_RESULT_COMPLAINT_RE = re.compile(
    r'合わ(?:ない|なく|ず)|おかしい|違(?:う|って)|ずれ|何も(?:書かれ|書け|出な|起き|変わら)|空(?:になる|のまま|になって)'
    r'|0\s*に(?:なる|なって)|のはず|足りない|多すぎ|消え(?:る|た|て)|重複して')
_EXPECT_RE = re.compile(r'([A-Z]{1,3}[1-9][0-9]*)\s*(?:は|が|に)\s*(?:「([^」]*)」|(空))\s*(?:の|に|と)?\s*(?:はず|なるはず|なっているはず)')
_SNAP_PATHS_RE = re.compile(r'実行前/後 snapshot:\s*(.+?\.before\.json)\s*/\s*(.+?\.after\.json)')
_SHEET_IN_REQ_RE = re.compile(r'シート「([^」]+)」|「([^」]+)」シート')


def _needs_run(request):
    """依頼が「結果が合わない」型か（直した後に動かして確かめるべきか）。純 Python。"""
    return bool(_RESULT_COMPLAINT_RE.search(request or ''))


_EXPECT_SIZE_RE = re.compile(r'(\d+)\s*(行|列)\s*の\s*はず')


def _request_expectations(request):
    """依頼の「F8 は「対応なし」のはず」「G2 は空のはず」「7 列のはずが 8 列」→ [(番地 or #行/#列, 期待の文字)]。純 Python。
    2026-09-17 夜: 表の大きさの期待を読まず、例のセル 2 つだけ ○ にした直しを合格にした（列はまだ 1 つ多かった）。"""
    out = []
    for m in _EXPECT_RE.finditer(request or ''):
        out.append((m.group(1), '' if m.group(3) else m.group(2)))
    for m in _EXPECT_SIZE_RE.finditer(request or ''):
        key = '#行' if m.group(2) == '行' else '#列'
        if all(k != key for k, _w in out):
            out.append((key, m.group(1)))
    return out


def _same_text(a, b):
    a, b = ('' if a is None else str(a)).strip(), ('' if b is None else str(b)).strip()
    if a == b:
        return True
    try:
        return abs(float(a.replace(',', '').replace('¥', '')) - float(b.replace(',', '').replace('¥', ''))) < 1e-6
    except ValueError:
        return False


def _cells_of_snapshot(path, sheet_hint=None):
    """rehearse の .after.json → {番地: 文字}（sheet_hint → アクティブ → 1 枚目の順でシートを選ぶ）。"""
    with open(path, encoding='utf-8') as f:
        doc = json.load(f)
    sheets = doc.get('sheets') or {}
    name = (sheet_hint if sheet_hint in sheets else None) or (doc.get('active') if doc.get('active') in sheets else None) \
        or next(iter(sheets), None)
    cells = {}
    info = sheets.get(name) or {}
    for row in (info.get('cells') or []):
        for col, v in (row.get('c') or {}).items():
            cells[f"{col}{row.get('r')}"] = v
    dm = re.match(r'(\d+)行 x (\d+)列', str(info.get('dims') or ''))
    if dm:
        cells['#行'], cells['#列'] = dm.group(1), dm.group(2)          # 使用範囲の大きさ（表の行数・列数の期待と比べる）
    return name, cells


def _expectation_check(expectations, cells):
    """期待と試し撃ちの後の値 → (全部合ったか, 行の並び)。純 Python。"""
    lines, ok = [], True
    for addr, want in expectations:
        got = cells.get(addr, '')
        hit = _same_text(want, got)
        ok = ok and hit
        if addr.startswith('#'):
            unit = addr[1:]
            lines.append(f"  表の{unit}数（使用範囲）: 期待 {want} {unit}→ 試し撃ちの後 {got or '?'} {unit} {'○' if hit else '×'}")
        else:
            lines.append(f"  {addr}: 期待「{want or '空'}」→ 試し撃ちの後「{got or '空'}」 {'○' if hit else '×'}")
    return ok, lines


def _score_note(prev, now, wrote):
    """直しの前後の ○ の数 → AI への知らせ（無ければ ''）。純 Python（2026-09-17 夜）。
    修理の試験で、当てずっぽうの直しが ○ を 1 つも増やさないのに戻さず、その上に次の直しを積んで 6 行壊した。"""
    if not wrote or prev is None or now is None:
        return ''
    (h0, t0), (h1, t1) = prev, now
    if t0 != t1:
        return ''
    if h1 < h0:
        return (f"\n⚠ この直しで ○ が減りました（{h0}/{t0} → {h1}/{t1}）＝この直しが壊しています。"
                "{\"op\":\"undo\"} で直す前に戻してください。")
    if h1 == h0 and h1 < t1:
        return (f"\n⚠ この直しの後も ○ は {h1}/{t1} のまま（直す前も {h0}/{t0}）＝直した行は症状と関係なかった可能性が高いです。"
                "{\"op\":\"undo\"} で直す前に戻してから、別の仮説を立ててください（関係ない直しを積まない）。")
    return ''


# ----------------------------------------------------------------
# 大きいマクロの続きを読む（2026-09-17 夜・表の書き方と罫線と列幅をそろえる 約 700 行の修理で見つけた穴）
#
# repair の本文は字数で切れる（…ここから 7,953 字を省略）のに、続きを読む手が無かった。AI は GoSub のラベル
# 「数を読む」を get で読もうとして「プロシージャが見つかりません」、grep で 1 行ずつ拾うしかなく 4 往復を読むだけで使い切った。
# ----------------------------------------------------------------

_LINES_MAX = 160


def _module_lines(wb, module):
    cm = wb.VBProject.VBComponents(module).CodeModule
    n = int(cm.CountOfLines)
    return str(cm.Lines(1, n)).replace('\r\n', '\n').split('\n') if n else []


def _numbered(lines, start):
    return "\n".join(f"{start + i:>4}: {l}" for i, l in enumerate(lines))


def _label_block(lines, label):
    """モジュールの行からラベル label: 〜 Return（無ければ次のラベルか End Sub の手前）を抜く → (開始の行番号, 行) or None。"""
    pat = re.compile(r'^\s*' + re.escape(label) + r'\s*:\s*$')
    s = next((i for i, l in enumerate(lines) if pat.match(l)), None)
    if s is None:
        return None
    e = s
    for k in range(s + 1, len(lines)):
        e = k
        if re.match(r'^\s*Return\s*$', lines[k], re.IGNORECASE) or _PROC_END_RE.match(lines[k]):
            break
    return s + 1, lines[s:e + 1]


def _read_lines_act(act, wb, module=None):
    """{"op":"lines","from":N,"to":M} → (表示名, ok, 出力)。行番号は grep の「マクロ名:661」と同じ（モジュールの行）。"""
    module = str(act.get('module') or '').strip() or (str(module).strip() if module else '')
    try:
        a, b = int(act.get('from')), int(act.get('to') or act.get('from'))
    except (TypeError, ValueError):
        return "lines", False, "実行していません: lines には from と to（行番号の数）が要ります"
    if not module:
        return "lines", False, "実行していません: lines には module が要ります"
    try:
        lines = _module_lines(wb, module)
    except Exception as ex:
        return "lines", False, f"実行していません: モジュール '{module}' を読めませんでした: {ex}"
    a, b = max(1, a), min(len(lines), b)
    if b < a:
        return "lines", False, f"実行していません: 行の範囲が空です（モジュールは {len(lines)} 行）"
    if b - a + 1 > _LINES_MAX:
        b = a + _LINES_MAX - 1
    tail = f"\n（{b + 1} 行目から先は from を {b + 1} にして読む）" if b < len(lines) else ""
    return f"lines {a}-{b}", True, _numbered(lines[a - 1:b], a) + tail


def _get_label_act(act, wb, module=None):
    """get の name がプロシージャでなく GoSub のラベルなら、その塊を返す。当たらなければ None（普通の get に任せる）。"""
    name = str(act.get('name') or '').strip()
    mods = [str(act.get('module') or '').strip() or (str(module).strip() if module else '')]
    for mod in [m for m in mods if m]:
        try:
            got = _label_block(_module_lines(wb, mod), name)
        except Exception:
            got = None
        if got:
            start, block = got
            return (f"get {name}（GoSub のラベル）", True,
                    f"[{mod}] ラベル {name}: の塊（{start}〜{start + len(block) - 1} 行）\n" + _numbered(block, start))
    return None


# ----------------------------------------------------------------
# 要らない直しを道具が戻す（2026-09-17 夜・修理の試験 種 77）
#
# AI は壊れた行（outRow）と一緒に関係ない行（uRow）も書き換え、依頼に書かれた 3 セルは ○ になったので done にした。
# 例に無い「対応なし」の行が壊れ、別の表 8 枚が全部外れた（嘘の完了）。例のセルだけでは「直しすぎ」が見えない。
# done の後に、直す前の姿との差を塊ごとに 1 つずつ戻して試し撃ちし、戻しても期待が全部 ○ の塊は要らない直しとして戻す。
# ----------------------------------------------------------------

_MINIMIZE_TRIALS = 6


def _fix_hunks(before_lines, after_lines):
    """直す前と今の行の差の塊（純 Python）→ [(tag, i1, i2, j1, j2)]。"""
    import difflib
    return [op for op in difflib.SequenceMatcher(None, before_lines, after_lines, autojunk=False).get_opcodes()
            if op[0] != 'equal']


def _revert_hunk(before_lines, after_lines, hunk):
    """今の行の塊 1 つだけを直す前に戻した行（純 Python）。"""
    _tag, i1, i2, j1, j2 = hunk
    return after_lines[:j1] + before_lines[i1:i2] + after_lines[j2:]


def _apply_proc(wb, text, module=None):
    with open(va.LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(text if text.endswith('\n') else text + '\n')
    ok, _out = va._run_cmd(['replace-procedure', '-y'] + (['--module', module] if module else []), wb)
    return ok


def _minimize_fix(wb, macro, module, start_proc, expectations, sheet_hint=None, max_trials=_MINIMIZE_TRIALS):
    """直しの塊を 1 つずつ戻して試し撃ち → 戻しても期待が全部 ○ の塊は戻したままにする。→ 知らせの行の並び。"""
    notes = []
    cur = _proc_text(wb, macro, module) if start_proc else ''
    if not cur:
        return notes

    def _lines(t):
        return t.replace('\r\n', '\n').rstrip('\n').split('\n')

    before = _lines(start_proc)
    trials = 0
    while trials < max_trials:
        now = _lines(cur)
        hunks = _fix_hunks(before, now)
        if len(hunks) < 2:
            break
        dropped = False
        for h in hunks:
            if trials >= max_trials:
                break
            trials += 1
            cand = '\n'.join(_revert_hunk(before, now, h))
            if not _apply_proc(wb, cand, module):
                continue
            ok_r, out_r = va._run_cmd(['rehearse', macro, '--timeout', '120'], wb)
            _res, v, _s = _verify_results([(f"rehearse {macro}", ok_r, out_r)], None, expectations, sheet_hint, wb, None)
            if v is True:
                gone = " / ".join(l.strip() for l in now[h[3]:h[4]] if l.strip()) or '（足した行なし）'
                back = " / ".join(l.strip() for l in before[h[1]:h[2]] if l.strip()) or '（行を抜く）'
                notes.append(f"要らない直しを道具が戻しました（戻しても依頼の期待は全部 ○）: 「{gone}」→「{back}」")
                cur = cand
                dropped = True
                break                                      # 塊を数え直す
            _apply_proc(wb, cur, module)                   # 要る直し＝元に戻す
        if not dropped:
            break
    return notes


def _undo_step(stack, wb, module=None):
    """{"op":"undo"}: 前の往復で直す前の姿（stack の最後）へプロシージャを戻す（replace-procedure＝控えつき）。"""
    label = "undo（直す前に戻す）"
    if not stack:
        return [(label, False, "実行していません: 戻す先がありません（まだ直していない）")]
    prev = stack.pop()
    if not prev:
        return [(label, False, "実行していません: 直す前の姿を読めていませんでした（get で読み、replace で戻す）")]
    with open(va.LAST_PROC_FILE, 'w', encoding='utf-8') as f:
        f.write(prev if prev.endswith('\n') else prev + '\n')
    ok, out = va._run_cmd(['replace-procedure', '-y'] + (['--module', module] if module else []), wb)
    res = [(label, ok, out)]
    if ok:
        okc, outc = va._run_cmd(['compile'], wb)
        res.append(("compile（書き換え後・道具が自動）", okc, outc))
    return res


def _verify_results(results, verified, expectations, sheet_hint=None, wb=None, score=None):
    """1 往復の結果を順に読み、「最後の書き換えの後に動かして確かめたか」を進める。→ (結果, verified, score)。
    書き換え（replace／code_replace／add）で None に戻す。rehearse／run が通れば、依頼の期待があれば値と突き合わせて
    ○× の行を結果に足し、全部○なら True（期待が無ければ通っただけで True）。実行が失敗なら False。
    score＝最後に突き合わせた (○ の数, 期待の数)。この往復で書き換えてから撃ったとき、前の score と比べた知らせも足す。"""
    out = []
    wrote = False
    for label, ok, text in results:
        lab = str(label)
        if ok and lab.startswith(_MACRO_WROTE_LABELS + ('undo',)):
            verified = None
            wrote = True
        elif lab.startswith(('rehearse ', 'run ')):
            if not ok:
                verified = False
            elif not expectations:
                verified = True
            else:
                cells, where = None, ''
                try:
                    if lab.startswith('rehearse '):
                        m = _SNAP_PATHS_RE.search(text or '')
                        if m:
                            where, cells = _cells_of_snapshot(m.group(2).strip(), sheet_hint)
                    elif wb is not None:
                        ws = wb.Worksheets(sheet_hint) if sheet_hint else wb.ActiveSheet
                        where = str(ws.Name)
                        cells = {a: va._text_of(ws.Range(a).Value) for a, _w in expectations if not a.startswith('#')}
                        cells['#行'], cells['#列'] = str(ws.UsedRange.Rows.Count), str(ws.UsedRange.Columns.Count)
                except Exception as ex:
                    text = (text or '').rstrip() + f"\n（依頼の期待と突き合わせられませんでした: {ex}）"
                if cells is None:
                    verified = True
                else:
                    hit, lines = _expectation_check(expectations, cells)
                    verified = hit
                    now = (sum(1 for l in lines if l.endswith('○')), len(lines))
                    text = ((text or '').rstrip() + f"\n--- 依頼の期待との突き合わせ（道具・シート {where}） ---\n"
                            + "\n".join(lines)
                            + ("\n全部合いました。" if hit else
                               "\n合っていない所があります＝まだ直っていません（直した行が症状と関係あるかを見直す）。")
                            + _score_note(score, now, wrote))
                    score = now
                    wrote = False
        out.append((label, ok, text))
    return out, verified, score


def _macro_action_to_tokens(act, allow_rehearse=False, existing=None, default_module=None, wb=None):
    """AI の手 1 つ → (表示名, コマンド引数列, 書き換えか)。規約外は ValueError。

    default_module＝いま修理しているマクロのモジュール。code_replace で AI が module を書かなければこれを使う
    （省くとブック全体が置換の対象になっていた・2026-09-04）。
    """
    if not isinstance(act, dict):
        raise ValueError("actions の要素がオブジェクトではありません")
    op = str(act.get('op') or '')
    if op == 'undo':
        raise ValueError("undo はマクロだけの修理（macro モード）の往復で、ほかの手と別に使えます")
    if op not in _MACRO_OPS:
        raise ValueError(f"許していない手です: {op!r}（" + " / ".join(_MACRO_OPS) + "）")
    module = str(act.get('module') or '').strip()
    if op == 'code_replace' and not module and default_module:
        module = str(default_module).strip()
    if op == 'list':
        return "list", ['list'], False
    if op == 'impact':
        name = str(act.get('name') or '').strip()
        if not name:
            raise ValueError("impact には name（マクロ名）が要ります")
        return f"impact {name}", ['impact', name], False
    if op == 'call_graph':
        name = str(act.get('name') or '').strip()
        # cmd_call_graph は位置引数を見ない（--macro だけ）＝名前が黙って無視されていた（2026-09-04）
        return (f"call_graph {name}" if name else "call_graph"), (['call-graph', '--macro', name] if name else ['call-graph']), False
    if op == 'run':
        name = str(act.get('name') or '').strip()
        if not name:
            raise ValueError("run には name（マクロ名）が要ります")
        if not allow_rehearse:
            raise ValueError("run（本物のブックで実際に動かす）は --rehearse を付けたときだけ使えます。"
                             "付いていないときは rehearse も run も使えないので、直してコンパイルが通ったら done にし、"
                             "report に「動かして確かめるところまではしていない」と書く")
        return f"run {name}", ['run-macro', name, '--timeout', '120'], False
    if op == 'add':
        name = str(act.get('name') or '').strip()
        code = act.get('code')
        if not name or not isinstance(code, str) or not code.strip():
            raise ValueError("add には name と code（Sub〜End Sub の全文）が要ります")
        if not module:
            raise ValueError("add には module（入れ先のモジュール名。材料の「モジュール」から選ぶ）が要ります")
        decl = _one_declared_name(code, 'add')
        if decl.lower() != name.lower():
            raise ValueError(f"add の code の名前（{decl}）が name（{name}）と違います")
        if existing and name.lower() in {n.lower() for n in existing}:
            raise ValueError(f"'{name}' は既にあります（新しく作るのでなく直すなら replace か code_replace）")
        return f"add {name} → {module}", ['add-procedure', '-y', module], True
    if op == 'get':
        name = str(act.get('name') or '').strip()
        if not name:
            raise ValueError("get には name が要ります")
        return f"get {name}", (['get', module, name] if module else ['get', name]), False
    if op == 'grep':
        text = str(act.get('text') or '')
        if not text.strip():
            raise ValueError("grep には text が要ります")
        return f"grep {text[:30]}", ['grep', text], False
    if op == 'code_replace':
        search, repl = act.get('search'), act.get('replace')
        if not isinstance(search, str) or not search or not isinstance(repl, str):
            raise ValueError("code_replace には search（文字列）と replace（文字列）が要ります")
        if '\n' in search or '\r' in search or '\n' in repl or '\r' in repl:
            raise ValueError("複数行の code_replace はマクロの手の並びでだけ使えます（_execute_macro が組み直す）")
        if not module:
            # module 無しはブックの全モジュールが対象。search の行があるモジュールを数え、1 つならそれを
            # 使う・複数なら断る（一律必須にしていたのを戻した・2026-09-04 夜。「同名が複数にあるときだけ必須」）
            hits = _modules_with_line(wb, search)
            if len(hits) == 1:
                module = hits[0]
            elif len(hits) > 1:
                raise ValueError(f"code_replace の search が {len(hits)} つのモジュールにあります（"
                                 + " / ".join(hits) + "）。module でどれか 1 つに絞ってください")
            # 0 件は CLI に任せる（「マッチなし」が返り AI に戻る）
        toks = ['code-replace', '-y', search, repl] + (['--module', module] if module else [])
        return f"code_replace {search[:30]}", toks, True
    if op == 'replace':
        name = str(act.get('name') or '').strip()
        code = act.get('code')
        if not name or not isinstance(code, str) or not code.strip():
            raise ValueError("replace には name と code（Sub〜End Sub の全文）が要ります")
        decl = _one_declared_name(code, 'replace')
        if decl.lower() != name.lower():
            raise ValueError(f"replace の code の名前（{decl}）が name（{name}）と違います")
        toks = ['replace-procedure', '-y'] + (['--module', module] if module else [])
        return f"replace {name}", toks, True
    if op == 'compile':
        return "compile", ['compile'], False
    name = str(act.get('name') or '').strip()
    if not name:
        raise ValueError("rehearse には name が要ります")
    if not allow_rehearse:
        raise ValueError("rehearse は --rehearse を付けたときだけ使えます")
    args = [str(a) for a in (act.get('args') or []) if a is not None]
    return f"rehearse {name}", ['rehearse', name] + args + ['--timeout', '120'], False


_NO_MATCH_RE = re.compile(r'にマッチする行はありません|置換後も同じ内容です')


def _execute_macro(actions, allow_rehearse=False, wb=None, existing=None, default_module=None,
                   compile_at_end=True, undo=None):
    """macro の actions を順に実行。書き換えの後は必ず compile を足す。[(表示名, ok, 出力)]

    existing（既にあるマクロ名）は list なら add が通るたびに足す（同じ名前を 2 回 add できていた・2026-09-04）。
    1 つでも失敗したら、それ以降の手は実行せず「止めました」を積んで終わる（2026-09-05。sheet 側の _execute と同じ型）。
    それまでに書き換えが済んでいれば（wrote）、止めた後もコンパイルは行う。
    compile_at_end=False にすると最後のコンパイルを足さない（シートの手と混ぜて実行する側が、
    1 往復につき 1 回だけ掛けるため・2026-09-05）。
    """
    results = []
    wrote = False
    for idx, act in enumerate(actions):
        # undo＝(積む先, 直しているマクロ名): 直す手の前にそのプロシージャの姿を読み、直しが通ったら積む
        # （{"op":"undo"} が戻すのは直し 1 つ分。往復単位だと同じ往復のコンパイルの直しまで戻る・2026-09-17 夜）
        snap = (_proc_text(wb, undo[1], default_module)
                if undo and undo[1] and isinstance(act, dict) and act.get('op') in ('replace', 'code_replace') else None)
        read = None
        if isinstance(act, dict) and act.get('op') == 'lines' and wb is not None:
            read = _read_lines_act(act, wb, default_module)
        elif (isinstance(act, dict) and act.get('op') == 'get' and wb is not None and existing is not None
              and str(act.get('name') or '').strip().lower() not in {str(n).lower() for n in existing}):
            read = _get_label_act(act, wb, default_module)
        if read:
            results.append(read)
            if not read[1]:
                for rest in actions[idx + 1:]:
                    rop = f"{rest.get('op', '?') if isinstance(rest, dict) else '?'}"
                    results.append((rop, False, f"実行していません: 前の手「{read[0]}」が失敗したので止めました。並べ直して返してください"))
                break
            continue
        if (isinstance(act, dict) and act.get('op') == 'code_replace' and isinstance(act.get('search'), str)
                and act.get('search') and isinstance(act.get('replace'), str)
                and any(ch in act['search'] + act['replace'] for ch in '\r\n')):
            label, ok, out, w = _block_replace(act, wb, default_module)     # 複数行＝プロシージャ 1 本の差し替えに組み直す
            results.append((label, ok, out))
            wrote = wrote or w
            if ok and snap:
                undo[0].append(snap)
            if not ok:
                for rest in actions[idx + 1:]:
                    rop = f"{rest.get('op', '?') if isinstance(rest, dict) else '?'}"
                    results.append((rop, False, f"実行していません: 前の手「{label}」が失敗したので止めました。"
                                                "並べ直して返してください"))
                break
            continue
        try:
            label, toks, is_write = _macro_action_to_tokens(act, allow_rehearse, existing, default_module, wb)
        except ValueError as ex:
            results.append((f"{act.get('op', '?') if isinstance(act, dict) else '?'}", False, f"実行していません: {ex}"))
        else:
            old_proc = ''
            if isinstance(act, dict) and act.get('op') in ('replace', 'add'):
                code = act['code'].replace('\r\n', '\n')
                with open(va.LAST_PROC_FILE, 'w', encoding='utf-8') as f:
                    f.write(code if code.endswith('\n') else code + '\n')
                if act.get('op') == 'replace' and wb is not None:
                    old_proc = _proc_text(wb, str(act.get('name') or ''), str(act.get('module') or '').strip() or None)
            ok, out = va._run_cmd(toks, wb)
            if ok and old_proc:
                out = (out or '').rstrip() + _replace_loss_note(old_proc, act['code'])
            if ok and isinstance(act, dict) and act.get('op') == 'code_replace' and _NO_MATCH_RE.search(out or ''):
                # code-replace は 1 行も当たらなくても成功で返る（コマンドとしては正しい）。
                # agent では「直した」と AI が誤解し、compile も通って done になっていた（2026-09-04）
                ok = False
                out = (out.rstrip() + "\n（1 行も書き換わっていません＝直っていません。get で今の字面を読み直し、"
                                      "search をそのとおりに書くか、replace でプロシージャの全文を入れ直してください）")
            results.append((label, ok, out))
            wrote = wrote or (is_write and ok)
            if ok and snap:
                undo[0].append(snap)
            if ok and isinstance(act, dict) and act.get('op') == 'add' and isinstance(existing, list):
                existing.append(str(act.get('name') or '').strip())
        if not results[-1][1]:
            failed_label = results[-1][0]
            for rest in actions[idx + 1:]:
                rop = f"{rest.get('op', '?') if isinstance(rest, dict) else '?'}"
                results.append((rop, False,
                                f"実行していません: 前の手「{failed_label}」が失敗したので止めました。"
                                "並べ直して返してください"))
            break
    if wrote and compile_at_end:
        ok, out = va._run_cmd(['compile'], wb)
        results.append(("compile（書き換え後・道具が自動）", ok, out))
    return results


_MACRO_WROTE_LABELS = ('replace', 'code_replace', 'add ')


def _macro_wrote(results):
    """結果の並びに「マクロを書き換えた」手があるか（混ぜて実行する側が compile を掛けるかの判定）。"""
    return any(ok and str(label).startswith(_MACRO_WROTE_LABELS) for label, ok, _out in results)


def _execute_mixed(actions, sheet, wb, allow_rehearse=False, existing=None, default_module=None):
    """シートの手とマクロの手が混ざった並びを、書かれた順に実行する（2026-09-05・統合ループ）。

    op の名前でどちらの実行系に渡すかを決める（シートの手とマクロの手に同じ名前は無い）。
    片方が失敗したら残りは実行しない（どちらの側でも同じ）。コンパイルは 1 往復に 1 回だけ、
    マクロを書き換えたときに最後へ足す（手ごとに掛けると往復が遅くなる）。
    """
    results = []
    for idx, act in enumerate(actions):
        op = (act or {}).get('op', '?') if isinstance(act, dict) else '?'
        if op in _MACRO_OPS:
            got = _execute_macro([act], allow_rehearse, wb, existing, default_module, compile_at_end=False)
        else:
            got = va._execute([act], sheet, wb)
        results += got
        if not got or not got[-1][1]:
            va._stop_rest(results, actions, idx)
            break
    if _macro_wrote(results):
        ok, out = va._run_cmd(['compile'], wb)
        results.append(("compile（書き換え後・道具が自動）", ok, out))
    return results


_GET_MODULE_RE = re.compile(r'^モジュール\s*[:：]\s*(\S+)', re.MULTILINE)


def _macro_materials(macro, wb=None):
    """(材料の文, repair が通ったか, 修理対象のモジュール名 or None)

    2026-09-17: get＋call-graph＋compile の 3 手を repair 1 手に（本文・入口・呼び元呼び先・コンパイルの行名指し・
    check の error 級・直前の控えとの差分が同じ形で揃う。AI に渡す材料と人が repair で見る物が同じになる）。
    """
    parts = []
    ok, out = va._run_cmd(['repair', macro], wb)
    parts.append(("repair（本文・入口・呼び元呼び先・コンパイル・check・控えとの差分）", ok, out))
    m = _GET_MODULE_RE.search(out or '')
    module = m.group(1) if m else None
    text = []
    for label, ok, out in parts:
        body = out.rstrip()
        # 直す本文（get／repair）は切らない。切ったまま渡すと、AI は見えていない後半ごと replace で書き直し、
        # 末尾が黙って消える（コンパイルは通るので誰も気づかない・2026-09-04）
        limit = va._RESULT_LIMIT_BIG if label == 'get' or label.startswith('repair') else va._RESULT_LIMIT
        if len(body) > limit:
            cut = body[:limit]
            shown = cut.count('\n')
            body = cut + (f"\n…（ここから {len(body) - limit:,} 字を省略しました。"
                          "**この続きは見えていないので、replace で全文を書き直さないでください。"
                          "直すのは code_replace で変更行だけにします**。続きは "
                          f"{{\"op\":\"lines\",\"from\":行,\"to\":行,\"module\":\"{module or 'モジュール名'}\"}} で読む"
                          f"（本文はおよそ {max(1, shown - 5)} 行目まで見えている。GoSub のラベルは get にラベル名）)")
        text.append(f"--- {label} {'' if ok else '（失敗）'}---\n" + body)
    return "\n".join(text), parts[0][1], module


def _macro_materials_create(wb=None):
    """新しく作るときの材料: モジュールの一覧（入れ先）・既にあるマクロ（名前を重ねない）・いまのコンパイル。"""
    parts = []
    for label, toks in (("モジュール（この中から入れ先を選ぶ）", ['list-modules']),
                        ("既にあるマクロ（名前を重ねない）", ['list', '--standard']),
                        ("compile（いまの状態）", ['compile'])):
        ok, out = va._run_cmd(toks, wb)
        parts.append((label, ok, out))
    text = [f"--- {label} {'' if ok else '（失敗）'}---\n" + out.rstrip()[:va._RESULT_LIMIT] for label, ok, out in parts]
    return "\n".join(text), parts[0][1]


def run_both(request, sheet, xl, wb, macro=None, ai=va._CC_AI, model=None, max_turns=va._DEFAULT_MAX_TURNS,
             dry_run=False, rehearse=False, show_image=True, grade=True,
             grade_ai=None, grade_model=None):
    """シートとマクロにまたがる依頼を、1 つのループで回す（2026-09-05・統合）。

    関所（plan・仕上げ検査・頼んでいない変化・自己採点・検査）は run_agent のものをそのまま使い、
    規約（BOTH_RULES）と実行系（_execute_mixed）だけ差し替える＝関所の実装を 2 つに増やさない。
    材料はシートの材料＋マクロの一覧（依頼文にマクロ名があれば、その 1 本のコードと呼び出し関係も）。
    """
    ai, model, api_key = va._ai_setup(ai, model)
    print(f"対象ブック: {wb.Name}   シート: {sheet}   ＋マクロ（1 つのループでまたいで直します）")
    names = []
    with contextlib.suppress(Exception):
        names = va._all_procedure_names(wb)
    macro = macro or _find_macro_in_request(request, names)
    ok_m, materials = va._run_cmd(['materials', sheet], wb)
    if not ok_m:
        raise RuntimeError("materials が失敗しました:\n" + materials.strip())
    try:
        materials += va._extra_materials(wb, wb.Sheets(sheet))
    except Exception as ex:
        materials += f"（追加の材料を読めませんでした: {ex}）\n"
    materials += va._notes_recall(str(wb.Name), sheet)
    if macro:                                   # 直す先が分かっているなら、そのコードと呼び出し関係も材料に
        mat_m, got, module = _macro_materials(macro, wb)
        materials += "\n--- マクロの材料（道具が読んだ現物） ---\n" + mat_m
    else:
        module = None
        mat_m, _got = _macro_materials_create(wb)
        materials += "\n--- マクロの材料（道具が読んだ現物） ---\n" + mat_m
    existing = list(names)

    def _mixed(actions, sh, book):
        return _execute_mixed(actions, sh, book, allow_rehearse=rehearse,
                              existing=existing, default_module=module)

    r = va.run_agent(request, sheet, wb, ai, model, max_turns, dry_run, materials=materials,
                  show_image=show_image, grade=grade, rules=BOTH_RULES, execute=va._mixed, mode='both',
                  grade_ai=grade_ai, grade_model=grade_model)
    if not dry_run:
        ok_c, out = va._run_cmd(['compile'], wb)     # マクロを触っていなくても、壊れたまま終わらせない
        print("検査: 全体コンパイル " + ("通過" if ok_c else "不通過") + ("" if ok_c else "\n" + out.rstrip()))
        r['compile'] = ok_c
        r['ok'] = r['ok'] and ok_c
        print("（--undo で戻せるのはシートだけです。マクロの書き換えは list-backups → restore で戻します）")
    return r


def run_macro_agent(request, xl, wb, macro=None, ai=va._CC_AI, model=None, max_turns=va._DEFAULT_MAX_TURNS,
                    dry_run=False, rehearse=False, ledger=True):
    """マクロ 1 本の修理（または新規作成）を回す。戻り値 dict（done・ok・turns・report…）。

    ledger=False は走行台帳に残さない（実射＝練習台の記録を実戦の記録に混ぜない・2026-09-06）。
    """
    ai, model, api_key = va._ai_setup(ai, model)
    names = va._all_procedure_names(wb)
    if not macro:
        macro = _find_macro_in_request(request, names)
    create = bool(va._CREATE_WORDS.search(request or '')) and (not macro or macro not in names)
    if not macro and not create:
        print("エラー: 直すマクロが分かりません。--macro 名 で指定してください")
        if names:
            print("  マクロ: " + ", ".join(names[:20]) + ("  …" if len(names) > 20 else ""))
        return {'done': False, 'ok': False, 'turns': 0, 'report': '', 'usage': {}}
    if not create and names and macro not in names:
        print(f"エラー: マクロ '{macro}' が見つかりません")
        _suggest_similar(macro, names)
        return {'done': False, 'ok': False, 'turns': 0, 'report': '', 'usage': {}}
    if create:
        print(f"対象ブック: {wb.Name}   新しいマクロを作ります"
              + (f"（名前の指定: {macro}）" if macro else "（名前は依頼から AI が決める）")
              + "（この往復のあいだ固定。別のブックには触らない）")
        materials, got = _macro_materials_create(wb)
        target_module = None
        head = f"\n\n【材料】（ブック {wb.Name}。道具が読んだ現物）\n"
    else:
        print(f"対象ブック: {wb.Name}   修理するマクロ: {macro}（この往復のあいだ固定。別のブックには触らない）")
        materials, got, target_module = _macro_materials(macro, wb)
        head = f"\n\n【材料】（ブック {wb.Name} のマクロ「{macro}」。道具が読んだ現物）\n"
    needs_run = _needs_run(request) and not create      # 「結果が合わない」型＝直した後に動かして確かめる（2026-09-17 夜）
    expectations = _request_expectations(request)
    sm = _SHEET_IN_REQ_RE.search(request or '')
    sheet_hint = (sm.group(1) or sm.group(2)) if sm else None
    score = None                       # 最後に突き合わせた (○ の数, 期待の数)
    if needs_run and rehearse and expectations and not dry_run and got:
        # 直す前に 1 回撃って ○× を取る（症状の確かめ＋直しの後に ○ が増えたかを比べる土台・2026-09-17 夜。
        # 土台が無いと、最初の直しが外れでも「増えていない」と言えず、外れの直しが残った）
        ok_b, out_b = va._run_cmd(['rehearse', macro, '--timeout', '120'], wb)
        res_b, _v, score = _verify_results([(f"rehearse {macro}", ok_b, out_b)], None, expectations, sheet_hint, wb, None)
        txt = res_b[0][2] or ''
        run_line = next((l for l in txt.split('\n') if l.startswith('マクロ実行')), '（撃てませんでした）')
        i = txt.find('--- 依頼の期待')
        materials += ("\n--- 直す前の試し撃ち（道具が rehearse を 1 回撃った。依頼の症状の確かめ） ---\n"
                      + run_line[:300] + ("\n" + txt[i:].rstrip() if i >= 0 else "") + "\n")
        if not ok_b:
            score = None      # 撃てない（コンパイルエラー等）は比べる土台にしない＝コンパイルの直しを「関係なかった」と言わない
    print(materials)
    if not got:
        return {'done': False, 'ok': False, 'turns': 0, 'report': '', 'usage': {}}
    if not dry_run:
        # シートの控えは sheet のループでしか取れない。マクロを直した後の --undo が、その前の
        # sheet 仕事の控えでシートを上書きしないように印を付ける（2026-09-04）
        va._mark_undo_stale(f"macro（マクロ「{macro or '新規'}」を直した。戻すなら list-backups → restore）",
                          str(getattr(wb, 'Name', '') or ''))
    prompt = (MACRO_RULES + "\n【依頼】\n" + request.strip() + head + materials + "\n"
              + ("（rehearse は使える）\n" if rehearse else ""))
    with open(va._LAST_AGENT_ASK_FILE, 'w', encoding='utf-8') as f:
        f.write(prompt)
    log = open(va._LAST_AGENT_LOG_FILE, 'w', encoding='utf-8')
    # meta を書かないと、この記録が --continue で sheet の往復として読み戻されていた（2026-09-04）
    log.write(json.dumps({'meta': {'book': str(wb.Name), 'sheet': None, 'mode': 'macro', 'macro': macro,
                                   'request': request, 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
                                   'ai': ai, 'model': model, 'max_turns': max_turns}},
                         ensure_ascii=False) + "\n")
    history = []
    total = {'in': 0, 'out': 0, 'think': 0, 'sent': 0, 'recv': 0, 'sec': 0.0}
    done = False
    report = ''
    wrote_any = False
    last_compile_ok = None
    gates = dict.fromkeys(va._GATE_LABELS, 0)   # 関所の差し戻し（2026-09-06・シート側と同じ数え方）
    run_id = va._run_id()
    plan = []                          # やることの一覧（2026-09-06・シート側と同じ落とし物の関所）
    dropped = []                       # 最後まで「未」だった項目
    code_before = None                 # 最初の書き換えの前の全コード（頼んでいない書き換えを見るため）
    code_notes = []                    # 頼んでいないコードの変化
    code_checked = False               # 差し戻しは 1 回だけ
    verified = None                    # 最後の書き換えの後に動かして確かめたか（None＝まだ・True／False＝結果）
    verify_rejects = 0
    undo_stack = []                    # 直し 1 つごとの「直す前の姿」（{"op":"undo"} で 1 つずつ戻す）
    start_proc = None                  # 最初に直す前の姿（done の後に要らない直しを戻す土台）
    minimized = []
    turn = 0
    t_start = time.time()
    try:
        msg = prompt
        while turn < max_turns:
            turn += 1
            history.append(('user', msg))
            print(f"\n===== 往復 {turn}/{max_turns} =====")
            va.progress_write('ask', book=str(wb.Name), turn=turn, max_turns=max_turns,
                           note=f"マクロ「{macro or '新規'}」", request=request[:80])
            text, usage, sec = va._ask(ai, model, api_key, history)
            history.append(('model', text))
            for k in ('in', 'out'):
                total[k] += usage[k]
            total['think'] += usage.get('think', 0)
            total['sent'] += len(msg)
            total['recv'] += len(text)
            total['sec'] += sec
            print(va._bill_line(len(msg), text, usage, sec, model))
            entry = {'turn': turn, 'prompt': msg, 'reply': text, 'usage': usage, 'sec': round(sec, 2)}
            try:
                d = _extract_json(text)
                if not isinstance(d, dict) or not isinstance(d.get('actions') or [], list):
                    raise ValueError("形が違います（最上位がオブジェクトで actions が配列）")
            except Exception as ex:
                print(f"AI の返事が規約に合っていません: {ex}")
                entry['error'] = str(ex)
                gates['format'] += 1
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                msg = f"【結果】\n返事が規約に合っていません: {ex}\n規約どおりの JSON で返してください。（残りの往復: {max_turns - turn} 回）\n"
                continue
            actions = d.get('actions') or []
            done = bool(d.get('done')) and not actions
            report = str(d.get('report') or '')
            plan = va._merge_plan(plan, va._parse_plan(d.get('plan')))     # 項目は消させない（state だけ更新）
            if d.get('hypothesis'):
                print(f"仮説: {d['hypothesis']}")
            if d.get('say'):
                print(f"AI: {d['say']}")
            if actions:
                hands = " / ".join((a.get('op', '?') if isinstance(a, dict) else '?') for a in actions)
                print(f"手 {len(actions)} 本: " + hands)
                va.progress_write('run', book=str(wb.Name), turn=turn, max_turns=max_turns,
                               note=f"マクロ「{macro or '新規'}」", hands=f"手 {len(actions)} 本: {hands}")
            if dry_run:
                print(json.dumps(actions, ensure_ascii=False, indent=1))
                print("（--dry-run: ブックには触っていません）")
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                break
            if actions:
                if code_before is None:
                    # 最初の書き換えの前に全コードを控える（頼んでいない書き換えを終わりに数えるため）
                    code_before = _code_snapshot(wb)
                if any(isinstance(a, dict) and a.get('op') == 'undo' for a in actions):
                    results = _undo_step(undo_stack, wb, target_module)          # 直す前に戻す（この往復の手はそれだけ）
                    rest = [a for a in actions if not (isinstance(a, dict) and a.get('op') == 'undo')]
                    if rest and results[-1][1]:
                        results += _execute_macro(rest, rehearse, wb, existing=names, default_module=target_module)
                else:
                    if (start_proc is None and macro and not create
                            and any(isinstance(a, dict) and a.get('op') in ('replace', 'code_replace') for a in actions)):
                        start_proc = _proc_text(wb, macro, target_module)     # 直す前の姿（要らない直しを戻す土台）
                    results = _execute_macro(actions, rehearse, wb, existing=names, default_module=target_module,
                                             undo=(undo_stack, macro))
                results, verified, score = _verify_results(results, verified, expectations, sheet_hint, wb, score)
                for label, ok, out in results:
                    print(f"--- {label} {'' if ok else '（失敗）'}---")
                    print(out.rstrip())
                    if label.startswith('compile'):
                        last_compile_ok = ok
                    # add（新規作成）も書き換え＝「書き換えはしていません」と言わない（2026-09-04）
                    if ok and (label.startswith('replace') or label.startswith('code_replace')
                               or label.startswith('add ')):
                        wrote_any = True
                if any(not ok for _l, ok, _o in results):
                    gates['hand'] += 1
                entry['results'] = [{'label': l, 'ok': o, 'out': t} for l, o, t in results]
                log.write(json.dumps(entry, ensure_ascii=False) + "\n")
                msg = va._result_prompt(results, max_turns - turn)
                continue
            log.write(json.dumps(entry, ensure_ascii=False) + "\n")
            if done and turn < max_turns:
                # (1) 落とし物の関所（2026-09-06・シート側と同じ。「未」のまま done は受け付けない）
                pending = va._plan_pending(plan)
                if pending:
                    print("未の項目が残っているので done を受け付けません: " + " / ".join(pending))
                    done = False
                    gates['plan'] += 1
                    msg = ("【結果】\n未の項目が残っています: " + " / ".join(pending)
                           + "。やったなら「済」、できないなら「不可」にして report に理由を書いてください。"
                             "まだやっていないなら、その手を actions に並べてください。"
                             f"（残りの往復: {max_turns - turn} 回）\n")
                    continue
                # (2) 頼んでいないコードの変化（道具が撃つ前の控えと照らす・AI を使わない）。1 回だけ
                if code_before is not None and not code_checked:
                    code_checked = True
                    code_notes = _code_violations(code_before, _code_snapshot(wb),
                                                  allowed=[macro] if macro else [],
                                                  allow_new=create)
                    if code_notes:
                        print("頼んでいないコードの変化（done を受け付けません）:")
                        for n in code_notes:
                            print("  - " + n)
                        done = False
                        gates['inv'] += 1
                        msg = ("【結果】\n道具が「撃つ前のコード」と「いまのコード」を照らしました。"
                               "直すと言った以外のところが変わっています:\n"
                               + "\n".join("- " + n for n in code_notes)
                               + "\n元に戻してください（get でそのプロシージャを読み、replace で戻す）。"
                                 "code_replace の search が他のプロシージャにもある行に当たったのが原因のことが"
                                 "多いので、search は前後を足して一意にしてください。"
                                 "直す必要があってそうしたのなら、report に理由を書いてから done にしてください。"
                               + f"（残りの往復: {max_turns - turn} 回）\n")
                        continue
                    print("頼んでいないコードの変化: なし")
                # (3) 結果の依頼は、直した後に動かして確かめたか（2026-09-17 夜・AI を使わない）。2 回まで
                if needs_run and rehearse and wrote_any and verified is not True and verify_rejects < 2:
                    verify_rejects += 1
                    done = False
                    gates['verify'] += 1
                    why = ("直した後に一度も動かしていません" if verified is None else
                           "直した後の試し撃ちで、まだ依頼の症状が消えていません（上の ○× か実行の失敗）")
                    print(f"動かして確かめていないので done を受け付けません: {why}")
                    msg = ("【結果】\n" + why + "。この依頼は「結果が合わない／何も書かれない」なので、コンパイルが通っても"
                           "直ったとは言えません。rehearse でこのマクロを撃ち、依頼の症状が消えたかを確かめてから done にしてください。"
                           + (" 依頼に書かれた期待（" + "・".join(f"{a}＝{w or '空'}" for a, w in expectations)
                              + "）は道具が試し撃ちの後の値と突き合わせます。" if expectations else "")
                           + " 合わなければ、直した行が症状と本当に関係あるかを見直し、関係ない行を直していたら元に戻してください。"
                           + f"（残りの往復: {max_turns - turn} 回）\n")
                    continue
            if done:
                pending = va._plan_pending(plan)
                if pending:                       # 最後の往復＝続きを頼めない。落とし物のまま合格にしない
                    done = False
                    dropped = pending
                break
            gates['empty'] += 1
            msg = ("【結果】\nactions が空で done でもありません。続きの actions を返すか、終わりなら "
                   f"done を true にして report を書いてください。（残りの往復: {max_turns - turn} 回）\n")
    finally:
        log.close()
    print()
    if plan:
        print(va._plan_line(plan))
    if done:
        print("AI の報告:\n" + (report.strip() or "（なし）"))
    elif dropped:
        print("落とし物: 未の項目が残ったまま往復が尽きました（{} 件）: ".format(len(dropped))
              + " / ".join(dropped) + "\n  → --max-turns を増やしてもう一度回してください。")
        print("AI の報告:\n" + (report.strip() or "（なし）"))
    elif not dry_run:
        print(f"往復の上限（{max_turns} 回）に達しました。AI は done を言っていません。")
    ok = True
    if done and not dry_run and rehearse and expectations and verified is True and start_proc:
        minimized = _minimize_fix(wb, macro, target_module, start_proc, expectations, sheet_hint)
        for n in minimized:
            print(n)
        if minimized:
            report = (report.rstrip() + "\n（道具: " + " ／ ".join(minimized) + "）").strip()
    if not dry_run:
        if code_before is not None:
            # 関所の後にも直しているので、終わりにもう一度照らす（直っていれば「なし」になる）
            code_notes = _code_violations(code_before, _code_snapshot(wb),
                                          allowed=[macro] if macro else [], allow_new=create)
            if code_notes:
                print(f"頼んでいないコードの変化: {len(code_notes)} 件"
                      "（直すと言った以外のところが変わっています。戻すなら list-backups → restore）")
                for n in code_notes:
                    print("  - " + n)
            else:
                print("頼んでいないコードの変化: なし（他のプロシージャ・宣言部）")
        ok_c, out = va._run_cmd(['compile'], wb)
        print("検査: 全体コンパイル " + ("通過" if ok_c else "不通過") + ("" if ok_c else "\n" + out.rstrip()))
        ok = ok_c and not code_notes
    print(va._gate_line(gates) + f"　＝ {va._gate_grade(ok and done, gates)}")
    print(va._bill_line(total['sent'], 'x' * total['recv'], total, total['sec'], model,
                     head=f"合計（往復 {turn} 回）"))
    print(f"経過: {time.time() - t_start:.1f} 秒")
    print(f"記録: {va._LAST_AGENT_LOG_FILE}（往復の全文）")
    if wrote_any:
        print("（書き換えはバックアップつきで保存済み。戻すなら list-backups → restore）")
    elif not dry_run:
        print("（書き換えはしていません）")
    va.progress_write('done' if (done or dry_run) else 'fail', book=str(wb.Name), turn=turn,
                   max_turns=max_turns, note=f"マクロ「{macro or '新規'}」",
                   sec=time.time() - t_start)
    r = {'done': done, 'ok': ok and (done or dry_run), 'turns': turn, 'report': report, 'usage': total,
         'wrote': wrote_any, 'gates': gates, 'plan': plan, 'dropped': dropped, 'run_id': run_id,
         'inv': code_notes, 'verified': verified, 'minimized': minimized}
    if ledger and not dry_run:
        # path は渡さない＝macro は道具が自分で保存する（「人が保存した」の印にならない）
        va._runs_record(run_id, 'macro', str(wb.Name), None, request, r, time.time() - t_start)
    return r


__all__ = ['BOTH_RULES', 'MACRO_RULES', 'run_both', 'run_macro_agent']
