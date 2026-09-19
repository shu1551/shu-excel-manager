# -*- coding: utf-8 -*-
"""vbam_form2vba.py — vba_manager 分割パート: UserForm を「VBAだけで組み立て直す作成マクロ」に変換

用途:
  .frx はバイナリで、メールの添付検査に弾かれることがある（2026-08-17）。
  フォームは結局「位置・大きさ・属性・コード」の集まりなので、全部テキストの
  VBA に書き下せる。出来た作成マクロを標準モジュールに貼って一回実行すれば、
  同じフォームがその場で組み上がる。持ち出しに .frx が要らなくなる。

ここに畳んである罠（どれも 2026-08-17 に実機で踏んで分かったもの。
文章で残すと次に書く人がまた踏むので、コード側に固定してある）:

  1. 削除した名前は Excel を閉じるまで VBE が握る。
     → 作り直しは Remove ではなく「改名して逃がす → 作る → 古い方を消す」。
        Remove してから同じ名前で Add すると、その実行でも次の実行でも失敗する。
  2. CodeModule.AddFromString は末尾に足さない（宣言部の直後に差し込む）。
     → 分割して注入すると後半が前半の頭にめり込み「プロシージャの外では無効です」。
        InsertLines(CountOfLines + 1, s) で末尾に追記する。
  3. VBA の1プロシージャは 64KB まで。
     → コードが大きいフォームは注入を分け、本体から「〜つづき2, 3…」を順に呼ぶ。
  4. リストボックスの高さは行数の倍数に丸められる。
     → フォントを決めてから大きさを当て、IntegralHeight を外し、
        全部そろえた後にもう一度当て直す。MultiPage の表示中ページにある
        リストは切り替えのたびに丸め直されるので、最後に当てる。
  5. 既定と違う属性は種類ごとに違う（Style, TabStop, TextAlign, ColumnWidths …）。
     → 素のコントロールを一時フォームに作って既定値を採り、差分だけ書き出す。
        個別に列挙すると必ず取りこぼす（MultiPage の Style=2 を落とした実例あり）。

  ※ 表示中のフォームは Designer が取れない。書き出す前に閉じておくこと。
"""
import io
import os
import re
import sys
import argparse
import difflib

from vbam_core import *  # noqa: F401,F403


# コントロール種別 → ProgID
_F2V_PROGID = {
    "TextBox": "Forms.TextBox.1",
    "Label": "Forms.Label.1",
    "CommandButton": "Forms.CommandButton.1",
    "ListBox": "Forms.ListBox.1",
    "ComboBox": "Forms.ComboBox.1",
    "CheckBox": "Forms.CheckBox.1",
    "OptionButton": "Forms.OptionButton.1",
    "ToggleButton": "Forms.ToggleButton.1",
    "Frame": "Forms.Frame.1",
    "MultiPage": "Forms.MultiPage.1",
    "Image": "Forms.Image.1",
    "SpinButton": "Forms.SpinButton.1",
    "TabStrip": "Forms.TabStrip.1",
    "ScrollBar": "Forms.ScrollBar.1",
}

# COM の GetTypeInfo は環境によりインターフェース名を返す（form_inspect と同じ正規化）
_F2V_NORM = {
    "ILabelControl": "Label", "IMdcText": "TextBox", "IMdcCombo": "ComboBox",
    "IMdcList": "ListBox", "IMdcCheckBox": "CheckBox", "IMdcOptionButton": "OptionButton",
    "IMdcToggleButton": "ToggleButton", "ICommandButton": "CommandButton",
    "IFrame": "Frame", "IMultiPage": "MultiPage", "IPage": "Page",
    "IImage": "Image", "ISpinbutton": "SpinButton", "ITabStrip": "TabStrip",
    "IScrollbar": "ScrollBar",
}

# 既定値と突き合わせて書き出す属性（幾何・フォント・名前・タブ順は別扱い）
_F2V_SWEEP = [
    "Style", "BackColor", "ForeColor", "BorderColor", "BorderStyle", "SpecialEffect",
    "TextAlign", "WordWrap", "AutoSize", "ControlTipText", "Enabled", "Locked",
    "Visible", "TabStop", "MultiLine", "ScrollBars", "MaxLength", "PasswordChar",
    "EnterKeyBehavior", "TabKeyBehavior", "AutoTab", "SelectionMargin", "HideSelection",
    "ListStyle", "ColumnCount", "ColumnWidths", "ColumnHeads", "BoundColumn",
    "TextColumn", "MultiSelect", "MatchEntry", "MatchRequired",
    "DropButtonStyle", "ShowDropButtonWhen", "ListRows", "Accelerator",
    "TripleState", "GroupName", "Alignment", "Orientation", "Min", "Max",
    "SmallChange", "LargeChange", "Delay", "TabOrientation", "MultiRow",
    "PictureAlignment", "PictureSizeMode", "PictureTiling", "Zoom", "Cycle",
    "KeepScrollBarsVisible", "ScrollHeight", "ScrollWidth", "DragBehavior",
    "EnterFieldBehavior", "MousePointer", "AutoWordSelect",
]

# フォーム自身の属性のうち、書き出さないもの（読み取り専用・COMオブジェクト・別扱い）
_F2V_SKIP_FORM = {
    "Caption", "Width", "Height", "StartUpPosition", "Name", "Left", "Top",
    "Font", "Picture", "MouseIcon", "Tag",
    "ActiveControl", "Controls", "Selected", "ShowToolbox",
    "InsideWidth", "InsideHeight",
}

# 1プロシージャに入れるコード行の文字数の目安（VBA の上限 64KB に対する余裕込み）
_F2V_CHUNK_LIMIT = 28000


def _f2v_q(s):
    """VBA の文字列リテラルにする（" は "" に）"""
    return '"' + str(s).replace('"', '""') + '"'


def _f2v_lit(v):
    """Python の値を VBA のリテラルにする"""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        if isinstance(v, int) and (v > 0x7FFFFFF or v < -0x7FFFFFF):
            return "&H%X" % (v & 0xFFFFFFFF)
        return "%g" % v
    return _f2v_q(v)


def _f2v_ctl_type(c):
    """コントロールの種別名を返す（不明なら None）"""
    try:
        n = c._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    except Exception:
        return None
    n = _F2V_NORM.get(n, n)
    if n in _F2V_PROGID:
        return n
    for k in _F2V_PROGID:
        if k.lower() in str(n).lower():
            return k
    return None


def _f2v_defaults(wb):
    """一時フォームに各種コントロールを1個ずつ置いて既定値を採る（罠5）"""
    tmp = wb.VBProject.VBComponents.Add(3)
    # 直前に消した名前は VBE がまだ握っているので、通る名前が出るまでずらす（罠1）
    for i in range(1, 100):
        try:
            tmp.Name = "tmpF2VDefaults%d" % i
            break
        except Exception:
            continue
    d = {}
    try:
        td = tmp.Designer
        for t, pid in _F2V_PROGID.items():
            try:
                c = td.Controls.Add(pid, "tmp" + t.lower())
            except Exception:
                continue
            props = {}
            for p in _F2V_SWEEP:
                try:
                    props[p] = c.__getattr__(p)
                except Exception:
                    pass
            d[t] = props
        d["_form"] = {}
        for pr in tmp.Properties:
            try:
                d["_form"][pr.Name] = pr.Value
            except Exception:
                pass
    finally:
        try:
            wb.VBProject.VBComponents.Remove(tmp)
        except Exception:
            pass
    return d


def build_form_builder(wb, src_form, new_form=None, sub_name=None):
    """UserForm を「組み立て直す作成マクロ」の VBA テキストにして返す。

    戻り値: (テキスト, 分割本数)
    """
    new_form = new_form or src_form
    sub_name = sub_name or (new_form + "作成")

    DEF = _f2v_defaults(wb)
    vbc = wb.VBProject.VBComponents(src_form)
    dsn = vbc.Designer
    if dsn is None:
        raise RuntimeError(
            "フォーム '%s' の Designer が取れません。"
            "表示中のフォームは取れないので、閉じてから実行してください。" % src_form)

    # --- フォーム自身のコード（自分の名前での参照は Me. に直して名前依存を外す）
    cm = vbc.CodeModule
    code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
    code = code.replace(src_form + ".", "Me.")
    code_lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    while code_lines and not code_lines[-1].strip():
        code_lines.pop()

    out = []
    a = out.append
    list_heights = []
    page_ctx = [None]

    def prop(nm, default=None):
        try:
            return vbc.Properties(nm).Value
        except Exception:
            return default

    a("Sub %s()" % sub_name)
    a("    '★VBAだけで「%s」を組み立て直す（一回実行すればフォームが出来る）" % src_form)
    a("    Const フォーム名 As String = %s" % _f2v_q(new_form))
    a("    Dim vbp As Object")
    a("    Dim vbc As Object")
    a("    Dim dsn As Object")
    a("    Dim c As Object")
    a("    Dim p As Object")
    a("    Dim 既存 As Object")
    a("    Dim i As Long")
    a("    Dim s As String")
    a("")
    a("    On Error Resume Next")
    a("    Set vbp = ThisWorkbook.VBProject")
    a("    On Error GoTo 0")
    a("    If vbp Is Nothing Then")
    a('        MsgBox "Excel のオプション → トラスト センター → トラスト センターの設定 → マクロの設定 で" & vbCrLf & _')
    a('               "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れてから、もう一度実行してください。", _')
    a("               vbExclamation")
    a("        Exit Sub")
    a("    End If")
    a("")
    a("    ' 同名のフォームがあれば、消さずに改名して逃がす。")
    a("    ' 削除した名前は Excel を閉じるまで VBE が握ったままで、その名前では作り直せない。")
    a("    ' 改名なら名前がすぐ解放されるので、この実行の中で作り直せる（最後に古い方を消す）。")
    a("    On Error Resume Next")
    a("    Set 既存 = vbp.VBComponents(フォーム名)")
    a("    On Error GoTo 0")
    a("    If Not 既存 Is Nothing Then")
    a("        For i = 1 To 999")
    a("            Err.Clear")
    a("            On Error Resume Next")
    a('            既存.Name = "旧" & フォーム名 & i')
    a("            On Error GoTo 0")
    a("            If Err.Number = 0 Then Exit For")
    a("        Next i")
    a("        If Err.Number <> 0 Then")
    a('            MsgBox "古い「" & フォーム名 & "」を逃がせませんでした。" & vbCrLf & _')
    a('                   "VBE でそのフォームを閉じてから、もう一度実行してください。", vbExclamation')
    a("            Exit Sub")
    a("        End If")
    a("    End If")
    a("")
    a("    ' フォーム本体（3 = UserForm）")
    a("    Set vbc = vbp.VBComponents.Add(3)")
    a("    vbc.Name = フォーム名")
    a("    ' VBE の「変数の宣言を強制する」が入っていると、新しいフォームに Option Explicit が先に入ってくる。")
    a("    ' 元のコードにも同じ行があると二重になるので、空にしてから中身を入れる")
    a("    With vbc.CodeModule")
    a("        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines")
    a("    End With")
    a("    vbc.Properties(\"Caption\") = %s" % _f2v_q(prop("Caption", src_form)))
    a("    vbc.Properties(\"StartUpPosition\") = %d" % int(prop("StartUpPosition", 1)))
    a("    vbc.Properties(\"Width\") = %g" % float(prop("Width", 300)))
    a("    vbc.Properties(\"Height\") = %g" % float(prop("Height", 300)))
    try:
        form_props = list(vbc.Properties)
    except Exception:
        form_props = []
    for pr in form_props:
        try:
            nm, v = pr.Name, pr.Value
        except Exception:
            continue
        if nm in _F2V_SKIP_FORM or v is None:
            continue
        if not isinstance(v, (bool, int, float, str)):
            continue
        if nm in DEF.get("_form", {}) and DEF["_form"][nm] != v:
            a('    vbc.Properties(%s) = %s' % (_f2v_q(nm), _f2v_lit(v)))
    a("    Set dsn = vbc.Designer")
    a("    ' 内寸 %s x %s（元フォームと同じ）" % (dsn.InsideWidth, dsn.InsideHeight))
    a("")

    # --- コントロールを集める（親子の判定つき）
    items = []
    for c in dsn.Controls:
        t = _f2v_ctl_type(c)
        parent_name = None
        try:
            pn = c.Parent.Name
            if pn != src_form:
                parent_name = pn
        except Exception:
            pass
        items.append((c, t, parent_name))

    def geom(c, var):
        a('    %s.Left = %g: %s.Top = %g: %s.Width = %g: %s.Height = %g'
          % (var, c.Left, var, c.Top, var, c.Width, var, c.Height))

    def emit(c, t, var, adder="dsn"):
        f = c.Font
        a('    Set %s = %s.Controls.Add(%s, %s)' % (var, adder, _f2v_q(_F2V_PROGID[t]), _f2v_q(c.Name)))
        geom(c, var)
        a('    %s.Font.Name = %s: %s.Font.Size = %g' % (var, _f2v_q(f.Name), var, f.Size))
        if getattr(f, "Bold", False):
            a('    %s.Font.Bold = True' % var)
        # フォントを変えると行の高さが変わり、リストは行数の倍数に丸め直される（罠4）
        ih = None
        if t in ("ListBox", "ComboBox"):
            try:
                ih = bool(c.IntegralHeight)
            except Exception:
                ih = None
        if ih:
            a('    %s.IntegralHeight = False' % var)
            list_heights.append((c.Name, c.Height, page_ctx[0]))
        geom(c, var)
        try:
            if t in ("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Frame"):
                a('    %s.Caption = %s' % (var, _f2v_q(c.Caption)))
        except Exception:
            pass
        if t == "CommandButton":
            try:
                if c.Default:
                    a('    %s.Default = True' % var)
                if c.Cancel:
                    a('    %s.Cancel = True' % var)
            except Exception:
                pass
        # 既定と違う属性だけを機械的に拾う（種類ごとの取りこぼしを防ぐ・罠5）
        base = DEF.get(t, {})
        for pname in _F2V_SWEEP:
            if pname not in base or pname == "IntegralHeight":
                continue
            try:
                v = c.__getattr__(pname)
            except Exception:
                continue
            if v is None or v == base[pname]:
                continue
            a('    %s.%s = %s' % (var, pname, _f2v_lit(v)))

    def _ti(x):
        try:
            return int(x[0].TabIndex)
        except Exception:
            return 999

    tops = [(c, t, pn) for (c, t, pn) in items if pn is None and t is not None]
    tops.sort(key=_ti)
    tab_order = []

    for c, t, parent_name in tops:
        a("    ' %s" % c.Name)
        emit(c, t, "c")
        try:
            tab_order.append((int(c.TabIndex), c.Name))
        except Exception:
            pass
        if t == "MultiPage":
            n_pages = c.Pages.Count
            a('    Set p = c')
            a('    Do While p.Pages.Count > %d' % n_pages)
            a('        p.Pages.Remove p.Pages.Count - 1')
            a('    Loop')
            a('    Do While p.Pages.Count < %d' % n_pages)
            a('        p.Pages.Add')
            a('    Loop')
            for i, pg in enumerate(c.Pages):
                a('    p.Pages(%d).Name = %s: p.Pages(%d).Caption = %s'
                  % (i, _f2v_q(pg.Name), i, _f2v_q(pg.Caption)))
            for pg_i, pg in enumerate(c.Pages):
                kids = [(k, kt, pn) for (k, kt, pn) in items if pn == pg.Name and kt]
                if not kids:
                    continue
                kids.sort(key=_ti)
                a('')
                for k, kt, pn in kids:
                    a("    ' %s（%s ページ）" % (k.Name, pg.Caption))
                    page_ctx[0] = (c.Name, pg_i, n_pages)
                    emit(k, kt, "c", adder="p.Pages(%d)" % pg_i)
                    page_ctx[0] = None
                    a('')
        a("")

    if tab_order:
        a("    ' タブ順（作った順で既に揃うが、念のため小さい方から明示する）")
        for ti, nm in sorted(tab_order):
            a('    dsn.Controls(%s).TabIndex = %d' % (_f2v_q(nm), ti))
        a("")

    if list_heights:
        a("    ' リストの高さ（行数の倍数への丸めを避けるため最後に当て直す）")
        a("    ' 表示中のページに載っているリストは当てた瞬間に丸め直される。")
        a("    ' そのリストのページを一度隠してから当てると、戻しても残る。")
        for nm, h, ctx in list_heights:
            if ctx:
                mp, pi, npg = ctx
                if npg > 1:
                    a('    dsn.Controls(%s).Value = %d' % (_f2v_q(mp), (pi + 1) % npg))
                    a('    DoEvents      ' + "' ページの切り替えを効かせてから当てる")
            a('    dsn.Controls(%s).Height = %g' % (_f2v_q(nm), h))
        for mp in sorted({ctx[0] for _, _, ctx in list_heights if ctx}):
            a('    dsn.Controls(%s).Value = 0      ' % _f2v_q(mp) + "' 既定のページへ戻す")
        a("")

    # --- コード注入（罠2・罠3）
    chunks = [[]]
    size = 0
    for ln in code_lines:
        emitted = '    s = s & %s & vbCrLf\n' % _f2v_q(ln)
        if size + len(emitted) > _F2V_CHUNK_LIMIT and chunks[-1]:
            chunks.append([])
            size = 0
        chunks[-1].append(ln)
        size += len(emitted)
    # かたまりの最後が空行だと、継ぎ目の空行が「元からあったもの」なのか
    # 追記でできた余りなのか区別できなくなる。空行は次のかたまりへ送る
    for i2 in range(len(chunks) - 1):
        while chunks[i2] and not chunks[i2][-1].strip():
            chunks[i2 + 1].insert(0, chunks[i2].pop())

    def emit_chunk(lines):
        a("    ' フォームの中身")
        a('    s = ""')
        # 最後の1行に vbCrLf を付けない。付けると末尾に空行が残り、
        # 次のかたまりを追記したときに継ぎ目へ空行が二重に入る
        for i2, ln in enumerate(lines):
            if i2 == len(lines) - 1:
                a('    s = s & %s' % _f2v_q(ln))
            else:
                a('    s = s & %s & vbCrLf' % _f2v_q(ln))
        a("")
        a("    ' 末尾に追記する（AddFromString は宣言部の直後へ差し込むので分割注入に使えない）")
        a("    ' 継ぎ目に余分な空行が残るので、追記の前に末尾の空行を落とす")
        a("    With vbc.CodeModule")
        a("        Do While .CountOfLines > 0")
        a("            If Len(Trim(.Lines(.CountOfLines, 1))) > 0 Then Exit Do")
        a("            .DeleteLines .CountOfLines")
        a("        Loop")
        a("        .InsertLines .CountOfLines + 1, s")
        a("    End With")

    emit_chunk(chunks[0])
    a("")
    for n in range(2, len(chunks) + 1):
        a("    ' コードの続き（1プロシージャ 64KB の上限に収めるため分けている）")
        a("    Call %sつづき%d" % (sub_name, n))
    if len(chunks) > 1:
        a("")
    a("    ' 逃がしておいた古いフォームを片付ける（新しい方が出来上がってから）")
    a("    If Not 既存 Is Nothing Then vbp.VBComponents.Remove 既存")
    a("End Sub")

    for n in range(2, len(chunks) + 1):
        a("")
        a("Sub %sつづき%d()" % (sub_name, n))
        a("    ' 「%s」の続き。単独では使わない。" % sub_name)
        a("    Dim vbc As Object")
        a("    Dim s As String")
        a("    Set vbc = ThisWorkbook.VBProject.VBComponents(%s)" % _f2v_q(new_form))
        emit_chunk(chunks[n - 1])
        a("End Sub")

    return "\n".join(out) + "\n", len(chunks)


def _f2v_verify(xl, wb, vba_text, form_name):
    """捨てブックに作成マクロを入れて実行し、元と突き合わせる（元ブックには触らない）"""
    tmp_path = os.path.join(SCRIPT_DIR, "_f2v_試作_%s.xlsm" % form_name)
    if os.path.exists(tmp_path):
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    src_cm = wb.VBProject.VBComponents(form_name).CodeModule
    orig = (src_cm.Lines(1, src_cm.CountOfLines) if src_cm.CountOfLines else "")
    orig = orig.replace(form_name + ".", "Me.")
    ref_ctl = [(c.Name, round(c.Left, 1), round(c.Top, 1), round(c.Width, 1), round(c.Height, 1))
               for c in wb.VBProject.VBComponents(form_name).Designer.Controls]

    ok = True
    tw = xl.Workbooks.Add()
    try:
        tw.SaveAs(tmp_path, 52)          # 52 = xlOpenXMLWorkbookMacroEnabled
        md = tw.VBProject.VBComponents.Add(1)
        md.Name = "組立"
        md.CodeModule.AddFromString(vba_text)
        subs = [l.strip()[4:].split("(")[0] for l in vba_text.split("\n") if l.startswith("Sub ")]
        xl.Run("'%s'!組立.%s" % (tw.Name, subs[0]))

        got_cm = tw.VBProject.VBComponents(form_name).CodeModule
        got = got_cm.Lines(1, got_cm.CountOfLines) if got_cm.CountOfLines else ""
        a1 = [x for x in orig.replace("\r\n", "\n").split("\n")]
        b1 = [x for x in got.replace("\r\n", "\n").split("\n")]
        while a1 and not a1[-1].strip():
            a1.pop()
        while b1 and not b1[-1].strip():
            b1.pop()
        # VBA は識別子の大小文字を勝手に揃える（.lines→.Lines 等）。中身の違いではないので
        # 照合は小文字化して行い、大小文字だけの差は「整形」として数える
        d_real = [x for x in difflib.unified_diff(
            [s.lower() for s in a1], [s.lower() for s in b1], n=0, lineterm="")]
        n_real = len([x for x in d_real if x.startswith(('+', '-')) and not x.startswith(('+++', '---'))])
        print("  コード  : 元 %d 行 / 組立 %d 行 → %s"
              % (len(a1), len(b1), "一致" if n_real == 0 else "差分 %d 行" % n_real))
        if n_real:
            ok = False
            for x in d_real[:20]:
                print("      " + x)

        got_ctl = [(c.Name, round(c.Left, 1), round(c.Top, 1), round(c.Width, 1), round(c.Height, 1))
                   for c in tw.VBProject.VBComponents(form_name).Designer.Controls]
        ng = [x for x in ref_ctl if x not in got_ctl]
        print("  コントロール: 元 %d 個 / 組立 %d 個 → %s"
              % (len(ref_ctl), len(got_ctl), "全一致" if not ng else "不一致 %d 個" % len(ng)))
        for x in ng[:10]:
            print("      " + str(x))
        if ng:
            ok = False
    finally:
        try:
            tw.Close(SaveChanges=False)
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
    return ok


def cmd_form_to_vba(args):
    """UserForm を作成マクロ(.vba)に書き出す: form-to-vba [excel_file] <フォーム名>

      form-to-vba ワークシート一覧                  → _ワークシート一覧_作成マクロ.vba
      form-to-vba ワークシート一覧 --add shu002      → 書き出して shu002 に入れるまで
      form-to-vba --all --add shu002                → 全 UserForm を一括
      form-to-vba AI作業窓フォーム --verify           → 捨てブックで組み立てて元と照合
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    all_opt = getattr(args, "all_opt", False)
    if not rest and not all_opt:
        print("使い方: form-to-vba [excel_file] <フォーム名> [--all] [--out f.vba] [--add モジュール名] [--verify]")
        print("  別名: フォーム書き出し")
        return False

    xl, wb = get_workbook(target_file)

    forms = [c.Name for c in wb.VBProject.VBComponents if int(c.Type) == 3]
    if all_opt:
        targets = forms
    else:
        name = rest[0]
        hit = [f for f in forms if f.lower() == name.lower()]
        if not hit:
            print(f"エラー: UserForm '{name}' が見つかりません")
            print("  存在するフォーム: " + ", ".join(forms))
            return False
        targets = hit

    out_opt = getattr(args, "out_opt", None)
    if out_opt and len(targets) > 1:
        print("エラー: --out は1フォームのときだけ使えます（--all との併用は不可）")
        return False

    add_mod = getattr(args, "add_opt", None)
    do_verify = getattr(args, "verify", False)
    new_name = getattr(args, "name_opt", None)
    sub_opt = getattr(args, "sub_opt", None)

    all_ok = True
    for form in targets:
        print(f"=== {form} ===")
        try:
            text, n_chunk = build_form_builder(
                wb, form,
                new_name if len(targets) == 1 else None,
                sub_opt if len(targets) == 1 else None)
        except Exception as e:
            print(f"  エラー: 書き出せませんでした: {e}")
            all_ok = False
            continue

        out_path = out_opt or os.path.join(SCRIPT_DIR, f"_{form}_作成マクロ.vba")
        with io.open(out_path, "w", encoding="utf-8") as fh:
            fh.write(text)
        n_lines = text.count("\n")
        print(f"  出力: {os.path.basename(out_path)}  ({n_lines}行 {len(text)}文字"
              + (f" / {n_chunk}本に分割" if n_chunk > 1 else "") + ")")

        if do_verify:
            if not _f2v_verify(xl, wb, text, form):
                all_ok = False

        if add_mod:
            # 追加は既存の add-procedure に通す（同名重複ガード・バックアップ・保存が乗る）
            from vbam_vba import cmd_add_procedure
            ns = argparse.Namespace(
                posargs=([target_file] if target_file else []) + [add_mod],
                code_file_opt=out_path, yes=True, force=False)
            if not cmd_add_procedure(ns):
                all_ok = False

    return all_ok
