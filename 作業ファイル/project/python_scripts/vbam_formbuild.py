# -*- coding: utf-8 -*-
"""vbam_formbuild.py — vba_manager 分割パート: 「フォームを組む」コマンド（form-build）

設計図（JSON）を受け取り、form_layout モジュールを通して UserForm を組み立てる。
フォームの自動生成・再構築・プレビューを 1 コマンドで完結させる。
"""
import os
import sys
import json
import time
import argparse

from vbam_core import (SCRIPT_DIR, get_workbook, parse_target_and_rest, job_clock_note,
                       call_log_write, BACKUP_DIR)
import form_layout
from form_builder import FormBuilder

_LAST_FORM_SPEC_FILE = os.path.join(SCRIPT_DIR, '_last_form_spec.json')
_LAST_FORM_PNG = os.path.join(SCRIPT_DIR, '_last_form.png')

SAMPLE_SPEC = {
    "name": "F_顧客入力",
    "caption": "顧客の登録",
    "width": None,
    "rows": [
        {"heading": "顧客"},
        {"row": [{"lbl": "名前"}, {"txt": "txtName", "required": True}]},
        {"row": [{"lbl": "区分"}, {"combo": "cmbKind", "items": ["法人", "個人"]}]},
        {"row": [{"lbl": "登録日"}, {"txt": "txtDate"}]},
        {"spacer": None},
        {"button_bar": [{"ok": "btnOK", "caption": "登録"}, {"cancel": "btnCancel", "caption": "閉じる"}]}
    ]
}

KNOWN_PARTS = {
    "heading", "lbl", "txt", "combo", "lst", "chk", "opt", "opt_group",
    "btn", "ok", "cancel", "row", "button_bar", "spacer", "frame", "multipage", "page"
}


def _validate_and_parse_spec(spec_data):
    """JSON 設計図を検査し、form_layout の rows 構造へ変換する。
    未知の部品名や引数の不足があれば、明確なメッセージとともに ValueError を投げる。
    """
    if not isinstance(spec_data, dict):
        raise ValueError("設計図の最上位はオブジェクト（dict）である必要があります。")

    form_name = spec_data.get("name") or spec_data.get("form_name")
    if not form_name:
        raise ValueError("設計図にフォーム名 'name' が指定されていません。")

    caption = spec_data.get("caption") or spec_data.get("title", form_name)
    width = spec_data.get("width")
    rows_def = spec_data.get("rows")
    if rows_def is None and "fields" in spec_data:
        # fields 形式の互換対応
        rows_def = []
        for f in spec_data["fields"]:
            c_type = f.get("control", "txt")
            f_name = f.get("name", "field")
            lbl_text = f.get("label", f_name)
            req = f.get("required", False)
            if c_type in ("textbox", "txt"):
                rows_def.append({"row": [{"lbl": lbl_text}, {"txt": f_name, "required": req}]})
            elif c_type in ("combobox", "combo"):
                rows_def.append({"row": [{"lbl": lbl_text}, {"combo": f_name, "items": f.get("items", [])}]})
            elif c_type in ("checkbox", "chk"):
                rows_def.append({"row": [{"chk": f_name, "caption": lbl_text}]})
            elif c_type in ("listbox", "lst"):
                rows_def.append({"row": [{"lbl": lbl_text}, {"lst": f_name, "items": f.get("items", [])}]})
            else:
                rows_def.append({"row": [{"lbl": lbl_text}, {"txt": f_name}]})
        if "buttons" in spec_data:
            btn_bar = []
            for b in spec_data["buttons"]:
                b_action = b.get("action")
                if b_action == "submit" or b.get("default"):
                    btn_bar.append({"ok": b.get("name", "btnOK"), "caption": b.get("caption", "OK")})
                elif b_action == "cancel" or b.get("cancel"):
                    btn_bar.append({"cancel": b.get("name", "btnCancel"), "caption": b.get("caption", "閉じる")})
                else:
                    btn_bar.append({"btn": b.get("name", "btn"), "caption": b.get("caption", "ボタン")})
            rows_def.append({"button_bar": btn_bar})
    if not isinstance(rows_def, list):
        raise ValueError("設計図の 'rows' はリストである必要があります。")

    parsed_rows = []
    for idx, r_item in enumerate(rows_def, start=1):
        parsed = _parse_row_item(r_item, line_no=idx)
        parsed_rows.append(parsed)

    meta = {'theme': spec_data.get('theme', '秀'), 'assist': spec_data.get('assist', True)}
    return form_name, caption, width, parsed_rows, meta


def _parse_row_item(item, line_no):
    if not isinstance(item, dict):
        raise ValueError(f"{line_no} 行目の項目がオブジェクトではありません。")

    matched_key = None
    for k in item.keys():
        if k in KNOWN_PARTS:
            matched_key = k
            break

    if not matched_key:
        unknown_keys = list(item.keys())
        first_uk = unknown_keys[0] if unknown_keys else "空の要素"
        parts_list = ", ".join(sorted(KNOWN_PARTS))
        raise ValueError(f"{line_no} 行目の {first_uk} が分からない（使える部品: {parts_list}）")

    val = item[matched_key]

    if matched_key == "heading":
        if isinstance(val, str):
            return form_layout.heading(val)
        elif isinstance(val, dict):
            return form_layout.heading(val.get("caption", ""), font=val.get("font"), name=val.get("name"))
        else:
            return form_layout.heading(str(val))

    elif matched_key == "spacer":
        h = val if isinstance(val, (int, float)) else None
        return form_layout.spacer(height=h)

    elif matched_key == "row":
        if not isinstance(val, list):
            raise ValueError(f"{line_no} 行目の row の中身はリストである必要があります。")
        elems = []
        for elem_idx, elem_def in enumerate(val, start=1):
            elems.append(_parse_elem_item(elem_def, f"{line_no} 行目の {elem_idx} 番目"))
        return form_layout.row(*elems)

    elif matched_key == "button_bar":
        if not isinstance(val, list):
            raise ValueError(f"{line_no} 行目の button_bar の中身はリストである必要があります。")
        btns = []
        for b_idx, b_def in enumerate(val, start=1):
            btns.append(_parse_elem_item(b_def, f"{line_no} 行目の button_bar の {b_idx} 番目"))
        return form_layout.button_bar(*btns)

    elif matched_key == "frame":
        caption = item.get("caption", val if isinstance(val, str) else "")
        f_rows = item.get("rows", [])
        if not isinstance(f_rows, list):
            raise ValueError(f"{line_no} 行目の frame の rows はリストである必要があります。")
        parsed_f_rows = [_parse_row_item(fr, line_no=f"{line_no}.{fri+1}") for fri, fr in enumerate(f_rows)]
        return form_layout.frame(caption, *parsed_f_rows, name=item.get("name"))

    elif matched_key == "multipage":
        mp_name = val if isinstance(val, str) else item.get("name", "mpMain")
        pages_def = item.get("pages", [])
        parsed_pages = []
        for p_idx, p_item in enumerate(pages_def, start=1):
            p_cap = p_item.get("caption", p_item.get("page", f"Page{p_idx}"))
            p_rows = p_item.get("rows", [])
            parsed_p_rows = [_parse_row_item(pr, line_no=f"{line_no}.{p_idx}.{pri+1}") for pri, pr in enumerate(p_rows)]
            parsed_pages.append(form_layout.page(p_cap, *parsed_p_rows))
        return form_layout.multipage(mp_name, *parsed_pages)

    else:
        elem = _parse_elem_item(item, str(line_no))
        return form_layout.row(elem)


def _parse_elem_item(elem, pos_desc):
    if not isinstance(elem, dict):
        raise ValueError(f"{pos_desc} の要素がオブジェクトではありません。")

    matched_key = None
    for k in elem.keys():
        if k in KNOWN_PARTS:
            matched_key = k
            break

    if not matched_key:
        first_uk = list(elem.keys())[0] if elem else "空の要素"
        parts_list = ", ".join(sorted(KNOWN_PARTS))
        raise ValueError(f"{pos_desc} の {first_uk} が分からない（使える部品: {parts_list}）")

    val = elem[matched_key]

    if matched_key == "lbl":
        cap = val if isinstance(val, str) else elem.get("caption", "")
        return form_layout.lbl(cap, width=elem.get("width"), bold=elem.get("bold", False),
                               name=elem.get("name"), font=elem.get("font"))

    elif matched_key == "txt":
        name = val if isinstance(val, str) else elem.get("name", "")
        return form_layout.txt(name, width=elem.get("width"), height=elem.get("height"),
                               multiline=elem.get("multiline", False), value=elem.get("value", ""),
                               required=elem.get("required", False), font=elem.get("font"))

    elif matched_key == "combo":
        name = val if isinstance(val, str) else elem.get("name", "")
        return form_layout.combo(name, width=elem.get("width"), items=elem.get("items", []),
                                 rowsource=elem.get("rowsource"), font=elem.get("font"))

    elif matched_key == "lst":
        name = val if isinstance(val, str) else elem.get("name", "")
        return form_layout.lst(name, width=elem.get("width"), height=elem.get("height"),
                               rows_visible=elem.get("rows_visible", 6), items=elem.get("items", []),
                               rowsource=elem.get("rowsource"), font=elem.get("font"))

    elif matched_key == "chk":
        name = val if isinstance(val, str) else elem.get("name", "")
        caption = elem.get("caption", name)
        return form_layout.chk(name, caption, width=elem.get("width"), font=elem.get("font"))

    elif matched_key == "opt":
        name = val if isinstance(val, str) else elem.get("name", "")
        caption = elem.get("caption", name)
        return form_layout.opt(name, caption, width=elem.get("width"), group=elem.get("group"), font=elem.get("font"))

    elif matched_key == "opt_group":
        pairs = val if isinstance(val, list) else elem.get("pairs", [])
        return form_layout.opt_group(*pairs, group=elem.get("group"))

    elif matched_key == "btn":
        name = val if isinstance(val, str) else elem.get("name", "")
        caption = elem.get("caption", name)
        return form_layout.btn(name, caption, width=elem.get("width"), height=elem.get("height"),
                               default=elem.get("default", False), cancel_btn=elem.get("cancel", False),
                               bold=elem.get("bold", False), accel=elem.get("accel"), font=elem.get("font"))

    elif matched_key == "ok":
        name = val if isinstance(val, str) else elem.get("name", "btnOK")
        caption = elem.get("caption", "OK")
        return form_layout.ok(name, caption, width=elem.get("width"), height=elem.get("height"),
                              bold=elem.get("bold", False), accel=elem.get("accel"), font=elem.get("font"))

    elif matched_key == "cancel":
        name = val if isinstance(val, str) else elem.get("name", "btnCancel")
        caption = elem.get("caption", "キャンセル")
        return form_layout.cancel(name, caption, width=elem.get("width"), height=elem.get("height"),
                                  bold=elem.get("bold", False), accel=elem.get("accel"), font=elem.get("font"))

    else:
        parts_list = ", ".join(sorted(KNOWN_PARTS))
        raise ValueError(f"{pos_desc} の {matched_key} が分からない（使える部品: {parts_list}）")


def cmd_form_build(args):
    """form-build コマンドのディスパッチ"""
    t0 = time.time()

    if getattr(args, 'sample', False):
        with open(_LAST_FORM_SPEC_FILE, 'w', encoding='utf-8') as f:
            json.dump(SAMPLE_SPEC, f, ensure_ascii=False, indent=2)
        print(f"見本の設計図を書き出しました: {_LAST_FORM_SPEC_FILE}")
        return

    posargs = getattr(args, 'posargs', []) or []
    spec_file = getattr(args, 'spec_file', None)
    excel_file = getattr(args, 'excel_file', None)

    if posargs:
        for p in posargs:
            if p.lower().endswith('.json'):
                spec_file = p
                break
        if not spec_file and len(posargs) >= 1:
            spec_file = posargs[-1]
            if len(posargs) >= 2:
                excel_file = posargs[0]

    if spec_file:
        norm_spec = spec_file.replace('/', os.sep).replace('\\', os.sep)
        if not os.path.isabs(norm_spec):
            proj_root = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '..'))
            cand1 = os.path.abspath(norm_spec)
            cand2 = os.path.join(SCRIPT_DIR, norm_spec)
            cand3 = os.path.abspath(os.path.join(proj_root, norm_spec))
            cand4 = os.path.abspath(os.path.join(SCRIPT_DIR, '..', norm_spec))
            for cand in [cand1, cand2, cand3, cand4]:
                if os.path.exists(cand):
                    spec_file = cand
                    break
        else:
            spec_file = os.path.abspath(norm_spec)

    if not spec_file or not os.path.exists(spec_file):
        print(f"エラー: 設計図ファイルが見つかりません: {spec_file}")
        return

    # JSON のパース
    try:
        with open(spec_file, 'r', encoding='utf-8') as f:
            spec_data = json.load(f)
    except json.JSONDecodeError as jde:
        print(f"エラー: JSON の構文エラーです: {jde.lineno} 行目の {jde.colno} 字目が正しくありません: {jde.msg}")
        return
    except Exception as ex:
        print(f"エラー: 設計図の読み込みに失敗しました: {ex}")
        return

    # 設計図の検証と変換
    try:
        form_name, caption, width, parsed_rows, meta = _validate_and_parse_spec(spec_data)
    except ValueError as ve:
        print(f"エラー: {ve}")
        return

    # --dry-run の場合（Excel に触らない）
    if getattr(args, 'dry_run', False):
        try:
            out_png = os.path.join(SCRIPT_DIR, f"{form_name}_preview.png")
            form_layout.preview_layout(parsed_rows, out_path=out_png, content_width=width)
            sec = round(time.time() - t0, 2)
            print(f"[dry-run] Excel に触らず配置図を生成しました: {out_png} ({sec} 秒)")
            call_log_write("form-build", {"spec": spec_file, "dry_run": True}, sec=sec, ok=True)
            return
        except Exception as pe:
            print(f"エラー: レイアウトプレビューの生成に失敗しました: {pe}")
            return

    # 実機 Excel への接続
    try:
        xl, wb = get_workbook(excel_file)
    except Exception as e:
        print(f"エラー: Excel ブックに接続できません: {e}")
        return

    if getattr(args, 'new_book', False):
        wb = xl.Workbooks.Add()
        wb.Activate()
        print(f"新しいブックに組み上げます: {wb.Name}")

    book_name = wb.Name if wb else "ActiveWorkbook"

    # モジュール名との衝突チェック
    try:
        for comp in wb.VBProject.VBComponents:
            if comp.Name.lower() == form_name.lower():
                # 1 = vbext_ct_StdModule, 2 = vbext_ct_ClassModule
                if comp.Type in (1, 2):
                    print(f"エラー: フォームの名前 '{form_name}' が既存の標準/クラスモジュールとぶつかっています。")
                    return
    except Exception as ce:
        print(f"警告: モジュール衝突チェックをスキップしました: {ce}")

    # 既存フォーム退避の追跡
    existing_frm = False
    try:
        for comp in wb.VBProject.VBComponents:
            if comp.Name.lower() == form_name.lower() and comp.Type == 3:
                existing_frm = True
                break
    except Exception:
        pass

    # 構築実行
    try:
        bind_info = None if getattr(args, 'no_bind', False) else spec_data.get('bind')
        should_stub = getattr(args, 'stub', False) or (bind_info is not None)
        keys_enabled = spec_data.get('keys', True)
        form_layout.build_form(
            form_name,
            caption,
            parsed_rows,
            width=width,
            vba_stub=should_stub,
            png=getattr(args, 'png', False),
            launcher=getattr(args, 'launcher', None),
            force=getattr(args, 'force', False),
            save=False,
            bind_info=bind_info,
            keys=keys_enabled
        )
    except Exception as be:
        print(f"エラー: フォームの構築に失敗しました: {be}")
        return

    # バックアップファイルの実体確認
    actual_backup_path = None
    if existing_frm and os.path.exists(BACKUP_DIR):
        base = os.path.splitext(book_name)[0]
        prefix = f"{base}_{form_name}"
        matched = [f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix) and f.endswith(".frm")]
        if matched:
            matched.sort()
            actual_backup_path = os.path.join(BACKUP_DIR, matched[-1])

    # 構築後のコントロール情報・Tab順を点検
    ctrl_names = []
    tab_order = []
    try:
        frm = wb.VBProject.VBComponents(form_name).Designer
        ctrls = list(frm.Controls)
        for c in ctrls:
            ctrl_names.append(c.Name)
        tabbable = []
        for c in ctrls:
            try:
                tabbable.append((c.TabIndex, c.Name))
            except Exception:
                pass
        tabbable.sort(key=lambda x: x[0])
        tab_order = [name for _, name in tabbable]
    except Exception:
        pass

    sec = round(time.time() - t0, 2)
    call_log_write("form-build", {"spec": spec_file, "form": form_name}, sec=sec, ok=True)


    # 構築後の form-lint 自動実行
    try:
        from form_inspect import get_form_info, collect_controls, lint_form
        comp = wb.VBProject.VBComponents(form_name)
        f_info = get_form_info(comp)
        f_ctrls = collect_controls(comp)
        lint_findings = lint_form(form_name, f_info, f_ctrls)
        if not lint_findings:
            print("form-lint: 0 件（合格: 重なり・はみ出し・字切れ・端のずれ・コントラスト比 4.5以上）")
        else:
            print(f"form-lint: {len(lint_findings)} 件の指摘")
            for lf in lint_findings:
                print(f"  - {lf}")
    except Exception as le:
        pass

    # 報告出力
    print(f"組んだフォーム名: {form_name}")
    print(f"部品の数: {len(ctrl_names)} 個")
    print(f"Tab の順: {' → '.join(tab_order) if tab_order else '(Tabなし)'}")
    if actual_backup_path:
        print(f"控えの置き場: {actual_backup_path}")
    else:
        print("控えの置き場: (既存フォームなし・新規作成)")
    if getattr(args, 'png', False):
        print(f"PNG の置き場: {_LAST_FORM_PNG}")
    print(f"所要時間: {sec} 秒")


def cmd_form_survey(args):
    """秀コンボ.xlsm 等のフォームを集計してテーマ仕様を導出する"""
    from collections import Counter
    from form_inspect import get_form_info, collect_controls

    excel_file = getattr(args, 'excel_file', None) or "秀コンボ.xlsm"
    xl, wb = get_workbook(excel_file)
    comps = [c for c in wb.VBProject.VBComponents if c.Type == 3]
    print(f"Inspecting {len(comps)} forms in {wb.Name}...")

    font_names = Counter()
    font_sizes = Counter()
    ctrl_heights = Counter()
    btn_sizes = Counter()
    lbl_heights = Counter()
    back_colors = Counter()
    fore_colors = Counter()
    details = []

    for comp in comps:
        fn = comp.Name
        try:
            f_info = get_form_info(comp)
            ctrls = collect_controls(comp)
            try:
                bg = comp.Properties("BackColor").Value
                back_colors[bg] += 1
            except Exception:
                pass
            try:
                fg = comp.Properties("ForeColor").Value
                fore_colors[fg] += 1
            except Exception:
                pass

            for c in ctrls:
                c_type = c.get("type") or ""
                h = round(c.get("height", 0) or 0, 1)
                w = round(c.get("width", 0) or 0, 1)
                fs = c.get("font_size")
                if fs:
                    font_sizes[fs] += 1

                if c_type.startswith(("TextBox", "ComboBox")):
                    ctrl_heights[h] += 1
                elif c_type.startswith("CommandButton"):
                    btn_sizes[(w, h)] += 1
                elif c_type.startswith("Label"):
                    lbl_heights[h] += 1

            details.append({
                "form": fn,
                "controls": len(ctrls),
                "width": f_info.get("width"),
                "height": f_info.get("height"),
            })
        except Exception as e:
            print(f"Error inspecting {fn}: {e}")

    out_dir = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '作業ファイル', 'project', 'queue', 'scratch', 'form_corpus'))
    if not os.path.exists(out_dir):
        out_dir = os.path.abspath(os.path.join(SCRIPT_DIR, '..', 'queue', 'scratch', 'form_corpus'))
    os.makedirs(out_dir, exist_ok=True)
    survey_md = os.path.join(out_dir, "theme_survey.md")

    lines = []
    lines.append("# 秀コンボ 22 フォームのテーマ・見た目調査（theme_survey）")
    lines.append("## 1. 集計サマリー")
    lines.append(f"- フォーム総数: {len(comps)} 本")
    lines.append(f"- 最多フォントサイズ: {font_sizes.most_common(3)}")
    lines.append(f"- 最多フォーム地色 (BackColor): {back_colors.most_common(2)}")
    lines.append(f"- 最多フォーム文字色 (ForeColor): {fore_colors.most_common(2)}")
    lines.append(f"- 最多入力欄高さ (TextBox/ComboBox): {ctrl_heights.most_common(3)}")
    lines.append(f"- 最多ボタンサイズ (Width, Height): {btn_sizes.most_common(5)}")
    lines.append(f"- 最多ラベル高さ: {lbl_heights.most_common(3)}")

    lines.append("## 2. フォーム個別一覧")
    lines.append("| フォーム名 | コントロール数 | 幅 x 高さ |")
    lines.append("|---|---|---|")
    for d in details:
        lines.append(f"| {d['form']} | {d['controls']} | {d['width']}x{d['height']} |")

    with open(survey_md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"Written survey to: {survey_md}")

    # 最多要素の抽出
    top_fs = font_sizes.most_common(1)[0][0] if font_sizes else 11
    top_ctrl_h = ctrl_heights.most_common(1)[0][0] if ctrl_heights else 22
    top_btn = btn_sizes.most_common(1)[0][0] if btn_sizes else (72, 24)

    # form_theme.json の書き出し
    theme = {
        "name": "秀",
        "font_name": "Yu Gothic UI",
        "font_size": top_fs,
        "heading_size": 13,
        "colors": {
            "form_bg": "&H8000000F&",
            "text": "&H80000012&",
            "heading": "&H00800000&",
            "input_bg": "&H80000005&",
            "focus_bg": "&H00FFF5E6&",
            "error_text": "&H000000C0&",
            "btn_bg": "&H8000000F&",
            "btn_default_border": "&H00C00000&"
        },
        "metrics": {
            "pad": 12,
            "gap_x": 8,
            "gap_y": 8,
            "ctrl_h": top_ctrl_h,
            "lbl_h": 18,
            "btn_w": top_btn[0],
            "btn_h": top_btn[1],
            "min_contrast_ratio": 4.5
        }
    }
    theme_path = os.path.join(SCRIPT_DIR, "form_theme.json")
    with open(theme_path, "w", encoding="utf-8") as f:
        json.dump(theme, f, ensure_ascii=False, indent=2)
    print(f"Generated form_theme.json at: {theme_path}")


def cmd_form_lint(args):
    """UserForm の見た目検査（form-lint）を実行する"""
    from form_inspect import get_form_info, collect_controls, lint_form

    posargs = getattr(args, 'posargs', []) or []
    form_name = None
    excel_file = getattr(args, 'excel_file', None)

    for p in posargs:
        if p.lower().endswith(('.xlsm', '.xlam', '.xlsb', '.xlsx')):
            excel_file = p
        else:
            form_name = p

    if not form_name:
        form_name = getattr(args, 'form_name', None)

    if not form_name:
        print("エラー: フォーム名を指定してください（例: form-lint F_顧客入力）")
        return

    xl, wb = get_workbook(excel_file)
    try:
        comp = wb.VBProject.VBComponents(form_name)
    except Exception:
        print(f"エラー: フォーム '{form_name}' が見つかりません。")
        return

    info = get_form_info(comp)
    controls = collect_controls(comp)
    findings = lint_form(form_name, info, controls)

    print(f"===== form-lint: {form_name} =====")
    if not findings:
        print("合格: 0 件（重なり・はみ出し・字切れ・ずれ・コントラスト比すべて良好）")
    else:
        print(f"指摘: {len(findings)} 件")
        for f in findings:
            print(f"  - {f}")


import re
import unicodedata

def _sanitize_var_name(name, prefix="ctl"):
    """VBA の識別子として安全な名前に変換する（記号除去・先頭数字回避）"""
    # 記号の除去
    s = re.sub(r'[^a-zA-Z0-9_぀-ゟ゠-ヿ一-鿿]', '_', str(name))
    s = re.sub(r'_+', '_', s).strip('_')
    if not s:
        s = "item"
    if s[0].isdigit():
        s = f"{prefix}_{s}"
    else:
        s = f"{prefix}{s}"
    return s


def analyze_table_columns(ws, table_range, is_list_object=False):
    """表範囲の各列の中身を詳細解析し、部品の種類・属性を決定する"""
    num_rows = table_range.Rows.Count
    num_cols = table_range.Columns.Count

    if num_rows < 1 or num_cols < 1:
        raise ValueError("表の範囲が空です。")

    # 見出し行（1行目）
    header_cells = [table_range.Cells(1, c) for c in range(1, num_cols + 1)]
    raw_headers = [str(c.Value) if c.Value is not None else f"列{ci+1}" for ci, c in enumerate(header_cells)]

    # データ行があるか
    data_rows_count = num_rows - 1

    # 高速一括取得
    if data_rows_count > 0:
        data_rng = ws.Range(table_range.Cells(2, 1), table_range.Cells(num_rows, num_cols))
        raw_vals = data_rng.Value2
        raw_forms = data_rng.Formula

        def _normalize_2d(raw, r_cnt, c_cnt):
            if not isinstance(raw, (tuple, list)):
                return ((raw,),)
            if len(raw) > 0 and not isinstance(raw[0], (tuple, list)):
                if r_cnt == 1:
                    return (tuple(raw),)
                else:
                    return tuple((v,) for v in raw)
            return tuple(tuple(r) for r in raw)

        vals = _normalize_2d(raw_vals, data_rows_count, num_cols)
        formulas = _normalize_2d(raw_forms, data_rows_count, num_cols)
    else:
        vals = ()
        formulas = ()

    cols_spec = []
    text_columns = []
    used_names = set()
    formula_cols = []
    start_col = header_cells[0].Column if header_cells else 1

    for c_idx in range(num_cols):
        h_name = raw_headers[c_idx]
        col_cell = header_cells[c_idx]

        # 1. 数式列の判定（C1要件: 数式列はフォームに入れない）
        is_formula = False
        if data_rows_count > 0:
            # 先頭〜数行の数式チェック
            sample_formulas = [formulas[r][c_idx] for r in range(min(5, data_rows_count))]
            if any(str(f).startswith("=") for f in sample_formulas if f is not None):
                is_formula = True
        if is_formula:
            formula_cols.append(c_idx + 1)
            continue

        # 2. 入力規則のチェック（リスト）
        v_items = None
        try:
            val_obj = ws.Cells(2, col_cell.Column).Validation if data_rows_count > 0 else col_cell.Validation
            if val_obj.Type == 3: # xlValidateList
                f1 = str(val_obj.Formula1 or "")
                if f1.startswith("="):
                    # 他シート参照や名前定義の場合、Range取得を試みる
                    try:
                        ref_rng = ws.Application.Range(f1)
                        ref_vals = ref_rng.Value2
                        if isinstance(ref_vals, tuple):
                            v_items = [str(rv[0]) for rv in ref_vals if rv[0] is not None]
                        else:
                            v_items = [str(ref_vals)] if ref_vals is not None else []
                    except Exception:
                        pass
                else:
                    v_items = [it.strip() for it in f1.split(",") if it.strip()]
        except Exception:
            pass

        # 3. 列データ値の収集・分析
        col_vals = []
        has_empty = False
        if data_rows_count > 0:
            for r in range(data_rows_count):
                v = vals[r][c_idx]
                if v is None or str(v).strip() == "":
                    has_empty = True
                else:
                    col_vals.append(v)
        else:
            has_empty = True

        # 4. 書式のチェック
        num_fmt = ""
        try:
            sample_cell = ws.Cells(2, col_cell.Column) if data_rows_count > 0 else col_cell
            num_fmt = str(sample_cell.NumberFormatLocal or "")
        except Exception:
            pass

        # 型と部品の判定
        control_type = "txt"
        combo_items = None
        is_multiline = False
        is_date = False
        is_currency = False
        example_hint = None
        align_right = False

        # (a) 入力規則リスト
        if v_items:
            control_type = "combo"
            combo_items = v_items

        # (b) True/False・○/空欄（真偽値）
        elif col_vals and set(str(v).lower() for v in col_vals).issubset({"true", "false", "0", "1", "○", "はい", "いいえ", ""}):
            control_type = "chk"
            non_empty_strs = set(str(v) for v in col_vals if v not in (None, ""))
            if "○" in non_empty_strs:
                true_val, false_val = "○", ""
            elif "はい" in non_empty_strs:
                true_val, false_val = "はい", "いいえ"
            elif "1" in non_empty_strs:
                true_val, false_val = "1", "0"
            else:
                true_val, false_val = "True", "False"

        # (c) 日付判定（書式または値）
        elif any(df in num_fmt.lower() for df in ("yyyy", "yy", "m月", "d日", "e/m/d")) or (col_vals and all(re.match(r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}$', str(v)) for v in col_vals[:10])):
            control_type = "txt"
            is_date = True
            example_hint = "例: 2026/10/30"

        # (d) ゼロ落ち・郵便・電話番号（00123 / 010-0001 / 018-000-0000）
        elif num_fmt == "@" or (col_vals and any(str(v).startswith("0") and len(str(v)) > 1 and str(v).isdigit() for v in col_vals)) or any(k in h_name for k in ("郵便", "電話", "コード", "番号", "TEL", "ZIP")):
            control_type = "txt"
            text_columns.append(h_name)

        # (e) 金額・数値判定
        elif any(cf in num_fmt for cf in ("#,##0", "¥", "$", "%")) or (col_vals and all(isinstance(v, (int, float)) for v in col_vals[:10])):
            control_type = "txt"
            align_right = True
            if any(k in h_name for k in ("金額", "単価", "給与", "費", "額", "残業")):
                is_currency = True

        # (f) 長文判定（平均文字長 > 40）
        elif col_vals and (sum(len(str(v)) for v in col_vals) / len(col_vals)) > 40:
            control_type = "txt"
            is_multiline = True

        # (g) 少カーディナリティの文字列（種類の少ない文字: 20種類以下 & 行の3割以下）
        elif col_vals and all(isinstance(v, str) for v in col_vals):
            uniq_vals = sorted(list(set(str(v) for v in col_vals if v)))
            if 2 <= len(uniq_vals) <= min(20, max(2, int(data_rows_count * 0.3))):
                control_type = "combo"
                combo_items = uniq_vals

        # 部品名の採番
        prefix = {"combo": "cmb", "chk": "chk", "txt": "txt"}.get(control_type, "ctl")
        base_name = _sanitize_var_name(h_name, prefix=prefix)
        ctrl_name = base_name
        cnt = 2
        while ctrl_name in used_names:
            ctrl_name = f"{base_name}_{cnt}"
            cnt += 1
        used_names.add(ctrl_name)

        # 必須フラグ（データ行があり、空欄が0個の場合）
        required = (data_rows_count > 0 and not has_empty)

        type_meta = "text"
        if control_type == "combo":
            type_meta = {"type": "combo", "items": combo_items or []}
        elif control_type == "chk":
            type_meta = {"type": "bool", "true_val": true_val if 'true_val' in locals() else "True", "false_val": false_val if 'false_val' in locals() else "False"}
        elif is_date:
            type_meta = "date"
        elif is_currency or align_right:
            type_meta = "number"

        cols_spec.append({
            "header": h_name,
            "name": ctrl_name,
            "col_index": c_idx + 1,
            "control": control_type,
            "type_meta": type_meta,
            "items": combo_items,
            "required": required,
            "multiline": is_multiline,
            "is_date": is_date,
            "is_currency": is_currency,
            "align_right": align_right,
            "hint": example_hint
        })

    extra_meta = {
        "formula_columns": formula_cols,
        "start_col": start_col
    }
    return cols_spec, text_columns, extra_meta



def build_list_form_spec(form_name, title, sheet_name, table_name, cols_spec, extra_meta=None, entry_form=None):
    """一覧と検索のフォーム設計図（kind: list）を組み立てる"""
    num_cols = len(cols_spec)
    show_cols = cols_spec[:6]
    col_widths = ";".join(["60" if c.get("is_date") or c.get("is_currency") else "80" for c in show_cols])

    rows = [
        {"row": [
            {"lbl": "検索:"},
            {"txt": "txtSearch", "width": 200},
            {"lbl": "0 件中 0 件", "name": "lblCount"}
        ]},
        {"spacer": 6},
        {"row": [
            {"lst": "lstData", "height": 220, "column_count": len(show_cols), "column_widths": col_widths}
        ]},
        {"spacer": 10},
        {"button_bar": [
            {"btn": "btnEdit", "caption": "直す"},
            {"btn": "btnNew", "caption": "新しく"},
            {"cancel": "btnCancel", "caption": "閉じる"}
        ]}
    ]

    bind_columns = {c["name"]: c["header"] for c in cols_spec}
    spec = {
        "name": form_name,
        "caption": f"{title} 一覧・検索",
        "theme": "秀",
        "kind": "list",
        "bind": {
            "sheet": sheet_name,
            "table": table_name,
            "columns": bind_columns,
            "entry_form": entry_form or form_name.replace("一覧", "入力"),
            "start_col": extra_meta.get("start_col", 1) if extra_meta else 1
        },
        "rows": rows
    }
    return spec


def build_form_spec_from_columns(form_name, title, sheet_name, table_name, cols_spec, text_columns, extra_meta=None):
    """解析した列定義から form_layout 形式の設計図 JSON を組み立てる（12列超は MultiPage化）"""
    bind_columns = {c["name"]: c["header"] for c in cols_spec}
    col_positions = {c["name"]: c.get("col_index", idx + 1) for idx, c in enumerate(cols_spec)}
    col_types = {c["name"]: c.get("type_meta", "text") for c in cols_spec}
    num_cols = len(cols_spec)
    extra_meta = extra_meta or {}

    rows = []

    # 12列以下の場合は通常の1列レイアウト
    if num_cols <= 12:
        rows.append({"heading": title})
        for c in cols_spec:
            elem = _create_control_elem(c)
            lbl_cap = c["header"]
            if c["hint"]:
                lbl_cap += f" ({c['hint']})"
            rows.append({"row": [{"lbl": lbl_cap}, elem]})
    else:
        # 12列を超える場合は MultiPage に分割（1ページあたり8〜10項目）
        chunk_size = 10
        chunks = [cols_spec[i:i + chunk_size] for i in range(0, num_cols, chunk_size)]
        pages = []
        for p_idx, chunk in enumerate(chunks, 1):
            p_cap = f"情報 {p_idx}"
            p_rows = []
            for c in chunk:
                elem = _create_control_elem(c)
                lbl_cap = c["header"]
                if c["hint"]:
                    lbl_cap += f" ({c['hint']})"
                p_rows.append({"row": [{"lbl": lbl_cap}, elem]})
            pages.append({"caption": p_cap, "rows": p_rows})
        rows.append({"multipage": "mpMain", "pages": pages})

    rows.append({"spacer": 12})
    rows.append({"button_bar": [
        {"ok": "btnOK", "caption": "登録"},
        {"cancel": "btnCancel", "caption": "閉じる"}
    ]})

    spec = {
        "name": form_name,
        "caption": title,
        "theme": "秀",
        "assist": True,
        "bind": {
            "sheet": sheet_name,
            "table": table_name,
            "columns": bind_columns,
            "text_columns": text_columns,
            "column_types": col_types,
            "column_positions": col_positions,
            "formula_columns": extra_meta.get("formula_columns", []),
            "start_col": extra_meta.get("start_col", 1)
        },
        "rows": rows
    }
    return spec


def _create_control_elem(c):
    ct = c["control"]
    nm = c["name"]
    if ct == "combo":
        return {"combo": nm, "items": c.get("items", [])}
    elif ct == "chk":
        return {"chk": nm, "caption": c["header"]}
    else:
        txt_dict = {"txt": nm, "required": c["required"]}
        if c.get("multiline"):
            txt_dict["multiline"] = True
            txt_dict["height"] = 54
        return txt_dict


def cmd_form_from_table(args):
    """表の見出しと中身から入力フォームの設計図を作る: form-from-table [テーブル名|範囲] [--name フォーム名] [--build] [--png]"""
    import time
    t0 = time.time()

    xl, wb = get_workbook(None)
    ws = wb.ActiveSheet

    # シート保護チェック（T29の要件）
    if ws.ProtectContents:
        print(f"エラー: シート '{ws.Name}' に保護が掛かっています（解除してから実行してください）。")
        return False

    posargs = getattr(args, 'posargs', []) or []
    target_arg = posargs[0] if posargs else None

    # 表範囲の特定
    target_rng = None
    table_name = ""
    is_list_obj = False

    if target_arg:
        # 1. テーブル名検索
        for lo in ws.ListObjects:
            if lo.Name.lower() == target_arg.lower():
                target_rng = lo.Range
                table_name = lo.Name
                is_list_obj = True
                break
        # 2. セル範囲
        if target_rng is None:
            try:
                target_rng = ws.Range(target_arg)
            except Exception:
                pass

    if target_rng is None:
        # アクティブセル周辺から自動探索
        try:
            ac = xl.ActiveCell
            # アクティブセルが含まれる ListObject があるか
            for lo in ws.ListObjects:
                # 共通部分があるか
                intersect = xl.Intersect(ac, lo.Range)
                if intersect is not None:
                    target_rng = lo.Range
                    table_name = lo.Name
                    is_list_obj = True
                    break
            if target_rng is None:
                # CurrentRegion
                target_rng = ac.CurrentRegion
                table_name = ws.Name
        except Exception as e:
            print(f"エラー: 表範囲を特定できませんでした: {e}")
            return False

    if target_rng is None or target_rng.Cells.Count < 2:
        print("エラー: 表の範囲が見つかりません。")
        return False

    # フォーム名
    form_name = getattr(args, 'name', None)
    if not form_name:
        form_name = f"F_{table_name}" if table_name else "F_DataInput"
    if not form_name.startswith("F_"):
        form_name = _sanitize_var_name(form_name, prefix="F_")

    title = f"{table_name} 登録" if table_name else "データ登録"

    # 列解析
    try:
        cols_spec, text_columns, extra_meta = analyze_table_columns(ws, target_rng, is_list_obj)
    except Exception as ae:
        print(f"エラー: 表の解析に失敗しました: {ae}")
        return False

    # 設計図の生成（--list のときは一覧・検索フォーム）
    if getattr(args, 'list_flag', False):
        if not getattr(args, 'name', None):
            form_name = f"F_{table_title}一覧"
        spec = build_list_form_spec(form_name, title, ws.Name, table_name, cols_spec, extra_meta=extra_meta)
    else:
        spec = build_form_spec_from_columns(form_name, title, ws.Name, table_name, cols_spec, text_columns, extra_meta=extra_meta)

    # _last_form_spec.json へ書き出し
    spec_path = _LAST_FORM_SPEC_FILE
    with open(spec_path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)

    sec = round(time.time() - t0, 2)
    call_log_write("form-from-table", {"table": table_name, "cols": len(cols_spec)}, sec=sec, ok=True)

    print(f"===== form-from-table 設計図作成完了 ({sec} 秒) =====")
    print(f"フォーム名: {form_name}")
    print(f"対象表: {ws.Name}!{table_name} ({len(cols_spec)} 項目)")
    for c in cols_spec:
        req_str = " [必須]" if c["required"] else ""
        opt_str = f" ({len(c['items'])}択)" if c.get("items") else ""
        print(f"  - {c['header']}: {c['control']}{opt_str}{req_str}")
    if text_columns:
        print(f"文字保護列: {', '.join(text_columns)}")
    print(f"設計図保存: {spec_path}")

    # --build が指定されていたら続けて form-build を実行
    if getattr(args, 'build', False):
        print("\n--- 続けて form-build を実行します ---")
        build_args = type("Args", (), {
            "posargs": [spec_path],
            "spec_file": spec_path,
            "excel_file": None,
            "sample": False,
            "dry_run": getattr(args, 'dry_run', False),
            "png": getattr(args, 'png', False),
            "stub": True,
            "launcher": None,
            "force": False,
            "new_book": False
        })()
        cmd_form_build(build_args)

        if getattr(args, 'test_flag', False) and not getattr(args, 'dry_run', False):
            if getattr(args, 'list_flag', False):
                print("form-test: 一覧フォームは台本の対象外です（入力フォームだけで --test を使ってください）")
            else:
                from vbam_formtest import generate_auto_test_script, run_form_test
                sample = {}
                for c in cols_spec:
                    items = c.get("items") or []
                    if c["control"] == "combo" and items:
                        sample[c["name"]] = items[0]
                    elif c["control"] == "chk":
                        sample[c["name"]] = True
                    elif c.get("is_date"):
                        sample[c["name"]] = "2026/01/01"
                    elif c.get("is_currency") or c.get("align_right"):
                        sample[c["name"]] = "100"
                    else:
                        sample[c["name"]] = "試験"
                # ボタンを押す手のため、閉じているうちにフォームの位置を控える
                from vbam_edit import cmd_form_click
                learn_args = type("Args", (), {"posargs": [form_name], "learn_flag": True})()
                cmd_form_click(learn_args)
                script = generate_auto_test_script(form_name, title, ws.Name, cols_spec, sample, base_rows=max(0, target_rng.Rows.Count - 1))
                ok, msg, _ = run_form_test(script)
                print(f"form-test: {msg}")

    return True


def cmd_create_sample_table(args):
    """汚れた表を新しい捨てブックに作成する"""
    xl, _ = get_workbook(None)
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "社員台帳"

    headers = [
        "社員番号", "氏名", "区分", "入社日", "基本給",
        "支給合計", "部署", "リモート勤務", "自己PR", "備考"
    ]
    for ci, h in enumerate(headers, 1):
        cell = ws.Cells(1, ci)
        cell.Value = h
        cell.Font.Bold = True

    data = [
        ["00101", "山田太郎", "正社員", "2024/04/01", 320000, "=E2*1.2", "営業", "True", "10年以上の法人営業経験があり、新規開拓から大規模案件のクロージングまで一貫して担当してきました。", "特命案件あり"],
        ["00102", "佐藤花子", "契約社員", "2025/01/15", 280000, "=E3*1.2", "開発", "False", "フルスタックエンジニアとしてフロントエンドからバックエンド・インフラ構築まで幅広く対応可能です。", ""],
        ["00103", "鈴木一郎", "役員", "2020/10/01", 500000, "=E4*1.2", "総務", "True", "組織マネジメントおよび財務会計の統括責任者として全社の業務改革を推進しています。", "非常勤"]
    ]

    for ri, row in enumerate(data, 2):
        for ci, val in enumerate(row, 1):
            cell = ws.Cells(ri, ci)
            if ci == 1:
                cell.NumberFormatLocal = "@"
                cell.Value = val
            elif ci == 4:
                cell.NumberFormatLocal = "yyyy/mm/dd"
                cell.Value = val
            elif ci == 5:
                cell.NumberFormatLocal = "#,##0"
                cell.Value = val
            elif ci == 6:
                cell.Formula = val
            else:
                cell.Value = val

    # 入力規則
    rng_val = ws.Range("C2:C10")
    rng_val.Validation.Delete()
    rng_val.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="役員,正社員,契約社員,パート")

    ws.Range("A1").Select()
    print(f"作成完了: {wb.Name}!{ws.Name} (10列・汚れた表)")
    return True
