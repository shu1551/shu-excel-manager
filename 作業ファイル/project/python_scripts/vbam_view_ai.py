# -*- coding: utf-8 -*-
"""vbam_view_ai.py — vba_manager 分割パート: AI に見せる見た目（画像の範囲・行列のヘッダー・添え文）と終わりの検査

2026-09-11 に vbam_agent.py から中身を変えずに切り出した（vbam_view.py＝人が見る目のコマンドとは別）。
vbam_agent は `from vbam_view_ai import *` で名前を引き継ぐ。
"""
import os
import re
import contextlib
import shutil
from vbam_core import (SCRIPT_DIR, _col_letter)
from vbam_hands import (_XL_MAX_COL, _XL_MAX_ROW)
from vbam_undo import (_same_path)

_LAST_AGENT_PNG = os.path.join(SCRIPT_DIR, '_last_agent.png')              # 終わったときの見た目
_LAST_AGENT_VIEW_PNG = os.path.join(SCRIPT_DIR, '_last_agent_view.png')    # 往復のあいだ AI に見せる見た目
_VIEW_MAX_ROWS, _VIEW_MAX_COLS = 40, 20                                    # AI に見せる画像の大きさ（左上から）


_VIEW_MARGIN = 2                # 画像の上下左右に付ける余白（行・列）


def _view_range(ws):
    """AI に見せる範囲（使用範囲の左上から 40 行 × 20 列まで＋上下左右 2 行 2 列の余白）。読めなければ None。

    余白が無いと、表がどこで終わっているのか・隣に何もないのかが画像から分からない（2026-09-04・
    Claude for Excel への取材で「書いた範囲＋数行数列の余白」が要ると言われた）。
    """
    try:
        ur = ws.UsedRange
        r0, c0 = int(ur.Row), int(ur.Column)
        nr, nc = min(int(ur.Rows.Count), _VIEW_MAX_ROWS), min(int(ur.Columns.Count), _VIEW_MAX_COLS)
        if nr < 1 or nc < 1:
            return None
        r1, c1 = r0 + nr - 1 + _VIEW_MARGIN, c0 + nc - 1 + _VIEW_MARGIN     # 下・右の余白
        r0, c0 = max(1, r0 - _VIEW_MARGIN), max(1, c0 - _VIEW_MARGIN)       # 上・左の余白（行 1・列 A で止める）
        r1 = min(r1, _XL_MAX_ROW, r0 + _VIEW_MAX_ROWS - 1)                  # 40×20 の上限は保つ
        c1 = min(c1, _XL_MAX_COL, c0 + _VIEW_MAX_COLS - 1)
        return f"{_col_letter(c0)}{r0}:{_col_letter(c1)}{r1}", r1 - r0 + 1, c1 - c0 + 1
    except Exception:
        return None


_VIEW_HEADED = False            # 直近の画像に行番号・列番号のヘッダーを描けたか（添え文がこれで変わる）
_VIEW_HEADER_BG = (232, 232, 232)
_VIEW_HEADER_LINE = (160, 160, 160)
_VIEW_HEADER_TEXT = (40, 40, 40)


def _header_font(px):
    """ヘッダーの文字（英数字だけ）。Windows の Arial → 無ければ PIL の既定。"""
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    for name in ('arial.ttf', 'segoeui.ttf', 'tahoma.ttf'):
        p = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', name)
        if os.path.isfile(p):
            with contextlib.suppress(Exception):
                return ImageFont.truetype(p, max(8, int(px)))
    with contextlib.suppress(Exception):
        return ImageFont.load_default()
    return None


def _text_size(draw, text, font):
    try:
        x0, y0, x1, y1 = draw.textbbox((0, 0), text, font=font)
        return x1 - x0, y1 - y0
    except Exception:
        return 6 * len(text), 10


def _draw_headers(png, ws, addr):
    """画像に行番号・列番号のヘッダーを描き足す（同じファイルに上書き）。描けなければ False＝画像はそのまま。

    CopyPicture はヘッダーを含められない。範囲の列幅・行高（ポイント）と画像の大きさから倍率を出し、
    上に列の文字・左に行の数字を Excel の画面と同じ帯で足す。
    2026-09-04 の取材で Claude for Excel が「行番号・列番号のヘッダーが写っていること」を注文していたが、
    それまでは添え文で「左上のセルが A1 です。ここから数えてください」と AI に数えさせていた＝
    20 列目・40 行目で 1 つずれても誰にも分からなかった（2026-09-06）。
    """
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return False
    try:
        rng = ws.Range(addr)
        r0, c0 = int(rng.Row), int(rng.Column)
        nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
        widths = [float(rng.Columns(i + 1).Width or 0) for i in range(nc)]
        heights = [float(rng.Rows(i + 1).Height or 0) for i in range(nr)]
        im = Image.open(png).convert('RGB')
    except Exception:
        return False
    tw, th = sum(widths), sum(heights)
    if tw <= 0 or th <= 0:
        return False
    try:
        W, H = im.size
        sx, sy = W / tw, H / th
        top = max(14, int(round(15 * sy)))                       # 列の帯＝標準の行 1 本ぶん
        font = _header_font(min(top - 3, 13))
        probe = ImageDraw.Draw(im)
        left = max(24, _text_size(probe, str(r0 + nr - 1), font)[0] + 10)    # 行の帯＝一番大きい行番号が入る幅
        out = Image.new('RGB', (W + left, H + top), _VIEW_HEADER_BG)
        out.paste(im, (left, top))
        d = ImageDraw.Draw(out)
        d.line([(0, top - 1), (W + left, top - 1)], fill=_VIEW_HEADER_LINE)
        d.line([(left - 1, 0), (left - 1, H + top)], fill=_VIEW_HEADER_LINE)
        x = float(left)
        for i, w in enumerate(widths):
            px = w * sx
            if px >= 6:                                          # 隠れた列・細すぎる列は文字を置かない
                label = _col_letter(c0 + i)
                lw, lh = _text_size(d, label, font)
                if lw <= px - 2:
                    d.text((x + (px - lw) / 2, (top - lh) / 2 - 1), label, fill=_VIEW_HEADER_TEXT, font=font)
            x += px
            d.line([(int(round(x)), 0), (int(round(x)), top - 1)], fill=_VIEW_HEADER_LINE)
        y = float(top)
        for i, h in enumerate(heights):
            py = h * sy
            if py >= 6:
                label = str(r0 + i)
                lw, lh = _text_size(d, label, font)
                if lh <= py - 1:
                    d.text(((left - lw) / 2 - 1, y + (py - lh) / 2 - 1), label, fill=_VIEW_HEADER_TEXT, font=font)
            y += py
            d.line([(0, int(round(y))), (left - 1, int(round(y)))], fill=_VIEW_HEADER_LINE)
        out.save(png)
        return True
    except Exception:
        return False


def _view_note(addr, tail=''):
    """画像の添え文。ヘッダーを描けた画像なら「行番号・列番号が写っている」、描けなければ左上の番地から数えさせる。"""
    head = f"（このシートの {addr} の見た目の画像を添えました＝"
    if _VIEW_HEADED:
        head += "Excel の画面と同じく、上に列の文字・左に行番号のヘッダーが写っています。番地はそれで読んでください。"
    else:
        head += f"画像の左上のセルが {addr.split(':')[0]} です。行番号・列番号は写っていないので、ここから数えてください。"
    return head + tail + "）\n"


def _shot_for_ai(sheet, wb):
    """書いた直後の見た目を PNG にして (パス, 番地) を返す（AI に添える）。撮れなければ (None, None)。

    文字の読み戻しでは罫線の継ぎはぎ・右寄せ・カンマ無しが見えない（「ひどい出来だ」の 3 連発は全部これ）。
    Gemini も Claude も画像を読めるので、人にだけ見せていた screenshot を AI にも見せる（2026-09-04）。
    番地も一緒に返す（添え文に書く）。画像には行番号・列番号のヘッダーを描き足す（2026-09-06。描けたかは
    _VIEW_HEADED に残り、添え文 _view_note がそれで変わる）。
    """
    from vbam_agent import (_run_cmd)   # 分割後の遅延 import（循環にしない・2026-09-11）
    global _VIEW_HEADED
    got = _view_range(wb.Sheets(sheet))
    if not got:
        return None, None
    addr, nr, nc = got
    ok, _out = _run_cmd(['screenshot', f"'{sheet}'!{addr}", '--out', _LAST_AGENT_VIEW_PNG], wb)
    if not ok or not os.path.isfile(_LAST_AGENT_VIEW_PNG):
        return None, None
    try:
        _VIEW_HEADED = _draw_headers(_LAST_AGENT_VIEW_PNG, wb.Sheets(sheet), addr)
    except Exception:
        _VIEW_HEADED = False
    print(f"画像: {addr}（{nr} 行 × {nc} 列・{os.path.getsize(_LAST_AGENT_VIEW_PNG):,} バイト"
          + ("・行列ヘッダー付き" if _VIEW_HEADED else "") + "）を AI に見せます")
    return _LAST_AGENT_VIEW_PNG, addr


_WANTS_FIRST_IMAGE_RE = re.compile(r'結合セル\s+(\d+)件|図形・ボタン:\s*(\d+)個')


def _wants_first_image(materials):
    """materials の文字だけでは崩れて見える表か（結合セル・図形が 1 つでもある）。1 往復目にも画像を添える判定（2026-09-05）。

    「結合セル: 未走査（大きい表）」は件数が無いので当たらない（見せられる画像も無い）。
    """
    return any(int(n1 or n2) >= 1 for n1, n2 in _WANTS_FIRST_IMAGE_RE.findall(materials or ''))



def _count_err_hash(text):
    """materials の出力から (エラーセルの数, ### の数)。行が無ければ 0（materials は 0 個のとき行を出さない）。"""
    m = re.search(r'エラーセル: (\d+)', text or '')
    n_err = int(m.group(1)) if m else 0
    m = re.search(r"'###' で読めないセル: (\d+)", text or '')
    return n_err, (int(m.group(1)) if m else 0)


def _verify_sheet(sheet, wb=None, base=(0, 0), materials=None, png_src=None):
    """終わりの検査。base＝始める前の (エラーセル, ###)。元からあった分は数えず、増えていなければ合格
    （2026-09-04: 元から #N/A のあるシートで、AI が done でも「要確認」＋終了コード 1 になっていた）。

    materials／png_src を渡すと読み直さず・撮り直さない（2026-09-06 夜: 採点が通った直後は採点に渡した
    材料と画像がそのまま現物＝同じシートを 3 回読み・4 回撮っていた道具側 6〜9 秒の削り先）。
    """
    from vbam_agent import (_run_cmd)   # 分割後の遅延 import（循環にしない・2026-09-11）
    if materials:
        mat = materials
    else:
        ok_r, mat = _run_cmd(['materials', sheet], wb)
    n_err, n_hash = _count_err_hash(mat)
    png = None
    # 検査の数字は materials 全体（使用範囲の全部）のまま。画像だけ AI 用と同じ 40×20 に絞る
    # （使用範囲が広いシートで、撮るのに時間がかかり、縮んで何も読めない画像になっていた・2026-09-04）
    if png_src and os.path.isfile(str(png_src)):
        import shutil
        with contextlib.suppress(Exception):                     # 採点に見せた画像（行列ヘッダー付き）をそのまま人にも
            if not _same_path(png_src, _LAST_AGENT_PNG):
                shutil.copyfile(png_src, _LAST_AGENT_PNG)
            png = _LAST_AGENT_PNG
    got = (_view_range(wb.Sheets(sheet)) if (wb is not None and not png) else None)
    if got:
        ok_s, _ = _run_cmd(['screenshot', f"'{sheet}'!{got[0]}", '--out', _LAST_AGENT_PNG], wb)
        if ok_s:
            png = _LAST_AGENT_PNG
            with contextlib.suppress(Exception):             # 人が見る画像にも行番号・列番号を付ける
                _draw_headers(_LAST_AGENT_PNG, wb.Sheets(sheet), got[0])
    ok = (n_err <= base[0] and n_hash <= base[1])
    print(f"検査: エラーセル {n_err}" + (f"（元から {base[0]}）" if base[0] else "")
          + f" / ### {n_hash}" + (f"（元から {base[1]}）" if base[1] else "")
          + " / " + (f"画像 {png}" if png else "画像なし") + ("  → 合格" if ok else "  → 要確認（増えている）"))
    return ok


__all__ = ['_LAST_AGENT_PNG', '_LAST_AGENT_VIEW_PNG', '_VIEW_HEADED', '_VIEW_HEADER_BG', '_VIEW_HEADER_LINE', '_VIEW_HEADER_TEXT', '_VIEW_MARGIN', '_VIEW_MAX_COLS', '_VIEW_MAX_ROWS', '_WANTS_FIRST_IMAGE_RE', '_count_err_hash', '_draw_headers', '_header_font', '_shot_for_ai', '_text_size', '_verify_sheet', '_view_note', '_view_range', '_wants_first_image']
