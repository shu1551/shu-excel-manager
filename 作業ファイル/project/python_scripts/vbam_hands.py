# -*- coding: utf-8 -*-
"""vbam_hands.py — vba_manager 分割パート: 道具が COM で直接やる手の中身と、材料の読み手

2026-09-11 に vbam_agent.py から中身を変えずに切り出した。normalize／fill（大きい表の手）、追加の材料
（見出し行の推定・番号列・膨張・外部リンク・名前定義・空欄…）、列プロファイル、read_file／CSV、番地の検査。
手の振り分け（_direct_action）は vbam_agent に残る。vbam_agent は `from vbam_hands import *` で名前を引き継ぐ。
"""
import os
import re
import io
import difflib
import unicodedata
import csv
import datetime
from vbam_core import (_col_letter)

_ADDR_RE = re.compile(r'^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$')
_XL_MAX_ROW = 1048576
_XL_MAX_COL = 16384                    # XFD
_FILL_MAX_CELLS = 200000      # fill で伸ばせるセル数の上限（normalize と同じ。列全体を指定されると使用範囲が膨らむ）

def _col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n

def _check_addr(a):
    a = str(a or '').strip()
    if not _ADDR_RE.match(a):
        raise ValueError(f"番地の形が違います（このシートの中の 'A1' か 'A1:D10'）: {a!r}")
    a = a.replace('$', '').upper()
    # A0・XFE1・A1048577 は形だけ合っていて Excel には無い（前は素通しで、COM の失敗になっていた）
    for part in a.split(':'):
        m = re.match(r'^([A-Z]{1,3})(\d+)$', part)
        col, row = _col_num(m.group(1)), int(m.group(2))
        if not (1 <= row <= _XL_MAX_ROW):
            raise ValueError(f"行が範囲の外です（1〜{_XL_MAX_ROW:,}）: {part}")
        if not (1 <= col <= _XL_MAX_COL):
            raise ValueError(f"列が範囲の外です（A〜XFD）: {part}")
    return a


# ----------------------------------------------------------------
# 道具が COM で直接やる手（2026-09-04・手順書 20 本が回るように足した 6 手）と、追加の材料
# ----------------------------------------------------------------

_CHART_TYPES = {'column': 51, 'bar': 57, 'line': 4, 'pie': 5, 'scatter': -4169, 'area': 1}
_CSV_ENCODINGS = ('utf-8-sig', 'utf-8', 'cp932')
# 文字なのに数値・日付に見える値（材料で番地を出す。格子の見た目では数値と区別がつかない＝2026-09-04 実射で AI が見落とした）
_NUMLIKE_RE = re.compile(r'^[-－]?(?:[0-9０-９]{1,3}(?:[,，][0-9０-９]{3})+|[0-9０-９]+)(?:[.．][0-9０-９]+)?$')
_DATELIKE_RE = re.compile(r'^(?:[0-9０-９]{2,4}[-/.年][0-9０-９]{1,2}[-/.月][0-9０-９]{1,2}日?|[0-9０-９]{1,2}月[0-9０-９]{1,2}日|'
                          r'[0-9０-９]{1,2}/[0-9０-９]{1,2})$')


def _date_fmt(spec):
    """Excel 流の日付の形（yyyy-mm-dd・yyyy/mm/dd・yyyymmdd）を strftime に。既に % があればそのまま。"""
    s = str(spec or 'yyyy-mm-dd')
    if '%' in s:
        return s
    return (s.replace('yyyy', '%Y').replace('YYYY', '%Y').replace('mm', '%m').replace('MM', '%m')
             .replace('dd', '%d').replace('DD', '%d'))


def _csv_text(v, datefmt):
    """CSV の 1 マスの文字。文字はそのまま（先頭ゼロを守る）・整数は桁を落とさない・日付は datefmt。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if _is_xl_error(v):
        # エラーセルは COM が大きな負の整数で返す。そのまま書くと CSV に -2146826246 が並び、
        # 開いた人には元が #N/A だったと分からない（2026-09-04 実測）
        return _XL_ERROR_CODES.get(int(v), '#ERROR!')
    if hasattr(v, 'year') and hasattr(v, 'month') and not isinstance(v, str):
        try:
            if '%H' in datefmt or '%M' in datefmt:
                return datetime.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second).strftime(datefmt)
            return datetime.date(v.year, v.month, v.day).strftime(datefmt)
        except Exception:
            return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v)


def _csv_check_path(path):
    """export_csv の出力先の検査（純 Python）: .csv・まだ無い・フォルダがある。"""
    p = str(path or '').strip()
    if not p.lower().endswith('.csv'):
        raise ValueError("path は .csv で指定してください")
    p = os.path.abspath(p)
    if os.path.exists(p):
        raise ValueError(f"出力先が既にあります（上書きしない。別の名前を報告で提案する）: {p}")
    if not os.path.isdir(os.path.dirname(p)):
        raise ValueError(f"出力先のフォルダがありません: {os.path.dirname(p)}")
    return p


# ---- 別のファイルを読む手（2026-09-09）----------------------------------
#   財務会計の移行（2027 年 1 月テスト）の仕事は、どれも「旧システムの出力を隣に置いて照らす」形。
#   道具は「別のブックには触らない」を安全の柱にしているが、**読むことすらできなかった**（手順書
#   「2つの表の突合」も「別シートなら read_sheet」で止まっている）。ここで足すのは読むだけの手で、
#   元のファイルには 1 バイトも書かない。書き先はこのブックの新しいシート＝undo で丸ごと消える。
_READ_FILE_EXTS = ('.csv', '.tsv', '.txt', '.xlsx', '.xlsm')
_READ_FILE_MAX_ROWS = 20000        # 既定の上限（人の目で確かめられる量。超える分は limit で明示させる）
_READ_FILE_MAX_COLS = 256


def _read_file_path(path):
    """read_file の入力の検査（純 Python）→ (絶対パス, 拡張子)。"""
    p = str(path or '').strip().strip('"').strip("'")
    if not p:
        raise ValueError('path にファイルの場所を書いてください（例 "C:/受領/旧システム.csv"）。'
                         '依頼にパスが無いなら手を出さず、report で人に聞いてください')
    p = os.path.abspath(p)
    ext = os.path.splitext(p)[1].lower()
    if ext not in _READ_FILE_EXTS:
        raise ValueError(f"読めるのは {'・'.join(_READ_FILE_EXTS)} です（{ext or '拡張子なし'}）")
    if not os.path.exists(p):
        raise ValueError(f"ファイルがありません: {p}（パスを人に聞いて report に書く）")
    return p, ext


def _read_text_decode(raw):
    """CSV の文字コードを当てる → (文字列, 名前)。BOM → utf-8 → cp932 の順（役所の CSV は cp932 が多い）。"""
    if raw.startswith(b'\xef\xbb\xbf'):
        return raw.decode('utf-8-sig'), 'utf-8-sig'
    for enc in ('utf-8', 'cp932'):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode('cp932', errors='replace'), 'cp932（読めない字は置き換え）'


def _read_csv_rows(path, limit):
    """CSV/TSV を 2 次元リストに（値は全部文字のまま）→ (行, 文字コード名)。"""
    with open(path, 'rb') as f:
        raw = f.read()
    text, enc = _read_text_decode(raw)
    delim = '\t' if path.lower().endswith(('.tsv', '.txt')) and '\t' in text[:4096] else ','
    rows = []
    for r in csv.reader(io.StringIO(text), delimiter=delim):
        rows.append([('' if v is None else str(v)) for v in r[:_READ_FILE_MAX_COLS]])
        if len(rows) >= limit:
            break
    return rows, enc


def _read_book_rows(path, want_sheet, limit):
    """xlsx/xlsm を openpyxl で読む（read_only・値だけ）→ (行, シート名, 全シート名)。Excel は起こさない。"""
    import openpyxl
    bk = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        names = list(bk.sheetnames)
        name = str(want_sheet) if want_sheet else names[0]
        if name not in names:
            raise ValueError(f"そのシートがありません: {name}（ある: {', '.join(names)}）")
        rows = []
        for r in bk[name].iter_rows(values_only=True):
            rows.append(list(r[:_READ_FILE_MAX_COLS]))
            if len(rows) >= limit:
                break
        while rows and all(v in (None, '') for v in rows[-1]):     # 末尾の空行を落とす
            rows.pop()
        return rows, name, names
    finally:
        bk.close()


def _open_book_by_path(wb, path):
    """同じ Excel の中でそのファイルが開いていれば、そのブックを返す（保存前の直しも読めるように）。"""
    try:
        app = wb.Application
        for i in range(1, int(app.Workbooks.Count) + 1):
            b = app.Workbooks(i)
            try:
                if os.path.abspath(str(b.FullName)).lower() == path.lower():
                    return b
            except Exception:
                continue
    except Exception:
        pass
    return None


def _new_sheet_name(wb, want, path):
    """写し先のシート名（31 字・使えない字を落とす・同名があれば _2）。"""
    base = str(want or '').strip() or ('取込_' + os.path.splitext(os.path.basename(path))[0])
    base = re.sub(r'[:\\/?*\[\]]', '_', base)[:31].strip() or '取込'
    try:
        names = {str(s.Name) for s in wb.Sheets}
    except Exception:
        names = set()
    if base not in names:
        return base
    for i in range(2, 100):
        cand = (base[:28] + f"_{i}")
        if cand not in names:
            return cand
    raise ValueError(f"シート名が決まりません: {base}")


def _read_file_write(ws, rows):
    """読んだ行を新しいシートへ。先頭ゼロの列は文字のまま守る（0001 が 1 になるのを防ぐ）。"""
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    if not n_rows or not n_cols:
        return 0, 0, []
    grid = [list(r) + [None] * (n_cols - len(r)) for r in rows]
    text_cols = []
    for c in range(n_cols):
        for r in range(1, n_rows):                      # 見出し行は見ない
            v = grid[r][c]
            if isinstance(v, str) and len(v) > 1 and v.isdigit() and v[0] == '0':
                text_cols.append(c)
                break
    for c in text_cols:                                  # 書く前に「文字」にしておく（後からでは戻らない）
        ws.Columns(c + 1).NumberFormat = '@'
    ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Value = tuple(tuple(r) for r in grid)
    return n_rows, n_cols, [_col_letter(c + 1) for c in text_cols]


# 依頼文の中のファイルのパス（引用符つき＝空白を含む道／裸＝空白で切れる道／UNC）
_PATH_QUOTED_RE = re.compile(r'["\'「『]([A-Za-z]:[\\/][^"\'」』\n]+?\.(?:csv|tsv|txt|xlsx|xlsm))["\'」』]', re.I)
_PATH_BARE_RE = re.compile(r'(?:[A-Za-z]:[\\/]|\\\\[^\s\\]+\\)[^\s"\'「」『』、。,]*?\.(?:csv|tsv|txt|xlsx|xlsm)', re.I)
_PEEK_MAX_FILES = 2
_PEEK_ROWS = 3          # 大きいファイルは頭だけ
_PEEK_ROWS_ALL = 25     # これ以下なら全部見せる（キーの一覧が要る突合で、読み戻しの往復を作らないため）


def _paths_in_request(request):
    """依頼文からファイルのパスを拾う（純 Python）。引用符つきを先に見る＝空白を含む道を切らない。"""
    s = str(request or '')
    out = []
    for m in _PATH_QUOTED_RE.finditer(s):
        p = m.group(1).strip()
        if p not in out:
            out.append(p)
    for m in _PATH_BARE_RE.finditer(s):
        p = m.group(0).strip()
        if not any(p in q for q in out):
            out.append(p)
    return out[:_PEEK_MAX_FILES]


def _peek_file(path):
    """ファイルの姿を説明 1 行と中身に（読むだけ）→ (見出し, 行たち, 頭だけか)。

    小さいファイルは**全部**見せる。突合の弾で頭 3 行しか見せなかったら、AI は「ファイルにしか
    無いキー」を報告できず、読み戻しの往復を足すか、黙って落とすかの二択になった（2026-09-09 実射）。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        bk = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            names = list(bk.sheetnames)
            ws = bk[names[0]]
            n_rows = int(ws.max_row or 0)
            n_cols = int(ws.max_column or 0)
            want = _PEEK_ROWS_ALL if n_rows <= _PEEK_ROWS_ALL else _PEEK_ROWS
            rows = []
            for r in ws.iter_rows(values_only=True):
                rows.append(['' if v is None else str(v) for v in r[:12]])
                if len(rows) >= want:
                    break
        finally:
            bk.close()
        head = (f"{os.path.basename(path)}（Excel・シート {len(names)} 枚: {'・'.join(names[:6])}"
                f"／先頭のシート「{names[0]}」は {n_rows} 行 x {n_cols} 列）")
        return head, rows, n_rows > len(rows)
    with open(path, 'rb') as f:
        raw = f.read()
    text, enc = _read_text_decode(raw)
    delim = '\t' if path.lower().endswith(('.tsv', '.txt')) and '\t' in text[:4096] else ','
    all_rows = [[str(v) for v in r[:12]] for r in csv.reader(io.StringIO(text), delimiter=delim)]
    n_rows = len(all_rows)
    rows = all_rows if n_rows <= _PEEK_ROWS_ALL else all_rows[:_PEEK_ROWS]
    n_cols = max((len(r) for r in rows), default=0)
    head = f"{os.path.basename(path)}（{enc}・{'タブ' if delim == chr(9) else 'カンマ'}区切り・{n_rows} 行 x {n_cols} 列）"
    return head, rows, n_rows > len(rows)


def _peek_files_note(request):
    """依頼文にファイルのパスがあれば、道具が先に読んで材料に出す（2026-09-09）。

    read_file を足した日の実射では、AI が「まず読む」だけの往復を 1 回使っていた（読んでから
    列を見て、次の往復で式を組む）。ここで姿を先に渡せば、読む手と式と tidy を**同じ返事に**
    並べられる＝往復が 1 回減る（Excelコンボと同じ「AI は短い指示だけ」の形）。
    道具はまだ**読み込んでいない**（シートは作っていない）。書くのは AI が read_file を並べたとき。
    """
    paths = _paths_in_request(request)
    if not paths:
        return ''
    blocks, readable = [], 0
    for p in paths:
        full = os.path.abspath(p.replace('/', os.sep))
        if not os.path.exists(full):
            blocks.append(f"{p} → **ありません**（パスを人に聞いて report に書く。read_file を撃たない）")
            continue
        try:
            head, rows, cut = _peek_file(full)
        except Exception as ex:
            blocks.append(f"{p} → いま読めません（{ex}）")
            continue
        readable += 1
        grid = "\n".join("    " + " | ".join(r) for r in rows)
        to = ("取込_" + os.path.splitext(os.path.basename(full))[0])[:31]
        label = f"先頭 {len(rows)} 行（この先は取り込んでから read で見る）" if cut else f"全 {len(rows)} 行（これで全部）"
        # 手の見本のパスは「/」で書く（JSON の中の \ は壊れる＝AI がそのまま写して落ちる）
        blocks.append(f"{head}\n  {label}:\n{grid}\n"
                      f'  取り込む手: {{"op":"read_file","path":"{full.replace(chr(92), "/")}","to":"{to}"}}')
    tail = ("\n**ファイルを読むためだけの往復を作らないでください**（姿はもう上にあります）。read_file と、"
            "そのあとの式・tidy は同じ返事に並べます。\n"
            "ただし、**式を入れてからでないと言えないこと**（件数・突き合わせの結果・片方にしかないキー）を"
            "報告に書く必要があるときは、そこだけは次の往復で read してから done にしてください"
            "（上が「全 N 行（これで全部）」なら、ファイル側のキーはもう分かっているので読み直しは要りません）。\n"
            if readable else "\n")
    return ("\n--- 依頼に出てきたファイル（道具が先に読みました。まだ取り込んでいません） ---\n"
            + "\n".join(blocks) + tail)


def _rows_of(v):
    """Range.Value（タプルのタプル／1 セルは素の値）を 2 次元リストに。"""
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) if isinstance(r, tuple) else [r] for r in v]


def _page_count(ws):
    """ページ数（アクティブシートなら GET.DOCUMENT(50)、それ以外は改ページの数から）。読めなければ None。"""
    try:
        if ws.Parent.ActiveSheet.Name == ws.Name:
            n = ws.Parent.Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
            if n:
                return int(n)
    except Exception:
        pass
    try:
        return (int(ws.HPageBreaks.Count) + 1) * (int(ws.VPageBreaks.Count) + 1)
    except Exception:
        return None


def _print_summary(ws):
    """印刷設定の要約 1 行（追加の材料と page_setup の結果に使う）。"""
    ps = ws.PageSetup
    area = str(ps.PrintArea or '').replace('$', '') or '（未設定＝使用範囲）'
    orient = '横' if int(ps.Orientation) == 2 else '縦'
    titles = str(ps.PrintTitleRows or '').replace('$', '') or 'なし'
    if ps.Zoom is False:
        fw, ft = ps.FitToPagesWide, ps.FitToPagesTall
        # 設定の名前も添える（2026-09-08）。「横 1 ページ…に収める」と日本語で出していたら、
        # 採点係が「FitToPage の情報が材料にない」を unmet に挙げて往復を 1 回捨てた
        fit = (f"横 {int(fw) if fw else '自動'} ページ×縦 {int(ft) if ft else '自動'} ページに収める"
               f"（FitToPage=ON・FitToPagesWide={int(fw) if fw else 0}"
               f"・FitToPagesTall={int(ft) if ft else 0}）")
    else:
        fit = f"倍率 {int(ps.Zoom)}%（FitToPage=OFF＝ページに収める設定ではない）"
    foot = " / ".join(s for s in (str(ps.LeftFooter or ''), str(ps.CenterFooter or ''), str(ps.RightFooter or ''))
                      if s) or 'なし'
    pages = _page_count(ws)
    return (f"印刷: 範囲 {area} / 向き {orient} / タイトル行 {titles} / {fit} / フッター {foot}"
            + (f" / ページ数 {pages}" if pages else ""))


_HEADER_SCAN_ROWS = 6          # 見出し行を探す深さ（1 行目タイトル・2 行空き・3 行目見出し に届く）


def _id_columns(header_row):
    """見出しの行 → 番号列（会員番号・コード・No・ID …）の列位置（0 起点）の集合（純 Python）。
    判定の語は tidy の列の型（vbam_edit._column_style）と同じ＝道具の中で番号列の定義を 1 つにする。"""
    import unicodedata
    from vbam_edit import _is_id_header
    out = set()
    for j, h in enumerate(header_row or []):
        hs = unicodedata.normalize('NFKC', str(h or '')).strip()
        if hs and _is_id_header(hs):
            out.add(j)
    return out


_ZENKAKU_DIGIT_RE = re.compile(r'[０-９]')


def _zenkaku_digits(s):
    """全角の数字を含むか（番号列でも全角は乱れ＝警告に残す）。"""
    return bool(_ZENKAKU_DIGIT_RE.search(str(s or '')))


_HANKANA_ANY_RE = re.compile('[｡-ﾟ]')
_ZEN_ALNUM_RE = re.compile('[０-９Ａ-Ｚａ-ｚ]')      # 全角の数字・英字


def _charkind_facts(grid, r0, c0, hdr_idx):
    """列ごとの文字種（見出しの下）→ 1 列 1 本の文（純 Python）。空白・カナ・全角英数のどれも無い列は出さない。

    掃除の判定に要る事実＝「姓名の間の空白は全角か半角か」「半角カナはあるか」。格子の文字では全角と半角の空白が
    見分けられず、AI が確かめようと read で読み直しても同じ見え方が返るだけ＝往復 1 回（sonnet 12.8 秒）を
    捨てていた（2026-09-08）。ここに数字で出しておけば読み直しは要らない。
    """
    if hdr_idx is None or not grid or hdr_idx + 1 >= len(grid):
        return []
    out = []
    for j in range(max(len(r) for r in grid)):
        zen = han = multi = hkana = zalnum = 0
        for row in grid[hdr_idx + 1:]:
            v = row[j] if j < len(row) else None
            if not isinstance(v, str) or not v:
                continue
            core = v.strip(_WS_CHARS)
            inner = _INNER_WS_RE.findall(core)
            if inner:
                if any(len(w) > 1 for w in inner):
                    multi += 1
                if any('　' in w for w in inner):
                    zen += 1
                if any(ch in w for w in inner for ch in '  \t'):
                    han += 1
            if _HANKANA_ANY_RE.search(core):
                hkana += 1
            if _ZEN_ALNUM_RE.search(core):
                zalnum += 1
        if not (zen or han or hkana or zalnum):
            continue
        parts = []
        if zen or han:
            parts.append(f"文字の間の空白=全角 {zen}・半角 {han}" + (f"・2 つ以上 {multi}" if multi else "")
                         + ("（この列へ書き足す値の空白も半角）" if han and not zen else
                            "（この列へ書き足す値の空白も全角）" if zen and not han else ""))
        if hkana:
            parts.append(f"半角カナ {hkana}")
        if zalnum:
            parts.append(f"全角英数 {zalnum}")
        out.append(f"{_col_letter(c0 + j)}列 " + "・".join(parts))
    return out


def _pick_header_row(grid):
    """先頭数行の値（2 次元）→ 見出しらしい行の位置（0 起点）と中身。無ければ None（純 Python）。

    見出しらしさ＝「文字で埋まっている数」が最も多い行。同数なら上の行。数だけの行は見出しにしない
    （タイトル行は 1 セルだけ埋まっているので、列数の多い本物の見出しに負ける）。
    """
    best = None
    for i, row in enumerate(grid or []):
        vals = [('' if v is None else str(v)).strip() for v in row]
        filled = [v for v in vals if v != '']
        if len(filled) < 2:
            continue                               # 1 セルだけ＝タイトル行
        words = sum(1 for v in filled if not _NUMLIKE_RE.match(v))
        if words < len(filled) * 0.6:
            continue                               # ほとんど数＝データ行
        score = (len(filled), words)
        if best is None or score > best[0]:
            best = (score, i, " ".join(filled[:8]))
    return (best[1], best[2]) if best else None


def _guess_header_row(ws, scan=_HEADER_SCAN_ROWS):
    """シートの見出し行を推定 → (行番号, 中身の文字) か None。COM は 1 回。"""
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    nr, nc = min(int(ur.Rows.Count), scan), min(int(ur.Columns.Count), 40)
    grid = _rows_of(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1)).Value)
    got = _pick_header_row(grid)
    return (r0 + got[0], got[1]) if got else None


def _table_blanks(grid, r0, c0, header_idx):
    """表の中の空欄（見出しの下で、同じ列の他の行は埋まっている）→ 番地の一覧（純 Python）。

    2026-09-06 夕（シュウさんの初実射）: テスト用2 の C10（フリガナ）が空欄なのに、道具も AI も採点係も
    言わなかった。「空白に見えるゴミ」は文字の入ったセルしか見ておらず、本当に空のセルは誰も数えていなかった。
    見出し行の下から最後の値のある行まで、その列に値が 3 つ以上あり 6 割以上埋まっている列だけ見る
    （備考のようなまばらな列は数えない）。行全体が空の行（区切り）は飛ばす。見出しの無い列も見ない。
    """
    rows = [list(r) for r in (grid or [])]
    if not rows:
        return []
    h = header_idx if (header_idx is not None and 0 <= header_idx < len(rows)) else 0

    def blank(v):
        return v is None or (isinstance(v, str) and v.strip(' 　 \t\r\n') == '')

    body = [(i, r) for i, r in enumerate(rows) if i > h and not all(blank(v) for v in r)]
    if not body:
        return []
    width = max(len(r) for r in rows)
    out = []
    for j in range(width):
        head = rows[h][j] if j < len(rows[h]) else None
        if blank(head):
            continue
        cells = [(i, (r[j] if j < len(r) else None)) for i, r in body]
        filled = sum(1 for _i, v in cells if not blank(v))
        if filled < 3 or filled / len(cells) < 0.6:
            continue
        out += [f"{_col_letter(c0 + j)}{r0 + i}" for i, v in cells if blank(v)]
    return sorted(out, key=lambda a: (int(re.sub(r'\D', '', a)), a))


def _extra_materials(wb, ws):
    """materials に無い材料（他のシート・枠固定・行高・印刷・条件付き書式・外部リンク・名前・膨張・ゴミ・空欄・ファイル）。

    手順書 20 本が「道具が読んだ現物」だけで動けるように足した（2026-09-04）。読めない項目は黙って飛ばす。
    """
    lines = ["--- 追加の材料（道具が読んだ現物） ---"]
    hr = None
    try:
        others = []
        for sh in wb.Sheets:
            if sh.Name == ws.Name:
                continue
            try:
                ur = str(sh.UsedRange.Address).replace('$', '')
            except Exception:
                ur = '?'
            others.append(f"{sh.Name}({ur}{'' if sh.Visible == -1 else '・非表示'})")
        lines.append("他のシート: " + (", ".join(others) if others else "なし")
                     + "（read_sheet か read の sheet で読める。書くときは手に sheet: 名前 を足す）")
    except Exception:
        pass
    # 見出し行の推定（2026-09-06）。日本の表は 1 行目がタイトル・3 行目が見出しのことが多く、
    # Claude for Excel が「メタ情報に見出し位置が無いため 1 行目を見出しと見て、テーブル化や
    # フィルタを掛けてしまう」を自分の失敗型トップ 5 の 1 番に挙げた。こちらは道具が数えて渡す。
    try:
        hr = _guess_header_row(ws)
        if hr:
            lines.append(f"見出し行の推定: {hr[0]} 行目（{hr[1]}）"
                         + ("　← 1 行目ではありません。並べ替え・テーブル化・フィルタ・tidy の"
                            "「表の範囲」はこの行から始めてください" if hr[0] != 1 else ""))
    except Exception:
        pass
    # 突き合わせのキーの型違い（2026-09-06 深夜のループ）。書いた後の検査でしか出ていなかったので、
    # 「会員番号で突き合わせ」は 15 回中 3 回しか一発で通らなかった（毎回 1 往復よけいに使う）
    try:
        from vbam_edit import key_type_mismatch
        vals = ws.UsedRange.Value
        if vals and not isinstance(vals, (str, int, float)):
            grid = [list(r) for r in vals]
            for note in key_type_mismatch(grid)[:3]:
                lines.append("キーの型: " + note)
    except Exception:
        pass
    # 計算モードと保護（2026-09-04）。どちらも「書いたのに直らない」の正体で、材料に無いと AI が
    # 数式を疑って書き直す方向へ走る（手動計算）か、失敗の理由が分からないまま往復を使い切る（保護）
    try:
        mode = int(wb.Application.Calculation)
        if mode != -4105:               # xlCalculationAutomatic
            name = {-4135: '手動', 2: 'データテーブル以外は自動'}.get(mode, f'自動でない({mode})')
            lines.append(f"計算モード: {name}  ← 数式を書いても値が更新されません。表示されている値は古い可能性があります"
                         "（直すのは人。CLI の calc-mode auto）")
        else:
            lines.append("計算モード: 自動")
    except Exception:
        pass
    try:
        prot = []
        if bool(ws.ProtectContents):
            prot.append('セル')
        if bool(ws.ProtectDrawingObjects):
            prot.append('図形')
        if prot:
            lines.append("保護: このシートは保護されています（" + "・".join(prot) + "）"
                         "  ← 書き込みは全部失敗します。解除は人の仕事なので、材料を読んで report で伝える")
    except Exception:
        pass
    try:
        prev = wb.ActiveSheet
        switched = (prev.Name != ws.Name)
        if switched:
            ws.Activate()                    # 枠固定・倍率はウィンドウの持ち物＝そのシートを一瞬前に出して読む（inspect 用）
        try:
            win = wb.Application.ActiveWindow
            if win.FreezePanes:
                lines.append(f"枠固定: {_col_letter(int(win.SplitColumn) + 1)}{int(win.SplitRow) + 1} の左上で固定"
                             f"  表示倍率: {int(win.Zoom)}%")
            else:
                lines.append(f"枠固定: なし  表示倍率: {int(win.Zoom)}%")
        finally:
            if switched:
                prev.Activate()
    except Exception:
        pass
    try:
        h = ws.UsedRange.EntireRow.RowHeight
        lines.append("行高: " + (f"均一 {float(h):g}" if h is not None else "不揃い"))
    except Exception:
        pass
    try:
        lines.append(_print_summary(ws))
    except Exception:
        pass
    try:
        n_cf = int(ws.Cells.FormatConditions.Count)
        try:
            n_dv = int(ws.Cells.SpecialCells(-4174).Count)          # xlCellTypeAllValidation
        except Exception:
            n_dv = 0
        try:
            n_ch = int(ws.ChartObjects().Count)
        except Exception:
            n_ch = 0
        lines.append(f"条件付き書式: {n_cf} 本  入力規則: {n_dv} セル  グラフ: {n_ch} 個")
    except Exception:
        pass
    try:
        lines += _heavy_materials(wb)
    except Exception as ex:
        lines.append(f"（テーブル・ピボットの材料を読めませんでした: {ex}）")
    formulas = {}                                                    # シート名 → [(番地, 式)]
    try:
        total = 0
        for sh in wb.Worksheets:
            try:
                cells = sh.UsedRange.SpecialCells(-4123)             # xlCellTypeFormulas
            except Exception:
                continue
            for c in cells:
                total += 1
                if total > 5000:
                    break
                try:
                    formulas.setdefault(sh.Name, []).append((str(c.Address).replace('$', ''), str(c.Formula)))
                except Exception:
                    pass
            if total > 5000:
                break
    except Exception:
        pass
    try:
        links = wb.LinkSources(1)                                    # xlExcelLinks
        links = list(links) if links else []
        refs = [a for a, f in formulas.get(ws.Name, []) if '[' in f]
        if links or refs:
            lines.append(f"外部リンク: {len(links)} 本  " + "; ".join(os.path.basename(str(l)) for l in links)
                         + (f"  このシートの参照セル: {' '.join(refs[:20])}" if refs else "  （このシートには参照セルなし）"))
        else:
            lines.append("外部リンク: なし")
    except Exception:
        pass
    try:
        allf = "\n".join(f for lst in formulas.values() for _, f in lst).lower()
        rows = []
        n_broken = n_unused = 0
        for nm in wb.Names:
            try:
                full = str(nm.Name)
                ref = str(nm.RefersTo)
                vis = bool(nm.Visible)
            except Exception:
                continue
            short = full.split('!')[-1]
            if short.startswith(('_xlfn.', '_xlnm.', '_xlpm.')) and not vis:
                continue                                             # Excel が内部で持つ隠し名（関数の互換用）＝人の名前ではない
            if short in ('Print_Area', 'Print_Titles', '_FilterDatabase', 'Criteria', 'Extract', 'Database'):
                rows.append(f"  {full} = {ref}  Excel が管理（印刷範囲・タイトル行・フィルタ）＝消さない")
                continue
            broken = '#REF!' in ref.upper()
            used = short.lower() in allf
            n_broken += int(broken)
            n_unused += int((not used) and (not broken))
            rows.append(f"  {full} = {ref}  " + ("壊れている(#REF!)" if broken else ("使用中" if used else "未使用"))
                        + ("" if vis else "・非表示"))
        if rows:
            lines.append(f"名前定義（ブック全体）: {len(rows)} 本  壊れ {n_broken}  未使用 {n_unused}"
                         "（消す手は無い。人が消すなら py vba_manager.py name delete 名前）")
            lines += rows[:40]
            if len(rows) > 40:
                lines.append(f"  …他{len(rows) - 40}本")
        else:
            lines.append("名前定義（ブック全体）: なし")
    except Exception:
        pass
    try:
        ur = ws.UsedRange
        last_r = ws.Cells.Find("*", ws.Cells(1, 1), -4163, 2, 1, 2)   # xlValues, xlPart, xlByRows, xlPrevious
        last_c = ws.Cells.Find("*", ws.Cells(1, 1), -4163, 2, 2, 2)   # xlByColumns
        if last_r is not None and last_c is not None:
            data_end = f"{_col_letter(int(last_c.Column))}{int(last_r.Row)}"
            ur_end = f"{_col_letter(int(ur.Column) + int(ur.Columns.Count) - 1)}{int(ur.Row) + int(ur.Rows.Count) - 1}"
            if data_end != ur_end:
                lines.append(f"データの末尾: {data_end}（使用範囲の端 {ur_end} のほうが大きい＝膨張。"
                             "余分な行・列は row_delete／col_delete で消せる＝承認の言葉があるときだけ。"
                             "無ければ番地で報告する）")
            else:
                lines.append(f"データの末尾: {data_end}（使用範囲と一致）")
    except Exception:
        pass
    try:
        ur = ws.UsedRange
        if int(ur.Cells.CountLarge) <= 20000:
            r0, c0 = int(ur.Row), int(ur.Column)
            junk, padded, numlike, datelike, idlike = [], [], [], [], []
            grid = _rows_of(ur.Value)
            # 番号列（会員番号・コード・No・ID）の文字の数字は正常＝「数値に見える文字」に混ぜない
            # （毎回 A6〜A20 が警告に並び、規則文の「番号の先頭ゼロは文字のまま」と矛盾していた・2026-09-06 夜）
            hdr_idx = (hr[0] - r0) if hr else None
            id_cols = _id_columns(grid[hdr_idx]) if (hdr_idx is not None and 0 <= hdr_idx < len(grid)) else set()
            for i, row in enumerate(grid):
                for j, v in enumerate(row):
                    if isinstance(v, str) and v != '':
                        core = v.strip(' \u3000\u00a0\t\r\n')
                        addr = f"{_col_letter(c0 + j)}{r0 + i}"
                        if core == '':
                            junk.append(addr)
                        elif core != v:
                            padded.append(addr)
                        if _NUMLIKE_RE.match(core):
                            if j in id_cols and hdr_idx is not None and i > hdr_idx and not _zenkaku_digits(core):
                                idlike.append(addr)
                            else:
                                numlike.append(addr)
                        elif _DATELIKE_RE.match(core):
                            datelike.append(addr)
            lines.append(f"空白に見えるゴミ（全角空白・空白だけ・改行だけ）: {len(junk)} セル " + " ".join(junk[:20])
                         + ("…" if len(junk) > 20 else ""))
            lines.append(f"前後に空白のある値: {len(padded)} セル " + " ".join(padded[:20]) + ("…" if len(padded) > 20 else ""))
            kinds = _charkind_facts(grid, r0, c0, hdr_idx)
            if kinds:
                lines.append("列の中の文字種（見出しの下。掃除の判定はここで足りる＝read で読み直さない）: " + "／".join(kinds))
            lines.append(f"数値に見える文字（左寄せの数字・全角数字。番号列は除く）: {len(numlike)} セル "
                         + " ".join(numlike[:30]) + ("…" if len(numlike) > 30 else ""))
            if idlike:
                lines.append(f"番号列の文字の数字（会員番号・コードなど＝文字のままが正常。直さない）: {len(idlike)} セル "
                             + " ".join(idlike[:12]) + ("…" if len(idlike) > 12 else ""))
            lines.append(f"日付に見える文字: {len(datelike)} セル " + " ".join(datelike[:30]) + ("…" if len(datelike) > 30 else ""))
            blanks = _table_blanks(grid, r0, c0, (hr[0] - r0) if hr else None)
            lines.append(f"表の中の空欄（見出しの下で、同じ列の他の行は埋まっている）: {len(blanks)} セル "
                         + " ".join(blanks[:20]) + ("…" if len(blanks) > 20 else "")
                         + ("（隣の列から一意に決まる値＝氏名→フリガナ・対応表のあるコード→名称は埋めて report に書く。"
                            "決まらない所は埋めない＝依頼に無ければ報告の気づきに、値が要るなら要判断で人に聞く）" if blanks else ""))
    except Exception:
        pass
    try:
        if int(ws.UsedRange.Rows.Count) >= _PROFILE_MIN_ROWS:      # 格子が全部出ない大きさ＝列プロファイルで代える
            lines.append(_column_profile(ws).rstrip())
    except Exception as ex:
        lines.append(f"（列プロファイルを読めませんでした: {ex}）")
    try:
        full = str(wb.FullName)
        if os.path.isfile(full):
            lines.append(f"ファイル: {os.path.getsize(full):,} バイト  {full}")
        else:
            lines.append("ファイル: 未保存（大きさは保存後に測れる）")
    except Exception:
        pass
    return "\n".join(lines) + "\n"


def _addr(rng):
    try:
        return str(rng.Address).replace('$', '')
    except Exception:
        return '?'


def _heavy_materials(wb):
    """重い道具の材料（ブック全体）: テーブル・ピボット・スライサー・パワークエリ・データモデル。読めない項目は黙って飛ばす。"""
    lines = []
    rows = []
    for sh in wb.Worksheets:
        try:
            for lo in sh.ListObjects:
                try:
                    heads = [str(c.Name) for c in lo.ListColumns][:12]
                except Exception:
                    heads = []
                rows.append(f"  [{sh.Name}] {lo.Name} 範囲={_addr(lo.Range)} 列={','.join(heads)}")
        except Exception:
            continue
    lines.append(f"テーブル（ブック全体）: {len(rows)} 本" + ("（table の手で読む・並べ替え・絞り込み・列追加）"
                                                          if rows else "（table create で作れる）"))
    lines += rows[:20]
    rows = []
    for sh in wb.Worksheets:
        try:
            pts = list(sh.PivotTables())
        except Exception:
            continue
        for pt in pts:
            try:
                parts = {1: [], 2: [], 3: [], 4: []}
                pfs = pt.PivotFields()
                for i in range(1, int(pfs.Count) + 1):
                    f = pfs.Item(i)
                    try:
                        o = int(f.Orientation)
                    except Exception:
                        o = 0
                    if o in parts:
                        parts[o].append(str(f.Name))
                try:
                    dfs = pt.DataFields
                    parts[4] = [f"{dfs.Item(i).Name}" for i in range(1, int(dfs.Count) + 1)] or parts[4]
                except Exception:
                    pass
                try:
                    src = str(pt.SourceData)
                except Exception:
                    src = '?'
                rows.append(f"  [{sh.Name}] {pt.Name} 出力={_addr(pt.TableRange2)} 元={src}"
                            f" 行={','.join(parts[1]) or '-'} 列={','.join(parts[2]) or '-'}"
                            f" フィルタ={','.join(parts[3]) or '-'} 値={','.join(parts[4]) or '-'}")
            except Exception as ex:
                rows.append(f"  [{sh.Name}] {getattr(pt, 'Name', '?')}（詳細を読めない: {ex}）")
    lines.append(f"ピボット（ブック全体）: {len(rows)} 本" + ("（pivot_field / pivot_calc の手で。中身は pivot_calc get_data）"
                                                          if rows else "（pivot create で作れる）"))
    lines += rows[:20]
    try:
        names = []
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                names.append(f"{sl.Name}（{sc.SourceName}）")
        lines.append(f"スライサー: {len(names)} 個 " + " ".join(names[:20]))
    except Exception:
        pass
    try:
        qs = wb.Queries
        names = [str(qs.Item(i).Name) for i in range(1, int(qs.Count) + 1)]
        lines.append(f"パワークエリ: {len(names)} 本 " + " ".join(names[:20])
                     + ("（表に出すには powerquery load）" if names else ""))
    except Exception:
        pass
    try:
        model = wb.Model
        tables = [str(model.ModelTables.Item(i).Name) for i in range(1, int(model.ModelTables.Count) + 1)]
        try:
            n_meas = int(model.ModelMeasures.Count)
        except Exception:
            n_meas = 0
        lines.append(f"データモデル: テーブル {len(tables)} 本 " + " ".join(tables[:20]) + f"  メジャー {n_meas} 本")
    except Exception:
        pass
    return lines


# ----------------------------------------------------------------
# 大きい表の手（2026-09-04）: AI に行を触らせない。
#   見る … 列プロファイル（列ごとの型・乱れ・形・種類。格子に出ていない行の代わり）
#   書く … normalize（列に規則を当てる。道具が全行を処理）／fill（数式を末尾まで）／find_replace（一括置換）
#   確かめる … 道具が件数と例と「結果の列の様子」を返す
# それまでは「materials で全体を見る → AI が全行の値を JSON で吐く」だったので、行数がそのまま AI の出力量になり
# 数百行で破綻していた（練習台は 9〜40 行だった）。
# ----------------------------------------------------------------
_PROFILE_MAX_ROWS = 20000
_PROFILE_MAX_COLS = 60
_PROFILE_MIN_ROWS = 41                # materials が格子を全部出さない行数（40 行超）からプロファイルを付ける
_XL_ERROR_CODES = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
                   -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!'}
_WS_CHARS = ' \u3000\u00a0\t\r\n'                      # 半角空白・全角空白・ノーブレークスペース・タブ・改行
_DASHES = '\u2010\u2011\u2012\u2013\u2014\u2015\u2212\uff0d'   # ‐ ‑ ‒ – — ― − －（ハイフンに見える記号）
_ZENKAKU_RE = re.compile('[\uff01-\uff5e\u3000]')             # 全角の英数字・記号・全角空白
_HANKANA_RE = re.compile('[｡-ﾟ]+')                 # 半角カナの連なり（｡｢｣､･ｦ〜ﾟ）。NFKC で全角に＝濁点・半濁点も合成される
_INNER_WS_RE = re.compile(r'(?<=\S)[ 　 	]+(?=\S)')   # 文字と文字のあいだの空白（半角・全角・NBSP・タブ）の連なり。前後は trim の仕事
# kana_zenkaku / space_zenkaku / space_hankaku は掃除の規則（2026-09-08）: 「氏名の空白を全角 1 つに」「半角カナを全角に」を
# AI が 30 セル書き直していた（sonnet で出力 1,200 トークン＝25 秒）。規則にすれば AI は列と規則名だけ返す
_NORMALIZE_RULES = ('trim', 'hankaku', 'hyphen', 'number', 'date', 'regex', 'as_text', 'fill_down',
                    'kana_zenkaku', 'space_zenkaku', 'space_hankaku', 'lower', 'phone', 'postal', 'corp')
# 書き方だけを変える規則（値の意味は変わらない）。「書き方をそろえて」の依頼なら承認の言葉が無くても当ててよい
# （2026-09-11 夜・試験: 経費精算の「そろえてください」で門が normalize を全部止め、75 秒・5 往復で何も直らなかった）
_FORMAT_RULES = frozenset(('trim', 'hankaku', 'hyphen', 'number', 'date', 'kana_zenkaku', 'space_zenkaku',
                           'space_hankaku', 'lower', 'phone', 'postal', 'corp'))
# 会社の種類の略（NFKC の後。（株）・㈱ は (株) になる）→ 正式の名前
_CORP_ABBR = {'(株)': '株式会社', '(有)': '有限会社', '(合)': '合同会社', '(資)': '合資会社', '(名)': '合名会社'}
_CORP_NAMES_RE = re.compile(r'\s*(株式会社|有限会社|合同会社|合資会社|合名会社)\s*')
_NUM_UNIT_RE = re.compile(r'\s*(円|個|本|枚|台|冊|箱|点|件)$')    # 数の後ろの単位（「12個」→ 12・2026-09-11 夜）
_DATE_TEXT_RE = re.compile(r'^\s*(\d{4})[/\-.年]\s*(\d{1,2})[/\-.月]\s*(\d{1,2})\s*日?\s*$')
_NUM_TEXT_RE = re.compile(r'^[-+]?\d+(?:\.\d+)?$')


def _is_xl_error(v):
    return isinstance(v, int) and not isinstance(v, bool) and v in _XL_ERROR_CODES


def _is_date_value(v):
    return hasattr(v, 'year') and hasattr(v, 'month') and not isinstance(v, str)


def _text_of(v):
    """値の文字の形（控えに使う）。日付は yyyy/m/d、整数の float は桁を落とさない。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if _is_date_value(v):
        # 時刻を落とすと、控え（as_text）や差分の照合で 9:30 の予定が 0:00 に見える（2026-09-04）
        if getattr(v, 'hour', 0) or getattr(v, 'minute', 0) or getattr(v, 'second', 0):
            return f"{v.year}/{v.month}/{v.day} {v.hour}:{v.minute:02d}" + (
                f":{v.second:02d}" if getattr(v, 'second', 0) else "")
        return f"{v.year}/{v.month}/{v.day}"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _hankaku(s):
    return ''.join(chr(ord(ch) - 0xFEE0) if 0xFF01 <= ord(ch) <= 0xFF5E else (' ' if ch == '　' else ch) for ch in s)


def _shape_of(s):
    """文字の形。数字→9・全角数字→９・英字→A・かな漢字→漢（連続は 1 つに畳む）。列の形の分布に使う。"""
    out = []
    for ch in str(s)[:24]:
        o = ord(ch)
        if '0' <= ch <= '9':
            t = '9'
        elif 0xFF10 <= o <= 0xFF19:
            t = '９'
        elif ch.isascii() and ch.isalpha():
            t = 'A'
        elif 0xFF21 <= o <= 0xFF5A:
            t = 'Ａ'
        elif ch.isascii() or o in (0x3000,) or ch in _DASHES or 0xFF01 <= o <= 0xFF5E:
            t = ch
        else:
            t = '漢'
        if t in ('A', 'Ａ', '漢') and out and out[-1] == t:
            continue
        out.append(t)
    return ''.join(out) + ('…' if len(str(s)) > 24 else '')


def _profile_columns(values, formulas=None, r0=1, c0=1):
    """列プロファイル（純 Python）。values は 2 次元リスト（1 行目が見出しならそれを名前にする）。
    formulas は同じ形（数式のセルは '=' 始まりの文字）か None。1 列 1 行の文で返す。"""
    if not values:
        return []
    ncols = max(len(r) for r in values)
    head = list(values[0]) + [None] * (ncols - len(values[0]))
    nonempty = [v for v in head if v not in (None, '')]
    has_header = len(nonempty) >= 2 and sum(1 for v in nonempty if isinstance(v, str)) * 2 >= len(nonempty)
    start = 1 if has_header else 0
    lines = []
    for j in range(min(ncols, _PROFILE_MAX_COLS)):
        n = {'num': 0, 'text': 0, 'date': 0, 'blank': 0, 'formula': 0, 'bool': 0, 'err': 0}
        dirty = {'前後空白': 0, '全角': 0, '数値に見える文字': 0, '日付に見える文字': 0, '空白だけ': 0, 'ハイフンの種類': 0}
        ex = {}
        shapes, tops = {}, {}
        lead0 = 0
        nmin = nmax = None
        for i in range(start, len(values)):
            row = values[i]
            v = row[j] if j < len(row) else None
            addr = f"{_col_letter(c0 + j)}{r0 + i}"
            f = formulas[i][j] if (formulas is not None and i < len(formulas) and j < len(formulas[i])) else None
            if isinstance(f, str) and f.startswith('='):
                n['formula'] += 1
                continue
            if v is None or v == '':
                n['blank'] += 1
                continue
            if isinstance(v, bool):
                n['bool'] += 1
                continue
            if _is_xl_error(v):
                n['err'] += 1
                continue
            if _is_date_value(v):
                n['date'] += 1
                continue
            if isinstance(v, (int, float)):
                n['num'] += 1
                fv = float(v)
                nmin = fv if nmin is None else min(nmin, fv)
                nmax = fv if nmax is None else max(nmax, fv)
                continue
            s = str(v)
            n['text'] += 1
            core = s.strip(_WS_CHARS)
            if core == '':
                dirty['空白だけ'] += 1
                ex.setdefault('空白だけ', (addr, s))
                continue
            if core != s:
                dirty['前後空白'] += 1
                ex.setdefault('前後空白', (addr, s))
            if _ZENKAKU_RE.search(core):
                dirty['全角'] += 1
                ex.setdefault('全角', (addr, s))
            if any(ch in _DASHES for ch in core) or re.search(r'\dー\d', core):
                dirty['ハイフンの種類'] += 1
                ex.setdefault('ハイフンの種類', (addr, s))
            hk = _hankaku(core)
            if _NUMLIKE_RE.match(hk):
                dirty['数値に見える文字'] += 1
                ex.setdefault('数値に見える文字', (addr, s))
                if re.match(r'[-+]?0\d', hk.replace(',', '')):
                    lead0 += 1
            elif _DATELIKE_RE.match(hk):
                dirty['日付に見える文字'] += 1
                ex.setdefault('日付に見える文字', (addr, s))
            sh = _shape_of(core)
            shapes[sh] = shapes.get(sh, 0) + 1
            tops[core] = tops.get(core, 0) + 1
        total = sum(n.values())
        if total == 0:
            continue
        name = _text_of(head[j]) if has_header and head[j] not in (None, '') else '（見出しなし）'
        parts = [f"{k} {c}" for k, c in (('数値', n['num']), ('文字', n['text']), ('日付', n['date']), ('数式', n['formula']),
                                         ('空白', n['blank']), ('エラー', n['err']), ('真偽', n['bool'])) if c]
        line = f"  {_col_letter(c0 + j)} {name}: " + " / ".join(parts)
        if n['text']:
            line += f"  種類 {len(tops)}"
            top_sh = sorted(shapes.items(), key=lambda kv: -kv[1])[:3]
            line += "  形: " + " / ".join(f"{s} ×{c}" for s, c in top_sh)
            if len(tops) <= 20 and n['text'] >= 5:
                top_v = sorted(tops.items(), key=lambda kv: -kv[1])[:3]
                if top_v and top_v[0][1] >= 2:                     # 全部 1 件ずつ（名前など）なら出さない
                    line += "  多い値: " + ", ".join(f"{_text_of(v)[:12]} ×{c}" for v, c in top_v)
        if n['num'] and nmin is not None:
            line += f"  数値の範囲 {nmin:,.10g}〜{nmax:,.10g}"
        dl = [f"{k} {c}（例 {ex[k][0]} {ex[k][1][:16]!r}）" if k in ex else f"{k} {c}" for k, c in dirty.items() if c]
        if dl:
            line += "  乱れ: " + "・".join(dl)
        if lead0:
            line += f"  先頭ゼロ {lead0}（番号なら文字のままが正しい）"
        lines.append(line)
    if ncols > _PROFILE_MAX_COLS:
        lines.append(f"  …他 {ncols - _PROFILE_MAX_COLS} 列（先頭 {_PROFILE_MAX_COLS} 列まで）")
    return lines


def _column_profile(ws, max_rows=_PROFILE_MAX_ROWS):
    """使用範囲の列プロファイルを文で返す（COM で値と数式を各 1 回読む）。"""
    ur = ws.UsedRange
    r0, c0 = int(ur.Row), int(ur.Column)
    nr, nc_all = int(ur.Rows.Count), int(ur.Columns.Count)
    nc = min(nc_all, _PROFILE_MAX_COLS)
    rows = min(nr, max_rows)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + rows - 1, c0 + nc - 1))
    values = _rows_of(rng.Value)
    formulas = None
    try:
        if rng.HasFormula is not False:               # True か None（混在）なら数式を読む
            formulas = _rows_of(rng.Formula)
    except Exception:
        formulas = None
    scope = f"先頭 {rows:,} 行（全 {nr:,} 行）" if rows < nr else f"全 {nr:,} 行"
    lines = [f"--- 列プロファイル（{scope}・{nc} 列。格子に出ていない行の代わり。乱れの番地は例 1 つ） ---"]
    lines += _profile_columns(values, formulas, r0, c0)
    lines.append("  行が多い表は normalize（列に規則を当てる）・fill（数式を末尾まで）・find_replace で直す。全行を write_grid で書かない")
    return "\n".join(lines) + "\n"


def _parse_date_text(s):
    t = _hankaku(s)
    m = _DATE_TEXT_RE.match(t)
    if not m:
        return _parse_wareki_text(t) or _parse_en_date(t)
    try:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return datetime.datetime(y, mo, d, tzinfo=datetime.timezone.utc)     # 素の datetime は pywin32 が 9 時間ずらす
    except ValueError:
        return None


# 和暦（2026-09-06・Claude for Excel の聞き取りで「元年・全角・R6 略記は DATEVALUE が通らない」と出た。
# こちらの normalize date も西暦 4 桁だけで、日本の名簿の「令和6年3月1日」「R6.3.1」「平成元年」を落としていた）。
# 全角は先に _hankaku で半角に落ちているので、ここでは半角だけ見る。
_WAREKI_BASE = {'令和': 2018, 'R': 2018, '平成': 1988, 'H': 1988, '昭和': 1925, 'S': 1925, '大正': 1911, 'T': 1911}
_WAREKI_RE = re.compile(r'^\s*(令和|平成|昭和|大正|[RHST])\s*(元|\d{1,2})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})\s*日?\s*$')


def _parse_wareki_text(s):
    """「令和6年3月1日」「R6.3.1」「平成元年1月8日」→ datetime。読めなければ None（純 Python）。"""
    m = _WAREKI_RE.match(str(s or ''))
    if not m:
        return None
    base = _WAREKI_BASE.get(m.group(1))
    if base is None:
        return None
    try:
        yy = 1 if m.group(2) == '元' else int(m.group(2))
        return datetime.datetime(base + yy, int(m.group(3)), int(m.group(4)),
                                 tzinfo=datetime.timezone.utc)
    except ValueError:
        return None


def _normalize_spec(act):
    """normalize の手の検査（純 Python）。(range, rules[(名, 引数)], to, header, overwrite, keep_leading_zero)"""
    rng = _check_addr(act.get('range'))
    raw = act.get('rules')
    raw = [] if raw is None else ([raw] if isinstance(raw, str) else raw)
    if not isinstance(raw, list):
        raise ValueError("normalize の rules は配列（規則名の文字か {\"regex\":{…}}）")
    rules = []
    for r in raw:
        if isinstance(r, str):
            name = r.strip().lower()
            if name not in _NORMALIZE_RULES:
                # 綴りの近い規則名は読み替える（kata_zenkaku → kana_zenkaku。AI の打ち間違い 1 字で手 7 本が連鎖停止し、
                # 往復が尽きて不合格になった・2026-09-08）。遠い名前は従来どおり断る
                near = difflib.get_close_matches(name, [n for n in _NORMALIZE_RULES if n != 'regex'], n=1, cutoff=0.75)
                if near:
                    name = near[0]
            if name not in _NORMALIZE_RULES or name == 'regex':
                raise ValueError(f"normalize の規則は {'/'.join(n for n in _NORMALIZE_RULES if n != 'regex')} か "
                                 f"{{\"regex\":{{\"search\":…,\"replace\":…}}}}（{r!r}）")
            rules.append((name, None))
        elif isinstance(r, dict) and 'regex' in r:
            rx = r['regex'] if isinstance(r['regex'], dict) else {}
            search, repl = rx.get('search'), rx.get('replace')
            if not isinstance(search, str) or not search or not isinstance(repl, str):
                raise ValueError('regex は {"regex":{"search":"^(\\\\d{3})(\\\\d{4})$","replace":"\\\\1-\\\\2"}} の形')
            try:
                pat = re.compile(search)
            except re.error as ex:
                raise ValueError(f"regex の search が正しくない: {ex}")
            # 置換の $1・${1}（JavaScript・Excel の書き方）も \1 と読む（2026-09-11: "$1-$2-$3" の文字が 13 セルに入った）
            repl = re.sub(r'\$\{?(\d+)\}?', lambda m: '\\g<' + m.group(1) + '>', repl)
            try:
                pat.sub(repl, '')
            except (re.error, IndexError) as ex:
                raise ValueError(f"regex の replace が正しくない（search に無い番号を使っている、など）: {ex}")
            rules.append(('regex', (pat, repl)))
        else:
            raise ValueError("normalize の rules の要素は規則名の文字か {\"regex\":{\"search\":…,\"replace\":…}}")
    to = act.get('to')
    to = _check_addr(to) if to else None
    if to and ':' in to:
        raise ValueError('to は書き先の左上 1 セル（例 "H2"）')
    overwrite = bool(act.get('overwrite'))
    if not to and not overwrite:
        raise ValueError('元の列に書くには "overwrite": true が要ります（依頼・補足・手順書に承認の言葉があるときだけ）。'
                         '控えを残すなら "to" で右の列へ')
    header = act.get('header')
    keep_zero = act.get('keep_leading_zero', True)
    return rng, rules, to, (str(header) if header not in (None, '') else None), overwrite, bool(keep_zero)


def _normalize_value(v, rules, keep_zero=True):
    """1 セルの値に規則を順に当てる → (新しい値, 当たった規則名)。fill_down は列単位なので _normalize_column が扱う。"""
    applied = []
    cur = v
    for name, arg in rules:
        if name in ('fill_down', 'phone'):
            continue                                   # 列単位の規則（_normalize_column が扱う）
        if not isinstance(cur, str):
            if name == 'as_text' and cur is not None:
                new = _text_of(cur)
                if new != cur:
                    applied.append(name)
                    cur = new
            continue                                   # 数値・日付には文字の規則は当たらない
        s = cur
        if name == 'trim':
            new = s.strip(_WS_CHARS)
        elif name == 'hankaku':
            new = _hankaku(s)
        elif name == 'hyphen':
            new = ''.join('-' if ch in _DASHES else ch for ch in s)
            new = re.sub(r'(?<=\d)ー(?=\d)', '-', new)
        elif name == 'regex':
            pat, repl = arg
            new = pat.sub(repl, s)
        elif name == 'number':
            t = _hankaku(s.strip(_WS_CHARS)).replace(',', '')
            new = s
            # 通貨の印と「万」（2026-09-11: ¥3,000,000・95万円・500万 が文字のまま残ったのに合格と出ていた）
            t = re.sub(r'^[¥￥$]\s*', '', t)
            t = _NUM_UNIT_RE.sub('', t)
            man = re.match(r'^([-+]?\d+(?:\.\d+)?)万$', t)
            if man:
                num = float(man.group(1)) * 10000
                new = int(num) if num.is_integer() else num
            elif _NUM_TEXT_RE.match(t):
                if keep_zero and re.match(r'[-+]?0\d', t):
                    pass                               # 先頭ゼロの番号は文字のまま
                elif len(re.sub(r'\D', '', t)) > 15:
                    # Excel の数値は 15 桁まで。16 桁超（口座・カード・伝票番号）を数値にすると
                    # 下位の桁が黙って変わる（"12345678901234567890" → 12345678901234567168・2026-09-04 実測）
                    pass                               # 桁の多い番号は文字のまま
                else:
                    num = float(t)
                    new = int(num) if num.is_integer() and 'e' not in t.lower() else num
        elif name == 'date':
            d = _parse_date_text(s)
            new = d if d is not None else s
        elif name == 'lower':
            new = s.lower()                            # 英字を小文字に（メールアドレス・2026-09-11）
        elif name == 'postal':
            # 郵便番号を 000-0000 に（〒・全角・区切りなし・空白も読む。7 桁でなければそのまま・2026-09-11 夜・試験で
            # 区切りなしの 0100951 と 〒 が残ったまま合格と出た）
            d = re.sub(r'[\s\-‐‑‒–—―−ー－〒]', '', unicodedata.normalize('NFKC', s))
            new = f"{d[:3]}-{d[3:]}" if re.fullmatch(r'\d{7}', d) else s
        elif name == 'corp':
            # 会社名の（株）・㈱ → 株式会社、(有) → 有限会社、会社の種類の前後の空白を取る（2026-09-11 夜・試験: 「株式会社 アクア
            # テック」「（株）北斗精機」が残り、採点係の差し戻しで往復が増えた）。種類の前後の位置（前株・後株）は変えない
            t = s
            for ab, full in _CORP_ABBR.items():
                t = re.sub(r'[（(]\s*' + re.escape(ab[1]) + r'\s*[）)]', full, t)
            t = t.replace('㈱', '株式会社').replace('㈲', '有限会社')
            new = _CORP_NAMES_RE.sub(lambda m: m.group(1), t).strip(_WS_CHARS)
        elif name == 'kana_zenkaku':
            new = _HANKANA_RE.sub(lambda m: unicodedata.normalize('NFKC', m.group(0)), s)
        elif name in ('space_zenkaku', 'space_hankaku'):
            new = _INNER_WS_RE.sub('　' if name == 'space_zenkaku' else ' ', s)
        else:                                          # as_text（文字はそのまま）
            new = s
        if new != cur:
            applied.append(name)
            cur = new
    return cur, applied


def _normalize_column(col, rules, keep_zero=True, skip=None):
    """1 列ぶん。col は値の並び、skip は触らない位置（数式）。→ (新しい並び, 規則ごとの件数, 例[(位置, 前, 後)])"""
    names = [n for n, _ in rules]
    out = list(col)
    counts = {}
    examples = []
    for i, v in enumerate(col):
        if (skip and skip[i]) or v is None or v == '':
            continue
        new, applied = _normalize_value(v, rules, keep_zero)
        if applied:
            out[i] = new
            for a in applied:
                counts[a] = counts.get(a, 0) + 1
            if len(examples) < 8:
                examples.append((i, v, new))
    if 'phone' in names:
        learned = _phone_learn(out)
        for i, v in enumerate(out):
            if (skip and skip[i]) or not isinstance(v, str):
                continue
            new = _phone_format(v, learned)
            if new != v:
                out[i] = new
                counts['phone'] = counts.get('phone', 0) + 1
                if len(examples) < 8:
                    examples.append((i, v, new))
    if 'fill_down' in names:
        last = None
        for i, v in enumerate(out):
            blank = v is None or (isinstance(v, str) and v.strip(_WS_CHARS) == '')
            if blank:
                if last is not None and not (skip and skip[i]):
                    out[i] = last
                    counts['fill_down'] = counts.get('fill_down', 0) + 1
                    if len(examples) < 8:
                        examples.append((i, v, last))
            else:
                last = v
    return out, counts, examples


def _cell_write_back(cell, new):
    """1 セルに値を書き、文字が数値や日付に化けたら文字列書式にして書き直す。日付は yyyy/m/d の書式に。"""
    if new is None or new == '':
        cell.ClearContents()
        return
    if isinstance(new, str):
        cell.Value = new
        try:
            back = cell.Value
        except Exception:
            back = new
        if not isinstance(back, str) or back != new:
            cell.NumberFormat = "@"
            cell.Value = new
        return
    from vbam_edit import _untext_one
    _untext_one(cell, new)
    cell.Value = new
    if _is_date_value(new):
        try:
            cell.NumberFormat = "yyyy/m/d"
        except Exception:
            pass


def _do_normalize(ws, act):
    """normalize の手（COM）。(表示名, ok, 出力)"""
    rng_addr, rules, to, header, overwrite, keep_zero = _normalize_spec(act)
    src = ws.Range(rng_addr)
    nr, nc = int(src.Rows.Count), int(src.Columns.Count)
    if nr * nc > 200000:
        raise ValueError(f"normalize は 1 回 20 万セルまで（{nr:,}×{nc}）。列ごとに分ける")
    values = _rows_of(src.Value)
    formulas = None
    try:
        if src.HasFormula is not False:
            formulas = _rows_of(src.Formula)
    except Exception:
        formulas = None
    new_grid = [list(r) + [None] * (nc - len(r)) for r in values]
    total_counts, examples, n_changed = {}, [], 0
    for j in range(nc):
        col = [new_grid[i][j] if j < len(new_grid[i]) else None for i in range(nr)]
        skip = [bool(formulas is not None and i < len(formulas) and j < len(formulas[i])
                     and isinstance(formulas[i][j], str) and str(formulas[i][j]).startswith('=')) for i in range(nr)]
        new_col, counts, ex = _normalize_column(col, rules, keep_zero, skip)
        for i in range(nr):
            if new_col[i] != col[i]:
                n_changed += 1
            new_grid[i][j] = new_col[i]
        for k, c in counts.items():
            total_counts[k] = total_counts.get(k, 0) + c
        examples += [(f"{_col_letter(src.Column + j)}{src.Row + i}", b, a) for i, b, a in ex]
    label = f"normalize {rng_addr}"
    lines = []
    if to:
        dr, dc = int(ws.Range(to).Row), int(ws.Range(to).Column)
        dest = ws.Range(ws.Cells(dr, dc), ws.Cells(dr + nr - 1, dc + nc - 1))
        n_exist = int(ws.Parent.Application.WorksheetFunction.CountA(dest))
        if n_exist and not overwrite:
            raise ValueError(f"書き先 {_addr(dest)} に値が {n_exist} セルある。空いている列を to にするか、上書きなら \"overwrite\": true")
        # 数式のセルは書き先に写さない（元の列の数式を別の場所へ動かすことになる）
        for i in range(nr):
            for j in range(nc):
                if formulas is not None and isinstance(formulas[i][j], str) and str(formulas[i][j]).startswith('='):
                    new_grid[i][j] = None
        # 書き先が文字列書式（@。CSV 取り込みの列に多い）だと、数値・日付が文字のまま入る（2026-09-04）
        from vbam_edit import _untext_grid
        _untext_grid(ws, dest, dr, dc, new_grid)
        dest.Value = [[None if (v == '' ) else v for v in row] for row in new_grid]
        # 文字が数値・日付に化けたセルは文字列書式で書き直す。日付は yyyy/m/d
        back = _rows_of(dest.Value)
        fixed = 0
        for i in range(nr):
            for j in range(nc):
                want = new_grid[i][j]
                got = back[i][j] if i < len(back) and j < len(back[i]) else None
                cell = ws.Cells(dr + i, dc + j)
                if isinstance(want, str) and want != '' and (not isinstance(got, str) or got != want):
                    cell.NumberFormat = "@"
                    cell.Value = want
                    fixed += 1
                elif _is_date_value(want):
                    try:
                        cell.NumberFormat = "yyyy/m/d"
                    except Exception:
                        pass
        if header and dr > 1:
            hcell = ws.Cells(dr - 1, dc)
            if hcell.Value in (None, ''):
                hcell.Value = header
                lines.append(f"見出し: {_addr(hcell)}「{header}」")
            else:
                lines.append(f"見出し: {_addr(hcell)} に既に「{_text_of(hcell.Value)}」があるので書かなかった")
        label += f" → {_addr(dest)}"
        lines.insert(0, f"書き先: {_addr(dest)}（{nr:,} 行 × {nc} 列を写した。変えた値 {n_changed:,}）"
                     + (f"・文字のまま守ったセル {fixed}" if fixed else ""))
        result_grid = _rows_of(dest.Value)      # 書こうとした値でなく、入った現物をプロファイルする
    else:
        wrote = 0
        for i in range(nr):
            for j in range(nc):
                if new_grid[i][j] != (values[i][j] if j < len(values[i]) else None):
                    _cell_write_back(ws.Cells(src.Row + i, src.Column + j), new_grid[i][j])
                    wrote += 1
        label += "（元の列に上書き）"
        lines.insert(0, f"元の列に書いた: {wrote:,} セルを変えた（{nr:,} 行 × {nc} 列。数式のセルは触っていない）")
        result_grid = _rows_of(src.Value)
    rule_names = {'trim': '前後の空白', 'hankaku': '全角→半角', 'hyphen': 'ハイフン統一', 'number': '数値化', 'date': '日付化',
                  'regex': '正規表現', 'as_text': '文字のまま写す', 'fill_down': '上の値で埋める'}
    lines.append("規則ごとの件数: " + (" / ".join(f"{rule_names.get(k, k)} {c:,}" for k, c in total_counts.items()) or "変えたセルなし"))
    if examples:
        lines.append("例: " + " / ".join(f"{a} {_text_of(b)[:18]!r}→{_text_of(n)[:18]!r}" for a, b, n in examples[:6]))
    prof = _profile_columns([[f"{_col_letter(src.Column + j)}" for j in range(nc)]] + result_grid, None,
                            (ws.Range(to).Row if to else src.Row) - 1, ws.Range(to).Column if to else src.Column)
    if prof:
        lines.append("結果の列の様子:")
        lines += prof
    return label, True, "\n".join(lines)


def _do_fill(ws, act):
    """fill の手（COM）: 先頭セルに数式か値を書いて末尾まで伸ばす。(表示名, ok, 出力)"""
    from vbam_edit import _write_value, _untext_one
    rng = ws.Range(_check_addr(act.get('range')))
    formula, value = act.get('formula'), act.get('value')
    if formula in (None, '') and value in (None, ''):
        raise ValueError('fill には formula（"=…"）か value が要ります')
    right = bool(act.get('right'))
    nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
    if nr > 1 and nc > 1:
        raise ValueError("fill の range は 1 列（下へ伸ばす）か 1 行（\"right\": true で右へ）")
    n_cells = nr if not right else nc
    if n_cells < 2:
        raise ValueError("fill は 2 セル以上の範囲に（先頭セル＋伸ばす先）")
    if n_cells > _FILL_MAX_CELLS:
        # 列全体（B2:B1048576）を「末尾まで」と読んだ手が通ると、使用範囲が 100 万行に膨らみ、
        # 以後の材料・画像・保存が重くなる。normalize には 20 万セルの関所があり fill には無かった
        raise ValueError(f"fill は 1 回 {_FILL_MAX_CELLS:,} セルまで（{n_cells:,} セル）。"
                         "表の最終行までの番地で指定する（列全体 B2:B1048576 のような書き方はしない）")
    if right:
        rest = ws.Range(ws.Cells(rng.Row, rng.Column + 1), ws.Cells(rng.Row, rng.Column + nc - 1))
    else:
        rest = ws.Range(ws.Cells(rng.Row + 1, rng.Column), ws.Cells(rng.Row + nr - 1, rng.Column))
    first = ws.Cells(rng.Row, rng.Column)
    n_exist = int(ws.Parent.Application.WorksheetFunction.CountA(rest))
    # 先頭セルは検査の外だった＝人が入れた値・数式を無条件に潰していた（2026-09-04）
    v = act.get('formula') if act.get('formula') not in (None, '') else act.get('value')
    try:
        cur = first.Formula
        head_used = cur is not None and str(cur) != '' and str(cur) != str(v)
        # 先頭に「これから書くのと同じ数式」があるとき（人の数式を末尾まで伸ばす頼み方）は上書きではない
        # （一律に「値がある」で断っていた＝以前通っていた頼み方が断られていた・2026-09-04 夜）
    except Exception:
        head_used = False
    if (n_exist or head_used) and not act.get('overwrite'):
        where = []
        if head_used:
            where.append(f"先頭 {_addr(first)}")
        if n_exist:
            where.append(f"{_addr(rest)} に {n_exist:,} セル")
        raise ValueError("fill 先に値がある（" + " / ".join(where)
                         + "）。上書きするなら \"overwrite\": true（承認があるときだけ）")
    if formula not in (None, '') and not str(formula).startswith('='):
        raise ValueError("formula は '=' で始める（値なら \"value\"）")
    v = formula if formula not in (None, '') else value
    _untext_one(first, v)      # 書き先が文字列書式（@）なら標準に戻す（数式が文字で入る・2026-09-04 実機）
    _write_value(first, v)
    if right:
        rng.FillRight()
    else:
        rng.FillDown()
    n_err = 0
    try:
        n_err = int(rng.SpecialCells(-4123, 16).Count)         # xlCellTypeFormulas, xlErrors
    except Exception:
        n_err = 0
    cells = [ws.Cells(rng.Row, rng.Column + k) if right else ws.Cells(rng.Row + k, rng.Column)
             for k in list(range(min(3, nr if not right else nc))) + list(range(max(3, (nr if not right else nc) - 2), nr if not right else nc))]
    seen = []
    for c in cells:
        try:
            seen.append(f"{_addr(c)}={str(c.Text)[:20]!r}")
        except Exception:
            pass
    txt = (f"fill（{'右' if right else '下'}）: {_addr(rng)} に {_text_of(v)[:60]!r} を伸ばした（{(nr if not right else nc):,} セル）"
           f"  エラー {n_err}  見え方: " + " ".join(seen))
    return f"fill {_addr(rng)}", True, txt


_IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.emf', '.wmf')
_XL_ERRORS = {2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!', 2029: '#NAME?', 2036: '#NUM!', 2042: '#N/A'}


def _eval_text(v):
    """Evaluate の戻りを人が読む文字に（エラーは名前・配列は先頭だけ・純 Python）。"""
    if isinstance(v, int) and not isinstance(v, bool) and v < -2146820000:
        return _XL_ERRORS.get(v & 0xFFFF, f"エラー({v & 0xFFFF})")
    if isinstance(v, tuple):
        flat = [x for row in v for x in (row if isinstance(row, tuple) else (row,))]
        head = " ".join(_text_of(x) for x in flat[:12])
        return f"配列 {len(flat)} 個: {head}" + (" …" if len(flat) > 12 else "")
    return _text_of(v)


# ----------------------------------------------------------------
# 英語の日付（2026-09-11: 「March 15, 2026」が date の規則で読めず、文字のまま残った）
# ----------------------------------------------------------------
_EN_MONTHS = {m: i for i, ms in enumerate((('jan', 'january'), ('feb', 'february'), ('mar', 'march'),
                                            ('apr', 'april'), ('may',), ('jun', 'june'), ('jul', 'july'),
                                            ('aug', 'august'), ('sep', 'sept', 'september'), ('oct', 'october'),
                                            ('nov', 'november'), ('dec', 'december')), 1) for m in ms}
_EN_DATE_RE = re.compile(r'^\s*([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\s*$')
_EN_DATE_RE2 = re.compile(r'^\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})\s*$')


# ----------------------------------------------------------------
# 電話番号の区切り（2026-09-11: AI が正規表現で区切りを当て推量し、名古屋の 052-133-4444 を 05-2133-4444 にした）。
# 市外局番の長さは番号だけでは決まらない（018 と 0182 のように前が同じで長さが違う）。道具が決めるのは、
#   ① 同じ列にハイフン付きの例がある市外局番（その区切りに合わせる）
#   ② 前が同じ長い市外局番の無いもの（03・06・011・045・052・075・078・092）と、携帯・IP・0120 などの特番
# だけ。決まらない番号はそのまま残す（当て推量で切らない）。
# ----------------------------------------------------------------
_PHONE_FIXED = {'03': 2, '06': 2, '011': 3, '045': 3, '052': 3, '075': 3, '078': 3, '092': 3}
_PHONE_SPECIAL = (('0120', (4, 3, 3)), ('0800', (4, 3, 3)), ('0570', (4, 3, 3)), ('0990', (4, 3, 3)))
_PHONE_MOBILE = ('090', '080', '070', '060', '050')
_PHONE_SEP_RE = re.compile(r'[\s\-‐‑‒–—―−ー－()（）]')
_PHONE_HYPHEN_RE = re.compile(r'(0\d{1,4})-(\d{1,4})-(\d{3,4})')


def _phone_split(d, parts):
    a, b, _c = parts
    return f"{d[:a]}-{d[a:a + b]}-{d[a + b:]}"


def _phone_plus81(t):
    """「+81-3-6666-5555」「+81 45 1111 2222」→ 区切りを保って 0 始まりに。当たらなければ None。"""
    m = re.fullmatch(r'\+81[-\s]?(\d{1,4})[-\s](\d{1,4})[-\s](\d{3,4})', t)
    return f"0{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None


def _phone_learn(col):
    """列の中のハイフン付きの番号から、市外局番 → (市外, 市内, 加入者) の桁数を覚える（純 Python）。"""
    learned = {}
    for v in col:
        if not isinstance(v, str):
            continue
        t = unicodedata.normalize('NFKC', v).strip()
        t = _phone_plus81(t) or t
        m = _PHONE_HYPHEN_RE.fullmatch(t)
        if m and len(m.group(1) + m.group(2) + m.group(3)) == 10 and not m.group(1).startswith(('0120', '0800', '0570')):
            learned[m.group(1)] = (len(m.group(1)), len(m.group(2)), len(m.group(3)))
    return learned


def _phone_format(v, learned=None):
    """1 つの電話番号を 03-1234-5678 の形に。区切りが決まらなければ元のまま（純 Python）。"""
    if not isinstance(v, str) or not v.strip():
        return v
    t = unicodedata.normalize('NFKC', v).strip()
    p81 = _phone_plus81(t)
    if p81:
        return p81
    d = _PHONE_SEP_RE.sub('', t)
    if d.startswith('+81'):
        d = '0' + d[3:]
    if not d.isdigit() or not d.startswith('0'):
        return v
    for pre, parts in _PHONE_SPECIAL:
        if d.startswith(pre) and len(d) == 10:
            return _phone_split(d, parts)
    if len(d) == 11 and d[:3] in _PHONE_MOBILE:
        return _phone_split(d, (3, 4, 4))
    if len(d) != 10:
        return v
    for area in sorted(learned or {}, key=len, reverse=True):
        parts = learned[area]
        if d.startswith(area) and sum(parts) == 10:
            return _phone_split(d, parts)
    for area, n in sorted(_PHONE_FIXED.items(), key=lambda x: -len(x[0])):
        if d.startswith(area) and d[n] != '0':                  # 次の桁が 0 なら長い市外局番（0920 壱岐 と 092 福岡 を分ける）
            return _phone_split(d, (n, 4 if n == 2 else 3, 4))
    return v


def _parse_en_date(s):
    """「March 15, 2026」「15 Mar 2026」→ datetime。読めなければ None（純 Python）。"""
    t = str(s or '')
    m = _EN_DATE_RE.match(t)
    if m:
        mon, d, y = m.group(1), m.group(2), m.group(3)
    else:
        m = _EN_DATE_RE2.match(t)
        if not m:
            return None
        d, mon, y = m.group(1), m.group(2), m.group(3)
    mo = _EN_MONTHS.get(mon.lower())
    if not mo:
        return None
    try:
        return datetime.datetime(int(y), mo, int(d), tzinfo=datetime.timezone.utc)
    except ValueError:
        return None


# ----------------------------------------------------------------
# 重複行を「同じ相手か」で見る（2026-09-11）。
# 動画のお題（BCG vs AI）を 51 行で撃ったら、材料の「重複行」は行の値が丸ごと同じ 3 組しか拾えず（本当は 26 組）、
# 残りを AI が自分で数えて、同じ会社の 2 行を両方消した。人は表を見て「空白や大文字が違うだけの同じ会社」と
# 判断している。その判断の材料（表記ゆれを無視した照合）を道具が持つ。
# ----------------------------------------------------------------
_LOOSE_DROP = set(_DASHES) | set('-,¥￥')
# 会社の種類の表記は相手を見分ける手がかりにならない（2026-09-11: 「株式会社アクアテック」と「アクアテック（株）」を
# 別の会社として残し、採点係の差し戻しで 1 往復増えた）。NFKC の後なので （株）・㈱ は (株) になっている
_CORP_RE = re.compile(r'株式会社|有限会社|合同会社|合資会社|合名会社|\((?:株|有|合|資|名)\)')
_CORP_ABBR_RE = re.compile(r'\((株|有|合|資|名)\)')
_CORP_FULL = {'株': '株式会社', '有': '有限会社', '合': '合同会社', '資': '合資会社', '名': '合名会社'}


def _loose_key(v):
    """照合用の値: 空白・全角半角・大文字小文字・ハイフン・カンマ・¥ を無視。日付は年月日。+81 は 0 に（純 Python）。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if _is_date_value(v):
        try:
            return f"{v.year}-{v.month}-{v.day}"
        except Exception:
            return str(v)
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if not isinstance(v, str):
        return str(v)
    s = v.strip(_WS_CHARS)
    if not s:
        return ''
    d = _parse_date_text(s)
    if d is not None:
        return f"{d.year}-{d.month}-{d.day}"
    s = unicodedata.normalize('NFKC', s).lower()
    # 会社の種類の略は正式の名前に戻して比べる（（株）アクアテック＝株式会社アクアテック）。前株・後株の位置は区別する
    # （2026-09-12・shu「重複判定は全列一致が基本。崩したら大変なことになる」＝株式会社A と A株式会社 を同じにしない）
    s = _CORP_ABBR_RE.sub(lambda m: _CORP_FULL[m.group(1)], s)
    # 負の数（-1,000・△1,000・(1,000)）は符号を残す（ハイフンを落とす前に見る。返金の行を支払いの行と同じと読まない）
    neg = bool(_LOOSE_NEG_RE.match(s))
    s = ''.join(ch for ch in s if not ch.isspace() and ch not in _LOOSE_DROP)
    if s.startswith('+81') and s[3:4].isdigit():
        s = '0' + s[3:]
    s = _loose_number_like(s)
    return ('-' + s) if neg and s[:1].isdigit() else s


_LOOSE_NEG_RE = re.compile(r'^\s*(?:[-−‐―△▲]\s*[\d,]+(?:\.\d+)?|\(\s*[\d,]+(?:\.\d+)?\s*\))\s*(?:円|個)?\s*$')
_LOOSE_NUM_RE = re.compile(r'^[△▲]?\(?(\d+(?:\.\d+)?)\)?(?:円|個|件|人|%)?$')


def _loose_number_like(s):
    """書き方だけ違う数・日付・郵便番号をそろえる（純 Python・_loose_key の続き。カンマ・¥・空白は落とした後）。
    2026-09-17: 重複を消した後の「重複でない行が消えた」検査が、〒240-7550／240-7550・20250526／2025-05-26・
    4422000円／4422000 を別の値と読み、マクロで合っている表を「行を消しすぎ」と言って AI を 3 往復させた。"""
    if s.startswith('〒'):
        s = s[1:]
    if len(s) == 8 and s.isdigit():
        y, mo, d = int(s[:4]), int(s[4:6]), int(s[6:])
        if 1900 <= y <= 2100 and 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y}-{mo}-{d}"
    m = _LOOSE_NUM_RE.match(s)
    if not m or s == m.group(1):
        return s
    num = m.group(1)
    if s.endswith('%'):
        try:
            f = float(num) / 100
            return str(int(f)) if f.is_integer() else repr(round(f, 10))
        except ValueError:
            return s
    if num.endswith('.0'):
        num = num[:-2]
    return num                     # 符号は呼び元（_loose_key）がハイフンを落とす前に見て付ける


def _identity_cols(body):
    """相手を見分ける列（値の種類が多い列）。状態・備考のように同じ値が並ぶ列は外す（純 Python）。"""
    width = max((len(r) for r in body), default=0)
    out = []
    for j in range(width):
        ks = [k for k in (_loose_key(r[j]) if j < len(r) else '' for r in body) if k]
        if len(ks) < 2:
            continue
        distinct = len(set(ks))
        if distinct >= 3 and distinct / len(ks) >= 0.3:
            out.append(j)
    return out


_ID_COL_HEAD_RE = re.compile(r'番号|コード|ｺｰﾄﾞ|^ID$|^No\.?$|№', re.I)
_CONTACT_HEAD_RE = re.compile(r'電話|TEL|FAX|携帯|郵便|〒', re.I)


def _heads_of(rows, header_idx):
    if header_idx is None or header_idx < 0 or header_idx >= len(rows):
        return None
    return rows[header_idx]


def _unique_cols(body, cols, heads=None):
    """番号の列（見出しが 番号・コード・ID・No。電話・郵便は除く）で、全部の値が違う列＝値が違えば別の相手。

    2026-09-11: 見出しを見ずに「全部の値が違う列」を番号扱いしていたので、重複を消した後の会社名の列が番号になり、
    「株式会社アクアテック」と「アクアテック（株）」（担当・電話・メールが同じ）を別の相手と判定していた。
    連番（1,2,3…＝行番号の写し）も除く（純 Python）。
    """
    out = []
    for j in cols:
        h = '' if not heads or j >= len(heads) or heads[j] is None else str(heads[j]).strip()
        if not h or not _ID_COL_HEAD_RE.search(h) or _CONTACT_HEAD_RE.search(h):
            continue
        ks = [k for k in (_loose_key(r[j]) if j < len(r) else '' for r in body) if k]
        if len(ks) >= 3 and len(set(ks)) == len(ks):
            try:
                nums = sorted(int(k) for k in ks)
                if nums == list(range(nums[0], nums[0] + len(nums))):
                    continue
            except ValueError:
                pass
            out.append(j)
    return out


_PLAIN_NUM_KEY_RE = re.compile(r'^\d+(?:\.\d+)?$')


def _weak_cols(body, cols):
    """数量・単価・金額のような「数だけの列」（先頭 0 の番号は除く）→ 列 index の集合（純 Python）。

    同じ数は別の取引でもよく並ぶ＝相手を見分ける手がかりとして弱い。2026-09-11 夜・試験の売上明細で、
    商品・数量・金額が同じなら日付も担当も違う取引を「同じ相手」と見て 4 行消した（見分け列の 6 割一致）。
    """
    out = set()
    for j in cols:
        ks = [k for k in (_loose_key(r[j]) if j < len(r) else '' for r in body) if k]
        num = sum(1 for k in ks if _PLAIN_NUM_KEY_RE.match(k) and not (len(k) > 1 and k.startswith('0')))
        if ks and num >= len(ks) * 0.8:
            out.add(j)
    return out


def _keys_same(a, b, cols, uniq, weak=()):
    """照合用の値の並び 2 つが「同じ相手」か。両方に値がある見分け列の 6 割以上（2 列以上）が同じ。番号の列が違えば別人。
    weak（数だけの列）の一致は数に入れても、数でない列（名前・日付・商品・電話…）が 2 列（両方に値のある列が
    1 列ならその 1 列）同じでなければ同じ相手にしない。"""
    both = same = sboth = ssame = 0
    for j in cols:
        x = a[j] if j < len(a) else ''
        y = b[j] if j < len(b) else ''
        if not x or not y:
            continue
        both += 1
        strong = j not in weak
        sboth += strong
        if x == y:
            same += 1
            ssame += strong
        elif j in uniq:
            return False
    return both >= 2 and same >= max(2, -(-both * 3 // 5)) and ssame >= min(2, sboth)


def _key_rows(rows, idx, width):
    return {i: [_loose_key(rows[i][j]) if j < len(rows[i]) else '' for j in range(width)] for i in idx}


def _body_idx(rows, header_idx):
    h = -1 if header_idx is None else header_idx
    return [i for i in range(h + 1, len(rows)) if any(_loose_key(v) for v in rows[i])]


def _dup_pairs(rows, header_idx=None):
    """見出しの下で、全部の列が同じ行（書き方の違いはならす）→ [(後の行 idx, 先に出る行 idx, 値が丸ごと同じか)]（純 Python）。

    2026-09-12・shu「重複判定は全列一致が基本。それ崩したら大変なことになる」。9/11 に足した「見分ける列の 6 割が
    同じなら同じ相手」は、備考だけ違う行まで重複として消させていた（例の表で 15 組）。消す判定は全列一致だけにし、
    1〜2 列だけ違う組は _near_dup_pairs で「重複の疑い」として報告に並べる（消さない・人が決める）。
    """
    body = _body_idx(rows, header_idx)
    if len(body) < 2:
        return []
    width = max(len(r) for r in rows)
    keys = _key_rows(rows, body, width)
    seen, out = {}, []
    for i in body:
        full = tuple(keys[i])
        if full in seen:
            k = seen[full]
            out.append((i, k, list(rows[i]) == list(rows[k])))
        else:
            seen[full] = i
    return out


def _near_dup_pairs(rows, header_idx=None):
    """全部の列は同じでないが、同じ相手に見える組（見分ける列の 6 割以上が同じ）→ [(後の行 idx, 先の行 idx, [違う列 idx])]。
    消さない。報告の「重複の疑い」に並べるだけ（純 Python）。"""
    full = {(i, k) for i, k, _e in _dup_pairs(rows, header_idx)}
    body = _body_idx(rows, header_idx)
    if len(body) < 2:
        return []
    width = max(len(r) for r in rows)
    keys = _key_rows(rows, body, width)
    return [(i, k, [j for j in range(width) if keys[i][j] != keys[k][j]])
            for i, k, _e in _fuzzy_pairs(rows, header_idx) if (i, k) not in full]


def _fuzzy_pairs(rows, header_idx=None):
    """見分ける列の 6 割以上が同じ組（9/11 の照合。いまは「重複の疑い」を数えるためだけに使う・消す判定には使わない）。"""
    body = _body_idx(rows, header_idx)
    if len(body) < 2:
        return []
    width = max(len(r) for r in rows)
    brows = [rows[i] for i in body]
    loose = len(body) * width <= 200000               # 大きい表は丸ごと同じ行だけ（照合の手間を抑える）
    cols = _identity_cols(brows) if loose else []
    uniq = set(_unique_cols(brows, cols, _heads_of(rows, header_idx)))
    weak = _weak_cols(brows, cols)
    keys = _key_rows(rows, body, width)
    exact_seen, idx, out = {}, {}, []
    for i in body:
        full = tuple(keys[i])
        if full in exact_seen:
            k = exact_seen[full]
            out.append((i, k, list(rows[i]) == list(rows[k])))
            continue
        found = None
        if cols:
            cands = sorted({k for j in cols if keys[i][j] for k in idx.get((j, keys[i][j]), ())})
            for k in cands:
                if _keys_same(keys[i], keys[k], cols, uniq, weak):
                    found = k
                    break
        if found is not None:
            out.append((i, found, False))
            continue
        exact_seen[full] = i
        for j in cols:
            if keys[i][j]:
                idx.setdefault((j, keys[i][j]), []).append(i)
    return out


def _lost_rows(before, after, header_before=None, header_after=None):
    """前の表の行のうち、後の表のどの行とも「同じ相手」にならない行 → [前の idx]（純 Python）。"""
    bbody = _body_idx(before, header_before)
    abody = _body_idx(after, header_after)
    if not bbody:
        return []
    width = max(max((len(r) for r in before), default=0), max((len(r) for r in after), default=0))
    brows = [before[i] for i in bbody]
    cols = _identity_cols(brows)
    uniq = set(_unique_cols(brows, cols, _heads_of(before, header_before)))
    weak = _weak_cols(brows, cols)
    kb = _key_rows(before, bbody, width)
    ka = _key_rows(after, abody, width)
    # 後の表で日付の列なのに、前の表では日付が数値のまま（シリアル値）だったセルは日付として照らす
    for j in range(width):
        avals = [after[i][j] for i in abody if j < len(after[i]) and after[i][j] not in (None, '')]
        if avals and sum(1 for v in avals if _is_date_value(v)) >= 0.6 * len(avals):
            for i in bbody:
                v = before[i][j] if j < len(before[i]) else None
                if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 80000 and float(v).is_integer():
                    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
                    kb[i][j] = f"{d.year}-{d.month}-{d.day}"
    full_after = {tuple(v) for v in ka.values()}
    idx = {}
    for i in abody:
        for j in cols:
            if ka[i][j]:
                idx.setdefault((j, ka[i][j]), []).append(i)
    lost = []
    for i in bbody:
        if tuple(kb[i]) in full_after:
            continue
        cands = {k for j in cols if kb[i][j] for k in idx.get((j, kb[i][j]), ())}
        if not any(_keys_same(kb[i], ka[k], cols, uniq, weak) for k in cands):
            lost.append(i)
    return lost


def _no_twin(rows, header_idx, drop):
    """消す行（idx の集合）のうち、残る行のどれとも同じ相手にならない行 → [idx]（純 Python）。"""
    body = _body_idx(rows, header_idx)
    if not body:
        return []
    width = max(len(r) for r in rows)
    brows = [rows[i] for i in body]
    cols = _identity_cols(brows)
    uniq = set(_unique_cols(brows, cols, _heads_of(rows, header_idx)))
    weak = _weak_cols(brows, cols)
    keys = _key_rows(rows, body, width)
    keep = [i for i in body if i not in drop]
    out = []
    for i in sorted(drop):
        if i not in keys:
            continue                                   # 空の行・見出しより上＝消しても相手は無くならない
        if any(keys[k] == keys[i] for k in keep):       # 相手＝全部の列が同じ行だけ（2026-09-12・全列一致が基本）
            continue
        out.append(i)
    return out


def _do_dedupe(ws, act):
    """dedupe の手（COM）。見出し行から末尾までの範囲で、keys の列が同じ行（表記ゆれは無視）の後の方を下から消す。"""
    rng = _check_addr(act.get('range'))
    if ':' not in rng:
        raise ValueError('dedupe の range は見出し行から表の末尾までです（例 "A5:I56"）')
    src = ws.Range(rng)
    r0, c0 = int(src.Row), int(src.Column)
    nr, nc = int(src.Rows.Count), int(src.Columns.Count)
    if nr < 3:
        raise ValueError("dedupe の range は見出し 1 行と本文 2 行以上です")
    if nr * nc > 200000:
        raise ValueError("dedupe は 20 万セルまでです")
    vals = [list(r) for r in _rows_of(src.Value)]
    heads = ['' if v is None else str(v).strip() for v in vals[0]]
    keys = act.get('keys') or act.get('key')
    if isinstance(keys, str):
        keys = [k for k in re.split(r'[,、\s]+', keys) if k]
    kidx = []
    for k in (keys or []):
        if isinstance(k, int) and not isinstance(k, bool) and 0 <= k < nc:
            kidx.append(k)
            continue
        k = str(k).strip()
        if k in heads:
            kidx.append(heads.index(k))
            continue
        if k.isascii() and k.isalpha():
            j = _col_num(k.upper()) - c0
            if 0 <= j < nc:
                kidx.append(j)
                continue
        raise ValueError(f"dedupe の keys「{k}」が、範囲の見出し（{' / '.join(h for h in heads if h)}）にも列文字にも当たりません")
    by_match = not kidx
    body = vals[1:]
    lk = [[_loose_key(r[j]) if j < len(r) else '' for j in range(nc)] for r in body]
    cols = _identity_cols(body)
    uniq = set(_unique_cols(body, cols, heads))
    weak = _weak_cols(body, cols)
    seen, drop = {}, []
    if by_match:
        # keys を省いた＝材料の「重複行」と同じ照合（表記ゆれ・担当・電話・メールなども見る）で消す（2026-09-11:
        # 会社名だけを keys にした回と、採点係が「アクアテック（株）も同じ会社」と差し戻した回とで結果が割れた）
        drop = [(i - 1, k - 1) for i, k, _e in _dup_pairs(vals, 0)]
        kidx = [0]
        kname = "材料の重複行と同じ照合（名前の書き方が違っても担当・電話・メールなどが同じなら同じ相手）"
    else:
        for i in range(len(body)):
            k = tuple(lk[i][j] for j in kidx)
            if not any(k):
                continue
            if k in seen:
                drop.append((i, seen[k]))
            else:
                seen[k] = i
        kname = '・'.join(heads[j] or _col_letter(c0 + j) for j in kidx)
    # 全列一致の門（2026-09-12・shu「重複判定は全列一致が基本」）: keys が同じでも、他の列が 1 つでも違う行は消さない
    # （備考だけ違う行を消すと備考が失われる。どちらを残すかは人の判断＝report に書く）
    bad = [] if by_match else [(i, k) for i, k in drop if lk[i] != lk[k]]
    if bad:
        ex = "／".join(f"行{r0 + 1 + i} と 行{r0 + 1 + k}" for i, k in bad[:5])
        raise ValueError(f"keys（{kname}）が同じでも、他の列が違う行があります（{ex}）。重複は全部の列が同じ行だけです"
                         "（書き方の違いは道具がならします）。1 行も消していません。keys を省けば全部の列が同じ行だけを消します。"
                         "備考などだけ違う行は消さずに report の【人に判断してほしいこと】に書いてください")
    if not drop:
        return (f"dedupe {rng}", True, f"重複はありません（keys: {kname}）。何も消していません\n")
    ur = ws.UsedRange
    uc0 = int(ur.Column)
    uc1 = uc0 + int(ur.Columns.Count) - 1
    kept = []
    for i, _k in sorted(drop, reverse=True):          # 下の行から＝上の行番号がずれない
        r = r0 + 1 + i
        outside = False
        if uc0 < c0 or uc1 > c0 + nc - 1:
            got = _rows_of(ws.Range(ws.Cells(r, uc0), ws.Cells(r, uc1)).Value)
            for jj, v in enumerate(got[0] if got else ()):
                col = uc0 + jj
                if (col < c0 or col > c0 + nc - 1) and v not in (None, ''):
                    outside = True
                    break
        if outside:
            # 表の外の同じ行に値がある。表の中だけ上へ詰めると隣の列が行とずれ、行ごと消すとその値を巻き込む
            # （2026-09-21 利用者の報告）＝消さずに知らせる
            kept.append((i, _k))
        else:
            ws.Rows(r).Delete()
    drop = [p for p in drop if p not in kept]
    if not drop:
        return (f"dedupe {rng}", True, "重複はありますが、どれも表の外の同じ行に値があるため消していません"
                f"（{'／'.join(f'行{r0 + 1 + i}' for i, _k in sorted(kept))}）\n")
    last = r0 + nr - 1 - len(drop)
    end = f"{_col_letter(c0)}{r0}:{_col_letter(c0 + nc - 1)}{last}"
    j0 = kidx[0]
    lines = [f"行{r0 + 1 + i} = 行{r0 + 1 + k}（{_text_of(body[i][j0]) if j0 < len(body[i]) else ''}）" for i, k in drop]
    out = (f"重複 {len(drop)} 行を下の行から消しました（keys: {kname}。先に出てくる行を残す。"
           "空白・全角半角・大文字小文字・ハイフンの違いは無視して照合）\n"
           + "消した行（消す前の行番号）: " + "／".join(lines) + "\n"
           + f"表はいま {end}（見出し + {nr - 1 - len(drop)} 行）。後ろの手はこの番地で"
             "（消した分だけ下の行が繰り上がっています）\n")
    if kept:
        out += ("重複でも消していない行（表の外の同じ行に値がある＝消すと巻き込む。消す前の行番号）: "
                + "／".join(f"行{r0 + 1 + i}" for i, _k in sorted(kept)) + "\n")
    return (f"dedupe {rng}", True, out)


__all__ = ['_ADDR_RE', '_CHART_TYPES', '_CSV_ENCODINGS', '_DASHES', '_DATELIKE_RE', '_DATE_TEXT_RE', '_FILL_MAX_CELLS', '_HANKANA_ANY_RE', '_HANKANA_RE', '_HEADER_SCAN_ROWS', '_IMAGE_EXTS', '_INNER_WS_RE', '_NORMALIZE_RULES', '_NUMLIKE_RE', '_NUM_TEXT_RE', '_PATH_BARE_RE', '_PATH_QUOTED_RE', '_PEEK_MAX_FILES', '_PEEK_ROWS', '_PEEK_ROWS_ALL', '_PROFILE_MAX_COLS', '_PROFILE_MAX_ROWS', '_PROFILE_MIN_ROWS', '_READ_FILE_EXTS', '_READ_FILE_MAX_COLS', '_READ_FILE_MAX_ROWS', '_WAREKI_BASE', '_WAREKI_RE', '_WS_CHARS', '_XL_ERRORS', '_XL_ERROR_CODES', '_XL_MAX_COL', '_XL_MAX_ROW', '_ZENKAKU_DIGIT_RE', '_ZENKAKU_RE', '_ZEN_ALNUM_RE', '_addr', '_cell_write_back', '_charkind_facts', '_check_addr', '_col_num', '_column_profile', '_csv_check_path', '_csv_text', '_date_fmt', '_do_fill', '_do_normalize', '_do_dedupe', '_CORP_RE', '_phone_format', '_phone_learn', '_phone_split', '_phone_plus81', '_PHONE_FIXED', '_PHONE_SPECIAL', '_PHONE_MOBILE', '_PHONE_SEP_RE', '_PHONE_HYPHEN_RE', '_heads_of', '_ID_COL_HEAD_RE', '_CONTACT_HEAD_RE', '_dup_pairs', '_lost_rows', '_no_twin', '_loose_key', '_identity_cols', '_unique_cols', '_keys_same', '_parse_en_date', '_EN_MONTHS', '_EN_DATE_RE', '_EN_DATE_RE2', '_LOOSE_DROP', '_key_rows', '_body_idx', '_eval_text', '_extra_materials', '_guess_header_row', '_hankaku', '_heavy_materials', '_id_columns', '_is_date_value', '_is_xl_error', '_new_sheet_name', '_normalize_column', '_normalize_spec', '_normalize_value', '_open_book_by_path', '_page_count', '_parse_date_text', '_parse_wareki_text', '_paths_in_request', '_peek_file', '_peek_files_note', '_pick_header_row', '_print_summary', '_profile_columns', '_read_book_rows', '_read_csv_rows', '_read_file_path', '_read_file_write', '_read_text_decode', '_rows_of', '_shape_of', '_table_blanks', '_text_of', '_zenkaku_digits']
