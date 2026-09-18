# -*- coding: utf-8 -*-
"""recipe_gui.py — 手順書を選んで撃つ窓（2026-09-07）

20 本の手順書は、これまでコマンドの引数（agent --recipe 名前）でしか選べなかった。
名前を覚えていないと選べない＝人が使う入口が無い状態だった。一覧から選んで、
補足（承認の言葉・番地・仕様）を書いて、押すだけにする。

  py recipe_gui.py              窓を出す
  py recipe_gui.py --selftest PNG   窓を組んで見た目を PNG に撮って閉じる（渡す前の目視用）

対象は「いま開いているブックのアクティブシート」。保存はしない（気に入らなければ保存せず閉じる）。
"""
import os
import subprocess
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk

from vbam_recipes import SHEET_RECIPES, recipe_when

FONT = ("Meiryo UI", 11)
FONT_B = ("Meiryo UI", 11, "bold")
FONT_BIG = ("Meiryo UI", 14, "bold")
FONT_LIST = ("Meiryo UI", 14)
OK_COLOR = "#1B7F3B"
NG_COLOR = "#B00020"

# 承認の言葉（補足に入れると、人の値を書き換える手が通る）。手順書の「承認:」の段と同じ語
APPROVAL_WORDS = ("置き換えてよい", "消してよい", "埋めてよい", "上書きしてよい")


class RecipeWindow:
    def __init__(self, root):
        self.root = root
        self.proc = None
        root.title("手順書を選ぶ")
        root.geometry("1180x700")
        pad = {"padx": 14, "pady": 6}

        tk.Label(root, text="手順書を選んでください", font=FONT_BIG).grid(
            row=0, column=0, columnspan=2, sticky="w", **pad)

        left = tk.Frame(root)
        left.grid(row=1, column=0, sticky="nsew", **pad)
        self.listbox = tk.Listbox(left, font=FONT_LIST, height=14, width=26,
                                  exportselection=False, activestyle="none")
        self.listbox.pack(side="left", fill="both", expand=True)
        bar = tk.Scrollbar(left, command=self.listbox.yview)
        bar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=bar.set)
        for name in SHEET_RECIPES:
            self.listbox.insert("end", name)
        self.listbox.bind("<<ListboxSelect>>", self.on_pick)

        right = tk.Frame(root)
        right.grid(row=1, column=1, sticky="nsew", **pad)
        tk.Label(right, text="いつ使うか", font=FONT_B).pack(anchor="w")
        self.when = tk.Message(right, text="", font=FONT, width=640, justify="left")
        self.when.pack(anchor="w", pady=(0, 10))
        tk.Label(right, text="手順（そのまま AI に渡します）", font=FONT_B).pack(anchor="w")
        self.body = tk.Text(right, font=("Meiryo UI", 10), width=62, height=13, wrap="word")
        self.body.pack(fill="both", expand=True)
        self.body.config(state="disabled")

        tk.Label(root, text="補足（承認の言葉・番地・仕様。空でも可）", font=FONT_B).grid(
            row=2, column=0, columnspan=2, sticky="w", **pad)
        sup = tk.Frame(root)
        sup.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.supplement = tk.Entry(sup, font=FONT)
        self.supplement.pack(fill="x", expand=True)
        words = tk.Frame(root)                       # 承認の言葉は入力欄の下に並べる（右端で切れないように）
        words.grid(row=4, column=0, columnspan=2, sticky="w", padx=14)
        tk.Label(words, text="よく使う承認の言葉:", font=FONT).pack(side="left")
        for word in APPROVAL_WORDS:
            tk.Button(words, text=word, font=FONT,
                      command=lambda w=word: self.add_word(w)).pack(side="left", padx=4)

        row = tk.Frame(root)
        row.grid(row=5, column=0, columnspan=2, sticky="ew", **pad)
        self.run_btn = tk.Button(row, text="この手順書で撃つ", font=FONT_B, command=self.run)
        self.run_btn.pack(side="left")
        tk.Button(row, text="閉じる", font=FONT, command=root.destroy).pack(side="right")
        self.status = tk.Label(row, text="対象はいま開いているブックのアクティブシート。保存はしません。",
                               font=FONT)
        self.status.pack(side="left", padx=14)

        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(1, weight=1)
        self.listbox.selection_set(0)
        self.on_pick()

    # -------- 画面 --------
    def current(self):
        sel = self.listbox.curselection()
        return self.listbox.get(sel[0]) if sel else None

    def on_pick(self, _event=None):
        name = self.current()
        if not name:
            return
        self.when.config(text=recipe_when(name))
        self.body.config(state="normal")
        self.body.delete("1.0", "end")
        self.body.insert("1.0", SHEET_RECIPES[name])
        self.body.config(state="disabled")

    def add_word(self, word):
        cur = self.supplement.get().strip()
        if word in cur:
            return
        self.supplement.delete(0, "end")
        self.supplement.insert(0, (cur + " " + word).strip())

    # -------- 実行 --------
    def run(self):
        name = self.current()
        if not name:
            return
        self.run_btn.config(state="disabled")
        self.status.config(text=f"「{name}」を撃っています…（窓はそのままで結果を待ってください）",
                           fg="black")
        threading.Thread(target=self._run_bg, args=(name, self.supplement.get().strip()),
                         daemon=True).start()

    def _run_bg(self, name, supplement):
        here = os.path.dirname(os.path.abspath(__file__))
        cmd = [sys.executable, os.path.join(here, "vba_manager.py"), "agent", "--recipe", name]
        if supplement:
            cmd.append(supplement)
        try:
            p = subprocess.run(cmd, cwd=here, capture_output=True, encoding="utf-8",
                               errors="replace")
            out = (p.stdout or "") + (p.stderr or "")
            ok = p.returncode == 0
        except Exception as ex:                      # 起動できない（py が無い等）も窓に出す
            out, ok = str(ex), False
        self.root.after(0, self._done, name, ok, out)

    def _done(self, name, ok, out):
        self.run_btn.config(state="normal")
        line = next((l for l in reversed(out.splitlines()) if l.strip()), "")
        self.status.config(text=("終わりました: " if ok else "止まりました: ") + line[:70],
                           fg=OK_COLOR if ok else NG_COLOR)
        self.body.config(state="normal")
        self.body.delete("1.0", "end")
        self.body.insert("1.0", out[-8000:] if out else "（出力がありません）")
        self.body.see("end")
        self.body.config(state="disabled")


def _dpi_aware():
    """画面の拡大率（125% など）に合わせる。しないと窓がぼやけ、撮った PNG も右下が欠ける。"""
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main(argv=None):
    _dpi_aware()
    argv = list(argv if argv is not None else sys.argv[1:])
    shot = None
    if argv and argv[0] == "--selftest":
        shot = argv[1] if len(argv) >= 2 else "_recipe_gui.png"
    root = tk.Tk()
    RecipeWindow(root)
    if shot:
        root.update_idletasks()
        root.update()
        root.update_idletasks()
        try:
            from PIL import ImageGrab
            root.update_idletasks()
            x, y = root.winfo_rootx(), root.winfo_rooty()
            box = (x, y, x + root.winfo_width(), y + root.winfo_height())
            ImageGrab.grab(box).save(shot)
            print(f"撮りました: {shot}")
        except Exception as ex:
            print(f"撮れませんでした: {ex}")
        root.destroy()
        return 0
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
