# -*- coding: utf-8 -*-
"""key_gui.py — APIキーの設定・削除の窓（2026-09-04）

黒い窓に見えない貼り付けをさせない。小さな窓を1枚出して、
「貼る → 押す → 通りましたと出る」で終わらせる。

  py key_gui.py                 窓を出す
  py key_gui.py --selftest PNG  窓を組んで見た目を PNG に撮って閉じる（渡す前の目視用）

キーは金庫（Windows のログインで暗号化＝DPAPI・期限つき）に預かる。いまのプロセスの環境変数にも入れるが、
窓を閉じれば消える。設定と削除の中身は vbam_agent が持っている。
"""
import os
import sys
import threading
import webbrowser

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk

from vbam_keys import (_KEY_ENV, _KEY_ISSUE_URL, _mask_key, _key_problem,
                        _user_env_get, _user_env_write, _ping_key,
                        _key_save, _key_forget, _key_info, _KEY_DAYS_DEFAULT)
from vbam_build import _AI_DEFAULT_MODEL

FONT = ("Meiryo UI", 11)
FONT_B = ("Meiryo UI", 11, "bold")
FONT_BIG = ("Meiryo UI", 14, "bold")
OK_COLOR = "#1B7F3B"
NG_COLOR = "#B00020"


class KeyWindow:
    def __init__(self, root):
        self.root = root
        root.title("APIキーの設定")
        root.resizable(False, False)
        pad = {"padx": 14, "pady": 6}

        tk.Label(root, text="APIキーの設定", font=FONT_BIG, anchor="w").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(12, 2))

        # 何のためのものか（事務の人が読んで分かる言葉で）
        tk.Label(root, font=FONT, anchor="w", justify="left", fg="#333", wraplength=560,
                 text=("「キー」は、この道具が AI に話しかけるための合言葉です。\n"
                       "AI の会社の画面で、自分の名前で無料で作れます（下の「キーを取りに行く」）。")
                 ).grid(row=1, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 6))
        self.env_warn = tk.Label(root, font=FONT_B, anchor="w", justify="left",
                                 fg=NG_COLOR, wraplength=560, text="")

        # どのAIか
        self.ai = tk.StringVar(value="gemini")
        box = tk.Frame(root)
        box.grid(row=2, column=0, columnspan=3, sticky="w", padx=14)
        tk.Label(box, text="使う AI:", font=FONT).pack(side="left")
        for name, label in (("gemini", "Gemini"), ("claude", "Claude")):
            tk.Radiobutton(box, text=label, value=name, variable=self.ai, font=FONT,
                           command=self.refresh).pack(side="left", padx=(8, 0))

        # いまの状態
        self.now = tk.Label(root, text="", font=FONT_B, anchor="w", fg="#333")
        self.now.grid(row=3, column=0, columnspan=3, sticky="w", **pad)

        # 貼る場所（見えるまま＝貼れたことが分かる）
        tk.Label(root, text="キーを貼り付けてください（Ctrl+V か 右クリック）", font=FONT_B,
                 anchor="w").grid(row=5, column=0, columnspan=3, sticky="w", padx=14)
        self.entry = tk.Entry(root, font=("Consolas", 11), width=52)
        self.entry.grid(row=6, column=0, columnspan=2, sticky="we", padx=(14, 4), pady=4)
        self.entry.bind("<Return>", lambda _e: self.on_set())
        tk.Button(root, text="貼り付け", font=FONT, width=9, command=self.on_paste).grid(
            row=6, column=2, sticky="w", padx=(0, 14))

        # 保管の期間（置きっぱなしにしない＝人の記憶に頼らない）
        exp = tk.Frame(root)
        exp.grid(row=7, column=0, columnspan=3, sticky="w", padx=14, pady=(8, 0))
        tk.Label(exp, text="預かる期間:", font=FONT).pack(side="left")
        self.days = tk.StringVar(value=str(_KEY_DAYS_DEFAULT))
        for label, val in (("30日", "30"), ("90日", "90"), ("期限なし", "0")):
            tk.Radiobutton(exp, text=label, value=val, variable=self.days, font=FONT).pack(
                side="left", padx=(8, 0))
        tk.Label(exp, text="（期間が過ぎたら、この道具が自分で消します）", font=FONT, fg="#555").pack(
            side="left", padx=(8, 0))

        # どこに保管されるのか（人に渡すときに何をすればいいかまで書く）
        tk.Label(root, font=FONT, anchor="w", justify="left", fg="#555", wraplength=560,
                 text=("保管の仕方: キーは暗号にしてから、このパソコンのあなた専用の場所にしまいます。\n"
                       "　・そのファイルを持ち出しても、他の人・他のパソコンでは開けません\n"
                       "　・画面にも記録にも、キーの全部は出しません\n"
                       "　・このパソコンを人に渡すときは「消す」を押してください\n"
                       "（詳しい人向け: Windows の DPAPI で暗号化し、LOCALAPPDATA に置いています）")
                 ).grid(row=8, column=0, columnspan=3, sticky="w", padx=14, pady=(6, 2))

        # ボタン
        bar = tk.Frame(root)
        bar.grid(row=9, column=0, columnspan=3, sticky="we", padx=14, pady=(10, 4))
        self.btn_set = tk.Button(bar, text="設定して確認", font=FONT_B, width=14, command=self.on_set)
        self.btn_set.pack(side="left")
        tk.Button(bar, text="キーを取りに行く", font=FONT, width=16, command=self.on_issue).pack(
            side="left", padx=8)
        tk.Button(bar, text="消す", font=FONT, width=8, command=self.on_clear).pack(side="left")
        tk.Button(bar, text="閉じる", font=FONT, width=8, command=root.destroy).pack(side="right")

        # 結果
        self.msg = tk.Label(root, text="", font=FONT_B, anchor="w", justify="left",
                            wraplength=560, fg="#444")
        self.msg.grid(row=10, column=0, columnspan=3, sticky="w", padx=14, pady=(2, 12))

        self.refresh()
        self.entry.focus_set()

    # ---- 表示 ----
    def refresh(self):
        ai = self.ai.get()
        env = _KEY_ENV[ai]
        have, info = _key_info(ai)
        self.now.config(text=("いまの状態: " + info) if info else "いまの状態: まだ入っていません")
        # 昔のやり方（環境変数に平文）で入っている人には、その場で片づけてもらう
        if _user_env_get(env):
            self.env_warn.config(
                text=(f"⚠ 古いやり方で {env} に平文のまま入っています（期限なし・他のソフトからも読めます）。\n"
                      "　　下の「消す」を押すと、そちらも一緒に片づけます。"))
            self.env_warn.grid(row=4, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 4))
        else:
            self.env_warn.grid_remove()

    def say(self, text, ok=None):
        self.msg.config(text=text, fg=(OK_COLOR if ok else NG_COLOR) if ok is not None else "#444")
        self.root.update_idletasks()

    # ---- 手 ----
    def on_paste(self):
        try:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, self.root.clipboard_get().strip())
            self.say("貼り付けました。「設定して確認」を押してください。")
        except Exception:
            self.say("クリップボードが空です。キーをコピーしてから押してください。", ok=False)

    def on_issue(self):
        webbrowser.open(_KEY_ISSUE_URL[self.ai.get()])
        self.say("ブラウザを開きました。キーを作ってコピーし、上の欄に貼り付けてください。")

    def on_set(self):
        ai = self.ai.get()
        env = _KEY_ENV[ai]
        key = self.entry.get().strip()
        bad = _key_problem(key)
        if bad:
            self.say(f"{bad}\nもう一度、キーだけを貼り付けてください。", ok=False)
            return
        os.environ[env] = key           # いまのプロセスにだけ（窓を閉じれば消える）
        self.say("疎通を確かめています…")
        self.btn_set.config(state="disabled")

        def work():
            # 疎通が通ってから金庫に入れる（前は先に預かっていたので、壊れたキーが残っていた・2026-09-04）
            ok, line = _ping_key(ai, _AI_DEFAULT_MODEL[ai], key)
            exp = err = None
            if ok:
                try:
                    exp = _key_save(ai, key, int(self.days.get()))
                except Exception as ex:
                    err = str(ex)
            def done():
                self.btn_set.config(state="normal")
                if not ok:
                    self.say(f"{line}\n貼り付けが途中で切れていないか確かめて、もう一度お試しください。"
                             "（預かっていません）", ok=False)
                    return
                if err:
                    self.say(f"保管できませんでした: {err}", ok=False)
                    return
                self.refresh()
                self.entry.delete(0, tk.END)
                self.say(f"{line}\n預かりました（" + (f"{exp} まで" if exp else "期限なし")
                         + "）。次からは、この設定は要りません。この窓は閉じてかまいません。", ok=True)
            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def on_clear(self):
        ai = self.ai.get()
        env = _KEY_ENV[ai]
        have, _ = _key_info(ai)
        in_env = bool(_user_env_get(env))
        if not have and not in_env:
            self.say("消すものはありません。")
            return
        from tkinter import messagebox
        if not messagebox.askyesno("APIキーを消す",
                                   f"{ai} のキーを、このパソコンから消します。よろしいですか。"):
            return
        done = []
        if _key_forget(ai):
            done.append("預かっていたぶん")
        if in_env:
            try:
                _user_env_write(env, None)
            except Exception as ex:
                self.say(f"消せませんでした: {ex}", ok=False)
                return
            if _user_env_get(env):
                self.say("消したつもりが残っています。もう一度お試しください。", ok=False)
                return
            done.append(f"古いやり方のぶん（{env}）")
        os.environ.pop(env, None)
        self.refresh()
        self.say("消しました（" + "・".join(done) + "）。\n"
                 "※ AI の会社に登録されたキー自体は、まだ生きています。使えなくするには、"
                 "「キーを取りに行く」の画面でそのキーを削除してください。", ok=True)


def _selftest(png_path):
    """窓を組んで見た目を PNG に撮る（渡す前の目視用。API は呼ばない）。"""
    root = tk.Tk()
    KeyWindow(root)
    root.geometry("+40+40")               # 主画面の決まった位置に出す（別の画面を撮る事故の防止）
    root.lift()
    root.attributes("-topmost", True)     # 手前に出してから撮る
    root.update()
    root.update_idletasks()
    try:
        from PIL import ImageGrab
        # この処理は DPI 非対応なので winfo_* は拡大前の値。画面の実寸と見比べて倍率を出す
        shot = ImageGrab.grab(all_screens=False)
        scale = shot.width / float(root.winfo_screenwidth())
        x, y = root.winfo_rootx(), root.winfo_rooty()
        w, h = root.winfo_width(), root.winfo_height()
        rect = tuple(int(v * scale) for v in (x, y, x + w, y + h))
        shot.crop(rect).save(png_path)
        print(f"撮りました: {png_path}（{rect[2] - rect[0]}x{rect[3] - rect[1]}）")
    except Exception as ex:
        print(f"撮れませんでした: {ex}")
    root.destroy()


def main():
    if '--selftest' in sys.argv:
        i = sys.argv.index('--selftest')
        _selftest(sys.argv[i + 1] if len(sys.argv) > i + 1 else '_key_gui.png')
        return
    root = tk.Tk()
    KeyWindow(root)
    root.mainloop()


if __name__ == '__main__':
    main()
