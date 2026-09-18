# -*- coding: utf-8 -*-
"""desktop_link.py — デスクトップのショートカットを置く／外す（2026-09-04）

**必ず Explorer に知らせる。** 知らせずに消すと、画面にアイコンだけ残って
「押しても消せない幽霊」になる（2026-09-04 に実際にやらかした）。
置くときも外すときも、この道具を通す。直接 os.remove しない。

  py desktop_link.py place  名前  実行するもの [引数]   ショートカットを置く
  py desktop_link.py remove 名前                        外す（実体を消して Explorer に知らせる）
  py desktop_link.py list                               置いてあるものの一覧
"""
import ctypes
import os
import sys

SHCNE_ASSOCCHANGED = 0x08000000


def desktop_dir():
    import win32com.client
    return win32com.client.Dispatch("WScript.Shell").SpecialFolders("Desktop")


def _notify():
    """Explorer に「変わった」と知らせる。これを忘れると幽霊が残る。"""
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, 0x0000, None, None)


def place(name, target, args='', workdir=None, desc='', icon=None):
    import win32com.client
    path = os.path.join(desktop_dir(), name if name.endswith('.lnk') else name + '.lnk')
    lnk = win32com.client.Dispatch("WScript.Shell").CreateShortcut(path)
    lnk.TargetPath = target
    lnk.Arguments = args
    lnk.WorkingDirectory = workdir or os.path.dirname(target)
    lnk.Description = desc
    if icon:
        lnk.IconLocation = icon
    lnk.Save()
    _notify()
    return path


def remove(name):
    path = os.path.join(desktop_dir(), name if name.endswith('.lnk') else name + '.lnk')
    existed = os.path.exists(path)
    if existed:
        os.remove(path)
    _notify()                       # 消しても消さなくても知らせる（画面と実物を必ず合わせる）
    return existed


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in ('place', 'remove', 'list'):
        print(__doc__.strip())
        return 1
    cmd = sys.argv[1]
    if cmd == 'list':
        d = desktop_dir()
        for n in sorted(x for x in os.listdir(d) if x.lower().endswith('.lnk')):
            print(n)
        return 0
    if cmd == 'remove':
        ok = remove(sys.argv[2])
        print(("外しました: " if ok else "もともとありません: ") + sys.argv[2]
              + "（Explorer に知らせ済み）")
        return 0
    name, target = sys.argv[2], sys.argv[3]
    args = ' '.join(f'"{a}"' for a in sys.argv[4:])
    print("置きました: " + place(name, target, args) + "（Explorer に知らせ済み）")
    return 0


if __name__ == '__main__':
    sys.exit(main())
