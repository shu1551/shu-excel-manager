# -*- coding: utf-8 -*-
"""お題のフォルダ 1 つを鍛える（2026-09-18）。request.txt・phrases.txt・*_本番_前/正解.xlsx・*_試験*_前/正解.xlsx を読んで
vbam_forge.forge に渡す（agent --forge --before --truth --test --phrases と同じ・日本語をコマンド行に書かない）。
  py -u forge_job.py <フォルダ> <仕事の名前> [--turns N] [--also 別のフォルダ …] [--continue]
  --also     … そのフォルダの *_前/正解 の組も別の表に足す（未見で外れた表など）
  --continue … 弾を作り直さず、台帳の弾に表を足して前回のマクロから続ける
会話している AI（Gemini・Claude など）が自分でマクロを書くとき（2026-09-19・API の鍵は要らない）:
  --prompt          … AI を呼ばずに、書き手への問い（決まり・依頼・表・前回の外れ）を <フォルダ>\\_forge_prompt.txt に書いて止まる
  --answer [ファイル] … 答え（Sub 全文・既定 <フォルダ>\\_forge_answer.txt・UTF-8 可）を 1 往復ぶん採点する。
                      不合格なら _forge_prompt.txt が次の問い（外れの説明つき）に書き換わる＝読んで答えを直し、また --answer
"""
import glob
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'python_scripts')))
import vbam_core
vbam_core.setup_encoding()
import vbam_forge as vf


def pairs(folder, pattern):
    out = []
    for b in sorted(glob.glob(os.path.join(folder, pattern))):
        t = b[:-len('_前.xlsx')] + '_正解.xlsx'
        if os.path.isfile(t):
            out += [b, t]
    return out


def _lock():
    """台帳（_agent_forge.json）を書く鍛えは 1 本ずつ（2026-09-18: 2 本同時に回し、後から終わった方が古い台帳を書き戻して
    合格した仕事を「不合格」に戻した＝2 回やった）。ほかの鍛えが走っていれば終わるまで待つ。"""
    import msvcrt
    import time
    path = vf._AGENT_FORGE_FILE + '.lock'
    f = open(path, 'a+')
    said = False
    while True:
        try:
            f.seek(0)
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            return f
        except OSError:
            if not said:
                print('ほかの鍛えが走っています。終わるまで待ちます（台帳を 2 本で書かないため）', flush=True)
                said = True
            time.sleep(5)


def main():
    lock = _lock()   # noqa: F841  プロセスが終わるまで持つ
    args = sys.argv[1:]
    folder, name = os.path.abspath(args[0]), args[1]
    turns = int(args[args.index('--turns') + 1]) if '--turns' in args else 6
    also = []
    if '--also' in args:
        for a in args[args.index('--also') + 1:]:
            if a.startswith('--'):
                break
            also.append(os.path.abspath(a))
    prompt_path = os.path.join(folder, '_forge_prompt.txt')
    answer = None
    if '--answer' in args:
        i = args.index('--answer') + 1
        answer = (os.path.abspath(args[i]) if i < len(args) and not args[i].startswith('--')
                  else os.path.join(folder, '_forge_answer.txt'))
        if name not in vf._forge_load():
            print('エラー: 台帳にこの仕事がありません。先に --prompt で問いを書き出してください')
            return
        if not os.path.isfile(answer):
            print(f'エラー: 答えのファイルがありません: {answer}')
            return
    extra = {}
    if answer:
        extra = {'ai': 'file', 'model': answer, 'prompt_out': prompt_path}
    elif '--prompt' in args:
        extra = {'prompt_only': True, 'prompt_out': prompt_path}
    req = open(os.path.join(folder, 'request.txt'), encoding='utf-8').read().strip()
    phrases = os.path.join(folder, 'phrases.txt')
    tests = pairs(folder, '*_試験*_前.xlsx')
    for a in also:
        tests += pairs(a, '*_前.xlsx')
    main_pair = pairs(folder, '*_本番_前.xlsx')
    if '--retest' in args:
        # 別の表の組を作り直す（覚書＝選んでいる列を読み直すときなど）
        d = vf._forge_load()
        if name in d:
            d[name]['tests'] = []
            d[name]['tests_passed'] = False
            d[name]['passed'] = False
            vf._forge_save(d)
            print('別の表の組を作り直します')
    if '--refresh' in args:
        # お題を作り直したとき: 弾（前・正解・別の表）を読み直し、前のマクロはそのまま残して続ける
        old = dict(vf._forge_load().get(name) or {})
        if not vf.harvest_pair(name, main_pair[0], main_pair[1], None, req):
            return
        d = vf._forge_load()
        keep = {k: old[k] for k in ('code', 'sub', 'bas', 'ask', 'handles', 'headers', 'keep_ask', 'ref_code')
                if old.get(k) is not None}
        d[name].update(keep)
        d[name]['passed'] = False
        vf._forge_save(d)
        print(f"弾を作り直しました（前のマクロは残します: {keep.get('sub')}）")
    if '--continue' in args or '--refresh' in args or answer:
        d = vf._forge_load()
        if name in d and d[name].get('request') != req:
            d[name]['request'] = req                           # 依頼文を書き足したら弾にも入れる（前回のマクロから続ける）
            vf._forge_save(d)
            print('依頼文を書き直しました')
        ok = vf.forge(name, tests=tests, phrases=phrases if os.path.isfile(phrases) else None,
                      max_turns=1 if answer else turns, **extra)
    else:
        ok = vf.forge(name, before=main_pair[0], truth=main_pair[1], tests=tests, request=req,
                      phrases=phrases if os.path.isfile(phrases) else None, max_turns=turns, **extra)
    if ok is None:
        print('FORGE_PROMPT', prompt_path)
        return
    if ok and extra:
        vf._write_prompt(prompt_path, '（合格しました。次の問いはありません）\n')
    print('FORGE_RESULT', name, 'passed' if ok else 'failed')


if __name__ == '__main__':
    main()
