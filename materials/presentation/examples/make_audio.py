#!/usr/bin/env python3
"""作例の音声を作る。narration/<題材>.json の原稿を edge-tts で読み上げ、audio/<題材>/ に書き出す。
   ・talk（発表の原稿）    → audio/<題材>/talk-<キー>.mp3      発表する学生の声（ja-JP-KeitaNeural）
   ・comment（教員の解説） → audio/<題材>/comment-<キー>.mp3   各回ワークシートの解説と同じ声（ja-JP-NanamiNeural）
   長さ（秒）と原稿の指紋を audio/<題材>/audio.json に書く。原稿・声・読み替えが前回から変わったものだけ作り直す。
   原稿を直したら、これを実行してから node build.js --examples で作例を組み直す（発表の音声が目安の秒数を超えると、build.js が知らせる）。
   キーは examples/<題材>.js の content と同じ。型によって原稿を変えるときは「型:キー」（例 story:cover）で、ファイル名では story-cover になる。
   usage: python3 make_audio.py            （全題材）
          python3 make_audio.py senior     （題材を指定）
          python3 make_audio.py --force    （すべて作り直す）
   必要なもの: pip install edge-tts mutagen（Microsoft Edge の読み上げ音声を使うので、ネットワークが必要）"""
import asyncio
import glob
import hashlib
import json
import os
import re
import sys

import edge_tts
from mutagen.mp3 import MP3

HERE = os.path.dirname(os.path.abspath(__file__))
VOICE = {'talk': 'ja-JP-KeitaNeural', 'comment': 'ja-JP-NanamiNeural'}
RATE = '+0%'

# 読み上げのときだけ使う読み替え（原稿の表記はそのまま。ノートには原稿のほうを載せる）。上から順に当てはめる
READ = [
    (r'高橋陽菜', '高橋ひな'), (r'陽菜', 'ひな'), (r'山本和夫', '山本かずお'), (r'和夫', 'かずお'),
    (r'AI', 'エーアイ'), (r'グループA', 'グループ・エー'), (r'グループB', 'グループ・ビー'), (r'グループC', 'グループ・シー'),
    (r'A案', 'エー案'), (r'A', 'エー'), (r'B', 'ビー'), (r'C', 'シー'),
    (r'基本型', 'きほんがた'), (r'場面起点型', '場面きてんがた'), (r'物語型', 'ものがたりがた'),
    (r'結論先行型', '結論せんこうがた'), (r'比較型', 'ひかくがた'), (r'型', 'かた'),
    (r'碁会所', 'ごかいしょ'), (r'行き先帳', 'いきさきちょう'), (r'行き先', 'いきさき'), (r'行きたい所', 'いきたいところ'), (r'いつもの所', 'いつものところ'),
    (r'一人', 'ひとり'), (r'二人', 'ふたり'), (r'一言', 'ひとこと'), (r'一続き', 'ひとつづき'), (r'一日', 'いちにち'),
    (r'一度目', 'いちどめ'), (r'三度目', 'さんどめ'), (r'三度', 'さんど'),
    (r'1つ1つ', 'ひとつひとつ'), (r'1つ', 'ひとつ'), (r'2つ', 'ふたつ'), (r'3つ', 'みっつ'), (r'4つ', 'よっつ'),
    (r'1文か2文', 'いちぶんか、にぶん'), (r'1文', 'いちぶん'), (r'2文', 'にぶん'), (r'20字', 'にじゅうじ'),
    (r'上の子', 'うえの子'), (r'下の子', 'したの子'), (r'来店組数', 'らいてんくみすう'), (r'人手', 'ひとで'),
    (r'(?<!関)係', 'かかり'), (r'性に合わない', 'しょうに合わない'), (r'想定の外', '想定のそと'), (r'(?<![反絶])対に', 'ついに'), (r'退けて', 'しりぞけて'),
    (r'間を置', 'まを置'), (r'(?<!発)表([をのは])', r'ひょう\1'),
    (r'家に', 'いえに'), (r'家の', 'いえの'), (r'他社', 'たしゃ'), (r'送り便', 'おくりびん'), (r'何の話', 'なんの話'),
    # 音声認識にかけて見つけた読み間違い（応える→いらえる など）と、その予防
    (r'応え', 'こたえ'), (r'終え', 'おえ'), (r'添え', 'そえ'), (r'通い', 'かよい'), (r'通う', 'かよう'),
    (r'(?<!判)断(?=[らりるっれ])', 'ことわ'), (r'対比', 'たいひ'), (r'杖', 'つえ'), (r'退店後', 'たいてんご'),
    # 「入れる」は「いれる」と読まれる（漢字とかな書きの音声の長さを比べて確かめた）。「はいれる」と読む箇所だけ読み替える
    (r'(から|のように|気軽に)入れ', r'\1はいれ'), (r'入れな', 'はいれな'),
    (r'(はっきり|別に)行', r'\1おこな'), (r'行った', 'おこなった'),
    (r'森田遥', '森田はるか'), (r'遥', 'はるか'), (r'学食', 'がくしょく'), (r'載', 'の'),
    (r'中の人', 'なかのひと'), (r'空いた', 'あいた'), (r'場違い', 'ばちがい'), (r'人柄', 'ひとがら'),
]


def spoken(text):
    for pat, rep in READ:
        text = re.sub(pat, rep, text)
    return text


def sha(kind, text):
    return hashlib.sha1(('%s|%s|%s' % (VOICE[kind], RATE, spoken(text))).encode('utf-8')).hexdigest()[:12]


def text_sha(text):
    """原稿そのものの指紋。build.js が、音声を作ったあとで原稿が変わっていないかを確かめるのに使う"""
    return hashlib.sha1(text.encode('utf-8')).hexdigest()[:12]


def fname(kind, key):
    return '%s-%s.mp3' % (kind, key.replace(':', '-'))


async def synth(kind, text, out, sem):
    async with sem:
        for attempt in range(4):
            try:
                await edge_tts.Communicate(spoken(text), VOICE[kind], rate=RATE).save(out)
                return
            except Exception as e:  # 一時的な通信の失敗はやり直す
                if attempt == 3:
                    raise
                print('  retry %s (%s)' % (os.path.basename(out), e))
                await asyncio.sleep(2 + attempt * 3)


async def topic(name, force, sem):
    spec = json.load(open(os.path.join(HERE, 'narration', name + '.json'), encoding='utf-8'))
    adir = os.path.join(HERE, 'audio', name)
    os.makedirs(adir, exist_ok=True)
    mf = os.path.join(adir, 'audio.json')
    old = json.load(open(mf, encoding='utf-8')) if os.path.exists(mf) else {}
    jobs, keep = [], set()
    for kind in ('talk', 'comment'):
        for key, text in spec[kind].items():
            out = os.path.join(adir, fname(kind, key))
            keep.add(os.path.basename(out))
            if force or not os.path.exists(out) or old.get(kind, {}).get(key, {}).get('sha1') != sha(kind, text):
                jobs.append((kind, text, out))
    await asyncio.gather(*(synth(kind, text, out, sem) for kind, text, out in jobs))
    # 原稿から消したキーの音声を消す
    for f in glob.glob(os.path.join(adir, '*.mp3')):
        if os.path.basename(f) not in keep:
            os.remove(f)
    manifest = {}
    for kind in ('talk', 'comment'):
        manifest[kind] = {key: {'sec': round(MP3(os.path.join(adir, fname(kind, key))).info.length, 2),
                                'sha1': sha(kind, text), 'text': text_sha(text)}
                          for key, text in spec[kind].items()}
    json.dump(manifest, open(mf, 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
    print('%s  発表 %d 本・解説 %d 本（作り直し %d）' % (name, len(manifest['talk']), len(manifest['comment']), len(jobs)))


async def main(argv):
    force = '--force' in argv
    names = [a for a in argv if not a.startswith('--')]
    if not names:
        names = sorted(f[:-5] for f in os.listdir(os.path.join(HERE, 'narration')) if f.endswith('.json'))
    sem = asyncio.Semaphore(4)
    for name in names:
        await topic(name, force, sem)

if __name__ == '__main__':
    asyncio.run(main(sys.argv[1:]))
