#!/usr/bin/env python3
"""content/NN.json のナレーションを edge-tts で音声（audio/NN/MM.mp3）にし、長さを audio/NN/audio.json に書く。
   原稿（と声・速さ）が前回から変わったスライドだけ作り直す。
   使い方: python3 make_audio.py              （全回）
           python3 make_audio.py 03 07        （第3回と第7回）
           python3 make_audio.py --force 03   （第3回をすべて作り直す）
   必要なもの: pip install edge-tts mutagen（Microsoft Edge の読み上げ音声を使うので、ネットワークが必要）"""
import asyncio, glob, hashlib, json, os, re, sys
import edge_tts
from mutagen.mp3 import MP3

HERE = os.path.dirname(os.path.abspath(__file__))
VOICE, RATE = 'ja-JP-NanamiNeural', '+0%'


def sha(text):
    return hashlib.sha1(('%s|%s|%s' % (VOICE, RATE, text)).encode('utf-8')).hexdigest()[:12]


async def synth(text, out, sem):
    async with sem:
        for attempt in range(4):
            try:
                await edge_tts.Communicate(text, VOICE, rate=RATE).save(out)
                return
            except Exception as e:  # 一時的な通信の失敗はやり直す
                if attempt == 3:
                    raise
                print('  retry %s (%s)' % (os.path.basename(out), e))
                await asyncio.sleep(2 + attempt * 3)


async def session(nn, force, sem):
    spec = json.load(open(os.path.join(HERE, 'content', nn + '.json'), encoding='utf-8'))
    adir = os.path.join(HERE, 'audio', nn)
    os.makedirs(adir, exist_ok=True)
    mf = os.path.join(adir, 'audio.json')
    old = json.load(open(mf, encoding='utf-8')) if os.path.exists(mf) else {}
    jobs, ids = [], []
    for i, s in enumerate(spec['slides'], 1):
        key, out = '%02d' % i, os.path.join(adir, '%02d.mp3' % i)
        ids.append(key)
        if force or not os.path.exists(out) or old.get(key, {}).get('sha1') != sha(s['narration']):
            jobs.append((key, s, out))
    await asyncio.gather(*(synth(s['narration'], out, sem) for _, s, out in jobs))
    # 回のスライドが減ったときは、余った音声を消す
    for f in glob.glob(os.path.join(adir, '[0-9][0-9].mp3')):
        if os.path.basename(f)[:2] not in ids:
            os.remove(f)
    manifest = {}
    for i, s in enumerate(spec['slides'], 1):
        key = '%02d' % i
        manifest[key] = {'sec': round(MP3(os.path.join(adir, key + '.mp3')).info.length, 2), 'sha1': sha(s['narration'])}
    json.dump(manifest, open(mf, 'w', encoding='utf-8'), indent=1)
    total = sum(m['sec'] for m in manifest.values())
    print('第%d回  %d枚（作り直し %d）  %.1f 分' % (spec['session'], len(manifest), len(jobs), total / 60))


async def main(argv):
    force = '--force' in argv
    nums = ['%02d' % int(a) for a in argv if a != '--force']
    if not nums:
        nums = sorted(f[:2] for f in os.listdir(os.path.join(HERE, 'content')) if re.fullmatch(r'\d\d\.json', f))
    sem = asyncio.Semaphore(4)
    for nn in nums:
        await session(nn, force, sem)

asyncio.run(main(sys.argv[1:]))
