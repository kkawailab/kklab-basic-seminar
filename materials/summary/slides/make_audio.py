#!/usr/bin/env python3
"""narration.json のナレーションを edge-tts で音声（mp3）にし、長さを durations.json に書く。
   使い方: python3 make_audio.py            （全スライド）
           python3 make_audio.py 3 7        （指定したスライドだけ作り直す）
   必要なもの: pip install edge-tts mutagen"""
import asyncio, json, sys, os
import edge_tts
from mutagen.mp3 import MP3

HERE = os.path.dirname(os.path.abspath(__file__))
AUDIO = os.path.join(HERE, 'audio')
os.makedirs(AUDIO, exist_ok=True)
spec = json.load(open(os.path.join(HERE, 'narration.json'), encoding='utf-8'))
only = {int(a) for a in sys.argv[1:]}

async def synth(slide):
    out = os.path.join(AUDIO, '%02d.mp3' % slide['id'])
    tts = edge_tts.Communicate(slide['narration'], spec['voice'], rate=spec.get('rate', '+0%'))
    await tts.save(out)
    return out

async def main():
    for s in spec['slides']:
        if only and s['id'] not in only:
            continue
        out = await synth(s)
        print('%02d  %5.1f s  %s' % (s['id'], MP3(out).info.length, s['title']))
    durations = {('%02d' % s['id']): round(MP3(os.path.join(AUDIO, '%02d.mp3' % s['id'])).info.length, 2)
                 for s in spec['slides'] if os.path.exists(os.path.join(AUDIO, '%02d.mp3' % s['id']))}
    json.dump(durations, open(os.path.join(AUDIO, 'durations.json'), 'w'), indent=1)
    print('total %.1f min' % (sum(durations.values()) / 60))

asyncio.run(main())
