#!/usr/bin/env python3
"""build.js が書き出した pptx を仕上げる。
   ・テーマの色とフォント（日本語フォントを Yu Gothic に）
   ・各スライドの音声を「自動再生・スライドショー中は非表示」にし、音声の長さで次のスライドへ自動送り
   ・mp3 のコンテンツタイプを audio/mpeg にする
   使い方: python3 finish.py  （summary-01-05-slides.pptx を上書き）"""
import json, os, re, sys, zipfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
PPTX = os.path.join(HERE, 'summary-01-05-slides.pptx')
durations = json.load(open(os.path.join(HERE, 'audio', 'durations.json')))
GAP_MS = 800  # 音声が終わってから次のスライドへ送るまでの間

COLORS = dict(dk1='1D2330', lt1='FFFFFF', dk2='4A5261', lt2='ECEEE8', accent1='2747C9', accent2='D68A17',
              accent3='DCE2F9', accent4='B3261E', accent5='757D8A', accent6='F6E4C3', hlink='2747C9', folHlink='4A5261')
JPFONT = 'Yu Gothic'

TIMING = '''<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="6" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq><p:audio><p:cMediaNode vol="80000" showWhenStopped="0"><p:cTn id="7" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''

def fix_slide(xml, n):
    # 音声の図形（pptxgenjs は videoFile と書くので audioFile に直す）
    m = re.search(r'<p:pic>\s*<p:nvPicPr><p:cNvPr id="(\d+)" name="Media[^"]*">', xml)
    if not m:
        return xml, False
    spid = m.group(1)
    xml = xml.replace('<a:videoFile r:link=', '<a:audioFile r:link=')
    dur_ms = int(round(durations.get('%02d' % n, 0) * 1000))
    trans = '<p:transition spd="med" advClick="1" advTm="%d"><p:fade/></p:transition>' % (dur_ms + GAP_MS)
    timing = TIMING.format(dur=dur_ms, spid=spid)
    xml = re.sub(r'<p:transition.*?</p:transition>', '', xml, flags=re.S)
    xml = re.sub(r'<p:timing>.*?</p:timing>', '', xml, flags=re.S)
    if '</p:clrMapOvr>' in xml:
        xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + trans + timing, 1)
    else:
        xml = xml.replace('</p:sld>', trans + timing + '</p:sld>')
    return xml, True

def fix_theme(xml):
    for k, v in COLORS.items():
        xml = re.sub(r'<a:%s>.*?</a:%s>' % (k, k), '<a:%s><a:srgbClr val="%s"/></a:%s>' % (k, v, k), xml, count=1, flags=re.S)
    xml = xml.replace('<a:clrScheme name="Office">', '<a:clrScheme name="Kiso Seminar">')
    xml = xml.replace('<a:fontScheme name="Office">', '<a:fontScheme name="Kiso Seminar">')
    xml = xml.replace('<a:ea typeface=""/>', '<a:ea typeface="%s"/>' % JPFONT)
    xml = re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>', '<a:font script="Jpan" typeface="%s"/>' % JPFONT, xml)
    return xml

def main():
    tmp = PPTX + '.tmp'
    done = []
    with zipfile.ZipFile(PPTX) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', item.filename)
            if m:
                xml, ok = fix_slide(data.decode('utf-8'), int(m.group(1)))
                if ok: done.append(int(m.group(1)))
                data = xml.encode('utf-8')
            elif item.filename == 'ppt/theme/theme1.xml':
                data = fix_theme(data.decode('utf-8')).encode('utf-8')
            elif item.filename == '[Content_Types].xml':
                data = data.decode('utf-8').replace('ContentType="audio/mp3"', 'ContentType="audio/mpeg"').encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, PPTX)
    print('audio set on slides:', sorted(done), '(%d)' % len(done))
    print('total %.1f min incl. gaps' % ((sum(durations.values()) + GAP_MS / 1000 * len(durations)) / 60))

main()
