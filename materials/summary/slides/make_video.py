#!/usr/bin/env python3
"""音声付きスライド（summary-01-05-slides.pptx）から MP4 動画を作る。

  各スライドを 1920x1080 の静止画に描き（video/frames/NN.png）、
  audio/NN.mp3 のナレーションを付けて video/NN.mp4 にする。
  最後に全スライドをつないだ通し動画 video/summary-01-05-video.mp4（章付き）を作る。

  usage: python3 make_video.py          # 全スライド
         python3 make_video.py 3 7      # 第3・7スライドだけ描き直す（通し動画も作り直す）

  必要なもの:
    pip install python-pptx Pillow
    ffmpeg（なければ pip install imageio-ffmpeg で同梱の ffmpeg を使う）
    Noto Sans CJK JP のフォント（Debian/Ubuntu なら fonts-noto-cjk）
  スライドのフォントは游ゴシックだが、動画は Noto Sans CJK JP で描く。
"""
import json
import os
import re
import shutil
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
DECK = os.path.join(HERE, 'summary-01-05-slides.pptx')
AUDIO = os.path.join(HERE, 'audio')
OUT = os.path.join(HERE, 'video')
FRAMES = os.path.join(OUT, 'frames')
FULL = os.path.join(OUT, 'summary-01-05-video.mp4')

W, H = 1920, 1080          # 出力サイズ
SS = 2                     # 描画の倍率（縮小してなめらかにする）
DPI = W / 10 * SS          # スライドは 10 インチ幅
FPS = 30
LEAD, TAIL = 0.5, 0.8      # ナレーションの前後に置く無音（秒）
AUDIO_BITRATE = '96k'

FONT_FILES = {
    False: ['/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf',
            '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc'],
    True: ['/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc',
           '/usr/share/fonts/opentype/noto/NotoSansCJKjp-Bold.otf',
           '/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc'],
}
INK, WHITE = '#1D2330', '#FFFFFF'

# 行頭に来てはいけない文字と、行末に来てはいけない文字（簡単な禁則処理）
NO_START = set('、。，．）」』】〕〉》］｝〙〗ゝゞ々ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶー・：；！？!?,.)]}%')
HANG = set('、。，．')
NO_END = set('（「『【〔〈《［｛〘〖(["{')
WORD = re.compile(r"[A-Za-z0-9'’.,%:/\-_&#@+]")


# ------------------------------------------------------------
# フォント
# ------------------------------------------------------------
_fonts = {}


def font(size_pt, bold):
    key = (round(size_pt * 4) / 4, bold)
    if key not in _fonts:
        path = next((p for p in FONT_FILES[bold] if os.path.exists(p)), None)
        if not path:
            sys.exit('Noto Sans CJK のフォントが見つかりません。FONT_FILES のパスを直してください。')
        px_size = max(4, int(round(size_pt / 72 * DPI)))
        _fonts[key] = ImageFont.truetype(path, px_size, index=0)
    return _fonts[key]


def px(emu):
    return Emu(emu).inches * DPI


# ------------------------------------------------------------
# 文字の描画（PowerPoint の文字枠のように折り返して配置する）
# ------------------------------------------------------------
def _srgb(el):
    if el is None:
        return None
    c = el.find(qn('a:solidFill'))
    if c is None:
        return None
    s = c.find(qn('a:srgbClr'))
    return '#' + s.get('val') if s is not None else None


def _run_props(rpr, defaults):
    size, bold, color = defaults
    if rpr is not None:
        if rpr.get('sz'):
            size = int(rpr.get('sz')) / 100
        if rpr.get('b') is not None:
            bold = rpr.get('b') in ('1', 'true')
        c = _srgb(rpr)
        if c:
            color = c
    return size, bold, color


def _tokens(text, f, color):
    """文字列を、折り返しの単位（英単語・空白・1文字）に分ける"""
    out, i = [], 0
    while i < len(text):
        ch = text[i]
        if WORD.match(ch):
            j = i
            while j < len(text) and WORD.match(text[j]):
                j += 1
            out.append((text[i:j], f, color))
            i = j
        else:
            out.append((ch, f, color))
            i += 1
    return out


def _wrap(tokens, width):
    lines, cur, cur_w = [], [], 0.0

    def tw(t):
        return t[1].getlength(t[0])

    for t in tokens:
        w = tw(t)
        if not cur or cur_w + w <= width:
            cur.append(t); cur_w += w
            continue
        if t[0] in HANG:                       # 句読点はぶら下げる
            cur.append(t); cur_w += w
            continue
        if t[0] in NO_START and len(cur) >= 2:  # 行頭禁則：前の文字ごと送る
            prev = cur.pop()
            lines.append(cur); cur = [prev, t]
        elif cur and cur[-1][0] in NO_END and len(cur) >= 2:  # 行末禁則
            prev = cur.pop()
            lines.append(cur); cur = [prev, t]
        else:
            lines.append(cur); cur = [t]
        if cur and cur[0][0] == ' ':
            cur = cur[1:]
        cur_w = sum(tw(x) for x in cur)
    lines.append(cur)
    return lines


def layout_text(txbody, box, default_color, margins=None):
    """txBody を箱に収めたときの行の配置を返す。描画は draw_text で行う。"""
    x0, y0, x1, y1 = box
    bp = txbody.find(qn('a:bodyPr'))
    ins = {'l': 91440, 'r': 91440, 't': 45720, 'b': 45720}
    if margins:
        ins = margins
    elif bp is not None:
        for k in ins:
            v = bp.get(k + 'Ins')
            if v is not None:
                ins[k] = int(v)
    ml, mr, mt, mb = (px(ins[k]) for k in 'lrtb')
    anchor = (bp.get('anchor') if bp is not None else None) or 't'
    width = max(10, x1 - x0 - ml - mr)
    # 既定の文字の設定（lstStyle の lvl1pPr/defRPr）
    base = (18.0, False, default_color)
    lst = txbody.find(qn('a:lstStyle'))
    if lst is not None:
        l1 = lst.find(qn('a:lvl1pPr'))
        if l1 is not None:
            base = _run_props(l1.find(qn('a:defRPr')), base)
    placed, total_h = [], 0.0
    for p in txbody.findall(qn('a:p')):
        ppr = p.find(qn('a:pPr'))
        align, ls, sa, sb = 'l', 1.0, 0.0, 0.0
        if ppr is not None:
            align = ppr.get('algn') or 'l'
            e = ppr.find(qn('a:lnSpc'))
            if e is not None and e.find(qn('a:spcPct')) is not None:
                ls = int(e.find(qn('a:spcPct')).get('val')) / 100000
            e = ppr.find(qn('a:spcAft'))
            if e is not None and e.find(qn('a:spcPts')) is not None:
                sa = int(e.find(qn('a:spcPts')).get('val')) / 100 / 72 * DPI
            e = ppr.find(qn('a:spcBef'))
            if e is not None and e.find(qn('a:spcPts')) is not None:
                sb = int(e.find(qn('a:spcPts')).get('val')) / 100 / 72 * DPI
        tokens, max_size = [], 0.0
        for child in p:
            tag = child.tag
            if tag in (qn('a:r'), qn('a:fld')):
                size, bold, color = _run_props(child.find(qn('a:rPr')), base)
                t = child.find(qn('a:t'))
                text = t.text if t is not None and t.text else ''
                if tag == qn('a:fld') and child.get('type') == 'slidenum':
                    text = str(layout_text.slide_no)
                tokens += _tokens(text, font(size, bold), color)
                max_size = max(max_size, size)
            elif tag == qn('a:br'):
                tokens.append(('\n', None, None))
        if not max_size:
            epr = p.find(qn('a:endParaRPr'))
            max_size = _run_props(epr, base)[0]
        pitch = max_size / 72 * DPI * 1.2 * ls
        # 改行で分けてから折り返す
        groups, cur = [], []
        for t in tokens:
            if t[0] == '\n':
                groups.append(cur); cur = []
            else:
                cur.append(t)
        groups.append(cur)
        lines = []
        for g in groups:
            lines += _wrap(g, width) if g else [[]]
        total_h += sb + len(lines) * pitch + sa
        placed.append((lines, align, pitch, max_size, sa, sb))
    inner_h = y1 - y0 - mt - mb
    y = y0 + mt
    if anchor == 'ctr':
        y += (inner_h - total_h) / 2
    elif anchor == 'b':
        y = y1 - mb - total_h
    return placed, (x0 + ml, y, width), total_h, inner_h


layout_text.slide_no = 0


def draw_text(d, txbody, box, default_color, margins=None):
    placed, (lx, y, width), total_h, inner_h = layout_text(txbody, box, default_color, margins)
    for lines, align, pitch, size, sa, sb in placed:
        y += sb
        for line in lines:
            w = sum(f.getlength(t) for t, f, _ in line)
            if align == 'ctr':
                x = lx + (width - w) / 2
            elif align == 'r':
                x = lx + width - w
            else:
                x = lx
            base_y = y + (pitch - size / 72 * DPI) / 2 + size / 72 * DPI * 0.88
            for t, f, color in line:
                d.text((x, base_y), t, font=f, fill=color, anchor='ls')
                x += f.getlength(t)
            y += pitch
        y += sa
    return total_h > inner_h + 2


# ------------------------------------------------------------
# 図形の描画
# ------------------------------------------------------------
def fill_of(shape):
    try:
        if shape.fill.type == 1:
            return '#' + str(shape.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def draw_arrow(d, box, name, fill):
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    if 'RIGHT' in name:
        head = min(w, h) * 0.5
        hx = x1 - head
        d.polygon([(x0, y0 + h * 0.25), (hx, y0 + h * 0.25), (hx, y0), (x1, y0 + h / 2),
                   (hx, y1), (hx, y1 - h * 0.25), (x0, y1 - h * 0.25)], fill=fill)
    else:
        head = min(w, h) * 0.5
        hy = y1 - head
        d.polygon([(x0 + w * 0.25, y0), (x1 - w * 0.25, y0), (x1 - w * 0.25, hy), (x1, hy),
                   (x0 + w / 2, y1), (x0, hy), (x0 + w * 0.25, hy)], fill=fill)


def draw_table(d, shape, box, default_color):
    tbl = shape.table
    issues = []
    y = box[1]
    widths = [px(c.width) for c in tbl.columns]
    for r in tbl.rows:
        # 文字が収まるように行の高さを伸ばす（PowerPoint と同じ）
        rh = px(r.height)
        for ci, c in enumerate(r.cells):
            m = {'l': c.margin_left, 'r': c.margin_right, 't': c.margin_top, 'b': c.margin_bottom}
            _, _, th, _ = layout_text(c._tc.txBody, (0, 0, widths[ci], rh), default_color, m)
            rh = max(rh, th + px(m['t']) + px(m['b']))
        x = box[0]
        for ci, c in enumerate(r.cells):
            cw = widths[ci]
            cell_box = (x, y, x + cw, y + rh)
            cf = None
            try:
                if c.fill.type == 1:
                    cf = '#' + str(c.fill.fore_color.rgb)
            except Exception:
                pass
            if cf:
                d.rectangle(cell_box, fill=cf)
            tcpr = c._tc.find(qn('a:tcPr'))
            if tcpr is not None:
                for side, seg in (('lnL', (x, y, x, y + rh)), ('lnR', (x + cw, y, x + cw, y + rh)),
                                  ('lnT', (x, y, x + cw, y)), ('lnB', (x, y + rh, x + cw, y + rh))):
                    ln = tcpr.find(qn('a:' + side))
                    if ln is None or ln.find(qn('a:noFill')) is not None:
                        continue
                    col = _srgb(ln) or '#B9BEB4'
                    lw = max(1, int(round(int(ln.get('w') or 9525) / 12700 / 72 * DPI)))
                    d.line(seg, fill=col, width=lw)
            m = {'l': c.margin_left, 'r': c.margin_right, 't': c.margin_top, 'b': c.margin_bottom}
            if draw_text(d, c._tc.txBody, cell_box, default_color, m):
                issues.append('table overflow: ' + c.text_frame.text[:20])
            x += cw
        y += rh
    return issues


def draw_shapes(d, shapes, default_color):
    issues = []
    for sh in shapes:
        try:
            if sh.shape_type in (MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.PICTURE):
                continue  # 音声のアイコンは動画には描かない
            if sh.left is None or sh.width is None:
                continue
            box = (px(sh.left), px(sh.top), px(sh.left + sh.width), px(sh.top + sh.height))
            if sh.has_table:
                issues += draw_table(d, sh, box, default_color)
                continue
            if not (sh.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX) or sh.is_placeholder):
                continue
            fill = fill_of(sh)
            name = ''
            try:
                if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    name = sh.auto_shape_type.name
            except Exception:
                pass
            if fill:
                if 'OVAL' in name or 'ELLIPSE' in name:
                    d.ellipse(box, fill=fill)
                elif 'ARROW' in name:
                    draw_arrow(d, box, name, fill)
                elif 'ROUNDED' in name:
                    adj = sh.adjustments[0] if len(sh.adjustments) else 0.16667
                    radius = adj * min(box[2] - box[0], box[3] - box[1])
                    d.rounded_rectangle(box, radius=radius, fill=fill)
                else:
                    d.rectangle(box, fill=fill)
            if sh.has_text_frame and sh.text_frame.text.strip():
                if draw_text(d, sh.text_frame._txBody, box, default_color):
                    issues.append('overflow: ' + sh.text_frame.text[:30])
        except Exception as e:
            issues.append('ERR %s: %s' % (getattr(sh, 'name', '?'), e))
    return issues


def render_frames(prs, wanted):
    os.makedirs(FRAMES, exist_ok=True)
    for i, slide in enumerate(prs.slides, 1):
        if wanted and i not in wanted:
            continue
        layout_text.slide_no = i
        bg = WHITE
        try:
            bg = fill_of(slide.slide_layout.background) or bg
        except Exception:
            pass
        try:
            if slide.background.fill.type == 1:
                bg = '#' + str(slide.background.fill.fore_color.rgb)
        except Exception:
            pass
        default = INK if bg.upper() == WHITE else WHITE
        img = Image.new('RGB', (W * SS, H * SS), bg)
        d = ImageDraw.Draw(img)
        issues = draw_shapes(d, [s for s in slide.slide_layout.shapes if not s.is_placeholder], default)
        issues += draw_shapes(d, slide.shapes, default)
        img = img.resize((W, H), Image.LANCZOS)
        img.save(os.path.join(FRAMES, '%02d.png' % i), optimize=True)
        print('frame %02d' % i, '; '.join(issues) if issues else 'ok')


# ------------------------------------------------------------
# ffmpeg
# ------------------------------------------------------------
def ffmpeg_exe():
    exe = shutil.which('ffmpeg')
    if exe:
        return exe
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError:
        sys.exit('ffmpeg が見つかりません。インストールするか、pip install imageio-ffmpeg を実行してください。')


def run(cmd):
    r = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
    if r.returncode:
        sys.exit('ffmpeg failed:\n' + r.stderr[-2000:])


def encode(ff, i, title, duration):
    total = round(LEAD + duration + TAIL, 3)
    out = os.path.join(OUT, '%02d.mp4' % i)
    run([ff, '-y', '-loglevel', 'error',
         '-loop', '1', '-framerate', str(FPS), '-i', os.path.join(FRAMES, '%02d.png' % i),
         '-i', os.path.join(AUDIO, '%02d.mp3' % i),
         '-af', 'adelay=delays=%d:all=1,apad,aresample=44100' % int(LEAD * 1000),
         '-t', str(total),
         '-c:v', 'libx264', '-preset', 'medium', '-tune', 'stillimage', '-crf', '22',
         '-pix_fmt', 'yuv420p', '-r', str(FPS), '-g', str(FPS * 10), '-sc_threshold', '0', '-profile:v', 'high', '-level', '4.0',
         '-c:a', 'aac', '-b:a', AUDIO_BITRATE, '-ac', '1',
         '-metadata', 'title=%s' % title, '-metadata', 'artist=河合 勝彦（名古屋市立大学 経済学部）',
         '-movflags', '+faststart', out])
    print('clip  %02d' % i, '%5.1f s' % total, '%5.1f MB' % (os.path.getsize(out) / 1e6), title)
    return total


def concat(ff, titles, totals):
    lst = os.path.join(FRAMES, 'concat.txt')
    with open(lst, 'w', encoding='utf-8') as f:
        for i in range(1, len(titles) + 1):
            f.write("file '%s'\n" % os.path.join(OUT, '%02d.mp4' % i))
    meta = os.path.join(FRAMES, 'chapters.txt')
    with open(meta, 'w', encoding='utf-8') as f:
        f.write(';FFMETADATA1\ntitle=基礎演習Ⅱ 第1回〜第5回 総まとめ\nartist=河合 勝彦（名古屋市立大学 経済学部）\n')
        t = 0.0
        for i, (title, total) in enumerate(zip(titles, totals), 1):
            f.write('[CHAPTER]\nTIMEBASE=1/1000\nSTART=%d\nEND=%d\ntitle=%d. %s\n'
                    % (round(t * 1000), round((t + total) * 1000), i, title))
            t += total
    run([ff, '-y', '-loglevel', 'error', '-f', 'concat', '-safe', '0', '-i', lst, '-i', meta,
         '-map', '0', '-map_metadata', '1', '-map_chapters', '1', '-c', 'copy', '-movflags', '+faststart', FULL])
    print('full  %s  %.1f min  %.1f MB' % (os.path.basename(FULL), t / 60, os.path.getsize(FULL) / 1e6))


def main(argv):
    wanted = {int(a) for a in argv}
    spec = json.load(open(os.path.join(HERE, 'narration.json'), encoding='utf-8'))
    titles = [s['title'] for s in spec['slides']]
    durations = json.load(open(os.path.join(AUDIO, 'durations.json'), encoding='utf-8'))
    prs = Presentation(DECK)
    assert len(prs.slides) == len(titles) == len(durations), 'スライド数と原稿・音声の数が合いません'
    os.makedirs(OUT, exist_ok=True)
    render_frames(prs, wanted)
    ff = ffmpeg_exe()
    totals = []
    for i, title in enumerate(titles, 1):
        dur = durations['%02d' % i]
        if wanted and i not in wanted:
            totals.append(round(LEAD + dur + TAIL, 3))
            continue
        totals.append(encode(ff, i, title, dur))
    concat(ff, titles, totals)


if __name__ == '__main__':
    main(sys.argv[1:])
