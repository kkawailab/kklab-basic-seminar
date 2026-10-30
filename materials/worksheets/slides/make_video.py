#!/usr/bin/env python3
"""各回ワークシートの解説スライド（worksheet-NN-slides.pptx）から MP4 動画を作る。

  各スライドを 1920x1080 の静止画に描き（video/build/NN/MM.png）、audio/NN/MM.mp3 の
  ナレーションを付けてスライドごとの動画にし、つないで video/worksheet-NN.mp4
  （スライドごとのチャプター付き）を書き出す。スライドごとの動画は video/build/ に残る。

  usage: python3 make_video.py              # 全回
         python3 make_video.py 03 07        # 第3回と第7回だけ
         python3 make_video.py --frames 03  # 静止画だけ描く（音声なしで、レイアウトを確かめる用）

  必要なもの:
    pip install python-pptx Pillow
    ffmpeg（なければ pip install imageio-ffmpeg で同梱の ffmpeg を使う）
    Noto Sans CJK JP のフォント（Debian/Ubuntu なら fonts-noto-cjk）
  スライドのフォントは游ゴシックだが、動画は Noto Sans CJK JP で描く。
  描画の部分は ../../summary/slides/make_video.py と同じ。
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
OUT = os.path.join(HERE, 'video')
BUILD = os.path.join(OUT, 'build')
ARTIST = '河合 勝彦（名古屋市立大学 経済学部）'

W, H = 1920, 1080          # 出力サイズ
SS = 2                     # 描画の倍率（縮小してなめらかにする）
DPI = W / 10 * SS          # スライドは 10 インチ幅
FPS = 15
LEAD, TAIL = 0.5, 0.8      # ナレーションの前後に置く無音（秒）
AUDIO_BITRATE = '64k'

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
        marl, indent, bullet = 0.0, 0.0, None   # 字下げと、ぶら下げの箇条書きの記号
        if ppr is not None:
            align = ppr.get('algn') or 'l'
            marl = px(int(ppr.get('marL') or 0))
            indent = px(int(ppr.get('indent') or 0))
            bu = ppr.find(qn('a:buChar'))
            if bu is not None:
                bullet = bu.get('char')
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
            lines += _wrap(g, width - marl) if g else [[]]
        total_h += sb + len(lines) * pitch + sa
        bu = None
        if bullet and lines and lines[0]:
            f0, c0 = lines[0][0][1], lines[0][0][2]
            bu = (bullet, f0, c0, marl + indent)
        placed.append((lines, align, pitch, max_size, sa, sb, marl, bu))
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
    for lines, align, pitch, size, sa, sb, marl, bu in placed:
        y += sb
        for li, line in enumerate(lines):
            w = sum(f.getlength(t) for t, f, _ in line)
            if align == 'ctr':
                x = lx + marl + (width - marl - w) / 2
            elif align == 'r':
                x = lx + width - w
            else:
                x = lx + marl
            base_y = y + (pitch - size / 72 * DPI) / 2 + size / 72 * DPI * 0.88
            if bu and li == 0:
                d.text((lx + bu[3], base_y), bu[0], font=bu[1], fill=bu[2], anchor='ls')
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
    elif 'UP' in name:
        head = min(w, h) * 0.5
        hy = y0 + head
        d.polygon([(x0 + w * 0.25, y1), (x1 - w * 0.25, y1), (x1 - w * 0.25, hy), (x1, hy),
                   (x0 + w / 2, y0), (x0, hy), (x0 + w * 0.25, hy)], fill=fill)
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


def render_frames(prs, frames):
    os.makedirs(frames, exist_ok=True)
    for i, slide in enumerate(prs.slides, 1):
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
        img.save(os.path.join(frames, '%02d.png' % i), optimize=True)
        print('  frame %02d' % i, '; '.join(issues) if issues else 'ok')


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


def encode(ff, build, nn, i, title, duration):
    total = round(LEAD + duration + TAIL, 3)
    out = os.path.join(build, '%02d.mp4' % i)
    run([ff, '-y', '-loglevel', 'error',
         '-loop', '1', '-framerate', str(FPS), '-i', os.path.join(build, '%02d.png' % i),
         '-i', os.path.join(HERE, 'audio', nn, '%02d.mp3' % i),
         '-af', 'adelay=delays=%d:all=1,apad,aresample=44100' % int(LEAD * 1000),
         '-t', str(total),
         '-c:v', 'libx264', '-preset', 'medium', '-tune', 'stillimage', '-crf', '22',
         '-pix_fmt', 'yuv420p', '-r', str(FPS), '-g', str(FPS * 60), '-sc_threshold', '0', '-profile:v', 'high', '-level', '4.0',
         '-c:a', 'aac', '-b:a', AUDIO_BITRATE, '-ac', '1',
         '-movflags', '+faststart', out])
    return total


def concat(ff, build, full, album, titles, totals):
    lst = os.path.join(build, 'concat.txt')
    with open(lst, 'w', encoding='utf-8') as f:
        for i in range(1, len(titles) + 1):
            f.write("file '%s'\n" % os.path.join(build, '%02d.mp4' % i))
    meta = os.path.join(build, 'chapters.txt')
    with open(meta, 'w', encoding='utf-8') as f:
        f.write(';FFMETADATA1\ntitle=%s\nartist=%s\n' % (album, ARTIST))
        t = 0.0
        for i, (title, total) in enumerate(zip(titles, totals), 1):
            f.write('[CHAPTER]\nTIMEBASE=1/1000\nSTART=%d\nEND=%d\ntitle=%d. %s\n'
                    % (round(t * 1000), round((t + total) * 1000), i, title))
            t += total
    run([ff, '-y', '-loglevel', 'error', '-f', 'concat', '-safe', '0', '-i', lst, '-i', meta,
         '-map', '0', '-map_metadata', '1', '-map_chapters', '1', '-c', 'copy', '-movflags', '+faststart', full])
    return t


def make_session(nn, frames_only):
    spec = json.load(open(os.path.join(HERE, 'content', nn + '.json'), encoding='utf-8'))
    deck = os.path.join(HERE, 'worksheet-%s-slides.pptx' % nn)
    build = os.path.join(BUILD, nn)
    prs = Presentation(deck)
    titles = [s['chapter'] for s in spec['slides']]
    assert len(prs.slides) == len(titles), '第%s回: スライド数と原稿の数が合いません' % nn
    print('第%d回 %s' % (spec['session'], spec['title']))
    render_frames(prs, build)
    if frames_only:
        return
    audio = json.load(open(os.path.join(HERE, 'audio', nn, 'audio.json'), encoding='utf-8'))
    assert len(audio) == len(titles), '第%s回: スライド数と音声の数が合いません' % nn
    ff = ffmpeg_exe()
    totals = [encode(ff, build, nn, i, t, audio['%02d' % i]['sec']) for i, t in enumerate(titles, 1)]
    full = os.path.join(OUT, 'worksheet-%s.mp4' % nn)
    album = '基礎演習Ⅱ 第%d回 %s　ワークシートの解説' % (spec['session'], spec['title'])
    t = concat(ff, build, full, album, titles, totals)
    print('  video  %s  %s  %.1f MB' % (os.path.basename(full), clock(t), os.path.getsize(full) / 1e6))


def clock(t):
    t = int(round(t))
    return '%d:%02d' % (t // 60, t % 60)


def write_page():
    """授業サイトの動画ページ（video/index.md）を、できている動画から書き出す"""
    url = "{{ '/materials/worksheets/slides/%s' | relative_url }}"
    rows, parts = [], []
    for f in sorted(os.listdir(os.path.join(HERE, 'content'))):
        if not re.fullmatch(r'\d\d\.json', f):
            continue
        nn = f[:2]
        mp4 = os.path.join(OUT, 'worksheet-%s.mp4' % nn)
        af = os.path.join(HERE, 'audio', nn, 'audio.json')
        if not (os.path.exists(mp4) and os.path.exists(af)):
            continue
        spec = json.load(open(os.path.join(HERE, 'content', f), encoding='utf-8'))
        audio = json.load(open(af, encoding='utf-8'))
        totals = [LEAD + audio['%02d' % i]['sec'] + TAIL for i in range(1, len(spec['slides']) + 1)]
        name = '第%d回　%s' % (spec['session'], spec['title'])
        rows.append('| [%s](#s%s) | %s |' % (name, nn, clock(sum(totals))))
        chapters, t = [], 0.0
        for s, d in zip(spec['slides'], totals):
            chapters.append('<li><span class="chapters__time">%s</span>%s</li>' % (clock(t), s['chapter']))
            t += d
        parts.append('\n'.join([
            '<h2 id="s%s">%s</h2>' % (nn, name),
            '',
            '<video controls preload="none" poster="%s" src="%s">お使いのブラウザは動画の再生に対応していません。</video>'
            % (url % ('video/poster-%s.jpg' % nn), url % ('video/worksheet-%s.mp4' % nn)),
            '',
            '[動画をダウンロード](%s)（MP4、%s、%.0fMB）　｜　[音声付きスライド（PowerPoint）](%s)'
            % (url % ('video/worksheet-%s.mp4' % nn), clock(sum(totals)), os.path.getsize(mp4) / 1e6,
               url % ('worksheet-%s-slides.pptx' % nn)),
            '',
            '<details class="chapters"><summary>チャプター（スライド%d枚）</summary><ol>%s</ol></details>'
            % (len(chapters), ''.join(chapters)),
        ]))
    page = '\n'.join([
        '---',
        'title: 各回ワークシートの解説　動画',
        'lead: 各回のワークシートの書き方を、記入例とともにスライドとナレーションで説明する動画です。ワークシートを手元に置いて見てください。',
        'description: 基礎演習Ⅱ 第1回〜第15回のワークシートの解説動画（MP4）と、音声付きスライド（PowerPoint）。',
        '---',
        '<!-- このページは make_video.py が書き出す。直すときは make_video.py の write_page を直す -->',
        '',
        '記入例は、書き方の目安です。AIの出力は毎回変わるので、記入例と同じ結果になるとは限りません。'
        '自分で確かめたことを、自分の言葉で書いてください。動画にはスライドごとのチャプターが入っているので、'
        '対応するプレーヤーでは見たいところへ飛べます。PowerPoint 版は、スライドショーを開始すると音声が自動で流れ、次のスライドへ進みます。',
        '',
        '<div class="table-scroll" markdown="1">',
        '',
        '| 回 | 長さ |',
        '|---|---:|',
    ] + rows + ['', '</div>', ''] + ['\n\n'.join(parts), ''])
    open(os.path.join(OUT, 'index.md'), 'w', encoding='utf-8').write(page)
    print('page  video/index.md  %d 回' % len(parts))


def poster(nn):
    """タイトルのスライドの静止画から、動画ページの表紙（video/poster-NN.jpg）を作る"""
    img = Image.open(os.path.join(BUILD, nn, '01.png')).convert('RGB').resize((960, 540), Image.LANCZOS)
    img.save(os.path.join(OUT, 'poster-%s.jpg' % nn), quality=82, optimize=True, progressive=True)


def main(argv):
    frames_only = '--frames' in argv
    nums = ['%02d' % int(a) for a in argv if a != '--frames']
    if not nums:
        nums = sorted(f[:2] for f in os.listdir(os.path.join(HERE, 'content')) if re.fullmatch(r'\d\d\.json', f))
    for nn in nums:
        make_session(nn, frames_only)
        if not frames_only:
            poster(nn)
    if not frames_only:
        write_page()


if __name__ == '__main__':
    main(sys.argv[1:])
