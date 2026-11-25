#!/usr/bin/env python3
"""作例のスライドに入れる図（ゾーニング図、場面のイラスト）を描く。
   images/<題材>/ に cover・avatar・plan・scene・story1〜3 の PNG を書き出す。大きさは、スライドの図の枠の縦横比に合わせてある。
   images/audio/ には、音声のボタン（発表の音声、教員の解説）の絵を書き出す。
   usage: python3 make_images.py            （全題材）
          python3 make_images.py family     （題材を指定）
   必要なもの: pip install Pillow、Noto Sans CJK JP のフォント（Debian/Ubuntu なら fonts-noto-cjk）"""
import math
import os
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
SS = 2            # 描画の倍率（縮小してなめらかにする）
PPI = 400         # 1インチあたりの画素数（書き出すとき）
FONT = next(p for p in ['/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc', '/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc'] if os.path.exists(p))

# 色（スライドのテーマに合わせる）
INK, INK2, PAPER, WHITE = '#1D2330', '#4A5261', '#ECEEE8', '#FFFFFF'
AMBER, AMBERTEXT, TINT = '#D68A17', '#8A570E', '#F6E4C3'
GREEN, SKY, RAIN, SKIN, GRAY = '#6F9A72', '#DCE2F9', '#9DB0F7', '#F0C9A4', '#B9BEB4'


class Canvas:
    """インチで大きさを決め、0〜100 の座標（横・縦それぞれ）か画素で描く"""

    def __init__(self, w_in, h_in, bg):
        self.W, self.H = int(w_in * PPI * SS), int(h_in * PPI * SS)
        self.im = Image.new('RGB', (self.W, self.H), bg)
        self.d = ImageDraw.Draw(self.im)
        self.u = self.W / 100          # 横幅の 1%

    def x(self, v): return v / 100 * self.W
    def y(self, v): return v / 100 * self.H
    def box(self, x0, y0, x1, y1): return (self.x(x0), self.y(y0), self.x(x1), self.y(y1))
    def font(self, size): return ImageFont.truetype(FONT, max(8, int(size * self.u)), index=0)

    def rr(self, b, r, fill=None, outline=None, width=0):
        self.d.rounded_rectangle(self.box(*b), radius=r * self.u, fill=fill, outline=outline, width=int(width * self.u))

    def circle(self, cx, cy, r, fill=None, outline=None, width=0):
        X, Y, R = self.x(cx), self.y(cy), r * self.u
        self.d.ellipse((X - R, Y - R, X + R, Y + R), fill=fill, outline=outline, width=int(width * self.u))

    def text(self, cx, cy, s, size, fill=INK, anchor='mm'):
        self.d.text((self.x(cx), self.y(cy)), s, font=self.font(size), fill=fill, anchor=anchor)

    def line(self, pts, fill, width, dash=None):
        P = [(self.x(a), self.y(b)) for a, b in pts]
        w = max(1, int(width * self.u))
        if not dash:
            self.d.line(P, fill=fill, width=w, joint='curve')
            return
        on, off = dash[0] * self.u, dash[1] * self.u
        for (x0, y0), (x1, y1) in zip(P, P[1:]):
            L = math.hypot(x1 - x0, y1 - y0)
            t = 0.0
            while t < L:
                a, b = t / L, min(t + on, L) / L
                self.d.line([(x0 + (x1 - x0) * a, y0 + (y1 - y0) * a), (x0 + (x1 - x0) * b, y0 + (y1 - y0) * b)], fill=fill, width=w)
                t += on + off

    def poly(self, pts, fill):
        self.d.polygon([(self.x(a), self.y(b)) for a, b in pts], fill=fill)

    def save(self, path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.im.resize((self.W // SS, self.H // SS), Image.LANCZOS).save(path, optimize=True)


# ------------------------------------------------------------
# 部品
# ------------------------------------------------------------
def pod(c, cx, cy, w, body=INK, heads=()):
    """自動運転の車（丸い箱型）。cx, cy は車体の中心、w は幅（横幅の%）。heads は窓に見える頭の色"""
    asp = c.W / c.H
    h = w * 0.46 * asp
    c.rr((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2), w * 0.16, fill=body)
    c.rr((cx - w * 0.4, cy - h * 0.36, cx + w * 0.4, cy + h * 0.08), w * 0.08, fill=SKY)
    for i, col in enumerate(heads):
        hx = cx - w * 0.28 + i * (w * 0.56 / max(1, len(heads) - 1) if len(heads) > 1 else 0)
        r = w * (0.075 if i % 2 == 0 else 0.058)
        c.circle(hx, cy - h * 0.08 + (0 if i % 2 == 0 else h * 0.06), r, fill=col)
    for sx in (-0.28, 0.28):
        c.circle(cx + w * sx, cy + h / 2, w * 0.085, fill=INK2)
        c.circle(cx + w * sx, cy + h / 2, w * 0.035, fill=PAPER)


def person(c, cx, base, h, body=AMBER, hair=INK, cane=False):
    """人（頭と胴）。base は足もとの縦位置（%）、h は身長（縦の%）"""
    asp = c.H / c.W
    r = h * 0.17 * asp
    hy = base - h + r / asp
    c.rr((cx - r * 1.25, hy + r / asp * 0.9, cx + r * 1.25, base), r * 1.1, fill=body)
    c.circle(cx, hy, r * 1.08, fill=hair)
    c.circle(cx, hy + r / asp * 0.12, r * 0.9, fill=SKIN)
    if cane:
        c.line([(cx + r * 1.7, base - h * 0.5), (cx + r * 1.7, base)], INK2, 0.5)


def house(c, x0, y0, x1, y1, door=AMBER):
    w = x1 - x0
    c.poly([(x0 - w * 0.08, y0 + (y1 - y0) * 0.38), ((x0 + x1) / 2, y0), (x1 + w * 0.08, y0 + (y1 - y0) * 0.38)], INK2)
    c.rr((x0, y0 + (y1 - y0) * 0.36, x1, y1), 0.6, fill=WHITE)
    c.rr((x0 + w * 0.56, y0 + (y1 - y0) * 0.58, x0 + w * 0.8, y1), 0.4, fill=door)
    c.rr((x0 + w * 0.14, y0 + (y1 - y0) * 0.52, x0 + w * 0.42, y0 + (y1 - y0) * 0.76), 0.4, fill=SKY)


def rain(c, x0, y0, x1, y1, n=26):
    for i in range(n):
        px = x0 + (x1 - x0) * ((i * 37) % 100) / 100
        py = y0 + (y1 - y0) * ((i * 53) % 100) / 100
        c.line([(px, py), (px - 1.2, py + 5)], RAIN, 0.45)


def tree(c, cx, base, s):
    asp = c.W / c.H
    c.line([(cx, base), (cx, base - s * asp * 1.0)], AMBERTEXT, s * 0.22)
    c.circle(cx, base - s * asp * 1.5, s, fill=GREEN)


# ------------------------------------------------------------
# ゾーニング図
#   zones: [(名前, x0, y0, x1, y1)]（%）、marks: {名前: 番号}（番号の丸を付け、色を付ける場所）
#   focus を渡すと、その場所だけを強調し、ほかは名前を出さない（物語型の場面用）
# ------------------------------------------------------------
def plan(path, w_in, h_in, zones, marks, route, entrance, focus=None):
    c = Canvas(w_in, h_in, '#F6F7F4')
    big = focus is not None
    c.rr((2, 3, 98, 97), 1.6, fill=WHITE, outline=INK2, width=0.35)
    centers = {}
    for name, x0, y0, x1, y1 in zones:
        hot = (name in focus) if big else (name in marks)
        c.rr((x0, y0, x1, y1), 1.2, fill=TINT if hot else PAPER, outline=AMBER if hot else None, width=0.45 if hot else 0)
        centers[name] = ((x0 + x1) / 2, (y0 + y1) / 2)
    pts = [(entrance, 97)] + [centers[n] for n in route]
    c.line(pts, AMBER, 0.5, dash=(1.6, 1.2))
    (xa, ya), (xb, yb) = pts[-2], pts[-1]
    ang = math.atan2((yb - ya) * c.H, (xb - xa) * c.W)
    tip = (c.x(xb), c.y(yb))
    L = 2.6 * c.u
    c.d.polygon([tip, (tip[0] - L * math.cos(ang - 0.5), tip[1] - L * math.sin(ang - 0.5)),
                 (tip[0] - L * math.cos(ang + 0.5), tip[1] - L * math.sin(ang + 0.5))], fill=AMBER)
    for name, x0, y0, x1, y1 in zones:
        hot = (name in focus) if big else (name in marks)
        if big and not hot:
            continue
        cx, cy = centers[name]
        size = (4.2 if big else 3.0) if hot else 2.5
        lines = name.split('\n')
        pitch = size * 1.25 * c.W / c.H
        for i, s in enumerate(lines):
            ty = cy + (i - (len(lines) - 1) / 2) * pitch
            tw = c.d.textlength(s, font=c.font(size)) / c.W * 100
            c.rr((cx - tw / 2 - 0.8, ty - pitch * 0.5, cx + tw / 2 + 0.8, ty + pitch * 0.5), 0.5, fill=TINT if hot else PAPER)
            c.text(cx, ty, s, size, fill=INK if hot else INK2)
        if name in marks and not big:
            c.circle(x0 + 3.2, y0 + 3.2 * c.W / c.H, 2.0, fill=AMBER)
            c.text(x0 + 3.2, y0 + 3.2 * c.W / c.H, str(marks[name]), 2.3)
    # 入口
    c.rr((entrance - (9 if big else 6), 95.2 - (2 if big else 0), entrance + (9 if big else 6), 99), 0.4, fill='#F6F7F4')
    c.text(entrance, 96.6, '▲ 入口', 3.4 if big else 2.5, fill=AMBERTEXT)
    c.save(path)


# ------------------------------------------------------------
# 題材ごとの図
# ------------------------------------------------------------
def family(out):
    zones = [('一週間\nシミュレーター', 5, 8, 33, 46), ('キャビン・リビング', 36, 8, 70, 60), ('見守りラウンジ', 73, 8, 95, 46),
             ('まちかどパーク', 5, 51, 33, 90), ('おためし車のりば', 36, 65, 70, 90), ('相談の小部屋', 73, 51, 95, 90)]
    marks = {'まちかどパーク': 1, 'キャビン・リビング': 2, '見守りラウンジ': 3}
    route = ['まちかどパーク', 'キャビン・リビング', '見守りラウンジ']
    plan(out + '/plan.png', 5.5, 2.95, zones, marks, route, 19)
    plan(out + '/story1.png', 2.8333, 1.6, zones, marks, ['まちかどパーク'], 19, focus=['まちかどパーク'])
    plan(out + '/story2.png', 2.8333, 1.6, zones, marks, ['まちかどパーク', '一週間\nシミュレーター', '見守りラウンジ'], 19, focus=['一週間\nシミュレーター', '見守りラウンジ'])

    # 表紙：家族を乗せた車が、一週間（月〜日）の行き先をめぐる
    c = Canvas(3.5, 4.4, TINT)
    c.circle(78, 12, 9, fill=AMBER)
    pts = [(22, 82), (70, 76), (80, 62), (38, 54), (22, 42), (60, 34), (78, 22)]
    c.line([(22, 97)] + pts, WHITE, 5.5)
    c.line([(22, 97)] + pts, INK2, 0.5, dash=(2.2, 1.8))
    for (px, py), day in zip(pts, '月火水木金土日'):
        c.circle(px, py, 5.2, fill=INK if day not in '土日' else AMBERTEXT)
        c.text(px, py, day, 5.0, fill=WHITE)
    tree(c, 12, 66, 4.5); tree(c, 90, 46, 4); tree(c, 48, 22, 3.5)
    pod(c, 56, 87, 46, heads=(SKIN, SKIN, SKIN, SKIN))
    c.save(out + '/cover.png')

    # 人物（丸く切り抜かれる）
    c = Canvas(1.5, 1.5, TINT)
    c.rr((18, 70, 82, 118), 16, fill=AMBER)
    c.circle(50, 44, 25, fill=INK)
    c.rr((25, 44, 75, 72), 8, fill=INK)
    c.circle(50, 48, 19.5, fill=SKIN)
    c.d.pieslice(c.box(29, 22, 71, 58), 180, 360, fill=INK)
    c.circle(42.5, 49, 2.0, fill=INK); c.circle(57.5, 49, 2.0, fill=INK)
    c.d.arc(c.box(43, 50, 57, 61), 20, 160, fill=AMBERTEXT, width=int(1.4 * c.u))
    c.save(out + '/avatar.png')

    # ある日：夜のリビングで、契約の画面を前に手が止まる
    c = Canvas(4.0, 2.95, TINT)
    c.rr((0, 78, 100, 100), 0, fill=PAPER)
    c.rr((6, 10, 30, 44), 1, fill=INK)                      # 夜の窓
    c.circle(22, 22, 3.2, fill=TINT); c.circle(23.6, 20.6, 2.8, fill=INK)
    c.rr((4, 60, 40, 84), 3, fill=INK2)                     # ソファ
    c.rr((4, 52, 12, 84), 3, fill=INK2)
    person(c, 24, 72, 36, body=AMBER, hair=INK)
    c.rr((54, 8, 94, 92), 3, fill=INK)                      # 契約の画面
    c.rr((56.5, 12, 91.5, 88), 1.5, fill=WHITE)
    c.rr((60, 17, 88, 43), 1.5, fill=SKY)
    pod(c, 74, 31, 20, heads=(SKIN,))
    for i, w in enumerate((26, 22, 24)):
        c.rr((60, 49 + i * 6, 60 + w, 52 + i * 6), 0.8, fill=PAPER)
    c.rr((60, 70, 88, 82), 2.5, fill=AMBER)
    c.text(74, 76, '同意する', 3.6, fill=INK)
    c.circle(44, 26, 6, fill=WHITE); c.text(44, 26, '？', 6.5, fill=AMBERTEXT)
    c.save(out + '/scene.png')

    # 新しい一日の3つ目：家に、おためしの車が届く
    c = Canvas(2.8333, 1.6, SKY)
    c.rr((0, 76, 100, 100), 0, fill=PAPER)
    c.circle(86, 18, 7, fill=AMBER)
    house(c, 8, 22, 44, 78)
    tree(c, 53, 78, 3.2)
    pod(c, 76, 71, 34, heads=(SKIN,))
    person(c, 30, 82, 40, body=AMBER, hair=INK)
    person(c, 39, 82, 26, body=AMBERTEXT, hair=INK)
    c.save(out + '/story3.png')


def senior(out):
    zones = [('ひと声ルーム', 5, 8, 33, 46), ('のりおり練習デッキ', 36, 8, 70, 60), ('茶の間ラウンジ', 73, 8, 95, 46),
             ('行き先の広場', 5, 51, 33, 90), ('送り便のりば', 36, 65, 70, 90), ('相談の小部屋', 73, 51, 95, 90)]
    marks = {'行き先の広場': 1, 'のりおり練習デッキ': 2, '茶の間ラウンジ': 3}
    route = ['行き先の広場', 'のりおり練習デッキ', '茶の間ラウンジ']
    plan(out + '/plan.png', 5.5, 2.95, zones, marks, route, 19)
    plan(out + '/story1.png', 2.8333, 1.6, zones, marks, ['行き先の広場'], 19, focus=['行き先の広場'])
    plan(out + '/story2.png', 2.8333, 1.6, zones, marks, ['行き先の広場', 'ひと声ルーム', 'のりおり練習デッキ'], 19, focus=['ひと声ルーム', 'のりおり練習デッキ'])

    # 表紙：自宅から、行きたい所へ道がのびる地図
    c = Canvas(3.5, 4.4, TINT)
    for gx in (20, 50, 80):
        c.line([(gx, 0), (gx, 100)], WHITE, 3.2)
    for gy in (18, 42, 66, 88):
        c.line([(0, gy), (100, gy)], WHITE, 3.2)
    home = (20, 88)
    dests = [((80, 18), '病院'), ((50, 42), '碁会所'), ((80, 66), '商店'), ((20, 18), '孫の家')]
    routes = [[home, (80, 88), (80, 66)], [home, (20, 42), (50, 42)], [home, (20, 66), (80, 66), (80, 18)], [home, (20, 18)]]
    for r in routes:
        c.line(r, AMBER, 0.9, dash=(2.4, 1.8))
    for (px, py), name in dests:
        c.circle(px, py - 4.6, 5.6, fill=INK)
        c.poly([(px - 4.0, py - 2.2), (px + 4.0, py - 2.2), (px, py + 2.4)], INK)
        c.circle(px, py - 4.6, 2.3, fill=AMBER)
        tw = c.d.textlength(name, font=c.font(4.6)) / c.W * 100
        c.rr((px - tw / 2 - 1.5, py + 3.2, px + tw / 2 + 1.5, py + 8.4), 1.2, fill=WHITE)
        c.text(px, py + 5.8, name, 4.6)
    house(c, 9, 78, 31, 94)
    pod(c, 56, 88, 26, heads=(GRAY,))
    c.save(out + '/cover.png')

    # 人物（丸く切り抜かれる）
    c = Canvas(1.5, 1.5, TINT)
    c.rr((16, 70, 84, 118), 16, fill=INK2)
    c.poly([(40, 70), (60, 70), (50, 84)], WHITE)
    c.circle(50, 45, 23, fill=GRAY)
    c.circle(50, 47, 20, fill=SKIN)
    c.d.pieslice(c.box(30, 25, 70, 52), 200, 340, fill=GRAY)
    for ex in (41.5, 58.5):
        c.circle(ex, 48, 5.6, outline=INK, width=1.3)
        c.circle(ex, 48, 1.7, fill=INK)
    c.line([(47.1, 48), (52.9, 48)], INK, 1.3)
    c.d.arc(c.box(43, 52, 57, 62), 20, 160, fill=AMBERTEXT, width=int(1.4 * c.u))
    c.save(out + '/avatar.png')

    # ある日：雨の朝、碁会所の誘いを断る
    c = Canvas(4.0, 2.95, TINT)
    c.rr((0, 78, 100, 100), 0, fill=PAPER)
    c.rr((6, 8, 44, 56), 1.2, fill=SKY)                     # 雨の窓
    rain(c, 8, 10, 42, 48)
    c.line([(25, 8), (25, 56)], WHITE, 1.2); c.line([(6, 32), (44, 32)], WHITE, 1.2)
    person(c, 58, 80, 52, body=INK2, hair=GRAY)
    c.rr((64, 40, 68, 54), 1.6, fill=INK)                   # 受話器
    c.rr((74, 10, 96, 46), 1.2, fill=WHITE)                 # カレンダー（×が3つ）
    c.rr((74, 10, 96, 18), 1.2, fill=AMBER)
    c.text(85, 14, '碁会所', 3.2, fill=INK)
    for i in range(3):
        c.text(79 + i * 6, 26, '×', 5.0, fill='#B3261E')
    for i in range(3):
        c.rr((76.5 + i * 6, 33, 81.5 + i * 6, 40), 0.6, fill=PAPER)
    c.rr((74, 62, 96, 80), 0.8, fill='#E3C68D')             # 碁盤
    for i in range(1, 5):
        c.line([(74 + i * 4.4, 62), (74 + i * 4.4, 80)], AMBERTEXT, 0.2)
    for i in range(1, 4):
        c.line([(74, 62 + i * 4.5), (96, 62 + i * 4.5)], AMBERTEXT, 0.2)
    c.circle(78.4, 66.5, 1.6, fill=INK); c.circle(87.2, 71, 1.6, fill=WHITE); c.circle(82.8, 75.5, 1.6, fill=INK)
    c.save(out + '/scene.png')

    # 新しい一日の3つ目：雨の朝、玄関先に車が来る
    c = Canvas(2.8333, 1.6, SKY)
    c.rr((0, 76, 100, 100), 0, fill=PAPER)
    rain(c, 2, 2, 98, 70, n=40)
    house(c, 6, 22, 42, 78)
    pod(c, 76, 71, 34, heads=())
    person(c, 52, 82, 44, body=INK2, hair=GRAY, cane=True)
    c.poly([(45, 32), (69, 32), (57, 20)], AMBER)           # 傘
    c.line([(57, 32), (57, 62)], INK2, 0.5)
    c.save(out + '/story3.png')


def student(out):
    zones = [('荷物の\n試し置き', 5, 8, 33, 46), ('おためしキャビン', 36, 8, 70, 60), ('呼び出し式の\n相談窓', 73, 8, 95, 46),
             ('計画テーブル', 5, 51, 33, 90), ('学生パスのりば', 36, 65, 70, 90), ('電源カウンター', 73, 51, 95, 90)]
    marks = {'計画テーブル': 1, 'おためしキャビン': 2, '呼び出し式の\n相談窓': 3}
    route = ['計画テーブル', 'おためしキャビン', '呼び出し式の\n相談窓']
    plan(out + '/plan.png', 5.5, 2.95, zones, marks, route, 19)
    plan(out + '/story1.png', 2.8333, 1.6, zones, marks, ['計画テーブル'], 19, focus=['計画テーブル'])
    plan(out + '/story2.png', 2.8333, 1.6, zones, marks, ['計画テーブル', 'おためしキャビン'], 19, focus=['おためしキャビン'])

    # 表紙：仲間が、入口の大きなテーブルで地図を囲む
    c = Canvas(3.5, 4.4, TINT)
    c.circle(80, 11, 8, fill=AMBER)
    for px, col, hair in ((18, GREEN, INK), (34, AMBER, AMBERTEXT), (50, INK2, INK), (66, AMBERTEXT, INK), (82, GREEN, AMBERTEXT)):
        person(c, px, 58, 30, body=col, hair=hair)
    c.rr((6, 46, 94, 72), 3, fill=WHITE)                    # テーブル（地図）
    c.rr((6, 70, 94, 74), 1.5, fill=INK2)
    pts = [(16, 64), (34, 55), (52, 62), (70, 53), (84, 60)]
    c.line(pts, AMBER, 0.9, dash=(2.4, 1.8))
    for px, py in (pts[0], pts[2], pts[4]):
        c.circle(px, py - 2.4, 3.0, fill=INK)
        c.poly([(px - 2.1, py - 1.2), (px + 2.1, py - 1.2), (px, py + 1.4)], INK)
        c.circle(px, py - 2.4, 1.2, fill=AMBER)
    tree(c, 12, 94, 4)
    pod(c, 58, 86, 46, heads=())
    c.save(out + '/cover.png')

    # 人物（丸く切り抜かれる）
    c = Canvas(1.5, 1.5, TINT)
    c.rr((16, 70, 84, 118), 16, fill=GREEN)
    c.rr((40, 66, 60, 78), 4, fill=WHITE)                   # フードのひも付きの襟
    c.circle(50, 43, 23.5, fill=INK)
    c.circle(50, 48, 19.5, fill=SKIN)
    c.d.pieslice(c.box(28, 20, 72, 60), 185, 355, fill=INK)
    c.circle(42.5, 50, 2.0, fill=INK); c.circle(57.5, 50, 2.0, fill=INK)
    c.d.arc(c.box(43, 51, 57, 62), 20, 160, fill=AMBERTEXT, width=int(1.4 * c.u))
    c.save(out + '/avatar.png')

    # ある日：夕方、販売店の前で立ち止まり、入らずに通りすぎる
    c = Canvas(4.0, 2.95, TINT)
    c.circle(12, 16, 7, fill=AMBER)                         # 夕日
    c.rr((0, 78, 100, 100), 0, fill=PAPER)
    c.rr((42, 12, 97, 78), 1.2, fill=WHITE)                 # 販売店
    c.rr((42, 12, 97, 25), 1.2, fill=INK)
    c.text(69.5, 18.5, 'モビリティ販売', 3.8, fill=AMBER)
    c.rr((46, 31, 76, 72), 1, fill=SKY)                     # ショーウインドー
    pod(c, 61, 57, 22, heads=())
    c.rr((81, 38, 93, 78), 0.8, fill=INK2)                  # 入口
    c.rr((79.5, 30, 94.5, 36), 0.8, fill=AMBER)
    c.text(87, 33, '受付', 3.0, fill=INK)
    person(c, 24, 84, 50, body=GREEN, hair=INK)
    c.rr((27.5, 52, 32.5, 68), 1.4, fill=INK2)              # かばん
    c.circle(35, 24, 6, fill=WHITE); c.text(35, 24, '？', 6.5, fill=AMBERTEXT)
    c.line([(15, 92), (4, 92)], AMBERTEXT, 0.6, dash=(1.6, 1.2))
    c.poly([(4, 89.4), (4, 94.6), (1, 92)], AMBERTEXT)
    c.save(out + '/scene.png')

    # 新しい一日の3つ目：合宿の朝、大学の前に車が来る
    c = Canvas(2.8333, 1.6, SKY)
    c.rr((0, 76, 100, 100), 0, fill=PAPER)
    c.circle(88, 16, 7, fill=AMBER)
    c.rr((4, 30, 36, 78), 0.8, fill=WHITE)                  # 大学
    c.rr((16, 12, 24, 32), 0.8, fill=WHITE)
    c.circle(20, 20, 2.4, fill=INK2)
    for i in range(3):
        for j in range(2):
            c.rr((8 + i * 9.5, 38 + j * 14, 13 + i * 9.5, 47 + j * 14), 0.4, fill=SKY)
    c.rr((16, 62, 24, 78), 0.6, fill=AMBERTEXT)
    pod(c, 74, 70, 40, heads=())
    for px, col, hair, h in ((41, GREEN, INK, 40), (48, AMBER, AMBERTEXT, 36), (55, INK2, INK, 40)):
        person(c, px, 84, h, body=col, hair=hair)
    c.rr((58.5, 68, 63.5, 84), 0.8, fill=AMBERTEXT)         # 荷物
    c.save(out + '/story3.png')


# ------------------------------------------------------------
# 音声のボタン（発表の音声、教員の解説）。スライドの右下に置く音声の絵として使う。角は build.js が丸く切り抜く
# ------------------------------------------------------------
def audio(out):
    for name, label, bg, fg in [('talk', '発表の音声', AMBER, INK), ('comment', '教員の解説', TINT, AMBERTEXT)]:
        c = Canvas(1.25, 0.3, bg)
        c.poly([(13, 28), (13, 72), (21.5, 50)], fg)        # 再生の三角
        c.text(58, 47, label, 10.6, fill=fg)
        c.save(out + '/%s.png' % name)


TOPICS = {'family': family, 'senior': senior, 'student': student, 'audio': audio}
for topic in (sys.argv[1:] or TOPICS):
    TOPICS[topic](os.path.join(HERE, 'images', topic))
    print('images/%s' % topic)
