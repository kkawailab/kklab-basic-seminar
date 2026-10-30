// 基礎演習Ⅱ 各回ワークシートの解説　音声付きスライドの組み立て
//   node build.js            全回（content/NN.json → worksheet-NN-slides.pptx）
//   node build.js 03 07      第3回と第7回だけ
//   音声（audio/NN/MM.mp3）がまだない回は、音声なしで組む（レイアウトの確認用）。
//   そのあと python3 finish.py で、テーマのフォント・色の設定と、音声の自動再生・自動送りを書き込む。
//   スライドの中身は content/NN.json に、部品（card、steps など）の組み合わせで書く。部品の一覧は README.md。
//   必要なもの: npm install pptxgenjs
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const HERE = __dirname;
const pad2 = n => String(n).padStart(2, '0');

// 色（授業サイト assets/css/style.css・ワークシートに合わせる）
const INK = '1D2330', INK2 = '4A5261', INK3 = '757D8A', PAPER = 'ECEEE8', RULE = 'B9BEB4', WHITE = 'FFFFFF';
const BAD = 'B3261E', BADTINT = 'F9DEDA', OK = '2E7D32', OKTINT = 'DCEFDD';
// フェーズの色。main は塗り、text は白地の文字、on は main の上の文字、sub は表紙の補助の文字、dot は表紙の他の回の丸
const PHASES = {
  1: { name: '探究の基礎', main: '2747C9', text: '1F3BA8', tint: 'DCE2F9', on: WHITE, sub: 'DCE2F9', dot: '5570D8' },
  2: { name: '実践プロジェクト', main: 'D68A17', text: '8A570E', tint: 'F6E4C3', on: INK, sub: '3F2E10', dot: 'E7AD55' },
  3: { name: 'キャリア講習', main: '6E7684', text: '4A5261', tint: 'E0E3E4', on: WHITE, sub: 'E0E3E4', dot: '8D94A0' },
};
const FONT = 'Yu Gothic';

// 本文の領域（10 x 5.625 in のスライド）
const X0 = 0.5, CW = 9.0, Y0 = 1.4, Y1 = 4.9;
const ROWGAP = 0.18, BLOCKGAP = 0.14, COLGAP = 0.35;

// ------------------------------------------------------------
// 文字の量から高さを見積もる（日本語は全角1文字＝フォントサイズ分の幅、行の高さは 1.2 倍×行間）
// ------------------------------------------------------------
function emWidth(str) {
  let em = 0;
  for (const ch of str.replace(/\*\*/g, '')) {
    const c = ch.codePointAt(0);
    if (ch === ' ') em += 0.3; else if (c >= 0x2000 || ch === '×') em += 1.0; else em += 0.58;
  }
  return em;
}
// 「・」「□ 」で始まる段落は、ぶら下げの箇条書きにする（記号の幅だけ字下げ）
const BULLETS = [['・', '30FB', 1.0], ['□ ', '25A1', 1.25]];
function bulletOf(p) { return BULLETS.find(([pre]) => p.startsWith(pre)); }
function lineCount(text, size, w) {
  let n = 0;
  for (let p of String(text).split('\n')) {
    const b = bulletOf(p);
    let ww = w;
    if (b) { p = p.slice(b[0].length); ww -= b[2] * size / 72; }
    n += Math.max(1, Math.ceil(emWidth(p) * size / 72 / Math.max(ww * 0.97, 0.1)));  // 3% の余裕を見る
  }
  return n;
}
const lineH = (size, ls) => size * 1.2 * ls / 72;
const textH = (text, size, w, ls = 1.3) => lineCount(text, size, w) * lineH(size, ls);

// 「**太字**」を runs にする。段落は \n で分ける
function runs(text, base, size) {
  const out = [];
  const paras = String(text).split('\n');
  paras.forEach((p, pi) => {
    const b = bulletOf(p);
    if (b) p = p.slice(b[0].length);
    const parts = p.split('**');
    const items = [];
    parts.forEach((t, i) => {
      if (!t) return;
      const strong = i % 2 === 1;
      items.push({ text: t, options: { bold: strong || !!base.bold, color: strong && base.em ? base.em : base.color } });
    });
    if (!items.length) items.push({ text: ' ', options: { color: base.color } });
    if (b) items[0].options.bullet = { characterCode: b[1], indent: Math.round(b[2] * size * 10) / 10 };
    if (pi < paras.length - 1) items[items.length - 1].options.breakLine = true;
    out.push(...items);
  });
  return out;
}

// ------------------------------------------------------------
// 描画の基本
// ------------------------------------------------------------
let pres, slide, ph, warnings, where;
function txt(text, o) {
  const size = o.size || 12, ls = o.ls || 1.3;
  if (!o.noCheck) {
    const need = textH(text, size, o.w, ls);
    if (need > o.h + 0.03) warnings.push(`${where} ${o.label || 'text'}: ${need.toFixed(2)}in > ${o.h.toFixed(2)}in「${String(text).slice(0, 16)}…」`);
  }
  slide.addText(runs(text, { color: o.color || INK, bold: !!o.bold, em: o.em }, size), {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: FONT, fontSize: size, align: o.align || 'left', valign: o.valign || 'top',
    margin: 0, isTextBox: true, lineSpacingMultiple: ls, paraSpaceAfter: o.paraSpaceAfter || 0,
  });
}
function rect(x, y, w, h, fill, round = true) {
  slide.addShape(round ? pres.ShapeType.roundRect : pres.ShapeType.rect,
    { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: round ? 0.08 : undefined });
}
function circle(label, x, y, d, fill, color) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  slide.addText(String(label), { x, y, w: d, h: d, fontFace: FONT, fontSize: d >= 0.38 ? 13 : 10.5, bold: true, color,
    align: 'center', valign: 'middle', margin: 0, isTextBox: true });
}
function arrow(dir, x, y, w, h, color = RULE) {
  const type = { down: pres.ShapeType.downArrow, up: pres.ShapeType.upArrow }[dir] || pres.ShapeType.rightArrow;
  slide.addShape(type, { x, y, w, h, fill: { color }, line: { color, width: 0 } });
}
function tone(name) {
  switch (name || 'paper') {
    case 'tint': return { fill: ph.tint, head: ph.text, body: INK, em: ph.text };
    case 'solid': return { fill: ph.main, head: ph.on, body: ph.on, em: ph.on };
    case 'bad': return { fill: BADTINT, head: BAD, body: INK, em: BAD };
    case 'ok': return { fill: OKTINT, head: OK, body: INK, em: OK };
    default: return { fill: PAPER, head: ph.text, body: INK, em: ph.text };
  }
}

// ------------------------------------------------------------
// 部品。measure(b, w) で自然な高さ、draw(b, x, y, w, h) で描く。grow は余った高さを受け取れるか
// ------------------------------------------------------------
const PADX = 0.18, PADY = 0.13;
const BLOCKS = {
  // 見出し付きの枠 {head, body, tone, size, headSize}
  card: {
    grow: true,
    headH: (b, w) => b.head ? textH(b.head, b.headSize || 13, w - 2 * PADX, 1.15) + 0.06 : 0,
    measure(b, w) { return PADY + this.headH(b, w) + (b.body ? textH(b.body, b.size || 12, w - 2 * PADX) : 0) + PADY + 0.02; },
    draw(b, x, y, w, h) {
      const t = tone(b.tone);
      rect(x, y, w, h, t.fill);
      let cy = y + PADY;
      if (b.head) {
        const hh = this.headH(b, w);
        txt(b.head, { x: x + PADX, y: cy, w: w - 2 * PADX, h: hh, size: b.headSize || 13, bold: true, color: t.head, ls: 1.15, label: 'card-head' });
        cy += hh;
      }
      if (b.body) txt(b.body, { x: x + PADX, y: cy, w: w - 2 * PADX, h: y + h - cy - PADY + 0.02, size: b.size || 12, color: t.body, em: t.em, label: 'card-body' });
    },
  },
  // 番号丸＋太字見出し＋説明の並び {head, items: [[見出し, 説明], …], marks, size, bodySize, tone}
  steps: {
    grow: 'spread',
    item(b, w, it) {
      const d = 0.38, tw = w - d - 0.16;
      const hh = textH(it[0], b.size || 13.5, tw, 1.15);
      const bh = it[1] ? textH(it[1], b.bodySize || 11.5, tw, 1.25) : 0;
      return Math.max(d + 0.04, hh + (it[1] ? 0.03 + bh : 0));
    },
    headH: b => b.head ? 0.38 : 0,
    measure(b, w) { return this.headH(b) + b.items.reduce((s, it) => s + this.item(b, w, it), 0) + (b.items.length - 1) * 0.12; },
    draw(b, x, y, w, h) {
      const nat = this.measure(b, w);
      const gap = 0.12 + Math.min(0.4, Math.max(0, h - nat) / Math.max(1, b.items.length - 1));
      let cy = y;
      if (b.head) { txt(b.head, { x, y: cy, w, h: 0.32, size: 14, bold: true, color: ph.text, label: 'steps-head' }); cy += this.headH(b); }
      const d = 0.38, tw = w - d - 0.16;
      b.items.forEach((it, i) => {
        const mark = b.marks ? b.marks[i] : i + 1;
        const bad = b.badMarks && b.badMarks.includes(i);
        circle(mark, x, cy + 0.01, d, bad ? BAD : ph.main, bad ? WHITE : ph.on);
        const ih = this.item(b, w, it);
        if (it[1]) {
          const hh = textH(it[0], b.size || 13.5, tw, 1.15);
          txt(it[0], { x: x + d + 0.16, y: cy, w: tw, h: hh, size: b.size || 13.5, bold: true, ls: 1.15, label: 'steps-item' });
          txt(it[1], { x: x + d + 0.16, y: cy + hh + 0.03, w: tw, h: ih - hh - 0.03, size: b.bodySize || 11.5, color: INK2, ls: 1.25, em: ph.text, label: 'steps-body' });
        } else {
          txt(it[0], { x: x + d + 0.16, y: cy, w: tw, h: ih, size: b.size || 13.5, bold: !b.plain, valign: 'middle', em: ph.text, ls: 1.15, label: 'steps-item' });
        }
        cy += ih + gap;
      });
    },
  },
  // チェック項目 {head, items: [...], tone, size}
  checks: {
    grow: true,
    measure(b, w) {
      const size = b.size || 12;
      return 0.15 + (b.head ? 0.4 : 0) + b.items.reduce((s, it) => s + textH('□ ' + it, size, w - 0.4, 1.3) + 0.06, 0) + 0.12;
    },
    draw(b, x, y, w, h) {
      const t = tone(b.tone || 'tint'), size = b.size || 12;
      rect(x, y, w, h, t.fill);
      let cy = y + 0.15;
      if (b.head) { txt(b.head, { x: x + 0.2, y: cy, w: w - 0.4, h: 0.34, size: 13.5, bold: true, color: t.head, ls: 1.15 }); cy += 0.4; }
      txt(b.items.map(s => '□ ' + s).join('\n'), { x: x + 0.2, y: cy, w: w - 0.4, h: y + h - cy - 0.08, size, color: t.body, em: t.em,
        paraSpaceAfter: 4.3, label: 'checks', noCheck: true });
      const need = b.items.reduce((s, it) => s + textH('□ ' + it, size, w - 0.4, 1.3) + 0.06, 0);
      if (need > y + h - cy - 0.06) warnings.push(`${where} checks: ${need.toFixed(2)}in > ${(y + h - cy).toFixed(2)}in`);
    },
  },
  // プロンプト例 {caption, text, size}
  prompt: {
    grow: true,
    measure(b, w) { return 0.12 + 0.3 + textH(b.text, b.size || 11, w - 0.44, 1.25) + 0.12; },
    draw(b, x, y, w, h) {
      rect(x, y, w, h, PAPER);
      rect(x, y, 0.07, h, ph.main, false);
      txt('プロンプト例　' + (b.caption || ''), { x: x + 0.24, y: y + 0.11, w: w - 0.4, h: 0.28, size: 10.5, bold: true, color: ph.text, ls: 1.1 });
      txt(b.text, { x: x + 0.24, y: y + 0.42, w: w - 0.42, h: h - 0.5, size: b.size || 11, color: INK2, ls: 1.25, label: 'prompt' });
    },
  },
  // 記入例 {head, rows: [[項目, 例], …], labelW, body, size}
  example: {
    grow: true,
    rowsH(b, w) {
      const size = b.size || 11.5, lw = b.labelW || 0;
      return b.rows.map(([l, t]) => lw ? Math.max(textH(l, size, lw - 0.08, 1.25), textH(t, size, w - 2 * PADX - lw, 1.25))
        : textH(l + '　' + t, size, w - 2 * PADX, 1.25));
    },
    measure(b, w) {
      let h = PADY + 0.36;
      if (b.rows) { const rh = this.rowsH(b, w); h += rh.reduce((s, x) => s + x, 0) + (rh.length - 1) * 0.08; }
      if (b.body) h += textH(b.body, b.size || 11.5, w - 2 * PADX, 1.25);
      return h + PADY;
    },
    draw(b, x, y, w, h) {
      const size = b.size || 11.5, lw = b.labelW || 0;
      rect(x, y, w, h, OKTINT);
      txt(b.head || '記入例', { x: x + PADX, y: y + PADY, w: w - 2 * PADX, h: 0.3, size: 12.5, bold: true, color: OK, ls: 1.1 });
      let cy = y + PADY + 0.36;
      if (b.rows) {
        const rh = this.rowsH(b, w);
        b.rows.forEach(([l, t], i) => {
          if (lw) {
            txt(l, { x: x + PADX, y: cy, w: lw - 0.08, h: rh[i], size, bold: true, color: OK, ls: 1.25, label: 'example-label' });
            txt(t, { x: x + PADX + lw, y: cy, w: w - 2 * PADX - lw, h: rh[i], size, color: INK, ls: 1.25, label: 'example' });
          } else {
            txt(`**${l}**　${t}`, { x: x + PADX, y: cy, w: w - 2 * PADX, h: rh[i], size, color: INK, em: OK, ls: 1.25, label: 'example' });
          }
          cy += rh[i] + 0.08;
        });
      }
      if (b.body) txt(b.body, { x: x + PADX, y: cy, w: w - 2 * PADX, h: y + h - cy - PADY + 0.02, size, color: INK, em: OK, ls: 1.25, label: 'example-body' });
    },
  },
  // 表 {head: [...], rows: [[...]], colW: [...], size, align: ['l','c',…], firstBold}
  table: {
    grow: false,
    widths(b, w) { const s = b.colW.reduce((a, c) => a + c, 0); return b.colW.map(c => c * w / s); },
    rowHeights(b, w) {
      const cw = this.widths(b, w), size = b.size || 11;
      return [b.head, ...b.rows].map(r => Math.max(0.3, ...r.map((c, i) => textH(c || ' ', size, cw[i] - 0.16, 1.2) + 0.1)));
    },
    measure(b, w) { return this.rowHeights(b, w).reduce((s, x) => s + x, 0); },
    draw(b, x, y, w) {
      const cw = this.widths(b, w), size = b.size || 11, al = b.align || [];
      const cell = (c, ri, ci) => ({ text: c, options: {
        fontFace: FONT, fontSize: size, bold: ri === 0 || (b.firstBold !== false && ci === 0),
        color: ri === 0 ? ph.text : (ci === 0 ? INK : INK2), fill: { color: ri === 0 ? ph.tint : WHITE },
        align: ri === 0 ? 'center' : ({ c: 'center', r: 'right' }[al[ci]] || 'left'), valign: 'middle', margin: [3, 6, 3, 6] } });
      const rows = [b.head, ...b.rows].map((r, ri) => r.map((c, ci) => cell(c, ri, ci)));
      slide.addTable(rows, { x, y, w, colW: cw, rowH: this.rowHeights(b, w), border: { type: 'solid', pt: 0.6, color: RULE } });
    },
  },
  // 矢印でつないだ流れ {items: [...], h, size, last, aw（矢印の幅）, gap（矢印の前後の間）}
  flow: {
    grow: false,
    measure(b) { return b.h || 0.62; },
    draw(b, x, y, w, h) {
      const n = b.items.length, aw = b.aw || 0.2, gap = b.gap ?? 0.08;
      const cw = (w - (n - 1) * (aw + 2 * gap)) / n, size = b.size || 12;
      b.items.forEach((t, i) => {
        const cx = x + i * (cw + aw + 2 * gap);
        const hi = b.last && i === n - 1;
        rect(cx, y, cw, h, hi ? ph.main : ph.tint);
        txt(t, { x: cx + 0.08, y, w: cw - 0.16, h, size, bold: true, color: hi ? ph.on : ph.text, align: 'center', valign: 'middle', ls: 1.15, label: 'flow' });
        if (i < n - 1) arrow('right', cx + cw + gap, y + h / 2 - 0.13, aw, 0.26);
      });
    },
  },
  // 1行の強調 {text, tone, size}
  banner: {
    grow: false,
    measure(b, w) { return Math.max(0.46, textH(b.text, b.size || 13, w - 0.4, 1.2) + 0.2); },
    draw(b, x, y, w, h) {
      const t = tone(b.tone || 'solid');
      rect(x, y, w, h, t.fill);
      txt(b.text, { x: x + 0.2, y, w: w - 0.4, h, size: b.size || 13, bold: true, color: t.head, em: t.em, valign: 'middle', ls: 1.2, label: 'banner' });
    },
  },
  // 小見出し {text, size}
  label: {
    grow: false,
    measure(b, w) { return textH(b.text, b.size || 14, w, 1.15) + 0.06; },
    draw(b, x, y, w, h) { txt(b.text, { x, y, w, h, size: b.size || 14, bold: true, color: { bad: BAD, ok: OK }[b.color] || ph.text, ls: 1.15, label: 'label' }); },
  },
  // 平面のゾーニング図 {caption, uw, uh, zones: [[x, y, 幅, 高さ, 名前, 'hi'|'bad'], …], entry: [x, 名前], height, size}
  //   座標は uw × uh の升目で、左上が原点。entry は下の辺の入口（矢印）
  zones: {
    grow: true,
    measure(b) { return b.height || 2.4; },
    draw(b, x, y, w, h) {
      const capH = b.caption ? 0.34 : 0, entH = b.entry ? 0.34 : 0;
      if (b.caption) txt(b.caption, { x, y, w, h: 0.3, size: 12.5, bold: true, color: ph.text, ls: 1.1 });
      const s = Math.min(w / b.uw, (h - capH - entH) / b.uh);
      const ox = x + (w - b.uw * s) / 2, oy = y + capH;
      rect(ox, oy, b.uw * s, b.uh * s, PAPER);
      b.zones.forEach(([zx, zy, zw, zh, label, mark]) => {
        const fill = { hi: ph.main, bad: BADTINT }[mark] || ph.tint, color = { hi: ph.on, bad: BAD }[mark] || ph.text;
        rect(ox + zx * s, oy + zy * s, zw * s, zh * s, fill);
        txt(label, { x: ox + zx * s + 0.03, y: oy + zy * s, w: zw * s - 0.06, h: zh * s, size: b.size || 10, bold: true, color,
          align: 'center', valign: 'middle', ls: 1.1, label: 'zone' });
      });
      if (b.entry) {
        const ex = ox + b.entry[0] * s, ey = oy + b.uh * s + 0.04;
        arrow('up', ex - 0.11, ey, 0.22, 0.28, ph.main);
        txt(b.entry[1], { x: ex + 0.16, y: ey, w: 1.2, h: 0.28, size: 10.5, bold: true, color: ph.text, valign: 'middle', ls: 1.1 });
      }
    },
  },
  // 大きな数字のタイル {items: [[数字, 説明], …], h}
  stat: {
    grow: false,
    measure(b) { return b.h || 1.15; },
    draw(b, x, y, w, h) {
      const n = b.items.length, gap = 0.15, tw = (w - gap * (n - 1)) / n;
      b.items.forEach(([v, l], i) => {
        const tx = x + i * (tw + gap);
        rect(tx, y, tw, h, PAPER);
        txt(v, { x: tx, y: y + 0.1, w: tw, h: h * 0.55, size: 32, bold: true, color: ph.text, align: 'center', valign: 'middle', ls: 1.0, noCheck: true });
        txt(l, { x: tx + 0.1, y: y + h * 0.62, w: tw - 0.2, h: h * 0.3, size: 12.5, color: INK2, align: 'center', ls: 1.1, label: 'stat' });
      });
    },
  },
  // 地の文 {text, size}
  text: {
    grow: false,
    measure(b, w) { return textH(b.text, b.size || 12, w, 1.3); },
    draw(b, x, y, w, h) { txt(b.text, { x, y, w, h, size: b.size || 12, color: INK2, em: ph.text, label: 'text' }); },
  },
};

function blockOf(b) {
  const def = BLOCKS[b.type];
  if (!def) throw new Error(`${where}: 部品 "${b.type}" はありません`);
  return def;
}
const measure = (b, w) => blockOf(b).measure(b, w);
const canGrow = b => b.grow !== undefined ? b.grow : !!blockOf(b).grow;

// 縦に積む。余った高さは最後の伸ばせる部品に渡す
function stack(blocks, x, y, w, h) {
  const hs = blocks.map(b => measure(b, w));
  const nat = hs.reduce((s, v) => s + v, 0) + (blocks.length - 1) * BLOCKGAP;
  if (nat > h + 0.03) warnings.push(`${where} column: ${nat.toFixed(2)}in > ${h.toFixed(2)}in`);
  const gi = blocks.map(canGrow).lastIndexOf(true);
  if (gi >= 0 && h > nat) hs[gi] += h - nat;
  let cy = y;
  blocks.forEach((b, i) => { blockOf(b).draw(b, x, cy, w, hs[i]); cy += hs[i] + BLOCKGAP; });
}
function colWidths(cols, w, gap) {
  const fixed = cols.reduce((s, c) => s + (c.w || 0), 0), free = cols.filter(c => !c.w).length;
  const rest = (w - fixed - gap * (cols.length - 1)) / Math.max(1, free);
  return cols.map(c => c.w || rest);
}
function rowMeasure(r) {
  if (!r.cols) return measure(r, CW);
  const gap = r.gap ?? COLGAP, ws = colWidths(r.cols, CW, gap);
  return Math.max(...r.cols.map((c, i) => c.blocks.reduce((s, b) => s + measure(b, ws[i]), 0) + (c.blocks.length - 1) * BLOCKGAP));
}
const rowGrows = r => r.cols ? r.cols.some(c => c.blocks.some(canGrow)) : canGrow(r);

function layoutBody(body) {
  const hs = body.map(rowMeasure);
  const nat = hs.reduce((s, v) => s + v, 0) + (body.length - 1) * ROWGAP;
  const avail = Y1 - Y0;
  if (nat > avail + 0.03) warnings.push(`${where} body: ${nat.toFixed(2)}in > ${avail.toFixed(2)}in`);
  const gi = body.map(rowGrows).lastIndexOf(true);
  if (gi >= 0 && avail > nat) hs[gi] += avail - nat;
  let y = Y0;
  body.forEach((r, i) => {
    if (r.cols) {
      const gap = r.gap ?? COLGAP, ws = colWidths(r.cols, CW, gap);
      let x = X0;
      r.cols.forEach((c, ci) => { stack(c.blocks, x, y, ws[ci], hs[i]); x += ws[ci] + gap; });
    } else {
      blockOf(r).draw(r, X0, y, CW, hs[i]);
    }
    y += hs[i] + ROWGAP;
  });
}

// ------------------------------------------------------------
// 1回分の組み立て
// ------------------------------------------------------------
function loadAudio(nn) {
  const f = path.join(HERE, 'audio', nn, 'audio.json');
  return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : {};
}
function fullDate(d) { return /^\d{4}年/.test(d) ? d : '2026年' + d; }

function buildSession(nn) {
  const spec = JSON.parse(fs.readFileSync(path.join(HERE, 'content', nn + '.json'), 'utf8'));
  const audio = loadAudio(nn);
  const totalSec = Object.values(audio).reduce((s, a) => s + a.sec + 1.3, 0);
  ph = PHASES[spec.phase];
  warnings = [];
  pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.author = '河合 勝彦';
  pres.title = `基礎演習Ⅱ 第${spec.session}回 ${spec.title}　ワークシートの解説`;
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  const FOOT = `基礎演習Ⅱ　2026年度後期　第${spec.session}回 ワークシートの解説`;
  pres.defineSlideMaster({ title: 'TITLE', background: { color: ph.main } });
  pres.defineSlideMaster({
    title: 'CONTENT', background: { color: WHITE },
    objects: [
      { placeholder: { options: { name: 'kicker', type: 'body', x: 0.5, y: 0.32, w: 8.8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: ph.text, align: 'left', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.6, w: 9.0, h: 0.62, fontFace: FONT, fontSize: 24, bold: true, color: INK, align: 'left', margin: 0, valign: 'top' }, text: '' } },
      { text: { text: FOOT, options: { x: 0.5, y: 5.22, w: 6, h: 0.25, fontFace: FONT, fontSize: 9, color: INK3, margin: 0 } } },
    ],
    slideNumber: { x: 8.7, y: 5.22, w: 0.5, h: 0.25, fontFace: FONT, fontSize: 9, color: INK3, align: 'right' },
  });

  spec.slides.forEach((s, i) => {
    const id = i + 1;
    where = `${nn}/${pad2(id)}`;
    if (s.type === 'title') {
      slide = pres.addSlide({ masterName: 'TITLE' });
      txt('名古屋市立大学 経済学部　2026年度後期　基礎演習Ⅱ', { x: 0.7, y: 0.6, w: 8.6, h: 0.32, size: 13, color: ph.sub });
      txt('ワークシートの解説', { x: 0.7, y: 1.0, w: 8.6, h: 0.5, size: 20, bold: true, color: ph.on });
      const head = `第${spec.session}回　${spec.title}`;
      txt(head, { x: 0.7, y: 1.5, w: 8.6, h: 0.85, size: Math.min(40, Math.floor(8.4 * 72 / emWidth(head))), bold: true, color: ph.on, ls: 1.1 });
      txt([fullDate(spec.date), ph.name, spec.tag].filter(Boolean).join('　｜　'), { x: 0.7, y: 2.45, w: 8.6, h: 0.4, size: 15, color: ph.sub });
      for (let k = 1; k <= 15; k++) {
        const cur = k === spec.session;
        circle(k, 0.7 + (k - 1) * 0.42, 3.35, 0.32, cur ? ph.on : ph.dot, cur ? ph.main : ph.on);
      }
      txt('河合 勝彦（名古屋市立大学 経済学部）', { x: 0.7, y: 4.45, w: 6, h: 0.3, size: 13, color: ph.on });
      txt(`音声付きスライド（約${Math.max(1, Math.round(totalSec / 60)) || '–'}分）　｜　ワークシートを手元に置いて見てください`,
        { x: 0.7, y: 4.8, w: 8.6, h: 0.3, size: 11, color: ph.sub });
    } else {
      slide = pres.addSlide({ masterName: 'CONTENT' });
      slide.addText(s.kicker || `第${spec.session}回　${spec.title}`, { placeholder: 'kicker' });
      if (lineCount(s.heading, 24, 9.0) > 1) warnings.push(`${where} heading が2行になる「${s.heading}」`);
      slide.addText(s.heading, { placeholder: 'title' });
      layoutBody(s.body);
    }
    slide.addNotes(s.narration);
    const mp3 = path.join(HERE, 'audio', nn, pad2(id) + '.mp3');
    if (fs.existsSync(mp3)) slide.addMedia({ type: 'audio', path: mp3, x: 9.32, y: 5.0, w: 0.42, h: 0.42 });
  });

  const out = path.join(HERE, `worksheet-${nn}-slides.pptx`);
  return pres.writeFile({ fileName: out }).then(() => {
    console.log(`wrote worksheet-${nn}-slides.pptx  ${spec.slides.length} slides  ${Object.keys(audio).length ? (totalSec / 60).toFixed(1) + ' min' : 'no audio'}`);
    warnings.forEach(w => console.log('  fit: ' + w));
  });
}

async function main() {
  const all = fs.readdirSync(path.join(HERE, 'content')).filter(f => /^\d\d\.json$/.test(f)).map(f => f.slice(0, 2)).sort();
  const wanted = process.argv.slice(2).map(a => pad2(parseInt(a, 10)));
  for (const nn of wanted.length ? wanted : all) await buildSession(nn);
}
main().catch(e => { console.error(e); process.exit(1); });
