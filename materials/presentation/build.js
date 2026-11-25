// 産学連携プロジェクト　提案発表のスライドテンプレート（発表時間5分）を作る
//   node build.js                    5つの型のテンプレートを作る（presentation-template-5min-<型>.pptx）
//   node build.js story compare      指定した型だけ作る（basic・situation・story・conclusion・compare）
//   node build.js --examples         作例を作る（examples/presentation-example-<題材>-<型>.pptx）
//   node build.js --examples family  指定した題材・型の作例だけ作る（題材は examples/<題材>.js）
//   必要なもの: npm install pptxgenjs。作例の図は examples/make_images.py で先に作っておく
//   作例の音声（発表の音声、教員の解説）は examples/make_audio.py で先に作っておく。音声がまだない題材は、音声なしで組む
//
// 入力枠は PowerPoint の「プレースホルダー」として、スライドのレイアウトに置いてある。
// 灰色の案内文はレイアウト側の文字なので、何も入力しなかった枠はスライドショーにも印刷にも出ない。
// pptxgenjs はプレースホルダーの余白・縦位置・自動調整や図の枠を書き出せないので、
// 書き出した pptx を開き直して、プレースホルダーとテーマ（色・フォント）、ノートの段落を書き直している（finish()）。
//
// 型（VARIANTS）は、共通のレイアウトの部品（patterns()）に、章の名前と案内文を渡して組み立てる。
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const crypto = require('crypto');

const EXAMPLES = process.argv.includes('--examples');

// ------------------------------------------------------------
// テーマ（授業サイト assets/css/style.css の「実践プロジェクト」の色に合わせる）
// スライドの色はすべてテーマの色で指定してあるので、PowerPoint の［デザイン］→［バリエーション］→［配色］で一括して変えられる
// ------------------------------------------------------------
const FONT = 'Yu Gothic';
const THEME = {
  name: 'Sangaku Teian',
  colors: {
    dk1: '1D2330', lt1: 'FFFFFF', dk2: '4A5261', lt2: 'ECEEE8',
    accent1: 'D68A17', accent2: '8A570E', accent3: 'F6E4C3', accent4: '2747C9', accent5: '2E7D32', accent6: 'B3261E',
    hlink: '1F3BA8', folHlink: '4A5261',
  },
};
const INK = 'tx1', WHITE = 'bg1', INK2 = 'tx2', PAPER = 'bg2', AMBER = 'accent1', AMBERTEXT = 'accent2', TINT = 'accent3';

const RADIUS = 0.08;                       // 角丸（インチ）
const X0 = 0.5, CW = 9.0, TOP = 1.62;      // 本文の領域（10 x 5.625 in）
const FOOTER = '産学連携プロジェクト　提案発表';
const TOTAL = 300;                         // 発表時間（秒）
const SEC_USAGE = 'はじめに（非表示）', SEC_TALK = '発表（5分）', SEC_APPENDIX = '補足（質疑用）';

// ------------------------------------------------------------
// レイアウトに置くもの
// ------------------------------------------------------------
// 塗りの図形（文字を載せることもできる）
function box(x, y, w, h, fill, o = {}) {
  const shape = o.shape || 'roundRect';
  return { text: { text: o.text || '', options: {
    shape, rectRadius: shape === 'roundRect' ? RADIUS : undefined, x, y, w, h, fill: { color: fill },
    fontSize: o.size || 12, bold: !!o.bold, color: o.color || INK, align: 'center', valign: 'middle', margin: 0,
    objectName: o.name || '図形',
  } } };
}
const circle = (n, x, y, d = 0.36) => box(x, y, d, d, AMBER, { shape: 'ellipse', text: String(n), size: d > 0.36 ? 13 : 12, bold: true, name: '番号 ' + n });
// 動かない文字（小見出しなど）
function label(text, x, y, w, h, o = {}) {
  return { text: { text, options: {
    x, y, w, h, fontSize: o.size || 12, bold: o.bold !== false, color: o.color || AMBERTEXT,
    align: o.align || 'left', valign: o.valign || 'middle', margin: 0, objectName: '見出し ' + text,
  } } };
}
// 入力枠（プレースホルダー）。type は title・body・pic
//   size・bold・color・align（l/ctr/r）・anchor（t/ctr/b）・ls（行間）・bullet・prompt（案内文）・promptColor・geom（図の枠の形）
function ph(key, type, x, y, w, h, o = {}) {
  return { ph: Object.assign({ key, type, x, y, w, h }, o) };
}

// ------------------------------------------------------------
// レイアウトの部品。どれも「レイアウト名、章の番号、章の名前、案内文など」を受け取る
// ------------------------------------------------------------
function patterns(frame, define) {
  const P = {};

  // 表紙（暗い地）：キャッチコピー、コンセプトの一文、グループ、提案先、画像
  P.cover = (c = {}) => define('表紙', INK, [
    label(FOOTER, 0.6, 0.6, 4.9, 0.3, { color: AMBER }),
    ph('title', 'title', 0.6, 1.15, 4.9, 1.75, { size: 32, bold: true, color: WHITE, anchor: 'b', prompt: c.title || 'キャッチコピー', promptColor: PAPER }),
    ph('concept', 'body', 0.6, 3.05, 4.9, 0.85, { color: PAPER, prompt: c.concept || '誰に・何を・どのように（コンセプトを一文で）', promptColor: PAPER }),
    ph('group', 'body', 0.6, 4.3, 4.9, 0.34, { bold: true, color: WHITE, anchor: 'ctr', prompt: 'グループ名｜メンバーの氏名', promptColor: PAPER }),
    ph('meta', 'body', 0.6, 4.66, 4.9, 0.3, { size: 12, color: PAPER, anchor: 'ctr', prompt: '提案先の企業名　御中｜発表日', promptColor: PAPER }),
    box(5.9, 0.6, 3.5, 4.4, AMBER, { name: '画像の台' }),
    ph('image', 'pic', 5.9, 0.6, 3.5, 4.4, { geom: 'roundRect', prompt: c.image || '提案を象徴する画像', promptColor: INK }),
  ], false);

  // 見出し付きの3行＋数字の枠
  P.rowsStat = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    c.rows.forEach(([t, p], i) => {
      const y = TOP + i * 1.15;
      o.push(box(X0, y, 5.3, 1.0, PAPER, { name: '枠 ' + t }));
      o.push(label(t, 0.7, y, 1.25, 1.0));
      o.push(ph('row' + (i + 1), 'body', 2.0, y + 0.08, 3.65, 0.84, { anchor: 'ctr', prompt: p }));
    });
    o.push(box(6.1, TOP, 3.4, 3.3, TINT, { name: '枠 数字' }));
    o.push(label(c.statHead, 6.3, TOP + 0.16, 3.0, 0.3));
    o.push(ph('stat', 'body', 6.3, 2.15, 3.0, 1.05, { size: 54, bold: true, anchor: 'ctr', prompt: '00%' }));
    o.push(ph('statNote', 'body', 6.3, 3.25, 3.0, 0.95, { prompt: 'この数字が示していること' }));
    o.push(ph('source', 'body', 6.3, 4.3, 3.0, 0.45, { size: 10, color: INK2, anchor: 'b', prompt: '出典：資料名、発行者、年' }));
    return define(name, WHITE, o);
  };

  // 人物の紹介：丸い画像、名前、属性、見出し付きの3行、本人の言葉
  P.persona = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(box(X0, TOP, 3.3, 3.3, PAPER, { name: '枠 人物' }));
    o.push(box(1.4, 1.8, 1.5, 1.5, TINT, { shape: 'ellipse', name: '画像の台' }));
    o.push(ph('image', 'pic', 1.4, 1.8, 1.5, 1.5, { geom: 'ellipse', size: 10, prompt: '人物の画像' }));
    o.push(ph('name', 'body', 0.7, 3.42, 2.9, 0.4, { size: 16, bold: true, align: 'ctr', anchor: 'ctr', prompt: c.name || 'ペルソナの名前（年齢）' }));
    o.push(ph('profile', 'body', 0.7, 3.86, 2.9, 0.9, { size: 12, color: INK2, align: 'ctr', prompt: '職業・家族構成・住んでいる地域' }));
    c.rows.forEach(([t, p], i) => {
      const y = TOP + i * 0.84;
      o.push(box(4.1, y + 0.19, 1.75, 0.34, TINT, { text: t, size: 11, bold: true, color: AMBERTEXT, name: '見出し ' + t }));
      o.push(ph('row' + (i + 1), 'body', 6.0, y, 3.5, 0.72, { anchor: 'ctr', prompt: p }));
    });
    o.push(box(4.1, 4.17, 5.4, 0.75, INK, { name: '枠 本人の言葉' }));
    o.push(ph('voice', 'body', 4.3, 4.17, 5.0, 0.75, { bold: true, color: WHITE, anchor: 'ctr', prompt: c.voice, promptColor: PAPER }));
    return define(name, WHITE, o);
  };

  // 場面の紹介：大きな画像、見出し付きの3行、本人の言葉
  P.sceneRows = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(box(X0, TOP, 4.0, 2.95, PAPER, { name: '図の台' }));
    o.push(ph('image', 'pic', X0, TOP, 4.0, 2.95, { geom: 'roundRect', prompt: c.image }));
    o.push(ph('caption', 'body', X0, 4.62, 4.0, 0.3, { size: 10, color: INK2, anchor: 'ctr', prompt: '画像の説明（生成画像はその旨を明記）' }));
    c.rows.forEach(([t, p], i) => {
      const y = TOP + i * 0.84;
      o.push(box(4.8, y + 0.19, 1.5, 0.34, TINT, { text: t, size: 11, bold: true, color: AMBERTEXT, name: '見出し ' + t }));
      o.push(ph('row' + (i + 1), 'body', 6.45, y, 3.05, 0.72, { anchor: 'ctr', prompt: p }));
    });
    o.push(box(4.8, 4.17, 4.7, 0.75, INK, { name: '枠 本人の言葉' }));
    o.push(ph('voice', 'body', 5.0, 4.17, 4.3, 0.75, { bold: true, color: WHITE, anchor: 'ctr', prompt: c.voice, promptColor: PAPER }));
    return define(name, WHITE, o);
  };

  // 大きな一文（暗い帯）＋3つの枠
  P.copy3 = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(box(X0, TOP, CW, 1.45, INK, { name: '枠 大きな一文' }));
    o.push(ph('copy', 'body', 0.8, TOP, 8.4, 1.45, { size: 32, bold: true, color: WHITE, align: 'ctr', anchor: 'ctr', prompt: c.copy, promptColor: PAPER }));
    const w = (CW - 2 * 0.25) / 3;
    c.cols.forEach(([t, p], i) => {
      const x = X0 + i * (w + 0.25);
      o.push(box(x, 3.27, w, 1.65, PAPER, { name: '枠 ' + t }));
      o.push(label(t, x + 0.2, 3.4, w - 0.4, 0.3));
      o.push(ph('col' + (i + 1), 'body', x + 0.2, 3.76, w - 0.4, 1.04, { prompt: p }));
    });
    return define(name, WHITE, o);
  };

  // 大きな図＋番号付きの3点
  P.imagePoints = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(box(X0, TOP, 5.5, 2.95, PAPER, { name: '図の台' }));
    o.push(ph('image', 'pic', X0, TOP, 5.5, 2.95, { geom: 'roundRect', prompt: 'ゾーニング図・平面図・イメージ画像' }));
    o.push(ph('caption', 'body', X0, 4.62, 5.5, 0.3, { size: 10, color: INK2, anchor: 'ctr', prompt: '図の説明（生成画像には「AIによる生成イメージ」と明記）' }));
    for (let i = 0; i < 3; i++) {
      const y = TOP + i * 1.1;
      o.push(circle(i + 1, 6.3, y));
      o.push(ph('head' + (i + 1), 'body', 6.82, y, 2.68, 0.36, { bold: true, anchor: 'ctr', prompt: '変えるところ' + (i + 1) }));
      o.push(ph('desc' + (i + 1), 'body', 6.82, y + 0.4, 2.68, 0.62, { prompt: c.desc || 'ねらいと、そう考えた根拠' }));
    }
    return define(name, WHITE, o);
  };

  // 4つの場面の流れ＋「これまで→これから」
  P.flow4 = (name, n, kicker, c = {}) => {
    const o = frame(n, kicker);
    const w = 2.1, gap = 0.2;
    o.push(box(X0 + w / 2, 1.805, 3 * (w + gap), 0.03, TINT, { shape: 'rect', name: '流れの線' }));
    ['来店前', '入店', '店内', '退店後'].forEach((t, i) => {
      const x = X0 + i * (w + gap);
      o.push(circle(i + 1, x + w / 2 - 0.2, TOP, 0.4));
      o.push(ph('step' + (i + 1), 'body', x, 2.12, w, 0.4, { bold: true, align: 'ctr', anchor: 'ctr', prompt: '場面（例：' + t + '）' }));
      o.push(box(x, 2.6, w, 1.5, PAPER, { name: '枠 場面' + (i + 1) }));
      o.push(ph('act' + (i + 1), 'body', x + 0.14, 2.7, w - 0.28, 1.3, { prompt: 'お客さまの行動と、店の対応' }));
    });
    o.push(box(X0, 4.27, CW, 0.65, TINT, { name: '枠 変化' }));
    o.push(ph('before', 'body', 0.7, 4.27, 3.7, 0.65, { anchor: 'ctr', prompt: 'これまで：お客さまの不安・面倒' }));
    o.push(box(4.55, 4.47, 0.4, 0.25, AMBER, { shape: 'rightArrow', name: '矢印' }));
    o.push(ph('after', 'body', 5.1, 4.27, 4.2, 0.65, { bold: true, anchor: 'ctr', prompt: 'これから：体験がこう変わる' }));
    return define(name, WHITE, o);
  };

  // 効果（数字2つと根拠）＋実現の道筋（3段階と費用）
  P.effectPlan = (name, n, kicker) => {
    const o = frame(n, kicker);
    o.push(label('期待される効果', X0, TOP, 4.2, 0.28));
    [0, 1].forEach(i => {
      const x = X0 + i * 2.2;
      o.push(box(x, 1.98, 2.0, 1.4, TINT, { name: '枠 効果' + (i + 1) }));
      o.push(ph('num' + (i + 1), 'body', x + 0.15, 2.04, 1.7, 0.78, { size: 36, bold: true, anchor: 'ctr', prompt: '+00%' }));
      o.push(ph('numLabel' + (i + 1), 'body', x + 0.15, 2.84, 1.7, 0.48, { size: 12, prompt: '指標の名前' }));
    });
    o.push(box(X0, 3.53, 4.2, 1.39, PAPER, { name: '枠 根拠' }));
    o.push(label('根拠', 0.7, 3.62, 1.0, 0.28));
    o.push(ph('basis', 'body', 0.7, 3.93, 3.8, 0.9, { prompt: 'アンケート・事例・試算の前提と出典' }));
    o.push(label('実現の道筋', 5.0, TOP, 4.5, 0.28));
    ['すぐに', '半年後', '1年後'].forEach((t, i) => {
      const y = 1.98 + i * 0.8;
      o.push(box(5.0, y, 1.2, 0.7, AMBER, { name: '枠 時期' + (i + 1) }));
      o.push(ph('when' + (i + 1), 'body', 5.0, y, 1.2, 0.7, { bold: true, align: 'ctr', anchor: 'ctr', prompt: t, promptColor: INK }));
      o.push(box(6.3, y, 3.2, 0.7, PAPER, { name: '枠 施策' + (i + 1) }));
      o.push(ph('plan' + (i + 1), 'body', 6.45, y, 2.9, 0.7, { anchor: 'ctr', prompt: 'この時期に実施すること' }));
    });
    o.push(label('費用の規模', 5.0, 4.38, 1.2, 0.54));
    o.push(ph('cost', 'body', 6.3, 4.38, 3.2, 0.54, { anchor: 'ctr', prompt: '小・中・大と、主な内訳' }));
    return define(name, WHITE, o);
  };

  // 4つの列（最後の列だけ暗い地で強調）。AI活用のスライド
  P.cols4 = (name, n, kicker) => {
    const o = frame(n, kicker);
    const w = 2.1, gap = 0.2;
    [['AIを使った場面', '発想、調査、画像の生成、批判的な検討など'], ['採用したもの', 'AIの出力から取り入れたこと'],
     ['捨てたものと理由', '採用しなかった出力と、その理由'], ['自分たちの判断', '最後に自分たちで決めたことと、判断の基準']].forEach(([t, p], i) => {
      const x = X0 + i * (w + gap), dark = i === 3;
      o.push(box(x, TOP, w, 3.3, dark ? INK : PAPER, { name: '枠 ' + t }));
      o.push(label(t, x + 0.15, TOP + 0.12, w - 0.3, 0.52, { color: dark ? AMBER : AMBERTEXT }));
      o.push(ph('col' + (i + 1), 'body', x + 0.15, TOP + 0.72, w - 0.3, 2.44, { color: dark ? WHITE : INK, prompt: p, promptColor: dark ? PAPER : INK2 }));
    });
    return define(name, WHITE, o);
  };

  // 「こんなとき／こうしたい／そうすれば」の3行（暗い地）＋3つの面
  P.statement = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(box(X0, TOP, 5.3, 3.3, INK, { name: '枠 したいこと' }));
    c.lines.forEach(([t, p], i) => {
      const y = TOP + 0.2 + i * 1.0;
      o.push(label(t, 0.75, y, 4.8, 0.3, { color: AMBER }));
      o.push(ph('line' + (i + 1), 'body', 0.75, y + 0.32, 4.8, 0.58, { size: 16, bold: true, color: WHITE, prompt: p, promptColor: PAPER }));
    });
    c.sides.forEach(([t, p], i) => {
      const y = TOP + i * 1.15;
      o.push(box(6.1, y, 3.4, 1.0, PAPER, { name: '枠 ' + t }));
      o.push(label(t, 6.3, y + 0.1, 3.0, 0.26));
      o.push(ph('side' + (i + 1), 'body', 6.3, y + 0.4, 3.0, 0.52, { prompt: p }));
    });
    return define(name, WHITE, o);
  };

  // いまの選択肢（3行：選んでいるもの｜足りないところ）＋踏み出せない理由（2つ、暗い地）
  P.alternatives = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    o.push(label('いま選んでいるもの', X0, TOP, 1.9, 0.28));
    o.push(label('足りないところ', 2.55, TOP, 3.35, 0.28));
    c.alts.forEach((t, i) => {
      const y = 1.98 + i * 0.99;
      o.push(box(X0, y, 1.9, 0.9, TINT, { name: '枠 選択肢' + (i + 1) }));
      o.push(ph('alt' + (i + 1), 'body', 0.62, y, 1.66, 0.9, { bold: true, anchor: 'ctr', prompt: '例：' + t }));
      o.push(box(2.55, y, 3.35, 0.9, PAPER, { name: '枠 足りないところ' + (i + 1) }));
      o.push(ph('lack' + (i + 1), 'body', 2.7, y, 3.05, 0.9, { anchor: 'ctr', prompt: 'それでは満たされないこと' }));
    });
    o.push(label('踏み出せない理由', 6.2, TOP, 3.3, 0.28));
    c.walls.forEach(([t, p], i) => {
      const y = 1.98 + i * 1.5;
      o.push(box(6.2, y, 3.3, 1.38, INK, { name: '枠 ' + t }));
      o.push(label(t, 6.4, y + 0.12, 2.9, 0.28, { color: AMBER }));
      o.push(ph('wall' + (i + 1), 'body', 6.4, y + 0.46, 2.9, 0.8, { color: WHITE, prompt: p, promptColor: PAPER }));
    });
    return define(name, WHITE, o);
  };

  // 3つの場面（番号、場面の名前、画像、説明）
  P.storyboard3 = (name, n, kicker, c) => {
    const o = frame(n, kicker);
    const w = (CW - 2 * 0.25) / 3;
    c.scenes.forEach((t, i) => {
      const x = X0 + i * (w + 0.25);
      o.push(circle(i + 1, x, TOP));
      o.push(ph('scene' + (i + 1), 'body', x + 0.46, TOP, w - 0.46, 0.36, { bold: true, anchor: 'ctr', prompt: '場面（例：' + t + '）' }));
      o.push(box(x, 2.08, w, 1.6, PAPER, { name: '図の台 ' + (i + 1) }));
      o.push(ph('image' + (i + 1), 'pic', x, 2.08, w, 1.6, { geom: 'roundRect', size: 12, prompt: 'この場面の画像・図' }));
      o.push(ph('text' + (i + 1), 'body', x, 3.78, w, 0.8, { prompt: c.text }));
    });
    o.push(ph('caption', 'body', X0, 4.64, CW, 0.28, { size: 10, color: INK2, anchor: 'ctr', prompt: '画像の説明（生成画像には「AIによる生成イメージ」と明記）' }));
    return define(name, WHITE, o);
  };

  // 3つの案を、3つの基準で比べる表（左端の列＝選んだ案を強調）＋選んだ理由
  P.compare3 = (name, n, kicker) => {
    const o = frame(n, kicker);
    const cx = i => X0 + 2.0 + 0.13 + i * (2.2 + 0.13), ry = j => 2.27 + j * 0.72;
    o.push(box(cx(0) - 0.06, TOP - 0.06, 2.32, 2.83, TINT, { name: '枠 選んだ案' }));
    ['A：選んだ案', 'B：比べた案', 'C：比べた案'].forEach((t, i) => {
      o.push(box(cx(i), TOP, 2.2, 0.55, i === 0 ? AMBER : PAPER, { name: '枠 案' + (i + 1) }));
      o.push(ph('opt' + (i + 1), 'body', cx(i) + 0.08, TOP, 2.04, 0.55, { bold: true, align: 'ctr', anchor: 'ctr', prompt: t, promptColor: i === 0 ? INK : INK2 }));
    });
    ['顧客に合うか', '実現しやすいか', '課題に効くか'].forEach((t, j) => {
      o.push(ph('crit' + (j + 1), 'body', X0, ry(j), 2.0, 0.62, { bold: true, color: AMBERTEXT, anchor: 'ctr', prompt: '例：' + t }));
      for (let i = 0; i < 3; i++) {
        if (i > 0) o.push(box(cx(i), ry(j), 2.2, 0.62, PAPER, { name: `枠 評価${j + 1}-${i + 1}` }));
        o.push(ph(`cell${j + 1}${i + 1}`, 'body', cx(i) + 0.1, ry(j), 2.0, 0.62, { align: 'ctr', anchor: 'ctr', prompt: '評価（◎○△）と理由' }));
      }
    });
    o.push(box(X0, 4.43, CW, 0.49, INK, { name: '枠 選んだ理由' }));
    o.push(ph('reason', 'body', 0.7, 4.43, 8.6, 0.49, { bold: true, color: WHITE, anchor: 'ctr', prompt: 'Aを選んだ理由を、一文で', promptColor: PAPER }));
    return define(name, WHITE, o);
  };

  // まとめ（暗い地）：大きな一文＋番号付きの3点
  P.closing = (name, c = {}) => {
    const o = [label(c.kicker || name, 0.6, 0.6, 3, 0.3, { color: AMBER })];
    o.push(ph('title', 'title', 0.6, 1.0, 8.8, 1.1, { size: 36, bold: true, color: WHITE, anchor: 'ctr', prompt: c.title || 'キャッチコピーをもう一度', promptColor: PAPER }));
    for (let i = 0; i < 3; i++) {
      const y = 2.45 + i * 0.68;
      o.push(circle(i + 1, 0.6, y + 0.06, 0.4));
      o.push(ph('point' + (i + 1), 'body', 1.2, y, 8.2, 0.52, { size: 16, color: WHITE, anchor: 'ctr', prompt: (c.points || [])[i] || '覚えて帰ってほしいこと' + (i + 1), promptColor: PAPER }));
    }
    o.push(ph('group', 'body', 0.6, 4.66, 8.8, 0.3, { size: 12, color: PAPER, anchor: 'ctr', prompt: 'グループ名｜メンバーの氏名', promptColor: PAPER }));
    return define(name, INK, o, false);
  };

  // 補足のレイアウト（どの型にも入れる）
  P.appendix = () => {
    define('参考・出典', WHITE, frame(null, '補足').concat([
      ph('body', 'body', X0, TOP, CW, 3.3, { bullet: true, prompt: '著者・発行者「資料名」発行年、URL（参照日）' })]));
    define('補足（見出しと本文）', WHITE, frame(null, '補足').concat([
      ph('body', 'body', X0, TOP, CW, 3.3, { bullet: true, prompt: '本文（箇条書き）' })]));
    define('見出しのみ', WHITE, frame(null, null));
  };
  return P;
}

// ------------------------------------------------------------
// ノート（話すこと、書き方）
// ------------------------------------------------------------
const NOTES = {
  cover: [
    '話すこと：グループ名と、提案を一言で（キャッチコピー）。「これから○○を提案します」と言い切って始める。',
    '書き方：キャッチコピーは20字以内。その下に「誰に・何を・どのように」の一文を置く。',
    '画像：提案を象徴する1枚（第8回の店舗イメージなど）。生成画像なら「AIによる生成イメージ」と入れる。'],
  persona: [
    '話すこと：提案の相手を1人に絞って紹介する。最後に本人の言葉を読み上げ、次のスライドにつなげる。',
    '書き方：第6回で作ったペルソナから、提案に関係する点だけを残す。属性の羅列にしない。',
    '画像：人物のイメージ。実在の人物の写真は使わない。生成画像なら「AIによる生成イメージ」と入れる。',
    '評価の観点：ペルソナとコンセプトに一貫性があるか。'],
  space: [
    '話すこと：図を指しながら、変えるところを3つまで。それぞれ「何を」「なぜ（誰のどの不安に効くか）」をセットで話す。',
    '書き方：図を大きく、文字は少なく。説明は口頭でする。第8回のゾーニング図やイメージを使う。',
    '画像：生成画像には「AIによる生成イメージ」と明記する。実在のメーカーのロゴに似たものが写っていないか確認する。',
    '評価の観点：空間と動線の提案に根拠があるか。'],
  flow: [
    '話すこと：お客さまが来店前から退店後までにどう動き、どこで不安が消えるかを順に話す。最後に「これまで→これから」でまとめる。',
    '書き方：場面は4つまで。場面の名前は提案に合わせて書き換える。第9回のジャーニーマップを使う。',
    '1つの場面は1〜2文で。細かい説明は質疑に回す。'],
  effect: [
    '話すこと：期待できる効果を数字2つまでで示し、根拠を一言添える。続けて、すぐにできることから順に3段階で実現の道筋を話す。',
    '書き方：数字には根拠（アンケート、他社の事例、試算の前提）と出典を。根拠のない数字は載せない。',
    '費用：正確な金額でなくてよい。「小・中・大」の規模と主な内訳を示す。',
    '評価の観点：実現性と費用を考えているか。'],
  ai: [
    '話すこと：AIをどの場面で使い、何を採用し、何を捨てたか。そして、最後に自分たちが何を基準に判断したかを話す。',
    '書き方：使った場面の列挙で終わらせない。「捨てたものと理由」「自分たちが判断したこと」を具体例で書く（第11回ワークシートの下書きを使う）。',
    'このスライドは必ず入れる。評価の観点：AIの活用が透明で、自分たちの判断を説明できるか。'],
  closing: [
    '話すこと：キャッチコピーをもう一度言い、覚えて帰ってほしいことを3つだけ。「以上で発表を終わります。ありがとうございました」で締める。',
    '書き方：新しい内容は足さない。3つは「誰のために」「何を変える」「まず何から」の順にすると伝わりやすい。',
    'このスライドを映したまま質疑に入る。'],
};
// ------------------------------------------------------------
// 型。layouts(P) でレイアウトを作り、slides に発表のスライドを順に並べる
//   slides の各行: layout（レイアウト名）、label（時間配分の表での名前）、sec（目安の秒数）、notes、
//                  use（作例でこのスライドに入れる内容の名前。examples/<題材>.js の content のキー）
// ------------------------------------------------------------
const PERSONA_ROWS = [['来店の目的', 'なぜ店に来るのか'], ['重視すること', '車選び・店選びで大事にすること'], ['不安・面倒なこと', '来店するときに感じること']];
const VARIANTS = {
  // 課題から順に積み上げる
  basic: {
    name: '基本型',
    idea: '課題から順に積み上げる基本の型。課題→顧客→コンセプト→空間→体験→効果と実現の道筋→AI活用の順に話す。初めての提案発表に向く。',
    steps: 7,
    layouts(P) {
      P.cover();
      P.rowsStat('1 課題', 1, '課題', { statHead: '私たちが着目した事実',
        rows: [['背景', '企業が置かれている状況'], ['制約', '守らなければならない条件'], ['期待される成果', '企業が提案に求めていること']] });
      P.persona('2 顧客', 2, '顧客', { rows: PERSONA_ROWS, voice: '「ペルソナの本音を、本人の言葉で」' });
      P.copy3('3 コンセプト', 3, 'コンセプト', { copy: 'キャッチコピー（20字以内）',
        cols: [['誰に', 'ペルソナを一言で'], ['何を', '提供する価値・体験'], ['どのように', '空間・接客・しくみ']] });
      P.imagePoints('4 空間', 4, '空間', {});
      P.flow4('5 体験', 5, '体験');
      P.effectPlan('6 効果・実現の道筋', 6, '効果・実現の道筋');
      P.cols4('7 AI活用', 7, 'AI活用と自分たちの思考');
      P.closing('まとめ');
    },
    slides: [
      { layout: '表紙', label: '表紙', sec: 15, notes: NOTES.cover, use: 'cover' },
      { layout: '1 課題', label: '1 課題', sec: 30, use: 'issue', notes: [
        '話すこと：企業から示された課題を、自分たちの言葉で言い直す。そのうえで、どこに着目したかを数字1つで示す。',
        '書き方：見出しは「〜が課題である」という文に。背景・制約・期待される成果は、それぞれ1〜2行。',
        '数字：出典を必ず書く。企業から受け取った非公開の情報は、スライドに載せてよいか確認する。',
        '評価の観点：企業の課題を正しく捉えているか。'] },
      { layout: '2 顧客', label: '2 顧客', sec: 30, notes: NOTES.persona, use: 'persona' },
      { layout: '3 コンセプト', label: '3 コンセプト', sec: 35, use: 'concept', notes: [
        '話すこと：キャッチコピーを読み上げ、「誰に・何を・どのように」で言い直す。前のスライドの本人の言葉に答える形にする。',
        '書き方：第7回のコンセプト文を使う。コピーはAIの案をそのまま使わず、自分たちの言葉に直したものを。',
        'ここが発表の中心。このスライドだけは、原稿を見ずに言えるようにしておく。'] },
      { layout: '4 空間', label: '4 空間', sec: 50, notes: NOTES.space, use: 'space' },
      { layout: '5 体験', label: '5 体験', sec: 40, notes: NOTES.flow, use: 'flow' },
      { layout: '6 効果・実現の道筋', label: '6 効果・実現', sec: 45, notes: NOTES.effect, use: 'effect' },
      { layout: '7 AI活用', label: '7 AI活用', sec: 30, notes: NOTES.ai, use: 'ai' },
      { layout: 'まとめ', label: 'まとめ', sec: 15, notes: NOTES.closing, use: 'closing' },
    ],
  },

  // お客さまの「ある場面」と、そこで本当にしたいことから組み立てる
  situation: {
    name: '場面起点型',
    idea: 'お客さまの「ある場面」から始める型。その場面で本当にしたいことを捉え、いまの選択肢では満たされない点を示してから、提案につなげる。',
    steps: 7,
    layouts(P) {
      P.cover();
      P.sceneRows('1 場面', 1, '場面', { image: '場面が伝わる画像・イラスト', voice: '「その場面での、本人のひとこと」',
        rows: [['いつ・どこで', 'どんな状況のときか'], ['きっかけ', '何が起きて、動こうと思ったか'], ['いまの不満', 'いまのやり方で困っていること']] });
      P.statement('2 したいこと', 2, '本当にしたいこと', {
        lines: [['こんなとき', '（場面）のとき'], ['こうしたい', '（片づけたいこと）をしたい'], ['そうすれば', '（その先の望み）ができる']],
        sides: [['実用の面', '実際に片づけたいこと'], ['気持ちの面', 'どんな気持ちになりたいか'], ['周りとの関係', '周りの人にどう見られたいか']] });
      P.alternatives('3 いまの選択肢', 3, 'いまの選択肢', { alts: ['ネットで検索', '知人に聞く', '何もしない'],
        walls: [['新しいことへの不安', '試すときに心配になること'], ['いまのやり方への慣れ', '変えるのが面倒に感じる理由']] });
      P.copy3('4 提案', 4, '提案', { copy: 'キャッチコピー（20字以内）',
        cols: [['したいことに', 'こう応える（提供する価値）'], ['気持ちに', 'こう応える（どんな気持ちになれるか）'], ['不安と慣れに', 'こう取り除く（安心して試せるしくみ）']] });
      P.imagePoints('5 空間と体験', 5, '空間と体験', { desc: 'どの「したいこと」に応えるか' });
      P.effectPlan('6 効果・実現の道筋', 6, '効果・実現の道筋');
      P.cols4('7 AI活用', 7, 'AI活用と自分たちの思考');
      P.closing('まとめ');
    },
    slides: [
      { layout: '表紙', label: '表紙', sec: 15, notes: NOTES.cover, use: 'cover' },
      { layout: '1 場面', label: '1 場面', sec: 30, notes: [
        '話すこと：お客さまが置かれている「ある場面」を1つ、目に浮かぶように話す。最後に本人のひとことを読み上げる。',
        '書き方：年齢や職業といった属性ではなく、状況で捉える。同じ人でも、場面が変われば求めるものが変わる。第6回のペルソナを、この場面に登場させる。',
        '「いまの不満」は、店への不満に限らない。車選びそのもので困っていることを書く。',
        '評価の観点：企業の課題を正しく捉えているか（この場面が、企業の課題とつながっていることを一言添える）。'],
        use: 'sit_scene' },
      { layout: '2 したいこと', label: '2 したいこと', sec: 35, notes: [
        '話すこと：「こんなとき、こうしたい。そうすれば、こうなれる」を、一続きの文として読み上げる。続けて、実用・気持ち・周りとの関係の3つの面で補う。',
        '書き方：お客さまが欲しいのは、車や店そのものではなく、その先にある「前に進むこと」。「車を買いたい」で止めず、なぜそうしたいのかをもう一段たずねる。',
        '気持ちの面と周りとの関係は見落としやすい。ペルソナの言葉やアンケートの自由記述から拾う。',
        'ここが発表の中心。この一文は、原稿を見ずに言えるようにしておく。'],
        use: 'sit_want' },
      { layout: '3 いまの選択肢', label: '3 いまの選択肢', sec: 40, notes: [
        '話すこと：お客さまが、いま代わりに何をしているか。それでは何が満たされないか。そして、なぜ新しい方法に踏み出せないのか。',
        '書き方：比べる相手は、同じ業種の店だけではない。ネットで調べる、知人に聞く、何もしないで先送りにする、も選択肢に入る。',
        '「踏み出せない理由」は、新しいことへの不安と、いまのやり方への慣れの2つに分けて書く。提案は、この2つを取り除くものでなければ選ばれない。',
        '事実（アンケート、インタビュー、観察）にもとづいて書き、出典を示す。'],
        use: 'sit_alts' },
      { layout: '4 提案', label: '4 提案', sec: 35, notes: [
        '話すこと：キャッチコピーを読み上げ、前の2枚で示した「したいこと」「気持ち」「不安と慣れ」のそれぞれに、どう応えるかを話す。',
        '書き方：第7回のコンセプトを、お客さまの側から言い直す。「店が何を提供するか」ではなく「お客さまが何をできるようになるか」で書く。',
        '評価の観点：ペルソナとコンセプトに一貫性があるか。'],
        use: 'sit_offer' },
      { layout: '5 空間と体験', label: '5 空間と体験', sec: 50, notes: NOTES.space.concat(['3つの点は、来店の流れ（入る→過ごす→相談する）の順に並べると、体験として伝わる。']), use: 'space' },
      { layout: '6 効果・実現の道筋', label: '6 効果・実現', sec: 40, notes: NOTES.effect, use: 'effect' },
      { layout: '7 AI活用', label: '7 AI活用', sec: 30, notes: NOTES.ai, use: 'ai' },
      { layout: 'まとめ', label: 'まとめ', sec: 15, notes: NOTES.closing, use: 'closing' },
    ],
  },

  // 一人の主人公の物語として語る
  story: {
    name: '物語型',
    idea: '一人の主人公の物語として語る型。ある日のつまずき→その原因→「もしも」の提案→新しい一日、の順に進め、聞き手に場面を思い描いてもらう。',
    steps: 7,
    layouts(P) {
      P.cover({ title: '物語のタイトル' });
      P.persona('1 主人公', 1, '主人公', { name: '主人公の名前（年齢）', voice: '「主人公の口ぐせ・本音」',
        rows: [['ふだんの暮らし', 'どんな毎日を送っているか'], ['大切にしていること', '何を大事に選ぶ人か'], ['叶えたいこと', 'この物語で手に入れたいもの']] });
      P.sceneRows('2 ある日', 2, 'ある日', { image: '出来事の場面の画像・イラスト', voice: '「そのとき、心の中でつぶやいたこと」',
        rows: [['どこで', '場所と時間'], ['何が起きたか', '主人公がつまずいた出来事'], ['どう感じたか', 'そのときの気持ち']] });
      P.rowsStat('3 なぜ起きるのか', 3, 'なぜ起きるのか', { statHead: '一人だけの話ではない',
        rows: [['店のしくみ', 'つまずきを生んでいる店の側の原因'], ['お客さまの思い', '主人公と同じ人たちが感じていること'], ['企業の課題', '企業から示された課題とのつながり']] });
      P.copy3('4 もしも', 4, 'もしも', { copy: 'もしも、〜な店だったら',
        cols: [['誰に', '主人公のような人たち'], ['何を', '提供する価値・体験'], ['どのように', '空間・接客・しくみ']] });
      P.storyboard3('5 新しい一日', 5, '新しい一日', { scenes: ['入店', '店内', '帰り道'], text: '主人公の行動と気持ち' });
      P.effectPlan('6 物語を現実に', 6, '物語を現実にする');
      P.cols4('7 AI活用', 7, 'AI活用と自分たちの思考');
      P.closing('結び', { title: '主人公の、その後のひとこと' });
    },
    slides: [
      { layout: '表紙', label: '表紙', sec: 15, use: 'cover', notes: [
        '話すこと：物語のタイトル（キャッチコピー）と、誰の物語かを一言で。「これから、ある一人のお客さまの話をします」と始める。',
        '書き方：タイトルは20字以内。その下に「誰に・何を・どのように」の一文を置く。',
        '画像：物語を象徴する1枚。生成画像なら「AIによる生成イメージ」と入れる。'] },
      { layout: '1 主人公', label: '1 主人公', sec: 25, notes: [
        '話すこと：主人公を、名前で呼んで紹介する。聞き手が「こういう人、いる」と思える具体的な点を2つか3つ。',
        '書き方：第6回のペルソナを主人公にする。属性の羅列ではなく、暮らしぶりと大切にしていることを書く。',
        '画像：実在の人物の写真は使わない。生成画像なら「AIによる生成イメージ」と入れる。'],
        use: 'st_hero' },
      { layout: '2 ある日', label: '2 ある日', sec: 35, notes: [
        '話すこと：主人公がつまずいた、ある日の出来事を1つ。場所、起きたこと、そのときの気持ちの順に、現在形で話すと場面が浮かぶ。',
        '書き方：出来事は1つに絞る。観察やインタビュー、アンケートで見つけた事実をもとにする（作り話にしない）。',
        '最後に、心の中のつぶやきを読み上げて、少し間を置く。'],
        use: 'st_day' },
      { layout: '3 なぜ起きるのか', label: '3 原因', sec: 35, notes: [
        '話すこと：主人公のつまずきは、なぜ起きるのか。店のしくみの側の原因を示し、数字1つで「一人だけの話ではない」ことを伝える。',
        '書き方：ここで企業の課題につなげる。「企業の課題」の行は、企業から示された課題を自分たちの言葉で言い直す。',
        '数字：出典を必ず書く。企業から受け取った非公開の情報は、スライドに載せてよいか確認する。',
        '評価の観点：企業の課題を正しく捉えているか。'],
        use: 'st_cause' },
      { layout: '4 もしも', label: '4 もしも', sec: 30, notes: [
        '話すこと：「もしも、〜な店だったら」と問いかけてから、コンセプトを「誰に・何を・どのように」で言う。物語の転換点。',
        '書き方：大きな一文は、問いかけの形でも、キャッチコピーでもよい。第7回のコンセプト文を使う。',
        '評価の観点：ペルソナとコンセプトに一貫性があるか。'],
        use: 'st_whatif' },
      { layout: '5 新しい一日', label: '5 新しい一日', sec: 60, notes: [
        '話すこと：同じ主人公が、新しい店を訪れた一日を、3つの場面で語る。「ある日」と同じ場所から始めると、変化がはっきりする。',
        '書き方：場面ごとに、空間のどこが変わり、主人公がどう動き、どう感じるかを1〜2文で。空間と動線の説明は、場面の中に織り込む。',
        '画像：第8回・第9回の図やイメージを使う。生成画像には「AIによる生成イメージ」と明記する。',
        '評価の観点：空間と動線の提案に根拠があるか（なぜその変更が主人公のつまずきに効くのかを、口頭で添える）。'],
        use: 'st_newday' },
      { layout: '6 物語を現実に', label: '6 現実に', sec: 40, notes: NOTES.effect, use: 'effect' },
      { layout: '7 AI活用', label: '7 AI活用', sec: 30, notes: NOTES.ai, use: 'ai' },
      { layout: '結び', label: '結び', sec: 15, notes: [
        '話すこと：主人公のその後のひとことで物語を閉じ、覚えて帰ってほしいことを3つだけ。「以上で発表を終わります。ありがとうございました」で締める。',
        '書き方：新しい内容は足さない。冒頭の「ある日」のつぶやきと対になる言葉にすると、変化が伝わる。',
        'このスライドを映したまま質疑に入る。'],
        use: 'st_end' },
    ],
  },

  // 結論を最初に言い、根拠を3つ示す
  conclusion: {
    name: '結論先行型',
    idea: '結論を最初に言う型。冒頭の1枚で提案の全体を伝え、そのあと根拠を3つ示す。時間が短いとき、判断する立場の人に聞いてもらうときに向く。',
    steps: 6,
    layouts(P) {
      P.cover();
      P.copy3('1 結論', 1, '結論', { copy: '提案を一文で（キャッチコピー）',
        cols: [['企業の課題', '示された課題を、自分たちの言葉で'], ['私たちの提案', '誰に・何を・どのように'], ['期待できる効果', '数字で示せる変化（目標）']] });
      P.persona('2 根拠1 顧客', 2, '根拠1　顧客が求めている', { rows: PERSONA_ROWS, voice: '「ペルソナの本音を、本人の言葉で」' });
      P.imagePoints('3 根拠2 空間', 3, '根拠2　空間で応えられる', {});
      P.flow4('4 根拠3 体験', 4, '根拠3　体験が変わる');
      P.effectPlan('5 効果と進め方', 5, '効果と進め方');
      P.cols4('6 AI活用', 6, 'AI活用と自分たちの思考');
      P.closing('まとめ', { title: '結論をもう一度', points: ['提案の要点1', '提案の要点2', 'まず始めること・ご検討いただきたいこと'] });
    },
    slides: [
      { layout: '表紙', label: '表紙', sec: 15, notes: NOTES.cover, use: 'cover' },
      { layout: '1 結論', label: '1 結論', sec: 45, notes: [
        '話すこと：最初に結論を言う。「課題は○○です。私たちは△△を提案します。これにより□□が期待できます」。このあと根拠を3つ話すことも予告する。',
        '書き方：この1枚だけで提案の全体がわかるようにする。聞き手がここで発表を止めても、判断の材料が残るように。',
        '効果の数字の根拠は「効果と進め方」のスライドで示す。ここでは結論だけを言い切る。',
        '評価の観点：企業の課題を正しく捉えているか。'],
        use: 'cc_conclusion' },
      { layout: '2 根拠1 顧客', label: '2 根拠1 顧客', sec: 35, notes: ['根拠の1つ目。「お客さまが、それを求めている」ことを示す。'].concat(NOTES.persona), use: 'persona' },
      { layout: '3 根拠2 空間', label: '3 根拠2 空間', sec: 50, notes: ['根拠の2つ目。「いまの店舗を変えることで、応えられる」ことを示す。'].concat(NOTES.space), use: 'space' },
      { layout: '4 根拠3 体験', label: '4 根拠3 体験', sec: 40, notes: ['根拠の3つ目。「お客さまの体験が、実際に変わる」ことを示す。'].concat(NOTES.flow), use: 'flow' },
      { layout: '5 効果と進め方', label: '5 効果と進め方', sec: 45, notes: NOTES.effect, use: 'effect' },
      { layout: '6 AI活用', label: '6 AI活用', sec: 30, notes: NOTES.ai, use: 'ai' },
      { layout: 'まとめ', label: 'まとめ', sec: 20, notes: [
        '話すこと：結論をもう一度言い、要点を2つ、最後に「まず始めること」または「ご検討いただきたいこと」を伝える。',
        '書き方：冒頭の結論と同じ言葉を使う。言い換えると、別の提案に聞こえる。',
        'このスライドを映したまま質疑に入る。'],
        use: 'closing' },
    ],
  },

  // 複数の案を同じ基準で比べ、選んだ理由を示す
  compare: {
    name: '比較型',
    idea: '複数の案を比べて、選んだ理由を示す型。判断の基準を先に示し、3つの案を同じ基準で比べる。自分たちの考えた道すじが伝わりやすい。',
    steps: 7,
    layouts(P) {
      P.cover();
      P.rowsStat('1 課題と判断の基準', 1, '課題と判断の基準', { statHead: '私たちが着目した事実',
        rows: [['企業の課題', '示された課題を、自分たちの言葉で'], ['守る条件', '制約（予算、広さ、期間など）'], ['判断の基準', '案を選ぶときに重視すること']] });
      P.persona('2 顧客', 2, '顧客', { rows: PERSONA_ROWS, voice: '「ペルソナの本音を、本人の言葉で」' });
      P.compare3('3 3つの案', 3, '検討した3つの案');
      P.copy3('4 選んだ案', 4, '選んだ案', { copy: 'キャッチコピー（20字以内）',
        cols: [['誰に', 'ペルソナを一言で'], ['何を', '提供する価値・体験'], ['どのように', '空間・接客・しくみ']] });
      P.imagePoints('5 空間と体験', 5, '空間と体験', {});
      P.effectPlan('6 効果・実現の道筋', 6, '効果・実現の道筋');
      P.cols4('7 AI活用', 7, 'AI活用と自分たちの思考');
      P.closing('まとめ');
    },
    slides: [
      { layout: '表紙', label: '表紙', sec: 15, notes: NOTES.cover, use: 'cover' },
      { layout: '1 課題と判断の基準', label: '1 課題と基準', sec: 30, notes: [
        '話すこと：企業の課題と守る条件を確認し、「この3つの基準で案を選びました」と先に示す。',
        '書き方：判断の基準は、企業の課題とペルソナから導く。案を決めてから基準を後づけしない。',
        '数字：出典を必ず書く。企業から受け取った非公開の情報は、スライドに載せてよいか確認する。',
        '評価の観点：企業の課題を正しく捉えているか。'],
        use: 'cp_criteria' },
      { layout: '2 顧客', label: '2 顧客', sec: 25, notes: NOTES.persona, use: 'persona' },
      { layout: '3 3つの案', label: '3 3つの案', sec: 45, notes: [
        '話すこと：検討した3つの案を、同じ基準で比べた結果を示す。表を全部読まず、差がついた点だけを話し、最後に選んだ理由を一文で言う。',
        '書き方：選んだ案を左端（A）に置く。評価は ◎○△ などの記号と、理由を短く。選ばなかった案の良い点も正直に書く。',
        '第7回で30案から絞った過程を使う。AIが出した案と自分たちの案が混ざっていてよいが、評価は自分たちで行う。',
        '選ばなかった案を低く見せるために、評価を作らない。質疑で必ず聞かれる。'],
        use: 'cp_options' },
      { layout: '4 選んだ案', label: '4 選んだ案', sec: 30, notes: [
        '話すこと：選んだ案を、キャッチコピーと「誰に・何を・どのように」で言い直す。',
        '書き方：第7回のコンセプト文を使う。コピーはAIの案をそのまま使わず、自分たちの言葉に直したものを。',
        '評価の観点：ペルソナとコンセプトに一貫性があるか。'],
        use: 'concept' },
      { layout: '5 空間と体験', label: '5 空間と体験', sec: 50, notes: NOTES.space.concat(['3つの点は、来店の流れ（入る→過ごす→相談する）の順に並べると、体験として伝わる。']), use: 'space' },
      { layout: '6 効果・実現の道筋', label: '6 効果・実現', sec: 40, notes: NOTES.effect, use: 'effect' },
      { layout: '7 AI活用', label: '7 AI活用', sec: 30, notes: NOTES.ai, use: 'ai' },
      { layout: 'まとめ', label: 'まとめ', sec: 15, notes: NOTES.closing, use: 'closing' },
    ],
  },
};

// ------------------------------------------------------------
// 作例の音声（発表の音声、教員の解説）
//   原稿は examples/narration/<題材>.json、音声と長さは examples/audio/<題材>/（make_audio.py が作る）。
//   キーは content と同じ（スライドの use）。型ごとの原稿「型:キー」があればそれを、なければ「キー」を使う
// ------------------------------------------------------------
const BUTTON = { talk: { x: 6.25, name: '発表の音声' }, comment: { x: 7.6, name: '教員の解説' }, y: 5.17, w: 1.25, h: 0.3 };
const sha1 = t => crypto.createHash('sha1').update(t, 'utf8').digest('hex').slice(0, 12);

function loadAudio(topic) {
  const dir = path.join(__dirname, 'examples');
  const nf = path.join(dir, 'narration', topic + '.json'), af = path.join(dir, 'audio', topic, 'audio.json');
  if (!fs.existsSync(nf) || !fs.existsSync(af)) return null;
  const text = JSON.parse(fs.readFileSync(nf, 'utf8')), info = JSON.parse(fs.readFileSync(af, 'utf8'));
  return (kind, id, use) => {
    const key = text[kind][`${id}:${use}`] !== undefined ? `${id}:${use}` : use;
    const t = text[kind][key], a = (info[kind] || {})[key];
    if (t === undefined) return null;
    const file = path.join(dir, 'audio', topic, `${kind}-${key.replace(':', '-')}.mp3`);
    if (!a || a.text !== sha1(t) || !fs.existsSync(file)) throw new Error(`${topic}: 音声「${kind} ${key}」がないか、原稿より古い（python3 examples/make_audio.py を実行する）`);
    return { kind, key, text: t, sec: a.sec, file, media: `${topic}-${path.basename(file)}` };
  };
}

// 音声の絵（ボタン）。絵をクリックすると、その音声を再生・一時停止する
function audioPic(c, id, rid) {
  const b = BUTTON[c.kind];
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${b.name}" descr="${b.name}（クリックで再生・一時停止）"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>` +
    `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="${rid}Audio"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">` +
    `<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${rid}Media"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rid}Icon"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${emu(b.x)}" y="${emu(BUTTON.y)}"/><a:ext cx="${emu(BUTTON.w)}" cy="${emu(BUTTON.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom></p:spPr></p:pic>`;
}
// アニメーションの設定。clips は [{ spid }]。
//   音声は自動では流さない。絵（ボタン）をクリックしたときだけ、その音声を再生・一時停止する。
//   （スライドの表示と同時に発表の音声を流し、続けて解説を流す設定は、PowerPoint では2つが同時に流れたのでやめた）
function audioTiming(clips) {
  let n = 1;   // id 1 は全体
  const tgt = c => `<p:tgtEl><p:spTgt spid="${c.spid}"/></p:tgtEl>`;
  const click = clips.map(c =>
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${++n}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
    `<p:stCondLst><p:cond evt="onClick" delay="0">${tgt(c)}</p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
    `<p:par><p:cTn id="${++n}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${++n}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${++n}" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="${++n}" dur="1" fill="hold"/>${tgt(c)}</p:cBhvr></p:cmd>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>` +
    `<p:nextCondLst><p:cond evt="onClick" delay="0">${tgt(c)}</p:cond></p:nextCondLst></p:seq>`).join('');
  const media = clips.map(c =>
    `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${++n}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>` +
    `<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn>${tgt(c)}</p:cMediaNode></p:audio>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${click}${media}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

// ------------------------------------------------------------
// 1つの型を pptx に書き出す
// ------------------------------------------------------------
// ex は作例の題材（examples/<題材>.js）。渡すと、入力枠に作例の内容を入れ、使い方のスライドを省く
async function build(id, ex) {
  const v = VARIANTS[id];
  const file = ex
    ? path.join(__dirname, 'examples', `presentation-example-${ex.id}-${id}.pptx`)
    : path.join(__dirname, `presentation-template-5min-${id}.pptx`);
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = ex ? `産学連携プロジェクト 提案発表の作例（5分・${v.name}・${ex.name}）` : `産学連携プロジェクト 提案発表テンプレート（5分・${v.name}）`;
  pres.subject = '基礎演習Ⅱ 産学連携プロジェクト';
  pres.author = '基礎演習Ⅱ';
  pres.company = '名古屋市立大学 経済学部';

  const LAYOUTS = {};   // レイアウト名 → { idx: 入力枠の定義 }
  const SLIDES = [];    // スライドごとの { layout, texts }
  function define(name, bg, objects, number = true) {
    const specs = {};
    const objs = objects.map((o, i) => {
      if (!o.ph) return o;
      specs[100 + i] = o.ph;   // pptxgenjs は入力枠に 100＋並び順 の番号を付ける
      return { placeholder: { options: { name: o.ph.key, type: o.ph.type === 'title' ? 'title' : 'body', x: o.ph.x, y: o.ph.y, w: o.ph.w, h: o.ph.h }, text: o.ph.prompt || ' ' } };
    });
    LAYOUTS[name] = specs;
    pres.defineSlideMaster({
      title: name, background: { color: bg }, objects: objs,
      slideNumber: number ? { x: 9.0, y: 5.2, w: 0.5, h: 0.25, fontSize: 10, color: INK2, align: 'right' } : undefined,
    });
    return name;
  }
  // 本文スライドの共通の枠：番号の丸＋章の名前、右上の進行の丸、見出し（1行のメッセージ）、フッター
  function frame(n, name) {
    const o = [];
    if (n) o.push(box(X0, 0.36, 0.34, 0.34, AMBER, { shape: 'ellipse', text: String(n), bold: true, name: '章の番号' }));
    if (name) o.push(label(name, n ? 0.94 : X0, 0.36, 5, 0.34));
    if (n) {
      for (let i = 1; i <= v.steps; i++) {
        const x = X0 + CW - (v.steps - i + 1) * 0.14 - (v.steps - i) * 0.1;
        o.push(box(x, 0.46, 0.14, 0.14, i === n ? AMBER : TINT, { shape: 'ellipse', name: '進行 ' + i }));
      }
    }
    o.push(ph('title', 'title', X0, 0.8, CW, 0.66, { size: 28, bold: true, anchor: 'ctr', prompt: 'この1枚で言いたいことを、1行の文で' }));
    o.push(label(FOOTER, X0, 5.2, 5, 0.25, { size: 10, bold: false, color: INK2 }));
    return o;
  }
  const P = patterns(frame, define);
  v.layouts(P);
  P.appendix();

  function add(layout, section, o = {}) {
    if (!LAYOUTS[layout]) throw new Error(`${id}: レイアウト「${layout}」がない`);
    const s = pres.addSlide({ masterName: layout, sectionTitle: section });
    if (ex && o.use && !ex.content[o.use]) throw new Error(`${ex.id}: 内容「${o.use}」がない（${id}・${layout}）`);
    // 作例の音声：発表のスライド（o.sec がある）には発表の音声と教員の解説、補足のスライドには解説だけを入れる
    const clips = [];
    if (ex && ex.audio && o.use) {
      const talk = o.sec ? ex.audio('talk', id, o.use) : null, comment = ex.audio('comment', id, o.use);
      if (o.sec && !talk) throw new Error(`${ex.id}: 発表の原稿「${o.use}」がない（${id}・${layout}）`);
      if (!comment) throw new Error(`${ex.id}: 解説の原稿「${o.use}」がない（${id}・${layout}）`);
      if (talk) {
        clips.push(talk);
        spoken.push({ label: layout, sec: o.sec, real: talk.sec });
        o.notes = o.notes.concat('', `■ この作例の発表の原稿（音声 ${Math.round(talk.sec)}秒・${talk.text.length}字）`, talk.text);
      }
      clips.push(comment);
      o.notes = o.notes.concat('', `■ 教員の解説（音声 ${Math.round(comment.sec)}秒）`, comment.text);
    }
    if (o.notes) s.addNotes(o.notes.join('\n'));
    if (o.hidden) s.hidden = true;
    SLIDES.push({ layout, texts: Object.assign({}, o.texts, ex && o.use ? ex.content[o.use] : null), clips });
    return s;
  }
  const spoken = [];   // 発表の音声の長さ（目安の秒数と比べる）
  (ex ? [SEC_TALK, SEC_APPENDIX] : [SEC_USAGE, SEC_TALK, SEC_APPENDIX]).forEach(t => pres.addSection({ title: t }));
  const total = v.slides.reduce((a, s) => a + s.sec, 0);
  if (total > TOTAL) throw new Error(`${id}: 目安の合計が ${total} 秒で、${TOTAL} 秒を超えている`);

  if (!ex) usage();
  // 使い方（非表示スライド）
  function usage() {
  const s = add('見出しのみ', SEC_USAGE, {
    hidden: true,
    texts: { title: `このテンプレートの使い方（${v.name}）` },
    notes: [
      'このスライドは「非表示スライド」です。スライドショーには出ませんが、印刷やPDFには含まれることがあるので、提出の前に削除してください。',
      '',
      '■ 入力のしかた',
      '・灰色の案内文をクリックして、自分たちの内容に置き換えます。入力しなかった枠は、スライドショーにも印刷にも出ません。',
      '・文字が枠に収まらないときは自動で小さくなります。小さくなりすぎたら、文字を減らしてください。',
      '・図や画像は、枠の中央のアイコンをクリックして入れます（ドラッグして入れることもできます）。',
      '',
      '■ スライドを増やす・減らす',
      '・［ホーム］→［新しいスライド］から、このテンプレートのレイアウト（各章、補足など）を選べます。',
      '・5分に収まらないときは、スライドを足すのではなく、話す内容を削ります。',
      '',
      '■ 色とフォントを変える',
      '・［デザイン］→［バリエーション］→［配色］／［フォント］で、全体をまとめて変えられます。',
      '',
      '■ 時間の目安',
      '・話す速さは1分あたり約300字。5分で約1,500字です。各スライドのノートに、目安の秒数と話す内容を書いてあります。',
      '・必ず時間を計って通しで練習し、超えたら削るところを決めます。',
      '',
      '■ 発表前のチェック',
      '・参考にしたデータや事例の出典を「参考・出典」のスライドに書いた。',
      '・生成画像に「AIによる生成イメージ」と明記した。',
      '・「AI活用と自分たちの思考」のスライドがある。',
      '・当日使う機材で表示を確認した。',
    ],
  });
  s.addText('非表示スライド（発表では表示されない）｜提出の前に削除する', {
    shape: 'roundRect', rectRadius: RADIUS, x: X0, y: 0.36, w: 4.7, h: 0.34, fill: { color: TINT },
    fontSize: 11, bold: true, color: AMBERTEXT, align: 'center', valign: 'middle', margin: 0, objectName: '注意書き',
  });
  const T = (text, o) => s.addText(text, Object.assign({ color: INK, valign: 'middle', margin: 0, isTextBox: true }, o));
  s.addText('', { shape: 'roundRect', rectRadius: RADIUS, x: X0, y: TOP, w: 4.8, h: 1.05, fill: { color: TINT }, objectName: '枠 この型の流れ' });
  T('この型の流れ', { x: 0.68, y: TOP + 0.08, w: 4.44, h: 0.26, fontSize: 11, bold: true, color: AMBERTEXT, objectName: '見出し この型の流れ' });
  T(v.idea, { x: 0.68, y: TOP + 0.35, w: 4.44, h: 0.64, fontSize: 12, valign: 'top', objectName: 'この型の流れ' });
  [['見出しは「文」で書く', '1枚1メッセージ。見出しだけで流れが伝わるように'],
   ['灰色の案内文を置き換える', '入力しない枠は表示されない。図は枠のアイコンから'],
   ['出典と「AIによる生成イメージ」を書く', '数字・事例・画像。出典は「参考・出典」にまとめる'],
   ['時間を計って、通しで練習する', '5分は約1,500字。ノートに目安の秒数がある'],
  ].forEach(([head, desc], i) => {
    const y = TOP + 1.2 + i * 0.53;
    s.addText(String(i + 1), { shape: 'ellipse', x: X0, y: y + 0.06, w: 0.36, h: 0.36, fill: { color: AMBER }, fontSize: 12, bold: true, color: INK, align: 'center', valign: 'middle', margin: 0, objectName: '番号 ' + (i + 1) });
    T(head, { x: 1.02, y, w: 4.28, h: 0.26, fontSize: 14, bold: true, objectName: '手順 ' + (i + 1) });
    T(desc, { x: 1.02, y: y + 0.26, w: 4.28, h: 0.22, fontSize: 12, color: INK2, objectName: '手順の説明 ' + (i + 1) });
  });
  // 時間配分（棒の長さが秒数）
  s.addText('', { shape: 'roundRect', rectRadius: RADIUS, x: 5.6, y: TOP, w: 3.9, h: 3.3, fill: { color: PAPER }, objectName: '枠 時間配分' });
  T('時間配分の目安', { x: 5.8, y: TOP + 0.12, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: AMBERTEXT, objectName: '見出し 時間配分' });
  const longest = Math.max(...v.slides.map(x => x.sec));
  v.slides.forEach((x, i) => {
    const y = TOP + 0.5 + i * 0.27;
    T(x.label, { x: 5.8, y, w: 1.5, h: 0.24, fontSize: 11, objectName: '時間 ' + x.label });
    s.addShape('rect', { x: 7.35, y: y + 0.05, w: x.sec / longest * 1.4, h: 0.14, fill: { color: AMBER }, objectName: '棒 ' + x.label });
    T(x.sec + '秒', { x: 8.8, y, w: 0.55, h: 0.24, fontSize: 11, align: 'right', objectName: '秒 ' + x.label });
  });
  T(`合計 ${Math.floor(total / 60)}分${total % 60}秒 ＋ 予備 ${TOTAL - total}秒 ＝ 5分`, { x: 5.8, y: TOP + 2.95, w: 3.55, h: 0.26, fontSize: 11, bold: true, align: 'right', objectName: '時間 合計' });
  }

  // 発表のスライド。ノートの先頭に、目安の秒数と原稿の字数（1秒5字）を入れる
  v.slides.forEach(x => add(x.layout, SEC_TALK, {
    notes: [`【目安 ${x.sec}秒｜原稿 約${Math.round(x.sec * 5 / 10) * 10}字】`].concat(x.notes), use: x.use, texts: x.texts, sec: x.sec,
  }));
  // 補足
  add('参考・出典', SEC_APPENDIX, {
    texts: { title: '参考・出典' },
    notes: ['発表の5分には含めません。質疑で根拠を聞かれたときに見せます。',
      '書き方：スライドに載せた数字、事例、図、画像の出典をすべて書く。Webページは URL と参照日も。',
      'AIの出力そのものは出典にならない。AIが示した情報は、もとの資料（一次資料）で確かめてから載せる。',
      '生成画像は、使ったツールの名前と「AIによる生成イメージ」であることを書く。'],
    use: 'sources',
  });
  add('補足（見出しと本文）', SEC_APPENDIX, {
    notes: ['発表の5分には含めません。質問されそうな点（費用の内訳、アンケートの詳細、ほかに検討した案など）を、1枚に1つずつ用意しておきます。',
      '使わなければ削除してください。増やすときは［ホーム］→［新しいスライド］から「補足」のレイアウトを選びます。'],
    use: 'appendix',
  });

  await pres.writeFile({ fileName: file });
  await finish(file, LAYOUTS, SLIDES, ex);
  console.log(`${path.basename(file)}  ${v.name}${ex ? '・' + ex.name : ''}  発表 ${v.slides.length} 枚（${Math.floor(total / 60)}分${total % 60}秒＋予備 ${TOTAL - total}秒）、全 ${SLIDES.length} 枚`);
  if (spoken.length) {
    // 発表の音声が、スライドごとの目安と、発表時間の5分に収まっているか
    const real = spoken.reduce((a, x) => a + x.real, 0);
    spoken.filter(x => x.real > x.sec).forEach(x => console.log(`  目安を超えている: ${x.label}  音声 ${x.real.toFixed(1)}秒 / 目安 ${x.sec}秒`));
    if (real > TOTAL) throw new Error(`${id}・${ex.id}: 発表の音声の合計が ${real.toFixed(1)} 秒で、${TOTAL} 秒を超えている`);
    const all = SLIDES.reduce((a, x) => a + x.clips.reduce((b, c) => b + c.sec, 0), 0);
    console.log(`  発表の音声 ${Math.floor(Math.round(real) / 60)}分${Math.round(real) % 60}秒、解説を含めた音声の合計 ${(all / 60).toFixed(1)}分`);
  }
}

// ------------------------------------------------------------
// 仕上げ：プレースホルダー、テーマ、ノートを書き直す
// ------------------------------------------------------------
const EMU = 914400;
const emu = v => Math.round(v * EMU);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const LANG = 'lang="ja-JP" altLang="en-US"';
const color = c => `<a:solidFill><a:schemeClr val="${c}"/></a:solidFill>`;
const phTag = (s, idx, prompt) => s.type === 'title'
  ? `<p:ph type="title"${prompt ? ' hasCustomPrompt="1"' : ''}/>`
  : `<p:ph type="${s.type}" sz="quarter" idx="${idx}"${prompt ? ' hasCustomPrompt="1"' : ''}/>`;
const phName = s => ({ title: '見出し', body: '本文', pic: '図' })[s.type] + ' ' + s.key;
const nv = (s, idx, id, prompt) =>
  `<p:nvSpPr><p:cNvPr id="${id}" name="${esc(phName(s))}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phTag(s, idx, prompt)}</p:nvPr></p:nvSpPr>`;

// レイアウト側：位置、余白、縦位置、文字の既定、自動調整（収まらないときは縮小）、案内文
function layoutPh(s, idx, id) {
  const size = s.size || 14, ls = s.ls || 1.1;
  let geom = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
  if (s.geom === 'ellipse') geom = '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
  if (s.geom === 'roundRect') geom = `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.round(RADIUS / Math.min(s.w, s.h) * 100000)}"/></a:avLst></a:prstGeom>`;
  const pic = s.type === 'pic';
  const ppr = lvl => {
    const ind = s.bullet ? ` marL="${emu(0.25 * lvl)}" indent="-${emu(0.25)}"` : ` marL="${emu(0.25 * (lvl - 1))}" indent="0"`;
    const bu = s.bullet ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : '<a:buNone/>';
    return `<a:lvl${lvl}pPr${ind} algn="${pic ? 'ctr' : s.align || 'l'}"><a:lnSpc><a:spcPct val="${Math.round(ls * 100000)}"/></a:lnSpc>` +
      `<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="${s.bullet ? 600 : 0}"/></a:spcAft>${bu}` +
      `<a:defRPr sz="${size * 100}" b="${s.bold ? 1 : 0}">${color(s.color || INK)}</a:defRPr></a:lvl${lvl}pPr>`;
  };
  return `<p:sp>${nv(s, idx, id, true)}<p:spPr><a:xfrm><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(s.w)}" cy="${emu(s.h)}"/></a:xfrm>${geom}</p:spPr>` +
    `<p:txBody><a:bodyPr vert="horz" wrap="square" lIns="${pic ? emu(0.15) : 0}" tIns="0" rIns="${pic ? emu(0.15) : 0}" bIns="0" rtlCol="0" anchor="${pic ? 'ctr' : s.anchor || 't'}">${pic ? '' : '<a:normAutofit/>'}</a:bodyPr>` +
    `<a:lstStyle>${[1, 2, 3, 4, 5].map(ppr).join('')}</a:lstStyle>` +
    `<a:p><a:r><a:rPr ${LANG}>${color(s.promptColor || INK2)}</a:rPr><a:t>${esc(s.prompt || '')}</a:t></a:r><a:endParaRPr ${LANG}/></a:p></p:txBody></p:sp>`;
}
// スライド側：位置も書式もレイアウトから引き継ぐ。文字があれば段落にする
//   図の枠に画像（{ rid, alt }）があれば、枠の形と位置を引き継いだ画像にする
function slidePh(s, idx, id, text) {
  if (s.type === 'pic' && text) {
    return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(phName(s))}" descr="${esc(text.alt)}"/><p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr>` +
      `<p:nvPr>${phTag(s, idx, false)}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${text.rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>`;
  }
  if (s.type === 'pic') return `<p:sp>${nv(s, idx, id, false)}<p:spPr/></p:sp>`;
  const paras = text
    ? String(text).split('\n').map(t => `<a:p><a:r><a:rPr ${LANG}/><a:t>${esc(t)}</a:t></a:r></a:p>`).join('')
    : `<a:p><a:endParaRPr ${LANG}/></a:p>`;
  return `<p:sp>${nv(s, idx, id, false)}<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`;
}
// pptxgenjs が書いたプレースホルダー（idx が 100 番台の図形）を、作り直したものに置き換える
function replacePh(xml, specs, make) {
  return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, sp => {
    const m = sp.match(/<p:ph\s+idx="(1\d\d)"/);
    if (!m || !specs[m[1]]) return sp;
    return make(specs[m[1]], Number(m[1]), sp.match(/<p:cNvPr id="(\d+)"/)[1]);
  });
}
// 1つの段落に重ねて書かれた a:pPr を、最初の1つだけにする
const onePPr = xml => xml.replace(/<a:p>[\s\S]*?<\/a:p>/g, p => {
  const i = p.indexOf('<a:pPr');
  if (i < 0) return p;
  return p.slice(0, i + 6) + p.slice(i + 6).replace(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g, '');
});

function fixTheme(xml) {
  for (const [k, v] of Object.entries(THEME.colors)) {
    xml = xml.replace(new RegExp(`<a:${k}>[\\s\\S]*?</a:${k}>`), `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`);
  }
  return xml.replace('<a:clrScheme name="Office">', `<a:clrScheme name="${THEME.name}">`)
    .replace('<a:fontScheme name="Office">', `<a:fontScheme name="${THEME.name}">`)
    .replace(/<a:ea typeface=""\/>/g, `<a:ea typeface="${FONT}"/>`)
    .replace(/<a:font script="Jpan" typeface="[^"]*"\/>/g, `<a:font script="Jpan" typeface="${FONT}"/>`);
}
// ノートは1つの段落に改行入りで書かれるので、行ごとの段落にする
function fixNotes(xml) {
  return xml.replace(/<a:p><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>([\s\S]*?)<\/a:t><\/a:r><a:endParaRPr lang="en-US" dirty="0"\/><\/a:p>/, (_, t) =>
    t.split('\n').map(line => line ? `<a:p><a:r><a:rPr ${LANG}/><a:t>${line}</a:t></a:r></a:p>` : `<a:p><a:endParaRPr ${LANG}/></a:p>`).join(''));
}

async function finish(file, LAYOUTS, SLIDES, ex) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const read = name => zip.file(name).async('string');
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
  const layoutName = {};   // slideLayoutN.xml → レイアウト名
  // 作例の音声：音声と絵を ppt/media に入れ、スライドに音声の絵（ボタン）とアニメーションの設定を足す
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', REL_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media';
  function addAudio(xml, clips, added) {
    if (!clips.length) return xml;
    let id = Math.max(...[...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map(x => Number(x[1])));
    const pics = clips.map(c => {
      const icon = `audio-${c.kind}.png`, rid = 'rId' + c.kind[0].toUpperCase() + c.kind.slice(1);
      if (!zip.file('ppt/media/' + c.media)) zip.file('ppt/media/' + c.media, fs.readFileSync(c.file));
      if (!zip.file('ppt/media/' + icon)) zip.file('ppt/media/' + icon, fs.readFileSync(path.join(__dirname, 'examples', 'images', 'audio', c.kind + '.png')));
      added.push(`<Relationship Id="${rid}Media" Type="${REL_MEDIA}" Target="../media/${c.media}"/>`,
        `<Relationship Id="${rid}Audio" Type="${REL}/audio" Target="../media/${c.media}"/>`,
        `<Relationship Id="${rid}Icon" Type="${REL}/image" Target="../media/${icon}"/>`);
      c.spid = ++id;
      return audioPic(c, id, rid);
    });
    const timing = audioTiming(clips);
    if (!xml.includes('</p:spTree>') || !xml.includes('</p:clrMapOvr>')) throw new Error('音声を入れる場所が見つからない');
    return xml.replace('</p:spTree>', pics.join('') + '</p:spTree>').replace('</p:clrMapOvr>', '</p:clrMapOvr>' + timing);
  }
  for (const name of names) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)) {
      let xml = await read(name);
      const title = xml.match(/<p:cSld name="([^"]*)"/)[1];
      layoutName[path.basename(name)] = title;
      if (title === 'DEFAULT') xml = xml.replace('<p:cSld name="DEFAULT"', '<p:cSld name="白紙"');
      if (LAYOUTS[title]) xml = replacePh(xml, LAYOUTS[title], layoutPh);
      zip.file(name, onePPr(xml).replace(/lang="en-US"/g, LANG));
    }
  }
  for (const name of names) {
    const m = name.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) {
      const rels = await read(`ppt/slides/_rels/slide${m[1]}.xml.rels`);
      const layout = layoutName[rels.match(/slideLayouts\/(slideLayout\d+\.xml)/)[1]];
      const slide = SLIDES[Number(m[1]) - 1];
      if (slide.layout !== layout) throw new Error(`${name}: レイアウトが合わない（${slide.layout} / ${layout}）`);
      const unknown = Object.keys(slide.texts).filter(k => !Object.values(LAYOUTS[layout]).some(s => s.key === k));
      if (unknown.length) throw new Error(`${name}: レイアウト「${layout}」にない入力枠 ${unknown.join(', ')}`);
      // 作例の図（図の枠に入れるファイル名）は、ppt/media に入れてスライドから参照する
      const added = [];
      const xml = replacePh(await read(name), LAYOUTS[layout], (s, idx, id) => {
        let value = slide.texts[s.key];
        if (s.type === 'pic' && value) {
          const media = `${ex.id}-${value}`;
          if (!zip.file('ppt/media/' + media)) zip.file('ppt/media/' + media, fs.readFileSync(path.join(__dirname, 'examples', 'images', ex.id, value)));
          const rid = 'rIdImage' + (added.length + 1);
          added.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${media}"/>`);
          value = { rid, alt: (ex.alts || {})[value] || '図' };
        }
        return slidePh(s, idx, id, value);
      });
      zip.file(name, addAudio(onePPr(xml).replace(/lang="en-US"/g, LANG), slide.clips, added));
      if (added.length) zip.file(`ppt/slides/_rels/slide${m[1]}.xml.rels`, rels.replace('</Relationships>', added.join('') + '</Relationships>'));
    } else if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) {
      zip.file(name, fixNotes(await read(name)));
    } else if (name === 'ppt/theme/theme1.xml') {
      zip.file(name, fixTheme(await read(name)));
    } else if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)) {
      zip.file(name, (await read(name)).replace(/lang="en-US"/g, LANG));
    }
  }
  const types = await read('[Content_Types].xml');
  const need = [['png', 'image/png']].concat(SLIDES.some(x => x.clips.length) ? [['mp3', 'audio/mpeg']] : []).filter(([e]) => !types.includes(`Extension="${e}"`));
  if (need.length) zip.file('[Content_Types].xml', types.replace('<Default ', need.map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('') + '<Default '));
  // フォルダーの項目を除いて詰め直す（フォルダーの項目があると、クラウド上で開けないことがある）
  const out = new JSZip();
  for (const name of Object.keys(zip.files).filter(n => !zip.files[n].dir)) out.file(name, await zip.files[name].async('nodebuffer'), { createFolders: false });
  fs.writeFileSync(file, await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

(async () => {
  const args = process.argv.slice(2).filter(a => !a.startsWith('--'));
  const dir = path.join(__dirname, 'examples');
  const topics = EXAMPLES ? fs.readdirSync(dir).filter(f => f.endsWith('.js')).map(f => f.slice(0, -3)).sort() : [];
  const unknown = args.filter(a => !VARIANTS[a] && !topics.includes(a));
  if (unknown.length) { console.error(`型・題材の名前が違う: ${unknown.join(', ')}（型は ${Object.keys(VARIANTS).join('・')}）`); process.exit(1); }
  const pick = (all, chosen) => (chosen.length ? chosen : all);
  const ids = pick(Object.keys(VARIANTS), args.filter(a => VARIANTS[a]));
  if (!EXAMPLES) { for (const id of ids) await build(id); return; }
  for (const topic of pick(topics, args.filter(a => topics.includes(a)))) {
    const ex = Object.assign({ id: topic }, require(path.join(dir, topic + '.js')), { audio: loadAudio(topic) });
    if (!ex.audio) console.log(`${topic}: 音声がないので、音声なしで組む（python3 examples/make_audio.py で作る）`);
    for (const id of ids) await build(id, ex);
  }
})();
