// 基礎演習Ⅱ 第1回〜第5回 総まとめ　音声付きスライドの組み立て
//   node build.js  →  summary-01-05-slides.pptx（音声は audio/NN.mp3、ナレーション原稿は narration.json）
//   そのあと python3 finish.py で、テーマのフォント・色の設定と、音声の自動再生・自動送りを書き込む。
//   必要なもの: npm install pptxgenjs
'use strict';
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const HERE = __dirname;
const spec = JSON.parse(fs.readFileSync(path.join(HERE, 'narration.json'), 'utf8'));
const durations = JSON.parse(fs.readFileSync(path.join(HERE, 'audio', 'durations.json'), 'utf8'));
const narr = Object.fromEntries(spec.slides.map(s => [s.id, s.narration]));

// 色（授業サイト・ワークシートに合わせる）
const INK = '1D2330', INK2 = '4A5261', INK3 = '757D8A', RULE = 'B9BEB4', PAPER = 'ECEEE8';
const BLUE = '2747C9', BLUETINT = 'DCE2F9', ORANGE = 'D68A17', ORANGETINT = 'F6E4C3', RED = 'B3261E', REDTINT = 'F9DEDA';
const FONT = 'Yu Gothic';

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.author = '河合 勝彦';
pres.title = '基礎演習Ⅱ 第1回〜第5回 総まとめ';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// ------------------------------------------------------------
// 収まりの確認（日本語は全角1文字＝フォントサイズ分の幅とみなす）
// ------------------------------------------------------------
const warnings = [];
function textWidthIn(str, size) {
  let em = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    if (ch === ' ') em += 0.3; else if (c > 0x2e7f) em += 1.0; else em += 0.56;
  }
  return em * size / 72;
}
function checkFit(label, paras, w, h, size, lineSp = 1.3) {
  let lines = 0;
  for (const p of paras) lines += Math.max(1, Math.ceil(textWidthIn(p, size) / Math.max(w, 0.1)));
  const need = lines * size / 72 * lineSp;
  if (need > h + 0.02) warnings.push(`${label}: ${lines}行 ${need.toFixed(2)}in > ${h}in  (${paras[0].slice(0, 18)}…)`);
}

// ------------------------------------------------------------
// 部品
// ------------------------------------------------------------
let slideNo = 0, curSlide = null;
function txt(text, o) {
  const paras = Array.isArray(text) ? text.filter(r => r.text).map(r => r.text) : String(text).split('\n');
  checkFit(`slide ${slideNo} ${o.label || ''}`, paras, o.w, o.h, o.fontSize || 14, o.lineSp || 1.3);
  const opts = { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: FONT, fontSize: o.fontSize || 14, color: o.color || INK,
    bold: !!o.bold, align: o.align || 'left', valign: o.valign || 'top', margin: 0, isTextBox: true,
    lineSpacingMultiple: o.lineSp || 1.3, paraSpaceAfter: o.paraSpaceAfter || 0 };
  if (o.fill) opts.fill = { color: o.fill };
  if (o.inset !== undefined) opts.margin = o.inset;
  curSlide.addText(Array.isArray(text) ? text : String(text), opts);
}
function rect(x, y, w, h, fill, o = {}) {
  curSlide.addShape(o.round ? pres.ShapeType.roundRect : pres.ShapeType.rect,
    { x, y, w, h, fill: { color: fill }, line: { color: o.line || fill, width: o.lineW || 0 }, rectRadius: o.round ? 0.08 : undefined });
}
function circleNum(n, x, y, d = 0.42, fill = BLUE, color = 'FFFFFF') {
  curSlide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  curSlide.addText(String(n), { x, y, w: d, h: d, fontFace: FONT, fontSize: d >= 0.4 ? 14 : 11, bold: true, color, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
}
function arrowRight(x, y, w = 0.3, h = 0.3, color = RULE) {
  curSlide.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill: { color }, line: { color, width: 0 } });
}
function arrowDown(x, y, w = 0.3, h = 0.28, color = RULE) {
  curSlide.addShape(pres.ShapeType.downArrow, { x, y, w, h, fill: { color }, line: { color, width: 0 } });
}
// 見出し付きの枠　head は枠の中の小見出し
function card(x, y, w, h, o) {
  rect(x, y, w, h, o.fill || PAPER, { round: true });
  let cy = y + 0.14;
  if (o.head) { txt(o.head, { x: x + 0.18, y: cy, w: w - 0.36, h: 0.32, fontSize: o.headSize || 13, bold: true, color: o.headColor || BLUE, label: 'card-head' }); cy += 0.36; }
  if (o.body) txt(o.body, { x: x + 0.18, y: cy, w: w - 0.36, h: y + h - cy - 0.12, fontSize: o.bodySize || 12, color: o.bodyColor || INK, lineSp: 1.3, label: 'card-body' });
}
// 番号丸＋太字見出し＋説明の1行
function numRow(n, x, y, w, head, body, o = {}) {
  const d = 0.4;
  circleNum(n, x, y + 0.02, d, o.fill || BLUE);
  const tx = x + d + 0.18, tw = w - d - 0.18;
  if (body) {
    txt(head, { x: tx, y, w: tw, h: 0.3, fontSize: o.headSize || 14, bold: true, label: 'numRow-head' });
    txt(body, { x: tx, y: y + 0.3, w: tw, h: o.bodyH || 0.5, fontSize: o.bodySize || 11.5, color: INK2, label: 'numRow-body' });
  } else {
    txt(head, { x: tx, y, w: tw, h: d + 0.04, fontSize: o.headSize || 14, bold: !!o.bold, valign: 'middle', label: 'numRow-head' });
  }
}
function chip(text, x, y, w, h, o = {}) {
  rect(x, y, w, h, o.fill || BLUETINT, { round: true });
  txt(text, { x: x + 0.1, y, w: w - 0.2, h, fontSize: o.fontSize || 12, bold: o.bold !== false, color: o.color || BLUE, align: 'center', valign: 'middle', label: 'chip' });
}
function addAudio(id) {
  const file = path.join(HERE, 'audio', String(id).padStart(2, '0') + '.mp3');
  curSlide.addMedia({ type: 'audio', path: file, x: 9.32, y: 5.0, w: 0.42, h: 0.42 });
}

// ------------------------------------------------------------
// レイアウト
// ------------------------------------------------------------
const FOOT = '基礎演習Ⅱ　2026年度後期　第1回〜第5回 総まとめ';
pres.defineSlideMaster({ title: 'TITLE', background: { color: BLUE } });
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: 'FFFFFF' },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: 0.5, y: 0.32, w: 8.8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: BLUE, margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.6, w: 9.0, h: 0.62, fontFace: FONT, fontSize: 26, bold: true, color: INK, margin: 0, valign: 'top' }, text: '' } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.22, w: 6, h: 0.25, fontFace: FONT, fontSize: 9, color: INK3, margin: 0 } } },
  ],
  slideNumber: { x: 8.7, y: 5.22, w: 0.5, h: 0.25, fontFace: FONT, fontSize: 9, color: INK3, align: 'right' },
});
pres.defineSlideMaster({ title: 'DARK', background: { color: INK } });

function content(id, kicker, title) {
  slideNo = id; curSlide = pres.addSlide({ masterName: 'CONTENT' });
  curSlide.addText(kicker, { placeholder: 'kicker' });
  checkFit(`slide ${id} title`, [title], 9.0, 0.62, 26, 1.2);
  curSlide.addText(title, { placeholder: 'title' });
  curSlide.addNotes(narr[id]);
  addAudio(id);
  return curSlide;
}

// ============================================================
// 1 タイトル
// ============================================================
slideNo = 1; curSlide = pres.addSlide({ masterName: 'TITLE' });
txt('名古屋市立大学 経済学部　2026年度後期　水曜3限', { x: 0.7, y: 0.75, w: 8.6, h: 0.35, fontSize: 14, color: 'FFFFFF' });
txt('基礎演習Ⅱ', { x: 0.7, y: 1.15, w: 8.6, h: 0.9, fontSize: 44, bold: true, color: 'FFFFFF', lineSp: 1.1 });
txt('第1回〜第5回　総まとめ', { x: 0.7, y: 2.05, w: 8.6, h: 0.7, fontSize: 30, bold: true, color: 'FFFFFF', lineSp: 1.1 });
txt('生成AIを「思考のパートナー」にする探究の基礎', { x: 0.7, y: 2.8, w: 8.6, h: 0.4, fontSize: 16, color: BLUETINT });
for (let i = 0; i < 5; i++) {
  const x = 0.7 + i * 0.62;
  curSlide.addShape(pres.ShapeType.ellipse, { x, y: 3.5, w: 0.48, h: 0.48, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF', width: 0 } });
  curSlide.addText(String(i + 1), { x, y: 3.5, w: 0.48, h: 0.48, fontFace: FONT, fontSize: 15, bold: true, color: BLUE, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
}
txt('ガイダンス　／　問いの立て方　／　先行研究の探索\n文献講読と批判的思考　／　リサーチ・クエスチョンの確定', { x: 3.9, y: 3.56, w: 5.6, h: 0.4, fontSize: 10.5, color: BLUETINT, lineSp: 1.25 });
txt('河合 勝彦（名古屋市立大学 経済学部）', { x: 0.7, y: 4.55, w: 6, h: 0.3, fontSize: 13, color: 'FFFFFF' });
txt('2026年10月28日　｜　音声付きスライド（約13分）', { x: 0.7, y: 4.85, w: 6, h: 0.3, fontSize: 11, color: BLUETINT });
curSlide.addNotes(narr[1]); addAudio(1);

// ============================================================
// 2 この動画で話すこと
// ============================================================
content(2, '総まとめ', 'この動画で話すこと');
[
  ['前半5回の流れ', '第1回から第5回まで、テーマが問いになるまでの一本道'],
  ['各回の要点', '第1回〜第5回。学んだことと、出席クイズで扱ったAIの用語'],
  ['5回を貫く3つの習慣', '分担、検証、記録'],
  ['後半へのつながりと、前半の提出物', '店舗設計提案プロジェクトと、確定レポートの要件'],
  ['付録：使ってみたいAIプロンプト集', '28個のプロンプトの使い方'],
].forEach(([h, b], i) => numRow(i + 1, 0.5, 1.45 + i * 0.72, 5.5, h, b));
card(6.5, 1.45, 3.0, 3.45, { head: '冊子と一緒に見てください', body: '第1部　総まとめ（4ページ）\n\n付録A　使ってみたいAIプロンプト集（28個、解説付き）\n\n付録B　出席確認クイズで扱ったAIの用語\n\nPDFは授業サイトの第5回のページから', bodySize: 12 });

// ============================================================
// 3 前半5回の流れ
// ============================================================
content(3, '探究の基礎（第1〜5回）', 'テーマが問いになるまでの一本道');
const flow = [
  ['ガイダンス', '嘘をつく相手', '関心テーマ3つ'],
  ['問いの立て方', 'ブレストの相手', '仮のリサーチ・クエスチョン'],
  ['先行研究の探索', '探す場所の案内役', '実在を確認した文献5件'],
  ['文献講読と批判的思考', '読む前の地図、議論の相手', '主張・根拠・限界、未解決の要素'],
  ['リサーチ・クエスチョンの確定', '厳しい査読者', '確定した問い、確定レポート'],
];
flow.forEach(([t, role, out], i) => {
  const x = 0.5 + i * 1.84, w = 1.72, y = 1.4, h = 2.55;
  rect(x, y, w, h, PAPER, { round: true });
  circleNum(i + 1, x + 0.15, y + 0.15, 0.4);
  txt(t, { x: x + 0.15, y: y + 0.62, w: w - 0.3, h: 0.5, fontSize: 13, bold: true, lineSp: 1.2, label: 'flow-title' });
  txt('AIの役割', { x: x + 0.15, y: y + 1.14, w: w - 0.3, h: 0.22, fontSize: 9.5, bold: true, color: BLUE });
  txt(role, { x: x + 0.15, y: y + 1.36, w: w - 0.3, h: 0.42, fontSize: 10.5, color: INK2, lineSp: 1.25, label: 'flow-role' });
  txt('持ち帰ったもの', { x: x + 0.15, y: y + 1.84, w: w - 0.3, h: 0.22, fontSize: 9.5, bold: true, color: BLUE });
  txt(out, { x: x + 0.15, y: y + 2.06, w: w - 0.3, h: 0.42, fontSize: 10.5, color: INK2, lineSp: 1.25, label: 'flow-out' });
  if (i < 4) arrowRight(x + w + 0.0, y + 1.1, 0.14, 0.26, RULE);
});
rect(0.5, 4.2, 9.0, 0.7, BLUETINT, { round: true });
txt('どの回でも自分の仕事：問いを選ぶ　・　事実を確かめる　・　自分の言葉で書く', { x: 0.7, y: 4.2, w: 8.6, h: 0.7, fontSize: 14, bold: true, color: BLUE, valign: 'middle' });

// ============================================================
// 4 第1回 仕組みと4つの考え方
// ============================================================
content(4, '第1回　ガイダンス｜9月30日', '授業の仕組みと、AI利用の4つの考え方');
rect(0.5, 1.4, 1.75, 1.35, PAPER, { round: true });
txt('50%', { x: 0.5, y: 1.5, w: 1.75, h: 0.7, fontSize: 36, bold: true, color: BLUE, align: 'center', lineSp: 1.0 });
txt('提出物', { x: 0.5, y: 2.2, w: 1.75, h: 0.4, fontSize: 13, align: 'center', color: INK2 });
rect(2.4, 1.4, 1.75, 1.35, PAPER, { round: true });
txt('50%', { x: 2.4, y: 1.5, w: 1.75, h: 0.7, fontSize: 36, bold: true, color: BLUE, align: 'center', lineSp: 1.0 });
txt('発表', { x: 2.4, y: 2.2, w: 1.75, h: 0.4, fontSize: 13, align: 'center', color: INK2 });
card(0.5, 2.95, 3.65, 1.95, { head: '二部構成', body: '探究の基礎（第1〜5回）\n実践プロジェクト（第6〜13回）\nキャリア講習（第14・15回）\n\nすべての提出物にAI活用記録を添付', bodySize: 12 });
[
  ['問いは自分が持つ', 'テーマを決め、何を知りたいかを言葉にするのは自分。AIは視点を広げる相手'],
  ['出力は素材であって完成品ではない', '読み、検証し、自分の言葉で書き直す'],
  ['事実は一次資料で確かめる', '数値、固有名詞、文献は、元の資料（論文、統計、公式サイト）で'],
  ['使ったことを隠さない', 'どこでどうAIを使ったかを記録し、提出物に明記する'],
].forEach(([h, b], i) => numRow(i + 1, 4.55, 1.45 + i * 0.88, 4.95, h, b, { bodyH: 0.5 }));

// ============================================================
// 5 第1回 ハルシネーション
// ============================================================
content(5, '第1回　ガイダンス｜9月30日', 'ハルシネーションと、出力を使う前のチェック');
const steps5 = [
  ['AIに嘘をつかせる', '大学の歴史を、参考文献3つ付きで説明させる'],
  ['実在しない文献が出てくる', '実在しそうな体裁。よく知っている話題でも年号や固有名詞がずれる'],
  ['なぜ起きるのか', 'LLMは「次に来る単語」を確率で予測しているだけ。意味を理解して調べてはいない'],
];
steps5.forEach(([h, b], i) => {
  const y = 1.4 + i * 1.2;
  card(0.5, y, 4.3, 1.06, { head: h, body: b, bodySize: 11.5 });
  if (i < 2) arrowDown(2.5, y + 1.05, 0.3, 0.16, RULE);
});
rect(5.2, 1.4, 4.3, 3.5, BLUETINT, { round: true });
txt('出力を使う前のチェック', { x: 5.45, y: 1.55, w: 3.9, h: 0.35, fontSize: 15, bold: true, color: BLUE });
txt('□ 固有名詞（人名、組織名、書名、論文名）は実在するか\n□ 数値には出典があるか。出典にあたって同じ数値か\n□ 根拠のない断定になっていないか\n□ 反対の見解や例外はないか\n□ 自分で説明できるか。説明できない内容は使わない',
  { x: 5.45, y: 2.0, w: 3.9, h: 2.8, fontSize: 12.5, lineSp: 1.35, paraSpaceAfter: 6, label: 'checks' });

// ============================================================
// 6 第2回 問いの立て方
// ============================================================
content(6, '第2回　問いの立て方｜10月7日', 'テーマを問いに変える');
card(0.5, 1.38, 4.3, 1.0, { head: 'テーマ＝話題の範囲', body: '例：地方の自動車販売', bodySize: 13 });
arrowRight(4.85, 1.72, 0.3, 0.3, BLUE);
card(5.2, 1.38, 4.3, 1.0, { head: '問い＝答えを求める疑問文', body: '例：なぜ若い世代の来店は減っているのか', bodySize: 13, fill: BLUETINT });
txt('良い問いの3条件', { x: 0.5, y: 2.55, w: 4, h: 0.3, fontSize: 13, bold: true, color: BLUE });
['調べて答えられる', '範囲が広すぎず狭すぎない', '答える意義がある'].forEach((t, i) => {
  const x = 0.5 + i * 3.05;
  circleNum(i + 1, x, 2.9, 0.4);
  txt(t, { x: x + 0.52, y: 2.9, w: 2.4, h: 0.44, fontSize: 13.5, bold: true, valign: 'middle' });
});
txt('第2回のワーク', { x: 0.5, y: 3.55, w: 4, h: 0.3, fontSize: 13, bold: true, color: BLUE });
const chips6 = ['AIに6つの視点から候補を出させる', '3条件で採点し、上位3つに絞る', 'ペアで「それはどうやって調べるの」', '仮の問いを1つ決める'];
chips6.forEach((t, i) => {
  const w = 2.05, x = 0.5 + i * (w + 0.27);
  chip(t, x, 3.9, w, 0.95, { fontSize: 11.5, fill: i === 3 ? BLUE : BLUETINT, color: i === 3 ? 'FFFFFF' : BLUE });
  if (i < 3) arrowRight(x + w + 0.03, 4.25, 0.2, 0.26, RULE);
});

// ============================================================
// 7 第2回 プロンプトの4要素
// ============================================================
content(7, '第2回　出席クイズ｜大規模言語モデルとプロンプト', 'プロンプトの4要素');
[
  ['役割', 'AIに立ってほしい立場。「あなたは経営学を専門とする大学教員です」「あなたは厳しい査読者です」'],
  ['目的と背景', '何のために、どんな状況で必要か。「2年生の演習で、リサーチ・クエスチョンを確定するレポートを書いています」'],
  ['条件', '範囲、制約、避けてほしいこと。「日本国内の事例に限る」「文献名は挙げないでください」'],
  ['出力の形式', '箇条書き、表、文字数、案の数。「300字以内」「10案を表で」'],
].forEach(([h, b], i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const x = 0.5 + col * 3.0, y = 1.4 + row * 1.78, w = 2.85, hh = 1.65;
  rect(x, y, w, hh, PAPER, { round: true });
  circleNum(i + 1, x + 0.15, y + 0.15, 0.38);
  txt(h, { x: x + 0.62, y: y + 0.15, w: w - 0.75, h: 0.38, fontSize: 15, bold: true, valign: 'middle' });
  txt(b, { x: x + 0.15, y: y + 0.62, w: w - 0.3, h: hh - 0.72, fontSize: 11, color: INK2, lineSp: 1.3, label: 'elem' });
});
card(6.6, 1.4, 2.9, 3.43, { fill: BLUETINT, head: '一度で終わらせない', body: '同じプロンプトでも出力は毎回少し違う。次の単語を確率で選んでいるから。\n\n続けて聞く：\n「なぜ」\n「他には」\n「反対の立場だと」\n「この案の弱点は」', bodySize: 12 });

// ============================================================
// 8 第3回 役割分担
// ============================================================
content(8, '第3回　先行研究の探索｜10月14日', 'AIは「どこを探すか」の見当、確認はデータベースで');
card(0.5, 1.4, 4.3, 1.6, { head: 'AIにしてもらうこと', body: '関連する学術分野と、検索キーワード（日本語と英語）の見当をつける。\n「文献名や著者名は挙げないでください」と条件に書く', bodySize: 12 });
card(0.5, 3.15, 4.3, 1.75, { head: '自分がすること', body: '学術データベースで実在と内容を確かめる', bodySize: 12, fill: BLUETINT });
['CiNii Research', 'Google Scholar', 'J-STAGE', '図書館の蔵書検索'].forEach((t, i) => chip(t, 0.68 + (i % 2) * 2.0, 4.1 + Math.floor(i / 2) * 0.4, 1.9, 0.34, { fontSize: 11, fill: 'FFFFFF' }));
txt('AIが挙げた文献の、見つけにくいずれ', { x: 5.2, y: 1.4, w: 4.3, h: 0.32, fontSize: 14, bold: true, color: BLUE });
[
  ['実在しない', 'それらしい著者、題名、年で作られる'],
  ['実在するが、内容が違う', '論文はあるが、テーマが自動車ではなく家電だった'],
  ['調査は実在するが、数値がない', '報告書を開くと、AIの示した数値が見当たらない'],
].forEach(([h, b], i) => {
  const y = 1.82 + i * 0.86;
  rect(5.2, y, 4.3, 0.76, REDTINT, { round: true });
  txt(h, { x: 5.38, y: y + 0.08, w: 3.95, h: 0.3, fontSize: 13, bold: true, color: RED });
  txt(b, { x: 5.38, y: y + 0.38, w: 3.95, h: 0.32, fontSize: 11, color: INK2 });
});
rect(5.2, 4.45, 4.3, 0.45, BLUETINT, { round: true });
txt('「本当ですか」と聞き直しても検証にはならない', { x: 5.38, y: 4.45, w: 4.0, h: 0.45, fontSize: 12, bold: true, color: BLUE, valign: 'middle' });

// ============================================================
// 9 第3回 文献リストの6項目
// ============================================================
content(9, '第3回　先行研究の探索｜10月14日', '文献リストの6項目と、AIが知らないこと');
txt('文献リストに揃える6項目', { x: 0.5, y: 1.4, w: 4.5, h: 0.32, fontSize: 14, bold: true, color: BLUE });
['著者', '発行年', '題名', '掲載誌または出版社', '巻号と頁', 'DOIまたはURL'].forEach((t, i) => {
  const col = i % 2, row = Math.floor(i / 2);
  numRow(i + 1, 0.5 + col * 2.35, 1.85 + row * 0.6, 2.3, t, null, { headSize: 12.5, bold: true });
});
card(0.5, 3.7, 4.4, 1.2, { head: '欠けやすいのは、年とURL', body: '検索した日と件数も記録しておくと、あとで再現できる。学術論文を少なくとも2件含める', bodySize: 11.5 });
[
  ['知識のカットオフ', '学習データはある時点まで。それ以降の出来事や最新の統計は知らない'],
  ['RAG（検索拡張生成）', '検索結果や文書を取り込んで回答する。最新情報を扱えるが、読み違いや誤った要約は起こる'],
  ['出典のURLは必ず開く', 'リンク先が存在しない、内容が違う、ということがある'],
].forEach(([h, b], i) => card(5.3, 1.4 + i * 1.18, 4.2, 1.05, { head: h, body: b, bodySize: 11, fill: i === 2 ? BLUETINT : PAPER }));

// ============================================================
// 10 第4回 論文の構造とAI要約のずれ
// ============================================================
content(10, '第4回　文献講読と批判的思考｜10月21日', '論文の構造と、AI要約の4つのずれ');
txt('論文の構造で読む', { x: 0.5, y: 1.4, w: 3, h: 0.32, fontSize: 14, bold: true, color: BLUE });
['問い', '方法', '結果', '考察', '限界'].forEach((t, i) => {
  const y = 1.82 + i * 0.56;
  chip(t, 0.5, y, 1.5, 0.44, { fontSize: 13 });
  if (i < 4) arrowDown(1.13, y + 0.44, 0.24, 0.12, RULE);
});
card(2.25, 1.4, 2.65, 3.5, { fill: BLUETINT, head: 'AI要約の使いどころ', body: '読む前の地図として使う。\n\n根拠は必ず原文で確かめる。原文の頁を書きながら照合する。\n\n要約は原文の代わりにならない', bodySize: 12 });
txt('AI要約に起きやすい4つのずれ', { x: 5.2, y: 1.4, w: 4.3, h: 0.32, fontSize: 14, bold: true, color: RED });
[
  ['問いを広げる', '「なぜ行かなくなったか」は論文の問いではない'],
  ['条件を落とす', '「購入経験者に限定」が抜ける'],
  ['結果を言い過ぎる', '「店に行かない」。後半の「商談は長い」が抜ける'],
  ['限界の節を見落とす', '原文には限界の節がある'],
].forEach(([h, b], i) => {
  const y = 1.82 + i * 0.78;
  rect(5.2, y, 4.3, 0.68, REDTINT, { round: true });
  txt(h, { x: 5.38, y: y + 0.06, w: 3.95, h: 0.3, fontSize: 13, bold: true, color: RED });
  txt(b, { x: 5.38, y: y + 0.35, w: 3.95, h: 0.3, fontSize: 11, color: INK2 });
});

// ============================================================
// 11 第4回 反論に答える
// ============================================================
content(11, '第4回　文献講読と批判的思考｜10月21日', '反論に、原文の根拠で答える');
txt('最も強い反論を3つ出させ、原文で判定する', { x: 0.5, y: 1.4, w: 4.5, h: 0.32, fontSize: 14, bold: true, color: BLUE });
[
  ['○', '原文の根拠（頁）で答えられる'],
  ['△', '一部は答えられるが、根拠が弱い'],
  ['×', '答えられない。論文の範囲の外'],
].forEach(([m, t], i) => {
  const y = 1.85 + i * 0.62;
  circleNum(m, 0.5, y, 0.44, i === 2 ? RED : BLUE);
  txt(t, { x: 1.1, y, w: 3.8, h: 0.44, fontSize: 13, valign: 'middle' });
});
rect(0.5, 3.85, 4.4, 1.05, BLUETINT, { round: true });
txt('答えられない反論が出るのは悪いことではない。「論文の範囲の外」と言えれば十分', { x: 0.68, y: 3.85, w: 4.05, h: 1.05, fontSize: 12.5, bold: true, color: BLUE, valign: 'middle', label: 'note11' });
['論文が答えていないこと', '自分の問いとの隙間', '問いを磨く材料（第5回へ）'].forEach((t, i) => {
  const y = 1.45 + i * 0.95;
  chip(t, 5.3, y, 4.2, 0.6, { fontSize: 13, fill: i === 2 ? BLUE : BLUETINT, color: i === 2 ? 'FFFFFF' : BLUE });
  if (i < 2) arrowDown(7.25, y + 0.62, 0.3, 0.28, RULE);
});
rect(5.3, 4.2, 4.2, 0.7, REDTINT, { round: true });
txt('全文を入力するときは、大学のAI利用規程と資料の利用条件を確認。要旨や必要な部分だけで足りることが多い', { x: 5.45, y: 4.2, w: 3.9, h: 0.7, fontSize: 10.5, color: RED, valign: 'middle', label: 'caution11' });

// ============================================================
// 12 第5回 査読者の4観点
// ============================================================
content(12, '第5回　リサーチ・クエスチョンの確定｜10月28日', '「厳しい査読者」の4観点で、補強か修正か');
['半年で調べて答えられるか', '先行研究に対する新しさがあるか', '範囲は適切か', '用語は明確か'].forEach((t, i) => numRow(i + 1, 0.5, 1.42 + i * 0.58, 4.3, t, null, { headSize: 13.5, bold: true }));
card(0.5, 3.68, 4.3, 1.22, { fill: BLUETINT, head: '指摘への応じ方', body: '答えられる指摘は「補強」、答えられない指摘は問いの「修正」。すべて受け入れる必要はなく、根拠を示して反論できた点は評価される', bodySize: 11, headSize: 12.5 });
txt('問いはこう変わる（例）', { x: 5.2, y: 1.4, w: 4.3, h: 0.32, fontSize: 14, bold: true, color: BLUE });
card(5.2, 1.8, 4.3, 1.2, { head: '修正前', headColor: INK3, body: '名古屋都市圏の20代にとって、オンラインでの情報収集は、ディーラーへの来店のしかたを「どう変えたのか」\n→ 過去との比較は半年では難しい', bodySize: 11 });
arrowDown(7.2, 3.03, 0.3, 0.22, BLUE);
card(5.2, 3.28, 4.3, 1.62, { fill: BLUETINT, head: '修正後（確定）', body: '車の購入を検討した名古屋都市圏の20代は、来店前にオンラインで何を調べ、何を目的にディーラーへ来店しているのか\n→ 現状を記述する。対象を絞り、「来店のしかた」を回数と目的に分けて定義', bodySize: 11 });

// ============================================================
// 13 第5回 1枚にまとめて発表、レポート
// ============================================================
content(13, '第5回　リサーチ・クエスチョンの確定｜10月28日', '1枚にまとめて発表し、確定レポートにする');
txt('1枚にまとめて、1人3分で発表', { x: 0.5, y: 1.4, w: 4.3, h: 0.32, fontSize: 14, bold: true, color: BLUE });
['問い', '背景と意義', '先行研究でわかっていること', '調べ方の見通し'].forEach((t, i) => chip(t, 0.5 + (i % 2) * 2.2, 1.82 + Math.floor(i / 2) * 0.58, 2.1, 0.5, { fontSize: 12 }));
card(0.5, 3.12, 4.3, 1.78, { head: '想定される質問', body: 'Q　20人のアンケートで「20代」と言えるのか\nA　一般化はできない。「傾向を探る調査」と位置づけ、限界として書く', bodySize: 11.5 });
rect(5.2, 1.4, 4.3, 3.5, BLUETINT, { round: true });
txt('確定レポートに含めるもの', { x: 5.45, y: 1.55, w: 3.9, h: 0.35, fontSize: 15, bold: true, color: BLUE });
txt('□ 確定したリサーチ・クエスチョンと、選んだ理由\n□ 背景と意義\n□ 先行研究の整理（実在を確認した文献リスト付き）\n□ 調べ方の見通し\n□ AI活用記録（AI利用ガイドの形式）\n□ 本文の冒頭または末尾に、AI利用を明記する一文',
  { x: 5.45, y: 2.0, w: 3.9, h: 2.8, fontSize: 12, lineSp: 1.35, paraSpaceAfter: 5, label: 'checks13' });

// ============================================================
// 14 3つの習慣
// ============================================================
content(14, '5回を貫く3つの習慣', '分担、検証、記録');
[
  ['分担', '広げるのはAI', '候補を出す、視点を増やす、反論を挙げる、言い換える\n\n選ぶ・確かめる・書くのは自分\n何を問うか、どの案を採用するか、事実か、自分の言葉でどう書くか'],
  ['検証', '一次資料で確かめる', '固有名詞、数値、文献、URL\n\n「根拠は」「出典は」「反対の見解は」と尋ね、示された資料を自分で開く'],
  ['記録', '使ったことを隠さない', 'AI活用記録（5項目）を提出物に添える\n\n本文の冒頭か末尾に、AI利用を明記する一文を入れる'],
].forEach(([h, sub, b], i) => {
  const x = 0.5 + i * 3.05, w = 2.9, y = 1.4, hh = 2.9;
  rect(x, y, w, hh, i === 0 ? BLUETINT : PAPER, { round: true });
  circleNum(i + 1, x + 0.18, y + 0.18, 0.44);
  txt(h, { x: x + 0.75, y: y + 0.18, w: w - 0.9, h: 0.44, fontSize: 20, bold: true, valign: 'middle' });
  txt(sub, { x: x + 0.18, y: y + 0.75, w: w - 0.36, h: 0.3, fontSize: 12.5, bold: true, color: BLUE });
  txt(b, { x: x + 0.18, y: y + 1.08, w: w - 0.36, h: hh - 1.2, fontSize: 11, color: INK2, lineSp: 1.3, label: 'habit' });
});
rect(0.5, 4.45, 9.0, 0.45, BLUE, { round: true });
txt('自分の考えを先に書いてからAIに批評させる。良い出力も「この案の弱点は」と聞いてから採用する', { x: 0.7, y: 4.45, w: 8.6, h: 0.45, fontSize: 12.5, bold: true, color: 'FFFFFF', valign: 'middle' });

// ============================================================
// 15 検証の表
// ============================================================
content(15, '習慣2　検証', '固有名詞、数値、文献、URLは一次資料で');
const rows15 = [
  ['確かめるもの', '確かめ方'],
  ['文献', 'CiNii Research、Google Scholar、J-STAGE、図書館の蔵書検索で、実在と内容の両方を確認する。確認できないものは使わない'],
  ['数値、統計', '公的統計や報告書の該当ページで、同じ数値であることを確かめる。確認できない数値は使わない'],
  ['固有名詞、年', '公式サイトや原資料で確認する。自分がよく知っている話題でもずれることがある'],
  ['出典のURL', '必ず開き、そのページに回答の内容が本当に書かれているかを見る'],
  ['主張', '「根拠は」「出典は」「反対の見解は」「例外は」と尋ね、示された資料を自分で確認する。自分で説明できない内容は使わない'],
].map((r, ri) => r.map((c, ci) => ({ text: c, options: {
  fontFace: FONT, fontSize: ri === 0 ? 12 : 12, bold: ri === 0 || ci === 0, color: ri === 0 ? BLUE : (ci === 0 ? INK : INK2),
  fill: { color: ri === 0 ? BLUETINT : 'FFFFFF' }, valign: 'middle', margin: [4, 8, 4, 8] } })));
rows15.slice(1).forEach(r => checkFit('slide 15 table', [r[1].text], 6.6, 0.56, 12, 1.25));
curSlide.addTable(rows15, { x: 0.5, y: 1.4, w: 9.0, colW: [1.7, 7.3], rowH: [0.36, 0.62, 0.62, 0.62, 0.5, 0.62], border: { type: 'solid', pt: 0.5, color: RULE } });

// ============================================================
// 16 記録
// ============================================================
content(16, '習慣3　記録', 'AI活用記録と、本文での明記');
txt('AI活用記録の5項目（提出物に添付）', { x: 0.5, y: 1.4, w: 4.5, h: 0.32, fontSize: 14, bold: true, color: BLUE });
[
  ['使用したツール', 'Copilot（大学提供）'],
  ['使用した場面', '問いの候補出し、要約、反論の生成、推敲など'],
  ['主なプロンプト', '実際に使った指示を1〜3個'],
  ['出力をどう扱ったか', '採用、修正、不採用と、その理由'],
  ['検証したこと', '事実、文献、数値を、何で確かめたか'],
].forEach(([h, b], i) => numRow(i + 1, 0.5, 1.82 + i * 0.62, 4.4, h, b, { headSize: 13, bodySize: 10.5, bodyH: 0.3 }));
card(5.2, 1.4, 4.3, 2.4, { fill: BLUETINT, head: '本文の冒頭または末尾に入れる一文（例）', headSize: 12.5, body: '本レポートの作成にあたり、問いの候補出しと構成案の検討に大学が提供する Copilot を使用した。出力はすべて筆者が一次資料に基づいて検証し、文章は筆者自身が執筆した。', bodySize: 12 });
card(5.2, 3.95, 4.3, 0.95, { head: 'AI活用の透明性は、評価の対象', headSize: 12.5, body: '出力を引用するときは、プロンプトと日付がわかる形で', bodySize: 11 });

// ============================================================
// 17 後半へ
// ============================================================
content(17, '第6回から', '後半の店舗設計提案プロジェクトへ');
txt('前半で身につけた型を、グループで使う', { x: 0.5, y: 1.4, w: 6, h: 0.32, fontSize: 14, bold: true, color: BLUE });
['広げる', '絞る', '批判させる', '記録する'].forEach((t, i) => {
  const w = 1.9, x = 0.5 + i * (w + 0.47);
  chip(t, x, 1.82, w, 0.55, { fontSize: 14, fill: i === 3 ? BLUE : BLUETINT, color: i === 3 ? 'FFFFFF' : BLUE });
  if (i < 3) arrowRight(x + w + 0.1, 1.95, 0.27, 0.28, RULE);
});
[
  ['ペルソナを作る', 'AIで複数の案を出し、現実味と課題との関連で取捨選択する'],
  ['コンセプトを出す', '案を広げ、絞り、弱点を指摘させる'],
  ['提案を仕上げる', 'グループでもAI活用記録を続ける'],
].forEach(([h, b], i) => card(0.5 + i * 3.05, 2.6, 2.9, 1.1, { head: h, body: b, bodySize: 11 }));
rect(0.5, 3.9, 9.0, 1.0, REDTINT, { round: true });
txt('企業から提供された非公開の情報の取り扱い', { x: 0.7, y: 3.98, w: 8.6, h: 0.3, fontSize: 13, bold: true, color: RED });
txt('AIに入力しない　・　SNSや外部に書かない　・　グループ外に共有しない　（判断に迷うときは担当教員に確認。第6回で確認します）', { x: 0.7, y: 4.3, w: 8.6, h: 0.5, fontSize: 12, color: INK, label: 'caution17' });

// ============================================================
// 18 プロンプト集
// ============================================================
content(18, '付録A', '使ってみたいAIプロンプト集（28個）');
[
  ['基本の型と、対話の続け方', '3'], ['問いを立てる（第2回）', '4'], ['先行研究を探す（第3回）', '5'], ['読む、批判的に考える（第4回）', '6'],
  ['問いを磨き、確定する（第5回）', '4'], ['レポートに仕上げる', '3'], ['検証のためのプロンプト', '3'],
].forEach(([t, n], i) => {
  const y = 1.4 + i * 0.4;
  rect(0.5, y, 3.9, 0.34, PAPER, { round: true });
  txt(t, { x: 0.65, y, w: 3.0, h: 0.34, fontSize: 11.5, bold: true, valign: 'middle' });
  txt(n + '個', { x: 3.6, y, w: 0.7, h: 0.34, fontSize: 11.5, color: BLUE, bold: true, align: 'right', valign: 'middle' });
});
card(0.5, 4.25, 3.9, 0.65, { fill: BLUETINT, body: '各プロンプトに「解説」「続けて聞く」「確かめる」を添付。（　）は自分の内容に置き換える', bodySize: 10.5 });
rect(4.7, 1.4, 4.8, 3.5, PAPER, { round: true });
txt('例　19　査読者として批判させる　［授業で使用］', { x: 4.9, y: 1.5, w: 4.4, h: 0.3, fontSize: 11, bold: true, color: BLUE });
txt('あなたは厳しい査読者です。次のリサーチ・クエスチョンについて、\n1. 半年で調べて答えられるか\n2. 先行研究に対する新しさがあるか\n3. 範囲は適切か\n4. 用語は明確か\nの観点から問題点を指摘し、それぞれ改善案を示してください。\n問い：（あなたの問い）\n背景：（200字程度）\n主な先行研究：（3件）',
  { x: 4.9, y: 1.85, w: 4.4, h: 2.3, fontSize: 10, lineSp: 1.25, label: 'prompt18' });
txt('続けて聞く：「指摘1を、問いを変えずに補強する方法は」\n確かめる：すべて受け入れなくてよい。修正前後の問いを記録する', { x: 4.9, y: 4.22, w: 4.4, h: 0.65, fontSize: 10, color: INK2, lineSp: 1.25, label: 'prompt18-notes' });

// ============================================================
// 19 おわりに
// ============================================================
slideNo = 19; curSlide = pres.addSlide({ masterName: 'DARK' });
txt('おわりに', { x: 0.7, y: 0.55, w: 8.6, h: 0.5, fontSize: 14, bold: true, color: BLUETINT });
txt('この4つを守れば、AIは思考のパートナーになる', { x: 0.7, y: 0.95, w: 8.6, h: 0.6, fontSize: 24, bold: true, color: 'FFFFFF' });
['問いは自分が持つ', '出力は素材であって完成品ではない', '事実は一次資料で確かめる', '使ったことを隠さない'].forEach((t, i) => {
  const y = 1.8 + i * 0.62;
  circleNum(i + 1, 0.7, y, 0.44, 'FFFFFF', BLUE);
  txt(t, { x: 1.3, y, w: 8.0, h: 0.44, fontSize: 18, bold: true, color: 'FFFFFF', valign: 'middle' });
});
txt('冊子（PDF）：授業サイト「第5回 リサーチ・クエスチョンの確定」のページから\nkklab.mobi/kklab-basic-seminar/sessions/05-rq-final/\n前半の提出物の書式と締切は、授業内で案内します', { x: 0.7, y: 4.35, w: 8.6, h: 0.9, fontSize: 11.5, color: BLUETINT, lineSp: 1.35 });
curSlide.addNotes(narr[19]); addAudio(19);

// ------------------------------------------------------------
pres.writeFile({ fileName: path.join(HERE, 'summary-01-05-slides.pptx') }).then(f => {
  console.log('wrote', f);
  if (warnings.length) { console.log('--- fit warnings ---'); warnings.forEach(w => console.log(w)); } else console.log('fit: ok');
});
