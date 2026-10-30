# 各回ワークシートの解説（音声付きスライドと動画）

第1回〜第15回のワークシート（`../sessions/NN.tex`）について、各欄の書き方を記入例とともに説明する、
ナレーション付きのスライド（PowerPoint）と、それを動画にした MP4 です。授業サイトの動画ページ
（`video/index.md`、/materials/worksheets/slides/video/）と、各回のページの資料欄からリンクしています。

| ファイル | 内容 |
|---|---|
| `worksheet-NN-slides.pptx` | 第NN回の音声付きスライド（完成品）。スライドショーを開始すると各スライドの音声が自動で流れ、終わると次のスライドへ進む。ナレーションの原稿は各スライドのノートにも入っている |
| `video/worksheet-NN.mp4` | 第NN回の解説動画（1920×1080、スライドごとのチャプター付き） |
| `video/poster-NN.jpg` | 動画ページで再生前に表示する表紙（タイトルのスライド） |
| `video/index.md` | 授業サイトの動画ページ。`make_video.py` が書き出すので、手では直さない |
| `content/NN.json` | 第NN回のスライドの中身とナレーション原稿 |
| `audio/NN/MM.mp3`・`audio/NN/audio.json` | ナレーションの音声と、その長さ・原稿の指紋（どのスライドを作り直すかの判定に使う） |
| `build.js` | `content/NN.json` からスライドを組む（pptxgenjs） |
| `make_audio.py` | 原稿から音声を作る。edge-tts（Microsoft Edge の読み上げ音声、`ja-JP-NanamiNeural`）を使うのでネットワークが必要 |
| `finish.py` | テーマのフォントと色、音声の自動再生・自動送りを書き込む |
| `make_video.py` | 各スライドを静止画に描き、ナレーションを付けて MP4 にし、動画ページと表紙を書き出す。動画の文字は Noto Sans CJK JP で描く |

## 作り方

次の順に実行します。回の番号を渡すと、その回だけ作り直します（渡さなければ全回）。

```sh
cd materials/worksheets/slides
pip install edge-tts mutagen python-pptx Pillow imageio-ffmpeg && npm install pptxgenjs   # 初回のみ
python3 make_audio.py 03     # 音声（原稿を変えたスライドだけ作り直す。--force ですべて）
node build.js 03             # スライド本体
python3 finish.py 03         # 仕上げ（音声の自動再生、テーマ）
python3 make_video.py 03     # 動画、表紙、動画ページ
```

レイアウトだけ確かめたいときは、音声を作らずに `node build.js 03` と `python3 make_video.py --frames 03`
を実行し、`video/build/03/MM.png` を見ます。`build.js` は、文字が枠に収まりそうにないところを
`fit:` の行で知らせます（見積もりは控えめなので、`video/build/` の静止画で実際の収まりを確かめる）。

動画を作るには、ffmpeg（なければ `imageio-ffmpeg` が同梱の ffmpeg を使う）と Noto Sans CJK JP のフォント
（Debian/Ubuntu なら `fonts-noto-cjk`）が必要です。描いた静止画とスライドごとの動画は `video/build/` に残り、
git には入れません。

## 中身の書き方（`content/NN.json`）

```json
{
  "session": 3, "phase": 1, "title": "先行研究の探索", "tag": "AI検索の活用", "date": "10月14日（水）",
  "slides": [
    {"type": "title", "chapter": "タイトル", "narration": "…"},
    {"chapter": "動画のチャプター名", "kicker": "ワーク1　…", "heading": "スライドの見出し（1行）",
     "body": [ 部品, {"cols": [{"w": 4.3, "blocks": [部品, …]}, {"blocks": [部品, …]}]} ],
     "narration": "…"}
  ]
}
```

- `phase` はフェーズ（1 探究の基礎・2 実践プロジェクト・3 キャリア講習）で、色が変わる。表紙はメタ情報から自動で組む
- `body` は上から順に積む。`cols` は横に並べる列で、`w`（インチ）を省いた列が残りの幅を使う。余った高さは、最後の伸ばせる部品が受け取る（`"grow": false` で止める）
- 文字の中の `**…**` は太字（フェーズの色）、行頭の `・` と `□ ` はぶら下げの箇条書きになる
- ナレーションは読み上げられるので、英字の固有名詞はカタカナで書く（CiNii Research は「サイニィ・リサーチ」）。「頁」は「ページ」と書く

| 部品（`type`） | 用途と主な項目 |
|---|---|
| `card` | 見出し付きの枠。`head`、`body`、`tone`（`paper`・`tint`・`solid`・`bad`・`ok`）、`size` |
| `steps` | 番号の丸と見出し・説明の並び。`items: [[見出し, 説明], …]`、`head`、`marks`（丸の中の文字）、`plain`（見出しを細字に） |
| `checks` | □ のチェック項目の枠。`head`、`items` |
| `prompt` | プロンプト例。`caption`、`text` |
| `example` | 記入例（緑の枠）。`rows: [[項目, 例], …]`（`labelW` で項目の列幅）か `body` |
| `table` | 表。`head`、`rows`、`colW`（比）、`align` |
| `flow` | 矢印でつないだ流れ。`items`、`last`（最後を塗る） |
| `zones` | 平面のゾーニング図。`uw`×`uh` の升目に `zones: [[x, y, 幅, 高さ, 名前, "hi"/"bad"], …]`、`entry: [x, "入口"]` |
| `banner` | 1行の強調。`text`、`tone` |
| `label`・`text`・`stat` | 小見出し、地の文、大きな数字のタイル |

## 内容のもと

各回の `content/NN.json` は、次のファイルをもとに書いています。ワークシートや授業の内容を変えたときは、該当する回を直してください。

| 箇所 | もとにしたファイル |
|---|---|
| ねらい、欄の構成、プロンプト例 | `../sessions/NN.tex`、`_sessions/NN-*.md` |
| 記入例、注意点（「ポイント」） | `../answers/NN.tex`（記入例・模範解答版）。第6回〜第13回の課題は記入例のための架空の設定、第14・15回は「書き方の例」 |
| AI利用の考え方、出力を使う前のチェック、AI活用記録 | `ai-guide.md` |
| 第4回の記入例（論文の頁） | `../../papers/paper-x.tex`（架空論文Ｘ） |
