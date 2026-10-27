# 総まとめ冊子

前半「探究の基礎」（第1回〜第5回）の講義内容の総まとめと、巻末の「使ってみたいAIプロンプト集」「出席確認クイズで扱ったAIの用語」をまとめた冊子です。第5回のページ（`_sessions/05-rq-final.md`）から PDF にリンクしています。

| ファイル | 内容 |
|---|---|
| `summary-01-05.tex` / `summary-01-05.pdf` | 第1回〜第5回 総まとめ。第1部が各回の要点と5回を貫く習慣（約5ページ）、付録Aが解説付きのプロンプト集（28個）、付録Bが用語集 |
| `preamble.tex` | 冊子の共通レイアウト（jlreq、tcolorbox、tabularray）。色と部品はワークシート（`../worksheets/preamble.tex`）に合わせてある |
| `Makefile` | ビルド（`make` で `summary-*.tex` をすべて PDF にする） |

## 作り方

LuaLaTeX（TeX Live の jlreq、tcolorbox、tabularray、enumitem、needspace、fancyhdr、hyperref）で組みます。

```sh
cd materials/summary
make                  # すべての冊子
make summary-01-05.pdf
make clean            # build/ を消す
```

目次とプロンプト一覧の頁番号を揃えるため、Makefile は2回組みます。中間ファイルは `build/` に入ります（Git には含めません）。

## 音声付きスライド（`slides/`）

冊子の内容を19枚のスライドとナレーション（約14分）で振り返る PowerPoint です。スライドショーを開始すると各スライドの音声が自動で再生され、音声が終わると次のスライドへ進みます（通常の表示では右下のスピーカーの印をクリックして再生）。ナレーションの原稿は各スライドのノートにも入っています。

| ファイル | 内容 |
|---|---|
| `slides/summary-01-05-slides.pptx` | 音声付きスライド（完成品） |
| `slides/narration.json` | スライドの題とナレーション原稿、音声合成の声と速さ |
| `slides/make_audio.py` | 原稿から音声（`audio/NN.mp3`）と長さ（`audio/durations.json`）を作る。edge-tts（Microsoft Edge の読み上げ音声、`ja-JP-NanamiNeural`）を使うのでネットワークが必要 |
| `slides/build.js` | スライドの中身とレイアウト（pptxgenjs） |
| `slides/finish.py` | テーマのフォントと色、音声の自動再生・自動送りを書き込む |
| `slides/make_video.py` | 各スライドを静止画に描き、ナレーションを付けて MP4 動画にする。`video/NN.mp4`（スライドごと）と `video/summary-01-05-video.mp4`（通し、スライドごとのチャプター付き）を書き出す。動画の文字は Noto Sans CJK JP で描く |
| `slides/video/index.md` | 授業サイトの動画ページ（通し動画の再生と、スライドごとの一覧）。原稿を変えて長さが変わったら、表の長さも直す |

作り直すときは、次の順に実行します。原稿だけ直したときは `make_audio.py` に回の番号を渡すと、その音声だけ作り直せます。

```sh
cd materials/summary/slides
pip install edge-tts mutagen python-pptx Pillow imageio-ffmpeg && npm install pptxgenjs   # 初回のみ
python3 make_audio.py        # 音声（python3 make_audio.py 3 7 で第3・7スライドだけ）
node build.js                # スライド本体
python3 finish.py            # 仕上げ（音声の自動再生、テーマ）
python3 make_video.py        # 動画（python3 make_video.py 18 で第18スライドだけ描き直し。通し動画も作り直す）
```

スライドの本文を変えたときは、`build.js` の該当箇所と `narration.json` の原稿を合わせて直してください。

動画を作るには、ffmpeg（なければ `imageio-ffmpeg` が同梱の ffmpeg を使う）と Noto Sans CJK JP のフォント（Debian/Ubuntu なら `fonts-noto-cjk`）が必要です。描いた静止画は `video/frames/` に残り、git には入れません。

## 内容のもと

冊子の内容は、次のファイルをもとに書いています。授業の内容を変えたときは、該当する箇所を直してください。

| 冊子の箇所 | もとにしたファイル |
|---|---|
| 各回の要点（第1部 2節） | `_sessions/01〜05`、`materials/worksheets/sessions/01〜05.tex`、同 `answers/01〜05.tex`（記入例の問いの変遷） |
| AI利用の4つの考え方、出力を使う前のチェック、AI活用記録、プロンプトの4要素 | `ai-guide.md` |
| 提出物のチェックリスト、評価で見ること | `_sessions/05-rq-final.md`、`grading.md` |
| 出席クイズの用語（付録B）と各回の「クイズでは」の段落 | `_attendance/01〜05.md` |
| 「授業で使用」の印があるプロンプト | `_sessions/02〜05` のプロンプト例（文言はそのまま） |

## 新しい冊子を足すとき

`summary-06-13.tex` のように `summary-` で始まる名前で置き、冒頭で `\input{preamble}` したうえで `\doctitle` と `\docdate` を `\renewcommand` してください。`make` が自動で拾います。付録の節番号（付録A、A.1）に切り替えるには、付録の最初で `\appendixmode` を呼びます。
