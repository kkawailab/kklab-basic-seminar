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
