# 地域の困りごとからビジネスを創出する（記入式テンプレート）

河津桜のデジタル管理（静岡県南伊豆町）の新聞記事を題材に、身近な資産の困りごとから事業案を組み立てる発想法と、記入シートA〜C・記入例をまとめた資料です。第7回「コンセプト立案」のページ（`_sessions/07-concept.md`）から PDF にリンクしています。

| ファイル | 内容 |
|---|---|
| `business_creation_template_ja.tex` / `business_creation_template_ja.pdf` | 本文（記事の要約、発想の構造、45分ワークショップの手順、記入シートA〜C、記入例、運営上の落とし穴、発問） |
| `Makefile` | ビルド（`make` で PDF を作る） |

## 作り方

LuaLaTeX（TeX Live の ltjsarticle、luatexja-preset（Harano Aji）、tabularx、booktabs、hyperref）で組みます。

```sh
cd materials/business
make          # PDF を作る
make clean    # build/ を消す
```

中間ファイルは `build/` に入ります（Git には含めません）。記入欄を広くするときは、冒頭の `\field` の定義にある `13mm` を変えます。
