# 基礎演習Ⅱ 授業サイト

名古屋市立大学 経済学部「基礎演習Ⅱ」（2026年度後期・水曜3限 13:00〜14:30・担当 河合勝彦）の授業サイトです。
Jekyll で作られており、GitHub Pages がリポジトリの内容から自動でサイトを生成します。

公開URL: https://kklab.mobi/kklab-basic-seminar/

## よくある更新

### お知らせを載せる

`_data/announcements.yml` の先頭に追記します（新しいものが上）。

```yaml
- date: 2026-10-07
  text: 第3回の資料を掲載しました。
  url: /sessions/03-literature-search/   # 任意。外部サイトは https:// から
```

### 日程を入れる

`_data/schedule.yml` の `date` に `YYYY-MM-DD` で書きます。日程を入れると、トップページの「次回の授業」と「15回の流れ」、授業計画ページの「次回」「終了した回」の表示が自動で切り替わります。時限や教室が通常と違う回は `period` / `room` を追加します。

```yaml
- number: 3
  date: 2026-10-07
  room: 3-101        # 教室が 2-204 でないときだけ
```

### 各回の資料（スライドやPDF）を載せる

1. ファイルを `materials/` に置く（例: `materials/03-slides.pdf`）
2. その回のファイル（`_sessions/03-literature-search.md`）の `materials:` に追記する

```yaml
materials:
  - title: スライド
    url: /materials/03-slides.pdf
  - title: ワークシート
    url: https://example.com/worksheet
    note: 授業中に配布
```

### 各回の内容を書き換える

`_sessions/` の各ファイルが1回分です。先頭の設定（`number`, `title`, `phase`, `tag`, `summary`, `note`）が授業計画ページやトップに使われ、それ以下の本文（Markdown）が各回のページ本文になります。プロンプト例は次の形で書きます。

```html
<figure class="prompt">
<figcaption>見出し</figcaption>
<pre>プロンプトの本文</pre>
</figure>
```

### 出席確認クイズの問題を直す

`_attendance/` の各ファイルが1回分です（`01.md` が第1回）。`questions:` に5問を書きます。`answer` は正解の選択肢の番号（0 が最初の選択肢）で、選択肢の表示順はページを開くたびに入れ替わります。不正解のときは `explain` の解説が表示されます。

```yaml
- q: 問題文
  choices:
  - 選択肢1
  - 選択肢2
  - 選択肢3
  - 選択肢4
  answer: 0
  explain: 解説
```

学生には https://kklab.mobi/kklab-basic-seminar/attendance/ から各回のクイズに回答してもらいます。全問正解すると、サーバー時刻と確認コードの付いた出席カードが表示されるので、そのスクリーンショットを提出させます（提出先は `_data/course.yml` の `attendance_submit_to`）。

確認コードは、学籍番号・氏名・提出時間から計算されます。教員用の照合ページ https://kklab.mobi/kklab-basic-seminar/verify/ にスクリーンショットの内容を転記すると、改ざんの有無を判定できます。`assets/js/code.js` の計算方法や鍵を変えると、それまでに提出されたコードは照合できなくなります。

### 科目情報・連絡先を変える

`_data/course.yml` を編集します。教室、曜日・時限、メールアドレス、オフィスアワーなどはここから全ページに反映されます。

## 構成

```
_config.yml           サイト設定（タイトル、URL）
_data/course.yml      科目情報・連絡先
_data/schedule.yml    各回の日程
_data/announcements.yml  お知らせ
_data/phases.yml      3つのフェーズの名前と説明
_data/nav.yml         ヘッダーのメニュー
_sessions/            各回のページ（01〜15）
_attendance/          出席確認クイズの問題（01〜15）
attendance/index.html 出席確認の一覧
verify.html           出席カードの照合（教員用）
index.html            トップページ
schedule.html         授業計画
syllabus.md           シラバス
ai-guide.md           AI利用ガイド
grading.md            成績評価
contact.md            連絡先
_layouts/ _includes/  ページの骨組み
assets/css/style.css  デザイン
assets/js/main.js     「次回」の自動判定
assets/js/attendance.js  クイズの採点と出席カードの表示
assets/js/code.js     確認コードの計算（クイズと照合ページで共用）
materials/            配布資料の置き場
```

## GitHub Pages の設定（初回のみ）

1. リポジトリの Settings → Pages を開く
2. Source を「Deploy from a branch」、Branch を `main` / `/ (root)` にして保存
3. 数分後に https://kklab.mobi/kklab-basic-seminar/ で公開される

以後は `main` に push するたびに自動で更新されます。

## ローカルで確認する（任意）

Ruby と Bundler がある環境で次を実行すると、http://localhost:4000/kklab-basic-seminar/ で確認できます。

```sh
bundle install
bundle exec jekyll serve
```
