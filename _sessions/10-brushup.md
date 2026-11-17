---
number: 10
title: 提案のブラッシュアップ
phase: 2
tag: AIによる批判的検証
summary: 企業のニーズに対し、提案の弱点をAIに指摘させ補強する
materials:
  - title: ワークシートの解説　動画（MP4、約7分）
    url: /materials/worksheets/slides/video/#s10
    note: ワークシートの各欄の書き方を、記入例とともにスライドとナレーションで説明する
  - title: ワークシートの解説　音声付きスライド（PowerPoint）
    url: /materials/worksheets/slides/worksheet-10-slides.pptx
    note: スライドショーを開始すると音声が自動で流れ、次のスライドへ進む
  - title: グループ活動記録シート（PDF、12ページ）
    url: /materials/group-log/group-log.pdf
    note: 第6〜12回のグループの話し合いごとに、出欠、成果、自分の役割、気づきを個人で記録する。第13回で全体を振り返る
  - title: グループ活動記録シート（Word）
    url: /materials/group-log/group-log.docx
    note: PDF 版と同じ内容。Word に直接入力して提出する
  - title: グループ活動記録シート（Markdown、記入用）
    url: /materials/group-log/group-log.md
    note: PDF 版と同じ内容。テキストエディタや Teams などで記入して提出する
  - title: グループ活動記録シート　汎用版（PDF、1ページ）
    url: /materials/group-log/group-log-simple.pdf
    note: 授業や日程に依存しない1ページの版。ほかのグループプロジェクトでも、話し合いのたびに1枚ずつ記入して使える
  - title: グループ活動記録シート　汎用版（Word）
    url: /materials/group-log/group-log-simple.docx
    note: PDF 版と同じ内容。回ごとにファイルをコピーして使う
  - title: グループ活動記録シート　汎用版（Markdown、記入用）
    url: /materials/group-log/group-log-simple.md
    note: PDF 版と同じ内容。回ごとにファイルをコピーして使う
---
## この回でやること

提案を企業の視点で検証します。AIに経営企画や店舗運営の担当者の役を与えて弱点を指摘させ、コスト、実現性、ブランドとの整合、法規や安全といった観点から補強します。グループ間で中間発表を行い、互いにフィードバックします。

## ねらい

- 企業が提案を見るときの観点（課題への適合、実現性、費用対効果、ブランド整合、運用負荷）を挙げられる
- AIに批判者の役割を与え、提案の弱点を体系的に洗い出せる
- 根拠となるデータの必要性を見極め、AIの数値ではなく一次資料で補強できる

## 進め方

1. ワーク1：提案の要点を1枚にまとめ、AIに企業側の立場から批判させる
2. ワーク2：指摘を「対応する」「根拠を示して反論する」「割り切る」に分類し、対応を決める
3. 中間発表：各グループ5分。他のグループから質問と指摘をもらう
4. 指摘をもとに提案を改訂する

## プロンプト例

<figure class="prompt">
<figcaption>企業側の視点で批判させる</figcaption>
<pre>あなたは自動車ディーラーを運営する会社の経営企画担当者です。
次の店舗設計提案について、課題への適合、初期費用と運営費、実現の難しさ、ブランドイメージとの整合、スタッフの運用負荷、安全と法規の観点から、厳しく問題点を指摘してください。指摘は重要な順に並べてください。
提案：（要約）</pre>
</figure>

<figure class="prompt">
<figcaption>根拠の不足を見つける</figcaption>
<pre>この提案の中で、根拠（データ、事例、調査）が示されていない主張を列挙し、それぞれについて、どんな一次資料で裏づけられそうかを提案してください。
数値やデータそのものは生成しないでください。</pre>
</figure>

## 次回までに

- 提案の改訂版（指摘への対応を明記）
- 発表の構成案のたたき台
