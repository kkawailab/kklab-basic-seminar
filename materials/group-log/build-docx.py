#!/usr/bin/env python3
"""グループ活動記録シートの Word 版（group-log.docx）を作る。

内容とページ構成は group-log.tex（PDF 版）と同じ。体裁の部品（色、表、記入欄、
フッター）は各回ワークシートの Word 版と共通で、../worksheets/tex2docx.py の関数を使う。
group-log.tex を直したときは、このファイルの対応する箇所も直すこと。

  python3 build-docx.py              group-log.docx を作る（python-docx が必要）
  uv run --with python-docx python3 build-docx.py
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "worksheets"))

import tex2docx as W  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Mm  # noqa: E402

W.Ctx.phase, W.Ctx.tint = W.PHASES["2"][0], W.PHASES["2"][1]
W.FOOTER_LEFT = "基礎演習Ⅱ　2026年度後期　グループ活動記録シート"
TW = W.TEXT_W

SESSIONS = [
    ("6", "11月4日（水）", "産学連携キックオフ"),
    ("7", "11月11日（水）", "コンセプト立案"),
    ("8", "11月18日（水）", "店舗設計・空間構成"),
    ("9", "11月25日（水）", "動線とUXのデザイン"),
    ("10", "12月2日（水）", "提案のブラッシュアップ"),
    ("11", "12月9日（水）", "プレゼン資料作成"),
    ("12", "2027年1月13日（水）", "最終プレゼン発表会"),
]
SPARE_SUB = "グループ活動記録　｜　第　　回　　　月　　　日（　　）　□ 授業内　□ 授業外"


# ------------------------------------------------------------
# 部品
# ------------------------------------------------------------
def header(doc, sub, title):
    t, cell = W.box(doc, TW, fill=W.Ctx.phase, border=(W.Ctx.phase, 4))
    W.set_cell_margins(cell, 2.5, 2.5, 4, 4)
    W.cell_paragraph(cell, [(sub, {})], {"gothic": True, "size": 9, "color": "FFFFFF"})
    W.cell_paragraph(cell, [(title, {})], {"gothic": True, "bold": True, "size": 18, "color": "FFFFFF"})
    W.name_line(doc, [("学籍番号", 50), ("氏名", 110), ("グループ", 160)], before=4, after=3)


def sec(doc, label, title, before=6):
    W.para(doc, [("■ " + label, {"color": W.Ctx.phase}), ("　", {}), (title, {})],
           {"gothic": True, "bold": True, "size": 10}, before=before, after=2, keep=True)


def line(doc, frags, size=9, after=2):
    return W.para(doc, frags, {"size": size}, after=after, keep=True)


def note(text):
    return (text, {"size": 8, "color": W.INK2})


def table(doc, spec, body, head=True):
    W.render_table(doc, TW, spec, body, head)


def new_page(doc, footer_right, first=False):
    section = doc.sections[0] if first else doc.add_section(WD_SECTION.NEW_PAGE)
    W.set_footer(section, footer_right)


# ------------------------------------------------------------
# ページ
# ------------------------------------------------------------
def page_intro(doc):
    new_page(doc, "使い方・名簿・ルール", first=True)
    header(doc, "基礎演習Ⅱ　2026年度後期　｜　実践プロジェクト（第6〜13回）",
           "産学連携プロジェクト　グループ活動記録シート")
    t, cell = W.box(doc, TW, fill=W.Ctx.tint, border=(W.Ctx.tint, 4))
    W.cell_paragraph(cell, [("このシートの使い方", {})], {"gothic": True, "bold": True, "size": 9, "color": W.Ctx.phase})
    items = [
        [("グループで話し合った回ごとに、授業の最後の5分で1枚（3〜9ページ）を", {}), ("個人で", {"bold": True}),
         ("記入する。授業外に集まったときなどは予備のページ（10・11ページ）を使う。", {})],
        [("毎回、①その日の出欠、②ディスカッションで得られた成果、③自分の役割と実際にやったこと、④気づき、を書く。"
          "成果は「話した」ではなく「何が決まったか、何ができたか」で書く。", {})],
        [("欠席・遅刻・早退した回も記入する。メンバーから共有を受けた内容と、追いつくためにしたことを書く。", {})],
        [("記入したら、2ページの「出席と役割の記録」に1行で転記する。第13回（プロジェクトの総括）で最後のページを書き、全体を振り返る。", {})],
        [("他のメンバーの出欠は、責めるためではなく、欠席者への共有と分担の見直しに使う。", {})],
        [("このWordファイルに直接入力するか、印刷して手で書く。", {})],
    ]
    for it in items:
        p = W.cell_paragraph(cell, [("・", {})] + it, {"size": 9})
        p.paragraph_format.left_indent = Mm(4)
        p.paragraph_format.first_line_indent = Mm(-3.5)
    W.spacer(doc)

    sec(doc, "名簿", "グループのメンバー（以後のページでは番号で書く）")
    table(doc, r"colspec={Q[c,wd=10mm,m] X[2,l,m] X[2,l,m] X[3,l,m]},row{2-Z}={ht=8mm}",
          r"番号 & 氏名 & 最初の役割 & 連絡方法・メモ \\ ① & & & \\ ② & & & \\ ③ & & & \\"
          r" ④ & & & \\ ⑤ & & & \\ ⑥ & & & \\")

    sec(doc, "ルール", "グループで決めた約束")
    table(doc, r"colspec={Q[l,wd=46mm,m] X[l,m]},row{1-Z}={ht=14mm}",
          r"\hc{欠席・遅刻の連絡\newline 誰に・いつまでに・何で} & \\"
          r"\hc{欠席者への共有のしかた\newline 誰が・何で・いつまでに} & \\"
          r"\hc{授業外の作業の進め方\newline 締め切り・集まり方} & \\"
          r"\hc{意見が分かれたときの決め方} & \\", head=False)


def page_list(doc):
    new_page(doc, "出席と役割の記録")
    header(doc, "グループ活動記録　｜　第6〜12回の一覧", "出席と役割の記録")
    W.render_instr(doc, "毎回の記録を書いたら、ここに1行で転記する。出欠は ○ 出席　△ 遅刻・早退　× 欠席。"
                        "欠席したメンバーは名簿の番号で書く。第13回の振り返りは、この一覧をもとに書く。")
    rows = "".join(r"%s & %s & & & & & \\" % (n, d) for n, d in
                   [("6", "11/4"), ("7", "11/11"), ("8", "11/18"), ("9", "11/25"), ("10", "12/2"),
                    ("11", "12/9"), ("12", "1/13"), ("予備", ""), ("予備", "")])
    table(doc, r"colspec={Q[c,wd=10mm,m] Q[c,wd=14mm,m] Q[c,wd=13mm,m] Q[c,wd=16mm,m] X[2,l,m] X[3,l,m] X[3,l,m]},row{2-Z}={ht=19mm}",
          r"回 & 日付 & 自分の\newline 出欠 & 欠席した\newline メンバー & 自分の役割 & グループの成果（一言） & 気づき（一言） \\" + rows)


def page_log(doc, sub, title):
    new_page(doc, title)
    header(doc, sub, title)

    sec(doc, "1", "出席", before=2)
    line(doc, [("自分：　□ 出席　□ 遅刻　□ 早退　□ 欠席　", {}), note("→ 出席以外のときは、次の2行も書く")])
    W.render_fillin(doc, TW, r"事前の連絡　□ した　□ していない\quad 共有してもらった内容と相手")
    W.render_fillin(doc, TW, "追いつくためにしたこと／グループのためにできること")
    line(doc, [("メンバーの出欠　", {}), note("番号は1ページの名簿の番号。○ 出席　△ 遅刻・早退　× 欠席")])
    table(doc, r"colspec={Q[l,wd=16mm,m] X[c,m] X[c,m] X[c,m] X[c,m] X[c,m] X[c,m]},row{2}={ht=7mm}",
          r"番号 & ① & ② & ③ & ④ & ⑤ & ⑥ \\ \hc{出欠} & & & & & & \\")
    W.render_fillin(doc, TW, "欠席者への共有（誰が・何で・何を）")

    sec(doc, "2", "ディスカッションの成果")
    W.render_fillin(doc, TW, "今日の目的・議題")
    W.render_answerbox(doc, TW, "グループで決まったこと・できあがったもの（成果物）", 18)
    W.render_twoboxes(doc, TW, "まだ決まっていないこと・次回に持ち越すこと", "意見が分かれた点と、どう決めたか", "13mm")

    sec(doc, "3", "自分の役割")
    line(doc, [("今日の役割：　□ 司会・進行　□ 記録　□ 調べ役　□ AI操作　□ 資料作成　□ 発表　□ 時間管理　□ その他（　　　　）", {})],
         size=8.5)
    W.render_answerbox(doc, TW, "実際にやったこと・グループへの貢献（具体的に）", 13)
    W.render_fillin(doc, TW, r"役割を果たせた度合い　1　・　2　・　3　・　4　・　5\quad 理由")

    sec(doc, "4", "気づき")
    W.render_twoboxes(doc, TW, "グループの進め方でうまくいったこと", "次に改善したいこと", "13mm")
    W.render_answerbox(doc, TW, "自分自身について気づいたこと（得意・苦手、メンバーから学んだこと）", 11)

    sec(doc, "5", "次回までの分担")
    table(doc, r"colspec={X[l,m] Q[l,wd=30mm,m] Q[l,wd=22mm,m] Q[c,wd=8mm,m]},row{2-Z}={ht=6mm}",
          r"やること & 担当 & 期限 & 済 \\ & & & □ \\ & & & □ \\ & & & □ \\")


def page_final(doc):
    title = "プロジェクト全体の振り返り"
    new_page(doc, title)
    header(doc, "グループ活動記録　｜　第13回　2027年1月20日（水）", title)

    sec(doc, "1", "出席の集計", before=2)
    line(doc, [("第6〜12回（全7回）のうち　出席＿＿＿回　遅刻・早退＿＿＿回　欠席＿＿＿回", {})])
    line(doc, [note("2ページの「出席と役割の記録」から数える。授業外の集まりへの参加は＿＿＿回")], after=3)

    sec(doc, "2", "自分の役割を振り返る")
    table(doc, r"colspec={Q[l,wd=50mm,m] X[l,m]},row{1-Z}={ht=14mm}",
          r"\hc{自分が担った役割は、回を追ってどう変わったか} & \\"
          r"\hc{グループの成果のうち、自分が最も貢献した部分と、その具体的な行動} & \\"
          r"\hc{果たせなかった役割、もっとできたこと。その理由} & \\"
          r"\hc{欠席や作業の遅れが出たとき、グループはどう対応したか。自分は何をしたか} & \\", head=False)

    sec(doc, "3", "気づきをまとめる")
    table(doc, r"colspec={Q[l,wd=50mm,m] X[l,m]},row{1-Z}={ht=14mm}",
          r"\hc{グループで成果を出すために大事だと気づいたこと} & \\"
          r"\hc{自分の強みと課題（チームの中での自分）} & \\"
          r"\hc{次にグループで活動するとき、自分が意識すること} & \\", head=False)

    sec(doc, "4", "メンバーの良かった行動")
    line(doc, [note("名簿の番号ごとに、助けられたこと・見習いたいことを具体的に書く（自分の欄は空けてよい）")])
    table(doc, r"colspec={Q[c,wd=10mm,m] X[l,m]},row{1-Z}={ht=6.5mm}",
          r"① & \\ ② & \\ ③ & \\ ④ & \\ ⑤ & \\ ⑥ & \\", head=False)


def main():
    doc = Document()
    W.setup(doc)
    page_intro(doc)
    page_list(doc)
    for n, d, t in SESSIONS:
        page_log(doc, "グループ活動記録　｜　第%s回　%s" % (n, d), t)
    for _ in range(2):
        page_log(doc, SPARE_SUB, "予備（追加の話し合い用）")
    page_final(doc)

    W.normalize(doc.element)
    for s in doc.sections:
        W.normalize(s.footer._element)
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None:
        zoom.set(qn("w:percent"), "100")
    doc.core_properties.title = "基礎演習Ⅱ 産学連携プロジェクト グループ活動記録シート"
    doc.core_properties.author = "河合 勝彦"
    out = HERE / "group-log.docx"
    doc.save(out)
    print(out.name)


if __name__ == "__main__":
    main()
