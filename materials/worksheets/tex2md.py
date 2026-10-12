#!/usr/bin/env python3
"""LuaLaTeX 版ワークシートの原稿（sessions/NN.tex）から、受講生が記入する Markdown 版を作る。

原稿は LaTeX 版・Word 版・Forms 版と共通。記入欄は「：」で終わる行にし、その後ろ
（長い答えは次の行から）に書いてもらう。表はできるだけ使わず、項目ごとの箇条書きに
読み替える。短い答え（○×や点数）だけの表は Markdown の表のままにする。

  記入欄 \\fillin・\\answerbox・\\writelines → **見出し**：
  左の列が項目名の表                        → - **項目名**：（書く列が複数なら項目ごとに小項目）
  空欄だけの表（候補を並べる表）            → 1. **1列目**： と、残りの列の小項目を、紙の行数ぶん
  チェック項目・次回までに・注意の□         → - [ ] 項目
  プロンプト例                              → 引用（> ）
  方眼（スケッチ）                          → 画像を貼る欄

  python3 tex2md.py            全回分を markdown/ に書き出す
  python3 tex2md.py 03         第3回だけ
"""
import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from tex2docx import (AIMEMO, PERSONA, PHASES, items_of, parse_blocks,  # noqa: E402
                      parse_colspec, split_top, strip_comments)
from tex2forms import COURSE, blank, split_label, text_of, with_sub  # noqa: E402

OUT = HERE / "markdown"
HOWTO = ("> 記入のしかた：「：」の後ろに書く（長い答えは次の行から書いてよい）。"
         "チェック項目は `- [ ]` を `- [x]` にする。AIの出力はそのまま写さず、要点を自分の言葉で書く。")


def label_of(raw):
    return with_sub(*split_label(raw))


class MD:
    def __init__(self):
        self.L = []
        self.work = ""

    # --- 行の出力 ---
    def line(self, s):
        self.L.append(s)

    def para(self, s):
        self.gap()
        self.L.append(s)
        self.L.append("")

    def gap(self):
        if self.L and self.L[-1] != "":
            self.L.append("")

    def field(self, label):
        """ブロックの記入欄（答えは次の行から）"""
        self.para("**%s**：" % label)
        self.L.append("")

    # --- 表 ---
    def table(self, spec, body, header):
        cs = re.search(r"colspec=\{(.*?)\}\s*(,|$)", spec, re.S)
        cols = parse_colspec(cs.group(1))
        short = [c["fixed"] is not None for c in cols]
        rows = [r for r in split_top(body, r"\\") if r.strip()]
        rows = [[re.sub(r"^\s*\\SetRow\{[^}]*\}", "", c).strip() for c in split_top(r.strip(), "&")]
                for r in rows]
        hdr = [label_of(c) for c in rows[0]] if header else None
        data = rows[1:] if header else rows
        ncol = len(cols)
        data = [list(r) + [""] * (ncol - len(r)) for r in data]
        # 全行に同じ文字が入っている列（「分」など）は単位扱いで見出しに折り込む
        unit = set()
        for c in range(ncol):
            vals = [text_of(r[c]) for r in data]
            if len(data) > 1 and all(vals) and len(set(vals)) == 1:
                unit.add(c)
                if hdr:
                    hdr[c] = "%s（%s）" % (hdr[c], vals[0])

        def inputs_of(r):
            return [c for c in range(ncol) if blank(r[c]) or c in unit]

        def labels_of(r):
            return [c for c in range(ncol) if c not in inputs_of(r)
                    and not re.fullmatch(r"[\d０-９]+", text_of(r[c]))]

        # 同じ見出しが2組ある表（キーワードを2列ずつ並べた表）は、列ごとの記入欄にする
        if hdr and len(set(h for h in hdr if h)) < len([h for h in hdr if h]) \
                and not any(labels_of(r) for r in data):
            self.gap()
            for h in dict.fromkeys(h for h in hdr if h):
                self.line("- **%s**：" % h)
            self.L.append("")
            return

        # 短い答えの列が2つ以上で、長い答えの列が1つ以下なら Markdown の表
        all_inputs = sorted({c for r in data for c in inputs_of(r)})
        longs = [c for c in all_inputs if not short[c]]
        if hdr and len(longs) <= 1 and len(all_inputs) - len(longs) >= 2:
            self.gap()
            self.line("| " + " | ".join(hdr) + " |")
            self.line("|" + "---|" * ncol)
            for r in data:
                cells = ["" if c in inputs_of(r) else text_of(r[c], nl=" ") for c in range(ncol)]
                self.line("| " + " | ".join(cells) + " |")
            self.L.append("")
            return

        self.gap()
        n = 0
        lead_done = False
        prev_labeled = False
        for r in data:
            inputs = inputs_of(r)
            labels = labels_of(r)
            nums = [text_of(r[c]) for c in range(ncol) if c not in inputs and c not in labels]
            if not inputs:   # 空欄のない行：記入例か、穴埋めの型
                t = [text_of(r[c], nl=" ") for c in range(ncol) if not blank(r[c])]
                if any("＿" in x or "①" in x for x in t):
                    self.line("- **%s**：%s" % (t[0], "　".join(t[1:])))
                else:
                    self.para("　".join(t))
                continue
            names = [hdr[c] if hdr else "" for c in inputs]
            if labels:
                prev_labeled = True
                title = "　".join(label_of(r[c]) for c in labels)
                if len(inputs) == 1:
                    if hdr and names[0] and not lead_done:
                        self.line("**%s**" % names[0])
                        self.line("")
                        lead_done = True
                    self.line("- **%s**：" % title)
                else:
                    self.line("- **%s**" % title)
                    for nm in names:
                        self.line("  - %s：" % (nm or "　"))
            else:
                if prev_labeled:
                    self.L.append("")
                    prev_labeled = False
                n = int(nums[0]) if nums and nums[0].isdigit() else n + 1
                first, rest = names[0], names[1:]
                mark = "%d. " % n
                self.line("%s**%s**：" % (mark, first or self.work))
                for nm in rest:
                    self.line(" " * len(mark) + "- %s：" % nm)
        self.L.append("")

    # --- ブロック列 ---
    def blocks(self, nodes):
        for n in nodes:
            if n[0] == "para":
                t = text_of(n[1], nl=" ")
                if "＿" in t:
                    head, sep, rest = t.partition("：")
                    self.para("**%s**：%s" % (head, rest) if sep else "**%s**：" % t)
                else:
                    self.para(t)
            elif n[0] == "cmd":
                self.cmd(n[1], n[2], n[3])
            else:
                self.env(n[1], n[2], n[3], n[4])

    def cmd(self, name, opts, args):
        if name == "session":
            num, ph, date, title, tag, group = args
            self.line("# 第%s回　%s" % (num, text_of(title)))
            self.line("")
            head = "%s　第%s回　%s｜%s" % (COURSE, num, date, PHASES[ph][2])
            if tag.strip():
                head += "｜" + text_of(tag)
            self.para(head)
            self.line("- 学籍番号：")
            self.line("- 氏名：")
            if group.strip():
                self.line("- グループ：")
            self.line("")
            self.para(HOWTO)
        elif name == "goals":
            self.gap()
            self.line("**この回のねらい**")
            self.line("")
            for it in items_of(args[0]):
                self.line("- " + text_of(it, nl=" "))
            self.line("")
        elif name == "work":
            self.work = text_of(args[1], nl=" ")
            self.para("## %s　%s" % (text_of(args[0]), self.work))
        elif name == "instr":
            self.para(text_of(args[0], nl=" "))
        elif name == "answerbox":
            self.field(text_of(opts[0], nl=" ") if opts[0] else self.work)
        elif name == "twoboxes":
            self.field(text_of(args[0], nl=" "))
            self.field(text_of(args[1], nl=" "))
        elif name == "gridbox":
            t = text_of(opts[0], nl=" ") if opts[0] else self.work
            self.para("**%s**：（手描きのスケッチを撮影して、ここに画像を貼る）" % t)
            self.line("")
        elif name == "writelines":
            self.field(self.work)
        elif name == "fillin":
            self.para("**%s**：" % text_of(args[0], nl=" "))
        elif name == "aimemo":
            self.work = "AI活用記録"
            self.para("## AI活用記録（この回の分）")
            self.blocks(parse_blocks(AIMEMO))
        elif name == "personacard":
            rows = [r for r in split_top(parse_blocks(PERSONA)[0][4], r"\\") if r.strip()]
            self.para("### ペルソナ%s" % args[0])
            self.line("- **名前**：")
            for r in rows[1:]:
                raw = re.sub(r"^\s*\\SetRow\{[^}]*\}", "", split_top(r.strip(), "&")[0])
                self.line("- **%s**：" % label_of(raw))
            self.line("")
        else:
            raise ValueError("未対応のコマンド: \\" + name)

    def env(self, name, opts, args, body):
        if name == "minipage":
            self.blocks(parse_blocks(body))
        elif name == "promptbox":
            self.gap()
            self.line("> **プロンプト例　%s**" % text_of(opts[0] or ""))
            self.line("> " + text_of(body, nl=" "))
            self.line("")
        elif name == "caution":
            parts = [p.strip() for p in text_of(body, nl=" ").split("□")]
            self.gap()
            self.line("> **注意：%s**" % text_of(args[0]))
            if parts[0]:
                self.line("> " + parts[0])
            for p in parts[1:]:
                if p:
                    self.line("> - [ ] " + p)
            self.line("")
        elif name == "checks":
            self.gap()
            for it in items_of(body):
                self.line("- [ ] " + text_of(it, nl=" "))
            self.line("")
        elif name == "nexttime":
            self.work = "次回までに"
            self.para("## 次回までに")
            for it in items_of(body):
                self.line("- [ ] " + text_of(it, nl=" "))
            self.line("")
        elif name in ("wstab", "wstabh"):
            self.table(args[0], body, name == "wstabh")
        else:
            raise ValueError("未対応の環境: " + name)


def convert(num):
    src = strip_comments((HERE / "sessions" / ("%s.tex" % num)).read_text(encoding="utf-8"))
    md = MD()
    md.blocks(parse_blocks(src))
    text = "\n".join(md.L)
    text = re.sub(r"\n{3,}", "\n\n", text)          # 空行は1つまで
    text = re.sub(r"：\n\n(?=\S)", "：\n\n", text)
    return text.rstrip() + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("nums", nargs="*", help="回番号（2桁）。省略すると全回")
    ap.add_argument("-o", "--outdir", type=Path, default=OUT)
    a = ap.parse_args()
    nums = a.nums or sorted(p.stem for p in (HERE / "sessions").glob("*.tex"))
    a.outdir.mkdir(parents=True, exist_ok=True)
    for num in nums:
        path = a.outdir / ("worksheet-%s.md" % num)
        path.write_text(convert(num), encoding="utf-8")
        print(path)


if __name__ == "__main__":
    main()
