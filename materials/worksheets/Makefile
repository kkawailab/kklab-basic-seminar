# 基礎演習Ⅱ 各回ワークシートのビルド
#   make            全回版（worksheets.pdf）と各回版（pdf/worksheet-NN.pdf）をすべて作る
#   make all-in-one 全回版だけ
#   make single     各回版だけ（make -j で並行ビルドできる）
#   make 03         第3回だけ（pdf/worksheet-03.pdf）
#   make word       Word 版（word/worksheets.docx と word/worksheet-NN.docx）を作る
#   make clean      中間ファイルを消す

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build
NUMS     := 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15
SINGLES  := $(NUMS:%=pdf/worksheet-%.pdf)
DOCXS    := word/worksheets.docx $(NUMS:%=word/worksheet-%.docx)
COMMON   := preamble.tex

.PHONY: all all-in-one single word clean $(NUMS)

all: all-in-one single word

all-in-one: worksheets.pdf

single: $(SINGLES)

$(NUMS): %: pdf/worksheet-%.pdf

# 目次のページ番号を揃えるため2回組む
worksheets.pdf: worksheets.tex cover.tex $(COMMON) $(NUMS:%=sessions/%.tex) | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) worksheets.tex >/dev/null
	$(LUALATEX) -output-directory=$(BUILD) worksheets.tex >/dev/null
	cp $(BUILD)/worksheets.pdf $@

pdf/worksheet-%.pdf: sessions/%.tex session.tex $(COMMON) | $(BUILD) pdf
	$(LUALATEX) -output-directory=$(BUILD) -jobname=worksheet-$* '\def\sessionno{$*}\input{session}' >/dev/null
	cp $(BUILD)/worksheet-$*.pdf $@

# Word 版は sessions/NN.tex を tex2docx.py で変換する（内容は PDF 版と同じ）
word: $(DOCXS)

word/worksheets.docx: tex2docx.py cover.tex $(NUMS:%=sessions/%.tex) | word
	python3 tex2docx.py -o $@

word/worksheet-%.docx: sessions/%.tex tex2docx.py | word
	python3 tex2docx.py -o $@ $*

$(BUILD) pdf word:
	mkdir -p $@

clean:
	rm -rf $(BUILD)
