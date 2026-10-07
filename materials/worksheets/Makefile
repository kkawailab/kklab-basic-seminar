# 基礎演習Ⅱ 各回ワークシートのビルド
#   make            全回版（worksheets.pdf）と各回版（pdf/worksheet-NN.pdf）をすべて作る
#   make all-in-one 全回版だけ
#   make single     各回版だけ（make -j で並行ビルドできる）
#   make 03         第3回だけ（pdf/worksheet-03.pdf）
#   make clean      中間ファイルを消す

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build
NUMS     := 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15
SINGLES  := $(NUMS:%=pdf/worksheet-%.pdf)
COMMON   := preamble.tex

.PHONY: all all-in-one single clean $(NUMS)

all: all-in-one single

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

$(BUILD) pdf:
	mkdir -p $@

clean:
	rm -rf $(BUILD)
