# 基礎演習Ⅱ　産学連携プロジェクト　企業への質問リストのビルド
#   make        company-qa.pdf を作る
#   make clean  中間ファイルを消す
# デザインは各回ワークシートと共通（../worksheets/preamble.tex を読み込む）

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build

.PHONY: all clean

all: company-qa.pdf

%.pdf: %.tex ../worksheets/preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	cp $(BUILD)/$@ $@

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
