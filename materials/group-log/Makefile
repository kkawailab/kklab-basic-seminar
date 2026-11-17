# 基礎演習Ⅱ　産学連携プロジェクト　グループ活動記録シートのビルド
#   make        group-log.pdf を作る
#   make clean  中間ファイルを消す
# デザインは各回ワークシートと共通（../worksheets/preamble.tex を読み込む）

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build

.PHONY: all clean

all: group-log.pdf

group-log.pdf: group-log.tex ../worksheets/preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) group-log.tex >/dev/null
	cp $(BUILD)/group-log.pdf $@

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
