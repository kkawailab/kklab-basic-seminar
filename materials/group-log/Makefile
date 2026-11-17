# 基礎演習Ⅱ　産学連携プロジェクト　グループ活動記録シートのビルド
#   make        group-log.pdf（PDF 版）と group-log.docx（Word 版）を作る
#   make pdf    PDF 版だけ
#   make word   Word 版だけ（python-docx が必要。なければ uv で一時的に入れて使う）
#   make clean  中間ファイルを消す
# デザインは各回ワークシートと共通（../worksheets/preamble.tex を読み込む）

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build

.PHONY: all pdf word clean

all: pdf word

pdf: group-log.pdf

word: group-log.docx

group-log.pdf: group-log.tex ../worksheets/preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) group-log.tex >/dev/null
	cp $(BUILD)/group-log.pdf $@

PYDOCX := $(shell python3 -c 'import docx' 2>/dev/null && echo python3 || echo uv run -q --with python-docx python3)

group-log.docx: build-docx.py ../worksheets/tex2docx.py
	$(PYDOCX) build-docx.py

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
