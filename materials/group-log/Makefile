# 基礎演習Ⅱ　産学連携プロジェクト　グループ活動記録シートのビルド
#   make        産学連携用（group-log）と汎用・シンプル版（group-log-simple）の PDF と Word をすべて作る
#   make pdf    PDF 版だけ（両方）
#   make word   Word 版だけ（両方。python-docx が必要。なければ uv で一時的に入れて使う）
#   make simple 汎用・シンプル版（group-log-simple.pdf / .docx）だけ
#   make clean  中間ファイルを消す
# デザインは各回ワークシートと共通（../worksheets/preamble.tex を読み込む）

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build

.PHONY: all pdf word simple clean

all: pdf word

pdf: group-log.pdf group-log-simple.pdf

word: group-log.docx group-log-simple.docx

simple: group-log-simple.pdf group-log-simple.docx

%.pdf: %.tex ../worksheets/preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	cp $(BUILD)/$@ $@

PYDOCX := $(shell python3 -c 'import docx' 2>/dev/null && echo python3 || echo uv run -q --with python-docx python3)

group-log.docx: build-docx.py ../worksheets/tex2docx.py
	$(PYDOCX) build-docx.py full

group-log-simple.docx: build-docx.py ../worksheets/tex2docx.py
	$(PYDOCX) build-docx.py simple

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
