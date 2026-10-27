# 基礎演習Ⅱ　総まとめ冊子のビルド
#   make                  すべての冊子（summary-*.tex）を PDF にする
#   make summary-01-05.pdf  第1回〜第5回 総まとめだけ
#   make clean            中間ファイルを消す

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build
DOCS     := $(patsubst %.tex,%.pdf,$(wildcard summary-*.tex))

.PHONY: all clean

all: $(DOCS)

# 目次とプロンプト一覧の頁番号を揃えるため2回組む
%.pdf: %.tex preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	cp $(BUILD)/$@ $@

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
