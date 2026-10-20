# 基礎演習Ⅱ　授業用の架空論文のビルド
#   make            すべての論文（paper-*.tex）を PDF にする
#   make paper-x.pdf 架空論文Ｘだけ
#   make clean      中間ファイルを消す

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build
PAPERS   := $(patsubst %.tex,%.pdf,$(wildcard paper-*.tex))

.PHONY: all clean

all: $(PAPERS)

# 表番号や相互参照を揃えるため2回組む
%.pdf: %.tex preamble.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	cp $(BUILD)/$@ $@

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
