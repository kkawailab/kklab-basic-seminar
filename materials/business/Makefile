# 基礎演習Ⅱ　事業創出テンプレートのビルド
#   make        PDF を作る
#   make clean  中間ファイルを消す

LUALATEX := lualatex -interaction=nonstopmode -halt-on-error -file-line-error
BUILD    := build
DOC      := business_creation_template_ja

.PHONY: all clean

all: $(DOC).pdf

# 相互参照を揃えるため2回組む
%.pdf: %.tex | $(BUILD)
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	$(LUALATEX) -output-directory=$(BUILD) $< >/dev/null
	cp $(BUILD)/$@ $@

$(BUILD):
	mkdir -p $@

clean:
	rm -rf $(BUILD)
