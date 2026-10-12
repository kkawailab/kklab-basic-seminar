#!/usr/bin/env python3
"""LuaLaTeX 版ワークシートの原稿（sessions/NN.tex）から Google フォーム版を作る。

Google フォームは外部から直接作れないため、Google Apps Script（FormApp）でフォームを
一括生成するスクリプトを書き出す。先生が script.google.com でこのスクリプトを実行すると、
Drive のフォルダに15回分のフォームと、回答を集める1つのスプレッドシートができる。

tex2forms.py の読み替え（セクション・質問・種類）をそのまま使う。
  セクション（■ ワーク）  → セクション見出し（タイトル＋説明。説明にプロンプト例と指示文）
  テキスト（短い回答）    → 記述式
  テキスト（長い回答）    → 段落
  選択肢（複数選択可）    → チェックボックス
  ファイル               → 段落（FormApp では作れないため、編集画面で「ファイルのアップロード」に置き換える）

  python3 tex2gforms.py            gforms/ に次を書き出す
    forms.json              フォームの定義
    build-forms.gs          定義を埋め込んだ Apps Script（これだけで動く）
    build-forms-github.gs   定義を GitHub の forms.json から読む軽い Apps Script（push 後に動く）
"""
import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from tex2forms import convert  # noqa: E402

OUT = HERE / "gforms"
FOLDER = "基礎演習Ⅱ 2026年度後期 ワークシート（Google フォーム）"
SHEET = "基礎演習Ⅱ 2026年度後期 ワークシート回答"
FILE_NOTE = "（編集画面でこの質問を「ファイルのアップロード」に置き換えてください）"
FORMS_URL = "https://raw.githubusercontent.com/kkawailab/kklab-basic-seminar/main/materials/worksheets/gforms/forms.json"


def build(num):
    form = convert(num)
    sections = []
    for s in form.sections:
        items = []
        for q in s.questions:
            hint = q.hint
            if q.kind == "file":
                hint = (hint + "\n" if hint else "") + FILE_NOTE
            items.append({"title": q.title, "kind": q.kind, "help": hint,
                          "choices": q.options, "required": q.required})
        sections.append({"title": s.title if s.title != "回答者" else "",
                         "desc": "\n".join(s.desc), "items": items})
    return {"num": num, "title": form.title, "description": "\n".join(form.desc), "sections": sections}


GS_HEAD = r"""// ============================================================
// 基礎演習Ⅱ 各回ワークシート　Google フォーム版を作る Apps Script
//   tex2gforms.py が sessions/NN.tex から生成したもの。手で直さず、原稿を直して作り直す。
//
// 使い方
//   1. https://script.google.com で「新しいプロジェクト」を作り、このファイルの内容を貼り付けて保存する
//   2. 関数 buildAllForms を選んで「実行」。初回は Drive とフォームへのアクセスを許可する
//   3. 実行ログに各フォームの編集用 URL と回答用 URL が出る。Drive のフォルダ
//      「%(folder)s」にフォームと回答スプレッドシートができる
//   同じ名前のフォームがフォルダにあれば飛ばす（作り直すときはフォームを削除してから実行）。
//   1回分だけ作るときは buildOne('03') のように実行する。
// ============================================================
var FOLDER_NAME = %(folder_js)s;
var SHEET_NAME = %(sheet_js)s;
var FORMS_URL = %(url_js)s;   // 定義が埋め込まれていないときはここから読む

function buildAllForms() {
  forms_().forEach(function (def) { buildForm_(def); });
}

function buildOne(num) {
  var def = forms_().filter(function (d) { return d.num === String(num); })[0];
  if (!def) throw new Error('第' + num + '回の定義がありません');
  buildForm_(def);
}

function forms_() {
  if (typeof FORMS !== 'undefined' && FORMS) return FORMS;
  return JSON.parse(UrlFetchApp.fetch(FORMS_URL).getContentText());
}

function folder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function sheet_(folder) {
  var it = folder.getFilesByName(SHEET_NAME);
  if (it.hasNext()) return it.next().getId();
  var ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss.getId();
}

function buildForm_(def) {
  var folder = folder_();
  if (folder.getFilesByName(def.title).hasNext()) {
    Logger.log('%%s: すでにあるので飛ばしました', def.title);
    return;
  }
  var form = FormApp.create(def.title);
  form.setDescription(def.description);
  form.setAllowResponseEdits(true);        // 送信後も自分の回答を直せる
  form.setShowLinkToRespondAgain(false);
  def.sections.forEach(function (sec) {
    if (sec.title) {
      var h = form.addSectionHeaderItem().setTitle(sec.title);
      if (sec.desc) h.setHelpText(sec.desc);
    }
    sec.items.forEach(function (q) {
      var item;
      if (q.kind === 'multi') {
        item = form.addCheckboxItem().setChoiceValues(q.choices);
      } else if (q.kind === 'short') {
        item = form.addTextItem();
      } else {
        item = form.addParagraphTextItem();
      }
      item.setTitle(q.title);
      if (q.help) item.setHelpText(q.help);
      if (q.required) item.setRequired(true);
    });
  });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet_(folder));
  DriveApp.getFileById(form.getId()).moveTo(folder);
  Logger.log('%%s\n  編集: %%s\n  回答: %%s', def.title, form.getEditUrl(), form.getPublishedUrl());
}
"""
GS_EMBED = "\n// 各回のフォームの定義（forms.json と同じ）\nvar FORMS = "
GS_GITHUB = ("\n// この版は定義を埋め込まず、実行時に GitHub の forms.json を読む。\n"
             "// 初回実行時に「外部サービスへの接続」の許可も求められる。\n")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--outdir", type=Path, default=OUT)
    a = ap.parse_args()
    a.outdir.mkdir(parents=True, exist_ok=True)
    nums = sorted(p.stem for p in (HERE / "sessions").glob("*.tex"))
    forms = [build(n) for n in nums]
    (a.outdir / "forms.json").write_text(json.dumps(forms, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    head = GS_HEAD % {"folder": FOLDER, "folder_js": json.dumps(FOLDER, ensure_ascii=False),
                      "sheet_js": json.dumps(SHEET, ensure_ascii=False),
                      "url_js": json.dumps(FORMS_URL)}
    gs = head + GS_EMBED + json.dumps(forms, ensure_ascii=False, indent=1) + ";\n"
    (a.outdir / "build-forms.gs").write_text(gs, encoding="utf-8")
    (a.outdir / "build-forms-github.gs").write_text(head + GS_GITHUB, encoding="utf-8")
    for f in forms:
        print("%s: %s  %dセクション %d問" % (f["num"], f["title"], len(f["sections"]),
                                           sum(len(s["items"]) for s in f["sections"])))


if __name__ == "__main__":
    main()
