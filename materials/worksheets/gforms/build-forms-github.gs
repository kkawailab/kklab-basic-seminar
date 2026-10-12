// ============================================================
// 基礎演習Ⅱ 各回ワークシート　Google フォーム版を作る Apps Script
//   tex2gforms.py が sessions/NN.tex から生成したもの。手で直さず、原稿を直して作り直す。
//
// 使い方
//   1. https://script.google.com で「新しいプロジェクト」を作り、このファイルの内容を貼り付けて保存する
//   2. 関数 buildAllForms を選んで「実行」。初回は Drive とフォームへのアクセスを許可する
//   3. 実行ログに各フォームの編集用 URL と回答用 URL が出る。Drive のフォルダ
//      「基礎演習Ⅱ 2026年度後期 ワークシート（Google フォーム）」にフォームと回答スプレッドシートができる
//   同じ名前のフォームがフォルダにあれば飛ばす（作り直すときはフォームを削除してから実行）。
//   1回分だけ作るときは buildOne('03') のように実行する。
// ============================================================
var FOLDER_NAME = "基礎演習Ⅱ 2026年度後期 ワークシート（Google フォーム）";
var SHEET_NAME = "基礎演習Ⅱ 2026年度後期 ワークシート回答";
var FORMS_URL = "https://raw.githubusercontent.com/kkawailab/kklab-basic-seminar/main/materials/worksheets/gforms/forms.json";   // 定義が埋め込まれていないときはここから読む

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
    Logger.log('%s: すでにあるので飛ばしました', def.title);
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
  Logger.log('%s\n  編集: %s\n  回答: %s', def.title, form.getEditUrl(), form.getPublishedUrl());
}

// この版は定義を埋め込まず、実行時に GitHub の forms.json を読む。
// 初回実行時に「外部サービスへの接続」の許可も求められる。
