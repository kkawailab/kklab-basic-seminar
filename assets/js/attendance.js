// 出席確認クイズの動作。
// 問題と正解は各ページ（_attendance/NN.md）から、レイアウトが
// <section class="quiz" data-page-id="..." data-answers='{"q1":"a",...}'> に書き出す。
// 全問正解すると、サーバー時刻と確認コード付きの出席カードを表示する。
(function () {
  'use strict';

  var quiz = document.querySelector('.quiz[data-answers]');
  if (!quiz) { return; }

  var ANSWERS = {};
  try { ANSWERS = JSON.parse(quiz.getAttribute('data-answers')); } catch (e) { ANSWERS = {}; }
  var PAGE_ID = quiz.getAttribute('data-page-id');
  var QUESTION_COUNT = Object.keys(ANSWERS).length;
  var STORE = 'bs2026-student-';

  // 選択肢の表示順をページ読み込みごとに入れ替える
  (function shuffleChoices() {
    var blocks = document.querySelectorAll('.question');
    for (var b = 0; b < blocks.length; b++) {
      var labels = Array.prototype.slice.call(blocks[b].querySelectorAll('.choice'));
      var anchor = blocks[b].querySelector('.explain');
      for (var i = labels.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var t = labels[i]; labels[i] = labels[j]; labels[j] = t;
      }
      for (var k = 0; k < labels.length; k++) { blocks[b].insertBefore(labels[k], anchor); }
    }
  })();

  // 学籍番号・氏名は端末内に記憶（入力の手間を減らすため）
  (function restoreStudent() {
    try {
      var id = localStorage.getItem(STORE + 'id');
      var name = localStorage.getItem(STORE + 'name');
      if (id) { document.getElementById('student-id').value = id; }
      if (name) { document.getElementById('student-name').value = name; }
    } catch (e) { /* 記憶できなくても動作に支障なし */ }
  })();

  function getChecked(name) {
    var checked = document.querySelector('input[name="' + name + '"]:checked');
    return checked ? checked.value : null;
  }

  function showError(message) {
    var box = document.getElementById('error-box');
    box.textContent = message;
    box.hidden = false;
    box.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function hideError() {
    document.getElementById('error-box').hidden = true;
  }

  function clearQuestion(i) {
    var radios = document.querySelectorAll('input[name="q' + i + '"]');
    for (var r = 0; r < radios.length; r++) { radios[r].checked = false; }
  }

  function setSubmitting(busy) {
    var btn = document.getElementById('submit-btn');
    btn.disabled = busy;
    btn.textContent = busy ? '提出処理中…' : '提出する';
  }

  async function submitQuiz() {
    hideError();

    var id = document.getElementById('student-id').value.trim();
    var name = document.getElementById('student-name').value.trim();

    if (!id || !name) {
      showError('学籍番号と氏名を入力してください。');
      return;
    }

    var unanswered = [], wrong = [];
    for (var i = 1; i <= QUESTION_COUNT; i++) {
      var v = getChecked('q' + i);
      var block = document.getElementById('q' + i);
      block.classList.remove('wrong');
      if (!v) { unanswered.push(i); }
      else if (v !== ANSWERS['q' + i]) { wrong.push(i); }
    }

    if (unanswered.length) {
      showError('Q' + unanswered.join('・Q') + ' に回答してください。');
      return;
    }

    if (wrong.length) {
      for (var w = 0; w < wrong.length; w++) {
        document.getElementById('q' + wrong[w]).classList.add('wrong');
        clearQuestion(wrong[w]);
      }
      showError(wrong.length + '問が不正解です（Q' + wrong.join('・Q') + '）。解説を読んで、もう一度回答してください。');
      return;
    }

    if (!window.crypto || !crypto.subtle) {
      showError('このページは https:// で開く必要があります。URLを確認してください。');
      return;
    }

    try {
      localStorage.setItem(STORE + 'id', id);
      localStorage.setItem(STORE + 'name', name);
    } catch (e) { /* 無視 */ }

    setSubmitting(true);
    var timeText = await fetchServerTime();
    if (!timeText) {
      setSubmitting(false);
      showError('時刻サーバーに接続できませんでした。通信環境を確認して、もう一度「提出する」を押してください。');
      return;
    }

    var code = await computeCode(PAGE_ID, id, name, timeText);

    document.getElementById('result-id').textContent = id;
    document.getElementById('result-name').textContent = name;
    document.getElementById('result-time').textContent = timeText;
    document.getElementById('result-code').textContent = code;
    document.getElementById('form-page').hidden = true;
    document.getElementById('result').hidden = false;
    window.scrollTo(0, 0);
  }

  document.getElementById('submit-btn').addEventListener('click', submitQuiz);
})();
