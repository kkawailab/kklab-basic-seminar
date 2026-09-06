/* 日程データをもとに「次回」「終了した回」の表示を切り替える。
   日程は _data/schedule.yml に書く。日程が空の回は「これから」の扱いになる。 */
(function () {
  'use strict';

  function parseDate(s) {
    if (!s) { return null; }
    var p = s.split('-');
    if (p.length !== 3) { return null; }
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  function formatDate(s) {
    var d = parseDate(s);
    if (!d) { return ''; }
    var wd = ['日', '月', '火', '水', '木', '金', '土'];
    return (d.getMonth() + 1) + '月' + d.getDate() + '日（' + wd[d.getDay()] + '）';
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // 全回のデータ。トップページでは JSON から、その他のページでは data-session 要素から集める
  var sessions = [];
  var jsonEl = document.getElementById('sessions-json');
  if (jsonEl) {
    try { sessions = JSON.parse(jsonEl.textContent); } catch (e) { sessions = []; }
  }
  var cells = Array.prototype.slice.call(document.querySelectorAll('[data-session]'));
  if (!sessions.length) {
    cells.forEach(function (el) {
      sessions.push({ number: +el.getAttribute('data-session'), date: el.getAttribute('data-date') || null });
    });
  }
  sessions.sort(function (a, b) { return a.number - b.number; });

  // 日付が今日より前なら「終了」。それ以外で最初に来る回が「次回」
  var status = {};
  var next = null;
  sessions.forEach(function (s) {
    var d = parseDate(s.date);
    if (d && d < today) { status[s.number] = 'past'; return; }
    if (!next) { next = s; status[s.number] = 'next'; return; }
    status[s.number] = 'future';
  });

  cells.forEach(function (el) {
    var st = status[+el.getAttribute('data-session')];
    if (st === 'past') { el.classList.add('is-past'); }
    if (st === 'next') { el.classList.add('is-next'); }
  });

  // トップページの「次回の授業」を書き換える
  var box = document.getElementById('next-session');
  if (box && next && next.title) {
    box.className = box.className.replace(/\bphase-\d\b/g, '').trim() + ' phase-' + next.phase;
    var set = function (name, text) {
      var el = box.querySelector('[data-field="' + name + '"]');
      if (el) { el.textContent = text; }
      return el;
    };
    set('num', '第' + next.number + '回');
    set('phase', next.phase_name || '');
    var title = set('title', next.title);
    if (title) { title.setAttribute('href', next.url); }
    var tag = set('tag', next.tag || '');
    if (tag) { tag.hidden = !next.tag; }
    set('summary', next.summary || '');
    set('date', formatDate(next.date) || '日程未定');
    if (next.place) { set('place', next.place); }
  }
})();
