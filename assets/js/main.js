/* 日程データをもとに「次回」「終了した回」の表示を切り替える。
   日程は _data/schedule.yml に書く。
   次回 = 今日以降で最も早い日付の回（日付が空の回は、日付のある回がすべて終わった後の候補）。
   確認用に ?today=2026-12-17 のように日付を指定すると、その日として表示できる。 */
(function () {
  'use strict';

  function parseDate(s) {
    if (!s) { return null; }
    var p = s.split('-');
    if (p.length !== 3) { return null; }
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }

  var jsonEl = document.getElementById('sessions-json');
  var courseYear = jsonEl ? +jsonEl.getAttribute('data-year') : NaN;

  function formatDate(s) {
    var d = parseDate(s);
    if (!d) { return ''; }
    var wd = ['日', '月', '火', '水', '木', '金', '土'];
    var y = (courseYear && d.getFullYear() !== courseYear) ? d.getFullYear() + '年' : '';
    return y + (d.getMonth() + 1) + '月' + d.getDate() + '日（' + wd[d.getDay()] + '）';
  }

  var today = new Date();
  var override = /[?&]today=(\d{4}-\d{2}-\d{2})/.exec(window.location.search);
  if (override) { today = parseDate(override[1]); }
  today.setHours(0, 0, 0, 0);

  // 全回のデータ。トップページでは JSON から、その他のページでは data-session 要素から集める
  var sessions = [];
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

  // 今日以降の回を日付順に並べ、最初のものを「次回」にする
  var upcoming = sessions.filter(function (s) {
    var d = parseDate(s.date);
    return d && d >= today;
  }).sort(function (a, b) { return parseDate(a.date) - parseDate(b.date); });
  var next = upcoming[0] || null;
  if (!next) {
    var undated = sessions.filter(function (s) { return !parseDate(s.date); });
    next = undated[0] || null;
  }

  var status = {};
  sessions.forEach(function (s) {
    var d = parseDate(s.date);
    if (d && d < today) { status[s.number] = 'past'; }
    else if (next && s.number === next.number) { status[s.number] = 'next'; }
    else { status[s.number] = 'future'; }
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
