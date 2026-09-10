'use strict';
/* Расписание: разбор выгрузок Excel прямо в браузере.
   s/ — группы, t/ — преподаватели, list.txt — сопоставление имён и файлов. */

var DIRS = { t: 't/', s: 's/' };

/* ── Разбор ─────────────────────────────────────────────── */

var LABELS = {
  'день недели': 'day', '№': 'num', 'пара': 'num',
  'начало': 'start', 'окончание': 'end', 'предмет': 'subject',
  'преподаватель': 'teacher', 'группа': 'group',
  'кабинет': 'room', 'каб.': 'room', 'каб': 'room'
};
var TIME_RE = /^\d{1,2}[:.]\d{2}$/;
var DAY_RE = /^(Понедельник|Вторник|Среда|Четверг|Пятница|Суббота|Воскресенье)\s*[-–—]?\s*(.*)$/i;
var WEEK_RE = /недел\S*\s*№?\s*(\d+)\s*(числител\S*|знаменател\S*)?/i;
var GROUP_RE = /^[A-Za-zА-Яа-яЁё]{1,6}[-‑]?\d{2,4}[А-Яа-яA-Za-z/\d-]*$/;

function norm(s) {
  return (s || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

/* Раскрывает rowspan/colspan: grid[r][c] = {text, origin} */
function buildGrid(table) {
  var grid = [], rows = table.rows;
  for (var ri = 0; ri < rows.length; ri++) {
    if (!grid[ri]) grid[ri] = [];
    var cells = rows[ri].cells, ci = 0;
    for (var i = 0; i < cells.length; i++) {
      while (grid[ri][ci]) ci++;
      var td = cells[i];
      var rs = parseInt(td.getAttribute('rowspan'), 10) || 1;
      var cs = parseInt(td.getAttribute('colspan'), 10) || 1;
      var text = norm(td.textContent);
      for (var r = ri; r < ri + rs; r++) {
        if (!grid[r]) grid[r] = [];
        for (var c = ci; c < ci + cs; c++) {
          grid[r][c] = { text: text, origin: r === ri && c === ci };
        }
      }
      ci += cs;
    }
  }
  return grid;
}

function at(grid, r, c) { var v = grid[r] && grid[r][c]; return v ? v.text : ''; }
function own(grid, r, c) { var v = grid[r] && grid[r][c]; return v && v.origin ? v.text : ''; }

function parseSchedule(doc, fallbackName) {
  var table = doc.querySelector('table');
  if (!table) throw new Error('таблица не найдена');

  var grid = buildGrid(table);
  var nrows = grid.length, ncols = 0, r, c;
  for (r = 0; r < nrows; r++) ncols = Math.max(ncols, (grid[r] || []).length);

  var head = -1, dayCol = 0;
  for (r = 0; r < Math.min(nrows, 12) && head < 0; r++) {
    for (c = 0; c < ncols; c++) {
      if (at(grid, r, c).toLowerCase() === 'день недели') { head = r; dayCol = c; break; }
    }
  }
  if (head < 0) throw new Error('не найдена строка заголовков');

  var cols = {};
  [head, head + 1].forEach(function (row) {
    for (var c = 0; c < ncols; c++) {
      var key = LABELS[own(grid, row, c).toLowerCase().replace(/:$/, '')];
      if (key) cols[key] = c;
    }
  });
  ['num', 'start', 'end', 'subject'].forEach(function (k) {
    if (cols[k] === undefined) throw new Error('нет колонки: ' + k);
  });

  var kind = cols.group !== undefined ? 'teacher' : 'group';
  var whoCol = kind === 'teacher' ? cols.group : cols.teacher;

  /* заголовок над таблицей */
  var title = '';
  for (r = 0; r < head; r++) title += ' ' + at(grid, r, 0);
  title = norm(title);

  var owner = '', week = '', parity = '';
  var m = WEEK_RE.exec(title);
  if (m) { week = m[1]; parity = (m[2] || '').toLowerCase(); }
  m = /Преподавател[ья]\s+(.+?)\s*(?:Расписание|$)/.exec(title);
  if (m) owner = norm(m[1]);

  if (kind === 'group' && !owner) {
    for (r = head; r < Math.min(head + 3, nrows) && !owner; r++) {
      for (c = 0; c < ncols; c++) {
        var t = own(grid, r, c);
        if (t && GROUP_RE.test(t) && !LABELS[t.toLowerCase()]) { owner = t; break; }
      }
    }
  }

  /* пары */
  var days = [], index = {}, total = 0;
  for (r = head + 1; r < nrows; r++) {
    if (!TIME_RE.test(at(grid, r, cols.start))) continue;
    var num = at(grid, r, cols.num);
    if (!/^\d+$/.test(num)) continue;

    var dayRaw = at(grid, r, dayCol);
    var dm = DAY_RE.exec(dayRaw);
    if (!dm) continue;

    if (!(dayRaw in index)) {
      index[dayRaw] = days.length;
      days.push({
        name: dm[1].charAt(0).toUpperCase() + dm[1].slice(1).toLowerCase(),
        date: norm(dm[2]), lessons: [], byNum: {}
      });
    }
    var day = days[index[dayRaw]];

    if (!day.byNum[num]) {
      day.byNum[num] = {
        num: num,
        start: at(grid, r, cols.start).replace('.', ':'),
        end: at(grid, r, cols.end).replace('.', ':'),
        subject: '', variants: []
      };
      day.lessons.push(day.byNum[num]);
    }
    var lesson = day.byNum[num];

    var subject = own(grid, r, cols.subject);
    if (subject && !lesson.subject) lesson.subject = subject;

    var who = whoCol !== undefined ? own(grid, r, whoCol) : '';
    var room = cols.room !== undefined ? own(grid, r, cols.room) : '';
    if (who || room) {
      var dup = lesson.variants.some(function (v) { return v.who === who && v.room === room; });
      if (!dup) lesson.variants.push({ who: who, room: room });
    }
  }

  days.forEach(function (day) {
    delete day.byNum;
    var ls = day.lessons;
    ls.forEach(function (l) { l.free = !l.subject && !l.variants.length; });
    /* хвост пустых пар — это просто конец дня, а не окно;
       пустые пары в начале дня показываем: люди привыкли их видеть */
    var b = ls.length;
    while (b > 0 && ls[b - 1].free) b--;
    day.lessons = ls.slice(0, b);
    day.lessons.forEach(function (l) { if (!l.free) total++; });
  });

  /* подвал: нагрузка, дата выгрузки, объявление */
  var load = '', stamp = '', note = '';
  for (r = 0; r < nrows; r++) {
    for (c = 0; c < ncols; c++) {
      var lab = own(grid, r, c).toLowerCase();
      if (lab.indexOf('недельная нагрузка') === 0) {
        load = firstAfter(grid, r, c, ncols);
      } else if (lab === 'дата и время' || lab === 'дата') {
        stamp = firstAfter(grid, r, c, ncols);
        if (lab === 'дата') {
          for (var r2 = r + 1; r2 < Math.min(r + 4, nrows); r2++) {
            if (own(grid, r2, c).toLowerCase() === 'время') {
              stamp += ' ' + firstAfter(grid, r2, c, ncols);
              break;
            }
          }
        }
      } else if (!note && /^уважаем/i.test(lab)) {
        note = own(grid, r, c);
      }
    }
  }

  return {
    kind: kind, owner: owner || fallbackName, week: week, parity: parity,
    load: load, stamp: norm(stamp), note: note, days: days, total: total
  };
}

function firstAfter(grid, r, c, ncols) {
  for (var x = c + 1; x < ncols; x++) if (own(grid, r, x)) return own(grid, r, x);
  return '';
}

/* ── Отрисовка ──────────────────────────────────────────── */

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function renderStamps(data) {
  var out = [];
  if (data.week) out.push('Неделя № ' + data.week + (data.parity ? ' · ' + data.parity : ''));
  var dates = data.days.map(function (d) { return d.date; }).filter(Boolean);
  if (dates.length) out.push(dates[0] + ' — ' + dates[dates.length - 1]);
  if (data.kind === 'teacher' && data.load) out.push('Нагрузка: ' + data.load);
  return out.map(function (s) { return '<span class="stamp">' + esc(s) + '</span>'; }).join('');
}

/* Часовой пояс расписания. 180 — Москва, 300 — Екатеринбург и т.д.
   Нужен, чтобы пара считалась идущей по времени колледжа, а не по часам телефона. */
var TZ_OFFSET_MIN = 180;

function moment(date, time) {
  var d = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(date || '');
  var t = /^(\d{1,2})[:.](\d{2})$/.exec(time || '');
  if (!d || !t) return null;
  return Date.UTC(+d[3], +d[2] - 1, +d[1], +t[1], +t[2]) - TZ_OFFSET_MIN * 60000;
}

function span(day, from, to) {
  var a = moment(day.date, from.start), b = moment(day.date, to.end);
  return a && b ? ' data-from="' + a + '" data-to="' + b + '"' : '';
}

function slot(label, start, end) {
  return '<p class="slot"><span class="slot__no">' + esc(label) + '</span>' +
    '<span class="slot__time">' + esc(start) + '<i>' + esc(end) + '</i></span></p>';
}

function renderLesson(day, lesson, kind) {
  var out = '<article class="lesson"' + span(day, lesson, lesson) + '>' +
    slot(lesson.num + ' пара', lesson.start, lesson.end) +
    '<div class="lesson__card"><h3 class="subject">' + esc(lesson.subject || 'Занятие') + '</h3>';

  var split = lesson.variants.length > 1;
  lesson.variants.forEach(function (v, i) {
    var parts = [];
    if (split) parts.push('<span class="tag">' + (i + 1) + ' п/г</span>');
    if (v.who) parts.push('<span class="who-name">' + esc(v.who) + '</span>');
    if (v.room) parts.push('<span class="room">' + esc(v.room) + '</span>');
    out += '<p class="meta">' + parts.join('') + '</p>';
  });
  if (!lesson.variants.length) {
    out += '<p class="meta"><span class="muted">' +
      (kind === 'group' ? 'Преподаватель не указан' : 'Группа не указана') + '</span></p>';
  }
  return out + '</div></article>';
}

function renderFree(day, lesson) {
  return '<article class="lesson lesson--free"' + span(day, lesson, lesson) + '>' +
    slot(lesson.num + ' пара', lesson.start, lesson.end) +
    '<div class="lesson__card"><h3 class="subject subject--free">Окно</h3></div></article>';
}

function renderSchedule(data) {
  if (!data.total) return '<p class="free free--week">На этой неделе занятий нет</p>';

  return data.days.map(function (day) {
    var out = '<section class="day"><h2 class="day__head">' +
      '<span class="day__name">' + esc(day.name) + '</span>' +
      '<span class="day__date">' + esc(day.date) + '</span></h2>';

    if (!day.lessons.length) return out + '<p class="free">Занятий нет</p></section>';

    day.lessons.forEach(function (lesson) {
      out += lesson.free ? renderFree(day, lesson) : renderLesson(day, lesson, data.kind);
    });
    return out + '</section>';
  }).join('');
}

function renderFoot(data) {
  var out = '';
  if (data.note) out += '<p class="note">' + esc(data.note) + '</p>';
  if (data.stamp) out += '<p class="stamp-line">Выгружено ' + esc(data.stamp) + '</p>';
  return out;
}

/* ── Приложение ─────────────────────────────────────────── */

if (typeof document !== 'undefined') (function () {

  var els = {
    eyebrow: document.getElementById('eyebrow'),
    title: document.getElementById('title'),
    stamps: document.getElementById('stamps'),
    schedule: document.getElementById('schedule'),
    foot: document.getElementById('foot')
  };
  var drops = { t: document.getElementById('drop-t'), s: document.getElementById('drop-s') };
  var entries = { t: [], s: [] };
  var cache = {};
  var site = { name: '', url: '' };

  /* Excel иногда отдаётся в другой кодировке — проверяем и перечитываем */
  function decode(buf, contentType) {
    var m = /charset=([\w-]+)/i.exec(contentType || '');
    if (m) return new TextDecoder(m[1]).decode(buf);
    var text = new TextDecoder('utf-8').decode(buf);
    if (text.indexOf('\uFFFD') >= 0) {
      try { return new TextDecoder('windows-1251').decode(buf); } catch (e) { /* оставляем как есть */ }
    }
    return text;
  }

  /* ── Тема ── */

  (function () {
    var btn = document.getElementById('theme');
    if (!btn) return;
    var root = document.documentElement;
    var system = window.matchMedia ? window.matchMedia('(prefers-color-scheme:dark)') : null;

    function stored() {
      try { return localStorage.getItem('theme'); } catch (e) { return null; }
    }
    function current() {
      return stored() || (system && system.matches ? 'dark' : 'light');
    }
    function paint() {
      var dark = current() === 'dark';
      btn.setAttribute('data-icon', dark ? 'sun' : 'moon');
      btn.setAttribute('aria-label', dark ? 'Светлая тема' : 'Тёмная тема');
      btn.title = btn.getAttribute('aria-label');
    }

    btn.addEventListener('click', function () {
      var next = current() === 'dark' ? 'light' : 'dark';
      var auto = system && (system.matches ? 'dark' : 'light');
      try {
        /* совпало с настройкой системы — снова следуем за ней */
        if (next === auto) localStorage.removeItem('theme');
        else localStorage.setItem('theme', next);
      } catch (e) { /* без сохранения, но в этой вкладке сработает */ }
      if (next === auto) root.removeAttribute('data-theme');
      else root.setAttribute('data-theme', next);
      paint();
    });

    if (system && system.addEventListener) system.addEventListener('change', paint);
    paint();
  })();

  /* Расхождение часов клиента и сервера. Берётся из заголовка Date каждого
     ответа — отдельный запрос ради времени не нужен. */
  var skew = 0;
  function serverNow() { return Date.now() + skew; }

  function get(url) {
    return fetch(url, { cache: 'no-cache' }).then(function (res) {
      var d = res.headers ? Date.parse(res.headers.get('date') || '') : NaN;
      if (!isNaN(d)) skew = d - Date.now();
      if (!res.ok) throw new Error(res.status === 404 ? 'нет файла' : 'ошибка ' + res.status);
      return res.arrayBuffer().then(function (buf) {
        return decode(buf, res.headers.get('content-type'));
      });
    });
  }

  /* Подсветка идущей сейчас пары или окна. Раз в минуту, без перерисовки. */
  function tick() {
    var now = serverNow();
    var cards = document.querySelectorAll('.lesson[data-from]');
    for (var i = 0; i < cards.length; i++) {
      var el = cards[i];
      var on = now >= +el.getAttribute('data-from') && now < +el.getAttribute('data-to');
      el.classList.toggle('lesson--now', on);
    }
  }
  setInterval(tick, 60000);

  /* ── Списки ── */

  function section(kind, title) {
    return kind === 's' && title.indexOf('-') > 0
      ? title.split('-')[0].toUpperCase()
      : title.charAt(0).toUpperCase();
  }

  function fillDrop(kind) {
    var list = drops[kind].querySelector('.drop__list');
    var prev = null, html = '';
    entries[kind].forEach(function (it) {
      var head = section(kind, it.title);
      if (head !== prev) {
        prev = head;
        html += '<li class="drop__section">' + esc(head) + '</li>';
      }
      html += '<li><a href="#' + kind + '/' + esc(it.file) + '">' + esc(it.title) + '</a></li>';
    });
    list.innerHTML = html || '<li class="drop__section">Список пуст</li>';
  }

  /* Название колледжа и адрес его сайта — в site.txt, чтобы правка не
     требовала лезть в код. Файла нет — просто шапка без ссылки. */
  function applySite() {
    var badge = document.querySelector('.badge');
    if (!badge) return;
    if (site.url) {
      var label = 'На главную сайта' + (site.name ? ' ' + site.name : ' колледжа');
      badge.setAttribute('href', site.url);
      badge.setAttribute('aria-label', label);
      badge.title = label;
      badge.removeAttribute('aria-hidden');
    } else {
      badge.removeAttribute('href');
      badge.removeAttribute('aria-label');
      badge.removeAttribute('title');
      badge.setAttribute('aria-hidden', 'true');
    }
  }

  function loadSite() {
    return get('site.txt').then(function (text) {
      text.split(/\r?\n/).forEach(function (line) {
        if (!line || line.charAt(0) === '#') return;
        var i = line.indexOf('=');
        if (i < 0) return;
        var key = line.slice(0, i).trim().toLowerCase();
        if (key in site) site[key] = line.slice(i + 1).trim();
      });
    }).catch(function () { /* без файла остаются пустые значения */ })
      .then(applySite);
  }

  function loadList() {
    return get('list.txt').then(function (text) {
      text.split(/\r?\n/).forEach(function (line) {
        if (!line || line.charAt(0) === '#') return;
        var p = line.split('|');
        if (p.length < 3) return;
        var kind = p[0].trim();
        if (!entries[kind]) return;
        entries[kind].push({ file: p[1].trim(), title: p[2].trim() });
      });
      ['t', 's'].forEach(function (kind) {
        entries[kind].sort(function (a, b) {
          return a.title.toLowerCase().replace(/ё/g, 'е')
            .localeCompare(b.title.toLowerCase().replace(/ё/g, 'е'), 'ru');
        });
        fillDrop(kind);
      });
    });
  }

  /* ── Поиск внутри списка ── */

  var finePointer = !!(window.matchMedia &&
    window.matchMedia('(hover:hover) and (pointer:fine)').matches);

  ['t', 's'].forEach(function (kind) {
    var input = drops[kind].querySelector('.drop__search');
    input.addEventListener('input', function () {
      var q = input.value.trim().toLowerCase();
      var list = drops[kind].querySelector('.drop__list');
      list.classList.toggle('is-filtered', q.length > 0);
      Array.prototype.forEach.call(list.children, function (li) {
        var a = li.firstChild;
        li.hidden = q && a.tagName === 'A' && a.textContent.toLowerCase().indexOf(q) < 0;
      });
    });
    drops[kind].addEventListener('toggle', function () {
      /* На телефоне и киоске фокус в поле сразу поднимал бы виртуальную
         клавиатуру поверх списка. Ставим его только там, где есть мышь. */
      if (drops[kind].open) {
        if (finePointer) input.focus();
      } else {
        input.value = '';
        input.dispatchEvent(new Event('input'));
      }
    });
  });

  document.addEventListener('click', function (ev) {
    ['t', 's'].forEach(function (kind) {
      if (drops[kind].open && !drops[kind].contains(ev.target)) drops[kind].open = false;
    });
  });

  /* ── Показ ── */

  function setValue(kind, title) {
    drops[kind].querySelector('.drop__value').textContent = title || 'Не выбрано';
    var list = drops[kind].querySelector('.drop__list');
    Array.prototype.forEach.call(list.querySelectorAll('a'), function (a) {
      if (a.textContent === title) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
  }

  function message(text, sub) {
    els.schedule.innerHTML = '<p class="free free--week">' + esc(text) +
      (sub ? '<span class="free__sub">' + esc(sub) + '</span>' : '') + '</p>';
    els.foot.innerHTML = '';
  }

  function show(data) {
    els.eyebrow.textContent = data.kind === 'teacher'
      ? 'Расписание преподавателя' : 'Расписание группы';
    els.title.textContent = data.owner;
    els.stamps.innerHTML = renderStamps(data);
    els.schedule.innerHTML = renderSchedule(data);
    els.foot.innerHTML = renderFoot(data);
    tick();
    document.title = data.owner + ' — расписание';
    setValue(data.kind === 'teacher' ? 't' : 's', data.owner);
    setValue(data.kind === 'teacher' ? 's' : 't', '');
  }

  function route() {
    var hash = decodeURIComponent(location.hash.replace(/^#/, ''));
    var parts = hash.split('/');
    var kind = parts[0], file = parts.slice(1).join('/');

    drops.t.open = drops.s.open = false;
    if (back) back.hidden = !(DIRS[kind] && file);

    if (!DIRS[kind] || !file) {
      els.eyebrow.textContent = site.name;
      els.title.textContent = 'Расписание занятий';
      els.stamps.innerHTML = '';
      document.title = 'Расписание занятий';
      setValue('t', ''); setValue('s', '');
      message('Выберите преподавателя или группу');
      return;
    }

    if (cache[hash]) { show(cache[hash]); return; }

    var known = entries[kind].filter(function (it) { return it.file === file; })[0];
    els.title.textContent = known ? known.title : file;
    els.stamps.innerHTML = '';
    message('Загрузка…');

    get(DIRS[kind] + file + '.htm').then(function (text) {
      var doc = new DOMParser().parseFromString(text, 'text/html');
      var data = parseSchedule(doc, known ? known.title : file);
      cache[hash] = data;
      show(data);
    }).catch(function (err) {
      els.eyebrow.textContent = kind === 't' ? 'Расписание преподавателя' : 'Расписание группы';
      message('Расписание не открылось', String(err.message || err));
    });
  }

  /* ── Кнопка «Назад» ──
     На киоске нет кнопок браузера, поэтому шаги считаем сами. Если человек
     пришёл сразу по ссылке на расписание и отматывать нечего — уводим на
     стартовый экран, а не за пределы сайта. */

  var back = document.getElementById('back');
  var steps = 0, goingBack = false;

  if (back) back.addEventListener('click', function () {
    goingBack = true;
    if (steps > 0) history.back();
    else location.hash = '';
  });

  window.addEventListener('hashchange', function () {
    if (goingBack) { goingBack = false; steps = Math.max(0, steps - 1); }
    else steps++;
    route();
  });

  Promise.all([loadSite(), loadList()]).then(route).catch(function (err) {
    message('Список расписаний не загрузился', 'list.txt — ' + String(err.message || err));
  });

})();

/* для тестов вне браузера */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseSchedule: parseSchedule, renderSchedule: renderSchedule, renderStamps: renderStamps };
}
