(function () {
  'use strict';

  const X = window.Stundenzettel;
  const VERSION = '2.2';
  const KEY = 'stundenzettel.v1';
  const TKEY = 'stundenzettel.template.v1';
  const BUNDLED_TEMPLATE = 'vorlage.xlsx';
  const DAY_NAMES = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'];
  const DAY_SHORT = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
  const MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
  const DAY_TYPES = ['Drehtag', 'Vor/Nachbereitung', 'Reisetag', 'Tag 6', 'Tag 7', 'Ausfalltag', 'Krank', 'Feiertag frei',
    'Feiertag', 'Urlaub', 'Bezahlter Ruhetag Sonntag', 'Bezahlter Ruhetag sonstiges', 'Ruhetag', 'Zeitkonto', 'AZV Tag'];
  const MAIL_DEFAULTS = {
    mailSubject: 'Stundenzettel KW {KW} ({Zeitraum}) – {Vorname}',
    mailBody: 'Hallo {Empfänger},\n\nanbei mein Stundenzettel für die KW {KW} ({Zeitraum}).\n\nViele Grüße\n{Vorname}'
  };
  const hoursFmt = new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const CHEVRON = '<span class="chev"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg></span>';
  const SHARE_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"/><polyline points="16 6 12 2 8 6"/><line x1="12" y1="2" x2="12" y2="15"/></svg>';

  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const pad = n => String(n).padStart(2, '0');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ---------- Datum & Zeit ----------
  const ymd = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const parseYmd = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d, 12); };
  const addDays = (s, n) => { const d = parseYmd(s); d.setDate(d.getDate() + n); return ymd(d); };
  const today = () => ymd(new Date());
  const mondayOf = s => { const d = parseYmd(s); d.setDate(d.getDate() - (d.getDay() + 6) % 7); return ymd(d); };
  const weekdayIndex = s => (parseYmd(s).getDay() + 6) % 7;
  const dm = s => { const d = parseYmd(s); return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.'; };
  const dmy = s => dm(s) + parseYmd(s).getFullYear();
  const weekRange = monday => dm(monday) + ' – ' + dmy(addDays(monday, 6));
  const longDay = s => { const d = parseYmd(s); return DAY_NAMES[weekdayIndex(s)] + ', ' + d.getDate() + '. ' + MONTHS[d.getMonth()]; };
  // Aktuelle Uhrzeit, auf die nächste Viertelstunde gerundet
  const nowTime = () => {
    const d = new Date();
    const t = (Math.round((d.getHours() * 60 + d.getMinutes()) / 15) * 15) % 1440;
    return pad(Math.floor(t / 60)) + ':' + pad(t % 60);
  };
  const stamp = iso => { const d = new Date(iso); return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '., ' + pad(d.getHours()) + ':' + pad(d.getMinutes()); };

  function isoWeek(s) {
    const d = parseYmd(s);
    const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    return { week: Math.ceil(((t - yearStart) / 86400000 + 1) / 7), year: t.getUTCFullYear() };
  }

  function addMinutes(t, min) {
    const [h, m] = t.split(':').map(Number);
    const x = (((h * 60 + m + min) % 1440) + 1440) % 1440;
    return pad(Math.floor(x / 60)) + ':' + pad(x % 60);
  }

  const fmtH = h => hoursFmt.format(Math.abs(h) < 0.005 ? 0 : h);
  const hm = h => { const t = Math.round(h * 60), a = Math.abs(t); return (t < 0 ? '-' : '') + Math.floor(a / 60) + ':' + pad(a % 60); };

  // ---------- Speicher ----------
  const blankState = () => ({
    settings: { name: '', position: '', toName: '', toEmail: '', sendFormat: 'xlsx' },
    days: {}, sent: {}, edited: {}, projects: [], signature: null, migratedVorname: true
  });

  function normalizeState(s) {
    const b = blankState();
    if (!s || typeof s !== 'object') return b;
    const st = Object.assign(b, { migratedVorname: false }, s, { settings: Object.assign(b.settings, s.settings) });
    // Einmalig: eigene Mail-Vorlagen auf Vorname statt vollem Namen umstellen
    if (!st.migratedVorname) {
      ['mailSubject', 'mailBody'].forEach(k => { if (st.settings[k]) st.settings[k] = st.settings[k].replace(/\{Name\}/g, '{Vorname}'); });
      st.migratedVorname = true;
    }
    if (!Array.isArray(st.projects)) st.projects = [];
    // Früherer "Projektbeginn" wird zu einem Projekt
    if (st.settings.projectStart) {
      if (!st.projects.length) st.projects.push({ id: 'p' + Date.now().toString(36), name: 'Projekt', start: st.settings.projectStart, end: '' });
      delete st.settings.projectStart;
    }
    return st;
  }

  function loadState() {
    try { return normalizeState(JSON.parse(localStorage.getItem(KEY))); } catch (e) { return blankState(); }
  }

  function loadTemplate() {
    try {
      const raw = localStorage.getItem(TKEY);
      if (!raw) return null;
      const tpl = X.deserialize(raw);
      try { tpl.info = X.inspect(tpl.parts); } catch (e) { /* gespeicherte Angaben behalten */ }
      return tpl;
    } catch (e) {
      return null;
    }
  }

  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) { toast('Speichern fehlgeschlagen – ist der Speicher voll?'); }
  }

  let state = loadState();
  let template = loadTemplate();
  let week = mondayOf(today());
  let followToday = true;
  let editing = null;

  // ---------- Tage ----------
  const getDay = key => state.days[key] || {};
  const hasTimes = d => !!(d.start || d.end);
  const isComplete = d => !!(d.start && d.end);

  // Angezeigte Stunden; eine halb eingetragene Pause wird ignoriert (dazu gibt es einen Hinweis).
  function hoursOf(d) {
    if (!isComplete(d)) return null;
    const { H } = X.calcDay(d.pStart && d.pEnd ? d : { start: d.start, end: d.end });
    return typeof H === 'number' ? H : null;
  }

  // Catering: an einem Drehtag automatisch an; ein bewusstes An- oder Abwählen wird gespeichert und hat Vorrang.
  const cateringOf = d => (typeof d.catering === 'boolean' ? d.catering : d.type === 'Drehtag');

  function setDay(key, patch) {
    const d = Object.assign({}, state.days[key], patch);
    for (const k of Object.keys(d)) if (d[k] === '' || d[k] == null || (d[k] === false && k !== 'catering')) delete d[k];
    if (Object.keys(d).length) state.days[key] = d;
    else delete state.days[key];
    state.edited[mondayOf(key)] = new Date().toISOString();
    save();
  }

  function dayIssues(d) {
    const out = [];
    if (d.start && !d.end) out.push('Arbeitsende fehlt');
    if (!d.start && d.end) out.push('Arbeitsbeginn fehlt');
    if (d.pStart && !d.pEnd) out.push('Pausenende fehlt');
    if (!d.pStart && d.pEnd) out.push('Pausenbeginn fehlt');
    if ((d.pStart || d.pEnd) && !hasTimes(d)) out.push('Pause ohne Arbeitszeit');
    if (d.pStart && d.pEnd && d.pEnd < d.pStart) out.push('Pause über Mitternacht – die Vorlage kann das nicht richtig rechnen');
    const h = hoursOf(d);
    if (h != null && h < 0) out.push('Pause ist länger als die Arbeitszeit');
    if (isComplete(d) && d.pStart && d.pEnd && d.end >= d.start && (d.pStart < d.start || d.pEnd > d.end)) out.push('Pause liegt außerhalb der Arbeitszeit');
    if (h != null && h > 6 && !(d.pStart && d.pEnd)) out.push('Über 6 Std. ohne Pause');
    return out;
  }

  function weekStats(monday) {
    let total = 0, days = 0;
    for (let i = 0; i < 7; i++) {
      const h = hoursOf(getDay(addDays(monday, i)));
      if (h != null) { total += h; days++; }
    }
    return { total, days };
  }

  const weekHasEntries = monday => [0, 1, 2, 3, 4, 5, 6].some(i => Object.keys(getDay(addDays(monday, i))).some(k => k !== 'catering'));

  function weekIssues(monday) {
    const out = [];
    if (!template) out.push({ block: true, text: 'Noch keine Vorlage geladen (Einstellungen → Vorlage).' });
    if (!state.settings.name.trim()) out.push({ text: 'Dein Name fehlt (Einstellungen).' });
    if (!state.settings.position.trim()) out.push({ text: 'Deine Position fehlt (Einstellungen).' });
    for (let i = 0; i < 7; i++) dayIssues(getDay(addDays(monday, i))).forEach(t => out.push({ text: DAY_SHORT[i] + ': ' + t }));
    if (!weekHasEntries(monday)) out.push({ text: 'In dieser Woche ist noch nichts eingetragen.' });
    return out;
  }

  // ---------- Hauptansicht ----------
  const isStandalone = () => window.navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches;
  const isIOS = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  if (isIOS) document.documentElement.classList.add('ios');

  function render() {
    const info = template && template.info;
    $('#prodName').textContent = info ? [info.production, info.productionNo].filter(Boolean).join(' · ') : 'Arbeitszeiten erfassen';
    renderBanners();
    renderWeek();
  }

  function renderBanners() {
    const out = [];
    if (isIOS && !isStandalone()) {
      out.push('<div class="banner"><div class="b-icon">' + SHARE_ICON + '</div><div><b>Als App installieren</b>' +
        '<p>Unten auf <b>Teilen</b> tippen, dann <b>Zum Home-Bildschirm</b>. Danach die App über das neue Icon öffnen und dort einrichten – Safari und die App speichern getrennt.</p></div></div>');
    }
    const steps = [];
    if (!template) steps.push(['Excel-Vorlage laden', '<label class="pill" for="templateInput">Laden</label>']);
    if (!state.settings.name.trim()) steps.push(['Name und Position eintragen', '<button class="pill" data-action="open-settings">Öffnen</button>']);
    if (steps.length) {
      out.push('<div class="banner setup"><b>Einrichten</b><ol class="steps">' +
        steps.map((s, i) => '<li><span class="num">' + (i + 1) + '</span><span>' + s[0] + '</span>' + s[1] + '</li>').join('') +
        '</ol></div>');
    }
    const current = mondayOf(today());
    const previous = addDays(current, -7);
    if (week === current && weekHasEntries(previous) && !state.sent[previous]) {
      out.push('<div class="banner warn"><div><b>KW ' + isoWeek(previous).week + ' ist noch nicht gesendet</b>' +
        '<p>Stundenzettel sind jeden Montag abzugeben.</p></div><button class="pill" data-action="go-week" data-week="' + previous + '">Öffnen</button></div>');
    }
    $('#banners').innerHTML = out.join('');
  }

  function renderWeek() {
    $('#kw').textContent = 'KW ' + isoWeek(week).week;
    $('#range').textContent = weekRange(week);
    $('#btnThisWeek').hidden = week === mondayOf(today());

    const stats = weekStats(week);
    $('#total').textContent = fmtH(stats.total) + ' Std.';
    $('#dayCount').textContent = stats.days === 1 ? '1 Arbeitstag' : stats.days + ' Arbeitstage';
    const sent = state.sent[week], edited = state.edited[week];
    $('#status').innerHTML = !sent ? ''
      : edited && edited > sent ? '<span class="chip warn">Nach dem Senden geändert</span>'
        : '<span class="chip ok">✓ Gesendet ' + esc(stamp(sent)) + '</span>';

    const t = today();
    let html = '';
    for (let i = 0; i < 7; i++) {
      const key = addDays(week, i);
      const d = getDay(key);
      const h = hoursOf(d);
      const issues = dayIssues(d);
      let main;
      if (hasTimes(d)) {
        const overnight = isComplete(d) && d.end < d.start ? ' <span class="plus1">+1</span>' : '';
        main = '<span class="d-times">' + (d.start || '––:––') + ' – ' + (d.end || '––:––') + overnight + '</span>';
      } else {
        const label = [d.type, d.note].filter(Boolean).map(esc).join(' · ');
        main = label ? '<span class="d-label">' + label + '</span>' : '<span class="d-empty">Kein Eintrag</span>';
      }
      const sub = [];
      if (d.type && hasTimes(d)) sub.push(esc(d.type));
      if (d.pStart || d.pEnd) sub.push('Pause ' + (d.pStart || '––:––') + '–' + (d.pEnd || '––:––'));
      if (template && template.info.catering && cateringOf(d)) sub.push('Catering');
      if (d.note && hasTimes(d)) sub.push(esc(d.note));
      html += '<li><button class="day' + (key === t ? ' today' : '') + (i > 4 ? ' weekend' : '') + '" data-action="edit-day" data-day="' + key + '">' +
        '<span class="d-date"><span class="wd">' + DAY_SHORT[i] + '</span><span class="dt">' + dm(key) + '</span></span>' +
        '<span class="d-main">' + main + (sub.length ? '<span class="d-sub">' + sub.join(' · ') + '</span>' : '') + '</span>' +
        '<span class="d-hours">' + (issues.length ? '<span class="dot" title="' + esc(issues.join(', ')) + '"></span>' : '') +
        (h != null ? fmtH(h) : '') + CHEVRON + '</span></button></li>';
    }
    $('#days').innerHTML = html;
  }

  // ---------- Tag bearbeiten ----------
  function openDay(key) {
    editing = key;
    fillDay();
    openOverlay('#dayOverlay');
  }

  function setTimeInput(field, value) {
    const b = $('#f-' + field);
    b.textContent = value || '--:--';
    b.classList.toggle('empty', !value);
  }

  function fillDay() {
    const d = getDay(editing);
    closePicker();
    $('#dayTitle').textContent = longDay(editing);
    ['start', 'end', 'pStart', 'pEnd'].forEach(f => setTimeInput(f, d[f]));
    $('#f-type').value = d.type || '';
    $('#f-type').classList.toggle('empty', !d.type);
    $('#f-note').value = d.note || '';
    $('#f-catering').checked = cateringOf(d);
    $('#cateringRow').hidden = !(template && template.info.catering);
    const notes = [];
    Object.keys(state.days).sort().reverse().forEach(k => { const n = (state.days[k].note || '').trim(); if (n && !notes.includes(n)) notes.push(n); });
    $('#noteList').innerHTML = notes.slice(0, 15).map(n => '<option value="' + esc(n) + '"></option>').join('');
    updateDayCalc();
  }

  function updateDayCalc() {
    const d = getDay(editing);
    const h = hoursOf(d);
    if (h != null) {
      const pause = d.pStart && d.pEnd ? X.calcDay(d).G : 0;
      const extra = [hm(h) + ' h'];
      if (pause) extra.push('Pause ' + hm(pause) + ' h');
      if (d.end < d.start) extra.push('endet am Folgetag');
      $('#dayCalc').innerHTML = '<span>Arbeitszeit</span><b>' + fmtH(h) + ' Std.</b><small>' + extra.join(' · ') + '</small>';
    } else {
      $('#dayCalc').innerHTML = '<span>Arbeitszeit</span><b>–</b>';
    }
    const issues = dayIssues(d);
    $('#dayWarn').innerHTML = issues.map(t => '<li>' + esc(t) + '</li>').join('');
    $('#dayWarn').hidden = !issues.length;
  }

  function changeDay(patch) {
    setDay(editing, patch);
    updateDayCalc();
    renderWeek();
  }

  // ---------- Uhrzeit-Rad in 15-Minuten-Schritten ----------
  // Zwei Räder (Stunde, Viertelstunde); der Wert wird übernommen, sobald das Rad stehen bleibt.
  let picking = null;
  const picker = $('#timePicker');
  const WHEEL_ITEM = 40;
  const HOURS = Array.from({ length: 24 }, (_, h) => pad(h));
  const QUARTERS = ['00', '15', '30', '45'];
  const wheelH = $('#wheelH'), wheelM = $('#wheelM');
  let wheelSilent = false;

  function fillWheel(col, values) {
    col.innerHTML = '<div class="wheel-pad"></div>' +
      values.map((v, i) => '<div class="wheel-item" data-index="' + i + '">' + v + '</div>').join('') + '<div class="wheel-pad"></div>';
  }
  fillWheel(wheelH, HOURS);
  fillWheel(wheelM, QUARTERS);

  const wheelIndex = (col, count) => Math.max(0, Math.min(count - 1, Math.round(col.scrollTop / WHEEL_ITEM)));
  const wheelValue = () => HOURS[wheelIndex(wheelH, 24)] + ':' + QUARTERS[wheelIndex(wheelM, 4)];

  function markWheel(col, count) {
    const i = wheelIndex(col, count);
    col.querySelectorAll('.wheel-item').forEach((el, j) => el.classList.toggle('sel', j === i));
  }

  // Startwert für ein leeres Feld (wird erst beim Drehen oder mit "Fertig" übernommen)
  function suggestedTime(field) {
    const d = getDay(editing);
    if (field === 'end' && d.start) return addMinutes(d.start, 600);
    if (field === 'pStart' && d.start) return addMinutes(d.start, 300);
    if (field === 'pEnd' && d.pStart) return addMinutes(d.pStart, 30);
    return nowTime();
  }

  function setWheel(value) {
    const total = +value.slice(0, 2) * 60 + Math.round(+value.slice(3, 5) / 15) * 15;
    wheelSilent = true;
    wheelH.scrollTop = (Math.floor(total / 60) % 24) * WHEEL_ITEM;
    wheelM.scrollTop = ((total % 60) / 15) * WHEEL_ITEM;
    markWheel(wheelH, 24);
    markWheel(wheelM, 4);
    setTimeout(() => { wheelSilent = false; }, 100);
  }

  function commitWheel() {
    if (!picking) return;
    const v = wheelValue();
    if ((getDay(editing)[picking] || '') !== v) setTime(picking, v);
  }

  [[wheelH, 24], [wheelM, 4]].forEach(([col, count]) => {
    let timer;
    col.addEventListener('scroll', () => {
      markWheel(col, count);
      if (wheelSilent) return;
      clearTimeout(timer);
      timer = setTimeout(commitWheel, 150);
    }, { passive: true });
    col.addEventListener('click', e => {
      const item = e.target.closest('.wheel-item');
      if (item) col.scrollTo({ top: +item.dataset.index * WHEEL_ITEM, behavior: 'smooth' });
    });
  });

  function openPicker(field) {
    if (picking === field) { closePicker(); return; }
    picking = field;
    $('#f-' + field).closest('.row').after(picker);
    picker.hidden = false;
    $$('.tbtn').forEach(b => b.classList.toggle('active', b.dataset.field === field));
    renderPicker();
    picker.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  function closePicker() {
    picking = null;
    picker.hidden = true;
    $$('.tbtn').forEach(b => b.classList.remove('active'));
  }

  function renderPicker() {
    if (!picking) return;
    setWheel(getDay(editing)[picking] || suggestedTime(picking));
  }

  function setTime(field, value) {
    setTimeInput(field, value);
    const patch = { [field]: value };
    // Ein Arbeitstag ohne Tagesart ist standardmäßig ein Drehtag.
    if (value && !getDay(editing).type) patch.type = 'Drehtag';
    changeDay(patch);
    if (patch.type) {
      $('#f-type').value = patch.type;
      $('#f-type').classList.remove('empty');
    }
    $('#f-catering').checked = cateringOf(getDay(editing));
    if (picking === field) renderPicker();
  }

  $('#f-type').innerHTML = '<option value="">Auswählen</option>' + DAY_TYPES.map(t => '<option>' + esc(t) + '</option>').join('');
  $('#f-type').addEventListener('change', e => {
    e.target.classList.toggle('empty', !e.target.value);
    changeDay({ type: e.target.value });
    $('#f-catering').checked = cateringOf(getDay(editing));
  });
  $('#f-note').addEventListener('input', e => changeDay({ note: e.target.value.trim() ? e.target.value : '' }));
  $('#f-note').addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
  $('#f-catering').addEventListener('change', e => changeDay({ catering: e.target.checked }));

  function stepDay(n) {
    editing = addDays(editing, n);
    if (mondayOf(editing) !== week) { week = mondayOf(editing); followToday = false; render(); }
    fillDay();
  }

  // ---------- Senden ----------
  // Lesbarer Dateiname, z. B. "Stundenzettel KW 42 2026 - Max Mustermann.xlsx"
  function fileName(monday) {
    const { week: kw, year } = isoWeek(monday);
    const name = state.settings.name.replace(/[^\p{L}\p{N} -]+/gu, ' ').replace(/\s+/g, ' ').trim();
    return 'Stundenzettel KW ' + kw + ' ' + year + (name ? ' - ' + name : '') + '.xlsx';
  }

  // Mail-Vorlage mit Platzhaltern: {Empfänger}, {Vorname}, {Name}, {KW}, {Zeitraum}, {Stunden}, {Position}
  function fillMail(text, monday) {
    const s = state.settings;
    const toName = s.toName.trim();
    const name = s.name.trim();
    const values = {
      KW: isoWeek(monday).week,
      Zeitraum: weekRange(monday),
      Stunden: fmtH(weekStats(monday).total),
      Vorname: name.split(/\s+/)[0] || '',
      Name: name,
      Position: s.position.trim()
    };
    return text
      .replace(/( ?)\{Empfänger\}/g, (m, space) => (toName ? space + toName : ''))
      .replace(/\{(KW|Zeitraum|Stunden|Vorname|Name|Position)\}/g, (m, k) => values[k]);
  }

  function mailText(monday) {
    const s = state.settings;
    return {
      subject: fillMail(s.mailSubject || MAIL_DEFAULTS.mailSubject, monday).replace(/\s+[–-]\s*$/, '').trim(),
      body: fillMail(s.mailBody || MAIL_DEFAULTS.mailBody, monday).trim()
    };
  }

  // Tagesart und Bemerkung landen zusammen in "Bemerkungen / Notiz"
  const remark = d => [d.type, (d.note || '').trim()].filter(Boolean).join(' – ');

  function makeXlsx(monday) {
    const sig = state.signature;
    const days = [];
    for (let i = 0; i < 7; i++) {
      const d = getDay(addDays(monday, i));
      days.push(Object.assign({}, d, { note: remark(d), catering: cateringOf(d) }));
    }
    const blob = X.build(template, {
      name: state.settings.name.trim(),
      position: state.settings.position.trim(),
      monday,
      days,
      signature: sig ? { png: X.fromBase64(sig.dataUrl.split(',')[1]), w: sig.w, h: sig.h } : null,
      signDate: today()
    });
    return new File([blob], fileName(monday), { type: X.XLSX_MIME });
  }

  function makePdf(monday) {
    const info = template.info;
    const sig = state.signature;
    const longDate = s => longDay(s) + ' ' + parseYmd(s).getFullYear();
    const days = [];
    for (let i = 0; i < 7; i++) {
      const key = addDays(monday, i), d = getDay(key), h = hoursOf(d);
      days.push({
        weekday: DAY_NAMES[i], date: dmy(key),
        start: d.start || '', end: d.end || '', pStart: d.pStart || '', pEnd: d.pEnd || '',
        pause: d.pStart && d.pEnd ? fmtH(X.calcDay(d).G) : '',
        hours: h != null ? fmtH(h) : '',
        note: remark(d),
        catering: withCatering() && cateringOf(d)
      });
    }
    const blob = window.StundenzettelPDF.build({
      title: fileName(monday).replace(/\.xlsx$/, ''),
      company: info.company, address: info.address || [], production: info.production, productionNo: info.productionNo,
      name: state.settings.name.trim(), position: state.settings.position.trim(),
      kwLabel: 'KW ' + isoWeek(monday).week + ' · ' + weekRange(monday),
      weekLabel: longDate(monday) + ' – ' + longDate(addDays(monday, 6)),
      days, total: fmtH(weekStats(monday).total), cateringColumn: withCatering(),
      signDate: sig ? dmy(today()) : '',
      signature: sig && sig.jpeg ? { jpeg: X.fromBase64(sig.jpeg.split(',')[1]), w: sig.w, h: sig.h } : null
    });
    return new File([blob], fileName(monday).replace(/\.xlsx$/, '.pdf'), { type: 'application/pdf' });
  }

  // Je nach Auswahl beim Senden: Excel, PDF oder beides
  function makeFiles(monday) {
    const format = state.settings.sendFormat || 'xlsx';
    const files = [];
    if (format !== 'pdf') files.push(makeXlsx(monday));
    if (format !== 'xlsx') files.push(makePdf(monday));
    return files;
  }

  function renderSendFormat() {
    const format = state.settings.sendFormat || 'xlsx';
    $$('#formatSeg button').forEach(b => b.classList.toggle('active', b.dataset.format === format));
    const xlsx = fileName(week), pdf = xlsx.replace(/\.xlsx$/, '.pdf');
    $('#sendFile').textContent = (format === 'pdf' ? [pdf] : format === 'both' ? [xlsx, pdf] : [xlsx]).join('\n');
  }

  function download(file) {
    const url = URL.createObjectURL(file);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function openSend() {
    const stats = weekStats(week);
    const issues = weekIssues(week);
    $('#sendTitle').textContent = 'Stundenzettel KW ' + isoWeek(week).week;
    $('#sendRange').textContent = weekRange(week);
    $('#sendTotal').textContent = fmtH(stats.total);
    $('#sendDays').textContent = stats.days;
    $('#sendIssues').innerHTML = issues.map(i => '<li class="' + (i.block ? 'block' : '') + '">' + esc(i.text) + '</li>').join('');
    $('#sendIssues').hidden = !issues.length;
    $('#sendOk').hidden = !!issues.length;
    const blocked = issues.some(i => i.block);
    $('#btnShare').disabled = blocked;
    $('#btnDownload').disabled = blocked;
    renderSendFormat();
    $('#sendSubject').textContent = mailText(week).subject;
    $('#sendSig').textContent = state.signature ? 'mit Datum ' + dmy(today()) : 'keine (Einstellungen)';
    const s = state.settings;
    $('#sendTo').hidden = !s.toEmail.trim();
    $('#sendToAddr').textContent = s.toEmail.trim();
    $('#btnMarkSent').textContent = state.sent[week] ? 'Markierung „gesendet“ entfernen' : 'Als gesendet markieren';
    openOverlay('#sendOverlay');
  }

  function markSent(monday, on) {
    if (on) state.sent[monday] = new Date().toISOString();
    else delete state.sent[monday];
    save();
    render();
  }

  async function share() {
    let files;
    try { files = makeFiles(week); } catch (e) { console.error(e); toast('Fehler beim Erstellen: ' + e.message); return; }
    const { subject, body } = mailText(week);
    // iOS übernimmt beim Teilen mit Datei keinen Betreff – darum liegt er in der Zwischenablage bereit.
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(subject).catch(() => {});
    if (!(navigator.canShare && navigator.canShare({ files }))) {
      files.forEach(download);
      toast('Teilen geht hier nicht – die Datei wurde gespeichert.');
      return;
    }
    try {
      await navigator.share({ files, title: subject, text: body });
      markSent(week, true);
      closeOverlay('#sendOverlay');
      toast('Stundenzettel verschickt');
    } catch (e) {
      if (e.name === 'AbortError') return;
      console.error(e);
      files.forEach(download);
      toast('Teilen nicht möglich – die Datei wurde gespeichert.');
    }
  }

  // ---------- Übersicht (Monat / Projekt) zum Abgleich mit der Abrechnung ----------
  const CHEV_LEFT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>';
  const CHEV_RIGHT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg>';
  let ovMode = 'month';
  let ovMonth = null;

  const monthKey = s => s.slice(0, 7);
  const monthLabel = ym => MONTHS[+ym.slice(5, 7) - 1] + ' ' + ym.slice(0, 4);
  const addMonths = (ym, n) => { const d = new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 1 + n, 1, 12); return d.getFullYear() + '-' + pad(d.getMonth() + 1); };
  const lastOfMonth = ym => ym + '-' + pad(new Date(+ym.slice(0, 4), +ym.slice(5, 7), 0).getDate());
  const hasEntry = d => !!d && Object.keys(d).some(k => k !== 'catering');
  const withCatering = () => !!(template && template.info.catering);

  function daysBetween(from, to) {
    const out = [];
    for (let k = from; k <= to; k = addDays(k, 1)) out.push(k);
    return out;
  }

  function summarize(keys) {
    const s = { hours: 0, workDays: 0, catering: 0, types: {}, days: [] };
    keys.forEach(k => {
      const d = state.days[k];
      if (!hasEntry(d)) return;
      const h = hoursOf(d);
      if (h != null) { s.hours += h; s.workDays++; }
      const type = d.type || (hasTimes(d) ? 'Ohne Tagesart' : '');
      if (type) s.types[type] = (s.types[type] || 0) + 1;
      if (withCatering() && cateringOf(d)) s.catering++;
      s.days.push(k);
    });
    return s;
  }

  function statsHtml(s) {
    return '<div class="stats three">' +
      '<div><b>' + fmtH(s.hours) + '</b><span>Stunden</span></div>' +
      '<div><b>' + s.workDays + '</b><span>Arbeitstage</span></div>' +
      (withCatering() ? '<div><b>' + s.catering + '</b><span>Catering</span></div>'
        : '<div><b>' + s.days.length + '</b><span>Tage mit Eintrag</span></div>') + '</div>';
  }

  function typesHtml(s) {
    const rows = DAY_TYPES.concat(['Ohne Tagesart']).filter(t => s.types[t]).map(t =>
      '<div class="row ov-row"><span class="label">' + esc(t) + '</span><span class="num">' + s.types[t] + '</span></div>');
    return '<div class="group-label">Tagesarten</div><div class="group">' + (rows.join('') || '<p class="ov-empty">Keine Tagesarten eingetragen.</p>') + '</div>';
  }

  function weekChip(monday) {
    if (state.sent[monday]) return '<span class="chip ok">gesendet</span>';
    return monday < mondayOf(today()) ? '<span class="chip warn">offen</span>' : '';
  }

  function dayLineHtml(k) {
    const d = state.days[k];
    const h = hoursOf(d);
    const info = [];
    if (hasTimes(d)) info.push((d.start || '––:––') + '–' + (d.end || '––:––'));
    if (d.pStart && d.pEnd) info.push('Pause ' + hm(X.calcDay(d).G));
    if (withCatering() && cateringOf(d)) info.push('Catering');
    if (d.note) info.push(esc(d.note));
    return '<div class="ov-day"><div class="l1"><span>' + DAY_SHORT[weekdayIndex(k)] + ' ' + dm(k) + ' · ' +
      esc(d.type || (hasTimes(d) ? 'Ohne Tagesart' : '–')) + '</span><b>' + (h != null ? fmtH(h) : '') + '</b></div>' +
      (info.length ? '<div class="l2">' + info.join(' · ') + '</div>' : '') + '</div>';
  }

  function renderMonth() {
    const ym = ovMonth, first = ym + '-01', last = lastOfMonth(ym);
    const s = summarize(daysBetween(first, last));
    let html = '<section class="week">' +
      '<button class="icon-btn" data-action="ov-month" data-step="-1" aria-label="Vorheriger Monat">' + CHEV_LEFT + '</button>' +
      '<div class="week-title"><div class="kw">' + monthLabel(ym) + '</div><div class="range muted">' + dm(first) + ' – ' + dmy(last) + '</div></div>' +
      '<button class="icon-btn" data-action="ov-month" data-step="1" aria-label="Nächster Monat">' + CHEV_RIGHT + '</button></section>';
    html += statsHtml(s);
    if (!s.days.length) return html + '<div class="group"><p class="ov-empty">In diesem Monat ist nichts eingetragen.</p></div>';
    html += typesHtml(s);
    const weeks = [];
    for (let mo = mondayOf(first); mo <= last; mo = addDays(mo, 7)) {
      const keys = daysBetween(mo, addDays(mo, 6)).filter(k => monthKey(k) === ym);
      const ws = summarize(keys);
      if (!ws.days.length) continue;
      weeks.push('<button class="row ov-row" data-action="ov-go-week" data-week="' + mo + '"><span class="label">KW ' + isoWeek(mo).week +
        '<small>' + weekRange(mo) + (keys.length < 7 ? ' · nur Tage im ' + MONTHS[+ym.slice(5, 7) - 1] : '') + '</small></span>' +
        '<span class="num">' + fmtH(ws.hours) + ' Std.</span>' + weekChip(mo) + '</button>');
    }
    html += '<div class="group-label">Wochen</div><div class="group">' + weeks.join('') + '</div>';
    html += '<div class="group-label">Tage</div><div class="group">' + s.days.map(dayLineHtml).join('') + '</div>';
    return html;
  }

  let ovProject = null; // Projekt-ID oder 'all'

  const entryKeys = () => Object.keys(state.days).filter(k => hasEntry(state.days[k])).sort();

  // Projekt, in dessen Zeitraum heute liegt – sonst das zuletzt angelegte
  function currentProjectId() {
    const t = today();
    const p = state.projects.find(x => (!x.start || x.start <= t) && (!x.end || t <= x.end)) || state.projects[state.projects.length - 1];
    return p ? p.id : 'all';
  }

  // Zeitraum eines Projekts (ohne Ende: bis heute bzw. bis zum letzten Eintrag); ohne Projekt: alle Einträge
  function projectRange(p) {
    const keys = entryKeys();
    if (!p) return keys.length ? { start: keys[0], end: keys[keys.length - 1] } : null;
    const start = p.start || keys[0] || today();
    let end = p.end;
    if (!end) {
      const last = keys[keys.length - 1] || today();
      end = last > today() ? last : today();
    }
    return { start, end: end < start ? start : end };
  }

  function createProject() {
    const keys = entryKeys();
    const first = !state.projects.length;
    const p = {
      id: 'p' + Date.now().toString(36),
      name: first && template && template.info.production ? template.info.production : 'Neues Projekt',
      start: first && keys.length ? keys[0] : today(),
      end: ''
    };
    state.projects.push(p);
    save();
    ovProject = p.id;
    toast('Projekt angelegt – Name und Zeitraum kannst du anpassen');
  }

  function renderProject() {
    const projects = state.projects;
    if (ovProject !== 'all' && !projects.some(x => x.id === ovProject)) ovProject = currentProjectId();
    const p = projects.find(x => x.id === ovProject) || null;
    let html = '<div class="group"><label class="row select-row" for="ovProjectSel"><span class="label">Projekt</span><select id="ovProjectSel">' +
      projects.map(x => '<option value="' + x.id + '"' + (x.id === ovProject ? ' selected' : '') + '>' + esc(x.name || 'Ohne Namen') + '</option>').join('') +
      '<option value="all"' + (ovProject === 'all' ? ' selected' : '') + '>Alle Einträge</option>' +
      '<option value="new">+ Neues Projekt</option></select><span class="select-chev" aria-hidden="true">›</span></label></div>';
    if (p) {
      html += '<div class="group">' +
        '<div class="row"><label for="pjName">Name</label><input id="pjName" type="text" value="' + esc(p.name) + '" placeholder="z. B. Staffel 20"></div>' +
        '<div class="row"><label for="pjStart">Beginn</label><input id="pjStart" type="date" value="' + (p.start || '') + '"></div>' +
        '<div class="row"><label for="pjEnd">Ende</label><input id="pjEnd" type="date" value="' + (p.end || '') + '"></div>' +
        '<button class="row-btn danger" data-action="pj-delete">Projekt löschen</button></div>' +
        '<p class="ov-note">Zum Projekt zählen alle Tage von Beginn bis Ende. Ohne Ende läuft es noch. Die Einträge selbst bleiben beim Löschen erhalten.</p>';
    } else if (!projects.length) {
      html += '<p class="ov-note">Lege über „+ Neues Projekt“ ein Projekt mit Name, Beginn und Ende an – dann werden seine Tage getrennt ausgewertet.</p>';
    }
    const r = projectRange(p);
    if (!r) return html + '<div class="group"><p class="ov-empty">Noch keine Einträge.</p></div>';
    html += '<div class="group"><div class="row"><span class="label">Zeitraum</span><span class="muted">' +
      (p && !p.end ? 'seit ' + dmy(r.start) : dmy(r.start) + ' – ' + dmy(r.end)) + '</span></div></div>';
    const s = summarize(daysBetween(r.start, r.end));
    html += statsHtml(s) + typesHtml(s);
    const months = [];
    for (let ym = monthKey(r.start); ym <= monthKey(r.end); ym = addMonths(ym, 1)) {
      const from = ym + '-01' < r.start ? r.start : ym + '-01';
      const to = lastOfMonth(ym) > r.end ? r.end : lastOfMonth(ym);
      const ms = summarize(daysBetween(from, to));
      if (!ms.days.length) continue;
      months.push('<button class="row ov-row" data-action="ov-open-month" data-month="' + ym + '"><span class="label">' + monthLabel(ym) +
        '<small>' + ms.workDays + ' Arbeitstage' + (withCatering() ? ' · ' + ms.catering + '× Catering' : '') + '</small></span>' +
        '<span class="num">' + fmtH(ms.hours) + ' Std.</span>' + CHEVRON + '</button>');
    }
    html += '<div class="group-label">Monate</div><div class="group">' + (months.join('') || '<p class="ov-empty">Keine Einträge in diesem Zeitraum.</p>') + '</div>';
    const current = mondayOf(today());
    const open = [];
    let total = 0, sent = 0;
    for (let mo = mondayOf(r.start); mo <= r.end; mo = addDays(mo, 7)) {
      if (!weekHasEntries(mo)) continue;
      total++;
      if (state.sent[mo]) sent++;
      else if (mo < current) {
        open.push('<button class="row ov-row" data-action="ov-go-week" data-week="' + mo + '"><span class="label">KW ' + isoWeek(mo).week +
          '<small>' + weekRange(mo) + '</small></span><span class="chip warn">offen</span></button>');
      }
    }
    html += '<div class="group-label">Gesendet: ' + sent + ' von ' + total + ' Wochen</div><div class="group">' +
      (open.join('') || '<p class="ov-empty">Alle abgeschlossenen Wochen sind als gesendet markiert.</p>') + '</div>';
    return html;
  }

  function renderOverview() {
    $$('#overviewOverlay .segmented button').forEach(b => b.classList.toggle('active', b.dataset.mode === ovMode));
    $('#ovBody').innerHTML = ovMode === 'month' ? renderMonth() : renderProject();
    const sel = $('#ovProjectSel');
    if (sel) {
      sel.addEventListener('change', () => {
        if (sel.value === 'new') createProject();
        else ovProject = sel.value;
        renderOverview();
      });
    }
    const p = ovMode === 'project' && state.projects.find(x => x.id === ovProject);
    if (p) {
      const name = $('#pjName'), start = $('#pjStart'), end = $('#pjEnd');
      name.addEventListener('input', () => { p.name = name.value; save(); });
      name.addEventListener('change', () => renderOverview());
      start.addEventListener('change', () => { p.start = start.value; save(); renderOverview(); });
      end.addEventListener('change', () => { p.end = end.value; save(); renderOverview(); });
    }
  }

  function openOverview() {
    ovMode = 'month';
    ovMonth = monthKey(followToday ? today() : week);
    if (!ovProject || (ovProject !== 'all' && !state.projects.some(x => x.id === ovProject))) ovProject = currentProjectId();
    openOverlay('#overviewOverlay');
    renderOverview();
  }

  // ---------- Einstellungen ----------
  function openSettings() {
    openOverlay('#settingsOverlay');
    $$('[data-setting]').forEach(el => { el.value = state.settings[el.dataset.setting] || MAIL_DEFAULTS[el.dataset.setting] || ''; });
    renderTemplateInfo();
    renderSignature();
  }

  $$('[data-setting]').forEach(el => el.addEventListener('input', () => {
    state.settings[el.dataset.setting] = el.value;
    save();
    render();
  }));

  function renderTemplateInfo() {
    if (!template) {
      $('#tplInfo').innerHTML = '<p class="tpl-empty">Noch keine Vorlage geladen.</p>';
      $('#tplButton').textContent = 'Vorlage laden';
      return;
    }
    const i = template.info;
    const bundled = template.source === 'bundled';
    $('#tplInfo').innerHTML = '<div class="tpl"><b>' + esc(i.production || 'Stundenzettel-Vorlage') + '</b>' +
      '<span>' + esc([i.productionNo, i.company].filter(Boolean).join(' · ')) + '</span>' +
      '<small>' + (bundled ? 'In der App enthalten' : esc(template.fileName) + ' · geladen ' + esc(stamp(template.importedAt))) + '</small></div>';
    $('#tplButton').textContent = 'Andere Vorlage laden';
    $('#tplBundled').hidden = bundled;
  }

  async function importTemplate(buffer, name) {
    const tpl = await X.readTemplate(buffer, name);
    tpl.source = 'user';
    localStorage.setItem(TKEY, X.serialize(tpl));
    template = tpl;
    render();
    renderTemplateInfo();
    toast('Vorlage geladen: ' + (tpl.info.production || name));
    return tpl.info;
  }

  // Die mitgelieferte Vorlage wird automatisch geladen und aktualisiert,
  // solange keine eigene Vorlage geladen wurde.
  async function loadBundledTemplate(force) {
    if (template && template.source !== 'bundled' && !force) return;
    let buffer;
    try {
      const res = await fetch(BUNDLED_TEMPLATE);
      if (!res.ok) return;
      buffer = await res.arrayBuffer();
    } catch (e) {
      return;
    }
    const sig = buffer.byteLength + '-' + X.crc32(new Uint8Array(buffer));
    if (!force && template && template.bundleSig === sig) return;
    try {
      const tpl = await X.readTemplate(buffer, 'Mitgelieferte Vorlage');
      tpl.source = 'bundled';
      tpl.bundleSig = sig;
      localStorage.setItem(TKEY, X.serialize(tpl));
      template = tpl;
      render();
      if (!$('#settingsOverlay').hidden) renderTemplateInfo();
      if (force) toast('Mitgelieferte Vorlage wird verwendet');
    } catch (e) {
      console.error(e);
    }
  }

  $('#templateInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try { await importTemplate(await file.arrayBuffer(), file.name); } catch (err) { console.error(err); alert(err.message || 'Die Datei konnte nicht gelesen werden.'); }
  });

  // ---------- Unterschrift ----------
  const sig = { ctx: null, drawing: false, last: null, dirty: false };
  const canvas = $('#sigCanvas');

  function renderSignature() {
    const has = !!state.signature;
    $('#sigPreview').hidden = !has;
    $('#sigPad').hidden = has;
    if (has) $('#sigImg').src = state.signature.dataUrl;
    else requestAnimationFrame(setupPad);
  }

  function setupPad() {
    const r = canvas.getBoundingClientRect();
    if (!r.width) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.round(r.width * dpr);
    canvas.height = Math.round(r.height * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.lineWidth = 3;
    ctx.strokeStyle = '#111827';
    sig.ctx = ctx;
    sig.dirty = false;
    $('#sigSave').disabled = true;
  }

  const point = e => { const r = canvas.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
  canvas.addEventListener('pointerdown', e => {
    if (!sig.ctx) setupPad();
    e.preventDefault();
    try { canvas.setPointerCapture(e.pointerId); } catch (err) { /* Zeiger schon weg */ }
    sig.drawing = true;
    sig.last = sig.lastMid = point(e);
    sig.ctx.beginPath();
    sig.ctx.moveTo(sig.last.x, sig.last.y);
    sig.ctx.lineTo(sig.last.x + 0.1, sig.last.y + 0.1);
    sig.ctx.stroke();
    sig.dirty = true;
    $('#sigSave').disabled = false;
  });
  canvas.addEventListener('pointermove', e => {
    if (!sig.drawing) return;
    const p = point(e);
    const mid = { x: (sig.last.x + p.x) / 2, y: (sig.last.y + p.y) / 2 };
    sig.ctx.beginPath();
    sig.ctx.moveTo(sig.lastMid.x, sig.lastMid.y);
    sig.ctx.quadraticCurveTo(sig.last.x, sig.last.y, mid.x, mid.y);
    sig.ctx.stroke();
    sig.last = p;
    sig.lastMid = mid;
  });
  const endStroke = () => {
    if (!sig.drawing) return;
    sig.drawing = false;
    sig.ctx.beginPath();
    sig.ctx.moveTo(sig.lastMid.x, sig.lastMid.y);
    sig.ctx.lineTo(sig.last.x, sig.last.y);
    sig.ctx.stroke();
  };
  canvas.addEventListener('pointerup', endStroke);
  canvas.addEventListener('pointercancel', endStroke);

  // PDF braucht die Unterschrift ohne Transparenz: auf weißem Grund als JPEG
  function signatureJpeg(source) {
    const c = document.createElement('canvas');
    c.width = source.width;
    c.height = source.height;
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(source, 0, 0, c.width, c.height);
    return c.toDataURL('image/jpeg', 0.92);
  }

  // Für Unterschriften aus älteren Versionen das JPEG nachträglich erzeugen
  function ensureSignatureJpeg() {
    const sg = state.signature;
    if (!sg || sg.jpeg) return;
    const img = new Image();
    img.onload = () => {
      if (state.signature !== sg) return;
      sg.jpeg = signatureJpeg(img);
      save();
    };
    img.src = sg.dataUrl;
  }

  function saveSignature() {
    const W = canvas.width, H = canvas.height;
    const px = sig.ctx.getImageData(0, 0, W, H).data;
    let x0 = W, y0 = H, x1 = -1, y1 = -1;
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        if (px[(y * W + x) * 4 + 3] > 10) {
          if (x < x0) x0 = x;
          if (x > x1) x1 = x;
          if (y < y0) y0 = y;
          if (y > y1) y1 = y;
        }
      }
    }
    if (x1 < 0) { toast('Bitte zuerst im Feld unterschreiben.'); return; }
    const m = Math.round(6 * (window.devicePixelRatio || 1));
    x0 = Math.max(0, x0 - m); y0 = Math.max(0, y0 - m);
    x1 = Math.min(W - 1, x1 + m); y1 = Math.min(H - 1, y1 + m);
    const w = x1 - x0 + 1, h = y1 - y0 + 1;
    const scale = Math.min(1, 200 / h);
    const out = document.createElement('canvas');
    out.width = Math.max(1, Math.round(w * scale));
    out.height = Math.max(1, Math.round(h * scale));
    const octx = out.getContext('2d');
    octx.imageSmoothingQuality = 'high';
    octx.drawImage(canvas, x0, y0, w, h, 0, 0, out.width, out.height);
    state.signature = { dataUrl: out.toDataURL('image/png'), jpeg: signatureJpeg(out), w: out.width, h: out.height };
    save();
    renderSignature();
    toast('Unterschrift gespeichert');
  }

  // ---------- Backup ----------
  async function exportBackup() {
    const payload = JSON.stringify({ app: 'stundenzettel', version: 1, exportedAt: new Date().toISOString(), state, template: localStorage.getItem(TKEY) });
    const file = new File([payload], 'Stundenzettel-Backup-' + today() + '.json', { type: 'application/json' });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file] }); } catch (e) { if (e.name !== 'AbortError') download(file); }
    } else {
      download(file);
    }
  }

  $('#backupInput').addEventListener('change', async e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try {
      const o = JSON.parse(await file.text());
      if (!o || o.app !== 'stundenzettel' || !o.state) throw new Error('Das ist kein Stundenzettel-Backup.');
      if (!confirm('Backup vom ' + stamp(o.exportedAt) + ' laden? Die aktuellen Einträge in der App werden dabei ersetzt.')) return;
      state = normalizeState(o.state);
      save();
      if (o.template) { localStorage.setItem(TKEY, o.template); template = loadTemplate(); }
      ensureSignatureJpeg();
      openSettings();
      render();
      toast('Backup geladen');
    } catch (err) {
      alert(err.message || 'Das Backup konnte nicht gelesen werden.');
    }
  });

  // ---------- Overlays & Meldungen ----------
  function openOverlay(sel) {
    const o = $(sel);
    o.hidden = false;
    void o.offsetWidth;
    o.classList.add('open');
    document.body.classList.add('locked');
  }

  function closeOverlay(sel) {
    const o = $(sel);
    if (o.hidden) return;
    if (document.activeElement && o.contains(document.activeElement)) document.activeElement.blur();
    o.classList.remove('open');
    setTimeout(() => {
      o.hidden = true;
      const sheet = o.querySelector('.sheet');
      sheet.style.transition = '';
      sheet.style.transform = '';
      if (!document.querySelector('.overlay.open')) document.body.classList.remove('locked');
    }, 220);
  }

  $$('.overlay').forEach(o => o.addEventListener('click', e => { if (e.target === o) closeOverlay('#' + o.id); }));

  // Fenster nach unten wegwischen (wenn ganz nach oben gescrollt)
  function enableSwipeToClose(overlay) {
    const sheet = overlay.querySelector('.sheet');
    let startX = 0, startY = 0, lastY = 0, startTime = 0, tracking = false, dragging = false;
    sheet.addEventListener('touchstart', e => {
      tracking = false;
      if (e.touches.length !== 1 || sheet.scrollTop > 0 || e.target.closest('input, textarea, select, canvas, .wheel')) return;
      tracking = true;
      dragging = false;
      startX = e.touches[0].clientX;
      startY = lastY = e.touches[0].clientY;
      startTime = Date.now();
    }, { passive: true });
    sheet.addEventListener('touchmove', e => {
      if (!tracking) return;
      const t = e.touches[0];
      const dx = t.clientX - startX, dy = t.clientY - startY;
      if (!dragging) {
        if (Math.abs(dx) > Math.abs(dy) || dy < 0 || sheet.scrollTop > 0) {
          if (Math.abs(dx) > 8 || dy < -8) tracking = false;
          return;
        }
        if (dy < 8) return;
        dragging = true;
        sheet.style.transition = 'none';
      }
      e.preventDefault();
      lastY = t.clientY;
      sheet.style.transform = 'translateY(' + Math.max(0, lastY - startY) + 'px)';
    }, { passive: false });
    const end = () => {
      const wasDragging = tracking && dragging;
      tracking = dragging = false;
      if (!wasDragging) return;
      const dy = lastY - startY;
      const speed = dy / Math.max(1, Date.now() - startTime);
      sheet.style.transition = 'transform 0.2s ease';
      if (dy > Math.min(140, sheet.offsetHeight * 0.3) || (speed > 0.5 && dy > 30)) {
        sheet.style.transform = 'translateY(100%)';
        if (overlay.id === 'dayOverlay') closePicker();
        closeOverlay('#' + overlay.id);
      } else {
        sheet.style.transform = '';
      }
    };
    sheet.addEventListener('touchend', end);
    sheet.addEventListener('touchcancel', end);
  }
  $$('.overlay').forEach(enableSwipeToClose);
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    const open = $$('.overlay.open').pop();
    if (open) closeOverlay('#' + open.id);
  });

  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2800);
  }

  // ---------- Aktionen ----------
  const actions = {
    'prev-week': () => { week = addDays(week, -7); followToday = false; render(); },
    'next-week': () => { week = addDays(week, 7); followToday = false; render(); },
    'this-week': () => { week = mondayOf(today()); followToday = true; render(); },
    'go-week': el => { week = el.dataset.week; followToday = false; render(); },
    'edit-day': el => openDay(el.dataset.day),
    'prev-day': () => stepDay(-1),
    'next-day': () => stepDay(1),
    'close-day': () => { closePicker(); closeOverlay('#dayOverlay'); },
    'pick-time': el => openPicker(el.dataset.field),
    'tp-done': () => { commitWheel(); closePicker(); },
    'now': el => setTime(el.dataset.field, nowTime()),
    'clear-field': el => setTime(el.dataset.field, ''),
    'pause': el => {
      const start = getDay(editing).pStart;
      if (!start) { toast('Bitte zuerst den Pausenbeginn eintragen.'); return; }
      setTime('pEnd', addMinutes(start, +el.dataset.min));
    },
    'copy-last': () => {
      for (let i = 1; i <= 21; i++) {
        const key = addDays(editing, -i);
        const d = state.days[key];
        if (d && hasTimes(d)) {
          changeDay({ type: d.type || 'Drehtag', start: d.start || '', end: d.end || '', pStart: d.pStart || '', pEnd: d.pEnd || '', catering: d.catering });
          fillDay();
          toast('Von ' + DAY_SHORT[weekdayIndex(key)] + ', ' + dm(key) + ' übernommen');
          return;
        }
      }
      toast('Keine früheren Zeiten gefunden.');
    },
    'clear-day': () => {
      if (!Object.keys(getDay(editing)).length) return;
      if (!confirm(longDay(editing) + ' leeren?')) return;
      changeDay({ type: '', start: '', end: '', pStart: '', pEnd: '', note: '', catering: null });
      fillDay();
    },
    'mail-reset': () => {
      delete state.settings.mailSubject;
      delete state.settings.mailBody;
      save();
      $('#s-mailSubject').value = MAIL_DEFAULTS.mailSubject;
      $('#s-mailBody').value = MAIL_DEFAULTS.mailBody;
      toast('Standardtext wiederhergestellt');
    },
    'open-overview': () => openOverview(),
    'close-overview': () => closeOverlay('#overviewOverlay'),
    'ov-mode': el => { ovMode = el.dataset.mode; renderOverview(); },
    'ov-month': el => { ovMonth = addMonths(ovMonth, +el.dataset.step); renderOverview(); },
    'ov-open-month': el => { ovMode = 'month'; ovMonth = el.dataset.month; renderOverview(); $('#overviewOverlay .sheet').scrollTop = 0; },
    'ov-go-week': el => { week = el.dataset.week; followToday = week === mondayOf(today()); closeOverlay('#overviewOverlay'); render(); },
    'open-settings': () => openSettings(),
    'close-settings': () => closeOverlay('#settingsOverlay'),
    'open-send': () => openSend(),
    'close-send': () => closeOverlay('#sendOverlay'),
    'copy-subject': async () => {
      try { await navigator.clipboard.writeText(mailText(week).subject); toast('Betreff kopiert – in der Mail einsetzen'); } catch (e) { toast(mailText(week).subject); }
    },
    'share': () => share(),
    'download': () => {
      try { makeFiles(week).forEach(download); toast('Datei gespeichert'); } catch (e) { console.error(e); toast('Fehler beim Erstellen: ' + e.message); }
    },
    'send-format': el => { state.settings.sendFormat = el.dataset.format; save(); renderSendFormat(); },
    'pj-delete': () => {
      const p = state.projects.find(x => x.id === ovProject);
      if (!p || !confirm('Projekt „' + (p.name || 'Ohne Namen') + '“ löschen? Die Einträge bleiben erhalten.')) return;
      state.projects = state.projects.filter(x => x !== p);
      save();
      ovProject = currentProjectId();
      renderOverview();
    },
    'toggle-sent': () => { markSent(week, !state.sent[week]); openSend(); },
    'copy-to': async () => {
      const addr = state.settings.toEmail.trim();
      try { await navigator.clipboard.writeText(addr); toast('Adresse kopiert – in Mail bei „An“ einfügen'); } catch (e) { toast(addr); }
    },
    'sig-clear': () => setupPad(),
    'sig-save': () => saveSignature(),
    'sig-redo': () => { state.signature = null; save(); renderSignature(); },
    'sig-remove': () => { if (confirm('Gespeicherte Unterschrift entfernen?')) { state.signature = null; save(); renderSignature(); } },
    'backup-export': () => exportBackup(),
    'use-bundled': () => loadBundledTemplate(true)
  };

  document.addEventListener('click', e => {
    const el = e.target.closest('[data-action]');
    if (!el || el.disabled) return;
    const fn = actions[el.dataset.action];
    if (fn) fn(el, e);
  });

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) return;
    if (followToday) week = mondayOf(today());
    render();
  });

  // ---------- Start ----------
  $('#version').textContent = VERSION;
  render();
  loadBundledTemplate();
  ensureSignatureJpeg();

  if ('serviceWorker' in navigator && (location.protocol === 'https:' || /[?&]sw\b/.test(location.search))) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});

  // Für Tests im Browser
  window.__stundenzettel = { importTemplate, makeFile: makeXlsx, makePdf, makeFiles, mailText, get state() { return state; }, get week() { return week; } };
})();
