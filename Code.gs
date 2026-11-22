/**
 * Workout-Coach – Backend (Google Apps Script)
 * Speichert Benutzer + Fortschritt in einer Google-Tabelle und
 * schickt täglich um 18 Uhr Erinnerungs-Mails (Mo–Sa), wenn das Workout fehlt.
 *
 * Einrichtung: einmal die Funktion setup() ausführen, dann
 * "Bereitstellen > Neue Bereitstellung > Web-App", Ausführen als: Ich,
 * Zugriff: Jeder. Die Web-App-URL in index.html bei API_URL eintragen.
 */
const APP_LINK = 'https://bobby983983.github.io/workout-coach/';
const PLAN = {1:'Tag A', 2:'Kurztag', 3:'Tag B', 4:'Kurztag', 5:'Tag A', 6:'Kurztag'};
const CH_START = '2026-12-01', CH_END = '2026-12-31';
const HEAD = ['name','nameKey','pinHash','salt','token','email','start','data','created','updated'];

function sheet_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('SHEET_ID');
  let ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Workout-Coach Daten');
    p.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName('Users');
  if (!sh) { sh = ss.getSheets()[0]; sh.setName('Users'); sh.getRange(1,1,1,HEAD.length).setValues([HEAD]); sh.setFrozenRows(1); }
  return sh;
}

function setup() {
  sheet_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(18).nearMinute(0).inTimezone('Europe/Berlin').create();
  ScriptApp.newTrigger('morningReminder').timeBased().everyDays(1).atHour(7).nearMinute(0).inTimezone('Europe/Berlin').create();
  ScriptApp.newTrigger('weeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(18).nearMinute(0).inTimezone('Europe/Berlin').create();
  Logger.log('Fertig. Tabelle: ' + SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getUrl());
}

function hash_(pin, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}
function rows_() {
  const sh = sheet_();
  const v = sh.getDataRange().getValues();
  return { sh, rows: v.slice(1).map((r, i) => ({ row: i + 2, o: Object.fromEntries(HEAD.map((h, j) => [h, r[j]])) })) };
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today_() { return Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd'); }
function validEmail_(e) { return !e || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }

function doGet() { return out_({ ok: true, app: 'workout-coach' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const q = JSON.parse(e.postData.contents || '{}');
    const { sh, rows } = rows_();
    const byToken = t => rows.find(r => t && r.o.token === t);
    const key = String(q.name || '').trim().toLowerCase();

    if (q.a === 'register') {
      const name = String(q.name || '').trim().slice(0, 20);
      const pin = String(q.pin || '');
      const email = String(q.email || '').trim().slice(0, 100);
      if (name.length < 2) return out_({ ok: false, err: 'Benutzername zu kurz' });
      if (!/^\d{4}$/.test(pin)) return out_({ ok: false, err: 'PIN muss 4 Ziffern haben' });
      if (!validEmail_(email)) return out_({ ok: false, err: 'E-Mail ungültig' });
      if (rows.some(r => r.o.nameKey === key)) return out_({ ok: false, err: 'Benutzername ist schon vergeben' });
      const salt = Utilities.getUuid(), token = Utilities.getUuid();
      const t = today_();
      const start = t < '2026-11-23' ? '2026-11-23' : t;
      const data = JSON.stringify({ days: {}, ms: [], ch: {} });
      sh.appendRow([name, key, hash_(pin, salt), salt, token, email, start, data, new Date(), new Date()]);
      return out_({ ok: true, token, name, email, start, data: JSON.parse(data) });
    }
    if (q.a === 'login') {
      const r = rows.find(x => x.o.nameKey === key);
      if (!r || hash_(String(q.pin || ''), r.o.salt) !== r.o.pinHash) return out_({ ok: false, err: 'Name oder PIN falsch' });
      return out_({ ok: true, token: r.o.token, name: r.o.name, email: r.o.email, start: r.o.start, data: JSON.parse(r.o.data || '{}') });
    }
    if (q.a === 'me') {
      const r = byToken(q.token);
      if (!r) return out_({ ok: false, err: 'auth' });
      return out_({ ok: true, name: r.o.name, email: r.o.email, start: r.o.start, data: JSON.parse(r.o.data || '{}') });
    }
    if (q.a === 'save') {
      const r = byToken(q.token);
      if (!r) return out_({ ok: false, err: 'auth' });
      const d = q.data || {};
      const clean = JSON.stringify({ days: d.days || {}, ms: d.ms || [], ch: d.ch || {} });
      if (clean.length > 45000) return out_({ ok: false, err: 'zu groß' });
      sh.getRange(r.row, HEAD.indexOf('data') + 1).setValue(clean);
      sh.getRange(r.row, HEAD.indexOf('updated') + 1).setValue(new Date());
      return out_({ ok: true });
    }
    if (q.a === 'email') {
      const r = byToken(q.token);
      if (!r) return out_({ ok: false, err: 'auth' });
      const email = String(q.email || '').trim().slice(0, 100);
      if (!validEmail_(email)) return out_({ ok: false, err: 'E-Mail ungültig' });
      sh.getRange(r.row, HEAD.indexOf('email') + 1).setValue(email);
      return out_({ ok: true, email });
    }
    if (q.a === 'list') {
      return out_({ ok: true, users: rows.map(r => {
        let d = {}; try { d = JSON.parse(r.o.data || '{}'); } catch (e) {}
        return { name: r.o.name, start: r.o.start, days: d.days || {}, ch: d.ch || {} };
      }) });
    }
    return out_({ ok: false, err: 'unbekannte Aktion' });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function addDays_(k, n) {
  const d = new Date(k + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
function dow_(k) { return new Date(k + 'T12:00:00Z').getUTCDay(); }

// gleiche Punkte-Logik wie in der App
function compute_(o, tt) {
  const t = tt > CH_END ? CH_END : tt;
  let data = {}; try { data = JSON.parse(o.data || '{}'); } catch (e) {}
  const days = data.days || {}, ch = data.ch || {};
  let start = o.start || '2026-11-23';
  if (tt >= CH_START && start < CH_START) start = CH_START;
  let pts = 0, streak = 0, total = 0, weekRun = 0;
  for (let k = start; k <= t; k = addDays_(k, 1)) {
    const d = days[k], rest = dow_(k) === 0, past = k < t;
    let st = 'open';
    if (d && d.done) { total++; pts += (d.at === k) ? 1 : -1; st = 'done'; if (!rest) { streak++; weekRun++; } }
    else if (rest) st = 'rest';
    else if (past) { pts -= 3; streak = 0; weekRun = 0; st = 'miss'; }
    if (dow_(k) === 6) { if (weekRun >= 6 && st !== 'open') pts += 3; weekRun = 0; }
  }
  Object.keys(ch).forEach(x => { const c = ch[x]; if (c && (c.at || '') >= start && (c.at || '') <= CH_END) pts += c.p || 0; });
  return { pts, streak, total, done: !!(days[tt] && days[tt].done), days };
}

function ranking_(rows, t) {
  return rows.filter(r => r.o.name).map(r => Object.assign({ row: r }, compute_(r.o, t), { name: r.o.name }))
    .sort((a, b) => b.pts - a.pts || b.streak - a.streak);
}

function dailyReminder() {
  const t = today_();
  const dow = dow_(t);
  if (dow === 0 || t > CH_END) return; // Sonntag Ruhetag / Challenge vorbei
  const y = addDays_(t, -1);
  const { rows } = rows_();
  const rank = ranking_(rows, t);
  const medal = ['🥇', '🥈', '🥉'];
  rank.forEach((me, i) => {
    const o = me.row.o;
    if (!o.email || !o.start || t < o.start || me.done) return;
    const missedY = dow_(y) !== 0 && y >= o.start && !(me.days[y] && me.days[y].done);
    let body = 'Hey ' + o.name + ',\n\n' +
      'dein Workout für heute (' + PLAN[dow] + ') ist noch offen! 😢\n' +
      (t < CH_START ? '🧪 Testphase – ab 1.12. zählt\'s richtig.\n' : 'Ausreden gibt\'s keine – sonst morgen −3 Punkte.\n');
    if (missedY) body += '\nGestern hast du auch verpasst. Heute nachholen bringt +2 zurück (2 Einheiten machen).\n';
    body += '\n🏆 Ranking\n';
    rank.slice(0, 3).forEach((r, j) => { body += medal[j] + ' ' + r.name + ' – ' + r.pts + ' Punkte (🔥 ' + r.streak + ')' + (r.done ? ' ✓ heute' : '') + '\n'; });
    body += '\nDu: Platz ' + (i + 1) + ' von ' + rank.length + ' mit ' + me.pts + ' Punkten.\n';
    if (i > 0) body += 'Bis Platz ' + i + ' fehlen dir ' + (rank[i - 1].pts - me.pts + 1) + ' Punkte.\n';
    body += '\nZum Coach: ' + APP_LINK + '\n\nZieh durch! 💪';
    MailApp.sendEmail({ to: o.email, subject: '😢 ' + o.name + ', vergiss nicht dein Workout!', body: body, name: 'Workout-Coach' });
  });
}

function weeklySummary() {
  const t = today_();
  if (t > addDays_(CH_END, 1)) return;
  const { rows } = rows_();
  const rank = ranking_(rows, t);
  const medal = ['🥇', '🥈', '🥉'];
  const mon = addDays_(t, -6);
  let table = '';
  rank.forEach((r, j) => { table += (medal[j] || (j + 1) + '.') + ' ' + r.name + ' – ' + r.pts + ' Punkte (🔥 ' + r.streak + ', ' + r.total + ' Workouts)\n'; });
  rank.forEach((me, i) => {
    const o = me.row.o;
    if (!o.email) return;
    let week = 0; for (let k = mon; k <= t; k = addDays_(k, 1)) if (me.days[k] && me.days[k].done) week++;
    let body = 'Hey ' + o.name + ',\n\nhier deine Wochenübersicht 💪\n\n' +
      (t < CH_START ? '🧪 Testphase – ab 1.12. startet die Challenge, dann beginnen alle bei 0 Punkten.\n\n' : '') +
      'Diese Woche: ' + week + ' Workouts\nDu: Platz ' + (i + 1) + ' von ' + rank.length + ' mit ' + me.pts + ' Punkten, Serie 🔥 ' + me.streak + '\n' +
      (i > 0 ? 'Bis Platz ' + i + ' fehlen dir ' + (rank[i - 1].pts - me.pts + 1) + ' Punkte.\n' : 'Du führst! Verteidige Platz 1. 🏆\n') +
      '\n🏆 Gesamt-Ranking\n' + table + '\nNeue Woche, neue Chance: ' + APP_LINK + '\n\nZieh durch! 💪';
    MailApp.sendEmail({ to: o.email, subject: '🏆 Workout-Challenge: deine Wochenübersicht', body: body, name: 'Workout-Coach' });
  });
}

function morningReminder() {
  const t = today_();
  const dow = dow_(t);
  if (dow === 0 || t > CH_END) return;
  const { rows } = rows_();
  const rank = ranking_(rows, t);
  rank.forEach((me, i) => {
    const o = me.row.o;
    if (!o.email || !o.start || t < o.start) return;
    const body = 'Guten Morgen ' + o.name + '! ☀️\n\nDenk an dein Workout heute: ' + PLAN[dow] + ' (max. 15 Minuten).\n' +
      'Aktuell: Platz ' + (i + 1) + ' von ' + rank.length + ' mit ' + me.pts + ' Punkten, Serie 🔥 ' + me.streak + '.\n' +
      (t < CH_START ? '🧪 Testphase – ab 1.12. zählt\'s richtig.\n' : '') +
      '\nJetzt starten: ' + APP_LINK + '\n\nZieh durch! 💪';
    MailApp.sendEmail({ to: o.email, subject: '☀️ Denk an dein Workout heute, ' + o.name + '!', body: body, name: 'Workout-Coach' });
  });
}

// Test: zeigt die Mail im Ausführungsprotokoll, ohne etwas zu senden
function testMailVorschau() {
  const t = today_();
  const rank = ranking_(rows_().rows, t);
  Logger.log(rank.map((r, i) => (i + 1) + '. ' + r.name + ' ' + r.pts + ' Pkt, Serie ' + r.streak).join('\n') || 'noch keine Benutzer');
}
