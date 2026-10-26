/**
 * OSA Pup Pause — results receiver (Google Apps Script, bound to a private Google Sheet).
 * - doPost: receives one anonymous check-in from the app and appends it to a tab for that month.
 * - updateSummary: rebuilds the "Summary" tab, one row per month.
 * - updateTestReport: rebuilds the "Test report" tab (Test 1 vs Test 2, go-forward verdicts). Both refresh every hour.
 * The web app can only ADD rows. It never returns data, so the public app cannot read results.
 */
const TZ = 'America/New_York';
const HEADERS = ['Received','Check-in ID','Device','Link','Language','Mood','Feeling words','Path','Named a feeling',
  'Kinds of help','Staff shown','Changed suggestion','Saw first step','Tapped booking','Booked with','Nudge shown',
  'Acted after nudge','What they would do first','Knew first step','Would make it easier','Opened support services',
  'Support links tapped','Finished','Seconds to finish','Season','Sunset (24h)','After midterms','Finals','Break'];
const ALLOWED_LINKS = ['test1','test2','demo','preview'];
const HELP = {wellness:'Stress & wellbeing',academic:'Classes & deadlines',access:'Access needs',intl:'Far from home',belong:'Belonging',housing:'Housing & groups',money:'Money & basics',safety:'Safety & reporting',systems:'Harvard systems'};
const STAFF = {rilda:'Rilda Kissel',coach:'Academic coaches',brooke:'Brooke Huynh',andrea:'Andrea Le',taylor:'Taylor Goyette-Frechette',crystal:'Crystal Quintanilla',kevin:'Kevin Boehm'};
const LINKNAME = {test1:'Test 1',test2:'Test 2',demo:'Everyday link'};
function map_(a, m) { return (Array.isArray(a) ? a : []).map(k => m[k] || clean_(k)).join(', '); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > 8000) return ok_();
    const msg = JSON.parse(raw);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (msg.type === 'followup') {
      const email = String(msg.email || '').slice(0, 120);
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return ok_();
      const sh = ss.getSheetByName('Check on me') || ss.insertSheet('Check on me');
      if (sh.getLastRow() === 0) sh.appendRow(['Received', 'Email', 'Link', 'Note sent?']);
      sh.appendRow([new Date(), email, clean_(msg.test), '']);
      return ok_();
    }
    if (msg.type !== 'checkin' || !msg.rec) return ok_();
    const r = msg.rec;
    if (r.test === 'preview') return ok_();            // staff previews are not real check-ins
    const now = new Date();
    const tab = Utilities.formatDate(now, TZ, 'yyyy-MM MMMM'); // e.g. "2026-10 October"
    const sh = ss.getSheetByName(tab) || newMonthTab_(ss, tab);
    const c = r.ctx || {};
    const row = [now, clean_(r.sid), clean_(r.device), LINKNAME[ALLOWED_LINKS.indexOf(r.test) > -1 ? r.test : 'demo'], clean_(r.lang),
      clean_(r.mood), list_(r.words), r.fine ? 'Thank-you (doing okay)' : (r.sawStep || (r.areas && r.areas.length) ? 'First step' : 'Started'),
      yn_(r.named), map_(r.areas, HELP), map_(r.staffList, STAFF), yn_(r.areaChanged), yn_(r.sawStep), yn_(r.calendly || r.pickedSlot),
      map_(Object.keys(r.calWho || {}), STAFF), yn_(r.nudged), yn_(r.pickedAfterNudge), clean_(r.check), yn_(r.knew), list_(r.friction),
      yn_(r.resOpened), list_(Object.keys(r.resClicks || {})), yn_(r.done), r.ms && r.ms.done ? Math.round(r.ms.done / 1000) : '',
      clean_(c.season), c.sunset || '', yn_(c.postMidterm), yn_(c.finals), yn_(c.brk)];
    // one row per check-in: if this ID already exists this month (an early partial send), replace it
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(v => v[0]) : [];
    const at = ids.indexOf(row[1]);
    if (at > -1) sh.getRange(at + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
  return ok_();
}

function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const enroll = settings ? Number(settings.getRange('B1').getValue()) || 0 : 0;
  const goal = settings ? Number(settings.getRange('B2').getValue()) || 100 : 100;
  const months = ss.getSheets().filter(s => /^\d{4}-\d{2} /.test(s.getName())).sort((a, b) => a.getName() < b.getName() ? -1 : 1);
  const out = [['Month','Check-ins','Students (devices)','% of HGSE','Goal %','Finished','Doing okay (thank-you)','Got a first step',
    'Named a feeling %','Knew first step %','Took a help action %','Acted after nudge %','Top kind of help','Test 1 check-ins','Test 2 check-ins']];
  months.forEach(sh => {
    const n = sh.getLastRow() - 1;
    if (n < 1) { out.push([sh.getName(), 0, 0, '', goal, 0, 0, 0, '', '', '', '', '', 0, 0]); return; }
    const v = sh.getRange(2, 1, n, HEADERS.length).getValues();
    const col = name => HEADERS.indexOf(name);
    const cnt = f => v.filter(f).length;
    const devices = new Set(v.map(x => x[col('Device')]).filter(Boolean)).size;
    const firstStep = v.filter(x => x[col('Path')] === 'First step');
    const finished = v.filter(x => x[col('Finished')] === 'Yes');
    const nudged = v.filter(x => x[col('Nudge shown')] === 'Yes');
    const help = {};
    firstStep.forEach(x => String(x[col('Kinds of help')]).split(', ').filter(Boolean).forEach(k => help[k] = (help[k] || 0) + 1));
    const top = Object.keys(help).sort((a, b) => help[b] - help[a])[0] || '';
    const pct = (a, b) => b ? Math.round(100 * a / b) : '';
    out.push([sh.getName(), n, devices, enroll ? Math.min(100, Math.round(1000 * devices / enroll) / 10) : 'add enrollment in Settings', goal,
      finished.length, cnt(x => x[col('Path')] === 'Thank-you (doing okay)'), firstStep.length,
      pct(cnt(x => x[col('Named a feeling')] === 'Yes'), n), pct(finished.filter(x => x[col('Knew first step')] === 'Yes').length, finished.filter(x => x[col('What they would do first')]).length),
      pct(firstStep.filter(x => x[col('Tapped booking')] === 'Yes' || x[col('Opened support services')] === 'Yes').length, firstStep.length),
      pct(nudged.filter(x => x[col('Acted after nudge')] === 'Yes').length, nudged.length), top,
      cnt(x => x[col('Link')] === 'Test 1'), cnt(x => x[col('Link')] === 'Test 2')]);
  });
  const sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary', 0);
  sum.clear();
  sum.getRange(1, 1, out.length, out[0].length).setValues(out);
  sum.getRange(1, 1, 1, out[0].length).setFontWeight('bold').setBackground('#A51C30').setFontColor('#FFFFFF');
  sum.setFrozenRows(1);
  sum.getRange(out.length + 2, 1).setValue('Updated ' + Utilities.formatDate(new Date(), TZ, 'MMM d, yyyy h:mm a') + ' · newsletter line: "X% of HGSE students have taken a Pup Pause."');
}

/** Run once after pasting: creates Settings, Summary and Test report tabs, plus an hourly refresh. Safe to run again. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  SETTINGS_ROWS.forEach((r, i) => {
    const row = i + 1;
    if (st.getRange(row, 1).getValue() !== r[0]) st.getRange(row, 1, 1, r.length).setValues([r]);
  });
  st.getRange(1, 1, SETTINGS_ROWS.length, 1).setFontWeight('bold');
  st.getRange('A11:C11').setFontWeight('bold').setBackground('#1F1A1B').setFontColor('#FFFFFF');
  st.setColumnWidth(1, 300);
  ScriptApp.getProjectTriggers()
    .filter(t => ['updateSummary', 'refreshReports'].indexOf(t.getHandlerFunction()) > -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshReports').timeBased().everyHours(1).create();
  refreshReports();
}

/** Adds a "Pup Pause" menu so staff can refresh the reports right after a test. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pup Pause').addItem('Refresh reports now', 'refreshReports').addToUi();
}

function refreshReports() {
  updateSummary();
  updateTestReport();
}

/* ---------------- MVP test report ----------------
 * Rebuilds the "Test report" tab from every month tab: Test 1 vs Test 2 (and the everyday link),
 * the go-forward verdicts from the evidence tracker, and a suggested call.
 * Totals and percentages only: no emails, IDs or individual rows, so it is safe to share with OSA staff. */

// Settings tab layout. Column A labels are how the script finds each value, so keep them as they are.
const SETTINGS_ROWS = [
  ['HGSE students enrolled', ''],
  ['Participation goal (%)', 100],
  ['', ''],
  ['Test 1 date', ''],
  ['Test 1 students in the room (your headcount)', ''],
  ['Test 2 date', ''],
  ['Test 2 students in the room (your headcount)', ''],
  ['', ''],
  ['Go-forward bars: set these before Test 1 and leave them alone after.', ''],
  ['', ''],
  ['Construct level', 'Go forward if at least', 'Rethink if below'],
  ['L3 · Self-check (started ÷ in the room)', 0.5, 0.3],
  ['L4 · Naming (named a feeling ÷ started)', 0.6, 0.4],
  ['L5 · Knows first step (picked specialist ÷ answered)', 0.6, 0.4],
  ['L5 · Acts (students who booked or opened a support link)', 3, 1],
  ['Nudge (acted after the nudge ÷ nudge shown)', 0.25, 0.1]
];
const FIRST = {specialist: 'Talk with a specialist', therapy: 'Book therapy directly', unsure: 'Not sure'};
const FRICTION = {process: 'Knowing the steps', information: 'Knowing what OSA offers', relational: 'Knowing someone there',
  temporal: 'A shorter wait', spatial: 'Easier to find', peer: 'A friend coming with me'};
const REPORT_LINKS = ['Test 1', 'Test 2', 'Everyday link'];

function updateTestReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName('Settings');
  const settings = {};
  if (st && st.getLastRow() > 0) {
    st.getRange(1, 1, st.getLastRow(), 3).getValues().forEach(r => { if (r[0]) settings[String(r[0])] = [r[1], r[2]]; });
  }
  const rows = [];
  ss.getSheets().filter(s => /^\d{4}-\d{2} /.test(s.getName())).forEach(sh => {
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 1, n, HEADERS.length).getValues().forEach(r => rows.push(r));
  });
  const stamp = Utilities.formatDate(new Date(), TZ, 'MMM d, yyyy h:mm a');
  const out = buildTestReport_(rows, settings, stamp);
  const width = 5;
  const grid = out.map(r => { const x = r.cells.slice(0, width); while (x.length < width) x.push(''); return x; });
  const sh = ss.getSheetByName('Test report') || ss.insertSheet('Test report', 1);
  sh.clear();
  sh.getRange(1, 1, grid.length, width).setValues(grid).setWrap(true).setVerticalAlignment('top');
  out.forEach((r, i) => {
    const rg = sh.getRange(i + 1, 1, 1, width);
    if (r.style === 'title') rg.setFontWeight('bold').setFontSize(14).setFontColor('#A51C30');
    if (r.style === 'section') rg.setFontWeight('bold').setBackground('#A51C30').setFontColor('#FFFFFF');
    if (r.style === 'head') rg.setFontWeight('bold').setBackground('#F3E8EA');
    if (r.style === 'note') rg.setFontStyle('italic').setFontColor('#555555');
  });
  sh.setColumnWidth(1, 330);
  [2, 3, 4].forEach(c => sh.setColumnWidth(c, 190));
  sh.setColumnWidth(5, 300);
}

/** Pure function (no Sheets calls) so it can be tested. rows = month-tab rows in HEADERS order. */
function buildTestReport_(rows, settings, stamp) {
  const col = name => HEADERS.indexOf(name);
  const get = (label, i) => (settings[label] || ['', ''])[i || 0];
  const out = [];
  const add = (cells, style) => out.push({cells: cells, style: style || ''});
  const by = {};
  REPORT_LINKS.forEach(l => by[l] = rows.filter(r => r[col('Link')] === l));
  const yes = (r, name) => r[col(name)] === 'Yes';
  const fmtDate = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'MMM d, yyyy') : String(v || '');
  const pct = (a, b) => b ? Math.round(100 * a / b) + '% (' + a + ' of ' + b + ')' : 'No data yet';

  const m = {};
  REPORT_LINKS.forEach(l => {
    const v = by[l];
    const answered = v.filter(r => String(r[col('What they would do first')]) !== '');
    const secs = v.map(r => Number(r[col('Seconds to finish')])).filter(x => x > 0).sort((a, b) => a - b);
    m[l] = {
      room: l === 'Test 1' ? Number(get('Test 1 students in the room (your headcount)')) || 0
          : l === 'Test 2' ? Number(get('Test 2 students in the room (your headcount)')) || 0 : 0,
      date: l === 'Test 1' ? fmtDate(get('Test 1 date')) : l === 'Test 2' ? fmtDate(get('Test 2 date')) : '',
      started: v.length,
      finished: v.filter(r => yes(r, 'Finished')).length,
      named: v.filter(r => yes(r, 'Named a feeling')).length,
      answered: answered.length,
      specialist: answered.filter(r => r[col('What they would do first')] === 'specialist').length,
      acted: v.filter(r => yes(r, 'Tapped booking') || yes(r, 'Opened support services')).length,
      nudged: v.filter(r => yes(r, 'Nudge shown')).length,
      nudgeActed: v.filter(r => yes(r, 'Nudge shown') && yes(r, 'Acted after nudge')).length,
      okay: v.filter(r => r[col('Path')] === 'Thank-you (doing okay)').length,
      median: secs.length ? secs[Math.floor(secs.length / 2)] : '',
      tally: (name, labels) => {
        const t = {};
        v.forEach(r => String(r[col(name)]).split(',').map(s => s.trim()).filter(Boolean)
          .forEach(k => { const lab = labels ? (labels[k] || k) : k; t[lab] = (t[lab] || 0) + 1; }));
        return t;
      }
    };
  });

  add(['OSA Pup Pause · MVP test report'], 'title');
  add(['Updated ' + stamp + '. Rebuilt every hour, or use the Pup Pause menu → Refresh reports now.'], 'note');
  add(['Totals and percentages only. No names, emails or individual answers appear on this tab.'], 'note');
  add(['']);

  // 1. Counts
  add(['1 · What students did', 'Test 1', 'Test 2', 'Everyday link', 'Where it comes from'], 'section');
  const line = (label, key, where, onlyTests) =>
    add([label].concat(REPORT_LINKS.map(l => onlyTests && l === 'Everyday link' ? '' : (m[l][key] === 0 && key === 'room' ? 'Add in Settings' : m[l][key]))).concat([where]));
  line('Date of test', 'date', 'Settings tab', true);
  line('Students in the room (headcount)', 'room', 'Settings tab: your count at the start of the last five minutes', true);
  line('Started a check-in', 'started', 'Check-ins sent from that link');
  line('Finished all four steps', 'finished', '"Finished" = Yes');
  line('Named a specific feeling', 'named', '"Named a feeling" = Yes');
  line('Answered "What would you do first?"', 'answered', '"What they would do first" not blank');
  line('Picked "Talk with a specialist"', 'specialist', '"What they would do first" = specialist');
  line('Tapped booking or opened a support link', 'acted', '"Tapped booking" or "Opened support services" = Yes');
  line('Nudge shown', 'nudged', '"Nudge shown" = Yes');
  line('Acted after the nudge', 'nudgeActed', '"Acted after nudge" = Yes');
  line('Doing okay (thank-you path)', 'okay', '"Path" = Thank-you');
  line('Typical seconds to finish (median)', 'median', '"Seconds to finish"');
  add(['']);

  // 2. Verdicts against the bars
  const bar = label => { const b = settings[label] || ['', '']; return {go: Number(b[0]), re: Number(b[1]), set: b[0] !== '' && b[1] !== ''}; };
  const levels = [
    {name: 'L3 · Self-check', label: SETTINGS_ROWS[11][0], num: x => x.started, den: x => x.room},
    {name: 'L4 · Naming', label: SETTINGS_ROWS[12][0], num: x => x.named, den: x => x.started},
    {name: 'L5 · Knows first step', label: SETTINGS_ROWS[13][0], num: x => x.specialist, den: x => x.answered},
    {name: 'L5 · Acts', label: SETTINGS_ROWS[14][0], count: x => x.acted},
    {name: 'Nudge', label: SETTINGS_ROWS[15][0], num: x => x.nudgeActed, den: x => x.nudged}
  ];
  const verdicts = {'Test 1': [], 'Test 2': []};
  add(['2 · Go-forward verdicts', 'Test 1', 'Test 2', 'Bar (go forward / rethink)', 'How it is calculated'], 'section');
  levels.forEach(L => {
    const b = bar(L.label);
    const cells = [L.name];
    ['Test 1', 'Test 2'].forEach(t => {
      const x = m[t];
      let value, shown, v;
      if (L.count) { value = L.count(x); shown = value + (value === 1 ? ' student' : ' students'); if (!x.started) value = null; }
      else { const d = L.den(x); value = d ? L.num(x) / d : null; shown = d ? pct(L.num(x), d) : (t !== 'Everyday link' && L.name === 'L3 · Self-check' && !x.room ? 'Add headcount in Settings' : 'No data yet'); }
      if (value === null || !b.set) v = '';
      else v = value >= b.go ? 'Go forward' : value < b.re ? 'Rethink' : 'Watch';
      verdicts[t].push(v);
      cells.push(v ? v + ': ' + shown : shown);
    });
    const f = n => L.count ? String(n) : Math.round(n * 100) + '%';
    cells.push(b.set ? f(b.go) + ' / ' + f(b.re) : 'Set in Settings');
    cells.push(L.label.replace(/^[^(]*\(|\)$/g, ''));
    add(cells);
  });
  add(['"Watch" means between the two bars: not a clear signal yet.'], 'note');
  add(['']);

  // 3. Suggested call (same rules as the evidence tracker's Decision tab)
  add(['3 · Suggested call', 'Test 1', 'Test 2'], 'section');
  const call = vs => {
    const go = vs.filter(v => v === 'Go forward').length, re = vs.filter(v => v === 'Rethink').length, w = vs.filter(v => v === 'Watch').length;
    if (!(go + re + w)) return 'Waiting for results';
    if (re >= 2) return 'Pivot: two or more levels below the bar. Revisit the first step or the format.';
    if (re === 1) return 'Adjust: one level is below the bar. Change that part and re-test.';
    if (go >= 4) return 'Persevere: run a bigger test with more classes.';
    return 'Mixed: run the next test before deciding.';
  };
  ['Go forward', 'Watch', 'Rethink'].forEach(k => add([k + ' (count)', verdicts['Test 1'].filter(v => v === k).length, verdicts['Test 2'].filter(v => v === k).length]));
  add(['Suggested call', call(verdicts['Test 1']), call(verdicts['Test 2'])], 'head');
  add(['The suggestion is a starting point. Write the final decision in the evidence tracker.'], 'note');
  add(['']);

  // 4-6. Breakdowns
  const breakdown = (title, name, labels, note) => {
    add([title, 'Test 1', 'Test 2', 'Everyday link'], 'section');
    const t = {}; REPORT_LINKS.forEach(l => t[l] = m[l].tally(name, labels));
    const keys = labels ? Object.keys(labels).map(k => labels[k]) : [];
    REPORT_LINKS.forEach(l => Object.keys(t[l]).forEach(k => { if (keys.indexOf(k) < 0) keys.push(k); }));
    if (!keys.length) add(['No answers yet']);
    keys.forEach(k => add([k].concat(REPORT_LINKS.map(l => t[l][k] || 0))));
    const top = l => { const o = t[l]; const ks = Object.keys(o).sort((a, b) => o[b] - o[a]); return ks.length ? ks[0] + (ks.length > 1 && o[ks[0]] === o[ks[1]] ? ' (tied)' : '') : ''; };
    add(['Most common'].concat(REPORT_LINKS.map(top)), 'head');
    if (note) add([note], 'note');
    add(['']);
  };
  breakdown('4 · What would make reaching out easier', 'Would make it easier', FRICTION,
    'Friction has no bar. The most common answer sets the next MVP; an even spread means no single fix stands out.');
  breakdown('5 · "If you wanted support, what would you do first?"', 'What they would do first', FIRST, '');
  breakdown('6 · Kinds of help students pointed to', 'Kinds of help', null, '');

  add(['Limits to keep in mind'], 'section');
  ['One class is a small sample: treat results as early signals, not proof.',
   'Howie may get a novelty boost the first time students see him.',
   'Most answers are self-reported; booking taps and finishing are behavior.',
   'Test 1 and Test 2 may reach different students in different weeks.'].forEach(s => add(['• ' + s]));
  return out;
}

function newMonthTab_(ss, name) {
  const sh = ss.insertSheet(name);
  sh.appendRow(HEADERS);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1F1A1B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}
function clean_(v) { return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, 200); }
function list_(a) { return Array.isArray(a) ? a.map(clean_).join(', ') : clean_(a); }
function yn_(b) { return b ? 'Yes' : ''; }
function ok_() { return ContentService.createTextOutput('ok'); }
