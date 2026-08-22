/**
 * BISC Union — form backend
 * Receives sign-ups and questions from the BISC Union web app,
 * writes them to a Google Sheet, and emails a notification.
 *
 * SETUP — about ten minutes, once.
 *
 *  1. Create a new Google Sheet. Name it something like "BISC Union — Intake".
 *  2. Copy its ID from the URL. In
 *       https://docs.google.com/spreadsheets/d/1AbC...XyZ/edit
 *     the ID is the 1AbC...XyZ part. Paste it into SHEET_ID below.
 *  3. Put the address that should get notified into NOTIFY_EMAIL.
 *  4. In the Sheet, go to Extensions → Apps Script. Delete whatever is
 *     in the editor and paste this whole file in. Save.
 *  5. Click Deploy → New deployment. Choose type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone
 *     Deploy, then approve the permissions prompt when it appears.
 *  6. Copy the Web app URL. It ends in /exec.
 *  7. In bisc-union-app.html, find the line   const ENDPOINT = '';
 *     and paste the URL between the quotes. Save and upload.
 *
 * IMPORTANT: after any edit to this file, you must click
 * Deploy → Manage deployments → edit → New version, or the live
 * URL will keep running the old code.
 */

const VERSION      = '2026-08-22-d · adds ?show diagnostics';
const SHEET_ID     = '1LmKjohQDfLLKVfEyjI1O4Yg0pccZvSef0-_91LoNH7I';
const NOTIFY_EMAIL = 'loren@dalbert.design';

const COLUMNS = {
  join: ['Timestamp', 'First name', 'Last name', 'Email', 'Phone',
         'Contact preference', 'Language', 'State', 'ZIP', 'Pilot',
         'Interested in', 'Requests', 'Anything else'],
  question: ['Timestamp', 'Name', 'Email', 'Question'],
  note: ['Timestamp', 'Taken at', 'Table', 'Taken by', 'Kind',
         'Name', 'Email', 'Who they are', 'Next step', 'Observation'],
  feedback: ['Timestamp', 'From', 'App', 'Page', 'Kind',
             'Highlighted text', 'Comment', 'Screen', 'Language', 'Status']
};

const TAB = { join: 'Members', question: 'Questions', note: 'Session Notes', feedback: 'Feedback' };


/* =========================================================
   AUTHENTICATION
   Members sign in with their email. We check it against the
   Members sheet, email them a six-digit code, and issue a
   token that lasts 30 days.

   The point of doing this server-side: gated content is
   never inside the HTML file. A person who views source
   sees the app shell and nothing else.
   ========================================================= */

const CODE_MINUTES  = 10;   // how long a code stays valid
const TOKEN_DAYS    = 30;   // how long someone stays signed in
const MAX_CODES_HOUR = 5;   // per email, stops abuse

function authRequest(data) {
  const email = String(data.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') < 0) return reply({ ok: false, error: 'bad email' });

  const props = PropertiesService.getScriptProperties();

  // simple rate limit
  const rlKey = 'rl:' + email;
  const rl = JSON.parse(props.getProperty(rlKey) || '[]')
    .filter(function (t) { return Date.now() - t < 3600000; });
  if (rl.length >= MAX_CODES_HOUR) return reply({ ok: true });  // silently stop
  rl.push(Date.now());
  props.setProperty(rlKey, JSON.stringify(rl));

  const member = findMember(email);
  const code = String(Math.floor(100000 + Math.random() * 900000));

  if (member) {
    props.setProperty('code:' + email, JSON.stringify({
      code: code, exp: Date.now() + CODE_MINUTES * 60000
    }));
    MailApp.sendEmail({
      to: email,
      subject: 'Your BISC Union sign-in code: ' + code,
      body: 'Hi ' + (member.first || 'there') + ',\n\n' +
            'Your sign-in code is ' + code + '\n\n' +
            'It works for the next ' + CODE_MINUTES + ' minutes. If you did not ask ' +
            'for it, you can ignore this email \u2014 nobody can get in without it.\n\n' +
            'BISC Union'
    });
  } else {
    // Same response either way, so the API never reveals who is a member.
    MailApp.sendEmail({
      to: email,
      subject: 'BISC Union \u2014 we could not find that email',
      body: 'Someone tried to sign in to the BISC Union app with this address, ' +
            'but we could not find it on our member list.\n\n' +
            'If that was you, it may be that you signed up with a different email, ' +
            'or have not joined yet. You can join any time from the app.\n\n' +
            'If it was not you, nothing has happened and you can ignore this.\n\n' +
            'BISC Union'
    });
  }
  return reply({ ok: true });
}

function findMember(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB.join);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COLUMNS.join.length).getValues();
  const emailCol = COLUMNS.join.indexOf('Email');
  const firstCol = COLUMNS.join.indexOf('First name');
  const lastCol  = COLUMNS.join.indexOf('Last name');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][emailCol]).trim().toLowerCase() === email) {
      return { first: rows[i][firstCol], last: rows[i][lastCol], email: email };
    }
  }
  return null;
}

function authVerify(data) {
  const email = String(data.email || '').trim().toLowerCase();
  const code  = String(data.code || '').trim();
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('code:' + email);
  if (!raw) return reply({ ok: false, error: 'no code' });

  const rec = JSON.parse(raw);
  if (Date.now() > rec.exp) { props.deleteProperty('code:' + email); return reply({ ok: false, error: 'expired' }); }
  if (rec.code !== code)    { return reply({ ok: false, error: 'wrong code' }); }

  props.deleteProperty('code:' + email);   // single use

  const member = findMember(email);
  if (!member) return reply({ ok: false, error: 'not a member' });

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().slice(0, 8);
  const expires = new Date(Date.now() + TOKEN_DAYS * 86400000);
  const sheet = getSheet('Auth Tokens',
    ['Token', 'Email', 'Name', 'Issued', 'Expires', 'Revoked']);
  sheet.appendRow([token, email, (member.first + ' ' + member.last).trim(), new Date(), expires, '']);

  return reply({ ok: true, token: token, name: member.first, expires: expires.toISOString() });
}

function tokenHolder(token) {
  if (!token) return null;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Auth Tokens');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === token) {
      if (rows[i][5]) return null;                       // revoked
      if (new Date(rows[i][4]) < new Date()) return null; // expired
      return { email: rows[i][1], name: rows[i][2] };
    }
  }
  return null;
}

function authCheck(data) {
  const who = tokenHolder(data.token);
  return who ? reply({ ok: true, name: String(who.name).split(' ')[0] })
             : reply({ ok: false });
}

/* =========================================================
   MEETING NOTES
   Lives in a "Meeting Notes" tab. Anyone with edit access to
   that sheet can post notes — no code, no developer.

   Columns, in order:
     A  Date            e.g. 2026-09-30
     B  Meeting         e.g. Full union meeting
     C  Attendance      e.g. 19 across both sessions
     D  Decisions       one per line
     E  Actions         one per line, each with a name and a date
     F  Open questions  one per line
     G  Next agenda     one per line
     H  Published       yes / no  (blank counts as no)
     I  Posted by       who wrote them up

   The four middle columns match the note-taking module in the
   curriculum on purpose: decisions, actions, open questions,
   next agenda. The sheet teaches the practice.
   ========================================================= */

const NOTES_TAB = 'Meeting Notes';
const NOTES_COLUMNS = ['Date', 'Meeting', 'Attendance', 'Decisions', 'Actions',
                       'Open questions', 'Next agenda', 'Published', 'Posted by'];

function lines(cell) {
  return String(cell || '').split(/\r?\n/).map(function (x) { return x.trim(); })
    .filter(function (x) { return x.length; });
}

function gatedFetch(data) {
  const who = tokenHolder(data.token);
  if (!who) return reply({ ok: false, error: 'not signed in' });

  const sheet = getSheet(NOTES_TAB, NOTES_COLUMNS);
  if (sheet.getLastRow() < 2) return reply({ ok: true, items: [], name: who.name });

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, NOTES_COLUMNS.length).getValues();
  const items = [];
  for (var i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const pub = String(r[7]).trim().toLowerCase();
    if (pub !== 'yes' && pub !== 'true' && pub !== 'y') continue;

    const d = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    items.push({
      iso: isNaN(d) ? '' : d.toISOString(),
      date: isNaN(d) ? String(r[0]) : Utilities.formatDate(d, 'America/New_York', 'MMMM d, yyyy'),
      meeting: String(r[1] || 'Meeting'),
      attendance: String(r[2] || ''),
      decisions: lines(r[3]),
      actions: lines(r[4]),
      questions: lines(r[5]),
      next: lines(r[6]),
      by: String(r[8] || '')
    });
  }
  items.sort(function (a, b) { return (b.iso || '').localeCompare(a.iso || ''); });
  return reply({ ok: true, items: items, name: who.name });
}

/* =========================================================
   MEETING AGENDAS
   Lives in an "Agendas" tab. Whoever holds the agenda role
   writes the next meeting's agenda by adding rows — no code.

   Columns, in order:
     A  Date          the meeting date, e.g. 2026-09-30
     B  Meeting       New member meeting  /  Full union meeting
     C  Order         1, 2, 3 … the order items run in
     D  Minutes       how long this item gets
     E  Item          short title
     F  Detail        one or two sentences
     G  Published     yes / no  (blank counts as no)
     H  Prepared by   who wrote it

   Clock times are worked out from the meeting's start time and
   the Minutes column, so changing a length or reordering rows
   reflows the whole agenda. Nobody types times.

   Posted agendas are members only. A signed-out visitor always
   sees the standard running order, whether or not one has been
   posted. No agenda posted at all? Everyone sees the standard
   running order, so a meeting is never left without one.
   ========================================================= */

/* Every type doPost knows how to route. Keep this and the routing
   in doPost in step — it is what tells you whether a deployment is
   actually current. */
const HANDLED_TYPES = ['join', 'question', 'note', 'feedback',
                       'auth_request', 'auth_verify', 'auth_check',
                       'gated_fetch', 'agenda_fetch', 'calendar_fetch'];

const AGENDA_TAB = 'Agendas';
const AGENDA_COLUMNS = ['Date', 'Meeting', 'Order', 'Minutes', 'Item', 'Detail',
                        'Published', 'Prepared by'];

function agendaFetch(data) {
  // members only — a signed-out visitor gets nothing and falls
  // back to the standard running order built into the app
  const who = tokenHolder(data && data.token);
  if (!who) return reply({ ok: false, error: 'not signed in' });

  const sheet = getSheet(AGENDA_TAB, AGENDA_COLUMNS);
  if (sheet.getLastRow() < 2) return reply({ ok: true, items: [] });

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, AGENDA_COLUMNS.length).getValues();
  const items = [];
  for (var i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || !r[4]) continue;
    const pub = String(r[6]).trim().toLowerCase();
    if (pub !== 'yes' && pub !== 'true' && pub !== 'y') continue;

    const d = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    if (isNaN(d)) continue;

    items.push({
      date: Utilities.formatDate(d, 'America/New_York', 'yyyy-MM-dd'),
      meeting: String(r[1] || '').trim(),
      order: Number(r[2]) || (i + 1),
      minutes: Number(r[3]) || 5,
      item: String(r[4]).trim(),
      detail: String(r[5] || '').trim(),
      by: String(r[7] || '').trim()
    });
  }
  items.sort(function (a, b) {
    return a.date === b.date ? a.order - b.order : a.date.localeCompare(b.date);
  });
  return reply({ ok: true, items: items });
}

/* =========================================================
   CALENDAR
   Two tabs, each doing one job.

   "Events" — one-off things: conferences, local actions,
   anything that is not part of the union's own cycle.
     A  Start date     2026-09-17
     B  End date       2026-09-18   (blank for a single day)
     C  Title          Basic Income Guarantee (BIG) Conference
     D  Where          Portland, Oregon
     E  Toolkit URL    big-2026-toolkit.html   (blank if none)
     F  Toolkit label  Open the BIG 2026 toolkit
     G  Members only   yes / no    (applies to the toolkit link)
     H  Published      yes / no

   "Meeting changes" — cancellations and moves to the union's
   own recurring meetings. The pattern stays in the app; this
   overrides a single date.
     A  Date        the date as it would normally fall
     B  Meeting     New member meeting / Full union meeting / Saturday mixer
     C  Change      cancelled  /  moved
     D  New date    only for a move
     E  Reason      shown to members, e.g. "Thanksgiving week"
     F  Published   yes / no

   Nobody has to add twelve rows a year. The recurring pattern
   is in the app; the sheet only handles exceptions.
   ========================================================= */

const EVENTS_TAB = 'Events';
const EVENTS_COLUMNS = ['Start date', 'End date', 'Title', 'Where',
                        'Toolkit URL', 'Toolkit label', 'Members only', 'Published'];
const CHANGES_TAB = 'Meeting changes';
const CHANGES_COLUMNS = ['Date', 'Meeting', 'Change', 'New date', 'Reason', 'Published'];

function yes(v) {
  const t = String(v).trim().toLowerCase();
  return t === 'yes' || t === 'true' || t === 'y';
}
function ymd(v) {
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d) ? '' : Utilities.formatDate(d, 'America/New_York', 'yyyy-MM-dd');
}

function calendarFetch(data) {
  const member = !!tokenHolder(data && data.token);

  const ev = getSheet(EVENTS_TAB, EVENTS_COLUMNS);
  const events = [];
  if (ev.getLastRow() > 1) {
    const rows = ev.getRange(2, 1, ev.getLastRow() - 1, EVENTS_COLUMNS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0] || !r[2] || !yes(r[7])) continue;
      const start = ymd(r[0]);
      if (!start) continue;
      const gated = yes(r[6]);
      events.push({
        start: start,
        end: ymd(r[1]) || start,
        title: String(r[2]).trim(),
        where: String(r[3] || '').trim(),
        // a gated toolkit link is withheld entirely from a signed-out visitor
        url: (r[4] && (!gated || member)) ? String(r[4]).trim() : '',
        label: String(r[5] || 'Open the toolkit').trim(),
        gated: gated,
        hasToolkit: !!r[4]
      });
    }
  }

  const ch = getSheet(CHANGES_TAB, CHANGES_COLUMNS);
  const changes = [];
  if (ch.getLastRow() > 1) {
    const rows = ch.getRange(2, 1, ch.getLastRow() - 1, CHANGES_COLUMNS.length).getValues();
    for (var j = 0; j < rows.length; j++) {
      const r = rows[j];
      if (!r[0] || !yes(r[5])) continue;
      const on = ymd(r[0]);
      if (!on) continue;
      changes.push({
        date: on,
        meeting: String(r[1] || '').trim(),
        change: String(r[2] || 'cancelled').trim().toLowerCase(),
        newDate: ymd(r[3]),
        reason: String(r[4] || '').trim()
      });
    }
  }

  return reply({ ok: true, events: events, changes: changes, member: member });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply({ ok: false, error: 'empty request' });
    }

    const data = JSON.parse(e.postData.contents);

    // Honeypot: real people never fill this in. Accept silently so
    // bots get a success and stop retrying, but write nothing.
    if (data.hp) return reply({ ok: true });

    if (data.type === 'auth_request') return authRequest(data);
    if (data.type === 'auth_verify')  return authVerify(data);
    if (data.type === 'auth_check')   return authCheck(data);
    if (data.type === 'gated_fetch')  return gatedFetch(data);
    if (data.type === 'agenda_fetch')   return agendaFetch(data);
    if (data.type === 'calendar_fetch') return calendarFetch(data);

    const type = data.type === 'join' ? 'join'
               : data.type === 'note' ? 'note'
               : data.type === 'feedback' ? 'feedback'
               : 'question';
    const sheet = getSheet(TAB[type], COLUMNS[type]);
    const now = new Date();

    if (type === 'feedback') {
      sheet.appendRow([now, data.from || '', data.app || '', data.page || '', data.kind || '',
        data.selected || '', data.comment || '', data.screen || '', data.lang || '', 'New']);
      return reply({ ok: true });
    }

    if (type === 'note') {
      sheet.appendRow([now, data.ts || '', data.table || '', data.by || '', data.kind || '',
        data.name || '', data.email || '', data.who || '', data.next || '', data.text || '']);
      return reply({ ok: true });
    }

    const row = type === 'join'
      ? [now, data.first || '', data.last || '', data.email || '', data.phone || '',
         data.contact || '', data.language || '', data.state || '', data.zip || '', data.pilot || '',
         data.interests || '', data.requests || '', data.note || '']
      : [now, data.name || '', data.email || '', data.message || ''];

    sheet.appendRow(row);
    notify(type, data);
    return reply({ ok: true });

  } catch (err) {
    console.error(err);
    return reply({ ok: false, error: String(err) });
  }
}

// Lets you open the /exec URL in a browser to confirm it is live.
function doGet(e) {
  // Self-test: open  <your /exec URL>?probe=feedback  in a browser.
  // It writes one test row, which proves the routing works without
  // involving the app at all. Also try probe=note or probe=question.
  // One-time setup: open  <your /exec URL>?setup=1  in a browser and
  // every tab is created with its headers. Safe to run more than once —
  // existing tabs are left exactly as they are.
  if (e && e.parameter && e.parameter.setup) {
    const made = [], already = [];
    const all = [
      [TAB.join, COLUMNS.join],
      [TAB.question, COLUMNS.question],
      [TAB.note, COLUMNS.note],
      [TAB.feedback, COLUMNS.feedback],
      [NOTES_TAB, NOTES_COLUMNS],
      [AGENDA_TAB, AGENDA_COLUMNS],
      [EVENTS_TAB, EVENTS_COLUMNS],
      [CHANGES_TAB, CHANGES_COLUMNS],
      ['Auth Tokens', ['Token', 'Email', 'Name', 'Issued', 'Expires', 'Revoked']]
    ];
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      for (var i = 0; i < all.length; i++) {
        const name = all[i][0];
        if (ss.getSheetByName(name)) { already.push(name); }
        else { getSheet(name, all[i][1]); made.push(name); }
      }
      return reply({ ok: true, created: made, alreadyThere: already,
                     sheetId: SHEET_ID, version: VERSION });
    } catch (err) {
      return reply({ ok: false, error: String(err), sheetId: SHEET_ID });
    }
  }

  // Read-only diagnostics: see exactly what the app will receive.
  //   ?show=calendar   what the Events and Meeting changes tabs parse to
  //   ?show=agendas    what the Agendas tab parses to (members only in the app,
  //                    shown here so you can debug without a token)
  //   ?show=notes      what the Meeting Notes tab parses to
  const show = e && e.parameter && e.parameter.show;
  if (show === 'calendar') return calendarFetch({});
  if (show === 'agendas') {
    const sheet = getSheet(AGENDA_TAB, AGENDA_COLUMNS);
    const n = Math.max(0, sheet.getLastRow() - 1);
    const rows = n ? sheet.getRange(2, 1, n, AGENDA_COLUMNS.length).getValues() : [];
    return reply({ ok: true, rowsInSheet: n, raw: rows });
  }
  if (show === 'notes') {
    const sheet = getSheet(NOTES_TAB, NOTES_COLUMNS);
    const n = Math.max(0, sheet.getLastRow() - 1);
    const rows = n ? sheet.getRange(2, 1, n, NOTES_COLUMNS.length).getValues() : [];
    return reply({ ok: true, rowsInSheet: n, raw: rows });
  }

  const probe = e && e.parameter && e.parameter.probe;
  if (probe) {
    try {
      const type = ['join', 'question', 'note', 'feedback'].indexOf(probe) > -1 ? probe : 'question';
      const sheet = getSheet(TAB[type], COLUMNS[type]);
      const now = new Date();
      const row = [now];
      for (var i = 1; i < COLUMNS[type].length; i++) row.push('PROBE TEST');
      sheet.appendRow(row);
      return reply({
        ok: true, probe: probe, wroteTo: TAB[type],
        columns: COLUMNS[type].length, version: VERSION,
        note: 'A test row was added to the ' + TAB[type] + ' tab. Delete it when you are done.'
      });
    } catch (err) {
      return reply({ ok: false, probe: probe, error: String(err), version: VERSION });
    }
  }

  return reply({
    ok: true,
    status: 'BISC Union intake is running',
    version: VERSION,
    setupRoute: true,
    diagnostics: ['?setup=1', '?show=calendar', '?show=agendas', '?show=notes', '?probe=feedback'],
    scriptId: ScriptApp.getScriptId(),
    sheetId: SHEET_ID,
    tabs: TAB,
    // derived from the routing below, so it can never go stale
    handles: HANDLED_TYPES
  });
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function notify(type, data) {
  if (!NOTIFY_EMAIL) return;
  try {
    if (type === 'join') {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New BISC member: ' + data.first + ' ' + data.last,
        body:
          data.first + ' ' + data.last + ' just joined.\n\n' +
          'Email: ' + data.email + '\n' +
          'Phone: ' + (data.phone || 'not given') + '\n' +
          'Reach by: ' + data.contact + '\n' +
          'Language: ' + (data.language || 'not given') + '\n' +
          'Where: ' + data.state + ' ' + data.zip + '\n' +
          'Pilot: ' + data.pilot + '\n' +
          'Interested in: ' + (data.interests || 'still deciding') + '\n' +
          'Asked for: ' + (data.requests || 'nothing specific') + '\n\n' +
          (data.note ? 'They wrote:\n' + data.note + '\n\n' : '') +
          (String(data.requests || '').indexOf('buddy') > -1
            ? '>> They asked for a BISC Buddy. Pair them within a week.\n'
            : '') +
          (data.language && data.language !== 'English' && data.language !== 'Either'
            ? '>> Language: ' + data.language + '. Pair with someone who speaks it.\n'
            : '')
      });
    } else {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        replyTo: data.email,
        subject: 'BISC question from ' + data.name,
        body: data.name + ' (' + data.email + ') asked:\n\n' + data.message +
              '\n\nReply straight to this email to answer them.'
      });
    }
  } catch (err) {
    // A failed notification must never lose the row.
    console.error('notify failed: ' + err);
  }
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
