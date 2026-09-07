/**
 * Job Tracker — pulls product roles straight from ATS APIs into this sheet.
 *
 * Greenhouse, Lever, Ashby and Workday each serve their public job boards
 * from a JSON API. Querying those directly means every row is live at fetch
 * time, so a closed req stops coming back instead of lingering the way an
 * indexed search result does.
 *
 * Setup:  Job Tracker menu -> Refresh now, then -> Install daily trigger.
 *
 * The sheet is the state. Status / Priority / Applied On / Notes are read
 * back before each refresh and written straight back, so edits survive.
 */

// ----------------------------------------------------------------- config

const CFG = {
  titleInclude: [
    'product manager', 'product management', 'product owner', 'product lead',
    'manager, product', 'manager - product', 'director, product',
    'director of product', 'head of product', 'vp of product', 'vp, product',
    'product strategy'
  ],
  titleExclude: [
    'intern', 'internship', 'co-op', 'apprentice', 'trainee',
    'program manager', 'project manager', 'technical program', 'scrum master',
    'product marketing', 'product support', 'product design', 'product designer',
    'recruiter', 'analyst', 'data scientist',
    'engineering manager', 'design manager'
  ],
  // Remove 'united states' to tighten to genuinely remote + Twin Cities.
  locationsAny: [
    'remote', 'anywhere', 'distributed', 'virtual', 'telecommute',
    'work at home', 'work from home', 'home office', 'nationwide',
    'multiple locations', 'minneapolis', 'st. paul', 'saint paul',
    'minnesota', ', mn', 'mn -', 'twin cities', 'united states', 'usa'
  ],
  maxAgeDays: 45,
  workdayMaxAgeDays: 30,
  workdayCategoryKeywords: ['product'],
  workdayRecencyPages: 10,   // 200 newest per tenant, on top of the category sweep

  greenhouse: {
    airtable: 'Airtable', anthropic: 'Anthropic', brex: 'Brex', carta: 'Carta',
    cloverhealth: 'Clover Health', databricks: 'Databricks', doximity: 'Doximity',
    elationhealth: 'Elation Health', figma: 'Figma', flatironhealth: 'Flatiron Health',
    garnerhealth: 'Garner Health', gusto: 'Gusto', healthjoy: 'HealthJoy',
    healthverity: 'HealthVerity', jamf: 'Jamf', komodohealth: 'Komodo Health',
    mavenclinic: 'Maven Clinic', mercury: 'Mercury', modernhealth: 'Modern Health',
    natera: 'Natera', omadahealth: 'Omada Health', papa: 'Papa', polaris: 'Polaris',
    sezzle: 'Sezzle', stripe: 'Stripe', suki: 'Suki AI', talkspace: 'Talkspace',
    tia: 'Tia', truveta: 'Truveta', zocdoc: 'Zocdoc'
  },
  lever: {
    aledade: 'Aledade', arcadia: 'Arcadia', canvasmedical: 'Canvas Medical',
    includedhealth: 'Included Health', palantir: 'Palantir', redoxengine: 'Redox',
    ro: 'Ro', swordhealth: 'Sword Health'
  },
  ashby: {
    abridge: 'Abridge', ambiencehealthcare: 'Ambience Healthcare',
    candidhealth: 'Candid Health', cedar: 'Cedar', commure: 'Commure',
    gravie: 'Gravie', hazel: 'Hazel Health', headway: 'Headway', nabla: 'Nabla',
    notion: 'Notion', openai: 'OpenAI', plaid: 'Plaid', ramp: 'Ramp',
    sondermind: 'SonderMind', trustedhealth: 'Trusted Health', virtahealth: 'Virta Health'
  },
  workday: {
    cvshealth:      { name: 'CVS Health',        dc: 1, site: 'CVS_Health_Careers' },
    target:         { name: 'Target',            dc: 5, site: 'targetcareers' },
    medtronic:      { name: 'Medtronic',         dc: 1, site: 'medtroniccareers' },
    cigna:          { name: 'Cigna / Evernorth', dc: 5, site: 'cignacareers' },
    humana:         { name: 'Humana',            dc: 5, site: 'Humana_External_Career_Site' },
    centene:        { name: 'Centene',           dc: 5, site: 'CENTENE_External' },
    hcsc:           { name: 'HCSC',              dc: 1, site: 'HCSC_External' },
    usbank:         { name: 'U.S. Bank',         dc: 1, site: 'US_Bank_Careers' },
    ecolab:         { name: 'Ecolab',            dc: 1, site: 'ECOLAB_External' },
    thomsonreuters: { name: 'Thomson Reuters',   dc: 5, site: 'External_Career_Site' }
  },

  googleQueries: [
    'site:myworkdayjobs.com "product manager" (remote OR minneapolis)',
    'site:job-boards.greenhouse.io "product manager"',
    'site:boards.greenhouse.io "product manager"',
    'site:jobs.lever.co "product manager"',
    'site:jobs.ashbyhq.com "product manager"',
    '(site:myworkdayjobs.com OR site:job-boards.greenhouse.io OR site:jobs.lever.co OR site:jobs.ashbyhq.com) "senior product manager" healthcare -intern'
  ]
};

const TRACK = ['Status', 'Priority', 'Applied On', 'Notes'];
const STATUSES = ['Interested', 'Applied', 'Screen', 'Interview', 'Onsite',
                  'Offer', 'Rejected', 'Passed', 'Closed'];
const PRIORITIES = ['A', 'B', 'C'];
const NAVY = '#1f3550', CREAM = '#f7f4ee', NEW_HL = '#fff3c4';

// ----------------------------------------------------------------- menu

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Job Tracker')
    .addItem('Refresh now', 'refresh')
    .addSeparator()
    .addItem('Install daily trigger (7am)', 'installTrigger')
    .addItem('Remove daily trigger', 'removeTrigger')
    .addToUi();
}

function installTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('refresh').timeBased().atHour(7).everyDays(1).create();
  SpreadsheetApp.getActive().toast('Daily refresh installed for ~7am.', 'Job Tracker', 5);
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'refresh'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

// ----------------------------------------------------------------- helpers

function today_() { var d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function ymd_(d) { return d ? Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd') : ''; }
function daysBetween_(a, b) { return Math.round((a - b) / 86400000); }

function parseIso_(s) {
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** "Posted 3 Days Ago" / "Posted Today" / "Posted 30+ Days Ago" */
function parseWorkdayDate_(text) {
  if (!text) return null;
  var t = String(text).toLowerCase(), now = today_(), m;
  if (t.indexOf('today') >= 0) return now;
  if (t.indexOf('yesterday') >= 0) return new Date(now.getTime() - 86400000);
  if ((m = t.match(/(\d+)\+?\s*day/)))   return new Date(now.getTime() - m[1] * 86400000);
  if ((m = t.match(/(\d+)\+?\s*month/))) return new Date(now.getTime() - m[1] * 30 * 86400000);
  return null;
}

function fetchAllJson_(requests) {
  if (!requests.length) return [];
  var out = [], CHUNK = 40;
  for (var i = 0; i < requests.length; i += CHUNK) {
    var slice = requests.slice(i, i + CHUNK);
    var responses;
    try {
      responses = UrlFetchApp.fetchAll(slice.map(function (r) { return r.req; }));
    } catch (e) {
      slice.forEach(function (r) { out.push({ meta: r.meta, data: null, err: String(e).slice(0, 120) }); });
      continue;
    }
    responses.forEach(function (resp, j) {
      var meta = slice[j].meta;
      try {
        if (resp.getResponseCode() >= 300) {
          out.push({ meta: meta, data: null, err: 'HTTP ' + resp.getResponseCode() });
        } else {
          out.push({ meta: meta, data: JSON.parse(resp.getContentText()), err: null });
        }
      } catch (e2) {
        out.push({ meta: meta, data: null, err: String(e2).slice(0, 120) });
      }
    });
  }
  return out;
}

function getReq_(url) {
  return { url: url, method: 'get', muteHttpExceptions: true,
           headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'application/json' } };
}

function wdReq_(tok, dc, site, offset, facets) {
  return {
    url: 'https://' + tok + '.wd' + dc + '.myworkdayjobs.com/wday/cxs/' + tok + '/' + site + '/jobs',
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'application/json' },
    payload: JSON.stringify({ appliedFacets: facets || {}, limit: 20,
                              offset: offset, searchText: '' })
  };
}

// ----------------------------------------------------------------- matching

function matches_(p) {
  var title = (p.title || '').toLowerCase();
  if (!title) return false;
  var inc = CFG.titleInclude.some(function (t) { return title.indexOf(t) >= 0; });
  if (!inc) return false;
  var exc = CFG.titleExclude.some(function (t) {
    return new RegExp('\\b' + t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(title);
  });
  if (exc) return false;

  if (CFG.locationsAny.length) {
    var loc = (p.location || '').toLowerCase().trim();
    // Workday collapses multi-site reqs to "5 Locations" and never says which,
    // so treat that as unknown rather than dropping a possible match.
    var unknown = !loc || /^\d+\s+locations?$/.test(loc);
    if (!unknown && !CFG.locationsAny.some(function (l) { return loc.indexOf(l) >= 0; })) return false;
  }
  if (p.posted && daysBetween_(today_(), p.posted) > CFG.maxAgeDays) return false;
  return true;
}

// ----------------------------------------------------------------- fetchers

function collectSimpleBoards_() {
  var reqs = [];
  Object.keys(CFG.greenhouse).forEach(function (tok) {
    reqs.push({ meta: { ats: 'Greenhouse', tok: tok, name: CFG.greenhouse[tok] },
      req: getReq_('https://boards-api.greenhouse.io/v1/boards/' + tok + '/jobs?content=false') });
  });
  Object.keys(CFG.lever).forEach(function (tok) {
    reqs.push({ meta: { ats: 'Lever', tok: tok, name: CFG.lever[tok] },
      req: getReq_('https://api.lever.co/v0/postings/' + tok + '?mode=json') });
  });
  Object.keys(CFG.ashby).forEach(function (tok) {
    reqs.push({ meta: { ats: 'Ashby', tok: tok, name: CFG.ashby[tok] },
      req: getReq_('https://api.ashbyhq.com/posting-api/job-board/' + tok) });
  });

  var postings = [], boardRows = [], rawUids = {};
  fetchAllJson_(reqs).forEach(function (r) {
    var m = r.meta, list = [];
    if (r.data) {
      if (m.ats === 'Greenhouse') list = normGreenhouse_(r.data, m);
      else if (m.ats === 'Lever') list = normLever_(r.data, m);
      else list = normAshby_(r.data, m);
    }
    list.forEach(function (p) { rawUids[p.uid] = 1; });
    var kept = list.filter(matches_);
    postings = postings.concat(kept);
    boardRows.push([m.name, m.ats, m.tok, list.length, kept.length, r.err || 'ok']);
  });
  return { postings: postings, boardRows: boardRows, rawUids: rawUids };
}

function normGreenhouse_(d, m) {
  return (d.jobs || []).map(function (j) {
    return {
      uid: 'greenhouse:' + m.tok + ':' + j.id,
      source: 'Greenhouse', company: m.name, title: (j.title || '').trim(),
      location: ((j.location || {}).name || '').trim(),
      url: j.absolute_url || '',
      posted: parseIso_(j.first_published) || parseIso_(j.updated_at),
      dept: '', req: j.requisition_id || ''
    };
  });
}

function normLever_(d, m) {
  if (!Array.isArray(d)) return [];
  return d.map(function (j) {
    var c = j.categories || {};
    var locs = c.allLocations || (c.location ? [c.location] : []);
    return {
      uid: 'lever:' + m.tok + ':' + j.id,
      source: 'Lever', company: m.name, title: (j.text || '').trim(),
      location: locs.join(' / '), url: j.hostedUrl || '',
      posted: j.createdAt ? new Date(Number(j.createdAt)) : null,
      dept: c.team || '', req: ''
    };
  });
}

function normAshby_(d, m) {
  return (d.jobs || []).filter(function (j) { return j.isListed !== false; }).map(function (j) {
    var locs = [j.location || ''].concat((j.secondaryLocations || []).map(function (s) { return s.location || ''; }))
      .filter(String);
    if (j.isRemote && !locs.some(function (x) { return x.toLowerCase().indexOf('remote') >= 0; })) locs.push('Remote');
    return {
      uid: 'ashby:' + m.tok + ':' + j.id,
      source: 'Ashby', company: m.name, title: (j.title || '').trim(),
      location: locs.join(' / '), url: j.jobUrl || j.applyUrl || '',
      posted: parseIso_(j.publishedAt), dept: j.department || j.team || '', req: ''
    };
  });
}

function normWorkdayJob_(j, tok, name, dc, site) {
  var path = j.externalPath || '';
  return {
    uid: 'workday:' + tok + ':' + path,
    source: 'Workday', company: name,
    title: (j.title || '').trim(),
    location: (j.locationsText || '').trim(),
    url: 'https://' + tok + '.wd' + dc + '.myworkdayjobs.com/' + site + path,
    posted: parseWorkdayDate_(j.postedOn),
    dept: j.timeType || '', req: (j.bulletFields || []).join(', ')
  };
}

/**
 * Workday in three parallel waves, to stay inside the Apps Script time limit.
 * Sequential paging would not finish; fetchAll does the same work in batches.
 */
function collectWorkday_() {
  var toks = Object.keys(CFG.workday);
  var seen = {}, boardCount = {}, boardErr = {}, rawCount = {};
  var cutoff = new Date(today_().getTime() - CFG.workdayMaxAgeDays * 86400000);

  function add(j, tok) {
    var c = CFG.workday[tok];
    var p = normWorkdayJob_(j, tok, c.name, c.dc, c.site);
    rawCount[tok] = (rawCount[tok] || 0) + 1;
    if (!seen[p.uid]) seen[p.uid] = p;
  }

  // wave 1 — first page per tenant, which also carries the facet list
  var w1 = toks.map(function (tok) {
    var c = CFG.workday[tok];
    return { meta: { tok: tok }, req: wdReq_(tok, c.dc, c.site, 0, {}) };
  });
  var catReqs = [], recReqs = [];
  fetchAllJson_(w1).forEach(function (r) {
    var tok = r.meta.tok, c = CFG.workday[tok];
    if (r.err || !r.data) { boardErr[tok] = r.err || 'no data'; return; }
    (r.data.jobPostings || []).forEach(function (j) { add(j, tok); });

    // category pass — turns 19,000 CVS postings into about 55
    var cats = [];
    (r.data.facets || []).forEach(function (f) {
      if (f.facetParameter !== 'jobFamilyGroup') return;
      (f.values || []).forEach(function (v) {
        var d = String(v.descriptor || '').toLowerCase();
        if (CFG.workdayCategoryKeywords.some(function (k) { return d.indexOf(k) >= 0; })) {
          cats.push({ id: v.id, count: v.count || 0 });
        }
      });
    });
    cats.slice(0, 3).forEach(function (cat) {
      for (var off = 0; off < Math.min(cat.count, 200); off += 20) {
        catReqs.push({ meta: { tok: tok },
          req: wdReq_(tok, c.dc, c.site, off, { jobFamilyGroup: [cat.id] }) });
      }
    });

    // recency pass — results are newest-first. Note `total` is only populated
    // on the first response and comes back 0 afterwards, so never page off it.
    var total = r.data.total || 0;
    for (var p = 1; p <= CFG.workdayRecencyPages; p++) {
      if (total && p * 20 >= total) break;
      recReqs.push({ meta: { tok: tok }, req: wdReq_(tok, c.dc, c.site, p * 20, {}) });
    }
  });

  // waves 2 and 3
  fetchAllJson_(catReqs).forEach(function (r) {
    if (r.data) (r.data.jobPostings || []).forEach(function (j) { add(j, r.meta.tok); });
  });
  fetchAllJson_(recReqs).forEach(function (r) {
    if (!r.data) return;
    (r.data.jobPostings || []).forEach(function (j) {
      var when = parseWorkdayDate_(j.postedOn);
      if (when && when < cutoff) return;
      add(j, r.meta.tok);
    });
  });

  var postings = [], boardRows = [];
  Object.keys(seen).forEach(function (uid) {
    var p = seen[uid];
    boardCount[p.company] = (boardCount[p.company] || 0) + 1;
    if (matches_(p)) postings.push(p);
  });
  toks.forEach(function (tok) {
    var c = CFG.workday[tok];
    var kept = postings.filter(function (p) { return p.company === c.name; }).length;
    boardRows.push([c.name, 'Workday', tok, rawCount[tok] || 0, kept, boardErr[tok] || 'ok']);
  });
  var rawUids = {};
  Object.keys(seen).forEach(function (uid) { rawUids[uid] = 1; });
  return { postings: postings, boardRows: boardRows, rawUids: rawUids };
}

// ----------------------------------------------------------------- sheet io

function readPrevious_(ss, tabName) {
  var sh = ss.getSheetByName(tabName);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var values = sh.getDataRange().getValues();
  var formulas = sh.getDataRange().getFormulas();
  var hdr = values[0];
  var iu = hdr.indexOf('UID');
  if (iu < 0) return map;
  var idx = {};
  TRACK.forEach(function (c) { var i = hdr.indexOf(c); if (i >= 0) idx[c] = i; });
  var iFirst = hdr.indexOf('First Seen');
  for (var r = 1; r < values.length; r++) {
    var uid = values[r][iu];
    if (!uid) continue;
    var rec = { firstSeen: iFirst >= 0 ? String(values[r][iFirst] || '') : '' };
    TRACK.forEach(function (c) {
      var v = idx[c] !== undefined ? values[r][idx[c]] : '';
      if (v instanceof Date) v = ymd_(v);
      rec[c] = v === null || v === undefined ? '' : String(v).trim();
    });
    // Prefer the formula: a Link cell reads back as "Open posting", so the
    // URL would be lost when the row moves to Closed.
    rec.row = values[r].map(function (v, i) {
      return formulas[r][i] ? formulas[r][i] : v;
    });
    rec.hdr = hdr;
    map[String(uid)] = rec;
  }
  return map;
}

function writeTab_(ss, name, headers, rows, opts) {
  opts = opts || {};
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  if (sh.getFilter()) sh.getFilter().remove();

  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(NAVY);
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  if (opts.widths) opts.widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  if (rows.length) sh.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  sh.setRowHeight(1, 26);
  return sh;
}

// ----------------------------------------------------------------- refresh

function refresh() {
  var ss = SpreadsheetApp.getActive();
  var t0 = new Date();
  ss.toast('Fetching boards…', 'Job Tracker', -1);

  var prevActive = readPrevious_(ss, 'Active');
  var prevClosed = readPrevious_(ss, 'Closed');

  var simple = collectSimpleBoards_();
  var wd = collectWorkday_();
  var all = simple.postings.concat(wd.postings);
  var boardRows = simple.boardRows.concat(wd.boardRows);
  var rawUids = simple.rawUids;
  Object.keys(wd.rawUids).forEach(function (u) { rawUids[u] = 1; });

  // de-dupe: the same role can appear on two systems
  var seenKey = {}, deduped = [];
  all.sort(function (a, b) { return (a.company + a.title).localeCompare(b.company + b.title); });
  all.forEach(function (p) {
    var k = (p.company + '|' + p.title + '|' + p.location).toLowerCase();
    if (seenKey[k]) return;
    seenKey[k] = 1;
    deduped.push(p);
  });

  var todayStr = ymd_(today_());
  var liveUids = {};
  deduped.forEach(function (p) {
    liveUids[p.uid] = 1;
    var prev = prevActive[p.uid] || prevClosed[p.uid];
    p.firstSeen = (prev && prev.firstSeen) ? prev.firstSeen : todayStr;
    TRACK.forEach(function (c) { p[c] = prev ? (prev[c] || '') : ''; });
  });

  deduped.sort(function (a, b) {
    return (b.posted ? b.posted.getTime() : 0) - (a.posted ? a.posted.getTime() : 0);
  });

  // ---- Active
  var hdr = TRACK.concat(['Title', 'Company', 'Location', 'Posted', 'Age (d)',
                          'First Seen', 'Source', 'Dept', 'Req', 'Link', 'UID']);
  var rows = deduped.map(function (p) {
    return TRACK.map(function (c) { return p[c]; }).concat([
      p.title, p.company, p.location,
      p.posted ? ymd_(p.posted) : '',
      p.posted ? daysBetween_(today_(), p.posted) : '',
      p.firstSeen, p.source, p.dept, p.req,
      p.url ? '=HYPERLINK("' + p.url + '","Open posting")' : '', p.uid
    ]);
  });
  var sh = writeTab_(ss, 'Active', hdr, rows,
    { widths: [90, 70, 90, 240, 330, 150, 210, 85, 60, 85, 90, 140, 100, 100, 260] });

  if (rows.length) {
    var n = rows.length + 1;
    sh.getRange(2, 1, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(true).build());
    sh.getRange(2, 2, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(PRIORITIES, true).setAllowInvalid(true).build());

    var range = sh.getRange(2, 1, rows.length, hdr.length);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$J2="' + todayStr + '"')
        .setBackground(NEW_HL).setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A2="Applied"')
        .setBackground(CREAM).setRanges([range]).build()
    ]);
    sh.hideColumns(hdr.length); // UID
  }

  // ---- New This Run
  var fresh = deduped.filter(function (p) { return p.firstSeen === todayStr; });
  writeTab_(ss, 'New This Run',
    ['Title', 'Company', 'Location', 'Posted', 'Source', 'Link'],
    fresh.map(function (p) {
      return [p.title, p.company, p.location, p.posted ? ymd_(p.posted) : '', p.source,
              p.url ? '=HYPERLINK("' + p.url + '","Open posting")' : ''];
    }), { widths: [330, 170, 220, 85, 90, 110] });

  // ---- Closed: anything previously live that stopped coming back
  var closedRows = [];
  Object.keys(prevClosed).forEach(function (uid) {
    if (liveUids[uid]) return;
    closedRows.push(prevClosed[uid].row);
  });
  Object.keys(prevActive).forEach(function (uid) {
    if (liveUids[uid] || prevClosed[uid]) return;
    // Still on the board, just no longer matching (usually a filter change).
    // That is not a closure, so don't report it as one.
    if (rawUids[uid]) return;
    var prev = prevActive[uid], row = prev.row, h = prev.hdr;
    function get(name) { var i = h.indexOf(name); return i >= 0 ? row[i] : ''; }
    var fs = prev.firstSeen || '';
    var days = '';
    if (fs) days = daysBetween_(today_(), new Date(fs));
    closedRows.push(TRACK.map(function (c) { return prev[c]; }).concat([
      get('Title'), get('Company'), get('Location'), fs, todayStr, days,
      get('Source'), get('Link'), uid
    ]));
  });
  writeTab_(ss, 'Closed',
    TRACK.concat(['Title', 'Company', 'Location', 'First Seen', 'Last Seen',
                  'Days Live', 'Source', 'Link', 'UID']),
    closedRows, { widths: [90, 70, 90, 220, 330, 170, 210, 85, 85, 75, 90, 110, 260] });

  // ---- Boards
  boardRows.sort(function (a, b) { return b[4] - a[4] || String(a[0]).localeCompare(String(b[0])); });
  writeTab_(ss, 'Boards',
    ['Company', 'ATS', 'Token', 'Postings Seen', 'Matched', 'Status'],
    boardRows, { widths: [210, 100, 180, 110, 80, 300] });

  // ---- Search Queries
  var since = ymd_(new Date(today_().getTime() - 7 * 86400000));
  writeTab_(ss, 'Search Queries', ['Query', 'With freshness filter', 'Run'],
    CFG.googleQueries.map(function (q) {
      var full = q + ' after:' + since;
      return [q, full, '=HYPERLINK("https://www.google.com/search?q=' +
              encodeURIComponent(full) + '","Search")'];
    }), { widths: [430, 480, 90] });

  // ---- Run Log
  writeTab_(ss, 'Run Log', ['Metric', 'Value'], [
    ['Last run', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')],
    ['Boards polled', boardRows.length],
    ['Boards erroring', boardRows.filter(function (b) { return b[5] !== 'ok'; }).length],
    ['Postings scanned', boardRows.reduce(function (a, b) { return a + b[3]; }, 0)],
    ['Matching (after de-dupe)', deduped.length],
    ['New since last run', fresh.length],
    ['Closed / removed', closedRows.length],
    ['Run time (s)', Math.round((new Date() - t0) / 1000)]
  ], { widths: [230, 380] });

  ss.setActiveSheet(ss.getSheetByName('New This Run'));
  ss.toast(deduped.length + ' roles, ' + fresh.length + ' new, ' +
           closedRows.length + ' closed.', 'Job Tracker', 8);
}
