/* Builds the retro deck (a list of slides, each a list of render.js elements)
 * from the /api/retro payload. Every content slide has an editable Insights
 * box (saved for everyone, exported into the .pptx). Every number is
 * clickable and opens the deals behind it. */
(function () {
  'use strict';
  var C = QR.C;
  var X0 = 96, XW = 1728;
  var INS = { y: 176, h: 84 };        // insights strip under the title
  var TOP = 282, BOTTOM = 936;        // content band on slides with the strip
  var GREEN = '#1E7F4F', RED = '#C8102E';
  var COL = { q1: '#AEB7D6', q2: C.navy, h: C.blue, target: '#D5DAE6' };
  var D = null;                        // payload

  // ---------- formatting ----------
  function num(v) { return v == null ? '-' : Number(v).toLocaleString('en-US'); }
  function money(v) {
    if (v == null) return '-';
    var a = Math.abs(v);
    if (a >= 1e6) return '$' + (v / 1e6).toFixed(2) + 'M';
    if (a >= 1e4) return '$' + Math.round(v / 1e3) + 'K';
    if (a >= 1e3) return '$' + (v / 1e3).toFixed(1) + 'K';
    return '$' + Math.round(v);
  }
  function pct(v) { return v == null || !isFinite(v) ? '-' : Math.round(v * 100) + '%'; }
  function days(v) { return v == null ? '-' : v + 'd'; }
  function first(n) { return String(n).split(' ')[0]; }
  // Chart labels: first name, plus last initial where two AEs share a first name.
  function shortNames(names) {
    return names.map(function (n) {
      var f = first(n), dup = names.filter(function (m) { return first(m) === f; }).length > 1, parts = String(n).split(' ');
      return dup && parts.length > 1 ? f + ' ' + parts[parts.length - 1][0] + '.' : f;
    });
  }
  // Stretch declared column widths so they add up to the table width (keeps group headers aligned).
  function fillCols(cols, w) { var sum = cols.reduce(function (a, c) { return a + c.w; }, 0); return cols.map(function (c) { return Object.assign({}, c, { w: c.w * w / sum }); }); }
  // Body row height that leaves `reserve` px for a chart under the table; null = no room for the chart.
  function rowsWithChart(n, reserve) { var rh = Math.min(40, Math.floor((BOTTOM - TOP - 38 - 22 - reserve) / (n + 1))); return rh >= 30 ? rh : null; }
  function attainFill(v) { return v == null ? null : v >= 1 ? '#C9E3C3' : v >= 0.85 ? '#E3F0DF' : v >= 0.6 ? '#FBF1D3' : '#F4D6D6'; }
  function targetTxt(t) { return t.status === 'set' ? money(t.value) : t.status === 'partial' ? money(t.value) + ' *' : t.status === 'na' ? 'N/A' : 'TBD'; }
  function P() { return D.periods; }
  var KEYS = ['compare', 'review', 'combined'];
  function lab(k) { return P()[k].short; }

  // ---------- shared chrome ----------
  function chrome(o, ctx) {
    var els = [{ type: 'image', src: '/logo.png', x: 1644, y: 66, w: 180, h: 35, position: 'right center' }];
    if (o.eyebrow) els.push({ type: 'text', x: X0, y: 62, w: 1400, h: 32, text: o.eyebrow, font: 'mono', size: 22, color: C.blue, tracking: 0.12, upper: true, bold: true });
    if (o.title) els.push({ type: 'text', x: X0, y: 96, w: 1520, h: 78, text: o.title, font: 'serif', size: 50, color: C.navy });
    if (o.definitions && o.definitions.length) {
      els.push({ type: 'line', x1: X0, x2: X0 + XW, y1: 948, y2: 948, color: C.soft, width: 1 });
      els.push({ type: 'text', x: X0, y: 954, w: XW, h: 66, size: 13, color: C.muted, lineHeight: 1.3,
        runs: o.definitions.map(function (d, i) { return [{ t: (i ? '     •  ' : '•  ') + d[0] + ': ', bold: true }, { t: d[1] }]; }).reduce(function (a, b) { return a.concat(b); }, []) });
    }
    els.push({ type: 'text', x: X0, y: 1026, w: 1300, h: 26, text: 'SpotDraft · Quarterly Sales Retro · ' + P().combined.label + (o.team ? ' · ' + o.team : ''), font: 'mono', size: 15, color: C.muted });
    els.push({ type: 'text', x: 1624, y: 1026, w: 200, h: 26, text: String(ctx.page).padStart(2, '0'), font: 'mono', size: 15, color: C.muted, align: 'right' });
    return els;
  }
  function insights(ctx, x, y, w, h, opts) {
    opts = opts || {};
    var key = ctx.slideId + (opts.suffix ? ':' + opts.suffix : '');
    return { type: 'insights', key: key, x: x, y: y, w: w, h: h, label: opts.label || 'Insights', text: ctx.insight(key), size: opts.size, placeholder: opts.placeholder || 'Click to add insights…' };
  }
  function strip(ctx) { return insights(ctx, X0, INS.y, XW, INS.h, { size: 18 }); }
  function heading(x, y, w, text) { return { type: 'text', x: x, y: y, w: w, h: 30, text: text, size: 21, bold: true, color: C.navy }; }

  // Column-group labels drawn above a table (e.g. "Q1 | Q2 | H1").
  function groupHeads(x, y, cols, groups) {
    var out = [], offs = [0];
    cols.forEach(function (c, i) { offs[i + 1] = offs[i] + c.w; });
    groups.forEach(function (g) {
      var gx = x + offs[g.from] + 6, gw = offs[g.to + 1] - offs[g.from] - 12;
      out.push({ type: 'text', x: gx, y: y, w: gw, h: 26, text: g.label, size: 17, bold: true, color: g.color || C.navy, align: 'center', font: 'mono', tracking: 0.06 });
      out.push({ type: 'rect', x: gx, y: y + 28, w: gw, h: 4, fill: g.color || C.navy });
    });
    return out;
  }
  function periodGroups(start, width, extra) {
    return KEYS.map(function (k, i) { return { label: k === 'combined' ? P().combined.short + ' (' + P().compare.short + ' + ' + P().review.short + ')' : P()[k].short, from: start + i * width, to: start + i * width + width - 1, color: k === 'combined' ? C.blue : k === 'review' ? C.navy : C.muted }; }).concat(extra || []);
  }

  // Row height that fits n body rows (plus header) into h px.
  function fitRows(n, h, max, min) { return Math.max(min || 32, Math.min(max || 46, Math.floor(h / (n + 1)))); }
  function cell(t, drill, title, o) { return Object.assign({ t: t, drill: drill || null, drillTitle: title }, o || {}); }

  function trend(a, b, lowerIsBetter) {
    if (a == null || b == null || a === b) return { t: '', color: C.muted };
    var up = b > a, good = lowerIsBetter ? !up : up;
    return { t: up ? ' ▲' : ' ▼', color: good ? GREEN : RED };
  }

  // ---------- opening ----------
  function titleSlide() {
    return function (ctx) {
      return [
        { type: 'rect', x: 0, y: 0, w: 1920, h: 1080, fill: C.navy },
        { type: 'image', src: '/logo-white.png', x: 120, y: 120, w: 300, h: 59 },
        { type: 'rect', x: 120, y: 450, w: 96, h: 8, fill: C.blue },
        { type: 'text', x: 120, y: 480, w: 1680, h: 140, text: 'Quarterly Sales Retro', font: 'serif', size: 100, color: C.white },
        { type: 'text', x: 120, y: 630, w: 1600, h: 60, text: P().combined.label + '  ·  ' + P().compare.short + ' vs ' + P().review.short, font: 'mono', size: 40, color: '#E2A58A', tracking: 0.06 },
        { type: 'text', x: 120, y: 700, w: 1600, h: 50, text: 'AE performance by team: revenue, MOFU and BOFU conversion, ACV and sales cycle', size: 26, color: C.grey },
        { type: 'text', x: 120, y: 1010, w: 1200, h: 30, text: (D.mock ? 'MOCK DATA · ' : '') + 'Live from HubSpot · data as of ' + new Date(D.generatedAt).toUTCString().replace(' GMT', ' UTC'), font: 'mono', size: 16, color: C.grey },
        { type: 'text', x: 1624, y: 1010, w: 200, h: 30, text: String(ctx.page).padStart(2, '0'), font: 'mono', size: 16, color: C.grey, align: 'right' }
      ];
    };
  }

  var COMMON_DEFS = [
    ['Team', 'AE\'s region on the Regional Interlocks roster; deals count by deal owner, whatever the deal region'],
    ['Achieved', 'Annual Subscription Fee $ of Sales Pipeline deals closed won (incl. Deal Desk Verified) in the period'],
    ['Target', 'Regional Interlocks New Business ARR targets; * = TBD for a quarter']
  ];

  function overviewSlide() {
    return function (ctx) {
      var els = chrome({ eyebrow: 'All teams', title: 'Teams at a glance: ' + P().compare.short + ' vs ' + P().review.short + ' and ' + P().combined.short,
        definitions: COMMON_DEFS.concat([['MOFU', 'MQLs created that became SQLs in the period ÷ MQLs created'], ['BOFU', 'SQLs won in the period ÷ SQLs in play (open at the start + produced)'], ['ACV / cycle', 'per closed-won deal; cycle = SQL date → close date']]) }, ctx);
      els.push(strip(ctx));
      var rows = D.teams.map(function (t) { return { name: t.title, p: t.total }; }).concat([{ name: 'All teams', p: D.company, bold: true }]);
      // Revenue table
      var cols = [{ label: 'Team', w: 228 }];
      KEYS.forEach(function () { cols.push({ label: 'Target', w: 166, align: 'right' }, { label: 'Achieved', w: 166, align: 'right' }, { label: 'Attainment', w: 168, align: 'right' }); });
      var y = TOP;
      els.push(heading(X0, y, 800, 'Revenue: target vs achieved'));
      els = els.concat(groupHeads(X0, y + 34, cols, periodGroups(1, 3)));
      var rh = 46;
      els.push({ type: 'table', x: X0, y: y + 72, w: XW, rowH: rh, size: 20, headSize: 16, cols: cols, rows: rows.map(function (r) {
        var out = [{ t: r.name, bold: true }];
        KEYS.forEach(function (k) {
          var m = r.p[k];
          out.push(cell(targetTxt(m.target), 'static', null, { bold: r.bold }));
          out.push(cell(money(m.revenue.achieved), m.revenue.drill.won, r.name + ' · ' + lab(k) + ' · Closed won', { bold: r.bold }));
          out.push(cell(pct(m.attainment), m.revenue.drill.won, r.name + ' · ' + lab(k) + ' · Closed won (attainment)', { bold: true, fill: attainFill(m.attainment) }));
        });
        return out;
      }) });
      // Conversion + deal dynamics table
      y = y + 72 + rh * 5 + 24;
      var cols2 = [{ label: 'Team', w: 228 }];
      ['MOFU', 'BOFU', 'ACV', 'Cycle'].forEach(function () { KEYS.forEach(function (k) { cols2.push({ label: lab(k), w: 125, align: 'right' }); }); });
      els.push(heading(X0, y, 1200, 'Conversion and deal dynamics'));
      var g2 = [['MOFU: MQL → SQL', 1], ['BOFU: SQL → won', 4], ['ACV', 7], ['Sales cycle', 10]].map(function (g) { return { label: g[0], from: g[1], to: g[1] + 2, color: C.navy }; });
      els = els.concat(groupHeads(X0, y + 34, cols2, g2));
      els.push({ type: 'table', x: X0, y: y + 72, w: XW, rowH: rh, size: 20, headSize: 16, cols: cols2, rows: rows.map(function (r) {
        var out = [{ t: r.name, bold: true }];
        KEYS.forEach(function (k) { var m = r.p[k]; out.push(cell(pct(m.mofu.funnel), m.mofu.drill.funnel, r.name + ' · ' + lab(k) + ' · MOFU funnel', { bold: k === 'combined' })); });
        KEYS.forEach(function (k) { var m = r.p[k]; out.push(cell(pct(m.bofu.winRate), m.bofu.drill.winRate, r.name + ' · ' + lab(k) + ' · BOFU win rate', { bold: k === 'combined' })); });
        KEYS.forEach(function (k) { var m = r.p[k]; out.push(cell(money(m.revenue.acv), m.revenue.drill.acv, r.name + ' · ' + lab(k) + ' · ACV', { bold: k === 'combined' })); });
        KEYS.forEach(function (k) { var m = r.p[k]; out.push(cell(days(m.revenue.cycle), m.revenue.drill.cycle, r.name + ' · ' + lab(k) + ' · Sales cycle', { bold: k === 'combined' })); });
        return out;
      }) });
      return els;
    };
  }

  // ---------- team slides ----------
  function teamRows(team) { return team.reps.map(function (r) { return { name: r.name, p: r.periods }; }).concat([{ name: team.title + ' team', p: team.total, total: true }]); }

  function sectionSlide(team) {
    return function (ctx) {
      var c = team.total.combined, q1 = team.total.compare, q2 = team.total.review;
      var els = [
        { type: 'rect', x: 0, y: 0, w: 1920, h: 1080, fill: C.navy },
        { type: 'image', src: '/logo-white.png', x: 120, y: 100, w: 240, h: 47 },
        { type: 'rect', x: 120, y: 236, w: 96, h: 8, fill: C.blue },
        { type: 'text', x: 120, y: 260, w: 1200, h: 130, text: team.title, font: 'serif', size: 110, color: C.white },
        { type: 'text', x: 120, y: 392, w: 1600, h: 40, text: team.long + ' · ' + team.reps.length + ' AE' + (team.reps.length === 1 ? '' : 's') + ' · ' + P().combined.label, font: 'mono', size: 26, color: '#E2A58A', tracking: 0.04 },
        { type: 'text', x: 120, y: 440, w: 1680, h: 70, text: team.reps.map(function (r) { return r.name; }).join('  ·  ') || 'No AEs on the roster', size: 22, color: C.grey, lineHeight: 1.4 }
      ];
      var tiles = [
        ['Target ' + P().combined.short, targetTxt(c.target), 'static', null, lab('compare') + ' ' + targetTxt(q1.target) + ' · ' + lab('review') + ' ' + targetTxt(q2.target)],
        ['Achieved ' + P().combined.short, money(c.revenue.achieved), c.revenue.drill.won, 'Closed won', lab('compare') + ' ' + money(q1.revenue.achieved) + ' · ' + lab('review') + ' ' + money(q2.revenue.achieved)],
        ['Attainment', pct(c.attainment), c.revenue.drill.won, 'Closed won', lab('compare') + ' ' + pct(q1.attainment) + ' · ' + lab('review') + ' ' + pct(q2.attainment)],
        ['Deals won', num(c.revenue.won), c.revenue.drill.won, 'Closed won', lab('compare') + ' ' + num(q1.revenue.won) + ' · ' + lab('review') + ' ' + num(q2.revenue.won)],
        ['MOFU: MQL → SQL', pct(c.mofu.funnel), c.mofu.drill.funnel, 'MOFU funnel', lab('compare') + ' ' + pct(q1.mofu.funnel) + ' · ' + lab('review') + ' ' + pct(q2.mofu.funnel)],
        ['BOFU: win rate', pct(c.bofu.winRate), c.bofu.drill.winRate, 'BOFU win rate', lab('compare') + ' ' + pct(q1.bofu.winRate) + ' · ' + lab('review') + ' ' + pct(q2.bofu.winRate)],
        ['ACV', money(c.revenue.acv), c.revenue.drill.acv, 'ACV', lab('compare') + ' ' + money(q1.revenue.acv) + ' · ' + lab('review') + ' ' + money(q2.revenue.acv)],
        ['Sales cycle', days(c.revenue.cycle), c.revenue.drill.cycle, 'Sales cycle', lab('compare') + ' ' + days(q1.revenue.cycle) + ' · ' + lab('review') + ' ' + days(q2.revenue.cycle)]
      ];
      var tw = (1680 - 3 * 24) / 4;
      tiles.forEach(function (t, i) {
        var x = 120 + (i % 4) * (tw + 24), y = 590 + Math.floor(i / 4) * 196;
        els.push({ type: 'rect', x: x, y: y, w: tw, h: 172, fill: '#16284A', radius: 12 });
        els.push({ type: 'text', x: x + 24, y: y + 18, w: tw - 48, h: 26, text: t[0], font: 'mono', size: 17, color: '#E2A58A', upper: true, tracking: 0.08, bold: true });
        els.push({ type: 'text', x: x + 24, y: y + 50, w: tw - 48, h: 70, text: t[1], font: 'serif', size: 54, color: C.white, drill: t[2], drillTitle: team.title + ' · ' + P().combined.short + ' · ' + (t[3] || t[0]) });
        els.push({ type: 'text', x: x + 24, y: y + 126, w: tw - 48, h: 30, text: t[4], size: 18, color: C.grey });
      });
      els.push({ type: 'text', x: 1624, y: 1010, w: 200, h: 30, text: String(ctx.page).padStart(2, '0'), font: 'mono', size: 16, color: C.grey, align: 'right' });
      return els;
    };
  }

  function scorecardSlide(team) {
    return function (ctx) {
      var els = chrome({ eyebrow: team.title + ' · AE scorecard', title: 'AE scorecard: ' + P().combined.short + ' with ' + P().compare.short + ' → ' + P().review.short, team: team.title,
        definitions: COMMON_DEFS.concat([['Big number', P().combined.short + ' (both quarters together)'], ['Small line', P().compare.short + ' → ' + P().review.short + ' with direction: ▲ / ▼ green = better (a shorter cycle is better)']]) }, ctx);
      els.push(strip(ctx));
      var rows = teamRows(team);
      var cols = [{ label: 'AE', w: 268 }, { label: 'Attainment', w: 210, align: 'right' }, { label: 'Achieved', w: 210, align: 'right' }, { label: 'Deals won', w: 170, align: 'right' },
        { label: 'MOFU: MQL → SQL', w: 220, align: 'right' }, { label: 'BOFU: win rate', w: 220, align: 'right' }, { label: 'ACV', w: 220, align: 'right' }, { label: 'Sales cycle', w: 210, align: 'right' }];
      var rh = fitRows(rows.length, BOTTOM - TOP - 10, 70, 46);
      var metric = function (r, get, fmt, drillOf, name, lower, fill) {
        var a = get(r.p.compare), b = get(r.p.review), h = get(r.p.combined), tr = trend(a, b, lower);
        return { runs: [{ t: fmt(h), bold: true, size: rh >= 60 ? 22 : 19 }, { t: '\n' + lab('compare') + ' ' + fmt(a) + ' → ' + lab('review') + ' ' + fmt(b), size: 13, color: C.muted }, { t: tr.t, size: 13, color: tr.color, bold: true }],
          drill: drillOf(r.p.combined), drillTitle: r.name + ' · ' + P().combined.short + ' · ' + name, fill: fill ? fill(h) : null };
      };
      els.push({ type: 'table', x: X0, y: TOP, w: XW, rowH: rh, headerH: 44, size: 14, headSize: 16, wrap: true, cols: cols, rows: rows.map(function (r) {
        return [{ runs: [{ t: r.name, bold: true, size: 19 }], fill: r.total ? C.soft : null },
          metric(r, function (m) { return m.attainment; }, pct, function (m) { return m.revenue.drill.won; }, 'Closed won (attainment)', false, attainFill),
          metric(r, function (m) { return m.revenue.achieved; }, money, function (m) { return m.revenue.drill.won; }, 'Closed won'),
          metric(r, function (m) { return m.revenue.won; }, num, function (m) { return m.revenue.drill.won; }, 'Deals won'),
          metric(r, function (m) { return m.mofu.funnel; }, pct, function (m) { return m.mofu.drill.funnel; }, 'MOFU funnel'),
          metric(r, function (m) { return m.bofu.winRate; }, pct, function (m) { return m.bofu.drill.winRate; }, 'BOFU win rate'),
          metric(r, function (m) { return m.revenue.acv; }, money, function (m) { return m.revenue.drill.acv; }, 'ACV'),
          metric(r, function (m) { return m.revenue.cycle; }, days, function (m) { return m.revenue.drill.cycle; }, 'Sales cycle', true)
        ];
      }) });
      return els;
    };
  }

  function revenueSlide(team) {
    return function (ctx) {
      var els = chrome({ eyebrow: team.title + ' · Revenue', title: 'Revenue: target vs achieved by AE', team: team.title,
        definitions: COMMON_DEFS.concat([['Attainment', 'achieved ÷ target; green ≥ 100%, light green ≥ 85%, amber ≥ 60%, red below'], ['N/A / TBD', 'not in the role that quarter / no target set yet']]) }, ctx);
      els.push(strip(ctx));
      var rows = teamRows(team);
      var cols = [{ label: 'AE', w: 220 }];
      KEYS.forEach(function () { cols.push({ label: 'Target', w: 108, align: 'right' }, { label: 'Achieved', w: 108, align: 'right' }, { label: 'Att.', w: 82, align: 'right' }); });
      var tw = cols.reduce(function (a, c) { return a + c.w; }, 0);
      var rh = fitRows(rows.length, BOTTOM - TOP - 40, 50, 34);
      els = els.concat(groupHeads(X0, TOP, cols, periodGroups(1, 3)));
      els.push({ type: 'table', x: X0, y: TOP + 38, w: tw, rowH: rh, size: 18, headSize: 15, cols: cols, rows: rows.map(function (r) {
        var out = [{ t: r.name, bold: true, fill: r.total ? C.soft : null }];
        KEYS.forEach(function (k) {
          var m = r.p[k], b = r.total || k === 'combined';
          out.push(cell(targetTxt(m.target), 'static', null, { bold: r.total, fill: r.total ? C.soft : null }));
          out.push(cell(money(m.revenue.achieved), m.revenue.drill.won, r.name + ' · ' + lab(k) + ' · Closed won', { bold: b, fill: r.total ? C.soft : null }));
          out.push(cell(pct(m.attainment), m.revenue.drill.won, r.name + ' · ' + lab(k) + ' · Closed won (attainment)', { bold: true, fill: attainFill(m.attainment) }));
        });
        return out;
      }) });
      // Chart: combined period, target vs achieved per AE.
      var cx = X0 + tw + 40, cw = X0 + XW - cx;
      els.push(heading(cx, TOP, cw, P().combined.short + ' target vs achieved'));
      var reps = team.reps;
      els.push({ type: 'bars', kind: 'barH', x: cx, y: TOP + 36, w: cw, h: Math.min(BOTTOM - TOP - 40, 70 + reps.length * 64), catWidth: 120,
        categories: shortNames(reps.map(function (r) { return r.name; })), fmt: money,
        series: [{ name: 'Target', color: COL.target, values: reps.map(function (r) { return r.periods.combined.target.value || 0; }) },
          { name: 'Achieved', color: C.navy, values: reps.map(function (r) { return r.periods.combined.revenue.achieved; }) }],
        drills: [reps.map(function () { return 'static'; }), reps.map(function (r) { return r.periods.combined.revenue.drill.won; })] });
      return els;
    };
  }

  // Shared layout for the two conversion slides: table on top, grouped bars below if there is room.
  function conversionSlide(team, o) {
    return function (ctx) {
      var els = chrome({ eyebrow: team.title + ' · ' + o.eyebrow, title: o.title, team: team.title, definitions: o.defs }, ctx);
      els.push(strip(ctx));
      var rows = teamRows(team);
      var cols = [{ label: 'AE', w: 240 }];
      KEYS.forEach(function () { o.cols.forEach(function (c) { cols.push({ label: c.label, w: c.w, align: 'right' }); }); });
      o.extra.cols.forEach(function (c) { cols.push({ label: c.label, w: c.w, align: 'right' }); });
      var n = o.cols.length;
      var groups = periodGroups(1, n, [{ label: o.extra.label, from: 1 + 3 * n, to: 3 * n + o.extra.cols.length, color: C.blue }]);
      cols = fillCols(cols, XW);
      var rh = rowsWithChart(rows.length, 230), withChart = rh != null;
      if (!withChart) rh = fitRows(rows.length, BOTTOM - TOP - 40, 44, 30);
      els = els.concat(groupHeads(X0, TOP, cols, groups));
      els.push({ type: 'table', x: X0, y: TOP + 38, w: XW, rowH: rh, size: 18, headSize: 15, cols: cols, rows: rows.map(function (r) {
        var out = [{ t: r.name, bold: true, fill: r.total ? C.soft : null }];
        KEYS.forEach(function (k) { o.cols.forEach(function (c) { out.push(c.cell(r.p[k], r.name + ' · ' + lab(k) + ' · ', r.total || (c.pct && k === 'combined'))); }); });
        o.extra.cols.forEach(function (c) { out.push(c.cell(r.p[c.key], r.name + ' · ' + lab(c.key) + ' · ', r.total)); });
        return out.map(function (x, i) { return i && r.total && typeof x === 'object' && !x.fill ? Object.assign({}, x, { fill: C.soft }) : x; });
      }) });
      if (withChart) {
        var ty = TOP + 38 + rh * (rows.length + 1) + 22, h = BOTTOM - ty;
        if (h >= 200) {
          els.push(heading(X0, ty, 1200, o.chartTitle));
          var cats = shortNames(rows.map(function (r) { return r.name; })).map(function (n, i) { return rows[i].total ? 'Team' : n; });
          els.push({ type: 'bars', kind: 'bar', x: X0, y: ty + 30, w: XW, h: Math.min(h - 30, 400), categories: cats, fmt: pct,
            series: KEYS.map(function (k) { return { name: k === 'combined' ? P().combined.short : lab(k), color: k === 'combined' ? COL.h : k === 'review' ? COL.q2 : COL.q1, values: rows.map(function (r) { var v = o.chartValue(r.p[k]); return v == null ? 0 : v; }) }; }),
            drills: KEYS.map(function (k) { return rows.map(function (r) { return o.chartDrill(r.p[k]); }); }) });
        }
      }
      return els;
    };
  }

  function mofuSlide(team) {
    return conversionSlide(team, {
      eyebrow: 'MOFU conversion', title: 'MOFU conversion: MQL → SQL by AE',
      cols: [
        { label: 'MQLs', w: 104, cell: function (m, t, b) { return cell(num(m.mofu.mqlNew), m.mofu.drill.mqlNew, t + 'MQLs created', { bold: b }); } },
        { label: '→ SQL', w: 104, cell: function (m, t, b) { return cell(num(m.mofu.mqlNewToSql), m.mofu.drill.mqlNewToSql, t + 'MQLs that became SQLs', { bold: b }); } },
        { label: 'Funnel %', w: 120, pct: true, cell: function (m, t, b) { return cell(pct(m.mofu.funnel), m.mofu.drill.funnel, t + 'MOFU funnel', { bold: true, color: b ? C.blue : null }); } }
      ],
      extra: { label: 'Production %', cols: KEYS.map(function (k) { return { key: k, label: lab(k), w: 118, cell: function (m, t, b) { return cell(pct(m.mofu.production), m.mofu.drill.production, t + 'MOFU production', { bold: b }); } }; }) },
      chartTitle: 'Funnel % by AE: ' + lab('compare') + ' vs ' + lab('review') + ' vs ' + P().combined.short,
      chartValue: function (m) { return m.mofu.funnel; }, chartDrill: function (m) { return m.mofu.drill.funnel; },
      defs: [['MQL', 'Sales Pipeline deal that entered Scheduled (by deal owner)'], ['SQL', 'SQL date in the period and stage Qualified to buy or later (closed-won deals with no SQL date count in their close quarter)'],
        ['Funnel %', 'MQLs created in the period that became SQLs in it ÷ MQLs created'], ['Production %', 'all SQLs produced in the period ÷ (MQLs open at the start + MQLs created)']]
    });
  }

  function bofuSlide(team) {
    return conversionSlide(team, {
      eyebrow: 'BOFU conversion', title: 'BOFU conversion: SQL → Closed won by AE',
      cols: [
        { label: 'In play', w: 104, cell: function (m, t, b) { return cell(num(m.bofu.inPlay), m.bofu.drill.inPlay, t + 'SQLs in play', { bold: b }); } },
        { label: 'Won', w: 104, cell: function (m, t, b) { return cell(num(m.bofu.won), m.bofu.drill.won, t + 'SQLs won', { bold: b }); } },
        { label: 'Win rate', w: 120, pct: true, cell: function (m, t, b) { return cell(pct(m.bofu.winRate), m.bofu.drill.winRate, t + 'BOFU win rate', { bold: true, color: b ? C.blue : null }); } }
      ],
      extra: { label: 'New-SQL win %', cols: KEYS.map(function (k) { return { key: k, label: lab(k), w: 118, cell: function (m, t, b) { return cell(pct(m.bofu.funnel), m.bofu.drill.funnel, t + 'Win rate on SQLs produced in the period', { bold: b }); } }; }) },
      chartTitle: 'Win rate by AE: ' + lab('compare') + ' vs ' + lab('review') + ' vs ' + P().combined.short,
      chartValue: function (m) { return m.bofu.winRate; }, chartDrill: function (m) { return m.bofu.drill.winRate; },
      defs: [['In play', 'SQLs open at the start of the period + SQLs produced in it (by deal owner)'], ['Won', 'in-play SQLs closed won in the period'],
        ['Win rate', 'won ÷ in play'], ['New-SQL win %', 'SQLs produced in the period and won in it ÷ SQLs produced'], ['Lost', 'closed lost, debooked or slipped back to a pre-SQL stage (see the drill-down)']]
    });
  }

  function dynamicsSlide(team) {
    return function (ctx) {
      var els = chrome({ eyebrow: team.title + ' · Deal dynamics', title: 'ACV and sales cycle by AE', team: team.title,
        definitions: [['ACV', 'Annual Subscription Fee $ won ÷ deals won in the period'], ['Sales cycle', 'average days from SQL date to close date, closed-won deals with an SQL date (others are flagged in the drill-down)'], ['Team', 'deals count by deal owner']] }, ctx);
      els.push(strip(ctx));
      var rows = teamRows(team);
      var cols = [{ label: 'AE', w: 300 }];
      KEYS.forEach(function () { cols.push({ label: 'Deals won', w: 136, align: 'right' }, { label: 'ACV', w: 158, align: 'right' }, { label: 'Sales cycle', w: 182, align: 'right' }); });
      // Two half-width charts only fit up to 6 AEs; beyond that the table carries the numbers.
      var rh = rows.length <= 7 ? rowsWithChart(rows.length, 260) : null, withChart = rh != null;
      if (!withChart) rh = fitRows(rows.length, BOTTOM - TOP - 40, 48, 30);
      els = els.concat(groupHeads(X0, TOP, cols, periodGroups(1, 3)));
      els.push({ type: 'table', x: X0, y: TOP + 38, w: XW, rowH: rh, size: 18, headSize: 15, cols: cols, rows: rows.map(function (r) {
        var out = [{ t: r.name, bold: true, fill: r.total ? C.soft : null }];
        KEYS.forEach(function (k) {
          var m = r.p[k], t = r.name + ' · ' + lab(k) + ' · ', b = r.total || k === 'combined', f = r.total ? C.soft : null;
          out.push(cell(num(m.revenue.won), m.revenue.drill.won, t + 'Deals won', { bold: b, fill: f }));
          out.push(cell(money(m.revenue.acv), m.revenue.drill.acv, t + 'ACV', { bold: b, fill: f }));
          out.push(cell(days(m.revenue.cycle) + (m.revenue.noSqlDate ? ' †' : ''), m.revenue.drill.cycle, t + 'Sales cycle', { bold: b, fill: f }));
        });
        return out;
      }) });
      if (withChart) {
        var ty = TOP + 38 + rh * (rows.length + 1) + 22, h = BOTTOM - ty;
        if (h >= 200) {
          var half = (XW - 48) / 2, cats = shortNames(rows.map(function (r) { return r.name; })).map(function (n, i) { return rows[i].total ? 'Team' : n; });
          var ser = function (get) { return KEYS.map(function (k) { return { name: k === 'combined' ? P().combined.short : lab(k), color: k === 'combined' ? COL.h : k === 'review' ? COL.q2 : COL.q1, values: rows.map(function (r) { return get(r.p[k]) || 0; }) }; }); };
          els.push(heading(X0, ty, half, 'ACV by AE'));
          els.push({ type: 'bars', kind: 'bar', x: X0, y: ty + 30, w: half, h: Math.min(h - 30, 400), categories: cats, fmt: money, series: ser(function (m) { return m.revenue.acv; }),
            drills: KEYS.map(function (k) { return rows.map(function (r) { return r.p[k].revenue.drill.acv; }); }) });
          els.push(heading(X0 + half + 48, ty, half, 'Sales cycle (days) by AE'));
          els.push({ type: 'bars', kind: 'bar', x: X0 + half + 48, y: ty + 30, w: half, h: Math.min(h - 30, 400), categories: cats, fmt: function (v) { return v + 'd'; }, series: ser(function (m) { return m.revenue.cycle; }),
            drills: KEYS.map(function (k) { return rows.map(function (r) { return r.p[k].revenue.drill.cycle; }); }) });
        }
      }
      if (rows.some(function (r) { return KEYS.some(function (k) { return r.p[k].revenue.noSqlDate; }); })) {
        var defsIdx = els.findIndex(function (e) { return e.type === 'text' && e.y === 954; });
        if (defsIdx > -1) els[defsIdx].runs = els[defsIdx].runs.concat([{ t: '     •  † ', bold: true }, { t: 'some won deals have no SQL date and are left out of the cycle average' }]);
      }
      return els;
    };
  }

  function retroNotesSlide(team) {
    return function (ctx) {
      var els = chrome({ eyebrow: team.title + ' · Looking back', title: team.title + ': what went right, what could be better, what\'s next', team: team.title }, ctx);
      var w = (XW - 2 * 32) / 3;
      [['right', 'What went right?'], ['better', 'What could have been better?'], ['next', 'What\'s next?']].forEach(function (b, i) {
        els.push(insights(ctx, X0 + i * (w + 32), 196, w, 740, { suffix: b[0], label: b[1], size: 20, placeholder: 'Click to add notes…' }));
      });
      return els;
    };
  }

  function definitionsSlide() {
    return function (ctx) {
      var els = chrome({ eyebrow: 'Appendix', title: 'How the numbers are built' }, ctx);
      var src = D.targets.source === 'live' ? 'live from the Regional Interlocks Inputs tab' : 'the Regional Interlocks Inputs tab snapshot of ' + D.targets.asOf + ' (live copy not reachable)';
      var left = [
        ['Periods: ', P().compare.label + ' and ' + P().review.label + ' (fiscal year starts 1 May), and ' + P().combined.short + ' = both quarters together, computed on the combined deal set (not an average of the two quarters).'],
        ['Teams: ', 'each AE\'s region on the Regional Interlocks roster. Every deal counts for its deal owner\'s team, whatever the deal region (e.g. an EMEA deal owned by a NAM AE counts for NAM). Team rows = all the team\'s AEs together.'],
        ['AEs shown: ', 'AEs on the roster with a target in either quarter, or with any MQL, SQL or closed-won deal in the period.'],
        ['Targets: ', 'New Business ARR targets, ' + src + '. ' + P().combined.short + ' target = sum of the two quarters; * = one quarter is TBD.'],
        ['Achieved: ', 'Annual Subscription Fee $ of Sales Pipeline deals closed won in the period (Closed won and Deal Desk Verified Closed Won), by close date.']
      ];
      var right = [
        ['MQL: ', 'Sales Pipeline deal that entered Scheduled.'],
        ['SQL: ', 'SQL date (Became Qualified to Buy Date) in the period and current stage Qualified to buy, Presentation done, Proposal shared, Contracting, Closed won, Closed lost, Debook or Deal Desk Verified Closed Won. A closed-won deal with no SQL date counts as an SQL in the quarter it closed.'],
        ['MOFU: ', 'funnel % = MQLs created in the period that became SQLs in it ÷ MQLs created; production % = SQLs produced ÷ (MQLs open at the start + MQLs created).'],
        ['BOFU: ', 'win rate = SQLs won in the period ÷ SQLs in play (open at the start + produced); new-SQL win % = SQLs produced and won in the period ÷ SQLs produced. Slip-backs to pre-SQL stages count as lost.'],
        ['ACV and sales cycle: ', 'ACV = ARR won ÷ deals won. Sales cycle = average days from SQL date to close date of deals won in the period.']
      ];
      var blk = function (x, list) { return { type: 'text', x: x, y: 196, w: 820, h: 740, size: 19, color: C.ink, lineHeight: 1.4, runs: list.map(function (l, i) { return [{ t: (i ? '\n\n' : ''), size: 10 }, { t: l[0], bold: true, color: C.navy }, { t: l[1] }]; }).reduce(function (a, b) { return a.concat(b); }, []) }; };
      els.push(blk(X0, left), blk(X0 + 908, right));
      return els;
    };
  }

  function buildDeck(data) {
    D = data;
    var slides = [
      { id: 'title', section: 'Opening', title: 'Title', build: titleSlide() },
      { id: 'overview', section: 'Opening', title: 'Teams at a glance', build: overviewSlide() }
    ];
    data.teams.forEach(function (t) {
      var sec = t.title + ' team';
      slides.push(
        { id: t.key + '-section', section: sec, title: t.title + ': section title', build: sectionSlide(t) },
        { id: t.key + '-scorecard', section: sec, title: t.title + ': AE scorecard', build: scorecardSlide(t) },
        { id: t.key + '-revenue', section: sec, title: t.title + ': revenue vs target', build: revenueSlide(t) },
        { id: t.key + '-mofu', section: sec, title: t.title + ': MOFU conversion', build: mofuSlide(t) },
        { id: t.key + '-bofu', section: sec, title: t.title + ': BOFU conversion', build: bofuSlide(t) },
        { id: t.key + '-dynamics', section: sec, title: t.title + ': ACV and sales cycle', build: dynamicsSlide(t) },
        { id: t.key + '-retro', section: sec, title: t.title + ': looking back', build: retroNotesSlide(t) }
      );
    });
    slides.push({ id: 'definitions', section: 'Appendix', title: 'How the numbers are built', build: definitionsSlide() });
    return slides;
  }

  window.QRDeck = { buildDeck: buildDeck };
})();
