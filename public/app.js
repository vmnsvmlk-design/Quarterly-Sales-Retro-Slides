(function () {
  'use strict';
  var state = { slides: [], idx: 0, data: null, busy: false, insights: {}, pick: new Set() };
  var $ = function (id) { return document.getElementById(id); };
  var retroKey = function () { return state.data.periods.combined.label; };

  function toast(msg) { var t = $('toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toast._t); toast._t = setTimeout(function () { t.classList.remove('on'); }, 3200); }

  function api(method, path, body) {
    var sep = path.indexOf('?') > -1 ? '&' : '?';
    return fetch(path + sep + 'retro=' + encodeURIComponent(retroKey()), {
      method: method, headers: body ? { 'Content-Type': 'application/json' } : {}, body: body ? JSON.stringify(body) : undefined
    }).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); });
  }

  function elementsFor(slide, page) {
    return slide.build({ page: page, slideId: slide.id, insight: function (key) { var v = state.insights[key]; return v ? v.text : ''; } });
  }
  function slideHtml(slide, page) { return '<div class="slide">' + QR.toHtml(elementsFor(slide, page)) + '</div>'; }

  function fit() {
    var st = $('stage'), fr = $('frame');
    var s = Math.max(0.15, Math.min((st.clientWidth - 64) / 1920, (st.clientHeight - 48) / 1080));
    fr.style.width = 1920 * s + 'px'; fr.style.height = 1080 * s + 'px';
    var inner = fr.querySelector('.slide'); if (inner) { inner.style.transform = 'scale(' + s + ')'; inner.style.transformOrigin = '0 0'; }
  }

  function wireInsights() {
    $('frame').querySelectorAll('[data-insights]').forEach(function (el) {
      var key = el.dataset.insights, cur = state.insights[key];
      el.contentEditable = 'true';
      el.classList.add('editable');
      el.innerText = cur ? cur.text : '';
      el.dataset.placeholder = el.dataset.ph || 'Click to add insights…';
      el.addEventListener('blur', function () {
        var text = el.innerText.replace(/ /g, ' ').trim();
        if ((cur ? cur.text : '') === text) return;
        api('PUT', '/api/insights/' + encodeURIComponent(key), { text: text }).then(function (v) {
          if (text) state.insights[key] = v; else delete state.insights[key];
          cur = state.insights[key]; refreshThumb(state.idx); toast('Insights saved for everyone');
        }).catch(function () { toast('Could not save. Check the connection and try again'); });
      });
      el.addEventListener('keydown', function (e) { e.stopPropagation(); });
    });
  }

  function renderMain() {
    var s = state.slides[state.idx];
    $('frame').innerHTML = slideHtml(s, state.idx + 1);
    fit();
    wireInsights();
    $('counter').textContent = String(state.idx + 1).padStart(2, '0') + ' / ' + String(state.slides.length).padStart(2, '0');
    document.querySelectorAll('.thumb').forEach(function (el) { el.classList.toggle('active', Number(el.dataset.i) === state.idx); });
    var active = document.querySelector('.thumb.active'); if (active) active.scrollIntoView({ block: 'nearest' });
  }

  function refreshThumb(i) {
    var pv = document.querySelector('.thumb[data-i="' + i + '"] .pv');
    if (pv) pv.innerHTML = slideHtml(state.slides[i], i + 1);
  }

  function renderThumbs() {
    var html = '', section = null;
    state.slides.forEach(function (s, i) {
      if (s.section !== section) { section = s.section; html += '<div class="secHead">' + QR.esc(section) + '</div>'; }
      html += '<div class="thumb" data-i="' + i + '"><div class="n">' + String(i + 1).padStart(2, '0') + '</div>' +
        '<div><div class="pv">' + slideHtml(s, i + 1) + '</div><div class="cap">' + QR.esc(s.title) + '</div></div></div>';
    });
    $('thumbs').innerHTML = html;
    $('slideCount').textContent = 'SLIDES · ' + state.slides.length;
  }

  function go(d) { state.idx = Math.max(0, Math.min(state.slides.length - 1, state.idx + d)); renderMain(); }

  // ---------- download picker ----------
  function hasInsights(s) { return Object.keys(state.insights).some(function (k) { return k === s.id || k.indexOf(s.id + ':') === 0; }); }
  function renderPicker() {
    var html = '', sections = [];
    state.slides.forEach(function (s) { if (sections.indexOf(s.section) < 0) sections.push(s.section); });
    sections.forEach(function (sec) {
      var list = state.slides.map(function (s, i) { return { s: s, i: i }; }).filter(function (x) { return x.s.section === sec; });
      var allOn = list.every(function (x) { return state.pick.has(x.s.id); });
      html += '<div class="pickSec"><span>' + QR.esc(sec) + '</span><label><input type="checkbox" data-sec="' + QR.esc(sec) + '"' + (allOn ? ' checked' : '') + '> All in section</label></div>';
      list.forEach(function (x) {
        html += '<label class="pickRow"><input type="checkbox" data-pick="' + x.s.id + '"' + (state.pick.has(x.s.id) ? ' checked' : '') + '><span class="n">' + String(x.i + 1).padStart(2, '0') + '</span><span>' + QR.esc(x.s.title) + '</span>' + (hasInsights(x.s) ? '<span class="ins">has insights</span>' : '') + '</label>';
      });
    });
    $('pickList').innerHTML = html;
    var n = state.pick.size;
    $('pickerMeta').textContent = n + ' of ' + state.slides.length + ' slides selected';
    $('pickGoLabel').textContent = state.busy ? 'Building .pptx…' : 'Download ' + n + ' slide' + (n === 1 ? '' : 's');
    $('pickGo').disabled = !n || state.busy;
  }
  function openPicker() { renderPicker(); $('picker').classList.add('on'); }

  async function exportPptx() {
    var list = state.slides.filter(function (s) { return state.pick.has(s.id); });
    if (!list.length || state.busy) return;
    state.busy = true; renderPicker();
    try {
      var pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.title = 'SpotDraft Sales Retro ' + retroKey();
      list.forEach(function (s, i) { QR.toPptx(pptx.addSlide(), elementsFor(s, i + 1), pptx); });
      await pptx.writeFile({ fileName: 'SpotDraft Sales Retro ' + retroKey() + '.pptx' });
      toast('Downloaded ' + list.length + ' slide' + (list.length > 1 ? 's' : '') + ' as .pptx');
      $('picker').classList.remove('on');
    } catch (e) { console.error(e); toast('Download failed. See the browser console'); }
    state.busy = false; renderPicker();
  }

  // ---------- click-through lists ----------
  var drillState = null;
  function fmtCell(v, i, data) {
    if (v == null || v === '') return '';
    if (i === data.moneyCol) return '$' + Math.round(v).toLocaleString('en-US');
    return String(v);
  }
  function renderDrill() {
    var d = drillState; if (!d) return;
    var q = ($('drillSearch').value || '').toLowerCase();
    var rows = d.data.rows.filter(function (r) { return !q || r.cells.join(' ').toLowerCase().indexOf(q) > -1; });
    var arr = rows.reduce(function (a, r) { return a + (Number(r.cells[d.data.moneyCol]) || 0); }, 0);
    $('drillMeta').textContent = rows.length + ' deal' + (rows.length === 1 ? '' : 's') + (q ? ' matching' : '') + ' · Annual Subscription Fee $' + Math.round(arr).toLocaleString('en-US');
    var tagCol = d.data.columns.length === 13 ? 12 : -1;
    $('drillTable').innerHTML = '<thead><tr>' + d.data.columns.map(function (c) { return '<th>' + QR.esc(c) + '</th>'; }).join('') + '</tr></thead><tbody>' +
      rows.map(function (r) {
        return '<tr>' + r.cells.map(function (v, i) {
          var txt = QR.esc(fmtCell(v, i, d.data));
          if (i === 0 && r.url) txt = '<a href="' + QR.esc(r.url) + '" target="_blank" rel="noopener">' + txt + '</a>';
          return '<td class="' + (i === d.data.moneyCol || i === 9 ? 'num' : i === tagCol ? 'tag' : '') + '">' + txt + '</td>';
        }).join('') + '</tr>';
      }).join('') +
      (rows.length ? '' : '<tr><td colspan="' + d.data.columns.length + '" class="empty">No deals</td></tr>') + '</tbody>';
  }
  function openDrill(key, title) {
    if (key === 'static') { toast('Target: static input from the Regional Interlocks Inputs tab (New Business ARR Targets)'); return; }
    if (key === 'none') { toast('No deals behind this number'); return; }
    $('drillTitle').textContent = title || 'Deals';
    $('drillMeta').textContent = 'Loading…';
    $('drillTable').innerHTML = '';
    $('drillSearch').value = '';
    $('drill').classList.add('on');
    drillState = null;
    fetch('/api/drill/' + encodeURIComponent(key)).then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.error || 'HTTP ' + r.status); return j; }); })
      .then(function (data) { drillState = { key: key, title: title, data: data }; renderDrill(); })
      .catch(function (e) { $('drillMeta').textContent = 'Could not load the list: ' + e.message; });
  }
  function drillCsv() {
    var d = drillState; if (!d) return;
    var q = function (v) { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var csv = [d.data.columns.map(q).join(',')].concat(d.data.rows.map(function (r) { return r.cells.map(q).join(','); })).join('\n');
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    a.download = (d.title || 'deals').replace(/[^\w\- ]+/g, ' ').replace(/\s+/g, ' ').trim() + '.csv';
    a.click();
  }

  function wire() {
    $('prev').onclick = function () { go(-1); };
    $('next').onclick = function () { go(1); };
    $('downloadBtn').onclick = openPicker;
    $('refreshBtn').onclick = function () { toast('Refreshing from HubSpot… this can take a minute'); load(true); };
    $('pickerClose').onclick = function () { $('picker').classList.remove('on'); };
    $('picker').addEventListener('click', function (e) { if (e.target.id === 'picker') $('picker').classList.remove('on'); });
    $('pickAll').onclick = function (e) { e.preventDefault(); state.slides.forEach(function (s) { state.pick.add(s.id); }); renderPicker(); };
    $('pickNone').onclick = function (e) { e.preventDefault(); state.pick.clear(); renderPicker(); };
    $('pickList').addEventListener('change', function (e) {
      var t = e.target;
      if (t.dataset.pick) { t.checked ? state.pick.add(t.dataset.pick) : state.pick.delete(t.dataset.pick); }
      if (t.dataset.sec) state.slides.forEach(function (s) { if (s.section === t.dataset.sec) { t.checked ? state.pick.add(s.id) : state.pick.delete(s.id); } });
      renderPicker();
    });
    $('pickGo').onclick = exportPptx;
    $('frame').addEventListener('click', function (e) {
      var dr = e.target.closest('[data-drill]');
      if (dr) openDrill(dr.getAttribute('data-drill'), dr.getAttribute('data-drill-title'));
    });
    $('drillClose').onclick = function () { $('drill').classList.remove('on'); };
    $('drill').addEventListener('click', function (e) { if (e.target.id === 'drill') $('drill').classList.remove('on'); });
    $('drillSearch').addEventListener('input', renderDrill);
    $('drillCsv').onclick = drillCsv;
    $('thumbs').addEventListener('click', function (e) { var t = e.target.closest('.thumb'); if (t) { state.idx = Number(t.dataset.i); renderMain(); } });
    window.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') { $('drill').classList.remove('on'); $('picker').classList.remove('on'); }
      if (e.target.isContentEditable || /INPUT|TEXTAREA/.test(e.target.tagName)) return;
      if (e.key === 'ArrowRight' || e.key === 'PageDown') go(1);
      if (e.key === 'ArrowLeft' || e.key === 'PageUp') go(-1);
    });
    window.addEventListener('resize', fit);
  }

  async function load(refresh) {
    try {
      var res = await fetch('/api/retro' + (refresh ? '?refresh=1' : ''), { headers: { Accept: 'application/json' } });
      var data = await res.json();
      if (!res.ok) throw new Error(data.error || ('HTTP ' + res.status));
      var first = !state.data;
      state.data = data;
      try { state.insights = await api('GET', '/api/insights'); } catch (e) { toast('Insights could not be loaded'); }
      var keepId = state.slides[state.idx] && state.slides[state.idx].id;
      state.slides = QRDeck.buildDeck(data);
      if (first) state.slides.forEach(function (s) { state.pick.add(s.id); });
      var ki = state.slides.findIndex(function (s) { return s.id === keepId; }); state.idx = ki > -1 ? ki : 0;
      $('periodChip').textContent = data.periods.compare.short + ' vs ' + data.periods.review.short + ' · ' + data.periods.combined.label;
      $('asOf').textContent = (data.mock ? 'MOCK DATA · ' : 'LIVE DATA · ') + 'AS OF ' + new Date(data.generatedAt).toUTCString().slice(5, 22).toUpperCase() + ' UTC';
      var notes = [];
      if (data.targets.source !== 'live') notes.push('Targets: Regional Interlocks not reachable, using the ' + data.targets.asOf + ' snapshot' + (data.targets.error ? ' (' + data.targets.error + ')' : ''));
      $('hint').textContent = notes.length ? notes.join(' · ') : 'Click any number to see the deals behind it · click an Insights box to type, it saves when you click away';
      var ld = $('loading'); if (ld) ld.remove();
      $('downloadBtn').disabled = false;
      renderThumbs(); renderMain();
      if (refresh) toast('Numbers refreshed');
    } catch (e) {
      var l = $('loading');
      if (l) l.outerHTML = '<div class="err"><b>Could not load the retro data.</b><br>' + QR.esc(e.message) + '</div>';
      else toast('Refresh failed: ' + e.message);
    }
  }
  wire();
  load(false);
})();
