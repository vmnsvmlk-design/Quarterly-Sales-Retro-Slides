/* Slide primitives -> HTML (on screen) and -> PptxGenJS (download).
 * Every slide is a list of elements positioned on a 1920x1080 canvas, so the
 * screen and the .pptx are drawn from exactly the same description.
 * 1920px = 13.333in, so 1px = 1/144in and a font size in px = px/2 pt. */
(function () {
  'use strict';

  var W = 1920, H = 1080, IN = 13.333 / W;
  var FONT = {
    // .pptx uses the brand fonts by name (all three are Google Fonts, so Google
    // Slides renders them; install them locally for PowerPoint to match).
    serif: { css: "'Libre Baskerville', Georgia, serif", pptx: 'Libre Baskerville' },
    sans: { css: "'Public Sans', Arial, sans-serif", pptx: 'Public Sans' },
    mono: { css: "'JetBrains Mono', Consolas, monospace", pptx: 'JetBrains Mono' }
  };
  var C = {
    navy: '#0D1B35', ink: '#1A2742', body: '#2F3D58', muted: '#6E7D98', faint: '#9BA5BB',
    line: '#C8CEDC', soft: '#E7EBF2', panel: '#F5F6F8', blue: '#9A3F1E', blueDark: '#6B2A12',
    blueSoft: '#F2DED3', green: '#1E7F4F', greenSoft: '#EAF5EF', red: '#C8102E', redMuted: '#C98A92',
    redSoft: '#FBEEF0', gold: '#D9A441', white: '#FFFFFF', grey: '#AEB7D6'
  };

  // Click-through marker: key = drill list key, 'static' = static input.
  function drillAttr(key, title) { return key ? ' data-drill="' + esc(key) + '" data-drill-title="' + esc(title || '') + '"' : ''; }

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function hex(c) { return String(c || '#000000').replace('#', '').toUpperCase(); }
  function inch(px) { return +(px * IN).toFixed(3); }
  function pt(px) { return +(px / 2).toFixed(1); }

  // ---------- HTML ----------
  function box(e, inner, extra) {
    return '<div style="position:absolute;left:' + e.x + 'px;top:' + e.y + 'px;width:' + e.w + 'px;height:' + e.h + 'px;' + (extra || '') + '">' + inner + '</div>';
  }

  function textCss(e) {
    var f = FONT[e.font || 'sans'];
    return 'font-family:' + f.css + ';font-size:' + (e.size || 24) + 'px;color:' + (e.color || C.ink) + ';font-weight:' + (e.bold ? 700 : 400) +
      ';text-align:' + (e.align || 'left') + ';line-height:' + (e.lineHeight || 1.25) + ';letter-spacing:' + (e.tracking || 0) + 'em;' +
      (e.italic ? 'font-style:italic;' : '') + (e.upper ? 'text-transform:uppercase;' : '');
  }

  function htmlText(e) {
    var v = e.valign || 'top';
    var jc = v === 'middle' ? 'center' : v === 'bottom' ? 'flex-end' : 'flex-start';
    var body = (e.runs || [{ t: e.text }]).map(function (r) {
      return '<span style="' + (r.bold ? 'font-weight:700;' : '') + (r.color ? 'color:' + r.color + ';' : '') + (r.size ? 'font-size:' + r.size + 'px;' : '') + (r.font ? 'font-family:' + FONT[r.font].css + ';' : '') + '">' + esc(r.t).replace(/\n/g, '<br>') + '</span>';
    }).join('');
    return box(e, '<div style="width:100%"' + drillAttr(e.drill, e.drillTitle) + '>' + body + '</div>', (e.drill ? 'cursor:pointer;' : '') + 'display:flex;flex-direction:column;justify-content:' + jc + ';' + textCss(e) + (e.fill ? 'background:' + e.fill + ';' : '') + (e.pad ? 'padding:' + e.pad + 'px;box-sizing:border-box;' : '') + 'overflow:hidden;');
  }

  function htmlRect(e) {
    return box(e, '', 'box-sizing:border-box;background:' + (e.fill || 'transparent') + ';' + (e.line ? 'border:' + (e.lineWidth || 2) + 'px ' + (e.dash ? 'dashed' : 'solid') + ' ' + e.line + ';' : '') + (e.radius ? 'border-radius:' + e.radius + 'px;' : ''));
  }

  function htmlTable(e) {
    var size = e.size || 20, rowH = e.rowH || 44, headH = e.headerH || rowH;
    var html = '<table style="border-collapse:collapse;width:100%;table-layout:fixed;font-family:' + FONT.sans.css + ';font-size:' + size + 'px;color:' + C.ink + '">';
    html += '<colgroup>' + e.cols.map(function (c) { return '<col style="width:' + c.w + 'px">'; }).join('') + '</colgroup>';
    if (e.cols.some(function (c) { return c.label != null; })) {
      html += '<tr style="height:' + headH + 'px;background:' + (e.headFill || C.soft) + '">' + e.cols.map(function (c) {
        return '<th style="padding:0 12px;text-align:' + (c.align || 'left') + ';font-weight:700;font-size:' + (e.headSize || size - 2) + 'px;color:' + C.navy + ';border-bottom:2px solid ' + C.line + ';' + (e.wrap ? 'white-space:normal;line-height:1.2' : 'white-space:nowrap;overflow:hidden;text-overflow:ellipsis') + '">' + esc(c.label) + '</th>';
      }).join('') + '</tr>';
    }
    var zi = 0;
    e.rows.forEach(function (row, ri) {
      // { spacer: px } = thin white gap between groups of rows.
      if (row && row.spacer) { html += '<tr style="height:' + row.spacer + 'px;background:#FFFFFF"><td colspan="' + e.cols.length + '" style="padding:0;border:0"></td></tr>'; zi = 0; return; }
      ri = zi++;
      html += '<tr style="height:' + rowH + 'px;' + (ri % 2 && e.zebra !== false ? 'background:' + C.panel + ';' : '') + '">';
      row.forEach(function (cell, ci) {
        var c = typeof cell === 'object' && cell !== null ? cell : { t: cell };
        var col = e.cols[ci];
        // e.wrap: text wraps inside the cell; c.runs: [{ t, bold, color, size }] for mixed styling.
        var inner = c.runs ? c.runs.map(function (r) { return '<span style="' + (r.bold ? 'font-weight:700;' : '') + (r.color ? 'color:' + r.color + ';' : '') + (r.size ? 'font-size:' + r.size + 'px;' : '') + '">' + esc(r.t).replace(/\n/g, '<br>') + '</span>'; }).join('') : esc(c.t);
        html += '<td' + drillAttr(c.drill, c.drillTitle || [row[0] && (row[0].t || row[0]), col.label].filter(Boolean).join(' · ')) + ' style="' + (c.drill ? 'cursor:pointer;text-decoration:underline dotted;text-underline-offset:4px;' : '') + 'padding:' + (e.wrap ? '6px 12px' : '0 12px') + ';text-align:' + (c.align || col.align || 'left') + ';border-bottom:1px solid ' + C.soft + ';' + (e.wrap ? 'white-space:normal;line-height:1.3;vertical-align:middle;' : 'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;') +
          (c.bold ? 'font-weight:700;' : '') + (c.color ? 'color:' + c.color + ';' : '') + (c.fill ? 'background:' + c.fill + ';' : '') + '">' + inner + '</td>';
      });
      html += '</tr>';
    });
    return box(e, html + '</table>', 'overflow:hidden');
  }

  function niceMax(v) {
    if (!v || v <= 0) return 1;
    var p = Math.pow(10, Math.floor(Math.log10(v)));
    var n = v / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p;
  }

  function svgWrap(e, inner) {
    return box(e, '<svg width="' + e.w + '" height="' + e.h + '" viewBox="0 0 ' + e.w + ' ' + e.h + '" style="overflow:visible;font-family:' + FONT.sans.css + '">' + inner + '</svg>');
  }

  // Bars above/below one zero line (pipeline movement).
  // Bars above/below one zero line. e.scalePos / e.scaleNeg let several
  // charts share one scale; bar.valueLabel overrides the number shown;
  // e.box = { from, to } draws a dashed outline around those bars.
  function wfGeometry(e) {
    var top = 34, bottom = e.bottom || 56, negLabel = 26;
    var plotH = e.h - top - bottom - negLabel;
    var pos = e.scalePos || Math.max.apply(null, e.bars.filter(function (b) { return b.dir > 0; }).map(function (b) { return b.value; }).concat([1]));
    var neg = e.scaleNeg != null ? e.scaleNeg : Math.max.apply(null, e.bars.filter(function (b) { return b.dir < 0; }).map(function (b) { return b.value; }).concat([0]));
    var total = pos + neg || 1;
    var n = e.bars.length, slot = e.w / n, bw = Math.min(slot * 0.62, 78);
    return { top: top, bottom: bottom, plotH: plotH, total: total, zeroY: top + plotH * (pos / total), slot: slot, bw: bw, n: n };
  }

  function htmlWaterfall(e) {
    var g = wfGeometry(e);
    var catY = e.h - g.bottom + 22;
    var out = '<line x1="0" x2="' + e.w + '" y1="' + g.zeroY + '" y2="' + g.zeroY + '" stroke="' + C.body + '" stroke-width="2"/>';
    e.bars.forEach(function (b, i) {
      var hgt = Math.max(b.value > 0 ? 2 : 0, g.plotH * (b.value / g.total));
      var x = g.slot * i + (g.slot - g.bw) / 2;
      var y = b.dir > 0 ? g.zeroY - hgt : g.zeroY;
      out += '<rect' + drillAttr(b.drill, (e.drillTitle ? e.drillTitle + ' · ' : '') + String(b.label).replace(/\n/g, ' ')) + ' style="' + (b.drill ? 'cursor:pointer' : '') + '" x="' + x + '" y="' + (b.value ? y : g.zeroY - 6) + '" width="' + g.bw + '" height="' + Math.max(hgt, b.value ? 0 : 12) + '" rx="3" fill="' + b.color + '"' + (b.value ? '' : ' fill-opacity="0"') + '/>';
      var ly = b.dir > 0 ? y - 10 : y + hgt + 22;
      out += '<text x="' + (x + g.bw / 2) + '" y="' + ly + '" text-anchor="middle" font-size="' + (e.valueSize || 21) + '" font-weight="700" fill="' + C.ink + '">' + esc(b.valueLabel != null ? b.valueLabel : b.value) + '</text>';
      String(b.label).split('\n').forEach(function (t, k) {
        out += '<text x="' + (x + g.bw / 2) + '" y="' + (catY + k * 19) + '" text-anchor="middle" font-size="' + (k ? 14 : 16) + '" fill="' + (k ? C.muted : C.body) + '">' + esc(t) + '</text>';
      });
      if (b.sub) out += '<text x="' + (x + g.bw / 2) + '" y="' + (catY + 22) + '" text-anchor="middle" font-size="14" font-weight="700" fill="' + C.green + '">' + esc(b.sub) + '</text>';
    });
    if (e.box) {
      var bx = g.slot * e.box.from + 4, bx2 = g.slot * (e.box.to + 1) - 4;
      out += '<rect x="' + bx + '" y="4" width="' + (bx2 - bx) + '" height="' + (e.h - 8) + '" rx="10" fill="none" stroke="' + C.blue + '" stroke-width="2.5" stroke-dasharray="9 6"/>';
    }
    return svgWrap(e, out);
  }

  // ---------- chart primitives ----------
  // Charts are built as rects / texts / lines in local px, then drawn as SVG
  // on screen and as plain shapes + text boxes in the .pptx, so both look
  // identical (no PowerPoint chart engine, no restyled labels or axes).
  function R(x, y, w, h, fill, o) { return Object.assign({ k: 'rect', x: x, y: y, w: w, h: h, fill: fill }, o || {}); }
  function T(x, y, t, o) { return Object.assign({ k: 'text', x: x, y: y, t: String(t), size: 17, color: C.ink, anchor: 'middle' }, o || {}); }
  function L(x1, y1, x2, y2, color, o) { return Object.assign({ k: 'line', x1: x1, y1: y1, x2: x2, y2: y2, color: color, width: 1.5 }, o || {}); }

  function primsSvg(e, prims) {
    var out = prims.map(function (p) {
      if (p.k === 'rect') return '<rect' + drillAttr(p.drill, p.drillTitle) + (p.drill ? ' style="cursor:pointer"' : '') + ' x="' + p.x + '" y="' + p.y + '" width="' + Math.max(0, p.w) + '" height="' + Math.max(0, p.h) + '" rx="' + (p.rx || 0) + '" fill="' + (p.fill || 'none') + '"' + (p.stroke ? ' stroke="' + p.stroke + '" stroke-width="' + (p.sw || 2) + '"' + (p.dash ? ' stroke-dasharray="9 6"' : '') : '') + '/>';
      if (p.k === 'wedge') {
        var da = drillAttr(p.drill, p.drillTitle) + (p.drill ? ' style="cursor:pointer"' : '');
        if (p.a1 - p.a0 >= 359.99) return '<circle' + da + ' cx="' + p.cx + '" cy="' + p.cy + '" r="' + p.r + '" fill="' + p.fill + '"/>';
        var rad = function (a) { return a * Math.PI / 180; };
        var x0 = p.cx + p.r * Math.sin(rad(p.a0)), y0 = p.cy - p.r * Math.cos(rad(p.a0)), x1 = p.cx + p.r * Math.sin(rad(p.a1)), y1 = p.cy - p.r * Math.cos(rad(p.a1));
        return '<path' + da + ' d="M' + p.cx + ' ' + p.cy + ' L' + x0 + ' ' + y0 + ' A' + p.r + ' ' + p.r + ' 0 ' + (p.a1 - p.a0 > 180 ? 1 : 0) + ' 1 ' + x1 + ' ' + y1 + ' Z" fill="' + p.fill + '" stroke="#FFFFFF" stroke-width="2"/>';
      }
      if (p.k === 'arrow') {
        var hl = Math.min(p.w, 0.5 * p.h), bt = p.h * 0.25, my = p.y + p.h / 2, xe = p.x + p.w, xb = xe - hl;
        var pts = [[p.x, my - bt], [xb, my - bt], [xb, p.y], [xe, my], [xb, p.y + p.h], [xb, my + bt], [p.x, my + bt]].map(function (q) { return q[0] + ',' + q[1]; }).join(' ');
        return '<polygon' + drillAttr(p.drill, p.drillTitle) + (p.drill ? ' style="cursor:pointer"' : '') + ' points="' + pts + '" fill="' + p.fill + '"/>';
      }
      if (p.k === 'line') return '<line x1="' + p.x1 + '" y1="' + p.y1 + '" x2="' + p.x2 + '" y2="' + p.y2 + '" stroke="' + p.color + '" stroke-width="' + p.width + '"' + (p.dash ? ' stroke-dasharray="5 4"' : '') + '/>';
      return '<text x="' + p.x + '" y="' + p.y + '" text-anchor="' + p.anchor + '" font-size="' + p.size + '"' + (p.bold ? ' font-weight="700"' : '') + ' fill="' + p.color + '">' + esc(p.t) + '</text>';
    }).join('');
    return svgWrap(e, out);
  }

  function primsPptx(slide, e, prims, pptx) {
    prims.forEach(function (p) {
      if (p.k === 'rect') {
        if (p.w <= 0 || p.h <= 0) return;
        var round = p.rx && Math.min(p.w, p.h) > 16;
        slide.addShape(round ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, { x: inch(e.x + p.x), y: inch(e.y + p.y), w: inch(p.w), h: inch(p.h),
          fill: p.fill && p.fill !== 'none' ? { color: hex(p.fill) } : { type: 'none' },
          line: p.stroke ? { color: hex(p.stroke), width: pt(p.sw || 2), dashType: p.dash ? 'dash' : 'solid' } : { type: 'none' },
          rectRadius: round ? Math.min(0.5, p.rx / Math.max(1, Math.min(p.w, p.h))) : undefined });
      } else if (p.k === 'arrow') {
        if (p.w <= 0) return;
        slide.addShape(pptx.ShapeType.rightArrow, { x: inch(e.x + p.x), y: inch(e.y + p.y), w: inch(p.w), h: inch(p.h), fill: { color: hex(p.fill) }, line: { type: 'none' } });
      } else if (p.k === 'wedge') {
        var norm = function (a) { return ((a % 360) + 360) % 360; };
        var wb = { x: inch(e.x + p.cx - p.r), y: inch(e.y + p.cy - p.r), w: inch(2 * p.r), h: inch(2 * p.r), fill: { color: hex(p.fill) }, line: { color: 'FFFFFF', width: 1 } };
        if (p.a1 - p.a0 >= 359.99) slide.addShape(pptx.ShapeType.ellipse, wb);
        else slide.addShape(pptx.ShapeType.pie, Object.assign(wb, { angleRange: [norm(p.a0 - 90), norm(p.a1 - 90)] }));
      } else if (p.k === 'line') {
        var x = Math.min(p.x1, p.x2), y = Math.min(p.y1, p.y2);
        slide.addShape(pptx.ShapeType.line, { x: inch(e.x + x), y: inch(e.y + y), w: inch(Math.abs(p.x2 - p.x1)), h: inch(Math.abs(p.y2 - p.y1)), line: { color: hex(p.color), width: pt(p.width), dashType: p.dash ? 'dash' : 'solid' } });
      } else {
        var bw = p.boxW || 420, h = p.size * 1.35, left = p.anchor === 'middle' ? p.x - bw / 2 : p.anchor === 'end' ? p.x - bw : p.x;
        slide.addText(p.t, { x: inch(e.x + left), y: inch(e.y + p.y - p.size * 1.02), w: inch(bw), h: inch(h), align: p.anchor === 'middle' ? 'center' : p.anchor === 'end' ? 'right' : 'left', valign: 'top',
          fontFace: FONT.sans.pptx, fontSize: pt(p.size), bold: !!p.bold, color: hex(p.color), margin: 0 });
      }
    });
  }

  // Grouped or stacked bar chart, vertical (bar) or horizontal (barH).
  function barsPrims(e) {
    var horiz = e.kind === 'barH';
    var cats = e.categories, ser = e.series, P = [];
    var legendH = e.legend === false ? 0 : 36;
    var padL = horiz ? (e.catWidth || 220) : 10, padB = horiz ? 10 : 44, padT = 26 + legendH, padR = horiz ? 60 : 10;
    var pw = e.w - padL - padR, ph = e.h - padT - padB;
    var totals = cats.map(function (_, ci) { return ser.reduce(function (a, s) { return a + (s.values[ci] || 0); }, 0); });
    var maxV = e.stacked ? Math.max.apply(null, totals.concat([1])) : Math.max.apply(null, ser.map(function (s) { return Math.max.apply(null, s.values.concat([0])); }).concat([1]));
    if (e.max) maxV = e.max;
    if (legendH) {
      var lx = horiz ? padL : 0;
      ser.forEach(function (s) { P.push(R(lx, 4, 16, 16, s.color, { rx: 3 }), T(lx + 24, 18, s.name, { anchor: 'start', size: 18, color: C.body, boxW: s.name.length * 11 + 20 })); lx += 40 + s.name.length * 9; });
    }
    var n = cats.length, slot = (horiz ? ph : pw) / Math.max(n, 1);
    var groupW = slot * 0.7, bw = e.stacked ? groupW : groupW / ser.length;
    var fmt = e.fmt || function (v) { return v; };
    cats.forEach(function (cat, ci) {
      var acc = 0;
      ser.forEach(function (s, si) {
        var v = s.values[ci] || 0, len = (horiz ? pw : ph) * (v / maxV), off = (slot - groupW) / 2 + (e.stacked ? 0 : si * bw);
        var dr = { drill: e.drills && e.drills[si] && e.drills[si][ci], drillTitle: s.name + ' · ' + cat };
        if (horiz) {
          var y = padT + slot * ci + off, x = padL + (e.stacked ? pw * (acc / maxV) : 0);
          if (v) P.push(R(x, y, len, bw, s.color, dr));
          if (v && e.showValues !== false && (!e.stacked || len > 34)) P.push(T(e.stacked ? x + len / 2 : x + len + 8, y + bw / 2 + 6, fmt(v), { anchor: e.stacked ? 'middle' : 'start', bold: true, color: e.stacked ? C.white : C.ink, boxW: e.stacked ? Math.max(len, 40) : 160 }));
        } else {
          var bx = padL + slot * ci + off, by = e.stacked ? padT + ph - (ph * ((acc + v) / maxV)) : padT + ph - len;
          if (v) P.push(R(bx, by, bw, len, s.color, dr));
          if (v && e.showValues !== false && (!e.stacked || len > 22)) P.push(T(bx + bw / 2, e.stacked ? by + len / 2 + 6 : by - 6, fmt(v), { bold: true, color: e.stacked ? C.white : C.ink, boxW: Math.max(bw, 80) }));
        }
        acc += v;
      });
      if (horiz) {
        var parts = String(cat).split('\n'), cy = padT + slot * ci + slot / 2 + 6 - (parts.length - 1) * 10;
        parts.forEach(function (t, k) { P.push(T(padL - 12, cy + k * 20, t, { anchor: 'end', size: k ? 13 : 18, bold: !k, color: k ? C.muted : C.ink, boxW: padL - 16 })); });
        if (e.stacked && e.showTotals !== false) P.push(T(padL + pw * (totals[ci] / maxV) + 8, padT + slot * ci + slot / 2 + 6, fmt(totals[ci]), { anchor: 'start', bold: true, boxW: 120 }));
      } else {
        P.push(T(padL + slot * ci + slot / 2, padT + ph + 28, cat, { size: 18, color: C.body, boxW: slot }));
      }
    });
    P.push(horiz ? L(padL, padT, padL, padT + ph, C.line, { width: 2 }) : L(padL, padT + ph, padL + pw, padT + ph, C.line, { width: 2 }));
    return P;
  }
  // Pie: e.slices = [{ value, color, label, drill }], drawn clockwise from 12 o'clock.
  function piePrims(e) {
    var r = Math.min(e.w, e.h) / 2, cx = e.w / 2, cy = e.h / 2, P = [], a = 0;
    var total = e.slices.reduce(function (s, x) { return s + (x.value || 0); }, 0);
    if (!total) return [{ k: 'wedge', cx: cx, cy: cy, r: r, a0: 0, a1: 360, fill: C.soft }];
    e.slices.forEach(function (s) {
      if (!s.value) return;
      var sweep = 360 * s.value / total;
      P.push({ k: 'wedge', cx: cx, cy: cy, r: r, a0: a, a1: a + sweep, fill: s.color, drill: s.drill, drillTitle: (e.drillTitle ? e.drillTitle + ' · ' : '') + s.label });
      a += sweep;
    });
    a = 0;
    e.slices.forEach(function (s) {
      if (!s.value) return;
      var sweep = 360 * s.value / total, mid = (a + sweep / 2) * Math.PI / 180;
      if (sweep >= 36) P.push(T(cx + Math.sin(mid) * r * 0.6, cy - Math.cos(mid) * r * 0.6 + 7, Math.round(s.value / total * 100) + '%', { size: 18, bold: true, color: C.white, boxW: 80 }));
      a += sweep;
    });
    return P;
  }
  function htmlBars(e) { return primsSvg(e, barsPrims(e)); }

  function htmlTrapezoid(e) {
    var i = e.inset || 0;
    var pts = '0,0 ' + e.w + ',0 ' + (e.w - i) + ',' + e.h + ' ' + i + ',' + e.h;
    return svgWrap(e, '<polygon points="' + pts + '" fill="' + e.fill + '"/>');
  }

  function htmlLine(e) {
    var x = Math.min(e.x1, e.x2), y = Math.min(e.y1, e.y2), w = Math.max(Math.abs(e.x2 - e.x1), 1), h = Math.max(Math.abs(e.y2 - e.y1), 1);
    return box({ x: x, y: y, w: w, h: h }, '<svg width="' + w + '" height="' + h + '" style="overflow:visible"><line x1="' + (e.x1 - x) + '" y1="' + (e.y1 - y) + '" x2="' + (e.x2 - x) + '" y2="' + (e.y2 - y) + '" stroke="' + e.color + '" stroke-width="' + (e.width || 2) + '"' + (e.dash ? ' stroke-dasharray="8 6"' : '') + '/></svg>');
  }

  // Floating waterfall (running total), Gong style.
  // e.steps = [{ label, value, total, valueLabel, sub, drill }], e.scaleMax
  // shared across charts, e.colors = { total, up, down }, e.negColor for the
  // value text of removals, e.labels (category labels under the bars),
  // e.bracket = { from, to, label } (line + total above those bars),
  // e.box = { from, to } (dashed highlight).
  function flowPrims(e) {
    var col = Object.assign({ total: C.navy, up: C.blue, down: C.redMuted }, e.colors || {});
    var top = e.bracket ? 74 : 30, bottom = e.labels ? 54 : 6, plotH = e.h - top - bottom;
    var n = e.steps.length, slot = e.w / n, bw = Math.min(slot * 0.62, 84);
    // e.scaleMin (<= 0) lets the running total dip below zero inside the chart.
    var lo = Math.min(0, e.scaleMin || 0);
    var y = function (v) { return top + plotH * (e.scaleMax - v) / (e.scaleMax - lo); };
    var run = 0, bars = [], P = [L(0, y(0), e.w, y(0), C.line)];
    e.steps.forEach(function (st, i) {
      var x = slot * i + (slot - bw) / 2;
      if (st.value == null) { bars.push({ x: x, empty: true, run: run }); return; }
      var a, b, color, label, neg = false;
      if (st.total) { a = 0; b = st.value; run = st.value; color = col.total; label = e.fmt(st.value); }
      else { a = run; b = run + st.value; run = b; neg = st.value < 0; color = neg ? col.down : col.up; label = (neg ? '−' : '+') + e.fmt(Math.abs(st.value)); }
      bars.push({ x: x, top: y(Math.max(a, b)), bottom: y(Math.min(a, b)), color: color, label: st.valueLabel != null ? st.valueLabel : label, run: run, neg: neg, st: st });
    });
    bars.forEach(function (b, i) {
      var next = bars[i + 1];
      if (!b.empty) {
        P.push(R(b.x, b.top, bw, Math.max(2, b.bottom - b.top), b.color, { rx: 3, drill: b.st.drill || e.drill, drillTitle: (e.drillTitle ? e.drillTitle + ' · ' : '') + (b.st.label || '') }));
        P.push(T(b.x + bw / 2, b.top - 8, b.label, { size: e.valueSize || 16, bold: true, color: b.neg && e.negColor ? e.negColor : C.ink, boxW: slot + 40 }));
      } else P.push(T(b.x + bw / 2, y(b.run) - 7, '-', { size: 14, color: C.faint, boxW: slot }));
      if (next) P.push(L(b.x + bw, y(b.run), next.x, y(b.run), C.grey, { dash: true }));
      if (e.labels && b.st) {
        String(b.st.label).split('\n').forEach(function (t, k) { P.push(T(b.x + bw / 2, y(lo) + 24 + k * 19, t, { size: k ? 14 : 16, color: k ? C.muted : C.body, boxW: slot })); });
        if (b.st.sub) P.push(T(b.x + bw / 2, y(lo) + 44, b.st.sub, { size: 14, bold: true, color: C.blueDark, boxW: slot + 30 }));
      }
    });
    if (e.bracket) {
      var bs = bars.slice(e.bracket.from, e.bracket.to + 1).filter(function (b) { return !b.empty; });
      if (bs.length) {
        var bx1 = bars[e.bracket.from].x, bx2 = bars[e.bracket.to].x + bw;
        var by = Math.min.apply(null, bs.map(function (b) { return b.top; })) - 34;
        P.push(L(bx1, by, bx2, by, C.navy, { width: 2 }), L(bx1, by, bx1, by + 10, C.navy, { width: 2 }), L(bx2, by, bx2, by + 10, C.navy, { width: 2 }));
        P.push(T((bx1 + bx2) / 2, by - 8, e.bracket.label, { size: 16, bold: true, color: C.navy, boxW: bx2 - bx1 + 80 }));
      }
    }
    if (e.box) {
      var x1 = slot * e.box.from + 4, x2 = slot * (e.box.to + 1) - 4;
      P.push(R(x1, 2, x2 - x1, e.h - 4, 'none', { stroke: C.blue, sw: 2.5, dash: true, rx: 10 }));
    }
    return P;
  }
  function htmlFlow(e) { return primsSvg(e, flowPrims(e)); }
  function pFlow(slide, e, pptx) { primsPptx(slide, e, flowPrims(e), pptx); }

  // Editable insights box. On screen the app turns [data-insights] into an
  // editable field; in thumbnails and the .pptx it is plain text.
  function htmlInsights(e) {
    var body = e.text ? esc(e.text).replace(/\n/g, '<br>') : '<span class="ph" style="color:' + C.faint + '">' + esc(e.placeholder || 'Add insights…') + '</span>';
    return box(e, '<div style="font-family:' + FONT.mono.css + ';font-size:15px;font-weight:700;letter-spacing:.12em;color:' + C.blue + ';text-transform:uppercase;margin-bottom:6px">' + esc(e.label || 'Insights') + '</div>' +
      '<div data-insights="' + esc(e.key) + '" style="font-family:' + FONT.sans.css + ';font-size:' + (e.size || 19) + 'px;line-height:1.35;color:' + C.ink + ';height:calc(100% - 28px);overflow:hidden;white-space:normal">' + body + '</div>',
      'box-sizing:border-box;padding:12px 20px;background:' + C.white + ';border:2px solid ' + C.blueSoft + ';border-radius:12px;');
  }

  function htmlLink(e) {
    return box(e, '<span data-goto="' + esc(e.target) + '" style="cursor:pointer;color:' + C.blue + ';font-family:' + FONT.sans.css + ';font-size:' + (e.size || 17) + 'px;font-weight:700;text-decoration:underline">' + esc(e.text) + '</span>', 'display:flex;align-items:center;justify-content:' + (e.align === 'right' ? 'flex-end' : 'flex-start') + ';');
  }

  function toHtml(elements) {
    return elements.filter(function (e) { return !e.pptxOnly; }).map(function (e) {
      switch (e.type) {
        case 'text': return htmlText(e);
        case 'rect': return htmlRect(e);
        case 'image': return box(e, '<img src="' + e.src + '" style="width:100%;height:100%;object-fit:contain;object-position:' + (e.position || 'left center') + '">');
        case 'table': return htmlTable(e);
        case 'bars': return htmlBars(e);
        case 'waterfall': return htmlWaterfall(e);
        case 'trapezoid': return htmlTrapezoid(e);
        case 'line': return htmlLine(e);
        case 'insights': return htmlInsights(e);
        case 'flow': return htmlFlow(e);
        case 'pie': return primsSvg(e, piePrims(e));
        case 'prims': return primsSvg(e, e.prims);
        case 'link': return htmlLink(e);
        default: return '';
      }
    }).join('');
  }

  // ---------- PPTX ----------
  function pText(slide, e) {
    // Split runs on newlines into explicit line breaks (a raw "\n" inside a run
    // makes PowerPoint start an odd empty paragraph).
    var runs = [];
    (e.runs || [{ t: e.text }]).forEach(function (r) {
      var opts = { bold: r.bold != null ? r.bold : !!e.bold, color: hex(r.color || e.color || C.ink), fontSize: pt(r.size || e.size || 24), fontFace: FONT[r.font || e.font || 'sans'].pptx };
      var txt = String(r.t == null ? '' : r.t); if (e.upper) txt = txt.toUpperCase();
      txt.split('\n').forEach(function (piece, k, arr) {
        if (k > 0 && runs.length) runs[runs.length - 1].options.breakLine = true;
        if (piece !== '' || k === arr.length - 1) runs.push({ text: piece, options: Object.assign({}, opts) });
        else runs.push({ text: '', options: Object.assign({}, opts) });
      });
    });
    slide.addText(runs, {
      x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), align: e.align || 'left', valign: e.valign || 'top', margin: e.pad ? pt(e.pad) : 0,
      fill: e.fill ? { color: hex(e.fill) } : undefined, charSpacing: e.tracking ? Math.round(e.tracking * pt(e.size || 24)) : 0,
      lineSpacingMultiple: e.lineHeight || 1.1, italic: !!e.italic, isTextBox: true, fit: 'shrink',
      shrinkText: true
    });
  }

  function pRect(slide, e, pptx) {
    slide.addShape(e.radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, {
      x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h),
      fill: e.fill ? { color: hex(e.fill) } : { type: 'none' },
      line: e.line ? { color: hex(e.line), width: pt(e.lineWidth || 2), dashType: e.dash ? 'dash' : 'solid' } : { type: 'none' },
      rectRadius: e.radius ? Math.min(0.5, e.radius / Math.min(e.w, e.h)) : undefined
    });
  }

  function pTable(slide, e) {
    var size = pt(e.size || 20), rows = [];
    if (e.cols.some(function (c) { return c.label != null; })) {
      rows.push(e.cols.map(function (c) { return { text: String(c.label || ''), options: { bold: true, color: hex(C.navy), fill: { color: hex(e.headFill || C.soft) }, align: c.align || 'left', fontSize: pt(e.headSize || (e.size || 20) - 2) } }; }));
    }
    var zi = 0, heights = [];
    e.rows.forEach(function (row, ri) {
      if (row && row.spacer) { heights.push(inch(row.spacer)); zi = 0; rows.push(e.cols.map(function () { return { text: '', options: { fill: { color: 'FFFFFF' }, fontSize: 1, border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }; })); return; }
      ri = zi++; heights.push(inch(e.rowH || 44));
      rows.push(row.map(function (cell, ci) {
        var c = typeof cell === 'object' && cell !== null ? cell : { t: cell };
        var fill = c.fill || (ri % 2 && e.zebra !== false ? C.panel : C.white);
        var opts = { bold: !!c.bold, color: hex(c.color || C.ink), fill: { color: hex(fill) }, align: c.align || e.cols[ci].align || 'left' };
        if (c.runs) {
          var parts = [];
          c.runs.forEach(function (r) {
            String(r.t).split('\n').forEach(function (piece, k) {
              if (k > 0 && parts.length) parts[parts.length - 1].options.breakLine = true;
              parts.push({ text: piece, options: { bold: !!r.bold, color: hex(r.color || c.color || C.ink), fontSize: r.size ? pt(r.size) : pt(e.size || 20) } });
            });
          });
          return { text: parts, options: opts };
        }
        return { text: String(c.t == null ? '' : c.t), options: opts };
      }));
    });
    slide.addTable(rows, {
      x: inch(e.x), y: inch(e.y), w: inch(e.w), colW: e.cols.map(function (c) { return inch(c.w); }),
      rowH: (function () { var hasHead = e.cols.some(function (c) { return c.label != null; }); return hasHead ? [inch(e.headerH || e.rowH || 44)].concat(heights) : heights; })(), fontFace: FONT.sans.pptx, fontSize: size, margin: e.wrap ? [0.04, 0.08, 0.04, 0.08] : [0, 0.06, 0, 0.06], valign: 'middle',
      border: { type: 'solid', pt: 0.5, color: hex(C.soft) }, autoPage: false
    });
  }

  function pBars(slide, e, pptx) { primsPptx(slide, e, barsPrims(e), pptx); }

  // Native chart: one stacked series per bar colour, won/lost as negatives.
  function pWaterfall(slide, e, pptx) {
    var labels = e.bars.map(function (b) { var l = String(b.label).replace(/\n/g, ' '); return b.pptxLabel || (b.sub ? l + ' (' + b.sub + ')' : l); });
    var colors = [], series = [];
    e.bars.forEach(function (b, i) {
      var s = series.filter(function (x) { return x.key === b.color; })[0];
      if (!s) { s = { key: b.color, name: b.seriesName || b.label, values: e.bars.map(function () { return 0; }) }; series.push(s); colors.push(hex(b.color)); }
      s.values[i] = b.dir < 0 ? -b.value : b.value;
    });
    var g = wfGeometry(e);
    var maxV = e.scalePos || Math.max.apply(null, e.bars.filter(function (b) { return b.dir > 0; }).map(function (b) { return b.value; }).concat([1]));
    var minV = -(e.scaleNeg != null ? e.scaleNeg : Math.max.apply(null, e.bars.filter(function (b) { return b.dir < 0; }).map(function (b) { return b.value; }).concat([0])));
    slide.addChart(pptx.ChartType.bar, series.map(function (s) { return { name: s.name, labels: labels, values: s.values }; }), {
      x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 50,
      chartColors: colors, showValue: true, dataLabelFormatCode: '#,##0;#,##0;;', dataLabelFontSize: 9, dataLabelColor: 'FFFFFF',
      catAxisLabelFontSize: 8, catAxisLabelColor: hex(C.body), catAxisLabelPos: 'low',
      valAxisHidden: true, valGridLine: { style: 'none' }, valAxisMaxVal: maxV, valAxisMinVal: minV, showLegend: false
    });
    if (e.box) slide.addShape(pptx.ShapeType.roundRect, { x: inch(e.x + g.slot * e.box.from), y: inch(e.y), w: inch(g.slot * (e.box.to - e.box.from + 1)), h: inch(e.h), fill: { type: 'none' }, line: { color: hex(C.blue), width: 1.25, dashType: 'dash' }, rectRadius: 0.1 });
  }

  function pInsights(slide, e, pptx) {
    slide.addShape(pptx.ShapeType.roundRect, { x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), fill: { color: 'FFFFFF' }, line: { color: hex(C.blueSoft), width: 1 }, rectRadius: 0.08 });
    slide.addText(String(e.label || 'Insights').toUpperCase(), { x: inch(e.x + 20), y: inch(e.y + 10), w: inch(e.w - 40), h: inch(22), fontFace: FONT.mono.pptx, fontSize: pt(15), bold: true, color: hex(C.blue), margin: 0, charSpacing: 2 });
    slide.addText(e.text || ' ', { x: inch(e.x + 20), y: inch(e.y + 36), w: inch(e.w - 40), h: inch(e.h - 44), fontFace: FONT.sans.pptx, fontSize: pt(e.size || 19), color: hex(C.ink), valign: 'top', margin: 0, fit: 'shrink' });
  }

  function pTrapezoid(slide, e, pptx) {
    slide.addShape(pptx.ShapeType.trapezoid, {
      x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), flipV: true,
      fill: { color: hex(e.fill) }, line: { type: 'none' },
      adjust: e.w ? Math.min(0.5, (e.inset || 0) / e.h) : 0.25
    });
  }

  function toPptx(slide, elements, pptx) {
    elements.filter(function (e) { return !e.screenOnly; }).forEach(function (e) {
      switch (e.type) {
        case 'text': return pText(slide, e);
        case 'rect': return pRect(slide, e, pptx);
        case 'image': return slide.addImage({ path: e.src, x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), sizing: { type: 'contain', w: inch(e.w), h: inch(e.h) } });
        case 'table': return pTable(slide, e);
        case 'bars': return pBars(slide, e, pptx);
        case 'waterfall': return pWaterfall(slide, e, pptx);
        case 'trapezoid': return pTrapezoid(slide, e, pptx);
        case 'insights': return pInsights(slide, e, pptx);
        case 'flow': return pFlow(slide, e, pptx);
        case 'pie': return primsPptx(slide, e, piePrims(e), pptx);
        case 'prims': return primsPptx(slide, e, e.prims, pptx);
        case 'link': return slide.addText([{ text: e.text, options: e.slideNumber ? { hyperlink: { slide: e.slideNumber }, color: hex(C.blue), bold: true, underline: { style: 'sng' } } : { color: hex(C.blue), bold: true } }], { x: inch(e.x), y: inch(e.y), w: inch(e.w), h: inch(e.h), fontFace: FONT.sans.pptx, fontSize: pt(e.size || 17), align: e.align || 'left', valign: 'middle', margin: 0 });
        case 'line': return slide.addShape(pptx.ShapeType.line, { x: inch(Math.min(e.x1, e.x2)), y: inch(Math.min(e.y1, e.y2)), w: inch(Math.abs(e.x2 - e.x1)), h: inch(Math.abs(e.y2 - e.y1)), line: { color: hex(e.color), width: pt(e.width || 2), dashType: e.dash ? 'dash' : 'solid' } });
        default: return null;
      }
    });
  }

  window.QR = { W: W, H: H, C: C, toHtml: toHtml, toPptx: toPptx, niceMax: niceMax, esc: esc };
})();
