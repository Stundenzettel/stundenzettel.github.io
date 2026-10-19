/*
 * Stundenzettel als PDF (A4 quer) – ohne externe Bibliotheken.
 * Verwendet die PDF-Standardschriften Helvetica und Helvetica-Bold (WinAnsi-Kodierung).
 */
(function (global) {
  'use strict';

  // Zeichenbreiten in 1/1000 em für die Codes 32–255 (WinAnsi), aus Helvetica / Helvetica-Bold.
  const REGULAR = [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,500,744,0,222,556,333,1000,556,556,333,1000,667,333,1000,0,611,0,0,222,222,333,333,350,556,1000,333,1000,500,333,944,0,500,667,278,333,556,556,556,556,260,556,333,737,370,556,584,333,737,333,400,549,333,333,333,576,537,278,333,333,365,556,834,834,834,611,667,667,667,667,667,667,1000,722,667,667,667,667,278,278,278,278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,500,556,556,556,556,278,278,278,278,556,556,556,556,556,556,556,549,611,556,556,556,556,500,556,500];
  const BOLD = [278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,500,744,0,278,556,500,1000,556,556,333,1000,667,333,1000,0,611,0,0,278,278,500,500,350,556,1000,333,1000,556,333,944,0,500,667,278,333,556,556,556,556,280,556,333,737,370,556,584,333,737,333,400,549,333,333,333,576,556,278,333,333,365,556,834,834,834,611,722,722,722,722,722,722,1000,722,667,667,667,667,278,278,278,278,722,722,778,778,778,778,778,584,778,722,722,722,722,667,667,611,556,556,556,556,556,556,889,556,556,556,556,556,278,278,278,278,611,611,611,611,611,611,611,549,611,611,611,611,611,556,611,556];

  const PAGE_W = 842, PAGE_H = 595;
  const FGF_GRAY = 0.75; // Grau des FGF-Bereichs in der Vorlage (#C0C0C0)
  const WINANSI_EXTRA = {
    0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88,
    0x2030: 0x89, 0x0160: 0x8A, 0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E, 0x2018: 0x91, 0x2019: 0x92, 0x201C: 0x93,
    0x201D: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B,
    0x0153: 0x9C, 0x017E: 0x9E, 0x0178: 0x9F
  };

  function encode(str) {
    const out = [];
    for (const ch of String(str == null ? '' : str)) {
      const c = ch.codePointAt(0);
      if ((c >= 0x20 && c < 0x7F) || (c >= 0xA0 && c <= 0xFF)) out.push(c);
      else if (WINANSI_EXTRA[c]) out.push(WINANSI_EXTRA[c]);
      else if (c === 0x09 || c === 0x0A || c === 0x0D) out.push(0x20);
      else out.push(0x3F);
    }
    return out;
  }

  function width(str, size, bold) {
    const table = bold ? BOLD : REGULAR;
    let w = 0;
    for (const c of encode(str)) w += table[c - 32] || 500;
    return w * size / 1000;
  }

  function fit(str, size, bold, maxW) {
    str = String(str == null ? '' : str);
    if (width(str, size, bold) <= maxW) return str;
    while (str.length && width(str + '…', size, bold) > maxW) str = str.slice(0, -1);
    return str.trimEnd() + '…';
  }

  function wrap(str, size, bold, maxW, maxLines) {
    const words = String(str).split(/\s+/).filter(Boolean);
    const lines = [];
    let cur = '';
    for (let i = 0; i < words.length; i++) {
      const next = cur ? cur + ' ' + words[i] : words[i];
      if (!cur || width(next, size, bold) <= maxW) { cur = next; continue; }
      lines.push(cur);
      cur = words[i];
      if (lines.length === maxLines - 1) { cur = words.slice(i).join(' '); break; }
    }
    if (cur) lines.push(cur);
    return lines.map(l => fit(l, size, bold, maxW));
  }

  const n = v => String(Math.round(v * 100) / 100);

  function escapeBytes(codes) {
    const out = [];
    for (const c of codes) {
      if (c === 0x28 || c === 0x29 || c === 0x5C) out.push(0x5C);
      out.push(c);
    }
    return new Uint8Array(out);
  }

  // Zeichenfläche mit Koordinaten von oben links (in Punkt)
  function page() {
    const ops = [];
    return {
      ops,
      text(x, top, str, size, o) {
        o = o || {};
        const bold = !!o.bold;
        let tx = x;
        if (o.align === 'center') tx -= width(str, size, bold) / 2;
        if (o.align === 'right') tx -= width(str, size, bold);
        ops.push('BT /' + (bold ? 'F2' : 'F1') + ' ' + n(size) + ' Tf ' + n(o.gray || 0) + ' g ' + n(tx) + ' ' + n(PAGE_H - top) + ' Td (');
        ops.push(escapeBytes(encode(str)));
        ops.push(') Tj ET\n');
      },
      line(x1, t1, x2, t2, lw, gray) {
        ops.push(n(lw || 0.6) + ' w ' + n(gray || 0) + ' G ' + n(x1) + ' ' + n(PAGE_H - t1) + ' m ' + n(x2) + ' ' + n(PAGE_H - t2) + ' l S\n');
      },
      fill(x, top, w, h, gray) {
        ops.push(n(gray) + ' g ' + n(x) + ' ' + n(PAGE_H - top - h) + ' ' + n(w) + ' ' + n(h) + ' re f\n');
      },
      box(x, top, w, h, lw) {
        ops.push(n(lw || 0.6) + ' w 0 G ' + n(x) + ' ' + n(PAGE_H - top - h) + ' ' + n(w) + ' ' + n(h) + ' re S\n');
      },
      polyline(points, lw) {
        ops.push(n(lw) + ' w 0 G 1 J 1 j ' + points.map((p, i) => n(p[0]) + ' ' + n(PAGE_H - p[1]) + (i ? ' l' : ' m')).join(' ') + ' S 0 J 0 j\n');
      },
      image(name, x, top, w, h) {
        ops.push('q ' + n(w) + ' 0 0 ' + n(h) + ' ' + n(x) + ' ' + n(PAGE_H - top - h) + ' cm /' + name + ' Do Q\n');
      }
    };
  }

  function checkbox(p, cx, cy, checked) {
    p.box(cx - 5, cy - 5, 10, 10, 0.9);
    if (checked) p.polyline([[cx - 3.2, cy + 0.2], [cx - 0.8, cy + 2.8], [cx + 3.6, cy - 3.4]], 1.4);
  }

  function concat(parts) {
    const enc = new TextEncoder();
    const bytes = parts.map(x => (typeof x === 'string' ? enc.encode(x) : x));
    const out = new Uint8Array(bytes.reduce((s, b) => s + b.length, 0));
    let o = 0;
    bytes.forEach(b => { out.set(b, o); o += b.length; });
    return out;
  }

  function assemble(content, image, title) {
    const objects = [
      ['<< /Type /Catalog /Pages 2 0 R >>'],
      ['<< /Type /Pages /Kids [3 0 R] /Count 1 >>'],
      ['<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + PAGE_W + ' ' + PAGE_H + '] /Resources << /Font << /F1 4 0 R /F2 5 0 R >>' +
        (image ? ' /XObject << /Im1 7 0 R >>' : '') + ' >> /Contents 6 0 R >>'],
      ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'],
      ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>'],
      ['<< /Length ' + content.length + ' >>\nstream\n', content, '\nendstream']
    ];
    if (image) {
      objects.push(['<< /Type /XObject /Subtype /Image /Width ' + image.w + ' /Height ' + image.h +
        ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + image.jpeg.length + ' >>\nstream\n', image.jpeg, '\nendstream']);
    }
    objects.push(['<< /Title (', escapeBytes(encode(title || 'Stundenzettel')), ') /Producer (Stundenzettel-App) >>']);
    const parts = [new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34, 0x0A, 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A])];
    let offset = parts[0].length;
    const offsets = [];
    objects.forEach((obj, i) => {
      const bytes = concat([(i + 1) + ' 0 obj\n'].concat(obj, ['\nendobj\n']));
      offsets.push(offset);
      parts.push(bytes);
      offset += bytes.length;
    });
    const xref = 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n' +
      offsets.map(o => String(o).padStart(10, '0') + ' 00000 n \n').join('') +
      'trailer\n<< /Size ' + (objects.length + 1) + ' /Root 1 0 R /Info ' + objects.length + ' 0 R >>\nstartxref\n' + offset + '\n%%EOF\n';
    parts.push(xref);
    return new Blob([concat(parts)], { type: 'application/pdf' });
  }

  // d: { title, company, address[], production, productionNo, name, position, kwLabel, weekLabel,
  //      days: [{ weekday, date, start, end, pStart, pEnd, pause, hours, note, catering }], total,
  //      cateringColumn, signDate, signature: { jpeg: Uint8Array, w, h } | null }
  function build(d) {
    const p = page();
    const L = 32, R = PAGE_W - 32;

    p.text(L, 46, fit(d.company || '', 18, true, 380), 18, { bold: true });
    (d.address || []).forEach((a, i) => p.text(L, 60 + i * 11, a, 8, { gray: 0.25 }));
    p.text(R, 44, 'Stundenzettel', 18, { bold: true, align: 'right' });
    p.text(R, 60, d.kwLabel || '', 10, { align: 'right', gray: 0.35 });

    const field = (label, value, x, top, labelW, valueW, size) => {
      p.text(x, top, label, 9, { bold: true });
      p.text(x + labelW, top, fit(value || '', size, false, valueW), size);
      p.line(x + labelW, top + 4, x + labelW + valueW, top + 4, 0.5, 0.6);
    };
    field('PRODUKTION:', d.production, L, 100, 92, 320, 11);
    field('PRODUKTIONSNUMMER:', d.productionNo, 470, 100, 124, R - 470 - 124, 11);
    field('NAME:', d.name, L, 128, 92, 320, 13);
    field('POSITION:', d.position, 470, 128, 124, R - 470 - 124, 13);
    field('WOCHE:', d.weekLabel, L, 156, 92, R - L - 92, 11);

    // Spalten wie im Formular. Die vier grauen Spalten sind der „Bereich für FGF“ und bleiben leer.
    const cols = [['Wochentag', 70], ['Datum', 62], ['Arbeits-\nbeginn', 52], ['Arbeits-\nende', 52], ['Beginn\nPause', 50],
      ['Ende\nPause', 50], ['abzgl.\nPause', 46], ['Zeit\ngesamt', 52], ['', 34], ['', 34], ['', 34], ['', 34], ['Bemerkungen / Notiz', 0]];
    const FGF = 8, NOTE = 12, CAT = 13;
    if (d.cateringColumn) cols.push(['Teilnahme\nSetcatering', 56]);
    const total = R - L;
    cols[NOTE][1] = total - cols.reduce((s, c) => s + c[1], 0);
    const xs = [];
    cols.reduce((x, c) => { xs.push(x); return x + c[1]; }, L);
    const fgfW = xs[NOTE] - xs[FGF];

    const top0 = 196, headH = 30, rowH = 30;
    const bottom = top0 + headH + d.days.length * rowH;
    p.fill(xs[FGF], top0 - 18, fgfW, 18, FGF_GRAY);
    p.text(xs[FGF] + 5, top0 - 6, 'Bereich für FGF', 8);
    p.text(xs[FGF] - 6, top0 - 6, 'bei Abweichungen korrigieren!', 8, { align: 'right' });
    p.fill(L, top0, total, headH, 0.85);
    d.days.forEach((day, r) => {
      const top = top0 + headH + r * rowH;
      p.fill(xs[0], top, cols[0][1], rowH, 0.9);
      p.fill(xs[1], top, cols[1][1], rowH, 0.96);
      p.fill(xs[7], top, cols[7][1], rowH, 0.96);
      p.fill(xs[FGF], top, fgfW, rowH, FGF_GRAY);
    });
    cols.forEach((c, i) => {
      if (!c[0]) return;
      const lines = c[0].split('\n');
      lines.forEach((ln, j) => p.text(xs[i] + c[1] / 2, top0 + headH / 2 + 3 + (j - (lines.length - 1) / 2) * 10, ln, 8, { bold: true, align: 'center' }));
    });
    d.days.forEach((day, r) => {
      const top = top0 + headH + r * rowH;
      const mid = top + rowH / 2 + 3.5;
      p.text(xs[0] + 6, mid, day.weekday, 10);
      p.text(xs[1] + cols[1][1] / 2, mid, day.date, 10, { align: 'center' });
      [day.start, day.end, day.pStart, day.pEnd, day.pause, day.hours].forEach((v, j) => {
        if (v) p.text(xs[2 + j] + cols[2 + j][1] / 2, mid, v, 10, { align: 'center', bold: j === 5 });
      });
      if (day.note) {
        // Passt die Bemerkung nicht in zwei Zeilen, wird sie kleiner und dreizeilig gesetzt.
        const maxW = cols[NOTE][1] - 10;
        const small = wrap(day.note, 8.5, false, maxW, 99).length > 2;
        const size = small ? 7.5 : 8.5, lh = small ? 8.5 : 10;
        const lines = wrap(day.note, size, false, maxW, small ? 3 : 2);
        lines.forEach((ln, j) => p.text(xs[NOTE] + 5, top + rowH / 2 + size * 0.35 + (j - (lines.length - 1) / 2) * lh, ln, size));
      }
      if (d.cateringColumn) checkbox(p, xs[CAT] + cols[CAT][1] / 2, top + rowH / 2, day.catering);
    });
    p.box(L, top0, total, bottom - top0, 0.8);
    xs.slice(1).forEach(x => p.line(x, top0, x, bottom, 0.5));
    for (let r = 0; r < d.days.length; r++) p.line(L, top0 + headH + r * rowH, R, top0 + headH + r * rowH, 0.5);

    p.fill(xs[7], bottom, cols[7][1], rowH, 0.96);
    p.box(xs[7], bottom, cols[7][1], rowH, 0.8);
    p.fill(xs[FGF], bottom, fgfW, rowH, FGF_GRAY);
    for (let i = FGF + 1; i < NOTE; i++) p.line(xs[i], bottom, xs[i], bottom + rowH, 0.5);
    p.box(xs[FGF], bottom, fgfW, rowH, 0.8);
    p.text(xs[7] - 8, bottom + rowH / 2 + 3.5, 'Summe', 10, { bold: true, align: 'right' });
    p.text(xs[7] + cols[7][1] / 2, bottom + rowH / 2 + 4, d.total || '', 11, { bold: true, align: 'center' });

    const sTop = bottom + rowH + 24;
    const lineY = sTop + 60;
    p.text(L, sTop, 'Ich bestätige die Richtigkeit der gemachten Angaben', 10, { bold: true });
    p.line(L, lineY, L + 130, lineY, 0.6);
    p.text(L, lineY + 11, 'Datum', 8, { gray: 0.35 });
    if (d.signDate) p.text(L + 65, lineY - 6, d.signDate, 11, { align: 'center' });
    p.line(L + 160, lineY, L + 400, lineY, 0.6);
    p.text(L + 160, lineY + 11, 'Unterschrift', 8, { gray: 0.35 });
    if (d.signature) {
      const sc = Math.min(240 / d.signature.w, 46 / d.signature.h);
      const w = d.signature.w * sc, h = d.signature.h * sc;
      p.image('Im1', L + 160 + (240 - w) / 2, lineY - 2 - h, w, h);
    }
    const gx = 470;
    p.text(gx, sTop, 'Genehmigt durch Produktionsleitung/Herstellungsleitung', 10, { bold: true });
    p.line(gx, lineY, gx + 120, lineY, 0.6);
    p.text(gx, lineY + 11, 'Datum', 8, { gray: 0.35 });
    p.line(gx + 150, lineY, R, lineY, 0.6);
    p.text(gx + 150, lineY + 11, 'Unterschrift', 8, { gray: 0.35 });

    return assemble(concat(p.ops), d.signature, d.title);
  }

  global.StundenzettelPDF = { build };
})(window);
