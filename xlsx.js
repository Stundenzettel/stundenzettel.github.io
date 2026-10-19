/*
 * Stundenzettel – liest die Excel-Vorlage und füllt sie aus.
 * Ohne externe Bibliotheken: ZIP lesen (DecompressionStream) und schreiben (unkomprimiert)
 * sind hier selbst umgesetzt, damit die App komplett offline läuft.
 */
(function (global) {
  'use strict';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const TEXT_PART = /\.(xml|rels|vml)$/i;
  const DAY_ROWS = [12, 13, 14, 15, 16, 17, 18];
  const EMU_PER_PX = 9525;
  // Catering-Kästchen als Unicode-Zeichen: Die Vorlage nutzt Wingdings ("q"), die es auf dem iPhone
  // nicht gibt – dort stünden sonst Buchstaben statt Kästchen.
  const BOX_CHECKED = '☑', BOX_EMPTY = '☐';
  const SYMBOL_FONT = 'Segoe UI Symbol';

  // ---------- CRC32 ----------
  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // ---------- ZIP lesen ----------
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Dieses Gerät kann die Vorlage nicht entpacken (iOS 16.4 oder neuer nötig).');
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buffer) {
    const u8 = new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Das ist keine Excel-Datei (.xlsx).');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const names = new TextDecoder();
    const files = [];
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Die Excel-Datei ist beschädigt.');
      const method = dv.getUint16(p + 10, true);
      const size = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = names.decode(u8.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (name.endsWith('/')) continue;
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const raw = u8.subarray(start, start + size);
      let data;
      if (method === 0) data = raw.slice();
      else if (method === 8) data = await inflateRaw(raw);
      else throw new Error('Die Excel-Datei verwendet eine unbekannte Komprimierung.');
      files.push({ name, data });
    }
    return files;
  }

  // ---------- ZIP schreiben ----------
  // entries: [{ name, crc, size, method (0 = gespeichert, 8 = komprimiert), data }]
  function zipEntries(entries) {
    const enc = new TextEncoder();
    const now = new Date();
    const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const chunks = [];
    const central = [];
    let offset = 0;
    for (const f of entries) {
      const name = enc.encode(f.name);
      const packed = f.data.length;

      const lh = new Uint8Array(30 + name.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(8, f.method, true);
      lv.setUint16(10, time, true);
      lv.setUint16(12, date, true);
      lv.setUint32(14, f.crc, true);
      lv.setUint32(18, packed, true);
      lv.setUint32(22, f.size, true);
      lv.setUint16(26, name.length, true);
      lh.set(name, 30);
      chunks.push(lh, f.data);

      const ch = new Uint8Array(46 + name.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(10, f.method, true);
      cv.setUint16(12, time, true);
      cv.setUint16(14, date, true);
      cv.setUint32(16, f.crc, true);
      cv.setUint32(20, packed, true);
      cv.setUint32(24, f.size, true);
      cv.setUint16(28, name.length, true);
      cv.setUint32(42, offset, true);
      ch.set(name, 46);
      central.push(ch);
      offset += lh.length + packed;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    return new Blob(chunks.concat(central, [end]), { type: XLSX_MIME });
  }

  // Ohne Komprimierung – synchron, damit das Teilen direkt im Tipp passiert.
  function zip(files) {
    return zipEntries(files.map(f => ({ name: f.name, crc: crc32(f.data), size: f.data.length, method: 0, data: f.data })));
  }

  // ---------- XML-Helfer ----------
  function attr(tag, name) {
    const m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag);
    return m ? m[1] : null;
  }

  function escXml(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function unescXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d)).replace(/&amp;/g, '&');
  }

  function colNum(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  }

  // Findet <c r="B7" …/> bzw. <row r="12" …>…</row> (Elemente sind nicht verschachtelt).
  function findEl(xml, tag, ref) {
    const re = new RegExp('<' + tag + ' r="' + ref + '"(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</' + tag + '>)');
    const m = re.exec(xml);
    return m ? { start: m.index, end: m.index + m[0].length, text: m[0] } : null;
  }

  function openTag(el) {
    return el.slice(0, el.indexOf('>') + 1);
  }

  function sharedStrings(xml) {
    if (!xml) return [];
    return (xml.match(/<si>[\s\S]*?<\/si>/g) || []).map(si => {
      const ts = si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || [];
      return unescXml(ts.map(t => t.replace(/<[^>]+>/g, '')).join(''));
    });
  }

  function cellValue(xml, ref, strings) {
    const c = findEl(xml, 'c', ref);
    if (!c) return '';
    const t = attr(openTag(c.text), 't');
    if (t === 'inlineStr') {
      return unescXml((c.text.match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || []).map(x => x.replace(/<[^>]+>/g, '')).join(''));
    }
    const v = /<v[^>]*>([\s\S]*?)<\/v>/.exec(c.text);
    if (!v) return '';
    return t === 's' ? (strings[+v[1]] || '') : unescXml(v[1]);
  }

  function unhideColumn(xml, col) {
    return xml.replace(/<col\b[^>]*>/g, tag =>
      +attr(tag, 'min') === col && +attr(tag, 'max') === col ? tag.replace(/\shidden="(1|true)"/, '') : tag);
  }

  function colWidthPx(xml, col) {
    let width = null;
    const re = /<col\b[^>]*>/g;
    let m;
    while ((m = re.exec(xml))) {
      if (col >= +attr(m[0], 'min') && col <= +attr(m[0], 'max')) { width = +attr(m[0], 'width'); break; }
    }
    if (width == null) {
      const fmt = /<sheetFormatPr\b[^>]*>/.exec(xml);
      width = +(fmt && attr(fmt[0], 'defaultColWidth')) || 8.43;
    }
    return Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7);
  }

  function rowHeightPx(xml, row) {
    const r = findEl(xml, 'row', row);
    let pt = r && +attr(openTag(r.text), 'ht');
    if (!pt) {
      const fmt = /<sheetFormatPr\b[^>]*>/.exec(xml);
      pt = +(fmt && attr(fmt[0], 'defaultRowHeight')) || 15;
    }
    return pt * 96 / 72;
  }

  function relTarget(rels, id) {
    const m = new RegExp('<Relationship\\b[^>]*\\sId="' + id + '"[^>]*>').exec(rels || '');
    return m ? attr(m[0], 'Target') : null;
  }

  function resolvePath(fromFile, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = fromFile.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg && seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }

  function relativePath(fromFile, toFile) {
    const a = fromFile.split('/');
    a.pop();
    const b = toFile.split('/');
    let i = 0;
    while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    return '../'.repeat(a.length - i) + b.slice(i).join('/');
  }

  function relsPathFor(part) {
    const i = part.lastIndexOf('/');
    return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels';
  }

  function nextRelId(rels) {
    let max = 0;
    (rels.match(/\sId="rId(\d+)"/g) || []).forEach(m => { max = Math.max(max, +/\d+/.exec(m)[0]); });
    return 'rId' + (max + 1);
  }

  function freeName(parts, prefix, ext) {
    let n = 1;
    while ((prefix + n + ext) in parts) n++;
    return prefix + n + ext;
  }

  // ---------- Zeiten und Datum ----------
  function timeToFraction(t) {
    if (!t) return 0;
    const [h, m] = t.split(':').map(Number);
    return (h * 60 + m) / 1440;
  }

  // Rechnet exakt wie die Formeln der Vorlage (Spalte G und H).
  function calcDay(d) {
    d = d || {};
    const C = timeToFraction(d.start), D = timeToFraction(d.end);
    const E = timeToFraction(d.pStart), F = timeToFraction(d.pEnd);
    const G = (F - E) * 24;
    let H = ' ';
    if (!(C === 0 && D === 0)) H = D < C ? (D - G / 24 - C + 1) * 24 : (D - G / 24 - C) * 24;
    return { G, H };
  }

  function excelSerial(isoDate) {
    const [y, m, d] = isoDate.split('-').map(Number);
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // ---------- Vorlage prüfen ----------
  function inspect(parts) {
    const wb = parts['xl/workbook.xml'];
    if (typeof wb !== 'string') throw new Error('Das ist keine Excel-Datei (.xlsx).');
    const sheetTag = /<sheet\b[^>]*>/.exec(wb);
    const target = sheetTag && relTarget(parts['xl/_rels/workbook.xml.rels'], attr(sheetTag[0], 'r:id'));
    const sheetPath = target && resolvePath('xl/workbook.xml', target);
    const sheet = sheetPath && parts[sheetPath];
    if (typeof sheet !== 'string') throw new Error('In der Datei wurde kein Arbeitsblatt gefunden.');
    const strings = sharedStrings(parts['xl/sharedStrings.xml']);
    const val = ref => cellValue(sheet, ref, strings).trim();
    const ok = /^montag$/i.test(val('A12')) && /^sonntag$/i.test(val('A18')) &&
      /arbeitsbeginn/i.test(val('C11')) && /arbeitsende/i.test(val('D11'));
    if (!ok) {
      throw new Error('Diese Datei sieht nicht wie die Stundenzettel-Vorlage aus ' +
        '(erwartet: Montag–Sonntag in Zeile 12–18, Arbeitsbeginn/-ende in Spalte C/D).');
    }
    return {
      sheetPath,
      sheetName: unescXml(attr(sheetTag[0], 'name') || ''),
      company: val('A1'),
      address: [val('A3'), val('A4')].filter(Boolean),
      production: val('C6'),
      productionNo: val('L6'),
      catering: /catering/i.test(val('N11'))
    };
  }

  async function readTemplate(buffer, fileName) {
    const files = await unzip(buffer);
    const parts = {};
    const dec = new TextDecoder();
    for (const f of files) parts[f.name] = TEXT_PART.test(f.name) ? dec.decode(f.data) : f.data;
    return { fileName, importedAt: new Date().toISOString(), order: files.map(f => f.name), parts, info: inspect(parts) };
  }

  function toBase64(u8) {
    let s = '';
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function fromBase64(b64) {
    const s = atob(b64);
    const u8 = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) u8[i] = s.charCodeAt(i);
    return u8;
  }

  function serialize(tpl) {
    const parts = {};
    for (const [k, v] of Object.entries(tpl.parts)) parts[k] = typeof v === 'string' ? { s: v } : { b: toBase64(v) };
    return JSON.stringify({
      v: 1, fileName: tpl.fileName, importedAt: tpl.importedAt, source: tpl.source, bundleSig: tpl.bundleSig,
      order: tpl.order, info: tpl.info, parts
    });
  }

  function deserialize(json) {
    const o = typeof json === 'string' ? JSON.parse(json) : json;
    const parts = {};
    for (const [k, v] of Object.entries(o.parts)) parts[k] = 's' in v ? v.s : fromBase64(v.b);
    return { fileName: o.fileName, importedAt: o.importedAt, source: o.source, bundleSig: o.bundleSig, order: o.order, info: o.info, parts };
  }

  // ---------- Zellen schreiben ----------
  const styleAttr = s => (s != null ? ' s="' + s + '"' : '');
  const numCell = v => (ref, s) => '<c r="' + ref + '"' + styleAttr(s) + '><v>' + v + '</v></c>';
  const strCell = text => (ref, s) => '<c r="' + ref + '"' + styleAttr(s) +
    ' t="inlineStr"><is><t xml:space="preserve">' + escXml(text) + '</t></is></c>';
  const emptyCell = () => (ref, s) => '<c r="' + ref + '"' + styleAttr(s) + '/>';
  // Formel behalten, nur den zwischengespeicherten Wert setzen (für Vorschauen ohne Neuberechnung).
  const formulaCell = v => (ref, s, old) => {
    const f = old && /<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(old);
    if (!f) return typeof v === 'number' ? numCell(v)(ref, s) : emptyCell()(ref, s);
    return typeof v === 'number'
      ? '<c r="' + ref + '"' + styleAttr(s) + '>' + f[0] + '<v>' + v + '</v></c>'
      : '<c r="' + ref + '"' + styleAttr(s) + ' t="str">' + f[0] + '<v xml:space="preserve">' + escXml(v) + '</v></c>';
  };

  function putCell(xml, ref, build) {
    const hit = findEl(xml, 'c', ref);
    if (hit) return xml.slice(0, hit.start) + build(ref, attr(openTag(hit.text), 's'), hit.text) + xml.slice(hit.end);

    const cell = build(ref, null, null);
    const rowNo = +/\d+$/.exec(ref)[0];
    const col = colNum(/^[A-Z]+/.exec(ref)[0]);
    const row = findEl(xml, 'row', rowNo);
    if (row) {
      let text = row.text;
      if (text.endsWith('/>')) {
        text = text.slice(0, -2) + '>' + cell + '</row>';
      } else {
        const re = /<c r="([A-Z]+)\d+"/g;
        let m, pos = -1;
        while ((m = re.exec(text))) if (colNum(m[1]) > col) { pos = m.index; break; }
        if (pos < 0) pos = text.lastIndexOf('</row>');
        text = text.slice(0, pos) + cell + text.slice(pos);
      }
      return xml.slice(0, row.start) + text + xml.slice(row.end);
    }
    const re = /<row r="(\d+)"/g;
    let m, pos = -1;
    while ((m = re.exec(xml))) if (+m[1] > rowNo) { pos = m.index; break; }
    if (pos < 0) pos = xml.indexOf('</sheetData>');
    return xml.slice(0, pos) + '<row r="' + rowNo + '">' + cell + '</row>' + xml.slice(pos);
  }

  // Listen in styles.xml, z. B. <cellXfs> mit <xf>-Einträgen oder <fonts> mit <font>-Einträgen.
  const styleList = (parts, list) => new RegExp('<' + list + '\\b[^>]*>([\\s\\S]*?)<\\/' + list + '>').exec(parts['xl/styles.xml'] || '');

  function styleItems(parts, list, item) {
    const block = styleList(parts, list);
    return block ? block[1].match(new RegExp('<' + item + '\\b[^>]*\\/>|<' + item + '\\b[^>]*>[\\s\\S]*?<\\/' + item + '>', 'g')) || [] : [];
  }

  // Hängt ein Element an eine Liste an und liefert seinen Index.
  function appendStyle(parts, list, item, el) {
    const index = styleItems(parts, list, item).length;
    const styles = parts['xl/styles.xml'];
    const block = styleList(parts, list);
    const updated = block[0]
      .replace(new RegExp('<\\/' + list + '>$'), () => el + '</' + list + '>')
      .replace(new RegExp('^(<' + list + '\\b[^>]*\\scount=")\\d+(")'), '$1' + (index + 1) + '$2');
    parts['xl/styles.xml'] = styles.slice(0, block.index) + updated + styles.slice(block.index + block[0].length);
    return index;
  }

  // Liefert zu einem Zellformat eine abgewandelte Kopie (change: xf-XML → xf-XML); je Format nur einmal.
  function styleVariants(parts, change) {
    const cache = {};
    return s => {
      if (s == null) return s;
      if (s in cache) return cache[s];
      const base = styleItems(parts, 'cellXfs', 'xf')[+s];
      if (!base) return (cache[s] = s);
      return (cache[s] = String(appendStyle(parts, 'cellXfs', 'xf', change(base))));
    };
  }

  // Kopie mit Zeilenumbruch (für lange Bemerkungen).
  const wrapStyles = parts => styleVariants(parts, base => {
    let xf;
    if (/<alignment\b/.test(base)) {
      xf = base.replace(/<alignment\b([^>]*?)(\/?)>/, (m, attrs, slash) =>
        '<alignment' + attrs.replace(/\swrapText="[^"]*"/, '') + ' wrapText="1"' + slash + '>');
    } else if (/\/>$/.test(base)) {
      xf = base.slice(0, -2) + '><alignment wrapText="1"/></xf>';
    } else {
      xf = base.replace(/^(<xf\b[^>]*>)/, '$1<alignment wrapText="1"/>');
    }
    if (!/\sapplyAlignment=/.test(xf)) xf = xf.replace(/^<xf\b/, '<xf applyAlignment="1"');
    return xf;
  });

  // Kopie mit einer Schrift, die die Kästchen-Zeichen enthält (Schriftgröße der Vorlage bleibt).
  const symbolStyles = parts => styleVariants(parts, base => {
    const fontId = +(attr(/^<xf\b[^>]*>/.exec(base)[0], 'fontId') || 0);
    const font = (styleItems(parts, 'fonts', 'font')[fontId] || '<font/>')
      .replace(/<(name|charset|family|scheme)\b[^>]*\/>/g, '')
      .replace(/^<font\b([^>]*?)\/>$/, '<font$1></font>')
      .replace(/<\/font>$/, '<name val="' + SYMBOL_FONT + '"/><family val="2"/></font>');
    const id = appendStyle(parts, 'fonts', 'font', font);
    let xf = /\sfontId="/.test(base) ? base.replace(/\sfontId="\d+"/, ' fontId="' + id + '"') : base.replace(/^<xf\b/, '<xf fontId="' + id + '"');
    if (!/\sapplyFont=/.test(xf)) xf = xf.replace(/^<xf\b/, '<xf applyFont="1"');
    return xf;
  });

  function insertBefore(xml, re, snippet) {
    const m = re.exec(xml);
    const at = m ? m.index : xml.length;
    return xml.slice(0, at) + snippet + xml.slice(at);
  }

  function forceRecalc(wb) {
    if (/<calcPr\b[^>]*fullCalcOnLoad=/.test(wb)) return wb;
    if (/<calcPr\b/.test(wb)) return wb.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    return insertBefore(wb, /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/,
      '<calcPr fullCalcOnLoad="1"/>');
  }

  // ---------- Unterschrift als Bild über "Unterschrift" (C22:D23) ----------
  function signatureAnchor(sheet, sig, relId, shapeId, declareNs) {
    const colW = [colWidthPx(sheet, 3), colWidthPx(sheet, 4)];
    const rowH = [rowHeightPx(sheet, 22), rowHeightPx(sheet, 23)];
    const boxW = colW[0] + colW[1], boxH = rowH[0] + rowH[1];
    const scale = Math.min((boxW - 8) / sig.w, (boxH - 4) / sig.h);
    const w = sig.w * scale, h = sig.h * scale;
    const x = (boxW - w) / 2, y = boxH - h - 2;
    const locate = (v, sizes, first) => {
      let i = 0;
      while (i < sizes.length - 1 && v >= sizes[i]) { v -= sizes[i]; i++; }
      return [first + i, Math.max(0, Math.round(v * EMU_PER_PX))];
    };
    const [c1, c1o] = locate(x, colW, 2), [r1, r1o] = locate(y, rowH, 21);
    const [c2, c2o] = locate(x + w, colW, 2), [r2, r2o] = locate(y + h, rowH, 21);
    const cx = Math.round(w * EMU_PER_PX), cy = Math.round(h * EMU_PER_PX);
    // Absolute Position: Excel nimmt den Zellanker, andere Programme (z. B. die iOS-Vorschau) die Koordinaten.
    let left = x, top = y;
    for (let c = 1; c <= 2; c++) left += colWidthPx(sheet, c);
    for (let r = 1; r <= 21; r++) top += rowHeightPx(sheet, r);
    const offX = Math.round(left * EMU_PER_PX), offY = Math.round(top * EMU_PER_PX);
    return '<xdr:twoCellAnchor' + (declareNs ? ' xmlns:xdr="' + NS_XDR + '" xmlns:a="' + NS_A + '"' : '') + ' editAs="oneCell">' +
      '<xdr:from><xdr:col>' + c1 + '</xdr:col><xdr:colOff>' + c1o + '</xdr:colOff><xdr:row>' + r1 + '</xdr:row><xdr:rowOff>' + r1o + '</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + c2 + '</xdr:col><xdr:colOff>' + c2o + '</xdr:colOff><xdr:row>' + r2 + '</xdr:row><xdr:rowOff>' + r2o + '</xdr:rowOff></xdr:to>' +
      '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + shapeId + '" name="Unterschrift" descr="Unterschrift"/>' +
      '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip xmlns:r="' + NS_R + '" r:embed="' + relId + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
      '<xdr:spPr><a:xfrm><a:off x="' + offX + '" y="' + offY + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';
  }

  function addSignature(parts, order, sheetPath, sheet, sig) {
    const put = (name, data) => { parts[name] = data; if (!order.includes(name)) order.push(name); };
    const emptyRels = XML_HEAD + '<Relationships xmlns="' + NS_PKG_REL + '"></Relationships>';
    let types = parts['[Content_Types].xml'];
    if (!/Extension="png"/i.test(types)) types = types.replace(/<Default\b/, '<Default Extension="png" ContentType="image/png"/><Default');

    const media = freeName(parts, 'xl/media/image', '.png');
    put(media, sig.png);

    const sheetRelsPath = relsPathFor(sheetPath);
    let sheetRels = parts[sheetRelsPath] || emptyRels;
    let drawingPath;
    const existing = /<drawing\b[^>]*>/.exec(sheet);
    if (existing) {
      drawingPath = resolvePath(sheetPath, relTarget(sheetRels, attr(existing[0], 'r:id')));
    } else {
      drawingPath = freeName(parts, 'xl/drawings/drawing', '.xml');
      parts[drawingPath] = XML_HEAD + '<xdr:wsDr xmlns:xdr="' + NS_XDR + '" xmlns:a="' + NS_A + '"></xdr:wsDr>';
      const rid = nextRelId(sheetRels);
      sheetRels = sheetRels.replace('</Relationships>', '<Relationship Id="' + rid +
        '" Type="' + NS_R + '/drawing" Target="' + relativePath(sheetPath, drawingPath) + '"/></Relationships>');
      put(sheetRelsPath, sheetRels);
      const nsR = /<worksheet\b[^>]*\sxmlns:r=/.test(sheet) ? '' : ' xmlns:r="' + NS_R + '"';
      sheet = insertBefore(sheet,
        /<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/,
        '<drawing' + nsR + ' r:id="' + rid + '"/>');
      types = types.replace('</Types>', '<Override PartName="/' + drawingPath + '" ContentType="' + CT_DRAWING + '"/></Types>');
    }
    parts['[Content_Types].xml'] = types;

    const drawingRelsPath = relsPathFor(drawingPath);
    let drawingRels = parts[drawingRelsPath] || emptyRels;
    const imgRel = nextRelId(drawingRels);
    drawingRels = drawingRels.replace('</Relationships>', '<Relationship Id="' + imgRel +
      '" Type="' + NS_R + '/image" Target="' + relativePath(drawingPath, media) + '"/></Relationships>');
    put(drawingRelsPath, drawingRels);

    let drawing = parts[drawingPath];
    let maxId = 1;
    (drawing.match(/cNvPr\b[^>]*\sid="(\d+)"/g) || []).forEach(m => { maxId = Math.max(maxId, +/id="(\d+)"/.exec(m)[1]); });
    const root = /<(\w+:)?wsDr\b[^>]*>/.exec(drawing)[0];
    const declareNs = !(root.includes('xmlns:xdr="' + NS_XDR + '"') && root.includes('xmlns:a="' + NS_A + '"'));
    drawing = drawing.replace(/<\/(\w+:)?wsDr>\s*$/, m => signatureAnchor(sheet, sig, imgRel, maxId + 1, declareNs) + m);
    put(drawingPath, drawing);
    return sheet;
  }

  // ---------- Stundenzettel erzeugen ----------
  // input: { name, position, monday: 'YYYY-MM-DD', days: [7 × {start,end,pStart,pEnd,note,catering}],
  //          signature: {png: Uint8Array, w, h} | null, signDate: 'YYYY-MM-DD' }
  function buildFiles(tpl, input) {
    const parts = Object.assign({}, tpl.parts);
    const order = tpl.order.slice();
    const path = tpl.info.sheetPath;
    let xml = parts[path];
    const base = excelSerial(input.monday);
    const wrap = wrapStyles(parts);
    const symbol = symbolStyles(parts);

    if (input.name) xml = putCell(xml, 'B7', strCell(input.name));
    if (input.position) xml = putCell(xml, 'K7', strCell(input.position));
    xml = putCell(xml, 'D9', numCell(base));
    xml = putCell(xml, 'I9', numCell(base + 6));

    const totals = [];
    DAY_ROWS.forEach((r, i) => {
      const d = input.days[i] || {};
      xml = putCell(xml, 'B' + r, numCell(base + i));
      ['start', 'end', 'pStart', 'pEnd'].forEach((k, j) => {
        xml = putCell(xml, 'CDEF'[j] + r, d[k] ? numCell(timeToFraction(d[k])) : emptyCell());
      });
      const { G, H } = calcDay(d);
      if (typeof H === 'number') totals.push(H);
      xml = putCell(xml, 'G' + r, formulaCell(G));
      xml = putCell(xml, 'H' + r, formulaCell(H));
      xml = putCell(xml, 'M' + r, d.note ? (ref, s) => strCell(d.note)(ref, wrap(s)) : emptyCell());
      if (tpl.info.catering) xml = putCell(xml, 'N' + r, (ref, s) => strCell(d.catering ? BOX_CHECKED : BOX_EMPTY)(ref, symbol(s)));
    });
    xml = putCell(xml, 'H19', formulaCell(totals.length ? totals.reduce((a, b) => a + b, 0) : ' '));
    // Spalte "Teilnahme Setcatering" ist in der Vorlage ausgeblendet – einblenden, damit die Häkchen sichtbar sind.
    if (tpl.info.catering) xml = unhideColumn(xml, 14);

    if (input.signature) {
      xml = putCell(xml, 'A23', numCell(excelSerial(input.signDate)));
      xml = addSignature(parts, order, path, xml, input.signature);
    }
    parts[path] = xml;
    parts['xl/workbook.xml'] = forceRecalc(parts['xl/workbook.xml']);

    const enc = new TextEncoder();
    return order.map(name => ({ name, data: typeof parts[name] === 'string' ? enc.encode(parts[name]) : parts[name] }));
  }

  const build = (tpl, input) => zip(buildFiles(tpl, input));

  global.Stundenzettel = { XLSX_MIME, readTemplate, inspect, serialize, deserialize, build, calcDay, fromBase64, crc32 };
})(window);
