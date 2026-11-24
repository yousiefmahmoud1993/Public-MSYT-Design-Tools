// Excel export helpers (offline, no libraries).
//  XLSX.build(sheets)                      -> Blob   new workbook with values + live formulas
//  XLSX.patchTemplate(buf, path, edits, tab) -> Blob   set input cells of a sheet in an .xlsx template
//  XLSX.save(blob, fileName)               -> download / native share
(function(){
  var CRC_T = (function(){ var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8){ var c = 0xFFFFFFFF; for (var i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  var enc = new TextEncoder(), dec = new TextDecoder();
  function esc(s){ return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }

  // ---- zip ----
  // entries: [{name, data(Uint8Array) | raw:{method, crc, csize, usize, bytes}}]
  function zipWrite(entries){
    var parts = [], central = [], off = 0;
    entries.forEach(function(e){
      var name = enc.encode(e.name), method, crc, csize, usize, body;
      if (e.raw) { method = e.raw.method; crc = e.raw.crc; csize = e.raw.csize; usize = e.raw.usize; body = e.raw.bytes; }
      else { method = 0; crc = crc32(e.data); csize = usize = e.data.length; body = e.data; }
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, method, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, csize, true); h.setUint32(22, usize, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, body);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, method, true);
      c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, csize, true); c.setUint32(24, usize, true);
      c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + csize;
    });
    var csz = central.reduce(function(a, p){ return a + p.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, csz, true); end.setUint32(16, off, true);
    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }
  function zipRead(buf){
    var u8 = new Uint8Array(buf), dv = new DataView(buf), e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error("Template is not a zip file");
    var n = dv.getUint16(e + 10, true), p = dv.getUint32(e + 16, true), out = [];
    for (var i = 0; i < n; i++) {
      var method = dv.getUint16(p + 10, true), crc = dv.getUint32(p + 16, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      var nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
      var name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
      var ds = lo + 30 + dv.getUint16(lo + 26, true) + dv.getUint16(lo + 28, true);
      out.push({name: name, raw: {method: method, crc: crc, csize: csize, usize: usize, bytes: u8.subarray(ds, ds + csize)}});
      p += 46 + nl + xl + cl;
    }
    return out;
  }

  // ---- template patch ----
  // edits: [{ref:"C11", v:number|string}]
  function setCell(xml, ref, v){
    var re = new RegExp('<c r="' + ref + '"(\\s[^>]*?)?(/>|>[\\s\\S]*?</c>)');
    var m = xml.match(re);
    if (!m) throw new Error("Cell " + ref + " not found in template");
    var s = (m[1] || "").match(/\ss="\d+"/);
    var cell = typeof v === "number" && isFinite(v)
      ? '<c r="' + ref + '"' + (s ? s[0] : "") + ' t="n"><v>' + v + '</v></c>'
      : '<c r="' + ref + '"' + (s ? s[0] : "") + ' t="inlineStr"><is><t>' + esc(v) + '</t></is></c>';
    return xml.replace(re, cell);
  }
  function patchTemplate(buf, sheetPath, edits, tab){
    var entries = zipRead(buf);
    entries.forEach(function(e){
      if (e.name !== sheetPath && e.name !== "xl/workbook.xml") return;
      if (e.raw.method !== 0) throw new Error(e.name + " must be stored uncompressed in the template");
      var xml = dec.decode(e.raw.bytes);
      if (e.name === sheetPath) edits.forEach(function(d){ xml = setCell(xml, d.ref, d.v); });
      else xml = xml.replace(/activeTab="\d+"/, 'activeTab="' + tab + '"');
      e.data = enc.encode(xml); delete e.raw;
    });
    // only the chosen tab is selected
    entries.forEach(function(e){
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name)) return;
      var xml = e.data ? dec.decode(e.data) : (e.raw.method === 0 ? dec.decode(e.raw.bytes) : null);
      if (xml === null) return;
      var nx = xml.replace(/ tabSelected="(true|1)"/, "");
      if (e.name === sheetPath) nx = nx.replace(/<sheetView /, '<sheetView tabSelected="1" ');
      if (nx !== xml || e.data) { e.data = enc.encode(nx); delete e.raw; }
    });
    return zipWrite(entries);
  }

  // ---- new workbook ----
  // sheets: [{name, cols:[widths], rows:[[cell,...]], merges:["A1:D1"]}], cell = null | number | string | {v, f, s}
  // styles: 0 normal, 1 title, 2 header, 3 input, 4 result, 5 note, 6 number, 7 ok, 8 bad
  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>'
    + '<fonts count="6"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="14"/><color rgb="FF1A4A6B"/><name val="Arial"/></font>'
    + '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF1A4A6B"/><name val="Arial"/></font>'
    + '<font><i/><sz val="9"/><color rgb="FF6B7A88"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF1E7D4F"/><name val="Arial"/></font></fonts>'
    + '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FF1A4A6B"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF99"/></patternFill></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF4F9"/></patternFill></fill></fills>'
    + '<borders count="2"><border/><border><left style="thin"><color rgb="FFD7E0E7"/></left><right style="thin"><color rgb="FFD7E0E7"/></right><top style="thin"><color rgb="FFD7E0E7"/></top><bottom style="thin"><color rgb="FFD7E0E7"/></bottom></border></borders>'
    + '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="9">'
    + '<xf/>'
    + '<xf fontId="1" applyFont="1"/>'
    + '<xf fontId="2" fillId="2" applyFont="1" applyFill="1"/>'
    + '<xf fontId="0" fillId="3" borderId="1" numFmtId="0" applyFill="1" applyBorder="1"><protection locked="0"/></xf>'
    + '<xf fontId="3" fillId="4" borderId="1" numFmtId="164" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/>'
    + '<xf fontId="4" applyFont="1"/>'
    + '<xf fontId="0" borderId="1" numFmtId="164" applyBorder="1" applyNumberFormat="1"/>'
    + '<xf fontId="5" applyFont="1"/>'
    + '<xf fontId="3" applyFont="1"/>'
    + '</cellXfs></styleSheet>';
  function colName(i){ var s = ""; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function cellXml(ref, c){
    if (c === null || c === undefined) return "";
    if (typeof c !== "object") c = {v: c};
    var s = c.s ? ' s="' + c.s + '"' : "";
    if (c.f) {
      var isStr = typeof c.v === "string";
      return '<c r="' + ref + '"' + s + (isStr ? ' t="str"' : "") + '><f>' + esc(c.f.replace(/^=/, "")) + '</f>' + (c.v !== undefined && c.v !== null && (isStr || isFinite(c.v)) ? '<v>' + esc(c.v) + '</v>' : "") + '</c>';
    }
    if (typeof c.v === "number") return isFinite(c.v) ? '<c r="' + ref + '"' + s + '><v>' + c.v + '</v></c>' : "";
    return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(c.v) + '</t></is></c>';
  }
  function sheetXml(sh){
    var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>';
    if (sh.cols) x += "<cols>" + sh.cols.map(function(w, i){ return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join("") + "</cols>";
    x += "<sheetData>";
    sh.rows.forEach(function(r, ri){
      var cs = (r || []).map(function(c, ci){ return cellXml(colName(ci) + (ri + 1), c); }).join("");
      x += '<row r="' + (ri + 1) + '">' + cs + "</row>";
    });
    x += "</sheetData>";
    if (sh.merges && sh.merges.length) x += '<mergeCells count="' + sh.merges.length + '">' + sh.merges.map(function(m){ return '<mergeCell ref="' + m + '"/>'; }).join("") + "</mergeCells>";
    x += '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToHeight="0"/></worksheet>';
    return x;
  }
  function build(sheets){
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + sheets.map(function(s, i){ return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join("") + "</Types>";
    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      + sheets.map(function(s, i){ return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join("") + '</sheets><calcPr fullCalcOnLoad="1"/></workbook>';
    var wrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + sheets.map(function(s, i){ return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join("")
      + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var files = [["[Content_Types].xml", ct], ["_rels/.rels", rels], ["xl/workbook.xml", wb], ["xl/_rels/workbook.xml.rels", wrels], ["xl/styles.xml", STYLES]];
    sheets.forEach(function(s, i){ files.push(["xl/worksheets/sheet" + (i + 1) + ".xml", sheetXml(s)]); });
    return zipWrite(files.map(function(f){ return {name: f[0], data: enc.encode(f[1])}; }));
  }

  function dl(blob, name){
    var url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 60000);
  }
  // iPhone / iPad (Safari or "Add to Home Screen"): the share sheet saves to Files or opens Excel / Books; else a normal download
  function isIOS(){ var u = navigator.userAgent || ""; return /iPad|iPhone|iPod/.test(u) || (/Macintosh/.test(u) && navigator.maxTouchPoints > 1); }
  function save(blob, name){
    if (window.isNativeApp && window.isNativeApp()) { window.nativeSave(blob, name); return; }
    if (isIOS() && navigator.canShare) {
      try {
        var file = new File([blob], name, {type: blob.type || (/\.pdf$/i.test(name) ? "application/pdf" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")});
        if (navigator.canShare({files: [file]})) {
          navigator.share({files: [file], title: name}).catch(function(e){ if (!e || e.name !== "AbortError") dl(blob, name); });
          return;
        }
      } catch (e) {}
    }
    dl(blob, name);
  }
  window.XLSX = {build: build, patchTemplate: patchTemplate, save: save, colName: colName, _zipRead: zipRead};
})();
