// Schreibt eine neue .xlsx-Datei (ohne Fremdbibliothek, Browser und Node).
// Nur zum einmaligen Anlegen neuer Dateien (Bauplan: danach nur noch Microsoft Graph).
// Formeln werden ohne Ergebnis gespeichert; Excel rechnet beim Öffnen neu (fullCalcOnLoad).

const enc = new TextEncoder();
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  // Steuerzeichen sind in XML verboten
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

// ---------------------------------------------------------------- ZIP (ohne Kompression)
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u) { let c = 0xffffffff; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }

export function zip(dateien) {
  const teile = []; const zentral = []; let off = 0;
  for (const [name, inhalt] of dateien) {
    const n = enc.encode(name); const d = typeof inhalt === 'string' ? enc.encode(inhalt) : inhalt;
    const crc = crc32(d);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
    lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true); lh.setUint32(14, crc, true); lh.setUint32(18, d.length, true);
    lh.setUint32(22, d.length, true); lh.setUint16(26, n.length, true); lh.setUint16(28, 0, true);
    teile.push(new Uint8Array(lh.buffer), n, d);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, d.length, true); ch.setUint32(24, d.length, true); ch.setUint16(28, n.length, true);
    ch.setUint32(42, off, true);
    zentral.push(new Uint8Array(ch.buffer), n);
    off += 30 + n.length + d.length;
  }
  const zgroesse = zentral.reduce((s, x) => s + x.length, 0);
  const ende = new DataView(new ArrayBuffer(22));
  ende.setUint32(0, 0x06054b50, true); ende.setUint16(8, dateien.length, true); ende.setUint16(10, dateien.length, true);
  ende.setUint32(12, zgroesse, true); ende.setUint32(16, off, true);
  const alle = [...teile, ...zentral, new Uint8Array(ende.buffer)];
  const out = new Uint8Array(alle.reduce((s, x) => s + x.length, 0));
  let p = 0; for (const x of alle) { out.set(x, p); p += x.length; }
  return out;
}

// ---------------------------------------------------------------- Stile
// Feste Stil-Nummern (cellXfs). Farben neutral, damit jede Datei gleich aussieht.
export const STIL = { normal: 0, kopf: 1, titel: 2, datum: 3, monat: 4, euro: 5, formelEuro: 6, hinweis: 7, formelText: 8, formelDatum: 9, eingabe: 10, fett: 11, fettEuro: 12, prozent: 13, formelMonat: 14 };
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="164" formatCode="DD.MM.YYYY"/><numFmt numFmtId="165" formatCode="MMM YYYY"/><numFmt numFmtId="166" formatCode="#,##0.00 &quot;€&quot;;[Red]-#,##0.00 &quot;€&quot;"/><numFmt numFmtId="167" formatCode="#,##0.00 &quot;€&quot;;[Red]-#,##0.00 &quot;€&quot;;&quot;&quot;"/></numFmts>
<fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="16"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF666666"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF555555"/><name val="Calibri"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDE5D8"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF6D5"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FF999999"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="15">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="167" fontId="4" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="4" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="165" fontId="1" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="166" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="10" fontId="4" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="165" fontId="4" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

// Fremde Stile (aus einer alten Datei) hinten anhängen. Liefert neues styles.xml und den Versatz der Zell-Stile.
const teile = (xml, tag, inner) => { const m = xml.match(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`)); return m ? [...m[1].matchAll(new RegExp(`<${inner}\\b[^>]*/>|<${inner}\\b[^>]*>[\\s\\S]*?</${inner}>`, 'g'))].map((x) => x[0]) : []; };
function stileZusammen(fremd) {
  if (!fremd) return { xml: STYLES, xfVersatz: 0 };
  const eigen = { numFmts: teile(STYLES, 'numFmts', 'numFmt'), fonts: teile(STYLES, 'fonts', 'font'), fills: teile(STYLES, 'fills', 'fill'), borders: teile(STYLES, 'borders', 'border'), xfs: teile(STYLES, 'cellXfs', 'xf'), dxfs: teile(STYLES, 'dxfs', 'dxf') };
  const f = { numFmts: teile(fremd, 'numFmts', 'numFmt'), fonts: teile(fremd, 'fonts', 'font'), fills: teile(fremd, 'fills', 'fill'), borders: teile(fremd, 'borders', 'border'), xfs: teile(fremd, 'cellXfs', 'xf'), dxfs: teile(fremd, 'dxfs', 'dxf') };
  const nf = new Map(); let naechste = 300;
  const numFmts = [...eigen.numFmts];
  for (const x of f.numFmts) { const id = Number(x.match(/numFmtId="(\d+)"/)[1]); const neu = naechste++; nf.set(id, neu); numFmts.push(x.replace(/numFmtId="\d+"/, `numFmtId="${neu}"`)); }
  const mapNf = (x) => x.replace(/numFmtId="(\d+)"/g, (a, id) => `numFmtId="${nf.get(+id) ?? id}"`);
  const fo = eigen.fonts.length, fi = eigen.fills.length, bo = eigen.borders.length;
  const xfs = [...eigen.xfs, ...f.xfs.map((x) => mapNf(x).replace(/fontId="(\d+)"/, (a, n) => `fontId="${+n + fo}"`).replace(/fillId="(\d+)"/, (a, n) => `fillId="${+n + fi}"`)
    .replace(/borderId="(\d+)"/, (a, n) => `borderId="${+n + bo}"`).replace(/xfId="\d+"/, 'xfId="0"'))];
  const dxfs = [...eigen.dxfs, ...f.dxfs.map(mapNf)];
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts><fonts count="${fo + f.fonts.length}">${[...eigen.fonts, ...f.fonts].join('')}</fonts><fills count="${fi + f.fills.length}">${[...eigen.fills, ...f.fills].join('')}</fills><borders count="${bo + f.borders.length}">${[...eigen.borders, ...f.borders].join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles><dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs></styleSheet>`;
  return { xml, xfVersatz: eigen.xfs.length, dxfVersatz: eigen.dxfs.length };
}

/** Kopiertes Blatt einbauen: Stile versetzen, Texte in die neue Textliste übernehmen, Beziehungen neu nummerieren. */
function kopieXml(k, sst, stile) {
  const neuerText = (t) => { if (!sst.index.has(t)) { sst.index.set(t, sst.liste.length); sst.liste.push(t); } return sst.index.get(t); };
  let x = k.xml;
  x = x.replace(/<(c|row)\b([^>]*?)\bs="(\d+)"/g, (a, tag, vor, n) => `<${tag}${vor}s="${+n + stile.xfVersatz}"`);
  x = x.replace(/<col\b([^>]*?)\bstyle="(\d+)"/g, (a, vor, n) => `<col${vor}style="${+n + stile.xfVersatz}"`);
  x = x.replace(/dxfId="(\d+)"/g, (a, n) => `dxfId="${+n + (stile.dxfVersatz || 0)}"`);
  x = x.replace(/(<c\b[^>]*\bt="s"[^>]*>)(\s*<v>)(\d+)(<\/v>)/g, (a, kopf, v, n, e) => { const t = k.strings[+n] ?? ''; return kopf + v + neuerText(k.textMap?.[t] ?? t) + e; });
  return x;
}

// ---------------------------------------------------------------- Blätter
const SP = (i) => { let s = ''; i += 1; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
const zelleXml = (ref, z, sst) => {
  if (z == null) return '';
  const s = z.s ? ` s="${z.s}"` : '';
  if (z.f != null) return `<c r="${ref}"${s}${z.t === 'str' ? ' t="str"' : ''}><f>${esc(z.f.replace(/^=/, ''))}</f></c>`;
  const v = z.v;
  if (v == null || v === '') return z.s ? `<c r="${ref}"${s}/>` : '';
  if (typeof v === 'number') return Number.isFinite(v) ? `<c r="${ref}"${s}><v>${v}</v></c>` : '';
  if (typeof v === 'boolean') return `<c r="${ref}"${s} t="b"><v>${v ? 1 : 0}</v></c>`;
  const t = String(v);
  if (!sst.index.has(t)) { sst.index.set(t, sst.liste.length); sst.liste.push(t); }
  return `<c r="${ref}"${s} t="s"><v>${sst.index.get(t)}</v></c>`;
};

function blattXml(b, tabellenRels, sst) {
  // b.zeilen: Map<zeilenNr, Array<zelle|null>> (Spalte 0 = A)
  const rows = [...b.zeilen.entries()].sort((x, y) => x[0] - y[0]).map(([r, zellen]) => {
    const inner = zellen.map((z, c) => zelleXml(SP(c) + r, z, sst)).join('');
    const ht = b.hoehen?.[r] ? ` ht="${b.hoehen[r]}" customHeight="1"` : '';
    return inner || ht ? `<row r="${r}"${ht}>${inner}</row>` : '';
  }).join('');
  const cols = b.breiten?.length ? `<cols>${b.breiten.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w || 10}" customWidth="1"/>`).join('')}</cols>` : '';
  const frei = b.fixiert ? `<pane ySplit="${b.fixiert - 1}" topLeftCell="A${b.fixiert}" activePane="bottomLeft" state="frozen"/>` : '';
  const views = `<sheetViews><sheetView workbookViewId="0"${b.tabFarbe ? '' : ''}>${frei}</sheetView></sheetViews>`;
  const dv = b.validierungen?.length ? `<dataValidations count="${b.validierungen.length}">${b.validierungen.map((v) =>
    `<dataValidation type="list" allowBlank="1" showErrorMessage="${v.streng ? 1 : 0}" errorStyle="${v.streng ? 'stop' : 'warning'}" sqref="${v.ref}"><formula1>${esc(v.formel)}</formula1></dataValidation>`).join('')}</dataValidations>` : '';
  const tp = tabellenRels.length ? `<tableParts count="${tabellenRels.length}">${tabellenRels.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : '';
  const pr = b.tabFarbe ? `<sheetPr><tabColor rgb="${b.tabFarbe}"/></sheetPr>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${pr}${views}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rows}</sheetData>${dv}${tp}</worksheet>`;
}

/**
 * Baut eine Arbeitsmappe.
 * @param {{blaetter: Array<{name, zeilen: Map, breiten?, fixiert?, validierungen?, tabelle?: {name, ref, spalten: string[]}, tabFarbe?}>, namen?: Array<{name, ref}>}} m
 * @returns {Uint8Array}
 */
export function mappeSchreiben(m) {
  const dateien = [];
  const ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'];
  const wbRels = [];
  const sst = { index: new Map(), liste: [] };
  const stile = stileZusammen(m.fremdeStile);
  let tabNr = 0, drawNr = 0, chartNr = 0;
  const tabNamen = new Set(m.blaetter.filter((b) => b.tabelle).map((b) => b.tabelle.name));
  const TYP = { table: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart' };
  m.blaetter.forEach((b, i) => {
    const nr = i + 1;
    const rels = [];
    if (b.tabelle) {
      tabNr++;
      const t = b.tabelle;
      const cols = t.spalten.map((n, j) => `<tableColumn id="${j + 1}" name="${esc(n)}"/>`).join('');
      dateien.push([`xl/tables/table${tabNr}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tabNr}" name="${t.name}" displayName="${t.name}" ref="${t.ref}"><autoFilter ref="${t.ref}"/><tableColumns count="${t.spalten.length}">${cols}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`]);
      ct.push(`<Override PartName="/xl/tables/table${tabNr}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      rels.push(['rIdT1', `../tables/table${tabNr}.xml`, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table']);
    }
    if (b.kopie) {
      for (const t of b.kopie.teile || []) {
        if (t.typ === 'table') {
          tabNr++;
          let txml = t.xml.replace(/(<table\b[^>]*?)\bid="\d+"/, `$1id="${tabNr}"`);
          const alt = txml.match(/\bname="([^"]+)"/)?.[1];
          if (alt && tabNamen.has(alt)) { const neu = alt + '_alt'; txml = txml.replace(/\bname="[^"]+"/, `name="${neu}"`).replace(/displayName="[^"]+"/, `displayName="${neu}"`); }
          if (alt) tabNamen.add(alt);
          dateien.push([`xl/tables/table${tabNr}.xml`, txml]);
          ct.push(`<Override PartName="/xl/tables/table${tabNr}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
          rels.push([t.rId, `../tables/table${tabNr}.xml`, TYP.table]);
        } else if (t.typ === 'drawing') {
          drawNr++;
          const dRels = [];
          for (const c of t.unter || []) {
            chartNr++;
            dateien.push([`xl/charts/chart${chartNr}.xml`, c.xml]);
            ct.push(`<Override PartName="/xl/charts/chart${chartNr}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
            dRels.push(`<Relationship Id="${c.rId}" Type="${TYP.chart}" Target="../charts/chart${chartNr}.xml"/>`);
          }
          dateien.push([`xl/drawings/drawing${drawNr}.xml`, t.xml]);
          dateien.push([`xl/drawings/_rels/drawing${drawNr}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${dRels.join('')}</Relationships>`]);
          ct.push(`<Override PartName="/xl/drawings/drawing${drawNr}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
          rels.push([t.rId, `../drawings/drawing${drawNr}.xml`, TYP.drawing]);
        }
      }
      dateien.push([`xl/worksheets/sheet${nr}.xml`, kopieXml(b.kopie, sst, stile)]);
    } else dateien.push([`xl/worksheets/sheet${nr}.xml`, blattXml(b, rels.map((r) => r[0]), sst)]);
    if (rels.length) dateien.push([`xl/worksheets/_rels/sheet${nr}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map(([id, t, ty]) => `<Relationship Id="${id}" Type="${ty}" Target="${t}"/>`).join('')}</Relationships>`]);
    ct.push(`<Override PartName="/xl/worksheets/sheet${nr}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbRels.push(`<Relationship Id="rId${nr}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nr}.xml"/>`);
  });
  const n = m.blaetter.length;
  wbRels.push(`<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
  wbRels.push(`<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`);
  if (m.metadata) {
    wbRels.push(`<Relationship Id="rId${n + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>`);
    ct.push('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>');
    dateien.push(['xl/metadata.xml', m.metadata]);
  }
  ct.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
  dateien.push(['xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.liste.length}" uniqueCount="${sst.liste.length}">${sst.liste.map((t) => `<si><t xml:space="preserve">${esc(t)}</t></si>`).join('')}</sst>`]);
  const namen = (m.namen || []).map((x) => `<definedName name="${x.name}">${esc(x.ref)}</definedName>`).join('');
  dateien.unshift(
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ct.join('')}</Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`],
    ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(m.titel || 'Unser Geld')}</dc:title><dc:creator>Unser Geld</dc:creator></cp:coreProperties>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${m.blaetter.map((b, i) => `<sheet name="${esc(b.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${namen ? `<definedNames>${namen}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join('')}</Relationships>`],
    ['xl/styles.xml', stile.xml],
  );
  return zip(dateien);
}

/** Hilfe zum Füllen: Zeilen-Map mit set(ref, zelle). */
export function neuesBlatt(name, opt = {}) {
  const zeilen = new Map();
  const set = (ref, z) => {
    const m = /^([A-Z]+)(\d+)$/.exec(ref); if (!m) throw new Error('Zelle? ' + ref);
    let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; c -= 1;
    const r = Number(m[2]);
    if (!zeilen.has(r)) zeilen.set(r, []);
    const z0 = zeilen.get(r); while (z0.length <= c) z0.push(null);
    z0[c] = z && typeof z === 'object' && !(z instanceof Date) ? z : { v: z };
  };
  return { name, zeilen, set, ...opt };
}
