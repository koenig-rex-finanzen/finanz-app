// Übernimmt Blätter unverändert aus einer alten .xlsx (Formeln, Formatierung, Diagramme, Tabellen, Auswahllisten).
// Für die Zusatzblätter beim Umzug: Gemeinsam (Aufteilung, Überweisung, Kredite, Jahr) und Kai (Depot).
// Formeln können dabei umgeschrieben werden (z. B. neue Spaltenbuchstaben).

const td = new TextDecoder();
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Alle Teile einer .xlsx als Text. */
export async function entpacken(bytes) {
  const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
  let e = u.length - 22; while (e > 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
  const out = {};
  for (let i = 0; i < n; i++) {
    const meth = dv.getUint16(p + 10, true), csz = dv.getUint32(p + 20, true), nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    const name = td.decode(u.subarray(p + 46, p + 46 + nl));
    const lnl = dv.getUint16(off + 26, true), lel = dv.getUint16(off + 28, true);
    const daten = u.subarray(off + 30 + lnl + lel, off + 30 + lnl + lel + csz);
    const roh = meth === 0 ? daten : new Uint8Array(await new Response(new Blob([daten]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
    out[name] = td.decode(roh);
    p += 46 + nl + el + cl;
  }
  return out;
}

const relsVon = (xml) => [...(xml || '').matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
  id: m[0].match(/Id="([^"]+)"/)[1], ziel: m[0].match(/Target="([^"]+)"/)[1], typ: m[0].match(/Type="([^"]+)"/)[1].split('/').pop(),
}));
const pfad = (basis, ziel) => { const teile = basis.split('/').slice(0, -1); for (const s of ziel.split('/')) { if (s === '..') teile.pop(); else if (s !== '.') teile.push(s); } return teile.join('/'); };

/**
 * Bereitet Blätter zum Kopieren vor.
 * @param {object} z aus entpacken()
 * @param {string[]} namen Blattnamen
 * @param {{formel?: (blatt, text) => string, texte?: Object<string, Object<string,string>>}} opt
 *   formel: schreibt eine Formel um (Text ohne „=“); texte: je Blatt Zuordnung alter → neuer Zelltext
 * @returns {{blaetter: Array, fremdeStile: string, fehlend: string[]}}
 */
export function blaetterVorbereiten(z, namen, opt = {}) {
  const wb = z['xl/workbook.xml'];
  const wbRels = relsVon(z['xl/_rels/workbook.xml.rels']);
  const strings = [...(z['xl/sharedStrings.xml'] || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unesc([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((x) => x[1]).join('')));
  const blaetter = []; const fehlend = [];
  for (const name of namen) {
    const m = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((x) => x[0]).find((s) => s.includes(`name="${esc(name)}"`));
    if (!m) { fehlend.push(name); continue; }
    const rid = m.match(/r:id="([^"]+)"/)[1];
    const sp = pfad('xl/workbook.xml', wbRels.find((r) => r.id === rid).ziel.replace(/^\/xl\//, ''));
    let xml = z[sp];
    if (opt.formel) xml = xml.replace(/(<f\b[^>]*>)([^<]*)(<\/f>)/g, (a, k, f, e) => k + esc(opt.formel(name, unesc(f))) + e);
    // Bestehende Ergebnisse bleiben stehen; Excel rechnet beim Öffnen neu.
    const relsPfad = sp.replace(/worksheets\//, 'worksheets/_rels/') + '.rels';
    const teile = [];
    for (const r of relsVon(z[relsPfad])) {
      const ziel = pfad(sp, r.ziel);
      if (r.typ === 'table') teile.push({ rId: r.id, typ: 'table', xml: z[ziel] });
      else if (r.typ === 'drawing') {
        const dRels = relsVon(z[ziel.replace(/drawings\//, 'drawings/_rels/') + '.rels']);
        const unter = dRels.filter((d) => d.typ === 'chart').map((d) => ({ rId: d.id, typ: 'chart', xml: opt.formel ? z[pfad(ziel, d.ziel)] : z[pfad(ziel, d.ziel)] }));
        teile.push({ rId: r.id, typ: 'drawing', xml: z[ziel], unter });
      }
      // Kommentare, VML usw. werden nicht übernommen
    }
    xml = xml.replace(/<legacyDrawing\b[^>]*\/>/g, '').replace(/<legacyDrawingHF\b[^>]*\/>/g, '');
    blaetter.push({ name, kopie: { xml, strings, textMap: opt.texte?.[name] || null, teile } });
  }
  // Teil für neuere Formeln (FILTER, LET …): Zellen mit cm="…" brauchen ihn
  return { blaetter, fremdeStile: z['xl/styles.xml'], metadata: z['xl/metadata.xml'] || null, fehlend };
}

/** Ersetzt in einer Formel Spaltenbuchstaben eines Blatts nach Zuordnung, z. B. {Kosten: {G: 'N'}}. */
export function spaltenUmschreiben(formel, zuordnung) {
  return formel.replace(/(?:'([^']+)'|([A-Za-zÄÖÜäöüß]+))!(\$?)([A-Z]{1,2})(\$?\d+)(?::(\$?)([A-Z]{1,2})(\$?\d+))?/g, (a, q, u, d1, c1, r1, d2, c2, r2) => {
    const blatt = q || u; const map = zuordnung[blatt];
    if (!map) return a;
    const vor = q ? `'${q}'` : u;
    const n1 = map[c1] || c1; const n2 = c2 ? map[c2] || c2 : null;
    return `${vor}!${d1}${n1}${r1}` + (c2 ? `:${d2}${n2}${r2}` : '');
  });
}

/**
 * Rechnet CHOOSECOLS(…tblX[]…, 2, 5, 4) auf eine neue Spaltenreihenfolge um.
 * @param {string} formel
 * @param {Object<string, {alt: string[], neu: string[], umbenannt?: Object<string,string>}>} tabellen
 * @param {(text: string) => void} [hinweis]
 */
export function chooseColsUmschreiben(formel, tabellen, hinweis = () => {}) {
  let out = ''; let i = 0;
  const key = 'CHOOSECOLS(';
  while (true) {
    const p = formel.indexOf(key, i);
    if (p < 0) { out += formel.slice(i); break; }
    const start = p + key.length;
    // erstes Argument bis zum Komma auf Ebene 0
    let tiefe = 0, j = start, inText = false;
    for (; j < formel.length; j++) {
      const c = formel[j];
      if (c === '"') inText = !inText;
      if (inText) continue;
      if (c === '(' || c === '[') tiefe++;
      else if (c === ')' || c === ']') tiefe--;
      else if (c === ',' && tiefe === 0) break;
    }
    const arg = formel.slice(start, j);
    const m = formel.slice(j).match(/^((?:,\s*\d+)+)\)/);
    const tab = Object.keys(tabellen).find((t) => arg.includes(t + '['));
    if (!m || !tab) { out += formel.slice(i, j); i = j; continue; }
    const t = tabellen[tab];
    const neu = m[1].split(',').slice(1).map((x) => {
      const name = t.alt[+x - 1]; const ziel = t.umbenannt?.[name] ?? name; const k = t.neu.indexOf(ziel);
      if (k < 0) { hinweis(`CHOOSECOLS: Spalte „${name}“ gibt es in ${tab} nicht mehr`); return x.trim(); }
      return String(k + 1);
    });
    out += formel.slice(i, j) + ',' + neu.join(',') + ')';
    i = j + m[0].length;
  }
  return out;
}
