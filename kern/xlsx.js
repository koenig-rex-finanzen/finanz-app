// Liest eine .xlsx-Datei (nur lesen) ohne Fremdbibliotheken.
// Eine .xlsx ist ein ZIP-Archiv mit XML-Dateien. Wir lesen die gespeicherten
// Zellwerte (bei Formeln: das zuletzt von Excel berechnete Ergebnis).
// Läuft im Browser und in Node (>= 18) über DecompressionStream.

const td = new TextDecoder('utf-8');

/** ZIP-Einträge auflisten (zentrales Verzeichnis). */
function zipEintraege(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  // Ende des zentralen Verzeichnisses suchen (Signatur 0x06054b50)
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 66000); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Keine gültige Excel-Datei (ZIP-Verzeichnis fehlt).');
  const anzahl = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const eintraege = new Map();
  for (let n = 0; n < anzahl; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP-Verzeichnis beschädigt.');
    const methode = dv.getUint16(p + 10, true);
    const groesseKomp = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const kommLen = dv.getUint16(p + 32, true);
    const lokal = dv.getUint32(p + 42, true);
    const name = td.decode(buf.subarray(p + 46, p + 46 + nameLen));
    eintraege.set(name, { methode, groesseKomp, lokal });
    p += 46 + nameLen + extraLen + kommLen;
  }
  return { dv, eintraege };
}

async function entpacken(buf, dv, e) {
  const lNameLen = dv.getUint16(e.lokal + 26, true);
  const lExtraLen = dv.getUint16(e.lokal + 28, true);
  const start = e.lokal + 30 + lNameLen + lExtraLen;
  const daten = buf.subarray(start, start + e.groesseKomp);
  if (e.methode === 0) return daten;
  if (e.methode !== 8) throw new Error('Unbekannte Kompression in der Excel-Datei.');
  const strom = new Blob([daten]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(strom).arrayBuffer());
}

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
function entities(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, e) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ENT[e]);
}
function attr(tag, name) {
  const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
  return m ? entities(m[1]) : null;
}
function textAus(xml) {
  // Alle <t>-Inhalte zusammenfügen (auch Rich-Text-Läufe)
  let s = '';
  for (const m of xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g)) s += m[1] ? entities(m[1]) : '';
  return s;
}

/** Spaltenbuchstaben → Index (A=0). */
export function spalte(buchst) {
  let n = 0;
  for (const c of buchst) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

/**
 * Liest die Arbeitsmappe.
 * @param {ArrayBuffer|Uint8Array} daten
 * @returns {Promise<{blaetter: Map<string, {zellen: Map<string, any>, zeilen: number}>}>}
 *   Zellwerte: Zahl, Text, true/false oder null. Datumswerte bleiben Excel-Zahlen.
 */
export async function leseArbeitsmappe(daten) {
  const buf = daten instanceof Uint8Array ? daten : new Uint8Array(daten);
  const { dv, eintraege } = zipEintraege(buf);
  const lies = async (pfad) => {
    const e = eintraege.get(pfad);
    return e ? td.decode(await entpacken(buf, dv, e)) : null;
  };

  // Gemeinsame Texte
  const sst = [];
  const sstXml = await lies('xl/sharedStrings.xml');
  if (sstXml) for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) sst.push(textAus(m[1]));

  // Blattnamen → Datei
  const wbXml = await lies('xl/workbook.xml');
  const relsXml = await lies('xl/_rels/workbook.xml.rels');
  const ziele = new Map();
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    ziele.set(attr(m[0], 'Id'), attr(m[0], 'Target'));
  }
  const blaetter = new Map();
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attr(m[0], 'name');
    const rid = attr(m[0], 'r:id');
    let ziel = ziele.get(rid) || '';
    ziel = ziel.startsWith('/') ? ziel.slice(1) : 'xl/' + ziel.replace(/^\.\//, '');
    const xml = await lies(ziel);
    if (!xml) continue;
    const zellen = new Map();
    let zeilen = 0;
    for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const kopf = '<c ' + c[1] + '>';
      const ref = attr(kopf, 'r');
      const typ = attr(kopf, 't');
      const inhalt = c[2] || '';
      let wert = null;
      if (typ === 'inlineStr') {
        wert = textAus(inhalt);
      } else {
        const v = inhalt.match(/<v>([\s\S]*?)<\/v>/);
        if (v) {
          const roh = entities(v[1]);
          if (typ === 's') wert = sst[parseInt(roh, 10)];
          else if (typ === 'str' || typ === 'e') wert = typ === 'e' ? null : roh;
          else if (typ === 'b') wert = roh === '1';
          else wert = Number(roh);
        }
      }
      if (wert === '' ) wert = null;
      if (wert !== null && ref) {
        zellen.set(ref, wert);
        const z = parseInt(ref.replace(/^[A-Z]+/, ''), 10);
        if (z > zeilen) zeilen = z;
      }
    }
    blaetter.set(name, { zellen, zeilen });
  }
  return { blaetter };
}
