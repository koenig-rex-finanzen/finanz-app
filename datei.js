// Dateien im einheitlichen Format finden, laden und über Microsoft Graph (Excel-Sitzung) beschreiben.
// Schreibt nur Spalten ohne Formel (kern/format.js). Nach dem Schreiben wird gegengelesen.
import * as od from './onedrive.js';
import { BLAETTER, KOPFZEILE, ERSTE_ZEILE, BEREICH_DATEI, spalten } from './kern/format.js';
import { leseDatei } from './kern/einheit.js';

const serial = (iso) => { const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number); return (Date.UTC(y, m - 1, d || 1) - Date.UTC(1899, 11, 30)) / 864e5; };
const zelle = (typ, v) => {
  if (v == null || v === '') return '';
  if (typ === 'd') return serial(v);
  if (typ === 'm') return serial(String(v).length === 7 ? v + '-01' : v);
  if (typ === 'z' || typ === 'n') return Number(v);
  return typeof v === 'number' ? v : String(v);
};
export class Doppelt extends od.Doppelt {}
// Excel-Sitzung (für Tests austauschbar)
let mitSitzung = (ort, f) => od.mitSitzung(ort, f);
export const _sitzungFuerTest = (f) => { mitSitzung = f; };

// ---------------------------------------------------------------- Finden

const ortMerken = (k, o) => { try { localStorage.setItem('ug.neu.ort.' + k, JSON.stringify(o)); } catch {} };
const ortGemerkt = (k) => { try { return JSON.parse(localStorage.getItem('ug.neu.ort.' + k)); } catch { return null; } };
let testOrdner = null;
/** Testbetrieb: Dateien aus einem Testordner (nur für die Prüfung vor dem Livegang). */
export const testModus = (ordner) => { testOrdner = ordner; };

async function kindFinden(ordner, name) {
  const it = await od.graphRoh(`/drives/${ordner.driveId}/items/${ordner.itemId}:/${encodeURIComponent(name)}?$select=id,name,eTag,webUrl,parentReference`);
  return it?.id ? { ort: { driveId: it.parentReference?.driveId || ordner.driveId, itemId: it.id }, item: it } : null;
}
/**
 * Sucht die Datei eines Bereichs. gemeinsam/tagesgeld: neben der alten gemeinsamen Datei; kai/natalie: im Ordner Apps/<Person>.
 * @returns {Promise<{ort, item}|null>}
 */
export async function finden(bereich) {
  const name = testOrdner ? `TEST_${BEREICH_DATEI[bereich].datei}` : BEREICH_DATEI[bereich].datei;
  const g = ortGemerkt(bereich + (testOrdner ? '-test' : ''));
  if (g) { const it = await od.info(g).catch(() => null); if (it?.id && it.name === name) return { ort: g, item: it }; }
  let ordner = null;
  if (testOrdner) ordner = await od.ordnerAusPfad(testOrdner[bereich] || testOrdner.standard);
  else if (bereich === 'gemeinsam' || bereich === 'tagesgeld') {
    const alt = await od.dateiFinden();
    if (alt?.item?.parentReference) ordner = { driveId: alt.item.parentReference.driveId, itemId: alt.item.parentReference.id };
  } else ordner = await od.privatOrdnerFinden(bereich === 'kai' ? 'Kai' : 'Natalie');
  if (!ordner) return null;
  const f = await kindFinden(ordner, name);
  if (f) ortMerken(bereich + (testOrdner ? '-test' : ''), f.ort);
  return f;
}

/** Lädt und liest eine Datei. @returns {{ort, eTag, webUrl, m, bytes}} */
export async function laden(f) {
  const bytes = await od.herunterladen(f.ort);
  const m = await leseDatei(bytes);
  const it = await od.info(f.ort).catch(() => f.item);
  return { ort: f.ort, eTag: it?.eTag || f.item.eTag, webUrl: it?.webUrl || f.item.webUrl, m, bytes };
}

// ---------------------------------------------------------------- Schreiben

/** Werte einer Zeile in Spaltenreihenfolge; Formelspalten bleiben undefiniert (werden nicht geschrieben). */
function zeilenWerte(blatt, obj) {
  return spalten(blatt).map((s) => (s.typ === 'f' ? undefined : zelle(s.typ, obj[s.key])));
}
/** Schreibt eine Zeile in zusammenhängenden Stücken (Formelspalten werden übersprungen). */
async function zeileSchreiben(api, blatt, z, werte) {
  const sps = spalten(blatt);
  let start = null;
  for (let i = 0; i <= sps.length; i++) {
    const frei = i < sps.length && werte[i] !== undefined;
    if (frei && start == null) start = i;
    if (!frei && start != null) {
      await api.schreib(blatt, `${sps[start].buchst}${z}:${sps[i - 1].buchst}${z}`, [werte.slice(start, i)]);
      start = null;
    }
  }
}
async function tabellenGrenzen(api, blatt) {
  const t = await api.tabelle(BLAETTER[blatt].tabelle);
  const m = t?.address?.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!m) throw new Error(`Tabelle im Blatt „${blatt}“ nicht gefunden.`);
  return { erste: +m[2] + 1, letzte: +m[4] };
}

/**
 * Neue Zeile in die erste freie Zeile eines Blatts.
 * @param {object} obj Werte nach Spalten-Schlüssel (kern/format.js)
 * @param {{schluessel?: string, doppelt?: (zeile: any[]) => boolean}} opt  schluessel = Spalte, die eine belegte Zeile erkennt (Standard: erste Spalte)
 * @returns {Promise<number>} Zeilennummer
 */
export function zeileNeu(ort, blatt, obj, { doppelt } = {}) {
  return mitSitzung(ort, async (api) => {
    const { erste, letzte } = await tabellenGrenzen(api, blatt);
    const sps = spalten(blatt);
    const vorhanden = (await api.lies(blatt, `A${erste}:${sps.at(-1).buchst}${letzte}`)).values;
    if (doppelt) { const t = vorhanden.filter((r) => r[0] !== '' && r[0] != null && doppelt(r)); if (t.length) throw new Doppelt(t.map((r) => ({ nr: String(r[0]) }))); }
    const ersteSp = sps.findIndex((s) => s.typ !== 'f');
    const i = vorhanden.findIndex((r) => sps.every((s, j) => s.typ === 'f' || r[j] === '' || r[j] == null) && (r[ersteSp] === '' || r[ersteSp] == null));
    if (i < 0) throw new Error(`Im Blatt „${blatt}“ ist kein Platz mehr. Bitte die Tabelle in Excel verlängern.`);
    const z = erste + i;
    const werte = zeilenWerte(blatt, obj);
    await zeileSchreiben(api, blatt, z, werte);
    const k = (await api.lies(blatt, `A${z}:A${z}`)).values[0][0];
    if (String(k ?? '') !== String(werte[0] ?? '')) throw new Error('Die Zeile wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return z;
  });
}

/** Ändert eine Zeile. Prüft vorher, dass in Spalte A noch der erwartete Wert steht (z. B. Nr), damit nie die falsche Zeile überschrieben wird. */
export function zeileAendern(ort, blatt, zeile, erwartet, obj) {
  return mitSitzung(ort, async (api) => {
    const a = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(a ?? '') !== String(erwartet ?? '')) throw new Error('Die Zeile hat sich in der Zwischenzeit geändert. Bitte neu laden.');
    await zeileSchreiben(api, blatt, zeile, zeilenWerte(blatt, obj));
    return zeile;
  });
}

/** Leert eine Zeile (alle Spalten ohne Formel). Prüft vorher Spalte A. */
export function zeileLeeren(ort, blatt, zeile, erwartet) {
  return mitSitzung(ort, async (api) => {
    const a = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(a ?? '') !== String(erwartet ?? '')) throw new Error('Die Zeile hat sich in der Zwischenzeit geändert. Bitte neu laden.');
    await zeileSchreiben(api, blatt, zeile, spalten(blatt).map((s) => (s.typ === 'f' ? undefined : '')));
    const k = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(k ?? '') !== '') throw new Error('Das Löschen wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
  });
}

/**
 * Import übernehmen – alles in einer Excel-Sitzung:
 * 1. Import-IDs an vorhandene Buchungen anhängen (ergaenzen: [{nr, ids}]),
 * 2. neue Buchungen in freie Zeilen schreiben (neu: [buchung]).
 * Vorher wird in der Datei selbst noch einmal geprüft, ob eine Import-ID schon drinsteht (z. B. vom anderen Handy) – solche Zeilen werden übersprungen.
 * @returns {Promise<{neu:number, ergaenzt:number, uebersprungen:string[]}>}
 */
export function importUebernehmen(ort, { neu = [], ergaenzen = [] }) {
  return mitSitzung(ort, async (api) => {
    const { erste, letzte } = await tabellenGrenzen(api, 'Buchungen');
    const sps = spalten('Buchungen');
    const iNr = sps.findIndex((s) => s.key === 'nr'), iId = sps.findIndex((s) => s.key === 'importId');
    const v = (await api.lies('Buchungen', `A${erste}:${sps.at(-1).buchst}${letzte}`)).values;
    const belegt = (r) => sps.some((s, j) => s.typ !== 'f' && r[j] !== '' && r[j] != null);
    const idsDa = new Set(v.flatMap((r) => String(r[iId] ?? '').split(/\s+/).filter(Boolean)));
    const nrDa = new Set(v.map((r) => String(r[iNr] ?? '')).filter(Boolean));
    const uebersprungen = [];
    // 1. Import-IDs ergänzen
    let ergaenzt = 0;
    for (const e of ergaenzen) {
      const i = v.findIndex((r) => String(r[iNr]) === e.nr);
      if (i < 0) { uebersprungen.push(`${e.nr} nicht gefunden`); continue; }
      const alt = String(v[i][iId] ?? '').split(/\s+/).filter(Boolean);
      const dazu = e.ids.filter((id) => !idsDa.has(id));
      if (!dazu.length) continue;
      const wert = [...alt, ...dazu].join(' ');
      await api.schreib('Buchungen', `${sps[iId].buchst}${erste + i}`, [[wert]]);
      v[i][iId] = wert; dazu.forEach((id) => idsDa.add(id)); ergaenzt++;
    }
    // 2. Neue Buchungen
    const rein = [];
    for (const b of neu) {
      const eigene = String(b.importId || '').split(/\s+/).filter(Boolean);
      if (eigene.some((id) => idsDa.has(id))) { uebersprungen.push(`${b.datum} ${b.betrag} € (schon importiert)`); continue; }
      if (nrDa.has(b.nr)) throw new Error(`Die Nummer ${b.nr} ist inzwischen vergeben (jemand hat gleichzeitig eingetragen). Bitte neu laden und den Import wiederholen.`);
      rein.push(b); eigene.forEach((id) => idsDa.add(id));
    }
    const freie = v.map((r, i) => (belegt(r) ? -1 : i)).filter((i) => i >= 0);
    if (freie.length < rein.length) throw new Error(`Im Blatt „Buchungen“ ist nicht genug Platz (${freie.length} frei, ${rein.length} nötig). Bitte die Tabelle in Excel verlängern.`);
    // zusammenhängende freie Zeilen in einem Rutsch schreiben
    let k = 0;
    while (k < rein.length) {
      let n = 1;
      while (k + n < rein.length && freie[k + n] === freie[k] + n) n++;
      const z = erste + freie[k];
      await api.schreib('Buchungen', `A${z}:${sps.at(-1).buchst}${z + n - 1}`, rein.slice(k, k + n).map((b) => zeilenWerte('Buchungen', b).map((x) => (x === undefined ? '' : x))));
      k += n;
    }
    // 3. Gegenlesen
    if (rein.length) {
      const a = (await api.lies('Buchungen', `A${erste}:A${letzte}`)).values.map((r) => String(r[0] ?? ''));
      const fehlt = rein.filter((b) => !a.includes(b.nr));
      if (fehlt.length) throw new Error(`${fehlt.length} Buchung(en) wurden nicht wie erwartet gespeichert (${fehlt.map((b) => b.nr).join(', ')}). Bitte in Excel prüfen.`);
    }
    return { neu: rein.length, ergaenzt, uebersprungen };
  });
}

/**
 * Buchung aufteilen (eine Excel-Sitzung): Die Original-Zeile behält den ersten Teil (Betrag, Kategorie),
 * die weiteren Teile werden neue Zeilen mit Quelle „Aufteilung“ und Kommentar „Teil von <Nr>“.
 * Prüft vorher Nummer und Betrag der Original-Zeile.
 */
export function aufteilen(ort, b, teile, nrn) {
  return mitSitzung(ort, async (api) => {
    const { erste, letzte } = await tabellenGrenzen(api, 'Buchungen');
    const sps = spalten('Buchungen');
    const i = (k) => sps.findIndex((s) => s.key === k);
    const v = (await api.lies('Buchungen', `A${erste}:${sps.at(-1).buchst}${letzte}`)).values;
    const r = b._zeile - erste;
    if (String(v[r]?.[i('nr')]) !== b.nr || Math.abs(Number(v[r][i('betrag')]) - b.betrag) >= 0.005) throw new Error(`${b.nr} hat sich in der Datei geändert. Bitte neu laden.`);
    const nrDa = new Set(v.map((x) => String(x[i('nr')] ?? '')));
    if (nrn.some((n) => nrDa.has(n))) throw new Error('Die neuen Nummern sind inzwischen vergeben. Bitte neu laden.');
    const [t1, ...rest] = teile;
    await zeileSchreiben(api, 'Buchungen', b._zeile, zeilenWerte('Buchungen', { ...b, betrag: t1.betrag, kategorie: t1.kategorie || null }));
    const belegt = (x) => sps.some((s, j) => s.typ !== 'f' && x[j] !== '' && x[j] != null);
    const freie = v.map((x, j) => (belegt(x) ? -1 : j)).filter((j) => j >= 0);
    if (freie.length < rest.length) throw new Error('Im Blatt „Buchungen“ ist nicht genug Platz. Bitte die Tabelle in Excel verlängern.');
    for (const [k, t] of rest.entries()) {
      const neu = { ...b, nr: nrn[k], betrag: t.betrag, kategorie: t.kategorie || null, posten: t.posten ?? null, quelle: 'Aufteilung', importId: null, gegenstueck: null, kommentar: `Teil von ${b.nr}` };
      await zeileSchreiben(api, 'Buchungen', erste + freie[k], zeilenWerte('Buchungen', neu));
    }
    const a = (await api.lies('Buchungen', `A${erste}:A${letzte}`)).values.map((x) => String(x[0] ?? ''));
    if (nrn.some((n) => !a.includes(n))) throw new Error('Die Aufteilung wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
  });
}

/** Setzt einen Wert im Blatt „Einstellungen“ (legt den Schlüssel an, wenn es ihn noch nicht gibt). */
export function einstellungSetzen(ort, schluessel, wert, erklaerung) {
  return mitSitzung(ort, async (api) => {
    const { erste, letzte } = await tabellenGrenzen(api, 'Einstellungen');
    const v = (await api.lies('Einstellungen', `A${erste}:C${letzte}`)).values;
    let i = v.findIndex((r) => r[0] === schluessel);
    if (i < 0) { i = v.findIndex((r) => r[0] === '' || r[0] == null); if (i < 0) throw new Error('Kein Platz mehr im Blatt Einstellungen.'); await api.schreib('Einstellungen', `A${erste + i}:C${erste + i}`, [[schluessel, wert, erklaerung || '']]); }
    else await api.schreib('Einstellungen', `B${erste + i}`, [[wert]]);
  });
}

/** Doppelt-Prüfung für Buchungen: gleiche Art, Datum, Betrag, Von, Nach. */
export function gleicheBuchung(b) {
  const sp = Object.fromEntries(spalten('Buchungen').map((s, i) => [s.key, i]));
  return (r) => Math.round(Number(r[sp.datum])) === serial(b.datum) && String(r[sp.art]) === b.art && String(r[sp.von]) === b.von
    && String(r[sp.nach]) === b.nach && Math.abs(Number(r[sp.betrag]) - Number(b.betrag)) < 0.005;
}

/** Mehrere Zeilen ans Ende eines Blatts (freie Zeilen, eine Sitzung). doppelt(row) → true = schon da, wird übersprungen. */
export function zeilenAnhaengen(ort, blatt, liste, { doppelt } = {}) {
  return mitSitzung(ort, async (api) => {
    const { erste, letzte } = await tabellenGrenzen(api, blatt);
    const sps = spalten(blatt);
    const v = (await api.lies(blatt, `A${erste}:${sps.at(-1).buchst}${letzte}`)).values;
    const neu = doppelt ? liste.filter((o) => !v.some((r) => doppelt(r, o))) : liste;
    const belegt = (r) => sps.some((s, j) => s.typ !== 'f' && r[j] !== '' && r[j] != null);
    const freie = v.map((r, j) => (belegt(r) ? -1 : j)).filter((j) => j >= 0);
    if (freie.length < neu.length) throw new Error(`Im Blatt „${blatt}“ ist nicht genug Platz (${freie.length} frei, ${neu.length} nötig). Bitte die Tabelle in Excel verlängern.`);
    for (const [k, o] of neu.entries()) await zeileSchreiben(api, blatt, erste + freie[k], zeilenWerte(blatt, o));
    return neu.length;
  });
}

/** Wert der ersten Spalte, wie er in Excel steht (für die Prüfung beim Ändern/Löschen). */
export const ersteSpalte = (blatt, obj) => zeilenWerte(blatt, obj)[0];

export { KOPFZEILE, ERSTE_ZEILE };
export const aufNeueVersionWarten = od.aufNeueVersionWarten;
export const sicherungAnlegen = od.sicherungAnlegen;
