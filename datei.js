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
  return od.mitSitzung(ort, async (api) => {
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
  return od.mitSitzung(ort, async (api) => {
    const a = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(a ?? '') !== String(erwartet ?? '')) throw new Error('Die Zeile hat sich in der Zwischenzeit geändert. Bitte neu laden.');
    await zeileSchreiben(api, blatt, zeile, zeilenWerte(blatt, obj));
    return zeile;
  });
}

/** Leert eine Zeile (alle Spalten ohne Formel). Prüft vorher Spalte A. */
export function zeileLeeren(ort, blatt, zeile, erwartet) {
  return od.mitSitzung(ort, async (api) => {
    const a = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(a ?? '') !== String(erwartet ?? '')) throw new Error('Die Zeile hat sich in der Zwischenzeit geändert. Bitte neu laden.');
    await zeileSchreiben(api, blatt, zeile, spalten(blatt).map((s) => (s.typ === 'f' ? undefined : '')));
    const k = (await api.lies(blatt, `A${zeile}`)).values[0][0];
    if (String(k ?? '') !== '') throw new Error('Das Löschen wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
  });
}

/** Setzt einen Wert im Blatt „Einstellungen“ (legt den Schlüssel an, wenn es ihn noch nicht gibt). */
export function einstellungSetzen(ort, schluessel, wert, erklaerung) {
  return od.mitSitzung(ort, async (api) => {
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

export { KOPFZEILE, ERSTE_ZEILE };
export const aufNeueVersionWarten = od.aufNeueVersionWarten;
export const sicherungAnlegen = od.sicherungAnlegen;
