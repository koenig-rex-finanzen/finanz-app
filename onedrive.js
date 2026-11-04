// Zugriff auf die Datei in OneDrive über Microsoft Graph.
import { CONFIG } from './config.js';
import { token } from './anmeldung.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const FELDER = 'id,name,eTag,size,lastModifiedDateTime,lastModifiedBy,parentReference,webUrl';

export class AnmeldungNoetig extends Error { constructor() { super('Bitte bei Microsoft anmelden.'); } }
export class KeinNetz extends Error { constructor() { super('Keine Verbindung zu OneDrive.'); } }
/** Beim Eintragen gefunden: dieselbe Buchung (Art, Datum, Betrag, Von, Nach) steht schon in der Datei. */
export class Doppelt extends Error {
  constructor(treffer) { super('Diese Buchung steht schon in der Datei.'); this.treffer = treffer; }
}
// Bei der Suche: „nicht da“ ist okay, aber Anmeldung/Netz-Probleme müssen durchgereicht werden
const nurFehlend = (e) => { if (e instanceof AnmeldungNoetig || e instanceof KeinNetz) throw e; return null; };

async function graph(pfad, { roh = false, ...opts } = {}) {
  let t;
  try { t = await token(); } catch { throw new KeinNetz(); }
  if (!t) throw new AnmeldungNoetig();
  let r;
  try { r = await fetch(pfad.startsWith('http') ? pfad : GRAPH + pfad, { ...opts, headers: { Authorization: `Bearer ${t}`, ...(opts.headers || {}) } }); }
  catch { throw new KeinNetz(); }
  if (r.status === 401) throw new AnmeldungNoetig();
  if (roh) { if (!r.ok) throw new Error(`OneDrive meldet Fehler ${r.status}`); return r; }
  if (r.status === 404) return null;
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error?.message || `OneDrive meldet Fehler ${r.status}`);
  return j;
}

export const ich = () => graph('/me?$select=displayName,userPrincipalName,mail');

const pfadKodiert = (p) => p.split('/').map(encodeURIComponent).join('/');
const ortVon = (item) => ({ driveId: item.parentReference?.driveId, itemId: item.id });
let gemerkt = null;
const ortLesen = () => { try { return gemerkt || JSON.parse(localStorage.getItem('ug.ort')); } catch { return gemerkt; } };
const ortMerken = (o) => { gemerkt = o; try { localStorage.setItem('ug.ort', JSON.stringify(o)); } catch {} };

/**
 * Liegt die gemeinsame Datei im eigenen OneDrive der angemeldeten Person?
 * @returns {Promise<'eigen'|'geteilt'>} – entscheidet, wer angemeldet ist (Besitzer = eigen).
 */
export async function dateiHerkunft(ort) {
  const d = await graph('/me/drive?$select=id');
  return String(d?.id || '').toLowerCase() === String(ort?.driveId || '').toLowerCase() ? 'eigen' : 'geteilt';
}

/** Wo die Datei zuletzt gefunden wurde. */
export const aktuellerOrt = () => ortLesen();

/** Metadaten eines Orts lesen. */
export const info = (o) => graph(`/drives/${o.driveId}/items/${o.itemId}?$select=${FELDER}`);

/**
 * Sucht die Buchhaltungsdatei: gemerkter Ort → eigenes OneDrive → „Mit mir geteilt“ → Suche.
 * @returns {Promise<{ort, item}|null>}
 */
export async function dateiFinden() {
  const alt = ortLesen();
  if (alt?.driveId) {
    const it = await info(alt).catch(nurFehlend);
    if (it) return { ort: alt, item: it };
  }
  // 1. Eigenes OneDrive (Kai)
  for (const p of CONFIG.pfade) {
    const it = await graph(`/me/drive/root:/${pfadKodiert(p + '/' + CONFIG.dateiName)}?$select=${FELDER}`).catch(nurFehlend);
    if (it?.id) { const o = ortVon(it); ortMerken(o); return { ort: o, item: it }; }
  }
  // 2. Mit mir geteilt (Natalie)
  const geteilt = await graph('/me/drive/sharedWithMe').catch(nurFehlend);
  for (const g of geteilt?.value || []) {
    const ri = g.remoteItem || g;
    const driveId = ri.parentReference?.driveId;
    if (!driveId) continue;
    if (g.name === CONFIG.dateiName) { const o = { driveId, itemId: ri.id }; const it = await info(o); if (it) { ortMerken(o); return { ort: o, item: it }; } }
    if (g.name === CONFIG.ordnerName) {
      const kinder = await graph(`/drives/${driveId}/items/${ri.id}/children?$select=${FELDER}`).catch(nurFehlend);
      const it = kinder?.value?.find((k) => k.name === CONFIG.dateiName);
      if (it) { const o = { driveId: it.parentReference?.driveId || driveId, itemId: it.id }; ortMerken(o); return { ort: o, item: it }; }
    }
  }
  // 3. Suche im eigenen OneDrive
  const treffer = await graph(`/me/drive/root/search(q='${encodeURIComponent(CONFIG.dateiName.replace('.xlsx', ''))}')?$select=${FELDER}`).catch(nurFehlend);
  const it = treffer?.value?.find((k) => k.name === CONFIG.dateiName);
  if (it) { const o = ortVon(it); ortMerken(o); return { ort: o, item: it }; }
  return null;
}

/** Datei über einen OneDrive-Freigabelink finden (falls die Suche nichts findet). */
export async function ueberLink(url) {
  const id = 'u!' + btoa(unescape(encodeURIComponent(url.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  const it = await graph(`/shares/${id}/driveItem?$select=${FELDER}`);
  if (!it?.id) throw new Error('Mit diesem Link wurde keine Datei gefunden.');
  if (it.name !== CONFIG.dateiName && !it.folder) throw new Error(`Der Link zeigt auf „${it.name}“, nicht auf die Buchhaltungsdatei.`);
  let ziel = it;
  if (it.folder) {
    const kinder = await graph(`/drives/${it.parentReference.driveId}/items/${it.id}/children?$select=${FELDER}`);
    ziel = kinder?.value?.find((k) => k.name === CONFIG.dateiName);
    if (!ziel) throw new Error('Im freigegebenen Ordner liegt keine Buchhaltungsdatei.');
  }
  const o = ortVon(ziel); ortMerken(o); return { ort: o, item: ziel };
}

/** Lädt den Dateiinhalt. */
export async function herunterladen(ort) {
  const r = await graph(`/drives/${ort.driveId}/items/${ort.itemId}/content`, { roh: true });
  return new Uint8Array(await r.arrayBuffer());
}

/**
 * Trägt eine Buchung in die erste freie Zeile der Tabelle „Buchungen“ ein (über eine Excel-Sitzung).
 * Die Formel-Spalten (Topf-Effekt, Zuordnung) bleiben unangetastet und rechnen selbst.
 * @param {{driveId,itemId}} ort
 * @param {object} b  { datum:'JJJJ-MM-TT', giltFuer:'JJJJ-MM', art, von, nach, betrag, bereich, posten, topf, beschreibung, kommentar, gezahltVon }
 * @returns {Promise<{nr:string, zeile:number}>}
 */
export async function buchungEintragen(ort, b, { trotzdem = false } = {}) {
  const W = `/drives/${ort.driveId}/items/${ort.itemId}/workbook`;
  const json = (o) => ({ ...o, headers: { 'Content-Type': 'application/json', ...(o.headers || {}) } });
  const sitzung = await graph(W + '/createSession', json({ method: 'POST', body: JSON.stringify({ persistChanges: true }) }));
  const S = { 'workbook-session-id': sitzung.id };
  try {
    // Tabelle finden und erste freie Zeile bestimmen
    const tab = await graph(W + "/tables('tblBuchungen')/range?$select=address,rowCount", { headers: S });
    if (!tab) throw new Error('Tabelle „Buchungen“ nicht gefunden.');
    const m = tab.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    const erste = Number(m[2]) + 1, letzte = Number(m[4]);
    const spA = await graph(W + `/worksheets('Buchungen')/range(address='A${erste}:I${letzte}')?$select=values`, { headers: S });
    const nummern = spA.values.map((r) => String(r[0] ?? '')).filter((x) => /^B\d+$/.test(x)).map((x) => Number(x.slice(1)));
    const frei = spA.values.findIndex((r) => r[0] === '' || r[0] == null);
    const nr = 'B' + String((nummern.length ? Math.max(...nummern) : 0) + 1).padStart(4, '0');
    let zeile;
    const serial = (iso) => { const [y, mo, d] = iso.split('-').map(Number); return (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
    // Dieselbe Buchung schon da? Geprüft wird in der Datei selbst, nicht in der (evtl. veralteten) Anzeige.
    if (!trotzdem) {
      const treffer = spA.values.filter((r) => /^B\d+$/.test(String(r[0] ?? ''))
        && Math.round(Number(r[1])) === serial(b.datum) && String(r[3]) === b.art && String(r[4]) === b.von && String(r[5]) === b.nach
        && Math.abs(Number(r[6]) - Number(b.betrag)) < 0.005)
        .map((r) => ({ nr: String(r[0]), posten: String(r[8] ?? '') }));
      if (treffer.length) throw new Doppelt(treffer);
    }
    const werteAN = [nr, serial(b.datum), serial(b.giltFuer + '-01'), b.art, b.von, b.nach, Number(b.betrag), b.bereich || '',
      b.posten || '', b.topf || '', b.beschreibung || '', b.kommentar || '', 'manuell', ''];
    if (frei >= 0) {
      zeile = erste + frei;
      await graph(W + `/worksheets('Buchungen')/range(address='A${zeile}:N${zeile}')`, json({ method: 'PATCH', headers: S, body: JSON.stringify({ values: [werteAN] }) }));
    } else {
      // Tabelle voll: neue Zeile anhängen (Formeln übernimmt Excel aus der Tabelle)
      const neu = await graph(W + "/tables('tblBuchungen')/rows/add", json({ method: 'POST', headers: S, body: JSON.stringify({ values: [[...werteAN, null, '', null]] }) }));
      if (!neu) throw new Error('Neue Zeile ließ sich nicht anhängen.');
      zeile = letzte + 1;
    }
    if (b.gezahltVon) {
      await graph(W + `/worksheets('Buchungen')/range(address='P${zeile}')`, json({ method: 'PATCH', headers: S, body: JSON.stringify({ values: [[b.gezahltVon]] }) }));
    }
    // Kontrolle: steht die Nummer wirklich in der Zeile?
    const kontrolle = await graph(W + `/worksheets('Buchungen')/range(address='A${zeile}')?$select=values`, { headers: S });
    if (kontrolle?.values?.[0]?.[0] !== nr) throw new Error('Die Buchung wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return { nr, zeile };
  } finally {
    await graph(W + '/closeSession', { method: 'POST', headers: S }).catch(() => {});
  }
}

/**
 * Schreibt mehrere Buchungen auf einmal (CSV-Import) und ergänzt Import-IDs bei schon eingetragenen.
 * Eine Excel-Sitzung, neue Zeilen als ein zusammenhängender Block in die ersten freien Zeilen.
 * @param {{driveId,itemId}} ort
 * @param {Array<object>} neue  Buchungen wie bei buchungEintragen, zusätzlich importId, quelle
 * @param {Array<{nr:string, importId:string}>} ergaenzen
 * @returns {Promise<{nummern:string[], ergaenzt:number}>}
 */
export async function buchungenSchreiben(ort, neue, ergaenzen = [], korrekturen = []) {
  const W = `/drives/${ort.driveId}/items/${ort.itemId}/workbook`;
  const json = (o) => ({ ...o, headers: { 'Content-Type': 'application/json', ...(o.headers || {}) } });
  const sitzung = await graph(W + '/createSession', json({ method: 'POST', body: JSON.stringify({ persistChanges: true }) }));
  const S = { 'workbook-session-id': sitzung.id };
  const patch = (adresse, values) => graph(W + `/worksheets('Buchungen')/range(address='${adresse}')`, json({ method: 'PATCH', headers: S, body: JSON.stringify({ values }) }));
  try {
    const tab = await graph(W + "/tables('tblBuchungen')/range?$select=address", { headers: S });
    const m = tab.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    const erste = Number(m[2]) + 1, letzte = Number(m[4]);
    const spA = await graph(W + `/worksheets('Buchungen')/range(address='A${erste}:A${letzte}')?$select=values`, { headers: S });
    const a = spA.values.map((r) => String(r[0] ?? ''));
    const nummern = a.filter((x) => /^B\d+$/.test(x)).map((x) => Number(x.slice(1)));
    let naechste = (nummern.length ? Math.max(...nummern) : 0) + 1;
    // Erste freie Zeile nach der letzten belegten
    let letzteBelegt = -1; a.forEach((x, i) => { if (x) letzteBelegt = i; });
    const start = erste + letzteBelegt + 1;
    if (start + neue.length - 1 > letzte) throw new Error(`In der Tabelle „Buchungen“ sind nur noch ${letzte - start + 1} Zeilen frei. Bitte in Excel die Tabelle verlängern.`);
    const serial = (iso) => { const [y, mo, d] = iso.split('-').map(Number); return (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
    const vergebene = [];
    if (neue.length) {
      const zeilenAN = neue.map((b) => {
        const nr = 'B' + String(naechste++).padStart(4, '0'); vergebene.push(nr);
        return [nr, serial(b.datum), serial(b.giltFuer + '-01'), b.art, b.von, b.nach, Number(b.betrag), b.bereich || '',
          b.posten || '', b.topf || '', b.beschreibung || '', b.kommentar || '', b.quelle || 'CSV', b.importId || ''];
      });
      const ende = start + neue.length - 1;
      await patch(`A${start}:N${ende}`, zeilenAN);
      if (neue.some((b) => b.gezahltVon)) await patch(`P${start}:P${ende}`, neue.map((b) => [b.gezahltVon || '']));
    }
    let ergaenzt = 0;
    for (const e of ergaenzen) {
      const i = a.indexOf(e.nr);
      if (i < 0) continue;
      await patch(`N${erste + i}`, [[e.importId]]); ergaenzt++;
    }
    for (const k of korrekturen) {
      const i = a.indexOf(k.nr);
      if (i < 0) continue;
      await patch(`G${erste + i}`, [[Number(k.betrag)]]);
    }
    // Kontrolle
    if (vergebene.length) {
      const k = await graph(W + `/worksheets('Buchungen')/range(address='A${start}:A${start + vergebene.length - 1}')?$select=values`, { headers: S });
      const ok = k.values.every((r, i) => r[0] === vergebene[i]);
      if (!ok) throw new Error('Die Buchungen wurden nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    }
    return { nummern: vergebene, ergaenzt };
  } finally {
    await graph(W + '/closeSession', { method: 'POST', headers: S }).catch(() => {});
  }
}


// ---------------------------------------------------------------- Einzelne Buchung ändern / löschen
export async function mitSitzung(ort, arbeit) {
  const W = `/drives/${ort.driveId}/items/${ort.itemId}/workbook`;
  const json = (o) => ({ ...o, headers: { 'Content-Type': 'application/json', ...(o.headers || {}) } });
  const sitzung = await graph(W + '/createSession', json({ method: 'POST', body: JSON.stringify({ persistChanges: true }) }));
  const S = { 'workbook-session-id': sitzung.id };
  const api = {
    lies: (blatt, adresse) => graph(W + `/worksheets('${blatt}')/range(address='${adresse}')?$select=values`, { headers: S }),
    schreib: (blatt, adresse, values) => graph(W + `/worksheets('${blatt}')/range(address='${adresse}')`, json({ method: 'PATCH', headers: S, body: JSON.stringify({ values }) })),
    tabelle: (name) => graph(W + `/tables('${name}')/range?$select=address`, { headers: S }),
  };
  try { return await arbeit(api); } finally { await graph(W + '/closeSession', { method: 'POST', headers: S }).catch(() => {}); }
}
const serialVon = (iso) => { const [y, mo, d] = iso.split('-').map(Number); return (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
async function zeileVon(api, nr) {
  const t = await api.tabelle('tblBuchungen');
  const m = t.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  const erste = Number(m[2]) + 1, letzte = Number(m[4]);
  const a = await api.lies('Buchungen', `A${erste}:A${letzte}`);
  const i = a.values.findIndex((r) => r[0] === nr);
  if (i < 0) throw new Error(`${nr} wurde in der Datei nicht gefunden.`);
  return erste + i;
}

/** Ändert eine Buchung (Spalten B bis L und P). Nummer, Quelle und Import-ID bleiben. */
export function buchungAendern(ort, nr, b) {
  return mitSitzung(ort, async (api) => {
    const z = await zeileVon(api, nr);
    await api.schreib('Buchungen', `B${z}:L${z}`, [[serialVon(b.datum), serialVon(b.giltFuer + '-01'), b.art, b.von, b.nach, Number(b.betrag),
      b.bereich || '', b.posten || '', b.topf || '', b.beschreibung || '', b.kommentar || '']]);
    await api.schreib('Buchungen', `P${z}`, [[b.gezahltVon || '']]);
    const k = await api.lies('Buchungen', `A${z}:G${z}`);
    if (k?.values?.[0]?.[0] !== nr || Math.abs(Number(k.values[0][6]) - Number(b.betrag)) >= 0.005) throw new Error('Die Änderung wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return z;
  });
}

/** Setzt nur „Gezahlt von“ (Spalte P) einer Buchung, z. B. „Ausgleich: aus Topf“ oder leer. Alles andere bleibt. */
export function gezahltVonSetzen(ort, nr, wert) {
  return mitSitzung(ort, async (api) => {
    const z = await zeileVon(api, nr);
    await api.schreib('Buchungen', `P${z}`, [[wert || '']]);
    const k = await api.lies('Buchungen', `A${z}:P${z}`);
    if (k?.values?.[0]?.[0] !== nr || String(k.values[0][15] ?? '') !== String(wert || '')) throw new Error('Die Änderung wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return z;
  });
}

/** Löscht eine Buchung: leert A bis N und P der Zeile. Die Formeln (O, Q) bleiben für die nächste Buchung. */
export function buchungLoeschen(ort, nr) {
  return mitSitzung(ort, async (api) => {
    const z = await zeileVon(api, nr);
    await api.schreib('Buchungen', `A${z}:N${z}`, [Array(14).fill('')]);
    await api.schreib('Buchungen', `P${z}`, [['']]);
    const k = await api.lies('Buchungen', `A${z}`);
    if (String(k?.values?.[0]?.[0] ?? '') !== '') throw new Error('Das Löschen wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return z;
  });
}

/**
 * Nach dem Schreiben über eine Excel-Sitzung liefert OneDrive die Datei manchmal noch einige Sekunden
 * im alten Stand aus. Wartet, bis sich die Version (eTag) gegenüber `alterETag` geändert hat.
 * @returns {Promise<object|null>} neue Metadaten, oder null, wenn sich nach `max` ms nichts geändert hat
 */
export async function aufNeueVersionWarten(ort, alterETag, { max = 45000, abstand = 2500 } = {}) {
  const ende = Date.now() + max;
  for (;;) {
    const it = await info(ort).catch(nurFehlend);
    if (it?.eTag && it.eTag !== alterETag) return it;
    if (Date.now() + abstand > ende) return null;
    await new Promise((r) => setTimeout(r, abstand));
  }
}

/** Trägt den Kontostand laut Bank zum Monatsende ins Blatt „Kontostände“ ein. */
export function kontostandEintragen(ort, monat, konto, betrag) {
  return mitSitzung(ort, async (api) => {
    const t = await api.tabelle('tblKontostaende');
    const m = t.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    const erste = Number(m[2]) + 1, letzte = Number(m[4]);
    const ab = await api.lies('Kontostände', `A${erste}:B${letzte}`);
    const ziel = serialVon(monat + '-01');
    const i = ab.values.findIndex((r) => Math.round(Number(r[0])) === ziel && r[1] === konto);
    if (i < 0) throw new Error(`Im Blatt „Kontostände“ gibt es keine Zeile für ${konto} im Monat ${monat}. Bitte in Excel ergänzen.`);
    await api.schreib('Kontostände', `C${erste + i}`, [[Number(betrag)]]);
    return erste + i;
  });
}

// ---------------------------------------------------------------- Regeldatei neben der Excel-Datei
export const REGELN_DATEI = 'Unser-Geld-Regeln.json';
export async function regelnLaden(item) {
  const d = item.parentReference?.driveId, p = item.parentReference?.id;
  if (!d || !p) return [];
  let t; try { t = await token(); } catch { throw new KeinNetz(); }
  if (!t) throw new AnmeldungNoetig();
  let r;
  try { r = await fetch(`${GRAPH}/drives/${d}/items/${p}:/${encodeURIComponent(REGELN_DATEI)}:/content`, { headers: { Authorization: `Bearer ${t}` }, cache: 'no-store' }); }
  catch { throw new KeinNetz(); }
  if (r.status === 404) return [];
  if (!r.ok) throw new Error(`Regeln konnten nicht geladen werden (${r.status})`);
  const j = await r.json().catch(() => null);
  return Array.isArray(j?.regeln) ? j.regeln : [];
}
export async function regelnSpeichern(item, regeln) {
  const d = item.parentReference?.driveId, p = item.parentReference?.id;
  const inhalt = JSON.stringify({ hinweis: 'Zuordnungsregeln für den CSV-Import von „Unser Geld“. Wird vom PC-Tool gepflegt.', regeln }, null, 2);
  return graph(`/drives/${d}/items/${p}:/${encodeURIComponent(REGELN_DATEI)}:/content`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: inhalt });
}

// ---------------------------------------------------------------- Privater Bereich: eigene Datei
// Ablage (so mit Kai und Natalie abgestimmt, 04.11.2026): ein Ordner mit dem Vornamen der Person unter
// „Apps“ in Kais OneDrive (z. B. Documents/Apps/Natalie). Kai teilt den Ordner nur mit dieser Person.
// Gefunden wird er im eigenen OneDrive (Kai) oder unter „Mit mir geteilt“ (Natalie).
// Gibt es keinen solchen Ordner, nimmt die App einen eigenen, nicht geteilten Ordner im eigenen OneDrive.
export const PRIVAT_DATEI = 'Mein-Bereich.json';
export const PRIVAT_ERSATZ = 'Unser-Geld-Privat';
export class Konflikt extends Error { constructor() { super('Die private Datei wurde gerade auf einem anderen Gerät geändert.'); } }

async function rohAnfrage(pfad, opts = {}) {
  let t;
  try { t = await token(); } catch { throw new KeinNetz(); }
  if (!t) throw new AnmeldungNoetig();
  try { return await fetch(pfad.startsWith('http') ? pfad : GRAPH + pfad, { cache: 'no-store', ...opts, headers: { Authorization: `Bearer ${t}`, ...(opts.headers || {}) } }); }
  catch { throw new KeinNetz(); }
}

const ordnerKey = (person) => 'ug.privatOrdner.' + person;
/** Sucht den privaten Ordner einer Person. @returns {{driveId, itemId, name, quelle}|null} */
export async function privatOrdnerFinden(person, { ersatz = false } = {}) {
  try { const o = JSON.parse(localStorage.getItem(ordnerKey(person))); if (o?.driveId) { const it = await graph(`/drives/${o.driveId}/items/${o.itemId}?$select=id,name,folder`).catch(nurFehlend); if (it?.folder) return o; } } catch (e) { if (e instanceof AnmeldungNoetig || e instanceof KeinNetz) throw e; }
  const merke = (o) => { try { localStorage.setItem(ordnerKey(person), JSON.stringify(o)); } catch {} return o; };
  // 1. Eigenes OneDrive: …/Apps/<Person>
  for (const p of ['Documents/Apps', 'Dokumente/Apps', 'Apps']) {
    const it = await graph(`/me/drive/root:/${pfadKodiert(p + '/' + person)}?$select=id,name,folder,parentReference`).catch(nurFehlend);
    if (it?.folder) return merke({ driveId: it.parentReference.driveId, itemId: it.id, name: it.name, quelle: 'eigen' });
  }
  // 2. Mit mir geteilt: Ordner mit dem Namen der Person
  const geteilt = await graph('/me/drive/sharedWithMe').catch(nurFehlend);
  for (const g of geteilt?.value || []) {
    const ri = g.remoteItem || g;
    if ((g.name || ri.name) === person && (ri.folder || g.folder) && ri.parentReference?.driveId) {
      return merke({ driveId: ri.parentReference.driveId, itemId: ri.id, name: person, quelle: 'geteilt' });
    }
  }
  if (!ersatz) return null;
  // 3. Ersatz (nur auf Wunsch): eigener, nicht geteilter Ordner
  let it = await graph(`/me/drive/root:/${PRIVAT_ERSATZ}?$select=id,name,folder,parentReference`).catch(nurFehlend);
  if (!it) {
    it = await graph('/me/drive/root/children', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: PRIVAT_ERSATZ, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) });
  }
  return merke({ driveId: it.parentReference.driveId, itemId: it.id, name: it.name, quelle: 'ersatz' });
}
export const privatOrdnerVergessen = (person) => { try { localStorage.removeItem(ordnerKey(person)); } catch {} };

/** Lädt die private Datei aus dem Ordner. null = gibt es noch nicht. */
export async function privatLaden(ordner) {
  const it = await graph(`/drives/${ordner.driveId}/items/${ordner.itemId}:/${encodeURIComponent(PRIVAT_DATEI)}?$select=id,eTag,webUrl,parentReference`);
  if (!it?.id) return null;
  const r = await rohAnfrage(`/drives/${ordner.driveId}/items/${it.id}/content`);
  if (r.status === 401) throw new AnmeldungNoetig();
  if (!r.ok) throw new Error(`Die private Datei konnte nicht geladen werden (${r.status}).`);
  const daten = await r.json().catch(() => { throw new Error('Die private Datei ist beschädigt (kein gültiges JSON).'); });
  return { daten, eTag: it.eTag, id: it.id, driveId: ordner.driveId, webUrl: it.webUrl };
}

/** Legt die private Datei im Ordner an (bricht ab, wenn es sie schon gibt). */
export async function privatAnlegen(ordner, daten) {
  const r = await rohAnfrage(`/drives/${ordner.driveId}/items/${ordner.itemId}:/${encodeURIComponent(PRIVAT_DATEI)}:/content?@microsoft.graph.conflictBehavior=fail`, {
    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(daten, null, 2) });
  if (r.status === 401) throw new AnmeldungNoetig();
  if (r.status === 409) throw new Konflikt();
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(r.status === 403 ? 'Kein Schreibrecht im Ordner. Ist der Ordner mit „Kann bearbeiten“ geteilt?' : j.error?.message || `Anlegen fehlgeschlagen (${r.status})`);
  return { id: j.id, eTag: j.eTag, driveId: ordner.driveId, webUrl: j.webUrl };
}

/** Speichert die private Datei, nur wenn sie seit dem Laden nicht verändert wurde (If-Match). */
export async function privatSpeichern(akt, daten) {
  const r = await rohAnfrage(`/drives/${akt.driveId}/items/${akt.id}/content`, {
    method: 'PUT', headers: { 'Content-Type': 'application/json', 'If-Match': akt.eTag }, body: JSON.stringify(daten, null, 2) });
  if (r.status === 401) throw new AnmeldungNoetig();
  if (r.status === 412) throw new Konflikt();
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error?.message || `Speichern fehlgeschlagen (${r.status})`);
  return { id: j.id, eTag: j.eTag, driveId: akt.driveId, webUrl: j.webUrl };
}

/** Liegt die gemeinsame Datei im eigenen OneDrive (also: bin ich der Besitzer)? */
export async function besitzerDerDatei(ort) {
  const d = await graph('/me/drive?$select=id').catch(nurFehlend);
  return Boolean(d?.id && ort?.driveId && d.id.toLowerCase() === ort.driveId.toLowerCase());
}

// ---------------------------------------------------------------- Töpfe in der gemeinsamen Datei ändern
// Blatt „Töpfe“ (Tabelle tblToepfe): A Topf, B Konto, C Bereich, D Sparrate / Monat, G Anfangsbestand, H Stichtag, L Bemerkung.
// E, F, I, J, K sind Formeln und werden nie beschrieben.
async function topfZeilen(api) {
  const t = await api.tabelle('tblToepfe');
  const m = t.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  const erste = Number(m[2]) + 1, letzte = Number(m[4]);
  const a = await api.lies('Töpfe', `A${erste}:A${letzte}`);
  return { erste, letzte, namen: a.values.map((r) => String(r[0] ?? '').trim()) };
}

/**
 * Ändert einen Topf: Name, Sparrate, Bemerkung. Beim Umbenennen werden auch alle Buchungen und Kosten
 * mit dem alten Namen umgestellt, damit Stand und Auswahllisten stimmen.
 */
export function topfAendern(ort, alterName, { name, sparrate, bemerkung }) {
  return mitSitzung(ort, async (api) => {
    const { erste, namen } = await topfZeilen(api);
    const i = namen.indexOf(alterName);
    if (i < 0) throw new Error(`Topf „${alterName}“ wurde im Blatt „Töpfe“ nicht gefunden.`);
    const z = erste + i;
    const neu = (name || '').trim();
    if (neu && neu !== alterName && namen.includes(neu)) throw new Error(`Einen Topf „${neu}“ gibt es schon.`);
    if (sparrate != null) await api.schreib('Töpfe', `D${z}`, [[Number(sparrate)]]);
    if (bemerkung != null) await api.schreib('Töpfe', `L${z}`, [[bemerkung]]);
    let umgestellt = 0;
    if (neu && neu !== alterName) {
      await api.schreib('Töpfe', `A${z}`, [[neu]]);
      for (const [blatt, tab] of [['Buchungen', 'tblBuchungen'], ['Kosten', 'tblKosten']]) {
        const t = await api.tabelle(tab);
        const m = t.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        const e = Number(m[2]) + 1, l = Number(m[4]);
        const j = await api.lies(blatt, `J${e}:J${l}`);
        for (let k = 0; k < j.values.length; k++) {
          if (String(j.values[k][0] ?? '').trim() === alterName) { await api.schreib(blatt, `J${e + k}`, [[neu]]); umgestellt++; }
        }
      }
    }
    const k = await api.lies('Töpfe', `A${z}:D${z}`);
    if (String(k.values[0][0]).trim() !== (neu || alterName)) throw new Error('Die Änderung wurde nicht wie erwartet gespeichert. Bitte in Excel prüfen.');
    return { zeile: z, umgestellt };
  });
}

/** Legt einen neuen Topf in der ersten freien Zeile der Tabelle „Töpfe“ an (Formeln sind dort schon vorbereitet). */
export function topfNeu(ort, { name, konto, bereich, sparrate, bemerkung, stichtag }) {
  return mitSitzung(ort, async (api) => {
    const { erste, namen } = await topfZeilen(api);
    if (namen.includes(name.trim())) throw new Error(`Einen Topf „${name.trim()}“ gibt es schon.`);
    const i = namen.findIndex((n, k) => !n && namen.slice(k).every((x) => !x || x === 'Summe'));
    if (i < 0 || namen[i] === 'Summe') throw new Error('Im Blatt „Töpfe“ ist keine Zeile mehr frei. Bitte in Excel die Tabelle verlängern.');
    const z = erste + i;
    await api.schreib('Töpfe', `A${z}:D${z}`, [[name.trim(), konto, bereich || '', Number(sparrate) || 0]]);
    await api.schreib('Töpfe', `G${z}:H${z}`, [[0, serialVon(stichtag)]]);
    if (bemerkung) await api.schreib('Töpfe', `L${z}`, [[bemerkung]]);
    const k = await api.lies('Töpfe', `A${z}`);
    if (String(k.values[0][0]).trim() !== name.trim()) throw new Error('Der Topf wurde nicht wie erwartet angelegt. Bitte in Excel prüfen.');
    return { zeile: z };
  });
}

// ---------------------------------------------------------------- Kais private Datei (Kai_Privat.xlsx im Ordner Apps/Kai)
export const KAI_DATEI = 'Kai_Privat.xlsx';
/** Sucht Kai_Privat.xlsx im privaten Ordner. @returns {{ort, item}|null} */
export async function kaiDateiFinden(ordner) {
  if (ordner) {
    const it = await graph(`/drives/${ordner.driveId}/items/${ordner.itemId}:/${encodeURIComponent(KAI_DATEI)}?$select=${FELDER}`).catch(nurFehlend);
    if (it?.id) return { ort: { driveId: it.parentReference?.driveId || ordner.driveId, itemId: it.id }, item: it };
  }
  const treffer = await graph(`/me/drive/root/search(q='Kai_Privat')?$select=${FELDER}`).catch(nurFehlend);
  const it = treffer?.value?.find((k) => k.name === KAI_DATEI);
  return it ? { ort: ortVon(it), item: it } : null;
}
