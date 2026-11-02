// Zugriff auf die Datei in OneDrive über Microsoft Graph.
import { CONFIG } from './config.js';
import { token } from './anmeldung.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const FELDER = 'id,name,eTag,size,lastModifiedDateTime,lastModifiedBy,parentReference,webUrl';

export class AnmeldungNoetig extends Error { constructor() { super('Bitte bei Microsoft anmelden.'); } }
export class KeinNetz extends Error { constructor() { super('Keine Verbindung zu OneDrive.'); } }
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
export async function buchungEintragen(ort, b) {
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
    const spA = await graph(W + `/worksheets('Buchungen')/range(address='A${erste}:A${letzte}')?$select=values`, { headers: S });
    const nummern = spA.values.map((r) => String(r[0] ?? '')).filter((x) => /^B\d+$/.test(x)).map((x) => Number(x.slice(1)));
    const frei = spA.values.findIndex((r) => r[0] === '' || r[0] == null);
    const nr = 'B' + String((nummern.length ? Math.max(...nummern) : 0) + 1).padStart(4, '0');
    let zeile;
    const serial = (iso) => { const [y, mo, d] = iso.split('-').map(Number); return (Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
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
async function mitSitzung(ort, arbeit) {
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
    return z;
  });
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
