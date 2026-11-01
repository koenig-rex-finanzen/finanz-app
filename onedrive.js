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
