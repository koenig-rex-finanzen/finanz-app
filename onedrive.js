// Zugriff auf die Datei in OneDrive über Microsoft Graph.
import { CONFIG } from './config.js';
import { token } from './anmeldung.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const FELDER = 'id,name,eTag,size,lastModifiedDateTime,lastModifiedBy,parentReference,webUrl';

export class AnmeldungNoetig extends Error { constructor() { super('Bitte bei Microsoft anmelden.'); } }

async function graph(pfad, { roh = false, ...opts } = {}) {
  const t = await token();
  if (!t) throw new AnmeldungNoetig();
  const r = await fetch(pfad.startsWith('http') ? pfad : GRAPH + pfad, { ...opts, headers: { Authorization: `Bearer ${t}`, ...(opts.headers || {}) } });
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

/** Metadaten eines Orts lesen. */
export const info = (o) => graph(`/drives/${o.driveId}/items/${o.itemId}?$select=${FELDER}`);

/**
 * Sucht die Buchhaltungsdatei: gemerkter Ort → eigenes OneDrive → „Mit mir geteilt“ → Suche.
 * @returns {Promise<{ort, item}|null>}
 */
export async function dateiFinden() {
  const alt = ortLesen();
  if (alt?.driveId) {
    const it = await info(alt).catch((e) => { if (e instanceof AnmeldungNoetig) throw e; return null; });
    if (it) return { ort: alt, item: it };
  }
  // 1. Eigenes OneDrive (Kai)
  for (const p of CONFIG.pfade) {
    const it = await graph(`/me/drive/root:/${pfadKodiert(p + '/' + CONFIG.dateiName)}?$select=${FELDER}`).catch(() => null);
    if (it?.id) { const o = ortVon(it); ortMerken(o); return { ort: o, item: it }; }
  }
  // 2. Mit mir geteilt (Natalie)
  const geteilt = await graph('/me/drive/sharedWithMe').catch(() => null);
  for (const g of geteilt?.value || []) {
    const ri = g.remoteItem || g;
    const driveId = ri.parentReference?.driveId;
    if (!driveId) continue;
    if (g.name === CONFIG.dateiName) { const o = { driveId, itemId: ri.id }; const it = await info(o); if (it) { ortMerken(o); return { ort: o, item: it }; } }
    if (g.name === CONFIG.ordnerName) {
      const kinder = await graph(`/drives/${driveId}/items/${ri.id}/children?$select=${FELDER}`).catch(() => null);
      const it = kinder?.value?.find((k) => k.name === CONFIG.dateiName);
      if (it) { const o = { driveId: it.parentReference?.driveId || driveId, itemId: it.id }; ortMerken(o); return { ort: o, item: it }; }
    }
  }
  // 3. Suche im eigenen OneDrive
  const treffer = await graph(`/me/drive/root/search(q='${encodeURIComponent(CONFIG.dateiName.replace('.xlsx', ''))}')?$select=${FELDER}`).catch(() => null);
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
