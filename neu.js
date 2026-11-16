// „Unser Geld“ – einheitliche Oberfläche für Gemeinsam, Natalie und Kai (Umbau „Einheitliche Bereiche“).
// Gleiche Reiter überall: Übersicht, Vorschau, Buchungen, Töpfe, Mehr. Die Farbe zeigt den Bereich.
import * as ms from './anmeldung.js';
import * as od from './onedrive.js';
import * as D from './datei.js';
import * as E from './kern/einheit.js';
import { FUNKTIONEN, WIDGETS_START, VORHER_AKTIV, LISTEN, SPARPLAN } from './kern/format.js';
import { CONFIG } from './config.js';
import * as hilfe from './hilfe.js';

export const VERSION = '1.0.0';
const $ = (s, r = document) => r.querySelector(s);
/** Gemeinsames Dialogfenster, immer ohne alten Schließen-Handler. */
const dialog = () => { const d = document.querySelector('#menu'); d.onclose = null; d.returnValue = ''; return d; };
const h = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const eur = (n) => (n < 0 ? '−' : '') + Math.abs(n || 0).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
const heute = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
const MONATE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
const monatsname = (m, jahr) => MONATE[+m.slice(5, 7) - 1] + (jahr ? ' ' + m.slice(0, 4) : '');
const datumKurz = (d) => (d ? `${d.slice(8, 10)}.${d.slice(5, 7)}.` : '');
const TABS = [
  ['uebersicht', 'Übersicht', '<path d="M3 11l9-7 9 7"/><path d="M5 10v10h14V10"/>'],
  ['vorschau', 'Vorschau', '<path d="M4 19V5"/><path d="M4 19h16"/><path d="M8 15l4-4 3 3 5-6"/>'],
  ['buchungen', 'Buchungen', '<path d="M4 6h16M4 12h16M4 18h10"/>'],
  ['toepfe', 'Töpfe', '<path d="M5 9h14l-1.5 10h-11z"/><path d="M8 9V6a4 4 0 0 1 8 0v3"/>'],
  ['mehr', 'Mehr', '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>'],
];
const FARBE = { gemeinsam: 'gemeinsam', kai: 'kai', natalie: 'natalie' };

const z = {
  person: null, bereich: 'gemeinsam', tab: 'uebersicht', monat: heute().slice(0, 7), suche: '',
  d: {}, // gemeinsam, tagesgeld, kai, natalie: {ort, eTag, webUrl, m}
  laedt: false, fehler: null, offenNr: null, topf: null, speichert: null, test: false,
};
const privat = () => (z.person === 'Kai' ? 'kai' : 'natalie');
const haupt = () => (z.bereich === 'gemeinsam' ? z.d.gemeinsam : z.d[privat()]);
const art = () => (z.bereich === 'gemeinsam' ? 'gemeinsam' : privat());

// ---------------------------------------------------------------- Start

/** Gibt es die neuen Dateien schon? (dann startet diese Oberfläche statt der alten) */
export async function aktiv() {
  if (new URLSearchParams(location.search).has('neu') || location.hash === '#umzug') return true;
  try { if (localStorage.getItem('ug.neu.aktiv') === '1') return true; } catch {}
  if (new URLSearchParams(location.search).has('code')) { try { await ms.rueckkehrVerarbeiten(); } catch {} }
  if (!ms.angemeldet()) return false;
  const f = await D.finden('gemeinsam').catch(() => null);
  if (f) { try { localStorage.setItem('ug.neu.aktiv', '1'); } catch {} return true; }
  return false;
}

export async function start() {
  document.body.classList.add('einheit');
  themaSetzen();
  hilfe.starten();
  const q = new URLSearchParams(location.search);
  if (q.get('neu') === 'test') {
    z.test = true;
    D.testModus({ standard: 'Umbau-Test', kai: 'Apps/Kai/Umbau-Test', natalie: 'Apps/Natalie/Umbau-Test' });
  }
  if (q.get('neu') === 'beispiel') return beispielStart();
  try { await ms.rueckkehrVerarbeiten(); } catch (e) { z.fehler = e.message; }
  if (!ms.angemeldet()) return anmeldeSeite();
  if (location.hash === '#umzug') return (await import('./umzug.js')).seite();
  if (location.hash === '#kai') z.bereich = 'privat';
  zeichnen();
  await allesLaden();
  window.addEventListener('hashchange', () => { if (location.hash === '#umzug') location.reload(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && !z.speichert) allesLaden(true); });
}

/** Beispielmodus: erfundene Daten, ohne Anmeldung, nichts wird gespeichert (zum Ansehen und Testen der Oberfläche). */
async function beispielStart() {
  const B = await import('./kern/beispiel.js'); const V = await import('./kern/vorlage.js');
  const H = heute(); z.person = new URLSearchParams(location.search).get('person') || 'Kai'; z.beispiel = true;
  const datei = async (d) => ({ ort: null, eTag: 'x', webUrl: '#', m: await E.leseDatei(V.dateiBauen(d)) });
  z.d.gemeinsam = await datei({ ...B.beispielPrivat(H), bereich: 'gemeinsam', titel: 'Beispiel Gemeinsam' });
  z.d.tagesgeld = await datei(B.beispielTagesgeld(H));
  z.d[privat()] = await datei({ ...B.beispielPrivat(H), bereich: privat() });
  for (const x of [z.d.gemeinsam, z.d[privat()]]) x.m.einstellungen.Funktionen = (VORHER_AKTIV[x.m.bereich] || []).join(',');
  if (location.hash === '#kai') z.bereich = 'privat';
  zeichnen();
}

function anmeldeSeite() {
  $('#app').innerHTML = `<main class="wrap einheit"><div class="start"><h1>Unser Geld</h1><p>Bitte mit dem Microsoft-Konto anmelden, in dem eure Dateien liegen.</p>
    ${z.fehler ? `<p class="fehler">${h(z.fehler)}</p>` : ''}<button class="knopf" id="anmelden">Anmelden</button></div></main>`;
  $('#anmelden').onclick = () => ms.anmelden();
}

// Letzter Stand auf dem Gerät (Cache-Speicher des Browsers), damit die App sofort und auch ohne Netz etwas zeigt
const SPEICHER = 'unser-geld-daten';
async function zwischenSpeichern(b, d) {
  try { const c = await caches.open(SPEICHER); await c.put(new Request(`./__daten/${z.test ? 'test-' : ''}${b}`), new Response(d.bytes, { headers: { 'x-etag': d.eTag || '', 'x-url': d.webUrl || '', 'x-ort': JSON.stringify(d.ort) } })); } catch {}
}
async function ausSpeicher(b) {
  try {
    const c = await caches.open(SPEICHER); const r = await c.match(`./__daten/${z.test ? 'test-' : ''}${b}`); if (!r) return null;
    const bytes = new Uint8Array(await r.arrayBuffer());
    return { ort: JSON.parse(r.headers.get('x-ort') || 'null'), eTag: r.headers.get('x-etag'), webUrl: r.headers.get('x-url'), m: await E.leseDatei(bytes), bytes, gespeichert: true };
  } catch { return null; }
}

async function allesLaden(still = false) {
  if (z.laedt) return; z.laedt = true;
  try {
    // Wer ist angemeldet? (gemerkt, sonst über die alte gemeinsame Datei)
    try { z.person ||= localStorage.getItem('ug.neu.person'); } catch {}
    if (!z.person) {
      const alt = await od.dateiFinden().catch(() => null);
      if (alt) z.person = (await od.dateiHerkunft(alt.ort)) === 'eigen' ? CONFIG.besitzer : CONFIG.besitzer === 'Kai' ? 'Natalie' : 'Kai';
      if (z.person) { try { localStorage.setItem('ug.neu.person', z.person); } catch {} }
    }
    z.person ||= 'Kai';
    const liste = ['gemeinsam', 'tagesgeld', privat()];
    // Sofort den letzten Stand vom Gerät zeigen
    if (!still) {
      for (const b of liste) if (!z.d[b]) { const x = await ausSpeicher(b); if (x) z.d[b] = x; }
      zeichnen();
    }
    // Dann alle Dateien gleichzeitig aus OneDrive
    const fehler = [];
    await Promise.all(liste.map(async (b) => {
      try {
        const f = await D.finden(b);
        if (!f) { if (b !== privat()) fehler.push(`Die Datei für „${b}“ wurde nicht gefunden.`); return; }
        const vorher = z.d[b];
        if (vorher && !vorher.gespeichert) { const it = await od.info(f.ort).catch(() => null); if (it?.eTag === vorher.eTag) return; }
        if (vorher?.gespeichert) { const it = await od.info(f.ort).catch(() => null); if (it?.eTag && it.eTag === vorher.eTag) { vorher.gespeichert = false; vorher.ort = f.ort; return; } }
        const d = await D.laden(f);
        z.d[b] = d; zwischenSpeichern(b, d);
        zeichnen();
      } catch (e) { fehler.push(e instanceof od.AnmeldungNoetig ? 'Die Anmeldung ist abgelaufen. Bitte neu anmelden.' : e.message); }
    }));
    const offline = liste.some((b) => z.d[b]?.gespeichert);
    z.fehler = fehler.length ? (offline ? 'Keine Verbindung zu OneDrive – ihr seht den zuletzt geladenen Stand. ' : '') + fehler[0] : null;
  } finally { z.laedt = false; zeichnen(); }
}
async function neuLaden(b, alterETag) {
  const it = await D.aufNeueVersionWarten(z.d[b].ort, alterETag).catch(() => null);
  const f = await D.finden(b);
  if (f) { z.d[b] = await D.laden(f); zwischenSpeichern(b, z.d[b]); }
  return !!it;
}

// ---------------------------------------------------------------- Funktionen / Widgets (Einstellungen in der Datei)

const einst = () => haupt()?.m.einstellungen || {};
const funktionenAn = () => new Set(String(einst().Funktionen || '').split(',').map((s) => s.trim()).filter(Boolean));
const fuerBereich = (id) => !FUNKTIONEN[id].bereiche || FUNKTIONEN[id].bereiche.includes(art());
const widgetsReihe = () => {
  const an = funktionenAn();
  const reihe = String(einst().Widgets || (WIDGETS_START[art()] || []).join(',')).split(',').map((s) => s.trim()).filter(Boolean);
  for (const id of an) if (FUNKTIONEN[id]?.art === 'widget' && !reihe.includes(id)) reihe.push(id);
  return reihe.filter((id) => an.has(id) && FUNKTIONEN[id]?.art === 'widget' && fuerBereich(id));
};

// ---------------------------------------------------------------- Rechnen für die Anzeige

function ansicht() {
  const H = heute();
  if (z.bereich === 'gemeinsam') {
    const g = z.d.gemeinsam?.m, t = z.d.tagesgeld?.m;
    if (!g) return null;
    const kg = E.kern(g, H, z.monat), kt = t ? E.kern(t, H, z.monat) : null;
    return { haupt: g, k: kg, konten: [...kg.konten, ...(kt?.konten || [])], toepfe: kt?.toepfe || [], topfModell: t, tagesgeld: kt };
  }
  const p = z.d[privat()]?.m; if (!p) return null;
  const k = E.kern(p, H, z.monat);
  const mitT = z.person === 'Natalie' && funktionenAn().has('tagesgeld') && z.d.tagesgeld;
  return { haupt: p, k, konten: k.konten, toepfe: mitT ? E.kern(z.d.tagesgeld.m, H).toepfe : k.toepfe, topfModell: mitT ? z.d.tagesgeld.m : p, tagesgeld: null };
}
/** Alle Buchungen der Ansicht (Gemeinsam: Gemeinsam + Tagesgeld, doppelte Überweisungen einmal). */
function buchungenAnsicht() {
  if (z.bereich === 'gemeinsam') {
    const g = (z.d.gemeinsam?.m.buchungen || []).map((b) => ({ ...b, _datei: 'gemeinsam' }));
    const t = (z.d.tagesgeld?.m.buchungen || []).filter((b) => !b.gegenstueck).map((b) => ({ ...b, _datei: 'tagesgeld' }));
    return [...g, ...t];
  }
  return (z.d[privat()]?.m.buchungen || []).map((b) => ({ ...b, _datei: privat() }));
}

// ---------------------------------------------------------------- Zeichnen

function zeichnen() {
  document.body.dataset.bereich = z.bereich === 'gemeinsam' ? 'gemeinsam' : FARBE[privat()];
  const a = ansicht();
  const inhalt = !a ? (z.laedt ? '<p class="leer">Wird geladen … (das kann bei langsamem OneDrive etwas dauern)</p>' : z.fehler ? '' : '<p class="leer">Noch nichts geladen.</p>')
    : z.tab === 'uebersicht' ? uebersicht(a) : z.tab === 'vorschau' ? vorschauSeite(a) : z.tab === 'buchungen' ? buchungenSeite()
      : z.tab === 'toepfe' ? (z.topf ? topfSeite(a) : toepfeSeite(a)) : mehrSeite();
  $('#app').innerHTML = `${kopf()}<main class="wrap einheit ${z.bereich === 'gemeinsam' ? '' : privat() === 'kai' ? 'kai' : 'meins'}">
    ${z.beispiel ? '<div class="hinweis info">Beispielmodus mit erfundenen Daten.</div>' : ''}
    ${z.test ? '<div class="hinweis info">Testbetrieb mit Testdateien – hier Eingetragenes landet nicht in euren echten Dateien.</div>' : ''}
    ${z.fehler ? `<div class="hinweis warn">${h(z.fehler)} ${/Anmeldung/.test(z.fehler) ? '<button class="link" id="neu-anmelden">Neu anmelden</button>' : ''}</div>` : ''}
    ${z.speichert ? `<div class="hinweis info">${h(z.speichert)}</div>` : ''}
    ${z.laedt && a ? '<div class="hinweis info">Wird aktualisiert …</div>' : ''}
    ${inhalt}</main>
    ${a && z.tab === 'buchungen' ? `<button class="fab" id="eintragen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Eintragen</button>` : ''}
    <nav><div class="wrap" role="tablist">${TABS.map(([id, name, pfad]) => `<button role="tab" aria-selected="${z.tab === id}" data-tab="${id}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${pfad}</svg>${name}</button>`).join('')}</div></nav>`;
  verdrahten();
}

function kopf() {
  const wahl = `<div class="bereich-wahl" role="tablist" aria-label="Bereich">
    <button role="tab" aria-selected="${z.bereich === 'gemeinsam'}" data-bereich="gemeinsam">Gemeinsam</button>
    <button role="tab" aria-selected="${z.bereich !== 'gemeinsam'}" data-bereich="privat">${z.person === 'Kai' ? 'Mein Geld' : 'Meins'}</button></div>`;
  const jetzt = heute().slice(0, 7);
  return `<header><div class="wrap"><div class="monat">
    <button id="m-zurueck" aria-label="Monat zurück"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M15 6l-6 6 6 6"/></svg></button>
    <div class="mitte"><h1>${monatsname(z.monat, true)}</h1><small>${z.monat === jetzt ? 'dieser Monat' : z.monat < jetzt ? 'vorbei' : 'Vorschau'}</small></div>
    <button id="m-vor" aria-label="Monat vor"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M9 6l6 6-6 6"/></svg></button>
  </div>${wahl}</div></header>`;
}

// ---- Übersicht: fester Kern, Widgets, Extras
function uebersicht(a) {
  const k = a.k;
  const kern = `<section class="hero"><p>Was ist frei? <strong class="${k.frei < 0 ? 'neg' : ''}">${eur(k.frei)}</strong></p>
    <div class="rechnung" data-hilfe-schluessel="Was ist frei?">
      <div class="rz"><span>Auf den Konten (zählen für Frei)</span><b>${eur(k.aufKonten)}</b></div>
      ${k.inToepfen ? `<div class="rz"><span>− in Töpfen</span><b>${eur(-k.inToepfen)}</b></div>` : ''}
      ${k.einnahmenOffen ? `<div class="rz"><span>+ kommt noch rein</span><b>${eur(k.einnahmenOffen)}</b></div>` : ''}
      ${k.kostenOffen ? `<div class="rz"><span>− geht noch ab</span><b>${eur(-k.kostenOffen)}</b></div>` : ''}
      ${k.soll ? `<div class="rz"><span>± Überweisungen zwischen den Bereichen</span><b>${eur(k.soll)}</b></div>` : ''}
    </div></section>
    <section><h2>Kontostände</h2><div class="konten">${a.konten.map((kt) => `<div class="konto"><span class="dot" style="background:${kt.rolle === 'Sparen' ? 'var(--spar)' : 'var(--gem)'}"></span>
      <div><div class="name">${h(kt.name)}</div><div class="sub">${h([kt.bank, kt.inhaber ? 'läuft auf ' + kt.inhaber : null, kt.geschlossenAb ? 'bis ' + datumKurz(kt.geschlossenAb) + kt.geschlossenAb.slice(0, 4) : null].filter(Boolean).join(' · '))}</div></div>
      <div class="betrag">${kt.standHeute == null ? '<span class="hint">Stand fehlt</span>' : eur(kt.standHeute)}</div></div>`).join('')}</div>
      ${a.tagesgeld ? `<p class="hint" style="margin-top:8px">Auf dem Tagesgeld nicht in Töpfen: <b>${eur(a.tagesgeld.freiHeute)}</b></p>` : ''}</section>`;
  const widgets = widgetsReihe().map((id) => widget(id, a)).filter(Boolean).join('');
  const seiten = [...funktionenAn()].filter((id) => FUNKTIONEN[id]?.art === 'seite' && fuerBereich(id));
  const extras = `<section><h2>Extras</h2>${seiten.length ? `<div class="kacheln">${seiten.map((id) => `<button class="kachel" data-seite="${id}">${h(FUNKTIONEN[id].titel)}</button>`).join('')}</div>`
    : '<p class="hint">Zusatzfunktionen schaltest du unter „Mehr“ ein. „Wie vor dem Umbau“ holt alles zurück, was vorher an war.</p>'}</section>`;
  return kern + widgets + extras;
}

function widget(id, a) {
  const m = a.haupt; const H = heute();
  if (id === 'wer-ueberweist' && z.bereich === 'gemeinsam') {
    const l = E.werUeberweist(z.d.gemeinsam.m, z.d.tagesgeld?.m, z.monat);
    if (!l.length) return `<section><h2>Wer überweist was</h2><p class="hint">Für ${monatsname(z.monat)} rechnet die Aufteilung noch nichts aus.</p></section>`;
    return `<section><h2>Wer überweist was</h2><div class="konten">${l.map((u) => `<div class="konto"><span class="dot" style="background:var(--${u.person === 'Kai' ? 'kai' : 'nat'})"></span>
      <div><div class="name">${u.person} → ${h(u.nach)}</div><div class="sub">${h(u.wann)}${u.hinweis ? ' · ' + h(u.hinweis) : ''}${u.gebucht && !u.erledigt ? ` · schon ${eur(u.gebucht)}` : ''}</div></div>
      <div class="betrag">${u.erledigt ? '✓ ' : ''}${eur(u.betrag)}</div></div>`).join('')}</div></section>`;
  }
  if (id === 'limits' || id === 'sprit-limit') {
    let l = E.limits(m, z.monat); if (id === 'sprit-limit') l = l.filter((x) => /sprit/i.test(x.kategorie)); else if (widgetsReihe().includes('sprit-limit')) l = l.filter((x) => !/sprit/i.test(x.kategorie));
    if (!l.length) return id === 'sprit-limit' ? '' : `<section><h2>Limits</h2><p class="hint">Noch keine Limits. Anlegen im Blatt „Limits“ (Mehr › Excel-Datei öffnen).</p></section>`;
    return `<section><h2>${id === 'sprit-limit' ? 'Sprit-Limit' : 'Limits'}</h2><div class="toepfe">${l.map((x) => `<div class="topf ${x.anteil >= 1 ? 'minus' : ''}"><div class="n">${h(x.kategorie)}</div>
      <div class="b">${eur(x.ist)} von ${eur(x.betrag)}</div><div class="s">${x.rest >= 0 ? `noch ${eur(x.rest)}` : `${eur(-x.rest)} drüber`}</div>
      <div class="bar"><span style="width:${Math.min(100, Math.round((x.anteil || 0) * 100))}%"></span></div></div>`).join('')}</div></section>`;
  }
  if (id === 'was-kommt') {
    const l = E.wasKommt(m, z.monat, H);
    return `<section><h2>Was noch kommt</h2>${l.length ? `<div class="konten">${l.map((x) => `<div class="konto"><span class="dot" style="background:${x.betrag > 0 ? 'var(--gem)' : 'var(--warn)'}"></span>
      <div><div class="name">${h(x.titel)}</div><div class="sub">${x.tag ? `am ${x.tag}.` : 'im Monat'}${x.ueberfaellig ? ' · noch nicht gebucht' : ''}</div></div><div class="betrag">${eur(x.betrag)}</div></div>`).join('')}</div>` : '<p class="hint">Alles Geplante ist gebucht.</p>'}</section>`;
  }
  if (id === 'warnung-vorschau') {
    const v = E.vorschau(m, H, z.monat); const n = Number(einst()['Vorschau Monate']) || 3;
    const l = v.slice(0, n).find((x) => x.luecke);
    return `<section><h2>Warnung Vorschau</h2>${l ? `<div class="hinweis warn">Im ${monatsname(l.monat, true)} fehlen voraussichtlich ${eur(-l.freiEnde)}.</div>` : `<p class="hint">In den nächsten ${n} Monaten keine Lücke.</p>`}</section>`;
  }
  if (id === 'kategorien') {
    const l = E.kategorienMonat(m, z.monat);
    return `<section><h2>Kategorien</h2>${l.length ? `<div class="konten">${l.map((x) => `<div class="konto"><span class="dot" style="background:var(--gem)"></span><div><div class="name">${h(x.kategorie)}</div></div><div class="betrag">${eur(x.betrag)}</div></div>`).join('')}</div>` : '<p class="hint">Diesen Monat noch keine Ausgaben.</p>'}</section>`;
  }
  if (id === 'erinnerungen') {
    const l = (m.erinnerungen || []).filter((e) => e.aktiv !== 'Nein');
    return `<section><h2>Erinnerungen</h2>${l.length ? `<div class="konten">${l.map((e) => `<div class="konto"><span class="dot" style="background:var(--gem)"></span><div><div class="name">${h(e.text)}</div><div class="sub">${e.tag ? `am ${e.tag}.` : ''}${String(e.erledigt || '').includes(z.monat) ? ' · erledigt' : ''}</div></div><div></div></div>`).join('')}</div>` : '<p class="hint">Keine Erinnerungen. Anlegen im Blatt „Erinnerungen“.</p>'}</section>`;
  }
  if (id === 'tagesgeld' && z.d.tagesgeld) {
    const t = E.kern(z.d.tagesgeld.m, H);
    return `<section><h2>Tagesgeld und Töpfe</h2><div class="konten"><div class="konto"><span class="dot" style="background:var(--spar)"></span><div><div class="name">In Töpfen</div></div><div class="betrag">${eur(t.inToepfen)}</div></div>
      <div class="konto"><span class="dot" style="background:var(--spar)"></span><div><div class="name">Nicht in Töpfen</div></div><div class="betrag">${eur(t.freiHeute)}</div></div></div></section>`;
  }
  return '';
}

// ---- Extras-Seiten (Etappe A: was es schon gibt, sonst Hinweis)
function extraSeite(id) {
  const a = ansicht(); const m = a.haupt;
  let inhalt = '';
  if (id === 'offene-punkte') inhalt = (m.offenePunkte || []).length ? `<ul class="offen-liste">${m.offenePunkte.map((p) => `<li><b>${h(p.thema || p.nr || '')}</b> ${h(p.text)} <span class="hint">${h(p.status || '')}</span></li>`).join('')}</ul>` : '<p class="hint">Keine offenen Punkte.</p>';
  else if (id === 'soll-ueberweisung') inhalt = m.soll.length ? `<div class="konten">${m.soll.map((s) => `<div class="konto"><span class="dot" style="background:var(--gem)"></span><div><div class="name">${monatsname(s.monat, true)}</div><div class="sub">${h(s.von)} → ${h(s.nach)}</div></div><div class="betrag">${eur(s.soll)}</div></div>`).join('')}</div>` : '<p class="hint">Noch keine Soll-Überweisungen eingetragen.</p>';
  else if (id === 'aufteilung' && z.bereich === 'gemeinsam') {
    const a2 = z.d.gemeinsam.m.aufteilung?.[z.monat];
    inhalt = a2 ? `<div class="konten">${Object.entries(a2).map(([k, v]) => `<div class="konto"><span></span><div><div class="name">${h(k)}</div></div><div class="betrag">${Math.abs(v) < 1.0001 && /anteil/i.test(k) ? Math.round(v * 100) + ' %' : eur(v)}</div></div>`).join('')}</div>` : '<p class="hint">Für diesen Monat gibt es noch keine Aufteilung.</p>';
  } else inhalt = `<p class="hint">Diese Seite kommt mit der nächsten Etappe in die neue App. Bis dahin steht alles in der Excel-Datei.</p>${haupt()?.webUrl ? `<p><a class="knopf leise" href="${h(haupt().webUrl)}" target="_blank" rel="noopener">Excel-Datei öffnen</a></p>` : ''}`;
  const dlg = dialog();
  dlg.innerHTML = `<h2>${h(FUNKTIONEN[id].titel)}</h2>${inhalt}<div class="knoepfe"><button class="knopf leise" id="dlg-zu">Schließen</button></div>`;
  dlg.showModal(); $('#dlg-zu').onclick = () => dlg.close();
}

// ---- Vorschau
function vorschauSeite(a) {
  const v = E.vorschau(a.haupt, heute(), heute().slice(0, 7));
  return `<section><h2>Vorschau</h2><p class="hint">Was in den nächsten 12 Monaten noch kommt und abgeht (Plan minus schon gebucht). Limits werden nicht abgezogen.</p>
    <div class="konten" style="margin-top:12px">${v.map((x) => `<div class="konto"><span class="dot" style="background:${x.luecke ? 'var(--warn)' : 'var(--gem)'}"></span>
      <div><div class="name">${monatsname(x.monat, true)}</div><div class="sub">+${eur(x.einnahmen)} · −${eur(x.kosten)}${x.sparraten ? ` · Töpfe −${eur(x.sparraten)}` : ''}${x.soll ? ` · Überweisungen ${eur(x.soll)}` : ''}</div></div>
      <div class="betrag ${x.freiEnde < 0 ? 'neg' : ''}">${eur(x.freiEnde)}<div class="sub">${x.ergebnis >= 0 ? '+' : ''}${eur(x.ergebnis)}</div></div></div>`).join('')}</div></section>`;
}

// ---- Buchungen
function buchungenSeite() {
  const q = z.suche.trim().toLowerCase();
  const alle = buchungenAnsicht().filter((b) => (q ? [b.nr, b.von, b.nach, b.posten, b.beschreibung, b.kommentar, b.kategorie, String(b.betrag).replace('.', ',')].join(' ').toLowerCase().includes(q)
    : b.giltFuer === z.monat || (b.datum || '').slice(0, 7) === z.monat)).sort((x, y) => (y.datum || '').localeCompare(x.datum || '') || String(y.nr).localeCompare(String(x.nr)));
  const eigene = new Set(eigeneKonten().map((k) => k.name));
  return `<section><h2>Buchungen</h2><input class="suche" id="suche" type="search" placeholder="Suchen (alle Monate)" value="${h(z.suche)}">
    ${alle.length ? `<div class="buchungen">${alle.map((b) => {
      const plus = eigene.has(b.nach) && !eigene.has(b.von);
      const offen = z.offenNr === b._datei + b.nr;
      return `<div class="bu-wrap"><button class="bu" data-bu="${h(b._datei + '|' + b.nr)}" aria-expanded="${offen}"><div><div class="t">${h(b.posten || b.beschreibung || (/^extern$/i.test(plus ? b.von : b.nach) ? b.kategorie : null) || (plus ? b.von : b.nach) || b.art)}</div>
        <div class="r">${datumKurz(b.datum)} · ${h(b.von)} → ${h(b.nach)}${b.kategorie ? ` · <span class="chip">${h(b.kategorie)}</span>` : ''}${b.topf ? ` · Topf ${h(b.topf)}` : ''}</div></div>
        <div class="betrag ${plus ? '' : 'neg'}">${plus ? '+' : b.art === 'Umbuchung' || /Topf/.test(b.art) ? '' : '−'}${eur(b.betrag)}</div></button>
        ${offen ? `<div class="bu-aktion"><dl class="more">${[['Nr', b.nr], ['Art', b.art], ['Gilt für', b.giltFuer && monatsname(b.giltFuer, true)], ['Kategorie', b.kategorie], ['Bank-Kategorie', b.bankKategorie], ['Posten', b.posten], ['Topf', b.topf], ['Beschreibung', b.beschreibung], ['Kommentar', b.kommentar], ['Gezahlt von', b.gezahltVon], ['Quelle', b.quelle], ['Gegenstück', b.gegenstueck]].filter(([, v]) => v).map(([k, v]) => `<dt>${k}</dt><dd>${h(v)}</dd>`).join('')}</dl>
          <div class="knoepfe"><button class="knopf klein" data-aendern="${h(b._datei + '|' + b.nr)}">Bearbeiten</button><button class="knopf klein leise" data-loeschen="${h(b._datei + '|' + b.nr)}">Löschen</button></div></div>` : ''}</div>`;
    }).join('')}</div>` : `<p class="leer">${q ? 'Nichts gefunden.' : `Keine Buchungen im ${monatsname(z.monat)}.`}</p>`}</section>`;
}

// ---- Töpfe
function toepfeSeite(a) {
  if (!a.toepfe.length) return `<section><h2>Töpfe</h2><p class="leer">${z.bereich !== 'gemeinsam' && z.person === 'Natalie' ? 'Die Töpfe liegen auf dem Tagesgeld (Bereich Gemeinsam). Unter „Mehr“ kannst du sie hier dazuschalten.' : 'Noch keine Töpfe.'}</p></section>`;
  const summe = E.r2(a.toepfe.reduce((s, t) => s + t.stand, 0));
  return `<section><h2>Töpfe</h2><p class="hint">Zusammen ${eur(summe)}</p><div class="toepfe">${a.toepfe.map((t) => {
    const ziel = t.ziel ? Math.max(0, Math.min(1, t.stand / t.ziel)) : null;
    return `<button class="topf ${t.stand < 0 ? 'minus' : ''}" data-topf="${h(t.name)}"><div class="n">${h(t.name)}</div><div class="b">${eur(t.stand)}</div>
      <div class="s">${t.sparrate ? `${eur(t.sparrate)} im Monat` : ''}${t.ziel ? ` · Ziel ${eur(t.ziel)}${t.zielBis ? ' bis ' + monatsname(t.zielBis, true) : ''}` : ''}${t.frei === 'Ja' ? ' · zählt als frei' : ''}</div>
      ${ziel != null ? `<div class="bar"><span style="width:${Math.round(ziel * 100)}%"></span></div>` : ''}</button>`;
  }).join('')}</div></section>`;
}
function topfSeite(a) {
  const v = E.topfVerlauf(a.topfModell, z.topf, heute().slice(0, 7), heute());
  if (!v) { z.topf = null; return toepfeSeite(a); }
  return `<section><button class="link" id="topf-zurueck">‹ Alle Töpfe</button><h2>${h(v.topf.name)}</h2><p class="hero"><strong>${eur(v.topf.stand)}</strong></p>
    <p class="hint">${v.topf.sparrate ? `Sparrate ${eur(v.topf.sparrate)}` : ''}${v.topf.bemerkung ? ' · ' + h(v.topf.bemerkung) : ''}</p>
    <div class="buchungen">${v.eintraege.map((e) => `<div class="bu"><div><div class="t">${h(e.b.posten || e.b.beschreibung || (e.sparplan ? 'Anteil Sparplan' : e.b.art))}</div><div class="r">${datumKurz(e.b.datum)}${e.b.datum.slice(0, 4)} · danach ${eur(e.danach)}</div></div><div class="betrag ${e.wirkung < 0 ? 'neg' : ''}">${e.wirkung >= 0 ? '+' : ''}${eur(e.wirkung)}</div></div>`).join('')}
    <div class="bu"><div><div class="t">Anfangsbestand${v.topf.zugeteilt ? ' + automatisch zugeteilt' : ''}</div></div><div class="betrag">${eur((v.topf.anfang || 0) + v.topf.zugeteilt)}</div></div></div></section>`;
}

// ---- Mehr = Einstellungen
function mehrSeite() {
  const an = funktionenAn(); const vorher = new Set(VORHER_AKTIV[art()] || []);
  const liste = Object.entries(FUNKTIONEN).filter(([id]) => fuerBereich(id));
  const reihe = widgetsReihe();
  const dateien = [['gemeinsam', 'Gemeinsam'], ['tagesgeld', 'Tagesgeld'], [privat(), z.person === 'Kai' ? 'Mein Geld (Kai)' : 'Meins (Natalie)']].filter(([b]) => z.d[b]);
  const thema = (() => { try { return localStorage.getItem('ug.thema') || 'auto'; } catch { return 'auto'; } })();
  return `<section><h2>Funktionen</h2><p class="hint">${z.bereich === 'gemeinsam' ? 'Gilt für euch beide.' : 'Nur für dich.'} Neu eingeschaltete Widgets erscheinen unten in der Übersicht.</p>
    <div class="konten funktionen">${liste.map(([id, f]) => `<label class="konto"><input type="checkbox" data-funktion="${id}" ${an.has(id) ? 'checked' : ''}>
      <div><div class="name">${h(f.titel)}</div><div class="sub">${f.art === 'widget' ? 'Widget in der Übersicht' : 'Seite unter „Extras“'}${vorher.has(id) ? ' · <b>vorher aktiv</b>' : ''}</div></div><span></span></label>`).join('')}</div>
    <p style="margin-top:12px"><button class="knopf leise" id="wie-vorher">Wie vor dem Umbau</button></p></section>
    ${reihe.length > 1 ? `<section><h2>Widgets anordnen</h2><div class="konten">${reihe.map((id, i) => `<div class="konto"><span></span><div><div class="name">${h(FUNKTIONEN[id].titel)}</div></div>
      <div><button class="link" data-hoch="${i}" ${i === 0 ? 'disabled' : ''} aria-label="nach oben">▲</button> <button class="link" data-runter="${i}" ${i === reihe.length - 1 ? 'disabled' : ''} aria-label="nach unten">▼</button></div></div>`).join('')}</div><p class="hint">„Extras“ bleibt immer ganz unten.</p></section>` : ''}
    <section><h2>Verwaltung</h2><p class="hint">Kosten und Einnahmen, Konten, Kategorien, Limits und Regeln pflegt ihr vorerst direkt in der Excel-Datei. Die Bearbeitung in der App kommt mit der nächsten Etappe.</p>
      <div class="konten">${dateien.map(([b, n]) => `<a class="konto" href="${h(z.d[b].webUrl)}" target="_blank" rel="noopener"><span></span><div><div class="name">${n}</div><div class="sub">Excel-Datei öffnen</div></div><span>›</span></a>`).join('')}</div></section>
    <section><h2>Darstellung</h2><div class="seg" role="radiogroup">${[['auto', 'Automatisch'], ['hell', 'Hell'], ['dunkel', 'Dunkel']].map(([w, n]) => `<label><input type="radio" name="thema" value="${w}" ${thema === w ? 'checked' : ''}><span>${n}</span></label>`).join('')}</div></section>
    <section><h2>App</h2><p class="hint">Version ${VERSION}${z.test ? ' · Testbetrieb' : ''} · angemeldet als ${h(z.person || '')}</p>
      <p style="margin-top:12px"><button class="knopf leise" id="abmelden">Abmelden</button></p></section>`;
}

function themaSetzen() {
  let t = 'auto'; try { t = localStorage.getItem('ug.thema') || 'auto'; } catch {}
  if (t === 'auto') document.documentElement.removeAttribute('data-theme'); else document.documentElement.dataset.theme = t === 'hell' ? 'light' : 'dark';
}

// ---------------------------------------------------------------- Bedienung

function verdrahten() {
  document.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { z.tab = b.dataset.tab; z.topf = null; z.offenNr = null; zeichnen(); scrollTo(0, 0); }));
  document.querySelectorAll('[data-bereich]').forEach((b) => (b.onclick = () => { z.bereich = b.dataset.bereich === 'gemeinsam' ? 'gemeinsam' : 'privat'; z.topf = null; zeichnen(); }));
  const mv = (n) => { const d = new Date(Date.UTC(+z.monat.slice(0, 4), +z.monat.slice(5, 7) - 1 + n, 1)); z.monat = d.toISOString().slice(0, 7); zeichnen(); };
  $('#m-zurueck')?.addEventListener('click', () => mv(-1)); $('#m-vor')?.addEventListener('click', () => mv(1));
  $('#neu-anmelden')?.addEventListener('click', () => ms.anmelden());
  $('#eintragen')?.addEventListener('click', () => formular());
  const s = $('#suche'); if (s) s.oninput = () => { z.suche = s.value; const pos = s.selectionStart; zeichnen(); const n = $('#suche'); n.focus(); n.setSelectionRange(pos, pos); };
  document.querySelectorAll('[data-bu]').forEach((b) => (b.onclick = () => { const id = b.dataset.bu.replace('|', ''); z.offenNr = z.offenNr === id ? null : id; zeichnen(); }));
  document.querySelectorAll('[data-aendern]').forEach((b) => (b.onclick = () => { const [d, nr] = b.dataset.aendern.split('|'); formular(z.d[d].m.buchungen.find((x) => x.nr === nr), d); }));
  document.querySelectorAll('[data-loeschen]').forEach((b) => (b.onclick = () => { const [d, nr] = b.dataset.loeschen.split('|'); loeschen(d, nr); }));
  document.querySelectorAll('[data-topf]').forEach((b) => (b.onclick = () => { z.topf = b.dataset.topf; zeichnen(); scrollTo(0, 0); }));
  $('#topf-zurueck')?.addEventListener('click', () => { z.topf = null; zeichnen(); });
  document.querySelectorAll('[data-seite]').forEach((b) => (b.onclick = () => extraSeite(b.dataset.seite)));
  document.querySelectorAll('[data-funktion]').forEach((c) => (c.onchange = () => funktionSetzen(c.dataset.funktion, c.checked)));
  $('#wie-vorher')?.addEventListener('click', wieVorher);
  document.querySelectorAll('[data-hoch]').forEach((b) => (b.onclick = () => widgetSchieben(+b.dataset.hoch, -1)));
  document.querySelectorAll('[data-runter]').forEach((b) => (b.onclick = () => widgetSchieben(+b.dataset.runter, 1)));
  document.querySelectorAll('input[name=thema]').forEach((r) => (r.onchange = () => { try { localStorage.setItem('ug.thema', r.value); } catch {} themaSetzen(); }));
  $('#abmelden')?.addEventListener('click', () => { ms.abmelden(); try { localStorage.removeItem('ug.neu.aktiv'); localStorage.removeItem('ug.neu.person'); caches.delete(SPEICHER); } catch {} location.reload(); });
}

async function speichern(text, arbeit) {
  if (z.speichert) return;
  if (Object.values(z.d).some((x) => x?.gespeichert)) { z.fehler = 'Gerade wird der zuletzt geladene Stand gezeigt. Bitte warten, bis die Dateien aus OneDrive geladen sind.'; return zeichnen(); }
  if (z.beispiel) { z.fehler = 'Beispielmodus: Es wird nichts gespeichert.'; return zeichnen(); }
  z.speichert = text; zeichnen();
  try { await arbeit(); z.fehler = null; } catch (e) { z.fehler = e instanceof od.AnmeldungNoetig ? 'Die Anmeldung ist abgelaufen. Bitte neu anmelden.' : e.message; }
  finally { z.speichert = null; zeichnen(); }
}
const einstDatei = () => (z.bereich === 'gemeinsam' ? 'gemeinsam' : privat());
async function einstSpeichern(an, reihe) {
  const b = einstDatei(); const e = z.d[b];
  await speichern('Wird gespeichert …', async () => {
    await D.einstellungSetzen(e.ort, 'Funktionen', [...an].join(','));
    await D.einstellungSetzen(e.ort, 'Widgets', reihe.join(','));
    e.m.einstellungen.Funktionen = [...an].join(','); e.m.einstellungen.Widgets = reihe.join(',');
    neuLaden(b, e.eTag).then(() => zeichnen());
  });
}
function funktionSetzen(id, ein) {
  const an = funktionenAn(); const reihe = String(einst().Widgets || '').split(',').filter(Boolean);
  if (ein) { an.add(id); if (FUNKTIONEN[id].art === 'widget') { const i = reihe.indexOf(id); if (i >= 0) reihe.splice(i, 1); reihe.push(id); } } else an.delete(id);
  einstSpeichern(an, reihe);
}
function wieVorher() {
  const an = funktionenAn(); const reihe = String(einst().Widgets || '').split(',').filter(Boolean).filter((id) => an.has(id));
  for (const id of VORHER_AKTIV[art()] || []) { if (!an.has(id)) { an.add(id); if (FUNKTIONEN[id].art === 'widget') reihe.push(id); } }
  // Standard-Reihenfolge für die Widgets, die vorher an waren und noch nicht angeordnet sind
  einstSpeichern(an, reihe);
}
function widgetSchieben(i, d) {
  const reihe = widgetsReihe(); const j = i + d; if (j < 0 || j >= reihe.length) return;
  [reihe[i], reihe[j]] = [reihe[j], reihe[i]];
  einstSpeichern(funktionenAn(), reihe);
}

// ---------------------------------------------------------------- Eintragen / Ändern / Löschen

function eigeneKonten() {
  if (z.bereich === 'gemeinsam') return [...(z.d.gemeinsam?.m.konten || []), ...(z.d.tagesgeld?.m.konten || [])].filter((k) => k.eigen === 'Ja' && k.status !== 'geschlossen');
  return (z.d[privat()]?.m.konten || []).filter((k) => k.eigen === 'Ja' && k.status !== 'geschlossen');
}
function dateiVonKonto(name) {
  if (z.bereich !== 'gemeinsam') return privat();
  if (z.d.tagesgeld?.m.konten.some((k) => k.name === name && k.eigen === 'Ja')) return 'tagesgeld';
  if (z.d.gemeinsam?.m.konten.some((k) => k.name === name && k.eigen === 'Ja')) return 'gemeinsam';
  return null;
}

function formular(alt, altDatei) {
  const a = ansicht(); const m = a.haupt; const H = heute();
  const konten = eigeneKonten();
  const kat = [...new Set([...(m.kategorien || []).filter((k) => k.aktiv !== 'Nein').map((k) => k.name)])];
  const posten = [...new Set([...m.kosten, ...m.einnahmen].map((p) => p.posten).filter(Boolean))];
  const toepfe = (a.topfModell?.toepfe || []).map((t) => t.name);
  const namen = [...new Set([...m.konten.map((k) => k.name), ...buchungenAnsicht().flatMap((b) => [b.von, b.nach])].filter(Boolean))].sort();
  const x = alt || { art: 'Ausgabe', datum: H, giltFuer: H.slice(0, 7), von: konten[0]?.name, nach: '' };
  const kontoSeite = ['Einnahme'].includes(x.art) ? x.nach : x.von;
  const gegen = x.art === 'Einnahme' ? x.von : x.nach;
  const dlg = dialog();
  dlg.innerHTML = `<form method="dialog" id="bu-form"><h2>${alt ? `Buchung ${h(alt.nr)} ändern` : 'Buchung eintragen'}</h2>
    <div class="seg" style="grid-template-columns:repeat(3,1fr)" aria-label="Art">${LISTEN.Arten.map((r) => `<label><input type="radio" name="art" value="${r}" ${x.art === r ? 'checked' : ''}><span>${r.replace('In Topf legen', 'In Topf').replace('Aus Topf nehmen', 'Aus Topf')}</span></label>`).join('')}</div>
    <div class="zwei"><label class="feld"><span>Betrag</span><input name="betrag" inputmode="decimal" required value="${alt ? String(alt.betrag).replace('.', ',') : ''}" placeholder="0,00"></label>
      <label class="feld"><span>Datum</span><input name="datum" type="date" required value="${h(x.datum)}"></label></div>
    <div class="zwei"><label class="feld"><span>Konto</span><select name="konto">${konten.map((k) => `<option ${k.name === kontoSeite ? 'selected' : ''}>${h(k.name)}</option>`).join('')}</select></label>
      <label class="feld"><span id="gegen-titel">Gegenüber</span><input name="gegen" list="namen" value="${h(gegen || '')}" placeholder="z. B. Laden, Person, Konto"></label></div>
    <datalist id="namen">${namen.map((n) => `<option value="${h(n)}">`).join('')}</datalist>
    <div class="zwei"><label class="feld"><span>Kategorie</span><select name="kategorie"><option value="">–</option>${kat.map((k) => `<option ${k === x.kategorie ? 'selected' : ''}>${h(k)}</option>`).join('')}</select></label>
      <label class="feld"><span>Gilt für Monat</span><input name="giltFuer" type="month" value="${h(x.giltFuer || H.slice(0, 7))}"></label></div>
    <div class="zwei"><label class="feld"><span>Posten</span><select name="posten"><option value="">–</option>${posten.map((p) => `<option ${p === x.posten ? 'selected' : ''}>${h(p)}</option>`).join('')}</select></label>
      <label class="feld"><span>Topf</span><select name="topf"><option value="">–</option>${[...toepfe, ...(toepfe.length ? [SPARPLAN] : [])].map((t) => `<option ${t === x.topf ? 'selected' : ''}>${h(t)}</option>`).join('')}</select></label></div>
    <label class="feld"><span>Beschreibung</span><input name="beschreibung" value="${h(x.beschreibung || '')}"></label>
    <label class="feld"><span>Kommentar</span><input name="kommentar" value="${h(x.kommentar || '')}"></label>
    ${z.bereich === 'gemeinsam' ? `<label class="feld"><span>Gezahlt von</span><select name="gezahltVon"><option value="">–</option>${LISTEN['Gezahlt von'].map((p) => `<option ${p === x.gezahltVon ? 'selected' : ''}>${h(p)}</option>`).join('')}</select></label>` : ''}
    <p class="hint" id="vorschlag"></p>
    <div class="knoepfe"><button class="knopf leise" value="abbrechen" formnovalidate>Abbrechen</button><button class="knopf" value="ok" id="bu-ok">${alt ? 'Speichern' : 'Eintragen'}</button></div></form>`;
  dlg.showModal();
  const f = $('#bu-form');
  const aktual = () => {
    const ar = f.art.value; $('#gegen-titel').textContent = ar === 'Einnahme' ? 'Von wem' : ar === 'Umbuchung' ? 'Auf welches Konto' : /Topf/.test(ar) ? 'Gegenüber (leer)' : 'An wen';
    f.gegen.disabled = /Topf/.test(ar);
    if (!f.kategorie.value && ar === 'Ausgabe') { const v = E.kategorieVorschlag(m, { art: ar, nach: f.gegen.value, beschreibung: f.beschreibung.value }, kat); $('#vorschlag').textContent = v ? `Vorschlag Kategorie: ${v}` : ''; if (v && kat.includes(v)) f.kategorie.value = v; }
  };
  f.querySelectorAll('input[name=art]').forEach((r) => (r.onchange = aktual)); f.gegen.onchange = aktual; f.beschreibung.onchange = aktual; aktual();
  f.datum.onchange = () => { if (!alt) f.giltFuer.value = f.datum.value.slice(0, 7); };
  dlg.onclose = () => { dlg.onclose = null; if (dlg.returnValue !== 'ok') return; const daten = Object.fromEntries(new FormData(f)); buchungSpeichern(daten, alt, altDatei); };
}

/** Art aus Sicht einer Datei: In Gemeinsam bleibt die gewählte Art (Sicht des Haushalts), sonst zählt, welche Seite ein eigenes Konto der Datei ist. */
function artFuerDatei(b, datei) {
  if (datei === 'gemeinsam' || /Topf/.test(b.art)) return b.art;
  const eigen = new Set(z.d[datei].m.konten.filter((k) => k.eigen === 'Ja').map((k) => k.name));
  const v = eigen.has(b.von), n = eigen.has(b.nach);
  return v && n ? 'Umbuchung' : n ? 'Einnahme' : v ? 'Ausgabe' : b.art;
}

function buchungAus(d) {
  const betrag = Number(String(d.betrag).replace(/\./g, '').replace(',', '.'));
  if (!(betrag > 0)) throw new Error('Bitte einen Betrag größer 0 eintragen.');
  const ar = d.art; const konto = d.konto; const gegen = (d.gegen || '').trim() || 'Extern';
  const von = ar === 'Einnahme' ? gegen : konto; const nach = ar === 'Einnahme' ? konto : /Topf/.test(ar) ? konto : gegen;
  return { datum: d.datum, giltFuer: d.giltFuer || d.datum.slice(0, 7), art: ar, von, nach, betrag, kategorie: d.kategorie || null, posten: d.posten || null, topf: d.topf || null,
    beschreibung: d.beschreibung || null, kommentar: d.kommentar || null, gezahltVon: d.gezahltVon || null };
}

function buchungSpeichern(daten, alt, altDatei) {
  let b; try { b = buchungAus(daten); } catch (e) { z.fehler = e.message; return zeichnen(); }
  speichern(alt ? 'Änderung wird gespeichert …' : 'Buchung wird eingetragen …', async () => {
    // Welche Dateien sind betroffen? Gemeinsam: Gemeinsam und/oder Tagesgeld (Überweisung dazwischen steht in beiden)
    const ziele = [...new Set([dateiVonKonto(b.von), dateiVonKonto(b.nach)].filter(Boolean))];
    if (!ziele.length) throw new Error('Mindestens Von oder Nach muss ein eigenes Konto sein.');
    if (alt) {
      const zeilen = [[altDatei, alt]];
      if (alt.gegenstueck) { const [dn, nr] = alt.gegenstueck.split(' '); const dd = dn.toLowerCase(); const g = z.d[dd]?.m.buchungen.find((x) => x.nr === nr); if (g) zeilen.push([dd, g]); }
      for (const [dd, x] of zeilen) {
        const neu = { ...x, ...b, art: artFuerDatei(b, dd), nr: x.nr, quelle: x.quelle, importId: x.importId, gegenstueck: x.gegenstueck, bankKategorie: x.bankKategorie, topf: dd === 'gemeinsam' && z.d.tagesgeld ? null : b.topf };
        const eTag = z.d[dd].eTag; await D.zeileAendern(z.d[dd].ort, 'Buchungen', x._zeile, x.nr, neu); await neuLaden(dd, eTag);
      }
      return;
    }
    const alle = ziele.flatMap((d) => z.d[d].m.buchungen);
    const nr = E.neueNr({ buchungen: alle });
    for (const d of ziele) {
      const anderes = ziele.find((x) => x !== d);
      const zeile = { ...b, art: artFuerDatei(b, d), nr, quelle: 'manuell', topf: d === 'gemeinsam' && ziele.includes('tagesgeld') ? null : b.topf,
        gegenstueck: anderes ? `${anderes === 'tagesgeld' ? 'Tagesgeld' : 'Gemeinsam'} ${nr}` : null };
      const eTag = z.d[d].eTag;
      try { await D.zeileNeu(z.d[d].ort, 'Buchungen', zeile, { doppelt: D.gleicheBuchung(b) }); }
      catch (e) {
        if (!(e instanceof od.Doppelt)) throw e;
        if (!(await frage(`Diese Buchung steht schon drin (${e.treffer.map((t) => t.nr).join(', ')}). Trotzdem eintragen?`, 'Trotzdem eintragen'))) return;
        await D.zeileNeu(z.d[d].ort, 'Buchungen', zeile);
      }
      z.d[d].m.buchungen.push({ ...zeile, _zeile: -1 });
      neuLaden(d, eTag).then(() => zeichnen());
    }
  });
}

/** Ja/Nein-Frage als App-Dialog (blockiert den Browser nicht wie confirm()). */
function frage(text, ja = 'Ja') {
  return new Promise((ok) => {
    const dlg = dialog();
    dlg.innerHTML = `<p>${h(text)}</p><div class="knoepfe"><button class="knopf leise" id="f-nein">Abbrechen</button><button class="knopf" id="f-ja">${h(ja)}</button></div>`;
    dlg.showModal();
    $('#f-nein').onclick = () => { dlg.close(); ok(false); };
    $('#f-ja').onclick = () => { dlg.close(); ok(true); };
  });
}

function loeschen(d, nr) {
  const x = z.d[d].m.buchungen.find((b) => b.nr === nr); if (!x) return;
  const dlg = dialog();
  dlg.innerHTML = `<h2>Buchung ${h(nr)} löschen?</h2><p>${h(x.posten || x.beschreibung || '')} · ${eur(x.betrag)} am ${datumKurz(x.datum)}${x.datum.slice(0, 4)}</p>
    ${x.gegenstueck ? `<p class="hint">Sie steht auch in der Datei ${h(x.gegenstueck.split(' ')[0])} und wird dort ebenfalls gelöscht.</p>` : ''}
    ${x.quelle === 'CSV' ? '<p class="hint">Sie kam aus einem Kontoauszug. Beim nächsten Abgleich wird die Bankzeile wieder als „neu“ angeboten.</p>' : ''}
    <div class="knoepfe"><button class="knopf leise" id="nein">Abbrechen</button><button class="knopf" id="ja">Löschen</button></div>`;
  dlg.showModal(); $('#nein').onclick = () => dlg.close();
  $('#ja').onclick = () => { dlg.close(); speichern('Wird gelöscht …', async () => {
    const zeilen = [[d, x]];
    if (x.gegenstueck) { const [dn, n2] = x.gegenstueck.split(' '); const dd = dn.toLowerCase(); const g = z.d[dd]?.m.buchungen.find((y) => y.nr === n2); if (g) zeilen.push([dd, g]); }
    for (const [dd, y] of zeilen) { const eTag = z.d[dd].eTag; await D.zeileLeeren(z.d[dd].ort, 'Buchungen', y._zeile, y.nr); await neuLaden(dd, eTag); }
    z.offenNr = null;
  }); };
}
