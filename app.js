// Handy-App „Unser Geld“ – zeigt die gemeinsame Buchhaltung an (nur lesen).
import { leseBuchhaltung } from './kern/lesen.js';
import { monatsansicht, monate, monatsname, monatPlus, monatsende, eur, AUSGLEICH, istAusgleich, ausgleichGrund, topfVerlauf, topfStandVor } from './kern/rechnen.js';
import * as ms from './anmeldung.js';
import * as od from './onedrive.js';
import { verstehen } from './kern/verstehen.js';
import * as meins from './meins.js';
import * as hilfe from './hilfe.js';

const VERSION = '0.13.0';
const app = document.getElementById('app');
const dateiInput = document.getElementById('datei');
const menu = document.getElementById('menu');

// ---------- kleine Helfer
const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const heute = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
const datumKurz = (iso) => (iso ? `${Number(iso.slice(8, 10))}.${Number(iso.slice(5, 7))}.` : '');
const vorzeichen = (n) => (n > 0 ? '+' : '') + eur(n);
const merken = (k, v) => { try { localStorage.setItem('ug.' + k, v); } catch {} };
const holen = (k) => { try { return localStorage.getItem('ug.' + k); } catch { return null; } };
const ICON = {
  links: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M15 5l-7 7 7 7"/></svg>',
  rechts: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M9 5l7 7-7 7"/></svg>',
  menu: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="5" r="1.4" fill="currentColor"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/><circle cx="12" cy="19" r="1.4" fill="currentColor"/></svg>',
  haken: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  pfeil: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>',
  zeichen: '<svg class="zeichen" viewBox="0 0 72 72" aria-hidden="true"><rect width="72" height="72" rx="18" fill="var(--surface)"/><circle cx="29" cy="36" r="15" fill="var(--kai)" opacity=".9"/><circle cx="43" cy="36" r="15" fill="var(--nat)" opacity=".9"/><path d="M36 23.2a15 15 0 010 25.6 15 15 0 010-25.6z" fill="var(--gem)"/></svg>',
};
const TABS = [
  ['uebersicht', 'Übersicht', '<path d="M4 11l8-6 8 6v8a1 1 0 01-1 1h-4v-5H9v5H5a1 1 0 01-1-1z"/>'],
  ['ablauf', 'Ablauf', '<path d="M8 3v18"/><circle cx="8" cy="6" r="2" fill="var(--surface)"/><circle cx="8" cy="13" r="2" fill="var(--surface)"/><path d="M13 6h7M13 13h5M13 19h6"/>'],
  ['toepfe', 'Töpfe', '<path d="M5 9h14l-1.5 10a1 1 0 01-1 .9h-9a1 1 0 01-1-.9z"/><path d="M8 9V7a4 4 0 018 0v2"/>'],
  ['buchungen', 'Buchungen', '<path d="M5 6h14M5 12h14M5 18h9"/>'],
];

// ---------- Datei auf dem Gerät merken (IndexedDB)
function db() {
  return new Promise((ok, fail) => {
    const r = indexedDB.open('unser-geld', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('dateien');
    r.onsuccess = () => ok(r.result);
    r.onerror = () => fail(r.error);
  });
}
async function speichereDatei(eintrag) {
  try { const d = await db(); await new Promise((ok, f) => { const t = d.transaction('dateien', 'readwrite'); t.objectStore('dateien').put(eintrag, 'aktuell'); t.oncomplete = ok; t.onerror = () => f(t.error); }); } catch {}
}
async function ladeDatei() {
  try { const d = await db(); return await new Promise((ok) => { const r = d.transaction('dateien').objectStore('dateien').get('aktuell'); r.onsuccess = () => ok(r.result || null); r.onerror = () => ok(null); }); } catch { return null; }
}
async function loescheDatei() {
  try { const d = await db(); await new Promise((ok) => { const t = d.transaction('dateien', 'readwrite'); t.objectStore('dateien').delete('aktuell'); t.oncomplete = ok; t.onerror = ok; }); } catch {}
}

// ---------- Gespeichert, aber von OneDrive noch nicht ausgeliefert
// Nach dem Schreiben über eine Excel-Sitzung liefert OneDrive die Datei oft noch einige Sekunden im alten Stand.
// Damit die Liste trotzdem sofort stimmt (und niemand dieselbe Buchung noch einmal einträgt), merkt sich die App
// ihre eigenen Änderungen und legt sie über den geladenen Stand, bis die Datei sie selbst enthält (höchstens 10 Min.).
const AUSSTEHEND_MAX = 10 * 60 * 1000;
let ausstehend = (() => { try { return (JSON.parse(holen('ausstehend')) || []).filter((a) => Date.now() - a.seit < AUSSTEHEND_MAX); } catch { return []; } })();
const ausstehendMerken = () => merken('ausstehend', JSON.stringify(ausstehend));
const gleicheBuchung = (x, b) => x.datum === b.datum && x.posten === b.posten && x.von === b.von && x.nach === b.nach
  && Math.abs(Number(x.betrag) - Number(b.betrag)) < 0.005 && String(x.gezahltVon || '') === String(b.gezahltVon || '');
/** Legt die noch nicht ausgelieferten eigenen Änderungen über eine Buchungsliste. Erledigte fallen weg. */
function ausstehendAnwenden(liste) {
  ausstehend = ausstehend.filter((a) => {
    if (Date.now() - a.seit > AUSSTEHEND_MAX) return false;
    const x = liste.find((b) => b.nr === a.nr);
    return a.typ === 'loeschen' ? Boolean(x) : !(x && gleicheBuchung(x, a.b));
  });
  ausstehendMerken();
  let l = liste.slice();
  for (const a of ausstehend) {
    const i = l.findIndex((b) => b.nr === a.nr);
    if (a.typ === 'loeschen') { if (i >= 0) l.splice(i, 1); }
    else if (i >= 0) l[i] = a.b; else l.push(a.b);
  }
  return l;
}
let nachpruefTimer = null;
function nachpruefenPlanen() {
  if (!ausstehend.length || nachpruefTimer) return;
  nachpruefTimer = setTimeout(() => { nachpruefTimer = null; ausOneDrive({ immer: true }); }, 6000);
}

// ---------- Zustand
const z = { modell: null, datei: null, monat: null, bereich: holen('bereich') || 'gemeinsam', tab: holen('tab') || 'uebersicht', filter: 'alle', status: null, laedt: false, geprueft: 0,
  topf: null, zeitraum: '12', eigenVon: null, eigenBis: null, topfScroll: 0 };

// ---------- Start
async function start() {
  hilfe.starten();
  let fehler = null;
  if (ms.aktiv()) {
    try {
      if (await ms.rueckkehrVerarbeiten()) {
        const me = await od.ich().catch(() => null);
        if (me) ms.kontoMerken(me.mail || me.userPrincipalName || me.displayName);
      }
    } catch (e) { fehler = e.message; }
  }
  if (location.hash === '#kai') {
    z.bereich = 'meins'; merken('bereich', 'meins');
    const l = document.querySelector('link[rel=manifest]'); if (l) l.href = 'kai.webmanifest';
    document.title = 'Mein Geld';
  }
  if (location.hash === '#pc') {
    const pc = await import('./pc.js');
    await pc.starte(app);
    return;
  }
  if (history.state?.topf) { try { history.replaceState(null, ''); } catch {} }
  const gespeichert = await ladeDatei();
  if (gespeichert) {
    try { await uebernehmen(gespeichert, false); } catch (e) { fehler = 'Die gespeicherte Datei ließ sich nicht lesen: ' + e.message; }
  }
  if (ms.aktiv() && ms.angemeldet()) { await ausOneDrive(); return; }
  if (!z.modell) zeigeStart(fehler);
  else if (fehler) { z.status = { art: 'warn', text: fehler }; zeichne(); }
}

async function uebernehmen(eintrag, speichern) {
  const modell = await leseBuchhaltung(eintrag.bytes);
  modell.buchungen = ausstehendAnwenden(modell.buchungen);
  z.modell = modell; z.datei = eintrag;
  const liste = monate(modell);
  const jetzt = heute().slice(0, 7);
  if (!z.monat || !liste.includes(z.monat)) z.monat = liste.includes(jetzt) ? jetzt : (liste.find((m) => m > jetzt) || liste[liste.length - 1]);
  if (speichern) await speichereDatei(eintrag);
  zeichne();
  nachpruefenPlanen();
}

/** Lädt die Datei aus OneDrive, wenn sie sich geändert hat. Läuft schon ein Laden, wird danach noch einmal geladen. */
async function ausOneDrive({ immer = false } = {}) {
  if (z.laedt) { z.nochmal = true; return; }
  do { z.nochmal = false; await einmalAusOneDrive({ immer: immer || z.nochmal }); } while (z.nochmal);
}
async function einmalAusOneDrive({ immer = false } = {}) {
  z.laedt = true; z.geprueft = Date.now();
  if (z.modell) { z.status = { art: 'info', text: 'Prüfe OneDrive auf Änderungen …' }; zeichne(); }
  try {
    const gefunden = await od.dateiFinden();
    if (!gefunden) {
      z.laedt = false;
      if (z.modell) { z.status = { art: 'warn', text: 'Die Datei wurde in OneDrive nicht mehr gefunden (umbenannt oder verschoben?). Ihr seht den zuletzt geladenen Stand.' }; zeichne(); }
      else zeigeNichtGefunden();
      return;
    }
    await ladeGefunden(gefunden, immer);
    if (z.bereich === 'meins' && meins.verfuegbar()) meins.laden({ immer: true });
  } catch (e) {
    z.laedt = false;
    if (e instanceof od.AnmeldungNoetig) {
      if (z.modell) { z.status = { art: 'warn', text: 'Die Microsoft-Anmeldung ist abgelaufen. Ihr seht den zuletzt geladenen Stand.', knopf: ['anmelden', 'Neu anmelden'] }; zeichne(); }
      else zeigeStart();
      return;
    }
    const offline = e instanceof od.KeinNetz;
    if (z.modell) { z.status = { art: offline ? 'info' : 'warn', text: offline ? 'Gerade keine Verbindung. Ihr seht den zuletzt geladenen Stand.' : 'OneDrive meldet einen Fehler (' + e.message + '). Ihr seht den zuletzt geladenen Stand.', knopf: ['aktualisieren', 'Noch mal versuchen'] }; zeichne(); }
    else zeigeStart('OneDrive war nicht erreichbar: ' + e.message);
    return;
  }
  z.laedt = false;
}

async function ladeGefunden({ ort, item }, immer) {
  const meta = { name: item.name, geaendert: item.lastModifiedDateTime ? Date.parse(item.lastModifiedDateTime) : null,
    geaendertVon: item.lastModifiedBy?.user?.displayName || null, eTag: item.eTag, quelle: 'onedrive', webUrl: item.webUrl || null };
  if (!immer && z.datei?.eTag && z.datei.eTag === item.eTag) {
    z.datei = { ...z.datei, ...meta }; z.status = null; zeichne(); return;
  }
  const bytes = await od.herunterladen(ort);
  z.status = null;
  await uebernehmen({ ...meta, geladen: new Date().toISOString(), bytes }, true);
}

function zeigeStart(fehler) {
  const msKnopf = ms.aktiv() ? `<div><button class="knopf" id="anmelden">Mit Microsoft anmelden</button></div>` : '';
  app.innerHTML = `<main class="wrap start">
    ${ICON.zeichen}
    <h1>Unser Geld</h1>
    <p>Hier seht ihr, wo euer gemeinsames Geld gerade liegt, wer was überweist und was im Monat noch abgeht. Alles kommt aus eurer Excel-Datei.</p>
    ${ms.aktiv() ? `${msKnopf}
      <p class="hint klein">Meldet euch mit dem Microsoft-Konto an, mit dem ihr die Datei in OneDrive öffnet. Die App lädt sie dann selbst und bleibt auf dem neuesten Stand.</p>
      <div style="margin-top:24px"><label class="knopf leise" for="datei">Stattdessen Datei vom Handy öffnen</label></div>`
    : `<div><label class="knopf" for="datei">Excel-Datei öffnen</label></div>
      <p class="hint klein">Wählt „Buchhaltung_Kai_Natalie.xlsx“ aus, zum Beispiel über OneDrive in der Dateien-App. Die Datei bleibt auf diesem Handy und wird nirgendwohin geschickt.</p>`}
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
  </main>`;
}

function zeigeNichtGefunden(fehler) {
  app.innerHTML = `<main class="wrap start">
    ${ICON.zeichen}
    <h1>Datei nicht gefunden</h1>
    <p>Mit ${h(ms.gemerktesKonto() || 'diesem Konto')} hat die App „Buchhaltung_Kai_Natalie.xlsx“ in OneDrive nicht gefunden.</p>
    <p class="hint">Kai kann in OneDrive bei der Datei auf „Teilen“ › „Link kopieren“ gehen und euch den Link schicken. Fügt ihn hier ein:</p>
    <form id="link-form" style="display:grid;gap:12px">
      <input id="link" type="url" inputmode="url" placeholder="https://1drv.ms/…" required
        style="height:52px;border-radius:14px;border:1px solid var(--line);background:var(--surface);color:var(--ink);padding:0 16px;font:inherit">
      <button class="knopf" type="submit">Datei über Link öffnen</button>
    </form>
    <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap"><button class="knopf leise" id="anmelden">Anderes Konto</button><label class="knopf leise" for="datei">Datei vom Handy</label></div>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
  </main>`;
}

document.addEventListener('submit', async (ev) => {
  if (ev.target.id !== 'link-form') return;
  ev.preventDefault();
  try { await ladeGefunden(await od.ueberLink(document.getElementById('link').value), true); }
  catch (e) { if (e instanceof od.AnmeldungNoetig) ms.anmelden(); else zeigeNichtGefunden(e.message); }
});

document.addEventListener('submit', async (ev) => {
  if (!/^[mk]-/.test(ev.target.id || '')) return;
  ev.preventDefault();
  await meins.absenden(ev.target);
});

document.addEventListener('change', (ev) => {
  if (/^k-/.test(ev.target.id || '') && ev.target.id !== 'k-suche') { meins.kaiAenderung(ev.target); return; }
  if (ev.target.id === 'm-csv') { const f = ev.target.files?.[0]; ev.target.value = ''; if (f) meins.csvGewaehlt(f); return; }
  if (ev.target.id === 'topf-von') { z.eigenVon = ev.target.value; if (z.eigenBis && z.eigenBis < z.eigenVon) z.eigenBis = z.eigenVon; zeichne(); }
  if (ev.target.id === 'topf-bis') { z.eigenBis = ev.target.value; if (z.eigenVon && z.eigenVon > z.eigenBis) z.eigenVon = z.eigenBis; zeichne(); }
});

document.addEventListener('input', (ev) => { if (ev.target.id === 'k-suche') meins.kaiAenderung(ev.target); });

document.addEventListener('visibilitychange', () => {
  if (document.body.classList.contains('pc-modus')) return;
  if (document.visibilityState === 'visible' && ms.aktiv() && ms.angemeldet() && Date.now() - z.geprueft > 5 * 60 * 1000) ausOneDrive();
});

window.addEventListener('hashchange', () => location.reload());

dateiInput.addEventListener('change', async () => {
  const f = dateiInput.files?.[0];
  dateiInput.value = '';
  if (!f) return;
  try {
    const bytes = new Uint8Array(await f.arrayBuffer());
    await uebernehmen({ name: f.name, geladen: new Date().toISOString(), geaendert: f.lastModified || null, quelle: 'geraet', bytes }, true);
    if (menu.open) menu.close();
  } catch (e) {
    zeigeStart('Diese Datei ließ sich nicht lesen. Ist es die richtige Excel-Datei? (' + e.message + ')');
  }
});

// ---------- Zeichnen
function zeichne() {
  if (z.topf && !z.modell.toepfe.some((t) => t.name === z.topf)) z.topf = null;
  if (z.topf) { zeichneTopf(); return; }
  if (z.bereich === 'meins' && meins.verfuegbar()) { meins.zeichne(); return; }
  meins.thema(false);
  const v = monatsansicht(z.modell, z.monat, heute());
  app.innerHTML = `
  ${kopf()}
  <main class="wrap">
    ${z.status ? `<div class="hinweis ${z.status.art}">${h(z.status.text)}${z.status.knopf ? ` <button class="link" id="${z.status.knopf[0]}">${h(z.status.knopf[1])}</button>` : ''}</div>` : ''}
    <div class="view ${z.tab === 'uebersicht' ? 'on' : ''}" role="tabpanel">${uebersicht(v)}</div>
    <div class="view ${z.tab === 'ablauf' ? 'on' : ''}" role="tabpanel">${ablauf(v)}</div>
    <div class="view ${z.tab === 'toepfe' ? 'on' : ''}" role="tabpanel">${toepfe(v)}</div>
    <div class="view ${z.tab === 'buchungen' ? 'on' : ''}" role="tabpanel">${buchungen(v)}</div>
  </main>
  ${kannEintragen() ? `<button class="fab" id="eintragen-knopf" aria-label="Buchung eintragen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Eintragen</button>` : ''}
  <nav><div class="wrap" role="tablist">${TABS.map(([id, name, pfad]) =>
    `<button role="tab" aria-selected="${z.tab === id}" data-tab="${id}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${pfad}</svg>${name}</button>`).join('')}</div></nav>`;
}

/** Kopf mit Monatswahl, Menü und (wenn angemeldet) Umschalter Gemeinsam/Meins. */
function kopf() {
  const liste = monate(z.modell);
  const i = liste.indexOf(z.monat);
  const [jahr] = z.monat.split('-');
  const jetzt = heute().slice(0, 7);
  const wann = z.monat === jetzt ? 'dieser Monat' : z.monat < jetzt ? 'vorbei' : 'Vorschau';
  const umschalter = meins.verfuegbar() ? `<div class="bereich-wahl" role="tablist" aria-label="Bereich">
      <button role="tab" aria-selected="${z.bereich !== 'meins'}" data-bereich="gemeinsam">Gemeinsam</button>
      <button role="tab" aria-selected="${z.bereich === 'meins'}" data-bereich="meins">Meins</button></div>` : '';
  return `<header><div class="wrap"><div class="monat">
    <button data-monat="${i > 0 ? liste[i - 1] : ''}" aria-label="Vorheriger Monat" ${i > 0 ? '' : 'disabled'}>${ICON.links}</button>
    <div class="mitte"><h1>${h(monatsname(z.monat))} ${jahr}</h1><small>${wann}</small></div>
    <div style="display:flex">
      <button data-monat="${i < liste.length - 1 ? liste[i + 1] : ''}" aria-label="Nächster Monat" ${i < liste.length - 1 ? '' : 'disabled'}>${ICON.rechts}</button>
      <button class="menu" id="menu-knopf" aria-label="Datei und Einstellungen">${ICON.menu}</button>
    </div>
  </div>${umschalter}</div></header>`;
}

function kontoFarbe(name) {
  const k = z.modell.konten.find((x) => x.name === name);
  if (!k) return 'c-ext';
  if (k.rolle === 'Sparen') return 'c-spar';
  if (k.rolle === 'gemeinsam') return /kredit|visa/i.test((k.art || '') + k.name) ? 'c-visa' : 'c-gem';
  if (k.rolle === 'privat') return k.inhaber === 'Natalie' ? 'c-nat' : 'c-kai';
  return 'c-ext';
}
const farbVar = { 'c-gem': 'var(--gem)', 'c-visa': 'var(--visa)', 'c-spar': 'var(--spar)', 'c-kai': 'var(--kai)', 'c-nat': 'var(--nat)', 'c-ext': 'var(--muted)' };

function uebersicht(v) {
  const gem = v.konten.find((k) => k.name === v.ablauf.konto);
  const stand = v.laufend ? gem?.standJetzt : v.vergangen ? gem?.standMonatsende : v.ablauf.ende;
  let satz, chip;
  if (v.laufend) {
    satz = `Auf dem Gemeinschaftskonto liegen gerade <strong class="${stand < 0 ? 'neg' : ''}">${eur(stand)}</strong>.`;
    chip = `<span class="check offen">Berechnet aus den Buchungen bis heute</span>`;
  } else if (v.vergangen) {
    satz = `Ende ${h(v.name)} lagen auf dem Gemeinschaftskonto <strong class="${stand < 0 ? 'neg' : ''}">${eur(stand)}</strong>.`;
    chip = gem?.bankMonatsende == null ? '<span class="check offen">Bankstand noch nicht eingetragen</span>'
      : Math.abs(gem.differenz) < 0.005 ? `<span class="check">${ICON.haken}Stimmt mit der Bank überein</span>`
        : `<span class="check warn">Bank: ${eur(gem.bankMonatsende)}, Differenz ${vorzeichen(gem.differenz)}</span>`;
  } else {
    satz = `Ende ${h(v.name)} liegen auf dem Gemeinschaftskonto voraussichtlich <strong class="${stand < 0 ? 'neg' : ''}">${eur(stand)}</strong>.`;
    chip = '<span class="check offen">Wenn alles wie geplant kommt und abgeht</span>';
  }

  const abw = v.abweichungen.length ? `<div class="hinweis warn"><b>Die App rechnet hier anders als Excel:</b><br>${v.abweichungen.map((a) =>
    `${h(a.was)}: App ${eur(a.app)}, Excel ${eur(a.excel)}`).join('<br>')}<br>Bitte die Datei einmal in Excel öffnen und speichern, dann neu laden.</div>` : '';

  const zeigeStand = (k) => (v.vergangen ? k.standMonatsende : k.standJetzt);
  const positiv = v.konten.map((k) => Math.max(0, zeigeStand(k)));
  const gesamt = positiv.reduce((s, x) => s + x, 0) || 1;
  const toepfeAnzahl = v.toepfe.length;
  const konten = `<section><h2>${v.zukunft ? 'Wo das Geld heute liegt' : v.vergangen ? `Wo das Geld Ende ${h(v.name)} lag` : 'Wo das Geld liegt'}</h2>
    <div class="strip" aria-hidden="true">${v.konten.map((k, i) => `<span style="flex:${positiv[i] / gesamt};background:${farbVar[kontoFarbe(k.name)]};min-width:${positiv[i] > 0 ? 4 : 0}px"></span>`).join('')}</div>
    <div class="konten">${v.konten.map((k) => {
      const sub = k.rolle === 'Sparen' ? `${toepfeAnzahl} Töpfe${Math.abs(v.nichtInToepfen) >= 0.005 ? `, davon ${eur(v.nichtInToepfen)} in keinem Topf` : ''}` : (k.bank || '') + (k.inhaber ? `, läuft auf ${k.inhaber}` : '');
      const s = zeigeStand(k);
      return `<div class="konto"><span class="dot" style="background:${farbVar[kontoFarbe(k.name)]}"></span><div><div class="name">${h(k.name)}</div><div class="sub">${h(sub)}</div></div><div class="betrag ${s < 0 ? 'neg' : ''}">${eur(s)}</div></div>`;
    }).join('')}</div>
    ${v.vergangen ? '<p class="hint" style="margin-top:10px">Stände am Monatsende laut Buchungen.</p>' : ''}
  </section>`;

  const ueb = v.ueberweisungen.length ? `<section><h2>Wer überweist was</h2><div class="ueb">${v.ueberweisungen.map((u) => {
    const status = u.erledigt ? `<span class="status ok">${ICON.haken} eingegangen</span>`
      : u.gebucht > 0 ? `<span class="status offen">${eur(u.gebucht)} da, ${eur(u.offen)} fehlen noch</span>` : '';
    return `<div class="zahlung ${u.erledigt ? 'erledigt' : ''}"><span class="who ${u.person === 'Kai' ? 'k' : 'n'}">${u.person[0]}</span>
      <div><div class="route">${h(u.person)} an ${h(u.nach)}</div><div class="bis">${h(u.hinweis || u.wann)}</div></div>
      <div class="betrag">${eur(u.betrag)}</div>${status}</div>`;
  }).join('')}</div>
  ${v.ausgleich.length ? `<p class="hint ausgleich-hinweis">Außerdem eingegangen: ${v.ausgleich.map((b) => `${eur(b.betrag)} ${h(b.posten || '')} (${h(b.grund)})`).join(', ')}. Zählt als Ausgleich, nicht als Überweisung.</p>` : ''}
  ${v.anteil ? `<div class="anteil">
    <div class="hint">So teilt ihr euch die Kosten im ${h(v.name)}, nach Einkommen</div>
    <div class="anteilbar" aria-hidden="true"><span style="flex:${v.anteil.anteilKai};background:var(--kai)"></span><span style="flex:${v.anteil.anteilNatalie};background:var(--nat)"></span></div>
    <div class="legend"><span><b style="color:var(--kai)">Kai ${(v.anteil.anteilKai * 100).toLocaleString('de-DE', { maximumFractionDigits: 2 })} %</b><br><span class="hint">trägt ${eur(v.anteil.traegtKai)}</span></span>
    <span style="text-align:right"><b style="color:var(--nat)">Natalie ${(v.anteil.anteilNatalie * 100).toLocaleString('de-DE', { maximumFractionDigits: 2 })} %</b><br><span class="hint">trägt ${eur(v.anteil.traegtNatalie)}</span></span></div>
    <p class="hint" style="margin-top:10px">Einkommen ${eur(v.anteil.einkommenKai)} und ${eur(v.anteil.einkommenNatalie)}. Gesamtbedarf ${eur(v.anteil.gesamtbedarf)}${v.anteil.schenkungen ? `, Mietgeschenk ${eur(v.anteil.schenkungen)} schon abgezogen` : ''}.</p>
  </div>` : ''}</section>` : '';

  const ende = (v.laufend || v.zukunft) ? `<section><h2>Ende ${h(v.name)}</h2>
    <div class="ende"><div><div style="font-weight:600">Gemeinschaftskonto, voraussichtlich</div><div class="hint">wenn alles wie geplant kommt und abgeht</div></div><div class="betrag ${v.ablauf.ende < 0 ? 'neg' : ''}">${eur(v.ablauf.ende)}</div></div>
    <p class="hint" style="margin-top:10px">Jeden Posten mit Tag zeigt der Reiter Ablauf.</p></section>` : '';

  const offen = v.offenePunkte.length ? `<section><h2>Noch zu klären</h2><ul class="offen-liste">${v.offenePunkte.map((p) =>
    `<li><b>${h(p.thema || 'Punkt ' + p.nr)}</b>${h(p.text)}</li>`).join('')}</ul></section>` : '';

  return `<div class="hero"><p>${satz}</p>${chip}</div>${abw}${konten}${ueb}${ende}${offen}`;
}

function ablauf(v) {
  const a = v.ablauf;
  const heuteTag = Number(v.heute.slice(8, 10));
  const vor = monatsname(monatPlus(v.monat, -1));
  let html = `<section><h2>${h(a.konto)} im ${h(v.name)}</h2>
    <p class="hint" style="margin-bottom:12px">Was wann reinkommt und abgeht, und was danach auf dem Konto ist.</p>
    <div class="legende"><span><i></i>gebucht</span><span><i class="d"></i>geplant</span><span><span style="display:inline-block;width:14px;height:4px;border-radius:2px;background:var(--warn);vertical-align:3px;margin-right:6px"></span>unter null</span></div>
    <div class="band">`;
  const zeile = (e, cls, tag, mon, titel, sub, betrag, danach) => `<div class="ev ${cls}" style="--seg:${danach < 0 ? 'var(--warn)' : 'var(--gem)'}">
      <div class="tag"><b>${tag}</b><span>${mon}</span></div><div class="rail"><span class="knot"></span></div>
      <div class="was"><div class="t">${titel}</div><div class="s">${sub}</div></div>
      <div class="zahlen">${betrag}<div class="st ${danach < 0 ? 'neg' : ''}">${e ? 'danach ' + eur(danach) : `<b style="color:var(--ink)">${eur(danach)}</b>`}</div></div></div>`;
  html += zeile(null, '', Number(a.startDatum.slice(8, 10)), vor.slice(0, 3), a.startIstPrognose ? `Stand Ende ${h(vor)}, voraussichtlich` : `Stand Ende ${h(vor)}`,
    a.startIstPrognose ? 'aus der Vorschau des Vormonats' : 'laut Buchungen', '', a.start);
  let markiert = !v.laufend;
  const mon = v.name.slice(0, 3);
  for (const e of a.ereignisse) {
    if (!markiert && (e.tag == null || e.tag > heuteTag || (e.datum && e.datum > v.heute))) {
      html += `<div class="heute"><span></span><div class="rail" style="height:28px"><span></span></div><span class="lbl">Heute, ${heuteTag}. ${h(v.name)}</span></div>`;
      markiert = true;
    }
    let sub;
    if (e.art === 'gebucht') {
      sub = [e.giltFuer && e.giltFuer !== v.monat ? `<span class="okt">gilt für ${h(monatsname(e.giltFuer))}</span>` : '', h(e.kommentar || (e.gegen ? (e.betrag > 0 ? 'von ' : 'an ') + e.gegen : ''))].filter(Boolean).join(' ');
    } else {
      sub = h(e.bemerkung && e.bemerkung.length <= 70 ? e.bemerkung : 'geplant');
    }
    const betrag = `<div class="b ${e.betrag > 0 ? 'plus' : ''}">${vorzeichen(e.betrag)}</div>`;
    html += zeile(e, (e.art === 'geplant' ? 'plan' : '') + (e.tag == null ? ' ohne' : ''), e.tag ?? '–', e.tag == null ? 'ohne Tag' : mon, h(e.titel), sub, betrag, e.danach);
  }
  if (!markiert) html += `<div class="heute"><span></span><div class="rail" style="height:28px"><span></span></div><span class="lbl">Heute, ${heuteTag}. ${h(v.name)}</span></div>`;
  html += `</div>
    <div class="ende" style="margin-top:8px"><div><div style="font-weight:600">Ende ${h(v.name)}${v.vergangen ? '' : ', voraussichtlich'}</div><div class="hint">${v.vergangen ? 'laut Buchungen' : 'wenn alles wie geplant kommt und abgeht'}</div></div><div class="betrag ${a.ende < 0 ? 'neg' : ''}">${eur(a.ende)}</div></div>
  </section>`;
  return html;
}

function toepfe(v) {
  const gruppen = new Map();
  for (const t of v.toepfe) { if (!gruppen.has(t.bereich)) gruppen.set(t.bereich, []); gruppen.get(t.bereich).push(t); }
  return `<section><h2>Töpfe auf dem Tagesgeld</h2>
    <div class="topfsumme">
      <div class="z"><span>Auf den Sparkonten</span><b>${eur(v.sparkonten)}</b></div>
      <div class="z"><span>Davon in Töpfen</span><b>${eur(v.toepfeSumme)}</b></div>
      <div class="z"><span>In keinem Topf</span><b style="color:${Math.abs(v.nichtInToepfen) < 0.005 ? 'var(--gem)' : 'var(--warn)'}">${eur(v.nichtInToepfen)}</b></div>
    </div>
    <p class="hint" style="margin-top:10px">Stand ${v.laufend ? 'heute' : `Ende ${h(v.name)}`}, laut Buchungen. Sparraten zusammen ${eur(v.toepfe.reduce((s, t) => s + t.sparrate, 0))} im Monat.</p>
    ${[...gruppen].map(([bereich, liste]) => `<div class="gruppe">${h(bereich || 'Ohne Bereich')}</div><div class="toepfe">${liste.map((t) =>
      `<button class="topf ${t.stand < 0 ? 'minus' : ''}" data-topf="${h(t.name)}" aria-label="${h(t.name)}: alle Buchungen ansehen"><span class="n">${h(t.name)}</span><span class="b">${eur(t.stand)}</span><span class="pf">${ICON.rechts}</span>
        <span class="s">${t.sparrate ? `${eur(t.sparrate)} im Monat` : 'keine feste Sparrate'}${t.bemerkung ? '. ' + h(t.bemerkung) : ''}</span></button>`).join('')}</div>`).join('')}
    <p class="hint" style="margin-top:12px">Topf antippen zeigt alle Buchungen dieses Topfs.</p>
  </section>`;
}

const datumLang = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
/** Detail-Zeilen einer Buchung (aufgeklappt). `extra` kommt vorne dazu. */
function details(b, extra = []) {
  return [...extra, ['Nummer', b.nr], ['Bezahlt am', datumLang(b.datum)], ['Gilt für', b.giltFuer ? monatsname(b.giltFuer, true) : ''],
    ['Art', b.art], ['Bereich', b.bereich], ['Topf', b.topf], ['Verwendungszweck', b.beschreibung], ['Kommentar', b.kommentar],
    ['Gezahlt von', istAusgleich(b) ? null : b.gezahltVon], ['Ausgleich', istAusgleich(b) ? ausgleichGrund(b) + ', wird nicht angerechnet' : null], ['Eingetragen', b.quelle === 'CSV' ? 'aus dem Bank-Export' : b.quelle]]
    .filter(([, w]) => w).map(([k, w]) => `<dt>${k}</dt><dd>${h(w)}</dd>`).join('');
}

// ---------- Ein Topf: alle Buchungen mit Zeitraum-Filter
const ZEITRAEUME = [['3', '3 Monate'], ['12', '12 Monate'], ['alles', 'Seit Beginn'], ['eigen', 'Zeitraum wählen']];
/** Bis wann der Topf gezeigt wird: wie im Reiter Töpfe (heute im laufenden Monat, sonst Monatsende). */
function topfBis() {
  const h0 = heute();
  return h0.slice(0, 7) === z.monat ? h0 : monatsende(z.monat);
}
function zeichneTopf() {
  const bis = topfBis();
  const vl = topfVerlauf(z.modell, z.topf, bis);
  const t = vl.topf;
  const bisMonat = bis.slice(0, 7);
  const startMonat = (vl.beginn || bis).slice(0, 7);
  // Zeitraum in Monaten [von, bisM]
  let von, bisM = bisMonat;
  if (z.zeitraum === 'alles') von = startMonat;
  else if (z.zeitraum === 'eigen') {
    von = z.eigenVon || monatPlus(bisMonat, -11);
    bisM = z.eigenBis || bisMonat;
    if (bisM > bisMonat) bisM = bisMonat;
    if (von > bisM) von = bisM;
  } else von = monatPlus(bisMonat, -(Number(z.zeitraum) - 1));
  const abgeschnitten = von <= startMonat;
  if (von < startMonat) von = startMonat;
  const vonDatum = abgeschnitten && vl.beginn ? vl.beginn : von + '-01';
  const bisDatum = bisM === bisMonat ? bis : monatsende(bisM);
  const imZeitraum = vl.eintraege.filter((e) => (abgeschnitten || e.b.datum >= vonDatum) && e.b.datum <= bisDatum);
  const standAnfang = abgeschnitten ? vl.anfang : topfStandVor(vl, vonDatum);
  const standEnde = imZeitraum.length ? imZeitraum[imZeitraum.length - 1].danach : standAnfang;
  const rein = imZeitraum.reduce((s, e) => s + Math.max(0, e.effekt), 0);
  const raus = imZeitraum.reduce((s, e) => s + Math.min(0, e.effekt), 0);

  // Monatsauswahl für „Zeitraum wählen“
  const alleMonate = [];
  for (let m = startMonat; m <= bisMonat; m = monatPlus(m, 1)) alleMonate.push(m);
  const monatOpt = (wert) => alleMonate.map((m) => `<option value="${m}" ${m === wert ? 'selected' : ''}>${monatsname(m, true)}</option>`).join('');

  // Liste: neu → alt, nach Monat gruppiert
  const gruppen = new Map();
  for (const e of [...imZeitraum].reverse()) {
    const m = e.b.datum.slice(0, 7);
    if (!gruppen.has(m)) gruppen.set(m, []);
    gruppen.get(m).push(e);
  }
  const zeileHtml = (e) => {
    const b = e.b;
    const titel = e.art === 'sparplan' ? 'Sparrate (Anteil)' : (b.posten || b.beschreibung);
    const sub = e.art === 'sparplan' ? `Anteil an ${eur(e.sparplan)} ${h(b.posten || 'Sparraten')}`
      : e.art === 'ohne' ? 'betrifft das Tagesgeld nicht, zählt nicht für den Topf' : '';
    const extra = e.art === 'sparplan'
      ? [['Für diesen Topf', `${eur(e.effekt)} = ${(vl.anteil * 100).toLocaleString('de-DE', { maximumFractionDigits: 1 })} % von ${eur(e.sparplan)}. So groß ist der Anteil dieses Topfs an allen Sparraten (${eur(t.sparrate)} von ${eur(t.sparrate / vl.anteil)} im Monat).`]]
      : [['Für diesen Topf', e.art === 'ohne' ? 'zählt nicht (kein Sparkonto beteiligt)' : vorzeichen(e.effekt)]];
    return `<button class="bu ${e.art === 'ohne' ? 'ohne-wirkung' : ''}" aria-expanded="false"><span class="t">${h(titel)}${istAusgleich(b) ? ' <span class="chip ausgleich">Ausgleich</span>' : ''}</span>
      <span class="betrag ${e.effekt > 0 ? 'plus' : ''}">${e.art === 'ohne' ? eur(b.betrag) : vorzeichen(e.effekt)}</span>
      <span class="r">${datumKurz(b.datum)} <span class="chip ${kontoFarbe(b.von)}">${h(b.von)}</span>${ICON.pfeil}<span class="chip ${kontoFarbe(b.nach)}">${h(b.nach)}</span></span>
      <span class="fuss"><span>${sub}</span><span class="danach">danach ${eur(e.danach)}</span></span>
      <span class="more"><dl>${details(b, extra)}</dl></span></button>`;
  };
  const liste = imZeitraum.length ? [...gruppen].map(([m, es]) => {
    const s = es.reduce((x, e) => x + e.effekt, 0);
    return `<div class="gruppe monatskopf"><span>${monatsname(m, true)}</span><span class="${s > 0.004 ? 'plus' : ''}">${Math.abs(s) < 0.005 ? '±0,00 €' : vorzeichen(Math.round(s * 100) / 100)}</span></div>
      <div class="buchungen">${es.map(zeileHtml).join('')}</div>`;
  }).join('') : `<div class="leer">In diesem Zeitraum gibt es keine Buchungen für diesen Topf.</div>`;

  const startZeile = `<div class="topf-start"><div><div style="font-weight:600">${abgeschnitten ? 'Anfangsbestand' : 'Stand davor'}</div>
      <div class="hint">${abgeschnitten ? (vl.beginn ? `am ${datumLang(vl.beginn)}, ab da wird gebucht` : 'ab da wird gebucht') : `am ${datumLang(vorTag(vonDatum))}`}</div></div>
    <div class="betrag">${eur(standAnfang)}</div></div>`;

  app.innerHTML = `
  <header><div class="wrap"><div class="monat">
    <button id="topf-zurueck" aria-label="Zurück zu den Töpfen">${ICON.links}</button>
    <div class="mitte"><h1>${h(t.name)}</h1><small>${h(t.bereich || 'Topf')} · ${h(t.konto || 'Tagesgeld')}</small></div>
    <span style="width:40px"></span>
  </div></div></header>
  <main class="wrap topf-seite">
    ${z.status ? `<div class="hinweis ${z.status.art}">${h(z.status.text)}${z.status.knopf ? ` <button class="link" id="${z.status.knopf[0]}">${h(z.status.knopf[1])}</button>` : ''}</div>` : ''}
    <div class="topf-kopf">
      <div class="hint">Stand ${bis === heute() ? 'heute' : 'am ' + datumLang(bis)}</div>
      <div class="gross ${vl.stand < 0 ? 'neg' : ''}">${eur(vl.stand)}</div>
      <div class="hint">${t.sparrate ? `Sparrate ${eur(t.sparrate)} im Monat` : 'keine feste Sparrate'}${t.bedarf ? ` · Bedarf ${eur(t.bedarf)} im Monat` : ''}</div>
      ${t.bemerkung ? `<p class="hint" style="margin-top:6px">${h(t.bemerkung)}</p>` : ''}
    </div>
    <section>
      <div class="filter umbruch" role="group" aria-label="Zeitraum">${ZEITRAEUME.map(([id, name]) => `<button aria-pressed="${z.zeitraum === id}" data-zeitraum="${id}">${name}</button>`).join('')}</div>
      ${z.zeitraum === 'eigen' ? `<div class="zwei zeitraum-wahl">
        <label class="feld"><span>Von</span><select id="topf-von">${monatOpt(von)}</select></label>
        <label class="feld"><span>Bis</span><select id="topf-bis">${monatOpt(bisM)}</select></label></div>` : ''}
      <p class="hint" style="margin:4px 0 12px">${abgeschnitten ? `Alles seit Beginn des Topfs${vl.beginn ? ` (${datumLang(vl.beginn)})` : ''}` : `${monatsname(von, true)}`}${bisM !== von || abgeschnitten ? ` bis ${bisM === bisMonat ? (bis === heute() ? 'heute' : datumLang(bis)) : monatsname(bisM, true)}` : ''}. ${imZeitraum.length} ${imZeitraum.length === 1 ? 'Buchung' : 'Buchungen'}.</p>
      <div class="topfsumme">
        <div class="z"><span>${abgeschnitten ? 'Anfangsbestand' : 'Stand am Anfang'}</span><b>${eur(standAnfang)}</b></div>
        <div class="z"><span>Hinein</span><b class="plus">${vorzeichen(Math.round(rein * 100) / 100)}</b></div>
        <div class="z"><span>Heraus</span><b>${eur(Math.round(raus * 100) / 100)}</b></div>
        <div class="z"><span>Stand am Ende</span><b>${eur(standEnde)}</b></div>
      </div>
    </section>
    <section>${liste}${startZeile}
      ${imZeitraum.length ? '<p class="hint" style="margin-top:10px">Antippen zeigt Details und Kommentar. „danach“ ist der Stand des Topfs nach der Buchung.</p>' : ''}</section>
  </main>`;
}
const vorTag = (iso) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() - 1); return d.toISOString().slice(0, 10); };

function oeffneTopf(name) {
  z.topfScroll = window.scrollY;
  z.topf = name;
  try { history.pushState({ topf: name }, ''); } catch {}
  zeichne(); window.scrollTo(0, 0);
}
function schliesseTopf() {
  z.topf = null; zeichne();
  window.scrollTo(0, z.topfScroll || 0);
}
window.addEventListener('popstate', () => { if (z.topf && !history.state?.topf) schliesseTopf(); });

function buchungen(v) {
  const liste = v.buchungen.filter((b) => z.filter === 'alle' || b.art === z.filter);
  const f = (id, name) => `<button aria-pressed="${z.filter === id}" data-filter="${id}">${name}</button>`;
  return `<section><h2>Buchungen</h2>
    <div class="filter" role="group" aria-label="Filter">${f('alle', 'Alle')}${f('Ausgabe', 'Ausgaben')}${f('Einnahme', 'Einnahmen')}${f('Umbuchung', 'Umbuchungen')}</div>
    ${liste.length ? `<div class="buchungen">${liste.map((b) => {
      const sign = b.art === 'Einnahme' ? '+' : b.art === 'Ausgabe' ? '−' : '';
      const andererMonat = b.giltFuer && b.giltFuer !== v.monat;
      const dl = details(b);
      const aktion = kannEintragen() ? `<div class="bu-aktion"><button class="knopf leise klein" data-bearbeiten="${h(b.nr)}">Bearbeiten</button><button class="knopf leise klein" data-loeschen-frage="${h(b.nr)}">Löschen</button>${istAusgleich(b)
        ? `<button class="knopf leise klein" data-ausgleich="" data-nr="${h(b.nr)}">Ausgleich aufheben, wieder anrechnen</button>`
        : `<button class="knopf leise klein" data-ausgleich-frage="${h(b.nr)}">Als Ausgleich markieren</button>`}</div>` : '';
      return `<div class="bu-wrap"><button class="bu" aria-expanded="false"><span class="t">${h(b.posten || b.beschreibung)}${istAusgleich(b) ? ' <span class="chip ausgleich">Ausgleich</span>' : ''}</span><span class="betrag ${b.art === 'Einnahme' ? 'plus' : ''}">${sign}${eur(b.betrag)}</span>
        <span class="r">${datumKurz(b.datum)} <span class="chip ${kontoFarbe(b.von)}">${h(b.von)}</span>${ICON.pfeil}<span class="chip ${kontoFarbe(b.nach)}">${h(b.nach)}</span>${andererMonat ? `<span class="okt">gilt für ${h(monatsname(b.giltFuer))}</span>` : ''}</span>
        <span class="more"><dl>${dl}</dl></span></button>${aktion}</div>`;
    }).join('')}</div><p class="hint" style="margin-top:10px">Antippen zeigt Details und Kommentar${kannEintragen() ? ', dort auch Bearbeiten und Löschen' : ''}.</p>`
      : `<div class="leer">${z.filter === 'alle' ? `Für ${h(v.name)} ist noch nichts gebucht.` : 'Keine Buchungen dieser Art in diesem Monat.'}</div>`}
  </section>`;
}

function zeigeMenu() {
  const d = z.datei;
  const zeit = (t) => new Date(t).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' });
  const ausOd = d?.quelle === 'onedrive';
  const konto = ms.gemerktesKonto();
  menu.innerHTML = `<h2 id="menu-titel">Datei</h2>
    <div class="zeile"><span class="hint">Quelle</span><span>${ausOd ? 'OneDrive' : 'von diesem Handy'}</span></div>
    ${d?.geaendert ? `<div class="zeile"><span class="hint">Zuletzt geändert</span><span>${zeit(d.geaendert)}${d.geaendertVon ? ' von ' + h(d.geaendertVon) : ''}</span></div>` : ''}
    <div class="zeile"><span class="hint">In die App geladen</span><span>${d?.geladen ? zeit(d.geladen) : '–'}</span></div>
    <div class="zeile"><span class="hint">Buchungen</span><span>${z.modell.buchungen.length}</span></div>
    ${ms.aktiv() && ms.angemeldet() ? `<div class="zeile"><span class="hint">Angemeldet</span><span>${h(konto || 'ja')}</span></div>` : ''}
    <div class="knoepfe">
      ${ms.aktiv() && ms.angemeldet() ? `<button class="knopf" id="aktualisieren">Jetzt aus OneDrive laden</button>`
        : ms.aktiv() ? `<button class="knopf" id="anmelden">Mit Microsoft anmelden</button>` : ''}
      ${d?.webUrl ? `<a class="knopf leise" href="${h(d.webUrl)}" target="_blank" rel="noopener">In Excel öffnen</a>` : ''}
      <label class="knopf leise" for="datei">Datei vom Handy öffnen</label>
      <a class="knopf leise" href="#pc">PC-Ansicht (Kontoauszüge importieren)</a>
      ${ms.aktiv() && ms.angemeldet() ? `<button class="knopf leise" id="abmelden">Abmelden</button>` : ''}
      <button class="knopf leise" id="entfernen">Daten von diesem Handy entfernen</button>
      <button class="knopf leise" id="schliessen">Schließen</button>
    </div>
    <p class="hint" style="margin-top:14px;text-align:center">Version ${VERSION}</p>`;
  menu.showModal();
}


// ---------- Eintragen
const SR = window.SpeechRecognition || window.webkitSpeechRecognition || null;
const kannEintragen = () => ms.aktiv() && ms.angemeldet() && z.datei?.quelle === 'onedrive' && !!od.aktuellerOrt();
const STANDARD = {
  Ausgabe: (k) => ({ von: k.gem, nach: k.extern }),
  Einnahme: (k) => ({ von: k.extern, nach: k.gem }),
  Umbuchung: (k) => ({ von: k.gem, nach: k.visa }),
};
function betragLesen(t) {
  const s = String(t || '').trim().replace(/\s|€/g, '');
  if (!s) return NaN;
  const norm = s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s;
  return Math.round(Number(norm) * 100) / 100;
}
function zeigeEintragen(fehler, vorher) {
  const m = z.modell;
  const konten = m.konten.filter((k) => k.status !== 'geschlossen').map((k) => k.name);
  const kn = {
    gem: m.konten.find((k) => k.rolle === 'gemeinsam' && /giro/i.test(k.art || ''))?.name || konten[0],
    visa: m.konten.find((k) => k.rolle === 'gemeinsam' && /kredit|visa/i.test((k.art || '') + k.name))?.name || konten[0],
    extern: m.konten.find((k) => k.rolle === 'extern')?.name || 'Extern',
  };
  const h0 = heute();
  const vormonat = monatPlus(h0.slice(0, 7), -1);
  const monateAuswahl = [...new Set([vormonat, ...(m.listen?.monate || monate(m))])].sort();
  const toepfe = ['Sparplan (verteilt)', ...m.toepfe.map((t) => t.name)];
  const postenListe = [...new Set([...m.kosten.map((k) => k.posten), ...m.einnahmen.map((e) => e.posten), ...m.buchungen.map((b) => b.posten)].filter(Boolean))].sort((a, b) => a.localeCompare(b, 'de'));
  const v = vorher || { art: 'Ausgabe', datum: h0, giltFuer: z.monat && z.monat >= vormonat ? (h0.slice(0, 7) === z.monat ? z.monat : z.monat) : h0.slice(0, 7), ...STANDARD.Ausgabe(kn), bereich: '', topf: '', posten: '', betrag: '', beschreibung: '', kommentar: '', gezahltVon: '' };
  const bearbeiten = Boolean(v.nr);
  const dazu = (liste, wert) => (wert && !liste.includes(wert) ? [...liste, wert] : liste);
  const kontenAuswahl = dazu(dazu(konten, v.von), v.nach);
  const bereiche = dazu(m.listen?.bereiche || [], v.bereich);
  const toepfeAuswahl = dazu(toepfe, v.topf);
  if (v.giltFuer && !monateAuswahl.includes(v.giltFuer)) { monateAuswahl.push(v.giltFuer); monateAuswahl.sort(); }
  const doppelt = v.doppelt?.length ? `<div class="hinweis warn"><b>Das steht schon drin:</b> ${v.doppelt.map((t) => `${h(t.nr)} ${h(t.posten)}`).join(', ')} – gleicher Tag, gleicher Betrag, gleiche Konten.
      Wenn es wirklich ein zweites Mal war, auf „Trotzdem eintragen“ tippen. Sonst „Abbrechen“.</div><input type="hidden" name="trotzdem" value="1">` : '';
  const ausgleichAn = v.ausgleich === 'on' || istAusgleich(v);
  const grund = v.grund || (istAusgleich(v) ? (/topf/i.test(v.gezahltVon) ? 'topf' : 'korrektur') : '');
  const opt = (liste, wert, leer) => (leer ? `<option value="">${leer}</option>` : '') + liste.map((x) => `<option value="${h(x)}" ${x === wert ? 'selected' : ''}>${h(x)}</option>`).join('');
  menu.innerHTML = `<form id="eintragen-form" novalidate>
    <h2 id="menu-titel">${bearbeiten ? `${h(v.nr)} ändern` : 'Buchung eintragen'}</h2>
    ${bearbeiten ? `<input type="hidden" name="nr" value="${h(v.nr)}">` : ''}
    <div class="sprechen" ${bearbeiten ? 'hidden' : ''}>
      ${SR ? `<button type="button" class="mikro" id="mikro" aria-label="Sprechen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg></button>` : ''}
      <input name="satz" id="satz" autocomplete="off" enterkeyhint="done" placeholder="${SR ? 'Sprechen oder schreiben, z. B. „12,50 Bäckerei mit der Visa“' : 'Schreiben oder per Tastatur-Mikrofon diktieren'}">
      <button type="button" class="knopf leise klein" id="satz-los">Übernehmen</button>
    </div>
    <p class="erkannt" id="erkannt" aria-live="polite" ${bearbeiten ? 'hidden' : ''}></p>
    <div class="seg" role="radiogroup" aria-label="Art">${['Ausgabe', 'Einnahme', 'Umbuchung'].map((a) =>
      `<label><input type="radio" name="art" value="${a}" ${v.art === a ? 'checked' : ''}><span>${a}</span></label>`).join('')}</div>
    <label class="feld"><span>Betrag in €</span><input name="betrag" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${h(v.betrag)}" required></label>
    <label class="feld"><span>Wofür</span><input name="posten" list="posten-liste" autocomplete="off" placeholder="z. B. Lidl-Einkauf" value="${h(v.posten)}" required>
      <datalist id="posten-liste">${postenListe.map((p) => `<option value="${h(p)}">`).join('')}</datalist></label>
    <div class="zwei">
      <label class="feld"><span>Bezahlt am</span><input type="date" name="datum" value="${h(v.datum)}" required></label>
      <label class="feld"><span>Gilt für</span><select name="giltFuer">${monateAuswahl.map((mo) => `<option value="${mo}" ${mo === v.giltFuer ? 'selected' : ''}>${monatsname(mo, true)}</option>`).join('')}</select></label>
    </div>
    <div class="zwei">
      <label class="feld"><span>Von</span><select name="von">${opt(kontenAuswahl, v.von)}</select></label>
      <label class="feld"><span>Nach</span><select name="nach">${opt(kontenAuswahl, v.nach)}</select></label>
    </div>
    <label class="feld"><span>Bereich</span><select name="bereich">${opt(bereiche, v.bereich, 'bitte wählen')}</select></label>
    <label class="feld"><span>Topf <small>(nur wenn Geld aufs Tagesgeld geht oder von dort kommt)</small></span><select name="topf">${opt(toepfeAuswahl, v.topf, 'kein Topf')}</select></label>
    <label class="feld"><span>Verwendungszweck <small>(optional)</small></span><input name="beschreibung" autocomplete="off" value="${h(v.beschreibung)}"></label>
    <label class="feld"><span>Kommentar <small>(optional)</small></span><input name="kommentar" autocomplete="off" value="${h(v.kommentar)}"></label>
    <label class="feld" id="gezahlt-feld" ${ausgleichAn ? 'hidden' : ''}><span>Gezahlt von <small>(nur wenn es nicht vom eigenen Konto kommt, z. B. Elterngeld)</small></span><select name="gezahltVon">${opt(['Kai', 'Natalie'], istAusgleich(v) ? '' : v.gezahltVon, '–')}</select></label>
    <div class="ausgleich-feld">
      <label class="haken"><input type="checkbox" name="ausgleich" ${ausgleichAn ? 'checked' : ''}><span><b>Ausgleich, nicht anrechnen</b><small>Die Buchung bleibt sichtbar, zählt aber nicht als Einzahlung.</small></span></label>
      <div class="grund" id="grund" ${ausgleichAn ? '' : 'hidden'} role="radiogroup" aria-label="Grund">
        <label><input type="radio" name="grund" value="topf" ${grund === 'topf' ? 'checked' : ''}><span><b>Aus Topf weitergeleitet</b><small>Geld kam aus einem Topf und lief nur übers eigene Konto</small></span></label>
        <label><input type="radio" name="grund" value="korrektur" ${grund === 'korrektur' ? 'checked' : ''}><span><b>Falsch bezahlt / Korrektur</b><small>Mit der falschen Karte bezahlt, oder die Rücküberweisung dazu</small></span></label>
      </div>
    </div>
    <p class="zusammen" id="zusammen"></p>
    ${doppelt}
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" type="submit" id="eintragen-los">${bearbeiten ? 'Speichern' : doppelt ? 'Trotzdem eintragen' : 'Eintragen'}</button><button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div>
  </form>`;
  const form = document.getElementById('eintragen-form');
  const zusammen = () => {
    const d = Object.fromEntries(new FormData(form));
    const b = betragLesen(d.betrag);
    document.getElementById('zusammen').textContent = isFinite(b) && b > 0
      ? `${d.art} über ${eur(b)} von ${d.von} an ${d.nach}, gilt für ${monatsname(d.giltFuer)}.${d.ausgleich ? ' Ausgleich, wird nicht angerechnet.' : ''}` : '';
  };
  form.addEventListener('input', zusammen);
  form.addEventListener('change', (ev) => {
    if (ev.target.name === 'art') { const s = STANDARD[ev.target.value](kn); form.von.value = s.von; form.nach.value = s.nach; zusammen(); }
    if (ev.target.name === 'ausgleich') {
      document.getElementById('grund').hidden = !ev.target.checked;
      document.getElementById('gezahlt-feld').hidden = ev.target.checked;
    }
    if (ev.target.name === 'datum' && ev.target.value) { const mo = ev.target.value.slice(0, 7); if ([...form.giltFuer.options].some((o) => o.value === mo)) form.giltFuer.value = mo; zusammen(); }
  });
  zusammen();

  // Satz verstehen und Formular vorausfüllen
  const satzFeld = document.getElementById('satz');
  const hinweis = document.getElementById('erkannt');
  const uebernehmen = () => {
    const satz = satzFeld.value.trim();
    if (!satz) return;
    const { felder: f, erkannt, fehlt } = verstehen(satz, z.modell, heute());
    const setze = (name, wert) => {
      const el = form.elements[name];
      if (wert == null || wert === '' || !el) return;
      if (el.tagName === 'SELECT' && ![...el.options].some((o) => o.value === wert)) {
        if (name !== 'giltFuer') return;
        el.add(new Option(monatsname(wert, true), wert));
      }
      el.value = wert;
      const feld = el.closest('.feld'); if (feld) { feld.classList.remove('neu'); void feld.offsetWidth; feld.classList.add('neu'); }
    };
    if (f.art) form.querySelector(`input[name=art][value="${f.art}"]`).checked = true;
    ['betrag', 'posten', 'datum', 'giltFuer', 'von', 'nach', 'bereich', 'topf', 'kommentar'].forEach((n) => setze(n, f[n]));
    if (f.datum && !f.giltFuer) { const mo = f.datum.slice(0, 7); if ([...form.giltFuer.options].some((o) => o.value === mo)) form.giltFuer.value = mo; }
    hinweis.innerHTML = `Vorausgefüllt${erkannt.length ? ': ' + h(erkannt.join(', ')) : ''}. Bitte kurz prüfen.${fehlt.length ? ` <b>Noch offen: ${h(fehlt.join(', '))}.</b>` : ''}`;
    zusammen();
  };
  document.getElementById('satz-los').addEventListener('click', uebernehmen);
  satzFeld.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); uebernehmen(); satzFeld.blur(); } });

  const mikro = document.getElementById('mikro');
  if (mikro) {
    let erkennung = null;
    mikro.addEventListener('click', () => {
      if (erkennung) { erkennung.stop(); return; }
      erkennung = new SR();
      erkennung.lang = 'de-DE'; erkennung.interimResults = true; erkennung.maxAlternatives = 1; erkennung.continuous = false;
      mikro.classList.add('hoert'); hinweis.textContent = 'Ich höre zu …';
      erkennung.onresult = (e) => {
        satzFeld.value = [...e.results].map((r) => r[0].transcript).join(' ');
        if (e.results[e.results.length - 1].isFinal) uebernehmen();
      };
      erkennung.onerror = (e) => {
        hinweis.textContent = e.error === 'not-allowed' || e.error === 'service-not-allowed'
          ? 'Das Mikrofon ist für diese Seite gesperrt. In den Browser-Einstellungen erlauben oder den Satz tippen.'
          : e.error === 'no-speech' ? 'Nichts gehört. Noch einmal auf das Mikrofon tippen.' : 'Spracherkennung gerade nicht möglich. Bitte tippen.';
      };
      erkennung.onend = () => { mikro.classList.remove('hoert'); erkennung = null; };
      try { erkennung.start(); } catch { mikro.classList.remove('hoert'); erkennung = null; }
    });
  }
  if (!menu.open) menu.showModal();
}

/** Nach erfolgreichem Schreiben: Änderung sofort anzeigen, dann warten, bis OneDrive die neue Version liefert, und neu laden. */
async function nachDemSchreiben(aenderung, text) {
  const vorher = z.datei?.eTag;
  ausstehend.push({ ...aenderung, seit: Date.now() }); ausstehendMerken();
  z.modell.buchungen = ausstehendAnwenden(z.modell.buchungen);
  z.status = { art: 'info', text }; zeichne();
  await od.aufNeueVersionWarten(od.aktuellerOrt(), vorher).catch(() => null);
  await ausOneDrive({ immer: true });
  z.status = { art: 'info', text }; zeichne();
}
const schreibFehler = (e, nichts = 'Es wurde nichts gespeichert') => (e instanceof od.AnmeldungNoetig ? `Die Anmeldung ist abgelaufen. Bitte im Menü neu anmelden. ${nichts}.`
  : e instanceof od.KeinNetz ? `Gerade keine Verbindung zu OneDrive. ${nichts}, bitte gleich noch einmal versuchen.`
    : /423|locked|gesperrt/i.test(e.message) ? 'Die Datei ist gerade in Excel geöffnet und gesperrt. Bitte dort schließen und noch einmal versuchen.'
      : 'Das hat nicht geklappt: ' + e.message + ' Bitte in Excel nachsehen, bevor ihr es noch einmal versucht.');

document.addEventListener('submit', async (ev) => {
  if (ev.target.id !== 'eintragen-form') return;
  ev.preventDefault();
  if (z.schreibt) return;
  const form = ev.target;
  const d = Object.fromEntries(new FormData(form));
  const betrag = betragLesen(d.betrag);
  const fehler = !(betrag > 0) ? 'Bitte einen Betrag größer als 0 eingeben.'
    : !d.posten.trim() ? 'Bitte eintragen, wofür das Geld war.'
      : !d.datum ? 'Bitte das Datum angeben.'
        : d.von === d.nach ? 'Von und Nach sind dasselbe Konto.'
          : !d.bereich ? 'Bitte einen Bereich wählen.'
            : d.ausgleich && !d.grund ? 'Bitte beim Ausgleich den Grund wählen.' : null;
  if (fehler) { zeigeEintragen(fehler, d); return; }
  const alt = d.nr ? z.modell.buchungen.find((x) => x.nr === d.nr) : null;
  if (d.nr && !alt) { zeigeEintragen(`${d.nr} gibt es nicht mehr. Bitte die Liste neu laden.`, d); return; }
  const gezahltVon = d.ausgleich ? AUSGLEICH[d.grund] : (d.gezahltVon || '');
  const posten = d.posten.trim();
  const daten = { datum: d.datum, giltFuer: d.giltFuer, art: d.art, von: d.von, nach: d.nach, betrag, bereich: d.bereich, posten, topf: d.topf || '',
    beschreibung: d.beschreibung || '', kommentar: d.kommentar || '', gezahltVon };
  // Schon in der Anzeige vorhanden? (Die Datei selbst wird beim Speichern noch einmal geprüft.)
  if (!alt && !d.trotzdem) {
    const treffer = z.modell.buchungen.filter((b) => b.datum === d.datum && b.art === d.art && b.von === d.von && b.nach === d.nach && Math.abs(b.betrag - betrag) < 0.005);
    if (treffer.length) { zeigeEintragen(null, { ...d, doppelt: treffer.map((b) => ({ nr: b.nr, posten: b.posten })) }); return; }
  }
  const knopf = document.getElementById('eintragen-los');
  knopf.disabled = true; knopf.textContent = alt ? 'Wird gespeichert …' : 'Wird eingetragen …';
  z.schreibt = true;
  let nr;
  try {
    if (alt) { await od.buchungAendern(od.aktuellerOrt(), alt.nr, daten); nr = alt.nr; }
    else ({ nr } = await od.buchungEintragen(od.aktuellerOrt(), daten, { trotzdem: Boolean(d.trotzdem) }));
  } catch (e) {
    z.schreibt = false;
    if (e instanceof od.Doppelt) { zeigeEintragen(null, { ...d, doppelt: e.treffer }); return; }
    zeigeEintragen(schreibFehler(e), d);
    return;
  }
  z.schreibt = false;
  menu.close();
  const b = { ...(alt || { quelle: 'manuell', importId: '', excelZuordnung: '' }), ...daten, nr };
  await nachDemSchreiben({ typ: alt ? 'aendern' : 'neu', nr, b }, alt ? `${nr} geändert: ${posten}, ${eur(betrag)}.` : `${nr} eingetragen: ${posten}, ${eur(betrag)}.`);
});

// ---------- Buchung bearbeiten / löschen
function zeigeBearbeiten(nr) {
  const b = z.modell.buchungen.find((x) => x.nr === nr);
  if (!b) return;
  zeigeEintragen(null, { ...b, betrag: b.betrag.toFixed(2).replace('.', ','), ausgleich: istAusgleich(b) ? 'on' : '' });
}
function zeigeLoeschen(nr, fehler) {
  const b = z.modell.buchungen.find((x) => x.nr === nr);
  if (!b) return;
  menu.innerHTML = `<h2 id="menu-titel">${h(nr)} löschen?</h2>
    <p class="ausgleich-was"><b>${h(b.posten || b.beschreibung)}</b>, ${eur(b.betrag)}<br><span class="hint">${datumKurz(b.datum)} ${h(b.von)} → ${h(b.nach)}</span></p>
    <p class="hint">Die Buchung wird aus der Datei entfernt. Rückgängig machen geht nur, indem man sie neu einträgt.${b.importId ? ' Sie kam aus dem Kontoauszug: Beim nächsten Abgleich wird die Bankzeile wieder als „neu“ angeboten.' : ''}</p>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" data-loeschen="${h(nr)}">Ja, löschen</button><button class="knopf leise" id="schliessen">Abbrechen</button></div>`;
  if (!menu.open) menu.showModal();
}
async function loeschen(nr, knopf) {
  const b = z.modell.buchungen.find((x) => x.nr === nr);
  if (!b || z.schreibt) return;
  knopf.disabled = true; knopf.textContent = 'Wird gelöscht …';
  z.schreibt = true;
  try { await od.buchungLoeschen(od.aktuellerOrt(), nr); }
  catch (e) { z.schreibt = false; zeigeLoeschen(nr, schreibFehler(e, 'Es wurde nichts gelöscht')); return; }
  z.schreibt = false;
  menu.close();
  await nachDemSchreiben({ typ: 'loeschen', nr, b }, `${nr} gelöscht: ${b.posten || b.beschreibung}, ${eur(b.betrag)}.`);
}

// ---------- Ausgleich an bestehender Buchung setzen oder aufheben
function zeigeAusgleichFrage(nr, fehler) {
  const b = z.modell.buchungen.find((x) => x.nr === nr);
  if (!b) return;
  menu.innerHTML = `<h2 id="menu-titel">Als Ausgleich markieren</h2>
    <p class="ausgleich-was"><b>${h(b.posten || b.beschreibung)}</b>, ${eur(b.betrag)}<br><span class="hint">${datumKurz(b.datum)} ${h(b.von)} → ${h(b.nach)}</span></p>
    <p class="hint">Die Buchung bleibt, wie sie ist, und zählt weiter bei den Kontoständen. Sie wird nur nicht mehr als Einzahlung angerechnet.</p>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe">
      <button class="knopf grund-knopf" data-ausgleich="topf" data-nr="${h(nr)}"><b>Aus Topf weitergeleitet</b><small>Geld kam aus einem Topf und lief nur übers eigene Konto</small></button>
      <button class="knopf grund-knopf" data-ausgleich="korrektur" data-nr="${h(nr)}"><b>Falsch bezahlt / Korrektur</b><small>Mit der falschen Karte bezahlt, oder die Rücküberweisung dazu</small></button>
      <button class="knopf leise" id="schliessen">Abbrechen</button>
    </div>`;
  if (!menu.open) menu.showModal();
}
async function ausgleichSetzen(nr, grund, knopf) {
  const b = z.modell.buchungen.find((x) => x.nr === nr);
  if (!b) return;
  const wert = grund ? AUSGLEICH[grund] : '';
  if (z.schreibt) return;
  if (knopf) { knopf.disabled = true; knopf.textContent = 'Wird gespeichert …'; }
  const text = grund ? `${nr} ist jetzt ein Ausgleich und wird nicht angerechnet.` : `${nr} wird wieder angerechnet.`;
  z.schreibt = true;
  try {
    await od.gezahltVonSetzen(od.aktuellerOrt(), nr, wert);
    z.schreibt = false;
    if (menu.open) menu.close();
    await nachDemSchreiben({ typ: 'aendern', nr, b: { ...b, gezahltVon: wert } }, text);
  } catch (e) {
    z.schreibt = false;
    const msg = e instanceof od.AnmeldungNoetig ? 'Die Anmeldung ist abgelaufen. Bitte im Menü neu anmelden.'
      : e instanceof od.KeinNetz ? 'Gerade keine Verbindung zu OneDrive. Es wurde nichts geändert.'
        : 'Das hat nicht geklappt: ' + e.message;
    if (grund) zeigeAusgleichFrage(nr, msg);
    else { z.status = { art: 'warn', text: msg }; zeichne(); }
  }
}

// ---------- Ereignisse
document.addEventListener('click', async (ev) => {
  if (document.body.classList.contains('pc-modus')) return;
  const t = ev.target.closest('button, [data-monat]');
  if (!t) { if (ev.target === menu) menu.close(); return; }
  if (t.dataset.bereich) { z.bereich = t.dataset.bereich; merken('bereich', z.bereich); zeichne(); window.scrollTo(0, 0); return; }
  if (z.bereich === 'meins' && !z.topf && meins.verfuegbar() && await meins.klick(t)) return;
  if (t.dataset.monat) { z.monat = t.dataset.monat; zeichne(); window.scrollTo(0, 0); return; }
  if (t.dataset.tab) { z.tab = t.dataset.tab; merken('tab', z.tab); zeichne(); window.scrollTo(0, 0); return; }
  if (t.dataset.filter) { z.filter = t.dataset.filter; zeichne(); return; }
  if (t.dataset.topf) { oeffneTopf(t.dataset.topf); return; }
  if (t.id === 'topf-zurueck') { if (history.state?.topf) history.back(); else schliesseTopf(); return; }
  if (t.dataset.zeitraum) {
    if (t.dataset.zeitraum === 'eigen' && z.zeitraum !== 'eigen') { z.eigenVon = null; z.eigenBis = null; }
    z.zeitraum = t.dataset.zeitraum; zeichne(); return;
  }
  if (t.dataset.bearbeiten) { zeigeBearbeiten(t.dataset.bearbeiten); return; }
  if (t.dataset.loeschenFrage) { zeigeLoeschen(t.dataset.loeschenFrage); return; }
  if (t.dataset.loeschen) { loeschen(t.dataset.loeschen, t); return; }
  if (t.dataset.ausgleichFrage) { zeigeAusgleichFrage(t.dataset.ausgleichFrage); return; }
  if (t.dataset.ausgleich != null && t.dataset.nr) { ausgleichSetzen(t.dataset.nr, t.dataset.ausgleich, t); return; }
  if (t.classList.contains('bu')) { t.setAttribute('aria-expanded', t.getAttribute('aria-expanded') === 'true' ? 'false' : 'true'); return; }
  if (t.id === 'menu-knopf') { zeigeMenu(); return; }
  if (t.id === 'eintragen-knopf') { zeigeEintragen(); return; }
  if (t.id === 'schliessen') { menu.close(); return; }
  if (t.id === 'anmelden') { ms.anmelden({ konto: t.closest('.start') ? null : ms.gemerktesKonto() }); return; }
  if (t.id === 'aktualisieren') { if (menu.open) menu.close(); ausOneDrive({ immer: true }); return; }
  if (t.id === 'abmelden') { meins.vergessen(); ms.abmelden(); menu.close(); z.status = { art: 'info', text: 'Abgemeldet. Ihr seht den zuletzt geladenen Stand.' }; zeichne(); return; }
  if (t.id === 'entfernen') { meins.vergessen(); await loescheDatei(); ms.abmelden(); menu.close(); z.modell = null; z.datei = null; zeigeStart(); }
});

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
meins.init({ z, app, menu, h, heute, zeichne, ausOneDrive, zeigeEintragen, zeigeMenu, ms, od, kopf, version: VERSION });
start();
