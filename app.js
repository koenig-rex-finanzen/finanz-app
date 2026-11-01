// Handy-App „Unser Geld“ – zeigt die gemeinsame Buchhaltung an (nur lesen).
import { leseBuchhaltung } from './kern/lesen.js';
import { monatsansicht, monate, monatsname, monatPlus, eur } from './kern/rechnen.js';
import * as ms from './anmeldung.js';
import * as od from './onedrive.js';
import { verstehen } from './kern/verstehen.js';

const VERSION = '0.5.1';
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

// ---------- Zustand
const z = { modell: null, datei: null, monat: null, tab: holen('tab') || 'uebersicht', filter: 'alle', status: null, laedt: false, geprueft: 0 };

// ---------- Start
async function start() {
  let fehler = null;
  if (ms.aktiv()) {
    try {
      if (await ms.rueckkehrVerarbeiten()) {
        const me = await od.ich().catch(() => null);
        if (me) ms.kontoMerken(me.mail || me.userPrincipalName || me.displayName);
      }
    } catch (e) { fehler = e.message; }
  }
  if (location.hash === '#pc') {
    const pc = await import('./pc.js');
    await pc.starte(app);
    return;
  }
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
  z.modell = modell; z.datei = eintrag;
  const liste = monate(modell);
  const jetzt = heute().slice(0, 7);
  if (!z.monat || !liste.includes(z.monat)) z.monat = liste.includes(jetzt) ? jetzt : (liste.find((m) => m > jetzt) || liste[liste.length - 1]);
  if (speichern) await speichereDatei(eintrag);
  zeichne();
}

/** Lädt die Datei aus OneDrive, wenn sie sich geändert hat. */
async function ausOneDrive({ immer = false } = {}) {
  if (z.laedt) return;
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
  const v = monatsansicht(z.modell, z.monat, heute());
  const liste = monate(z.modell);
  const i = liste.indexOf(z.monat);
  const [jahr] = z.monat.split('-');
  app.innerHTML = `
  <header><div class="wrap"><div class="monat">
    <button data-monat="${i > 0 ? liste[i - 1] : ''}" aria-label="Vorheriger Monat" ${i > 0 ? '' : 'disabled'}>${ICON.links}</button>
    <div class="mitte"><h1>${h(v.name)} ${jahr}</h1><small>${v.laufend ? 'dieser Monat' : v.vergangen ? 'vorbei' : 'Vorschau'}</small></div>
    <div style="display:flex">
      <button data-monat="${i < liste.length - 1 ? liste[i + 1] : ''}" aria-label="Nächster Monat" ${i < liste.length - 1 ? '' : 'disabled'}>${ICON.rechts}</button>
      <button class="menu" id="menu-knopf" aria-label="Datei und Einstellungen">${ICON.menu}</button>
    </div>
  </div></div></header>
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
      `<div class="topf ${t.stand < 0 ? 'minus' : ''}"><span class="n">${h(t.name)}</span><span class="b">${eur(t.stand)}</span>
        <span class="s" style="grid-column:1/-1">${t.sparrate ? `${eur(t.sparrate)} im Monat` : 'keine feste Sparrate'}${t.bemerkung ? '. ' + h(t.bemerkung) : ''}</span></div>`).join('')}</div>`).join('')}
  </section>`;
}

function buchungen(v) {
  const liste = v.buchungen.filter((b) => z.filter === 'alle' || b.art === z.filter);
  const f = (id, name) => `<button aria-pressed="${z.filter === id}" data-filter="${id}">${name}</button>`;
  return `<section><h2>Buchungen</h2>
    <div class="filter" role="group" aria-label="Filter">${f('alle', 'Alle')}${f('Ausgabe', 'Ausgaben')}${f('Einnahme', 'Einnahmen')}${f('Umbuchung', 'Umbuchungen')}</div>
    ${liste.length ? `<div class="buchungen">${liste.map((b) => {
      const sign = b.art === 'Einnahme' ? '+' : b.art === 'Ausgabe' ? '−' : '';
      const andererMonat = b.giltFuer && b.giltFuer !== v.monat;
      const dl = [['Nummer', b.nr], ['Bezahlt am', b.datum ? b.datum.split('-').reverse().join('.') : ''], ['Gilt für', b.giltFuer ? monatsname(b.giltFuer, true) : ''],
        ['Art', b.art], ['Bereich', b.bereich], ['Topf', b.topf], ['Verwendungszweck', b.beschreibung], ['Kommentar', b.kommentar],
        ['Gezahlt von', b.gezahltVon], ['Eingetragen', b.quelle === 'CSV' ? 'aus dem Bank-Export' : b.quelle]]
        .filter(([, w]) => w).map(([k, w]) => `<dt>${k}</dt><dd>${h(w)}</dd>`).join('');
      return `<button class="bu" aria-expanded="false"><span class="t">${h(b.posten || b.beschreibung)}</span><span class="betrag ${b.art === 'Einnahme' ? 'plus' : ''}">${sign}${eur(b.betrag)}</span>
        <span class="r">${datumKurz(b.datum)} <span class="chip ${kontoFarbe(b.von)}">${h(b.von)}</span>${ICON.pfeil}<span class="chip ${kontoFarbe(b.nach)}">${h(b.nach)}</span>${andererMonat ? `<span class="okt">gilt für ${h(monatsname(b.giltFuer))}</span>` : ''}</span>
        <span class="more"><dl>${dl}</dl></span></button>`;
    }).join('')}</div><p class="hint" style="margin-top:10px">Antippen zeigt Details und Kommentar.</p>`
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
  const opt = (liste, wert, leer) => (leer ? `<option value="">${leer}</option>` : '') + liste.map((x) => `<option value="${h(x)}" ${x === wert ? 'selected' : ''}>${h(x)}</option>`).join('');
  menu.innerHTML = `<form id="eintragen-form" novalidate>
    <h2 id="menu-titel">Buchung eintragen</h2>
    <div class="sprechen">
      ${SR ? `<button type="button" class="mikro" id="mikro" aria-label="Sprechen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg></button>` : ''}
      <input name="satz" id="satz" autocomplete="off" enterkeyhint="done" placeholder="${SR ? 'Sprechen oder schreiben, z. B. „12,50 Bäckerei mit der Visa“' : 'Schreiben oder per Tastatur-Mikrofon diktieren'}">
      <button type="button" class="knopf leise klein" id="satz-los">Übernehmen</button>
    </div>
    <p class="erkannt" id="erkannt" aria-live="polite"></p>
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
      <label class="feld"><span>Von</span><select name="von">${opt(konten, v.von)}</select></label>
      <label class="feld"><span>Nach</span><select name="nach">${opt(konten, v.nach)}</select></label>
    </div>
    <label class="feld"><span>Bereich</span><select name="bereich">${opt(m.listen?.bereiche || [], v.bereich, 'bitte wählen')}</select></label>
    <label class="feld"><span>Topf <small>(nur wenn Geld aufs Tagesgeld geht oder von dort kommt)</small></span><select name="topf">${opt(toepfe, v.topf, 'kein Topf')}</select></label>
    <label class="feld"><span>Verwendungszweck <small>(optional)</small></span><input name="beschreibung" autocomplete="off" value="${h(v.beschreibung)}"></label>
    <label class="feld"><span>Kommentar <small>(optional)</small></span><input name="kommentar" autocomplete="off" value="${h(v.kommentar)}"></label>
    <label class="feld"><span>Gezahlt von <small>(nur wenn es nicht vom eigenen Konto kommt, z. B. Elterngeld)</small></span><select name="gezahltVon">${opt(['Kai', 'Natalie'], v.gezahltVon, '–')}</select></label>
    <p class="zusammen" id="zusammen"></p>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" type="submit" id="eintragen-los">Eintragen</button><button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div>
  </form>`;
  const form = document.getElementById('eintragen-form');
  const zusammen = () => {
    const d = Object.fromEntries(new FormData(form));
    const b = betragLesen(d.betrag);
    document.getElementById('zusammen').textContent = isFinite(b) && b > 0
      ? `${d.art} über ${eur(b)} von ${d.von} an ${d.nach}, gilt für ${monatsname(d.giltFuer)}.` : '';
  };
  form.addEventListener('input', zusammen);
  form.addEventListener('change', (ev) => {
    if (ev.target.name === 'art') { const s = STANDARD[ev.target.value](kn); form.von.value = s.von; form.nach.value = s.nach; zusammen(); }
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

document.addEventListener('submit', async (ev) => {
  if (ev.target.id !== 'eintragen-form') return;
  ev.preventDefault();
  const form = ev.target;
  const d = Object.fromEntries(new FormData(form));
  const betrag = betragLesen(d.betrag);
  const fehler = !(betrag > 0) ? 'Bitte einen Betrag größer als 0 eingeben.'
    : !d.posten.trim() ? 'Bitte eintragen, wofür das Geld war.'
      : !d.datum ? 'Bitte das Datum angeben.'
        : d.von === d.nach ? 'Von und Nach sind dasselbe Konto.'
          : !d.bereich ? 'Bitte einen Bereich wählen.' : null;
  if (fehler) { zeigeEintragen(fehler, d); return; }
  const knopf = document.getElementById('eintragen-los');
  knopf.disabled = true; knopf.textContent = 'Wird eingetragen …';
  try {
    const { nr } = await od.buchungEintragen(od.aktuellerOrt(), { ...d, betrag, posten: d.posten.trim() });
    menu.close();
    z.status = { art: 'info', text: `${nr} eingetragen: ${d.posten.trim()}, ${eur(betrag)}.` };
    zeichne();
    await ausOneDrive({ immer: true });
    z.status = { art: 'info', text: `${nr} eingetragen: ${d.posten.trim()}, ${eur(betrag)}.` };
    zeichne();
  } catch (e) {
    if (e instanceof od.AnmeldungNoetig) { zeigeEintragen('Die Anmeldung ist abgelaufen. Bitte im Menü neu anmelden, dann noch einmal eintragen.', d); return; }
    if (e instanceof od.KeinNetz) { zeigeEintragen('Gerade keine Verbindung zu OneDrive. Es wurde nichts eingetragen, bitte gleich noch einmal versuchen.', d); return; }
    zeigeEintragen('Das Eintragen hat nicht geklappt: ' + e.message + ' Es wurde nichts doppelt gespeichert; bitte in Excel nachsehen, bevor ihr es noch einmal versucht.', d);
  }
});

// ---------- Ereignisse
document.addEventListener('click', async (ev) => {
  if (document.body.classList.contains('pc-modus')) return;
  const t = ev.target.closest('button, [data-monat]');
  if (!t) { if (ev.target === menu) menu.close(); return; }
  if (t.dataset.monat) { z.monat = t.dataset.monat; zeichne(); window.scrollTo(0, 0); return; }
  if (t.dataset.tab) { z.tab = t.dataset.tab; merken('tab', z.tab); zeichne(); window.scrollTo(0, 0); return; }
  if (t.dataset.filter) { z.filter = t.dataset.filter; zeichne(); return; }
  if (t.classList.contains('bu')) { t.setAttribute('aria-expanded', t.getAttribute('aria-expanded') === 'true' ? 'false' : 'true'); return; }
  if (t.id === 'menu-knopf') { zeigeMenu(); return; }
  if (t.id === 'eintragen-knopf') { zeigeEintragen(); return; }
  if (t.id === 'schliessen') { menu.close(); return; }
  if (t.id === 'anmelden') { ms.anmelden({ konto: t.closest('.start') ? null : ms.gemerktesKonto() }); return; }
  if (t.id === 'aktualisieren') { if (menu.open) menu.close(); ausOneDrive({ immer: true }); return; }
  if (t.id === 'abmelden') { ms.abmelden(); menu.close(); z.status = { art: 'info', text: 'Abgemeldet. Ihr seht den zuletzt geladenen Stand.' }; zeichne(); return; }
  if (t.id === 'entfernen') { await loescheDatei(); ms.abmelden(); menu.close(); z.modell = null; z.datei = null; zeigeStart(); }
});

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
start();
