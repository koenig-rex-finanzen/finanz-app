// „Meins“: der private Bereich in „Unser Geld“.
// Daten: private Datei „Mein-Bereich.json“ (Ordner der Person, siehe onedrive.js) + gemeinsame Datei (nur gelesen,
// außer Töpfe ändern und Überweisungen eintragen – das geht über die gemeinsamen Funktionen).
import * as P from './kern/privat.js';
import { eur, monatsname, monatPlus, monatsende, monatsansicht, toepfeStand } from './kern/rechnen.js';
import { leseCsv, csvText, kontoErkennen, kurzname } from './kern/csv.js';

// Bereiche, die schon gebaut sind. Für alle anderen zeigt die App „kommt als Nächstes“.
const FERTIG = new Set(['Natalie']);
const TABS = [
  ['m-uebersicht', 'Übersicht', '<path d="M4 11l8-6 8 6v8a1 1 0 01-1 1h-4v-5H9v5H5a1 1 0 01-1-1z"/>'],
  ['m-konto', 'Konto', '<rect x="3" y="6" width="18" height="13" rx="2.5"/><path d="M3 10h18M7 15h4"/>'],
  ['m-toepfe', 'Töpfe', '<path d="M5 9h14l-1.5 10a1 1 0 01-1 .9h-9a1 1 0 01-1-.9z"/><path d="M8 9V7a4 4 0 018 0v2"/>'],
  ['m-challenges', 'Challenges', '<path d="M8 4h8v5a4 4 0 01-8 0z"/><path d="M8 6H5a3 3 0 003 4M16 6h3a3 3 0 01-3 4M12 13v4M8 20h8"/>'],
];

let C; // Verbindung zur App: { z, app, menu, h, heute, zeichne, ausOneDrive, zeigeEintragen, ms, od, ICON, kopf }
const M = {
  status: 'aus', fehler: null, akt: null, person: null, ordner: null, tab: null, csv: null, csvFehler: null,
  kontoFilter: 'alle', meldung: null, laedt: false, geladenFuer: null,
};
const h = (s) => C.h(s);
const zahl = (t) => { const s = String(t ?? '').trim().replace(/\s|€/g, ''); if (!s) return NaN; return Math.round(Number(s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s) * 100) / 100; };
const betragText = (n) => (n == null || !isFinite(n) ? '' : Number(n).toFixed(2).replace('.', ','));
const datumKurz = (iso) => (iso ? `${Number(iso.slice(8, 10))}.${Number(iso.slice(5, 7))}.` : '');
const datumLang = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
const vz = (n) => (n > 0 ? '+' : '') + eur(n);
const merk = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch {} };
const hol = (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } };
const konto = () => C.ms.gemerktesKonto() || 'ohne';
const d = () => M.akt?.daten;

export function init(ctx) { C = ctx; M.tab = hol('ug.mtab') || 'm-uebersicht'; }

/** Thema (Farben) passend zum Bereich setzen. */
export function thema(an) {
  if (an && FERTIG.has(M.person)) document.documentElement.dataset.bereich = 'natalie';
  else delete document.documentElement.dataset.bereich;
}
export const verfuegbar = () => C.ms.aktiv() && C.ms.angemeldet() && C.z.datei?.quelle === 'onedrive';

// ---------------------------------------------------------------- Laden
export async function laden({ immer = false } = {}) {
  if (!verfuegbar() || M.laedt) return;
  M.person = hol('ug.person.' + konto());
  if (!M.person) { M.status = 'person'; C.zeichne(); return; }
  if (!FERTIG.has(M.person)) { M.status = 'spaeter'; C.zeichne(); return; }
  // Zwischenspeicher (offline)
  const cache = hol('ug.privat.' + konto());
  if (!M.akt && cache?.person === M.person) { M.akt = cache.akt; M.status = 'bereit'; }
  if (!immer && M.geladenFuer === konto() && M.status === 'bereit') return;
  M.laedt = true;
  try {
    M.ordner = await C.od.privatOrdnerFinden(M.person);
    if (!M.ordner) { M.status = 'kein-ordner'; return; }
    const akt = await C.od.privatLaden(M.ordner);
    if (!akt) { M.status = 'einrichten'; M.akt = null; return; }
    P.vervollstaendigen(akt.daten);
    M.akt = akt; M.status = 'bereit'; M.fehler = null; M.geladenFuer = konto();
    merk('ug.privat.' + konto(), { person: M.person, akt });
    await auszeichnungenPruefen();
  } catch (e) {
    if (M.status !== 'bereit') { M.status = 'fehler'; }
    M.fehler = e instanceof C.od.AnmeldungNoetig ? 'Die Microsoft-Anmeldung ist abgelaufen. Bitte im Menü neu anmelden.'
      : e instanceof C.od.KeinNetz ? 'Gerade keine Verbindung. Du siehst den zuletzt geladenen Stand.' : e.message;
  } finally { M.laedt = false; C.zeichne(); }
}
export function vergessen() { M.akt = null; M.status = 'aus'; M.geladenFuer = null; merk('ug.privat.' + konto(), null); }

/**
 * Ändert die private Datei: lädt den neuesten Stand, wendet `fn` an, speichert mit If-Match.
 * Bei gleichzeitiger Änderung auf einem anderen Gerät einmal neu versuchen.
 */
async function aendern(fn) {
  for (let versuch = 0; ; versuch++) {
    const frisch = await C.od.privatLaden(M.ordner);
    if (!frisch) throw new Error('Die private Datei wurde nicht gefunden.');
    P.vervollstaendigen(frisch.daten);
    const daten = structuredClone(frisch.daten);
    const ergebnis = fn(daten);
    daten.geaendert = new Date().toISOString();
    try {
      const it = await C.od.privatSpeichern(frisch, daten);
      M.akt = { ...frisch, ...it, daten };
      merk('ug.privat.' + konto(), { person: M.person, akt: M.akt });
      return ergebnis;
    } catch (e) { if (!(e instanceof C.od.Konflikt) || versuch >= 2) throw e; }
  }
}
const fehlerText = (e) => (e instanceof C.od.AnmeldungNoetig ? 'Die Anmeldung ist abgelaufen. Bitte im Menü neu anmelden. Es wurde nichts gespeichert.'
  : e instanceof C.od.KeinNetz ? 'Gerade keine Verbindung zu OneDrive. Es wurde nichts gespeichert, bitte gleich noch einmal versuchen.'
    : 'Das hat nicht geklappt: ' + e.message);

async function speichern(fn, meldung, { knopf, dialogFehler } = {}) {
  if (knopf) { knopf.disabled = true; knopf.dataset.text = knopf.textContent; knopf.textContent = 'Wird gespeichert …'; }
  try {
    const r = await aendern(fn);
    if (C.menu.open && !dialogFehler?.offenLassen) C.menu.close();
    M.meldung = typeof meldung === 'function' ? meldung(r) : meldung;
    await auszeichnungenPruefen();
    C.zeichne();
    return true;
  } catch (e) {
    if (dialogFehler) dialogFehler(fehlerText(e));
    else { M.meldung = { art: 'warn', text: fehlerText(e) }; C.zeichne(); }
    if (knopf) { knopf.disabled = false; knopf.textContent = knopf.dataset.text; }
    return false;
  }
}

// ---------------------------------------------------------------- Ruhmeshalle: neue Auszeichnungen
async function auszeichnungenPruefen() {
  if (!d()) return;
  const neu = P.neueAuszeichnungen(d(), { modell: C.z.modell, p: d(), heute: C.heute() });
  if (!neu.length) return;
  try {
    await aendern((x) => { const da = new Set(x.ruhm.map((r) => r.id)); for (const n of neu) if (!da.has(n.id)) x.ruhm.push(n); });
    M.meldung = { art: 'feier', text: (neu.some((n) => n.art === 'pokal') ? '🏆 ' : '🏅 ') + neu.map((n) => n.titel).join(' · ') + ' – steht jetzt in deiner Ruhmeshalle.' };
    setTimeout(konfetti, 150);
  } catch { /* nächstes Mal */ }
}

export function konfetti() {
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const cv = document.createElement('canvas');
  cv.className = 'konfetti'; document.body.appendChild(cv);
  const ctx = cv.getContext('2d');
  const dpr = devicePixelRatio || 1;
  const W = cv.width = innerWidth * dpr, H = cv.height = innerHeight * dpr;
  const farben = ['#E46A9E', '#F4A6C6', '#5B8FD6', '#9CC1F0', '#47A985', '#9ADBBE', '#F6C85F'];
  const teile = Array.from({ length: 160 }, () => ({
    x: W / 2 + (Math.random() - 0.5) * W * 0.3, y: H * 0.35, vx: (Math.random() - 0.5) * 16 * dpr, vy: (-Math.random() * 15 - 6) * dpr,
    g: 0.42 * dpr, w: (6 + Math.random() * 6) * dpr, hh: (8 + Math.random() * 8) * dpr, r: Math.random() * 6, vr: (Math.random() - 0.5) * 0.35,
    f: farben[Math.floor(Math.random() * farben.length)], rund: Math.random() < 0.3,
  }));
  let t0 = performance.now();
  const schritt = (t) => {
    const dt = Math.min(2, (t - t0) / 16.7); t0 = t;
    ctx.clearRect(0, 0, W, H);
    let lebt = false;
    for (const s of teile) {
      s.vy += s.g * dt; s.vx *= 0.99; s.x += s.vx * dt; s.y += s.vy * dt; s.r += s.vr * dt;
      if (s.y < H + 40) lebt = true;
      ctx.save(); ctx.translate(s.x, s.y); ctx.rotate(s.r); ctx.fillStyle = s.f;
      if (s.rund) { ctx.beginPath(); ctx.arc(0, 0, s.w / 2, 0, 7); ctx.fill(); } else ctx.fillRect(-s.w / 2, -s.hh / 2, s.w, s.hh * Math.abs(Math.cos(s.r)));
      ctx.restore();
    }
    if (lebt) requestAnimationFrame(schritt); else cv.remove();
  };
  requestAnimationFrame(schritt);
}

// ---------------------------------------------------------------- Zeichnen
export function zeichne() {
  thema(true);
  const kopf = C.kopf();
  const wahl = (inhalt, mitNav = true) => {
    C.app.innerHTML = `${kopf}<main class="wrap meins">${M.fehler && M.status === 'bereit' ? `<div class="hinweis info">${h(M.fehler)}</div>` : ''}${inhalt}</main>${mitNav ? nav() : ''}`;
  };
  if (M.status === 'aus' || (M.laedt && !d())) { wahl(`<div class="leer m-laedt">Dein Bereich wird geladen …</div>`, false); if (M.status === 'aus') laden(); return; }
  if (M.status === 'person') { wahl(personWahl(), false); return; }
  if (M.status === 'spaeter') { wahl(spaeter(), false); return; }
  if (M.status === 'kein-ordner') { wahl(keinOrdner(), false); return; }
  if (M.status === 'einrichten') { wahl(einrichten(), false); return; }
  if (M.status === 'fehler' && !d()) { wahl(`<div class="hinweis warn">${h(M.fehler || 'Unbekannter Fehler')}</div><button class="knopf" data-m="neu-laden">Noch mal versuchen</button>`, false); return; }
  const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
  const meldung = M.meldung ? `<div class="hinweis ${M.meldung.art}">${h(M.meldung.text)}${M.meldung.knopf ? ` <button class="link" data-m="${M.meldung.knopf[0]}">${h(M.meldung.knopf[1])}</button>` : ''}</div>` : '';
  const inhalt = { 'm-uebersicht': uebersicht, 'm-konto': kontoAnsicht, 'm-toepfe': toepfeAnsicht, 'm-challenges': challengesAnsicht }[M.tab] || uebersicht;
  wahl(meldung + inhalt(mp));
  C.app.insertAdjacentHTML('beforeend', `<button class="fab" data-m="eintragen" aria-label="Eintragen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Eintragen</button>`);
}
const nav = () => `<nav><div class="wrap" role="tablist">${TABS.map(([id, name, pfad]) =>
  `<button role="tab" aria-selected="${M.tab === id}" data-m="tab" data-arg="${id}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${pfad}</svg>${name}</button>`).join('')}</div></nav>`;

function personWahl() {
  const leute = P.personen(C.z.modell);
  return `<section class="m-start"><h2>Wer bist du?</h2>
    <p class="hint">Damit die App weiß, welchen privaten Bereich sie laden soll. Das merkt sich nur dieses Gerät für dein Microsoft-Konto (${h(konto())}).</p>
    <div class="knoepfe-liste">${leute.map((p) => `<button class="knopf leise gross-knopf" data-m="person" data-arg="${h(p)}">${h(p)}</button>`).join('')}</div></section>`;
}
function spaeter() {
  return `<section class="m-start"><h2>Dein Bereich kommt als Nächstes</h2>
    <p>Der private Bereich für ${h(M.person)} wird mit eigenen Wünschen gebaut (z. B. Vorschau mit Minus-Warnung). Bis dahin ist hier noch nichts.</p>
    <p class="hint" style="margin-top:12px">Falsch gewählt? <button class="link" data-m="person-wechseln">Andere Person wählen</button></p></section>`;
}
function keinOrdner() {
  return `<section class="m-start"><h2>Dein Ordner ist noch nicht da</h2>
    <p>Die App sucht den Ordner <b>„${h(M.person)}“</b>, den Kai in seinem OneDrive unter <b>Apps</b> angelegt und mit dir geteilt hat.</p>
    <p class="hint" style="margin-top:8px">Hat Kai ihn schon geteilt? Dann öffne einmal den Link aus seiner Einladung (Mail von OneDrive), danach hier auf „Noch mal suchen“ tippen.</p>
    <div class="knoepfe-liste" style="margin-top:16px"><button class="knopf" data-m="neu-laden">Noch mal suchen</button>
    <button class="knopf leise" data-m="ersatz-ordner">Stattdessen in meinem eigenen OneDrive anlegen</button></div>
    <p class="hint" style="margin-top:12px">Falsche Person? <button class="link" data-m="person-wechseln">Andere Person wählen</button></p></section>`;
}
function einrichten(fehler) {
  const ort = M.ordner?.quelle === 'geteilt' ? `im Ordner „${h(M.ordner.name)}“, den Kai mit dir geteilt hat` : M.ordner?.quelle === 'eigen' ? `im Ordner Apps › ${h(M.ordner.name)}` : 'in deinem eigenen OneDrive (Ordner „Unser-Geld-Privat“)';
  return `<section class="m-start"><h2>Hallo ${h(M.person)}! Dein Bereich wird eingerichtet</h2>
    <p>Hier hast du dein Giro, PayPal, deine Töpfe, Limits, Erinnerungen und Spar-Challenges. Die Daten liegen ${ort} in der Datei „${h(C.od.PRIVAT_DATEI)}“.</p>
    <form id="m-einrichten" class="m-form" novalidate>
      <p class="hint">Damit die App deinen Giro-Stand ausrechnen kann, braucht sie einmal den Stand laut Bank und das Datum dazu. Danach zählt sie deine Buchungen und alles aus der gemeinsamen Datei, was dein Giro betrifft, selbst dazu.</p>
      <label class="feld"><span>Stand Giro laut Bank in €</span><input name="anfang" inputmode="decimal" placeholder="z. B. 412,37" autocomplete="off"></label>
      <label class="feld"><span>Am Ende von diesem Tag</span><input type="date" name="stichtag" value="${C.heute()}"></label>
      <p class="hint klein">Alles, was bis einschließlich zu diesem Tag gebucht wurde, steckt schon im Stand. Weißt du es gerade nicht? Leer lassen, das geht später im Menü ⋮.</p>
      ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
      <button class="knopf" type="submit">Bereich anlegen</button>
    </form></section>`;
}

// ---------- Übersicht
function uebersicht(mp) {
  const s = mp.summen;
  const name = monatsname(mp.monat);
  const vergangen = monatsende(mp.monat) < C.heute();
  const frei = `<strong class="${s.frei < 0 ? 'neg' : ''}">${eur(s.frei)}</strong>`;
  const satz = mp.laufend ? `Dir bleiben im ${h(name)} noch ${frei} frei.` : vergangen ? `Im ${h(name)} sind dir ${frei} geblieben.` : `Im ${h(name)} bleiben dir voraussichtlich ${frei} frei.`;
  const chip = mp.laufend ? `<span class="check">nach allen festen Kosten und Überweisungen · noch ${mp.tageUebrig} ${mp.tageUebrig === 1 ? 'Tag' : 'Tage'}</span>` : `<span class="check offen">nach allen festen Kosten und Überweisungen</span>`;

  const erin = P.erinnerungen(d(), mp, C.heute()).filter((e) => e.faellig);
  const erinHtml = erin.length ? `<section class="m-erin">${erin.map((e) => `<div class="erin ${e.ueberfaellig ? 'ueber' : ''}">
      <span class="glocke" aria-hidden="true">🔔</span><div><b>${h(e.text)}</b>${e.auto && e.betrag ? `<div class="hint">${eur(e.betrag)} ${e.betrag < e.soll ? 'fehlen noch' : ''}</div>` : ''}</div>
      ${e.auto ? `<button class="knopf klein" data-m="ueberweisung" data-arg="${h(e.nach)}">Eintragen</button>` : `<button class="knopf klein leise" data-m="erin-erledigt" data-arg="${h(e.id)}">Erledigt</button>`}</div>`).join('')}</section>` : '';

  const zeile = (t, b, cls = '', sub = '') => `<div class="rz ${cls}"><span>${t}${sub ? `<small>${sub}</small>` : ''}</span><b>${b}</b></div>`;
  const rechnung = `<section><h2>So rechnet die App</h2><div class="rechnung">
    ${zeile('Einnahmen aufs Giro', vz(s.einnahmen + s.sonstigeEin), 'plus', mp.einnahmen.length ? mp.einnahmen.map((e) => e.posten).join(', ') : 'keine geplant')}
    ${zeile('Überweisungen an gemeinsame Konten', eur(-s.ueberweisungen), '', mp.ueberweisungen.map((u) => `${u.nach} ${eur(u.betragMonat)}`).join(', ') || 'keine')}
    ${zeile('Feste Kosten', eur(-s.feste), '', `${mp.feste.length} Posten${mp.feste.some((f) => f.geschaetzt) ? ', teils geschätzt' : ''}`)}
    ${zeile('Schon ausgegeben', eur(-s.ausgegeben), '', s.ausTopf ? `ohne ${eur(s.ausTopf)} aus Töpfen bezahlt` : 'deine eingetragenen Ausgaben')}
    <div class="rz summe"><span>Bleibt frei</span><b class="${s.frei < 0 ? 'neg' : ''}">${eur(s.frei)}</b></div></div>
    ${mp.unbekannt.length ? `<p class="hint" style="margin-top:8px">Nicht eingerechnet, weil der Monat unbekannt ist: ${mp.unbekannt.map((u) => `${h(u.posten)} (${eur(u.betrag)}, ${h(u.rhythmus)})`).join(', ')}.</p>` : ''}</section>`;

  const v = monatsansicht(C.z.modell, mp.monat, C.heute());
  const spar = v.konten.filter((k) => k.rolle === 'Sparen');
  const meine = P.meineToepfe(C.z.modell, d());
  const meineSumme = v.toepfe.filter((t) => meine.includes(t.name)).reduce((x, t) => x + t.stand, 0);
  const konten = `<section><h2>Deine Konten</h2><div class="konten">
    <div class="konto"><span class="dot" style="background:var(--rosa)"></span><div><div class="name">${h(mp.giro)}</div><div class="sub">${mp.giroStand == null ? 'Stand fehlt noch – im Menü ⋮ eintragen' : `berechnet, ${mp.laufend ? 'heute' : 'am ' + datumLang(mp.giroStichtag)}`}</div></div><div class="betrag ${mp.giroStand < 0 ? 'neg' : ''}">${mp.giroStand == null ? '–' : eur(mp.giroStand)}</div></div>
    <div class="konto"><span class="dot" style="background:var(--blau)"></span><div><div class="name">PayPal</div><div class="sub">im ${h(name)} bezahlt, geht vom Giro ab</div></div><div class="betrag">${eur(s.paypal)}</div></div>
    ${spar.map((k) => `<div class="konto"><span class="dot" style="background:var(--gruen)"></span><div><div class="name">${h(k.name)}</div><div class="sub">aus der gemeinsamen Datei · deine Töpfe ${eur(Math.round(meineSumme * 100) / 100)}</div></div><div class="betrag">${eur(v.vergangen ? k.standMonatsende : k.standJetzt)}</div></div>`).join('')}
  </div></section>`;

  const kommt = mp.kommt.length ? `<section><h2>${vergangen ? 'Nicht erledigt' : 'Was noch kommt'}</h2><div class="kommt">${mp.kommt.map((k) => {
    const knopf = k.art === 'einnahme' ? `<button class="knopf klein leise" data-m="plan-ist" data-arg="${h(k.id)}">Ist da</button>`
      : k.art === 'fest' ? `<button class="knopf klein leise" data-m="plan-ist" data-arg="${h(k.id)}">Bezahlt</button>`
        : `<button class="knopf klein leise" data-m="ueberweisung" data-arg="${h(k.nach)}">Eintragen</button>`;
    return `<div class="k-zeile ${k.ueberfaellig ? 'ueber' : ''}"><span class="k-tag">${k.tag ? `<b>${k.tag}.</b>` : '<b>–</b>'}</span>
      <div><div class="t">${h(k.titel)}</div><div class="hint">${k.art === 'einnahme' ? 'Einnahme' : k.art === 'ueberweisung' ? h(k.wann) : `feste Kosten${k.geschaetzt ? ', geschätzt' : ''}`}${k.quelle === 'gemeinsam' ? ' · aus der gemeinsamen Datei' : ''}</div></div>
      <b class="${k.betrag > 0 ? 'plus' : ''}">${vz(k.betrag)}</b>${knopf}</div>`;
  }).join('')}</div></section>` : '';

  const bereiche = mp.bereiche.length ? `<section><h2>Ausgaben nach Bereich</h2><div class="bereiche-liste">${mp.bereiche.map((b) => `<div class="ber ${b.status || ''}">
      <div class="ber-kopf"><span>${h(b.bereich)}</span><b>${eur(b.betrag)}${b.limit != null ? `<small> von ${eur(b.limit)}</small>` : ''}</b></div>
      ${b.limit != null ? `<div class="balken"><span style="width:${Math.min(100, b.anteil * 100)}%"></span></div>
      <div class="hint klein">${b.status === 'erreicht' ? 'Limit erreicht' : b.status === 'fast' ? `Fast erreicht, noch ${eur(b.limit - b.betrag)}` : `noch ${eur(b.limit - b.betrag)}`}</div>` : ''}</div>`).join('')}</div>
    <p class="hint" style="margin-top:8px"><button class="link" data-m="limits">Limits einstellen</button></p></section>`
    : `<section><h2>Ausgaben nach Bereich</h2><div class="leer">Im ${h(name)} hast du noch nichts eingetragen. <button class="link" data-m="limits">Limits einstellen</button></div></section>`;

  return `<div class="hero m-hero"><p>${satz}</p>${chip}</div>${erinHtml}${rechnung}${kommt}${konten}${bereiche}`;
}

// ---------- Konto
function kontoAnsicht(mp) {
  const von = mp.monat + '-01', bis = monatsende(mp.monat);
  const liste = P.giroBewegungen(C.z.modell, d(), von, bis)
    .filter((x) => M.kontoFilter === 'alle' || (M.kontoFilter === 'PayPal' ? x.b.weg === 'PayPal' : M.kontoFilter === 'gemeinsam' ? x.quelle === 'gemeinsam' : x.quelle === 'privat' && x.b.weg !== 'PayPal'));
  const f = (id, n) => `<button aria-pressed="${M.kontoFilter === id}" data-m="kfilter" data-arg="${id}">${n}</button>`;
  const zeilen = liste.map((x) => {
    const b = x.b;
    if (x.quelle === 'gemeinsam') {
      return `<div class="bu-wrap"><button class="bu" aria-expanded="false"><span class="t">${h(b.posten || b.beschreibung)}</span><span class="betrag ${x.betrag > 0 ? 'plus' : ''}">${vz(x.betrag)}</span>
        <span class="r">${datumKurz(b.datum)} <span class="chip c-gem">gemeinsame Datei</span> ${h(b.von === mp.giro ? 'an ' + b.nach : 'von ' + b.von)}</span>
        <span class="more"><dl><dt>Nummer</dt><dd>${h(b.nr)}</dd>${b.topf ? `<dt>Topf</dt><dd>${h(b.topf)}</dd>` : ''}${b.kommentar ? `<dt>Kommentar</dt><dd>${h(b.kommentar)}</dd>` : ''}<dt>Hinweis</dt><dd>Steht in der gemeinsamen Datei. Ändern im Bereich „Gemeinsam“.</dd></dl></span></button></div>`;
    }
    const plan = b.plan ? [...mp.feste, ...mp.einnahmen].find((p) => p.id === b.plan) : null;
    return `<div class="bu-wrap"><button class="bu" aria-expanded="false"><span class="t">${h(b.posten)}${b.ausTopf ? ' <span class="chip aus-topf">aus Topf</span>' : ''}</span><span class="betrag ${x.betrag > 0 ? 'plus' : ''}">${vz(x.betrag)}</span>
      <span class="r">${datumKurz(b.datum)} <span class="chip ${b.weg === 'PayPal' ? 'c-pp' : 'c-giro'}">${h(b.weg || 'Giro')}</span>${b.bereich ? ` ${h(b.bereich)}` : ''}${b.monat && b.monat !== b.datum.slice(0, 7) ? ` <span class="okt">gilt für ${h(monatsname(b.monat))}</span>` : ''}</span>
      <span class="more"><dl><dt>Nummer</dt><dd>${h(b.nr)}</dd>${plan ? `<dt>Gehört zu</dt><dd>${h(plan.posten)}</dd>` : ''}${b.ausTopf ? `<dt>Aus Topf</dt><dd>${h(b.ausTopf)}, zählt nicht gegen „frei“</dd>` : ''}${b.kommentar ? `<dt>Kommentar</dt><dd>${h(b.kommentar)}</dd>` : ''}<dt>Eingetragen</dt><dd>${b.quelle === 'CSV' ? 'aus dem Kontoauszug' : 'von Hand'}</dd></dl></span></button>
      <div class="bu-aktion"><button class="knopf leise klein" data-m="bearbeiten" data-arg="${h(b.nr)}">Bearbeiten</button><button class="knopf leise klein" data-m="loeschen-frage" data-arg="${h(b.nr)}">Löschen</button></div></div>`;
  }).join('');
  return `<section><h2>${h(mp.giro)}</h2>
    <div class="m-giro"><div><div class="hint">Stand ${mp.laufend ? 'heute' : 'am ' + datumLang(mp.giroStichtag)}, berechnet</div><div class="gross ${mp.giroStand < 0 ? 'neg' : ''}">${mp.giroStand == null ? '–' : eur(mp.giroStand)}</div>
      <div class="hint">${d().giro.stichtag ? `ab ${eur(d().giro.anfang)} am ${datumLang(d().giro.stichtag)}` : 'Anfangsstand fehlt noch'} · <button class="link" data-m="giro-start">ändern</button></div></div>
      <label class="knopf leise klein" for="m-csv">Kontoauszug abgleichen</label><input type="file" id="m-csv" accept=".csv,text/csv" hidden></div>
    ${abgleichHtml()}
  </section>
  <section><h2>Buchungen im ${h(monatsname(mp.monat))}</h2>
    <div class="filter" role="group">${f('alle', 'Alle')}${f('Giro', 'Giro')}${f('PayPal', 'PayPal')}${f('gemeinsam', 'Aus der gemeinsamen Datei')}</div>
    ${liste.length ? `<div class="buchungen">${zeilen}</div><p class="hint" style="margin-top:8px">Antippen zeigt Details. Was ein gemeinsames Konto oder einen Topf betrifft, steht in der gemeinsamen Datei und wird hier nur mitgezählt.</p>` : `<div class="leer">Keine Buchungen in diesem Monat.</div>`}
  </section>
  <section><h2>Feste Kosten und Einnahmen</h2>${festeHtml(mp)}</section>`;
}

function festeHtml() {
  const giro = d().giro.konto;
  const gem = [...C.z.modell.kosten.filter((k) => k.konto === giro).map((k) => ({ ...k, art: 'Ausgabe' })), ...C.z.modell.einnahmen.filter((e) => e.konto === giro).map((e) => ({ ...e, art: 'Einnahme', rhythmus: 'monatlich' }))];
  const zeile = (f, eigen) => `<div class="f-zeile"><div><div class="t">${h(f.posten)}</div><div class="hint">${f.art === 'Einnahme' ? 'Einnahme' : 'Ausgabe'} · ${h(f.rhythmus || 'monatlich')}${f.tag ? ` · am ${f.tag}.` : ''}${f.faellig ? ` · ${h(f.faellig)}` : ''}${f.ab ? ` · ab ${monatsname(f.ab, true)}` : ''}${f.bis ? ` · bis ${monatsname(f.bis, true)}` : ''}</div></div>
    <b class="${f.art === 'Einnahme' ? 'plus' : ''}">${eur(f.betrag)}</b>${eigen ? `<button class="knopf leise klein" data-m="fest-bearbeiten" data-arg="${h(f.id)}">Ändern</button>` : '<span class="chip c-gem">gemeinsam</span>'}</div>`;
  return `<div class="feste">${gem.map((f) => zeile(f, false)).join('')}${d().feste.map((f) => zeile(f, true)).join('')}</div>
    <p class="hint" style="margin-top:8px">„gemeinsam“ steht in der gemeinsamen Datei (Blätter Kosten und Einnahmen) und wird dort geändert.</p>
    <button class="knopf leise" style="margin-top:10px" data-m="fest-neu">Eigenen festen Posten anlegen</button>`;
}

function abgleichHtml() {
  if (M.csvFehler) return `<p class="fehler">${h(M.csvFehler)}</p>`;
  const a = M.csv;
  if (!a) return '<p class="hint" style="margin-top:8px">Kontoauszug als CSV aus dem VW-Banking laden: Die App zeigt nur, was fehlt oder abweicht. Übernommen wird nichts ohne deinen Klick.</p>';
  const neu = a.zeilen.filter((r) => r.status === 'neu');
  const ok = a.zeilen.filter((r) => r.status === 'eingetragen').length;
  const vor = a.zeilen.filter((r) => r.status === 'vor-stichtag').length;
  return `<div class="m-abgleich">
    <div class="m-ab-kopf"><b>Abgleich ${datumLang(a.csv.von)} bis ${datumLang(a.csv.bis)}</b><button class="link" data-m="csv-zu">schließen</button></div>
    <div class="topfsumme">
      <div class="z"><span>Schon eingetragen</span><b>${ok}</b></div>
      <div class="z"><span>Fehlt noch</span><b style="color:${neu.length ? 'var(--warn)' : 'var(--gem)'}">${neu.length}</b></div>
      ${vor ? `<div class="z"><span>Vor deinem Anfangsstand</span><b>${vor}</b></div>` : ''}
      <div class="z"><span>Stand laut Bank</span><b>${eur(a.saldo.bank)}</b></div>
      <div class="z"><span>Stand laut App am ${datumLang(a.saldo.bis)}</span><b>${a.saldo.app == null ? '–' : eur(a.saldo.app)}</b></div>
      ${a.saldo.differenz != null ? `<div class="z"><span>Unterschied</span><b style="color:${Math.abs(a.saldo.differenz) < 0.005 ? 'var(--gem)' : 'var(--warn)'}">${Math.abs(a.saldo.differenz) < 0.005 ? 'stimmt ✓' : vz(a.saldo.differenz)}</b></div>` : ''}
    </div>
    ${neu.length ? `<h3>Steht auf dem Auszug, fehlt bei dir</h3><div class="buchungen">${neu.map((r) => `<div class="m-ab-zeile"><div><div class="t">${h(kurzname(r.z.info, r.z.umsatzart))}</div><div class="hint">${datumKurz(r.z.datum)} · ${h(r.z.info.slice(0, 80))}</div>
      ${r.gemeinsamKonto ? `<div class="hint warnfarbe">Betrifft ${h(r.gemeinsamKonto)} – gehört in die gemeinsame Datei.</div>` : ''}</div>
      <b class="${r.z.betrag > 0 ? 'plus' : ''}">${vz(r.z.betrag)}</b>
      ${r.gemeinsamKonto ? `<button class="knopf leise klein" data-m="csv-gemeinsam" data-arg="${h(r.id)}">Gemeinsam eintragen</button>` : `<button class="knopf leise klein" data-m="csv-uebernehmen" data-arg="${h(r.id)}">Übernehmen</button>`}</div>`).join('')}</div>` : '<p class="hint" style="margin-top:8px">Alles auf dem Auszug ist eingetragen.</p>'}
    ${a.fehlen.length ? `<h3>Eingetragen, aber nicht auf dem Auszug</h3><div class="buchungen">${a.fehlen.map((k) => `<div class="m-ab-zeile"><div><div class="t">${h(k.b.posten || k.b.beschreibung)}</div><div class="hint">${datumKurz(k.datum)} · ${k.quelle === 'gemeinsam' ? 'gemeinsame Datei ' + h(k.b.nr) : h(k.b.weg || 'Giro') + ' ' + h(k.b.nr)}</div></div><b>${vz(k.betrag)}</b></div>`).join('')}</div>
      <p class="hint" style="margin-top:6px">Vielleicht noch nicht abgebucht, ein anderes Datum oder ein Tippfehler im Betrag.</p>` : ''}
  </div>`;
}

// ---------- Töpfe
function toepfeAnsicht(mp) {
  const heute = C.heute();
  const bis = mp.laufend ? heute : monatsende(mp.monat);
  const alle = toepfeStand(C.z.modell, bis);
  const meine = P.meineToepfe(C.z.modell, d());
  const v = monatsansicht(C.z.modell, mp.monat, heute);
  const spar = v.konten.filter((k) => k.rolle === 'Sparen');
  const karte = (t) => {
    const ziel = P.zielStand(t.stand, d().ziele[t.name], heute);
    return `<div class="m-topf"><button class="m-topf-kopf" data-topf="${h(t.name)}"><span class="n">${h(t.name)}</span><span class="b ${t.stand < 0 ? 'neg' : ''}">${eur(t.stand)}</span></button>
      <div class="hint">${t.sparrate ? `${eur(t.sparrate)} im Monat` : 'keine feste Sparrate'}${t.bemerkung ? ' · ' + h(t.bemerkung) : ''}</div>
      ${ziel ? `<div class="balken ziel"><span style="width:${ziel.anteil * 100}%"></span></div>
        <div class="hint klein">${ziel.geschafft ? `Ziel ${eur(ziel.betrag)} erreicht 🎉` : `Ziel ${eur(ziel.betrag)}${ziel.datum ? ` bis ${datumLang(ziel.datum)}` : ''} · fehlen ${eur(ziel.fehlt)}${ziel.rate != null ? ` · dafür ${eur(ziel.rate)} im Monat` : ''}`}</div>` : ''}
      <div class="m-topf-knoepfe"><button class="knopf leise klein" data-m="topf-bearbeiten" data-arg="${h(t.name)}">Bearbeiten</button><button class="knopf leise klein" data-topf="${h(t.name)}">Buchungen</button></div></div>`;
  };
  const meineListe = alle.filter((t) => meine.includes(t.name));
  return `<section><h2>Deine Töpfe</h2>
    <div class="topfsumme">${spar.map((k) => `<div class="z"><span>${h(k.name)}</span><b>${eur(v.vergangen ? k.standMonatsende : k.standJetzt)}</b></div>`).join('')}
      <div class="z"><span>Davon in deinen Töpfen</span><b>${eur(Math.round(meineListe.reduce((s, t) => s + t.stand, 0) * 100) / 100)}</b></div>
      <div class="z"><span>Deine Sparraten im Monat</span><b>${eur(Math.round(meineListe.reduce((s, t) => s + t.sparrate, 0) * 100) / 100)}</b></div></div>
    <p class="hint" style="margin-top:8px">Stand ${mp.laufend ? 'heute' : `Ende ${h(monatsname(mp.monat))}`}, direkt aus der gemeinsamen Datei. Name, Sparrate und Bemerkung änderst du hier für beide Ansichten zugleich.</p>
    <div class="m-toepfe">${meineListe.map(karte).join('') || '<div class="leer">Noch keine Töpfe als deine markiert.</div>'}</div>
    <div class="knoepfe-liste" style="margin-top:14px"><button class="knopf leise" data-m="topf-neu">Neuen Topf anlegen</button><button class="knopf leise" data-m="meine-toepfe">Welche Töpfe sind meine?</button></div>
  </section>`;
}

// ---------- Challenges + Ruhmeshalle
const POKAL = (stufe) => `<svg viewBox="0 0 48 48" class="pokal ${stufe}" aria-hidden="true"><path d="M14 6h20v10a10 10 0 01-20 0z" fill="var(--p1)"/><path d="M14 9H8a6 6 0 006 8M34 9h6a6 6 0 01-6 8" fill="none" stroke="var(--p1)" stroke-width="3"/><path d="M22 26h4v8h-4z" fill="var(--p2)"/><rect x="15" y="34" width="18" height="6" rx="2" fill="var(--p2)"/><path d="M19 10q2 6 6 7" stroke="#fff" stroke-opacity=".55" stroke-width="2.4" fill="none" stroke-linecap="round"/></svg>`;
const MEDAILLE = (stufe) => `<svg viewBox="0 0 48 48" class="medaille ${stufe}" aria-hidden="true"><path d="M15 4h7l4 14h-7zM33 4h-7l-4 14h7z" fill="var(--band)"/><circle cx="24" cy="30" r="12" fill="var(--p1)"/><circle cx="24" cy="30" r="8" fill="none" stroke="var(--p2)" stroke-width="2"/><path d="M24 25l1.6 3.3 3.6.5-2.6 2.5.6 3.6-3.2-1.7-3.2 1.7.6-3.6-2.6-2.5 3.6-.5z" fill="var(--p2)"/></svg>`;

function challengesAnsicht() {
  const ctx = { modell: C.z.modell, p: d(), heute: C.heute() };
  const ruhm = [...d().ruhm].sort((a, b) => b.datum.localeCompare(a.datum));
  const halle = `<section><h2>Ruhmeshalle</h2><div class="halle">${ruhm.length ? ruhm.map((r) => `<div class="preis" title="${h(r.titel)}">${r.art === 'pokal' ? POKAL(r.stufe) : MEDAILLE(r.stufe)}<span>${h(r.titel)}</span><small>${datumLang(r.datum)}</small></div>`).join('')
    : '<p class="hint halle-leer">Hier stehen bald deine Pokale und Medaillen. Starte eine Challenge!</p>'}</div></section>`;
  const aktiv = d().challenges.filter((c) => !c.archiviert);
  const karten = aktiv.map((c) => challengeKarte(c, P.challengeStand(c, ctx))).join('');
  const archiv = d().challenges.filter((c) => c.archiviert);
  return `${halle}<section><h2>Deine Challenges</h2>${karten || '<div class="leer">Noch keine Challenge. Alle sind freiwillig, ein Topf kann dazu, muss aber nicht.</div>'}
    <button class="knopf" style="margin-top:14px;width:100%" data-m="challenge-neu">Neue Challenge starten</button></section>
    ${archiv.length ? `<section><h2>Abgeschlossen</h2><div class="feste">${archiv.map((c) => `<div class="f-zeile"><div><div class="t">${h(titelVon(c))}</div><div class="hint">${h(P.CHALLENGE_ARTEN[c.typ].titel)}</div></div><button class="knopf leise klein" data-m="ch-zurueck" data-arg="${h(c.id)}">Zurückholen</button></div>`).join('')}</div></section>` : ''}`;
}
const titelVon = (c) => c.name || (c.typ === 'wochen52' ? '52 Wochen' : c.typ === 'limit' ? `${c.bereich}: höchstens ${eur(c.betrag)} im ${monatsname(c.monat)}` : c.typ === 'nospend' ? `${c.ziel} No-Spend-Tage im ${monatsname(c.monat)}` : P.CHALLENGE_ARTEN[c.typ].titel);

function challengeKarte(c, s) {
  const kopf = `<div class="ch-kopf"><div><div class="ch-art">${h(P.CHALLENGE_ARTEN[c.typ].titel)}${c.topf ? ` · Topf ${h(c.topf)}` : ''}</div><div class="ch-titel">${h(titelVon(c))}</div></div>${s.geschafft ? POKAL('gold') : ''}</div>`;
  const balken = `<div class="balken ch"><span style="width:${Math.round(s.anteil * 1000) / 10}%"></span></div>`;
  let mitte = '';
  if (c.typ === 'ziel') {
    mitte = `${balken}<div class="ch-zahlen"><span><b>${eur(s.stand)}</b> von ${eur(s.betrag)}</span><span>${Math.round(s.anteil * 100)} %</span></div>
      <p class="hint">${s.geschafft ? 'Geschafft! 🎉' : `Es fehlen ${eur(s.fehlt)}.${s.rate != null ? ` Bis ${datumLang(c.datum)} sind das ${eur(s.rate)} im Monat (${s.monate} ${s.monate === 1 ? 'Monat' : 'Monate'}).` : ''}`}</p>
      ${!c.topf && !s.geschafft ? `<button class="knopf leise klein" data-m="ch-einzahlung" data-arg="${h(c.id)}">Zurückgelegt eintragen</button>` : ''}
      ${c.topf ? '<p class="hint klein">Der Stand kommt aus dem Topf in der gemeinsamen Datei.</p>' : ''}`;
  } else if (c.typ === 'wochen52') {
    const erledigt = new Set(c.erledigt || []);
    mitte = `${balken}<div class="ch-zahlen"><span><b>${eur(s.betrag)}</b> von ${eur(s.gesamt)}</span><span>${s.wochen} / 52 Wochen</span></div>
      <p class="hint">Diese Woche: Woche ${s.aktuelleWoche} → ${eur(s.aktuelleWoche)}${s.offen.length ? ` · offen bis jetzt: ${s.offen.length}` : ' · alles erledigt bis jetzt'}</p>
      <div class="wochen">${Array.from({ length: 52 }, (_, i) => i + 1).map((w) => `<button class="woche ${erledigt.has(w) ? 'an' : ''} ${w === s.aktuelleWoche ? 'jetzt' : ''}" data-m="woche" data-arg="${h(c.id)}" data-w="${w}" aria-pressed="${erledigt.has(w)}" aria-label="Woche ${w}, ${w} Euro">${w}</button>`).join('')}</div>
      <p class="hint klein">Antippen hakt eine Woche ab (oder wieder ab).${c.topf ? ' Das Geld selbst überweist du wie gewohnt in den Topf.' : ''}</p>`;
  } else if (c.typ === 'limit') {
    mitte = `${balken}<div class="ch-zahlen"><span><b>${eur(s.ausgegeben)}</b> von ${eur(c.betrag)}</span><span>${s.rest >= 0 ? `noch ${eur(s.rest)}` : `${eur(-s.rest)} drüber`}</span></div>
      <p class="hint">Belohnung: <b>${h(c.belohnung || '–')}</b>${c.topf ? ` · Rest geht in „${h(c.topf)}“` : ''}</p>
      ${s.geschafft ? `<p class="feier-text">Geschafft! Gönn dir: ${h(c.belohnung || 'deine Belohnung')} 🎁</p>${c.topf && s.rest > 0 && !C.z.modell.buchungen.some((b) => b.posten === `Rest ${c.bereich} ${monatsname(c.monat)}`) ? `<button class="knopf klein" data-m="ch-rest" data-arg="${h(c.id)}">Rest ${eur(s.rest)} in den Topf eintragen</button>` : ''}`
        : s.vorbei ? '<p class="hint">Diesmal nicht geklappt – nächster Monat, neue Chance.</p>' : ''}`;
  } else if (c.typ === 'nospend') {
    const erster = new Date(c.monat + '-01T12:00:00').getDay();
    const versatz = (erster + 6) % 7;
    mitte = `${balken}<div class="ch-zahlen"><span><b>${s.geschafftTage}</b> von ${c.ziel} Tagen</span><span>${h(monatsname(c.monat, true))}</span></div>
      <div class="kalender"><span>Mo</span><span>Di</span><span>Mi</span><span>Do</span><span>Fr</span><span>Sa</span><span>So</span>${'<i></i>'.repeat(versatz)}${s.tage.map((t) => `<b class="${t.vergangen ? (t.ohne ? 'ohne' : 'mit') : t.heute ? 'heute' : 'spaeter'}" title="${datumLang(t.datum)}">${Number(t.datum.slice(8))}</b>`).join('')}</div>
      <p class="hint klein">Grün = Tag ohne eigene Ausgaben (feste Kosten und Ausgaben aus Töpfen zählen nicht). Gezählt wird ein Tag, wenn er vorbei ist.</p>`;
  }
  return `<div class="ch-karte ${s.geschafft ? 'geschafft' : ''}">${kopf}${mitte}
    <div class="ch-fuss"><button class="link" data-m="ch-bearbeiten" data-arg="${h(c.id)}">Bearbeiten</button>${s.vorbei ? `<button class="link" data-m="ch-archiv" data-arg="${h(c.id)}">Abschließen</button>` : ''}</div></div>`;
}

// ---------------------------------------------------------------- Dialoge
const dialog = (html) => { C.menu.innerHTML = html; if (!C.menu.open) C.menu.showModal(); };
const opt = (liste, wert, leer) => (leer != null ? `<option value="">${h(leer)}</option>` : '') + liste.map((x) => {
  const [v, t] = Array.isArray(x) ? x : [x, x];
  return `<option value="${h(v)}" ${v === wert ? 'selected' : ''}>${h(t)}</option>`;
}).join('');
const bereichListe = () => [...new Set([...d().limits.map((l) => l.bereich), ...d().buchungen.map((b) => b.bereich).filter(Boolean), ...P.BEREICHE_VORSCHLAG])];
const SR = window.SpeechRecognition || window.webkitSpeechRecognition || null;

function eintragenDialog(fehler, vorher) {
  const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
  const v = vorher || { art: 'Ausgabe', weg: 'Giro', datum: C.heute(), betrag: '', posten: '', bereich: '', plan: '', ausTopf: '', kommentar: '' };
  const bearbeiten = Boolean(v.nr);
  const plaene = [...mp.einnahmen.map((e) => ['E', e]), ...mp.feste.map((f) => ['A', f])]
    .filter(([, x]) => !x.ist || x.ist.nr === v.nr).map(([a, x]) => [x.id, `${a === 'E' ? 'Einnahme' : 'Feste Kosten'}: ${x.posten} (${eur(x.plan)})`]);
  const postenListe = [...new Set([...d().buchungen.map((b) => b.posten), ...mp.feste.map((f) => f.posten)].filter(Boolean))];
  dialog(`<form id="m-eintragen" novalidate>
    <h2 id="menu-titel">${bearbeiten ? `${h(v.nr)} ändern` : 'In meinen Bereich eintragen'}</h2>
    ${bearbeiten ? '' : `<div class="sprechen">
      ${SR ? `<button type="button" class="mikro" id="m-mikro" aria-label="Sprechen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg></button>` : ''}
      <input id="m-satz" autocomplete="off" enterkeyhint="done" placeholder="z. B. „19,99 Zalando mit PayPal“">
      <button type="button" class="knopf leise klein" id="m-satz-los">Übernehmen</button></div>
    <p class="erkannt" id="m-erkannt" aria-live="polite"></p>`}
    <div class="seg zwei-seg" role="radiogroup" aria-label="Art">${['Ausgabe', 'Einnahme'].map((a) => `<label><input type="radio" name="art" value="${a}" ${v.art === a ? 'checked' : ''}><span>${a}</span></label>`).join('')}</div>
    <div class="seg zwei-seg" role="radiogroup" aria-label="Bezahlt mit">${P.WEGE.map((w) => `<label><input type="radio" name="weg" value="${w}" ${v.weg === w ? 'checked' : ''}><span>${w === 'Giro' ? 'Giro' : 'PayPal'}</span></label>`).join('')}</div>
    <label class="feld"><span>Betrag in €</span><input name="betrag" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${h(v.betrag === '' ? '' : betragText(Number(v.betrag)) || v.betrag)}"></label>
    <label class="feld"><span>Wofür</span><input name="posten" list="m-posten" autocomplete="off" placeholder="z. B. Friseur" value="${h(v.posten)}"><datalist id="m-posten">${postenListe.map((p) => `<option value="${h(p)}">`).join('')}</datalist></label>
    <div class="zwei"><label class="feld"><span>Datum</span><input type="date" name="datum" value="${h(v.datum)}"></label>
      <label class="feld"><span>Bereich</span><input name="bereich" list="m-bereiche" autocomplete="off" placeholder="z. B. Kleidung" value="${h(v.bereich)}"><datalist id="m-bereiche">${bereichListe().map((b) => `<option value="${h(b)}">`).join('')}</datalist></label></div>
    <label class="feld"><span>Gehört zu <small>(geplanter Posten, dann zählt er nicht doppelt)</small></span><select name="plan">${opt(plaene, v.plan, '– nichts Geplantes –')}</select></label>
    <label class="feld" id="m-topf-feld" ${v.art === 'Einnahme' ? 'hidden' : ''}><span>Aus einem Topf bezahlt? <small>(zählt dann nicht gegen „frei“)</small></span><select name="ausTopf">${opt(P.meineToepfe(C.z.modell, d()), v.ausTopf, 'nein')}</select></label>
    <label class="feld"><span>Kommentar <small>(optional)</small></span><input name="kommentar" autocomplete="off" value="${h(v.kommentar || '')}"></label>
    ${v.importId ? `<input type="hidden" name="importId" value="${h(v.importId)}">` : ''}${bearbeiten ? `<input type="hidden" name="nr" value="${h(v.nr)}">` : `<p class="hint">Geld aufs Gemeinschaftskonto, aufs Tagesgeld oder in einen Topf? <button type="button" class="link" data-m="gemeinsam-statt">In die gemeinsame Datei eintragen</button></p>`}
    <p class="zusammen" id="m-zusammen"></p>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" type="submit">${bearbeiten ? 'Speichern' : 'Eintragen'}</button><button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div>
  </form>`);
  const form = document.getElementById('m-eintragen');
  const zusammen = () => {
    const x = Object.fromEntries(new FormData(form));
    const b = zahl(x.betrag);
    document.getElementById('m-zusammen').textContent = b > 0 ? `${x.art} über ${eur(b)} ${x.weg === 'PayPal' ? 'mit PayPal (geht vom Giro ab)' : 'vom Giro'}${x.posten ? ' für ' + x.posten : ''}.` : '';
  };
  form.addEventListener('input', zusammen);
  form.addEventListener('change', (e) => {
    if (e.target.name === 'art') document.getElementById('m-topf-feld').hidden = e.target.value === 'Einnahme';
    if (e.target.name === 'posten' && !form.bereich.value) { const b = P.bereichVorschlag(d(), e.target.value); if (b) form.bereich.value = b; }
    if (e.target.name === 'plan' && e.target.value) {
      const p = [...mp.einnahmen, ...mp.feste].find((x) => x.id === e.target.value);
      if (p) { if (!form.posten.value) form.posten.value = p.posten; if (!form.betrag.value) form.betrag.value = betragText(p.plan); form.querySelector(`input[name=art][value="${mp.einnahmen.includes(p) ? 'Einnahme' : 'Ausgabe'}"]`).checked = true; }
    }
    zusammen();
  });
  zusammen();
  const satz = document.getElementById('m-satz');
  if (satz) {
    const hinweis = document.getElementById('m-erkannt');
    const los = () => {
      const { felder: f, erkannt, fehlt } = P.verstehenPrivat(satz.value, d(), C.heute());
      if (f.art) form.querySelector(`input[name=art][value="${f.art}"]`).checked = true;
      if (f.weg) form.querySelector(`input[name=weg][value="${f.weg}"]`).checked = true;
      if (f.betrag) form.betrag.value = betragText(f.betrag);
      if (f.posten) form.posten.value = f.posten;
      if (f.datum) form.datum.value = f.datum;
      if (f.bereich) form.bereich.value = f.bereich;
      document.getElementById('m-topf-feld').hidden = f.art === 'Einnahme';
      hinweis.innerHTML = `Vorausgefüllt${erkannt.length ? ': ' + h(erkannt.join(', ')) : ''}. Bitte kurz prüfen.${fehlt.length ? ` <b>Noch offen: ${h(fehlt.join(', '))}.</b>` : ''}`;
      zusammen();
    };
    document.getElementById('m-satz-los').addEventListener('click', los);
    satz.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); los(); satz.blur(); } });
    const mikro = document.getElementById('m-mikro');
    if (mikro) {
      let er = null;
      mikro.addEventListener('click', () => {
        if (er) { er.stop(); return; }
        er = new SR(); er.lang = 'de-DE'; er.interimResults = true; er.continuous = false;
        mikro.classList.add('hoert'); hinweis.textContent = 'Ich höre zu …';
        er.onresult = (e) => { satz.value = [...e.results].map((r) => r[0].transcript).join(' '); if (e.results[e.results.length - 1].isFinal) los(); };
        er.onerror = (e) => { hinweis.textContent = e.error === 'not-allowed' ? 'Das Mikrofon ist gesperrt. Bitte in den Browser-Einstellungen erlauben oder tippen.' : 'Spracherkennung gerade nicht möglich. Bitte tippen.'; };
        er.onend = () => { mikro.classList.remove('hoert'); er = null; };
        try { er.start(); } catch { mikro.classList.remove('hoert'); er = null; }
      });
    }
  }
}

function einfacherDialog(id, titel, felder, fehler, { knopfText = 'Speichern', loeschen = null, text = '' } = {}) {
  dialog(`<form id="${id}" novalidate><h2 id="menu-titel">${titel}</h2>${text}${felder}
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" type="submit">${knopfText}</button>${loeschen ? `<button class="knopf leise" type="button" data-m="${loeschen[0]}" data-arg="${h(loeschen[1])}">${loeschen[2] || 'Löschen'}</button>` : ''}<button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div></form>`);
}
const feld = (name, label, wert = '', extra = '') => `<label class="feld"><span>${label}</span><input name="${name}" value="${h(wert ?? '')}" autocomplete="off" ${extra}></label>`;

function giroStartDialog(fehler) {
  einfacherDialog('m-giro', 'Anfangsstand Giro', feld('anfang', 'Stand laut Bank in €', betragText(d().giro.anfang), 'inputmode="decimal"')
    + feld('stichtag', 'Am Ende von diesem Tag', d().giro.stichtag || C.heute(), 'type="date"'), fehler,
  { text: '<p class="hint">Alles bis einschließlich zu diesem Tag steckt im Stand. Danach rechnet die App selbst weiter.</p>' });
}

function limitsDialog(fehler) {
  const zeilen = [...d().limits, { bereich: '', betrag: '' }];
  dialog(`<form id="m-limits" novalidate><h2 id="menu-titel">Deine Limits</h2>
    <p class="hint">Ein Limit pro Bereich und Monat. Bei 80 % und beim Erreichen gibt die App einen Hinweis. Leeren Betrag = Limit weg.</p>
    ${zeilen.map((l, i) => `<div class="zwei"><label class="feld"><span>Bereich</span><input name="b${i}" list="m-bereiche" value="${h(l.bereich)}" autocomplete="off"></label>
      <label class="feld"><span>Limit im Monat €</span><input name="l${i}" inputmode="decimal" value="${h(betragText(l.betrag === '' ? null : l.betrag))}" autocomplete="off"></label></div>`).join('')}
    <datalist id="m-bereiche">${bereichListe().map((b) => `<option value="${h(b)}">`).join('')}</datalist>
    <input type="hidden" name="n" value="${zeilen.length}">
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <div class="knoepfe"><button class="knopf" type="submit">Speichern</button><button class="knopf leise" type="button" data-m="limit-zeile">Weiteres Limit</button><button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div></form>`);
}

function erinnerungenDialog(fehler) {
  const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
  const alle = P.erinnerungen(d(), mp, C.heute());
  dialog(`<h2 id="menu-titel">Erinnerungen</h2>
    <p class="hint">Fällige Erinnerungen zeigt die App beim Öffnen ganz oben. Damit dein Handy von selbst erinnert, „In den Kalender“ tippen: Das legt einen monatlichen Termin an.</p>
    <div class="feste">${alle.map((e) => `<div class="f-zeile"><div><div class="t">${h(e.text)}</div><div class="hint">${e.tag ? `jeden Monat am ${e.tag}.` : 'ohne Tag'}${e.auto ? ' · automatisch aus der gemeinsamen Datei' : ''}${e.erledigt ? ' · diesen Monat erledigt' : ''}</div></div>
      <button class="knopf leise klein" data-m="ics" data-arg="${h(e.id)}">In den Kalender</button>${e.auto ? '' : `<button class="knopf leise klein" data-m="erin-loeschen" data-arg="${h(e.id)}">Löschen</button>`}</div>`).join('')}</div>
    <form id="m-erin" novalidate style="margin-top:14px"><h3>Neue Erinnerung</h3>
      ${feld('text', 'Woran?', '', 'placeholder="z. B. Nagelstudio-Termin machen"')}
      ${feld('tag', 'Jeden Monat am Tag', '', 'inputmode="numeric" placeholder="1–31"')}
      ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
      <div class="knoepfe"><button class="knopf" type="submit">Hinzufügen</button><button class="knopf leise" type="button" id="schliessen">Fertig</button></div></form>`);
}

function festDialog(id, fehler) {
  const f = id ? d().feste.find((x) => x.id === id) : { art: 'Ausgabe', rhythmus: 'monatlich', posten: '', betrag: '', tag: '', bereich: '', startMonat: C.z.monat, ab: '', bis: '' };
  einfacherDialog('m-fest', id ? 'Festen Posten ändern' : 'Eigener fester Posten',
    `<div class="seg zwei-seg">${['Ausgabe', 'Einnahme'].map((a) => `<label><input type="radio" name="art" value="${a}" ${f.art === a ? 'checked' : ''}><span>${a}</span></label>`).join('')}</div>
    ${feld('posten', 'Wofür', f.posten, 'placeholder="z. B. Handyvertrag"')}
    <div class="zwei">${feld('betrag', 'Betrag €', betragText(f.betrag === '' ? null : f.betrag), 'inputmode="decimal"')}${feld('tag', 'Tag im Monat', f.tag || '', 'inputmode="numeric"')}</div>
    <div class="zwei"><label class="feld"><span>Rhythmus</span><select name="rhythmus">${opt(['monatlich', 'vierteljährlich', 'halbjährlich', 'jährlich'], f.rhythmus)}</select></label>
      ${feld('startMonat', 'Zuerst fällig im', f.startMonat || C.z.monat, 'type="month"')}</div>
    <div class="zwei">${feld('ab', 'Ab (optional)', f.ab || '', 'type="month"')}${feld('bis', 'Bis (optional)', f.bis || '', 'type="month"')}</div>
    <label class="feld"><span>Bereich</span><input name="bereich" list="m-bereiche" value="${h(f.bereich || '')}" autocomplete="off"><datalist id="m-bereiche">${bereichListe().map((b) => `<option value="${h(b)}">`).join('')}</datalist></label>
    ${id ? `<input type="hidden" name="id" value="${h(id)}">` : ''}`, fehler,
  { loeschen: id ? ['fest-loeschen', id] : null, text: '<p class="hint">Nur Posten, die von deinem Giro gehen und nicht schon in der gemeinsamen Datei stehen (z. B. Handy, Ratenzahlung).</p>' });
}

function topfDialog(name, fehler) {
  const t = C.z.modell.toepfe.find((x) => x.name === name);
  const ziel = d().ziele[name] || {};
  einfacherDialog('m-topf', `Topf „${h(name)}“`,
    `<input type="hidden" name="alt" value="${h(name)}">
    ${feld('name', 'Name', t.name)}${feld('sparrate', 'Sparrate im Monat €', betragText(t.sparrate), 'inputmode="decimal"')}${feld('bemerkung', 'Bemerkung', t.bemerkung || '')}
    <h3>Dein Ziel <small class="hint">(nur für dich)</small></h3>
    <div class="zwei">${feld('zielBetrag', 'Zielbetrag €', betragText(ziel.betrag), 'inputmode="decimal"')}${feld('zielDatum', 'bis (optional)', ziel.datum || '', 'type="date"')}</div>`, fehler,
  { text: '<p class="hint warnfarbe">Name, Sparrate und Bemerkung stehen in der gemeinsamen Datei – Kai sieht die Änderung. Eine andere Sparrate ändert auch, was du jeden Monat aufs Tagesgeld überweist.</p>' });
}
function topfNeuDialog(fehler) {
  einfacherDialog('m-topf-neu', 'Neuer Topf', `${feld('name', 'Name', '', 'placeholder="z. B. Urlaub Natalie"')}${feld('sparrate', 'Sparrate im Monat €', '', 'inputmode="decimal" placeholder="0,00"')}
    <label class="feld"><span>Bereich</span><select name="bereich">${opt(C.z.modell.listen?.bereiche || [], 'Persönliches')}</select></label>${feld('bemerkung', 'Bemerkung (optional)')}
    <div class="zwei">${feld('zielBetrag', 'Ziel € (optional, nur für dich)', '', 'inputmode="decimal"')}${feld('zielDatum', 'bis (optional)', '', 'type="date"')}</div>`, fehler,
  { knopfText: 'Topf anlegen', text: '<p class="hint warnfarbe">Der Topf wird in der gemeinsamen Datei angelegt (Blatt „Töpfe“, auf dem Tagesgeld). Kai sieht ihn. Eine Sparrate erhöht deine monatliche Überweisung aufs Tagesgeld.</p>' });
}
function meineToepfeDialog() {
  const meine = P.meineToepfe(C.z.modell, d());
  dialog(`<form id="m-meine" novalidate><h2 id="menu-titel">Welche Töpfe sind meine?</h2><p class="hint">Diese Töpfe zeigt dein Bereich. Die Auswahl ist nur für dich.</p>
    <div class="haken-liste">${C.z.modell.toepfe.map((t) => `<label class="haken"><input type="checkbox" name="t" value="${h(t.name)}" ${meine.includes(t.name) ? 'checked' : ''}><span><b>${h(t.name)}</b><small>${h(t.bereich || '')}</small></span></label>`).join('')}</div>
    <div class="knoepfe"><button class="knopf" type="submit">Speichern</button><button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div></form>`);
}

function challengeNeuDialog(typ, fehler, vorher = {}) {
  if (!typ) {
    dialog(`<h2 id="menu-titel">Neue Challenge</h2><div class="knoepfe">${Object.entries(P.CHALLENGE_ARTEN).map(([k, a]) => `<button class="knopf grund-knopf" data-m="ch-typ" data-arg="${k}"><b>${h(a.titel)}</b><small>${h(a.kurz)}</small></button>`).join('')}<button class="knopf leise" id="schliessen">Abbrechen</button></div>`);
    return;
  }
  const v = vorher;
  const topfWahl = `<label class="feld"><span>Topf <small>(optional)</small></span><select name="topf">${opt(P.meineToepfe(C.z.modell, d()), v.topf || '', 'ohne Topf')}</select></label>`;
  const bearbeiten = Boolean(v.id);
  let felder = '';
  if (typ === 'ziel') felder = `${feld('name', 'Wofür sparst du?', v.name || '', 'placeholder="z. B. neue Jacke"')}<div class="zwei">${feld('betrag', 'Ziel €', betragText(v.betrag), 'inputmode="decimal"')}${feld('datum', 'Wunschdatum', v.datum || '', 'type="date"')}</div>${topfWahl}<p class="hint klein">Mit Topf zählt der Stand des Topfs. Ohne Topf trägst du selbst ein, was du zurückgelegt hast.</p>`;
  if (typ === 'wochen52') felder = `${feld('start', 'Start (Woche 1)', v.start || C.heute(), 'type="date"')}${topfWahl}`;
  if (typ === 'limit') felder = `<label class="feld"><span>Bereich</span><input name="bereich" list="m-bereiche" value="${h(v.bereich || '')}" autocomplete="off"><datalist id="m-bereiche">${bereichListe().map((b) => `<option value="${h(b)}">`).join('')}</datalist></label>
    <div class="zwei">${feld('betrag', 'Höchstens €', betragText(v.betrag), 'inputmode="decimal"')}${feld('monat', 'Monat', v.monat || C.z.monat, 'type="month"')}</div>
    ${feld('belohnung', 'Deine Belohnung', v.belohnung || '', 'placeholder="z. B. Kino mit einer Freundin"')}${topfWahl.replace('Topf <small>(optional)</small>', 'Rest in den Topf <small>(optional)</small>')}`;
  if (typ === 'nospend') felder = `<div class="zwei">${feld('ziel', 'Wie viele Tage?', v.ziel || '', 'inputmode="numeric"')}${feld('monat', 'Monat', v.monat || C.z.monat, 'type="month"')}</div>`;
  einfacherDialog('m-challenge', `${bearbeiten ? 'Ändern: ' : ''}${h(P.CHALLENGE_ARTEN[typ].titel)}`, `<input type="hidden" name="typ" value="${typ}">${bearbeiten ? `<input type="hidden" name="id" value="${h(v.id)}">` : ''}${felder}`, fehler,
    { knopfText: bearbeiten ? 'Speichern' : 'Los geht’s', loeschen: bearbeiten ? ['ch-loeschen', v.id, 'Challenge löschen'] : null, text: `<p class="hint">${h(P.CHALLENGE_ARTEN[typ].kurz)}</p>` });
}

function menuDialog() {
  dialog(`<h2 id="menu-titel">Mein Bereich</h2>
    <div class="zeile"><span class="hint">Person</span><span>${h(M.person || '–')}</span></div>
    <div class="zeile"><span class="hint">Gespeichert in</span><span>${M.ordner ? (M.ordner.quelle === 'ersatz' ? 'deinem OneDrive › Unser-Geld-Privat' : `OneDrive › Apps › ${h(M.ordner.name)}`) : '–'}</span></div>
    <div class="zeile"><span class="hint">Zuletzt geändert</span><span>${d()?.geaendert ? new Date(d().geaendert).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' }) : '–'}</span></div>
    <div class="knoepfe">
      ${d() ? `<button class="knopf leise" data-m="giro-start">Anfangsstand Giro</button>
      <button class="knopf leise" data-m="limits">Limits</button>
      <button class="knopf leise" data-m="erinnerungen">Erinnerungen</button>
      <button class="knopf leise" data-m="fest-neu">Eigenen festen Posten anlegen</button>
      <button class="knopf leise" data-m="meine-toepfe">Welche Töpfe sind meine?</button>` : ''}
      <button class="knopf leise" data-m="neu-laden">Neu laden</button>
      <button class="knopf leise" data-m="gemeinsames-menu">Datei & Anmeldung (gemeinsam)</button>
      <button class="knopf leise" data-m="person-wechseln">Andere Person wählen</button>
      <button class="knopf leise" id="schliessen">Schließen</button></div>
    <p class="hint" style="margin-top:14px;text-align:center">Version ${h(C.version)}</p>`);
}

// ---------------------------------------------------------------- Ereignisse
/** Klicks im privaten Bereich. Gibt true zurück, wenn behandelt. */
export async function klick(t) {
  const a = t.dataset.m;
  if (t.id === 'menu-knopf') { menuDialog(); return true; }
  if (!a) return false;
  const arg = t.dataset.arg;
  const ctx = () => ({ modell: C.z.modell, p: d(), heute: C.heute() });
  switch (a) {
    case 'tab': M.tab = arg; merk('ug.mtab', arg); M.meldung = null; C.zeichne(); window.scrollTo(0, 0); return true;
    case 'person': merk('ug.person.' + konto(), arg); M.person = arg; M.status = 'aus'; M.akt = null; laden({ immer: true }); return true;
    case 'person-wechseln': if (C.menu.open) C.menu.close(); merk('ug.person.' + konto(), null); vergessen(); M.status = 'person'; M.person = null; thema(false); C.zeichne(); return true;
    case 'neu-laden': if (C.menu.open) C.menu.close(); if (M.person) C.od.privatOrdnerVergessen(M.person); M.meldung = null; M.status = d() ? 'bereit' : 'aus'; await laden({ immer: true }); return true;
    case 'ersatz-ordner': try { M.ordner = await C.od.privatOrdnerFinden(M.person, { ersatz: true }); M.status = 'aus'; laden({ immer: true }); } catch (e) { M.fehler = fehlerText(e); M.status = 'fehler'; C.zeichne(); } return true;
    case 'gemeinsames-menu': C.menu.close(); C.zeigeMenu(); return true;
    case 'eintragen': eintragenDialog(); return true;
    case 'gemeinsam-statt': C.menu.close(); C.zeigeEintragen(null, gemeinsamVorlage({})); return true;
    case 'kfilter': M.kontoFilter = arg; C.zeichne(); return true;
    case 'csv-zu': M.csv = null; M.csvFehler = null; C.zeichne(); return true;
    case 'bearbeiten': { const b = d().buchungen.find((x) => x.nr === arg); if (b) eintragenDialog(null, { ...b }); return true; }
    case 'loeschen-frage': {
      const b = d().buchungen.find((x) => x.nr === arg);
      dialog(`<h2 id="menu-titel">${h(b.nr)} löschen?</h2><p><b>${h(b.posten)}</b>, ${eur(b.betrag)} am ${datumLang(b.datum)}</p>
        <div class="knoepfe"><button class="knopf" data-m="loeschen" data-arg="${h(b.nr)}">Ja, löschen</button><button class="knopf leise" id="schliessen">Abbrechen</button></div>`);
      return true;
    }
    case 'loeschen': await speichern((x) => { x.buchungen = x.buchungen.filter((b) => b.nr !== arg); }, { art: 'info', text: `${arg} gelöscht.` }, { knopf: t }); return true;
    case 'plan-ist': {
      const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
      const e = mp.einnahmen.find((x) => x.id === arg); const f = mp.feste.find((x) => x.id === arg);
      const p = e || f;
      const datum = C.heute().slice(0, 7) === C.z.monat ? C.heute() : `${C.z.monat}-${String(Math.min(p.tag || 1, 28)).padStart(2, '0')}`;
      eintragenDialog(null, { art: e ? 'Einnahme' : 'Ausgabe', weg: 'Giro', datum, betrag: p.plan, posten: p.posten, bereich: p.bereich || '', plan: p.id, ausTopf: '', kommentar: '', monat: C.z.monat });
      return true;
    }
    case 'ueberweisung': {
      const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
      const u = mp.ueberweisungen.find((x) => x.nach === arg);
      C.zeigeEintragen(null, gemeinsamVorlage({ nach: arg, betrag: u?.offen, posten: `Überweisung ${monatsname(C.z.monat)}`, giltFuer: C.z.monat }));
      return true;
    }
    case 'erin-erledigt': await speichern((x) => { const e = x.erinnerungen.find((y) => y.id === arg); if (e) { e.erledigt = e.erledigt || {}; e.erledigt[C.z.monat] = true; } }, { art: 'info', text: 'Erledigt für diesen Monat.' }, { knopf: t }); return true;
    case 'erinnerungen': erinnerungenDialog(); return true;
    case 'erin-loeschen': await speichern((x) => { x.erinnerungen = x.erinnerungen.filter((e) => e.id !== arg); }, null, { dialogFehler: (f) => erinnerungenDialog(f), knopf: t }); erinnerungenDialog(); return true;
    case 'ics': {
      const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
      const e = P.erinnerungen(d(), mp, C.heute()).find((x) => x.id === arg);
      if (!e?.tag) return true;
      const text = P.ics({ titel: e.auto ? `Überweisung an ${e.nach}` : e.text, tag: e.tag, start: C.heute(), uid: (e.id + '-' + (M.person || '')).replace(/\W+/g, '-'), text: e.auto ? 'Betrag steht in „Unser Geld“ unter Meins.' : '' });
      const url = URL.createObjectURL(new Blob([text], { type: 'text/calendar' }));
      const link = Object.assign(document.createElement('a'), { href: url, download: 'Erinnerung-' + (e.auto ? 'Ueberweisung' : e.text).replace(/[^\wäöüÄÖÜß-]+/g, '-').slice(0, 40) + '.ics' });
      document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 5000);
      return true;
    }
    case 'limits': limitsDialog(); return true;
    case 'limit-zeile': {
      const form = document.getElementById('m-limits');
      const x = Object.fromEntries(new FormData(form));
      const n = Number(x.n);
      const tmp = Array.from({ length: n }, (_, i) => ({ bereich: x['b' + i], betrag: x['l' + i] ? zahl(x['l' + i]) : '' }));
      const alt = d().limits; d().limits = tmp.filter((l, i) => l.bereich || i < n - 1); limitsDialog(); d().limits = alt;
      return true;
    }
    case 'giro-start': giroStartDialog(); return true;
    case 'fest-neu': festDialog(null); return true;
    case 'fest-bearbeiten': festDialog(arg); return true;
    case 'fest-loeschen': await speichern((x) => { x.feste = x.feste.filter((f) => f.id !== arg); }, { art: 'info', text: 'Fester Posten gelöscht.' }, { knopf: t }); return true;
    case 'topf-bearbeiten': topfDialog(arg); return true;
    case 'topf-neu': topfNeuDialog(); return true;
    case 'meine-toepfe': meineToepfeDialog(); return true;
    case 'challenge-neu': challengeNeuDialog(); return true;
    case 'ch-typ': challengeNeuDialog(arg); return true;
    case 'ch-bearbeiten': { const c = d().challenges.find((x) => x.id === arg); challengeNeuDialog(c.typ, null, c); return true; }
    case 'ch-loeschen': await speichern((x) => { x.challenges = x.challenges.filter((c) => c.id !== arg); }, { art: 'info', text: 'Challenge gelöscht. Pokale in der Ruhmeshalle bleiben.' }, { knopf: t }); return true;
    case 'ch-archiv': await speichern((x) => { const c = x.challenges.find((y) => y.id === arg); if (c) c.archiviert = C.heute(); }, { art: 'info', text: 'Challenge abgeschlossen.' }, { knopf: t }); return true;
    case 'ch-zurueck': await speichern((x) => { const c = x.challenges.find((y) => y.id === arg); if (c) delete c.archiviert; }, null, { knopf: t }); return true;
    case 'woche': {
      const w = Number(t.dataset.w);
      t.classList.toggle('an');
      await speichern((x) => { const c = x.challenges.find((y) => y.id === arg); const s = new Set(c.erledigt || []); s.has(w) ? s.delete(w) : s.add(w); c.erledigt = [...s].sort((p, q) => p - q); }, null);
      return true;
    }
    case 'ch-einzahlung': einfacherDialog('m-einzahlung', 'Zurückgelegt', `<input type="hidden" name="id" value="${h(arg)}">${feld('betrag', 'Betrag €', '', 'inputmode="decimal"')}${feld('datum', 'Datum', C.heute(), 'type="date"')}`, null); return true;
    case 'ch-rest': {
      const c = d().challenges.find((x) => x.id === arg);
      const s = P.challengeStand(c, ctx());
      C.zeigeEintragen(null, gemeinsamVorlage({ nach: sparkonto(), betrag: s.rest, posten: `Rest ${c.bereich} ${monatsname(c.monat)}`, topf: c.topf, giltFuer: C.heute().slice(0, 7), kommentar: 'Challenge: Monatslimit mit Belohnung' }));
      return true;
    }
    case 'csv-uebernehmen': {
      const r = M.csv?.zeilen.find((x) => x.id === arg);
      if (!r) return true;
      const posten = kurzname(r.z.info, r.z.umsatzart);
      eintragenDialog(null, { art: r.z.betrag > 0 ? 'Einnahme' : 'Ausgabe', weg: /paypal/i.test(r.z.info) ? 'PayPal' : 'Giro', datum: r.z.datum, betrag: Math.abs(r.z.betrag), posten, bereich: P.bereichVorschlag(d(), posten) || '', plan: '', ausTopf: '', kommentar: '', importId: r.id, quelle: 'CSV' });
      return true;
    }
    case 'csv-gemeinsam': {
      const r = M.csv?.zeilen.find((x) => x.id === arg);
      if (!r) return true;
      const ein = r.z.betrag > 0;
      C.zeigeEintragen(null, gemeinsamVorlage({ von: ein ? r.gemeinsamKonto : undefined, nach: ein ? d().giro.konto : r.gemeinsamKonto, betrag: Math.abs(r.z.betrag), posten: kurzname(r.z.info, r.z.umsatzart), datum: r.z.datum, giltFuer: r.z.datum.slice(0, 7), beschreibung: r.z.info.slice(0, 120) }));
      return true;
    }
    default: return false;
  }
}
const sparkonto = () => C.z.modell.konten.find((k) => k.rolle === 'Sparen' && k.status !== 'geschlossen')?.name || '';
/** Vorlage fürs Formular der gemeinsamen Datei: Umbuchung vom eigenen Giro. */
function gemeinsamVorlage({ von, nach, betrag, posten, topf, giltFuer, kommentar, datum, beschreibung }) {
  const m = C.z.modell;
  const bereich = (m.listen?.bereiche || []).find((b) => /sparen|umbuch/i.test(b)) || '';
  return { art: 'Umbuchung', datum: datum || C.heute(), giltFuer: giltFuer || C.heute().slice(0, 7), von: von || d().giro.konto, nach: nach || m.konten.find((k) => k.rolle === 'gemeinsam')?.name,
    betrag: betrag ? betragText(betrag) : '', posten: posten || '', bereich, topf: topf || '', beschreibung: beschreibung || '', kommentar: kommentar || '', gezahltVon: '' };
}

/** Formulare im privaten Bereich. Gibt true zurück, wenn behandelt. */
export async function absenden(form) {
  const x = Object.fromEntries(new FormData(form));
  const knopf = form.querySelector('button[type=submit]');
  switch (form.id) {
    case 'm-einrichten': {
      const anfang = x.anfang.trim() ? zahl(x.anfang) : null;
      if (x.anfang.trim() && !isFinite(anfang)) { C.app.querySelector('main').innerHTML = einrichten('Bitte den Stand als Zahl eingeben, z. B. 412,37.'); return true; }
      const daten = P.leer(M.person, P.privatesKonto(C.z.modell, M.person), C.heute());
      daten.giro.anfang = anfang; daten.giro.stichtag = anfang == null ? null : (x.stichtag || C.heute());
      knopf.disabled = true; knopf.textContent = 'Wird angelegt …';
      try { await C.od.privatAnlegen(M.ordner, daten); M.status = 'aus'; M.meldung = { art: 'feier', text: 'Dein Bereich ist angelegt. Willkommen! 💐' }; await laden({ immer: true }); setTimeout(konfetti, 200); }
      catch (e) { if (e instanceof C.od.Konflikt) { M.status = 'aus'; await laden({ immer: true }); } else C.app.querySelector('main').innerHTML = einrichten(fehlerText(e)); }
      return true;
    }
    case 'm-eintragen': {
      const betrag = zahl(x.betrag);
      const f = !(betrag > 0) ? 'Bitte einen Betrag größer als 0 eingeben.' : !x.posten.trim() ? 'Bitte eintragen, wofür.' : !x.datum ? 'Bitte das Datum angeben.' : null;
      const vorher = { ...x, betrag: x.betrag, nr: x.nr };
      if (f) { eintragenDialog(f, vorher); return true; }
      const mp = P.monatPrivat(C.z.modell, d(), C.z.monat, C.heute());
      const plan = x.plan ? [...mp.einnahmen, ...mp.feste].find((p) => p.id === x.plan) : null;
      const alt = x.nr ? d().buchungen.find((b) => b.nr === x.nr) : null;
      const eintrag = { datum: x.datum, monat: plan ? C.z.monat : (alt?.plan && !x.plan ? x.datum.slice(0, 7) : alt?.monat && x.plan === alt.plan ? alt.monat : x.datum.slice(0, 7)), art: x.art, weg: x.weg, betrag,
        posten: x.posten.trim(), bereich: x.bereich.trim() || null, plan: x.plan || null, ausTopf: x.art === 'Ausgabe' && x.ausTopf ? x.ausTopf : null, kommentar: x.kommentar.trim() || null };
      const csvId = !x.nr && x.importId ? x.importId : null;
      const hinweis = x.art === 'Ausgabe' && !x.nr && eintrag.bereich ? P.limitHinweis(d(), eintrag.monat, eintrag.bereich, betrag) : null;
      const ok = await speichern((dd) => {
        if (x.nr) { const b = dd.buchungen.find((y) => y.nr === x.nr); if (!b) throw new Error(`${x.nr} gibt es nicht mehr.`); Object.assign(b, eintrag, { geaendert: C.heute() }); return x.nr; }
        if (csvId && dd.buchungen.some((y) => y.importId === csvId)) throw new Error('Diese Bankzeile wurde schon übernommen.');
        const nr = P.neueNr(dd);
        dd.buchungen.push({ nr, ...eintrag, quelle: csvId ? 'CSV' : 'manuell', importId: csvId, angelegt: C.heute() });
        return nr;
      }, (nr) => ({ art: hinweis ? 'warn' : 'info', text: `${nr} ${x.nr ? 'geändert' : 'eingetragen'}: ${eintrag.posten}, ${eur(betrag)}.`
        + (hinweis ? (hinweis.stufe === 'erreicht' ? ` Dein Limit für ${hinweis.bereich} (${eur(hinweis.limit)}) ist jetzt erreicht.` : ` Du hast 80 % deines Limits für ${hinweis.bereich} erreicht (${eur(hinweis.summe)} von ${eur(hinweis.limit)}).`) : '') }),
      { knopf, dialogFehler: (fe) => eintragenDialog(fe, vorher) });
      if (ok && M.csvDatei) { abgleichNeu(); C.zeichne(); }
      return true;
    }
    case 'm-giro': {
      const anfang = zahl(x.anfang);
      if (!isFinite(anfang) || !x.stichtag) { giroStartDialog('Bitte Stand und Datum angeben.'); return true; }
      await speichern((dd) => { dd.giro.anfang = anfang; dd.giro.stichtag = x.stichtag; }, { art: 'info', text: 'Anfangsstand gespeichert.' }, { knopf, dialogFehler: giroStartDialog });
      return true;
    }
    case 'm-limits': {
      const n = Number(x.n); const limits = [];
      for (let i = 0; i < n; i++) {
        const b = (x['b' + i] || '').trim(); const l = zahl(x['l' + i]);
        if (b && l > 0) { if (limits.some((y) => y.bereich === b)) { limitsDialog(`„${b}“ steht doppelt.`); return true; } limits.push({ bereich: b, betrag: l }); }
      }
      await speichern((dd) => { dd.limits = limits; }, { art: 'info', text: 'Limits gespeichert.' }, { knopf, dialogFehler: limitsDialog });
      return true;
    }
    case 'm-erin': {
      const tag = Number(x.tag);
      if (!x.text.trim() || !(tag >= 1 && tag <= 31)) { erinnerungenDialog('Bitte Text und einen Tag von 1 bis 31 angeben.'); return true; }
      await speichern((dd) => { dd.erinnerungen.push({ id: P.neueNr(dd, 'R'), text: x.text.trim(), tag, aktiv: true, erledigt: {} }); }, null, { knopf, dialogFehler: erinnerungenDialog });
      erinnerungenDialog();
      return true;
    }
    case 'm-fest': {
      const betrag = zahl(x.betrag); const tag = x.tag ? Number(x.tag) : null;
      if (!x.posten.trim() || !(betrag > 0) || (tag != null && !(tag >= 1 && tag <= 31))) { festDialog(x.id, 'Bitte Wofür, Betrag und (optional) einen Tag von 1 bis 31 angeben.'); return true; }
      const f = { art: x.art, posten: x.posten.trim(), betrag, tag, rhythmus: x.rhythmus, startMonat: x.startMonat || null, ab: x.ab || null, bis: x.bis || null, bereich: x.bereich.trim() || null };
      await speichern((dd) => { if (x.id) Object.assign(dd.feste.find((y) => y.id === x.id), f); else dd.feste.push({ id: P.neueNr(dd, 'F'), ...f }); },
        { art: 'info', text: `Fester Posten „${f.posten}“ gespeichert.` }, { knopf, dialogFehler: (fe) => festDialog(x.id, fe) });
      return true;
    }
    case 'm-meine': {
      const namen = new FormData(form).getAll('t');
      await speichern((dd) => { dd.meineToepfe = namen; }, null, { knopf });
      return true;
    }
    case 'm-topf': {
      const t = C.z.modell.toepfe.find((y) => y.name === x.alt);
      const sparrate = zahl(x.sparrate); const zb = x.zielBetrag ? zahl(x.zielBetrag) : null;
      if (!x.name.trim() || !(sparrate >= 0) || (x.zielBetrag && !(zb > 0))) { topfDialog(x.alt, 'Bitte Name und Sparrate (0 oder mehr) angeben, Ziel als Zahl.'); return true; }
      knopf.disabled = true; knopf.textContent = 'Wird gespeichert …';
      try {
        const neuName = x.name.trim();
        const geaendert = neuName !== t.name || Math.abs(sparrate - t.sparrate) > 0.004 || (x.bemerkung || '') !== (t.bemerkung || '');
        let umgestellt = 0;
        if (geaendert) ({ umgestellt } = await C.od.topfAendern(C.od.aktuellerOrt(), t.name, { name: neuName, sparrate, bemerkung: x.bemerkung }));
        await aendern((dd) => {
          const z0 = dd.ziele[t.name]; delete dd.ziele[t.name];
          if (zb) dd.ziele[neuName] = { betrag: zb, datum: x.zielDatum || null }; else if (z0 && !x.zielBetrag) { /* Ziel entfernt */ }
          if (Array.isArray(dd.meineToepfe) && neuName !== t.name) dd.meineToepfe = dd.meineToepfe.map((n) => (n === t.name ? neuName : n));
          for (const c of dd.challenges) if (c.topf === t.name) c.topf = neuName;
        });
        C.menu.close();
        M.meldung = { art: 'info', text: `Topf „${neuName}“ gespeichert${geaendert ? ' – auch in der gemeinsamen Datei' : ''}${umgestellt ? ` (${umgestellt} Buchungen/Kosten auf den neuen Namen umgestellt)` : ''}.` };
        if (geaendert) await C.ausOneDrive({ immer: true }); else C.zeichne();
      } catch (e) { topfDialog(x.alt, fehlerText(e)); }
      return true;
    }
    case 'm-topf-neu': {
      const sparrate = x.sparrate ? zahl(x.sparrate) : 0; const zb = x.zielBetrag ? zahl(x.zielBetrag) : null;
      if (!x.name.trim() || !(sparrate >= 0)) { topfNeuDialog('Bitte einen Namen und eine Sparrate (0 oder mehr) angeben.'); return true; }
      knopf.disabled = true; knopf.textContent = 'Wird angelegt …';
      try {
        await C.od.topfNeu(C.od.aktuellerOrt(), { name: x.name.trim(), konto: sparkonto(), bereich: x.bereich, sparrate, bemerkung: x.bemerkung, stichtag: C.heute() });
        await aendern((dd) => { const meine = P.meineToepfe(C.z.modell, dd); dd.meineToepfe = [...new Set([...meine, x.name.trim()])]; if (zb) dd.ziele[x.name.trim()] = { betrag: zb, datum: x.zielDatum || null }; });
        C.menu.close(); M.meldung = { art: 'info', text: `Topf „${x.name.trim()}“ angelegt.` };
        await C.ausOneDrive({ immer: true });
      } catch (e) { topfNeuDialog(fehlerText(e)); }
      return true;
    }
    case 'm-challenge': {
      const c = { typ: x.typ };
      let f = null;
      if (x.typ === 'ziel') { c.name = x.name.trim(); c.betrag = zahl(x.betrag); c.datum = x.datum || null; c.topf = x.topf || null; if (!c.name || !(c.betrag > 0)) f = 'Bitte Namen und Zielbetrag angeben.'; }
      if (x.typ === 'wochen52') { c.start = x.start || C.heute(); c.topf = x.topf || null; }
      if (x.typ === 'limit') { c.bereich = x.bereich.trim(); c.betrag = zahl(x.betrag); c.monat = x.monat; c.belohnung = x.belohnung.trim(); c.topf = x.topf || null; if (!c.bereich || !(c.betrag > 0) || !c.monat) f = 'Bitte Bereich, Betrag und Monat angeben.'; }
      if (x.typ === 'nospend') { c.ziel = Number(x.ziel); c.monat = x.monat; if (!(c.ziel >= 1 && c.ziel <= 31) || !c.monat) f = 'Bitte eine Zahl von 1 bis 31 und den Monat angeben.'; }
      if (f) { challengeNeuDialog(x.typ, f, { ...c, id: x.id }); return true; }
      await speichern((dd) => {
        if (x.id) { const alt = dd.challenges.find((y) => y.id === x.id); Object.assign(alt, c); return; }
        dd.challenges.push({ id: P.neueNr(dd, 'C'), angelegt: C.heute(), ...c, ...(c.typ === 'wochen52' ? { erledigt: [] } : {}), ...(c.typ === 'ziel' ? { einzahlungen: [] } : {}) });
      }, { art: 'info', text: x.id ? 'Challenge geändert.' : 'Challenge gestartet. Viel Erfolg! 💪' }, { knopf, dialogFehler: (fe) => challengeNeuDialog(x.typ, fe, { ...c, id: x.id }) });
      M.tab = 'm-challenges'; C.zeichne();
      return true;
    }
    case 'm-einzahlung': {
      const b = zahl(x.betrag);
      if (!(b > 0)) { C.menu.close(); return true; }
      await speichern((dd) => { const c = dd.challenges.find((y) => y.id === x.id); c.einzahlungen = [...(c.einzahlungen || []), { datum: x.datum || C.heute(), betrag: b }]; }, { art: 'info', text: `${eur(b)} zurückgelegt. Weiter so!` }, { knopf });
      return true;
    }
    default: return false;
  }
}

// ---------------------------------------------------------------- Kontoauszug laden
export async function csvGewaehlt(datei) {
  M.csvFehler = null;
  try {
    const bytes = new Uint8Array(await datei.arrayBuffer());
    M.csvDatei = { bytes, name: datei.name };
    abgleichNeu();
  } catch (e) { M.csvFehler = 'Die Datei ließ sich nicht lesen: ' + e.message; M.csv = null; }
  C.zeichne();
}
function abgleichNeu() {
  try {
    const csv = leseCsv(csvText(M.csvDatei.bytes), M.csvDatei.name);
    const kt = kontoErkennen(csv, C.z.modell);
    if (kt && kt.name !== d().giro.konto) throw new Error(`Das ist der Auszug von „${kt.name}“, nicht von deinem Giro. Den gleicht das PC-Tool im Bereich „Gemeinsam“ ab.`);
    M.csv = { csv, ...P.abgleichen(csv, C.z.modell, d()) };
    M.csvFehler = null;
  } catch (e) { M.csv = null; M.csvFehler = e.message; }
}
