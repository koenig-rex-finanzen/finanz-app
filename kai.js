// Kais privater Bereich („Mein Geld“) in „Unser Geld“: liest und schreibt Kai_Privat.xlsx (Ordner Apps/Kai, nicht geteilt).
// Aus der gemeinsamen Datei wird nur gelesen (Kais Soll aufs Gemeinschaftskonto). Nichts fließt zurück.
import * as K from './kern/kai.js';
import { eur, monatsname } from './kern/rechnen.js';
import { betragFinden, datumFinden } from './kern/verstehen.js';

let C;
const S = { status: 'aus', fehler: null, ort: null, item: null, k: null, tab: null, meldung: null, laedt: false,
  suche: '', fMonat: '', fArt: '', offen: null, c24: null, tr: null, gelesen: null, topf: null };
const h = (s) => C.h(s);
const zahl = (t) => { const s = String(t ?? '').trim().replace(/\s|€/g, ''); if (!s) return NaN; return Math.round(Number(s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s) * 100) / 100; };
const bt = (n) => (n == null || !isFinite(n) ? '' : Number(n).toFixed(2).replace('.', ','));
const dl = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
const dk = (iso) => (iso ? `${Number(iso.slice(8, 10))}.${Number(iso.slice(5, 7))}.` : '');
const vz = (n) => (n > 0 ? '+' : '') + eur(n);
const mn = (m, j = true) => (m ? monatsname(m, j) : '');
const TABS = [
  ['k-uebersicht', 'Übersicht', '<path d="M4 11l8-6 8 6v8a1 1 0 01-1 1h-4v-5H9v5H5a1 1 0 01-1-1z"/>'],
  ['k-vorschau', 'Vorschau', '<path d="M4 19h16M6 16V9M11 16V5M16 16v-5"/>'],
  ['k-buchungen', 'Buchungen', '<path d="M5 6h14M5 12h14M5 18h9"/>'],
  ['k-toepfe', 'Töpfe', '<path d="M5 9h14l-1.5 10a1 1 0 01-1 .9h-9a1 1 0 01-1-.9z"/><path d="M8 9V7a4 4 0 018 0v2"/>'],
  ['k-depot', 'Depot', '<path d="M4 18l5-6 4 3 7-9"/><path d="M15 6h5v5"/>'],
  ['k-mehr', 'Mehr', '<circle cx="5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="19" cy="12" r="1.6"/>'],
];
const merk = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch {} };
const hol = (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } };

export function init(ctx) { C = ctx; S.tab = hol('ug.ktab') || 'k-uebersicht'; }
export function vergessen() { S.k = null; S.status = 'aus'; S.ort = null; }

// ---------------------------------------------------------------- Laden
export async function laden({ immer = false } = {}) {
  if (S.laedt) return;
  if (!immer && S.status === 'bereit') return;
  S.laedt = true;
  try {
    const ordner = await C.od.privatOrdnerFinden('Kai').catch((e) => { if (e instanceof C.od.AnmeldungNoetig || e instanceof C.od.KeinNetz) throw e; return null; });
    const f = await C.od.kaiDateiFinden(ordner);
    if (!f) { S.status = 'keine-datei'; return; }
    if (!immer && S.item?.eTag === f.item.eTag && S.k) { S.status = 'bereit'; return; }
    const bytes = await C.od.herunterladen(f.ort);
    S.k = await K.leseKai(bytes); S.ort = f.ort; S.item = f.item; S.status = 'bereit'; S.fehler = null; S.gelesen = new Date().toISOString();
  } catch (e) {
    S.fehler = e instanceof C.od.AnmeldungNoetig ? 'Die Microsoft-Anmeldung ist abgelaufen. Bitte im Menü neu anmelden.'
      : e instanceof C.od.KeinNetz ? 'Gerade keine Verbindung zu OneDrive.' : 'Kai_Privat.xlsx ließ sich nicht lesen: ' + e.message;
    if (!S.k) S.status = 'fehler';
  } finally { S.laedt = false; C.zeichne(); }
}

/** Schreibt über eine Excel-Sitzung (Formeln und Auswahllisten bleiben), lädt danach neu. */
async function schreiben(arbeit, meldung, { knopf, dialogFehler } = {}) {
  if (knopf) { knopf.disabled = true; knopf.dataset.t = knopf.textContent; knopf.textContent = 'Wird gespeichert …'; }
  try {
    const r = await C.od.mitSitzung(S.ort, arbeit);
    if (C.menu.open) C.menu.close();
    S.meldung = { art: 'info', text: typeof meldung === 'function' ? meldung(r) : meldung };
    await laden({ immer: true });
    return true;
  } catch (e) {
    const t = e instanceof C.od.AnmeldungNoetig ? 'Die Anmeldung ist abgelaufen. Bitte neu anmelden. Es wurde nichts gespeichert.'
      : e instanceof C.od.KeinNetz ? 'Keine Verbindung zu OneDrive. Es wurde nichts gespeichert.'
        : /423|locked|gesperrt/i.test(e.message) ? 'Die Datei ist gerade in Excel geöffnet und gesperrt. Bitte dort schließen und noch einmal versuchen.'
          : 'Das hat nicht geklappt: ' + e.message + ' Bitte in Excel nachsehen, bevor du es noch einmal versuchst.';
    if (dialogFehler) dialogFehler(t); else { S.meldung = { art: 'warn', text: t }; C.zeichne(); }
    if (knopf) { knopf.disabled = false; knopf.textContent = knopf.dataset.t; }
    return false;
  }
}
const ser = (iso) => K.serial(iso);
const mSer = (m) => (m ? K.serial(m + '-01') : '');
const buchungZeile = (b, nr) => [nr, ser(b.datum), mSer(b.giltFuer), b.art, b.konto, b.gegen || '', Number(b.betrag), b.kategorie || '', b.posten || '', b.topf || '', b.beschreibung || '', b.kommentar || '', b.quelle || 'manuell', b.importId || ''];
/** Neue Buchungen in die ersten freien Zeilen (nur A–N, O/P sind Formeln). */
async function buchungenAnhaengen(api, liste) {
  const [von, bis] = K.BEREICHE.buchungen;
  const a = (await api.lies('Buchungen', `A${von}:A${bis}`)).values.map((r) => String(r[0] ?? ''));
  let letzte = -1; a.forEach((x, i) => { if (x) letzte = i; });
  const start = von + letzte + 1;
  if (start + liste.length - 1 > bis) throw new Error('Im Blatt „Buchungen“ ist kein Platz mehr (Formeln reichen bis Zeile 1004).');
  let n = Math.max(0, ...a.filter((x) => /^P\d+$/.test(x)).map((x) => Number(x.slice(1)))) + 1;
  const nummern = [];
  const werte = liste.map((b) => { const nr = 'P' + String(n++).padStart(4, '0'); nummern.push(nr); return buchungZeile(b, nr); });
  if (werte.length) await api.schreib('Buchungen', `A${start}:N${start + werte.length - 1}`, werte);
  const k = await api.lies('Buchungen', `A${start}:A${start + werte.length - 1}`);
  if (!k.values.every((r, i) => r[0] === nummern[i])) throw new Error('Die Buchungen wurden nicht wie erwartet gespeichert.');
  return nummern;
}
async function zeileVonNr(api, nr) {
  const [von, bis] = K.BEREICHE.buchungen;
  const a = (await api.lies('Buchungen', `A${von}:A${bis}`)).values;
  const i = a.findIndex((r) => r[0] === nr);
  if (i < 0) throw new Error(`${nr} wurde in der Datei nicht gefunden.`);
  return von + i;
}

// ---------------------------------------------------------------- Zeichnen
export function zeichne() {
  document.documentElement.dataset.bereich = 'kai';
  const kopf = C.kopf();
  const wahl = (inhalt, nav = true) => { C.app.innerHTML = `${kopf}<main class="wrap meins kai">${inhalt}</main>${nav ? navHtml() : ''}`; };
  if (S.status === 'aus' || (S.laedt && !S.k)) { wahl('<div class="leer m-laedt">Mein Geld wird geladen …</div>', false); if (S.status === 'aus') laden(); return; }
  if (S.status === 'keine-datei') {
    wahl(`<section class="m-start"><h2>Kai_Privat.xlsx nicht gefunden</h2><p>Die App sucht die Datei im Ordner <b>OneDrive › Apps › Kai</b> (nicht geteilt). Lade sie dort hoch, dann auf „Noch mal suchen“.</p>
      <div class="knoepfe-liste"><button class="knopf" data-k="neu-laden">Noch mal suchen</button></div></section>`, false);
    return;
  }
  if (S.status === 'fehler' && !S.k) { wahl(`<div class="hinweis warn">${h(S.fehler)}</div><button class="knopf" data-k="neu-laden">Noch mal versuchen</button>`, false); return; }
  const r = K.rechneKai(S.k, C.heute());
  const meld = (S.fehler ? `<div class="hinweis info">${h(S.fehler)}</div>` : '') + (S.meldung ? `<div class="hinweis ${S.meldung.art}">${h(S.meldung.text)}</div>` : '');
  const f = { 'k-uebersicht': uebersicht, 'k-vorschau': vorschau, 'k-buchungen': buchungen, 'k-toepfe': toepfeAnsicht, 'k-depot': depotAnsicht, 'k-mehr': mehr }[S.tab] || uebersicht;
  wahl(meld + f(r));
  if (['k-uebersicht', 'k-buchungen'].includes(S.tab) || (S.tab === 'k-toepfe' && !S.topf)) C.app.insertAdjacentHTML('beforeend', '<button class="fab" data-k="eintragen" aria-label="Buchung eintragen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Eintragen</button>');
}
const navHtml = () => `<nav class="sechs"><div class="wrap" role="tablist">${TABS.map(([id, n, p]) => `<button role="tab" aria-selected="${S.tab === id}" data-k="tab" data-arg="${id}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${p}</svg>${n}</button>`).join('')}</div></nav>`;

/** Soll laut gemeinsamer Datei (Aufteilung, „Kai → Gemeinschaftskonto“) – nur gelesen. */
function sollGemeinsam() {
  const a = C.z.modell?.aufteilung || {};
  const out = {};
  for (const [m, w] of Object.entries(a)) { const key = Object.keys(w).find((x) => x.startsWith('Kai → Gemeinschaftskonto')); if (key) out[m] = Math.round(w[key] * 100) / 100; }
  return out;
}
function sollAbweichungen() {
  const soll = sollGemeinsam();
  return S.k.gemeinsam.filter((g) => soll[g.monat] != null && Math.abs((g.soll || 0) - soll[g.monat]) > 0.005).map((g) => ({ ...g, neu: soll[g.monat] }))
    .concat(Object.keys(soll).filter((m) => !S.k.gemeinsam.some((g) => g.monat === m) && m >= S.k.gemeinsam[0]?.monat).map((m) => ({ monat: m, soll: null, neu: soll[m], fehlt: true })));
}

function uebersicht(r) {
  if (r.fehlt) {
    return `<section class="k-warn"><b>Kontostand fehlt</b><p>Ohne den Stand vom c24-Konto zum Stichtag kann die App nicht rechnen.</p><button class="knopf hell" data-k="konto">Kontostand eintragen</button></section>${planIstHtml()}`;
  }
  const luecke = r.status === 'Lücke';
  const warn = luecke ? `<section class="k-warn" role="alert"><span class="k-warn-symbol" aria-hidden="true">!</span>
    <b>Im ${h(mn(r.ersterMinusMonat, false))} fehlen ${eur(r.fehlbetrag)}.</b>
    <p>${r.behalten > 0 ? `Behalte diesen Monat ${eur(r.behalten)} übrig.` : 'Auch ohne variable Ausgaben reicht es nicht – schau in die Vorschau.'}</p>
    ${r.inRuecklage > 0 ? `<button class="knopf hell" data-k="ruecklage">${eur(r.inRuecklage)} in Rücklage legen</button>` : ''}</section>` : '';
  const abw = sollAbweichungen();
  const sollHinweis = abw.length ? `<div class="hinweis warn">Das Soll aufs Gemeinschaftskonto hat sich in der gemeinsamen Datei geändert (${abw.length > 3 ? `${abw.length} Monate, ab ${mn(abw[0].monat)}` : abw.map((a) => mn(a.monat, false)).join(', ')}). <button class="link" data-k="soll-uebernehmen">Übernehmen</button></div>` : '';
  const zeilen = r.vorschau.filter((v) => v.imHorizont);
  return `${warn}${sollHinweis}
    <div class="k-hero"><div class="k-hero-oben"><span>Frei</span><span class="k-ampel ${luecke ? 'rot' : 'gruen'}">${luecke ? 'Lücke' : '✓ Alles gedeckt'}</span></div>
      <div class="k-gross ${r.frei < 0 ? 'neg' : ''}">${eur(r.frei)}</div>
      <div class="k-hero-zeilen"><span>c24 <b>${eur(r.stand)}</b></span>${r.spartopf ? `<span>Spartopf <b>${r.spartopf.fehlt ? 'Stand fehlt' : eur(r.spartopf.stand)}</b></span>` : ''}<span>verplant in Töpfen <b>${eur(r.verplant)}</b></span><span>Rücklage <b>${eur(r.ruecklage)}</b></span></div></div>
    <section><h2>Empfehlung für ${h(mn(r.akt, false))}</h2><div class="rechnung">
      <div class="rz"><span>Bis Monatsende behalten<small>damit die nächsten Monate gedeckt sind</small></span><b>${eur(r.behalten)}</b></div>
      <div class="rz"><span>Davon noch in die Rücklage<small>Topf „${h(K.RUECKLAGE)}“, Stand ${eur(r.ruecklage)}</small></span><b>${eur(r.inRuecklage)}</b></div>
      <div class="rz summe"><span>Frei für Variables diesen Monat</span><b class="${r.variabel < 0 ? 'neg' : ''}">${eur(r.variabel)}</b></div></div>
      ${r.inRuecklage > 0 && !luecke ? `<button class="knopf leise" style="margin-top:10px" data-k="ruecklage">${eur(r.inRuecklage)} in Rücklage legen</button>` : ''}</section>
    <section><h2>Die nächsten ${r.horizont} Monate</h2><div class="k-tabelle"><div class="k-tz kopf"><span>Monat</span><span>Frei am Ende</span><span>Behalten</span><span>Variables</span></div>
      ${zeilen.map((v) => `<button class="k-tz ${v.status === 'Lücke' ? 'luecke' : ''}" data-k="vmonat" data-arg="${v.monat}"><span>${h(mn(v.monat, false))}${v.status === 'Lücke' ? ' ⚠' : ''}</span><span class="${v.freiEnde < 0 ? 'neg' : ''}">${eur(v.freiEnde)}</span><span>${eur(v.behalten)}</span><span class="${v.variabel < 0 ? 'neg' : ''}">${eur(v.variabel)}</span></button>`).join('')}</div>
      <p class="hint" style="margin-top:6px">Ohne variable Ausgaben gerechnet. Antippen zeigt die Posten.</p></section>
    ${planIstHtml()}
    ${spartopfHinweis(r)}
    <p class="hint">Stichtag ${dl(r.stichtag)} · letzte Buchung ${dl(r.letzteBuchung) || '–'} · geladen ${S.gelesen ? new Date(S.gelesen).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '–'}</p>`;
}
function planIstHtml() {
  const m = C.z.monat;
  const p = K.planIst(S.k, m);
  const z = (t, x) => `<div class="k-pz"><span>${t}</span><span>${x.plan == null ? '–' : eur(x.plan)}</span><span>${eur(x.ist)}</span></div>`;
  return `<section><h2>Plan und Ist im ${h(mn(m, false))}</h2><div class="k-pi"><div class="k-pz kopf"><span></span><span>Plan</span><span>Ist</span></div>
    ${z('Einnahmen', p.einnahmen)}${z('Gemeinschaftskonto', p.gemeinsam)}${z('Feste Kosten', p.kosten)}${z('In Töpfe', p.sparraten)}${z('Variables', p.variabel)}</div>
    <p class="hint" style="margin-top:6px">Monat oben mit den Pfeilen wechseln.</p></section>`;
}

function vorschau(r) {
  if (r.fehlt) return `<section class="k-warn"><b>Kontostand fehlt</b><p>Die Vorschau braucht den Stand zum Stichtag.</p><button class="knopf hell" data-k="konto">Kontostand eintragen</button></section>`;
  return `<section><h2>Vorschau ab ${h(mn(r.akt, false))}</h2><p class="hint">Gewarnt wird für die ersten ${r.horizont} Monate (ab 0,01 € Lücke). Danach „Ausblick“.</p>
    <div class="k-vorschau">${r.vorschau.map((v) => {
      const auf = S.offen === v.monat;
      const posten = (titel, liste, feld = 'rest') => liste.filter((x) => x[feld] || x.plan).map((x) => `<div class="k-pp"><span>${h(x.posten || x.name)}${x.ist ? ` <small>schon ${eur(x.ist)}</small>` : ''}</span><b>${eur(x[feld] ?? x.betrag)}</b></div>`).join('') || `<div class="k-pp hint"><span>${titel}: nichts</span></div>`;
      return `<div class="k-vm ${v.status === 'Lücke' ? 'luecke' : v.status === 'Ausblick' ? 'ausblick' : ''}"><button class="k-vm-kopf" data-k="vmonat" data-arg="${v.monat}" aria-expanded="${auf}">
        <span class="t">${h(mn(v.monat))}</span><span class="k-status ${v.status === 'Lücke' ? 'rot' : v.status === 'OK' ? 'gruen' : ''}">${h(v.status)}</span>
        <span class="k-vm-zahlen"><span>Ergebnis <b class="${v.ergebnis < 0 ? 'neg' : ''}">${vz(v.ergebnis)}</b></span><span>frei am Ende <b class="${v.freiEnde < 0 ? 'neg' : ''}">${eur(v.freiEnde)}</b></span>
        <span>behalten <b>${eur(v.behalten)}</b></span><span>Variables <b class="${v.variabel < 0 ? 'neg' : ''}">${eur(v.variabel)}</b></span></span></button>
        ${auf ? `<div class="k-vm-detail"><h3>Einnahmen ${eur(v.summen.einnahmen)}</h3>${posten('Einnahmen', v.einnahmen)}
          <h3>Gemeinschaftskonto ${eur(v.summen.gemeinsam)}</h3><div class="k-pp"><span>Soll ${v.gemeinsam.soll == null ? 'fehlt in „Gemeinsam“' : eur(v.gemeinsam.soll)}${v.gemeinsam.ist ? ` <small>schon ${eur(v.gemeinsam.ist)}</small>` : ''}</span><b>${eur(v.gemeinsam.rest)}</b></div>
          <h3>Feste Kosten ${eur(v.summen.kosten)}</h3>${posten('Kosten', v.kosten)}
          <h3>Sparraten ${eur(v.summen.sparraten)}</h3>${v.sparraten.filter((s) => s.betrag).map((s) => `<div class="k-pp"><span>${h(s.name)}</span><b>${eur(s.betrag)}</b></div>`).join('') || '<div class="k-pp hint"><span>keine (im aktuellen Monat schon zugeteilt)</span></div>'}
          <div class="k-pp summe"><span>Frei am Anfang ${eur(v.freiAnfang)} + Ergebnis</span><b>${eur(v.freiEnde)}</b></div></div>` : ''}</div>`;
    }).join('')}</div></section>`;
}

function buchungen() {
  const k = S.k;
  const q = S.suche.toLowerCase();
  const liste = k.buchungen.filter((b) => (!S.fMonat || b.giltFuer === S.fMonat) && (!S.fArt || b.art === S.fArt)
    && (!q || [b.posten, b.gegen, b.beschreibung, b.kommentar, b.kategorie, b.nr, b.topf].some((x) => (x || '').toLowerCase().includes(q))))
    .sort((a, b) => (b.datum || '').localeCompare(a.datum || '') || b.nr.localeCompare(a.nr));
  const monate = [...new Set(k.buchungen.map((b) => b.giltFuer).filter(Boolean))].sort().reverse();
  return `<section><h2>Buchungen</h2>
    <div class="k-filter"><input type="search" id="k-suche" placeholder="Suchen …" value="${h(S.suche)}" autocomplete="off">
      <select id="k-fmonat"><option value="">alle Monate</option>${monate.map((m) => `<option value="${m}" ${m === S.fMonat ? 'selected' : ''}>${h(mn(m))}</option>`).join('')}</select>
      <select id="k-fart"><option value="">alle Arten</option>${K.ARTEN.map((a) => `<option ${a === S.fArt ? 'selected' : ''}>${a}</option>`).join('')}</select></div>
    <p class="hint" style="margin:6px 0 10px">${liste.length} Buchungen · Summe Wirkung aufs c24-Konto ${vz(Math.round(liste.reduce((s, b) => s + K.wirkungAuf(b, k.konto.name), 0) * 100) / 100)}</p>
    ${liste.length ? `<div class="buchungen">${liste.map((b) => buchungHtml(b)).join('')}</div>`
    : '<div class="leer">Keine Buchungen gefunden.</div>'}
  </section>`;
}
function buchungHtml(b, topfName) {
  const umb = b.art === K.UMBUCHUNG;
  const fremd = !umb && b.konto && b.konto !== S.k.konto.name;
  const wirk = topfName ? K.topfeffekt(b) : null;
  const zeichen = topfName ? (wirk > 0 ? '+' : wirk < 0 ? '−' : '') : umb ? '⇄ ' : b.art === 'Einnahme' ? '+' : b.art === 'Ausgabe' ? '−' : '';
  return `<div class="bu-wrap"><button class="bu" aria-expanded="false"><span class="t">${umb ? `${h(b.konto)} → ${h(b.gegen || '?')}` : h(b.posten || b.gegen || b.beschreibung || b.art)}${b.topf && !topfName && !umb ? ` <span class="chip aus-topf">${h(b.topf)}</span>` : ''}${fremd ? ` <span class="chip">${h(b.konto)}</span>` : ''}</span>
      <span class="betrag ${(topfName ? wirk > 0 : b.art === 'Einnahme') ? 'plus' : ''}">${zeichen}${eur(b.betrag)}</span>
      <span class="r">${dk(b.datum)} · ${h(b.art)}${b.giltFuer && b.giltFuer !== b.datum?.slice(0, 7) ? ` <span class="okt">gilt für ${h(mn(b.giltFuer, false))}</span>` : ''}${b.gegen && b.posten && !umb ? ' · ' + h(b.gegen) : ''}</span>
      <span class="more"><dl>${[['Nummer', b.nr], ['Konto', umb ? `${b.konto} → ${b.gegen}` : b.konto], ['Datum', dl(b.datum)], ['Gilt für', mn(b.giltFuer)], ['Kategorie (Bank)', b.kategorie], ['Topf', b.topf], ['Beschreibung', b.beschreibung], ['Kommentar', b.kommentar], ['Quelle', b.quelle]].filter(([, w]) => w).map(([a, w]) => `<dt>${a}</dt><dd>${h(w)}</dd>`).join('')}</dl></span></button>
      <div class="bu-aktion"><button class="knopf leise klein" data-k="bearbeiten" data-arg="${h(b.nr)}">Bearbeiten</button><button class="knopf leise klein" data-k="loeschen-frage" data-arg="${h(b.nr)}">Löschen</button></div></div>`;
}

// ---------------------------------------------------------------- Töpfe (eigener Reiter)
const eigeneKonten = () => (S.k.konten || [S.k.konto]).filter((kt) => kt === S.k.konto || kt.rolle === 'Sparen');
function spartopfKarte(r, kurz) {
  const sp = r.spartopf;
  if (!sp) return '';
  if (sp.fehlt) return `<section class="k-spar"><div class="k-spar-kopf"><span data-hilfe-schluessel="spartopf">${h(sp.name)}</span></div><p>Für den Spartopf fehlt noch der Stand zum Stichtag. Bis dahin rechnet die App so, als läge alles auf c24.</p>
    <button class="knopf hell" data-k="konto" data-arg="${sp.zeile}">Stand eintragen</button></section>`;
  const o = sp.offen;
  const zeile = Math.abs(o) < 0.005 ? '<span class="k-ampel gruen">✓ passt genau</span>'
    : o > 0 ? `<span class="k-ampel rot">noch ${eur(o)} überweisen</span>` : `<span class="k-ampel">${eur(-o)} mehr drauf</span>`;
  return `<section class="k-spar"><div class="k-spar-kopf"><span data-hilfe-schluessel="spartopf">${h(sp.name)}</span>${zeile}</div>
    <div class="k-gross">${eur(sp.stand)}</div>
    <div class="k-hero-zeilen"><span>verplant in Töpfen <b>${eur(sp.verplant)}</b></span><span>Rücklage auf c24 <b>${eur(r.ruecklage)}</b></span></div>
    ${!kurz && o > 0.005 ? `<button class="knopf hell" style="margin-top:12px" data-k="umbuchen" data-arg="${o}">${eur(o)} aufs Spartopf umbuchen</button>` : ''}</section>`;
}
function spartopfHinweis(r) {
  const sp = r.spartopf;
  if (sp && !sp.fehlt && sp.offen > 0.005) return `<div class="hinweis info">Aufs ${h(sp.name)} fehlen noch ${eur(sp.offen)} für deine Töpfe. <button class="link" data-k="tab" data-arg="k-toepfe">Zu den Töpfen</button></div>`;
  return '';
}
function toepfeAnsicht(r) {
  if (S.topf) { const t = r.toepfe.find((x) => x.name === S.topf); if (t) return topfSeite(t); S.topf = null; }
  const liste = (arr) => arr.map((t) => `<button class="k-topf" data-k="topf-auf" data-arg="${h(t.name)}"><span class="n">${h(t.name)}<small>${t.sparrate ? `${eur(t.sparrate)} im Monat` : 'keine Sparrate'}${t.bis ? ` · bis ${h(mn(t.bis))}` : ''}</small></span><b class="${t.stand < 0 ? 'neg' : ''}">${eur(t.stand)}</b><span class="pfeil" aria-hidden="true">›</span></button>`).join('');
  const sp = r.toepfe.filter((t) => !t.frei), c = r.toepfe.filter((t) => t.frei);
  return `${spartopfKarte(r)}
    <section><h2>${r.spartopf ? `Auf dem ${h(r.spartopf.name)}` : 'Töpfe'}</h2><div class="k-toepfe">${liste(sp) || '<div class="leer">Keine Töpfe.</div>'}</div>
      <div class="k-pp summe"><span>Summe</span><b>${eur(r.verplant)}</b></div></section>
    ${c.length ? `<section><h2>Auf c24 (zählt als frei)</h2><div class="k-toepfe">${liste(c)}</div></section>` : ''}
    ${S.k.toepfe.length < 10 ? '<button class="knopf leise" data-k="topf">Neuer Topf</button>' : ''}
    <p class="hint" style="margin-top:10px">Die Sparraten werden jeden Monat zugeteilt. Wenn du Geld aufs Spartopf überweist, trägst du das als „Umbuchung“ ein. Töpfe ändern sich dadurch nicht.</p>`;
}
function topfSeite(t) {
  const k = S.k;
  const nach = (b) => b.datum && (!k.konto.stichtag || b.datum > k.konto.stichtag);
  const bu = k.buchungen.filter((b) => b.topf === t.name && b.art !== K.UMBUCHUNG).sort((a, b) => (b.datum || '').localeCompare(a.datum || '') || b.nr.localeCompare(a.nr));
  return `<div class="k-topfkopf"><button class="k-zurueck" data-k="topf-zu" aria-label="Zurück"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5l-7 7 7 7"/></svg></button><h2>${h(t.name)}</h2></div>
    <section class="k-spar"><div class="k-spar-kopf"><span>Stand ${h(mn(K.rechneKai(k, C.heute()).akt, false))}</span>${t.frei ? '<span class="k-ampel">zählt als frei</span>' : ''}</div><div class="k-gross ${t.stand < 0 ? 'neg' : ''}">${eur(t.stand)}</div>
      <div class="k-hero-zeilen"><span>Sparrate <b>${eur(t.sparrate)}</b></span>${t.ab ? `<span>ab <b>${h(mn(t.ab))}</b></span>` : ''}${t.bis ? `<span>bis <b>${h(mn(t.bis))}</b></span>` : ''}</div></section>
    <section><h2>So setzt sich der Stand zusammen</h2><div class="rechnung">
      <div class="rz"><span>Anfangsbestand<small>zum Stichtag ${dl(k.konto.stichtag) || '–'}</small></span><b>${eur(t.anfang)}</b></div>
      <div class="rz"><span>Sparraten zugeteilt<small>${t.ab ? `seit ${h(mn(t.ab))}` : 'kein Beginn gesetzt'}</small></span><b>${vz(t.zugeteilt)}</b></div>
      <div class="rz"><span>Hineingelegt / entnommen<small>Buchungen mit diesem Topf nach dem Stichtag</small></span><b>${vz(t.bewegt)}</b></div>
      <div class="rz summe"><span>Stand</span><b>${eur(t.stand)}</b></div></div>
      ${t.bemerkung ? `<p class="hint" style="margin-top:6px">${h(t.bemerkung)}</p>` : ''}
      <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap"><button class="knopf leise klein" data-k="topf-buchen" data-arg="In Topf legen">In Topf legen</button><button class="knopf leise klein" data-k="topf-buchen" data-arg="Aus Topf nehmen">Aus Topf nehmen</button><button class="knopf leise klein" data-k="topf" data-arg="${t.zeile}">Topf ändern</button></div></section>
    <section><h2>Buchungen</h2>${bu.length ? `<div class="buchungen">${bu.map((b) => buchungHtml(b, t.name).replace('class="bu-wrap"', `class="bu-wrap${nach(b) ? '' : ' alt'}"`)).join('')}</div><p class="hint" style="margin-top:6px">Ausgegraut: vor dem Stichtag, steckt schon im Anfangsbestand.</p>` : '<div class="leer">Noch keine Buchungen mit diesem Topf.</div>'}</section>`;
}

function depotAnsicht() {
  const d = K.depot(S.k);
  const pos = d.positionen.filter((p) => p.stueck > 1e-9 || S.depotAlle);
  const pz = (n) => (n == null ? '–' : (n * 100).toLocaleString('de-DE', { maximumFractionDigits: 1, minimumFractionDigits: 1 }) + ' %');
  const st = (n) => n.toLocaleString('de-DE', { maximumFractionDigits: 6 });
  return `<section><h2>Trade Republic</h2>
    <div class="topfsumme">
      <div class="z"><span>Bargeld bei TR</span><b>${eur(d.bargeld)}</b></div>
      <div class="z"><span>Eingesetzt in offene Positionen</span><b>${eur(d.eingesetzt)}</b></div>
      <div class="z"><span>Aktueller Wert <small class="hint">(nur mit Kurs)</small></span><b>${eur(d.wert)}</b></div>
      <div class="z"><span>Gewinn/Verlust offen <small class="hint">(nur mit Kurs)</small></span><b class="${d.gv < 0 ? 'neg' : 'plus'}">${vz(d.gv)}</b></div>
      <div class="z"><span>Dividenden netto</span><b>${eur(d.dividenden)}</b></div>
      <div class="z"><span>Realisiert</span><b>${vz(d.realisiert)}</b></div>
      <div class="z"><span>Eingezahlt / ausgezahlt</span><b>${eur(d.eingezahlt)} / ${eur(d.ausgezahlt)}</b></div></div>
    <p class="hint" style="margin-top:6px">Stand der Daten ${dl(d.stand) || '–'}. Kurse trägst du von Hand ein (Position antippen). Nur Zahlen, keine Empfehlungen.</p>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap"><label class="knopf leise klein" for="k-tr">TR-Export einlesen</label><input type="file" id="k-tr" accept=".csv,.xlsx,text/csv" hidden>
      <button class="knopf leise klein" data-k="depot-alle">${S.depotAlle ? 'Nur offene Positionen' : 'Auch geschlossene zeigen'}</button></div>
    ${trImportHtml()}
  </section>
  <section><h2>Positionen</h2><div class="k-depot">${pos.map((p) => `<button class="k-pos" data-k="kurs" data-arg="${h(p.position)}">
      <span class="n">${h(p.name || p.position)}<small>${h(p.art || '')} · ${st(p.stueck)} Stück</small></span>
      <span class="w">${p.wert != null ? eur(p.wert) : '<span class="hint">Kurs fehlt</span>'}<small class="${p.gv < 0 ? 'neg' : p.gv > 0 ? 'plus' : ''}">${p.gv != null ? `${vz(p.gv)} · ${pz(p.gvProzent)}` : `Einstand ${eur(p.einstand)}`}</small></span>
      <span class="k-pos-detail">Einstand ${eur(p.einstand)} · Ø ${p.kaufkurs != null ? eur(p.kaufkurs) : '–'} · Kurs ${p.kurs != null ? eur(p.kurs) + (p.kursVom ? ` vom ${dl(p.kursVom)}` : '') : '–'}${p.dividenden ? ` · Div. ${eur(p.dividenden)}` : ''}${p.realisiert ? ` · real. ${vz(p.realisiert)}` : ''}${p.letzterKurs != null ? ` · letzter TR-Kurs ${eur(p.letzterKurs)} (${dl(p.letzterKursAm)})` : ''}</span></button>`).join('')}</div></section>`;
}
function trImportHtml() {
  if (!S.tr) return '';
  if (S.tr.fehler) return `<p class="fehler">${h(S.tr.fehler)}</p>`;
  const t = S.tr;
  return `<div class="m-abgleich"><div class="m-ab-kopf"><b>TR-Export: ${t.neu.length} neue Zeilen</b><button class="link" data-k="tr-zu">schließen</button></div>
    ${t.neu.length ? `<div class="buchungen">${t.neu.slice(0, 50).map((n) => `<div class="m-ab-zeile"><div><div class="t">${h(n.typ)} · ${h(n.name || n.beschreibung || '')}</div><div class="hint">${dl(n.datum)}${n.stueck != null ? ` · ${n.stueck} Stück` : ''}</div></div><b class="${n.betrag > 0 ? 'plus' : ''}">${vz(n.betrag)}</b></div>`).join('')}</div>
    ${t.neu.length > 50 ? `<p class="hint">… und ${t.neu.length - 50} weitere.</p>` : ''}
    ${t.positionenNeu.length ? `<p class="hint">Neue Positionen im Depot: ${t.positionenNeu.map((p) => h(p.name)).join(', ')}.</p>` : ''}
    ${t.unbekannt.length ? `<p class="hint warnfarbe">${t.unbekannt.length} Zeilen mit unbekanntem Typ werden nicht übernommen.</p>` : ''}
    <button class="knopf" data-k="tr-uebernehmen">${t.neu.length} Zeilen übernehmen</button>` : '<p class="hint">Alles schon in der Datei.</p>'}</div>`;
}

function mehr(r) {
  const k = S.k;
  const soll = sollGemeinsam();
  const abw = sollAbweichungen();
  const offen = k.offen.filter((o) => o.status !== 'erledigt');
  return `<section><h2>c24-Kontoauszug</h2><p class="hint">CSV aus dem Online-Banking (Transaktionen › Export › CSV). Die App zeigt, was neu ist und was du schon von Hand eingetragen hast. Nichts wird ohne deinen Klick übernommen.</p>
      <label class="knopf leise" for="k-c24" style="margin-top:10px">c24-Export einlesen</label><input type="file" id="k-c24" accept=".csv,text/csv" hidden>${c24Html()}</section>
    <section><h2>Gemeinsam (Gemeinschaftskonto)</h2>
      <div class="k-tabelle gem"><div class="k-tz kopf"><span>Monat</span><span>Soll</span><span>überwiesen</span><span>offen</span></div>
      ${k.gemeinsam.map((g) => { const v = r.vorschau.find((x) => x.monat === g.monat); const istB = K.planIst(k, g.monat).gemeinsam.ist; const neu = soll[g.monat];
        return `<div class="k-tz"><span>${h(mn(g.monat, false))} ${g.monat.slice(2, 4)}</span><span>${eur(g.soll || 0)}${neu != null && Math.abs(neu - (g.soll || 0)) > 0.005 ? `<small class="warnfarbe"> neu ${eur(neu)}</small>` : ''}</span><span>${eur(istB)}</span><span>${eur(Math.max(0, (g.soll || 0) - istB))}</span></div>`; }).join('')}</div>
      <p class="hint" style="margin-top:6px">Soll kommt aus der gemeinsamen Datei (Blatt Aufteilung, nur gelesen, bei jedem Öffnen). ${abw.length ? `<button class="link" data-k="soll-uebernehmen">${abw.length} Änderung${abw.length > 1 ? 'en' : ''} übernehmen</button>` : 'Alles aktuell.'}</p></section>
    <section><h2>Feste Kosten</h2><div class="feste">${k.kosten.map((x) => `<div class="f-zeile"><div><div class="t">${h(x.posten)}</div><div class="hint">${h(x.rhythmus)}${x.rhythmus === 'jährlich' && x.faellig ? ` im ${mn('2000-' + String(x.faellig).padStart(2, '0'), false)}` : ''}${x.tag ? ` · am ${x.tag}.` : ''}${x.ab ? ` · ab ${mn(x.ab)}` : ''}${x.bis ? ` · bis ${mn(x.bis)}` : ''}</div></div><b>${eur(x.betrag)}</b><button class="knopf leise klein" data-k="posten" data-arg="kosten:${x.zeile}">Ändern</button></div>`).join('')}</div>
      ${k.kosten.length < 20 ? '<button class="knopf leise" style="margin-top:10px" data-k="posten" data-arg="kosten:">Neue feste Kosten</button>' : ''}</section>
    <section><h2>Einnahmen</h2><div class="feste">${k.einnahmen.map((x) => `<div class="f-zeile"><div><div class="t">${h(x.posten)}</div><div class="hint">${h(x.rhythmus)}${x.eingang ? ` · ${h(x.eingang)}` : ''}${x.ab ? ` · ab ${mn(x.ab)}` : ''}${x.bis ? ` · bis ${mn(x.bis)}` : ''}</div></div><b class="plus">${eur(x.betrag)}</b><button class="knopf leise klein" data-k="posten" data-arg="einnahmen:${x.zeile}">Ändern</button></div>`).join('')}</div>
      ${k.einnahmen.length < 8 ? '<button class="knopf leise" style="margin-top:10px" data-k="posten" data-arg="einnahmen:">Neue Einnahme</button>' : ''}</section>
    <section><h2>Offene Punkte</h2>${offen.length ? `<div class="feste">${offen.map((o) => `<div class="f-zeile"><div><div class="t">${h(o.punkt)}</div>${o.notiz ? `<div class="hint">${h(o.notiz)}</div>` : ''}</div><span></span><button class="knopf leise klein" data-k="erledigt" data-arg="${o.zeile}">Erledigt</button></div>`).join('')}</div>` : '<div class="leer">Nichts offen.</div>'}
      ${k.offen.length - offen.length ? `<p class="hint" style="margin-top:6px">${k.offen.length - offen.length} erledigt.</p>` : ''}</section>
    <section><h2>Konten</h2><div class="feste">${eigeneKonten().map((kt) => `<div class="f-zeile"><div><div class="t">${h(kt.name)}</div><div class="hint">Stand zum Stichtag ${kt.anfang == null ? 'fehlt' : `${eur(kt.anfang)} am ${dl(kt.stichtag) || '–'}`}</div></div><b>${(() => { const st = K.standVon(k, kt); return st == null ? '–' : eur(st); })()}</b><button class="knopf leise klein" data-k="konto" data-arg="${kt.zeile}">Ändern</button></div>`).join('')}</div>
      ${S.item?.webUrl ? `<a class="knopf leise" style="margin-top:10px" href="${h(S.item.webUrl)}" target="_blank" rel="noopener">In Excel öffnen</a>` : ''}</section>`;
}

function c24Html() {
  const a = S.c24;
  if (!a) return '';
  if (a.fehler) return `<p class="fehler">${h(a.fehler)}</p>`;
  const zaehl = (s) => a.zeilen.filter((x) => x.status === s).length;
  const postenListe = postenAuswahl();
  return `<div class="m-abgleich"><div class="m-ab-kopf"><b>${a.zeilen.length} Zeilen im Export</b><button class="link" data-k="c24-zu">schließen</button></div>
    <div class="topfsumme"><div class="z"><span>Schon importiert</span><b>${zaehl('importiert')}</b></div><div class="z"><span>Von Hand eingetragen (nur Import-ID ergänzen)</span><b>${zaehl('ergaenzen')}</b></div>
      ${zaehl('gegenbuchung') ? `<div class="z"><span>Spiegelzeilen vom Spartopf (stehen schon als Umbuchung)</span><b>${zaehl('gegenbuchung')}</b></div>` : ''}<div class="z"><span>Bitte prüfen</span><b class="${zaehl('pruefen') ? 'warnfarbe' : ''}">${zaehl('pruefen')}</b></div><div class="z"><span>Neu</span><b>${zaehl('neu')}</b></div></div>
    <form id="k-c24-form"><div class="buchungen">${a.zeilen.map((x, i) => x.status === 'importiert' || x.status === 'gegenbuchung' ? '' : `<div class="m-ab-zeile k-c24 ${x.status}">
      <label class="haken"><input type="checkbox" name="w${i}" ${x.status === 'pruefen' ? '' : 'checked'}><span><b>${h(x.vorschlag?.gegen || x.z.empfaenger || x.z.typ)}</b><small>${dk(x.z.datum)} · ${h((x.z.zweck || x.z.typ || '').slice(0, 60))}</small></span></label>
      <b class="${x.z.betrag > 0 ? 'plus' : ''}">${vz(x.z.betrag)}</b>
      ${x.status === 'ergaenzen' ? `<div class="hint">Eingetragen als ${h(x.treffer.nr)} (${h(x.treffer.posten || x.treffer.gegen || '')}) – nur die Import-ID wird ergänzt.</div>`
        : `${x.status === 'pruefen' ? `<div class="hint warnfarbe">Unklar: ${h(x.grund)}${x.treffer ? ` (${h(x.treffer.nr)})` : ''}. Nur übernehmen, wenn es wirklich neu ist.</div>` : ''}
        ${x.vorschlag.art === K.UMBUCHUNG ? `<div class="hint">Umbuchung ${h(x.vorschlag.konto)} → ${h(x.vorschlag.gegen)}</div>` : ''}
        <div class="zwei"><label class="feld" ${x.vorschlag.art === K.UMBUCHUNG ? 'hidden' : ''}><span>Posten</span><select name="p${i}">${['', ...postenListe].map((p) => `<option ${p === (x.vorschlag.posten || '') ? 'selected' : ''} value="${h(p)}">${h(p || '– variabel –')}</option>`).join('')}</select></label>
        <label class="feld"><span>Gilt für</span><input type="month" name="g${i}" value="${x.vorschlag.giltFuer}"></label></div>`}</div>`).join('')}</div>
    <div class="knoepfe"><button class="knopf" type="submit">Ausgewählte übernehmen</button></div></form></div>`;
}
const postenAuswahl = () => [...new Set([...S.k.einnahmen.map((e) => e.posten), ...S.k.kosten.map((x) => x.posten), ...K.POSTEN_EXTRA].filter(Boolean))];

// ---------------------------------------------------------------- Dialoge
const dialog = (html) => { C.menu.innerHTML = html; if (!C.menu.open) C.menu.showModal(); };
const feld = (name, label, wert = '', extra = '') => `<label class="feld"><span>${label}</span><input name="${name}" value="${h(wert ?? '')}" autocomplete="off" ${extra}></label>`;
const auswahl = (name, label, liste, wert, leer) => `<label class="feld"><span>${label}</span><select name="${name}">${leer != null ? `<option value="">${h(leer)}</option>` : ''}${liste.map((x) => `<option ${x === wert ? 'selected' : ''} value="${h(x)}">${h(x)}</option>`).join('')}</select></label>`;
const knoepfe = (fehler, text = 'Speichern', extra = '') => `${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}<div class="knoepfe"><button class="knopf" type="submit">${text}</button>${extra}<button class="knopf leise" type="button" id="schliessen">Abbrechen</button></div>`;
const SR = window.SpeechRecognition || window.webkitSpeechRecognition || null;

function eintragenDialog(fehler, v0) {
  const v = { art: 'Ausgabe', datum: C.heute(), giltFuer: C.heute().slice(0, 7), betrag: '', posten: '', topf: '', gegen: '', beschreibung: '', kommentar: '', konto: S.k.konto.name, ...(v0 || {}) };
  const kontoNamen = eigeneKonten().map((kt) => kt.name);
  const umb = v.art === K.UMBUCHUNG;
  const nachVorschlag = umb && kontoNamen.includes(v.gegen) ? v.gegen : kontoNamen.find((n) => n !== v.konto) || '';
  const bearb = Boolean(v.nr);
  dialog(`<form id="k-eintragen" novalidate><h2 id="menu-titel">${bearb ? h(v.nr) + ' ändern' : 'Buchung eintragen'}</h2>
    ${bearb ? '' : `<div class="sprechen">${SR ? '<button type="button" class="mikro" id="k-mikro" aria-label="Sprechen"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg></button>' : ''}
      <input id="k-satz" autocomplete="off" placeholder="z. B. „gestern 26,40 Tanken“"><button type="button" class="knopf leise klein" id="k-satz-los">Übernehmen</button></div><p class="erkannt" id="k-erkannt"></p>`}
    <div class="seg fuenf-seg">${K.ARTEN.map((a) => `<label><input type="radio" name="art" value="${a}" ${v.art === a ? 'checked' : ''}><span>${a.replace(' legen', '').replace(' nehmen', '')}</span></label>`).join('')}</div>
    ${feld('betrag', 'Betrag in €', v.betrag === '' ? '' : bt(Number(v.betrag)), 'inputmode="decimal" placeholder="0,00"')}
    <div class="zwei">${feld('datum', 'Datum', v.datum, 'type="date"')}${feld('giltFuer', 'Gilt für', v.giltFuer, 'type="month"')}</div>
    ${kontoNamen.length > 1 ? `<div class="zwei">${auswahl('konto', umb ? 'Von Konto' : 'Konto', kontoNamen, v.konto)}<span class="nur-umb" ${umb ? '' : 'hidden'}>${auswahl('nach', 'Nach Konto', kontoNamen, nachVorschlag)}</span></div>` : `<input type="hidden" name="konto" value="${h(v.konto)}">`}
    <div class="nicht-umb" ${umb ? 'hidden' : ''}>${auswahl('posten', 'Posten', postenAuswahl(), v.posten, '– variabel –')}
    ${auswahl('topf', 'Topf', S.k.toepfe.map((t) => t.name), v.topf, 'kein Topf')}
    ${feld('gegen', 'Gegenüber', umb ? '' : v.gegen, 'placeholder="z. B. Aral"')}</div>${feld('beschreibung', 'Beschreibung (optional)', v.beschreibung)}${feld('kommentar', 'Kommentar (optional)', v.kommentar)}
    ${bearb ? `<input type="hidden" name="nr" value="${h(v.nr)}">` : ''}
    <p class="zusammen" id="k-zus"></p>${knoepfe(fehler, bearb ? 'Speichern' : 'Eintragen')}</form>`);
  const f = document.getElementById('k-eintragen');
  const zus = () => {
    const x = Object.fromEntries(new FormData(f)); const b = zahl(x.betrag); const u = x.art === K.UMBUCHUNG;
    f.querySelectorAll('.nur-umb').forEach((e) => { e.hidden = !u; }); f.querySelectorAll('.nicht-umb').forEach((e) => { e.hidden = u; });
    const lab = f.querySelector('select[name=konto]')?.closest('label')?.querySelector('span'); if (lab?.firstChild) { lab.firstChild.textContent = u ? 'Von Konto' : 'Konto'; const q = lab.querySelector('.hilfe'); if (q) q.dataset.hilfe = u ? 'Von Konto' : 'Konto'; }
    document.getElementById('k-zus').textContent = !(b > 0) ? '' : u ? `Umbuchung über ${eur(b)} von ${x.konto} nach ${x.nach || '?'}. Töpfe bleiben gleich.`
      : `${x.art} über ${eur(b)}${x.posten ? ', ' + x.posten : ''}${x.topf ? ', Topf ' + x.topf : ''}${eigeneKonten().length > 1 ? ', ' + x.konto : ''}, gilt für ${mn(x.giltFuer)}.`;
  };
  f.addEventListener('input', zus);
  f.addEventListener('change', (e) => { if (e.target.name === 'datum' || e.target.name === 'posten') f.giltFuer.value = K.giltFuer(f.datum.value || C.heute(), f.posten.value); zus(); });
  zus();
  const satz = document.getElementById('k-satz');
  if (satz) {
    const hin = document.getElementById('k-erkannt');
    const los = () => {
      let rest = ' ' + satz.value + ' ';
      const er = [];
      const d = datumFinden(rest, C.heute()); if (d) { f.datum.value = d.wert; const i = rest.toLowerCase().indexOf(d.treffer); if (i >= 0) rest = rest.slice(0, i) + ' ' + rest.slice(i + d.treffer.length); er.push('Datum'); }
      const b = betragFinden(rest); if (b) { f.betrag.value = bt(b.wert); rest = rest.replace(b.treffer, ' '); er.push('Betrag'); }
      if (/\b(bekommen|gehalt|gutschrift|erstatt)/i.test(rest)) f.querySelector('input[name=art][value=Einnahme]').checked = true;
      const wort = rest.replace(/\s(für|bei|beim|im|an|euro|€)\s/gi, ' ').replace(/[,.;:!?]/g, ' ').replace(/\s+/g, ' ').trim();
      const p = postenAuswahl().find((x) => wort && x.toLowerCase().includes(wort.toLowerCase().split(' ')[0]));
      if (p) { f.posten.value = p; er.push('Posten'); } else if (wort) { f.gegen.value = wort.charAt(0).toUpperCase() + wort.slice(1); er.push('Gegenüber'); }
      f.giltFuer.value = K.giltFuer(f.datum.value, f.posten.value);
      hin.textContent = `Vorausgefüllt${er.length ? ': ' + er.join(', ') : ''}. Bitte kurz prüfen.`; zus();
    };
    document.getElementById('k-satz-los').addEventListener('click', los);
    satz.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); los(); } });
    const mik = document.getElementById('k-mikro');
    if (mik) {
      let er = null;
      mik.addEventListener('click', () => {
        if (er) { er.stop(); return; }
        er = new SR(); er.lang = 'de-DE'; er.interimResults = true;
        mik.classList.add('hoert'); hin.textContent = 'Ich höre zu …';
        er.onresult = (e) => { satz.value = [...e.results].map((r) => r[0].transcript).join(' '); if (e.results[e.results.length - 1].isFinal) los(); };
        er.onerror = () => { hin.textContent = 'Spracherkennung gerade nicht möglich. Bitte tippen.'; };
        er.onend = () => { mik.classList.remove('hoert'); er = null; };
        try { er.start(); } catch { mik.classList.remove('hoert'); er = null; }
      });
    }
  }
}

function kontoDialog(zeile, fehler) {
  const k = eigeneKonten().find((x) => x.zeile === Number(zeile)) || S.k.konto;
  dialog(`<form id="k-konto" novalidate><h2 id="menu-titel">Kontostand ${h(k.name)}</h2><p class="hint">Stand laut Bank am Ende des Stichtags. Buchungen bis dahin ändern den Stand nicht, zählen aber im Ist.</p>
    ${feld('anfang', 'Stand in €', bt(k.anfang), 'inputmode="decimal"')}${feld('stichtag', 'Stichtag', k.stichtag || C.heute(), 'type="date"')}<input type="hidden" name="zeile" value="${k.zeile}">${knoepfe(fehler)}</form>`);
}
function ruecklageDialog(betrag, fehler) {
  dialog(`<form id="k-ruecklage" novalidate><h2 id="menu-titel">In Rücklage legen</h2>
    <p>Legt eine Buchung <b>„In Topf legen“</b> an, Topf <b>„${h(K.RUECKLAGE)}“</b>. Das Geld bleibt auf dem c24-Konto, ist aber für später reserviert.</p>
    ${feld('betrag', 'Betrag in €', bt(betrag), 'inputmode="decimal"')}${feld('datum', 'Datum', C.heute(), 'type="date"')}${knoepfe(fehler, 'In Rücklage legen')}</form>`);
}
function topfDialog(zeile, fehler) {
  const t = zeile ? S.k.toepfe.find((x) => x.zeile === Number(zeile)) : { name: '', sparrate: '', ab: C.heute().slice(0, 7), bis: '', anfang: 0, frei: false, bemerkung: '' };
  dialog(`<form id="k-topf" novalidate><h2 id="menu-titel">${zeile ? 'Topf ändern' : 'Neuer Topf'}</h2>
    ${feld('name', 'Name', t.name)}<div class="zwei">${feld('sparrate', 'Sparrate im Monat €', bt(t.sparrate === '' ? null : t.sparrate), 'inputmode="decimal"')}${feld('anfang', 'Anfangsbestand €', bt(t.anfang), 'inputmode="decimal"')}</div>
    <div class="zwei">${feld('ab', 'Ab (Monat)', t.ab || '', 'type="month"')}${feld('bis', 'Bis (Monat)', t.bis || '', 'type="month"')}</div>
    ${auswahl('frei', 'Zählt als frei?', ['Nein', 'Ja'], t.frei ? 'Ja' : 'Nein')}${feld('bemerkung', 'Bemerkung', t.bemerkung || '')}
    ${zeile ? `<input type="hidden" name="zeile" value="${zeile}"><p class="hint">Beenden: „Bis“ setzen statt löschen.</p>` : ''}${knoepfe(fehler)}</form>`);
}
function postenDialog(arg, fehler) {
  const [blatt, zeile] = arg.split(':');
  const liste = blatt === 'kosten' ? S.k.kosten : S.k.einnahmen;
  const x = zeile ? liste.find((y) => y.zeile === Number(zeile)) : { posten: '', bereich: '', betrag: '', rhythmus: 'monatlich', faellig: '', tag: '', ab: C.heute().slice(0, 7), bis: '', bemerkung: '', eingang: '' };
  dialog(`<form id="k-posten" novalidate><h2 id="menu-titel">${zeile ? 'Ändern' : 'Neu'}: ${blatt === 'kosten' ? 'Feste Kosten' : 'Einnahme'}</h2>
    ${feld('posten', 'Posten', x.posten)}${blatt === 'kosten' ? feld('bereich', 'Bereich', x.bereich || '') : feld('eingang', 'Eingang (wann)', x.eingang || '', 'placeholder="z. B. Ende Vormonat"')}
    <div class="zwei">${feld('betrag', 'Betrag €', bt(x.betrag === '' ? null : x.betrag), 'inputmode="decimal"')}${auswahl('rhythmus', 'Rhythmus', ['monatlich', 'jährlich'], x.rhythmus)}</div>
    <div class="zwei">${feld('faellig', 'Fällig im Monat (1–12, nur jährlich)', x.faellig || '', 'inputmode="numeric"')}${blatt === 'kosten' ? feld('tag', 'Abbuchungstag', x.tag || '', 'inputmode="numeric"') : '<span></span>'}</div>
    <div class="zwei">${feld('ab', blatt === 'kosten' ? 'Ab (Monat)' : 'Ab (gilt für)', x.ab || '', 'type="month"')}${feld('bis', blatt === 'kosten' ? 'Bis (Monat)' : 'Bis (gilt für)', x.bis || '', 'type="month"')}</div>
    ${feld('bemerkung', 'Bemerkung', x.bemerkung || '')}<input type="hidden" name="ziel" value="${h(arg)}">
    ${zeile ? '<p class="hint">Beenden: „Bis“ setzen statt löschen. Der Name wird in den Buchungen nicht umgestellt.</p>' : ''}${knoepfe(fehler)}</form>`);
}
function kursDialog(position, fehler) {
  const p = K.depot(S.k).positionen.find((x) => x.position === position);
  dialog(`<form id="k-kurs" novalidate><h2 id="menu-titel">${h(p.name || p.position)}</h2><p class="hint">${h(p.position)} · Einstand ${eur(p.einstand)}${p.letzterKurs != null ? ` · letzter TR-Kurs ${eur(p.letzterKurs)} am ${dl(p.letzterKursAm)}` : ''}</p>
    <div class="zwei">${feld('kurs', 'Kurs aktuell €', bt(p.kurs), 'inputmode="decimal"')}${feld('vom', 'vom', (typeof p.kursVom === 'string' && /^\d{4}-/.test(p.kursVom) ? p.kursVom : C.heute()), 'type="date"')}</div>
    <input type="hidden" name="zeile" value="${p.zeile}">${knoepfe(fehler)}</form>`);
}

// ---------------------------------------------------------------- Ereignisse
export async function klick(t) {
  const a = t.dataset.k;
  if (!a) return false;
  const arg = t.dataset.arg;
  switch (a) {
    case 'tab': S.tab = arg; merk('ug.ktab', arg); S.meldung = null; C.zeichne(); window.scrollTo(0, 0); return true;
    case 'neu-laden': if (C.menu.open) C.menu.close(); C.od.privatOrdnerVergessen('Kai'); S.status = 'aus'; await laden({ immer: true }); return true;
    case 'eintragen': eintragenDialog(); return true;
    case 'konto': kontoDialog(arg); return true;
    case 'topf-auf': S.topf = arg; C.zeichne(); window.scrollTo(0, 0); return true;
    case 'topf-zu': S.topf = null; C.zeichne(); window.scrollTo(0, 0); return true;
    case 'topf-buchen': eintragenDialog(null, { art: arg, topf: S.topf, posten: 'Sonstiges' }); return true;
    case 'umbuchen': { const sp = S.k.spartopf; eintragenDialog(null, { art: K.UMBUCHUNG, konto: S.k.konto.name, gegen: sp?.name, betrag: Number(arg) || '', beschreibung: 'Sparraten aufs Spartopf' }); return true; }
    case 'ruecklage': { const r = K.rechneKai(S.k, C.heute()); ruecklageDialog(r.inRuecklage || r.behalten); return true; }
    case 'vmonat': S.offen = S.offen === arg ? null : arg; if (S.tab !== 'k-vorschau') { S.tab = 'k-vorschau'; merk('ug.ktab', S.tab); } C.zeichne(); return true;
    case 'bearbeiten': { const b = S.k.buchungen.find((x) => x.nr === arg); eintragenDialog(null, { ...b, betrag: b.betrag }); return true; }
    case 'loeschen-frage': {
      const b = S.k.buchungen.find((x) => x.nr === arg);
      dialog(`<h2 id="menu-titel">${h(b.nr)} löschen?</h2><p><b>${h(b.posten || b.gegen || b.art)}</b>, ${eur(b.betrag)} am ${dl(b.datum)}</p><p class="hint">Die Zeile wird geleert (A–N), die Formeln bleiben.</p>
        <div class="knoepfe"><button class="knopf" data-k="loeschen" data-arg="${h(b.nr)}">Ja, löschen</button><button class="knopf leise" id="schliessen">Abbrechen</button></div>`);
      return true;
    }
    case 'loeschen': await schreiben(async (api) => { const z = await zeileVonNr(api, arg); await api.schreib('Buchungen', `A${z}:N${z}`, [Array(14).fill('')]); }, `${arg} gelöscht.`, { knopf: t }); return true;
    case 'soll-uebernehmen': {
      const abw = sollAbweichungen().filter((x) => !x.fehlt);
      await schreiben(async (api) => { for (const g of abw) await api.schreib('Gemeinsam', `C${S.k.gemeinsam.find((x) => x.monat === g.monat).zeile}`, [[g.neu]]); },
        `Soll für ${abw.length} Monat${abw.length > 1 ? 'e' : ''} aus der gemeinsamen Datei übernommen.`, { knopf: t });
      return true;
    }
    case 'topf': topfDialog(arg); return true;
    case 'posten': postenDialog(arg); return true;
    case 'kurs': kursDialog(arg); return true;
    case 'depot-alle': S.depotAlle = !S.depotAlle; C.zeichne(); return true;
    case 'erledigt': await schreiben((api) => api.schreib('Offene Punkte', `C${arg}`, [['erledigt']]), 'Als erledigt markiert.', { knopf: t }); return true;
    case 'c24-zu': S.c24 = null; C.zeichne(); return true;
    case 'c24-los': await c24Los(t); return true;
    case 'tr-zu': S.tr = null; C.zeichne(); return true;
    case 'tr-uebernehmen': {
      const tr = S.tr;
      await schreiben(async (api) => {
        const [von, bis] = K.BEREICHE.tr;
        const b = (await api.lies('TR Buchungen', `B${von}:B${bis}`)).values;
        let letzte = -1; b.forEach((r, i) => { if (r[0] !== '' && r[0] != null) letzte = i; });
        const start = von + letzte + 1;
        const werte = tr.neu.map((n) => [n.datum ? ser(n.datum) : '', n.typ, n.name || '', n.isin || '', n.position || '', n.stueck ?? '', n.kurs ?? '', n.betrag, n.gebuehr ?? '', n.steuer ?? '', n.beschreibung || '', n.id]);
        if (start + werte.length - 1 > bis) throw new Error('Im Blatt „TR Buchungen“ ist kein Platz mehr.');
        if (werte.length) await api.schreib('TR Buchungen', `A${start}:L${start + werte.length - 1}`, werte);
        if (tr.positionenNeu.length) {
          const [dv, db] = K.BEREICHE.depot;
          const a = (await api.lies('Depot', `A${dv}:A${db}`)).values;
          let frei = a.findIndex((r) => r[0] === '' || r[0] == null);
          for (const p of tr.positionenNeu) { if (frei < 0) break; await api.schreib('Depot', `A${dv + frei}:C${dv + frei}`, [[p.position, p.name, p.art || '']]); frei = a.findIndex((r, i) => i > frei && (r[0] === '' || r[0] == null)); }
        }
        return werte.length;
      }, (n) => `${n} TR-Zeilen übernommen.`, { knopf: t });
      S.tr = null; C.zeichne();
      return true;
    }
    default: return false;
  }
}

export async function absenden(form) {
  const x = Object.fromEntries(new FormData(form));
  const knopf = form.querySelector('button[type=submit]');
  switch (form.id) {
    case 'k-eintragen': {
      const betrag = zahl(x.betrag);
      const f = !(betrag > 0) ? 'Bitte einen Betrag größer als 0 eingeben.' : !x.datum ? 'Bitte das Datum angeben.' : !x.giltFuer ? 'Bitte „Gilt für“ angeben.'
        : x.art.includes('Topf') && !x.topf && x.art !== K.UMBUCHUNG ? 'Bei „In Topf legen“ und „Aus Topf nehmen“ bitte den Topf wählen.' : null;
      const umb = x.art === K.UMBUCHUNG;
      const f2 = f || (umb && (!x.nach || x.nach === x.konto) ? 'Bei einer Umbuchung bitte zwei verschiedene Konten wählen.' : null);
      if (f2) { eintragenDialog(f2, { ...x, gegen: umb ? x.nach : x.gegen, betrag: x.betrag }); return true; }
      const konto = x.konto || S.k.konto.name;
      const b = umb ? { datum: x.datum, giltFuer: x.giltFuer, art: x.art, konto, gegen: x.nach, betrag, posten: '', topf: '', beschreibung: x.beschreibung.trim(), kommentar: x.kommentar.trim() }
        : { datum: x.datum, giltFuer: x.giltFuer, art: x.art, konto, gegen: (x.gegen || '').trim(), betrag, posten: x.posten, topf: x.topf, beschreibung: x.beschreibung.trim(), kommentar: x.kommentar.trim() };
      if (x.nr) {
        const alt = S.k.buchungen.find((y) => y.nr === x.nr);
        await schreiben(async (api) => { const z = await zeileVonNr(api, x.nr); await api.schreib('Buchungen', `B${z}:L${z}`, [buchungZeile({ ...alt, ...b }, x.nr).slice(1, 12)]); }, `${x.nr} geändert.`, { knopf, dialogFehler: (fe) => eintragenDialog(fe, { ...x }) });
      } else {
        await schreiben(async (api) => (await buchungenAnhaengen(api, [{ ...b, quelle: 'manuell' }]))[0], (nr) => `${nr} eingetragen: ${b.posten || b.gegen || b.art}, ${eur(betrag)}.`, { knopf, dialogFehler: (fe) => eintragenDialog(fe, { ...x }) });
      }
      return true;
    }
    case 'k-ruecklage': {
      const betrag = zahl(x.betrag);
      if (!(betrag > 0)) { ruecklageDialog(betrag, 'Bitte einen Betrag größer als 0 eingeben.'); return true; }
      if (!S.k.toepfe.some((t) => t.name === K.RUECKLAGE)) { ruecklageDialog(betrag, `Den Topf „${K.RUECKLAGE}“ gibt es in der Datei nicht.`); return true; }
      await schreiben(async (api) => (await buchungenAnhaengen(api, [{ datum: x.datum, giltFuer: x.datum.slice(0, 7), art: 'In Topf legen', konto: S.k.konto.name, gegen: '', betrag, posten: 'Sonstiges', topf: K.RUECKLAGE, beschreibung: 'Rücklage für die Vorschau', quelle: 'manuell' }]))[0],
        (nr) => `${nr}: ${eur(betrag)} in „${K.RUECKLAGE}“ gelegt.`, { knopf, dialogFehler: (fe) => ruecklageDialog(betrag, fe) });
      return true;
    }
    case 'k-konto': {
      const anfang = zahl(x.anfang);
      const z = Number(x.zeile) || 5;
      if (!isFinite(anfang) || !x.stichtag) { kontoDialog(z, 'Bitte Stand und Stichtag angeben.'); return true; }
      await schreiben((api) => api.schreib('Konten', `E${z}:F${z}`, [[anfang, ser(x.stichtag)]]), 'Kontostand gespeichert.', { knopf, dialogFehler: (fe) => kontoDialog(z, fe) });
      return true;
    }
    case 'k-topf': {
      const sr = zahl(x.sparrate); const an = x.anfang ? zahl(x.anfang) : 0;
      if (!x.name.trim() || !(sr >= 0) || !isFinite(an)) { topfDialog(x.zeile, 'Bitte Name, Sparrate (0 oder mehr) und Anfangsbestand als Zahl angeben.'); return true; }
      if (x.ab && x.bis && x.bis < x.ab) { topfDialog(x.zeile, '„Bis“ liegt vor „Ab“.'); return true; }
      await schreiben(async (api) => {
        let z = Number(x.zeile);
        if (!z) {
          const [v, b] = K.BEREICHE.toepfe; const n = (await api.lies('Töpfe', `B${v}:B${b}`)).values; const i = n.findIndex((r) => r[0] === '' || r[0] == null);
          if (i < 0) throw new Error('Im Blatt „Töpfe“ ist keine Zeile mehr frei.'); z = v + i;
          const ids = (await api.lies('Töpfe', `A${v}:A${b}`)).values.map((r) => String(r[0] || '')).filter((s) => /^T\d+$/.test(s)).map((s) => Number(s.slice(1)));
          await api.schreib('Töpfe', `A${z}`, [['T' + String(Math.max(0, ...ids) + 1).padStart(2, '0')]]);
        }
        await api.schreib('Töpfe', `B${z}:F${z}`, [[x.name.trim(), sr, mSer(x.ab), mSer(x.bis), an]]);
        await api.schreib('Töpfe', `J${z}:K${z}`, [[x.frei, x.bemerkung || '']]);
      }, `Topf „${x.name.trim()}“ gespeichert.`, { knopf, dialogFehler: (fe) => topfDialog(x.zeile, fe) });
      return true;
    }
    case 'k-posten': {
      const [blatt, zeile] = x.ziel.split(':');
      const betrag = zahl(x.betrag); const fae = x.faellig ? Number(x.faellig) : null; const tag = x.tag ? Number(x.tag) : null;
      const f = !x.posten.trim() || !(betrag >= 0) ? 'Bitte Posten und Betrag angeben.' : x.rhythmus === 'jährlich' && !(fae >= 1 && fae <= 12) ? 'Bei jährlich bitte den Monat 1–12 angeben.' : tag != null && !(tag >= 1 && tag <= 31) ? 'Abbuchungstag 1–31.' : null;
      if (f) { postenDialog(x.ziel, f); return true; }
      const kosten = blatt === 'kosten';
      const B = kosten ? 'Kosten' : 'Einnahmen'; const [v, b] = K.BEREICHE[blatt];
      await schreiben(async (api) => {
        let z = Number(zeile);
        if (!z) {
          const n = (await api.lies(B, `B${v}:B${b}`)).values; const i = n.findIndex((r) => r[0] === '' || r[0] == null);
          if (i < 0) throw new Error(`Im Blatt „${B}“ ist keine Zeile mehr frei.`); z = v + i;
          const ids = (await api.lies(B, `A${v}:A${b}`)).values.map((r) => String(r[0] || '')).filter((s) => /^[KE]\d+$/.test(s)).map((s) => Number(s.slice(1)));
          await api.schreib(B, `A${z}`, [[(kosten ? 'K' : 'E') + String(Math.max(0, ...ids) + 1).padStart(2, '0')]]);
        }
        if (kosten) {
          await api.schreib(B, `B${z}:J${z}`, [[x.posten.trim(), x.bereich || '', betrag, x.rhythmus, fae ?? '', tag ?? '', S.k.konto.name, mSer(x.ab), mSer(x.bis)]]);
          await api.schreib(B, `N${z}`, [[x.bemerkung || '']]);
        } else {
          await api.schreib(B, `B${z}:I${z}`, [[x.posten.trim(), betrag, x.rhythmus, fae ?? '', x.eingang || '', S.k.konto.name, mSer(x.ab), mSer(x.bis)]]);
          await api.schreib(B, `M${z}`, [[x.bemerkung || '']]);
        }
      }, `„${x.posten.trim()}“ gespeichert.`, { knopf, dialogFehler: (fe) => postenDialog(x.ziel, fe) });
      return true;
    }
    case 'k-kurs': {
      const kurs = x.kurs ? zahl(x.kurs) : '';
      if (x.kurs && !(kurs > 0)) { kursDialog(K.depot(S.k).positionen.find((p) => p.zeile === Number(x.zeile)).position, 'Bitte den Kurs als Zahl angeben.'); return true; }
      await schreiben((api) => api.schreib('Depot', `L${x.zeile}:M${x.zeile}`, [[kurs, kurs === '' ? '' : ser(x.vom || C.heute())]]), 'Kurs gespeichert.', { knopf });
      return true;
    }
    case 'k-c24-form': {
      const a = S.c24;
      const neu = []; const ergaenzen = [];
      a.zeilen.forEach((r, i) => {
        if (r.status === 'importiert' || r.status === 'gegenbuchung' || !x['w' + i]) return;
        if (r.status === 'ergaenzen') ergaenzen.push({ nr: r.treffer.nr, id: r.z.id });
        else neu.push({ ...r.vorschlag, posten: r.vorschlag.art === K.UMBUCHUNG ? null : x['p' + i] || null, giltFuer: x['g' + i] || r.vorschlag.giltFuer });
      });
      if (!neu.length && !ergaenzen.length) return true;
      dialog(`<h2 id="menu-titel">Übernehmen?</h2><p>${neu.length} neue Buchungen eintragen${ergaenzen.length ? ` und bei ${ergaenzen.length} von Hand eingetragenen die Import-ID ergänzen` : ''}.</p>
        <div class="knoepfe"><button class="knopf" data-k="c24-los">Ja, übernehmen</button><button class="knopf leise" id="schliessen">Abbrechen</button></div>`);
      S.c24.plan = { neu, ergaenzen };
      return true;
    }
    default: return false;
  }
}
export async function c24Los(t) {
  const { neu, ergaenzen } = S.c24.plan;
  const ok = await schreiben(async (api) => {
    const nummern = await buchungenAnhaengen(api, neu);
    for (const e of ergaenzen) { const z = await zeileVonNr(api, e.nr); await api.schreib('Buchungen', `N${z}`, [[e.id]]); }
    return nummern.length;
  }, (n) => `${n} Buchungen übernommen${ergaenzen.length ? `, ${ergaenzen.length} Import-IDs ergänzt` : ''}.`, { knopf: t });
  if (ok) S.c24 = null;
  C.zeichne();
}

export async function dateiGewaehlt(id, datei) {
  const bytes = new Uint8Array(await datei.arrayBuffer());
  const text = () => { let s = new TextDecoder('utf-8').decode(bytes); if (s.includes('�')) s = new TextDecoder('windows-1252').decode(bytes); return s; };
  if (id === 'k-c24') {
    try {
      const gk = C.z.modell?.konten.find((k) => k.rolle === 'gemeinsam' && /giro/i.test(k.art || ''));
      const gkEndung = (String(gk?.ibanEnde || '').match(/(\d{4})\D*$/) || [])[1];
      S.c24 = { zeilen: await K.abgleichC24(K.leseC24(text()), S.k, { gkEndung }) };
    } catch (e) { S.c24 = { fehler: e.message }; }
  } else if (id === 'k-tr') {
    try {
      const roh = /\.xlsx$/i.test(datei.name) ? K.leseTr(await K.trTextAusXlsx(bytes)) : K.leseTr(text());
      S.tr = K.trNeu(roh, S.k);
    } catch (e) { S.tr = { fehler: e.message }; }
  }
  C.zeichne();
}
export function eingabe(el) {
  if (el.id === 'k-suche') { S.suche = el.value; const pos = el.selectionStart; C.zeichne(); const n = document.getElementById('k-suche'); if (n) { n.focus(); n.setSelectionRange(pos, pos); } return true; }
  if (el.id === 'k-fmonat') { S.fMonat = el.value; C.zeichne(); return true; }
  if (el.id === 'k-fart') { S.fArt = el.value; C.zeichne(); return true; }
  return false;
}
