// Kais privater Bereich („Mein Geld“) in „Unser Geld“: liest und schreibt Kai_Privat.xlsx (Ordner Apps/Kai, nicht geteilt).
// Aus der gemeinsamen Datei wird nur gelesen (Kais Soll aufs Gemeinschaftskonto). Nichts fließt zurück.
import * as K from './kern/kai.js';
import { eur, monatsname } from './kern/rechnen.js';
import { betragFinden, datumFinden } from './kern/verstehen.js';

let C;
const S = { status: 'aus', fehler: null, ort: null, item: null, k: null, tab: null, meldung: null, laedt: false,
  suche: '', fMonat: '', fArt: '', offen: null, c24: null, tr: null, gelesen: null, topf: null, fWofuer: '', wZeit: 'monat' };
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
  if (S.schreibt) return false;
  if (knopf) { knopf.disabled = true; knopf.dataset.t = knopf.textContent; knopf.textContent = 'Wird gespeichert …'; }
  S.schreibt = true;
  try {
    const vorher = S.item?.eTag;
    const r = await C.od.mitSitzung(S.ort, arbeit);
    S.schreibt = false;
    if (C.menu.open) C.menu.close();
    const text = typeof meldung === 'function' ? meldung(r) : meldung;
    // OneDrive liefert die Datei nach dem Schreiben oft noch kurz im alten Stand: erst neu laden, wenn die neue Version da ist.
    S.meldung = { art: 'info', text: text + ' Anzeige wird aktualisiert …' }; C.zeichne();
    const neu = await C.od.aufNeueVersionWarten(S.ort, vorher).catch(() => null);
    S.meldung = { art: neu ? 'info' : 'warn', text: neu ? text : text + ' OneDrive zeigt die Änderung noch nicht an. Bitte gleich über das Menü neu laden.' };
    await laden({ immer: true });
    return true;
  } catch (e) {
    S.schreibt = false;
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
const WL = () => K.wofuerListe(S.k);
/** Budget als „Wofür“ gewählt → Posten = Budget (zählt dann in Kosten/Vorschau als Ist). */
const budgetPosten = (b) => { if (b.wofuer && K.istBudget(S.k, b.wofuer) && (!b.posten || b.posten === 'Sonstiges')) b.posten = b.wofuer; return b; };
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
  if (liste.some((b) => b.wofuer)) { await wofuerKopf(api); await api.schreib('Buchungen', `${W}${start}:${W}${start + liste.length - 1}`, liste.map((b) => [b.wofuer || ''])); }
  const k = await api.lies('Buchungen', `A${start}:A${start + werte.length - 1}`);
  if (!k.values.every((r, i) => r[0] === nummern[i])) throw new Error('Die Buchungen wurden nicht wie erwartet gespeichert.');
  return nummern;
}
const W = K.WOFUER_SPALTE;
/** Überschrift „Wofür“ in Buchungen!Q4 setzen, falls sie noch fehlt (neue Spalte, ändert keine Formel). */
async function wofuerKopf(api) {
  const kopf = (await api.lies('Buchungen', `${W}4`)).values?.[0]?.[0];
  if (!kopf) await api.schreib('Buchungen', `${W}4`, [['Wofür']]);
  else if (kopf !== 'Wofür') throw new Error(`In Buchungen!${W}4 steht „${kopf}“ statt „Wofür“. Bitte in Excel nachsehen.`);
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
  const f = { 'k-ausgaben': ausgabenAnsicht, 'k-uebersicht': uebersicht, 'k-vorschau': vorschau, 'k-buchungen': buchungen, 'k-toepfe': toepfeAnsicht, 'k-depot': depotAnsicht, 'k-mehr': mehr }[S.tab] || uebersicht;
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
    ${budgetDrueberText()}
    ${r.inRuecklage > 0 ? `<button class="knopf hell" data-k="ruecklage">${eur(r.inRuecklage)} in Rücklage legen</button>` : ''}</section>` : '';
  const abw = sollAbweichungen();
  const sollHinweis = abw.length ? `<div class="hinweis warn">Das Soll aufs Gemeinschaftskonto hat sich in der gemeinsamen Datei geändert (${abw.length > 3 ? `${abw.length} Monate, ab ${mn(abw[0].monat)}` : abw.map((a) => mn(a.monat, false)).join(', ')}). <button class="link" data-k="soll-uebernehmen">Übernehmen</button></div>` : '';
  const zeilen = r.vorschau.filter((v) => v.imHorizont);
  return `${kontostaende(r)}${warn}${sollHinweis}
    <div class="k-hero"><div class="k-hero-oben"><span>Frei</span><span class="k-ampel ${luecke ? 'rot' : 'gruen'}">${luecke ? 'Lücke' : '✓ Alles gedeckt'}</span></div>
      <div class="k-gross ${r.frei < 0 ? 'neg' : ''}">${eur(r.frei)}</div>
      <div class="k-hero-zeilen"><span>Kontostände <b>${eur(Math.round((r.stand + (r.spartopf?.stand || 0)) * 100) / 100)}</b></span><span>verplant in Töpfen <b>${eur(r.verplant)}</b></span><span>Rücklage <b>${eur(r.ruecklage)}</b></span></div></div>
    <section><h2>Empfehlung für ${h(mn(r.akt, false))}</h2><div class="rechnung">
      <div class="rz"><span>Bis Monatsende behalten<small>damit die nächsten Monate gedeckt sind</small></span><b>${eur(r.behalten)}</b></div>
      <div class="rz"><span>Davon noch in die Rücklage<small>Topf „${h(K.RUECKLAGE)}“, Stand ${eur(r.ruecklage)}</small></span><b>${eur(r.inRuecklage)}</b></div>
      <div class="rz summe"><span>Frei für Variables diesen Monat</span><b class="${r.variabel < 0 ? 'neg' : ''}">${eur(r.variabel)}</b></div></div>
      ${r.inRuecklage > 0 && !luecke ? `<button class="knopf leise" style="margin-top:10px" data-k="ruecklage">${eur(r.inRuecklage)} in Rücklage legen</button>` : ''}</section>
    <section><h2>Die nächsten ${r.horizont} Monate</h2><div class="k-tabelle"><div class="k-tz kopf"><span>Monat</span><span>Frei am Ende</span><span>Behalten</span><span>Variables</span></div>
      ${zeilen.map((v) => `<button class="k-tz ${v.status === 'Lücke' ? 'luecke' : ''}" data-k="vmonat" data-arg="${v.monat}"><span>${h(mn(v.monat, false))}${v.status === 'Lücke' ? ' ⚠' : ''}</span><span class="${v.freiEnde < 0 ? 'neg' : ''}">${eur(v.freiEnde)}</span><span>${eur(v.behalten)}</span><span class="${v.variabel < 0 ? 'neg' : ''}">${eur(v.variabel)}</span></button>`).join('')}</div>
      <p class="hint" style="margin-top:6px">Ohne variable Ausgaben gerechnet. Antippen zeigt die Posten.</p></section>
    ${wofuerHtml()}
    ${planIstHtml()}
    ${spartopfHinweis(r)}
    <p class="hint">Stichtag ${dl(r.stichtag)} · letzte Buchung ${dl(r.letzteBuchung) || '–'} · geladen ${S.gelesen ? new Date(S.gelesen).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '–'}</p>`;
}
/** Name wie in der c24-App: „c24 Smartkonto“ und „Spartopf“. */
const anzeigeName = (name) => (name === S.k.konto.name ? 'c24 Smartkonto' : String(name).replace(/^c24\s+/i, ''));
/** Große Karte ganz oben: echte Kontostände (berechnet) und der letzte Vergleich mit der Bank-App. */
function kontostaende(r) {
  const konten = K.eigeneKonten(S.k).map((kt) => ({ kt, stand: K.standVon(S.k, kt) }));
  const p = hol('ug.kpruef');
  const gleich = p && konten.every(({ kt, stand }) => { const x = p.konten.find((y) => y.name === kt.name); return !x || Math.abs((x.stand ?? NaN) - (stand ?? NaN)) < 0.005; });
  const status = !p ? '<span class="k-ampel">noch nicht verglichen</span>'
    : !gleich ? '<span class="k-ampel">seit dem Vergleich geändert</span>'
      : p.ok ? '<span class="k-ampel gruen">✓ stimmt mit der Bank</span>' : '<span class="k-ampel rot">weicht von der Bank ab</span>';
  return `<section class="k-konten" aria-label="Aktuelle Kontostände"><div class="k-spar-kopf"><span data-hilfe-schluessel="Aktuelle Kontostände">Aktuelle Kontostände</span>${status}</div>
    ${konten.map(({ kt, stand }) => `<div class="k-konto-zeile"><span>${h(anzeigeName(kt.name))}</span><b class="${stand < 0 ? 'neg' : ''}">${stand == null ? 'Stand fehlt' : eur(stand)}</b></div>`).join('')}
    ${konten.length > 1 ? `<div class="k-konto-zeile summe"><span>Zusammen</span><b>${eur(Math.round(konten.reduce((a, x) => a + (x.stand || 0), 0) * 100) / 100)}</b></div>` : ''}
    <p class="k-konten-fuss">Berechnet aus deinen Buchungen bis ${dl(r.letzteBuchung) || '–'}${p ? ` · zuletzt verglichen ${new Date(p.zeit).toLocaleDateString('de-DE', { day: 'numeric', month: 'numeric' })} ${new Date(p.zeit).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })}` : ''}</p>
    <button class="knopf hell" data-k="pruefen">Mit Bank-App vergleichen</button></section>`;
}
/** „Über Budget: Dampfen +30,00 €“ für die Warnung (nur laufender Monat). */
function budgetDrueberText() {
  const akt = C.heute().slice(0, 7);
  const d = K.budgets(S.k).map((x) => K.budgetStand(S.k, x, akt, C.heute())).filter((b) => b.drueber > 0);
  return d.length ? `<p><span class="k-warn-budget">Über Budget:</span> ${d.map((b) => `${h(b.name)} +${eur(b.drueber)}`).join(', ')}. <button class="link hell" data-k="ausgaben">Ansehen</button></p>` : '';
}
/** Ring für ein Budget: äußerer Ring = verbraucht vom Plan, innerer roter Ring = was darüber hinausgeht, Strich = heute. */
function budgetRing(b, gross = true) {
  const g = gross ? 136 : 64, sw = gross ? 13 : 7, r1 = (g - sw) / 2, r2 = r1 - sw - (gross ? 4 : 2), c = g / 2;
  const U1 = 2 * Math.PI * r1, U2 = 2 * Math.PI * r2;
  const teil = b.plan > 0 ? Math.min(1, b.verbraucht / b.plan) : (b.verbraucht > 0 ? 1 : 0);
  const ueber = b.plan > 0 ? Math.min(1, b.drueber / b.plan) : 0;
  const winkel = b.anteil * 2 * Math.PI - Math.PI / 2;
  const strich = b.anteil > 0 && b.anteil < 1 ? `<line x1="${c + (r1 - sw / 2 - 2) * Math.cos(winkel)}" y1="${c + (r1 - sw / 2 - 2) * Math.sin(winkel)}" x2="${c + (r1 + sw / 2 + 2) * Math.cos(winkel)}" y2="${c + (r1 + sw / 2 + 2) * Math.sin(winkel)}" class="k-ring-heute"><title>bis heute vorgesehen: ${eur(b.sollBisHeute)}</title></line>` : '';
  return `<svg class="k-ring ${b.drueber > 0 ? 'drueber' : ''}" width="${g}" height="${g}" viewBox="0 0 ${g} ${g}" role="img" aria-label="${h(b.name)}: ${eur(b.verbraucht)} von ${eur(b.plan)} verbraucht">
    <circle cx="${c}" cy="${c}" r="${r1}" class="k-ring-spur" stroke-width="${sw}" fill="none"/>
    <circle cx="${c}" cy="${c}" r="${r1}" class="k-ring-wert" stroke-width="${sw}" fill="none" stroke-linecap="round" stroke-dasharray="${(teil * U1).toFixed(2)} ${U1.toFixed(2)}" transform="rotate(-90 ${c} ${c})"><title>verbraucht ${eur(b.verbraucht)} von ${eur(b.plan)}</title></circle>
    ${ueber > 0 ? `<circle cx="${c}" cy="${c}" r="${r2}" class="k-ring-ueber" stroke-width="${sw}" fill="none" stroke-linecap="round" stroke-dasharray="${(ueber * U2).toFixed(2)} ${U2.toFixed(2)}" transform="rotate(-90 ${c} ${c})"><title>${eur(b.drueber)} über Budget</title></circle>` : ''}
    ${strich}
    ${gross ? `<text x="${c}" y="${c - 2}" text-anchor="middle" class="k-ring-zahl">${h(eur(b.verbraucht).replace(/,00 €$/, ' €'))}</text><text x="${c}" y="${c + 17}" text-anchor="middle" class="k-ring-von">von ${h(eur(b.plan).replace(/,00 €$/, ' €'))}</text>` : ''}</svg>`;
}
function budgetText(b, lauf) {
  const restTage = lauf ? (() => { const h0 = C.heute(); const t = new Date(Date.UTC(+h0.slice(0, 4), +h0.slice(5, 7), 0)).getUTCDate(); return t - Number(h0.slice(8, 10)); })() : 0;
  if (b.drueber > 0) return `<span class="k-b-status rot">⚠ ${eur(b.drueber)} über Budget</span>`;
  if (lauf && b.verbraucht > b.sollBisHeute + 0.005) return `<span class="k-b-status gelb">↗ schneller als geplant</span><small>Bis heute wären ${eur(b.sollBisHeute)} dran. Noch ${eur(b.rest)} für ${restTage} Tag${restTage === 1 ? '' : 'e'}.</small>`;
  return `<span class="k-b-status gruen">✓ im Plan</span><small>Noch ${eur(b.rest)} übrig${lauf ? ` für ${restTage} Tag${restTage === 1 ? '' : 'e'}` : ''}.</small>`;
}
/** Eigene Seite „Ausgaben“: Budgets als Ringe, Wofür mit Referenzwert, Verlauf 6 Monate. */
function ausgabenAnsicht() {
  const k = S.k; const heute = C.heute(); const akt = heute.slice(0, 7);
  const m = S.aMonat || akt; const lauf = m === akt;
  const bud = K.budgets(k).map((x) => K.budgetStand(k, x, m, heute)).filter((b) => b.plan > 0 || b.verbraucht > 0);
  const a = K.wofuerAuswertung(k, m, m);
  const ref = K.referenzWerte(k, m);
  const max = Math.max(1, ...a.reihen.map((x) => Math.max(x.summe, ref.werte.get(x.name) || 0)));
  const monate6 = Array.from({ length: 6 }, (_, i) => K.monatPlus(m, i - 5));
  const v = K.wofuerVerlauf(k, monate6);
  const reihen = v.gruppen.filter((g) => g.summe > 0).sort((x, y) => (K.WOFUER_EXTRA.indexOf(x.name) - K.WOFUER_EXTRA.indexOf(y.name)) || y.summe - x.summe);
  const ersterMonat = [...k.buchungen.map((b) => b.giltFuer).filter(Boolean)].sort()[0] || akt;
  const kopf = `<div class="k-topfkopf"><button class="k-zurueck" data-k="a-zurueck" aria-label="Zurück"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5l-7 7 7 7"/></svg></button><h2>Ausgaben</h2></div>
    <div class="k-amonat"><button class="knopf leise klein" data-k="a-monat" data-arg="${K.monatPlus(m, -1)}" ${K.monatPlus(m, -1) < ersterMonat ? 'disabled' : ''} aria-label="Vorheriger Monat">‹</button><b>${h(mn(m))}${lauf ? ' <small>läuft</small>' : ''}</b><button class="knopf leise klein" data-k="a-monat" data-arg="${K.monatPlus(m, 1)}" ${m >= akt ? 'disabled' : ''} aria-label="Nächster Monat">›</button></div>`;
  const budHtml = bud.length ? `<section><h2>Budgets</h2>${bud.map((b) => `<div class="k-budget">${budgetRing(b)}<div class="k-b-info"><b class="k-b-name">${h(b.name)}</b>${budgetText(b, lauf)}
      <button class="link" data-k="wofuer-filter" data-arg="${h(b.name)}">${b.anzahl} Buchung${b.anzahl === 1 ? '' : 'en'} ansehen</button></div></div>`).join('')}
    <p class="hint">Blauer Ring: verbraucht vom Budget (voll = aufgebraucht). Roter innerer Ring: was darüber hinausgeht. Der Strich zeigt, wie viel bis heute dran wäre, wenn du gleichmäßig verbrauchst.</p></section>`
    : `<section><h2>Budgets</h2><div class="leer">Noch kein Monatsbudget. Ein Topf, an den du jeden Monat gehst, lässt sich auf seiner Seite unter „Töpfe“ in ein Budget umwandeln.</div></section>`;
  const wofHtml = `<section><h2>Wofür ausgegeben</h2>${a.reihen.length ? `<div class="k-wofuer">${a.reihen.map((x) => { const rw = ref.werte.get(x.name);
    const bx = K.budgets(k).find((y) => y.posten === x.name);
    return `<button class="k-w-zeile ${x.name === 'Nicht zugeordnet' ? 'leise' : ''}" data-k="wofuer-filter" data-arg="${h(x.name)}"><span class="k-w-name">${h(x.name)} <small>${x.anzahl}×${bx ? ` · Budget ${eur(K.plan(bx, m))}` : rw != null ? ` · Ø ${eur(rw)}` : ''}</small></span><b>${eur(x.summe)}</b>
      <span class="k-w-balken mit-ref"><i style="width:${Math.round((x.summe / max) * 100)}%"></i>${rw != null ? `<em class="k-w-ref" style="left:${Math.min(100, Math.round((rw / max) * 100))}%" title="Ø ${eur(rw)}"></em>` : ''}</span></button>`; }).join('')}
      <div class="k-w-zeile summe"><span class="k-w-name">Ausgaben gesamt</span><b>${eur(a.gesamt)}</b></div></div>
      <p class="hint" style="margin-top:6px">${ref.werte.size ? `Strich und „Ø“ = Durchschnitt aus ${ref.monate.map((x) => mn(x, false)).join(', ')}. <b>Nur ein Referenzwert, kein Plan</b> – damit du siehst, ob ein Monat aus der Reihe fällt.` : 'Sobald es zwei Monate mit „Wofür“ gibt, zeigt die App hier einen Durchschnitt als Referenz.'}</p>`
    : '<div class="leer">In diesem Monat gibt es noch keine Ausgaben.</div>'}
    ${a.ohne ? `<p class="hint">${a.ohne} Ausgabe${a.ohne > 1 ? 'n' : ''} ohne „Wofür“. <button class="link" data-k="wofuer-alle">Jetzt zuordnen</button></p>` : ''}</section>`;
  const vmax = (g) => Math.max(1, ...g.werte);
  const verlauf = reihen.length ? `<section><h2>Verlauf</h2><div class="k-verlauf">
      <div class="k-vl kopf"><span></span><span class="k-vl-bars">${monate6.map((x) => `<i>${h(mn(x, false).slice(0, 3))}</i>`).join('')}</span><span>Ø</span></div>
      ${reihen.map((g) => { const avg = g.werte.filter((w) => w > 0).length ? g.summe / Math.max(1, g.werte.filter((w) => w > 0).length) : 0;
        return `<div class="k-vl"><span class="k-vl-name">${h(g.name)}</span><span class="k-vl-bars">${g.werte.map((w, i) => `<i class="${monate6[i] === m ? 'akt' : ''}" style="--h:${Math.round((w / vmax(g)) * 100)}%" title="${h(mn(monate6[i], false))}: ${eur(w)}"></i>`).join('')}</span><span class="k-vl-avg">${eur(Math.round(avg * 100) / 100).replace(/,\d\d €$/, ' €')}</span></div>`; }).join('')}
    </div><p class="hint" style="margin-top:6px">Je Zeile eigener Maßstab, der hervorgehobene Balken ist ${h(mn(m, false))}. Ø = Durchschnitt der Monate mit Ausgaben in dieser Gruppe.</p></section>` : '';
  return kopf + budHtml + wofHtml + verlauf;
}
/** Auswertung „Wofür“ für den aktuellen Monat, 3 oder 12 Monate. */
function wofuerHtml() {
  const akt = C.heute().slice(0, 7);
  const von = S.wZeit === '12' ? K.monatPlus(akt, -11) : S.wZeit === '3' ? K.monatPlus(akt, -2) : akt;
  const a = K.wofuerAuswertung(S.k, von, akt);
  const max = Math.max(1, ...a.reihen.map((x) => x.summe));
  const wahl = [['monat', mn(akt, false)], ['3', '3 Monate'], ['12', '12 Monate']];
  const budK = K.budgets(S.k).map((x) => K.budgetStand(S.k, x, akt, C.heute())).filter((b) => b.plan > 0);
  return `<section><h2>Wofür ausgegeben</h2>
    ${budK.map((b) => `<button class="k-budget klein" data-k="ausgaben">${budgetRing(b, false)}<div class="k-b-info"><b class="k-b-name">${h(b.name)} <span>${eur(b.verbraucht)} von ${eur(b.plan)}</span></b>${budgetText(b, true)}</div></button>`).join('')}
    <div class="k-wzeit" role="group">${wahl.map(([w, t]) => `<button class="${S.wZeit === w ? 'an' : ''}" data-k="wzeit" data-arg="${w}">${h(t)}</button>`).join('')}</div>
    ${a.reihen.length ? `<div class="k-wofuer">${a.reihen.map((x) => `<button class="k-w-zeile ${x.name === 'Nicht zugeordnet' ? 'leise' : ''}" data-k="wofuer-filter" data-arg="${h(x.name)}"><span class="k-w-name">${h(x.name)} <small>${x.anzahl}×</small></span><b>${eur(x.summe)}</b><span class="k-w-balken"><i style="width:${Math.round((x.summe / max) * 100)}%"></i></span></button>`).join('')}
      <div class="k-w-zeile summe"><span class="k-w-name">Ausgaben gesamt</span><b>${eur(a.gesamt)}</b></div></div>`
    : '<div class="leer">In diesem Zeitraum gibt es noch keine Ausgaben.</div>'}
    ${a.ohne ? `<p class="hint" style="margin-top:8px">${a.ohne} Ausgabe${a.ohne > 1 ? 'n' : ''} ohne „Wofür“. <button class="link" data-k="wofuer-alle">Vorschläge für alle Buchungen ansehen</button></p>` : ''}
    ${a.nichtGezaehlt.length ? `<p class="hint">Nicht gezählt, weil es Geld zwischen deinen Konten ist: ${a.nichtGezaehlt.map((b) => `<button class="link" data-k="pruef-ansehen" data-arg="${h(b.nr)}">${h(b.nr)}</button> (${eur(b.betrag)})`).join(', ')}. Am besten als „Umbuchung“ ändern.</p>` : ''}
    <button class="knopf leise" style="width:100%;margin-top:10px" data-k="ausgaben">Ausgaben-Übersicht mit Verlauf öffnen</button>
    <p class="hint">Gezählt nach „Gilt für“. Ohne „Wofür“ wird nach Posten sortiert („Gemeinschaftskonto“, „Trade Republic“, „Feste Posten“). Umbuchungen und Topf-Arten zählen nicht.</p></section>`;
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
  const wf = (b) => (b.art === 'Ausgabe' ? K.wofuerGruppe(b, k) : b.wofuer || '');
  const liste = k.buchungen.filter((b) => (!S.fMonat || b.giltFuer === S.fMonat) && (!S.fArt || b.art === S.fArt) && (!S.fWofuer || wf(b) === S.fWofuer)
    && (!q || [b.posten, b.gegen, b.beschreibung, b.kommentar, b.kategorie, b.nr, b.topf, b.wofuer].some((x) => (x || '').toLowerCase().includes(q))))
    .sort((a, b) => (b.datum || '').localeCompare(a.datum || '') || b.nr.localeCompare(a.nr));
  const monate = [...new Set(k.buchungen.map((b) => b.giltFuer).filter(Boolean))].sort().reverse();
  return `<section><h2>Buchungen</h2>
    <div class="k-filter"><input type="search" id="k-suche" placeholder="Suchen …" value="${h(S.suche)}" autocomplete="off">
      <select id="k-fmonat"><option value="">alle Monate</option>${monate.map((m) => `<option value="${m}" ${m === S.fMonat ? 'selected' : ''}>${h(mn(m))}</option>`).join('')}</select>
      <select id="k-fart"><option value="">alle Arten</option>${K.ARTEN.map((a) => `<option ${a === S.fArt ? 'selected' : ''}>${a}</option>`).join('')}</select>
      <select id="k-fwofuer"><option value="">Wofür: alle</option>${[...WL(), ...K.WOFUER_EXTRA].map((a) => `<option ${a === S.fWofuer ? 'selected' : ''}>${a}</option>`).join('')}</select></div>
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
  return `<div class="bu-wrap"><button class="bu" aria-expanded="false"><span class="t">${umb ? `${h(b.konto)} → ${h(b.gegen || '?')}` : h(b.posten || b.gegen || b.beschreibung || b.art)}${b.topf && !topfName && !umb ? ` <span class="chip aus-topf">${h(b.topf)}</span>` : ''}${fremd ? ` <span class="chip">${h(b.konto)}</span>` : ''}${b.wofuer ? ` <span class="chip wofuer">${h(b.wofuer)}</span>` : ''}${K.istTeil(b) ? ` <span class="chip wofuer">${h(b.kommentar.match(/^Teil von P\d+/)[0])}</span>` : ''}</span>
      <span class="betrag ${(topfName ? wirk > 0 : b.art === 'Einnahme') ? 'plus' : ''}">${zeichen}${eur(b.betrag)}</span>
      <span class="r">${dk(b.datum)} · ${h(b.art)}${b.giltFuer && b.giltFuer !== b.datum?.slice(0, 7) ? ` <span class="okt">gilt für ${h(mn(b.giltFuer, false))}</span>` : ''}${b.gegen && b.posten && !umb ? ' · ' + h(b.gegen) : ''}</span>
      <span class="more"><dl>${[['Nummer', b.nr], ['Konto', umb ? `${b.konto} → ${b.gegen}` : b.konto], ['Datum', dl(b.datum)], ['Gilt für', mn(b.giltFuer)], ['Kategorie (Bank)', b.kategorie], ['Wofür', b.wofuer], ['Topf', b.topf], ['Beschreibung', b.beschreibung], ['Kommentar', b.kommentar], ['Quelle', b.quelle]].filter(([, w]) => w).map(([a, w]) => `<dt>${a}</dt><dd>${h(w)}</dd>`).join('')}</dl></span></button>
      <div class="bu-aktion"><button class="knopf leise klein" data-k="bearbeiten" data-arg="${h(b.nr)}">Bearbeiten</button>${b.art === 'Ausgabe' && !K.istTeil(b) ? `<button class="knopf leise klein" data-k="aufteilen" data-arg="${h(b.nr)}">Aufteilen</button>` : ''}<button class="knopf leise klein" data-k="loeschen-frage" data-arg="${h(b.nr)}">Löschen</button></div></div>`;
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
    ${!t.frei ? `<p class="hint" style="margin:-4px 0 14px">Gehst du jeden Monat an diesen Topf (z. B. für etwas, das du laufend kaufst)? Dann passt ein Monatsbudget besser. <button class="link" data-k="topf-budget" data-arg="${h(t.name)}">In Monatsbudget umwandeln</button></p>` : ''}
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
        <label class="feld"><span>Gilt für</span><input type="month" name="g${i}" value="${x.vorschlag.giltFuer}"></label></div>
        ${x.vorschlag.art === 'Ausgabe' ? `<label class="feld"><span>Wofür</span><select name="w2${i}">${['', ...WL()].map((w) => `<option ${w === (x.vorschlag.wofuer || '') ? 'selected' : ''} value="${h(w)}">${h(w || '– nicht zugeordnet –')}</option>`).join('')}</select></label>` : ''}`}</div>`).join('')}</div>
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
    ${auswahl('wofuer', 'Wofür', WL(), v.wofuer, '– nicht zugeordnet –')}
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
  // „Wofür“ vorschlagen, solange es noch niemand gewählt hat
  let wofuerVonHand = Boolean(v.wofuer);
  const wofuerRaten = () => { if (wofuerVonHand || !f.wofuer) return; const x = Object.fromEntries(new FormData(f)); f.wofuer.value = K.wofuerVorschlag({ art: x.art, gegen: x.gegen, beschreibung: x.beschreibung, posten: x.posten }, S.k) || ''; };
  f.addEventListener('change', (e) => { if (e.target.name === 'wofuer') { wofuerVonHand = true; if (K.istBudget(S.k, e.target.value) && f.posten && [...f.posten.options].some((o) => o.value === e.target.value) && (!f.posten.value || f.posten.value === 'Sonstiges')) f.posten.value = e.target.value; } else if (['gegen', 'beschreibung', 'art', 'posten'].includes(e.target.name)) wofuerRaten(); });
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
      wofuerRaten();
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

function pruefDialog(fehler, werte = {}) {
  const konten = K.eigeneKonten(S.k);
  dialog(`<form id="k-pruefen" novalidate><h2 id="menu-titel">Mit Bank-App vergleichen</h2>
    <p class="hint">Trag die Stände ein, die die c24-App gerade anzeigt. Die App vergleicht sie mit dem, was sie aus deinen Buchungen rechnet. Es wird nichts gespeichert, außer dem Ergebnis auf diesem Gerät.</p>
    ${konten.map((kt, i) => feld('b' + i, `${h(anzeigeName(kt.name))} laut Bank in €`, werte['b' + i] ?? '', 'inputmode="decimal" placeholder="0,00"')).join('')}
    ${knoepfe(fehler, 'Vergleichen')}</form>`);
}
function pruefErgebnis(e) {
  const zeile = (x) => `<div class="k-pr-konto ${x.ok ? 'ok' : 'ab'}"><div class="k-pr-kopf"><b>${h(anzeigeName(x.name))}</b>${x.ok ? '<span class="k-ampel gruen">✓ stimmt</span>' : `<span class="k-ampel rot">${x.diff > 0 ? '+' : '−'}${eur(Math.abs(x.diff))}</span>`}</div>
    <div class="rechnung"><div class="rz"><span>laut Bank</span><b>${eur(x.bank)}</b></div><div class="rz"><span>laut App<small>Stichtag ${dl(x.stichtag)} + Buchungen danach</small></span><b>${x.stand == null ? 'fehlt' : eur(x.stand)}</b></div></div>
    ${x.verdacht.length ? `<p class="hint" style="margin:8px 0 4px">${h(x.verdacht.find((v) => v.hinweis)?.hinweis || '')}</p>
      ${x.verdacht.filter((v) => v.nr).length ? `<p style="margin:6px 0 4px"><b>Diese Buchungen solltest du dir ansehen:</b></p><div class="k-pr-liste">${x.verdacht.filter((v) => v.nr).slice(0, 12).map((v) => `<button class="k-pr-b" data-k="pruef-ansehen" data-arg="${h(v.nr)}"><span>${h(v.nr)} · ${dk(v.datum)} · ${h(v.posten || v.gegen || v.art)}<small>${h(v.grund)}</small></span><b>${vz(v.wirkung)}</b></button>`).join('')}</div>`
        : '<p class="hint">Keine Buchung passt genau zur Differenz. Vergleich am besten die letzten Tage mit der Umsatzliste in der c24-App.</p>'}` : ''}</div>`;
  const falsch = e.falsch.length ? `<div class="hinweis info" style="margin-top:12px"><b>Nicht als Umbuchung eingetragen:</b> ${e.falsch.map((b) => `<button class="link" data-k="pruef-ansehen" data-arg="${h(b.nr)}">${h(b.nr)} (${eur(b.betrag)} ${h(b.konto)} → ${h(b.gegen)})</button>`).join(', ')}. Das sieht nach Geld zwischen deinen eigenen Konten aus. Der Kontostand stimmt trotzdem, aber in „Wofür ausgegeben“ zählt es als Ausgabe. Über „Bearbeiten“ als Art „Umbuchung“ eintragen.</div>` : '';
  dialog(`<h2 id="menu-titel">${e.ok ? '✓ Alles stimmt' : 'Abweichung gefunden'}</h2>${e.konten.map(zeile).join('')}${falsch}
    <div class="knoepfe"><button class="knopf" id="schliessen">Fertig</button></div>`);
}
/** „Wofür“ nachtragen: Liste mit Vorschlägen, Buchungen anhaken → „Ausgewählte ändern“ → zurück zur neu sortierten Liste → alle übernehmen. */
const OHNE = 'Ohne Vorschlag';
function wofuerAlleDialog(neu = true) {
  if (neu || !S.wPlan) {
    S.wPlan = S.k.buchungen.filter((b) => !b.wofuer && b.art === 'Ausgabe' && b.posten !== K.GK_POSTEN && b.posten !== 'Trade Republic (extra)' && !K.falscheUmbuchungen(S.k).includes(b)).sort((x, y) => (x.datum || '').localeCompare(y.datum || '') || x.nr.localeCompare(y.nr))
      .map((b) => ({ nr: b.nr, datum: b.datum, text: b.gegen || b.posten || b.beschreibung || '', betrag: b.betrag, wofuer: K.wofuerVorschlag(b, S.k) || '', vonHand: false }));
    S.wWahl = new Set(); S.wOffen = new Set();
  } else {
    document.querySelectorAll('#menu details[data-gruppe]').forEach((d) => { if (d.open) S.wOffen.add(d.dataset.gruppe); else S.wOffen.delete(d.dataset.gruppe); });
  }
  const plan = S.wPlan;
  if (!plan.length) { dialog(`<h2 id="menu-titel">Alles zugeordnet</h2><p>Alle Ausgaben haben schon ein „Wofür“.</p><div class="knoepfe"><button class="knopf" id="schliessen">OK</button></div>`); return; }
  const mit = plan.filter((x) => x.wofuer);
  const gruppen = [...WL(), OHNE].map((w) => [w, plan.filter((x) => (x.wofuer || OHNE) === w)]).filter(([, l]) => l.length);
  dialog(`<div class="k-wa"><h2 id="menu-titel">„Wofür“ nachtragen</h2>
    <p class="hint">Für <b>${mit.length}</b> von ${plan.length} Ausgaben steht ein „Wofür“ bereit. Hake an, was du ändern willst, dann „Ausgewählte ändern“. Gespeichert wird erst mit „Alle übernehmen“, und nur die Spalte „Wofür“.</p>
    ${gruppen.map(([w, l], gi) => `<details data-gruppe="${h(w)}" ${S.wOffen.has(w) ? 'open' : ''}><summary><b>${h(w)}</b> · ${l.length}× · ${eur(Math.round(l.reduce((s2, x) => s2 + x.betrag, 0) * 100) / 100)}</summary>
      <label class="k-wa-zeile alle"><input type="checkbox" id="k-wg-${gi}" data-gruppe="${h(w)}" ${l.every((x) => S.wWahl.has(x.nr)) ? 'checked' : ''}><span>alle in „${h(w)}“ anhaken</span></label>
      ${l.map((x) => `<label class="k-wa-zeile"><input type="checkbox" id="k-wh-${h(x.nr)}" data-nr="${h(x.nr)}" ${S.wWahl.has(x.nr) ? 'checked' : ''}><span>${h(x.nr)} · ${dk(x.datum)} · ${h(x.text)}${x.vonHand ? ' <i class="k-wa-neu">geändert</i>' : ''}</span><b>${eur(x.betrag)}</b></label>`).join('')}</details>`).join('')}
    <div class="k-wa-leiste"><button class="knopf leise" data-k="w-aendern" id="k-wa-aendern" ${S.wWahl.size ? '' : 'disabled'}>Ausgewählte ändern (${S.wWahl.size})</button>
      <button class="knopf" data-k="wofuer-alle-los" ${mit.length ? '' : 'disabled'}>Alle ${mit.length} übernehmen</button>
      <button class="knopf leise" id="schliessen">Abbrechen</button></div></div>`);
}
function wofuerAendernDialog() {
  const liste = S.wPlan.filter((x) => S.wWahl.has(x.nr));
  if (!liste.length) { wofuerAlleDialog(false); return; }
  const opts = (wert) => `${WL().map((w) => `<option ${w === wert ? 'selected' : ''}>${h(w)}</option>`).join('')}<option value="" ${!wert ? 'selected' : ''}>– leer lassen –</option>`;
  dialog(`<div class="k-wa"><h2 id="menu-titel">${liste.length} Buchung${liste.length > 1 ? 'en' : ''} ändern</h2>
    <p class="hint">Für jede Buchung einzeln wählen. Mit „Alle auf …“ kannst du alle auf einmal umstellen und danach einzelne wieder ändern.</p>
    ${liste.length > 1 ? `<label class="feld"><span>Alle auf …</span><select id="k-wa-alle"><option value="-">– einzeln wählen –</option>${opts('-')}</select></label>` : ''}
    <div class="k-wa-einzeln">${liste.map((x) => `<div class="k-wa-e"><div class="k-wa-e-kopf"><span>${h(x.nr)} · ${dk(x.datum)} · ${h(x.text)}</span><b>${eur(x.betrag)}</b></div>
      <select id="k-wa-n-${h(x.nr)}" data-nr="${h(x.nr)}" aria-label="Wofür für ${h(x.nr)}">${opts(x.wofuer)}</select></div>`).join('')}</div>
    <div class="k-wa-leiste"><button class="knopf" data-k="w-setzen">Übernehmen</button><button class="knopf leise" data-k="w-zurueck">Zurück</button></div></div>`);
}
function topfBudgetDialog(name, fehler) {
  const r = K.rechneKai(S.k, C.heute());
  const t = r.toepfe.find((x) => x.name === name); if (!t) return;
  const mitTopf = S.k.buchungen.filter((b) => b.topf === name && b.art !== K.UMBUCHUNG);
  dialog(`<h2 id="menu-titel">„${h(name)}“ als Monatsbudget</h2>
    <p>Statt zu sparen planst du jeden Monat <b>${eur(t.sparrate)}</b> für „${h(name)}“ ein. Die App zeigt dir, wie viel davon schon verbraucht ist und ob du drüber liegst.</p>
    <div class="rechnung" style="margin:10px 0">
      <div class="rz"><span>Topf „${h(name)}“<small>wird entfernt</small></span><b>${eur(t.stand)}</b></div>
      <div class="rz"><span>Neues Budget<small>Kosten, Bereich „Budget“, ab ${h(mn(r.akt))}</small></span><b>${eur(t.sparrate)}/Monat</b></div>
    </div>
    <p class="hint">Der Stand des Topfs (${eur(t.stand)}) zählt danach wieder als frei. In der Vorschau zieht das Budget jeden Monat ${eur(t.sparrate)} ab, im laufenden Monat nur, was noch übrig ist.${mitTopf.length ? ` ${mitTopf.length} Buchung${mitTopf.length > 1 ? 'en' : ''} mit diesem Topf ${mitTopf.length > 1 ? 'zählen' : 'zählt'} danach zum Budget.` : ''} Vorher legt die App eine Sicherungskopie der Datei in OneDrive an.</p>
    ${knoepfe(fehler, 'Umwandeln').replace('type="submit"', `type="button" data-k="topf-budget-los" data-arg="${h(name)}"`)}`);
}
function aufteilenDialog(nr, fehler, werte) {
  const b = S.k.buchungen.find((x) => x.nr === nr); if (!b) return;
  const teile = werte || [{ betrag: b.betrag, wofuer: b.wofuer || (K.istBudget(S.k, b.posten) ? b.posten : '') || K.wofuerVorschlag(b, S.k) || '' }, { betrag: '', wofuer: '' }];
  const zeile = (t, i) => `<div class="k-wa-e"><div class="k-wa-e-kopf"><span>${i === 0 ? 'Teil 1 <small>(bleibt ' + h(b.nr) + ')</small>' : 'Teil ' + (i + 1)}</span>${i > 1 ? `<button type="button" class="link" data-k="teil-weg" data-arg="${i}">entfernen</button>` : ''}</div>
    <div class="k-teil"><input name="tb${i}" inputmode="decimal" placeholder="0,00" value="${t.betrag === '' ? '' : String(t.betrag).replace('.', ',')}" aria-label="Betrag Teil ${i + 1}"><select name="tw${i}" aria-label="Wofür Teil ${i + 1}"><option value="">– Wofür –</option>${WL().map((w) => `<option ${w === t.wofuer ? 'selected' : ''}>${h(w)}</option>`).join('')}</select></div></div>`;
  dialog(`<form id="k-aufteilen" novalidate><h2 id="menu-titel">Buchung aufteilen</h2>
    <p class="hint">${h(b.nr)} · ${dl(b.datum)} · ${h(b.gegen || b.beschreibung || '')} · <b>${eur(b.betrag)}</b>. Teil 1 bleibt in der Buchung, jeder weitere Teil wird eine eigene Zeile („Teil von ${h(b.nr)}“). Zusammen bleibt es ${eur(b.betrag)}, der Kontostand ändert sich nicht.</p>
    <input type="hidden" name="nr" value="${h(nr)}"><input type="hidden" name="n" value="${teile.length}">
    <div class="k-wa-einzeln">${teile.map(zeile).join('')}</div>
    ${teile.length < 5 ? '<button type="button" class="knopf leise klein" data-k="teil-dazu">+ weiterer Teil</button>' : ''}
    <p class="k-teil-rest" id="k-teil-rest"></p>
    ${knoepfe(fehler, 'Aufteilen')}</form>`);
  const f = document.getElementById('k-aufteilen');
  const rest = () => {
    const n = Number(f.n.value); let sum = 0;
    for (let i = 0; i < n; i++) sum += zahl(f['tb' + i].value) || 0;
    const r2 = Math.round((b.betrag - sum) * 100) / 100;
    document.getElementById('k-teil-rest').innerHTML = Math.abs(r2) < 0.005 ? '✓ Die Teile ergeben genau ' + eur(b.betrag) : r2 > 0 ? `Noch <b>${eur(r2)}</b> zu verteilen` : `<b>${eur(-r2)}</b> zu viel`;
  };
  f.addEventListener('input', rest);
  // Teil 1 automatisch auf den Rest setzen, wenn ein anderer Teil geändert wird
  f.addEventListener('change', (e) => {
    if (!/^tb[1-9]$/.test(e.target.name)) return;
    const n = Number(f.n.value); let andere = 0;
    for (let i = 1; i < n; i++) andere += zahl(f['tb' + i].value) || 0;
    const r1 = Math.round((b.betrag - andere) * 100) / 100;
    if (r1 > 0) f.tb0.value = String(r1.toFixed(2)).replace('.', ',');
    rest();
  });
  rest();
}
const teileAusForm = (f) => Array.from({ length: Number(f.n.value) }, (_, i) => ({ betrag: f['tb' + i].value.trim() === '' ? '' : zahl(f['tb' + i].value), wofuer: f['tw' + i].value }));
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
    case 'pruefen': pruefDialog(); return true;
    case 'pruef-ansehen': if (C.menu.open) C.menu.close(); S.suche = arg; S.fMonat = ''; S.fArt = ''; S.fWofuer = ''; S.tab = 'k-buchungen'; merk('ug.ktab', S.tab); C.zeichne(); window.scrollTo(0, 0); return true;
    case 'wzeit': S.wZeit = arg; C.zeichne(); return true;
    case 'wofuer-filter': S.fWofuer = arg; S.suche = ''; S.fArt = ''; S.fMonat = S.tab === 'k-ausgaben' ? S.aMonat : S.wZeit === 'monat' ? C.heute().slice(0, 7) : ''; S.tab = 'k-buchungen'; merk('ug.ktab', S.tab); C.zeichne(); window.scrollTo(0, 0); return true;
    case 'wofuer-alle': wofuerAlleDialog(); return true;
    case 'topf-budget': topfBudgetDialog(arg); return true;
    case 'aufteilen': aufteilenDialog(arg); return true;
    case 'teil-dazu': case 'teil-weg': {
      const f = document.getElementById('k-aufteilen'); const tl = teileAusForm(f);
      if (a === 'teil-dazu') tl.push({ betrag: '', wofuer: '' }); else tl.splice(Number(arg), 1);
      aufteilenDialog(f.nr.value, null, tl); return true;
    }
    case 'ausgaben': S.tab = 'k-ausgaben'; S.aMonat = S.aMonat || C.heute().slice(0, 7); C.zeichne(); window.scrollTo(0, 0); return true;
    case 'a-monat': S.aMonat = arg; C.zeichne(); return true;
    case 'a-zurueck': S.tab = 'k-uebersicht'; merk('ug.ktab', S.tab); C.zeichne(); window.scrollTo(0, 0); return true;
    case 'topf-budget-los': {
      const name = arg;
      const r = K.rechneKai(S.k, C.heute()); const tp = r.toepfe.find((x) => x.name === name);
      if (!tp) return true;
      if (K.istBudget(S.k, name)) { topfBudgetDialog(name, `Ein Budget „${name}“ gibt es schon.`); return true; }
      const mitTopf = S.k.buchungen.filter((b) => b.topf === name && b.art !== K.UMBUCHUNG);
      await schreiben(async (api) => {
        await C.od.sicherungAnlegen(S.ort, 'vor-Budget-' + name.replace(/[^A-Za-z0-9äöüÄÖÜß-]+/g, '_'));
        const [tv, tb] = K.BEREICHE.toepfe;
        const tn = (await api.lies('Töpfe', `B${tv}:B${tb}`)).values.map((x) => String(x[0] ?? ''));
        const ti = tn.indexOf(name); if (ti < 0) throw new Error(`Topf „${name}“ nicht gefunden.`);
        const [kv, kb] = K.BEREICHE.kosten;
        const kn = (await api.lies('Kosten', `B${kv}:B${kb}`)).values; const ki = kn.findIndex((x) => x[0] === '' || x[0] == null);
        if (ki < 0) throw new Error('Im Blatt „Kosten“ ist keine Zeile mehr frei.');
        const kz = kv + ki;
        const ids = (await api.lies('Kosten', `A${kv}:A${kb}`)).values.map((x) => String(x[0] || '')).filter((x) => /^K\d+$/.test(x)).map((x) => Number(x.slice(1)));
        await api.schreib('Kosten', `A${kz}:J${kz}`, [['K' + String(Math.max(0, ...ids) + 1).padStart(2, '0'), name, K.BUDGET, tp.sparrate, 'monatlich', '', '', S.k.konto.name, mSer(r.akt), '']]);
        await api.schreib('Kosten', `N${kz}`, [[`Monatsbudget, vorher Topf (umgestellt ${dl(C.heute())}).`]]);
        for (const b of mitTopf) { const z = await zeileVonNr(api, b.nr); await api.schreib('Buchungen', `J${z}`, [['']]); if (!b.posten || b.posten === 'Sonstiges') await api.schreib('Buchungen', `I${z}`, [[name]]); }
        const tz = tv + ti;
        await api.schreib('Töpfe', `A${tz}:F${tz}`, [['', '', '', '', '', '']]);
        await api.schreib('Töpfe', `J${tz}:K${tz}`, [['', '']]);
      }, `„${name}“ ist jetzt ein Monatsbudget (${eur(tp.sparrate)}). Sicherung liegt in OneDrive.`, { knopf: t, dialogFehler: (fe) => topfBudgetDialog(name, fe) });
      S.topf = null;
      return true;
    }
    case 'w-aendern': wofuerAlleDialog(false); wofuerAendernDialog(); return true;
    case 'w-zurueck': wofuerAlleDialog(false); return true;
    case 'w-setzen': {
      for (const x of S.wPlan) {
        const sel = document.getElementById('k-wa-n-' + x.nr);
        if (!sel) continue;
        if (sel.value !== x.wofuer) { x.wofuer = sel.value; x.vonHand = true; }
        S.wOffen.add(x.wofuer || OHNE);
      }
      S.wWahl = new Set();
      wofuerAlleDialog(false); return true;
    }
    case 'wofuer-alle-los': {
      const plan = (S.wPlan || []).filter((p) => p.wofuer);
      await schreiben(async (api) => {
        const [von, bis] = K.BEREICHE.buchungen;
        const a = (await api.lies('Buchungen', `A${von}:A${bis}`)).values.map((r) => String(r[0] ?? ''));
        let letzte = -1; a.forEach((x, i) => { if (x) letzte = i; });
        if (letzte < 0) return 0;
        await wofuerKopf(api);
        const q = (await api.lies('Buchungen', `${W}${von}:${W}${von + letzte}`)).values.map((r) => String(r[0] ?? ''));
        const neu = new Map(plan.map((p) => [p.nr, p.wofuer]));
        let n = 0;
        const werte = a.slice(0, letzte + 1).map((nr, i) => { if (!q[i] && neu.has(nr)) { n++; return [neu.get(nr)]; } return [q[i]]; });
        if (n) await api.schreib('Buchungen', `${W}${von}:${W}${von + letzte}`, werte);
        // Budget als „Wofür“ → auch Posten (Spalte I), wenn dort noch nichts steht
        const bud = plan.filter((p) => K.istBudget(S.k, p.wofuer));
        if (bud.length) {
          const iw = (await api.lies('Buchungen', `I${von}:I${von + letzte}`)).values.map((r) => String(r[0] ?? ''));
          for (const p of bud) { const i = a.indexOf(p.nr); if (i >= 0 && (!iw[i] || iw[i] === 'Sonstiges')) await api.schreib('Buchungen', `I${von + i}`, [[p.wofuer]]); }
        }
        return n;
      }, (n) => `„Wofür“ bei ${n} Buchungen eingetragen.`, { knopf: t });
      S.wPlan = null;
      return true;
    }
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
    case 'loeschen': await schreiben(async (api) => { const z = await zeileVonNr(api, arg); await api.schreib('Buchungen', `A${z}:N${z}`, [Array(14).fill('')]); if (S.k.buchungen.find((b) => b.nr === arg)?.wofuer) await api.schreib('Buchungen', `${W}${z}`, [['']]); }, `${arg} gelöscht.`, { knopf: t }); return true;
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
        : { datum: x.datum, giltFuer: x.giltFuer, art: x.art, konto, gegen: (x.gegen || '').trim(), betrag, posten: x.posten, topf: x.topf, beschreibung: x.beschreibung.trim(), kommentar: x.kommentar.trim(), wofuer: x.wofuer || '' };
      if (b.wofuer) budgetPosten(b);
      if (x.nr) {
        const alt = S.k.buchungen.find((y) => y.nr === x.nr);
        await schreiben(async (api) => {
          const z = await zeileVonNr(api, x.nr); await api.schreib('Buchungen', `B${z}:L${z}`, [buchungZeile({ ...alt, ...b }, x.nr).slice(1, 12)]);
          if ((b.wofuer || '') !== (alt.wofuer || '')) { await wofuerKopf(api); await api.schreib('Buchungen', `${W}${z}`, [[b.wofuer || '']]); }
        }, `${x.nr} geändert.`, { knopf, dialogFehler: (fe) => eintragenDialog(fe, { ...x }) });
      } else {
        await schreiben(async (api) => (await buchungenAnhaengen(api, [{ ...b, quelle: 'manuell' }]))[0], (nr) => `${nr} eingetragen: ${b.posten || b.gegen || b.art}, ${eur(betrag)}.`, { knopf, dialogFehler: (fe) => eintragenDialog(fe, { ...x }) });
      }
      return true;
    }
    case 'k-aufteilen': {
      const b = S.k.buchungen.find((y) => y.nr === x.nr);
      const alle = teileAusForm(form);
      const tl = alle.filter((t) => t.betrag !== '');
      const f = K.aufteilungPruefen(b, tl);
      if (f) { aufteilenDialog(x.nr, f, alle); return true; }
      const [t1, ...rest] = tl;
      const p1 = K.istBudget(S.k, b.posten) && b.posten !== t1.wofuer ? '' : b.posten || '';
      const neue = rest.map((t) => budgetPosten({ datum: b.datum, giltFuer: b.giltFuer, art: 'Ausgabe', konto: b.konto, gegen: b.gegen, betrag: t.betrag, kategorie: b.kategorie,
        posten: '', topf: b.topf, beschreibung: b.beschreibung, kommentar: `Teil von ${b.nr}`, quelle: 'Aufteilung', importId: '', wofuer: t.wofuer }));
      await schreiben(async (api) => {
        const z = await zeileVonNr(api, b.nr);
        const g = (await api.lies('Buchungen', `G${z}`)).values[0][0];
        if (Math.abs(Number(g) - b.betrag) > 0.005) throw new Error(`${b.nr} hat in der Datei gerade einen anderen Betrag. Bitte neu laden.`);
        await api.schreib('Buchungen', `G${z}`, [[t1.betrag]]);
        await api.schreib('Buchungen', `I${z}`, [[budgetPosten({ wofuer: t1.wofuer, posten: p1 }).posten || '']]);
        await wofuerKopf(api); await api.schreib('Buchungen', `${W}${z}`, [[t1.wofuer || '']]);
        await buchungenAnhaengen(api, neue);
      }, `${b.nr} aufgeteilt in ${tl.length} Teile.`, { knopf, dialogFehler: (fe) => aufteilenDialog(x.nr, fe, alle) });
      return true;
    }
    case 'k-pruefen': {
      const konten = K.eigeneKonten(S.k);
      const bank = {}; let fehlt = false;
      konten.forEach((kt, i) => { const v = zahl(x['b' + i]); if (isFinite(v)) bank[kt.name] = v; else if (String(x['b' + i] || '').trim()) fehlt = true; });
      if (fehlt || !Object.keys(bank).length) { pruefDialog('Bitte mindestens einen Stand als Zahl eintragen, z. B. 245,30.', x); return true; }
      const e = K.kontenPruefen(S.k, bank);
      merk('ug.kpruef', { zeit: new Date().toISOString(), ok: e.ok, konten: e.konten.map(({ name, bank: b, stand, diff, ok }) => ({ name, bank: b, stand, diff, ok })) });
      pruefErgebnis(e); C.zeichne();
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
        else neu.push({ ...r.vorschlag, posten: r.vorschlag.art === K.UMBUCHUNG ? null : x['p' + i] || null, giltFuer: x['g' + i] || r.vorschlag.giltFuer, wofuer: r.vorschlag.art === 'Ausgabe' ? x['w2' + i] || '' : '' });
        budgetPosten(neu[neu.length - 1]);
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
  if (el.id === 'k-fwofuer') { S.fWofuer = el.value; C.zeichne(); return true; }
  if (el.id === 'k-wa-alle') { if (el.value !== '-') document.querySelectorAll('select[id^="k-wa-n-"]').forEach((x) => { x.value = el.value; }); return true; }
  if (el.id?.startsWith('k-wa-n-')) return true;
  if (el.id?.startsWith('k-wh-') || el.id?.startsWith('k-wg-')) {
    const nrs = el.dataset.nr ? [el.dataset.nr] : S.wPlan.filter((x) => (x.wofuer || OHNE) === el.dataset.gruppe).map((x) => x.nr);
    for (const nr of nrs) { if (el.checked) S.wWahl.add(nr); else S.wWahl.delete(nr); const cb = document.getElementById('k-wh-' + nr); if (cb) cb.checked = el.checked; }
    const k = document.getElementById('k-wa-aendern'); if (k) { k.disabled = !S.wWahl.size; k.textContent = `Ausgewählte ändern (${S.wWahl.size})`; }
    return true;
  }
  return false;
}
