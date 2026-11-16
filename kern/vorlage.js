// Baut eine Datei im einheitlichen Format (kern/format.js) aus Daten.
// Wird für die Vorlage (erfundene Daten) und beim Umzug (echte Daten, im Browser) benutzt.
// Alle Rechnungen in Excel stehen hier; der Rechenkern (kern/einheit.js) rechnet genau so.

import { BLAETTER, GRUNDBLAETTER, LISTEN, KATEGORIEN_START, FORMAT_VERSION, KOPFZEILE, ERSTE_ZEILE, spalten, sp, WIDGETS_START, VORHER_AKTIV } from './format.js';
import { mappeSchreiben, neuesBlatt, STIL } from './mappe.js';

const serial = (iso) => { if (!iso) return null; const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number); return (Date.UTC(y, m - 1, d || 1) - Date.UTC(1899, 11, 30)) / 864e5; };
const q = (b) => `'${b}'`;
const letzte = (blatt) => KOPFZEILE + BLAETTER[blatt].zeilen;
/** Absoluter Bereich einer Spalte, z. B. 'Buchungen'!$G$5:$G$2004 */
export const R = (blatt, key) => { const c = sp(blatt, key); return `${q(blatt)}!$${c}$${ERSTE_ZEILE}:$${c}$${letzte(blatt)}`; };
const ART_GELD = ['"<>In Topf legen"', '"<>Aus Topf nehmen"'];
const ABSTAND = (zelle) => `IF(${zelle}="",1,IFERROR(CHOOSE(MATCH(${zelle},{"monatlich","vierteljährlich","halbjährlich","jährlich"},0),1,3,6,12),1))`;
// Konto zählt für Frei (eigenes Konto, Zählt für Frei = Ja). Leeres Konto zählt mit.
const ZAEHLT = (zelle) => `OR(${zelle}="",COUNTIFS(${R('Konten', 'name')},${zelle},${R('Konten', 'eigen')},"Ja",${R('Konten', 'frei')},"Ja")>0)`;

const BU = (k) => R('Buchungen', k);

// ---------------------------------------------------------------- Formeln je Zeile
const FORMELN = {
  Konten: {
    stand: (r, c) => `IF(OR(${c.eigen}${r}<>"Ja",${c.anfang}${r}=""),"",${c.anfang}${r}`
      + `+SUMIFS(${BU('betrag')},${BU('nach')},${c.name}${r},${BU('datum')},">"&N(${c.stichtag}${r}),${BU('datum')},"<="&Heute,${BU('art')},${ART_GELD[0]},${BU('art')},${ART_GELD[1]})`
      + `-SUMIFS(${BU('betrag')},${BU('von')},${c.name}${r},${BU('datum')},">"&N(${c.stichtag}${r}),${BU('datum')},"<="&Heute,${BU('art')},${ART_GELD[0]},${BU('art')},${ART_GELD[1]}))`,
  },
  Töpfe: {
    stand: (r, c) => {
      const f = (art) => `SUMIFS(${BU('betrag')},${BU('topf')},${c.name}${r},${BU('art')},"${art}",${BU('datum')},">"&N(${c.stichtag}${r}),${BU('datum')},"<="&Heute)`;
      const sp_ = (art) => `SUMIFS(${BU('betrag')},${BU('topf')},"Sparplan (verteilt)",${BU('art')},"${art}",${BU('datum')},"<="&Heute)`;
      const bisM = `MIN(IF(${c.bis}${r}="",AktMonat,${c.bis}${r}),AktMonat)`;
      const monate = `MAX(0,(YEAR(${bisM})-YEAR(${c.ab}${r}))*12+MONTH(${bisM})-MONTH(${c.ab}${r})+1)`;
      return `IF(${c.name}${r}="","",N(${c.anfang}${r})`
        + `+IF(AND(${c.automatisch}${r}="Ja",${c.ab}${r}<>""),N(${c.sparrate}${r})*${monate},0)`
        + `+${f('Einnahme')}+${f('In Topf legen')}-${f('Ausgabe')}-${f('Aus Topf nehmen')}`
        + `+IF(SUM(${R('Töpfe', 'sparrate')})=0,0,N(${c.sparrate}${r})/SUM(${R('Töpfe', 'sparrate')})*(${sp_('Einnahme')}-${sp_('Ausgabe')})))`;
    },
  },
  Kontostände: {
    datei: (r, c) => {
      const kz = `MATCH(${c.konto}${r},${R('Konten', 'name')},0)`;
      const st = `N(INDEX(${R('Konten', 'stichtag')},${kz}))`;
      const s = (seite) => `SUMIFS(${BU('betrag')},${BU(seite)},${c.konto}${r},${BU('datum')},">"&${st},${BU('datum')},"<="&EOMONTH(${c.monat}${r},0),${BU('art')},${ART_GELD[0]},${BU('art')},${ART_GELD[1]})`;
      return `IF(OR(${c.monat}${r}="",${c.konto}${r}=""),"",IFERROR(INDEX(${R('Konten', 'anfang')},${kz})+${s('nach')}-${s('von')},""))`;
    },
    differenz: (r, c) => `IF(OR(${c.bank}${r}="",${c.datei}${r}=""),"",ROUND(${c.bank}${r}-${c.datei}${r},2))`,
  },
  Kosten: {
    monatswert: (r, c) => `IF(${c.posten}${r}="",0,IF(${c.status}${r}="beendet",0,${c.betrag}${r}/${ABSTAND(c.rhythmus + r)}))`,
    planAkt: (r, c) => planFormel(r, c, 'AktMonat'),
    istAkt: (r, c) => `IF(${c.posten}${r}="","",SUMIFS(${BU('betrag')},${BU('posten')},${c.posten}${r},${BU('giltFuer')},AktMonat,${BU('art')},"Ausgabe"))`,
    offenAkt: (r, c) => `IF(${c.posten}${r}="","",IF(${ZAEHLT(c.konto + r)},MAX(0,${c.planAkt}${r}-${c.istAkt}${r}),0))`,
  },
  Einnahmen: {
    planAkt: (r, c) => planFormel(r, c, 'AktMonat'),
    istAkt: (r, c) => `IF(${c.posten}${r}="","",SUMIFS(${BU('betrag')},${BU('posten')},${c.posten}${r},${BU('giltFuer')},AktMonat,${BU('art')},"Einnahme"))`,
    offenAkt: (r, c) => `IF(${c.posten}${r}="","",IF(${ZAEHLT(c.konto + r)},MAX(0,${c.planAkt}${r}-${c.istAkt}${r}),0))`,
  },
  Limits: {
    istAkt: (r, c) => `IF(${c.kategorie}${r}="","",SUMIFS(${BU('betrag')},${BU('kategorie')},${c.kategorie}${r},${BU('giltFuer')},AktMonat,${BU('art')},"Ausgabe"))`,
    restAkt: (r, c) => `IF(${c.kategorie}${r}="","",${c.betrag}${r}-${c.istAkt}${r})`,
  },
  'Soll-Überweisungen': {
    ist: (r, c) => `IF(${c.monat}${r}="",0,SUMIFS(${BU('betrag')},${BU('von')},${c.von}${r},${BU('nach')},${c.nach}${r},${BU('giltFuer')},${c.monat}${r}))`,
    offen: (r, c) => `IF(${c.monat}${r}="",0,MAX(0,N(${c.soll}${r})-${c.ist}${r}))`,
    wirkung: (r, c) => `IF(${c.monat}${r}="",0,IF(COUNTIFS(${R('Konten', 'name')},${c.nach}${r},${R('Konten', 'eigen')},"Ja",${R('Konten', 'frei')},"Ja")>0,1,IF(COUNTIFS(${R('Konten', 'name')},${c.von}${r},${R('Konten', 'eigen')},"Ja",${R('Konten', 'frei')},"Ja")>0,-1,0)))`,
  },
};
/** Plan eines Postens in einem Monat (wie kern/einheit.js plan()). mZelle = Zelle mit dem Monat (1. des Monats). */
function planFormel(r, c, mZelle, pre = '') {
  const x = (k) => `${pre}${c[k]}${r}`;
  const ab = ABSTAND(x('rhythmus'));
  return `IF(${x('posten')}="",0,IF(OR(${x('status')}="beendet",AND(${x('ab')}<>"",${mZelle}<${x('ab')}),AND(${x('bis')}<>"",${mZelle}>${x('bis')})),0,`
    + `IF(${ab}=1,N(${x('betrag')}),IF(${x('faellig')}="",0,IF(MOD(MONTH(${mZelle})-${x('faellig')},${ab})=0,N(${x('betrag')}),0)))))`;
}

// ---------------------------------------------------------------- Werte
function wert(typ, v) {
  if (v == null || v === '') return null;
  if (typ === 'd' || typ === 'm') return typeof v === 'number' ? v : serial(typ === 'm' && String(v).length === 7 ? v + '-01' : v);
  if (typ === 'z' || typ === 'n') return typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
  if (typ === 'j') return v === true || v === 'Ja' ? 'Ja' : v === false || v === 'Nein' ? 'Nein' : String(v);
  return typeof v === 'number' ? v : String(v);
}
const STIL_TYP = { d: STIL.datum, m: STIL.monat, z: STIL.euro };
const FORMEL_STIL = { Töpfe: STIL.formelEuro, Konten: STIL.formelEuro, Kontostände: STIL.formelEuro, Kosten: STIL.formelEuro, Einnahmen: STIL.formelEuro, Limits: STIL.formelEuro, 'Soll-Überweisungen': STIL.formelEuro };

function tabellenBlatt(name, zeilen = []) {
  const def = BLAETTER[name];
  const sps = spalten(name);
  const c = Object.fromEntries(sps.map((s) => [s.key, s.buchst]));
  const b = neuesBlatt(name, { breiten: sps.map((s) => s.breite), fixiert: KOPFZEILE + 1 });
  b.set('A1', { v: def.titel, s: STIL.titel });
  b.set('A2', { v: def.text, s: STIL.hinweis });
  sps.forEach((s) => b.set(s.buchst + KOPFZEILE, { v: s.name, s: STIL.kopf }));
  b.hoehen = { [KOPFZEILE]: 30 };
  const bis = letzte(name);
  if (zeilen.length > def.zeilen) throw new Error(`${name}: ${zeilen.length} Zeilen, Platz für ${def.zeilen}`);
  const formeln = FORMELN[name] || {};
  for (let i = 0; i < def.zeilen; i++) {
    const r = ERSTE_ZEILE + i; const d = zeilen[i];
    for (const s of sps) {
      if (s.typ === 'f') {
        if (!formeln[s.key]) continue;
        // Formeln nur so weit vorbelegen, wie sinnvoll: Daten + Puffer (Buchungen haben keine Formeln)
        if (i >= Math.max(zeilen.length + 40, Math.min(def.zeilen, 60))) continue;
        b.set(s.buchst + r, { f: formeln[s.key](r, c), s: FORMEL_STIL[name] || STIL.formelText });
      } else if (d) {
        const v = wert(s.typ, d[s.key]);
        if (v != null) b.set(s.buchst + r, { v, s: STIL_TYP[s.typ] || 0 });
      }
    }
  }
  b.tabelle = { name: def.tabelle, ref: `A${KOPFZEILE}:${sps.at(-1).buchst}${bis}`, spalten: sps.map((s) => s.name) };
  b.validierungen = sps.filter((s) => s.liste).map((s) => ({ ref: `${s.buchst}${ERSTE_ZEILE}:${s.buchst}${bis}`, formel: 'lst' + s.liste.replace(/\W/g, ''), streng: false }));
  // Hilfreiche Auswahl ohne Zwang: Kategorie und Topf
  for (const s of sps) {
    if (s.key === 'kategorie' && name !== 'Kategorien') b.validierungen.push({ ref: `${s.buchst}${ERSTE_ZEILE}:${s.buchst}${bis}`, formel: R('Kategorien', 'name') });
    if (s.key === 'topf' && name !== 'Töpfe') b.validierungen.push({ ref: `${s.buchst}${ERSTE_ZEILE}:${s.buchst}${bis}`, formel: R('Töpfe', 'name') });
  }
  return { b, c };
}

function listenBlatt() {
  const b = neuesBlatt('Listen', { fixiert: 2 });
  const namen = [];
  const keys = ['Monate', ...Object.keys(LISTEN)];
  b.breiten = keys.map(() => 18);
  keys.forEach((k, i) => {
    const col = String.fromCharCode(65 + i);
    b.set(col + '1', { v: k, s: STIL.kopf });
    // Monate ab November 2026 (wie die alte gemeinsame Datei: Aufteilung/Überweisung zeigen auf Listen!A2 …)
    const werte = k === 'Monate' ? Array.from({ length: 72 }, (_, j) => serial(`${2026 + Math.floor((j + 10) / 12)}-${String((j + 10) % 12 + 1).padStart(2, '0')}-01`)) : LISTEN[k];
    werte.forEach((v, j) => b.set(col + (j + 2), { v, s: k === 'Monate' ? STIL.monat : 0 }));
    namen.push({ name: 'lst' + k.replace(/\W/g, ''), ref: `Listen!$${col}$2:$${col}$${werte.length + 1}` });
  });
  return { b, namen };
}

function uebersicht(bereichTitel) {
  const b = neuesBlatt('Übersicht', { breiten: [3, 38, 16, 4, 40] });
  const s = (ref, v, st) => b.set(ref, typeof v === 'string' && v.startsWith('=') ? { f: v, s: st ?? STIL.formelEuro } : { v, s: st ?? 0 });
  s('B1', bereichTitel, STIL.titel);
  s('B2', 'Oben den Monat wählen. Grau = Formel, nicht überschreiben. Die App rechnet genauso.', STIL.hinweis);
  s('B4', 'Monat', STIL.fett);
  b.set('C4', { f: 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)', s: STIL.eingabe });
  s('E4', '← hier einen Monat eintragen (z. B. 01.12.2026); leer lassen = dieser Monat', STIL.hinweis);
  s('B5', 'Heute', STIL.fett); b.set('C5', { f: 'TODAY()', s: STIL.formelDatum });
  s('B7', 'Kontostände (heute)', STIL.fett);
  const kStart = 8;
  for (let i = 0; i < 10; i++) {
    const r = kStart + i; const k = ERSTE_ZEILE + i; const c = (key) => `'Konten'!$${sp('Konten', key)}$${k}`;
    b.set('B' + r, { f: `IF(AND(${c('eigen')}="Ja",${c('status')}<>"geschlossen"),${c('name')},"")`, s: STIL.formelText, t: 'str' });
    b.set('C' + r, { f: `IF(B${r}="","",${c('stand')})`, s: STIL.formelEuro });
  }
  const z = kStart + 11;
  s('B' + z, 'Geld auf Konten, die für „Frei“ zählen', STIL.fett);
  s('C' + z, `=SUMIFS(${R('Konten', 'stand')},${R('Konten', 'frei')},"Ja",${R('Konten', 'status')},"<>geschlossen")`);
  s('B' + (z + 1), 'davon in Töpfen (verplant)');
  s('C' + (z + 1), `=SUMIFS(${R('Töpfe', 'stand')},${R('Töpfe', 'frei')},"<>Ja")`);
  s('B' + (z + 2), 'Frei heute (ohne Töpfe)', STIL.fett);
  s('C' + (z + 2), `=C${z}-C${z + 1}`, STIL.fettEuro);
  s('B' + (z + 4), 'Noch offen in diesem Monat', STIL.fett);
  s('B' + (z + 5), '+ Einnahmen, die noch kommen'); s('C' + (z + 5), `=SUM(${R('Einnahmen', 'offenAkt')})`);
  s('B' + (z + 6), '− Kosten, die noch abgehen'); s('C' + (z + 6), `=SUM(${R('Kosten', 'offenAkt')})`);
  s('B' + (z + 7), '± Überweisungen zwischen den Bereichen');
  s('C' + (z + 7), `=SUMPRODUCT((${R('Soll-Überweisungen', 'monat')}=AktMonat)*${R('Soll-Überweisungen', 'offen')}*${R('Soll-Überweisungen', 'wirkung')})`);
  s('B' + (z + 9), 'Was ist frei? (bis Monatsende)', STIL.fett);
  s('C' + (z + 9), `=C${z + 2}+C${z + 5}-C${z + 6}+C${z + 7}`, STIL.fettEuro);
  s('E' + (z + 9), 'Limits werden nicht abgezogen, sie warnen nur.', STIL.hinweis);
  return { b, namen: [{ name: 'AktMonat', ref: "'Übersicht'!$C$4" }, { name: 'Heute', ref: "'Übersicht'!$C$5" }], zeilen: { freiHeute: z + 2, frei: z + 9 }, freiHeuteRef: `'Übersicht'!$C$${z + 2}` };
}

function vorschau(freiHeuteRef) {
  const b = neuesBlatt('Vorschau', { fixiert: 5 });
  const M = 'CDEFGHIJKLMN'.split('');
  b.breiten = [8, 30, ...M.map(() => 12), 9, 9];
  b.set('B1', { v: 'Vorschau 12 Monate', s: STIL.titel });
  b.set('B2', { v: 'Was im Monat noch kommt und abgeht (Plan minus schon gebucht). Limits werden nicht abgezogen.', s: STIL.hinweis });
  b.set('B4', { v: 'Monat', s: STIL.kopf });
  M.forEach((c, i) => b.set(c + '4', { f: i === 0 ? 'AktMonat' : `EDATE(${M[i - 1]}4,1)`, s: STIL.formelMonat }));
  const sum = {};
  let r = 6;
  const block = (blatt, titel, art) => {
    b.set('B' + r, { v: titel, s: STIL.fett }); r++;
    const von = r;
    const c = Object.fromEntries(spalten(blatt).map((s) => [s.key, s.buchst]));
    const n = Math.min(BLAETTER[blatt].zeilen, blatt === 'Kosten' ? 60 : 20);
    for (let i = 0; i < n; i++, r++) {
      const k = ERSTE_ZEILE + i; const pre = `${q(blatt)}!$`;
      const cc = Object.fromEntries(Object.entries(c).map(([kk, v]) => [kk, v + '$']));
      b.set('A' + r, { v: k, s: STIL.hinweis });
      b.set('B' + r, { f: `IF(${pre}${c.posten}$${k}="","",${pre}${c.posten}$${k})`, s: STIL.formelText, t: 'str' });
      b.set('P' + r, { f: `IF(${ZAEHLT(`${pre}${c.konto}$${k}`)},1,0)`, s: STIL.formelText });
      M.forEach((mc) => {
        const plan = planFormel(k, cc, `${mc}$4`, pre);
        const ist = `SUMIFS(${BU('betrag')},${BU('posten')},${pre}${c.posten}$${k},${BU('giltFuer')},${mc}$4,${BU('art')},"${art}")`;
        b.set(mc + r, { f: `IF(OR(${pre}${c.posten}$${k}="",$P${r}=0),0,MAX(0,${plan}-${ist}))`, s: STIL.formelEuro });
      });
    }
    sum[art] = [von, r - 1];
    r++;
  };
  block('Einnahmen', 'Einnahmen (noch offen)', 'Einnahme');
  block('Kosten', 'Kosten (noch offen)', 'Ausgabe');
  const zeile = (titel, f, st = STIL.formelEuro) => { b.set('B' + r, { v: titel, s: STIL.fett }); M.forEach((mc, i) => b.set(mc + r, { f: f(mc, i), s: st })); return r++; };
  const rE = zeile('Summe Einnahmen', (mc) => `SUM(${mc}${sum.Einnahme[0]}:${mc}${sum.Einnahme[1]})`);
  const rK = zeile('Summe Kosten', (mc) => `SUM(${mc}${sum.Ausgabe[0]}:${mc}${sum.Ausgabe[1]})`);
  const T = (k) => R('Töpfe', k);
  const rS = zeile('Sparraten (automatisch zugeteilt)', (mc, i) => i === 0 ? '0'
    : `SUMPRODUCT((${T('automatisch')}="Ja")*(${T('frei')}<>"Ja")*(((${T('ab')}="")+(${T('ab')}<=${mc}$4))>0)*(((${T('bis')}="")+(${T('bis')}>=${mc}$4))>0),${T('sparrate')})`);
  const S = (k) => R('Soll-Überweisungen', k);
  const rU = zeile('Überweisungen zwischen Bereichen (±)', (mc) => `SUMPRODUCT((${S('monat')}=${mc}$4)*${S('offen')}*${S('wirkung')})`);
  const rG = zeile('Ergebnis des Monats', (mc) => `${mc}${rE}-${mc}${rK}-${mc}${rS}+${mc}${rU}`, STIL.fettEuro);
  const rA = zeile('Frei am Monatsanfang', (mc, i) => i === 0 ? freiHeuteRef : `${M[i - 1]}${r + 1}`);
  const rZ = zeile('Frei am Monatsende', (mc) => `${mc}${rA}+${mc}${rG}`, STIL.fettEuro);
  zeile('Hinweis', (mc) => `IF(${mc}${rZ}<0,"Lücke","OK")`, STIL.formelText);
  return { b, zeilen: { freiEnde: rZ, ergebnis: rG } };
}

const ERKLAERUNG = [
  ['So funktioniert die Datei', ''],
  ['Gleich in allen Bereichen', 'Gemeinsam, Tagesgeld, Natalie und Kai haben dieselben Blätter mit denselben Spalten. Die App „Unser Geld“ liest und schreibt diese Datei.'],
  ['Buchungen', 'Jede Zahlung eine Zeile. Von → Nach zeigt, woher das Geld kommt und wohin es geht. Eigene Konten stehen im Blatt „Konten“ mit „Eigenes Konto = Ja“.'],
  ['Art', 'Einnahme: Geld kommt auf ein eigenes Konto. Ausgabe: Geld geht von einem eigenen Konto weg. Umbuchung: zwischen zwei eigenen Konten. In Topf legen / Aus Topf nehmen: Geld bleibt auf dem Konto, wird nur einem Topf zugeteilt.'],
  ['Posten', 'Geplante feste Kosten und Einnahmen (Blätter „Kosten“ und „Einnahmen“). Steht in einer Buchung derselbe Posten, gilt der Plan für den Monat als bezahlt.'],
  ['Topf', 'Gespartes Geld für einen Zweck. Einnahme mit Topf = mehr im Topf, Ausgabe mit Topf = weniger im Topf. „Sparplan (verteilt)“ wird nach Sparrate auf alle Töpfe verteilt.'],
  ['Kategorie', 'Wofür wurde das Geld ausgegeben? Eine Liste für alle Bereiche (Blatt „Kategorien“). Freiwillig.'],
  ['Bank-Kategorie', 'Die Einordnung, die die Bank im Export mitliefert. Nur zur Info.'],
  ['Limit', 'Obergrenze pro Monat für eine Kategorie. Warnt nur und wird nicht vorab abgezogen.'],
  ['Was ist frei?', 'Geld auf den Konten, die für Frei zählen, minus Töpfe, plus Einnahmen, die noch kommen, minus Kosten, die noch abgehen.'],
  ['Kontostände', 'Monatsabschluss: Stand aus der Bank-App eintragen. Die Differenz zeigt, ob etwas fehlt oder doppelt ist.'],
  ['Gegenstück', 'Steht dieselbe Überweisung auch in einer anderen Datei (z. B. Gemeinschaftskonto → Tagesgeld), steht hier deren Nummer.'],
  ['Grau', 'Graue Zellen sind Formeln. Nicht überschreiben.'],
];

/**
 * Baut die Datei.
 * @param {object} d {bereich, titel, aktMonat?, konten, buchungen, kosten, einnahmen, toepfe, limits, kategorien?, regeln, soll, kontostaende,
 *   einstellungen?, zusatz?: Array<{name, ...blatt}>}
 * @returns {Uint8Array}
 */
export function dateiBauen(d) {
  const ue = uebersicht(d.titel || 'Unser Geld');
  const vs = vorschau(ue.freiHeuteRef);
  const li = listenBlatt();
  const kat = d.kategorien || KATEGORIEN_START.map((name) => ({ name, aktiv: 'Ja', angelegt: 'Start' }));
  const einst = d.einstellungen || [
    { schluessel: 'Format-Version', wert: String(FORMAT_VERSION), erklaerung: 'Version des Dateiformats' },
    { schluessel: 'Bereich', wert: d.bereich, erklaerung: 'gemeinsam, tagesgeld, natalie oder kai' },
    { schluessel: 'Funktionen', wert: '', erklaerung: 'Eingeschaltete Zusatzfunktionen (durch Komma getrennt). Leer = nur der feste Kern.' },
    { schluessel: 'Widgets', wert: (WIDGETS_START[d.bereich] || []).join(','), erklaerung: 'Reihenfolge der Widgets in der Übersicht' },
    { schluessel: 'Vorher aktiv', wert: (VORHER_AKTIV[d.bereich] || []).join(','), erklaerung: 'Was vor dem Umbau (16.11.2026) an war – für „Wie vor dem Umbau“' },
  ];
  const erk = neuesBlatt('Erklärung', { breiten: [3, 22, 100] });
  ERKLAERUNG.forEach(([a, t], i) => { erk.set('B' + (i + 1), { v: a, s: i === 0 ? STIL.titel : STIL.fett }); if (t) erk.set('C' + (i + 1), { v: t }); });
  const daten = {
    Buchungen: d.buchungen, Konten: d.konten, Kontostände: d.kontostaende, Kosten: d.kosten, Einnahmen: d.einnahmen, Töpfe: d.toepfe,
    Limits: d.limits, 'Soll-Überweisungen': d.soll, Kategorien: kat, Regeln: d.regeln, Einstellungen: einst,
  };
  const fertig = { Übersicht: ue.b, Vorschau: vs.b, Erklärung: erk, Listen: li.b };
  for (const name of GRUNDBLAETTER) if (!fertig[name]) fertig[name] = tabellenBlatt(name, daten[name] || []).b;
  const blaetter = GRUNDBLAETTER.map((n) => fertig[n]);
  for (const z of d.zusatz || []) blaetter.splice(blaetter.length - 2, 0, z);
  return mappeSchreiben({ titel: d.titel, blaetter, namen: [...ue.namen, ...li.namen, ...(d.namen || [])], fremdeStile: d.fremdeStile, metadata: d.metadata });
}

/** Zusatzblatt mit fester Tabelle (kern/format.js ZUSATZ_TABELLEN), z. B. „Offene Punkte“, „TR Buchungen“. */
export const zusatzTabelle = (name, zeilen = []) => tabellenBlatt(name, zeilen).b;
