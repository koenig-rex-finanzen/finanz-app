// Kais privater Bereich („Mein Geld“): liest Kai_Privat.xlsx und rechnet wie die Formeln der Datei.
// Ohne Browser, ohne Fremdbibliotheken. Hier stehen keine persönlichen Daten.
// Quelle der Wahrheit ist die Datei; der Kern muss dieselben Werte liefern (test/kai.mjs).

import { leseArbeitsmappe } from './xlsx.js';

const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const summe = (arr, f = (x) => x) => r2(arr.reduce((s, x) => s + (Number(f(x)) || 0), 0));
const excelDatum = (n) => (typeof n === 'number' && isFinite(n) ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : null);
const mon = (n) => excelDatum(n)?.slice(0, 7) || null;
const txt = (v) => (v == null ? null : String(v).trim() || null);
const zahl = (v) => (typeof v === 'number' ? v : v == null || v === '' ? null : Number(String(v).replace(/\./g, '').replace(',', '.')));
export const monatPlus = (m, n) => { const d = new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)); return d.toISOString().slice(0, 7); };
const monateZwischen = (a, b) => (+b.slice(0, 4) - +a.slice(0, 4)) * 12 + (+b.slice(5, 7) - +a.slice(5, 7));
export const serial = (iso) => { const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d || 1) - Date.UTC(1899, 11, 30)) / 864e5; };

export const ARTEN = ['Einnahme', 'Ausgabe', 'In Topf legen', 'Aus Topf nehmen', 'Umbuchung'];
export const UMBUCHUNG = 'Umbuchung'; // von Konto (E) nach Konto (F), z. B. c24 → Spartopf; ändert keine Töpfe
export const POSTEN_EXTRA = ['Überweisung Gemeinschaftskonto', 'Trade Republic (extra)', 'Sparraten-Erstattung (alt)', 'Sonstiges'];
export const GK_POSTEN = 'Überweisung Gemeinschaftskonto';
export const RUECKLAGE = 'Rücklage Lücke';
/** „Wofür“ (Buchungen!Q): wofür das Geld ausgegeben wurde – unabhängig von Töpfen und Posten, nur zum Auswerten. */
export const WOFUER = ['Freizeit', 'Essen & Trinken', 'Lieferdienst', 'Shopping', 'Amazon', 'Haushalt', 'Handwerksbedarf', 'Auto', 'Kinder', 'Abos', 'Geschenke', 'Gesundheit', 'Bargeld', 'Sonstiges'];
export const WOFUER_SPALTE = 'Q';
// Zeilenbereiche der Datei (so angelegt, Formeln dort vorbelegt)
export const BEREICHE = { buchungen: [5, 1004], kosten: [5, 24], einnahmen: [5, 12], toepfe: [5, 14], gemeinsam: [5, 18], depot: [17, 56], tr: [5, 2004], offen: [5, 200] };

/** Liest Kai_Privat.xlsx in ein Modell (Zeilennummern bleiben für das Schreiben erhalten). */
export async function leseKai(daten) {
  const mappe = await leseArbeitsmappe(daten);
  const bl = (n) => { const b = mappe.blaetter.get(n); if (!b) throw new Error(`Blatt „${n}“ fehlt in Kai_Privat.xlsx.`); return b.zellen; };
  const zeilen = (name, [von, bis], fn, pflicht) => {
    const z = bl(name); const out = [];
    for (let r = von; r <= bis; r++) { if (z.get(pflicht + r) == null) continue; out.push({ zeile: r, ...fn((c) => z.get(c + r)) }); }
    return out;
  };
  const K = bl('Konten');
  const kontoZeile = (r) => ({ zeile: r, id: txt(K.get('A' + r)), name: txt(K.get('B' + r)), bank: txt(K.get('C' + r)), rolle: txt(K.get('D' + r)),
    anfang: zahl(K.get('E' + r)), stichtag: excelDatum(K.get('F' + r)), endung: txt(K.get('I' + r)), excelStand: zahl(K.get('G' + r)) });
  const konto = { ...kontoZeile(5), name: txt(K.get('B5')) || 'Kai c24' };
  const konten = [konto];
  for (let r = 6; r <= 12; r++) { const x = kontoZeile(r); if (x.name) konten.push(x); }
  // Pocket „Spartopf“: echtes Konto mit Rolle „Sparen“; dort liegen die Töpfe mit „Zählt als frei = Nein“.
  const spartopf = konten.find((x) => x.rolle === 'Sparen') || null;
  const buchungen = zeilen('Buchungen', BEREICHE.buchungen, (c) => ({
    nr: txt(c('A')), datum: excelDatum(c('B')), giltFuer: mon(c('C')), art: txt(c('D')), konto: txt(c('E')), gegen: txt(c('F')),
    betrag: zahl(c('G')) || 0, kategorie: txt(c('H')), posten: txt(c('I')), topf: txt(c('J')), beschreibung: txt(c('K')),
    kommentar: txt(c('L')), quelle: txt(c('M')), importId: txt(c('N')), wofuer: txt(c(WOFUER_SPALTE)),
  }), 'A').filter((b) => b.nr);
  const kosten = zeilen('Kosten', BEREICHE.kosten, (c) => ({
    id: txt(c('A')), posten: txt(c('B')), bereich: txt(c('C')), betrag: zahl(c('D')) || 0, rhythmus: txt(c('E')) || 'monatlich',
    faellig: zahl(c('F')), tag: zahl(c('G')), konto: txt(c('H')), ab: mon(c('I')), bis: mon(c('J')), bemerkung: txt(c('N')),
  }), 'B');
  const einnahmen = zeilen('Einnahmen', BEREICHE.einnahmen, (c) => ({
    id: txt(c('A')), posten: txt(c('B')), betrag: zahl(c('C')) || 0, rhythmus: txt(c('D')) || 'monatlich', faellig: zahl(c('E')),
    eingang: txt(c('F')), konto: txt(c('G')), ab: mon(c('H')), bis: mon(c('I')), bemerkung: txt(c('M')),
  }), 'B');
  const toepfe = zeilen('Töpfe', BEREICHE.toepfe, (c) => ({
    id: txt(c('A')), name: txt(c('B')), sparrate: zahl(c('C')) || 0, ab: mon(c('D')), bis: mon(c('E')), anfang: zahl(c('F')) || 0,
    frei: txt(c('J')) === 'Ja', bemerkung: txt(c('K')), excelStand: zahl(c('I')),
  }), 'B');
  const gemeinsam = zeilen('Gemeinsam', BEREICHE.gemeinsam, (c) => ({ monat: mon(c('B')), soll: zahl(c('C')), bemerkung: txt(c('F')) }), 'B');
  const depot = zeilen('Depot', BEREICHE.depot, (c) => ({ position: txt(c('A')), name: txt(c('B')), art: txt(c('C')), kurs: zahl(c('L')), kursVom: excelDatum(c('M')) || txt(c('M')) }), 'A');
  const tr = zeilen('TR Buchungen', BEREICHE.tr, (c) => ({
    datum: excelDatum(c('A')), typ: txt(c('B')), name: txt(c('C')), isin: txt(c('D')), position: txt(c('E')), stueck: zahl(c('F')),
    kurs: zahl(c('G')), betrag: zahl(c('H')) || 0, gebuehr: zahl(c('I')), steuer: zahl(c('J')), beschreibung: txt(c('K')), id: txt(c('L')),
  }), 'B');
  const offen = zeilen('Offene Punkte', BEREICHE.offen, (c) => ({ nr: c('A'), punkt: txt(c('B')), status: txt(c('C')) || 'offen', notiz: txt(c('D')) }), 'B');
  const S = bl('Start');
  const V = bl('Vorschau');
  const sp = 'CDEFGHIJKLMN'.split('');
  const excel = {
    aktMonat: mon(S.get('C4')), kontostand: zahl(K.get('G5')), spartopf: spartopf?.excelStand ?? null, nochAufsPocket: zahl(bl('Töpfe').get('I17')), verplant: zahl(S.get('C11')), ruecklage: zahl(S.get('C12')), frei: zahl(S.get('C14')),
    status: txt(S.get('C17')), fehlbetrag: zahl(S.get('C19')), behalten: zahl(S.get('C20')), variabel: zahl(S.get('C22')),
    vorschau: sp.map((c) => ({ monat: mon(V.get(c + '4')), ergebnis: zahl(V.get(c + '61')), freiEnde: zahl(V.get(c + '62')), behalten: zahl(V.get(c + '64')), variabel: zahl(V.get(c + '65')), status: txt(V.get(c + '66')) })),
  };
  return {
    konto, konten, spartopf, buchungen, kosten, einnahmen, toepfe, gemeinsam, depot, tr, offen,
    start: { rechnenAb: mon(S.get('D4')), planMonat: mon(S.get('C5')), horizont: zahl(S.get('C6')) || 3 }, excel,
  };
}

// ---------------------------------------------------------------- Rechnen (wie die Formeln)

/** Wirkung aufs Konto in Spalte E (wie Buchungen!O). Umbuchung: geht von E ab (und kommt bei F an, siehe standVon). */
export const kontoeffekt = (b) => (b.art === 'Einnahme' ? b.betrag : b.art === 'Ausgabe' || b.art === UMBUCHUNG ? -b.betrag : 0);
/** Wie Buchungen!P. Eine Umbuchung ändert keinen Topf. */
export const topfeffekt = (b) => (!b.topf || b.art === UMBUCHUNG ? 0 : b.art === 'Ausgabe' || b.art === 'Aus Topf nehmen' ? -b.betrag : b.betrag);
/** Wirkung einer Buchung auf ein bestimmtes Konto (für Anzeige und Abgleich). */
export const wirkungAuf = (b, name) => r2((b.konto === name ? kontoeffekt(b) : 0) + (b.art === UMBUCHUNG && b.gegen === name ? b.betrag : 0));

/** Plan eines Postens im Monat (Kosten L / Vorschau). */
export function plan(p, m) {
  if ((p.ab && m < p.ab) || (p.bis && m > p.bis)) return 0;
  if (p.rhythmus === 'jährlich') return Number(m.slice(5, 7)) === Number(p.faellig) ? p.betrag : 0;
  return p.betrag;
}
const ist = (k, posten, m, art) => summe(k.buchungen.filter((b) => b.posten === posten && b.giltFuer === m && b.art === art), (b) => b.betrag);

/** Stand eines Kontos (wie Konten!G): Anfangsbestand + Buchungen nach dem Stichtag dieses Kontos. */
export function standVon(k, kt) {
  if (!kt || kt.anfang == null) return null;
  const nach = (b) => b.datum && (!kt.stichtag || b.datum > kt.stichtag);
  return r2(kt.anfang + summe(k.buchungen.filter(nach), (b) => wirkungAuf(b, kt.name)));
}
export const kontostand = (k) => standVon(k, k.konto);

export function toepfeStand(k, akt) {
  return k.toepfe.map((t) => {
    const bis = t.bis && t.bis < akt ? t.bis : akt;
    const zugeteilt = t.ab ? r2(t.sparrate * Math.max(0, monateZwischen(t.ab, bis) + 1)) : 0;
    const bewegt = summe(k.buchungen.filter((b) => b.topf === t.name && b.datum && (!k.konto.stichtag || b.datum > k.konto.stichtag)), topfeffekt);
    return { ...t, zugeteilt, bewegt, stand: r2(t.anfang + zugeteilt + bewegt) };
  });
}

/**
 * Alles, was Start und Vorschau zeigen.
 * @param {object} k aus leseKai()
 * @param {string} heute 'JJJJ-MM-TT'
 */
export function rechneKai(k, heute, { horizont } = {}) {
  const akt = k.start.rechnenAb || heute.slice(0, 7);
  const H = horizont || k.start.horizont || 3;
  const stand = kontostand(k);
  const toepfe = toepfeStand(k, akt);
  const verplant = summe(toepfe.filter((t) => !t.frei), (t) => t.stand);
  const ruecklage = summe(toepfe.filter((t) => t.frei), (t) => t.stand);
  const fehlt = stand == null;
  // Spartopf (Pocket): Geld der Töpfe liegt dort. Frei = c24 + Spartopf − Töpfe. Ohne Spartopf-Stand zählt er mit 0 (wie N() in Excel).
  const sp = k.spartopf || null;
  const spStand = sp ? standVon(k, sp) : null;
  const spartopf = sp ? { name: sp.name, zeile: sp.zeile, stand: spStand, fehlt: spStand == null, anfang: sp.anfang, stichtag: sp.stichtag, verplant, offen: spStand == null ? null : r2(verplant - spStand) } : null;
  const frei = fehlt ? 0 : r2(stand + (spStand || 0) - verplant);

  const monate = Array.from({ length: 12 }, (_, i) => monatPlus(akt, i));
  const vorschau = monate.map((m, i) => {
    const einnahmen = k.einnahmen.map((e) => ({ posten: e.posten, id: e.id, plan: plan(e, m), ist: ist(k, e.posten, m, 'Einnahme') }))
      .map((x) => ({ ...x, rest: r2(Math.max(0, x.plan - x.ist)) }));
    const g = k.gemeinsam.find((x) => x.monat === m);
    const gIst = ist(k, GK_POSTEN, m, 'Ausgabe');
    const gemeinsam = { soll: g?.soll ?? null, ist: gIst, rest: g ? r2(Math.max(0, (g.soll || 0) - gIst)) : 0, fehlt: !g };
    const kosten = k.kosten.map((x) => ({ posten: x.posten, id: x.id, plan: plan(x, m), ist: ist(k, x.posten, m, 'Ausgabe') }))
      .map((x) => ({ ...x, rest: r2(Math.max(0, x.plan - x.ist)) }));
    const sparraten = k.toepfe.map((t) => ({ name: t.name, betrag: t.frei || m <= akt || (t.ab && m < t.ab) || (t.bis && m > t.bis) ? 0 : t.sparrate }));
    const s = { einnahmen: summe(einnahmen, (x) => x.rest), gemeinsam: gemeinsam.rest, kosten: summe(kosten, (x) => x.rest), sparraten: summe(sparraten, (x) => x.betrag) };
    return { monat: m, nr: i + 1, imHorizont: i + 1 <= H, einnahmen, gemeinsam, kosten, sparraten, summen: s, ergebnis: r2(s.einnahmen - s.gemeinsam - s.kosten - s.sparraten) };
  });
  let lauf = frei;
  for (const v of vorschau) { v.freiAnfang = r2(lauf); v.freiEnde = r2(lauf + v.ergebnis); lauf = v.freiEnde; }
  for (let i = 0; i < 12; i++) {
    const v = vorschau[i];
    const spaeter = vorschau.slice(i + 1).filter((x) => x.imHorizont).map((x) => x.freiEnde);
    v.behalten = i + 1 >= H || !spaeter.length ? 0 : r2(Math.max(0, v.freiEnde - Math.min(...spaeter)));
    v.variabel = r2((i === 0 ? v.freiAnfang : vorschau[i - 1].behalten) + v.ergebnis - v.behalten);
    v.status = !v.imHorizont ? 'Ausblick' : v.freiEnde < 0 || v.variabel < 0 ? 'Lücke' : 'OK';
  }
  const imH = vorschau.filter((v) => v.imHorizont);
  const luecke = imH.find((v) => v.status === 'Lücke') || null;
  const fehlbetrag = r2(Math.max(0, -Math.min(...imH.map((v) => v.freiEnde)), -vorschau[0].variabel));
  const behalten = vorschau[0].behalten;
  return {
    akt, horizont: H, fehlt, stand, spartopf, toepfe, verplant, ruecklage, frei, vorschau,
    status: fehlt ? 'Kontostand fehlt' : luecke ? 'Lücke' : 'OK', ersterMinusMonat: luecke?.monat || null,
    fehlbetrag, behalten, inRuecklage: r2(Math.max(0, behalten - ruecklage)), variabel: vorschau[0].variabel,
    stichtag: k.konto.stichtag, letzteBuchung: k.buchungen.reduce((x, b) => (b.datum > x ? b.datum : x), ''),
  };
}

// ---------------------------------------------------------------- Kontostände prüfen

const tagNr = (iso) => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
/** Eigene Konten mit echtem Geld (c24 + Pockets mit Rolle „Sparen“). */
export const eigeneKonten = (k) => (k.konten || [k.konto]).filter((kt) => kt === k.konto || kt.name === k.konto.name || kt.rolle === 'Sparen');

/** Sieht aus wie eine Überweisung zwischen den eigenen Konten, ist aber nicht als Umbuchung eingetragen. */
export function falscheUmbuchungen(k) {
  const namen = eigeneKonten(k).map((kt) => kt.name);
  const kurz = namen.map((n) => n.replace(/^c24\s+/i, '').toLowerCase()).filter((n) => n.length > 3);
  return k.buchungen.filter((b) => (b.art === 'Ausgabe' || b.art === 'Einnahme') && !b.topf && (
    namen.includes(b.gegen) || /umbuchung/i.test(b.kategorie || '') || kurz.some((n) => (b.gegen || '').toLowerCase() === n)));
}

/**
 * Vergleicht den Stand laut Bank mit dem berechneten Stand.
 * @param {object} k aus leseKai()
 * @param {Object<string, number>} bank Kontoname → Stand laut Bank-App
 * @returns {{konten: Array, ok: boolean, falsch: Array}}
 */
export function kontenPruefen(k, bank) {
  const konten = eigeneKonten(k).filter((kt) => bank[kt.name] != null && isFinite(bank[kt.name])).map((kt) => {
    const stand = standVon(k, kt);
    const diff = stand == null ? null : r2(bank[kt.name] - stand);
    const nach = k.buchungen.filter((b) => b.datum && (!kt.stichtag || b.datum > kt.stichtag) && wirkungAuf(b, kt.name) !== 0);
    let verdacht = [];
    if (diff != null && Math.abs(diff) >= 0.005) {
      // 1. Buchung, deren Betrag genau der Differenz entspricht (doppelt drin, Vorzeichen falsch oder fehlt auf dem Auszug)
      const passt = (b) => Math.abs(Math.abs(wirkungAuf(b, kt.name)) - Math.abs(diff)) < 0.005 || Math.abs(Math.abs(wirkungAuf(b, kt.name)) * 2 - Math.abs(diff)) < 0.005;
      const doppelt = new Set();
      for (const b of nach) for (const c of nach) if (b.nr < c.nr && !istTeil(b) && !istTeil(c) && Math.abs(wirkungAuf(b, kt.name) - wirkungAuf(c, kt.name)) < 0.005 && Math.abs(tagNr(b.datum) - tagNr(c.datum)) <= 3) { doppelt.add(b); doppelt.add(c); }
      verdacht = nach.filter((b) => passt(b) || doppelt.has(b)).map((b) => ({ ...b, wirkung: wirkungAuf(b, kt.name),
        grund: [passt(b) ? 'Betrag passt zur Differenz' : null, doppelt.has(b) ? 'gleicher Betrag innerhalb von 3 Tagen' : null].filter(Boolean).join(' · ') }))
        .sort((a, b) => (b.grund.includes('Differenz') - a.grund.includes('Differenz')) || b.datum.localeCompare(a.datum));
      verdacht.push({ hinweis: diff < 0
        ? `Bei der Bank sind ${eur2(-diff)} weniger. Fehlt in der App eine Ausgabe, oder steht eine Einnahme doppelt drin?`
        : `Bei der Bank sind ${eur2(diff)} mehr. Steht in der App eine Ausgabe doppelt drin, oder fehlt eine Einnahme?` });
    }
    return { name: kt.name, zeile: kt.zeile, bank: r2(bank[kt.name]), stand, diff, ok: diff != null && Math.abs(diff) < 0.005, verdacht, stichtag: kt.stichtag };
  });
  return { konten, ok: konten.length > 0 && konten.every((x) => x.ok), falsch: falscheUmbuchungen(k) };
}
const eur2 = (n) => n.toFixed(2).replace('.', ',') + ' €';

// ---------------------------------------------------------------- Wofür

const WOFUER_REGELN = [
  ['Bargeld', /bargeld|geldautomat|viacash/i],
  ['Amazon', /amazon|amzn/i],
  ['Lieferdienst', /lieferando|wolt|uber\s?eats|domino|flink|getir|pizza/i],
  ['Handwerksbedarf', /obi\b|hornbach|bauhaus|toom|hagebau|globus baumarkt|baumarkt|werkzeug|würth|wuerth/i],
  ['Abos', /anthropic|openai|chatgpt|netflix|spotify|disney|dazn|youtube|apple\.com|google\s?(one|play)|lovable|runpod|digitale produkte/i],
  ['Auto', /aral|shell|esso|jet\b|total|tankstelle|tanken|kfz|werkstatt|atu\b|parkhaus|parken|mobilität/i],
  ['Gesundheit', /apotheke|arzt|zahn|gesundheit|optik|physio/i],
  ['Essen & Trinken', /netto|rewe|edeka|aldi|lidl|kaufland|penny|wurst|bäcker|baecker|metzger|restaurant|café|cafe|imbiss|döner|lebensmittel/i],
  ['Kinder', /kita|spielzeug|mytoys|kinder/i],
  ['Geschenke', /geschenk/i],
  ['Haushalt', /dm\b|rossmann|ikea|haushalt|drogerie/i],
  ['Shopping', /shopping|bekleidung|zalando|h&m|new yorker|primark|kleidung/i],
  ['Freizeit', /freizeit|kino|schwimmbad|zoo|museum|tickets?\b|eventim|steam|playstation|nintendo/i],
];
/** Vorschlag fürs „Wofür“ aus Gegenüber, Beschreibung und Bank-Kategorie. Nur für Ausgaben ohne festen Posten. */
export function wofuerVorschlag(b, k) {
  if (!b || b.art !== 'Ausgabe' || b.posten === GK_POSTEN || b.posten === 'Trade Republic (extra)') return null;
  if (k && b.posten && istBudget(k, b.posten)) return b.posten;
  // Gelernt: was zuletzt beim selben Gegenüber gewählt wurde
  const g = (b.gegen || '').trim().toLowerCase();
  if (k && g.length > 2) {
    const frueher = k.buchungen.filter((x) => x.wofuer && x !== b && x.nr !== b.nr && (x.gegen || '').trim().toLowerCase() === g).sort((x, y) => (y.datum || '').localeCompare(x.datum || ''))[0];
    if (frueher) return frueher.wofuer;
  }
  const text = [b.gegen, b.beschreibung, b.kategorie].filter(Boolean).join(' ');
  for (const [w, re] of WOFUER_REGELN) if (re.test(text)) return w;
  return null;
}

/**
 * Ausgaben nach „Wofür“ für die Monate von…bis („Gilt für“). Nur echte Ausgaben der eigenen Konten, keine Umbuchungen.
 * Ohne „Wofür“: nach Posten gruppiert (Gemeinschaftskonto, Trade Republic, Feste Posten), sonst „Nicht zugeordnet“.
 * Buchungen, die eigentlich Umbuchungen sind (falscheUmbuchungen), zählen nicht.
 */
export const WOFUER_EXTRA = ['Gemeinschaftskonto', 'Trade Republic', 'Feste Posten', 'Nicht zugeordnet'];

// ---------------------------------------------------------------- Budgets und Aufteilen

/** Bereich in Kosten!C, der einen Posten zum Monatsbudget macht (z. B. Dampfen). Plan = Budget, Ist = Ausgaben mit diesem Posten. */
export const BUDGET = 'Budget';
export const budgets = (k) => (k.kosten || []).filter((x) => x.posten && x.bereich === BUDGET);
/** Auswahl fürs „Wofür“: erst die Budgets, dann die feste Liste. Ein Budget als „Wofür“ setzt auch den Posten. */
export const wofuerListe = (k) => [...new Set([...budgets(k).map((x) => x.posten), ...WOFUER])];
export const istBudget = (k, name) => budgets(k).some((x) => x.posten === name);

/** Teil-Zeilen einer aufgeteilten Buchung: Quelle „Aufteilung“, Kommentar beginnt mit „Teil von P0001“. */
export const istTeil = (b) => b.quelle === 'Aufteilung' && /^Teil von P\d+/.test(b.kommentar || '');
export const teilVon = (b) => (istTeil(b) ? b.kommentar.match(/^Teil von (P\d+)/)[1] : null);
export function teileVon(k) {
  const m = new Map();
  for (const b of k.buchungen) { const v = teilVon(b); if (v) m.set(v, [...(m.get(v) || []), b]); }
  return m;
}
/**
 * Prüft eine Aufteilung. teile = [{betrag, wofuer}], der erste Teil bleibt in der Original-Zeile.
 * @returns {string|null} Fehlertext
 */
export function aufteilungPruefen(b, teile) {
  if (!b || b.art !== 'Ausgabe') return 'Aufteilen geht nur bei Ausgaben.';
  if (istTeil(b)) return 'Das ist schon ein Teil. Bitte die ursprüngliche Buchung aufteilen.';
  if (teile.length < 2) return 'Bitte mindestens zwei Teile angeben.';
  if (teile.some((t) => !(t.betrag > 0))) return 'Jeder Teil braucht einen Betrag über 0.';
  const sum = r2(teile.reduce((a, t) => a + t.betrag, 0));
  if (Math.abs(sum - b.betrag) >= 0.005) return `Die Teile ergeben ${sum.toFixed(2).replace('.', ',')} €, die Buchung hat ${b.betrag.toFixed(2).replace('.', ',')} €.`;
  return null;
}

const tageImMonat = (m) => new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7), 0)).getUTCDate();
/**
 * Budget im Monat m: Plan, verbraucht, übrig/drüber und wie viel bis heute „dran“ wäre (gleichmäßig über den Monat).
 * heute = 'JJJJ-MM-TT' (nur für den laufenden Monat wichtig).
 */
export function budgetStand(k, x, m, heute) {
  const p = plan(x, m);
  const ausg = k.buchungen.filter((b) => b.posten === x.posten && b.giltFuer === m && b.art === 'Ausgabe');
  const verbraucht = summe(ausg, (b) => b.betrag);
  const lauf = heute && heute.slice(0, 7) === m;
  const anteil = lauf ? Math.min(1, Number(heute.slice(8, 10)) / tageImMonat(m)) : (heute && m < heute.slice(0, 7) ? 1 : 0);
  return { name: x.posten, plan: p, verbraucht, rest: r2(p - verbraucht), drueber: r2(Math.max(0, verbraucht - p)), anteil, sollBisHeute: r2(p * anteil), anzahl: ausg.length, buchungen: ausg };
}

/** Summen je Gruppe (wofuerGruppe) für mehrere Monate: {monate, gruppen: [{name, werte: [..], summe}]} */
export function wofuerVerlauf(k, monate) {
  const namen = new Set(eigeneKonten(k).map((kt) => kt.name));
  const falsch = new Set(falscheUmbuchungen(k));
  const g = new Map();
  for (const b of k.buchungen) {
    if (b.art !== 'Ausgabe' || !namen.has(b.konto) || falsch.has(b) || !monate.includes(b.giltFuer)) continue;
    const n = wofuerGruppe(b, k);
    const x = g.get(n) || { name: n, werte: monate.map(() => 0) };
    const i = monate.indexOf(b.giltFuer); x.werte[i] = r2(x.werte[i] + b.betrag); g.set(n, x);
  }
  return { monate, gruppen: [...g.values()].map((x) => ({ ...x, summe: r2(x.werte.reduce((a, v) => a + v, 0)) })) };
}

/**
 * Referenzwert je Gruppe: Durchschnitt der letzten 3 vollen Monate vor akt, in denen schon mit „Wofür“ gearbeitet wurde.
 * Nur ein Anhaltspunkt, kein Plan. Erst ab 2 solchen Monaten.
 */
export function referenzWerte(k, akt) {
  const vorher = [3, 2, 1].map((i) => monatPlus(akt, -i));
  const genutzt = vorher.filter((m) => k.buchungen.some((b) => b.giltFuer === m && b.wofuer));
  if (genutzt.length < 2) return { monate: genutzt, werte: new Map() };
  const v = wofuerVerlauf(k, genutzt);
  return { monate: genutzt, werte: new Map(v.gruppen.map((x) => [x.name, r2(x.summe / genutzt.length)])) };
}
/** Gruppe einer Ausgabe in der Auswertung: „Wofür“, sonst nach Posten, sonst „Nicht zugeordnet“. */
export function wofuerGruppe(b, k) {
  if (b.wofuer) return b.wofuer;
  if (k && b.posten && istBudget(k, b.posten)) return b.posten;
  if (b.posten === GK_POSTEN) return 'Gemeinschaftskonto';
  if (b.posten === 'Trade Republic (extra)') return 'Trade Republic';
  if (b.posten && b.posten !== 'Sonstiges') return 'Feste Posten';
  return 'Nicht zugeordnet';
}
export function wofuerAuswertung(k, von, bis) {
  const namen = new Set(eigeneKonten(k).map((kt) => kt.name));
  const falsch = new Set(falscheUmbuchungen(k));
  const liste = k.buchungen.filter((b) => b.art === 'Ausgabe' && namen.has(b.konto) && !falsch.has(b) && b.giltFuer && b.giltFuer >= von && b.giltFuer <= bis);
  const gruppen = new Map();
  for (const b of liste) {
    const g = wofuerGruppe(b, k);
    const x = gruppen.get(g) || { name: g, summe: 0, anzahl: 0 };
    x.summe = r2(x.summe + b.betrag); x.anzahl++; gruppen.set(g, x);
  }
  const rang = (n) => WOFUER_EXTRA.indexOf(n);
  const reihen = [...gruppen.values()].sort((a, b) => rang(a.name) - rang(b.name) || b.summe - a.summe);
  return { von, bis, gesamt: summe(liste, (b) => b.betrag), reihen, ohne: liste.filter((b) => wofuerGruppe(b, k) === 'Nicht zugeordnet').length, nichtGezaehlt: [...falsch].filter((b) => b.giltFuer >= von && b.giltFuer <= bis) };
}

/** Plan/Ist eines Monats (wie Start, Abschnitt Plan/Ist). */
export function planIst(k, m) {
  const ein = { plan: summe(k.einnahmen, (e) => plan(e, m)), ist: summe(k.buchungen.filter((b) => b.giltFuer === m && b.art === 'Einnahme'), (b) => b.betrag) };
  const g = k.gemeinsam.find((x) => x.monat === m);
  const gem = { plan: g?.soll || 0, ist: ist(k, GK_POSTEN, m, 'Ausgabe') };
  const kos = { plan: summe(k.kosten, (x) => plan(x, m)), ist: summe(k.kosten, (x) => ist(k, x.posten, m, 'Ausgabe')) };
  const spar = { plan: summe(k.toepfe, (t) => t.sparrate), ist: -summe(k.buchungen.filter((b) => b.giltFuer === m && b.art === 'Ausgabe'), topfeffekt) };
  const ausgaben = summe(k.buchungen.filter((b) => b.giltFuer === m && b.art === 'Ausgabe'), (b) => b.betrag);
  const variabel = { plan: null, ist: r2(ausgaben - gem.ist - kos.ist - spar.ist) };
  return { einnahmen: ein, gemeinsam: gem, kosten: kos, sparraten: spar, variabel };
}

// ---------------------------------------------------------------- Gilt für, Regeln, c24-CSV

/** Gehalt ab dem 20. gilt für den Folgemonat, sonst der Monat des Datums. */
export function giltFuer(datum, posten) {
  if (/gehalt/i.test(posten || '') && Number(datum.slice(8, 10)) >= 20) return monatPlus(datum.slice(0, 7), 1);
  return datum.slice(0, 7);
}

function zerlegen(z) {
  const out = []; let f = ''; let q = false;
  for (let i = 0; i < z.length; i++) {
    const c = z[i];
    if (q) { if (c === '"') { if (z[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { out.push(f); f = ''; } else f += c;
  }
  out.push(f); return out;
}
const deDatum = (s) => { const m = (s || '').match(/(\d{2})\.(\d{2})\.(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : null; };
const deBetrag = (s) => { const t = String(s || '').replace(/[€\s ]/g, '').replace(/\./g, '').replace(',', '.'); return t ? Number(t) : NaN; };

/** Liest einen c24-Export (Transaktionen › Export › CSV). */
export function leseC24(text) {
  const zeilen = text.replace(/^﻿/, '').replace(/\r/g, '').split('\n').filter((z) => z.trim());
  const kopfIdx = zeilen.findIndex((z) => /Transaktionstyp/i.test(z) && /Buchungsdatum/i.test(z));
  if (kopfIdx < 0) throw new Error('Das ist kein c24-Export (Spalte „Transaktionstyp“ fehlt).');
  const kopf = zerlegen(zeilen[kopfIdx]).map((s) => s.trim());
  const i = (n) => kopf.indexOf(n);
  const daten = [];
  for (const z of zeilen.slice(kopfIdx + 1)) {
    const f = zerlegen(z);
    const datum = deDatum(f[i('Buchungsdatum')]);
    const betrag = r2(deBetrag(f[i('Betrag')]));
    if (!datum || !isFinite(betrag)) continue;
    const g = (n) => (i(n) >= 0 ? (f[i(n)] || '').trim() : '');
    daten.push({ datum, betrag, typ: g('Transaktionstyp'), karteneinsatz: g('Karteneinsatz'), empfaenger: g('Zahlungsempfänger'),
      iban: g('IBAN'), bic: g('BIC'), zweck: g('Verwendungszweck'), beschreibung: g('Beschreibung'), kategorie: g('Kategorie'), unterkategorie: g('Unterkategorie'),
      bargeld: g('Bargeldabhebung'), kontonummer: g('Kontonummer'), kontoname: g('Kontoname') });
  }
  return daten;
}

/** SHA-1 als Hex (WebCrypto, im Browser und in Node vorhanden). */
async function sha1(text) {
  const h = await globalThis.crypto.subtle.digest('SHA-1', new TextEncoder().encode(text));
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, '0')).join('');
}
/** Import-IDs genau wie beim Anlegen der Datei: c24- + SHA-1 (12 Zeichen) inkl. laufender Nummer bei gleichem Schlüssel. */
export async function importIds(zeilen) {
  const zaehler = new Map();
  for (const z of zeilen) {
    const schluessel = `c24|${z.datum}|${z.betrag.toFixed(2)}|${z.empfaenger}|${z.zweck}|${z.karteneinsatz}`;
    const n = (zaehler.get(schluessel) || 0) + 1; zaehler.set(schluessel, n);
    z.id = 'c24-' + (await sha1(`${schluessel}|${n}`)).slice(0, 12);
  }
  return zeilen;
}

/**
 * Feste Regeln (Kais Bau-Skript) + eigene Regeln „Text enthält … → Posten/Topf/Gilt für“.
 * @param {object} z Bankzeile
 * @param {object} kk Kai-Modell
 * @param {{gkEndung?:string, regeln?:Array}} opt gkEndung: Kontonummer-Endung des Gemeinschaftskontos (aus der gemeinsamen Datei)
 */
export function vorschlag(z, kk, { gkEndung, regeln = [] } = {}) {
  const text = `${z.empfaenger} ${z.zweck} ${z.beschreibung}`;
  let art = z.betrag > 0 ? 'Einnahme' : 'Ausgabe';
  let posten = null, topf = null, gegen = z.empfaenger || 'Extern', folgemonat = false, konto = kk.konto.name;
  const sp = kk.spartopf;
  if (sp && aufSpartopf(z, sp)) konto = sp.name; // Zeile gehört zum Pocket selbst (z. B. Zinsen)
  else if (sp && ((sp.endung && (z.iban || '').replace(/\s/g, '').endsWith(sp.endung))
    || (/umbuchung/i.test(`${z.kategorie} ${z.unterkategorie}`) && z.empfaenger && z.empfaenger.trim().length > 3 && sp.name.toLowerCase().includes(z.empfaenger.trim().toLowerCase())))) {
    // Überweisung zwischen c24 und Spartopf
    return { datum: z.datum, giltFuer: z.datum.slice(0, 7), art: UMBUCHUNG, konto: z.betrag < 0 ? kk.konto.name : sp.name, gegen: z.betrag < 0 ? sp.name : kk.konto.name,
      betrag: r2(Math.abs(z.betrag)), kategorie: [z.kategorie, z.unterkategorie].filter(Boolean).join(' / ') || null, posten: null, topf: null,
      beschreibung: [z.zweck || z.typ, z.empfaenger].filter(Boolean).join(' · ').slice(0, 250), kommentar: null, quelle: 'CSV', importId: z.id };
  }
  for (const r of regeln) if (r.text && text.toLowerCase().includes(r.text.toLowerCase())) { posten = r.posten || posten; topf = r.topf || topf; folgemonat = !!r.folgemonat; break; }
  if (!posten) {
    const rate = kk.kosten.find((k) => k.posten === 'PayPal-Rate');
    if (art === 'Einnahme' && /gehalt|lohn/i.test(text)) { posten = 'Gehalt Kai'; }
    else if (/anthropic/i.test(text)) posten = 'Claude-Abo';
    else if (rate && /paypal/i.test(text) && Math.abs(Math.abs(z.betrag) - rate.betrag) < 0.005) posten = 'PayPal-Rate';
    else if (art === 'Ausgabe' && gkEndung && (z.iban || '').replace(/\s/g, '').endsWith(gkEndung)) { posten = GK_POSTEN; gegen = 'Gemeinschaftskonto'; }
    else if (/trade\s?republic/i.test(text)) posten = 'Trade Republic (extra)';
    else {
      const p = [...kk.kosten, ...kk.einnahmen].find((x) => x.posten && x.posten.length > 3 && text.toLowerCase().includes(x.posten.toLowerCase()));
      if (p) posten = p.posten;
    }
  }
  const datum = z.datum;
  const gf = folgemonat ? monatPlus(datum.slice(0, 7), 1) : giltFuer(datum, posten);
  const v = {
    datum, giltFuer: gf, art, konto, gegen, betrag: r2(Math.abs(z.betrag)),
    kategorie: [z.kategorie, z.unterkategorie].filter(Boolean).join(' / ') || null, posten, topf,
    beschreibung: [z.zweck || [z.typ, z.karteneinsatz].filter(Boolean).join(' '), z.empfaenger].filter(Boolean).join(' · ').slice(0, 250),
    kommentar: null, quelle: 'CSV', importId: z.id,
  };
  v.wofuer = wofuerVorschlag(v, kk);
  if (v.wofuer && !v.posten && istBudget(kk, v.wofuer)) v.posten = v.wofuer;
  return v;
}

/** Gehört die Exportzeile zum Pocket selbst (Kontonummer/Kontoname im Export)? */
export const aufSpartopf = (z, sp) => Boolean(sp && ((sp.endung && (z.kontonummer || '').replace(/\s/g, '').endsWith(sp.endung)) || (z.kontoname && sp.name && sp.name.toLowerCase().includes(z.kontoname.toLowerCase()) && z.kontoname.length > 3)));
const tag = (iso) => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
/**
 * Abgleich c24-Export mit der Datei. Schreibt nichts.
 * Status: importiert (Import-ID da) · ergaenzen (von Hand eingetragen: gleicher Betrag ±3 Tage, nur Import-ID ergänzen)
 *         · pruefen (unklar, nie automatisch) · neu
 */
export async function abgleichC24(zeilen, kk, opt = {}) {
  await importIds(zeilen);
  const ids = new Set(kk.buchungen.map((b) => b.importId).filter(Boolean));
  const benutzt = new Set();
  const sp = kk.spartopf;
  const c24 = kk.konto.name;
  const teile = teileVon(kk);
  // Aufgeteilte Buchungen zählen beim Abgleich mit ihrer ursprünglichen Summe; die Teil-Zeilen selbst nie.
  const signiertAuf = (b, name) => (b.art === 'Einnahme' || b.art === 'Ausgabe' || b.art === UMBUCHUNG) && (b.konto === name || b.gegen === name) && !istTeil(b)
    ? r2(wirkungAuf(b, name) + (teile.get(b.nr) || []).reduce((a, t) => a + wirkungAuf(t, name), 0)) : null;
  return zeilen.map((z) => {
    if (ids.has(z.id)) return { z, status: 'importiert' };
    const eigen = sp && aufSpartopf(z, sp) ? sp.name : c24;
    // Pocket-Zeile, die eine Überweisung von/zum c24 spiegelt: steht schon als Umbuchung (oder kommt mit der c24-Zeile)
    if (eigen !== c24) {
      const spiegel = zeilen.some((y) => y !== z && !aufSpartopf(y, sp) && Math.abs(y.betrag + z.betrag) < 0.005 && Math.abs(tag(y.datum) - tag(z.datum)) <= 1)
        || kk.buchungen.some((b) => b.art === UMBUCHUNG && b.datum && Math.abs(wirkungAuf(b, sp.name) - z.betrag) < 0.005 && Math.abs(tag(b.datum) - tag(z.datum)) <= 3);
      if (spiegel) return { z, status: 'gegenbuchung' };
    }
    const signiert = (b) => signiertAuf(b, eigen);
    const kand = kk.buchungen.filter((b) => !benutzt.has(b) && !b.importId && signiert(b) != null && Math.abs(signiert(b) - z.betrag) < 0.005 && b.datum && Math.abs(tag(b.datum) - tag(z.datum)) <= 3);
    const csvGleich = kk.buchungen.filter((b) => b.importId && signiert(b) != null && b.datum === z.datum && Math.abs(signiert(b) - z.betrag) < 0.005);
    const v = vorschlag(z, kk, opt);
    if (kand.length === 1 && !csvGleich.length) { benutzt.add(kand[0]); return { z, status: 'ergaenzen', treffer: kand[0], vorschlag: v }; }
    if (kand.length > 1 || csvGleich.length) return { z, status: 'pruefen', treffer: kand[0] || csvGleich[0], grund: kand.length > 1 ? 'mehrere passende Buchungen' : 'gleicher Tag und Betrag schon aus einem Export da', vorschlag: v };
    return { z, status: 'neu', vorschlag: v };
  });
}

// ---------------------------------------------------------------- Trade Republic

/** Depot wie Blatt „Depot“ (Formeln Zeile 6–13 und 17 ff.). */
export function depot(k) {
  const tr = k.tr;
  const s = (pos, typ, feld) => { const x = tr.filter((t) => t.position === pos && t.typ === typ).reduce((y, t) => y + (Number(t[feld]) || 0), 0); return feld === 'stueck' ? r2x(x) : r2(x); };
  const positionen = k.depot.map((d) => {
    const stueck = r2x(s(d.position, 'Kauf', 'stueck') + s(d.position, 'Verkauf', 'stueck') + s(d.position, 'Split', 'stueck'));
    const gekauft = r2(-(s(d.position, 'Kauf', 'betrag') + s(d.position, 'Zeichnung', 'betrag') + s(d.position, 'Geschenk', 'betrag')));
    const verkauftStueck = -s(d.position, 'Verkauf', 'stueck');
    const erloes = s(d.position, 'Verkauf', 'betrag');
    const einstand = stueck + verkauftStueck <= 0 ? 0 : gekauft * stueck / (stueck + verkauftStueck);
    const letzte = [...tr].filter((t) => t.position === d.position && t.kurs != null).sort((a, b) => (a.datum || '').localeCompare(b.datum || '')).pop();
    const wert = d.kurs != null ? r2(stueck * d.kurs) : null;
    const einstandR = r2(einstand);
    return { ...d, stueck, gekauft, verkauftStueck, erloes, einstand: r2(einstand), realisiert: r2(erloes - (gekauft - einstand)), dividenden: s(d.position, 'Dividende', 'betrag'),
      kaufkurs: stueck > 0 ? einstand / stueck : null, wert, gv: wert == null ? null : r2(wert - einstandR), gvProzent: wert == null || !einstandR ? null : (wert - einstandR) / einstandR,
      letzterKurs: letzte?.kurs ?? null, letzterKursAm: letzte?.datum ?? null };
  });
  const offen = positionen.filter((p) => p.stueck > 1e-9);
  return {
    positionen, bargeld: summe(tr, (t) => t.betrag), eingesetzt: summe(positionen, (p) => p.einstand),
    wert: summe(positionen.filter((p) => p.wert != null), (p) => p.wert), gv: summe(positionen.filter((p) => p.gv != null), (p) => p.gv),
    dividenden: summe(tr.filter((t) => t.typ === 'Dividende'), (t) => t.betrag), realisiert: summe(positionen, (p) => p.realisiert),
    eingezahlt: summe(tr.filter((t) => t.typ === 'Einzahlung'), (t) => t.betrag), ausgezahlt: -summe(tr.filter((t) => t.typ === 'Auszahlung'), (t) => t.betrag),
    stand: tr.reduce((x, t) => (t.datum > x ? t.datum : x), ''), offen: offen.length,
  };
}
const r2x = (n) => Math.round(n * 1e8) / 1e8;

const TR_TYP = {
  TRANSFER_INBOUND: 'Einzahlung', TRANSFER_INSTANT_INBOUND: 'Einzahlung', CUSTOMER_INPAYMENT: 'Einzahlung', TRANSFER_INSTANT_OUTBOUND: 'Auszahlung',
  TRANSFER_OUTBOUND: 'Auszahlung', REVERSE_SPLIT: 'Split', IPO_SUBSCRIPTION: 'Zeichnung', STOCKPERK: 'Geschenk', TAX_OPTIMIZATION: 'Steuer',
  BUY: 'Kauf', SELL: 'Verkauf', SAVINGS_PLAN: 'Kauf', DIVIDEND: 'Dividende',
};
export function trTyp(z) {
  const t = (z.type || '').toUpperCase(); const d = z.description || '';
  if (TR_TYP[t]) return TR_TYP[t];
  if (/reverse_split/i.test(d)) return 'Split';
  if (/IPO access/i.test(d)) return 'Zeichnung';
  if (/stockperk/i.test(d)) return 'Geschenk';
  if (/tax optimi/i.test(d)) return 'Steuer';
  if (/dividend/i.test(d + t)) return 'Dividende';
  if (/^sell|sell trade/i.test(d) || /SELL/.test(t)) return 'Verkauf';
  if (/buy trade|savings plan|^buy/i.test(d) || /BUY|SAVINGS/.test(t)) return 'Kauf';
  if (/incoming|top up|inpayment/i.test(d) || /INBOUND|INPAYMENT/.test(t)) return 'Einzahlung';
  if (/outgoing/i.test(d) || /OUTBOUND/.test(t)) return 'Auszahlung';
  return null;
}

/** TR-Transaktionsexport lesen: CSV oder Text aus einer .xlsx mit der ganzen CSV in Spalte A. */
export function leseTr(text) {
  const zeilen = text.replace(/^﻿/, '').replace(/\r/g, '').split('\n').filter((z) => z.trim());
  const kopfIdx = zeilen.findIndex((z) => /(^|,)datetime,/.test(z) && /transaction_id/.test(z));
  if (kopfIdx < 0) throw new Error('Das ist kein Trade-Republic-Transaktionsexport (Spalte „transaction_id“ fehlt).');
  const kopf = zerlegen(zeilen[kopfIdx]).map((s) => s.trim());
  const out = [];
  for (const z of zeilen.slice(kopfIdx + 1)) {
    const f = zerlegen(z); const o = {};
    kopf.forEach((k, i) => { o[k] = (f[i] ?? '').trim(); });
    if (!o.transaction_id) continue;
    out.push(o);
  }
  return out;
}
/** Text aus einer TR-.xlsx (ganze CSV in Spalte A). */
export async function trTextAusXlsx(bytes) {
  const m = await leseArbeitsmappe(bytes);
  const b = [...m.blaetter.values()][0];
  const zeilen = [];
  for (let r = 1; r <= b.zeilen; r++) { const v = b.zellen.get('A' + r); if (v != null) zeilen.push(String(v)); }
  return zeilen.join('\n');
}

/** Neue TR-Zeilen (noch nicht in der Datei) als Zeilen fürs Blatt „TR Buchungen“. */
export function trNeu(roh, k) {
  const da = new Set(k.tr.map((t) => t.id));
  const num = (s) => (s === '' || s == null ? null : Number(s));
  const posVon = new Map(k.tr.filter((t) => t.isin).map((t) => [t.isin, t.position || t.isin]));
  const neu = [];
  const unbekannt = [];
  for (const z of roh) {
    if (da.has(z.transaction_id)) continue;
    const typ = trTyp(z);
    if (!typ) { unbekannt.push(z); continue; }
    const betrag = r2((num(z.amount) || 0) + (num(z.fee) || 0) + (num(z.tax) || 0));
    const stueck = ['Kauf', 'Verkauf', 'Split'].includes(typ) ? num(z.shares) : typ === 'Dividende' ? num(z.shares) : null;
    neu.push({ datum: (z.date || z.datetime || '').slice(0, 10), typ, name: z.name || null, isin: z.symbol || null, position: z.symbol || null,
      stueck: typ === 'Verkauf' && stueck > 0 ? -stueck : stueck, kurs: typ === 'Kauf' || typ === 'Verkauf' ? num(z.price) : null,
      betrag, gebuehr: num(z.fee), steuer: num(z.tax), beschreibung: z.description || null, id: z.transaction_id, _zeit: z.datetime });
  }
  // Reverse-Split: alle ISINs eines Split-Vorgangs (gleiche Zeit) bekommen die schon bekannte Position
  const gruppen = new Map();
  for (const n of neu.filter((x) => x.typ === 'Split')) { const g = gruppen.get(n._zeit) || []; g.push(n); gruppen.set(n._zeit, g); }
  for (const g of gruppen.values()) {
    const bekannt = g.map((n) => posVon.get(n.isin)).find(Boolean) || g.find((n) => (n.stueck || 0) < 0)?.isin;
    if (bekannt) for (const n of g) { n.position = bekannt; posVon.set(n.isin, bekannt); }
  }
  for (const n of neu) if (n.isin && posVon.has(n.isin)) n.position = posVon.get(n.isin);
  neu.sort((a, b) => (a._zeit || a.datum).localeCompare(b._zeit || b.datum));
  const positionenNeu = [...new Set(neu.map((n) => n.position).filter((p) => p && !k.depot.some((d) => d.position === p)))]
    .map((p) => { const n = neu.find((x) => x.position === p); return { position: p, name: n?.name || p, art: /krypto|crypto/i.test(n?.beschreibung || '') ? 'Krypto' : null }; });
  return { neu: neu.map(({ _zeit, ...r }) => r), unbekannt, positionenNeu };
}
