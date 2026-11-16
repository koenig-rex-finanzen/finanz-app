// Einheitlicher Rechenkern für alle Bereiche (neues Dateiformat, kern/format.js).
// Liest jede Datei gleich und rechnet genau wie die Formeln aus kern/vorlage.js.

import { leseArbeitsmappe } from './xlsx.js';
import { BLAETTER, KOPFZEILE, spalten, RHYTHMUS_ABSTAND, NUR_TOPF, TOPF_PLUS, TOPF_MINUS, SPARPLAN, KARTEN_ARTEN } from './format.js';

export const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const summe = (arr, f = (x) => x) => r2(arr.reduce((s, x) => s + (f(x) || 0), 0));
export const isoAus = (n) => (typeof n === 'number' && isFinite(n) ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : typeof n === 'string' && /^\d{4}-\d\d-\d\d/.test(n) ? n.slice(0, 10) : null);
export const monatPlus = (m, n) => new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)).toISOString().slice(0, 7);
export const monatsende = (m) => new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7), 0)).toISOString().slice(0, 10);

// ---------------------------------------------------------------- Lesen

function umwandeln(typ, v) {
  if (v == null || v === '') return null;
  if (typ === 'd') return isoAus(v);
  if (typ === 'm') return isoAus(v)?.slice(0, 7) ?? null;
  if (typ === 'z' || typ === 'n') return typeof v === 'number' ? v : Number(String(v).replace(',', '.')) || null;
  if (typ === 'j') return v === true || v === 'Ja' ? 'Ja' : v === false || v === 'Nein' ? 'Nein' : String(v).trim() || null;
  if (typ === 'f') return v;
  return String(v).trim() || null;
}

/** Zeilen eines Blatts als Objekte (Spalten über die Überschrift gefunden). Jede Zeile kennt ihre Zeilennummer (_zeile). */
export function tabelleLesen(mappe, name) {
  const b = mappe.blaetter.get(name);
  if (!b) return [];
  const sps = spalten(name);
  const kopf = new Map();
  for (let c = 0; c < 40; c++) {
    const buchst = c < 26 ? String.fromCharCode(65 + c) : 'A' + String.fromCharCode(65 + c - 26);
    const t = b.zellen.get(buchst + KOPFZEILE);
    if (typeof t === 'string') kopf.set(t.trim(), buchst);
  }
  const fehlt = sps.filter((s) => !kopf.has(s.name)).map((s) => s.name);
  if (fehlt.length) throw new Error(`Blatt „${name}“: Spalte(n) ${fehlt.join(', ')} fehlen.`);
  const out = [];
  for (let r = KOPFZEILE + 1; r <= b.zeilen; r++) {
    const z = { _zeile: r };
    let leer = true;
    for (const s of sps) {
      const v = umwandeln(s.typ, b.zellen.get(kopf.get(s.name) + r));
      z[s.key] = v;
      if (v != null && s.typ !== 'f') leer = false;
    }
    if (!leer) out.push(z);
  }
  return out;
}

/** Liest eine Datei im neuen Format. */
export async function leseDatei(daten) {
  const mappe = await leseArbeitsmappe(daten);
  const t = (n) => tabelleLesen(mappe, n);
  const einst = Object.fromEntries(t('Einstellungen').map((e) => [e.schluessel, e.wert ?? '']));
  const ue = mappe.blaetter.get('Übersicht')?.zellen;
  const zus = (n) => (mappe.blaetter.has(n) ? tabelleLesen(mappe, n) : []);
  return {
    bereich: einst.Bereich || null, einstellungen: einst,
    aufteilung: mappe.blaetter.has('Aufteilung') ? leseRaster(mappe.blaetter.get('Aufteilung')) : {},
    erinnerungen: zus('Erinnerungen'), offenePunkte: zus('Offene Punkte'),
    konten: t('Konten'), buchungen: t('Buchungen'), kontostaende: t('Kontostände'), kosten: t('Kosten'), einnahmen: t('Einnahmen'),
    toepfe: t('Töpfe'), limits: t('Limits'), soll: t('Soll-Überweisungen'), kategorien: t('Kategorien'), regeln: t('Regeln'),
    excel: { monat: isoAus(ue?.get('C4'))?.slice(0, 7) || null },
    blaetter: [...mappe.blaetter.keys()],
  };
}

// ---------------------------------------------------------------- Rechnen (wie die Formeln)

const eigen = (m, name) => m.konten.find((k) => k.name === name && k.eigen === 'Ja');
const zaehltFrei = (m, name) => !name || m.konten.some((k) => k.name === name && k.eigen === 'Ja' && k.frei === 'Ja');
const bewegtGeld = (b) => !NUR_TOPF.includes(b.art);

/** Stand eines Kontos bis einschließlich `bis` (wie Konten!Stand bzw. Kontostände!Stand laut Datei). */
export function kontostand(m, name, bis) {
  const k = eigen(m, name);
  if (!k || k.anfang == null) return null;
  const im = (b) => b.datum && b.datum > (k.stichtag || '') && b.datum <= bis && bewegtGeld(b);
  return r2(k.anfang + summe(m.buchungen.filter((b) => im(b) && b.nach === name), (b) => b.betrag) - summe(m.buchungen.filter((b) => im(b) && b.von === name), (b) => b.betrag));
}

const monateZwischen = (a, b) => (+b.slice(0, 4) - +a.slice(0, 4)) * 12 + +b.slice(5, 7) - +a.slice(5, 7);
/** Wirkung einer Buchung auf ihren Topf (nur über die Art). */
export const topfWirkung = (b) => (TOPF_PLUS.includes(b.art) ? b.betrag : TOPF_MINUS.includes(b.art) ? -b.betrag : 0);

/** Stände aller Töpfe (wie Töpfe!Stand aktuell). akt = 'JJJJ-MM', heute = 'JJJJ-MM-TT'. */
export function toepfeStand(m, akt, heute) {
  const raten = summe(m.toepfe, (t) => t.sparrate);
  const sparplan = summe(m.buchungen.filter((b) => b.topf === SPARPLAN && b.datum && b.datum <= heute), topfWirkung);
  return m.toepfe.map((t) => {
    const bis = t.bis && t.bis < akt ? t.bis : akt;
    const zugeteilt = t.automatisch === 'Ja' && t.ab ? r2((t.sparrate || 0) * Math.max(0, monateZwischen(t.ab, bis) + 1)) : 0;
    const bewegt = summe(m.buchungen.filter((b) => b.topf === t.name && b.datum && b.datum > (t.stichtag || '') && b.datum <= heute), topfWirkung);
    const ausSparplan = raten ? r2(((t.sparrate || 0) / raten) * sparplan) : 0;
    return { ...t, zugeteilt, bewegt, ausSparplan, stand: r2((t.anfang || 0) + zugeteilt + bewegt + ausSparplan) };
  });
}

const abstand = (p) => RHYTHMUS_ABSTAND[p.rhythmus] || 1;
/** Plan eines Postens im Monat (wie Kosten!Plan Monat). */
export function plan(p, monat) {
  if (p.status === 'beendet' || (p.ab && monat < p.ab) || (p.bis && monat > p.bis)) return 0;
  const n = abstand(p);
  if (n === 1) return p.betrag || 0;
  if (!p.faellig) return 0;
  return (((+monat.slice(5, 7) - p.faellig) % n) + n) % n === 0 ? p.betrag || 0 : 0;
}
/** Gebucht auf einen Posten: Buchungen mit diesem Posten, bei Ausgaben auch mit gleichnamiger Kategorie (z. B. „Sprit“). */
const gebucht = (m, posten, monat, art) => summe(m.buchungen.filter((b) => b.giltFuer === monat && b.art === art
  && (b.posten === posten || (art === 'Ausgabe' && b.kategorie === posten && b.posten !== posten))), (b) => b.betrag);
const istKarte = (m, konto) => m.konten.some((k) => k.name === konto && KARTEN_ARTEN.includes(k.art));
/**
 * Offener Rest eines Postens im Monat (0, wenn das Konto nicht für Frei zählt).
 * karte = true: Kosten einer Karte (Kreditkarte/Haushalt) zählen als ein Topf: Plan × (1 − mit der Karte ausgegeben / Plan der Karte).
 */
export function offen(m, p, monat, art, karte = true) {
  if (!zaehltFrei(m, p.konto)) return 0;
  const pl = plan(p, monat);
  if (karte && art === 'Ausgabe' && p.konto && istKarte(m, p.konto)) {
    const planKarte = summe(m.kosten.filter((q) => q.konto === p.konto), (q) => plan(q, monat));
    const ausgegeben = summe(m.buchungen.filter((b) => b.von === p.konto && b.giltFuer === monat && b.art === 'Ausgabe'), (b) => b.betrag);
    return planKarte > 0 ? r2(pl * Math.max(0, 1 - ausgegeben / planKarte)) : 0;
  }
  return r2(Math.max(0, pl - gebucht(m, p.posten, monat, art)));
}
/** Soll-Überweisungen eines Monats: Wirkung aufs Freie (+ kommt rein, − geht raus). */
export function sollOffen(m, monat) {
  return summe(m.soll.filter((s) => s.monat === monat), (s) => {
    const ist = summe(m.buchungen.filter((b) => b.von === s.von && b.nach === s.nach && b.giltFuer === monat), (b) => b.betrag);
    const rest = Math.max(0, (s.soll || 0) - ist);
    const w = zaehltFrei(m, s.nach) && s.nach ? 1 : zaehltFrei(m, s.von) && s.von ? -1 : 0;
    return rest * w;
  });
}

/**
 * Fester Kern: Kontostände, Töpfe, Was ist frei?
 * @param {object} m aus leseDatei()
 * @param {string} heute 'JJJJ-MM-TT'
 * @param {string} [akt] 'JJJJ-MM' (Standard: Monat von heute)
 */
export function kern(m, heute, akt = heute.slice(0, 7)) {
  const konten = m.konten.filter((k) => k.eigen === 'Ja' && k.status !== 'geschlossen')
    .map((k) => ({ ...k, standHeute: kontostand(m, k.name, heute) }));
  const toepfe = toepfeStand(m, akt, heute);
  const aufKonten = summe(konten.filter((k) => k.frei === 'Ja'), (k) => k.standHeute);
  const inToepfen = summe(toepfe.filter((t) => t.frei !== 'Ja'), (t) => t.stand);
  const freiHeute = r2(aufKonten - inToepfen);
  const einnahmenOffen = summe(m.einnahmen, (p) => offen(m, p, akt, 'Einnahme'));
  const kostenOffen = summe(m.kosten, (p) => offen(m, p, akt, 'Ausgabe'));
  const soll = sollOffen(m, akt);
  return { akt, heute, konten, toepfe, aufKonten, inToepfen, freiHeute, einnahmenOffen, kostenOffen, soll, frei: r2(freiHeute + einnahmenOffen - kostenOffen + soll) };
}

/** Vorschau über 12 Monate (wie Blatt „Vorschau“). Sparraten nur für Töpfe mit „Automatisch zuteilen“, ab dem 2. Monat. */
export function vorschau(m, heute, akt = heute.slice(0, 7)) {
  const k = kern(m, heute, akt);
  let lauf = k.freiHeute;
  return Array.from({ length: 12 }, (_, i) => {
    const mon = monatPlus(akt, i);
    const einnahmen = summe(m.einnahmen, (p) => offen(m, p, mon, 'Einnahme'));
    const kosten = summe(m.kosten, (p) => offen(m, p, mon, 'Ausgabe', i === 0));
    const sparraten = i === 0 ? 0 : summe(m.toepfe.filter((t) => t.automatisch === 'Ja' && t.frei !== 'Ja' && (!t.ab || t.ab <= mon) && (!t.bis || t.bis >= mon)), (t) => t.sparrate);
    const soll = sollOffen(m, mon);
    const ergebnis = r2(einnahmen - kosten - sparraten + soll);
    const anfang = lauf; lauf = r2(lauf + ergebnis);
    return { monat: mon, einnahmen, kosten, sparraten, soll, ergebnis, freiAnfang: anfang, freiEnde: lauf, luecke: lauf < 0 };
  });
}

/** Limits im Monat: ausgegeben (Ausgaben mit dieser Kategorie, „Gilt für“ = Monat) und Rest. Nur Warnung. */
export function limits(m, monat) {
  return m.limits.filter((l) => (!l.ab || l.ab <= monat) && (!l.bis || l.bis >= monat)).map((l) => {
    const ist = summe(m.buchungen.filter((b) => b.kategorie === l.kategorie && b.giltFuer === monat && b.art === 'Ausgabe'), (b) => b.betrag);
    const anteil = l.betrag ? ist / l.betrag : null;
    return { ...l, ist, rest: r2((l.betrag || 0) - ist), anteil, status: anteil == null ? null : anteil >= 1 ? 'erreicht' : anteil >= 0.8 ? 'fast' : 'ok' };
  });
}

// ---------------------------------------------------------------- Aufteilung (Raster aus Excel, nur gelesen)

const SPN = (i) => (i < 26 ? String.fromCharCode(65 + i) : String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26)));
/** Monatsraster: Zeile mit „Monat“ in B und Datumswerten daneben; Bezeichnung in B. */
function leseRaster(b) {
  const z = b.zellen; let kopf = -1;
  for (let r = 1; r <= b.zeilen; r++) if (z.get('B' + r) === 'Monat' && typeof z.get('D' + r) === 'number') { kopf = r; break; }
  if (kopf < 0) return {};
  const monate = [];
  for (let c = 2; c < 40; c++) { const v = z.get(SPN(c) + kopf); if (typeof v !== 'number') break; monate.push([SPN(c), isoAus(v).slice(0, 7)]); }
  const raster = Object.fromEntries(monate.map(([, m]) => [m, {}]));
  for (let r = kopf + 1; r <= b.zeilen; r++) {
    const l = z.get('B' + r); if (typeof l !== 'string') continue;
    for (const [c, m] of monate) { const v = z.get(c + r); if (typeof v === 'number') raster[m][l.trim()] = v; }
  }
  return raster;
}

// ---------------------------------------------------------------- Wer überweist was (Gemeinsam + Tagesgeld)

const istAusgleich = (b) => /^ausgleich/i.test(b?.gezahltVon || '');
/** Konto mit einer ID (z. B. KT01 = Gemeinschaftskonto, wie kGem in der Aufteilung). */
export const kontoMitId = (m, id) => m.konten.find((k) => k.id === id) || null;
/**
 * Überweisungen von Kai/Natalie an Gemeinschaftskonto und Tagesgeld laut Aufteilung, mit Stand aus den Buchungen.
 * @param {object} g Gemeinsam (leseDatei)  @param {object|null} t Tagesgeld (leseDatei)
 */
export function werUeberweist(g, t, monat) {
  const a = g.aufteilung?.[monat];
  if (!a) return [];
  const wert = (anf) => { const k = Object.keys(a).find((x) => x.startsWith(anf)); return k ? r2(a[k]) : 0; };
  const gk = kontoMitId(g, 'KT01')?.name || 'Gemeinschaftskonto';
  const tg = t?.konten.find((k) => k.eigen === 'Ja' && k.rolle === 'Sparen' && k.status !== 'geschlossen')?.name || 'Tagesgeld';
  const privat = (m, person) => m.konten.find((k) => k.rolle === 'privat' && k.inhaber === person)?.name || null;
  const zahler = (m, b) => (istAusgleich(b) ? null : b.gezahltVon || m.konten.find((k) => k.name === b.von && k.rolle === 'privat')?.inhaber || null);
  const eingezahlt = (m, person, nach) => (m ? summe(m.buchungen.filter((b) => b.giltFuer === monat && b.nach === nach && zahler(m, b) === person
    && !(m === t && b.topf && b.topf !== SPARPLAN)), (b) => b.betrag) : 0);
  const tagN = Math.max(0, ...g.einnahmen.filter((e) => e.gehoertZu === 'Natalie' && (!e.ab || e.ab <= monat) && (!e.bis || e.bis >= monat) && e.tag).map((e) => e.tag));
  const wannN = tagN > 1 ? `wenn ihr Geld da ist, um den ${tagN}.` : 'bis zum 1.';
  return [
    { person: 'Kai', von: privat(g, 'Kai'), nach: gk, betrag: r2(wert('Kai → Gemeinschaftskonto') + wert('Kai: schon direkt')), tag: 1, wann: 'bis zum 1.', m: g },
    { person: 'Kai', von: privat(g, 'Kai'), nach: tg, betrag: wert('Kai → Tagesgeld'), tag: 1, wann: 'bis zum 1.', m: t },
    { person: 'Natalie', von: privat(g, 'Natalie'), nach: tg, betrag: wert('Natalie → Tagesgeld'), tag: tagN || 15, wann: wannN, hinweis: 'Sparraten für die Töpfe', m: t },
    { person: 'Natalie', von: privat(g, 'Natalie'), nach: gk, betrag: wert('Natalie → Gemeinschaftskonto'), tag: tagN || 15, wann: wannN, m: g },
  ].filter((u) => u.betrag > 0).map(({ m, ...u }) => {
    const gebucht = eingezahlt(m, u.person, u.nach);
    return { ...u, gebucht, offen: r2(Math.max(0, u.betrag - gebucht)), erledigt: gebucht >= u.betrag - 0.005 };
  });
}

// ---------------------------------------------------------------- Was noch kommt / Kategorien / Töpfe

/** Offene Posten des Monats (Einnahmen, Kosten, Soll-Überweisungen) nach Tag. Nur Konten, die für Frei zählen. */
export function wasKommt(m, monat, heute) {
  const tagHeute = heute.slice(0, 7) === monat ? +heute.slice(8, 10) : heute.slice(0, 7) > monat ? 99 : 0;
  const liste = [
    ...m.einnahmen.map((p) => ({ art: 'einnahme', titel: p.posten, tag: p.tag || null, betrag: offen(m, p, monat, 'Einnahme') })),
    ...m.kosten.map((p) => ({ art: 'kosten', titel: p.posten, tag: p.tag || null, betrag: -offen(m, p, monat, 'Ausgabe') })),
    ...m.soll.filter((s) => s.monat === monat).map((s) => {
      const ist = summe(m.buchungen.filter((b) => b.von === s.von && b.nach === s.nach && b.giltFuer === monat), (b) => b.betrag);
      const rest = r2(Math.max(0, (s.soll || 0) - ist)); const rein = zaehltFrei(m, s.nach) && s.nach;
      return { art: 'ueberweisung', titel: rein ? `Überweisung von ${s.von}` : `Überweisung an ${s.nach}`, tag: null, betrag: rein ? rest : -rest };
    }),
  ].filter((x) => Math.abs(x.betrag) > 0.005);
  for (const x of liste) x.ueberfaellig = x.tag != null && x.tag < tagHeute;
  return liste.sort((a, b) => (a.tag ?? 99) - (b.tag ?? 99));
}

/** Ausgaben nach Kategorie im Monat („Gilt für“), größte zuerst. Ohne Kategorie = „Ohne Kategorie“. */
export function kategorienMonat(m, monat) {
  const map = new Map();
  for (const b of m.buchungen.filter((x) => x.art === 'Ausgabe' && x.giltFuer === monat && !istAusgleich(x))) {
    const k = b.kategorie || 'Ohne Kategorie'; map.set(k, r2((map.get(k) || 0) + b.betrag));
  }
  return [...map].map(([kategorie, betrag]) => ({ kategorie, betrag })).sort((a, b) => b.betrag - a.betrag);
}

/** Buchungen eines Topfs mit Wirkung und Stand danach (alt → neu). Rechnet wie toepfeStand. */
export function topfVerlauf(m, name, akt, heute) {
  const t = toepfeStand(m, akt, heute).find((x) => x.name === name); if (!t) return null;
  const raten = summe(m.toepfe, (x) => x.sparrate); const anteil = raten ? (t.sparrate || 0) / raten : 0;
  const ein = [];
  for (const b of m.buchungen) {
    if (!b.datum || b.datum > heute) continue;
    if (b.topf === SPARPLAN && anteil) ein.push({ b, wirkung: r2(topfWirkung(b) * anteil), sparplan: true });
    else if (b.topf === name && b.datum > (t.stichtag || '')) ein.push({ b, wirkung: topfWirkung(b) });
  }
  ein.sort((x, y) => x.b.datum.localeCompare(y.b.datum));
  let lauf = (t.anfang || 0) + t.zugeteilt;
  for (const e of ein) { lauf = r2(lauf + e.wirkung); e.danach = lauf; }
  return { topf: t, eintraege: ein.reverse() };
}

/** Nächste freie Nummer (häufigster Buchstabe der Datei, sonst B). */
export function neueNr(m) {
  const z = {};
  for (const b of m.buchungen) { const p = (b.nr || '').match(/^([A-Z])\d+$/)?.[1]; if (p) z[p] = (z[p] || 0) + 1; }
  const pre = Object.entries(z).sort((x, y) => y[1] - x[1])[0]?.[0] || 'B';
  const max = Math.max(0, ...m.buchungen.map((b) => +((b.nr || '').match(new RegExp(`^${pre}(\\d+)$`))?.[1] || 0)));
  return pre + String(max + 1).padStart(4, '0');
}

const KAT_REGELN = [
  ['Sprit', /\b(sprit|tank(en|stelle\w*)?|getankt|benzin|diesel|aral|shell|esso|agip|avia|totalenergies|jet tank\w*|orlen)\b/i],
  ['Bargeld', /bargeld|geldautomat|viacash/i], ['Amazon', /amazon|amzn/i], ['Lieferdienst', /lieferando|wolt|uber\s?eats|domino|flink|pizza/i],
  ['Drogerie', /\bdm\b|rossmann|müller drogerie|drogerie/i],
  ['Handwerksbedarf', /\bobi\b|hornbach|bauhaus|toom|hagebau|globus baumarkt|baumarkt|werkzeug/i],
  ['Abos', /anthropic|openai|chatgpt|netflix|spotify|disney|dazn|youtube|apple\.com|google\s?(one|play)/i],
  ['Auto', /kfz|werkstatt|atu\b|parkhaus|parken/i], ['Gesundheit', /apotheke|arzt|zahn|optik|physio/i],
  ['Essen & Trinken', /netto|rewe|edeka|aldi|lidl|kaufland|penny|bäcker|baecker|metzger|restaurant|café|cafe|imbiss|döner/i],
  ['Kinder', /kita|spielzeug|mytoys|kinder/i], ['Geschenke', /geschenk/i], ['Haushalt', /ikea|haushalt/i],
  ['Shopping', /zalando|h&m|new yorker|primark|kleidung/i], ['Freizeit', /kino|schwimmbad|zoo|museum|eventim|steam|playstation|nintendo/i],
];
/**
 * Kategorie-Vorschlag: 1. eigene Regeln (Blatt „Regeln“), 2. was zuletzt beim selben Gegenüber gewählt wurde, 3. feste Wörter.
 * @param {object} b {von, nach, beschreibung, bankKategorie, art}
 */
export function kategorieVorschlag(m, b, kategorien = []) {
  if (!b || b.art !== 'Ausgabe') return null;
  const text = [b.nach, b.beschreibung, b.bankKategorie].filter(Boolean).join(' ');
  for (const r of m.regeln || []) if (r.stichwort && r.kategorie && text.toLowerCase().includes(String(r.stichwort).toLowerCase())) return r.kategorie;
  const g = (b.nach || '').trim().toLowerCase();
  if (g.length > 2) {
    const fr = m.buchungen.filter((x) => x.kategorie && x.nr !== b.nr && (x.nach || '').trim().toLowerCase() === g).sort((x, y) => (y.datum || '').localeCompare(x.datum || ''))[0];
    if (fr) return fr.kategorie;
  }
  const erlaubt = new Set(kategorien.length ? kategorien : KAT_REGELN.map(([k]) => k));
  for (const [k, re] of KAT_REGELN) if (erlaubt.has(k) && re.test(text)) return k;
  return null;
}
