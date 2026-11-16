// Einheitlicher Rechenkern für alle Bereiche (neues Dateiformat, kern/format.js).
// Liest jede Datei gleich und rechnet genau wie die Formeln aus kern/vorlage.js.

import { leseArbeitsmappe } from './xlsx.js';
import { BLAETTER, KOPFZEILE, spalten, RHYTHMUS_ABSTAND, NUR_TOPF, TOPF_PLUS, TOPF_MINUS, SPARPLAN } from './format.js';

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
  return {
    bereich: einst.Bereich || null, einstellungen: einst,
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
const gebucht = (m, posten, monat, art) => summe(m.buchungen.filter((b) => b.posten === posten && b.giltFuer === monat && b.art === art), (b) => b.betrag);
/** Offener Rest eines Postens im Monat (0, wenn das Konto nicht für Frei zählt). */
export function offen(m, p, monat, art) {
  if (!zaehltFrei(m, p.konto)) return 0;
  return r2(Math.max(0, plan(p, monat) - gebucht(m, p.posten, monat, art)));
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
    const kosten = summe(m.kosten, (p) => offen(m, p, mon, 'Ausgabe'));
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
