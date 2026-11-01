// Macht aus der gelesenen Arbeitsmappe ein sauberes Datenmodell.
// Spalten werden über ihre Überschrift gefunden, nicht über feste Buchstaben.
// So bricht nichts, wenn in Excel eine Spalte eingefügt wird.

import { leseArbeitsmappe } from './xlsx.js';

const SPALTEN = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
const spaltenName = (i) => (i < 26 ? SPALTEN[i] : SPALTEN[Math.floor(i / 26) - 1] + SPALTEN[i % 26]);

/** Excel-Seriennummer → 'JJJJ-MM-TT' */
export function excelDatum(n) {
  if (typeof n !== 'number' || !isFinite(n)) return null;
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return d.toISOString().slice(0, 10);
}
const monatVon = (datum) => (datum ? datum.slice(0, 7) : null);
const zahl = (v) => (typeof v === 'number' ? v : v == null ? null : Number(String(v).replace(/\./g, '').replace(',', '.')) || null);
const text = (v) => (v == null ? null : String(v).trim() || null);

function tabelle(blatt, ersteSpalte, { maxKopfZeile = 10 } = {}) {
  // Kopfzeile finden: die Zeile, deren erste belegte Zelle = ersteSpalte ist
  const z = blatt.zellen;
  let kopf = -1;
  for (let r = 1; r <= maxKopfZeile && kopf < 0; r++) {
    for (let c = 0; c < 4; c++) {
      if (z.get(spaltenName(c) + r) === ersteSpalte) { kopf = r; break; }
    }
  }
  if (kopf < 0) throw new Error(`Überschrift „${ersteSpalte}“ nicht gefunden.`);
  const spalten = new Map();
  for (let c = 0; c < 40; c++) {
    const t = z.get(spaltenName(c) + kopf);
    if (typeof t === 'string') spalten.set(t.trim(), spaltenName(c));
  }
  const zeilen = [];
  for (let r = kopf + 1; r <= blatt.zeilen; r++) {
    const zeile = { _zeile: r };
    for (const [name, buchst] of spalten) zeile[name] = z.get(buchst + r) ?? null;
    zeilen.push(zeile);
  }
  return zeilen;
}

function blatt(mappe, name) {
  const b = mappe.blaetter.get(name);
  if (!b) throw new Error(`Blatt „${name}“ fehlt in der Datei.`);
  return b;
}

/**
 * Liest die Buchhaltungsdatei in ein Modell.
 * @param {ArrayBuffer|Uint8Array} daten Inhalt von Buchhaltung_Kai_Natalie.xlsx
 */
export async function leseBuchhaltung(daten) {
  const mappe = await leseArbeitsmappe(daten);

  const konten = tabelle(blatt(mappe, 'Konten'), 'Konto')
    .filter((r) => /^KT\d+/.test(r['ID'] || ''))
    .map((r) => ({
      name: text(r['Konto']), id: r['ID'], bank: text(r['Bank']), art: text(r['Art']),
      inhaber: text(r['Inhaber']), rolle: text(r['Rolle']), imTool: r['Im Tool'] === 'ja',
      csv: r['CSV-Import'] === 'ja', status: text(r['Status']), ibanEnde: text(r['IBAN-Ende']),
      anfang: zahl(r['Anfangsbestand']), stichtag: excelDatum(r['Stichtag']), bemerkung: text(r['Bemerkung']),
    }));

  const kosten = tabelle(blatt(mappe, 'Kosten'), 'ID')
    .filter((r) => /^K\d+/.test(r['ID'] || ''))
    .map((r) => ({
      id: r['ID'], posten: text(r['Posten']), bereich: text(r['Bereich']), betrag: zahl(r['Betrag']),
      rhythmus: text(r['Rhythmus']), faellig: text(r['Fällig']), monatswert: zahl(r['Monatswert']),
      konto: text(r['Bezahlt aus (Konto)']), topf: text(r['Topf']), vertragAuf: text(r['Vertrag auf']),
      status: text(r['Status']), ab: monatVon(excelDatum(r['Ab (Monat)'])), bis: monatVon(excelDatum(r['Bis (Monat)'])),
      bemerkung: text(r['Bemerkung']), tag: zahl(r['Abbuchungstag']),
    }));

  const einnahmen = tabelle(blatt(mappe, 'Einnahmen'), 'ID')
    .filter((r) => /^E\d+/.test(r['ID'] || ''))
    .map((r) => ({
      id: r['ID'], posten: text(r['Posten']), bereich: text(r['Bereich']), betrag: zahl(r['Betrag']),
      eingang: text(r['Eingang (wann)']), konto: text(r['Konto (Eingang)']), giltFuer: text(r['Gilt für']),
      ab: monatVon(excelDatum(r['Ab (Monat)'])), bis: monatVon(excelDatum(r['Bis (Monat)'])),
      status: text(r['Status']), bemerkung: text(r['Bemerkung']), tag: zahl(r['Tag (verfügbar ab)']),
      gehoertZu: text(r['Gehört zu']),
    }));

  const toepfeRoh = tabelle(blatt(mappe, 'Töpfe'), 'Topf');
  const toepfe = [];
  for (const r of toepfeRoh) {
    const name = text(r['Topf']);
    if (!name) continue;
    if (name === 'Summe') break;
    toepfe.push({
      name, konto: text(r['Konto']), bereich: text(r['Bereich']), sparrate: zahl(r['Sparrate / Monat']) || 0,
      bedarf: zahl(r['Bedarf / Monat']) || 0, anfang: zahl(r['Anfangsbestand']) || 0,
      stichtag: excelDatum(r['Stichtag']), excelStand: zahl(r['Stand aktuell']), bemerkung: text(r['Bemerkung']),
    });
  }

  const buchungen = tabelle(blatt(mappe, 'Buchungen'), 'Nr')
    .filter((r) => /^B\d+/.test(r['Nr'] || '') && zahl(r['Betrag']) != null)
    .map((r) => ({
      nr: r['Nr'], datum: excelDatum(r['Datum (Zahlung)']), giltFuer: monatVon(excelDatum(r['Gilt für Monat'])),
      art: text(r['Art']), von: text(r['Von']), nach: text(r['Nach']), betrag: zahl(r['Betrag']),
      bereich: text(r['Bereich']), posten: text(r['Posten']), topf: text(r['Topf']),
      beschreibung: text(r['Beschreibung / Verwendungszweck']), kommentar: text(r['Kommentar']),
      quelle: text(r['Quelle']), importId: text(r['Import-ID']), gezahltVon: text(r['Gezahlt von']),
      excelZuordnung: text(r['Zuordnung']),
    }));

  const kontostaendeBank = tabelle(blatt(mappe, 'Kontostände'), 'Monat')
    .filter((r) => typeof r['Monat'] === 'number' && r['Konto'])
    .map((r) => ({
      monat: monatVon(excelDatum(r['Monat'])), konto: text(r['Konto']), bank: zahl(r['Ist-Stand laut Bank']),
      kommentar: text(r['Kommentar']),
    }));

  // Aufteilung: Monatsraster (Monate in einer Zeile nebeneinander, Bezeichnung in Spalte B)
  const aufteilung = leseRaster(blatt(mappe, 'Aufteilung'));

  // Offene Punkte
  let offenePunkte = [];
  try {
    offenePunkte = tabelle(blatt(mappe, 'Offene Punkte'), 'Nr')
      .filter((r) => r['Nr'] != null && r['Was ist offen'])
      .map((r) => ({ nr: r['Nr'], thema: text(r['Thema']), text: text(r['Was ist offen']), erledigt: r['Erledigt'] === 'ja' }));
  } catch { /* Blatt optional */ }

  // Was Excel gerade auf „Start“ ausgerechnet hat – zum Abgleich
  const start = blatt(mappe, 'Start').zellen;
  const excelMonat = monatVon(excelDatum(start.get('C4')));
  const excelKontostaende = {};
  for (let r = 30; r <= 40; r++) {
    const name = start.get('B' + r);
    if (typeof name === 'string' && konten.some((k) => k.name === name)) excelKontostaende[name] = zahl(start.get('C' + r));
  }

  return { konten, kosten, einnahmen, toepfe, buchungen, kontostaendeBank, aufteilung, offenePunkte,
    excel: { monat: excelMonat, kontostaende: excelKontostaende } };
}

function leseRaster(b) {
  // Sucht die Zeile mit „Monat“ in Spalte B, auf die mehrere Datumswerte folgen.
  const z = b.zellen;
  let kopf = -1;
  for (let r = 1; r <= b.zeilen; r++) {
    if (z.get('B' + r) === 'Monat' && typeof z.get('D' + r) === 'number') { kopf = r; break; }
  }
  if (kopf < 0) return {};
  const monate = [];
  for (let c = 2; c < 40; c++) {
    const v = z.get(spaltenName(c) + kopf);
    if (typeof v !== 'number') break;
    monate.push([spaltenName(c), monatVon(excelDatum(v))]);
  }
  const raster = {};
  for (const [, m] of monate) raster[m] = {};
  for (let r = kopf + 1; r <= b.zeilen; r++) {
    const label = z.get('B' + r);
    if (typeof label !== 'string') continue;
    for (const [buchst, m] of monate) {
      const v = z.get(buchst + r);
      if (typeof v === 'number') raster[m][label.trim()] = v;
    }
  }
  return raster;
}
