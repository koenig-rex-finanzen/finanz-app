// CSV-Import der VW-Bank-Exporte: lesen, Konto erkennen, Buchungen vorschlagen, Doppelte erkennen.
// Reine Funktionen ohne Browser – läuft auch in Node (z. B. später für Jarvis).
// Wichtig: Hier stehen keine persönlichen Daten (keine IBANs, keine Namen). Konten werden über
// die Kontonummer-Endungen aus der Excel-Datei erkannt, persönliche Regeln kommen aus einer Regeldatei.

import { kontostand, SPRIT_MUSTER } from './rechnen.js';

const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const tag = (iso) => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
const monatPlus = (m, n) => { const d = new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)); return d.toISOString().slice(0, 7); };
const deDatum = (s) => { const m = (s || '').match(/(\d{2})\.(\d{2})\.(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : null; };
const deZahl = (s) => { const t = String(s || '').trim(); if (!t) return 0; return Number(t.replace(/\./g, '').replace(',', '.')) || 0; };
const endung = (s) => (String(s || '').match(/(\d{4})\D*$/) || [])[1] || null;

// ---------------------------------------------------------------- Lesen

function zeileZerlegen(z) {
  const out = []; let feld = ''; let inQ = false;
  for (let i = 0; i < z.length; i++) {
    const c = z[i];
    if (inQ) { if (c === '"') { if (z[i + 1] === '"') { feld += '"'; i++; } else inQ = false; } else feld += c; }
    else if (c === '"') inQ = true;
    else if (c === ';') { out.push(feld); feld = ''; }
    else feld += c;
  }
  out.push(feld);
  return out;
}

/** Text einer CSV-Datei (Bytes) – UTF-8, sonst Windows-1252. */
export function csvText(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let t = new TextDecoder('utf-8').decode(u8);
  if (t.includes('�')) t = new TextDecoder('windows-1252').decode(u8);
  return t.replace(/^﻿/, '');
}

/** Liest einen VW-Bank-Export. */
export function leseCsv(text, dateiname = '') {
  const zeilen = text.replace(/\r/g, '').split('\n');
  const kopf = zeilen.slice(0, 7).map(zeileZerlegen);
  const finde = (name) => kopf.find((z) => (z[0] || '').trim().toLowerCase().startsWith(name));
  const inhaber = (finde('kontoinhaber') || [])[1] || '';
  // Kontozeile: normalerweise Zeile 3; sonst die erste Kopfzeile mit „Nr. 123456…“
  const zeileText = (z) => (z || []).filter(Boolean).join(' ').trim();
  const kontoText = /Nr\.?\s*\d{6,}/.test(zeileText(kopf[2])) ? zeileText(kopf[2]) : zeileText(kopf.find((z) => /Nr\.?\s*\d{6,}/.test(zeileText(z)))) || zeileText(kopf[2]);
  const kontoNr = (kontoText.match(/Nr\.?\s*(\d{6,})/) || dateiname.match(/KtoNr_(\d+)/) || [])[1] || '';
  const saldo = deZahl((finde('saldo') || [])[1]);
  const zr = ((finde('zeitraum') || [])[1] || '').match(/(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})/);
  // Spaltenkopf: mit oder ohne Anführungszeichen, auch mit Leerzeichen/BOM davor
  const kopfIdx = zeilen.findIndex((z) => /^\s*"?Nr\.?"?\s*;\s*"?Buchungsdatum"?\s*;/i.test(z.replace(/^\uFEFF/, '')));
  if (kopfIdx < 0) throw new Error(`„${dateiname || 'Datei'}“ ist kein VW-Bank-Export (Spaltenkopf fehlt).`);
  const spalten = zeileZerlegen(zeilen[kopfIdx]).map((s) => s.trim());
  const i = (n) => spalten.indexOf(n);
  const daten = [];
  for (const z of zeilen.slice(kopfIdx + 1)) {
    if (!z.trim()) continue;
    const f = zeileZerlegen(z);
    const datum = deDatum(f[i('Buchungsdatum')]);
    if (!datum) continue;
    const info = (f[i('Umsatzinformation')] || '').replace(/\s+/g, ' ').trim();
    const betrag = r2(deZahl(f[i('Haben (EUR)')]) - deZahl(f[i('Soll (EUR)')]));
    const karte = info.match(/\*{4}\d{4}\s+(\d{2}\.\d{2}\.\d{4})/);
    daten.push({
      nr: Number(f[i('Nr.')]) || daten.length + 1, datum, wertstellung: deDatum(f[i('Wertstellung')]),
      umsatzart: (f[i('Umsatzart')] || '').trim(), info, betrag,
      referenz: (f[i('Referenznummer')] || '').trim(), mandat: (f[i('Mandat ID')] || '').trim(),
      kartendatum: karte ? deDatum(karte[1]) : null,
    });
  }
  return { dateiname, inhaber, kontoText, kontoNr, saldo, von: zr ? deDatum(zr[1]) : null, bis: zr ? deDatum(zr[2]) : null, zeilen: daten };
}

// ---------------------------------------------------------------- Konten erkennen

/** Konto der Excel-Datei zu einem Export. */
export function kontoErkennen(csv, modell) {
  const e = endung(csv.kontoNr);
  let k = modell.konten.find((x) => endung(x.ibanEnde) && endung(x.ibanEnde) === e);
  if (!k && /visa|kreditkarte/i.test(`${csv.kontoText} ${csv.dateiname || ''}`)) k = modell.konten.find((x) => /kredit|visa/i.test((x.art || '') + x.name));
  return k || null;
}

const BIC_BANK = { DEFFDEFF: /c\s?24/i };

/** Gegenkonto aus dem Buchungstext (eigene Konten über Kontonummer-Endung), sonst Extern. */
function gegenkonto(info, modell, eigenesKonto, importNummern) {
  const extern = modell.konten.find((x) => x.rolle === 'extern')?.name || 'Extern';
  const kandidaten = [];
  for (const m of info.matchAll(/IBAN:\s*([A-Z]{2}\d{2}[A-Z0-9]{8,30})/g)) kandidaten.push(m[1]);
  for (const m of info.matchAll(/Kto:\s*(\d{5,})/g)) kandidaten.push(m[1]);
  for (const c of kandidaten) {
    const k = modell.konten.find((x) => endung(x.ibanEnde) && c.endsWith(endung(x.ibanEnde)));
    if (k && k.name !== eigenesKonto) return k.name;
  }
  // Kartennummer/Kontonummer eines anderen importierten Kontos im Text (z. B. Visa-Ausgleich)
  for (const [nr, name] of importNummern) if (name !== eigenesKonto && nr.length >= 10 && info.includes(nr)) return name;
  // Bank über BIC (z. B. C24), nur wenn genau ein privates Konto dieser Bank existiert
  const bic = (info.match(/BIC:\s*([A-Z0-9]{8,11})/) || [])[1];
  if (bic && BIC_BANK[bic]) {
    const k = modell.konten.filter((x) => BIC_BANK[bic].test(x.bank || '') && x.name !== eigenesKonto);
    if (k.length === 1) return k[0].name;
  }
  return extern;
}

// ---------------------------------------------------------------- Vorschläge

/** Kurzer, lesbarer Name aus dem Buchungstext. */
export function kurzname(info, umsatzart = '') {
  let t = info;
  const laden = LAEDEN.find(([re]) => re.test(t));
  if (laden) return laden[1];
  const pp = t.match(/Ihr Einkauf bei\s+(.+?)\s+PayPal/i);
  if (/AMZN Mktp|Amazon\.de|AMAZON EU/i.test(t)) return 'Amazon';
  if (/Amazon Music/i.test(t)) return 'Amazon Music';
  if (pp && pp[1].trim()) t = pp[1];
  else if (/PayPal/i.test(t)) return 'PayPal';
  const kz = t.match(/Kartenzahlung\s+(.+?)\/\//i); if (kz) t = kz[1];
  if (/Visa Card/i.test(umsatzart)) t = t.split(/\s{2,}|\s+[A-Z]{2}\s+\d{4}\*{4}/)[0];
  if (/Bank Card/i.test(umsatzart) && !kz) t = t.replace(/^\/+/, '').split('/')[0];
  t = t.replace(/\s*(BLZ|BIC|IBAN|Kto):.*$/i, '').replace(/\b[0-9a-f]{8}-[0-9a-f-]{20,}\b/gi, ' ').replace(/\b\d{6,}[\w./-]*/g, ' ').replace(/\b(wolfs|burg|fil|\d+)\b/gi, ' ').replace(/\s-\s*$/, '').replace(/^[\d\W]+/, '')
    .replace(/\b(sagt danke|marken-discount|marken-discou|gmbh|s\.?a\.?r\.?l\.?.*|et cie.*|& co\.? kg|kg|ag|e\.v\.|wolfsburg|braunschweig-?|wofsburg)\b/gi, ' ')
    .replace(/[,;]+/g, ' ').replace(/\s+/g, ' ').trim();
  const w = t.split(' ').slice(0, 4).join(' ');
  return w ? w.replace(/\b([A-ZÄÖÜ])([A-ZÄÖÜ]+)\b/g, (_, a, b) => a + b.toLowerCase()) : (umsatzart || 'Unbekannt');
}

// Bekannte Läden → einheitlicher Name
const LAEDEN = [
  [/\blidl\b/i, 'Lidl'], [/\bnetto\b/i, 'Netto'], [/\baldi\b/i, 'Aldi'], [/\bedeka\b/i, 'Edeka'], [/\bkaufland\b/i, 'Kaufland'],
  [/\brewe\b/i, 'Rewe'], [/\bpenny\b/i, 'Penny'], [/\bdm[- ]drogerie\b/i, 'dm'], [/\brossmann\b/i, 'Rossmann'], [/\btedi\b/i, 'TEDi'],
  [/apotheke/i, 'Apotheke'], [/b(?:ae|ä)cker/i, 'Bäckerei'], [/\b(tgtg|toogoodtogo)/i, 'Too Good To Go'], [/takeaway|lieferando/i, 'Lieferando'],
  [/\bklarna\b/i, 'Klarna'], [/trade republic/i, 'Trade Republic'], [/disneyplus/i, 'Disney+'], [/\bnetflix\b/i, 'Netflix'],
  [/congstar/i, 'Congstar'], [/glasfaser/i, 'Glasfaser'], [/rundfunk/i, 'Rundfunkbeitrag'], [/wuertt\.?\s*versicherg/i, 'Zahnzusatz Württembergische'],
];

// Allgemeine Stichworte (keine persönlichen Daten)
const STICH = [
  { re: /\b(lidl|netto|aldi|edeka|kaufland|rewe|penny|norma|dm drogerie|rossmann|tedi|apotheke|b(?:ae|ä)ckerei|tgtg|toogoodtogo|cadera)\b/i, bereich: 'Wohnen & Haushalt', posten: null },
  { re: SPRIT_MUSTER, bereich: 'Autos', posten: 'Sprit' },
  { re: /\bparken\b/i, bereich: 'Autos', posten: 'Parken' },
  { re: /\b(autohaus|werkstatt|tüv|tuv)\b/i, bereich: 'Autos', posten: null },
  { re: /\bdrk\b/i, bereich: 'Kinder', posten: null },
  { re: /\bIHR ELTERNGELD\b/i, bereich: 'Staatliche Leistungen', posten: 'Elterngeld' },
  { re: /\b(familienkasse|kindergeld)\b/i, bereich: 'Staatliche Leistungen', posten: 'Kindergeld' },
  { re: /\blohn-?gehalt\b/i, bereich: 'Einkommen', posten: 'Gehalt' },
  { re: /\bhabenzinsen\b/i, bereich: 'Sonstige Einnahmen', posten: 'Zinsen' },
  { re: /\b(rundfunk)\b/i, bereich: 'Abos & Telefon', posten: null },
  { re: /\b(congstar|glasfaser|netflix|disneyplus|disney|amazon music|spotify)\b/i, bereich: 'Abos & Telefon', posten: null },
  { re: /\b(wuertt|versicherung|huk|allianz|adac|hansemerkur)\b/i, bereich: 'Versicherungen', posten: null },
  { re: /\b(trade republic)\b/i, bereich: null, posten: null },
];

const schluessel = (s) => (s || '').toLowerCase().replace(/[^a-zäöüß ]/g, ' ').replace(/\b(gutschrift|lastschrift|dauerauftrag|echtzeit|überweisung|ihr|einkauf|bei|paypal|europe|der|die|das|und|blz|bic|iban|kto)\b/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter((w) => w.length > 2).slice(0, 4).join(' ');
const aktivIn = (p, m) => (!p.ab || p.ab <= m) && (!p.bis || p.bis >= m);

function vorschlagen(z, kontoName, gegen, modell, regeln) {
  const rolle = (n) => modell.konten.find((k) => k.name === n)?.rolle;
  const eigen = rolle(gegen) && rolle(gegen) !== 'extern' && rolle(gegen) !== 'offen';
  const art = eigen ? 'Umbuchung' : z.betrag < 0 ? 'Ausgabe' : 'Einnahme';
  const von = z.betrag < 0 ? kontoName : gegen;
  const nach = z.betrag < 0 ? gegen : kontoName;
  const betrag = Math.abs(z.betrag);
  const monat = z.datum.slice(0, 7);
  const v = { art, von, nach, betrag, datum: z.datum, giltFuer: monat, bereich: '', posten: '', topf: '', beschreibung: z.info.replace(/\s*(BLZ|BIC):.*$/i, '').slice(0, 250), kommentar: '', gezahltVon: '' };
  if (z.kartendatum && z.kartendatum !== z.datum) v.kommentar = `Karte am ${z.kartendatum.split('-').reverse().join('.')}`;
  const gruende = [];
  let ampel = 'unbekannt';
  const txt = z.info;

  // 1. Eigene Regeln (Regeldatei): „Text enthält …“
  const regel = (regeln || []).find((r) => r.enthaelt && txt.toLowerCase().includes(r.enthaelt.toLowerCase()) && (!r.konto || r.konto === kontoName) && (!r.betrag || Math.abs(r.betrag - betrag) < 0.005));
  if (regel) {
    Object.assign(v, Object.fromEntries(Object.entries(regel.setze || {}).filter(([, w]) => w != null && w !== '')));
    if (regel.setze?.giltFuer === '+1') v.giltFuer = monatPlus(monat, 1);
    ampel = 'sicher'; gruende.push(`Regel „${regel.enthaelt}“`);
  }

  // 2. Frühere Buchungen mit gleichem Text (aus früheren Importen)
  if (ampel === 'unbekannt') {
    const sk = schluessel(kurzname(txt, z.umsatzart) + ' ' + txt);
    const frueher = [...modell.buchungen].reverse().find((b) => b.quelle === 'CSV' && b.beschreibung && schluessel(kurzname(b.beschreibung) + ' ' + b.beschreibung) === sk && (b.von === kontoName || b.nach === kontoName));
    if (frueher) {
      v.posten = frueher.posten || ''; v.bereich = frueher.bereich || ''; if (frueher.topf) v.topf = frueher.topf;
      ampel = 'sicher'; gruende.push(`wie ${frueher.nr}`);
    }
  }

  // 3. Feste Kosten / Einnahmen mit gleichem Betrag und Konto
  if (ampel === 'unbekannt' || !v.posten) {
    const fest = art === 'Einnahme'
      ? modell.einnahmen.filter((e) => e.konto === kontoName && Math.abs(e.betrag - betrag) < 0.005 && (aktivIn(e, monat) || aktivIn(e, monatPlus(monat, 1))))
      : modell.kosten.filter((k) => Math.abs((k.betrag ?? 0) - betrag) < 0.005 && (aktivIn(k, monat) || aktivIn(k, monatPlus(monat, 1))));
    const passendesKonto = fest.filter((x) => (x.konto === kontoName));
    const wahl = passendesKonto.length === 1 ? passendesKonto[0] : fest.length === 1 ? fest[0] : null;
    const laden = LAEDEN.find(([re]) => re.test(txt));
    const passtZumLaden = (p) => !laden || p.toLowerCase().includes(laden[1].toLowerCase()) || laden[1].toLowerCase().includes(p.toLowerCase().split(/[\s(]/)[0]);
    if (wahl && passtZumLaden(wahl.posten)) {
      v.posten = wahl.posten; v.bereich = wahl.bereich || v.bereich;
      if (wahl.topf && (rolle(von) === 'Sparen' || rolle(nach) === 'Sparen')) v.topf = wahl.topf;
      const wortGleich = wahl.posten.toLowerCase().split(/[\s()-]+/).some((w) => w.length > 3 && txt.toLowerCase().includes(w));
      // Monat: Einnahmen „Folgemonat“, Kosten „Monatsanfang“ am Monatsende gebucht → nächster Monat
      if (art === 'Einnahme' && wahl.giltFuer === 'Folgemonat') v.giltFuer = monatPlus(monat, 1);
      if (art !== 'Einnahme' && /monatsanfang/i.test(wahl.faellig || '') && +z.datum.slice(8, 10) >= 25) v.giltFuer = monatPlus(monat, 1);
      ampel = wortGleich && wahl.konto === kontoName ? 'sicher' : 'pruefen';
      gruende.push(`${wahl.id}: ${wahl.posten}${wahl.konto !== kontoName ? ` (laut Datei über ${wahl.konto})` : ''}`);
    }
  }

  // 4. Umbuchung zwischen eigenen Konten
  if (art === 'Umbuchung') {
    v.bereich = v.bereich || 'Sparen & Umbuchungen';
    if (!v.posten) v.posten = /kredit|visa/i.test(nach) ? 'Aufladung Visa' : `Umbuchung ${von} → ${nach}`;
    if (ampel === 'unbekannt') { ampel = 'pruefen'; gruende.push('Umbuchung zwischen euren Konten'); }
  }

  // 5. Allgemeine Stichworte
  if (!v.bereich || !v.posten) {
    const s = STICH.find((x) => x.re.test(txt));
    if (s) {
      if (!v.bereich && s.bereich) v.bereich = s.bereich;
      if (!v.posten) v.posten = s.posten || kurzname(txt, z.umsatzart);
      if (ampel === 'unbekannt' && v.bereich) { ampel = 'pruefen'; gruende.push('Stichwort'); }
    }
  }
  if (!v.posten) v.posten = kurzname(txt, z.umsatzart);
  if (!v.bereich && ampel !== 'unbekannt') ampel = 'pruefen';
  return { buchung: v, ampel, grund: gruende.join(', ') };
}

// ---------------------------------------------------------------- Import-ID

function fnv(s) { // 53-Bit-Hash, stabil
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 2654435761); h2 = Math.imul(h2 ^ c, 1597334677); }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}
export function importId(kontoId, z, n = 1) {
  const text = z.info.toLowerCase().replace(/\s+/g, ' ').trim();
  return `${kontoId}-${z.datum.replace(/-/g, '')}-${fnv(`${kontoId}|${z.datum}|${z.betrag.toFixed(2)}|${text}`)}${n > 1 ? '-' + n : ''}`;
}

// ---------------------------------------------------------------- Analyse

/**
 * Analysiert einen oder mehrere Exporte gegen die Excel-Datei.
 * @param {Array<{text:string, name:string}>} dateien
 * @param {object} modell aus leseBuchhaltung()
 * @param {Array} regeln eigene Regeln (optional)
 */
export function analysieren(dateien, modell, regeln = []) {
  const csvs = dateien.map((d) => { try { return { ...leseCsv(d.text, d.name) }; } catch (e) { return { dateiname: d.name, fehler: e.message, zeilen: [] }; } });
  const importNummern = [];
  const dateiInfo = csvs.map((c) => {
    if (c.fehler) return { csv: c, konto: null, importieren: false, hinweis: c.fehler };
    const k = kontoErkennen(c, modell);
    if (!k) return { csv: c, konto: null, importieren: false, hinweis: `Konto Nr. …${endung(c.kontoNr)} ist in eurer Datei nicht angelegt.` };
    importNummern.push([c.kontoNr, k.name]);
    if (!k.csv) return { csv: c, konto: k, importieren: false, hinweis: `${k.name} ist ${k.rolle === 'privat' ? 'ein privates Konto' : 'nicht für den Import vorgesehen'} („CSV-Import: nein“). Wird nicht importiert.` };
    return { csv: c, konto: k, importieren: true, hinweis: null };
  });

  // Import-IDs, die schon in der Datei stehen (eine Buchung kann mehrere haben, getrennt durch Leerzeichen)
  const idsDatei = new Map();
  for (const b of modell.buchungen) for (const id of (b.importId || '').split(/\s+/).filter(Boolean)) idsDatei.set(id, b.nr);
  const zeilen = [];
  dateiInfo.forEach((d, di) => {
    if (!d.importieren) return;
    const k = d.konto;
    const zaehler = new Map();
    for (const z of d.csv.zeilen) {
      const basis = `${z.datum}|${z.betrag}|${z.info}`;
      const n = (zaehler.get(basis) || 0) + 1; zaehler.set(basis, n);
      const id = importId(k.id, z, n);
      const gegen = gegenkonto(z.info, modell, k.name, importNummern);
      const { buchung, ampel, grund } = vorschlagen(z, k.name, gegen, modell, regeln);
      buchung.importId = id;
      const eintrag = { id, datei: di, konto: k.name, kontoId: k.id, zeile: z, buchung, ampel, grund, status: 'neu', auswahl: true, treffer: null };
      if (idsDatei.has(id)) { eintrag.status = 'importiert'; eintrag.auswahl = false; eintrag.treffer = [idsDatei.get(id)]; eintrag.grund = `schon importiert (${idsDatei.get(id)})`; }
      zeilen.push(eintrag);
    }
  });

  // ---- Abgleich mit dem, was ihr selbst eingetragen habt (je Konto)
  const cent = (x) => Math.round(x * 100);
  const daten = (l) => [l.zeile.datum, l.zeile.wertstellung, l.zeile.kartendatum].filter(Boolean);
  const abstand = (l, b) => Math.min(...daten(l).map((d) => Math.abs(tag(b.datum) - tag(d))));
  const kurz = (b) => `${b.nr} ${b.posten || b.beschreibung || ''}`.trim();
  const teilmenge = (werte, ziel, max = 4) => { // kleinste Kombination aus 2..max Werten mit Summe = ziel
    const n = Math.min(werte.length, 14);
    for (let g = 2; g <= max; g++) {
      const idx = [];
      const suche = (start, rest, tiefe) => {
        if (tiefe === g) return rest === 0;
        for (let i = start; i < n; i++) { idx.push(i); if (suche(i + 1, rest - werte[i], tiefe + 1)) return true; idx.pop(); }
        return false;
      };
      if (suche(0, ziel, 0)) return idx.slice();
    }
    return null;
  };
  const nurDatei = {};
  for (const d of dateiInfo.filter((x) => x.importieren)) {
    const k = d.konto;
    const s = (b) => (b.nach === k.name ? b.betrag : -b.betrag);
    // Buchungen dieses Kontos, die noch keine Import-ID von diesem Konto haben
    const frei = modell.buchungen.filter((b) => b.datum && (b.von === k.name || b.nach === k.name)
      && !(b.importId || '').split(/\s+/).some((id) => id.startsWith(k.id + '-')));
    const offen = () => zeilen.filter((l) => l.konto === k.name && l.status === 'neu');
    const nutze = (b) => frei.splice(frei.indexOf(b), 1);
    const markiere = (l, bs, ampel, grund) => { l.status = 'eingetragen'; l.treffer = bs.map((b) => b.nr); l.auswahl = true; l.ampel = ampel; l.grund = grund; };

    // 1. gleicher Betrag, Datum höchstens 3 Tage auseinander
    const paare = [];
    for (const l of offen()) for (const b of frei) if (cent(s(b)) === cent(l.zeile.betrag) && abstand(l, b) <= 3) paare.push([abstand(l, b), l, b]);
    paare.sort((a, b) => a[0] - b[0]);
    for (const [, l, b] of paare) if (l.status === 'neu' && frei.includes(b)) { nutze(b); markiere(l, [b], 'sicher', `schon eingetragen als ${kurz(b)}`); }

    // 2. mehrere eurer Buchungen = eine Bankzeile
    for (const l of offen()) {
      const kand = frei.filter((b) => Math.sign(s(b)) === Math.sign(l.zeile.betrag) && abstand(l, b) <= 5);
      const t = kand.length >= 2 && teilmenge(kand.map((b) => cent(s(b))), cent(l.zeile.betrag));
      if (t) { const bs = t.map((i) => kand[i]); bs.forEach(nutze); markiere(l, bs, 'pruefen', `zusammen eingetragen als ${bs.map((b) => b.nr).join(' + ')}`); }
    }
    // 3. eine eurer Buchungen = mehrere Bankzeilen
    for (const b of [...frei]) {
      const kand = offen().filter((l) => Math.sign(l.zeile.betrag) === Math.sign(s(b)) && abstand(l, b) <= 5);
      const t = kand.length >= 2 && teilmenge(kand.map((l) => cent(l.zeile.betrag)), cent(s(b)));
      if (t) { nutze(b); t.map((i) => kand[i]).forEach((l, _, alle) => markiere(l, [b], 'pruefen', `${alle.length} Bankzeilen ergeben zusammen ${kurz(b)}`)); }
    }
    // 4. Unklar: gleicher Betrag mit größerem Datumsabstand oder fast gleicher Betrag (Tippfehler)
    const ziffern = (x) => String(Math.abs(cent(x))).split('').sort().join('');
    for (const l of offen()) {
      const kand = frei.map((b) => {
        const a = abstand(l, b), gleich = cent(s(b)) === cent(l.zeile.betrag);
        const fast = !gleich && Math.sign(s(b)) === Math.sign(l.zeile.betrag)
          && (Math.abs(s(b) - l.zeile.betrag) <= Math.max(1, Math.abs(l.zeile.betrag) * 0.05) || ziffern(s(b)) === ziffern(l.zeile.betrag));
        return (gleich && a <= 10) || (fast && a <= 3) ? { b, a, gleich } : null;
      }).filter(Boolean).sort((x, y) => (y.gleich - x.gleich) || (x.a - y.a));
      if (!kand.length) continue;
      const { b, a, gleich } = kand[0];
      nutze(b);
      l.status = 'vielleicht'; l.auswahl = false; l.ampel = 'pruefen'; l.entscheidung = null;
      l.kandidat = { nr: b.nr, posten: b.posten, datum: b.datum, betrag: s(b) };
      l.grund = gleich ? `Vielleicht schon eingetragen als ${kurz(b)} – Datum weicht ${a} Tage ab`
        : `Vielleicht schon eingetragen als ${kurz(b)} – dort steht ${Math.abs(s(b)).toFixed(2).replace('.', ',')} € statt ${Math.abs(l.zeile.betrag).toFixed(2).replace('.', ',')} €`;
    }
    // Rest: vor dem Stichtag oder neu
    for (const l of offen()) if (k.stichtag && l.zeile.datum <= k.stichtag) { l.status = 'vor-stichtag'; l.auswahl = false; }
    // In der Datei, aber nicht auf dem Auszug (im Zeitraum des Auszugs)
    const vonD = d.csv.von || '0000', bisD = d.csv.bis || '9999';
    nurDatei[k.name] = frei.filter((b) => b.datum >= vonD && b.datum <= bisD && !(b.importId || '').trim())
      .map((b) => ({ nr: b.nr, posten: b.posten, datum: b.datum, betrag: s(b), vorStichtag: !!(k.stichtag && b.datum <= k.stichtag) }));
  }

  // ---- Umbuchungen, die in zwei Exporten vorkommen → nur einmal buchen
  for (const a of zeilen) {
    if (a.status !== 'neu' && a.status !== 'vor-stichtag') continue;
    if (a.buchung.art !== 'Umbuchung') continue;
    const b = zeilen.find((x) => x !== a && (x.status === 'neu' || x.status === 'vor-stichtag') && x.konto !== a.konto
      && x.buchung.art === 'Umbuchung' && Math.abs(x.zeile.betrag + a.zeile.betrag) < 0.005
      && x.buchung.von === a.buchung.von && x.buchung.nach === a.buchung.nach
      && Math.abs(tag(x.zeile.datum) - tag(a.zeile.datum)) <= 3);
    if (!b) continue;
    const [haupt, gegen] = a.zeile.betrag < 0 ? [a, b] : [b, a];
    gegen.status = 'gegenbuchung'; gegen.auswahl = false; gegen.grund = `Gegenbuchung zu ${haupt.konto}, ${haupt.zeile.datum.split('-').reverse().join('.')}`;
    gegen.gegenId = haupt.id; haupt.gegenId = gegen.id;
  }

  // ---- Saldo-Abgleich je Konto
  const abgleich = dateiInfo.filter((d) => d.importieren).map((d) => {
    const k = d.konto; const bis = d.csv.bis || d.csv.zeilen.reduce((m, z) => (z.datum > m ? z.datum : m), '');
    const datei = kontostand(modell, k.name, bis);
    const eigene = zeilen.filter((z) => z.konto === k.name);
    const nachStichtag = (z) => !k.stichtag || z.zeile.datum > k.stichtag;
    const hauptNeu = (z) => { const h = zeilen.find((x) => x.id === z.gegenId); return h && h.status === 'neu'; };
    const neu = r2(eigene.filter((z) => nachStichtag(z) && (z.status === 'neu' || (z.status === 'gegenbuchung' && hauptNeu(z)))).reduce((s, z) => s + z.zeile.betrag, 0));
    const unklar = eigene.filter((z) => z.status === 'vielleicht');
    const fehlen = (nurDatei[k.name] || []).filter((b) => !b.vorStichtag && b.datum <= bis);
    const fehlenSumme = r2(fehlen.reduce((s, b) => s + b.betrag, 0));
    const differenz = r2(d.csv.saldo - datei);
    return { konto: k.name, bis, bank: d.csv.saldo, datei, differenz, neu, fehlen, fehlenSumme, unklar: unklar.length, rest: r2(differenz - neu + fehlenSumme) };
  });

  return { dateien: dateiInfo.map((d) => ({ name: d.csv.dateiname, konto: d.konto?.name || null, kontoText: d.csv.kontoText, saldo: d.csv.saldo, von: d.csv.von, bis: d.csv.bis, anzahl: d.csv.zeilen.length, importieren: d.importieren, hinweis: d.hinweis })), zeilen, abgleich, nurDatei };
}
