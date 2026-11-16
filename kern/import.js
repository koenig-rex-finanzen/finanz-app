// Import von Kontoauszügen (VW Bank, c24) ins einheitliche Format – für alle Bereiche.
// Reine Funktionen ohne Browser (auch in Node). Keine persönlichen Daten: Konten werden über die
// Kontonummer-Endung im Blatt „Konten“ erkannt, Vorschläge kommen aus Blatt „Regeln“, Kosten, Einnahmen und früheren Buchungen.
//
// Ablauf: lesen() je Datei → abgleichen() gegen die Datei(en) des Bereichs → Nutzer prüft → übernehmen (app/datei.js).
// Status je Bankzeile:
//   importiert   Import-ID steht schon in der Datei – nichts tun
//   eingetragen  von Hand eingetragen (gleicher Betrag, Datum ±3 Tage) – nur Import-ID ergänzen
//   vielleicht   unsicherer Treffer – Nutzer entscheidet (nie automatisch)
//   neu          wird als Buchung angelegt (Vorschlag)
//   gegenbuchung dieselbe Überweisung steht im Export des anderen eigenen Kontos – nur einmal buchen
//   vor-stichtag liegt vor dem Stichtag des Kontos – zählt nicht

import { leseCsv, importId as vwImportId, kurzname } from './csv.js';
import { leseC24, importIds as c24ImportIds } from './kai.js';
import { kategorieVorschlag } from './einheit.js';
import { NUR_TOPF } from './format.js';

const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const tag = (iso) => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
const monatPlus = (m, n) => { const d = new Date(Date.UTC(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)); return d.toISOString().slice(0, 7); };
const endung = (s) => (String(s || '').replace(/\s/g, '').match(/(\d{4})\D*$/) || [])[1] || null;
const cent = (x) => Math.round(x * 100);
const aktivIn = (p, m) => (!p.ab || p.ab <= m) && (!p.bis || p.bis >= m) && p.status !== 'beendet' && p.status !== 'pausiert';
const ids = (b) => String(b.importId || '').split(/\s+/).filter(Boolean);
const istTeil = (b) => b.quelle === 'Aufteilung' && /^Teil von \S+/.test(b.kommentar || '');
const teilVon = (b) => (istTeil(b) ? b.kommentar.match(/^Teil von (\S+)/)[1] : null);

// ---------------------------------------------------------------- Lesen

/** Welcher Export ist das? */
export function formatErkennen(text) {
  if (/Transaktionstyp/i.test(text) && /Buchungsdatum/i.test(text)) return 'c24';
  if (/^\s*"?Nr\.?"?\s*;\s*"?Buchungsdatum"?\s*;/im.test(text.replace(/﻿/g, ''))) return 'vw';
  if (/transaction_id/.test(text) && /datetime/.test(text)) return 'tr';
  return null;
}

/**
 * Liest einen Export. Die Zeilen bekommen ihr Konto erst in abgleichen() (über die Endung).
 * @returns {{name, format, kontoEndung, kontoText, saldo, von, bis, zeilen:Array}}
 */
export async function lesen(text, name = '') {
  const format = formatErkennen(text);
  if (format === 'vw') {
    const c = leseCsv(text, name);
    const zeilen = c.zeilen.map((z) => ({
      datum: z.datum, wertstellung: z.wertstellung, kartendatum: z.kartendatum, betrag: z.betrag, text: z.info, umsatzart: z.umsatzart,
      gegenNummern: [...z.info.matchAll(/IBAN:\s*([A-Z]{2}\d{2}[A-Z0-9]{8,30})/g), ...z.info.matchAll(/Kto:\s*(\d{5,})/g)].map((m) => m[1]),
      empfaenger: null, bankKategorie: null, _vw: z,
    }));
    return { name, format, kontoEndung: endung(c.kontoNr), kontoNr: c.kontoNr, kontoText: c.kontoText, saldo: c.saldo, von: c.von, bis: c.bis, zeilen };
  }
  if (format === 'c24') {
    const roh = await c24ImportIds(leseC24(text));
    const zeilen = roh.map((z) => ({
      datum: z.datum, wertstellung: null, kartendatum: null, betrag: z.betrag, text: [z.empfaenger, z.zweck, z.beschreibung].filter(Boolean).join(' '),
      umsatzart: z.typ, gegenNummern: z.iban ? [z.iban.replace(/\s/g, '')] : [], empfaenger: z.empfaenger || null,
      bankKategorie: [z.kategorie, z.unterkategorie].filter(Boolean).join(' / ') || null,
      kontoEndung: endung(z.kontonummer), kontoName: z.kontoname || null, id: z.id,
      beschreibung: [z.zweck || [z.typ, z.karteneinsatz].filter(Boolean).join(' '), z.empfaenger].filter(Boolean).join(' · ').slice(0, 250),
    }));
    const daten = zeilen.map((z) => z.datum).sort();
    return { name, format, kontoEndung: null, kontoText: 'c24', saldo: null, von: daten[0] || null, bis: daten.at(-1) || null, zeilen };
  }
  if (format === 'tr') throw new Error(`„${name}“ ist ein Trade-Republic-Export. Der kommt ins Depot (eigene Seite), nicht in die Buchungen.`);
  throw new Error(`„${name || 'Datei'}“ ist weder ein VW-Bank- noch ein c24-Export.`);
}

// ---------------------------------------------------------------- Konten

/**
 * Alle Konten des Bereichs mit ihrer Datei. Eigene Konten zuerst (ein Konto kann in zwei Dateien stehen: einmal eigen, einmal als Gegenüber).
 * @param {Object<string, object>} dateien z. B. { gemeinsam: m, tagesgeld: m } (m aus leseDatei)
 */
export function kontenListe(dateien) {
  const l = [];
  for (const [datei, m] of Object.entries(dateien)) for (const k of m.konten || []) if (k.eigen === 'Ja') l.push({ ...k, datei });
  for (const [datei, m] of Object.entries(dateien)) for (const k of m.konten || []) if (k.eigen !== 'Ja' && !l.some((x) => x.name === k.name)) l.push({ ...k, datei });
  return l;
}

/** Konto einer Exportzeile bzw. eines Exports (nur eigene Konten). */
function kontoFuer(exp, z, konten) {
  const eigene = konten.filter((k) => k.eigen === 'Ja');
  const e = z.kontoEndung || exp.kontoEndung;
  let k = e && eigene.find((x) => endung(x.endung) === e);
  if (!k && z.kontoName) k = eigene.find((x) => x.name.toLowerCase().includes(z.kontoName.toLowerCase()) && z.kontoName.length > 3);
  if (!k && exp.format === 'vw' && /visa|kreditkarte/i.test(`${exp.kontoText} ${exp.name}`)) k = eigene.find((x) => /kredit/i.test(x.art || '') || /visa/i.test(x.name));
  if (!k && exp.format === 'c24' && !z.kontoEndung && !z.kontoName) {
    // Export ohne Kontospalte: das eine c24-Hauptkonto des Bereichs
    const c24 = eigene.filter((x) => /c\s?24/i.test(`${x.bank} ${x.name}`) && x.csv === 'Ja' && !/pocket|spar/i.test(`${x.art} ${x.rolle}`));
    if (c24.length === 1) k = c24[0];
  }
  return k || null;
}

/** Gegenkonto aus Bank-Text/IBAN: eigenes oder bekanntes Konto über die Endung, sonst null. */
function gegenkontoFuer(z, konten, eigenesKonto, exporte) {
  for (const n of z.gegenNummern || []) {
    const k = konten.find((x) => endung(x.endung) && n.endsWith(endung(x.endung)) && x.name !== eigenesKonto);
    if (k) return k;
  }
  // Kontonummer eines anderen importierten Kontos im Text (z. B. Visa-Ausgleich)
  for (const e of exporte) if (e.konto && e.konto.name !== eigenesKonto && e.kontoNr && e.kontoNr.length >= 10 && z.text.includes(e.kontoNr)) return e.konto;
  // c24: Umbuchung zwischen Konto und Pocket ohne IBAN – Empfänger heißt wie das Pocket
  if (z.empfaenger && z.empfaenger.trim().length > 3 && /umbuchung/i.test(z.bankKategorie || '')) {
    const k = konten.find((x) => x.name !== eigenesKonto && x.name.toLowerCase().includes(z.empfaenger.trim().toLowerCase()));
    if (k) return k;
  }
  return null;
}

// ---------------------------------------------------------------- Vorschlag

const schluessel = (s) => (s || '').toLowerCase().replace(/[^a-zäöüß ]/g, ' ')
  .replace(/\b(gutschrift|lastschrift|dauerauftrag|echtzeit|überweisung|ihr|einkauf|bei|paypal|europe|der|die|das|und|blz|bic|iban|kto|kartenzahlung)\b/g, ' ')
  .replace(/\s+/g, ' ').trim().split(' ').filter((w) => w.length > 2).slice(0, 4).join(' ');

/**
 * Buchungsvorschlag für eine Bankzeile (Sicht der Datei, in der das Konto eigen ist).
 * @returns {{buchung, ampel: 'sicher'|'pruefen'|'unbekannt', grund}}
 */
export function vorschlagen(z, konto, gegen, m, { kategorien = [], extern } = {}) {
  // Überweisung zwischen zwei eigenen Konten des Bereichs (auch Gemeinschaftskonto ↔ Tagesgeld) = Umbuchung
  const art = gegen && gegen.eigen === 'Ja' ? 'Umbuchung' : z.betrag < 0 ? 'Ausgabe' : 'Einnahme';
  const gegenName = gegen ? gegen.name : (extern || z.empfaenger || kurzname(z.text, z.umsatzart) || 'Extern');
  const von = z.betrag < 0 ? konto.name : gegenName;
  const nach = z.betrag < 0 ? gegenName : konto.name;
  const betrag = r2(Math.abs(z.betrag));
  const monat = z.datum.slice(0, 7);
  const v = {
    datum: z.datum, giltFuer: monat, art, von, nach, betrag, kategorie: null, bankKategorie: z.bankKategorie, posten: null, topf: null,
    beschreibung: (z.beschreibung || z.text.replace(/\s*(BLZ|BIC):.*$/i, '')).slice(0, 250), kommentar: null, quelle: 'CSV', importId: z.id, gezahltVon: null,
  };
  if (z.kartendatum && z.kartendatum !== z.datum) v.kommentar = `Karte am ${z.kartendatum.split('-').reverse().join('.')}`;
  const txt = z.text.toLowerCase();
  const gruende = [];
  let ampel = 'unbekannt';

  // 1. Eigene Regeln (Blatt „Regeln“)
  const regel = (m.regeln || []).find((r) => r.stichwort && txt.includes(String(r.stichwort).toLowerCase()) && (!r.konto || r.konto === konto.name));
  if (regel) {
    if (regel.kategorie) v.kategorie = regel.kategorie;
    if (regel.posten) v.posten = regel.posten;
    if (regel.topf) v.topf = regel.topf;
    ampel = 'sicher'; gruende.push(`Regel „${regel.stichwort}“`);
  }

  // 2. Frühere Buchung mit gleichem Bank-Text (aus einem früheren Import desselben Kontos)
  if (ampel === 'unbekannt') {
    const sk = schluessel(`${kurzname(z.text, z.umsatzart)} ${z.text}`);
    const frueher = sk && [...m.buchungen].sort((a, b) => (b.datum || '').localeCompare(a.datum || ''))
      .find((b) => b.quelle === 'CSV' && b.beschreibung && (b.von === konto.name || b.nach === konto.name) && b.art === art
        && schluessel(`${kurzname(b.beschreibung)} ${b.beschreibung}`) === sk);
    if (frueher) {
      v.posten = frueher.posten || null; v.kategorie = frueher.kategorie || null; v.topf = frueher.topf || null;
      if (!gegen && frueher.art === 'Ausgabe' && frueher.nach) v.nach = frueher.nach;
      if (!gegen && frueher.art === 'Einnahme' && frueher.von) v.von = frueher.von;
      ampel = 'sicher'; gruende.push(`wie ${frueher.nr}`);
    }
  }

  // 3. Feste Kosten / Einnahmen mit gleichem Betrag
  if (!v.posten && art !== 'Umbuchung') {
    const liste = art === 'Einnahme' ? m.einnahmen || [] : m.kosten || [];
    const fest = liste.filter((p) => p.posten && Math.abs((p.betrag ?? 0) - betrag) < 0.005 && (aktivIn(p, monat) || aktivIn(p, monatPlus(monat, 1))));
    const amKonto = fest.filter((p) => p.konto === konto.name);
    const wahl = amKonto.length === 1 ? amKonto[0] : fest.length === 1 ? fest[0] : null;
    if (wahl) {
      v.posten = wahl.posten; if (!v.kategorie && wahl.kategorie) v.kategorie = wahl.kategorie;
      const wortGleich = wahl.posten.toLowerCase().split(/[\s()/-]+/).some((w) => w.length > 3 && txt.includes(w));
      if (ampel === 'unbekannt') ampel = wortGleich && (!wahl.konto || wahl.konto === konto.name) ? 'sicher' : 'pruefen';
      gruende.push(`fester Posten „${wahl.posten}“`);
    }
  }
  // 4. Posten-Name steht im Bank-Text (Kosten und Einnahmen)
  if (!v.posten && art !== 'Umbuchung') {
    const p = [...(art === 'Einnahme' ? m.einnahmen || [] : m.kosten || [])].filter((x) => x.posten && x.posten.length > 3 && aktivIn(x, monat))
      .find((x) => txt.includes(x.posten.toLowerCase().replace(/\s*\(.*\)$/, '')));
    if (p) { v.posten = p.posten; if (!v.kategorie && p.kategorie) v.kategorie = p.kategorie; if (ampel === 'unbekannt') ampel = 'pruefen'; gruende.push(`Posten „${p.posten}“ im Text`); }
  }
  // 5. Gehalt: ab dem 20. gilt es für den Folgemonat
  if (art === 'Einnahme' && /gehalt|lohn|bezüge|bezuege/i.test(z.text)) {
    if (!v.posten) { const g = (m.einnahmen || []).find((e) => /gehalt|lohn/i.test(e.posten || '')); if (g) v.posten = g.posten; }
    if (+z.datum.slice(8, 10) >= 20) v.giltFuer = monatPlus(monat, 1);
  } else if (art === 'Einnahme' && v.posten && /gehalt|lohn/i.test(v.posten) && +z.datum.slice(8, 10) >= 20) v.giltFuer = monatPlus(monat, 1);
  // 6. Kategorie (nur Ausgaben)
  if (art === 'Ausgabe' && !v.kategorie) {
    const k = kategorieVorschlag(m, v, kategorien);
    if (k) { v.kategorie = k; if (ampel === 'unbekannt') ampel = 'pruefen'; gruende.push(`Kategorie „${k}“`); }
  }
  if (art === 'Umbuchung') { if (ampel === 'unbekannt') ampel = 'pruefen'; gruende.push(`Umbuchung zwischen euren Konten`); }
  return { buchung: v, ampel, grund: gruende.join(', ') };
}

// ---------------------------------------------------------------- Abgleich

/** Vorzeichen einer Buchung aus Sicht eines Kontos (null = bewegt das Konto nicht). Aufgeteilte Buchungen zählen mit ihren Teilen. */
function signiertFuer(b, kontoName, teile) {
  if (NUR_TOPF.includes(b.art) || istTeil(b)) return null;
  const s = (x) => (x.nach === kontoName ? x.betrag : x.von === kontoName ? -x.betrag : 0);
  if (b.von !== kontoName && b.nach !== kontoName) return null;
  return r2(s(b) + (teile.get(b.nr) || []).reduce((a, t) => a + s(t), 0));
}
/** Gehört eine Import-ID zu diesem Konto? (VW: beginnt mit der Konto-ID; c24: c24-… an einem c24-Konto) */
const idVonKonto = (id, k) => id.startsWith(k.id + '-') || (id.startsWith('c24-') && /c\s?24/i.test(`${k.bank} ${k.name}`));

/** Kleinste Kombination aus 2..max Werten mit Summe ziel (Cent). */
function teilmenge(werte, ziel, max = 4) {
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
}

/**
 * Gleicht gelesene Exporte mit den Dateien eines Bereichs ab. Schreibt nichts.
 * @param {Array} exporte aus lesen()
 * @param {Object<string, object>} dateien { gemeinsam: m, tagesgeld: m } oder { kai: m }
 * @returns {{exporte, zeilen, abgleich, nurDatei}}
 */
export function abgleichen(exporte, dateien, { kategorien = [] } = {}) {
  const konten = kontenListe(dateien);
  const zeilen = [];
  // VW-IDs: laufende Nummer bei gleicher Zeile (wie bisher)
  for (const [ei, e] of exporte.entries()) {
    if (e.fehler) continue;
    e.konto = e.format === 'vw' ? kontoFuer(e, {}, konten) : null;
    const zaehler = new Map();
    for (const z of e.zeilen) {
      const k = e.format === 'vw' ? e.konto : kontoFuer(e, z, konten);
      if (!k) { zeilen.push({ export: ei, zeile: z, konto: null, status: 'ohne-konto', auswahl: false, grund: `Konto …${z.kontoEndung || e.kontoEndung || '?'} ist in euren Dateien nicht als eigenes Konto angelegt` }); continue; }
      if (e.format === 'vw') {
        const basis = `${z.datum}|${z.betrag}|${z.text}`;
        const n = (zaehler.get(basis) || 0) + 1; zaehler.set(basis, n);
        z.id = vwImportId(k.id, z._vw, n);
      }
      zeilen.push({ export: ei, zeile: z, konto: k, id: z.id, status: 'neu', auswahl: true });
    }
  }
  for (const e of exporte) if (!e.konto && e.format === 'c24') { const k = zeilen.find((l) => exporte[l.export] === e && l.konto)?.konto; if (k) e.konto = k; }

  // Import-IDs je Datei
  const idsDatei = new Map();
  for (const [datei, m] of Object.entries(dateien)) for (const b of m.buchungen) for (const id of ids(b)) idsDatei.set(`${datei}|${id}`, b.nr);
  const alleIds = new Set(zeilen.map((l) => l.id).filter(Boolean));
  for (const l of zeilen) {
    if (!l.konto) continue;
    const nr = idsDatei.get(`${l.konto.datei}|${l.id}`);
    if (nr) { l.status = 'importiert'; l.auswahl = false; l.treffer = [nr]; l.grund = `schon importiert (${nr})`; }
  }

  // Vorschläge
  for (const l of zeilen.filter((x) => x.status === 'neu')) {
    const m = dateien[l.konto.datei];
    const gegen = gegenkontoFuer(l.zeile, konten, l.konto.name, exporte);
    const extern = m.konten.find((k) => k.rolle === 'extern')?.name;
    Object.assign(l, vorschlagen(l.zeile, l.konto, gegen, m, { kategorien, extern }));
    l.gegen = gegen?.name || null;
    // Stichwort für „Regel merken“: Empfänger (c24) bzw. kurzer Name aus dem Bank-Text (VW)
    l.stichwort = (l.zeile.empfaenger || kurzname(l.zeile.text, l.zeile.umsatzart) || '').trim().slice(0, 40) || null;
  }

  // Abgleich mit dem, was ihr selbst eingetragen habt – je Konto
  const nurDatei = {};
  const kontenImImport = [...new Map(zeilen.filter((l) => l.konto).map((l) => [l.konto.name, l.konto])).values()];
  const abstandVon = (l, b) => Math.min(...[l.zeile.datum, l.zeile.wertstellung, l.zeile.kartendatum].filter(Boolean).map((d) => Math.abs(tag(b.datum) - tag(d))));
  const kurz = (b) => `${b.nr} ${b.posten || b.beschreibung || ''}`.trim().slice(0, 60);
  for (const k of kontenImImport) {
    const m = dateien[k.datei];
    const teile = new Map();
    for (const t of m.buchungen) { const p = teilVon(t); if (p) teile.set(p, [...(teile.get(p) || []), t]); }
    const s = (b) => signiertFuer(b, k.name, teile);
    const frei = m.buchungen.filter((b) => b.datum && s(b) != null && !ids(b).some((id) => alleIds.has(id) || idVonKonto(id, k)));
    const offen = () => zeilen.filter((l) => l.konto?.name === k.name && l.status === 'neu');
    const nutze = (b) => frei.splice(frei.indexOf(b), 1);
    const markiere = (l, bs, ampel, grund) => { l.status = 'eingetragen'; l.treffer = bs.map((b) => b.nr); l.auswahl = true; l.ampel = ampel; l.grund = grund; };

    // 1. gleicher Betrag, Datum höchstens 3 Tage auseinander
    const paare = [];
    for (const l of offen()) for (const b of frei) if (cent(s(b)) === cent(l.zeile.betrag) && abstandVon(l, b) <= 3) paare.push([abstandVon(l, b), l, b]);
    paare.sort((a, b) => a[0] - b[0]);
    for (const [, l, b] of paare) if (l.status === 'neu' && frei.includes(b)) { nutze(b); markiere(l, [b], 'sicher', `schon eingetragen als ${kurz(b)}`); }
    // 2. mehrere eurer Buchungen = eine Bankzeile
    for (const l of offen()) {
      const kand = frei.filter((b) => Math.sign(s(b)) === Math.sign(l.zeile.betrag) && abstandVon(l, b) <= 5);
      const t = kand.length >= 2 && teilmenge(kand.map((b) => cent(s(b))), cent(l.zeile.betrag));
      if (t) { const bs = t.map((i) => kand[i]); bs.forEach(nutze); markiere(l, bs, 'pruefen', `zusammen eingetragen als ${bs.map((b) => b.nr).join(' + ')}`); }
    }
    // 3. eine eurer Buchungen = mehrere Bankzeilen
    for (const b of [...frei]) {
      const kand = offen().filter((l) => Math.sign(l.zeile.betrag) === Math.sign(s(b)) && abstandVon(l, b) <= 5);
      const t = kand.length >= 2 && teilmenge(kand.map((l) => cent(l.zeile.betrag)), cent(s(b)));
      if (t) { nutze(b); const alle = t.map((i) => kand[i]); alle.forEach((l) => markiere(l, [b], 'pruefen', `${alle.length} Bankzeilen ergeben zusammen ${kurz(b)}`)); }
    }
    // 4. unsicher: gleicher Betrag mit größerem Abstand oder fast gleicher Betrag (Tippfehler)
    const ziffern = (x) => String(Math.abs(cent(x))).split('').sort().join('');
    for (const l of offen()) {
      const kand = frei.map((b) => {
        const a = abstandVon(l, b), gleich = cent(s(b)) === cent(l.zeile.betrag);
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
    for (const l of offen()) if (k.stichtag && l.zeile.datum <= k.stichtag) { l.status = 'vor-stichtag'; l.auswahl = false; l.grund = 'vor dem Stichtag des Kontos'; }
    // In der Datei, aber nicht auf dem Auszug (im Zeitraum des Auszugs)
    const eigeneZeilen = zeilen.filter((l) => l.konto?.name === k.name);
    const vonD = eigeneZeilen.reduce((x, l) => (l.zeile.datum < x ? l.zeile.datum : x), '9999');
    const bisD = eigeneZeilen.reduce((x, l) => (l.zeile.datum > x ? l.zeile.datum : x), '0000');
    nurDatei[k.name] = frei.filter((b) => b.datum >= vonD && b.datum <= bisD && !ids(b).length)
      .map((b) => ({ nr: b.nr, posten: b.posten || b.beschreibung, datum: b.datum, betrag: s(b), vorStichtag: !!(k.stichtag && b.datum <= k.stichtag) }));
  }

  // Überweisungen zwischen zwei eigenen Konten, die in beiden Exporten stehen → nur einmal buchen
  for (const a of zeilen) {
    if (!['neu', 'vor-stichtag'].includes(a.status) || !a.gegen) continue;
    const b = zeilen.find((x) => x !== a && ['neu', 'vor-stichtag'].includes(x.status) && x.konto?.name === a.gegen && x.gegen === a.konto.name
      && Math.abs(x.zeile.betrag + a.zeile.betrag) < 0.005 && Math.abs(tag(x.zeile.datum) - tag(a.zeile.datum)) <= 3);
    if (!b) continue;
    const [haupt, gegen] = a.zeile.betrag < 0 ? [a, b] : [b, a];
    gegen.status = 'gegenbuchung'; gegen.auswahl = false; gegen.grund = `steht schon im Export von ${haupt.konto.name} (${haupt.zeile.datum.split('-').reverse().join('.')})`;
    gegen.gegenId = haupt.id; haupt.gegenId = gegen.id;
  }
  // Gegenbuchung, deren Hauptzeile schon in der Datei steht (z. B. Überweisung schon über das andere Konto importiert)
  for (const l of zeilen.filter((x) => x.status === 'neu' && x.gegen)) {
    const g = konten.find((k) => k.name === l.gegen);
    if (!g || g.eigen !== 'Ja') continue;
    const m = dateien[l.konto.datei];
    const da = m.buchungen.find((b) => b.datum && Math.abs(tag(b.datum) - tag(l.zeile.datum)) <= 3 && Math.abs(b.betrag - Math.abs(l.zeile.betrag)) < 0.005
      && ((l.zeile.betrag < 0 && b.von === l.konto.name && b.nach === l.gegen) || (l.zeile.betrag > 0 && b.nach === l.konto.name && b.von === l.gegen)) && ids(b).length);
    if (da) { l.status = 'eingetragen'; l.treffer = [da.nr]; l.ampel = 'sicher'; l.grund = `schon über ${l.gegen} importiert (${da.nr}) – Import-ID wird ergänzt`; }
  }

  // Saldo-Abgleich (nur Exporte mit Saldo, z. B. VW)
  const abgleich = exporte.filter((e) => e.konto && e.saldo != null).map((e) => {
    const k = e.konto; const m = dateien[k.datei];
    const bis = e.bis || e.zeilen.reduce((x, z) => (z.datum > x ? z.datum : x), '');
    const teile = new Map();
    for (const t of m.buchungen) { const p = teilVon(t); if (p) teile.set(p, [...(teile.get(p) || []), t]); }
    const datei = r2((k.anfang || 0) + m.buchungen.filter((b) => b.datum && b.datum > (k.stichtag || '') && b.datum <= bis)
      .reduce((x, b) => x + (signiertFuer(b, k.name, teile) ?? 0), 0));
    const eigene = zeilen.filter((l) => l.konto?.name === k.name);
    const nach = (l) => !k.stichtag || l.zeile.datum > k.stichtag;
    const neu = r2(eigene.filter((l) => nach(l) && l.status === 'neu').reduce((x, l) => x + l.zeile.betrag, 0));
    const fehlen = (nurDatei[k.name] || []).filter((b) => !b.vorStichtag && b.datum <= bis);
    const fehlenSumme = r2(fehlen.reduce((x, b) => x + b.betrag, 0));
    const differenz = r2(e.saldo - datei);
    return { konto: k.name, bis, bank: e.saldo, datei, differenz, neu, fehlenSumme, unklar: eigene.filter((l) => l.status === 'vielleicht').length, rest: r2(differenz - neu + fehlenSumme) };
  });

  return { exporte: exporte.map((e) => ({ name: e.name, format: e.format, konto: e.konto?.name || null, anzahl: e.zeilen?.length || 0, von: e.von, bis: e.bis, saldo: e.saldo, fehler: e.fehler || null })), zeilen, abgleich, nurDatei };
}

/**
 * Was beim Übernehmen geschrieben wird, je Datei.
 * - neu: neue Buchungen (Überweisungen zwischen zwei Dateien des Bereichs in beide, mit Gegenstück)
 * - ergaenzen: Import-ID an vorhandene Buchungen anhängen
 * @param {Function} artFuer (buchung, datei) → Art aus Sicht der Datei
 * @param {Function} nrFuer () → nächste freie Nummer (über alle betroffenen Dateien)
 */
export function plan(ergebnis, dateien, { artFuer, nrFuer }) {
  const konten = kontenListe(dateien);
  const eigenIn = (name) => konten.find((k) => k.name === name && k.eigen === 'Ja')?.datei || null;
  const p = Object.fromEntries(Object.keys(dateien).map((d) => [d, { neu: [], ergaenzen: [] }]));
  for (const l of ergebnis.zeilen) {
    if (!l.auswahl) continue;
    if (l.status === 'eingetragen' || (l.status === 'vielleicht' && l.entscheidung === 'gleich')) {
      const nrs = l.status === 'vielleicht' ? [l.kandidat.nr] : l.treffer;
      for (const nr of nrs) {
        const ziel = p[l.konto.datei].ergaenzen;
        const x = ziel.find((e) => e.nr === nr);
        if (x) x.ids.push(l.id); else ziel.push({ nr, ids: [l.id] });
      }
      continue;
    }
    if (l.status !== 'neu' && !(l.status === 'vielleicht' && l.entscheidung === 'neu')) continue;
    const b = { ...l.buchung, importId: [l.id, l.gegenId].filter(Boolean).join(' ') };
    const ziele = [...new Set([l.konto.datei, eigenIn(b.von), eigenIn(b.nach)].filter(Boolean))];
    const nr = nrFuer();
    for (const d of ziele) {
      const anderes = ziele.find((x) => x !== d);
      p[d].neu.push({ ...b, nr, art: artFuer(b, d), topf: d === 'gemeinsam' && ziele.includes('tagesgeld') ? null : b.topf,
        gegenstueck: anderes ? `${anderes === 'tagesgeld' ? 'Tagesgeld' : 'Gemeinsam'} ${nr}` : null });
    }
  }
  return p;
}
