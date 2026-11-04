// Privater Bereich („Meins“): Rechnungen ohne Browser, ohne Fremdbibliotheken.
// Die private Datei (JSON im eigenen OneDrive) enthält nur, was nicht in die gemeinsame Datei gehört:
// Giro-Anfangsstand, eigene Buchungen (Giro/PayPal), eigene feste Kosten, Limits, Erinnerungen,
// Ziele für Töpfe, Challenges und die Ruhmeshalle.
// Töpfe, Tagesgeld und Überweisungen kommen immer aus der gemeinsamen Datei (eine Quelle).
// Hier stehen keine persönlichen Daten.

import { monatsansicht, toepfeStand, monatPlus, monatsende, istAusgleich } from './rechnen.js';

const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const summe = (arr, f = (x) => x) => r2(arr.reduce((s, x) => s + (f(x) || 0), 0));
const norm = (s) => (s || '').toLowerCase().replace(/\([^)]*\)/g, '').replace(/[^a-z0-9äöüß]+/g, ' ').trim();
const tagNr = (iso) => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 864e5;
const aktivIn = (p, monat) => (!p.ab || p.ab <= monat) && (!p.bis || p.bis >= monat);
const MONATSNAMEN = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];

export const VERSION_DATEI = 1;
export const WEGE = ['Giro', 'PayPal'];
// Vorschläge für eigene Bereiche (frei änderbar, keine Beträge)
export const BEREICHE_VORSCHLAG = ['Lebensmittel', 'Drogerie', 'Kleidung', 'Essen & Trinken', 'Freizeit', 'Geschenke', 'Kinder',
  'Auto', 'Gesundheit & Beauty', 'Abos & Telefon', 'Wohnen', 'Sonstiges'];

/** Leere private Datei. */
export function leer(person, giroKonto, heute) {
  return {
    hinweis: 'Privater Bereich von „Unser Geld“. Liegt nur in deinem OneDrive und ist nicht geteilt. Wird von der App gepflegt.',
    version: VERSION_DATEI, person, angelegt: heute, geaendert: null,
    giro: { konto: giroKonto, anfang: null, stichtag: null },
    buchungen: [], feste: [], limits: [], erinnerungen: [], meineToepfe: null, ziele: {},
    challenges: [], ruhm: [], naechsteNr: 1,
  };
}

/** Fehlende Felder ergänzen (ältere oder von Hand geänderte Dateien). */
export function vervollstaendigen(d) {
  const l = leer(d.person, d.giro?.konto, d.angelegt);
  for (const k of Object.keys(l)) if (d[k] == null) d[k] = l[k];
  d.giro = { ...l.giro, ...d.giro };
  return d;
}

export function neueNr(d, praefix = 'P') {
  const n = d.naechsteNr || 1; d.naechsteNr = n + 1;
  return praefix + String(n).padStart(4, '0');
}

/** Das private Konto einer Person laut gemeinsamer Datei (Rolle „privat“). */
export const privatesKonto = (modell, person) => modell.konten.find((k) => k.rolle === 'privat' && k.inhaber === person)?.name || null;
/** Personen mit privatem Konto in der gemeinsamen Datei. */
export const personen = (modell) => [...new Set(modell.konten.filter((k) => k.rolle === 'privat' && k.inhaber).map((k) => k.inhaber))];

// ---------------------------------------------------------------- Fälligkeit fester Kosten

function faelligMonatAus(text) {
  if (!text) return null;
  const t = String(text).toLowerCase();
  const m = t.match(/(\d{1,2})\.(\d{1,2})\./);
  if (m) return { monat: Number(m[2]), tag: Number(m[1]) };
  const i = MONATSNAMEN.findIndex((n) => t.includes(n));
  if (i >= 0) return { monat: i + 1, tag: null };
  if (/monatsanfang/.test(t)) return { monat: null, tag: 1 };
  return null;
}
const ABSTAND = { monatlich: 1, vierteljährlich: 3, halbjährlich: 6, jährlich: 12 };

/**
 * Ist ein fester Posten im Monat fällig?
 * Gemeinsame Kosten: rhythmus, faellig („03.07.“, „Oktober“, „Monatsanfang“), ab/bis, tag.
 * Eigene feste Kosten: rhythmus, startMonat ('JJJJ-MM', erster fälliger Monat), ab/bis, tag.
 * @returns {{faellig: boolean|null, tag: number|null}}  null = Monat unbekannt
 */
export function faelligkeit(p, monat) {
  if (!aktivIn(p, monat)) return { faellig: false, tag: null };
  const f = faelligMonatAus(p.faellig);
  const tag = p.tag || f?.tag || null;
  const n = ABSTAND[(p.rhythmus || 'monatlich').toLowerCase()] || 1;
  if (n === 1) return { faellig: true, tag };
  const mon = Number(monat.slice(5, 7));
  let anker = p.startMonat ? Number(p.startMonat.slice(5, 7)) : f?.monat || (p.ab ? Number(p.ab.slice(5, 7)) : null);
  if (!anker) return { faellig: null, tag };
  return { faellig: ((mon - anker) % n + 12) % 12 === 0, tag };
}

// ---------------------------------------------------------------- Buchungen

/** Wirkung einer privaten Buchung aufs Giro (PayPal wird vom Giro abgebucht). */
export const wirkung = (b) => (b.art === 'Einnahme' ? b.betrag : -b.betrag);
/** Wirkung einer Buchung der gemeinsamen Datei aufs Giro. */
const wirkungGemeinsam = (b, giro) => (b.nach === giro ? b.betrag : 0) - (b.von === giro ? b.betrag : 0);

/** Giro-Stand bis einschließlich `bis` ('JJJJ-MM-TT'). null, solange kein Anfangsstand eingetragen ist. */
export function giroStand(modell, p, bis) {
  const g = p.giro;
  if (g.anfang == null || !g.stichtag) return null;
  const nach = (d) => d && d > g.stichtag && d <= bis;
  const privat = summe(p.buchungen.filter((b) => nach(b.datum)), wirkung);
  const gemeinsam = summe(modell.buchungen.filter((b) => nach(b.datum) && (b.von === g.konto || b.nach === g.konto)), (b) => wirkungGemeinsam(b, g.konto));
  return r2(g.anfang + privat + gemeinsam);
}

/** Alle Bewegungen aufs/vom Giro in einem Zeitraum, aus beiden Dateien, neu → alt. */
export function giroBewegungen(modell, p, von, bis) {
  const g = p.giro.konto;
  const liste = [
    ...p.buchungen.filter((b) => b.datum >= von && b.datum <= bis).map((b) => ({ quelle: 'privat', b, datum: b.datum, betrag: wirkung(b) })),
    ...modell.buchungen.filter((b) => b.datum && b.datum >= von && b.datum <= bis && (b.von === g || b.nach === g))
      .map((b) => ({ quelle: 'gemeinsam', b, datum: b.datum, betrag: wirkungGemeinsam(b, g) })),
  ];
  return liste.sort((x, y) => y.datum.localeCompare(x.datum) || String(y.b.nr).localeCompare(String(x.b.nr)));
}

// ---------------------------------------------------------------- Monat: was bleibt frei?

/**
 * Monatsrechnung für den privaten Bereich.
 * Frei = Einnahmen − Überweisungen an gemeinsame Konten − feste Kosten − eigene Ausgaben + sonstige Einnahmen.
 * Ausgaben, die aus einem Topf bezahlt wurden („aus Topf“), zählen nicht gegen das Freie.
 */
export function monatPrivat(modell, p, monat, heute) {
  const person = p.person;
  const giro = p.giro.konto;
  const extern = modell.konten.find((k) => k.rolle === 'extern')?.name || 'Extern';
  const imMonat = (b) => (b.monat || b.datum?.slice(0, 7)) === monat;
  const priv = p.buchungen.filter(imMonat);
  const v = monatsansicht(modell, monat, heute);
  const laufend = heute.slice(0, 7) === monat;
  const heuteTag = Number(heute.slice(8, 10));

  // Einnahmen: geplante aus der gemeinsamen Datei (Konto = eigenes Giro), eigene geplante, dazu tatsächliche
  const einnahmen = [];
  for (const e of modell.einnahmen) {
    if (e.konto !== giro || !aktivIn(e, monat)) continue;
    const privatIst = priv.find((b) => b.art === 'Einnahme' && b.plan === e.id);
    const gemIst = !privatIst && modell.buchungen.find((b) => b.giltFuer === monat && b.nach === giro && b.von === extern
      && (norm(b.posten) === norm(e.posten) || Math.abs(b.betrag - e.betrag) < 0.005));
    const ist = privatIst || gemIst || null;
    einnahmen.push({ quelle: 'gemeinsam', id: e.id, posten: e.posten, plan: r2(e.betrag), tag: e.tag || null, bereich: e.bereich,
      ist, betrag: ist ? r2(ist.betrag) : r2(e.betrag), bemerkung: e.eingang || e.bemerkung || null });
  }
  for (const f of p.feste.filter((x) => x.art === 'Einnahme')) {
    const fl = faelligkeit(f, monat);
    if (!fl.faellig) continue;
    const ist = priv.find((b) => b.art === 'Einnahme' && b.plan === f.id) || null;
    einnahmen.push({ quelle: 'privat', id: f.id, posten: f.posten, plan: r2(f.betrag), tag: fl.tag, bereich: f.bereich, ist, betrag: ist ? r2(ist.betrag) : r2(f.betrag) });
  }

  // Überweisungen an gemeinsame Konten (Soll aus „Aufteilung“, Stand aus den Buchungen)
  const ueberweisungen = v.ueberweisungen.filter((u) => u.person === person)
    .map((u) => ({ ...u, betragMonat: r2(Math.max(u.betrag, u.gebucht)) }));

  // Feste Kosten: aus der gemeinsamen Datei (Konto = eigenes Giro) und eigene
  const feste = [];
  const unbekannt = [];
  const festeQuellen = [
    ...modell.kosten.filter((k) => k.konto === giro && k.status !== 'beendet').map((k) => ({ ...k, quelle: 'gemeinsam' })),
    ...p.feste.filter((f) => f.art !== 'Einnahme').map((f) => ({ ...f, quelle: 'privat' })),
  ];
  for (const k of festeQuellen) {
    const fl = faelligkeit(k, monat);
    if (fl.faellig === false) continue;
    if (fl.faellig == null) { unbekannt.push({ id: k.id, posten: k.posten, betrag: r2(k.betrag), rhythmus: k.rhythmus }); continue; }
    const privatIst = priv.find((b) => b.art === 'Ausgabe' && b.plan === k.id);
    const gemIst = !privatIst && k.quelle === 'gemeinsam' && modell.buchungen.find((b) => b.giltFuer === monat && b.von === giro && b.nach === extern && norm(b.posten) === norm(k.posten));
    const ist = privatIst || gemIst || null;
    feste.push({ quelle: k.quelle, id: k.id, posten: k.posten, plan: r2(k.betrag), tag: fl.tag, bereich: k.bereich, rhythmus: k.rhythmus,
      geschaetzt: k.status === 'geschätzt', ist, betrag: ist ? r2(ist.betrag) : r2(k.betrag) });
  }

  // Eigene Buchungen ohne Planbezug
  const ausgaben = priv.filter((b) => b.art === 'Ausgabe' && !b.plan);
  const sonstigeEin = priv.filter((b) => b.art === 'Einnahme' && !b.plan);
  const ausTopf = ausgaben.filter((b) => b.ausTopf);

  const s = {
    einnahmen: summe(einnahmen, (e) => e.betrag),
    sonstigeEin: summe(sonstigeEin, (b) => b.betrag),
    ueberweisungen: summe(ueberweisungen, (u) => u.betragMonat),
    feste: summe(feste, (f) => f.betrag),
    ausgegeben: summe(ausgaben.filter((b) => !b.ausTopf), (b) => b.betrag),
    ausTopf: summe(ausTopf, (b) => b.betrag),
    paypal: summe(priv.filter((b) => b.weg === 'PayPal' && b.art === 'Ausgabe'), (b) => b.betrag),
  };
  s.frei = r2(s.einnahmen + s.sonstigeEin - s.ueberweisungen - s.feste - s.ausgegeben);
  s.verfuegbarVorAusgaben = r2(s.einnahmen + s.sonstigeEin - s.ueberweisungen - s.feste);

  // Ausgaben nach Bereich, mit Limits
  const nachBereich = new Map();
  for (const b of priv.filter((x) => x.art === 'Ausgabe')) {
    const k = b.bereich || 'Ohne Bereich';
    nachBereich.set(k, r2((nachBereich.get(k) || 0) + b.betrag));
  }
  for (const l of p.limits) if (!nachBereich.has(l.bereich)) nachBereich.set(l.bereich, 0);
  const bereiche = [...nachBereich].map(([bereich, betrag]) => {
    const limit = p.limits.find((l) => l.bereich === bereich)?.betrag ?? null;
    const anteil = limit ? betrag / limit : null;
    return { bereich, betrag, limit, anteil, status: limit == null ? null : anteil >= 1 ? 'erreicht' : anteil >= 0.8 ? 'fast' : 'ok' };
  }).sort((a, b) => (b.limit != null) - (a.limit != null) || b.betrag - a.betrag);

  // Was noch kommt (offene Posten), nach Tag
  const kommt = [
    ...einnahmen.filter((e) => !e.ist).map((e) => ({ art: 'einnahme', tag: e.tag, titel: e.posten, betrag: e.plan, id: e.id, quelle: e.quelle })),
    ...ueberweisungen.filter((u) => !u.erledigt).map((u) => ({ art: 'ueberweisung', tag: u.tag, titel: `Überweisung an ${u.nach}`, betrag: -u.offen, nach: u.nach, wann: u.wann })),
    ...feste.filter((f) => !f.ist).map((f) => ({ art: 'fest', tag: f.tag, titel: f.posten, betrag: -f.plan, id: f.id, quelle: f.quelle, geschaetzt: f.geschaetzt })),
  ].sort((a, b) => (a.tag ?? 99) - (b.tag ?? 99));
  for (const k of kommt) k.ueberfaellig = laufend && k.tag != null && k.tag < heuteTag;

  const stichtag = laufend ? heute : monatsende(monat);
  return {
    monat, laufend, person, giro, einnahmen, ueberweisungen, feste, unbekannt, ausgaben, sonstigeEin, summen: s, bereiche, kommt,
    giroStand: giroStand(modell, p, stichtag), giroStichtag: stichtag,
    tageUebrig: laufend ? Number(monatsende(monat).slice(8, 10)) - heuteTag + 1 : null,
  };
}

/** Limit-Hinweis nach einer neuen Ausgabe: wurde mit dieser Ausgabe eine Schwelle überschritten? */
export function limitHinweis(p, monat, bereich, neuBetrag) {
  const l = p.limits.find((x) => x.bereich === bereich);
  if (!l || !l.betrag) return null;
  const vorher = summe(p.buchungen.filter((b) => b.art === 'Ausgabe' && b.bereich === bereich && (b.monat || b.datum.slice(0, 7)) === monat), (b) => b.betrag);
  const nachher = r2(vorher + neuBetrag);
  if (nachher >= l.betrag && vorher < l.betrag) return { stufe: 'erreicht', bereich, limit: l.betrag, summe: nachher };
  if (nachher >= 0.8 * l.betrag && vorher < 0.8 * l.betrag) return { stufe: 'fast', bereich, limit: l.betrag, summe: nachher };
  return null;
}

// ---------------------------------------------------------------- Töpfe

/** Welche Töpfe „meine“ sind: gespeicherte Auswahl, sonst Töpfe mit dem Namen oder Vertrag der Person. */
export function meineToepfe(modell, p) {
  if (Array.isArray(p.meineToepfe)) return modell.toepfe.filter((t) => p.meineToepfe.includes(t.name)).map((t) => t.name);
  const vertrag = new Set(modell.kosten.filter((k) => k.vertragAuf === p.person && k.topf).map((k) => k.topf));
  return modell.toepfe.filter((t) => vertrag.has(t.name) || norm(t.name).split(' ').includes(norm(p.person))).map((t) => t.name);
}

/** Ziel eines Topfs: Stand, Fortschritt, nötige Sparrate bis zum Wunschdatum. */
export function zielStand(stand, ziel, heute) {
  if (!ziel?.betrag) return null;
  const fehlt = r2(Math.max(0, ziel.betrag - stand));
  let monate = null, rate = null;
  if (ziel.datum && fehlt > 0) {
    const a = heute.slice(0, 7), b = ziel.datum.slice(0, 7);
    monate = (Number(b.slice(0, 4)) - Number(a.slice(0, 4))) * 12 + Number(b.slice(5, 7)) - Number(a.slice(5, 7)) + 1;
    rate = monate > 0 ? r2(fehlt / monate) : fehlt;
  }
  return { betrag: ziel.betrag, datum: ziel.datum || null, stand, fehlt, anteil: Math.max(0, Math.min(1, stand / ziel.betrag)), monate, rate, geschafft: stand >= ziel.betrag - 0.005 };
}

// ---------------------------------------------------------------- Erinnerungen

/** Erinnerungen für einen Monat: eigene + automatische (Überweisungen). Fällig = Tag erreicht oder in den nächsten 3 Tagen. */
export function erinnerungen(p, mp, heute) {
  const monat = mp.monat;
  const heuteTag = heute.slice(0, 7) === monat ? Number(heute.slice(8, 10)) : heute.slice(0, 7) > monat ? 99 : 0;
  const liste = [];
  for (const u of mp.ueberweisungen) {
    liste.push({ id: 'auto-' + u.nach, auto: true, text: `${u.offen > 0 && u.gebucht > 0 ? 'Rest überweisen' : 'Überweisen'} an ${u.nach}: ${u.wann}`,
      tag: u.tag, erledigt: u.erledigt, betrag: u.offen, nach: u.nach, von: u.von, soll: u.betrag });
  }
  for (const e of p.erinnerungen.filter((x) => x.aktiv !== false)) {
    liste.push({ id: e.id, auto: false, text: e.text, tag: e.tag, erledigt: !!e.erledigt?.[monat] });
  }
  for (const e of liste) {
    e.ueberfaellig = !e.erledigt && e.tag != null && e.tag < heuteTag;
    e.faellig = !e.erledigt && (e.tag == null || e.tag <= heuteTag + 3);
  }
  return liste.sort((a, b) => (a.erledigt - b.erledigt) || (a.tag ?? 99) - (b.tag ?? 99));
}

/** Kalender-Datei (.ics) für eine monatlich wiederkehrende Erinnerung. */
export function ics({ titel, tag, start, uid, text }) {
  const ab = start.replace(/-/g, '').slice(0, 6) + String(Math.min(tag, 28)).padStart(2, '0');
  const esc = (s) => String(s || '').replace(/([,;\\])/g, '\\$1').replace(/\n/g, '\\n');
  const jetzt = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z';
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Unser Geld//DE', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT',
    `UID:${uid}@unser-geld`, `DTSTAMP:${jetzt}`, `DTSTART;VALUE=DATE:${ab}`, `RRULE:FREQ=MONTHLY;BYMONTHDAY=${tag}`,
    `SUMMARY:${esc(titel)}`, `DESCRIPTION:${esc(text || 'Erinnerung aus „Unser Geld“')}`,
    'BEGIN:VALARM', 'TRIGGER:PT9H', 'ACTION:DISPLAY', `DESCRIPTION:${esc(titel)}`, 'END:VALARM',
    'END:VEVENT', 'END:VCALENDAR'].join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------- Kontoauszug (VW-Giro) abgleichen

/**
 * Gleicht einen gelesenen VW-Export (kern/csv.js leseCsv) mit den eigenen Buchungen
 * und den Buchungen der gemeinsamen Datei ab, die das Giro betreffen. Schreibt nichts.
 * @returns {{zeilen, fehlen, saldo}}
 *  zeilen: je Bankzeile { z, status: 'eingetragen'|'neu'|'vor-stichtag', treffer: [...] , gemeinsamKonto }
 *  fehlen: eingetragen, aber nicht auf dem Auszug (im Zeitraum des Auszugs)
 */
export function abgleichen(csv, modell, p) {
  const g = p.giro;
  const kandidaten = [
    ...p.buchungen.map((b) => ({ quelle: 'privat', b, datum: b.datum, betrag: r2(wirkung(b)) })),
    ...modell.buchungen.filter((b) => b.datum && (b.von === g.konto || b.nach === g.konto))
      .map((b) => ({ quelle: 'gemeinsam', b, datum: b.datum, betrag: r2(wirkungGemeinsam(b, g.konto)) })),
  ];
  const benutzt = new Set();
  const importIds = new Set(p.buchungen.map((b) => b.importId).filter(Boolean));
  const zeilen = csv.zeilen.map((z) => ({ z, id: abgleichId(z), status: null, treffer: [] }));
  // 0. Schon übernommen (Import-ID)
  for (const r of zeilen) if (importIds.has(r.id)) { r.status = 'eingetragen'; const k = kandidaten.find((x) => x.b.importId === r.id); if (k) { r.treffer = [k]; benutzt.add(k); } }
  // 1. Vor dem Stichtag
  for (const r of zeilen) if (!r.status && g.stichtag && r.z.datum <= g.stichtag) r.status = 'vor-stichtag';
  // 2. Gleicher Betrag, ±4 Tage, nächstes Datum zuerst
  const abstand = (k, r) => Math.abs(tagNr(k.datum) - tagNr(r.z.datum));
  for (const r of zeilen) {
    if (r.status) continue;
    const pp = /paypal/i.test(r.z.info);
    const istPP = (k) => (k.b.weg === 'PayPal' ? 1 : 0);
    const pass = kandidaten.filter((k) => !benutzt.has(k) && Math.abs(k.betrag - r.z.betrag) < 0.005 && abstand(k, r) <= 4)
      .sort((a, b) => abstand(a, r) - abstand(b, r) || (pp ? istPP(b) - istPP(a) : istPP(a) - istPP(b)));
    if (pass[0]) { r.status = 'eingetragen'; r.treffer = [pass[0]]; benutzt.add(pass[0]); }
  }
  // 3. Zusammengerechnet: 2–3 eigene Buchungen (z. B. PayPal) = eine Bankzeile, ±5 Tage
  for (const r of zeilen) {
    if (r.status) continue;
    const pool = kandidaten.filter((k) => !benutzt.has(k) && k.quelle === 'privat' && Math.sign(k.betrag) === Math.sign(r.z.betrag) && abstand(k, r) <= 5);
    let gefunden = null;
    for (let i = 0; i < pool.length && !gefunden; i++) for (let j = i + 1; j < pool.length && !gefunden; j++) {
      if (Math.abs(pool[i].betrag + pool[j].betrag - r.z.betrag) < 0.005) gefunden = [pool[i], pool[j]];
      for (let k = j + 1; k < pool.length && !gefunden; k++) if (Math.abs(pool[i].betrag + pool[j].betrag + pool[k].betrag - r.z.betrag) < 0.005) gefunden = [pool[i], pool[j], pool[k]];
    }
    if (gefunden) { r.status = 'eingetragen'; r.treffer = gefunden; gefunden.forEach((k) => benutzt.add(k)); }
  }
  // 4. Rest ist neu. Hinweis, wenn ein gemeinsames Konto beteiligt ist (gehört dann in die gemeinsame Datei)
  const endung = (s) => (String(s || '').match(/(\d{4})\D*$/) || [])[1] || null;
  for (const r of zeilen) {
    if (r.status) continue;
    r.status = 'neu';
    const kt = modell.konten.find((k) => k.name !== g.konto && endung(k.ibanEnde) && ['gemeinsam', 'Sparen'].includes(k.rolle)
      && new RegExp(`(IBAN:\\s*[A-Z]{2}\\d{2}\\w*|Kto:\\s*\\d*)${endung(k.ibanEnde)}\\b`).test(r.z.info));
    if (kt) r.gemeinsamKonto = kt.name;
  }
  const fehlen = kandidaten.filter((k) => !benutzt.has(k) && csv.von && csv.bis && k.datum >= csv.von && k.datum <= csv.bis
    && (!g.stichtag || k.datum > g.stichtag)).sort((a, b) => a.datum.localeCompare(b.datum));
  const stand = csv.bis ? giroStand(modell, p, csv.bis) : null;
  return { zeilen, fehlen, saldo: { bank: csv.saldo, app: stand, differenz: stand == null ? null : r2(csv.saldo - stand), bis: csv.bis } };
}
/** Eindeutige Kennung einer Bankzeile (ohne persönliche Daten im Klartext). */
export function abgleichId(z) {
  const s = `${z.datum}|${z.betrag.toFixed(2)}|${z.info}`;
  let h1 = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) { h1 ^= s.charCodeAt(i); h1 = Math.imul(h1, 0x01000193) >>> 0; }
  return 'G-' + z.datum.replace(/-/g, '') + '-' + h1.toString(36);
}

/** Bereich für eine neue Buchung vorschlagen: zuletzt benutzter Bereich bei gleichem Wofür. */
export function bereichVorschlag(p, posten) {
  const n = norm(posten);
  if (!n) return null;
  const b = [...p.buchungen].reverse().find((x) => x.bereich && (norm(x.posten) === n || norm(x.posten).startsWith(n) || n.startsWith(norm(x.posten))));
  return b?.bereich || null;
}

// ---------------------------------------------------------------- Challenges

export const CHALLENGE_ARTEN = {
  ziel: { titel: 'Sparziel', kurz: 'Betrag und Wunschdatum, mit Fortschrittsbalken und nötiger Rate.' },
  wochen52: { titel: '52-Wochen-Challenge', kurz: 'Woche 1 → 1 €, Woche 2 → 2 € … Woche 52 → 52 €. Zusammen 1.378 €.' },
  limit: { titel: 'Monatslimit mit Belohnung', kurz: 'Unter deinem Limit für einen Bereich bleiben, dann gibt es deine Belohnung.' },
  nospend: { titel: 'No-Spend-Tage', kurz: 'Tage ohne eigene Ausgaben sammeln, der Kalender zeigt die geschafften.' },
};

/** Ein Tag ohne eigene Ausgaben? Feste Kosten (mit Planbezug) und Ausgaben aus einem Topf zählen nicht. */
export const ausgabenAmTag = (p, datum) => p.buchungen.filter((b) => b.art === 'Ausgabe' && b.datum === datum && !b.plan);

/**
 * Stand einer Challenge.
 * @returns {{anteil:number, text:string, geschafft:boolean, vorbei:boolean, auszeichnungen:Array<{key,art,stufe,titel}>, ...}}
 */
export function challengeStand(c, { modell, p, heute }) {
  const aus = [];
  const topfStand = (name) => toepfeStand(modell, heute).find((t) => t.name === name)?.stand ?? null;
  if (c.typ === 'ziel') {
    const eigen = summe(c.einzahlungen || [], (e) => e.betrag);
    const stand = c.topf ? (topfStand(c.topf) ?? 0) : eigen;
    const zs = zielStand(stand, { betrag: c.betrag, datum: c.datum }, heute);
    if (zs.anteil >= 0.5) aus.push({ key: 'halb', art: 'medaille', stufe: 'silber', titel: `Halbzeit: ${c.name}` });
    if (zs.geschafft) aus.push({ key: 'ziel', art: 'pokal', stufe: 'gold', titel: `Sparziel erreicht: ${c.name}` });
    return { ...zs, anteil: zs.anteil, geschafft: zs.geschafft, vorbei: zs.geschafft, auszeichnungen: aus };
  }
  if (c.typ === 'wochen52') {
    const erledigt = new Set(c.erledigt || []);
    const betrag = [...erledigt].reduce((s, w) => s + w, 0);
    const start = c.start || heute;
    const aktuelleWoche = Math.max(1, Math.min(52, Math.floor((tagNr(heute) - tagNr(start)) / 7) + 1));
    const faelligBis = Array.from({ length: aktuelleWoche }, (_, i) => i + 1).filter((w) => !erledigt.has(w));
    for (const [n, stufe, t] of [[13, 'bronze', 'Ein Viertel geschafft'], [26, 'silber', 'Halbzeit'], [39, 'gold', 'Drei Viertel geschafft']]) {
      if (erledigt.size >= n) aus.push({ key: 'w' + n, art: 'medaille', stufe, titel: `52 Wochen: ${t} (${n} Wochen)` });
    }
    if (erledigt.size >= 52) aus.push({ key: 'fertig', art: 'pokal', stufe: 'gold', titel: '52-Wochen-Challenge geschafft' });
    return { anteil: erledigt.size / 52, betrag, gesamt: 1378, wochen: erledigt.size, aktuelleWoche, offen: faelligBis, geschafft: erledigt.size >= 52, vorbei: erledigt.size >= 52, auszeichnungen: aus };
  }
  if (c.typ === 'limit') {
    const ausgegeben = summe(p.buchungen.filter((b) => b.art === 'Ausgabe' && b.bereich === c.bereich && (b.monat || b.datum.slice(0, 7)) === c.monat), (b) => b.betrag);
    const vorbei = monatsende(c.monat) < heute;
    const rest = r2(c.betrag - ausgegeben);
    const geschafft = vorbei && ausgegeben <= c.betrag + 0.005;
    if (geschafft) aus.push({ key: 'limit', art: 'pokal', stufe: 'gold', titel: `Unter dem Limit geblieben: ${c.bereich}` });
    return { anteil: Math.min(1, c.betrag ? ausgegeben / c.betrag : 0), ausgegeben, rest, geschafft, vorbei, verfehlt: vorbei && !geschafft || (!vorbei && ausgegeben > c.betrag + 0.005), auszeichnungen: aus };
  }
  if (c.typ === 'nospend') {
    const ende = monatsende(c.monat);
    const tage = [];
    for (let d = c.monat + '-01'; d <= ende; d = naechsterTag(d)) {
      const vergangen = d < heute;
      const leer = ausgabenAmTag(p, d).length === 0;
      tage.push({ datum: d, vergangen, heute: d === heute, ohne: (vergangen || d === heute) && leer, mit: !leer });
    }
    const geschafftTage = tage.filter((t) => t.vergangen && t.ohne).length;
    const vorbei = ende < heute;
    if (geschafftTage >= Math.ceil(c.ziel / 2)) aus.push({ key: 'halb', art: 'medaille', stufe: 'bronze', titel: `No-Spend: Hälfte geschafft (${monatKurz(c.monat)})` });
    if (geschafftTage >= c.ziel) aus.push({ key: 'ziel', art: 'pokal', stufe: 'gold', titel: `${c.ziel} No-Spend-Tage im ${monatKurz(c.monat)}` });
    return { anteil: Math.min(1, geschafftTage / c.ziel), tage, geschafftTage, geschafft: geschafftTage >= c.ziel, vorbei: vorbei || geschafftTage >= c.ziel, auszeichnungen: aus };
  }
  return { anteil: 0, auszeichnungen: [], geschafft: false, vorbei: false };
}
const naechsterTag = (d) => new Date(tagNr(d) * 864e5 + 864e5).toISOString().slice(0, 10);
const MK = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
const monatKurz = (m) => MK[Number(m.slice(5, 7)) - 1] + ' ' + m.slice(0, 4);

/** Neue Auszeichnungen, die noch nicht in der Ruhmeshalle stehen. */
export function neueAuszeichnungen(p, ctx) {
  const da = new Set(p.ruhm.map((r) => r.id));
  const neu = [];
  for (const c of p.challenges) {
    for (const a of challengeStand(c, ctx).auszeichnungen) {
      const id = `${c.id}:${a.key}`;
      if (!da.has(id)) neu.push({ id, challenge: c.id, art: a.art, stufe: a.stufe, titel: a.titel, datum: ctx.heute });
    }
  }
  return neu;
}

export { monatPlus, istAusgleich };

// ---------------------------------------------------------------- Satz verstehen (privat)
import { betragFinden, datumFinden } from './verstehen.js';
/**
 * „12,50 Zalando mit PayPal“, „gestern 30 € Friseur“, „Gehalt 565,30 bekommen“ → Formularfelder.
 * Feste Regeln, keine KI. Speichern passiert erst nach Bestätigung im Formular.
 */
export function verstehenPrivat(satz, p, heute) {
  let rest = ' ' + (satz || '').trim() + ' ';
  const weg = (s) => { if (s) rest = rest.replace(s, ' '); };
  const felder = {}; const erkannt = [];
  const d = datumFinden(rest, heute);
  if (d) { felder.datum = d.wert; const i = rest.toLowerCase().indexOf(d.treffer); if (i >= 0) rest = rest.slice(0, i) + ' ' + rest.slice(i + d.treffer.length); erkannt.push('Datum'); }
  const b = betragFinden(rest);
  if (b) { felder.betrag = b.wert; weg(b.treffer); erkannt.push('Betrag'); }
  const pp = rest.match(/\s(?:mit|über|per|via|bei)?\s*pay\s?pal\b/i);
  if (pp) { felder.weg = 'PayPal'; weg(pp[0]); erkannt.push('PayPal'); }
  const giro = rest.match(/\s(?:vom|mit|über|per)\s+(?:dem\s+)?(?:giro\w*|konto|karte)\b/i);
  if (giro) { felder.weg = felder.weg || 'Giro'; weg(giro[0]); }
  const ein = rest.match(/\s(bekommen|erhalten|eingegangen|gutschrift|einnahme|zurück(?:bekommen|erstattet)?|erstattet)\b/i);
  if (ein) { felder.art = 'Einnahme'; weg(ein[0]); erkannt.push('Einnahme'); }
  const aus = rest.match(/\s(ausgegeben|bezahlt|gezahlt|gekauft|ausgabe)\b/i);
  if (aus) { felder.art = felder.art || 'Ausgabe'; weg(aus[0]); }
  const posten = rest.replace(/\s(für|beim?|im|in|an|den|die|das|der|euro|€|und)\s/gi, ' ').replace(/[,.;:!?„“"]/g, ' ').replace(/\s+/g, ' ').trim();
  if (posten) { felder.posten = posten.charAt(0).toUpperCase() + posten.slice(1); erkannt.push('Wofür'); }
  const bereich = bereichVorschlag(p, felder.posten);
  if (bereich) { felder.bereich = bereich; erkannt.push('Bereich'); }
  const fehlt = [];
  if (!felder.betrag) fehlt.push('Betrag');
  if (!felder.posten) fehlt.push('Wofür');
  return { felder, erkannt, fehlt };
}
