// Umzug: alte Dateien → neues einheitliches Format (Bauplan Phase 2).
// Reine Funktionen: bekommen die alten Modelle (kern/lesen.js, kern/kai.js, Natalies JSON) und liefern
// die Daten für kern/vorlage.js dateiBauen() plus eine Vergleichsliste alt ↔ neu.
// Läuft im Browser mit der Anmeldung der Person; Daten verlassen das Gerät nur Richtung OneDrive.

import { BEREICH_ZU_KATEGORIE, KATEGORIEN_START, SPARPLAN, VORHER_AKTIV, WIDGETS_START, FORMAT_VERSION } from './format.js';
import { kern, r2, kontostand as neuStand } from './einheit.js';
import { SPRIT_MUSTER, toepfeStand as altToepfe, kontostand as altKontostand, monatsansicht } from './rechnen.js';
import * as P from './privat.js';
import * as K from './kai.js';

const MONATSNAMEN = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];
/** „03.07.“ → {monat:7, tag:3}, „Oktober“ → {monat:10}, „Monatsanfang“ → {tag:1}. */
export function faelligAus(text) {
  if (text == null || text === '') return {};
  if (typeof text === 'number') return text >= 1 && text <= 12 ? { monat: text } : {};
  const t = String(text).toLowerCase();
  const m = t.match(/(\d{1,2})\.(\d{1,2})\./);
  if (m) return { monat: +m[2], tag: +m[1] };
  const i = MONATSNAMEN.findIndex((n) => t.includes(n));
  if (i >= 0) return { monat: i + 1 };
  if (/monatsanfang/.test(t)) return { tag: 1 };
  if (/monatsende/.test(t)) return { tag: 28 };
  return {};
}
const kategorieAus = (bereich) => BEREICH_ZU_KATEGORIE[bereich] ?? null;
const istSprit = (b) => SPRIT_MUSTER.test(`${b.posten || ''} ${b.beschreibung || ''}`);
const einstellungen = (bereich, extra = []) => [
  { schluessel: 'Format-Version', wert: String(FORMAT_VERSION), erklaerung: 'Version des Dateiformats' },
  { schluessel: 'Bereich', wert: bereich, erklaerung: 'gemeinsam, tagesgeld, natalie oder kai' },
  { schluessel: 'Funktionen', wert: '', erklaerung: 'Eingeschaltete Zusatzfunktionen (durch Komma getrennt). Leer = nur der feste Kern.' },
  { schluessel: 'Widgets', wert: (WIDGETS_START[bereich] || []).join(','), erklaerung: 'Reihenfolge der Widgets in der Übersicht' },
  { schluessel: 'Vorher aktiv', wert: (VORHER_AKTIV[bereich] || []).join(','), erklaerung: 'Was vor dem Umbau (16.11.2026) an war – für „Wie vor dem Umbau“' },
  { schluessel: 'Umzug', wert: new Date().toISOString().slice(0, 10), erklaerung: 'Daten aus der alten Datei übernommen am' },
  ...extra,
];
/** Art aus Sicht einer Datei: welche Seite ist ein eigenes Konto? */
function artFuer(b, eigene) {
  if (['In Topf legen', 'Aus Topf nehmen'].includes(b.art)) return b.art;
  const v = eigene.has(b.von), n = eigene.has(b.nach);
  return v && n ? 'Umbuchung' : n ? 'Einnahme' : v ? 'Ausgabe' : b.art;
}
const neuesKonto = (k, eigen, frei) => ({
  id: k.id, name: k.name, bank: k.bank, art: k.art, inhaber: k.inhaber, rolle: k.rolle, eigen: eigen ? 'Ja' : 'Nein', frei: frei ? 'Ja' : 'Nein',
  endung: k.ibanEnde ?? k.endung ?? null, csv: k.csv === true || k.csv === 'ja' ? 'Ja' : 'Nein', anfang: eigen ? k.anfang : null, stichtag: eigen ? k.stichtag : null,
  status: k.status === 'geschlossen' ? 'geschlossen' : 'aktiv', bemerkung: k.bemerkung,
});

// ---------------------------------------------------------------- Gemeinsam + Tagesgeld (aus Buchhaltung_Kai_Natalie.xlsx)

/**
 * Teilt die alte gemeinsame Datei auf: Tagesgeld (alle Sparkonten mit Töpfen) und Gemeinsam (Gemeinschaftskonto, Visa, neue c24-Konten).
 * @param {object} alt aus kern/lesen.js leseBuchhaltung()
 * @param {{neueKonten?: Array, vwSchluss?: string, sparratenFormel?: boolean}} opt
 */
export function gemeinsamUndTagesgeld(alt, opt = {}) {
  const spar = new Set(alt.konten.filter((k) => k.rolle === 'Sparen').map((k) => k.name));
  const gem = new Set(alt.konten.filter((k) => k.rolle === 'gemeinsam').map((k) => k.name));
  const hinweise = [];
  const zahler = (b) => b.gezahltVon || alt.konten.find((k) => k.name === b.von && k.rolle === 'privat')?.inhaber || null;
  const basis = (b) => ({
    nr: b.nr, datum: b.datum, giltFuer: b.giltFuer, von: b.von, nach: b.nach, betrag: b.betrag,
    kategorie: istSprit(b) && b.art !== 'Einnahme' ? 'Sprit' : kategorieAus(b.bereich), posten: b.posten, beschreibung: b.beschreibung,
    kommentar: b.kommentar, quelle: b.quelle, importId: b.importId, gezahltVon: zahler(b),
  });
  const tg = [], gm = [];
  for (const b of alt.buchungen) {
    const inT = spar.has(b.von) || spar.has(b.nach), inG = gem.has(b.von) || gem.has(b.nach);
    if (inT) tg.push({ ...basis(b), art: artFuer(b, spar), topf: b.topf, gegenstueck: inG ? `Gemeinsam ${b.nr}` : null });
    if (inG || !inT) {
      const x = { ...basis(b), art: artFuer(b, gem), topf: null, gegenstueck: inT ? `Tagesgeld ${b.nr}` : null };
      if (b.topf && !inT) x.kommentar = [b.kommentar, `früher Topf: ${b.topf} (zählte nicht für den Topf)`].filter(Boolean).join(' · ');
      if (!inG && !inT) hinweise.push(`${b.nr}: berührt weder Gemeinschafts- noch Sparkonto, liegt jetzt in Gemeinsam`);
      gm.push(x);
    }
  }
  const kontenFuer = (eigene, freiAuch) => alt.konten.map((k) => neuesKonto(k, eigene.has(k.name), eigene.has(k.name) && freiAuch(k)));
  const vw = opt.vwSchluss || '2027-01-01';
  const gKonten = kontenFuer(gem, () => true).map((k) => (gem.has(k.name) ? { ...k, geschlossenAb: vw, bemerkung: [k.bemerkung, 'ab 01.01.2027 nicht mehr genutzt'].filter(Boolean).join(' · ') } : k));
  for (const n of opt.neueKonten || []) gKonten.push({ eigen: 'Ja', frei: 'Ja', status: 'aktiv', rolle: 'gemeinsam', csv: 'Ja', ...n });
  const tKonten = kontenFuer(spar, () => true);
  const kosten = alt.kosten.map((k) => {
    const f = faelligAus(k.faellig);
    if ((k.rhythmus || 'monatlich') !== 'monatlich' && !f.monat) hinweise.push(`Kosten ${k.id} ${k.posten}: Fälligkeit „${k.faellig ?? ''}“ ohne Monat – bitte prüfen`);
    return { id: k.id, posten: k.posten, kategorie: /^sprit\b/i.test(k.posten || '') ? 'Sprit' : kategorieAus(k.bereich), betrag: k.betrag, rhythmus: k.rhythmus || 'monatlich',
      faellig: f.monat ?? null, tag: k.tag ?? f.tag ?? null, konto: k.konto, topf: k.topf, vertragAuf: k.vertragAuf, status: k.status || 'aktiv', ab: k.ab, bis: k.bis,
      bemerkung: [k.bemerkung, k.faellig && !f.monat && !f.tag ? `Fällig: ${k.faellig}` : null].filter(Boolean).join(' · ') || null };
  });
  const einnahmen = alt.einnahmen.map((e) => ({ id: e.id, posten: e.posten, kategorie: kategorieAus(e.bereich), betrag: e.betrag, rhythmus: 'monatlich', tag: e.tag,
    konto: e.konto, gehoertZu: e.gehoertZu, giltFuer: e.giltFuer, eingang: e.eingang, status: e.status || 'aktiv', ab: e.ab, bis: e.bis, bemerkung: e.bemerkung }));
  const sprit = alt.kosten.filter((k) => /^sprit\b/i.test(k.posten || '') && k.status !== 'beendet');
  const limits = sprit.map((k) => ({ kategorie: 'Sprit', betrag: r2(k.monatswert ?? k.betrag), ab: k.ab, bis: k.bis, bemerkung: 'Sprit-Limit (früher Sprit-Budget); der Posten bleibt in Kosten, damit die Aufteilung gleich bleibt' }));
  const toepfe = alt.toepfe.map((t) => ({ name: t.name, konto: t.konto, kategorie: kategorieAus(t.bereich), sparrate: t.sparrate, automatisch: 'Nein', bedarf: t.bedarf,
    anfang: t.anfang, stichtag: t.stichtag, frei: 'Nein', bemerkung: [t.bemerkung, t.bereich && !kategorieAus(t.bereich) ? `früher Bereich: ${t.bereich}` : null].filter(Boolean).join(' · ') || null }));
  const ks = (eig) => alt.kontostaendeBank.filter((x) => eig.has(x.konto)).map((x) => ({ monat: x.monat, konto: x.konto, bank: x.bank, kommentar: x.kommentar }));
  const kategorien = [...new Set([...KATEGORIEN_START, ...(limits.length ? ['Sprit'] : [])])].map((name) => ({ name, aktiv: 'Ja', angelegt: 'Start' }));
  const sparraten = r2(toepfe.reduce((s, t) => s + (t.sparrate || 0), 0));
  return {
    hinweise,
    tagesgeld: { bereich: 'tagesgeld', titel: 'Tagesgeld und Töpfe', konten: tKonten, buchungen: tg, kosten: [], einnahmen: [], toepfe, limits: [], soll: [],
      kontostaende: ks(spar), regeln: [], kategorien, einstellungen: einstellungen('tagesgeld') },
    gemeinsam: { bereich: 'gemeinsam', titel: 'Gemeinsam – Kai & Natalie', konten: gKonten, buchungen: gm, kosten, einnahmen, toepfe: [], limits, soll: [],
      kontostaende: ks(gem), regeln: [], kategorien,
      einstellungen: einstellungen('gemeinsam', [{ schluessel: 'Sparraten Tagesgeld', wert: String(sparraten), erklaerung: 'Summe der Sparraten aus Tagesgeld.xlsx (für die Aufteilung). Die App hält den Wert aktuell.' }]) },
  };
}

// ---------------------------------------------------------------- Natalie (aus Mein-Bereich.json + gemeinsamer Datei)

/**
 * @param {object} p Natalies JSON (Mein-Bereich.json)
 * @param {object} alt altes gemeinsames Modell (für Giro-Bewegungen, feste Kosten/Einnahmen, Überweisungen)
 * @param {string} heute 'JJJJ-MM-TT'
 */
export function natalie(p, alt, heute) {
  const giro = p.giro.konto;
  const hinweise = [];
  const B = [];
  const planName = (id) => alt.einnahmen.find((e) => e.id === id)?.posten || alt.kosten.find((k) => k.id === id)?.posten || p.feste.find((f) => f.id === id)?.posten || null;
  for (const b of p.buchungen) {
    B.push({ nr: b.nr, datum: b.datum, giltFuer: b.monat || b.datum?.slice(0, 7), art: b.art, von: b.art === 'Einnahme' ? (b.posten || 'Extern') : giro, nach: b.art === 'Einnahme' ? giro : (b.posten || 'Extern'),
      betrag: b.betrag, kategorie: b.bereich || null, posten: b.plan ? planName(b.plan) : null, topf: null, beschreibung: b.posten,
      kommentar: [b.kommentar, b.weg === 'PayPal' ? 'über PayPal' : null, b.ausTopf ? `aus Topf: ${b.ausTopf}` : null].filter(Boolean).join(' · ') || null,
      quelle: b.quelle === 'CSV' ? 'CSV' : 'manuell', importId: b.importId });
  }
  // Bewegungen aus der gemeinsamen Datei, die ihr Giro berühren (zählten bisher für den Giro-Stand)
  const ab = p.giro.stichtag || '0000';
  for (const b of alt.buchungen.filter((x) => x.datum && x.datum > ab && (x.von === giro || x.nach === giro))) {
    const spar = alt.konten.find((k) => k.name === (b.von === giro ? b.nach : b.von))?.rolle === 'Sparen';
    B.push({ nr: 'G' + b.nr.slice(1), datum: b.datum, giltFuer: b.giltFuer, art: b.nach === giro ? 'Einnahme' : 'Ausgabe', von: b.von, nach: b.nach, betrag: b.betrag,
      kategorie: kategorieAus(b.bereich), beschreibung: b.posten || b.beschreibung, kommentar: b.kommentar, quelle: 'Umzug', gezahltVon: b.gezahltVon,
      gegenstueck: `${spar ? 'Tagesgeld' : 'Gemeinsam'} ${b.nr}` });
  }
  const kosten = [
    ...alt.kosten.filter((k) => k.konto === giro && k.status !== 'beendet').map((k) => { const f = faelligAus(k.faellig); return { id: k.id, posten: k.posten, kategorie: kategorieAus(k.bereich),
      betrag: k.betrag, rhythmus: k.rhythmus || 'monatlich', faellig: f.monat ?? null, tag: k.tag ?? f.tag ?? null, konto: giro, status: k.status || 'aktiv', ab: k.ab, bis: k.bis, bemerkung: 'aus der gemeinsamen Datei' }; }),
    ...p.feste.filter((f) => f.art !== 'Einnahme').map((f) => ({ id: f.id, posten: f.posten, kategorie: f.bereich, betrag: f.betrag, rhythmus: f.rhythmus || 'monatlich',
      faellig: f.startMonat ? +f.startMonat.slice(5, 7) : null, tag: f.tag, konto: giro, status: 'aktiv', ab: f.ab, bis: f.bis })),
  ];
  const einnahmen = [
    ...alt.einnahmen.filter((e) => e.konto === giro).map((e) => ({ id: e.id, posten: e.posten, kategorie: kategorieAus(e.bereich), betrag: e.betrag, rhythmus: 'monatlich', tag: e.tag,
      konto: giro, gehoertZu: e.gehoertZu, status: e.status || 'aktiv', ab: e.ab, bis: e.bis, bemerkung: 'aus der gemeinsamen Datei' })),
    ...p.feste.filter((f) => f.art === 'Einnahme').map((f) => ({ id: f.id, posten: f.posten, betrag: f.betrag, rhythmus: f.rhythmus || 'monatlich', tag: f.tag, konto: giro, status: 'aktiv', ab: f.ab, bis: f.bis })),
  ];
  // Soll-Überweisungen aus der Aufteilung (Monate, die die gemeinsame Datei ausrechnet)
  const soll = [];
  for (const monat of Object.keys(alt.aufteilung).sort()) {
    for (const u of monatsansicht(alt, monat, heute).ueberweisungen.filter((x) => x.person === p.person)) soll.push({ monat, von: giro, nach: u.nach, soll: u.betrag, bemerkung: `aus der Aufteilung (${u.wann})` });
  }
  const konten = [
    { id: 'KT1', name: giro, art: 'Giro', inhaber: p.person, rolle: 'privat', eigen: 'Ja', frei: 'Ja', csv: 'Ja', anfang: p.giro.anfang, stichtag: p.giro.stichtag, status: 'aktiv' },
    ...alt.konten.filter((k) => k.name !== giro).map((k) => neuesKonto(k, false, false)),
  ];
  const zusatz = {
    Erinnerungen: p.erinnerungen.map((e) => ({ id: e.id, text: e.text, tag: e.tag, aktiv: e.aktiv === false ? 'Nein' : 'Ja', erledigt: Object.keys(e.erledigt || {}).filter((m) => e.erledigt[m]).join(',') || null })),
    Challenges: p.challenges.map((c) => ({ id: c.id, art: c.art, titel: c.titel, start: c.start, ende: c.ende, status: c.status, daten: JSON.stringify(c) })),
    Ruhmeshalle: p.ruhm.map((r) => ({ datum: r.datum, auszeichnung: r.titel || r.art, challenge: r.challenge, text: r.text })),
  };
  // Ziele der Töpfe gehen in die Tagesgeld-Datei (Entscheidung 16.11.2026)
  const ziele = Object.entries(p.ziele || {}).map(([topf, z]) => ({ topf, ziel: z.betrag, zielBis: z.datum?.slice(0, 7) || null }));
  return {
    hinweise, ziele,
    daten: { bereich: 'natalie', titel: `${p.person} – mein Bereich`, konten, buchungen: B.sort((a, b) => (a.datum || '').localeCompare(b.datum || '')), kosten, einnahmen, toepfe: [],
      limits: p.limits.map((l) => ({ kategorie: l.bereich, betrag: l.betrag })), soll, kontostaende: [], regeln: [],
      einstellungen: einstellungen('natalie', [{ schluessel: 'Person', wert: p.person, erklaerung: 'Wem der Bereich gehört' }]) },
    zusatz,
    neueKategorien: [...new Set([...p.buchungen.map((b) => b.bereich), ...p.limits.map((l) => l.bereich)].filter((x) => x && !KATEGORIEN_START.includes(x)))],
  };
}

// ---------------------------------------------------------------- Kai (aus Kai_Privat.xlsx)

const KAI_KOSTEN_KATEGORIE = { Abos: 'Abos', Auto: 'Auto', Versicherung: 'Versicherungen', Raten: 'Kredite' };
/** @param {object} k aus kern/kai.js leseKai() */
export function kai(k) {
  const hinweise = [];
  const budgets = new Set(K.budgets(k).map((x) => x.posten));
  const B = k.buchungen.map((b) => {
    const ein = b.art === 'Einnahme';
    let gegen = b.gegen || 'Extern';
    if (b.posten === K.GK_POSTEN) gegen = 'Gemeinschaftskonto';
    const istBudget = budgets.has(b.posten);
    return {
      nr: b.nr, datum: b.datum, giltFuer: b.giltFuer, art: b.art,
      von: ein ? gegen : b.konto, nach: ein ? b.konto : b.art === K.UMBUCHUNG ? b.gegen : ['In Topf legen', 'Aus Topf nehmen'].includes(b.art) ? b.konto : gegen,
      betrag: b.betrag, kategorie: istBudget ? b.posten : b.wofuer && !K.WOFUER_EXTRA.includes(b.wofuer) ? b.wofuer : null, bankKategorie: b.kategorie,
      posten: istBudget ? null : b.posten, topf: b.topf,
      beschreibung: [b.beschreibung, b.posten === K.GK_POSTEN && b.gegen && b.gegen !== 'Gemeinschaftskonto' ? b.gegen : null].filter(Boolean).join(' · ') || null,
      kommentar: b.kommentar, quelle: b.quelle, importId: b.importId,
    };
  });
  const eigen = new Set([k.konto.name, ...(k.spartopf ? [k.spartopf.name] : [])]);
  for (const b of B) { const a = artFuer(b, eigen); if (a !== b.art) { hinweise.push(`${b.nr}: Art ${b.art} → ${a}`); b.art = a; } }
  const konten = k.konten.map((x) => ({ id: x.id, name: x.name, bank: x.bank, rolle: x.rolle, eigen: eigen.has(x.name) ? 'Ja' : 'Nein', frei: eigen.has(x.name) ? 'Ja' : 'Nein',
    endung: x.endung, csv: eigen.has(x.name) ? 'Ja' : 'Nein', anfang: eigen.has(x.name) ? x.anfang : null, stichtag: eigen.has(x.name) ? x.stichtag : null, status: 'aktiv' }));
  const kosten = k.kosten.filter((x) => x.bereich !== K.BUDGET).map((x) => ({ id: x.id, posten: x.posten, kategorie: KAI_KOSTEN_KATEGORIE[x.bereich] ?? null, betrag: x.betrag, rhythmus: x.rhythmus,
    faellig: x.faellig, tag: x.tag, konto: x.konto, status: 'aktiv', ab: x.ab, bis: x.bis, bemerkung: [x.bemerkung, x.bereich && !KAI_KOSTEN_KATEGORIE[x.bereich] ? `früher Bereich: ${x.bereich}` : null].filter(Boolean).join(' · ') || null }));
  const limits = k.kosten.filter((x) => x.bereich === K.BUDGET).map((x) => ({ kategorie: x.posten, betrag: x.betrag, ab: x.ab, bis: x.bis, bemerkung: 'früher Monatsbudget' }));
  const einnahmen = k.einnahmen.map((e) => ({ id: e.id, posten: e.posten, betrag: e.betrag, rhythmus: e.rhythmus, faellig: e.faellig, eingang: e.eingang, konto: e.konto, status: 'aktiv', ab: e.ab, bis: e.bis, bemerkung: e.bemerkung }));
  const toepfe = k.toepfe.map((t) => ({ name: t.name, sparrate: t.sparrate, automatisch: 'Ja', ab: t.ab, bis: t.bis, anfang: t.anfang, stichtag: k.konto.stichtag, frei: t.frei ? 'Ja' : 'Nein', bemerkung: t.bemerkung }));
  const soll = k.gemeinsam.filter((g) => g.monat).map((g) => ({ monat: g.monat, von: k.konto.name, nach: 'Gemeinschaftskonto', soll: g.soll, bemerkung: g.bemerkung }));
  const zusatz = {
    'TR Buchungen': k.tr.map((t) => ({ ...t })),
    'Offene Punkte': k.offen.map((o) => ({ nr: o.nr, text: o.punkt, status: o.status, notiz: o.notiz })),
  };
  return {
    hinweise, neueKategorien: [...limits.map((l) => l.kategorie)],
    daten: { bereich: 'kai', titel: 'Kai – mein Bereich', konten, buchungen: B, kosten, einnahmen, toepfe, limits, soll, kontostaende: [], regeln: [],
      einstellungen: einstellungen('kai', [{ schluessel: 'Person', wert: 'Kai', erklaerung: 'Wem der Bereich gehört' }, { schluessel: 'Vorschau Monate', wert: String(k.start.horizont || 3), erklaerung: 'Wie viele Monate die Warnung vorausschaut' }]) },
    zusatz,
  };
}

// ---------------------------------------------------------------- Vergleich alt ↔ neu

const zeile = (was, alt, neu, erklaerung) => ({ was, alt: alt == null ? null : r2(alt), neu: neu == null ? null : r2(neu), gleich: alt == null || neu == null ? alt === neu : Math.abs(alt - neu) < 0.006, erklaerung: erklaerung || null });

/** Gemeinsam/Tagesgeld: Kontostände, Töpfe, „Nicht in Töpfen“. m = neue Modelle (kern/einheit.js leseDatei). */
export function vergleichGemeinsam(alt, neuG, neuT, heute) {
  const L = [];
  for (const k of alt.konten.filter((x) => x.anfang != null && ['gemeinsam', 'Sparen'].includes(x.rolle))) {
    const m = k.rolle === 'Sparen' ? neuT : neuG;
    L.push(zeile(`Kontostand ${k.name}`, altKontostand(alt, k.name, heute), neuStand(m, k.name, heute)));
  }
  const tAlt = altToepfe(alt, heute); const tNeu = kern(neuT, heute).toepfe;
  for (const t of tAlt) L.push(zeile(`Topf ${t.name}`, t.stand, tNeu.find((x) => x.name === t.name)?.stand));
  const sparAlt = alt.konten.filter((k) => k.rolle === 'Sparen').reduce((s, k) => s + (altKontostand(alt, k.name, heute) || 0), 0) - tAlt.reduce((s, t) => s + t.stand, 0);
  L.push(zeile('Nicht in Töpfen', sparAlt, kern(neuT, heute).freiHeute));
  L.push(zeile('Anzahl Buchungen (Gemeinsam + Tagesgeld, doppelte Überweisungen einmal gezählt)', alt.buchungen.length,
    neuG.buchungen.length + neuT.buchungen.filter((b) => !b.gegenstueck).length));
  return L;
}
/** Natalie: Giro-Stand. „Was ist frei?“ wird jetzt wie überall aus dem Kontostand gerechnet (siehe Erklärung). */
export function vergleichNatalie(p, alt, neu, heute) {
  const altGiro = P.giroStand(alt, p, heute);
  const k = kern(neu, heute);
  const mp = P.monatPrivat(alt, p, heute.slice(0, 7), heute);
  return [
    zeile('Giro-Stand heute', altGiro, k.konten[0]?.standHeute),
    zeile('Was ist frei?', mp.summen.frei, k.frei, 'Bisher: Monatsrechnung (Einnahmen − Überweisungen − feste Kosten − Ausgaben dieses Monats). Neu wie in allen Bereichen: Giro-Stand + was noch kommt − was noch abgeht. Unterschied = Geld, das schon vor dem Monat auf dem Giro lag.'),
    zeile('Anzahl eigene Buchungen', p.buchungen.length, neu.buchungen.filter((b) => b.quelle !== 'Umzug').length),
  ];
}
/** Kai: Kontostand c24, Spartopf, Töpfe, Frei heute. Vorschau weicht nur durch Limits (früher Budgets) ab. */
export function vergleichKai(k, neu, heute) {
  const alt = K.rechneKai(k, heute);
  const n = kern(neu, heute, alt.akt);
  const L = [zeile(`Kontostand ${k.konto.name}`, alt.stand, n.konten.find((x) => x.name === k.konto.name)?.standHeute)];
  if (alt.spartopf) L.push(zeile(`Kontostand ${alt.spartopf.name}`, alt.spartopf.stand, n.konten.find((x) => x.name === alt.spartopf.name)?.standHeute));
  for (const t of alt.toepfe) L.push(zeile(`Topf ${t.name}`, t.stand, n.toepfe.find((x) => x.name === t.name)?.stand));
  L.push(zeile('Frei (Konten − Töpfe)', alt.frei, n.freiHeute));
  L.push(zeile('Anzahl Buchungen', k.buchungen.length, neu.buchungen.length));
  return L;
}
