// Rechnet aus dem Modell die Monatsansicht.
// Die Regeln entsprechen den Formeln in der Excel-Datei (Blätter Start, Töpfe,
// Überweisung, Aufteilung). Wo die App mehr kann als Excel (jeder Monat, Prognose),
// steht das am jeweiligen Abschnitt dabei.

export const eur = (n) => (n < 0 ? '−' : '') + Math.abs(n).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const summe = (arr, f = (x) => x) => r2(arr.reduce((s, x) => s + (f(x) || 0), 0));

export function monatsende(monat) {
  const [j, m] = monat.split('-').map(Number);
  const d = new Date(Date.UTC(j, m, 0));
  return d.toISOString().slice(0, 10);
}
export function monatPlus(monat, n) {
  const [j, m] = monat.split('-').map(Number);
  const d = new Date(Date.UTC(j, m - 1 + n, 1));
  return d.toISOString().slice(0, 7);
}
const tagVon = (datum) => Number(datum.slice(8, 10));
const norm = (s) => (s || '').toLowerCase().replace(/\([^)]*\)/g, '').replace(/\s+/g, ' ').trim();
const aktivIn = (posten, monat) => (!posten.ab || posten.ab <= monat) && (!posten.bis || posten.bis >= monat);

/** Kontostand aus Anfangsbestand + Buchungen nach dem Stichtag bis einschließlich `bis`. */
export function kontostand(modell, kontoName, bis) {
  const k = modell.konten.find((x) => x.name === kontoName);
  if (!k) return null;
  const nachStichtag = (b) => b.datum && (!k.stichtag || b.datum > k.stichtag) && b.datum <= bis;
  const zu = summe(modell.buchungen.filter((b) => b.nach === kontoName && nachStichtag(b)), (b) => b.betrag);
  const ab = summe(modell.buchungen.filter((b) => b.von === kontoName && nachStichtag(b)), (b) => b.betrag);
  return r2((k.anfang || 0) + zu - ab);
}

function rolle(modell, kontoName) {
  return modell.konten.find((k) => k.name === kontoName)?.rolle || 'offen';
}

/** Topf-Stände bis einschließlich `bis` (wie Blatt „Töpfe“, nur mit Datumsgrenze). */
export function toepfeStand(modell, bis) {
  const bis_ = (b) => b.datum && b.datum <= bis;
  const sparplan = summe(modell.buchungen.filter((b) => b.topf === 'Sparplan (verteilt)' && bis_(b)), (b) => b.betrag);
  const sparrateSumme = summe(modell.toepfe, (t) => t.sparrate);
  return modell.toepfe.map((t) => {
    const ausSparplan = sparrateSumme ? (t.sparrate / sparrateSumme) * sparplan : 0;
    const einzeln = summe(
      modell.buchungen.filter((b) => b.topf === t.name && bis_(b)),
      (b) => (rolle(modell, b.nach) === 'Sparen' ? b.betrag : 0) - (rolle(modell, b.von) === 'Sparen' ? b.betrag : 0),
    );
    return { ...t, ausSparplan: r2(ausSparplan), einzeln, stand: r2(t.anfang + ausSparplan + einzeln) };
  });
}

/** Alle Monate, die die App anbietet. */
export function monate(modell) {
  const s = new Set(Object.keys(modell.aufteilung));
  for (const b of modell.buchungen) if (b.giltFuer) s.add(b.giltFuer);
  return [...s].sort();
}

function privatesKonto(modell, person) {
  return modell.konten.find((k) => k.rolle === 'privat' && k.inhaber === person)?.name || null;
}

/**
 * Monatsansicht.
 * @param {object} modell aus leseBuchhaltung()
 * @param {string} monat 'JJJJ-MM'
 * @param {string} heute 'JJJJ-MM-TT'
 */
export function monatsansicht(modell, monat, heute) {
  const ende = monatsende(monat);
  const vormonatsEnde = monatsende(monatPlus(monat, -1));
  const laufend = heute.slice(0, 7) === monat;
  const vergangen = ende < heute;
  const stichtagFuerJetzt = laufend ? heute : ende;

  const gem = modell.konten.find((k) => k.rolle === 'gemeinsam' && /giro/i.test(k.art || ''))?.name || 'Gemeinschaftskonto';
  const visa = modell.konten.find((k) => k.rolle === 'gemeinsam' && /kredit|visa/i.test((k.art || '') + k.name))?.name || 'Visa Haushalt';

  // ---- Konten (nur die, die im Tool angezeigt werden und offen sind)
  const konten = modell.konten
    .filter((k) => k.imTool && k.status !== 'geschlossen')
    .map((k) => {
      const bank = modell.kontostaendeBank.find((x) => x.monat === monat && x.konto === k.name)?.bank ?? null;
      const standEnde = kontostand(modell, k.name, ende);
      return {
        name: k.name, rolle: k.rolle, inhaber: k.inhaber, bank: k.bank, bemerkung: k.bemerkung,
        standJetzt: kontostand(modell, k.name, stichtagFuerJetzt), standMonatsende: standEnde,
        bankMonatsende: bank, differenz: bank == null ? null : r2(bank - standEnde),
      };
    });

  // ---- Töpfe
  const toepfe = toepfeStand(modell, stichtagFuerJetzt);
  const toepfeSumme = summe(toepfe, (t) => t.stand);
  const sparkonten = summe(modell.konten.filter((k) => k.rolle === 'Sparen'), (k) => kontostand(modell, k.name, stichtagFuerJetzt));
  const nichtInToepfen = r2(sparkonten - toepfeSumme);

  // ---- Aufteilung / Überweisungen (aus dem Blatt „Aufteilung“, das jeden Monat ausrechnet)
  const a = modell.aufteilung[monat] || null;
  const roh = (anfang) => {
    if (!a) return 0;
    const key = Object.keys(a).find((k) => k.startsWith(anfang));
    return key ? a[key] : 0;
  };
  const wert = (anfang) => r2(roh(anfang));
  const tagNatalie = Math.max(0, ...modell.einnahmen
    .filter((e) => e.gehoertZu === 'Natalie' && aktivIn(e, monat) && e.tag).map((e) => e.tag));
  const sparkonto = modell.konten.find((k) => k.rolle === 'Sparen' && k.status !== 'geschlossen')?.name || 'Tagesgeld Natalie';
  const gebuchtVonNach = (von, nach) => summe(
    modell.buchungen.filter((b) => b.giltFuer === monat && b.von === von && b.nach === nach), (b) => b.betrag);
  // Wer hat eingezahlt? „Gezahlt von“, sonst Inhaber des privaten Quellkontos (wie Spalte Zuordnung).
  const zahler = (b) => b.gezahltVon || modell.konten.find((k) => k.name === b.von && k.rolle === 'privat')?.inhaber || null;
  const eingezahlt = (person, nach) => summe(modell.buchungen.filter((b) => b.giltFuer === monat && b.nach === nach
    && zahler(b) === person
    // Geld für einen bestimmten Topf (z. B. geparktes Elterngeld) ist keine Sparrate
    && !(rolle(modell, nach) === 'Sparen' && b.topf && b.topf !== 'Sparplan (verteilt)')), (b) => b.betrag);
  const kaiPriv = privatesKonto(modell, 'Kai');
  const natPriv = privatesKonto(modell, 'Natalie');
  const ueberweisungen = [
    { person: 'Kai', von: kaiPriv, nach: gem, betrag: r2(wert('Kai → Gemeinschaftskonto') + wert('Kai: schon direkt')), tag: 1, wann: 'bis zum 1.',
      hinweis: wert('Kai: schon direkt') > 0 ? `davon ${eur(wert('Kai: schon direkt'))} schon zurückgelegt (z. B. Elterngeld)` : null },
    { person: 'Kai', von: kaiPriv, nach: sparkonto, betrag: wert('Kai → Tagesgeld'), tag: 1, wann: 'bis zum 1.' },
    { person: 'Natalie', von: natPriv, nach: sparkonto, betrag: wert('Natalie → Tagesgeld'), tag: tagNatalie || 15,
      wann: tagNatalie > 1 ? `wenn ihr Geld da ist, um den ${tagNatalie}.` : 'bis zum 1.', hinweis: 'Sparraten für die Töpfe' },
    { person: 'Natalie', von: natPriv, nach: gem, betrag: wert('Natalie → Gemeinschaftskonto'), tag: tagNatalie || 15,
      wann: tagNatalie > 1 ? `wenn ihr Geld da ist, um den ${tagNatalie}.` : 'bis zum 1.' },
  ].filter((u) => u.betrag > 0)
    .map((u) => {
      const gebucht = eingezahlt(u.person, u.nach);
      return { ...u, gebucht, offen: r2(Math.max(0, u.betrag - gebucht)), erledigt: gebucht >= u.betrag - 0.005 };
    });

  const anteil = a ? {
    einkommenKai: wert('Einkommen Kai'), einkommenNatalie: wert('Einkommen Natalie'),
    anteilKai: roh('Anteil Kai'), anteilNatalie: roh('Anteil Natalie'),
    traegtKai: wert('Kai trägt'), traegtNatalie: wert('Natalie trägt'), gesamtbedarf: wert('Gesamtbedarf'),
    schenkungen: Math.abs(wert('abzüglich Schenkungen')), kaiDirekt: wert('Kai: schon direkt'),
    ueberschussKai: wert('Überschuss Kai'),
  } : null;

  // ---- Ablauf Gemeinschaftskonto: gebuchte Bewegungen + offene Planposten
  const ereignisse = [];
  for (const b of modell.buchungen) {
    if (!b.datum || b.datum.slice(0, 7) !== monat) continue;
    if (b.von !== gem && b.nach !== gem) continue;
    const k = modell.konten.find((x) => x.name === gem);
    if (k?.stichtag && b.datum <= k.stichtag) continue;
    ereignisse.push({
      art: 'gebucht', tag: tagVon(b.datum), datum: b.datum, titel: b.posten || b.beschreibung,
      betrag: b.nach === gem ? b.betrag : -b.betrag, nr: b.nr, giltFuer: b.giltFuer,
      gegen: b.nach === gem ? b.von : b.nach, kommentar: b.kommentar,
    });
  }
  const bucheGefunden = (posten, von, nach) => modell.buchungen.some((b) => b.giltFuer === monat
    && (von == null || b.von === von) && (nach == null || b.nach === nach) && norm(b.posten) === norm(posten));

  for (const k of modell.kosten) {
    if (k.konto !== gem || !aktivIn(k, monat) || !k.monatswert) continue;
    if (bucheGefunden(k.posten, gem, null)) continue;
    ereignisse.push({ art: 'geplant', tag: k.tag || null, titel: k.posten, betrag: -r2(k.monatswert), quelle: k.id, bemerkung: k.bemerkung });
  }
  const visaPlan = summe(modell.kosten.filter((k) => k.konto === visa && aktivIn(k, monat)), (k) => k.monatswert);
  const visaGebucht = gebuchtVonNach(gem, visa);
  if (visaPlan - visaGebucht > 0.005) {
    const tag = Math.min(...modell.kosten.filter((k) => k.konto === visa && aktivIn(k, monat) && k.tag).map((k) => k.tag), 99);
    ereignisse.push({ art: 'geplant', tag: tag === 99 ? null : tag, titel: 'Aufladung Visa', betrag: -r2(visaPlan - visaGebucht),
      bemerkung: 'Haushaltsgeld und alles, was über die Visa läuft' });
  }
  for (const e of modell.einnahmen) {
    if (e.konto !== gem || !aktivIn(e, monat)) continue;
    if (bucheGefunden(e.posten, null, gem)) continue;
    ereignisse.push({ art: 'geplant', tag: e.tag || null, titel: e.posten, betrag: r2(e.betrag), quelle: e.id, bemerkung: e.bemerkung });
  }
  for (const u of ueberweisungen) {
    if (u.nach !== gem || u.offen <= 0) continue;
    ereignisse.push({ art: 'geplant', tag: u.tag, titel: `Überweisung von ${u.person}`, betrag: u.offen,
      bemerkung: `Anteil für ${monatsname(monat)}`, person: u.person });
  }
  // Sortierung: nach Tag; am selben Tag erst Gebuchtes, dann Eingänge, dann Abgänge. Ohne Tag ans Ende.
  ereignisse.sort((x, y) => (x.tag ?? 99) - (y.tag ?? 99)
    || (x.art === y.art ? 0 : x.art === 'gebucht' ? -1 : 1)
    || (x.datum && y.datum ? x.datum.localeCompare(y.datum) : 0)
    || (y.betrag > 0) - (x.betrag > 0));
  // Zukünftiger Monat: Start ist die Prognose vom Monatsende davor.
  const zukunft = monat > heute.slice(0, 7);
  const startstand = zukunft && monat <= monatPlus(heute.slice(0, 7), 24)
    ? monatsansicht(modell, monatPlus(monat, -1), heute).ablauf.ende
    : kontostand(modell, gem, vormonatsEnde);
  let lauf = startstand;
  for (const e of ereignisse) { lauf = r2(lauf + e.betrag); e.danach = lauf; }
  const ablauf = { konto: gem, start: startstand, startDatum: vormonatsEnde, startIstPrognose: zukunft, ereignisse, ende: lauf };

  // ---- Buchungen des Monats (gilt für diesen Monat oder in diesem Monat gezahlt)
  const buchungen = modell.buchungen
    .filter((b) => b.giltFuer === monat || (b.datum && b.datum.slice(0, 7) === monat))
    .sort((x, y) => (y.datum || '').localeCompare(x.datum || '') || y.nr.localeCompare(x.nr));

  // ---- Abgleich mit dem, was Excel gerade ausgerechnet hat
  const abweichungen = [];
  if (modell.excel.monat === monat) {
    for (const [name, wertExcel] of Object.entries(modell.excel.kontostaende)) {
      const eigen = kontostand(modell, name, ende);
      if (wertExcel != null && Math.abs(eigen - wertExcel) > 0.005) abweichungen.push({ was: `Kontostand ${name}`, app: eigen, excel: r2(wertExcel) });
    }
  }
  for (const t of toepfeStand(modell, '9999-12-31')) {
    if (t.excelStand != null && Math.abs(t.stand - t.excelStand) > 0.005) abweichungen.push({ was: `Topf ${t.name}`, app: t.stand, excel: r2(t.excelStand) });
  }

  return {
    monat, name: monatsname(monat), laufend, vergangen, zukunft, heute, konten, toepfe, toepfeSumme, sparkonten, nichtInToepfen,
    ueberweisungen, anteil, ablauf, buchungen, abweichungen,
    offenePunkte: modell.offenePunkte.filter((p) => !p.erledigt),
  };
}

const MONATE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
export function monatsname(monat, mitJahr = false) {
  const [j, m] = monat.split('-').map(Number);
  return MONATE[m - 1] + (mitJahr ? ' ' + j : '');
}
