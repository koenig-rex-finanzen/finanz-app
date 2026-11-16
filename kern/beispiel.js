// Erfundene Beispieldaten im neuen Format (keine echten Namen, Konten oder Beträge).
const m = (heute, n) => { const d = new Date(Date.UTC(+heute.slice(0, 4), +heute.slice(5, 7) - 1 + n, 1)); return d.toISOString().slice(0, 7); };
const tag = (monat, t) => `${monat}-${String(t).padStart(2, '0')}`;

/** Privater Bereich (wie Kai/Natalie): Giro + Spar-Pocket, automatische Töpfe, Limits, Soll-Überweisung. */
export function beispielPrivat(heute) {
  const akt = heute.slice(0, 7); const vor = m(heute, -1); const vor2 = m(heute, -2);
  const t = Math.min(Number(heute.slice(8, 10)), 28);
  const B = []; let n = 1;
  const b = (datum, art, von, nach, betrag, x = {}) => B.push({ nr: 'B' + String(n++).padStart(4, '0'), datum, giltFuer: x.giltFuer || datum.slice(0, 7), art, von, nach, betrag, quelle: 'manuell', ...x });
  b(tag(vor2, 28), 'Einnahme', 'Arbeitgeber', 'Giro A', 2500, { posten: 'Gehalt', giltFuer: vor });
  b(tag(vor, 1), 'Ausgabe', 'Giro A', 'Extern', 400, { posten: 'Miete-Anteil' });
  b(tag(vor, 3), 'Ausgabe', 'Giro A', 'Extern', 54.9, { kategorie: 'Essen & Trinken', beschreibung: 'Restaurant' });
  b(tag(vor, 4), 'Umbuchung', 'Giro A', 'Spar-Pocket', 300);
  b(tag(vor, 10), 'Ausgabe', 'Giro A', 'Extern', 89.99, { kategorie: 'Shopping', topf: 'Kleidung' });
  b(tag(vor, 12), 'In Topf legen', 'Giro A', 'Giro A', 50, { topf: 'Urlaub' });
  b(tag(vor, 20), 'Ausgabe', 'Giro A', 'Gemeinschaft', 700, { posten: 'Überweisung Gemeinsam' });
  b(tag(vor, 28), 'Einnahme', 'Arbeitgeber', 'Giro A', 2500, { posten: 'Gehalt', giltFuer: akt });
  b(tag(akt, 1), 'Ausgabe', 'Giro A', 'Extern', 400, { posten: 'Miete-Anteil' });
  b(tag(akt, 2), 'Ausgabe', 'Giro A', 'Extern', 30, { kategorie: 'Dampfen' });
  b(tag(akt, Math.max(1, t - 1)), 'Ausgabe', 'Giro A', 'Extern', 45.5, { kategorie: 'Dampfen' });
  b(tag(akt, Math.max(1, t - 1)), 'Aus Topf nehmen', 'Giro A', 'Giro A', 20, { topf: 'Urlaub' });
  b(tag(akt, 1), 'Ausgabe', 'Giro A', 'Gemeinschaft', 500, { posten: 'Überweisung Gemeinsam' });
  b(tag(m(heute, 1), 2), 'Ausgabe', 'Giro A', 'Extern', 999, { kategorie: 'Shopping', kommentar: 'Zukunft: zählt noch nicht' });
  // Karte als Topf: Haushaltsgeld + Streaming über die Karte, Einkäufe ohne Posten
  b(tag(akt, 1), 'Umbuchung', 'Giro A', 'Karte A', 310);
  b(tag(akt, 3), 'Ausgabe', 'Karte A', 'Supermarkt', 95.4, { kategorie: 'Essen & Trinken' });
  b(tag(akt, Math.max(1, t - 1)), 'Ausgabe', 'Karte A', 'Drogerie', 24.6, { kategorie: 'Drogerie' });
  // Posten mit gleichnamiger Kategorie (Sprit): Tanken ohne Posten zählt für den Posten
  b(tag(akt, Math.max(1, t - 1)), 'Ausgabe', 'Giro A', 'Tankstelle', 40, { kategorie: 'Sprit' });
  b(tag(akt, Math.max(1, t - 1)), 'Ausgabe', 'Giro A', 'Tankstelle', 15, { kategorie: 'Sprit', posten: 'Sprit' });
  return {
    bereich: 'kai', titel: 'Beispiel privat',
    konten: [
      { id: 'KT1', name: 'Giro A', bank: 'Bank X', art: 'Giro', inhaber: 'A', rolle: 'privat', eigen: 'Ja', frei: 'Ja', endung: '0001', csv: 'Ja', anfang: 1000, stichtag: tag(vor2, 27), status: 'aktiv' },
      { id: 'KT2', name: 'Spar-Pocket', bank: 'Bank X', art: 'Pocket', inhaber: 'A', rolle: 'Sparen', eigen: 'Ja', frei: 'Ja', anfang: 200, stichtag: tag(vor2, 27), status: 'aktiv' },
      { id: 'KT3', name: 'Gemeinschaft', art: 'Giro', rolle: 'gemeinsam', eigen: 'Nein', frei: 'Nein', status: 'aktiv' },
      { id: 'KT4', name: 'Karte A', art: 'Kreditkarte', rolle: 'privat', eigen: 'Ja', frei: 'Ja', anfang: 0, stichtag: tag(vor2, 27), status: 'aktiv' },
    ],
    buchungen: B,
    kosten: [
      { id: 'K01', posten: 'Miete-Anteil', kategorie: 'Haushalt', betrag: 400, rhythmus: 'monatlich', tag: 1, konto: 'Giro A', status: 'aktiv' },
      { id: 'K02', posten: 'Handy', kategorie: 'Abos', betrag: 15, rhythmus: 'monatlich', tag: 15, konto: 'Giro A', status: 'aktiv' },
      { id: 'K03', posten: 'Versicherung', kategorie: 'Versicherungen', betrag: 120, rhythmus: 'jährlich', faellig: +m(heute, 2).slice(5, 7), konto: 'Giro A', status: 'aktiv' },
      { id: 'K04', posten: 'Abo quartal', kategorie: 'Abos', betrag: 30, rhythmus: 'vierteljährlich', faellig: +akt.slice(5, 7), konto: 'Giro A', status: 'aktiv' },
      { id: 'K05', posten: 'Altes Abo', betrag: 9.99, rhythmus: 'monatlich', konto: 'Giro A', status: 'aktiv', bis: vor },
      { id: 'K06', posten: 'Fremdes Konto', betrag: 50, rhythmus: 'monatlich', konto: 'Gemeinschaft', status: 'aktiv' },
      { id: 'K07', posten: 'Haushaltsgeld', betrag: 300, rhythmus: 'monatlich', konto: 'Karte A', status: 'aktiv' },
      { id: 'K08', posten: 'Streaming', betrag: 10, rhythmus: 'monatlich', konto: 'Karte A', status: 'aktiv' },
      { id: 'K09', posten: 'Sprit', betrag: 100, rhythmus: 'monatlich', konto: 'Giro A', status: 'aktiv' },
    ],
    einnahmen: [{ id: 'E01', posten: 'Gehalt', betrag: 2500, rhythmus: 'monatlich', tag: 28, konto: 'Giro A', status: 'aktiv' }],
    toepfe: [
      { name: 'Urlaub', sparrate: 50, automatisch: 'Ja', ab: vor2, anfang: 100, stichtag: tag(vor2, 27), frei: 'Nein' },
      { name: 'Kleidung', sparrate: 30, automatisch: 'Ja', ab: vor, anfang: 0, stichtag: tag(vor2, 27), frei: 'Nein' },
      { name: 'Rücklage', sparrate: 0, automatisch: 'Nein', anfang: 75, stichtag: tag(vor2, 27), frei: 'Ja' },
    ],
    limits: [{ kategorie: 'Dampfen', betrag: 90 }, { kategorie: 'Shopping', betrag: 150 }],
    soll: [{ monat: akt, von: 'Giro A', nach: 'Gemeinschaft', soll: 700 }, { monat: m(heute, 1), von: 'Giro A', nach: 'Gemeinschaft', soll: 650 }],
    kontostaende: [{ monat: vor, konto: 'Giro A', bank: 2105.11 }, { monat: vor, konto: 'Spar-Pocket', bank: 500 }],
    regeln: [{ stichwort: 'TANKSTELLE', kategorie: 'Auto' }],
  };
}

/** Tagesgeld: ein Sparkonto, Töpfe mit Sparplan (verteilt), Einzahlungen von außen. */
export function beispielTagesgeld(heute) {
  const akt = heute.slice(0, 7); const vor = m(heute, -1);
  const B = []; let n = 1;
  const b = (datum, art, von, nach, betrag, x = {}) => B.push({ nr: 'T' + String(n++).padStart(4, '0'), datum, giltFuer: datum.slice(0, 7), art, von, nach, betrag, quelle: 'manuell', ...x });
  b(tag(vor, 1), 'Einnahme', 'Person A Giro', 'Tagesgeld', 300, { topf: 'Sparplan (verteilt)', gezahltVon: 'A' });
  b(tag(vor, 15), 'Einnahme', 'Person B Giro', 'Tagesgeld', 200, { topf: 'Sparplan (verteilt)', gezahltVon: 'B' });
  b(tag(vor, 20), 'Ausgabe', 'Tagesgeld', 'Person B Giro', 53.3, { topf: 'Kinder' });
  b(tag(akt, 1), 'Einnahme', 'Gemeinschaft', 'Tagesgeld', 120, { topf: 'Auto' });
  b(tag(akt, 1), 'Umbuchung', 'Tagesgeld alt', 'Tagesgeld', 40, { topf: 'Auto' });
  return {
    bereich: 'tagesgeld', titel: 'Beispiel Tagesgeld',
    konten: [
      { id: 'KT1', name: 'Tagesgeld', art: 'Tagesgeld', rolle: 'Sparen', eigen: 'Ja', frei: 'Ja', anfang: 1500, stichtag: tag(m(heute, -2), 28), status: 'aktiv' },
      { id: 'KT2', name: 'Tagesgeld alt', art: 'Tagesgeld', rolle: 'Sparen', eigen: 'Ja', frei: 'Ja', anfang: 40, stichtag: tag(m(heute, -2), 28), status: 'aktiv' },
    ],
    buchungen: B,
    kosten: [], einnahmen: [],
    toepfe: [
      { name: 'Auto', konto: 'Tagesgeld', sparrate: 100, anfang: 600, frei: 'Nein' },
      { name: 'Kinder', konto: 'Tagesgeld', sparrate: 150, anfang: 400, frei: 'Nein' },
      { name: 'Notfall', konto: 'Tagesgeld', sparrate: 250, anfang: 300, frei: 'Nein', ziel: 3000 },
    ],
    limits: [], soll: [], kontostaende: [{ monat: vor, konto: 'Tagesgeld', bank: 1946.7 }], regeln: [],
  };
}
