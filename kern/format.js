// Einheitliches Dateiformat für alle Bereiche (Gemeinsam, Tagesgeld, Natalie, Kai).
// Eine Quelle für: Vorlage (kern/vorlage.js), Lesen (kern/einheit.js), Umzug (kern/umzug.js).
// Spalten werden über ihre Überschrift gefunden. Die Reihenfolge hier ist die Reihenfolge in der Datei.
// Namen, die die Formeln der Zusatzblätter in Gemeinsam benutzen (Aufteilung, Überweisung),
// bleiben wie in der alten gemeinsamen Datei (z. B. „Bezahlt aus (Konto)“, „Sparrate / Monat“).
// Hier stehen keine persönlichen Daten.

export const FORMAT_VERSION = 1;
export const KOPFZEILE = 4; // Zeile 1 Titel, 2 Erklärung, 4 Überschriften, ab 5 Daten
export const ERSTE_ZEILE = 5;

// Typen: t = Text, d = Datum, m = Monat (Datum am 1.), z = Betrag, n = Zahl, j = Ja/Nein, f = Formel (nicht beschreiben)
// liste = Auswahlliste aus dem Blatt „Listen“
export const BLAETTER = {
  Buchungen: {
    tabelle: 'tblBuchungen', zeilen: 2000, titel: 'Buchungen',
    text: 'Jede Zahlung eine Zeile. Von → Nach: woher das Geld kommt und wohin es geht. Kategorie = wofür das Geld ausgegeben wurde.',
    spalten: [
      ['nr', 'Nr', 't', 8], ['datum', 'Datum (Zahlung)', 'd', 12], ['giltFuer', 'Gilt für Monat', 'm', 11],
      ['art', 'Art', 't', 14, 'Arten'], ['von', 'Von', 't', 18], ['nach', 'Nach', 't', 18], ['betrag', 'Betrag', 'z', 11],
      ['kategorie', 'Kategorie', 't', 16], ['bankKategorie', 'Bank-Kategorie', 't', 16], ['posten', 'Posten', 't', 22],
      ['topf', 'Topf', 't', 16], ['beschreibung', 'Beschreibung / Verwendungszweck', 't', 36], ['kommentar', 'Kommentar', 't', 24],
      ['quelle', 'Quelle', 't', 10, 'Quellen'], ['importId', 'Import-ID', 't', 14], ['gezahltVon', 'Gezahlt von', 't', 14, 'Gezahlt von'],
      ['gegenstueck', 'Gegenstück', 't', 14],
    ],
  },
  Konten: {
    tabelle: 'tblKonten', zeilen: 30, titel: 'Konten',
    text: 'Eigenes Konto = Ja: der Stand wird hier geführt. Nein: nur Gegenüber (z. B. das Konto der anderen Person).',
    spalten: [
      ['id', 'ID', 't', 7], ['name', 'Konto', 't', 22], ['bank', 'Bank', 't', 12], ['art', 'Art', 't', 12, 'Kontoarten'],
      ['inhaber', 'Inhaber', 't', 10], ['rolle', 'Rolle', 't', 11, 'Rollen'], ['eigen', 'Eigenes Konto', 'j', 9, 'JaNein'],
      ['frei', 'Zählt für Frei', 'j', 9, 'JaNein'], ['endung', 'Kontonummer-Endung', 't', 12], ['csv', 'CSV-Import', 'j', 9, 'JaNein'],
      ['anfang', 'Anfangsbestand', 'z', 13], ['stichtag', 'Stichtag', 'd', 12], ['stand', 'Stand', 'f', 13],
      ['status', 'Status', 't', 11, 'Kontostatus'], ['geschlossenAb', 'Geschlossen ab', 'd', 12], ['bemerkung', 'Bemerkung', 't', 30],
    ],
  },
  Kontostände: {
    tabelle: 'tblKontostaende', zeilen: 400, titel: 'Kontostände (Monatsabschluss)',
    text: 'Am Monatsende den Stand aus der Bank-App eintragen. Differenz ≠ 0 heißt: etwas fehlt oder ist doppelt.',
    spalten: [
      ['monat', 'Monat', 'm', 11], ['konto', 'Konto', 't', 22], ['bank', 'Ist-Stand laut Bank', 'z', 14], ['datei', 'Stand laut Datei', 'f', 14],
      ['differenz', 'Differenz', 'f', 11], ['kommentar', 'Kommentar', 't', 34], ['geprueft', 'Geprüft am', 'd', 12],
    ],
  },
  Kosten: {
    tabelle: 'tblKosten', zeilen: 120, titel: 'Kosten (Posten)',
    text: 'Geplante feste Ausgaben. Fällig (Monat) nur bei jährlich/halbjährlich/vierteljährlich: erster Monat im Jahr (1–12).',
    spalten: [
      ['id', 'ID', 't', 7], ['posten', 'Posten', 't', 26], ['kategorie', 'Kategorie', 't', 16], ['betrag', 'Betrag', 'z', 11],
      ['rhythmus', 'Rhythmus', 't', 13, 'Rhythmus'], ['faellig', 'Fällig (Monat)', 'n', 9], ['tag', 'Abbuchungstag', 'n', 9],
      ['konto', 'Bezahlt aus (Konto)', 't', 20], ['topf', 'Topf', 't', 16], ['vertragAuf', 'Vertrag auf', 't', 11],
      ['status', 'Status', 't', 11, 'Poststatus'], ['ab', 'Ab (Monat)', 'm', 11], ['bis', 'Bis (Monat)', 'm', 11],
      ['monatswert', 'Monatswert', 'f', 11], ['planAkt', 'Plan Monat', 'f', 11], ['istAkt', 'Gebucht Monat', 'f', 11], ['offenAkt', 'Offen Monat', 'f', 11],
      ['bemerkung', 'Bemerkung', 't', 30],
    ],
  },
  Einnahmen: {
    tabelle: 'tblEinnahmen', zeilen: 40, titel: 'Einnahmen (Posten)',
    text: 'Geplante feste Einnahmen. Tag = ab wann das Geld da ist.',
    spalten: [
      ['id', 'ID', 't', 7], ['posten', 'Posten', 't', 26], ['kategorie', 'Kategorie', 't', 16], ['betrag', 'Betrag', 'z', 11],
      ['rhythmus', 'Rhythmus', 't', 13, 'Rhythmus'], ['faellig', 'Fällig (Monat)', 'n', 9], ['tag', 'Tag (verfügbar ab)', 'n', 9],
      ['konto', 'Konto (Eingang)', 't', 20], ['gehoertZu', 'Gehört zu', 't', 11], ['giltFuer', 'Gilt für', 't', 14], ['eingang', 'Eingang (wann)', 't', 16],
      ['status', 'Status', 't', 11, 'Poststatus'], ['ab', 'Ab (Monat)', 'm', 11], ['bis', 'Bis (Monat)', 'm', 11],
      ['planAkt', 'Plan Monat', 'f', 11], ['istAkt', 'Gebucht Monat', 'f', 11], ['offenAkt', 'Offen Monat', 'f', 11], ['bemerkung', 'Bemerkung', 't', 30],
    ],
  },
  Töpfe: {
    tabelle: 'tblToepfe', zeilen: 60, titel: 'Töpfe',
    text: 'Gespartes Geld für einen Zweck. Automatisch zuteilen = Ja: die Sparrate wird jeden Monat dem Topf gutgeschrieben, ohne Überweisung.',
    spalten: [
      ['name', 'Topf', 't', 22], ['konto', 'Konto', 't', 18], ['kategorie', 'Kategorie', 't', 16], ['sparrate', 'Sparrate / Monat', 'z', 11],
      ['automatisch', 'Automatisch zuteilen', 'j', 10, 'JaNein'], ['ab', 'Ab (Monat)', 'm', 11], ['bis', 'Bis (Monat)', 'm', 11],
      ['bedarf', 'Bedarf / Monat', 'z', 11], ['anfang', 'Anfangsbestand', 'z', 12], ['stichtag', 'Stichtag', 'd', 12],
      ['frei', 'Zählt als frei', 'j', 9, 'JaNein'], ['ziel', 'Ziel', 'z', 11], ['zielBis', 'Ziel bis', 'm', 11],
      ['stand', 'Stand aktuell', 'f', 12], ['bemerkung', 'Bemerkung', 't', 30],
    ],
  },
  Limits: {
    tabelle: 'tblLimits', zeilen: 40, titel: 'Limits',
    text: 'Ein Limit warnt nur. Es wird nicht vorab vom Freien abgezogen, auch nicht in der Vorschau.',
    spalten: [
      ['kategorie', 'Kategorie', 't', 18], ['betrag', 'Betrag / Monat', 'z', 12], ['ab', 'Ab (Monat)', 'm', 11], ['bis', 'Bis (Monat)', 'm', 11],
      ['istAkt', 'Ausgegeben Monat', 'f', 12], ['restAkt', 'Rest Monat', 'f', 11], ['bemerkung', 'Bemerkung', 't', 30],
    ],
  },
  Kategorien: {
    tabelle: 'tblKategorien', zeilen: 80, titel: 'Kategorien',
    text: 'Wofür wurde das Geld ausgegeben? Eine gemeinsame Liste für alle Bereiche. Neue Kategorien erscheinen überall.',
    spalten: [['name', 'Kategorie', 't', 20], ['erklaerung', 'Erklärung', 't', 40], ['aktiv', 'Aktiv', 'j', 8, 'JaNein'], ['angelegt', 'Angelegt von', 't', 12]],
  },
  Regeln: {
    tabelle: 'tblRegeln', zeilen: 200, titel: 'Regeln',
    text: 'Beim Import: Enthält der Bank-Text das Stichwort, schlägt die App Kategorie/Topf/Posten vor. Gespeichert wird erst nach Bestätigung.',
    spalten: [['stichwort', 'Bank-Text enthält', 't', 24], ['kategorie', 'Kategorie', 't', 16], ['topf', 'Topf', 't', 16], ['posten', 'Posten', 't', 20], ['konto', 'Nur für Konto', 't', 16], ['angelegt', 'Angelegt am', 'd', 12]],
  },
  'Soll-Überweisungen': {
    tabelle: 'tblSoll', zeilen: 120, titel: 'Soll-Überweisungen',
    text: 'Was pro Monat zwischen den Bereichen überwiesen werden soll (z. B. Anteil aufs Gemeinschaftskonto).',
    spalten: [['monat', 'Monat', 'm', 11], ['von', 'Von', 't', 18], ['nach', 'Nach', 't', 18], ['soll', 'Soll', 'z', 11], ['ist', 'Gebucht', 'f', 11], ['offen', 'Offen', 'f', 11],
      ['wirkung', 'Wirkung auf Frei', 'f', 9], ['bemerkung', 'Bemerkung', 't', 30]],
  },
  Einstellungen: {
    tabelle: 'tblEinstellungen', zeilen: 40, titel: 'Einstellungen',
    text: 'Wird von der App gepflegt (Funktionen an/aus, Reihenfolge der Widgets). Hier nichts von Hand ändern.',
    spalten: [['schluessel', 'Schlüssel', 't', 22], ['wert', 'Wert', 't', 60], ['erklaerung', 'Erklärung', 't', 50]],
  },
};

// Zusatzblätter mit fester Tabelle (werden neu angelegt; Aufteilung, Überweisung, Kredite, Jahr, Depot werden aus der alten Datei kopiert)
export const ZUSATZ_TABELLEN = {
  'Offene Punkte': {
    tabelle: 'tblOffen', zeilen: 200, titel: 'Offene Punkte', text: 'Was noch zu klären oder zu erledigen ist.',
    spalten: [['nr', 'Nr', 'n', 6], ['thema', 'Thema', 't', 18], ['text', 'Was ist offen', 't', 60], ['status', 'Status', 't', 11], ['notiz', 'Notiz', 't', 40]],
  },
  'TR Buchungen': {
    tabelle: 'tblTR', zeilen: 2000, titel: 'Trade Republic – Buchungen', text: 'Aus dem Trade-Republic-Export. Wird vom Blatt „Depot“ ausgewertet.',
    spalten: [['datum', 'Datum', 'd', 12], ['typ', 'Typ', 't', 12], ['name', 'Name', 't', 28], ['isin', 'ISIN', 't', 14], ['position', 'Position', 't', 16],
      ['stueck', 'Stück', 'n', 10], ['kurs', 'Kurs', 'n', 10], ['betrag', 'Betrag', 'z', 11], ['gebuehr', 'Gebühr', 'z', 9], ['steuer', 'Steuer', 'z', 9],
      ['beschreibung', 'Beschreibung', 't', 30], ['id', 'ID', 't', 20]],
  },
  Erinnerungen: {
    tabelle: 'tblErinnerungen', zeilen: 60, titel: 'Erinnerungen', text: 'Monatliche Erinnerungen. Erledigt = Monate, in denen sie abgehakt wurde.',
    spalten: [['id', 'ID', 't', 8], ['text', 'Text', 't', 40], ['tag', 'Tag', 'n', 6], ['aktiv', 'Aktiv', 'j', 8], ['erledigt', 'Erledigt (Monate)', 't', 40]],
  },
  Challenges: {
    tabelle: 'tblChallenges', zeilen: 60, titel: 'Challenges', text: 'Spar-Challenges. Daten = Einstellungen der Challenge (von der App gepflegt).',
    spalten: [['id', 'ID', 't', 8], ['art', 'Art', 't', 14], ['titel', 'Titel', 't', 26], ['start', 'Start', 'd', 12], ['ende', 'Ende', 'd', 12], ['status', 'Status', 't', 11], ['daten', 'Daten', 't', 60]],
  },
  Ruhmeshalle: {
    tabelle: 'tblRuhm', zeilen: 200, titel: 'Ruhmeshalle', text: 'Geschaffte Challenges und Auszeichnungen.',
    spalten: [['datum', 'Datum', 'd', 12], ['auszeichnung', 'Auszeichnung', 't', 22], ['challenge', 'Challenge', 't', 26], ['text', 'Text', 't', 50]],
  },
};
Object.assign(BLAETTER, ZUSATZ_TABELLEN);

// Grundblätter in jeder Datei, in dieser Reihenfolge. „Übersicht“ und „Vorschau“ erzeugt kern/vorlage.js (nur Formeln),
// „Erklärung“ und „Listen“ ebenfalls.
export const GRUNDBLAETTER = ['Übersicht', 'Vorschau', 'Buchungen', 'Konten', 'Kontostände', 'Kosten', 'Einnahmen', 'Töpfe',
  'Limits', 'Soll-Überweisungen', 'Kategorien', 'Regeln', 'Einstellungen', 'Erklärung', 'Listen'];

// Zusatzblätter nur dort, wo nötig (werden beim Umzug aus der alten Datei übernommen oder leer angelegt).
export const ZUSATZBLAETTER = {
  gemeinsam: ['Aufteilung', 'Überweisung', 'Kredite', 'Jahr', 'Offene Punkte'],
  tagesgeld: [],
  kai: ['Depot', 'TR Buchungen', 'Offene Punkte'],
  natalie: ['Challenges', 'Ruhmeshalle', 'Erinnerungen'],
};
export const BEREICH_DATEI = {
  gemeinsam: { datei: 'Gemeinsam.xlsx', ordner: 'Finanz-Tool_König-Rex', farbe: 'gruen' },
  tagesgeld: { datei: 'Tagesgeld.xlsx', ordner: 'Finanz-Tool_König-Rex', farbe: 'gruen' },
  kai: { datei: 'Kai.xlsx', ordner: 'Apps/Kai', farbe: 'blau' },
  natalie: { datei: 'Natalie.xlsx', ordner: 'Apps/Natalie', farbe: 'rosa' },
};

export const LISTEN = {
  Arten: ['Einnahme', 'Ausgabe', 'Umbuchung', 'In Topf legen', 'Aus Topf nehmen'],
  Rhythmus: ['monatlich', 'vierteljährlich', 'halbjährlich', 'jährlich'],
  JaNein: ['Ja', 'Nein'],
  Quellen: ['manuell', 'CSV', 'Aufteilung', 'Umzug'],
  'Gezahlt von': ['Kai', 'Natalie', 'Ausgleich: aus Topf', 'Ausgleich: falsch bezahlt'],
  Rollen: ['gemeinsam', 'privat', 'Sparen', 'extern'],
  Kontoarten: ['Giro', 'Haushalt', 'Kreditkarte', 'Pocket', 'Tagesgeld', 'PayPal', 'Depot', 'Extern'],
  Kontostatus: ['aktiv', 'geschlossen'],
  Poststatus: ['aktiv', 'geschätzt', 'beendet'],
};
export const RHYTHMUS_ABSTAND = { monatlich: 1, vierteljährlich: 3, halbjährlich: 6, jährlich: 12 };

// Arten: Topf-Wirkung hängt nur an der Art (eigene Konten der Datei bestimmen die Art).
// Einnahme / In Topf legen: +; Ausgabe / Aus Topf nehmen: −; Umbuchung (zwischen eigenen Konten): 0.
export const TOPF_PLUS = ['Einnahme', 'In Topf legen'];
export const TOPF_MINUS = ['Ausgabe', 'Aus Topf nehmen'];
export const NUR_TOPF = ['In Topf legen', 'Aus Topf nehmen']; // bewegen kein Geld zwischen Konten
export const SPARPLAN = 'Sparplan (verteilt)';
// Konten, deren Kosten wie ein gemeinsamer Topf zählen (Haushaltsgeld über die Karte): offen = Plan der Karte − mit der Karte ausgegeben
export const KARTEN_ARTEN = ['Kreditkarte', 'Haushalt']; // Topf-Wert: wird nach Sparrate auf alle Töpfe verteilt

// Gemeinsame Kategorien-Liste (Kais bisherige Liste + Drogerie; Versicherungen, Kredite aus Gemeinsam, Entscheidung 16.11.2026)
export const KATEGORIEN_START = ['Freizeit', 'Essen & Trinken', 'Lieferdienst', 'Shopping', 'Amazon', 'Haushalt', 'Drogerie',
  'Handwerksbedarf', 'Auto', 'Kinder', 'Abos', 'Geschenke', 'Gesundheit', 'Bargeld', 'Versicherungen', 'Kredite', 'Sonstiges'];

// Alte „Bereich“-Werte der gemeinsamen Datei → Kategorie (Entscheidung Kai 16.11.2026). Nicht genannte → leer.
export const BEREICH_ZU_KATEGORIE = {
  'Wohnen & Haushalt': 'Haushalt', Autos: 'Auto', Kinder: 'Kinder', 'Abos & Telefon': 'Abos', Sonstiges: 'Sonstiges',
  Versicherungen: 'Versicherungen', Kredite: 'Kredite',
};

// ---------------------------------------------------------------- Funktionen und Widgets
// art: widget = Karte in der Übersicht, seite = Kachel im Widget „Extras“ (eigene Seite)
export const FUNKTIONEN = {
  'wer-ueberweist': { titel: 'Wer überweist was', art: 'widget', bereiche: ['gemeinsam'] },
  'sprit-limit': { titel: 'Sprit-Limit', art: 'widget' },
  'was-kommt': { titel: 'Was noch kommt', art: 'widget' },
  limits: { titel: 'Limits', art: 'widget' },
  erinnerungen: { titel: 'Erinnerungen', art: 'widget' },
  'warnung-vorschau': { titel: 'Warnung Vorschau', art: 'widget' },
  kategorien: { titel: 'Kategorien', art: 'widget' },
  tagesgeld: { titel: 'Tagesgeld und Töpfe', art: 'widget', bereiche: ['natalie'] },
  import: { titel: 'Import (Kontoauszug)', art: 'seite' },
  'kontostand-pruefen': { titel: 'Kontostand prüfen', art: 'seite' },
  ausgaben: { titel: 'Ausgaben', art: 'seite' },
  depot: { titel: 'Depot', art: 'seite' },
  challenges: { titel: 'Challenges', art: 'seite' },
  ruhmeshalle: { titel: 'Ruhmeshalle', art: 'seite' },
  aufteilung: { titel: 'Aufteilung', art: 'seite', bereiche: ['gemeinsam'] },
  'offene-punkte': { titel: 'Offene Punkte', art: 'seite' },
  'soll-ueberweisung': { titel: 'Überweisung an Gemeinsam', art: 'seite', bereiche: ['kai', 'natalie'] },
};
// Standard-Reihenfolge der Widgets (Konzept, Runde 2). „Extras“ steht immer ganz unten und ist nicht Teil der Liste.
export const WIDGETS_START = {
  gemeinsam: ['wer-ueberweist', 'sprit-limit', 'was-kommt'],
  natalie: ['was-kommt', 'limits', 'erinnerungen'],
  kai: ['warnung-vorschau', 'kategorien', 'limits'],
};
// „Wie vor dem Umbau“: was am 16.11.2026 (App 0.17.0) im jeweiligen Bereich sichtbar war.
export const VORHER_AKTIV = {
  gemeinsam: ['wer-ueberweist', 'sprit-limit', 'was-kommt', 'offene-punkte', 'import', 'kontostand-pruefen', 'aufteilung'],
  natalie: ['was-kommt', 'limits', 'erinnerungen', 'tagesgeld', 'import', 'challenges', 'ruhmeshalle', 'soll-ueberweisung'],
  kai: ['warnung-vorschau', 'kategorien', 'limits', 'import', 'kontostand-pruefen', 'ausgaben', 'depot', 'offene-punkte', 'soll-ueberweisung'],
};

/** Spalten eines Blatts als Liste {key, name, typ, breite, liste, buchst}. */
export function spalten(blatt) {
  const b = BLAETTER[blatt];
  if (!b) throw new Error(`Unbekanntes Blatt ${blatt}`);
  return b.spalten.map(([key, name, typ, breite, liste], i) => ({ key, name, typ, breite, liste, buchst: buchstabe(i) }));
}
export function buchstabe(i) { let s = ''; i += 1; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; }
/** Buchstabe der Spalte mit diesem Schlüssel. */
export const sp = (blatt, key) => { const c = spalten(blatt).find((x) => x.key === key); if (!c) throw new Error(`${blatt}.${key}?`); return c.buchst; };
