// Versteht einen gesprochenen oder getippten Satz und macht daraus Vorschläge fürs Buchungsformular.
// Feste Regeln, keine KI, läuft komplett auf dem Gerät. Ergebnis ist immer nur ein Vorschlag.

const MONATE = ['januar', 'februar', 'märz', 'april', 'mai', 'juni', 'juli', 'august', 'september', 'oktober', 'november', 'dezember'];
const MONAT_RE = '(januar|jänner|februar|märz|maerz|april|mai|juni|juli|august|september|oktober|november|dezember)';
const WOCHENTAGE = ['sonntag', 'montag', 'dienstag', 'mittwoch', 'donnerstag', 'freitag', 'samstag'];
const ZAHLWORTE = { null: 0, eins: 1, ein: 1, eine: 1, zwei: 2, drei: 3, vier: 4, fünf: 5, sechs: 6, sieben: 7, acht: 8, neun: 9, zehn: 10, elf: 11, zwölf: 12, zwanzig: 20, dreißig: 30, vierzig: 40, fünfzig: 50, hundert: 100 };

// Läden und Stichworte → Bereich (und typisches Konto). Ergänzt die bekannten Posten aus der Datei.
const STICHWORTE = [
  { re: /\b(lidl|netto|aldi|rewe|edeka|kaufland|penny|norma|dm|rossmann|müller|bäcker\w*|backshop|apotheke|einkauf\w*|lebensmittel|drogerie|supermarkt)\b/i, bereich: 'Wohnen & Haushalt', konto: 'visa' },
  { re: /\b(tank\w*|sprit|benzin|diesel|aral|shell|esso|jet|total|werkstatt|tüv|inspektion|reifen|waschanlage|parken|parkhaus)\b/i, bereich: 'Autos' },
  { re: /\b(kita|drk|turnen|windel\w*|baby\w*|kinder\w*|spielzeug|aaron|nora)\b/i, bereich: 'Kinder' },
  { re: /\b(versicherung\w*|huk|allianz|adac|württembergische|wuertt\w*)\b/i, bereich: 'Versicherungen' },
  { re: /\b(netflix|disney\w*|amazon music|spotify|prime|congstar|glasfaser|handy\w*|telefon|internet|abo)\b/i, bereich: 'Abos & Telefon' },
  { re: /\b(miete|strom|heizung|möbel|ikea|baumarkt|obi|hornbach|toom|haushalt\w*)\b/i, bereich: 'Wohnen & Haushalt' },
  { re: /\b(geschenk\w*|geburtstag|weihnacht\w*)\b/i, bereich: 'Sonstiges' },
  { re: /\b(gehalt|lohn)\b/i, bereich: 'Einkommen', art: 'Einnahme' },
  { re: /\b(kindergeld|elterngeld|kinderzuschlag)\b/i, bereich: 'Staatliche Leistungen', art: 'Einnahme' },
];

const norm = (s) => (s || '').toLowerCase().replace(/[()„“"',.!?]/g, ' ').replace(/\s+/g, ' ').trim();
const gross = (s) => s.charAt(0).toUpperCase() + s.slice(1);
const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

/** Betrag aus dem Text holen. Gibt {wert, treffer} oder null. */
export function betragFinden(text) {
  const t = text.replace(/(\d)\s*[.,]\s*(\d{2})(?!\d)/g, '$1,$2'); // „12 , 50“ → „12,50“
  const muster = [
    // 12 Euro 50 / 12 € 50 (Cent)
    { re: /(\d{1,3}(?:\.\d{3})*|\d+)\s*(?:euro|€|eur)\s+(?:und\s+)?(\d{1,2})(?:\s*(?:cent|ct))?(?!\s*[.,]?\d)(?!\s*(?:uhr|\.))/i, f: (m) => Number(m[1].replace(/\./g, '')) + Number(m[2]) / 100 },
    // 1.234,56 / 12,50 / 12,5
    { re: /(\d{1,3}(?:\.\d{3})+|\d+),(\d{1,2})\b/, f: (m) => Number(m[1].replace(/\./g, '')) + Number(m[2].padEnd(2, '0')) / 100 },
    // 12.50 (englische Schreibweise der Spracherkennung)
    { re: /\b(\d+)\.(\d{2})\b(?!\.\d)/, f: (m) => Number(m[1]) + Number(m[2]) / 100 },
    // 50 Cent
    { re: /\b(\d{1,2})\s*(?:cent|ct)\b/i, f: (m) => Number(m[1]) / 100 },
    // 12 Euro / 12 € / 12€
    { re: /(\d{1,3}(?:\.\d{3})+|\d+)\s*(?:euro|€|eur)\b/i, f: (m) => Number(m[1].replace(/\./g, '')) },
  ];
  for (const { re, f } of muster) {
    const m = t.match(re);
    if (m) return { wert: Math.round(f(m) * 100) / 100, treffer: m[0] };
  }
  // Zahlwort + Euro („zwölf Euro“)
  const zw = t.match(new RegExp(`\\b(${Object.keys(ZAHLWORTE).join('|')})\\s*(?:euro|€)`, 'i'));
  if (zw) return { wert: ZAHLWORTE[zw[1].toLowerCase()], treffer: zw[0] };
  // Letzte Möglichkeit: eine einzelne Zahl, die kein Datum ist
  const zahlen = [...t.matchAll(/(?<![.\d])(\d+(?:,\d{1,2})?)(?![.\d])(?!\s*\.?\s*(?:uhr|tag|januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember))/gi)]
    .filter((m) => !/\bam\s*$/i.test(t.slice(0, m.index)));
  if (zahlen.length === 1) return { wert: Number(zahlen[0][1].replace(',', '.')), treffer: zahlen[0][0] };
  return null;
}

/** Datum aus dem Text. */
export function datumFinden(text, heute) {
  const h = new Date(heute + 'T12:00:00');
  const t = text.toLowerCase();
  if (/\bvorgestern\b/.test(t)) { const d = new Date(h); d.setDate(d.getDate() - 2); return { wert: iso(d), treffer: 'vorgestern' }; }
  if (/\bgestern\b/.test(t)) { const d = new Date(h); d.setDate(d.getDate() - 1); return { wert: iso(d), treffer: 'gestern' }; }
  if (/\bheute\b/.test(t)) return { wert: heute, treffer: 'heute' };
  // am 3.11. / am 3.11.2026 / am 3. November
  let m = t.match(/\bam\s+(\d{1,2})\.\s*(\d{1,2})\.(?:\s*(\d{4}|\d{2}(?![,.\d])))?/);
  if (m) {
    const j = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3])) : h.getFullYear();
    let d = new Date(j, Number(m[2]) - 1, Number(m[1]), 12);
    if (!m[3] && d > h && (d - h) > 30 * 864e5) d = new Date(j - 1, Number(m[2]) - 1, Number(m[1]), 12);
    return { wert: iso(d), treffer: m[0] };
  }
  m = t.match(new RegExp(`\\bam\\s+(\\d{1,2})\\.?\\s+${MONAT_RE}`));
  if (m) {
    const mi = monatIndex(m[2]);
    let d = new Date(h.getFullYear(), mi, Number(m[1]), 12);
    if (d > h && (d - h) > 30 * 864e5) d = new Date(h.getFullYear() - 1, mi, Number(m[1]), 12);
    return { wert: iso(d), treffer: m[0] };
  }
  m = t.match(/\bam\s+(\d{1,2})\.?(?!\d)(?!\s*(?:euro|€))/);
  if (m) {
    let d = new Date(h.getFullYear(), h.getMonth(), Number(m[1]), 12);
    if (d > h) d = new Date(h.getFullYear(), h.getMonth() - 1, Number(m[1]), 12);
    return { wert: iso(d), treffer: m[0] };
  }
  // (am) Montag – der letzte vergangene
  for (let i = 0; i < 7; i++) {
    if (new RegExp(`\\b(?:am\\s+|letzten\\s+)?${WOCHENTAGE[i]}\\b`).test(t)) {
      const d = new Date(h);
      let diff = (h.getDay() - i + 7) % 7; if (diff === 0) diff = 7;
      d.setDate(d.getDate() - diff);
      return { wert: iso(d), treffer: t.match(new RegExp(`\\b(?:am\\s+|letzten\\s+)?${WOCHENTAGE[i]}\\b`))[0] };
    }
  }
  return null;
}
function monatIndex(name) { const n = name.replace('jänner', 'januar').replace('maerz', 'märz'); return MONATE.indexOf(n); }

/** „für Oktober“, „gilt für Dezember“, „für nächsten Monat“ → 'JJJJ-MM' */
export function giltFuerFinden(text, heute) {
  const t = text.toLowerCase();
  const [hj, hm] = heute.split('-').map(Number);
  const plus = (n) => { const d = new Date(hj, hm - 1 + n, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  if (/\bfür\s+(?:den\s+)?(?:nächsten|kommenden)\s+monat\b/.test(t)) return { wert: plus(1), treffer: t.match(/für\s+(?:den\s+)?(?:nächsten|kommenden)\s+monat/)[0] };
  if (/\bfür\s+(?:den\s+)?(?:letzten|vorigen|vergangenen)\s+monat\b/.test(t)) return { wert: plus(-1), treffer: t.match(/für\s+(?:den\s+)?(?:letzten|vorigen|vergangenen)\s+monat/)[0] };
  const m = t.match(new RegExp(`\\b(?:gilt\\s+)?für\\s+(?:den\\s+)?${MONAT_RE}(?:\\s+(\\d{4}))?`));
  if (!m) return null;
  const mi = monatIndex(m[1]);
  let j = m[2] ? Number(m[2]) : hj;
  if (!m[2]) { // nächstliegendes Jahr wählen
    const diff = (j * 12 + mi) - (hj * 12 + hm - 1);
    if (diff > 6) j -= 1; else if (diff < -6) j += 1;
  }
  return { wert: `${j}-${String(mi + 1).padStart(2, '0')}`, treffer: m[0] };
}

/**
 * Hauptfunktion.
 * @param {string} satz
 * @param {object} modell aus leseBuchhaltung()
 * @param {string} heute 'JJJJ-MM-TT'
 * @returns {{felder: object, erkannt: string[], fehlt: string[]}}
 */
export function verstehen(satz, modell, heute) {
  const roh = (satz || '').trim();
  let rest = ' ' + roh + ' ';
  const weg = (s) => { if (s) rest = rest.replace(s, ' '); };
  const felder = {};
  const erkannt = [];

  // Konten mit Spitznamen
  const k = (pred) => modell.konten.find(pred)?.name;
  const gem = k((x) => x.rolle === 'gemeinsam' && /giro/i.test(x.art || '')) || 'Gemeinschaftskonto';
  const visa = k((x) => x.rolle === 'gemeinsam' && /kredit|visa/i.test((x.art || '') + x.name)) || 'Visa Haushalt';
  const spar = k((x) => x.rolle === 'Sparen' && x.status !== 'geschlossen') || 'Tagesgeld Natalie';
  const extern = k((x) => x.rolle === 'extern') || 'Extern';
  const kaiPriv = k((x) => x.rolle === 'privat' && x.inhaber === 'Kai');
  const natPriv = k((x) => x.rolle === 'privat' && x.inhaber === 'Natalie');
  const KONTO_NAMEN = [
    [/gemeinschaftskont\w*|gemeinsame[ns]?\s+kont\w*|gemeinschaft|girokonto|giro(?!\s*natalie)/i, gem],
    [/visa(?:\s*haushalt)?|kreditkarte|haushaltskarte|(?:der|die)\s+karte/i, visa],
    [/tagesgeld\w*(?:\s*natalie)?|sparkont\w*|spartopf|töpfe/i, spar],
    ...(kaiPriv ? [[/c\s?24|kais?\s+kont\w*|mein(?:em)?\s+kont\w*\s+kai/i, kaiPriv]] : []),
    ...(natPriv ? [[/natalies?\s+(?:giro|kont\w*)|natalie\s+giro/i, natPriv]] : []),
  ];

  // Art
  let art = null;
  if (/\b(umbuch\w*|umgebucht|aufgeladen|aufladen|aufladung|rübergeschoben|überwiesen\s+auf)\b/i.test(roh)) art = 'Umbuchung';
  else if (/\b(einnahme|bekommen|eingegangen|gutschrift|erstattung|erstattet|zurückbekommen|eingang)\b/i.test(roh)) art = 'Einnahme';
  else if (/\b(ausgabe|bezahlt|gekauft|ausgegeben|abgebucht)\b/i.test(roh)) art = 'Ausgabe';
  ['umbuchung', 'einnahme', 'ausgabe'].forEach((w) => { rest = rest.replace(new RegExp(`\\b${w}\\b`, 'i'), ' '); });

  // Datum / Monat (vor dem Betrag, damit „am 28.10.“ nicht als Betrag gilt)
  const esc = (x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const gf = giltFuerFinden(roh, heute);
  if (gf) { felder.giltFuer = gf.wert; rest = rest.replace(new RegExp(esc(gf.treffer), 'i'), ' '); erkannt.push('Monat'); }
  const dt = datumFinden(rest, heute);
  if (dt) { felder.datum = dt.wert; rest = rest.replace(new RegExp(esc(dt.treffer), 'i'), ' '); erkannt.push('Datum'); }

  // Betrag
  const b = betragFinden(rest);
  if (b) { felder.betrag = b.wert.toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); weg(b.treffer); erkannt.push('Betrag'); }

  // Kommentar: alles nach „Kommentar“ / „Notiz“
  const km = rest.match(/\b(?:kommentar|notiz|anmerkung)\s*:?\s*(.+)$/i);
  if (km) { felder.kommentar = km[1].trim(); rest = rest.slice(0, km.index); erkannt.push('Kommentar'); }

  // Topf: „Topf Urlaub“ / „in den Topf Aaron“
  const topfNamen = ['Sparplan (verteilt)', ...modell.toepfe.map((t) => t.name)];
  const tm = rest.match(/\b(?:in\s+den\s+|aus\s+dem\s+|vom\s+|für\s+den\s+)?topf\s+(.+?)(?=\s+(?:von|vom|aus|auf|aufs|an|nach|mit|über|per|bereich)\b|\s*$)/i);
  if (tm) {
    const gesucht = norm(tm[1]);
    const t = topfNamen.find((n) => norm(n) === gesucht) || topfNamen.find((n) => norm(n).startsWith(gesucht) || gesucht.startsWith(norm(n)));
    if (t) { felder.topf = t; weg(tm[0]); erkannt.push('Topf'); }
  } else if (/\bsparplan\b/i.test(rest)) { felder.topf = 'Sparplan (verteilt)'; rest = rest.replace(/\bsparplan\w*\b/i, ' '); erkannt.push('Topf'); }

  // Konten mit Richtung
  const gefunden = [];
  for (const [re, name] of KONTO_NAMEN) {
    const m = rest.match(new RegExp(`(?:\\b(von|vom|aus|ab|mit\\s+der|mit\\s+dem|über\\s+die|über\\s+das|per|auf|aufs|an|ans|nach|zum|zur|in\\s+die|in\\s+das|ins)\\s+)?(?:das\\s+|die\\s+|dem\\s+|der\\s+)?(?:${re.source})`, 'i'));
    if (!m) continue;
    const praep = (m[1] || '').toLowerCase();
    const richtung = /^(von|vom|aus|ab|mit|über|per)/.test(praep) ? 'von' : /^(auf|aufs|an|ans|nach|zum|zur|in|ins)/.test(praep) ? 'nach' : null;
    gefunden.push({ name, richtung, pos: m.index });
    weg(m[0]);
  }
  gefunden.sort((a, b) => a.pos - b.pos);

  // Bekannte Posten aus der Datei (Kosten, Einnahmen, frühere Buchungen)
  const bekannte = [
    ...modell.buchungen.map((x) => ({ posten: x.posten, bereich: x.bereich, von: x.von, nach: x.nach, art: x.art, topf: x.topf })),
    ...modell.kosten.map((x) => ({ posten: x.posten, bereich: x.bereich, von: x.konto, nach: extern, art: 'Ausgabe', topf: x.konto === spar ? x.topf : null })),
    ...modell.einnahmen.map((x) => ({ posten: x.posten, bereich: x.bereich, von: extern, nach: x.konto, art: 'Einnahme' })),
  ].filter((x) => x.posten);
  const restNorm = ' ' + norm(rest) + ' ';
  let treffer = null;
  for (const x of bekannte) {
    const p = norm(x.posten);
    const kern = p.split(' ')[0];
    if (p.length >= 3 && restNorm.includes(' ' + p + ' ')) { if (!treffer || p.length > norm(treffer.posten).length) treffer = { ...x, voll: true }; }
    else if (!treffer && !p.includes(' ') && kern.length >= 4 && restNorm.includes(' ' + kern + ' ')) treffer = { ...x, voll: false, wort: kern };
  }
  const stich = STICHWORTE.find((s) => s.re.test(rest));

  // Art festlegen
  if (!art) art = treffer?.art || stich?.art || (gefunden.length >= 2 && gefunden.every((g) => g.name !== extern) ? 'Umbuchung' : 'Ausgabe');
  felder.art = art;

  // Von/Nach festlegen
  let von = gefunden.find((g) => g.richtung === 'von')?.name;
  let nach = gefunden.find((g) => g.richtung === 'nach')?.name;
  const ohne = gefunden.filter((g) => !g.richtung).map((g) => g.name);
  if (art === 'Ausgabe') { von = von || ohne.shift() || treffer?.von || (stich?.konto === 'visa' ? visa : null) || gem; nach = nach || extern; }
  else if (art === 'Einnahme') { nach = nach || ohne.shift() || treffer?.nach || gem; von = von || extern; }
  else {
    // „Visa aufgeladen“, „aufs Tagesgeld“: ein Konto ohne Richtung ist das Ziel, außer es ist das Gemeinschaftskonto
    if (!von && !nach && ohne.length === 1 && ohne[0] !== gem) nach = ohne.shift();
    von = von || ohne.shift() || gem; nach = nach || ohne.shift() || (von === gem ? visa : gem);
  }
  if (von === nach) nach = art === 'Ausgabe' ? extern : nach;
  felder.von = von; felder.nach = nach;
  if (gefunden.length) erkannt.push('Konten');

  // Bereich: ausdrücklich genannt > bekannter Posten > Stichwort
  const bereiche = modell.listen?.bereiche || [];
  const bm = rest.match(/\bbereich\s+(.+?)(?=\s*$|\s+(?:von|vom|auf|an)\b)/i);
  let bereich = null;
  const passend = (w) => bereiche.find((x) => norm(x).includes(norm(w)) || norm(w).includes(norm(x).split(' ')[0]));
  if (bm) { bereich = passend(bm[1]); if (bereich) weg(bm[0]); }
  if (!bereich) {
    for (const x of bereiche) { const erstes = norm(x).split(' ')[0]; if (erstes.length >= 4 && new RegExp(`\\b${erstes}\\b`, 'i').test(rest)) { bereich = x; rest = rest.replace(new RegExp(`\\b${erstes}\\b`, 'i'), ' '); break; } }
  }
  bereich = bereich || treffer?.bereich || stich?.bereich || (art === 'Umbuchung' ? 'Sparen & Umbuchungen' : null);
  if (bereich) { felder.bereich = bereich; erkannt.push('Bereich'); }
  if (!felder.topf && treffer?.topf && (von === spar || nach === spar)) felder.topf = treffer.topf;

  // Posten: bekannter Name oder der übrige Text
  const fuell = /\b(euro|€|eur|ich|wir|habe|hab|haben|hat|hast|bei|im|in|beim|für|fürs|den|die|der|das|dem|des|ein|eine|einen|und|mit|von|vom|aus|auf|aufs|an|nach|zum|zur|per|über|heute|gestern|bezahlt|gekauft|ausgegeben|abgebucht|bekommen|eingegangen|umgebucht|aufgeladen|konto|euros|getankt)\b/gi;
  let posten = treffer?.voll ? treffer.posten : rest.replace(fuell, ' ').replace(/[.,;:!?]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!treffer?.voll && treffer && posten.toLowerCase() === treffer.wort) posten = treffer.posten;
  if (!posten && art === 'Umbuchung') posten = nach === visa ? 'Aufladung Visa' : nach === spar ? (felder.topf && felder.topf !== 'Sparplan (verteilt)' ? `Sparen ${felder.topf}` : 'Sparraten') : `Umbuchung an ${nach}`;
  if (posten) { felder.posten = gross(posten); erkannt.push('Wofür'); }

  const fehlt = [];
  if (!felder.betrag) fehlt.push('Betrag');
  if (!felder.posten) fehlt.push('Wofür');
  if (!felder.bereich) fehlt.push('Bereich');
  return { felder, erkannt, fehlt };
}
