// PC-Ansicht „Unser Geld“ (…/finanz-app/#pc): Kontoauszüge abgleichen und importieren,
// Monatsabschluss, Buchungsliste mit Bearbeiten, Zuordnungsregeln.
import { leseBuchhaltung } from './kern/lesen.js';
import { eur, monatsname, monatsansicht, monate, monatsende, monatPlus, AUSGLEICH, istAusgleich } from './kern/rechnen.js';
import { analysieren, csvText, kurzname } from './kern/csv.js';
import * as ms from './anmeldung.js';
import * as od from './onedrive.js';

const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const datum = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
const zeit = (t) => new Date(t).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' });
const heute = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
const zahl = (t) => { const s = String(t ?? '').trim().replace(/\s|€/g, ''); return s ? Math.round(Number(s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s) * 100) / 100 : NaN; };
const STATUS = {
  neu: ['Neu', 'Steht auf dem Auszug, aber noch nicht in der Datei. Wird als neue Buchung eingetragen.'],
  vielleicht: ['Bitte entscheiden', 'Es gibt eine ähnliche Buchung in der Datei. Bitte entscheiden, ob es dieselbe ist. Ohne Entscheidung wird nichts übernommen.'],
  eingetragen: ['Schon eingetragen', 'Habt ihr schon eingetragen (auch zusammengerechnet). Es wird nur die Import-ID ergänzt, damit es nie doppelt wird.'],
  'vor-stichtag': ['Vor dem Stichtag', 'Steckt schon im Anfangsbestand – normalerweise nicht importieren.'],
  gegenbuchung: ['Gegenbuchung', 'Die andere Seite einer Umbuchung – wird nicht doppelt gebucht.'],
  importiert: ['Schon importiert', 'Diese Bankzeile ist schon in der Datei.'],
};

const z = {
  modell: null, item: null, ort: null, regeln: [], ansicht: 'import', meldung: null,
  dateien: [], analyse: null, filter: 'neu', schreibt: false, bestaetigen: null,
  abschlussMonat: null, suche: '', bMonat: '', bKonto: '', bArt: '', bearbeiten: null, loeschenFrage: false,
};
let root;

export async function starte(el) {
  root = el;
  document.body.classList.add('pc-modus');
  if (!ms.aktiv() || !ms.angemeldet()) return zeigeAnmelden();
  await laden();
}

function zeigeAnmelden(fehler) {
  root.innerHTML = `<main class="pc-start">
    <h1>Unser Geld am PC</h1>
    <p>Hier gleicht ihr die Kontoauszüge mit euren Einträgen ab, importiert, was fehlt, und schließt den Monat ab. Dafür ist die Anmeldung mit dem Microsoft-Konto nötig, mit dem ihr die Datei in OneDrive öffnet.</p>
    <div><button class="knopf" id="pc-anmelden">Mit Microsoft anmelden</button></div>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <p class="hint"><a href="./">Zur Handy-Ansicht</a></p>
  </main>`;
}

async function laden(versuch = 1) {
  if (!z.modell) root.innerHTML = '<main class="pc-start"><p class="hint">Lade die Datei aus OneDrive …</p></main>';
  try {
    const g = await od.dateiFinden();
    if (!g) return zeigeAnmelden('Die Datei wurde in OneDrive nicht gefunden.');
    const bytes = await od.herunterladen(g.ort);
    z.modell = await leseBuchhaltung(bytes); z.item = g.item; z.ort = g.ort;
    z.regeln = await od.regelnLaden(g.item).catch(() => z.regeln || []);
    if (z.dateien.length) neuAnalysieren(); else z.analyse = null;
    if (!z.abschlussMonat) { const t = heute(); z.abschlussMonat = Number(t.slice(8, 10)) <= 15 ? monatPlus(t.slice(0, 7), -1) : t.slice(0, 7); }
    zeichne();
  } catch (e) {
    if (e instanceof od.AnmeldungNoetig) return zeigeAnmelden('Die Anmeldung ist abgelaufen.');
    if (versuch < 3) { await new Promise((r) => setTimeout(r, 1500 * versuch)); return laden(versuch + 1); }
    zeigeAnmelden('Fehler beim Laden: ' + e.message + '. Bitte die Seite neu laden.');
  }
}

function neuAnalysieren() {
  // Entscheidungen und Änderungen aus der alten Analyse übernehmen (gleiche Bankzeile = gleiche ID)
  const alt = new Map((z.analyse?.zeilen || []).map((x) => [x.id, x]));
  z.analyse = analysieren(z.dateien, z.modell, z.regeln);
  for (const x of z.analyse.zeilen) {
    const a = alt.get(x.id);
    if (!a || a.status !== x.status) continue;
    x.auswahl = a.auswahl; x.entscheidung = a.entscheidung; if (a.geaendert) { x.buchung = a.buchung; x.geaendert = true; }
  }
}

// ---------------------------------------------------------------- Rahmen
const NAV = [['import', 'Kontoauszüge abgleichen'], ['abschluss', 'Monatsabschluss'], ['buchungen', 'Buchungen'], ['regeln', 'Regeln']];
function zeichne() {
  const konto = ms.gemerktesKonto();
  const scroll = window.scrollY;
  root.innerHTML = `<div class="pc">
    <aside class="pc-nav">
      <div class="pc-logo">Unser Geld</div>
      ${NAV.map(([id, name]) => `<button class="${z.ansicht === id ? 'an' : ''}" data-ansicht="${id}">${name}</button>`).join('')}
      <div class="pc-fuss">
        <div class="hint">${h(z.item?.name)}</div>
        <div class="hint">geändert ${z.item?.lastModifiedDateTime ? zeit(z.item.lastModifiedDateTime) : '–'}${z.item?.lastModifiedBy?.user?.displayName ? ' von ' + h(z.item.lastModifiedBy.user.displayName) : ''}</div>
        <div class="hint">${z.modell.buchungen.length} Buchungen, angemeldet als ${h(konto || '–')}</div>
        <button class="knopf leise klein" id="pc-neu-laden">Neu laden</button>
        <a class="hint" href="./">Zur Handy-Ansicht</a>
      </div>
    </aside>
    <main class="pc-inhalt">
      ${z.meldung ? `<div class="hinweis ${z.meldung.art}">${z.meldung.html}</div>` : ''}
      ${{ import: importAnsicht, abschluss: abschlussAnsicht, buchungen: buchungenAnsicht, regeln: regelnAnsicht }[z.ansicht]()}
    </main>
  </div>`;
  window.scrollTo(0, scroll);
}

const opts = (liste, wert, leer, anzeige = (x) => x) => (leer != null ? `<option value="">${h(leer)}</option>` : '') + liste.map((x) => `<option value="${h(x)}" ${x === wert ? 'selected' : ''}>${h(anzeige(x))}</option>`).join('');
const kontenListe = () => z.modell.konten.map((k) => k.name);
const toepfeListe = () => ['Sparplan (verteilt)', ...z.modell.toepfe.map((t) => t.name)];

// ---------------------------------------------------------------- 1. Kontoauszüge abgleichen
function importAnsicht() {
  const a = z.analyse;
  let html = `<h1>Kontoauszüge abgleichen</h1>
    <label class="ablage" id="ablage">
      <input type="file" id="csv-wahl" accept=".csv,text/csv" multiple>
      <strong>CSV-Exporte hierher ziehen</strong> oder klicken und auswählen
      <span class="hint">Gemeinschaftskonto, Visa und Tagesgeld, gern alle auf einmal. Die Dateien bleiben auf diesem Rechner.</span>
    </label>`;
  if (!a) return html + `<p class="hint" style="margin-top:14px">Das Tool vergleicht jede Bankzeile mit euren Einträgen: gleich, zusammengerechnet, vielleicht gleich oder neu. Und es zeigt, was ihr eingetragen habt, das noch nicht auf dem Auszug steht.</p>`;

  // Dateien mit Saldo-Abgleich
  html += `<section><h2>Konten</h2><div class="karten">${a.dateien.map((d) => {
    const s = a.abgleich.find((x) => x.konto === d.konto && d.importieren);
    let abgleich = '';
    if (s) {
      const ok = Math.abs(s.differenz) < 0.005;
      const erklaert = Math.abs(s.rest) < 0.005;
      abgleich = `<div class="abgleich ${ok ? 'ok' : erklaert ? 'erklaert' : 'warn'}">
        <div><span>Bank am ${datum(s.bis)}</span><b>${eur(s.bank)}</b></div>
        <div><span>Datei bis dahin</span><b>${eur(s.datei)}</b></div>
        ${ok ? '<p>Stimmt überein.</p>' : `<p>Differenz ${eur(s.differenz)}${erklaert ? ', vollständig erklärt:' : `, davon unerklärt ${eur(s.rest)}${s.unklar ? ` (vermutlich bei „Bitte entscheiden“)` : ''}:`}</p>
          <ul>${s.neu ? `<li>${eur(s.neu)} neu auf dem Auszug</li>` : ''}
          ${s.fehlen.map((b) => `<li>${h(b.nr)} ${h(b.posten)} (${eur(b.betrag)}, ${datum(b.datum)}) ist eingetragen, aber noch nicht bei der Bank</li>`).join('')}</ul>`}
      </div>`;
    }
    return `<div class="karte ${d.importieren ? '' : 'aus'}"><div class="k-kopf"><b>${h(d.konto || d.kontoText || d.name)}</b><span class="hint">${d.anzahl} ${d.anzahl === 1 ? 'Umsatz' : 'Umsätze'}${d.von ? `, ${datum(d.von)} – ${datum(d.bis)}` : ''}</span></div>
      ${d.hinweis ? `<p class="hint">${h(d.hinweis)}</p>` : ''}${abgleich}</div>`;
  }).join('')}</div>
  <button class="knopf leise klein" id="csv-leeren">Dateien entfernen</button></section>`;

  // Nur in der Datei
  const nur = Object.entries(a.nurDatei || {}).flatMap(([konto, l]) => l.filter((b) => !b.vorStichtag).map((b) => ({ ...b, konto })));
  if (nur.length) {
    html += `<section><h2>Eingetragen, aber nicht auf dem Auszug <span class="anz">${nur.length}</span></h2>
      <p class="hint">Diese Buchungen habt ihr eingetragen, im Zeitraum des Auszugs findet die Bank sie aber nicht. Meist ist die Zahlung noch nicht gebucht. Sonst stimmt Betrag oder Konto nicht – dann unter „Buchungen“ korrigieren.</p>
      <div class="tabelle-rahmen"><table class="umsaetze"><thead><tr><th>Nr</th><th>Datum</th><th>Konto</th><th class="r">Betrag</th><th>Wofür</th><th></th></tr></thead><tbody>
      ${nur.map((b) => `<tr class="gewaehlt"><td>${h(b.nr)}</td><td>${datum(b.datum)}</td><td>${h(b.konto)}</td><td class="r">${eur(b.betrag)}</td><td>${h(b.posten)}</td><td><button class="link" data-zu-buchung="${h(b.nr)}">bearbeiten</button></td></tr>`).join('')}
      </tbody></table></div></section>`;
  }

  // Bankzeilen
  const zaehle = (st) => a.zeilen.filter((x) => x.status === st).length;
  const filter = [['neu', 'Neu'], ['vielleicht', 'Bitte entscheiden'], ['eingetragen', 'Schon eingetragen'], ['vor-stichtag', 'Vor dem Stichtag'], ['gegenbuchung', 'Gegenbuchungen'], ['importiert', 'Schon importiert'], ['alle', 'Alle']]
    .filter(([id]) => id === 'alle' || id === 'neu' || zaehle(id));
  const liste = a.zeilen.filter((x) => z.filter === 'alle' || x.status === z.filter);
  const m = z.modell;
  const bereiche = m.listen?.bereiche || [];
  const monateWahl = [...new Set([...a.zeilen.map((x) => x.buchung.giltFuer), ...(m.listen?.monate || [])])].sort();
  const sel = (feld, wert, liste, leer, anz) => `<select data-feld="${feld}">${opts(liste, wert, leer, anz)}</select>`;
  const felder = (b) => `<td>${sel('art', b.art, ['Ausgabe', 'Einnahme', 'Umbuchung'])}</td>
    <td class="vn">${sel('von', b.von, kontenListe())}${sel('nach', b.nach, kontenListe())}</td>
    <td><input data-feld="posten" value="${h(b.posten)}"></td>
    <td>${sel('bereich', b.bereich, bereiche, 'bitte wählen')}</td>
    <td>${sel('topf', b.topf, toepfeListe(), '–')}</td>
    <td>${sel('giltFuer', b.giltFuer, monateWahl, null, (x) => monatsname(x, true))}</td>
    <td><input data-feld="kommentar" value="${h(b.kommentar)}"></td>`;

  html += `<section><h2>Bankzeilen</h2>
    <div class="filter">${filter.map(([id, name]) => `<button aria-pressed="${z.filter === id}" data-filter="${id}" class="${id === 'vielleicht' ? 'achtung' : ''}">${name} <span class="anz">${id === 'alle' ? a.zeilen.length : zaehle(id)}</span></button>`).join('')}</div>
    ${z.filter !== 'alle' ? `<p class="hint">${h(STATUS[z.filter]?.[1] || '')}</p>` : ''}
    ${liste.length ? `<div class="tabelle-rahmen"><table class="umsaetze"><thead><tr>
      <th><input type="checkbox" id="alle-wahl" aria-label="Alle auswählen"></th><th></th><th>Datum</th><th>Konto</th><th class="r">Betrag</th><th>Text der Bank</th>
      <th>Art</th><th>Von → Nach</th><th>Wofür</th><th>Bereich</th><th>Topf</th><th>Gilt für</th><th>Kommentar</th><th></th></tr></thead><tbody>
      ${liste.map((x) => {
        const b = x.buchung;
        const wirdNeu = x.status === 'neu' || x.status === 'vor-stichtag' || (x.status === 'vielleicht' && x.entscheidung === 'neu');
        const fehlt = x.auswahl && wirdNeu && !b.bereich;
        const sperre = x.status === 'importiert' || x.status === 'gegenbuchung' || (x.status === 'vielleicht' && !x.entscheidung);
        let rechts;
        if (x.status === 'vielleicht') {
          const k = x.kandidat;
          const anderer = Math.abs(Math.abs(k.betrag) - Math.abs(x.zeile.betrag)) >= 0.005;
          rechts = x.entscheidung === 'neu' ? felder(b) + `<td><button class="link" data-entscheid="">zurück</button></td>`
            : `<td colspan="7"><div class="entscheid"><span>${h(x.grund)}</span>
              <button class="knopf klein ${x.entscheidung === 'gleich' ? '' : 'leise'}" data-entscheid="gleich">Ist dieselbe</button>
              ${anderer ? `<button class="knopf klein ${x.entscheidung === 'korrigieren' ? '' : 'leise'}" data-entscheid="korrigieren">Dieselbe, Betrag auf ${eur(Math.abs(x.zeile.betrag))} korrigieren</button>` : ''}
              <button class="knopf klein leise" data-entscheid="neu">Ist neu</button></div></td><td></td>`;
        } else if (x.status === 'neu' || x.status === 'vor-stichtag') {
          rechts = felder(b) + `<td><button class="link" data-regel-aus="${h(x.id)}" title="Für ähnliche Bankzeilen künftig automatisch so zuordnen">als Regel</button></td>`;
        } else {
          rechts = `<td colspan="7" class="hint">${x.status === 'eingetragen' ? `→ ${h(x.grund)}` : h(x.grund || STATUS[x.status][1])}</td><td></td>`;
        }
        return `<tr data-id="${h(x.id)}" class="${x.auswahl ? 'gewaehlt' : ''} ${fehlt ? 'fehlt' : ''} st-${x.status}">
          <td><input type="checkbox" data-wahl ${x.auswahl ? 'checked' : ''} ${sperre ? 'disabled' : ''}></td>
          <td><span class="ampel ${x.ampel}" title="${{ sicher: 'sicher erkannt', pruefen: 'bitte prüfen', unbekannt: 'nicht erkannt' }[x.ampel]}"></span></td>
          <td class="nw">${datum(x.zeile.datum)}</td><td class="nw">${h(x.konto)}</td>
          <td class="r nw ${x.zeile.betrag > 0 ? 'plus' : ''}">${eur(x.zeile.betrag)}</td>
          <td title="${h(x.zeile.info)}"><div class="banktext">${h(x.zeile.info)}</div>${x.grund && (x.status === 'neu' || x.status === 'vor-stichtag') ? `<div class="grund">${h(x.grund)}</div>` : ''}</td>
          ${rechts}</tr>`;
      }).join('')}</tbody></table></div>` : `<div class="leer">${z.filter === 'neu' ? 'Nichts Neues: Alles auf dem Auszug ist schon eingetragen oder liegt vor dem Stichtag.' : 'Nichts in dieser Gruppe.'}</div>`}
  </section>`;

  const p = plan();
  const offen = a.zeilen.filter((x) => x.status === 'vielleicht' && !x.entscheidung).length;
  if (z.bestaetigen) {
    return html + `<div class="pc-leiste frage"><span>Wirklich ${beschreibe(p)} in die Datei schreiben?</span>
      <span><button class="knopf leise" id="abbrechen">Abbrechen</button> <button class="knopf" id="uebernehmen">Ja, übernehmen</button></span></div>`;
  }
  const ohneBereich = p.neu.filter((x) => !x.buchung.bereich).length;
  const nichts = !p.neu.length && !p.ergaenzen.length && !p.korrekturen.length;
  html += `<div class="pc-leiste">
    <span>${nichts ? 'Nichts zu übernehmen' : beschreibe(p)}${offen ? ` · <b class="neg">${offen} noch zu entscheiden</b>` : ''}${ohneBereich ? ` · <b class="neg">${ohneBereich} ohne Bereich</b>` : ''}</span>
    <button class="knopf" id="uebernehmen" ${nichts || ohneBereich || z.schreibt ? 'disabled' : ''}>${z.schreibt ? 'Wird geschrieben …' : 'In die Datei übernehmen'}</button>
  </div>`;
  return html;
}

function plan() {
  const a = z.analyse;
  const neu = a.zeilen.filter((x) => x.auswahl && (x.status === 'neu' || x.status === 'vor-stichtag' || (x.status === 'vielleicht' && x.entscheidung === 'neu')))
    .sort((p, q) => p.zeile.datum.localeCompare(q.zeile.datum));
  const ids = new Map(); // Nr → Import-IDs
  const add = (nr, id) => { if (!ids.has(nr)) ids.set(nr, []); ids.get(nr).push(id); };
  for (const x of a.zeilen) {
    if (x.auswahl && x.status === 'eingetragen') x.treffer.forEach((nr) => add(nr, x.id));
    if (x.status === 'vielleicht' && (x.entscheidung === 'gleich' || x.entscheidung === 'korrigieren')) add(x.kandidat.nr, x.id);
  }
  const ergaenzen = [...ids].map(([nr, l]) => {
    const alt = z.modell.buchungen.find((b) => b.nr === nr)?.importId || '';
    return { nr, importId: [...new Set([...alt.split(/\s+/).filter(Boolean), ...l])].join(' ') };
  });
  const korrekturen = a.zeilen.filter((x) => x.status === 'vielleicht' && x.entscheidung === 'korrigieren').map((x) => ({ nr: x.kandidat.nr, betrag: Math.abs(x.zeile.betrag) }));
  return { neu, ergaenzen, korrekturen };
}
const beschreibe = (p) => [p.neu.length ? `${p.neu.length} neue Buchung${p.neu.length === 1 ? '' : 'en'} (${eur(p.neu.reduce((s, x) => s + x.zeile.betrag, 0))})` : '',
  p.ergaenzen.length ? `${p.ergaenzen.length}× Import-ID ergänzen` : '', p.korrekturen.length ? `${p.korrekturen.length}× Betrag korrigieren` : ''].filter(Boolean).join(', ');

async function dateienHinzu(files) {
  for (const f of files) {
    if (!/\.csv$/i.test(f.name)) continue;
    const text = csvText(new Uint8Array(await f.arrayBuffer()));
    z.dateien = z.dateien.filter((d) => d.name !== f.name).concat({ name: f.name, text });
  }
  neuAnalysieren();
  z.filter = z.analyse.zeilen.some((x) => x.status === 'vielleicht') ? 'vielleicht' : 'neu';
  z.meldung = null; z.bestaetigen = null;
  zeichne();
}

async function uebernehmen() {
  const p = plan();
  if (!z.bestaetigen) { z.bestaetigen = true; zeichne(); return; }
  z.bestaetigen = null; z.schreibt = true; zeichne();
  try {
    const r = await od.buchungenSchreiben(z.ort, p.neu.map((x) => ({ ...x.buchung, quelle: 'CSV' })), p.ergaenzen, p.korrekturen);
    z.schreibt = false;
    z.meldung = { art: 'info', html: `Übernommen: ${r.nummern.length} neue Buchungen${r.nummern.length ? ` (${h(r.nummern[0])}${r.nummern.length > 1 ? ' bis ' + h(r.nummern[r.nummern.length - 1]) : ''})` : ''}${r.ergaenzt ? `, ${r.ergaenzt}× Import-ID ergänzt` : ''}${p.korrekturen.length ? `, ${p.korrekturen.length}× Betrag korrigiert` : ''}. Die Bankzeilen stehen jetzt unter „Schon importiert“.` };
    await laden();
  } catch (e) {
    z.schreibt = false;
    z.meldung = { art: 'warn', html: `Das Übernehmen hat nicht geklappt: ${h(e.message)} Bitte „Neu laden“ und dann nachsehen, bevor ihr es noch einmal versucht – schon Übernommenes erkennt das Tool dann als „Schon importiert“.` };
    zeichne();
  }
}

// ---------------------------------------------------------------- 2. Monatsabschluss
function abschlussAnsicht() {
  const m = z.modell;
  const erster = m.listen?.monate?.[0] || '0000-00';
  const jetzt = heute().slice(0, 7);
  const liste = monate(m).filter((x) => x >= erster && x <= jetzt);
  if (!liste.length) liste.push(jetzt);
  const mo = liste.includes(z.abschlussMonat) ? z.abschlussMonat : liste[liste.length - 1];
  z.abschlussMonat = mo;
  const v = monatsansicht(m, mo, heute());
  const ende = monatsende(mo);
  const konten = m.konten.filter((k) => k.imTool && k.status !== 'geschlossen');
  // Vorschlag aus geladenem Auszug: Saldo minus alles, was nach dem Monatsende gebucht wurde
  const ausAuszug = (konto) => {
    if (!z.analyse) return null;
    const d = z.analyse.dateien.find((x) => x.konto === konto && x.importieren);
    if (!d || !d.bis || d.bis < ende) return null;
    const nachher = z.analyse.zeilen.filter((x) => x.konto === konto && x.zeile.datum > ende).reduce((s, x) => s + x.zeile.betrag, 0);
    return Math.round((d.saldo - nachher) * 100) / 100;
  };
  const punkte = [];
  const p = (ok, text, detail = '') => punkte.push({ ok, text, detail });
  for (const u of v.ueberweisungen) p(u.erledigt, `${u.person} an ${u.nach}: ${eur(u.betrag)}`, u.erledigt ? 'eingegangen' : u.gebucht ? `${eur(u.gebucht)} da, ${eur(u.offen)} fehlen` : 'noch nicht gebucht');
  const offenGeplant = v.vergangen || v.laufend ? v.ablauf.offenGeplant.filter((e) => v.vergangen || (e.tag ?? 99) <= Number(heute().slice(8, 10))) : [];
  p(!offenGeplant.length, 'Geplante Zahlungen auf dem Gemeinschaftskonto gebucht', offenGeplant.length ? 'noch offen: ' + offenGeplant.map((e) => `${e.titel} ${eur(e.betrag)}`).join(', ') : 'alle gebucht');
  p(Math.abs(v.nichtInToepfen) < 0.005, '„Nicht in Töpfen“ ist 0', eur(v.nichtInToepfen));
  for (const k of v.konten) p(k.bankMonatsende != null && Math.abs(k.differenz) < 0.005, `Bankstand ${k.name} eingetragen und stimmt`, k.bankMonatsende == null ? 'noch nicht eingetragen' : `Differenz ${eur(k.differenz)}`);
  if (z.analyse) {
    const offenImport = z.analyse.zeilen.filter((x) => x.zeile.datum.slice(0, 7) === mo && (x.status === 'neu' || x.status === 'vielleicht'));
    p(!offenImport.length, 'Kontoauszüge abgeglichen', offenImport.length ? `${offenImport.length} Bankzeilen noch nicht übernommen oder entschieden` : 'alles übernommen');
  } else p(null, 'Kontoauszüge abgleichen', 'noch keine Exporte geladen');
  p(!v.offenePunkte.length, 'Offene Punkte in der Datei', v.offenePunkte.length ? `${v.offenePunkte.length} offen` : 'keine');

  return `<h1>Monatsabschluss</h1>
    <div class="zeile-wahl"><label>Monat <select id="abschluss-monat">${opts(liste, mo, null, (x) => monatsname(x, true))}</select></label>
    <span class="hint">Stand ${v.vergangen ? 'am Monatsende' : 'heute, Monat läuft noch'}</span></div>
    <section><h2>Kontostände am ${datum(ende)}</h2>
    <div class="tabelle-rahmen"><table class="umsaetze"><thead><tr><th>Konto</th><th class="r">Laut Datei</th><th class="r">Laut Bank (eingetragen)</th><th class="r">Differenz</th><th>Bankstand eintragen</th></tr></thead><tbody>
    ${konten.map((k) => {
      const kk = v.konten.find((x) => x.name === k.name);
      const vorschlag = ausAuszug(k.name);
      return `<tr class="gewaehlt"><td>${h(k.name)}</td><td class="r">${eur(kk.standMonatsende)}</td>
        <td class="r">${kk.bankMonatsende == null ? '–' : eur(kk.bankMonatsende)}</td>
        <td class="r ${kk.differenz && Math.abs(kk.differenz) >= 0.005 ? 'neg' : ''}">${kk.differenz == null ? '' : eur(kk.differenz)}</td>
        <td class="nw"><input data-bank="${h(k.name)}" placeholder="0,00" value="${vorschlag != null && kk.bankMonatsende == null ? vorschlag.toFixed(2).replace('.', ',') : ''}" style="width:110px">
          <button class="knopf klein leise" data-bank-speichern="${h(k.name)}">Speichern</button>
          ${vorschlag != null ? `<span class="hint">aus dem Auszug: ${eur(vorschlag)}</span>` : ''}</td></tr>`;
    }).join('')}</tbody></table></div>
    <p class="hint">Den Stand laut Bank zum Monatsende eintragen. Liegt ein Auszug bis nach dem Monatsende vor, rechnet das Tool den Stand am Monatsende selbst aus.</p></section>
    <section><h2>Checkliste</h2><ul class="checkliste">${punkte.map((x) => `<li class="${x.ok === null ? 'offen' : x.ok ? 'ok' : 'nein'}"><span class="mark">${x.ok === null ? '–' : x.ok ? '✓' : '!'}</span><div><b>${h(x.text)}</b><div class="hint">${h(x.detail)}</div></div></li>`).join('')}</ul></section>
    <section><h2>Ende ${h(v.name)}</h2><div class="ende"><div><div style="font-weight:600">Gemeinschaftskonto${v.vergangen ? '' : ', voraussichtlich'}</div><div class="hint">${v.vergangen ? 'laut Buchungen' : 'wenn alles wie geplant kommt und abgeht'}</div></div><div class="betrag ${v.ablauf.ende < 0 ? 'neg' : ''}">${eur(v.ablauf.ende)}</div></div></section>`;
}

// ---------------------------------------------------------------- 3. Buchungen
function buchungenAnsicht() {
  const m = z.modell;
  const q = z.suche.trim().toLowerCase();
  const liste = m.buchungen.filter((b) => (!z.bMonat || b.giltFuer === z.bMonat) && (!z.bKonto || b.von === z.bKonto || b.nach === z.bKonto) && (!z.bArt || b.art === z.bArt)
    && (!q || [b.nr, b.posten, b.beschreibung, b.kommentar, b.bereich, b.topf, String(b.betrag).replace('.', ',')].some((x) => (x || '').toLowerCase().includes(q))))
    .sort((a, b) => (b.datum || '').localeCompare(a.datum || '') || b.nr.localeCompare(a.nr));
  const summe = (konto) => liste.reduce((s, b) => s + (b.nach === konto ? b.betrag : 0) - (b.von === konto ? b.betrag : 0), 0);
  const monateWahl = [...new Set(m.buchungen.map((b) => b.giltFuer).filter(Boolean))].sort();
  return `<h1>Buchungen</h1>
    <div class="zeile-wahl">
      <input id="b-suche" placeholder="Suchen: Nummer, Wofür, Text, Betrag …" value="${h(z.suche)}" style="flex:1;min-width:220px">
      <select id="b-monat">${opts(monateWahl, z.bMonat, 'alle Monate', (x) => 'gilt für ' + monatsname(x, true))}</select>
      <select id="b-konto">${opts(kontenListe(), z.bKonto, 'alle Konten')}</select>
      <select id="b-art">${opts(['Ausgabe', 'Einnahme', 'Umbuchung'], z.bArt, 'alle Arten')}</select>
    </div>
    <p class="hint">${liste.length} Buchungen${z.bKonto ? `, Saldo auf ${h(z.bKonto)}: ${eur(summe(z.bKonto))}` : ''}. Zum Bearbeiten eine Zeile anklicken.</p>
    <div class="tabelle-rahmen"><table class="umsaetze buchliste"><thead><tr><th>Nr</th><th>Datum</th><th>Gilt für</th><th>Art</th><th>Von → Nach</th><th class="r">Betrag</th><th>Wofür</th><th>Bereich</th><th>Topf</th><th>Kommentar</th><th>Quelle</th></tr></thead><tbody>
    ${liste.map((b) => z.bearbeiten === b.nr ? editor(b) : `<tr class="gewaehlt klickbar" data-nr="${h(b.nr)}"><td>${h(b.nr)}</td><td class="nw">${datum(b.datum)}</td><td class="nw">${b.giltFuer ? monatsname(b.giltFuer).slice(0, 3) + ' ' + b.giltFuer.slice(2, 4) : ''}</td>
      <td>${h(b.art)}</td><td class="nw">${h(b.von)} → ${h(b.nach)}</td><td class="r nw">${eur(b.betrag)}</td><td>${h(b.posten)}${istAusgleich(b) ? ' <span class="chip ausgleich">Ausgleich</span>' : ''}</td><td>${h(b.bereich)}</td><td>${h(b.topf)}</td>
      <td class="hint">${h(b.kommentar)}</td><td class="hint">${b.quelle === 'CSV' ? 'Auszug' : 'von Hand'}${b.importId ? ' ✓' : ''}</td></tr>`).join('')}
    </tbody></table></div>`;
}
function editor(b) {
  const m = z.modell;
  const monateWahl = [...new Set([b.giltFuer, ...(m.listen?.monate || []), ...m.buchungen.map((x) => x.giltFuer)].filter(Boolean))].sort();
  const f = (name, inhalt) => `<label class="feld"><span>${name}</span>${inhalt}</label>`;
  return `<tr class="editor gewaehlt"><td colspan="11"><form id="b-form" data-nr="${h(b.nr)}">
    <div class="editor-kopf"><b>${h(b.nr)} bearbeiten</b><span class="hint">${b.quelle === 'CSV' ? 'aus dem Auszug importiert' : 'von Hand eingetragen'}${b.importId ? ', mit dem Auszug abgeglichen' : ''}</span></div>
    <div class="editor-felder">
      ${f('Art', `<select name="art">${opts(['Ausgabe', 'Einnahme', 'Umbuchung'], b.art)}</select>`)}
      ${f('Betrag', `<input name="betrag" value="${String(b.betrag.toFixed(2)).replace('.', ',')}">`)}
      ${f('Bezahlt am', `<input type="date" name="datum" value="${h(b.datum)}">`)}
      ${f('Gilt für', `<select name="giltFuer">${opts(monateWahl, b.giltFuer, null, (x) => monatsname(x, true))}</select>`)}
      ${f('Von', `<select name="von">${opts(kontenListe(), b.von)}</select>`)}
      ${f('Nach', `<select name="nach">${opts(kontenListe(), b.nach)}</select>`)}
      ${f('Wofür', `<input name="posten" value="${h(b.posten)}">`)}
      ${f('Bereich', `<select name="bereich">${opts(m.listen?.bereiche || [], b.bereich, '–')}</select>`)}
      ${f('Topf', `<select name="topf">${opts(toepfeListe(), b.topf, '–')}</select>`)}
      ${f('Gezahlt von / Ausgleich', `<select name="gezahltVon">${opts(['Kai', 'Natalie', AUSGLEICH.topf, AUSGLEICH.korrektur], b.gezahltVon, '–')}</select>`)}
      ${f('Verwendungszweck', `<input name="beschreibung" value="${h(b.beschreibung)}">`)}
      ${f('Kommentar', `<input name="kommentar" value="${h(b.kommentar)}">`)}
    </div>
    <div class="editor-knoepfe">
      <button class="knopf" type="submit">Speichern</button>
      <button class="knopf leise" type="button" id="b-abbrechen">Abbrechen</button>
      ${z.loeschenFrage ? `<span class="neg">Wirklich ${h(b.nr)} löschen? Das lässt sich nur in Excel rückgängig machen.</span><button class="knopf leise" type="button" id="b-loeschen-ja">Ja, löschen</button>`
        : `<button class="knopf leise" type="button" id="b-loeschen">Löschen</button>`}
    </div></form></td></tr>`;
}

// ---------------------------------------------------------------- 4. Regeln
function regelnAnsicht() {
  const m = z.modell;
  return `<h1>Regeln</h1>
    <p class="hint">Mit Regeln ordnet der Abgleich Bankzeilen automatisch zu, z. B. „Text enthält ‚Kita‘ → Wofür ‚Kita-Beitrag‘, Bereich ‚Kinder‘“. Gespeichert in „${h(od.REGELN_DATEI)}“ neben eurer Excel-Datei in OneDrive. Die Excel-Datei selbst bleibt unverändert.</p>
    <div class="tabelle-rahmen"><table class="umsaetze"><thead><tr><th>Bank-Text enthält</th><th>nur Konto</th><th>nur Betrag</th><th>→ Art</th><th>Wofür</th><th>Bereich</th><th>Topf</th><th>Gilt für</th><th></th></tr></thead><tbody>
    ${z.regeln.map((r, i) => `<tr class="gewaehlt" data-regel="${i}">
      <td><input data-r="enthaelt" value="${h(r.enthaelt)}"></td>
      <td><select data-r="konto">${opts(kontenListe(), r.konto, 'alle')}</select></td>
      <td><input data-r="betrag" value="${r.betrag ? String(r.betrag).replace('.', ',') : ''}" placeholder="egal" style="width:80px"></td>
      <td><select data-r="setze.art">${opts(['Ausgabe', 'Einnahme', 'Umbuchung'], r.setze?.art, 'wie erkannt')}</select></td>
      <td><input data-r="setze.posten" value="${h(r.setze?.posten)}"></td>
      <td><select data-r="setze.bereich">${opts(m.listen?.bereiche || [], r.setze?.bereich, '–')}</select></td>
      <td><select data-r="setze.topf">${opts(toepfeListe(), r.setze?.topf, '–')}</select></td>
      <td><select data-r="setze.giltFuer">${opts(['+1'], r.setze?.giltFuer, 'Monat der Zahlung', () => 'Folgemonat')}</select></td>
      <td><button class="link" data-regel-weg="${i}">entfernen</button></td></tr>`).join('') || '<tr><td colspan="9" class="hint">Noch keine Regeln. Im Abgleich bei einer Bankzeile auf „als Regel“ klicken oder hier eine neue anlegen.</td></tr>'}
    </tbody></table></div>
    <div class="editor-knoepfe"><button class="knopf leise" id="regel-neu">Neue Regel</button><button class="knopf" id="regeln-speichern">Regeln speichern</button></div>`;
}

async function regelnSichern(meldung) {
  try {
    z.regeln = z.regeln.filter((r) => (r.enthaelt || '').trim());
    await od.regelnSpeichern(z.item, z.regeln);
    if (z.dateien.length) neuAnalysieren();
    z.meldung = { art: 'info', html: meldung || 'Regeln gespeichert.' };
  } catch (e) { z.meldung = { art: 'warn', html: 'Regeln konnten nicht gespeichert werden: ' + h(e.message) }; }
  zeichne();
}

// ---------------------------------------------------------------- Ereignisse
const imPc = () => document.body.classList.contains('pc-modus');
document.addEventListener('click', async (ev) => {
  if (!imPc()) return;
  const t = ev.target.closest('button, input[type=checkbox], tr.klickbar');
  if (!t) return;
  if (t.id === 'pc-anmelden') { ms.anmelden(); return; }
  if (t.id === 'pc-neu-laden') { z.meldung = null; laden(); return; }
  if (t.dataset.ansicht) { z.ansicht = t.dataset.ansicht; z.meldung = null; z.bearbeiten = null; zeichne(); window.scrollTo(0, 0); return; }
  // Abgleich
  if (t.dataset.filter) { z.filter = t.dataset.filter; zeichne(); return; }
  if (t.id === 'csv-leeren') { z.dateien = []; z.analyse = null; zeichne(); return; }
  if (t.id === 'uebernehmen') { uebernehmen(); return; }
  if (t.id === 'abbrechen') { z.bestaetigen = null; zeichne(); return; }
  if (t.id === 'alle-wahl') {
    const ids = [...document.querySelectorAll('tr[data-id]')].map((r) => r.dataset.id);
    z.analyse.zeilen.forEach((x) => { if (ids.includes(x.id) && x.status !== 'importiert' && x.status !== 'gegenbuchung' && !(x.status === 'vielleicht' && !x.entscheidung)) x.auswahl = t.checked; });
    z.bestaetigen = null; zeichne(); return;
  }
  if (t.hasAttribute('data-wahl')) { const x = z.analyse.zeilen.find((r) => r.id === t.closest('tr').dataset.id); x.auswahl = t.checked; z.bestaetigen = null; zeichne(); return; }
  if (t.dataset.entscheid != null) {
    const x = z.analyse.zeilen.find((r) => r.id === t.closest('tr').dataset.id);
    x.entscheidung = t.dataset.entscheid || null; x.auswahl = x.entscheidung === 'neu'; z.bestaetigen = null; zeichne(); return;
  }
  if (t.dataset.regelAus) {
    const x = z.analyse.zeilen.find((r) => r.id === t.dataset.regelAus);
    const b = x.buchung;
    z.regeln.push({ enthaelt: kurzname(x.zeile.info, x.zeile.umsatzart), konto: x.konto, setze: { art: b.art, posten: b.posten, bereich: b.bereich, topf: b.topf || undefined } });
    z.ansicht = 'regeln';
    z.meldung = { art: 'info', html: 'Neue Regel aus der Bankzeile vorbereitet. Den Suchtext bei Bedarf kürzen und dann „Regeln speichern“.' };
    zeichne(); return;
  }
  if (t.dataset.zuBuchung) { z.ansicht = 'buchungen'; z.suche = t.dataset.zuBuchung; z.bMonat = ''; z.bKonto = ''; z.bArt = ''; z.bearbeiten = t.dataset.zuBuchung; z.loeschenFrage = false; zeichne(); return; }
  // Abschluss
  if (t.dataset.bankSpeichern) {
    const konto = t.dataset.bankSpeichern;
    const wert = zahl(document.querySelector(`input[data-bank="${CSS.escape(konto)}"]`).value);
    if (!isFinite(wert)) { z.meldung = { art: 'warn', html: 'Bitte einen Betrag eingeben, z. B. 396,27.' }; zeichne(); return; }
    t.disabled = true; t.textContent = 'Speichert …';
    try {
      await od.kontostandEintragen(z.ort, z.abschlussMonat, konto, wert);
      z.meldung = { art: 'info', html: `Bankstand ${h(konto)} für ${monatsname(z.abschlussMonat, true)} gespeichert: ${eur(wert)}.` };
      await laden();
    } catch (e) { z.meldung = { art: 'warn', html: h(e.message) }; zeichne(); }
    return;
  }
  // Buchungen
  if (t.classList.contains('klickbar')) { z.bearbeiten = t.dataset.nr; z.loeschenFrage = false; zeichne(); return; }
  if (t.id === 'b-abbrechen') { z.bearbeiten = null; z.loeschenFrage = false; zeichne(); return; }
  if (t.id === 'b-loeschen') { z.loeschenFrage = true; zeichne(); return; }
  if (t.id === 'b-loeschen-ja') {
    const nr = z.bearbeiten;
    try { await od.buchungLoeschen(z.ort, nr); z.meldung = { art: 'info', html: `${h(nr)} gelöscht.` }; z.bearbeiten = null; z.loeschenFrage = false; await laden(); }
    catch (e) { z.meldung = { art: 'warn', html: 'Löschen hat nicht geklappt: ' + h(e.message) }; zeichne(); }
    return;
  }
  // Regeln
  if (t.id === 'regel-neu') { z.regeln.push({ enthaelt: '', setze: {} }); zeichne(); return; }
  if (t.dataset.regelWeg) { z.regeln.splice(Number(t.dataset.regelWeg), 1); zeichne(); return; }
  if (t.id === 'regeln-speichern') { regelnSichern(); }
});

document.addEventListener('change', (ev) => {
  if (!imPc()) return;
  const t = ev.target;
  if (t.id === 'csv-wahl') { dateienHinzu([...t.files]); return; }
  if (t.id === 'abschluss-monat') { z.abschlussMonat = t.value; zeichne(); return; }
  if (t.id === 'b-monat') { z.bMonat = t.value; zeichne(); return; }
  if (t.id === 'b-konto') { z.bKonto = t.value; zeichne(); return; }
  if (t.id === 'b-art') { z.bArt = t.value; zeichne(); return; }
  if (t.dataset.r != null) {
    const r = z.regeln[Number(t.closest('tr').dataset.regel)];
    const [a, b] = t.dataset.r.split('.');
    const wert = a === 'betrag' ? (isFinite(zahl(t.value)) ? zahl(t.value) : undefined) : (t.value || undefined);
    if (b) { r.setze = r.setze || {}; r.setze[b] = wert; } else r[a] = wert;
    return;
  }
  const feld = t.dataset?.feld; if (!feld || !z.analyse) return;
  const x = z.analyse.zeilen.find((r) => r.id === t.closest('tr').dataset.id);
  x.buchung[feld] = t.value; x.geaendert = true;
  if (feld === 'bereich' || feld === 'art') zeichne();
});
document.addEventListener('input', (ev) => {
  if (!imPc() || ev.target.id !== 'b-suche') return;
  z.suche = ev.target.value;
  clearTimeout(z.sucheTimer);
  z.sucheTimer = setTimeout(() => { const pos = ev.target.selectionStart; zeichne(); const el = document.getElementById('b-suche'); el.focus(); el.setSelectionRange(pos, pos); }, 250);
});
document.addEventListener('submit', async (ev) => {
  if (!imPc() || ev.target.id !== 'b-form') return;
  ev.preventDefault();
  const d = Object.fromEntries(new FormData(ev.target));
  const betrag = zahl(d.betrag);
  if (!(betrag > 0)) { z.meldung = { art: 'warn', html: 'Bitte einen Betrag größer als 0 eingeben.' }; zeichne(); return; }
  if (d.von === d.nach) { z.meldung = { art: 'warn', html: 'Von und Nach sind dasselbe Konto.' }; zeichne(); return; }
  const nr = ev.target.dataset.nr;
  try {
    await od.buchungAendern(z.ort, nr, { ...d, betrag });
    z.meldung = { art: 'info', html: `${h(nr)} gespeichert.` }; z.bearbeiten = null;
    await laden();
  } catch (e) { z.meldung = { art: 'warn', html: 'Speichern hat nicht geklappt: ' + h(e.message) }; zeichne(); }
});
document.addEventListener('dragover', (ev) => { if (imPc()) { ev.preventDefault(); document.getElementById('ablage')?.classList.add('drueber'); } });
document.addEventListener('dragleave', () => document.getElementById('ablage')?.classList.remove('drueber'));
document.addEventListener('drop', (ev) => {
  if (!imPc()) return;
  ev.preventDefault();
  if (z.modell && ev.dataTransfer?.files?.length) { z.ansicht = 'import'; dateienHinzu([...ev.dataTransfer.files]); }
});

