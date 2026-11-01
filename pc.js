// PC-Ansicht „Unser Geld“ (Adresse …/finanz-app/#pc): CSV-Import, später Abschluss, Buchungsliste, Regeln.
import { leseBuchhaltung } from './kern/lesen.js';
import { eur, monatsname } from './kern/rechnen.js';
import { analysieren, csvText } from './kern/csv.js';
import * as ms from './anmeldung.js';
import * as od from './onedrive.js';

const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const datum = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
const zeit = (t) => new Date(t).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' });
const STATUS = {
  neu: ['Neu', 'Wird als neue Buchung eingetragen'],
  eingetragen: ['Schon eingetragen', 'Gibt es schon in der Datei – es wird nur die Import-ID ergänzt'],
  'vor-stichtag': ['Vor dem Stichtag', 'Steckt schon im Anfangsbestand – normalerweise nicht importieren'],
  gegenbuchung: ['Gegenbuchung', 'Die andere Seite einer Umbuchung – wird nicht doppelt gebucht'],
  importiert: ['Schon importiert', 'Diese Bankzeile ist schon in der Datei'],
};

const z = { modell: null, item: null, ort: null, ansicht: 'import', dateien: [], analyse: null, filter: 'neu', meldung: null, schreibt: false };
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
    <p>Hier importiert ihr die Kontoauszüge (CSV) und schließt den Monat ab. Dafür ist die Anmeldung mit dem Microsoft-Konto nötig, mit dem ihr die Datei in OneDrive öffnet.</p>
    <button class="knopf" id="pc-anmelden">Mit Microsoft anmelden</button>
    ${fehler ? `<p class="fehler">${h(fehler)}</p>` : ''}
    <p class="hint"><a href="./">Zur Handy-Ansicht</a></p>
  </main>`;
}

async function laden() {
  root.innerHTML = '<main class="pc-start"><p class="hint">Lade die Datei aus OneDrive …</p></main>';
  try {
    const g = await od.dateiFinden();
    if (!g) return zeigeAnmelden('Die Datei wurde in OneDrive nicht gefunden.');
    const bytes = await od.herunterladen(g.ort);
    z.modell = await leseBuchhaltung(bytes); z.item = g.item; z.ort = g.ort;
    if (z.dateien.length) z.analyse = analysieren(z.dateien, z.modell);
    zeichne();
  } catch (e) {
    if (e instanceof od.AnmeldungNoetig) return zeigeAnmelden('Die Anmeldung ist abgelaufen.');
    zeigeAnmelden('Fehler beim Laden: ' + e.message);
  }
}

// ---------------------------------------------------------------- Rahmen
function zeichne() {
  const konto = ms.gemerktesKonto();
  const nav = [['import', 'Kontoauszüge importieren'], ['abschluss', 'Monatsabschluss'], ['buchungen', 'Buchungen'], ['regeln', 'Regeln']];
  root.innerHTML = `<div class="pc">
    <aside class="pc-nav">
      <div class="pc-logo">Unser Geld</div>
      ${nav.map(([id, name]) => `<button class="${z.ansicht === id ? 'an' : ''}" data-ansicht="${id}">${name}</button>`).join('')}
      <div class="pc-fuss">
        <div class="hint">${h(z.item?.name)}</div>
        <div class="hint">geändert ${z.item?.lastModifiedDateTime ? zeit(z.item.lastModifiedDateTime) : '–'}${z.item?.lastModifiedBy?.user?.displayName ? ' von ' + h(z.item.lastModifiedBy.user.displayName) : ''}</div>
        <div class="hint">${z.modell.buchungen.length} Buchungen · angemeldet als ${h(konto || '–')}</div>
        <button class="knopf leise klein" id="pc-neu-laden">Neu laden</button>
        <a class="hint" href="./">Zur Handy-Ansicht</a>
      </div>
    </aside>
    <main class="pc-inhalt">
      ${z.meldung ? `<div class="hinweis ${z.meldung.art}">${z.meldung.html}</div>` : ''}
      ${z.ansicht === 'import' ? importAnsicht() : `<h1>${h(nav.find((n) => n[0] === z.ansicht)[1])}</h1><div class="leer">Kommt als Nächstes.</div>`}
    </main>
  </div>`;
}

// ---------------------------------------------------------------- Import
function importAnsicht() {
  const a = z.analyse;
  let html = `<h1>Kontoauszüge importieren</h1>
    <label class="ablage" id="ablage">
      <input type="file" id="csv-wahl" accept=".csv,text/csv" multiple>
      <strong>CSV-Exporte hierher ziehen</strong> oder klicken und auswählen
      <span class="hint">Gemeinschaftskonto, Visa und Tagesgeld, gern alle auf einmal. Die Dateien bleiben auf diesem Rechner.</span>
    </label>`;
  if (!a) return html;

  // Dateien und Saldo-Abgleich
  html += `<section><h2>Dateien</h2><div class="karten">${a.dateien.map((d) => {
    const s = a.abgleich.find((x) => x.konto === d.konto && d.importieren);
    let abgleich = '';
    if (s) {
      const ok = Math.abs(s.differenz) < 0.005;
      const erklaert = Math.abs(s.rest) < 0.005;
      abgleich = `<div class="abgleich ${ok ? 'ok' : erklaert ? 'erklaert' : 'warn'}">
        <div><span>Bank am ${datum(s.bis)}</span><b>${eur(s.bank)}</b></div>
        <div><span>Datei bis dahin</span><b>${eur(s.datei)}</b></div>
        ${ok ? '<p>Stimmt überein.</p>' : `<p>Differenz ${eur(s.differenz)}${erklaert ? ', erklärt:' : ', davon unerklärt ' + eur(s.rest) + ':'}</p>
          <ul>${s.neu ? `<li>${eur(s.neu)} neu im Export (wird mit dem Import eingetragen)</li>` : ''}
          ${s.fehlen.map((b) => `<li>${h(b.nr)} ${h(b.posten)} (${eur(b.betrag)}, ${datum(b.datum)}) steht in der Datei, aber noch nicht bei der Bank</li>`).join('')}</ul>`}
      </div>`;
    }
    return `<div class="karte ${d.importieren ? '' : 'aus'}"><div class="k-kopf"><b>${h(d.konto || d.kontoText || d.name)}</b><span class="hint">${d.anzahl} ${d.anzahl === 1 ? 'Umsatz' : 'Umsätze'}${d.von ? `, ${datum(d.von)} – ${datum(d.bis)}` : ''}</span></div>
      ${d.hinweis ? `<p class="hint">${h(d.hinweis)}</p>` : ''}${abgleich}</div>`;
  }).join('')}</div>
  <button class="knopf leise klein" id="csv-leeren">Dateien entfernen</button></section>`;

  // Tabelle
  const zaehle = (st) => a.zeilen.filter((x) => x.status === st).length;
  const filter = [['neu', 'Neu'], ['eingetragen', 'Schon eingetragen'], ['vor-stichtag', 'Vor dem Stichtag'], ['gegenbuchung', 'Gegenbuchungen'], ['importiert', 'Schon importiert'], ['alle', 'Alle']]
    .filter(([id]) => id === 'alle' || id === 'neu' || zaehle(id));
  const liste = a.zeilen.filter((x) => z.filter === 'alle' || x.status === z.filter);
  const m = z.modell;
  const konten = m.konten.filter((k) => k.status !== 'geschlossen' || a.zeilen.some((x) => x.buchung.von === k.name || x.buchung.nach === k.name)).map((k) => k.name);
  const bereiche = m.listen?.bereiche || [];
  const toepfe = ['Sparplan (verteilt)', ...m.toepfe.map((t) => t.name)];
  const monateWahl = [...new Set([...a.zeilen.map((x) => x.buchung.giltFuer), ...(m.listen?.monate || [])])].sort();
  const sel = (feld, wert, liste, leer) => `<select data-feld="${feld}">${leer != null ? `<option value="">${leer}</option>` : ''}${liste.map((x) => `<option value="${h(x)}" ${x === wert ? 'selected' : ''}>${h(feld === 'giltFuer' ? monatsname(x, true) : x)}</option>`).join('')}</select>`;

  html += `<section><h2>Umsätze</h2>
    <div class="filter">${filter.map(([id, name]) => `<button aria-pressed="${z.filter === id}" data-filter="${id}">${name} <span class="anz">${id === 'alle' ? a.zeilen.length : zaehle(id)}</span></button>`).join('')}</div>
    ${z.filter !== 'alle' ? `<p class="hint">${h(STATUS[z.filter]?.[1] || '')}</p>` : ''}
    ${liste.length ? `<div class="tabelle-rahmen"><table class="umsaetze"><thead><tr>
      <th><input type="checkbox" id="alle-wahl" aria-label="Alle auswählen"></th><th></th><th>Datum</th><th>Konto</th><th class="r">Betrag</th><th>Text der Bank</th>
      <th>Art</th><th>Von → Nach</th><th>Wofür</th><th>Bereich</th><th>Topf</th><th>Gilt für</th><th>Kommentar</th></tr></thead><tbody>
      ${liste.map((x) => {
        const b = x.buchung;
        const fehlt = x.auswahl && x.status === 'neu' && !b.bereich;
        const bearbeitbar = x.status === 'neu' || x.status === 'vor-stichtag';
        return `<tr data-id="${h(x.id)}" class="${x.auswahl ? 'gewaehlt' : ''} ${fehlt ? 'fehlt' : ''}">
          <td><input type="checkbox" data-wahl ${x.auswahl ? 'checked' : ''} ${x.status === 'importiert' || x.status === 'gegenbuchung' ? 'disabled' : ''}></td>
          <td><span class="ampel ${x.ampel}" title="${{ sicher: 'sicher erkannt', pruefen: 'bitte prüfen', unbekannt: 'nicht erkannt' }[x.ampel]}${x.grund ? ': ' + h(x.grund) : ''}"></span></td>
          <td class="nw">${datum(x.zeile.datum)}</td><td class="nw">${h(x.konto)}</td>
          <td class="r nw ${x.zeile.betrag > 0 ? 'plus' : ''}">${eur(x.zeile.betrag)}</td>
          <td title="${h(x.zeile.info)}"><div class="banktext">${h(x.zeile.info)}</div>${x.grund ? `<div class="grund">${h(x.grund)}</div>` : ''}</td>
          ${bearbeitbar ? `<td>${sel('art', b.art, ['Ausgabe', 'Einnahme', 'Umbuchung'])}</td>
            <td class="vn">${sel('von', b.von, konten)}${sel('nach', b.nach, konten)}</td>
            <td><input data-feld="posten" value="${h(b.posten)}"></td>
            <td>${sel('bereich', b.bereich, bereiche, 'bitte wählen')}</td>
            <td>${sel('topf', b.topf, toepfe, '–')}</td>
            <td>${sel('giltFuer', b.giltFuer, monateWahl)}</td>
            <td><input data-feld="kommentar" value="${h(b.kommentar)}"></td>`
          : `<td colspan="7" class="hint">${x.status === 'eingetragen' ? `→ ${h(x.treffer)}: Import-ID wird ergänzt` : h(STATUS[x.status][1])}</td>`}
        </tr>`;
      }).join('')}</tbody></table></div>` : `<div class="leer">${z.filter === 'neu' ? 'Keine neuen Umsätze. Alles schon in der Datei oder vor dem Stichtag.' : 'Nichts in dieser Gruppe.'}</div>`}
  </section>`;

  const neu = a.zeilen.filter((x) => x.auswahl && (x.status === 'neu' || x.status === 'vor-stichtag'));
  const erg = a.zeilen.filter((x) => x.auswahl && x.status === 'eingetragen');
  const ohneBereich = neu.filter((x) => !x.buchung.bereich).length;
  if (z.bestaetigen) {
    const b = z.bestaetigen;
    return html + `<div class="pc-leiste frage"><span>Wirklich ${b.anzahl} neue Buchung${b.anzahl === 1 ? '' : 'en'} (zusammen ${eur(b.summe)}) in die Datei schreiben${b.erg ? ` und ${b.erg} Import-ID${b.erg === 1 ? '' : 's'} ergänzen` : ''}?</span>
      <span><button class="knopf leise" id="abbrechen">Abbrechen</button> <button class="knopf" id="uebernehmen">Ja, übernehmen</button></span></div>`;
  }
  html += `<div class="pc-leiste">
    <span>${neu.length} neue Buchung${neu.length === 1 ? '' : 'en'}${erg.length ? `, ${erg.length} Import-ID${erg.length === 1 ? '' : 's'} ergänzen` : ''}${ohneBereich ? ` · <b class="neg">${ohneBereich} ohne Bereich</b>` : ''}</span>
    <button class="knopf" id="uebernehmen" ${(!neu.length && !erg.length) || ohneBereich || z.schreibt ? 'disabled' : ''}>${z.schreibt ? 'Wird geschrieben …' : 'In die Datei übernehmen'}</button>
  </div>`;
  return html;
}

async function dateienHinzu(files) {
  for (const f of files) {
    if (!/\.csv$/i.test(f.name)) continue;
    const text = csvText(new Uint8Array(await f.arrayBuffer()));
    z.dateien = z.dateien.filter((d) => d.name !== f.name).concat({ name: f.name, text });
  }
  z.analyse = analysieren(z.dateien, z.modell);
  z.filter = 'neu'; z.meldung = null;
  zeichne();
}

async function uebernehmen() {
  const a = z.analyse;
  const neu = a.zeilen.filter((x) => x.auswahl && (x.status === 'neu' || x.status === 'vor-stichtag'))
    .sort((p, q) => p.zeile.datum.localeCompare(q.zeile.datum));
  const erg = a.zeilen.filter((x) => x.auswahl && x.status === 'eingetragen');
  const summe = neu.reduce((s, x) => s + x.zeile.betrag, 0);
  if (!z.bestaetigen) { z.bestaetigen = { anzahl: neu.length, summe, erg: erg.length }; zeichne(); return; }
  z.bestaetigen = null; z.schreibt = true; zeichne();
  try {
    const r = await od.buchungenSchreiben(z.ort, neu.map((x) => ({ ...x.buchung, quelle: 'CSV' })), erg.map((x) => ({ nr: x.treffer, importId: x.id })));
    z.schreibt = false;
    z.meldung = { art: 'info', html: `Übernommen: ${r.nummern.length} neue Buchungen${r.nummern.length ? ` (${h(r.nummern[0])}${r.nummern.length > 1 ? ' bis ' + h(r.nummern[r.nummern.length - 1]) : ''})` : ''}${r.ergaenzt ? `, ${r.ergaenzt} Import-IDs ergänzt` : ''}.` };
    await laden();
  } catch (e) {
    z.schreibt = false;
    z.meldung = { art: 'warn', html: `Das Übernehmen hat nicht geklappt: ${h(e.message)} Bitte in Excel nachsehen, bevor ihr es noch einmal versucht.` };
    zeichne();
  }
}

// ---------------------------------------------------------------- Ereignisse
document.addEventListener('click', (ev) => {
  if (!document.body.classList.contains('pc-modus')) return;
  const t = ev.target.closest('button, input[type=checkbox]');
  if (!t) return;
  if (t.id === 'pc-anmelden') { ms.anmelden(); return; }
  if (t.id === 'pc-neu-laden') { z.meldung = null; laden(); return; }
  if (t.dataset.ansicht) { z.ansicht = t.dataset.ansicht; zeichne(); return; }
  if (t.dataset.filter) { z.filter = t.dataset.filter; zeichne(); return; }
  if (t.id === 'csv-leeren') { z.dateien = []; z.analyse = null; zeichne(); return; }
  if (t.id === 'uebernehmen') { uebernehmen(); return; }
  if (t.id === 'abbrechen') { z.bestaetigen = null; zeichne(); return; }
  if (t.id === 'alle-wahl') {
    const ids = [...document.querySelectorAll('tr[data-id]')].map((r) => r.dataset.id);
    z.analyse.zeilen.forEach((x) => { if (ids.includes(x.id) && x.status !== 'importiert' && x.status !== 'gegenbuchung') x.auswahl = t.checked; });
    zeichne(); return;
  }
  if (t.hasAttribute('data-wahl')) {
    const x = z.analyse.zeilen.find((r) => r.id === t.closest('tr').dataset.id);
    x.auswahl = t.checked; z.bestaetigen = null; zeichne();
  }
});
document.addEventListener('change', (ev) => {
  if (!document.body.classList.contains('pc-modus')) return;
  const t = ev.target;
  if (t.id === 'csv-wahl') { dateienHinzu([...t.files]); return; }
  const feld = t.dataset?.feld; if (!feld) return;
  const x = z.analyse.zeilen.find((r) => r.id === t.closest('tr').dataset.id);
  x.buchung[feld] = t.value;
  if (feld === 'bereich' || feld === 'art') zeichne();
});
document.addEventListener('dragover', (ev) => { if (document.body.classList.contains('pc-modus')) { ev.preventDefault(); document.getElementById('ablage')?.classList.add('drueber'); } });
document.addEventListener('dragleave', () => document.getElementById('ablage')?.classList.remove('drueber'));
document.addEventListener('drop', (ev) => {
  if (!document.body.classList.contains('pc-modus')) return;
  ev.preventDefault();
  if (z.modell && ev.dataTransfer?.files?.length) dateienHinzu([...ev.dataTransfer.files]);
});
