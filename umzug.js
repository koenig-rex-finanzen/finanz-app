// Umzug alt → neu als Seite in der App (#umzug). Nur für den Besitzer der gemeinsamen Datei (Kai).
// 1. Sicherung der aktuellen Dateien  2. neue Dateien bauen  3. Vergleichsliste  4. erst nach Bestätigung anlegen (überschreibt nie).
import * as od from './onedrive.js';
import * as U from './kern/umzug.js';
import * as V from './kern/vorlage.js';
import { leseDatei } from './kern/einheit.js';
import { leseBuchhaltung } from './kern/lesen.js';
import { leseKai } from './kern/kai.js';
import { BEREICH_DATEI } from './kern/format.js';

const $ = (s) => document.querySelector(s);
const h = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const eur = (n) => (n == null ? '–' : (n < 0 ? '−' : '') + Math.abs(n).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');
const heute = () => new Date().toISOString().slice(0, 10);
let log = [];
const schreib = (t) => { log.push(t); const el = $('#u-log'); if (el) el.innerHTML = log.map((x) => `<li>${x}</li>`).join(''); };

export function seite() {
  document.body.dataset.bereich = 'gemeinsam';
  $('#app').innerHTML = `<main class="wrap einheit"><section><h2>Umzug in die neuen Dateien</h2>
    <p class="hint">Baut aus den aktuellen Dateien die vier neuen (Gemeinsam, Tagesgeld, Kai, Natalie). Vorher wird alles gesichert. Die alten Dateien bleiben unverändert als Archiv.</p>
    <div class="zwei"><label class="feld"><span>Endung c24 Haushalt</span><input id="u-e1" inputmode="numeric" maxlength="4"></label><label class="feld"><span>Endung c24 Laufende Kosten</span><input id="u-e2" inputmode="numeric" maxlength="4"></label></div>
    <label class="feld"><span>Korrekturen (JSON, optional)</span><input id="u-korr" placeholder='{"K004":{"faellig":1,"tag":15}}'></label>
    <p style="margin-top:12px"><button class="knopf" id="u-probe">Probelauf (Testordner)</button></p>
    <ol id="u-log" class="offen-liste"></ol><div id="u-ergebnis"></div></section></main>`;
  $('#u-probe').onclick = () => lauf(false);
}

function parameter() {
  const korr = $('#u-korr').value.trim();
  return {
    neueKonten: [
      { id: 'KT20', name: 'c24 Haushalt', bank: 'C24', art: 'Giro', inhaber: 'Natalie', endung: $('#u-e1').value.trim() || null, anfang: null, bemerkung: 'Hauptkonto mit Karte, ersetzt die Visa ab 01.01.2027' },
      { id: 'KT21', name: 'c24 Laufende Kosten', bank: 'C24', art: 'Pocket', inhaber: 'Natalie', endung: $('#u-e2').value.trim() || null, anfang: null, bemerkung: 'Daueraufträge und Abbuchungen ab 01.01.2027' },
    ],
    kostenKorrektur: korr ? JSON.parse(korr) : {},
  };
}

/** Baut alle vier Dateien aus den aktuellen alten Dateien. */
export async function bauen(param) {
  const alt = await od.dateiFinden(); if (!alt) throw new Error('Alte gemeinsame Datei nicht gefunden.');
  const bG = await od.herunterladen(alt.ort);
  const kOrd = await od.privatOrdnerFinden('Kai'); const kf = await od.kaiDateiFinden(kOrd); if (!kf) throw new Error('Kai_Privat.xlsx nicht gefunden.');
  const bK = await od.herunterladen(kf.ort);
  const nOrd = await od.privatOrdnerFinden('Natalie');
  const nJson = nOrd ? (await od.privatLaden(nOrd))?.daten : null;
  const altG = await leseBuchhaltung(bG); const altK = await leseKai(bK);
  const H = heute();
  const g = U.gemeinsamUndTagesgeld(altG, param);
  const zg = await U.zusatzGemeinsam(bG, g.gemeinsam);
  const ka = U.kai(altK); const zk = await U.zusatzKai(bK);
  const n = nJson ? U.natalie(nJson, altG, H) : null;
  // Natalies Topf-Ziele in die Tagesgeld-Datei
  for (const zi of n?.ziele || []) { const t = g.tagesgeld.toepfe.find((x) => x.name === zi.topf); if (t) { t.ziel = zi.ziel; t.zielBis = zi.zielBis; } }
  const kat = (extra) => [...g.gemeinsam.kategorien, ...extra.filter((x) => !g.gemeinsam.kategorien.some((k) => k.name === x)).map((name) => ({ name, aktiv: 'Ja', angelegt: 'Umzug' }))];
  const alleNeu = [...(ka.neueKategorien || []), ...(n?.neueKategorien || [])];
  const dateien = {
    gemeinsam: V.dateiBauen({ ...g.gemeinsam, soll: zg.soll, kategorien: kat(alleNeu), zusatz: zg.zusatz, fremdeStile: zg.fremdeStile, namen: zg.namen, metadata: zg.metadata }),
    tagesgeld: V.dateiBauen({ ...g.tagesgeld, kategorien: kat(alleNeu) }),
    kai: V.dateiBauen({ ...ka.daten, kategorien: kat(alleNeu), zusatz: [...Object.entries(ka.zusatz).map(([x, zz]) => V.zusatzTabelle(x, zz)), ...zk.zusatz], fremdeStile: zk.fremdeStile, metadata: zk.metadata }),
    ...(n ? { natalie: V.dateiBauen({ ...n.daten, kategorien: kat(alleNeu), zusatz: Object.entries(n.zusatz).map(([x, zz]) => V.zusatzTabelle(x, zz)) }) } : {}),
  };
  const neu = {}; for (const [k, b] of Object.entries(dateien)) neu[k] = await leseDatei(b);
  const vergleich = {
    'Gemeinsam + Tagesgeld': U.vergleichGemeinsam(altG, neu.gemeinsam, neu.tagesgeld, H),
    Kai: U.vergleichKai(altK, neu.kai, H),
    ...(n ? { Natalie: U.vergleichNatalie(nJson, altG, neu.natalie, H) } : {}),
  };
  return { dateien, vergleich, hinweise: [...g.hinweise, ...zg.hinweise, ...ka.hinweise, ...zk.hinweise, ...(n?.hinweise || [])], quellen: { alt, kf, nOrd } };
}

async function lauf(echt) {
  log = []; $('#u-ergebnis').innerHTML = '';
  try {
    const param = parameter();
    schreib('Lese die aktuellen Dateien …');
    const r = await bauen(param);
    schreib(`Neue Dateien gebaut: ${Object.keys(r.dateien).join(', ')}.`);
    const unerklaert = Object.values(r.vergleich).flat().filter((z) => !z.gleich && !z.erklaerung);
    $('#u-ergebnis').innerHTML = `<h2>Vergleich alt ↔ neu</h2>${Object.entries(r.vergleich).map(([t, l]) => `<h3>${h(t)}: ${l.filter((z) => z.gleich).length}/${l.length} gleich</h3>
      <div class="konten">${l.filter((z) => !z.gleich).map((z) => `<div class="konto"><span class="dot" style="background:${z.erklaerung ? 'var(--gem)' : 'var(--warn)'}"></span><div><div class="name">${h(z.was)}</div><div class="sub">${h(z.erklaerung || 'nicht erklärt')}</div></div><div class="betrag">${eur(z.alt)} → ${eur(z.neu)}</div></div>`).join('')}</div>`).join('')}
      ${r.hinweise.length ? `<h3>Hinweise</h3><ul class="offen-liste">${r.hinweise.map((x) => `<li>${h(x)}</li>`).join('')}</ul>` : ''}
      ${!echt ? (unerklaert.length ? `<div class="hinweis warn">${unerklaert.length} Abweichung(en) ohne Erklärung – der echte Umzug ist gesperrt.</div>`
        : `<p style="margin-top:16px">Alles gleich oder erklärt. Für den echten Umzug „UMZUG“ eintippen:</p><input id="u-ok" class="suche" autocomplete="off"><p><button class="knopf" id="u-echt">Echten Umzug starten</button></p>`) : ''}`;
    if (!echt) {
      await hochladen(r, false);
      const b = $('#u-echt'); if (b) b.onclick = () => { if ($('#u-ok').value.trim() === 'UMZUG') lauf(true); else schreib('Bitte „UMZUG“ eintippen.'); };
    } else {
      if (unerklaert.length) throw new Error('Unerklärte Abweichungen – abgebrochen.');
      schreib('Sichere die alten Dateien …');
      await od.sicherungAnlegen(r.quellen.alt.ort, 'vor-Umzug'); await od.sicherungAnlegen(r.quellen.kf.ort, 'vor-Umzug');
      await hochladen(r, true);
      try { localStorage.setItem('ug.neu.aktiv', '1'); } catch {}
      schreib('<b>Fertig.</b> Die App arbeitet ab jetzt mit den neuen Dateien. <a href="./">Zur App</a>');
    }
  } catch (e) { schreib(`<span class="fehler">${h(e.message)}</span>`); }
}

async function hochladen(r, echt) {
  const ordner = {};
  if (echt) {
    const it = r.quellen.alt.item; ordner.gemeinsam = ordner.tagesgeld = { driveId: it.parentReference.driveId, itemId: it.parentReference.id };
    ordner.kai = await od.privatOrdnerFinden('Kai'); ordner.natalie = r.quellen.nOrd;
  } else {
    ordner.gemeinsam = ordner.tagesgeld = await od.ordnerAusPfad('Umbau-Test');
    ordner.kai = await od.ordnerAusPfad('Apps/Kai/Umbau-Test'); ordner.natalie = await od.ordnerAusPfad('Apps/Natalie/Umbau-Test');
  }
  for (const [b, bytes] of Object.entries(r.dateien)) {
    const name = (echt ? '' : 'TEST_') + BEREICH_DATEI[b].datei;
    if (!ordner[b]) { schreib(`Ordner für ${b} fehlt – übersprungen.`); continue; }
    if (!echt) { const alt = await od.graphRoh(`/drives/${ordner[b].driveId}/items/${ordner[b].itemId}:/${encodeURIComponent(name)}?$select=id`); if (alt?.id) await od.dateiErsetzen({ driveId: ordner[b].driveId, itemId: alt.id }, bytes); else await od.dateiNeu(ordner[b], name, bytes); }
    else await od.dateiNeu(ordner[b], name, bytes);
    schreib(`${name} angelegt.`);
  }
  // Einmal in Excel rechnen lassen, damit alle Formel-Ergebnisse in den Dateien stehen
  for (const b of Object.keys(r.dateien)) {
    const name = (echt ? '' : 'TEST_') + BEREICH_DATEI[b].datei;
    const it = ordner[b] && await od.graphRoh(`/drives/${ordner[b].driveId}/items/${ordner[b].itemId}:/${encodeURIComponent(name)}?$select=id`);
    if (it?.id) { await od.berechnen({ driveId: ordner[b].driveId, itemId: it.id }); schreib(`${name} in Excel gerechnet.`); }
  }
}
