// Kleine Fragezeichen mit einfachen Erklärungen – für alle Bereiche (Gemeinsam, Natalie, Kai).
// Funktioniert ohne Änderungen an den einzelnen Seiten: Nach jedem Zeichnen sucht die Datei bekannte
// Feld-Namen, Überschriften und Rechenzeilen und hängt ein „?“ an. Antippen zeigt eine Sprechblase.
// Hier stehen keine persönlichen Daten.

// Schlüssel = sichtbarer Text (ohne <small>). Endet ein Schlüssel auf „…“, reicht der Anfang.
const GEMEINSAM = {
  'Art': '<b>Ausgabe:</b> Geld geht an jemand anderen (Laden, Firma).<br><b>Einnahme:</b> Geld kommt von außen (Gehalt, Kindergeld).<br><b>Umbuchung:</b> Geld wandert zwischen unseren eigenen Konten, z. B. aufs Tagesgeld.',
  'Wofür': 'Ein kurzer Name für die Zahlung, z. B. „Lidl“. Passt er zu einem festen Posten aus der Liste, wird er dort als bezahlt gezählt.',
  'Bezahlt am': 'Der Tag, an dem das Geld vom Konto gegangen ist (wie auf dem Kontoauszug).',
  'Gilt für': 'Der Monat, zu dem die Zahlung gehört. Beispiel: Das Gehalt kommt am 30. Oktober, gilt aber für November.',
  'Von': 'Von welchem Konto das Geld kommt. „Extern“ heißt: von außerhalb, z. B. Arbeitgeber.',
  'Nach': 'Auf welches Konto das Geld geht. „Extern“ heißt: an jemand anderen, z. B. einen Laden.',
  'Bereich': 'Grobe Einordnung, z. B. Haushalt oder Kinder. Damit die Auswertung weiß, wofür das Geld war.',
  'Topf': 'Nur nötig, wenn Geld aufs Tagesgeld geht oder von dort kommt. Dann steigt oder sinkt der Stand dieses Topfs.',
  'Gezahlt von': 'Nur ausfüllen, wenn das Geld nicht vom eigenen Konto kommt. Dann zählt die Einzahlung für diese Person.',
  'Ausgleich, nicht anrechnen': 'Die Buchung ist echt und zählt beim Kontostand und bei den Töpfen. Sie zählt aber <b>nicht</b> als Einzahlung und nicht bei Einnahmen und Ausgaben. Zum Beispiel, wenn Geld nur weitergeleitet oder falsch bezahlt und zurücküberwiesen wurde.',
  'Wo das Geld…': 'Die Stände der gemeinsamen Konten, so wie die App sie aus den Buchungen ausrechnet.',
  'Wer überweist was': 'Die gemeinsamen Kosten werden nach dem Anteil am Einkommen getragen. Hier steht, wer wie viel wohin überweisen soll und was schon angekommen ist.',
  'Töpfe auf dem Tagesgeld': 'Das Geld liegt zusammen auf einem Tagesgeld. Die Töpfe teilen es gedanklich auf, z. B. für Auto oder Geschenke. Jeder Topf hat eine eigene Sparrate.',
  'Noch zu klären': 'Punkte aus dem Blatt „Offene Punkte“, die noch nicht erledigt sind.',
  'Sprit im…': 'Budget = Posten „Sprit“ im Blatt Kosten. Dazu zählt jede Ausgabe, deren „Wofür“ oder Bank-Text nach Tanken aussieht (z. B. Sprit, Tanken, Aral, Shell, Esso), im Monat „Gilt für“. Ab zwei vollen Monaten rechnet die App hoch, wo ihr am Monatsende landet: was schon weg ist + euer Durchschnitt für die restlichen Tage.',
};

const NATALIE = {
  'Art': '<b>Ausgabe:</b> Geld geht von deinem Giro oder PayPal weg.<br><b>Einnahme:</b> Geld kommt auf dein Giro.',
  'Bezahlt mit': 'Womit du bezahlt hast. PayPal wird einzeln eingetragen und geht am Ende vom Giro ab.',
  'Wofür': 'Ein kurzer Name für die Zahlung, z. B. „Friseur“.',
  'Bereich': 'Grobe Einordnung, z. B. Kleidung. Damit siehst du, wofür du wie viel ausgibst, und die Limits greifen.',
  'Gehört zu': 'Wähle hier den geplanten Posten, wenn die Zahlung schon bei den festen Kosten steht. Dann wird sie nicht doppelt abgezogen.',
  'Aus einem Topf bezahlt?': 'Wenn du etwas mit Geld aus einem deiner Töpfe bezahlt hast. Dann zählt es nicht gegen dein „frei“, weil das Geld dafür schon zurückgelegt war.',
  'So rechnet die App': 'Einnahmen − Überweisungen an die gemeinsamen Konten − feste Kosten − was du schon ausgegeben hast = was dir in diesem Monat noch frei bleibt.',
  'Bleibt frei': 'So viel kannst du in diesem Monat noch ausgeben, ohne dass es für die festen Kosten und Überweisungen eng wird.',
  'Überweisungen an gemeinsame Konten': 'Was du laut Aufteilung aufs Gemeinschaftskonto und aufs Tagesgeld überweist. Kommt aus der gemeinsamen Datei.',
  'Feste Kosten': 'Regelmäßige Zahlungen von deinem Giro, z. B. Versicherungen oder Abos.',
  'Schon ausgegeben': 'Alles, was du in diesem Monat eingetragen hast, ohne das, was aus Töpfen bezahlt wurde.',
  'Deine Konten': 'Dein Giro (mit PayPal) und das Tagesgeld mit den Töpfen.',
  'Was noch kommt': 'Geplante Einnahmen und Zahlungen dieses Monats. Daneben siehst du, ob sie schon da oder bezahlt sind.',
  'Ausgaben nach Bereich': 'Deine Ausgaben, nach Bereich zusammengezählt. Der Balken zeigt, wie viel vom Limit schon weg ist.',
  'Deine Töpfe': 'Deine Töpfe auf dem Tagesgeld. Sie stehen in der gemeinsamen Datei, deshalb sieht man dort denselben Stand.',
  'Limit im Monat €': 'So viel möchtest du in diesem Bereich höchstens ausgeben. Die App sagt Bescheid, wenn du es erreichst.',
  'Stand Giro laut Bank in €': 'Der Kontostand laut Bank am Ende des gewählten Tags. Ab da rechnet die App mit deinen Buchungen weiter.',
};

const KAI = {
  'Art': '<b>Einnahme:</b> Geld kommt aufs Konto.<br><b>Ausgabe:</b> Geld geht vom Konto weg. Mit Topf: aus diesem Topf bezahlt.<br><b>In Topf:</b> extra Geld in einen Topf legen, zusätzlich zur Sparrate. Es bewegt sich kein Geld.<br><b>Aus Topf:</b> Geld aus einem Topf zurück ins „Frei“, ohne es auszugeben.<br><b>Umbuchung:</b> echtes Geld wandert zwischen c24 und Spartopf. Die Töpfe ändern sich nicht.',
  'Konto': 'Auf welchem deiner Konten die Buchung passiert.',
  'Von Konto': 'Von hier geht das Geld weg.',
  'Nach Konto': 'Hier kommt das Geld an.',
  'Posten': 'Der geplante Posten (feste Kosten oder Einnahme), zu dem die Buchung gehört. Dann weiß die Vorschau, dass er schon bezahlt ist. „Variabel“ ist alles Ungeplante.',
  'Topf': 'Nur bei „In Topf“, „Aus Topf“ oder einer Ausgabe, die du aus einem Topf bezahlst. Sonst „kein Topf“.',
  'Gilt für': 'Der Monat, zu dem die Buchung gehört. Das Gehalt ab dem 20. gilt für den Folgemonat.',
  'Gegenüber': 'Wer das Geld bekommt oder schickt, z. B. „Aral“.',
  'Frei': 'c24 + Spartopf − alles, was in Töpfen verplant ist. Das ist Geld, das für nichts reserviert ist.',
  'Bis Monatsende behalten': 'So viel sollte Ende des Monats noch übrig sein, damit die nächsten Monate ohne Minus auskommen.',
  'Davon noch in die Rücklage': 'Der Teil von „behalten“, der noch nicht im Topf „Rücklage Lücke“ liegt. Mit dem Knopf legst du ihn dort zurück (das Geld bleibt auf c24).',
  'Frei für Variables diesen Monat': 'So viel kannst du diesen Monat für Ungeplantes ausgeben (Einkauf, Essen, Spontanes), ohne dass später eine Lücke entsteht.',
  'Die nächsten…': 'Die Vorschau ohne variable Ausgaben: Feste Einnahmen, Überweisung aufs Gemeinschaftskonto, feste Kosten und Sparraten.',
  'Vorschau ab…': 'Zwölf Monate im Voraus. Gewarnt wird für die ersten Monate, sobald irgendwo auch nur 1 Cent fehlt.',
  'Plan und Ist im…': 'Plan = was laut festen Kosten und Einnahmen kommen sollte. Ist = was du schon gebucht hast.',
  'Empfehlung für…': 'Was die App dir für diesen Monat rät, damit die nächsten Monate gedeckt sind.',
  'Frei am Ende': 'So viel ist Ende des Monats frei, wenn du nichts Variables ausgibst.',
  'Behalten': 'So viel musst du von diesem Monat übrig lassen, damit ein späterer Monat nicht ins Minus rutscht.',
  'Variables': 'So viel ist in diesem Monat für Ungeplantes da.',
  'Auf dem…': 'Diese Töpfe liegen als echtes Geld auf dem Spartopf (Pocket). Die Summe sollte so groß sein wie der Stand vom Spartopf.',
  'Auf c24 (zählt als frei)': 'Diese Töpfe bleiben auf dem c24-Konto. Sie zählen beim „Frei“ mit, sind aber für einen Zweck markiert.',
  'So setzt sich der Stand zusammen': 'Anfangsbestand + alle Sparraten seit Beginn + extra hineingelegt − entnommen oder ausgegeben = Stand.',
  'Anfangsbestand': 'So viel war am Stichtag schon im Topf.',
  'Sparraten zugeteilt': 'Jeden Monat kommt die Sparrate automatisch dazu. Dafür musst du nichts eintragen.',
  'Hineingelegt / entnommen': 'Buchungen mit diesem Topf: „In Topf“ (+), „Aus Topf“ und Ausgaben aus dem Topf (−).',
  'Zählt als frei?': '„Ja“ nur für Töpfe, die auf c24 bleiben, z. B. „Rücklage Lücke“. „Nein“ für Töpfe auf dem Spartopf.',
  'Sparrate im Monat €': 'So viel kommt jeden Monat automatisch in den Topf.',
  'spartopf': 'Das Pocket bei c24. Hier liegt das Geld deiner Töpfe. Steht „noch … überweisen“, ist auf dem Pocket weniger als in den Töpfen verplant. Dann buchst du den Betrag mit einer Umbuchung um.',
};

const TEXTE = { gemeinsam: GEMEINSAM, natalie: NATALIE, kai: KAI };
const ZIELE = 'label.feld > span, main h2, .rz > span, .k-hero-oben > span:first-child, .k-tz.kopf > span, .haken b, [data-hilfe-schluessel]';

function bereichVon(el) {
  const m = el.closest('main') || document.querySelector('main');
  if (m?.classList.contains('kai')) return 'kai';
  if (m?.classList.contains('meins')) return 'natalie';
  return 'gemeinsam';
}
const eigenerText = (el) => [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join('').replace(/\s+/g, ' ').trim();
function finde(liste, text) {
  if (!text) return null;
  if (liste[text]) return liste[text];
  const k = Object.keys(liste).find((x) => x.endsWith('…') && text.startsWith(x.slice(0, -1)));
  return k ? liste[k] : null;
}
const knopf = (schluessel) => `<span class="hilfe" role="button" tabindex="0" aria-label="Erklärung" data-hilfe="${schluessel}">?</span>`;

let laufend = false;
function schmuecken() {
  laufend = false;
  // „Art“-Auswahl (Kacheln) bekommt eine kleine Zeile darüber
  document.querySelectorAll('.seg').forEach((seg) => {
    if (seg.previousElementSibling?.classList.contains('hilfe-ueber')) return;
    const name = seg.getAttribute('aria-label') || (seg.querySelector('input[name=art]') ? 'Art' : null);
    const b = bereichVon(seg.closest('dialog') ? document.body : seg);
    if (!name || !finde(TEXTE[b], name)) return;
    seg.insertAdjacentHTML('beforebegin', `<div class="hilfe-ueber" data-bereich-hilfe="${b}">${name} ${knopf(name)}</div>`);
  });
  document.querySelectorAll(ZIELE).forEach((el) => {
    if (el.dataset.hilfeFertig) return;
    el.dataset.hilfeFertig = '1';
    const b = bereichVon(el.closest('dialog') ? document.body : el);
    const schluessel = el.dataset.hilfeSchluessel || eigenerText(el);
    if (!finde(TEXTE[b], schluessel)) return;
    const small = el.querySelector(':scope > small');
    const html = knopf(schluessel);
    if (small) small.insertAdjacentHTML('beforebegin', ' ' + html); else el.insertAdjacentHTML('beforeend', ' ' + html);
    el.querySelector(':scope > .hilfe').dataset.bereichHilfe = b;
  });
}

let blase = null;
function zu() { if (blase) { blase.remove(); blase = null; } }
function auf(t) {
  const alt = blase?.dataset.fuer === t.dataset.hilfe;
  zu();
  if (alt) return;
  const text = finde(TEXTE[t.dataset.bereichHilfe || t.closest('[data-bereich-hilfe]')?.dataset.bereichHilfe || bereichVon(t)], t.dataset.hilfe);
  if (!text) return;
  blase = document.createElement('div');
  blase.className = 'hilfe-blase'; blase.setAttribute('role', 'tooltip'); blase.dataset.fuer = t.dataset.hilfe;
  blase.innerHTML = text;
  (t.closest('dialog') || document.body).appendChild(blase);
  const r = t.getBoundingClientRect();
  const w = Math.min(320, window.innerWidth - 32);
  blase.style.width = w + 'px';
  blase.style.left = Math.max(16, Math.min(r.left + r.width / 2 - w / 2, window.innerWidth - w - 16)) + 'px';
  const h = blase.offsetHeight;
  blase.style.top = (r.bottom + 8 + h < window.innerHeight ? r.bottom + 8 : Math.max(8, r.top - h - 8)) + 'px';
}

export function starten() {
  document.addEventListener('click', (ev) => {
    const t = ev.target.closest('.hilfe');
    if (t) { ev.preventDefault(); ev.stopPropagation(); auf(t); return; }
    if (blase && !ev.target.closest('.hilfe-blase')) zu();
  }, true);
  document.addEventListener('keydown', (ev) => {
    const t = ev.target.closest?.('.hilfe');
    if (t && (ev.key === 'Enter' || ev.key === ' ')) { ev.preventDefault(); auf(t); }
    if (ev.key === 'Escape') zu();
  }, true);
  window.addEventListener('scroll', zu, { passive: true, capture: true });
  new MutationObserver(() => { if (!laufend) { laufend = true; requestAnimationFrame(schmuecken); } })
    .observe(document.body, { childList: true, subtree: true });
  schmuecken();
}
