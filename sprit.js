// Feld „Sprit im <Monat>“ in der Übersicht (Gemeinsam): Budget, schon weg, Tendenz, Buchungen.
// Rechnung in kern/rechnen.js (spritBudget). Hier stehen keine persönlichen Daten.
import { eur, monatsname } from './kern/rechnen.js';

const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const datumKurz = (iso) => (iso ? `${Number(iso.slice(8, 10))}.${Number(iso.slice(5, 7))}.` : '');

/** HTML-Abschnitt für die Monatsansicht `v` (aus monatsansicht()). Leer, wenn es keinen Posten „Sprit“ gibt. */
export function spritAbschnitt(v) {
  const sp = v.sprit;
  if (!sp) return '';
  const anteil = sp.budget > 0 ? Math.min(1, sp.ist / sp.budget) : 1;
  const drueber = sp.warnung === 'drueber';
  const kopfText = drueber ? `${eur(sp.ist - sp.budget)} über dem Budget` : `noch ${eur(sp.rest)} übrig`;
  let tendenz = '';
  if (sp.tendenz != null) {
    tendenz = sp.warnung === 'tendenz'
      ? `<p class="sprit-warn"><b>Achtung:</b> Bei eurem bisherigen Tempo landet ihr Ende ${h(v.name)} bei etwa ${eur(sp.tendenz)}, also ${eur(sp.tendenz - sp.budget)} über dem Budget.</p>`
      : `<p class="hint">Tendenz: Ende ${h(v.name)} etwa ${eur(sp.tendenz)}${drueber ? '' : ` (${eur(sp.budget - sp.tendenz)} unter dem Budget)`}.</p>`;
  } else if (sp.laufend && sp.tendenzAb) {
    tendenz = `<p class="hint">Eine Tendenz, wo ihr am Monatsende landet, zeigt die App ab ${h(monatsname(sp.tendenzAb, true))} (dann gibt es zwei volle Monate zum Vergleichen).</p>`;
  }
  const schnitt = sp.schnitt != null ? `<p class="hint">Im Schnitt der letzten ${sp.schnittMonate} Monate: ${eur(sp.schnitt)} im Monat${sp.schnitt > sp.budget + 0.005 ? ` – mehr als das Budget von ${eur(sp.budget)}` : ''}.</p>` : '';
  const liste = sp.buchungen.length ? `<details class="sprit-liste"><summary>${sp.buchungen.length} ${sp.buchungen.length === 1 ? 'Buchung zählt' : 'Buchungen zählen'} dazu</summary>
    ${sp.buchungen.map((b) => `<div class="sprit-zeile"><span>${datumKurz(b.datum)} ${h(b.posten)} <small>${h(b.konto)}</small></span><b>${eur(b.betrag)}</b></div>`).join('')}</details>`
    : '<p class="hint">Noch keine Tank-Buchung in diesem Monat.</p>';
  return `<section><h2>Sprit im ${h(v.name)}</h2>
    <div class="ber sprit ${drueber ? 'erreicht' : sp.warnung === 'tendenz' || anteil >= 0.85 ? 'fast' : ''}">
      <div class="ber-kopf"><span>${eur(sp.ist)} <small>von ${eur(sp.budget)}</small></span><small>${kopfText}</small></div>
      <div class="balken"><span style="width:${Math.round(anteil * 100)}%"></span></div>
      ${tendenz}${schnitt}${liste}
    </div></section>`;
}
