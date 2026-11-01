// Microsoft-Anmeldung für eine Web-App ohne Server (OAuth 2.0 mit PKCE), ohne Fremdbibliothek.
// Nur private Microsoft-Konten („consumers“). Es gibt kein Geheimnis im Code:
// die Client-ID ist öffentlich, die Sicherheit kommt aus PKCE und der festen Rücksprungadresse.

import { CONFIG } from './config.js';

const BASIS = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite.All';
const SPEICHER = 'ug.token';

export const aktiv = () => Boolean(CONFIG.clientId);
const rueckAdresse = () => location.origin + location.pathname;

const b64url = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const zufall = (n = 32) => b64url(crypto.getRandomValues(new Uint8Array(n)));
const lesen = (k, s = localStorage) => { try { return JSON.parse(s.getItem(k)); } catch { return null; } };
const schreiben = (k, v, s = localStorage) => { try { v == null ? s.removeItem(k) : s.setItem(k, JSON.stringify(v)); } catch {} };

/** Leitet zur Microsoft-Anmeldung weiter. */
export async function anmelden({ konto } = {}) {
  const verifier = zufall(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = zufall(16);
  schreiben('ug.pkce', { verifier, state }, sessionStorage);
  const p = new URLSearchParams({
    client_id: CONFIG.clientId, response_type: 'code', redirect_uri: rueckAdresse(), response_mode: 'query',
    scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state,
  });
  if (konto) p.set('login_hint', konto); else p.set('prompt', 'select_account');
  location.assign(`${BASIS}/authorize?${p}`);
}

async function tokenAnfrage(felder) {
  const r = await fetch(`${BASIS}/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CONFIG.clientId, scope: SCOPES, ...felder }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error_description?.split('\r\n')[0] || j.error || `Anmeldung fehlgeschlagen (${r.status})`);
  const t = { access: j.access_token, refresh: j.refresh_token, ablauf: Date.now() + (j.expires_in - 120) * 1000,
    konto: lesen(SPEICHER)?.konto || null };
  schreiben(SPEICHER, t);
  return t;
}

/**
 * Nach der Rückkehr von Microsoft: Code gegen Token tauschen. Gibt true zurück, wenn gerade eine
 * Anmeldung abgeschlossen wurde. Wirft bei Fehlern.
 */
export async function rueckkehrVerarbeiten() {
  const q = new URLSearchParams(location.search);
  if (!q.has('code') && !q.has('error')) return false;
  history.replaceState(null, '', rueckAdresse());
  if (q.has('error')) {
    if (q.get('error') === 'access_denied') throw new Error('Die Anmeldung wurde abgebrochen.');
    throw new Error(q.get('error_description') || q.get('error'));
  }
  const pkce = lesen('ug.pkce', sessionStorage);
  schreiben('ug.pkce', null, sessionStorage);
  if (!pkce || pkce.state !== q.get('state')) throw new Error('Anmeldung ungültig (Sicherheitsprüfung). Bitte noch einmal anmelden.');
  await tokenAnfrage({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: rueckAdresse(), code_verifier: pkce.verifier });
  return true;
}

/** Gültiger Zugriffstoken oder null (dann muss neu angemeldet werden). */
export async function token() {
  const t = lesen(SPEICHER);
  if (!t) return null;
  if (t.access && t.ablauf > Date.now()) return t.access;
  if (!t.refresh) return null;
  try { return (await tokenAnfrage({ grant_type: 'refresh_token', refresh_token: t.refresh })).access; }
  catch { const k = t.konto; schreiben(SPEICHER, k ? { konto: k } : null); return null; }
}

export function kontoMerken(konto) { const t = lesen(SPEICHER) || {}; t.konto = konto; schreiben(SPEICHER, t); }
export const gemerktesKonto = () => lesen(SPEICHER)?.konto || null;
export const angemeldet = () => Boolean(lesen(SPEICHER)?.refresh);
export function abmelden() { schreiben(SPEICHER, null); schreiben('ug.ort', null); }
