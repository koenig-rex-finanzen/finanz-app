// Startet die neue einheitliche Oberfläche, sobald es die neuen Dateien gibt – sonst die bisherige App.
// Die bisherige App wird dann gar nicht geladen (sie hängt eigene Klick-Steuerungen an die Seite).
let neu = null;
try { neu = await import('./neu.js'); } catch (e) { console.error('Neue Oberfläche nicht ladbar:', e); }
if (neu && await neu.aktiv().catch(() => false)) neu.start();
else await import('./app.js');
