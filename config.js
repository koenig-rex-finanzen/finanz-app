// Öffentliche Einstellungen der App. Hier steht nichts Geheimes.
export const CONFIG = {
  // Anwendungs-ID (Client-ID) aus der Microsoft-Entra-Registrierung „Unser Geld“.
  // Leer = Microsoft-Anmeldung aus, die App öffnet die Datei nur per Auswahl.
  clientId: '',
  dateiName: 'Buchhaltung_Kai_Natalie.xlsx',
  ordnerName: 'Finanz-Tool_König-Rex',
  // Wo die Datei in Kais OneDrive liegt (relativ zu „Eigene Dateien“)
  pfade: ['Documents/Finanz-Tool_König-Rex', 'Dokumente/Finanz-Tool_König-Rex', 'Finanz-Tool_König-Rex'],
};
