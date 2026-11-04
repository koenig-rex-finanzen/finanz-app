// Öffentliche Einstellungen der App. Hier steht nichts Geheimes.
export const CONFIG = {
  // Anwendungs-ID (Client-ID) aus der Microsoft-Entra-Registrierung „Unser Geld“.
  // Leer = Microsoft-Anmeldung aus, die App öffnet die Datei nur per Auswahl.
  clientId: '3beb2282-42a9-45a8-8ec4-8ab0d690f232',
  dateiName: 'Buchhaltung_Kai_Natalie.xlsx',
  ordnerName: 'Finanz-Tool_König-Rex',
  // Wo die Datei in Kais OneDrive liegt (relativ zu „Eigene Dateien“)
  // Wem die gemeinsame Datei gehört (liegt in seinem OneDrive). Wer sie nur geteilt bekommt, ist die andere Person.
  besitzer: 'Kai',
  pfade: ['Finanz-Tool_König-Rex', 'Documents/Finanz-Tool_König-Rex', 'Dokumente/Finanz-Tool_König-Rex'],
};
