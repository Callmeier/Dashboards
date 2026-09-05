# Microsoft / OneDrive einmalig einrichten

Die PWA schreibt Monatsstände über Microsoft Graph direkt in `03_Monatsstaende` der ausgewählten Excel-Datei.

## 1. App in Microsoft Entra registrieren

1. Microsoft Entra Admin Center öffnen.
2. **App-Registrierungen** → **Neue Registrierung**.
3. Name z. B. `Callmeier Finance Cockpit`.
4. Bei den unterstützten Kontotypen eine Option wählen, die zu deinem Microsoft-Konto passt. Wenn auch ein persönliches Microsoft-Konto verwendet werden soll, muss dieses zugelassen sein.
5. Registrierung abschließen.

## 2. Single-Page-App konfigurieren

1. In der App-Registrierung **Authentifizierung** öffnen.
2. **Plattform hinzufügen** → **Single-Page-Anwendung (SPA)**.
3. Als Umleitungs-URI exakt die URL eintragen, die die PWA unter **Daten erfassen → Diese Umleitungs-URL in Microsoft Entra eintragen** anzeigt.

Es wird **kein Client-Secret** benötigt. Eine Browser-PWA ist ein öffentlicher SPA-Client und verwendet den Authorization-Code-Flow mit PKCE.

## 3. Microsoft Graph-Berechtigung

Unter **API-Berechtigungen** die delegierte Microsoft-Graph-Berechtigung hinzufügen:

- `Files.ReadWrite`

Damit kann die PWA Dateien des angemeldeten Benutzers lesen und die ausgewählte Excel-Datei aktualisieren.

## 4. App-ID in der PWA hinterlegen

1. Auf der Übersichtsseite der App-Registrierung die **Anwendungs-ID (Client)** kopieren.
2. In der PWA **Daten erfassen** öffnen.
3. Die ID in **Microsoft App-ID** eintragen.
4. **App-ID speichern** wählen.
5. **Microsoft verbinden** wählen und Zugriff bestätigen.

Die App-ID wird nur im Browser gespeichert und nicht in GitHub eingecheckt.

## 5. Excel-Datei auswählen

1. Die Excel-Datei muss in deinem OneDrive liegen.
2. In der PWA nach einem Teil des Dateinamens suchen.
3. Datei auswählen.
4. Monat wählen → **Monatsstände laden**.
5. Werte ändern → **Änderungen in Excel speichern**.

Die PWA schreibt nur in die Monatswert-Spalten K bis V des Tabellenblatts `03_Monatsstaende`. Formeln und andere Tabellenblätter werden nicht durch die PWA überschrieben.
