# Dropbox-Einrichtung – Finance Cockpit

Die PWA verwendet Dropbox als Speicher für die Excel-Datei. Es ist kein Microsoft 365, Entra oder eigener Server erforderlich.

## 1. Dropbox-App einmalig anlegen

1. Öffne die Dropbox App Console: https://www.dropbox.com/developers/apps
2. Wähle **Create app**.
3. API: **Scoped access**.
4. Zugriff: **App folder** (empfohlen).
5. Vergib einen Namen, z. B. `Callmeier Finance Cockpit`.

Mit **App folder** kann die PWA ausschließlich ihren eigenen Ordner unter `Apps/...` sehen und nicht deine übrige Dropbox.

## 2. Berechtigungen

Im Tab **Permissions** aktivieren:

- `files.metadata.read`
- `files.content.read`
- `files.content.write`

Danach Änderungen speichern/übernehmen.

## 3. Umleitungs-URL

Im Finance Cockpit auf **Daten erfassen** klicken. Dort wird unter „Diese Umleitungs-URL in Dropbox eintragen“ die exakte URL angezeigt.

Diese URL in der Dropbox App Console unter **OAuth 2 → Redirect URIs** eintragen.

## 4. App-Key

In der Dropbox App Console unter **Settings** den **App key** kopieren und im Finance Cockpit in das Feld **Dropbox App-Key** eintragen.

Kein App secret in die PWA eintragen. Die Browser-App verwendet OAuth 2 mit PKCE.

## 5. Excel-Datei

Nach dem Verbinden kannst du im Finance Cockpit auf **Aktuelle Excel hochladen** klicken. Die Datei wird in den geschützten Dropbox-App-Ordner gelegt. Danach kann die PWA Monatsstände lesen und aktualisieren.

## Technischer Schutz

- Finanzdaten liegen nicht in GitHub.
- Das Zugriffstoken wird nur lokal im Browser gespeichert und ist kurzlebig.
- Vor dem Überschreiben prüft die PWA die Dropbox-Dateirevision, damit zwischenzeitliche Änderungen nicht unbemerkt überschrieben werden.
- Beim Speichern werden gezielt die Monatszellen in `03_Monatsstaende` geändert; Formeln, Formatierungen, Diagramme und sonstige Bestandteile der XLSX/XLSM-Datei bleiben im ZIP-Paket erhalten.
- Die berechneten Werte in `04_Monatsauswertung` werden als Cache mitgeführt, damit die PWA sofort aktualisiert ist. Excel selbst wird beim nächsten Öffnen zu einer vollständigen Neuberechnung aufgefordert.
