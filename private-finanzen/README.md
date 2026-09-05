# Dashboard Private Finanzen

Eigenständige PWA für das **Callmeier Finance Cockpit**.

## Funktionen

- lokaler Excel-Import im Browser
- responsive KPI- und Vermögensübersicht
- Monatswahl, Jahresverlauf und Asset-Allokation
- automatische Hinweise anhand der Excel-Parameter
- lokale Zwischenspeicherung auf dem Gerät
- PWA-Installation und Offline-Grundfunktion
- **Dropbox-Synchronisierung für Monatsstände**

## Dropbox-Modus

Die PWA kann eine `.xlsx`- oder `.xlsm`-Datei aus einem Dropbox-App-Ordner laden. Unter **Daten erfassen** können die Werte für `03_Monatsstaende` eingegeben und anschließend wieder in dieselbe Dropbox-Datei gespeichert werden.

Dropbox dient als Dateispeicher; die PWA berechnet die für das Dashboard benötigten Kennzahlen direkt im Browser und aktualisiert zusätzlich die Excel-Berechnungscaches. Excel wird beim nächsten Öffnen zur vollständigen Neuberechnung aufgefordert.

Die einmalige Dropbox-Einrichtung ist in [`DROPBOX_SETUP.md`](./DROPBOX_SETUP.md) beschrieben.

## Datenschutz

**Keine echten Finanzdaten gehören in dieses GitHub-Repository.**

Die Excel-Datei liegt entweder lokal auf dem Gerät oder im vom Nutzer autorisierten Dropbox-App-Ordner. Bei empfohlenem **App Folder**-Zugriff hat die PWA keinen Zugriff auf die übrigen Dropbox-Dateien.
