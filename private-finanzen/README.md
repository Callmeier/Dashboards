# Dashboard Private Finanzen

Installierbare PWA für das Excel-basierte **Callmeier Finance Cockpit**.

## Funktionsumfang

- Excel-Datei wird direkt im Browser eingelesen
- keine Finanzdaten werden an GitHub oder den Webserver übertragen
- letzte importierte Excel-Datei wird lokal per IndexedDB auf dem Gerät gespeichert
- Monatsauswahl Januar bis Dezember
- Nettovermögen, MoM, YTD und Finanzscore
- Liquidität, Rücklagen, Investments und Verbindlichkeiten
- Jahresverlauf als SVG-Chart
- Asset-Allokation
- automatische Hinweise anhand der Parameter aus `05_Ziele_Parameter`
- Zusammenfassung regelmäßiger Zahlungen
- installierbar als PWA
- Offline-App-Shell per Service Worker

## Erwartete Excel-Struktur

Die App liest insbesondere:

- `04_Monatsauswertung`
- `05_Ziele_Parameter`
- `11_Regelmaessige_Zahlungen` (optional)

Die Formeln werden weiterhin in Excel berechnet. Nach Änderungen an der Datei sollte sie in Excel gespeichert werden, bevor sie erneut in die PWA importiert wird.

## Cloudflare Pages

Dieses Dashboard kann als eigenes Cloudflare-Pages-Projekt aus dem Monorepo veröffentlicht werden.

- Repository: `Callmeier/Dashboards`
- Production branch: `main`
- Root directory: `private-finanzen`
- Build command: leer
- Build output directory: `.`

## Datenschutz

Keine `.xlsx`, `.xls`, `.csv` oder Datenbankdateien mit echten Finanzdaten in GitHub ablegen. Die `.gitignore` im Repository schützt zusätzlich vor versehentlichen Uploads solcher Dateien.
