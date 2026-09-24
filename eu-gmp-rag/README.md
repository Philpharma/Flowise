# EU-GMP-Leitfaden → Word-Dateien für RAG

Lädt alle PDFs des Abschnitts „EU-GMP Leitfaden“ von
https://www.bundesgesundheitsministerium.de/service/gesetze-und-verordnungen/bekanntmachungen
und erzeugt daraus kompakte Word-Dateien (max. ca. 15 Seiten) mit dem reinen Richtlinientext.

- Entfernt: Deckblatt der Bekanntmachung, Dokumenthistorie-Tabelle, Kopf-/Fußzeilen, Seitenzahlen
- Tabellen werden zeilenweise als Text wiedergegeben („Spalte: Wert; Spalte: Wert“), damit ein RAG sie versteht
- Lange Dokumente (Teil II, Teil IV …) werden an Kapitelgrenzen geteilt, z. B.
  `Teil II Wirkstoffe - Teil 2 von 5 - Kap 3-4 Personal, Gebäude und Anlagen.docx`
- Überschriften sind echte Word-Überschriften; jede Datei beginnt mit Titel und Quellenangabe

## Ergebnis herunterladen
- Ordner `eu-gmp-rag/ausgabe/` bzw. `eu-gmp-rag/ausgabe.zip` in diesem Branch, oder
- GitHub → Actions → „EU-GMP PDFs -> Word (RAG)“ → letzter Lauf → Artefakt `eu-gmp-word`

## Lokal ausführen
```
pip install -r requirements.txt
python eu_gmp_to_docx.py --out ausgabe            # lädt direkt von der BMG-Seite
python eu_gmp_to_docx.py --pdf-dir pdfs --out ausgabe   # oder mit bereits geladenen PDFs
```
