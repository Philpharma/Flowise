# EU-GMP-Leitfaden → Word-Dateien für RAG

Lädt alle PDFs des Abschnitts „EU-GMP Leitfaden“ von
https://www.bundesgesundheitsministerium.de/service/gesetze-und-verordnungen/bekanntmachungen
und erzeugt daraus kompakte Word-Dateien (max. ca. 6.500 Wörter ≈ 13–15 Seiten) mit dem reinen Richtlinientext.

- Entfernt: Deckblatt der Bekanntmachung, Status/Dokumenthistorie, Inhaltsverzeichnis, Kopf-/Fußzeilen
  (inkl. Bundesanzeiger-Zeilen), Seitenzahlen – der Text beginnt bei „Grundsätze“/„Einleitung“/„Anwendungsbereich“
- Tabellen werden zeilenweise als Text wiedergegeben („Spalte: Wert; Spalte: Wert“), damit ein RAG sie versteht;
  grau hinterlegte Zellen (z. B. Teil II Tabelle 1, Anhang 7) werden als „(grau markiert)“ ausgegeben;
  Tabellen ohne vollständigen Rahmen (z. B. Reinraumklassen in Teil IV) bleiben Zeile für Zeile erhalten
- Fußnoten werden hinter den Absatz gestellt, statt ihn zu unterbrechen
- Lange Dokumente (Teil II, Teil IV …) werden an Kapitelgrenzen geteilt, z. B.
  `Teil II Wirkstoffe (1 von 3) Kap 1-6 Einleitung, Qualitätsmanagement, Personal.docx`
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
