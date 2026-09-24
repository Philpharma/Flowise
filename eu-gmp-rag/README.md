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

## Zusatzdokumente (`quellen.json`)
Dokumente, die das BMG nicht (oder nur überholt) auf Deutsch bekanntgemacht hat, werden über
`quellen_laden.py` geladen und mit demselben Skript konvertiert:
- **Deutsch (amtlich)** aus dem Amtsblatt der EU (EUR-Lex bzw. Cellar des Amts für Veröffentlichungen):
  RL (EU) 2017/1572, DVO (EU) 2017/1569, VO (EU) 2019/6, DVO (EU) 2025/2091 und 2025/2154,
  Leitlinien Hilfsstoffe (2015/C 95/02), GDP-Leitlinien (2013/C 343/01, 2015/C 95/01)
- **Englisch** (nur englisch verfügbar, Dateiname endet auf „(EN)“): Teil III, Anhänge 1, 2, 17, 21,
  neue Fassungen von Anhang 13 (Leitlinie 2017) und Anhang 19 (2026), Glossar, Korrespondenztabellen

Pro Quelle lässt sich in `quellen.json` festlegen, wo der Text beginnt (`start`, regulärer Ausdruck)
und ob Zeilennummern entfernt werden (`zeilennummern`).

```
python quellen_laden.py quellen.json ausgabe/_quellen
python eu_gmp_to_docx.py --out ausgabe --quellen quellen.json            # BMG + Zusatzdokumente
python eu_gmp_to_docx.py --pdf-dir - --out ausgabe --quellen quellen.json  # nur Zusatzdokumente
```
