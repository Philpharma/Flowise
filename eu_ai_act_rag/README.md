# RAG-Dokumentenbestand EU-KI-Verordnung (VO (EU) 2024/1689)

Python-Pipeline, die die amtlichen Texte der KI-Verordnung von EUR-Lex abruft, deterministisch
(ohne LLM) gliedert und als RAG-taugliche Word-Dateien mit Metadaten und automatischer
Qualitätsprüfung ausgibt.

## Quellen

| Verwendung | Dokument | CELEX |
|---|---|---|
| Artikel, Kapitel, Abschnitte, Anhänge (aktuell geltend) | VO (EU) 2024/1689, konsolidierte Fassung, Stand 27.07.2026 | `02024R1689-20260727` |
| Erwägungsgründe 1–180, Bezugsvermerke | VO (EU) 2024/1689, Amtsblattfassung | `32024R1689` |
| Erwägungsgründe und verfügender Teil der Änderungsverordnung | VO (EU) 2026/1744 (Digital-Omnibus KI) | `32026R1744` |
| Leitlinien der Kommission (getrennt, kein Gesetzestext) | Kommissionsseiten auf `*.europa.eu`, siehe `leitlinien_quellen.json` | – |

Gesetzestext wird ausschließlich von `eur-lex.europa.eu` geladen (Rückfall: Cellar des Amts für
Veröffentlichungen, `publications.europa.eu`, das EUR-Lex technisch speist). Andere Hosts werden abgewiesen.

## Installation und Aufruf

```bash
pip install -r requirements.txt
python build_corpus.py                       # abrufen, erzeugen, prüfen
python build_corpus.py --entwuerfe           # zusätzlich Leitlinien-ENTWÜRFE (gekennzeichnet)
python build_corpus.py --leitlinien-discovery  # Liste weiterer Leitlinien-Kandidaten zur manuellen Prüfung
python build_corpus.py --seitenzahl          # echte Seitenzahlen per LibreOffice (optional)
```

Ergebnis: `output/EU_KI_Verordnung_RAG/` (anpassbar mit `--out`). Exit-Code 0 = alle Prüfungen bestanden,
1 = Qualitätsfehler (siehe `00_Qualitaetsbericht.md`), 2 = Quelle nicht abrufbar.

**Ohne direkten Zugriff auf EUR-Lex** (Firewall, Bot-Schutz): die drei Dokumente im Browser über
`https://eur-lex.europa.eu/legal-content/DE/TXT/HTML/?uri=CELEX:<CELEX>` öffnen, als
`<CELEX>.html` speichern (z. B. `02024R1689-20260727.html`) und aufrufen mit
`python build_corpus.py --html-dir <ordner>`. Leitlinien-PDFs analog mit `--leitlinien-pdf-dir`
(Dateiname `<id>.pdf`, ids aus `leitlinien_quellen.json`).

## Ausgabestruktur

```
00_Dokumentenuebersicht.xlsx / .csv   Dateiname, Typ, Kapitel/Abschnitt, Artikel von/bis, EG von/bis,
                                      Anhang, CELEX, Fassungsstand, Seiten, QS-Status
00_Qualitaetsbericht.md / .json       alle Einzelprüfungen
01_EU_KI_Verordnung/                  NN_Kapitel_<röm>[_Abschnitt_<n>]_<Titel>_Art_<von>-<bis>.docx
02_Erwaegungsgruende/VO_2024-1689_Originalfassung/   Erwaegungsgruende_001-0xx.docx …
02_Erwaegungsgruende/VO_2026-1744_Digital-Omnibus/   VO_2026-1744_Erwaegungsgruende_…docx
03_Anhaenge/                          NN_Anhang_<röm>_<Titel>.docx (je Anhang, bei Überlänge in Teile)
04_Aenderungsverordnungen/            VO_2026-1744_NN_Art_<n>_<Titel>[_Teil_k_Nr_x-y].docx
05_EU_Kommission_Leitlinien/          <id>_S001-012.docx … (Entwürfe: ENTWURF_…)
```

Die Aufteilung ergibt sich automatisch aus der tatsächlichen Gliederung: je Kapitel bzw. Abschnitt
ein Dokument; nur wenn ein Abschnitt größer als ca. 15 Seiten ist, wird an Artikelgrenzen geteilt
(Anhänge und Artikel der Änderungsverordnung an Gliederungspunkten der obersten Ebene). Artikel werden
nicht zerteilt. Erwägungsgründe werden in Blöcke von ca. 12 Seiten gebündelt.

## Wortlauttreue

* Text wird nur technisch normalisiert: Leerraum, geschützte Leerzeichen, Weichtrennzeichen und die
  redaktionellen EUR-Lex-Konsolidierungsmarker `▼B`, `►M1`, `◄` werden entfernt. Nummern, Buchstaben,
  Reihenfolge und Wortlaut bleiben unverändert; Hoch-/Tiefstellungen (z. B. 10²⁵) bleiben als Formatierung erhalten.
* Aufzählungen aus EUR-Lex-Layouttabellen werden als eingerückte Absätze mit Originalzeichen ausgegeben,
  echte Tabellen als Word-Tabellen; Tabellen mit verbundenen Zellen werden in strukturierten Text
  („Zeile n: Spalte: Wert | …“) überführt. Die dabei ergänzten Spaltenbezeichnungen tragen die
  Zeichenformatvorlage `RAG Tabellenlabel`.
* Metadaten, wiederholte Kontextüberschriften und Fußnotenüberschriften tragen Formatvorlagen `RAG …`
  und gehören nicht zum amtlichen Text. Fußnoten stehen am Ende des jeweiligen Dokuments.

## Qualitätsprüfung (automatisch)

* Kapitel I–XIII, Artikel 1–113 (plus erkannte eingefügte Artikel, z. B. „4a“), Anhänge I–XIII vorhanden;
  keine Dubletten, keine Lücken, aufsteigende Reihenfolge.
* Abgleich der Überschriftenfolge mit einem zweiten, unabhängigen HTML-Scan und mit den ELI-Kennungen (`art_*`, `anx_*`, `rct_*`).
* Erwägungsgründe 1–180 (bzw. 1–n bei VO 2026/1744) lückenlos.
* **Vollständigkeit:** der gesamte sichtbare Quelltext (ohne Leerraum) muss zeichengenau der Folge aller
  erfassten Blöcke entsprechen; jeder Block muss genau einem Word-Dokument zugeordnet sein.
* **Wortlaut der Word-Dateien:** jede Datei wird zurückgelesen und zeichengenau (ohne Leerraum) mit
  den Quellblöcken verglichen, einschließlich Fußnoten; Metadaten (CELEX) vorhanden.
* Plausibilität des Textendes je Artikel/Erwägungsgrund (Hinweis auf Abschneiden).

## Tests

`python -m pytest tests` erzeugt synthetische HTML-Dateien im EUR-Lex-Markup (Platzhaltertext, kein
Gesetzestext) und prüft Parser, Aufteilung, Word-Erzeugung und die Erkennung fehlender Artikel bzw.
veränderter Word-Texte.
