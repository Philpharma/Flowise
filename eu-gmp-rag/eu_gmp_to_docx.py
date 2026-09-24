#!/usr/bin/env python3
"""
EU-GMP-Leitfaden (BMG-Bekanntmachungen) -> RAG-taugliche Word-Dateien.

Ablauf:
  1. Seite https://www.bundesgesundheitsministerium.de/service/gesetze-und-verordnungen/bekanntmachungen
     laden und alle PDF-Links des Abschnitts "EU-GMP Leitfaden" einsammeln
     (oder mit --pdf-dir bereits heruntergeladene PDFs verwenden).
  2. Text wortgetreu extrahieren, dabei entfernen:
       - Kopf-/Fußzeilen und Seitenzahlen (Zeilen, die sich auf vielen Seiten wiederholen)
       - die Dokumenthistorie-Tabelle am Anfang
       - das Deckblatt der Bekanntmachung (Text vor der ersten nummerierten Überschrift)
  3. Tabellen im Fließtext zeilenweise in Sätze umwandeln
     ("Spalte A: Wert; Spalte B: Wert"), damit ein RAG sie lesen kann.
  4. Lange Dokumente (Teil II, Teil IV, große Anhänge) an Kapitelgrenzen in
     Teile von max. ca. 15 Seiten aufteilen und sprechend benennen.
  5. Jede Word-Datei beginnt mit Titel + Quellenzeile; Überschriften sind echte
     Word-Überschriften (Überschrift 1/2), damit Chunker sie erkennen.

Aufruf:
  python eu_gmp_to_docx.py --out ausgabe/            # lädt alles von der BMG-Seite
  python eu_gmp_to_docx.py --pdf-dir pdfs/ --out ausgabe/   # nutzt lokale PDFs
"""
import argparse
import re
import statistics
import sys
import zipfile
from collections import Counter
from pathlib import Path
from urllib.parse import urljoin

import pdfplumber
from docx import Document
from docx.shared import Pt

SOURCE_PAGE = "https://www.bundesgesundheitsministerium.de/service/gesetze-und-verordnungen/bekanntmachungen"

# ~15 Word-Seiten bei 11 pt ≈ 6.000 Wörter. Etwas Luft lassen.
MAX_WORDS = 5500

# Sprechende Namen (Fallback, falls der Titel im PDF nicht sauber erkennbar ist)
TITLES = {
    "Einleitung": "Einleitung EU-GMP-Leitfaden",
    "Kapitel 1": "Teil I Kapitel 1 Pharmazeutisches Qualitätssystem",
    "Kapitel 2": "Teil I Kapitel 2 Personal",
    "Kapitel 3": "Teil I Kapitel 3 Räumlichkeiten und Ausrüstung",
    "Kapitel 4": "Teil I Kapitel 4 Dokumentation",
    "Kapitel 5": "Teil I Kapitel 5 Produktion",
    "Kapitel 6": "Teil I Kapitel 6 Qualitätskontrolle",
    "Kapitel 7": "Teil I Kapitel 7 Ausgelagerte Tätigkeiten",
    "Kapitel 8": "Teil I Kapitel 8 Beanstandungen, Qualitätsmängel und Produktrückrufe",
    "Kapitel 9": "Teil I Kapitel 9 Selbstinspektionen",
    "Teil II": "Teil II Wirkstoffe",
    "Anhang 3": "Anhang 3 Herstellung von Radiopharmaka",
    "Anhang 6": "Anhang 6 Herstellung medizinischer Gase",
    "Anhang 7": "Anhang 7 Herstellung pflanzlicher Arzneimittel",
    "Anhang 8": "Anhang 8 Probenahme von Ausgangs- und Verpackungsmaterial",
    "Anhang 9": "Anhang 9 Herstellung von Liquida, Cremes und Salben",
    "Anhang 10": "Anhang 10 Herstellung von Dosieraerosolen zur Inhalation",
    "Anhang 11": "Anhang 11 Computergestützte Systeme",
    "Anhang 12": "Anhang 12 Ionisierende Strahlung in der Arzneimittelherstellung",
    "Anhang 13": "Anhang 13 Prüfpräparate",
    "Anhang 14": "Anhang 14 Arzneimittel aus menschlichem Blut oder Plasma",
    "Anhang 15": "Anhang 15 Qualifizierung und Validierung",
    "Anhang 16": "Anhang 16 Zertifizierung durch die Sachkundige Person und Chargenfreigabe",
    "Anhang 19": "Anhang 19 Referenz- und Rückstellmuster",
    "Teil IV": "Teil IV Arzneimittel für neuartige Therapien (ATMP)",
}

LINK_RE = re.compile(r"^(Einleitung|Kapitel \d+|Teil [IV]+|Anhang \d+)\b")
HISTORY_KEYWORDS = ("historie", "history", "grund der änderung", "überarbeitung",
                    "inkrafttreten", "datum der", "revision", "stand der")
PAGE_NO_RE = re.compile(r"^(seite\s*)?\d{1,3}(\s*(von|/)\s*\d{1,3})?$", re.I)
H1_RE = re.compile(r"^(\d{1,2})\.?\s+([A-ZÄÖÜ][^.]{2,120})$")
H2_RE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.?\s+([A-ZÄÖÜ][^.]{2,120})$")


# --------------------------------------------------------------------------- Download
def download_pdfs(target: Path) -> list[tuple[str, Path]]:
    import requests
    from bs4 import BeautifulSoup

    target.mkdir(parents=True, exist_ok=True)
    html = requests.get(SOURCE_PAGE, timeout=60).text
    soup = BeautifulSoup(html, "html.parser")

    # Abschnitt "EU-GMP Leitfaden" suchen (Akkordeon-Überschrift) und seine Links nehmen
    anchor = soup.find(string=re.compile(r"EU-GMP[- ]Leitfaden"))
    if anchor is None:
        sys.exit("Abschnitt 'EU-GMP Leitfaden' auf der Seite nicht gefunden.")
    container = anchor.find_parent()
    while container is not None and len([a for a in container.find_all("a") if ".pdf" in (a.get("href") or "")]) < 5:
        container = container.find_parent()

    result, seen = [], set()
    for a in container.find_all("a"):
        href = a.get("href") or ""
        text = " ".join(a.get_text(" ").split())
        m = LINK_RE.match(text)
        if not m or ".pdf" not in href.lower() or m.group(1) in seen:
            continue
        seen.add(m.group(1))
        url = urljoin(SOURCE_PAGE, href)
        path = target / f"{m.group(1)}.pdf"
        if not path.exists():
            r = requests.get(url, timeout=120)
            r.raise_for_status()
            path.write_bytes(r.content)
        print(f"  geladen: {m.group(1):12s} <- {url}")
        result.append((m.group(1), path))
    return result


# --------------------------------------------------------------------------- Extraktion
def table_to_sentences(rows: list[list[str | None]]) -> list[str]:
    """Tabelle -> eine Zeile Text pro Tabellenzeile, mit Spaltenköpfen als Schlüssel."""
    rows = [[" ".join((c or "").split()) for c in r] for r in rows if any(c for c in r)]
    if not rows:
        return []
    # verbundene Zellen: leere Zellen von oben auffüllen (nur erste Spalte)
    for i in range(1, len(rows)):
        if not rows[i][0] and rows[i - 1][0]:
            rows[i][0] = rows[i - 1][0]
    header, body = rows[0], rows[1:]
    if not body or not any(header):
        return [" | ".join(c for c in r if c) for r in rows]
    out = []
    for r in body:
        parts = []
        for h, c in zip(header, r):
            if c:
                parts.append(f"{h}: {c}" if h else c)
        if parts:
            out.append("; ".join(parts) + ("" if parts[-1].endswith(".") else "."))
    return out


def is_history_table(rows) -> bool:
    text = " ".join(" ".join(c or "" for c in r) for r in rows).lower()
    return any(k in text for k in HISTORY_KEYWORDS)


def extract_blocks(pdf_path: Path) -> list[tuple[str, str]]:
    """Liefert eine Liste von (typ, text); typ in {'line', 'table', 'gap'}."""
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages):
            items = []  # (top, typ, payload)
            tables = page.find_tables()
            bboxes = []
            for t in tables:
                rows = t.extract()
                bboxes.append(t.bbox)
                if pno < 2 and is_history_table(rows):
                    continue  # Dokumenthistorie am Anfang verwerfen
                items.append((t.bbox[1], "table", rows))

            def outside(obj):
                return not any(b[0] - 1 <= obj["x0"] and obj["x1"] <= b[2] + 1 and
                               b[1] - 1 <= obj["top"] and obj["bottom"] <= b[3] + 1 for b in bboxes)

            filtered = page.filter(lambda o: o.get("object_type") != "char" or outside(o))
            for ln in filtered.extract_text_lines(layout=False, strip=True, return_chars=True):
                chars = [c for c in ln["chars"] if c["text"].strip()]
                bold = bool(chars) and sum("bold" in c["fontname"].lower() for c in chars) > len(chars) / 2
                items.append((ln["top"], "line", (ln["text"], ln["bottom"], page.height, bold)))
            items.sort(key=lambda x: x[0])
            pages.append(items)

    # Kopf-/Fußzeilen: Zeilen im oberen/unteren 10 % der Seite, die auf >= 40 % der Seiten vorkommen
    norm = lambda s: re.sub(r"\d+", "#", s.strip().lower())
    counter = Counter()
    for items in pages:
        seen = set()
        for top, typ, p in items:
            if typ == "line" and (top < p[2] * 0.1 or top > p[2] * 0.9):
                seen.add(norm(p[0]))
        counter.update(seen)
    repeated = {k for k, v in counter.items() if len(pages) > 2 and v >= max(2, 0.4 * len(pages))}

    blocks = []
    for items in pages:
        heights = [p[1] - top for top, typ, p in items if typ == "line"]
        lh = statistics.median(heights) if heights else 12
        prev_bottom = None
        for top, typ, p in items:
            if typ == "table":
                blocks.append(("table", p))
                prev_bottom = None
                continue
            text, bottom, ph, bold = p
            edge = top < ph * 0.1 or top > ph * 0.9
            if PAGE_NO_RE.match(text) and edge or norm(text) in repeated:
                continue
            if prev_bottom is not None and top - prev_bottom > lh * 0.8:
                blocks.append(("gap", ""))
            blocks.append(("line", (text, bold)))
            prev_bottom = bottom
    return blocks


def blocks_to_paragraphs(blocks) -> list[tuple[str, str]]:
    """Zeilen zu Absätzen zusammenführen; Überschriften erkennen.
    Rückgabe: (typ, text) mit typ in {'h1','h2','p','table_row'}."""
    paras, buf = [], []
    last_h1 = 0

    def flush():
        if buf:
            paras.append(("p", " ".join(buf).strip()))
            buf.clear()

    for typ, payload in blocks:
        if typ == "gap":
            flush()
            continue
        if typ == "table":
            flush()
            paras.extend(("table_row", s) for s in table_to_sentences(payload))
            continue
        line, bold = payload
        m1, m2 = H1_RE.match(line), H2_RE.match(line)
        # Überschrift = fortlaufend nummeriert UND fett (bzw. kurze Zeile, falls keine Fettschrift erkennbar)
        if m1 and int(m1.group(1)) == last_h1 + 1 and (bold or (len(line) < 60 and not buf)):
            flush()
            last_h1 = int(m1.group(1))
            paras.append(("h1", line))
            continue
        if m2 and int(m2.group(1)) == last_h1 and bold and len(line) < 110:
            flush()
            paras.append(("h2", line))
            continue
        if re.match(r"^([•\-–▪]|\(?[a-z0-9ivx]{1,4}[\).])\s", line):
            flush()  # Aufzählungspunkt beginnt neuen Absatz
        if buf and buf[-1].endswith("-") and not re.match(r"^(und|oder|bzw\.|sowie)\b", line):
            # Silbentrennung auflösen ("Ausgangs-/material"); vor Großbuchstaben Bindestrich behalten ("API-/Ausgangsmaterial")
            buf[-1] = (buf[-1] if line[:1].isupper() else buf[-1][:-1]) + line
            continue
        buf.append(line)
    flush()
    return paras


def strip_front_matter(paras):
    """Deckblatt der Bekanntmachung weglassen: alles vor der ersten Kapitelüberschrift."""
    for i, (t, _) in enumerate(paras):
        if t == "h1":
            return paras[i:] if i < len(paras) * 0.3 else paras
    return paras


# --------------------------------------------------------------------------- Aufteilen
def words(paras) -> int:
    return sum(len(t.split()) for _, t in paras)


def split_sections(paras, level="h1"):
    sections, cur = [], []
    for p in paras:
        if p[0] == level and cur:
            sections.append(cur)
            cur = []
        cur.append(p)
    if cur:
        sections.append(cur)
    return sections


def chunk(paras):
    """Kapitel zusammenfassen, bis MAX_WORDS erreicht; zu große Kapitel an Unterkapiteln teilen."""
    if words(paras) <= MAX_WORDS:
        return [paras]
    units = []
    for sec in split_sections(paras, "h1"):
        if words(sec) <= MAX_WORDS:
            units.append(sec)
        else:
            subs = split_sections(sec, "h2")
            heading = sec[0] if sec[0][0] == "h1" else None
            for i, s in enumerate(subs):
                if i > 0 and heading:  # Kapitelüberschrift als Kontext wiederholen
                    s = [("h1", heading[1] + " (Fortsetzung)")] + s
                units.append(s)
    parts, cur = [], []
    for u in units:
        if cur and words(cur) + words(u) > MAX_WORDS:
            parts.append(cur)
            cur = []
        cur = cur + u
    if cur:
        parts.append(cur)
    return parts


def topic_of(part) -> str:
    heads = []
    for t, text in part:
        if t == "h1" and "(Fortsetzung)" not in text:
            heads.append(re.sub(r"^\d{1,2}\.?\s+", "", text))
    if not heads:
        heads = [re.sub(r"^\d{1,2}\.\d{1,2}\.?\s+", "", t2) for t, t2 in part if t == "h2"][:2]
    nums = [int(H1_RE.match(t2).group(1)) for t, t2 in part if t == "h1" and H1_RE.match(t2)]
    rng = ""
    if nums:
        rng = f"Kap {nums[0]}" if nums[0] == nums[-1] else f"Kap {nums[0]}-{nums[-1]}"
    topic = ", ".join(heads[:3]) + (" u.a." if len(heads) > 3 else "")
    return f"{rng} {topic}".strip()


def safe_name(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*]', "-", s)
    return re.sub(r"\s+", " ", s).strip()[:120]


# --------------------------------------------------------------------------- Word
def write_docx(path: Path, title: str, source: str, part):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(11)
    doc.core_properties.title = title
    doc.core_properties.subject = source
    doc.add_heading(title, level=0)
    doc.add_paragraph(f"Quelle: {source}").italic = True
    in_table = False
    for typ, text in part:
        if typ == "h1":
            doc.add_heading(text, level=1)
        elif typ == "h2":
            doc.add_heading(text, level=2)
        elif typ == "table_row":
            if not in_table:
                doc.add_paragraph("Tabelle (zeilenweise wiedergegeben):").runs[0].bold = True
            doc.add_paragraph(text, style="List Bullet")
        else:
            doc.add_paragraph(text)
        in_table = typ == "table_row"
    doc.save(path)


def process(key: str, pdf: Path, out: Path):
    base = TITLES.get(key, key)
    paras = strip_front_matter(blocks_to_paragraphs(extract_blocks(pdf)))
    parts = chunk(paras)
    source = f"EU-GMP-Leitfaden, {key} (Bekanntmachung des BMG), {SOURCE_PAGE}"
    files = []
    for i, part in enumerate(parts, 1):
        if len(parts) == 1:
            name = base
        else:
            name = f"{base} - Teil {i} von {len(parts)} - {topic_of(part)}"
        name = safe_name(name)
        write_docx(out / f"{name}.docx", name, source, part)
        files.append((name, words(part)))
    return files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="ausgabe")
    ap.add_argument("--pdf-dir", help="Ordner mit bereits geladenen PDFs (Dateiname = Linktext, z.B. 'Kapitel 1.pdf')")
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    if args.pdf_dir:
        pdfs = [(p.stem, p) for p in sorted(Path(args.pdf_dir).glob("*.pdf"))]
    else:
        pdfs = download_pdfs(out / "_pdf")

    for key, pdf in pdfs:
        for name, w in process(key, pdf, out):
            print(f"  {w:6d} Wörter  {name}.docx")

    with zipfile.ZipFile(out.with_suffix(".zip"), "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(out.glob("*.docx")):
            z.write(f, f.name)
    print(f"Fertig: {out.with_suffix('.zip')}")


if __name__ == "__main__":
    main()
