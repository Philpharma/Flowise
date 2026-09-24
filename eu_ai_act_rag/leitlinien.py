"""Offizielle Leitlinien der Europäischen Kommission zur KI-Verordnung (getrennt vom Gesetzestext).

Die PDF-Dokumente werden von *.europa.eu geladen und ihr Text wird seitenweise mit pdfplumber
extrahiert. Zeilenumbrüche innerhalb eines Absatzes werden zusammengeführt (reine
Leerraumänderung); der Text selbst wird nicht verändert.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

import config
from docx_writer import DocSpec, Section, write_docx, read_back
from eurlex_parser import Block, Run, squash
from planner import PlannedDoc, slug

MAX_SEITEN_PRO_DATEI = 15
PARA_END = re.compile(r"[.:;!?)]$")
ITEM_START = re.compile(r"^(\(?\d{1,3}[.)]|\(?[a-z][.)]|\(?[ivx]{1,5}\)|[•\-–—▪]|\d+(\.\d+)+\.?\s)")


@dataclass
class GuidelineResult:
    eintrag: dict
    ok: bool
    meldung: str
    docs: list[PlannedDoc] = field(default_factory=list)
    kandidaten: list[tuple[str, str]] = field(default_factory=list)


def _allowed(url: str) -> bool:
    host = urlparse(url).hostname or ""
    return host == "europa.eu" or host.endswith(config.ALLOWED_GUIDELINE_HOST_SUFFIX)


def _get(url: str, cache: Path) -> bytes:
    if not _allowed(url):
        raise ValueError(f"keine offizielle EU-Adresse: {url}")
    cache.mkdir(parents=True, exist_ok=True)
    f = cache / hashlib.sha256(url.encode()).hexdigest()[:24]
    if f.exists():
        return f.read_bytes()
    r = requests.get(url, headers={"User-Agent": config.USER_AGENT}, timeout=120, allow_redirects=True)
    r.raise_for_status()
    if not _allowed(r.url):
        raise ValueError(f"Weiterleitung auf nicht offizielle Adresse: {r.url}")
    f.write_bytes(r.content)
    return r.content


def find_pdf_links(html: bytes, base: str) -> list[tuple[str, str]]:
    soup = BeautifulSoup(html, "lxml")
    out = []
    for a in soup.find_all("a", href=True):
        href = urljoin(base, a["href"])
        text = " ".join(a.get_text(" ").split())
        if ("/redirection/document/" in href or href.lower().endswith(".pdf")) and _allowed(href):
            if (text, href) not in out:
                out.append((text, href))
    return out


def choose_pdf(links: list[tuple[str, str]]) -> tuple[str, str] | None:
    if not links:
        return None
    de = [l for l in links if re.search(r"\b(DE|German|Deutsch)\b", l[0])]
    en = [l for l in links if re.search(r"\b(EN|English)\b", l[0])]
    main = [l for l in links if re.search(r"guideline|leitlinie", l[0], re.I)]
    return (de or en or main or links)[0]


def _pdfium_pages(data: bytes) -> list[list[str]]:
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(data)
    pages = []
    for i in range(len(doc)):
        # pdfium gibt Bindestrich-Glyphen (v. a. an Zeilenenden) als U+FFFE aus -> "-"; andere Steuerzeichen entfernen
        txt = doc[i].get_textpage().get_text_range().replace("\ufffe", "-")
        txt = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\uffff]", "", txt)
        pages.append([ln.strip() for ln in txt.replace("\r\n", "\n").replace("\r", "\n").split("\n")])
    return pages


def _plumber_pages(data: bytes) -> list[list[str]]:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [[ln.strip() for ln in (p.extract_text() or "").splitlines()] for p in pdf.pages]


def pdf_pages(data: bytes) -> list[list[str]]:
    """Text je PDF-Seite. Deckblatt (S. 1) über pypdfium2 – pdfplumber vertauscht dort die Zeichen des
    gestalteten Kopfs –, alle übrigen Seiten über pdfplumber (bessere Zeilenbildung bei hochgestellten
    Fußnotenziffern). Beide Extraktoren werden in cross_check gegeneinander geprüft."""
    first = _pdfium_pages(data)
    rest = _plumber_pages(data)
    return first[:1] + rest[1:]


HEADING_RE = re.compile(r"^(?:\d{1,2}(?:\.\d{1,2}){0,3}\.?|[IVX]{1,4}\.|[A-H]\.)\s+\S")


def is_heading(para: str) -> bool:
    """Nummerierte Gliederungsüberschrift (kurz, ohne Satzendezeichen) – nur Formatierung, Text bleibt gleich."""
    return (len(para) <= 110 and HEADING_RE.match(para) is not None and not PARA_END.search(para)
            and not para.rstrip().endswith(",") and not re.search(r"\.{4,}|…{2,}|\s\d+$", para))


def doc_facts(pages: list[list[str]]) -> dict:
    """Datum, Aktenzeichen und Sprache deterministisch aus dem Deckblatt lesen."""
    first = " ".join(" ".join(p) for p in pages[:2])
    facts = {}
    if m := re.search(r"(?:Brüssel, den|Brussels,)\s*([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4}|XXX)", first):
        facts["datum"] = m.group(1)
    if m := re.search(r"(C\(\d{4}\)\s*\d+\s*final|\[…\]\(\d{4}\)\s*XXX\s*draft)", first):
        facts["aktenzeichen"] = m.group(1)
    facts["sprache"] = "DE" if re.search(r"\bDE DE\b|Brüssel", first) else ("EN" if re.search(r"\bEN EN\b|Brussels", first) else "")
    return facts


def cross_check(data: bytes, pages: list[list[str]]) -> list[int]:
    """Unabhängige Zweitextraktion (pypdfium2): Seiten (ab S. 2), deren Zeichenbestand abweicht."""
    import collections
    other = _pdfium_pages(data)
    bad = []
    for i in range(1, min(len(pages), len(other))):
        if collections.Counter(squash("".join(pages[i]))) != collections.Counter(squash("".join(other[i]))):
            bad.append(i + 1)
    return bad


def lines_to_paragraphs(lines: list[str], min_full_line: int = 55) -> list[str]:
    """Umbrochene PDF-Zeilen zu Absätzen verbinden (nur Leerraum wird verändert).

    Eine Zeile wird mit der nächsten verbunden, wenn sie "voll" ist (typischer Zeilenumbruch im
    Fließtext), nicht mit Satzzeichen endet und die nächste Zeile keinen Gliederungspunkt beginnt.
    """
    paras: list[str] = []
    prev_line = ""
    for ln in lines:
        if not ln:
            prev_line = ""
            continue
        if (paras and prev_line and (len(prev_line) >= min_full_line or prev_line.endswith("-"))
                and not PARA_END.search(prev_line)
                and not ITEM_START.match(ln)):
            sep = "" if prev_line.endswith("-") else " "
            paras[-1] = paras[-1] + sep + ln
        else:
            paras.append(ln)
        prev_line = ln
    return paras


def build(out: Path, cache: Path, include_drafts: bool, quellen: Path, pdf_dir: Path | None = None) -> list[GuidelineResult]:
    cfg = json.loads(quellen.read_text(encoding="utf-8"))
    ordner = config.ORDNER["leitlinien"]
    results = []
    for e in cfg["leitlinien"]:
        draft = e["status"] != "final"
        if draft and not include_drafts:
            results.append(GuidelineResult(e, True, "Entwurf – nicht übernommen (Option --entwuerfe)"))
            continue
        local_pdf = pdf_dir / f"{e['id']}.pdf" if pdf_dir else None
        local_md = pdf_dir / f"{e['id']}.md" if pdf_dir else None
        if local_md is not None and local_md.exists() and not local_pdf.exists():
            try:
                results.append(build_from_markdown(e, local_md, out, draft))
            except Exception as exc:  # noqa: BLE001
                results.append(GuidelineResult(e, False, f"Markdown-Verarbeitung fehlgeschlagen: {exc}"))
            continue
        try:
            if local_pdf is not None and local_pdf.exists():
                page, links, title = b"", [], ""
                chosen = ("lokal gespeichertes PDF", f"{e['seite']} (Datei {local_pdf.name})")
                data = local_pdf.read_bytes()
            else:
                page = _get(e["seite"], cache)
                soup = BeautifulSoup(page, "lxml")
                title = " ".join((soup.title.get_text() if soup.title else "").split())
                links = find_pdf_links(page, e["seite"])
                chosen = choose_pdf(links)
                if not chosen:
                    results.append(GuidelineResult(e, False, "kein PDF-Link auf der Kommissionsseite gefunden", kandidaten=links))
                    continue
                data = _get(chosen[1], cache)
            warn = ""
            if not draft and re.search(r"\bdraft\b|consultation|entwurf", title, re.I):
                warn = f"Seitentitel deutet auf Entwurf hin ({title!r}) – als ENTWURF gekennzeichnet"
                draft = True
            if not data.startswith(b"%PDF"):
                results.append(GuidelineResult(e, False, f"Download ist kein PDF: {chosen[1]}", kandidaten=links))
                continue
            pages = pdf_pages(data)
            facts = doc_facts(pages)
            xbad = cross_check(data, pages)
            sha = hashlib.sha256(data).hexdigest()
            docs = []
            n_files = -(-len(pages) // MAX_SEITEN_PRO_DATEI)
            per_file = -(-len(pages) // n_files)  # gleichmäßig verteilen, keine Kleinstreste
            for start in range(0, len(pages), per_file):
                chunk = pages[start:start + per_file]
                blocks: list[Block] = []
                sections: list[Section] = []
                for pno, lines in enumerate(chunk, start + 1):
                    for para in lines_to_paragraphs(lines):
                        b = Block(idx=-1, kind="p", runs=[Run(para)])
                        blocks.append(b)
                        if pno > 1 and is_heading(para):
                            sections.append(Section("heading", 2 if para.count(".") <= 1 else 3, [b]))
                        elif sections and sections[-1].kind == "blocks":
                            sections[-1].blocks.append(b)
                        else:
                            sections.append(Section("blocks", blocks=[b]))
                von, bis = start + 1, start + len(chunk)
                prefix = "ENTWURF_" if draft and "ENTWURF" not in e["id"] else ""
                name = f"{prefix}{e['id']}_S{von:03d}-{bis:03d}.docx"
                meta = [
                    ("Titel", e["titel"]),
                    ("Herausgeber", "Europäische Kommission"),
                    ("Status", "ENTWURF / DRAFT" if draft else "endgültig veröffentlicht"),
                    ("Datum / Aktenzeichen", f"{facts.get('datum', e.get('datum', ''))}; {facts.get('aktenzeichen', '–')}"),
                    ("Sprache", facts.get("sprache", "") or "–"),
                    ("Bezug", f"Verordnung (EU) 2024/1689, {e.get('artikel', '')}"),
                    ("Quelle", f"{e['seite']} → {chosen[1]} ({chosen[0]})"),
                    ("Abgerufen am", f"{dt.date.today().isoformat()}; SHA-256 PDF: {sha[:16]}…"),
                    ("Enthaltener Inhalt", f"PDF-Seiten {von}–{bis} von {len(pages)}"),
                    ("Hinweis", "Nicht verbindliche Leitlinien, kein Gesetzestext. Text per PDF-Extraktion "
                                "übernommen; Tabellen/Grafiken können im PDF-Original abweichend dargestellt sein." + (" " + warn if warn else "")),
                ]
                spec = DocSpec(out / ordner / name, e["titel"] + f" (S. {von}–{bis})", meta,
                               sections, entwurf=draft,
                               core_subject="Leitlinien der Kommission zur KI-Verordnung",
                               core_keywords=f"Leitlinien; {e.get('artikel', '')}")
                write_docx(spec)
                exp = squash("".join(b.text for b in blocks))
                got = squash("".join(read_back(spec.path)["content"]))
                src = squash("".join("".join(l) for l in chunk))
                ok = exp == got == src and not any(von <= x <= bis for x in xbad)
                pd = PlannedDoc(spec, "Leitlinie (ENTWURF)" if draft else "Leitlinie", ordner,
                                celex=f"– ({facts.get('aktenzeichen', 'Kommissionsdokument')})",
                                fassung=facts.get("datum", e.get("datum", "")),
                                qa_status="OK" if ok else "FEHLER")
                pd.anhang = ""
                pd.kapitel = e.get("artikel", "")
                if not ok:
                    pd.qa_details.append("Extrahierter PDF-Text weicht vom Word-Inhalt oder von der Zweitextraktion ab"
                                         + (f" (Seiten {[x for x in xbad if von <= x <= bis]})" if xbad else ""))
                else:
                    pd.qa_details.append("Text von S. 2 an mit zweitem Extraktor (pypdfium2) gegengeprüft; Deckblatt nur pypdfium2")
                docs.append(pd)
            results.append(GuidelineResult(e, all(d.qa_status == "OK" for d in docs),
                                           f"{len(pages)} PDF-Seiten, {len(docs)} Datei(en), Sprache {facts.get('sprache') or '?'}, {facts.get('aktenzeichen', '')}" + (f"; {warn}" if warn else ""),
                                           docs, links))
        except Exception as exc:  # noqa: BLE001 – Fehler je Leitlinie berichten, Lauf fortsetzen
            hint = f"kein lokales PDF {local_pdf.name} vorhanden; " if local_pdf is not None else ""
            results.append(GuidelineResult(e, False, f"{hint}Abruf/Verarbeitung fehlgeschlagen: {str(exc)[:160]}"))
    return results


def build_from_markdown(e: dict, md_path: Path, out: Path, draft: bool) -> GuidelineResult:
    """Leitlinie aus einer vom Nutzer bereitgestellten Markdown-Konvertierung des Kommissions-PDF."""
    from markdown_quelle import parse_markdown, reference_text
    from planner import pack

    md = md_path.read_text(encoding="utf-8")
    items, stats = parse_markdown(md)
    facts = doc_facts([[md[:3000]]])
    ordner = config.ORDNER["leitlinien"]
    # Segmente an Überschriften der Ebenen 1–3 beginnen
    segments: list[list[tuple[str, int, Block]]] = []
    for it in items:
        if not segments or (it[0] == "heading" and it[1] <= 3):
            segments.append([])
        segments[-1].append(it)
    size = lambda seg: sum(len(b.text) if b.kind == "p" else len(b.text) for _, _, b in seg)  # noqa: E731
    groups = pack(segments, size, config.MAX_ZEICHEN, config.ZIEL_ZEICHEN)
    sha = hashlib.sha256(md.encode()).hexdigest()
    docs = []
    heads_all = [b.text for k, _, b in items if k == "heading"]
    for gi, g in enumerate(groups, 1):
        flat = [it for seg in g for it in seg]
        sections: list[Section] = []
        for kind, lvl, b in flat:
            if kind == "heading":
                sections.append(Section("heading", lvl, [b]))
            elif sections and sections[-1].kind == "blocks":
                sections[-1].blocks.append(b)
            else:
                sections.append(Section("blocks", blocks=[b]))
        heads = [b.text for k, _, b in flat if k == "heading"]
        nums = [n for n in (h.split(" ", 1)[0].rstrip(".") for h in heads) if re.fullmatch(r"[IVX]+|\d+(?:\.\d+)*", n)]
        rng = f"_Abschn_{nums[0]}-{nums[-1]}" if nums else ""
        name = f"{e['id']}_Teil_{gi:02d}{rng}.docx" if len(groups) > 1 else f"{e['id']}.docx"
        meta = [
            ("Titel", e["titel"]),
            ("Herausgeber", "Europäische Kommission"),
            ("Status", "ENTWURF / DRAFT" if draft else "endgültig veröffentlicht"),
            ("Datum / Aktenzeichen", f"{facts.get('datum', e.get('datum', ''))}; {facts.get('aktenzeichen', '–')}"),
            ("Sprache", facts.get("sprache", "") or "–"),
            ("Bezug", f"Verordnung (EU) 2024/1689, {e.get('artikel', '')}"),
            ("Quelle", f"{e['seite']} – Textgrundlage: vom Nutzer bereitgestellte Markdown-Konvertierung "
                       f"des Kommissions-PDF ({md_path.name}, Konverter-Merkmale: marker)"),
            ("Übernommen am", f"{dt.date.today().isoformat()}; SHA-256 Markdown: {sha[:16]}…"),
            ("Enthaltener Inhalt", f"Teil {gi} von {len(groups)}: " + (f"{heads[0]} … {heads[-1]}" if heads else "Vorspann")),
            ("Hinweis", "Nicht verbindliche Leitlinien, kein Gesetzestext. Nicht gegen das Original-PDF geprüft; "
                        f"{stats.bilder} Bild(er) der Vorlage (Logo/Grafik) nicht übernommen. Link-Ziele in <…> ergänzt."),
        ]
        spec = DocSpec(out / ordner / name, e["titel"] + (f" (Teil {gi} von {len(groups)})" if len(groups) > 1 else ""),
                       meta, sections, entwurf=draft, core_subject="Leitlinien der Kommission zur KI-Verordnung",
                       core_keywords=f"Leitlinien; {e.get('artikel', '')}")
        write_docx(spec)
        exp = squash("".join(b.text if b.kind == "p" else "".join(c.text for r in b.rows for c in r) for _, _, b in flat))
        got = squash("".join(read_back(spec.path)["content"])).replace("•", "")
        ok = exp.replace("•", "") == got
        pd = PlannedDoc(spec, "Leitlinie (ENTWURF)" if draft else "Leitlinie", ordner,
                        celex=f"– ({facts.get('aktenzeichen', 'Kommissionsdokument')})",
                        fassung=facts.get("datum", e.get("datum", "")), qa_status="OK" if ok else "FEHLER")
        pd.kapitel = e.get("artikel", "")
        pd.qa_details.append("Quelle: Markdown-Konvertierung (nicht gegen PDF geprüft)" if ok
                             else "Word-Inhalt weicht vom Markdown-Text ab")
        docs.append(pd)
    ref = squash(reference_text(md))
    parsed = squash("".join(b.text if b.kind == "p" else "".join(c.text for r in b.rows for c in r) for _, _, b in items))
    complete = ref == parsed
    if not complete:
        for d in docs:
            d.qa_status = "FEHLER"
            d.qa_details.append("Markdown-Text nicht vollständig übernommen (Abgleich mit Referenztext)")
    ok = complete and all(d.qa_status == "OK" for d in docs)
    return GuidelineResult(e, ok, f"Markdown ({md_path.name}): {len(heads_all)} Überschriften, {len(docs)} Datei(en), "
                                  f"Sprache {facts.get('sprache') or '?'}, {facts.get('aktenzeichen', '')}; "
                                  f"Vollständigkeitsabgleich {'bestanden' if complete else 'NICHT bestanden'}", docs)


def discover(cache: Path, quellen: Path) -> list[tuple[str, str, str]]:
    """Listet Leitlinien-Kandidaten auf den Übersichtsseiten der Kommission (nur zur manuellen Prüfung)."""
    cfg = json.loads(quellen.read_text(encoding="utf-8"))
    known = {e["seite"] for e in cfg["leitlinien"]}
    found = []
    for seed in cfg.get("discovery_seiten", []):
        try:
            html = _get(seed, cache)
        except Exception as exc:  # noqa: BLE001
            found.append((seed, "", f"nicht abrufbar: {exc}"))
            continue
        soup = BeautifulSoup(html, "lxml")
        for a in soup.find_all("a", href=True):
            href = urljoin(seed, a["href"])
            text = " ".join(a.get_text(" ").split())
            if re.search(r"guideline|leitlinie", text + href, re.I) and _allowed(href) and href not in known:
                status = "draft" if re.search(r"draft|consultation", text + href, re.I) else "unklar/final?"
                if (text, href, status) not in found:
                    found.append((text, href, status))
    return found
