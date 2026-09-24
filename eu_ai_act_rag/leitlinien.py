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

SEITEN_PRO_DATEI = 12
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


def pdf_pages(data: bytes) -> list[list[str]]:
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for p in pdf.pages:
            txt = p.extract_text() or ""
            pages.append([ln.strip() for ln in txt.splitlines()])
    return pages


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
            sha = hashlib.sha256(data).hexdigest()
            docs = []
            for start in range(0, len(pages), SEITEN_PRO_DATEI):
                chunk = pages[start:start + SEITEN_PRO_DATEI]
                blocks: list[Block] = []
                for pno, lines in enumerate(chunk, start + 1):
                    for para in lines_to_paragraphs(lines):
                        blocks.append(Block(idx=-1, kind="p", runs=[Run(para)]))
                von, bis = start + 1, start + len(chunk)
                prefix = "ENTWURF_" if draft else ""
                name = f"{prefix}{e['id']}_S{von:03d}-{bis:03d}.docx"
                meta = [
                    ("Dokument", e["titel"]),
                    ("Herausgeber", "Europäische Kommission"),
                    ("Status", "ENTWURF / DRAFT" if draft else "endgültig veröffentlicht"),
                    ("Datum", e.get("datum", "")),
                    ("Bezug", f"Verordnung (EU) 2024/1689, {e.get('artikel', '')}"),
                    ("Quelle", f"{e['seite']} → {chosen[1]} ({chosen[0]})"),
                    ("Abgerufen am", f"{dt.date.today().isoformat()}; SHA-256 PDF: {sha[:16]}…"),
                    ("Enthaltener Inhalt", f"PDF-Seiten {von}–{bis} von {len(pages)}"),
                    ("Hinweis", "Nicht verbindliche Leitlinien, kein Gesetzestext. Text per PDF-Extraktion "
                                "übernommen; Tabellen/Grafiken können im PDF-Original abweichend dargestellt sein." + (" " + warn if warn else "")),
                ]
                spec = DocSpec(out / ordner / name, e["titel"] + f" (S. {von}–{bis})", meta,
                               [Section("blocks", blocks=blocks)], entwurf=draft,
                               core_subject="Leitlinien der Kommission zur KI-Verordnung",
                               core_keywords=f"Leitlinien; {e.get('artikel', '')}")
                write_docx(spec)
                exp = squash("".join(b.text for b in blocks))
                got = squash("".join(read_back(spec.path)["content"]))
                src = squash("".join("".join(l) for l in chunk))
                ok = exp == got == src
                pd = PlannedDoc(spec, "Leitlinie (ENTWURF)" if draft else "Leitlinie", ordner,
                                celex="– (keine CELEX; Kommissionsdokument)", fassung=e.get("datum", ""),
                                qa_status="OK" if ok else "FEHLER")
                pd.anhang = ""
                pd.kapitel = e.get("artikel", "")
                if not ok:
                    pd.qa_details.append("Extrahierter PDF-Text weicht vom Word-Inhalt ab")
                docs.append(pd)
            results.append(GuidelineResult(e, all(d.qa_status == "OK" for d in docs),
                                           f"{len(pages)} PDF-Seiten, {len(docs)} Datei(en)" + (f"; {warn}" if warn else ""),
                                           docs, links))
        except Exception as exc:  # noqa: BLE001 – Fehler je Leitlinie berichten, Lauf fortsetzen
            results.append(GuidelineResult(e, False, f"Abruf/Verarbeitung fehlgeschlagen: {exc}"))
    return results


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
