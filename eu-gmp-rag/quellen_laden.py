#!/usr/bin/env python3
"""Lädt die Zusatzquellen aus quellen.json (EU-Kommission, EUR-Lex, EMA) nach <ziel>/.

Aufruf: python quellen_laden.py quellen.json ausgabe/_quellen
Ergebnis: <key>.pdf / <key>.html / <key>.docx sowie download.log
"""
import json
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36",
      "Accept-Language": "de-DE,de;q=0.9,en;q=0.8"}
EURLEX = "https://eur-lex.europa.eu/legal-content/DE/TXT/HTML/?uri=CELEX:{}"


def get(url):
    r = requests.get(url, headers=UA, timeout=120, allow_redirects=True)
    r.raise_for_status()
    return r


def ext_of(r) -> str:
    head = r.content[:5]
    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"PK"):
        return "docx"
    return "html"


def ema_pdf(page_url: str, match: str) -> str:
    soup = BeautifulSoup(get(page_url).text, "html.parser")
    links = [urljoin(page_url, a["href"]) for a in soup.find_all("a", href=True) if a["href"].lower().endswith(".pdf")]
    for l in links:  # bevorzugt: Link, der das Stichwort enthält und kein "overview"/"comments" ist
        if match.lower() in l.lower() and not re.search(r"overview|comment|presentation|concept", l, re.I):
            return l
    if not links:
        raise RuntimeError("keine PDF-Links auf der Seite gefunden")
    return links[0]


def main():
    manifest, target = Path(sys.argv[1]), Path(sys.argv[2])
    target.mkdir(parents=True, exist_ok=True)
    log = []
    for q in json.loads(manifest.read_text(encoding="utf-8"))["quellen"]:
        try:
            url = EURLEX.format(q["celex"]) if q["typ"] == "eurlex" else q["url"]
            if q["typ"] == "ema":
                url = ema_pdf(url, q.get("pdf_match", ""))
            r = get(url)
            ext = ext_of(r)
            (target / f"{q['key']}.{ext}").write_bytes(r.content)
            log.append(f"OK    {q['key']:26s} {ext:4s} {len(r.content):9d} B  {url}")
        except Exception as e:  # einzelne Fehler sollen den Rest nicht aufhalten
            log.append(f"FEHLER {q['key']:25s} {e}")
        print(log[-1])
    (target / "download.log").write_text("\n".join(log) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
