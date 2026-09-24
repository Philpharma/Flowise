"""Abruf der amtlichen Dokumente von EUR-Lex (mit lokalem Cache und Prüfsummen)."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

import requests

from config import ALLOWED_LAW_HOSTS, CELLAR_URL, USER_AGENT, Source


class FetchError(RuntimeError):
    pass


@dataclass
class FetchedDocument:
    source: Source
    url: str
    html: str
    sha256: str
    abgerufen_am: str
    herkunft: str  # "download" | "cache" | "lokale Datei"


def _check_host(url: str, allowed: tuple[str, ...]) -> None:
    host = urlparse(url).hostname or ""
    if host not in allowed:
        raise FetchError(f"Host {host!r} ist keine zugelassene amtliche Quelle für Gesetzestext")


def _plausible(source: Source, html: str) -> bool:
    # EUR-Lex liefert bei Überlast/Bot-Schutz gelegentlich eine Hinweis- statt der Dokumentseite.
    if len(html) < 20_000:
        return False
    number = {"32024R1689": "2024/1689", "32026R1744": "2026/1744"}.get(source.celex, "2024/1689")
    return number in html


def fetch(source: Source, cache_dir: Path, html_dir: Path | None = None, offline: bool = False,
          refresh: bool = False, retries: int = 4) -> FetchedDocument:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"{source.celex}.html"
    meta_file = cache_dir / f"{source.celex}.json"
    url = source.url()
    _check_host(url, ALLOWED_LAW_HOSTS)

    if html_dir is not None:
        local = html_dir / f"{source.celex}.html"
        if local.exists():
            html = local.read_text(encoding="utf-8")
            if not _plausible(source, html):
                raise FetchError(f"{local} scheint nicht das Dokument {source.celex} zu enthalten")
            return FetchedDocument(source, url, html, hashlib.sha256(html.encode()).hexdigest(),
                                   dt.datetime.fromtimestamp(local.stat().st_mtime).isoformat(timespec="seconds"),
                                   "lokale Datei")

    if cache_file.exists() and not refresh:
        html = cache_file.read_text(encoding="utf-8")
        meta = json.loads(meta_file.read_text(encoding="utf-8")) if meta_file.exists() else {}
        return FetchedDocument(source, meta.get("url", url), html, hashlib.sha256(html.encode()).hexdigest(),
                               meta.get("abgerufen_am", ""), "cache")
    if offline:
        raise FetchError(f"Offline-Modus: keine zwischengespeicherte Fassung für {source.celex} in {cache_dir}")

    last_error: Exception | None = None
    attempts = [(url, {"Accept-Language": "de"})] * retries
    # Rückfall: Cellar des Amts für Veröffentlichungen (technische Quelle hinter EUR-Lex, identischer Inhalt)
    attempts.append((CELLAR_URL.format(celex=source.celex), {"Accept": "application/xhtml+xml, text/html",
                                                             "Accept-Language": "deu"}))
    for attempt, (get_url, headers) in enumerate(attempts):
        try:
            resp = requests.get(get_url, headers={"User-Agent": USER_AGENT, **headers}, timeout=120)
            _check_host(resp.url, ALLOWED_LAW_HOSTS)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            html = resp.text
            if not _plausible(source, html):
                raise FetchError(f"Antwort für {source.celex} ist nicht das erwartete Dokument (Länge {len(html)})")
            ts = dt.datetime.now().isoformat(timespec="seconds")
            sha = hashlib.sha256(html.encode()).hexdigest()
            cache_file.write_text(html, encoding="utf-8")
            meta_file.write_text(json.dumps({"celex": source.celex, "url": get_url, "abgerufen_am": ts,
                                             "sha256": sha}, indent=2), encoding="utf-8")
            herkunft = "download EUR-Lex" if get_url == url else "download Cellar (Amt für Veröffentlichungen)"
            return FetchedDocument(source, url, html, sha, ts, herkunft)
        except (requests.RequestException, FetchError) as exc:
            last_error = exc
            if attempt < len(attempts) - 1:
                time.sleep(2 ** min(attempt + 1, 4))
    raise FetchError(f"Abruf von {url} fehlgeschlagen: {last_error}")
