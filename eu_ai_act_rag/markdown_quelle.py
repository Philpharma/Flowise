"""Einlesen von Markdown-Konvertierungen (z. B. mit "marker" aus Kommissions-PDFs erzeugt).

Es wird ausschließlich Markdown-Syntax aufgelöst (Überschriftszeichen, **fett**, *kursiv*,
<sup>, Tabellen, Listen, Links, Seitenanker des Konverters). Der Text selbst bleibt unverändert.
Überschriften werden anhand der amtlichen Nummerierung erkannt ("I.", "2.", "2.1.", "a)", "i."),
weil die #-Ebenen des Konverters nicht zuverlässig sind.
"""
from __future__ import annotations

import html
import re
from dataclasses import dataclass

from eurlex_parser import Block, Cell, Run

ESCAPABLE = "\\`*_{}[]()#+-.!|<>~"
_ESC = {c: chr(0xE000 + i) for i, c in enumerate(ESCAPABLE)}
_UNESC = {v: k for k, v in _ESC.items()}

PAGE_ANCHOR_RE = re.compile(r"<span\s+id=\"[^\"]*\"\s*>\s*</span>|<a\s+id=\"[^\"]*\"\s*>\s*</a>")
IMAGE_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")
HEADING_MD_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
LIST_RE = re.compile(r"^(\s*)[-*+]\s+(.*)$")
TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$")
NUM_HEADING_RE = re.compile(
    r"^(?P<rom>[IVX]{1,5}\.)\s|^(?P<num>\d{1,2}(?:\.\d{1,2}){0,4}\.?)\s|^(?P<let>[a-z]\))\s|^(?P<low>[ivx]{1,5}\.)\s")

INLINE_RE = re.compile(
    r"(?P<sup><sup>(?P<sup_t>.*?)</sup>)"
    r"|(?P<sub><sub>(?P<sub_t>.*?)</sub>)"
    r"|(?P<bold>\*\*(?P<bold_t>.+?)\*\*|__(?P<bold2_t>.+?)__)"
    r"|(?P<ital>(?<![\w*])\*(?!\s)(?P<ital_t>.+?)(?<!\s)\*(?![\w*])|(?<!\w)_(?!\s)(?P<ital2_t>.+?)(?<!\s)_(?!\w))"
    r"|(?P<link>\[(?P<link_t>[^\]]*)\]\((?P<link_u>[^)\s]*)(?:\s+\"[^\"]*\")?\))"
    r"|(?P<auto><(?P<auto_u>https?://[^>\s]+)>)"
    r"|(?P<br><br\s*/?>)"
    r"|(?P<tag></?(?:span|b|i|em|strong|u)\b[^>]*>)",
    re.S)


@dataclass
class MdStats:
    bilder: int = 0
    links: int = 0


def _protect(s: str) -> str:
    return re.sub(r"\\([\\`*_{}\[\]()#+\-.!|<>~])", lambda m: _ESC[m.group(1)], s)


def _restore(s: str) -> str:
    return "".join(_UNESC.get(ch, ch) for ch in s)


def inline_runs(s: str, stats: MdStats, sup=False, sub=False, bold=False, italic=False) -> list[Run]:
    """Inline-Markdown in formatierte Textläufe zerlegen (rekursiv für Verschachtelungen)."""
    runs: list[Run] = []
    pos = 0
    for m in INLINE_RE.finditer(s):
        if m.start() > pos:
            runs.append(Run(s[pos:m.start()], sup, sub, bold, italic))
        g = m.lastgroup
        if m.group("sup"):
            runs += inline_runs(m.group("sup_t"), stats, True, sub, bold, italic)
        elif m.group("sub"):
            runs += inline_runs(m.group("sub_t"), stats, sup, True, bold, italic)
        elif m.group("bold"):
            runs += inline_runs(m.group("bold_t") or m.group("bold2_t"), stats, sup, sub, True, italic)
        elif m.group("ital"):
            runs += inline_runs(m.group("ital_t") or m.group("ital2_t"), stats, sup, sub, bold, True)
        elif m.group("link"):
            stats.links += 1
            runs += inline_runs(m.group("link_t"), stats, sup, sub, bold, italic)
            url = m.group("link_u")
            if url and url not in m.group("link_t") and not url.startswith("#"):
                runs.append(Run(f" <{url}>", label=True))
        elif m.group("auto"):
            runs.append(Run(m.group("auto_u"), sup, sub, bold, italic))
        elif m.group("br"):
            runs.append(Run(" ", sup, sub, bold, italic))
        # "tag": einfache HTML-Formatierungs-Tags ohne Textinhalt entfallen
        pos = m.end()
        del g
    if pos < len(s):
        runs.append(Run(s[pos:], sup, sub, bold, italic))
    out: list[Run] = []
    for r in runs:
        r.text = html.unescape(_restore(r.text))
        r.text = re.sub(r"[ \t]+", " ", r.text)
        if not r.text:
            continue
        if out and out[-1].text.endswith(" ") and r.text.startswith(" "):
            r.text = r.text[1:]
        if out and (out[-1].sup, out[-1].sub, out[-1].bold, out[-1].italic, out[-1].label) == \
                (r.sup, r.sub, r.bold, r.italic, r.label):
            out[-1].text += r.text
        elif r.text:
            out.append(r)
    if out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text]


def plain(runs: list[Run]) -> str:
    return "".join(r.text for r in runs if not r.label)


def heading_level(text: str) -> int | None:
    """Gliederungsebene aus der Nummerierung; None = keine echte Überschrift."""
    t = text.rstrip()
    if len(t) > 160 or t.endswith((":", ";", ",")) or (t.endswith(".") and len(t) > 60):
        return None
    m = NUM_HEADING_RE.match(text)
    if not m:
        return None
    if m.group("rom"):
        return 1
    if m.group("num"):
        depth = m.group("num").rstrip(".").count(".") + 1
        return min(depth + 1, 4)
    return 4


def parse_markdown(md: str) -> tuple[list[tuple[str, int, Block]], MdStats]:
    """Liefert eine Liste (art, ebene, Block) mit art in {"heading", "p", "table"}."""
    stats = MdStats()
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    lines = md.split("\n")
    items: list[tuple[str, int, Block]] = []
    para: list[str] = []
    i = 0

    def flush():
        nonlocal para
        if para:
            runs = inline_runs(_protect(" ".join(x.strip() for x in para)), stats)
            if plain(runs).strip():
                items.append(("p", 0, Block(idx=-1, kind="p", runs=runs)))
        para = []

    while i < len(lines):
        raw = lines[i]
        line = PAGE_ANCHOR_RE.sub("", raw)
        if IMAGE_RE.search(line):
            stats.bilder += len(IMAGE_RE.findall(line))
            line = IMAGE_RE.sub("", line)
        if not line.strip():
            flush()
            i += 1
            continue
        if line.lstrip().startswith("|"):
            flush()
            rows: list[list[Cell]] = []
            while i < len(lines) and PAGE_ANCHOR_RE.sub("", lines[i]).lstrip().startswith("|"):
                row = PAGE_ANCHOR_RE.sub("", lines[i]).strip()
                i += 1
                if TABLE_SEP_RE.match(row):
                    continue
                cells = [c for c in re.split(r"(?<!\\)\|", _protect(row).strip().strip("|"))]
                rows.append([Cell("\n".join(plain(inline_runs(part, stats))
                                            for part in re.split(r"<br\s*/?>", c) if plain(inline_runs(part, stats))))
                             for c in cells])
            if rows:
                items.append(("table", 0, Block(idx=-1, kind="table", rows=rows)))
            continue
        if (m := HEADING_MD_RE.match(line)) and LIST_RE.match(m.group(2)):
            line = m.group(2)  # Konverter hat einen Aufzählungspunkt als Überschrift markiert ("# - …")
        if m := HEADING_MD_RE.match(line):
            flush()
            runs = inline_runs(_protect(m.group(2)), stats)
            txt = plain(runs)
            if txt.strip():
                lvl = heading_level(txt)
                items.append(("heading" if lvl else "p", lvl or 0, Block(idx=-1, kind="p", runs=runs)))
            i += 1
            continue
        if m := LIST_RE.match(line):
            flush()
            level = 1 + len(m.group(1).replace("\t", "    ")) // 2
            runs = inline_runs(_protect(m.group(2)), stats)
            items.append(("p", 0, Block(idx=-1, kind="p", runs=runs, level=level)))  # kein künstliches Aufzählungszeichen
            i += 1
            # Folgezeilen eines Listenpunkts
            while i < len(lines) and lines[i].strip() and not LIST_RE.match(lines[i]) \
                    and not lines[i].lstrip().startswith(("|", "#")):
                items[-1][2].runs += inline_runs(_protect(" " + lines[i].strip()), stats)
                i += 1
            continue
        para.append(line)
        i += 1
    flush()
    return items, stats


def reference_text(md: str) -> str:
    """Unabhängige, rein zeichenbasierte Entfernung der Markdown-Syntax (Referenz für die QS)."""
    t = md.replace("\r\n", "\n")
    t = re.sub(r"\\([\\`*_{}\[\]()#+\-.!|<>~])", lambda m: _ESC[m.group(1)], t)  # maskierte Zeichen zuerst schützen
    t = PAGE_ANCHOR_RE.sub("", t)
    t = IMAGE_RE.sub("", t)
    t = re.sub(r"\[([^\]]*)\]\([^)\s]*(?:\s+\"[^\"]*\")?\)", r"\1", t)
    t = re.sub(r"<(https?://[^>\s]+)>", r"\1", t)
    t = re.sub(r"</?(?:sup|sub|span|b|i|em|strong|u)\b[^>]*>|<br\s*/?>", "", t)
    t = "\n".join(x for x in t.split("\n") if not TABLE_SEP_RE.match(x.strip()) or not x.strip())
    t = re.sub(r"(?m)^\s*#{1,6}\s+", "", t)
    t = re.sub(r"(?m)^\s*[-*+]\s+", "", t)
    t = "\n".join(x.replace("|", "") if x.lstrip().startswith("|") else x for x in t.split("\n"))
    t = t.replace("**", "").replace("__", "")
    t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"\1", t)
    t = re.sub(r"(?<!\w)_(?!\s)(.+?)(?<!\s)_(?!\w)", r"\1", t)
    return html.unescape(_restore(t))
