"""Deterministischer Parser für EUR-Lex-HTML (Amtsblatt-/ELI-Format und konsolidierte Fassungen).

Arbeitsweise:
1. ``linearize`` zerlegt das HTML verlustfrei in eine geordnete Folge von Textblöcken (Absätze,
   Aufzählungspunkte, echte Tabellen). Jeder Textknoten des Dokuments landet in genau einem Block;
   Fußnoten werden separat erfasst. Der Text wird nur technisch normalisiert (Whitespace,
   geschützte Leerzeichen, Weichtrennzeichen, EUR-Lex-Änderungsmarker ▼/►/◄), nie inhaltlich.
2. ``parse_structure`` ordnet die Blöcke anhand der amtlichen Überschriften (KAPITEL, ABSCHNITT,
   Artikel, ANHANG, nummerierte Erwägungsgründe) einer Gliederung zu.

Es wird kein Text erzeugt, umformuliert, ergänzt oder umsortiert.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from bs4 import BeautifulSoup, NavigableString, Tag, Comment

BLOCK_TAGS = {
    "address", "article", "aside", "blockquote", "body", "center", "dd", "div", "dl", "dt", "fieldset",
    "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li",
    "main", "nav", "ol", "p", "pre", "section", "table", "tbody", "td", "tfoot", "th", "thead", "tr", "ul",
    "html",
}
SKIP_TAGS = {"script", "style", "noscript", "head", "title", "meta", "link", "img", "col", "colgroup", "svg",
             "button", "input", "select", "template"}
NOTE_CLASSES = {"oj-note", "note", "footnote"}

# Aufzählungszeichen: (1)  1.  a)  (a)  i)  (iv)  1.1.  —
_E = r"(?:\d{1,3}(?:\.\d{1,3})*[a-z]?|[a-zA-Z]{1,2}|[ivxlcIVXLC]{1,6})"
ENUM_RE = re.compile(rf"^(?:\({_E}\)|{_E}[.)]|[—–\-•·▪])$")

# EUR-Lex-Konsolidierungsmarker (▼B, ►M1, ►C1, ◄) – redaktionelle Kennzeichen, kein Normtext.
MARKER_RE = re.compile(r"[▼►]\s?(?:B|[MACR]\d{1,3})\b|◄")
INVISIBLE = dict.fromkeys(map(ord, "­​‌‍﻿"), None)
WS_RE = re.compile(r"[ \t\r\n\f\v  -   　]+")
ELI_ID_RE = re.compile(r"^(?:tit|pbl|cit|rct|enc|cpt|sct|art|anx|fnp|prt|ttl|sbs)_")


@dataclass
class Run:
    text: str
    sup: bool = False
    sub: bool = False
    bold: bool = False
    italic: bool = False


@dataclass
class Cell:
    text: str
    colspan: int = 1
    rowspan: int = 1
    header: bool = False


@dataclass
class Block:
    idx: int
    kind: str  # "p" | "table"
    runs: list[Run] = field(default_factory=list)
    enum: str = ""
    level: int = 0
    cls: str = ""
    eli_id: str = ""
    notes: list[str] = field(default_factory=list)
    rows: list[list[Cell]] | None = None

    @property
    def text(self) -> str:
        if self.kind == "table":
            return " ".join(c.text for r in self.rows for c in r if c.text)
        return "".join(r.text for r in self.runs)

    @property
    def full_text(self) -> str:
        return f"{self.enum} {self.text}".strip() if self.enum else self.text

    @property
    def is_complex_table(self) -> bool:
        if self.kind != "table":
            return False
        widths = {sum(c.colspan for c in r) for r in self.rows}
        return len(widths) > 1 or any(c.colspan > 1 or c.rowspan > 1 for r in self.rows for c in r)


def normalize(text: str) -> str:
    text = text.translate(INVISIBLE)
    text = MARKER_RE.sub(" ", text)
    return WS_RE.sub(" ", text)


def squash(text: str) -> str:
    """Vergleichsform: sämtlicher Leerraum entfernt (für die wortlautgetreue Kontrolle)."""
    return WS_RE.sub("", MARKER_RE.sub("", text.translate(INVISIBLE)))


def _classes(el: Tag) -> list[str]:
    c = el.get("class") or []
    return c if isinstance(c, list) else str(c).split()


def _is_note(el: Tag) -> bool:
    return el.name in ("p", "div", "dd", "li") and bool(NOTE_CLASSES & set(_classes(el)))


def _has_block_desc(el: Tag) -> bool:
    return any(isinstance(d, Tag) and d.name in BLOCK_TAGS for d in el.descendants)


class Linearizer:
    def __init__(self, soup: BeautifulSoup):
        self.blocks: list[Block] = []
        self.notes: dict[str, Block] = {}
        self._note_ids: dict[str, str] = {}
        self._pending_enum: str | None = None
        for br in soup.find_all("br"):
            br.replace_with(" ")
        for c in soup.find_all(string=lambda s: isinstance(s, Comment)):
            c.extract()
        self._collect_notes(soup)

    # -- Fußnoten ---------------------------------------------------------------------------------
    def _collect_notes(self, soup):
        for el in soup.find_all(_is_note):
            if el.find_parent(_is_note):
                continue
            ids = [el.get("id")] + [a.get("id") for a in el.find_all(id=True)] + [a.get("name") for a in el.find_all("a", attrs={"name": True})]
            ids = [i for i in ids if i]
            key = ids[0] if ids else f"note_{len(self.notes) + 1}"
            runs = self._runs(list(el.children), el)
            self.notes[key] = Block(idx=-1, kind="p", runs=runs, cls=" ".join(_classes(el)))
            for i in ids:
                self._note_ids[i] = key

    def _note_refs(self, nodes) -> list[str]:
        refs = []
        for n in nodes:
            if not isinstance(n, Tag):
                continue
            anchors = ([n] if n.name == "a" else []) + n.find_all("a", href=True)
            for a in anchors:
                href = a.get("href") or ""
                if "#" in href:
                    key = self._note_ids.get(href.split("#", 1)[1])
                    if key and key not in refs:
                        refs.append(key)
        return refs

    # -- Textläufe --------------------------------------------------------------------------------
    def _runs(self, nodes, container: Tag) -> list[Run]:
        raw: list[Run] = []

        def fmt(node) -> tuple[bool, bool, bool, bool]:
            sup = sub = bold = italic = False
            p = node.parent
            while p is not None and p is not container.parent:
                cls = " ".join(_classes(p)) if isinstance(p, Tag) else ""
                name = p.name if isinstance(p, Tag) else ""
                sup |= name == "sup" or "super" in cls
                sub |= name == "sub" or "oj-sub" in cls or cls.endswith(" sub")
                bold |= name in ("b", "strong") or "bold" in cls
                italic |= name in ("i", "em") or "italic" in cls
                if p is container:
                    break
                p = p.parent
            return sup, sub, bold, italic

        def walk(node):
            if isinstance(node, Comment):
                return
            if isinstance(node, NavigableString):
                t = normalize(str(node))
                if t:
                    raw.append(Run(t, *fmt(node)))
                return
            if node.name in SKIP_TAGS or _is_note(node):
                return
            for ch in node.children:
                walk(ch)

        for n in nodes:
            walk(n)
        runs: list[Run] = []
        for r in raw:
            if runs and runs[-1].text.endswith(" ") and r.text.startswith(" "):
                r.text = r.text.lstrip(" ")
            if not r.text:
                continue
            if runs and (runs[-1].sup, runs[-1].sub, runs[-1].bold, runs[-1].italic) == (r.sup, r.sub, r.bold, r.italic):
                runs[-1].text += r.text
            else:
                runs.append(r)
        if runs:
            runs[0].text = runs[0].text.lstrip(" ")
            runs[-1].text = runs[-1].text.rstrip(" ")
        return [r for r in runs if r.text]

    # -- Blockbildung -----------------------------------------------------------------------------
    @staticmethod
    def _eli_id(el: Tag) -> str:
        for p in [el, *el.parents]:
            if isinstance(p, Tag) and p.get("id") and ELI_ID_RE.match(p["id"]):
                return p["id"]
        return ""

    def _emit(self, runs, level, el: Tag, notes, kind="p", rows=None):
        enum = self._pending_enum or ""
        self._pending_enum = None
        self.blocks.append(Block(idx=len(self.blocks), kind=kind, runs=runs, enum=enum, level=level,
                                 cls=" ".join(_classes(el)), eli_id=self._eli_id(el), notes=notes, rows=rows))

    def _flush(self, group, level, el: Tag, more_follow: bool):
        if not group:
            return
        runs = self._runs(group, el)
        text = "".join(r.text for r in runs)
        if not text:
            return
        if more_follow and ENUM_RE.match(text) and self._pending_enum is None:
            self._pending_enum = text  # z. B. <span class="no-parag">(1)</span> vor dem Absatztext
            return
        self._emit(runs, level, el, self._note_refs(group))

    def process(self, el: Tag, level: int = 0):
        if el.name in SKIP_TAGS or _is_note(el):
            return
        if el.name == "table":
            return self._table(el, level)
        if "grid-list" in _classes(el) or ("grid-container" in _classes(el) and el.find(class_="grid-list-column-1")):
            return self._grid_list(el, level)
        if el.name in ("hr",):
            return
        if not _has_block_desc(el):
            self._flush(list(el.children), level, el, more_follow=False)
            return
        children = list(el.children)
        group: list = []
        for i, ch in enumerate(children):
            if isinstance(ch, Tag) and (ch.name in BLOCK_TAGS or _has_block_desc(ch)):
                self._flush(group, level, el, more_follow=True)
                group = []
                self.process(ch, level + 1 if ch.name == "li" else level)
            else:
                group.append(ch)
        self._flush(group, level, el, more_follow=False)
        if self._pending_enum is not None and el.name in ("td", "li"):
            # Aufzählungszeichen ohne nachfolgenden Text: als eigener Block erhalten
            enum, self._pending_enum = self._pending_enum, None
            self._emit([Run(enum)], level, el, [])

    def _grid_list(self, el: Tag, level: int):
        c1 = el.find(class_="grid-list-column-1")
        c2 = el.find(class_="grid-list-column-2")
        if c1 is None or c2 is None:
            return self.process_children(el, level)
        enum = "".join(r.text for r in self._runs([c1], c1))
        self._pending_enum = enum or None
        self.process(c2, level + 1)
        if self._pending_enum is not None:
            self._emit([Run(self._pending_enum)], level + 1, el, [])
            self._pending_enum = None

    def process_children(self, el: Tag, level: int):
        for ch in el.children:
            if isinstance(ch, Tag):
                self.process(ch, level)

    @staticmethod
    def _rows(tbl: Tag) -> list[list[Tag]]:
        rows = []
        for tr in tbl.find_all("tr"):
            if tr.find_parent("table") is tbl:
                rows.append([c for c in tr.find_all(["td", "th"], recursive=False)])
        return [r for r in rows if r]

    def _cell_text(self, cell: Tag) -> str:
        # Zelleninhalt mit derselben (verlustfreien) Logik linearisieren, Absätze mit Zeilenumbruch trennen
        saved, saved_enum = self.blocks, self._pending_enum
        self.blocks, self._pending_enum = [], None
        self.process(cell, 0)
        texts = [b.full_text for b in self.blocks]
        self.blocks, self._pending_enum = saved, saved_enum
        return "\n".join(t for t in texts if t)

    def _table(self, tbl: Tag, level: int):
        rows = self._rows(tbl)
        if not rows:
            return
        first_texts = [normalize(r[0].get_text()).strip() for r in rows]
        is_list = (all(len(r) == 2 for r in rows)
                   and any(ENUM_RE.match(t) for t in first_texts)
                   and all(t == "" or ENUM_RE.match(t) for t in first_texts))
        if is_list:
            for r, t in zip(rows, first_texts):
                if self._pending_enum is not None:  # verwaistes Zeichen erhalten
                    self._emit([Run(self._pending_enum)], level + 1, tbl, [])
                self._pending_enum = t or None
                self.process(r[1], level + 1)
                if self._pending_enum is not None:
                    self._emit([Run(self._pending_enum)], level + 1, tbl, [])
                    self._pending_enum = None
            return
        if any(c.find("table") for r in rows for c in r):
            # Layouttabelle mit verschachtelten Tabellen: zellenweise linearisieren
            for r in rows:
                for c in r:
                    self.process(c, level)
            return
        cells = [[Cell(self._cell_text(c), int(c.get("colspan", 1) or 1), int(c.get("rowspan", 1) or 1),
                       c.name == "th") for c in r] for r in rows]
        if not any(c.text for r in cells for c in r):
            return
        notes = self._note_refs([tbl])
        self._emit([], level, tbl, notes, kind="table", rows=cells)


def linearize(html: str) -> tuple[list[Block], dict[str, Block], BeautifulSoup]:
    soup = BeautifulSoup(html, "lxml")
    root = soup.body or soup
    lin = Linearizer(soup)
    lin.process(root, 0)
    return lin.blocks, lin.notes, soup


# ================================================================================================
# Gliederung
# ================================================================================================
KAPITEL_RE = re.compile(r"^KAPITEL\s+([IVXLC]+)$")
ABSCHNITT_RE = re.compile(r"^ABSCHNITT\s+(\d+)$")
ARTIKEL_RE = re.compile(r"^Artikel\s+(\d+[a-z]*)$")
ANHANG_RE = re.compile(r"^ANHANG\s+([IVXLC]+[a-z]?)$")
ANHANG_SOLO_RE = re.compile(r"^ANHANG$")
RECITAL_RE = re.compile(r"^\((\d+)\)$")
SCHLUSS_RE = re.compile(r"^Diese Verordnung ist in allen ihren Teilen verbindlich")
ERLASSEN_RE = re.compile(r"HAT FOLGENDE VERORDNUNG ERLASSEN|HABEN FOLGENDE VERORDNUNG ERLASSEN")
BACK_RE = re.compile(r"^(ELI:|ISSN\b)")


@dataclass
class Unit:
    typ: str  # artikel | erwaegungsgrund | anhang
    nummer: str
    heading: list[Block] = field(default_factory=list)
    body: list[Block] = field(default_factory=list)
    kapitel: str | None = None
    abschnitt: str | None = None

    @property
    def titel(self) -> str:
        return " ".join(b.text for b in self.heading[1:]) if len(self.heading) > 1 else ""

    @property
    def blocks(self) -> list[Block]:
        return self.heading + self.body

    @property
    def chars(self) -> int:
        return sum(len(b.full_text) for b in self.blocks)


@dataclass
class Gliederung:
    typ: str  # kapitel | abschnitt
    nummer: str
    heading: list[Block] = field(default_factory=list)
    kapitel: str | None = None

    @property
    def titel(self) -> str:
        return " ".join(b.text for b in self.heading[1:])


@dataclass
class ParsedAct:
    blocks: list[Block]
    notes: dict[str, Block]
    front: list[Block] = field(default_factory=list)
    recitals: list[Unit] = field(default_factory=list)
    preamble_end: list[Block] = field(default_factory=list)
    kapitel: list[Gliederung] = field(default_factory=list)
    abschnitte: list[Gliederung] = field(default_factory=list)
    articles: list[Unit] = field(default_factory=list)
    final: list[Block] = field(default_factory=list)
    annexes: list[Unit] = field(default_factory=list)
    back: list[Block] = field(default_factory=list)
    unassigned: list[Block] = field(default_factory=list)
    stand: str = ""
    titel: str = ""


def _looks_like_art_heading(b: Block) -> bool:
    if b.kind != "p" or b.enum or not ARTIKEL_RE.match(b.text):
        return False
    return b.level == 0 or "art" in b.cls


def parse_structure(blocks: list[Block], notes: dict[str, Block], track_quotes: bool = False) -> ParsedAct:
    act = ParsedAct(blocks=blocks, notes=notes)
    state = "front"
    cur_kap = cur_abs = None
    cur_unit: Unit | None = None
    expect_title: list[Block] | None = None
    recital_level = 0
    depth = 0  # Tiefe innerhalb zitierter (eingefügter) Texte in Änderungsrechtsakten

    for b in blocks:
        t = b.text if b.kind == "p" else ""
        quoted = track_quotes and depth > 0
        if track_quotes:
            depth = max(0, depth + b.full_text.count("„") - b.full_text.count("“"))

        if expect_title is not None:
            is_heading = bool(KAPITEL_RE.match(t) or ABSCHNITT_RE.match(t) or _looks_like_art_heading(b)
                              or ANHANG_RE.match(t))
            if not is_heading and b.kind == "p" and not b.enum and len(t) < 400:
                expect_title.append(b)
                expect_title = None
                continue
            expect_title = None

        if not quoted and state in ("final", "anhang") and b.kind == "table" and BACK_RE.match(b.text):
            state = "back"
        if not quoted and not b.enum and b.kind == "p" and state != "back":
            if state != "anhang" and (m := KAPITEL_RE.match(t)):
                cur_kap = Gliederung("kapitel", m.group(1), [b])
                act.kapitel.append(cur_kap)
                cur_abs = None
                expect_title = cur_kap.heading
                state, cur_unit = "enacting", None
                continue
            if state == "enacting" and (m := ABSCHNITT_RE.match(t)):
                cur_abs = Gliederung("abschnitt", m.group(1), [b], kapitel=cur_kap.nummer if cur_kap else None)
                act.abschnitte.append(cur_abs)
                expect_title = cur_abs.heading
                cur_unit = None
                continue
            if state in ("front", "recitals", "preamble_end", "enacting") and _looks_like_art_heading(b):
                m = ARTIKEL_RE.match(t)
                cur_unit = Unit("artikel", m.group(1), [b], kapitel=cur_kap.nummer if cur_kap else None,
                                abschnitt=cur_abs.nummer if cur_abs else None)
                act.articles.append(cur_unit)
                expect_title = cur_unit.heading
                state = "enacting"
                continue
            if state in ("enacting", "final", "anhang") and (m := ANHANG_RE.match(t)):
                cur_unit = Unit("anhang", m.group(1), [b])
                act.annexes.append(cur_unit)
                expect_title = cur_unit.heading
                state = "anhang"
                continue
            if state in ("final", "anhang") and BACK_RE.match(b.text):
                state = "back"
            elif state == "enacting" and SCHLUSS_RE.match(t):
                state = "final"

        if state in ("front", "recitals") and b.kind == "p" and not quoted and (m := RECITAL_RE.match(b.enum or "")) \
                and not ERLASSEN_RE.search(t):
            cur_unit = Unit("erwaegungsgrund", m.group(1), [], body=[b])
            act.recitals.append(cur_unit)
            recital_level = b.level
            state = "recitals"
            continue
        if state == "recitals":
            if ERLASSEN_RE.search(t) or b.level < recital_level:
                state = "preamble_end"
            else:
                cur_unit.body.append(b)
                continue

        if state == "front":
            act.front.append(b)
        elif state == "preamble_end":
            act.preamble_end.append(b)
        elif state in ("enacting", "anhang") and cur_unit is not None:
            cur_unit.body.append(b)
        elif state == "final":
            act.final.append(b)
        elif state == "back":
            act.back.append(b)
        else:
            act.unassigned.append(b)

    # Stand / Titel aus dem Vorspann
    for b in act.front[:40]:
        if m := re.search(r"\b\d{5}R\d{4}\s*[—-]\s*[A-Z]{2}\s*[—-]\s*(\d{2}\.\d{2}\.\d{4})", b.text):
            act.stand = m.group(1)
        if not act.titel and re.match(r"^(VERORDNUNG|DURCHFÜHRUNGSVERORDNUNG|DELEGIERTE VERORDNUNG)\b", b.text):
            act.titel = b.text
    return act


def parse(html: str, track_quotes: bool = False) -> ParsedAct:
    blocks, notes, _ = linearize(html)
    return parse_structure(blocks, notes, track_quotes=track_quotes)


def independent_heading_scan(html: str) -> dict[str, list[str]]:
    """Zweiter, vom Linearisierer unabhängiger Scan der Überschriften direkt im HTML (für die QS)."""
    soup = BeautifulSoup(html, "lxml")
    for n in soup.find_all(_is_note):
        n.decompose()
    res = {"kapitel": [], "artikel": [], "anhaenge": [], "eli_artikel": [], "eli_anhaenge": [], "eli_rct": []}
    for p in soup.find_all(["p", "div", "span", "td"]):
        if p.find(["p", "div", "table"]):
            continue
        t = normalize(p.get_text()).strip()
        if (m := KAPITEL_RE.match(t)):
            res["kapitel"].append(m.group(1))
        elif (m := ARTIKEL_RE.match(t)) and p.find_parent("table") is None:
            res["artikel"].append(m.group(1))
        elif (m := ANHANG_RE.match(t)):
            res["anhaenge"].append(m.group(1))
    for el in soup.find_all(id=True):
        i = el["id"]
        if (m := re.fullmatch(r"art_(\d+[a-z]*)", i)):
            res["eli_artikel"].append(m.group(1))
        elif (m := re.fullmatch(r"anx_([IVXLC]+[a-z]?)", i)):
            res["eli_anhaenge"].append(m.group(1))
        elif (m := re.fullmatch(r"rct_(\d+)", i)):
            res["eli_rct"].append(m.group(1))
    return res


def body_text_without_notes(html: str) -> str:
    """Gesamter sichtbarer Text des Dokuments ohne Fußnoten (Referenz für den Vollständigkeitsvergleich)."""
    soup = BeautifulSoup(html, "lxml")
    for n in soup.find_all(_is_note):
        n.decompose()
    for t in soup.find_all(SKIP_TAGS):
        t.decompose()
    for c in soup.find_all(string=lambda s: isinstance(s, Comment)):
        c.extract()
    root = soup.body or soup
    return root.get_text(" ")
