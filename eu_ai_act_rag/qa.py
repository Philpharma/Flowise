"""Automatisierte Qualitätsprüfung der erzeugten Dokumente gegen die EUR-Lex-Quellen."""
from __future__ import annotations

import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import config
from docx_writer import read_back
from eurlex_parser import ParsedAct, body_text_without_notes, independent_heading_scan, squash
from planner import PlannedDoc

ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}
END_OK = re.compile(r"(?:[.;:!?)\]»“\"’—]|gestrichen\)?|\(aufgehoben\)|entfällt)$")


def roman(s: str) -> int:
    s = re.sub(r"[a-z]$", "", s)
    total = 0
    for i, ch in enumerate(s):
        v = ROMAN[ch]
        total += -v if i + 1 < len(s) and ROMAN[s[i + 1]] > v else v
    return total


def art_key(n: str) -> tuple[int, str]:
    m = re.match(r"(\d+)([a-z]*)", n)
    return int(m.group(1)), m.group(2)


@dataclass
class Check:
    bereich: str
    pruefung: str
    ok: bool
    details: str = ""
    schwere: str = "FEHLER"  # FEHLER | WARNUNG


@dataclass
class QAResult:
    checks: list[Check] = field(default_factory=list)

    def add(self, bereich, pruefung, ok, details="", schwere="FEHLER"):
        self.checks.append(Check(bereich, pruefung, bool(ok), details, schwere))

    @property
    def fehler(self):
        return [c for c in self.checks if not c.ok and c.schwere == "FEHLER"]

    @property
    def warnungen(self):
        return [c for c in self.checks if not c.ok and c.schwere == "WARNUNG"]


def _dups(seq):
    seen, d = set(), []
    for x in seq:
        if x in seen and x not in d:
            d.append(x)
        seen.add(x)
    return d


def _first_diff(a: str, b: str) -> str:
    i = next((k for k, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    return f"erste Abweichung bei Zeichen {i}: Quelle …{a[max(0, i - 40):i + 40]}… / Ausgabe …{b[max(0, i - 40):i + 40]}…"


def check_completeness(res: QAResult, name: str, html: str, act: ParsedAct):
    """Linearisierung verlustfrei? Vergleich des gesamten Quelltexts (ohne Fußnoten) mit allen Blöcken."""
    src = squash(body_text_without_notes(html))
    parsed = squash("".join(
        (b.enum + b.text) if b.kind == "p" else "".join(c.text for r in b.rows for c in r) for b in act.blocks))
    res.add(name, "Quelltext vollständig und in Originalreihenfolge erfasst (zeichengenauer Vergleich ohne Leerraum)",
            src == parsed, "" if src == parsed else f"Quelle {len(src)} / erfasst {len(parsed)} Zeichen; " + _first_diff(src, parsed))


def check_consolidated(res: QAResult, html: str, act: ParsedAct, docs: list[PlannedDoc]):
    exp = config.EXPECTED["konsolidiert"]
    name = "Konsolidierte Fassung"
    scan = independent_heading_scan(html)

    kap = [k.nummer for k in act.kapitel]
    missing = [k for k in exp["kapitel"] if k not in kap]
    res.add(name, "Alle Kapitel vorhanden", not missing, f"gefunden: {', '.join(kap)}" + (f"; fehlend: {missing}" if missing else ""))
    res.add(name, "Keine doppelten Kapitel", not _dups(kap), str(_dups(kap)))
    res.add(name, "Kapitelreihenfolge aufsteigend", kap == sorted(kap, key=roman))
    res.add(name, "Kapitelüberschriften identisch mit unabhängigem HTML-Scan", kap == scan["kapitel"],
            f"Scan: {scan['kapitel']}")
    for k in act.kapitel:
        res.add(name, f"Kapitel {k.nummer} hat Titel", len(k.heading) >= 2 and k.titel.strip())

    arts = [a.nummer for a in act.articles]
    base = {str(i) for i in range(exp["artikel_von"], exp["artikel_bis"] + 1)}
    missing = sorted(base - set(arts), key=art_key)
    extra = sorted(set(arts) - base, key=art_key)
    res.add(name, f"Alle Artikel {exp['artikel_von']}–{exp['artikel_bis']} vorhanden (keine fehlenden Nummern)", not missing,
            f"{len(arts)} Artikel erfasst" + (f"; fehlend: {missing}" if missing else ""))
    res.add(name, "Zusätzliche (eingefügte) Artikel", True, ", ".join(extra) if extra else "keine", "WARNUNG")
    res.add(name, "Keine doppelten Artikel", not _dups(arts), str(_dups(arts)))
    res.add(name, "Artikelreihenfolge wie in EUR-Lex (aufsteigend)", arts == sorted(arts, key=art_key))
    res.add(name, "Artikelfolge identisch mit unabhängigem HTML-Scan", arts == scan["artikel"],
            "" if arts == scan["artikel"] else f"Scan: {scan['artikel'][:10]}… ({len(scan['artikel'])})")
    if scan["eli_artikel"]:
        res.add(name, "Artikelfolge identisch mit ELI-Kennungen (art_*)", arts == scan["eli_artikel"],
                "" if arts == scan["eli_artikel"] else f"ELI: {len(scan['eli_artikel'])} Einträge")
    for a in act.articles:
        body_txt = [b.full_text for b in a.body if b.full_text.strip()]
        if not a.titel.strip():
            res.add(name, f"Artikel {a.nummer}: Überschrift vorhanden", False, schwere="WARNUNG")
        if not body_txt:
            ok = "gestrichen" in a.titel.lower()
            res.add(name, f"Artikel {a.nummer}: Text vorhanden", ok, "leer")
        elif not END_OK.search(body_txt[-1].rstrip()):
            res.add(name, f"Artikel {a.nummer}: Textende plausibel (nicht abgeschnitten)", False,
                    f"endet mit: …{body_txt[-1][-60:]}", "WARNUNG")

    anx = [a.nummer for a in act.annexes]
    missing = [a for a in exp["anhaenge"] if a not in anx]
    res.add(name, "Alle Anhänge vorhanden", not missing, f"gefunden: {', '.join(anx)}" + (f"; fehlend: {missing}" if missing else ""))
    res.add(name, "Keine doppelten Anhänge", not _dups(anx), str(_dups(anx)))
    res.add(name, "Anhangreihenfolge aufsteigend", anx == sorted(anx, key=lambda x: (roman(x), x)))
    res.add(name, "Anhangfolge identisch mit unabhängigem HTML-Scan", anx == scan["anhaenge"], f"Scan: {scan['anhaenge']}")
    for a in act.annexes:
        body_txt = [b.full_text for b in a.body if b.full_text.strip()]
        res.add(name, f"Anhang {a.nummer}: Titel und Inhalt vorhanden", bool(a.titel.strip()) and bool(body_txt))

    check_units_vs_eli(res, name, html, act.articles, "art", "Artikel")
    check_units_vs_eli(res, name, html, act.annexes, "anx", "Anhänge")
    check_completeness(res, name, html, act)
    res.add(name, "Keine Blöcke ohne Zuordnung", not act.unassigned,
            "; ".join(b.full_text[:80] for b in act.unassigned[:5]))
    check_assignment(res, name, act, docs, required=act.front + [b for a in act.articles for b in a.blocks]
                     + act.final + [b for a in act.annexes for b in a.blocks])


def check_recitals(res: QAResult, name: str, html: str, act: ParsedAct, docs: list[PlannedDoc], expected: int | None,
                   extra_required=()):
    nums = [int(u.nummer) for u in act.recitals]
    n = expected or (max(nums) if nums else 0)
    missing = [i for i in range(1, n + 1) if i not in nums]
    res.add(name, f"Erwägungsgründe vollständig (1–{n})", bool(nums) and not missing and (expected is None or max(nums) == expected),
            f"{len(nums)} erfasst" + (f"; fehlend: {missing[:20]}" if missing else ""))
    res.add(name, "Keine doppelten Erwägungsgründe", not _dups(nums), str(_dups(nums)))
    res.add(name, "Reihenfolge der Erwägungsgründe lückenlos aufsteigend", nums == list(range(1, len(nums) + 1)))
    scan = independent_heading_scan(html)
    if scan["eli_rct"]:
        res.add(name, "Erwägungsgründe identisch mit ELI-Kennungen (rct_*)", [str(x) for x in nums] == scan["eli_rct"],
                f"ELI: {len(scan['eli_rct'])}")
    for u in act.recitals:
        last = [b.full_text for b in u.body if b.full_text.strip()][-1]
        if not END_OK.search(last.rstrip()):
            res.add(name, f"Erwägungsgrund {u.nummer}: Textende plausibel", False, f"…{last[-60:]}", "WARNUNG")
    check_units_vs_eli(res, name, html, act.recitals, "rct", "Erwägungsgründe")
    check_completeness(res, name, html, act)
    check_assignment(res, name, act, docs, required=act.front + [b for u in act.recitals for b in u.body]
                     + act.preamble_end + list(extra_required))


def check_articles_vs_eli(res: QAResult, name: str, html: str, act: ParsedAct):
    scan = independent_heading_scan(html)
    arts = [a.nummer for a in act.articles]
    if scan["eli_artikel"]:
        res.add(name, "Artikelfolge identisch mit ELI-Kennungen (art_*)", arts == scan["eli_artikel"],
                f"erfasst: {arts}; ELI: {scan['eli_artikel']}")
    check_units_vs_eli(res, name, html, act.articles, "art", "Artikel")
    res.add(name, "Schlussformel erkannt (Ende des letzten Artikels korrekt abgegrenzt)", bool(act.final))


def check_units_vs_eli(res: QAResult, name: str, html: str, units, prefix: str, label: str):
    """Unabhängiger Vergleich je Einheit mit dem Text des zugehörigen ELI-Containers (<div id="art_5"> usw.)."""
    from bs4 import BeautifulSoup
    from eurlex_parser import _is_note, doc_root
    soup = doc_root(BeautifulSoup(html, "lxml"))
    for n in soup.find_all(_is_note):
        n.decompose()
    checked, bad = 0, []
    for u in units:
        el = soup.find(id=f"{prefix}_{u.nummer}")
        if el is None:
            continue
        checked += 1
        mine = squash("".join((b.enum + b.text) if b.kind == "p" else "".join(c.text for r in b.rows for c in r)
                              for b in u.blocks))
        if squash(el.get_text()) != mine:
            bad.append(u.nummer)
    if checked:
        res.add(name, f"{label}: Text je Einheit identisch mit ELI-Container ({prefix}_*), {checked} geprüft",
                not bad, f"abweichend: {bad[:20]}")


def check_assignment(res: QAResult, name: str, act: ParsedAct, docs: list[PlannedDoc], required):
    counts: dict[int, int] = {}
    for d in docs:
        for b in d.spec.content_blocks():
            counts[b.idx] = counts.get(b.idx, 0) + 1
    missing = [b for b in required if counts.get(b.idx, 0) == 0]
    double = [b for b in required if counts.get(b.idx, 0) > 1]
    res.add(name, "Jeder Textblock genau einem Word-Dokument zugeordnet (nichts verloren, nichts doppelt)",
            not missing and not double,
            (f"{len(missing)} fehlen, z. B.: {missing[0].full_text[:80]!r}; " if missing else "") +
            (f"{len(double)} doppelt" if double else ""))


def expected_doc_text(d: PlannedDoc) -> str:
    parts = []
    for b in d.spec.content_blocks():
        parts.append((b.enum + b.text) if b.kind == "p" else "".join(c.text for r in b.rows for c in r))
    return squash("".join(parts))


def check_docx(res: QAResult, d: PlannedDoc):
    back = read_back(d.spec.path)
    got = squash("".join(back["content"]))
    exp = expected_doc_text(d)
    problems = []
    if got != exp:
        problems.append("Word-Inhalt weicht von der Quelle ab: " + _first_diff(exp, got))
    exp_notes = sorted(squash(n.enum + n.text) for n in d.spec.notes)
    if sorted(squash(n) for n in back["notes"]) != exp_notes:
        problems.append("Fußnoten weichen ab")
    if any(ch in "".join(back["content"]) for ch in "▼►◄"):
        problems.append("EUR-Lex-Änderungsmarker im Text verblieben")
    if not any(m.startswith("CELEX-Nummer:") for m in back["meta"]):
        problems.append("Metadaten (CELEX) fehlen")
    heads = [b.text for b in d.spec.content_blocks() if b.kind == "p" and re.match(r"^Artikel \d+[a-z]*$", b.text)]
    if heads != [f"Artikel {a}" for a in d.artikel] and d.dokumenttyp.startswith("Artikel"):
        problems.append(f"Artikelüberschriften/Reihenfolge abweichend: {heads}")
    ok = not problems
    res.add(d.spec.path.name, "Word-Datei wortlautgetreu (Rücklesen und Vergleich), Metadaten, Fußnoten", ok, "; ".join(problems))
    size_pages = d.zeichen / config.ZEICHEN_PRO_SEITE
    if size_pages > config.MAX_SEITEN * 1.15:
        d.qa_details.append(f"Umfang ca. {size_pages:.0f} Seiten (> Ziel), strukturbedingt nicht weiter teilbar")
    d.qa_details.extend(problems)
    d.qa_status = "OK" if ok else "FEHLER"


def count_pages(paths: list[Path]) -> dict[Path, int]:
    """Echte Seitenzahl über LibreOffice (optional)."""
    result = {}
    with tempfile.TemporaryDirectory() as tmp:
        for i in range(0, len(paths), 20):
            chunk = paths[i:i + 20]
            try:
                subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, *map(str, chunk)],
                               check=True, capture_output=True, timeout=600)
            except (OSError, subprocess.SubprocessError):
                return result
        def reader(p):  # ohne Zusatzbibliothek: Seitenobjekte im PDF zählen
            return len(re.findall(rb"/Type\s*/Page(?!s)", Path(p).read_bytes()))
        for p in paths:
            pdf = Path(tmp) / (p.stem + ".pdf")
            if pdf.exists():
                result[p] = reader(pdf)
    return result
