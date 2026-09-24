"""Aufteilung der geparsten Rechtsakte in Word-Dokumente passender Größe (strukturgeführt)."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import config
from docx_writer import DocSpec, Section
from eurlex_parser import Block, Gliederung, ParsedAct, Unit
from fetch import FetchedDocument


@dataclass
class PlannedDoc:
    spec: DocSpec
    dokumenttyp: str
    ordner: str
    kapitel: str = ""
    abschnitt: str = ""
    artikel: list[str] = field(default_factory=list)
    erwaegungsgruende: list[str] = field(default_factory=list)
    anhang: str = ""
    celex: str = ""
    fassung: str = ""
    quelle_key: str = ""
    units: list[Unit] = field(default_factory=list)
    qa_status: str = "offen"
    qa_details: list[str] = field(default_factory=list)
    seiten: str = ""

    @property
    def zeichen(self) -> int:
        return sum(len(b.full_text) for b in self.spec.content_blocks())


def slug(text: str, maxlen: int = 60) -> str:
    t = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("Ä", "Ae").replace("Ö", "Oe") \
        .replace("Ü", "Ue").replace("ß", "ss")
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9\-]+", "_", t).strip("_")
    t = re.sub(r"_+", "_", t)
    if len(t) > maxlen:
        t = t[:maxlen].rsplit("_", 1)[0]
    return t


def title_case_de(text: str) -> str:
    """Nur für Dateinamen: amtliche Überschriften in VERSALIEN lesbar machen (Text im Dokument bleibt unverändert)."""
    if not text.isupper():
        return text
    small = {"und", "oder", "für", "von", "mit", "der", "die", "das", "des", "den", "dem", "in", "im", "zur", "zum",
             "auf", "an", "bei", "zu", "über", "unter", "nach", "sowie", "einschließlich", "gegen", "aus"}

    def part(p: str) -> str:
        return p if len(p) <= 2 else p[:1] + p[1:].lower()

    out = []
    for i, w in enumerate(text.split(" ")):
        lw = w.lower()
        out.append(lw if i and lw in small else "-".join(part(p) for p in w.split("-")))
    return " ".join(out)


def fassung_label(fd: FetchedDocument, act: ParsedAct) -> str:
    src = fd.source
    if src.art == "konsolidiert":
        m = re.search(r"-(\d{4})(\d{2})(\d{2})$", src.celex)
        stand = act.stand or (f"{m.group(3)}.{m.group(2)}.{m.group(1)}" if m else "")
        return f"Konsolidierte Fassung, Stand {stand}"
    m = re.search(r"vom (\d{1,2}\. \w+ \d{4})", act.titel or "")
    return f"Amtsblattfassung (ursprünglicher Rechtsakt{', vom ' + m.group(1) if m else ''})"


def metadata(fd: FetchedDocument, act: ParsedAct, inhalt: str) -> list[tuple[str, str]]:
    return [
        ("Quelldokument", fd.source.beschreibung),
        ("CELEX-Nummer", fd.source.celex),
        ("Fassungsstand", fassung_label(fd, act)),
        ("Quelle (EUR-Lex)", fd.source.page_url()),
        ("Abgerufen am", f"{fd.abgerufen_am} ({fd.herkunft}); SHA-256 des HTML: {fd.sha256[:16]}…"),
        ("Enthaltener Inhalt", inhalt),
        ("Hinweis", "Amtlicher Wortlaut, automatisiert und unverändert aus EUR-Lex übernommen; EUR-Lex-"
                    "Änderungsmarker (▼/►/◄) entfernt. Rechtsverbindlich ist nur die im Amtsblatt veröffentlichte Fassung."),
    ]


def notes_for(blocks: list[Block], act: ParsedAct) -> list[Block]:
    keys: list[str] = []
    for b in blocks:
        for k in b.notes:
            if k not in keys:
                keys.append(k)
    return [act.notes[k] for k in keys if k in act.notes]


def pack(units: list, size, max_chars: int, target: int) -> list[list]:
    """Aufeinanderfolgende Einheiten gierig zu Gruppen ≤ max_chars bündeln (Einheiten werden nie geteilt)."""
    groups, cur, cur_size = [], [], 0
    for u in units:
        s = size(u)
        if cur and cur_size + s > target:
            groups.append(cur)
            cur, cur_size = [], 0
        cur.append(u)
        cur_size += s
    if cur:
        groups.append(cur)
    # sehr kleine Restgruppe an die vorherige anhängen, wenn Obergrenze eingehalten wird
    if len(groups) > 1 and sum(size(u) for u in groups[-1]) < config.MIN_SEITEN * config.ZEICHEN_PRO_SEITE \
            and sum(size(u) for u in groups[-2] + groups[-1]) <= max_chars:
        groups[-2].extend(groups.pop())
    return groups


def split_blocks(blocks: list[Block], max_chars: int, target: int) -> list[list[Block]]:
    """Überlange Einheit (z. B. Anhang) an Gliederungspunkten der obersten Ebene teilen."""
    if sum(len(b.full_text) for b in blocks) <= max_chars:
        return [blocks]
    base = min((b.level for b in blocks if b.enum), default=0)
    segments: list[list[Block]] = []
    for b in blocks:
        if not segments or (b.enum and b.level == base):
            segments.append([])
        segments[-1].append(b)
    groups = pack(segments, lambda seg: sum(len(x.full_text) for x in seg), max_chars, target)
    return [[b for seg in g for b in seg] for g in groups]


def _range(nums: list[str]) -> str:
    return nums[0] if len(nums) == 1 else f"{nums[0]}-{nums[-1]}"


# ------------------------------------------------------------------------------------------------
def plan_articles(fd: FetchedDocument, act: ParsedAct, out: Path) -> list[PlannedDoc]:
    ordner = config.ORDNER["artikel"]
    docs: list[PlannedDoc] = []
    n = 0
    fassung = fassung_label(fd, act)

    if act.front:
        n += 1
        spec = DocSpec(out / ordner / f"{n:02d}_Titel_und_Fassungsinformation.docx",
                       "Titel und Fassungsinformation der Verordnung (EU) 2024/1689",
                       metadata(fd, act, "Titel, Vorspann und Änderungsübersicht der konsolidierten Fassung"),
                       [Section("blocks", blocks=act.front)], notes_for(act.front, act),
                       core_subject=fd.source.celex, core_keywords="KI-Verordnung; Titel; Fassung")
        docs.append(PlannedDoc(spec, "Titel/Fassungsinformation", ordner, celex=fd.source.celex, fassung=fassung,
                               quelle_key=fd.source.key))

    kap_by_nr = {k.nummer: k for k in act.kapitel}
    abs_by_key = {(a.kapitel, a.nummer): a for a in act.abschnitte}
    # Einheiten: (Kapitel, Abschnitt) -> Artikel in Quellreihenfolge
    groups: list[tuple[str | None, list[str | None], list[Unit]]] = []
    for a in act.articles:
        if groups and groups[-1][0] == a.kapitel and groups[-1][1][-1] == a.abschnitt:
            groups[-1][2].append(a)
        else:
            groups.append((a.kapitel, [a.abschnitt], [a]))
    # Benachbarte kleine Abschnitte desselben Kapitels zusammenfassen (Kapitelgrenzen bleiben erhalten)
    min_chars = config.MIN_SEITEN * config.ZEICHEN_PRO_SEITE
    size = lambda arts: sum(u.chars for u in arts)  # noqa: E731
    merged: list[tuple[str | None, list[str | None], list[Unit]]] = []
    for g in groups:
        if (merged and merged[-1][0] == g[0] and size(merged[-1][2]) < min_chars and size(g[2]) < min_chars
                and size(merged[-1][2]) + size(g[2]) <= config.ZIEL_ZEICHEN):
            merged[-1][1].extend(g[1])
            merged[-1][2].extend(g[2])
        else:
            merged.append((g[0], list(g[1]), list(g[2])))

    seen_kap, seen_abs = set(), set()
    for kap_nr, abs_nrs, arts in merged:
        kap = kap_by_nr.get(kap_nr)
        abs_list = [abs_by_key[(kap_nr, x)] for x in abs_nrs if (kap_nr, x) in abs_by_key]
        parts = pack(arts, lambda u: u.chars, config.MAX_ZEICHEN, config.ZIEL_ZEICHEN)
        for pi, part in enumerate(parts, 1):
            n += 1
            is_last_group = part[-1] is act.articles[-1]
            sections: list[Section] = []
            if kap:
                if kap_nr not in seen_kap:
                    sections.append(Section("heading", 1, kap.heading))
                    seen_kap.add(kap_nr)
                else:
                    sections.append(Section("context", 1, text=" – ".join(b.text for b in kap.heading) + " (Fortsetzung)"))
            cur_abs = object()
            for a in part:
                if a.abschnitt != cur_abs:
                    cur_abs = a.abschnitt
                    ab = abs_by_key.get((kap_nr, a.abschnitt))
                    if ab is not None:
                        if (kap_nr, ab.nummer) not in seen_abs:
                            sections.append(Section("heading", 2, ab.heading))
                            seen_abs.add((kap_nr, ab.nummer))
                        else:
                            sections.append(Section("context", 2, text=" – ".join(b.text for b in ab.heading) + " (Fortsetzung)"))
                sections.append(Section("heading", 3, a.heading))
                sections.append(Section("blocks", blocks=a.body))
            if is_last_group and act.final:
                sections.append(Section("context", 3, text="Schlussformel"))
                sections.append(Section("blocks", blocks=act.final))
            nums = [a.nummer for a in part]
            part_abs = [x for x in abs_list if any(a.abschnitt == x.nummer for a in part)]
            name_parts = [f"{n:02d}"]
            titel_parts = []
            if kap:
                name_parts.append(f"Kapitel_{kap.nummer}")
                titel_parts.append(f"Kapitel {kap.nummer} – {kap.titel}")
            if len(part_abs) == 1:
                ab = part_abs[0]
                name_parts.append(f"Abschnitt_{ab.nummer}")
                titel_parts.append(f"Abschnitt {ab.nummer} – {ab.titel}")
                topic = ab.titel
            else:
                if part_abs:
                    name_parts.append(f"Abschnitt_{part_abs[0].nummer}-{part_abs[-1].nummer}")
                    titel_parts.append(f"Abschnitte {part_abs[0].nummer}–{part_abs[-1].nummer}")
                topic = kap.titel if kap else ""
            if topic:
                name_parts.append(slug(title_case_de(topic), 55))
            name_parts.append(f"Art_{_range(nums)}")
            if len(parts) > 1:
                name_parts.append(f"Teil_{pi}")
            titel = " / ".join(titel_parts) + f" (Art. {_range(nums)})"
            blocks_all = [b for s in sections for b in s.blocks]
            spec = DocSpec(out / ordner / ("_".join(name_parts) + ".docx"), titel,
                           metadata(fd, act, f"Artikel {', '.join(nums)}" + (" sowie Schlussformel" if is_last_group and act.final else "")),
                           sections, notes_for(blocks_all, act), core_subject=fd.source.celex,
                           core_keywords=f"KI-Verordnung; Artikel {_range(nums)}")
            docs.append(PlannedDoc(spec, "Artikel (verfügender Teil)", ordner,
                                   kapitel=f"Kapitel {kap.nummer} – {kap.titel}" if kap else "",
                                   abschnitt=" | ".join(f"Abschnitt {x.nummer} – {x.titel}" for x in part_abs),
                                   artikel=nums, celex=fd.source.celex, fassung=fassung, quelle_key=fd.source.key,
                                   units=part))
    return docs


def plan_annexes(fd: FetchedDocument, act: ParsedAct, out: Path) -> list[PlannedDoc]:
    ordner = config.ORDNER["anhaenge"]
    fassung = fassung_label(fd, act)
    docs = []
    for i, anx in enumerate(act.annexes, 1):
        parts = split_blocks(anx.body, config.MAX_ZEICHEN, config.ZIEL_ZEICHEN)
        for pi, part in enumerate(parts, 1):
            sections = [Section("heading", 1, anx.heading) if pi == 1 else
                        Section("context", 1, text=" – ".join(b.text for b in anx.heading) + f" (Fortsetzung, Teil {pi})"),
                        Section("blocks", blocks=part)]
            name = f"Anhang_{anx.nummer}_{slug(title_case_de(anx.titel), 60)}"
            if len(parts) > 1:
                name += f"_Teil_{pi}"
            titel = f"Anhang {anx.nummer} – {anx.titel}" + (f" (Teil {pi} von {len(parts)})" if len(parts) > 1 else "")
            blocks_all = [b for s in sections for b in s.blocks]
            spec = DocSpec(out / ordner / f"{i:02d}_{name}.docx", titel,
                           metadata(fd, act, f"Anhang {anx.nummer}" + (f", Teil {pi} von {len(parts)}" if len(parts) > 1 else " (vollständig)")),
                           sections, notes_for(blocks_all, act), core_subject=fd.source.celex,
                           core_keywords=f"KI-Verordnung; Anhang {anx.nummer}")
            docs.append(PlannedDoc(spec, "Anhang", ordner, anhang=anx.nummer, celex=fd.source.celex, fassung=fassung,
                                   quelle_key=fd.source.key, units=[anx] if pi == 1 else []))
    return docs


def plan_recitals(fd: FetchedDocument, act: ParsedAct, out: Path, subdir: str, prefix: str,
                  label: str, include_front: bool = True) -> list[PlannedDoc]:
    ordner = f"{config.ORDNER['erwaegungsgruende']}/{subdir}"
    fassung = fassung_label(fd, act)
    docs = []
    width = 3
    if include_front and act.front:
        blocks = act.front
        spec = DocSpec(out / ordner / f"{prefix}000_Titel_und_Bezugsvermerke.docx", f"{label} – Titel und Bezugsvermerke",
                       metadata(fd, act, "Titel und Bezugsvermerke (vor den Erwägungsgründen)"),
                       [Section("blocks", blocks=blocks)], notes_for(blocks, act),
                       core_subject=fd.source.celex, core_keywords=f"{label}; Bezugsvermerke")
        docs.append(PlannedDoc(spec, "Präambel (Bezugsvermerke)", ordner, celex=fd.source.celex, fassung=fassung,
                               quelle_key=fd.source.key))
    groups = pack(act.recitals, lambda u: u.chars, config.MAX_ZEICHEN, config.ZIEL_ZEICHEN)
    for gi, g in enumerate(groups):
        von, bis = g[0].nummer, g[-1].nummer
        sections: list[Section] = [Section("context", 1, text=f"{label} – Erwägungsgründe ({von}) bis ({bis})")]
        blocks = [b for u in g for b in u.body]
        sections.append(Section("blocks", blocks=blocks))
        if gi == len(groups) - 1 and act.preamble_end:
            sections.append(Section("blocks", blocks=act.preamble_end))
        blocks_all = [b for s in sections for b in s.blocks]
        spec = DocSpec(out / ordner / f"{prefix}Erwaegungsgruende_{int(von):0{width}d}-{int(bis):0{width}d}.docx",
                       f"{label} – Erwägungsgründe {von} bis {bis}",
                       metadata(fd, act, f"Erwägungsgründe ({von}) bis ({bis})" +
                                (" sowie Einleitungsformel des verfügenden Teils" if gi == len(groups) - 1 and act.preamble_end else "")),
                       sections, notes_for(blocks_all, act), core_subject=fd.source.celex,
                       core_keywords=f"{label}; Erwägungsgründe {von}-{bis}")
        docs.append(PlannedDoc(spec, "Erwägungsgründe", ordner, erwaegungsgruende=[u.nummer for u in g],
                               celex=fd.source.celex, fassung=fassung, quelle_key=fd.source.key, units=g))
    return docs


def plan_amending_act(fd: FetchedDocument, act: ParsedAct, out: Path, prefix: str, label: str) -> list[PlannedDoc]:
    """Verfügender Teil (und ggf. Anhänge) einer Änderungsverordnung; Erwägungsgründe separat in 02_."""
    ordner = config.ORDNER["aenderungen"]
    fassung = fassung_label(fd, act)
    docs = []
    n = 0
    for a in act.articles:
        parts = split_blocks(a.body, config.MAX_ZEICHEN, config.ZIEL_ZEICHEN)
        for pi, part in enumerate(parts, 1):
            n += 1
            sections = [Section("heading", 3, a.heading) if pi == 1 else
                        Section("context", 3, text=" – ".join(b.text for b in a.heading) + f" (Fortsetzung, Teil {pi})"),
                        Section("blocks", blocks=part)]
            last = a is act.articles[-1] and pi == len(parts)
            if last and act.final:
                sections += [Section("context", 3, text="Schlussformel"), Section("blocks", blocks=act.final)]
            pts = [b.enum for b in part if b.enum and b.level == min((x.level for x in part if x.enum), default=0)]
            teil = f"_Teil_{pi}_Nr_{pts[0].rstrip('.')}-{pts[-1].rstrip('.')}" if len(parts) > 1 and pts else (f"_Teil_{pi}" if len(parts) > 1 else "")
            name = f"{prefix}{n:02d}_Art_{a.nummer}_{slug(title_case_de(a.titel), 50)}{teil}.docx"
            blocks_all = [b for s in sections for b in s.blocks]
            spec = DocSpec(out / ordner / name, f"{label} – Artikel {a.nummer} {a.titel}" + (f" (Teil {pi} von {len(parts)})" if len(parts) > 1 else ""),
                           metadata(fd, act, f"Artikel {a.nummer}" + (f", Teil {pi} von {len(parts)}" if len(parts) > 1 else "")),
                           sections, notes_for(blocks_all, act), core_subject=fd.source.celex,
                           core_keywords=f"{label}; Artikel {a.nummer}")
            docs.append(PlannedDoc(spec, "Änderungsverordnung – verfügender Teil", ordner, artikel=[a.nummer],
                                   celex=fd.source.celex, fassung=fassung, quelle_key=fd.source.key,
                                   units=[a] if pi == 1 else []))
    for anx in act.annexes:
        n += 1
        sections = [Section("heading", 1, anx.heading), Section("blocks", blocks=anx.body)]
        spec = DocSpec(out / ordner / f"{prefix}{n:02d}_Anhang_{anx.nummer}.docx", f"{label} – Anhang {anx.nummer}",
                       metadata(fd, act, f"Anhang {anx.nummer}"), sections, notes_for(anx.blocks, act),
                       core_subject=fd.source.celex, core_keywords=f"{label}; Anhang {anx.nummer}")
        docs.append(PlannedDoc(spec, "Änderungsverordnung – Anhang", ordner, anhang=anx.nummer, celex=fd.source.celex,
                               fassung=fassung, quelle_key=fd.source.key, units=[anx]))
    return docs
