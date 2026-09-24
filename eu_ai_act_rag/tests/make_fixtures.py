"""Erzeugt SYNTHETISCHE Test-HTML-Dateien im Markup-Stil von EUR-Lex (Amtsblatt/ELI und konsolidiert).

Achtung: Die Texte sind Platzhalter ("Testtext …") und KEIN Gesetzestext. Sie dienen nur dazu,
Parser, Aufteilung, Word-Erzeugung und Qualitätsprüfung offline zu testen.
"""
from __future__ import annotations

import random
from pathlib import Path

LOREM = ("Testtext zur Strukturpruefung mit Umlauten äöü ß und Sonderzeichen „Anführung“ – Gedankenstrich; "
         "dieser Satz dient ausschliesslich als Platzhalter fuer amtlichen Wortlaut").split()

STRUCTURE = [  # (Kapitel, Titel, [(Abschnitt, Titel, von, bis)])
    ("I", "ALLGEMEINE BESTIMMUNGEN", [(None, None, 1, 4)]),
    ("II", "VERBOTENE PRAKTIKEN IM KI-BEREICH", [(None, None, 5, 5)]),
    ("III", "HOCHRISIKO-KI-SYSTEME", [("1", "Einstufung von KI-Systemen als Hochrisiko-KI-Systeme", 6, 7),
                                      ("2", "Anforderungen an Hochrisiko-KI-Systeme", 8, 15),
                                      ("3", "Pflichten der Anbieter und Betreiber", 16, 27),
                                      ("4", "Notifizierende Behörden und notifizierte Stellen", 28, 39),
                                      ("5", "Normen, Konformitätsbewertung, Bescheinigungen, Registrierung", 40, 49)]),
    ("IV", "TRANSPARENZPFLICHTEN FÜR ANBIETER UND BETREIBER BESTIMMTER KI-SYSTEME", [(None, None, 50, 50)]),
    ("V", "KI-MODELLE MIT ALLGEMEINEM VERWENDUNGSZWECK", [("1", "Einstufungsvorschriften", 51, 52),
                                                         ("2", "Pflichten für Anbieter", 53, 54),
                                                         ("3", "Pflichten bei systemischem Risiko", 55, 55),
                                                         ("4", "Praxisleitfäden", 56, 56)]),
    ("VI", "MASSNAHMEN ZUR INNOVATIONSFÖRDERUNG", [(None, None, 57, 63)]),
    ("VII", "GOVERNANCE", [("1", "Governance auf Unionsebene", 64, 69), ("2", "Zuständige nationale Behörden", 70, 70)]),
    ("VIII", "EU-DATENBANK FÜR HOCHRISIKO-KI-SYSTEME", [(None, None, 71, 71)]),
    ("IX", "BEOBACHTUNG NACH DEM INVERKEHRBRINGEN", [("1", "Beobachtung", 72, 72), ("2", "Meldung", 73, 73),
                                                    ("3", "Durchsetzung", 74, 84), ("4", "Rechtsbehelfe", 85, 87),
                                                    ("5", "Aufsicht GPAI", 88, 94)]),
    ("X", "VERHALTENSKODIZES UND LEITLINIEN", [(None, None, 95, 96)]),
    ("XI", "BEFUGNISÜBERTRAGUNG UND AUSSCHUSSVERFAHREN", [(None, None, 97, 98)]),
    ("XII", "SANKTIONEN", [(None, None, 99, 101)]),
    ("XIII", "SCHLUSSBESTIMMUNGEN", [(None, None, 102, 113)]),
]
ANNEXES = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII"]


def words(rng, n):
    return " ".join(rng.choice(LOREM) for _ in range(n)) + "."


# ------------------------------------------------------------------------------------------ OJ
def oj_list(items, rng):
    rows = "".join(f'<tr><td valign="top"><p class="oj-normal">{e}</p></td><td valign="top"><p class="oj-normal">{t}</p>{sub}</td></tr>'
                   for e, t, sub in items)
    return f'<table width="100%" border="0" cellspacing="0" cellpadding="0"><col width="4%"/><col width="96%"/><tbody>{rows}</tbody></table>'


def oj_article(n, rng, note_id=None, quoted=False):
    q = "„" if quoted else ""
    ref = (f' <a id="ntc{note_id}" href="#ntr{note_id}">(<span class="oj-super oj-note-tag">{note_id}</span>)</a>'
           if note_id else "")
    sub = oj_list([("i)", words(rng, 8), ""), ("ii)", words(rng, 6), "")], rng)
    lst = oj_list([(f"{c})", words(rng, 12), sub if c == "b" else "") for c in "abc"], rng)
    return (f'<div class="eli-subdivision" id="art_{n}"><p id="d1e{n}" class="oj-ti-art">{q}Artikel {n}</p>'
            f'<div class="eli-title" id="art_{n}.tit_1"><p class="oj-sti-art">Titel des Artikels {n}</p></div>'
            f'<div id="{n:03d}.001"><p class="oj-normal">(1)   {words(rng, 40)}{ref}</p></div>'
            f'<div id="{n:03d}.002"><p class="oj-normal">(2)   {words(rng, 20)}:</p>{lst}</div></div>')


def make_oj_original(path: Path, rng):
    notes = []
    parts = ['<html><head><title>L_202401689DE.000101.fmx.xml</title></head><body><div id="docHtml">'
             '<table width="100%" border="0" cellspacing="0" cellpadding="0" class="oj-table"><col width="50%"/><col width="20%"/><col width="30%"/>'
             '<tbody><tr><td><p class="oj-hd-date">Amtsblatt der Europäischen Union</p></td><td><p class="oj-hd-lg">DE</p></td>'
             '<td><p class="oj-hd-ti">Reihe L</p></td></tr></tbody></table>',
             '<div class="eli-container"><div class="eli-main-title" id="tit_1"><p class="oj-doc-ti">VERORDNUNG (EU) 2024/1689 DES EUROPÄISCHEN PARLAMENTS UND DES RATES</p>'
             '<p class="oj-doc-ti">vom 13. Juni 2024</p><p class="oj-doc-ti">zur Festlegung harmonisierter Vorschriften (Testfixture)</p></div>'
             '<div class="eli-subdivision" id="pbl_1"><p class="oj-normal">DAS EUROPÄISCHE PARLAMENT UND DER RAT DER EUROPÄISCHEN UNION —</p>'
             '<div class="eli-subdivision" id="cit_1"><p class="oj-normal">gestützt auf den Vertrag über die Arbeitsweise der Europäischen Union'
             ' <a id="ntc1-L" href="#ntr1-L">(<span class="oj-super oj-note-tag">1</span>)</a>,</p></div>'
             '<p class="oj-normal">in Erwägung nachstehender Gründe:</p>']
    notes.append('<p class="oj-note"><a id="ntr1-L" href="#ntc1-L">(<span class="oj-super">1</span>)</a>  ABl. C 1 vom 1.1.2024, S. 1 (Testfußnote).</p>')
    for i in range(1, 181):
        extra = ""
        if i == 7:
            extra = f' <a id="ntc2-L" href="#ntr2-L">(<span class="oj-super oj-note-tag">2</span>)</a>'
            notes.append('<p class="oj-note"><a id="ntr2-L" href="#ntc2-L">(<span class="oj-super">2</span>)</a>  Richtlinie 95/46/EG (Testfußnote).</p>')
        body = f'<p class="oj-normal">{words(rng, rng.randint(90, 260))}{extra}</p>'
        if i == 12:
            body += f'<p class="oj-normal">{words(rng, 30)}</p>'
        parts.append(f'<div class="eli-subdivision" id="rct_{i}"><table width="100%" border="0" cellspacing="0" cellpadding="0">'
                     f'<col width="4%"/><col width="96%"/><tbody><tr><td valign="top"><p class="oj-normal">({i})</p></td>'
                     f'<td valign="top">{body}</td></tr></tbody></table></div>')
    parts.append('<p class="oj-normal">HABEN FOLGENDE VERORDNUNG ERLASSEN:</p></div><div id="enc_1">')
    parts += oj_enacting(rng)
    parts.append('</div>')
    parts += oj_annexes(rng)
    parts.append('<hr class="oj-note"/>' + "".join(notes))
    parts.append('<table width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td><p class="oj-normal">ELI: http://data.europa.eu/eli/reg/2024/1689/oj</p>'
                 '<p class="oj-normal">ISSN 1977-0642 (electronic edition)</p></td></tr></tbody></table>')
    parts.append('</div></div></body></html>')
    path.write_text("".join(parts), encoding="utf-8")


def oj_enacting(rng):
    parts = []
    for kap, titel, secs in STRUCTURE:
        parts.append(f'<div id="cpt_{kap}"><p id="d1e{kap}" class="oj-ti-section-1">KAPITEL {kap}</p>'
                     f'<div class="eli-title" id="cpt_{kap}.tit_1"><p class="oj-ti-section-2">{titel}</p></div>')
        for sct, stitel, a, b in secs:
            if sct:
                parts.append(f'<div id="cpt_{kap}.sct_{sct}"><p class="oj-ti-section-1">ABSCHNITT {sct}</p>'
                             f'<div class="eli-title"><p class="oj-ti-section-2">{stitel}</p></div>')
            for n in range(a, b + 1):
                parts.append(oj_article(n, rng))
            if sct:
                parts.append("</div>")
        parts.append("</div>")
    parts.append('<p class="oj-normal">Diese Verordnung ist in allen ihren Teilen verbindlich und gilt unmittelbar in jedem Mitgliedstaat.</p>'
                 '<p class="oj-signatory">Geschehen zu Brüssel am 13. Juni 2024.</p>'
                 '<table width="100%"><tbody><tr><td><p class="oj-signatory">Im Namen des Europäischen Parlaments</p><p class="oj-signatory">Die Präsidentin</p></td>'
                 '<td><p class="oj-signatory">Im Namen des Rates</p><p class="oj-signatory">Der Präsident</p></td></tr></tbody></table>')
    return parts


def oj_annexes(rng):
    parts = []
    for a in ANNEXES:
        items = oj_list([(f"{k}.", words(rng, 60 if a != "III" else 900), "") for k in range(1, 9)], rng)
        parts.append(f'<div class="eli-subdivision" id="anx_{a}"><p class="oj-doc-ti" id="d1e{a}x">ANHANG {a}</p>'
                     f'<p class="oj-doc-ti">Titel des Anhangs {a}</p><p class="oj-normal">{words(rng, 20)}</p>{items}</div>')
    return parts


# ------------------------------------------------------------------------------------------ Konsolidiert
def cons_article(n, rng):
    grid = "".join(f'<div class="grid-container grid-list"><div class="list grid-list-column-1"><span>{c})</span></div>'
                   f'<div class="grid-list-column-2"><p class="norm">{words(rng, 14)}</p></div></div>' for c in "abc")
    sup = ' mehr als 10<sup>25</sup> Gleitkommaoperationen' if n == 51 else ""
    return (f'<div class="eli-subdivision" id="art_{n}"><p class="title-article-norm">Artikel {n}</p>'
            f'<div class="eli-title"><p class="stitle-article-norm">Titel des Artikels {n}</p></div>'
            f'<div class="norm"><span class="no-parag">(1)   </span><div class="inline-element"><p class="inline-element norm">{words(rng, 60)}{sup}</p></div></div>'
            f'<p class="modref">▼M1</p>'
            f'<div class="norm"><span class="no-parag">(2)   </span><div class="inline-element"><p class="inline-element norm">►M1 {words(rng, 25)} ◄ Folgendes:</p>{grid}</div></div></div>')


def make_consolidated(path: Path, rng):
    parts = ['<html><body><div id="document1">'
             '<p class="hd-date">02024R1689 — DE — 27.07.2026 — 001.001</p>'
             '<p class="disclaimer">Dieser Text dient lediglich zu Informationszwecken und hat keine Rechtswirkung.</p>'
             '<p class="title-doc-first">►B VERORDNUNG (EU) 2024/1689 DES EUROPÄISCHEN PARLAMENTS UND DES RATES vom 13. Juni 2024 (Testfixture)</p>'
             '<p class="hd-modifiers">Geändert durch:</p>'
             '<table class="hd-modifiers"><tbody><tr><td></td><td></td><th colspan="3">Amtsblatt</th></tr>'
             '<tr><td></td><td></td><th>Nr.</th><th>Seite</th><th>Datum</th></tr>'
             '<tr><td>►M1</td><td>Verordnung (EU) 2026/1744 (Testeintrag)</td><td>L 1744</td><td>1</td><td>24.7.2026</td></tr></tbody></table>']
    for kap, titel, secs in STRUCTURE:
        parts.append(f'<p class="modref">▼B</p><p class="title-division-1">KAPITEL {kap}</p><p class="title-division-2">{titel}</p>')
        for sct, stitel, a, b in secs:
            if sct:
                parts.append(f'<p class="title-division-1">ABSCHNITT {sct}</p><p class="title-division-2">{stitel}</p>')
            for n in range(a, b + 1):
                parts.append(cons_article(n, rng))
                if n == 4:
                    parts.append('<p class="modref">▼M1</p>' + cons_article(4, rng).replace("art_4", "art_4a")
                                 .replace("Artikel 4<", "Artikel 4a<").replace("Artikels 4<", "Artikels 4a<"))
    parts.append('<p class="norm">Diese Verordnung ist in allen ihren Teilen verbindlich und gilt unmittelbar in jedem Mitgliedstaat.</p>')
    for a in ANNEXES:
        body = "".join(f'<div class="grid-container grid-list"><div class="grid-list-column-1"><span>{k}.</span></div>'
                       f'<div class="grid-list-column-2"><p class="norm">{words(rng, 60 if a != "III" else 900)}</p></div></div>'
                       for k in range(1, 9))
        table = ('<table><tbody><tr><th>Nummer</th><th>Kategorie</th></tr><tr><td>1</td><td>Testkategorie A</td></tr>'
                 '<tr><td>2</td><td>Testkategorie B</td></tr></tbody></table>') if a == "I" else ""
        parts.append(f'<div class="eli-subdivision" id="anx_{a}"><p class="title-annex-1">ANHANG {a}</p>'
                     f'<p class="title-annex-2">Titel des Anhangs {a}</p>{table}{body}</div>')
    parts.append("</div></body></html>")
    path.write_text("".join(parts), encoding="utf-8")


# ------------------------------------------------------------------------------------------ Omnibus
def make_omnibus(path: Path, rng):
    parts = ['<html><body><div class="eli-container"><div class="eli-main-title" id="tit_1">'
             '<p class="oj-doc-ti">VERORDNUNG (EU) 2026/1744 DES EUROPÄISCHEN PARLAMENTS UND DES RATES</p>'
             '<p class="oj-doc-ti">vom 15. Juli 2026</p><p class="oj-doc-ti">zur Änderung der Verordnung (EU) 2024/1689 (Testfixture)</p></div>'
             '<div class="eli-subdivision" id="pbl_1"><p class="oj-normal">in Erwägung nachstehender Gründe:</p>']
    for i in range(1, 41):
        parts.append(f'<div class="eli-subdivision" id="rct_{i}"><table width="100%"><col width="4%"/><col width="96%"/><tbody><tr>'
                     f'<td valign="top"><p class="oj-normal">({i})</p></td><td valign="top"><p class="oj-normal">{words(rng, 150)}</p></td></tr></tbody></table></div>')
    parts.append('<p class="oj-normal">HABEN FOLGENDE VERORDNUNG ERLASSEN:</p></div><div id="enc_1">')
    points = []
    for k in range(1, 44):
        inner = ""
        if k == 3:
            inner = (f'<p class="oj-ti-art">„Artikel 4a</p><p class="oj-sti-art">Eingefügter Testartikel</p>'
                     f'<p class="oj-normal">{words(rng, 40)}“</p>')
        if k == 20:
            inner = (f'<p class="oj-ti-section-1">„ABSCHNITT 6</p><p class="oj-ti-section-2">Eingefügter Abschnitt</p>'
                     f'<p class="oj-ti-art">Artikel 60a</p><p class="oj-sti-art">Testtitel</p><p class="oj-normal">{words(rng, 30)}“</p>')
        points.append((f"{k}.", words(rng, 180), inner))
    parts.append('<div class="eli-subdivision" id="art_1"><p class="oj-ti-art">Artikel 1</p><div class="eli-title"><p class="oj-sti-art">Änderung der Verordnung (EU) 2024/1689</p></div>'
                 '<p class="oj-normal">Die Verordnung (EU) 2024/1689 wird wie folgt geändert:</p>' + oj_list(points, rng) + '</div>')
    for n, t in ((2, "Änderung der Verordnung (EU) 2018/1139"), (3, "Änderung der Verordnung (EU) 2023/1230"), (4, "Inkrafttreten")):
        parts.append(f'<div class="eli-subdivision" id="art_{n}"><p class="oj-ti-art">Artikel {n}</p><div class="eli-title"><p class="oj-sti-art">{t}</p></div>'
                     f'<p class="oj-normal">{words(rng, 60)}</p></div>')
    parts.append('<p class="oj-normal">Diese Verordnung ist in allen ihren Teilen verbindlich und gilt unmittelbar in jedem Mitgliedstaat.</p>'
                 '<p class="oj-signatory">Geschehen zu Straßburg am 15. Juli 2026.</p></div></div></body></html>')
    path.write_text("".join(parts), encoding="utf-8")


def make_all(target: Path, seed: int = 1689):
    target.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)
    make_oj_original(target / "32024R1689.html", rng)
    make_consolidated(target / "02024R1689-20260727.html", rng)
    make_omnibus(target / "32026R1744.html", rng)


if __name__ == "__main__":
    make_all(Path(__file__).parent / "fixtures")
