"""Offline-Tests mit synthetischen EUR-Lex-ähnlichen Fixtures (kein Gesetzestext)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

import build_corpus  # noqa: E402
import qa  # noqa: E402
from eurlex_parser import linearize, normalize, parse  # noqa: E402
from make_fixtures import make_all  # noqa: E402


@pytest.fixture(scope="session")
def fixtures(tmp_path_factory):
    d = tmp_path_factory.mktemp("fixtures")
    make_all(d)
    return d


def run(fixtures, tmp_path, html_dir=None):
    out = tmp_path / "out"
    rc = build_corpus.main(["--html-dir", str(html_dir or fixtures), "--ohne-leitlinien", "--out", str(out),
                            "--cache", str(tmp_path / "cache")])
    return rc, out


def test_normalize_removes_only_markers_and_whitespace():
    assert normalize("►M1 Text mit  ◄ Rest") .split() == ["Text", "mit", "Rest"]
    assert normalize("Anbie­ter") == "Anbieter"
    assert normalize("Absatz (1) Buchstabe a)") == "Absatz (1) Buchstabe a)"


def test_list_tables_and_grid_lists_keep_enumeration():
    html = ('<html><body><p class="oj-ti-art">Artikel 1</p><p>Titel</p><table><tr><td><p>a)</p></td><td><p>erstens</p>'
            '<table><tr><td><p>i)</p></td><td><p>unter</p></td></tr></table></td></tr>'
            '<tr><td><p>b)</p></td><td><p>zweitens</p></td></tr></table>'
            '<div class="grid-container grid-list"><div class="grid-list-column-1"><span>c)</span></div>'
            '<div class="grid-list-column-2"><p>drittens</p></div></div>'
            '<div class="norm"><span class="no-parag">(2)  </span><div><p>Absatz</p></div></div></body></html>')
    blocks, _, _ = linearize(html)
    got = [(b.enum, b.text, b.level) for b in blocks]
    assert ("a)", "erstens", 1) in got and ("i)", "unter", 2) in got and ("b)", "zweitens", 1) in got
    assert ("c)", "drittens", 1) in got and ("(2)", "Absatz", 0) in got


def test_full_run_passes_qa(fixtures, tmp_path):
    rc, out = run(fixtures, tmp_path)
    assert rc == 0
    assert (out / "00_Dokumentenuebersicht.xlsx").exists() and (out / "00_Dokumentenuebersicht.csv").exists()
    csv_text = (out / "00_Dokumentenuebersicht.csv").read_text(encoding="utf-8-sig")
    assert ";FEHLER;" not in csv_text
    names = [p.name for p in out.rglob("*.docx")]
    assert any(n.startswith("02_Kapitel_I_") for n in names)
    assert sum(n.startswith("Erwaegungsgruende_") for n in names) >= 6
    assert any("VO_2026-1744_Erwaegungsgruende" in n for n in names)
    assert len([n for n in names if re.match(r"\d\d_Anhang_", n)]) >= 13


def test_quoted_insertions_in_amending_act_are_not_structure(fixtures):
    act = parse((fixtures / "32026R1744.html").read_text(encoding="utf-8"), track_quotes=True)
    assert [a.nummer for a in act.articles] == ["1", "2", "3", "4"]
    assert len(act.recitals) == 40


def test_missing_article_is_detected(fixtures, tmp_path):
    broken = tmp_path / "broken"
    broken.mkdir()
    for f in fixtures.iterdir():
        (broken / f.name).write_text(f.read_text(encoding="utf-8"), encoding="utf-8")
    p = broken / "02024R1689-20260727.html"
    html = p.read_text(encoding="utf-8")
    start = html.index('<div class="eli-subdivision" id="art_7">')
    end = html.index('<div class="eli-subdivision" id="art_8">')
    p.write_text(html[:start] + html[end:], encoding="utf-8")
    rc, out = run(fixtures, tmp_path, broken)
    assert rc == 1
    report = (out / "00_Qualitaetsbericht.md").read_text(encoding="utf-8")
    assert "fehlend: ['7']" in report


def test_altered_word_text_is_detected(fixtures, tmp_path):
    rc, out = run(fixtures, tmp_path)
    assert rc == 0
    from docx import Document
    import config
    import planner
    from fetch import fetch
    fd = fetch(config.SOURCES["konsolidiert"], tmp_path / "c", fixtures)
    act = parse(fd.html)
    docs = planner.plan_articles(fd, act, out)
    target = docs[3]
    d = Document(target.spec.path)
    para = next(p for p in d.paragraphs if p.style.name == "Normal" and len(p.text) > 50)
    para.runs[-1].text = para.runs[-1].text[:-10]  # "abgeschnitten"
    d.save(target.spec.path)
    res = qa.QAResult()
    qa.check_docx(res, target)
    assert res.fehler and "weicht" in res.fehler[0].details


def test_guideline_pdf_lines_are_joined_without_changing_text():
    import leitlinien
    from eurlex_parser import squash
    lines = ["1. Introduction", "The AI Act lays down harmonised rules for the placing on the market of",
             "AI systems and sets specific requirements for high-", "risk systems.", "", "(a) first item"]
    paras = leitlinien.lines_to_paragraphs(lines)
    assert paras == ["1. Introduction", "The AI Act lays down harmonised rules for the placing on the market of "
                     "AI systems and sets specific requirements for high-risk systems.", "(a) first item"]
    assert squash("".join(paras)) == squash("".join(lines))


def test_guideline_link_selection_prefers_german():
    import leitlinien
    html = (b'<a href="https://ec.europa.eu/newsroom/dae/redirection/document/1">Guidelines EN</a>'
            b'<a href="https://ec.europa.eu/newsroom/dae/redirection/document/2">Leitlinien DE</a>'
            b'<a href="https://example.com/x.pdf">fremd</a>')
    links = leitlinien.find_pdf_links(html, "https://digital-strategy.ec.europa.eu/en/library/x")
    assert len(links) == 2 and leitlinien.choose_pdf(links)[0] == "Leitlinien DE"
