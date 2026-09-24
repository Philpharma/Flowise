"""Zentrale Konfiguration: amtliche Quellen, Erwartungswerte für die Qualitätsprüfung, Größenziele.

Es werden ausschließlich EUR-Lex-Dokumente (eur-lex.europa.eu) als Quelle für Gesetzestext verwendet.
"""
from dataclasses import dataclass, field

EURLEX_HTML_URL = "https://eur-lex.europa.eu/legal-content/{lang}/TXT/HTML/?uri=CELEX:{celex}"
EURLEX_PAGE_URL = "https://eur-lex.europa.eu/legal-content/{lang}/TXT/?uri=CELEX:{celex}"
LANG = "DE"

# Hosts für Gesetzestext: EUR-Lex; Cellar des Amts für Veröffentlichungen nur als technischer Rückfall.
CELLAR_URL = "https://publications.europa.eu/resource/celex/{celex}"
ALLOWED_LAW_HOSTS = ("eur-lex.europa.eu", "publications.europa.eu")
ALLOWED_GUIDELINE_HOST_SUFFIX = ".europa.eu"


@dataclass(frozen=True)
class Source:
    key: str
    celex: str
    kurzname: str
    beschreibung: str
    art: str  # "konsolidiert" | "amtsblatt"
    verwendung: tuple = field(default_factory=tuple)

    def url(self, lang: str = LANG) -> str:
        return EURLEX_HTML_URL.format(lang=lang, celex=self.celex)

    def page_url(self, lang: str = LANG) -> str:
        return EURLEX_PAGE_URL.format(lang=lang, celex=self.celex)


SOURCES = {
    "konsolidiert": Source(
        key="konsolidiert",
        celex="02024R1689-20260727",
        kurzname="VO (EU) 2024/1689 – konsolidierte Fassung",
        beschreibung="Verordnung (EU) 2024/1689 (KI-Verordnung), konsolidierte Fassung, Stand 27.07.2026",
        art="konsolidiert",
        verwendung=("artikel", "anhaenge"),
    ),
    "original": Source(
        key="original",
        celex="32024R1689",
        kurzname="VO (EU) 2024/1689 – Amtsblattfassung",
        beschreibung="Verordnung (EU) 2024/1689 (KI-Verordnung), ursprüngliche Fassung, ABl. L, 12.7.2024",
        art="amtsblatt",
        verwendung=("erwaegungsgruende",),
    ),
    "omnibus": Source(
        key="omnibus",
        celex="32026R1744",
        kurzname="VO (EU) 2026/1744 – Digital-Omnibus KI",
        beschreibung="Verordnung (EU) 2026/1744 (Digital-Omnibus-Verordnung zur KI), Amtsblattfassung",
        art="amtsblatt",
        verwendung=("erwaegungsgruende", "aenderungen"),
    ),
}

# Erwartungswerte für die Qualitätsprüfung (Struktur der Verordnung (EU) 2024/1689).
# Eingefügte Artikel/Anhänge (z. B. "4a") werden automatisch erkannt und im Bericht ausgewiesen.
EXPECTED = {
    "konsolidiert": {
        "kapitel": ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII"],
        "artikel_von": 1,
        "artikel_bis": 113,
        "anhaenge": ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII"],
    },
    "original": {"erwaegungsgruende": 180},
    "omnibus": {"erwaegungsgruende": None},  # Anzahl wird aus der Quelle ermittelt; Lückenlosigkeit wird geprüft
}

# Größensteuerung: grobe Schätzung ~3.300 Zeichen je Word-Seite bei 11 pt / A4 / 2 cm Rand.
ZEICHEN_PRO_SEITE = 3300
MAX_SEITEN = 15
ZIEL_SEITEN = 12
MIN_SEITEN = 5
MAX_ZEICHEN = MAX_SEITEN * ZEICHEN_PRO_SEITE
ZIEL_ZEICHEN = ZIEL_SEITEN * ZEICHEN_PRO_SEITE

ORDNER = {
    "artikel": "01_EU_KI_Verordnung",
    "erwaegungsgruende": "02_Erwaegungsgruende",
    "anhaenge": "03_Anhaenge",
    "aenderungen": "04_Aenderungsverordnungen",
    "leitlinien": "05_EU_Kommission_Leitlinien",
}

USER_AGENT = "eu-ai-act-rag-builder/1.0 (+Python requests; Abruf amtlicher EUR-Lex-Texte)"
