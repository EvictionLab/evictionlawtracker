#!/usr/bin/env python3
"""Match LSC statutory citations to individual sections of historical XML codes.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import html
import re
import sqlite3
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator, Sequence

try:
    from openpyxl import load_workbook
except ImportError as exc:
    raise SystemExit("This script requires openpyxl: pip install openpyxl") from exc


REQUIRED_COLUMNS = (
    "Jurisdiction",
    "Question",
    "Citation/Statute",
    "XML Title",
    "Chapter",
    "Section",
)

REVIEW_COLUMNS = (
    "XML File",
    "XML Year",
    "Matched ShortName",
    "Match Status",
    "Match Score",
    "Years Matched",
    "Candidates",
    "Match Notes",
)

LONG_REPORT_COLUMNS = (
    "Workbook Row",
    "Jurisdiction",
    "Question",
    "Position",
    "Citation",
    "Citation Type",
    "Section Key",
    "Year",
    "Status",
    "Score",
    "Matched ShortName",
    "XML Heading",
    "XML File",
    "XML Title",
    "Chapter",
    "Hierarchy",
    "Content Hash",
    "Changed From Prior Year",
    "Candidates",
    "Notes",
)

JURISDICTION_CODES = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "District of Columbia": "DC", "Florida": "FL", "Georgia": "GA", "Hawaii": "HI",
    "Idaho": "ID", "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", "Kansas": "KS",
    "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK",
    "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY", "American Samoa": "AS",
    "Federated States of Micronesia": "FM", "Guam": "GU", "Marshall Islands": "MH",
    "Northern Mariana Islands": "MP", "Palau": "PW", "Puerto Rico": "PR",
    "U.S. Virgin Islands": "VI",
}

# A section number: starts with a digit, may contain letters, dots, dashes,
# colons, commas (Kan. 58-25,120) and slashes (ILCS 5/9-209).
NUM_TOKEN = r"[0-9][0-9A-Za-z.\-:,/\u2013]*[0-9A-Za-z]|[0-9]"
TRAILING_NUMBER_RE = re.compile(r"(" + NUM_TOKEN + r")\s*$")
ILCS_SHORTNAME_RE = re.compile(r"^(\d+)\s+ILCS\s+(\S+)\s*$", re.IGNORECASE)

CITATION_RE = re.compile(
    # 735 Ill. Comp. Stat. 5/9-209
    r"(?P<ilcs_ch>\d+)\s+Ill\.?\s*Comp\.?\s*Stat\.?\s*(?P<ilcs_num>\d+/[0-9A-Za-z.\-]*[0-9A-Za-z])"
    # § 35-9A-141 / §§ ... / Art. 4701 / Rule 13
    r"|(?:§+|\bArt(?:icle)?\.|\bRule)\s*(?P<num>" + NUM_TOKEN + r")"
    # Fla. Stat. 83.46 / N.J. Stat 2A:18-61.1 (no section sign)
    r"|\b(?:Stat|Code|Laws)\.?,?\s+(?P<bare_num>\d[0-9A-Za-z.\-:]*[0-9A-Za-z])",
    re.IGNORECASE,
)

NON_STATUTE_RE = re.compile(
    r"executive order|\bE\.O\.|declaration|state of emergency|admin\.? code|"
    r"\bR\.\s*Civ\.\s*P\.|civil form|\ban act\b|\bNo\.\s*\d+\.",
    re.IGNORECASE,
)

QUALIFIER_STOPWORDS = {
    "code", "codes", "stat", "stats", "statutes", "rev", "revised", "gen", "general",
    "laws", "law", "the", "and", "ann", "annotated", "cons", "uncons", "codified",
    "cent", "comp", "compiled", "sec", "section", "tit", "title", "art", "article",
    "for", "chapter",
}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def unescape_all(value: str) -> str:
    """Undo (possibly repeated) HTML escaping, e.g. '&amp;amp;' -> '&'."""
    text = value or ""
    for _ in range(3):
        unescaped = html.unescape(text)
        if unescaped == text:
            break
        text = unescaped
    return text


def strip_markup(value: str) -> str:
    return normalize_text(re.sub(r"<[^>]+>", " ", unescape_all(value)))


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].lower()


def normalize_section_number(value: str) -> str:
    """Canonical section key: lower case, unified dashes, no subsections,
    no leading zeros at the start of the number or after - : / separators
    (09.45.100 -> 9.45.100; 1-01-01 -> 1-1-1). Zeros after '.' are kept
    because they are significant in decimal-style numbering (83.056 vs 83.56)."""
    text = unescape_all(value).lower()
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"(?:\([a-z0-9]+\))+$", "", text)
    text = text.strip(".,;:-")
    text = re.sub(r"(^|[-:/])0+(?=\d)", r"\1", text)
    return text


def normalize_heading(value: str) -> str:
    text = strip_markup(value).lower()
    text = text.replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    text = re.sub(r"[\u2013\u2014]", "-", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return normalize_text(text)


def heading_similarity(left: str, right: str) -> float:
    a, b = normalize_heading(left), normalize_heading(right)
    if not a or not b:
        return 0.0
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    ta, tb = set(a.split()), set(b.split())
    jaccard = len(ta & tb) / len(ta | tb) if ta | tb else 0.0
    # Workbook headings sometimes carry extra clauses, so containment counts too.
    contained = 1.0 if (len(a) > 8 and len(b) > 8 and (a in b or b in a)) else 0.0
    return max(ratio, jaccard, 0.9 * contained)


NUMBER_TOKEN_RE = re.compile(r"\b\d+[a-z]*(?:-[a-z0-9]+)*\b")


def qualifier_tokens(prefix: str) -> tuple[set[str], set[str]]:
    """Split a citation/ShortName prefix into title/chapter numbers and
    code-name words: 'Del Code tit. 25' -> ({'25'}, set());
    'Md. Code, Real Property' -> (set(), {'real', 'property'}).
    All-caps abbreviations (RPA, CCP) are left to code_abbreviations()."""
    text = unescape_all(prefix).replace("&", " ")
    text = re.sub(r"\b[A-Z]{2,5}\b(?!\.)", " ", text)
    lower = text.lower()
    numbers = set(NUMBER_TOKEN_RE.findall(lower))
    words = {
        word
        for word in re.findall(r"[a-z]+", NUMBER_TOKEN_RE.sub(" ", lower))
        if len(word) > 2 and word not in QUALIFIER_STOPWORDS
    }
    return numbers, words


def words_match(left: str, right: str) -> bool:
    """'civ' ~ 'civil', 'prop' ~ 'property'; 'prop' !~ 'procedure'."""
    return min(len(left), len(right)) >= 3 and (left.startswith(right) or right.startswith(left))


def drop_jurisdiction_words(words: set[str], jurisdiction: str) -> set[str]:
    """Remove state-name abbreviations ('cal', 'mass', 'tex', 'virginia')."""
    name_words = re.findall(r"[a-z]+", jurisdiction.lower())
    return {w for w in words if not any(n.startswith(w) or w.startswith(n) for n in name_words)}


def file_qualifiers(path: str) -> tuple[set[str], set[str]]:
    """Code-name words and abbreviations from the XML file name, e.g.
    'Civil Code - CIV.xml' -> ({'civil'}, {'civ'})."""
    stem = Path(path).stem
    _, words = qualifier_tokens(stem)
    abbreviations = {t.lower() for t in re.findall(r"\b[A-Z]{2,5}\b", stem)}
    return words, abbreviations


def prefix_initials(prefix: str) -> tuple[str, str]:
    """Initials and joined letters of a citation prefix, used to recognize
    code abbreviations in ShortNames: 'N.Y. Real Prop. Acts. Law' ->
    ('nyrpal', 'nyrealpropactslaw'); 'La. Code of Civil Proc.' -> ('lccp', ...)."""
    words = [w for w in re.findall(r"[a-z]+", unescape_all(prefix).lower()) if w not in {"of", "the", "and"}]
    return "".join(w[0] for w in words), "".join(words)


def is_subsequence(short: str, long: str) -> bool:
    remaining = iter(long)
    return all(ch in remaining for ch in short)


def code_abbreviations(prefix: str, jurisdiction: str) -> set[str]:
    """All-caps tokens in a ShortName prefix that look like code
    abbreviations ('NY RPA Law' -> {'rpa'}; 'LA CCP Art.' -> {'ccp'}),
    excluding the jurisdiction's own postal code."""
    state = JURISDICTION_CODES.get(jurisdiction, "").lower()
    return {
        token.lower()
        for token in re.findall(r"\b[A-Z]{2,5}\b", unescape_all(prefix))
        if token.lower() != state and token not in {"ILCS", "SDCL"}
    }


def split_identifier(value: str) -> tuple[str, str]:
    """Split a ShortName/Caption into (prefix, section_key).
    'Ala. Code 35-9A-141' -> ('Ala. Code', '35-9a-141');
    '735 ILCS 5/9-209'    -> ('735 ILCS', '5/9-209')."""
    text = normalize_text(unescape_all(value))
    if not text:
        return "", ""
    match = ILCS_SHORTNAME_RE.match(text)
    if match:
        return f"{match.group(1)} ILCS", normalize_section_number(match.group(2))
    match = TRAILING_NUMBER_RE.search(text)
    if not match:
        return text, ""
    prefix = text[: match.start()].strip()
    prefix = re.sub(r"(?:§+|&sect;|\bSec\.?|\bSection|\bArt\.?)\s*$", "", prefix, flags=re.IGNORECASE)
    return prefix.strip(" ,"), normalize_section_number(match.group(1))


# ---------------------------------------------------------------------------
# Years
# ---------------------------------------------------------------------------

def list_year_dirs(jurisdiction_dir: Path) -> list[tuple[int, Path]]:
    if not jurisdiction_dir.is_dir():
        return []
    years = [
        (int(path.name), path)
        for path in jurisdiction_dir.iterdir()
        if path.is_dir() and re.fullmatch(r"(?:19|20)\d{2}", path.name)
    ]
    return sorted(years)


def detect_year(target_year: int, xml_path: Path) -> tuple[int, Path] | None:
    """
    Use exact or prior year matching to find the closest year to the target year.
    ``xml_path`` is the jurisdiction's directory (xml_root / Jurisdiction).
    Returns (resolved_year, year_directory) or None when no folder is at or
    before the target year.
    """
    prior = [(year, path) for year, path in list_year_dirs(Path(xml_path)) if year <= target_year]
    return prior[-1] if prior else None


# ---------------------------------------------------------------------------
# Citations
# ---------------------------------------------------------------------------

def classify_citation(text: str, has_number: bool, prefix: str) -> str:
    if NON_STATUTE_RE.search(text):
        return "other"
    if not has_number:
        return "comprehensive" if re.search(r"\b(act|chapter|ch\.)\b", text, re.IGNORECASE) else "other"
    if re.search(r"\bcode\b", prefix, re.IGNORECASE):
        return "code"
    return "statute"


def parse_citation_cell(cell: object) -> list[dict]:
    """Split one workbook cell into citations. Parts without any digit
    ('Fees.') are treated as the continuation of the previous heading.
    Duplicate citations within a cell are dropped."""
    text = normalize_text(cell)
    if not text:
        return []
    parts: list[str] = []
    for part in re.split(r";\s*", text):
        part = part.strip()
        if not part:
            continue
        if parts and not re.search(r"\d", part):
            parts[-1] = f"{parts[-1]}; {part}"
        else:
            parts.append(part)

    citations = []
    seen = set()
    for part in parts:
        match = CITATION_RE.search(part)
        if match and match.group("ilcs_ch"):
            number = match.group("ilcs_num")
            prefix = f"{match.group('ilcs_ch')} ILCS"
        elif match and match.group("num"):
            number = match.group("num")
            prefix = part[: match.start()]
        elif match and match.group("bare_num"):
            number = match.group("bare_num")
            prefix = part[: match.start("bare_num")]
        else:
            number, prefix = "", part
        heading = part[match.end():] if match else ""
        heading = heading.strip(" .,:;-\u2013\u2014")
        prefix = normalize_text(prefix).strip(" ,")
        key = normalize_section_number(number) if number else ""
        citation_type = classify_citation(part, bool(key), prefix)
        if citation_type == "other":
            key = ""
        dedupe = (key, prefix.lower()) if key else ("", part.lower())
        if dedupe in seen:
            continue
        seen.add(dedupe)
        numbers, words = qualifier_tokens(prefix)
        initials, letters = prefix_initials(prefix)
        citations.append({
            "citation_raw": part,
            "cite_number": number,
            "section_key": key,
            "cite_normalized": normalize_text(f"{re.sub(r'[§,]', ' ', prefix)} {number}") if key else "",
            "heading": heading,
            "prefix": prefix,
            "citation_type": citation_type,
            "position": len(citations) + 1,
            "qual_numbers": numbers,
            "qual_words": words,
            "qual_initials": initials,
            "qual_letters": letters,
        })
    return citations


def detect_citation(workbook_path, sheet: str | None = None) -> list[dict]:
    """
    Detect the citations in the workbook. One cell may contain several
    citations; each becomes its own dict with the keys
        row, jurisdiction, question, citation_raw, cite_number,
        cite_normalized (prefix + number, e.g. 'Ala. Code 35-9A-141'),
        section_key (normalized number used for matching, e.g. '35-9a-141'),
        heading (LSC's heading text), prefix (code / title / chapter text),
        citation_type (code, statute, comprehensive, other),
        position (order within the cell), qual_numbers, qual_words.
    """
    workbook = load_workbook(workbook_path, read_only=True)
    worksheet = workbook[sheet] if sheet else workbook.active
    rows = worksheet.iter_rows(values_only=True)
    headers = {normalize_text(name): index for index, name in enumerate(next(rows)) if name}
    output = []
    for row_number, values in enumerate(rows, start=2):
        jurisdiction = normalize_text(values[headers["Jurisdiction"]])
        if not jurisdiction:
            continue
        question = normalize_text(values[headers["Question"]])
        for citation in parse_citation_cell(values[headers["Citation/Statute"]]):
            citation.update(row=row_number, jurisdiction=jurisdiction, question=question)
            output.append(citation)
    return output


# ---------------------------------------------------------------------------
# Section index
# ---------------------------------------------------------------------------

INDEX_SCHEMA = """
CREATE TABLE IF NOT EXISTS files (
    path TEXT PRIMARY KEY, jurisdiction TEXT, year INTEGER,
    mtime REAL, size INTEGER, n_sections INTEGER, error TEXT
);
CREATE TABLE IF NOT EXISTS sections (
    id INTEGER PRIMARY KEY,
    path TEXT, jurisdiction TEXT, year INTEGER,
    title TEXT, chapter TEXT, hierarchy TEXT,
    caption TEXT, heading TEXT, shortname TEXT, prefix TEXT, section_key TEXT,
    content_hash TEXT, content_length INTEGER, revision_history TEXT
);
CREATE INDEX IF NOT EXISTS ix_sections_key ON sections (jurisdiction, section_key);
CREATE INDEX IF NOT EXISTS ix_sections_year ON sections (jurisdiction, year);
CREATE INDEX IF NOT EXISTS ix_sections_path ON sections (path);
"""

SECTION_FIELDS = (
    "path", "jurisdiction", "year", "title", "chapter", "hierarchy", "caption",
    "heading", "shortname", "prefix", "section_key", "content_hash",
    "content_length", "revision_history",
)


def iter_index_nodes(path: Path) -> Iterator[tuple[dict, list[str]]]:
    """Stream <Index> nodes, yielding (fields, ancestor_captions)."""
    stack: list[dict] = []
    wanted = {"caption", "description", "shortname", "content", "revisionhistory"}
    for event, element in ET.iterparse(path, events=("start", "end")):
        tag = local_name(element.tag)
        if event == "start":
            if tag == "index":
                stack.append({})
            continue
        if tag in wanted and stack:
            stack[-1][tag] = element.text or ""
            element.clear()
        elif tag == "index" and stack:
            node = stack.pop()
            yield node, [normalize_text(unescape_all(a.get("caption", ""))) for a in stack]
            element.clear()


def section_from_node(node: dict, ancestors: Sequence[str]) -> dict | None:
    shortname = normalize_text(unescape_all(node.get("shortname", "")))
    caption = normalize_text(unescape_all(node.get("caption", "")))
    prefix, key = split_identifier(shortname)
    if not key:
        _, key = split_identifier(caption)
    if not key or not (shortname or node.get("content")):
        return None
    content = strip_markup(node.get("content", ""))
    heading = strip_markup(node.get("description", ""))
    history = strip_markup(node.get("revisionhistory", ""))
    comparable = comparable_text(content, caption, heading, history)
    ancestors = [a for a in ancestors if a]
    chapter = next(
        (a for a in ancestors if re.match(r"(?i)(chapter|ch\.)\b", a)),
        ancestors[1] if len(ancestors) > 1 else "",
    )
    return {
        "title": ancestors[0] if ancestors else "",
        "chapter": chapter,
        "hierarchy": " > ".join(ancestors),
        "caption": caption,
        "heading": heading,
        "shortname": shortname,
        "prefix": prefix,
        "section_key": key,
        "content_hash": hashlib.sha1(comparable.encode("utf-8")).hexdigest()[:16] if comparable else "",
        "content_length": len(content),
        "revision_history": history,
    }


def comparable_text(content: str, caption: str, heading: str, history: str) -> str:
    """Section text reduced to what matters for change detection: no
    session-law history (stored separately), no repeated caption/heading at
    the start, and only lower-case words and numbers, so formatting-only
    edits do not count as a change in the law."""
    text = content
    if history and history in text:
        text = text.replace(history, " ")
    else:
        # The history usually closes the text as a parenthetical: (Acts 2006, No. 316, ...)
        text = re.sub(r"\((?:Acts|Code|Laws|Added|Amended|L\.)\b[^()]*\)\s*$", " ", text)
    tokens = re.findall(r"[a-z0-9]+", text.lower())
    for lead in (caption, heading):
        lead_tokens = re.findall(r"[a-z0-9]+", lead.lower())
        if lead_tokens and tokens[: len(lead_tokens)] == lead_tokens:
            tokens = tokens[len(lead_tokens):]
    return " ".join(tokens)


def open_index(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(str(db_path))
    connection.row_factory = sqlite3.Row
    connection.executescript(INDEX_SCHEMA)
    return connection


def index_jurisdiction(
    connection: sqlite3.Connection,
    xml_root: Path,
    jurisdiction: str,
    rebuild: bool = False,
) -> None:
    """Index all XML files under xml_root/<jurisdiction>/<year>/, skipping
    files whose path, mtime and size are unchanged since the last run."""
    files = [
        (year, path)
        for year, year_dir in list_year_dirs(xml_root / jurisdiction)
        for path in sorted(year_dir.glob("*.xml"))
    ]
    known = {
        row["path"]: (row["mtime"], row["size"])
        for row in connection.execute(
            "SELECT path, mtime, size FROM files WHERE jurisdiction = ?", (jurisdiction,)
        )
    }
    current = {str(path.relative_to(xml_root)) for _, path in files}
    for stale in set(known) - current:
        connection.execute("DELETE FROM sections WHERE path = ?", (stale,))
        connection.execute("DELETE FROM files WHERE path = ?", (stale,))

    todo = []
    for year, path in files:
        stat = path.stat()
        relative = str(path.relative_to(xml_root))
        if rebuild or known.get(relative) != (stat.st_mtime, stat.st_size):
            todo.append((year, path, relative, stat))
    if not todo:
        return

    insert = (
        f"INSERT INTO sections ({', '.join(SECTION_FIELDS)}) "
        f"VALUES ({', '.join('?' for _ in SECTION_FIELDS)})"
    )
    for number, (year, path, relative, stat) in enumerate(todo, start=1):
        connection.execute("DELETE FROM sections WHERE path = ?", (relative,))
        rows, error = [], ""
        try:
            for node, ancestors in iter_index_nodes(path):
                section = section_from_node(node, ancestors)
                if section:
                    section.update(path=relative, jurisdiction=jurisdiction, year=year)
                    rows.append(tuple(section[name] for name in SECTION_FIELDS))
        except (ET.ParseError, OSError) as exc:
            error = str(exc)
        connection.executemany(insert, rows)
        connection.execute(
            "INSERT OR REPLACE INTO files VALUES (?, ?, ?, ?, ?, ?, ?)",
            (relative, jurisdiction, year, stat.st_mtime, stat.st_size, len(rows), error),
        )
        if number % 25 == 0 or number == len(todo):
            connection.commit()
            print(f"  {jurisdiction}: indexed {number:,}/{len(todo):,} files", file=sys.stderr)


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

@dataclass
class Candidate:
    section: dict
    score: int
    reasons: list[str] = field(default_factory=list)

    def label(self) -> str:
        return f"{self.section['shortname']} [{self.score}] {self.section['heading'][:60]}"


@dataclass
class YearMatch:
    year: int
    status: str  # MATCHED, AMBIGUOUS, LOW_SCORE, HEADING_MATCH, NOT_FOUND
    score: int
    section: dict | None
    candidates: list[Candidate]

    @property
    def accepted(self) -> bool:
        return self.status == "MATCHED"


def score_section(citation: dict, section: dict, unique_in_year: bool) -> Candidate:
    """Score one section that shares the citation's section number.

    base 60
    title/chapter numbers: +20, or -40 if they contradict
    code-name words shared with the ShortName or file name: +5 each (max 10)
    ShortName code words the citation lacks ('Proc.' for a Civil Code cite): -5 each (max 15)
    file name names exactly the cited code: +10
    code abbreviation (RPA, CCP, CIV): +15 exact initials / +8 loose
    heading similarity: up to +20, or -15 when both headings exist and disagree
    only section with this number in the year and nothing contradicts it: +15"""
    jurisdiction = citation["jurisdiction"]
    score, reasons = 60, ["section number"]
    section_numbers, section_words = qualifier_tokens(section["prefix"])
    section_words = drop_jurisdiction_words(section_words, jurisdiction)
    cite_words = drop_jurisdiction_words(citation["qual_words"], jurisdiction)
    file_words, file_abbreviations = file_qualifiers(section["path"])
    file_words = drop_jurisdiction_words(file_words, jurisdiction)
    contradicted = False

    if citation["qual_numbers"]:
        if citation["qual_numbers"] <= section_numbers:
            score += 20
            reasons.append("title/chapter number")
        else:
            score -= 40
            contradicted = True
            reasons.append("title/chapter mismatch")

    matched = {w for w in cite_words if any(words_match(w, x) for x in section_words | file_words)}
    if matched:
        score += min(10, 5 * len(matched))
        reasons.append(f"code words: {', '.join(sorted(matched))}")
    extra = {w for w in section_words if not any(words_match(w, c) for c in cite_words)}
    if extra and cite_words:
        score -= min(15, 5 * len(extra))
        reasons.append(f"other code words: {', '.join(sorted(extra))}")
    if cite_words and file_words and all(any(words_match(f, c) for c in cite_words) for f in file_words) \
            and all(any(words_match(c, f) for f in file_words) for c in cite_words):
        score += 10
        reasons.append("file name is the cited code")

    abbreviations = code_abbreviations(section["prefix"], jurisdiction) | file_abbreviations
    if any(a in citation["qual_initials"] for a in abbreviations if len(a) >= 2):
        score += 15
        reasons.append("code abbreviation")
    elif any(len(a) >= 3 and is_subsequence(a, citation["qual_letters"]) for a in abbreviations):
        score += 8
        reasons.append("code abbreviation (loose)")

    similarity = heading_similarity(citation["heading"], section["heading"])
    if citation["heading"] and section["heading"] and similarity < 0.35:
        score -= 15
        contradicted = True
        reasons.append(f"heading differs {similarity:.2f}")
    elif similarity:
        score += round(20 * similarity)
        reasons.append(f"heading {similarity:.2f}")

    if unique_in_year and not contradicted:
        score += 15
        reasons.append("only section with this number")
    return Candidate(section, score, reasons)


def title_hint(citation: dict) -> str:
    """Leading number of the section key ('35-9a-141' -> '35', '55.1-1245'
    -> '55'), used to keep the heading fallback inside the same title even
    when the title itself was renumbered (Va. 55 -> 55.1)."""
    match = re.match(r"^(\d+)", citation["section_key"])
    return match.group(1) if match else ""


def heading_fallback(
    connection: sqlite3.Connection,
    citation: dict,
    year: int,
    threshold: float,
    max_candidates: int,
) -> list[Candidate]:
    """Sections in the same title whose heading closely matches the
    citation's heading. Offered for review only (status HEADING_MATCH)."""
    hint = title_hint(citation)
    if not hint or not citation["heading"]:
        return []
    rows = connection.execute(
        "SELECT * FROM sections WHERE jurisdiction = ? AND year = ? AND section_key LIKE ?",
        (citation["jurisdiction"], year, hint + "%"),
    ).fetchall()
    same_title = re.compile(re.escape(hint) + r"(?!\d)")
    scored = []
    for row in rows:
        if not same_title.match(row["section_key"]):
            continue
        similarity = heading_similarity(citation["heading"], row["heading"])
        if similarity >= threshold:
            scored.append(Candidate(dict(row), 40 + round(30 * similarity), [f"heading {similarity:.2f} (number differs)"]))
    scored = best_per_shortname(scored)
    return scored[:max_candidates]


def best_per_shortname(candidates: Iterable[Candidate]) -> list[Candidate]:
    best: dict[tuple[str, str], Candidate] = {}
    for candidate in candidates:
        key = (candidate.section["shortname"] or candidate.section["section_key"], candidate.section["content_hash"])
        if key not in best or candidate.score > best[key].score:
            best[key] = candidate
    return sorted(best.values(), key=lambda c: (-c.score, c.section["shortname"]))


def match_citation_all_years(
    connection: sqlite3.Connection,
    citation: dict,
    years: Sequence[int],
    args: argparse.Namespace,
) -> dict[int, YearMatch]:
    rows_by_year: dict[int, list[dict]] = defaultdict(list)
    for row in connection.execute(
        "SELECT * FROM sections WHERE jurisdiction = ? AND section_key = ?",
        (citation["jurisdiction"], citation["section_key"]),
    ):
        rows_by_year[row["year"]].append(dict(row))

    results = {}
    for year in years:
        rows = rows_by_year.get(year, [])
        distinct = {(row["shortname"], row["content_hash"]) for row in rows}
        candidates = best_per_shortname(score_section(citation, row, len(distinct) == 1) for row in rows)
        candidates = candidates[: args.max_candidates]
        if not candidates:
            fallback = heading_fallback(connection, citation, year, args.heading_threshold, args.max_candidates)
            if fallback:
                results[year] = YearMatch(year, "HEADING_MATCH", fallback[0].score, None, fallback)
            else:
                results[year] = YearMatch(year, "NOT_FOUND", 0, None, [])
            continue
        top = candidates[0]
        if top.score < args.min_score:
            status = "LOW_SCORE"
        elif len(candidates) > 1 and top.score - candidates[1].score < args.ambiguity_margin:
            status = "AMBIGUOUS"
        else:
            status = "MATCHED"
        results[year] = YearMatch(year, status, top.score, top.section if status == "MATCHED" else None, candidates)
    return results


def summarize_years(matches: dict[int, YearMatch]) -> tuple[str, dict[int, str]]:
    """'8/8 years (2009-2019); text changed 2013, 2016' plus, per year,
    whether the matched text differs from the previous matched year."""
    matched = [m for _, m in sorted(matches.items()) if m.accepted]
    changed: dict[int, str] = {}
    previous_hash = None
    change_years = []
    for match in matched:
        current = match.section["content_hash"]
        if previous_hash is None:
            changed[match.year] = ""
        elif current != previous_hash:
            changed[match.year] = "yes"
            change_years.append(str(match.year))
        else:
            changed[match.year] = "no"
        previous_hash = current
    if not matched:
        return f"0/{len(matches)} years", changed
    summary = f"{len(matched)}/{len(matches)} years ({matched[0].year}-{matched[-1].year})"
    if change_years:
        summary += f"; text changed {', '.join(change_years)}"
    return summary, changed


def unique_join(values: Iterable[object]) -> str:
    output, seen = [], set()
    for value in values:
        value = normalize_text(value)
        if value and value not in seen:
            output.append(value)
            seen.add(value)
    return "; ".join(output)


def combine_statuses(statuses: Sequence[str]) -> str:
    if not statuses:
        return "NO_CITATION"
    unique = set(statuses)
    if unique == {"MATCHED"}:
        return "MATCHED"
    if "MATCHED" in unique:
        return "PARTIAL"
    for status in ("AMBIGUOUS", "HEADING_MATCH", "LOW_SCORE", "NO_XML_YEAR", "NOT_FOUND"):
        if status in unique:
            return status
    return "NO_MATCH"


# ---------------------------------------------------------------------------
# Workbook output
# ---------------------------------------------------------------------------

def header_map(worksheet) -> dict[str, int]:
    return {normalize_text(cell.value): cell.column for cell in worksheet[1] if normalize_text(cell.value)}


def ensure_review_columns(worksheet, headers: dict[str, int]) -> dict[str, int]:
    source_header = worksheet.cell(1, max(headers.values()))
    for column_name in REVIEW_COLUMNS:
        if column_name in headers:
            continue
        column = worksheet.max_column + 1
        cell = worksheet.cell(1, column, column_name)
        if source_header.has_style:
            cell._style = copy(source_header._style)
        headers[column_name] = column
    return headers


def write_row(worksheet, row_number: int, headers: dict[str, int], values: dict, overwrite: bool) -> None:
    for name, value in values.items():
        cell = worksheet.cell(row_number, headers[name])
        if name in REQUIRED_COLUMNS and not overwrite and normalize_text(cell.value):
            continue
        cell.value = value


def process_workbook(args: argparse.Namespace) -> tuple[Path, Path, dict[str, int]]:
    workbook_path = Path(args.workbook).expanduser().resolve()
    xml_root = Path(args.xml_root).expanduser().resolve()
    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else workbook_path.with_name(f"{workbook_path.stem}_matched.xlsx")
    )
    report_path = (
        Path(args.report).expanduser().resolve()
        if args.report
        else output_path.with_suffix(".by_year.csv")
    )
    index_path = Path(args.index_db).expanduser().resolve()

    citations = detect_citation(workbook_path, args.sheet)
    selected = set(args.jurisdiction or [])
    if selected:
        citations = [c for c in citations if c["jurisdiction"] in selected]
    jurisdictions = sorted({c["jurisdiction"] for c in citations})

    connection = open_index(index_path)
    for jurisdiction in jurisdictions:
        index_jurisdiction(connection, xml_root, jurisdiction, rebuild=args.rebuild_index)

    workbook = load_workbook(workbook_path)
    worksheet = workbook[args.sheet] if args.sheet else workbook.active
    headers = header_map(worksheet)
    missing = [name for name in REQUIRED_COLUMNS if name not in headers]
    if missing:
        raise ValueError(f"Workbook is missing required column(s): {', '.join(missing)}")
    headers = ensure_review_columns(worksheet, headers)

    by_row: dict[int, list[dict]] = defaultdict(list)
    for citation in citations:
        by_row[citation["row"]].append(citation)

    report_rows = []
    summary: dict[str, int] = defaultdict(int)
    for row_number in range(2, worksheet.max_row + 1):
        jurisdiction = normalize_text(worksheet.cell(row_number, headers["Jurisdiction"]).value)
        if not jurisdiction or (selected and jurisdiction not in selected):
            continue
        row_citations = by_row.get(row_number, [])
        years = [year for year, _ in list_year_dirs(xml_root / jurisdiction)]
        resolved = detect_year(args.target_year, xml_root / jurisdiction)
        target = resolved[0] if resolved else None

        statuses, titles, chapters, sections, files, shortnames = [], [], [], [], [], []
        scores, year_notes, candidates_text, notes = [], [], [], []
        for citation in row_citations:
            label = citation["cite_normalized"] or citation["citation_raw"][:60]
            if not citation["section_key"]:
                statuses.append("UNPARSED_CITATION")
                notes.append(f"{label}: {citation['citation_type']} citation, not matched to code sections")
                continue
            matches = match_citation_all_years(connection, citation, years, args)
            year_summary, changed = summarize_years(matches)
            year_notes.append(f"{label}: {year_summary}")

            for year, match in sorted(matches.items()):
                section = match.section or {}
                report_rows.append({
                    "Workbook Row": row_number,
                    "Jurisdiction": jurisdiction,
                    "Question": citation["question"],
                    "Position": citation["position"],
                    "Citation": citation["citation_raw"],
                    "Citation Type": citation["citation_type"],
                    "Section Key": citation["section_key"],
                    "Year": year,
                    "Status": match.status,
                    "Score": match.score,
                    "Matched ShortName": section.get("shortname", ""),
                    "XML Heading": section.get("heading", ""),
                    "XML File": section.get("path", ""),
                    "XML Title": section.get("title", ""),
                    "Chapter": section.get("chapter", ""),
                    "Hierarchy": section.get("hierarchy", ""),
                    "Content Hash": section.get("content_hash", ""),
                    "Changed From Prior Year": changed.get(year, ""),
                    "Candidates": " | ".join(c.label() for c in match.candidates) if not match.accepted else "",
                    "Notes": "; ".join(match.candidates[0].reasons) if match.candidates else "",
                })

            if target is None:
                statuses.append("NO_XML_YEAR")
                notes.append(f"{label}: no XML year at or before {args.target_year}")
                continue
            match = matches[target]
            statuses.append(match.status)
            scores.append(match.score)
            if match.accepted:
                section = match.section
                titles.append(section["title"])
                chapters.append(section["chapter"])
                sections.append(section["shortname"])
                files.append(section["path"])
                shortnames.append(section["shortname"])
            else:
                candidates_text.extend(f"{label} -> {c.label()}" for c in match.candidates)
            notes.append(
                f"{label}: {match.status}"
                + (f" ({', '.join(match.candidates[0].reasons)})" if match.candidates else "")
            )

        status = combine_statuses(statuses)
        summary[status] += 1
        write_row(worksheet, row_number, headers, {
            "XML Title": unique_join(titles),
            "Chapter": unique_join(chapters),
            "Section": unique_join(sections),
            "XML File": unique_join(files),
            "XML Year": target or "",
            "Matched ShortName": unique_join(shortnames),
            "Match Status": status,
            "Match Score": min(scores) if scores else 0,
            "Years Matched": " | ".join(year_notes),
            "Candidates": " | ".join(candidates_text),
            "Match Notes": " | ".join(notes),
        }, args.overwrite_existing)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(LONG_REPORT_COLUMNS))
        writer.writeheader()
        writer.writerows(report_rows)
    connection.close()
    return output_path, report_path, dict(summary)


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

def _write_test_title(path: Path, title: str, chapter: str, sections: Sequence[tuple[str, str, str, str]]) -> None:
    body = "".join(
        f'<Index Level="3" HasChildren="0"><Caption>Section {number}</Caption>'
        f"<Description>{heading}</Description>"
        f"<Content>&lt;P&gt;{text}&lt;/P&gt;</Content>"
        f"<ShortName>{shortname}</ShortName><RevisionHistory>(Acts 2006)</RevisionHistory></Index>"
        for number, shortname, heading, text in sections
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        '\ufeff<?xml version="1.0" encoding="UTF-8"?><Content Type="Statutes"><Indexes>'
        f'<Index Level="1" HasChildren="1"><Caption>{title}</Caption><Description>PROPERTY</Description><Indexes>'
        f'<Index Level="2" HasChildren="1"><Caption>{chapter}</Caption><Description>LANDLORD</Description>'
        f"<Indexes>{body}</Indexes></Index></Indexes></Index></Indexes></Content>",
        encoding="utf-8",
    )


def run_self_test() -> None:
    with tempfile.TemporaryDirectory() as temporary:
        root = Path(temporary)
        xml_root = root / "xml"
        heading = "Noncompliance with rental agreement; failure to pay rent"
        _write_test_title(xml_root / "Alabama" / "2009" / "Title 35.xml", "Title 35", "Chapter 9A", [
            ("35-9A-400", "Ala. Code 35-9A-400", heading, "old text"),  # renumbered later
            ("35-9A-141", "Ala. Code 35-9A-141", "Definitions", "defs"),
        ])
        _write_test_title(xml_root / "Alabama" / "2019" / "Title 35.xml", "Title 35", "Chapter 9A", [
            ("35-9A-421", "Ala. Code 35-9A-421", heading, "new text"),
            ("35-9A-141", "Ala. Code 35-9A-141", "Definitions", "defs v2"),
        ])
        _write_test_title(xml_root / "Alabama" / "2019" / "Title 1.xml", "Title 1", "Chapter 1", [
            ("1-1-1", "Ala. Code 1-1-1", "Meaning of certain words", "x"),
        ])
        _write_test_title(xml_root / "New York" / "2019" / "Consolidated Laws.xml", "Laws", "Article 7", [
            ("711", "NY RPA Law 711", "Grounds where landlord-tenant relationship exists", "a"),
            ("711", "NY EDN Law 711", "Textbooks; boards of education", "b"),
        ])
        _write_test_title(xml_root / "Delaware" / "2019" / "Title 25.xml", "Title 25", "Chapter 51", [
            ("5106", "Del. Code Tit. 25 Sec. 5106", "Rental agreement; term and termination", "c"),
            ("5106", "Del. Code Tit. 7 Sec. 5106", "Something else", "d"),
        ])

        _write_test_title(xml_root / "California" / "2019" / "Civil Code - CIV.xml", "Civil Code", "Chapter 2", [
            ("1946", "CA Civ. Sec. 1946", "", "hiring renewed"),
        ])
        _write_test_title(xml_root / "California" / "2019" / "Code of Civil Procedure - CCP.xml", "CCP", "Chapter 4", [
            ("1946", "CA Civ. Proc. Sec. 1946", "", "other"),
            ("1161", "CA Civ. Proc. Sec. 1161", "", "unlawful detainer"),
        ])
        _write_test_title(xml_root / "California" / "2019" / "Education Code - EDC.xml", "Education", "Chapter 9", [
            ("1946", "CA Educ. Sec. 1946", "", "schools"),
            ("1161", "CA Educ. Sec. 1161", "", "schools"),
        ])

        cites = parse_citation_cell(
            "Ala. Code § 35-9A-421. Noncompliance with rental agreement; failure to pay rent.; "
            "Ala. Code § 35-9A-421. Noncompliance with rental agreement; failure to pay rent.; "
            "Executive Order 2020-72 (COVID-19)"
        )
        assert len(cites) == 2, cites
        assert cites[0]["section_key"] == "35-9a-421", cites[0]
        assert cites[0]["heading"].endswith("failure to pay rent"), cites[0]
        assert cites[1]["citation_type"] == "other", cites[1]
        ilcs = parse_citation_cell("735 Ill. Comp. Stat. 5/9-209 Demand for rent")[0]
        assert ilcs["section_key"] == "5/9-209" and ilcs["qual_numbers"] == {"735"}, ilcs
        assert split_identifier("735 ILCS 5/9-209") == ("735 ILCS", "5/9-209")
        assert split_identifier("Alaska Stat. 09.45.100")[1] == "9.45.100"
        assert split_identifier("CA Bus. &amp;amp; Prof. Sec. 2") == ("CA Bus. & Prof.", "2")

        connection = open_index(root / "index.sqlite")
        for jurisdiction in ("Alabama", "New York", "Delaware", "California"):
            index_jurisdiction(connection, xml_root, jurisdiction)
        args = argparse.Namespace(min_score=70, ambiguity_margin=10, max_candidates=5, heading_threshold=0.85)

        def cite(jurisdiction: str, text: str) -> dict:
            parsed = parse_citation_cell(text)[0]
            parsed.update(jurisdiction=jurisdiction, question="test", row=2)
            return parsed

        al = match_citation_all_years(connection, cite("Alabama", "Ala. Code § 35-9A-421. " + heading), [2009, 2019], args)
        assert al[2019].status == "MATCHED" and al[2019].section["title"] == "Title 35", al[2019]
        assert al[2019].section["chapter"] == "Chapter 9A", al[2019]
        assert al[2009].status == "HEADING_MATCH", al[2009]
        assert al[2009].candidates[0].section["shortname"] == "Ala. Code 35-9A-400", al[2009]

        defs = match_citation_all_years(connection, cite("Alabama", "Ala. Code § 35-9A-141. Definitions."), [2009, 2019], args)
        assert all(m.status == "MATCHED" for m in defs.values()), defs
        summary, changed = summarize_years(defs)
        assert changed[2019] == "yes" and "text changed 2019" in summary, summary

        ny = match_citation_all_years(connection, cite(
            "New York", "N.Y. Real Prop. Acts. Law § 711. Grounds where landlord-tenant relationship exists"), [2019], args)
        assert ny[2019].status == "MATCHED" and ny[2019].section["shortname"] == "NY RPA Law 711", ny[2019]

        de = match_citation_all_years(connection, cite(
            "Delaware", "Del Code tit. 25 § 5106. Rental agreement; term and termination of rental agreement"), [2019], args)
        assert de[2019].status == "MATCHED" and "Tit. 25" in de[2019].section["shortname"], de[2019]

        ca = match_citation_all_years(connection, cite("California", "Cal. Civil code § 1946. Renewable hiring; notice"), [2019], args)
        assert ca[2019].status == "MATCHED" and ca[2019].section["shortname"] == "CA Civ. Sec. 1946", ca[2019]
        ccp = match_citation_all_years(connection, cite(
            "California", "Cal. Civil Procedure Code § 1161. Unlawful detainer defined"), [2019], args)
        assert ccp[2019].status == "MATCHED" and ccp[2019].section["shortname"] == "CA Civ. Proc. Sec. 1161", ccp[2019]

        assert detect_year(2021, xml_root / "Alabama")[0] == 2019
        assert detect_year(2008, xml_root / "Alabama") is None

        # Re-indexing unchanged files is a no-op.
        before = connection.execute("SELECT COUNT(*) FROM sections").fetchone()[0]
        index_jurisdiction(connection, xml_root, "Alabama")
        assert connection.execute("SELECT COUNT(*) FROM sections").fetchone()[0] == before
        connection.close()

    print("Self-test passed.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Match LSC citations to sections of historical state-statute XML files.",
    )
    parser.add_argument("--workbook", help="Input citation workbook (.xlsx).")
    parser.add_argument("--xml-root", help="Root directory laid out as <Jurisdiction>/<Year>/*.xml.")
    parser.add_argument("--output", help="Output workbook path.")
    parser.add_argument("--report", help="Long-format CSV (one row per citation x year).")
    parser.add_argument("--sheet", help="Worksheet name. Defaults to the active sheet.")
    parser.add_argument(
        "--index-db",
        default="data/section_index.sqlite",
        help="SQLite section-index cache. Default: data/section_index.sqlite.",
    )
    parser.add_argument("--rebuild-index", action="store_true", help="Re-parse every XML file.")
    parser.add_argument(
        "--jurisdiction",
        action="append",
        help="Process one jurisdiction only. Repeat the flag for multiple jurisdictions.",
    )
    parser.add_argument(
        "--target-year",
        type=int,
        default=2021,
        help="Year whose section fills the workbook columns (exact or latest prior). Default: 2021.",
    )
    parser.add_argument("--min-score", type=int, default=70, help="Minimum score for a match. Default: 70.")
    parser.add_argument(
        "--ambiguity-margin", type=int, default=10,
        help="Minimum lead over the second candidate. Default: 10.",
    )
    parser.add_argument(
        "--max-candidates", type=int, default=5,
        help="Maximum candidates retained for review. Default: 5.",
    )
    parser.add_argument(
        "--heading-threshold", type=float, default=0.85,
        help="Heading similarity needed to propose a renumbered section. Default: 0.85.",
    )
    parser.add_argument(
        "--overwrite-existing", action="store_true",
        help="Replace existing XML Title, Chapter, and Section values.",
    )
    parser.add_argument("--self-test", action="store_true", help="Run the built-in tests and exit.")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.self_test:
        run_self_test()
        return 0
    if not args.workbook or not args.xml_root:
        parser.error("--workbook and --xml-root are required unless --self-test is used")

    output, report, summary = process_workbook(args)
    print(f"Workbook: {output}")
    print(f"By-year report: {report}")
    print("Match summary (workbook rows, target year):")
    for status, count in sorted(summary.items()):
        print(f"  {status}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
