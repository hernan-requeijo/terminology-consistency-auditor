"""
TERMINOLOGY CONSISTENCY AUDITOR
================================
Project: TERMINOLOGY CONSISTENCY
Module : auditor.py  —  Core engine (v0.2)

Supports:
  - TMX files  (bilingual Eng/Spa segments in one file)
  - Parallel document pairs  (one Eng file + one Spa file)
    * .docx   (Word)
    * .pdf    (PDF via pdfminer)
    * .txt    (plain text)
    * .xlsx   (Excel — column-pair mode)
    * .csv    (CSV — column-pair mode)

Output:
  - Console summary table
  - CSV report  (inconsistency_report_<timestamp>.csv)
  - JSON report (inconsistency_report_<timestamp>.json)

Usage examples:
  python auditor.py --tmx file1.tmx file2.tmx
  python auditor.py --src manual_en.docx --tgt manual_es.docx
  python auditor.py --tmx corp1.tmx --src addendum_en.pdf --tgt addendum_es.pdf
  python auditor.py --tmx corp1.tmx --min-freq 2 --min-variants 2 --output report
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field, asdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import List, Dict, Tuple, Optional


# ─────────────────────────────────────────────
#  Data structures
# ─────────────────────────────────────────────

@dataclass
class TranslationUnit:
    """One bilingual segment pair."""
    source: str
    target: str
    origin: str          # filename or label
    unit_id: str = ""    # segment id if available
    segment_only: bool = False  # if True: skip n-gram extraction, use whole segment as term


@dataclass
class TermCluster:
    """All translations found for one normalised source term."""
    source_norm: str                           # normalised source string
    source_display: str                        # most common raw form
    variants: Dict[str, List[str]] = field(default_factory=dict)
    # variants = { target_norm -> [origin, origin, …] }

    @property
    def variant_count(self) -> int:
        return len(self.variants)

    @property
    def total_occurrences(self) -> int:
        return sum(len(v) for v in self.variants.values())

    @property
    def is_inconsistent(self) -> bool:
        return self.variant_count > 1

    @property
    def severity(self) -> str:
        """Simple severity based on frequency and dispersion."""
        if self.variant_count >= 4:
            return "CRITICAL"
        if self.variant_count == 3:
            return "HIGH"
        if self.variant_count == 2 and self.total_occurrences >= 5:
            return "MEDIUM"
        return "LOW"

    @property
    def dominant_translation(self) -> str:
        """The most frequently used translation."""
        return max(self.variants, key=lambda t: len(self.variants[t]))

    def to_dict(self) -> dict:
        return {
            "source": self.source_display,
            "source_normalised": self.source_norm,
            "variant_count": self.variant_count,
            "total_occurrences": self.total_occurrences,
            "severity": self.severity,
            "dominant_translation": self.dominant_translation,
            "variants": {
                tgt: {"occurrences": len(orgs), "origins": list(set(orgs))}
                for tgt, orgs in self.variants.items()
            },
        }


# ─────────────────────────────────────────────
#  Text normalisation helpers
# ─────────────────────────────────────────────

def normalise(text: str) -> str:
    """
    Light normalisation for comparison keys:
    lowercase, collapse whitespace, strip leading/trailing punctuation.
    Preserves enough form for a linguist to recognise the term.
    """
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"^[\s\W]+|[\s\W]+$", "", text)
    return text


def similarity(a: str, b: str) -> float:
    """SequenceMatcher ratio — 0.0 to 1.0."""
    return SequenceMatcher(None, a, b).ratio()


def is_trivial(text: str) -> bool:
    """Skip segments that carry no terminology value."""
    t = text.strip()
    if not t:
        return True
    # Pure numbers / dates / URLs / single characters
    if re.fullmatch(r"[\d\s.,/:;-]+", t):
        return True
    if re.match(r"https?://", t, re.I):
        return True
    if len(t) <= 2:
        return True
    return False


# ─────────────────────────────────────────────
#  Spanish-specific helpers
# ─────────────────────────────────────────────

# Articles, prepositions, and clitic pronouns that commonly attach to
# the edges of extracted windows but are not part of the term itself.
_ES_FUNCTION_WORDS: frozenset = frozenset({
    "el", "la", "los", "las", "un", "una", "unos", "unas",
    "de", "del", "al", "en", "a", "por", "para", "con", "sin",
    "se", "le", "les", "lo", "su", "sus",
    "este", "esta", "estos", "estas", "ese", "esa", "esos", "esas",
    "que", "y", "o", "pero", "más", "como", "también", "ya", "si",
    "es", "son", "está", "están", "ser", "estar",
    "no", "ni", "e", "u",
})


def strip_es_function_words(phrase: str) -> str:
    """
    Remove leading and trailing Spanish function words / articles / prepositions
    from an extracted target window before storing as a variant key.

    Preserves a minimum of 2 words so that short windows (2 words) are not
    reduced to single tokens, which would discard the term content.

    Example: 'los datos personales' → 'datos personales'
             'el plazo de'          → 'plazo'
             'los datos'            → 'los datos'  (would strip to 1 word — preserved)
    """
    words = phrase.split()
    MIN_WORDS = 2  # never reduce below this
    # Strip from left
    while len(words) > MIN_WORDS and words[0].lower().rstrip(".,;:") in _ES_FUNCTION_WORDS:
        words = words[1:]
    # Strip from right
    while len(words) > MIN_WORDS and words[-1].lower().rstrip(".,;:") in _ES_FUNCTION_WORDS:
        words = words[:-1]
    return " ".join(words)


# ─────────────────────────────────────────────
#  File parsers
# ─────────────────────────────────────────────

class TMXParser:
    """Parse a TMX file and return a list of TranslationUnit objects."""

    SRC_LANGS = {"en", "eng", "en-us", "en-gb", "english"}
    TGT_LANGS = {"es", "spa", "es-es", "es-mx", "es-419", "spanish", "español"}

    def parse(self, path: Path) -> List[TranslationUnit]:
        units: List[TranslationUnit] = []
        try:
            tree = ET.parse(path)
        except ET.ParseError as e:
            print(f"  [WARN] Could not parse TMX {path.name}: {e}")
            return units

        root = tree.getroot()
        # TMX namespace handling
        ns = ""
        if root.tag.startswith("{"):
            ns = root.tag.split("}")[0] + "}"

        for tu in root.iter(f"{ns}tu"):
            seg_map: Dict[str, str] = {}
            for tuv in tu.findall(f"{ns}tuv"):
                lang = (
                    tuv.get("lang") or
                    tuv.get("{http://www.w3.org/XML/1998/namespace}lang") or
                    ""
                ).lower().strip()
                seg_el = tuv.find(f"{ns}seg")
                text = ("".join(seg_el.itertext()) if seg_el is not None else "").strip()
                if text:
                    seg_map[lang] = text

            src_text = self._pick(seg_map, self.SRC_LANGS)
            tgt_text = self._pick(seg_map, self.TGT_LANGS)

            if src_text and tgt_text:
                units.append(TranslationUnit(
                    source=src_text,
                    target=tgt_text,
                    origin=path.name,
                    unit_id=tu.get("id", ""),
                ))
        return units

    def _pick(self, seg_map: Dict[str, str], lang_set: set) -> Optional[str]:
        for lang, text in seg_map.items():
            if any(lang.startswith(l) for l in lang_set):
                return text
        return None


class DocxParser:
    """Extract paragraphs from a .docx file."""

    def parse(self, path: Path) -> List[str]:
        try:
            from docx import Document
            doc = Document(str(path))
            paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            return paras
        except Exception as e:
            print(f"  [WARN] Could not read DOCX {path.name}: {e}")
            return []


class PDFParser:
    """Extract text blocks from a PDF using pdfminer."""

    def parse(self, path: Path) -> List[str]:
        try:
            from pdfminer.high_level import extract_text
            raw = extract_text(str(path))
            # Split on double newlines to approximate paragraphs
            blocks = [b.strip() for b in re.split(r"\n{2,}", raw) if b.strip()]
            return blocks
        except Exception as e:
            print(f"  [WARN] Could not read PDF {path.name}: {e}")
            return []


class TextParser:
    """Extract lines/paragraphs from plain text."""

    def parse(self, path: Path) -> List[str]:
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
            blocks = [b.strip() for b in re.split(r"\n{2,}", text) if b.strip()]
            return blocks if blocks else [l.strip() for l in text.splitlines() if l.strip()]
        except Exception as e:
            print(f"  [WARN] Could not read TXT {path.name}: {e}")
            return []


class SpreadsheetParser:
    """
    Extract bilingual pairs from Excel/CSV.
    Assumes columns: col 0 = source (Eng), col 1 = target (Spa).
    Or auto-detects if headers contain 'en'/'es' or 'source'/'target'.
    """

    def parse_excel(self, path: Path) -> List[Tuple[str, str]]:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            return self._extract_pairs(rows, path.name)
        except Exception as e:
            print(f"  [WARN] Could not read XLSX {path.name}: {e}")
            return []

    def parse_csv(self, path: Path) -> List[Tuple[str, str]]:
        try:
            rows = []
            with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
                reader = csv.reader(f)
                rows = list(reader)
            return self._extract_pairs(rows, path.name)
        except Exception as e:
            print(f"  [WARN] Could not read CSV {path.name}: {e}")
            return []

    def _extract_pairs(self, rows: list, name: str) -> List[Tuple[str, str]]:
        if not rows:
            return []
        src_col, tgt_col = 0, 1
        # Try header detection
        header = [str(c).lower() if c else "" for c in rows[0]]
        for i, h in enumerate(header):
            if any(x in h for x in ["source", "english", " en", "src"]):
                src_col = i
            if any(x in h for x in ["target", "spanish", " es", "tgt", "translation"]):
                tgt_col = i
        start = 1 if any(header) else 0
        pairs = []
        for row in rows[start:]:
            if len(row) > max(src_col, tgt_col):
                s = str(row[src_col]).strip() if row[src_col] else ""
                t = str(row[tgt_col]).strip() if row[tgt_col] else ""
                if s and t:
                    pairs.append((s, t))
        return pairs


# ─────────────────────────────────────────────
#  Parallel document aligner
# ─────────────────────────────────────────────

class ParallelAligner:
    """
    Align paragraphs from a source document and a target document
    into TranslationUnit pairs.

    Strategy:
      1. If counts match exactly → zip 1:1.
      2. If off by ≤20% → greedy similarity alignment.
      3. Otherwise → sentence-count-weighted block merge.
    """

    FUZZY_THRESHOLD = 0.55  # minimum similarity to accept an alignment

    def align(
        self,
        src_blocks: List[str],
        tgt_blocks: List[str],
        src_name: str,
        tgt_name: str,
    ) -> List[TranslationUnit]:

        # Filter trivial blocks
        src = [b for b in src_blocks if not is_trivial(b)]
        tgt = [b for b in tgt_blocks if not is_trivial(b)]

        if not src or not tgt:
            print(f"  [WARN] One or both documents yielded no usable segments.")
            return []

        ratio = len(src) / len(tgt) if tgt else 0
        label = f"{src_name} ↔ {tgt_name}"

        if len(src) == len(tgt):
            print(f"  [INFO] Perfect 1:1 alignment ({len(src)} segments) — {label}")
            return self._zip(src, tgt, label)

        if 0.8 <= ratio <= 1.25:
            print(f"  [INFO] Near-match alignment ({len(src)} src / {len(tgt)} tgt) — {label}")
            return self._greedy_align(src, tgt, label)

        print(
            f"  [INFO] Loose alignment ({len(src)} src / {len(tgt)} tgt, ratio={ratio:.2f}) "
            f"— using sentence-weighted merge — {label}"
        )
        return self._sentence_weighted_merge(src, tgt, label)

    def _zip(self, src, tgt, label) -> List[TranslationUnit]:
        return [
            TranslationUnit(source=s, target=t, origin=label,
                            unit_id=str(i), segment_only=True)
            for i, (s, t) in enumerate(zip(src, tgt), 1)
        ]

    def _greedy_align(self, src, tgt, label) -> List[TranslationUnit]:
        """
        Simple greedy alignment: for each source block, find the
        best-matching unused target block (by character-length ratio
        as a proxy for sentence length similarity).
        """
        units = []
        used = set()
        for i, s in enumerate(src):
            best_j, best_score = -1, -1.0
            s_len = len(s)
            for j, t in enumerate(tgt):
                if j in used:
                    continue
                len_ratio = min(s_len, len(t)) / max(s_len, len(t), 1)
                if len_ratio > best_score:
                    best_score = len_ratio
                    best_j = j
            if best_j >= 0 and best_score >= 0.3:
                used.add(best_j)
                units.append(TranslationUnit(
                    source=s,
                    target=tgt[best_j],
                    origin=label,
                    unit_id=str(i),
                ))
        return units

    def _sentence_weighted_merge(self, src, tgt, label) -> List[TranslationUnit]:
        """
        When block counts differ significantly, merge blocks until
        sentence counts are roughly balanced, then zip.
        """
        def sentence_count(text):
            return max(1, len(re.split(r"[.!?]+", text)))

        def merge_until(blocks, target_count):
            result = []
            buf = []
            buf_count = 0
            per_chunk = max(1, sum(sentence_count(b) for b in blocks) // target_count)
            for b in blocks:
                buf.append(b)
                buf_count += sentence_count(b)
                if buf_count >= per_chunk:
                    result.append(" ".join(buf))
                    buf, buf_count = [], 0
            if buf:
                result.append(" ".join(buf))
            return result

        target_n = min(len(src), len(tgt))
        merged_src = merge_until(src, target_n)
        merged_tgt = merge_until(tgt, target_n)
        return self._zip(merged_src, merged_tgt, label)


# ─────────────────────────────────────────────
#  Term extractor  (noun-phrase heuristic)
# ─────────────────────────────────────────────

class TermExtractor:
    """
    Lightweight heuristic term extraction from segments.

    Extracts candidate terms as:
      - Full segment (if short, ≤ 10 words)
      - Noun phrases via capitalised-word sequences
      - Multi-word sequences repeated across the corpus

    For v0.1 we operate at full-segment level for TMX
    (each TU is already a terminological unit) and apply
    n-gram extraction for parallel documents.
    """

    MAX_SEGMENT_WORDS = 12   # treat whole segment as term candidate
    MIN_TERM_CHARS = 3

    def extract_from_unit(self, unit: TranslationUnit) -> List[Tuple[str, str]]:
        """
        Return (src_term, tgt_term) pairs from one TranslationUnit.

        segment_only=True  → whole segment only (parallel docs, TMX with long segments).
        segment_only=False → whole segment + noun-phrase n-grams (short TMX TUs).
        """
        # Always include whole segment
        pairs = [(unit.source, unit.target)]

        if unit.segment_only:
            return pairs

        src_words = unit.source.split()
        if len(src_words) <= self.MAX_SEGMENT_WORDS:
            return pairs  # short enough — whole segment is sufficient

        # Long segments with n-gram expansion (TMX mode only)
        src_candidates = self._noun_phrases(unit.source)
        tgt_candidates = self._noun_phrases(unit.target)
        for sc, tc in zip(src_candidates, tgt_candidates):
            if len(sc) >= self.MIN_TERM_CHARS and len(tc) >= self.MIN_TERM_CHARS:
                pairs.append((sc, tc))

        return pairs

    def _noun_phrases(self, text: str) -> List[str]:
        """
        Heuristic: sequences of capitalised or title-case words,
        plus all 1–4 word subsequences from the text.
        No NLP library required.
        """
        # Capitalised multi-word runs
        caps_pattern = re.compile(r"(?:[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+)+)")
        caps = caps_pattern.findall(text)

        # All 2–4 word n-grams
        words = re.findall(r"\b\w[\w\-]*\b", text)
        ngrams = []
        for n in (2, 3, 4):
            for i in range(len(words) - n + 1):
                phrase = " ".join(words[i:i+n])
                if len(phrase) >= self.MIN_TERM_CHARS:
                    ngrams.append(phrase)

        return caps + ngrams


# ─────────────────────────────────────────────
#  Corpus analyser
# ─────────────────────────────────────────────

class CorpusAnalyser:
    """
    Builds the source→target mapping from a list of TranslationUnits
    and identifies inconsistencies.

    Two-pass strategy:
      Pass 1 — Whole-segment indexing: every TU as a full-segment pair.
      Pass 2 — Sub-segment term extraction: find source phrases that
               appear in MULTIPLE TUs (i.e. they are recurring terms),
               then track how they are translated each time they appear.
               This is the core of real-world terminology consistency checking.
    """

    def __init__(self, min_freq: int = 2, min_variants: int = 2):
        self.min_freq = min_freq
        self.min_variants = min_variants

        # Minimum source phrase length (in words) for sub-segment extraction.
        # Single-word tokens are too positionally unstable for ratio-based
        # target projection — they generate spurious multi-variant clusters.
        self.MIN_SRC_WORDS = 2

        # source_norm -> { target_norm -> [origin, …] }
        self._map: Dict[str, Dict[str, List[str]]] = defaultdict(lambda: defaultdict(list))
        # source_norm -> display form
        self._display: Dict[str, str] = {}

    def ingest(self, units: List[TranslationUnit]) -> None:
        """Two-pass ingestion."""
        # Pass 1: collect all 1–5 word source phrases across corpus
        phrase_units: Dict[str, List[TranslationUnit]] = defaultdict(list)

        for unit in units:
            if is_trivial(unit.source) or is_trivial(unit.target):
                continue
            for phrase in self._source_phrases(unit.source):
                norm = normalise(phrase)
                if norm:
                    phrase_units[norm].append(unit)

        # Pass 2: for phrases appearing in ≥ min_freq units, extract their translations
        for src_norm, unit_list in phrase_units.items():
            if len(unit_list) < self.min_freq:
                continue

            # Skip single-word source phrases: too positionally unstable for
            # ratio-based target projection.  They generate spurious variants
            # when the same word appears in different sentence positions.
            if len(src_norm.split()) < self.MIN_SRC_WORDS:
                continue

            # For each unit containing this source phrase, extract the
            # corresponding target phrase (using sub-segment alignment)
            for unit in unit_list:
                tgt_phrase = self._extract_target_phrase(
                    src_norm, unit.source, unit.target
                )
                if tgt_phrase and not is_trivial(tgt_phrase):
                    # Strip leading/trailing Spanish function words before hashing
                    tgt_clean = strip_es_function_words(tgt_phrase)
                    tgt_norm = normalise(tgt_clean)
                    if tgt_norm and len(tgt_norm) >= 3:
                        self._map[src_norm][tgt_norm].append(unit.origin)
                        raw = self._find_raw(src_norm, unit.source)
                        if src_norm not in self._display or (
                            raw and len(raw) < len(self._display.get(src_norm, raw))
                        ):
                            self._display[src_norm] = raw or src_norm

    def _source_phrases(self, text: str) -> List[str]:
        """
        Extract 1–5 word candidate phrases from a source segment.
        Skips stopword-only phrases.
        """
        STOPWORDS = {
            "the","a","an","of","in","to","and","or","for","with",
            "is","are","was","were","be","been","by","from","at","on",
            "as","it","its","that","this","which","not","no","any",
            "all","each","both","such","their","they","them","have",
            "has","had","will","would","shall","should","may","might",
            "must","can","could","do","does","did","but","so","yet",
        }
        words = re.findall(r"\b\w[\w\-']*\b", text.lower())
        phrases = []
        for n in range(1, 6):  # 1 to 5 word phrases
            for i in range(len(words) - n + 1):
                chunk = words[i:i+n]
                # Skip if all stopwords
                if all(w in STOPWORDS for w in chunk):
                    continue
                # Skip if starts or ends with stopword (for n>1)
                if n > 1 and (chunk[0] in STOPWORDS or chunk[-1] in STOPWORDS):
                    continue
                phrase = " ".join(chunk)
                if len(phrase) >= 3:
                    phrases.append(phrase)
        return phrases

    def _find_raw(self, norm: str, source: str) -> str:
        """Find the original-casing version of norm in source text."""
        words = norm.split()
        src_lower = source.lower()
        idx = src_lower.find(norm)
        if idx >= 0:
            return source[idx: idx + len(norm)]
        return norm

    def _extract_target_phrase(
        self, src_norm: str, src_text: str, tgt_text: str
    ) -> Optional[str]:
        """
        Given a known source phrase (src_norm) within src_text,
        identify the corresponding target phrase in tgt_text.

        Strategy (v0.2):
          - Find character position of src phrase in src_text.
          - Apply an EN→ES length-ratio correction (~1.12) to the projected
            position, compensating for Spanish translations being systematically
            longer than their English sources.
          - Extract a TIGHT window of exactly src_word_count words centred
            on the projected position (v0.1 used a ± half-width that caused
            context bleed, inflating the variant count).
          - Caller applies strip_es_function_words() before normalising.

        Limitations: heuristic, no word alignment model.  Works well for
        2–5 word technical terms; less reliable for longer phrases.
        """
        src_lower = src_text.lower()
        idx = src_lower.find(src_norm)
        if idx < 0:
            return None

        # Character midpoint of the source phrase
        mid_char = idx + len(src_norm) / 2

        # EN→ES positional correction: Spanish text runs ~12% longer on average,
        # so the same concept appears slightly later in the target by character ratio.
        ES_LENGTH_CORRECTION = 1.12
        adjusted_mid = mid_char * ES_LENGTH_CORRECTION
        ratio = adjusted_mid / max(len(src_text), 1)

        tgt_words = tgt_text.split()
        if not tgt_words:
            return None

        center_word = max(0, min(int(ratio * len(tgt_words)), len(tgt_words) - 1))
        src_word_count = len(src_norm.split())

        # Tight window: same word count as source phrase, centred on projection.
        # v0.1 used centre ± (half + src_word_count) which captured too much context.
        half = src_word_count // 2
        start = max(0, center_word - half)
        end = min(len(tgt_words), start + src_word_count)

        # If window is too small (edge of text), extend forward
        if (end - start) < src_word_count:
            end = min(len(tgt_words), start + src_word_count)

        window = " ".join(tgt_words[start:end])
        return window if len(window) >= 3 else None

    def build_clusters(self) -> List[TermCluster]:
        """Return TermCluster objects sorted by severity and frequency."""
        clusters = []
        for src_norm, tgt_map in self._map.items():
            total = sum(len(v) for v in tgt_map.values())
            if total < self.min_freq:
                continue
            if len(tgt_map) < self.min_variants:
                continue
            cluster = TermCluster(
                source_norm=src_norm,
                source_display=self._display.get(src_norm, src_norm),
                variants={t: list(orgs) for t, orgs in tgt_map.items()},
            )
            clusters.append(cluster)

        sev_order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}
        clusters.sort(key=lambda c: (sev_order[c.severity], -c.total_occurrences))
        return clusters

    def stats(self) -> dict:
        total_terms = len(self._map)
        inconsistent = sum(1 for tgt_map in self._map.values() if len(tgt_map) > 1)
        return {
            "unique_source_terms": total_terms,
            "inconsistent_terms": inconsistent,
            "consistency_rate_pct": round(
                100 * (total_terms - inconsistent) / total_terms, 1
            ) if total_terms else 0,
        }


# ─────────────────────────────────────────────
#  File dispatcher
# ─────────────────────────────────────────────

def load_any_file(path: Path) -> List[str]:
    """Dispatch to the right parser based on file extension."""
    ext = path.suffix.lower()
    if ext == ".docx":
        return DocxParser().parse(path)
    if ext == ".pdf":
        return PDFParser().parse(path)
    if ext in (".txt", ".text"):
        return TextParser().parse(path)
    if ext in (".xlsx", ".xls"):
        # Spreadsheets in single-file mode: return flat text blocks
        pairs = SpreadsheetParser().parse_excel(path)
        return [f"{s} ||| {t}" for s, t in pairs]
    if ext == ".csv":
        pairs = SpreadsheetParser().parse_csv(path)
        return [f"{s} ||| {t}" for s, t in pairs]
    print(f"  [WARN] Unsupported extension '{ext}' for {path.name} — skipping.")
    return []


# ─────────────────────────────────────────────
#  Report writers
# ─────────────────────────────────────────────

def write_csv(clusters: List[TermCluster], path: Path) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Severity", "Source (EN)", "# Variants", "Total Occurrences",
            "Dominant Translation (ES)", "All Variants (ES)", "Origins"
        ])
        for c in clusters:
            all_variants = " | ".join(
                f"{t} ({len(orgs)}x)" for t, orgs in c.variants.items()
            )
            origins = " | ".join(sorted(set(
                o for orgs in c.variants.values() for o in orgs
            )))
            writer.writerow([
                c.severity,
                c.source_display,
                c.variant_count,
                c.total_occurrences,
                c.dominant_translation,
                all_variants,
                origins,
            ])


def write_json(clusters: List[TermCluster], stats: dict, path: Path) -> None:
    data = {
        "generated": datetime.now().isoformat(timespec="seconds"),
        "stats": stats,
        "inconsistencies": [c.to_dict() for c in clusters],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def print_summary(clusters: List[TermCluster], stats: dict) -> None:
    print()
    print("=" * 70)
    print("  TERMINOLOGY CONSISTENCY AUDIT — SUMMARY")
    print("=" * 70)
    print(f"  Unique source terms indexed : {stats['unique_source_terms']:,}")
    print(f"  Inconsistent terms found    : {stats['inconsistent_terms']:,}")
    print(f"  Corpus consistency rate     : {stats['consistency_rate_pct']}%")
    print()

    if not clusters:
        print("  ✓ No inconsistencies found above the configured thresholds.")
        print("=" * 70)
        return

    # Show top 20 in a plain-text table
    print(f"  Top inconsistencies (showing up to 20 of {len(clusters)} total):")
    print()
    header = f"  {'SEV':<9} {'#VAR':<6} {'OCC':<5}  {'SOURCE (EN)':<35} {'DOMINANT (ES)'}"
    print(header)
    print("  " + "-" * 68)

    for c in clusters[:20]:
        src = (c.source_display[:33] + "..") if len(c.source_display) > 35 else c.source_display
        dom = (c.dominant_translation[:28] + "..") if len(c.dominant_translation) > 30 else c.dominant_translation
        print(f"  {c.severity:<9} {c.variant_count:<6} {c.total_occurrences:<5}  {src:<35} {dom}")

    print()
    print(f"  Full details in the CSV and JSON reports.")
    print("=" * 70)


# ─────────────────────────────────────────────
#  Main pipeline
# ─────────────────────────────────────────────

def run(
    tmx_files: List[Path],
    src_files: List[Path],
    tgt_files: List[Path],
    spreadsheet_pairs: List[Tuple[Path, Path]],
    min_freq: int,
    min_variants: int,
    output_stem: str,
) -> None:

    analyser = CorpusAnalyser(min_freq=min_freq, min_variants=min_variants)
    aligner = ParallelAligner()
    all_units: List[TranslationUnit] = []

    # ── 1. TMX files ──────────────────────────────────────────────
    if tmx_files:
        print(f"\n[TMX] Parsing {len(tmx_files)} TMX file(s)…")
        parser = TMXParser()
        for tmx_path in tmx_files:
            units = parser.parse(tmx_path)
            print(f"  {tmx_path.name}: {len(units)} translation units")
            all_units.extend(units)

    # ── 2. Parallel document pairs ────────────────────────────────
    if src_files and tgt_files:
        if len(src_files) != len(tgt_files):
            print("[ERROR] Number of --src and --tgt files must match.")
            sys.exit(1)
        print(f"\n[DOCS] Aligning {len(src_files)} document pair(s)…")
        for src_path, tgt_path in zip(src_files, tgt_files):
            print(f"  {src_path.name}  ↔  {tgt_path.name}")
            src_blocks = load_any_file(src_path)
            tgt_blocks = load_any_file(tgt_path)
            units = aligner.align(
                src_blocks, tgt_blocks,
                src_path.name, tgt_path.name,
            )
            print(f"    → {len(units)} aligned segments")
            all_units.extend(units)

    # ── 3. Spreadsheet pairs (bilingual columns) ──────────────────
    if spreadsheet_pairs:
        print(f"\n[XLSX/CSV] Parsing {len(spreadsheet_pairs)} spreadsheet pair(s)…")
        sp = SpreadsheetParser()
        for src_path, tgt_path in spreadsheet_pairs:
            ext = src_path.suffix.lower()
            if ext in (".xlsx", ".xls"):
                pairs = sp.parse_excel(src_path)
            else:
                pairs = sp.parse_csv(src_path)
            units = [
                TranslationUnit(source=s, target=t,
                                origin=f"{src_path.name}+{tgt_path.name}")
                for s, t in pairs
            ]
            print(f"  {src_path.name}: {len(units)} bilingual pairs")
            all_units.extend(units)

    if not all_units:
        print("\n[ERROR] No translation units could be loaded. Check your input files.")
        sys.exit(1)

    print(f"\n[ANALYSE] Indexing {len(all_units):,} translation units…")
    analyser.ingest(all_units)
    stats = analyser.stats()
    clusters = analyser.build_clusters()

    # ── Output ────────────────────────────────────────────────────
    print_summary(clusters, stats)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = output_stem or f"consistency_report_{ts}"

    csv_path = Path(stem + ".csv")
    json_path = Path(stem + ".json")

    write_csv(clusters, csv_path)
    write_json(clusters, stats, json_path)

    print(f"\n  Reports saved:")
    print(f"    {csv_path.resolve()}")
    print(f"    {json_path.resolve()}")
    print()


# ─────────────────────────────────────────────
#  CLI
# ─────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="auditor",
        description="TERMINOLOGY CONSISTENCY AUDITOR — v0.1\n"
                    "Detects Eng→Spa translation inconsistencies across TMX and document corpora.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument(
        "--tmx", nargs="+", metavar="FILE.tmx",
        help="One or more TMX files (bilingual Eng/Spa)",
    )
    p.add_argument(
        "--src", nargs="+", metavar="FILE",
        help="Source (English) document(s): .docx, .pdf, .txt",
    )
    p.add_argument(
        "--tgt", nargs="+", metavar="FILE",
        help="Target (Spanish) document(s): must match --src count",
    )
    p.add_argument(
        "--min-freq", type=int, default=2, metavar="N",
        help="Minimum total occurrences for a term to appear in report (default: 2)",
    )
    p.add_argument(
        "--min-variants", type=int, default=2, metavar="N",
        help="Minimum number of distinct translations to flag as inconsistent (default: 2)",
    )
    p.add_argument(
        "--output", metavar="STEM", default="",
        help="Output file stem (default: consistency_report_<timestamp>)",
    )
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    if not args.tmx and not (args.src and args.tgt):
        parser.print_help()
        sys.exit(0)

    tmx_files = [Path(f) for f in (args.tmx or [])]
    src_files = [Path(f) for f in (args.src or [])]
    tgt_files = [Path(f) for f in (args.tgt or [])]

    for f in tmx_files + src_files + tgt_files:
        if not f.exists():
            print(f"[ERROR] File not found: {f}")
            sys.exit(1)

    run(
        tmx_files=tmx_files,
        src_files=src_files,
        tgt_files=tgt_files,
        spreadsheet_pairs=[],
        min_freq=args.min_freq,
        min_variants=args.min_variants,
        output_stem=args.output,
    )


if __name__ == "__main__":
    main()
