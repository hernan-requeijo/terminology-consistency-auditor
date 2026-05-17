# TERMINOLOGY CONSISTENCY AUDITOR

A standalone Python tool for detecting EN→ES terminology inconsistencies
across TMX corpora and parallel document pairs.

## Why I built this

As a professional EN>ES translator with 25 years of experience in medical and
legal texts, I found that terminology inconsistency is one of the most
time-consuming quality issues to detect manually. I built this tool to automate
the detection process across TMX corpora and parallel document pairs, using
Python's standard library with no external NLP dependencies.


## What it does

1. **Ingests** bilingual content from multiple sources in one run
2. **Extracts** recurring source phrases (1–5 words) across all segments
3. **Maps** each phrase to all Spanish translations found in the corpus
4. **Reports** phrases that have been translated in more than one way,
   ranked by severity and frequency

---

## Supported input formats

| Format | Mode | Notes |
|--------|------|-------|
| `.tmx` | Single bilingual file | Standard CAT tool export; en/es auto-detected |
| `.docx` | Source + target pair | One English file, one Spanish file |
| `.pdf` | Source + target pair | Text-based PDFs; scanned PDFs not supported |
| `.txt` | Source + target pair | Plain text, paragraph-separated |
| `.xlsx` | Bilingual columns | Col 0 = EN, Col 1 = ES (or auto-detect by header) |
| `.csv` | Bilingual columns | Same column logic as XLSX |

---

## Installation

Requires Python 3.9+ and these standard packages:

```bash
pip install python-docx pdfminer.six openpyxl lxml
```

No other dependencies. `difflib` is built into Python.

---

## Usage

### TMX files only
```bash
python auditor.py --tmx corpus_01.tmx corpus_02.tmx corpus_03.tmx
```

### Parallel document pair (one source, one translation)
```bash
python auditor.py --src manual_en.docx --tgt manual_es.docx
```

### PDF pair
```bash
python auditor.py --src report_en.pdf --tgt report_es.pdf
```

### Mix TMX + documents
```bash
python auditor.py \
  --tmx main_corpus.tmx legacy.tmx \
  --src addendum_en.docx --tgt addendum_es.docx
```

### Multiple document pairs (must be same count and order)
```bash
python auditor.py \
  --src doc1_en.pdf doc2_en.docx doc3_en.txt \
  --tgt doc1_es.pdf doc2_es.docx doc3_es.txt
```

### Tuning thresholds
```bash
# Only flag terms appearing ≥ 3 times with ≥ 3 different translations
python auditor.py --tmx corpus.tmx --min-freq 3 --min-variants 3

# Save report with custom name
python auditor.py --tmx corpus.tmx --output gdpr_audit_2025
```

---

## Output

Every run produces two files:

**`<stem>.csv`** — spreadsheet-ready report, one row per inconsistent term:
- Severity (CRITICAL / HIGH / MEDIUM / LOW)
- Source phrase (English)
- Number of distinct Spanish translations found
- Total occurrences across corpus
- Dominant translation (most frequent)
- All variants with occurrence counts
- Origin files where each variant appears

**`<stem>.json`** — structured report for programmatic processing,
including corpus statistics and full variant detail per cluster.

Console output shows the top 20 inconsistencies in a summary table.

---

## Severity scale

| Severity | Condition |
|----------|-----------|
| CRITICAL | 4+ distinct translations |
| HIGH | 3 distinct translations |
| MEDIUM | 2 translations, ≥ 5 total occurrences |
| LOW | 2 translations, < 5 total occurrences |

---

## How alignment works (parallel documents)

When you provide `--src` / `--tgt` document pairs, the tool aligns
the two files into bilingual segment pairs before analysis:

- **Perfect match** (same paragraph count): 1:1 zip
- **Near match** (≤ 20% difference): greedy length-ratio alignment
- **Large difference**: sentence-count-weighted block merge before zip

For best results with parallel documents:
- Use documents with consistent paragraph structure
- Avoid documents where images or tables interrupt text flow
- TMX files always produce the cleanest results

---

## Architecture notes (for developers)

```
auditor.py
├── TranslationUnit        — bilingual segment data class
├── TermCluster            — inconsistency cluster data class
├── TMXParser              — parse .tmx files
├── DocxParser             — extract paragraphs from .docx
├── PDFParser              — extract blocks from .pdf
├── TextParser             — read .txt files
├── SpreadsheetParser      — extract bilingual pairs from .xlsx/.csv
├── ParallelAligner        — align src/tgt doc blocks into TU pairs
├── CorpusAnalyser         — two-pass term extraction + inconsistency detection
│   ├── _source_phrases()  — 1-5 word phrase extraction (no stopword-only)
│   ├── _extract_target()  — position-ratio target phrase projection
│   └── build_clusters()   — rank and return TermCluster objects
├── write_csv()            — CSV report writer
├── write_json()           — JSON report writer
└── main() / run()         — CLI entry point
```

---

## Roadmap — next modules

- [ ] LLM context-aware false-positive filter (Claude API)
- [ ] Eng-Spa specific rule engine (gender agreement, ser/estar, cognates)
- [ ] TBX export for direct import into Trados / memoQ / Phrase
- [ ] TM health score (composite corpus quality metric)
- [ ] Retroactive termbase miner (promote stable pairs to glossary)
- [ ] Web UI wrapper (Flask/FastAPI)

---

## Known limitations (v0.1)

- Target phrase extraction for sub-segment terms uses position-ratio
  heuristics, not word alignment — accurate for technical terminology
  (short, dense phrases) but may produce window artifacts for long sentences.
- Scanned PDFs are not supported (no OCR layer in v0.1).
- Language auto-detection relies on xml:lang tags in TMX; unusual
  tag values may need the TMXParser.SRC_LANGS / TGT_LANGS sets extended.

---

*TERMINOLOGY CONSISTENCY Project — © 2025. Internal development tool.*
