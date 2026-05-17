# Roadmap

## Current version: v0.2

Core engine is complete and validated on real bilingual corpora.
Two-pass phrase extraction → source→target mapping → inconsistency clustering → severity scoring.

---

## Planned modules

### ✅ v0.1–v0.2 — Core engine (complete)
- TMX, DOCX, PDF, TXT, XLSX, CSV parsing
- Parallel document alignment (1:1, greedy, sentence-weighted)
- Two-pass term extraction and inconsistency clustering
- CSV + JSON report output
- Extraction quality improvements: ES function-word stripping,
  EN→ES length-ratio correction, single-token noise filtering

### 🔲 v0.3 — Extraction refinements
- Wider window for 2-word source phrases
- Verb-phrase filter (suppress "X is Y" patterns from term candidates)
- Near-duplicate cluster deduplication
- All-caps / heading token filter

### 🔲 v0.4 — TBX export
- Export validated term pairs to TermBase eXchange format
- Compatible with memoQ, Trados, Phrase termbase import

### 🔲 v0.5 — Spanish linguistic rule engine
- Gender agreement propagation detection
- Latin American vs. Castilian variant flagging
- False cognate detection (actual/current, embarazada/embarrassed, etc.)
- ser/estar disambiguation in translated state descriptions

### 🔲 v0.6 — LLM filter
- Optional Claude API integration for context-aware false-positive reduction
- Distinguishes genuine inconsistencies from acceptable stylistic variation

### 🔲 v0.7 — Web UI
- Local Flask drag-and-drop interface
- No cloud dependency; runs entirely on localhost