# Cancer Multiomics Journal Monitor

Codex skill for generating monthly Excel reports of cancer-related strict multi-omics papers from a fixed journal registry.

The skill discovers candidate papers through PubMed, filters for cancer relevance and multi-omics evidence, enriches DOI and dataset accession metadata, then writes an Excel workbook with article-level sheets and source status tracking.

## Repository Layout

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   └── journal_registry.json
└── scripts/
    └── generate_monthly_report.py
```

## Install

Clone this repository into the Codex skills directory:

```bash
git clone https://github.com/FanFPu/journal-monitor.git \
  ~/.codex/skills/cancer-multiomics-journal-monitor
```

Install Python dependencies:

```bash
python3 -m pip install -r ~/.codex/skills/cancer-multiomics-journal-monitor/requirements.txt
```

## Usage

Run from the workspace where the report should be written:

```bash
python3 ~/.codex/skills/cancer-multiomics-journal-monitor/scripts/generate_monthly_report.py
```

By default, the report covers the latest 30 days and writes:

```text
./reports/journal_cancer_multiomics_YYYY-MM.xlsx
```

Useful options:

```bash
# Rebuild even if the monthly workbook already exists
python3 ~/.codex/skills/cancer-multiomics-journal-monitor/scripts/generate_monthly_report.py --force

# Quick validation against the first N journals
python3 ~/.codex/skills/cancer-multiomics-journal-monitor/scripts/generate_monthly_report.py --journal-limit 3

# Custom date window
python3 ~/.codex/skills/cancer-multiomics-journal-monitor/scripts/generate_monthly_report.py --days 60
```

## Output

The workbook contains:

- `All_Articles`: all included papers
- one sheet per journal with included papers
- `Sources`: journal registry, query status, candidate count, and included count

Core article fields include title, journal, publication date, DOI, cancer type, omics layers, Chinese abstract summary, sample or patient information, data accession fields, dataset download links, and PubMed/article URLs.

## Journal Registry

Edit `references/journal_registry.json` to add, remove, or update journals. Each entry should include the display journal name, publisher, official URLs, and the PubMed journal name used for search.
