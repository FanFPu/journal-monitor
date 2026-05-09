---
name: cancer-multiomics-journal-monitor
description: Monitor a fixed set of oncology and top-tier biomedical journals, collect articles from the latest 30 days, keep only cancer-related and strict multi-omics papers, and generate an Excel monthly report with one dedicated Chinese abstract-summary column, one dedicated sample or patient-information column, and one dedicated dataset section covering repository name, accession numbers, whether data are downloadable, and download links when available. Use when the user asks for 肿瘤多组学期刊月报, 月度文献抓取, 癌症多组学监控, or periodic journal screening across Cancer Cell, Cancer Discovery, Nature Cancer, Cancer Research, Clinical Cancer Research, Nature family journals, Science journals, Cell journals, and related titles.
---

# Cancer Multiomics Journal Monitor

## Goal

Generate a monthly Excel report for a fixed journal list. Only include papers that are both:

- directly cancer related
- clearly multi-omics related, meaning at least two omics layers are integrated, or the paper explicitly uses single-cell/spatial multi-omics

The default report window is the latest 30 days counted from the run date.

## Workflow

1. Run the bundled script:

   ```bash
   python3 ~/.codex/skills/cancer-multiomics-journal-monitor/scripts/generate_monthly_report.py
   ```

2. The script writes the workbook to `./reports/journal_cancer_multiomics_YYYY-MM.xlsx` in the current workspace.
3. If the target file already exists, the script skips by default. Use `--force` to rebuild.
4. Use `--journal-limit N` for quick validation runs.

## Output Contract

The workbook must contain:

- `All_Articles`: all included papers
- one sheet per journal with included papers only
- `Sources`: journal source registry, fetch status, and candidate counts

Default article columns:

- `report_month`
- `journal`
- `publish_date`
- `title`
- `article_type`
- `cancer_type`
- `omics_layers`
- `article_url`
- `doi`
- `abstract_summary_cn`
- `sample_patient_info_cn`
- `database_name`
- `data_downloadable`
- `data_accessions`
- `geo_accessions`
- `data_download_url`
- `pubmed_url`

Column intent:

- `abstract_summary_cn`: one dedicated summary column for the paper abstract. This is not a full translation and not sentence-by-sentence paraphrase. It should be a compact Chinese summary based mainly on the abstract, rewritten in your own words.
- `sample_patient_info_cn`: one dedicated column for sample and patient information. Prefer concrete details such as patient count, case count, cohort size, lesion count, sample count, validation cohort, and if available major sample types.
- `database_name`: one dedicated column for the database or repository name, for example `GEO`, `SRA`, `BioProject`, `EGA`, `ArrayExpress`, `dbGaP`, `PRIDE`, or publisher-hosted supplementary repository.
- `data_downloadable`: explicitly mark whether the article states that data can be downloaded or deposited, for example `是`, `否`, or `摘要/元数据未明确`.
- `data_accessions`: one dedicated column for all accession numbers mentioned in the article metadata or abstract.
- `geo_accessions`: a dedicated GEO-only helper column. Fill only when GEO accessions are explicitly available.
- `data_download_url`: one dedicated column for concrete download or landing links for the dataset or repository entry. Prefer real accession landing pages over generic homepages.

## Filtering Rules

- Journal must be present in `references/journal_registry.json`.
- Publication date must fall within the latest 30 days unless the operator passes a different `--days`.
- Paper must be cancer related based on title or abstract.
- Paper must be strict multi-omics:
  - explicit multi-omics or multiomics phrasing, or
  - at least two omics layers detected from title or abstract, or
  - explicit single-cell/spatial multi-omics phrasing
- Pure single-omics papers are excluded by default even if they are cancer related.
- Reviews are kept only if they are still clearly about cancer multi-omics; they are labeled as reviews.

## Notes

- The bundled script uses PubMed as the universal discovery layer, then enriches article links with DOI resolution and collects dataset accessions from PubMed metadata plus abstract text.
- The `Sources` sheet is the journal-level link inventory. Update `references/journal_registry.json` when you want to add or remove journals.
- `abstract_summary_cn` is a standalone abstract-summary field. It should summarize the paper's abstract in Chinese with 2-4 compact sentences, written in your own words. Do not paste the original abstract, do not fully translate it line by line, and do not make up findings that are not supported by the abstract.
- `sample_patient_info_cn` is also a standalone field. It should focus only on sample source, patient count, cohort size, lesion or sample count, and any clearly stated validation cohort or sample type. If the abstract does not clearly provide this information, write `摘要未明确说明`.
- The dataset section should answer three separate questions whenever possible: `数据库名字是什么`、`数据能不能下载`、`下载链接是什么`。 These should be reflected through `database_name`, `data_downloadable`, and `data_download_url`.
- If accession numbers are present, place them in `data_accessions`, and copy GEO-specific ones into `geo_accessions`.
- If repository or download information is not clearly stated in metadata, abstract, or accessible landing pages, leave the unknown field blank or mark it as `未明确`; do not fabricate repository names or links.
