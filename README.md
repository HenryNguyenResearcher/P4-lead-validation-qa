# Lead Validation & QA Pipeline

**Automated QA system for LinkedIn lead research processes 15 leads, scores confidence 0–100, flags incomplete records in under 1 second.**

---

## Problem

Fulfilment VA teams manually review each lead record to check completeness and accuracy before handing off to clients - a process prone to human error and inconsistency, especially at scale. A single missed field or broken LinkedIn URL can damage client trust.

## Solution

A Python pipeline that ingests raw lead data (CSV), applies a 7-point validation framework, scores each record 0–100, and outputs a colour-coded Excel QA report ready for Google Sheets upload. Built to mirror the exact workflow of a LinkedIn research VA role.

**Validation framework** (maps to real VA QA criteria):

- Name completeness (first + last)
- Job title presence + U.S. executive tier classification (C-Suite, VP, Director, Manager…)
- LinkedIn URL format validation (regex against linkedin.com/in/ pattern)
- Email format validation
- U.S. location verification (50-state abbreviation lookup)
- Confidence score (0–100) → auto-assigns status: Ready / Review / Flag for QA

## Results

| Metric | Value |
| --- | --- |
| Leads processed | 15 |
| Ready (score ≥85, no flags) | 10 (67%) |
| Needs review | 5 (33%) |
| Avg confidence score | 87.7/100 |
| Processing time | < 1 second |

## Tools Used

- Python · pandas · openpyxl
- Google Sheets (output-compatible CSV/XLSX)
- Regex validation
- Excel with conditional formatting (colour-coded by status)

## Output

`output/validated_leads.xlsx` - two sheets:

1. **Validated Leads** - full record with QA Status, Confidence Score, Flags, colour-coded rows
2. **QA Summary** - aggregate metrics for client reporting

## How to Run

```bash
pip install pandas openpyxl
python validate_leads.py
```

Input: `data/raw_leads.csv`
Output: `output/validated_leads.xlsx`

## Links

- Notion case study: [notion.so/Lead-Validation-QA-Pipeline](https://app.notion.com/p/3525000d62a881c6ab19dafe150b3ed4)
- Contact: hoangnvm.hust@gmail.com
