# Lead Validation & QA Pipeline
Python QA system: 15 leads validated with 7-point accuracy framework in < 1 second

## Problem
Fulfilment VA teams manually review each lead record before client handoff — a process prone to inconsistency at scale. Missing a blank job title, a broken LinkedIn URL, or a mis-formatted name can damage client trust and require costly re-work.

## Solution
A Python pipeline that ingests raw lead data (CSV), applies a 7-point validation framework (name completeness, LinkedIn URL regex, email format, U.S. state verification, executive tier classification, confidence scoring, auto-flag), and outputs a colour-coded Excel QA report compatible with Google Sheets upload.

## Results

| Metric | Before | After |
|---|---|---|
| QA time per lead | ~4 min manual | < 0.07 sec automated |
| Human error risk | High (manual scan) | Eliminated for format checks |
| Leads processed | N/A | 15 in < 1 second |
| Ready rate | N/A | 67% (10/15) |
| Avg confidence score | N/A | 87.7/100 |

## Tools Used
- Python · pandas · openpyxl
- Google Sheets (output-compatible XLSX)
- Excel with conditional colour formatting
- Regex validation

## Demo
[placeholder — GIF or video link]

## Links
- Notion: [this page]
- Contact: hoangnvm.hust@gmail.com
