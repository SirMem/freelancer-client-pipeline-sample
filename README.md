# Freelancer Client & Project Pipeline — Free Sample

**Trust hub (disclosure · ledger · refund):** https://sirmem.github.io/freelancer-money-sheets/

A complete, usable **mini CRM pipeline**: client roster + projects with status, fee type, auto Value, Paid?, due date, and Overdue flag. Structural seam of the paid **Freelancer Client & Project Tracker** — same quality Clients + Projects core, not a crippled demo.

## What's inside

| Sheet | Role |
|-------|------|
| **Start Here** | TOC seam map + gap questions (PIPELINE / TIMELOG / SETTINGS) |
| **Clients** | Small roster (name / email / company / notes / Active?) |
| **Projects** | Status Lead/Active/Waiting/Done/Lost · Fee type Hourly/Flat · auto Value · Paid? · Due · Overdue · pipeline snapshot totals |

- EXAMPLE rows only (fictional clients/jobs) — replace with your own.
- Yellow cells are inputs; Value / Overdue / snapshot totals are formulas.
- Hand-check on sample data (LibreOffice recalc, as of 2026-10-28): **Active 2** · **pipeline $5,650** · **won $900** · **paid $900** · **overdue 1** (Landing page due 2026-10-20, Waiting).

## What this sample does **not** do

It does **not** include Settings (default rate / currency), Time Log (hours roll into projects), or a Dashboard (by-status / by-client KPIs). Those live in the paid Freelancer Client & Project Tracker. Hourly Value requires Rate filled on the project row (no studio default-rate fallback).

## Download

Raw `.xlsx`:

https://raw.githubusercontent.com/SirMem/freelancer-client-pipeline-sample/main/Freelancer-Client-Pipeline-Sample.xlsx

## Gap questions (optional — only if one fits)

If one fits after you try the sample, open a GitHub issue with **one word**:

1. **PIPELINE** — Dashboard of pipeline vs won (by status + by client) without a second sheet?
2. **TIMELOG** — Hours logged against a project that roll into Hours logged automatically?
3. **SETTINGS** — Default rate / currency / studio settings so blank project rates inherit a studio default?

If none apply, keep the free sheet — no pitch.

## Rebuild

```bash
python3 build.py
```

Requires `openpyxl`.

## Source / methodology

Method: standard freelance project-tracking practice (client roster + status pipeline + fee type + due dates). Example data is fictional — not personal client data. Not accounting software or tax advice.

## License

Free for personal use. Example data is fictional. Not accounting or tax advice.
