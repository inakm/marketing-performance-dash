# LOGIC.md — How the Dashboard Worked

This document describes the internal logic of the dashboard for future reference. All processing is client-side JavaScript in `script.js`; no server exists.

## Overview

The app is a two-stage pipeline:

```
Excel files  ->  Read & normalize  ->  Classify & aggregate  ->  Render charts/search
```

Two inputs are supported:

1. **AKM Marketing Report** (`.xlsx`, required) — primary data. Sheet name `DATA`.
2. **Neha's Remarks File** (`.xlsx`, optional) — secondary call-remark data. The first sheet whose name contains `neha` or `remark` is used (fallback: first sheet).

## Column layout (0-indexed rows after a header row)

### AKM `DATA` sheet

| Index | Field       | Notes                                 |
| ----- | ----------- | ------------------------------------- |
| 0     | —           | Unused by logic (often serial no.)    |
| 1     | `date`      | Parsed by `normalizeDate`             |
| 2     | `person`    | Normalized via `normalizePersonName`  |
| 3     | `clientName`| Trimmed, dedupe key for unique visits |
| 4     | `priority`  | `1`–`4` (string coerced to int)       |
| 5     | `dsoRemark` | Free-text field used for classification |
| 6     | `area`      | Geographic area string                |
| 7     | `km`        | Distance, parsed via `parseFloat`     |
| 8     | `phone`     | Counted when length ≥ 7               |
| 9     | `email`     | Counted when it contains `@`          |

A row where `person` normalizes to `ABSENT` (equals "ABSENT" case-insensitively) is recorded as an absence day with the remark `ABSENT` and a `date`.

### Neha remarks sheet

| Index | Field         |
| ----- | ------------- |
| 1     | `date`        |
| 2     | `person`      |
| 3     | `clientName`  |
| 4     | `area`        |
| 5     | `email`       |
| 6     | `phone`       |
| 7     | `nehaRemark`  |
| 8     | `dsoRemark`   |

## Reading & normalization

- `readExcel` loads the workbook with `cellFormula: true` and `cellDates: true`, so formula results and real dates survive.
- `sheetToArray` converts each sheet to a 2D array (`header: 1`, empty cells default to `''`, raw values kept).
- `resolveAllRefs` / `resolveCellRef` resolves `=CELL` references such as `=B5` by looking up the referred cell's value in the same array. This handles reports that alias one column to another via formula.
- `normalizePersonName` collapses whitespace, fixes common misspellings (`kumar r` → `Kumar`, `kuma r` → `Kumar`), keeps `Kumar & Praveen` as a shared-person token, and title-cases the result.

## Lead-status classification (`classifyStatus`)

Each AKM `dsoRemark` is lowercased and matched against keyword rules **in order of precedence**:

1. `ABSENT` — remark is `absent` (< 20 chars) or contains "not went to market".
2. `Not Interested` — contains "not interested".
3. `No Requirement` — "not needed", "no requirement", "presently not needed", "no requir".
4. `Already has Vendor` — "already vendor", "already order given", "already family dealer", "already given order", "already regentra", "already purchased".
5. `Hot Lead` — "order received", "delivered", or any quotation request ("send quotation", "based on quotation", ...), or the word "interested".
6. `Sample Needed` — "send/give sample", "sample testing", "sample we check", "give me some sample".
7. `Not Reachable` — "busy", "not responding", "wrong number", "call disconnected", "unreachable", "not available/aviliable", "security not allowed", "photos not allowed", "concern person not available", "not there", "in meeting", "out of city/station/town", "sir/owner/manager is out", "head sir is".
8. `Follow-up Needed` — any future-tense signal: "tomorrow"/"tommrow"/"tommorw"/"tommorow", "will update/contact/discuss/check/call/text", "come tomorrow", "come next time", "check & update", "month end", "next month", "follow up".

Anything unmatched falls back to `Follow-up Needed`.

`classifyNehaRemark` mirrors this for the Neha file with categories: `No Remark`, `Not Interested`, `No Requirement`, `Not Reachable`, `Interested`, `Quotation Requested`, `Follow-up Pending`, `Email Shared` (contains "email"/"mail"), else `Other`.

`extractActionItems` converts remark keywords into human action strings joined by `|`, e.g. "Send quotation | Awaiting update".

`normalizeDate` accepts `Date` objects, Excel serial numbers (`40000–60000`), ISO strings, `YYYY-MM-DD`, and date-major/minor slash or dash formats, outputting `YYYY-MM-DD`.

## Aggregation (`processAllData`)

- **Persons** are derived from every non-`ABSENT` AKM `person` cell, splitting on `&` so "Kumar & Praveen" feeds both person tabs.
- Per person it computes:
  - `priorityBreakdown` {1,2,3,4} counts
  - `statusBreakdown` via `classifyStatus`
  - `areaMap` per-area lead counts
  - `dailyMap` per-date lead counts
  - `totalKm` (rounded), `phones`, `emails`
  - `uniqueVisits` = distinct lowercased `clientName`s
- **Summary** aggregates across all leads: totals, P1 count, follow-up count (statuses `Follow-up Needed`, `Quotation Needed`, `Sample Needed`), `absentDays`, `uniqueClients`.
- **Follow-ups** collect leads whose status is one of `Follow-up Needed`, `Quotation Needed`, `Sample Needed`, `Hot Lead`, enriched with `actionItems`, sorted Hot Lead → Quotation → Sample → Follow-up.
- **Search index** (`allClientNames`) merges AKM leads with Neha rows by lowercased client name, attaching Neha remarks to the AKM record when matched.
- **Neha analysis** buckets remark categories, DSO responses, and a daily timeline.

## Rendering

- ApexCharts instances are tracked in a `charts` array and destroyed before each re-render (`destroyCharts`) to avoid leaks when switching person tabs.
- Person content is injected with unique chart-container IDs (`id + Date.now() + random suffix`) so charts re-attach correctly.
- Charts: priority donut, status horizontal bar, daily activity area chart, top-12 area bar (leads), top-12 area bar (KM), plus Neha pie/bar/timeline.
- Dark theme constants: `#f3c623` gold accent, `#3b82f6` blue, `#8b5cf6` purple, per-person colors `Praveen`, `Pavan`, `Kumar`, `Kumar & Praveen`.
- The search dropdown filters `allClientNames` on 2+ characters across client, DSO remark, Neha remark, area, and person; top 20 results shown; clicking anywhere outside the bar dismisses it.

## Dependencies

Loaded via CDN in `index.html`:

- ApexCharts — `cdn.jsdelivr.net/npm/apexcharts`
- SheetJS — `cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`
- Font Awesome 6.5.1 — `cdnjs.cloudflare.com/.../font-awesome/6.5.1/css/all.min.css`
- Manrope font — Google Fonts

## Data retention

Uploaded files are held only as in-memory `File`/workbook references. Refreshing the page or clearing files discards everything — no storage APIs are used.