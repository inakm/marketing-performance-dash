# SkyLimit Marketing Performance Dashboard

A client-side, privacy-first analytics dashboard for tracking marketing team performance from Excel reports. Upload the team's daily marketing reports and the dashboard generates lead-status classification, priority breakdowns, area/travel coverage, unique client visits, follow-up tracking, and searchable records — entirely in the browser.

> **Status:** Archived. This repository is maintained in read-only form for reference and reuse.

## Features

- **Excel upload & parse** — drop in the `AKM Marketing Report` (primary) and optionally `Neha's Remarks` file (`.xlsx` / `.xls`). Uses SheetJS with formula-cell and date resolution.
- **Lead status classification** — each DSO remark is scanned for keywords and tagged as `Hot Lead`, `Follow-up Needed`, `Quotation/Sample Needed`, `No Requirement`, `Already has Vendor`, `Not Reachable`, `Not Interested`, or `ABSENT`.
- **Summary KPIs** — total leads, active persons, total KM traveled, Priority-1 leads, unique clients, phones/emails collected, and pending follow-ups.
- **Per-person dashboards** — tabs per marketing person with priority donut, status analysis, daily activity, area coverage (leads and KM), and unique-visit KPI cards.
- **Follow-up tracker** — auto-extracted action items (e.g. "Send quotation", "Visit tomorrow") with client, area, date, and priority.
- **Global search** — real-time search across client name, person, area, DSO remark, and Neha's remark.
- **Neha's remarks analysis** — remark-category pie, DSO response status, and daily volume timeline.

## How to run

No build step, no server, no dependencies to install — it's a static site.

1. Open `index.html` in any modern browser, or serve the folder over HTTP:
   `npx serve .`
2. Upload the `AKM Marketing Report` (required). The optional `Neha's Remarks` file enables the remarks analysis section.
3. Click **Generate Dashboard**. Charts render with ApexCharts; all processing happens locally.

## Privacy

All data is processed in the browser via `FileReader` — nothing is uploaded to any server, stored, or shared. See the in-app footer for the full statement.

## Tech stack

- Vanilla HTML + CSS (custom dark "liquid glass" theme, Manrope font)
- [ApexCharts](https://apexcharts.com) — interactive charts
- [SheetJS (xlsx)](https://sheetjs.com) — Excel parsing
- [Font Awesome](https://fontawesome.com) — icons

## Repository structure

| File         | Purpose                                                        |
| ------------ | -------------------------------------------------------------- |
| `index.html` | Single-page layout: upload, dashboard, FAQs, footer             |
| `script.js`  | All logic: parsing, classification, aggregation, charts, search |
| `style.css`  | Styling and theme                                               |
| `DESIGN.md`  | Design-language reference (Sentry-inspired) used for theme       |
| `LOGIC.md`   | Detailed explanation of data flow and classification internals   |

## License & credits

Maintained by [Anjani Kumar Mishra](https://inakm.github.io/). Built by the [F9XR Team](https://f9xr.org/).