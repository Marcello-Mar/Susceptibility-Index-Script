# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Static HTML/CSS/JS, no build server, published with GitHub Pages from the repository root. Libraries vendored locally (no CDN calls at runtime) where possible. Delegated choice: plain JS modules, SheetJS for spreadsheet I/O.

## Users

Clinicians and clinical researchers in (pediatric) infectious diseases and antimicrobial stewardship who hold a line-list of isolates with S/I/R results exported from the microbiology lab, and want empiric-therapy coverage estimates for single agents and two-drug combinations, by Gram group or genus/species.

## Product Purpose

Replace the Colab notebook "Susceptibility Index Script" with a browser tool: load a file, map its columns, set parameters, and explore susceptibility and coverage estimates that recompute live; export the results for papers, posters and local guidelines.

## Positioning

Runs entirely in the browser: patient-level microbiology data never leaves the user's machine. Reports two denominators side by side (susceptible among tested, and susceptible among all isolates with untested counted as not covered), with Jeffreys intervals and paired McNemar tests for combination vs single agent.

## Operating Context

Desktop browser, institutional PCs, often while preparing an abstract, a local antibiogram or a stewardship protocol. Input files are Excel/CSV exports, one row per isolate, one column per antibiotic. Outputs end up in Excel, Word manuscripts and conference slides.

## Capabilities and Constraints

- Input: .xlsx/.xls/.csv; user maps organism column, antibiotic columns, optional patient ID and date columns.
- Values S/R/N (I configurable as S, R or not tested); empty = not tested.
- Gram classification from an editable map (default bundled, genus-based, longest-key match), unmapped organisms resolved in-app; map exportable.
- Optional exclusion list, deduplication (first isolate per patient, optionally per organism), minimum-tested threshold, confidence level, subgroup by Gram or genus/species.
- No server, no storage beyond the session; exports xlsx, csv, svg/png charts, printable report.

## Evidence on Hand

Bundled Gram map (229 genera) from the previous repo. No real patient dataset in the repo; any demo dataset must be synthetic and labelled as such.

## Product Principles

1. Data stays local.
2. Every number traceable: counts beside every percentage.
3. Defaults match the previous script; changes in method are stated in the UI.
4. Recompute on every parameter change, without a "run" button.

## Accessibility & Inclusion

Colour never the only carrier of S/R meaning; keyboard-operable tables and controls.
