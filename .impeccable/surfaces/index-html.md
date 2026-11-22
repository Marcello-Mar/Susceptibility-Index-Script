---
version: 1
slug: "index-html"
primary_target: "index.html"
related_targets: []
---

Scope: index.html, the whole app (single page). Visitor mode: Operate.
Audience/job: ID clinicians and researchers turning an isolate line-list into single-agent and two-drug coverage estimates for abstracts, antibiograms and stewardship protocols. Data never leaves the browser.
Constraints: English UI; static GitHub Pages; vendored libs; synthetic demo data labelled as such.

## Direction contract

THESIS: The app reads like a clinical breakpoint reference table that computes itself: dense, ruled, footnoted. It refuses the category default of a grey dashboard with KPI tiles and a lone accent colour.

OWN-WORLD: Cool blue-grey ground, deep petrol-ink shell band, tinted header bands per antibiotic class (penicillins, cephalosporins, carbapenems, aminoglycosides, quinolones, glycopeptides and so on), %S cells on a three-step threshold scale with the number always printed. One system sans with tabular figures; small caps for class labels; hairline rules, no cards with shadows.

STORY: The user sees the data is local, loads a file, confirms columns and Gram classes, then reads ranked agents and combinations, opening any figure to see how it was derived, and exports.

FIRST VIEWPORT: Ink header band with name and "runs locally" status; left rail of setup sections (Data, Columns, Organisms, Parameters) with live counts; main area with subset tabs (All, Gram+, Gram−, chosen genus/species) over the result views (Overview, Single agents, Combinations, Antibiogram) and an Export menu top right. Empty state: drop zone plus "Load synthetic demo".

FORM: Reference-table grammar, candidate 4 of 7. seed:f66b8f91df02d533. Signature move: every computed figure carries a numbered note marker; clicking it opens the derivation (numerator, denominator, interval, test) in a notes column, like breakpoint-table footnotes.

FINISH: unreviewed and undocumented is unfinished; this build ends with the finish review, the verdict, DESIGN.md, and every shipping raster carrying its provenance
