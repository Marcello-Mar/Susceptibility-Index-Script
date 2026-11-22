# Susceptibility Index

Browser tool that turns an isolate line-list (one row per isolate, one column per antibiotic) into susceptibility and coverage estimates for single agents and every two-drug combination, overall and by Gram group, genus or species.

**Open it:** https://marcello-mar.github.io/Susceptibility-Index-Script/

The page reads your file locally. Nothing is uploaded: there is no server, and the page's Content-Security-Policy blocks network requests (`connect-src 'none'`). Closing the tab discards the session. It also works offline: download the repository and open `index.html`.

## Input

Excel (.xlsx, .xls) or CSV, first row = headers.

| Microorganism | Patient ID | Collection date | Ampicillin | Gentamicin | Ciprofloxacin | … |
|---|---|---|---|---|---|---|
| Escherichia coli | P0017 | 2025-03-02 | R | S | S | … |
| Staphylococcus aureus | P0021 | 2025-03-04 | R | S | | … |

- Column order and names are free: organism, patient ID, date and antibiotic columns are mapped on screen (with automatic detection).
- Values: `S`, `I`, `R` (case-insensitive; `susceptible`/`resistant`/`sensibile`/`resistente` also accepted). Empty, `N`, `NT`, `NA`, `-` = not tested. Any other value is counted as not tested and listed in a warning.
- Patient ID and date are optional; they are only used to keep the first isolate per patient.

## Settings

| Setting | Options | v1 behaviour |
|---|---|---|
| I results | count as S, as R, or as not tested | as S |
| Repeat isolates | keep all; first per patient and organism; first per patient (earliest date, else file order) | keep all |
| Exclusions | organism names containing given text (typed or from a file) | exclusion file |
| Minimum tested | agents tested in fewer than X% of the subset's isolates are dropped | same |
| Confidence level | 90, 95, 99% | free |
| Combination denominator | tested for at least one agent; tested for both | at least one |
| Colour bands | thresholds for green/amber/red | – |

Organism groups (Gram-positive, Gram-negative, fungi, other) come from a bundled map of 250 genera and common aliases, editable in the Organisms view. Matching uses the exact name first, then the longest map entry found at the start of a word. Names not in the map are listed for assignment. The map can be exported and re-imported (same two-column format as v1: `Microorganism`, `Gram_stain`).

All settings, including column mapping and organism assignments, can be saved as a JSON file and loaded in a later session.

## Statistics

For each agent and each pair, in each subset:

- **%S** = S / (S + R) among isolates with a result.
- **Coverage** = S / all isolates in the subset; untested isolates count as not covered.
- **Confidence intervals**: Jeffreys, quantiles of Beta(S + ½, failures + ½); lower bound 0 when S = 0, upper bound 1 when there are no failures.
- **Combination**: S when at least one agent is S.
- **Gain over a single agent**: isolates R to that agent and S to its partner, over the isolates tested for that agent. p-value: exact two-sided McNemar test on the same isolates (paired), not adjusted for multiple comparisons.

Every figure in the tables opens its derivation (counts, denominator, interval, test) in the notes column.

## Output

- **Results workbook (.xlsx)**: Summary, Methods (parameters and notes), singles and combinations for All, Gram+, Gram− and the selected genus/species, antibiogram, data checks.
- **Current table (.csv)**, with filters and sorting as on screen.
- **Figures** (.svg, or .png at 300 dpi): single-agent interval plot, combination matrix, organism × antibiotic antibiogram.
- **Printable report** (print or save as PDF).
- **Organism map (.xlsx)** and **settings (.json)**.

## Changes from v1 (Colab notebook)

- Graphical interface; no Python or Colab needed.
- Combination vs single agent: exact McNemar on paired data replaces the chi-square for independent samples.
- Combination tables report "R to both" and "R to one, other untested" separately; in v1 the exported `r_count` (R to both) did not add up with the percentage when results were missing.
- Organism classification no longer depends on the order of the map file (v1 took the first entry contained in the name).
- New: I handling, deduplication, genus/species subsets, combination denominator option, figures, report, settings files.
- With repeat isolates kept, I as S and the "at least one agent" denominator, %S, coverage and intervals equal those of v1 (`tests/validate.py` compares the two on the synthetic dataset; maximum difference 5 × 10⁻¹¹ percentage points).

## Development

Static files only: `index.html`, `css/`, `js/engine.js` (calculations, no DOM), `js/charts.js` (SVG figures), `js/app.js` (interface). SheetJS Community Edition 0.20.3 is vendored in `vendor/` (Apache-2.0).

```
python3 tests/make_synthetic.py data/synthetic_demo.xlsx   # regenerate the demo data
python3 tests/validate.py <v1 notebook .ipynb>              # compare with the v1 Python logic
```

`data/synthetic_demo.*` is invented data for the demo and the tests, not patient data. The v1 notebooks are in the git history (commit `1ea31e4`).

## License

GPL-3.0, see `LICENSE`.
