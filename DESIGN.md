---
name: Susceptibility Index
description: A clinical breakpoint reference table that computes itself; dense, ruled, footnoted, local-only.
colors:
  ground: "#e9eef1"
  paper: "#fbfcfd"
  paper-2: "#f3f6f8"
  ink: "#13252d"
  ink-2: "#4a626d"
  ink-3: "#6d838d"
  rule: "#cfd9de"
  rule-strong: "#93a7b0"
  band: "#0d3743"
  band-2: "#134a58"
  band-ink: "#e6f1f4"
  band-ink-2: "#a5c4cc"
  accent: "#0f5c6e"
  accent-ink: "#ffffff"
  accent-soft: "#d5e8ed"
  focus: "#1b7f97"
  sel: "#fff3c4"
  good: "#236b4c"
  good-bg: "#d7ebe0"
  mid: "#7a5700"
  mid-bg: "#f4e6bf"
  low: "#9c2f27"
  low-bg: "#f4d6d1"
  na-bg: "#eef1f3"
  gram-positive: "#5b3f9a"
  gram-negative: "#b3325a"
  gram-fungus: "#7b6a2c"
  gram-unclassified: "#6d838d"
typography:
  headline:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "1.25rem"
    fontWeight: 650
    lineHeight: 1.45
    letterSpacing: "-0.01em"
    fontFeature: "tnum"
  title:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.875rem"
    fontWeight: 650
    lineHeight: 1.45
    fontFeature: "tnum"
  body:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.875rem"
    fontWeight: 400
    lineHeight: 1.45
    fontFeature: "tnum"
  table:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.8125rem"
    fontWeight: 400
    lineHeight: 1.45
    fontFeature: "tnum"
  label:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.75rem"
    fontWeight: 600
    lineHeight: 1.45
    fontFeature: "tnum"
  section-caps:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.8125rem"
    fontWeight: 650
    letterSpacing: "0.06em"
  class-caps:
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif"
    fontSize: "0.6875rem"
    fontWeight: 650
    letterSpacing: "0.08em"
  derivation:
    fontFamily: "ui-monospace, SF Mono, Cascadia Mono, Menlo, Consolas, monospace"
    fontSize: "0.75rem"
    fontWeight: 400
rounded:
  hairline: "2px"
  r: "3px"
spacing:
  xs: "4px"
  sm: "8px"
  md: "12px"
  lg: "16px"
  xl: "20px"
  rail: "320px"
  notes: "280px"
  band-height: "52px"
components:
  button:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    typography: "{typography.table}"
    rounded: "{rounded.r}"
    padding: "0 12px"
    height: "30px"
  button-primary:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-ink}"
    typography: "{typography.table}"
    rounded: "{rounded.r}"
    padding: "0 12px"
    height: "30px"
  button-ghost:
    backgroundColor: "transparent"
    textColor: "{colors.accent}"
    rounded: "{rounded.r}"
    padding: "0 6px"
    height: "30px"
  button-ghost-hover:
    backgroundColor: "{colors.accent-soft}"
  band-button:
    backgroundColor: "{colors.band-2}"
    textColor: "{colors.band-ink}"
    typography: "{typography.table}"
    rounded: "{rounded.r}"
    padding: "0 12px"
    height: "32px"
  segmented-on:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.accent-ink}"
    height: "28px"
    padding: "0 10px"
  input:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    rounded: "{rounded.r}"
    padding: "0 8px"
    height: "30px"
  pct-good:
    backgroundColor: "{colors.good-bg}"
    textColor: "{colors.good}"
    rounded: "{rounded.hairline}"
    padding: "1px 6px"
  pct-mid:
    backgroundColor: "{colors.mid-bg}"
    textColor: "{colors.mid}"
    rounded: "{rounded.hairline}"
    padding: "1px 6px"
  pct-low:
    backgroundColor: "{colors.low-bg}"
    textColor: "{colors.low}"
    rounded: "{rounded.hairline}"
    padding: "1px 6px"
  sheet:
    backgroundColor: "{colors.paper}"
    rounded: "{rounded.r}"
  header-band:
    backgroundColor: "{colors.band}"
    textColor: "{colors.band-ink}"
    height: "52px"
    padding: "0 20px"
---

# Design System: Susceptibility Index

## Overview

**Creative North Star: "The Breakpoint Table That Computes Itself"**

The interface reads like a printed clinical breakpoint reference (EUCAST/CLSI-style tables): a dense, ruled sheet of agents grouped under tinted class bands, every figure printed with its count, and a numbered footnote apparatus in a right-hand notes column. It is an operating tool for a desk session, not a dashboard: there are no KPI tiles, no hero, no cards floating on shadows. Structure is carried by hairline rules, tonal steps between ground, rail and sheet, and a single deep petrol-ink header band.

Density is high and deliberate: 13px table text, 5px by 8px cell padding, tabular figures everywhere. Colour is spent on meaning in three places only: the three-step %S threshold scale, the per-class band tints and edge marks, and the Gram-group dots. Chrome stays blue-grey and quiet so those carriers stay legible.

**Key Characteristics:**
- Cool blue-grey ground with near-white paper sheets; one dark petrol band as the shell.
- Hairline rules (1px) and tonal layering instead of shadows.
- One system sans with tabular figures for every role; monospace only for derivation equations.
- Every computed figure carries a numbered note marker that opens its derivation in the notes column.
- %S always printed as a number inside its threshold cell; colour is never the sole carrier.

## Colors

A cool, desaturated blue-grey world anchored by petrol ink, with saturated colour reserved for data semantics.

### Primary
- **Petrol Teal** (`accent`): primary buttons, the pressed segment, links, note markers, ghost buttons, checkbox and range accents, the coverage diamond in interval marks.
- **Petrol Ink Band** (`band`, `band-2`): the sticky 52px header band, the active subset-tab underline (3px), the toast, bar-chart bars (`band-2`). `band-ink` and `band-ink-2` are its text tones.

### Secondary
- **Threshold Scale** (`good`/`good-bg`, `mid`/`mid-bg`, `low`/`low-bg`, `na-bg`): the three-step %S scale (green from 90%, amber from 70%, red below; user-adjustable) on %S chips, antibiogram cells and matrix cells. Each step is a dark text tone on its pale tint; untested is pale grey with a dash.
- **Antibiotic Class Bands** (per `[data-class]`: `--cb` tint, `--ce` edge): 21 class pairs, from Penicillins (`#dce7f5` / `#2f5f9e`) to Other (`#e7ebee` / `#5d6f78`). The tint fills class header rows, antibiogram column heads and agent-class selects; the edge colours the class label text and the 8px class square beside each agent. These are defined in `css/app.css` and mirrored literally in `js/charts.js` for exports.

### Tertiary
- **Gram Dots** (`gram-positive` violet, `gram-negative` raspberry, `gram-fungus` olive, `gram-unclassified` grey): 8px round dots on subset tabs and swatches in the Organisms counts.

### Neutral
- **Cool Ground** (`ground`): page background behind the rail and sheets.
- **Paper** (`paper`) and **Paper Shade** (`paper-2`): sheets, inputs, popovers; `paper-2` for the rail, sticky table heads, derivation equations and the empty-state side panel.
- **Ink ramp** (`ink`, `ink-2`, `ink-3`): primary text, secondary text and labels, tertiary hints and counts.
- **Rules** (`rule`, `rule-strong`): hairline dividers and sheet borders; `rule-strong` for input and button strokes and the table-head underline.
- **Selection Butter** (`sel`): selected table row and highlighted footnote.
- **Focus Teal** (`focus`): the 2px focus outline.

### Named Rules
**The Meaning-Only Colour Rule.** Saturated colour appears only where it encodes data: threshold step, antibiotic class, Gram group. Chrome is ink, paper and rule.

**The Number-Always-Printed Rule.** A threshold colour never stands alone: the %S value is printed inside every coloured cell, and the band boundaries are stated in a note.

## Typography

**Body Font:** system-ui (with -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial)
**Label/Mono Font:** ui-monospace (SF Mono, Cascadia Mono, Menlo, Consolas), derivation equations only

**Character:** One neutral UI sans set small and tight, with tabular numerals on the body so every column of counts aligns; hierarchy comes from weight (650 headings, 600 labels) and case, not size.

### Hierarchy
- **Headline** (650, 1.25rem; 1.5rem on the empty state, -0.01 to -0.015em): sheet titles ("Single agents", "Antibiogram").
- **Title** (650, 0.875rem): sheet section heads, ledger heads, derivation agent name.
- **Body** (400, 0.875rem, 1.45): base text; prose in the empty state capped at 62ch.
- **Table** (400, 0.8125rem): reference tables, controls, buttons, notes column.
- **Label** (600, 0.75rem): field labels, table heads (650), hints (400).
- **Section caps** (650, 0.8125rem, uppercase, 0.06em): numbered rail step headings ("1 DATA"); notes-column headings use 0.75rem, 0.08em with a 2px ink underline.
- **Class caps** (650, 0.6875rem, uppercase, 0.08em): antibiotic class band rows, in the class edge colour.

### Named Rules
**The Tabular Figures Rule.** `font-variant-numeric: tabular-nums` is set on the body; every numeric column right-aligns.

**The Weight-Not-Size Rule.** The ramp spans 11px to 24px; distinctions inside a sheet are made with weight, case and ink tone.

## Layout

A fixed application grid: a 52px sticky band; below it a 320px setup rail (sticky, scrolls independently) beside the main column. The main column stacks subset tabs (on paper, ruled below), view tabs (file-folder tabs that join the sheet), then a work grid of the sheet plus a 280px sticky notes column. Gutters are 20px; sheet sections pad 16px; rail steps pad 18px.

Spacing is a 4px-based rhythm (4, 8, 10, 12, 14, 16, 18, 20) applied tightly; table cells are 5px by 8px. Breakpoints: at 1560px the rail narrows to 296px and notes to 260px; at 1360px notes drop below the sheet; at 820px the rail stacks above the main column, band buttons become icon-only and the empty state goes single-column. Print hides the app and shows a dedicated report with point-based sizes.

## Elevation & Depth

Flat. Depth is tonal (ground, paper-2, paper) and ruled (1px `rule`, 1px `rule-strong` under table heads, 2px `ink` under notes headings, 1px `ink` above totals). Shadows exist only on two transient overlays that must lift off the sheet.

### Shadow Vocabulary
- **Popover** (`box-shadow: 0 10px 30px -8px rgba(0,0,0,0.28), 0 2px 6px rgba(0,0,0,0.1)`): the Export menu.
- **Toast** (`box-shadow: 0 8px 24px -6px rgba(0,0,0,0.35)`): the status toast on the band colour.

### Named Rules
**The Ruled-Not-Raised Rule.** Persistent surfaces never take a shadow; separate them with a hairline rule or a tonal step. Shadows are reserved for transient overlays.

## Shapes

Nearly square. One 3px radius (`r`) for controls, sheets, panels and popovers; 2px for value chips, class squares, menu rows and footnote highlights. Circles only for the Gram dots, the local-status dot and the %S point mark; the coverage mark is a 45-degree diamond. The selected view tab rounds only its top corners and the sheet squares its top-left corner so the two join like a folder tab. Dashed 1px border only on the file drop zone.

## Components

### Buttons
- **Shape:** gently squared (3px), 30px tall, 12px horizontal padding; 24px small variant.
- **Default:** paper fill, `rule-strong` stroke, ink text; hover lifts to white with an `ink-3` stroke (150ms, `cubic-bezier(0.16, 1, 0.3, 1)`).
- **Primary:** petrol teal fill and stroke, white text; hover darkens.
- **Ghost:** no fill or stroke, teal text, 6px padding; hover fills `accent-soft`.
- **Band button:** 32px, `band-2` fill, a darker teal stroke, band-ink text with a 16px stroked icon; icon-only below 820px.

### Segmented Control
Joined buttons in one `rule-strong` frame, 28px tall, separated by hairlines; the pressed segment fills petrol teal with white text.

### Inputs / Fields
- **Style:** 30px, paper fill, 1px `rule-strong` stroke, 3px radius; label above in 12px/600 `ink-2`, hint below in `ink-3`.
- **Hover:** stroke darkens to `ink-3`. **Focus:** global 2px `focus` outline, 2px offset.
- **Disabled step:** the whole rail step drops to 50% opacity and ignores pointer events.

### Navigation
- **Subset tabs:** text tabs on a paper strip with Gram dot and count; the selected tab gets a 3px band-colour underline. A genus/species select sits after a hairline and takes a 3px band underline when active.
- **View tabs:** folder tabs; the selected one fills paper, takes the rule border and drops 1px to join the sheet.

### Reference Table (signature)
Sticky `paper-2` head with 12px/650 labels and sort arrows; 1px hairline rows; right-aligned numerics; hover tints rows pale blue; the selected row fills `sel`. Agents group under class band rows (class tint, uppercase class caps in the edge colour). Each row carries the %S chip, its interval in `ink-2`, and a 110px interval mark (point dot for %S, diamond for coverage, bar for the CI).

### %S Chip
Min 46px, right-aligned, 600 weight, 2px radius, threshold tint with dark text; untested uses `na-bg` and `ink-3`.

### Note Marker and Notes Column (signature)
Every derived header or figure carries a small superscript numbered button in teal (hover `accent-soft`). The notes column holds a Derivation block (definition list, then a monospace equation on `paper-2`, flashing pale yellow for 220ms on change) above a counter-numbered notes list; the referenced note highlights in `sel`.

### Antibiogram Grid
Rotated class-tinted column heads (150px), italic organism row heads sticky on the left, centred 12px/600 cells in threshold tints; cells with too few tested fade to 55%; selection is a 2px inset ink outline.

### Drop Zone
Dashed `rule-strong` panel on paper; hover or drag-over turns the dash teal and the fill pale blue.

## Do's and Don'ts

### Do:
- **Do** print the count beside every percentage and the %S value inside every threshold cell.
- **Do** attach a numbered note marker to every computed figure and route its derivation to the notes column.
- **Do** group agents under their class band (`--cb` tint, `--ce` edge) wherever agents are listed, and mirror the same hex values in exported figures.
- **Do** separate surfaces with 1px rules and the ground/paper-2/paper tonal steps.
- **Do** keep controls at 30px with the 3px radius, and keep numerics tabular and right-aligned.

### Don't:
- **Don't** add KPI tiles, shadowed cards or a lone-accent dashboard layout; the system is a ruled reference sheet.
- **Don't** put a shadow on any persistent surface; only the Export popover and toast are lifted.
- **Don't** use threshold green, amber or red for chrome, status or decoration; they mean %S bands only.
- **Don't** let colour be the only carrier of a susceptibility result.
