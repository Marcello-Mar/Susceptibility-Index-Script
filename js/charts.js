/* Standalone SVG figures (used on screen and for export). Colours are literal so exported files keep them. */
(function (root) {
  'use strict';

  const C = {
    ink: '#13252d', ink2: '#4a626d', ink3: '#6d838d', rule: '#cfd9de', grid: '#e3e9ec', paper: '#ffffff',
    good: '#236b4c', goodBg: '#d7ebe0', mid: '#7a5700', midBg: '#f4e6bf', low: '#9c2f27', lowBg: '#f4d6d1', naBg: '#eef1f3',
    accent: '#0f5c6e',
  };
  const CLASS_EDGE = {
    'Penicillins': '#2f5f9e', 'Cephalosporins': '#5a46a8', 'Carbapenems': '#a3365f', 'Monobactams': '#86468f',
    'Aminoglycosides': '#2c7652', 'Fluoroquinolones': '#8c5b12', 'Glycopeptides': '#1d6b7c', 'Lipopeptides': '#557a24',
    'Oxazolidinones': '#7d6a14', 'Macrolides & lincosamides': '#9b4a2e', 'Streptogramins': '#8a4b3b', 'Tetracyclines': '#6b5d3b',
    'Folate pathway inhibitors': '#46607e', 'Polymyxins': '#74468c', 'Fosfomycin': '#3c6f6f', 'Nitrofurans': '#75652f',
    'Rifamycins': '#a14d22', 'Fusidanes': '#4d6838', 'Nitroimidazoles': '#5d5070', 'Amphenicols': '#665a46', 'Other': '#5d6f78',
  };
  const FONT = 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif';

  function esc(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function f1(x) { return isFinite(x) ? (x * 100).toFixed(1) : '–'; }
  function band(p, th) {
    if (!isFinite(p)) return 'na';
    if (p * 100 >= th.good) return 'good';
    if (p * 100 >= th.mid) return 'mid';
    return 'low';
  }
  function textWidth(s, size) { return String(s).length * size * 0.56; }

  function header(title, subtitle, w) {
    return `<text x="0" y="16" font-size="15" font-weight="650" fill="${C.ink}">${esc(title)}</text>` +
      (subtitle ? `<text x="0" y="34" font-size="11" fill="${C.ink2}">${esc(subtitle)}</text>` : '') +
      `<line x1="0" x2="${w}" y1="44" y2="44" stroke="${C.ink}" stroke-width="1.5"/>`;
  }

  function wrap(w, h, body) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="-16 -16 ${w + 32} ${h + 32}" font-family="${FONT}">` +
      `<rect x="-16" y="-16" width="${w + 32}" height="${h + 32}" fill="${C.paper}"/>${body}</svg>`;
  }

  /* Forest-style plot: %S among tested (dot, CI) and coverage of all isolates (diamond, CI) */
  function singlesPlot(rows, o) {
    const labelW = Math.max(150, Math.min(260, Math.max(...rows.map((r) => textWidth(r.agent, 12))) + 30));
    const plotW = 420, valW = 190;
    const w = labelW + plotW + valW;
    const top = 92, rowH = 22;
    const h = top + rows.length * rowH + 40;
    const x = (p) => labelW + p * plotW;
    let b = header(o.title, o.subtitle, w);
    // legend
    b += `<g transform="translate(0,64)" font-size="11" fill="${C.ink2}">` +
      `<circle cx="6" cy="-4" r="4.5" fill="${C.ink}"/><text x="16" y="0">% susceptible among tested, ${o.confPct}% CI</text>` +
      `<rect x="${226}" y="-8.5" width="8" height="8" transform="rotate(45 ${230} -4.5)" fill="#fff" stroke="${C.accent}" stroke-width="1.5"/>` +
      `<text x="${242}" y="0">coverage of all isolates (untested = not covered)</text></g>`;
    // grid
    for (let t = 0; t <= 100; t += 10) {
      const xx = x(t / 100);
      b += `<line x1="${xx}" x2="${xx}" y1="${top - 8}" y2="${top + rows.length * rowH}" stroke="${t % 50 === 0 ? C.rule : C.grid}"/>`;
      b += `<text x="${xx}" y="${top + rows.length * rowH + 16}" font-size="10" fill="${C.ink3}" text-anchor="middle">${t}</text>`;
    }
    if (o.thresholds) {
      for (const t of [o.thresholds.mid, o.thresholds.good]) {
        const xx = x(t / 100);
        b += `<line x1="${xx}" x2="${xx}" y1="${top - 8}" y2="${top + rows.length * rowH}" stroke="${C.ink3}" stroke-dasharray="2 3"/>`;
      }
    }
    b += `<text x="${x(0.5)}" y="${top + rows.length * rowH + 32}" font-size="11" fill="${C.ink2}" text-anchor="middle">Percent</text>`;
    b += `<text x="${labelW + plotW + 12}" y="${top - 12}" font-size="10" font-weight="650" fill="${C.ink2}">%S (CI) · S/tested</text>`;
    rows.forEach((r, i) => {
      const y = top + i * rowH + rowH / 2;
      if (i % 2 === 1) b += `<rect x="0" y="${y - rowH / 2}" width="${w}" height="${rowH}" fill="#f6f8f9"/>`;
      const ce = CLASS_EDGE[r.cls] || CLASS_EDGE.Other;
      b += `<rect x="0" y="${y - 5}" width="8" height="10" rx="1.5" fill="${ce}"/>`;
      b += `<text x="14" y="${y + 4}" font-size="12" fill="${C.ink}">${esc(r.agent)}</text>`;
      if (isFinite(r.glo)) {
        b += `<line x1="${x(r.glo)}" x2="${x(r.ghi)}" y1="${y + 4}" y2="${y + 4}" stroke="${C.accent}" stroke-opacity="0.5" stroke-width="1.5"/>`;
        b += `<rect x="${x(r.pG) - 4}" y="${y}" width="8" height="8" transform="rotate(45 ${x(r.pG)} ${y + 4})" fill="#fff" stroke="${C.accent}" stroke-width="1.5"/>`;
      }
      if (isFinite(r.pS)) {
        b += `<line x1="${x(r.lo)}" x2="${x(r.hi)}" y1="${y - 2}" y2="${y - 2}" stroke="${C.ink}" stroke-width="2"/>`;
        b += `<circle cx="${x(r.pS)}" cy="${y - 2}" r="4.5" fill="${C.ink}"/>`;
      }
      b += `<text x="${labelW + plotW + 12}" y="${y + 4}" font-size="11" fill="${C.ink}">${f1(r.pS)} <tspan fill="${C.ink2}">(${f1(r.lo)}–${f1(r.hi)})</tspan> <tspan fill="${C.ink3}">${r.s}/${r.tested}</tspan></text>`;
    });
    if (o.footnote) b += `<text x="0" y="${h - 2}" font-size="10" fill="${C.ink3}">${esc(o.footnote)}</text>`;
    return wrap(w, h, b);
  }

  /* Combination matrix: lower triangle = combination %S, diagonal = single agent %S */
  function comboMatrix(agents, singles, combos, o) {
    const n = agents.length;
    const cell = n > 22 ? 26 : 32;
    const labelW = Math.max(130, Math.min(240, Math.max(...agents.map((a) => textWidth(a.agent, 11))) + 24));
    const top = 64 + labelW * 0.72;
    const lastLabel = agents.length ? Math.max(...agents.map((ag, j) => String(ag.agent).length * 11 * 0.62 * Math.cos(50 * Math.PI / 180) - (n - 1 - j) * cell)) + 4 : 0;
    const w = labelW + n * cell + Math.max(10, lastLabel - cell / 2 + 8);
    const h = top + n * cell + 48;
    const sMap = new Map(singles.map((s) => [s.agent, s]));
    const cMap = new Map();
    for (const c of combos) { cMap.set(c.a + '\u0000' + c.b, c); cMap.set(c.b + '\u0000' + c.a, c); }
    let b = header(o.title, o.subtitle, w);
    // column labels, rotated
    agents.forEach((a, j) => {
      const cx = labelW + j * cell + cell / 2;
      b += `<text transform="translate(${cx + 4},${top - 6}) rotate(-50)" font-size="11" fill="${C.ink}">${esc(a.agent)}</text>`;
    });
    agents.forEach((a, i) => {
      const y = top + i * cell;
      const ce = CLASS_EDGE[a.cls] || CLASS_EDGE.Other;
      b += `<rect x="0" y="${y + cell / 2 - 5}" width="8" height="10" rx="1.5" fill="${ce}"/>`;
      b += `<text x="14" y="${y + cell / 2 + 4}" font-size="11" fill="${C.ink}">${esc(a.agent)}</text>`;
      for (let j = 0; j <= i; j++) {
        const x = labelW + j * cell;
        let p, title;
        if (i === j) { const s = sMap.get(a.agent); p = s ? s.pS : NaN; title = `${a.agent}: ${f1(p)}%`; }
        else { const c = cMap.get(a.agent + '\u0000' + agents[j].agent); p = c ? c.pS : NaN; title = `${a.agent} + ${agents[j].agent}: ${f1(p)}%`; }
        const bd = band(p, o.thresholds);
        const fill = { good: C.goodBg, mid: C.midBg, low: C.lowBg, na: C.naBg }[bd];
        const fg = { good: C.good, mid: C.mid, low: C.low, na: C.ink3 }[bd];
        const key = i === j ? '' : ` data-a="${esc(a.agent)}" data-b="${esc(agents[j].agent)}"`;
        b += `<g class="mx"${key}><title>${esc(title)}</title><rect x="${x + 0.5}" y="${y + 0.5}" width="${cell - 1}" height="${cell - 1}" fill="${fill}" ${i === j ? `stroke="${C.ink}" stroke-width="1"` : ''}/>` +
          `<text x="${x + cell / 2}" y="${y + cell / 2 + 3.5}" font-size="${cell > 28 ? 10 : 9}" font-weight="600" fill="${fg}" text-anchor="middle">${isFinite(p) ? Math.floor(p * 100 + 1e-9) : '–'}</text></g>`;
      }
    });
    const ly = top + n * cell + 22;
    b += `<g font-size="10" fill="${C.ink2}" transform="translate(0,${ly})">` +
      `<rect x="0" y="-9" width="12" height="12" fill="${C.goodBg}"/><text x="16" y="1">≥ ${o.thresholds.good}%</text>` +
      `<rect x="70" y="-9" width="12" height="12" fill="${C.midBg}"/><text x="86" y="1">${o.thresholds.mid}–&lt;${o.thresholds.good}%</text>` +
      `<rect x="170" y="-9" width="12" height="12" fill="${C.lowBg}"/><text x="186" y="1">&lt; ${o.thresholds.mid}%</text>` +
      `<rect x="240" y="-9" width="12" height="12" fill="#fff" stroke="${C.ink}"/><text x="256" y="1">diagonal: single agent · values truncated</text></g>`;
    if (o.footnote) b += `<text x="0" y="${h - 2}" font-size="10" fill="${C.ink3}">${esc(o.footnote)}</text>`;
    return wrap(w, h, b);
  }

  /* Organism x agent antibiogram */
  function antibiogramPlot(grid, agents, o) {
    const cellW = 36, cellH = 22;
    const labelW = Math.max(170, Math.min(280, Math.max(...grid.map((g) => textWidth(g.name, 12))) + 50));
    const top = 64 + Math.min(170, Math.max(...agents.map((a) => textWidth(a.agent, 11))) * 0.8 + 10);
    const w = labelW + agents.length * cellW;
    const h = top + grid.length * cellH + 48;
    let b = header(o.title, o.subtitle, w);
    agents.forEach((a, j) => {
      const cx = labelW + j * cellW + cellW / 2;
      const ce = CLASS_EDGE[a.cls] || CLASS_EDGE.Other;
      b += `<rect x="${labelW + j * cellW + 1}" y="${top - 5}" width="${cellW - 2}" height="3" fill="${ce}"/>`;
      b += `<text transform="translate(${cx + 4},${top - 10}) rotate(-60)" font-size="11" fill="${C.ink}">${esc(a.agent)}</text>`;
    });
    grid.forEach((g, i) => {
      const y = top + i * cellH;
      b += `<text x="0" y="${y + cellH / 2 + 4}" font-size="12" font-style="italic" fill="${C.ink}">${esc(g.name)}</text>`;
      b += `<text x="${labelW - 8}" y="${y + cellH / 2 + 4}" font-size="10" fill="${C.ink3}" text-anchor="end">${g.n}</text>`;
      g.cells.forEach((c, j) => {
        const x = labelW + j * cellW;
        const few = c.tested > 0 && c.tested < o.minTested;
        const bd = c.tested ? band(c.pS, o.thresholds) : 'na';
        const fill = { good: C.goodBg, mid: C.midBg, low: C.lowBg, na: '#ffffff' }[bd];
        const fg = { good: C.good, mid: C.mid, low: C.low, na: '#b5c2c8' }[bd];
        b += `<rect x="${x + 0.5}" y="${y + 0.5}" width="${cellW - 1}" height="${cellH - 1}" fill="${fill}" ${few ? 'fill-opacity="0.5"' : ''}/>`;
        b += `<text x="${x + cellW / 2}" y="${y + cellH / 2 + 3.5}" font-size="10" font-weight="600" fill="${fg}" text-anchor="middle" ${few ? 'fill-opacity="0.6"' : ''}>${c.tested ? Math.floor(c.pS * 100 + 1e-9) + (few ? '*' : '') : '–'}</text>`;
      });
      b += `<line x1="0" x2="${w}" y1="${y + cellH}" y2="${y + cellH}" stroke="${C.grid}"/>`;
    });
    b += `<text x="0" y="${top + grid.length * cellH + 20}" font-size="10" fill="${C.ink2}">Cells: % susceptible among tested, truncated. n = isolates. * fewer than ${o.minTested} tested. – not tested.</text>`;
    if (o.footnote) b += `<text x="0" y="${h - 2}" font-size="10" fill="${C.ink3}">${esc(o.footnote)}</text>`;
    return wrap(w, h, b);
  }

  root.SICharts = { singlesPlot, comboMatrix, antibiogramPlot, band, CLASS_EDGE };
})(typeof window !== 'undefined' ? window : globalThis);
