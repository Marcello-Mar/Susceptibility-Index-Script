/* Susceptibility Index v2 — UI */
(function () {
  'use strict';
  const E = window.SIEngine, CH = window.SICharts;

  /* ---------------- state ---------------- */
  const S = {
    file: null,          // { name, synthetic, sheets, sheet, wb, rows, columns }
    map: null,           // { orgCol, patientCol, dateCol, agents: [{ name, include, cls, tested }] }
    params: { iMode: 'S', dedup: 'patient-organism', threshold: 0, conf: 0.95, combDen: 'either', good: 90, mid: 70, abgBy: 'species', abgMin: 10, abgMinTested: 10 },
    gramMap: (window.SI_DEFAULT_GRAM_MAP || []).map((e) => e.slice()),
    overrides: [],       // [[exact organism, category]] set in this session
    exclusions: '',
    subset: { type: 'all' },
    view: 'overview',
    singlesMode: 'table', combosMode: 'table', groupByClass: true,
    sort: { singles: { key: 'pS', dir: -1 }, combos: { key: 'pS', dir: -1 } },
    classFilter: '', query: '', comboAgent: '', comboSigOnly: false, comboShowAll: false, orgQuery: '',
    sel: null, hlNote: null,
    prep: null, cache: new Map(), pendingSettings: null,
  };

  /* ---------------- helpers ---------------- */
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pct = (x, d = 1) => (isFinite(x) ? (x * 100).toFixed(d) : '–');
  const pval = (p) => (!isFinite(p) ? '–' : p < 0.001 ? '<0.001' : p.toFixed(3));
  const pEq = (p) => (isFinite(p) && p < 0.001 ? '< 0.001' : '= ' + pval(p));
  const fmtInt = (n) => (n == null ? '–' : Number(n).toLocaleString('en-US'));
  const confPct = () => Math.round(S.params.conf * 1000) / 10;
  const th = () => ({ good: S.params.good, mid: S.params.mid });
  const bandOf = (p) => CH.band(p, th());
  const MIN_LIST = 10;
  const SHORT_CLASS = { 'Penicillins': 'PEN', 'Cephalosporins': 'CEPH', 'Carbapenems': 'CARB', 'Monobactams': 'MONO', 'Aminoglycosides': 'AG', 'Fluoroquinolones': 'FQ',
    'Glycopeptides': 'GLY', 'Lipopeptides': 'LIP', 'Oxazolidinones': 'OXA', 'Macrolides & lincosamides': 'MLS', 'Streptogramins': 'STG', 'Tetracyclines': 'TET',
    'Folate pathway inhibitors': 'FOL', 'Polymyxins': 'POL', 'Fosfomycin': 'FOS', 'Nitrofurans': 'NIT', 'Rifamycins': 'RIF', 'Fusidanes': 'FUS', 'Nitroimidazoles': 'NIM', 'Amphenicols': 'AMP', 'Other': 'OTH' };
  const CAT_LABEL = { positive: 'Gram-positive', negative: 'Gram-negative', fungus: 'Fungi', other: 'Other', unclassified: 'Unclassified' };
  const CAT_COLOR = { positive: 'var(--gpos)', negative: 'var(--gneg)', fungus: 'var(--gfun)', other: 'var(--gunc)', unclassified: '#c99a2e' };
  const ICON = {
    up: '<svg class="i" viewBox="0 0 24 24"><path d="M4 15v4h16v-4M12 4v11M7.5 9.5 12 4l4.5 5.5"/></svg>',
    table: '<svg class="i" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 10h18M3 15h18M9 4v16"/></svg>',
    chart: '<svg class="i" viewBox="0 0 24 24"><path d="M4 6h3M4 12h9M4 18h6"/><circle cx="16" cy="6" r="2"/><circle cx="18" cy="12" r="2"/><circle cx="13" cy="18" r="2"/></svg>',
    grid: '<svg class="i" viewBox="0 0 24 24"><path d="M4 4h4v4H4zM4 10h4v4H4zM10 10h4v4h-4zM4 16h4v4H4zM10 16h4v4h-4zM16 16h4v4h-4z"/></svg>',
    x: '<svg class="i" viewBox="0 0 24 24"><path d="m6 6 12 12M18 6 6 18"/></svg>',
  };

  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 2600);
  }
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  }
  function baseName() { return S.file ? S.file.name.replace(/\.[^.]+$/, '') : 'susceptibility'; }
  function subsetLabel(sub = S.subset) {
    if (sub.type === 'all') return 'All isolates';
    if (sub.type === 'gram') return sub.value === 'positive' ? 'Gram-positive' : 'Gram-negative';
    return sub.value;
  }
  function subsetKey(sub) { return sub.type + ':' + (sub.value || ''); }

  /* ---------------- file loading ---------------- */
  function looksLikeAgentColumn(rows, col) {
    let valid = 0, nonEmpty = 0;
    const lim = Math.min(rows.length, 400);
    for (let i = 0; i < lim; i++) {
      const v = rows[i][col];
      if (v == null || String(v).trim() === '') continue;
      nonEmpty++;
      if (E.normalizeValue(v, 'S') !== -1) valid++;
    }
    return nonEmpty > 0 && valid / nonEmpty >= 0.6;
  }
  function guessColumns(columns, rows) {
    const pick = (re, not = []) => columns.find((c) => re.test(c) && !not.includes(c)) || null;
    const orgCol = pick(/micro|organ|organism|germe|isolat|species|specie|batter|pathogen|patogen/i) || columns[0];
    const patientCol = pick(/patient|pazient|^pid$|mrn|nosolog|record|cartella|codice/i, [orgCol]);
    const dateCol = pick(/date|data|prelievo|collect|sample.?time|giorno/i, [orgCol, patientCol]);
    const used = new Set([orgCol, patientCol, dateCol]);
    const agents = columns.filter((c) => !used.has(c) && looksLikeAgentColumn(rows, c)).map((c) => ({ name: c, include: true, cls: E.guessClass(c), tested: 0 }));
    return { orgCol, patientCol, dateCol, agents };
  }
  function readSheet(wb, name) {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    const header = (XLSX.utils.sheet_to_json(ws, { header: 1, range: 0, blankrows: false })[0] || []).map((h) => (h == null ? '' : String(h)));
    const columns = header.filter((h) => h !== '');
    return { rows, columns };
  }
  function loadWorkbook(wb, name, synthetic) {
    const sheets = wb.SheetNames;
    const sheet = sheets[0];
    const { rows, columns } = readSheet(wb, sheet);
    S.file = { name, synthetic: !!synthetic, wb, sheets, sheet, rows, columns };
    afterRowsLoaded();
  }
  function afterRowsLoaded() {
    const f = S.file;
    if (!f.rows.length || !f.columns.length) { toast('No rows found in this sheet.'); }
    S.map = guessColumns(f.columns, f.rows);
    S.subset = { type: 'all' }; S.sel = null; S.view = 'overview';
    if (S.pendingSettings) { applySettings(S.pendingSettings, true); S.pendingSettings = null; }
    recompute(true);
    renderRail(); renderMain();
  }
  function handleFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true, dense: false });
        loadWorkbook(wb, file.name, false);
        toast(`Loaded ${file.name}`);
      } catch (err) {
        toast('Could not read this file: ' + (err && err.message ? err.message : err));
      }
    };
    reader.onerror = () => toast('The browser could not read this file.');
    reader.readAsArrayBuffer(file);
  }
  function loadDemo() {
    const d = window.SI_DEMO;
    const aoa = [d.columns].concat(d.rows.map((r) => r.map((v) => (v === '' ? null : v))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Isolates');
    loadWorkbook(wb, d.name, true);
  }

  /* ---------------- computation ---------------- */
  function cfg() {
    const m = S.map;
    return {
      orgCol: m.orgCol, patientCol: m.patientCol || null, dateCol: m.dateCol || null,
      agentCols: m.agents.filter((a) => a.include).map((a) => a.name),
      iMode: S.params.iMode, dedup: m.patientCol ? S.params.dedup : 'none',
      exclusions: S.exclusions.split(/\n|;/).map((x) => x.trim()).filter(Boolean),
      gramMap: S.overrides.concat(S.gramMap),
    };
  }
  function agentClass(name) {
    const a = S.map && S.map.agents.find((x) => x.name === name);
    return a ? a.cls : E.guessClass(name);
  }
  function recompute(prepAgain) {
    if (!S.file || !S.map) return;
    if (prepAgain) {
      // tested fraction per column on raw rows, for the column list
      for (const a of S.map.agents) {
        let t = 0;
        for (const r of S.file.rows) if (E.normalizeValue(r[a.name], S.params.iMode) > 0) t++;
        a.tested = S.file.rows.length ? t / S.file.rows.length : 0;
      }
      S.prep = E.prepare(S.file.rows, cfg());
      S.cache.clear();
      // keep a genus/species subset only if it still exists
      if (S.subset.type === 'genus' || S.subset.type === 'species') {
        const ok = S.prep.isolates.some((i) => i[S.subset.type] === S.subset.value);
        if (!ok) S.subset = { type: 'all' };
      }
    }
  }
  function resultsFor(sub) {
    const key = subsetKey(sub) + '|' + S.params.threshold + '|' + S.params.conf + '|' + S.params.combDen;
    if (S.cache.has(key)) return S.cache.get(key);
    const iso = E.subsetIsolates(S.prep.isolates, sub);
    const r = E.analyze(iso, S.prep.agents, { threshold: S.params.threshold, conf: S.params.conf, combDenominator: S.params.combDen });
    r.isolates = iso;
    for (const s of r.singles) s.cls = agentClass(s.agent);
    S.cache.set(key, r);
    return r;
  }
  let pending = 0;
  function schedule(prepAgain, railToo) {
    cancelAnimationFrame(pending);
    pending = requestAnimationFrame(() => {
      recompute(prepAgain);
      if (railToo) renderRailDynamic();
      renderMain();
    });
  }

  /* ---------------- rail ---------------- */
  function renderRail() {
    const rail = $('#rail');
    const has = !!S.file;
    rail.innerHTML = `
      <section class="step" id="step-data"><h2><span class="n">1</span>Data<span class="meta" id="meta-data"></span></h2><div class="body" id="sec-data"></div></section>
      <section class="step" id="step-cols" aria-disabled="${!has}"><h2><span class="n">2</span>Columns<span class="meta" id="meta-cols"></span></h2><div class="body" id="sec-cols"></div></section>
      <section class="step" id="step-org" aria-disabled="${!has}"><h2><span class="n">3</span>Organisms<span class="meta" id="meta-org"></span></h2><div class="body"><div id="sec-org"></div>
        <div class="field" style="margin-top:14px"><label for="in-excl">Exclude organisms containing</label>
        <textarea id="in-excl" rows="3" placeholder="One per line, e.g. contaminant&#10;Micrococcus">${esc(S.exclusions)}</textarea>
        <div class="row"><span class="hint">Case-insensitive text match on the organism name.</span><button class="btn ghost small" data-action="excl-file" type="button">From file</button></div>
        <input type="file" id="file-excl" accept=".xlsx,.xls,.csv,.txt" hidden></div></div></section>
      <section class="step" id="step-par" aria-disabled="${!has}"><h2><span class="n">4</span>Parameters</h2><div class="body" id="sec-par"></div></section>`;
    renderData(); renderCols(); renderOrg(); renderParams();
  }
  function renderRailDynamic() { renderData(); renderOrg(); renderColsMeta(); }

  function renderData() {
    const el = $('#sec-data');
    const f = S.file;
    let h = `<label class="drop" id="drop" tabindex="0">
        ${ICON.up}<span><strong>${f ? 'Replace file' : 'Choose a file'}</strong> or drop it here</span>
        <small>.xlsx, .xls or .csv · one row per isolate</small>
        <input type="file" id="file-data" accept=".xlsx,.xls,.xlsm,.csv,.tsv,.ods" hidden></label>`;
    if (f) {
      h = `<div class="file-line"><b title="${esc(f.name)}">${esc(f.name)}</b>${f.synthetic ? '<span class="synthetic-tag" style="margin-left:0">Synthetic</span>' : ''}<span>${fmtInt(f.rows.length)} rows</span></div>` +
        (f.sheets.length > 1 ? `<div class="field"><label for="in-sheet">Sheet</label><select id="in-sheet">${f.sheets.map((s) => `<option ${s === f.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div>` : '') +
        `<div style="margin-top:10px">${h}</div>`;
    } else {
      h += `<div class="row" style="margin-top:10px"><button class="btn" data-action="demo" type="button">Load synthetic demo</button><span class="hint" style="font-size:var(--t-xs)">640 invented isolates</span></div>`;
    }
    el.innerHTML = h;
    $('#meta-data').textContent = f ? `${fmtInt(S.prep ? S.prep.stats.analysed : 0)} analysed` : '';
  }

  function colOptions(selected, allowNone) {
    const cols = S.file ? S.file.columns : [];
    return (allowNone ? `<option value="">None</option>` : '') + cols.map((c) => `<option ${c === selected ? 'selected' : ''} value="${esc(c)}">${esc(c)}</option>`).join('');
  }
  function renderCols() {
    const el = $('#sec-cols');
    if (!S.file) { el.innerHTML = '<p class="hint" style="margin:0">Load a file to map its columns.</p>'; return; }
    const m = S.map;
    const classes = E.DISPLAY_CLASS_ORDER;
    el.innerHTML = `
      <div class="field"><label for="in-org">Organism</label><select id="in-org">${colOptions(m.orgCol)}</select></div>
      <div class="row" style="margin-top:12px; align-items:start">
        <div class="field" style="margin:0; flex:1"><label for="in-pid">Patient ID</label><select id="in-pid">${colOptions(m.patientCol, true)}</select></div>
        <div class="field" style="margin:0; flex:1"><label for="in-date">Date</label><select id="in-date">${colOptions(m.dateCol, true)}</select></div>
      </div>
      <div class="field"><div class="row"><span class="lbl" style="flex:1">Antibiotics <span class="muted" id="ag-count"></span></span>
        <button class="btn ghost small" data-action="ag-all" type="button">All</button><button class="btn ghost small" data-action="ag-none" type="button">None</button></div>
      <div class="agents"><table><tbody>${m.agents.map((a, i) => `
        <tr data-class="${esc(a.cls)}"><td><input type="checkbox" data-ag="${i}" ${a.include ? 'checked' : ''} aria-label="Include ${esc(a.name)}"></td>
        <td class="nm" title="${esc(a.name)}">${esc(a.name)}</td>
        <td><select data-agcls="${i}" aria-label="Class of ${esc(a.name)}" title="${esc(a.cls)}">${classes.map((c) => `<option value="${esc(c)}" ${c === a.cls ? 'selected' : ''} label="${esc(SHORT_CLASS[c] || c)}">${esc(c)}</option>`).join('')}</select></td>
        <td class="pc">${Math.round(a.tested * 100)}%</td></tr>`).join('')}</tbody></table></div>
      <span class="hint">Detected from columns whose values are S, I, R or blank. Percent = rows with a result.</span>
      </div>`;
    renderColsMeta();
  }
  function renderColsMeta() {
    if (!S.map) return;
    const n = S.map.agents.filter((a) => a.include).length;
    $('#meta-cols').textContent = `${n} agents`;
    const c = $('#ag-count'); if (c) c.textContent = `${n} of ${S.map.agents.length}`;
  }

  function renderOrg() {
    const el = $('#sec-org');
    if (!S.prep) { el.innerHTML = '<p class="hint" style="margin:0">Classification uses the bundled genus map; unknown organisms are listed here for you to assign.</p>'; $('#meta-org').textContent = ''; return; }
    const sum = E.summary(S.prep.isolates);
    const unc = S.prep.stats.unclassified;
    el.innerHTML = `<div class="cats">${sum.categories.map(([c, n]) => `<span class="sw" style="background:${CAT_COLOR[c]}"></span><span>${CAT_LABEL[c] || c}</span><span class="ct">${fmtInt(n)}</span>`).join('')}</div>` +
      (unc.length ? `<div class="unc"><p>${unc.length} organism name${unc.length > 1 ? 's' : ''} not in the map</p><ul>${unc.slice(0, 40).map(([o, n]) => `
        <li><span title="${esc(o)}">${esc(o)}</span><em>${n}</em><select data-assign="${esc(o)}" aria-label="Classify ${esc(o)}"><option value="">Assign…</option><option value="positive">Gram+</option><option value="negative">Gram−</option><option value="fungus">Fungus</option><option value="other">Other</option></select></li>`).join('')}</ul>
        ${unc.length > 40 ? `<p class="muted" style="margin:6px 0 0;font-weight:400">${unc.length - 40} more in Organisms view.</p>` : ''}</div>` : '') +
      `<div class="row" style="margin-top:10px"><button class="btn small" data-action="view-org" type="button">Review all organisms</button></div>`;
    $('#meta-org').textContent = unc.length ? `${unc.length} to assign` : `${sum.organisms.length} names`;
  }

  function renderParams() {
    const el = $('#sec-par');
    const p = S.params;
    const hasPid = S.map && S.map.patientCol;
    const seg = (name, opts, val) => `<div class="seg" role="group">${opts.map(([v, l]) => `<button type="button" data-seg="${name}" data-val="${v}" aria-pressed="${String(v) === String(val)}">${l}</button>`).join('')}</div>`;
    el.innerHTML = `
      <div class="field"><span class="lbl">Intermediate (I) results</span>${seg('iMode', [['S', 'Count as S'], ['R', 'As R'], ['N', 'Not tested']], p.iMode)}</div>
      <div class="field"><label for="in-dedup">Duplicate isolates</label>
        <select id="in-dedup" ${hasPid ? '' : 'disabled'}><option value="none" ${p.dedup === 'none' ? 'selected' : ''}>Keep all isolates</option><option value="patient-organism" ${p.dedup === 'patient-organism' ? 'selected' : ''}>First isolate per patient and organism</option><option value="patient" ${p.dedup === 'patient' ? 'selected' : ''}>First isolate per patient</option></select>
        <span class="hint">${hasPid ? 'Earliest by date when a date column is set, otherwise file order.' : 'Needs a patient ID column (step 2).'}</span></div>
      <div class="field"><label for="in-thr">Minimum tested, % of isolates in the subset</label>
        <div class="range-row"><input type="range" id="in-thr" min="0" max="100" step="5" value="${p.threshold}"><input type="number" id="in-thr-n" min="0" max="100" step="1" value="${p.threshold}" aria-label="Minimum tested percent"></div>
        <span class="hint">Agents tested in fewer isolates are left out of the subset.</span></div>
      <div class="field"><label for="in-conf">Confidence level</label>
        <select id="in-conf">${[0.9, 0.95, 0.99].map((c) => `<option value="${c}" ${c === p.conf ? 'selected' : ''}>${c * 100}%</option>`).join('')}</select></div>
      <div class="field"><span class="lbl">Combination denominator</span>${seg('combDen', [['either', 'Tested for ≥1 agent'], ['both', 'Tested for both']], p.combDen)}
        <span class="hint">“≥1 agent” reproduces the v1 script.</span></div>
      <div class="field"><span class="lbl">Colour bands for %S</span>
        <div class="row"><label class="hint" for="in-good">Green from</label><input type="number" id="in-good" min="1" max="100" value="${p.good}" style="width:64px"><label class="hint" for="in-mid">amber from</label><input type="number" id="in-mid" min="0" max="99" value="${p.mid}" style="width:64px"></div></div>`;
  }

  /* ---------------- main ---------------- */
  function renderMain() {
    const main = $('#main');
    $('#btn-export').disabled = !S.prep;
    if (!S.prep) { main.innerHTML = emptyState(); return; }
    const all = S.prep.isolates;
    const nPos = all.filter((i) => i.category === 'positive').length;
    const nNeg = all.filter((i) => i.category === 'negative').length;
    const genera = countBy(all, 'genus'), species = countBy(all, 'species');
    const pickOn = S.subset.type === 'genus' || S.subset.type === 'species';
    const tab = (sub, label, n, dot) => `<button role="tab" data-sub="${sub.type}" data-subval="${esc(sub.value || '')}" aria-selected="${subsetKey(sub) === subsetKey(S.subset)}">${dot ? `<span class="gdot" style="background:${dot}"></span>` : ''}<b>${label}</b><small>${fmtInt(n)}</small></button>`;
    const views = [['overview', 'Overview'], ['singles', 'Single agents'], ['combos', 'Combinations'], ['abg', 'Antibiogram'], ['org', 'Organisms']];
    main.innerHTML = `
      <nav class="subsets" role="tablist" aria-label="Isolate subset">
        ${tab({ type: 'all' }, 'All isolates', all.length)}
        ${tab({ type: 'gram', value: 'positive' }, 'Gram-positive', nPos, 'var(--gpos)')}
        ${tab({ type: 'gram', value: 'negative' }, 'Gram-negative', nNeg, 'var(--gneg)')}
        <div class="pick ${pickOn ? 'on' : ''}"><label class="sr-only" for="in-pick">Genus or species</label>
          <select id="in-pick"><option value="">Genus or species…</option>
            <optgroup label="Genus">${genera.map(([g, n]) => `<option value="genus|${esc(g)}" ${S.subset.type === 'genus' && S.subset.value === g ? 'selected' : ''}>${esc(g)} (${n})</option>`).join('')}</optgroup>
            <optgroup label="Species">${species.map(([g, n]) => `<option value="species|${esc(g)}" ${S.subset.type === 'species' && S.subset.value === g ? 'selected' : ''}>${esc(g)} (${n})</option>`).join('')}</optgroup>
          </select></div>
      </nav>
      <div class="views" role="tablist" aria-label="View">${views.map(([v, l]) => `<button role="tab" data-view="${v}" aria-selected="${S.view === v}">${l}</button>`).join('')}
        <div class="tools" id="view-tools"></div></div>
      <div class="work"><section class="sheet" id="sheet"></section><aside class="notes" id="notes" aria-label="Notes"></aside></div>
      <footer class="colophon">Susceptibility Index v2 · ${S.file.synthetic ? 'Showing synthetic demo data, not patient data. · ' : ''}Calculations run in this browser tab. GPL-3.0.</footer>`;
    const r = resultsFor(S.subset);
    ({ overview: viewOverview, singles: viewSingles, combos: viewCombos, abg: viewAbg, org: viewOrg })[S.view](r);
    renderNotes(r);
  }
  function countBy(list, key) {
    const m = new Map();
    for (const i of list) m.set(i[key], (m.get(i[key]) || 0) + 1);
    return Array.from(m.entries()).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(b[0]));
  }

  function emptyState() {
    return `<div class="work" style="padding-top:20px"><section class="sheet" style="border-radius:var(--r)"><div class="empty-state">
      <div><h1>Susceptibility and coverage from an isolate line-list</h1>
        <p>Load an Excel or CSV export with one row per isolate and one column per antibiotic. The page estimates the percentage of susceptible isolates for each agent and every two-drug combination, with confidence intervals, overall and by Gram group, genus or species.</p>
        <p>The file is read by this page and stays on your computer. Close the tab and the session is gone; export what you need first.</p>
        <div class="spec-wrap"><table class="spec" aria-label="Expected layout"><thead><tr><th>Microorganism</th><th>Patient ID</th><th>Ampicillin</th><th>Gentamicin</th><th>Ciprofloxacin</th><th>…</th></tr></thead>
        <tbody><tr><td><i>Escherichia coli</i></td><td>P0017</td><td class="r">R</td><td class="s">S</td><td class="s">S</td><td></td></tr>
        <tr><td><i>Staphylococcus aureus</i></td><td>P0021</td><td class="r">R</td><td class="s">S</td><td class="n">–</td><td></td></tr></tbody></table></div>
        <p class="hint" style="font-size:var(--t-xs)">S, I, R; empty, N, NT or “–” mean not tested. Patient ID and date are optional and only used to remove repeat isolates.</p>
        <div class="row" style="margin-top:16px"><button class="btn primary" data-action="pick-file" type="button">${ICON.up}Choose a file</button><button class="btn" data-action="demo" type="button">Load synthetic demo</button></div>
      </div>
      <div><h2 style="margin:0 0 10px;font-size:var(--t-md)">What changed from the Colab script</h2>
        <ol>
          <li><b>Columns are mapped on screen.</b> The organism column no longer has to come first; patient ID and date columns are optional.</li>
          <li><b>Combination vs single agent is a paired test.</b> Exact McNemar on the same isolates replaces the chi-square for independent samples.</li>
          <li><b>Organism classification matches on the longest name at a word start</b>, so Actinobacillus is never read as Bacillus. Unknown names are assigned here instead of in a prompt.</li>
          <li><b>New options:</b> I as S, R or not tested; first isolate per patient; genus and species subsets; denominator for combinations.</li>
          <li><b>Exports:</b> Excel workbook, CSV, SVG and PNG figures, printable report, reusable settings and organism map.</li>
        </ol>
        <p style="margin-top:14px;font-size:var(--t-sm)">With repeat isolates kept, I counted as S and the “≥1 agent” denominator, percentages and intervals equal those of the v1 script (checked to 10<sup>−10</sup> on the demo data).</p>
      </div></div></section></div>`;
  }

  /* ---- overview ---- */
  function viewOverview(r) {
    const st = S.prep.stats;
    const all = S.prep.isolates;
    const sum = E.summary(r.isolates);
    const top = sum.organisms.slice(0, 10);
    const maxTop = top.length ? top[0][1] : 1;
    const least = (cat) => {
      const rr = resultsFor({ type: 'gram', value: cat });
      return rr.singles.filter((s) => s.tested >= MIN_LIST).slice().sort((a, b) => a.pS - b.pS).slice(0, 10);
    };
    const leastTable = (rows) => rows.length ? `<table class="flow"><colgroup><col><col class="c-pct"><col class="c-v"></colgroup><thead class="sr-only"><tr><th>Agent</th><th>%S</th><th>S/tested</th></tr></thead>${rows.map((s) => `<tr data-class="${esc(s.cls)}"><td class="nm" title="${esc(s.agent)}"><span class="clsdot"></span>${esc(s.agent)}</td><td class="num"><span class="pct ${bandOf(s.pS)}">${pct(s.pS)}</span></td><td class="nn">${s.s}/${s.tested}</td></tr>`).join('')}</table>` : '<p class="muted" style="margin:0">No isolates in this group.</p>';
    const nr = (n) => `<button class="nref" data-note="${n}" type="button" aria-label="Note ${n}">${n}</button>`;
    $('#sheet').innerHTML = `
      <div class="sheet-head"><h1>${esc(subsetLabel())}</h1><span class="sub">${fmtInt(r.n)} isolates · ${r.kept.length} agents analysed · ${r.combos.length} combinations</span></div>
      ${warnings(r)}
      <div class="ledger">
        <section><h3>Isolates</h3><table class="flow">
          <tr><td>Rows read</td><td class="num">${fmtInt(st.rowsRead)}</td></tr>
          ${st.emptyOrganism ? `<tr class="minus"><td>− no organism name</td><td class="num">${fmtInt(st.emptyOrganism)}</td></tr>` : ''}
          <tr class="minus"><td>− excluded by name</td><td class="num">${fmtInt(st.excluded)}</td></tr>
          <tr class="minus"><td>− repeat isolates${nr(7)}</td><td class="num">${S.map.patientCol && S.params.dedup !== 'none' ? fmtInt(st.dedupRemoved) : '<span class="muted" style="font-weight:400">not applied</span>'}</td></tr>
          <tr class="total"><td>Analysed</td><td class="num">${fmtInt(all.length)}</td></tr>
          ${S.subset.type !== 'all' ? `<tr><td>In ${esc(subsetLabel())}</td><td class="num">${fmtInt(r.n)}</td></tr>` : ''}
        </table></section>
        <section><h3>Groups${nr(8)}</h3><div class="bars">${E.summary(all).categories.map(([c, n]) => `<div class="b"><span title="${CAT_LABEL[c] || c}"><span class="gdot" style="display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;background:${CAT_COLOR[c]}"></span>${CAT_LABEL[c] || c}</span><i style="width:${(n / all.length) * 100}%;background:${CAT_COLOR[c]}"></i><em>${fmtInt(n)} · ${pct(n / all.length)}%</em></div>`).join('')}</div></section>
        <section><h3>Most frequent organisms${S.subset.type !== 'all' ? ' in subset' : ''}</h3><div class="bars">${top.map(([o, n]) => `<div class="b it"><span title="${esc(o)}">${esc(o)}</span><i style="width:${(n / maxTop) * 100}%"></i><em>${fmtInt(n)} · ${pct(n / Math.max(1, r.n))}%</em></div>`).join('')}</div></section>
        <section><h3>Least active agents, Gram-positive · %S${nr(1)}</h3><p class="hint muted" style="margin:-6px 0 8px;font-size:var(--t-xs)">Agents with at least ${MIN_LIST} isolates tested.</p>${leastTable(least('positive'))}</section>
        <section><h3>Least active agents, Gram-negative · %S${nr(1)}</h3><p class="hint muted" style="margin:-6px 0 8px;font-size:var(--t-xs)">Agents with at least ${MIN_LIST} isolates tested.</p>${leastTable(least('negative'))}</section>
        <section><h3>Most active in ${esc(subsetLabel())} · %S${nr(1)}</h3><p class="hint muted" style="margin:-6px 0 8px;font-size:var(--t-xs)">Agents with at least ${MIN_LIST} isolates tested.</p>${leastTable(r.singles.filter((s) => s.tested >= MIN_LIST).slice(0, 10))}</section>
      </div>
      <div class="ledger-foot"><button class="btn small" data-view="singles" type="button">All single agents</button><button class="btn small" data-view="combos" type="button">Combinations</button></div>`;
    $('#view-tools').innerHTML = '';
  }
  function warnings(r) {
    const st = S.prep.stats;
    const out = [];
    if (st.invalidValues.length) out.push(`${st.invalidValues.reduce((a, b) => a + b[1], 0)} cells with unrecognised values were counted as not tested: ${st.invalidValues.slice(0, 6).map(([v, n]) => `“${esc(v)}” ×${n}`).join(', ')}${st.invalidValues.length > 6 ? '…' : ''}.`);
    if (st.unclassified.length) out.push(`${st.unclassified.length} organism names are unclassified and appear only under All isolates and genus/species. Assign them in step 3.`);
    if (r.dropped.length && r.n) out.push(`${r.dropped.length} agent${r.dropped.length > 1 ? 's' : ''} below the ${S.params.threshold}% testing threshold in this subset: ${r.dropped.map((d) => esc(d.agent)).join(', ')}.`);
    return out.map((t) => `<div class="warn" style="margin-top:12px">${t}</div>`).join('');
  }

  /* ---- single agents ---- */
  function sortRows(rows, sort, getters) {
    const g = getters[sort.key] || ((x) => x[sort.key]);
    return rows.slice().sort((a, b) => {
      const va = g(a), vb = g(b);
      if (typeof va === 'string') return sort.dir * va.localeCompare(vb);
      return sort.dir * ((isFinite(va) ? va : -1) - (isFinite(vb) ? vb : -1));
    });
  }
  function filteredSingles(r) {
    let rows = r.singles;
    if (S.classFilter) rows = rows.filter((s) => s.cls === S.classFilter);
    if (S.query) rows = rows.filter((s) => s.agent.toLowerCase().includes(S.query.toLowerCase()));
    return sortRows(rows, S.sort.singles, { agent: (x) => x.agent.toLowerCase() });
  }
  function ival(lo, hi, p, gp) {
    const ticks = [0.25, 0.5, 0.75].map((t) => `<span class="t" style="left:${t * 100}%"></span>`).join('');
    return `<div class="ival" aria-hidden="true">${ticks}${isFinite(lo) ? `<span class="ci-l" style="left:${lo * 100}%;width:${Math.max(0.5, (hi - lo) * 100)}%"></span>` : ''}${isFinite(gp) ? `<span class="cv" style="left:${gp * 100}%"></span>` : ''}${isFinite(p) ? `<span class="pt" style="left:${p * 100}%"></span>` : ''}</div>`;
  }
  function th_(key, label, sortState, cls = '', note) {
    const s = sortState.key === key ? (sortState.dir === 1 ? 'ascending' : 'descending') : 'none';
    return `<th scope="col" class="sortable ${cls}" data-sort="${key}" aria-sort="${s}" tabindex="0">${label}${note ? `<button class="nref" data-note="${note}" type="button" aria-label="Note ${note}">${note}</button>` : ''}</th>`;
  }
  function viewSingles(r) {
    const rows = filteredSingles(r);
    const classes = Array.from(new Set(r.singles.map((s) => s.cls)));
    const head = `<div class="sheet-head"><h1>Single agents</h1><span class="sub">${esc(subsetLabel())} · ${fmtInt(r.n)} isolates · ${confPct()}% CI</span><span class="grow"></span>
      <div class="filters"><label>Class <select id="in-cls"><option value="">All classes</option>${E.DISPLAY_CLASS_ORDER.filter((c) => classes.includes(c)).map((c) => `<option ${c === S.classFilter ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
      <label><input type="search" id="in-q" placeholder="Find agent" value="${esc(S.query)}" style="width:140px"></label>
      ${S.singlesMode === 'figure' ? '' : `<label><input type="checkbox" id="in-group" ${S.groupByClass ? 'checked' : ''}> Group by class</label>`}</div></div>${warnings(r)}`;
    $('#view-tools').innerHTML = modeToggle('singlesMode', [['table', 'Table', ICON.table], ['figure', 'Figure', ICON.chart]], S.singlesMode);
    if (S.singlesMode === 'figure') {
      $('#sheet').innerHTML = head + `<div class="figure" id="figure">${singlesFigure(r, rows)}</div>`;
      return;
    }
    const body = groupRows(rows, (s) => `
      <tr data-class="${esc(s.cls)}" data-sel="single" data-key="${esc(s.agent)}" class="${S.sel && S.sel.kind === 'single' && S.sel.key === s.agent ? 'sel' : ''}">
        <td><span class="clsdot"></span>${esc(s.agent)}</td>
        <td class="num">${fmtInt(s.tested)} <span class="ci">${pct(s.testedPct, 0)}%</span></td>
        <td class="num">${fmtInt(s.s)}</td><td class="num">${fmtInt(s.r)}</td>
        <td class="num"><span class="pct ${bandOf(s.pS)}">${pct(s.pS)}</span> <span class="ci">${pct(s.lo)}–${pct(s.hi)}</span></td>
        <td>${ival(s.lo, s.hi, s.pS, s.pG)}</td>
        <td class="num">${pct(s.pG)} <span class="ci">${pct(s.glo)}–${pct(s.ghi)}</span></td>
      </tr>`, (s) => s.cls, 10);
    $('#sheet').innerHTML = head + `<div class="tbl-wrap"><table class="ref" id="tbl"><thead><tr>
      ${th_('agent', 'Agent', S.sort.singles)}${th_('tested', 'Tested · % of n', S.sort.singles, 'num', 4)}
      ${th_('s', 'S', S.sort.singles, 'num')}${th_('r', 'R', S.sort.singles, 'num', 1)}
      ${th_('pS', `%S · ${confPct()}% CI`, S.sort.singles, 'num', 1)}
      <th scope="col"><span class="legend"><span>● %S</span><span>◇ coverage</span><button class="nref" data-note="3" type="button" aria-label="Note 3">3</button></span></th>
      ${th_('pG', 'Coverage · CI', S.sort.singles, 'num', 2)}</tr></thead><tbody>${body}</tbody></table></div>
      ${rows.length ? '' : '<p class="muted" style="padding:16px;margin:0">No agents match.</p>'}`;
  }
  function groupRows(rows, rowHtml, clsOf, colspan) {
    if (!S.groupByClass) return rows.map(rowHtml).join('');
    const groups = new Map();
    for (const r of rows) { const c = clsOf(r); if (!groups.has(c)) groups.set(c, []); groups.get(c).push(r); }
    const order = E.DISPLAY_CLASS_ORDER.filter((c) => groups.has(c));
    return order.map((c) => `<tr class="classband" data-class="${esc(c)}"><td colspan="${colspan}">${esc(c)}</td></tr>` +
      groups.get(c).map(rowHtml).join('')).join('');
  }
  function modeToggle(name, opts, val) {
    return `<div class="seg" role="group" aria-label="Display">${opts.map(([v, l, ic]) => `<button type="button" data-mode="${name}" data-val="${v}" aria-pressed="${v === val}" style="display:inline-flex;gap:6px;align-items:center">${ic}${l}</button>`).join('')}</div>`;
  }
  function figFoot() {
    return `${S.file.synthetic ? 'Synthetic demo data. ' : ''}Jeffreys ${confPct()}% CI. I counted as ${({ S: 'S', R: 'R', N: 'not tested' })[S.params.iMode]}. Susceptibility Index v2.`;
  }
  function singlesFigure(r, rows) {
    const list = (rows || filteredSingles(r)).filter((s) => s.tested > 0);
    if (!list.length) return '<p class="muted">Nothing to plot.</p>';
    return CH.singlesPlot(list.map((s) => Object.assign({}, s, { cls: s.cls })), {
      title: `Single agents · ${subsetLabel()} (n = ${r.n})`, subtitle: `${baseName()}`, confPct: confPct(), thresholds: th(), footnote: figFoot(),
    });
  }

  /* ---- combinations ---- */
  function filteredCombos(r) {
    let rows = r.combos;
    if (S.comboAgent) rows = rows.filter((c) => c.a === S.comboAgent || c.b === S.comboAgent);
    if (S.query) { const q = S.query.toLowerCase(); rows = rows.filter((c) => (c.a + ' ' + c.b).toLowerCase().includes(q)); }
    if (S.comboSigOnly) rows = rows.filter((c) => c.pValA < 0.05 && c.pValB < 0.05);
    return sortRows(rows, S.sort.combos, { name: (x) => (x.a + ' + ' + x.b).toLowerCase(), minDelta: (x) => Math.min(x.deltaA, x.deltaB) });
  }
  function viewCombos(r) {
    const rows = filteredCombos(r);
    const head = `<div class="sheet-head"><h1>Combinations</h1><span class="sub">${esc(subsetLabel())} · ${fmtInt(r.n)} isolates · ${r.combos.length} pairs</span><span class="grow"></span>
      <div class="filters"><label>With <select id="in-cagent"><option value="">any agent</option>${r.kept.map((a) => `<option ${a === S.comboAgent ? 'selected' : ''}>${esc(a)}</option>`).join('')}</select></label>
      <label><input type="search" id="in-q" placeholder="Find pair" value="${esc(S.query)}" style="width:140px"></label>
      <label title="Exact McNemar p < 0.05 against both components, no multiplicity correction"><input type="checkbox" id="in-sig" ${S.comboSigOnly ? 'checked' : ''}> Adds coverage to both (p&lt;0.05)</label></div></div>${warnings(r)}`;
    $('#view-tools').innerHTML = modeToggle('combosMode', [['table', 'Table', ICON.table], ['matrix', 'Matrix', ICON.grid]], S.combosMode);
    if (S.combosMode === 'matrix') {
      $('#sheet').innerHTML = head + `<div class="figure" id="figure">${comboFigure(r)}</div><p class="hint" style="padding:0 16px 16px;margin:0;color:var(--ink-3);font-size:var(--t-xs)">Select a cell for its derivation.</p>`;
      return;
    }
    const lim = S.comboShowAll ? rows.length : Math.min(rows.length, 150);
    const gain = (d, p) => `<span class="${p < 0.05 ? '' : 'muted'}">+${pct(d)}</span> <span class="ci">${pval(p)}</span>`;
    const nameCell = 'style="white-space:normal;min-width:190px"';
    $('#sheet').innerHTML = head + `<div class="tbl-wrap"><table class="ref" id="tbl"><thead><tr>
      ${th_('name', 'Combination', S.sort.combos)}${th_('tested', 'Tested', S.sort.combos, 'num', 5)}
      ${th_('s', 'S', S.sort.combos, 'num', 5)}
      ${th_('pS', `%S · ${confPct()}% CI`, S.sort.combos, 'num', 1)}
      ${th_('pG', 'Coverage', S.sort.combos, 'num', 2)}
      ${th_('deltaA', 'vs 1st: gain · p', S.sort.combos, 'num', 6)}${th_('deltaB', 'vs 2nd: gain · p', S.sort.combos, 'num', 6)}</tr></thead>
      <tbody>${rows.slice(0, lim).map((c) => `
        <tr data-sel="combo" data-key="${esc(c.a)}\t${esc(c.b)}" class="${S.sel && S.sel.kind === 'combo' && S.sel.key === c.a + '\t' + c.b ? 'sel' : ''}">
          <td ${nameCell}><span data-class="${esc(agentClass(c.a))}"><span class="clsdot"></span></span>${esc(c.a)} + <span data-class="${esc(agentClass(c.b))}"><span class="clsdot" style="margin-left:2px"></span></span>${esc(c.b)}</td>
          <td class="num">${fmtInt(c.tested)}</td><td class="num">${fmtInt(c.s)}</td>
          <td class="num"><span class="pct ${bandOf(c.pS)}">${pct(c.pS)}</span> <span class="ci">${pct(c.lo)}–${pct(c.hi)}</span></td>
          <td class="num">${pct(c.pG)}</td>
          <td class="num">${gain(c.deltaA, c.pValA)}</td><td class="num">${gain(c.deltaB, c.pValB)}</td></tr>`).join('')}</tbody></table></div>
      ${rows.length > lim ? `<div style="padding:12px 16px"><button class="btn small" data-action="combo-all" type="button">Show all ${fmtInt(rows.length)} pairs</button></div>` : ''}
      ${rows.length ? '' : '<p class="muted" style="padding:16px;margin:0">No combinations match.</p>'}`;
  }
  function orderedAgents(r) {
    const list = r.singles.map((s) => ({ agent: s.agent, cls: s.cls }));
    const ord = (c) => E.DISPLAY_CLASS_ORDER.indexOf(c);
    return list.sort((a, b) => ord(a.cls) - ord(b.cls) || a.agent.localeCompare(b.agent));
  }
  function comboFigure(r) {
    if (r.singles.length < 2) return '<p class="muted">At least two agents are needed.</p>';
    return CH.comboMatrix(orderedAgents(r), r.singles, r.combos, {
      title: `Two-drug combinations · ${subsetLabel()} (n = ${r.n})`,
      subtitle: `% susceptible to at least one agent, denominator: tested for ${S.params.combDen === 'both' ? 'both agents' : 'at least one agent'}`,
      thresholds: th(), footnote: figFoot(),
    });
  }

  /* ---- antibiogram ---- */
  function abgData(r) {
    const agents = orderedAgents(r);
    const idx = agents.map((a) => S.prep.agents.indexOf(a.agent));
    const grid = E.antibiogram(r.isolates, S.prep.agents, { by: S.params.abgBy, minIsolates: S.params.abgMin, maxRows: 40 })
      .map((g) => Object.assign({}, g, { cells: idx.map((j) => g.cells[j]) }));
    return { agents, grid };
  }
  function viewAbg(r) {
    const { agents, grid } = abgData(r);
    const head = `<div class="sheet-head"><h1>Antibiogram</h1><span class="sub">${esc(subsetLabel())} · %S among tested<button class="nref" data-note="1" type="button" aria-label="Note 1">1</button> · colour bands<button class="nref" data-note="9" type="button" aria-label="Note 9">9</button> · select a cell for its derivation</span><span class="grow"></span>
      <div class="filters"><label>Rows <select id="in-abgby"><option value="species" ${S.params.abgBy === 'species' ? 'selected' : ''}>Species</option><option value="genus" ${S.params.abgBy === 'genus' ? 'selected' : ''}>Genus</option></select></label>
      <label>Min. isolates <input type="number" id="in-abgmin" min="1" value="${S.params.abgMin}" style="width:60px"></label>
      <label title="Cells with fewer tested isolates are faded and starred">Faded below <input type="number" id="in-abgmt" min="1" value="${S.params.abgMinTested}" style="width:60px"> tested</label></div></div>`;
    $('#view-tools').innerHTML = '';
    if (!grid.length) { $('#sheet').innerHTML = head + `<p class="muted" style="padding:16px;margin:0">No ${S.params.abgBy} has at least ${S.params.abgMin} isolates in this subset.</p>`; return; }
    $('#sheet').innerHTML = head + `<div class="tbl-wrap"><table class="ref abg" id="tbl"><thead><tr><th scope="col" style="vertical-align:bottom">${S.params.abgBy === 'genus' ? 'Genus' : 'Species'}</th><th class="num" scope="col" style="vertical-align:bottom">n</th>
      ${agents.map((a) => `<th class="ag" scope="col" data-class="${esc(a.cls)}">${esc(a.agent)}</th>`).join('')}</tr></thead>
      <tbody>${grid.map((g) => `<tr><th scope="row">${esc(g.name)}</th><td class="num dim">${g.n}</td>${g.cells.map((c, j) => {
        const few = c.tested > 0 && c.tested < S.params.abgMinTested;
        const bd = c.tested ? bandOf(c.pS) : 'na';
        return `<td class="cell ${bd} ${few ? 'few' : ''}" data-sel="abg" data-key="${esc(g.name)}\t${esc(agents[j].agent)}" title="${esc(g.name)} · ${esc(agents[j].agent)}: ${c.s}/${c.tested} S">${c.tested ? Math.floor(c.pS * 100 + 1e-9) + (few ? '*' : '') : '–'}</td>`;
      }).join('')}</tr>`).join('')}</tbody></table></div>
      <p class="hint" style="padding:10px 16px 16px;margin:0;color:var(--ink-3);font-size:var(--t-xs)">Cells show % susceptible among tested, truncated to the integer; * fewer than ${S.params.abgMinTested} tested; – not tested. Agents follow the subset's testing threshold.</p>`;
  }
  function abgFigure(r) {
    const { agents, grid } = abgData(r);
    if (!grid.length) return null;
    return CH.antibiogramPlot(grid, agents, { title: `Antibiogram · ${subsetLabel()}`, subtitle: `${baseName()} · rows: ${S.params.abgBy}, ≥ ${S.params.abgMin} isolates`, thresholds: th(), minTested: S.params.abgMinTested, footnote: figFoot() });
  }

  /* ---- organisms ---- */
  function viewOrg() {
    const names = countBy(S.prep.isolates, 'organism');
    const byName = new Map(S.prep.isolates.map((i) => [i.organism, i]));
    const q = S.orgQuery.toLowerCase();
    const rows = names.filter(([o]) => !q || o.toLowerCase().includes(q));
    rows.sort((a, b) => (byName.get(a[0]).category === 'unclassified' ? -1 : 0) - (byName.get(b[0]).category === 'unclassified' ? -1 : 0) || b[1] - a[1]);
    $('#view-tools').innerHTML = '';
    $('#sheet').innerHTML = `<div class="sheet-head"><h1>Organisms</h1><span class="sub">${names.length} names in the analysed isolates · map of ${S.gramMap.length + S.overrides.length} entries</span><span class="grow"></span>
      <div class="map-tools"><input type="search" id="in-orgq" placeholder="Find organism" value="${esc(S.orgQuery)}" style="width:170px">
      <button class="btn small" data-action="map-import" type="button">Import map</button><button class="btn small" data-action="map-reset" type="button">Reset to bundled map</button>
      <input type="file" id="file-map" accept=".xlsx,.xls,.csv" hidden></div></div>
      <div class="tbl-wrap"><table class="ref" id="tbl"><thead><tr><th scope="col">Organism as written</th><th class="num" scope="col">Isolates</th><th scope="col">Group<button class="nref" data-note="8" type="button" aria-label="Note 8">8</button></th><th scope="col">Matched map entry</th></tr></thead>
      <tbody>${rows.map(([o, n]) => { const it = byName.get(o); return `<tr><td><i>${esc(o)}</i></td><td class="num">${fmtInt(n)}</td>
        <td><select class="cat" data-assign="${esc(o)}" aria-label="Group for ${esc(o)}">${['positive', 'negative', 'fungus', 'other', 'unclassified'].map((c) => `<option value="${c === 'unclassified' ? '' : c}" ${c === it.category ? 'selected' : ''}>${CAT_LABEL[c]}</option>`).join('')}</select></td>
        <td class="dim">${it.matchKey ? esc(it.matchKey) : '<span style="color:var(--mid)">none</span>'}${S.overrides.some((x) => E.normName(x[0]) === E.normName(o)) ? ' · set here' : ''}</td></tr>`; }).join('')}</tbody></table></div>
      <p class="hint" style="padding:10px 16px 16px;margin:0;color:var(--ink-3);font-size:var(--t-xs)">Changing a group adds an exact-name entry to this session's map. Export it from Export → Organism map and import it next time.</p>`;
  }

  /* ---- notes column: derivations + numbered method notes ---- */
  function methodNotes(r) {
    const p = S.params;
    const iTxt = { S: 'counted as S', R: 'counted as R', N: 'treated as not tested' }[p.iMode];
    const dedupTxt = !S.map.patientCol || p.dedup === 'none' ? 'All isolates kept (no patient column or deduplication off).' :
      `First isolate per ${p.dedup === 'patient' ? 'patient' : 'patient and organism'}, earliest by ${S.map.dateCol ? `<b>${esc(S.map.dateCol)}</b>` : 'file order'}; ${fmtInt(S.prep.stats.dedupRemoved)} removed.`;
    return [
      `<b>%S</b> = S / (S + R) among isolates with a result for the agent. I results ${iTxt}.`,
      `<b>Coverage</b> = S / all isolates in the subset; untested isolates count as not covered.`,
      `<b>Jeffreys ${confPct()}% interval</b>: quantiles of Beta(S + ½, failures + ½); lower bound 0 when S = 0, upper bound 100 when there are no failures.`,
      p.threshold > 0 ? `Agents tested in fewer than <b>${p.threshold}%</b> of the subset's isolates are left out.${r.dropped.length ? ` Here: ${r.dropped.map((d) => esc(d.agent)).join(', ')}.` : ''}` : `No minimum-tested threshold: every agent with at least one result is analysed.`,
      `A <b>combination</b> counts as S when either agent is S. Denominator: isolates tested for ${p.combDen === 'both' ? 'both agents' : 'at least one agent (v1 rule)'}. “Not S” includes R to both${p.combDen === 'both' ? '' : ' and R to one with the other untested'}.`,
      `<b>Gain vs an agent</b>: isolates R to that agent and S to its partner, over isolates tested for that agent. p: exact two-sided McNemar on the same isolates; not adjusted for multiple comparisons.`,
      `<b>Repeat isolates</b>: ${dedupTxt}`,
      `<b>Groups</b> from the organism map: exact name first, otherwise the longest map entry found at the start of a word. Fungi and unclassified names are outside the Gram subsets.`,
      `<b>Colour bands</b>: green ≥ ${p.good}%, amber ${p.mid}–${p.good}%, red &lt; ${p.mid}%. The number is always printed.`,
    ];
  }
  function derivation(r) {
    const sel = S.sel;
    if (!sel) return `<div class="deriv"><p class="empty">Select a row or cell to see how its figures were derived.</p></div>`;
    if (sel.kind === 'single') {
      const s = r.singles.find((x) => x.agent === sel.key);
      if (!s) return '';
      const [a, b] = [s.s + 0.5, s.r + 0.5];
      return `<div class="deriv" data-class="${esc(s.cls)}"><h4><span class="clsdot"></span>${esc(s.agent)}</h4><dl>
        <dt>Isolates in subset</dt><dd>${fmtInt(s.n)}</dd><dt>With a result</dt><dd>${fmtInt(s.tested)}</dd><dt>Not tested</dt><dd>${fmtInt(s.nt)}</dd>
        <dt>S</dt><dd>${fmtInt(s.s)}</dd><dt>R</dt><dd>${fmtInt(s.r)}</dd></dl>
        <div class="eq">%S = ${s.s} / ${s.tested} = ${pct(s.pS, 2)}%
CI ${pct(s.lo)}–${pct(s.hi)}%, Beta(${a}, ${b})
Coverage = ${s.s} / ${s.n} = ${pct(s.pG, 2)}%
CI ${pct(s.glo)}–${pct(s.ghi)}%, Beta(${s.s + 0.5}, ${s.n - s.s + 0.5})</div></div>`;
    }
    if (sel.kind === 'combo') {
      const [ka, kb] = sel.key.split('\t');
      const c = r.combos.find((x) => x.a === ka && x.b === kb);
      if (!c) return '';
      const A = r.singles.find((x) => x.agent === c.a), B = r.singles.find((x) => x.agent === c.b);
      return `<div class="deriv"><h4>${esc(c.a)} + ${esc(c.b)}</h4><dl>
        <dt>Tested (${S.params.combDen === 'both' ? 'both' : '≥1'})</dt><dd>${fmtInt(c.tested)}</dd><dt>Tested for both</dt><dd>${fmtInt(c.bothTested)}</dd>
        <dt>S to either</dt><dd>${fmtInt(c.s)}</dd><dt>R to both</dt><dd>${fmtInt(c.bothR)}</dd>${S.params.combDen === 'both' ? '' : `<dt>R to one, other untested</dt><dd>${fmtInt(c.rN)}</dd>`}</dl>
        <div class="eq">%S = ${c.s} / ${c.tested} = ${pct(c.pS, 2)}%
CI ${pct(c.lo)}–${pct(c.hi)}%
Coverage = ${c.sG} / ${c.n} = ${pct(c.pG, 2)}%

vs ${c.a}: ${A.s} S of ${A.tested} tested
 + ${c.gainA} R to ${c.a}, S to ${c.b}
 = ${A.s + c.gainA}/${A.tested} = ${pct(c.pA, 2)}%  (+${pct(c.deltaA, 2)} pp)
 McNemar exact p ${pEq(c.pValA)}

vs ${c.b}: ${B.s} S of ${B.tested} tested
 + ${c.gainB} R to ${c.b}, S to ${c.a}
 = ${B.s + c.gainB}/${B.tested} = ${pct(c.pB, 2)}%  (+${pct(c.deltaB, 2)} pp)
 McNemar exact p ${pEq(c.pValB)}</div></div>`;
    }
    if (sel.kind === 'abg') {
      const [org, ag] = sel.key.split('\t');
      const j = S.prep.agents.indexOf(ag);
      const key = S.params.abgBy === 'genus' ? 'genus' : 'species';
      const list = r.isolates.filter((i) => i[key] === org);
      let s = 0, rr = 0; for (const i of list) { if (i.vals[j] === 1) s++; else if (i.vals[j] === 2) rr++; }
      const [lo, hi] = E.jeffreys(s, rr, S.params.conf);
      return `<div class="deriv"><h4><i>${esc(org)}</i> · ${esc(ag)}</h4><dl><dt>Isolates</dt><dd>${list.length}</dd><dt>Tested</dt><dd>${s + rr}</dd><dt>S</dt><dd>${s}</dd><dt>R</dt><dd>${rr}</dd></dl>
        <div class="eq">%S = ${s} / ${s + rr} = ${s + rr ? pct(s / (s + rr), 2) : '–'}%
${confPct()}% CI ${s + rr ? pct(lo, 1) + '–' + pct(hi, 1) : '–'}%</div></div>`;
    }
    return '';
  }
  function renderNotes(r) {
    const notes = methodNotes(r);
    $('#notes').innerHTML = `<h2>Derivation</h2>${derivation(r)}<h2>Notes</h2><ol class="mnotes">${notes.map((t, i) => `<li id="note-${i + 1}" class="${S.hlNote === i + 1 ? 'hl' : ''}"><span>${t}</span></li>`).join('')}</ol>`;
  }

  /* ---------------- exports ---------------- */
  function singlesRows(r) {
    return r.singles.map((s) => ({
      Agent: s.agent, Class: s.cls, 'Isolates in subset': s.n, Tested: s.tested, 'Tested %': round(s.testedPct * 100), S: s.s, R: s.r,
      '%S': round(s.pS * 100), [`%S CI low (${confPct()}%)`]: round(s.lo * 100), '%S CI high': round(s.hi * 100),
      'Coverage %': round(s.pG * 100), 'Coverage CI low': round(s.glo * 100), 'Coverage CI high': round(s.ghi * 100),
    }));
  }
  function combosRows(r) {
    return r.combos.map((c) => ({
      Combination: `${c.a} + ${c.b}`, 'Agent 1': c.a, 'Agent 2': c.b, 'Isolates in subset': c.n, Tested: c.tested, 'Tested for both': c.bothTested,
      S: c.s, 'Not S': c.f, 'R to both': c.bothR, 'R to one, other untested': c.rN,
      '%S': round(c.pS * 100), [`%S CI low (${confPct()}%)`]: round(c.lo * 100), '%S CI high': round(c.hi * 100),
      'Coverage %': round(c.pG * 100), 'Coverage CI low': round(c.glo * 100), 'Coverage CI high': round(c.ghi * 100),
      'Gain vs agent 1 (isolates)': c.gainA, 'Gain vs agent 1 (pp)': round(c.deltaA * 100), 'p vs agent 1 (McNemar exact)': sig(c.pValA),
      'Gain vs agent 2 (isolates)': c.gainB, 'Gain vs agent 2 (pp)': round(c.deltaB * 100), 'p vs agent 2 (McNemar exact)': sig(c.pValB),
    }));
  }
  function round(x) { return isFinite(x) ? Math.round(x * 100) / 100 : null; }
  function sig(p) { return isFinite(p) ? Number(p.toPrecision(4)) : null; }
  function stripTags(h) { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; }

  function exportXlsx() {
    const wb = XLSX.utils.book_new();
    const st = S.prep.stats, all = S.prep.isolates;
    const sum = E.summary(all);
    const aoa = [
      ['Susceptibility Index v2 — results'], [],
      ['File', S.file.name + (S.file.synthetic ? ' (SYNTHETIC demo data)' : '')], ['Sheet', S.file.sheet], ['Exported', new Date().toISOString().slice(0, 16).replace('T', ' ')], [],
      ['Rows read', st.rowsRead], ['No organism name', st.emptyOrganism], ['Excluded by name', st.excluded], ['Repeat isolates removed', st.dedupRemoved], ['Analysed', st.analysed], [],
      ['Group', 'Isolates', '%'], ...sum.categories.map(([c, n]) => [CAT_LABEL[c] || c, n, round((n / all.length) * 100)]), [],
      ['Top 10 organisms', 'Isolates', '%'], ...sum.organisms.slice(0, 10).map(([o, n]) => [o, n, round((n / all.length) * 100)]), [],
    ];
    for (const cat of ['positive', 'negative']) {
      const rr = resultsFor({ type: 'gram', value: cat });
      aoa.push([`10 least active agents, ${CAT_LABEL[cat]} (≥${MIN_LIST} tested)`, '%S', 'S', 'Tested']);
      rr.singles.filter((s) => s.tested >= MIN_LIST).sort((a, b) => a.pS - b.pS).slice(0, 10).forEach((s) => aoa.push([s.agent, round(s.pS * 100), s.s, s.tested]));
      aoa.push([]);
    }
    const wsS = XLSX.utils.aoa_to_sheet(aoa); wsS['!cols'] = [{ wch: 42 }, { wch: 12 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsS, 'Summary');
    const p = S.params;
    const meth = [['Parameter', 'Value'], ['Organism column', S.map.orgCol], ['Patient ID column', S.map.patientCol || '—'], ['Date column', S.map.dateCol || '—'],
      ['Antibiotics', S.prep.agents.join(', ')], ['I results', { S: 'counted as S', R: 'counted as R', N: 'not tested' }[p.iMode]],
      ['Repeat isolates', !S.map.patientCol ? 'not applied' : { none: 'kept', 'patient-organism': 'first per patient and organism', patient: 'first per patient' }[p.dedup]],
      ['Exclusion terms', S.exclusions.split(/\n|;/).map((x) => x.trim()).filter(Boolean).join('; ') || '—'],
      ['Minimum tested %', p.threshold], ['Confidence level', p.conf], ['Combination denominator', p.combDen === 'both' ? 'tested for both' : 'tested for at least one'],
      [], ['Notes'], ...methodNotes(resultsFor({ type: 'all' })).map((t, i) => [i + 1, stripTags(t)])];
    const wsM = XLSX.utils.aoa_to_sheet(meth); wsM['!cols'] = [{ wch: 26 }, { wch: 110 }];
    XLSX.utils.book_append_sheet(wb, wsM, 'Methods');
    const subs = [[{ type: 'all' }, 'All'], [{ type: 'gram', value: 'positive' }, 'Gram+'], [{ type: 'gram', value: 'negative' }, 'Gram-']];
    if (S.subset.type === 'genus' || S.subset.type === 'species') subs.push([S.subset, S.subset.value.replace(/[\\/?*[\]:]/g, '').slice(0, 20)]);
    for (const [sub, label] of subs) {
      const rr = resultsFor(sub);
      const a = XLSX.utils.json_to_sheet(singlesRows(rr)); a['!cols'] = [{ wch: 30 }, { wch: 22 }];
      XLSX.utils.book_append_sheet(wb, a, `${label} - Singles`.slice(0, 31));
      const b = XLSX.utils.json_to_sheet(combosRows(rr)); b['!cols'] = [{ wch: 44 }, { wch: 24 }, { wch: 24 }];
      XLSX.utils.book_append_sheet(wb, b, `${label} - Combs`.slice(0, 31));
    }
    const ab = abgData(resultsFor(S.subset));
    if (ab.grid.length) {
      const rows = ab.grid.map((g) => { const o = { [S.params.abgBy === 'genus' ? 'Genus' : 'Species']: g.name, n: g.n }; g.cells.forEach((c, j) => { o[ab.agents[j].agent] = c.tested ? round(c.pS * 100) : null; o[ab.agents[j].agent + ' (tested)'] = c.tested; }); return o; });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), `Antibiogram ${S.subset.type === 'all' ? 'All' : subsetLabel().slice(0, 18)}`.slice(0, 31));
    }
    if (st.invalidCells.length || st.unclassified.length) {
      const chk = [['Unrecognised values (first 200)'], ['Row', 'Agent', 'Value'], ...st.invalidCells.map((c) => [c.row, c.agent, c.value]), [], ['Unclassified organisms', 'Isolates'], ...st.unclassified];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(chk), 'Data checks');
    }
    XLSX.writeFile(wb, `${baseName()}_susceptibility_index.xlsx`, { compression: true });
    toast('Workbook exported');
  }
  function currentTableRows() {
    const r = resultsFor(S.subset);
    if (S.view === 'singles') { const keep = new Set(filteredSingles(r).map((s) => s.agent)); const order = filteredSingles(r).map((s) => s.agent); return singlesRows(r).filter((x) => keep.has(x.Agent)).sort((a, b) => order.indexOf(a.Agent) - order.indexOf(b.Agent)); }
    if (S.view === 'combos') { const fr = filteredCombos(r); const m = new Map(combosRows(r).map((x) => [x['Agent 1'] + '\t' + x['Agent 2'], x])); return fr.map((c) => m.get(c.a + '\t' + c.b)); }
    if (S.view === 'abg') { const ab = abgData(r); return ab.grid.map((g) => { const o = { Organism: g.name, n: g.n }; g.cells.forEach((c, j) => { o[ab.agents[j].agent] = c.tested ? round(c.pS * 100) : ''; }); return o; }); }
    if (S.view === 'org') { const by = new Map(S.prep.isolates.map((i) => [i.organism, i])); return countBy(S.prep.isolates, 'organism').map(([o, n]) => ({ Organism: o, Isolates: n, Group: CAT_LABEL[by.get(o).category], 'Matched entry': by.get(o).matchKey || '' })); }
    return singlesRows(r);
  }
  function exportCsv() {
    const rows = currentTableRows();
    if (!rows.length) { toast('Nothing to export'); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const name = { singles: 'singles', combos: 'combinations', abg: 'antibiogram', org: 'organisms' }[S.view] || 'singles';
    download(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }), `${baseName()}_${slug(subsetLabel())}_${name}.csv`);
    toast('CSV exported');
  }
  function slug(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, ''); }
  function currentFigure() {
    const r = resultsFor(S.subset);
    if (S.view === 'combos') return { svg: comboFigure(r), name: 'combinations' };
    if (S.view === 'abg') return { svg: abgFigure(r), name: 'antibiogram' };
    return { svg: singlesFigure(r), name: 'singles' };
  }
  function exportSvg() {
    const f = currentFigure();
    if (!f.svg || !f.svg.startsWith('<svg')) { toast('No figure for this view'); return; }
    download(new Blob(['<?xml version="1.0" encoding="UTF-8"?>\n' + f.svg], { type: 'image/svg+xml' }), `${baseName()}_${slug(subsetLabel())}_${f.name}.svg`);
    toast('SVG exported');
  }
  function exportPng() {
    const f = currentFigure();
    if (!f.svg || !f.svg.startsWith('<svg')) { toast('No figure for this view'); return; }
    const m = f.svg.match(/viewBox="([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+)"/);
    const w = +m[3], h = +m[4], scale = 300 / 96;
    const img = new Image();
    const url = URL.createObjectURL(new Blob([f.svg.replace(/width="[\d.]+" height="[\d.]+"/, `width="${w}" height="${h}"`)], { type: 'image/svg+xml' }));
    img.onload = () => {
      const c = document.createElement('canvas'); c.width = Math.round(w * scale); c.height = Math.round(h * scale);
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      c.toBlob((b) => withDpi(b, (out) => { download(out, `${baseName()}_${slug(subsetLabel())}_${f.name}.png`); toast('PNG exported at 300 dpi'); }), 'image/png');
    };
    img.onerror = () => toast('PNG rendering failed in this browser; use SVG.');
    img.src = url;
  }
  // Write a pHYs chunk (300 dpi) after IHDR so layout apps size the image correctly
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(bytes) { let c = 0xffffffff; for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  function withDpi(blob, cb) {
    blob.arrayBuffer().then((buf) => {
      const src = new Uint8Array(buf);
      const ppm = Math.round(300 / 0.0254);
      const chunk = new Uint8Array(21);
      const dv = new DataView(chunk.buffer);
      dv.setUint32(0, 9); chunk.set([0x70, 0x48, 0x59, 0x73], 4); dv.setUint32(8, ppm); dv.setUint32(12, ppm); chunk[16] = 1;
      dv.setUint32(17, crc32(chunk.subarray(4, 17)));
      const at = 8 + 25; // signature + IHDR chunk
      const out = new Uint8Array(src.length + 21);
      out.set(src.subarray(0, at), 0); out.set(chunk, at); out.set(src.subarray(at), at + 21);
      cb(new Blob([out], { type: 'image/png' }));
    });
  }
  function exportMap() {
    const rows = S.overrides.concat(S.gramMap).map(([o, c]) => ({ Microorganism: o, Gram_stain: c === 'positive' ? 'Gram-positive' : c === 'negative' ? 'Gram-negative' : c }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Map');
    XLSX.writeFile(wb, 'organism_map.xlsx');
    toast('Organism map exported');
  }
  function settingsObj() {
    return { app: 'susceptibility-index', version: 2, mapping: S.map ? { orgCol: S.map.orgCol, patientCol: S.map.patientCol, dateCol: S.map.dateCol, agents: S.map.agents.map((a) => ({ name: a.name, include: a.include, cls: a.cls })) } : null, params: S.params, exclusions: S.exclusions, organismOverrides: S.overrides };
  }
  function exportSettings() {
    download(new Blob([JSON.stringify(settingsObj(), null, 2)], { type: 'application/json' }), `${baseName()}_settings.json`);
    toast('Settings exported');
  }
  function applySettings(o, silent) {
    if (!o || o.app !== 'susceptibility-index') { toast('Not a Susceptibility Index settings file'); return; }
    Object.assign(S.params, o.params || {});
    S.exclusions = o.exclusions || '';
    S.overrides = Array.isArray(o.organismOverrides) ? o.organismOverrides : [];
    if (o.mapping && S.file) {
      const cols = new Set(S.file.columns);
      const m = o.mapping;
      if (cols.has(m.orgCol)) S.map.orgCol = m.orgCol;
      S.map.patientCol = cols.has(m.patientCol) ? m.patientCol : null;
      S.map.dateCol = cols.has(m.dateCol) ? m.dateCol : null;
      const saved = new Map((m.agents || []).map((a) => [a.name, a]));
      for (const a of S.map.agents) { const s = saved.get(a.name); if (s) { a.include = s.include; a.cls = s.cls; } else a.include = false; }
      for (const [name, s] of saved) if (!S.map.agents.some((a) => a.name === name) && cols.has(name)) S.map.agents.push({ name, include: s.include, cls: s.cls, tested: 0 });
    }
    if (!silent) toast('Settings applied');
  }
  function printReport() {
    const r = resultsFor(S.subset);
    const notes = methodNotes(r);
    const top = r.combos.slice(0, 25);
    $('#report').innerHTML = `<h1>Susceptibility Index · ${esc(subsetLabel())}</h1>
      <div class="meta">${esc(S.file.name)}${S.file.synthetic ? ' (synthetic demo data)' : ''} · ${fmtInt(r.n)} isolates · ${new Date().toISOString().slice(0, 10)}</div>
      <h2>Single agents</h2><table><thead><tr><th>Agent</th><th>Class</th><th class="num">Tested</th><th class="num">S</th><th class="num">R</th><th class="num">%S (${confPct()}% CI)</th><th class="num">Coverage (CI)</th></tr></thead>
      <tbody>${r.singles.map((s) => `<tr><td>${esc(s.agent)}</td><td>${esc(s.cls)}</td><td class="num">${s.tested}</td><td class="num">${s.s}</td><td class="num">${s.r}</td><td class="num">${pct(s.pS)} (${pct(s.lo)}–${pct(s.hi)})</td><td class="num">${pct(s.pG)} (${pct(s.glo)}–${pct(s.ghi)})</td></tr>`).join('')}</tbody></table>
      <div class="fig">${singlesFigure(r, r.singles)}</div>
      <h2>Top ${top.length} combinations by %S</h2><table><thead><tr><th>Combination</th><th class="num">Tested</th><th class="num">S</th><th class="num">%S (CI)</th><th class="num">Coverage</th><th class="num">Gain vs 1st (p)</th><th class="num">Gain vs 2nd (p)</th></tr></thead>
      <tbody>${top.map((c) => `<tr><td>${esc(c.a)} + ${esc(c.b)}</td><td class="num">${c.tested}</td><td class="num">${c.s}</td><td class="num">${pct(c.pS)} (${pct(c.lo)}–${pct(c.hi)})</td><td class="num">${pct(c.pG)}</td><td class="num">+${pct(c.deltaA)} (${pval(c.pValA)})</td><td class="num">+${pct(c.deltaB)} (${pval(c.pValB)})</td></tr>`).join('')}</tbody></table>
      <h2>Methods</h2><ol>${notes.map((t) => `<li>${t}</li>`).join('')}</ol>`;
    window.print();
  }

  /* ---------------- events ---------------- */
  function bind() {
    document.addEventListener('click', (e) => {
      const t = e.target.closest('button, [data-sel], .mx, [data-action], label.drop');
      if (!t) {
        if (!e.target.closest('.menu')) closeMenu();
        return;
      }
      if (!t.closest('.menu')) closeMenu();
      const act = t.dataset.action;
      if (t.matches('label.drop') && e.target.tagName !== 'INPUT') { e.preventDefault(); $('#file-data').click(); return; }
      if (act === 'demo') { loadDemo(); return; }
      if (act === 'pick-file') { const f = $('#file-data'); if (f) f.click(); return; }
      if (act === 'excl-file') { $('#file-excl').click(); return; }
      if (act === 'ag-all' || act === 'ag-none') { S.map.agents.forEach((a) => { a.include = act === 'ag-all'; }); renderCols(); schedule(true, true); return; }
      if (act === 'view-org') { S.view = 'org'; renderMain(); return; }
      if (act === 'combo-all') { S.comboShowAll = true; renderMain(); return; }
      if (act === 'map-import') { $('#file-map').click(); return; }
      if (act === 'map-reset') { S.gramMap = (window.SI_DEFAULT_GRAM_MAP || []).map((x) => x.slice()); S.overrides = []; schedule(true, true); toast('Bundled organism map restored'); return; }
      if (t.id === 'btn-export') { const m = $('#export-menu'); const open = m.hidden; m.hidden = !open; t.setAttribute('aria-expanded', String(open)); return; }
      if (t.id === 'btn-settings-load') { $('#file-settings').click(); return; }
      if (t.dataset.export) {
        closeMenu();
        ({ xlsx: exportXlsx, csv: exportCsv, svg: exportSvg, png: exportPng, print: printReport, map: exportMap, settings: exportSettings })[t.dataset.export]();
        return;
      }
      if (t.dataset.seg) { S.params[t.dataset.seg] = t.dataset.val; renderParams(); schedule(t.dataset.seg === 'iMode', true); return; }
      if (t.dataset.mode) { S[t.dataset.mode] = t.dataset.val; renderMain(); return; }
      if (t.dataset.sub) {
        S.subset = t.dataset.sub === 'all' ? { type: 'all' } : { type: t.dataset.sub, value: t.dataset.subval };
        S.sel = null; S.comboAgent = ''; S.comboShowAll = false; renderMain(); return;
      }
      if (t.dataset.view) { S.view = t.dataset.view; S.query = ''; S.sel = null; renderMain(); return; }
      if (t.dataset.note) { S.hlNote = +t.dataset.note; const li = $('#note-' + t.dataset.note); document.querySelectorAll('.mnotes li').forEach((x) => x.classList.remove('hl')); if (li) { li.classList.add('hl'); li.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } e.stopPropagation(); return; }
      if (t.matches('th.sortable')) return;
      if (t.dataset.sel) { S.sel = { kind: t.dataset.sel, key: t.dataset.key }; markSel(t); return; }
      if (t.matches('.mx') && t.dataset.a) {
        const r = resultsFor(S.subset);
        const c = r.combos.find((x) => (x.a === t.dataset.a && x.b === t.dataset.b) || (x.a === t.dataset.b && x.b === t.dataset.a));
        if (c) { S.sel = { kind: 'combo', key: c.a + '\t' + c.b }; renderNotes(r); }
      }
    });
    document.addEventListener('click', (e) => {
      const h = e.target.closest('th.sortable');
      if (!h || e.target.closest('.nref')) return;
      sortBy(h.dataset.sort);
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeMenu();
      const h = e.target.closest && e.target.closest('th.sortable');
      if (h && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); sortBy(h.dataset.sort); }
      const d = e.target.closest && e.target.closest('label.drop');
      if (d && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); $('#file-data').click(); }
    });
    document.addEventListener('change', (e) => {
      const t = e.target;
      if (t.id === 'file-data') { handleFile(t.files[0]); t.value = ''; return; }
      if (t.id === 'file-settings') { readJson(t.files[0]); t.value = ''; return; }
      if (t.id === 'file-excl') { readExclusions(t.files[0]); t.value = ''; return; }
      if (t.id === 'file-map') { readMap(t.files[0]); t.value = ''; return; }
      if (t.id === 'in-sheet') { const { rows, columns } = readSheet(S.file.wb, t.value); Object.assign(S.file, { sheet: t.value, rows, columns }); afterRowsLoaded(); return; }
      if (t.id === 'in-org' || t.id === 'in-pid' || t.id === 'in-date') {
        const k = { 'in-org': 'orgCol', 'in-pid': 'patientCol', 'in-date': 'dateCol' }[t.id];
        S.map[k] = t.value || null;
        // a column used as organism/ID/date is not an antibiotic
        S.map.agents.forEach((a) => { if (a.name === t.value) a.include = false; });
        renderCols(); renderParams(); schedule(true, true); return;
      }
      if (t.dataset.ag !== undefined) { S.map.agents[+t.dataset.ag].include = t.checked; schedule(true, true); return; }
      if (t.dataset.agcls !== undefined) { S.map.agents[+t.dataset.agcls].cls = t.value; t.closest('tr').dataset.class = t.value; t.title = t.value; S.cache.clear(); schedule(false, false); return; }
      if (t.dataset.assign !== undefined) {
        const name = t.dataset.assign;
        S.overrides = S.overrides.filter((x) => E.normName(x[0]) !== E.normName(name));
        if (t.value) S.overrides.unshift([name, t.value]);
        schedule(true, true); return;
      }
      if (t.id === 'in-dedup') { S.params.dedup = t.value; schedule(true, true); return; }
      if (t.id === 'in-conf') { S.params.conf = +t.value; schedule(false, false); return; }
      if (t.id === 'in-good' || t.id === 'in-mid') {
        const g = clamp(+$('#in-good').value, 1, 100), m = clamp(+$('#in-mid').value, 0, g - 1);
        S.params.good = g; S.params.mid = m; schedule(false, false); return;
      }
      if (t.id === 'in-pick') {
        if (!t.value) { S.subset = { type: 'all' }; } else { const [ty, ...v] = t.value.split('|'); S.subset = { type: ty, value: v.join('|') }; }
        S.sel = null; S.comboAgent = ''; renderMain(); return;
      }
      if (t.id === 'in-cls') { S.classFilter = t.value; renderMain(); return; }
      if (t.id === 'in-group') { S.groupByClass = t.checked; renderMain(); return; }
      if (t.id === 'in-cagent') { S.comboAgent = t.value; renderMain(); return; }
      if (t.id === 'in-sig') { S.comboSigOnly = t.checked; renderMain(); return; }
      if (t.id === 'in-abgby') { S.params.abgBy = t.value; S.sel = null; renderMain(); return; }
      if (t.id === 'in-abgmin') { S.params.abgMin = Math.max(1, +t.value || 1); renderMain(); return; }
      if (t.id === 'in-abgmt') { S.params.abgMinTested = Math.max(1, +t.value || 1); renderMain(); return; }
    });
    document.addEventListener('input', (e) => {
      const t = e.target;
      if (t.id === 'in-thr' || t.id === 'in-thr-n') {
        const v = clamp(+t.value || 0, 0, 100);
        S.params.threshold = v;
        const other = t.id === 'in-thr' ? $('#in-thr-n') : $('#in-thr'); if (other) other.value = v;
        schedule(false, false); return;
      }
      if (t.id === 'in-excl') { S.exclusions = t.value; clearTimeout(bind._x); bind._x = setTimeout(() => schedule(true, true), 300); return; }
      if (t.id === 'in-q') { S.query = t.value; clearTimeout(bind._q); bind._q = setTimeout(() => { renderMain(); const q = $('#in-q'); if (q) { q.focus(); q.setSelectionRange(q.value.length, q.value.length); } }, 180); return; }
      if (t.id === 'in-orgq') { S.orgQuery = t.value; clearTimeout(bind._o); bind._o = setTimeout(() => { renderMain(); const q = $('#in-orgq'); if (q) { q.focus(); q.setSelectionRange(q.value.length, q.value.length); } }, 180); return; }
    });
    // drag and drop
    document.addEventListener('dragover', (e) => { e.preventDefault(); const d = $('#drop'); if (d) d.classList.add('over'); });
    document.addEventListener('dragleave', (e) => { if (e.target === document.documentElement || !e.relatedTarget) { const d = $('#drop'); if (d) d.classList.remove('over'); } });
    document.addEventListener('drop', (e) => { e.preventDefault(); const d = $('#drop'); if (d) d.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
    window.addEventListener('beforeunload', (e) => { if (S.file && !S.file.synthetic) { e.preventDefault(); e.returnValue = ''; } });
  }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }
  function closeMenu() { const m = $('#export-menu'); if (m && !m.hidden) { m.hidden = true; $('#btn-export').setAttribute('aria-expanded', 'false'); } }
  function sortBy(key) {
    const k = S.view === 'combos' ? 'combos' : 'singles';
    const s = S.sort[k];
    if (s.key === key) s.dir = -s.dir; else { s.key = key; s.dir = (key === 'agent' || key === 'name') ? 1 : -1; }
    renderMain();
  }
  function markSel(el) {
    document.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel'));
    el.classList.add('sel');
    renderNotes(resultsFor(S.subset));
  }
  function readJson(file) {
    if (!file) return;
    file.text().then((txt) => {
      let o; try { o = JSON.parse(txt); } catch (err) { toast('Not a valid JSON file'); return; }
      if (S.file) { applySettings(o); recompute(true); renderRail(); renderMain(); }
      else { S.pendingSettings = o; toast('Settings will be applied when you load a data file'); }
    });
  }
  function firstColumn(file, cb) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, blankrows: false });
        cb(aoa);
      } catch (err) { toast('Could not read this file'); }
    };
    reader.readAsArrayBuffer(file);
  }
  function readExclusions(file) {
    if (!file) return;
    firstColumn(file, (aoa) => {
      const terms = aoa.map((r) => r[0]).filter((x) => x != null && String(x).trim()).map(String);
      S.exclusions = (S.exclusions ? S.exclusions.trim() + '\n' : '') + terms.join('\n');
      $('#in-excl').value = S.exclusions;
      schedule(true, true); toast(`${terms.length} exclusion terms added`);
    });
  }
  function readMap(file) {
    if (!file) return;
    firstColumn(file, (aoa) => {
      const out = [];
      aoa.forEach((r, i) => {
        if (!r[0] || !r[1]) return;
        const c = String(r[1]).toLowerCase();
        const cat = c.includes('pos') ? 'positive' : c.includes('neg') ? 'negative' : c.includes('fung') ? 'fungus' : c.includes('other') ? 'other' : null;
        if (cat) out.push([String(r[0]).trim(), cat]);
      });
      if (!out.length) { toast('No rows with an organism and a Gram group found'); return; }
      S.gramMap = out.concat(S.gramMap);
      schedule(true, true); toast(`${out.length} map entries imported`);
    });
  }

  bind();
  renderRail();
  renderMain();
})();
