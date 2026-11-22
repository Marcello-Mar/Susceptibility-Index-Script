/*
 * Susceptibility Index — analysis engine.
 * Pure functions, no DOM. Loaded as a classic script (window.SIEngine) and by Node tests (module.exports).
 */
(function (root) {
  'use strict';

  /* ---------- special functions ---------- */

  const LANCZOS = [676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.13857109526572012,
    9.9843695780195716e-6, 1.5056327351493116e-7];

  function lgamma(x) {
    if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lgamma(1 - x);
    x -= 1;
    let a = 0.99999999999980993;
    const t = x + 7.5;
    for (let i = 0; i < 8; i++) a += LANCZOS[i] / (x + i + 1);
    return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
  }

  // Continued fraction for the incomplete beta (Lentz)
  function betacf(a, b, x) {
    const MAXIT = 300, EPS = 1e-15, FPMIN = 1e-300;
    const qab = a + b, qap = a + 1, qam = a - 1;
    let c = 1, d = 1 - qab * x / qap;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    d = 1 / d;
    let h = d;
    for (let m = 1; m <= MAXIT; m++) {
      const m2 = 2 * m;
      let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; h *= d * c;
      aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d;
      const del = d * c; h *= del;
      if (Math.abs(del - 1) < EPS) break;
    }
    return h;
  }

  // Regularized incomplete beta I_x(a, b)
  function ibeta(x, a, b) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log(1 - x));
    if (x < (a + 1) / (a + b + 2)) return bt * betacf(a, b, x) / a;
    return 1 - bt * betacf(b, a, 1 - x) / b;
  }

  // Quantile of Beta(a, b) by bisection (monotone, robust)
  function betaPpf(p, a, b) {
    let lo = 0, hi = 1;
    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2;
      if (ibeta(mid, a, b) < p) lo = mid; else hi = mid;
      if (hi - lo < 1e-12) break;
    }
    return (lo + hi) / 2;
  }

  // Jeffreys interval, with the boundary rule used by the v1 script
  function jeffreys(s, f, conf) {
    const n = s + f;
    if (n === 0) return [NaN, NaN];
    const alpha = 1 - conf;
    let lo = betaPpf(alpha / 2, s + 0.5, f + 0.5);
    let hi = betaPpf(1 - alpha / 2, s + 0.5, f + 0.5);
    if (s === 0) lo = 0;
    if (f === 0) hi = 1;
    return [lo, hi];
  }

  // Exact two-sided McNemar test (binomial on discordant pairs)
  function mcnemarExact(b, c) {
    const n = b + c;
    if (n === 0) return 1;
    const k = Math.min(b, c);
    let cdf = 0;
    const logHalfN = n * Math.log(0.5);
    for (let i = 0; i <= k; i++) {
      cdf += Math.exp(lgamma(n + 1) - lgamma(i + 1) - lgamma(n - i + 1) + logHalfN);
    }
    return Math.min(1, 2 * cdf);
  }

  /* ---------- value normalisation ---------- */

  const N_TOKENS = new Set(['', 'N', 'NT', 'NA', 'N/A', '-', '--', '.', 'ND', 'NOT TESTED']);

  // codes: 0 = not tested, 1 = S, 2 = R ; returns -1 for unrecognised
  function normalizeValue(v, iMode) {
    if (v === null || v === undefined) return 0;
    const t = String(v).trim().toUpperCase();
    if (N_TOKENS.has(t)) return 0;
    if (t === 'S' || t === 'SUSCEPTIBLE' || t === 'SENSIBILE') return 1;
    if (t === 'R' || t === 'RESISTANT' || t === 'RESISTENTE') return 2;
    if (t === 'I' || t === 'SDD' || t === 'INTERMEDIATE' || t === 'INTERMEDIO') {
      if (iMode === 'R') return 2;
      if (iMode === 'N') return 0;
      return 1;
    }
    return -1;
  }

  /* ---------- organism classification ---------- */

  function normName(s) {
    return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();
  }

  function isWordStart(str, idx) {
    if (idx === 0) return true;
    return !/[a-z0-9]/i.test(str[idx - 1]);
  }

  // map: array of [name, category]. Exact name wins, then the longest key found at a word start.
  function buildClassifier(map) {
    const entries = map
      .filter((e) => e && e[0] && e[1])
      .map(([k, v]) => [normName(k), v])
      .sort((a, b) => b[0].length - a[0].length);
    const exact = new Map(entries);
    const cache = new Map();
    return function classify(name) {
      const n = normName(name);
      if (cache.has(n)) return cache.get(n);
      let out = { category: 'unclassified', key: null };
      if (exact.has(n)) out = { category: exact.get(n), key: n };
      else {
        for (const [k, v] of entries) {
          let idx = n.indexOf(k);
          while (idx !== -1) {
            if (isWordStart(n, idx)) { out = { category: v, key: k }; break; }
            idx = n.indexOf(k, idx + 1);
          }
          if (out.key) break;
        }
      }
      cache.set(n, out);
      return out;
    };
  }

  function genusOf(name) {
    const n = String(name == null ? '' : name).trim().replace(/\s+/g, ' ');
    const first = n.split(' ')[0] || '';
    return first.charAt(0).toUpperCase() + first.slice(1).toLowerCase();
  }

  function speciesOf(name) {
    const n = String(name == null ? '' : name).trim().replace(/\s+/g, ' ');
    const parts = n.split(' ');
    if (!parts[0]) return '';
    const g = parts[0].charAt(0).toUpperCase() + parts[0].slice(1).toLowerCase();
    return parts.length > 1 ? g + ' ' + parts.slice(1).join(' ').toLowerCase() : g;
  }

  /* ---------- antibiotic classes ---------- */

  const CLASS_RULES = [
    ['Carbapenems', /(imipen|meropen|ertapen|doripen|carbapen|biapen|tebipen)/],
    ['Monobactams', /(aztreon)/],
    ['Cephalosporins', /(cef|ceph|cefa|cefo|ceft|cefur|cefep|ceftaz|ceftri|ceftol|ceftar|ceftob|cefid|cefix|cefpod|cefox|cefaz|loracarb)/],
    ['Penicillins', /(penicil|amoxi|ampi|oxacil|cloxacil|dicloxacil|flucloxacil|methicil|meticil|nafcil|piperac|ticarc|temocil|mecillin|sulbact|benzylpen|tazobact)/],
    ['Aminoglycosides', /(gentam|amikac|tobram|netilm|streptomyc|kanam|plazomic)/],
    ['Fluoroquinolones', /(floxac|nalidix|quinolon)/],
    ['Glycopeptides', /(vancom|teicopl|dalbav|oritav|telavan)/],
    ['Lipopeptides', /(daptom)/],
    ['Oxazolidinones', /(linezol|tedizol)/],
    ['Macrolides & lincosamides', /(erythrom|eritrom|azithrom|azitrom|clarithrom|claritrom|clindam|lincom|spiram|josam|telithrom)/],
    ['Tetracyclines', /(tetracic|tetracyc|doxyc|doxic|minoc|tigec|eravac|omadac)/],
    ['Folate pathway inhibitors', /(trimet|sulfamet|cotrim|co-trim|tmp|smx|sulfon)/],
    ['Polymyxins', /(colist|polymyx|polimix)/],
    ['Fosfomycin', /(fosfom)/],
    ['Nitrofurans', /(nitrofur)/],
    ['Rifamycins', /(rifamp|rifab|rifax)/],
    ['Fusidanes', /(fusid)/],
    ['Nitroimidazoles', /(metronid|tinid)/],
    ['Streptogramins', /(quinupr|dalfopr|pristinam)/],
    ['Amphenicols', /(chloramph|cloramf)/],
  ];
  const CLASS_ORDER = CLASS_RULES.map((r) => r[0]).concat(['Other']);
  // Order used when listing agents: beta-lactams first, as reference tables do
  const DISPLAY_CLASS_ORDER = ['Penicillins', 'Cephalosporins', 'Carbapenems', 'Monobactams', 'Aminoglycosides',
    'Fluoroquinolones', 'Glycopeptides', 'Lipopeptides', 'Oxazolidinones', 'Macrolides & lincosamides',
    'Streptogramins', 'Tetracyclines', 'Folate pathway inhibitors', 'Polymyxins', 'Fosfomycin', 'Nitrofurans',
    'Rifamycins', 'Fusidanes', 'Nitroimidazoles', 'Amphenicols', 'Other'];

  function guessClass(agent) {
    const a = String(agent).toLowerCase();
    // beta-lactam/inhibitor combinations whose base is a cephalosporin
    if (/(ceftaz|ceftol|cefep|cefid|cefoper)/.test(a)) return 'Cephalosporins';
    for (const [cls, re] of CLASS_RULES) if (re.test(a)) return cls;
    return 'Other';
  }

  /* ---------- preparation ---------- */

  /*
   * rows: array of objects (header -> cell)
   * cfg: { orgCol, agentCols[], patientCol|null, dateCol|null, iMode: 'S'|'R'|'N',
   *        dedup: 'none'|'patient-organism'|'patient', exclusions: string[], gramMap: [[name, cat]] }
   */
  function prepare(rows, cfg) {
    const classify = buildClassifier(cfg.gramMap || []);
    const agents = cfg.agentCols.slice();
    const exclusions = (cfg.exclusions || []).map(normName).filter(Boolean);
    const invalid = new Map(); // value -> count
    const invalidCells = [];
    let excluded = 0, emptyOrg = 0;
    let list = [];

    rows.forEach((row, i) => {
      const orgRaw = row[cfg.orgCol];
      const org = orgRaw == null ? '' : String(orgRaw).trim().replace(/\s+/g, ' ');
      if (!org) { emptyOrg++; return; }
      const on = normName(org);
      if (exclusions.some((e) => on.includes(e))) { excluded++; return; }
      const vals = new Uint8Array(agents.length);
      for (let j = 0; j < agents.length; j++) {
        const raw = row[agents[j]];
        const code = normalizeValue(raw, cfg.iMode);
        if (code === -1) {
          const key = String(raw).trim();
          invalid.set(key, (invalid.get(key) || 0) + 1);
          if (invalidCells.length < 200) invalidCells.push({ row: i + 2, agent: agents[j], value: key });
          vals[j] = 0;
        } else vals[j] = code;
      }
      const cls = classify(org);
      let date = null;
      if (cfg.dateCol) {
        const d = row[cfg.dateCol];
        if (d instanceof Date) date = d.getTime();
        else if (typeof d === 'number') date = d; // Excel serial or epoch: only order matters
        else if (d != null && String(d).trim()) {
          const str = String(d).trim();
          // dd/mm/yyyy (day first, as exported by European lab systems)
          const m = str.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})(?:[ T](\d{1,2}):(\d{2}))?/);
          if (m) {
            const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
            date = Date.UTC(y, +m[2] - 1, +m[1], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
          } else {
            const t = Date.parse(str);
            date = isNaN(t) ? null : t;
          }
        }
      }
      list.push({
        idx: i,
        organism: org,
        species: speciesOf(org),
        genus: genusOf(org),
        category: cls.category,
        matchKey: cls.key,
        patient: cfg.patientCol ? String(row[cfg.patientCol] == null ? '' : row[cfg.patientCol]).trim() : null,
        date,
        vals,
      });
    });

    let dedupRemoved = 0;
    if (cfg.patientCol && cfg.dedup && cfg.dedup !== 'none') {
      const sorted = list.slice().sort((a, b) => {
        const da = a.date == null ? Infinity : a.date;
        const db = b.date == null ? Infinity : b.date;
        return da - db || a.idx - b.idx;
      });
      const seen = new Set();
      const keep = new Set();
      for (const it of sorted) {
        if (!it.patient) { keep.add(it); continue; }
        const key = cfg.dedup === 'patient' ? it.patient : it.patient + '\u0000' + normName(it.species);
        if (seen.has(key)) continue;
        seen.add(key); keep.add(it);
      }
      dedupRemoved = list.length - keep.size;
      list = list.filter((it) => keep.has(it));
    }

    const unclassified = new Map();
    for (const it of list) if (it.category === 'unclassified') unclassified.set(it.organism, (unclassified.get(it.organism) || 0) + 1);

    return {
      agents,
      isolates: list,
      stats: {
        rowsRead: rows.length,
        emptyOrganism: emptyOrg,
        excluded,
        dedupRemoved,
        analysed: list.length,
        invalidValues: Array.from(invalid.entries()).sort((a, b) => b[1] - a[1]),
        invalidCells,
        unclassified: Array.from(unclassified.entries()).sort((a, b) => b[1] - a[1]),
      },
    };
  }

  /* ---------- subsets ---------- */

  function subsetIsolates(isolates, sub) {
    switch (sub.type) {
      case 'all': return isolates;
      case 'gram': return isolates.filter((i) => i.category === sub.value);
      case 'genus': return isolates.filter((i) => i.genus === sub.value);
      case 'species': return isolates.filter((i) => i.species === sub.value);
      default: return isolates;
    }
  }

  /* ---------- analysis ---------- */

  function testedFraction(isolates, j) {
    if (!isolates.length) return 0;
    let t = 0;
    for (const it of isolates) if (it.vals[j] !== 0) t++;
    return t / isolates.length;
  }

  /*
   * opts: { threshold (0-100), conf (0-1), combDenominator: 'either'|'both' }
   * returns { n, kept: [agentIndex], dropped: [{agent, tested}], singles, combos }
   */
  function analyze(isolates, agents, opts) {
    const n = isolates.length;
    const thr = opts.threshold / 100;
    const kept = [], dropped = [];
    agents.forEach((a, j) => {
      const tf = testedFraction(isolates, j);
      if (n > 0 && tf >= thr - 1e-12) kept.push(j); else dropped.push({ agent: a, tested: tf });
    });

    const singles = kept.map((j) => {
      let s = 0, r = 0;
      for (const it of isolates) { const v = it.vals[j]; if (v === 1) s++; else if (v === 2) r++; }
      const tested = s + r;
      const [lo, hi] = jeffreys(s, r, opts.conf);
      const [glo, ghi] = jeffreys(s, n - s, opts.conf);
      return {
        agent: agents[j], j, s, r, nt: n - tested, tested, n,
        pS: tested ? s / tested : NaN, lo, hi,
        testedPct: n ? tested / n : NaN,
        pG: n ? s / n : NaN, glo, ghi,
      };
    }).sort((a, b) => (b.pS || 0) - (a.pS || 0) || b.tested - a.tested);

    const singleByJ = new Map(singles.map((x) => [x.j, x]));
    const combos = [];
    for (let x = 0; x < kept.length; x++) {
      for (let y = x + 1; y < kept.length; y++) {
        const a = kept[x], b = kept[y];
        let s = 0, tested = 0, bothR = 0, rN = 0, bothTested = 0;
        // paired gains vs each component, on isolates in that component's denominator
        let gainA = 0, gainB = 0;
        for (const it of isolates) {
          const va = it.vals[a], vb = it.vals[b];
          const anyS = va === 1 || vb === 1;
          const inDen = opts.combDenominator === 'both' ? (va !== 0 && vb !== 0) : (va !== 0 || vb !== 0);
          if (va !== 0 && vb !== 0) bothTested++;
          if (inDen) {
            tested++;
            if (anyS) s++;
            else if (va === 2 && vb === 2) bothR++;
            else rN++;
          }
          if (va === 2 && vb === 1) gainA++;
          if (vb === 2 && va === 1) gainB++;
        }
        const f = tested - s;
        const [lo, hi] = jeffreys(s, f, opts.conf);
        let sG = 0;
        for (const it of isolates) if (it.vals[a] === 1 || it.vals[b] === 1) sG++;
        const [glo, ghi] = jeffreys(sG, n - sG, opts.conf);
        const A = singleByJ.get(a), B = singleByJ.get(b);
        combos.push({
          a: agents[a], b: agents[b], ja: a, jb: b,
          s, f, tested, bothR, rN, bothTested, n,
          pS: tested ? s / tested : NaN, lo, hi,
          testedPct: n ? tested / n : NaN,
          sG, pG: n ? sG / n : NaN, glo, ghi,
          // vs A: among isolates tested for A, combination S where A is S or (A is R and B is S)
          gainA, pA: A.tested ? (A.s + gainA) / A.tested : NaN, deltaA: A.tested ? gainA / A.tested : NaN, pValA: mcnemarExact(gainA, 0),
          gainB, pB: B.tested ? (B.s + gainB) / B.tested : NaN, deltaB: B.tested ? gainB / B.tested : NaN, pValB: mcnemarExact(gainB, 0),
        });
      }
    }
    combos.sort((p, q) => (q.pS || 0) - (p.pS || 0) || q.tested - p.tested);
    return { n, kept: kept.map((j) => agents[j]), dropped, singles, combos };
  }

  // Organism x agent grid for the top organisms in a set
  function antibiogram(isolates, agents, opts) {
    const groups = new Map();
    const keyFn = opts.by === 'genus' ? (i) => i.genus : (i) => i.species;
    for (const it of isolates) {
      const k = keyFn(it);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(it);
    }
    const rows = Array.from(groups.entries())
      .map(([name, list]) => ({ name, list, n: list.length, category: list[0].category }))
      .filter((g) => g.n >= (opts.minIsolates || 1))
      .sort((a, b) => b.n - a.n)
      .slice(0, opts.maxRows || 20);
    return rows.map((g) => ({
      name: g.name, n: g.n, category: g.category,
      cells: agents.map((a, j) => {
        let s = 0, r = 0;
        for (const it of g.list) { if (it.vals[j] === 1) s++; else if (it.vals[j] === 2) r++; }
        return { s, r, tested: s + r, pS: s + r ? s / (s + r) : NaN };
      }),
    }));
  }

  function summary(isolates) {
    const cat = new Map(), org = new Map();
    for (const it of isolates) {
      cat.set(it.category, (cat.get(it.category) || 0) + 1);
      org.set(it.species, (org.get(it.species) || 0) + 1);
    }
    return {
      n: isolates.length,
      categories: Array.from(cat.entries()).sort((a, b) => b[1] - a[1]),
      organisms: Array.from(org.entries()).sort((a, b) => b[1] - a[1]),
    };
  }

  /* ---------- redundant combinations ---------- */

  const BETA_LACTAMS = new Set(['Penicillins', 'Cephalosporins', 'Carbapenems', 'Monobactams']);
  const ANTI_GRAMPOS = new Set(['Glycopeptides', 'Lipopeptides', 'Oxazolidinones']);
  const has = (s, re) => re.test(String(s).toLowerCase());
  // Clinically used pairs that the class rules would otherwise remove. Matching is on name stems (English and Italian).
  const COMBO_EXCEPTIONS = [
    { id: 'ampCro', label: 'Ampicillin + ceftriaxone or cefotaxime', note: 'E. faecalis endocarditis',
      doi: '10.1093/cid/cit052', test: (a, b) => has(a, /ampic/) && !has(a, /sulbact/) && has(b, /ceftriax|cefotax/) },
    { id: 'czaAtm', label: 'Ceftazidime-avibactam + aztreonam', note: 'metallo-β-lactamase producers',
      doi: '10.1093/cid/ciaa586', test: (a, b) => has(a, /avibact/) && has(a, /ceftaz/) && has(b, /aztreon/) && !has(b, /avibact/) },
    { id: 'czErt', label: 'Cefazolin + ertapenem', note: 'persistent MSSA bacteraemia, salvage',
      doi: '10.1093/cid/ciz995', test: (a, b) => has(a, /cefazol/) && has(b, /ertapen/) },
    { id: 'dblCarb', label: 'Ertapenem + meropenem, doripenem or imipenem', note: 'KPC salvage; limited in vitro benefit',
      doi: '10.1007/s10096-017-2936-5', test: (a, b) => has(a, /ertapen/) && has(b, /meropen|doripen|imipen/) },
  ];

  function userPairMatch(line, a, b) {
    const parts = String(line).split('+').map((x) => normName(x)).filter(Boolean);
    if (parts.length !== 2) return false;
    const na = normName(a), nb = normName(b);
    return (na.includes(parts[0]) && nb.includes(parts[1])) || (na.includes(parts[1]) && nb.includes(parts[0]));
  }

  /*
   * rules: { sameClass, betaLactam, antiGramPos: bool, exceptions: { id: bool }, keep: string[] ("A + B") }
   * returns null when the pair is kept, else { rule, reason }
   */
  function redundancy(a, clsA, b, clsB, rules) {
    for (const line of rules.keep || []) if (userPairMatch(line, a, b)) return null;
    for (const ex of COMBO_EXCEPTIONS) {
      if (rules.exceptions && rules.exceptions[ex.id] && (ex.test(a, b) || ex.test(b, a))) return null;
    }
    if (rules.sameClass && clsA === clsB && clsA !== 'Other') return { rule: 'sameClass', reason: `Two ${clsA.toLowerCase()}` };
    if (rules.betaLactam && BETA_LACTAMS.has(clsA) && BETA_LACTAMS.has(clsB)) return { rule: 'betaLactam', reason: `Two β-lactams (${clsA.toLowerCase()} + ${clsB.toLowerCase()})` };
    if (rules.antiGramPos && ANTI_GRAMPOS.has(clsA) && ANTI_GRAMPOS.has(clsB)) return { rule: 'antiGramPos', reason: `Overlapping anti-Gram-positive agents (${clsA.toLowerCase()} + ${clsB.toLowerCase()})` };
    return null;
  }

  const api = {
    redundancy, COMBO_EXCEPTIONS,
    lgamma, ibeta, betaPpf, jeffreys, mcnemarExact, normalizeValue,
    buildClassifier, genusOf, speciesOf, guessClass, CLASS_ORDER, DISPLAY_CLASS_ORDER,
    prepare, subsetIsolates, analyze, antibiogram, summary, testedFraction, normName,
  };
  root.SIEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
