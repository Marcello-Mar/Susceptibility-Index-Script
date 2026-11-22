const XLSX = require('../vendor/xlsx.full.min.js');
const E = require('../js/engine.js');
const wb = XLSX.read(require('fs').readFileSync(process.argv[2]), { cellDates: true });
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
const cols = Object.keys(rows[0]);
const prep = E.prepare(rows, { orgCol: 'Microorganism', agentCols: cols.slice(1), iMode: 'S', dedup: 'none', gramMap: [] });
const out = {};
for (const conf of [0.95, 0.90]) for (const thr of [0, 50]) {
  const r = E.analyze(prep.isolates, prep.agents, { threshold: thr, conf, combDenominator: 'either' });
  const p = (x) => x * 100;
  out[`${conf}-${thr}`] = {
    singles: Object.fromEntries(r.singles.map((s) => [s.agent, [p(s.pS || 0), p(s.lo || 0), p(s.hi || 0), p(s.pG), p(s.glo), p(s.ghi)]])),
    combos: Object.fromEntries(r.combos.map((c) => [`${c.a}|${c.b}`, [p(c.pS || 0), p(c.lo || 0), p(c.hi || 0), p(c.pG), p(c.glo), p(c.ghi)]])),
  };
}
console.log(JSON.stringify(out));
