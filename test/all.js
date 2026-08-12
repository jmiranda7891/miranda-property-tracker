#!/usr/bin/env node
// Runs every suite. Exits non-zero on the first failing assertion anywhere, so this is
// safe to use as a deploy gate.
//
//   npm test                  # everything
//   node test/all.js server   # one suite (substring match)
const { makeRecorder } = require('./lib/harness');

const SUITES = [
  ['static', './static'], // fastest, no browser — fails loudest on a syntax error
  ['server', './server'], // real Code.js in a bare vm — RBAC, CRUD, audit, export, seed
  ['client', './client']  // real Styles.html/JavaScript.html in Chromium
];

(async () => {
  const filter = process.argv[2];
  const picked = filter ? SUITES.filter(([n]) => n.includes(filter)) : SUITES;
  if (!picked.length) {
    console.error('No suite matches "' + filter + '". Available: ' + SUITES.map(([n]) => n).join(', '));
    process.exit(2);
  }

  let failures = 0, total = 0;
  const started = Date.now();
  for (const [name, mod] of picked) {
    console.log('\n=== ' + name + ' ===');
    const t = makeRecorder(name);
    try {
      await require(mod)(t);
    } catch (e) {
      console.log('    FAIL  [suite crashed] ' + (e && e.message ? e.message : e));
      failures++; total++;
    }
    const r = t.result();
    failures += r.failures; total += r.total;
  }

  const secs = ((Date.now() - started) / 1000).toFixed(1);
  console.log('\n' + (failures
    ? failures + ' FAILED of ' + total + ' checks (' + secs + 's)'
    : 'All ' + total + ' checks passed (' + secs + 's)'));
  process.exit(failures ? 1 : 0);
})();
