// Static safety checks — no browser needed. Catches the two mistakes that break an Apps
// Script deploy in ways that are painful to debug from @HEAD:
//   1. a syntax error in Code.js or in the served client JS (blank page, cryptic error)
//   2. a /* */ sequence anywhere in a file served through HtmlService — HtmlService's
//      sanitizer can corrupt these even outside real comments, so the served files simply
//      never contain the two-character sequence at all (see CLAUDE.md "Apps Script gotchas").
// Ported from the CL Social Media App repo's test/static.js, adapted for this app's files
// and oauth scope needs.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { REPO } = require('./lib/harness');

function scriptBlock(file) {
  const html = fs.readFileSync(path.join(REPO, file), 'utf8');
  const m = html.match(/<script>([\s\S]*)<\/script>/);
  return m ? m[1] : '';
}

module.exports = async function run(t) {
  t.group('parses');
  for (const f of ['Code.js']) {
    let ok = true, err = '';
    try { new vm.Script(fs.readFileSync(path.join(REPO, f), 'utf8'), { filename: f }); }
    catch (e) { ok = false; err = e.message; }
    t.check(f + ' parses', ok, err);
  }
  let ok = true, err = '';
  try { new vm.Script(scriptBlock('JavaScript.html'), { filename: 'JavaScript.html' }); }
  catch (e) { ok = false; err = e.message; }
  t.check('JavaScript.html script block parses', ok, err);

  t.group('HtmlService comment-stripping hazard');
  for (const f of ['JavaScript.html', 'Index.html', 'Styles.html']) {
    const body = fs.readFileSync(path.join(REPO, f), 'utf8');
    const bad = body.split('\n')
      .map((l, i) => ({ n: i + 1, l }))
      .filter((x) => x.l.includes('/*') || x.l.includes('*/'));
    t.check(f + ': no /* */ anywhere in the file', bad.length === 0,
      bad.slice(0, 3).map((x) => 'line ' + x.n).join(', '));
  }

  t.group('deploy surface');
  const ignore = fs.readFileSync(path.join(REPO, '.claspignore'), 'utf8');
  t.check('.claspignore is still a whitelist (nothing else can ship)', /^\*\*\/\*\*/m.test(ignore));
  for (const f of ['appsscript.json', 'Code.js', 'Index.html', 'JavaScript.html', 'Styles.html']) {
    t.check('whitelisted: ' + f, ignore.includes('!' + f));
  }

  const manifest = JSON.parse(fs.readFileSync(path.join(REPO, 'appsscript.json'), 'utf8'));
  t.check('webapp access is DOMAIN (outer gate, app-level allow-list is the inner gate)',
    manifest.webapp && manifest.webapp.access === 'DOMAIN');
  t.check('manifest grants spreadsheets scope', manifest.oauthScopes.some((s) => s.includes('spreadsheets')));
  t.check('manifest grants userinfo.email scope (Session.getActiveUser)', manifest.oauthScopes.some((s) => s.includes('userinfo.email')));

  const code = fs.readFileSync(path.join(REPO, 'Code.js'), 'utf8');
  const scopeFor = [
    ['ScriptApp.newTrigger', 'script.scriptapp'],
    ['UrlFetchApp.fetch', 'script.external_request'],
    ['MailApp.sendEmail', 'script.send_mail'],
    ['DriveApp.', 'auth/drive'],
    ['DocumentApp.', 'auth/documents']
  ];
  scopeFor.forEach(([api, scope]) => {
    t.check('v1 does not call ' + api + ' (no unneeded scope)', code.indexOf(api) < 0);
  });
  t.check('SpreadsheetApp.create is used (export + registry) so drive.file scope is granted',
    code.indexOf('SpreadsheetApp.create') >= 0 && manifest.oauthScopes.some((s) => s.includes('drive.file')));

  t.check('.clasprc.json is gitignored (live OAuth token)',
    fs.readFileSync(path.join(REPO, '.gitignore'), 'utf8').includes('.clasprc.json'));

  t.group('seed data');
  const seedMatch = code.match(/var SEED_PROPERTIES_ = (\[[\s\S]*?\n\]);/);
  t.check('SEED_PROPERTIES_ array literal found', !!seedMatch);
  if (seedMatch) {
    let seed = [];
    let ok2 = true, err2 = '';
    try { seed = new Function('return ' + seedMatch[1])(); } catch (e) { ok2 = false; err2 = e.message; }
    t.check('SEED_PROPERTIES_ evaluates as a plain array', ok2, err2);
    t.check('SEED_PROPERTIES_ has 41 rows (34 MX + 7 USA)', seed.length === 41, 'got ' + seed.length);
    t.check('every seed row has a referencia and direccion', seed.every((p) => p.referencia && p.direccion));
    t.check('every seed row has bilingual research text', seed.every((p) => p.aiResearchEN && p.aiResearchES));
    const mx = seed.filter((p) => p.countryCode === 'MX').length;
    const us = seed.filter((p) => p.countryCode === 'US').length;
    t.check('34 MX rows', mx === 34, 'got ' + mx);
    t.check('7 US rows', us === 7, 'got ' + us);
    const willow = seed.filter((p) => p.referencia === 'Casa Willow')[0];
    t.check('the sold Chicago house is pre-archived in the seed', !!willow && willow.archived === true);

    t.group('seed data - schema v2 (estate-planning rework)');
    const VALID_STATUSES = ['en_uso', 'libre', 'en_venta', 'vendida'];
    const VALID_PLANES = ['mantener_individual', 'fideicomiso', 'vender'];
    t.check('every seed row has one of the 4 new status values', seed.every((p) => VALID_STATUSES.includes(p.status)),
      [...new Set(seed.map((p) => p.status))].filter((s) => !VALID_STATUSES.includes(s)).join(','));
    t.check('every seed row has escrituras Si or No (never Ejido/Proceso/etc)', seed.every((p) => p.escrituras === 'Si' || p.escrituras === 'No'),
      [...new Set(seed.map((p) => p.escrituras))].filter((s) => s !== 'Si' && s !== 'No').join(','));
    t.check('every seed row has an explicit boolean esEjido', seed.every((p) => typeof p.esEjido === 'boolean'));
    t.check('every seed row has one of the 3 planLargoPlazo values', seed.every((p) => VALID_PLANES.includes(p.planLargoPlazo)),
      [...new Set(seed.map((p) => p.planLargoPlazo))].filter((s) => !VALID_PLANES.includes(s)).join(','));
    t.check('every seed row has lat/lng for the map view', seed.every((p) => typeof p.lat === 'number' && typeof p.lng === 'number'));
    const ejidoCount = seed.filter((p) => p.esEjido).length;
    t.check('at least one property is flagged as ejido land', ejidoCount > 0, 'got ' + ejidoCount);
  }
};
