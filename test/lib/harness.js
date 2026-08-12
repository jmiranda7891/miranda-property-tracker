// Shared test harness for the Miranda Property Tracker client.
//
// The point of this file: browser suites run the REAL Styles.html + JavaScript.html that
// get deployed, with `google.script.run` swapped for an in-page mock. A test failure means
// the shipping code is broken, not a copy of it. Ported from the CL Social Media App repo's
// test/lib/harness.js (same mechanism, unchanged).
const fs = require('fs');
const path = require('path');

const REPO = path.resolve(__dirname, '..', '..');

// Chromium: prefer the image's pre-installed browser (Claude Code on the web sets
// PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers), otherwise let playwright find its own.
function launchOpts() {
  const preinstalled = '/opt/pw-browsers/chromium';
  return fs.existsSync(preinstalled) ? { executablePath: preinstalled } : {};
}

function readCss() {
  return fs.readFileSync(path.join(REPO, 'Styles.html'), 'utf8').replace(/<\/?style>/g, '');
}
function readClientJs() {
  const html = fs.readFileSync(path.join(REPO, 'JavaScript.html'), 'utf8');
  const m = html.match(/<script>([\s\S]*)<\/script>/);
  if (!m) throw new Error('Could not find the <script> block in JavaScript.html');
  return m[1];
}

// Pull a single named function out of the client source, so a suite can unit-test one
// helper without booting the whole app.
function extractFn(name) {
  const js = readClientJs();
  const start = js.indexOf('function ' + name + '(');
  if (start < 0) throw new Error('Function not found in JavaScript.html: ' + name);
  let depth = 0;
  for (let i = js.indexOf('{', start); i < js.length; i++) {
    if (js[i] === '{') depth++;
    else if (js[i] === '}') { depth--; if (!depth) return js.slice(start, i + 1); }
  }
  throw new Error('Unbalanced braces reading: ' + name);
}

// Builds a full page: real CSS, a mock google.script.run backed by `apiSource`
// (a string of JS defining `var API = {...}`), the real client code, then `bootJs`.
function buildPage(opts) {
  const { apiSource, bootJs = '', hangFns = [], preludeJs = '' } = opts;
  const hang = JSON.stringify(hangFns);
  return `<!doctype html><html><head><meta charset="utf-8"><style>${readCss()}</style></head>
<body><div id="app"></div>
<script>
var CALLS = [];
var HANG_FNS = ${hang};
${apiSource}
window.google = { script: { run: (function () {
  function mk(ok, bad) {
    var o = {};
    Object.keys(API).forEach(function (fn) {
      o[fn] = function () {
        var args = Array.prototype.slice.call(arguments);
        CALLS.push(fn);
        if (HANG_FNS.indexOf(fn) >= 0) return;
        setTimeout(function () {
          try { var r = API[fn].apply(null, args); if (ok) ok(r); }
          catch (e) { if (bad) bad(e); }
        }, 5);
      };
    });
    o.withSuccessHandler = function (f) { return mk(f, bad); };
    o.withFailureHandler = function (f) { return mk(ok, f); };
    return o;
  }
  return mk(null, null);
})() } };
${preludeJs}
${readClientJs()}
${bootJs}
<\/script></body></html>`;
}

function writePage(name, html) {
  const dir = path.join(REPO, 'test', '.tmp');
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, name);
  fs.writeFileSync(file, html);
  return 'file://' + file;
}

// Tiny assertion recorder. Suites export run(t) and call t.check(...).
function makeRecorder(label) {
  const state = { failures: 0, total: 0 };
  return {
    group(name) { console.log('  ' + name); },
    check(name, cond, detail) {
      state.total++;
      if (cond) { console.log('    PASS  ' + name); }
      else {
        state.failures++;
        console.log('    FAIL  ' + name + (detail ? '  -> ' + detail : ''));
      }
    },
    watch(page) {
      page.on('pageerror', (e) => {
        state.failures++; state.total++;
        console.log('    FAIL  [page error] ' + e.message);
      });
    },
    result() { return state; },
    label
  };
}

module.exports = { REPO, launchOpts, readCss, readClientJs, extractFn, buildPage, writePage, makeRecorder };
