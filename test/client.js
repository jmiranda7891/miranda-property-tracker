// Client-side UI tests. Runs the REAL Styles.html + JavaScript.html in Chromium with a
// scripted mock of google.script.run (see test/lib/harness.js), so a failure here means the
// shipping client code is broken, not a copy of it.
const { chromium } = require('playwright');
const { launchOpts, buildPage, writePage, makeRecorder } = require('./lib/harness');

const SAMPLE_PROPS = [
  {
    id: 'p1', countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Casa Test',
    propietario: 'Owner A', direccion: 'Calle 1', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250',
    observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 100000, precioEstimadoIsPlaceholder: false,
    escrituras: 'Si', propEscriturado: 'Owner A', propuestaTraspaso: '', pin: '',
    aiResearchEN: 'English research text for Casa Test.', aiResearchES: 'Texto de investigación en español para Casa Test.',
    aiValueEstimateEN: 'English value estimate.', aiValueEstimateES: 'Estimación de valor en español.',
    researchDate: '', archived: false, archivedReason: '', createdBy: 'a@x.com', createdAt: '', updatedBy: 'a@x.com', updatedAt: ''
  },
  {
    id: 'p2', countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Edificio', referencia: 'US Test',
    propietario: 'Owner B', direccion: '1 Main St', ciudad: 'Chicago', estado: 'Illinois', cp: '60654',
    observacionesRaw: '', status: 'for_sale', precioEstimadoUSD: 500000, precioEstimadoIsPlaceholder: false,
    escrituras: '', propEscriturado: '', propuestaTraspaso: '', pin: '17-00-000-000-0000',
    aiResearchEN: 'English research text for US Test.', aiResearchES: 'Texto de investigación en español para US Test.',
    aiValueEstimateEN: 'US value estimate.', aiValueEstimateES: 'Estimación de valor en EE. UU.',
    researchDate: '', archived: false, archivedReason: '', createdBy: 'a@x.com', createdAt: '', updatedBy: 'a@x.com', updatedAt: ''
  }
];

function apiSourceFor(role) {
  return `
    var SAMPLE_PROPS = ${JSON.stringify(SAMPLE_PROPS)};
    var API = {
      getBootstrap: function () {
        if ('${role}' === 'none') return { email: 'stranger@example.com', role: 'none', build: 'v1' };
        return { email: 'admin@example.com', role: '${role}', build: 'v1', properties: SAMPLE_PROPS };
      },
      listUsers: function () { return [{ email: 'admin@example.com', role: 'admin', addedBy: 'system', addedAt: '' }]; },
      getAuditLog: function () { return []; },
      exportProperties: function (json) { return { url: 'https://fake.export.test/sheet', count: SAMPLE_PROPS.length }; },
      addProperty: function (form) { return SAMPLE_PROPS; },
      updateProperty: function (id, form) { return SAMPLE_PROPS; },
      archiveProperty: function (id, reason) { return SAMPLE_PROPS; },
      unarchiveProperty: function (id) { return SAMPLE_PROPS; },
      deleteProperty: function (id, ref) { return SAMPLE_PROPS; },
      saveUser: function (email, role) { return [{ email: 'admin@example.com', role: 'admin' }]; },
      removeUser: function (email) { return [{ email: 'admin@example.com', role: 'admin' }]; }
    };
  `;
}

module.exports = async function run(t) {
  const browser = await chromium.launch(launchOpts());

  t.group('access control screen');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-none.html', buildPage({ apiSource: apiSourceFor('none') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    const cardCount = await page.locator('.prop-card').count();
    const hasCenterMsg = await page.locator('.center-msg').count();
    t.check('an unauthorized role sees no property cards', cardCount === 0);
    t.check('an unauthorized role sees the access-denied message', hasCenterMsg > 0);
    await page.close();
  }

  t.group('dashboard renders and filters');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-admin.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    const cardCount = await page.locator('.prop-card').count();
    t.check('both sample properties render as cards', cardCount === 2, 'got ' + cardCount);

    await page.selectOption('.filter-bar select >> nth=0', 'MX');
    await page.waitForTimeout(50);
    const filteredCount = await page.locator('.prop-card').count();
    t.check('filtering by country narrows the list', filteredCount === 1, 'got ' + filteredCount);
    await page.selectOption('.filter-bar select >> nth=0', '');
    await page.waitForTimeout(50);
    await page.close();
  }

  t.group('list view + sorting');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-list.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    t.check('cards view is the default', (await page.locator('.prop-card').count()) === 2 && (await page.locator('table.prop-table').count()) === 0);

    await page.click('button:has-text("Lista"), button:has-text("List")');
    await page.waitForTimeout(50);
    t.check('switching to List view shows a table row per filtered property', (await page.locator('table.prop-table tbody tr').count()) === 2);
    t.check('switching to List view hides the card grid', (await page.locator('.prop-card').count()) === 0);

    const refHeader = 'table.prop-table th:has-text("Referencia"), table.prop-table th:has-text("Reference")';
    const firstCellText = async () => (await page.locator('table.prop-table tbody tr').first().locator('td').first().textContent()).trim();
    await page.click(refHeader);
    await page.waitForTimeout(50);
    const ascFirst = await firstCellText();
    await page.click(refHeader);
    await page.waitForTimeout(50);
    const descFirst = await firstCellText();
    t.check('clicking a sortable column header once, then again, reverses row order', ascFirst !== descFirst, ascFirst + ' / ' + descFirst);
    await page.click(refHeader);
    await page.waitForTimeout(50);
    const backToAscFirst = await firstCellText();
    t.check('a third click returns to the original ascending order', backToAscFirst === ascFirst, ascFirst + ' / ' + backToAscFirst);

    await page.click('button:has-text("Tarjetas"), button:has-text("Cards")');
    await page.waitForTimeout(50);
    t.check('switching back to Cards view restores the card grid', (await page.locator('.prop-card').count()) === 2);
    await page.close();
  }

  t.group('card polish');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-cards.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    const borderColors = await page.locator('.prop-card').evaluateAll((els) => els.map((el) => getComputedStyle(el).borderLeftColor));
    t.check('cards with different statuses get different accent border colors', borderColors[0] !== borderColors[1], borderColors.join(' / '));
    const mapsHref = await page.locator('.card-maps-link').first().getAttribute('href');
    t.check('each card has a working Maps quick-link', mapsHref && mapsHref.indexOf('google.com/maps') >= 0 && mapsHref.indexOf('Maravatio') >= 0, mapsHref);
    await page.close();
  }

  t.group('EN/ES toggle: chrome switches, raw data does not');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-lang.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    const titleEs = await page.locator('.topbar h1').textContent();
    await page.click('.lang-toggle button:has-text("EN")');
    await page.waitForTimeout(50);
    const titleEn = await page.locator('.topbar h1').textContent();
    t.check('the app title text changes between ES and EN', titleEs !== titleEn, titleEs + ' / ' + titleEn);

    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    const researchEn = await page.locator('.field-group .prose').first().textContent();
    t.check('detail shows the English research field after switching to EN', researchEn.indexOf('English research text') >= 0, researchEn);
    const ownerText = await page.locator('.kv:has-text("Owner(s)") .v, .kv:has-text("Propietario(s)") .v').first().textContent();
    t.check('the raw owner field is never translated', ownerText.trim() === 'Owner A', ownerText);

    await page.click('.modal-close');
    await page.click('.lang-toggle button:has-text("ES")');
    await page.waitForTimeout(50);
    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    const researchEs = await page.locator('.field-group .prose').first().textContent();
    t.check('detail shows the Spanish research field after switching back to ES', researchEs.indexOf('investigación en español') >= 0, researchEs);
    await page.close();
  }

  t.group('admin-only actions are hidden from a member');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-member.html', buildPage({ apiSource: apiSourceFor('member') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    const adminButtonCount = await page.locator('button:has-text("Admin"), button:has-text("⚙")').count();
    t.check('a member does not see the Admin button', adminButtonCount === 0);
    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    const archiveButtonCount = await page.locator('.modal-footer button:has-text("Archivar"), .modal-footer button:has-text("Archive")').count();
    t.check('a member does not see Archive/Delete in the detail modal', archiveButtonCount === 0);
    await page.close();
  }

  await browser.close();
};
