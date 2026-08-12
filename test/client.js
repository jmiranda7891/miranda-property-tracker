// Client-side UI tests. Runs the REAL Styles.html + JavaScript.html in Chromium with a
// scripted mock of google.script.run (see test/lib/harness.js), so a failure here means the
// shipping client code is broken, not a copy of it.
const { chromium } = require('playwright');
const { launchOpts, buildPage, writePage, makeRecorder } = require('./lib/harness');

const SAMPLE_PROPS = [
  {
    id: 'p1', countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Casa Test',
    propietario: 'Owner A', direccion: 'Calle 1', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250',
    observacionesRaw: 'En Uso', status: 'en_uso', precioEstimadoUSD: 100000, precioEstimadoIsPlaceholder: false,
    escrituras: 'Si', esEjido: false, propEscriturado: 'Titled Name A', planLargoPlazo: 'fideicomiso', pin: '',
    lat: 19.8942, lng: -100.4436,
    aiResearchEN: 'English research text for Casa Test.', aiResearchES: 'Texto de investigación en español para Casa Test.',
    aiValueEstimateEN: 'English value estimate.', aiValueEstimateES: 'Estimación de valor en español.',
    researchDate: '', archived: false, archivedReason: '', createdBy: 'a@x.com', createdAt: '', updatedBy: 'a@x.com', updatedAt: ''
  },
  {
    id: 'p2', countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Edificio', referencia: 'US Test',
    propietario: 'Owner B', direccion: '1 Main St', ciudad: 'Chicago', estado: 'Illinois', cp: '60654',
    observacionesRaw: '', status: 'en_venta', precioEstimadoUSD: 500000, precioEstimadoIsPlaceholder: false,
    escrituras: 'No', esEjido: false, propEscriturado: '', planLargoPlazo: 'mantener_individual', pin: '17-00-000-000-0000',
    lat: 41.8781, lng: -87.6298,
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

// A minimal fake of the two Leaflet calls initMap() makes, so the map view's OWN logic (right
// container id, one marker per property with coordinates, a working popup link back to
// openDetail) is verified without depending on live network access to the Leaflet CDN in the
// test sandbox - the real CDN load is a manual-smoke-test item (see CLAUDE.md), not this suite.
const FAKE_LEAFLET_PRELUDE = `
  window.__mapCalls = [];
  window.__markers = [];
  window.L = {
    map: function (id) {
      window.__mapCalls.push(id);
      var inst = { setView: function () { return inst; }, remove: function () {}, fitBounds: function () { return inst; } };
      return inst;
    },
    tileLayer: function () { return { addTo: function () { return this; } }; },
    marker: function (latlng) {
      var m = { latlng: latlng, addTo: function () { window.__markers.push(m); return m; }, bindPopup: function (html) { m.popup = html; return m; } };
      return m;
    }
  };
`;

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
    const ownerText = await page.locator('.kv:has-text("Beneficial Owner") .v, .kv:has-text("Dueño Beneficiario") .v').first().textContent();
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

  t.group('estate planning fields (add/edit form)');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-form.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);

    // Detail view shows the deed status as Sí/No text and the ejido flag, not a raw enum key.
    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    const deedText = await page.locator('.kv:has-text("Escrituras"), .kv:has-text("Deed status")').first().locator('.v').textContent();
    t.check('deed status renders as a yes/no word, not a raw code', ['Sí', 'Yes', 'No'].includes(deedText.trim()), deedText);
    await page.click('.modal-close');

    await page.click('button:has-text("Agregar propiedad"), button:has-text("Add property")');
    await page.waitForTimeout(50);
    const escrituraOptions = await page.locator('#f_escrituras option').allTextContents();
    t.check('Escrituras is a two-option Si/No select, not free text', escrituraOptions.length === 2);
    const planOptions = await page.locator('#f_planLargoPlazo option').allTextContents();
    t.check('Plan a Largo Plazo has exactly 3 options', planOptions.length === 3, planOptions.join(','));
    t.check('an Es Ejido checkbox exists', (await page.locator('#f_esEjido').count()) === 1);

    const datalistOptions = await page.locator('#dl_names option').evaluateAll((els) => els.map((e) => e.value));
    t.check('the name datalist is populated from existing properties', datalistOptions.includes('Owner A') && datalistOptions.includes('Titled Name A'), datalistOptions.join(','));
    t.check('both name fields point at the shared datalist', (await page.locator('#f_propietario').getAttribute('list')) === 'dl_names'
      && (await page.locator('#f_propEscriturado').getAttribute('list')) === 'dl_names');
    await page.close();
  }

  t.group('map view (fake Leaflet - verifies initMap()\'s own logic, not the real CDN)');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-map.html', buildPage({ apiSource: apiSourceFor('admin'), preludeJs: FAKE_LEAFLET_PRELUDE }));
    await page.goto(url);
    await page.waitForTimeout(150);
    await page.click('button:has-text("Mapa"), button:has-text("Map")');
    await page.waitForTimeout(100);
    const mapCalls = await page.evaluate(() => window.__mapCalls);
    t.check('switching to Map view initializes Leaflet on #mapContainer', mapCalls.includes('mapContainer'), mapCalls.join(','));
    const markers = await page.evaluate(() => window.__markers.map((m) => ({ latlng: m.latlng, popup: m.popup })));
    t.check('one marker per property with coordinates', markers.length === 2, JSON.stringify(markers.map((m) => m.latlng)));
    t.check('a marker popup links back to the same detail modal', markers.some((m) => m.popup.includes('Casa Test') && m.popup.includes("MPT.openDetail('p1')")), JSON.stringify(markers));

    // The actual bug: Leaflet assigns its internal panes/controls z-index up to 1000, and
    // .leaflet-container doesn't establish its own stacking context by default, so those
    // values used to leak above .modal-backdrop (z-index 50) and paint the map over an open
    // modal. The fix is .map-container getting its OWN stacking context (position+z-index) so
    // Leaflet's internal z-index values stay contained regardless of how high they are.
    const mapContainerStyle = await page.locator('#mapContainer').evaluate((el) => {
      const cs = getComputedStyle(el);
      return { position: cs.position, zIndex: cs.zIndex };
    });
    t.check('#mapContainer establishes its own stacking context (position set, z-index not auto)',
      mapContainerStyle.position !== 'static' && mapContainerStyle.zIndex !== 'auto', JSON.stringify(mapContainerStyle));

    // Fake Leaflet renders no real clickable pins, so open the detail modal directly (this is
    // exactly what a real pin's popup link does under the hood) while Map view stays active.
    await page.evaluate(() => window.MPT.openDetail('p1'));
    await page.waitForTimeout(50);
    const stackOrder = await page.evaluate(() => {
      const modal = document.querySelector('.modal-backdrop');
      const map = document.getElementById('mapContainer');
      if (!modal || !map) return null;
      const modalZ = parseInt(getComputedStyle(modal).zIndex, 10) || 0;
      const mapZ = parseInt(getComputedStyle(map).zIndex, 10) || 0;
      return { modalZ, mapZ };
    });
    t.check('the modal backdrop\'s z-index is higher than the (now self-contained) map container\'s',
      stackOrder && stackOrder.modalZ > stackOrder.mapZ, JSON.stringify(stackOrder));
    await page.close();
  }

  t.group('lat/lng are editable (map pin self-service fix)');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-latlng.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    await page.click('button:has-text("Editar"), button:has-text("Edit")');
    await page.waitForTimeout(50);
    const latVal = await page.locator('#f_lat').inputValue();
    const lngVal = await page.locator('#f_lng').inputValue();
    t.check('the edit form prefills the existing lat/lng', parseFloat(latVal) === 19.8942 && parseFloat(lngVal) === -100.4436, latVal + ',' + lngVal);
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
