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
    propietario: 'Owner B', direccion: '1 Main St', direccion2: 'Unit 501', ciudad: 'Chicago', estado: 'Illinois', cp: '60654',
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
      removeUser: function (email) { return [{ email: 'admin@example.com', role: 'admin' }]; },
      listProperties: function () { return SAMPLE_PROPS; },
      relocateAllPins: function () {
        return {
          updated: SAMPLE_PROPS.length, kept: 0,
          details: SAMPLE_PROPS.map(function (p) {
            return {
              referencia: p.referencia, method: 'address', status: 'OK', query: p.direccion + ', ' + p.ciudad,
              addressStatus: 'OK', addressQuery: p.direccion + ', ' + p.ciudad,
              before: { lat: p.lat, lng: p.lng }, after: { lat: p.lat, lng: p.lng }
            };
          })
        };
      }
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

    // Regression: render() rebuilds the whole #app innerHTML on every keystroke (needed so the
    // filtered list updates live) - without id-based focus restoration, that blurs the search
    // box after each character, so only the FIRST typed character ever lands and every
    // subsequent one needs a fresh click. page.type() sends real per-character key events,
    // unlike page.fill() (which sets the whole value in one shot and would never catch this).
    await page.click('.filter-bar input[type="text"]');
    await page.type('.filter-bar input[type="text"]', 'casa', { delay: 30 });
    await page.waitForTimeout(50);
    const searchValue = await page.locator('.filter-bar input[type="text"]').inputValue();
    t.check('typing multiple characters into the search box does not lose focus after each keystroke',
      searchValue === 'casa', searchValue);
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

  t.group('list view: per-column filters (AutoFilter-style)');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-list-filters.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    await page.click('button:has-text("Lista"), button:has-text("List")');
    await page.waitForTimeout(50);
    t.check('a filter row renders under the column headers, one control per column', (await page.locator('tr.filter-row th').count()) > 0);

    // Text filter: "Casa Test" vs "US Test" only differ in referencia.
    await page.fill('#colf_referencia', 'US');
    await page.waitForTimeout(50);
    t.check('a text column filter narrows rows by a case-insensitive contains match', (await page.locator('table.prop-table tbody tr').count()) === 1);
    t.check('the surviving row is the one that matches', (await page.locator('table.prop-table tbody tr td').first().textContent()).includes('US Test'));
    const clearBtn = page.locator('.list-filter-bar button');
    t.check('a "clear column filters" control appears once a filter is active', (await clearBtn.count()) === 1);
    await clearBtn.click();
    await page.waitForTimeout(50);
    t.check('clearing resets to all rows visible again', (await page.locator('table.prop-table tbody tr').count()) === 2);

    // Select filter: countryCode narrows to exactly the MX or US row.
    await page.selectOption('#colf_countryCode', 'US');
    await page.waitForTimeout(50);
    t.check('a select column filter (country) narrows to exactly the matching row', (await page.locator('table.prop-table tbody tr').count()) === 1);
    await page.selectOption('#colf_countryCode', '');
    await page.waitForTimeout(50);

    // Range filter: precioEstimadoUSD is 100000 (Casa Test) vs 500000 (US Test).
    await page.fill('#colf_precioEstimadoUSD_min', '200000');
    await page.waitForTimeout(50);
    t.check('a range column filter (price min) excludes rows below the minimum', (await page.locator('table.prop-table tbody tr').count()) === 1);
    const survivorText = await page.locator('table.prop-table tbody tr td').first().textContent();
    t.check('the range filter kept the higher-priced row', survivorText.includes('US Test'), survivorText);
    await page.fill('#colf_precioEstimadoUSD_min', '');
    await page.waitForTimeout(50);

    // Column filters compose with a query that would otherwise match zero rows.
    await page.fill('#colf_referencia', 'nonexistent-xyz');
    await page.waitForTimeout(50);
    t.check('a filter matching nothing shows an in-table empty message, not a broken table', (await page.locator('table.prop-table tbody tr').count()) === 1
      && (await page.locator('td.empty-filtered').count()) === 1);
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

    // US Test (p2) carries a direccion2 ("Unit 501") in the fixture - confirm the detail view's
    // Location section shows both the country and the apt/interior number as their own lines.
    await page.click('.prop-card >> nth=1');
    await page.waitForTimeout(50);
    const locationGroupText = await page.locator('.field-group:has-text("Ubicación"), .field-group:has-text("Location")').first().textContent();
    t.check('the detail view\'s Location section shows the country (EE. UU./USA)', locationGroupText.includes('EE. UU.') || locationGroupText.includes('USA'), locationGroupText);
    t.check('the detail view\'s Location section shows the apt/interior number', locationGroupText.includes('Unit 501'), locationGroupText);
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

    const countryGroupHeading = await page.locator('.field-group:has(#f_countryCode) h4').textContent();
    t.check('País/Country now lives in the Ubicación/Location field-group, not Identification',
      countryGroupHeading.includes('Ubicación') || countryGroupHeading.includes('Location'), countryGroupHeading);
    t.check('a second address field exists for the apt/interior number', (await page.locator('#f_direccion2').count()) === 1);
    await page.close();
  }

  t.group('tracking fields (lot/construction size, acquisition basis, ownership/legal, registry IDs) - v1.8');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-tracking.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);

    // Casa Test (p1) is MX + tipo Casa: both lot and construction size rows show, and the
    // registry section shown is the Mexico one, not USA - this is also the regression check
    // for the bug where "USA - Parcel" used to render unconditionally regardless of country.
    await page.click('button:has-text("Agregar propiedad"), button:has-text("Add property")');
    await page.waitForTimeout(50);
    t.check('a new property (defaults to MX) shows lot size and construction size fields', (await page.locator('#f_lotSizeValue').count()) === 1 && (await page.locator('#f_constructionSizeValue').count()) === 1);
    t.check('a new property (defaults to MX) shows the Mexico registry fields, not USA', (await page.locator('#f_folioReal').count()) === 1 && (await page.locator('#f_county').count()) === 0);

    // Switching Tipo to Terreno hides construction size (a lot has no construction); switching
    // to Departamento hides lot size (a condo unit has no exclusive lot).
    await page.selectOption('#f_tipo', 'Terreno');
    await page.waitForTimeout(50);
    t.check('switching Tipo to Terreno hides construction size but keeps lot size', (await page.locator('#f_constructionSizeValue').count()) === 0 && (await page.locator('#f_lotSizeValue').count()) === 1);
    await page.selectOption('#f_tipo', 'Departamento');
    await page.waitForTimeout(50);
    t.check('switching Tipo to Departamento hides lot size but keeps construction size', (await page.locator('#f_lotSizeValue').count()) === 0 && (await page.locator('#f_constructionSizeValue').count()) === 1);
    await page.selectOption('#f_tipo', 'Casa');
    await page.waitForTimeout(50);

    // Switching Country to US swaps the registry section live, without needing to save/reopen.
    await page.fill('#f_referencia', 'Reflow Test Ref');
    await page.selectOption('#f_countryCode', 'US');
    await page.waitForTimeout(50);
    t.check('switching Country to USA swaps in the USA registry fields', (await page.locator('#f_county').count()) === 1 && (await page.locator('#f_folioReal').count()) === 0);
    t.check('switching Country does not lose a value already typed in another field', await page.locator('#f_referencia').inputValue() === 'Reflow Test Ref');
    await page.selectOption('#f_countryCode', 'MX');
    await page.waitForTimeout(50);
    t.check('switching Country back to Mexico restores the Mexico registry fields', (await page.locator('#f_folioReal').count()) === 1 && (await page.locator('#f_county').count()) === 0);
    await page.click('button:has-text("Cancelar"), button:has-text("Cancel")');

    // US Test (p2) has no tracking-field data in the fixture, so this exercises rendering with
    // everything blank (should not throw / show "null" text) plus a documentsUrl link.
    await page.click('.prop-card >> nth=1');
    await page.waitForTimeout(50);
    await page.click('button:has-text("Editar"), button:has-text("Edit")');
    await page.waitForTimeout(50);
    t.check('editing an existing USA property shows its own USA registry fields, not Mexico\'s', (await page.locator('#f_county').count()) === 1 && (await page.locator('#f_folioReal').count()) === 0);
    await page.close();
  }

  t.group('documentsUrl link (safe scheme allowlist)');
  {
    const withDocsUrl = SAMPLE_PROPS.map((p, i) => i === 0
      ? Object.assign({}, p, { documentsUrl: 'https://drive.google.com/folder/abc' })
      : Object.assign({}, p, { documentsUrl: 'javascript:alert(1)' }));
    const docsApiSource = `
      var SAMPLE_PROPS = ${JSON.stringify(withDocsUrl)};
      var API = { getBootstrap: function () { return { email: 'admin@example.com', role: 'admin', build: 'v1', properties: SAMPLE_PROPS }; } };
    `;
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-docs-url.html', buildPage({ apiSource: docsApiSource }));
    await page.goto(url);
    await page.waitForTimeout(150);

    await page.click('.prop-card >> nth=0');
    await page.waitForTimeout(50);
    const goodHref = await page.locator('a:has-text("Documentos"), a:has-text("Documents")').getAttribute('href');
    t.check('a genuine https:// documentsUrl renders as a clickable link', goodHref === 'https://drive.google.com/folder/abc', goodHref);
    await page.click('.modal-close');

    await page.click('.prop-card >> nth=1');
    await page.waitForTimeout(50);
    const badLinkCount = await page.locator('a:has-text("Documentos"), a:has-text("Documents")').count();
    t.check('a javascript: documentsUrl is never rendered as a clickable link', badLinkCount === 0);
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
    t.check('a property with its own unique coordinate renders at the EXACT geocoded lat/lng, not nudged off by the shared-coordinate jitter',
      markers.some((m) => m.latlng[0] === SAMPLE_PROPS[0].lat && m.latlng[1] === SAMPLE_PROPS[0].lng)
      && markers.some((m) => m.latlng[0] === SAMPLE_PROPS[1].lat && m.latlng[1] === SAMPLE_PROPS[1].lng),
      JSON.stringify(markers.map((m) => m.latlng)));

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

  t.group('map view: the shared-coordinate jitter only applies to properties that actually share one point');
  {
    const SHARED_PROPS = SAMPLE_PROPS.map((p) => Object.assign({}, p, { lat: 19.7, lng: -101.2 }));
    const sharedApiSource = `
      var SAMPLE_PROPS = ${JSON.stringify(SHARED_PROPS)};
      var API = {
        getBootstrap: function () { return { email: 'admin@example.com', role: 'admin', build: 'v1', properties: SAMPLE_PROPS }; },
        listProperties: function () { return SAMPLE_PROPS; }
      };
    `;
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-map-shared.html', buildPage({ apiSource: sharedApiSource, preludeJs: FAKE_LEAFLET_PRELUDE }));
    await page.goto(url);
    await page.waitForTimeout(150);
    await page.click('button:has-text("Mapa"), button:has-text("Map")');
    await page.waitForTimeout(100);
    const markers = await page.evaluate(() => window.__markers.map((m) => m.latlng));
    t.check('two properties sharing the exact same coordinate DO get nudged apart (still distinguishable pins)',
      markers.length === 2 && (markers[0][0] !== markers[1][0] || markers[0][1] !== markers[1][1]), JSON.stringify(markers));
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
    const hint = await page.locator('.field-group:has(#f_lat) p').textContent();
    t.check('the form explains coordinates are automatic (manual entry is the override)',
      hint.includes('automáticamente') || hint.includes('automatically'), hint);
    await page.close();
  }

  t.group('admin pins tab: diagnostics after re-locating');
  {
    const page = await browser.newPage();
    t.watch(page);
    const url = writePage('client-relocate.html', buildPage({ apiSource: apiSourceFor('admin') }));
    await page.goto(url);
    await page.waitForTimeout(150);
    await page.click('button:has-text("⚙")');
    await page.waitForTimeout(100);
    await page.click('.admin-tabs button.tab:has-text("Pines"), .admin-tabs button.tab:has-text("Pins")');
    await page.waitForTimeout(100);
    const btn = page.locator('.modal-body button:has-text("Reubicar"), .modal-body button:has-text("Re-locate")');
    t.check('the Pins tab offers a re-locate-all-pins button', (await btn.count()) === 1);
    await btn.click();
    await page.waitForTimeout(150);
    const calls = await page.evaluate(() => window.CALLS);
    t.check('clicking it calls relocateAllPins then refreshes the list',
      calls.includes('relocateAllPins') && calls.indexOf('listProperties') > calls.indexOf('relocateAllPins'), calls.join(','));
    const rowCount = await page.locator('.modal-body table.admin-table tbody tr').count();
    t.check('a diagnostic table row is shown per property after re-locating', rowCount === SAMPLE_PROPS.length, rowCount);
    const firstRowText = await page.locator('.modal-body table.admin-table tbody tr').first().innerText();
    t.check('a diagnostic row shows the property reference and the query that was sent',
      firstRowText.includes(SAMPLE_PROPS[0].referencia), firstRowText);
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
