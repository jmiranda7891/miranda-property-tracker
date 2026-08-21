// Server-side logic tests. Unlike the client suites, this runs the REAL Code.js in a bare
// vm context with fake Apps Script services (see test/lib/gas-stubs.js) — no browser needed.
// Covers RBAC, property CRUD guards, admin user-management guards, audit logging, export,
// and the one-time seed function.
const { loadServer } = require('./lib/gas-stubs');

module.exports = async function run(t) {
  t.group('bootstrap & RBAC');
  {
    const S = loadServer();
    S.__setUser('first@example.com');
    const boot = S.getBootstrap();
    t.check('first person to open the app becomes admin', boot.role === 'admin', boot.role);
    t.check('bootstrap includes the properties list for an authorized user', Array.isArray(boot.properties));

    S.__setUser('stranger@example.com');
    const boot2 = S.getBootstrap();
    t.check('an unlisted email gets role "none"', boot2.role === 'none', boot2.role);
    t.check('a "none" role gets no properties back', boot2.properties === undefined);

    let threw = false;
    try { S.listProperties(); } catch (e) { threw = true; }
    t.check('requireUser_ throws for an unlisted email', threw);
  }

  t.group('property CRUD');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.getBootstrap(); // bootstraps the registry + seeds admin@example.com as admin

    const afterAdd = S.addProperty({ referencia: 'Test House', direccion: '123 Main St', countryCode: 'MX', ciudad: 'Maravatio', estado: 'Michoacan' });
    t.check('addProperty returns the updated list including the new row', afterAdd.some((p) => p.referencia === 'Test House'));
    const added = afterAdd.filter((p) => p.referencia === 'Test House')[0];
    t.check('new property defaults to libre status', added.status === 'libre');
    t.check('new property defaults escrituras to No', added.escrituras === 'No');
    t.check('new property defaults esEjido to false', added.esEjido === false);
    t.check('new property defaults planLargoPlazo to mantener_individual', added.planLargoPlazo === 'mantener_individual');
    t.check('new property is not archived', added.archived === false);

    let threwNoRef = false;
    try { S.addProperty({ direccion: 'no reference' }); } catch (e) { threwNoRef = true; }
    t.check('addProperty requires a reference name', threwNoRef);

    const updated = S.updateProperty(added.id, { estado: 'Guanajuato', status: 'en_venta', esEjido: true, planLargoPlazo: 'vender' });
    const afterUpdate = updated.filter((p) => p.id === added.id)[0];
    t.check('updateProperty applies field changes', afterUpdate.estado === 'Guanajuato' && afterUpdate.status === 'en_venta'
      && afterUpdate.esEjido === true && afterUpdate.planLargoPlazo === 'vender');

    // A member (non-admin) can add and edit, but not archive/delete.
    S.saveUser('member@example.com', 'member');
    S.__setUser('member@example.com');
    const memberAdded = S.addProperty({ referencia: 'Member Added', direccion: '1 Member Way' });
    t.check('a member can add a property', memberAdded.some((p) => p.referencia === 'Member Added'));
    let memberArchiveThrew = false;
    try { S.archiveProperty(added.id, 'test'); } catch (e) { memberArchiveThrew = true; }
    t.check('a member cannot archive a property', memberArchiveThrew);
    let memberDeleteThrew = false;
    try { S.deleteProperty(added.id, 'Test House'); } catch (e) { memberDeleteThrew = true; }
    t.check('a member cannot delete a property', memberDeleteThrew);

    S.__setUser('admin@example.com');
    const afterArchive = S.archiveProperty(added.id, 'testing archive');
    t.check('admin can archive a property', afterArchive.filter((p) => p.id === added.id)[0].archived === true);
    const afterUnarchive = S.unarchiveProperty(added.id);
    t.check('admin can unarchive a property', afterUnarchive.filter((p) => p.id === added.id)[0].archived === false);

    let mismatchThrew = false;
    try { S.deleteProperty(added.id, 'wrong name'); } catch (e) { mismatchThrew = true; }
    t.check('delete refuses when the confirmation text does not match the reference', mismatchThrew);
    const afterDelete = S.deleteProperty(added.id, 'Test House');
    t.check('delete removes the property once confirmation matches', !afterDelete.some((p) => p.id === added.id));

    const log = S.getAuditLog(50);
    t.check('delete is captured in the audit log with a full snapshot', log.some((e) => e.action === 'delete_property' && e.entityId === added.id));
    const deleteEntry = log.filter((e) => e.action === 'delete_property' && e.entityId === added.id)[0];
    const snapshot = JSON.parse(deleteEntry.detail);
    t.check('the audit snapshot preserves the deleted row’s data', snapshot.referencia === 'Test House');
  }

  t.group('users admin guards');
  {
    const S = loadServer();
    S.__setUser('owner@example.com');
    S.getBootstrap();
    S.saveUser('two@example.com', 'admin');

    let selfRemoveThrew = false;
    try { S.removeUser('owner@example.com'); } catch (e) { selfRemoveThrew = true; }
    t.check('an admin cannot remove themselves', selfRemoveThrew);

    S.removeUser('two@example.com'); // back down to one admin
    let lastAdminThrew = false;
    try { S.saveUser('owner@example.com', 'member'); } catch (e) { lastAdminThrew = true; }
    t.check('cannot demote the last admin', lastAdminThrew);

    S.__setUser('member2@example.com');
    let notAdminThrew = false;
    try { S.listUsers(); } catch (e) { notAdminThrew = true; }
    t.check('a non-admin (even unlisted) cannot list users', notAdminThrew);
  }

  t.group('export');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.getBootstrap();
    S.addProperty({ referencia: 'Export Me', direccion: '9 Export Ave' });
    const res = S.exportProperties(JSON.stringify({}));
    t.check('exportProperties returns a url and a count', typeof res.url === 'string' && res.count >= 1);
    const exportedSs = S.__spreadsheetApp.__registry[Object.keys(S.__spreadsheetApp.__registry).filter((id) => res.url.indexOf(id) >= 0)[0]];
    t.check('the export spreadsheet has a header row with bilingual labels',
      exportedSs.getSheetByName('Properties').getRange(1, 1, 1, 1).getValues()[0][0].indexOf('/') >= 0);
  }

  t.group('one-time seed');
  {
    // The registry seeds itself automatically the moment it's first created (inside
    // ensureRegistry_, triggered by the very first authorized call) - there is no separate
    // manual step to remember, since a headless clasp deploy has no browser to click "Run" in.
    const S = loadServer();
    S.__setUser('admin@example.com');
    const boot = S.getBootstrap();
    t.check('the first bootstrap call auto-seeds all 41 rows', boot.properties.length === 41, 'got ' + boot.properties.length);
    const listAfterSeed = S.listProperties();
    t.check('seeded rows are readable via listProperties', listAfterSeed.length === 41);
    const r2 = S.seedProperties_();
    t.check('seedProperties_() called by hand afterward is a no-op (SEED_DONE guard)', r2.alreadyDone === true, JSON.stringify(r2));
    const listAfterSecondSeed = S.listProperties();
    t.check('re-running seed does not duplicate rows', listAfterSecondSeed.length === listAfterSeed.length);
  }

  t.group('schema v2 migration (legacy-shaped rows on an existing live sheet)');
  {
    // Simulates a spreadsheet seeded before the schema v2 rework: append an old-shaped row
    // directly (bypassing SEED_PROPERTIES_, which already writes the new shape), then force a
    // fresh execution (bustReg_ - CACHE_ never survives a real Apps Script request the way it
    // survives repeated calls within one Node process) and confirm ensureRegistry_'s
    // migrateSchemaV2_ pass normalizes it - while leaving an already-new-shape row untouched.
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.getBootstrap();
    const ssId = Object.keys(S.__spreadsheetApp.__registry)[0];
    const ss = S.__spreadsheetApp.__registry[ssId];
    const sh = ss.getSheetByName('Properties');
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const legacy = {
      id: 'legacy1', countryCode: 'MX', referencia: 'Legacy Test', direccion: 'Calle X', ciudad: 'Maravatio', estado: 'Michoacan',
      status: 'in_process', escrituras: 'Ejido', propuestaTraspaso: 'Fideicomiso (Renta)', observacionesRaw: 'nota vieja'
    };
    sh.appendRow(head.map((h) => (legacy[h] != null ? legacy[h] : '')));
    S.setProp_('SCHEMA_V2_DONE', '');
    S.bustReg_();
    S.getBootstrap();

    const vals = sh.getDataRange().getValues();
    const h2 = vals[0];
    const idx = {}; h2.forEach((x, i) => { idx[x] = i; });
    const migrated = vals.filter((r) => r[idx.id] === 'legacy1')[0];
    const row = (name) => migrated[idx[name]];
    t.check('legacy Ejido escrituras collapses to No', row('escrituras') === 'No', row('escrituras'));
    t.check('legacy Ejido escrituras sets the separate esEjido flag', row('esEjido') === true, row('esEjido'));
    t.check('legacy in_process status collapses to libre', row('status') === 'libre', row('status'));
    t.check('the collapsed legacy status is preserved as a note', String(row('observacionesRaw')).includes('Estatus anterior: Proceso'));
    t.check('legacy "Fideicomiso (Renta)" succession text maps to the fideicomiso plan', row('planLargoPlazo') === 'fideicomiso', row('planLargoPlazo'));
    t.check('the original succession text is preserved as a note', String(row('observacionesRaw')).includes('Traspaso (dato original): Fideicomiso (Renta)'));

    const fresh = vals.filter((r) => r[idx.referencia] === 'Casa Maravatio')[0];
    const freshRow = (name) => fresh[idx[name]];
    t.check('an already-new-shape row is left untouched by the migration', freshRow('status') === 'en_uso' && freshRow('escrituras') === 'Si');
  }

  t.group('map pin corrections (LATLNG_FIX_V1)');
  {
    // Simulates a live sheet already seeded under the wrong (hand-estimated) v1.2 coordinates,
    // then confirms the one-time correction pass fixes exactly the named properties and
    // leaves every other row's coordinates untouched.
    const S = loadServer();
    S.__setUser('admin@example.com');
    const boot = S.getBootstrap();
    const ssId = Object.keys(S.__spreadsheetApp.__registry)[0];
    const ss = S.__spreadsheetApp.__registry[ssId];
    const sh = ss.getSheetByName('Properties');
    const vals0 = sh.getDataRange().getValues();
    const head0 = vals0[0];
    const idx0 = {}; head0.forEach((x, i) => { idx0[x] = i; });
    const oldWrongCoord = 41.0; // stand-in "wrong" value distinct from any real corrected one
    const targetRow = vals0.findIndex((r) => r[idx0.referencia] === 'Dpto. Prairie');
    sh.getRange(targetRow + 1, idx0.lat + 1).setValue(oldWrongCoord);
    const untouchedBefore = vals0.find((r) => r[idx0.referencia] === 'Casa Maravatio')[idx0.lat];

    S.setProp_('LATLNG_FIX_V1_DONE', '');
    S.bustReg_();
    S.getBootstrap();

    const vals1 = sh.getDataRange().getValues();
    const idx1 = {}; vals1[0].forEach((x, i) => { idx1[x] = i; });
    const fixedRow = vals1.find((r) => r[idx1.referencia] === 'Dpto. Prairie');
    t.check('the corrected referencia gets the new coordinates', Math.abs(fixedRow[idx1.lat] - 41.8671) < 0.0001, fixedRow[idx1.lat]);
    const otherRow = vals1.find((r) => r[idx1.referencia] === 'Casa Maravatio');
    t.check('an uncorrected referencia is untouched', otherRow[idx1.lat] === untouchedBefore, otherRow[idx1.lat]);

    const targetId = boot.properties.filter((p) => p.referencia === 'Dpto. Prairie')[0].id;
    const updated = S.updateProperty(targetId, { lat: 10, lng: 20 });
    const afterUpdate = updated.filter((p) => p.id === targetId)[0];
    t.check('updateProperty can change lat/lng', Number(afterUpdate.lat) === 10 && Number(afterUpdate.lng) === 20);
  }

  t.group('automatic geocoding (built-in Maps service, faked)');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    // The geocoder "knows" one street address and every city query; the street hit for the
    // known address, the city-center hit for everything else. A bogus out-of-country result
    // for one specific address exercises the bounds check.
    S.__setGeocode((query) => {
      if (query.includes('Calle Conocida 1')) return { lat: 19.7001, lng: -101.1001 };
      if (query.includes('Calle Rara 9')) return { lat: 40.4168, lng: -3.7038 }; // Madrid, Spain - a wrong-continent geocoder match
      if (query.startsWith('Morelia')) return { lat: 19.7008, lng: -101.1844 };
      return null;
    });
    const boot = S.getBootstrap();
    t.check('the one-time re-geocode pass ran at bootstrap (GEOCODE_ALL_V1)', S.prop_('GEOCODE_ALL_V1_DONE') === 'true');
    const callsAfterBoot = S.__geocodeCalls.length;
    t.check('the pass geocoded every seeded row', callsAfterBoot >= 41, 'calls: ' + callsAfterBoot);
    S.bustReg_();
    S.getBootstrap();
    t.check('a second bootstrap does not re-geocode (flag guard)', S.__geocodeCalls.length === callsAfterBoot, 'calls: ' + S.__geocodeCalls.length);

    const added = S.addProperty({ referencia: 'Geo Test', direccion: 'Calle Conocida 1', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' })
      .filter((p) => p.referencia === 'Geo Test')[0];
    t.check('a new property with blank coordinates is geocoded from its address',
      Math.abs(added.lat - 19.7001) < 0.0001 && Math.abs(added.lng - (-101.1001)) < 0.0001, added.lat + ',' + added.lng);

    const manual = S.addProperty({ referencia: 'Manual Geo', direccion: 'Calle Conocida 1', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX', lat: 1, lng: 2 })
      .filter((p) => p.referencia === 'Manual Geo')[0];
    t.check('manually typed coordinates are respected over the geocoder', Number(manual.lat) === 1 && Number(manual.lng) === 2);

    const fallback = S.addProperty({ referencia: 'Fallback Geo', direccion: 'Calle Desconocida 77', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' })
      .filter((p) => p.referencia === 'Fallback Geo')[0];
    t.check('an unknown street falls back to the city-center coordinates',
      Math.abs(fallback.lat - 19.7008) < 0.0001, fallback.lat + ',' + fallback.lng);

    const bounded = S.addProperty({ referencia: 'Bounds Geo', direccion: 'Calle Rara 9', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' })
      .filter((p) => p.referencia === 'Bounds Geo')[0];
    t.check('an out-of-country geocoder match is rejected and falls back to the city',
      Math.abs(bounded.lat - 19.7008) < 0.0001, bounded.lat + ',' + bounded.lng);

    // Editing the address with untouched (prefilled) coordinates moves the pin; typing new
    // coordinates pins it exactly there instead.
    const moved = S.updateProperty(added.id, { direccion: 'Otra Calle 5', lat: added.lat, lng: added.lng })
      .filter((p) => p.id === added.id)[0];
    t.check('editing the address re-geocodes when coordinates were left as prefilled',
      Math.abs(moved.lat - 19.7008) < 0.0001, moved.lat + ',' + moved.lng);
    const pinned = S.updateProperty(added.id, { direccion: 'Tercera Calle 8', lat: 21.5, lng: -100.5 })
      .filter((p) => p.id === added.id)[0];
    t.check('typing coordinates while editing overrides the geocoder', Number(pinned.lat) === 21.5 && Number(pinned.lng) === -100.5);

    const relocated = S.relocateAllPins();
    t.check('relocateAllPins (admin) re-runs the pass on demand and reports counts',
      relocated.updated > 0 && typeof relocated.kept === 'number', JSON.stringify(relocated));
  }

  t.group('geocoding diagnostics (surfaced in the Admin panel)');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.__setGeocode((query) => {
      if (query.includes('Broken Address 1')) throw new Error('simulated geocoder outage');
      if (query.includes('No Match 2')) return null; // ZERO_RESULTS
      if (query.startsWith('Morelia')) return { lat: 19.7008, lng: -101.1844 };
      return { lat: 19.7001, lng: -101.1001 };
    });
    S.getBootstrap();

    const good = S.addProperty({ referencia: 'Diag Good', direccion: 'Working Address 3', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' });
    const relocated1 = S.relocateAllPins();
    t.check('relocateAllPins returns a details array with one entry per property', Array.isArray(relocated1.details) && relocated1.details.length === good.length, relocated1.details && relocated1.details.length);
    const goodEntry = relocated1.details.find((d) => d.referencia === 'Diag Good');
    t.check('a successful geocode reports method+status+the query actually sent', goodEntry && goodEntry.method === 'address' && goodEntry.status === 'OK' && goodEntry.query.includes('Working Address 3'), JSON.stringify(goodEntry));
    t.check('a successful geocode reports before/after coordinates', goodEntry && goodEntry.after.lat === 19.7001 && goodEntry.after.lng === -101.1001, JSON.stringify(goodEntry));

    S.addProperty({ referencia: 'Diag Error', direccion: 'Broken Address 1', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' });
    S.addProperty({ referencia: 'Diag NoMatch', direccion: 'No Match 2', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' });
    const relocated2 = S.relocateAllPins();
    const errorEntry = relocated2.details.find((d) => d.referencia === 'Diag Error');
    t.check('a geocoder exception on the address attempt is reported even if a city fallback then succeeds',
      errorEntry && errorEntry.method === 'city' && errorEntry.addressStatus.indexOf('ERROR') === 0, JSON.stringify(errorEntry));
    const noMatchEntry = relocated2.details.find((d) => d.referencia === 'Diag NoMatch');
    t.check('a genuine no-result falls back to the city and reports method "city"',
      noMatchEntry && noMatchEntry.method === 'city' && noMatchEntry.status === 'OK', JSON.stringify(noMatchEntry));

    // Dpto. Prairie's seed address no longer has "Unit 2201" inline (v1.7 split it into the
    // separate direccion2 field) - so this checks stripUnitNoise_ directly, as a defense-in-depth
    // safety net for anyone who still types a unit number straight into the address field.
    S.addProperty({ referencia: 'Diag Inline Unit', direccion: '1500 W Test St, Unit 501', ciudad: 'Morelia', estado: 'Michoacan', countryCode: 'MX' });
    S.relocateAllPins();
    t.check('geocodeProperty_ strips unit/apartment noise from the query before geocoding', (() => {
      const calls = S.__geocodeCalls;
      const unitCall = calls.find((c) => c.query.includes('1500 W Test St'));
      return unitCall && !unitCall.query.includes('Unit');
    })(), JSON.stringify(S.__geocodeCalls.filter((c) => c.query.includes('Test St'))));

    // Regression: the bare "Uni" alternative used to have no trailing word-boundary check, so
    // it matched as a PREFIX of an ordinary word too - it silently ate the real street name out
    // of "22 N Union St" (Edificio Oficina Aurora), because "Uni" + a greedy trailing run of
    // word characters devoured "on". Confirmed live via test/static.js's seed-data check.
    S.addProperty({ referencia: 'Diag Union St', direccion: '22 N Union St', ciudad: 'Aurora', estado: 'Illinois', countryCode: 'US' });
    S.relocateAllPins();
    t.check('a real street name starting with "Uni" ("Union") is never mistaken for unit/apartment noise', (() => {
      const calls = S.__geocodeCalls;
      const unionCall = calls.find((c) => c.query.includes('Union'));
      return !!unionCall;
    })(), JSON.stringify(S.__geocodeCalls.filter((c) => c.query.includes('Aurora'))));
  }

  t.group('separate apt/interior number field (direccion2) - v1.7');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.getBootstrap();

    const afterAdd = S.addProperty({ referencia: 'Depto Test', direccion: '100 Main St', direccion2: 'Apt 4B', countryCode: 'MX', ciudad: 'Morelia', estado: 'Michoacan' });
    const added = afterAdd.filter((p) => p.referencia === 'Depto Test')[0];
    t.check('addProperty stores direccion2 separately from direccion', added.direccion === '100 Main St' && added.direccion2 === 'Apt 4B', JSON.stringify(added));

    const updated = S.updateProperty(added.id, { direccion2: 'Apt 9Z' });
    const afterUpdate = updated.filter((p) => p.id === added.id)[0];
    t.check('updateProperty can change direccion2', afterUpdate.direccion2 === 'Apt 9Z', afterUpdate.direccion2);

    // splitAddressUnitV1_ migration: a legacy row with the unit number still embedded in
    // direccion (as every property looked before v1.7) gets it pulled out into direccion2,
    // while a row someone has already given its own direccion2 value is left untouched.
    const ssId = Object.keys(S.__spreadsheetApp.__registry)[0];
    const ss = S.__spreadsheetApp.__registry[ssId];
    const sh = ss.getSheetByName('Properties');
    const vals0 = sh.getDataRange().getValues();
    const head0 = vals0[0];
    const idx0 = {}; head0.forEach((x, i) => { idx0[x] = i; });
    const legacyRow = vals0.findIndex((r) => r[idx0.referencia] === 'Depto Test');
    sh.getRange(legacyRow + 1, idx0.direccion + 1).setValue('200 Legacy Ave, Unit 12');
    sh.getRange(legacyRow + 1, idx0.direccion2 + 1).setValue('');
    const alreadySplitBefore = vals0.find((r) => r[idx0.referencia] === 'Casa Maravatio')[idx0.direccion2];

    S.setProp_('SPLIT_UNIT_V1_DONE', '');
    S.bustReg_();
    S.getBootstrap();

    const vals1 = sh.getDataRange().getValues();
    const idx1 = {}; vals1[0].forEach((x, i) => { idx1[x] = i; });
    const splitRow = vals1.find((r) => r[idx1.referencia] === 'Depto Test');
    t.check('the migration pulls an embedded unit number out into direccion2',
      splitRow[idx1.direccion] === '200 Legacy Ave' && splitRow[idx1.direccion2] === 'Unit 12', JSON.stringify({ direccion: splitRow[idx1.direccion], direccion2: splitRow[idx1.direccion2] }));
    const untouchedRow = vals1.find((r) => r[idx1.referencia] === 'Casa Maravatio');
    t.check('a row with no embedded unit noise is left untouched by the migration',
      untouchedRow[idx1.direccion2] === alreadySplitBefore, untouchedRow[idx1.direccion2]);
  }

  t.group('tracking fields (lot/construction size, acquisition basis, ownership/legal, registry IDs) - v1.8');
  {
    const S = loadServer();
    S.__setUser('admin@example.com');
    S.getBootstrap();

    const afterAdd = S.addProperty({
      referencia: 'Tracking Test', direccion: '300 Test Ave', countryCode: 'MX', ciudad: 'Morelia', estado: 'Michoacan',
      lotSizeValue: '500', lotSizeUnit: 'm2', constructionSizeValue: '220', constructionSizeUnit: 'm2',
      acquisitionDate: '2010-05-01', acquisitionPriceUSD: '80000',
      ownershipPct: '50', liensNotes: 'Second mortgage with BBVA', hoaFeeUSD: '150', documentsUrl: 'https://drive.google.com/folder/xyz',
      folioReal: 'FR-12345', claveCatastral: 'CAT-9999', cuentaPredial: 'PRED-001', usoDeSuelo: 'Habitacional', notario: 'Lic. Juan Perez',
      county: 'should be ignored for MX rows but stored anyway', titlePolicyInfo: '', legalDescription: ''
    });
    const added = afterAdd.filter((p) => p.referencia === 'Tracking Test')[0];
    t.check('addProperty stores lot/construction size + unit', added.lotSizeValue === 500 && added.lotSizeUnit === 'm2' && added.constructionSizeValue === 220 && added.constructionSizeUnit === 'm2', JSON.stringify(added));
    t.check('addProperty stores acquisition date + price separately from the current estimate', added.acquisitionDate === '2010-05-01' && added.acquisitionPriceUSD === 80000, JSON.stringify(added));
    t.check('addProperty stores ownership %, liens, HOA fee, documents URL', added.ownershipPct === 50 && added.liensNotes === 'Second mortgage with BBVA' && added.hoaFeeUSD === 150 && added.documentsUrl === 'https://drive.google.com/folder/xyz', JSON.stringify(added));
    t.check('addProperty stores MX registry IDs', added.folioReal === 'FR-12345' && added.claveCatastral === 'CAT-9999' && added.cuentaPredial === 'PRED-001' && added.usoDeSuelo === 'Habitacional' && added.notario === 'Lic. Juan Perez', JSON.stringify(added));

    const updated = S.updateProperty(added.id, { county: 'Cook County', titlePolicyInfo: 'Chicago Title #4567', legalDescription: 'Lot 4, Block 2' });
    const afterUpdate = updated.filter((p) => p.id === added.id)[0];
    t.check('updateProperty can set USA-only registry fields', afterUpdate.county === 'Cook County' && afterUpdate.titlePolicyInfo === 'Chicago Title #4567' && afterUpdate.legalDescription === 'Lot 4, Block 2', JSON.stringify(afterUpdate));

    t.check('a property added with none of the new fields gets clean blank/null defaults, not errors', (() => {
      const bare = S.addProperty({ referencia: 'Bare Test', direccion: '1 Bare St', countryCode: 'US', ciudad: 'Chicago', estado: 'Illinois' })
        .filter((p) => p.referencia === 'Bare Test')[0];
      return bare.lotSizeValue === null && bare.acquisitionPriceUSD === null && bare.ownershipPct === null
        && bare.liensNotes === '' && bare.folioReal === '' && bare.county === '';
    })());

    // addTrackingFieldsV1_ must retrofit a LIVE sheet that predates these columns without
    // disturbing existing data - simulate a pre-v1.8 sheet by truncating the fake sheet's
    // header row back to before the tracking columns existed, then re-bootstrap.
    const ssId = Object.keys(S.__spreadsheetApp.__registry)[0];
    const ss = S.__spreadsheetApp.__registry[ssId];
    const sh = ss.getSheetByName('Properties');
    // Removing headers alone would misalign every surviving row's values (a plain filter
    // shifts column positions without moving the data in sh.rows to match) - rebuild both
    // headers and row data together, keeping only the columns that predate v1.8.
    const survivingIdx = [];
    const newHeaders = [];
    sh.headers.forEach((h, i) => {
      if (S.TRACKING_FIELDS_V1_.indexOf(h) === -1) { survivingIdx.push(i); newHeaders.push(h); }
    });
    sh.rows = sh.rows.map((row) => survivingIdx.map((i) => row[i]));
    sh.headers = newHeaders;
    t.check('the simulated pre-v1.8 sheet no longer has the tracking columns', !sh.headers.includes('lotSizeValue') && !sh.headers.includes('folioReal'), sh.headers.join(','));

    S.setProp_('TRACKING_FIELDS_V1_DONE', '');
    S.bustReg_();
    const beforeRetrofit = S.listProperties().filter((p) => p.referencia === 'Tracking Test')[0];
    S.getBootstrap();

    t.check('addTrackingFieldsV1_ retrofits a pre-v1.8 sheet with the missing columns',
      S.TRACKING_FIELDS_V1_.every((c) => sh.headers.includes(c)), sh.headers.join(','));
    const afterRetrofit = S.listProperties().filter((p) => p.referencia === 'Tracking Test')[0];
    t.check('retrofitting the columns does not disturb pre-existing row data',
      afterRetrofit.referencia === beforeRetrofit.referencia && afterRetrofit.direccion === beforeRetrofit.direccion,
      JSON.stringify({ before: beforeRetrofit, after: afterRetrofit }));
  }
};
