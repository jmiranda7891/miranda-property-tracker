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
    t.check('new property defaults to unspecified status', added.status === 'unspecified');
    t.check('new property is not archived', added.archived === false);

    let threwNoRef = false;
    try { S.addProperty({ direccion: 'no reference' }); } catch (e) { threwNoRef = true; }
    t.check('addProperty requires a reference name', threwNoRef);

    const updated = S.updateProperty(added.id, { estado: 'Guanajuato', status: 'for_sale' });
    const afterUpdate = updated.filter((p) => p.id === added.id)[0];
    t.check('updateProperty applies field changes', afterUpdate.estado === 'Guanajuato' && afterUpdate.status === 'for_sale');

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
};
