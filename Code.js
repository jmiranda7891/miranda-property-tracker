/**
 * Miranda Family Property Tracker - server (Apps Script)
 * Entry point + bridge functions. All client<->server calls go through google.script.run.
 * See CLAUDE.md for architecture, data model, deploy workflow, and gotchas.
 */

// ============================ Web app entry ============================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Miranda Family Property Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ============================ Identity ============================
function getUserEmail() {
  return Session.getActiveUser().getEmail() || '';
}

// ============================ Script properties ============================
// REGISTRY_ID (the one spreadsheet's id) and SEED_DONE live here, never in a sheet cell.
function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}
function setProp_(key, val) {
  PropertiesService.getScriptProperties().setProperty(key, val);
}

// ============================ Sheet utilities ============================
// Read a tab into array-of-objects keyed by the header row.
function rows_(sheet) {
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return [];
  var head = vals[0];
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var o = {};
    for (var c = 0; c < head.length; c++) o[head[c]] = vals[r][c];
    out.push(o);
  }
  return out;
}
function tab_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}
function appendObj_(sh, obj) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = head.map(function (h) { return obj[h] == null ? '' : obj[h]; });
  sh.appendRow(row);
}
function uid_() { return Utilities.getUuid().slice(0, 8); }
function nowIso_() { return new Date().toISOString(); }
// Add a column to a sheet's header row if it doesn't exist yet (safe on live sheets).
function ensureColumn_(sh, name) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (head.indexOf(name) >= 0) return false;
  sh.getRange(1, lastCol + 1).setValue(name);
  return true;
}
// Find a row by a key column's value. Returns {row (1-based sheet row), head, values} or null.
function findRowByKey_(sh, keyCol, keyVal) {
  var vals = sh.getDataRange().getValues();
  var head = vals[0];
  var kIdx = head.indexOf(keyCol);
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][kIdx]) === String(keyVal)) return { row: r + 1, head: head, values: vals[r] };
  }
  return null;
}

// ============================ The one spreadsheet (single-tenant registry) ============================
var PROPERTIES_HEADERS_ = [
  'id', 'countryCode', 'paisRaw', 'tipo', 'referencia', 'propietario', 'direccion', 'ciudad',
  'estado', 'cp', 'observacionesRaw', 'status', 'precioEstimadoUSD', 'precioEstimadoIsPlaceholder',
  'escrituras', 'propEscriturado', 'propuestaTraspaso', 'pin',
  'aiResearchEN', 'aiResearchES', 'aiValueEstimateEN', 'aiValueEstimateES', 'researchDate',
  'archived', 'archivedReason', 'createdBy', 'createdAt', 'updatedBy', 'updatedAt'
];

function ensureRegistry_() {
  var id = prop_('REGISTRY_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }
  var ss = SpreadsheetApp.create('Miranda Family - Property Registry');
  tab_(ss, 'Properties', PROPERTIES_HEADERS_);
  tab_(ss, 'Users', ['email', 'role', 'addedBy', 'addedAt']);
  tab_(ss, 'AuditLog', ['ts', 'user', 'action', 'entityType', 'entityId', 'detail']);
  // seed the deploying/first-opening user as the first admin
  var me = getUserEmail();
  if (me) appendObj_(ss.getSheetByName('Users'), { email: me, role: 'admin', addedBy: 'system', addedAt: nowIso_() });
  var s1 = ss.getSheetByName('Sheet1'); if (s1) ss.deleteSheet(s1);
  setProp_('REGISTRY_ID', ss.getId());
  // Populate the family's property data right away, since this only runs the moment the
  // registry itself is first created - there is no separate manual "run this once from the
  // editor" step to remember (clasp's headless deploy path has no browser to click Run in
  // anyway). seedProperties_() below is guarded by the same SEED_DONE flag, so it stays a
  // safe no-op if ever called again by hand.
  seedIntoRegistry_(ss, me || 'system');
  return ss;
}

// NOTE: takes the spreadsheet directly and writes audit rows itself, rather than going
// through ss_()/audit_() - this runs from inside ensureRegistry_ itself, before CACHE_.ss is
// set, so calling ss_() here would recurse back into ensureRegistry_() and create a second
// spreadsheet.
function seedIntoRegistry_(ss, actor) {
  if (!SEED_PROPERTIES_.length || prop_('SEED_DONE') === 'true') return 0;
  var sh = ss.getSheetByName('Properties');
  var now = nowIso_();
  SEED_PROPERTIES_.forEach(function (p) {
    var row = Object.assign({}, SEED_DEFAULTS_, {
      id: uid_(), researchDate: now, createdBy: actor, createdAt: now, updatedBy: actor, updatedAt: now
    }, p);
    appendObj_(sh, row);
  });
  setProp_('SEED_DONE', 'true');
  appendObj_(ss.getSheetByName('AuditLog'), {
    ts: nowIso_(), user: String(actor || ''), action: 'seed_properties',
    entityType: 'property', entityId: 'bulk', detail: JSON.stringify({ count: SEED_PROPERTIES_.length })
  });
  return SEED_PROPERTIES_.length;
}

// Execution-scoped cache. Apps Script rebuilds globals on every invocation, so this only
// kills repeated openById/reads WITHIN one request, never goes stale across requests.
var CACHE_ = { ss: null, users: null };
function bustReg_() { CACHE_.ss = null; CACHE_.users = null; }

function ss_() {
  if (CACHE_.ss) return CACHE_.ss;
  CACHE_.ss = ensureRegistry_();
  return CACHE_.ss;
}
function usersRows_() {
  if (CACHE_.users) return CACHE_.users;
  CACHE_.users = rows_(ss_().getSheetByName('Users'));
  return CACHE_.users;
}

// ============================ RBAC ============================
var ROLES_ = ['member', 'admin'];

function roleFor_(email) {
  if (!email) return 'none';
  var users = usersRows_();
  if (!users.length) return 'admin'; // bootstrap: no users yet => first opener is admin
  var u = users.filter(function (x) { return String(x.email).toLowerCase() === String(email).toLowerCase(); })[0];
  return u ? String(u.role) : 'none'; // unlisted email => NO access (private family app, not org-open)
}

function requireUser_() {
  var role = roleFor_(getUserEmail());
  if (role === 'none') throw new Error('You are not authorized to use this app. Ask an admin to add your email.');
  return role;
}
function requireAdmin_() {
  var role = requireUser_();
  if (role !== 'admin') throw new Error('Admin only.');
  return role;
}

// ============================ Bootstrap ============================
function getBootstrap() {
  var email = getUserEmail();
  var role = roleFor_(email);
  if (role === 'none') {
    return { email: email, role: role, build: BUILD_LABEL_ };
  }
  return {
    email: email,
    role: role,
    build: BUILD_LABEL_,
    properties: listProperties()
  };
}

var BUILD_LABEL_ = 'v1';

// ============================ Properties: read ============================
var STATUSES_ = ['in_use', 'for_sale', 'in_process', 'ejido', 'family_matter', 'investment', 'held_by_company', 'sold', 'unspecified'];
var TIPOS_ = ['Casa', 'Departamento', 'Edificio', 'Local', 'Oficina', 'Terreno', 'Otro'];

function listProperties() {
  requireUser_();
  return rows_(ss_().getSheetByName('Properties')).map(function (p) {
    return {
      id: String(p.id),
      countryCode: String(p.countryCode || ''),
      paisRaw: String(p.paisRaw || ''),
      tipo: String(p.tipo || ''),
      referencia: String(p.referencia || ''),
      propietario: String(p.propietario || ''),
      direccion: String(p.direccion || ''),
      ciudad: String(p.ciudad || ''),
      estado: String(p.estado || ''),
      cp: String(p.cp || ''),
      observacionesRaw: String(p.observacionesRaw || ''),
      status: String(p.status || 'unspecified'),
      precioEstimadoUSD: p.precioEstimadoUSD === '' || p.precioEstimadoUSD == null ? null : Number(p.precioEstimadoUSD),
      precioEstimadoIsPlaceholder: p.precioEstimadoIsPlaceholder === true || p.precioEstimadoIsPlaceholder === 'TRUE',
      escrituras: String(p.escrituras || ''),
      propEscriturado: String(p.propEscriturado || ''),
      propuestaTraspaso: String(p.propuestaTraspaso || ''),
      pin: String(p.pin || ''),
      aiResearchEN: String(p.aiResearchEN || ''),
      aiResearchES: String(p.aiResearchES || ''),
      aiValueEstimateEN: String(p.aiValueEstimateEN || ''),
      aiValueEstimateES: String(p.aiValueEstimateES || ''),
      researchDate: String(p.researchDate || ''),
      archived: p.archived === true || p.archived === 'TRUE',
      archivedReason: String(p.archivedReason || ''),
      createdBy: String(p.createdBy || ''),
      createdAt: String(p.createdAt || ''),
      updatedBy: String(p.updatedBy || ''),
      updatedAt: String(p.updatedAt || '')
    };
  });
}

// ============================ Properties: write ============================
var EDITABLE_FIELDS_ = [
  'countryCode', 'paisRaw', 'tipo', 'referencia', 'propietario', 'direccion', 'ciudad', 'estado', 'cp',
  'observacionesRaw', 'status', 'precioEstimadoUSD', 'precioEstimadoIsPlaceholder',
  'escrituras', 'propEscriturado', 'propuestaTraspaso', 'pin'
];

function addProperty(form) {
  requireUser_();
  form = form || {};
  var referencia = String(form.referencia || '').trim();
  if (!referencia) throw new Error('Reference name is required.');
  var direccion = String(form.direccion || '').trim();
  if (!direccion) throw new Error('Address is required.');
  var countryCode = form.countryCode === 'US' ? 'US' : 'MX';
  var status = STATUSES_.indexOf(form.status) >= 0 ? form.status : 'unspecified';
  var id = uid_();
  var me = getUserEmail();
  var now = nowIso_();
  var row = {
    id: id, countryCode: countryCode,
    paisRaw: String(form.paisRaw || (countryCode === 'US' ? 'Estados Unidos' : 'Mexico')),
    tipo: String(form.tipo || ''), referencia: referencia, propietario: String(form.propietario || ''),
    direccion: direccion, ciudad: String(form.ciudad || ''), estado: String(form.estado || ''), cp: String(form.cp || ''),
    observacionesRaw: String(form.observacionesRaw || ''), status: status,
    precioEstimadoUSD: form.precioEstimadoUSD ? Number(form.precioEstimadoUSD) : '',
    precioEstimadoIsPlaceholder: false,
    escrituras: String(form.escrituras || ''), propEscriturado: String(form.propEscriturado || ''),
    propuestaTraspaso: String(form.propuestaTraspaso || ''), pin: String(form.pin || ''),
    aiResearchEN: '', aiResearchES: '', aiValueEstimateEN: '', aiValueEstimateES: '', researchDate: '',
    archived: false, archivedReason: '', createdBy: me, createdAt: now, updatedBy: me, updatedAt: now
  };
  appendObj_(ss_().getSheetByName('Properties'), row);
  bustReg_();
  audit_(me, 'add_property', 'property', id, row);
  return listProperties();
}

function updateProperty(id, form) {
  requireUser_();
  form = form || {};
  var sh = ss_().getSheetByName('Properties');
  var found = findRowByKey_(sh, 'id', id);
  if (!found) throw new Error('Property not found.');
  var head = found.head;
  var before = {};
  head.forEach(function (h, i) { before[h] = found.values[i]; });
  EDITABLE_FIELDS_.forEach(function (k) {
    if (Object.prototype.hasOwnProperty.call(form, k)) {
      var c = head.indexOf(k);
      if (c >= 0) sh.getRange(found.row, c + 1).setValue(form[k]);
    }
  });
  var me = getUserEmail();
  var uCol = head.indexOf('updatedBy'), tCol = head.indexOf('updatedAt');
  sh.getRange(found.row, uCol + 1).setValue(me);
  sh.getRange(found.row, tCol + 1).setValue(nowIso_());
  bustReg_();
  audit_(me, 'edit_property', 'property', id, { before: before, changed: form });
  return listProperties();
}

function archiveProperty(id, reason) {
  requireAdmin_();
  var sh = ss_().getSheetByName('Properties');
  var found = findRowByKey_(sh, 'id', id);
  if (!found) throw new Error('Property not found.');
  var head = found.head;
  sh.getRange(found.row, head.indexOf('archived') + 1).setValue(true);
  sh.getRange(found.row, head.indexOf('archivedReason') + 1).setValue(String(reason || ''));
  bustReg_();
  audit_(getUserEmail(), 'archive_property', 'property', id, { reason: String(reason || '') });
  return listProperties();
}

function unarchiveProperty(id) {
  requireAdmin_();
  var sh = ss_().getSheetByName('Properties');
  var found = findRowByKey_(sh, 'id', id);
  if (!found) throw new Error('Property not found.');
  var head = found.head;
  sh.getRange(found.row, head.indexOf('archived') + 1).setValue(false);
  sh.getRange(found.row, head.indexOf('archivedReason') + 1).setValue('');
  bustReg_();
  audit_(getUserEmail(), 'unarchive_property', 'property', id, {});
  return listProperties();
}

// Hard delete. Admin-only; the client makes the user type the property's referencia to
// confirm, and we check it again here server-side as defense in depth. The full row is
// snapshotted into AuditLog.detail first, since this is the only recovery path.
function deleteProperty(id, confirmReferencia) {
  requireAdmin_();
  var sh = ss_().getSheetByName('Properties');
  var found = findRowByKey_(sh, 'id', id);
  if (!found) throw new Error('Property not found.');
  var head = found.head;
  var snapshot = {};
  head.forEach(function (h, i) { snapshot[h] = found.values[i]; });
  if (String(confirmReferencia || '').trim() !== String(snapshot.referencia || '').trim()) {
    throw new Error('Confirmation text does not match the property reference name.');
  }
  sh.deleteRow(found.row);
  bustReg_();
  audit_(getUserEmail(), 'delete_property', 'property', id, snapshot);
  return listProperties();
}

// ============================ Users admin ============================
function listUsers() {
  requireAdmin_();
  return usersRows_().map(function (u) {
    return { email: String(u.email), role: String(u.role), addedBy: String(u.addedBy || ''), addedAt: String(u.addedAt || '') };
  });
}

function saveUser(email, role) {
  requireAdmin_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Invalid email.');
  if (ROLES_.indexOf(role) < 0) throw new Error('Invalid role.');
  var sh = ss_().getSheetByName('Users');
  var vals = sh.getDataRange().getValues();
  var head = vals[0];
  var eIdx = head.indexOf('email'), rIdx = head.indexOf('role');
  var admins = 0, targetRow = -1;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][rIdx]) === 'admin') admins++;
    if (String(vals[r][eIdx]).toLowerCase() === email) targetRow = r;
  }
  var isNew = targetRow < 0;
  if (!isNew && String(vals[targetRow][rIdx]) === 'admin' && role !== 'admin' && admins <= 1) {
    throw new Error('Cannot demote the last admin.');
  }
  if (isNew) {
    appendObj_(sh, { email: email, role: role, addedBy: getUserEmail(), addedAt: nowIso_() });
  } else {
    sh.getRange(targetRow + 1, rIdx + 1).setValue(role);
  }
  bustReg_();
  audit_(getUserEmail(), isNew ? 'add_user' : 'edit_user_role', 'user', email, { role: role });
  return listUsers();
}

function removeUser(email) {
  requireAdmin_();
  email = String(email || '').trim().toLowerCase();
  if (email === String(getUserEmail()).toLowerCase()) throw new Error('You cannot remove yourself.');
  var sh = ss_().getSheetByName('Users');
  var vals = sh.getDataRange().getValues();
  var head = vals[0];
  var eIdx = head.indexOf('email'), rIdx = head.indexOf('role');
  var admins = 0, targetRow = -1;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][rIdx]) === 'admin') admins++;
    if (String(vals[r][eIdx]).toLowerCase() === email) targetRow = r;
  }
  if (targetRow < 0) throw new Error('User not found.');
  if (String(vals[targetRow][rIdx]) === 'admin' && admins <= 1) throw new Error('Cannot remove the last admin.');
  sh.deleteRow(targetRow + 1);
  bustReg_();
  audit_(getUserEmail(), 'remove_user', 'user', email, {});
  return listUsers();
}

// ============================ Audit log ============================
function audit_(user, action, entityType, entityId, detail) {
  appendObj_(ss_().getSheetByName('AuditLog'), {
    ts: nowIso_(), user: String(user || ''), action: String(action || ''),
    entityType: String(entityType || ''), entityId: String(entityId || ''),
    detail: JSON.stringify(detail || {})
  });
}

function getAuditLog(limit) {
  requireAdmin_();
  var all = rows_(ss_().getSheetByName('AuditLog'));
  var n = Math.min(all.length, limit || 100);
  var out = [];
  for (var i = all.length - 1; i >= 0 && out.length < n; i--) {
    var e = all[i];
    out.push({
      ts: String(e.ts), user: String(e.user), action: String(e.action),
      entityType: String(e.entityType), entityId: String(e.entityId), detail: String(e.detail)
    });
  }
  return out;
}

// ============================ Export ============================
var EXPORT_HEADERS_ = [
  ['referencia', 'Reference / Referencia'],
  ['tipo', 'Type / Tipo'],
  ['countryCode', 'Country / País'],
  ['direccion', 'Address / Dirección'],
  ['ciudad', 'City / Ciudad'],
  ['estado', 'State / Estado'],
  ['cp', 'Zip / C.P.'],
  ['propietario', 'Owner(s) / Propietario(s)'],
  ['status', 'Status / Estatus'],
  ['observacionesRaw', 'Notes / Observaciones'],
  ['precioEstimadoUSD', 'Estimated Price USD / Precio Estimado USD'],
  ['escrituras', 'Deed Status / Escrituras'],
  ['propEscriturado', 'Titled To / Prop. Escriturado'],
  ['propuestaTraspaso', 'Succession Plan / Propuesta de Traspaso'],
  ['pin', 'Parcel PIN'],
  ['archived', 'Archived / Archivado'],
  ['mapsLink', 'Google Maps']
];

function mapsLink_(p) {
  var q = [p.direccion, p.ciudad, p.estado, p.countryCode === 'US' ? 'USA' : 'Mexico'].filter(String).join(', ');
  return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(q);
}

// filterJson: JSON string, either '{}'/omitted (export everything) or '{"ids":["..."]}'
function exportProperties(filterJson) {
  requireUser_();
  var filter = {};
  try { filter = JSON.parse(filterJson || '{}'); } catch (e) { filter = {}; }
  var all = listProperties();
  var selected = (filter.ids && filter.ids.length) ? all.filter(function (p) { return filter.ids.indexOf(p.id) >= 0; }) : all;
  var name = 'Miranda Properties Export - ' + nowIso_().slice(0, 10);
  var ss = SpreadsheetApp.create(name);
  var sh = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  sh.setName('Properties');
  sh.getRange(1, 1, 1, EXPORT_HEADERS_.length).setValues([EXPORT_HEADERS_.map(function (h) { return h[1]; })]);
  var out = selected.map(function (p) {
    return EXPORT_HEADERS_.map(function (h) {
      var k = h[0];
      if (k === 'mapsLink') return mapsLink_(p);
      if (k === 'archived') return p.archived ? 'Yes / Sí' : 'No';
      return p[k] == null ? '' : p[k];
    });
  });
  if (out.length) sh.getRange(2, 1, out.length, EXPORT_HEADERS_.length).setValues(out);
  sh.autoResizeColumns(1, EXPORT_HEADERS_.length);
  return { url: ss.getUrl(), count: out.length };
}

// ============================ One-time seed (run once from the Apps Script editor) ============================
// SEED_PROPERTIES_ below is hand-authored from the family's original Excel listing plus a
// one-time web-research pass (city/area context + an informal value-estimate rationale,
// bilingual). It is NOT a live AI integration - nothing here calls an external API at
// runtime. See CLAUDE.md "Data migration" for the source-cleanup decisions this reflects
// (country-spelling normalization, the sold USA house archived, the one placeholder price
// flagged, etc).
var SEED_DEFAULTS_ = {
  propietario: '', cp: '', observacionesRaw: '', escrituras: '', propEscriturado: '',
  propuestaTraspaso: '', pin: '', precioEstimadoIsPlaceholder: false,
  aiResearchEN: '', aiResearchES: '', aiValueEstimateEN: '', aiValueEstimateES: '',
  archived: false, archivedReason: ''
};

var SEED_PROPERTIES_ = [
  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Casa Maravatio', propietario: 'Jorge Miranda Juárez', direccion: 'Leona Vicario #110', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 1400000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'Fideicomiso (Compartido)',
    aiResearchEN: 'Maravatío is a small agricultural-service city in eastern Michoacán along the Mexico City–Morelia corridor, with a modest, largely local housing market rather than a resort or metro-fed one. Recent listings for houses with meaningful land run roughly $3-6 million pesos, but plain in-town homes without acreage trade far lower, reflecting the town’s limited buyer pool and modest local incomes.',
    aiResearchES: 'Maravatío es una pequeña ciudad de servicios agrícolas en el oriente de Michoacán, sobre el corredor Ciudad de México–Morelia, con un mercado de vivienda modesto y mayormente local, no impulsado por turismo o zona metropolitana. Listados recientes de casas con terreno amplio rondan los $3-6 millones de pesos, pero casas urbanas comunes sin mucho terreno se cotizan muy por debajo.',
    aiValueEstimateEN: 'An asking price of $1,400,000 for a house on a standard in-town lot looks notably high for Maravatío’s market, where comparable homes without large land parcels more typically list well under $1 million pesos. This is not a formal appraisal, but the number is worth double-checking against lot size, construction quality, and whether it was set years ago and never updated.',
    aiValueEstimateES: 'Un precio de venta de $1,400,000 por una casa en un lote urbano estándar se ve notablemente alto para el mercado de Maravatío, donde casas comparables sin terreno extenso normalmente se listan muy por debajo de $1 millón de pesos. Esto no es un avalúo formal, pero conviene revisar el tamaño del lote, la calidad de construcción y si el precio quedó fijado hace años sin actualizarse.' },

  { countryCode: 'MX', paisRaw: 'México', tipo: 'Edificio', referencia: 'Edificio IESA', propietario: 'Inorporated Express, Pamela Miranda Molina', direccion: 'Ocampo 51 Col. Centro Maravatío, Mich.', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 1370000, escrituras: 'Si', propEscriturado: 'Incorporated Express, S.A de C.V', propuestaTraspaso: 'Fideicomiso (Renta)',
    aiResearchEN: 'Maravatío’s small commercial core along streets like Ocampo serves a town-and-farmland customer base rather than a big-city retail market, so commercial buildings there are valued mainly on location within the compact centro and rentability to local businesses, schools, or offices. Commercial space of any real size in a town this size is relatively scarce, which can support pricing for a well-placed building.',
    aiResearchES: 'El pequeño centro comercial de Maravatío, sobre calles como Ocampo, atiende a la población del pueblo y del campo circundante, no a un mercado comercial de gran ciudad, así que los edificios comerciales se valoran sobre todo por su ubicación en el centro compacto y su potencial de renta para negocios, escuelas u oficinas locales. Espacio comercial de tamaño considerable es relativamente escaso en un pueblo así.',
    aiValueEstimateEN: 'At $1,370,000 for a commercial building in a small-city center like Maravatío, the price sits toward the higher end of what this market typically supports, though a multi-story or multi-unit building with rentable space could justify it. This is only an informal read, not an appraisal — a real assessment would need square footage, condition, and current rental income.',
    aiValueEstimateES: 'A $1,370,000 por un edificio comercial en el centro de Maravatío, el precio está en la parte alta de lo que normalmente soporta este mercado, aunque un edificio de varios niveles o unidades rentables podría justificarlo. Esto es solo una lectura informal, no un avalúo — necesitaría metros cuadrados, estado del inmueble e ingresos de renta actuales.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Edificio', referencia: 'Oficina OE Maravatio', propietario: 'Jorge Miranda Juarez', direccion: 'Madero #55', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'a nombre de CC', status: 'held_by_company', precioEstimadoUSD: null, precioEstimadoIsPlaceholder: true, escrituras: 'Si', propEscriturado: 'OECC', propuestaTraspaso: 'Fideicomiso (Renta)',
    aiResearchEN: 'This small office building on Madero street in central Maravatío sits within the same modest small-town commercial market as the family’s other Maravatío properties — limited demand, mostly local tenants, and no resort or metro pricing pressure.',
    aiResearchES: 'Este pequeño edificio de oficinas en la calle Madero, en el centro de Maravatío, está dentro del mismo mercado comercial modesto de pueblo pequeño que otras propiedades de la familia en Maravatío — demanda limitada, inquilinos mayormente locales, sin presión de precios de zona turística o metropolitana.',
    aiValueEstimateEN: 'The listed price in the original source file was literally "1" — plainly a data placeholder, not a real figure, and it should not be read as any indication of value. A realistic figure for a small office building in Maravatío’s centro would need to be researched from comparable local commercial listings.',
    aiValueEstimateES: 'El precio en el archivo original era literalmente "1" — claramente un marcador de datos (placeholder), no una cifra real, y no debe interpretarse como indicación de valor. Una cifra realista tendría que investigarse comparando con listados comerciales locales similares.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Oficina Cd. Hidalgo', propietario: 'Casa de Cambio', direccion: 'Cuauhtemoc Sur 6', ciudad: 'Cd. Hidalgo', estado: 'Michoacan', cp: '61100', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 1370000, escrituras: 'Si', propEscriturado: 'OECC', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'Ciudad Hidalgo is a small Michoacán city of roughly 115,000 people near Maravatío, with an economy historically built on forestry and furniture-making and tourism being developed as a newer sector; the commercial real estate market there is small-city scale, not comparable to a state capital or resort town.',
    aiResearchES: 'Ciudad Hidalgo es una pequeña ciudad de Michoacán de alrededor de 115,000 habitantes cerca de Maravatío, con una economía históricamente basada en la silvicultura y la fabricación de muebles, y el turismo desarrollándose como sector más nuevo; el mercado inmobiliario comercial ahí es de escala de ciudad pequeña.',
    aiValueEstimateEN: '$1,370,000 for a storefront/office in Ciudad Hidalgo looks unusually high for this size of market — commercial land there has been seen priced from a few hundred to a few thousand pesos per square meter depending on location, and a plain small-city storefront at this price would need substantial size or a standout location to justify it. This is an informal flag, not an appraisal.',
    aiValueEstimateES: '$1,370,000 por un local/oficina en Ciudad Hidalgo se ve inusualmente alto para este tamaño de mercado — se ha visto terreno comercial en la zona con precios de unos cientos a unos miles de pesos por metro cuadrado. Esto es una alerta informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Departamento', referencia: 'Departamento D.F', propietario: 'Jorge Miranda Juarez', direccion: 'Av. Santa Fe #449, Dpto. 1402 Torre B', ciudad: 'Mexico', estado: 'Distrito Federal', cp: '05384', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 400000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'Santa Fe is a major modern high-rise business district in western Mexico City, home to corporate headquarters, malls, and universities, and one of the capital’s more expensive submarkets — average prices there run around $45,000 pesos per square meter in 2026, up roughly 30% over five years, with heavy recent new supply.',
    aiResearchES: 'Santa Fe es un importante distrito de negocios moderno y de rascacielos en el poniente de la Ciudad de México, sede de corporativos, centros comerciales y universidades, y uno de los submercados más caros de la capital — los precios promedio rondan los $45,000 pesos por metro cuadrado en 2026, con un alza de casi 30% en cinco años.',
    aiValueEstimateEN: '$400,000 for an apartment in a Santa Fe tower looks low relative to current pricing, where typical units start closer to $3 million pesos and often run much higher; this figure likely reflects an old or outdated listing price rather than current value. This is only an informal read, not a formal appraisal.',
    aiValueEstimateES: '$400,000 por un departamento en una torre de Santa Fe se ve bajo comparado con los precios actuales, donde las unidades típicas arrancan cerca de $3 millones de pesos y suelen ser mucho más altas; esta cifra probablemente refleje un precio de listado viejo y no el valor actual. Esto es solo una lectura informal, no un avalúo formal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Casa Guanajuato', propietario: 'Georgina M. Lopez', direccion: 'Tenaza #28, Col. San Javier', ciudad: 'Guanajuato', estado: 'Guanajuato', cp: '', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 350000, escrituras: 'Si', propEscriturado: 'Georgina Molina López', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'Guanajuato city is a UNESCO World Heritage colonial capital with a tourism- and university-driven economy, where housing demand is supported by steady visitor interest plus the Universidad de Guanajuato; the city’s listing inventory has grown but remains smaller than nearby Querétaro, which supports relatively firm pricing in established neighborhoods.',
    aiResearchES: 'La ciudad de Guanajuato es una capital colonial Patrimonio de la Humanidad por la UNESCO con una economía impulsada por el turismo y la universidad; el inventario de listados ha crecido pero sigue siendo menor que el de Querétaro, lo que sostiene precios relativamente firmes en colonias establecidas.',
    aiValueEstimateEN: '$350,000 for a house in Colonia San Javier is a plausible, moderate figure for a smaller or older home outside the most tourist-premium blocks of central Guanajuato, where colonial-core properties can command significantly more per square meter. This is only an informal comment, not an appraisal.',
    aiValueEstimateES: '$350,000 por una casa en la Colonia San Javier es una cifra moderada y creíble para una vivienda fuera de las cuadras de mayor plusvalía turística del centro de Guanajuato, donde las propiedades del casco colonial pueden costar bastante más por metro cuadrado. Esto es solo un comentario informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Buzon Purpecha', propietario: 'Jorge Miranda Juarez', direccion: 'Leona Vicario #235, esq. con Josefa Ortiz de Dominguez', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 230000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'Small storefronts along Leona Vicario in Maravatío’s commercial strip serve foot traffic from the town center and are typically occupied by shops, small food businesses, or service counters — modest but steady local commerce rather than tourist-driven retail. This kind of small commercial local is one of the more liquid property types in a town this size.',
    aiResearchES: 'Los pequeños locales comerciales sobre Leona Vicario, en la franja comercial de Maravatío, atienden el tránsito peatonal del centro y suelen ser ocupados por tiendas, negocios de comida o mostradores de servicio — comercio local modesto pero constante. Este tipo de local suele ser uno de los inmuebles más líquidos en un pueblo de este tamaño.',
    aiValueEstimateEN: '$230,000 for a small storefront/local in central Maravatío looks broadly in line with what similar small commercial spaces fetch in this market, where local retail units often trade in the low hundreds of thousands of pesos. This is an informal read, not an appraisal.',
    aiValueEstimateES: '$230,000 por un local comercial pequeño en el centro de Maravatío se ve en línea con lo que se cotizan locales similares en este mercado, donde el retail local suele venderse en cientos de miles de pesos bajos. Esto es una lectura informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Cantalagua', propietario: 'Jorge A. Miranda', direccion: 'Lote #11, Zona D. Fraccionamiento y Club de Golf Hacienda Cantalagua', ciudad: 'Contepec', estado: 'Michoacan', cp: '', observacionesRaw: '', status: 'unspecified', precioEstimadoUSD: 137000, escrituras: 'Si', propEscriturado: 'Jorge A. Miranda', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'Hacienda Cantalagua is an established golf-and-country-club residential development in Contepec, Michoacán, built around a longstanding 18-hole course and hotel dating to the 1970s — a niche, low-density resort-style community rather than a general urban market, with home lots offered from around 400 square meters, roughly two hours from Mexico City.',
    aiResearchES: 'Hacienda Cantalagua es un desarrollo residencial establecido de club de golf en Contepec, Michoacán, construido alrededor de un campo de 18 hoyos y un hotel que datan de los años setenta; es una comunidad de nicho, de baja densidad y estilo resort, con lotes ofrecidos desde alrededor de 400 metros cuadrados.',
    aiValueEstimateEN: '$137,000 for a lot in this development is a plausible figure for an entry-size parcel in a niche golf-community market, though exact pricing depends heavily on lot size and proximity to the course itself. This is an informal comment only.',
    aiValueEstimateES: '$137,000 por un lote en este desarrollo es una cifra creíble para un predio de tamaño de entrada en un mercado de nicho de comunidad de golf, aunque el precio exacto depende mucho del tamaño del lote y su cercanía al campo. Esto es solo un comentario informal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Terreno Autopista', propietario: 'Jorge A. Miranda', direccion: 'Tramo Federal Maravatio-Atlacomulco, Ejido la Huaracha (Juan Francisco)', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: '', status: 'unspecified', precioEstimadoUSD: 80000, escrituras: 'Ejido', propEscriturado: 'Jorge A. Miranda', propuestaTraspaso: 'JORGE A',
    aiResearchEN: 'This parcel sits near the highway corridor around Maravatío, an agricultural region where rural land near major roads can carry a modest premium for visibility or future commercial use, but is held under ejido (communal) tenure — meaning converting to sellable private property (dominio pleno) requires a formal, sometimes lengthy legal process.',
    aiResearchES: 'Este predio está cerca del corredor carretero de Maravatío, una región agrícola donde el terreno junto a vías importantes puede tener un ligero premio por visibilidad o uso comercial futuro, pero está bajo régimen ejidal (comunal) — convertirlo a propiedad privada vendible (dominio pleno) requiere un trámite legal formal, a veces largo.',
    aiValueEstimateEN: '$80,000 for a highway-adjacent ejido parcel is plausible for raw rural land in this area, where unimproved farmland can run well under $1,000 pesos per square meter, but the ejido status is the bigger factor in valuing this correctly. This is purely an informal read.',
    aiValueEstimateES: '$80,000 por un predio ejidal junto a la carretera es una cifra creíble para terreno rural sin urbanizar en esta zona, pero el estatus ejidal es el factor más importante para valorarlo bien — hay que confirmar el estatus legal de traspaso antes de comparar precios. Esto es solo una lectura informal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Oficina Acambaro', propietario: 'Jorge Miranda Juárez, Jorge A. Miranda', direccion: 'Miguel Hidalgo 335-C', ciudad: 'Acambaro', estado: 'Guanajuato', cp: '38600', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 300000, escrituras: 'Si', propEscriturado: 'Jorge A. Miranda', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Acámbaro is a mid-size Guanajuato town with a diversified local economy and an active if modest real estate market, where land and small commercial buildings are frequently listed with negotiable pricing; office space there serves local professional and administrative needs rather than a corporate market.',
    aiResearchES: 'Acámbaro es un pueblo mediano de Guanajuato con una economía local diversificada y un mercado inmobiliario activo aunque modesto, donde el terreno y los pequeños edificios comerciales suelen listarse con precio negociable; el espacio de oficinas ahí atiende necesidades profesionales y administrativas locales.',
    aiValueEstimateEN: '$300,000 for a small office building in Acámbaro looks like a reasonable, mid-range figure for this market, broadly consistent with local small commercial property pricing, though without square footage it’s hard to judge precisely. This is only an informal estimate.',
    aiValueEstimateES: '$300,000 por un pequeño edificio de oficinas en Acámbaro parece una cifra razonable y media para este mercado, aunque sin metros cuadrados es difícil juzgarlo con precisión. Esto es solo una estimación informal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Dpto. Cancun', propietario: 'Jorge Miranda Juárez', direccion: 'Dpto. 16-C Torre II, Condominio Cancun Towers', ciudad: 'Cancun', estado: 'Quintana Roo', cp: '77500', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 750000, escrituras: 'Si', propEscriturado: 'OECC', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Cancún is one of Mexico’s leading resort and tourism markets, with average condo pricing around $30,000 pesos per square meter citywide and up to $38,000-42,000 in premium zones like Puerto Cancún and the Hotel Zone; the market posted roughly 14% appreciation in 2025 and continues to draw both vacation-rental investors and permanent residents.',
    aiResearchES: 'Cancún es uno de los mercados turísticos y de resort líderes de México, con precios promedio de departamentos alrededor de $30,000 pesos por metro cuadrado a nivel ciudad y hasta $38,000-42,000 en zonas premium como Puerto Cancún y la Zona Hotelera; el mercado tuvo una plusvalía de casi 14% en 2025.',
    aiValueEstimateEN: '$750,000 for a condo in a tower like Cancún Towers looks plausible but on the lower end for Cancún’s current market, where a mid-size unit even outside the priciest zones commonly runs into the low millions of pesos; worth confirming unit size and whether the price reflects an older listing. This is an informal read only.',
    aiValueEstimateES: '$750,000 por un departamento en una torre como Cancún Towers parece creíble pero en la parte baja para el mercado actual de Cancún, donde una unidad de tamaño medio incluso fuera de las zonas más caras suele costar millones de pesos bajos; vale la pena confirmar el tamaño y si el precio refleja un listado antiguo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Carmen Lopez', propietario: 'Jorge Miranda Juarez', direccion: 'Leona Vicario 895 con Trece de Septiembre', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 200000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'This house sits on the same Leona Vicario corridor in Maravatío, in a residential stretch of the same small, locally-driven housing market, where plain in-town homes without extra land typically list well under $1 million pesos. Buyer demand here comes mostly from local families rather than investors.',
    aiResearchES: 'Esta casa está sobre el mismo corredor de Leona Vicario en Maravatío, en un tramo residencial del mismo mercado de vivienda pequeño y de demanda local, donde las casas urbanas sencillas sin terreno adicional suelen listarse muy por debajo de $1 millón de pesos.',
    aiValueEstimateEN: '$200,000 for a house on Leona Vicario looks like a modest, plausible figure for a smaller or older home in Maravatío’s centro, consistent with local resale prices for basic housing stock in this size of town. This isn’t a formal appraisal.',
    aiValueEstimateES: '$200,000 por una casa sobre Leona Vicario parece una cifra modesta y creíble para una vivienda más pequeña o antigua en el centro de Maravatío, consistente con precios de reventa locales para vivienda básica en un pueblo de este tamaño. Esto no es un avalúo formal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Local Acambaro 2-A', propietario: 'Georgina M. Lopez', direccion: 'Zaragoza 2-A', ciudad: 'Acambaro', estado: 'Guanajuato', cp: '38600', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 120000, escrituras: 'Si', propEscriturado: 'Georgina Molina López', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'This storefront on Zaragoza street sits within Acámbaro’s compact commercial center, the kind of small retail unit commonly bought by local shopkeepers and service businesses; it sits directly next to a matching adjacent unit (2-B), which adds flexibility since the two can be combined or operated independently.',
    aiResearchES: 'Este local sobre la calle Zaragoza está dentro del centro comercial compacto de Acámbaro, del tipo de local pequeño que suelen comprar comerciantes y negocios de servicio locales; está justo junto a una unidad contigua equivalente (2-B), lo que añade flexibilidad.',
    aiValueEstimateEN: '$120,000 for this small storefront looks reasonable for Acámbaro’s market, where basic commercial locals tend to price in the low hundreds of thousands of pesos. This is only an informal read, not an appraisal.',
    aiValueEstimateES: '$120,000 por este pequeño local parece razonable para el mercado de Acámbaro, donde los locales comerciales básicos suelen tener precios en cientos de miles de pesos bajos. Esto es solo una lectura informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Local Acambaro 2-B', propietario: 'Georgina M. Lopez', direccion: 'Zaragoza 2-B', ciudad: 'Acambaro', estado: 'Guanajuato', cp: '38600', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 120000, escrituras: 'Si', propEscriturado: 'Georgina Molina López', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'This is the adjacent storefront to 2-A on the same stretch of Zaragoza street in Acámbaro’s commercial center, sharing the same small local-retail market dynamics — modest demand from shopkeepers and service businesses rather than large chain retail.',
    aiResearchES: 'Este es el local contiguo al 2-A, sobre el mismo tramo de la calle Zaragoza en el centro comercial de Acámbaro, y comparte la misma dinámica de mercado de retail local pequeño.',
    aiValueEstimateEN: '$120,000 for this second storefront matches its neighboring unit’s price and looks equally reasonable for Acámbaro’s commercial market. This is only an informal read, not an appraisal.',
    aiValueEstimateES: '$120,000 por este segundo local coincide con el precio de la unidad vecina y se ve igualmente razonable para el mercado comercial de Acámbaro. Esto es solo una lectura informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Local', referencia: 'Oficina Salvatierra', propietario: 'Jorge Miranda Juárez', direccion: 'Juarez 215-A', ciudad: 'Salvatierra', estado: 'Guanajuato', cp: '38900', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 90000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Salvatierra is an agricultural Guanajuato town and designated Pueblo Mágico, with its economy centered on diverse crop production and a growing commercial sector; the town’s general real estate market remains modest in scale compared to larger Guanajuato cities.',
    aiResearchES: 'Salvatierra es un pueblo agrícola de Guanajuato y Pueblo Mágico designado, con una economía centrada en producción diversa de cultivos y un sector comercial en crecimiento, aunque el mercado inmobiliario general sigue siendo modesto.',
    aiValueEstimateEN: '$90,000 for a small office space in Salvatierra looks low but plausible for this size of market, where basic commercial units in a modest agricultural town would be expected to price cheaply relative to larger Guanajuato cities. This is only an informal estimate.',
    aiValueEstimateES: '$90,000 por un pequeño espacio de oficina en Salvatierra se ve bajo pero creíble para este tamaño de mercado. Esto es solo una estimación informal.' },

  { countryCode: 'MX', paisRaw: 'México', tipo: 'Terreno', referencia: 'Casa Blanca', propietario: 'Pamela Miranda Molina', direccion: 'Sta Rosa camino a Palomas', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: 34000, escrituras: 'Ejido', propEscriturado: '', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Like other rural parcels around Maravatío, this land is part of the area’s agricultural ejido system, where plots are typically valued for farming or grazing use rather than development, and legal transfer outside the ejido membership requires the dominio pleno conversion process.',
    aiResearchES: 'Como otros predios rurales alrededor de Maravatío, este terreno forma parte del sistema ejidal agrícola de la zona, donde los lotes suelen valorarse para uso de cultivo o pastoreo, y el traspaso legal fuera de los miembros del ejido requiere el trámite de dominio pleno.',
    aiValueEstimateEN: '$34,000 for an ejido land parcel is a small, plausible figure consistent with basic agricultural land pricing in rural Michoacán. This is an informal estimate only — actual value depends heavily on the parcel’s size, access, and whether dominio pleno has been obtained.',
    aiValueEstimateES: '$34,000 por un predio ejidal es una cifra pequeña y creíble, consistente con precios básicos de tierra agrícola en el Michoacán rural. Esto es solo una estimación informal — el valor real depende mucho del tamaño del predio, el acceso, y si ya se obtuvo el dominio pleno.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Yurecuaro', propietario: 'Jorge Miranda Juarez', direccion: 'Yurecuaro', ciudad: 'Yurecuaro', estado: 'Michoacan', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 30000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Yurécuaro is a small agricultural town in western Michoacán near the Jalisco border, part of a farming region historically built on sugarcane and livestock, with a modest, largely local land market. Rural parcels and commercial lots with an ejido pedigree in nearby areas have been seen priced around 650 pesos per square meter.',
    aiResearchES: 'Yurécuaro es un pequeño pueblo agrícola en el occidente de Michoacán, cerca del límite con Jalisco, parte de una región agrícola históricamente basada en la caña de azúcar y la ganadería, con un mercado de tierra modesto y mayormente local.',
    aiValueEstimateEN: '$30,000 for this parcel is a plausible, modest figure for rural land in the Yurécuaro area, consistent with the town’s small local land market. This is only an informal comment, not an appraisal.',
    aiValueEstimateES: '$30,000 por este predio es una cifra modesta y creíble para terreno rural en la zona de Yurécuaro, consistente con el pequeño mercado de tierra local del pueblo. Esto es solo un comentario informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'El Tejero', propietario: 'Jorge Miranda Juarez', direccion: 'El Tejero', ciudad: 'Tejero', estado: 'Michoacan', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 30000, escrituras: 'Ejido', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'El Tejero is a rural locality within Maravatío municipality, in the same agricultural belt as the family’s other Maravatío-area land, and shares that region’s ejido-dominated land tenure, where sale to an outside buyer requires the formal dominio pleno conversion process.',
    aiResearchES: 'El Tejero es una localidad rural dentro del municipio de Maravatío, en el mismo cinturón agrícola que otros predios de la familia en esa zona, y comparte el régimen de tenencia ejidal dominante en la región.',
    aiValueEstimateEN: '$30,000 for this ejido parcel is a plausible, modest figure in line with other small rural land holdings around Maravatío. This is only an informal comment — confirming the dominio pleno/ejido status is more important here than the price itself.',
    aiValueEstimateES: '$30,000 por este predio ejidal es una cifra modesta y creíble, en línea con otros predios rurales pequeños alrededor de Maravatío. Confirmar el estatus de dominio pleno/ejido es más importante aquí que el precio en sí.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Edificio', referencia: 'La Victoria', propietario: 'Jorge Miranda Juarez', direccion: 'Prol. Galeana #1657, Fracc. San Antonio', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'Asunto JMJ', status: 'family_matter', precioEstimadoUSD: 450000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'Prolongación Galeana is a commercial artery on the edge of Maravatío’s built-up area, suited to auto-oriented retail, workshops, or storage rather than dense pedestrian shopping. This specific listing is flagged internally as tied to a family/ownership matter, which matters more to a buyer than general area trends.',
    aiResearchES: 'Prolongación Galeana es una vía comercial en la orilla de la zona urbanizada de Maravatío. Este listado está marcado internamente como ligado a un asunto familiar/de propiedad, lo cual pesa más para un comprador que las tendencias generales de la zona.',
    aiValueEstimateEN: '$450,000 for a commercial building on this corridor is a plausible mid-range figure for Maravatío’s commercial market. Given the noted family/ownership matter, any value read should be treated as secondary to resolving legal title first.',
    aiValueEstimateES: '$450,000 por un edificio comercial en este corredor es una cifra media creíble para el mercado comercial de Maravatío. Dado el asunto familiar/de propiedad señalado, cualquier lectura de valor debe verse como secundaria a resolver primero el título legal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Panuco, Zac', propietario: 'Jorge Miranda Juarez', direccion: 'Zacatecas, Zac.', ciudad: 'Panuco', estado: 'Zacatecas', cp: '', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: null, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'Pánuco, Zacatecas is a rural farming and ranching municipality ranking among the state’s top agricultural producers, with alfalfa and forage corn as leading crops and much of its land organized under the ejido communal system.',
    aiResearchES: 'Pánuco, Zacatecas es un municipio rural agrícola y ganadero que figura entre los principales productores agrícolas del estado, con la alfalfa y el maíz forrajero como cultivos principales, y buena parte de su tierra bajo el sistema ejidal.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible. Informally, plain agricultural land in a rural Zacatecas municipality like Pánuco tends to be inexpensive, especially if it carries ejido status requiring conversion to sell freely.',
    aiValueEstimateES: 'Sin precio listado, no es posible una comparación directa. De forma informal, el terreno agrícola sencillo en un municipio rural de Zacatecas como Pánuco suele ser barato, especialmente si tiene estatus ejidal.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Funera Express', propietario: 'Pamela Miranda Molina', direccion: 'Ejido Tzinzingareo, Municipio de Irimbo', ciudad: 'Tzinzingareo', estado: 'Michoacan', cp: '', observacionesRaw: '', status: 'unspecified', precioEstimadoUSD: null, escrituras: '', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: 'PAMELA',
    aiResearchEN: 'Irimbo and the Tzinzingareo area sit in Michoacán’s agricultural highlands near Maravatío, a region dominated by small-scale farming, forestry, and increasingly avocado cultivation; land here is commonly held under ejido tenure, requiring the dominio pleno conversion process to sell.',
    aiResearchES: 'Irimbo y la zona de Tzinzingareo se ubican en las tierras altas agrícolas de Michoacán, cerca de Maravatío, una región dominada por agricultura de pequeña escala, silvicultura y, cada vez más, cultivo de aguacate; el terreno aquí comúnmente está bajo régimen ejidal.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible. Informally, rural ejido land in this part of Michoacán tends to be inexpensive per hectare unless it has water access or avocado-orchard potential, which can raise value substantially.',
    aiValueEstimateES: 'Sin precio listado, no es posible una comparación directa. De forma informal, el terreno ejidal rural en esta parte de Michoacán suele ser barato por hectárea, a menos que tenga acceso a agua o potencial para huerta de aguacate.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Terreno Manzanillo', propietario: 'Jorge, Francisco, Toño', direccion: 'Manzanillo, Col.', ciudad: 'Manzanillo', estado: 'Colima', cp: '', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 950000, escrituras: 'Pendientes', propEscriturado: '', propuestaTraspaso: 'VENDER',
    aiResearchEN: 'Manzanillo is a Pacific coast port and resort city in Colima, where land is in demand both for its working port/logistics economy and for residential and tourism development along the coastline; residential ocean-view land has been seen priced around $7,000 pesos per square meter in desirable areas.',
    aiResearchES: 'Manzanillo es una ciudad portuaria y turística del Pacífico en Colima, donde el terreno tiene demanda tanto por su economía portuaria/logística activa como por el desarrollo residencial y turístico costero; se ha visto terreno con vista al mar con precios de alrededor de $7,000 pesos por metro cuadrado.',
    aiValueEstimateEN: '$950,000 for land in Manzanillo is plausible depending on size and proximity to the coast, though shared partial ownership among siblings (Jorge, Francisco, Toño) adds complexity that matters more than the raw price — any sale would need agreement among all co-owners first.',
    aiValueEstimateES: '$950,000 por terreno en Manzanillo es creíble dependiendo del tamaño y cercanía a la costa, aunque la propiedad parcial compartida entre hermanos (Jorge, Francisco, Toño) agrega complejidad que importa más que el precio.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Predio Acambaro', propietario: 'Consorcio Inmobiliario JMJ', direccion: 'Predio Rustico CERRO DE LA CRUZ', ciudad: 'Acambaro', estado: 'Guanajuato', cp: '38600', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 364000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'Cerro de la Cruz is a rural/hillside area on Acámbaro’s outskirts, typically used for agriculture, grazing, or eventual low-density development rather than dense urban use; local land listings in and around Acámbaro range widely, with subdivided residential parcels seen around $700 pesos per square meter.',
    aiResearchES: 'Cerro de la Cruz es una zona rural/de cerro en las afueras de Acámbaro, típicamente usada para agricultura, pastoreo o desarrollo de baja densidad; los listados de terreno alrededor de Acámbaro varían bastante, con lotes residenciales subdivididos vistos alrededor de $700 pesos por metro cuadrado.',
    aiValueEstimateEN: '$364,000 for rustic land depends heavily on parcel size, which isn’t specified here — at typical rural land rates this could represent anywhere from a modest plot to several hectares. This is only an informal comment, not an appraisal.',
    aiValueEstimateES: '$364,000 por terreno rústico depende mucho del tamaño del predio, que no se especifica aquí. Esto es solo un comentario informal, no un avalúo.' },

  { countryCode: 'MX', paisRaw: '', tipo: 'Terreno', referencia: 'Rancho Grande', propietario: 'Jorge Miranda Juárez', direccion: 'Salida Salvatierra', ciudad: 'Acambaro', estado: 'Guanajuato', cp: '', observacionesRaw: '', status: 'unspecified', precioEstimadoUSD: 200000, escrituras: '', propEscriturado: 'en tramite Escritura', propuestaTraspaso: '',
    aiResearchEN: 'This parcel sits on the road out toward Salvatierra, in the agricultural belt connecting Acámbaro to the wider Bajío farming region; land along this corridor is typically valued for farming use, with any development potential tied to future road-frontage commercial interest.',
    aiResearchES: 'Este predio está sobre la salida hacia Salvatierra, en el cinturón agrícola que conecta Acámbaro con la región del Bajío; el terreno sobre este corredor se valora típicamente para uso agrícola.',
    aiValueEstimateEN: '$200,000 for rural land here is a plausible mid-range figure, but the pending deed status ("en trámite") is more important — a buyer should treat the price as provisional until the title process is resolved.',
    aiValueEstimateES: '$200,000 por terreno rural aquí es una cifra media creíble, pero el estatus de escritura "en trámite" es más importante — un comprador debería tratar el precio como provisional hasta resolver el trámite de título.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Jungapeo', propietario: 'Jorge Miranda Juárez', direccion: 'La Garita, Lazaro Cardenas', ciudad: 'Jungapeo', estado: 'Michoacan', cp: '', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 180000, escrituras: 'Proceso', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'Jungapeo sits in Michoacán’s Tierra Caliente region, an agricultural area increasingly known for mango and avocado production, including export-qualified avocado orchards on ejido land; La Garita falls within this same rural farming context, where land value is closely tied to crop potential and water access.',
    aiResearchES: 'Jungapeo se ubica en la región de Tierra Caliente de Michoacán, una zona agrícola cada vez más conocida por la producción de mango y aguacate; La Garita está dentro de este mismo contexto agrícola rural.',
    aiValueEstimateEN: '$180,000 for rural land in this area is plausible if the parcel has decent farming potential, particularly for avocado — but the deed being "in process" means legal title isn’t finalized, which should factor into the price more than area trends.',
    aiValueEstimateES: '$180,000 por terreno rural en esta zona es creíble si el predio tiene buen potencial agrícola, particularmente para aguacate — pero que la escritura esté "en proceso" significa que el título no está finalizado.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'INFONAVIT', propietario: 'Jorge Miranda Juárez', direccion: 'And. Onesimo Lopez Couto #86, Fracc. Francisco J. Mujica', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 23000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'INFONAVIT-financed housing is Mexico’s mass mortgage program for salaried workers, and small units originally built under it are common starter homes in towns like Maravatío; nationwide, INFONAVIT-financed homes average close to $1.86 million pesos, but that reflects new financed purchases, not resale of small older units.',
    aiResearchES: 'La vivienda financiada por INFONAVIT es el programa hipotecario masivo de México para trabajadores asalariados, y las unidades pequeñas construidas bajo este esquema son viviendas de entrada comunes en pueblos como Maravatío.',
    aiValueEstimateEN: '$23,000 for a small INFONAVIT-program house is very low even by resale standards, closer to the range seen for INFONAVIT’s own foreclosed/recovered ("recuperada") housing stock. Worth checking whether this unit has a bank lien or came through a recovery/auction process.',
    aiValueEstimateES: '$23,000 por una casa pequeña del programa INFONAVIT es muy bajo incluso para estándares de reventa, más cercano al rango de las viviendas "recuperadas" de INFONAVIT. Vale la pena revisar si tiene algún gravamen bancario o si vino de un proceso de recuperación/remate.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'San Nicolas', propietario: 'Jorge Miranda Juárez', direccion: 'Salazar 204 Col. Centro', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: '', status: 'unspecified', precioEstimadoUSD: null, escrituras: '', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'This house is on Salazar street, within Maravatío’s same modest local housing market — a small city where in-town homes without significant land typically sell well under $1 million pesos, with demand coming mainly from local families.',
    aiResearchES: 'Esta casa está sobre la calle Salazar, dentro del mismo mercado de vivienda local y modesto de Maravatío, con demanda principalmente de familias locales.',
    aiValueEstimateEN: 'No price is listed for this property. Based on the general Maravatío market, a plain in-town house on a street like Salazar would informally be expected to fall in the low hundreds of thousands of pesos, but this is a rough guess only.',
    aiValueEstimateES: 'No hay precio listado para esta propiedad. Con base en el mercado general de Maravatío, informalmente se esperaría un rango de cientos de miles de pesos bajos, pero esto es solo un cálculo aproximado.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'Puebla', propietario: 'Jorge Miranda Juarez', direccion: 'Calle 5 Sur, Col. el Humilladero', ciudad: 'Tehuacan', estado: 'Puebla', cp: '', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 14000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'Tehuacán is a mid-size industrial city in Puebla known historically for textiles and bottled mineral water, with an active land market along its federal highway corridors; industrial-zoned land there has been seen priced around $500 pesos per square meter, though smaller in-town residential lots price much lower.',
    aiResearchES: 'Tehuacán es una ciudad industrial mediana de Puebla conocida históricamente por sus textiles y su agua mineral embotellada, con un mercado de terreno activo a lo largo de sus corredores carreteros federales.',
    aiValueEstimateEN: '$14,000 for a small lot in Colonia El Humilladero is a modest, plausible figure for a small residential parcel in a lower-cost neighborhood of Tehuacán, well below the pricing seen for larger industrial-zoned land along the highway.',
    aiValueEstimateES: '$14,000 por un lote pequeño en la Colonia El Humilladero es una cifra modesta y creíble para un predio residencial pequeño en una colonia de menor costo de Tehuacán.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Casa', referencia: 'Francisco I Madero', propietario: 'Jorge Miranda Juarez', direccion: 'Venustiano Carranza 429, Frac. Francisco I. Madero', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: 12000, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'This is a house on Venustiano Carranza street, within Maravatío’s small local housing market where basic in-town homes are priced modestly relative to larger cities, with demand driven by local buyers.',
    aiResearchES: 'Esta es una casa sobre la calle Venustiano Carranza, dentro del mercado de vivienda local y pequeño de Maravatío, con demanda impulsada por compradores locales.',
    aiValueEstimateEN: '$12,000 is an extremely low figure for a full house, even in a low-cost market like Maravatío — it likely reflects a very small/deteriorated structure, an outdated price, a partial interest in the property, or possibly a data entry issue. Worth verifying before relying on it.',
    aiValueEstimateES: '$12,000 es una cifra extremadamente baja para una casa completa, incluso en un mercado de bajo costo como Maravatío — probablemente refleje una estructura muy pequeña o deteriorada, un precio desactualizado, o un error de captura. Vale la pena verificar la cifra.' },

  { countryCode: 'MX', paisRaw: '', tipo: 'Terreno', referencia: 'Inversion Qro.', propietario: 'Centro Comercial Tecnologico 998 S.A de C.V', direccion: 'Lot. 6, Manzana ZRUM 05, Fracc. Zaru', ciudad: 'Queretaro', estado: 'Queretaro', cp: '', observacionesRaw: 'Inversion', status: 'investment', precioEstimadoUSD: 45000, escrituras: '', propEscriturado: '', propuestaTraspaso: '',
    aiResearchEN: 'Querétaro city is one of Mexico’s fastest-growing industrial and tech hubs, benefiting strongly from nearshoring-driven manufacturing investment, with land and housing in developing subdivisions ("fraccionamientos") appreciating around 8-12% annually in growth areas.',
    aiResearchES: 'La ciudad de Querétaro es uno de los polos industriales y tecnológicos de más rápido crecimiento en México, beneficiada por la inversión manufacturera impulsada por el nearshoring, con terreno en fraccionamientos apreciándose entre 8% y 12% anual.',
    aiValueEstimateEN: '$45,000 for an investment land parcel in a development like Fracc. Zaru sounds like an entry-level lot price, plausible for a smaller or pre-development parcel in Querétaro’s fast-growing periphery, though established, closer-in subdivisions command considerably more.',
    aiValueEstimateES: '$45,000 por un predio de inversión en un desarrollo como Fracc. Zaru suena a un precio de lote de entrada, creíble para un predio más pequeño o en preventa en la periferia de rápido crecimiento de Querétaro.' },

  { countryCode: 'MX', paisRaw: 'Mexico', tipo: 'Terreno', referencia: 'max', propietario: 'Jorge Miranda Juárez', direccion: 'Comonfort 13-B Centro', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: null, escrituras: 'Si', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'This parcel on Comonfort street sits within Maravatío’s small urban core, where in-town lots are valued based on location, size, and utility access, and land without a structure typically trades cheaper than developed property nearby.',
    aiResearchES: 'Este predio sobre la calle Comonfort está dentro del núcleo urbano pequeño de Maravatío, donde los lotes se valoran según ubicación, tamaño y acceso a servicios.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible. Informally, a small in-town lot in Maravatío’s centro would be expected to fall in the tens of thousands of pesos depending on size and services, but this is a rough guess only.',
    aiValueEstimateES: 'Sin precio listado, no es posible una comparación directa. De forma informal, se esperaría un rango de decenas de miles de pesos según tamaño y servicios, pero esto es solo un cálculo aproximado.' },

  { countryCode: 'MX', paisRaw: '', tipo: 'Terreno', referencia: 'El Salto', propietario: 'Jorge Miranda Juárez', direccion: 'El Gigante', ciudad: 'Maravatio', estado: 'Michoacan', cp: '61250', observacionesRaw: 'En Uso', status: 'in_use', precioEstimadoUSD: null, escrituras: 'Ejido', propEscriturado: 'Jorge Miranda Juarez', propuestaTraspaso: '',
    aiResearchEN: 'El Salto is rural land in the Maravatío area held under ejido tenure, part of the same communal agricultural land system as the region’s other ejido parcels — valued mainly for farming or grazing, with any sale requiring conversion to private title (dominio pleno).',
    aiResearchES: 'El Salto es terreno rural en la zona de Maravatío bajo régimen ejidal, parte del mismo sistema de tierra agrícola comunal que otros predios de la región.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible. Informally, unconverted ejido farmland in this region tends to sell cheaply, often in the tens of thousands of pesos — but the ejido status itself should be confirmed before assuming the land is freely transferable.',
    aiValueEstimateES: 'Sin precio listado, no es posible una comparación directa. De forma informal, la tierra ejidal sin convertir en esta región suele venderse barata, a menudo en decenas de miles de pesos.' },

  { countryCode: 'MX', paisRaw: '', tipo: 'Terreno', referencia: 'Torreón La Perla', propietario: 'Jorge Miranda Juárez', direccion: 'Torreon', ciudad: 'Torreon', estado: 'Coahuila', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: null, escrituras: 'Si', propEscriturado: 'IEX', propuestaTraspaso: '',
    aiResearchEN: 'Torreón is a major industrial city in the Comarca Lagunera region of Coahuila, with an economy built on manufacturing, agriculture, commerce, and strong highway/rail logistics infrastructure; land there is available across residential, commercial, and industrial categories at prices generally considered attractive relative to other Mexican industrial hubs.',
    aiResearchES: 'Torreón es una ciudad industrial importante en la Comarca Lagunera de Coahuila, con una economía basada en manufactura, agricultura, comercio, e infraestructura logística sólida de carreteras y ferrocarril.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible for La Perla. Informally, given Torreón’s active and comparatively affordable land market, a parcel here could range from modest to solidly priced depending on zoning and location relative to logistics corridors.',
    aiValueEstimateES: 'Sin precio listado, no es posible una comparación directa para La Perla. De forma informal, dado el mercado de terreno activo y accesible de Torreón, un predio aquí podría ir de modesto a bien cotizado según la zonificación.' },

  { countryCode: 'MX', paisRaw: '', tipo: 'Terreno', referencia: 'Torreón Monte Bello', propietario: 'Jorge Miranda Juárez', direccion: 'Torreon', ciudad: 'Torreon', estado: 'Coahuila', cp: '61250', observacionesRaw: 'VENDER', status: 'for_sale', precioEstimadoUSD: null, escrituras: 'Si', propEscriturado: 'IEX', propuestaTraspaso: '',
    aiResearchEN: 'Monte Bello, like other Torreón-area parcels, sits within the Comarca Lagunera’s industrial and agricultural economy, a region with strong highway and rail connectivity that supports steady demand for land across residential, commercial, and industrial uses.',
    aiResearchES: 'Monte Bello, como otros predios de la zona de Torreón, se ubica dentro de la economía industrial y agrícola de la Comarca Lagunera, una región con fuerte conectividad de carreteras y ferrocarril.',
    aiValueEstimateEN: 'With no listed price, no direct comparison is possible for Monte Bello either. As with La Perla, an informal expectation would tie a Torreón-area parcel’s value to zoning and proximity to logistics infrastructure.',
    aiValueEstimateES: 'Sin precio listado, tampoco es posible una comparación directa para Monte Bello. Igual que con La Perla, una expectativa informal ligaría el valor a la zonificación y la cercanía a infraestructura logística.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Terreno', referencia: 'Estacionamiento P443', propietario: 'J&G Properties', direccion: '501 N. Clinton St., Parking #P443', ciudad: 'Chicago', estado: 'Illinois', cp: '60654', observacionesRaw: 'STATE PLANNING', status: 'unspecified', precioEstimadoUSD: 30000, pin: '17-09-112-107-1439',
    aiResearchEN: '501 N. Clinton sits in River North, where dedicated condo parking spaces are commonly bought and sold as separate deeded units apart from residential condos, reflecting the neighborhood’s dense urban footprint and limited street parking; demand for such units tracks the broader River North condo market.',
    aiResearchES: '501 N. Clinton está en River North, donde los espacios de estacionamiento de condominio se compran y venden comúnmente como unidades escrituradas separadas de los condominios residenciales.',
    aiValueEstimateEN: '$30,000 for a deeded parking spot in River North is a plausible, moderate figure for this type of unit — Chicago high-rise parking spaces in this neighborhood have historically traded anywhere from the high five figures to low six figures. This is only an informal comment, not an appraisal.',
    aiValueEstimateES: '$30,000 dólares por un espacio de estacionamiento escriturado en River North es una cifra creíble y moderada para este tipo de unidad. Esto es solo un comentario informal, no un avalúo.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Edificio', referencia: 'Edificio J.A', propietario: 'Jorge A. Miranda', direccion: '669 W. Ohio St', ciudad: 'Chicago', estado: 'Illinois', cp: '60654', observacionesRaw: 'JORGE A', status: 'unspecified', precioEstimadoUSD: 800000, pin: '17-09-103-002-0000',
    aiResearchEN: 'This address sits in River North, a dense, high-amenity Near North Side neighborhood known for its mix of commercial buildings, condo towers, and nightlife/dining. Overall neighborhood pricing reached a median around $450,000 in early 2026, up about 3.4% year over year.',
    aiResearchES: 'Esta dirección está en River North, un vecindario denso y de muchas amenidades del Near North Side, conocido por su mezcla de edificios comerciales, torres de condominios y vida nocturna/restaurantes.',
    aiValueEstimateEN: '$800,000 for a building on Ohio Street is plausible for a smaller commercial or mixed-use building in River North, though the figure sits well above the neighborhood’s residential median, which makes sense if this is a full building rather than a single condo unit.',
    aiValueEstimateES: '$800,000 dólares por un edificio en Ohio Street es creíble para un edificio comercial o de uso mixto más pequeño en River North, aunque la cifra está por arriba de la mediana residencial del vecindario.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Edificio', referencia: 'Edificio Corporativo', propietario: 'J&G Services', direccion: '685 W. Ohio St.', ciudad: 'Chicago', estado: 'Illinois', cp: '60654', observacionesRaw: 'STATE PLANNING', status: 'unspecified', precioEstimadoUSD: 2000000,
    aiResearchEN: 'Also on the Ohio Street corridor in River North, this location benefits from the same dense mixed commercial-residential setting, close to major expressway access, though the broader office market’s soft vacancy is a relevant headwind for larger corporate-style buildings specifically.',
    aiResearchES: 'También sobre el corredor de Ohio Street en River North, esta ubicación se beneficia del mismo entorno denso de uso mixto comercial-residencial, cerca de accesos a autopistas importantes.',
    aiValueEstimateEN: '$2,000,000 for a corporate-style building at this address is a plausible figure for a larger commercial building in River North. This is only an informal comment, not an appraisal — condition, square footage, and lease-up status would move the real figure substantially.',
    aiValueEstimateES: '$2,000,000 dólares por un edificio de tipo corporativo en esta dirección es una cifra creíble para un edificio comercial más grande en River North. Esto es solo un comentario informal, no un avalúo.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Edificio', referencia: 'Edificio Oficina Aurora', propietario: 'J&G Properties', direccion: '22 N Union St', ciudad: 'Aurora', estado: 'Illinois', cp: '60505', observacionesRaw: 'STATE PLANNING (source data listed the city as Chicago; corrected to Aurora, IL to match the address)', status: 'unspecified', precioEstimadoUSD: 550000,
    aiResearchEN: 'This address is in downtown Aurora, Illinois, a Chicago-area suburb whose historic downtown has the city’s largest concentration of office listings and is undergoing active revitalization; commercial rents in Aurora citywide average around $20 per square foot annually, well below downtown Chicago pricing.',
    aiResearchES: 'Esta dirección está en el centro de Aurora, Illinois, un suburbio del área de Chicago cuyo centro histórico tiene la mayor concentración de oficinas en renta de la ciudad y está en revitalización activa.',
    aiValueEstimateEN: '$550,000 for an office building in downtown Aurora is a plausible figure for a small-to-mid-size commercial building in this suburban market, which prices well below Chicago proper but benefits from ongoing downtown investment.',
    aiValueEstimateES: '$550,000 dólares por un edificio de oficinas en el centro de Aurora es una cifra creíble para un edificio comercial pequeño a mediano en este mercado suburbano.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Departamento', referencia: 'Dpto. Prairie', propietario: 'Georgina M. Lopez', direccion: '1211 S. Prairie St, Unit 2201', ciudad: 'Chicago', estado: 'Illinois', cp: '60605', observacionesRaw: 'STATE PLANNING', status: 'unspecified', precioEstimadoUSD: 1600000,
    aiResearchEN: '1211 S. Prairie is One Museum Park East, a well-known high-rise condo tower in Chicago’s South Loop, completed in 2007 near the lakefront and Museum Campus; higher floors with skyline or lake views in this building command a premium over lower or interior-facing units.',
    aiResearchES: '1211 S. Prairie es One Museum Park East, una conocida torre de condominios de gran altura en el South Loop de Chicago, terminada en 2007 cerca del lago y del Museum Campus.',
    aiValueEstimateEN: '$1,600,000 for Unit 2201 (a high floor) at One Museum Park East is plausible for a premium, view-oriented unit in this well-regarded tower, since higher floors in the building have historically carried a notable price step-up over typical units.',
    aiValueEstimateES: '$1,600,000 dólares por la Unidad 2201 (un piso alto) en One Museum Park East es creíble para una unidad premium orientada a la vista en esta torre bien valorada.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Casa', referencia: 'Casa Willow', propietario: 'Alemap, LLC', direccion: '(sold — street address not recorded in source data)', ciudad: 'Chicago', estado: 'Illinois', cp: '60605', observacionesRaw: 'Marked "vendida" (sold) in the original source data', status: 'sold', precioEstimadoUSD: 750000, archived: true, archivedReason: "Sold ('vendida' in source data) — archived at migration",
    aiResearchEN: 'This property was marked as sold in the family’s original records. No further area research was carried out since it is no longer part of the active portfolio.',
    aiResearchES: 'Esta propiedad estaba marcada como vendida en los registros originales de la familia. No se realizó más investigación de la zona porque ya no forma parte del portafolio activo.',
    aiValueEstimateEN: 'Not applicable — the property was sold; the $750,000 figure carried over from the source file reflects historical record-keeping, not a current estimate.',
    aiValueEstimateES: 'No aplica — la propiedad fue vendida; la cifra de $750,000 que se conserva del archivo original refleja un registro histórico, no una estimación actual.' },

  { countryCode: 'US', paisRaw: 'Estados Unidos', tipo: 'Terreno', referencia: 'Estacionamiento P52', propietario: 'J&G Properties', direccion: '501 N. Clinton St., Parking #P52', ciudad: 'Chicago', estado: 'Illinois', cp: '60654', observacionesRaw: 'STATE PLANNING', status: 'unspecified', precioEstimadoUSD: 45000, pin: '17-09-112-107-1254',
    aiResearchEN: 'Like other deeded parking units at 501 N. Clinton in River North, this spot is part of a market where garage spaces trade separately from condo units, with pricing shaped by garage type (self-park vs. valet), floor level, and spot size within the building.',
    aiResearchES: 'Como otras unidades de estacionamiento escrituradas en 501 N. Clinton en River North, este espacio forma parte de un mercado donde los lugares de garaje se venden por separado de las unidades de condominio.',
    aiValueEstimateEN: '$45,000 for a deeded parking spot at the same address is on the higher end of typical River North parking-unit pricing, though not unusual for a larger space or one in a more convenient garage location.',
    aiValueEstimateES: '$45,000 dólares por un espacio de estacionamiento escriturado en la misma dirección está en la parte alta de lo típico para River North, aunque no es inusual para un espacio más grande.' }
];

// Normally never needed - ensureRegistry_() calls seedIntoRegistry_() itself the moment the
// registry is first created (see above), so a fresh deploy seeds itself on first open with
// no manual step. This is kept as an admin-callable fallback (e.g. the registry was created
// before SEED_PROPERTIES_ was filled in, or a schema needs a manual redo of the SEED_DONE
// flag first).
function seedProperties_() {
  requireAdmin_();
  if (prop_('SEED_DONE') === 'true') {
    Logger.log('Seed already run - SEED_DONE is true. Not re-seeding.');
    return { seeded: 0, alreadyDone: true };
  }
  if (!SEED_PROPERTIES_.length) throw new Error('SEED_PROPERTIES_ is empty - nothing to seed.');
  var n = seedIntoRegistry_(ss_(), getUserEmail());
  bustReg_();
  return { seeded: n };
}
