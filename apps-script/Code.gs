/**
 * Recipe Box — Google Apps Script backend
 * ------------------------------------------------------------
 * Lives inside its own Google Sheet (Extensions → Apps Script).
 * Completely separate from the MMT dashboard.
 *
 * Sheet tabs (created by running setup() once):
 *   Recipes      one row per recipe (ingredients + steps readable as text)
 *   Ingredients  your ingredient database (notes, substitutions, aisle)
 *   Cook Log     notes, adjustments and edit history per recipe
 *   Shopping     the shopping list
 *
 * Settings live in Script Properties (use the "Recipe Box" menu in the Sheet):
 *   APP_PIN            4-digit PIN the app asks for. If it isn't set, the first PIN typed
 *                      into the app becomes the PIN.
 *   ANTHROPIC_API_KEY  optional — turns on photo, social-video and smart text imports
 *   CLAUDE_MODEL       optional — defaults to claude-sonnet-5-5
 */

const BACKEND_VERSION = '11.08 v2';
const DEFAULT_MODEL = 'claude-sonnet-5-5';

const SHEETS = {
  Recipes: ['id', 'title', 'description', 'source_type', 'source_url', 'photo', 'servings',
    'prep_time', 'cook_time', 'total_time', 'tags', 'favorite', 'rating',
    'ingredients', 'steps', 'notes', 'times_cooked', 'last_cooked',
    'created', 'updated', 'ingredients_data', 'steps_data', 'original'],
  Ingredients: ['id', 'name', 'aliases', 'category', 'notes', 'substitutions', 'created', 'updated'],
  'Cook Log': ['id', 'recipe_id', 'recipe_title', 'date', 'type', 'text', 'changes', 'rating', 'created'],
  Shopping: ['id', 'name', 'qty', 'category', 'ingredient_id', 'recipes', 'checked', 'added', 'sort']
};

// Columns that hold JSON (kept at the far right of the sheet so the readable columns come first).
const JSON_COLS = { ingredients_data: 1, steps_data: 1, original: 1, recipes: 1, changes: 1 };

// ════════════════════════════════════════════════════════════
// SHEET MENU + ONE-TIME SETUP
// ════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Recipe Box')
    .addItem('1. Set up tabs', 'setup')
    .addItem('2. Set app PIN', 'menuSetPin')
    .addItem('3. Set Claude API key (optional)', 'menuSetKey')
    .addSeparator()
    .addItem('Check settings', 'menuStatus')
    .addToUi();
}

/** The data Sheet: the one this script is attached to, or one it makes for itself. */
function book() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create('McAfee Cookbook Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function setup() {
  const ss = book();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SHEETS[name];
    const existing = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    // Add any missing headers on the right; never reorder or delete what's there.
    const missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
    if (!existing.filter(String).length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (missing.length) {
      sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  });
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  Logger.log('Cookbook data Sheet: ' + ss.getUrl());
  try { SpreadsheetApp.getUi().alert('Recipe Box tabs are ready.'); } catch (e) {}
}

function menuSetPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('App PIN', 'Enter a 4-digit PIN for the Recipe Box app:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (!/^\d{4}$/.test(pin)) { ui.alert('The PIN must be exactly 4 digits.'); return; }
  PropertiesService.getScriptProperties().setProperty('APP_PIN', pin);
  ui.alert('PIN saved.');
}

function menuSetKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Claude API key',
    'Paste your Anthropic API key (starts with sk-ant-). Leave blank to remove it.\n' +
    'It turns on photo, social-video and smart text imports.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (key) props.setProperty('ANTHROPIC_API_KEY', key); else props.deleteProperty('ANTHROPIC_API_KEY');
  ui.alert(key ? 'API key saved.' : 'API key removed.');
}

function menuStatus() {
  const p = PropertiesService.getScriptProperties();
  SpreadsheetApp.getUi().alert(
    'Backend ' + BACKEND_VERSION + '\n\n' +
    'PIN: ' + (p.getProperty('APP_PIN') ? 'set' : 'NOT SET') + '\n' +
    'Claude API key: ' + (p.getProperty('ANTHROPIC_API_KEY') ? 'set' : 'not set (photo/social imports off)') + '\n' +
    'Model: ' + (p.getProperty('CLAUDE_MODEL') || DEFAULT_MODEL));
}

// ════════════════════════════════════════════════════════════
// REQUEST ROUTING
// ════════════════════════════════════════════════════════════

function doGet() {
  return out({ ok: true, app: 'recipe-box', version: BACKEND_VERSION });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return out({ ok: false, error: 'Bad request' }); }

  const props = PropertiesService.getScriptProperties();
  let pin = props.getProperty('APP_PIN');
  let pinCreated = false;
  if (!pin) {
    // First run: the first 4-digit PIN entered in the app becomes the PIN.
    if (data.action === 'ping' && /^\d{4}$/.test(String(data.pin || ''))) {
      pin = String(data.pin); props.setProperty('APP_PIN', pin); pinCreated = true;
      try { setup(); } catch (e) {}
    } else {
      return out({ ok: false, error: 'Pick a 4-digit PIN to finish setting up.', code: 'pin' });
    }
  }
  const cache = CacheService.getScriptCache();
  const fails = +(cache.get('pin_fails') || 0);
  if (fails >= 10) return out({ ok: false, error: 'Too many wrong PINs. Try again in 15 minutes.', code: 'locked' });
  if (String(data.pin || '') !== pin) {
    cache.put('pin_fails', String(fails + 1), 900);
    return out({ ok: false, error: 'Wrong PIN', code: 'pin' });
  }
  if (fails) cache.remove('pin_fails');

  try {
    switch (data.action) {
      case 'ping':              return out({ ok: true, version: BACKEND_VERSION, ai: !!aiKey(), pinCreated: pinCreated });
      case 'load':              return out(loadAll());
      case 'save_recipe':       return out(withLock(function () { return saveRecipe(data); }));
      case 'delete_recipe':     return out(withLock(function () { return deleteRecipe(data.id); }));
      case 'save_ingredients':  return out(withLock(function () { return saveIngredients(data.items || []); }));
      case 'delete_ingredient': return out(withLock(function () { return deleteRows('Ingredients', [data.id]); }));
      case 'save_log':          return out(withLock(function () { upsert('Cook Log', [data.entry]); return bumpCooked(data.entry, data.recipe); }));
      case 'delete_log':        return out(withLock(function () { return deleteRows('Cook Log', [data.id]); }));
      case 'shop_ops':          return out(withLock(function () { return shopOps(data.ops || []); }));
      case 'import_url':        return out(importUrl(String(data.url || '').trim()));
      case 'ai_parse':          return out(aiParse(data));
      case 'upload_photo':      return out(uploadPhoto(data));
      default:                  return out({ ok: false, error: 'Unknown action: ' + data.action });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════
// SHEET HELPERS (rows ↔ objects, keyed by header name)
// ════════════════════════════════════════════════════════════

function sheet(name) {
  const ss = book();
  let sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}

function headersOf(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function readAll(name) {
  const sh = sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = headersOf(sh);
  const values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return values.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    headers.forEach(function (h, i) {
      if (!h) return;
      let v = r[i];
      if (v instanceof Date) v = v.toISOString();
      if (JSON_COLS[h]) { try { v = v ? JSON.parse(v) : null; } catch (e) { v = null; } }
      o[h] = v;
    });
    return o;
  });
}

function toCell(h, v) {
  if (v === undefined || v === null) return '';
  if (JSON_COLS[h]) return typeof v === 'string' ? v : JSON.stringify(v);
  if (typeof v === 'boolean') return v;
  if (Array.isArray(v)) return v.join(', ');
  const s = String(v);
  return s.length > 49000 ? s.slice(0, 49000) : s;
}

/** Insert or replace rows by id. Only the fields given are written; others are kept. */
function upsert(name, objs) {
  objs = (objs || []).filter(function (o) { return o && o.id; });
  if (!objs.length) return { ok: true, n: 0 };
  const sh = sheet(name);
  const headers = headersOf(sh);
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const appends = [];
  objs.forEach(function (o) {
    const idx = ids.indexOf(String(o.id));
    if (idx > -1) {
      const rowNum = idx + 2;
      const row = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      headers.forEach(function (h, i) { if (h && Object.prototype.hasOwnProperty.call(o, h)) row[i] = toCell(h, o[h]); });
      sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
    } else {
      appends.push(headers.map(function (h) { return h ? toCell(h, o[h]) : ''; }));
      ids.push(String(o.id));
    }
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
  return { ok: true, n: objs.length };
}

function deleteRows(name, idList) {
  const sh = sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, n: 0 };
  const want = {};
  idList.forEach(function (id) { want[String(id)] = 1; });
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  let n = 0;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (want[String(ids[i][0])]) { sh.deleteRow(i + 2); n++; }
  }
  return { ok: true, n: n };
}

// ════════════════════════════════════════════════════════════
// DATA ACTIONS
// ════════════════════════════════════════════════════════════

function loadAll() {
  return {
    ok: true,
    version: BACKEND_VERSION,
    ai: !!aiKey(),
    recipes: readAll('Recipes'),
    ingredients: readAll('Ingredients'),
    log: readAll('Cook Log'),
    shopping: readAll('Shopping')
  };
}

function saveRecipe(data) {
  const r = data.recipe;
  if (!r || !r.id) throw new Error('Recipe is missing');
  // Readable copies for the Sheet; the app reads the *_data columns.
  r.ingredients = (r.ingredients_data || []).map(function (l) {
    return l.section ? '— ' + l.section + ' —' : l.text;
  }).join('\n');
  r.steps = (r.steps_data || []).map(function (s) {
    return s.section ? '— ' + s.section + ' —' : s.text;
  }).join('\n');
  if (data.ingredients && data.ingredients.length) upsert('Ingredients', data.ingredients);
  upsert('Recipes', [r]);
  if (data.log && data.log.length) upsert('Cook Log', data.log);
  return { ok: true };
}

function deleteRecipe(id) {
  deleteRows('Recipes', [id]);
  const logIds = readAll('Cook Log').filter(function (l) { return String(l.recipe_id) === String(id); }).map(function (l) { return l.id; });
  if (logIds.length) deleteRows('Cook Log', logIds);
  return { ok: true };
}

function saveIngredients(items) {
  upsert('Ingredients', items);
  return { ok: true };
}

function bumpCooked(entry, recipe) {
  // When a "made it" entry is logged, keep the recipe's counters in the Sheet up to date.
  if (recipe && recipe.id) {
    const row = { id: recipe.id };
    ['times_cooked', 'last_cooked', 'rating'].forEach(function (k) { if (recipe[k] !== undefined) row[k] = recipe[k]; });
    upsert('Recipes', [row]);
  }
  return { ok: true };
}

function shopOps(ops) {
  const puts = [], dels = [];
  ops.forEach(function (op) {
    if (op.op === 'put' && op.item) puts.push(op.item);
    if (op.op === 'del' && op.id) dels.push(op.id);
  });
  if (puts.length) upsert('Shopping', puts);
  if (dels.length) deleteRows('Shopping', dels);
  return { ok: true };
}

function uploadPhoto(data) {
  if (!data.data) throw new Error('No photo');
  const props = PropertiesService.getScriptProperties();
  let folder;
  const fid = props.getProperty('PHOTO_FOLDER_ID');
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder('Recipe Box Photos');
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }
  const bytes = Utilities.base64Decode(data.data);
  const blob = Utilities.newBlob(bytes, data.mediaType || 'image/jpeg', (data.name || 'recipe') + '.jpg');
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600', fileId: file.getId() };
}

// ════════════════════════════════════════════════════════════
// IMPORTS — websites (recipe schema), social video captions, AI
// ════════════════════════════════════════════════════════════

const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';

function fetchText(url, extraHeaders) {
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true,
    headers: Object.assign({ 'User-Agent': UA, 'Accept-Language': 'en-US,en;q=0.9' }, extraHeaders || {})
  });
  return { code: res.getResponseCode(), text: res.getContentText() };
}

function socialKind(url) {
  if (/(?:youtube\.com|youtu\.be)\//i.test(url)) return 'youtube';
  if (/tiktok\.com\//i.test(url)) return 'tiktok';
  if (/instagram\.com\//i.test(url)) return 'instagram';
  if (/(?:facebook\.com|fb\.watch)\//i.test(url)) return 'facebook';
  if (/pinterest\.[a-z.]+\/|pin\.it\//i.test(url)) return 'pinterest';
  return '';
}

function importUrl(url) {
  if (!/^https?:\/\//i.test(url)) return { ok: false, error: 'That doesn\'t look like a link.' };
  const kind = socialKind(url);
  if (kind && kind !== 'pinterest') return importSocial(url, kind);

  const page = fetchText(url);
  if (page.code >= 400) return { ok: false, error: 'The site refused the request (error ' + page.code + '). Try copying the recipe text and using Paste text instead.' };
  const r = recipeFromHtml(page.text, url);
  if (r && r.ingredients.length) {
    r.source_type = kind || 'website';
    return { ok: true, recipe: r, via: 'schema' };
  }
  if (aiKey()) {
    const text = htmlToText(page.text).slice(0, 40000);
    const meta = metaTags(page.text);
    const res = aiParse({ text: text, hint: 'This is the text of a web page that may contain a recipe.' });
    if (res.ok) {
      res.recipe.source_url = url;
      res.recipe.source_type = kind || 'website';
      if (!res.recipe.photo && meta.image) res.recipe.photo = meta.image;
    }
    return res;
  }
  return { ok: false, error: 'No recipe found on that page automatically. Add a Claude API key for smarter imports, or use Paste text.' };
}

function importSocial(url, kind) {
  let title = '', caption = '', photo = '', author = '';

  if (kind === 'youtube') {
    const id = (url.match(/(?:v=|youtu\.be\/|shorts\/|embed\/|live\/)([\w-]{11})/) || [])[1];
    try {
      const o = JSON.parse(fetchText('https://www.youtube.com/oembed?format=json&url=' + encodeURIComponent(url)).text);
      title = o.title || ''; author = o.author_name || ''; photo = o.thumbnail_url || '';
    } catch (e) {}
    if (id) {
      photo = 'https://i.ytimg.com/vi/' + id + '/hqdefault.jpg';
      const page = fetchText('https://www.youtube.com/watch?v=' + id).text;
      const m = page.match(/"shortDescription":"((?:[^"\\]|\\.)*)"/);
      if (m) { try { caption = JSON.parse('"' + m[1] + '"'); } catch (e) {} }
      if (!title) title = metaTags(page).title;
    }
  } else if (kind === 'tiktok') {
    try {
      const o = JSON.parse(fetchText('https://www.tiktok.com/oembed?url=' + encodeURIComponent(url)).text);
      caption = o.title || ''; author = o.author_name || ''; photo = o.thumbnail_url || '';
    } catch (e) {}
    if (!caption) {
      const page = fetchText(url).text;
      const m = metaTags(page); caption = m.description; photo = photo || m.image; title = m.title;
    }
  } else {
    // Instagram / Facebook: public pages expose the caption in their meta tags (when not behind a login).
    const page = fetchText(url).text;
    const m = metaTags(page);
    caption = m.description; photo = m.image; title = m.title;
    const q = caption.match(/:\s*["“]([\s\S]+)["”]\s*$/); // IG: '123 likes, … - user on May 1: "caption"'
    if (q) caption = q[1];
  }

  caption = decodeEntities(caption || '').trim();
  const base = { source_url: url, source_type: kind, photo: photo, author: author };

  // Many creators link the full written recipe in the caption — try that first.
  const links = (caption.match(/https?:\/\/[^\s"'<>)]+/g) || []).filter(function (l) { return !socialKind(l); }).slice(0, 3);
  for (let i = 0; i < links.length; i++) {
    try {
      const page = fetchText(links[i]);
      const r = recipeFromHtml(page.text, links[i]);
      if (r && r.ingredients.length) {
        r.source_url = url; r.source_type = kind;
        r.notes = ((r.notes || '') + '\nWritten recipe: ' + links[i]).trim();
        if (!r.photo) r.photo = photo;
        return { ok: true, recipe: r, via: 'linked recipe' };
      }
    } catch (e) {}
  }

  if (!caption && !title) {
    return { ok: false, error: 'Couldn\'t read that post (it may be private or need a login). Copy the caption and use Paste text, or screenshot it and use Photo.' };
  }

  if (aiKey()) {
    const res = aiParse({
      text: 'Title: ' + title + '\nBy: ' + author + '\n\nCaption / description:\n' + caption,
      hint: 'This is the caption of a ' + kind + ' cooking video. Extract the recipe. If amounts are missing, list the ingredients without amounts. If there is no recipe at all, set found to false.'
    });
    if (res.ok) {
      Object.keys(base).forEach(function (k) { if (base[k] && !res.recipe[k]) res.recipe[k] = base[k]; });
      res.recipe.source_url = url; res.recipe.source_type = kind;
      return res;
    }
  }
  // No AI: hand back what we have so the user can tidy it up.
  return {
    ok: true, via: 'caption', partial: true,
    recipe: Object.assign(base, { title: title || 'Recipe from ' + kind, description: '', ingredients: [], steps: [], notes: caption, tags: [] }),
    message: 'Pulled the caption into Notes — move the ingredients and steps across, or add a Claude API key to do this automatically.'
  };
}

function metaTags(html) {
  function get(prop) {
    const re1 = new RegExp('<meta[^>]+(?:property|name)=["\']' + prop + '["\'][^>]*content=["\']([^"\']*)["\']', 'i');
    const re2 = new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]*(?:property|name)=["\']' + prop + '["\']', 'i');
    const m = html.match(re1) || html.match(re2);
    return m ? decodeEntities(m[1]) : '';
  }
  return {
    title: get('og:title') || get('twitter:title') || ((html.match(/<title[^>]*>([^<]*)/i) || [])[1] || ''),
    description: get('og:description') || get('description') || get('twitter:description'),
    image: get('og:image') || get('twitter:image')
  };
}

/** Finds schema.org Recipe data (JSON-LD, used by almost every recipe site). */
function recipeFromHtml(html, url) {
  const blocks = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi) || [];
  let found = null;
  blocks.forEach(function (b) {
    if (found) return;
    const json = b.replace(/^<script[^>]*>/i, '').replace(/<\/script>$/i, '').trim();
    let data;
    try { data = JSON.parse(json); } catch (e) {
      try { data = JSON.parse(json.replace(/[\u0000-\u001F]+/g, ' ')); } catch (e2) { return; }
    }
    found = findRecipeNode(data);
  });
  if (!found) return null;

  const r = found;
  const meta = metaTags(html);
  const ingredients = [].concat(r.recipeIngredient || r.ingredients || []).map(function (s) { return clean(s); }).filter(String);
  const steps = [];
  flattenInstructions(r.recipeInstructions, steps);
  return {
    title: clean(r.name) || meta.title,
    description: clean(r.description),
    photo: imageUrl(r.image) || meta.image,
    source_url: url,
    servings: clean(Array.isArray(r.recipeYield) ? r.recipeYield[0] : r.recipeYield),
    prep_time: isoDur(r.prepTime),
    cook_time: isoDur(r.cookTime),
    total_time: isoDur(r.totalTime),
    tags: uniq([].concat(splitList(r.recipeCategory), splitList(r.recipeCuisine), splitList(r.keywords).slice(0, 6))),
    ingredients: ingredients,
    steps: steps,
    notes: ''
  };
}

function findRecipeNode(d) {
  if (!d) return null;
  if (Array.isArray(d)) { for (let i = 0; i < d.length; i++) { const f = findRecipeNode(d[i]); if (f) return f; } return null; }
  if (typeof d !== 'object') return null;
  const t = [].concat(d['@type'] || []);
  if (t.indexOf('Recipe') > -1) return d;
  if (d['@graph']) return findRecipeNode(d['@graph']);
  if (d.mainEntity) return findRecipeNode(d.mainEntity);
  return null;
}

function flattenInstructions(x, out, section) {
  if (!x) return;
  if (typeof x === 'string') {
    // Some sites put all steps in one string.
    htmlToText(x).split(/\n+/).map(clean).filter(String).forEach(function (s) { out.push(s); });
    return;
  }
  if (Array.isArray(x)) { x.forEach(function (i) { flattenInstructions(i, out); }); return; }
  const t = [].concat(x['@type'] || []);
  if (t.indexOf('HowToSection') > -1) {
    if (x.name) out.push({ section: clean(x.name) });
    flattenInstructions(x.itemListElement, out);
    return;
  }
  const text = clean(x.text || x.name || '');
  if (text) out.push(text);
}

function imageUrl(img) {
  if (!img) return '';
  if (typeof img === 'string') return img;
  if (Array.isArray(img)) return imageUrl(img[img.length > 1 ? img.length - 1 : 0]) || imageUrl(img[0]);
  return img.url || img.contentUrl || '';
}

function isoDur(s) {
  if (!s || typeof s !== 'string') return '';
  const m = s.match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?/i);
  if (!m) return s;
  const mins = (+m[1] || 0) * 1440 + (+m[2] || 0) * 60 + (+m[3] || 0);
  if (!mins) return '';
  const h = Math.floor(mins / 60), mm = mins % 60;
  return (h ? h + ' hr' + (mm ? ' ' : '') : '') + (mm ? mm + ' min' : '');
}

function splitList(v) {
  if (!v) return [];
  return [].concat(v).join(',').split(',').map(function (s) { return clean(s); }).filter(String);
}
function uniq(a) {
  const seen = {}; return a.filter(function (x) { const k = x.toLowerCase(); if (seen[k]) return false; seen[k] = 1; return true; });
}
function clean(s) {
  if (s === undefined || s === null) return '';
  return decodeEntities(String(s).replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
}
function decodeEntities(s) {
  return String(s || '')
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(+n); })
    .replace(/&#x([0-9a-f]+);/gi, function (_, n) { return String.fromCharCode(parseInt(n, 16)); })
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&frac12;/g, '½').replace(/&frac14;/g, '¼').replace(/&frac34;/g, '¾')
    .replace(/&ndash;/g, '–').replace(/&mdash;/g, '—').replace(/&deg;/g, '°').replace(/&rsquo;|&lsquo;/g, "'").replace(/&rdquo;|&ldquo;/g, '"');
}
function htmlToText(html) {
  return decodeEntities(String(html)
    .replace(/<(script|style|noscript|svg|nav|footer|header)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|li|h\d|div|tr|section)>/gi, '\n')
    .replace(/<[^>]+>/g, ' '))
    .replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim();
}

// ════════════════════════════════════════════════════════════
// CLAUDE (optional) — reads photos, screenshots, captions and messy text
// ════════════════════════════════════════════════════════════

function aiKey() { return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || ''; }

const AI_SYSTEM =
  'You turn recipes from photos, screenshots, video captions and messy text into clean structured data. ' +
  'Reply with ONLY a JSON object, no prose, in exactly this shape:\n' +
  '{"found":true,"title":"","description":"","servings":"","prep_time":"","cook_time":"","total_time":"",' +
  '"tags":[""],"ingredients":["2 cups all-purpose flour", {"section":"For the sauce"}, "1 tbsp soy sauce"],' +
  '"steps":["Preheat the oven to 350°F.", {"section":"Sauce"}, "Whisk…"],"notes":""}\n' +
  'Rules: one ingredient per string, written "amount unit ingredient, prep" (e.g. "3 cloves garlic, minced"). ' +
  'Use {"section": "..."} objects only when the recipe has named parts. Keep the author\'s amounts and wording; do not invent amounts. ' +
  'Steps are full sentences, one step per string, without numbers. Times like "15 min" or "1 hr 10 min". ' +
  'tags: up to 5 short lowercase tags (meal type, cuisine, main ingredient). notes: tips, storage or substitutions mentioned by the author. ' +
  'For handwriting, read carefully. If there is no recipe, return {"found":false}.';

function aiParse(data) {
  const key = aiKey();
  if (!key) return { ok: false, error: 'Photo and smart imports need a Claude API key. In the Sheet: Recipe Box → Set Claude API key.', code: 'noai' };
  const content = [];
  (data.images || []).slice(0, 5).forEach(function (img) {
    content.push({ type: 'image', source: { type: 'base64', media_type: img.mediaType || 'image/jpeg', data: img.data } });
  });
  let prompt = (data.hint ? data.hint + '\n\n' : '');
  if (data.text) prompt += data.text;
  if (!data.text && content.length) prompt += 'Extract the recipe from ' + (content.length > 1 ? 'these images (they are pages of the same recipe).' : 'this image.');
  content.push({ type: 'text', text: prompt });

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', muteHttpExceptions: true, contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: PropertiesService.getScriptProperties().getProperty('CLAUDE_MODEL') || DEFAULT_MODEL,
      max_tokens: 4096,
      system: AI_SYSTEM,
      messages: [{ role: 'user', content: content }]
    })
  });
  const code = res.getResponseCode();
  let body;
  try { body = JSON.parse(res.getContentText()); } catch (e) { body = {}; }
  if (code !== 200) {
    const msg = body.error && body.error.message || ('error ' + code);
    return { ok: false, error: 'Claude couldn\'t read it: ' + msg };
  }
  const text = (body.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
  const s = text.indexOf('{'), e = text.lastIndexOf('}');
  let r;
  try { r = JSON.parse(text.slice(s, e + 1)); } catch (err) { return { ok: false, error: 'Couldn\'t understand the reply. Try again.' }; }
  if (r.found === false) return { ok: false, error: 'No recipe found in that. Check the photo or text and try again.' };
  r.ingredients = r.ingredients || []; r.steps = r.steps || []; r.tags = r.tags || [];
  return { ok: true, recipe: r, via: 'ai' };
}

// Run this from the editor to check imports work (View → Logs).
function testImport() {
  Logger.log(JSON.stringify(importUrl('https://www.allrecipes.com/recipe/10813/best-chocolate-chip-cookies/')).slice(0, 2000));
}
