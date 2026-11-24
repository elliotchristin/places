/* Lieux — personal photo map over OneDrive (Microsoft Graph), no backend.
   Reads only the OneDrive "Images" folder. Write access is requested only when
   you first create an album or like a photo. */
'use strict';

const CLIENT_ID = '2255226b-3bd7-48fd-bed7-e45d25432566';
const READ = ['Files.Read'];
const WRITE = ['Files.ReadWrite'];
const GRAPH = 'https://graph.microsoft.com/v1.0';
const PIN_SIZE = 'c360x360_crop';
const VIEW_SIZE = 'c2048x2048';
const FAV_ALBUM = 'Favoris';
const DATA_VERSION = 4; // bump to force a full re-read when the stored item shape changes
const DEBUG = new URLSearchParams(location.search).has('debug');

const $ = (id) => document.getElementById(id);
const log = (...a) => DEBUG && console.log('[lieux]', ...a);
const fmtDate = new Intl.DateTimeFormat('fr-CA', { day: 'numeric', month: 'long', year: 'numeric' });
const fmtMonth = new Intl.DateTimeFormat('fr-CA', { month: 'long', year: 'numeric' });
const fmtNum = new Intl.NumberFormat('fr-CA');
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
const plural = (n, one, many) => `${fmtNum.format(n)} ${n > 1 ? many : one}`;
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const VID_ICON = '<svg class="vid" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z" fill="currentColor"/></svg>';
const CHECK_ICON = '<svg class="check" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M7.5 12.5l3 3 6-6.5" fill="none" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';

/* ================= Auth ================= */
const redirectUri = new URL('./', location.href).href;
const msalApp = new msal.PublicClientApplication({
  auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', redirectUri },
  cache: { cacheLocation: 'localStorage' },
});
let account = null;

async function initAuth() {
  await msalApp.initialize();
  const res = await msalApp.handleRedirectPromise();
  account = res?.account || msalApp.getAllAccounts()[0] || null;
  if (account) msalApp.setActiveAccount(account);
  return account;
}
const signIn = () => msalApp.loginRedirect({ scopes: READ, prompt: 'select_account' });

async function getToken(scopes = READ) {
  try {
    return (await msalApp.acquireTokenSilent({ scopes, account })).accessToken;
  } catch (e) {
    log('silent token failed', e);
    await msalApp.acquireTokenRedirect({ scopes, account });
    return new Promise(() => {}); // navigating away
  }
}

/* Write access: ask Microsoft once, with an explanation first. */
async function ensureWrite() {
  try {
    await msalApp.acquireTokenSilent({ scopes: WRITE, account });
    return true;
  } catch {
    openModal('Autoriser les albums', `
      <p class="modal-text">Pour créer des albums et ajouter des favoris, Lieux a besoin de l'autorisation de <strong>modifier</strong> ton OneDrive. Microsoft va te demander de confirmer.</p>
      <p class="modal-text fine">Lieux n'utilise cette autorisation que pour créer des albums et y ajouter des photos. Une photo dans un album n'est pas copiée : c'est la même photo, rangée à un autre endroit.</p>
      <button class="primary" id="grantWrite">Continuer</button>`);
    $('grantWrite').onclick = () => msalApp.acquireTokenRedirect({ scopes: WRITE, account });
    return false;
  }
}

/* ================= Graph ================= */
async function graph(url, opts = {}, scopes = READ, attempt = 0) {
  const token = await getToken(scopes);
  let res;
  try {
    res = await fetch(url.startsWith('http') ? url : GRAPH + url, {
      ...opts,
      headers: { Authorization: 'Bearer ' + token, ...(opts.body ? { 'Content-Type': 'application/json' } : {}) },
    });
  } catch (netErr) { // flaky mobile network: retry a few times
    if (attempt >= 4) throw netErr;
    await new Promise((r) => setTimeout(r, 1000 * 2 ** attempt));
    return graph(url, opts, scopes, attempt + 1);
  }
  if ((res.status === 429 || res.status >= 500) && attempt < 5) {
    await new Promise((r) => setTimeout(r, (Number(res.headers.get('Retry-After')) || 2 ** attempt) * 1000));
    return graph(url, opts, scopes, attempt + 1);
  }
  if (res.status === 204) return null;
  if (!res.ok) {
    const err = new Error(`Graph ${res.status}`);
    err.status = res.status;
    err.body = await res.text().catch(() => '');
    throw err;
  }
  return res.json();
}
async function graphAll(url, scopes = READ, maxPages = 50) {
  const out = [];
  for (let i = 0; url && i < maxPages; i++) {
    const page = await graph(url, {}, scopes);
    out.push(...page.value);
    url = page['@odata.nextLink'];
  }
  return out;
}
/* Run up to 20 requests per call through JSON batching. */
async function graphBatch(requests, scopes = READ, parallel = 1) {
  const chunks = [];
  for (let i = 0; i < requests.length; i += 20) chunks.push(requests.slice(i, i + 20));
  const results = new Array(chunks.length);
  let next = 0;
  const worker = async () => {
    while (next < chunks.length) {
      const k = next++;
      const chunk = chunks[k].map((r, j) => ({ id: String(j), ...r, ...(r.body ? { headers: { 'Content-Type': 'application/json' } } : {}) }));
      const res = await graph('/$batch', { method: 'POST', body: JSON.stringify({ requests: chunk }) }, scopes);
      const m = new Map(res.responses.map((r) => [r.id, r]));
      results[k] = chunk.map((_, j) => m.get(String(j)));
    }
  };
  await Promise.all(Array.from({ length: Math.min(parallel, chunks.length) }, worker));
  return results.flat();
}

/* ================= Local store (IndexedDB) ================= */
const db = {
  _db: null,
  open() {
    if (this._db) return Promise.resolve(this._db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('lieux', 1);
      req.onupgradeneeded = () => { req.result.createObjectStore('items', { keyPath: 'id' }); req.result.createObjectStore('meta'); };
      req.onsuccess = () => resolve((this._db = req.result));
      req.onerror = () => reject(req.error);
    });
  },
  async tx(store, mode, fn) {
    const d = await this.open();
    return new Promise((resolve, reject) => {
      const t = d.transaction(store, mode);
      let out;
      Promise.resolve(fn(t.objectStore(store))).then((v) => (out = v));
      t.oncomplete = () => resolve(out);
      t.onerror = () => reject(t.error);
    });
  },
  all() { return this.tx('items', 'readonly', (s) => new Promise((r) => { const q = s.getAll(); q.onsuccess = () => r(q.result); })); },
  getMeta(k) { return this.tx('meta', 'readonly', (s) => new Promise((r) => { const q = s.get(k); q.onsuccess = () => r(q.result); })); },
  setMeta(k, v) { return this.tx('meta', 'readwrite', (s) => { s.put(v, k); }); },
  apply(puts, dels) { return this.tx('items', 'readwrite', (s) => { puts.forEach((p) => s.put(p)); dels.forEach((id) => s.delete(id)); }); },
  async reset() {
    await this.tx('items', 'readwrite', (s) => s.clear());
    await this.tx('meta', 'readwrite', (s) => { s.clear(); s.put(DATA_VERSION, 'version'); });
  },
};

/* ================= Sync (delta query on the Images folder only) ================= */
const SELECT = 'id,name,file,location,photo,deleted,webUrl,fileSystemInfo';

function toItem(it) {
  const loc = it.location;
  const mime = it.file?.mimeType || '';
  if (it.deleted || !it.file || !/^(image|video)\//.test(mime)) return null;
  if (!loc || loc.latitude == null || loc.longitude == null) return null;
  if (Math.abs(loc.latitude) < 1e-6 && Math.abs(loc.longitude) < 1e-6) return null; // "null island"
  const taken = it.photo?.takenDateTime || it.fileSystemInfo?.createdDateTime;
  return {
    id: it.id,
    n: it.name, // album entries get their own ids, so favourites are matched by file name
    la: +loc.latitude.toFixed(6),
    lo: +loc.longitude.toFixed(6),
    t: taken ? Date.parse(taken) : 0,
    v: mime.startsWith('video/') ? 1 : 0, // not the "video" facet: motion photos carry it too
    u: it.webUrl,
  };
}

async function startUrl() {
  let rootId = await db.getMeta('rootId');
  if (!rootId) {
    rootId = (await graph('/me/drive/special/photos?$select=id')).id;
    await db.setMeta('rootId', rootId);
  }
  return `${GRAPH}/me/drive/items/${rootId}/delta?$select=${SELECT}&$top=1000`;
}

async function sync(onProgress) {
  let url = (await db.getMeta('resumeLink')) || (await db.getMeta('deltaLink')) || (await startUrl());
  let seen = 0, changed = 0;
  while (url) {
    let page;
    try {
      page = await graph(url);
    } catch (e) {
      if (e.status === 410) { await db.reset(); url = await startUrl(); continue; } // delta token expired
      throw e;
    }
    const puts = [], dels = [];
    for (const it of page.value) { const item = toItem(it); if (item) puts.push(item); else dels.push(it.id); }
    await db.apply(puts, dels);
    seen += page.value.length;
    changed += puts.length + dels.filter((id) => byId.has(id)).length;
    onProgress?.(seen, changed);
    if (page['@odata.nextLink']) {
      url = page['@odata.nextLink'];
      await db.setMeta('resumeLink', url);
    } else {
      await db.setMeta('deltaLink', page['@odata.deltaLink']);
      await db.setMeta('resumeLink', null);
      url = null;
    }
  }
  return { seen, changed };
}

/* ================= Thumbnails (batched, links expire after a few hours) ================= */
const thumbCache = new Map();
const thumbQueue = [];
let thumbTimer = null, inflight = 0;

function thumb(id, size = PIN_SIZE) {
  const key = id + '|' + size;
  if (thumbCache.has(key)) return thumbCache.get(key);
  const p = new Promise((resolve) => thumbQueue.push({ id, size, resolve }));
  thumbCache.set(key, p);
  if (!thumbTimer) thumbTimer = setTimeout(flushThumbs, 30);
  return p;
}
function flushThumbs() {
  thumbTimer = null;
  while (thumbQueue.length && inflight < 4) {
    const batch = thumbQueue.splice(0, 20);
    inflight++;
    runThumbBatch(batch).finally(() => { inflight--; if (thumbQueue.length) flushThumbs(); });
  }
}
async function runThumbBatch(batch) {
  try {
    const res = await graphBatch(batch.map((b) => ({ method: 'GET', url: `/me/drive/items/${b.id}/thumbnails?select=${b.size}` })));
    batch.forEach((b, i) => {
      const r = res[i];
      const url = r && r.status === 200 ? r.body?.value?.[0]?.[b.size]?.url : null;
      if (!url) thumbCache.delete(b.id + '|' + b.size);
      b.resolve(url || null);
    });
  } catch (e) {
    log('thumb batch failed', e);
    batch.forEach((b) => { thumbCache.delete(b.id + '|' + b.size); b.resolve(null); });
  }
}
const setBg = (el, id, size = PIN_SIZE) =>
  thumb(id, size).then((url) => { if (url) el.style.backgroundImage = `url("${url}")`; return url; });

/* ================= Place names ================= */
const placeCache = new Map();
function placeName(lat, lon) {
  const key = lat.toFixed(2) + ',' + lon.toFixed(2);
  if (placeCache.has(key)) return placeCache.get(key);
  const p = fetch(`https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=${lat}&longitude=${lon}&localityLanguage=fr`)
    .then((r) => (r.ok ? r.json() : null))
    .then((j) => {
      if (!j) return '';
      const parts = [j.city || j.locality || j.principalSubdivision, j.countryName].filter(Boolean);
      return [...new Set(parts)].join(', ');
    })
    .catch(() => '');
  placeCache.set(key, p);
  return p;
}

/* ================= State & filters ================= */
let map, index = null;
let items = [], byId = new Map();
let yearFilter = 'all';
let setFilter = null; // { label, ids: Set, kind: 'album'|'search', album? }

function filtered() {
  let list = items;
  if (setFilter) list = setFilter.list || list.filter((i) => setFilter.ids.has(i.id));
  if (yearFilter !== 'all') list = list.filter((i) => new Date(i.t).getFullYear() === yearFilter);
  return list;
}

function setItems(list) {
  items = list;
  byId = new Map(list.map((i) => [i.id, i]));
  rebuild();
}

function rebuild() {
  index = new Supercluster({
    radius: 62,
    maxZoom: 17,
    map: (p) => ({ rep: p.id, rt: p.t }),
    reduce: (a, b) => { if (b.rt > a.rt) { a.rt = b.rt; a.rep = b.rep; } },
  });
  index.load(filtered().map((i) => ({ type: 'Feature', geometry: { type: 'Point', coordinates: [i.lo, i.la] }, properties: { id: i.id, t: i.t } })));
  renderChrome();
  clearMarkers();
  scheduleMarkers();
}

/* ================= Map ================= */
const darkMQ = matchMedia('(prefers-color-scheme: dark)');
const store = {
  get(k, d) { try { return localStorage.getItem('lieux.' + k) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('lieux.' + k, v); } catch {} },
};
let styleMode = store.get('style', 'auto');
const resolvedStyle = () => (styleMode === 'auto' ? (darkMQ.matches ? 'dark' : 'light') : styleMode);

async function applyStyle() {
  const s = await MapStyles.build(resolvedStyle());
  map.setStyle(s, { diff: false });
  document.querySelectorAll('#styleSeg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.style === styleMode)));
}

async function initMap() {
  const style = await MapStyles.build(resolvedStyle());
  map = new maplibregl.Map({
    container: 'map',
    style,
    center: [-30, 30],
    zoom: 1.6,
    attributionControl: { compact: true },
    dragRotate: false,
    pitchWithRotate: false,
  });
  map.touchZoomRotate.disableRotation();
  map.on('move', scheduleMarkers);
  map.on('moveend', scheduleMarkers);
  map.once('load', () => document.querySelector('.maplibregl-ctrl-attrib')?.classList.remove('maplibregl-compact-show'));
  darkMQ.addEventListener('change', () => styleMode === 'auto' && applyStyle());
  document.querySelectorAll('#styleSeg button').forEach((b) => {
    b.setAttribute('aria-pressed', String(b.dataset.style === styleMode));
    b.onclick = () => { styleMode = b.dataset.style; store.set('style', styleMode); applyStyle(); closePopover(); };
  });
}

/* ---------- Markers (clustered with supercluster: one marker per cluster, no duplicates) ---------- */
const markers = new Map();
let markerFrame = null;
function scheduleMarkers() {
  if (markerFrame || !map) return;
  const run = () => { markerFrame = null; updateMarkers(); };
  markerFrame = document.hidden ? setTimeout(run, 50) : requestAnimationFrame(run);
}
function clearMarkers() { for (const m of markers.values()) m.remove(); markers.clear(); }
const wrapLng = (x) => ((((x + 180) % 360) + 360) % 360) - 180;

function updateMarkers() {
  if (!index) return;
  const b = map.getBounds();
  const span = b.getEast() - b.getWest();
  const bbox = span >= 360
    ? [-180, Math.max(b.getSouth(), -85), 180, Math.min(b.getNorth(), 85)]
    : [wrapLng(b.getWest()), Math.max(b.getSouth(), -85), wrapLng(b.getEast()), Math.min(b.getNorth(), 85)];
  const z = Math.floor(map.getZoom());
  const seen = new Set();
  for (const f of index.getClusters(bbox, z)) {
    const p = f.properties;
    const key = p.cluster ? `c${p.cluster_id}` : `p${p.id}`;
    seen.add(key);
    if (markers.has(key)) continue;
    const rep = byId.get(p.cluster ? p.rep : p.id);
    if (!rep) continue;
    const el = buildPin(p, rep);
    // MapLibre positions markers with a CSS transform on the element it gets, so the pin
    // (which has its own hover/transition transforms) must live inside a plain wrapper.
    const wrap = document.createElement('div');
    wrap.className = 'pin-wrap';
    wrap.appendChild(el);
    // Pin sits on a real photo (the most recent of the group), never on an average position.
    const m = new maplibregl.Marker({ element: wrap, anchor: 'bottom', offset: [0, -8] }).setLngLat([rep.lo, rep.la]).addTo(map);
    el.addEventListener('click', (ev) => { ev.stopPropagation(); onPinClick(p, rep); });
    markers.set(key, m);
  }
  for (const [k, m] of markers) if (!seen.has(k)) { m.remove(); markers.delete(k); }
}

function buildPin(p, rep) {
  const el = document.createElement('button');
  el.className = 'pin';
  const count = p.cluster ? p.point_count : 1;
  if (count === 1) el.classList.add('small');
  if (count > 1) el.innerHTML = `<span class="count">${count >= 1000 ? Math.round(count / 100) / 10 + 'k' : count}</span>`;
  else if (rep.v) el.innerHTML = VID_ICON;
  el.setAttribute('aria-label', count > 1 ? `${count} photos` : 'Photo');
  setBg(el, rep.id).then(async (u) => {
    if (u || !p.cluster) return; // no thumbnail: try other photos of the group
    for (const l of index.getLeaves(p.cluster_id, 10)) if (l.properties.id !== rep.id && (await setBg(el, l.properties.id))) return;
  });
  return el;
}

function onPinClick(p, rep) {
  if (!p.cluster) return openViewer([rep], 0);
  if (map.getZoom() < 9) {
    const z = index.getClusterExpansionZoom(p.cluster_id);
    map.easeTo({ center: [rep.lo, rep.la], zoom: Math.min(z + 0.5, 17), duration: 600 });
    return;
  }
  const list = index.getLeaves(p.cluster_id, Infinity).map((l) => byId.get(l.properties.id)).filter(Boolean);
  openSheet(list, { fit: true });
}

function fitList(list, { sheet = false, animate = true } = {}) {
  if (!list.length) return;
  const b = new maplibregl.LngLatBounds();
  list.forEach((i) => b.extend([i.lo, i.la]));
  const wide = innerWidth >= 760;
  const padding = sheet
    ? (wide ? { top: 120, bottom: 60, left: 460, right: 80 } : { top: 110, bottom: innerHeight * 0.62 + 30, left: 50, right: 70 })
    : { top: 120, bottom: 110, left: 50, right: 70 };
  map.fitBounds(b, { padding, maxZoom: sheet ? 16 : 13, duration: animate ? 700 : 0 });
}
const fitAll = (animate = true) => fitList(filtered(), { animate });

/* ================= Chrome: summary, years, filter chip ================= */
function renderChrome() {
  const list = filtered();
  const videos = list.filter((i) => i.v).length, photos = list.length - videos;
  $('summary').textContent = [plural(photos, 'photo', 'photos'), videos ? plural(videos, 'vidéo', 'vidéos') : ''].filter(Boolean).join(' · ');

  const base = setFilter ? (setFilter.list || items.filter((i) => setFilter.ids.has(i.id))) : items;
  const years = [...new Set(base.map((i) => new Date(i.t).getFullYear()))].filter((y) => y > 1990).sort((a, b) => b - a);
  if (yearFilter !== 'all' && !years.includes(yearFilter)) yearFilter = 'all';
  const nav = $('years');
  nav.hidden = years.length < 2;
  nav.innerHTML = '';
  for (const y of ['all', ...years]) {
    const b = document.createElement('button');
    b.textContent = y === 'all' ? 'Tout' : y;
    b.setAttribute('aria-pressed', String(y === yearFilter));
    b.onclick = () => { yearFilter = y; rebuild(); fitAll(); closeSheet(); };
    nav.appendChild(b);
  }

  $('filterChip').hidden = !setFilter;
  if (setFilter) $('filterLabel').textContent = setFilter.label;
}

function applySetFilter(f, fit = true) {
  setFilter = f;
  f?.list?.forEach((i) => byId.set(i.id, i)); // album photos may live outside the Images folder
  yearFilter = 'all';
  rebuild();
  if (fit) fitAll();
}
$('filterClear').onclick = () => { applySetFilter(null); closeSheet(); };

/* ================= Sheet (photo grid) ================= */
let sheetList = [], selecting = false;
const selected = new Set();
const io = new IntersectionObserver((entries) => {
  for (const e of entries) if (e.isIntersecting) { io.unobserve(e.target); setBg(e.target, e.target.dataset.id); }
}, { root: $('grid'), rootMargin: '300px' });

function openSheet(list, { title = null, fit = false } = {}) {
  sheetList = list.slice().sort((a, b) => b.t - a.t);
  setSelecting(false);
  const n = sheetList.length;
  const oldest = sheetList[n - 1], newest = sheetList[0];
  const range = oldest && newest
    ? (fmtMonth.format(oldest.t) === fmtMonth.format(newest.t) ? fmtMonth.format(newest.t) : `${fmtMonth.format(oldest.t)} – ${fmtMonth.format(newest.t)}`)
    : '';
  $('sheetSub').textContent = `${plural(n, 'élément', 'éléments')}${range ? ' · ' + range : ''}`;
  if (title) $('sheetTitle').textContent = title;
  else {
    $('sheetTitle').textContent = '…';
    const lat = sheetList.reduce((s, i) => s + i.la, 0) / n, lon = sheetList.reduce((s, i) => s + i.lo, 0) / n;
    const mine = sheetList;
    placeName(lat, lon).then((name) => { if (sheetList === mine) $('sheetTitle').textContent = name || 'Ici'; });
  }

  const grid = $('grid');
  io.disconnect();
  grid.innerHTML = '';
  grid.scrollTop = 0;
  let month = '';
  const frag = document.createDocumentFragment();
  sheetList.forEach((it, idx) => {
    const m = fmtMonth.format(it.t);
    if (m !== month && n > 12) {
      month = m;
      const h = document.createElement('div');
      h.className = 'month';
      h.textContent = cap(m);
      frag.appendChild(h);
    }
    const c = document.createElement('button');
    c.className = 'cell';
    c.dataset.id = it.id;
    c.setAttribute('aria-label', fmtDate.format(it.t));
    c.innerHTML = (it.v ? VID_ICON : '') + CHECK_ICON;
    c.onclick = () => {
      if (!selecting) return openViewer(sheetList, idx);
      if (selected.has(it.id)) selected.delete(it.id); else selected.add(it.id);
      c.classList.toggle('selected', selected.has(it.id));
      renderSelbar();
    };
    frag.appendChild(c);
    io.observe(c);
  });
  grid.appendChild(frag);
  $('sheet').classList.add('open');
  $('sheet').setAttribute('aria-hidden', 'false');
  if (fit) fitList(sheetList, { sheet: true });
}
function closeSheet() {
  $('sheet').classList.remove('open');
  $('sheet').setAttribute('aria-hidden', 'true');
  setSelecting(false);
}
function setSelecting(on) {
  selecting = on;
  selected.clear();
  $('sheet').classList.toggle('selecting', on);
  $('selectBtn').textContent = on ? 'Annuler' : 'Sélectionner';
  document.querySelectorAll('.grid .cell.selected').forEach((c) => c.classList.remove('selected'));
  renderSelbar();
}
function renderSelbar() {
  $('selbar').hidden = !selecting;
  const n = selected.size;
  $('selCount').textContent = n ? plural(n, 'sélectionnée', 'sélectionnées') : 'Touche des photos';
  $('selAlbum').disabled = $('selLike').disabled = !n;
}
$('selectBtn').onclick = () => setSelecting(!selecting);
$('selAlbum').onclick = () => openAlbumPicker([...selected]);
$('selLike').onclick = async () => {
  const ids = [...selected];
  if (await setLiked(ids, true)) { toast(`${plural(ids.length, 'photo ajoutée', 'photos ajoutées')} aux favoris`); setSelecting(false); }
};

// Swipe the sheet down to close.
(() => {
  let y0 = null;
  const s = $('sheet');
  s.addEventListener('touchstart', (e) => { y0 = e.target.closest('.grid') && $('grid').scrollTop > 0 ? null : e.touches[0].clientY; }, { passive: true });
  s.addEventListener('touchend', (e) => { if (y0 != null && e.changedTouches[0].clientY - y0 > 80) closeSheet(); y0 = null; });
})();

/* ================= Viewer ================= */
let vList = [], vIdx = 0;
function openViewer(list, idx) {
  vList = list; vIdx = idx;
  $('viewer').classList.add('open');
  $('viewer').setAttribute('aria-hidden', 'false');
  showViewer();
  loadFavs(); // to show the heart state
}
function showViewer() {
  const it = vList[vIdx];
  if (!it) return;
  const img = $('viewerImg');
  img.classList.add('loading');
  img.dataset.full = '';
  thumbCache.get(it.id + '|' + PIN_SIZE)?.then((u) => { if (u && vList[vIdx] === it && !img.dataset.full) img.src = u; });
  thumb(it.id, VIEW_SIZE).then((u) => {
    if (vList[vIdx] !== it) return;
    if (u) { img.src = u; img.dataset.full = '1'; }
    img.onload = () => img.classList.remove('loading');
  });
  stopVideo();
  $('viewerPlay').hidden = !it.v;
  $('viewerDate').textContent = it.t ? fmtDate.format(it.t) : '';
  $('viewerPlace').textContent = '';
  placeName(it.la, it.lo).then((n) => { if (vList[vIdx] === it) $('viewerPlace').textContent = n; });
  $('viewerOpen').href = it.u;
  $('viewerCount').textContent = vList.length > 1 ? `${vIdx + 1} sur ${fmtNum.format(vList.length)}` : '';
  renderHeart();
  [vIdx - 1, vIdx + 1].forEach((j) => vList[j] && thumb(vList[j].id, VIEW_SIZE).then((u) => { if (u) new Image().src = u; }));
}
function stepViewer(d) {
  const j = vIdx + d;
  if (j < 0 || j >= vList.length) return;
  vIdx = j;
  showViewer();
}
function closeViewer() {
  stopVideo();
  $('viewer').classList.remove('open');
  $('viewer').setAttribute('aria-hidden', 'true');
  $('viewerImg').removeAttribute('src');
}

/* Videos play inside the app, streamed from OneDrive only when you tap play. */
function stopVideo() {
  const v = $('viewerVideo');
  v.pause();
  v.removeAttribute('src');
  v.load();
  v.hidden = true;
  $('viewerImg').hidden = false;
}
async function playVideo() {
  const it = vList[vIdx];
  if (!it?.v) return;
  const play = $('viewerPlay');
  play.classList.add('loading');
  try {
    const r = await graph(`/me/drive/items/${it.id}?select=id,@microsoft.graph.downloadUrl`);
    if (vList[vIdx] !== it) return;
    const v = $('viewerVideo');
    v.src = r['@microsoft.graph.downloadUrl'];
    v.hidden = false;
    $('viewerImg').hidden = true;
    play.hidden = true;
    await v.play().catch(() => {}); // iOS may need a second tap on the native control
  } catch (e) {
    console.error(e);
    toast('Impossible de lire cette vidéo ici.', it.u);
  } finally {
    play.classList.remove('loading');
  }
}
$('viewerPlay').onclick = (e) => { e.stopPropagation(); playVideo(); };
function renderHeart() {
  const it = vList[vIdx];
  const on = isFav(it);
  $('viewerLike').setAttribute('aria-pressed', String(on));
  $('viewerLike').setAttribute('aria-label', on ? 'Retirer des favoris' : 'Ajouter aux favoris');
}
$('viewerLike').onclick = async () => {
  const it = vList[vIdx];
  if (!it || !(await ensureWrite())) return;
  await loadFavs();
  const on = !isFav(it);
  if (await setLiked([it.id], on)) { renderHeart(); toast(on ? 'Ajoutée aux favoris' : 'Retirée des favoris'); }
};
$('viewerAlbum').onclick = () => vList[vIdx] && openAlbumPicker([vList[vIdx].id]);
(() => {
  let x0 = null, y0 = null;
  const st = $('viewerStage');
  st.addEventListener('touchstart', (e) => {
    if (e.target.closest('video')) { x0 = y0 = null; return; } // scrubbing a video is not a swipe
    x0 = e.touches[0].clientX; y0 = e.touches[0].clientY;
  }, { passive: true });
  st.addEventListener('touchend', (e) => {
    if (x0 == null) return;
    const dx = e.changedTouches[0].clientX - x0, dy = e.changedTouches[0].clientY - y0;
    if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) stepViewer(dx < 0 ? 1 : -1);
    else if (dy > 90) closeViewer();
    x0 = y0 = null;
  });
  st.addEventListener('click', (e) => {
    if (e.target.closest('video')) return; // let the native video controls work
    if (vList[vIdx]?.v && !$('viewerPlay').hidden) { playVideo(); return; }
    const r = st.getBoundingClientRect();
    if (e.clientX < r.left + r.width / 3) stepViewer(-1);
    else if (e.clientX > r.right - r.width / 3) stepViewer(1);
  });
})();

/* ================= Albums & favourites (OneDrive albums = Graph "bundles") ================= */
let albumsCache = null;
// The webUrl Graph returns for albums (photos.onedrive.com/album/…) lands on an error page; this one opens the album.
const albumUrl = (id) => `https://onedrive.live.com/photos?photosData=${encodeURIComponent('/photos/album/' + id)}`;
// Favourites album: file name -> id of the entry inside the album (OneDrive gives album entries their own ids).
let favByName = null, favAlbum = undefined;

async function listAlbums(force = false) {
  if (albumsCache && !force) return albumsCache;
  // No $select here: Graph drops the "bundle" facet when fields are selected.
  const all = await graphAll(`${GRAPH}/me/drive/bundles?$top=200`);
  albumsCache = all.filter((b) => b.bundle?.album || /\/album\//.test(b.webUrl || ''))
    .map((b) => ({ id: b.id, name: b.name, count: b.bundle?.childCount ?? 0, url: albumUrl(b.id), mod: Date.parse(b.lastModifiedDateTime) || 0 }))
    .sort((a, b) => b.mod - a.mod);
  return albumsCache;
}
async function albumIds(albumId) {
  const kids = await graphAll(`${GRAPH}/me/drive/items/${albumId}/children?$select=id&$top=1000`);
  return new Set(kids.map((k) => k.id));
}
async function addToAlbum(albumId, ids) {
  const res = await graphBatch(ids.map((id) => ({ method: 'POST', url: `/drive/bundles/${albumId}/children`, body: { id } })), WRITE);
  return res.filter((r) => r && (r.status < 300 || r.status === 409)).length;
}
async function createAlbum(name, ids) {
  const album = await graph('/me/drive/bundles', {
    method: 'POST',
    body: JSON.stringify({ name, '@microsoft.graph.conflictBehavior': 'rename', bundle: { album: {} }, children: ids.slice(0, 1).map((id) => ({ id })) }),
  }, WRITE);
  if (ids.length > 1) await addToAlbum(album.id, ids.slice(1));
  albumsCache = null;
  return album;
}

async function refreshFavs() {
  const kids = favAlbum ? await graphAll(`${GRAPH}/me/drive/items/${favAlbum.id}/children?$select=id,name&$top=1000`) : [];
  favByName = new Map(kids.map((k) => [k.name, k.id]));
}
async function loadFavs() {
  if (favByName) return favByName;
  try {
    const albums = await listAlbums();
    // Find it by remembered id first, so renaming the album in OneDrive ("Favoris (Maps)") keeps working.
    const savedId = store.get('favAlbumId', '');
    favAlbum = albums.find((a) => a.id === savedId) || albums.find((a) => a.name.startsWith(FAV_ALBUM)) || null;
    if (favAlbum) store.set('favAlbumId', favAlbum.id);
    await refreshFavs();
  } catch (e) { log('favs', e); favByName = new Map(); }
  renderHeart();
  return favByName;
}
const isFav = (it) => !!(it && favByName?.has(it.n));
async function setLiked(ids, on) {
  if (!(await ensureWrite())) return false;
  await loadFavs();
  const list = ids.map((id) => byId.get(id)).filter(Boolean);
  try {
    if (on) {
      const todo = list.filter((i) => !isFav(i)).map((i) => i.id);
      if (!favAlbum) {
        const a = await createAlbum(FAV_ALBUM, todo);
        favAlbum = { id: a.id, name: a.name, url: albumUrl(a.id) };
        store.set('favAlbumId', a.id);
      }
      else if (todo.length) await addToAlbum(favAlbum.id, todo);
    } else if (favAlbum) {
      const entries = list.map((i) => favByName.get(i.n)).filter(Boolean);
      await graphBatch(entries.map((eid) => ({ method: 'DELETE', url: `/drive/bundles/${favAlbum.id}/children/${eid}` })), WRITE);
    }
    await refreshFavs();
    albumsCache = null;
    return true;
  } catch (e) {
    console.error(e);
    toast("OneDrive a refusé la modification. Réessaie dans un instant.");
    return false;
  }
}

/* Album picker: add the given photos to an existing album or a new one. */
async function openAlbumPicker(ids) {
  if (!ids.length || !(await ensureWrite())) return;
  openModal(`Ajouter ${ids.length > 1 ? plural(ids.length, 'photo', 'photos') : 'la photo'} à un album`, `
    <form class="new-album" id="newAlbumForm">
      <input id="newAlbumName" placeholder="Nouvel album" maxlength="120" autocomplete="off">
      <button class="pill-btn" type="submit">Créer</button>
    </form>
    <div class="album-list" id="albumList"><p class="modal-text fine">Chargement des albums…</p></div>`);
  $('newAlbumForm').onsubmit = async (e) => {
    e.preventDefault();
    const name = $('newAlbumName').value.trim();
    if (!name) return $('newAlbumName').focus();
    busy(true);
    try {
      const a = await createAlbum(name, ids);
      closeModal(); setSelecting(false);
      toast(`Album « ${esc(a.name)} » créé`, albumUrl(a.id));
    } catch (err) { console.error(err); toast("Impossible de créer l'album. Réessaie dans un instant."); }
    busy(false);
  };
  try {
    await loadFavs();
    const albums = (await listAlbums()).filter((a) => a.id !== favAlbum?.id);
    $('albumList').innerHTML = albums.length ? '' : '<p class="modal-text fine">Aucun album pour l’instant. Crée le premier ci-dessus.</p>';
    for (const a of albums) {
      const row = document.createElement('button');
      row.className = 'album-row';
      row.innerHTML = `<span class="album-thumb"></span><span class="album-text"><strong>${esc(a.name)}</strong><small>${plural(a.count, 'élément', 'éléments')}</small></span>`;
      row.onclick = async () => {
        busy(true);
        try {
          const n = await addToAlbum(a.id, ids);
          albumsCache = null;
          closeModal(); setSelecting(false);
          const already = ids.length - n;
          toast(`${n ? plural(n, 'photo ajoutée', 'photos ajoutées') : 'Rien de nouveau'} à « ${esc(a.name)} »${already > 0 ? ` (${plural(already, 'y était déjà', 'y étaient déjà')})` : ''}`, a.url);
        } catch (err) { console.error(err); toast("OneDrive a refusé l'ajout. Réessaie dans un instant."); }
        busy(false);
      };
      $('albumList').appendChild(row);
      albumCover(a.id, row.querySelector('.album-thumb'));
    }
  } catch (e) { console.error(e); $('albumList').innerHTML = '<p class="modal-text fine">Impossible de charger tes albums.</p>'; }
}

/* Albums browser: show an album's photos on the map. */
async function openAlbums() {
  openModal('Albums', '<div class="album-list" id="albumList"><p class="modal-text fine">Chargement des albums…</p></div>');
  try {
    const albums = await listAlbums(true);
    $('albumList').innerHTML = albums.length ? '' : '<p class="modal-text fine">Aucun album dans ton OneDrive.</p>';
    for (const a of albums) {
      const row = document.createElement('div');
      row.className = 'album-row';
      row.innerHTML = `<span class="album-thumb"></span>
        <button class="album-text"><strong>${esc(a.name)}</strong><small>${plural(a.count, 'élément', 'éléments')}</small></button>
        <a class="icon-btn" href="${esc(a.url)}" target="_blank" rel="noopener" aria-label="Ouvrir « ${esc(a.name)} » dans OneDrive">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </a>`;
      row.querySelector('.album-text').onclick = async () => {
        busy(true);
        try {
          // Album photos carry their own location; they can be copies stored outside "Images".
          const kids = await graphAll(`${GRAPH}/me/drive/items/${a.id}/children?$top=1000&$select=${SELECT}`);
          const list = kids.map(toItem).filter(Boolean);
          closeModal();
          if (!list.length) { toast(`Aucune photo géolocalisée dans « ${esc(a.name)} »`, a.url); }
          else {
            applySetFilter({ kind: 'album', label: a.name, list, ids: new Set(list.map((i) => i.id)) }, false);
            openSheet(list, { title: a.name, fit: true });
          }
        } catch (e) { console.error(e); toast("Impossible d'ouvrir l'album."); }
        busy(false);
      };
      $('albumList').appendChild(row);
      albumCover(a.id, row.querySelector('.album-thumb'));
    }
  } catch (e) { console.error(e); $('albumList').innerHTML = '<p class="modal-text fine">Impossible de charger tes albums.</p>'; }
}
async function albumCover(albumId, el) {
  try {
    const r = await graph(`/me/drive/items/${albumId}/children?$top=1&$select=id`);
    if (r.value[0]) setBg(el, r.value[0].id);
  } catch {}
}

/* ================= Search: places (Photon) + photo content (OneDrive search) ================= */
let searchTimer = null, searchSeq = 0;
function openSearch() {
  $('search').hidden = false;
  closePopover();
  setTimeout(() => $('searchInput').focus(), 30);
}
function closeSearch() {
  $('search').hidden = true;
  $('searchInput').value = '';
  $('results').hidden = true;
  $('searchInput').blur();
}
$('searchInput').addEventListener('input', () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(runSearch, 280);
});
$('searchInput').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); $('results').querySelector('button')?.click(); }
  if (e.key === 'Escape') closeSearch();
});

function nearby(lat, lon, extent) {
  if (extent) {
    const [w, n, e, s] = extent;
    return items.filter((i) => i.lo >= w && i.lo <= e && i.la >= s && i.la <= n);
  }
  const d = 0.12; // ~13 km
  return items.filter((i) => Math.abs(i.la - lat) < d && Math.abs(i.lo - lon) < d / Math.cos((lat * Math.PI) / 180));
}

async function runSearch() {
  const q = $('searchInput').value.trim();
  const box = $('results');
  if (q.length < 2) { box.hidden = true; return; }
  const seq = ++searchSeq;
  box.hidden = false;
  box.innerHTML = '';
  const photoRow = document.createElement('button');
  photoRow.className = 'result';
  photoRow.innerHTML = `<span class="r-icon photo"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="5" width="17" height="14" rx="2.5" fill="none" stroke="currentColor" stroke-width="2"/><path d="M6 17l4-4.5 3 3 2-2 3 3.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg></span>
    <span class="r-text"><strong>Photos : « ${esc(q)} »</strong><small>Ce qu'il y a sur la photo ou son nom</small></span>`;
  photoRow.onclick = () => photoSearch(q);
  box.appendChild(photoRow);

  try {
    const c = map.getCenter();
    const bias = map.getZoom() > 4 ? `&lat=${c.lat.toFixed(3)}&lon=${wrapLng(c.lng).toFixed(3)}` : '';
    const r = await fetch(`https://photon.komoot.io/api/?q=${encodeURIComponent(q)}&limit=6&lang=fr${bias}`);
    const j = await r.json();
    if (seq !== searchSeq) return;
    for (const f of j.features || []) {
      const p = f.properties;
      const [lon, lat] = f.geometry.coordinates;
      const mine = nearby(lat, lon, p.extent);
      const sub = [p.city !== p.name ? p.city : null, p.state, p.country].filter(Boolean).join(', ');
      const row = document.createElement('button');
      row.className = 'result';
      row.innerHTML = `<span class="r-icon place"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 21s-6.5-6.2-6.5-11a6.5 6.5 0 0 1 13 0c0 4.8-6.5 11-6.5 11z" fill="currentColor"/><circle cx="12" cy="10" r="2.4" fill="#fff"/></svg></span>
        <span class="r-text"><strong>${esc(p.name || sub)}</strong><small>${esc(sub)}</small></span>
        ${mine.length ? `<span class="r-count">${plural(mine.length, 'photo', 'photos')}</span>` : ''}`;
      row.onclick = () => {
        closeSearch();
        if (p.extent) { const [w, n, e, s] = p.extent; map.fitBounds([[w, s], [e, n]], { padding: 60, maxZoom: 15, duration: 900 }); }
        else map.flyTo({ center: [lon, lat], zoom: 13, duration: 900 });
        if (mine.length) setTimeout(() => openSheet(mine, { title: p.name || sub }), 950);
      };
      box.appendChild(row);
    }
  } catch (e) { log('photon', e); }
}

/* OneDrive tags photos automatically, sometimes in French, sometimes in English
   ("neige" finds 6 photos, "snow" 97). Search both languages. */
const TAG_PAIRS = `chien:dog, chat:cat, plage:beach, neige:snow, montagne:mountain, mer:sea, lac:lake, riviere:river,
foret:forest, arbre:tree, fleur:flower, ciel:sky, nuage:cloud, coucher de soleil:sunset, lever de soleil:sunrise, nuit:night,
ville:city, rue:street, batiment:building, eglise:church, pont:bridge, voiture:car, velo:bicycle, bateau:boat, avion:airplane,
nourriture:food, repas:food, gateau:cake, cafe:coffee, biere:beer, vin:wine, fete:party,
ultimate:frisbee, football:soccer, ski:skiing, randonnee:hiking, piscine:swimming pool,
herbe:grass, jardin:garden, parc:park, musee:museum, chateau:castle, plante:plant, oiseau:bird, cheval:horse, poisson:fish,
enfant:child, bebe:baby, personne:person, groupe:group, sourire:smile, texte:text, capture d'ecran:screenshot, glace:ice,
hiver:winter, ete:summer, automne:autumn, feu:fire, soleil:sun, lune:moon, livre:book, ordinateur:computer, telephone:phone,
sable:sand, vague:wave, ile:island, cascade:waterfall, route:road, stade:stadium, marche:market`
  .split(/\s*,\s*/).map((p) => p.split(':').map((s) => s.trim()));
const norm = (s) => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
function searchTerms(q) {
  const n = norm(q), single = n.replace(/s$/, '');
  const terms = new Set([q.trim()]);
  for (const [fr, en] of TAG_PAIRS) {
    if (n === fr || single === fr || n === fr + 'x') terms.add(en);
    if (n === en || single === en) terms.add(fr);
  }
  return [...terms];
}

async function photoSearch(q) {
  busy(true);
  status('Recherche dans tes photos…');
  try {
    const terms = searchTerms(q);
    const url = (t) => `/me/drive/root/search(q='${encodeURIComponent(t.replace(/'/g, "''"))}')?$select=id&$top=200`;
    const ids = new Set();
    const collect = (res) => res.forEach((r) => (r?.body?.value || []).forEach((h) => byId.has(h.id) && ids.add(h.id)));
    const first = await graphBatch(terms.map((t) => ({ method: 'GET', url: url(t) })), READ, 2);
    collect(first);
    const capped = terms.filter((_, i) => (first[i]?.body?.value || []).length >= 200);
    closeSearch();
    busy(false);
    if (!ids.size) {
      status('');
      toast(`Aucune photo pour « ${esc(q)} ». La recherche utilise les étiquettes de OneDrive (objets, lieux), pas les visages.`);
      return;
    }
    const show = (fit) => {
      applySetFilter({ kind: 'search', label: `« ${q} »`, ids: new Set(ids) }, false);
      openSheet(items.filter((i) => ids.has(i.id)), { title: `« ${q} »`, fit });
    };
    show(true);
    // OneDrive search stops at 200 results per word and ignores folder scoping, so say so.
    status(capped.length ? 'OneDrive limite la recherche à 200 résultats par mot' : '', true);
  } catch (e) { console.error(e); status(''); toast('La recherche OneDrive a échoué. Réessaie.'); }
  finally { busy(false); }
}

/* ================= Modal, popover, toast, status ================= */
function openModal(title, html) {
  $('modalTitle').textContent = title;
  $('modalBody').innerHTML = html;
  $('modal').classList.add('open');
  $('modal').setAttribute('aria-hidden', 'false');
}
function closeModal() { $('modal').classList.remove('open'); $('modal').setAttribute('aria-hidden', 'true'); }
$('modalClose').onclick = closeModal;
$('modal').addEventListener('click', (e) => { if (e.target === $('modal')) closeModal(); });

function togglePopover() { $('stylePop').hidden = !$('stylePop').hidden; }
function closePopover() { $('stylePop').hidden = true; }

let toastTimer = null;
function toast(html, link) {
  const t = $('toast');
  t.innerHTML = `<span>${html}</span>${link ? `<a href="${esc(link)}" target="_blank" rel="noopener">Ouvrir</a>` : ''}`;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), link ? 6000 : 3200);
}
function busy(on) { document.body.classList.toggle('busy', on); }

let statusTimer = null;
function status(text, done = false) {
  const s = $('status');
  clearTimeout(statusTimer);
  if (!text) { s.hidden = true; return; }
  s.hidden = false;
  s.classList.toggle('done', done);
  $('statusText').textContent = text;
  if (done) statusTimer = setTimeout(() => (s.hidden = true), 2200);
}

document.addEventListener('keydown', (e) => {
  if ($('viewer').classList.contains('open')) {
    if (e.key === 'ArrowRight') stepViewer(1);
    else if (e.key === 'ArrowLeft') stepViewer(-1);
    else if (e.key === 'Escape') closeViewer();
  } else if (e.key === 'Escape') { closeModal(); closePopover(); closeSheet(); }
});

/* ================= Boot ================= */
async function main() {
  $('signinBtn').onclick = signIn;
  $('fitBtn').onclick = () => fitAll();
  $('styleBtn').onclick = togglePopover;
  $('albumsBtn').onclick = () => { closePopover(); openAlbums(); };
  $('searchBtn').onclick = openSearch;
  $('searchCancel').onclick = closeSearch;
  $('sheetClose').onclick = closeSheet;
  $('viewerClose').onclick = closeViewer;

  let acct = null;
  try { acct = await initAuth(); } catch (e) { console.error(e); }
  if (!acct) { $('signin').hidden = false; return; }

  $('topbar').hidden = false;
  $('controls').hidden = false;
  await initMap();
  map.on('click', closePopover);
  if ((await db.getMeta('version')) !== DATA_VERSION) await db.reset();
  const cached = await db.all();
  const firstRun = !cached.length;
  setItems(cached);
  if (!firstRun) map.once('load', () => fitAll(false));

  await runSync(firstRun);

  // Coming back to the app (iOS keeps it in memory): refresh thumbnails and pull new photos.
  let lastActive = Date.now();
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'visible') { lastActive = Date.now(); return; }
    const away = Date.now() - lastActive;
    if (away > 2 * 3600e3) { thumbCache.clear(); clearMarkers(); scheduleMarkers(); }
    if (away > 5 * 60e3) runSync(false);
  });
}

let syncing = false, syncRetries = 0;
async function runSync(firstRun) {
  if (syncing) return;
  syncing = true;
  status(firstRun ? 'Lecture de ton OneDrive…' : 'Mise à jour…');
  try {
    let shown = 0, fitted = false;
    const { changed } = await sync(async (seen, ch) => {
      if (!firstRun) return;
      status(`Lecture de ton OneDrive… ${fmtNum.format(seen)} fichiers`);
      if (ch - shown >= 500) { // show photos as they arrive on the first run
        shown = ch;
        setItems(await db.all());
        if (!fitted) { fitted = true; fitAll(); }
      }
    });
    const fresh = await db.all();
    if (firstRun || changed || fresh.length !== items.length) {
      setItems(fresh);
      if (firstRun) (map.loaded() ? fitAll() : map.once('load', () => fitAll()));
    }
    status(fresh.length ? 'À jour' : 'Aucune photo géolocalisée trouvée', true);
  } catch (e) {
    console.error(e);
    const denied = e.status === 401 || e.status === 403;
    status(denied ? 'Accès refusé par OneDrive — reconnecte-toi' : 'Connexion interrompue — nouvel essai dans 15 s', true);
    if (!denied && (syncRetries = (syncRetries || 0) + 1) <= 3) setTimeout(() => runSync(firstRun), 15000);
  } finally {
    syncing = false;
  }
}

main();
