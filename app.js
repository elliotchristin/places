/* Lieux — personal photo map over OneDrive (Microsoft Graph), no backend. */
'use strict';

const CLIENT_ID = '2255226b-3bd7-48fd-bed7-e45d25432566';
const SCOPES = ['Files.Read'];
const GRAPH = 'https://graph.microsoft.com/v1.0';
const PIN_SIZE = 'c360x360_crop';
const VIEW_SIZE = 'c2048x2048';
const STYLES = {
  light: 'https://tiles.openfreemap.org/styles/liberty',
  dark: 'https://tiles.openfreemap.org/styles/dark',
};
const DEBUG = new URLSearchParams(location.search).has('debug');

const $ = (id) => document.getElementById(id);
const log = (...a) => DEBUG && console.log('[lieux]', ...a);
const fmtDate = new Intl.DateTimeFormat('fr-CA', { day: 'numeric', month: 'long', year: 'numeric' });
const fmtMonth = new Intl.DateTimeFormat('fr-CA', { month: 'long', year: 'numeric' });
const fmtNum = new Intl.NumberFormat('fr-CA');
const VID_ICON = '<svg class="vid" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z" fill="currentColor"/></svg>';

/* ---------------- Auth ---------------- */
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
function signIn() {
  return msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
}
async function getToken() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    log('silent token failed', e);
    await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
    return new Promise(() => {}); // page is navigating away
  }
}

/* ---------------- Graph ---------------- */
async function graph(url, opts = {}, attempt = 0) {
  const token = await getToken();
  const res = await fetch(url.startsWith('http') ? url : GRAPH + url, {
    ...opts,
    headers: { Authorization: 'Bearer ' + token, ...(opts.body ? { 'Content-Type': 'application/json' } : {}), ...(opts.headers || {}) },
  });
  if ((res.status === 429 || res.status >= 500) && attempt < 5) {
    const wait = (Number(res.headers.get('Retry-After')) || 2 ** attempt) * 1000;
    await new Promise((r) => setTimeout(r, wait));
    return graph(url, opts, attempt + 1);
  }
  if (!res.ok) {
    const err = new Error(`Graph ${res.status}`);
    err.status = res.status;
    err.body = await res.text().catch(() => '');
    throw err;
  }
  return res.json();
}

/* ---------------- Local store (IndexedDB) ---------------- */
const db = {
  _db: null,
  open() {
    if (this._db) return Promise.resolve(this._db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('lieux', 1);
      req.onupgradeneeded = () => {
        req.result.createObjectStore('items', { keyPath: 'id' });
        req.result.createObjectStore('meta');
      };
      req.onsuccess = () => resolve((this._db = req.result));
      req.onerror = () => reject(req.error);
    });
  },
  async tx(store, mode, fn) {
    const d = await this.open();
    return new Promise((resolve, reject) => {
      const t = d.transaction(store, mode);
      const s = t.objectStore(store);
      let out;
      Promise.resolve(fn(s)).then((v) => (out = v));
      t.oncomplete = () => resolve(out);
      t.onerror = () => reject(t.error);
    });
  },
  all() {
    return this.tx('items', 'readonly', (s) => new Promise((r) => { const q = s.getAll(); q.onsuccess = () => r(q.result); }));
  },
  getMeta(k) {
    return this.tx('meta', 'readonly', (s) => new Promise((r) => { const q = s.get(k); q.onsuccess = () => r(q.result); }));
  },
  setMeta(k, v) { return this.tx('meta', 'readwrite', (s) => { s.put(v, k); }); },
  apply(puts, dels) {
    return this.tx('items', 'readwrite', (s) => { puts.forEach((p) => s.put(p)); dels.forEach((id) => s.delete(id)); });
  },
  async reset() { await this.tx('items', 'readwrite', (s) => s.clear()); await this.tx('meta', 'readwrite', (s) => s.clear()); },
};

/* ---------------- Sync (delta query) ---------------- */
const SELECT = 'id,name,file,location,photo,video,deleted,webUrl,fileSystemInfo';

function toItem(it) {
  const loc = it.location;
  if (it.deleted || !it.file || !loc || loc.latitude == null || loc.longitude == null) return null;
  if (Math.abs(loc.latitude) < 1e-6 && Math.abs(loc.longitude) < 1e-6) return null; // null island
  const taken = it.photo?.takenDateTime || it.fileSystemInfo?.createdDateTime;
  return {
    id: it.id,
    la: +loc.latitude.toFixed(6),
    lo: +loc.longitude.toFixed(6),
    t: taken ? Date.parse(taken) : 0,
    v: !!(it.video || (it.file.mimeType || '').startsWith('video/')) ? 1 : 0,
    u: it.webUrl,
  };
}

async function sync(onProgress) {
  let url = (await db.getMeta('resumeLink')) || (await db.getMeta('deltaLink'))
    || `${GRAPH}/me/drive/root/delta?$select=${SELECT}`;
  let seen = 0, changed = 0;
  while (url) {
    let page;
    try {
      page = await graph(url);
    } catch (e) {
      if (e.status === 410) { // delta token expired: start over
        await db.reset();
        url = `${GRAPH}/me/drive/root/delta?$select=${SELECT}`;
        continue;
      }
      throw e;
    }
    const puts = [], dels = [];
    for (const it of page.value) {
      const item = toItem(it);
      if (item) puts.push(item); else dels.push(it.id);
    }
    await db.apply(puts, dels);
    seen += page.value.length;
    changed += puts.length;
    onProgress?.(seen, changed);
    if (page['@odata.nextLink']) {
      url = page['@odata.nextLink'];
      await db.setMeta('resumeLink', url);
    } else {
      await db.setMeta('deltaLink', page['@odata.deltaLink']);
      await db.setMeta('resumeLink', null);
      await db.setMeta('lastSync', Date.now());
      url = null;
    }
  }
  return { seen, changed };
}

/* ---------------- Thumbnails (batched) ---------------- */
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
async function flushThumbs() {
  thumbTimer = null;
  while (thumbQueue.length && inflight < 4) {
    const batch = thumbQueue.splice(0, 20);
    inflight++;
    runBatch(batch).finally(() => { inflight--; if (thumbQueue.length) flushThumbs(); });
  }
}
async function runBatch(batch) {
  try {
    const body = {
      requests: batch.map((b, i) => ({ id: String(i), method: 'GET', url: `/me/drive/items/${b.id}/thumbnails?select=${b.size}` })),
    };
    const res = await graph('/$batch', { method: 'POST', body: JSON.stringify(body) });
    const byId = new Map(res.responses.map((r) => [r.id, r]));
    batch.forEach((b, i) => {
      const r = byId.get(String(i));
      const url = r && r.status === 200 ? r.body?.value?.[0]?.[b.size]?.url : null;
      if (!url) thumbCache.delete(b.id + '|' + b.size);
      b.resolve(url || null);
    });
  } catch (e) {
    log('thumb batch failed', e);
    batch.forEach((b) => { thumbCache.delete(b.id + '|' + b.size); b.resolve(null); });
  }
}
function setBg(el, id, size) {
  return thumb(id, size).then((url) => { if (url) el.style.backgroundImage = `url("${url}")`; return url; });
}

/* ---------------- Reverse geocoding (place names) ---------------- */
const placeCache = new Map();
function placeName(lat, lon) {
  const key = lat.toFixed(2) + ',' + lon.toFixed(2);
  if (placeCache.has(key)) return placeCache.get(key);
  const p = fetch(`https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=${lat}&longitude=${lon}&localityLanguage=fr`)
    .then((r) => (r.ok ? r.json() : null))
    .then((j) => {
      if (!j) return '';
      const city = j.city || j.locality || j.principalSubdivision || '';
      const country = j.countryName || '';
      return [city, country].filter(Boolean).filter((v, i, a) => a.indexOf(v) === i).join(', ');
    })
    .catch(() => '');
  placeCache.set(key, p);
  return p;
}

/* ---------------- Map ---------------- */
let map, items = [], byId = new Map(), byT = new Map(), yearFilter = 'all';
const markers = new Map();
const dark = matchMedia('(prefers-color-scheme: dark)');

function frenchLabels() {
  for (const layer of map.getStyle().layers) {
    if (layer.type !== 'symbol') continue;
    const tf = map.getLayoutProperty(layer.id, 'text-field');
    if (!tf || !JSON.stringify(tf).includes('name')) continue;
    map.setLayoutProperty(layer.id, 'text-field', ['coalesce', ['get', 'name:fr'], ['get', 'name:latin'], ['get', 'name']]);
  }
}

function addPhotoLayer() {
  if (map.getSource('photos')) return;
  map.addSource('photos', {
    type: 'geojson',
    data: toGeoJSON(),
    cluster: true,
    clusterRadius: 64,
    clusterMaxZoom: 17,
    clusterProperties: { latest: ['max', ['get', 't']] },
  });
  // Invisible layer so the source tiles get loaded and can be queried.
  map.addLayer({ id: 'photos-hit', type: 'circle', source: 'photos', paint: { 'circle-radius': 1, 'circle-opacity': 0 } });
}

function initMap() {
  map = new maplibregl.Map({
    container: 'map',
    style: dark.matches ? STYLES.dark : STYLES.light,
    center: [-71.9, 45.4],
    zoom: 3,
    attributionControl: { compact: true },
    dragRotate: false,
    pitchWithRotate: false,
  });
  map.touchZoomRotate.disableRotation();
  map.on('style.load', () => { frenchLabels(); addPhotoLayer(); clearMarkers(); });
  map.on('render', scheduleMarkers);
  dark.addEventListener('change', () => map.setStyle(dark.matches ? STYLES.dark : STYLES.light));
}

function filtered() {
  return yearFilter === 'all' ? items : items.filter((i) => new Date(i.t).getFullYear() === yearFilter);
}
function toGeoJSON() {
  return {
    type: 'FeatureCollection',
    features: filtered().map((i) => ({ type: 'Feature', geometry: { type: 'Point', coordinates: [i.lo, i.la] }, properties: { id: i.id, t: i.t, v: i.v } })),
  };
}
function refreshSource() {
  const src = map?.getSource('photos');
  if (src) { src.setData(toGeoJSON()); clearMarkers(); }
}

let markerFrame = null;
function scheduleMarkers() {
  if (markerFrame) return;
  markerFrame = requestAnimationFrame(() => { markerFrame = null; updateMarkers(); });
}
function clearMarkers() { for (const m of markers.values()) m.remove(); markers.clear(); }

function updateMarkers() {
  if (!map.getSource('photos') || !map.isSourceLoaded('photos')) return;
  const seen = new Set();
  for (const f of map.querySourceFeatures('photos')) {
    const p = f.properties;
    const key = p.cluster ? 'c' + p.cluster_id + '@' + Math.floor(map.getZoom()) : 'p' + p.id;
    if (seen.has(key)) continue;
    seen.add(key);
    if (markers.has(key)) continue;
    const el = buildPin(p);
    const m = new maplibregl.Marker({ element: el, anchor: 'bottom', offset: [0, -8] }).setLngLat(f.geometry.coordinates).addTo(map);
    el.addEventListener('click', (ev) => { ev.stopPropagation(); onPinClick(p, f.geometry.coordinates); });
    markers.set(key, m);
  }
  for (const [k, m] of markers) if (!seen.has(k)) { m.remove(); markers.delete(k); }
}

function buildPin(p) {
  const el = document.createElement('button');
  el.className = 'pin';
  const count = p.cluster ? p.point_count : 1;
  const rep = p.cluster ? byT.get(p.latest) : byId.get(p.id);
  if (count === 1) el.classList.add('small');
  if (count > 1) el.innerHTML = `<span class="count">${count >= 1000 ? Math.round(count / 100) / 10 + 'k' : count}</span>`;
  else if (rep?.v) el.innerHTML = VID_ICON;
  el.setAttribute('aria-label', count > 1 ? `${count} photos` : 'Photo');
  const fallback = async () => { // representative has no thumbnail: try other photos of the cluster
    if (!p.cluster) return;
    const leaves = await map.getSource('photos').getClusterLeaves(p.cluster_id, 12, 0).catch(() => []);
    for (const l of leaves) if (l.properties.id !== rep?.id && (await setBg(el, l.properties.id, PIN_SIZE))) return;
  };
  if (rep) setBg(el, rep.id, PIN_SIZE).then((u) => u || fallback());
  else fallback();
  return el;
}

async function onPinClick(p, coords) {
  if (!p.cluster) {
    const it = byId.get(p.id);
    if (it) openViewer([it], 0);
    return;
  }
  const src = map.getSource('photos');
  if (map.getZoom() < 9) {
    const z = await src.getClusterExpansionZoom(p.cluster_id);
    map.easeTo({ center: coords, zoom: Math.min(z + 0.5, 17), duration: 600 });
    return;
  }
  const leaves = await src.getClusterLeaves(p.cluster_id, Infinity, 0);
  const list = leaves.map((l) => byId.get(l.properties.id)).filter(Boolean);
  openSheet(list, coords);
}

function fitAll(animate = true) {
  const list = filtered();
  if (!list.length) return;
  const b = new maplibregl.LngLatBounds();
  list.forEach((i) => b.extend([i.lo, i.la]));
  map.fitBounds(b, { padding: { top: 120, bottom: 90, left: 50, right: 50 }, maxZoom: 13, duration: animate ? 800 : 0 });
}

/* ---------------- UI: summary + years ---------------- */
function renderChrome() {
  const list = filtered();
  const photos = list.filter((i) => !i.v).length, videos = list.length - photos;
  const parts = [`${fmtNum.format(photos)} photo${photos > 1 ? 's' : ''}`];
  if (videos) parts.push(`${fmtNum.format(videos)} vidéo${videos > 1 ? 's' : ''}`);
  $('summary').textContent = parts.join(' · ');

  const years = [...new Set(items.map((i) => new Date(i.t).getFullYear()))].filter((y) => y > 1990).sort((a, b) => b - a);
  const nav = $('years');
  nav.hidden = years.length < 2;
  nav.innerHTML = '';
  for (const y of ['all', ...years]) {
    const b = document.createElement('button');
    b.textContent = y === 'all' ? 'Tout' : y;
    b.setAttribute('aria-pressed', String(y === yearFilter));
    b.onclick = () => { yearFilter = y; renderChrome(); refreshSource(); fitAll(); closeSheet(); };
    nav.appendChild(b);
  }
}

/* ---------------- Sheet ---------------- */
let sheetList = [];
const io = new IntersectionObserver((entries) => {
  for (const e of entries) {
    if (!e.isIntersecting) continue;
    io.unobserve(e.target);
    setBg(e.target, e.target.dataset.id, PIN_SIZE);
  }
}, { root: $('grid'), rootMargin: '300px' });

function openSheet(list, coords) {
  sheetList = list.slice().sort((a, b) => b.t - a.t);
  const n = sheetList.length;
  const first = sheetList[n - 1], last = sheetList[0];
  $('sheetTitle').textContent = '…';
  const range = first && last ? (fmtMonth.format(first.t) === fmtMonth.format(last.t) ? fmtMonth.format(last.t) : `${fmtMonth.format(first.t)} – ${fmtMonth.format(last.t)}`) : '';
  $('sheetSub').textContent = `${fmtNum.format(n)} élément${n > 1 ? 's' : ''} · ${range}`;
  const lat = sheetList.reduce((s, i) => s + i.la, 0) / n, lon = sheetList.reduce((s, i) => s + i.lo, 0) / n;
  placeName(lat, lon).then((name) => { if (sheetList.length === n) $('sheetTitle').textContent = name || 'Ici'; });

  const grid = $('grid');
  io.disconnect();
  grid.innerHTML = '';
  grid.scrollTop = 0;
  let month = '';
  sheetList.forEach((it, idx) => {
    const m = fmtMonth.format(it.t);
    if (m !== month && n > 12) {
      month = m;
      const h = document.createElement('div');
      h.className = 'month';
      h.textContent = m.charAt(0).toUpperCase() + m.slice(1);
      grid.appendChild(h);
    }
    const c = document.createElement('button');
    c.className = 'cell';
    c.dataset.id = it.id;
    c.setAttribute('aria-label', fmtDate.format(it.t));
    if (it.v) c.innerHTML = VID_ICON;
    c.onclick = () => openViewer(sheetList, idx);
    grid.appendChild(c);
    io.observe(c);
  });
  $('sheet').classList.add('open');
  $('sheet').setAttribute('aria-hidden', 'false');

  const b = new maplibregl.LngLatBounds();
  sheetList.forEach((i) => b.extend([i.lo, i.la]));
  const wide = innerWidth >= 760;
  map.fitBounds(b, { padding: wide ? { top: 120, bottom: 60, left: 460, right: 60 } : { top: 110, bottom: innerHeight * 0.62 + 30, left: 40, right: 40 }, maxZoom: 16, duration: 700 });
}
function closeSheet() {
  $('sheet').classList.remove('open');
  $('sheet').setAttribute('aria-hidden', 'true');
}

// Swipe the sheet down to close.
(() => {
  let y0 = null;
  const s = $('sheet');
  s.addEventListener('touchstart', (e) => { if (e.target.closest('.grid') && $('grid').scrollTop > 0) return; y0 = e.touches[0].clientY; }, { passive: true });
  s.addEventListener('touchend', (e) => { if (y0 != null && e.changedTouches[0].clientY - y0 > 80) closeSheet(); y0 = null; });
})();

/* ---------------- Viewer ---------------- */
let vList = [], vIdx = 0;
function openViewer(list, idx) {
  vList = list; vIdx = idx;
  $('viewer').classList.add('open');
  $('viewer').setAttribute('aria-hidden', 'false');
  showViewer();
}
function showViewer() {
  const it = vList[vIdx];
  if (!it) return;
  const img = $('viewerImg');
  img.classList.add('loading');
  const small = thumbCache.get(it.id + '|' + PIN_SIZE);
  small?.then((u) => { if (u && vList[vIdx] === it && !img.dataset.full) img.src = u; });
  img.dataset.full = '';
  thumb(it.id, VIEW_SIZE).then((u) => {
    if (vList[vIdx] !== it) return;
    if (u) { img.src = u; img.dataset.full = '1'; }
    img.onload = () => img.classList.remove('loading');
  });
  $('viewerPlay').hidden = !it.v;
  $('viewerDate').textContent = it.t ? fmtDate.format(it.t) : '';
  $('viewerPlace').textContent = '';
  placeName(it.la, it.lo).then((n) => { if (vList[vIdx] === it) $('viewerPlace').textContent = n; });
  $('viewerOpen').href = it.u;
  $('viewerCount').textContent = vList.length > 1 ? `${vIdx + 1} sur ${fmtNum.format(vList.length)}` : '';
  // Preload neighbours.
  [vIdx - 1, vIdx + 1].forEach((j) => vList[j] && thumb(vList[j].id, VIEW_SIZE).then((u) => { if (u) new Image().src = u; }));
}
function stepViewer(d) {
  const j = vIdx + d;
  if (j < 0 || j >= vList.length) return;
  vIdx = j;
  showViewer();
}
function closeViewer() {
  $('viewer').classList.remove('open');
  $('viewer').setAttribute('aria-hidden', 'true');
  $('viewerImg').removeAttribute('src');
}
(() => {
  let x0 = null, y0 = null;
  const st = $('viewerStage');
  st.addEventListener('touchstart', (e) => { x0 = e.touches[0].clientX; y0 = e.touches[0].clientY; }, { passive: true });
  st.addEventListener('touchend', (e) => {
    if (x0 == null) return;
    const dx = e.changedTouches[0].clientX - x0, dy = e.changedTouches[0].clientY - y0;
    if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) stepViewer(dx < 0 ? 1 : -1);
    else if (dy > 90) closeViewer();
    x0 = y0 = null;
  });
  st.addEventListener('click', (e) => {
    if (vList[vIdx]?.v) { window.open(vList[vIdx].u, '_blank', 'noopener'); return; }
    const r = st.getBoundingClientRect();
    stepViewer(e.clientX < r.left + r.width / 3 ? -1 : e.clientX > r.right - r.width / 3 ? 1 : 0);
  });
})();
document.addEventListener('keydown', (e) => {
  if ($('viewer').classList.contains('open')) {
    if (e.key === 'ArrowRight') stepViewer(1);
    else if (e.key === 'ArrowLeft') stepViewer(-1);
    else if (e.key === 'Escape') closeViewer();
  } else if (e.key === 'Escape') closeSheet();
});

/* ---------------- Status ---------------- */
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

/* ---------------- Boot ---------------- */
function setItems(list) {
  items = list;
  byId = new Map(list.map((i) => [i.id, i]));
  byT = new Map();
  for (const i of list) if (!byT.has(i.t)) byT.set(i.t, i);
  renderChrome();
  refreshSource();
}

async function main() {
  $('signinBtn').onclick = signIn;
  $('fitBtn').onclick = () => fitAll();
  $('sheetClose').onclick = closeSheet;
  $('viewerClose').onclick = closeViewer;

  let acct;
  try {
    acct = await initAuth();
  } catch (e) {
    console.error(e);
  }
  if (!acct) { $('signin').hidden = false; return; }

  $('topbar').hidden = false;
  initMap();
  const cached = await db.all();
  const firstRun = !cached.length;
  setItems(cached);
  if (!firstRun) map.once('load', () => fitAll(false));

  status(firstRun ? 'Lecture de ton OneDrive…' : 'Mise à jour…');
  try {
    let shown = 0;
    const { changed } = await sync(async (seen, ch) => {
      if (!firstRun) return;
      status(`Lecture de ton OneDrive… ${fmtNum.format(seen)} fichiers`);
      if (ch - shown >= 400) { // show photos as they arrive on the first run
        shown = ch;
        setItems(await db.all());
        if (shown === ch && items.length && !map.__fitted) { map.__fitted = true; fitAll(); }
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
    status(e.status === 403 || e.status === 401 ? 'Accès refusé par OneDrive — reconnecte-toi' : 'Synchronisation interrompue — elle reprendra à la prochaine ouverture', true);
  }
}

main();
