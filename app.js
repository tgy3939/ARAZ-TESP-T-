/* Arazi Tespit — arayüz (ÇKS Excel'i ile çalışır, tutanağı kendisi üretir) */
const C = globalThis.TespitCore;
const FONTS = globalThis.TESPIT_FONTS;
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = (s) => { const n = parseInt(String(s).replace(/\D/g, ''), 10); return isNaN(n) ? 1e9 : n; };
const upper = (s) => String(s || '').toLocaleUpperCase('tr-TR');
const nz = (s) => upper(s).replace(/\s+/g, '');

// ---------- Ayarlar ----------
const SKEY = 'tespit.settings.v1';
let S = { ...C.DEFAULTS };
try {
  const old = JSON.parse(localStorage.getItem(SKEY) || '{}');
  for (const k in C.DEFAULTS) if (old[k] !== undefined) S[k] = old[k];
  if (!old.tarihMod) S.tarihMod = old.tarih ? 'sabit' : 'bos';
  S._t = old._t || 0;
} catch (e) {}
function saveSettings() { try { localStorage.setItem(SKEY, JSON.stringify(S)); } catch (e) {} }

// ---------- Cihazda saklama ----------
let dbp = null;
function idb() {
  if (dbp) return dbp;
  dbp = new Promise((res) => {
    try {
      const r = indexedDB.open('tespit', 3);
      r.onupgradeneeded = () => {
        const d = r.result;
        for (const n of ['docs', 'files', 'kv', 'photos']) if (!d.objectStoreNames.contains(n)) d.createObjectStore(n, { keyPath: 'id' });
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => res(null);
      r.onblocked = () => res(null);
    } catch (e) { res(null); }
  });
  return dbp;
}
async function tx(store, mode, fn) {
  const d = await idb();
  if (!d) throw new Error('nodb');
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode);
    let out;
    const r = fn(t.objectStore(store));
    if (r) r.onsuccess = () => { out = r.result; };
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error);
  });
}
let storageOk = true;
const safe = async (p, fallback) => { try { return await p; } catch (e) { storageOk = false; return fallback; } };
const putDoc = (d) => safe(tx('docs', 'readwrite', (st) => st.put({ id: d.id, name: d.name, kind: 'liste', year: d.year || '', farmers: d.farmers, updated: Date.now() })));
const allDocs = () => safe(tx('docs', 'readonly', (st) => st.getAll()), []);
const delDocDb = async (id) => { await safe(tx('docs', 'readwrite', (st) => st.delete(id))); await safe(tx('files', 'readwrite', (st) => st.delete(id))); };
const getMarks = () => safe(tx('kv', 'readonly', (st) => st.get('marks')), null);
const putPhoto = (ph) => safe(tx('photos', 'readwrite', (st) => st.put(ph)));
const delPhotoDb = (id) => safe(tx('photos', 'readwrite', (st) => st.delete(id)));
const allPhotos = () => safe(tx('photos', 'readonly', (st) => st.getAll()), []);
let saveTimer = null;
async function saveLocal() {
  clearTimeout(saveTimer);
  storageOk = true;
  await safe(tx('kv', 'readwrite', (st) => st.put({ id: 'marks', data: M })));
  renderSaveState();
}
function queueSave() { clearTimeout(saveTimer); saveTimer = setTimeout(saveLocal, 250); }

// ---------- Bulut kaydı ----------
// İşaretler Claude hesabındaki kişisel alana da yazılır (yalnız bu kullanıcı görür).
// Buluta TC, ad, köy, ada/parsel açık gitmez: parsel ve köy, geri çözülemeyen kısa kodlarla tutulur.
const h32 = (str, seed) => { let h = seed >>> 0; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h; };
const hkey = (str) => h32(str, 2166136261).toString(36) + h32(str, 0x9747b28c).toString(36);
const vhOf = (k) => 'm-' + hkey(String(k).split('|')[2] || '');
let cloud = null, uid = null, pulled = false;
const cloudDocs = {};          // belge adı -> { marks: { kod: işaret } }
const cloudPh = {};            // p-köy -> { photos: { kod: [kayıt] } }
const dirtyP = new Set();      // fotoğraf listesi buluta yazılmamış köyler
let assetsApi = null;
let cloudState = 'off';        // off | loading | ok | offline | error
let lastSaved = 0, settingsDirty = false;
const withTimeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej({ code: 'timeout' }), ms))]);
const userDoc = (id) => cloud.doc(`data/users/${uid}/${id}`);
const offlineErr = (e) => !navigator.onLine || (e && (e.code === 'timeout' || e.code === 'unavailable'));

async function cloudInit() {
  if (!(window.claude && typeof window.claude.use === 'function')) { renderSaveState(); return; }
  try {
    const [db, user, assets] = await Promise.all([window.claude.use('db'), window.claude.use('user'), window.claude.use('assets')]);
    assetsApi = assets || null;
    const id = user ? await user.id() : null;
    if (!db || !id) { renderSaveState(); return; }
    cloud = db; uid = id;
    await cloudPull();
  } catch (e) { console.error(e); cloudState = 'error'; renderSaveState(); }
}
async function cloudPull() {
  if (!cloud) return;
  cloudState = 'loading'; renderSaveState();
  try {
    const snap = await withTimeout(cloud.collection(`data/users/${uid}`).get(), 15000);
    const rev = new Map();
    PARCELS.forEach((_, k) => rev.set(hkey(k), k));
    for (const k in M) rev.set(hkey(k), k);
    let got = false, settings = null;
    for (const d of snap.docs) {
      const data = d.data() || {};
      if (d.id === 'settings') { settings = data; continue; }
      if (d.id.startsWith('p-')) {
        cloudPh[d.id] = { photos: { ...(data.photos || {}) } };
        for (const h in data.photos || {}) {
          const k = rev.get(h);
          if (!k) continue;
          for (const e of data.photos[h] || []) {
            if (photoById(e.id)) continue;
            if (deletedPh.has(e.id)) continue;
            const rec = { id: e.id, key: k, t: e.t, lat: e.lat, lon: e.lon, asset: e.a ? { id: e.a, url: e.u } : null, up: true, blob: null };
            addPhotoRec(rec); putPhoto(rec); got = true;
          }
        }
        continue;
      }
      if (!d.id.startsWith('m-')) continue;
      cloudDocs[d.id] = { marks: { ...(data.marks || {}) } };
      for (const h in data.marks || {}) {
        const k = rev.get(h);
        if (!k) continue;
        const a = data.marks[h], b = M[k];
        if (!b || (a.t || 0) > (b.t || 0)) { M[k] = a; got = true; }
      }
    }
    if (settings && settings.S && (settings.t || 0) > (S._t || 0)) {
      for (const k in C.DEFAULTS) if (settings.S[k] !== undefined) S[k] = settings.S[k];
      S._t = settings.t; saveSettings();
    } else if ((S._t || 0) > ((settings && settings.t) || 0)) settingsDirty = true;
    dirty.clear();
    for (const k in M) {
      const vh = vhOf(k);
      const c = cloudDocs[vh] && cloudDocs[vh].marks[hkey(k)];
      if (!c || (c.t || 0) < (M[k].t || 0)) dirty.add(vh);
    }
    PH.forEach((list) => list.forEach((ph) => { if (!ph.up) dirtyP.add(phVh(ph.key)); }));
    pulled = true;
    cloudState = 'ok';
    if (!lastSaved && !dirty.size) lastSaved = Date.now();
    if (got) { await saveLocal(); if (!$('work').hidden) renderAll(); }
    renderSaveState();
    if (dirty.size || settingsDirty || dirtyP.size) cloudPush();
  } catch (e) {
    console.error(e);
    cloudState = offlineErr(e) ? 'offline' : 'error';
    renderSaveState();
  }
}
let pushing = null, pushTimer = null;
function schedulePush() { clearTimeout(pushTimer); pushTimer = setTimeout(() => cloudPush(), 4000); renderSaveState(); }
function cloudPush() {
  clearTimeout(pushTimer);
  if (!cloud) { renderSaveState(); return Promise.resolve(); }
  if (pushing) return pushing;
  pushing = (async () => {
    try {
      if (!pulled) { await cloudPull(); if (!pulled) return; }
      for (const vh of [...dirty]) {
        const marks = { ...((cloudDocs[vh] || {}).marks || {}) };
        for (const k in M) if (vhOf(k) === vh) marks[hkey(k)] = M[k];
        await withTimeout(userDoc(vh).set({ v: 1, marks, updated: Date.now() }), 15000);
        cloudDocs[vh] = { marks };
        dirty.delete(vh);
      }
      if (dirtyP.size && assetsApi) {
        for (const vh of [...dirtyP]) {
          const own = [];
          PH.forEach((list, k) => { if (phVh(k) === vh) own.push(...list); });
          for (const ph of own) {
            if (ph.up || !ph.blob) continue;
            const r = await withTimeout(assetsApi.upload(ph.blob), 60000);
            ph.asset = { id: r.id, url: r.url };
            ph.up = true;
            await putPhoto(ph);
          }
          const photos = { ...((cloudPh[vh] || {}).photos || {}) };
          PH.forEach((list, k) => { if (phVh(k) === vh) photos[hkey(k)] = list.filter((p) => p.up && p.asset).map((p) => ({ id: p.id, a: p.asset.id, u: p.asset.url, t: p.t, lat: p.lat ?? null, lon: p.lon ?? null })); });
          for (const id of deletedPh) for (const h in photos) photos[h] = photos[h].filter((e) => e.id !== id);
          await withTimeout(userDoc(vh).set({ v: 1, photos, updated: Date.now() }), 15000);
          cloudPh[vh] = { photos };
          dirtyP.delete(vh);
        }
      }
      if (settingsDirty) {
        const out = {};
        for (const k in C.DEFAULTS) out[k] = S[k];
        await withTimeout(userDoc('settings').set({ S: out, t: S._t || Date.now() }), 15000);
        settingsDirty = false;
      }
      cloudState = 'ok';
      lastSaved = Date.now();
    } catch (e) {
      console.error(e);
      cloudState = offlineErr(e) ? 'offline' : 'error';
    } finally {
      pushing = null;
      renderSaveState();
    }
  })();
  return pushing;
}
const hhmm = (t) => new Date(t).toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });
function renderSaveState() {
  const el = $('saveState');
  if (!el) return;
  let txt, cls;
  const pend = dirty.size || settingsDirty || (assetsApi && dirtyP.size);
  if (!cloud) {
    txt = storageOk ? 'Bu tablette kayıtlı' : 'Kaydedilemiyor'; cls = storageOk ? 'ok' : 'no';
  } else if (pushing || cloudState === 'loading') {
    txt = 'Kaydediliyor…'; cls = '';
  } else if (cloudState === 'offline') {
    txt = 'İnternet yok · tablette duruyor'; cls = 'ed';
  } else if (cloudState === 'error') {
    txt = 'Buluta yazılamadı · tablette duruyor'; cls = 'ed';
  } else if (pend) {
    txt = 'Kaydedilmemiş değişiklik var'; cls = 'ed';
  } else {
    txt = lastSaved ? `Kaydedildi · ${hhmm(lastSaved)}` : 'Kaydedildi'; cls = 'ok';
  }
  el.textContent = txt;
  el.className = 'savestate ' + cls;
  const note = $('saveNote');
  if (note) note.textContent = cloud
    ? 'İşaretler tablete hemen, Claude hesabınıza internet olduğunda kaydedilir. Başka cihazda Excel\'i yükleyince kaldığınız yerden devam edersiniz.'
    : (storageOk ? 'İşaretler bu tablete kaydediliyor.' : 'Bu cihaza kaydedilemiyor. Gün sonunda tutanak PDF\'ini indirerek saklayın.');
}
async function saveNow() {
  const b = $('saveBtn');
  b.disabled = true;
  await saveLocal();
  await cloudPush();
  b.disabled = false;
  renderSaveState();
}
window.addEventListener('online', () => { if (cloud) (pulled ? cloudPush() : cloudPull()); });
document.addEventListener('visibilitychange', () => { if (document.hidden) { saveLocal(); if (cloud && (dirty.size || settingsDirty || dirtyP.size)) cloudPush(); } });

// ---------- Fotoğraflar ----------
const PH = new Map();          // parsel anahtarı -> [{ id, key, t, lat, lon, blob, asset, up }]
const deletedPh = new Set((() => { try { return JSON.parse(localStorage.getItem('tespit.phdel') || '[]'); } catch (e) { return []; } })());
const saveDeleted = () => { try { localStorage.setItem('tespit.phdel', JSON.stringify([...deletedPh])); } catch (e) {} };
const phVh = (k) => 'p-' + hkey(String(k).split('|')[2] || '');
const photoById = (id) => { for (const l of PH.values()) { const p = l.find((x) => x.id === id); if (p) return p; } return null; };
function addPhotoRec(rec) { if (!PH.has(rec.key)) PH.set(rec.key, []); const l = PH.get(rec.key); if (!l.some((x) => x.id === rec.id)) l.push(rec); l.sort((a, b) => a.t - b.t); }
const urlCache = new Map();
function phSrc(ph) {
  if (urlCache.has(ph.id)) return urlCache.get(ph.id);
  let u = '';
  if (ph.blob) u = URL.createObjectURL(ph.blob);
  else if (ph.asset) u = ph.asset.url || ('/_blob/' + ph.asset.id);
  urlCache.set(ph.id, u);
  return u;
}
async function loadPhotos() {
  const all = await allPhotos();
  for (const p of all) addPhotoRec(p);
}
const fmtCoord = (lat, lon) => `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
const exifTime = (dt) => { const m = String(dt || '').match(/^(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2})(?::(\d{2}))?/); return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)).getTime() : 0; };
const trDT = (t) => new Date(t).toLocaleString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
// Fotoğrafı küçültür ve altına parsel/tespit/konum bilgisini gömer.
async function stampPhoto(file, k, fb) {
  const buf = await file.arrayBuffer();
  const ex = C.readExif(buf);
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch (e) {
    bmp = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = URL.createObjectURL(file); });
  }
  const iw = bmp.width, ih = bmp.height;
  const sc = Math.min(1, 2048 / Math.max(iw, ih));
  const w = Math.round(iw * sc), h = Math.round(ih * sc);
  const { r, f } = PARCELS.get(k);
  const m = live(k);
  const t = C.tespit(r, m || { m: '+' }, S);
  const when = exifTime(ex.dt) || Date.now();
  const lines = [
    `${C.villageOf(r)} · Ada ${r.ada} / Parsel ${r.parsel} · ${f.ad}`,
    `Beyan: ${r.urun} ${r.alan} da  →  Tespit: ${m ? `${t.urun} ${t.alan} da${t.olc ? ' (ölçüm)' : ''}` : 'işaretlenmedi'}`,
    `${trDT(when)} · ${ex.lat != null ? 'GPS: ' + fmtCoord(ex.lat, ex.lon) : fb ? `GPS (${fb.src}): ${fmtCoord(fb.lat, fb.lon)} ±${fb.acc} m` : 'GPS: konum bilgisi yok'}`
  ];
  const fs = Math.max(16, Math.round(w / 52));
  const pad = Math.round(fs * 0.6), lh = Math.round(fs * 1.3);
  const band = pad * 2 + lh * lines.length;
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = h + band;
  const g = cv.getContext('2d');
  g.drawImage(bmp, 0, 0, w, h);
  g.fillStyle = '#111'; g.fillRect(0, h, w, band);
  g.fillStyle = '#fff'; g.textBaseline = 'top';
  lines.forEach((ln, i) => { g.font = `${i === 0 ? 700 : 500} ${fs}px Barlow, "Segoe UI", Roboto, Arial, sans-serif`; g.fillText(ln, pad, h + pad + i * lh, w - pad * 2); });
  const blob = await new Promise((res) => cv.toBlob(res, 'image/jpeg', 0.85));
  return { blob, t: when, lat: ex.lat ?? (fb ? fb.lat : null), lon: ex.lon ?? (fb ? fb.lon : null), gpsSrc: ex.lat != null ? 'foto' : fb ? fb.src : '' };
}
let camKey = null;
const fileSig = (f) => `${f.name}|${f.size}|${f.lastModified}`;

// ---------- Kameradaki fotoğrafları toplu ekleme ----------
// Sahada tabletin kendi kamerasıyla çekilir; sonra hepsi birden seçilir. Her fotoğraf, çekim saatine
// en yakın (bu cihazda, elle verilmiş) işarete eşlenir; kullanıcı listeyi kontrol edip onaylar.
let BATCH = [];
function batchCandidates() {
  const out = [];
  for (const k in M) {
    const m = M[k];
    if (!m || m.m === 'x' || m.bulk || !PARCELS.has(k)) continue;
    if (m.dev && m.dev !== DEVICE) continue;
    out.push({ k, t: m.t });
  }
  return out.sort((a, b) => a.t - b.t);
}
function matchTime(cands, pt) {
  if (!pt) return null;
  let best = null;
  // Önce işaretten sonra çekilmiş olma (20 dk içinde), yoksa işaretten hemen önce çekilmiş olma (5 dk içinde).
  for (const c of cands) if (c.t <= pt + 1000 && pt - c.t <= 20 * 60000) best = c;
  if (best) return best.k;
  for (const c of cands) if (c.t > pt && c.t - pt <= 5 * 60000) return c.k;
  return null;
}
const parcelLabel = (k) => { const { r, f } = PARCELS.get(k); const m = live(k); return `${C.villageOf(r)} ${r.ada}/${r.parsel} · ${f.ad} · ${m ? (m.m === '+' ? '+' : m.m === '-' ? '−' : 'elle') : ''} ${m ? hhmm(m.t) : ''}`; };
function openBatch() { $('batchIn').value = ''; $('batchIn').click(); }
$('batchIn').addEventListener('change', async (e) => {
  const files = [...e.target.files].filter((f) => !f.type || f.type.startsWith('image/') || /\.jpe?g$/i.test(f.name));
  if (!files.length) return;
  const known = new Set();
  PH.forEach((l) => l.forEach((p) => p.sig && known.add(p.sig)));
  const cands = batchCandidates();
  BATCH = [];
  for (const f of files) {
    const ex = C.readExif(await f.slice(0, 256 * 1024).arrayBuffer());
    const t = exifTime(ex.dt) || f.lastModified || 0;
    const dup = known.has(fileSig(f));
    BATCH.push({ f, t, gps: ex.lat != null, k: dup ? '' : (matchTime(cands, t) || ''), dup, url: URL.createObjectURL(f) });
  }
  BATCH.sort((a, b) => a.t - b.t);
  const opts = (sel) => `<option value="">— Ekleme —</option>` + cands.slice().reverse().map((c) => `<option value="${esc(c.k)}" ${c.k === sel ? 'selected' : ''}>${esc(parcelLabel(c.k))}</option>`).join('');
  $('batchList').innerHTML = BATCH.map((b, i) => `<li>
      <img src="${esc(b.url)}" alt="">
      <div class="bmeta"><div><b class="num">${b.t ? trDT(b.t) : 'saat yok'}</b> · ${b.gps ? '<span class="okc">GPS var</span>' : '<span class="noc">GPS yok</span>'}${b.dup ? ' · <span class="noc">zaten eklenmiş</span>' : ''}</div>
        <select data-bi="${i}" aria-label="Fotoğrafın parseli">${opts(b.k)}</select></div>
    </li>`).join('');
  updateBatchCount();
  openDlg('dlgBatch', 'batchGo');
});
function updateBatchCount() {
  const n = BATCH.filter((b) => b.k).length;
  $('batchGo').textContent = `${n} fotoğrafı ekle`;
  $('batchGo').disabled = !n;
  $('batchSum').textContent = `${BATCH.length} fotoğraf seçildi · ${n} tanesi bir parsele eşlendi · ${BATCH.filter((b) => !b.k).length} eşlenmedi`;
}
$('batchList').addEventListener('change', (e) => { const s = e.target.closest('[data-bi]'); if (!s) return; BATCH[+s.dataset.bi].k = s.value; updateBatchCount(); });
$('batchGo').addEventListener('click', async () => {
  $('batchGo').disabled = true;
  let n = 0;
  for (const b of BATCH) {
    if (!b.k) continue;
    $('batchSum').textContent = `Ekleniyor… ${n + 1} / ${BATCH.filter((x) => x.k).length}`;
    try {
      const mk = live(b.k);
      const st = await stampPhoto(b.f, b.k, mk && mk.gps ? { ...mk.gps, src: 'işaret anı' } : null);
      const rec = { id: 'ph' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), key: b.k, t: st.t, lat: st.lat, lon: st.lon, blob: st.blob, asset: null, up: false, sig: fileSig(b.f), addedAt: Date.now() };
      addPhotoRec(rec); await putPhoto(rec); dirtyP.add(phVh(b.k)); n++;
    } catch (err) { console.error(err); }
  }
  BATCH.forEach((b) => URL.revokeObjectURL(b.url));
  BATCH = [];
  closeDlg('dlgBatch');
  schedulePush();
  if (!$('work').hidden) renderDetail();
  const m = $('batchNote'); if (m) { m.hidden = false; m.textContent = `${n} fotoğraf parsellerine eklendi.`; setTimeout(() => { m.hidden = true; }, 6000); }
});
function openCamera(k) { camKey = k; $('camIn').value = ''; $('camIn').click(); }
// Tabletin kendi kamerasıyla (konum etiketi açık) çekilmiş fotoğrafı galeriden/dosyalardan seçmek için.
function openGallery(k) { camKey = k; $('galIn').value = ''; $('galIn').click(); }
async function onPhotoPicked(e) {
  const file = e.target.files[0];
  const k = camKey;
  camKey = null;
  if (!file || !k) return;
  if (file.type && !file.type.startsWith('image/') && !/\.jpe?g$/i.test(file.name)) {
    const rw = [...document.querySelectorAll('[data-k]')].find((x) => x.dataset.k === k);
    const nt = rw && rw.querySelector('.phnote');
    if (nt) nt.textContent = 'Seçilen dosya fotoğraf değil. DCIM → Camera klasöründen bir JPG seçin.';
    return;
  }
  const row = [...document.querySelectorAll('[data-k]')].find((x) => x.dataset.k === k);
  const note = row && row.querySelector('.phnote');
  if (note) note.textContent = 'Fotoğraf hazırlanıyor…';
  try {
    // Uygulamanın kamerasıyla çekildiyse tabletin şu anki konumu, galeriden seçildiyse işaretin verildiği anın konumu.
    let fb = null;
    if (e.target.id === 'camIn') { const p = await getPos(60000); if (p) fb = { lat: p.lat, lon: p.lon, acc: p.acc, src: 'cihaz' }; }
    if (!fb) { const m = live(k); if (m && m.gps) fb = { ...m.gps, src: 'işaret anı' }; }
    const st = await stampPhoto(file, k, fb);
    const rec = { id: 'ph' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), key: k, t: st.t, lat: st.lat, lon: st.lon, gpsSrc: st.gpsSrc, blob: st.blob, asset: null, up: false, sig: fileSig(file), addedAt: Date.now() };
    addPhotoRec(rec);
    await putPhoto(rec);
    dirtyP.add(phVh(k));
    updateRow(k);
    schedulePush();
    const r2 = [...document.querySelectorAll('[data-k]')].find((x) => x.dataset.k === k);
    const n2 = r2 && r2.querySelector('.phnote');
    if (n2) {
      if (rec.lat != null) n2.textContent = rec.gpsSrc === 'foto' ? `Konum fotoğraftan alındı: ${fmtCoord(rec.lat, rec.lon)}` : `Konum ${rec.gpsSrc === 'cihaz' ? 'tabletin GPS’inden' : 'işaretin verildiği andan'} alındı: ${fmtCoord(rec.lat, rec.lon)}`;
      else n2.innerHTML = 'Fotoğrafta konum bilgisi yok. Tabletin kendi kamera uygulamasıyla (konum etiketi açık) çekip buradan seçmeyi deneyin: <button class="lnk" type="button" data-gal>Galeriden / Dosyalardan seç</button>';
    }
  } catch (err) {
    console.error(err);
    if (note) note.textContent = 'Fotoğraf eklenemedi: ' + (err && err.message ? err.message : err);
  }
}
$('camIn').addEventListener('change', onPhotoPicked);
$('galIn').addEventListener('change', onPhotoPicked);
let viewPh = null;
function openPhoto(id) {
  const ph = photoById(id);
  if (!ph) return;
  viewPh = ph;
  const { r, f } = PARCELS.get(ph.key) || { r: {}, f: {} };
  $('phImg').src = phSrc(ph);
  $('phInfo').textContent = `${f.ad || ''} · Ada ${r.ada} / Parsel ${r.parsel} · ${trDT(ph.t)} · ${ph.lat != null ? 'GPS ' + fmtCoord(ph.lat, ph.lon) : 'konum yok'}${ph.up ? '' : ' · henüz buluta yüklenmedi'}`;
  $('phMsg').textContent = '';
  openDlg('dlgPhoto', 'phClose');
}
async function photoBytes(ph) {
  if (ph.blob) return new Uint8Array(await ph.blob.arrayBuffer());
  const res = await fetch(phSrc(ph));
  if (!res.ok) throw new Error('fotoğraf indirilemedi');
  return new Uint8Array(await res.arrayBuffer());
}

// ---------- Durum ----------
let DOCS = [];        // [{ id, name, year, farmers }]
let M = {};           // parsel anahtarı -> { m: '+'|'-'|'e'|'x', v?, d, t }
const dirty = new Set();  // buluta yazılmamış köyler (köy anahtarının özeti)
let mode = 'koy';
let selKoy = null;
try { const v = JSON.parse(localStorage.getItem('tespit.view') || '{}'); if (['koy', 'ciftci', 'tbs'].includes(v.mode)) mode = v.mode; selKoy = v.koy || null; } catch (e) {}
const saveView = () => { try { localStorage.setItem('tespit.view', JSON.stringify({ mode, koy: selKoy, f: selF })); } catch (e) {} };
let selF = (() => { try { return JSON.parse(localStorage.getItem('tespit.view') || '{}').f || null; } catch (e) { return null; } })();      // TC
let query = '';
let koyF = '';
let showCayir = false;

let PARCELS = new Map();  // anahtar -> { r, f }
let VILLAGES = [];        // [{ name, keys }]
let FARMERS = [];         // [{ id: TC, f, keys, koy }]
let PLACES = new Map();   // köy|ada|parsel -> parsel anahtarları (ortak kullanılan parseller için)
let olcK = null;          // ölçüm girişi açık olan parsel
let OVER = new Map();     // köy|ada|parsel -> { palan, sum } beyan toplamı parsel alanını aşıyorsa
let showOver = false;

const byPlace = (a, b) => C.villageOf(a).localeCompare(C.villageOf(b), 'tr') || num(a.ada) - num(b.ada) || num(a.parsel) - num(b.parsel);
function rebuild() {
  PARCELS = new Map();
  const vmap = new Map();
  const fmap = new Map();
  for (const d of DOCS) {
    C.assignKeys(d.farmers);
    for (const f of d.farmers) {
      if (!fmap.has(f.tc)) fmap.set(f.tc, { id: f.tc, f, keys: [] });
      const x = fmap.get(f.tc);
      for (const r of f.rows) {
        if (PARCELS.has(r.key)) continue;
        PARCELS.set(r.key, { r, f });
        x.keys.push(r.key);
        const v = C.villageOf(r);
        if (!vmap.has(v)) vmap.set(v, []);
        vmap.get(v).push(r.key);
      }
    }
  }
  FARMERS = [...fmap.values()].filter((x) => x.keys.length);
  PLACES = new Map();
  PARCELS.forEach((p, k) => { const pk = C.placeKey(p.r); if (!PLACES.has(pk)) PLACES.set(pk, []); PLACES.get(pk).push(k); });
  OVER = new Map();
  PLACES.forEach((keys, pk) => {
    let palan = 0, sum = 0;
    for (const k of keys) { const r = PARCELS.get(k).r; palan = Math.max(palan, parseFloat(r.palan) || 0); sum += parseFloat(r.alan) || 0; }
    if (palan > 0 && sum > palan + 0.0005) OVER.set(pk, { palan, sum });
  });
  for (const x of FARMERS) {
    x.keys.sort((a, b) => byPlace(PARCELS.get(a).r, PARCELS.get(b).r));
    const c = {};
    x.keys.forEach((k) => { const v = C.villageOf(PARCELS.get(k).r); c[v] = (c[v] || 0) + 1; });
    x.koy = Object.entries(c).sort((a, b) => b[1] - a[1])[0][0];
  }
  FARMERS.sort((a, b) => a.f.ad.localeCompare(b.f.ad, 'tr'));
  VILLAGES = [...vmap.entries()].map(([name, keys]) => ({ name, keys })).sort((a, b) => a.name.localeCompare(b.name, 'tr'));
  for (const v of VILLAGES) v.keys.sort((a, b) => {
    const A = PARCELS.get(a), B = PARCELS.get(b);
    return num(A.r.ada) - num(B.r.ada) || num(A.r.parsel) - num(B.r.parsel) || A.f.ad.localeCompare(B.f.ad, 'tr');
  });
  if (!selKoy || !VILLAGES.find((v) => v.name === selKoy)) {
    const best = VILLAGES.slice().sort((a, b) => checkCount(b.keys) - checkCount(a.keys))[0];
    selKoy = best ? best.name : null;
  }
  if (!selF || !FARMERS.find((x) => x.id === selF)) selF = FARMERS[0] ? FARMERS[0].id : null;
  const prods = new Set([S.eksiUrun]);
  PARCELS.forEach((p) => prods.add(p.r.urun));
  $('urunList').innerHTML = [...prods].map((u) => `<option value="${esc(u)}"></option>`).join('');
}
const isCheck = (k) => !C.isCayir(PARCELS.get(k).r);
const checkCount = (keys) => keys.filter(isCheck).length;
// M'de silinen işaret { m: 'x', t } olarak kalır; böylece cihazlar arası birleştirmede geri gelmez.
const live = (k) => { const m = M[k]; return m && m.m !== 'x' ? m : null; };
const eff = (k) => live(k) || (C.isCayir(PARCELS.get(k).r) ? C.AUTO : null);
const markOfRow = (r) => live(r.key) || (C.isCayir(r) ? C.AUTO : null);
// Aynı köy/ada/parseli başka çiftçiler de beyan ettiyse onların satırları (ortak kullanım).
const sharesOf = (k) => { const p = PARCELS.get(k); return (PLACES.get(C.placeKey(p.r)) || []).filter((x) => x !== k && PARCELS.get(x).f.tc !== p.f.tc); };
const DEVICE = (() => { try { let d = localStorage.getItem('tespit.dev'); if (!d) { d = 'd' + Math.random().toString(36).slice(2, 10); localStorage.setItem('tespit.dev', d); } return d; } catch (e) { return 'd-anon'; } })();
// ---------- Cihaz GPS'i (bağımsız sürüm) ----------
// Claude içinde konum izni yok; uygulama kendi adresinden (https) açıldığında tabletin GPS'i okunur.
const STANDALONE = !(window.claude && typeof window.claude.use === 'function');
let gpsState = 'bekleniyor', lastPos = null;
function getPos(maxAge) {
  maxAge = maxAge || 30000;
  return new Promise((res) => {
    if (!STANDALONE || !('geolocation' in navigator)) { if (STANDALONE) { gpsState = 'yok'; renderGpsState(); } return res(null); }
    if (lastPos && Date.now() - lastPos.at < maxAge) return res(lastPos);
    navigator.geolocation.getCurrentPosition(
      (p) => { lastPos = { lat: p.coords.latitude, lon: p.coords.longitude, acc: Math.round(p.coords.accuracy || 0), at: Date.now() }; gpsState = 'tamam'; renderGpsState(); res(lastPos); },
      (err) => { gpsState = err && err.code === 1 ? 'izinyok' : 'alinamadi'; renderGpsState(); res(null); },
      { enableHighAccuracy: true, timeout: 20000, maximumAge: maxAge });
  });
}
// İşaret verildiği anın konumunu işarete ekler (işaret bu arada değişmediyse).
function attachGps(keys) {
  if (!STANDALONE) return;
  const t0 = keys.map((k) => M[k] && M[k].t);
  getPos().then((pos) => {
    if (!pos) return;
    keys.forEach((k, i) => {
      const m = M[k];
      if (!m || m.m === 'x' || m.t !== t0[i]) return;
      M[k] = { ...m, gps: { lat: +pos.lat.toFixed(6), lon: +pos.lon.toFixed(6), acc: pos.acc } };
      dirty.add(vhOf(k));
      const inDom = mode !== 'tbs' && $('rows') && [...$('rows').querySelectorAll('[data-k]')].some((x) => x.dataset.k === k);
      if (inDom) updateRow(k);
    });
    queueSave();
  });
}
function renderGpsState() {
  const el = document.getElementById('gpsState');
  if (!el) return;
  el.hidden = !STANDALONE;
  const map = {
    bekleniyor: 'GPS: ilk işarette konum izni istenecek',
    tamam: `GPS açık${lastPos ? ` (±${lastPos.acc} m)` : ''}`,
    izinyok: 'GPS izni verilmedi: tarayıcı ayarlarından bu uygulamaya konum izni verin',
    alinamadi: 'GPS alınamadı: tabletin konumu açık mı?',
    yok: 'Bu cihazda GPS yok'
  };
  el.textContent = map[gpsState] || '';
  el.className = 'foot gpsstate ' + (gpsState === 'tamam' ? 'ok' : gpsState === 'bekleniyor' ? '' : 'ed');
}

function setMark(k, m) { const o = { ...m, d: m.d || C.isoToday(), t: Date.now(), by: S.p1 || undefined, dev: DEVICE, res: undefined }; for (const x in o) if (o[x] === undefined) delete o[x]; M[k] = o; dirty.add(vhOf(k)); }
function clearMark(k) { if (M[k]) { M[k] = { m: 'x', t: Date.now() }; dirty.add(vhOf(k)); } }

function keyStats(keys) {
  const s = { total: keys.length, check: 0, checked: 0, plus: 0, minus: 0, ed: 0, auto: 0, done: 0 };
  for (const k of keys) {
    const m = eff(k);
    if (isCheck(k)) { s.check++; if (live(k)) s.checked++; }
    if (!m) continue;
    s.done++;
    if (m.auto) s.auto++; else if (m.m === '+') s.plus++; else if (m.m === '-') s.minus++; else s.ed++;
  }
  return s;
}
const farmerDone = (x) => x.keys.every((k) => eff(k));
function totals() {
  const s = keyStats([...PARCELS.keys()]);
  return { ...s, farmers: FARMERS.length, fdone: FARMERS.filter(farmerDone).length };
}

// ---------- Excel okuma ----------
async function readXlsx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = async (path) => { const f = zip.file(path); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
  // Biçimli metin parçalarını (<r><t>) birleştirir, okunuş bilgisini (<rPh>) atlar.
  const richText = (el) => {
    let t = '';
    for (const n of el.childNodes) {
      if (n.nodeName === 't') t += n.textContent;
      else if (n.nodeName === 'r') for (const tt of n.getElementsByTagName('t')) t += tt.textContent;
    }
    return t;
  };
  const strings = [];
  const ss = await xml('xl/sharedStrings.xml');
  if (ss) for (const si of ss.getElementsByTagName('si')) strings.push(richText(si));
  let sheetPath = 'xl/worksheets/sheet1.xml';
  try {
    const wb = await xml('xl/workbook.xml');
    const rels = await xml('xl/_rels/workbook.xml.rels');
    const first = wb.getElementsByTagName('sheet')[0];
    const rid = first.getAttribute('r:id') || first.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    for (const r of rels.getElementsByTagName('Relationship')) if (r.getAttribute('Id') === rid) {
      const tg = r.getAttribute('Target');
      sheetPath = tg.startsWith('/') ? tg.slice(1) : 'xl/' + tg.replace(/^\.\//, '');
    }
  } catch (e) {}
  const sh = await xml(sheetPath);
  if (!sh) throw new Error('çalışma sayfası bulunamadı');
  const colIdx = (ref) => { let n = 0; for (const ch of ref) { const c = ch.charCodeAt(0); if (c < 65 || c > 90) break; n = n * 26 + (c - 64); } return n - 1; };
  const rows = [];
  for (const row of sh.getElementsByTagName('row')) {
    const arr = [];
    for (const c of row.getElementsByTagName('c')) {
      const t = c.getAttribute('t');
      const v = c.getElementsByTagName('v')[0];
      let val = '';
      if (t === 's') val = v ? strings[+v.textContent] ?? '' : '';
      else if (t === 'inlineStr') { const is = c.getElementsByTagName('is')[0]; val = is ? richText(is) : ''; }
      else val = v ? v.textContent : '';
      arr[colIdx(c.getAttribute('r') || 'A')] = val;
    }
    rows[(+row.getAttribute('r') || rows.length + 1) - 1] = arr;
  }
  return rows;
}


function mergeMarks(incoming) {
  let n = 0;
  for (const k in incoming) {
    if (!k.startsWith('k3|')) continue;
    const a = incoming[k];
    const b = M[k];
    if (!b || (a.t || 0) > (b.t || 0)) { M[k] = a; n++; dirty.add(vhOf(k)); }
  }
  return n;
}
async function addFiles(files) {
  const st = $('loadStatus');
  st.className = 'status';
  let added = 0, restored = 0;
  const errors = [];
  const diffs = [];
  for (const file of files) {
    try {
      st.textContent = `${file.name} okunuyor…`;
      const head = new Uint8Array(await file.slice(0, 5).arrayBuffer());
      if (head[0] === 0x50 && head[1] === 0x4b) {
        const rows = await readXlsx(await file.arrayBuffer());
        const { year, farmers } = C.fromSheet(rows);
        if (!farmers.length) { errors.push(`${file.name}: ÇKS parsel üretim listesi bulunamadı (başlık satırında “İşletme Adı”, “TC / Vergi No”, “Köy”, “Ada No”, “Parsel No”, “Ürün” olmalı)`); continue; }
        C.assignKeys(farmers);
        const sig = farmers.map((f) => f.tc + ':' + f.rows.map((r) => r.key).join(',')).join('|');
        let h = 2166136261 >>> 0;
        for (let i = 0; i < sig.length; i++) { h ^= sig.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
        const id = 'x' + h.toString(36) + '-' + sig.length.toString(36);
        const doc = { id, name: file.name.replace(/\.xlsx?$/i, ''), year, farmers };
        const at = DOCS.findIndex((d) => d.id === id);
        if (at >= 0) { DOCS[at] = doc; await putDoc(doc); added++; diffs.push({ name: doc.name, same: true }); continue; }
        // Aynı yılın, aynı köyleri kapsayan eski listesi varsa yeni liste onun yerine geçer.
        const vset = (fs) => new Set(fs.flatMap((f) => f.rows.map(C.villageOf)));
        const nv = vset(farmers);
        const olds = DOCS.filter((d) => {
          if ((d.year || '') !== (year || '')) return false;
          const ov = vset(d.farmers);
          let inter = 0; nv.forEach((x) => { if (ov.has(x)) inter++; });
          return inter > 0 && inter >= Math.min(nv.size, ov.size) * 0.5;
        });
        if (olds.length) diffs.push({ name: doc.name, ...diffLists(olds, farmers) });
        for (const o of olds) await delDocDb(o.id);
        DOCS = DOCS.filter((d) => !olds.includes(d));
        DOCS.push(doc);
        await putDoc(doc);
        added++;
      } else if (String.fromCharCode(...head).startsWith('%PDF')) {
        const b = await C.readBackup(PDFLib, new Uint8Array(await file.arrayBuffer()));
        if (!b || !b.marks) { errors.push(`${file.name}: bu ekrandan indirilmiş bir tutanak değil, içinde işaret yok`); continue; }
        restored += mergeMarks(b.marks);
        const ps = b.settings || {};
        for (const k in C.DEFAULTS) if (ps[k] !== undefined && (S[k] === C.DEFAULTS[k] || S[k] === '')) S[k] = ps[k];
        saveSettings();
      } else {
        errors.push(`${file.name}: Excel (.xlsx) ya da bu ekrandan indirilmiş PDF seçin`);
      }
    } catch (e) {
      console.error(e);
      errors.push(`${file.name}: açılamadı (${e && e.message ? e.message : e})`);
    }
  }
  queueSave();
  rebuild();
  if (added) resolvePending();
  if (cloud) { if (added) await cloudPull(); else schedulePush(); }
  const msgs = errors.slice();
  if (restored) msgs.unshift(`${restored} parselin işareti geri yüklendi.`);
  st.className = errors.length ? 'status err' : 'status';
  st.textContent = msgs.join(' · ');
  if (added && DOCS.length) showWork();
  else renderHome();
  if (diffs.some((d) => !d.same)) showDiff(diffs.filter((d) => !d.same));
}

async function loadAll() {
  const docs = await allDocs();
  const out = [];
  let legacy = false;
  for (const d of docs.sort((a, b) => (a.updated || 0) - (b.updated || 0))) {
    if (d.kind !== 'liste') { await delDocDb(d.id); legacy = true; continue; } // önceki sürümün TBS PDF kayıtları artık kullanılmıyor
    out.push({ id: d.id, name: d.name, year: d.year || '', farmers: d.farmers });
  }
  DOCS = out;
  const saved = await getMarks();
  M = {};
  if (legacy) { queueSave(); return; } // eski PDF denemelerinin işaretleriyle başlamasın
  mergeMarks((saved && saved.data) || {});
  dirty.clear();
}

// Eski ve yeni ÇKS listesini karşılaştırır.
function diffLists(oldDocs, newFarmers) {
  const om = new Map(), nm = new Map();
  oldDocs.forEach((d) => { C.assignKeys(d.farmers); d.farmers.forEach((f) => f.rows.forEach((r) => om.set(r.key, { r, f }))); });
  newFarmers.forEach((f) => f.rows.forEach((r) => nm.set(r.key, { r, f })));
  const base = (k) => k.split('|').slice(0, 5).join('|');
  let added = [...nm.keys()].filter((k) => !om.has(k));
  let removed = [...om.keys()].filter((k) => !nm.has(k));
  const changed = [];
  const addByBase = new Map();
  added.forEach((k) => { const b = base(k); if (!addByBase.has(b)) addByBase.set(b, []); addByBase.get(b).push(k); });
  const pairedAdd = new Set(), pairedRem = new Set();
  removed.forEach((rk) => { const l = addByBase.get(base(rk)); if (l && l.length) { const ak = l.shift(); changed.push({ o: rk, n: ak }); pairedAdd.add(ak); pairedRem.add(rk); } });
  added = added.filter((k) => !pairedAdd.has(k));
  removed = removed.filter((k) => !pairedRem.has(k));
  const alan = [...nm.keys()].filter((k) => om.has(k) && om.get(k).r.alan !== nm.get(k).r.alan);
  const ot = new Set([...om.values()].map((x) => x.f.tc)), nt = new Set([...nm.values()].map((x) => x.f.tc));
  const marked = (k) => { const m = live(k); return m && !m.auto ? m : null; };
  const affected = [
    ...removed.filter(marked).map((k) => ({ k, why: 'beyan listeden kalktı', o: om.get(k) })),
    ...changed.filter((c) => marked(c.o)).map((c) => ({ k: c.o, why: `ürün değişti: ${om.get(c.o).r.urun} → ${nm.get(c.n).r.urun}`, o: om.get(c.o) })),
    ...alan.filter(marked).map((k) => ({ k, why: `alan değişti: ${om.get(k).r.alan} → ${nm.get(k).r.alan} da`, o: om.get(k) }))
  ];
  return {
    oldN: om.size, newN: nm.size, newFarmers: [...nt].filter((t) => !ot.has(t)).length, goneFarmers: [...ot].filter((t) => !nt.has(t)).length,
    added: added.map((k) => nm.get(k)), removed: removed.map((k) => om.get(k)), changed: changed.map((c) => ({ o: om.get(c.o), n: nm.get(c.n) })),
    alan: alan.map((k) => ({ o: om.get(k), n: nm.get(k) })), affected
  };
}
function showDiff(diffs) {
  const line = (x) => `${esc(C.villageOf(x.r))} ${esc(x.r.ada)}/${esc(x.r.parsel)} · ${esc(x.f.ad)}`;
  const sec = (title, arr, fn, cls) => arr.length ? `<h4 class="${cls || ''}">${title} (${arr.length})</h4><ul class="difflist">${arr.slice(0, 300).map((x) => `<li>${fn(x)}</li>`).join('')}${arr.length > 300 ? `<li>… ve ${arr.length - 300} tane daha</li>` : ''}</ul>` : '';
  $('diffBody').innerHTML = diffs.map((d) => `
    <p class="num"><b>${esc(d.name)}</b>: önceki listede ${d.oldN}, yenisinde ${d.newN} parsel. ${d.newFarmers} yeni çiftçi, ${d.goneFarmers} çiftçi listeden çıkmış.</p>
    ${d.affected.length ? `<div class="warn">İşaretlediğiniz ${d.affected.length} parselin beyanı değişmiş. Bunları yeniden kontrol edin; ürünü değişenler kontrol listesine kendiliğinden geri döndü.</div>` : '<div class="okmsg">İşaretlediğiniz parsellerin hiçbirinin beyanı değişmemiş.</div>'}
    ${sec('Yeniden kontrol edilecek işaretli parseller', d.affected, (a) => `${line(a.o)} — ${esc(a.why)}`, 'no')}
    ${sec('Ürünü değişen', d.changed, (c) => `${line(c.n)} — ${esc(c.o.r.urun)} → <b>${esc(c.n.r.urun)}</b>`)}
    ${sec('Yeni parsel', d.added, (x) => `${line(x)} — ${esc(x.r.urun)} ${esc(x.r.alan)} da`)}
    ${sec('Listeden kalkan', d.removed, (x) => `${line(x)} — ${esc(x.r.urun)} ${esc(x.r.alan)} da`)}
    ${sec('Alanı değişen', d.alan, (c) => `${line(c.n)} — ${esc(c.o.r.alan)} → <b>${esc(c.n.r.alan)}</b> da`)}
  `).join('<hr>');
  openDlg('dlgDiff', 'diffClose');
}

// ---------- Dosyalar ekranı ----------
function renderHome() {
  $('home').hidden = false; $('work').hidden = true;
  $('prog').hidden = true; $('outBtn').hidden = true; $('filesBtn').hidden = true; $('modeSeg').hidden = true;
  $('saveBtn').hidden = true; $('saveState').hidden = true;
  if ($('syncBar')) $('syncBar').hidden = true;
  $('backBtn').style.visibility = 'hidden';
  $('toWork').hidden = !DOCS.length;
  $('savedCard').hidden = !DOCS.length;
  $('syncCard').hidden = !DOCS.length;
  renderConflictNote();
  renderSyncLog();
  $('syncPhN').textContent = String([...PH.values()].reduce((a, l) => a + l.length, 0));
  $('savedList').innerHTML = DOCS.map((d) => {
    const keys = d.farmers.flatMap((f) => f.rows.map((r) => r.key)).filter((k) => PARCELS.has(k));
    const s = keyStats(keys);
    const vs = new Set(d.farmers.flatMap((f) => f.rows.map(C.villageOf)));
    return `<li><div class="meta"><b>${esc(d.name)}</b><span class="num">ÇKS listesi${d.year ? ' · ' + esc(d.year) : ''} · ${d.farmers.length} çiftçi · ${keys.length} parsel · ${vs.size} köy · kontrol ${s.checked}/${s.check}</span></div>
      <button class="btn danger" type="button" data-del="${esc(d.id)}" aria-label="${esc(d.name)} dosyasını kaldır">Kaldır</button></li>`;
  }).join('');
}
$('savedList').addEventListener('click', (e) => {
  const d = e.target.closest('[data-del]');
  if (!d) return;
  const doc = DOCS.find((x) => x.id === d.dataset.del);
  ask('Dosya kaldırılsın mı?', `${doc ? doc.name : 'Dosya'} bu cihazdan kaldırılır. Yalnız bu dosyadaki parsellerin işaretleri de silinir. Daha önce indirdiğiniz tutanak PDF'ini yeniden yüklerseniz işaretler geri gelir.`, 'Kaldır', async () => {
    await delDocDb(d.dataset.del);
    DOCS = DOCS.filter((x) => x.id !== d.dataset.del);
    rebuild();
    for (const k in M) if (!PARCELS.has(k)) delete M[k];
    queueSave();
    renderHome();
  });
});
$('fileIn').addEventListener('change', (e) => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) addFiles(fs); });
$('toWork').addEventListener('click', () => showWork());
$('filesBtn').addEventListener('click', () => { $('loadStatus').textContent = ''; renderHome(); });

// ---------- Çalışma ekranı ----------
const narrow = () => window.matchMedia('(max-width: 900px)').matches;
function showWork() {
  $('home').hidden = true; $('work').hidden = false;
  $('prog').hidden = false; $('outBtn').hidden = false; $('filesBtn').hidden = false; $('modeSeg').hidden = false;
  $('saveBtn').hidden = false; $('saveState').hidden = false;
  setView(narrow() ? 'list' : 'detail');
  renderAll();
}
function renderAll() { renderMode(); renderSide(); renderDetail(); renderProg(); renderSaveState(); renderSyncBar(); }
function setView(v) {
  $('work').classList.toggle('show-list', v === 'list');
  $('work').classList.toggle('show-detail', v === 'detail');
  $('backBtn').style.visibility = v === 'detail' ? 'visible' : 'hidden';
}
$('backBtn').addEventListener('click', () => setView('list'));
function renderMode() {
  $('modeSeg').querySelectorAll('[data-mode]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.mode === mode)));
  $('q').placeholder = mode === 'ciftci' ? 'Ad, TC ya da ada/parsel' : 'Ada/parsel ya da çiftçi adı';
}
$('modeSeg').addEventListener('click', (e) => {
  const b = e.target.closest('[data-mode]');
  if (!b || b.dataset.mode === mode) return;
  mode = b.dataset.mode;
  query = ''; $('q').value = '';
  setView(narrow() ? 'list' : 'detail');
  renderAll();
});
function renderProg() {
  const t = totals();
  $('progTxt').textContent = `Kontrol ${t.checked}/${t.check} · Tamam ${t.fdone}/${t.farmers} çiftçi`;
  $('progBar').style.width = (t.check ? (100 * t.checked) / t.check : 100) + '%';
}
// Arama
const apQuery = () => query.match(/^(\d+)\s*[\/ ]\s*(\d+)$/);
function parcelMatches(k) {
  if (!query) return true;
  const p = PARCELS.get(k);
  const ap = apQuery();
  if (ap) return p.r.ada === ap[1] && p.r.parsel === ap[2];
  if (/^\d+$/.test(query)) return p.r.ada === query || p.r.parsel === query || p.f.tc.includes(query);
  return upper(p.f.ad).includes(upper(query));
}
function farmerMatches(x) {
  if (koyF && x.koy !== koyF) return false;
  if (!query) return true;
  if (apQuery()) return x.keys.some(parcelMatches);
  if (/^\d+$/.test(query)) return x.f.tc.includes(query) || x.keys.some(parcelMatches);
  return upper(x.f.ad).includes(upper(query));
}
$('q').addEventListener('input', (e) => { query = e.target.value.trim(); renderSide(); renderDetail(); });

// Yan liste
const LIST_MAX = 300;
function renderSide() {
  if (mode === 'tbs') {
    $('chips').hidden = true;
    const all = tbsKeys([...PARCELS.keys()]);
    const items = VILLAGES.map((v) => ({ v, keys: tbsKeys(v.keys) })).filter((x) => x.keys.length);
    const it = (name, label, keys) => {
      const done = keys.filter((k) => live(k).tbs).length;
      const cls = keys.length && done === keys.length ? 'done' : done ? 'part' : '';
      return `<li><button type="button" data-tv="${esc(name)}" aria-current="${tbsKoy === name}"><span class="dot ${cls}"></span><span><b>${esc(label)}</b><small>${keys.length} değişiklik</small></span><span class="cnt num">${done}/${keys.length}<em style="color:var(--muted);font-weight:500">işlendi</em></span></button></li>`;
    };
    $('flist').innerHTML = `<li class="subhead">TBS'ye girilecek değişiklikler</li>` + it('', 'Bütün köyler', all) + items.map((x) => it(x.v.name, x.v.name, x.keys)).join('') +
      (all.length ? '' : '<li class="empty">Henüz değişiklik yok. − verilen, ölçülen ya da elle düzenlenen parseller burada listelenir.</li>');
    return;
  }
  if (mode === 'koy') {
    $('chips').hidden = true;
    const vs = VILLAGES.filter((v) => !query || v.keys.some(parcelMatches));
    $('flist').innerHTML = `<li class="subhead">Parselin bulunduğu köy</li>` + (vs.length ? vs.map(villageItem).join('') : '<li class="empty">Eşleşen parsel yok.</li>');
  } else {
    const counts = {};
    FARMERS.forEach((x) => { counts[x.koy] = (counts[x.koy] || 0) + 1; });
    const ks = Object.keys(counts).sort((a, b) => a.localeCompare(b, 'tr'));
    $('chips').hidden = ks.length < 2;
    $('chips').innerHTML = [`<button class="chip" type="button" data-koy="" aria-pressed="${koyF === ''}">Tümü ${FARMERS.length}</button>`]
      .concat(ks.map((k) => `<button class="chip" type="button" data-koy="${esc(k)}" aria-pressed="${koyF === k}">${esc(k)} ${counts[k]}</button>`)).join('');
    const fs = FARMERS.filter(farmerMatches);
    $('flist').innerHTML = fs.length
      ? fs.slice(0, LIST_MAX).map(farmerItem).join('') + (fs.length > LIST_MAX ? `<li class="empty">${fs.length - LIST_MAX} çiftçi daha var; köy seçin ya da arayın.</li>` : '')
      : '<li class="empty">Eşleşen çiftçi yok.</li>';
  }
}
function villageItem(v) {
  const s = keyStats(v.keys);
  const cls = s.check === 0 || s.checked === s.check ? 'done' : s.checked ? 'part' : '';
  return `<li><button type="button" data-v="${esc(v.name)}" aria-current="${v.name === selKoy}"><span class="dot ${cls}"></span>
    <span><b>${esc(v.name)}</b><small>${s.total} parsel · ${s.total - s.check} çayır otu</small></span>
    <span class="cnt num">${s.checked}/${s.check}<em style="color:var(--muted);font-weight:500">kontrol</em></span></button></li>`;
}
function farmerItem(x) {
  const s = keyStats(x.keys);
  const cls = s.done === s.total ? 'done' : s.checked ? 'part' : '';
  return `<li><button type="button" data-f="${x.id}" aria-current="${x.id === selF}"><span class="dot ${cls}"></span>
    <span><b>${esc(x.f.ad)}</b><small>${esc(x.koy)} · ${s.total} parsel · ${s.check} kontrol</small></span>
    <span class="cnt num">${s.done}/${s.total}${s.minus ? `<em>${s.minus} −</em>` : ''}</span></button></li>`;
}
$('chips').addEventListener('click', (e) => { const b = e.target.closest('[data-koy]'); if (!b) return; koyF = b.dataset.koy; renderSide(); });
$('flist').addEventListener('click', (e) => {
  const tv = e.target.closest('[data-tv]');
  if (tv) { tbsKoy = tv.dataset.tv; renderSide(); renderDetail(); setView('detail'); $('main').scrollTop = 0; return; }
  const v = e.target.closest('[data-v]');
  const f = e.target.closest('[data-f]');
  if (v) selKoy = v.dataset.v;
  else if (f) selF = f.dataset.f;
  else return;
  renderSide(); renderDetail();
  setView('detail');
  $('main').scrollTop = 0;
});

// Parsel kartı
function tspText(t) { return [esc(t.urun), t.alan ? esc(t.alan) + ' da' + (t.olc ? ' <b>(ölçüm)</b>' : '') : '', esc(t.acik)].filter(Boolean).join(' · '); }
function rowHtml(k, ctx) {
  const { r, f } = PARCELS.get(k);
  const m = eff(k);
  const t = C.tespit(r, m, S);
  const cls = !m ? '' : m.auto ? 'm-plus auto' : m.m === '+' ? 'm-plus' : m.m === '-' ? 'm-minus' : 'm-edit';
  const ap = apQuery();
  const hit = ap && r.ada === ap[1] && r.parsel === ap[2];
  const who = ctx === 'koy' ? `<div class="who">${esc(f.ad)} <span>· TC ${esc(f.tc)}</span></div>` : '';
  const where = ctx === 'ciftci' ? `<span class="where">${esc(C.villageOf(r))}</span>` : '';
  const sh = sharesOf(k);
  const shared = sh.length ? `<div class="shared"><b>Ortak parsel${r.palan ? ` · parsel ${esc(r.palan)} da` : ''}:</b> ${sh.map((x) => { const q = PARCELS.get(x); return `${esc(q.f.ad)} (${esc(q.r.alan)} da, ${esc(q.r.urun.split('\\')[0])})`; }).join(', ')}. + ve − bu çiftçilere de işlenir.</div>` : '';
  const olcOpen = olcK === k;
  const cur = live(k);
  const phs = PH.get(k) || [];
  const ov = OVER.get(C.placeKey(r));
  const overHtml = ov ? `<div class="overflag">Alan aşımı: parsel ${ov.palan.toFixed(3)} da, beyanların toplamı ${ov.sum.toFixed(3)} da (+${(ov.sum - ov.palan).toFixed(3)})</div>` : '';
  return `<div class="card prow koy ${cls}" data-k="${esc(k)}" ${hit ? 'style="outline:3px solid var(--accent)"' : ''}>
    <div class="info">
      <div class="ap num">${esc(r.ada)} / ${esc(r.parsel)} <span>· ${esc(r.alan)} da</span>${where}</div>
      ${who}
      <div class="urun">Beyan: <b>${esc(r.urun)}</b></div>
      ${shared}
      ${overHtml}
    </div>
    <div class="marks">
      <button class="mk plus" type="button" data-mk="+" aria-pressed="${m ? m.m === '+' : false}" aria-label="${esc(r.ada)}/${esc(r.parsel)}: beyan doğru"><span class="s">+</span><span class="t">Uygun</span></button>
      <button class="mk minus" type="button" data-mk="-" aria-pressed="${m ? m.m === '-' : false}" aria-label="${esc(r.ada)}/${esc(r.parsel)}: uygun değil, çayır otu yaz"><span class="s">−</span><span class="t">Çayır otu</span></button>
      <button class="mk cam" type="button" data-cam aria-label="${esc(r.ada)}/${esc(r.parsel)}: fotoğraf çek"><span class="s r"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>${phs.length ? `<em class="phn">${phs.length}</em>` : ''}</span><span class="t">Fotoğraf</span></button>
      <button class="mk olc" type="button" data-olcbtn aria-pressed="${cur && cur.olc ? 'true' : 'false'}" aria-expanded="${olcOpen}" aria-label="${esc(r.ada)}/${esc(r.parsel)}: ölçülen alanı gir"><span class="s r">${cur && cur.olc ? esc(cur.olc) : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 17l14-14 4 4L7 21z"/><path d="M7 13l2 2M10 10l2 2M13 7l2 2"/></svg>'}</span><span class="t">${cur && cur.olc ? 'da ölçüm' : 'Ölçüm'}</span></button>
    </div>
    ${olcOpen ? `<form class="olcform" data-olcform>
      <label for="olcIn">Ölçülen alan (da)</label>
      <input id="olcIn" inputmode="decimal" autocomplete="off" value="${esc((cur && cur.olc) || '')}" placeholder="ör. 10,5">
      <button class="btn primary" type="submit">Kaydet</button>
      ${cur && cur.olc ? '<button class="btn danger" type="button" data-olcdel>Ölçümü sil</button>' : ''}
      <button class="btn" type="button" data-olcx>Vazgeç</button>
      <span class="err" id="olcErr" role="alert"></span>
    </form>` : ''}
    <div class="tsp">${m && m.auto ? '<span class="auto-tag">Çayır otu beyanı · otomatik +</span>' : t ? `<span class="val">Tespit: ${tspText(t)}</span>${m.by ? `<span class="by">${esc(m.by)} · ${esc(C.trDate(m.d || ''))}</span>` : ''}${m.gps ? `<span class="gpsok" title="${esc(fmtCoord(m.gps.lat, m.gps.lon))} ±${m.gps.acc} m">GPS</span>` : ''}` : '<span class="note">Kontrol edilmedi</span>'}
      <button class="lnk" type="button" data-edit>${m && m.m === 'e' ? 'Düzenlendi · değiştir' : 'Düzenle / açıklama'}</button>
      ${cur && cur.m === '-' && !phs.length ? '<button class="lnk warnlnk" type="button" data-cam>İtiraza karşı fotoğraf çek</button>' : ''}
      ${cur ? '<button class="lnk" type="button" data-gal>Galeriden fotoğraf seç</button>' : ''}</div>
    ${phs.length ? `<div class="thumbs">${phs.map((ph) => `<button type="button" class="thumb" data-ph="${esc(ph.id)}" aria-label="Fotoğrafı aç"><img src="${esc(phSrc(ph))}" alt="" loading="lazy">${ph.lat != null ? '<span class="gps">GPS</span>' : ''}</button>`).join('')}</div>` : ''}
    <div class="phnote note" role="status"></div>
  </div>`;
}
function pills(s) {
  return (s.check ? `<span class="pill num">${s.checked} / ${s.check} kontrol</span>` : '') +
    (s.plus ? `<span class="pill ok num">${s.plus} +</span>` : '') +
    (s.minus ? `<span class="pill no num">${s.minus} −</span>` : '') +
    (s.ed ? `<span class="pill ed num">${s.ed} düzenlendi</span>` : '') +
    (s.auto ? `<span class="pill num">${s.auto} çayır otu otomatik +</span>` : '');
}

let curKeys = [];
function renderDetail() { saveView(); if (mode === 'koy') renderVillage(); else if (mode === 'tbs') renderTbs(); else renderFarmer(); }
// Köy içinde ada ada çalışılır: her köy için seçili ada hatırlanır ("" = bütün adalar).
let selAda = {};
try { selAda = JSON.parse(localStorage.getItem('tespit.ada') || '{}') || {}; } catch (e) {}
const saveAda = () => { try { localStorage.setItem('tespit.ada', JSON.stringify(selAda)); } catch (e) {} };
function adaGroups(v) {
  const m = new Map();
  for (const k of v.keys) { const a = PARCELS.get(k).r.ada; if (!m.has(a)) m.set(a, []); m.get(a).push(k); }
  return [...m.entries()].map(([ada, keys]) => {
    const check = keys.filter(isCheck);
    return { ada, keys, check: check.length, checked: check.filter((k) => live(k)).length };
  }).sort((a, b) => num(a.ada) - num(b.ada));
}
function adaChips(v, groups, cur) {
  const shown = groups.filter((g) => showCayir || g.check);
  const all = keyStats(v.keys);
  return `<button class="adachip all" type="button" data-ada="" aria-pressed="${cur === ''}">Tümü<small class="num">${all.checked}/${all.check}</small></button>` +
    shown.map((g) => {
      const st = g.check === 0 || g.checked === g.check ? 'done' : g.checked ? 'part' : '';
      return `<button class="adachip ${st}" type="button" data-ada="${esc(g.ada)}" aria-pressed="${cur === g.ada}">${esc(g.ada)}<small class="num">${g.check ? `${g.checked}/${g.check}` : '✓'}</small></button>`;
    }).join('');
}
function renderVillage() {
  const v = VILLAGES.find((x) => x.name === selKoy);
  if (!v) { $('detail').innerHTML = '<div class="empty">Köy seçin.</div>'; curKeys = []; return; }
  const groups = adaGroups(v);
  let cur = selAda[v.name] || '';
  if (cur && !groups.some((g) => g.ada === cur)) cur = '';
  const scope = cur ? groups.find((g) => g.ada === cur).keys : v.keys;
  const s = keyStats(scope);
  const nf = new Set(scope.map((k) => PARCELS.get(k).f.tc)).size;
  // Arama yapılınca seçili adaya bakılmaz, bütün köyde aranır.
  const base = query ? v.keys : scope;
  const nOver = v.keys.filter((k) => OVER.has(C.placeKey(PARCELS.get(k).r))).length;
  curKeys = showOver && nOver
    ? base.filter((k) => OVER.has(C.placeKey(PARCELS.get(k).r)) && parcelMatches(k))
    : base.filter((k) => (showCayir || query || isCheck(k) || live(k)) && parcelMatches(k));
  const left = scope.filter((k) => isCheck(k) && !live(k)).length;
  const where = cur ? `${v.name} köyü ${cur} adada` : `${v.name} köyünde`;
  let body = '';
  if (!curKeys.length) body = `<div class="card empty">${query ? 'Aramaya uyan parsel yok.' : cur ? 'Bu adada kontrol edilecek parsel yok; hepsi çayır otu.' : 'Bu köyde kontrol edilecek parsel yok; hepsi çayır otu.'}</div>`;
  else {
    let last = null;
    for (const k of curKeys) {
      const a = PARCELS.get(k).r.ada;
      if (a !== last && (!cur || query)) {
        const g = groups.find((x) => x.ada === a);
        body += `<div class="adahead" data-adah="${esc(a)}">Ada ${esc(a)} <span class="num">${g.checked}/${g.check} kontrol</span></div>`;
        last = a;
      }
      body += rowHtml(k, 'koy');
    }
  }
  const nextG = cur ? groups.filter((g) => g.check && g.checked < g.check && num(g.ada) > num(cur))[0] || groups.filter((g) => g.check && g.checked < g.check && g.ada !== cur)[0] : null;
  $('detail').innerHTML = `
    <div class="card fhead">
      <div class="l1"><h2>${esc(v.name)}${cur ? ` · Ada ${esc(cur)}` : ''}</h2><span class="koy">parsellerin bulunduğu köy</span></div>
      <div class="facts num"><span><b>${s.check}</b> parsel kontrol edilecek</span><span><b>${s.total - s.check}</b> çayır otu (otomatik +)</span><span><b>${nf}</b> çiftçi</span></div>
      <div class="adachips" id="adaChips" role="group" aria-label="Ada seçimi">${adaChips(v, groups, cur)}</div>
      <div class="acts">
        <div class="pills" id="pills">${pills(s)}</div>
        <span style="flex:1"></span>
        ${nOver ? `<button class="btn toggle overbtn" type="button" id="overBtn" aria-pressed="${showOver}">Alan aşımı (${nOver})</button>` : ''}
        <button class="btn" type="button" id="batchBtn">Kameradaki fotoğrafları ekle</button>
        <button class="btn toggle" type="button" id="cayirBtn" aria-pressed="${showCayir}">Çayır otlarını ${showCayir ? 'gizle' : 'göster'}</button>
        <button class="btn" type="button" id="allPlus" ${left ? '' : 'disabled'}>${cur ? `${esc(cur)} adada kalanlara + ver` : 'Kalanlara + ver'}</button>
      </div>
    </div>
    <div class="okmsg" id="batchNote" hidden></div>
    <div class="rows" id="rows">${body}</div>
    <div class="next">${nextG ? `<button class="btn primary" type="button" id="nextAda">Sonraki ada: ${esc(nextG.ada)} →</button>` : ''}</div>`;
  $('cayirBtn').onclick = () => { showCayir = !showCayir; renderVillage(); };
  $('batchBtn').onclick = openBatch;
  if ($('overBtn')) $('overBtn').onclick = () => { showOver = !showOver; renderVillage(); };
  $('adaChips').onclick = (e) => {
    const b = e.target.closest('[data-ada]');
    if (!b) return;
    selAda[v.name] = b.dataset.ada; saveAda();
    renderVillage(); $('main').scrollTop = 0;
  };
  if (nextG) $('nextAda').onclick = () => { selAda[v.name] = nextG.ada; saveAda(); renderVillage(); $('main').scrollTop = 0; };
  $('allPlus').onclick = () => ask('Kalan parsellere + verilsin mi?', `${where} kontrol edilmemiş ${scope.filter((k) => isCheck(k) && !live(k)).length} parsel “beyan doğru” (+) olarak işaretlenecek.`, '+ ver', () => {
    scope.forEach((k) => { if (isCheck(k) && !live(k)) setMark(k, { m: '+', bulk: 1 }); });
    changed(); renderDetail();
  });
}
function renderFarmer() {
  const x = FARMERS.find((y) => y.id === selF);
  if (!x) { $('detail').innerHTML = '<div class="empty">Çiftçi seçin.</div>'; curKeys = []; return; }
  const f = x.f;
  const keys = x.keys;
  curKeys = keys;
  const s = keyStats(keys);
  const alan = keys.reduce((a, k) => a + (parseFloat(PARCELS.get(k).r.alan) || 0), 0);
  const vc = {};
  keys.forEach((k) => { const v = C.villageOf(PARCELS.get(k).r); vc[v] = (vc[v] || 0) + 1; });
  const left = keys.filter((k) => !eff(k)).length;
  const list = FARMERS.filter(farmerMatches);
  const pos = list.findIndex((y) => y.id === x.id);
  const next = pos >= 0 && pos + 1 < list.length ? list[pos + 1] : null;
  const done = s.done === s.total;
  const nph = keys.reduce((a, k) => a + (PH.get(k) || []).length, 0);
  $('detail').innerHTML = `
    <div class="card fhead">
      <div class="l1"><h2>${esc(f.ad)}</h2><span class="koy">${esc(x.koy)}</span></div>
      <div class="facts num"><span>TC <b>${esc(f.tc)}</b></span><span><b>${keys.length}</b> parsel</span><span><b>${alan.toFixed(3)}</b> da</span></div>
      <div class="facts">Parseller: ${Object.entries(vc).map(([k, n]) => `<span><b>${esc(k)}</b> ${n}</span>`).join('')}</div>
      <div class="acts">
        <div class="pills" id="pills">${pills(s)}</div>
        <span style="flex:1"></span>
        <span class="pill ${done ? 'ok' : ''}" id="fstate">${done ? 'Tutanak tamam' : `${s.total - s.done} parsel bekliyor`}</span>
        <button class="btn" type="button" id="allPlus" ${left ? '' : 'disabled'}>Kalanlara + ver</button>
        <button class="btn" type="button" id="phOut" ${nph ? '' : 'disabled'}>Fotoğraflar (${nph})</button>
        <button class="btn" type="button" id="oneXlsx">Excel</button>
        <button class="btn primary" type="button" id="oneOut">Tutanağı indir</button>
      </div>
      <div id="oneMsg" role="status"></div>
    </div>
    <div class="rows" id="rows">${keys.map((k) => rowHtml(k, 'ciftci')).join('')}</div>
    <div class="next">${next ? `<button class="btn" type="button" id="nextBtn">Sonraki: ${esc(next.f.ad)} →</button>` : ''}</div>`;
  $('allPlus').onclick = () => ask('Kalan parsellere + verilsin mi?', `${f.ad} için işaretlenmemiş ${keys.filter((k) => !eff(k)).length} parsel “beyan doğru” (+) olarak işaretlenecek.`, '+ ver', () => {
    const d = C.isoToday(); const t = Date.now();
    keys.forEach((k) => { if (!eff(k)) setMark(k, { m: '+', d, bulk: 1 }); });
    changed(); renderDetail();
  });
  $('oneOut').onclick = () => download([x], $('oneMsg'), $('oneOut'));
  $('oneXlsx').onclick = () => download([x], $('oneMsg'), $('oneXlsx'), 'xlsx');
  $('phOut').onclick = () => downloadPhotos(x, $('oneMsg'), $('phOut'));
  if (next) $('nextBtn').onclick = () => { selF = next.id; renderSide(); renderDetail(); $('main').scrollTop = 0; };
}
function changed() { queueSave(); schedulePush(); renderProg(); renderSide(); renderSyncBar(); }
function updateRow(k) {
  const el = [...$('rows').querySelectorAll('[data-k]')].find((e) => e.dataset.k === k);
  if (el) el.outerHTML = rowHtml(k, mode);
  let keys = curKeys;
  if (mode === 'koy') {
    const v = VILLAGES.find((x) => x.name === selKoy);
    const groups = adaGroups(v);
    const cur = selAda[v.name] && groups.some((g) => g.ada === selAda[v.name]) ? selAda[v.name] : '';
    keys = cur ? groups.find((g) => g.ada === cur).keys : v.keys;
    $('adaChips').innerHTML = adaChips(v, groups, cur);
    document.querySelectorAll('[data-adah]').forEach((h) => { const g = groups.find((x) => x.ada === h.dataset.adah); if (g) h.querySelector('span').textContent = `${g.checked}/${g.check} kontrol`; });
  }
  const s = keyStats(keys);
  $('pills').innerHTML = pills(s);
  $('allPlus').disabled = mode === 'koy' ? !keys.some((x) => isCheck(x) && !live(x)) : !keys.some((x) => !eff(x));
  const fs = $('fstate');
  if (fs) { const d = s.done === s.total; fs.className = 'pill' + (d ? ' ok' : ''); fs.textContent = d ? 'Tutanak tamam' : `${s.total - s.done} parsel bekliyor`; }
}
function refreshRows(keys) { for (const k of keys) updateRow(k); }
$('detail').addEventListener('click', (e) => {
  const row = e.target.closest('[data-k]');
  if (!row) return;
  const k = row.dataset.k;
  const mk = e.target.closest('[data-mk]');
  if (mk) {
    const v = mk.dataset.mk;
    const cur = live(k);
    const sh = sharesOf(k);
    if (cur && cur.m === v) {
      clearMark(k);
      sh.forEach((x) => { const c = live(x); if (c && c.m === v && !c.olc) clearMark(x); else if (c && c.m === v) setMark(x, { m: v, olc: c.olc }); });
    } else {
      setMark(k, { m: v, olc: cur ? cur.olc : undefined });
      sh.forEach((x) => { const c = live(x); if (!c || c.m !== 'e') setMark(x, { m: v, olc: c ? c.olc : undefined }); });
      attachGps([k, ...sh.filter((x) => live(x))]);
    }
    refreshRows([k, ...sh]); changed();
    return;
  }
  const th = e.target.closest('[data-ph]');
  if (th) { openPhoto(th.dataset.ph); return; }
  if (e.target.closest('[data-cam]')) { openCamera(k); return; }
  if (e.target.closest('[data-gal]')) { openGallery(k); return; }
  if (e.target.closest('[data-olcbtn]')) {
    const prev = olcK;
    olcK = olcK === k ? null : k;
    if (prev && prev !== k) updateRow(prev);
    updateRow(k);
    if (olcK) setTimeout(() => { const i = $('olcIn'); if (i) { i.focus(); i.select(); } }, 20);
    return;
  }
  if (e.target.closest('[data-olcx]')) { olcK = null; updateRow(k); return; }
  if (e.target.closest('[data-olcdel]')) {
    const c = live(k);
    if (c) setMark(k, { m: c.m, v: c.v });
    olcK = null; updateRow(k); changed();
    return;
  }
  if (e.target.closest('[data-edit]')) openEdit(k);
});
$('detail').addEventListener('submit', (e) => {
  const fm = e.target.closest('[data-olcform]');
  if (!fm) return;
  e.preventDefault();
  const k = fm.closest('[data-k]').dataset.k;
  const olc = C.fmtArea($('olcIn').value);
  if (!olc) { $('olcErr').textContent = 'Dönüm olarak bir sayı girin, ör. 10 ya da 7,5'; return; }
  const c = live(k);
  setMark(k, c ? { m: c.m, v: c.v, olc } : { m: '+', olc });
  attachGps([k]);
  olcK = null; updateRow(k); changed();
});

// ---------- TBS'ye girilecek değişiklikler ----------
let tbsKoy = '';
let tbsFilter = 'bekleyen';   // bekleyen | islenen | hepsi
// Tespiti beyandan farklı olan (− , ölçüm, elle düzenleme) parseller: TBS'de düzeltilmesi gerekenler.
function tbsChange(k) {
  const m = live(k);
  if (!m || m.auto) return null;
  const { r } = PARCELS.get(k);
  const t = C.tespit(r, m, S);
  if (!t) return null;
  if (t.urun === r.urun && t.alan === r.alan) return null;   // yalnız açıklama eklenmişse TBS'de değişecek bir şey yok
  return t;
}
function tbsKind(k) {
  const { r } = PARCELS.get(k); const m = live(k); const t = C.tespit(r, m, S);
  const a = [];
  if (t.urun !== r.urun) a.push('Ürün değişti');
  if (t.alan !== r.alan) a.push(m.olc ? 'Alan değişti (ölçüm)' : 'Alan değişti');
  return a.join(', ');
}
const tbsKeys = (keys) => keys.filter((k) => PARCELS.has(k) && tbsChange(k)).sort((a, b) => {
  const A = PARCELS.get(a), B = PARCELS.get(b);
  return C.villageOf(A.r).localeCompare(C.villageOf(B.r), 'tr') || num(A.r.ada) - num(B.r.ada) || num(A.r.parsel) - num(B.r.parsel) || A.f.ad.localeCompare(B.f.ad, 'tr');
});
function tbsScope() {
  const v = tbsKoy ? VILLAGES.find((x) => x.name === tbsKoy) : null;
  const keys = tbsKeys(v ? v.keys : [...PARCELS.keys()]).filter((k) => parcelMatches(k));
  return keys;
}
function tbsRowHtml(k) {
  const { r, f } = PARCELS.get(k);
  const m = live(k);
  const t = tbsChange(k);
  const kind = tbsKind(k);
  return `<div class="card tbsrow ${m.tbs ? 'done' : ''}" data-k="${esc(k)}">
    <div class="tinfo">
      <div class="ap num">${esc(C.villageOf(r))} · ${esc(r.ada)} / ${esc(r.parsel)}</div>
      <div class="who">${esc(f.ad)} <span>· TC ${esc(f.tc)}</span></div>
      <div class="chg"><span class="kind">${kind}</span> Beyan: <b>${esc(r.urun)} ${esc(r.alan)} da</b> → Tespit: <b>${esc(t.urun)} ${esc(t.alan)} da</b>${t.acik ? ` · ${esc(t.acik)}` : ''}</div>
      <div class="note">${m.by ? esc(m.by) + ' · ' : ''}${esc(C.trDate(m.d || ''))}${m.tbs ? ` · TBS'ye işlendi: ${esc(m.tbs.by || '')} ${esc(trDT(m.tbs.t))}` : ''}</div>
    </div>
    <label class="tbscheck"><input type="checkbox" data-tbs ${m.tbs ? 'checked' : ''}> TBS'ye işlendi</label>
  </div>`;
}
function renderTbs() {
  const keys = tbsScope();
  const done = keys.filter((k) => live(k).tbs).length;
  const shown = keys.filter((k) => tbsFilter === 'hepsi' || (tbsFilter === 'islenen' ? live(k).tbs : !live(k).tbs));
  const chip = (id, label, n) => `<button class="chip" type="button" data-tf="${id}" aria-pressed="${tbsFilter === id}">${label} ${n}</button>`;
  $('detail').innerHTML = `
    <div class="card fhead">
      <div class="l1"><h2>${tbsKoy ? esc(tbsKoy) : 'Bütün köyler'}</h2><span class="koy">TBS'ye girilecek değişiklikler</span></div>
      <div class="facts num"><span><b>${keys.length}</b> değişiklik</span><span><b>${done}</b> işlendi</span><span><b>${keys.length - done}</b> bekliyor</span></div>
      <div class="acts">
        <div class="pills">${chip('bekleyen', 'Bekleyen', keys.length - done)}${chip('islenen', 'İşlenen', done)}${chip('hepsi', 'Hepsi', keys.length)}</div>
        <span style="flex:1"></span>
        <button class="btn primary" type="button" id="tbsXlsx" ${keys.length ? '' : 'disabled'}>Listeyi Excel olarak indir</button>
      </div>
      <div id="tbsMsg" role="status"></div>
    </div>
    <div class="rows" id="rows">${shown.length ? shown.map(tbsRowHtml).join('') : `<div class="card empty">${keys.length ? 'Bu süzgece uyan kayıt yok.' : 'Değişiklik yok. − verilen, ölçülen ya da elle düzenlenen parseller burada listelenir.'}</div>`}</div>`;
  $('detail').querySelectorAll('[data-tf]').forEach((b) => { b.onclick = () => { tbsFilter = b.dataset.tf; renderTbs(); }; });
  $('tbsXlsx').onclick = () => downloadTbs(keys);
}
$('detail').addEventListener('change', (e) => {
  const cb = e.target.closest('[data-tbs]');
  if (!cb) return;
  const k = cb.closest('[data-k]').dataset.k;
  const m = live(k);
  if (!m) return;
  const o = { ...m, t: Date.now() };
  if (cb.checked) o.tbs = { t: Date.now(), by: S.p1 || '' }; else delete o.tbs;
  M[k] = o; dirty.add(vhOf(k));
  queueSave(); schedulePush(); renderSide();
  const row = cb.closest('[data-k]');
  row.outerHTML = tbsRowHtml(k);
  // Başlıktaki sayıları güncelle (satır, süzgeç değişene kadar yerinde kalır).
  const keys = tbsScope(); const done = keys.filter((x) => live(x).tbs).length;
  const facts = $('detail').querySelector('.fhead .facts');
  if (facts) facts.innerHTML = `<span><b>${keys.length}</b> değişiklik</span><span><b>${done}</b> işlendi</span><span><b>${keys.length - done}</b> bekliyor</span>`;
  const set = (id, label, n) => { const b = $('detail').querySelector(`[data-tf="${id}"]`); if (b) b.textContent = `${label} ${n}`; };
  set('bekleyen', 'Bekleyen', keys.length - done); set('islenen', 'İşlenen', done); set('hepsi', 'Hepsi', keys.length);
});
async function downloadTbs(keys) {
  const msg = $('tbsMsg');
  msg.innerHTML = '<p class="status">Liste hazırlanıyor…</p>';
  try {
    const cols = [
      { h: 'Köy', w: 14 }, { h: 'Ada', w: 7 }, { h: 'Parsel', w: 8 }, { h: 'Çiftçi', w: 24 }, { h: 'TC / Vergi No', w: 14 },
      { h: 'Beyan ürün', w: 22 }, { h: 'Beyan alan (da)', w: 10, num: true }, { h: 'Tespit ürün', w: 22 }, { h: 'Tespit alan (da)', w: 10, num: true },
      { h: 'Değişiklik', w: 16 }, { h: 'Açıklama', w: 20 }, { h: 'Tespit eden', w: 16 }, { h: 'Tespit tarihi', w: 12 }, { h: 'Konum (tespit anı)', w: 22 }, { h: "TBS'ye işlendi", w: 22 }
    ];
    const n = (v) => { const x = parseFloat(String(v).replace(',', '.')); return isNaN(x) ? v : x; };
    const rows = keys.map((k) => {
      const { r, f } = PARCELS.get(k); const m = live(k); const t = tbsChange(k);
      const kind = tbsKind(k);
      return [C.villageOf(r), r.ada, r.parsel, f.ad, f.tc, r.urun, n(r.alan), t.urun, n(t.alan), kind, t.acik || '', m.by || '', C.trDate(m.d || ''), m.gps ? `${m.gps.lat}, ${m.gps.lon}` : '', m.tbs ? `Evet · ${m.tbs.by || ''} ${trDT(m.tbs.t)}` : ''];
    });
    const out = await C.buildTableXlsx(JSZip, 'TBS değişiklikleri', `TBS'ye girilecek değişiklikler — ${tbsKoy || 'Bütün köyler'} — ${C.trDate(C.isoToday())}`, cols, rows);
    const name = `TBS_Degisiklikleri_${fileSafe(tbsKoy || 'TUMU')}_${C.isoToday()}.xlsx`;
    msg.innerHTML = '<p class="status">İndirme onayı bekleniyor…</p>';
    await saveFile(name, out, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    msg.innerHTML = `<div class="okmsg">${esc(name)} indirildi (${rows.length} satır).</div>`;
  } catch (err) {
    console.error(err);
    msg.innerHTML = `<div class="warn">${err && err.code === 'declined' ? 'İndirme onaylanmadı.' : 'Liste oluşturulamadı.'}</div>`;
  }
}

// ---------- Elle düzenleme ----------
let edK = null;
// ---------- Sesle açıklama ----------
// İnternet varken uygulamanın mikrofonu; yoksa klavyenin kendi mikrofonu (çevrimdışı Türkçe paketiyle) kullanılır.
let dict = null, dictTr = true, dictBase = '';
async function initDictation() {
  try { dict = window.claude && typeof window.claude.use === 'function' ? await window.claude.use('dictation') : null; } catch (e) { dict = null; }
  $('micBtn').hidden = !dict;
  $('micHint').textContent = dict ? '' : 'Sesle yazmak için klavyedeki mikrofon tuşunu kullanabilirsiniz.';
  if (!dict) return;
  dict.onPartial(({ text }) => { $('micLive').textContent = text; });
  dict.onFinal(({ text }) => {
    const ta = $('edAcik');
    ta.value = ta.value ? ta.value + text : text.replace(/^\s+/, '');
    $('micLive').textContent = '';
  });
  dict.onState(({ state, reason }) => {
    const on = state !== 'idle';
    $('micBtn').setAttribute('aria-pressed', String(on));
    $('micLbl').textContent = state === 'listening' ? 'Dinliyor… bitirmek için dokunun' : state === 'starting' ? 'Mikrofon açılıyor…' : state === 'finishing' ? 'Son sözler yazılıyor…' : 'Sesle yaz';
    if (!on && reason === 'unavailable' && dictTr) { dictTr = false; $('micHint').textContent = 'Türkçe ses tanıma bu görünümde açılamadı, yeniden deneyin.'; }
    else if (!on && reason === 'not_granted') $('micHint').textContent = 'Mikrofon izni verilmedi. Klavyedeki mikrofon tuşunu kullanabilirsiniz.';
    else if (!on && (reason === 'upstream_error' || reason === 'rate_limited')) $('micHint').textContent = 'Ses tanıma şu an çalışmıyor (internet gerekir). Klavyedeki mikrofonu deneyin.';
    else if (!on) $('micHint').textContent = '';
  });
}
$('micBtn').addEventListener('click', () => {
  if (!dict) return;
  if (dict.active()) { dict.stop().catch(() => {}); return; }
  $('micHint').textContent = '';
  dict.start(dictTr ? { language: 'tr' } : undefined).catch((e) => {
    if (e && e.code === 'unavailable' && dictTr) dictTr = false;
    if (e && e.code === 'not_granted') $('micHint').textContent = 'Mikrofon izni verilmedi. Klavyedeki mikrofon tuşunu kullanabilirsiniz.';
    else if (e && e.code !== 'stopped') $('micHint').textContent = 'Ses tanıma başlatılamadı (internet gerekir). Klavyedeki mikrofonu deneyin.';
  });
});
function stopDictation() { if (dict && dict.active()) dict.stop().catch(() => {}); $('micLive').textContent = ''; }

function openEdit(k) {
  edK = k;
  const { r, f } = PARCELS.get(k);
  const t = C.tespit(r, live(k) || { m: '+' }, S);
  $('edSub').textContent = `${f.ad} · Ada ${r.ada} / Parsel ${r.parsel} · ${C.villageOf(r)} · Beyan: ${r.urun}, ${r.alan} da`;
  $('edUrun').value = t.urun; $('edAlan').value = t.alan; $('edAcik').value = t.acik;
  openDlg('dlgEdit', 'edUrun');
}
$('edForm').addEventListener('submit', (e) => {
  e.preventDefault();
  const { r } = PARCELS.get(edK);
  const urun = $('edUrun').value.trim();
  const acik = $('edAcik').value.trim();
  const alanIn = $('edAlan').value.trim();
  const alan = C.fmtArea(alanIn) || alanIn;
  const olc = alan && alan !== r.alan ? alan : undefined;   // beyandan farklı alan = ölçüm
  if (urun === r.urun && !acik) setMark(edK, { m: '+', olc });
  else if (urun === S.eksiUrun && acik === (S.eksiAcik || '')) setMark(edK, { m: '-', olc });
  else setMark(edK, { m: 'e', v: { urun, acik }, olc });
  attachGps([edK]);
  stopDictation();
  closeDlg('dlgEdit'); updateRow(edK); changed();
});
$('edClear').addEventListener('click', () => { clearMark(edK); closeDlg('dlgEdit'); updateRow(edK); changed(); });

// ---------- Ayarlar ----------
function syncTarih() { $('stTarihF').hidden = $('stMod').value !== 'sabit'; }
$('stMod').addEventListener('change', syncTarih);
$('setBtn').addEventListener('click', () => {
  $('stUrun').value = S.eksiUrun; $('stAcik').value = S.eksiAcik;
  $('stP1').value = S.p1; $('stP2').value = S.p2; $('stP3').value = S.p3; $('stMod').value = S.tarihMod || 'bos'; $('stTarih').value = S.tarih;
  syncTarih();
  openDlg('dlgSet', 'stUrun');
});
$('stToday').addEventListener('click', () => { $('stTarih').value = C.trDate(C.isoToday()); });
$('stForm').addEventListener('submit', (e) => {
  e.preventDefault();
  S.eksiUrun = $('stUrun').value.trim() || C.DEFAULTS.eksiUrun; S.eksiAcik = $('stAcik').value.trim();
  S.p1 = $('stP1').value.trim(); S.p2 = $('stP2').value.trim(); S.p3 = $('stP3').value.trim();
  S.tarihMod = $('stMod').value; S.tarih = $('stTarih').value.trim();
  S._t = Date.now(); settingsDirty = true;
  saveSettings(); closeDlg('dlgSet'); schedulePush();
  if (DOCS.length && !$('work').hidden) renderDetail();
});

// ---------- Tutanak PDF ----------
function yearOf() { return (DOCS.find((d) => d.year) || {}).year || String(new Date().getFullYear()); }
async function saveFile(name, bytes, type) {
  const blob = bytes instanceof Blob ? bytes : new Blob([bytes], { type: type || 'application/pdf' });
  if (window.claude && typeof window.claude.use === 'function') {
    const dl = await window.claude.use('downloads');
    if (!dl) throw { code: 'unavailable' };
    return dl.save({ filename: name, data: blob });
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return { status: 'saved' };
}
const fileSafe = (s) => String(s).replace(/[^\p{L}\p{N}_ .-]+/gu, '_').replace(/\s+/g, '_');
async function download(xs, msgEl, btn, fmt) {
  if (btn) btn.disabled = true;
  msgEl.innerHTML = '<p class="status">PDF hazırlanıyor…</p>';
  try {
    const list = xs.map((x) => ({ f: x.f, koy: x.koy, rows: x.keys.map((k) => PARCELS.get(k).r) }));
    const xlsx = fmt === 'xlsx';
    const out = xlsx
      ? await C.buildXlsx(JSZip, list, markOfRow, S, FONTS, { year: yearOf() })
      : await C.buildPdf(PDFLib, list, markOfRow, S, FONTS, { year: yearOf(), backup: { v: 3, marks: M, settings: S, at: Date.now() } });
    const ext = xlsx ? 'xlsx' : 'pdf';
    const name = xs.length === 1
      ? `Tutanak_${fileSafe(xs[0].f.ad)}_${C.isoToday()}.${ext}`
      : `Tutanaklar_${fileSafe($('outKoy').value || 'TUMU')}_${xs.length}_ciftci_${C.isoToday()}.${ext}`;
    msgEl.innerHTML = '<p class="status">İndirme onayı bekleniyor…</p>';
    await saveFile(name, out, xlsx ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf');
    msgEl.innerHTML = `<div class="okmsg">${esc(name)} indirildi.</div>`;
  } catch (err) {
    console.error(err);
    const code = err && err.code;
    msgEl.innerHTML = `<div class="warn">${code === 'declined' ? 'İndirme onaylanmadı. Yeniden deneyebilirsiniz.'
      : code === 'unavailable' || code === 'not_granted' || code === 'capability_disabled' ? 'Bu görünümde dosya indirilemiyor. Sayfayı Claude uygulamasında ya da tarayıcıda açıp yeniden deneyin.'
      : code === 'rate_limited' ? 'Açık bir indirme onayı var. Biraz bekleyip yeniden deneyin.'
      : 'PDF oluşturulamadı: ' + esc(err && err.message ? err.message : code || err)}</div>`;
  }
  if (btn) btn.disabled = false;
}
function outSelection() {
  const koy = $('outKoy').value;
  const onlyDone = $('outDone').checked;
  return FARMERS.filter((x) => (!koy || x.koy === koy) && (!onlyDone || farmerDone(x)));
}
function renderOutCount() {
  const xs = outSelection();
  const rows = xs.reduce((a, x) => a + x.keys.length, 0);
  $('outCount').textContent = xs.length ? `${xs.length} çiftçi, ${rows} parsel` : 'Seçime uyan çiftçi yok.';
  $('outGo').disabled = !xs.length;
  $('outXlsx').disabled = !xs.length;
}
$('outBtn').addEventListener('click', () => {
  const t = totals();
  $('outSum').innerHTML = `<div><b class="num">${t.fdone}/${t.farmers}</b><span>tutanağı tamam</span></div><div><b class="num">${t.checked}/${t.check}</b><span>kontrol</span></div><div><b class="num" style="color:var(--no)">${t.minus}</b><span>− çayır otu</span></div><div><b class="num" style="color:var(--ed)">${t.ed}</b><span>düzenlendi</span></div>`;
  const ks = [...new Set(FARMERS.map((x) => x.koy))].sort((a, b) => a.localeCompare(b, 'tr'));
  const prev = $('outKoy').value;
  $('outKoy').innerHTML = '<option value="">Bütün köyler</option>' + ks.map((k) => `<option value="${esc(k)}">${esc(k)}</option>`).join('');
  $('outKoy').value = ks.includes(prev) ? prev : (mode === 'koy' && ks.includes(selKoy) ? selKoy : '');
  $('outMsg').innerHTML = '';
  renderOutCount();
  openDlg('dlgOut', 'outKoy');
});
$('outKoy').addEventListener('change', renderOutCount);
$('outDone').addEventListener('change', renderOutCount);
$('outGo').addEventListener('click', () => download(outSelection(), $('outMsg'), $('outGo')));
$('outXlsx').addEventListener('click', () => download(outSelection(), $('outMsg'), $('outXlsx'), 'xlsx'));

// ---------- Fotoğraf eki PDF ----------
async function downloadPhotos(x, msgEl, btn) {
  btn.disabled = true;
  msgEl.innerHTML = '<p class="status">Fotoğraflar hazırlanıyor…</p>';
  try {
    const items = [];
    for (const k of x.keys) {
      const { r } = PARCELS.get(k);
      for (const ph of PH.get(k) || []) {
        items.push({
          jpg: await photoBytes(ph),
          title: `FOTOĞRAF EKİ — ${x.f.ad} (TC ${x.f.tc})`,
          sub: `${C.villageOf(r)} · Ada ${r.ada} / Parsel ${r.parsel} · Beyan: ${r.urun} ${r.alan} da · ${trDT(ph.t)} · ${ph.lat != null ? 'GPS ' + fmtCoord(ph.lat, ph.lon) : 'konum yok'}`
        });
      }
    }
    const out = await C.buildPhotoPdf(PDFLib, items, FONTS, 'Fotoğraf eki — ' + x.f.ad);
    const name = `Fotograflar_${fileSafe(x.f.ad)}_${C.isoToday()}.pdf`;
    msgEl.innerHTML = '<p class="status">İndirme onayı bekleniyor…</p>';
    await saveFile(name, out);
    msgEl.innerHTML = `<div class="okmsg">${esc(name)} indirildi (${items.length} fotoğraf).</div>`;
  } catch (err) {
    console.error(err);
    const code = err && err.code;
    msgEl.innerHTML = `<div class="warn">${code === 'declined' ? 'İndirme onaylanmadı.' : 'Fotoğraf PDF\'i oluşturulamadı: ' + esc(err && err.message ? err.message : code || err)}</div>`;
  }
  btn.disabled = false;
}
$('phClose').addEventListener('click', () => closeDlg('dlgPhoto'));
$('phDl').addEventListener('click', async () => {
  if (!viewPh) return;
  try {
    const { r } = PARCELS.get(viewPh.key);
    const bytes = await photoBytes(viewPh);
    const name = `Foto_${fileSafe(C.villageOf(r))}_${r.ada}-${r.parsel}_${new Date(viewPh.t).toISOString().slice(0, 10)}.jpg`;
    const blob = new Blob([bytes], { type: 'image/jpeg' });
    if (window.claude && typeof window.claude.use === 'function') {
      const dl = await window.claude.use('downloads');
      if (!dl) throw { code: 'unavailable' };
      await dl.save({ filename: name, data: blob });
    } else { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); }
    $('phMsg').textContent = name + ' indirildi.';
  } catch (err) { $('phMsg').textContent = err && err.code === 'declined' ? 'İndirme onaylanmadı.' : 'İndirilemedi.'; }
});
$('phDel').addEventListener('click', () => {
  const ph = viewPh;
  if (!ph) return;
  closeDlg('dlgPhoto');
  ask('Fotoğraf silinsin mi?', 'Fotoğraf bu tabletten ve hesabınızdan kalıcı olarak silinir.', 'Sil', async () => {
    const l = PH.get(ph.key) || [];
    PH.set(ph.key, l.filter((x) => x.id !== ph.id));
    await delPhotoDb(ph.id);
    if (ph.asset && assetsApi) { try { await assetsApi.delete(ph.asset.id); } catch (e) {} }
    deletedPh.add(ph.id); saveDeleted();
    dirtyP.add(phVh(ph.key));
    updateRow(ph.key); schedulePush();
    if (mode === 'ciftci') renderDetail();
  });
});

// ---------- Ekip eşitlemesi (dosyayla) ----------
// Dosyada TC, ad, ada/parsel açık yazılmaz; parseller buluttaki gibi kodla tutulur.
// Alan cihazda aynı Excel yüklüyse kodlar parsellere çözülür.
let PENDING = {};   // çözülemeyen kod -> işaret (Excel yüklenince yeniden denenir)
let SYNCLOG = (() => { try { return JSON.parse(localStorage.getItem('tespit.synclog') || '{}'); } catch (e) { return {}; } })();
const saveSyncLog = () => { try { localStorage.setItem('tespit.synclog', JSON.stringify(SYNCLOG)); } catch (e) {} };
const lastExport = () => { try { return +localStorage.getItem('tespit.lastExport') || 0; } catch (e) { return 0; } };
// Bu cihazda verilip henüz eşitleme dosyasına alınmamış işaretler/fotoğraflar.
function unexported() {
  const le = lastExport();
  let n = 0, oldest = 0, ph = 0;
  for (const k in M) { const m = M[k]; if (m.dev === DEVICE && (m.t || 0) > le) { n++; if (!oldest || m.t < oldest) oldest = m.t; } }
  PH.forEach((l) => l.forEach((p) => { if ((p.addedAt || 0) > le) ph++; }));
  return { n, ph, oldest };
}
function renderSyncBar() {
  const bar = $('syncBar');
  if (!bar) return;
  const u = unexported();
  const today0 = new Date(); today0.setHours(0, 0, 0, 0);
  const late = new Date().getHours() >= 16;
  const show = DOCS.length && (u.n || u.ph) && (late || (u.oldest && u.oldest < today0.getTime()));
  bar.hidden = !show;
  if (show) $('syncBarTxt').textContent = `Bu cihazda eşitleme dosyasına alınmamış ${u.n} işaret${u.ph ? ` ve ${u.ph} fotoğraf` : ''} var. Gün bitmeden dosyayı çıkarıp merkeze gönderin.`;
}
function renderSyncLog() {
  const el = $('syncLog');
  if (!el) return;
  const items = Object.values(SYNCLOG).sort((a, b) => b.at - a.at);
  el.innerHTML = items.length ? `<h4>Gelen dosyalar</h4><ul class="difflist">${items.map((x) => {
    const old = Date.now() - x.at > 24 * 3600e3;
    return `<li class="${old ? 'stale' : ''}"><b>${esc(x.by || 'adsız cihaz')}</b> · dosya ${esc(trDT(x.at))} · yüklendi ${esc(trDT(x.imported))}${old ? ' · <b>1 günden eski</b>' : ''}</li>`;
  }).join('')}</ul>` : '';
}
let CONFLICTS = []; // [{ k, a: işaret, b: işaret }] — iki kişinin farklı işaret verdiği parseller
async function loadConflicts() { const r = await safe(tx('kv', 'readonly', (st) => st.get('conflicts')), null); CONFLICTS = (r && r.data) || []; }
const saveConflicts = () => safe(tx('kv', 'readwrite', (st) => st.put({ id: 'conflicts', data: CONFLICTS })));
// coreSame: tespitin kendisi aynı mı (TBS'ye işlendi bilgisi hariç). sameMark: her şey aynı mı.
const coreSame = (a, b) => (a.m === 'x' && b.m === 'x') || (a.m === b.m && (a.olc || '') === (b.olc || '') && JSON.stringify(a.v || null) === JSON.stringify(b.v || null));
const sameMark = (a, b) => coreSame(a, b) && !!a.tbs === !!b.tbs;
function markLabel(m) {
  if (!m || m.m === 'x') return 'işaret kaldırılmış';
  const base = m.m === '+' ? '+ Uygun' : m.m === '-' ? '− Çayır otu' : 'Elle: ' + ((m.v && m.v.urun) || '');
  return base + (m.olc ? ` · ölçüm ${m.olc} da` : '') + (m.by ? ` · ${m.by}` : '') + (m.t ? ` · ${trDT(m.t)}` : '');
}
const b64enc = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
const b64dec = (str) => { const bin = atob(str); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
async function loadPending() { const r = await safe(tx('kv', 'readonly', (st) => st.get('pending')), null); PENDING = (r && r.data) || {}; }
const savePending = () => safe(tx('kv', 'readwrite', (st) => st.put({ id: 'pending', data: PENDING })));
function revMap() { const rev = new Map(); PARCELS.forEach((_, k) => rev.set(hkey(k), k)); for (const k in M) rev.set(hkey(k), k); return rev; }
function resolvePending() {
  const rev = revMap();
  let n = 0;
  for (const h in PENDING) {
    const k = rev.get(h);
    if (!k) continue;
    const a = PENDING[h], b = M[k];
    if (!b || (a.t || 0) > (b.t || 0)) { M[k] = a; dirty.add(vhOf(k)); n++; }
    delete PENDING[h];
  }
  if (n) { queueSave(); schedulePush(); }
  savePending();
  return n;
}
async function exportSync(withPhotos) {
  const msg = $('syncMsg');
  msg.className = 'status'; msg.textContent = 'Eşitleme dosyası hazırlanıyor…';
  try {
    const marks = {};
    for (const k in M) marks[hkey(k)] = M[k];
    for (const h in PENDING) if (!marks[h]) marks[h] = PENDING[h];
    const photos = [];
    if (withPhotos) {
      for (const [k, list] of PH) for (const ph of list) {
        try { photos.push({ id: ph.id, h: hkey(k), t: ph.t, lat: ph.lat ?? null, lon: ph.lon ?? null, jpg: b64enc(await photoBytes(ph)) }); }
        catch (e) { console.error(e); }
      }
    }
    const who = S.p1 || 'cihaz';
    const body = JSON.stringify({ app: 'arazi-tespit', v: 1, by: who, dev: DEVICE, at: Date.now(), year: yearOf(), marks, photos, deleted: [...deletedPh] });
    const d = new Date();
    const name = `Esitleme_${fileSafe(who)}_${C.isoToday()}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}.json`;
    msg.textContent = 'İndirme onayı bekleniyor…';
    await saveFile(name, new Blob([body], { type: 'application/json' }));
    const nm = Object.values(marks).filter((m) => m.m !== 'x').length;
    try { localStorage.setItem('tespit.lastExport', String(Date.now())); } catch (e) {}
    renderSyncBar();
    msg.className = 'okmsg'; msg.textContent = `${name} indirildi: ${nm} işaret${withPhotos ? `, ${photos.length} fotoğraf` : ''}. Bu dosyayı merkezdeki cihaza gönderin.`;
  } catch (err) {
    console.error(err);
    msg.className = 'warn'; msg.textContent = err && err.code === 'declined' ? 'İndirme onaylanmadı.' : 'Eşitleme dosyası oluşturulamadı.';
  }
}
async function importSync(files) {
  const msg = $('syncMsg');
  msg.className = 'status';
  const out = [];
  for (const file of files) {
    try {
      msg.textContent = `${file.name} okunuyor…`;
      const data = JSON.parse(await file.text());
      if (!data || data.app !== 'arazi-tespit' || !data.marks) { out.push(`${file.name}: eşitleme dosyası değil`); continue; }
      const rev = revMap();
      let added = 0, updated = 0, waiting = 0, phAdded = 0, phDel = 0;
      let same = 0, conf = 0;
      const resolved = [];
      for (const h in data.marks) {
        const a = data.marks[h];
        const k = rev.get(h);
        if (!k) { const p = PENDING[h]; if (!p || (a.t || 0) > (p.t || 0)) { PENDING[h] = a; waiting++; } continue; }
        resolved.push([k, a]);
        const b = M[k];
        if (b && sameMark(a, b)) { same++; continue; }
        // İki tarafta da gerçek (kaldırılmamış) ama farklı işaret varsa çakışma olarak kaydet.
        if (b && b.m !== 'x' && a.m !== 'x' && !a.res && !coreSame(a, b) && (a.dev || a.by || '') !== (b.dev || b.by || '')) {
          CONFLICTS = CONFLICTS.filter((c) => c.k !== k);
          CONFLICTS.push({ k, a: b, b: a });
          conf++;
        }
        if (!b || (a.t || 0) > (b.t || 0)) { if (!b || b.m === 'x') added++; else updated++; M[k] = a; dirty.add(vhOf(k)); }
      }
      // Doğrulama: dosyadaki her işaret ya aynen ya da daha yenisiyle bu cihazda bulunmalı.
      const bad = resolved.filter(([k, a]) => !M[k] || (!sameMark(M[k], a) && (M[k].t || 0) < (a.t || 0))).length;
      for (const id of data.deleted || []) {
        if (deletedPh.has(id)) continue;
        deletedPh.add(id);
        const ph = photoById(id);
        if (ph) {
          PH.set(ph.key, (PH.get(ph.key) || []).filter((x) => x.id !== id));
          await delPhotoDb(id);
          dirtyP.add(phVh(ph.key));
          phDel++;
        }
      }
      saveDeleted();
      for (const e of data.photos || []) {
        if (deletedPh.has(e.id) || photoById(e.id)) continue;
        const k = rev.get(e.h);
        if (!k) continue;
        const rec = { id: e.id, key: k, t: e.t, lat: e.lat, lon: e.lon, blob: new Blob([b64dec(e.jpg)], { type: 'image/jpeg' }), asset: null, up: false };
        addPhotoRec(rec); await putPhoto(rec); dirtyP.add(phVh(k)); phAdded++;
      }
      const total = Object.keys(data.marks).length;
      SYNCLOG[data.dev || data.by || file.name] = { by: data.by || '', at: data.at || 0, imported: Date.now(), n: Object.keys(data.marks).length };
      saveSyncLog();
      out.push(`${data.by || file.name} (${trDT(data.at || Date.now())}): dosyada ${total} kayıt → ${added} yeni, ${updated} güncellenen, ${same} zaten aynıydı${phAdded ? `; ${phAdded} fotoğraf` : ''}${phDel ? `; ${phDel} silinen fotoğraf` : ''}${waiting ? `; ${waiting} kayıt bu cihazdaki Excel'de yok, Excel yüklenince işlenecek` : ''}${conf ? `; ${conf} çakışma` : ''}. Doğrulama: ${bad ? `${bad} kayıt işlenemedi!` : 'eksiksiz'}.`);
    } catch (err) {
      console.error(err);
      out.push(`${file.name}: okunamadı`);
    }
  }
  await savePending();
  await saveConflicts();
  await saveLocal();
  schedulePush();
  renderHome();
  msg.className = 'okmsg';
  msg.innerHTML = out.map(esc).join('<br>');
  renderSyncLog();
  renderConflictNote();
}
function renderConflictNote() {
  const el = $('confNote');
  if (!el) return;
  el.hidden = !CONFLICTS.length;
  el.innerHTML = CONFLICTS.length ? `<div class="warn">${CONFLICTS.length} parselde iki kişi farklı işaret vermiş. Şimdilik en son verilen geçerli. <button class="btn" type="button" id="confOpen">Çakışmaları incele</button></div>` : '';
  const b = $('confOpen'); if (b) b.onclick = openConflicts;
}
function openConflicts() {
  $('confList').innerHTML = CONFLICTS.map((c, i) => {
    const p = PARCELS.get(c.k);
    const head = p ? `${esc(C.villageOf(p.r))} · Ada ${esc(p.r.ada)} / Parsel ${esc(p.r.parsel)} · ${esc(p.f.ad)} · Beyan: ${esc(p.r.urun)} ${esc(p.r.alan)} da` : 'Bu cihazda yüklü olmayan parsel';
    const cur = M[c.k];
    const pick = (m, side) => `<label class="confopt ${cur && sameMark(cur, m) ? 'on' : ''}"><input type="radio" name="cf${i}" value="${side}" ${cur && sameMark(cur, m) ? 'checked' : ''}> ${esc(markLabel(m))}</label>`;
    return `<li><div class="confhead">${head}</div>${pick(c.a, 'a')}${pick(c.b, 'b')}</li>`;
  }).join('');
  openDlg('dlgConf', 'confDone');
}
$('confDone').addEventListener('click', async () => {
  // Seçilen işaret yeni zamanla kaydedilir ki bütün cihazlarda bu karar geçerli olsun.
  const keep = [];
  CONFLICTS.forEach((c, i) => {
    const sel = document.querySelector(`input[name="cf${i}"]:checked`);
    if (!sel) { keep.push(c); return; }
    const m = sel.value === 'a' ? c.a : c.b;
    const cur = M[c.k];
    M[c.k] = { ...(cur && sameMark(cur, m) ? cur : m), t: Date.now(), res: 1 };
    dirty.add(vhOf(c.k));
  });
  CONFLICTS = keep;
  await saveConflicts(); await saveLocal(); schedulePush();
  closeDlg('dlgConf'); renderConflictNote();
  if (!$('work').hidden) renderAll();
});
$('syncOut').addEventListener('click', () => exportSync($('syncPh').checked));
$('syncBarBtn').addEventListener('click', () => { $('loadStatus').textContent = ''; renderHome(); $('syncCard').scrollIntoView({ block: 'start' }); });
setInterval(renderSyncBar, 5 * 60000);
$('syncIn').addEventListener('change', (e) => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) importSync(fs); });

// ---------- Diyaloglar ----------
let lastFocus = null;
function openDlg(id, focusId) { lastFocus = document.activeElement; $(id).hidden = false; setTimeout(() => { const el = $(focusId); if (el) el.focus(); }, 30); }
function closeDlg(id) { if (id === 'dlgEdit') stopDictation(); $(id).hidden = true; if (lastFocus && lastFocus.focus) lastFocus.focus(); }
document.querySelectorAll('.dlg').forEach((d) => {
  d.addEventListener('click', (e) => { if (e.target === d || e.target.closest('[data-close]')) closeDlg(d.id); });
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') document.querySelectorAll('.dlg:not([hidden])').forEach((d) => closeDlg(d.id)); });
let askFn = null;
function ask(title, text, okLabel, fn) { $('askTitle').textContent = title; $('askText').textContent = text; $('askOk').textContent = okLabel; askFn = fn; openDlg('dlgAsk', 'askOk'); }
$('askOk').addEventListener('click', async () => { closeDlg('dlgAsk'); const f = askFn; askFn = null; if (f) await f(); });

// ---------- Başlangıç ----------
$('saveBtn').addEventListener('click', saveNow);
(async () => {
  await loadAll();
  await loadPhotos();
  await loadPending();
  await loadConflicts();
  rebuild();
  resolvePending();
  if (DOCS.length) showWork(); else renderHome();
  renderSaveState();
  cloudInit();
  initDictation();
  renderGpsState();
  if (STANDALONE && navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
})();
