/* Planlı Üretim Arazi Tespit Tutanağı — ÇKS Excel'inden tutanak üretimi.
   Tarayıcıda ve Node'da aynı şekilde çalışır. */
(function (root) {
  'use strict';

  // ---------- Sayfa ölçüleri (TBS tutanağından ölçüldü, pt) ----------
  const W = 841.89, H = 595.276;
  const COLS = [14.17, 38.45, 83.36, 141.09, 168.39, 200.93, 232.22, 270.82, 311.68, 391.01, 482.93, 526.48, 558.35, 639.02, 732.95, 782.8];
  const LINE = 9.9;        // 9 pt yazıda satır aralığı
  const PAD_L = 2.1;       // hücre sol boşluğu
  const PAD_R = 1.9;       // sağa dayalı sayılar için
  const GRAY = '0.827 g';
  const TOP_FIRST = H - 271.21;  // ilk sayfada tablo gövdesinin üstü
  const TOP_NEXT = H - 28.6;     // devam sayfalarında
  const BOTTOM = 40;
  const colL = (no) => COLS[no - 7];
  const colR = (no) => COLS[no - 6];

  // ---------- Yazı tipi ölçüleri ----------
  function fontApi(fm) {
    const q = fm.map['?'];
    const gid = (ch) => (fm.map[ch] !== undefined ? fm.map[ch] : q);
    return {
      width(s, size) { let w = 0; for (const ch of String(s)) w += fm.widths[gid(ch)]; return (w * size) / 1000; },
      hex(s) { let h = ''; for (const ch of String(s)) h += gid(ch).toString(16).padStart(4, '0'); return h; }
    };
  }

  // TBS'deki gibi satır kırar: boşlukta, "\" işaretinden önce, sığmayan kelimede harf harf.
  function wrap(text, maxW, size, F) {
    const lines = [];
    let cur = '';
    for (const word of String(text || '').split(/\s+/).filter(Boolean)) {
      word.split(/(?=\\)/).forEach((piece, i) => {
        const cand = cur + (i === 0 && cur ? ' ' : '') + piece;
        if (F.width(cand, size) <= maxW) { cur = cand; return; }
        if (cur) { lines.push(cur); cur = ''; }
        let p = piece;
        while (F.width(p, size) > maxW && p.length > 1) {
          let n = 1;
          while (n < p.length && F.width(p.slice(0, n + 1), size) <= maxW) n++;
          lines.push(p.slice(0, n));
          p = p.slice(n);
        }
        cur = p;
      });
    }
    if (cur) lines.push(cur);
    return lines;
  }

  // ---------- Parsel anahtarı, köy ----------
  const nz = (s) => String(s || '').toLocaleUpperCase('tr-TR').replace(/\s+/g, '');
  function villageOf(r) {
    const p = String(r.yer || '').split('\\').map((x) => x.replace(/\s+/g, ' ').trim()).filter(Boolean);
    return p.length ? p[p.length - 1] : '—';
  }
  // TC + parselin köyü + ada + parsel + ürün; aynı çiftçide aynısından iki satır varsa alan da eklenir.
  function assignKeys(farmers) {
    for (const f of farmers) {
      const base = f.rows.map((r) => ['k3', nz(f.tc), nz(villageOf(r)), nz(r.ada), nz(r.parsel), nz(r.urun)].join('|'));
      const cnt = {};
      base.forEach((b) => { cnt[b] = (cnt[b] || 0) + 1; });
      const seen = {};
      f.rows.forEach((r, i) => {
        let k = cnt[base[i]] > 1 ? base[i] + '|' + nz(r.alan) : base[i];
        seen[k] = (seen[k] || 0) + 1;
        if (seen[k] > 1) k += '#' + seen[k];
        r.key = k;
      });
    }
    return farmers;
  }

  // ---------- ÇKS "Köy Genelinde Parsel Üretim Belgesi" (Excel) ----------
  function fromSheet(rows) {
    const Hd = (s) => String(s || '').toLocaleUpperCase('tr-TR').replace(/\s+/g, '');
    const want = {
      ad: 'İŞLETMEADI', tc: 'TC/VERGİNO', il: 'İL', ilce: 'İLÇE', koy: 'KÖY', ada: 'ADANO', parsel: 'PARSELNO',
      kul: 'KULLANILANALAN(DA)', tpn: 'TARIMPARSELNO', palan: 'PARSELALANI(DA)', urun: 'ÜRÜN', sekil: 'ÜRÜNKULLANIMŞEKLİ', tarim: 'TARIMŞEKLİ',
      ekili: 'EKİLİALAN(DA)', cesit: 'ÜRETİMÇEŞİDİ', dil: 'DİLEKÇEİLİ', dilce: 'DİLEKÇEİLÇESİ'
    };
    let hi = -1, year = '';
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const r = rows[i] || [];
      for (const c of r) { const m = String(c || '').match(/Üretim\s*Yılı\s*:?\s*(\d{4})/i); if (m) year = m[1]; }
      if (r.some((c) => Hd(c) === want.ad) && r.some((c) => Hd(c) === want.tc)) { hi = i; break; }
    }
    if (hi < 0) return { year, farmers: [] };
    const col = {};
    (rows[hi] || []).forEach((c, j) => { const h = Hd(c); for (const k in want) if (want[k] === h && col[k] === undefined) col[k] = j; });
    if ([col.tc, col.ada, col.parsel, col.urun, col.koy].includes(undefined)) return { year, farmers: [] };
    const g = (r, k) => (col[k] === undefined ? '' : String(r[col[k]] ?? '').replace(/\s+/g, ' ').trim());
    const fx = (v) => { const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? '' : n.toFixed(3); };
    const map = new Map();
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const tc = g(r, 'tc').replace(/\D/g, '');
      if (!tc) continue;
      if (!map.has(tc)) map.set(tc, { ad: g(r, 'ad'), tc, il: g(r, 'il'), ilce: g(r, 'ilce'), dil: g(r, 'dil'), dilce: g(r, 'dilce'), rows: [] });
      const f = map.get(tc);
      f.rows.push({
        urun: g(r, 'sekil') ? g(r, 'urun') + '\\' + g(r, 'sekil') : g(r, 'urun'),
        yer: [g(r, 'il'), g(r, 'ilce'), g(r, 'koy')].filter(Boolean).join('\\'),
        ada: g(r, 'ada'), parsel: g(r, 'parsel'), alan: fx(g(r, 'ekili') || g(r, 'kul')), palan: fx(g(r, 'palan')),
        tip: g(r, 'cesit').replace(/^(\d+)\.\s*/, '$1. '), sulama: g(r, 'tarim').toLocaleUpperCase('tr-TR'), tpn: g(r, 'tpn')
      });
    }
    return { year, farmers: [...map.values()] };
  }

  // ---------- Tespit değerleri ----------
  const DEFAULTS = { eksiUrun: 'ÇAYIR OTU\\MUHTELİF', eksiAcik: '', p1: '', p2: '', p3: '', tarihMod: 'bos', tarih: '' };
  const isCayir = (r) => /^ÇAYIR OTU/.test(String(r.urun || '').toLocaleUpperCase('tr-TR'));
  const AUTO = Object.freeze({ m: '+', auto: true });
  // Planlamaya uygunluk her zaman "Uygun"; olmama nedeni boş.
  // Arazide ölçüm yapıldıysa (mark.olc) tespit alanına (18) ölçülen alan yazılır.
  function tespit(row, mark, st) {
    if (!mark) return null;
    const v = mark.v || {};
    const olc = mark.olc || '';
    if (mark.m === 'e') return { urun: v.urun || '', alan: olc || v.alan || row.alan, acik: v.acik || '', olc };
    if (mark.m === '+') return { urun: row.urun, alan: olc || row.alan, acik: '', olc };
    if (mark.m === '-') return { urun: st.eksiUrun, alan: olc || row.alan, acik: st.eksiAcik || '', olc };
    return null;
  }
  const placeKey = (r) => [nz(villageOf(r)), nz(r.ada), nz(r.parsel)].join('|');
  // "10", "7,5", "7.5" -> "10.000", "7.500"; geçersizse ''
  function fmtArea(s) {
    const t = String(s || '').trim().replace(/\s+/g, '');
    if (!t) return '';
    const n = parseFloat(t.replace(',', '.'));
    return isNaN(n) || n <= 0 || n > 100000 ? '' : n.toFixed(3);
  }
  const trDate = (iso) => { const [y, m, d] = String(iso).split('-'); return d && m && y ? `${d}.${m}.${y}` : ''; };
  const isoToday = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };

  // ---------- PDF ----------
  function b64(s) {
    if (typeof atob === 'function') { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
    return new Uint8Array(Buffer.from(s, 'base64'));
  }
  function embedFont(PDFLib, doc, fm, name, flags) {
    const { PDFString } = PDFLib;
    const ctx = doc.context;
    const bytes = b64(fm.font);
    const ff = ctx.register(ctx.flateStream(bytes, { Length1: bytes.length }));
    const fd = ctx.register(ctx.obj({ Type: 'FontDescriptor', FontName: name, Flags: flags, FontBBox: fm.bbox, ItalicAngle: 0, Ascent: fm.ascent, Descent: fm.descent, CapHeight: fm.capHeight, StemV: 80, FontFile2: ff }));
    const cid = ctx.register(ctx.obj({
      Type: 'Font', Subtype: 'CIDFontType2', BaseFont: name,
      CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 },
      FontDescriptor: fd, DW: 1000, W: [0, fm.widths], CIDToGIDMap: 'Identity'
    }));
    const tu = ctx.register(ctx.flateStream(fm.toUnicode));
    return ctx.register(ctx.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: name, Encoding: 'Identity-H', DescendantFonts: [cid], ToUnicode: tu }));
  }
  const n2 = (v) => (Math.round(v * 100) / 100).toString();

  const HEAD = {
    7: ['Sıra', 'No', '(7)'], 8: ['Ürün Adı', '(8)'], 9: ['İl\\İlçe\\Köy', '(9)'], 10: ['Ada', 'No', '(10)'], 11: ['Parsel', 'No', '(11)'],
    12: ['Alan', '(da)', '(12)'], 13: ['Üretim', 'Tipi', '(13)'], 14: ['Sulama', 'Durumu', '(14)'], 15: ['Üretim', 'Planlamasına', 'Uygunluğu', '(15)'],
    16: ['Üretim Planlamasına', 'Uygun Olmama Nedeni', '(16)'], 17: ['Ürün Adı', '(17)'], 18: ['Alan', '(da)', '(18)'],
    19: ['Üretim Planlamasına', 'Uygunluğu', '(19)'], 20: ['Üretim Planlamasına', 'Uygun Olmama  Nedeni', '(20)'], 21: ['Açıklama', '(21)']
  };
  const FIELDS = [['İLİ', '(1)'], ['İLÇESİ', '(2)'], ['KÖYÜ / MAHALLESİ', '(3)'], ['ADI SOYADI / TİCARİ ÜNVANI', '(4)'], ['T.C. KİMLİK NO / VERGİ NO', '(5)'], ['CEP TELEFON NUMARASI', '(6)']];
  const FIELD_Y = [507.9, 485.5, 462.7, 439.8, 418.0, 395.1];

  // list: [{ f, rows, koy }] — rows tutanaktaki sırasıyla. markOf(row) -> işaret.
  async function buildPdf(PDFLib, list, markOf, st, FONTS, opts) {
    const { PDFDocument, PDFName, PDFHexString } = PDFLib;
    const o = opts || {};
    const doc = await PDFDocument.create();
    doc.setTitle('Planlı Üretim Arazi Tespit Tutanağı');
    doc.setCreator('Arazi Tespit');
    const api = { R: fontApi(FONTS.R), B: fontApi(FONTS.B), S: fontApi(FONTS.S) };
    const refs = {
      R: embedFont(PDFLib, doc, FONTS.R, 'TSPAAA+LiberationSerif', 34),
      B: embedFont(PDFLib, doc, FONTS.B, 'TSPAAB+LiberationSerif-Bold', 34),
      S: embedFont(PDFLib, doc, FONTS.S, 'TSPAAC+LiberationSans', 32)
    };
    const pages = [];
    let cur = null;
    const newPage = () => { cur = { ops: '0 g 0 G\n' }; pages.push(cur); };
    const T = (f, size, x, y, s) => { if (s !== '' && s != null) cur.ops += `BT /F${f} ${n2(size)} Tf ${n2(x)} ${n2(y)} Td <${api[f].hex(s)}> Tj ET\n`; };
    const TC = (f, size, x0, x1, y, s) => T(f, size, (x0 + x1) / 2 - api[f].width(s, size) / 2, y, s);
    const fill = (x0, y0, x1, y1) => { cur.ops += `${GRAY} ${n2(x0)} ${n2(y0)} ${n2(x1 - x0)} ${n2(y1 - y0)} re f 0 g\n`; };
    const box = (x0, y0, x1, y1, lw) => { cur.ops += `${lw} w ${n2(x0)} ${n2(y0)} ${n2(x1 - x0)} ${n2(y1 - y0)} re S\n`; };
    const year = o.year || String(new Date().getFullYear());

    const drawHead = (it) => {
      fill(14.17, H - 72.92, 782.8, H - 28.35); box(14.17, H - 72.92, 782.8, H - 28.35, 0.5);
      TC('B', 12, 14.17, 782.8, 547.2, `${year} YILI PLANLI ÜRETİM DESTEĞİ`);
      TC('B', 12, 14.17, 782.8, 533.9, 'PLANLI ÜRETİM ARAZİ TESPİT TUTANAĞI');
      const vals = [it.f.dil || it.f.il, it.f.dilce || it.f.ilce, it.koy || '', it.f.ad, it.f.tc, it.f.tel || ''];
      FIELDS.forEach(([lab, no], i) => {
        T('R', 12, 16.2, FIELD_Y[i], lab); T('B', 7, 211.5, FIELD_Y[i], no); T('R', 12, 234.6, FIELD_Y[i], ':');
        T('R', 12, 241.9, FIELD_Y[i], vals[i]);
      });
      const g0 = H - 225.94, g1 = H - 208.93;
      fill(14.17, g0, 782.8, g1);
      [[14.17, 38.45], [38.45, 482.93], [482.93, 782.8]].forEach(([a, b]) => box(a, g0, b, g1, 0.8));
      TC('R', 11, 38.45, 482.93, 373.7, 'BEYAN EDİLEN'); TC('R', 11, 482.93, 782.8, 373.7, 'TESPİT EDİLEN');
      const h0 = H - 271.21;
      fill(14.17, h0, 782.8, g0);
      for (let no = 7; no <= 21; no++) {
        box(colL(no), h0, colR(no), g0, 0.8);
        const ls = HEAD[no];
        ls.forEach((s, i) => TC('R', 9, colL(no), colR(no), 328 + (ls.length - 1 - i) * LINE, s));
      }
    };

    const layoutRow = (r, i, mk) => {
      const F = api.R;
      const mw = (no) => colR(no) - colL(no) - PAD_L - 1.0;
      const t = tespit(r, mk, st);
      const cells = {
        7: [String(i + 1)], 8: wrap(r.urun, mw(8), 9, F), 9: wrap(r.yer, mw(9), 9, F), 10: [r.ada], 11: [r.parsel], 12: [r.alan],
        13: wrap(r.tip, mw(13), 9, F), 14: wrap(r.sulama, mw(14), 9, F), 15: ['Uygun'], 16: [],
        17: t ? wrap(t.urun, mw(17), 9, F) : [], 18: t ? [t.alan] : [], 19: t ? ['Uygun'] : [], 20: [], 21: t ? wrap(t.acik, mw(21), 9, F) : []
      };
      let n = 1;
      for (const no in cells) n = Math.max(n, cells[no].length);
      return { cells, n };
    };
    const RIGHT = { 7: 1, 10: 1, 11: 1, 12: 1, 18: 1 };
    const drawRow = (L, top) => {
      const h = L.n * LINE + 4.15;
      for (let no = 7; no <= 21; no++) {
        box(colL(no), top - h, colR(no), top, 0.8);
        L.cells[no].forEach((s, k) => {
          const y = top - 10.0 - k * LINE;
          if (RIGHT[no]) T('R', 9, colR(no) - PAD_R - api.R.width(s, 9), y, s);
          else T('R', 9, colL(no) + PAD_L, y, s);
        });
      }
      return h;
    };
    const drawSig = (tb, it) => {
      const xs = [209.4, 404.6, 641.0];
      ['Teknik Personel', 'Teknik Personel', 'CBS Sorumlusu'].forEach((s, i) => T('S', 10, xs[i], tb - 17.8, s));
      const names = [st.p1, st.p2, st.p3].map((x) => String(x || '').trim());
      let tarih = '';
      if (st.tarihMod === 'sabit') tarih = String(st.tarih || '').trim();
      else if (st.tarihMod === 'son') {
        let best = '';
        it.rows.forEach((r) => { const m = markOf(r); if (m && !m.auto && m.d && m.d > best) best = m.d; });
        tarih = trDate(best || isoToday());
      }
      [['Adı Soyadı', names], ['Tarih', [tarih, tarih, tarih]], ['İmza', ['', '', '']]].forEach(([lab, v], j) => {
        const y = tb - 38.3 - j * 16.5;
        T('S', 10, 87.9, y, lab); T('S', 10, 209.4, y, ':');
        v.forEach((s, i) => { if (s) T(i === 0 ? 'S' : 'S', 10, i === 0 ? 215.2 : xs[i], y, s); });
      });
    };

    for (const it of list) {
      newPage();
      drawHead(it);
      let top = TOP_FIRST;
      it.rows.forEach((r, i) => {
        const L = layoutRow(r, i, markOf(r));
        const h = L.n * LINE + 4.15;
        if (top - h < BOTTOM) { newPage(); top = TOP_NEXT; }
        top -= drawRow(L, top);
      });
      if (top - 80 < 28) { newPage(); top = TOP_NEXT; }
      drawSig(top, it);
    }

    const ctx = doc.context;
    const N = pages.length;
    pages.forEach((p, i) => {
      const s = `${i + 1} / ${N}`;
      cur = p; TC('S', 8, 384.5, 404.5, 19.1, s);
      const page = doc.addPage([W, H]);
      const ref = ctx.register(ctx.flateStream(p.ops));
      page.node.set(PDFName.of('Contents'), ref);
      page.node.setFontDictionary(PDFName.of('FR'), refs.R);
      page.node.setFontDictionary(PDFName.of('FB'), refs.B);
      page.node.setFontDictionary(PDFName.of('FS'), refs.S);
    });
    if (o.backup) doc.getInfoDict().set(PDFName.of('TespitData'), PDFHexString.fromText(JSON.stringify(o.backup)));
    return doc.save();
  }

  // Bu uygulamadan indirilmiş bir tutanağın içine saklanan işaretleri okur.
  async function readBackup(PDFLib, bytes) {
    const { PDFDocument, PDFName } = PDFLib;
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    const td = doc.getInfoDict().lookup(PDFName.of('TespitData'));
    if (!td) return null;
    try { return JSON.parse(td.decodeText()); } catch (e) { return null; }
  }

  // ---------- Fotoğraf ----------
  // JPEG içindeki EXIF'ten çekim zamanı ve GPS koordinatını okur (kamera konum etiketi açıksa).
  function readExif(buf) {
    try {
      const dv = new DataView(buf);
      if (dv.getUint16(0) !== 0xFFD8) return {};
      let off = 2;
      while (off + 4 < dv.byteLength) {
        const marker = dv.getUint16(off);
        if ((marker & 0xFF00) !== 0xFF00) break;
        const len = dv.getUint16(off + 2);
        if (marker === 0xFFE1 && dv.getUint32(off + 4) === 0x45786966) return parseTiff(dv, off + 10);
        if (marker === 0xFFDA) break;
        off += 2 + len;
      }
    } catch (e) {}
    return {};
  }
  function parseTiff(dv, t) {
    const le = dv.getUint16(t) === 0x4949;
    const u16 = (o) => dv.getUint16(o, le), u32 = (o) => dv.getUint32(o, le);
    const SZ = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8 };
    const ifd = (o) => { const n = u16(o); const e = {}; for (let i = 0; i < n; i++) { const p = o + 2 + i * 12; e[u16(p)] = { type: u16(p + 2), count: u32(p + 4), vo: p + 8 }; } return e; };
    const at = (en) => ((SZ[en.type] || 1) * en.count > 4 ? t + u32(en.vo) : en.vo);
    const rat = (o) => u32(o) / (u32(o + 4) || 1);
    const str = (en) => { const o = at(en); let s = ''; for (let i = 0; i < en.count; i++) { const c = dv.getUint8(o + i); if (!c) break; s += String.fromCharCode(c); } return s; };
    const out = {};
    const ifd0 = ifd(t + u32(t + 4));
    if (ifd0[0x8769]) { const ex = ifd(t + u32(ifd0[0x8769].vo)); if (ex[0x9003]) out.dt = str(ex[0x9003]); }
    if (ifd0[0x8825]) {
      const g = ifd(t + u32(ifd0[0x8825].vo));
      const coord = (tag, ref) => {
        if (!g[tag] || !g[ref] || g[tag].count < 3) return null;
        const o = at(g[tag]);
        const v = rat(o) + rat(o + 8) / 60 + rat(o + 16) / 3600;
        const r = str(g[ref]);
        return r === 'S' || r === 'W' ? -v : v;
      };
      const lat = coord(2, 1), lon = coord(4, 3);
      if (lat != null && lon != null && isFinite(lat) && isFinite(lon) && (lat || lon)) { out.lat = lat; out.lon = lon; }
    }
    return out;
  }

  // Fotoğraf eki: her fotoğraf ayrı yatay A4 sayfada, üstte çiftçi ve parsel bilgisi.
  async function buildPhotoPdf(PDFLib, items, FONTS, title) {
    const { PDFDocument, PDFName } = PDFLib;
    const doc = await PDFDocument.create();
    doc.setTitle(title || 'Fotoğraf eki');
    doc.setCreator('Arazi Tespit');
    const R = fontApi(FONTS.R), B = fontApi(FONTS.B);
    const refR = embedFont(PDFLib, doc, FONTS.R, 'TSPAAA+LiberationSerif', 34);
    const refB = embedFont(PDFLib, doc, FONTS.B, 'TSPAAB+LiberationSerif-Bold', 34);
    const ctx = doc.context;
    const N = items.length;
    for (let i = 0; i < N; i++) {
      const it = items[i];
      const img = await doc.embedJpg(it.jpg);
      const page = doc.addPage([W, H]);
      const top = H - 30;
      let ops = '0 g\n';
      const T = (api, f, size, x, y, s) => { ops += `BT /${f} ${n2(size)} Tf ${n2(x)} ${n2(y)} Td <${api.hex(s)}> Tj ET\n`; };
      T(B, 'FB', 13, 30, top - 12, it.title);
      T(R, 'FR', 10.5, 30, top - 28, it.sub);
      const pn = `${i + 1} / ${N}`;
      T(R, 'FR', 9, W - 30 - R.width(pn, 9), 18, pn);
      const boxW = W - 60, boxH = top - 42 - 34;
      const sc = Math.min(boxW / img.width, boxH / img.height);
      const w = img.width * sc, h = img.height * sc;
      const x = (W - w) / 2, y = 34 + (boxH - h) / 2;
      const name = PDFName.of('Im' + i);
      page.node.normalize();
      page.node.setXObject(name, img.ref);
      ops += `q ${n2(w)} 0 0 ${n2(h)} ${n2(x)} ${n2(y)} cm /Im${i} Do Q\n`;
      const ref = ctx.register(ctx.flateStream(ops));
      page.node.addContentStream(ref);
      page.node.setFontDictionary(PDFName.of('FR'), refR);
      page.node.setFontDictionary(PDFName.of('FB'), refB);
    }
    return doc.save();
  }

  // ---------- Excel (.xlsx) çıktısı ----------
  // Her çiftçi ayrı sayfa; düzen PDF tutanağıyla aynı (başlık, alanlar, 7-21 sütunları, imza bölümü).
  const xesc = (v) => String(v == null ? '' : v).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  const colName = (i) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  function sheetName(name, used) {
    let b = String(name || 'Çiftçi').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 28) || 'Çiftçi';
    let n = b, i = 2;
    while (used.has(n.toLocaleUpperCase('tr-TR'))) { n = b.slice(0, 27) + ' ' + i; i++; }
    used.add(n.toLocaleUpperCase('tr-TR'));
    return n;
  }
  async function buildXlsx(JSZip, list, markOf, st, FONTS, opts) {
    const o = opts || {};
    const F = fontApi(FONTS.R);
    const year = o.year || String(new Date().getFullYear());
    // Excel'de sütunlar biraz daha geniş: ürün ve köy adları 1-2 satıra sığsın, sayfa yazdırırken genişliğe sığdırılır.
    const colW = [5, 17, 21, 6.5, 7, 8, 9, 8, 12, 14, 17, 8, 12, 14, 16];
    const colPt = colW.map((w) => (w * 7 + 5) * 0.75);
    const HEADS = []; for (let no = 7; no <= 21; no++) HEADS.push(HEAD[no].join('\n').replace('Uygun Olmama  Nedeni', 'Uygun Olmama Nedeni'));
    // stil sırası styles.xml'deki cellXfs ile aynı
    const S_TITLE = 1, S_LABEL = 2, S_GROUP = 3, S_HEAD = 4, S_TXT = 5, S_NUM = 6, S_RIGHT = 7, S_SIG = 8, S_GRAYCELL = 9;
    const sheets = [];
    const used = new Set();
    for (const it of list) {
      const rows = [];      // { r, ht, cells: [[col, style, value, isNum]] }
      const merges = [];
      const put = (r, cells, ht) => rows.push({ r, ht, cells });
      put(1, [[0, S_TITLE, `${year} YILI PLANLI ÜRETİM DESTEĞİ`]], 18); merges.push('A1:O1');
      put(2, [[0, S_TITLE, 'PLANLI ÜRETİM ARAZİ TESPİT TUTANAĞI']], 18); merges.push('A2:O2');
      for (let c = 1; c < 15; c++) { rows[0].cells.push([c, S_TITLE, '']); rows[1].cells.push([c, S_TITLE, '']); }
      const vals = [it.f.dil || it.f.il, it.f.dilce || it.f.ilce, it.koy || '', it.f.ad, it.f.tc, it.f.tel || ''];
      FIELDS.forEach(([lab, no], i) => {
        const r = 4 + i;
        put(r, [[0, S_LABEL, `${lab} ${no} :`], [6, S_LABEL, vals[i]]], 17);
        merges.push(`A${r}:F${r}`, `G${r}:O${r}`);
      });
      const g = 11;
      put(g, [[0, S_GRAYCELL, ''], [1, S_GROUP, 'BEYAN EDİLEN'], ...[2, 3, 4, 5, 6, 7, 8, 9].map((c) => [c, S_GROUP, '']), [10, S_GROUP, 'TESPİT EDİLEN'], ...[11, 12, 13, 14].map((c) => [c, S_GROUP, ''])], 16);
      merges.push(`B${g}:J${g}`, `K${g}:O${g}`);
      put(g + 1, HEADS.map((h, i) => [i, S_HEAD, h]), 62);
      let r = g + 2;
      it.rows.forEach((row, i) => {
        const t = tespit(row, markOf(row), st);
        const mw = (no) => colPt[no - 7] - 8;
        const lines = Math.max(1, wrap(row.urun, mw(8), 9, F).length, wrap(row.yer, mw(9), 9, F).length, t ? wrap(t.urun, mw(17), 9, F).length : 1, t ? wrap(t.acik, mw(21), 9, F).length : 1);
        const num = (v) => { const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? null : n; };
        const cells = [
          [0, S_RIGHT, i + 1, true], [1, S_TXT, row.urun], [2, S_TXT, row.yer], [3, S_RIGHT, row.ada], [4, S_RIGHT, row.parsel],
          num(row.alan) != null ? [5, S_NUM, num(row.alan), true] : [5, S_RIGHT, row.alan], [6, S_TXT, row.tip], [7, S_TXT, row.sulama], [8, S_TXT, 'Uygun'], [9, S_TXT, ''],
          [10, S_TXT, t ? t.urun : ''], t && num(t.alan) != null ? [11, S_NUM, num(t.alan), true] : [11, S_RIGHT, t ? t.alan : ''], [12, S_TXT, t ? 'Uygun' : ''], [13, S_TXT, ''], [14, S_TXT, t ? t.acik : '']
        ];
        put(r, cells, Math.round(lines * 12 + 4));
        r++;
      });
      // imza bölümü
      r += 1;
      let tarih = '';
      if (st.tarihMod === 'sabit') tarih = String(st.tarih || '').trim();
      else if (st.tarihMod === 'son') {
        let best = '';
        it.rows.forEach((row) => { const m = markOf(row); if (m && !m.auto && m.d && m.d > best) best = m.d; });
        tarih = trDate(best || isoToday());
      }
      const names = [st.p1, st.p2, st.p3].map((x) => String(x || '').trim());
      put(r, [[6, S_SIG, 'Teknik Personel'], [9, S_SIG, 'Teknik Personel'], [12, S_SIG, 'CBS Sorumlusu']], 16); r++;
      [['Adı Soyadı :', names], ['Tarih :', [tarih, tarih, tarih]], ['İmza :', ['', '', '']]].forEach(([lab, v]) => {
        put(r, [[2, S_SIG, lab], [6, S_SIG, v[0]], [9, S_SIG, v[1]], [12, S_SIG, v[2]]], lab === 'İmza :' ? 30 : 18);
        r++;
      });
      const xmlRows = rows.map((row) => `<row r="${row.r}" ht="${row.ht}" customHeight="1">${row.cells.sort((a, b) => a[0] - b[0]).map(([c, sIdx, v, isNum]) => {
        const ref = colName(c) + row.r;
        if (isNum) return `<c r="${ref}" s="${sIdx}"><v>${v}</v></c>`;
        if (v === '' || v == null) return `<c r="${ref}" s="${sIdx}"/>`;
        return `<c r="${ref}" s="${sIdx}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
      }).join('')}</row>`).join('');
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" zoomScale="110"/></sheetViews>
<sheetFormatPr defaultRowHeight="13"/>
<cols>${colW.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>
<sheetData>${xmlRows}</sheetData>
<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>
<printOptions horizontalCentered="1"/>
<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter>
</worksheet>`;
      sheets.push({ name: sheetName(it.f.ad, used), xml, printRows: `$${g + 1}:$${g + 1}` });
    }
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="5">
<font><sz val="9"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="12"/><name val="Times New Roman"/><family val="1"/></font>
<font><sz val="12"/><name val="Times New Roman"/><family val="1"/></font>
<font><sz val="11"/><name val="Times New Roman"/><family val="1"/></font>
<font><sz val="10"/><name val="Arial"/><family val="2"/></font>
</fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD3D3D3"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="10">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="bottom" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Planlı Üretim Arazi Tespit Tutanağı</dc:title><dc:creator>Arazi Tespit</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((sh, i) => `<sheet name="${xesc(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    zip.file('xl/styles.xml', styles);
    sheets.forEach((sh, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sh.xml));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // Basit tablo halinde Excel (başlık satırı + satırlar). cols: [{ h, w, num }]
  async function buildTableXlsx(JSZip, sheet, title, cols, rows) {
    const cell = (c, r, v, st, num) => {
      const ref = colName(c) + r;
      if (v === '' || v == null) return `<c r="${ref}" s="${st}"/>`;
      if (num && typeof v === 'number') return `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
    };
    let x = `<row r="1" ht="20" customHeight="1">${cell(0, 1, title, 1)}</row>`;
    x += `<row r="2" ht="30" customHeight="1">${cols.map((c, i) => cell(i, 2, c.h, 2)).join('')}</row>`;
    rows.forEach((row, ri) => { x += `<row r="${ri + 3}">${row.map((v, i) => cell(i, ri + 3, v, cols[i].num ? 4 : 3, cols[i].num)).join('')}</row>`; });
    const last = colName(cols.length - 1);
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.w}" customWidth="1"/>`).join('')}</cols>
<sheetData>${x}</sheetData>
<autoFilter ref="A2:${last}${rows.length + 2}"/>
<mergeCells count="1"><mergeCell ref="A1:${last}1"/></mergeCells>
<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="3"><font><sz val="10"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="12"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><name val="Arial"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1E8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="${xesc(String(sheet).slice(0, 31))}" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'${xesc(String(sheet).slice(0, 31))}'!$A$2:$${last}$${rows.length + 2}</definedName></definedNames></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    zip.file('xl/styles.xml', styles);
    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  root.TespitCore = { W, H, COLS, wrap, fontApi, villageOf, assignKeys, fromSheet, DEFAULTS, isCayir, AUTO, tespit, placeKey, fmtArea, trDate, isoToday, buildPdf, readBackup, readExif, buildPhotoPdf, buildXlsx, buildTableXlsx };
})(typeof globalThis !== 'undefined' ? globalThis : this);
