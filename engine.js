/* Motor de cálculo: replica las fórmulas de la pestaña "Data" del Excel
   Cobertura Inventarios Villavicencio a partir de ZEXI_01, TMS y Preventa. */
(function (root) {
  const COLUMNS = ['Llave','Semana','Tipo PT','Marca','Lines','Centro','Material','Descripcion','UM','Existencia','Cant entrega Pendien','Total Disponible','Cantidad en Pedidos','Inv. Final','tipo','presentacio','conte','Inv. Piso','Inv. Reservado','Venta Total','Preventa Hl','Venta tat','Venta Ka','Inv. Transito','Stock out','FC Dia','Dias cobertura','Stock Seguridad HL','Stock Seguridad DC','TMS','Fecha Arribo TMS','Fecha Prox Prod','Planta','Htls Prod Plan','Riesgo HL','Riesgo DC','FC TOTAL CD','FC MES','Volumen Entregado','Forescast KA','Forescaste TAT','FC MTD KA','Volumen Solicitado','%Cumpl. Ventas','Volumen recibido2','%Cumpl. T1','Venta W+1','Cumpl. CD','Cumpl. Ventas','Entregado TAT','Entregado KA','Entregado Socios','DC TMS','DC BODEGA','DC PREVENTA','DC TRANSITOS','VENTA D+3','PRD 1.5 EST','PRD 1,5 PQ'];

  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const num = v => { if (v == null || v === '') return 0; if (typeof v === 'number') return v; const n = parseFloat(String(v).replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; };
  const key = v => { if (v == null) return ''; const s = String(v).trim(); return /^\d+(\.0+)?$/.test(s) ? String(parseInt(s, 10)) : s.toUpperCase(); };
  const div = (a, b) => (b ? a / b : 0); // IFERROR(a/b,0)

  function findHeader(rows, mustHave, maxScan = 15) {
    for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
      const h = (rows[i] || []).map(norm);
      if (mustHave.every(m => h.some(x => x === m || x.startsWith(m)))) return i;
    }
    return -1;
  }
  function colIdx(hdr, ...cands) {
    const h = hdr.map(norm);
    for (const c of cands) { const i = h.findIndex(x => x === c); if (i >= 0) return i; }
    for (const c of cands) { const i = h.findIndex(x => x.startsWith(c)); if (i >= 0) return i; }
    return -1;
  }

  /* ---------- ZEXI_01 ---------- */
  function parseZexi(rows) {
    const hi = findHeader(rows, ['material', 'alm']);
    if (hi < 0) throw new Error('ZEXI_01: no encontré los encabezados "Material" y "Alm."');
    const h = rows[hi];
    const c = {
      mat: colIdx(h, 'material'), ce: colIdx(h, 'ce.', 'centro', 'ce'), alm: colIdx(h, 'alm.', 'almacen', 'alm'),
      libre: colIdx(h, 'libre utiliz.', 'libre utilizacion', 'libre'), ctrl: colIdx(h, 'in control-cal', 'en control calidad', 'control'),
      bloq: colIdx(h, 'bloqueado')
    };
    const miss = Object.entries(c).filter(([k, v]) => v < 0 && k !== 'ce').map(([k]) => k);
    if (miss.length) throw new Error('ZEXI_01: faltan columnas (' + miss.join(', ') + ')');
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r || r[c.mat] == null || r[c.mat] === '') continue;
      out.push({ mat: key(r[c.mat]), ce: c.ce >= 0 ? key(r[c.ce]) : '', alm: key(r[c.alm]), libre: num(r[c.libre]), ctrl: num(r[c.ctrl]), bloq: num(r[c.bloq]) });
    }
    if (!out.length) throw new Error('ZEXI_01: el archivo no tiene filas de datos');
    return out;
  }

  /* ---------- TMS ---------- */
  function parseTms(rows) {
    const hi = findHeader(rows, ['llave']) >= 0 ? findHeader(rows, ['llave']) : findHeader(rows, ['sku']);
    if (hi < 0) throw new Error('TMS: no encontré la columna "LLAVE" o "SKU"');
    const h = rows[hi];
    const cL = colIdx(h, 'llave'), cSku = colIdx(h, 'sku'), cDest = colIdx(h, 'shipment destination location id', 'destino');
    const cHl = colIdx(h, 'hl', 'suma de hl', 'hectolitros'), cF = colIdx(h, 'fecha arribo');
    if (cHl < 0) throw new Error('TMS: falta la columna "HL"');
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      let k = cL >= 0 ? r[cL] : null;
      if ((k == null || k === '') && cSku >= 0 && cDest >= 0 && r[cSku] != null) k = key(r[cSku]) + key(r[cDest]);
      if (k == null || k === '' || /total|en blanco/i.test(String(k))) continue;
      out.push({ llave: String(k).trim().toUpperCase(), hl: num(r[cHl]), fecha: cF >= 0 ? r[cF] : null });
    }
    return out; // puede estar vacío (semana sin TMS)
  }

  /* ---------- Preventa ---------- */
  function parsePreventa(rows) {
    // Busca los bloques "PREVENTA D+1" y "PREVENTA D+3" (SKU, CO GC Llanos, CO KA OFF Sur, Total)
    const blocks = {};
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      (rows[i] || []).forEach((v, j) => { const n = norm(v).replace(/\s/g, ''); if (n.includes('d+1')) blocks.d1 = blocks.d1 || { r: i, c: j }; if (n.includes('d+3')) blocks.d3 = blocks.d3 || { r: i, c: j }; });
    }
    if (!blocks.d1 && !blocks.d3) throw new Error('Preventa: no encontré los títulos "PREVENTA D+1" / "PREVENTA D+3"');
    const read = b => {
      const m = {}; if (!b) return m;
      const hdr = rows[b.r] || [];
      // columna Total: la que dice "total" a la derecha del título; si no, la 4ta
      let cT = b.c + 3; for (let j = b.c + 1; j <= b.c + 6; j++) if (norm(hdr[j]) === 'total') { cT = j; break; }
      let cTat = b.c + 1, cKa = b.c + 2;
      for (let j = b.c + 1; j < cT; j++) { const n = norm(hdr[j]); if (n.includes('llanos') || n.includes('gc')) cTat = j; if (n.includes('ka')) cKa = j; }
      for (let i = b.r + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue; const s = r[b.c];
        if (s == null || s === '' || isNaN(num(s)) || norm(s) === 'sku' || norm(s).includes('total')) continue;
        const k = key(s); if (!(k in m)) m[k] = { tat: num(r[cTat]), ka: num(r[cKa]), total: num(r[cT]) };
      }
      return m;
    };
    return { d1: read(blocks.d1), d3: read(blocks.d3) };
  }

  /* Excel WEEKNUM(fecha) tipo 1 (semana inicia domingo) */
  function weeknum(d) {
    const j1 = new Date(d.getFullYear(), 0, 1);
    const days = Math.floor((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - j1) / 864e5);
    return Math.floor((days + j1.getDay()) / 7) + 1;
  }
  function toDateStr(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) { const d = new Date(Math.round(v.getTime() / 36e5) * 36e5); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
    return String(v);
  }

  /* ---------- Cálculo de Data ---------- */
  function compute(master, zexiRows, tmsRows, prev, today = new Date()) {
    const day = today.getDate();
    const BJ = day / 31 > 1 ? 1 : day / 30;
    const semana = weeknum(today);
    const conteOf = {}; master.forEach(m => conteOf[key(m.Material)] = num(m.conte));
    // índices ZEXI
    const zByKey = {}, zFirstByMat = {};
    zexiRows.forEach(z => {
      const k = z.mat + z.ce + z.alm; if (!(k in zByKey)) zByKey[k] = z;
      if (!(z.mat in zFirstByMat)) zFirstByMat[z.mat] = z;
    });
    // TMS: suma de HL y primera fecha de arribo por llave
    const tmsHl = {}, tmsF = {};
    tmsRows.forEach(t => { tmsHl[t.llave] = (tmsHl[t.llave] || 0) + t.hl; const f = toDateStr(t.fecha); if (f && (!tmsF[t.llave] || f < tmsF[t.llave])) tmsF[t.llave] = f; });

    const rows = master.map(m => {
      const mat = key(m.Material), centro = key(m.Centro || 'AV19'), llave = mat + centro;
      const Q = num(m.conte), P = num(m.presentacio);
      const zA = zByKey[llave + '1901'];
      const J = zA ? zA.libre : 0;
      const K = num(m['Cant entrega Pendien']), M = num(m['Cantidad en Pedidos']);
      const L = J - K, N = L - M;
      const R = div(J * Q, 1e5), S = div(K * Q, 1e5), T = div(M * P * Q, 1e5);
      const U = N < 0 ? (M + N) * P * Q / 1e5 : M * P * Q / 1e5;
      const p1 = prev.d1[mat] || {}, p3 = prev.d3[mat] || {};
      const V = p1.tat || 0, W = p1.ka || 0, AV = p1.total || 0, BG = p3.total || 0;
      const zT = zByKey[llave + '1940'];
      const X = zT ? (zT.libre + zT.ctrl + zT.bloq) * Q / 1e5 : 0;
      const AL = num(m['FC TOTAL CD']), AM = AL * BJ;
      const Z = div(AM, day);
      const zF = zFirstByMat[mat];
      const Y = zF ? div(zF.ctrl * Q / 1e5, Z) : 0;
      const AA = div(R, Z);
      const AD = tmsHl[llave] || 0, AE = tmsF[llave] || null;
      const AI = R + X + AD - S - T - AV;
      const AJ = Z === 0 ? null : AI / Z;
      const AO = num(m['Forescast KA']), AR = num(m['Volumen Solicitado']), AT = num(m['Volumen recibido2']);
      const AY = num(m['Entregado TAT']), AZ = num(m['Entregado KA']), BB = num(m['Entregado Socios']);
      const AN = AY + AZ + BB;
      const und = num(m._undEst), cj = num(m._cjEst);
      const BH = (Q && und) ? Math.round((R - Z * 1.5) * 100 / (Q / 1000)) / und : 0;
      const v = {
        'Llave': llave, 'Semana': semana, 'Tipo PT': m['Tipo PT'], 'Marca': m.Marca, 'Lines': m.Lines, 'Centro': m.Centro, 'Material': m.Material,
        'Descripcion': m.Descripcion, 'UM': m.UM, 'Existencia': J, 'Cant entrega Pendien': K, 'Total Disponible': L, 'Cantidad en Pedidos': M,
        'Inv. Final': N, 'tipo': m.tipo, 'presentacio': P, 'conte': Q, 'Inv. Piso': R, 'Inv. Reservado': S, 'Venta Total': T, 'Preventa Hl': U,
        'Venta tat': V, 'Venta Ka': W, 'Inv. Transito': X, 'Stock out': Y, 'FC Dia': Z, 'Dias cobertura': AA, 'Stock Seguridad HL': 0,
        'Stock Seguridad DC': 0, 'TMS': AD, 'Fecha Arribo TMS': AE, 'Fecha Prox Prod': m['Fecha Prox Prod'], 'Planta': m.Planta || null, 'Htls Prod Plan': m['Htls Prod Plan'],
        'Riesgo HL': AI, 'Riesgo DC': AJ, 'FC TOTAL CD': AL, 'FC MES': AM, 'Volumen Entregado': AN, 'Forescast KA': AO,
        'Forescaste TAT': Math.max(0, AM - AO), 'FC MTD KA': AO * BJ, 'Volumen Solicitado': AR, '%Cumpl. Ventas': div(AR, AM),
        'Volumen recibido2': AT, '%Cumpl. T1': div(AT, AM), 'Venta W+1': AV, 'Cumpl. CD': div(AN, AM), 'Cumpl. Ventas': div(AR, AM),
        'Entregado TAT': AY, 'Entregado KA': AZ, 'Entregado Socios': BB, 'DC TMS': div(AD, Z), 'DC BODEGA': div(R, Z), 'DC PREVENTA': div(AV, Z),
        'DC TRANSITOS': div(X, Z), 'VENTA D+3': BG, 'PRD 1.5 EST': BH, 'PRD 1,5 PQ': BH * cj
      };
      return COLUMNS.map(c => { const x = v[c]; return typeof x === 'number' ? Math.round(x * 100) / 100 : (x == null ? null : x); });
    });
    const matched = master.filter(m => zByKey[key(m.Material) + key(m.Centro || 'AV19') + '1901']).length;
    return { columns: COLUMNS, rows, semana, centro: master[0] ? master[0].Centro : 'AV19', stats: { zexi: zexiRows.length, tms: tmsRows.length, prevD1: Object.keys(prev.d1).length, prevD3: Object.keys(prev.d3).length, skusConZexi: matched } };
  }

  /* Elige la hoja correcta dentro de un libro (por nombre, o la primera) */
  function pickSheet(wb, XLSX, names) {
    const n = wb.SheetNames.find(s => names.some(x => norm(s) === norm(x))) || wb.SheetNames.find(s => names.some(x => norm(s).includes(norm(x)))) || wb.SheetNames[0];
    return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) };
  }

  const api = { COLUMNS, parseZexi, parseTms, parsePreventa, compute, pickSheet, weeknum };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Engine = api;
})(typeof self !== 'undefined' ? self : this);
