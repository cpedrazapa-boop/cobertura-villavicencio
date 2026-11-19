/* Motor de cálculo: replica las fórmulas de la pestaña "Data" del Excel
   Cobertura Inventarios Villavicencio a partir de ZEXI_01, TMS y Preventa. */
(function (root) {
  const COLUMNS = ['Llave','Semana','Tipo PT','Marca','Lines','Centro','Material','Descripcion','UM','Existencia','Cant entrega Pendien','Total Disponible','Cantidad en Pedidos','Inv. Final','tipo','presentacio','conte','Inv. Piso','Inv. Reservado','Venta Total','Preventa Hl','Venta tat','Venta Ka','Inv. Transito','Stock out','FC Dia','Dias cobertura','Stock Seguridad HL','Stock Seguridad DC','TMS','Fecha Arribo TMS','Fecha Prox Prod','Planta','Htls Prod Plan','Riesgo HL','Riesgo DC','FC TOTAL CD','FC MES','Volumen Entregado','Forescast KA','Forescaste TAT','FC MTD KA','Volumen Solicitado','%Cumpl. Ventas','Volumen recibido2','%Cumpl. T1','Venta W+1','Cumpl. CD','Cumpl. Ventas','Entregado TAT','Entregado KA','Entregado Socios','DC TMS','DC BODEGA','DC PREVENTA','DC TRANSITOS','VENTA D+3','PRD 1.5 EST','PRD 1,5 PQ','Inv. Calidad','Inv. Bloqueo','Inv. Calle','ABC'];

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
  function parseTms(rows, centros) {
    // Formato crudo (hoja "Data TMS" del Shipment Plan, o Libro1): SKU + Shipment Destination Location ID + HL
    const hr = findHeader(rows, ['shipment destination location id', 'sku', 'hl']);
    if (hr >= 0) {
      const h = rows[hr];
      const cSku = colIdx(h, 'sku'), cDest = colIdx(h, 'shipment destination location id'), cHl = colIdx(h, 'hl'),
            cF = colIdx(h, 'fecha arribo'), cCom = colIdx(h, 'shipment commodity');
      const out = [];
      for (let i = hr + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const dest = key(r[cDest]); if (!dest || (centros && !centros.includes(dest))) continue;
        if (cCom >= 0 && r[cCom] != null && norm(r[cCom]) !== 'fgoods') continue; // solo producto terminado
        const sku = key(r[cSku]); if (!/^\d+$/.test(sku)) continue;
        out.push({ llave: sku + dest, hl: num(r[cHl]), fecha: cF >= 0 ? r[cF] : null });
      }
      return out;
    }
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
  function parsePreventa(rows, today = new Date()) {
    // Formato crudo (CSV de surtido): Día, CD, CodCliente, TipoPedido, Sku, Suma de Hectolitros
    const hr = findHeader(rows, ['cd', 'sku', 'tipopedido']);
    if (hr >= 0) {
      const h = rows[hr];
      const cDia = colIdx(h, 'dia'), cCd = colIdx(h, 'cd'), cCli = colIdx(h, 'codcliente', 'cliente'), cTp = colIdx(h, 'tipopedido'),
            cSku = colIdx(h, 'sku'), cHl = colIdx(h, 'suma de hectolitros', 'hectolitros', 'hl'), cDoc = colIdx(h, 'documentoventa');
      if (cHl < 0) throw new Error('Preventa: falta la columna de hectolitros');
      const t = today.getDate(), dist = d => ((d - t) % 31 + 31) % 31;
      const data = [];
      for (let i = hr + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const sku = key(r[cSku]), cd = key(r[cCd]); if (!/^\d+$/.test(sku) || !cd) continue;
        data.push({ cd, sku, dia: num(r[cDia]), tp: key(r[cTp]), cli: r[cCli] == null ? '' : String(r[cCli]).trim(), doc: cDoc >= 0 ? r[cDoc] : null, hl: num(r[cHl]) });
      }
      if (!data.length) throw new Error('Preventa: el archivo no tiene filas de datos');
      const out = {};
      for (const cd of [...new Set(data.map(x => x.cd))]) {
        const rs = data.filter(x => x.cd === cd);
        const dias = [...new Set(rs.map(x => x.dia))].filter(d => dist(d) > 0).sort((a, b) => dist(a) - dist(b));
        const d1day = dias[0];
        const add = (m, x) => { const o = m[x.sku] || (m[x.sku] = { tat: 0, ka: 0, total: 0 }); if (x.tp === 'ZCP1') o.tat += x.hl; else if (x.tp === 'ZCP6') o.ka += x.hl; o.total += x.hl; };
        const d1 = {}, d3 = {}, cl = {};
        rs.forEach(x => {
          add(d3, x); if (x.dia === d1day) add(d1, x);
          const c = (cl[x.sku] || (cl[x.sku] = {})), k = x.cli;
          const o = c[k] || (c[k] = { hl: 0, hl1: 0, docs: new Set(), tp: x.tp }); o.hl += x.hl; if (x.dia === d1day) o.hl1 += x.hl; if (x.doc != null) o.docs.add(x.doc);
        });
        const clientes = {};
        for (const sku in cl) clientes[sku] = Object.entries(cl[sku]).map(([c, o]) => [c, Math.round(o.hl * 100) / 100, o.docs.size, o.tp, '', Math.round(o.hl1 * 100) / 100]).sort((a, b) => b[1] - a[1]);
        out[cd] = { d1, d3, clientes, d1day, dias: [...new Set(rs.map(x => x.dia))].sort((a, b) => dist(a) - dist(b)) };
      }
      return out;
    }
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
    return { '*': { d1: read(blocks.d1), d3: read(blocks.d3) } };
  }

  /* ---------- SL (solicitado / entregado) ---------- */
  function filtroCentros(rows) {
    for (let i = rows.length - 1; i >= Math.max(0, rows.length - 6); i--) {
      const t = (rows[i] || []).filter(v => typeof v === 'string' && /filtros aplicados/i.test(v)).join(' ');
      if (t) { const m = t.match(/codigo centro es ([^\n]+)/i); if (m) return (m[1].match(/AV\d+/gi) || []).map(x => x.toUpperCase()); }
    }
    return null;
  }
  function parseSL(rows) {
    // Base 1: Codigo Centro, SKU, Volumen Solicitado HL, Volumen Planeado HL, Volumen Entregado HL
    const h1 = findHeader(rows, ['codigo centro', 'sku', 'volumen solicitado hl']);
    if (h1 >= 0) {
      const h = rows[h1], cC = colIdx(h, 'codigo centro'), cS = colIdx(h, 'sku'), cSol = colIdx(h, 'volumen solicitado hl'), cPl = colIdx(h, 'volumen planeado hl'), cEn = colIdx(h, 'volumen entregado hl');
      const sol = {};
      for (let i = h1 + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue; const c = key(r[cC]), k = key(r[cS]);
        if (!/^AV\d+$/.test(c) || !/^\d+$/.test(k)) continue;
        (sol[c] || (sol[c] = {}))[k] = { sol: num(r[cSol]), plan: cPl >= 0 ? num(r[cPl]) : 0, ent: cEn >= 0 ? num(r[cEn]) : 0 };
      }
      return { tipo: 'sl1', sol };
    }
    // Base 2 en formato largo: una fila por Codigo Centro + SKU + Clase Documento
    const hl = findHeader(rows, ['codigo centro', 'sku', 'clase documento', 'volumen entregado hl']);
    if (hl >= 0) {
      const h = rows[hl], cC = colIdx(h, 'codigo centro'), cS = colIdx(h, 'sku'), cD = colIdx(h, 'clase documento'), cV = colIdx(h, 'volumen entregado hl');
      const canal = {}, MAP = { ZCP1: 'tat', ZCP6: 'ka', ZCPV: 'soc' };
      for (let i = hl + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue; const c = key(r[cC]), k = key(r[cS]), t = MAP[key(r[cD])];
        if (!/^AV\d+$/.test(c) || !/^\d+$/.test(k) || !t) continue;
        const o = ((canal[c] || (canal[c] = {}))[k] || (canal[c][k] = { tat: 0, ka: 0, soc: 0 })); o[t] += num(r[cV]);
      }
      return { tipo: 'sl2', canal };
    }
    // Base 2: Clase Documento (ZCP1 = TAT, ZCP6 = KA, ZCPV = Socios) por SKU
    const h2 = findHeader(rows, ['clase documento']);
    if (h2 >= 0) {
      const top = rows[h2], sub = rows[h2 + 1] || [];
      const cols = {}; top.forEach((v, j) => { const n = key(v); if (n === 'ZCP1' || n === 'ZCP6' || n === 'ZCPV') cols[n] = j; });
      const cS = sub.findIndex(v => norm(v) === 'sku'), cC = sub.findIndex(v => norm(v) === 'codigo centro');
      const filt = filtroCentros(rows);
      const fijo = cC < 0 ? (filt && filt.length === 1 ? filt[0] : null) : null;
      if (cC < 0 && !fijo) return { tipo: 'sl2', combinado: true, centros: filt || [] };
      const canal = {};
      for (let i = h2 + 2; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue; const k = key(r[cS < 0 ? 0 : cS]); if (!/^\d+$/.test(k)) continue;
        const c = fijo || key(r[cC]); if (!/^AV\d+$/.test(c)) continue;
        (canal[c] || (canal[c] = {}))[k] = { tat: num(r[cols.ZCP1]), ka: num(r[cols.ZCP6]), soc: num(r[cols.ZCPV]) };
      }
      return { tipo: 'sl2', canal };
    }
    return null;
  }

  /* ---------- Pronóstico (FC del mes) y ABC ---------- */
  function abcDe(m) { // m: {sku: fc}. A = SKUs que suman el 80 %, B el siguiente 15 %, C el resto (regla de la pestaña ABC)
    const arr = Object.entries(m).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]), tot = arr.reduce((a, [, v]) => a + v, 0);
    const out = {}; let acc = 0;
    for (const [k, v] of arr) { out[k] = acc / tot < 0.8 ? 'A' : acc / tot < 0.95 ? 'B' : 'C'; acc += v; }
    return out;
  }
  function parseFC(rows) {
    const h = findHeader(rows, ['cd', 'sku_demanda']);
    if (h < 0) throw new Error('Pronóstico: no encontré las columnas CD y Sku_Demanda');
    const hdr = rows[h], cC = colIdx(hdr, 'cd'), cS = colIdx(hdr, 'sku_demanda');
    // columnas de semanas: las que dicen "hl" y, en la fila "Tipo", "Pronostico"
    const tipoRow = rows.slice(0, h).find(r => r && norm(r[0]) === 'tipo') || [];
    const semRow = rows.slice(0, h).find(r => r && norm(r[0]) === 'semana') || [];
    const wcols = hdr.map((v, j) => j).filter(j => norm(hdr[j]) === 'hl' && (!tipoRow.length || norm(tipoRow[j] || tipoRow.slice(0, j + 1).filter(x => x != null).pop()).startsWith('pronostico')));
    if (!wcols.length) throw new Error('Pronóstico: no encontré las columnas de semanas (hl)');
    const fc = {};
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue; const c = key(r[cC]), k = key(r[cS]);
      if (!/^AV\d+$/.test(c) || !/^\d+$/.test(k)) continue;
      const m = fc[c] || (fc[c] = {}); m[k] = (m[k] || 0) + wcols.reduce((a, j) => a + num(r[j]), 0);
    }
    let mes = null; const f = rows.slice(-6).flat().find(v => typeof v === 'string' && /mes es/i.test(v)); if (f) { const mm = f.match(/mes es (\d+)/i); if (mm) mes = +mm[1]; }
    const out = {};
    for (const c in fc) { const abc = abcDe(fc[c]); out[c] = {}; for (const k in fc[c]) out[c][k] = { fc: Math.round(fc[c][k] * 100) / 100, ka: 0, abc: abc[k] || 'C' }; }
    return { tipo: 'fc', mes, semanas: wcols.map(j => semRow[j]).filter(v => v != null), fc: out };
  }

  /* ---------- Maestro de clientes (código -> nombre) ---------- */
  function parseClientes(rows) {
    const h = findHeader(rows, ['cliente', 'nombrecliente']);
    if (h < 0) throw new Error('Clientes: no encontré las columnas CLIENTE y NOMBRECLIENTE');
    const cC = colIdx(rows[h], 'cliente'), cN = colIdx(rows[h], 'nombrecliente'), out = {};
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r || r[cC] == null) continue;
      const k = String(r[cC]).trim().replace(/\.0+$/, ''), n = r[cN] == null ? '' : String(r[cN]).trim();
      if (k && n && !(k in out)) out[k] = n;
    }
    return out;
  }

  /* ---------- Reconocer qué archivo es ---------- */
  function classify(rows) {
    if (findHeader(rows, ['material', 'alm']) >= 0) return 'zexi';
    if (findHeader(rows, ['cd', 'sku_demanda']) >= 0) return 'fc';
    if (findHeader(rows, ['shipment destination location id', 'sku', 'hl']) >= 0 || findHeader(rows, ['llave', 'hl']) >= 0) return 'tms';
    if (findHeader(rows, ['cd', 'sku', 'tipopedido']) >= 0) return 'prev';
    for (let i = 0; i < Math.min(rows.length, 10); i++) if ((rows[i] || []).some(v => /preventa\s*d\+[13]/i.test(String(v ?? '')))) return 'prev';
    if (findHeader(rows, ['cliente', 'nombrecliente']) >= 0) return 'cli';
    if (findHeader(rows, ['codigo centro', 'sku', 'volumen solicitado hl']) >= 0) return 'sl1';
    if (findHeader(rows, ['clase documento']) >= 0) return 'sl2';
    return null;
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
  function compute(master, zexiRows, tmsRows, prevAll, today = new Date(), CENTRO = 'AV19', extra = {}) {
    const prev = prevAll && prevAll.d1 ? prevAll : ((prevAll && (prevAll[CENTRO] || prevAll['*'])) || { d1: {}, d3: {} });
    // Los maestros traen FC, entregas y ABC de Villavicencio: para otro centro se usan solo si se cargan sus archivos
    const nativo = CENTRO === ((master[0] && master[0].Centro) || 'AV19');
    const sl = extra.sl || null, canal = extra.canal || null, fc = extra.fc || null;
    const day = today.getDate();
    const BJ = day / 31 > 1 ? 1 : day / 30;
    const semana = weeknum(today);
    const conteOf = {}; master.forEach(m => conteOf[key(m.Material)] = num(m.conte));
    // índices ZEXI
    const zByKey = {}, zFirstByMat = {};
    zexiRows.forEach(z => {
      const k = z.mat + z.ce + z.alm; if (!(k in zByKey)) zByKey[k] = z;
      if (!((z.mat + z.ce) in zFirstByMat)) zFirstByMat[z.mat + z.ce] = z;
    });
    // TMS: suma de HL y primera fecha de arribo por llave
    const tmsHl = {}, tmsF = {};
    tmsRows.forEach(t => { tmsHl[t.llave] = (tmsHl[t.llave] || 0) + t.hl; const f = toDateStr(t.fecha); if (f && (!tmsF[t.llave] || f < tmsF[t.llave])) tmsF[t.llave] = f; });

    const rows = master.map(m => {
      const mat = key(m.Material), centro = CENTRO, llave = mat + centro;
      const pre = centro.replace(/\D/g, ''); // AV19 -> 19xx, AV24 -> 24xx
      const Q = num(m.conte), P = num(m.presentacio);
      const zA = zByKey[llave + pre + '01'];
      const J = zA ? zA.libre : 0;
      const K = nativo ? num(m['Cant entrega Pendien']) : 0, M = nativo ? num(m['Cantidad en Pedidos']) : 0;
      const L = J - K, N = L - M;
      const R = div(J * Q, 1e5), S = div(K * Q, 1e5), T = div(M * P * Q, 1e5);
      const U = N < 0 ? (M + N) * P * Q / 1e5 : M * P * Q / 1e5;
      const p1 = prev.d1[mat] || {}, p3 = prev.d3[mat] || {};
      const V = p1.tat || 0, W = p1.ka || 0, AV = p1.total || 0, BG = p3.total || 0;
      const zT = zByKey[llave + pre + '40'];
      const X = zT ? (zT.libre + zT.ctrl + zT.bloq) * Q / 1e5 : 0;
      const zC = zByKey[llave + pre + '41'];
      const CAL = zA ? zA.ctrl * Q / 1e5 : 0, BLQ = zA ? zA.bloq * Q / 1e5 : 0;
      const CALLE = zC ? (zC.libre + zC.ctrl + zC.bloq) * Q / 1e5 : 0;
      const F = fc ? fc[mat] : null;
      const AL = fc ? (F ? F.fc : 0) : (nativo ? num(m['FC TOTAL CD']) : 0), AM = AL * BJ;
      const Z = div(AM, day);
      const zF = zFirstByMat[mat + CENTRO];
      const Y = zF ? div(zF.ctrl * Q / 1e5, Z) : 0;
      const AA = div(R, Z);
      const AD = tmsHl[llave] || 0, AE = tmsF[llave] || null;
      const AI = R + X + AD - S - T - AV;
      const AJ = Z === 0 ? null : AI / Z;
      const AO = nativo ? num(m['Forescast KA']) : (F ? F.ka : 0);
      const AR = sl ? (sl[mat] ? sl[mat].sol : 0) : (nativo ? num(m['Volumen Solicitado']) : 0), AT = nativo ? num(m['Volumen recibido2']) : 0;
      const Cn = canal ? (canal[mat] || { tat: 0, ka: 0, soc: 0 }) : null;
      const AY = Cn ? Cn.tat : (nativo ? num(m['Entregado TAT']) : 0), AZ = Cn ? Cn.ka : (nativo ? num(m['Entregado KA']) : 0), BB = Cn ? Cn.soc : (nativo ? num(m['Entregado Socios']) : 0);
      const AN = AY + AZ + BB;
      const und = num(m._undEst), cj = num(m._cjEst);
      const BH = (Q && und) ? Math.round((R - Z * 1.5) * 100 / (Q / 1000)) / und : 0;
      const v = {
        'Llave': llave, 'Semana': semana, 'Tipo PT': m['Tipo PT'], 'Marca': m.Marca, 'Lines': m.Lines, 'Centro': m.Centro, 'Material': m.Material,
        'Descripcion': m.Descripcion, 'UM': m.UM, 'Existencia': J, 'Cant entrega Pendien': K, 'Total Disponible': L, 'Cantidad en Pedidos': M,
        'Inv. Final': N, 'tipo': m.tipo, 'presentacio': P, 'conte': Q, 'Inv. Piso': R, 'Inv. Reservado': S, 'Venta Total': T, 'Preventa Hl': U,
        'Venta tat': V, 'Venta Ka': W, 'Inv. Transito': X, 'Stock out': Y, 'FC Dia': Z, 'Dias cobertura': AA, 'Stock Seguridad HL': 0,
        'Stock Seguridad DC': 0, 'TMS': AD, 'Fecha Arribo TMS': AE, 'Fecha Prox Prod': nativo ? m['Fecha Prox Prod'] : null, 'Planta': nativo ? (m.Planta || null) : null, 'Htls Prod Plan': nativo ? m['Htls Prod Plan'] : null,
        'Riesgo HL': AI, 'Riesgo DC': AJ, 'FC TOTAL CD': AL, 'FC MES': AM, 'Volumen Entregado': AN, 'Forescast KA': AO,
        'Forescaste TAT': Math.max(0, AM - AO), 'FC MTD KA': AO * BJ, 'Volumen Solicitado': AR, '%Cumpl. Ventas': div(AR, AM),
        'Volumen recibido2': AT, '%Cumpl. T1': div(AT, AM), 'Venta W+1': AV, 'Cumpl. CD': div(AN, AM), 'Cumpl. Ventas': div(AR, AM),
        'Entregado TAT': AY, 'Entregado KA': AZ, 'Entregado Socios': BB, 'DC TMS': div(AD, Z), 'DC BODEGA': div(R, Z), 'DC PREVENTA': div(AV, Z),
        'DC TRANSITOS': div(X, Z), 'VENTA D+3': BG, 'PRD 1.5 EST': BH, 'PRD 1,5 PQ': BH * cj,
        'Inv. Calidad': CAL, 'Inv. Bloqueo': BLQ, 'Inv. Calle': CALLE, 'ABC': fc ? (F ? F.abc : '') : (nativo ? (m.ABC || '') : '')
      };
      return COLUMNS.map(c => { const x = v[c]; return typeof x === 'number' ? Math.round(x * 100) / 100 : (x == null ? null : x); });
    });
    const matched = master.filter(m => zByKey[key(m.Material) + CENTRO + CENTRO.replace(/\D/g, '') + '01']).length;
    return { columns: COLUMNS, rows, dia: day, semana, centro: CENTRO, clientes: prev.clientes || null, d1day: prev.d1day || null, dias: prev.dias || null, stats: { zexi: zexiRows.length, tms: tmsRows.length, prevD1: Object.keys(prev.d1).length, prevD3: Object.keys(prev.d3).length, skusConZexi: matched } };
  }

  /* Elige la hoja correcta dentro de un libro (por nombre, o la primera) */
  function pickSheet(wb, XLSX, names) {
    const n = wb.SheetNames.find(s => names.some(x => norm(s) === norm(x))) || wb.SheetNames.find(s => names.some(x => norm(s).includes(norm(x)))) || wb.SheetNames[0];
    return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) };
  }

  const api = { COLUMNS, parseZexi, parseTms, parsePreventa, parseSL, parseFC, parseClientes, abcDe, classify, compute, pickSheet, weeknum, norm, num, key };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Engine = api;
})(typeof self !== 'undefined' ? self : this);
