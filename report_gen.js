/**
 * SCM Dashboard — PPT 주간보고서 자동 생성 (PptxGenJS)
 * ====================================================
 * 경영진 주간보고용: 탭별 상세 인사이트 + MoM 비교 + 액션아이템
 */

async function generateReport() {
  if (typeof PptxGenJS === 'undefined') {
    alert('PPT 라이브러리 로딩 실패. 네트워크를 확인해주세요.');
    return;
  }
  const overlay = document.createElement('div');
  overlay.id = 'rpt-overlay';
  overlay.innerHTML = `<div style="position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center">
    <div style="background:#fff;border-radius:12px;padding:32px 48px;text-align:center;box-shadow:0 20px 40px rgba(0,0,0,.3)">
      <div style="font-size:28px;margin-bottom:12px">📊</div>
      <p style="font-size:16px;font-weight:600;color:#1e293b;margin-bottom:8px">주간 보고서 생성 중...</p>
      <p id="rpt-status" style="font-size:12px;color:#64748b">준비</p>
      <div style="width:220px;height:4px;background:#e2e8f0;border-radius:2px;margin-top:12px;overflow:hidden"><div id="rpt-bar" style="width:0%;height:100%;background:#2563eb;transition:width .3s"></div></div>
    </div></div>`;
  document.body.appendChild(overlay);
  const setStatus = (msg, pct) => {
    const s = document.getElementById('rpt-status'); if (s) s.textContent = msg;
    const b = document.getElementById('rpt-bar'); if (b) b.style.width = pct + '%';
  };
  try { await _buildReport(setStatus); } catch (e) { alert('보고서 생성 실패: ' + e.message); console.error(e); }
  finally { const ov = document.getElementById('rpt-overlay'); if (ov) ov.remove(); }
}

async function _buildReport(setStatus) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.author = 'BeaverWorks SCM';
  pres.title = 'BeaverWorks SCM 주간 경영보고서';

  // ═══ PALETTE ═══
  const C = {
    navy: '1E2761', ice: 'CADCFC', white: 'FFFFFF', accent: '2563EB',
    green: '15803D', red: 'DC2626', gray: '64748B', ltGray: 'F1F5F9',
    dark: '1E293B', sub: '475569', orange: 'C2410C', purple: '7E22CE',
    teal: '0E7490', yellow: '92400E', bg: 'F8FAFC'
  };

  // ═══ HELPERS ═══
  const _f = n => n == null ? '0' : Number(n).toLocaleString('ko-KR');
  const _m = n => { if (!n || n === 0) return '0.00M'; return (n / 1e6).toFixed(2) + 'M'; };
  const _p = (v, base) => base > 0 ? ((v - base) / base * 100).toFixed(1) : null;
  const _chg = (cur, prev) => { const p = _p(cur, prev); return p === null ? '' : (p >= 0 ? '+' : '') + p + '%'; };
  const _shortNm = (s, max) => s && s.length > max ? s.slice(0, max) + '..' : (s || '-');

  // ═══ DATA CONTEXT ═══
  const yr = S.selYear;
  const selMs = getSelectedMonths();
  const lastMk = (() => {
    if (S.selMonth === 'all') return yr === '2025' ? '2025-12' : M26_DATA[M26_DATA.length - 1];
    if (S.selMonth.startsWith('Q')) { const qm = getQuarterMonths(S.selMonth, yr); return qm[qm.length - 1]; }
    return yr + '-' + S.selMonth;
  })();
  const allMsFull = [...M25, ...M26];
  const lastIdx = allMsFull.indexOf(lastMk);
  const prevMk = lastIdx > 0 ? allMsFull[lastIdx - 1] : null;
  const reportDate = new Date().toISOString().slice(0, 10);
  const periodStr = S.selMonth === 'all' ? yr + '년 연간' : yr + '.' + parseInt(lastMk.slice(5)) + '월';

  // Previous period months (for MoM)
  let prevMs = [];
  if (selMs.length === 1 && lastIdx > 0) prevMs = [allMsFull[lastIdx - 1]];
  else if (S.selMonth !== 'all' && S.selMonth.startsWith('Q')) {
    const qMap = { Q1: 'Q4', Q2: 'Q1', Q3: 'Q2', Q4: 'Q3' };
    const pYr = S.selMonth === 'Q1' ? String(+yr - 1) : yr;
    prevMs = getQuarterMonths(qMap[S.selMonth], pYr);
  }

  setStatus('데이터 수집 중...', 5);

  // ═══════════════════════════════════════════════
  // DATA AGGREGATION
  // ══════════════════════════════════��════════════

  // ── INVENTORY ──
  let pIn = 0, pOut = 0, pOther = 0, pInAmt = 0, pOutAmt = 0;
  Object.values(mio).forEach(d => { selMs.forEach(m => { const r = d[m]; if (r) { pIn += r.iq || 0; pOut += r.on || 0; pOther += r.ot || 0; pInAmt += r.ia || 0; pOutAmt += r.oa || 0; } }); });
  let endQty = 0, endAmt = 0;
  Object.values(mio).forEach(d => { if (d[lastMk]) { endQty += d[lastMk].eq || 0; endAmt += d[lastMk].ea || 0; } });
  let beginQty = 0, beginAmt = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { beginQty += d[prevMk].eq || 0; beginAmt += d[prevMk].ea || 0; } });
  else { beginQty = endQty - pIn + pOut + pOther; beginAmt = endAmt - pInAmt + pOutAmt; }

  // Health
  const getEq = sku => mio[sku] && mio[sku][lastMk] ? mio[sku][lastMk].eq : 0;
  const activeInv = D.inv.filter(i => (i.qty !== 0 || i.amount !== 0) && i.sku_id.startsWith('B-'));
  let cntNormal = 0, cntLow = 0, cntExcess = 0, cntNoDemand = 0, cntNoStock = 0;
  const lowItems = [], excessItems = [], noDemandItems = [];
  activeInv.forEach(i => {
    const ss = calcSS(i.sku_id); const eq = getEq(i.sku_id);
    const ratio = ss > 0 ? eq / ss : 999;
    if (eq === 0) { cntNoStock++; }
    else if (ratio < 1) { cntLow++; lowItems.push(i); }
    else if (isExcess(i.sku_id, eq)) { cntExcess++; excessItems.push(i); }
    else if (isNoDemand(i.sku_id)) { cntNoDemand++; noDemandItems.push(i); }
    else cntNormal++;
  });
  const totalActive = activeInv.length;
  const healthScore = totalActive > 0 ? Math.round((cntNormal * 100 + cntExcess * 60 + cntNoDemand * 35 + cntLow * 15) / (totalActive * 100) * 100) : 0;

  // Previous period inventory for MoM
  let prevEndAmt = 0, prevEndQty = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { prevEndAmt += d[prevMk].ea || 0; prevEndQty += d[prevMk].eq || 0; } });

  // Category breakdown (inventory by B-category)
  const invByCat = {};
  activeInv.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = B_CATS[cat] || cat;
    if (!invByCat[cat]) invByCat[cat] = { nm, qty: 0, amt: 0, cnt: 0 };
    invByCat[cat].cnt++;
    const eq = getEq(i.sku_id);
    invByCat[cat].qty += eq;
    invByCat[cat].amt += (mio[i.sku_id] && mio[i.sku_id][lastMk] ? mio[i.sku_id][lastMk].ea : 0) || 0;
  });
  const invCatList = Object.values(invByCat).sort((a, b) => b.amt - a.amt);

  // ── SALES ──
  const salesTotal = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const salesQty = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].qty : 0), s), 0);
  const cSalesAmt = D.isd.filter(i => i.type.includes('C')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const bSalesAmt = D.isd.filter(i => i.type.includes('B')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const prevSalesTotal = prevMs.length > 0 ? D.isd.reduce((s, i) => prevMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0) : 0;

  // Monthly sales trend
  const salesMonthly = {};
  const trendMs = (selMs.length > 1 ? selMs : allMsFull.slice(Math.max(0, lastIdx - 5), lastIdx + 1));
  trendMs.forEach(m => { salesMonthly[m] = D.isd.reduce((s, i) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0); });

  // Sales by category
  const salesByCat = {};
  D.isd.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = { ...C_CATS, ...B_CATS }[cat] || cat;
    if (!salesByCat[cat]) salesByCat[cat] = { nm, amt: 0, qty: 0, prevAmt: 0 };
    selMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].amt += d.amt || 0; salesByCat[cat].qty += d.qty || 0; } });
    prevMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].prevAmt += d.amt || 0; } });
  });
  const topSalesCat = Object.values(salesByCat).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);

  // Top selling items
  const topSalesItems = D.isd.map(i => ({
    sku: i.sku_id, name: i.name || (skuMap[i.sku_id] ? skuMap[i.sku_id].name : i.sku_id),
    amt: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0),
    qty: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].qty : 0), 0)
  })).filter(i => i.amt > 0).sort((a, b) => b.amt - a.amt).slice(0, 10);

  // ── COST ──
  const parents = new Set(D.bom.map(b => b.parent_sku));
  let costItems = [];
  parents.forEach(sku => {
    if (S.bomDisc && S.bomDisc[sku]) return;
    const inv = D.inv.find(i => i.sku_id === sku); if (!inv) return;
    const tx = D.sales_tx ? D.sales_tx.filter(t => t.sku_id === sku) : [];
    let paidAmt = 0, paidQty = 0;
    tx.forEach(t => { if ((t.amt || 0) > 0) { paidQty += t.qty || 0; paidAmt += t.amt || 0; } });
    const sp = paidQty > 0 ? Math.round(paidAmt / paidQty) : 0;
    const children = D.bom.filter(b => b.parent_sku === sku);
    let matCost = 0;
    children.forEach(ch => { const ci = D.inv.find(x => x.sku_id === ch.child_sku); matCost += (ci ? ci.unit_cost || 0 : 0) * (ch.qty || 1); });
    const cr = sp > 0 ? Math.round(matCost / sp * 100) : 0;
    const margin = sp > 0 ? sp - matCost : 0;
    if (sp > 0 || matCost > 0) costItems.push({ sku, name: inv.name || sku, sp, matCost, cr, margin, children: children.length });
  });
  costItems.sort((a, b) => b.cr - a.cr);
  const avgCR = costItems.length > 0 ? Math.round(costItems.reduce((s, c) => s + c.cr, 0) / costItems.length) : 0;
  const highCR = costItems.filter(c => c.cr > 70);
  const totalMargin = costItems.reduce((s, c) => s + c.margin, 0);

  // ── PO ──
  const allPO = D.po.filter(p => p.vendor !== 'nan' && p.item_name !== 'nan');
  let periodPO = allPO;
  if (S.selMonth !== 'all') {
    const poSelSet = new Set(selMs.map(m => m.replace('-', '/')));
    periodPO = allPO.filter(p => { const pm = p.date_no.substring(0, 7).replace('-', '/'); return poSelSet.has(pm); });
  }
  const poAmt = periodPO.reduce((s, p) => s + (p.amount || 0), 0);
  const poVendors = [...new Set(periodPO.map(p => p.vendor))];
  const poByVendor = {};
  periodPO.forEach(p => { if (!poByVendor[p.vendor]) poByVendor[p.vendor] = { cnt: 0, amt: 0 }; poByVendor[p.vendor].cnt++; poByVendor[p.vendor].amt += p.amount || 0; });
  const topPOVendors = Object.entries(poByVendor).sort((a, b) => b[1].amt - a[1].amt).slice(0, 7);
  const poByStatus = {};
  periodPO.forEach(p => { const st = p.status || '미확인'; if (!poByStatus[st]) poByStatus[st] = { cnt: 0, amt: 0 }; poByStatus[st].cnt++; poByStatus[st].amt += p.amount || 0; });
  const compRate = periodPO.length > 0 ? Math.round(periodPO.filter(p => p.status === '완료').length / periodPO.length * 100) : 0;

  // Previous PO for MoM
  let prevPOAmt = 0;
  if (prevMs.length > 0) {
    const prevSet = new Set(prevMs.map(m => m.replace('-', '/')));
    allPO.filter(p => { const pm = p.date_no.substring(0, 7).replace('-', '/'); return prevSet.has(pm); }).forEach(p => prevPOAmt += p.amount || 0);
  }

  // ── BOM ──
  const bomCnt = D.bom.length;
  const bomParentCnt = parents.size;
  const usedMats = new Set(D.bom.map(b => b.child_sku));
  const bomMatCnt = usedMats.size;
  // Production feasibility
  let canProduce = 0, cantProduce = 0;
  parents.forEach(sku => {
    if (S.bomDisc && S.bomDisc[sku]) return;
    const children = D.bom.filter(b => b.parent_sku === sku);
    const feasible = children.every(ch => {
      const eq = getEq(ch.child_sku);
      return eq >= (ch.qty || 1);
    });
    if (feasible) canProduce++; else cantProduce++;
  });

  // ── SUPPLIER (거래처) ──
  const csd = D.csd || [];
  const custTotal = csd.reduce((s, c) => selMs.reduce((ss, m) => ss + (c.monthly[m] ? c.monthly[m].amt : 0), s), 0);
  const topCusts = csd.map(c => ({
    name: c.name, amt: selMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0),
    prevAmt: prevMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0)
  })).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);
  // Concentration (top 5 share)
  const top5CustAmt = topCusts.slice(0, 5).reduce((s, c) => s + c.amt, 0);
  const custConcentration = custTotal > 0 ? Math.round(top5CustAmt / custTotal * 100) : 0;

  // ── TA (타계정) ──
  const taAll = D.ta || [];
  const taYr = taAll.filter(t => (t.d || '').startsWith(yr));
  const taTotal = taYr.reduce((s, t) => s + (t.a || 0), 0);
  const taQty = taYr.reduce((s, t) => s + (t.q || 0), 0);
  const taByAcct = {};
  taYr.forEach(t => { const ac = t.ac || '(미분류)'; if (!taByAcct[ac]) taByAcct[ac] = { qty: 0, amt: 0 }; taByAcct[ac].qty += t.q || 0; taByAcct[ac].amt += t.a || 0; });
  const taAcctList = Object.entries(taByAcct).sort((a, b) => b[1].amt - a[1].amt);
  // Monthly TA trend
  const taMonthly = {};
  const taMs = yr === '2026' ? M26_DATA : M25;
  taMs.forEach(m => { taMonthly[m] = taYr.filter(t => (t.d || '').substring(5, 7) === m.slice(5)).reduce((s, t) => s + (t.a || 0), 0); });

  setStatus('슬라이드 생성 시작...', 10);

  // ══════════���════════════════════════════════════════════════
  // SLIDE GENERATION
  // ═══════════════════════════════════════════════════════════

  // Helper: Section title slide
  function sectionSlide(title, subtitle, icon) {
    const sl = pres.addSlide();
    sl.background = { color: C.navy };
    sl.addText(icon || '', { x: 0.8, y: 1.8, w: 1, h: 1, fontSize: 40 });
    sl.addText(title, { x: 2.0, y: 1.8, w: 7, h: 0.8, fontSize: 32, fontFace: 'Arial Black', color: C.white });
    sl.addText(subtitle, { x: 2.0, y: 2.7, w: 7, h: 0.5, fontSize: 14, color: C.ice });
    sl.addShape(pres.shapes.RECTANGLE, { x: 2.0, y: 3.5, w: 1.5, h: 0.05, fill: { color: C.accent } });
    return sl;
  }

  // Helper: Insight box at bottom of slide
  function addInsights(sl, insights, yPos) {
    const y = yPos || 4.3;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.04, fill: { color: 'E2E8F0' } });
    sl.addText([
      { text: '💡 Insights', options: { bold: true, fontSize: 11, breakLine: true } },
      ...insights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < insights.length - 1, fontSize: 10 } }))
    ], { x: 0.5, y: y + 0.1, w: 9, h: 1.1, color: C.dark, valign: 'top' });
  }

  // Helper: MoM badge
  function momBadge(cur, prev) {
    if (!prev || prev === 0) return '';
    const pct = ((cur - prev) / prev * 100).toFixed(1);
    return (pct >= 0 ? '▲' : '▼') + Math.abs(pct) + '% MoM';
  }

  // ══════════════════════════════════════════���════════════════
  // 1. COVER
  // ═════════════════════════════════════��═════════════════════
  setStatus('표지 생성...', 12);
  const s1 = pres.addSlide();
  s1.background = { color: C.navy };
  s1.addText('BeaverWorks', { x: 0.8, y: 1.2, w: 8.5, h: 0.9, fontSize: 48, fontFace: 'Arial Black', color: C.white, bold: true });
  s1.addText('SCM 주간 경영보고서', { x: 0.8, y: 2.1, w: 8.5, h: 0.6, fontSize: 28, color: C.ice });
  s1.addText(`보고 기간: ${periodStr}`, { x: 0.8, y: 3.0, w: 5, h: 0.4, fontSize: 14, color: C.gray });
  s1.addText(`생성일: ${reportDate}`, { x: 0.8, y: 3.4, w: 5, h: 0.4, fontSize: 12, color: C.gray });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.3, w: 2.0, h: 0.06, fill: { color: C.accent } });
  s1.addText('비버웍스 공급망관리 시스템 | 자동생성 보고서', { x: 0.8, y: 4.6, w: 6, h: 0.4, fontSize: 11, color: C.gray });

  // ══════════════════════════════════���════════════════════════
  // 2. EXECUTIVE SUMMARY
  // ═════════════════════════════���═════════════════════════════
  setStatus('Executive Summary...', 15);
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addText('Executive Summary', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 26, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s2.addText(periodStr + ' 핵심 경영지표 요약', { x: 0.5, y: 0.75, w: 9, h: 0.3, fontSize: 12, color: C.sub });

  const kpis = [
    { label: '총 매출', val: _m(salesTotal), chg: _chg(salesTotal, prevSalesTotal), color: C.accent },
    { label: '재고자산', val: _m(endAmt), chg: _chg(endAmt, prevEndAmt), color: C.teal },
    { label: '재고건강도', val: healthScore + '점', chg: '', color: healthScore >= 70 ? C.green : C.orange },
    { label: '발주규모', val: _m(poAmt), chg: _chg(poAmt, prevPOAmt), color: C.purple },
    { label: '평균원가율', val: avgCR + '%', chg: '', color: avgCR > 60 ? C.red : C.green },
    { label: '타계정', val: _m(taTotal), chg: '', color: C.orange },
    { label: '납기완료율', val: compRate + '%', chg: '', color: compRate >= 80 ? C.green : C.red },
    { label: '거래처집중도', val: custConcentration + '%', chg: 'TOP5 비중', color: custConcentration > 80 ? C.orange : C.teal },
  ];
  kpis.forEach((kpi, idx) => {
    const col = idx % 4; const row = Math.floor(idx / 4);
    const x = 0.4 + col * 2.4; const y = 1.2 + row * 1.65;
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 1.45, fill: { color: C.ltGray }, shadow: { type: 'outer', color: '000000', blur: 3, offset: 1, angle: 135, opacity: 0.06 } });
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.05, h: 1.45, fill: { color: kpi.color } });
    s2.addText(kpi.label, { x: x + 0.15, y: y + 0.08, w: 1.9, h: 0.3, fontSize: 10, color: C.sub, margin: 0 });
    s2.addText(kpi.val, { x: x + 0.15, y: y + 0.4, w: 1.9, h: 0.55, fontSize: 22, bold: true, color: C.dark, margin: 0 });
    if (kpi.chg) {
      const chgColor = kpi.chg.includes('+') || kpi.chg.includes('▲') ? C.green : kpi.chg.includes('-') || kpi.chg.includes('▼') ? C.red : C.sub;
      s2.addText(kpi.chg, { x: x + 0.15, y: y + 1.0, w: 1.9, h: 0.3, fontSize: 9, color: chgColor, margin: 0 });
    }
  });

  // Quick summary text
  const summaryBullets = [];
  if (prevSalesTotal > 0) summaryBullets.push(`매출 ${salesTotal > prevSalesTotal ? '증가' : '���소'} (${_chg(salesTotal, prevSalesTotal)})`);
  if (cntLow + cntNoStock > 3) summaryBullets.push(`재고부족 ${cntLow + cntNoStock}건 주의`);
  if (highCR.length > 3) summaryBullets.push(`고원가율 ${highCR.length}건 관리 필요`);
  if (summaryBullets.length > 0) {
    s2.addText(summaryBullets.join(' | '), { x: 0.4, y: 4.8, w: 9.2, h: 0.35, fontSize: 10, color: C.sub, italic: true });
  }

  // ══════════════��════════════════════════════════════════════
  // 3-4. 재고 현황 (2 slides)
  // ═════════════════════════════════��═════════════════════════
  setStatus('재고 현황 분석...', 25);

  // Slide 3: 재고 흐름 + 건강도
  const s3 = pres.addSlide();
  s3.background = { color: C.white };
  s3.addText('📦 재고 현황', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s3.addText(periodStr, { x: 7, y: 0.3, w: 2.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Waterfall chart
  const wfLabels = ['기초', '입고', '출고', '��계정', '기말'];
  const wfAmts = [beginAmt, pInAmt, pOutAmt, beginAmt + pInAmt - pOutAmt - endAmt, endAmt];
  s3.addChart(pres.charts.BAR, [{
    name: '금액(백만)', labels: wfLabels, values: wfAmts.map(v => Math.round(v / 1e6))
  }], {
    x: 0.4, y: 0.9, w: 5.0, h: 2.8, barDir: 'col',
    showTitle: true, title: '재고 흐름 워터풀 (백만원)', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.accent], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  // Health gauge (right) - as pie chart
  s3.addChart(pres.charts.DOUGHNUT, [{
    name: '상태', labels: ['정상', '부족', '재고없음', '과다', '무수요'],
    values: [cntNormal, cntLow, cntNoStock, cntExcess, cntNoDemand]
  }], {
    x: 5.6, y: 0.9, w: 4.0, h: 2.8,
    showTitle: true, title: `재고 건강도: ${healthScore}점`, titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.green, C.yellow, C.red, C.orange, C.purple],
    showPercent: true, showLegend: true, legendPos: 'b', legendFontSize: 9
  });

  // Insights
  const invIns = [];
  invIns.push(`기말재고 ${_m(endAmt)} (${_f(endQty)}개)${prevEndAmt > 0 ? ' — 전월대비 ' + _chg(endAmt, prevEndAmt) : ''}`);
  invIns.push(`입고 ${_m(pInAmt)} / 출고 ${_m(pOutAmt)} — 회전율 ${endAmt > 0 ? (pOutAmt / endAmt * 100).toFixed(0) + '%' : '-'}`);
  if (cntLow + cntNoStock > 0) invIns.push(`⚠️ 재고부족 ${cntLow}건 + 재고없음 ${cntNoStock}건 → 긴급발주 필요`);
  if (cntExcess > 5) invIns.push(`과다재고 ${cntExcess}건 — 적정재고 대비 6개월분 초과`);
  addInsights(s3, invIns, 4.0);

  // Slide 4: 카테고리별 재고 + 위험 품목
  setStatus('재고 상세 분석...', 30);
  const s4 = pres.addSlide();
  s4.background = { color: C.white };
  s4.addText('📦 재고 상세 — 카테고리별 & 위험품목', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // Category table (left)
  const catHdr = [
    { text: '카테고리', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '품목수', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '수량', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const catRows = invCatList.slice(0, 8).map(c => [
    _shortNm(c.nm, 8), { text: _f(c.cnt), options: { align: 'right' } },
    { text: _f(c.qty), options: { align: 'right' } }, { text: _m(c.amt), options: { align: 'right' } }
  ]);
  s4.addTable([catHdr, ...catRows], { x: 0.4, y: 0.9, w: 4.8, h: 0.3 + catRows.length * 0.38, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.4, 0.9, 1.1, 1.4] });

  // Risk items table (right)
  s4.addText('⚠️ 부족 품목 TOP', { x: 5.5, y: 0.85, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.red, margin: 0 });
  const riskItems = lowItems.slice(0, 6).map(i => {
    const nm = (skuMap[i.sku_id] ? skuMap[i.sku_id].name : i.sku_id) || i.sku_id;
    const eq = getEq(i.sku_id); const ss = calcSS(i.sku_id);
    return [_shortNm(nm, 10), { text: _f(eq), options: { align: 'right' } }, { text: _f(ss), options: { align: 'right', color: C.red } }];
  });
  if (riskItems.length > 0) {
    s4.addTable([
      [{ text: '품목', options: { bold: true, fill: { color: C.red }, color: C.white } },
       { text: '현재고', options: { bold: true, fill: { color: C.red }, color: C.white, align: 'right' } },
       { text: '안전재고', options: { bold: true, fill: { color: C.red }, color: C.white, align: 'right' } }],
      ...riskItems
    ], { x: 5.5, y: 1.2, w: 4.1, h: 0.3 + riskItems.length * 0.35, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.8, 1.0, 1.3] });
  }

  // Excess items
  const excessTop = excessItems.slice(0, 4).map(i => {
    const nm = (skuMap[i.sku_id] ? skuMap[i.sku_id].name : i.sku_id) || i.sku_id;
    return _shortNm(nm, 12) + ' (' + _f(getEq(i.sku_id)) + '개)';
  });
  if (excessTop.length > 0) {
    s4.addText([
      { text: '📈 과다재고 TOP', options: { bold: true, breakLine: true, fontSize: 11, color: C.orange } },
      ...excessTop.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < excessTop.length - 1, fontSize: 9 } }))
    ], { x: 5.5, y: 1.2 + 0.3 + riskItems.length * 0.35 + 0.3, w: 4.1, h: 1.5, color: C.dark, valign: 'top' });
  }

  // ══════════════════════��════════════════════════════════���═══
  // 5-6. 판매 분석 (2 slides)
  // ═══════════════════════════════════════════════════════════
  setStatus('판매 분석...', 40);

  const s5 = pres.addSlide();
  s5.background = { color: C.white };
  s5.addText('💰 판매 분석', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s5.addText(periodStr + (prevSalesTotal > 0 ? ' | MoM ' + _chg(salesTotal, prevSalesTotal) : ''), { x: 5.5, y: 0.3, w: 4, h: 0.35, fontSize: 11, color: prevSalesTotal > 0 && salesTotal >= prevSalesTotal ? C.green : C.red, align: 'right' });

  // Monthly trend line chart
  const tLabels = Object.keys(salesMonthly).map(m => m.slice(5) + '월');
  const tValues = Object.values(salesMonthly).map(v => Math.round(v / 1e6 * 100) / 100);
  if (tLabels.length > 1) {
    s5.addChart(pres.charts.LINE, [{ name: '매출(백만)', labels: tLabels, values: tValues }], {
      x: 0.4, y: 0.9, w: 5.5, h: 2.8,
      showTitle: true, title: '월별 매출 추이', titleColor: C.sub, titleFontSize: 11,
      lineSize: 3, lineSmooth: true, chartColors: [C.accent],
      showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // Category pie (right)
  if (topSalesCat.length > 0) {
    s5.addChart(pres.charts.DOUGHNUT, [{
      name: '카테고리', labels: topSalesCat.slice(0, 6).map(c => c.nm),
      values: topSalesCat.slice(0, 6).map(c => Math.round(c.amt / 1e6))
    }], {
      x: 6.1, y: 0.9, w: 3.6, h: 2.8,
      showTitle: true, title: '카테고리별', titleColor: C.sub, titleFontSize: 11,
      chartColors: ['2563EB', '0E7490', '15803D', '7E22CE', 'C2410C', '64748B'],
      showPercent: true, showLegend: true, legendPos: 'b', legendFontSize: 8
    });
  }

  // Sales insights
  const sIns = [];
  sIns.push(`총 매출 ${_m(salesTotal)} — 완제품 ${Math.round(cSalesAmt / (salesTotal || 1) * 100)}% / 원재료 ${Math.round(bSalesAmt / (salesTotal || 1) * 100)}%`);
  if (topSalesCat.length > 0) sIns.push(`주력 카테고리: ${topSalesCat[0].nm} (${_m(topSalesCat[0].amt)}, ${salesTotal > 0 ? Math.round(topSalesCat[0].amt / salesTotal * 100) : 0}%)`);
  if (topSalesCat.length > 0 && topSalesCat[0].prevAmt > 0) sIns.push(`${topSalesCat[0].nm} MoM: ${_chg(topSalesCat[0].amt, topSalesCat[0].prevAmt)}`);
  addInsights(s5, sIns, 4.0);

  // Slide 6: TOP 품목 + 카테고리 상세
  setStatus('판매 TOP 품목...', 48);
  const s6 = pres.addSlide();
  s6.background = { color: C.white };
  s6.addText('💰 판매 상세 — TOP 품목 & 카테고리 비교', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // Top items table
  const itemHdr = [
    { text: '순위', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '품목명', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '수량', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const itemRows = topSalesItems.slice(0, 8).map((it, i) => [
    String(i + 1), _shortNm(it.name, 14),
    { text: _f(it.qty), options: { align: 'right' } },
    { text: _m(it.amt), options: { align: 'right' } },
    { text: salesTotal > 0 ? Math.round(it.amt / salesTotal * 100) + '%' : '-', options: { align: 'right' } }
  ]);
  s6.addTable([itemHdr, ...itemRows], { x: 0.4, y: 0.85, w: 5.5, h: 0.3 + itemRows.length * 0.38, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [0.5, 2.2, 0.8, 1.1, 0.9] });

  // Category comparison bar chart (right)
  if (topSalesCat.length > 1) {
    const catLabels = topSalesCat.slice(0, 5).map(c => _shortNm(c.nm, 6));
    const catVals = topSalesCat.slice(0, 5).map(c => Math.round(c.amt / 1e6));
    const catPrev = topSalesCat.slice(0, 5).map(c => Math.round(c.prevAmt / 1e6));
    const chartData = [{ name: '당기', labels: catLabels, values: catVals }];
    if (prevMs.length > 0) chartData.push({ name: '전기', labels: catLabels, values: catPrev });
    s6.addChart(pres.charts.BAR, chartData, {
      x: 6.1, y: 0.85, w: 3.6, h: 3.0, barDir: 'col',
      showTitle: true, title: '카테고리 비교', titleColor: C.sub, titleFontSize: 10,
      chartColors: [C.accent, 'B0C4DE'], showValue: false,
      showLegend: prevMs.length > 0, legendPos: 'b', legendFontSize: 8,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // ═══════════════════════════════════════════════════════════
  // 7-8. 원가 분석 (2 slides)
  // ══════════════════════════════════���════════════════════════
  setStatus('원가 분석...', 55);

  const s7 = pres.addSlide();
  s7.background = { color: C.white };
  s7.addText('💹 원가 분석', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s7.addText(`분석 대상: ${costItems.length}개 완제품`, { x: 6, y: 0.3, w: 3.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Cost ratio distribution
  const crBands = [
    { label: '~30%', cnt: costItems.filter(c => c.cr > 0 && c.cr <= 30).length, color: C.green },
    { label: '30~50%', cnt: costItems.filter(c => c.cr > 30 && c.cr <= 50).length, color: C.teal },
    { label: '50~70%', cnt: costItems.filter(c => c.cr > 50 && c.cr <= 70).length, color: C.yellow },
    { label: '70%~', cnt: costItems.filter(c => c.cr > 70).length, color: C.red },
  ];
  s7.addChart(pres.charts.BAR, [{
    name: '품목수', labels: crBands.map(b => b.label), values: crBands.map(b => b.cnt)
  }], {
    x: 0.4, y: 0.9, w: 4.5, h: 2.8, barDir: 'col',
    showTitle: true, title: '원가율 분포', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.green, C.teal, C.yellow, C.red],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  // Summary KPIs (right)
  s7.addText('원가 요약', { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.sub, margin: 0 });
  const costSummary = [
    ['평균 원가율', avgCR + '%'],
    ['고위험(>70%)', highCR.length + '건'],
    ['저원가(<30%)', costItems.filter(c => c.cr > 0 && c.cr <= 30).length + '건'],
    ['총 마진합계', _m(totalMargin)],
    ['최고 원가율', costItems.length > 0 ? costItems[0].cr + '% (' + _shortNm(costItems[0].name, 8) + ')' : '-'],
  ];
  const csRows = [
    [{ text: '지표', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '값', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
    ...costSummary.map(r => [r[0], { text: r[1], options: { align: 'right' } }])
  ];
  s7.addTable(csRows, { x: 5.3, y: 1.3, w: 4.3, h: 2.2, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.0, 2.3] });

  const cIns = [];
  cIns.push(`평균 원가율 ${avgCR}% — ${avgCR > 50 ? '마진 관리 주의' : '양호한 수준'}`);
  if (highCR.length > 0) cIns.push(`고원가율(>70%) ${highCR.length}건: ${highCR.slice(0, 3).map(c => _shortNm(c.name, 6)).join(', ')}`);
  cIns.push(`총 마진합계 ${_m(totalMargin)} (${costItems.length}개 품목 기준)`);
  addInsights(s7, cIns, 4.0);

  // Slide 8: 고위험 품목 상세
  setStatus('원가 상세...', 60);
  const s8 = pres.addSlide();
  s8.background = { color: C.white };
  s8.addText('💹 원가 상세 — 고위험 품목 & 마진 분석', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // High CR table
  const crHdr = [
    { text: '품목', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '판매가', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '원가', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '원가율', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '마진', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const crRows = costItems.slice(0, 10).map(c => [
    _shortNm(c.name, 14),
    { text: _f(c.sp), options: { align: 'right' } },
    { text: _f(c.matCost), options: { align: 'right' } },
    { text: c.cr + '%', options: { align: 'right', color: c.cr > 70 ? C.red : c.cr > 50 ? C.orange : C.green, bold: c.cr > 70 } },
    { text: _f(c.margin), options: { align: 'right', color: c.margin < 0 ? C.red : C.dark } }
  ]);
  s8.addTable([crHdr, ...crRows], { x: 0.4, y: 0.85, w: 9.2, h: 0.3 + crRows.length * 0.38, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [3.0, 1.5, 1.5, 1.2, 2.0] });

  // Margin scatter description
  s8.addText([
    { text: '원가 개선 권장 사항', options: { bold: true, breakLine: true, fontSize: 11 } },
    { text: '원가율 70% 이상 품목은 BOM 원재료 단가 재협상 또는 설계변경(VA/VE) 검토 필요', options: { bullet: true, breakLine: true, fontSize: 10 } },
    { text: '마진 마이너스 품목은 판매가 조정 또는 단종 검토 대상', options: { bullet: true, breakLine: true, fontSize: 10 } },
    { text: '원가율 30% 미만 품목��� 프리미엄 전략 유지 및 판매 확대 권장', options: { bullet: true, fontSize: 10 } }
  ], { x: 0.4, y: 0.85 + 0.3 + crRows.length * 0.38 + 0.2, w: 9.2, h: 1.2, color: C.dark, valign: 'top' });

  // ═══════════════════════════════���═══════════════════════════
  // 9-10. 발주 관리 (2 slides)
  // ═══════════════════════════════════════════��═══════════════
  setStatus('발주 분석...', 68);

  const s9 = pres.addSlide();
  s9.background = { color: C.white };
  s9.addText('📋 발주 관리', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s9.addText(`${periodStr} | ${_f(periodPO.length)}건 / ${_m(poAmt)}`, { x: 5, y: 0.3, w: 4.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Vendor bar chart
  if (topPOVendors.length > 0) {
    s9.addChart(pres.charts.BAR, [{
      name: '발주액(백만)', labels: topPOVendors.map(v => _shortNm(v[0], 8)),
      values: topPOVendors.map(v => Math.round(v[1].amt / 1e6))
    }], {
      x: 0.4, y: 0.9, w: 5.0, h: 2.8, barDir: 'bar',
      showTitle: true, title: 'TOP 공급사별 발주금액', titleColor: C.sub, titleFontSize: 11,
      chartColors: [C.purple], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // Status breakdown (right)
  s9.addText('발주 상태별 현황', { x: 5.7, y: 0.9, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.sub, margin: 0 });
  const stRows = [
    [{ text: '상태', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '건수', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
     { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
    ...Object.entries(poByStatus).sort((a, b) => b[1].amt - a[1].amt).map(([st, v]) => [
      st, { text: _f(v.cnt), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
    ])
  ];
  s9.addTable(stRows, { x: 5.7, y: 1.3, w: 3.9, h: 0.3 + Object.keys(poByStatus).length * 0.38, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.4, 1.0, 1.5] });

  const poIns = [];
  poIns.push(`총 ${_f(periodPO.length)}건 발주, ${_m(poAmt)}${prevPOAmt > 0 ? ' (전월대비 ' + _chg(poAmt, prevPOAmt) + ')' : ''}`);
  poIns.push(`납기완료율 ${compRate}% — ${compRate >= 80 ? '양호' : '개선 필요'}`);
  if (topPOVendors.length > 0) poIns.push(`최대 거래처: ${topPOVendors[0][0]} (${_m(topPOVendors[0][1].amt)}, ${poAmt > 0 ? Math.round(topPOVendors[0][1].amt / poAmt * 100) : 0}%)`);
  addInsights(s9, poIns, 4.1);

  // ═══════════════════════════��══════════════════════════��════
  // 11. BOM 관리
  // ═══════════════════════════════════════════════════════════
  setStatus('BOM 분석...', 75);

  const s10 = pres.addSlide();
  s10.background = { color: C.white };
  s10.addText('🔧 BOM 관리 & 생산가능성', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // BOM summary + feasibility
  const bomKpis = [
    ['BOM 총 관계', _f(bomCnt) + '건'],
    ['완제품 (Parent)', _f(bomParentCnt) + '종'],
    ['원재료 (Child)', _f(bomMatCnt) + '종'],
    ['생산가능', _f(canProduce) + '종'],
    ['생산불가 (자재부족)', _f(cantProduce) + '종'],
    ['생산가능률', bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) + '%' : '-'],
  ];
  const bomTbl = [
    [{ text: '지표', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '값', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
    ...bomKpis.map(r => [r[0], { text: r[1], options: { align: 'right' } }])
  ];
  s10.addTable(bomTbl, { x: 0.4, y: 0.9, w: 4.5, h: 2.5, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.5, 2.0] });

  // Feasibility pie
  s10.addChart(pres.charts.PIE, [{
    name: '생산가능성', labels: ['생산가능', '자재부족'],
    values: [canProduce, cantProduce]
  }], {
    x: 5.5, y: 0.9, w: 4.0, h: 2.5,
    showTitle: true, title: '생산가능 현황', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.green, C.red], showPercent: true, showLegend: true, legendPos: 'b'
  });

  const bomIns = [];
  bomIns.push(`${bomParentCnt}개 완제품 중 ${canProduce}개 즉시 생산가능 (${bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) : 0}%)`);
  if (cantProduce > 0) bomIns.push(`⚠️ ${cantProduce}개 품목 자재부족으로 생산불가 — 긴급 자재 확보 필요`);
  bomIns.push(`평균 BOM 깊이: 완제품당 ${bomParentCnt > 0 ? (bomCnt / bomParentCnt).toFixed(1) : 0}개 자재 사용`);
  addInsights(s10, bomIns, 3.8);

  // ═════════════════════════════════════════════════════════���═
  // 12. 거래처 분석
  // ═════════════════════════════��═════════════════════════════
  setStatus('거래처 분석...', 82);

  const s11 = pres.addSlide();
  s11.background = { color: C.white };
  s11.addText('🏭 거래처 분석 (판매)', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s11.addText(`TOP5 집중도: ${custConcentration}%`, { x: 6, y: 0.3, w: 3.5, h: 0.35, fontSize: 11, color: custConcentration > 80 ? C.orange : C.teal, align: 'right' });

  // Top customers table
  const custHdr = [
    { text: '순위', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '거래처', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '매출', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: 'MoM', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const custDataRows = topCusts.slice(0, 10).map((c, i) => {
    const share = custTotal > 0 ? Math.round(c.amt / custTotal * 100) : 0;
    const mom = c.prevAmt > 0 ? _chg(c.amt, c.prevAmt) : '-';
    const momColor = mom.includes('+') ? C.green : mom.includes('-') ? C.red : C.sub;
    return [
      String(i + 1), _shortNm(c.name, 12),
      { text: _m(c.amt), options: { align: 'right' } },
      { text: share + '%', options: { align: 'right' } },
      { text: mom, options: { align: 'right', color: momColor } }
    ];
  });
  s11.addTable([custHdr, ...custDataRows], { x: 0.4, y: 0.85, w: 9.2, h: 0.3 + custDataRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [0.6, 3.0, 1.8, 1.4, 2.4] });

  const custIns = [];
  custIns.push(`총 ${topCusts.length}개 거래처, 총 매출 ${_m(custTotal)}`);
  custIns.push(`TOP5 집중도 ${custConcentration}% — ${custConcentration > 80 ? '특정 거래처 의존도 높음, 분산 필요' : '적정 수준'}`);
  if (topCusts.length > 0) custIns.push(`최대: ${topCusts[0].name} (${_m(topCusts[0].amt)})`);
  addInsights(s11, custIns, 0.85 + 0.3 + custDataRows.length * 0.36 + 0.2);

  // ══════════════════════════════════════���════════════════════
  // 13-14. 타계정 (2 slides)
  // ════════════════════���═════════════════════════════════���════
  setStatus('타계정 분석...', 88);

  const s12 = pres.addSlide();
  s12.background = { color: C.white };
  s12.addText('📑 타계정 내역', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s12.addText(`${yr}년 총 ${_m(taTotal)} / ${_f(taQty)}건`, { x: 5.5, y: 0.3, w: 4, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Monthly TA trend
  const taLabels = Object.keys(taMonthly).map(m => m.slice(5) + '월');
  const taValues = Object.values(taMonthly).map(v => Math.round(v / 1e6 * 100) / 100);
  if (taLabels.length > 1) {
    s12.addChart(pres.charts.BAR, [{
      name: '타계정(백만)', labels: taLabels, values: taValues
    }], {
      x: 0.4, y: 0.9, w: 5.2, h: 2.8, barDir: 'col',
      showTitle: true, title: '월별 타계정 추이', titleColor: C.sub, titleFontSize: 11,
      chartColors: [C.orange], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // Account breakdown (right)
  s12.addText('계정별 집계', { x: 5.8, y: 0.9, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.sub, margin: 0 });
  if (taAcctList.length > 0) {
    const taHdr = [
      { text: '계��', options: { bold: true, fill: { color: C.navy }, color: C.white } },
      { text: '건수', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
      { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
    ];
    const taRows = taAcctList.slice(0, 7).map(([ac, v]) => [
      _shortNm(ac, 10), { text: _f(v.qty), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
    ]);
    s12.addTable([taHdr, ...taRows], { x: 5.8, y: 1.3, w: 3.8, h: 0.3 + taRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.5, 0.9, 1.4] });
  }

  const taIns = [];
  taIns.push(`${yr}년 타계정 총 ${_m(taTotal)} (${_f(taQty)}건, ${taAcctList.length}개 계정)`);
  if (taAcctList.length > 0) taIns.push(`최대 계정: ${taAcctList[0][0]} — ${_m(taAcctList[0][1].amt)} (${taTotal > 0 ? Math.round(taAcctList[0][1].amt / taTotal * 100) : 0}%)`);
  if (taTotal > 50000000) taIns.push(`⚠️ 타계정 규모 ${_m(taTotal)} — 정리/회수 계획 수립 권장`);
  addInsights(s12, taIns, 4.1);

  // ═══════════════════════════════════════════════════════════
  // 15. 공급사 관리 (구매측)
  // ═════════════════════════════════��═════════════════════════
  setStatus('공급사 관리...', 92);

  const s13 = pres.addSlide();
  s13.background = { color: C.white };
  s13.addText('🏭 공급사 관리 (구매)', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s13.addText(`활동 ${poVendors.length}개사 | 납기완료율 ${compRate}%`, { x: 5, y: 0.3, w: 4.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Vendor detail table
  const vHdr = [
    { text: '공급사', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '발주건', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const vRows = topPOVendors.slice(0, 10).map(([nm, v]) => [
    _shortNm(nm, 14), { text: _f(v.cnt), options: { align: 'right' } },
    { text: _m(v.amt), options: { align: 'right' } },
    { text: poAmt > 0 ? Math.round(v.amt / poAmt * 100) + '%' : '-', options: { align: 'right' } }
  ]);
  s13.addTable([vHdr, ...vRows], { x: 0.4, y: 0.85, w: 9.2, h: 0.3 + vRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [3.5, 1.5, 2.0, 2.2] });

  const vIns = [];
  vIns.push(`총 ${poVendors.length}개 공급사, 발주 ${_f(periodPO.length)}건 / ${_m(poAmt)}`);
  const topVShare = topPOVendors.length > 0 && poAmt > 0 ? Math.round(topPOVendors[0][1].amt / poAmt * 100) : 0;
  if (topVShare > 50) vIns.push(`⚠️ ${topPOVendors[0][0]} 집중도 ${topVShare}% — 대안 공급사 확보 검토`);
  vIns.push(`납기완료율 ${compRate}% — ${compRate >= 90 ? '우수' : compRate >= 70 ? '양��' : '미흡, 공급사 관리 강화 필요'}`);
  addInsights(s13, vIns, 0.85 + 0.3 + vRows.length * 0.36 + 0.2);

  // ═══════════════════════════════��═══════════════════════════
  // 16. ACTION ITEMS
  // ════════════════════════════════���═════════════════════���════
  setStatus('액션 아이템 생성...', 95);

  const sAct = pres.addSlide();
  sAct.background = { color: C.white };
  sAct.addText('🎯 Action Items — 주간 후속 조치', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  const actions = [];
  if (cntLow + cntNoStock > 3) actions.push({ pri: '긴급', text: `재고부족 ${cntLow + cntNoStock}건 긴급 발주 검토 (부족 ${cntLow} + 재고없음 ${cntNoStock})`, dept: '구매팀', area: '재고' });
  if (cantProduce > 3) actions.push({ pri: '긴급', text: `생산불가 ${cantProduce}건 — 자재 긴급 확보 필요`, dept: '구매/생산팀', area: 'BOM' });
  if (highCR.length > 3) actions.push({ pri: '높음', text: `고원가율(>70%) ${highCR.length}건 원가절감 방안 수립`, dept: '생산팀', area: '원가' });
  if (cntExcess > 8) actions.push({ pri: '높음', text: `과다재고 ${cntExcess}건 — 재고 최적화/처분 계획`, dept: '물류팀', area: '재고' });
  if (prevSalesTotal > 0 && salesTotal < prevSalesTotal * 0.9) actions.push({ pri: '높음', text: `매출 ${_chg(salesTotal, prevSalesTotal)} 감소 — 원인 분석`, dept: '영업팀', area: '판매' });
  if (compRate < 70) actions.push({ pri: '높음', text: `납기완료율 ${compRate}% — 공급사 관��� 강화`, dept: '구매팀', area: '발주' });
  if (custConcentration > 85) actions.push({ pri: '중간', text: `거래처 집중도 ${custConcentration}% — 신규 거래처 개발`, dept: '영업팀', area: '거래처' });
  if (cntNoDemand > 5) actions.push({ pri: '중간', text: `무수요 ${cntNoDemand}건 — 폐기/이관 검토`, dept: '물류팀', area: '재고' });
  if (taTotal > 50000000) actions.push({ pri: '중간', text: `타계정 ${_m(taTotal)} — 정리/회수 계획 수립`, dept: '재무팀', area: '타계정' });
  if (actions.length === 0) actions.push({ pri: '정보', text: '특이사항 없음 — 현 수준 유지 관리', dept: '전체', area: '-' });

  const actHdr = [
    { text: '우선순위', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '영역', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '조치 사항', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '담당', options: { bold: true, fill: { color: C.navy }, color: C.white } }
  ];
  const actRows = actions.map(a => {
    const pc = a.pri === '긴급' ? C.red : a.pri === '높음' ? C.orange : C.gray;
    return [
      { text: a.pri, options: { color: pc, bold: true } },
      a.area,
      a.text,
      a.dept
    ];
  });
  sAct.addTable([actHdr, ...actRows], { x: 0.4, y: 0.85, w: 9.2, h: 0.3 + actRows.length * 0.45, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.2, 1.0, 5.0, 2.0] });

  sAct.addText(`본 보고서는 ${reportDate} 기준 SCM 대시보드 데이터로 자동 생성되었습니다. | BeaverWorks SCM v12.3`, { x: 0.4, y: 5.1, w: 9.2, h: 0.3, fontSize: 9, color: C.gray, italic: true });

  // ═══════════��════════════════���══════════════════════════════
  // SAVE
  // ════════════════════════════════════════════���══════════════
  setStatus('파일 저장 중...', 98);
  const fileName = `BW_SCM_Weekly_${yr}${S.selMonth === 'all' ? '' : '_' + S.selMonth}_${reportDate.replace(/-/g, '')}.pptx`;
  await pres.writeFile({ fileName });
  setStatus('완료!', 100);
}
