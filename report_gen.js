/**
 * SCM Dashboard — PPT 주간보고서 자동 생성 (PptxGenJS)
 * ====================================================
 * 경영진 주간보고용: 선택월 중심 분석 + 전월 비교 + 탭별 인사이트
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
  const _chg = (cur, prev) => { if (!prev || prev === 0) return ''; const p = ((cur - prev) / prev * 100).toFixed(1); return (p >= 0 ? '+' : '') + p + '%'; };
  // 품목명 풀네임 (truncation 없음) — 여러 소스에서 이름 탐색
  const _nm = (skuId) => {
    // 1) SKU 마스터
    const s = skuMap[skuId];
    if (s && (s.name || s.n)) return (s.name || s.n);
    // 2) 재고 (item_name 필드)
    const inv = D.inv.find(i => i.sku_id === skuId);
    if (inv && (inv.item_name || inv.name)) return (inv.item_name || inv.name);
    // 3) 판매집계 (isd)
    const isd = D.isd.find(i => i.sku_id === skuId);
    if (isd && isd.name) return isd.name;
    // 4) 발주 (po)
    const po = D.po.find(p => p.sku_id === skuId);
    if (po && po.item_name) return po.item_name;
    return skuId;
  };

  // 최근 N개월 출고량 합계
  const _recentOut = (skuId, months) => {
    const io = mio[skuId];
    if (!io) return 0;
    let total = 0;
    const ms = months || 3;
    const allMs = [...M25, ...M26];
    const curIdx = allMs.indexOf(lastMk);
    for (let i = curIdx; i >= Math.max(0, curIdx - ms + 1); i--) {
      const r = io[allMs[i]];
      if (r) total += (r.on || 0);
    }
    return total;
  };

  // ═══ DATA CONTEXT — 선택월 중심 ═══
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
  const prevPeriodStr = prevMk ? prevMk.slice(0, 4) + '.' + parseInt(prevMk.slice(5)) + '월' : '';

  // 전월 비교용 months
  let prevMs = [];
  if (selMs.length === 1 && lastIdx > 0) prevMs = [allMsFull[lastIdx - 1]];
  else if (S.selMonth !== 'all' && S.selMonth.startsWith('Q')) {
    const qMap = { Q1: 'Q4', Q2: 'Q1', Q3: 'Q2', Q4: 'Q3' };
    const pYr = S.selMonth === 'Q1' ? String(+yr - 1) : yr;
    prevMs = getQuarterMonths(qMap[S.selMonth], pYr);
  }

  // 최근 6개월 추이용 months
  const trendMs = allMsFull.slice(Math.max(0, lastIdx - 5), lastIdx + 1);

  setStatus('데이터 수집 중...', 5);

  // ══════════════════════════════════════════
  // DATA AGGREGATION
  // ══════════════════════════════════════════

  // ── 재고 (선택월 기준) ──
  let pIn = 0, pOut = 0, pOther = 0, pInAmt = 0, pOutAmt = 0;
  Object.values(mio).forEach(d => { selMs.forEach(m => { const r = d[m]; if (r) { pIn += r.iq || 0; pOut += r.on || 0; pOther += r.ot || 0; pInAmt += r.ia || 0; pOutAmt += r.oa || 0; } }); });
  let endQty = 0, endAmt = 0;
  Object.values(mio).forEach(d => { if (d[lastMk]) { endQty += d[lastMk].eq || 0; endAmt += d[lastMk].ea || 0; } });
  let beginQty = 0, beginAmt = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { beginQty += d[prevMk].eq || 0; beginAmt += d[prevMk].ea || 0; } });
  else { beginQty = endQty - pIn + pOut + pOther; beginAmt = endAmt - pInAmt + pOutAmt; }

  // 전월 재고 (비교용)
  let prevEndAmt = 0, prevEndQty = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { prevEndAmt += d[prevMk].ea || 0; prevEndQty += d[prevMk].eq || 0; } });

  // Health — 대시보드와 동일한 getEq (fallback: D.inv.qty)
  const getEq = sku => mio[sku] && mio[sku][lastMk] ? mio[sku][lastMk].eq : ((D.inv.find(x => x.sku_id === sku) || {}).qty || 0);
  const activeInv = D.inv.filter(i => (i.qty !== 0 || i.amount !== 0) && i.sku_id.startsWith('B-'));
  let cntNormal = 0, cntLow = 0, cntExcess = 0, cntNoDemand = 0, cntNoStock = 0;
  const lowItems = [], excessItems = [], noDemandItems = [];
  activeInv.forEach(i => {
    const ss = calcSS(i.sku_id); const eq = getEq(i.sku_id);
    const ratio = ss > 0 ? eq / ss : 999;
    if (eq === 0) cntNoStock++;
    else if (ratio < 1) { cntLow++; lowItems.push(i); }
    else if (isExcess(i.sku_id, eq)) { cntExcess++; excessItems.push(i); }
    else if (isNoDemand(i.sku_id)) { cntNoDemand++; noDemandItems.push(i); }
    else cntNormal++;
  });
  const totalActive = activeInv.length;
  const healthScore = totalActive > 0 ? Math.round((cntNormal * 100 + cntExcess * 60 + cntNoDemand * 35 + cntLow * 15) / (totalActive * 100) * 100) : 0;

  // 카테고리별 재고
  const invByCat = {};
  activeInv.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = B_CATS[cat] || cat;
    if (!invByCat[cat]) invByCat[cat] = { nm, qty: 0, amt: 0, cnt: 0 };
    invByCat[cat].cnt++;
    invByCat[cat].qty += getEq(i.sku_id);
    invByCat[cat].amt += (mio[i.sku_id] && mio[i.sku_id][lastMk] ? mio[i.sku_id][lastMk].ea : 0) || 0;
  });
  const invCatList = Object.values(invByCat).sort((a, b) => b.amt - a.amt);

  // ── 판매 (선택월 vs 전월) ──
  const salesTotal = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const salesQty = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].qty : 0), s), 0);
  const cSalesAmt = D.isd.filter(i => i.type.includes('C')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const bSalesAmt = D.isd.filter(i => i.type.includes('B')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const prevSalesTotal = prevMs.length > 0 ? D.isd.reduce((s, i) => prevMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0) : 0;

  // 월별 매출 추이 (최근 6개월)
  const salesMonthly = {};
  trendMs.forEach(m => { salesMonthly[m] = D.isd.reduce((s, i) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0); });

  // 카테고리별 매출
  const salesByCat = {};
  D.isd.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = { ...C_CATS, ...B_CATS }[cat] || cat;
    if (!salesByCat[cat]) salesByCat[cat] = { nm, amt: 0, qty: 0, prevAmt: 0 };
    selMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].amt += d.amt || 0; salesByCat[cat].qty += d.qty || 0; } });
    prevMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].prevAmt += d.amt || 0; } });
  });
  const topSalesCat = Object.values(salesByCat).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);

  // TOP 판매 품목 (품목명 풀네임)
  const topSalesItems = D.isd.map(i => ({
    sku: i.sku_id, name: _nm(i.sku_id),
    amt: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0),
    qty: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].qty : 0), 0),
    prevAmt: prevMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0)
  })).filter(i => i.amt > 0).sort((a, b) => b.amt - a.amt).slice(0, 10);

  // ── 원가 ──
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
    if (sp > 0 || matCost > 0) costItems.push({ sku, name: _nm(sku), sp, matCost, cr, margin, children: children.length });
  });
  costItems.sort((a, b) => b.cr - a.cr);
  const avgCR = costItems.length > 0 ? Math.round(costItems.reduce((s, c) => s + c.cr, 0) / costItems.length) : 0;
  const highCR = costItems.filter(c => c.cr > 70);
  const totalMargin = costItems.reduce((s, c) => s + c.margin, 0);

  // ── 발주 (선택월 기준) ──
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

  // 미입고(진행중) 잔량
  const pendingPO = periodPO.filter(p => p.status === '진행중');
  const pendingAmt = pendingPO.reduce((s, p) => s + (p.amount || 0), 0);
  const pendingQty = pendingPO.reduce((s, p) => s + (p.qty || 0), 0);

  // 품목별 발주 TOP
  const poByItem = {};
  periodPO.forEach(p => {
    const nm = p.item_name || p.sku_id || '-';
    if (!poByItem[nm]) poByItem[nm] = { cnt: 0, amt: 0, qty: 0 };
    poByItem[nm].cnt++; poByItem[nm].amt += p.amount || 0; poByItem[nm].qty += p.qty || 0;
  });
  const topPOItems = Object.entries(poByItem).sort((a, b) => b[1].amt - a[1].amt).slice(0, 7);

  // 전월 발주
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
  let canProduce = 0, cantProduce = 0;
  const cantProduceList = [];
  parents.forEach(sku => {
    if (S.bomDisc && S.bomDisc[sku]) return;
    const children = D.bom.filter(b => b.parent_sku === sku);
    const shortMats = [];
    children.forEach(ch => {
      const eq = getEq(ch.child_sku);
      if (eq < (ch.qty || 1)) shortMats.push({ sku: ch.child_sku, name: _nm(ch.child_sku), need: ch.qty || 1, have: eq });
    });
    if (shortMats.length === 0) canProduce++;
    else { cantProduce++; cantProduceList.push({ parent: sku, parentName: _nm(sku), shortMats }); }
  });

  // ── 거래처 (판매) ──
  const csd = D.csd || [];
  const custTotal = csd.reduce((s, c) => selMs.reduce((ss, m) => ss + (c.monthly[m] ? c.monthly[m].amt : 0), s), 0);
  const topCusts = csd.map(c => ({
    name: c.name, amt: selMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0),
    prevAmt: prevMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0)
  })).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);
  const top5CustAmt = topCusts.slice(0, 5).reduce((s, c) => s + c.amt, 0);
  const custConcentration = custTotal > 0 ? Math.round(top5CustAmt / custTotal * 100) : 0;

  // ── 타계정 ──
  const taAll = D.ta || [];
  const taYr = taAll.filter(t => (t.d || '').startsWith(yr));
  // 연간 합계 (추이 차트용)
  const taYrTotal = taYr.reduce((s, t) => s + (t.a || 0), 0);
  const taYrQty = taYr.reduce((s, t) => s + (t.q || 0), 0);
  // 당월분 (계정별 집계용) — selMs 기준 필터
  const selMsSet = new Set(selMs.map(m => m.slice(5, 7)));
  const taMonth = taYr.filter(t => selMsSet.has((t.d || '').substring(5, 7)));
  const taTotal = taMonth.reduce((s, t) => s + (t.a || 0), 0);
  const taQty = taMonth.reduce((s, t) => s + (t.q || 0), 0);
  const taByAcct = {};
  taMonth.forEach(t => { const ac = t.ac || '(미분류)'; if (!taByAcct[ac]) taByAcct[ac] = { qty: 0, amt: 0 }; taByAcct[ac].qty += t.q || 0; taByAcct[ac].amt += t.a || 0; });
  const taAcctList = Object.entries(taByAcct).sort((a, b) => b[1].amt - a[1].amt);
  const taMs = yr === '2026' ? M26_DATA : M25;
  const taMonthly = {};
  taMs.forEach(m => { taMonthly[m] = taYr.filter(t => (t.d || '').substring(5, 7) === m.slice(5)).reduce((s, t) => s + (t.a || 0), 0); });

  setStatus('슬라이드 생성...', 10);

  // ══════════════════════════════════════════════════════
  // SLIDE HELPERS
  // ══════════════════════════════════════════════════════

  function addInsights(sl, insights, yPos) {
    const y = yPos || 4.3;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.04, fill: { color: 'E2E8F0' } });
    sl.addText([
      { text: '💡 Insights', options: { bold: true, fontSize: 11, breakLine: true } },
      ...insights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < insights.length - 1, fontSize: 10 } }))
    ], { x: 0.5, y: y + 0.1, w: 9, h: 1.2, color: C.dark, valign: 'top' });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 1: 표지
  // ══════════════════════════════════════════════════════
  setStatus('표지...', 12);
  const s1 = pres.addSlide();
  s1.background = { color: C.navy };
  s1.addText('BeaverWorks', { x: 0.8, y: 1.2, w: 8.5, h: 0.9, fontSize: 48, fontFace: 'Arial Black', color: C.white, bold: true });
  s1.addText('SCM 주간 경영보고서', { x: 0.8, y: 2.1, w: 8.5, h: 0.6, fontSize: 28, color: C.ice });
  s1.addText(`보고 기간: ${periodStr}${prevPeriodStr ? ' (비교: ' + prevPeriodStr + ')' : ''}`, { x: 0.8, y: 3.0, w: 8, h: 0.4, fontSize: 14, color: C.gray });
  s1.addText(`생성일: ${reportDate}`, { x: 0.8, y: 3.4, w: 5, h: 0.4, fontSize: 12, color: C.gray });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.3, w: 2.0, h: 0.06, fill: { color: C.accent } });
  s1.addText('비버웍스 공급망관리 시스템 | 자동생성 보고서', { x: 0.8, y: 4.6, w: 6, h: 0.4, fontSize: 11, color: C.gray });

  // ══════════════════════════════════════════════════════
  // SLIDE 2: Executive Summary
  // ══════════════════════════════════════════════════════
  setStatus('Executive Summary...', 15);
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addText('Executive Summary', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 26, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s2.addText(`${periodStr} 핵심지표${prevPeriodStr ? ' | vs ' + prevPeriodStr : ''}`, { x: 0.5, y: 0.75, w: 9, h: 0.3, fontSize: 12, color: C.sub });

  const kpis = [
    { label: '총 매출', val: _m(salesTotal), chg: _chg(salesTotal, prevSalesTotal), color: C.accent },
    { label: '재고자산', val: _m(endAmt), chg: _chg(endAmt, prevEndAmt), color: C.teal },
    { label: '재고건강도', val: healthScore + '점', chg: '', color: healthScore >= 70 ? C.green : C.orange },
    { label: '발주규모', val: _m(poAmt), chg: _chg(poAmt, prevPOAmt), color: C.purple },
    { label: '평균원가율', val: avgCR + '%', chg: '', color: avgCR > 60 ? C.red : C.green },
    { label: '타계정', val: _m(taTotal), chg: '', color: C.orange },
    { label: '미입고 잔량', val: _f(pendingPO.length) + '건', chg: _m(pendingAmt), color: pendingPO.length > 10 ? C.red : C.teal },
    { label: '거래처집중도', val: custConcentration + '%', chg: 'TOP5 비중', color: custConcentration > 80 ? C.orange : C.teal },
  ];
  kpis.forEach((kpi, idx) => {
    const col = idx % 4; const row = Math.floor(idx / 4);
    const x = 0.4 + col * 2.4; const y = 1.15 + row * 1.65;
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 1.45, fill: { color: C.ltGray }, shadow: { type: 'outer', color: '000000', blur: 3, offset: 1, angle: 135, opacity: 0.06 } });
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.05, h: 1.45, fill: { color: kpi.color } });
    s2.addText(kpi.label, { x: x + 0.15, y: y + 0.08, w: 1.9, h: 0.3, fontSize: 10, color: C.sub, margin: 0 });
    s2.addText(kpi.val, { x: x + 0.15, y: y + 0.4, w: 1.9, h: 0.55, fontSize: 22, bold: true, color: C.dark, margin: 0 });
    if (kpi.chg) {
      const chgColor = kpi.chg.includes('+') ? C.green : kpi.chg.includes('-') ? C.red : C.sub;
      s2.addText(kpi.chg, { x: x + 0.15, y: y + 1.0, w: 1.9, h: 0.3, fontSize: 9, color: chgColor, margin: 0 });
    }
  });
  s2.addText(`${periodStr} 기준 | 전월(${prevPeriodStr || '-'}) 대비 변동률 표기`, { x: 0.4, y: 4.9, w: 9.2, h: 0.3, fontSize: 9, color: C.gray, italic: true });

  // ══════════════════════════════════════════════════════
  // SLIDE 3: 재고 현황 — 흐름 + 건강도
  // ══════════════════════════════════════════════════════
  setStatus('재고 현황...', 25);
  const s3 = pres.addSlide();
  s3.background = { color: C.white };
  s3.addText('📦 재고 현황', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s3.addText(`${periodStr} | 전월대비 자산 ${_chg(endAmt, prevEndAmt)}`, { x: 5.5, y: 0.3, w: 4, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Waterfall chart
  s3.addChart(pres.charts.BAR, [{
    name: '금액(백만)', labels: ['기초', '입고', '출고', '타계정', '기말'],
    values: [beginAmt, pInAmt, pOutAmt, Math.max(0, beginAmt + pInAmt - pOutAmt - endAmt), endAmt].map(v => Math.round(v / 1e6))
  }], {
    x: 0.4, y: 0.9, w: 5.0, h: 2.8, barDir: 'col',
    showTitle: true, title: '재고 흐름 (백만원)', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.accent], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  // Health gauge (반원 게이지)
  const gaugeColor = healthScore >= 70 ? C.green : healthScore >= 40 ? C.orange : C.red;
  s3.addChart(pres.charts.DOUGHNUT, [{
    name: '건강도', labels: ['건강도', '잔여', '하단(숨김)'],
    values: [healthScore, 100 - healthScore, 100]
  }], {
    x: 5.6, y: 0.7, w: 4.0, h: 3.2,
    showTitle: false,
    chartColors: [gaugeColor, 'E2E8F0', C.white],
    showPercent: false, showValue: false, showLegend: false,
    dataLabelPosition: 'none',
  });
  // 게이지 중심 텍스트
  s3.addText('건강도', { x: 6.5, y: 2.1, w: 2.2, h: 0.3, fontSize: 11, color: C.sub, align: 'center', margin: 0 });
  s3.addText(`${healthScore}/100`, { x: 6.5, y: 2.35, w: 2.2, h: 0.5, fontSize: 26, bold: true, color: gaugeColor, align: 'center', margin: 0 });
  // 범례 텍스트
  s3.addText(`●정상 ${cntNormal}  ●부족 ${cntLow}  ●재고없음 ${cntNoStock}  ●과다 ${cntExcess}  ●무수요 ${cntNoDemand}`, {
    x: 5.6, y: 3.55, w: 4.0, h: 0.3, fontSize: 8, color: C.sub, align: 'center', margin: 0
  });

  const invIns = [];
  invIns.push(`${periodStr} 기말재고 ${_m(endAmt)} (${_f(endQty)}개)${prevEndAmt ? ' ← 전월 ' + _m(prevEndAmt) : ''}`);
  invIns.push(`입고 ${_m(pInAmt)} / 출고 ${_m(pOutAmt)} — 회전율 ${endAmt > 0 ? (pOutAmt / endAmt * 100).toFixed(0) + '%' : '-'}`);
  if (cntLow + cntNoStock > 0) invIns.push(`⚠️ 부족 ${cntLow}건 + 재고없음 ${cntNoStock}건 → 긴급발주 대상`);
  addInsights(s3, invIns, 4.0);

  // ══════════════════════════════════════════════════════
  // SLIDE 4: 재고 상세 — 위험품목 (현재고 + 최근출고)
  // ══════════════════════════════════════════════════════
  setStatus('재고 위험품목...', 30);
  const s4 = pres.addSlide();
  s4.background = { color: C.white };
  s4.addText('📦 재고 상세 — 위험품목 & 카테고리', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // 부족 품목: 품목명(풀네임) + 현재고 + 최근3개월 출고
  s4.addText(`⚠️ 재고부족 품목 (현재고 < 안전재고) — ${cntLow}건`, { x: 0.4, y: 0.85, w: 5, h: 0.3, fontSize: 11, bold: true, color: C.red, margin: 0 });
  const riskHdr = [
    { text: '품목명', options: { bold: true, fill: { color: C.red }, color: C.white } },
    { text: '현재고', options: { bold: true, fill: { color: C.red }, color: C.white, align: 'right' } },
    { text: '최근3개월 출고', options: { bold: true, fill: { color: C.red }, color: C.white, align: 'right' } }
  ];
  const riskRows = lowItems.slice(0, 7).map(i => {
    const eq = getEq(i.sku_id);
    const recentOut = _recentOut(i.sku_id, 3);
    return [
      _nm(i.sku_id),
      { text: _f(eq) + '개', options: { align: 'right' } },
      { text: _f(recentOut) + '개', options: { align: 'right', color: C.orange } }
    ];
  });
  if (riskRows.length > 0) {
    s4.addTable([riskHdr, ...riskRows], { x: 0.4, y: 1.15, w: 5.5, h: 0.3 + riskRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.8, 1.2, 1.5] });
  }

  // 카테고리별 재고 (right)
  s4.addText('카테고리별 재고', { x: 6.1, y: 0.85, w: 3.5, h: 0.3, fontSize: 11, bold: true, color: C.sub, margin: 0 });
  const catHdr = [
    { text: '카테고리', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '수량', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const catRows = invCatList.slice(0, 7).map(c => [
    c.nm, { text: _f(c.qty), options: { align: 'right' } }, { text: _m(c.amt), options: { align: 'right' } }
  ]);
  s4.addTable([catHdr, ...catRows], { x: 6.1, y: 1.15, w: 3.5, h: 0.3 + catRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.4, 0.9, 1.2] });

  // 과다재고 (하단)
  if (excessItems.length > 0) {
    const exTop = excessItems.slice(0, 4).map(i => `${_nm(i.sku_id)} (현재고 ${_f(getEq(i.sku_id))}개, 최근출고 ${_f(_recentOut(i.sku_id, 3))}개)`);
    s4.addText([
      { text: `📈 과다재고 TOP ${Math.min(excessItems.length, 4)}건 (6개월분 초과)`, options: { bold: true, breakLine: true, fontSize: 10, color: C.orange } },
      ...exTop.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < exTop.length - 1, fontSize: 9 } }))
    ], { x: 0.4, y: 1.15 + 0.3 + riskRows.length * 0.36 + 0.3, w: 9.2, h: 1.5, color: C.dark, valign: 'top' });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 5: 판매 분석 — 추이 + 카테고리
  // ══════════════════════════════════════════════════════
  setStatus('판매 분석...', 40);
  const s5 = pres.addSlide();
  s5.background = { color: C.white };
  s5.addText('💰 판매 분석', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s5.addText(`${periodStr} ${_m(salesTotal)}${prevSalesTotal > 0 ? ' | 전월대비 ' + _chg(salesTotal, prevSalesTotal) : ''}`, { x: 4.5, y: 0.3, w: 5, h: 0.35, fontSize: 11, color: prevSalesTotal > 0 && salesTotal >= prevSalesTotal ? C.green : C.red, align: 'right' });

  // Line chart (최근 6개월)
  const tLabels = Object.keys(salesMonthly).map(m => m.slice(5) + '월');
  const tValues = Object.values(salesMonthly).map(v => Math.round(v / 1e6 * 100) / 100);
  if (tLabels.length > 1) {
    s5.addChart(pres.charts.LINE, [{ name: '매출(백만)', labels: tLabels, values: tValues }], {
      x: 0.4, y: 0.9, w: 5.5, h: 2.8,
      showTitle: true, title: '최근 6개월 매출 추이', titleColor: C.sub, titleFontSize: 11,
      lineSize: 3, lineSmooth: true, chartColors: [C.accent],
      showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // Category doughnut
  if (topSalesCat.length > 0) {
    s5.addChart(pres.charts.DOUGHNUT, [{
      name: '카테고리', labels: topSalesCat.slice(0, 6).map(c => c.nm),
      values: topSalesCat.slice(0, 6).map(c => Math.round(c.amt / 1e6))
    }], {
      x: 6.1, y: 0.9, w: 3.6, h: 2.8,
      showTitle: true, title: '카테고리별 매출', titleColor: C.sub, titleFontSize: 11,
      chartColors: ['2563EB', '0E7490', '15803D', '7E22CE', 'C2410C', '64748B'],
      showPercent: true, showLegend: true, legendPos: 'b', legendFontSize: 8
    });
  }

  const sIns = [];
  sIns.push(`${periodStr} 총 매출 ${_m(salesTotal)} — 완제품 ${Math.round(cSalesAmt / (salesTotal || 1) * 100)}% / 원재료 ${Math.round(bSalesAmt / (salesTotal || 1) * 100)}%`);
  if (prevSalesTotal > 0) sIns.push(`전월(${prevPeriodStr}) ${_m(prevSalesTotal)} → 당월 ${_m(salesTotal)} (${_chg(salesTotal, prevSalesTotal)})`);
  if (topSalesCat.length > 0) sIns.push(`주력: ${topSalesCat[0].nm} ${_m(topSalesCat[0].amt)} (전체 ${salesTotal > 0 ? Math.round(topSalesCat[0].amt / salesTotal * 100) : 0}%)`);
  addInsights(s5, sIns, 4.0);

  // ══════════════════════════════════════════════════════
  // SLIDE 6: 판매 TOP 품목 (풀네임)
  // ══════════════════════════════════════════════════════
  setStatus('판매 TOP 품목...', 48);
  const s6 = pres.addSlide();
  s6.background = { color: C.white };
  s6.addText('💰 판매 TOP 10 품목', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s6.addText(`${periodStr} 기준 | 전월 비교`, { x: 0.5, y: 0.72, w: 9, h: 0.25, fontSize: 11, color: C.sub });

  const itemHdr = [
    { text: '#', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '품목명', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '수량', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: 'vs전월', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const itemRows = topSalesItems.map((it, i) => {
    const chg = _chg(it.amt, it.prevAmt);
    const chgColor = chg.includes('+') ? C.green : chg.includes('-') ? C.red : C.sub;
    return [
      String(i + 1), it.name,
      { text: _f(it.qty), options: { align: 'right' } },
      { text: _m(it.amt), options: { align: 'right' } },
      { text: salesTotal > 0 ? Math.round(it.amt / salesTotal * 100) + '%' : '-', options: { align: 'right' } },
      { text: chg || '-', options: { align: 'right', color: chgColor } }
    ];
  });
  s6.addTable([itemHdr, ...itemRows], { x: 0.3, y: 1.0, w: 9.4, h: 0.3 + itemRows.length * 0.38, fontSize: 9.5, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [0.4, 3.5, 0.9, 1.2, 0.8, 1.2] });

  // ══════════════════════════════════════════════════════
  // SLIDE 7: 원가 분석
  // ══════════════════════════════════════════════════════
  setStatus('원가 분석...', 55);
  const s7 = pres.addSlide();
  s7.background = { color: C.white };
  s7.addText('💹 원가 분석', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s7.addText(`${costItems.length}개 품목 | 평균 원가율 ${avgCR}%`, { x: 5.5, y: 0.3, w: 4, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Distribution chart
  const crBands = [
    { label: '~30%', cnt: costItems.filter(c => c.cr > 0 && c.cr <= 30).length },
    { label: '30~50%', cnt: costItems.filter(c => c.cr > 30 && c.cr <= 50).length },
    { label: '50~70%', cnt: costItems.filter(c => c.cr > 50 && c.cr <= 70).length },
    { label: '70%~', cnt: costItems.filter(c => c.cr > 70).length },
  ];
  s7.addChart(pres.charts.BAR, [{
    name: '품목수', labels: crBands.map(b => b.label), values: crBands.map(b => b.cnt)
  }], {
    x: 0.4, y: 0.9, w: 4.5, h: 2.8, barDir: 'col',
    showTitle: true, title: '원가율 구간별 분포', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.green, C.teal, C.yellow, C.red],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  // Summary table
  const costSummary = [
    ['평균 원가율', avgCR + '%'],
    ['고위험(>70%)', highCR.length + '건'],
    ['저원가(<30%)', costItems.filter(c => c.cr > 0 && c.cr <= 30).length + '건'],
    ['총 마진합계', _m(totalMargin)],
    ['최고 원가율', costItems.length > 0 ? costItems[0].cr + '%' : '-'],
    ['최고 원가 품목', costItems.length > 0 ? costItems[0].name : '-'],
  ];
  s7.addTable([
    [{ text: '지표', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '값', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
    ...costSummary.map(r => [r[0], { text: r[1], options: { align: 'right' } }])
  ], { x: 5.3, y: 0.9, w: 4.3, h: 2.6, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.0, 2.3] });

  const cIns = [];
  cIns.push(`평균 원가율 ${avgCR}% — ${avgCR > 50 ? '마진 관리 주의 구간' : '수익성 양호'}`);
  if (highCR.length > 0) cIns.push(`고위험 ${highCR.length}건: ${highCR.slice(0, 2).map(c => c.name).join(', ')}`);
  addInsights(s7, cIns, 4.0);

  // ══════════════════════════════════════════════════════
  // SLIDE 8: 원가 상세 — 품목별 (풀네임)
  // ══════════════════════════════════════════════════════
  setStatus('원가 상세...', 60);
  const s8 = pres.addSlide();
  s8.background = { color: C.white };
  s8.addText('💹 원가 상세 — 품목별 원가율 & 마진', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 20, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  const crHdr = [
    { text: '품목명', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '판매가', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '원가', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '원가율', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '마진', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const crRows = costItems.slice(0, 12).map(c => [
    c.name,
    { text: _f(c.sp), options: { align: 'right' } },
    { text: _f(c.matCost), options: { align: 'right' } },
    { text: c.cr + '%', options: { align: 'right', color: c.cr > 70 ? C.red : c.cr > 50 ? C.orange : C.green, bold: c.cr > 70 } },
    { text: _f(c.margin), options: { align: 'right', color: c.margin < 0 ? C.red : C.dark } }
  ]);
  s8.addTable([crHdr, ...crRows], { x: 0.3, y: 0.85, w: 9.4, h: 0.3 + crRows.length * 0.36, fontSize: 9.5, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [3.5, 1.4, 1.4, 1.1, 2.0] });

  // ══════════════════════════════════════════════════════
  // SLIDE 9: 발주 관리
  // ══════════════════════════════════════════════════════
  setStatus('발주 분석...', 68);
  const s9 = pres.addSlide();
  s9.background = { color: C.white };
  s9.addText('📋 발주 관리', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s9.addText(`${periodStr} | ${_f(periodPO.length)}건 ${_m(poAmt)}${prevPOAmt > 0 ? ' (전월 ' + _chg(poAmt, prevPOAmt) + ')' : ''}`, { x: 4, y: 0.3, w: 5.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  if (topPOVendors.length > 0) {
    s9.addChart(pres.charts.BAR, [{
      name: '발주액(백만)', labels: topPOVendors.map(v => v[0]),
      values: topPOVendors.map(v => Math.round(v[1].amt / 1e6))
    }], {
      x: 0.4, y: 0.9, w: 5.0, h: 2.8, barDir: 'bar',
      showTitle: true, title: 'TOP 공급사별 발주금액', titleColor: C.sub, titleFontSize: 11,
      chartColors: [C.purple], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  // 미입고 현황 + 발주 품목 TOP
  s9.addText(`📌 미입고 잔량 — ${_f(pendingPO.length)}건 / ${_m(pendingAmt)}`, { x: 5.7, y: 0.9, w: 3.9, h: 0.3, fontSize: 10, bold: true, color: C.red, margin: 0 });
  if (topPOItems.length > 0) {
    s9.addText('품목별 발주 TOP', { x: 5.7, y: 1.25, w: 3.9, h: 0.25, fontSize: 10, bold: true, color: C.sub, margin: 0 });
    s9.addTable([
      [{ text: '품목명', options: { bold: true, fill: { color: C.navy }, color: C.white } },
       { text: '수량', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
       { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
      ...topPOItems.map(([nm, v]) => [
        nm, { text: _f(v.qty), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
      ])
    ], { x: 5.7, y: 1.5, w: 3.9, h: 0.3 + topPOItems.length * 0.34, fontSize: 9, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.7, 0.8, 1.4] });
  }

  const poIns = [];
  poIns.push(`${periodStr} 발주 ${_f(periodPO.length)}건, ${_m(poAmt)}${prevPOAmt > 0 ? ' (전월 ' + _m(prevPOAmt) + ', ' + _chg(poAmt, prevPOAmt) + ')' : ''}`);
  poIns.push(`미입고 ${_f(pendingPO.length)}건 (${_m(pendingAmt)}) — ${pendingPO.length > 10 ? '⚠️ 지연 관리 필요' : '정상 범위'}`);
  if (topPOVendors.length > 0) poIns.push(`최대: ${topPOVendors[0][0]} (${_m(topPOVendors[0][1].amt)}, 비중 ${poAmt > 0 ? Math.round(topPOVendors[0][1].amt / poAmt * 100) : 0}%)`);
  addInsights(s9, poIns, 4.1);

  // ══════════════════════════════════════════════════════
  // SLIDE 10: BOM 관리 & 생산가능성
  // ══════════════════════════════════════════════════════
  setStatus('BOM 분석...', 75);
  const s10 = pres.addSlide();
  s10.background = { color: C.white };
  s10.addText('🔧 BOM 관리 & 생산가능성', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });

  // BOM summary
  s10.addTable([
    [{ text: '지표', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '값', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
    ['BOM 총 관계', { text: _f(bomCnt) + '건', options: { align: 'right' } }],
    ['완제품 (Parent)', { text: _f(bomParentCnt) + '종', options: { align: 'right' } }],
    ['원재료 (Child)', { text: _f(bomMatCnt) + '종', options: { align: 'right' } }],
    ['즉시 생산가능', { text: _f(canProduce) + '종', options: { align: 'right', color: C.green } }],
    ['자재부족 (생산불가)', { text: _f(cantProduce) + '종', options: { align: 'right', color: C.red } }],
    ['생산가능률', { text: (bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) : 0) + '%', options: { align: 'right' } }],
  ], { x: 0.4, y: 0.9, w: 4.5, h: 2.8, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.5, 2.0] });

  // Feasibility pie
  s10.addChart(pres.charts.PIE, [{
    name: '생산가능성', labels: ['생산가능', '자재부족'], values: [canProduce, cantProduce]
  }], {
    x: 5.5, y: 0.9, w: 4.0, h: 2.5,
    showTitle: true, title: '생산가능 현황', titleColor: C.sub, titleFontSize: 11,
    chartColors: [C.green, C.red], showPercent: true, showLegend: true, legendPos: 'b'
  });

  // 생산불가 품목 리스트 (풀네임)
  if (cantProduceList.length > 0) {
    const cpTxt = cantProduceList.slice(0, 3).map(cp => `${cp.parentName}: 부족자재 ${cp.shortMats.slice(0, 2).map(m => m.name).join(', ')}`);
    s10.addText([
      { text: '⚠️ 생산불가 품목 (자재부족)', options: { bold: true, breakLine: true, fontSize: 10, color: C.red } },
      ...cpTxt.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < cpTxt.length - 1, fontSize: 9 } }))
    ], { x: 5.5, y: 3.5, w: 4.1, h: 1.3, color: C.dark, valign: 'top' });
  }

  const bomIns = [];
  bomIns.push(`${bomParentCnt}개 완제품 중 ${canProduce}개 즉시 생산가능 (${bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) : 0}%)`);
  if (cantProduce > 0) bomIns.push(`⚠️ ${cantProduce}개 자재부족 — 해당 원재료 긴급 발주 필요`);
  addInsights(s10, bomIns, 4.2);

  // ══════════════════════════════════════════════════════
  // SLIDE 11: 거래처 분석 (판매)
  // ══════════════════════════════════════════════════════
  setStatus('거래처 분석...', 82);
  const s11 = pres.addSlide();
  s11.background = { color: C.white };
  s11.addText('🏭 거래처 분석 (판매)', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s11.addText(`TOP5 집중도 ${custConcentration}% | ${periodStr}`, { x: 5, y: 0.3, w: 4.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  const custHdr = [
    { text: '#', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '거래처명', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '당월 매출', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: 'vs전월', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const custRows = topCusts.slice(0, 10).map((c, i) => {
    const share = custTotal > 0 ? Math.round(c.amt / custTotal * 100) : 0;
    const chg = _chg(c.amt, c.prevAmt);
    const chgColor = chg.includes('+') ? C.green : chg.includes('-') ? C.red : C.sub;
    return [
      String(i + 1), c.name,
      { text: _m(c.amt), options: { align: 'right' } },
      { text: share + '%', options: { align: 'right' } },
      { text: chg || '-', options: { align: 'right', color: chgColor } }
    ];
  });
  s11.addTable([custHdr, ...custRows], { x: 0.3, y: 0.85, w: 9.4, h: 0.3 + custRows.length * 0.36, fontSize: 9.5, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [0.4, 3.5, 1.8, 1.2, 2.5] });

  const custIns = [];
  custIns.push(`${periodStr} 활동 거래처 ${topCusts.length}개, 총 ${_m(custTotal)}`);
  custIns.push(`TOP5 집중도 ${custConcentration}% — ${custConcentration > 80 ? '⚠️ 의존도 높음' : '적정 분산'}`);
  addInsights(s11, custIns, 0.85 + 0.3 + custRows.length * 0.36 + 0.15);

  // ══════════════════════════════════════════════════════
  // SLIDE 12: 타계정
  // ══════════════════════════════════════════════════════
  setStatus('타계정...', 88);
  const s12 = pres.addSlide();
  s12.background = { color: C.white };
  s12.addText('📑 타계정 내역', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s12.addText(`${periodStr} ${_m(taTotal)} / ${_f(taQty)}건  (${yr}년 누계 ${_m(taYrTotal)})`, { x: 4.0, y: 0.3, w: 5.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  // Monthly trend
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

  // Account breakdown
  s12.addText(`계정별 집계 (${periodStr})`, { x: 5.8, y: 0.9, w: 4, h: 0.3, fontSize: 11, bold: true, color: C.sub, margin: 0 });
  if (taAcctList.length > 0) {
    s12.addTable([
      [{ text: '계정', options: { bold: true, fill: { color: C.navy }, color: C.white } },
       { text: '건수', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
       { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }],
      ...taAcctList.slice(0, 7).map(([ac, v]) => [
        ac, { text: _f(v.qty), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
      ])
    ], { x: 5.8, y: 1.2, w: 3.8, h: 0.3 + Math.min(taAcctList.length, 7) * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.5, 0.9, 1.4] });
  }

  const taIns = [];
  taIns.push(`${periodStr} 타계정 ${_m(taTotal)} (${_f(taQty)}건) — ${yr}년 누계 ${_m(taYrTotal)}`);
  if (taAcctList.length > 0) taIns.push(`당월 최대: ${taAcctList[0][0]} — ${_m(taAcctList[0][1].amt)} (${taTotal > 0 ? Math.round(taAcctList[0][1].amt / taTotal * 100) : 0}%)`);
  if (taYrTotal > 50000000) taIns.push(`⚠️ 연간 누계 ${_m(taYrTotal)} — 정리/회수 계획 수립 필요`);
  addInsights(s12, taIns, 4.1);

  // ══════════════════════════════════════════════════════
  // SLIDE 13: 공급사 관리 (구매측)
  // ══════════════════════════════════════════════════════
  setStatus('공급사 관리...', 92);
  const s13 = pres.addSlide();
  s13.background = { color: C.white };
  s13.addText('🏭 공급사 관리 (구매)', { x: 0.5, y: 0.25, w: 6, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s13.addText(`${poVendors.length}개사 | 납기 ${compRate}%`, { x: 6, y: 0.3, w: 3.5, h: 0.35, fontSize: 11, color: C.sub, align: 'right' });

  const vHdr = [
    { text: '공급사', options: { bold: true, fill: { color: C.navy }, color: C.white } },
    { text: '건수', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '금액', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } },
    { text: '비중', options: { bold: true, fill: { color: C.navy }, color: C.white, align: 'right' } }
  ];
  const vRows = topPOVendors.map(([nm, v]) => [
    nm, { text: _f(v.cnt), options: { align: 'right' } },
    { text: _m(v.amt), options: { align: 'right' } },
    { text: poAmt > 0 ? Math.round(v.amt / poAmt * 100) + '%' : '-', options: { align: 'right' } }
  ]);
  s13.addTable([vHdr, ...vRows], { x: 0.3, y: 0.85, w: 9.4, h: 0.3 + vRows.length * 0.36, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [3.5, 1.5, 2.2, 2.2] });

  const vIns = [];
  vIns.push(`${periodStr} ${poVendors.length}개 공급사, ${_f(periodPO.length)}건 / ${_m(poAmt)}`);
  const topVShare = topPOVendors.length > 0 && poAmt > 0 ? Math.round(topPOVendors[0][1].amt / poAmt * 100) : 0;
  if (topVShare > 50) vIns.push(`⚠️ ${topPOVendors[0][0]} 비중 ${topVShare}% — 리스크 분산 필요`);
  vIns.push(`납기완료율 ${compRate}% — ${compRate >= 90 ? '우수' : compRate >= 70 ? '양호' : '⚠️ 미흡'}`);
  addInsights(s13, vIns, 0.85 + 0.3 + vRows.length * 0.36 + 0.15);

  // ══════════════════════════════════════════════════════
  // SLIDE 14: ACTION ITEMS
  // ══════════════════════════════════════════════════════
  setStatus('액션 아이템...', 95);
  const sAct = pres.addSlide();
  sAct.background = { color: C.white };
  sAct.addText('🎯 Action Items — 주간 후속 조치', { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  sAct.addText(`${periodStr} 기준 자동 도출`, { x: 0.5, y: 0.72, w: 9, h: 0.25, fontSize: 11, color: C.sub });

  const actions = [];
  if (cntLow + cntNoStock > 3) actions.push({ pri: '긴급', text: `재고부족 ${cntLow + cntNoStock}건 긴급 발주 (부족${cntLow} + 재고없음${cntNoStock})`, dept: '구매팀', area: '재고' });
  if (cantProduce > 3) actions.push({ pri: '긴급', text: `생산불가 ${cantProduce}건 — BOM 자재 긴급 확보`, dept: '구매/생산', area: 'BOM' });
  if (highCR.length > 3) actions.push({ pri: '높음', text: `고원가율(>70%) ${highCR.length}건 — 원가절감 방안`, dept: '생산팀', area: '원가' });
  if (cntExcess > 8) actions.push({ pri: '높음', text: `과다재고 ${cntExcess}건 — 최적화/처분`, dept: '물류팀', area: '재고' });
  if (prevSalesTotal > 0 && salesTotal < prevSalesTotal * 0.9) actions.push({ pri: '높음', text: `매출 ${_chg(salesTotal, prevSalesTotal)} 감소 — 원인분석`, dept: '영업팀', area: '판매' });
  if (compRate < 70) actions.push({ pri: '높음', text: `납기완료율 ${compRate}% — 공급사 관리 강화`, dept: '구매팀', area: '발주' });
  if (custConcentration > 85) actions.push({ pri: '중간', text: `거래처 집중도 ${custConcentration}% — 분산 필요`, dept: '영업팀', area: '거래처' });
  if (cntNoDemand > 5) actions.push({ pri: '중간', text: `무수요 ${cntNoDemand}건 — 폐기/이관 검토`, dept: '물류팀', area: '재고' });
  if (taTotal > 50000000) actions.push({ pri: '중간', text: `타계정 ${_m(taTotal)} — 정리/회수`, dept: '재무팀', area: '타계정' });
  if (actions.length === 0) actions.push({ pri: '정보', text: '특이사항 없음 — 현 수준 유지', dept: '전체', area: '-' });

  sAct.addTable([
    [{ text: '우선순위', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '영역', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '조치 사항', options: { bold: true, fill: { color: C.navy }, color: C.white } },
     { text: '담당', options: { bold: true, fill: { color: C.navy }, color: C.white } }],
    ...actions.map(a => [
      { text: a.pri, options: { color: a.pri === '긴급' ? C.red : a.pri === '높음' ? C.orange : C.gray, bold: true } },
      a.area, a.text, a.dept
    ])
  ], { x: 0.3, y: 1.0, w: 9.4, h: 0.3 + actions.length * 0.45, fontSize: 10, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.1, 1.0, 5.3, 2.0] });

  sAct.addText(`${reportDate} 기준 SCM 대시보드 데이터 자동생성 | BeaverWorks v12.3`, { x: 0.3, y: 5.1, w: 9.4, h: 0.3, fontSize: 9, color: C.gray, italic: true });

  // ══════════════════════════════════════════════════════
  // SAVE
  // ══════════════════════════════════════════════════════
  setStatus('저장 중...', 98);
  const fileName = `BW_SCM_Weekly_${yr}${S.selMonth === 'all' ? '' : '_' + S.selMonth}_${reportDate.replace(/-/g, '')}.pptx`;
  await pres.writeFile({ fileName });
  setStatus('완료!', 100);
}
