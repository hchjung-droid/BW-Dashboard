/**
 * SCM Dashboard — PPT 보고서 자동 생성 (PptxGenJS)
 * ================================================
 * 경영진 보고용: 전체 탭 인사이트를 슬라이드로 요약
 */

// eslint-disable-next-line no-unused-vars
async function generateReport() {
  if (typeof PptxGenJS === 'undefined') {
    alert('PPT 라이브러리 로딩 실패. 네트워크를 확인해주세요.');
    return;
  }

  // Show loading overlay
  const overlay = document.createElement('div');
  overlay.id = 'rpt-overlay';
  overlay.innerHTML = `<div style="position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;flex-direction:column">
    <div style="background:#fff;border-radius:12px;padding:32px 48px;text-align:center;box-shadow:0 20px 40px rgba(0,0,0,.3)">
      <div style="font-size:28px;margin-bottom:12px">📊</div>
      <p style="font-size:16px;font-weight:600;color:#1e293b;margin-bottom:8px">보고서 생성 중...</p>
      <p id="rpt-status" style="font-size:12px;color:#64748b">슬라이드 구성 준비</p>
      <div style="width:200px;height:4px;background:#e2e8f0;border-radius:2px;margin-top:12px;overflow:hidden"><div id="rpt-bar" style="width:0%;height:100%;background:#2563eb;transition:width .3s"></div></div>
    </div></div>`;
  document.body.appendChild(overlay);
  const setStatus = (msg, pct) => {
    const s = document.getElementById('rpt-status'); if (s) s.textContent = msg;
    const b = document.getElementById('rpt-bar'); if (b) b.style.width = pct + '%';
  };

  try {
    await _buildReport(setStatus);
  } catch (e) {
    alert('보고서 생성 실패: ' + e.message);
    console.error(e);
  } finally {
    const ov = document.getElementById('rpt-overlay');
    if (ov) ov.remove();
  }
}

async function _buildReport(setStatus) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.author = 'BeaverWorks SCM';
  pres.title = 'BeaverWorks SCM 경영보고서';

  // ====== COLOR PALETTE (Midnight Executive) ======
  const CLR = {
    navy: '1E2761', ice: 'CADCFC', white: 'FFFFFF', accent: '2563EB',
    green: '15803D', red: 'DC2626', gray: '64748B', lightGray: 'F1F5F9',
    darkText: '1E293B', subText: '475569', orange: 'C2410C', purple: '7E22CE',
    teal: '0E7490'
  };

  // ====== HELPER ======
  const _fmt = n => n == null ? '0' : Number(n).toLocaleString('ko-KR');
  const _fmtM = n => { if (!n || n === 0) return '0.00M'; return (n / 1e6).toFixed(2) + 'M'; };
  const _pct = (a, b) => b > 0 ? ((a - b) / b * 100).toFixed(1) + '%' : '-';
  const _monthNm = mk => mk ? mk.slice(2, 4) + '.' + parseInt(mk.slice(5)) + '월' : '';

  // ====== DATA PREP ======
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

  setStatus('데이터 분석 중...', 10);

  // ─── Inventory KPIs ───
  let pIn = 0, pOut = 0, pOther = 0, pInAmt = 0, pOutAmt = 0;
  Object.values(mio).forEach(skuData => {
    selMs.forEach(m => { const d = skuData[m]; if (d) { pIn += d.iq || 0; pOut += d.on || 0; pOther += d.ot || 0; pInAmt += d.ia || 0; pOutAmt += d.oa || 0; } });
  });
  let endQty = 0, endAmt = 0;
  Object.values(mio).forEach(d => { if (d[lastMk]) { endQty += d[lastMk].eq || 0; endAmt += d[lastMk].ea || 0; } });
  let beginQty = 0, beginAmt = 0;
  if (prevMk) { Object.values(mio).forEach(d => { if (d[prevMk]) { beginQty += d[prevMk].eq || 0; beginAmt += d[prevMk].ea || 0; } }); }
  else { beginQty = endQty - pIn + pOut + pOther; beginAmt = endAmt - pInAmt + pOutAmt; }

  // Health scores
  const getEq = sku => mio[sku] && mio[sku][lastMk] ? mio[sku][lastMk].eq : (D.inv.find(x => x.sku_id === sku) || {}).qty || 0;
  const activeInv = D.inv.filter(i => (i.qty !== 0 || i.amount !== 0) && i.sku_id.startsWith('B-'));
  let cntNormal = 0, cntLow = 0, cntExcess = 0, cntNoDemand = 0, cntNoStock = 0;
  activeInv.forEach(i => {
    const ss = calcSS(i.sku_id); const eq = getEq(i.sku_id);
    const ratio = ss > 0 ? eq / ss : 999;
    if (eq === 0) cntNoStock++;
    else if (ratio < 1) cntLow++;
    else if (isExcess(i.sku_id, eq)) cntExcess++;
    else if (isNoDemand(i.sku_id)) cntNoDemand++;
    else cntNormal++;
  });
  const totalActive = activeInv.length;
  const healthW = cntNormal * 100 + cntExcess * 60 + cntNoDemand * 35 + cntLow * 15;
  const healthScore = totalActive > 0 ? Math.round(healthW / (totalActive * 100) * 100) : 0;

  // ─── Sales KPIs ───
  const cItems = D.isd.filter(i => i.type.includes('C'));
  const bItems = D.isd.filter(i => i.type.includes('B'));
  const salesTotal = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const salesQty = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].qty : 0), s), 0);
  const cSalesAmt = cItems.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const bSalesAmt = bItems.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);

  // Monthly sales trend
  const salesMonthly = {};
  selMs.forEach(m => {
    salesMonthly[m] = D.isd.reduce((s, i) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0);
  });

  // MoM growth
  let prevSalesTotal = 0;
  if (selMs.length === 1) {
    const idx = allMsFull.indexOf(selMs[0]);
    if (idx > 0) { const pm = allMsFull[idx - 1]; prevSalesTotal = D.isd.reduce((s, i) => s + (i.monthly[pm] ? i.monthly[pm].amt : 0), 0); }
  }

  // ─── Cost KPIs ───
  const parents = new Set(D.bom.map(b => b.parent_sku));
  let costItems = [];
  parents.forEach(sku => {
    if (S.bomDisc && S.bomDisc[sku]) return;
    const inv = D.inv.find(i => i.sku_id === sku);
    if (!inv) return;
    const tx = D.sales_tx ? D.sales_tx.filter(t => t.sku_id === sku) : [];
    let paidAmt = 0, paidQty = 0;
    tx.forEach(t => { if ((t.amt || 0) > 0) { paidQty += (t.qty || 0); paidAmt += (t.amt || 0); } });
    const sp = paidQty > 0 ? Math.round(paidAmt / paidQty) : 0;
    const children = D.bom.filter(b => b.parent_sku === sku);
    let matCost = 0;
    children.forEach(ch => {
      const cInv = D.inv.find(x => x.sku_id === ch.child_sku);
      matCost += (cInv ? cInv.unit_cost || 0 : 0) * (ch.qty || 1);
    });
    const cr = sp > 0 ? Math.round(matCost / sp * 100) : 0;
    if (sp > 0 || matCost > 0) costItems.push({ sku, name: inv.name, sp, matCost, cr, children: children.length });
  });
  costItems.sort((a, b) => b.cr - a.cr);
  const avgCR = costItems.length > 0 ? Math.round(costItems.reduce((s, c) => s + c.cr, 0) / costItems.length) : 0;
  const highCR = costItems.filter(c => c.cr > 70);
  const lowCR = costItems.filter(c => c.cr > 0 && c.cr <= 30);

  // ─── PO KPIs ───
  const filteredPO = D.po.filter(p => p.vendor !== 'nan' && p.item_name !== 'nan');
  const poCnt = filteredPO.length;
  const poAmt = filteredPO.reduce((s, p) => s + (p.amount || 0), 0);
  const poVendors = [...new Set(filteredPO.map(p => p.vendor))];
  const poByVendor = {};
  filteredPO.forEach(p => {
    if (!poByVendor[p.vendor]) poByVendor[p.vendor] = { cnt: 0, amt: 0 };
    poByVendor[p.vendor].cnt++;
    poByVendor[p.vendor].amt += (p.amount || 0);
  });
  const topVendors = Object.entries(poByVendor).sort((a, b) => b[1].amt - a[1].amt).slice(0, 5);

  // ─── BOM KPIs ───
  const bomCnt = D.bom.length;
  const bomParents = parents.size;
  const usedMats = new Set(D.bom.map(b => b.child_sku));
  const bomMats = usedMats.size;

  // ─── TA (타계정) KPIs ───
  const taAll = D.ta || [];
  const taYr = taAll.filter(t => (t.d || '').startsWith(yr));
  const taTotal = taYr.reduce((s, t) => s + (t.a || 0), 0);
  const taQty = taYr.reduce((s, t) => s + (t.q || 0), 0);
  const taByAcct = {};
  taYr.forEach(t => { const ac = t.ac || '(미분류)'; if (!taByAcct[ac]) taByAcct[ac] = { qty: 0, amt: 0 }; taByAcct[ac].qty += t.q || 0; taByAcct[ac].amt += t.a || 0; });
  const taAcctList = Object.entries(taByAcct).sort((a, b) => b[1].amt - a[1].amt);

  // ─── Supplier KPIs ───
  const csd = D.csd || [];
  const topCusts = csd.slice(0, 10);

  setStatus('슬라이드 생성 중...', 30);

  // ══════════════════════════════════════════════════════
  // SLIDE 1: 표지
  // ══════════════════════════════════════════════════════
  const s1 = pres.addSlide();
  s1.background = { color: CLR.navy };
  s1.addText('BeaverWorks', { x: 0.8, y: 1.5, w: 8.5, h: 0.8, fontSize: 44, fontFace: 'Arial Black', color: CLR.white, bold: true });
  s1.addText('SCM 경영보고서', { x: 0.8, y: 2.3, w: 8.5, h: 0.6, fontSize: 28, fontFace: 'Arial', color: CLR.ice });
  s1.addText(`${yr}년 ${S.selMonth === 'all' ? '연간' : parseInt(S.selMonth) + '월'} | 생성일: ${reportDate}`, { x: 0.8, y: 3.2, w: 8.5, h: 0.4, fontSize: 14, color: CLR.gray });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.2, w: 2.0, h: 0.06, fill: { color: CLR.accent } });
  s1.addText('비버웍스 공급망관리 시스템', { x: 0.8, y: 4.5, w: 5, h: 0.4, fontSize: 12, color: CLR.gray });

  setStatus('대시보드 요약...', 40);

  // ══════════════════════════════════════════════════════
  // SLIDE 2: Executive Summary (핵심 KPI)
  // ══════════════════════════════════════════════════════
  const s2 = pres.addSlide();
  s2.background = { color: CLR.white };
  s2.addText('Executive Summary', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 28, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });
  s2.addText(`${yr}년 ${S.selMonth === 'all' ? '연간' : parseInt(S.selMonth) + '월'} 핵심 경영지표`, { x: 0.5, y: 0.85, w: 9, h: 0.35, fontSize: 13, color: CLR.subText });

  // KPI Cards (2x3 grid)
  const kpis = [
    { label: '매출', value: _fmtM(salesTotal), sub: `완제품 ${_fmtM(cSalesAmt)} / 원재료 ${_fmtM(bSalesAmt)}`, color: CLR.accent },
    { label: '재고자산', value: _fmtM(endAmt), sub: `기말수량 ${_fmt(endQty)}개`, color: CLR.teal },
    { label: '재고건강도', value: healthScore + '점', sub: `정상${cntNormal} 부족${cntLow} 과다${cntExcess}`, color: healthScore >= 70 ? CLR.green : CLR.orange },
    { label: '발주', value: _fmt(poCnt) + '건', sub: `${_fmtM(poAmt)} / ${poVendors.length}개 거래처`, color: CLR.purple },
    { label: '평균 원가율', value: avgCR + '%', sub: `고위험(>70%) ${highCR.length}건`, color: avgCR > 60 ? CLR.red : CLR.green },
    { label: '타계정', value: _fmtM(taTotal), sub: `${_fmt(taQty)}건 / ${taAcctList.length}개 계정`, color: CLR.orange },
  ];

  kpis.forEach((kpi, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const x = 0.5 + col * 3.1;
    const y = 1.5 + row * 1.9;
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.7, fill: { color: CLR.lightGray }, shadow: { type: 'outer', color: '000000', blur: 4, offset: 2, angle: 135, opacity: 0.08 } });
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.7, fill: { color: kpi.color } });
    s2.addText(kpi.label, { x: x + 0.2, y: y + 0.15, w: 2.5, h: 0.35, fontSize: 11, color: CLR.subText, margin: 0 });
    s2.addText(kpi.value, { x: x + 0.2, y: y + 0.5, w: 2.5, h: 0.6, fontSize: 24, bold: true, color: CLR.darkText, margin: 0 });
    s2.addText(kpi.sub, { x: x + 0.2, y: y + 1.15, w: 2.5, h: 0.35, fontSize: 9, color: CLR.gray, margin: 0 });
  });

  setStatus('재고 분석...', 50);

  // ══════════════════════════════════════════════════════
  // SLIDE 3: 재고 현황
  // ══════════════════════════════════════════════════════
  const s3 = pres.addSlide();
  s3.background = { color: CLR.white };
  s3.addText('재고 현황', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });

  // Waterfall data
  s3.addText('재고 흐름 (금액 기준)', { x: 0.5, y: 1.0, w: 4, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  const wfData = [
    ['기초재고', _fmtM(beginAmt)],
    ['입고(+)', _fmtM(pInAmt)],
    ['출고(-)', _fmtM(pOutAmt)],
    ['타계정(-)', _fmtM(beginAmt + pInAmt - pOutAmt - endAmt)],
    ['기말재고', _fmtM(endAmt)]
  ];
  const wfRows = [
    [{ text: '구분', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '금액', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
    ...wfData.map(r => [r[0], { text: r[1], options: { align: 'right' } }])
  ];
  s3.addTable(wfRows, { x: 0.5, y: 1.4, w: 4.3, h: 2.2, fontSize: 12, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.3, 2.0] });

  // Health breakdown (right side)
  s3.addText('재고 상태 분포', { x: 5.2, y: 1.0, w: 4, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  const healthRows = [
    [{ text: '상태', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '건수', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } },
     { text: '비율', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
    ['정상', { text: _fmt(cntNormal), options: { align: 'right' } }, { text: totalActive > 0 ? Math.round(cntNormal / totalActive * 100) + '%' : '-', options: { align: 'right' } }],
    ['부족', { text: _fmt(cntLow), options: { align: 'right', color: CLR.orange } }, { text: totalActive > 0 ? Math.round(cntLow / totalActive * 100) + '%' : '-', options: { align: 'right' } }],
    ['재고없음', { text: _fmt(cntNoStock), options: { align: 'right', color: CLR.red } }, { text: totalActive > 0 ? Math.round(cntNoStock / totalActive * 100) + '%' : '-', options: { align: 'right' } }],
    ['과다', { text: _fmt(cntExcess), options: { align: 'right', color: CLR.orange } }, { text: totalActive > 0 ? Math.round(cntExcess / totalActive * 100) + '%' : '-', options: { align: 'right' } }],
    ['무수요', { text: _fmt(cntNoDemand), options: { align: 'right', color: CLR.purple } }, { text: totalActive > 0 ? Math.round(cntNoDemand / totalActive * 100) + '%' : '-', options: { align: 'right' } }],
  ];
  s3.addTable(healthRows, { x: 5.2, y: 1.4, w: 4.3, h: 2.5, fontSize: 12, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.8, 1.2, 1.3] });

  // Insight text
  const invInsights = [];
  if (cntLow + cntNoStock > 5) invInsights.push(`재고부족 ${cntLow + cntNoStock}건 — 긴급 발주 검토 필요`);
  if (cntExcess > 10) invInsights.push(`과다재고 ${cntExcess}건 — 재고 최적화 필요`);
  if (cntNoDemand > 5) invInsights.push(`무수요 ${cntNoDemand}건 — 폐기/처분 검토`);
  if (invInsights.length === 0) invInsights.push('재고 상태 전반적으로 양호');
  s3.addText([
    { text: 'Key Insights', options: { bold: true, breakLine: true } },
    ...invInsights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < invInsights.length - 1 } }))
  ], { x: 0.5, y: 4.0, w: 9, h: 1.2, fontSize: 12, color: CLR.darkText, valign: 'top' });

  setStatus('매출 분석...', 60);

  // ══════════════════════════════════════════════════════
  // SLIDE 4: 판매 분석
  // ══════════════════════════════════════════════════════
  const s4 = pres.addSlide();
  s4.background = { color: CLR.white };
  s4.addText('판매 분석', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });

  // Sales trend chart
  const salesLabels = selMs.map(m => m.slice(5) + '월');
  const salesValues = selMs.map(m => Math.round((salesMonthly[m] || 0) / 1e6 * 100) / 100);
  if (salesLabels.length > 1) {
    s4.addChart(pres.charts.LINE, [{
      name: '매출(백만)', labels: salesLabels, values: salesValues
    }], {
      x: 0.5, y: 1.0, w: 5.5, h: 3.0,
      showTitle: true, title: '월별 매출 추이 (백만원)',
      lineSize: 3, lineSmooth: true,
      chartColors: [CLR.accent],
      showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: CLR.darkText,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: CLR.gray, valAxisLabelColor: CLR.gray
    });
  }

  // Sales category breakdown (right side)
  const salesByCat = {};
  D.isd.forEach(i => {
    const cat = getSkuCat(i.sku_id); const catNm = { ...C_CATS, ...B_CATS }[cat] || cat;
    if (!salesByCat[cat]) salesByCat[cat] = { catNm, amt: 0 };
    selMs.forEach(m => { const d = i.monthly[m]; if (d) salesByCat[cat].amt += d.amt || 0; });
  });
  const topCats = Object.values(salesByCat).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt).slice(0, 6);
  if (topCats.length > 0) {
    s4.addChart(pres.charts.PIE, [{
      name: '카테고리별', labels: topCats.map(c => c.catNm), values: topCats.map(c => Math.round(c.amt / 1e6))
    }], {
      x: 6.2, y: 1.0, w: 3.5, h: 3.0,
      showTitle: true, title: '카테고리별 매출',
      showPercent: true, showLegend: true, legendPos: 'b',
      chartColors: ['2563EB', '0E7490', '15803D', '7E22CE', 'C2410C', '64748B']
    });
  }

  // Sales insights
  const salesInsights = [];
  salesInsights.push(`총 매출: ${_fmtM(salesTotal)} (완제품 ${Math.round(cSalesAmt / (salesTotal || 1) * 100)}%)`);
  if (prevSalesTotal > 0) {
    const gr = ((salesTotal - prevSalesTotal) / prevSalesTotal * 100).toFixed(1);
    salesInsights.push(`전월 대비: ${gr > 0 ? '+' : ''}${gr}%`);
  }
  if (topCats.length > 0) salesInsights.push(`주력: ${topCats[0].catNm} (${_fmtM(topCats[0].amt)})`);
  s4.addText([
    { text: 'Key Insights', options: { bold: true, breakLine: true } },
    ...salesInsights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < salesInsights.length - 1 } }))
  ], { x: 0.5, y: 4.2, w: 9, h: 1.0, fontSize: 12, color: CLR.darkText, valign: 'top' });

  setStatus('원가 분석...', 70);

  // ══════════════════════════════════════════════════════
  // SLIDE 5: 원가 분석
  // ══════════════════════════════════════════════════════
  const s5 = pres.addSlide();
  s5.background = { color: CLR.white };
  s5.addText('원가 분석', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });

  // Cost ratio distribution chart
  const crBands = [
    { label: '~30%', cnt: costItems.filter(c => c.cr > 0 && c.cr <= 30).length },
    { label: '30~50%', cnt: costItems.filter(c => c.cr > 30 && c.cr <= 50).length },
    { label: '50~70%', cnt: costItems.filter(c => c.cr > 50 && c.cr <= 70).length },
    { label: '70%~', cnt: costItems.filter(c => c.cr > 70).length },
  ];
  s5.addChart(pres.charts.BAR, [{
    name: '품목수', labels: crBands.map(b => b.label), values: crBands.map(b => b.cnt)
  }], {
    x: 0.5, y: 1.0, w: 4.5, h: 2.8, barDir: 'col',
    showTitle: true, title: '원가율 분포',
    chartColors: [CLR.accent],
    showValue: true, dataLabelPosition: 'outEnd',
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: CLR.gray, valAxisLabelColor: CLR.gray
  });

  // Top high-cost items table
  s5.addText('고원가율 품목 TOP 5', { x: 5.3, y: 1.0, w: 4.5, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  const topCost = highCR.slice(0, 5);
  const costRows = [
    [{ text: '품목', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '원가율', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
    ...topCost.map(c => [
      c.name ? (c.name.length > 12 ? c.name.slice(0, 12) + '..' : c.name) : c.sku,
      { text: c.cr + '%', options: { align: 'right', color: c.cr > 80 ? CLR.red : CLR.orange } }
    ])
  ];
  if (costRows.length > 1) {
    s5.addTable(costRows, { x: 5.3, y: 1.4, w: 4.3, h: 0.3 + topCost.length * 0.4, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [3.0, 1.3] });
  }

  // Cost insights
  const costInsights = [];
  costInsights.push(`평균 원가율: ${avgCR}% (${costItems.length}개 품목 분석)`);
  if (highCR.length > 0) costInsights.push(`고위험(>70%) ${highCR.length}건 — 마진 개선 필요`);
  if (lowCR.length > 0) costInsights.push(`저원가(<30%) ${lowCR.length}건 — 수익성 양호`);
  s5.addText([
    { text: 'Key Insights', options: { bold: true, breakLine: true } },
    ...costInsights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < costInsights.length - 1 } }))
  ], { x: 0.5, y: 4.2, w: 9, h: 1.0, fontSize: 12, color: CLR.darkText, valign: 'top' });

  setStatus('발주/BOM 분석...', 80);

  // ══════════════════════════════════════════════════════
  // SLIDE 6: 발주 & BOM
  // ══════════════════════════════════════════════════════
  const s6 = pres.addSlide();
  s6.background = { color: CLR.white };
  s6.addText('발주 현황 & BOM 관리', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });

  // PO by vendor chart
  if (topVendors.length > 0) {
    s6.addChart(pres.charts.BAR, [{
      name: '발주금액(백만)', labels: topVendors.map(v => v[0].length > 8 ? v[0].slice(0, 8) + '..' : v[0]),
      values: topVendors.map(v => Math.round(v[1].amt / 1e6))
    }], {
      x: 0.5, y: 1.0, w: 4.5, h: 2.8, barDir: 'bar',
      showTitle: true, title: 'TOP 5 공급사 (발주금액)',
      chartColors: [CLR.purple],
      showValue: true, dataLabelPosition: 'outEnd',
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: CLR.gray, valAxisLabelColor: CLR.gray
    });
  }

  // BOM summary (right)
  s6.addText('BOM 현황', { x: 5.3, y: 1.0, w: 4.5, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  const bomRows = [
    [{ text: '지표', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '값', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
    ['BOM 관계', { text: _fmt(bomCnt) + '건', options: { align: 'right' } }],
    ['완제품 (Parent)', { text: _fmt(bomParents) + '종', options: { align: 'right' } }],
    ['원재료 (Child)', { text: _fmt(bomMats) + '종', options: { align: 'right' } }],
    ['발주 건수', { text: _fmt(poCnt) + '건', options: { align: 'right' } }],
    ['발주 금액', { text: _fmtM(poAmt), options: { align: 'right' } }],
    ['거래처 수', { text: poVendors.length + '개', options: { align: 'right' } }],
  ];
  s6.addTable(bomRows, { x: 5.3, y: 1.4, w: 4.3, h: 2.8, fontSize: 12, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.5, 1.8] });

  // PO insights
  const poInsights = [];
  poInsights.push(`총 발주 ${_fmt(poCnt)}건, ${_fmtM(poAmt)} (${poVendors.length}개 거래처)`);
  if (topVendors.length > 0) poInsights.push(`최대 거래처: ${topVendors[0][0]} (${_fmtM(topVendors[0][1].amt)})`);
  poInsights.push(`BOM ${bomParents}개 완제품 × ${bomMats}개 원재료 구성`);
  s6.addText([
    { text: 'Key Insights', options: { bold: true, breakLine: true } },
    ...poInsights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < poInsights.length - 1 } }))
  ], { x: 0.5, y: 4.2, w: 9, h: 1.0, fontSize: 12, color: CLR.darkText, valign: 'top' });

  setStatus('거래처/타계정...', 88);

  // ══════════════════════════════════════════════════════
  // SLIDE 7: 거래처 & 타계정
  // ══════════════════════════════════════════════════════
  const s7 = pres.addSlide();
  s7.background = { color: CLR.white };
  s7.addText('거래처 분석 & 타계정', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });

  // Top customers table (left)
  s7.addText('주요 거래처 (판매)', { x: 0.5, y: 1.0, w: 4.5, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  const custRows = [
    [{ text: '거래처', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '매출', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
    ...topCusts.slice(0, 6).map(c => [
      (c.nm || c.name || '').length > 12 ? (c.nm || c.name || '').slice(0, 12) + '..' : (c.nm || c.name || ''),
      { text: _fmtM(c.total || c.amt || 0), options: { align: 'right' } }
    ])
  ];
  if (custRows.length > 1) {
    s7.addTable(custRows, { x: 0.5, y: 1.4, w: 4.3, h: 0.3 + Math.min(topCusts.length, 6) * 0.4, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.8, 1.5] });
  }

  // 타계정 (right)
  s7.addText('타계정 계정별 집계', { x: 5.3, y: 1.0, w: 4.5, h: 0.35, fontSize: 13, color: CLR.subText, bold: true, margin: 0 });
  if (taAcctList.length > 0) {
    const taRows = [
      [{ text: '계정', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
       { text: '건수', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } },
       { text: '금액', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white, align: 'right' } }],
      ...taAcctList.slice(0, 6).map(([ac, v]) => [
        ac.length > 10 ? ac.slice(0, 10) + '..' : ac,
        { text: _fmt(v.qty), options: { align: 'right' } },
        { text: _fmtM(v.amt), options: { align: 'right' } }
      ])
    ];
    s7.addTable(taRows, { x: 5.3, y: 1.4, w: 4.3, h: 0.3 + Math.min(taAcctList.length, 6) * 0.4, fontSize: 11, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [2.0, 1.0, 1.3] });
  } else {
    s7.addText('타계정 데이터 없음', { x: 5.3, y: 1.6, w: 4.3, h: 0.5, fontSize: 12, color: CLR.gray });
  }

  // Combined insights
  const taInsights = [];
  if (topCusts.length > 0) taInsights.push(`최대 거래처: ${topCusts[0].nm || topCusts[0].name || '-'} (${_fmtM(topCusts[0].total || topCusts[0].amt || 0)})`);
  taInsights.push(`타계정 총계: ${_fmtM(taTotal)} (${_fmt(taQty)}건)`);
  if (taAcctList.length > 0) taInsights.push(`최대 계정: ${taAcctList[0][0]} (${_fmtM(taAcctList[0][1].amt)})`);
  s7.addText([
    { text: 'Key Insights', options: { bold: true, breakLine: true } },
    ...taInsights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < taInsights.length - 1 } }))
  ], { x: 0.5, y: 4.2, w: 9, h: 1.0, fontSize: 12, color: CLR.darkText, valign: 'top' });

  setStatus('액션 아이템...', 93);

  // ══════════════════════════════════════════════════════
  // SLIDE 8: Action Items & 결론
  // ══════════════════════════════════════════════════════
  const s8 = pres.addSlide();
  s8.background = { color: CLR.white };
  s8.addText('Action Items', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, fontFace: 'Arial Black', color: CLR.darkText, margin: 0 });
  s8.addText('주요 후속 조치 사항', { x: 0.5, y: 0.85, w: 9, h: 0.35, fontSize: 13, color: CLR.subText });

  const actions = [];
  // Generate contextual action items
  if (cntLow + cntNoStock > 3) actions.push({ pri: '긴급', text: `재고부족 ${cntLow + cntNoStock}건 긴급 발주 검토`, dept: '구매팀' });
  if (cntExcess > 8) actions.push({ pri: '높음', text: `과다재고 ${cntExcess}건 재고 조정/처분 계획 수립`, dept: '물류팀' });
  if (cntNoDemand > 5) actions.push({ pri: '중간', text: `무수요 ${cntNoDemand}건 폐기/이관 검토`, dept: '물류팀' });
  if (highCR.length > 3) actions.push({ pri: '높음', text: `고원가율(>70%) ${highCR.length}건 원가절감 방안 수립`, dept: '생산팀' });
  if (prevSalesTotal > 0 && salesTotal < prevSalesTotal) actions.push({ pri: '높음', text: `매출 전월대비 감소 — 원인 분석 및 대응`, dept: '영업팀' });
  if (taTotal > 50000000) actions.push({ pri: '중간', text: `타계정 ${_fmtM(taTotal)} — 정리/회수 계획 필요`, dept: '재무팀' });
  if (actions.length === 0) actions.push({ pri: '정보', text: '특이사항 없음 — 현 수준 유지', dept: '전체' });

  const actRows = [
    [{ text: '우선순위', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '조치 사항', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } },
     { text: '담당', options: { bold: true, fill: { color: CLR.navy }, color: CLR.white } }],
    ...actions.map(a => {
      const priColor = a.pri === '긴급' ? CLR.red : a.pri === '높음' ? CLR.orange : CLR.gray;
      return [
        { text: a.pri, options: { color: priColor, bold: true } },
        a.text,
        a.dept
      ];
    })
  ];
  s8.addTable(actRows, { x: 0.5, y: 1.3, w: 9, h: 0.4 + actions.length * 0.55, fontSize: 12, border: { type: 'solid', pt: 0.5, color: 'E2E8F0' }, colW: [1.5, 5.5, 2.0] });

  // Footer note
  s8.addText(`본 보고서는 ${reportDate} 기준 SCM 대시보드 데이터로 자동 생성되었습니다.`, { x: 0.5, y: 5.0, w: 9, h: 0.35, fontSize: 10, color: CLR.gray, italic: true });

  setStatus('파일 저장 중...', 97);

  // ══════════════════════════════════════════════════════
  // SAVE
  // ══════════════════════════════════════════════════════
  const fileName = `BW_SCM_Report_${yr}${S.selMonth === 'all' ? '' : '_' + S.selMonth}_${reportDate.replace(/-/g, '')}.pptx`;
  await pres.writeFile({ fileName });

  setStatus('완료!', 100);
}
