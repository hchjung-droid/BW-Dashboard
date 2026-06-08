/**
 * SCM Dashboard — PPT 주간보고서 자동 생성 (PptxGenJS)
 * ====================================================
 * 경영진 주간보고용: 선택월 중심 분석 + 전월 비교 + 탭별 인사이트
 * v3 — 통일 레이아웃 + 5건 피드백 반영
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

  // ═══ 통일 색상 ═══
  const C = {
    navy: '1E2761', ice: 'CADCFC', white: 'FFFFFF', accent: '2563EB',
    green: '15803D', red: 'DC2626', gray: '64748B', ltGray: 'F1F5F9',
    dark: '1E293B', sub: '475569', orange: 'C2410C', purple: '7E22CE',
    teal: '0E7490', yellow: '92400E', bg: 'F8FAFC'
  };

  // ═══ 통일 레이아웃 상수 ═══
  const L = {
    mx: 0.4,           // 좌우 마진
    tw: 9.2,           // 전체 너비 (10 - 0.4*2)
    titleY: 0.2,       // 제목 Y
    titleH: 0.5,       // 제목 높이
    subY: 0.65,        // 부제 Y
    subH: 0.25,        // 부제 높이
    bodyY: 0.95,       // 본문 시작 Y
    insY: 4.25,        // Insights Y (고정)
    splitL: 5.15,      // 2단 좌측 너비
    splitRx: 5.65,     // 2단 우측 X
    splitRw: 3.95,     // 2단 우측 너비
    tblFull: 9.2,      // 전체폭 테이블 너비
  };

  const _f = n => n == null ? '0' : Number(n).toLocaleString('ko-KR');
  const _m = n => { if (!n || n === 0) return '0'; return (n / 1e6).toFixed(2) + 'M'; };
  const _chg = (cur, prev) => { if (!prev || prev === 0) return ''; const p = ((cur - prev) / prev * 100).toFixed(1); return (p >= 0 ? '+' : '') + p + '%'; };

  // 품목명 풀네임 — 여러 소스 탐색
  const _nm = (skuId) => {
    const s = skuMap[skuId];
    if (s && (s.name || s.n)) return (s.name || s.n);
    const inv = D.inv.find(i => i.sku_id === skuId);
    if (inv && (inv.item_name || inv.name)) return (inv.item_name || inv.name);
    const isd = D.isd.find(i => i.sku_id === skuId);
    if (isd && isd.name) return isd.name;
    const po = D.po.find(p => p.sku_id === skuId);
    if (po && po.item_name) return po.item_name;
    return skuId;
  };

  const _recentOut = (skuId, months) => {
    const io = mio[skuId]; if (!io) return 0;
    let total = 0; const ms = months || 3;
    const allMs = [...M25, ...M26];
    const curIdx = allMs.indexOf(lastMk);
    for (let i = curIdx; i >= Math.max(0, curIdx - ms + 1); i--) {
      const r = io[allMs[i]]; if (r) total += (r.on || 0);
    }
    return total;
  };

  // ═══ 통일 슬라이드 헬퍼 ═══
  function addTitle(sl, icon, title, sub, subAlign) {
    sl.background = { color: C.white };
    sl.addText(`${icon} ${title}`, { x: L.mx, y: L.titleY, w: 6, h: L.titleH, fontSize: 22, fontFace: 'Arial Black', color: C.dark, margin: 0 });
    if (sub) sl.addText(sub, { x: L.mx, y: L.subY, w: L.tw, h: L.subH, fontSize: 11, color: C.sub, align: subAlign || 'right', margin: 0 });
  }

  function addInsights(sl, insights) {
    sl.addText([
      { text: '💡 Insights', options: { bold: true, fontSize: 10, breakLine: true, color: C.accent } },
      ...insights.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < insights.length - 1, fontSize: 9.5 } }))
    ], { x: L.mx, y: L.insY, w: L.tw, h: 0.95, color: C.dark, valign: 'top' });
  }

  // 테이블 공통 스타일
  const tblBorder = { type: 'solid', pt: 0.5, color: 'E2E8F0' };
  const hdrStyle = (bgColor) => ({ bold: true, fill: { color: bgColor || C.navy }, color: C.white, fontSize: 9.5 });
  const hdrR = (bgColor) => ({ bold: true, fill: { color: bgColor || C.navy }, color: C.white, align: 'right', fontSize: 9.5 });

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
  const prevPeriodStr = prevMk ? prevMk.slice(0, 4) + '.' + parseInt(prevMk.slice(5)) + '월' : '';

  let prevMs = [];
  if (selMs.length === 1 && lastIdx > 0) prevMs = [allMsFull[lastIdx - 1]];
  else if (S.selMonth !== 'all' && S.selMonth.startsWith('Q')) {
    const qMap = { Q1: 'Q4', Q2: 'Q1', Q3: 'Q2', Q4: 'Q3' };
    const pYr = S.selMonth === 'Q1' ? String(+yr - 1) : yr;
    prevMs = getQuarterMonths(qMap[S.selMonth], pYr);
  }

  const trendMs = allMsFull.slice(Math.max(0, lastIdx - 5), lastIdx + 1);

  setStatus('데이터 수집 중...', 5);

  // ══════════════════════════════════════════
  // DATA AGGREGATION
  // ══════════════════════════════════════════

  // ── 재고 ──
  let pIn = 0, pOut = 0, pOther = 0, pInAmt = 0, pOutAmt = 0;
  Object.values(mio).forEach(d => { selMs.forEach(m => { const r = d[m]; if (r) { pIn += r.iq || 0; pOut += r.on || 0; pOther += r.ot || 0; pInAmt += r.ia || 0; pOutAmt += r.oa || 0; } }); });
  let endQty = 0, endAmt = 0;
  Object.values(mio).forEach(d => { if (d[lastMk]) { endQty += d[lastMk].eq || 0; endAmt += d[lastMk].ea || 0; } });
  let beginQty = 0, beginAmt = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { beginQty += d[prevMk].eq || 0; beginAmt += d[prevMk].ea || 0; } });
  else { beginQty = endQty - pIn + pOut + pOther; beginAmt = endAmt - pInAmt + pOutAmt; }

  let prevEndAmt = 0, prevEndQty = 0;
  if (prevMk) Object.values(mio).forEach(d => { if (d[prevMk]) { prevEndAmt += d[prevMk].ea || 0; prevEndQty += d[prevMk].eq || 0; } });

  // Health — 대시보드와 동일한 getEq
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

  const invByCat = {};
  activeInv.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = B_CATS[cat] || cat;
    if (!invByCat[cat]) invByCat[cat] = { nm, qty: 0, amt: 0, cnt: 0 };
    invByCat[cat].cnt++; invByCat[cat].qty += getEq(i.sku_id);
    invByCat[cat].amt += (mio[i.sku_id] && mio[i.sku_id][lastMk] ? mio[i.sku_id][lastMk].ea : 0) || 0;
  });
  const invCatList = Object.values(invByCat).sort((a, b) => b.amt - a.amt);

  // ── 판매 ──
  const salesTotal = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const salesQty = D.isd.reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].qty : 0), s), 0);
  const cSalesAmt = D.isd.filter(i => i.type.includes('C')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const bSalesAmt = D.isd.filter(i => i.type.includes('B')).reduce((s, i) => selMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0);
  const prevSalesTotal = prevMs.length > 0 ? D.isd.reduce((s, i) => prevMs.reduce((ss, m) => ss + (i.monthly[m] ? i.monthly[m].amt : 0), s), 0) : 0;

  // ★ 선택월 판매된 SKU set (원가 필터용)
  const soldSkus = new Set();
  D.isd.forEach(i => { const amt = selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0); if (amt > 0) soldSkus.add(i.sku_id); });

  const salesMonthly = {};
  trendMs.forEach(m => { salesMonthly[m] = D.isd.reduce((s, i) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0); });

  const salesByCat = {};
  D.isd.forEach(i => {
    const cat = getSkuCat(i.sku_id); const nm = { ...C_CATS, ...B_CATS }[cat] || cat;
    if (!salesByCat[cat]) salesByCat[cat] = { nm, amt: 0, qty: 0, prevAmt: 0 };
    selMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].amt += d.amt || 0; salesByCat[cat].qty += d.qty || 0; } });
    prevMs.forEach(m => { const d = i.monthly[m]; if (d) { salesByCat[cat].prevAmt += d.amt || 0; } });
  });
  const topSalesCat = Object.values(salesByCat).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);

  const topSalesItems = D.isd.map(i => ({
    sku: i.sku_id, name: _nm(i.sku_id),
    amt: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0),
    qty: selMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].qty : 0), 0),
    prevAmt: prevMs.reduce((s, m) => s + (i.monthly[m] ? i.monthly[m].amt : 0), 0)
  })).filter(i => i.amt > 0).sort((a, b) => b.amt - a.amt).slice(0, 10);

  // ── 원가 (★ 대시보드와 동일 로직: 수불부 단가 / 판매단가) ──
  const DATA_MONTHS = [...M25, ...M26_DATA];
  const parents = new Set(D.bom.map(b => b.parent_sku));
  const allPur = [...(D.purch25 || []), ...(D.purch26 || [])].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  // 원가 기준월: 선택월의 마지막 월
  const costRefMk = lastMk;
  let costItems = [];
  parents.forEach(sku => {
    if (S.bomDisc && S.bomDisc[sku]) return;
    if (!soldSkus.has(sku)) return; // ★ 선택월에 판매된 품목만
    const inv = D.inv.find(i => i.sku_id === sku); if (!inv) return;

    // ── 수불부 단가 (uc): 대시보드와 동일 — 기말→입고→출고→직전월 역추적 ──
    const io = mio[sku];
    let uc = 0;
    if (io) {
      // ① 해당월 기말단가 → 입고단가 → 출고단가
      const r = io[costRefMk];
      if (r && r.eq > 0 && r.ea > 0) uc = Math.round(r.ea / r.eq);
      else if (r && r.iq > 0 && r.ia > 0) uc = Math.round(r.ia / r.iq);
      else if (r && (r.on > 0 || r.ot > 0) && r.oa > 0) uc = Math.round(r.oa / ((r.on || 0) + (r.ot || 0)));
      // ② 직전월 역추적
      if (uc <= 0) {
        const idx = DATA_MONTHS.indexOf(costRefMk);
        for (let j = idx - 1; j >= 0; j--) {
          const pm = DATA_MONTHS[j], rr = io[pm];
          if (!rr) continue;
          if (rr.eq > 0 && rr.ea > 0) { uc = Math.round(rr.ea / rr.eq); break; }
          if (rr.iq > 0 && rr.ia > 0) { uc = Math.round(rr.ia / rr.iq); break; }
          if ((rr.on > 0 || rr.ot > 0) && rr.oa > 0) { uc = Math.round(rr.oa / ((rr.on || 0) + (rr.ot || 0))); break; }
        }
      }
    }
    if (uc <= 0 && inv.unit_cost > 0) uc = inv.unit_cost;

    // ── 판매단가 (avgSP): 선택월 sales_tx 기준 ──
    const tx = D.sales_tx ? D.sales_tx.filter(t => t.sku_id === sku) : [];
    let paidAmt = 0, paidQty = 0;
    const selPrefixes = selMs.map(m => m.replace('-', '/'));
    tx.forEach(t => {
      if ((t.amt || 0) <= 0) return;
      const tPrefix = (t.date || '').substring(0, 7);
      if (selPrefixes.some(p => tPrefix === p)) { paidQty += t.qty || 0; paidAmt += t.amt || 0; }
    });
    // fallback: 선택월에 없으면 전체 기간 최신 판매월
    if (paidQty === 0) {
      const rMs = [...DATA_MONTHS].reverse();
      for (const mk of rMs) {
        const mPrefix = mk.replace('-', '/');
        const mTx = tx.filter(t => (t.date || '').startsWith(mPrefix) && (t.amt || 0) > 0);
        const mQ = mTx.reduce((s, t) => s + (t.qty || 0), 0);
        const mA = mTx.reduce((s, t) => s + (t.amt || 0), 0);
        if (mQ > 0) { paidQty = mQ; paidAmt = mA; break; }
      }
    }
    const avgSP = paidQty > 0 ? Math.round(paidAmt / paidQty) : 0;

    // ── 원가율: 수불부 단가 / 판매단가 (대시보드와 동일) ──
    const cr = avgSP > 0 && uc > 0 ? +(uc / avgSP * 100).toFixed(1) : 0;
    const margin = cr > 0 ? +(100 - cr).toFixed(1) : 0;

    // ── BOM 자재비 (참고용, 대시보드 bomCR과 동일) ──
    const children = D.bom.filter(b => b.parent_sku === sku);
    let bomCost = 0;
    children.forEach(ch => {
      let p = 0;
      const cio = mio[ch.child_sku];
      if (cio && cio[costRefMk]) {
        const r = cio[costRefMk];
        if (r.eq > 0 && r.ea > 0) p = Math.round(r.ea / r.eq);
        else if (r.iq > 0 && r.ia > 0) p = Math.round(r.ia / r.iq);
      }
      if (p <= 0 && cio) {
        const idx = DATA_MONTHS.indexOf(costRefMk);
        for (let j = idx - 1; j >= 0; j--) {
          const pm = DATA_MONTHS[j], rr = cio[pm];
          if (!rr) continue;
          if (rr.eq > 0 && rr.ea > 0) { p = Math.round(rr.ea / rr.eq); break; }
          if (rr.iq > 0 && rr.ia > 0) { p = Math.round(rr.ia / rr.iq); break; }
        }
      }
      if (p <= 0) {
        const cmP = costRefMk.replace('-', '/');
        const pur = allPur.find(x => x.sku_id === ch.child_sku && x.price > 0 && (x.date || '') <= cmP + '/31');
        if (pur) p = pur.price;
      }
      if (p <= 0) {
        const pur = allPur.find(x => x.sku_id === ch.child_sku && x.price > 0);
        if (pur) p = pur.price;
        else { const ci = D.inv.find(i => i.sku_id === ch.child_sku); if (ci && ci.qty > 0 && ci.amount > 0) p = Math.round(ci.amount / ci.qty); }
      }
      bomCost += p * (ch.qty || 1);
    });
    const bomCR = avgSP > 0 && bomCost > 0 ? +(bomCost / avgSP * 100).toFixed(1) : 0;

    // 선택월 매출 (금액 + 수량)
    const isdItem = D.isd.find(x => x.sku_id === sku);
    const periodSales = selMs.reduce((s, m) => { const d = isdItem ? isdItem.monthly : null; return s + (d && d[m] ? d[m].amt : 0); }, 0);
    const periodQty = selMs.reduce((s, m) => { const d = isdItem ? isdItem.monthly : null; return s + (d && d[m] ? d[m].qty : 0); }, 0);
    if (avgSP > 0 || uc > 0) costItems.push({ sku, name: _nm(sku), uc, avgSP, cr, margin, bomCost, bomCR, children: children.length, periodSales, periodQty });
  });
  costItems.sort((a, b) => b.cr - a.cr);
  const withCR = costItems.filter(c => c.cr > 0);
  const avgCR = withCR.length > 0 ? +(withCR.reduce((s, c) => s + c.cr, 0) / withCR.length).toFixed(1) : 0;
  const highCR = costItems.filter(c => c.cr > 70);
  const totalMargin = withCR.reduce((s, c) => s + (c.avgSP - c.uc), 0);

  // ── 발주/공급사 ──
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

  let prevPOAmt = 0;
  if (prevMs.length > 0) {
    const prevSet = new Set(prevMs.map(m => m.replace('-', '/')));
    allPO.filter(p => { const pm = p.date_no.substring(0, 7).replace('-', '/'); return prevSet.has(pm); }).forEach(p => prevPOAmt += p.amount || 0);
  }

  // ★ 품목별 구매 TOP (공급사관리 슬라이드에 통합)
  const poByItem = {};
  periodPO.forEach(p => {
    const nm = p.item_name || p.sku_id || '-';
    if (!poByItem[nm]) poByItem[nm] = { cnt: 0, amt: 0, qty: 0 };
    poByItem[nm].cnt++; poByItem[nm].amt += p.amount || 0; poByItem[nm].qty += p.qty || 0;
  });
  const topPOItems = Object.entries(poByItem).sort((a, b) => b[1].amt - a[1].amt).slice(0, 10);

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

  // ── 거래처 ──
  const csd = D.csd || [];
  const custTotal = csd.reduce((s, c) => selMs.reduce((ss, m) => ss + (c.monthly[m] ? c.monthly[m].amt : 0), s), 0);
  const topCusts = csd.map(c => ({
    name: c.name, amt: selMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0),
    prevAmt: prevMs.reduce((s, m) => s + (c.monthly[m] ? c.monthly[m].amt : 0), 0)
  })).filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);
  const top5CustAmt = topCusts.slice(0, 5).reduce((s, c) => s + c.amt, 0);
  const custConcentration = custTotal > 0 ? Math.round(top5CustAmt / custTotal * 100) : 0;

  // ── 타계정 (★ 당월분만) ──
  const taAll = D.ta || [];
  const taYr = taAll.filter(t => (t.d || '').startsWith(yr));
  const taYrTotal = taYr.reduce((s, t) => s + (t.a || 0), 0);
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
  // SLIDE 1: 표지
  // ══════════════════════════════════════════════════════
  setStatus('표지...', 12);
  const s1 = pres.addSlide();
  s1.background = { color: C.navy };
  s1.addText('BeaverWorks', { x: 0.8, y: 1.2, w: 8.5, h: 0.9, fontSize: 48, fontFace: 'Arial Black', color: C.white, bold: true });
  s1.addText('SCM 주간 경영보고서', { x: 0.8, y: 2.1, w: 8.5, h: 0.6, fontSize: 28, color: C.ice });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.85, w: 3.0, h: 0.04, fill: { color: C.accent } });
  s1.addText(`보고 기간: ${periodStr}${prevPeriodStr ? ' (비교: ' + prevPeriodStr + ')' : ''}`, { x: 0.8, y: 3.1, w: 8, h: 0.4, fontSize: 14, color: C.gray });
  s1.addText(`생성일: ${reportDate}`, { x: 0.8, y: 3.5, w: 5, h: 0.4, fontSize: 12, color: C.gray });
  s1.addText('비버웍스 공급망관리 시스템 | 자동생성 보고서', { x: 0.8, y: 4.6, w: 6, h: 0.4, fontSize: 11, color: C.gray });

  // ══════════════════════════════════════════════════════
  // SLIDE 2: Executive Summary
  // ══════════════════════════════════════════════════════
  setStatus('Executive Summary...', 15);
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addText('Executive Summary', { x: L.mx, y: L.titleY, w: 9, h: L.titleH, fontSize: 24, fontFace: 'Arial Black', color: C.dark, margin: 0 });
  s2.addText(`${periodStr} 핵심지표${prevPeriodStr ? ' | vs ' + prevPeriodStr : ''}`, { x: L.mx, y: L.subY, w: L.tw, h: L.subH, fontSize: 11, color: C.sub });

  const kpis = [
    { label: '총 매출', val: _m(salesTotal), chg: _chg(salesTotal, prevSalesTotal), color: C.accent },
    { label: '재고자산', val: _m(endAmt), chg: _chg(endAmt, prevEndAmt), color: C.teal },
    { label: '재고건강도', val: healthScore + '점', chg: '', color: healthScore >= 70 ? C.green : C.orange },
    { label: '발주규모', val: _m(poAmt), chg: _chg(poAmt, prevPOAmt), color: C.purple },
    { label: '평균원가율', val: avgCR + '%', chg: costItems.length + '개 판매품목', color: avgCR > 60 ? C.red : C.green },
    { label: '타계정', val: _m(taTotal), chg: '', color: C.orange },
    { label: '고위험원가', val: highCR.length + '건', chg: '원가율 70%↑', color: highCR.length > 3 ? C.red : C.teal },
    { label: '거래처집중도', val: custConcentration + '%', chg: 'TOP5 비중', color: custConcentration > 80 ? C.orange : C.teal },
  ];
  kpis.forEach((kpi, idx) => {
    const col = idx % 4; const row = Math.floor(idx / 4);
    const x = L.mx + col * 2.35; const y = L.bodyY + row * 1.55;
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.15, h: 1.35, fill: { color: C.ltGray }, shadow: { type: 'outer', color: '000000', blur: 3, offset: 1, angle: 135, opacity: 0.06 } });
    s2.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.05, h: 1.35, fill: { color: kpi.color } });
    s2.addText(kpi.label, { x: x + 0.15, y: y + 0.08, w: 1.85, h: 0.25, fontSize: 10, color: C.sub, margin: 0 });
    s2.addText(kpi.val, { x: x + 0.15, y: y + 0.35, w: 1.85, h: 0.5, fontSize: 22, bold: true, color: C.dark, margin: 0 });
    if (kpi.chg) {
      const chgColor = kpi.chg.includes('+') ? C.green : kpi.chg.includes('-') ? C.red : C.sub;
      s2.addText(kpi.chg, { x: x + 0.15, y: y + 0.95, w: 1.85, h: 0.25, fontSize: 9, color: chgColor, margin: 0 });
    }
  });
  s2.addText(`${periodStr} 기준 | 전월(${prevPeriodStr || '-'}) 대비 변동률 표기`, { x: L.mx, y: 4.85, w: L.tw, h: 0.3, fontSize: 9, color: C.gray, italic: true });

  // ══════════════════════════════════════════════════════
  // SLIDE 3: 재고 현황
  // ══════════════════════════════════════════════════════
  setStatus('재고 현황...', 20);
  const s3 = pres.addSlide();
  addTitle(s3, '📦', '재고 현황', `${periodStr} | 전월대비 자산 ${_chg(endAmt, prevEndAmt)}`);

  s3.addChart(pres.charts.BAR, [{
    name: '금액(백만)', labels: ['기초', '입고', '출고', '타계정', '기말'],
    values: [beginAmt, pInAmt, pOutAmt, Math.max(0, beginAmt + pInAmt - pOutAmt - endAmt), endAmt].map(v => Math.round(v / 1e6))
  }], {
    x: L.mx, y: L.bodyY, w: L.splitL, h: 3.0, barDir: 'col',
    showTitle: true, title: '재고 흐름 (백만원)', titleColor: C.sub, titleFontSize: 10,
    chartColors: [C.accent], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  // Health gauge
  const gaugeColor = healthScore >= 70 ? C.green : healthScore >= 40 ? C.orange : C.red;
  s3.addChart(pres.charts.DOUGHNUT, [{
    name: '건강도', labels: ['건강도', '잔여', '하단'],
    values: [healthScore, 100 - healthScore, 100]
  }], {
    x: L.splitRx, y: L.bodyY - 0.1, w: L.splitRw, h: 2.8, showTitle: false,
    chartColors: [gaugeColor, 'E2E8F0', C.white],
    showPercent: false, showValue: false, showLegend: false, dataLabelPosition: 'none',
  });
  s3.addText('재고 건강도', { x: L.splitRx + 0.7, y: 2.0, w: 2.5, h: 0.3, fontSize: 10, color: C.sub, align: 'center', margin: 0 });
  s3.addText(`${healthScore}/100`, { x: L.splitRx + 0.7, y: 2.25, w: 2.5, h: 0.5, fontSize: 26, bold: true, color: gaugeColor, align: 'center', margin: 0 });
  s3.addText(`정상 ${cntNormal}  |  부족 ${cntLow}  |  없음 ${cntNoStock}  |  과다 ${cntExcess}  |  무수요 ${cntNoDemand}`, {
    x: L.splitRx, y: 3.5, w: L.splitRw, h: 0.25, fontSize: 8, color: C.sub, align: 'center', margin: 0
  });

  addInsights(s3, [
    `${periodStr} 기말재고 ${_m(endAmt)} (${_f(endQty)}개)${prevEndAmt ? ' ← 전월 ' + _m(prevEndAmt) : ''}`,
    `입고 ${_m(pInAmt)} / 출고 ${_m(pOutAmt)} — 회전율 ${endAmt > 0 ? (pOutAmt / endAmt * 100).toFixed(0) + '%' : '-'}`,
    ...(cntLow + cntNoStock > 0 ? [`⚠️ 부족 ${cntLow}건 + 재고없음 ${cntNoStock}건 → 긴급발주 대상`] : [])
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE 4: 재고 상세
  // ══════════════════════════════════════════════════════
  setStatus('재고 위험품목...', 25);
  const s4 = pres.addSlide();
  addTitle(s4, '📦', '재고 상세 — 위험품목 & 카테고리', '');

  s4.addText(`⚠️ 재고부족 — ${cntLow}건`, { x: L.mx, y: 0.8, w: 5, h: 0.25, fontSize: 10, bold: true, color: C.red, margin: 0 });
  const riskRows = lowItems.slice(0, 7).map(i => [
    _nm(i.sku_id),
    { text: _f(getEq(i.sku_id)) + '개', options: { align: 'right' } },
    { text: _f(_recentOut(i.sku_id, 3)) + '개', options: { align: 'right', color: C.orange } }
  ]);
  if (riskRows.length > 0) {
    s4.addTable([
      [{ text: '품목명', options: hdrStyle(C.red) },
       { text: '현재고', options: hdrR(C.red) },
       { text: '3개월 출고', options: hdrR(C.red) }],
      ...riskRows
    ], { x: L.mx, y: 1.05, w: L.splitL, fontSize: 9.5, border: tblBorder, colW: [2.65, 1.1, 1.4], autoPage: false });
  }

  s4.addText('카테고리별 재고', { x: L.splitRx, y: 0.8, w: L.splitRw, h: 0.25, fontSize: 10, bold: true, color: C.sub, margin: 0 });
  const catRows = invCatList.slice(0, 7).map(c => [
    c.nm, { text: _f(c.qty), options: { align: 'right' } }, { text: _m(c.amt), options: { align: 'right' } }
  ]);
  s4.addTable([
    [{ text: '카테고리', options: hdrStyle() },
     { text: '수량', options: hdrR() },
     { text: '금액', options: hdrR() }],
    ...catRows
  ], { x: L.splitRx, y: 1.05, w: L.splitRw, fontSize: 9.5, border: tblBorder, colW: [1.55, 1.0, 1.4], autoPage: false });

  if (excessItems.length > 0) {
    const exTop = excessItems.slice(0, 4).map(i => `${_nm(i.sku_id)} (현재고 ${_f(getEq(i.sku_id))}개, 출고 ${_f(_recentOut(i.sku_id, 3))}개)`);
    s4.addText([
      { text: `📈 과다재고 TOP ${Math.min(excessItems.length, 4)}건`, options: { bold: true, breakLine: true, fontSize: 10, color: C.orange } },
      ...exTop.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < exTop.length - 1, fontSize: 9 } }))
    ], { x: L.mx, y: 3.6, w: L.tw, h: 1.0, color: C.dark, valign: 'top' });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 5: 판매 분석
  // ══════════════════════════════════════════════════════
  setStatus('판매 분석...', 35);
  const s5 = pres.addSlide();
  addTitle(s5, '💰', '판매 분석', `${periodStr} ${_m(salesTotal)}${prevSalesTotal > 0 ? ' | 전월 ' + _chg(salesTotal, prevSalesTotal) : ''}`);

  const tLabels = Object.keys(salesMonthly).map(m => m.slice(5) + '월');
  const tValues = Object.values(salesMonthly).map(v => Math.round(v / 1e6 * 100) / 100);
  if (tLabels.length > 1) {
    s5.addChart(pres.charts.LINE, [{ name: '매출(백만)', labels: tLabels, values: tValues }], {
      x: L.mx, y: L.bodyY, w: L.splitL, h: 3.0,
      showTitle: true, title: '최근 6개월 매출 추이', titleColor: C.sub, titleFontSize: 10,
      lineSize: 3, lineSmooth: true, chartColors: [C.accent],
      showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  if (topSalesCat.length > 0) {
    s5.addChart(pres.charts.DOUGHNUT, [{
      name: '카테고리', labels: topSalesCat.slice(0, 6).map(c => c.nm),
      values: topSalesCat.slice(0, 6).map(c => Math.round(c.amt / 1e6))
    }], {
      x: L.splitRx, y: L.bodyY, w: L.splitRw, h: 3.0,
      showTitle: true, title: '카테고리별 매출', titleColor: C.sub, titleFontSize: 10,
      chartColors: ['2563EB', '0E7490', '15803D', '7E22CE', 'C2410C', '64748B'],
      showPercent: true, showLegend: true, legendPos: 'b', legendFontSize: 8
    });
  }

  addInsights(s5, [
    `${periodStr} 총 매출 ${_m(salesTotal)} — 완제품 ${Math.round(cSalesAmt / (salesTotal || 1) * 100)}% / 원재료 ${Math.round(bSalesAmt / (salesTotal || 1) * 100)}%`,
    ...(prevSalesTotal > 0 ? [`전월(${prevPeriodStr}) ${_m(prevSalesTotal)} → 당월 ${_m(salesTotal)} (${_chg(salesTotal, prevSalesTotal)})`] : []),
    ...(topSalesCat.length > 0 ? [`주력: ${topSalesCat[0].nm} ${_m(topSalesCat[0].amt)} (전체 ${salesTotal > 0 ? Math.round(topSalesCat[0].amt / salesTotal * 100) : 0}%)`] : [])
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE 6: 판매 TOP 품목
  // ══════════════════════════════════════════════════════
  setStatus('판매 TOP 품목...', 40);
  const s6 = pres.addSlide();
  addTitle(s6, '💰', '판매 TOP 10 품목', `${periodStr} 기준 | 전월 비교`);

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
  s6.addTable([
    [{ text: '#', options: hdrStyle() },
     { text: '품목명', options: hdrStyle() },
     { text: '수량', options: hdrR() },
     { text: '금액', options: hdrR() },
     { text: '비중', options: hdrR() },
     { text: 'vs전월', options: hdrR() }],
    ...itemRows
  ], { x: L.mx, y: L.bodyY, w: L.tblFull, fontSize: 9.5, border: tblBorder, colW: [0.4, 3.3, 1.0, 1.3, 0.9, 2.3], autoPage: false });

  // ══════════════════════════════════════════════════════
  // SLIDE 7: 원가 분석 (★ 선택월 판매 품목 기준)
  // ══════════════════════════════════════════════════════
  setStatus('원가 분석...', 50);
  const s7 = pres.addSlide();
  addTitle(s7, '💹', '원가 분석', `${periodStr} 판매 ${costItems.length}개 품목 | 평균 원가율 ${avgCR}%`);

  const crBands = [
    { label: '~50%', cnt: withCR.filter(c => c.cr <= 50).length },
    { label: '50~70%', cnt: withCR.filter(c => c.cr > 50 && c.cr <= 70).length },
    { label: '70~90%', cnt: withCR.filter(c => c.cr > 70 && c.cr < 90).length },
    { label: '90%↑', cnt: withCR.filter(c => c.cr >= 90).length },
  ];
  s7.addChart(pres.charts.BAR, [{
    name: '품목수', labels: crBands.map(b => b.label), values: crBands.map(b => b.cnt)
  }], {
    x: L.mx, y: L.bodyY, w: L.splitL, h: 3.0, barDir: 'col',
    showTitle: true, title: '원가율 구간별 분포', titleColor: C.sub, titleFontSize: 10,
    chartColors: [C.green, C.teal, C.yellow, C.red],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
  });

  s7.addTable([
    [{ text: '지표', options: hdrStyle() },
     { text: '값', options: hdrR() }],
    ['분석대상 (판매품목)', { text: costItems.length + '개', options: { align: 'right' } }],
    ['원가율 산출 기준', { text: '수불부단가/판매단가', options: { align: 'right', fontSize: 9 } }],
    ['평균 원가율', { text: avgCR + '%', options: { align: 'right' } }],
    ['고위험(>70%)', { text: highCR.length + '건', options: { align: 'right', color: C.red, bold: true } }],
    ['양호(<50%)', { text: withCR.filter(c => c.cr <= 50).length + '건', options: { align: 'right' } }],
    ['최고 원가율', { text: costItems.length > 0 ? costItems[0].name + ' ' + costItems[0].cr + '%' : '-', options: { align: 'right' } }],
  ], { x: L.splitRx, y: L.bodyY, w: L.splitRw, h: 2.5, fontSize: 10, border: tblBorder, colW: [1.8, 2.15], autoPage: false });

  addInsights(s7, [
    `${periodStr} 판매된 ${costItems.length}개 품목 대상 분석 — 평균 원가율 ${avgCR}%`,
    ...(highCR.length > 0 ? [`⚠️ 고위험 ${highCR.length}건: ${highCR.slice(0, 3).map(c => c.name).join(', ')}`] : ['수익성 양호'])
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE 8: 원가 상세 (★ 선택월 판매 품목만)
  // ══════════════════════════════════════════════════════
  setStatus('원가 상세...', 55);
  const s8 = pres.addSlide();
  addTitle(s8, '💹', `원가 상세 — ${periodStr} 판매 품목`, '원가율 높은 순');

  const crRows = costItems.slice(0, 12).map(c => {
    const diffPct = c.cr > 0 && c.bomCR > 0 ? +(c.bomCR - c.cr).toFixed(1) : 0;
    return [
      c.name,
      { text: _f(c.uc), options: { align: 'right' } },
      { text: _f(c.avgSP), options: { align: 'right' } },
      { text: c.cr + '%', options: { align: 'right', color: c.cr > 70 ? C.red : c.cr > 50 ? C.orange : C.green, bold: c.cr > 70 } },
      { text: c.bomCR + '%', options: { align: 'right', color: c.bomCR > 70 ? C.red : c.bomCR > 50 ? C.orange : C.green } },
      { text: (diffPct >= 0 ? '+' : '') + diffPct + '%p', options: { align: 'right', color: Math.abs(diffPct) > 15 ? C.orange : C.sub, fontSize: 8 } },
      { text: _f(c.periodQty), options: { align: 'right' } },
      { text: _m(c.periodSales), options: { align: 'right' } }
    ];
  });
  s8.addTable([
    [{ text: '품목명', options: hdrStyle() },
     { text: '수불부원가', options: hdrR() },
     { text: '판매단가', options: hdrR() },
     { text: '원가율', options: hdrR() },
     { text: 'BOM원가율', options: hdrR() },
     { text: '차이', options: hdrR() },
     { text: '수량', options: hdrR() },
     { text: '매출', options: hdrR() }],
    ...crRows
  ], { x: L.mx, y: L.bodyY, w: L.tblFull, fontSize: 8.5, border: tblBorder, colW: [2.2, 1.0, 1.0, 0.85, 0.95, 0.7, 0.7, 1.8], autoPage: false });

  // ══════════════════════════════════════════════════════
  // SLIDE 9: ★ 고위험 원가 상세 (>70%) — 신규
  // ══════════════════════════════════════════════════════
  if (highCR.length > 0) {
    setStatus('고위험 원가 상세...', 60);
    const s9 = pres.addSlide();
    s9.background = { color: C.white };
    s9.addText('⚠️ 고위험 원가 상세 (원가율 70% 초과)', { x: L.mx, y: L.titleY, w: 9, h: L.titleH, fontSize: 22, fontFace: 'Arial Black', color: C.red, margin: 0 });
    s9.addText(`${periodStr} 판매 품목 중 ${highCR.length}건`, { x: L.mx, y: L.subY, w: L.tw, h: L.subH, fontSize: 11, color: C.sub });

    const hrRows = highCR.slice(0, 10).map((c, i) => {
      return [
        String(i + 1),
        c.name,
        { text: c.cr + '%', options: { align: 'right', color: C.red, bold: true } },
        { text: _f(c.uc), options: { align: 'right' } },
        { text: _f(c.avgSP), options: { align: 'right' } },
        { text: c.margin + '%', options: { align: 'right', color: parseFloat(c.margin) < 30 ? C.red : C.dark } },
        { text: c.bomCR + '%', options: { align: 'right', color: c.bomCR > 70 ? C.orange : C.sub } }
      ];
    });

    s9.addTable([
      [{ text: '#', options: hdrStyle(C.red) },
       { text: '품목명', options: hdrStyle(C.red) },
       { text: '원가율', options: hdrR(C.red) },
       { text: '수불부원가', options: hdrR(C.red) },
       { text: '판매단가', options: hdrR(C.red) },
       { text: '마진율', options: hdrR(C.red) },
       { text: 'BOM원가율', options: hdrR(C.red) }],
      ...hrRows
    ], { x: L.mx, y: L.bodyY, w: L.tblFull, fontSize: 9, border: tblBorder, colW: [0.35, 2.65, 0.8, 1.2, 1.2, 0.9, 2.1], autoPage: false });

    const hrTblEnd = L.bodyY + (hrRows.length + 1) * 0.3 + 0.15;
    addInsights(s9, [
      `고위험 ${highCR.length}건 — 원가율 70% 초과 품목의 마진 개선 시급`,
      ...(highCR.filter(c => c.margin < 0).length > 0 ? [`🚨 마진 적자 ${highCR.filter(c => c.margin < 0).length}건 — 가격 인상 또는 원가 절감 필요`] : []),
      `최대 원가 자재를 활용하여 BOM 원가 최적화 검토`
    ]);
  }

  // ══════════════════════════════════════════════════════
  // SLIDE: BOM 관리 & 생산가능성
  // ══════════════════════════════════════════════════════
  setStatus('BOM 분석...', 68);
  const sBom = pres.addSlide();
  addTitle(sBom, '🔧', 'BOM 관리 & 생산가능성', '');

  sBom.addTable([
    [{ text: '지표', options: hdrStyle() },
     { text: '값', options: hdrR() }],
    ['BOM 총 관계', { text: _f(bomCnt) + '건', options: { align: 'right' } }],
    ['완제품 (Parent)', { text: _f(bomParentCnt) + '종', options: { align: 'right' } }],
    ['원재료 (Child)', { text: _f(bomMatCnt) + '종', options: { align: 'right' } }],
    ['즉시 생산가능', { text: _f(canProduce) + '종', options: { align: 'right', color: C.green } }],
    ['자재부족 (생산불가)', { text: _f(cantProduce) + '종', options: { align: 'right', color: C.red } }],
    ['생산가능률', { text: (bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) : 0) + '%', options: { align: 'right' } }],
  ], { x: L.mx, y: L.bodyY, w: L.splitL, h: 2.6, fontSize: 10, border: tblBorder, colW: [2.7, 2.45], autoPage: false });

  sBom.addChart(pres.charts.PIE, [{
    name: '생산가능성', labels: ['생산가능', '자재부족'], values: [canProduce, cantProduce]
  }], {
    x: L.splitRx, y: L.bodyY, w: L.splitRw, h: 2.5,
    showTitle: true, title: '생산가능 현황', titleColor: C.sub, titleFontSize: 10,
    chartColors: [C.green, C.red], showPercent: true, showLegend: true, legendPos: 'b'
  });

  if (cantProduceList.length > 0) {
    const cpTxt = cantProduceList.slice(0, 3).map(cp => `${cp.parentName}: ${cp.shortMats.slice(0, 2).map(m => m.name).join(', ')}`);
    sBom.addText([
      { text: '⚠️ 자재부족 품목', options: { bold: true, breakLine: true, fontSize: 10, color: C.red } },
      ...cpTxt.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < cpTxt.length - 1, fontSize: 9 } }))
    ], { x: L.splitRx, y: 3.5, w: L.splitRw, h: 1.0, color: C.dark, valign: 'top' });
  }

  addInsights(sBom, [
    `${bomParentCnt}개 완제품 중 ${canProduce}개 즉시 생산가능 (${bomParentCnt > 0 ? Math.round(canProduce / bomParentCnt * 100) : 0}%)`,
    ...(cantProduce > 0 ? [`⚠️ ${cantProduce}개 자재부족 — 해당 원재료 긴급 발주 필요`] : [])
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE: 거래처 분석 (판매)
  // ══════════════════════════════════════════════════════
  setStatus('거래처 분석...', 75);
  const sCust = pres.addSlide();
  addTitle(sCust, '🏢', '거래처 분석 (판매)', `TOP5 집중도 ${custConcentration}% | ${periodStr}`);

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
  sCust.addTable([
    [{ text: '#', options: hdrStyle() },
     { text: '거래처명', options: hdrStyle() },
     { text: '당월 매출', options: hdrR() },
     { text: '비중', options: hdrR() },
     { text: 'vs전월', options: hdrR() }],
    ...custRows
  ], { x: L.mx, y: L.bodyY, w: L.tblFull, fontSize: 9.5, border: tblBorder, colW: [0.4, 3.3, 1.8, 1.2, 2.5], autoPage: false });

  addInsights(sCust, [
    `${periodStr} 활동 거래처 ${topCusts.length}개, 총 ${_m(custTotal)}`,
    `TOP5 집중도 ${custConcentration}% — ${custConcentration > 80 ? '⚠️ 의존도 높음' : '적정 분산'}`
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE: 타계정 (★ 당월분만)
  // ══════════════════════════════════════════════════════
  setStatus('타계정...', 82);
  const sTa = pres.addSlide();
  addTitle(sTa, '📑', '타계정 내역', `${periodStr} ${_m(taTotal)} / ${_f(taQty)}건 (${yr}년 누계 ${_m(taYrTotal)})`);

  const taLabels = Object.keys(taMonthly).map(m => m.slice(5) + '월');
  const taValues = Object.values(taMonthly).map(v => Math.round(v / 1e6 * 100) / 100);
  if (taLabels.length > 1) {
    sTa.addChart(pres.charts.BAR, [{
      name: '타계정(백만)', labels: taLabels, values: taValues
    }], {
      x: L.mx, y: L.bodyY, w: L.splitL, h: 3.0, barDir: 'col',
      showTitle: true, title: '월별 타계정 추이', titleColor: C.sub, titleFontSize: 10,
      chartColors: [C.orange], showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.dark, dataLabelFontSize: 9,
      valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: C.gray, valAxisLabelColor: C.gray
    });
  }

  sTa.addText(`계정별 집계 (${periodStr})`, { x: L.splitRx, y: L.bodyY, w: L.splitRw, h: 0.25, fontSize: 10, bold: true, color: C.sub, margin: 0 });
  if (taAcctList.length > 0) {
    sTa.addTable([
      [{ text: '계정', options: hdrStyle() },
       { text: '건수', options: hdrR() },
       { text: '금액', options: hdrR() }],
      ...taAcctList.slice(0, 7).map(([ac, v]) => [
        ac, { text: _f(v.qty), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
      ])
    ], { x: L.splitRx, y: L.bodyY + 0.3, w: L.splitRw, fontSize: 9.5, border: tblBorder, colW: [1.5, 1.0, 1.45], autoPage: false });
  }

  addInsights(sTa, [
    `${periodStr} 타계정 ${_m(taTotal)} (${_f(taQty)}건) — ${yr}년 누계 ${_m(taYrTotal)}`,
    ...(taAcctList.length > 0 ? [`당월 최대: ${taAcctList[0][0]} — ${_m(taAcctList[0][1].amt)} (${taTotal > 0 ? Math.round(taAcctList[0][1].amt / taTotal * 100) : 0}%)`] : []),
    ...(taYrTotal > 50000000 ? [`⚠️ 연간 누계 ${_m(taYrTotal)} — 정리/회수 계획 수립 필요`] : [])
  ]);

  // ══════════════════════════════════════════════════════
  // SLIDE: ★ 공급사 관리 (구매) + 구매 품목 TOP 통합
  // (★ 발주관리 슬라이드 제거, 구매 품목 상위를 여기에 통합)
  // ══════════════════════════════════════════════════════
  setStatus('공급사 관리...', 90);
  const sVen = pres.addSlide();
  addTitle(sVen, '🏭', '공급사 관리 (구매)', `${periodStr} | ${poVendors.length}개사 ${_m(poAmt)}${prevPOAmt > 0 ? ' (전월 ' + _chg(poAmt, prevPOAmt) + ')' : ''}`);

  // 공급사 테이블 (좌)
  sVen.addText('공급사별 발주', { x: L.mx, y: 0.8, w: 5, h: 0.25, fontSize: 10, bold: true, color: C.sub, margin: 0 });
  const vRows = topPOVendors.slice(0, 7).map(([nm, v]) => [
    nm, { text: _f(v.cnt), options: { align: 'right' } },
    { text: _m(v.amt), options: { align: 'right' } },
    { text: poAmt > 0 ? Math.round(v.amt / poAmt * 100) + '%' : '-', options: { align: 'right' } }
  ]);
  sVen.addTable([
    [{ text: '공급사', options: hdrStyle() },
     { text: '건수', options: hdrR() },
     { text: '금액', options: hdrR() },
     { text: '비중', options: hdrR() }],
    ...vRows
  ], { x: L.mx, y: 1.05, w: L.splitL, fontSize: 9.5, border: tblBorder, colW: [1.75, 0.8, 1.3, 1.3], autoPage: false });

  // ★ 구매 품목 TOP (우 — 발주관리에서 이동)
  sVen.addText('구매 품목 상위', { x: L.splitRx, y: 0.8, w: L.splitRw, h: 0.25, fontSize: 10, bold: true, color: C.sub, margin: 0 });
  const poItemRows = topPOItems.slice(0, 8).map(([nm, v]) => [
    nm, { text: _f(v.qty), options: { align: 'right' } }, { text: _m(v.amt), options: { align: 'right' } }
  ]);
  sVen.addTable([
    [{ text: '품목명', options: hdrStyle(C.purple) },
     { text: '수량', options: hdrR(C.purple) },
     { text: '금액', options: hdrR(C.purple) }],
    ...poItemRows
  ], { x: L.splitRx, y: 1.05, w: L.splitRw, fontSize: 9, border: tblBorder, colW: [1.75, 0.8, 1.4], autoPage: false });

  addInsights(sVen, [
    `${periodStr} ${poVendors.length}개 공급사 / ${_f(periodPO.length)}건 / ${_m(poAmt)}`,
    ...(topPOVendors.length > 0 && poAmt > 0 && Math.round(topPOVendors[0][1].amt / poAmt * 100) > 50 ? [`⚠️ ${topPOVendors[0][0]} 비중 ${Math.round(topPOVendors[0][1].amt / poAmt * 100)}% — 리스크 분산 필요`] : []),
    ...(topPOItems.length > 0 ? [`구매 최대 품목: ${topPOItems[0][0]} (${_m(topPOItems[0][1].amt)})`] : [])
  ]);

  // ★ 발주관리 슬라이드 — 제거됨
  // ★ 액션아이템즈 슬라이드 — 제거됨

  // ══════════════════════════════════════════════════════
  // SAVE
  // ══════════════════════════════════════════════════════
  setStatus('저장 중...', 98);
  const fileName = `BW_SCM_Weekly_${yr}${S.selMonth === 'all' ? '' : '_' + S.selMonth}_${reportDate.replace(/-/g, '')}.pptx`;
  await pres.writeFile({ fileName });
  setStatus('완료!', 100);
}
