/**
 * SCM Dashboard — Google Sheets CSV 설정
 * ========================================
 *
 * Excel 데이터를 Google Sheets로 변환 → 공유(뷰어) 설정 후
 * gviz CSV endpoint로 대시보드에서 fetch합니다.
 *
 * URL 형식: https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv
 *
 * 월별 업데이트 방법:
 *   1. extract_to_csv.py 실행하여 새 CSV 생성
 *   2. Google Sheets에 새 데이터 복사/붙여넣기 또는 CSV 임포트
 *   3. 대시보드 새로고침 (gviz endpoint가 자동으로 최신 데이터 제공)
 */

const SHEETS_CONFIG = {
  // ===== 파일별 Google Sheets ID =====
  // Google Sheets "웹에 게시" CSV endpoint 사용 (CORS 지원)
  SHEETS: {
    subul_b:    { id: '1F7IbHocw-0LSNZ-NYQZs76cQ-t7M_sVL_LOsb7jETZo', name: '원재료B' },
    subul_c:    { id: '1rhO4EEdnUhlx21qqx5oJXE8lp94lhOCww-Ho0VwKH8E', name: '제품C' },
    ta_detail:  { id: '1htpbNVn5HRfoXfeXYEDyD8HsnDYJyxxsnaeNi4nR0lI', name: '타계정내역' },
    sales:      { id: '1RvRiCMZUFw4VCIwWjhTq3v27_zySYM6iokF3Vw5tnWA', name: '판매현황' },
    purchase:   { id: '1av1bb3PYH3vgJ9eSEH4sFumn6GD1YmlDw_V18wxukEM', name: '구매현황' },
    orders:     { id: '1f043aCULe5AoBvlYCsvoxclTwTDFCuboq1hUxC-i9J8', name: '발주서' },
    production: { id: '15UmZo39D-oxCXGwyRnTvX5xhWBQX9NdQrwPOAVOdmww', name: '생산입고소모' },
    sku_master: { id: '1TfDoVA2U9fMYf4DBrLYdE-AGWyRavGrdwf2UHLr5j44', name: '품목코드' },
    pl_data:    { id: '1-xAyzdvTvcIslDZTtIa_ZkLfIZHSodS5KlMmnYkoftI', name: 'PL_DATA' },
  },

  // ===== Google Sheets gviz CSV endpoint =====
  csvUrl(sheetKey) {
    const cfg = this.SHEETS[sheetKey];
    if (!cfg || !cfg.id) return null;
    // Google Sheets gviz CSV endpoint (CORS 지원, 공유 설정 필요: 링크가 있는 모든 사용자 → 뷰어)
    return `https://docs.google.com/spreadsheets/d/${cfg.id}/gviz/tq?tqx=out:csv`;
  },

  // ===== 수불부 원천 Excel 컬럼 오프셋 =====
  // CSV에서 A열="월(YYYY-MM)" 추가 → 나머지 컬럼이 1칸씩 밀림
  SUBUL_MONTH_COL: 0,   // A열: 월 키 (YYYY-MM)
  SUBUL_DATA_OFFSET: 1, // B열부터 기존 Excel 컬럼 시작
};
