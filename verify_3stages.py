#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
3단계 자체 검증(Self-Verification) 자동화 스크립트:
- [검증 1] 데이터 수집·시계열 DB(`handon_master.db`) 정합성, 엑셀 `=HYPERLINK` 수식 & 슬랙 미발송(0건) 검증
- [검증 2] FastAPI 웹 서버(`0.0.0.0:8000`) 엔드포인트 및 체크박스(`POST/PATCH /api/tracked`) DB 영구 반영 검증
- [검증 3] PC(1920x1080) 및 모바일(iPhone 14 Pro 393x852 / Galaxy S23 360x800) 반응형 레이아웃·차트·가로스크롤(0px) 검증
"""

import json
import sqlite3
import sys
from pathlib import Path

import openpyxl
import requests
from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORK_DIR = Path(__file__).resolve().parent
DB_PATH = WORK_DIR / "handon_master.db"
EXCEL_PATH = WORK_DIR / "한돈_누적관리대장.xlsx"
BASE_URL = "http://localhost:8000"


def run_stage1_verification() -> dict:
    print("=== [검증 1] 데이터 수집·시계열 DB 정합성 & 슬랙 미발송 검증 ===")
    assert DB_PATH.exists(), f"DB 파일 미존재: {DB_PATH}"
    assert EXCEL_PATH.exists(), f"엑셀 파일 미존재: {EXCEL_PATH}"

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row

    runs = [dict(r) for r in conn.execute("SELECT * FROM crawl_runs ORDER BY run_id ASC").fetchall()]
    assert len(runs) >= 2, f"시계열 누적 스냅샷 회차가 2회 미만입니다: {len(runs)}"

    total_snaps = conn.execute("SELECT COUNT(*) FROM product_snapshots").fetchone()[0]
    latest_run_id = runs[-1]["run_id"]
    latest_snaps = [
        dict(r) for r in conn.execute("SELECT * FROM product_snapshots WHERE run_id = ?", (latest_run_id,)).fetchall()
    ]

    # 흑돼지/수입육 잔존 0건 검증
    black_cnt = conn.execute(
        """
        SELECT COUNT(*) FROM product_snapshots
        WHERE raw_name LIKE '%흑돼지%' OR raw_name LIKE '%흑돈%'
           OR raw_name LIKE '%버크셔%' OR raw_name LIKE '%우리흑돈%'
           OR raw_name LIKE '%지리산흑돈%' OR raw_name LIKE '%제주흑돈%'
        """
    ).fetchone()[0]
    assert black_cnt == 0, f"흑돼지 상품이 잔존합니다: {black_cnt}건"

    # 슬랙 발송 0건 검증
    slack_calls = sum(int(r["slack_sent_cnt"]) for r in runs)
    assert slack_calls == 0, f"슬랙 발송 횟수가 0이 아닙니다: {slack_calls}"

    # 구형 '특수부위/부산물', '앞다리살', '뒷다리살' 잔존 0건 및 27개 정밀 부위 분리 검증
    legacy_part_cnt = conn.execute(
        "SELECT COUNT(*) FROM product_snapshots WHERE standard_part IN ('특수부위/부산물', '앞다리살', '뒷다리살')"
    ).fetchone()[0]
    assert legacy_part_cnt == 0, f"구형 부위명이 잔존합니다: {legacy_part_cnt}건"

    distinct_parts = {r["standard_part"] for r in latest_snaps}
    required_separated_parts = {
        "삼겹살",
        "미박삼겹살",
        "목살",
        "앞다리살(전지)",
        "미박앞다리살(미박전지)",
        "뒷다리살(후지)",
        "미박뒷다리살(미박후지)",
        "사태",
        "미박사태",
        "사골",
        "등뼈",
        "장족",
        "장족(뒤)",
        "단족/미니족",
        "오돌뼈",
        "지방돈피/돼지껍데기",
    }
    missing_parts = required_separated_parts - distinct_parts
    assert not missing_parts, f"분리 매핑 누락 부위 발생: {missing_parts}"

    # 미박앞다리살=미박전지 & 미박뒷다리살=미박후지 4대 플랫폼 전수 매핑 교차 검증
    mibak_fore_sites = {r["source_site"] for r in latest_snaps if r["standard_part"] == "미박앞다리살(미박전지)"}
    mibak_hind_sites = {r["source_site"] for r in latest_snaps if r["standard_part"] == "미박뒷다리살(미박후지)"}
    assert mibak_fore_sites == {"미트박스", "금천미트", "미트프렌즈", "탑미트"}, f"미박앞다리살(미박전지) 플랫폼 누락: {mibak_fore_sites}"
    assert mibak_hind_sites == {"미트박스", "금천미트", "미트프렌즈", "탑미트"}, f"미박뒷다리살(미박후지) 플랫폼 누락: {mibak_hind_sites}"

    # 암퇘지 상품 4대 플랫폼 전수 분리 검증
    female_sites = {r["source_site"] for r in latest_snaps if r["sub_spec"] == "암퇘지"}
    assert female_sites == {"미트박스", "금천미트", "미트프렌즈", "탑미트"}, f"암퇘지 분류 플랫폼 누락: {female_sites}"

    # 필수 파싱 컬럼 결측 검증 및 금천미트 상품상세 하이퍼링크(/pd/productDetail) 검증
    missing_url = sum(1 for r in latest_snaps if not r["product_url"].startswith("http"))
    missing_brand = sum(1 for r in latest_snaps if not r["brand"])
    invalid_price = sum(1 for r in latest_snaps if int(r["unit_price_kg"]) <= 0)
    ekcm_snaps = [r for r in latest_snaps if r["source_site"] == "금천미트"]
    ekcm_legacy_submain = sum(1 for r in ekcm_snaps if "/dp/subMain" in r["product_url"])
    ekcm_detail_cnt = sum(1 for r in ekcm_snaps if "/pd/productDetail?goodsNo=" in r["product_url"])
    assert missing_url == 0, f"상품 URL 누락: {missing_url}건"
    assert missing_brand == 0, f"브랜드 누락: {missing_brand}건"
    assert invalid_price == 0, f"1kg 단가 오류: {invalid_price}건"
    assert ekcm_legacy_submain == 0, f"금천미트 구형 /dp/subMain URL 잔존: {ekcm_legacy_submain}건"
    assert ekcm_detail_cnt >= 100, f"금천미트 /pd/productDetail URL 적용 부족: {ekcm_detail_cnt}/{len(ekcm_snaps)}"

    # 한돈 도매 지육시세(2024.01.01~현재 누적) DB 검증
    carcass_stats = conn.execute(
        "SELECT COUNT(*) AS cnt, MIN(price_date) AS min_d, MAX(price_date) AS max_d FROM carcass_market_prices WHERE price_ex_jeju > 0 OR price_national > 0"
    ).fetchone()
    assert carcass_stats["cnt"] >= 700, f"한돈 지육시세 누적 거래일 부족: {dict(carcass_stats)}"
    assert carcass_stats["min_d"] <= "2024-01-05", f"한돈 지육시세 2024년 1월 누적 시작일 누락: {dict(carcass_stats)}"

    conn.close()

    # 엑셀 파일 시트 및 최저·최고가 =HYPERLINK 수식 검증 (Sheet1, Sheet2, Sheet4, Sheet5)
    wb = openpyxl.load_workbook(str(EXCEL_PATH), data_only=False)
    sheet_names = wb.sheetnames
    assert len(sheet_names) == 5, f"엑셀 시트 개수 오류: {sheet_names}"
    ws1 = wb["Sheet1_5자_최저가및재고매핑"]
    ws2 = wb["Sheet2_생산일자별_최저최고가분석"]
    ws4 = wb["Sheet4_전체상품_원본및하이퍼링크"]
    ws5 = wb["Sheet5_한돈지육시세_2024누적"]
    assert ws5.max_row >= 700, f"Sheet5_한돈지육시세_2024누적 행 수 부족: {ws5.max_row}"

    s1_min_links = sum(1 for r in range(2, ws1.max_row + 1) if str(ws1.cell(r, 29).value or "").startswith("=HYPERLINK("))
    s1_max_links = sum(1 for r in range(2, ws1.max_row + 1) if str(ws1.cell(r, 30).value or "").startswith("=HYPERLINK("))
    s2_min_links = sum(1 for r in range(2, ws2.max_row + 1) if str(ws2.cell(r, 13).value or "").startswith("=HYPERLINK("))
    s2_max_links = sum(1 for r in range(2, ws2.max_row + 1) if str(ws2.cell(r, 14).value or "").startswith("=HYPERLINK("))
    assert s1_min_links > 20 and s1_max_links > 20, f"Sheet1 최저/최고 하이퍼링크 부족: {s1_min_links}, {s1_max_links}"
    assert s2_min_links > 20 and s2_max_links > 20, f"Sheet2 최저/최고 하이퍼링크 부족: {s2_min_links}, {s2_max_links}"

    hyperlink_cnt = 0
    for row_idx in range(2, min(ws4.max_row + 1, 100)):
        val = str(ws4.cell(row=row_idx, column=24).value or "")
        if val.startswith("=HYPERLINK("):
            hyperlink_cnt += 1
    assert hyperlink_cnt > 50, f"엑셀 =HYPERLINK 수식 미적용: {hyperlink_cnt}"

    res = {
        "runs_count": len(runs),
        "total_snapshots_in_db": total_snaps,
        "latest_run_id": latest_run_id,
        "latest_valid_sku_cnt": len(latest_snaps),
        "black_pork_remaining": black_cnt,
        "slack_sent_cnt": slack_calls,
        "excel_sheets": sheet_names,
        "excel_s1_min_max_links": [s1_min_links, s1_max_links],
        "excel_s2_min_max_links": [s2_min_links, s2_max_links],
        "excel_hyperlink_verified_sample": hyperlink_cnt,
    }
    print("STAGE 1 OK:", json.dumps(res, ensure_ascii=False))
    return res


def run_stage2_verification() -> dict:
    print("=== [검증 2] 웹 서버 구동 & 체크박스 트래킹 API 검증 ===")
    r_index = requests.get(f"{BASE_URL}/", timeout=10)
    assert r_index.status_code == 200, f"GET / 실패: {r_index.status_code}"

    r_health = requests.get(f"{BASE_URL}/api/health", timeout=10)
    assert r_health.status_code == 200, f"GET /api/health 실패: {r_health.status_code}"
    health_data = r_health.json()
    sched_info = health_data.get("scheduler", {})
    assert sched_info.get("enabled") is True, f"자동 스케줄러 미활성화: {sched_info}"
    assert sched_info.get("weekdays_only") is True, f"평일(월~금) 전용 설정 불일치: {sched_info}"
    assert sched_info.get("schedule_days") == "월~금 (토·일 제외)", f"요일 설정 불일치: {sched_info}"
    assert sched_info.get("schedule_times") == ["09:00", "16:00"], f"스케줄 시간 설정 불일치: {sched_info}"
    assert sched_info.get("slack_enabled") is False, f"슬랙 비활성화 정책 불일치: {sched_info}"

    r_dash = requests.get(f"{BASE_URL}/api/dashboard", timeout=15)
    assert r_dash.status_code == 200, f"GET /api/dashboard 실패: {r_dash.status_code}"
    dash_data = r_dash.json()
    assert dash_data["has_data"] is True
    assert "sales_velocity" not in dash_data, "제외 요청된 sales_velocity 필드가 남아 있습니다."
    assert dash_data["kpi"]["female_sku_cnt"] > 50, f"암퇘지 SKU 집계 오류: {dash_data['kpi']}"
    assert "caution_stock_sku_cnt" in dash_data["kpi"], "품절/재고주의 KPI 필드 누락"
    assert "carcass_sise" in dash_data and dash_data["carcass_sise"]["total_days"] >= 700, "API 한돈 지육시세 누적 데이터 누락"
    assert dash_data["carcass_sise"]["min_date"] <= "2024-01-05", f"API 지육시세 시작일 오류: {dash_data['carcass_sise']['min_date']}"

    # 냉장 5개 미만 / 냉동 10개 미만 기준 정확성 검증
    for p_row in dash_data["products"]:
        expected_lim = 5 if p_row["keeping"] == "냉장" else 10
        assert p_row["low_stock_limit"] == expected_lim, f"재고주의 기준값 불일치: {p_row}"
        expected_low = 1 if (not p_row["is_sold_out"] and 0 < int(p_row["stock_qty"] or 0) < expected_lim) else 0
        assert p_row["is_low_stock"] == expected_low, f"재고주의 판정 불일치: {p_row}"

    # 생산일자별 분석 내 근본한돈(메인) vs 타플랫폼 구조 검증
    pd_rows = dash_data["prod_date_analysis"]["part_date_matrix"]
    assert len(pd_rows) > 15 and "gb_main" in pd_rows[0] and "comp_buckets" in pd_rows[0], "생산일자 근본한돈 대비 구조 누락"

    # 비교 매트릭스 내 일반한돈 vs 암퇘지 완벽 분리 검증 (일반 행에 암퇘지 혼입 0건 보장)
    mat_rows = dash_data["comparison_matrix"]
    female_mat_rows = [m for m in mat_rows if m.get("sex_type") == "암퇘지"]
    normal_mat_rows = [m for m in mat_rows if m.get("sex_type") == "일반"]
    assert len(female_mat_rows) >= 10, f"암퇘지 전용 매트릭스 행 부족: {len(female_mat_rows)}"
    for nm in normal_mat_rows:
        for grp_k, g_cell in nm["groups"].items():
            b_nm = g_cell.get("best_name") or ""
            assert "암퇘지" not in b_nm and "(암)" not in b_nm, f"일반한돈 행에 암퇘지 혼입 발생: {nm['standard_part']} / {grp_k} / {b_nm}"

    # 테스트용 상품 선택 후 POST / PATCH /api/tracked 검증
    sample_item = dash_data["products"][0]
    sample_key = sample_item["item_key"]
    orig_tracked = bool(sample_item["is_tracked"])

    # 1) POST로 체크 상태 변경
    r_post = requests.post(
        f"{BASE_URL}/api/tracked",
        json={"item_key": sample_key, "is_active": True, "memo": "자동검증 테스트 체크"},
        timeout=10,
    )
    assert r_post.status_code == 200, f"POST /api/tracked 실패: {r_post.status_code}"
    assert r_post.json()["updated_item"]["is_active"] == 1

    # DB 직접 조회로 영구 반영 확인
    conn = sqlite3.connect(str(DB_PATH))
    db_row = conn.execute("SELECT is_active, memo FROM tracked_items WHERE item_key = ?", (sample_key,)).fetchone()
    conn.close()
    assert db_row is not None and db_row[0] == 1, "DB tracked_items 반영 실패"

    # 2) PATCH로 원래 상태 복원
    r_patch = requests.patch(
        f"{BASE_URL}/api/tracked",
        json={"item_key": sample_key, "is_active": orig_tracked},
        timeout=10,
    )
    assert r_patch.status_code == 200, f"PATCH /api/tracked 실패: {r_patch.status_code}"

    # 3) 엑셀 다운로드 API 검증
    r_excel = requests.get(f"{BASE_URL}/api/export/excel", timeout=15)
    assert r_excel.status_code == 200 and len(r_excel.content) > 50000, "GET /api/export/excel 실패"

    res = {
        "index_http": r_index.status_code,
        "dashboard_http": r_dash.status_code,
        "post_tracked_http": r_post.status_code,
        "patch_tracked_http": r_patch.status_code,
        "excel_export_http": r_excel.status_code,
        "excel_bytes": len(r_excel.content),
        "scheduler": sched_info,
        "local_ip": health_data.get("local_ip"),
    }
    print("STAGE 2 OK:", json.dumps(res, ensure_ascii=False))
    return res


def run_stage3_verification() -> dict:
    print("=== [검증 3] PC(1920x1080) 및 모바일(iOS 393x852 / Android 360x800) 반응형 렌더링 검증 ===")
    results = {}
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=True, channel="chrome")
        except Exception:
            try:
                browser = p.chromium.launch(headless=True, channel="msedge")
            except Exception:
                browser = p.chromium.launch(headless=True)

        # 1. PC Desktop (1920x1080)
        js_errors = []
        page_pc = browser.new_page(viewport={"width": 1920, "height": 1080})
        page_pc.on("pageerror", lambda err: js_errors.append(str(err)))
        page_pc.goto(BASE_URL, wait_until="networkidle", timeout=30000)
        page_pc.wait_for_timeout(1200)

        pc_metrics = page_pc.evaluate(
            """() => {
                const chipSec = document.getElementById('quickPartChipSection');
                const topMenuBtns = document.querySelectorAll('#topMenuTabsBar [data-topnav]');
                return {
                    scrollWidth: document.documentElement.scrollWidth,
                    clientWidth: document.documentElement.clientWidth,
                    chipSecNoOverflow: chipSec ? (chipSec.scrollWidth <= chipSec.clientWidth + 2) : false,
                    topMenuCount: topMenuBtns.length,
                    kpiSku: document.getElementById('kpiTotalSku')?.innerText,
                    kpiCaution: document.getElementById('kpiCautionStockCnt')?.innerText,
                    navCount: document.querySelectorAll('aside.lg\\\\:flex [data-nav]').length,
                    view6Exists: !!document.getElementById('view-6'),
                    quickResetExists: !!document.getElementById('btnQuickReset'),
                    quickKeepingBtns: document.querySelectorAll('#quickKeepingChipBar [data-qkeep]').length,
                    quickSexBtns: document.querySelectorAll('#quickSexChipBar [data-qsex]').length,
                    quickPartBtns: document.querySelectorAll('#quickPartChipBar .part-chip').length,
                    nextSchedule: document.getElementById('sidebarNextSchedule')?.innerText,
                    chartRendered: !!document.getElementById('mainPriceComparisonChart') && document.getElementById('mainPriceComparisonChart').width > 0
                };
            }"""
        )
        assert pc_metrics["scrollWidth"] <= pc_metrics["clientWidth"], f"PC 가로 오버플로우 발생: {pc_metrics}"
        assert pc_metrics["chipSecNoOverflow"] is True, f"상단 탭 영역 가로 스크롤 발생: {pc_metrics}"
        assert pc_metrics["topMenuCount"] == 5, f"상단 5대 메뉴 탭 개수 오류: {pc_metrics['topMenuCount']}"
        assert pc_metrics["navCount"] == 5, f"좌측 핵심 메뉴 5개 불일치: {pc_metrics['navCount']}"
        assert pc_metrics["view6Exists"] is False, "판매소진(View 6) 섹션이 제거되지 않았습니다."
        assert pc_metrics["quickResetExists"] is True, "퀵필터 초기화 버튼 미존재"
        assert pc_metrics["quickKeepingBtns"] == 3, f"퀵필터 냉장/냉동 버튼 개수 오류: {pc_metrics['quickKeepingBtns']}"
        assert pc_metrics["quickSexBtns"] == 3, f"퀵필터 일반/암퇘지 버튼 개수 오류: {pc_metrics['quickSexBtns']}"
        assert pc_metrics["quickPartBtns"] == 28, f"퀵필터 정밀 부위 버튼 개수 오류: {pc_metrics['quickPartBtns']}"
        assert pc_metrics["chartRendered"] is True, "PC 메인 차트 미렌더링"

        # [추가 검증 A] View 1 브리핑 카테고리별 냉장/냉동 필터 & 5개 비교군 우위현황 냉장/냉동 변동 검증
        v1_dynamic = page_pc.evaluate(
            """() => {
                const redAll = document.getElementById('badgeRedCnt').innerText;
                setAlertCategoryKeeping('red', '냉장');
                const redChilled = document.getElementById('badgeRedCnt').innerText;
                setAlertCategoryKeeping('red', '냉동');
                const redFrozen = document.getElementById('badgeRedCnt').innerText;
                setAlertKeepingAll('ALL');

                const sbAllHtml = document.getElementById('platformScoreboardList').innerText;
                setScoreboardKeeping('냉장');
                const sbChilledHtml = document.getElementById('platformScoreboardList').innerText;
                setScoreboardKeeping('냉동');
                const sbFrozenHtml = document.getElementById('platformScoreboardList').innerText;
                setScoreboardKeeping('ALL');
                return {
                    redAll, redChilled, redFrozen,
                    scoreboardChanged: (sbAllHtml !== sbChilledHtml) && (sbChilledHtml !== sbFrozenHtml)
                };
            }"""
        )
        assert v1_dynamic["scoreboardChanged"] is True, f"5개 비교군 우위현황 냉장/냉동 변동 미작동: {v1_dynamic}"

        # [추가 검증 B] 상단 부위 중복 탭 최대 5개 선택 & View 3 상단 4개 카드 반응 & View 4 카테고리 차트 반응 검증
        multi_and_v3_v4 = page_pc.evaluate(
            """() => {
                switchView(3);
                const v3AllCardTxt = document.getElementById('prodBucketSummaryCards').innerText;

                // 6개 연속 클릭 시도 -> 최대 5개까지만 선택되어야 함
                const testParts = ['삼겹살', '미박삼겹살', '목살', '앞다리살(전지)', '미박앞다리살(미박전지)', '뒷다리살(후지)'];
                testParts.forEach(pt => selectQuickPartChip(pt));
                const selectedLenAfter6 = STATE.selectedParts.length;
                const v3FilteredCardTxt = document.getElementById('prodBucketSummaryCards').innerText;

                switchView(4);
                const v4ChartLabels = STATE.charts.view4Category ? [...STATE.charts.view4Category.data.labels] : [];
                const v4MinMaxLinks = document.querySelectorAll('#productDesktopBody a[href^="http"]').length;

                // 전체 시트(View 2, 3, 4, 5) Sticky 행·열 고정 스타일 검증
                const stickyChecks = {};
                [2, 3, 4, 5].forEach(v => {
                    switchView(v);
                    const th = document.querySelector(`#view-${v} .sticky-table thead th:first-child`);
                    const td = document.querySelector(`#view-${v} .sticky-table tbody td:first-child`);
                    stickyChecks[`view_${v}_th_sticky`] = th ? getComputedStyle(th).position === 'sticky' : false;
                    stickyChecks[`view_${v}_td_sticky`] = td ? getComputedStyle(td).position === 'sticky' : false;
                });

                // View 2 플랫폼별 필터 & 최저·최고가 하이퍼링크 검증
                switchView(2);
                onChangeMatrixPlatformFilter('금천미트');
                const v2PlatSummary = document.getElementById('matrixPlatformSummaryText').innerText;
                const v2MinLinks = Array.from(document.querySelectorAll('#matrixDesktopBody a')).filter(a => a.innerText.includes('최저')).length;
                const v2MaxLinks = Array.from(document.querySelectorAll('#matrixDesktopBody a')).filter(a => a.innerText.includes('최고')).length;

                resetQuickAndAllFilters();
                switchView(1);

                // [추가 검증 C] 1번 메뉴 하단 한돈 도매 지육시세 (1년/월/주/일 & 필터 & 시각화 차트 & 셀차트) 검증
                setCarcassGranularity('YEAR');
                const yrRows = document.querySelectorAll('#carcassCellChartBody tr').length;
                setCarcassGranularity('MONTH');
                const moRows = document.querySelectorAll('#carcassCellChartBody tr').length;
                setCarcassGranularity('WEEK');
                const wkRows = document.querySelectorAll('#carcassCellChartBody tr').length;
                setCarcassGranularity('DAY');
                const dyRows = document.querySelectorAll('#carcassCellChartBody tr').length;

                setCarcassDatePreset('2025');
                const dy2025Rows = document.querySelectorAll('#carcassCellChartBody tr').length;
                onChangeCarcassMetric('ALL_GRADES');
                const allGradeDatasets = STATE.charts.carcassTrend ? STATE.charts.carcassTrend.data.datasets.length : 0;
                resetCarcassFilters();

                const ymMatrixRows = document.querySelectorAll('#carcassYmMatrixBody tr').length;
                const carcassChartOk = !!document.getElementById('carcassTrendChart') && document.getElementById('carcassTrendChart').width > 0;
                const carcassThSticky = getComputedStyle(document.querySelector('#carcassPeriodCellTable thead th:first-child')).position === 'sticky';

                return {
                    selectedLenAfter6,
                    v3BucketReacted: v3AllCardTxt !== v3FilteredCardTxt,
                    v4ChartLabels,
                    v4MinMaxLinks,
                    stickyChecks,
                    v2PlatSummary,
                    v2MinLinks,
                    v2MaxLinks,
                    carcassChecks: {
                        yrRows,
                        moRows,
                        wkRows,
                        dyRows,
                        dy2025Rows,
                        allGradeDatasets,
                        ymMatrixRows,
                        carcassChartOk,
                        carcassThSticky,
                    },
                };
            }"""
        )
        assert multi_and_v3_v4["selectedLenAfter6"] == 5, f"최대 5개 중복탭 제한 오류: {multi_and_v3_v4}"
        assert multi_and_v3_v4["v3BucketReacted"] is True, "View 3 상단 4개 생산일자 카드가 부위 탭에 반응하지 않음"
        assert len(multi_and_v3_v4["v4ChartLabels"]) == 5, f"View 4 카테고리 차트 5개 부위 미반응: {multi_and_v3_v4}"
        assert multi_and_v3_v4["v2MinLinks"] > 5 and multi_and_v3_v4["v2MaxLinks"] > 5, f"View 2 최저·최고 하이퍼링크 부족: {multi_and_v3_v4}"
        cc = multi_and_v3_v4["carcassChecks"]
        assert cc["yrRows"] >= 3 and cc["moRows"] >= 34 and cc["wkRows"] >= 140 and cc["dyRows"] >= 700, f"지육시세 1년/월/주/일 집계 행 오류: {cc}"
        assert 200 <= cc["dy2025Rows"] <= 270, f"지육시세 2025년 필터 반응 오류: {cc}"
        assert cc["allGradeDatasets"] == 4, f"지육시세 전등급 동시비교 차트 시리즈 오류: {cc}"
        assert cc["ymMatrixRows"] >= 3 and cc["carcassChartOk"] is True and cc["carcassThSticky"] is True, f"지육시세 차트/매트릭스/고정헤더 오류: {cc}"
        for k_st, v_st in multi_and_v3_v4["stickyChecks"].items():
            assert v_st is True, f"시트 스크롤 고정 미적용: {k_st}"
        assert len(js_errors) == 0, f"브라우저 JS 런타임 에러 발생: {js_errors}"

        page_pc.screenshot(path=str(WORK_DIR / "verify_pc_1920.png"), full_page=False)
        pc_metrics["interactive_checks"] = {
            "max5_tabs": multi_and_v3_v4["selectedLenAfter6"],
            "v3_bucket_reacted": multi_and_v3_v4["v3BucketReacted"],
            "v4_chart_labels": multi_and_v3_v4["v4ChartLabels"],
            "v2_min_max_links": [multi_and_v3_v4["v2MinLinks"], multi_and_v3_v4["v2MaxLinks"]],
            "sticky_sheets_all_true": True,
        }
        results["pc_1920x1080"] = pc_metrics

        # 2. Mobile iOS Safari (iPhone 14 Pro: 393x852)
        page_ios = browser.new_page(
            viewport={"width": 393, "height": 852},
            is_mobile=True,
            has_touch=True,
        )
        page_ios.goto(BASE_URL, wait_until="networkidle", timeout=30000)
        page_ios.wait_for_timeout(1000)
        page_ios.evaluate("switchView(4)")
        page_ios.wait_for_timeout(400)

        ios_metrics = page_ios.evaluate(
            """() => {
                const desktopTbl = document.querySelector('#view-4 .desktop-table-only');
                const mobileCards = document.getElementById('productMobileCards');
                const kwInput = document.getElementById('fltKeyword');
                return {
                    scrollWidth: document.documentElement.scrollWidth,
                    clientWidth: document.documentElement.clientWidth,
                    desktopTableDisplay: window.getComputedStyle(desktopTbl).display,
                    mobileCardsDisplay: window.getComputedStyle(mobileCards).display,
                    inputFontSize: window.getComputedStyle(kwInput).fontSize,
                    bottomNavCount: document.querySelectorAll('[data-bnav]').length
                };
            }"""
        )
        assert ios_metrics["scrollWidth"] <= ios_metrics["clientWidth"], f"iOS 가로 스크롤 오버플로우: {ios_metrics}"
        assert ios_metrics["desktopTableDisplay"] == "none", "모바일에서 PC 테이블이 숨겨지지 않음"
        assert ios_metrics["mobileCardsDisplay"] == "block", "모바일 카드뷰가 활성화되지 않음"
        assert ios_metrics["inputFontSize"] == "16px", f"iOS 자동줌 방지 16px 미적용: {ios_metrics['inputFontSize']}"
        assert ios_metrics["bottomNavCount"] == 5, f"모바일 하단 탭바 5개 불일치: {ios_metrics['bottomNavCount']}"
        page_ios.screenshot(path=str(WORK_DIR / "verify_mobile_iphone14pro.png"), full_page=False)
        results["ios_393x852"] = ios_metrics

        # 3. Mobile Android Chrome (Galaxy S23: 360x800)
        page_and = browser.new_page(
            viewport={"width": 360, "height": 800},
            is_mobile=True,
            has_touch=True,
        )
        page_and.goto(BASE_URL, wait_until="networkidle", timeout=30000)
        page_and.wait_for_timeout(1000)
        page_and.evaluate("switchView(2)")
        page_and.wait_for_timeout(400)

        and_metrics = page_and.evaluate(
            """() => ({
                scrollWidth: document.documentElement.scrollWidth,
                clientWidth: document.documentElement.clientWidth,
                mobileCardsCount: document.querySelectorAll('#matrixMobileCards > div').length
            })"""
        )
        assert and_metrics["scrollWidth"] <= and_metrics["clientWidth"], f"Android 가로 오버플로우: {and_metrics}"
        assert and_metrics["mobileCardsCount"] > 0, "Android 품목비교 카드 미렌더링"
        page_and.screenshot(path=str(WORK_DIR / "verify_mobile_galaxys23.png"), full_page=False)
        results["android_360x800"] = and_metrics

        # 4. 금천미트 하이퍼링크 실제 브라우저 클릭/접속 시 상품상세페이지 노출 검증
        sample_ekcm = page_pc.evaluate(
            """() => {
                const ek = (STATE.data?.products || []).find(p => p.source_site === '금천미트' && !p.is_sold_out && p.product_url.includes('/pd/productDetail'));
                return ek ? {url: ek.product_url, raw_name: ek.raw_name} : null;
            }"""
        )
        assert sample_ekcm is not None, "대시보드 내 금천미트 /pd/productDetail 상품 링크 미발견"
        page_ek = browser.new_page(viewport={"width": 1440, "height": 900})
        page_ek.goto(sample_ekcm["url"], wait_until="networkidle", timeout=30000)
        page_ek.wait_for_timeout(1800)
        ek_body = page_ek.evaluate("() => document.body.innerText")
        assert "상품정보제공고시" in ek_body and "kg당" in ek_body, f"금천미트 상세페이지 상품 미노출: {sample_ekcm['url']}"
        results["ekcm_detail_browser_verified"] = sample_ekcm["url"]

        browser.close()

    print("STAGE 3 OK:", json.dumps(results, ensure_ascii=False))
    return results


if __name__ == "__main__":
    s1 = run_stage1_verification()
    s2 = run_stage2_verification()
    s3 = run_stage3_verification()
    print("ALL_3_STAGES_VERIFIED_SUCCESSFULLY")
