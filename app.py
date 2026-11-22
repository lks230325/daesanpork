#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[대산포크 한돈 B2B 인텔리전스] FastAPI 반응형 웹 대시보드 서버 & 일일 자동 크롤링 스케줄러
- PC & 모바일(iOS Safari / Android Chrome) 완벽 반응형 웹 애플리케이션
- 4대 축산 플랫폼(미트박스, 금천미트, 미트프렌즈, 탑미트) 시계열 누적 DB (`handon_master.db`) 연동
- 매일 오전 09:00 / 오후 16:00 자동 크롤링 및 DB·엑셀 누적(Append) 백그라운드 스케줄러 내장
- 중점관리 체크박스 실시간 DB 저장 (`POST/PATCH /api/tracked`) 및 누적 엑셀(`=HYPERLINK`) 다운로드
- 슬랙(Slack) 웹훅 발송 100% 제외 (0회 호출 보장)
"""

import json
import logging
import os
import socket
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

import crawler_db

logger = logging.getLogger("HandonAutoScheduler")

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = WORK_DIR / "templates"
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

DAILY_SCHEDULE_HOURS: List[int] = [9, 16]  # 평일(월~금) 09:00, 16:00 자동 크롤링
WEEKDAYS_ONLY = (0, 1, 2, 3, 4)  # 0:월, 1:화, 2:수, 3:목, 4:금 (5:토, 6:일 제외)
_SCHEDULER_LOCK = threading.Lock()
_CRAWL_EXEC_LOCK = threading.Lock()
_SCHEDULER_STATE: Dict[str, Any] = {
    "enabled": True,
    "schedule_days": "월~금 (토·일 제외)",
    "weekdays_only": True,
    "schedule_times": ["09:00", "16:00"],
    "slack_enabled": False,
    "is_running_now": False,
    "last_executed_slot": None,
    "last_auto_run_at": None,
    "last_auto_run_id": None,
    "thread_alive": False,
}
_SCHEDULER_THREAD: Optional[threading.Thread] = None

app = FastAPI(
    title="대산포크 한돈 4대 플랫폼 통합 인텔리전스 대시보드",
    description="미트박스(근본한돈 분리)·금천미트·미트프렌즈·탑미트 시계열 누적 관리 및 실시간 문제해결 웹 시스템",
    version="2.1.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def compute_next_scheduled_time(now: Optional[datetime] = None) -> str:
    now_dt = now or datetime.now()
    candidates: List[datetime] = []
    for day_offset in range(0, 5):
        base_date = (now_dt + timedelta(days=day_offset)).date()
        if base_date.weekday() not in WEEKDAYS_ONLY:
            continue
        for hr in DAILY_SCHEDULE_HOURS:
            cand = datetime(base_date.year, base_date.month, base_date.day, hr, 0, 0)
            if cand > now_dt:
                candidates.append(cand)
    candidates.sort()
    return candidates[0].strftime("%Y-%m-%d %H:%M:%S") if candidates else "-"


def get_scheduler_status() -> Dict[str, Any]:
    with _SCHEDULER_LOCK:
        st = dict(_SCHEDULER_STATE)
    st["next_scheduled_at"] = compute_next_scheduled_time()
    st["thread_alive"] = bool(_SCHEDULER_THREAD and _SCHEDULER_THREAD.is_alive())
    return st


def _has_slot_already_run_in_db(date_str: str, hour: int) -> bool:
    """동일 일자·시간대(09시 또는 16시)에 이미 자동 크롤링이 수행되었는지 DB에서 확인"""
    slot_tag = f"scheduled_{hour:02d}00"
    hour_prefix = f"{date_str} {hour:02d}:"
    try:
        conn = crawler_db.get_db_connection()
        row = conn.execute(
            """
            SELECT run_id FROM crawl_runs
            WHERE (trigger_source = ? AND collected_date = ?)
               OR (collected_at LIKE ? AND trigger_source LIKE 'scheduled%')
            LIMIT 1
            """,
            (slot_tag, date_str, f"{hour_prefix}%"),
        ).fetchone()
        conn.close()
        return row is not None
    except Exception:
        return False


def _daily_scheduler_loop() -> None:
    """
    평일(월~금, 토·일 제외) 09:00 및 16:00 정기 자동 크롤링 실행:
    - 정각 실행은 물론, PC 재부팅 등으로 09:00/16:00 정각을 넘겨 기동된 경우에도 당일 미실행 슬롯을 즉시 보정 수집(Catch-up)
    """
    logger.info("평일(월~금, 토·일 제외) 자동 크롤링 스케줄러 가동 시작 (09:00, 16:00 / 슬랙 알림 제외)")
    while True:
        try:
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            due_hour: Optional[int] = None
            if now.weekday() in WEEKDAYS_ONLY:
                if 9 <= now.hour < 16:
                    due_hour = 9
                elif now.hour >= 16:
                    due_hour = 16

            if due_hour is not None:
                slot_key = f"{date_str}_{due_hour:02d}00"
                with _SCHEDULER_LOCK:
                    already_mem = _SCHEDULER_STATE.get("last_executed_slot") == slot_key
                if not already_mem and not _has_slot_already_run_in_db(date_str, due_hour):
                    if _CRAWL_EXEC_LOCK.acquire(blocking=False):
                        try:
                            with _SCHEDULER_LOCK:
                                _SCHEDULER_STATE["is_running_now"] = True
                                _SCHEDULER_STATE["last_executed_slot"] = slot_key
                            trigger_name = f"scheduled_{due_hour:02d}00"
                            logger.info("[자동 스케줄러] %s 정기 크롤링 시작 (%s)", slot_key, trigger_name)
                            res = crawler_db.run_full_crawl(trigger_source=trigger_name)
                            with _SCHEDULER_LOCK:
                                _SCHEDULER_STATE["last_auto_run_at"] = res.get("collected_at")
                                _SCHEDULER_STATE["last_auto_run_id"] = res.get("run_id")
                            logger.info("[자동 스케줄러] %s 정기 크롤링 완료: Run #%s", slot_key, res.get("run_id"))
                        except Exception as exc:
                            logger.error("[자동 스케줄러] 크롤링 오류: %s", exc)
                        finally:
                            with _SCHEDULER_LOCK:
                                _SCHEDULER_STATE["is_running_now"] = False
                            _CRAWL_EXEC_LOCK.release()
        except Exception as outer_exc:
            logger.error("[자동 스케줄러 루프 예외] %s", outer_exc)
        time.sleep(20)


def ensure_scheduler_started() -> None:
    global _SCHEDULER_THREAD
    if _SCHEDULER_THREAD is None or not _SCHEDULER_THREAD.is_alive():
        _SCHEDULER_THREAD = threading.Thread(
            target=_daily_scheduler_loop,
            name="HandonDailyCrawlScheduler_09_16",
            daemon=True,
        )
        _SCHEDULER_THREAD.start()
        with _SCHEDULER_LOCK:
            _SCHEDULER_STATE["thread_alive"] = True


class TrackedToggleRequest(BaseModel):
    item_key: str
    is_active: Optional[bool] = None
    memo: Optional[str] = None


class CrawlTriggerRequest(BaseModel):
    trigger_source: str = "web_manual"


@app.on_event("startup")
def on_startup() -> None:
    crawler_db.init_db()
    ensure_scheduler_started()


def get_deploy_status_info() -> Dict[str, Any]:
    status_file = WORK_DIR / "deploy" / "deploy_status.json"
    if status_file.exists():
        try:
            return json.loads(status_file.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request) -> HTMLResponse:
    local_ip = get_local_ip()
    dep = get_deploy_status_info()
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "local_ip": local_ip,
            "pc_url": "http://localhost:8000",
            "mobile_url": f"http://{local_ip}:8000",
            "public_https_url": dep.get("public_https_url", ""),
        },
    )


@app.get("/api/health")
async def api_health() -> Dict[str, Any]:
    crawler_db.init_db()
    ensure_scheduler_started()
    conn = crawler_db.get_db_connection()
    run_cnt = conn.execute("SELECT COUNT(*) FROM crawl_runs").fetchone()[0]
    snap_cnt = conn.execute("SELECT COUNT(*) FROM product_snapshots").fetchone()[0]
    tracked_cnt = conn.execute("SELECT COUNT(*) FROM tracked_items WHERE is_active = 1").fetchone()[0]
    black_cnt = conn.execute(
        "SELECT COUNT(*) FROM product_snapshots WHERE raw_name LIKE '%흑돼지%' OR raw_name LIKE '%흑돈%' OR raw_name LIKE '%버크셔%' OR raw_name LIKE '%우리흑돈%'"
    ).fetchone()[0]
    conn.close()
    dep = get_deploy_status_info()
    return {
        "status": "ok",
        "db_path": str(crawler_db.DB_PATH),
        "excel_path": str(crawler_db.EXCEL_MASTER_PATH),
        "run_count": run_cnt,
        "snapshot_count": snap_cnt,
        "tracked_active_count": tracked_cnt,
        "black_pork_remaining_count": black_cnt,
        "slack_webhook_calls": 0,
        "scheduler": get_scheduler_status(),
        "local_ip": get_local_ip(),
        "public_https_url": dep.get("public_https_url", ""),
    }


@app.get("/api/scheduler")
async def api_scheduler_status() -> Dict[str, Any]:
    ensure_scheduler_started()
    return get_scheduler_status()


@app.get("/api/dashboard")
async def api_dashboard(run_id: Optional[int] = Query(default=None)) -> JSONResponse:
    ensure_scheduler_started()
    payload = crawler_db.get_dashboard_payload(run_id=run_id)
    local_ip = get_local_ip()
    sched = get_scheduler_status()
    dep = get_deploy_status_info()
    payload["server_info"] = {
        "pc_url": "http://localhost:8000",
        "mobile_url": f"http://{local_ip}:8000",
        "public_https_url": dep.get("public_https_url", ""),
        "local_ip": local_ip,
        "slack_webhook_enabled": False,
        "slack_sent_cnt": 0,
        "scheduler": sched,
    }
    return JSONResponse(content=payload)


@app.get("/api/tracked")
async def api_get_tracked(run_id: Optional[int] = Query(default=None)) -> JSONResponse:
    payload = crawler_db.get_dashboard_payload(run_id=run_id)
    return JSONResponse(
        content={
            "status": "ok",
            "tracked_active_cnt": len(payload.get("tracked_items", [])),
            "tracked_items": payload.get("tracked_items", []),
        }
    )


@app.post("/api/tracked")
@app.patch("/api/tracked")
async def api_toggle_tracked(req: TrackedToggleRequest) -> JSONResponse:
    try:
        updated = crawler_db.toggle_tracked_item(
            item_key=req.item_key,
            is_active=req.is_active,
            memo=req.memo,
        )
        # 엑셀 대장의 Sheet3(중점관리 체크상품 누적추이)도 최신 상태로 동기화
        try:
            crawler_db.export_cumulative_excel()
        except Exception:
            pass
        payload = crawler_db.get_dashboard_payload()
        return JSONResponse(
            content={
                "status": "ok",
                "updated_item": updated,
                "tracked_active_cnt": len(payload.get("tracked_items", [])),
                "tracked_items": payload.get("tracked_items", []),
            }
        )
    except ValueError as ve:
        raise HTTPException(status_code=404, detail=str(ve))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"중점관리 상태 저장 오류: {exc}")


@app.post("/api/crawl")
async def api_trigger_crawl(req: Optional[CrawlTriggerRequest] = None) -> JSONResponse:
    source = req.trigger_source if req else "web_manual"
    if not _CRAWL_EXEC_LOCK.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="현재 크롤링 작업이 이미 진행 중입니다.")
    try:
        crawl_res = crawler_db.run_full_crawl(trigger_source=source)
        payload = crawler_db.get_dashboard_payload(run_id=crawl_res["run_id"])
        payload["server_info"] = {
            "pc_url": "http://localhost:8000",
            "mobile_url": f"http://{get_local_ip()}:8000",
            "local_ip": get_local_ip(),
            "slack_webhook_enabled": False,
            "slack_sent_cnt": 0,
            "scheduler": get_scheduler_status(),
        }
        return JSONResponse(
            content={
                "status": "ok",
                "crawl_result": crawl_res,
                "dashboard": payload,
            }
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"크롤링 실행 오류: {exc}")
    finally:
        _CRAWL_EXEC_LOCK.release()


@app.get("/api/export/excel")
async def api_export_excel(run_id: Optional[int] = Query(default=None)) -> FileResponse:
    try:
        excel_path = crawler_db.export_cumulative_excel(run_id=run_id)
        filename = f"한돈_누적관리대장_{run_id or 'latest'}.xlsx"
        return FileResponse(
            path=str(excel_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"엑셀 파일 생성 실패: {exc}")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
