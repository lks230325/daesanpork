#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[대산포크 한돈 B2B 인텔리전스] 윈도우 작업 스케줄러 연동 평일(09:00 / 16:00) 안전 크롤링 보장기
- 슈퍼바이저 미실행 시 백그라운드 자동 기동
- 이미 당일 해당 슬롯(09시/16시) 크롤링이 완료되었거나 진행 중이면 중복 실행 없이 즉시 종료
- 미실행 상태일 경우에만 안전하게 수집 실행 후 DB/엑셀 누적
"""

import json
import os
import subprocess
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

DEPLOY_DIR = Path(__file__).resolve().parent
WORK_DIR = DEPLOY_DIR.parent
sys.path.insert(0, str(WORK_DIR))

import crawler_db  # noqa: E402

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def ensure_background_supervisor() -> None:
    sup_script = DEPLOY_DIR / "server_supervisor.py"
    pyw = Path(sys.executable).with_name("pythonw.exe")
    exe = str(pyw if pyw.exists() else Path(sys.executable))
    subprocess.Popen(
        [exe, str(sup_script)],
        cwd=str(WORK_DIR),
        creationflags=CREATE_NO_WINDOW,
    )


def has_slot_run(date_str: str, hour: int) -> bool:
    slot_tag = f"scheduled_{hour:02d}00"
    hour_prefix = f"{date_str} {hour:02d}:"
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


def main() -> int:
    ensure_background_supervisor()
    now = datetime.now()
    if now.weekday() not in (0, 1, 2, 3, 4):
        return 0

    slot_hour = 9 if now.hour < 15 else 16
    if len(sys.argv) > 1 and sys.argv[1] in ("9", "09", "16"):
        slot_hour = int(sys.argv[1])

    date_str = now.strftime("%Y-%m-%d")
    # app.py 내장 스케줄러가 먼저 시작할 수 있도록 8초 대기 후 확인
    time.sleep(8)
    if has_slot_run(date_str, slot_hour):
        return 0

    try:
        with urllib.request.urlopen("http://127.0.0.1:8000/api/scheduler", timeout=5) as resp:
            sched = json.loads(resp.read().decode("utf-8"))
            if sched.get("is_running_now"):
                return 0
    except Exception:
        pass

    if not has_slot_run(date_str, slot_hour):
        trigger_source = f"scheduled_{slot_hour:02d}00"
        try:
            req = urllib.request.Request(
                "http://127.0.0.1:8000/api/crawl",
                data=json.dumps({"trigger_source": trigger_source}).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=600) as resp:
                _ = resp.read()
        except Exception:
            if not has_slot_run(date_str, slot_hour):
                crawler_db.run_full_crawl(trigger_source=trigger_source)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
