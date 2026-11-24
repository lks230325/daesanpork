#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Pages 정적 배포 빌드 스크립트 (build_gh_pages.py)
- 로컬 실행 시: SQLite(`handon_master.db`)에서 최신 대시보드 및 회차별 JSON + `한돈_누적관리대장.xlsx`를 `static_data/`로 추출한 뒤 `dist/` 정적 사이트 생성
- GitHub Actions 실행 시: 저장소에 커밋된 `static_data/`와 `templates/index.html`을 결합하여 `dist/`(`https://<아이디>.github.io/<저장소명>/` 하위 경로 완벽 호환) 빌드
"""

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any, Dict

from jinja2 import Environment, FileSystemLoader, select_autoescape

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORK_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = WORK_DIR / "templates"
STATIC_DATA_DIR = WORK_DIR / "static_data"
DIST_DIR = WORK_DIR / "dist"
DB_PATH = WORK_DIR / "handon_master.db"
EXCEL_PATH = WORK_DIR / "한돈_누적관리대장.xlsx"


DEFAULT_GH_PAGES_URL = "https://lks230325.github.io/daesanpork/"


def export_db_to_static_data(public_pages_url: str = "") -> Dict[str, Any]:
    """로컬 SQLite DB가 존재할 경우 static_data/ 폴더에 최신 JSON 및 회차별 스냅샷·엑셀 추출"""
    if not public_pages_url:
        public_pages_url = os.environ.get("GH_PAGES_URL", "").strip() or DEFAULT_GH_PAGES_URL
    if not DB_PATH.exists():
        latest_json = STATIC_DATA_DIR / "dashboard_latest.json"
        if latest_json.exists():
            return json.loads(latest_json.read_text(encoding="utf-8"))
        raise FileNotFoundError(f"DB 및 static_data 파일이 모두 없습니다: {DB_PATH}")

    import crawler_db

    STATIC_DATA_DIR.mkdir(parents=True, exist_ok=True)
    runs_dir = STATIC_DATA_DIR / "runs"
    runs_dir.mkdir(parents=True, exist_ok=True)

    latest_payload = crawler_db.get_dashboard_payload(run_id=None, db_path=DB_PATH)
    server_info = {
        "pc_url": public_pages_url or "./",
        "mobile_url": "GitHub Pages 정적 배포",
        "public_https_url": public_pages_url,
        "local_ip": "github.io",
        "slack_webhook_enabled": False,
        "slack_sent_cnt": 0,
        "scheduler": {
            "enabled": True,
            "schedule_days": "월~금 (토·일 제외)",
            "weekdays_only": True,
            "schedule_times": ["09:00", "16:00"],
            "slack_enabled": False,
            "next_scheduled_at": "평일 09:00 · 16:00 자동동기화",
        },
    }
    latest_payload["server_info"] = server_info

    (STATIC_DATA_DIR / "dashboard_latest.json").write_text(
        json.dumps(latest_payload, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    # 최근 10개 회차 스냅샷 개별 JSON 저장 (드롭다운 회차 전환 지원)
    runs_list = latest_payload.get("runs") or []
    keep_run_ids = set()
    for r_meta in runs_list[:10]:
        rid = int(r_meta["run_id"])
        keep_run_ids.add(rid)
        r_payload = crawler_db.get_dashboard_payload(run_id=rid, db_path=DB_PATH)
        r_payload["server_info"] = server_info
        # 지육시세는 대용량 중복 방지를 위해 최신 회차와 동일하게 유지하되 정상 작동 보장
        (runs_dir / f"run_{rid}.json").write_text(
            json.dumps(r_payload, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )

    for old_f in runs_dir.glob("run_*.json"):
        try:
            num = int(old_f.stem.split("_")[1])
            if num not in keep_run_ids:
                old_f.unlink()
        except Exception:
            pass

    if not EXCEL_PATH.exists():
        crawler_db.export_cumulative_excel(db_path=DB_PATH)
    if EXCEL_PATH.exists():
        shutil.copy2(EXCEL_PATH, STATIC_DATA_DIR / "한돈_누적관리대장.xlsx")

    return latest_payload


def build_dist(public_pages_url: str = "") -> Path:
    """templates/index.html + static_data/ -> dist/ 정적 웹사이트 빌드"""
    env_pages_url = os.environ.get("GH_PAGES_URL", "").strip()
    effective_url = public_pages_url or env_pages_url

    if DB_PATH.exists() and "--from-static-only" not in sys.argv:
        payload = export_db_to_static_data(public_pages_url=effective_url)
    else:
        latest_json = STATIC_DATA_DIR / "dashboard_latest.json"
        if not latest_json.exists():
            raise FileNotFoundError(f"정적 데이터 파일이 없습니다: {latest_json}")
        payload = json.loads(latest_json.read_text(encoding="utf-8"))
        if effective_url and "server_info" in payload:
            payload["server_info"]["public_https_url"] = effective_url
            (STATIC_DATA_DIR / "dashboard_latest.json").write_text(
                json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
                encoding="utf-8",
            )

    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    dist_data_dir = DIST_DIR / "data"
    shutil.copytree(STATIC_DATA_DIR, dist_data_dir)

    # Jinja2 템플릿을 정적 HTML로 렌더링 (하위 경로 https://<id>.github.io/<repo>/ 상대경로 100% 호환)
    j_env = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
    )
    tmpl = j_env.get_template("index.html")
    rendered_html = tmpl.render(
        pc_url=effective_url or "./",
        mobile_url=effective_url or "GitHub Pages",
        public_https_url=effective_url,
    )

    # <head> 내부에 정적 페이지 모드 플래그 주입 (로컬 서브패스 테스트 및 GitHub Pages 즉시 활성화)
    static_flag_script = (
        "  <script>window.__GH_PAGES_STATIC__ = true;</script>\n</head>"
    )
    rendered_html = rendered_html.replace("</head>", static_flag_script, 1)

    (DIST_DIR / "index.html").write_text(rendered_html, encoding="utf-8")
    (DIST_DIR / ".nojekyll").write_text("", encoding="utf-8")
    robots_txt = (
        "User-agent: Googlebot\nDisallow: /\n\n"
        "User-agent: Yeti\nDisallow: /\n\n"
        "User-agent: NaverBot\nDisallow: /\n\n"
        "User-agent: Daumoa\nDisallow: /\n\n"
        "User-agent: Bingbot\nDisallow: /\n\n"
        "User-agent: GPTBot\nDisallow: /\n\n"
        "User-agent: *\nDisallow: /\n"
    )
    (DIST_DIR / "robots.txt").write_text(robots_txt, encoding="utf-8")

    print(
        f"[BUILD_OK] dist 생성 완료: index.html ({len(rendered_html):,} bytes), "
        f"최신 회차 #{payload.get('current_run', {}).get('run_id')}, "
        f"상품 {len(payload.get('products', [])):,}종, "
        f"지육시세 {payload.get('carcass_sise', {}).get('total_days', 0):,}일"
    )
    return DIST_DIR


def sync_static_data_to_github(run_id: int = 0, collected_at: str = "") -> bool:
    """정기 크롤링 완료 후 static_data/를 갱신하고 Git 원격 저장소(origin main)가 연결된 경우 자동 커밋·푸시"""
    import subprocess

    try:
        export_db_to_static_data()
        git_dir = WORK_DIR / ".git"
        if not git_dir.exists():
            return False
        rem = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            cwd=str(WORK_DIR),
            capture_output=True,
            text=True,
            timeout=10,
        )
        if rem.returncode != 0 or not rem.stdout.strip():
            return False

        subprocess.run(["git", "add", "static_data"], cwd=str(WORK_DIR), check=False, timeout=20)
        diff = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            cwd=str(WORK_DIR),
            capture_output=True,
            timeout=10,
        )
        if diff.returncode != 0:
            msg = f"chore(data): sync crawl snapshot #{run_id} ({collected_at})"
            subprocess.run(["git", "commit", "-m", msg], cwd=str(WORK_DIR), check=False, timeout=20)
            subprocess.run(["git", "push", "origin", "main"], cwd=str(WORK_DIR), check=False, timeout=60)
        return True
    except Exception as exc:
        print(f"[WARN] GitHub Pages 데이터 자동 푸시 건너뜀: {exc}")
        return False


if __name__ == "__main__":
    build_dist()
