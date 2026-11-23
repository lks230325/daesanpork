#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[대산포크 한돈 B2B 인텔리전스] 상시 서버 & 무료 고정 HTTPS 배포 URL 슈퍼바이저 (Self-Healing Supervisor)
- 단일 인스턴스 락(127.0.0.1:48899)으로 중복 실행 100% 방지 (윈도우 스케줄러 워치독 연동)
- FastAPI(uvicorn :8000) 상태 상시 감시 및 비정상 종료 시 자동 재기동
- 무료 영구 고정 HTTPS 배포 URL(tunwg + 로컬 WireGuard 고유키 + Let's Encrypt SSL) 상시 유지 (`TUNWG_PATH=deploy/tunwg_keys`)
  -> PC 재부팅·프로세스 재시작 시에도 동일한 고정 도메인(`https://co46iwxtjywrmy3w.l.tunwg.com`) 영구 유지
- 보조 백업 Cloudflare HTTPS 터널(`cloudflared.exe`) 동시 가동으로 이중화 보장
"""

import json
import logging
import os
import re
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

DEPLOY_DIR = Path(__file__).resolve().parent
WORK_DIR = DEPLOY_DIR.parent
LOGS_DIR = DEPLOY_DIR / "logs"
TUNWG_KEYS_DIR = DEPLOY_DIR / "tunwg_keys"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
TUNWG_KEYS_DIR.mkdir(parents=True, exist_ok=True)

STATUS_JSON_PATH = DEPLOY_DIR / "deploy_status.json"
URL_GUIDE_TXT_PATH = WORK_DIR / "외부접속_URL_안내.txt"
TUNWG_EXE = DEPLOY_DIR / "tunwg.exe"
CLOUDFLARED_EXE = DEPLOY_DIR / "cloudflared.exe"

SINGLE_INSTANCE_PORT = 48899
APP_PORT = 8000
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] Supervisor - %(message)s",
    handlers=[
        logging.FileHandler(LOGS_DIR / "supervisor.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("HandonSupervisor")


def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def is_fastapi_healthy() -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{APP_PORT}/api/health", timeout=4) as resp:
            return resp.status == 200
    except Exception:
        return False


def is_remote_url_healthy(url: str) -> bool:
    if not url:
        return False
    try:
        req = urllib.request.Request(
            f"{url.rstrip('/')}/api/health",
            headers={"User-Agent": "HandonSupervisorHealthCheck/1.0"},
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            return resp.status == 200
    except Exception:
        return False


def get_python_exe() -> str:
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        cand = exe.with_name("python.exe")
        if cand.exists():
            return str(cand)
    return str(exe)


class DeploymentSupervisor:
    def __init__(self) -> None:
        self.uvicorn_proc: Optional[subprocess.Popen] = None
        self.fixed_tunnel_proc: Optional[subprocess.Popen] = None
        self.cf_tunnel_proc: Optional[subprocess.Popen] = None
        self.fixed_public_url: str = ""
        self.fixed_started_at: str = ""
        self.backup_cf_url: str = ""
        self.cf_started_at: str = ""
        self.started_at: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._last_remote_probe_ts: float = 0.0
        self._remote_fail_streak: int = 0
        self._lock = threading.Lock()

        # 기존 고정 URL 캐시가 있으면 초기값으로 복원 (키 고정으로 항상 동일 주소)
        if STATUS_JSON_PATH.exists():
            try:
                prev = json.loads(STATUS_JSON_PATH.read_text(encoding="utf-8"))
                prev_fixed = str(prev.get("fixed_https_url") or "")
                if ".tunwg.com" in prev_fixed:
                    self.fixed_public_url = prev_fixed
            except Exception:
                pass

    def write_status_files(self) -> None:
        local_ip = get_local_ip()
        pc_url = f"http://localhost:{APP_PORT}"
        lan_url = f"http://{local_ip}:{APP_PORT}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self._lock:
            fixed_url = self.fixed_public_url
            f_start = self.fixed_started_at
            cf_url = self.backup_cf_url
            c_start = self.cf_started_at

        primary_url = fixed_url or cf_url

        status_data: Dict[str, Any] = {
            "supervisor_pid": os.getpid(),
            "supervisor_started_at": self.started_at,
            "updated_at": now_str,
            "pc_url": pc_url,
            "lan_mobile_url": lan_url,
            "public_https_url": primary_url,
            "url_is_permanent_fixed": bool(fixed_url),
            "fixed_https_url": fixed_url,
            "fixed_tunnel_started_at": f_start,
            "backup_cloudflare_url": cf_url,
            "backup_tunnel_started_at": c_start,
            "fastapi_healthy": is_fastapi_healthy(),
            "schedule": "월~금 09:00 · 16:00 자동수집 (슬랙 미발송)",
        }
        try:
            STATUS_JSON_PATH.write_text(json.dumps(status_data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            logger.warning("상태 JSON 저장 실패: %s", exc)

        guide_lines = [
            "==========================================================================",
            " [대산포크 B2B 인텔리전스] 한돈 4대 플랫폼 대시보드 배포 접속 안내",
            "==========================================================================",
            f"- 상태 갱신 시각          : {now_str}",
            f"- [고정 배포 URL · 무료]  : {fixed_url or '(고정 터널 연결 중...)'}",
            "  * PC 재부팅/재기동 시에도 주소가 변경되지 않는 영구 고정 HTTPS 주소입니다.",
            "  * 즐겨찾기(북마크) 및 모바일 홈화면에 고정 등록하여 사용하세요.",
            f"- [보조 백업 HTTPS URL]   : {cf_url or '(보조 터널 연결 중...)'}",
            f"- 사내 동일 Wi-Fi 접속    : {lan_url}",
            f"- 현재 PC 로컬 접속       : {pc_url}",
            "- 자동 크롤링 스케줄      : 매주 월~금 오전 09:00 / 오후 16:00 (DB 및 엑셀 자동 누적)",
            "==========================================================================",
            "",
        ]
        try:
            URL_GUIDE_TXT_PATH.write_text("\n".join(guide_lines), encoding="utf-8")
        except Exception as exc:
            logger.warning("안내 텍스트 저장 실패: %s", exc)

    def ensure_uvicorn(self) -> None:
        if is_fastapi_healthy():
            return
        if self.uvicorn_proc is not None and self.uvicorn_proc.poll() is None:
            time.sleep(2)
            if is_fastapi_healthy():
                return

        py_exe = get_python_exe()
        uv_log = open(LOGS_DIR / "uvicorn.log", "a", encoding="utf-8")
        cmd = [py_exe, "-X", "utf8", "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", str(APP_PORT)]
        logger.info("FastAPI 서버 기동 시작: %s", " ".join(cmd))
        self.uvicorn_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=uv_log,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
        )
        for _ in range(20):
            time.sleep(0.5)
            if is_fastapi_healthy():
                logger.info("FastAPI 서버 정상 응답 확인 (PID=%s)", self.uvicorn_proc.pid)
                break

    def _watch_fixed_tunnel_output(self, proc: subprocess.Popen) -> None:
        url_regex = re.compile(r"(https://[a-zA-Z0-9\-]+\.l\.tunwg\.com)")
        tun_log_path = LOGS_DIR / "fixed_tunnel.log"
        with open(tun_log_path, "a", encoding="utf-8") as tlog:
            assert proc.stdout is not None
            for raw_line in iter(proc.stdout.readline, ""):
                if not raw_line:
                    break
                tlog.write(raw_line)
                tlog.flush()
                m = url_regex.search(raw_line)
                if m:
                    found_url = m.group(1).strip()
                    with self._lock:
                        changed = self.fixed_public_url != found_url
                        self.fixed_public_url = found_url
                        self.fixed_started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self._remote_fail_streak = 0
                    if changed:
                        logger.info("무료 고정 HTTPS 배포 터널 개통 완료: %s", found_url)
                    self.write_status_files()

    def ensure_fixed_tunnel(self) -> None:
        if not TUNWG_EXE.exists():
            logger.error("tunwg.exe 파일이 존재하지 않습니다: %s", TUNWG_EXE)
            return

        # 프로세스가 살아 있더라도 3분마다 외부 HTTPS 엔드포인트 생존 확인 후 응답 불가 누적 시 자동 복구
        if self.fixed_tunnel_proc is not None and self.fixed_tunnel_proc.poll() is None:
            now_ts = time.time()
            if self.fixed_public_url and (now_ts - self._last_remote_probe_ts >= 180):
                self._last_remote_probe_ts = now_ts
                if is_remote_url_healthy(self.fixed_public_url):
                    self._remote_fail_streak = 0
                    return
                self._remote_fail_streak += 1
                logger.warning(
                    "고정 HTTPS 터널 응답 지연 감지 (%d/2): %s",
                    self._remote_fail_streak,
                    self.fixed_public_url,
                )
                if self._remote_fail_streak >= 2:
                    logger.warning("고정 HTTPS 터널 재연결 실행 (동일 고정 URL 유지)")
                    try:
                        self.fixed_tunnel_proc.terminate()
                        self.fixed_tunnel_proc.wait(timeout=5)
                    except Exception:
                        try:
                            self.fixed_tunnel_proc.kill()
                        except Exception:
                            pass
                    self._remote_fail_streak = 0
                else:
                    return
            else:
                return

        env = os.environ.copy()
        env["TUNWG_PATH"] = str(TUNWG_KEYS_DIR)
        cmd = [str(TUNWG_EXE), f"--forward=http://127.0.0.1:{APP_PORT}"]
        logger.info("무료 영구 고정 HTTPS 터널(tunwg) 프로세스 시작 (키 경로: %s)", TUNWG_KEYS_DIR)
        self.fixed_tunnel_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            creationflags=CREATE_NO_WINDOW,
        )
        self._last_remote_probe_ts = time.time()
        t = threading.Thread(target=self._watch_fixed_tunnel_output, args=(self.fixed_tunnel_proc,), daemon=True)
        t.start()

    def _watch_cf_tunnel_output(self, proc: subprocess.Popen) -> None:
        url_regex = re.compile(r"(https://[a-zA-Z0-9\-]+\.trycloudflare\.com)")
        tun_log_path = LOGS_DIR / "tunnel.log"
        with open(tun_log_path, "a", encoding="utf-8") as tlog:
            assert proc.stderr is not None
            for raw_line in iter(proc.stderr.readline, ""):
                if not raw_line:
                    break
                tlog.write(raw_line)
                tlog.flush()
                m = url_regex.search(raw_line)
                if m:
                    found_url = m.group(1).strip()
                    with self._lock:
                        changed = self.backup_cf_url != found_url
                        self.backup_cf_url = found_url
                        self.cf_started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if changed:
                        logger.info("보조 Cloudflare HTTPS 터널 개통 완료: %s", found_url)
                        self.write_status_files()

    def ensure_cf_backup_tunnel(self) -> None:
        if not CLOUDFLARED_EXE.exists():
            return
        if self.cf_tunnel_proc is not None and self.cf_tunnel_proc.poll() is None:
            return

        cmd = [
            str(CLOUDFLARED_EXE),
            "tunnel",
            "--url",
            f"http://127.0.0.1:{APP_PORT}",
            "--no-autoupdate",
        ]
        logger.info("보조 Cloudflare HTTPS 터널 프로세스 시작")
        self.cf_tunnel_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=CREATE_NO_WINDOW,
        )
        t = threading.Thread(target=self._watch_cf_tunnel_output, args=(self.cf_tunnel_proc,), daemon=True)
        t.start()

    def run_forever(self) -> None:
        logger.info("슈퍼바이저 상시 가동 시작 (PID=%d)", os.getpid())
        self.ensure_uvicorn()
        self.ensure_fixed_tunnel()
        self.ensure_cf_backup_tunnel()
        self.write_status_files()

        while True:
            time.sleep(15)
            try:
                self.ensure_uvicorn()
                self.ensure_fixed_tunnel()
                self.ensure_cf_backup_tunnel()
                self.write_status_files()
            except Exception as exc:
                logger.error("슈퍼바이저 루프 오류: %s", exc)


def main() -> int:
    lock_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        lock_sock.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
        lock_sock.listen(1)
    except OSError:
        # 이미 슈퍼바이저가 정상 실행 중이므로 조용히 종료
        return 0

    sup = DeploymentSupervisor()
    sup.run_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
