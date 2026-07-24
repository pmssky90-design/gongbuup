from __future__ import annotations

import json
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
SUMMARY = REPORTS / "production_priority_1_candidate_summary.json"


def main() -> int:
    summary = json.loads(SUMMARY.read_text(encoding="utf-8"))
    if not summary.get("passed"):
        raise RuntimeError("검증을 통과한 후보만 미리보기 서버를 실행할 수 있습니다.")
    candidate = Path(str(summary["candidate_output"])).resolve()
    if candidate.parent != (ROOT / "preview").resolve() or not candidate.is_dir():
        raise RuntimeError(f"후보 경로가 안전 범위를 벗어났습니다: {candidate}")

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        port = int(probe.getsockname()[1])

    log_path = REPORTS / "production_priority_1_candidate_preview_server.log"
    log_handle = log_path.open("ab")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "http.server",
            str(port),
            "--bind",
            "127.0.0.1",
            "--directory",
            str(candidate),
        ],
        cwd=str(ROOT),
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        creationflags=flags,
    )
    url = f"http://localhost:{port}/"
    status = None
    for _ in range(30):
        if process.poll() is not None:
            raise RuntimeError(f"미리보기 서버가 종료되었습니다: exit={process.returncode}")
        try:
            with urllib.request.urlopen(url, timeout=2) as response:
                status = int(response.status)
            break
        except OSError:
            time.sleep(0.2)
    if status != 200:
        process.terminate()
        raise RuntimeError(f"미리보기 HTTP 확인 실패: {status}")

    result = {
        "candidate_output": str(candidate),
        "url": url,
        "port": port,
        "pid": process.pid,
        "http_status": status,
        "deployment": False,
    }
    (REPORTS / "production_priority_1_candidate_preview_server.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
