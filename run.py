"""검사성적서 관리 시스템 실행기.

    python run.py                 사내망 전체 공개 (0.0.0.0:8000)
    python run.py --local         이 PC 에서만 (127.0.0.1:8000)
    python run.py --port 9000     포트 변경

한국어 Windows 콘솔의 cp949 출력 오류를 막기 위해 UTF-8 로 맞춘다.
"""

from __future__ import annotations

import argparse
import os
import sys


def main() -> None:
    ap = argparse.ArgumentParser(description="검사성적서 관리 시스템")
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--local", action="store_true", help="127.0.0.1 로만 바인딩")
    ap.add_argument("--reload", action="store_true", help="개발용 자동 재시작")
    args = ap.parse_args()

    os.environ.setdefault("PYTHONUTF8", "1")
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    host = "127.0.0.1" if args.local else args.host
    shown = "127.0.0.1" if host in ("0.0.0.0", "127.0.0.1") else host
    print("=" * 62)
    print(" 검사성적서 관리 시스템")
    print(f" 브라우저에서 접속:  http://{shown}:{args.port}")
    if host == "0.0.0.0":
        print(" (사내망의 다른 PC 는 이 PC 의 IP 주소로 접속합니다)")
    print(" 종료: Ctrl+C")
    print("=" * 62)

    import uvicorn
    uvicorn.run("app:app", host=host, port=args.port, reload=args.reload, log_level="info")


if __name__ == "__main__":
    main()
