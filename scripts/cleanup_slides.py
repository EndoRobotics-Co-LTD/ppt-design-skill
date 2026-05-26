"""중간 테스트 시도로 생성된 쓰레기 슬라이드 정리.

전제:
- 슬라이드 1~6: Reference 원본 (보존)
- 슬라이드 27 이후: 마지막 chrome 통일 테스트 결과 (보존)
- 슬라이드 7~26: 중간 테스트 시도 — 일관성 없는 쓰레기 (삭제)

사용:
    python scripts/cleanup_slides.py [keep_until] [delete_after]
    기본값: 1~6 보존, 27 이후 보존, 7~26 삭제

안전 장치:
- 파일은 저장하지 않음 — 사용자가 결과 확인 후 직접 저장
- 차트 데이터 워크북 먼저 정리
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"

# 기본: 1~6 보존, 27 이후 보존, 7~26 삭제
KEEP_HEAD = 6        # 처음 N장 보존
KEEP_TAIL_FROM = 27  # 이 슬라이드 이후 보존


def cleanup_chart_workbooks(session):
    closed = 0
    for i in range(1, session.slide_count + 1):
        slide = session._prs.Slides(i)
        for shape in slide.Shapes:
            try:
                if shape.HasChart == -1:
                    try:
                        shape.Chart.ChartData.Workbook.Close(False)
                        closed += 1
                    except Exception:
                        pass
            except Exception:
                pass
    if closed:
        print(f"[CLEANUP] 차트 워크북 {closed}개 닫음")


def main():
    head = int(sys.argv[1]) if len(sys.argv) > 1 else KEEP_HEAD
    tail_from = int(sys.argv[2]) if len(sys.argv) > 2 else KEEP_TAIL_FROM

    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 현재 슬라이드 수: {session.slide_count}")
    print(f"[INFO] 보존: 1~{head} + {tail_from}~끝")
    print(f"[INFO] 삭제 대상: {head + 1}~{tail_from - 1}")

    cleanup_chart_workbooks(session)

    # 끝 → 앞 순으로 삭제 (인덱스 시프트 방지)
    deleted = 0
    for idx in range(tail_from - 1, head, -1):
        try:
            session._prs.Slides(idx).Delete()
            deleted += 1
        except Exception as e:
            print(f"[WARN] 슬라이드 {idx} 삭제 실패: {e}")

    print(f"[DONE] {deleted}장 삭제. 현재 슬라이드 수: {session.slide_count}")
    print(f"       1~{head}: 원본, {head + 1}~{session.slide_count}: 최종 테스트 결과")
    print(f"       파일은 저장되지 않았습니다 — 결과 확인 후 직접 저장하세요.")


if __name__ == "__main__":
    main()
