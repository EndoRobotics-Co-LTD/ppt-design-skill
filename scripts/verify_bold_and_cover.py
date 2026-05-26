"""두 변경사항 동시 검증:
1) Chrome 상단 바의 제목 Bold 적용 (본문 슬라이드)
2) 표지 제목 크기↑ + Bold + 완전 중앙배치

기존 표지/본문 테스트 결과를 정리하고 새로 만든다.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def _delete_after_index(session, since_idx: int):
    """since_idx 이후 슬라이드 모두 삭제 (끝→앞 순)."""
    deleted = 0
    for idx in range(session.slide_count, since_idx, -1):
        try:
            session._prs.Slides(idx).Delete()
            deleted += 1
        except Exception:
            pass
    return deleted


def main():
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 시작 슬라이드 수: {session.slide_count}")

    # Reference 원본 6장만 남기고 정리 (테스트 슬라이드 제거)
    removed = _delete_after_index(session, 6)
    print(f"[CLEANUP] 테스트 슬라이드 {removed}장 삭제. 현재 {session.slide_count}장")

    # 본문 한 장 — chrome 제목 bold 확인용
    print("[STEP 1] 본문 1단 슬라이드 추가 (chrome 제목 bold 확인)...")
    session.add_single_column(
        title="핵심 성과 요약",
        bullets=[
            "매출 전년 동기 대비 18% 증가",
            "신규 고객 12개사 확보",
            "EBITDA 마진 14% 유지",
        ],
    )

    # KPI 카드 — chrome 제목 bold 추가 확인
    print("[STEP 2] KPI 카드 슬라이드 추가...")
    session.add_kpi_cards(
        title="핵심 지표",
        cards=[
            {"label": "매출 (YoY)", "value": "+18%", "note": "전년 동기 대비"},
            {"label": "신규 고객", "value": "12사", "note": "글로벌 5개사 포함"},
            {"label": "EBITDA 마진", "value": "14%", "note": "원가 개선 효과"},
        ],
    )

    # 표지 — bold + 크기↑ + 중앙배치 확인
    print("[STEP 3] 표지 슬라이드 추가 (제목 크기↑ + bold + 중앙)...")
    cover_idx = session.add_cover(
        title="2026년 1분기 사업 보고",
        presenter="이은상",
        organization="전략기획팀",
        date="2026.05.26",
    )

    print(f"\n[DONE] 슬라이드 {session.slide_count}장")
    print(f"       - 슬라이드 1~6: Reference 원본")
    print(f"       - 슬라이드 7: 본문 1단 (chrome 제목 bold 확인)")
    print(f"       - 슬라이드 8: KPI 카드 (chrome 제목 bold 확인)")
    print(f"       - 슬라이드 {cover_idx}: 표지 (큰 bold 제목 + 중앙)")
    print(f"       파일은 저장되지 않았습니다.")


if __name__ == "__main__":
    main()
