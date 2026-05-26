"""디자인 시스템 전체 검증 — 표지 + 12종 본문 중 대표 8종을 한꺼번에 생성.

기존 테스트 슬라이드를 모두 제거하고, Reference 원본 6장 뒤에 새로 만든다.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


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
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 시작 슬라이드 수: {session.slide_count}")

    cleanup_chart_workbooks(session)

    # 7번 이후 정리
    removed = 0
    for idx in range(session.slide_count, 6, -1):
        try:
            session._prs.Slides(idx).Delete()
            removed += 1
        except Exception:
            pass
    print(f"[CLEANUP] 테스트 슬라이드 {removed}장 삭제. 현재 {session.slide_count}장")

    # 표지
    print("[1] 표지...")
    session.add_cover(
        title="2026년 1분기 사업 보고",
        presenter="이은상",
        organization="전략기획팀",
        date="2026.05.26",
    )

    print("[2] 목차...")
    session.add_agenda(
        items=["회사 소개", "1분기 실적", "주요 전략", "시장 분석", "Q&A"],
    )

    print("[3] 섹션 구분...")
    session.add_section_divider(
        title="1분기 실적",
        number=1,
        subtitle="2026 Q1 Performance Review",
    )

    print("[4] 본문 1단...")
    session.add_single_column(
        title="핵심 성과 요약",
        bullets=[
            "매출 전년 동기 대비 18% 증가, 분기 최대 실적 달성",
            "신규 고객 12개사 확보 (글로벌 5개사 포함)",
            {"text": "EBITDA 마진 14% 유지", "level": 0},
            {"text": "원가 절감 프로젝트 효과 본격 반영", "level": 1},
            "다음 분기 신제품 출시 준비 완료",
        ],
    )

    print("[5] 본문 2단...")
    session.add_two_column(
        title="시장 환경 변화",
        left_sub="기회 요인",
        left_bullets=[
            "글로벌 헬스케어 수요 확대",
            "정부 의료기기 R&D 지원 강화",
            "동남아 시장 진입 장벽 완화",
        ],
        right_sub="리스크 요인",
        right_bullets=[
            "주요 부품 가격 상승",
            "환율 변동성 확대",
            "경쟁사 신제품 출시 임박",
        ],
    )

    print("[6] KPI 카드...")
    session.add_kpi_cards(
        title="핵심 지표",
        cards=[
            {"label": "매출 (YoY)", "value": "+18%", "note": "전년 동기 대비"},
            {"label": "신규 고객", "value": "12사", "note": "글로벌 5개사 포함"},
            {"label": "EBITDA 마진", "value": "14%", "note": "원가 개선 효과"},
        ],
    )

    print("[7] 비교 (Before/After)...")
    session.add_comparison(
        title="프로세스 개선 효과",
        left={
            "label": "Before",
            "title": "수기 보고서 작성",
            "bullets": [
                "주간 보고 작성 평균 4시간",
                "디자인 편차 심함",
                "개정 시 전면 재작성",
            ],
        },
        right={
            "label": "After",
            "title": "PPTMaker AI 자동화",
            "bullets": [
                "주간 보고 작성 30분 이내",
                "전사 표준 디자인 자동 적용",
                "데이터만 갱신, 즉시 반영",
            ],
        },
        arrow_label="PPTMaker 도입",
    )

    print("[8] 핵심 메시지...")
    session.add_highlight_quote(
        message="2026년, 글로벌 시장 진출 원년",
        attribution="대표이사",
    )

    print("[9] 프로세스 다이어그램...")
    session.add_process_flow(
        title="PPTMaker 도입 로드맵",
        steps=[
            {"label": "01", "title": "디자인 토큰", "note": "Reference 분석 완료"},
            {"label": "02", "title": "레이아웃 12종", "note": "본문 빌더 구현"},
            {"label": "03", "title": "사내 PoC", "note": "전략기획팀 우선"},
            {"label": "04", "title": "전사 배포", "note": "Q2 내 완료 목표"},
        ],
    )

    print(f"\n[DONE] 총 {session.slide_count}장")
    print("       슬라이드 1~6: Reference 원본")
    print("       슬라이드 7: 표지 (Cover)")
    print("       슬라이드 8~15: 12종 중 8종 본문")
    print("       파일은 저장되지 않았습니다.")


if __name__ == "__main__":
    main()
