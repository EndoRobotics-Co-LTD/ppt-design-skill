"""12종 레이아웃 라이브 검증 — 대표 8종을 실제 PowerPoint에 즉시 생성.

실행:
    python scripts/verify_layouts.py

PowerPoint가 열린 상태(또는 자동 실행)에서 Reference 템플릿에 다음 슬라이드를 추가:
1. 목차 (A1)
2. 섹션 구분 (A2)
3. 본문 1단 (A3)
4. 본문 2단 (A4)
5. KPI 카드 (B1)
6. 차트 (B3)
7. 비교 (B4)
8. 프로세스 다이어그램 (C3)

파일은 저장하지 않음 — 사용자가 화면으로 확인 후 결정.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def cleanup_open_chart_workbooks(session=None):
    """이전 실행에서 열려있을 수 있는 차트 데이터 워크북 정리.

    PowerPoint 내부에 임베디드된 chart의 워크북을 찾아 닫는다.
    """
    if session is None:
        return
    closed = 0
    try:
        for i in range(1, session.slide_count + 1):
            slide = session._prs.Slides(i)
            for shape in slide.Shapes:
                try:
                    if shape.HasChart == -1:  # MSO_TRUE
                        try:
                            shape.Chart.ChartData.Workbook.Close(False)
                            closed += 1
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass
    if closed:
        print(f"[CLEANUP] 차트 데이터 워크북 {closed}개 닫음")


def main():
    print(f"[INFO] 템플릿: {TEMPLATE.name}")
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 현재 슬라이드 수: {session.slide_count}")
    cleanup_open_chart_workbooks(session)

    print("[1/8] 목차 추가...")
    session.add_agenda(
        items=["회사 소개", "1분기 실적", "주요 전략", "시장 분석", "Q&A"],
        page=2,
    )

    print("[2/8] 섹션 구분 추가...")
    session.add_section_divider(
        title="1분기 실적",
        number=1,
        subtitle="2026 Q1 Performance Review",
        page=3,
    )

    print("[3/8] 본문 1단 추가...")
    session.add_single_column(
        title="핵심 성과 요약",
        bullets=[
            "매출 전년 동기 대비 18% 증가, 분기 최대 실적 달성",
            "신규 고객 12개사 확보 (글로벌 5개사 포함)",
            {"text": "EBITDA 마진 14% 유지", "level": 0},
            {"text": "원가 절감 프로젝트 효과 본격 반영", "level": 1},
            "다음 분기 신제품 출시 준비 완료",
        ],
        eyebrow="01 / 경영 보고",
        page=4,
    )

    print("[4/8] 본문 2단 추가...")
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
        page=5,
    )

    print("[5/8] KPI 카드 추가...")
    session.add_kpi_cards(
        title="핵심 지표",
        cards=[
            {"label": "매출 (YoY)", "value": "+18%", "note": "전년 동기 대비"},
            {"label": "신규 고객", "value": "12사", "note": "글로벌 5개사 포함"},
            {"label": "EBITDA 마진", "value": "14%", "note": "원가 개선 효과"},
        ],
        page=6,
    )

    print("[6/8] 차트 추가...")
    session.add_chart(
        title="분기별 매출 추이 (억원)",
        categories=["2025 Q1", "2025 Q2", "2025 Q3", "2025 Q4", "2026 Q1"],
        series=[{"name": "매출", "values": [82, 88, 95, 102, 97]}],
        chart_type="column",
        note="* 2026 Q1은 잠정치 기준",
        page=7,
    )

    print("[7/8] 비교 추가...")
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
        page=8,
    )

    print("[8/8] 프로세스 다이어그램 추가...")
    session.add_process_flow(
        title="PPTMaker 도입 로드맵",
        steps=[
            {"label": "01", "title": "디자인 토큰", "note": "Reference 분석 완료"},
            {"label": "02", "title": "레이아웃 12종", "note": "본문 빌더 구현"},
            {"label": "03", "title": "사내 PoC", "note": "전략기획팀 우선"},
            {"label": "04", "title": "전사 배포", "note": "Q2 내 완료 목표"},
        ],
        page=9,
    )

    print(f"\n[DONE] 슬라이드 {session.slide_count}장 — PowerPoint에서 확인하세요.")
    print("       파일은 저장되지 않았습니다 (사용자 통제).")


if __name__ == "__main__":
    main()
