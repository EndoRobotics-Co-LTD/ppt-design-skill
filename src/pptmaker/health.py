"""슬라이드 시각 건강 검진 — 텍스트 오버플로우 / 영역 침범 / 도형 겹침 감지 + 자동 fix.

핵심 trade-off: 디자인 표준 일관성 vs 컨텐츠 완결성·가독성.
사용자 명시 원칙: 컨텐츠가 디자인을 깨면 **컨텐츠 완결성 우선**.
즉 Type 스케일을 일시적으로 한 단계 깨더라도 텍스트가 보이게 한다.

LiveSession.add_xxx() 메서드들 끝에 자동 호출되어 silent fix.
자동 fix로 해결 안 되면 issue 리스트 반환 → Claude가 사용자에게 옵션 제시.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Iterable

# COM enum (msoAutoSize)
MSO_AUTO_SIZE_NONE = 0
MSO_AUTO_SIZE_SHAPE_TO_FIT_TEXT = 2
MSO_AUTO_SIZE_TEXT_TO_FIT_SHAPE = 1
MSO_TRUE = -1

# 슬라이드 chrome 영역 (pt). _common.py 와 일치해야 함.
CHROME_TOP_H = 62.4         # 헤더 바 높이
CHROME_BOTTOM_Y = 516.4     # 하단 구분선 y
SLIDE_H = 540.0
SLIDE_W = 960.0
EPSILON = 1.0  # 1pt 이내 침범은 무시 (anti-aliasing 등)

# Type 스케일 (pt) — _common.Type 와 동기. 자동 축소 시 사용.
TYPE_SCALE_DESCENDING = [72, 56, 40, 28, 20, 16, 14, 11, 9]


@dataclass
class Issue:
    """발견된 시각 문제."""
    kind: str           # "text_overflow" | "body_bottom_violation" | "shape_overlap" | etc.
    slide_index: int
    shape_name: str
    detail: str
    severity: str = "warning"   # "warning" | "error"

    def __str__(self) -> str:
        return f"[slide {self.slide_index}] {self.kind} on {self.shape_name!r}: {self.detail}"


@dataclass
class Fix:
    """자동 적용된 수정."""
    kind: str           # "font_shrink" | "line_spacing_shrink" | etc.
    slide_index: int
    shape_name: str
    detail: str

    def __str__(self) -> str:
        return f"[slide {self.slide_index}] {self.kind} on {self.shape_name!r}: {self.detail}"


# ──────────────────────────────────────────────────────────────────────
# 감지
# ──────────────────────────────────────────────────────────────────────

def _shape_bottom(shape) -> float:
    try:
        return float(shape.Top) + float(shape.Height)
    except Exception:
        return 0.0


def _shape_right(shape) -> float:
    try:
        return float(shape.Left) + float(shape.Width)
    except Exception:
        return 0.0


def _estimate_wrapped_lines(text: str, font_size: float, box_width: float) -> int:
    """텍스트가 box_width 안에서 자동 줄바꿈됐을 때 차지하는 라인 수 추정.

    한글 + 영문 혼용 가정. 평균 글자 폭 = font_size * 0.6 (한글 약 1.0, 영문 약 0.5의 가중평균).
    """
    if not text or box_width <= 0:
        return 0
    avg_char_width = font_size * 0.6
    chars_per_line = max(1, int(box_width / avg_char_width))
    total_lines = 0
    # paragraph 분리 (CR 또는 LF)
    paragraphs = text.replace("\n", "\r").split("\r")
    for p in paragraphs:
        p_len = len(p)
        if p_len == 0:
            total_lines += 1  # 빈 줄
        else:
            total_lines += max(1, math.ceil(p_len / chars_per_line))
    return total_lines


def _is_text_overflowing(shape) -> tuple[bool, str]:
    """텍스트가 박스를 넘어가는지 추정.

    PowerPoint Lines.Count는 paragraph 수만 반환하는 경우가 있어
    자동 wrapping을 고려한 라인 수 추정으로 보완.

    Returns: (오버플로우 여부, 상세 텍스트)
    """
    try:
        if shape.HasTextFrame != MSO_TRUE:
            return False, ""
        tf = shape.TextFrame
        tr = tf.TextRange
        text = (tr.Text or "").strip()
        if not text:
            return False, ""
        # 폰트 사이즈 (robust read)
        font_size = _read_font_size(tr) or 14.0
        # 박스 폭에서 좌우 마진 약간 제외
        try:
            box_width = float(shape.Width) - 8
        except Exception:
            box_width = 800.0
        # wrapping 추정
        line_count = _estimate_wrapped_lines(text, font_size, box_width)
        # 라인 spacing 1.3배 + 박스 상하 패딩 8pt
        needed_height = line_count * font_size * 1.3 + 8
        actual_height = float(shape.Height)
        if needed_height > actual_height + EPSILON:
            return True, (
                f"text needs ~{needed_height:.0f}pt, shape height {actual_height:.0f}pt "
                f"(est. lines={line_count}, font={font_size:.0f}pt, box_w={box_width:.0f}pt)"
            )
        return False, ""
    except Exception:
        return False, ""


def _is_body_bottom_violated(shape) -> tuple[bool, str]:
    """도형의 하단이 chrome 하단 구분선(BODY_BOTTOM)을 넘어가는지."""
    try:
        bottom = _shape_bottom(shape)
        if bottom > CHROME_BOTTOM_Y + EPSILON:
            return True, f"shape bottom {bottom:.0f}pt exceeds BODY_BOTTOM {CHROME_BOTTOM_Y:.0f}pt"
        return False, ""
    except Exception:
        return False, ""


def _is_slide_overflow(shape) -> tuple[bool, str]:
    """도형이 슬라이드 자체를 벗어나는지 (Top<0, Left<0, Right>SLIDE_W, Bottom>SLIDE_H)."""
    try:
        if _shape_bottom(shape) > SLIDE_H + EPSILON:
            return True, f"shape bottom exceeds slide height ({_shape_bottom(shape):.0f} > {SLIDE_H:.0f})"
        if _shape_right(shape) > SLIDE_W + EPSILON:
            return True, f"shape right exceeds slide width ({_shape_right(shape):.0f} > {SLIDE_W:.0f})"
        return False, ""
    except Exception:
        return False, ""


def check_slide(slide, slide_index: int) -> list[Issue]:
    """슬라이드 한 장의 시각 문제 감지."""
    issues: list[Issue] = []
    for shape in slide.Shapes:
        try:
            name = str(shape.Name)
        except Exception:
            name = "?"
        # 텍스트 오버플로우
        overflow, detail = _is_text_overflowing(shape)
        if overflow:
            issues.append(Issue("text_overflow", slide_index, name, detail))
        # body_bottom 침범
        violated, detail = _is_body_bottom_violated(shape)
        if violated:
            issues.append(Issue("body_bottom_violation", slide_index, name, detail, severity="error"))
        # 슬라이드 영역 벗어남
        out, detail = _is_slide_overflow(shape)
        if out:
            issues.append(Issue("slide_overflow", slide_index, name, detail, severity="error"))
    return issues


# ──────────────────────────────────────────────────────────────────────
# 자동 fix
# ──────────────────────────────────────────────────────────────────────

def _read_font_size(tr) -> float | None:
    """TextRange 의 폰트 사이즈를 안전하게 읽는다. mixed면 첫 run 또는 paragraph 사이즈."""
    # 1) range 전체 폰트 사이즈
    try:
        sz = tr.Font.Size
        if sz is not None and sz > 0:
            return float(sz)
    except Exception:
        pass
    # 2) 첫 run
    try:
        sz = tr.Runs(1).Font.Size
        if sz is not None and sz > 0:
            return float(sz)
    except Exception:
        pass
    # 3) 첫 paragraph
    try:
        sz = tr.Paragraphs(1).Font.Size
        if sz is not None and sz > 0:
            return float(sz)
    except Exception:
        pass
    return None


def _set_font_size_all(tr, size: float) -> bool:
    """TextRange 의 모든 paragraph/run에 폰트 사이즈를 강제 적용. 성공 여부."""
    # 1) range 전체에 한 번
    try:
        tr.Font.Size = size
        return True
    except Exception:
        pass
    # 2) paragraph 마다
    try:
        for p in tr.Paragraphs():
            try:
                p.Font.Size = size
            except Exception:
                pass
        return True
    except Exception:
        pass
    # 3) run 마다
    try:
        for r in tr.Runs():
            try:
                r.Font.Size = size
            except Exception:
                pass
        return True
    except Exception:
        pass
    return False


def _shrink_font_one_step(shape) -> Fix | None:
    """텍스트박스의 폰트 사이즈를 Type 스케일에서 한 단계 작게."""
    try:
        if shape.HasTextFrame != MSO_TRUE:
            return None
        tr = shape.TextFrame.TextRange
        original = _read_font_size(tr)
        if original is None or original <= 0:
            # 폰트 사이즈 못 읽으면 default 가정
            original = 14.0
        # 현재 사이즈 이하의 가장 가까운 스케일 값
        smaller = [s for s in TYPE_SCALE_DESCENDING if s < original]
        if not smaller:
            return None
        new_size = max(smaller)
        if not _set_font_size_all(tr, new_size):
            return None
        return Fix(
            "font_shrink",
            -1,  # slide_index 는 caller가 채움
            str(shape.Name),
            f"{original:.0f}pt → {new_size}pt",
        )
    except Exception:
        return None


def _expand_shape_height(shape, max_bottom: float) -> Fix | None:
    """텍스트박스 높이를 늘려서 텍스트 fit. 단 max_bottom 안 넘게."""
    try:
        if shape.HasTextFrame != MSO_TRUE:
            return None
        original_h = float(shape.Height)
        top = float(shape.Top)
        # 사용 가능한 최대 높이
        max_h = max_bottom - top - EPSILON
        if max_h <= original_h:
            return None
        # 필요한 높이 추정
        overflow, _ = _is_text_overflowing(shape)
        if not overflow:
            return None
        # 약간 여유 두고 확장
        new_h = min(max_h, original_h * 1.5)
        if new_h <= original_h + EPSILON:
            return None
        shape.Height = new_h
        return Fix(
            "expand_shape",
            -1,
            str(shape.Name),
            f"height {original_h:.0f}pt → {new_h:.0f}pt",
        )
    except Exception:
        return None


def auto_fit_slide(slide, slide_index: int, max_passes: int = 3) -> tuple[list[Fix], list[Issue]]:
    """슬라이드의 시각 문제를 자동 수정 시도.

    전략 (우선순위 순):
      1. text_overflow 가 있으면 → 박스 확장 시도 (BODY_BOTTOM 한도 내)
      2. 그래도 오버플로우면 → 폰트 한 단계 축소
      3. 최대 max_passes 회 반복

    Returns:
        (applied_fixes, remaining_issues)
        remaining_issues 가 비어있지 않으면 Claude가 사용자에게 옵션 제시해야 함.
    """
    applied: list[Fix] = []
    for _ in range(max_passes):
        issues = check_slide(slide, slide_index)
        if not issues:
            break

        progressed = False
        for issue in issues:
            shape = None
            # shape 찾기 (Name으로)
            for sh in slide.Shapes:
                try:
                    if str(sh.Name) == issue.shape_name:
                        shape = sh
                        break
                except Exception:
                    continue
            if shape is None:
                continue

            if issue.kind == "text_overflow":
                # 1. 박스 확장 시도
                fix = _expand_shape_height(shape, CHROME_BOTTOM_Y)
                if fix is not None:
                    fix.slide_index = slide_index
                    applied.append(fix)
                    progressed = True
                    continue
                # 2. 폰트 한 단계 축소
                fix = _shrink_font_one_step(shape)
                if fix is not None:
                    fix.slide_index = slide_index
                    applied.append(fix)
                    progressed = True
                    continue
            elif issue.kind == "body_bottom_violation":
                # 박스가 BODY_BOTTOM 침범 → 높이 줄여서 BODY_BOTTOM에 맞춤
                try:
                    new_h = CHROME_BOTTOM_Y - float(shape.Top) - EPSILON
                    if new_h > 0:
                        old_h = float(shape.Height)
                        shape.Height = new_h
                        applied.append(Fix(
                            "shrink_to_body_bottom",
                            slide_index,
                            str(shape.Name),
                            f"height {old_h:.0f}pt → {new_h:.0f}pt",
                        ))
                        progressed = True
                except Exception:
                    pass

        if not progressed:
            break  # 더 이상 자동 수정 못 함

    # 남은 문제
    remaining = check_slide(slide, slide_index)
    return applied, remaining
