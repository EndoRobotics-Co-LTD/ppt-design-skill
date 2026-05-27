"""PPTMaker — EndoRobotics 전사 PPT 빌더.

라이브 모드(PowerPoint COM)와 오프라인 모드(python-pptx)를 모두 지원한다.
2개의 사내 표준 테마(theme1, theme2) 지원 — `themes` 모듈 참조.
"""
from pptmaker import themes
from pptmaker.design_tokens import TOKENS

__all__ = ["TOKENS", "themes"]
__version__ = "0.1.0"
