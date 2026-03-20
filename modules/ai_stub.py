"""AI統合スタブ - Phase 2でClaude APIに置き換え"""

from modules.models import PresentationSpec, SlideContent


def enrich_presentation(spec: PresentationSpec) -> PresentationSpec:
    """Phase 2: Claudeでナラティブ追加、スライド再構成、要約生成。
    MVP: そのまま返す。"""
    return spec


def generate_speaker_notes(slide: SlideContent, audience: str, purpose: str) -> str:
    """Phase 2: スライドごとの発表者ノートを生成。
    MVP: 空文字を返す。"""
    return ""


def suggest_slide_order(tables: list, purpose: str) -> list[int]:
    """Phase 2: AIが最適なスライド順序を提案。
    MVP: 元の順序を返す。"""
    return list(range(len(tables)))
