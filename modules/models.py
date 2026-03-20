from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ParsedTable:
    """Excelから抽出された1つのテーブル"""
    sheet_name: str
    title: Optional[str]
    headers: list[str]
    rows: list[list]
    source_range: str  # e.g. "Sheet1!A1:D10"


@dataclass
class ChartSpec:
    """チャートの仕様"""
    chart_type: str  # "bar" | "line" | "pie" | "stacked_bar" | "column"
    title: str
    category_column: int  # カテゴリ列のインデックス
    value_columns: list[int]  # 値列のインデックス
    reason: str = ""  # このチャート種類を選んだ理由


@dataclass
class SlideContent:
    """1スライドの内容"""
    title: str
    body_text: Optional[str] = None
    table: Optional[ParsedTable] = None
    chart: Optional[ChartSpec] = None
    layout: str = "table"  # "table" | "title" | "bullets" | "blank" | "chart" | "table_chart"
    notes: str = ""


@dataclass
class PresentationSpec:
    """プレゼンテーション全体の仕様"""
    title: str
    audience: str
    purpose: str
    slides: list[SlideContent] = field(default_factory=list)


# --- テンプレート更新機能用モデル ---

@dataclass
class TemplateTextSlot:
    """Gammaスライド内の1つのテキスト領域"""
    slide_index: int
    shape_index: int
    shape_name: str
    role: str  # "title" | "heading" | "body"
    original_text: str
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    font_bold: Optional[bool] = None
    font_color_rgb: Optional[str] = None


@dataclass
class TemplateSlotGroup:
    """1つのケースエントリ（見出し+本文）"""
    slide_index: int
    heading_slot: Optional[TemplateTextSlot] = None
    body_slot: Optional[TemplateTextSlot] = None


@dataclass
class TemplateMap:
    """解析済みテンプレート構造"""
    source_path: str
    slide_count: int
    slide_titles: list[TemplateTextSlot] = field(default_factory=list)
    slot_groups: list[list[TemplateSlotGroup]] = field(default_factory=list)
    all_slots: list[TemplateTextSlot] = field(default_factory=list)


@dataclass
class FieldMapping:
    """Excelフィールド→テンプレートスロットの対応"""
    slot: TemplateTextSlot
    excel_column: Optional[str] = None
    transform: str = "direct"
