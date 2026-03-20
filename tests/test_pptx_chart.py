"""PPTXチャート生成のテスト"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.models import PresentationSpec, SlideContent, ParsedTable, ChartSpec
from modules.pptx_generator import generate_presentation, to_bytes


def _make_table():
    return ParsedTable(
        sheet_name="売上",
        title="月次売上データ",
        headers=["月", "売上", "前年比"],
        rows=[
            ["4月", 1200, 105],
            ["5月", 1350, 112],
            ["6月", 980, 89],
        ],
        source_range="売上!A1:C4",
    )


def _make_chart_spec(chart_type="column"):
    return ChartSpec(
        chart_type=chart_type,
        title="月次売上グラフ",
        category_column=0,
        value_columns=[1, 2],
        reason="テスト",
    )


def test_chart_only_slide():
    """グラフのみスライドの生成"""
    spec = PresentationSpec(
        title="チャートテスト",
        audience="テスト",
        purpose="テスト",
        slides=[
            SlideContent(
                title="売上グラフ",
                table=_make_table(),
                chart=_make_chart_spec("column"),
                layout="chart",
            ),
        ],
    )
    prs = generate_presentation(spec)
    assert len(prs.slides) == 2  # タイトル + チャート


def test_table_chart_slide():
    """テーブル+チャート並列スライドの生成"""
    spec = PresentationSpec(
        title="並列テスト",
        audience="テスト",
        purpose="テスト",
        slides=[
            SlideContent(
                title="売上テーブル＋グラフ",
                table=_make_table(),
                chart=_make_chart_spec("line"),
                layout="table_chart",
            ),
        ],
    )
    prs = generate_presentation(spec)
    assert len(prs.slides) == 2  # タイトル + テーブル+チャート
    data = to_bytes(prs)
    assert data[:2] == b"PK"


def test_pie_chart():
    """円グラフの生成"""
    table = ParsedTable(
        sheet_name="地域",
        title="地域別構成比",
        headers=["地域", "構成比"],
        rows=[["東京", 35], ["大阪", 25], ["名古屋", 20], ["その他", 20]],
        source_range="地域!A1:B5",
    )
    chart = ChartSpec(
        chart_type="pie",
        title="地域別構成比",
        category_column=0,
        value_columns=[1],
        reason="テスト",
    )
    spec = PresentationSpec(
        title="円グラフテスト",
        audience="テスト",
        purpose="テスト",
        slides=[SlideContent(title="構成比", table=table, chart=chart, layout="chart")],
    )
    prs = generate_presentation(spec)
    assert len(prs.slides) == 2


def test_all_chart_types():
    """全チャートタイプが生成可能"""
    for chart_type in ["column", "bar", "line", "pie", "stacked_bar"]:
        table = _make_table()
        chart = _make_chart_spec(chart_type)
        if chart_type == "pie":
            chart.value_columns = [1]  # 円グラフは1系列のみ
        spec = PresentationSpec(
            title=f"{chart_type}テスト",
            audience="テスト",
            purpose="テスト",
            slides=[SlideContent(title=f"{chart_type}", table=table, chart=chart, layout="chart")],
        )
        prs = generate_presentation(spec)
        assert len(prs.slides) == 2, f"{chart_type}の生成に失敗"
