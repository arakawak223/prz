"""チャート分析モジュールのテスト"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.models import ParsedTable
from modules.chart_analyzer import analyze_table


def test_time_series_detection():
    """時系列データは折れ線グラフを推奨"""
    table = ParsedTable(
        sheet_name="売上",
        title="月次売上",
        headers=["月", "売上（万円）", "前年比（%）"],
        rows=[
            ["4月", 1200, 105],
            ["5月", 1350, 112],
            ["6月", 980, 89],
        ],
        source_range="売上!A1:C4",
    )
    result = analyze_table(table)
    assert result is not None
    assert result.chart_type == "line"
    assert result.category_column == 0
    assert 1 in result.value_columns
    assert 2 in result.value_columns


def test_proportion_data_pie():
    """構成比データは円グラフを推奨"""
    table = ParsedTable(
        sheet_name="地域",
        title="地域別構成比",
        headers=["地域", "構成比（%）"],
        rows=[
            ["東京", 35],
            ["大阪", 25],
            ["名古屋", 20],
            ["その他", 20],
        ],
        source_range="地域!A1:B5",
    )
    result = analyze_table(table)
    assert result is not None
    assert result.chart_type == "pie"


def test_comparison_data_column():
    """比較データは棒グラフを推奨"""
    table = ParsedTable(
        sheet_name="製品",
        title="製品別売上",
        headers=["製品", "売上"],
        rows=[
            ["A", 500],
            ["B", 300],
            ["C", 200],
        ],
        source_range="製品!A1:B4",
    )
    result = analyze_table(table)
    assert result is not None
    assert result.chart_type == "column"


def test_no_numeric_columns():
    """数値列がない場合はNoneを返す"""
    table = ParsedTable(
        sheet_name="テスト",
        title="テキストのみ",
        headers=["名前", "部署"],
        rows=[
            ["田中", "営業"],
            ["鈴木", "開発"],
        ],
        source_range="テスト!A1:B3",
    )
    result = analyze_table(table)
    assert result is None


def test_many_categories_bar():
    """カテゴリが多い場合は横棒グラフを推奨"""
    table = ParsedTable(
        sheet_name="データ",
        title="部署別",
        headers=["部署", "人数"],
        rows=[[f"部署{i}", i * 10] for i in range(8)],
        source_range="データ!A1:B9",
    )
    result = analyze_table(table)
    assert result is not None
    assert result.chart_type == "bar"
