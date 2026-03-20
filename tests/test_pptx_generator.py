"""PPTX generatorのテスト"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.models import PresentationSpec, SlideContent, ParsedTable
from modules.pptx_generator import generate_presentation, to_bytes


def _make_spec():
    table = ParsedTable(
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
    return PresentationSpec(
        title="テストプレゼンテーション",
        audience="経営会議",
        purpose="Q3実績報告",
        slides=[
            SlideContent(title="月次売上データ", table=table, layout="table"),
        ],
    )


def test_generate_presentation():
    spec = _make_spec()
    prs = generate_presentation(spec)
    # タイトルスライド + テーブルスライド = 2
    assert len(prs.slides) == 2


def test_to_bytes():
    spec = _make_spec()
    prs = generate_presentation(spec)
    data = to_bytes(prs)
    assert len(data) > 0
    # PPTX is a ZIP file, check magic bytes
    assert data[:2] == b"PK"


def test_large_table_split():
    """大きいテーブルが分割されることを確認"""
    rows = [[f"行{i}", i * 100, i * 10] for i in range(30)]
    table = ParsedTable(
        sheet_name="大量データ",
        title="大きいテーブル",
        headers=["項目", "値1", "値2"],
        rows=rows,
        source_range="大量データ!A1:C31",
    )
    spec = PresentationSpec(
        title="分割テスト",
        audience="テスト",
        purpose="テスト",
        slides=[SlideContent(title="大きいテーブル", table=table, layout="table")],
    )
    prs = generate_presentation(spec)
    # タイトル(1) + テーブル分割(2) = 3スライド
    assert len(prs.slides) == 3
