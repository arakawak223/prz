"""AI engineのテスト（モック使用）"""

import os
import sys
import json
from unittest.mock import patch, MagicMock

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.models import PresentationSpec, SlideContent, ParsedTable
from modules.ai_engine import _table_to_text, _tables_to_context


def _make_table():
    return ParsedTable(
        sheet_name="売上",
        title="月次売上",
        headers=["月", "売上", "前年比"],
        rows=[["4月", 1200, 105], ["5月", 1350, 112]],
        source_range="売上!A1:C3",
    )


def _make_spec():
    table = _make_table()
    return PresentationSpec(
        title="テスト",
        audience="経営会議",
        purpose="報告",
        slides=[SlideContent(title="月次売上", table=table, layout="table")],
    )


def test_table_to_text():
    table = _make_table()
    text = _table_to_text(table)
    assert "月次売上" in text
    assert "月" in text
    assert "1200" in text


def test_tables_to_context():
    tables = [_make_table()]
    context = _tables_to_context(tables)
    assert "月次売上" in context


@patch.dict(os.environ, {"ANTHROPIC_API_KEY": "test-key"})
@patch("modules.ai_engine.Anthropic")
def test_enrich_presentation_mock(mock_anthropic_cls):
    from modules.ai_engine import enrich_presentation

    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client

    # ストーリー構成のレスポンス
    structure_response = MagicMock()
    structure_response.content = [MagicMock(text=json.dumps([
        {"title": "導入", "type": "intro", "table_index": None, "key_points": ["目的の説明"]},
        {"title": "売上分析", "type": "analysis", "table_index": 0, "key_points": ["売上推移"]},
        {"title": "まとめ", "type": "summary", "table_index": None, "key_points": ["次期施策"]},
    ]))]

    # 要約のレスポンス
    summary_response = MagicMock()
    summary_response.content = [MagicMock(text="売上は前年比で増加傾向にあります。")]

    # ノートのレスポンス
    notes_response = MagicMock()
    notes_response.content = [MagicMock(text=json.dumps([
        "導入のノート",
        "分析のノート",
        "まとめのノート",
    ]))]

    mock_client.messages.create.side_effect = [
        structure_response,
        summary_response,
        notes_response,
    ]

    spec = _make_spec()
    result = enrich_presentation(spec)

    assert len(result.slides) == 3
    assert result.slides[0].title == "導入"
    assert result.slides[1].title == "売上分析"
    assert result.slides[1].body_text == "売上は前年比で増加傾向にあります。"
    assert result.slides[2].title == "まとめ"
    assert result.slides[0].notes == "導入のノート"
