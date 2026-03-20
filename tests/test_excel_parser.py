"""Excel parserのテスト"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.excel_parser import load_workbook, get_sheet_names, detect_tables

FIXTURE_PATH = os.path.join(os.path.dirname(__file__), "fixtures", "sample.xlsx")


@pytest.fixture(scope="module", autouse=True)
def create_fixture():
    """テスト用Excelファイルを生成"""
    if not os.path.exists(FIXTURE_PATH):
        import subprocess
        subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "create_sample_excel.py")])


def test_load_workbook():
    wb = load_workbook(FIXTURE_PATH)
    assert wb is not None


def test_get_sheet_names():
    wb = load_workbook(FIXTURE_PATH)
    names = get_sheet_names(wb)
    assert "月次売上" in names
    assert "地域別売上" in names
    assert len(names) == 3


def test_detect_tables():
    wb = load_workbook(FIXTURE_PATH)
    ws = wb["月次売上"]
    tables = detect_tables(ws, "月次売上")
    assert len(tables) >= 1
    table = tables[0]
    assert len(table.headers) == 4
    assert len(table.rows) == 6
    assert table.headers[0] == "月"


def test_detect_tables_region():
    wb = load_workbook(FIXTURE_PATH)
    ws = wb["地域別売上"]
    tables = detect_tables(ws, "地域別売上")
    assert len(tables) >= 1
    table = tables[0]
    assert len(table.headers) == 4
    assert len(table.rows) == 6
