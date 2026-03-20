"""テンプレート解析・更新のテスト"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

GAMMA_PPTX = os.path.join(os.path.dirname(__file__), "..", "sl05ocy2ew7cuqw.pptx")


@pytest.mark.skipif(not os.path.exists(GAMMA_PPTX), reason="Gamma PPTX not found")
class TestTemplateParser:
    def test_parse_template(self):
        from modules.pptx_template_parser import parse_template
        template = parse_template(GAMMA_PPTX)
        assert template.slide_count == 10
        assert len(template.all_slots) > 0
        assert len(template.slide_titles) > 0

    def test_slot_roles(self):
        from modules.pptx_template_parser import parse_template
        template = parse_template(GAMMA_PPTX)
        roles = set(s.role for s in template.all_slots)
        assert "title" in roles or "heading" in roles or "body" in roles

    def test_slot_groups(self):
        from modules.pptx_template_parser import parse_template
        template = parse_template(GAMMA_PPTX)
        total_groups = sum(len(g) for g in template.slot_groups)
        assert total_groups > 0

    def test_update_by_slot_text(self):
        from modules.pptx_template_parser import parse_template
        from modules.pptx_template_updater import update_by_slot_text
        template = parse_template(GAMMA_PPTX)

        # 最初のスロットのテキストを差し替え
        first_slot = template.all_slots[0]
        replacements = [(first_slot, "テスト差し替えテキスト")]

        with open(GAMMA_PPTX, "rb") as f:
            result = update_by_slot_text(f, replacements)

        assert len(result) > 0
        assert result[:2] == b"PK"  # valid PPTX/ZIP

    def test_format_preservation(self):
        """更新後もファイルサイズが大幅に変わらないことで書式保持を確認"""
        from modules.pptx_template_parser import parse_template
        from modules.pptx_template_updater import update_by_slot_text

        original_size = os.path.getsize(GAMMA_PPTX)
        template = parse_template(GAMMA_PPTX)

        first_slot = template.all_slots[0]
        replacements = [(first_slot, "更新テスト")]

        with open(GAMMA_PPTX, "rb") as f:
            result = update_by_slot_text(f, replacements)

        # 書式保持されていればサイズはほぼ同じ（±20%以内）
        assert len(result) > original_size * 0.8
        assert len(result) < original_size * 1.2
