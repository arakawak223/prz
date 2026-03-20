"""Excelフィールドとテンプレートスロットのマッピング"""

from modules.models import TemplateMap, TemplateTextSlot, TemplateSlotGroup, FieldMapping, ParsedTable


def auto_map(template: TemplateMap, tables: list[ParsedTable]) -> list[FieldMapping]:
    """テンプレートスロットとExcelカラムを自動マッピング。

    戦略: テーブルのヘッダーとスロットのoriginal_textの部分一致を探す。
    """
    mappings = []

    if not tables:
        return [FieldMapping(slot=s) for s in template.all_slots]

    # 全テーブルのヘッダーをフラットに集める
    all_headers = []
    for table in tables:
        all_headers.extend(table.headers)

    for slot in template.all_slots:
        best_match = _find_best_match(slot, all_headers, tables)
        mappings.append(FieldMapping(
            slot=slot,
            excel_column=best_match,
        ))

    return mappings


def _find_best_match(slot: TemplateTextSlot, headers: list[str], tables: list[ParsedTable]) -> str | None:
    """スロットに最も一致するExcelカラムを探す"""
    text = slot.original_text.lower()

    # ヘッダー名がテキストに含まれるか
    for header in headers:
        if header.lower() in text or text in header.lower():
            return header

    # セル値とのマッチング
    for table in tables:
        for row in table.rows:
            for j, val in enumerate(row):
                if val is not None and str(val).strip() and str(val).strip().lower() in text:
                    if j < len(table.headers):
                        return table.headers[j]

    return None


def build_data_rows(tables: list[ParsedTable]) -> list[dict]:
    """テーブルデータを辞書リストに変換"""
    rows = []
    for table in tables:
        for row in table.rows:
            row_dict = {}
            for j, header in enumerate(table.headers):
                if j < len(row):
                    row_dict[header] = row[j]
            rows.append(row_dict)
    return rows


def get_slot_summary(template: TemplateMap) -> list[dict]:
    """テンプレートのスロット情報をサマリーとして返す"""
    summary = []
    for slot in template.all_slots:
        text_preview = slot.original_text[:60] + "..." if len(slot.original_text) > 60 else slot.original_text
        summary.append({
            "スライド": slot.slide_index + 1,
            "役割": {"title": "タイトル", "heading": "見出し", "body": "本文"}.get(slot.role, slot.role),
            "テキスト": text_preview,
            "シェイプ": slot.shape_name,
        })
    return summary
