"""AI統合エンジン - Claude APIによるプレゼンテーション強化"""

import json
import os
from anthropic import Anthropic

from modules.models import PresentationSpec, SlideContent, ParsedTable, ChartSpec
from modules.chart_analyzer import analyze_table
import config


def _get_client() -> Anthropic:
    """Anthropicクライアントを取得"""
    api_key = os.environ.get(config.ANTHROPIC_API_KEY_ENV)
    if not api_key:
        raise ValueError(
            f"環境変数 {config.ANTHROPIC_API_KEY_ENV} が設定されていません。\n"
            "export ANTHROPIC_API_KEY=your-api-key を実行してください。"
        )
    return Anthropic(api_key=api_key)


def _table_to_text(table: ParsedTable) -> str:
    """テーブルをテキスト表現に変換"""
    lines = [f"テーブル: {table.title or table.source_range}"]
    lines.append(f"シート: {table.sheet_name}")
    lines.append("| " + " | ".join(table.headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(table.headers)) + " |")
    for row in table.rows:
        lines.append("| " + " | ".join(str(v) if v is not None else "" for v in row) + " |")
    return "\n".join(lines)


def _tables_to_context(tables: list[ParsedTable]) -> str:
    """全テーブルをコンテキスト文字列に変換"""
    return "\n\n".join(_table_to_text(t) for t in tables)


def enrich_presentation(spec: PresentationSpec, progress_callback=None) -> PresentationSpec:
    """Claude APIでプレゼンテーションを強化する。

    以下を自動生成:
    - ストーリー構成（導入→現状分析→課題→解決策→結語）
    - 各テーブルの要約テキストと考察
    - 発表者ノート
    """
    client = _get_client()

    # 全テーブルデータを収集
    tables = [s.table for s in spec.slides if s.table]
    tables_context = _tables_to_context(tables)

    if progress_callback:
        progress_callback("ストーリー構成を生成中...")

    # Step 1: ストーリー構成とスライド計画を生成
    structure = _generate_story_structure(client, spec, tables_context)

    if progress_callback:
        progress_callback("各スライドのコンテンツを生成中...")

    # Step 2: 構成に基づいてスライドを再構築
    new_slides = _build_slides_from_structure(client, spec, tables, structure)

    if progress_callback:
        progress_callback("発表者ノートを生成中...")

    # Step 3: 発表者ノートを追加
    new_slides = _add_speaker_notes(client, spec, new_slides)

    spec.slides = new_slides
    return spec


def _generate_story_structure(client: Anthropic, spec: PresentationSpec, tables_context: str) -> list[dict]:
    """ストーリー構成を生成する"""
    prompt = f"""あなたはプレゼンテーションの構成を設計する専門家です。
以下のデータと条件に基づき、効果的なプレゼンテーションのスライド構成を設計してください。

## プレゼンテーション情報
- タイトル: {spec.title}
- 対象者: {spec.audience}
- 目的: {spec.purpose}

## データ
{tables_context}

## 指示
以下の流れでスライド構成を設計してください:
1. 導入（目的・背景の提示）
2. 現状分析（データに基づく現状の説明）
3. 課題・ポイント（データから読み取れる課題や重要ポイント）
4. 解決策・提案（課題に対するアクション提案）
5. まとめ（結論と次のステップ）

各スライドについて、以下を指定してください:
- どのテーブルのデータを使うか（使わない場合はnull）
- データスライドの表示形式: "table"（表のみ）、"chart"（グラフのみ）、"table_chart"（表＋グラフ並列）
- グラフを使う場合のグラフ種類: "bar"（横棒）、"column"（縦棒）、"line"（折れ線）、"pie"（円）

以下のJSON形式で出力してください（JSON以外のテキストは出力しないでください）:
[
  {{
    "title": "スライドタイトル",
    "type": "intro|analysis|issue|proposal|summary",
    "table_index": null または 0始まりのインデックス,
    "key_points": ["ポイント1", "ポイント2"],
    "display": "table|chart|table_chart",
    "chart_type": "bar|column|line|pie|null"
  }}
]"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    # JSONブロックを抽出
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()

    return json.loads(text)


def _build_slides_from_structure(
    client: Anthropic,
    spec: PresentationSpec,
    tables: list[ParsedTable],
    structure: list[dict],
) -> list[SlideContent]:
    """構成に基づいてスライドコンテンツを生成"""
    slides = []

    for item in structure:
        table_idx = item.get("table_index")
        table = tables[table_idx] if table_idx is not None and table_idx < len(tables) else None
        key_points = item.get("key_points", [])
        display = item.get("display", "table")
        ai_chart_type = item.get("chart_type")

        if table:
            # テーブル付きスライド: 要約テキストを生成
            summary = _generate_table_summary(client, spec, table, key_points)

            # チャート仕様を決定
            chart_spec = None
            if display in ("chart", "table_chart"):
                # まずAI推奨のチャート種類を使用、なければ自動分析
                auto_chart = analyze_table(table)
                if auto_chart:
                    chart_spec = auto_chart
                    if ai_chart_type and ai_chart_type != "null":
                        chart_spec.chart_type = ai_chart_type
                        chart_spec.reason = f"AI推奨: {ai_chart_type}"

            layout = display if chart_spec and display in ("chart", "table_chart") else "table"

            slides.append(SlideContent(
                title=item["title"],
                body_text=summary,
                table=table,
                chart=chart_spec,
                layout=layout,
            ))
        else:
            # テキストのみスライド
            body = "\n".join(f"• {p}" for p in key_points) if key_points else None
            layout = "bullets" if key_points else "blank"
            slides.append(SlideContent(
                title=item["title"],
                body_text=body,
                layout=layout,
            ))

    return slides


def _generate_table_summary(
    client: Anthropic,
    spec: PresentationSpec,
    table: ParsedTable,
    key_points: list[str],
) -> str:
    """テーブルデータの要約と考察を生成"""
    table_text = _table_to_text(table)
    points_text = "\n".join(f"- {p}" for p in key_points) if key_points else "特になし"

    prompt = f"""以下のテーブルデータについて、プレゼンテーション用の簡潔な要約と考察を日本語で生成してください。

## 条件
- 対象者: {spec.audience}
- 目的: {spec.purpose}
- 着目ポイント:
{points_text}

## データ
{table_text}

## 指示
- 2〜3文で要約してください
- 数値の具体的な言及を含めてください
- 「〜が読み取れます」「〜が課題です」等のビジネス文体で
- 箇条書きではなく、文章で書いてください"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}],
    )

    return response.content[0].text.strip()


def _add_speaker_notes(
    client: Anthropic,
    spec: PresentationSpec,
    slides: list[SlideContent],
) -> list[SlideContent]:
    """全スライドに発表者ノートを追加"""
    slides_summary = []
    for i, s in enumerate(slides):
        slides_summary.append(f"スライド{i+1}: {s.title} ({s.layout})")

    slides_text = "\n".join(slides_summary)

    prompt = f"""以下のプレゼンテーションの各スライドに対して、発表者ノート（話すべきポイント）を生成してください。

## プレゼンテーション
- タイトル: {spec.title}
- 対象者: {spec.audience}
- 目的: {spec.purpose}

## スライド一覧
{slides_text}

## 指示
各スライドについて、発表時に話すべきポイントを2〜3文で簡潔に書いてください。

以下のJSON形式で出力してください（JSON以外のテキストは出力しないでください）:
[
  "スライド1のノート",
  "スライド2のノート",
  ...
]"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()

    try:
        notes_list = json.loads(text)
        for i, slide in enumerate(slides):
            if i < len(notes_list):
                slide.notes = notes_list[i]
    except (json.JSONDecodeError, TypeError):
        pass

    return slides


# ===========================================
# テンプレート更新用AI機能
# ===========================================

def improve_text(text: str, context: str = "") -> str:
    """テキストを推敲・ブラッシュアップする"""
    client = _get_client()
    prompt = f"""以下のプレゼンテーション用テキストを推敲・改善してください。

## 元テキスト
{text}

{f"## 補足コンテキスト{chr(10)}{context}" if context else ""}

## 指示
- ビジネスプレゼンテーションにふさわしい文体にしてください
- 冗長な表現を簡潔にしてください
- 重要なポイントを強調してください
- 元テキストの意味や情報を変えないでください
- 改善後のテキストのみを出力してください（説明不要）"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def generate_text_from_data(table_text: str, slot_role: str, original_text: str) -> str:
    """Excelデータに基づいてテンプレートスロット用テキストを生成する"""
    client = _get_client()
    role_desc = {
        "title": "スライドタイトル（短く端的に）",
        "heading": "セクション見出し（簡潔に）",
        "body": "本文（2〜4文で要約）",
    }.get(slot_role, "テキスト")

    prompt = f"""以下のデータに基づいて、プレゼンテーションの{role_desc}を生成してください。

## 参考データ
{table_text}

## 元テキスト（参考）
{original_text[:200]}

## 指示
- {role_desc}として適切な長さと文体で書いてください
- データの具体的な数値やポイントを反映してください
- 生成テキストのみを出力してください（説明不要）"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def translate_text(text: str, target_lang: str) -> str:
    """テキストを翻訳する"""
    client = _get_client()
    lang_name = {"en": "英語", "ja": "日本語", "zh": "中国語", "ko": "韓国語"}.get(target_lang, target_lang)

    prompt = f"""以下のテキストを{lang_name}に翻訳してください。

## 元テキスト
{text}

## 指示
- プレゼンテーション資料として自然な{lang_name}に翻訳してください
- 専門用語はそのまま維持するか、適切に翻訳してください
- 翻訳結果のみを出力してください（説明不要）"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def ai_update_all_slots(slots_text: list[tuple[str, str, str]], purpose: str = "") -> list[str]:
    """全スロットのテキストを一括で改善する。

    Args:
        slots_text: [(role, original_text, slot_id), ...] のリスト
        purpose: 改善の目的
    Returns:
        改善後テキストのリスト
    """
    client = _get_client()

    slots_desc = []
    for i, (role, text, _) in enumerate(slots_text):
        role_label = {"title": "タイトル", "heading": "見出し", "body": "本文"}.get(role, role)
        slots_desc.append(f"[{i}] ({role_label}) {text[:150]}")

    slots_joined = "\n\n".join(slots_desc)

    prompt = f"""以下のプレゼンテーションの各テキストを推敲・改善してください。

{f"## 目的{chr(10)}{purpose}" if purpose else ""}

## テキスト一覧
{slots_joined}

## 指示
- 各テキストをビジネスプレゼンにふさわしい文体に改善してください
- タイトルは短く端的に、本文は簡潔かつ具体的にしてください
- 元の意味や情報を変えないでください

以下のJSON形式で出力してください（JSON以外のテキストは出力しないでください）:
[
  "改善後テキスト0",
  "改善後テキスト1",
  ...
]"""

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()

    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return []
