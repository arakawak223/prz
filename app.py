"""prz - プレゼン自動生成アプリ"""

import streamlit as st
import pandas as pd

import os
from modules.excel_parser import load_workbook, get_sheet_names, detect_tables
from modules.pptx_generator import generate_presentation, to_bytes
from modules.chart_analyzer import analyze_table
from modules.pptx_template_parser import parse_template
from modules.pptx_template_updater import update_by_slot_text
from modules.template_mapper import get_slot_summary, build_data_rows
from modules.models import PresentationSpec, SlideContent
import config


def _make_unique_headers(headers: list[str]) -> list[str]:
    """重複ヘッダー名に連番を付与"""
    result = []
    seen = {}
    for h in headers:
        if h in seen:
            seen[h] += 1
            result.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            result.append(h)
    return result


def _table_to_df(table) -> pd.DataFrame:
    """ParsedTableをDataFrameに変換（重複ヘッダー対応）"""
    return pd.DataFrame(table.rows, columns=_make_unique_headers(table.headers))


CHART_TYPE_LABELS = {
    "none": "なし（テーブルのみ）",
    "column": "棒グラフ（縦）",
    "bar": "棒グラフ（横）",
    "line": "折れ線グラフ",
    "pie": "円グラフ",
}

LAYOUT_LABELS = {
    "table": "テーブルのみ",
    "chart": "グラフのみ",
    "table_chart": "テーブル+グラフ並列",
}

st.set_page_config(
    page_title=config.APP_TITLE,
    page_icon="📊",
    layout="wide",
)

st.title(config.APP_TITLE)

# --- AI機能の利用可否チェック ---
ai_available = bool(os.environ.get(config.ANTHROPIC_API_KEY_ENV))

# --- サイドバー設定 ---
with st.sidebar:
    st.header("設定")
    use_ai = st.toggle(
        "AI機能を使用（Claude API）",
        value=ai_available,
        disabled=not ai_available,
    )
    if not ai_available:
        st.caption(
            f"AI機能を有効にするには環境変数 `{config.ANTHROPIC_API_KEY_ENV}` を設定してください。"
        )
    if use_ai:
        st.success("AI機能: 有効")
    else:
        st.info("AI機能: 無効（テーブル転記のみ）")

    st.divider()
    st.header("グラフ設定")
    auto_chart = st.toggle("グラフ自動生成", value=True, key="auto_chart")
    if auto_chart:
        st.caption("数値データを含むテーブルにグラフを自動追加します。")
    default_layout = st.selectbox(
        "デフォルト表示形式",
        options=list(LAYOUT_LABELS.keys()),
        format_func=lambda x: LAYOUT_LABELS[x],
        index=2,
        key="default_layout",
    )

# --- セッション状態の初期化 ---
for key, default in [
    ("parsed_tables", []), ("selected_tables", []), ("spec", None),
    ("pptx_bytes", None), ("chart_settings", {}),
    ("template_map", None), ("updated_pptx_bytes", None), ("ai_suggestions", {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ===========================================
# タブ切替
# ===========================================
tab1, tab2 = st.tabs(["新規プレゼン生成", "テンプレート更新（Gamma等）"])

# ===========================================
# タブ1: 新規プレゼン生成（既存機能）
# ===========================================
with tab1:
    st.header("Step 1: Excelファイルのアップロード")

    uploaded_file = st.file_uploader(
        "ファイルを選択してください（Excel / PowerPoint）",
        type=["xlsx", "xls", "pptx"],
        key="excel_upload",
    )

    if uploaded_file is not None and uploaded_file.name.endswith(".pptx"):
        st.info("PowerPointファイルが検出されました。「テンプレート更新（Gamma等）」タブに切り替えてアップロードしてください。")

    if uploaded_file is not None and not uploaded_file.name.endswith(".pptx"):
        try:
            wb = load_workbook(uploaded_file)
            sheet_names = get_sheet_names(wb)

            st.success(f"読み込み完了: {len(sheet_names)} シート検出")

            selected_sheets = st.multiselect(
                "対象シートを選択", sheet_names, default=sheet_names,
            )

            all_tables = []
            for sheet_name in selected_sheets:
                ws = wb[sheet_name]
                tables = detect_tables(ws, sheet_name)
                all_tables.extend(tables)

            st.session_state.parsed_tables = all_tables

            if all_tables:
                st.info(f"{len(all_tables)} 個のテーブルを検出しました")
                for i, table in enumerate(all_tables):
                    with st.expander(f"テーブル {i+1}: {table.source_range} ({len(table.headers)}列 x {len(table.rows)}行)"):
                        st.dataframe(_table_to_df(table), use_container_width=True)
            else:
                st.warning("テーブルが検出されませんでした。")

        except Exception as e:
            st.error(f"ファイル読み込みエラー: {e}")

    # Step 2: コンテキスト設定
    if st.session_state.parsed_tables:
        st.divider()
        st.header("Step 2: コンテキスト設定")

        col1, col2 = st.columns(2)
        with col1:
            pres_title = st.text_input("プレゼンテーションタイトル", value="プレゼンテーション資料", key="pres_title")
            audience = st.text_input("対象者（誰に向けた資料か）", placeholder="例: 経営会議メンバー", key="audience")
        with col2:
            purpose = st.text_area("目的（何のための資料か）", placeholder="例: Q3売上実績の報告", key="purpose", height=120)

        st.subheader("スライドに含めるテーブル")
        selected_indices = []
        chart_settings = {}

        for i, table in enumerate(st.session_state.parsed_tables):
            col_check, col_chart = st.columns([3, 2])
            with col_check:
                selected = st.checkbox(
                    f"テーブル {i+1}: {table.source_range} ({len(table.headers)}列 x {len(table.rows)}行)",
                    value=True, key=f"table_select_{i}",
                )
            if selected:
                selected_indices.append(i)
                with col_chart:
                    if auto_chart:
                        auto_analysis = analyze_table(table)
                        if auto_analysis:
                            recommended = auto_analysis.chart_type
                            chart_type = st.selectbox(
                                f"グラフ種類",
                                options=list(CHART_TYPE_LABELS.keys()),
                                format_func=lambda x: CHART_TYPE_LABELS[x],
                                index=list(CHART_TYPE_LABELS.keys()).index(recommended) if recommended in CHART_TYPE_LABELS else 0,
                                key=f"chart_type_{i}",
                            )
                            if chart_type != "none":
                                chart_settings[i] = chart_type
                                st.caption(f"推奨理由: {auto_analysis.reason}")
                        else:
                            st.caption("数値列なし - グラフ不可")

        st.session_state.selected_tables = selected_indices
        st.session_state.chart_settings = chart_settings

    # Step 3: プレビューと生成
    if st.session_state.parsed_tables and st.session_state.selected_tables:
        st.divider()
        st.header("Step 3: プレビューと生成")

        st.subheader("スライド構成")
        slides_preview = [{"スライド": 1, "種類": "タイトル", "内容": pres_title, "グラフ": "-"}]
        chart_settings = st.session_state.get("chart_settings", {})
        for idx, table_idx in enumerate(st.session_state.selected_tables):
            table = st.session_state.parsed_tables[table_idx]
            chart_info = CHART_TYPE_LABELS.get(chart_settings.get(table_idx, "none"), "なし")
            layout_info = LAYOUT_LABELS.get(default_layout, "テーブルのみ") if table_idx in chart_settings else "テーブルのみ"
            slides_preview.append({
                "スライド": idx + 2, "種類": layout_info,
                "内容": f"{table.source_range} ({len(table.headers)}列 x {len(table.rows)}行)",
                "グラフ": chart_info,
            })
        st.table(pd.DataFrame(slides_preview))

        button_label = "AIでプレゼンを生成" if use_ai else "PowerPointを生成"
        if st.button(button_label, type="primary", use_container_width=True):
            slides = []
            chart_settings = st.session_state.get("chart_settings", {})
            for table_idx in st.session_state.selected_tables:
                table = st.session_state.parsed_tables[table_idx]
                slide_title = table.title or table.source_range
                chart_spec = None
                slide_layout = "table"
                if table_idx in chart_settings:
                    auto_analysis = analyze_table(table)
                    if auto_analysis:
                        auto_analysis.chart_type = chart_settings[table_idx]
                        chart_spec = auto_analysis
                        slide_layout = default_layout

                slides.append(SlideContent(title=slide_title, table=table, chart=chart_spec, layout=slide_layout))

            spec = PresentationSpec(
                title=pres_title, audience=audience or "未設定",
                purpose=purpose or "未設定", slides=slides,
            )

            if use_ai:
                from modules.ai_engine import enrich_presentation
                progress_bar = st.progress(0, text="AI処理を開始...")
                step_count = [0]
                def on_progress(msg):
                    step_count[0] += 1
                    progress_bar.progress(min(step_count[0] * 30, 90), text=msg)
                try:
                    spec = enrich_presentation(spec, progress_callback=on_progress)
                    progress_bar.progress(100, text="AI処理完了")
                except Exception as e:
                    st.error(f"AI処理エラー: {e}")
                    st.stop()
            else:
                from modules.ai_stub import enrich_presentation
                with st.spinner("生成中..."):
                    spec = enrich_presentation(spec)

            st.session_state.spec = spec
            with st.spinner("PowerPointファイルを生成中..."):
                prs = generate_presentation(spec)
                st.session_state.pptx_bytes = to_bytes(prs)

            st.success(f"生成完了！（{len(spec.slides) + 1} スライド）")

            if spec.slides:
                st.subheader("生成コンテンツプレビュー")
                for i, slide in enumerate(spec.slides):
                    chart_label = ""
                    if slide.chart:
                        chart_label = f" | {CHART_TYPE_LABELS.get(slide.chart.chart_type, slide.chart.chart_type)}"
                    with st.expander(f"スライド {i+1}: {slide.title} [{LAYOUT_LABELS.get(slide.layout, slide.layout)}{chart_label}]", expanded=(i == 0)):
                        if slide.body_text:
                            st.markdown(f"**要約:** {slide.body_text}")
                        if slide.chart:
                            st.markdown(f"**グラフ:** {CHART_TYPE_LABELS.get(slide.chart.chart_type, slide.chart.chart_type)} - {slide.chart.reason}")
                        if slide.table:
                            st.dataframe(_table_to_df(slide.table), use_container_width=True)
                        if slide.notes:
                            st.markdown(f"**発表者ノート:** {slide.notes}")

    # Step 4: ダウンロード
    if st.session_state.pptx_bytes:
        st.divider()
        st.header("Step 4: ダウンロード")
        file_size_kb = len(st.session_state.pptx_bytes) / 1024
        st.info(f"ファイルサイズ: {file_size_kb:.1f} KB")
        st.download_button(
            label="PowerPointファイルをダウンロード (.pptx)",
            data=st.session_state.pptx_bytes,
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", use_container_width=True,
        )


# ===========================================
# タブ2: テンプレート更新（Gamma等）
# ===========================================
with tab2:
    st.markdown("Gammaなどで作成した美しいPPTXのデザインを保持したまま、テキスト内容を編集・更新します。")

    # Step 1: テンプレートPPTXアップロード
    st.header("Step 1: PPTXファイルのアップロード")
    template_file = st.file_uploader(
        "編集したいPPTXファイルを選択",
        type=["pptx"],
        key="template_upload",
    )

    if template_file is not None:
        try:
            template_map = parse_template(template_file)
            st.session_state.template_map = template_map

            st.success(f"解析完了: {template_map.slide_count} スライド、{len(template_map.all_slots)} テキスト領域を検出")

        except Exception as e:
            st.error(f"テンプレート解析エラー: {e}")

    # Step 2: AI機能 & テキスト編集
    if st.session_state.template_map and template_file is not None:
        st.divider()
        st.header("Step 2: テキスト編集")

        template_map = st.session_state.template_map

        # --- AI一括処理ボタン ---
        if use_ai:
            st.subheader("AI一括処理")
            ai_col1, ai_col2, ai_col3 = st.columns(3)

            with ai_col1:
                if st.button("AIで全テキストを推敲", use_container_width=True, key="ai_improve_all"):
                    from modules.ai_engine import ai_update_all_slots
                    with st.spinner("AIがテキストを推敲中..."):
                        slots_data = [(s.role, s.original_text, f"{s.slide_index}_{s.shape_index}") for s in template_map.all_slots]
                        improved = ai_update_all_slots(slots_data)
                        if improved:
                            st.session_state.ai_suggestions = {}
                            for i, slot in enumerate(template_map.all_slots):
                                if i < len(improved):
                                    st.session_state.ai_suggestions[f"{slot.slide_index}_{slot.shape_index}"] = improved[i]
                            st.success(f"{len(improved)} 箇所の改善案を生成しました。下の編集欄に反映されています。")
                            st.rerun()
                        else:
                            st.error("AI推敲に失敗しました。")

            with ai_col2:
                target_lang = st.selectbox(
                    "翻訳先",
                    options=["en", "zh", "ko", "ja"],
                    format_func=lambda x: {"en": "英語", "zh": "中国語", "ko": "韓国語", "ja": "日本語"}.get(x, x),
                    key="translate_lang",
                    label_visibility="collapsed",
                )
                if st.button("AIで全テキストを翻訳", use_container_width=True, key="ai_translate_all"):
                    from modules.ai_engine import translate_text
                    with st.spinner(f"AIが翻訳中..."):
                        st.session_state.ai_suggestions = {}
                        for slot in template_map.all_slots:
                            try:
                                translated = translate_text(slot.original_text, target_lang)
                                st.session_state.ai_suggestions[f"{slot.slide_index}_{slot.shape_index}"] = translated
                            except Exception as e:
                                st.warning(f"スライド{slot.slide_index+1}の翻訳に失敗: {e}")
                        st.success("翻訳完了！下の編集欄に反映されています。")
                        st.rerun()

            with ai_col3:
                excel_for_ai = st.file_uploader(
                    "Excelデータ（任意）",
                    type=["xlsx", "xls"],
                    key="ai_excel_upload",
                )
                if excel_for_ai and st.button("AIでデータ反映", use_container_width=True, key="ai_data_update"):
                    from modules.ai_engine import generate_text_from_data
                    with st.spinner("AIがデータを解析中..."):
                        wb_ai = load_workbook(excel_for_ai)
                        ai_tables = []
                        for sn in get_sheet_names(wb_ai):
                            ai_tables.extend(detect_tables(wb_ai[sn], sn))

                        if ai_tables:
                            from modules.ai_engine import _tables_to_context
                            table_context = _tables_to_context(ai_tables)
                            st.session_state.ai_suggestions = {}
                            for slot in template_map.all_slots:
                                try:
                                    new_text = generate_text_from_data(table_context, slot.role, slot.original_text)
                                    st.session_state.ai_suggestions[f"{slot.slide_index}_{slot.shape_index}"] = new_text
                                except Exception as e:
                                    st.warning(f"スライド{slot.slide_index+1}のテキスト生成に失敗: {e}")
                            st.success("データ反映完了！下の編集欄に反映されています。")
                            st.rerun()
                        else:
                            st.warning("Excelからテーブルが検出されませんでした。")

            st.divider()

        # --- スライドごとのテキスト編集 ---
        st.caption("変更したいテキスト領域だけ編集してください。空欄のまま＝変更なし。")

        ai_suggestions = st.session_state.get("ai_suggestions", {})
        replacements = []

        for slide_title in template_map.slide_titles:
            slide_idx = slide_title.slide_index
            slide_slots = [s for s in template_map.all_slots if s.slide_index == slide_idx]
            text_preview = slide_title.original_text[:50]

            with st.expander(f"スライド {slide_idx + 1}: {text_preview}", expanded=False):
                for slot in slide_slots:
                    role_label = {"title": "タイトル", "heading": "見出し", "body": "本文"}.get(slot.role, slot.role)
                    slot_key = f"{slot.slide_index}_{slot.shape_index}"

                    st.markdown(f"**[{role_label}]** 現在のテキスト:")
                    st.caption(slot.original_text[:200] + ("..." if len(slot.original_text) > 200 else ""))

                    # AI提案があればデフォルト値に設定
                    default_val = ai_suggestions.get(slot_key, "")

                    col_edit, col_ai = st.columns([4, 1])
                    with col_edit:
                        new_text = st.text_area(
                            f"新テキスト（S{slide_idx+1} {role_label}）",
                            value=default_val,
                            placeholder="空欄＝変更しない",
                            key=f"replace_{slot_key}",
                            height=68,
                            label_visibility="collapsed",
                        )
                    with col_ai:
                        if use_ai:
                            if st.button("AI推敲", key=f"ai_single_{slot_key}"):
                                from modules.ai_engine import improve_text
                                with st.spinner("推敲中..."):
                                    source = new_text.strip() if new_text.strip() else slot.original_text
                                    improved = improve_text(source)
                                    st.session_state.ai_suggestions[slot_key] = improved
                                    st.rerun()

                    if new_text.strip():
                        replacements.append((slot, new_text.strip()))
                    st.divider()

        # Step 3: 生成とダウンロード
        st.divider()
        st.header("Step 3: 更新とダウンロード")

        if replacements:
            st.info(f"{len(replacements)} 箇所のテキストを差し替えます")
            with st.expander("差し替え内容の確認"):
                for slot, new_text in replacements:
                    role_label = {"title": "タイトル", "heading": "見出し", "body": "本文"}.get(slot.role, slot.role)
                    old_preview = slot.original_text[:50] + "..." if len(slot.original_text) > 50 else slot.original_text
                    new_preview = new_text[:50] + "..." if len(new_text) > 50 else new_text
                    st.markdown(f"**S{slot.slide_index+1} [{role_label}]:**")
                    st.markdown(f"  {old_preview}")
                    st.markdown(f"  → **{new_preview}**")

        if st.button("テンプレートを更新", type="primary", use_container_width=True, key="update_template"):
            if not replacements:
                st.warning("変更するテキストを入力してください。")
            else:
                with st.spinner("デザインを保持してテキストを更新中..."):
                    template_file.seek(0)
                    updated_bytes = update_by_slot_text(template_file, replacements)
                    st.session_state.updated_pptx_bytes = updated_bytes
                st.success("更新完了！デザインはそのまま、テキストだけ差し替えました。")

        if st.session_state.updated_pptx_bytes:
            file_size_kb = len(st.session_state.updated_pptx_bytes) / 1024
            st.info(f"ファイルサイズ: {file_size_kb:.1f} KB")
            st.download_button(
                label="更新済みPPTXをダウンロード",
                data=st.session_state.updated_pptx_bytes,
                file_name="updated_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary", use_container_width=True,
                key="download_updated",
            )
