"""チャート分析モジュール - データ特性から最適なグラフ種類を自動判定"""

import re
from modules.models import ParsedTable, ChartSpec


# 時系列を示唆するパターン
TIME_PATTERNS = re.compile(
    r"(月|年|日|期|四半期|Q[1-4]|[0-9]{4}|1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月"
    r"|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
    r"|Week|Month|Year|Date|Quarter)",
    re.IGNORECASE,
)

# 構成比・割合を示唆するパターン
PROPORTION_PATTERNS = re.compile(
    r"(構成比|割合|シェア|比率|%|パーセント|percent|share|ratio)",
    re.IGNORECASE,
)


def analyze_table(table: ParsedTable) -> ChartSpec | None:
    """テーブルデータを分析し、最適なチャート仕様を返す。
    数値列がない場合はNoneを返す。
    """
    if not table.headers or not table.rows:
        return None

    # 数値列とカテゴリ列を判定
    num_cols = _detect_numeric_columns(table)
    if not num_cols:
        return None

    # カテゴリ列（最初の非数値列、通常はインデックス0）
    cat_col = _detect_category_column(table, num_cols)

    # データ特性を分析
    is_time_series = _is_time_series(table, cat_col)
    is_proportion = _is_proportion_data(table, num_cols)
    num_categories = len(table.rows)
    num_value_cols = len(num_cols)

    # チャート種類を決定
    chart_type, reason = _select_chart_type(
        is_time_series, is_proportion, num_categories, num_value_cols, table.headers, num_cols
    )

    return ChartSpec(
        chart_type=chart_type,
        title=table.title or "グラフ",
        category_column=cat_col,
        value_columns=num_cols,
        reason=reason,
    )


def _detect_numeric_columns(table: ParsedTable) -> list[int]:
    """数値データが含まれる列のインデックスを返す"""
    num_cols = []
    for col_idx in range(len(table.headers)):
        numeric_count = 0
        for row in table.rows:
            if col_idx < len(row):
                val = row[col_idx]
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    numeric_count += 1
                elif isinstance(val, str):
                    try:
                        float(val.replace(",", "").replace("%", ""))
                        numeric_count += 1
                    except (ValueError, AttributeError):
                        pass
        # 半数以上が数値ならば数値列とみなす
        if numeric_count > len(table.rows) / 2:
            num_cols.append(col_idx)
    return num_cols


def _detect_category_column(table: ParsedTable, num_cols: list[int]) -> int:
    """カテゴリ列（ラベル列）を検出"""
    for i in range(len(table.headers)):
        if i not in num_cols:
            return i
    return 0


def _is_time_series(table: ParsedTable, cat_col: int) -> bool:
    """時系列データかどうかを判定"""
    # ヘッダーに時間を示唆する語があるか
    header = table.headers[cat_col] if cat_col < len(table.headers) else ""
    if TIME_PATTERNS.search(header):
        return True

    # カテゴリ値に時間パターンがあるか
    for row in table.rows:
        if cat_col < len(row) and row[cat_col] is not None:
            val = str(row[cat_col])
            if TIME_PATTERNS.search(val):
                return True
    return False


def _is_proportion_data(table: ParsedTable, num_cols: list[int]) -> bool:
    """構成比・割合データかどうかを判定"""
    for col_idx in num_cols:
        header = table.headers[col_idx] if col_idx < len(table.headers) else ""
        if PROPORTION_PATTERNS.search(header):
            return True

        # 値の合計が約100なら割合データの可能性
        values = []
        for row in table.rows:
            if col_idx < len(row):
                val = _to_number(row[col_idx])
                if val is not None:
                    values.append(val)
        if values and 95 <= sum(values) <= 105:
            return True
    return False


def _to_number(val) -> float | None:
    """値を数値に変換"""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    if isinstance(val, str):
        try:
            return float(val.replace(",", "").replace("%", ""))
        except (ValueError, AttributeError):
            return None
    return None


def _select_chart_type(
    is_time_series: bool,
    is_proportion: bool,
    num_categories: int,
    num_value_cols: int,
    headers: list[str],
    num_cols: list[int],
) -> tuple[str, str]:
    """最適なチャート種類を選定"""

    # 時系列データ → 折れ線グラフ（構成比より優先）
    if is_time_series:
        if num_value_cols >= 2:
            return "line", "時系列の複数指標を折れ線グラフで推移を比較"
        return "line", "時系列データのため、折れ線グラフで推移を表示"

    # 構成比データ → 円グラフ（カテゴリが少ない場合）
    if is_proportion and num_categories <= 8:
        return "pie", "構成比・割合データのため、円グラフで内訳を視覚化"

    # 複数の値列がある比較データ → グループ棒グラフ
    if num_value_cols >= 2:
        return "column", "複数指標の比較のため、棒グラフでカテゴリ別に表示"

    # カテゴリが多い → 横棒グラフ
    if num_categories > 6:
        return "bar", "カテゴリ数が多いため、横棒グラフで比較"

    # デフォルト → 縦棒グラフ
    return "column", "カテゴリ別の数値比較のため、棒グラフで表示"
