# prz - プレゼン自動生成

ExcelデータからPowerPointプレゼンテーションを自動生成するアプリケーション。

## セットアップ

```bash
pip install -r requirements.txt
```

## 起動

```bash
streamlit run app.py
```

## AI機能を有効にする

```bash
export ANTHROPIC_API_KEY=your-api-key
```

## 機能

### Phase 1 - テーブル転記
- Excelファイルのインポートとテーブル自動検出
- 対象者・目的のコンテキスト設定
- スライド構成のプレビュー
- PowerPointファイル（.pptx）のエクスポート

### Phase 2 - AI統合（Claude API）
- ストーリー構成の自動生成（導入→現状分析→課題→解決策→まとめ）
- テーブルデータの要約テキスト・考察の自動生成
- 発表者ノートの自動生成
- サイドバーでAI機能のON/OFF切替

### Phase 3 - グラフ自動描画
- データ特性の自動分析（時系列/構成比/比較）
- 最適なグラフ種類の自動選定（棒/折れ線/円グラフ）
- テーブル+グラフ並列レイアウト
- チャート種類のカスタマイズUI

## 開発ロードマップ

- **Phase 1 (MVP)**: Excelテーブル → PowerPointへの転記 ✅
- **Phase 2 (AI統合)**: Claude APIによる要約テキスト・考察の自動生成 ✅
- **Phase 3 (グラフ化)**: 数値データの最適チャート自動描画 ✅

## テスト

```bash
python -m pytest tests/ -v
```
