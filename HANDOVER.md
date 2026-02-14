# オプション建玉分析機能 実装引継ぎ

## セッション概要
日経225オプションの建玉・取引高分析機能を既存の先物分析アプリに追加する作業を開始しました。

## 完了した作業

### 1. データモデル定義 ✅
**ファイル**: `models.py`

以下の3つのdataclassを追加：
```python
@dataclass
class OptionParticipantOI:
    """オプション建玉（参加者別）"""
    report_date: date
    option_type: str            # "PUT" or "CALL"
    strike_price: int           # 権利行使価格
    participant_id: str
    participant_name_jp: str
    long_volume: Optional[float]
    short_volume: Optional[float]

@dataclass
class OptionParticipantVolume:
    """オプション取引高（参加者別）"""
    trade_date: date
    option_type: str
    strike_price: int
    participant_id: str
    participant_name_en: str
    participant_name_jp: str
    rank: int
    volume: float
    volume_day: float
    volume_night: float

@dataclass
class OptionStrikeRow:
    """行使価格別集約行（UI表示用）"""
    strike_price: int
    put_start_oi_net: Optional[float] = None
    put_daily_volumes: dict = field(default_factory=dict)
    put_week_total: Optional[float] = None
    call_start_oi_net: Optional[float] = None
    call_daily_volumes: dict = field(default_factory=dict)
    call_week_total: Optional[float] = None
```

### 2. 設定ファイル拡張 ✅
**ファイル**: `config.py`

追加した定数：
```python
PRODUCT_DISPLAY_NAMES["NK225OP"] = "日経225オプション"

OPTION_OI_SECTION_KEYWORDS = {
    "PUT": ["put", "プット", "PUT"],
    "CALL": ["call", "コール", "CALL"],
}

OPTION_STRIKE_DISPLAY_RANGE = {
    "ATM±5": 5,
    "ATM±10": 10,
    "ATM±20": 20,
    "全て": 9999,
}
```

### 3. データ取得 ✅
- オプション建玉Excelファイルをダウンロード済み
- 場所: `cache/oi/20260130_nk225op_oi_by_tp.xlsx`
- URLパターン確認済み: `/automation/markets/derivatives/open-interest/files/{YYYY}/{YYYYMMDD}_nk225op_oi_by_tp.xlsx`

### 4. 実装計画策定 ✅
**ファイル**: `C:\Users\kawai\.claude\plans\linked-purring-flame.md`

詳細な実装計画を作成済み。

## 未完了の作業

### Phase 3: パーサー実装（次のステップ）
#### 3.1 オプション建玉パーサー
**新規ファイル**: `data/parser_option_oi.py`

実装内容：
- `parse_option_oi_excel(content: bytes) -> list[OptionParticipantOI]`
- PUT/CALLセクション検出
- 行使価格ごとのグループ化
- 参加者データ抽出

再利用する既存パターン：
- `parser_oi.py`の`_find_section_headers()` - セクション検出
- `_extract_report_date()` - 日付抽出
- `_consolidate_long_short()` - ロング/ショート統合

#### 3.2 取引高パーサー拡張
**修正ファイル**: `data/parser_volume.py`

修正内容：
- `parse_volume_excel()`に`include_options`パラメータ追加
- product列で"NK225 CALL", "NK225 PUT"等を判定
- contract列から行使価格抽出
- option_typeとstrike_priceをパース

### Phase 4: 集約ロジック
**ファイル**: `data/aggregator.py`

新規関数：
- `load_option_weekly_data(week, session_keys) -> list[OptionStrikeRow]`
- `_aggregate_by_strike()` - 行使価格別に全参加者を集約

### Phase 5: UI実装
#### 5.1 テーブルコンポーネント
**新規ファイル**: `ui/option_strike_table.py`

実装内容：
- `render_option_strike_table(rows, week)`
- `_build_option_display_df(rows, week)` - PUT/CALL左右分割レイアウト

列構成：
```
PUT_前週Net | PUT_1/6 | PUT_1/7 | ... | PUT_1/10 | PUT_週計 | 行使価格 |
CALL_週計 | CALL_1/10 | ... | CALL_1/7 | CALL_1/6 | CALL_前週Net
```

**重要**:
- PUT側日付列は左から右へ（古→新）
- CALL側日付列は右から左へ（新→古）
- 行使価格は降順ソート

#### 5.2 サイドバー拡張
**修正ファイル**: `ui/sidebar.py`

追加内容：
- オプション用セレクタ
- 表示範囲選択（ATM±5, ±10, ±20, 全て）

#### 5.3 メインアプリ統合
**修正ファイル**: `app.py`

追加内容：
```python
main_tab1, main_tab2 = st.tabs(["先物分析", "オプション分析"])

with main_tab2:
    selections = render_sidebar_option()
    rows = _cached_option_weekly_data(...)
    render_option_strike_table(rows, week)
```

## データソース情報

### 建玉データ
- **URL**: `https://www.jpx.co.jp/automation/markets/derivatives/open-interest/files/{YYYY}/{YYYYMMDD}_nk225op_oi_by_tp.xlsx`
- **取得方法**: 既存の`fetcher.get_oi_index()`でentry["IndexOptions"]を使用
- **構造**: Excel内部構造は未確認（実装時に確認必要）

### 取引高データ
- **URL**: 先物と同じファイル内に混在
- **場所**: `/automation/markets/derivatives/participant-volume/files/daily/{YYYYMM}/{YYYYMMDD}_volume_by_participant_{session}.xlsx`
- **判定**: product列で区別（"NK225 CALL", "NK225 PUT"等を想定）

## UI要件（重要）

### レイアウト
```
PUT側（左）                                      | CALL側（右）
-------------------------------------------------|--------------------------------------------------
前週Net | 1/6 1/7 1/8 1/9 1/10 | 合計 | 行使価格 | 合計 | 1/10 1/9 1/8 1/7 1/6 | 前週Net
38500   | 1200 ... 900          | 5800 | 38500   | 4200 | 950 ... 1100        | -1500
38000   | ...                   | ...  | 38000   | ...  | ...                 | ...
```

- **行使価格**: 縦軸、高い順に上から下へ
- **PUT（左側）**: 前週Net → 日次取引高（左が古く右が新しい）→ 週合計 → 行使価格
- **CALL（右側）**: 行使価格 → 週合計 → 日次取引高（右が古く左が新しい）→ 前週Net
- **集計**: 各行使価格について全業者合計のネット建玉と取引高

## 技術的課題

### 解決済み
- ✅ データソースURL確認
- ✅ オプションExcelファイルダウンロード
- ✅ データモデル設計

### 未解決
- ⚠️ **Excel内部構造未確認** - 実装前に`cache/oi/20260130_nk225op_oi_by_tp.xlsx`の構造を解析する必要あり
  - セクション分割（PUT/CALL別か、行使価格別か）
  - 列配置（権利行使価格、建玉数、参加者の位置）
  - ヘッダー構造
- ⚠️ 取引高Excelのproduct列値未確認
- ⚠️ contract列からの行使価格抽出方法未確認

### 解析スクリプト
Excel構造解析用スクリプトを作成済み：
- `scripts/analyze_option_structure.py` - 構造解析
- `scripts/fetch_option_sample.py` - サンプルダウンロード

実行方法：
```bash
cd "C:\Users\kawai\Desktop\実験室01\先物手口分析"
python scripts/analyze_option_structure.py
```

## 実装優先順位

1. **最優先**: Excel構造確認（`scripts/analyze_option_structure.py`実行）
2. `data/parser_option_oi.py` 新規作成
3. `data/parser_volume.py` 拡張
4. `data/aggregator.py` 集約関数追加
5. `ui/option_strike_table.py` 新規作成
6. `ui/sidebar.py` 拡張
7. `app.py` タブ統合

## 参考ファイル

既存コードのパターンを参照：
- **セクション検出**: `data/parser_oi.py` の`_find_section_headers()`
- **日付抽出**: `data/parser_oi.py` の`_extract_report_date()`
- **DataFrame構築**: `ui/weekly_table.py` の`_build_display_dataframe()`
- **スタイリング**: `ui/weekly_table.py` の`_apply_table_styling()`
- **Night sessionシフト**: `data/aggregator.py` の該当ロジック

## 検証項目

実装完了後の確認事項：
1. 実データで2週間分の動作確認
2. 行使価格ソート（降順）確認
3. PUT/CALL左右配置の視認性確認
4. 全業者合計の計算精度検証
5. Night session日付シフト検証

## 注意事項

- 先物は「参加者別」分析、オプションは「行使価格別・全業者合計」分析
- PUT/CALLで日付列の並び順が逆（UI要件参照）
- 行使価格数が多いため、デフォルトATM±10に制限推奨
- 既存のキャッシュ機構を活用してパフォーマンス確保

## 次のセッションで最初にやること

1. `python scripts/analyze_option_structure.py` を実行してExcel構造確認
2. 出力結果を基に`data/parser_option_oi.py`の実装を開始
3. 先物パーサー（`data/parser_oi.py`）のパターンを参考に実装

---

作成日: 2026-02-08
プロジェクト: JPX先物手口分析ツール
機能: 日経225オプション建玉・取引高分析
