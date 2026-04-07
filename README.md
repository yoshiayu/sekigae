# 受講生席替えアプリ (Streamlit + SQLite)

毎週の席替え運用を想定したMVPです。  
73名を13テーブル(最大6名/テーブル)に対して、以下を考慮しながら自動割当します。

- スキル分散（高い / 並 / 低い / ヤバい）
- 同じ会社の偏り回避
- 前回同席ペアの再同席回避
- 収容制約（必ず全員配置、上限超過なし）

## 採用技術

- フロント/アプリ: Streamlit
- ロジック: Python
- DB: SQLite (`seating_app.db`)
- データ表示: pandas

## ディレクトリ構成

```text
.
├─ app.py
├─ requirements.txt
├─ README.md
├─ seating_app.db               # 初回起動時に自動作成
├─ data
│  ├─ sample_students.csv
│  └─ dummy_students_73.csv
├─ scripts
│  └─ smoke_test.py
└─ src
   ├─ __init__.py
   ├─ constants.py
   ├─ db.py
   ├─ repository.py
   ├─ csv_service.py
   └─ seating.py
```

## セットアップ

```bash
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 実行手順

```bash
streamlit run app.py
```

ブラウザで表示されたURLを開いて利用します。

## 主な機能

- 受講生管理
- 追加 / 編集 / 削除
- 検索
- CSVインポート（追記 or 全置換）
- CSVテンプレートダウンロード
- ダミー73名データ投入

- 席替え実行
- 条件重み設定（同じ会社回避・前回同席回避・スキル分散・人数均等化）
- 試行回数指定（100〜500）
- 乱数シード指定（任意）

- 結果画面
- テーブル1〜13を表示
- スキル色分け表示
- 会社重複警告 / 前回ペア重複表示
- CSV出力
- 手動調整（`table_no` 直接編集）
- 履歴保存

- 履歴管理
- 週ごとの結果一覧
- 履歴詳細表示
- 履歴CSV出力
- 過去履歴の結果タブ再読み込み

## データ構造

### Student (`students`)

- id
- name
- company
- skill_level
- created_at
- updated_at

### SeatingHistory (`seating_histories`)

- id
- target_week
- settings_json
- total_score
- overlap_rate
- created_at

### SeatingAssignment (`seating_assignments`)

- id
- seating_history_id
- table_no
- student_id

## CSV仕様

インポート列:

```csv
name,company,skill_level
```

`skill_level` の許可値:

- 高い
- 並
- 低い
- ヤバい

不正時は行番号付きでエラー表示します。

## 席替えアルゴリズム

`src/seating.py` の `generate_best_assignment` を使用。

1. 初期化
- 受講生をスキル人数・会社人数情報付きで並べ替え（レアスキル優先、人数多い会社優先）
- テーブル期待人数（73名なら6名×8卓、5名×5卓）を計算

2. 1回の割当試行
- 各受講生について、全テーブル候補のスコアを計算
- スコア項目:
- 同じ会社が既にいる人数（大ペナルティ）
- 前回同席ペア人数（中ペナルティ）
- スキル偏り悪化度（大ペナルティ）
- 期待人数からのずれ（小ペナルティ）
- ランダム揺らぎ（低優先）
- 上位候補から確率的に選択して配置

3. 複数回試行
- 100〜500回試行
- 総合スコア最小案を採用

4. 評価
- 会社重複数
- 前回同席ペア重複数 / 重複率
- スキル偏差
- テーブル人数偏差

## エラー時の挙動

- 上限超過（13×6を超える）は実行不可
- CSV不正は行番号付きで表示
- 手動調整で6名超過や範囲外テーブル指定があれば反映拒否
- 席替え生成失敗時は例外メッセージを表示

## スモークテスト

```bash
python scripts/smoke_test.py
```

期待出力:

```text
smoke test passed
score=...
```

## 今後の改善案

- ドラッグ&ドロップUIの追加（現状はtable_no編集方式）
- 男女/年齢/コース分散を重み付き制約として追加
- 履歴を使った複数週連続の重複最小化
- 手動調整の操作履歴（Undo/Redo）
- 権限管理（教務/講師）
- PostgreSQL対応

