# 受講生席替えアプリ (Streamlit + SQLite)

毎週の席替え運用を想定したMVPです。  
73名を13テーブル(最大6名/テーブル)に対して、以下を考慮しながら自動割当します。

- スキル固め（`高い` / `並` / `低い+ヤバい` の3グループ）
- 同じ会社の偏り回避
- 前回同席ペアの再同席回避
- 収容制約（必ず全員配置、上限超過なし）

## 採用技術

- フロント/アプリ: Streamlit
- ロジック: Python
- DB: SQLite (`seating_app.db`)
- データ表示: pandas

## 動作環境

- Python 3.11 以上推奨（3.13で動作確認）
- `pip` が使えること
- ブラウザで `http://127.0.0.1:8501` にアクセスできること

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
├─ logs
│  ├─ streamlit.out.log         # バックグラウンド起動時に作成
│  └─ streamlit.err.log         # バックグラウンド起動時に作成
├─ scripts
│  ├─ smoke_test.py
│  ├─ run_streamlit_background.ps1
│  └─ install_startup_task.ps1
└─ src
   ├─ __init__.py
   ├─ constants.py
   ├─ db.py
   ├─ repository.py
   ├─ csv_service.py
   ├─ layout_export.py
   └─ seating.py
```

## セットアップ

Windows (PowerShell):

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## 実行手順

どのPCでもコマンド差分を減らすため、`python -m streamlit` で実行します。

```bash
python -m streamlit run app.py
```

ポート変更したい場合:

```bash
python -m streamlit run app.py --server.port 8502
```

## Windowsでバックグラウンド常駐起動

以下を1回実行すると、ログオン時に自動起動するよう設定します。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\install_startup_task.ps1
```

補足:

- 可能なら Scheduled Task を作成します
- 権限不足で失敗した場合は Startup フォルダ方式に自動フォールバックします
- ログは `logs/streamlit.out.log` と `logs/streamlit.err.log` に出力されます

手動でバックグラウンド起動する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run_streamlit_background.ps1
```

## 主な機能

- 受講生管理（追加 / 編集 / 削除 / 検索）
- CSVインポート（追記 or 全置換）
- スキル一括更新（名前は変更せず、スキルのみ更新）
- ダミー73名データ投入
- 席替え実行（テーブル数 + テーブル最小人数/最大人数設定、重み設定、試行回数、乱数シード）
- テーブル表示（テーブルごとの枠表示）
- 手動調整（`table_no` 直接編集）
- CSV出力 / 履歴保存 / 履歴再読み込み
- 「○月○日の座席表」命名での Excel(.xlsx) 出力（通常版 / 反転版の2シート）
  - 通常版: 後ろから前（スクリーン方向）を見た配置
  - 反転版: 前から後ろを見た配置
- Googleスプレッドシートへの自動作成（テンプレート複製 + 通常版 / 反転版）

## Googleスプレッドシート自動出力の事前準備

1. Google Cloud でサービスアカウントを作成し、JSONキーを発行します。
2. JSON内の `client_email` を出力先スプレッドシートへ「編集者」で共有します。
3. アプリの「席替え結果」タブで以下を入力します。
   - 出力先スプレッドシートURL/ID
   - テンプレートシート名（例: `座席表（テンプレ）`）
   - サービスアカウントJSON
4. 「Googleスプレッドシートへ2種類を作成」を押すと、通常版と反転版を自動作成します。

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

## スキル一括更新の貼り付け形式

以下のいずれかで貼り付け可能です。

- `受講生<TAB>スキル`（Excelの2列コピペ）
- `受講生=スキル`
- `受講生:スキル`
- `受講生,スキル`

## 席替えアルゴリズム概要

`src/seating.py` の `generate_best_assignment` を使用します。

- 同じ会社の重複
- 前回同席ペアの再発
- スキルグループ混在ペア（混ざるほどペナルティ）
- テーブル人数の偏差

をスコア化し、複数回試行で最小スコア案を採用します。

補足:

- 配席は「各テーブル min〜max の範囲」に収め、できるだけ均等になるように割り当てます。
- 同スキル固めの重みを強化しています。
- スキル同席制約:
  - 禁止: `高い×低い`, `高い×ヤバい`
  - `並×ヤバい` は同席可
- 配席探索は段階フォールバックしますが、上記の禁止同席制約は常に維持します。

## スモークテスト

```bash
python scripts/smoke_test.py
```

期待出力:

```text
smoke test passed
score=...
```

## トラブルシュート

- `streamlit` が見つからない:
  - `python -m pip install -r requirements.txt`
  - 実行は `python -m streamlit run app.py` を使用
- `Port 8501 is not available`:
  - 既存プロセスを停止するか、`--server.port` で別ポートを使用
