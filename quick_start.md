# JAMA要件管理ツール クイックスタートガイド

## 🚀 5分で始める

### 1. まずはテンプレートで練習（NEW!）

JAMAの設定なしで、すぐに試せます：

```bash
# 必要なパッケージをインストール
pip install requests openpyxl

# テンプレートを作成
python main.py template -o practice.xlsx
```

生成されたExcelファイルを開いて、どんな形式か確認してみましょう！

### 2. 本番用セットアップ（初回のみ）

```bash
# 必要なパッケージをインストール
pip install requests openpyxl

# config.json を作成
# 以下の内容をコピーして、あなたのAPI情報を入力
```

```json
{
  "base_url": "https://stargate.jamacloud.com",
  "project_id": 124,
  "api_id": "あなたのAPI ID",
  "api_secret": "あなたのAPI Secret",
  "proxies": {
    "http": "http://proxy1000.co.jp:15520",
    "https": "http://proxy1000.co.jp:15520"
  }
}
```

### 3. 基本的な使い方

#### 📝 練習用テンプレートで試す

```bash
# テンプレートを作成（JAMAへの接続不要）
python main.py template -o 練習用.xlsx
```

#### 📥 要件を取得する

```bash
# プロジェクト全体を取得
python main.py fetch -o 要件一覧.xlsx

# 「公共駐車場」以下だけを取得
python main.py fetch -o 公共駐車場.xlsx -n "公共駐車場"

# 6.1.5 以下だけを取得
python main.py fetch -o 部分取得.xlsx -s 6.1.5
```

#### ✏️ Excelで編集する

1. **新規追加**: JAMA_IDを空欄のまま（自動的に新規作成）
2. **更新**: W列（要件更新）を「する」に設定
3. **削除**: B列（メモ/コメント）に「削除」と入力
4. **スキップ**: W列を空欄または「しない」
5. **メモ**: B列に自由にコメント記入可能（「削除」以外）

#### 📤 JAMAに反映する

```bash
# まず確認（実際の更新はしない）
# 更新対象を全件表示
python main.py update -i 要件一覧.xlsx --dry-run

# 問題なければ実行
python main.py update -i 要件一覧.xlsx
```

### 更新時の進捗表示

```
Excelファイル読み込み進捗: 5000/13393 (37.3%)
要件分類進捗: 10000/13393 (74.6%)
  新規作成: 10件
  更新: 123件
  削除: 0件

【更新予定】123件
  1. ID: 12345, SYSP: ドライバがAP SWを押下した時...
     更新フィールド: description, tags, reason
  ...（全件表示）

更新開始: 123件
  進捗: 50/123 (40.7%)
```

## 📁 ファイル構成

```
jama-tool/
├── main.py              # メインプログラム
├── jama_client.py       # JAMA API通信
├── excel_handler.py     # Excel処理
├── config.py            # 設定管理
├── config.json          # 設定ファイル（要作成）
├── requirements.txt     # 必要パッケージ
└── jama_tool.log       # 実行ログ
```

## 💡 便利な使い方

### 特定の階層だけ取得

```bash
# 3階層目まで
python main.py fetch -o 浅い階層.xlsx -d 3

# 「出庫機能」以下を5階層まで
python main.py fetch -o 出庫機能.xlsx -n "出庫機能" -d 5
```

### バッチ処理

```bash
# 複数のコンポーネントを個別に取得
python main.py fetch -o 駐車場.xlsx -n "公共駐車場"
python main.py fetch -o 出庫.xlsx -n "出庫機能"
python main.py fetch -o 起動.xlsx -n "Advanced Park起動"
```

## ⚠️ 注意事項

1. **API Secret は一度しか表示されません** - 必ずメモしてください
2. **大量更新の前にはバックアップ** - JAMAの既存データを保護
3. **ドライラン推奨** - `--dry-run` で事前確認

## 🆘 困ったときは

### よくあるエラー

| エラー | 原因 | 対処法 |
|--------|------|---------|
| OAuth認証失敗 | API認証情報が間違い | config.json を確認 |
| ファイルが見つからない | ファイル名の間違い | ファイル名を確認 |
| プロキシエラー | ネットワーク設定 | プロキシ設定を確認 |
| 処理が止まったように見える | 大量データ処理中 | ログで進捗を確認 |

### 大量データ処理時のTips

```bash
# ログをリアルタイムで監視（別ターミナルで実行）
tail -f jama_tool.log

# 進捗表示の例
# 取得進捗: 50000/108708 (46.0%)
# 要件シート作成進捗: 10000/108708 (9.2%)
# SYSPアイテム数: 5000件
# Descriptionテンプレート作成進捗: 1000/5000 (20.0%)
```

**デバッグモードでフィールドを確認：**
```bash
python main.py fetch -o debug.xlsx --debug

# ログに以下のような情報が出力されます：
# === DEBUG: First Item structure ===
# Available fields: ['name', 'description', ...]
# 
# === DEBUG: SYSP Item structure ===
# SYSP Available fields: ['name', 'description', 'tags', 'reason', ...]
# SYSP Fields content ===
#   tags: タグ1, タグ2
#   reason: 理由の内容
```

**1時間以上かかる場合の対処法：**
1. 階層を制限して部分的に取得（`-d 3`オプション）
2. 特定コンポーネントのみ取得（`-n "コンポーネント名"`）
3. デバッグモードで問題を特定（`--debug`オプション）

### ログを確認

```bash
# 最新のエラーを確認
tail -n 50 jama_tool.log

# リアルタイムで監視
tail -f jama_tool.log
```

## 📞 サポート

問題が解決しない場合は、以下の情報と共に管理者に連絡：
- エラーメッセージ
- `jama_tool.log` の該当部分
- 実行したコマンド