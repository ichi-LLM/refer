# JAMA要件管理ツール クイックスタートガイド

## 🚀 5分で始める

### 1. まずは練習用テンプレートで試す

JAMAの設定なしで、すぐに試せます：

```bash
# 必要なパッケージをインストール
pip install -r requirements.txt

# テンプレートを作成
python main.py template -o practice.xlsx
```

生成されたExcelファイルを開いて、どんな形式か確認してみましょう！

### 2. 本番用セットアップ（初回のみ）

#### config.json を作成

以下の内容をコピーして、あなたのAPI情報を入力：

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

**⚠️ API Secret は一度しか表示されません！必ずメモしてください**

### 3. 基本的な使い方

#### 📥 要件を取得する

```bash
# プロジェクト全体を取得
python main.py fetch -o 要件一覧.xlsx

# 特定の場所だけ取得（名前で指定）
python main.py fetch -o 公共駐車場.xlsx -n "公共駐車場"

# 特定の場所だけ取得（番号で指定）
python main.py fetch -o 部分取得.xlsx -s 6.1.5
```

#### ✏️ Excelで編集する

| 操作 | 方法 |
|------|------|
| **新規追加** | A列（JAMA_ID）を空欄のまま |
| **更新** | W列（要件更新）を「する」に設定 |
| **削除** ⚠️ | B列（メモ/コメント）に「削除」と入力 |
| **スキップ** | W列を空欄または「しない」 |

**⚠️ 削除の注意**: B列に完全一致で「削除」と入力すると削除されます。「削除予定」などは削除されません。

#### 📤 JAMAに反映する

```bash
# 必ず最初はドライラン（確認のみ）
python main.py update -i 要件一覧.xlsx --dry-run

# 確認して問題なければ実行
python main.py update -i 要件一覧.xlsx
```

## 📋 Excel操作早見表

### 更新ロジック

| A列(JAMA_ID) | B列(メモ) | W列(要件更新) | 結果 |
|--------------|-----------|---------------|------|
| 空欄 | (任意) | (任意) | **新規作成** |
| あり | **削除** | (任意) | **削除** ⚠️ |
| あり | それ以外 | **する** | **更新** |
| あり | それ以外 | それ以外 | スキップ |

### 編集のポイント

- 空欄のフィールドは更新されません（既存値を保持）
- 複数フィールドの同時更新も可能
- SYSPアイテムはDescription編集シートで詳細編集可能

## 💡 便利な使い方

### 階層や件数を制限して取得

```bash
# 3階層目まで
python main.py fetch -o 浅い階層.xlsx -d 3

# 最大5000件まで
python main.py fetch -o 制限付き.xlsx --count 5000

# 組み合わせも可能
python main.py fetch -o 出庫機能.xlsx -n "出庫機能" -d 5 --count 1000
```

### 複数コンポーネントを個別に取得

```bash
python main.py fetch -o 駐車場.xlsx -n "公共駐車場"
python main.py fetch -o 出庫.xlsx -n "出庫機能"
python main.py fetch -o 起動.xlsx -n "Advanced Park起動"
```

## ⚠️ 注意事項

1. **プロキシ設定は必須** - config.jsonに正しく設定してください
2. **削除は取り消せません** - B列の「削除」入力は慎重に
3. **大量更新の前にバックアップ** - JAMAの既存データを保護
4. **ドライラン推奨** - `--dry-run` で必ず事前確認

## 🆘 困ったときは

### よくあるエラー

| エラー | 対処法 |
|--------|--------|
| OAuth認証失敗 | config.jsonのapi_id/api_secretを確認 |
| ファイルが見つからない | ファイル名とパスを確認 |
| プロキシエラー | プロキシ設定を確認 |
| Excelが開けない | 他のプログラムで開いていないか確認 |

### ログを確認

```bash
# 最新のエラーを確認
tail -n 50 jama_tool.log
```

## 📁 ファイル構成

```
jama-tool/
├── main.py              # メインプログラム
├── config.json          # 設定ファイル（要作成）
├── requirements.txt     # 必要パッケージ
└── jama_tool.log       # 実行ログ
```

## 🎯 典型的な作業フロー

1. **月初めに全体を取得**
   ```bash
   python main.py fetch -o 要件_202401.xlsx
   ```

2. **Excelで1ヶ月分の更新作業**
   - 新規要件を追加
   - 既存要件を更新
   - 不要な要件を削除

3. **月末にJAMAへ反映**
   ```bash
   python main.py update -i 要件_202401.xlsx --dry-run
   python main.py update -i 要件_202401.xlsx
   ```

---

**準備完了！** 実際に試してみましょう 🚀