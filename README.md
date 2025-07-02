# JAMA要件管理ツール

JAMAの要件を取得・更新するためのPythonツールです。

## 概要

このツールは、JAMA REST APIを使用して要件データをExcelファイルで管理するためのソフトウェアです。ソフトウェア開発の要件管理において、JAMAとExcelを連携させることで、効率的な要件の一括編集・更新を実現します。

## 主な機能

1. **要件取得機能**: JAMAプロジェクトから要件構造を取得し、Excelファイルに出力
2. **要件更新機能**: Excelファイルで編集した内容をJAMAに反映（新規作成・更新・削除）
3. **テンプレート作成機能**: JAMAに接続せずに練習用のExcelテンプレートを作成
4. **Description編集機能**: SYSPアイテムのDescriptionを5行形式のテーブルで編集

## 動作環境

- Python 3.10.0 以上
- Windows環境で動作確認済み
- プロキシ環境での使用が必須

## セットアップ

### 1. リポジトリのクローン

```bash
git clone [リポジトリURL]
cd jama-tool
```

### 2. 必要なパッケージのインストール

```bash
pip install -r requirements.txt
```

### 3. 設定ファイルの作成

`config.json` ファイルを作成し、以下の内容を記入してください：

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

**API認証情報の取得方法:**
1. JAMAにログイン
2. ユーザープロファイルページへ移動
3. 「Set API Credentials」をクリック
4. アプリケーション名を入力
5. 「Create API Credentials」をクリック
6. 表示されるClient IDとClient Secretをメモ
   - **重要**: Client Secretは一度しか表示されません！必ずメモしてください

## 使い方

### コマンド一覧

#### 1. テンプレート作成（練習用）

JAMAに接続せずに、サンプルデータ入りのExcelテンプレートを作成します：

```bash
python main.py template -o template.xlsx
```

#### 2. 要件の取得 (fetch)

##### プロジェクト全体を取得
```bash
python main.py fetch -o requirements.xlsx
```

##### 特定のコンポーネント以下を取得（sequenceで指定）
```bash
python main.py fetch -o requirements.xlsx -s 6.1.5
```

##### 特定のコンポーネント以下を取得（名前で指定）
```bash
python main.py fetch -o requirements.xlsx -n "公共駐車場"
```

##### オプション
- `-d, --max-depth`: 取得する最大階層数を指定
- `--count`: 取得する最大件数を指定（推奨上限: 16000）
- `-c, --config`: 設定ファイルのパス（デフォルト: config.json）

例：
```bash
# 3階層まで取得
python main.py fetch -o requirements.xlsx -d 3

# 最大5000件まで取得
python main.py fetch -o requirements.xlsx --count 5000
```

#### 3. 要件の更新 (update)

##### ドライラン（実際の更新は行わない）
```bash
python main.py update -i requirements.xlsx --dry-run
```

##### 実際に更新を実行
```bash
python main.py update -i requirements.xlsx
```

## Excelファイルの構造

### シート1: Requirement_of_Driver

要件の一覧を管理するメインシートです。

| 列 | 項目名 | 説明 | 編集可否 |
|---|--------|------|----------|
| A | JAMA_ID | JAMAのアイテムID | 新規作成時は空欄 |
| B | メモ/コメント | 自由記入欄 | ○ |
| C | Sequence | 階層位置（例: 6.1.5.3） | × |
| D～N | 階層1～11 | 各階層のアイテム名 | × |
| O | アイテムタイプ | 固定値: Requirement | × |
| P | Assignee | 担当者 | ○ |
| Q | Status | ステータス | ○ |
| R | Tags | タグ（カンマ区切り） | ○ |
| S | Reason | 理由 | ○ |
| T | Preconditions | 前提条件 | ○ |
| U | Target_system | 対象システム | ○ |
| V | 現在のDescription | 現在の説明（参照用） | × |
| W | 要件更新 | 「する」で更新対象 | ○ |
| X | 新Description参照 | Description編集シートへのリンク | × |

### シート2: Description_edit

SYSPアイテムのDescriptionを編集するための専用シートです。

#### テーブル構造（5行×81列）
- **1行目**: I/O Type（IN/OUT）
- **2行目**: 項目名
  - (a) Trigger action: 1列
  - (b) Behavior of ego-vehicle: 64列
  - (c) HMI: 10列
  - (d) Other: 5列
- **3行目**: Data Name
- **4行目**: Data Label
- **5行目**: Data

## 操作ロジック

### 更新判定ロジック

| A列（JAMA_ID） | B列（メモ） | W列（要件更新） | 動作 |
|----------------|------------|----------------|------|
| 空欄 | （任意） | （任意） | **新規作成** |
| あり | **削除**（完全一致） | （任意） | **削除** ⚠️ |
| あり | それ以外 | **する** | **更新** |
| あり | それ以外 | それ以外 | スキップ |

### ⚠️ 削除操作の注意事項

**B列（メモ/コメント）に「削除」と完全一致で入力した場合のみ、その要件は削除されます。**

- ✅ 削除される: 「削除」（完全一致）
- ❌ 削除されない: 「削除予定」「削除したい」「要削除」など

誤って「削除」と入力しないよう、十分注意してください。

### 更新時の動作

- **空欄のフィールドは更新対象外**（既存の値を保持）
- Descriptionのみの更新も可能
- 複数フィールドの同時更新も可能

## 操作の流れ（典型的な使用例）

### 1. 初回：要件構造の取得

```bash
# プロジェクト全体を取得
python main.py fetch -o requirements_20240101.xlsx

# または特定コンポーネントのみ
python main.py fetch -o driver_requirements.xlsx -n "Requirements of Driver"
```

### 2. Excelで編集

1. **新規要件の追加**
   - 最下行に追加
   - A列（JAMA_ID）は空欄のまま
   - 必要な情報を記入

2. **既存要件の更新**
   - W列（要件更新）を「する」に変更
   - 更新したいフィールドのみ編集
   - 空欄は既存値を保持

3. **要件の削除**
   - B列（メモ/コメント）に「削除」と入力
   - **注意**: 完全一致の「削除」のみ有効

4. **Description編集**（SYSPアイテムのみ）
   - X列のリンクをクリック
   - Description_editシートで5行テーブルを編集

### 3. JAMAへの反映

```bash
# 必ずドライランで確認
python main.py update -i requirements_20240101.xlsx --dry-run

# 確認後、実行
python main.py update -i requirements_20240101.xlsx
```

## エラーハンドリング

### よくあるエラーと対処法

| エラー | 原因 | 対処法 |
|--------|------|--------|
| OAuth認証に失敗しました | API認証情報の誤り | config.jsonのapi_id/api_secretを確認 |
| 設定ファイルが見つかりません | config.jsonが存在しない | config.jsonを作成（サンプルが自動生成される） |
| プロキシエラー | ネットワーク接続の問題 | プロキシ設定を確認 |
| Excelファイルが開けない | ファイルが使用中 | 他のプログラムで開いていないか確認 |

### ログファイル

詳細なエラー情報は `jama_tool.log` に記録されます：

```bash
# 最新のログを確認
tail -n 50 jama_tool.log

# リアルタイムで監視
tail -f jama_tool.log
```

## 注意事項

1. **プロキシ設定は必須です**
   - 社内ネットワークからJAMAにアクセスする場合、プロキシ設定が必要です

2. **大量データの更新**
   - 更新前に必ずJAMAのバックアップを取ってください
   - ドライラン機能で事前確認を推奨

3. **API制限**
   - 大量のAPI呼び出しを行う場合、レート制限に注意してください

4. **Excel編集時の注意**
   - 結合セルを作成しないでください
   - 数式は使用できません（値のみ）

## トラブルシューティング

### Q: 取得に時間がかかる

A: 大量のデータを取得する場合、以下の方法で高速化できます：
- 階層を制限: `-d 3` オプション
- 件数を制限: `--count 5000` オプション
- 特定コンポーネントのみ取得: `-n "コンポーネント名"`

### Q: Descriptionが更新されない

A: 以下を確認してください：
1. W列（要件更新）が「する」になっているか
2. SYSPアイテムの場合、Description_editシートで編集したか
3. 5行テーブルの形式が正しいか

### Q: 削除が実行されない

A: B列に入力した文字が完全に「削除」と一致しているか確認してください。前後の空白や他の文字が含まれていると削除されません。

## ファイル構成

```
jama-tool/
├── main.py              # メインプログラム
├── jama_client.py       # JAMA API通信モジュール
├── excel_handler.py     # Excel処理モジュール
├── config.py            # 設定管理モジュール
├── config.json          # 設定ファイル（要作成）
├── requirements.txt     # 必要パッケージリスト
├── README.md           # このファイル
└── jama_tool.log       # 実行ログ（自動生成）
```

## ライセンス

[ライセンス情報を記載]

## 作者

[作者情報を記載]