# JAMA要件管理ツール

JAMAの要件を取得・更新するためのPythonツールです。

## 機能

1. **構造取得機能**: JAMAプロジェクトから要件構造を取得し、Excelファイルに出力
2. **要件更新機能**: Excelファイルで編集した内容をJAMAに反映（追加・修正・削除）

## セットアップ

### 1. 必要なパッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. 設定ファイルの作成

`config.json` ファイルを作成し、以下の内容を記入してください：

```json
{
  "base_url": "URL",
  "project_id": xxx,
  "api_id": "あなたのAPI ID",
  "api_secret": "あなたのAPI Secret",
  "proxies": {
    "http": "TEST",
    "https": "TEST"
  }
}
```

**API認証情報の取得方法:**
1. JAMAにログイン
2. ユーザープロファイルページへ移動
3. 「Set API Credentials」ボタンをクリック
4. アプリケーション名を入力
5. 「Create API Credentials」ボタンをクリック
6. 表示されるClient IDとClient Secretをメモ（Secretは一度しか表示されません！）

## 使い方

### テンプレートの作成（NEW!）

JAMAに接続せずに、練習用のExcelテンプレートを作成できます：

```bash
python main.py template -o template.xlsx
```

このコマンドで、サンプルデータが入ったExcelファイルが作成されます。JAMAの設定は不要なので、すぐに試せます。

### 要件構造の取得

#### プロジェクト全体を取得
```bash
python main.py fetch -o requirements.xlsx
```

#### 特定のコンポーネント以下を取得（sequenceで指定）
```bash
python main.py fetch -o requirements.xlsx -s 6.1.5
```

#### 特定のコンポーネント以下を取得（名前で指定）
```bash
python main.py fetch -o requirements.xlsx -n "TEST"
```

#### 最大3階層まで取得
```bash
python main.py fetch -o requirements.xlsx -d 3
```

### 要件の更新

#### ドライラン（実際の更新は行わない）
```bash
python main.py update -i requirements.xlsx --dry-run
```

#### 実際に更新を実行
```bash
python main.py update -i requirements.xlsx
```

## Excelファイルの構造

### シート1: Requirement_of_Driver

要件の一覧を表示します。各列の説明：

| 列 | 項目名 | 説明 |
|---|--------|------|
| A | JAMA_ID | JAMAのアイテムID（空欄=新規作成） |
| B | 操作 | 自動判定（新規/更新）または手動で「削除」を入力 |
| C | Sequence | 階層位置（例: 6.1.5.3） |
| D～N | 階層1～11 | 各階層のアイテム名 |
| O | アイテムタイプ | JAMAのアイテムタイプ |
| P | Assignee | 担当者 |
| Q | Status | ステータス |
| R | Tags | タグ（カンマ区切り） |
| S | Reason | 理由 |
| T | Preconditions | 前提条件 |
| U | Target_system | 対象システム |
| V | 現在のDescription | JAMAから取得した現在の説明 |
| W | Description更新 | 「する」または「しない」 |
| X | 新Description参照 | Description編集シートへのリンク |

### シート2: Description_edit

SYSPのDescriptionを5行形式で編集するためのシートです。

テーブル構造：
- 1行目: I/O Type（IN/OUT）
- 2行目: 項目名（(a)～(d)）
- 3行目: Data Name
- 4行目: Data Label
- 5行目: Data

列の配分：
- A列: 項目名
- B列: (a)Trigger action（1列）
- C～BQ列: (b)Behavior of ego-vehicle（64列）
- BR～CA列: (c)HMI（10列）
- CB～CF列: (d)Other（5列）

## 操作の流れ

1. **要件取得**
   ```bash
   python main.py fetch -o requirements.xlsx -n "Requirements of Driver"
   ```

2. **Excelで編集**
   - 新規要件の追加: 新しい行を追加し、必要な情報を記入
   - 要件の更新: 既存行の内容を変更
   - 要件の削除: B列（操作）に「削除」と入力
   - Descriptionの更新: W列を「する」に変更し、Description_editシートで新しい表を作成

3. **JAMAに反映**
   ```bash
   # まずドライランで確認
   python main.py update -i requirements.xlsx --dry-run
   
   # 問題なければ実行
   python main.py update -i requirements.xlsx
   ```

## トラブルシューティング

### エラー: "OAuth認証に失敗しました"
- API IDとAPI Secretが正しいか確認してください
- プロキシ設定が正しいか確認してください

### エラー: "設定ファイルが見つかりません"
- `config.json` ファイルを作成してください
- サンプルファイル `config.json.sample` が自動生成されるので、それを参考にしてください

### Excelファイルが開けない
- ファイルが他のプログラムで開かれていないか確認してください
- ファイルの拡張子が `.xlsx` であることを確認してください

## ログファイル

実行ログは `jama_tool.log` に保存されます。エラーが発生した場合は、このファイルを確認してください。

## 注意事項

- 大量の要件を一度に更新する場合は、必ずバックアップを取ってから実行してください
- API呼び出しにはレート制限がある可能性があります
- Description編集時は、Excelの結合セルに注意してください
