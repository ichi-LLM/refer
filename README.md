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

### デバッグモード（拡張版）

フィールド名が不明な場合、デバッグモードで確認できます：

```bash
# コマンドラインオプション
python main.py fetch -o debug.xlsx --debug

# または設定ファイルで
{
  "debug": true
}
```

デバッグモードでは以下の情報がログに出力されます：
1. **最初のアイテム**のフィールド構造
2. **最初のSYSPアイテム**のフィールド構造（Tags、Reason等を確認）

### SYSP項目の拡張フィールド取得

SYSPアイテムは自動的に以下のフィールドを検索・取得します：
- assignee / Assignee
- status / Status  
- tags / Tags / tag / Tag
- reason / Reason / reasons / Reasons
- preconditions / Preconditions / precondition / Precondition
- target_system / Target_system / targetSystem / target / Target

フィールド名の大文字小文字や区切り文字の違いを自動的に処理します。

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
python main.py fetch -o requirements.xlsx -n "公共駐車場"
```

#### 最大3階層まで取得
```bash
python main.py fetch -o requirements.xlsx -d 3
```

#### デバッグモードで実行（フィールド情報を確認）
```bash
python main.py fetch -o requirements.xlsx --debug
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
| B | メモ/コメント | 自由記入欄（処理には影響しない） |
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
| W | 要件更新 | 「する」で更新対象（空欄=スキップ） |
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
   - 新規要件の追加: JAMA_IDを空欄にして、必要な情報を記入
   - 要件の更新: 既存要件のW列（要件更新）を「する」に変更
   - 要件の削除: B列（メモ/コメント）に「削除」と入力
   - Descriptionの更新: Description_editシートで新しい表を作成

3. **JAMAに反映**
   ```bash
   # まずドライランで確認（更新対象を全件表示）
   python main.py update -i requirements.xlsx --dry-run
   
   # 問題なければ実行
   python main.py update -i requirements.xlsx
   ```

### 更新ロジック

| A列（JAMA_ID） | B列（メモ） | W列（要件更新） | 動作 |
|----------------|------------|----------------|------|
| 空欄 | （任意） | （任意） | 新規作成 |
| あり | 「削除」 | （任意） | 削除 |
| あり | それ以外 | 「する」 | 更新（空欄フィールドは保持） |
| あり | それ以外 | それ以外 | スキップ |

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
- **大量データの処理**: 1000件以上の要件を処理する場合、Excel作成に時間がかかることがあります。進捗表示を確認しながらお待ちください。

## 重要な変更点（最新）

### 更新ロジックの改善
- B列を「メモ/コメント」欄に変更（自由記入、処理には影響しない）
- W列を「要件更新」に変更（「する」で更新対象）
- JAMA_IDが空欄 → 新規作成
- JAMA_IDあり＋W列「する」 → 更新
- それ以外 → スキップ

### dry-run表示の改善
- 更新対象を**全件表示**（件数に関係なく）
- 各要件の更新フィールドを明示
- 意図しない更新を事前に確認可能

#### dry-run表示の例

```
【更新予定】123件
  1. ID: 12345, SYSP: ドライバがAP SWを押下した時、AdvancedParkを起動する
     更新フィールド: description, tags, reason
  2. ID: 12346, SYSP: センサーが障害物なしを検知した時、車両を後退させる
     更新フィールド: description, preconditions
  ...
  123. ID: 99999, SYSP: 最後の要件名
     更新フィールド: description, tags, reason, status
```

### 進捗表示の追加
- Excelファイル読み込み時
- 要件分類時
- 更新処理時（10件ごと）

### SYSP項目の処理（拡張版）
- nameフィールドに「SYSP」を含む項目のみDescriptionテンプレートを作成
- すべてのSYSP項目に対してテンプレートを自動生成（制限なし）
- 「編集画面へ」リンクと「一覧に戻る」リンクが正しく機能
- **各テンプレートに要件名を表示**（リンク先の上のセル）
- **SYSPアイテムの全フィールドを自動検出**（Tags、Reason等）

### データ保護機能
- Excelで空欄のセルは更新対象から除外（既存値を保持）
- Description更新時も他のフィールドに影響しない
- 安全なアップロードを実現

### デバッグ機能（拡張版）
- `--debug`オプションでフィールド構造を確認可能
- 最初のアイテムと最初のSYSPアイテムの両方を表示
- 実際のフィールド名が不明な場合に有用

## パフォーマンスと進捗表示

大量のデータを扱う際の進捗表示：

- **JAMA取得時**: `取得進捗: 5000/10000 (50.0%)`
- **Excel作成時**: 
  - `要件シート作成進捗: 1000/5000 (20.0%)`
  - `Descriptionテンプレート作成進捗: 10/50 (20.0%)`
- **列幅調整時**: `列幅調整進捗: 10/24`

10万件以上のデータでも、進捗を確認しながら安心して処理できます。