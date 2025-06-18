import pandas as pd

class ExcelManager:
    """
    Excelファイルの読み書きを管理するクラス。
    """
    def __init__(self, filepath):
        self.filepath = filepath
        self.columns = [
            'JAMA ID', '操作', '階層レベル', '要件/項目名', 'Item Type', '担当者',
            'Description種別', 'Data Name', 'Data Label', 'Data'
        ]

    def load_requirements_from_excel(self):
        """Excelから要件データを読み込み、DataFrameとして返す。"""
        try:
            df = pd.read_excel(self.filepath, sheet_name='Requirements', dtype={'JAMA ID': 'Int64'})
            # 空の行を除外
            df = df.dropna(how='all')
            # JAMA IDが数値でない場合は欠損値にする
            df['JAMA ID'] = pd.to_numeric(df['JAMA ID'], errors='coerce').astype('Int64')
            return df
        except FileNotFoundError:
            print(f"エラー: Excelファイルが見つかりません: {self.filepath}")
            return pd.DataFrame()
        except Exception as e:
            print(f"エラー: Excelファイルの読み込み中にエラーが発生しました: {e}")
            return pd.DataFrame()

    def export_to_excel(self, jama_data):
        """JAMAから取得したデータを整形してExcelファイルに出力する。"""
        print(f"取得したデータをExcelファイルに出力しています: {self.filepath}")
        
        records = []
        for item in sorted(jama_data, key=lambda x: x['location']['globalSortOrder']):
            # ここでHTMLのDescriptionを解析する処理が必要になる
            # 今回は主要なフィールドのみをフラットに出力する
            fields = item.get('fields', {})
            record = {
                'JAMA ID': item.get('id'),
                '操作': '',
                '階層レベル': len(item.get('location', {}).get('sequence', '').split('.')),
                '要件/項目名': fields.get('name', ''),
                'Item Type': item.get('itemType'), # 本来はitemtype名に変換すべき
                '担当者': fields.get('assigned', ''), # 本来は担当者名に変換すべき
                'Description種別': '',
                'Data Name': '',
                'Data Label': '',
                'Data': ''
            }
            records.append(record)
            
        df = pd.DataFrame(records, columns=self.columns)
        
        try:
            with pd.ExcelWriter(self.filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Requirements', index=False)
            print("Excelファイルへの出力が完了しました。")
        except Exception as e:
            print(f"エラー: Excelファイルへの書き込み中にエラーが発生しました: {e}")
            
    def create_template(self):
        """空のExcelテンプレートを作成する。"""
        print(f"新しいExcelテンプレートを作成しています: {self.filepath}")
        df = pd.DataFrame(columns=self.columns)
        try:
            with pd.ExcelWriter(self.filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Requirements', index=False)
            print(f"テンプレートファイル '{self.filepath}' を作成しました。")
        except Exception as e:
            print(f"エラー: テンプレートファイルの作成中にエラーが発生しました: {e}")