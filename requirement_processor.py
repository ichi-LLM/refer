import pandas as pd

class RequirementProcessor:
    """
    JAMAとExcel間のデータ処理を行うクラス。
    """
    def __init__(self, jama_client, field_mapping):
        self.client = jama_client
        self.field_mapping = field_mapping
        self.itemtype_map = {}
        self.user_map = {}

    def _initialize_maps(self):
        """ItemTypeとUserのIDと名前の対応表を初期化する。"""
        print("ItemTypeとユーザーの情報をJAMAから取得しています...")
        # ItemType ID -> Name
        itemtypes = self.client.get_itemtypes()
        self.itemtype_map = {it['id']: it['display'] for it in itemtypes}
        # User ID -> Full Name
        users = self.client.get_users()
        self.user_map = {u['id']: f"{u['firstName']} {u['lastName']}".strip() for u in users}

    def _get_name_from_id(self, map_dict, item_id):
        return map_dict.get(item_id, str(item_id))

    def _get_id_from_name(self, map_dict, name):
        # 逆引きマップを作成
        reverse_map = {v: k for k, v in map_dict.items()}
        return reverse_map.get(name)

    def _build_description_html(self, desc_rows):
        """Excelの複数行データからDescription用のHTMLテーブルを生成する。"""
        if desc_rows.empty:
            return ""

        # ここでご要望の複雑なHTMLテーブル構造を生成します
        # 今回は簡易的なサンプルとして基本的なテーブルを作成します
        html = "<table border='1'><tbody>"
        html += "<tr><th>Type</th><th>Name</th><th>Label</th><th>Data</th></tr>"
        for _, row in desc_rows.iterrows():
            html += (
                f"<tr>"
                f"<td>{row.get('Description種別', '')}</td>"
                f"<td>{row.get('Data Name', '')}</td>"
                f"<td>{row.get('Data Label', '')}</td>"
                f"<td>{row.get('Data', '')}</td>"
                f"</tr>"
            )
        html += "</tbody></table>"
        return html

    def jama_to_excel(self, excel_manager, project_id, component_id=None):
        """JAMAからデータを取得し、Excelに出力する。"""
        self._initialize_maps()
        items = self.client.get_items_in_project(project_id, component_id)
        
        # IDを名前に変換
        for item in items:
            item['itemType'] = self._get_name_from_id(self.itemtype_map, item['itemType'])
            if 'fields' in item and self.field_mapping['assignee'] in item['fields']:
                assignee_id = item['fields'][self.field_mapping['assignee']]
                item['fields']['assigned'] = self._get_name_from_id(self.user_map, assignee_id)
        
        excel_manager.export_to_excel(items)

    def excel_to_jama(self, excel_manager, project_id):
        """Excelからデータを読み込み、JAMAに反映させる。"""
        self._initialize_maps()
        df = excel_manager.load_requirements_from_excel()
        if df.empty:
            print("Excelデータが空のため、処理を中断します。")
            return
            
        # 親アイテムの行を特定 (Description種別が空の行)
        parent_rows = df[df['Description種別'].isna() | (df['Description種別'] == '')]

        for index, row in parent_rows.iterrows():
            jama_id = row['JAMA ID'] if pd.notna(row['JAMA ID']) else None
            operation = row.get('操作', '').upper()

            # --- 削除処理 ---
            if operation == 'DELETE' and jama_id:
                try:
                    self.client.delete_item(jama_id)
                except Exception as e:
                    print(f"エラー: アイテム (ID: {jama_id}) の削除に失敗しました。")
                continue

            # --- Description部分の子行を取得 ---
            desc_rows = pd.DataFrame()
            if jama_id:
                desc_rows = df[(df['JAMA ID'] == jama_id) & (df.index > index) & (df['Description種別'].notna())]
            else: # 新規作成の場合
                # 次の親行までの間を子行とみなす
                next_parent_indices = parent_rows[parent_rows.index > index].index
                next_parent_index = next_parent_indices[0] if len(next_parent_indices) > 0 else len(df)
                desc_rows = df.loc[index + 1 : next_parent_index - 1]

            # --- データペイロードの構築 ---
            fields = {}
            fields[self.field_mapping['name']] = row['要件/項目名']
            fields[self.field_mapping['description']] = self._build_description_html(desc_rows)
            
            # 担当者を名前からIDに変換
            assignee_name = row['担当者']
            if pd.notna(assignee_name):
                assignee_id = self._get_id_from_name(self.user_map, assignee_name)
                if assignee_id:
                    fields[self.field_mapping['assignee']] = assignee_id
                else:
                    print(f"警告: 担当者 '{assignee_name}' がJAMAに見つかりません。")

            payload = {
                "fields": fields
            }

            # --- 新規作成または更新処理 ---
            if jama_id: # 更新
                try:
                    self.client.update_item(jama_id, payload)
                except Exception as e:
                    print(f"エラー: アイテム (ID: {jama_id}) の更新に失敗しました。")
            else: # 新規作成
                item_type_name = row['Item Type']
                item_type_id = self._get_id_from_name(self.itemtype_map, item_type_name)
                if not item_type_id:
                    print(f"警告: Item Type '{item_type_name}' が見つかりません。この行はスキップされます。")
                    continue
                
                payload['project'] = project_id
                payload['itemType'] = item_type_id
                # locationの指定 (親アイテムの指定など) が必要
                # ここではプロジェクトルートに作成する単純な例
                payload['location'] = { "parent": { "project": project_id } }

                try:
                    self.client.create_item(payload)
                except Exception as e:
                    print(f"エラー: アイテム '{row['要件/項目名']}' の新規作成に失敗しました。")