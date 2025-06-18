import requests
import base64
import time

class JamaClient:
    """
    JAMA REST APIとの通信を管理するクライアントクラス。
    """
    def __init__(self, base_url, client_id, client_secret, proxies):
        self.base_url = base_url.rstrip('/')
        self.client_id = client_id
        self.client_secret = client_secret
        self.proxies = proxies if any(proxies.values()) else None
        self.access_token = None
        self.token_expires_at = 0

    def _get_access_token(self):
        """OAuth 2.0のアクセストークンを取得または更新する。"""
        if self.access_token and time.time() < self.token_expires_at:
            return

        print("JAMAから新しいアクセストークンを取得しています...")
        url = f"{self.base_url}/rest/oauth/token"
        auth_str = f"{self.client_id}:{self.client_secret}"
        auth_b64 = base64.b64encode(auth_str.encode()).decode()
        
        headers = {
            'Authorization': f'Basic {auth_b64}',
            'Content-Type': 'application/x-www-form-urlencoded',
        }
        data = {'grant_type': 'client_credentials'}

        try:
            response = requests.post(url, headers=headers, data=data, proxies=self.proxies)
            response.raise_for_status()
            token_data = response.json()
            self.access_token = token_data['access_token']
            # トークンの有効期限に5秒のバッファを持たせる
            self.token_expires_at = time.time() + token_data['expires_in'] - 5
            print("アクセストークンの取得に成功しました。")
        except requests.exceptions.RequestException as e:
            print(f"エラー: アクセストークンの取得に失敗しました。{e}")
            raise

    def _make_request(self, method, endpoint, params=None, json=None):
        """APIリクエストを送信する共通メソッド。"""
        self._get_access_token()
        url = f"{self.base_url}/rest/v1{endpoint}"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        try:
            response = requests.request(method, url, headers=headers, params=params, json=json, proxies=self.proxies)
            response.raise_for_status()
            # DELETEのように中身が空のレスポンスもある
            if response.status_code == 204 or not response.content:
                return None
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"エラー: APIリクエストに失敗しました ({method} {url}) - {e}")
            if e.response:
                print(f"レスポンス内容: {e.response.text}")
            raise

    def get_all_paginated_data(self, endpoint, params):
        """ページネーションを考慮してすべてのデータを取得する。"""
        all_data = []
        start_at = 0
        max_results = 50 # APIの最大値
        params['maxResults'] = max_results
        
        while True:
            params['startAt'] = start_at
            print(f"{endpoint} から {start_at+1} 件目以降のデータを取得中...")
            response = self._make_request('GET', endpoint, params=params)
            
            if not response or 'data' not in response or not response['data']:
                break
                
            data = response['data']
            all_data.extend(data)
            
            page_info = response['meta']['pageInfo']
            if (page_info['startIndex'] + page_info['resultCount']) >= page_info['totalResults']:
                break
                
            start_at += max_results
            
        return all_data

    def get_items_in_project(self, project_id, component_id=None):
        """プロジェクト内のすべてのアイテムを取得する。"""
        if component_id:
             # 特定コンポーネント配下を取得
            return self.get_item_children(component_id)
        else:
            # プロジェクト全体を取得
            params = {'project': project_id, 'include': 'location,fields'}
            return self.get_all_paginated_data('/items', params=params)

    def get_item_children(self, item_id):
        """指定したアイテムの子アイテムをすべて取得する。"""
        params = {'include': 'location,fields'}
        return self.get_all_paginated_data(f'/items/{item_id}/children', params)

    def create_item(self, data):
        """新しいアイテムを作成する。"""
        print(f"アイテムを作成中: {data['fields'].get('name', '名称未設定')}")
        return self._make_request('POST', '/items', json=data)

    def update_item(self, item_id, data):
        """既存のアイテムを更新する。"""
        print(f"アイテムを更新中 (ID: {item_id})")
        return self._make_request('PUT', f'/items/{item_id}', json=data)

    def delete_item(self, item_id):
        """アイテムを削除する。"""
        print(f"アイテムを削除中 (ID: {item_id})")
        return self._make_request('DELETE', f'/items/{item_id}')
        
    def get_itemtypes(self):
        """すべてのアイテムタイプを取得する。"""
        return self.get_all_paginated_data('/itemtypes', params={})

    def get_users(self):
        """すべてのユーザーを取得する。"""
        return self.get_all_paginated_data('/users', params={})