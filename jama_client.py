"""
JAMA REST APIクライアント
OAuth認証とAPI操作を処理
"""

import requests
import time
import base64
import logging
from typing import Dict, List, Optional, Any
from urllib.parse import urljoin

logger = logging.getLogger(__name__)


class JAMAClient:
    """JAMA REST APIクライアント"""
    
    def __init__(self, config):
        """
        初期化
        
        Args:
            config: 設定オブジェクト
        """
        self.config = config
        self.base_url = config.base_url
        self.project_id = config.project_id
        self.proxies = config.proxies
        
        # OAuth用
        self.client_id = config.api_id
        self.client_secret = config.api_secret
        self.access_token = None
        self.token_expires = 0
        
        # セッション設定
        self.session = requests.Session()
        self.session.proxies = self.proxies
        
    def _get_access_token(self) -> str:
        """
        アクセストークンを取得（必要に応じて更新）
        
        Returns:
            アクセストークン
        """
        # トークンがまだ有効な場合はそのまま返す
        if self.access_token and time.time() < self.token_expires:
            return self.access_token
            
        logger.info("新しいアクセストークンを取得")
        
        # Basic認証用のヘッダー作成
        credentials = f"{self.client_id}:{self.client_secret}"
        auth_header = base64.b64encode(credentials.encode()).decode()
        
        # トークン取得
        token_url = urljoin(self.base_url, "/rest/oauth/token")
        headers = {
            "Authorization": f"Basic {auth_header}",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        data = {
            "grant_type": "client_credentials"
        }
        
        try:
            response = self.session.post(token_url, headers=headers, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data["access_token"]
            # 有効期限を設定（少し余裕を持たせる）
            self.token_expires = time.time() + token_data["expires_in"] - 60
            
            logger.info("アクセストークン取得成功")
            return self.access_token
            
        except requests.exceptions.RequestException as e:
            logger.error(f"アクセストークン取得失敗: {str(e)}")
            raise Exception(f"OAuth認証に失敗しました: {str(e)}")
            
    def _make_request(self, method: str, endpoint: str, 
                     params: Optional[Dict] = None, 
                     json_data: Optional[Dict] = None) -> Dict:
        """
        API リクエストを実行
        
        Args:
            method: HTTPメソッド
            endpoint: エンドポイント
            params: クエリパラメータ
            json_data: リクエストボディ
            
        Returns:
            レスポンスデータ
        """
        url = urljoin(self.base_url, endpoint)
        headers = {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json"
        }
        
        try:
            response = self.session.request(
                method=method,
                url=url,
                headers=headers,
                params=params,
                json=json_data
            )
            response.raise_for_status()
            
            # DELETEリクエストはレスポンスボディが空の場合がある
            if method == "DELETE" and not response.content:
                return {"status": "success"}
                
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"APIリクエスト失敗: {method} {url}, エラー: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"レスポンス: {e.response.text}")
            raise
            
    def get_project_info(self) -> Dict:
        """
        プロジェクト情報を取得
        
        Returns:
            プロジェクト情報
        """
        endpoint = f"/rest/v1/projects/{self.project_id}"
        response = self._make_request("GET", endpoint)
        return response.get("data", {})
        
    def get_all_items(self, max_depth: Optional[int] = None) -> List[Dict]:
        """
        プロジェクト内の全アイテムを取得
        
        Args:
            max_depth: 取得する最大階層数
            
        Returns:
            アイテムリスト
        """
        items = []
        start_at = 0
        max_results = 50
        
        while True:
            params = {
                "project": self.project_id,
                "startAt": start_at,
                "maxResults": max_results
            }
            
            response = self._make_request("GET", "/rest/v1/items", params=params)
            data = response.get("data", [])
            
            if not data:
                break
                
            # アイテムを処理
            for item in data:
                # 階層深度チェック
                if max_depth:
                    sequence = item.get("location", {}).get("sequence", "")
                    depth = len(sequence.split(".")) if sequence else 0
                    if depth > max_depth:
                        continue
                        
                items.append(self._process_item(item))
                
            # ページング
            meta = response.get("meta", {}).get("pageInfo", {})
            total = meta.get("totalResults", 0)
            
            start_at += max_results
            if start_at >= total:
                break
                
            logger.info(f"取得進捗: {min(start_at, total)}/{total}")
            
        return items
        
    def get_items_by_component(self, sequence: Optional[str] = None,
                              name: Optional[str] = None,
                              max_depth: Optional[int] = None) -> List[Dict]:
        """
        特定コンポーネント以下のアイテムを取得
        
        Args:
            sequence: コンポーネントのsequence
            name: コンポーネント名
            max_depth: 取得する最大階層数
            
        Returns:
            アイテムリスト
        """
        # まず全アイテムを取得
        all_items = self.get_all_items()
        
        # フィルタリング
        filtered_items = []
        target_found = False
        target_sequence = None
        
        for item in all_items:
            item_sequence = item.get("sequence", "")
            item_name = item.get("name", "")
            
            # ターゲットを探す
            if not target_found:
                if (sequence and item_sequence == sequence) or \
                   (name and item_name == name):
                    target_found = True
                    target_sequence = item_sequence
                    filtered_items.append(item)
                continue
                
            # ターゲット以下のアイテムをフィルタ
            if target_sequence and item_sequence.startswith(target_sequence + "."):
                # 相対的な深度チェック
                if max_depth:
                    target_depth = len(target_sequence.split("."))
                    item_depth = len(item_sequence.split("."))
                    if item_depth - target_depth > max_depth:
                        continue
                        
                filtered_items.append(item)
                
        return filtered_items
        
    def create_item(self, item_data: Dict) -> int:
        """
        新規アイテムを作成
        
        Args:
            item_data: アイテムデータ
            
        Returns:
            作成されたアイテムのID
        """
        # リクエストボディ作成
        request_data = {
            "project": self.project_id,
            "itemType": item_data.get("item_type_id", 1),  # デフォルトは1
            "childItemType": item_data.get("child_item_type_id", 1),
            "location": {
                "parent": {
                    "item": item_data.get("parent_id"),
                    "project": self.project_id
                }
            },
            "fields": self._prepare_fields(item_data)
        }
        
        response = self._make_request("POST", "/rest/v1/items", json_data=request_data)
        return response.get("id")
        
    def update_item(self, item_id: int, item_data: Dict) -> None:
        """
        既存アイテムを更新
        
        Args:
            item_id: アイテムID
            item_data: 更新データ
        """
        # リクエストボディ作成
        request_data = {
            "fields": self._prepare_fields(item_data)
        }
        
        # アイテムタイプなどは更新時は不要
        endpoint = f"/rest/v1/items/{item_id}"
        self._make_request("PUT", endpoint, json_data=request_data)
        
    def delete_item(self, item_id: int) -> None:
        """
        アイテムを削除（非アクティブ化）
        
        Args:
            item_id: アイテムID
        """
        endpoint = f"/rest/v1/items/{item_id}"
        self._make_request("DELETE", endpoint)
        
    def _process_item(self, item: Dict) -> Dict:
        """
        APIレスポンスのアイテムを内部形式に変換
        
        Args:
            item: APIレスポンスのアイテム
            
        Returns:
            内部形式のアイテム
        """
        fields = item.get("fields", {})
        location = item.get("location", {})
        
        return {
            "jama_id": item.get("id"),
            "sequence": location.get("sequence", ""),
            "parent_id": location.get("parent", {}).get("item"),
            "item_type_id": item.get("itemType"),
            "child_item_type_id": item.get("childItemType"),
            "name": fields.get("name", ""),
            "description": fields.get("description", ""),
            "assignee": fields.get("assignee", ""),
            "status": fields.get("status", ""),
            "tags": fields.get("tags", ""),
            "reason": fields.get("reason", ""),
            "preconditions": fields.get("preconditions", ""),
            "target_system": fields.get("target_system", ""),
            "created_date": item.get("createdDate"),
            "modified_date": item.get("modifiedDate"),
            "created_by": item.get("createdBy"),
            "modified_by": item.get("modifiedBy")
        }
        
    def _prepare_fields(self, item_data: Dict) -> Dict:
        """
        アイテムデータからfieldsオブジェクトを作成
        
        Args:
            item_data: アイテムデータ
            
        Returns:
            fieldsオブジェクト
        """
        fields = {}
        
        # フィールドマッピング
        field_mapping = {
            "name": "name",
            "description": "description",
            "assignee": "assignee",
            "status": "status",
            "tags": "tags",
            "reason": "reason",
            "preconditions": "preconditions",
            "target_system": "target_system"
        }
        
        for internal_key, api_key in field_mapping.items():
            if internal_key in item_data and item_data[internal_key] is not None:
                fields[api_key] = item_data[internal_key]
                
        return fields
