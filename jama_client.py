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
        
        # デバッグモード
        self.debug_mode = False
        self._debug_printed = False
        self._sysp_debug_printed = False
        
        # セッション設定
        self.session = requests.Session()
        self.session.proxies = self.proxies
        
    def set_debug_mode(self, enabled: bool) -> None:
        """
        デバッグモードを設定
        
        Args:
            enabled: デバッグモードの有効/無効
        """
        self.debug_mode = enabled
        if enabled:
            logger.info("JAMAクライアントのデバッグモードが有効になりました")
        
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
        
        # 最初に総数を取得
        logger.info("アイテム総数を確認中...")
        initial_response = self._make_request("GET", "/rest/v1/items", 
                                            params={"project": self.project_id, "startAt": 0, "maxResults": 1})
        total_items = initial_response.get("meta", {}).get("pageInfo", {}).get("totalResults", 0)
        logger.info(f"取得予定アイテム総数: {total_items}件")
        
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
                
            logger.info(f"取得進捗: {min(start_at, total)}/{total} ({min(start_at, total)/total*100:.1f}%)")
            
        logger.info(f"アイテム取得完了: {len(items)}件")
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
        # 現在のアイテム情報を取得して親情報を保持
        endpoint = f"/rest/v1/items/{item_id}"
        current_item = self._make_request("GET", endpoint)
        
        # デバッグモード時のみ詳細ログを出力
        if self.debug_mode:
            logger.debug(f"Current item response keys: {current_item.keys()}")
        
        # データ構造に応じて親IDを取得
        if "data" in current_item:
            location = current_item["data"].get("location", {})
            if self.debug_mode:
                logger.debug(f"Using data.location path")
        else:
            location = current_item.get("location", {})
            if self.debug_mode:
                logger.debug(f"Using direct location path")
        
        parent_id = location.get("parent", {}).get("item")
        
        # デバッグモード時は詳細情報、通常時は簡潔な情報
        if self.debug_mode:
            logger.debug(f"Updating item {item_id}, parent_id: {parent_id}, location: {location}")
        else:
            logger.info(f"アイテム更新中: ID={item_id}")
        
        # リクエストボディ作成（親情報に応じて分岐）
        if parent_id:
            # 親がアイテムの場合
            if self.debug_mode:
                logger.debug(f"Parent is an item: {parent_id}")
            request_data = {
                "project": self.project_id,
                "location": {
                    "parent": {
                        "item": parent_id
                    }
                },
                "fields": self._prepare_fields(item_data)
            }
        else:
            # 親がプロジェクト（ルート）の場合
            if self.debug_mode:
                logger.debug(f"Parent is project root")
            request_data = {
                "project": self.project_id,
                "location": {
                    "parent": {
                        "project": self.project_id
                    }
                },
                "fields": self._prepare_fields(item_data)
            }
        
        # デバッグモード時のみリクエストデータを表示
        if self.debug_mode:
            logger.debug(f"Request data: {request_data}")
        
        # 更新を実行
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
        
        # デバッグモード：最初の1件だけfieldsの中身を出力
        if self.debug_mode and not self._debug_printed:
            logger.info("=== DEBUG: First Item structure ===")
            logger.info(f"Item ID: {item.get('id')}")
            logger.info(f"Item Type: {item.get('itemType')}")
            logger.info(f"Available fields: {list(fields.keys())}")
            logger.info("=== DEBUG: Fields content sample ===")
            for key, value in fields.items():
                # 長いテキストは省略
                if isinstance(value, str) and len(value) > 100:
                    logger.info(f"  {key}: {value[:100]}...")
                else:
                    logger.info(f"  {key}: {value}")
            logger.info("=== DEBUG: End ===")
            self._debug_printed = True
            
        # デバッグモード：最初のSYSPアイテムのフィールドを出力
        name = fields.get("name", "")
        if self.debug_mode and "SYSP" in name and not self._sysp_debug_printed:
            logger.info("\n=== DEBUG: SYSP Item structure ===")
            logger.info(f"SYSP Item ID: {item.get('id')}")
            logger.info(f"SYSP Item Name: {name}")
            logger.info(f"SYSP Available fields: {list(fields.keys())}")
            logger.info("=== DEBUG: SYSP Fields content ===")
            for key, value in fields.items():
                # 長いテキストは省略
                if isinstance(value, str) and len(value) > 100:
                    logger.info(f"  {key}: {value[:100]}...")
                else:
                    logger.info(f"  {key}: {value}")
            logger.info("=== DEBUG: SYSP End ===\n")
            self._sysp_debug_printed = True
        
        # SYSPアイテムの場合、すべてのフィールドを取得
        result = {
            "jama_id": item.get("id"),
            "sequence": location.get("sequence", ""),
            "parent_id": location.get("parent", {}).get("item"),
            "item_type_id": item.get("itemType"),
            "child_item_type_id": item.get("childItemType"),
            "name": name,
            "description": fields.get("description", ""),
            "created_date": item.get("createdDate"),
            "modified_date": item.get("modifiedDate"),
            "created_by": item.get("createdBy"),
            "modified_by": item.get("modifiedBy")
        }
        
        # SYSPアイテムの場合、すべてのフィールドを含める
        if "SYSP" in name:
            # 既知のフィールドを試す
            possible_fields = [
                "assignee", "status", "tags", "reason", "preconditions", "target_system",
                "Assignee", "Status", "Tags", "Reason", "Preconditions", "Target_system",
                "tag", "Tag", "reasons", "Reasons", "precondition", "Precondition",
                "targetSystem", "target", "Target"
            ]
            
            for field_name in possible_fields:
                if field_name in fields:
                    # 正規化したキー名で保存
                    normalized_key = field_name.lower().replace("_", "")
                    if "assignee" in normalized_key:
                        result["assignee"] = fields[field_name]
                    elif "status" in normalized_key:
                        result["status"] = fields[field_name]
                    elif "tag" in normalized_key:
                        result["tags"] = fields[field_name]
                    elif "reason" in normalized_key:
                        result["reason"] = fields[field_name]
                    elif "precondition" in normalized_key:
                        result["preconditions"] = fields[field_name]
                    elif "target" in normalized_key:
                        result["target_system"] = fields[field_name]
            
            # その他のフィールドもすべて含める（デバッグ用）
            for key, value in fields.items():
                if key not in ["name", "description", "documentKey", "globalID"]:
                    # まだ処理していないフィールドを追加
                    if key not in [k for k in result.keys()]:
                        result[f"custom_{key}"] = value
        else:
            # 非SYSPアイテムは従来通り
            result["assignee"] = fields.get("assignee", "")
            result["status"] = fields.get("status", "")
            result["tags"] = fields.get("tags", "")
            result["reason"] = fields.get("reason", "")
            result["preconditions"] = fields.get("preconditions", "")
            result["target_system"] = fields.get("target_system", "")
        
        return result
        
    def _prepare_fields(self, item_data: Dict) -> Dict:
        """
        アイテムデータからfieldsオブジェクトを作成
        空欄のフィールドは更新対象から除外する
        
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
            # 値が存在し、None でも空文字でもない場合のみ更新対象に含める
            if internal_key in item_data and item_data[internal_key] is not None and item_data[internal_key] != "":
                fields[api_key] = item_data[internal_key]
                
        return fields