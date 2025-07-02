"""
JAMA REST APIクライアント
OAuth認証とAPI操作を処理
"""

import requests
import time
import base64
import logging
import json
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
        self.sample_mode = False
        self._item_count = 0
        self._sysp_found = False
        
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
            
    def set_sample_mode(self, enabled: bool) -> None:
        """
        サンプルモードを設定
        
        Args:
            enabled: サンプルモードの有効/無効
        """
        self.sample_mode = enabled
        if enabled:
            logger.info("JAMAクライアントのサンプルモードが有効になりました")
        
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
        
    def get_sample_items(self, count: int = 100) -> List[Dict]:
        """
        サンプルアイテムを取得（デバッグ・構造調査用）
        
        Args:
            count: 取得する件数
            
        Returns:
            アイテムリスト
        """
        logger.info(f"サンプルアイテム取得開始: {count}件")
        items = []
        
        # 一度のリクエストで取得
        params = {
            "project": self.project_id,
            "startAt": 0,
            "maxResults": min(count, 50)  # APIの制限に合わせる
        }
        
        # 必要に応じて複数回リクエスト
        while len(items) < count:
            response = self._make_request("GET", "/rest/v1/items", params=params)
            data = response.get("data", [])
            
            if not data:
                break
                
            # アイテムを処理
            for item in data:
                items.append(self._process_item(item))
                if len(items) >= count:
                    break
                    
            # 次のページ
            params["startAt"] += params["maxResults"]
            
            # 残りのアイテム数に応じてmaxResultsを調整
            remaining = count - len(items)
            if remaining < params["maxResults"]:
                params["maxResults"] = remaining
                
        logger.info(f"サンプルアイテム取得完了: {len(items)}件")
        
        # SYSPアイテムが見つかったか報告
        sysp_items = [item for item in items if "SYSP" in item.get("name", "")]
        logger.info(f"SYSPアイテム数: {len(sysp_items)}件")
        
        return items
        
    def get_all_items(self, max_depth: Optional[int] = None, max_count: Optional[int] = None) -> List[Dict]:
        """
        プロジェクト内の全アイテムを取得
        
        Args:
            max_depth: 取得する最大階層数
            max_count: 取得する最大件数
            
        Returns:
            アイテムリスト
        """
        items = []
        start_at = 0
        max_results = 50
        
        # カウンターリセット
        self._item_count = 0
        self._sysp_found = False
        
        # 最初に総数を取得
        logger.info("アイテム総数を確認中...")
        initial_response = self._make_request("GET", "/rest/v1/items", 
                                            params={"project": self.project_id, "startAt": 0, "maxResults": 1})
        total_items = initial_response.get("meta", {}).get("pageInfo", {}).get("totalResults", 0)
        logger.info(f"取得予定アイテム総数: {total_items}件")
        
        # 進捗表示用の上限値を決定
        progress_total = total_items
        if max_count:
            logger.info(f"最大取得件数制限: {max_count}件")
            progress_total = min(total_items, max_count)
        
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
                
                # 最大件数チェック
                if max_count and len(items) >= max_count:
                    logger.info(f"最大取得件数（{max_count}件）に到達しました")
                    logger.info(f"取得進捗: {max_count}/{progress_total} (100.0%)")
                    return items
                
            # ページング
            meta = response.get("meta", {}).get("pageInfo", {})
            total = meta.get("totalResults", 0)
            
            start_at += max_results
            if start_at >= total:
                break
            
            # 進捗表示（制限値を考慮）
            current_progress = min(start_at, progress_total)
            logger.info(f"取得進捗: {current_progress}/{progress_total} ({current_progress/progress_total*100:.1f}%)")
            
        logger.info(f"アイテム取得完了: {len(items)}件")
        return items
        
    def get_items_by_component(self, sequence: Optional[str] = None,
                              name: Optional[str] = None,
                              max_depth: Optional[int] = None,
                              max_count: Optional[int] = None) -> List[Dict]:
        """
        特定コンポーネント以下のアイテムを取得
        
        Args:
            sequence: コンポーネントのsequence
            name: コンポーネント名
            max_depth: 取得する最大階層数
            max_count: 取得する最大件数
            
        Returns:
            アイテムリスト
        """
        # まず全アイテムを取得（max_countは渡さない。フィルタリング後に適用）
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
                    
                    # 最大件数チェック
                    if max_count and len(filtered_items) >= max_count:
                        return filtered_items
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
                
                # 最大件数チェック
                if max_count and len(filtered_items) >= max_count:
                    logger.info(f"最大取得件数（{max_count}件）に到達しました")
                    return filtered_items
                
        return filtered_items
        
    def _find_field_value(self, fields: Dict, field_patterns: List[str]) -> Any:
        """
        パターンに一致するフィールドを検索
        
        Args:
            fields: フィールド辞書
            field_patterns: 検索パターンのリスト（優先順位順）
            
        Returns:
            見つかった値、または空文字列
        """
        # 完全一致を優先
        for pattern in field_patterns:
            if pattern in fields:
                return self._convert_field_value(fields[pattern])
        
        # 部分一致（大文字小文字を無視）
        for pattern in field_patterns:
            pattern_lower = pattern.lower()
            for key in fields.keys():
                if pattern_lower in key.lower():
                    return self._convert_field_value(fields[key])
        
        return ""
    
    def _convert_field_value(self, value: Any) -> str:
        """
        フィールド値を文字列に変換
        
        Args:
            value: 変換する値
            
        Returns:
            文字列化された値
        """
        if value is None:
            return ""
        elif isinstance(value, str):
            return value
        elif isinstance(value, (int, float)):
            return str(value)
        elif isinstance(value, list):
            # 配列はカンマ区切りの文字列に変換
            return ", ".join(str(v) for v in value)
        elif isinstance(value, dict):
            # 辞書は文字列化
            return str(value)
        else:
            return str(value)
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
        self._item_count += 1
        
        # サンプルモードで詳細出力（デバッグモードでは簡潔な出力のみ）
        if self.sample_mode:
            # 最初の3件とSYSPアイテムは必ず出力
            name = item.get("fields", {}).get("name", "")
            is_sysp = "SYSP" in name
            
            if self._item_count <= 3 or is_sysp:
                logger.info(f"\n=== Item #{self._item_count} Structure ===")
                logger.info(f"Item ID: {item.get('id')}")
                logger.info(f"Item Type: {item.get('itemType')}")
                logger.info(f"Item Name: {name}")
                logger.info(f"Is SYSP: {is_sysp}")
                
                # item直下のキーを確認
                logger.info(f"Item top-level keys: {list(item.keys())}")
                
                # 各トップレベルキーの内容を確認（fieldsとlocation以外）
                for key in item.keys():
                    if key not in ['fields', 'location']:
                        value = item[key]
                        if isinstance(value, dict):
                            logger.info(f"{key}: {list(value.keys()) if value else 'empty dict'}")
                        elif isinstance(value, (list, str, int, bool)):
                            logger.info(f"{key}: {value}")
                
                # fieldsの内容
                fields = item.get("fields", {})
                logger.info(f"Fields keys: {list(fields.keys())}")
                
                # SYSPの場合は全フィールドを出力
                if is_sysp:
                    logger.info("=== SYSP Fields Content ===")
                    for key, value in fields.items():
                        if isinstance(value, str) and len(value) > 100:
                            logger.info(f"  {key}: {value[:100]}...")
                        else:
                            logger.info(f"  {key}: {value}")
                    
                    # 標準フィールドの可能性がある場所を探す
                    if "status" in item:
                        logger.info(f"Found status at item level: {item['status']}")
                    if "assigned" in item:
                        logger.info(f"Found assigned at item level: {item['assigned']}")
                    if "assignedUser" in item:
                        logger.info(f"Found assignedUser at item level: {item['assignedUser']}")
                
                logger.info("=== End ===\n")
                
                if is_sysp and not self._sysp_found:
                    self._sysp_found = True
                    # 完全なJSON構造を出力
                    logger.info("\n=== First SYSP Item Full JSON ===")
                    logger.info(json.dumps(item, indent=2, ensure_ascii=False))
                    logger.info("=== End Full JSON ===\n")
        
        # データ抽出
        fields = item.get("fields", {})
        location = item.get("location", {})
        name = fields.get("name", "")
        
        # 基本情報を構築
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
        
        # フィールドの柔軟な検索
        # status - システムフィールド（item直下の可能性もある）
        if "status" in item:
            result["status"] = self._convert_field_value(item["status"])
        else:
            result["status"] = self._find_field_value(fields, ["status_STD", "status", "Status"])
        
        # assigned/assignee - ユーザーフィールド（item直下の可能性もある）
        if "assigned" in item:
            result["assignee"] = self._convert_field_value(item["assigned"])
        elif "assignedUser" in item:
            result["assignee"] = self._convert_field_value(item["assignedUser"])
        else:
            result["assignee"] = self._find_field_value(fields, ["assigned", "assignedUser", "assignee", "Assignee"])
        
        # tags - 特殊なフィールド（item直下の可能性）
        if "tags" in item:
            result["tags"] = self._convert_field_value(item["tags"])
        else:
            result["tags"] = self._find_field_value(fields, ["tags", "Tags"])
        
        # カスタムフィールド（fields内）
        # Reason
        result["reason"] = self._find_field_value(fields, ["reason", "Reason"])
        
        # Preconditions
        result["preconditions"] = self._find_field_value(fields, ["preconditions", "Preconditions", "precondition", "Precondition"])
        
        # Target System
        result["target_system"] = self._find_field_value(fields, ["target_system", "targetSystem", "Target_system", "TargetSystem", "target", "Target"])
        
        # SYSPアイテムの場合、追加のデバッグ情報
        if "SYSP" in name and self.debug_mode:
            logger.info(f"SYSP Item {item.get('id')} extracted fields:")
            logger.info(f"  status: '{result['status']}'")
            logger.info(f"  assignee: '{result['assignee']}'")
            logger.info(f"  tags: '{result['tags']}'")
            logger.info(f"  reason: '{result['reason']}'")
            logger.info(f"  preconditions: '{result['preconditions']}'")
            logger.info(f"  target_system: '{result['target_system']}'")
        
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