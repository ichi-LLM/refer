"""
JAMA REST APIクライアント
OAuth認証とAPI操作を効率的に処理
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
        self.config = config
        self.base_url = config.base_url
        self.project_id = config.project_id
        self.proxies = config.proxies
        self.client_id = config.api_id
        self.client_secret = config.api_secret
        self.access_token = None
        self.token_expires = 0
        self.debug_mode = False
        self._debug_printed = False
        self.session = requests.Session()
        self.session.proxies = self.proxies

    def set_debug_mode(self, enabled: bool) -> None:
        self.debug_mode = enabled
        if enabled:
            logger.info("JAMAクライアントのデバッグモードが有効になりました")

    def _get_access_token(self) -> str:
        if self.access_token and time.time() < self.token_expires:
            return self.access_token
        logger.info("新しいアクセストークンを取得")
        credentials = f"{self.client_id}:{self.client_secret}"
        auth_header = base64.b64encode(credentials.encode()).decode()
        token_url = urljoin(self.base_url, "/rest/oauth/token")
        headers = {
            "Authorization": f"Basic {auth_header}",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        data = {"grant_type": "client_credentials"}
        try:
            response = self.session.post(token_url, headers=headers, data=data)
            response.raise_for_status()
            token_data = response.json()
            self.access_token = token_data["access_token"]
            self.token_expires = time.time() + token_data["expires_in"] - 60
            logger.info("アクセストークン取得成功")
            return self.access_token
        except requests.exceptions.RequestException as e:
            logger.error(f"アクセストークン取得失敗: {str(e)}")
            raise Exception(f"OAuth認証に失敗しました: {str(e)}")

    def _make_request(self, method: str, endpoint: str,
                     params: Optional[Dict] = None,
                     json_data: Optional[Dict] = None) -> Dict:
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
            if method.upper() == "DELETE" and not response.content:
                return {"status": "success"}
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"APIリクエスト失敗: {method} {url}, エラー: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"レスポンス: {e.response.text}")
            raise

    def get_project_info(self) -> Dict:
        endpoint = f"/rest/v1/projects/{self.project_id}"
        response = self._make_request("GET", endpoint)
        return response.get("data", {})

    def get_all_items(self, max_depth: Optional[int] = None) -> List[Dict]:
        """プロジェクト内の全アイテムを取得（全件取得モード）"""
        logger.info("プロジェクト全体のアイテムを取得開始")
        all_items = []
        start_at = 0
        max_results = 50
        
        while True:
            params = {
                "project": self.project_id,
                "startAt": start_at,
                "maxResults": max_results,
                "include": "fields,location" # 必須フィールドを要求
            }
            response = self._make_request("GET", "/rest/v1/items", params=params)
            data = response.get("data", [])
            if not data:
                break
            
            for item in data:
                if max_depth:
                    sequence = item.get("location", {}).get("sequence", "")
                    depth = len(sequence.split(".")) if sequence else 0
                    if depth > max_depth:
                        continue
                all_items.append(self._process_item(item))
            
            total = response.get("meta", {}).get("pageInfo", {}).get("totalResults", 0)
            start_at += len(data)
            logger.info(f"取得進捗: {min(start_at, total)}/{total} ({min(start_at, total)/total*100:.1f}%)")
            if start_at >= total:
                break
        
        logger.info(f"アイテム取得完了: {len(all_items)}件")
        return all_items

    def get_items_by_component(self, sequence: Optional[str] = None,
                              name: Optional[str] = None,
                              max_depth: Optional[int] = None) -> List[Dict]:
        """特定コンポーネント以下のアイテムを効率的に取得"""
        logger.info(f"コンポーネント指定での取得開始: sequence='{sequence}', name='{name}'")
        
        # ステップ1: 起点となるアイテムのIDを特定
        start_item = self._find_item(sequence=sequence, name=name)
        if not start_item:
            logger.warning(f"指定されたコンポーネントが見つかりませんでした。")
            return []
        
        start_item_id = start_item['id']
        logger.info(f"起点アイテムを発見: ID={start_item_id}, Name='{start_item['fields']['name']}'")
        
        # ステップ2: 起点アイテムとその子孫をすべて取得
        all_descendants = []
        # まず起点アイテム自身をリストに追加
        all_descendants.append(self._process_item(start_item))
        
        # 次にすべての子孫を再帰的に取得
        self._get_descendants(start_item_id, all_descendants)

        # 深度（max_depth）によるフィルタリング
        if max_depth is not None:
            logger.info(f"最大深度 {max_depth} でフィルタリング")
            start_depth = len(start_item.get("location", {}).get("sequence", "").split('.'))
            
            filtered_list = []
            for item in all_descendants:
                item_depth = len(item.get("sequence", "").split('.'))
                if item_depth - start_depth < max_depth:
                    filtered_list.append(item)
            return filtered_list
            
        return all_descendants

    def _find_item(self, sequence: Optional[str] = None, name: Optional[str] = None) -> Optional[Dict]:
        """指定されたsequenceまたはnameに一致するアイテムを1件見つける"""
        # sequenceが指定されている場合、優先的に処理
        if sequence:
            parts = sequence.split('.')
            current_parent_id = None
            found_item = None
            
            for i, part_str in enumerate(parts):
                logger.info(f"Sequence '{'.'.join(parts[:i+1])}' を検索中...")
                target_sort_order = int(part_str)
                
                # 親IDを指定して子アイテムを取得
                params = {
                    "project": self.project_id,
                    "include": "fields,location"
                }
                if current_parent_id:
                    params['parent'] = current_parent_id
                
                endpoint = "/rest/v1/items"
                if not current_parent_id:
                     endpoint = f"/rest/v1/projects/{self.project_id}/rootitems"

                children = self._make_request("GET", endpoint, params=params).get("data", [])
                
                found_child = None
                for child in children:
                    if child.get("location", {}).get("sortOrder") == target_sort_order:
                        found_child = child
                        break
                
                if not found_child:
                    logger.warning(f"Sequence '{sequence}' の途中でアイテムが見つかりませんでした。")
                    return None
                
                current_parent_id = found_child['id']
                found_item = found_child

            return found_item

        # nameが指定されている場合（全件取得よりは効率的な方法で探す）
        if name:
            logger.info(f"名前 '{name}' でアイテムを検索中...")
            # まずルートアイテムから検索
            root_items = self._make_request("GET", f"/rest/v1/projects/{self.project_id}/rootitems", params={"include":"fields,location"}).get("data", [])
            for item in root_items:
                if item.get("fields", {}).get("name") == name:
                    return item
            
            # TODO: さらに深い階層を名前で検索する場合は、より複雑なロジックが必要。
            # 現状はルートから見つからない場合はNoneを返す。
            logger.warning(f"ルート直下で名前 '{name}' のアイテムは見つかりませんでした。")

        return None

    def _get_descendants(self, parent_id: int, all_items_list: List[Dict]):
        """指定された親ID配下の子孫アイテムを再帰的に全て取得する"""
        start_at = 0
        max_results = 50
        while True:
            params = {
                "parent": parent_id,
                "startAt": start_at,
                "maxResults": max_results,
                "include": "fields,location"
            }
            response = self._make_request("GET", "/rest/v1/items", params=params)
            children = response.get("data", [])
            if not children:
                break
            
            for child in children:
                all_items_list.append(self._process_item(child))
                # さらにその子孫を取得
                self._get_descendants(child['id'], all_items_list)
                
            total = response.get("meta", {}).get("pageInfo", {}).get("totalResults", 0)
            start_at += len(children)
            if start_at >= total:
                break

    def create_item(self, item_data: Dict) -> int:
        """新規アイテムを作成"""
        # (実装は変更なし)
        pass
        
    def update_item(self, item_id: int, item_data: Dict) -> None:
        """既存アイテムを更新"""
        # (実装は変更なし)
        pass
        
    def delete_item(self, item_id: int) -> None:
        """アイテムを削除"""
        # (実装は変更なし)
        pass

    def _process_item(self, item: Dict) -> Dict:
        """APIレスポンスのアイテムを内部形式に変換"""
        fields = item.get("fields", {})
        location = item.get("location", {})
        
        if self.debug_mode and not self._debug_printed:
            logger.info("=== DEBUG: Item structure ===")
            logger.info(f"Item ID: {item.get('id')}")
            logger.info(f"Available fields: {list(fields.keys())}")
            logger.info("=== DEBUG: Fields content sample ===")
            for key, value in fields.items():
                sample = str(value)
                if len(sample) > 100:
                    logger.info(f"  {key}: {sample[:100]}...")
                else:
                    logger.info(f"  {key}: {sample}")
            logger.info("=== DEBUG: End ===")
            self._debug_printed = True
        
        # `item.get("fields", {})` から直接値を取得するように修正
        return {
            "jama_id": item.get("id"),
            "sequence": location.get("sequence", ""),
            "parent_id": location.get("parent", {}).get("item"),
            "item_type_id": item.get("itemType"),
            "name": fields.get("name", ""),
            "description": fields.get("description", ""),
            # その他のカスタムフィールドも同様に取得
            "reason": fields.get("reason", ""),
            "tags": fields.get("tags", ""),
            "preconditions": fields.get("preconditions", ""),
            # 他のフィールドもここに追加
        }
        
    def _prepare_fields(self, item_data: Dict) -> Dict:
        """アイテムデータからfieldsオブジェクトを作成"""
        # (実装は変更なし)
        pass