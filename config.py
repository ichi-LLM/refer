"""
設定ファイルの読み込みと管理
"""

import json
import os
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)


class Config:
    """設定管理クラス"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初期化
        
        Args:
            config_file: 設定ファイルのパス
        """
        self.config_file = config_file
        self.config_data = {}
        
        # デフォルト値
        self.base_url = "https://stargate.jamacloud.com"
        self.project_id = 124
        self.api_id = ""
        self.api_secret = ""
        self.proxies = {
            "http": "http://proxy1000.co.jp:15520",
            "https": "http://proxy1000.co.jp:15520"
        }
        
        # パフォーマンス設定
        self.column_width_check_rows = 100   # 列幅調整時にチェックする最大行数
        
        # デバッグ設定
        self.debug = False
        
        # アイテムタイプマッピング（階層レベル -> item_type_id）
        self.item_type_mapping = {
            "1": 30,   # Component
            "2": 30,   # Component
            "3": 30,   # Component
            "4": 30,   # Component
            "5": 30,   # Component
            "6": 30,   # Component
            "7": 30,   # Component
            "8": 30,   # Component
            "9": 31    # Set
            # 10, 11は要件名による動的判定のため除外
        }
        
        # アイテムタイプ名マッピング（ID -> 表示名）
        self.item_type_names = {
            30: "Component",
            31: "Set",
            266: "STD User Requirement",
            301: "STD Oneteam System Requirement"
        }
        
        # 10-11階層のデフォルトアイテムタイプ（判定できない場合）
        self.default_item_type_for_10_11 = 301
        
        # 設定ファイルを読み込み
        self._load_config()
        
    def _load_config(self) -> None:
        """設定ファイルを読み込み"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                    
                # 設定値を適用
                self.base_url = self.config_data.get('base_url', self.base_url)
                self.project_id = self.config_data.get('project_id', self.project_id)
                self.api_id = self.config_data.get('api_id', self.api_id)
                self.api_secret = self.config_data.get('api_secret', self.api_secret)
                
                # プロキシ設定
                if 'proxies' in self.config_data:
                    self.proxies = self.config_data['proxies']
                    
                # パフォーマンス設定
                self.column_width_check_rows = self.config_data.get('performance', {}).get('column_width_check_rows', 100)
                
                # デバッグ設定
                self.debug = self.config_data.get('debug', False)
                
                # アイテムタイプマッピング
                if 'item_type_mapping' in self.config_data:
                    # 設定ファイルの値で上書き（Noneも含めて）
                    for level, type_id in self.config_data['item_type_mapping'].items():
                        self.item_type_mapping[str(level)] = type_id
                
                # アイテムタイプ名マッピング
                if 'item_type_names' in self.config_data:
                    for type_id, name in self.config_data['item_type_names'].items():
                        self.item_type_names[int(type_id)] = name
                        
                # 10-11階層のデフォルト
                self.default_item_type_for_10_11 = self.config_data.get('default_item_type_for_10_11', 301)
                    
                logger.info(f"設定ファイルを読み込みました: {self.config_file}")
                
            except Exception as e:
                logger.error(f"設定ファイルの読み込みに失敗: {str(e)}")
                raise Exception(f"設定ファイルの読み込みエラー: {str(e)}")
        else:
            # 設定ファイルが存在しない場合、サンプルを作成
            self._create_sample_config()
            raise Exception(f"設定ファイルが見つかりません。{self.config_file} を作成してください。")
            
    def _create_sample_config(self) -> None:
        """サンプル設定ファイルを作成"""
        sample_config = {
            "base_url": "https://stargate.jamacloud.com",
            "project_id": 124,
            "api_id": "YOUR_API_ID_HERE",
            "api_secret": "YOUR_API_SECRET_HERE",
            "proxies": {
                "http": "http://proxy1000.co.jp:15520",
                "https": "http://proxy1000.co.jp:15520"
            },
            "performance": {
                "column_width_check_rows": 100,
                "_comment": "大量データ処理時のパフォーマンス調整"
            },
            "item_type_mapping": {
                "1": 30,
                "2": 30,
                "3": 30,
                "4": 30,
                "5": 30,
                "6": 30,
                "7": 30,
                "8": 30,
                "9": 31,
                "_comment": "階層レベルごとのアイテムタイプID。10,11は要件名により動的判定"
            },
            "item_type_names": {
                "30": "Component",
                "31": "Set",
                "266": "STD User Requirement",
                "301": "STD Oneteam System Requirement"
            },
            "default_item_type_for_10_11": 301,
            "debug": False
        }
        
        sample_file = self.config_file + ".sample"
        with open(sample_file, 'w', encoding='utf-8') as f:
            json.dump(sample_config, f, indent=2, ensure_ascii=False)
            
        logger.info(f"サンプル設定ファイルを作成しました: {sample_file}")
        
    def validate(self) -> bool:
        """
        設定値の検証
        
        Returns:
            検証成功時True
        """
        errors = []
        
        if not self.api_id:
            errors.append("api_id が設定されていません")
            
        if not self.api_secret:
            errors.append("api_secret が設定されていません")
            
        if not self.base_url:
            errors.append("base_url が設定されていません")
            
        if not self.project_id:
            errors.append("project_id が設定されていません")
            
        if errors:
            for error in errors:
                logger.error(f"設定エラー: {error}")
            return False
            
        return True
        
    def get(self, key: str, default: Any = None) -> Any:
        """
        設定値を取得
        
        Args:
            key: 設定キー
            default: デフォルト値
            
        Returns:
            設定値
        """
        return self.config_data.get(key, default)
        
    def get_item_type_for_level(self, level: int) -> Optional[int]:
        """
        階層レベルに対応するアイテムタイプIDを取得
        
        Args:
            level: 階層レベル（1-11）
            
        Returns:
            アイテムタイプID、未定義の場合はNone
        """
        return self.item_type_mapping.get(str(level))
        
    def get_item_type_name(self, type_id: int) -> str:
        """
        アイテムタイプIDに対応する表示名を取得
        
        Args:
            type_id: アイテムタイプID
            
        Returns:
            表示名、未定義の場合は"Unknown"
        """
        return self.item_type_names.get(type_id, "Unknown")
        
    def is_item_type_defined_for_level(self, level: int) -> bool:
        """
        指定階層レベルのアイテムタイプIDが定義されているか確認
        
        Args:
            level: 階層レベル（1-11）
            
        Returns:
            定義されている場合True
        """
        return self.item_type_mapping.get(str(level)) is not None
