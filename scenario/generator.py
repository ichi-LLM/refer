"""
異常系シナリオ生成のメイン処理
JAMAの要件データから異常系トリガーを読み込み、ベースシナリオに挿入して新しいシナリオを生成
"""

import json
import os
import yaml
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any
import copy

logger = logging.getLogger(__name__)


class ScenarioGenerator:
    """異常系シナリオ生成クラス"""
    
    def __init__(self, config_path: str = None):
        """
        初期化
        
        Args:
            config_path: 設定ファイルのパス
        """
        self.config_path = config_path
        self.scenario_mappings = {}
        self.base_scenarios_dir = Path(__file__).parent / "base_scenarios"
        self.config_dir = Path(__file__).parent / "config"
        self.output_dir = Path(__file__).parent / "output"
        
        # 設定ファイル読み込み
        self._load_scenario_mappings()
        
    def _load_scenario_mappings(self) -> None:
        """scenario_mappings.yamlを読み込み"""
        mapping_file = self.config_dir / "scenario_mappings.yaml"
        
        if not mapping_file.exists():
            logger.warning(f"設定ファイルが見つかりません: {mapping_file}")
            # デフォルト設定を作成
            self._create_default_mappings(mapping_file)
            
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                self.scenario_mappings = yaml.safe_load(f)
                logger.info(f"設定ファイルを読み込みました: {mapping_file}")
        except Exception as e:
            logger.error(f"設定ファイルの読み込みに失敗: {str(e)}")
            raise
            
    def _create_default_mappings(self, mapping_file: Path) -> None:
        """デフォルトのscenario_mappings.yamlを作成"""
        default_mappings = {
            "common_settings": {
                "injection_points": {
                    "開始前": {
                        "description": "機能開始前のチェック",
                        "default_delay": 3.0
                    },
                    "自動制御中": {
                        "description": "自動制御実行中の異常",
                        "default_delay": 0.5
                    }
                }
            },
            "scenarios": {
                "AP_出庫": {
                    "base_file": "AP_basic_sample1.json",
                    "injection_points": {
                        "開始前": {
                            "after_event": 2
                        },
                        "自動制御中": {
                            "after_event": 10
                        }
                    }
                }
            }
        }
        
        # ディレクトリ作成
        mapping_file.parent.mkdir(parents=True, exist_ok=True)
        
        # ファイル書き込み
        with open(mapping_file, 'w', encoding='utf-8') as f:
            yaml.dump(default_mappings, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            
        logger.info(f"デフォルト設定ファイルを作成しました: {mapping_file}")
        self.scenario_mappings = default_mappings
        
    def generate_scenarios(self, base_scenario_name: str, trigger_requirements: List[Dict], 
                         output_dir: Optional[str] = None) -> List[str]:
        """
        トリガー要件を基に異常系シナリオを生成
        
        Args:
            base_scenario_name: ベースシナリオ名（例: "AP_出庫"）
            trigger_requirements: トリガー要件リスト
            output_dir: 出力ディレクトリ（省略時はデフォルト）
            
        Returns:
            生成したシナリオファイルパスのリスト
        """
        logger.info(f"シナリオ生成開始: ベース={base_scenario_name}, トリガー数={len(trigger_requirements)}")
        
        # 出力ディレクトリ設定
        if output_dir:
            self.output_dir = Path(output_dir)
        
        # ベースシナリオの設定を取得
        if base_scenario_name not in self.scenario_mappings.get("scenarios", {}):
            raise ValueError(f"シナリオ '{base_scenario_name}' の設定が見つかりません")
            
        scenario_config = self.scenario_mappings["scenarios"][base_scenario_name]
        
        # ベースシナリオファイルを読み込み
        base_scenario = self._load_base_scenario(base_scenario_name, scenario_config["base_file"])
        
        # 共通設定を取得
        common_settings = self.scenario_mappings.get("common_settings", {})
        injection_points = scenario_config.get("injection_points", {})
        
        # 生成したファイルパスのリスト
        generated_files = []
        
        # 出力ディレクトリ作成（日付_シナリオ名）
        date_str = datetime.now().strftime("%Y-%m-%d")
        scenario_output_dir = self.output_dir / f"{date_str}_{base_scenario_name}"
        scenario_output_dir.mkdir(parents=True, exist_ok=True)
        
        # 各トリガー × 各挿入位置でシナリオ生成
        total_count = len(trigger_requirements) * len(injection_points)
        current_count = 0
        
        print(f"\n=== シナリオ生成開始 ===")
        print(f"ベースシナリオ: {base_scenario_name}")
        print(f"トリガー要件: {len(trigger_requirements)}件")
        print(f"挿入位置: {len(injection_points)}箇所（{', '.join(injection_points.keys())}）")
        print(f"生成予定: {total_count}シナリオ\n")
        
        for trigger in trigger_requirements:
            trigger_name = trigger.get("signal_name", "UNKNOWN")
            
            for point_name, point_config in injection_points.items():
                current_count += 1
                
                # 進捗表示
                print(f"[{current_count}/{total_count}] {trigger_name} × {point_name}", end=" → ")
                
                try:
                    # シナリオ生成
                    new_scenario = self._generate_single_scenario(
                        base_scenario=base_scenario,
                        trigger=trigger,
                        injection_point=point_name,
                        injection_config=point_config,
                        common_settings=common_settings
                    )
                    
                    # ファイル名生成
                    output_filename = f"{date_str}_{base_scenario_name}_{trigger_name}_{point_name}.json"
                    output_path = scenario_output_dir / output_filename
                    
                    # ファイル保存
                    self._save_scenario(new_scenario, output_path)
                    generated_files.append(str(output_path))
                    
                    print(f"{output_filename} ✓")
                    
                except Exception as e:
                    logger.error(f"シナリオ生成失敗: {trigger_name} × {point_name}: {str(e)}")
                    print(f"✗ エラー: {str(e)}")
                    
        # レポート生成
        self._generate_report(scenario_output_dir, generated_files, trigger_requirements, injection_points)
        
        print(f"\n=== 生成完了 ===")
        print(f"生成数: {len(generated_files)}シナリオ")
        print(f"出力先: {scenario_output_dir}")
        print(f"レポート: generation_report.txt")
        
        return generated_files
        
    def _load_base_scenario(self, scenario_name: str, base_file: str) -> Dict:
        """
        ベースシナリオファイルを読み込み
        
        Args:
            scenario_name: シナリオ名
            base_file: ベースファイル名
            
        Returns:
            シナリオデータ（辞書）
        """
        scenario_path = self.base_scenarios_dir / scenario_name / base_file
        
        if not scenario_path.exists():
            raise FileNotFoundError(f"ベースシナリオファイルが見つかりません: {scenario_path}")
            
        try:
            with open(scenario_path, 'r', encoding='utf-8') as f:
                scenario_data = json.load(f)
                logger.info(f"ベースシナリオを読み込みました: {scenario_path}")
                return scenario_data
        except Exception as e:
            logger.error(f"ベースシナリオの読み込みに失敗: {str(e)}")
            raise
            
    def _generate_single_scenario(self, base_scenario: Dict, trigger: Dict, 
                                 injection_point: str, injection_config: Dict,
                                 common_settings: Dict) -> Dict:
        """
        単一の異常系シナリオを生成
        
        Args:
            base_scenario: ベースシナリオ
            trigger: トリガー情報
            injection_point: 挿入位置名
            injection_config: 挿入位置設定
            common_settings: 共通設定
            
        Returns:
            新しいシナリオ
        """
        from .fault_injector import FaultInjector
        
        # FaultInjectorを使用してイベント挿入
        injector = FaultInjector()
        
        # 挿入位置の設定
        after_event = injection_config.get("after_event")
        default_delay = common_settings.get("injection_points", {}).get(injection_point, {}).get("default_delay", 3.0)
        delay = injection_config.get("delay", default_delay)
        
        # イベント挿入
        new_scenario = injector.inject_fault_event(
            base_scenario=base_scenario,
            trigger_info=trigger,
            after_event_no=after_event,
            delay=delay
        )
        
        # シナリオメタデータ更新
        new_scenario["scenario_summary"] = f"{base_scenario.get('scenario_summary', '')}_{trigger.get('signal_name', '')}_{injection_point}"
        new_scenario["variation"] = f"異常系_{injection_point}"
        
        return new_scenario
        
    def _save_scenario(self, scenario: Dict, output_path: Path) -> None:
        """
        シナリオをJSONファイルとして保存
        
        Args:
            scenario: シナリオデータ
            output_path: 出力パス
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(scenario, f, ensure_ascii=False, indent=2)
                logger.info(f"シナリオを保存しました: {output_path}")
        except Exception as e:
            logger.error(f"シナリオの保存に失敗: {str(e)}")
            raise
            
    def _generate_report(self, output_dir: Path, generated_files: List[str], 
                        triggers: List[Dict], injection_points: Dict) -> None:
        """
        生成レポートを作成
        
        Args:
            output_dir: 出力ディレクトリ
            generated_files: 生成したファイルリスト
            triggers: トリガー要件リスト
            injection_points: 挿入位置設定
        """
        report_path = output_dir / "generation_report.txt"
        
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("=" * 50 + "\n")
                f.write("異常系シナリオ生成レポート\n")
                f.write("=" * 50 + "\n\n")
                
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"出力先: {output_dir}\n\n")
                
                f.write("【生成統計】\n")
                f.write(f"- トリガー要件数: {len(triggers)}\n")
                f.write(f"- 挿入位置数: {len(injection_points)}\n")
                f.write(f"- 生成シナリオ数: {len(generated_files)}\n")
                f.write(f"- 期待シナリオ数: {len(triggers) * len(injection_points)}\n\n")
                
                f.write("【トリガー要件一覧】\n")
                for trigger in triggers:
                    f.write(f"- {trigger.get('signal_name', 'UNKNOWN')}: ")
                    f.write(f"JAMA_ID={trigger.get('jama_id', 'N/A')}, ")
                    f.write(f"値={trigger.get('value', 'N/A')}\n")
                    
                f.write("\n【挿入位置一覧】\n")
                for point_name, config in injection_points.items():
                    f.write(f"- {point_name}: イベント{config.get('after_event', 'N/A')}の後\n")
                    
                f.write("\n【生成ファイル一覧】\n")
                for file_path in generated_files:
                    f.write(f"- {Path(file_path).name}\n")
                    
                if len(generated_files) < len(triggers) * len(injection_points):
                    f.write("\n【エラー/スキップ】\n")
                    f.write(f"- {len(triggers) * len(injection_points) - len(generated_files)}件のシナリオ生成に失敗しました\n")
                    f.write("- 詳細はログファイルを確認してください\n")
                    
                logger.info(f"レポートを生成しました: {report_path}")
                
        except Exception as e:
            logger.error(f"レポート生成に失敗: {str(e)}")