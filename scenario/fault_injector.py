"""
異常イベント挿入処理
ベースシナリオに異常系トリガーイベントを挿入し、イベント番号を振り直す
"""

import copy
import logging
from typing import Dict, List, Any, Optional

logger = logging.getLogger(__name__)


class FaultInjector:
    """異常イベント挿入クラス"""
    
    def inject_fault_event(self, base_scenario: Dict, trigger_info: Dict, 
                          after_event_no: int, delay: float = 3.0) -> Dict:
        """
        ベースシナリオに異常イベントを挿入
        
        Args:
            base_scenario: ベースシナリオ（変更しない）
            trigger_info: トリガー情報（signal_name, value, action_type等）
            after_event_no: この番号のイベントの後に挿入
            delay: トリガー遅延時間（秒）
            
        Returns:
            新しいシナリオ（異常イベント挿入済み）
        """
        # ディープコピーで元のシナリオを保護
        new_scenario = copy.deepcopy(base_scenario)
        
        # storyセクションのイベントリストを取得
        if "story" not in new_scenario or len(new_scenario["story"]) == 0:
            raise ValueError("シナリオにstoryセクションがありません")
            
        # 通常はegoアクターのイベントを対象とする
        ego_story = None
        for actor_story in new_scenario["story"]:
            if actor_story.get("actor_name") == "ego":
                ego_story = actor_story
                break
                
        if not ego_story:
            raise ValueError("egoアクターのstoryが見つかりません")
            
        events = ego_story.get("events", [])
        if not events:
            raise ValueError("イベントリストが空です")
            
        # 挿入位置を特定
        insert_index = self._find_insert_position(events, after_event_no)
        if insert_index == -1:
            raise ValueError(f"イベント番号 {after_event_no} が見つかりません")
            
        logger.info(f"イベント{after_event_no}の後（インデックス{insert_index}）に異常イベントを挿入します")
        
        # 異常イベントを作成
        fault_event = self._create_fault_event(
            trigger_info=trigger_info,
            new_event_no=after_event_no + 1,
            trigger_event_no=after_event_no,
            delay=delay
        )
        
        # イベントを挿入
        events.insert(insert_index, fault_event)
        
        # 後続イベントの番号を振り直し
        self._renumber_events(events, insert_index + 1, after_event_no + 2)
        
        # イベント参照を更新
        self._update_event_references(events, after_event_no + 1)
        
        # test_procedure_orderを更新
        self._update_test_procedure_order(events)
        
        logger.info(f"異常イベント挿入完了: {trigger_info.get('signal_name', 'UNKNOWN')}")
        
        return new_scenario
        
    def _find_insert_position(self, events: List[Dict], after_event_no: int) -> int:
        """
        挿入位置のインデックスを検索
        
        Args:
            events: イベントリスト
            after_event_no: この番号のイベントの後に挿入
            
        Returns:
            挿入位置のインデックス（見つからない場合は-1）
        """
        for i, event in enumerate(events):
            if event.get("no") == after_event_no:
                return i + 1  # 次の位置に挿入
        return -1
        
    def _create_fault_event(self, trigger_info: Dict, new_event_no: int, 
                           trigger_event_no: int, delay: float) -> Dict:
        """
        異常イベントを作成
        
        Args:
            trigger_info: トリガー情報
            new_event_no: 新しいイベント番号
            trigger_event_no: トリガーとなるイベント番号
            delay: 遅延時間
            
        Returns:
            異常イベント
        """
        # 基本構造
        fault_event = {
            "no": new_event_no,
            "times": 1,
            "action": {
                "type": trigger_info.get("action_type", "can_communication_error"),
                "params": []
            },
            "start_trigger": {
                "condition_groups": [
                    {
                        "conditions": [
                            {
                                "type": "event_state",
                                "params": [
                                    {
                                        "rule": "equalTo",
                                        "name": "event_no",
                                        "value": trigger_event_no,
                                        "unit": ""
                                    },
                                    {
                                        "rule": "equalTo",
                                        "name": "state",
                                        "value": "completeState",
                                        "unit": ""
                                    }
                                ],
                                "delay": delay
                            }
                        ]
                    }
                ]
            },
            "criteria": [
                {
                    "target_name": "-",
                    "expressions": []
                }
            ],
            "remarks": [
                f"異常系トリガー: {trigger_info.get('signal_name', 'UNKNOWN')}"
            ],
            "test_procedure_order": new_event_no  # 後で更新される
        }
        
        # アクションパラメータ設定
        action_params = []
        
        # valueパラメータ（ほとんどのアクションで使用）
        if "value" in trigger_info:
            action_params.append({
                "name": "value",
                "value": str(trigger_info["value"]),
                "unit": ""
            })
        else:
            # デフォルト値
            action_params.append({
                "name": "value",
                "value": "1",
                "unit": ""
            })
            
        # 特定のアクションタイプに応じた追加パラメータ
        action_type = trigger_info.get("action_type", "")
        
        if "communication_error" in action_type and "target" in trigger_info:
            action_params.append({
                "name": "target",
                "value": trigger_info["target"],
                "unit": ""
            })
            
        elif action_type == "brake":
            # ブレーキの場合は％単位
            if len(action_params) > 0:
                action_params[0]["unit"] = "%"
                
        elif action_type == "steering":
            # ステアリングの場合は角度と時間
            action_params = [
                {
                    "name": "target_rudder_angle",
                    "value": trigger_info.get("value", 45),
                    "unit": "deg"
                },
                {
                    "name": "steering_time",
                    "value": 1,
                    "unit": "s"
                }
            ]
            
        fault_event["action"]["params"] = action_params
        
        return fault_event
        
    def _renumber_events(self, events: List[Dict], start_index: int, start_no: int) -> None:
        """
        イベント番号を振り直し
        
        Args:
            events: イベントリスト
            start_index: 振り直し開始インデックス
            start_no: 振り直し開始番号
        """
        for i in range(start_index, len(events)):
            old_no = events[i]["no"]
            new_no = start_no + (i - start_index)
            events[i]["no"] = new_no
            logger.debug(f"イベント番号変更: {old_no} → {new_no}")
            
    def _update_event_references(self, events: List[Dict], inserted_event_no: int) -> None:
        """
        イベント参照を更新（start_triggerのevent_no参照）
        
        Args:
            events: イベントリスト
            inserted_event_no: 挿入されたイベント番号
        """
        for event in events:
            # start_trigger内のevent_no参照を更新
            if "start_trigger" in event:
                self._update_trigger_references(event["start_trigger"], inserted_event_no)
                
        # end_storyも更新が必要な場合（今回は省略）
        
    def _update_trigger_references(self, trigger: Dict, inserted_event_no: int) -> None:
        """
        トリガー内のイベント参照を更新
        
        Args:
            trigger: トリガー設定
            inserted_event_no: 挿入されたイベント番号
        """
        if "condition_groups" not in trigger:
            return
            
        for group in trigger["condition_groups"]:
            if "conditions" not in group:
                continue
                
            for condition in group["conditions"]:
                if condition.get("type") == "event_state":
                    params = condition.get("params", [])
                    for param in params:
                        if param.get("name") == "event_no":
                            old_value = param.get("value", 0)
                            # 挿入イベント番号以降の参照は+1する
                            if isinstance(old_value, int) and old_value >= inserted_event_no:
                                param["value"] = old_value + 1
                                logger.debug(f"イベント参照更新: {old_value} → {old_value + 1}")
                                
    def _update_test_procedure_order(self, events: List[Dict]) -> None:
        """
        test_procedure_orderを連番に更新
        
        Args:
            events: イベントリスト
        """
        for i, event in enumerate(events, 1):
            event["test_procedure_order"] = i
            
    def validate_scenario(self, scenario: Dict) -> List[str]:
        """
        生成されたシナリオの妥当性を検証
        
        Args:
            scenario: 検証対象のシナリオ
            
        Returns:
            エラーメッセージのリスト（空なら問題なし）
        """
        errors = []
        
        # イベント番号の連続性チェック
        if "story" in scenario:
            for actor_story in scenario["story"]:
                events = actor_story.get("events", [])
                expected_no = 1
                for event in events:
                    if event.get("no") != expected_no:
                        errors.append(f"イベント番号が不連続: 期待値={expected_no}, 実際={event.get('no')}")
                    expected_no += 1
                    
        # トリガー参照の妥当性チェック
        if "story" in scenario:
            for actor_story in scenario["story"]:
                events = actor_story.get("events", [])
                event_numbers = {event.get("no") for event in events}
                
                for event in events:
                    if "start_trigger" in event:
                        referenced_events = self._extract_referenced_events(event["start_trigger"])
                        for ref_no in referenced_events:
                            if ref_no not in event_numbers and ref_no != 0:  # 0は時間トリガー
                                errors.append(f"イベント{event.get('no')}が存在しないイベント{ref_no}を参照")
                                
        return errors
        
    def _extract_referenced_events(self, trigger: Dict) -> List[int]:
        """
        トリガーから参照されているイベント番号を抽出
        
        Args:
            trigger: トリガー設定
            
        Returns:
            参照イベント番号のリスト
        """
        referenced = []
        
        if "condition_groups" in trigger:
            for group in trigger["condition_groups"]:
                if "conditions" in group:
                    for condition in group["conditions"]:
                        if condition.get("type") == "event_state":
                            params = condition.get("params", [])
                            for param in params:
                                if param.get("name") == "event_no":
                                    value = param.get("value")
                                    if isinstance(value, int):
                                        referenced.append(value)
                                        
        return referenced