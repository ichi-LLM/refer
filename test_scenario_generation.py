"""
シナリオ生成機能のテストスクリプト
CAN_BRAKE_ERRORのサンプルでテストを実行
"""

import json
import os
from pathlib import Path
import sys
import difflib
from datetime import datetime

# パスを追加（必要に応じて調整）
sys.path.insert(0, str(Path(__file__).parent))

from scenario.generator import ScenarioGenerator
from scenario.fault_injector import FaultInjector
from scenario.requirement_parser import RequirementParser


def test_fault_injection():
    """イベント挿入のテスト"""
    print("\n=== イベント挿入テスト ===")
    
    # サンプルのベースシナリオ（簡略版）
    base_scenario = {
        "scenario_summary": "テストシナリオ",
        "story": [
            {
                "actor_name": "ego",
                "events": [
                    {
                        "no": 1,
                        "action": {"type": "engine_startup_operation"},
                        "test_procedure_order": 1
                    },
                    {
                        "no": 2,
                        "action": {"type": "shift"},
                        "start_trigger": {
                            "condition_groups": [{
                                "conditions": [{
                                    "type": "event_state",
                                    "params": [
                                        {"name": "event_no", "value": 1},
                                        {"name": "state", "value": "completeState"}
                                    ]
                                }]
                            }]
                        },
                        "test_procedure_order": 2
                    },
                    {
                        "no": 3,
                        "action": {"type": "appcssw"},
                        "start_trigger": {
                            "condition_groups": [{
                                "conditions": [{
                                    "type": "event_state",
                                    "params": [
                                        {"name": "event_no", "value": 2},
                                        {"name": "state", "value": "completeState"}
                                    ]
                                }]
                            }]
                        },
                        "test_procedure_order": 3
                    }
                ]
            }
        ]
    }
    
    # トリガー情報
    trigger_info = {
        "jama_id": "115200",
        "signal_name": "CAN_BRAKE_ERROR",
        "value": 1,
        "action_type": "can_communication_error"
    }
    
    # FaultInjectorのテスト
    injector = FaultInjector()
    
    # イベント2の後に挿入
    new_scenario = injector.inject_fault_event(
        base_scenario=base_scenario,
        trigger_info=trigger_info,
        after_event_no=2,
        delay=3.0
    )
    
    # 結果確認
    events = new_scenario["story"][0]["events"]
    
    print(f"元のイベント数: 3")
    print(f"新しいイベント数: {len(events)}")
    
    print("\nイベント一覧:")
    for event in events:
        print(f"  Event {event['no']}: {event['action']['type']}")
        
    # 挿入されたイベントの確認
    inserted_event = events[2]  # インデックス2が新しく挿入されたイベント
    assert inserted_event["no"] == 3
    assert inserted_event["action"]["type"] == "can_communication_error"
    
    # 後続イベントの番号確認
    assert events[3]["no"] == 4  # 元のイベント3が4になる
    
    # トリガー参照の更新確認
    last_trigger = events[3]["start_trigger"]["condition_groups"][0]["conditions"][0]
    trigger_params = last_trigger["params"]
    event_no_param = next(p for p in trigger_params if p["name"] == "event_no")
    assert event_no_param["value"] == 3  # 2→3に更新されている
    
    print("\n✅ イベント挿入テスト成功")
    
    return new_scenario


def test_scenario_comparison(base_file: str, generated_file: str):
    """生成されたシナリオとベースシナリオの差分を表示"""
    print("\n=== シナリオ差分確認 ===")
    
    # ファイル読み込み
    with open(base_file, 'r', encoding='utf-8') as f:
        base_data = json.load(f)
    
    with open(generated_file, 'r', encoding='utf-8') as f:
        generated_data = json.load(f)
    
    # イベント数の比較
    base_events = base_data["story"][0]["events"]
    gen_events = generated_data["story"][0]["events"]
    
    print(f"ベースシナリオ: {len(base_events)}イベント")
    print(f"生成シナリオ: {len(gen_events)}イベント")
    print(f"差分: +{len(gen_events) - len(base_events)}イベント")
    
    # 挿入されたイベントを特定
    print("\n挿入されたイベント:")
    for i, event in enumerate(gen_events):
        # 対応するベースイベントを探す
        is_new = True
        for base_event in base_events:
            if event["action"]["type"] == base_event["action"]["type"] and \
               "can_" not in event["action"]["type"]:  # 異常系以外で一致
                is_new = False
                break
        
        if is_new or "can_" in event["action"]["type"]:
            print(f"  → Event {event['no']}: {event['action']['type']}")
            if "params" in event["action"]:
                for param in event["action"]["params"]:
                    print(f"      {param['name']}: {param['value']}{param.get('unit', '')}")
    
    # JSON差分（詳細）
    print("\n詳細な差分を表示しますか？ (y/N): ", end="")
    if input().lower() == 'y':
        base_json = json.dumps(base_data, indent=2, ensure_ascii=False).splitlines()
        gen_json = json.dumps(generated_data, indent=2, ensure_ascii=False).splitlines()
        
        diff = difflib.unified_diff(base_json, gen_json, lineterm='',
                                   fromfile='base_scenario.json',
                                   tofile='generated_scenario.json')
        
        for line in diff:
            if line.startswith('+'):
                print(f"\033[92m{line}\033[0m")  # 緑色
            elif line.startswith('-'):
                print(f"\033[91m{line}\033[0m")  # 赤色
            else:
                print(line)


def main():
    """メインテスト処理"""
    print("=" * 50)
    print("異常系シナリオ生成機能テスト")
    print("=" * 50)
    
    # 1. イベント挿入テスト
    new_scenario = test_fault_injection()
    
    # 2. テスト用シナリオをファイルに保存
    test_output_dir = Path("scenario/output/test")
    test_output_dir.mkdir(parents=True, exist_ok=True)
    
    date_str = datetime.now().strftime("%Y-%m-%d")
    output_file = test_output_dir / f"{date_str}_test_CAN_BRAKE_ERROR.json"
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(new_scenario, f, ensure_ascii=False, indent=2)
    
    print(f"\nテストシナリオを保存: {output_file}")
    
    # 3. 実際のベースシナリオでテスト（存在する場合）
    base_scenario_path = Path("scenario/base_scenarios/AP_出庫/AP_basic_sample1.json")
    if base_scenario_path.exists():
        print("\n=== 実際のベースシナリオでテスト ===")
        
        # ScenarioGeneratorのテスト
        generator = ScenarioGenerator()
        
        # サンプルトリガー
        sample_triggers = [
            {
                "jama_id": "115200",
                "requirement_name": "ブレーキ通信途絶",
                "signal_name": "CAN_BRAKE_ERROR",
                "value": 1,
                "action_type": "can_communication_error",
                "remarks": "テスト用"
            }
        ]
        
        try:
            generated_files = generator.generate_scenarios(
                base_scenario_name="AP_出庫",
                trigger_requirements=sample_triggers,
                output_dir=str(test_output_dir)
            )
            
            if generated_files:
                print(f"\n生成されたファイル:")
                for file_path in generated_files:
                    print(f"  - {Path(file_path).name}")
                    
                # 差分確認
                test_scenario_comparison(
                    str(base_scenario_path),
                    generated_files[0]
                )
        except Exception as e:
            print(f"⚠️ エラー: {str(e)}")
            print("scenario_mappings.yamlの設定を確認してください")
    
    print("\n" + "=" * 50)
    print("テスト完了")
    print("=" * 50)


if __name__ == "__main__":
    main()