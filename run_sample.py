#!/usr/bin/env python3
"""
JAMA要件管理ツールのサンプル実行スクリプト
基本的な使い方を示すためのサンプルコード
"""

import subprocess
import sys
import os


def run_command(cmd):
    """コマンドを実行して結果を表示"""
    print(f"\n実行コマンド: {cmd}")
    print("-" * 50)
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("エラー:", result.stderr)
    return result.returncode == 0


def main():
    """サンプル実行"""
    print("JAMA要件管理ツール - サンプル実行")
    print("=" * 50)
    
    # 1. テンプレートの作成（設定ファイル不要）
    print("\n### Excelテンプレートの作成 ###")
    if run_command("python main.py template -o sample_template.xlsx"):
        print("✅ テンプレート作成完了: sample_template.xlsx")
        print("   → JAMAに接続せずに、Excelの形式を確認できます")
    
    # 2. 設定ファイルの確認
    if not os.path.exists("config.json"):
        print("\n❌ config.json が見つかりません。")
        print("以下の機能はJAMAへの接続が必要なため、スキップします。")
        print("README.md を参照して設定ファイルを作成してください。")
        print("\n生成されたファイル:")
        print("  - sample_template.xlsx（テンプレート）")
        return
        
    # 2. ヘルプの表示
    print("\n### ヘルプ表示 ###")
    run_command("python main.py -h")
    
    # 3. 要件構造の取得（サンプル）
    print("\n### 要件構造の取得 ###")
    
    # プロジェクト全体を取得
    if run_command("python main.py fetch -o sample_all.xlsx"):
        print("✅ プロジェクト全体の取得完了: sample_all.xlsx")
        
    # 特定コンポーネント以下を取得（sequenceで指定）
    if run_command("python main.py fetch -o sample_component.xlsx -s 6"):
        print("✅ コンポーネント6以下の取得完了: sample_component.xlsx")
        
    # 最大3階層まで取得
    if run_command("python main.py fetch -o sample_depth3.xlsx -d 3"):
        print("✅ 3階層までの取得完了: sample_depth3.xlsx")
        
    # 4. 更新のサンプル（ドライランのみ）
    print("\n### 更新のドライラン ###")
    print("注意: 実際の更新を行う場合は、Excelファイルを編集してから")
    print("      python main.py update -i ファイル名.xlsx")
    print("      を実行してください。")
    
    if os.path.exists("sample_all.xlsx"):
        run_command("python main.py update -i sample_all.xlsx --dry-run")
        
    print("\n" + "=" * 50)
    print("サンプル実行完了")
    print("\n生成されたファイル:")
    for file in ["sample_all.xlsx", "sample_component.xlsx", "sample_depth3.xlsx"]:
        if os.path.exists(file):
            print(f"  - {file}")
            
    print("\n次のステップ:")
    print("1. 生成されたExcelファイルを開いて内容を確認")
    print("2. 必要に応じて編集")
    print("3. python main.py update -i ファイル名.xlsx で更新")


if __name__ == "__main__":
    main()
