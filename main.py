#!/usr/bin/env python3
"""
JAMA要件管理ツール
JAMAの要件を取得・更新するためのコマンドラインツール
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime
import logging
from typing import Optional

from jama_client import JAMAClient
from excel_handler import ExcelHandler
from config import Config

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('jama_tool.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class JAMATool:
    """JAMA要件管理ツールのメインクラス"""
    
    def __init__(self, config_path: str = "config.json"):
        """
        初期化
        
        Args:
            config_path: 設定ファイルのパス
        """
        self.excel = ExcelHandler()
        
        # テンプレート作成時は設定不要
        if config_path is None:
            self.config = None
            self.jama = None
        else:
            self.config = Config(config_path)
            self.jama = JAMAClient(self.config)
        
    def fetch_structure(self, 
                       output_file: str,
                       component_sequence: Optional[str] = None,
                       component_name: Optional[str] = None,
                       max_depth: Optional[int] = None) -> None:
        """
        JAMAから要件構造を取得してExcelに出力
        
        Args:
            output_file: 出力Excelファイル名
            component_sequence: 取得開始位置のsequence（例: "6.1.5"）
            component_name: 取得開始位置の名前
            max_depth: 取得する最大階層数
        """
        try:
            logger.info("JAMAから要件構造を取得開始")
            
            # プロジェクト情報取得
            project_info = self.jama.get_project_info()
            logger.info(f"プロジェクト: {project_info.get('fields', {}).get('name', 'Unknown')}")
            
            # 要件一覧取得
            items = []
            
            if component_sequence or component_name:
                # 特定コンポーネント以下を取得
                logger.info(f"コンポーネント指定: sequence={component_sequence}, name={component_name}")
                items = self.jama.get_items_by_component(
                    sequence=component_sequence,
                    name=component_name,
                    max_depth=max_depth
                )
            else:
                # プロジェクト全体を取得
                logger.info("プロジェクト全体の要件を取得")
                items = self.jama.get_all_items(max_depth=max_depth)
            
            logger.info(f"取得した要件数: {len(items)}")
            
            # Excelファイルに出力
            output_path = Path(output_file)
            if not output_path.suffix:
                output_path = output_path.with_suffix('.xlsx')
                
            self.excel.create_requirement_excel(items, str(output_path))
            logger.info(f"Excelファイル作成完了: {output_path}")
            
            print(f"\n✅ 要件構造を正常に取得しました")
            print(f"📄 出力ファイル: {output_path}")
            print(f"📊 取得した要件数: {len(items)}")
            
        except Exception as e:
            logger.error(f"要件構造の取得に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)
            
    def create_template(self, output_file: str) -> None:
        """
        空のExcelテンプレートを作成
        
        Args:
            output_file: 出力ファイル名
        """
        try:
            logger.info("Excelテンプレート作成開始")
            
            # 出力パス設定
            output_path = Path(output_file)
            if not output_path.suffix:
                output_path = output_path.with_suffix('.xlsx')
                
            # 空のアイテムリストでテンプレート作成
            sample_items = [
                {
                    "jama_id": "",
                    "sequence": "1",
                    "name": "サンプル要件1",
                    "assignee": "田中太郎",
                    "status": "Draft",
                    "tags": "サンプル,テスト",
                    "reason": "テンプレート例",
                    "preconditions": "特になし",
                    "target_system": "システムA",
                    "description": "<table><tr><td>IN</td><td>OUT</td></tr></table>"
                },
                {
                    "jama_id": "12345",
                    "sequence": "1.1",
                    "name": "既存要件の更新例",
                    "assignee": "佐藤花子",
                    "status": "Review",
                    "tags": "更新,サンプル",
                    "reason": "",
                    "preconditions": "",
                    "target_system": "",
                    "description": ""
                },
                {
                    "jama_id": "",
                    "sequence": "2",
                    "name": "SYSP: Description編集の例",
                    "assignee": "山田次郎",
                    "status": "Draft",
                    "tags": "SYSP,新規",
                    "reason": "",
                    "preconditions": "",
                    "target_system": "",
                    "description": ""
                }
            ]
            
            # Excelファイル作成
            self.excel.create_requirement_excel(sample_items, str(output_path))
            
            print(f"\n✅ Excelテンプレートを作成しました")
            print(f"📄 出力ファイル: {output_path}")
            print("\n📝 テンプレートの使い方:")
            print("  1. 新規要件: JAMA_ID を空欄にして、必要な情報を記入")
            print("  2. 既存要件の更新: JAMA_ID を記入して、変更したい内容を編集")
            print("  3. 要件の削除: 操作列に「削除」と記入")
            print("  4. Description編集: W列を「する」にして、Description_editシートで編集")
            print("\n💡 ヒント: まずは少量のデータで試してみることをお勧めします")
            
        except Exception as e:
            logger.error(f"テンプレート作成に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)
            
    def update_requirements(self, input_file: str, dry_run: bool = False) -> None:
        """
        Excelファイルから要件を読み込んでJAMAを更新
        
        Args:
            input_file: 入力Excelファイル名
            dry_run: True の場合、実際の更新は行わない
        """
        try:
            logger.info(f"Excelファイルから要件を読み込み: {input_file}")
            
            # Excelから要件データを読み込み
            requirements = self.excel.read_requirement_excel(input_file)
            
            if not requirements:
                print("更新対象の要件がありません")
                return
                
            print(f"\n📋 更新対象の要件数: {len(requirements)}")
            
            # 操作別に分類
            new_items = [r for r in requirements if r['operation'] == '新規']
            update_items = [r for r in requirements if r['operation'] == '更新']
            delete_items = [r for r in requirements if r['operation'] == '削除']
            
            print(f"  新規作成: {len(new_items)}件")
            print(f"  更新: {len(update_items)}件")
            print(f"  削除: {len(delete_items)}件")
            
            if dry_run:
                print("\n🔍 ドライランモード - 実際の更新は行いません")
                self._show_update_preview(new_items, update_items, delete_items)
                return
                
            # 確認
            response = input("\n実行しますか？ (y/N): ")
            if response.lower() != 'y':
                print("キャンセルしました")
                return
                
            # 更新実行
            results = {
                'success': [],
                'failed': []
            }
            
            # 新規作成
            for item in new_items:
                try:
                    logger.info(f"新規作成: {item.get('name', 'Unknown')}")
                    item_id = self.jama.create_item(item)
                    results['success'].append(f"✅ 新規作成: ID={item_id}, {item.get('name', '')}")
                except Exception as e:
                    logger.error(f"新規作成失敗: {str(e)}")
                    results['failed'].append(f"❌ 新規作成失敗: {item.get('name', '')}, エラー: {str(e)}")
                    
            # 更新
            for item in update_items:
                try:
                    logger.info(f"更新: ID={item['jama_id']}, {item.get('name', 'Unknown')}")
                    self.jama.update_item(item['jama_id'], item)
                    results['success'].append(f"✅ 更新: ID={item['jama_id']}, {item.get('name', '')}")
                except Exception as e:
                    logger.error(f"更新失敗: {str(e)}")
                    results['failed'].append(f"❌ 更新失敗: ID={item['jama_id']}, エラー: {str(e)}")
                    
            # 削除
            for item in delete_items:
                try:
                    logger.info(f"削除: ID={item['jama_id']}, {item.get('name', 'Unknown')}")
                    self.jama.delete_item(item['jama_id'])
                    results['success'].append(f"✅ 削除: ID={item['jama_id']}, {item.get('name', '')}")
                except Exception as e:
                    logger.error(f"削除失敗: {str(e)}")
                    results['failed'].append(f"❌ 削除失敗: ID={item['jama_id']}, エラー: {str(e)}")
                    
            # 結果表示
            print("\n📊 実行結果:")
            print(f"成功: {len(results['success'])}件")
            print(f"失敗: {len(results['failed'])}件")
            
            if results['success']:
                print("\n成功した操作:")
                for msg in results['success']:
                    print(f"  {msg}")
                    
            if results['failed']:
                print("\n失敗した操作:")
                for msg in results['failed']:
                    print(f"  {msg}")
                    
        except Exception as e:
            logger.error(f"要件の更新に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)
            
    def _show_update_preview(self, new_items, update_items, delete_items):
        """更新内容のプレビューを表示"""
        if new_items:
            print("\n【新規作成予定】")
            for item in new_items[:5]:  # 最初の5件のみ表示
                print(f"  - {item.get('name', 'Unknown')}")
            if len(new_items) > 5:
                print(f"  ... 他 {len(new_items) - 5}件")
                
        if update_items:
            print("\n【更新予定】")
            for item in update_items[:5]:
                print(f"  - ID: {item['jama_id']}, {item.get('name', 'Unknown')}")
            if len(update_items) > 5:
                print(f"  ... 他 {len(update_items) - 5}件")
                
        if delete_items:
            print("\n【削除予定】")
            for item in delete_items[:5]:
                print(f"  - ID: {item['jama_id']}, {item.get('name', 'Unknown')}")
            if len(delete_items) > 5:
                print(f"  ... 他 {len(delete_items) - 5}件")


def main():
    """メイン処理"""
    parser = argparse.ArgumentParser(
        description='JAMA要件管理ツール',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 空のテンプレートを作成（JAMAへの接続不要）
  %(prog)s template -o template.xlsx
  
  # プロジェクト全体の要件を取得
  %(prog)s fetch -o requirements.xlsx
  
  # 特定のコンポーネント以下を取得（sequenceで指定）
  %(prog)s fetch -o requirements.xlsx -s 6.1.5
  
  # 特定のコンポーネント以下を取得（名前で指定）
  %(prog)s fetch -o requirements.xlsx -n "公共駐車場"
  
  # 最大3階層まで取得
  %(prog)s fetch -o requirements.xlsx -d 3
  
  # Excelファイルから要件を更新（ドライラン）
  %(prog)s update -i requirements.xlsx --dry-run
  
  # Excelファイルから要件を更新（実行）
  %(prog)s update -i requirements.xlsx
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='実行するコマンド')
    
    # fetchコマンド
    fetch_parser = subparsers.add_parser('fetch', help='JAMAから要件構造を取得')
    fetch_parser.add_argument('-o', '--output', required=True,
                             help='出力Excelファイル名')
    fetch_parser.add_argument('-s', '--sequence',
                             help='取得開始位置のsequence（例: 6.1.5）')
    fetch_parser.add_argument('-n', '--name',
                             help='取得開始位置のアイテム名')
    fetch_parser.add_argument('-d', '--max-depth', type=int,
                             help='取得する最大階層数')
    fetch_parser.add_argument('-c', '--config', default='config.json',
                             help='設定ファイルのパス（デフォルト: config.json）')
    
    # updateコマンド
    update_parser = subparsers.add_parser('update', help='Excelファイルから要件を更新')
    update_parser.add_argument('-i', '--input', required=True,
                              help='入力Excelファイル名')
    update_parser.add_argument('--dry-run', action='store_true',
                              help='実際の更新は行わない（プレビューのみ）')
    update_parser.add_argument('-c', '--config', default='config.json',
                              help='設定ファイルのパス（デフォルト: config.json）')
    
    # templateコマンド
    template_parser = subparsers.add_parser('template', help='空のExcelテンプレートを作成')
    template_parser.add_argument('-o', '--output', required=True,
                                help='出力Excelファイル名')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(1)
        
    # テンプレート作成の場合は設定ファイル不要
    if args.command == 'template':
        tool = JAMATool(config_path=None)
        tool.create_template(output_file=args.output)
    else:
        # その他のコマンドは設定ファイルが必要
        config_path = args.config if hasattr(args, 'config') else 'config.json'
        tool = JAMATool(config_path)
        
        # コマンド実行
        if args.command == 'fetch':
            tool.fetch_structure(
                output_file=args.output,
                component_sequence=args.sequence,
                component_name=args.name,
                max_depth=args.max_depth
            )
        elif args.command == 'update':
            tool.update_requirements(
                input_file=args.input,
                dry_run=args.dry_run
            )


if __name__ == '__main__':
    main()
