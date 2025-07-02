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
        # テンプレート作成時は設定不要
        if config_path is None:
            self.config = None
            self.jama = None
            self.excel = ExcelHandler()
        else:
            self.config = Config(config_path)
            self.jama = JAMAClient(self.config)
            self.excel = ExcelHandler(self.config)
        
    def fetch_structure(self, 
                       output_file: str,
                       component_sequence: Optional[str] = None,
                       component_name: Optional[str] = None,
                       max_depth: Optional[int] = None,
                       debug: bool = False,
                       sample_mode: bool = False,
                       sample_count: int = 100,
                       count: Optional[int] = None) -> None:
        """
        JAMAから要件構造を取得してExcelに出力
        
        Args:
            output_file: 出力Excelファイル名
            component_sequence: 取得開始位置のsequence（例: "6.1.5"）
            component_name: 取得開始位置の名前
            max_depth: 取得する最大階層数
            debug: デバッグモードフラグ
            sample_mode: サンプルモード（少数のアイテムで構造調査）
            sample_count: サンプルモードで取得する件数
            count: 取得する最大件数（通常のfetchでも使用可能）
        """
        try:
            logger.info("JAMAから要件構造を取得開始")
            
            if debug:
                logger.info("デバッグモードが有効です")
                self.jama.set_debug_mode(True)
            
            if sample_mode:
                logger.info(f"サンプルモードが有効です（{sample_count}件取得）")
                self.jama.set_sample_mode(True)
            
            # プロジェクト情報取得
            project_info = self.jama.get_project_info()
            logger.info(f"プロジェクト: {project_info.get('fields', {}).get('name', 'Unknown')}")
            
            # 要件一覧取得
            items = []
            
            if sample_mode:
                # サンプルモード：指定件数のみ取得
                logger.info(f"サンプルモード: {sample_count}件のアイテムを取得")
                items = self.jama.get_sample_items(sample_count)
            elif component_sequence or component_name:
                # 特定コンポーネント以下を取得
                logger.info(f"コンポーネント指定: sequence={component_sequence}, name={component_name}")
                items = self.jama.get_items_by_component(
                    sequence=component_sequence,
                    name=component_name,
                    max_depth=max_depth,
                    max_count=count
                )
            else:
                # プロジェクト全体を取得
                logger.info("プロジェクト全体の要件を取得")
                if count:
                    logger.info(f"最大取得件数: {count}件")
                items = self.jama.get_all_items(max_depth=max_depth, max_count=count)
            
            logger.info(f"取得した要件数: {len(items)}")
            
            if len(items) > 1000:
                logger.warning(f"大量のデータ（{len(items)}件）を処理します。時間がかかる場合があります。")
                print(f"\n⚠️  大量のデータ（{len(items)}件）を処理します。")
                print("Excel作成には時間がかかる場合があります。しばらくお待ちください...")
            
            # Excelファイルに出力（サンプルモードでも出力する）
            output_path = Path(output_file)
            if not output_path.suffix:
                output_path = output_path.with_suffix('.xlsx')
                
            self.excel.create_requirement_excel(items, str(output_path))
            logger.info(f"Excelファイル作成完了: {output_path}")
            
            print(f"\n✅ 要件構造を正常に取得しました")
            print(f"📄 出力ファイル: {output_path}")
            print(f"📊 取得した要件数: {len(items)}")
            
            if sample_mode:
                print("\n📊 サンプルモードで実行されました")
                print("詳細なデバッグ情報はログファイルを確認してください: jama_tool.log")
            
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
                    "name": "サンプル要件1（新規作成例）",
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
                    "name": "既存要件の更新例（W列を「する」に設定）",
                    "assignee": "佐藤花子",
                    "status": "Review",
                    "tags": "更新,サンプル",
                    "reason": "",
                    "preconditions": "",
                    "target_system": "",
                    "description": ""
                },
                {
                    "jama_id": "12346",
                    "sequence": "1.2",
                    "name": "スキップ例（W列が空欄）",
                    "assignee": "山田次郎",
                    "status": "Approved",
                    "tags": "スキップ",
                    "reason": "",
                    "preconditions": "",
                    "target_system": "",
                    "description": ""
                },
                {
                    "jama_id": "12347",
                    "sequence": "1.3",
                    "name": "削除例（B列に「削除」と記入）",
                    "assignee": "高橋三郎",
                    "status": "Obsolete",
                    "tags": "削除予定",
                    "reason": "不要になった",
                    "preconditions": "",
                    "target_system": "",
                    "description": ""
                },
                {
                    "jama_id": "",
                    "sequence": "2",
                    "name": "SYSP: Description編集の例",
                    "assignee": "鈴木一郎",
                    "status": "Draft",
                    "tags": "SYSP,新規",
                    "reason": "サンプル理由",
                    "preconditions": "サンプル前提条件",
                    "target_system": "システムB",
                    "description": ""
                }
            ]
            
            # Excelファイル作成
            self.excel.create_requirement_excel(sample_items, str(output_path))
            
            print(f"\n✅ Excelテンプレートを作成しました")
            print(f"📄 出力ファイル: {output_path}")
            print("\n📝 テンプレートの使い方:")
            print("  1. 新規要件: JAMA_ID を空欄にして、必要な情報を記入")
            print("  2. 既存要件の更新: W列（要件更新）を「する」に設定")
            print("  3. 要件の削除: B列（メモ/コメント）に「削除」と入力")
            print("  4. スキップ: W列を空欄または「しない」に設定")
            print("  5. メモ: B列に自由にコメントを記入可能（「削除」以外）")
            print("  6. Description編集: SYSPアイテムは自動的にテンプレート作成")
            print("\n💡 ヒント: まずは少量のデータで試してみることをお勧めします")
            
        except Exception as e:
            logger.error(f"テンプレート作成に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)
            
    def update_requirements(self, input_file: str, dry_run: bool = False, debug: bool = False) -> None:
        """
        Excelファイルから要件を読み込んでJAMAを更新
        
        Args:
            input_file: 入力Excelファイル名
            dry_run: True の場合、実際の更新は行わない
            debug: True の場合、デバッグモードを有効にする
        """
        try:
            logger.info(f"Excelファイルから要件を読み込み: {input_file}")
            
            if debug:
                logger.info("デバッグモードが有効です")
                self.jama.set_debug_mode(True)
            
            # Excelから要件データを読み込み（進捗表示あり）
            requirements = self.excel.read_requirement_excel(input_file)
            
            if not requirements:
                print("更新対象の要件がありません")
                return
                
            print(f"\n📋 更新対象の要件数: {len(requirements)}")
            
            # 操作別に分類（進捗表示付き）
            logger.info("要件の分類開始")
            new_items = []
            update_items = []
            delete_items = []
            
            total_reqs = len(requirements)
            if total_reqs > 0:
                for idx, r in enumerate(requirements, 1):
                    if idx % 1000 == 0 or idx == total_reqs:
                        logger.info(f"要件分類進捗: {idx}/{total_reqs} ({idx/total_reqs*100:.1f}%)")
                        
                    if r['operation'] == '新規':
                        new_items.append(r)
                    elif r['operation'] == '更新':
                        update_items.append(r)
                    elif r['operation'] == '削除':
                        delete_items.append(r)
            
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
            
            # 新規作成（進捗表示付き）
            if new_items:
                print(f"\n新規作成開始: {len(new_items)}件")
                for idx, item in enumerate(new_items, 1):
                    # 最初、50件ごと、最後に表示
                    if idx == 1 or idx % 50 == 0 or idx == len(new_items):
                        print(f"  進捗: {idx}/{len(new_items)} ({idx/len(new_items)*100:.1f}%)")
                        
                    try:
                        logger.info(f"新規作成: {item.get('name', 'Unknown')}")
                        item_id = self.jama.create_item(item)
                        results['success'].append(f"✅ 新規作成: ID={item_id}, {item.get('name', '')}")
                    except Exception as e:
                        logger.error(f"新規作成失敗: {str(e)}")
                        results['failed'].append(f"❌ 新規作成失敗: {item.get('name', '')}, エラー: {str(e)}")
                        
            # 更新（進捗表示付き）
            if update_items:
                print(f"\n更新開始: {len(update_items)}件")
                for idx, item in enumerate(update_items, 1):
                    # 最初、50件ごと、最後に表示
                    if idx == 1 or idx % 50 == 0 or idx == len(update_items):
                        print(f"  進捗: {idx}/{len(update_items)} ({idx/len(update_items)*100:.1f}%)")
                        
                    try:
                        logger.info(f"更新: ID={item['jama_id']}, {item.get('name', 'Unknown')}")
                        self.jama.update_item(item['jama_id'], item)
                        results['success'].append(f"✅ 更新: ID={item['jama_id']}, {item.get('name', '')}")
                    except Exception as e:
                        logger.error(f"更新失敗: {str(e)}")
                        results['failed'].append(f"❌ 更新失敗: ID={item['jama_id']}, エラー: {str(e)}")
                        
            # 削除（進捗表示付き）
            if delete_items:
                print(f"\n削除開始: {len(delete_items)}件")
                for idx, item in enumerate(delete_items, 1):
                    # 最初、50件ごと、最後に表示
                    if idx == 1 or idx % 50 == 0 or idx == len(delete_items):
                        print(f"  進捗: {idx}/{len(delete_items)} ({idx/len(delete_items)*100:.1f}%)")
                        
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
            for idx, item in enumerate(new_items, 1):
                print(f"  {idx}. {item.get('name', 'Unknown')}")
            
        if update_items:
            print(f"\n【更新予定】{len(update_items)}件")
            for idx, item in enumerate(update_items, 1):
                print(f"  {idx}. ID: {item['jama_id']}, {item.get('name', 'Unknown')}")
            
                # 更新されるフィールドを特定
                update_fields = []
                if item.get('description'):
                    update_fields.append('description')
                if item.get('tags'):
                    update_fields.append('tags')
                if item.get('reason'):
                    update_fields.append('reason')
                if item.get('status'):
                    update_fields.append('status')
                if item.get('assignee'):
                    update_fields.append('assignee')
                if item.get('preconditions'):
                    update_fields.append('preconditions')
                if item.get('target_system'):
                    update_fields.append('target_system')
                
                if update_fields:
                    print(f"     更新フィールド: {', '.join(update_fields)}")
                else:
                    print(f"     更新フィールド: なし（変更なし）")
            
        if delete_items:
            print("\n【削除予定】")
            for idx, item in enumerate(delete_items, 1):
                print(f"  {idx}. ID: {item['jama_id']}, {item.get('name', 'Unknown')}")


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
  
  # 最大500件のみ取得
  %(prog)s fetch -o requirements.xlsx --count 500
  
  # サンプルモードで構造調査（100件取得）
  %(prog)s fetch -o test.xlsx --sample-mode --sample-count 100
  
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
    fetch_parser.add_argument('--debug', action='store_true',
                             help='デバッグモードを有効にする')
    fetch_parser.add_argument('--sample-mode', action='store_true',
                             help='サンプルモード（少数のアイテムで構造調査）')
    fetch_parser.add_argument('--sample-count', type=int, default=100,
                             help='サンプルモードで取得する件数（デフォルト: 100）')
    fetch_parser.add_argument('--count', type=int,
                             help='取得する最大件数（通常のfetchでも使用可能）')
    
    # updateコマンド
    update_parser = subparsers.add_parser('update', help='Excelファイルから要件を更新')
    update_parser.add_argument('-i', '--input', required=True,
                            help='入力Excelファイル名')
    update_parser.add_argument('--dry-run', action='store_true',
                            help='実際の更新は行わない（プレビューのみ）')
    update_parser.add_argument('--debug', action='store_true',
                            help='デバッグモードを有効にする')
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
                max_depth=args.max_depth,
                debug=args.debug if hasattr(args, 'debug') else False,
                sample_mode=args.sample_mode if hasattr(args, 'sample_mode') else False,
                sample_count=args.sample_count if hasattr(args, 'sample_count') else 100,
                count=args.count if hasattr(args, 'count') else None
            )
        elif args.command == 'update':
            tool.update_requirements(
                input_file=args.input,
                dry_run=args.dry_run,
                debug=args.debug if hasattr(args, 'debug') else False
            )


if __name__ == '__main__':
    main()