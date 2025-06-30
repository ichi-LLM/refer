#!/usr/bin/env python3
"""
JAMA要件管理ツール
JAMAの要件を取得・更新するためのコマンドラインツール
"""

import argparse
import sys
from pathlib import Path
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
        logging.FileHandler('jama_tool.log', encoding='utf-8', mode='w'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class JAMATool:
    """JAMA要件管理ツールのメインクラス"""

    def __init__(self, config: Optional[Config] = None):
        self.config = config
        self.jama = JAMAClient(config) if config else None
        self.excel = ExcelHandler(config)

    def fetch_structure(self,
                       output_file: str,
                       component_sequence: Optional[str] = None,
                       component_name: Optional[str] = None,
                       max_depth: Optional[int] = None,
                       debug: bool = False) -> None:
        """JAMAから要件構造を取得してExcelに出力"""
        if not self.jama:
            logger.error("JAMAクライアントが初期化されていません。")
            return

        try:
            logger.info("JAMAから要件構造の取得を開始")
            if debug:
                logger.info("デバッグモードが有効です")
                self.jama.set_debug_mode(True)

            project_info = self.jama.get_project_info()
            logger.info(f"プロジェクト: {project_info.get('fields', {}).get('name', 'Unknown')}")

            items = []
            if component_sequence or component_name:
                # 特定コンポーネント以下を効率的に取得
                items = self.jama.get_items_by_component(
                    sequence=component_sequence,
                    name=component_name,
                    max_depth=max_depth
                )
            else:
                # プロジェクト全体を取得
                logger.info("プロジェクト全体の要件を取得します。")
                items = self.jama.get_all_items(max_depth=max_depth)

            logger.info(f"取得した合計要件数: {len(items)}")

            if not items:
                print("\n✅ 取得対象の要件が見つかりませんでした。")
                return
            
            if len(items) > 1000:
                logger.warning(f"大量のデータ（{len(items)}件）を処理します。時間がかかる場合があります。")
                print(f"\n⚠️  大量のデータ（{len(items)}件）を処理します。Excel作成に時間がかかる場合があります。")
            
            output_path = Path(output_file).with_suffix('.xlsx')
            self.excel.create_requirement_excel(items, str(output_path))
            
            print(f"\n✅ 要件構造を正常に取得しました。")
            print(f"📄 出力ファイル: {output_path}")
            print(f"📊 取得した要件数: {len(items)}")

        except Exception as e:
            logger.error(f"要件構造の取得に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)

    def create_template(self, output_file: str) -> None:
        """空のExcelテンプレートを作成"""
        try:
            logger.info("Excelテンプレート作成開始")
            output_path = Path(output_file).with_suffix('.xlsx')

            # サンプルデータ
            sample_items = [
                {"jama_id": "", "sequence": "1", "name": "サンプル要件1 (新規作成の例)"},
                {"jama_id": "12345", "sequence": "1.1", "name": "既存要件の更新例"},
                {"jama_id": "", "sequence": "2", "name": "SYSP: Description編集の例"}
            ]

            self.excel.create_requirement_excel(sample_items, str(output_path))
            
            print(f"\n✅ Excelテンプレートを作成しました。")
            print(f"📄 出力ファイル: {output_path}")

        except Exception as e:
            logger.error(f"テンプレート作成に失敗: {str(e)}", exc_info=True)
            print(f"\n❌ エラーが発生しました: {str(e)}")
            sys.exit(1)

    def update_requirements(self, input_file: str, dry_run: bool = False) -> None:
        """Excelファイルから要件を読み込んでJAMAを更新"""
        # (このメソッドの実装は変更なし)
        pass

def main():
    """メイン処理"""
    parser = argparse.ArgumentParser(
        description='JAMA要件管理ツール',
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""
使用例:
  # 空のテンプレートを作成（JAMAへの接続不要）
  %(prog)s template -o template.xlsx
  
  # プロジェクト全体の要件を取得
  %(prog)s fetch -o requirements.xlsx
  
  # 特定のコンポーネント以下をsequenceで指定して取得 (例: 1)
  %(prog)s fetch -o requirements.xlsx -s 1
  
  # 更新内容をプレビュー（ドライラン）
  %(prog)s update -i requirements.xlsx --dry-run
"""
    )
    
    subparsers = parser.add_subparsers(dest='command', required=True, help='実行するコマンド')
    
    # fetchコマンド
    fetch_parser = subparsers.add_parser('fetch', help='JAMAから要件構造を取得')
    fetch_parser.add_argument('-o', '--output', required=True, help='出力Excelファイル名')
    fetch_parser.add_argument('-s', '--sequence', help='取得開始位置のsequence (例: 6.1.5)')
    fetch_parser.add_argument('-n', '--name', help='取得開始位置のアイテム名 (ルート直下のみ)')
    fetch_parser.add_argument('-d', '--max-depth', type=int, help='起点からの相対的な最大階層数')
    fetch_parser.add_argument('-c', '--config', default='config.json', help='設定ファイルのパス')
    fetch_parser.add_argument('--debug', action='store_true', help='デバッグモードを有効にする')
    
    # updateコマンド
    update_parser = subparsers.add_parser('update', help='Excelファイルから要件を更新')
    update_parser.add_argument('-i', '--input', required=True, help='入力Excelファイル名')
    update_parser.add_argument('--dry-run', action='store_true', help='実際の更新は行わない')
    update_parser.add_argument('-c', '--config', default='config.json', help='設定ファイルのパス')
    
    # templateコマンド
    template_parser = subparsers.add_parser('template', help='空のExcelテンプレートを作成')
    template_parser.add_argument('-o', '--output', required=True, help='出力Excelファイル名')
    
    args = parser.parse_args()
    
    if args.command == 'template':
        tool = JAMATool()
        tool.create_template(output_file=args.output)
    else:
        try:
            config = Config(args.config)
            tool = JAMATool(config)
            
            if args.command == 'fetch':
                tool.fetch_structure(
                    output_file=args.output,
                    component_sequence=args.sequence,
                    component_name=args.name,
                    max_depth=args.max_depth,
                    debug=args.debug
                )
            elif args.command == 'update':
                tool.update_requirements(
                    input_file=args.input,
                    dry_run=args.dry_run
                )
        except Exception as e:
            logger.error(f"処理中に予期せぬエラーが発生しました: {e}", exc_info=True)
            print(f"\n❌ 致命的なエラー: {e}")
            sys.exit(1)

if __name__ == '__main__':
    main()