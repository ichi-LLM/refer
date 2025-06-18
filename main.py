import configparser
import argparse
from jama_client import JamaClient
from excel_manager import ExcelManager
from requirement_processor import RequirementProcessor

def main():
    # --- コマンドライン引数の設定 ---
    parser = argparse.ArgumentParser(description="JAMAとExcelを連携させるツール")
    parser.add_argument('action', choices=['get', 'update', 'template'], help="実行する操作を選択: 'get' (JAMA->Excel), 'update' (Excel->JAMA), 'template' (空のExcel作成)")
    parser.add_argument('--file', required=True, help="使用するExcelファイル名 (例: reqs.xlsx)")
    parser.add_argument('--component', type=int, help="取得対象を特定のコンポーネント配下に限定する場合、そのコンポーネントのJAMA IDを指定")
    args = parser.parse_args()

    # --- 設定ファイルの読み込み ---
    config = configparser.ConfigParser()
    try:
        config.read('config.ini', encoding='utf-8')
        jama_config = config['JAMA']
        proxy_config = config['PROXY']
        field_mapping = config['FIELD_MAPPING']
    except Exception as e:
        print(f"エラー: config.iniの読み込みに失敗しました。ファイルが存在し、形式が正しいか確認してください。 {e}")
        return

    # --- クラスの初期化 ---
    proxies = {
        'http': proxy_config.get('http_proxy') or None,
        'https': proxy_config.get('https_proxy') or None,
    }
    
    try:
        client = JamaClient(
            base_url=jama_config.get('base_url'),
            client_id=jama_config.get('client_id'),
            client_secret=jama_config.get('client_secret'),
            proxies=proxies
        )
        processor = RequirementProcessor(client, field_mapping)
        excel_manager = ExcelManager(args.file)

        # --- アクションの実行 ---
        if args.action == 'get':
            print("JAMAからデータを取得し、Excelに出力します。")
            processor.jama_to_excel(excel_manager, jama_config.getint('project_id'), args.component)
        
        elif args.action == 'update':
            print("Excelのデータを読み込み、JAMAに反映します。")
            processor.excel_to_jama(excel_manager, jama_config.getint('project_id'))

        elif args.action == 'template':
            excel_manager.create_template()
            
    except Exception as e:
        print(f"処理中に予期せぬエラーが発生しました: {e}")


if __name__ == '__main__':
    main()