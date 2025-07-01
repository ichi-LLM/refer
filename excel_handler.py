"""
Excel ファイルの読み書きを処理
要件一覧とDescription編集の2シート構造を管理
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import logging
from typing import List, Dict, Optional, Tuple
from html.parser import HTMLParser
import re
import os

logger = logging.getLogger(__name__)


class HTMLTableParser(HTMLParser):
    """HTMLテーブルをパースするクラス"""
    
    def __init__(self):
        super().__init__()
        self.tables = []
        self.current_table = []
        self.current_row = []
        self.current_cell = ""
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        
    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
            self.current_table = []
        elif tag == "tr" and self.in_table:
            self.in_row = True
            self.current_row = []
        elif tag in ["td", "th"] and self.in_row:
            self.in_cell = True
            self.current_cell = ""
            
    def handle_endtag(self, tag):
        if tag == "table":
            self.in_table = False
            if self.current_table:
                self.tables.append(self.current_table)
        elif tag == "tr" and self.in_row:
            self.in_row = False
            if self.current_row:
                self.current_table.append(self.current_row)
        elif tag in ["td", "th"] and self.in_cell:
            self.in_cell = False
            self.current_row.append(self.current_cell.strip())
            
    def handle_data(self, data):
        if self.in_cell:
            self.current_cell += data


class ExcelHandler:
    """Excelファイルの読み書きハンドラー"""
    
    # カラム定義
    COLUMNS = {
        'A': 'JAMA_ID',
        'B': 'メモ/コメント',
        'C': 'Sequence',
        'D': '階層1',
        'E': '階層2',
        'F': '階層3',
        'G': '階層4',
        'H': '階層5',
        'I': '階層6',
        'J': '階層7',
        'K': '階層8',
        'L': '階層9',
        'M': '階層10',
        'N': '階層11',
        'O': 'アイテムタイプ',
        'P': 'Assignee',
        'Q': 'Status',
        'R': 'Tags',
        'S': 'Reason',
        'T': 'Preconditions',
        'U': 'Target_system',
        'V': '現在のDescription',
        'W': '要件更新',
        'X': '新Description参照'
    }
    
    # Description編集シートの列幅（(a)～(d)の列数）
    DESC_COLS = {
        'a': 1,      # (a)Trigger action: 1列
        'b': 64,     # (b)Behavior of ego-vehicle: 64列
        'c': 10,     # (c)HMI: 10列
        'd': 5       # (d)Other: 5列
    }
    
    def __init__(self, config=None):
        """初期化"""
        self.wb = None
        self.requirement_sheet = None
        self.description_sheet = None
        self.sequence_index = {}  # sequence検索用インデックス
        self.sysp_template_map = {}  # SYSPテンプレートのマッピング
        
        # パフォーマンス設定
        if config:
            # max_description_templates は使用しない（すべてのSYSPにテンプレート作成）
            self.column_width_check_rows = config.column_width_check_rows
        else:
            self.column_width_check_rows = 100
        
    def create_requirement_excel(self, items: List[Dict], output_file: str) -> None:
        """
        要件一覧をExcelファイルに出力
        
        Args:
            items: 要件アイテムリスト
            output_file: 出力ファイル名
        """
        logger.info(f"Excelファイル作成開始: {output_file}")
        logger.info(f"処理対象アイテム数: {len(items)}件")
        
        # 新規ワークブック作成
        logger.info("新規ワークブック作成中...")
        self.wb = Workbook()
        
        # 最初にDescription_editシートを作成（マッピング情報を確立）
        logger.info("Description編集シート作成開始...")
        self.description_sheet = self.wb.create_sheet("Description_edit")
        self._create_description_sheet(items)
        
        # その後でRequirement_of_Driverシートを作成
        logger.info("要件一覧シート作成開始...")
        self.requirement_sheet = self.wb.active
        self.requirement_sheet.title = "Requirement_of_Driver"
        self._create_requirement_sheet(items)
        
        # 保存
        logger.info("ファイル保存中...")
        self.wb.save(output_file)
        logger.info(f"Excelファイル作成完了: {output_file}")
        
    def _create_requirement_sheet(self, items: List[Dict]) -> None:
        """要件一覧シートを作成"""
        ws = self.requirement_sheet
        
        # ヘッダー行作成
        logger.info("ヘッダー行作成中...")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for col, header in self.COLUMNS.items():
            cell = ws[f"{col}1"]
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # sequence検索用インデックスを作成（高速化のため）
        logger.info("階層情報インデックス作成中...")
        self.sequence_index = {item.get("sequence", ""): item for item in items if item.get("sequence")}
        logger.info(f"インデックス作成完了: {len(self.sequence_index)}件")
            
        # データ行作成
        total_items = len(items)
        logger.info(f"データ行作成開始: {total_items}件")
        
        # SYSPカウンター初期化
        self._sysp_count = 0
        
        row_num = 2
        for idx, item in enumerate(items, 1):
            # 進捗表示（100件ごと、または最後）
            if idx % 100 == 0 or idx == total_items:
                logger.info(f"要件シート作成進捗: {idx}/{total_items} ({idx/total_items*100:.1f}%)")
                
            # 階層情報を解析（高速化版）
            hierarchy = self._parse_hierarchy_fast(item)
            
            # 基本情報
            ws[f"A{row_num}"] = item.get("jama_id", "")
            ws[f"B{row_num}"] = ""  # メモ/コメント欄（空欄）
            ws[f"C{row_num}"] = item.get("sequence", "")
            
            # 階層1～11
            for i, level in enumerate(hierarchy, 1):
                if i <= 11:
                    ws[f"{get_column_letter(3 + i)}{row_num}"] = level
                    
            # その他の情報
            ws[f"O{row_num}"] = "Requirement"  # デフォルト値
            ws[f"P{row_num}"] = item.get("assignee", "")
            ws[f"Q{row_num}"] = item.get("status", "")
            ws[f"R{row_num}"] = item.get("tags", "")
            ws[f"S{row_num}"] = item.get("reason", "")
            ws[f"T{row_num}"] = item.get("preconditions", "")
            ws[f"U{row_num}"] = item.get("target_system", "")
            
            # SYSPアイテムのカスタムフィールドをログに出力（デバッグ用）
            if "SYSP" in item.get("name", ""):
                custom_fields = {k: v for k, v in item.items() if k.startswith("custom_")}
                # 最初の3件のSYSPアイテムのみカスタムフィールドを表示
                sysp_count = getattr(self, '_sysp_count', 0)
                if custom_fields and sysp_count < 3:
                    logger.info(f"SYSP Item #{sysp_count + 1} custom fields: {list(custom_fields.keys())}")
                    self._sysp_count = sysp_count + 1
            
            # Description関連
            description = item.get("description", "")
            if description:
                # HTMLテーブルを簡易表示
                ws[f"V{row_num}"] = self._extract_table_preview(description)
                
                # SYSPのDescriptionがある場合は編集リンクを作成
                if "SYSP" in item.get("name", ""):
                    ws[f"W{row_num}"] = ""  # デフォルトは空欄（更新しない）
                    # sysp_template_mapから正しいテンプレート位置を取得
                    if hasattr(self, 'sysp_template_map') and (idx - 1) in self.sysp_template_map:
                        template_row = self.sysp_template_map[idx - 1]
                        ws[f"X{row_num}"] = f"編集画面へ"
                        # 要件名が見えるように、1行上にリンク
                        ws[f"X{row_num}"].hyperlink = f"#Description_edit!A{template_row - 1}"
                        ws[f"X{row_num}"].font = Font(color="0000FF", underline="single")
                else:
                    ws[f"W{row_num}"] = ""  # 非SYSPも空欄
            else:
                ws[f"W{row_num}"] = ""  # Descriptionがない場合も空欄
                    
            row_num += 1
            
        logger.info("要件シート作成完了")
        
        # 列幅調整
        logger.info("列幅調整中...")
        self._adjust_column_widths(ws)
        logger.info("列幅調整完了")
        
    def _create_description_sheet(self, items: List[Dict]) -> None:
        """Description編集シートを作成"""
        ws = self.description_sheet
        
        # スタイル定義
        header_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # SYSPアイテムをフィルタリング
        logger.info("SYSPアイテムをフィルタリング中...")
        sysp_items = []
        for idx, item in enumerate(items):
            # nameフィールドにSYSPを含むかチェック
            if "SYSP" in item.get("name", ""):
                sysp_items.append((item, idx))
        
        logger.info(f"SYSPアイテム数: {len(sysp_items)}件")
        
        if not sysp_items:
            logger.info("SYSPアイテムが見つかりませんでした")
            ws["A1"] = "SYSPアイテムが見つかりませんでした"
            return
            
        # すべてのSYSPアイテムにテンプレートを作成（制限なし）
        current_row = 10
        total_sysp = len(sysp_items)
        
        logger.info(f"Descriptionテンプレート作成開始: {total_sysp}件")
        
        # SYSPアイテムのインデックスを保存（リンク作成用）
        self.sysp_template_map = {}  # row_index -> template_start_row
        
        for idx, (item, original_idx) in enumerate(sysp_items, 1):
            # 進捗表示
            if idx % 100 == 0 or idx == total_sysp:
                logger.info(f"Descriptionテンプレート作成進捗: {idx}/{total_sysp} ({idx/total_sysp*100:.1f}%)")
                
            # マッピング情報を保存
            self.sysp_template_map[original_idx] = current_row
            
            # ヘッダー行（【JAMA_ID】要件名の形式で表示）
            jama_id = item.get('jama_id', '新規')
            name = item.get("name", "")
            ws[f"A{current_row - 2}"] = f"【{jama_id}】{name}"
            ws[f"A{current_row - 2}"].font = Font(bold=True)
            ws.merge_cells(f"A{current_row - 2}:J{current_row - 2}")
            
            # 要件名を表示（リンク先の1つ上）
            ws[f"A{current_row - 1}"] = name
            ws[f"A{current_row - 1}"].font = Font(bold=True, size=12)
            ws.merge_cells(f"A{current_row - 1}:J{current_row - 1}")
            
            # 一覧に戻るリンク
            ws[f"K{current_row - 1}"] = "一覧に戻る"
            ws[f"K{current_row - 1}"].hyperlink = f"#Requirement_of_Driver!A{original_idx + 2}"
            ws[f"K{current_row - 1}"].font = Font(color="0000FF", underline="single")
            
            # 既存のDescriptionをパース
            existing_data = None
            if item.get("description"):
                existing_data = self._parse_existing_description_table(item.get("description"))
                if existing_data:
                    logger.info(f"既存のDescriptionテーブルを検出: JAMA_ID={jama_id}")
            
            # 5行テーブルのテンプレート作成（既存データがあれば渡す）
            self._create_description_template(ws, current_row, existing_data)
            
            # 「【参考】現在のDescription」セクションは削除
            
            current_row += 16  # 次のテンプレートまでの間隔
            
        logger.info(f"Descriptionテンプレート作成完了: {total_sysp}件")
                    
    def _create_description_template(self, ws, start_row: int, existing_data: Optional[List[List[str]]] = None) -> None:
        """
        5行形式のDescriptionテンプレートを作成
        
        Args:
            ws: ワークシート
            start_row: 開始行
            existing_data: 既存のテーブルデータ（5行のリスト）
        """
        # 列の定義
        col_idx = 1
        
        # スタイル
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        
        # 1行目: I/O Type
        ws.cell(row=start_row, column=1, value="I/O Type").border = border
        ws.cell(row=start_row, column=2, value="IN").border = border
        ws.cell(row=start_row, column=3, value="OUT").border = border
        # OUTを最後まで結合
        ws.merge_cells(f"C{start_row}:CF{start_row}")
        
        # 2行目: 項目名
        ws.cell(row=start_row + 1, column=1, value="").border = border
        ws.cell(row=start_row + 1, column=2, value="(a)Trigger action").border = border
        
        # (b)Behavior of ego-vehicle (64列)
        b_start = 3
        b_end = b_start + self.DESC_COLS['b'] - 1
        ws.cell(row=start_row + 1, column=b_start, value="(b)Behavior of ego-vehicle").border = border
        if b_end > b_start:
            ws.merge_cells(f"{get_column_letter(b_start)}{start_row + 1}:{get_column_letter(b_end)}{start_row + 1}")
            
        # (c)HMI (10列)
        c_start = b_end + 1
        c_end = c_start + self.DESC_COLS['c'] - 1
        ws.cell(row=start_row + 1, column=c_start, value="(c)HMI").border = border
        if c_end > c_start:
            ws.merge_cells(f"{get_column_letter(c_start)}{start_row + 1}:{get_column_letter(c_end)}{start_row + 1}")
            
        # (d)Other (5列)
        d_start = c_end + 1
        d_end = d_start + self.DESC_COLS['d'] - 1
        ws.cell(row=start_row + 1, column=d_start, value="(d)Other").border = border
        if d_end > d_start:
            ws.merge_cells(f"{get_column_letter(d_start)}{start_row + 1}:{get_column_letter(d_end)}{start_row + 1}")
            
        # 3-5行目: Data Name, Data Label, Data
        data_rows = ["Data Name", "Data Label", "Data"]
        for i, row_name in enumerate(data_rows, 2):
            ws.cell(row=start_row + i, column=1, value=row_name).border = border
            # 入力セルを配置
            for col in range(2, d_end + 1):
                cell = ws.cell(row=start_row + i, column=col)
                
                # 既存データがある場合は値を設定
                if existing_data and i >= 2 and col <= len(existing_data[i]):
                    # existing_dataのインデックスは0ベース、iは2から始まる
                    cell.value = existing_data[i][col - 1]
                else:
                    cell.value = ""
                    
                cell.border = border
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                
    def _parse_hierarchy_fast(self, item: Dict) -> List[str]:
        """
        アイテムの階層構造を高速に解析（インデックス使用）
        
        Args:
            item: 対象アイテム
            
        Returns:
            階層名のリスト
        """
        hierarchy = []
        sequence = item.get("sequence", "")
        
        if not sequence:
            return [item.get("name", "")]
            
        # sequence の各レベルでアイテムを探す（インデックス使用で高速化）
        parts = sequence.split(".")
        for i in range(1, len(parts) + 1):
            current_seq = ".".join(parts[:i])
            # インデックスから直接取得
            if current_seq in self.sequence_index:
                hierarchy.append(self.sequence_index[current_seq].get("name", ""))
                    
        return hierarchy
        
    def _parse_hierarchy(self, item: Dict, all_items: List[Dict]) -> List[str]:
        """
        アイテムの階層構造を解析
        
        Args:
            item: 対象アイテム
            all_items: 全アイテムリスト
            
        Returns:
            階層名のリスト
        """
        hierarchy = []
        sequence = item.get("sequence", "")
        
        if not sequence:
            return [item.get("name", "")]
            
        # sequence の各レベルでアイテムを探す
        parts = sequence.split(".")
        for i in range(1, len(parts) + 1):
            current_seq = ".".join(parts[:i])
            # 該当するアイテムを探す
            for other_item in all_items:
                if other_item.get("sequence") == current_seq:
                    hierarchy.append(other_item.get("name", ""))
                    break
                    
        return hierarchy
        
    def _extract_table_preview(self, html_description: str) -> str:
        """
        HTMLテーブルから簡易プレビューを作成
        
        Args:
            html_description: HTML形式のDescription
            
        Returns:
            プレビュー文字列
        """
        if not html_description:
            return ""
            
        # HTMLパーサーでテーブルを抽出
        parser = HTMLTableParser()
        parser.feed(html_description)
        
        if parser.tables:
            # 最初のテーブルの最初の3行を表示
            table = parser.tables[0]
            preview_rows = table[:3]
            preview = []
            for row in preview_rows:
                preview.append(" | ".join(row[:4]))  # 最初の4列のみ
            return "\n".join(preview)
            
        # テーブルがない場合はテキストを抽出
        text = re.sub('<[^<]+?>', '', html_description)
        return text[:100] + "..." if len(text) > 100 else text
        
    def _parse_existing_description_table(self, html_description: str) -> Optional[List[List[str]]]:
        """
        既存のHTML Descriptionテーブルをパースして値を取得
        
        Args:
            html_description: HTML形式のDescription
            
        Returns:
            5行×81列のテーブルデータ、または None（形式が異なる場合）
        """
        if not html_description or '<table' not in html_description:
            return None
            
        # HTMLパーサーでテーブルを抽出
        parser = HTMLTableParser()
        parser.feed(html_description)
        
        if not parser.tables:
            return None
            
        table = parser.tables[0]
        
        # テーブル構造の検証
        if len(table) != 5:
            return None
            
        # 固定ヘッダーの検証
        # 1行目
        if len(table[0]) < 3:
            return None
        if table[0][0] != "I/O Type" or table[0][1] != "IN" or table[0][2] != "OUT":
            return None
            
        # 2行目
        if len(table[1]) < 5:
            return None
        if table[1][0] != "" or table[1][1] != "(a)Trigger action":
            return None
        if "(b)Behavior of ego-vehicle" not in table[1][2]:
            return None
        # 2行目の後半のヘッダーは、HTMLパーサーの結果によって位置が異なる可能性があるため
        # より柔軟な検証にする
        row2_text = " ".join(table[1])
        if "(c)HMI" not in row2_text or "(d)Other" not in row2_text:
            return None
            
        # 3-5行目
        if table[2][0] != "Data Name" or table[3][0] != "Data Label" or table[4][0] != "Data":
            return None
            
        # 構造が正しいことを確認したら、値を返す
        return table
    
    def _adjust_column_widths(self, ws) -> None:
        """列幅を自動調整"""
        total_columns = len(self.COLUMNS)
        logger.info(f"列幅自動調整開始: {total_columns}列")
        
        for idx, column in enumerate(ws.columns, 1):
            if idx > total_columns:
                break
                
            # 進捗表示
            if idx % 5 == 0 or idx == total_columns:
                logger.info(f"列幅調整進捗: {idx}/{total_columns}")
                
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            # 設定値に基づいてチェックする行数を制限（高速化のため）
            check_rows = min(self.column_width_check_rows, len(column))
            for cell in list(column)[:check_rows]:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
                    
            adjusted_width = min(max_length + 2, 50)  # 最大幅を50に制限
            ws.column_dimensions[column_letter].width = adjusted_width
            
    def read_requirement_excel(self, input_file: str) -> List[Dict]:
        """
        Excelファイルから要件データを読み込み

        Args:
            input_file: 入力ファイル名

        Returns:
            要件データリスト
        """
        logger.info(f"Excelファイル読み込み開始: {input_file}")
    
        try:
            # ファイルサイズをチェック
            file_size = os.path.getsize(input_file) / 1024  # KB単位
            logger.info(f"ファイルサイズ: {file_size:.1f}KB")
        
            # 読み込み開始メッセージ
            print(f"\nExcelファイルを開いています... (ファイルサイズ: {file_size:.1f}KB)")
            print("5MB程度のファイルで約2分かかる場合があります。しばらくお待ちください...")
        
            # ファイルを読み込み
            wb = openpyxl.load_workbook(input_file)
            print("Excelファイルを開きました。データを処理しています...")
        
            requirement_sheet = wb["Requirement_of_Driver"]
            description_sheet = wb["Description_edit"] if "Description_edit" in wb.sheetnames else None
        
            requirements = []
            total_rows = requirement_sheet.max_row - 1  # ヘッダー行を除く
        
            logger.info(f"総行数: {total_rows}")
        
            # ヘッダー行をスキップして、データ行を読み込み
            for row_idx, row_num in enumerate(range(2, requirement_sheet.max_row + 1), 1):
                # 進捗表示（100行ごと、最初と最後）
                if row_idx == 1 or row_idx % 100 == 0 or row_idx == total_rows:
                    logger.info(f"Excelファイル読み込み進捗: {row_idx}/{total_rows} ({row_idx/total_rows*100:.1f}%)")
            
                # JAMA_IDとW列（要件更新）を確認
                jama_id = requirement_sheet[f"A{row_num}"].value
                update_flag = requirement_sheet[f"W{row_num}"].value
                memo_comment = requirement_sheet[f"B{row_num}"].value
            
                # 削除操作の判定（B列に「削除」と記載）- 最優先
                if str(memo_comment).strip() == "削除" and jama_id:
                    operation = "削除"
                # 処理対象の判定
                elif not jama_id:
                    # JAMA_IDが空欄の場合は新規作成
                    operation = "新規"
                elif update_flag == "する":
                    operation = "更新"
                else:
                    # JAMA_IDはあるが更新フラグが「する」でない場合はスキップ
                    continue
                
                # 基本情報を読み込み（空欄は除外）
                item = {
                    "operation": operation,
                    "sequence": requirement_sheet[f"C{row_num}"].value,
                    "name": requirement_sheet[f"N{row_num}"].value or requirement_sheet[f"M{row_num}"].value,  # 階層11または10
                }
            
                # JAMA IDがある場合のみ設定
                if jama_id:
                    item["jama_id"] = jama_id
            
                # その他のフィールドは、値がある場合のみ設定
                field_mapping = {
                    "P": "assignee",
                    "Q": "status",
                    "R": "tags",
                    "S": "reason",
                    "T": "preconditions",
                    "U": "target_system"
                }
            
                for col, field_name in field_mapping.items():
                    value = requirement_sheet[f"{col}{row_num}"].value
                    if value is not None and value != "":
                        item[field_name] = value
            
                # Description更新チェック
                if update_flag == "する" and description_sheet:
                    # 新しいDescriptionを読み込み
                    desc_ref = requirement_sheet[f"X{row_num}"].value
                    if desc_ref:
                        # Description_editシートから読み込み
                        new_description = self._read_description_from_sheet(
                            description_sheet, 
                            item.get("jama_id", "新規")
                        )
                        if new_description:
                            item["description"] = new_description
                        
                requirements.append(item)
            
            logger.info(f"読み込み完了: {len(requirements)}件")
            logger.info(f"  新規作成: {len([r for r in requirements if r['operation'] == '新規'])}件")
            logger.info(f"  更新: {len([r for r in requirements if r['operation'] == '更新'])}件")
            logger.info(f"  削除: {len([r for r in requirements if r['operation'] == '削除'])}件")
        
            return requirements
        
        except Exception as e:
            logger.error(f"Excelファイル読み込みエラー: {str(e)}")
            raise
            
    def _read_description_from_sheet(self, ws, jama_id: str) -> Optional[str]:
        """
        Description編集シートから新しいDescriptionを読み込み
        
        Args:
            ws: Description_editワークシート
            jama_id: JAMA ID
            
        Returns:
            HTML形式のDescription
        """
        # シート内で【JAMA_ID】のパターンを検索
        search_pattern = f"【{jama_id}】"
        
        for row in range(1, ws.max_row + 1):
            cell_value = ws[f"A{row}"].value
            if cell_value and search_pattern in str(cell_value):
                # テーブル開始行を特定（【JAMA_ID】要件名の2行下）
                table_start = row + 2
                
                # 5行分のデータを読み込み
                table_data = []
                for i in range(5):
                    row_data = []
                    for col in range(1, 87):  # A～CF列（1+1+64+10+5 = 81列 + 余裕）
                        value = ws.cell(row=table_start + i, column=col).value
                        # 0も正しく表示されるように修正
                        row_data.append(str(value) if value is not None else "")
                    table_data.append(row_data)
                    
                # HTMLテーブルに変換
                return self._convert_to_html_table(table_data)
                
        return None
        
    def _convert_to_html_table(self, table_data: List[List[str]]) -> str:
        """
        テーブルデータをHTMLに変換
        
        Args:
            table_data: テーブルデータ
            
        Returns:
            HTMLテーブル
        """
        html = "<table border='1' cellpadding='5' cellspacing='0'>\n"
        
        for row_idx, row in enumerate(table_data):
            html += "<tr>\n"
            
            # 特殊な結合処理
            if row_idx == 0:  # I/O Type行
                html += f"<td>{row[0]}</td>\n"
                # IN列に薄い青の背景色
                html += f"<td style='background-color: #E3F2FD;'>{row[1]}</td>\n"
                # OUT列に薄い緑の背景色
                html += f"<td colspan='84' style='background-color: #E8F5E9;'>{row[2]}</td>\n"
            elif row_idx == 1:  # 項目名行
                html += f"<td>{row[0]}</td>\n"
                # (a)Trigger actionに薄い青の背景色
                html += f"<td style='background-color: #E3F2FD;'>{row[1]}</td>\n"
                # (b)Behavior of ego-vehicleに薄い緑の背景色
                html += f"<td colspan='64' style='background-color: #E8F5E9;'>{row[2]}</td>\n"
                # (c)HMIに薄い緑の背景色
                html += f"<td colspan='10' style='background-color: #E8F5E9;'>{row[66]}</td>\n"
                # (d)Otherに薄い緑の背景色
                html += f"<td colspan='5' style='background-color: #E8F5E9;'>{row[76]}</td>\n"
            else:  # データ行（色なし）
                for cell in row[:81]:  # 必要な列数のみ
                    html += f"<td>{cell}</td>\n"
                    
            html += "</tr>\n"
            
        html += "</table>"
        return html