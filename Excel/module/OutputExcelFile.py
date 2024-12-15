import pandas as pd
from openpyxl import load_workbook
import datetime
from module.InputExcelFile import InputExcelFile
from module.BulkInfo import BulkInfo

class OutputExcelFile:
    """
    新しいExcelファイルに保存するクラス
    output_file_name (str): 出力Excelファイルの名前
    output_sheet_name (str): 出力Excelファイルのシート名
    """
    def __init__(self, df,output_folder_path,management_id,alphabet,management_number,category,output_sheet_name):
        self.df = df
        self.output_file_name = self.getfile_name(management_id,alphabet,management_number,category)
        self.output_file_path = output_folder_path + self.output_file_name
        self.output_sheet_name = output_sheet_name

        self.main()
    
    def getfile_name(self,management_id,alphabet,management_number,category):
        date_str = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        file_name = management_id + '_' + alphabet+'_'+management_number + '_' + category + '_検査結果一括更新_' + date_str+'.xlsx'
        return file_name

    def write_excel(self, df):
        with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=self.output_sheet_name, index=False)

    def main(self):
        # Excelファイルの書き込み
        self.write_excel(self.df)

if __name__ == '__main__':
    bulk_path='C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/0000_AAA_01_リンク_検査結果一括更新_20241206122532.xlsx'
    bulk_info = BulkInfo(bulk_path)

    # テスト用
    input_file_path = input('ファイルパスを入力してください')
    input_sheet_name = '検査'
    condition = {'管理番号':'AAA0001','検査カテゴリ':'リンク'}
    start_row = 3
    input_excel = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)
    input_excel_file = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)

    # テスト用
    folder_path = 'C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/'
    OutputExcelFile(input_excel_file.df,folder_path,bulk_info.management_id,bulk_info.management_alphabet,bulk_info.management_number,bulk_info.category,'テスト')