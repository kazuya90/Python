import pandas as pd
from openpyxl import load_workbook
import os

class ExcelToExcel:
    """
    Excelファイルを読み込み、条件に一致するデータを抽出して新しいExcelファイルに保存するクラス
    Args:
    input_file_path (str): 入力Excelファイルのパス
    input_sheet_name (str): 入力Excelファイルのシート名
    output_file_name (str): 出力Excelファイルの名前
    output_sheet_name (str): 出力Excelファイルのシート名
    filter_conditions_dict (dict): フィルタ条件を指定する辞書　列名:一致する値
    sort_column (str): ソートに使用する列名(デフォルトはNone指定なし)
    input_start_row (int, optional): データの開始行（デフォルトは0）
    """
    def __init__(self, input_file_path, input_sheet_name, output_file_name, output_sheet_name,filter_conditions_dict,sort_column=None,input_start_row=0):
        self.input_file_path = input_file_path
        self.input_sheet_name = input_sheet_name
        self.output_file_path = self.get_dirname(self.input_file_path)+output_file_name
        self.output_sheet_name = output_sheet_name
        self.filter_conditions_dict = filter_conditions_dict
        self.sort_column = sort_column
        self.input_start_row = input_start_row

    # pathからファイル名を除いたディレクトリを取得
    def get_dirname(self, path):
        return os.path.dirname(path)+"/"

    def read_excel(self):
        return pd.read_excel(self.input_file_path, sheet_name=self.input_sheet_name,header=self.input_start_row)

    def write_excel(self, df):
        with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=self.output_sheet_name, index=False)

    def add_sheet_to_existing_excel(self, df):
    # 既存のワークブックを読み込む
        book = load_workbook(self.output_file_path)
    
        # ExcelWriterに既存のワークブックを渡す
        with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
            writer.book = book
            df.to_excel(writer, sheet_name=self.output_sheet_name, index=False)
            
            # 保存
            writer.save()
    
    # keyを列名としてvalueと一致するものをフィルタリングする
    def filter_data_dict(self, df, filter_dict):
        for key, value in filter_dict.items():
            df = df[df[key] == value]
        return df

    # def process_data(self, df):
    #     df.loc[:, '年齢'] = df['年齢'] * 2
    #     return df
    
    # ソートする列名を指定
    def sort_data(self, df, sort_column):
        if sort_column is None:
            return df
        return df.sort_values(sort_column)
    
    # 書き込む列を絞り込む
    def select_columns(self, df):
        return df[['都道府県', '年齢']]

    def main(self):
        # Excelファイルの読み込み
        df = self.read_excel()
        # 条件によるフィルタリング
        filtered_df = self.filter_data_dict(df,self.filter_conditions_dict)
        # 加工
        # processed_df = self.process_data(filtered_df)
        # 列の絞り込み
        selected_df = self.select_columns(filtered_df)
        # ソート
        sorted_df = self.sort_data(selected_df, self.sort_column)
        # Excelファイルの書き込み
        self.write_excel(sorted_df)

        # self.add_sheet_to_existing_excel(processed_df)
        print("フィルタリングされたデータが新しいExcelファイルに保存されました。")

if __name__ == '__main__':
    input_file_path = 'C:/Users/user/OneDrive/プロジェクト/Python/excel/test_data/input.xlsx'
    input_sheet_name = 'Sheet'
    output_file_name = '__output.xlsx'
    output_sheet_name = 'Sheet0'
    filter_conditions_dict = {'都道府県':'福井県'}
    sort_column = '年齢'
    start_row = 2

    excel_to_excel = ExcelToExcel(input_file_path, input_sheet_name, output_file_name, output_sheet_name,filter_conditions_dict,sort_column,start_row)
    excel_to_excel.main()