import pandas as pd
from openpyxl import load_workbook
import os

class InputExcelFile:
    """
    Excelファイルを読み込み、条件に一致するデータを抽出するクラス
    Args:
    input_file_path (str): 入力Excelファイルのパス
    input_sheet_name (str): 入力Excelファイルのシート名
    filter_conditions_dict (dict): フィルタ条件を指定する辞書　列名:一致する値
    sort_column (str): ソートに使用する列名(デフォルトはNone指定なし)
    input_start_row (int, optional): データの開始行（デフォルトは0）
    """
    def __init__(self, input_file_path, input_sheet_name,filter_conditions_dict,sort_column=None,input_start_row=0):
        #出力用データフレーム
        self.df = None

        #加工用
        self.input_file_path = input_file_path
        self.input_sheet_name = input_sheet_name
        self.filter_conditions_dict = filter_conditions_dict
        self.sort_column = sort_column
        self.input_start_row = input_start_row

        self.main()

    def read_excel(self):
        return pd.read_excel(self.input_file_path, sheet_name=self.input_sheet_name,header=self.input_start_row)
    
    # keyを列名としてvalueと一致するものをフィルタリングする
    def filter_data_dict(self, df, filter_dict):
        for key, value in filter_dict.items():
            df = df[df[key] == value]
        return df
    
    # ソートする列名を指定
    def sort_data(self, df, sort_column):
        if sort_column is None:
            return df
        return df.sort_values(sort_column)
    
    #書き込む列を絞り込む
    def select_columns(self, df):
        return df[['管理番号','検査カテゴリ', '検査結果','行番号','コメント','対象ソースコード','修正ソースコード']]

    def main(self):
        print('反映元:',self.input_file_path)
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
        # 抽出したdf
        self.df = sorted_df

if __name__ == '__main__':
    input_file_path = input('ファイルパスを入力してください')
    input_sheet_name = '検査'
    condition = {'管理番号':'NUL0001','検査カテゴリ':'リンク'}
    #0始まり？
    start_row = 12
    input_excel = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)
    print(input_excel.df.head)