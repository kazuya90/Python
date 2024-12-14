import pandas as pd
from openpyxl import load_workbook
import os
import datetime

class OutputExcelFile:
    """
    新しいExcelファイルに保存するクラス
    output_file_name (str): 出力Excelファイルの名前
    output_sheet_name (str): 出力Excelファイルのシート名
    sort_column (str): ソートに使用する列名(デフォルトはNone指定なし)
    output_start_row (int, optional): データの開始行（デフォルトは0）
    """
    def __init__(self, df,output_file_path,id,alphaget,management_number,category,output_sheet_name):
        self.df = df
        self.output_file_name = self.getfile_name(id,alphaget,management_number,category)
        self.output_file_path = output_file_path + self.output_file_name
        self.output_sheet_name = output_sheet_name

        self.main()
    
    def getfile_name(self,id,alphabet,management_number,category):
        date_str = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        file_name = id + '_' + alphabet+'_'+management_number + '_' + category + '_検査結果一括更新_' + date_str+'.xlsx'
        return file_name

    def write_excel(self, df):
        with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=self.output_sheet_name, index=False)

    def main(self):
        # Excelファイルの書き込み
        self.write_excel(self.df)

        # self.add_sheet_to_existing_excel(processed_df)
        print("フィルタリングされたデータが新しいExcelファイルに保存されました。")

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
        return df[['管理番号','検査カテゴリ', 'コメント']]

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
        # 抽出したdf
        self.df = sorted_df


class BulkRegistrationSheetInput:
    """
    一括登録シートの情報を取得するクラス
    bulk_file_paths (tuple): 一括登録シートのファイルパスのタプル
    ids (list): IDを保持
    management_numbers (list): 管理番号を保持
    management_alphabets (list): 識別のアルファベットを保持
    management_numbers (list): ページ番号を保持
    inspection_categories (list): 検査カテゴリを保持
    """
    def __init__(self,file_paths:tuple):
        self.bulk_file_paths = file_paths
        #IDを保持
        self.ids = list()
        #識別のアルファベットを保持
        self.management_alphabets = list()
        #ページ番号を保持
        self.management_numbers = list()
        #検査カテゴリを保持
        self.inspection_categories = list()
        self.set_info()
        #OutputExcelFileに引き渡すための情報
        self.conditions:list = self.get_conditions()

    def get_conditions(self):
        conditions = list()
        for condition in zip(self.management_alphabets,self.management_numbers,self.inspection_categories):
            #管理番号を加工
            #condition[1]を四桁の0埋め
            management = condition[0] + condition[1].zfill(4)
            condition_dict = {'管理番号':management,'検査カテゴリ':condition[2]}
            conditions.append(condition_dict)
        return conditions
        
    # IDと管理番号と検査カテゴリを各リストに格納
    def set_info(self):
        for file_path in self.bulk_file_paths:
            file_name=self.get_filenames(file_path)
            print(file_name)

            #file_nameを'_'で区切る
            file_names = file_name.split('_')

            self.ids.append(file_names[0])
            self.management_alphabets.append(file_names[1])
            self.management_numbers.append(file_names[2])
            self.inspection_categories.append(file_names[3])

    # pathsからファイル名のみを取得
    def get_filenames(self, paths):
        return os.path.basename(paths)

if __name__ == '__main__':
    bulk_paths=('C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/0000_AAA_01_リンク_検査結果一括更新_20241206122532.xlsx','C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/0000_AAA_03_リンク_検査結果一括更新_20241206122532.xlsx')
    brs = BulkRegistrationSheetInput(bulk_paths)
    input_file_path = input('ファイルパスを入力してください')
    input_sheet_name = '検査'
    start_row = 3
    print(brs.conditions)
    for condition,id,alphabet,number,category in zip(brs.conditions,brs.ids,brs.management_alphabets,brs.management_numbers,brs.inspection_categories):
        input_excel = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)
        print(input_excel.df.head)
        OutputExcelFile(input_excel.df,'C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/',id,alphabet,number,category,'output_sheet_name')