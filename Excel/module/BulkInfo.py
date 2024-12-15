import pandas as pd
from openpyxl import load_workbook
import os

class BulkInfo:
    """
    一括登録シートの情報を取得するクラス
    bulk_file_paths (tuple): 一括登録シートのファイルパスのタプル
    management_id (str): IDを保持
    management_number (str): 管理番号を保持
    management_alphabet (str): 識別のアルファベットを保持
    management_number (str): ページ番号を保持
    category (str): 検査カテゴリを保持
    """
    def __init__(self,file_path):
        self.bulk_file_path = file_path
        #IDを保持
        self.management_id = str()
        #識別のアルファベットを保持
        self.management_alphabet = str()
        #ページ番号を保持
        self.management_number = str()
        #検査カテゴリを保持
        self.category = str()
        self.set_info()
        #OutputExcelFileに引き渡すための情報
        self.condition:dict = self.get_condition()

    def get_condition(self):
        condition_dict = dict()
        #管理番号を加工
        #condition[1]を四桁の0埋め
        management = self.management_alphabet + self.management_number.zfill(4)
        condition_dict = {'管理番号':management,'検査カテゴリ':self.category}
        return condition_dict
        
    # IDと管理番号と検査カテゴリを各リストに格納
    def set_info(self):
        file_name=self.get_filename(self.bulk_file_path)
        print(file_name)

        #file_nameを'_'で区切る
        file_names = file_name.split('_')

        self.management_id = file_names[0]
        self.management_alphabet = file_names[1]
        self.management_number = file_names[2]
        self.category = file_names[3]

    # pathからファイル名のみを取得
    def get_filename(self, path):
        return os.path.basename(path)

if __name__ == '__main__':
    bulk_path='C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/0000_AAA_01_リンク_検査結果一括更新_20241206122532.xlsx'
    bulk_info = BulkInfo(bulk_path)
    print(bulk_info.management_id)
    print(bulk_info.management_alphabet)
    print(bulk_info.management_number)
    print(bulk_info.category)
    print(bulk_info.condition)