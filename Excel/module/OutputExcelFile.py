import pandas as pd
import openpyxl
import datetime
from module.InputExcelFile import InputExcelFile
from module.BulkInfo import BulkInfo
import shutil

class OutputExcelFile:
    """
    新しいExcelファイルに保存するクラス
    output_file_name (str): 出力Excelファイルの名前
    output_sheet_name (str): 出力Excelファイルのシート名
    """
    def __init__(self, df,management_id,management_number,category,bulk_excel_file_path):
        self.data_list = self.df_to_list(df)
        self.bulk_excel_file_path = bulk_excel_file_path
        self.wb = self.load_bulk_excel()
        self.all_ws = self.wb['検査結果(ページ単位)']
        self.individual_ws = self.wb['検査結果(検査箇所単位)']
        self.copy_all_ws = self.copy_ws(self.all_ws)
        self.copy_individual_ws = self.copy_ws(self.individual_ws)

        self.main()
    #一括登録ファイルを取得
    def load_bulk_excel(self):
        wb = openpyxl.load_workbook(self.bulk_excel_file_path)
        return wb
    # 一括登録ファイルのシートをコピーする
    def copy_ws(self,ws):
        ws_copy = self.wb.copy_worksheet(ws)
        ws_copy.title = '_' + ws.title
        # コピーしたシートを表紙シートのあとに移動
        self.wb._sheets.remove(ws_copy)
        self.wb._sheets.append(ws_copy)
        return ws_copy

    # new_bulk_excel_file_pathにdfを書き込む
    def write_excel(self):
        #2行目から書き込む
        all_index = 2
        individual_index = 2
        for row in self.data_list:
            if(row[3] == 1):
                self.write_cell(self.all_ws,all_index,row)
                all_index += 1
            else:
                self.write_cell(self.individual_ws,individual_index,row)
                individual_index += 1
        self.wb.save(self.bulk_excel_file_path)
    # write_excelのサブモジュール
    # 引数にワークシートオブジェクト
    # ワークシートオブジェクトのセルに値を代入
    def write_cell(self,ws,index,row):
        #コメント列が結合セルか判定
        is_merged = self.is_merged_cell_not_top_left(ws[index][6],ws)
        #検査結果は結合されないため分岐から除外
        ws[index][3].value = row[2]

        #/rを削除
        comment = self.trim_r(row[4]) if row[4] is not None else ''
        target = self.trim_r(row[5]) if row[5] is not None else ''
        ammend = self.trim_r(row[6]) if row[6] is not None else ''


        if (not(is_merged)):
            # 結合セルではない場合はそのまま代入
            ws[index][5].value = comment
            ws[index][6].value = target
            ws[index][7].value = ammend
        else:
        #結合セルの扱い
            if(is_merged == index):
                #結合セルの最初の行
                ws[is_merged][5].value = ""
                #対象ソースコードは先頭のもののみ代入
                ws[index][6].value = target
                ws[is_merged][7].value = ""
            
            if(row[2]=="いいえ"):
                #検査結果がいいえの場合のみ検査.xlsmの内容を反映
                new_comment = comment if pd.notna(comment) else ""
                ws[is_merged][5].value = "\n".join([str(ws[is_merged][5].value),new_comment]).strip()
                new_ammend = ammend if pd.notna(ammend) else ""
                ws[is_merged][7].value = "\n".join([str(ws[is_merged][7].value),new_ammend]).strip()


            #対象ソースコードは結合不要なため除外
            #ws[is_merged][6].value += target
    def trim_r(self,cell_value):
        # 値に改行コード(_x000D_)が含まれている場合は削除
        _value = cell_value
        if isinstance(cell_value, str):
            _value = cell_value.replace('_x000D_', '')
        return _value

    # 結合セルであり結合の最初のセルではない場合は最初の行番号を返す
    def is_merged_cell_not_top_left(self,cell, ws):
        for merged_range in ws.merged_cells.ranges:
            # cell.coordinateはセルのアドレス
            if cell.coordinate in merged_range:
                return merged_range.min_row  # 結合セルの場合は最初の行番号を返す
        return False  # 結合セルでない場合
    
    # dfをopenpyxlで扱えるようにリストに変換
    def df_to_list(self,df):
        data_list = df.values.tolist()
        return data_list

    def main(self):
        
        # Excelファイルの書き込み
        self.write_excel()

if __name__ == '__main__':
    bulk_path="C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/本物1867_NUL0000_リンク_検査結果一括更新_20250101163437.xlsx"
    bulk_info = BulkInfo(bulk_path)

    # テスト用
    input_file_path = input('ファイルパスを入力してください')
    input_sheet_name = '検査'
    condition = {'管理番号':'NUL0000','検査カテゴリ':'リンク'}
    start_row = 12
    input_excel = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)
    input_excel_file = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)


    OutputExcelFile(input_excel_file.df,
                         bulk_info.management_id,
                         bulk_info.management_number,
                         bulk_info.category,
                         bulk_path
                         )
    