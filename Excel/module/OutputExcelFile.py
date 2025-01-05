import pandas as pd
import openpyxl
import datetime
from .InputExcelFile import InputExcelFile
from .BulkInfo import BulkInfo
import shutil

class OutputExcelFile:
    """
    新しいExcelファイルに保存するクラス
    output_file_name (str): 出力Excelファイルの名前
    output_sheet_name (str): 出力Excelファイルのシート名
    """
    def __init__(self, df,output_folder_path,management_id,management_number,category,bulk_excel_file_path):
        self.data_list = self.df_to_list(df)
        self.output_file_name = self.getfile_name(management_id,management_number,category)
        self.output_file_path = output_folder_path +self.output_file_name
        self.bulk_excel_file_path = bulk_excel_file_path
        self.wb = self.copy_excel()
        self.all_ws = self.wb['検査結果(ページ単位)']
        self.individual_ws = self.wb['検査結果(検査箇所単位)']

        self.main()
    #一括登録ファイルのコピー、コピーしたものに編集する
    def copy_excel(self):
        print(self.output_file_path)
        shutil.copyfile(self.bulk_excel_file_path, self.output_file_path)
        wb = openpyxl.load_workbook(self.output_file_path)
        return wb
    
    def getfile_name(self,management_id,management_number,category):
        date_str = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        file_name = management_id +'_'+management_number + '_' + category + '_検査結果一括更新_' + date_str+'.xlsx'
        return file_name

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
        self.wb.save(self.output_file_path)
    # write_excelのサブモジュール
    # 引数にワークシートオブジェクト
    # ワークシートオブジェクトのセルに値を代入
    def write_cell(self,ws,index,row):
        #コメント列が結合セルか判定
        is_merged = self.is_merged_cell_not_top_left(ws[index][6],ws)
        #検査結果は結合されないため分岐から除外
        ws[index][3].value = row[2]
        if (not(is_merged)):
            ws[index][5].value = row[4]
            ws[index][6].value = row[5]
            ws[index][7].value = row[6]
        else:
            comment_list = list()
            amend_list = list()

            comment_list.append(ws[is_merged][5].value)
            comment_list.append(row[4])

            amend_list.append(ws[is_merged][6].value)
            amend_list.append(row[5])

            ws[is_merged][5].value = "\n".join(comment_list)
            #対象ソースコードは結合不要なため除外
            #ws[is_merged][6].value += row[5]
            ws[is_merged][7].value = "\n".join(amend_list)

    # 結合セルであり結合の最初のセルではない場合は最初の行番号を返す
    def is_merged_cell_not_top_left(self,cell, ws):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                # 結合セルの左上セルかを判定
                if cell.row == merged_range.min_row and cell.column == merged_range.min_col:
                    return False  # 左上セルなら False
                return merged_range.min_row  # 結合セルで左上セルでない場合は最初の行番号を返す
        return False  # 結合セルでない場合
    
    # dfをopenpyxlで扱えるようにリストに変換
    def df_to_list(self,df):
        data_list = df.values.tolist()
        return data_list

    def main(self):
        
        # Excelファイルの書き込み
        self.write_excel()

if __name__ == '__main__':
    bulk_path='C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/本物1867_NUL0000_リンク_検査結果一括更新_20241227085441.xlsx'
    bulk_info = BulkInfo(bulk_path)

    # テスト用
    input_file_path = input('ファイルパスを入力してください')
    input_sheet_name = '検査'
    condition = {'管理番号':'NUL0000','検査カテゴリ':'リンク'}
    start_row = 12
    input_excel = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)
    input_excel_file = InputExcelFile(input_file_path, input_sheet_name,condition,input_start_row=start_row)

    # テスト用
    folder_path = 'C:/Users/user/OneDrive/プロジェクト/Python/Excel/test_data/'

    OutputExcelFile(input_excel_file.df,
                         folder_path,
                         bulk_info.management_id,
                         bulk_info.management_number,
                         bulk_info.category,
                         bulk_path
                         )
    