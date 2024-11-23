import pandas as pd
from openpyxl import load_workbook

# 読み込むExcelファイルのパス
input_file_path = 'C:/Users/user/OneDrive/プロジェクト/Python/Python/excel/test_data/input.xlsx'
input_sheet_name = 'Sheet'

# 書き込むExcelファイルのパス
output_file_path = 'C:/Users/user/OneDrive/プロジェクト/Python/Python/excel/test_data/___output.xlsx'
output_sheet_name = 'Sheet'

# Excelファイルを読み込む
df = pd.read_excel(input_file_path, sheet_name=input_sheet_name)

# 特定の条件でフィルタリング
filtered_df = df[(df['都道府県'] == '福井県') & (df['年齢'] >= 20)]

# フィルタリングしたデータの特定の列を加工
# 例: '年齢'列の値を2倍にする
filtered_df.loc[:, '年齢'] = filtered_df['年齢'] * 2

# excelファイルを新しく作成してデータを書き込む
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    filtered_df.to_excel(writer, sheet_name=output_sheet_name, index=False)

print("フィルタリングされたデータが新しいExcelファイルに保存されました。")