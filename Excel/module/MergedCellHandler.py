import openpyxl
import pandas as pd

class MergedCellHandler:

  def __init__(self, df, merged_cell):
    self.df = df
    self.merged_cell = merged_cell
  
  # merged_cellに対応する文字列を結合する
  #結合セルの中身は['A1:A3','B1:B3']など
  def merge_cell(self, df, merged_cell):
    for cell in merged_cell:
      #':'で区切る
      cell = cell.split(':')
      #開始セルと終了セルを取得
      start_cell = cell[0]
      end_cell = cell[1]
      #開始セルの値を取得
      value = df[start_cell].values[0]
      #終了セルの値を取得
      end_value = df[end_cell].values[0]
      #結合セルの値を取得
      merged_value = value + end_value
      #結合セルの値を開始セルに代入
      df[start_cell] = merged_value
    return df
  
  # A1などの文字列を行と列のインデックスに変換する
  # A1:ZZ10にも対応
  # 戻り値は((行インデックス,列インデックス),(行インデックス,列インデックス))
  def cell_to_index(self, cell):
    #':'で区切る
    cell = cell.split(':')
    #開始セルと終了セルを取得
    start_cell = cell[0]
    end_cell = cell[1]
    #開始セルの行と列を取得
    start_column = ord(start_cell[0]) - ord('A')
    start_row = int(start_cell[1]) - 1
    #終了セルの行と列を取得
    end_column = ord(end_cell[0]) - ord('A')
    end_row = int(end_cell[1]) - 1
    return ((start_row, start_column), (end_row, end_column))
  
  # 受け取った結合セルのタプルから対応するセルの文字列を結合する
  # merged_cellの中身は((0, 0), (1, 1))など
  def merge_cell_by_index(self, df, merged_cell):
    start_cell = merged_cell[0]
    end_cell = merged_cell[1]
    start_row = int(start_cell[0])
    end_row = int(end_cell[0])
    col = int(start_cell[1])
    print("col",col)
    #開始セルの値を取得
    #結合する行の値を追記していく
    value_list = list()
    for row in range(start_row,end_row+1):
      value_list.append(df.iat[row, col])
    print(value_list)
    return "\n".join(value_list)
  
if __name__ == "__main__":
  df = pd.DataFrame({'A': ['A1', 'A2', 'A3'], 'B': ['B1', 'B2', 'B3']})
  merged_cell = ['A1:A3', 'B1:B3']
  mc = MergedCellHandler(df, merged_cell)
  merged_cell_index = mc.cell_to_index("A1:A3")
  print("merge_cell_index",merged_cell_index)
  merge_str = mc.merge_cell_by_index(df, merged_cell_index)
  print(merge_str)