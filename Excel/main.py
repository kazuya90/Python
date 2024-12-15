from module.InputExcelFile import InputExcelFile
from module.BulkInfo  import BulkInfo
from module.OutputExcelFile import OutputExcelFile
from module.UI import UI
import tkinter as tk

class EventReceiver:
    def __init__(self, root):
        # rootにイベントをバインド
        root.bind("<<CustomEvent>>", self.execute_action)

    def execute_action(self, event):
        # カスタムイベントに応じて実行する処理
        button_click(ui)

if __name__ == '__main__':
  def test_print(message):
    print(message)
  def button_click(ui):
    for file in ui.bulk_files:
      bulk_info = BulkInfo(file)
      start_row = 3
      input_sheet_name = '検査'
      input_excel_file = InputExcelFile(ui.inspection_file ,input_sheet_name, bulk_info.condition,input_start_row=start_row)
      print(input_excel_file.df)
      print(ui.folder_path)
      print(bulk_info.management_id)
      print(bulk_info.management_alphabet)
      print(bulk_info.management_number)
      print(bulk_info.category)
      print(ui.output_sheet_name)
      folder_path = ui.folder_path + '/'

      OutputExcelFile(input_excel_file.df, folder_path, bulk_info.management_id, bulk_info.management_alphabet, bulk_info.management_number, bulk_info.category, ui.output_sheet_name)
    print('処理が完了しました。')
  root = tk.Tk()
  event_receiver = EventReceiver(root)
  ui = UI(root)
  ui.root.mainloop()




