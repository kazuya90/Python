import tkinter as tk
from tkinter import filedialog
from module.OutputExcelFile import OutputExcelFile

class UI:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple UI")
        self.root.geometry("400x300")  # ウィンドウサイズを設定
        self.root.resizable(False, False)  # ウィンドウサイズを固定

        self.bulk_files = []
        self.inspection_file = ""

        self.create_widgets()

    def create_widgets(self):
        ## ラベル
        self.label = tk.Label(self.root, text="反映元の検査.xlsmファイルを選択してください", font=("Meiryo UI", 10, "bold"))
        self.label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="w")

        ## ファイル選択ボタンとラベル
        self.file_button = tk.Button(self.root, text="ファイルを選択", command=self.select_file)
        self.file_button.grid(row=1, column=0, padx=10, pady=0)

        self.inspection_file_label = tk.Label(self.root, text="ファイルが選択されていません")
        self.inspection_file_label.grid(row=1, column=1, padx=10, pady=0, sticky="ew")

        ## ラベル
        self.label = tk.Label(self.root, text="反映先の一括登録ファイルを選択してください", font=("Meiryo UI", 10, "bold"))
        self.label.grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="w")

        ## ファイル選択ボタンとラベル
        self.file_button_2 = tk.Button(self.root, text="ファイルを選択", command=self.select_files)
        self.file_button_2.grid(row=3, column=0, padx=10, pady=0)

        self.bulk_files_label = tk.Label(self.root, text="ファイルが選択されていません")
        self.bulk_files_label.grid(row=3, column=1, padx=10, pady=0, sticky="ew")

        ## 実行ボタン
        self.button = tk.Button(self.root, text="反映", command=lambda: self.trigger_event())
        self.button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

        ## 列の引き伸ばし
        self.root.grid_columnconfigure(1, weight=1)

        # 調整
        # ウィンドウサイズをウィジェットに合わせて調整
        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height())

    def trigger_event(self):
        # カスタムイベントを発生させる
        print("Triggering Custom Event...")
        self.root.event_generate("<<CustomEvent>>", when="tail")

    def select_file(self):

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx, *.xlsm")])
        if file_path:
            self.inspection_file = file_path
            self.inspection_file_label.config(text=file_path)

    # 複数のファイルを選択可能
    def select_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if file_paths:
            self.bulk_files = file_paths
            self.bulk_files_label.config(text=", ".join(file_paths))

if __name__ == "__main__":
    def print_test(message):
        print(message)    

    root = tk.Tk()
    app = UI(root)
    app.call_method(print_test, "Hello, World!")
    root.mainloop()
