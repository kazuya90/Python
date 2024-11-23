import tkinter as tk
from tkinter import filedialog

def on_button_click():
    pass

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        file_label.config(text=file_path)

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=folder_path)

# メインウィンドウの作成
root = tk.Tk()
root.title("Simple UI")


# ウィジェットの作成

## 入力フォーム1とそのラベル
label1 = tk.Label(root, text="入力フォーム1")
label1.grid(row=1, column=0, padx=10, pady=10)

entry = tk.Entry(root)
entry.grid(row=1, column=1, padx=10, pady=10,sticky="ew")

## フォルダ選択ボタンとラベル
folder_button = tk.Button(root, text="フォルダを選択", command=select_folder)
folder_button.grid(row=3, column=0, padx=10, pady=10)

folder_label = tk.Label(root, text="フォルダが選択されていません")
folder_label.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

## ファイル選択ボタンとラベル
file_button = tk.Button(root, text="ファイルを選択", command=select_file)
file_button.grid(row=4, column=0, padx=10, pady=10)

file_label = tk.Label(root, text="ファイルが選択されていません")
file_label.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

## 実行ボタン
button = tk.Button(root, text="実行", command=on_button_click)
button.grid(row=5, column=0,columnspan=4, padx=10, pady=10)

## 列の引き伸ばし
root.grid_columnconfigure(1, weight=1)


# 調整
# ウィンドウサイズをウィジェットに合わせて調整
root.update_idletasks()
root.minsize(root.winfo_width(), root.winfo_height())

# メインループ
root.mainloop()