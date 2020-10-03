from os import walk
import time
import pandas as pd
import glob


def scan_file_xlsx(path):
    filenames = []
    for root, dirs, file in walk(path):
        files_xlsx = glob.glob(root + r"\*.xlsx")
        for file_xlsx in files_xlsx:
            filenames.append(file_xlsx)
        files_xls = glob.glob(root + r"\*.xls")
        for file_xls in files_xls:
            filenames.append(file_xls)
    return filenames


def result_out(filenames, key_word):
    key_files = []
    for filename_root in filenames:
        key_file = scan_table(filename_root, key_word)
        if key_file != "":
            key_files.append(key_file)
            print(key_file)
        
    if key_files == []:
        print("無任何關鍵字!")
    input("完成！任意鍵！\n")


def scan_table(filename_root, key_word):
    #df = pd.read_excel(filename_root, encoding="utf_8_sig")
    df = pd.read_excel(filename_root, encoding="utf-8")
    rows = df.shape[0]
    columns = df.shape[1]
    key_file = ""

    for column in range(columns):
        for row in range(rows):
            data_item = str(df.iloc[row, column])
            if data_item.find(key_word) != -1:
                key_file = filename_root
                return key_file
    return key_file


while True:
    path = r"{}".format(input("請輸入位址："))
    filenames = scan_file_xlsx(path)
    key_word = input("請輸入關鍵字：")
    result_out(filenames, key_word)