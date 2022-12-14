import streamlit as st
import openpyxl
import subprocess
import win32print
import win32api
import os
import glob


st.title('切断資料自動作成')

filename = st.text_input("切断資料ファイル名")
path = r"\\dwhnas1\DWH1\y_yuyama\000_inc_share\006_共有資料\s切断資料自動化\\" + filename + ".xlsx"

sep = "\\" if os.name =="nt" else "/"
file_paths = sorted(glob.glob(os.path.join(r"\\dwhnas1\DWH1\y_yuyama\000_inc_share\006_共有資料\s切断資料自動化\*.xlsx"), recursive=True))

filename_list_duplication = [""]
for file_path in file_paths:
    filename = os.path.split(file_path)[-1].replace(f".xlsx", "")
    filename_list_duplication.append(filename)
filename_list = list(dict.fromkeys(filename_list_duplication))

filename = st.selectbox("切断資料ファイル名", filename_list)

path = r"\\dwhnas1\DWH1\y_yuyama\000_inc_share\006_共有資料\s切断資料自動化\\" + filename + ".xlsx"

wb = openpyxl.load_workbook(path)
sheet = wb[filename]

st.sidebar.header('編集画面')
edit_form = st.sidebar.form('edit_form')
disconnection_date = edit_form.date_input('切断期日')
order_number = edit_form.number_input('受注番号', 0,  1000000, 0)
number_of_order = edit_form.number_input('受注数', 0,  10000, 0)
deadline = edit_form.date_input('納期')
push_button = edit_form.form_submit_button()


def PrintOut():
    win32api.ShellExecute(
        0,
        "print",
        path,
        "/c:""%s" % win32print.GetDefaultPrinter(),
        ".",
        0
    )

if push_button:
    sheet["O1"] = disconnection_date.strftime("%Y/%m/%d")
    sheet["O2"].value = order_number
    sheet["O3"].value = number_of_order
    sheet["O4"] = deadline.strftime("%Y/%m/%d")
    
    wb.save(path)        
    wb.close() 
    
    """
    ## 作成しました！
    """
    
    subprocess.Popen(['start', path], shell=True)
    
    wb.active = 0
    sheet.sheet_view.tabSelected = True
    
    wb.save(path)
    
    # printOut()
    
    wb.close()
    """
    ## 印刷しました！
    """
    
    wb = openpyxl.load_workbook(path)
    sheet.row_dimensions[1:4].hidden= True
    wb.save(path)
    wb.close()
    """
    ## 非表示にしました！
    """

# pyhtonからvbaに変数を渡して，vbaで変数を書き込む
# streamlitをキルして，エクセルファイルを立ち上げるコマンドを入力