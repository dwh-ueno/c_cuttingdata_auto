def main():
    import subprocess
    import win32api
    
    path = r"\\dwhnas1\DWH1\u-ueno\c_cuttingdata_auto\切断資料雛形(自動化テスト用).xlsx"
    subprocess.Popen(['start', path], shell=True)
    
if __name__ == "__main__":
    main()
