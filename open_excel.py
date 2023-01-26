def main():
    import subprocess

    path = "//dwhnas1/DWH1/u_ueno/c_cuttingdata_auto/切断資料雛形(自動化テスト用).xlsx"
    subprocess.Popen(['start', path], shell=True)

if __name__ == "__main__":
    main()
