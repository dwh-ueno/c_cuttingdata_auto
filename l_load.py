import subprocess
import os


def streamlit_run():
    cmd = "streamlit run c_cuttingdata_auto.py"
    subprocess.call(cmd.split())


if __name__ == '__main__':
    os.chdir("//dwhnas1/DWH1/u-ueno/c_cuttingdata_auto")
    print('現在のディレクトリ:', os.getcwd())
    streamlit_run()