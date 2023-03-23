import os
import subprocess
from time import sleep



def streamlit_run():
    cmd = "streamlit run c_cuttingdata_auto.py"
    subprocess.call(cmd.split())

           

if __name__ == "__main__":
    os.chdir(r"//dwhnas1/DWH1/u_ueno/c_cuttingdata_auto")
    streamlit_run()