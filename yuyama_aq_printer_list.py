import subprocess
import win32print

print(win32print.GetDefaultPrinter())
# data = subprocess.check_output(['wmic', 'printer' , 'list' ,'brief']).decode('utf-8').split('\r\r\n')
# data = data[1:]
# print(data)
# for line in data:
#     for pritername in line.split("  "):
#         if pritername != "":
#             print(pritername)

