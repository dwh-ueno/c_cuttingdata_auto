import win32print
import win32api

def PrintOut():
    win32api.ShellExecute(
        0,
        "print",
        path,
        "/c:""%s" % win32print.GetDefaultPrinter(),
        ".",
        0
    )