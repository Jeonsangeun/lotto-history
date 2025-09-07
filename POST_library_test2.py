import os
import sys
from tkinter import *
import pandas as pd
import tkinter.ttk as ttk

if getattr(sys, 'frozen', False):
    # test.exe로 실행한 경우,test.exe를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    # python test.py로 실행한 경우,test.py를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(__file__))

meta_data = pd.read_excel('POST_CODE.xlsx', sheet_name=None, na_values="TT")

def close_window():
    root.destroy()

root = Tk()
root.geometry("260x520")
root.title('POST Library')
root.resizable(True, False)

# h6
btn3 = Button(root, text='Exit', command=close_window)
btn3.place(x=5, y=490)

mainloop()