import os
import sys
import pandas as pd
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as meg

if getattr(sys, 'frozen', False):
    # test.exe로 실행한 경우,test.exe를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    # python test.py로 실행한 경우,test.py를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(__file__))

meta_data = pd.read_excel('POST_CODE.xlsx', sheet_name=None, na_values="TT")

def search():
    if str(combobox.get()) == 'None':
        meg.showwarning(title='Alert', message='Please select manufacturer')
    else:
        Spec_text.delete('1.0', "end")
        sheet = str(combobox.get())
        data = meta_data[sheet]
        data = data.fillna('')
        find_code = line1.get().upper()
        if find_code == '' or len(find_code) > 2:
            Spec_text.insert('1.0', 'Please correct code')
        else:
            description = ''
            for value in data.values:
                if value[0] == find_code:
                    description += value[1]
                    description += '\n'
            if description == '':
                Spec_text.insert('1.0', 'No found')
            else:
                Spec_text.insert('1.0', description)

def DCT_search():
    relate_task, note_history, date, DC_code = 3, 4, 6, 7
    if str(combobox.get()) == 'None':
        pass
    else:
        DCT_text.delete('1.0', "end")
        data = meta_data['DCT']
        data = data.fillna('')
        find_code = line1.get().upper()
        if find_code == '' or len(find_code) > 2:
            Spec_text.insert('1.0', 'Please correct code')
        else:
            note = ''
            for value in data.values:
                if value[0] == find_code:
                    if value[date] != '':
                        time_date = value[date]
                        note += ('GDCO : ' + value[relate_task][-10:] + " | ")
                        note += ('date : ' + time_date.strftime('%Y-%m-%d') + " | ")
                        note += ('DC : ' + value[DC_code])
                        note += '\n'
                        note += ('note : ' + value[note_history])
                        note += '\n'
                    else:
                        pass
            if note == '':
                DCT_text.insert('1.0', 'No found')
            else:
                DCT_text.insert('1.0', note)

def close_window():
    root.destroy()

root = Tk()
root.geometry("260x520")
root.title('POST Library')
root.resizable(True, False)

# h1
label1 = Label(root, text='Manufacturer')
label1.place(x=5, y=5)
combo = ['Wiwynn', 'Quanta', "Ingrasys", "Lenovo", "DELL", "HPE"]
combobox = ttk.Combobox(root)
combobox.config(values=combo, state='readonly', width=8)
combobox.set('None')
combobox.place(x=90, y=5)

# h2
label2 = Label(root, text='POST CODE')
label2.place(x=5, y=30)
line1 = Entry(root, width=10, border=1, relief='sunken')
line1.place(x=90, y=30)
btn1 = Button(root, text='Search!', command=lambda:[search(), DCT_search()])
btn1.place(x=180, y=30)

# h3
label3 = Label(root, text='Specification')
label3.place(x=5, y=60)
Spec_text = Text(root, width=30, height=10)
Spec_text.place(x=5, y=90)

# h4
label4 = Label(root, text='DCT Comments')
label4.place(x=5, y=270)
DCT_text = Text(root, width=30, height=10)
DCT_text.place(x=5, y=300)

# h5
label5 = Label(root, text='Note')
label5.place(x=5, y=470)
line5 = Entry(root, width=10, border=1, relief='sunken')
line5.place(x=60, y=470)
btn2 = Button(root, text='Save')
btn2.place(x=180, y=470)

# h6
btn3 = Button(root, text='Exit', command=close_window)
btn3.place(x=5, y=490)

mainloop()
