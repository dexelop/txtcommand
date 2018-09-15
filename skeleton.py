from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
from tkinter.messagebox import showerror
import tkinter.scrolledtext as tkst
import requests
from evalpy import DoReMi


win = Tk()
win.title('Hello title')


def btn_openxlsx_func():
    fname = askopenfilename(filetypes=(
        # ("Template files", "*.tplate"),
        # ("HTML files", "*.html;*.htm"),
        # ("All files", "*.*"),
        # ("text files", "*.txt"),))
        ("xlsx files", "*.xlsx"),))
    if fname:
        try:
            print(fname)
        except:  # <- naked except is a bad idea
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return


def btn_add_range_func():
    s = ent_start_value.get()
    e = ent_end_value.get()
    try:
        s = int(s)
        e = int(e)
        if e > s:
            print(str(e))
            print(str(s))
            btn_copy_column.configure(state='normal')
            btn_change_lbl_ex.configure(state='normal')
            btn_add_range.configure(state='disabled')
            btn_typing.configure(state='disabled')
        else:
            ent_end_value.set('끝 행은 시작행 보다 큰 수여야 합니다.')
    except:
        ent_start_value.set('행 숫자를 넣어 주세요')


def btn_copy_column_func():
    print('hi')


def btn_change_lbl_ex_func():
    text = combo_str.get()
    if text == '명령을 선택하거나, 소리질러':
        pass
    else:
        combo.current(0)


def btn_typing_func():
    print('typing is working!')


def delete_from_scrt():
    scrt.delete(1.0, END)


def btn_open_txt_func():
    fname = askopenfilename(filetypes=(
        # ("Template files", "*.tplate"),
        # ("HTML files", "*.html;*.htm"),
        # ("All files", "*.*"),
        ("text files", "*.txt"),))
        # ("xlsx files", "*.xlsx"),))
    if fname:
        try:
            print(fname)
        except:  # <- naked except is a bad idea
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return


def btn_save_txt_func():
    fout = asksaveasfile(mode='w', defaultextension=".txt")
    text2save = str(scrt.get(0.0,END))
    try:
        fout.write(text2save)
        fout.close()
    except:
        pass



btn_openxls = Button(win, text='Open xlsx file', command=btn_openxlsx_func, width=50)
btn_openxls.grid(row=0, column=0, columnspan=3)#, sticky=W)
###

lbl_start = Label(win, text='시작행 번호').grid(row=1, column=0)
ent_start_value = StringVar()
ent_start = Entry(win, textvariable=ent_start_value).grid(row=2, column=0)
###

lbl_end = Label(win, text='끝 행 번호').grid(row=1, column=1)
ent_end_value = StringVar()
ent_end = Entry(win, textvariable=ent_end_value).grid(row=2, column=1) #, sticky=W)
###

btn_add_range = Button(win, text="행 범위 선택 추가", command=btn_add_range_func)#, width=30)
btn_add_range.grid(row=2, column=2)#, columnspan=2)
###

lbl_info = Label(win, text=" 아래 세가지 옵션 중 선택하세요 ")
lbl_info.grid(row=4, column=0, columnspan=3)

###
lbl_info_1 = Label(win, text="명령어 선택 ")
lbl_info_1.grid(row=5, column=0)
combo_str = StringVar()
combo = ttk.Combobox(win, textvariable=combo_str, state='readonly')# width=15
combo['values'] = ('명령을 선택','apple', 'banana', 'pineapple', 'pear')
combo.grid(row=6, column=0)
combo.current(0)
btn_change_lbl_ex = Button(win, text='레이블 변경', command=btn_change_lbl_ex_func, state='disabled')
###

btn_change_lbl_ex.grid(row=7, column=0)
lbl_info_2 = Label(win, text="| 복사할 Column 입력 ")
lbl_info_2.grid(row=5, column=1)
ent_copy_column_value = StringVar()
ent_copy_column = Entry(win, textvariable=ent_copy_column_value).grid(row=6, column=1)
btn_copy_column = Button(win, text='복사할 컬럼 추가', command=btn_copy_column_func, state='disabled')
btn_copy_column.grid(row=7, column=1)
###

lbl_info_3 = Label(win, text="| 입력할 값 수동 입력")
lbl_info_3.grid(row=5, column=2)
ent_typing_value = StringVar()
ent_typing = Entry(win, textvariable=ent_typing_value).grid(row=6, column=2)
btn_typing = Button(win, text='Typing', command=btn_typing_func, state='disabled')
btn_typing.grid(row=7, column=2)
###

scrt = tkst.ScrolledText(win, height=4, wrap=WORD)  # , width=50 Create a scrolledtext
scrt.grid(row=8, column=0, columnspan=3, sticky=W+E)#columnspan=3)
scrt.focus_set()  # Default focus
###

btn_delete_scrt = Button(win, text="삭제하기", command=delete_from_scrt)#, width=10)
btn_delete_scrt.grid(row=10, column=1, sticky=E)

btn_open_txt = Button(win, text='불러오기', command=btn_open_txt_func)
btn_open_txt.grid(row=10, column=0)

btn_save_txt = Button(win, text='저장하기', command=btn_save_txt_func)
btn_save_txt.grid(row=10, column=2)

win.mainloop()
