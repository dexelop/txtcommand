from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
from tkinter.messagebox import showerror
import tkinter.scrolledtext as tkst
import requests
from evalpy import DoReMi


win = Tk()
win.title('자동 기장 입력')


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
            pass
            # print(str(s))
            # print(str(e))
            # btn_copy_cell.configure(state='normal')
            # btn_command_toscrt.configure(state='normal')
            # btn_add_range.configure(state='normal')#'disabled')
            # btn_typing.configure(state='normal')
        else:
            ent_end_value.set('끝 행은 시작행 보다 큰 수여야 합니다.')
    except:
        ent_start_value.set('행 숫자를 넣어 주세요')


def btn_command_toscrt_func():
    text = combo_command_str.get()
    # # print(text)
    # combo_command.current(0)
    if text == combo_command['values'][0]:
        pass
    else:
        scrt.insert(INSERT, text)
        scrt.insert(INSERT, '\n')
        combo_command.current(0)


def btn_copy_cell_func():
    text = combo_copy_cell_str.get()
    if text == combo_copy_cell['values'][0]:
        pass
    else:
        scrt.insert(INSERT, text)
        scrt.insert(INSERT, '\n')
        combo_copy_cell.current(0)


def btn_typing_func():
    text = ent_typing_value.get()
    if text == '':
        pass
    else:
        scrt.insert(INSERT, text)
        scrt.insert(INSERT, '\n')


def btn_pos_add_func():
    text = lbl_pos_value.get()
    if text == '좌표':
        pass
    else:
        scrt.insert(INSERT, text)
        scrt.insert(INSERT, '\n')


def delete_from_scrt():
    scrt.delete(1.0, END)


def btn_open_txt_func():
    fname = askopenfilename(filetypes=(
        ("text files", "*.txt"),))
    if fname:
        try:
            with open(fname, 'r+') as f:
                text = f.readlines()
                print(text)
                for t in text:
                    scrt.insert(INSERT, t)
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


lbl_hi = Label(win, text='환영합니다요')
lbl_hi.grid(row=0, column=0, columnspan=4)


btn_openxls = Button(win, text='Open xlsx file', command=btn_openxlsx_func, width=10)
btn_openxls.grid(row=1, column=0, rowspan=2)#, sticky=W)
###

lbl_start = Label(win, text='시작행 번호', width=10).grid(row=1, column=1)
ent_start_value = StringVar()
ent_start = Entry(win, textvariable=ent_start_value, width=10).grid(row=2, column=1)
###

lbl_end = Label(win, text='끝 행 번호', width=10).grid(row=1, column=2)
ent_end_value = StringVar()
ent_end = Entry(win, textvariable=ent_end_value, width=10).grid(row=2, column=2) #, sticky=W)
###

btn_add_range = Button(win, text="행 범위 선택 추가", command=btn_add_range_func, width=10)#, width=30)
btn_add_range.grid(row=2, column=3)#, columnspan=2)
###

lbl_info = Label(win, text=" 아래 네 가지 옵션 중 선택하세요 ")
lbl_info.grid(row=4, column=0, columnspan=4)

###
lbl_command = Label(win, text="명령어 선택 ", width=10)
lbl_command.grid(row=5, column=0)
combo_command_str = StringVar()
combo_command = ttk.Combobox(win, textvariable=combo_command_str, width=10, state='readonly')# width=15
combo_command['values'] = ('명령을 선택','apple', 'banana', 'pineapple', 'pear')
combo_command.grid(row=6, column=0)
combo_command.current(0)
btn_command_toscrt = Button(win, text='명령 추가', command=btn_command_toscrt_func, width=10, state='normal')#''disabled')
btn_command_toscrt.grid(row=7, column=0)
###

lbl_copy_cell = Label(win, text="복사할 Cell ", width=10)
lbl_copy_cell.grid(row=5, column=1)
combo_copy_cell_str = StringVar()
combo_copy_cell = ttk.Combobox(win, textvariable=combo_copy_cell_str, width=10, state='readonly')# width=15
combo_copy_cell['values'] = ('복사할 셀 선택','A', 'B', 'C', 'D', 'E', 'F','G','H','I','J', 'K', 'L', 'M')
combo_copy_cell.grid(row=6, column=1)
combo_copy_cell.current(0)
btn_copy_cell = Button(win, text='복사할 컬럼 추가', command=btn_copy_cell_func, width=10, state='normal')#''disabled')
btn_copy_cell.grid(row=7, column=1)
# ent_copy_column_value = StringVar()
# ent_copy_column = Entry(win, textvariable=ent_copy_column_value).grid(row=6, column=1)
# btn_copy_column = Button(win, text='복사할 컬럼 추가', command=btn_copy_column_func, state='disabled')
# btn_copy_column.grid(row=7, column=1)
###

lbl_typing = Label(win, text="수동 입력 값", width=10)
lbl_typing.grid(row=5, column=2)
ent_typing_value = StringVar()
ent_typing = Entry(win, textvariable=ent_typing_value, width=10).grid(row=6, column=2)
btn_typing = Button(win, text='Typing', command=btn_typing_func, width=10, state='normal')
btn_typing.grid(row=7, column=2)
###

scrt = tkst.ScrolledText(win, height=10, width=30, wrap=WORD)  # , width=50 Create a scrolledtext
scrt.grid(row=8, column=0, columnspan=3, sticky=W+E)#columnspan=3)
# scrt.focus_set()  # Default focus
###

btn_delete_scrt = Button(win, text="삭제하기", command=delete_from_scrt, width=10)#, width=10)
btn_delete_scrt.grid(row=10, column=1, sticky=E)

btn_open_txt = Button(win, text='불러오기', command=btn_open_txt_func, width=10)
btn_open_txt.grid(row=10, column=0)

btn_save_txt = Button(win, text='저장하기', command=btn_save_txt_func, width=10)
btn_save_txt.grid(row=10, column=2)


from pymouse import PyMouse
import time


def btn_pos_func():
    m = PyMouse()
    pos = m.position()
    time.sleep(0.8)
    pos2 = m.position()
    if pos == pos2:
        pos_str = str(int(pos[0])) + ':' + str(int(pos[1]))
        lbl_pos_value.set(pos_str)
    else:
        btn_pos_func()


lbl_pos_value = StringVar()
lbl_pos_value.set('좌표')
lbl_pos = Label(win, textvariable=lbl_pos_value, width=10)
lbl_pos.grid(row=5, column=3)
btn_pos = Button(win, text='좌표 받기', command=btn_pos_func, width=10)
btn_pos.grid(row=6, column=3)
btn_pos_add = Button(win, text='좌표 추가', command=btn_pos_add_func, width=10)
btn_pos_add.grid(row=7, column=3)


# import pyautogui as pa
#
#
# def btn_pos_add_func():
#     pos = lbl_pos_value.get()
#     x = int(pos.split(':')[0])
#     y = int(pos.split(':')[1])
#     pa.click(x,y)




win.mainloop()
