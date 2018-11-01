from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfile
from tkinter.messagebox import showerror
import tkinter.scrolledtext as tkst
import requests
from pymouse import PyMouse
import time
import action1
from action1 import MyAction
import openpyxl
import requests

class MyFrame(Frame, MyAction):
    def __init__(self):
        Frame.__init__(self, height=510, width=770)#, bg='red')
        self.fname = ''
        self.master.title("서찬영 세무회계 사무소")
        self.master.rowconfigure(10)
        self.master.columnconfigure(10)
        self.grid(sticky=W+E+N+S)

        self.hello_msg = requests.get('http://ctascy.dothome.co.kr/hello.txt')
        self.hello_msg.encoding=None
        self.hello_msg = self.hello_msg.text
        self.lbl_hi = Label(self, text=self.hello_msg)
        self.lbl_hi.grid(row=0, column=0, columnspan=4)

        self.btn_openxls = Button(self, text='Open xlsx file', command=self.btn_openxlsx_func, width=10)
        self.btn_openxls.grid(row=1, column=0) #, rowspan=2)  # , sticky=W)

        self.btn_noxls = Button(self, text='No Excel File', command=self.btn_noxls_func, width=10)
        self.btn_noxls.grid(row=2, column=0)

        self.lbl_start = Label(self, text='시작행 번호', width=10).grid(row=1, column=1)
        self.ent_start_value = StringVar()
        self.ent_start_value.set("1")
        self.ent_start = Entry(self, textvariable=self.ent_start_value, width=10).grid(row=2, column=1)
        ###

        self.lbl_end = Label(self, text='끝 행 번호', width=10).grid(row=1, column=2)
        self.ent_end_value = StringVar()
        self.ent_end = Entry(self, textvariable=self.ent_end_value, width=10).grid(row=2, column=2)  # , sticky=W)
        ###

        # self.lbl_pw = Label(self, text='Password', width=5).grid(row=1, column=3)
        self.ent_pw_value = StringVar()
        self.ent_pw = Entry(self, textvariable=self.ent_pw_value,show="*", width=5).grid(row=1, column=3)
        self.btn_pw = Button(self, text="PW", command = self.btn_pw_func, width=5)
        self.btn_pw.grid(row=1, column=4)

        self.btn_add_range = Button(self, text="행 범위 선택 추가", command=self.btn_add_range_func, width=10)  # , width=30)
        self.btn_add_range.grid(row=2, column=3)  # , columnspan=2)
        ###

        self.lbl_info = Label(self, text=" 아래 네 가지 옵션 중 선택하세요 ")
        self.lbl_info.grid(row=4, column=0, columnspan=4)

        ###
        self.lbl_command = Label(self, text="명령어 선택 ", width=10)
        self.lbl_command.grid(row=5, column=0)
        self.combo_command_str = StringVar()
        self.combo_command = ttk.Combobox(self, textvariable=self.combo_command_str, width=10, state='readonly')  # width=15
        self.combo_command['values'] = ('명령을 선택', '{{Enter()}}', '{{SpaceBar()}}', "{{Sleep()}}", "{{MoveToRight()}}", "{{MoveToLeft()}}", "{{MoveToUp()}}", "{{MoveToDown()}}")
        self.combo_command.grid(row=6, column=0)
        self.combo_command.current(0)
        self.btn_command_toscrt = Button(self, text='명령 추가', command=self.btn_command_toscrt_func, width=10,
                                    state='normal')  # ''disabled')
        self.btn_command_toscrt.grid(row=7, column=0)
        ###

        self.lbl_copy_cell = Label(self, text="복사할 Cell ", width=10)
        self.lbl_copy_cell.grid(row=5, column=1)
        self.combo_copy_cell_str = StringVar()
        self.combo_copy_cell = ttk.Combobox(self, textvariable=self.combo_copy_cell_str, width=10, state='readonly')  # width=15
        self.combo_copy_cell['values'] = ('복사할 셀 선택', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M')
        self.combo_copy_cell.grid(row=6, column=1)
        self.combo_copy_cell.current(0)
        self.btn_copy_cell = Button(self, text='복사할 컬럼 추가', command=self.btn_copy_cell_func, width=10,
                               state='disabled')  # ''disabled')
        self.btn_copy_cell.grid(row=7, column=1)
        # self.ent_copy_column_value = StringVar()
        # self.ent_copy_column = Entry(self, textvariable=ent_copy_column_value).grid(row=6, column=1)
        # self.btn_copy_column = Button(self, text='복사할 컬럼 추가', command=btn_copy_column_func, state='disabled')
        # self.btn_copy_column.grid(row=7, column=1)
        ###

        self.lbl_typing = Label(self, text="수동 입력 값", width=10)
        self.lbl_typing.grid(row=5, column=2)
        self.ent_typing_value = StringVar()
        self.ent_typing = Entry(self, textvariable=self.ent_typing_value, width=10).grid(row=6, column=2)
        self.btn_typing = Button(self, text='Typing', command=self.btn_typing_func, width=10, state='normal')
        self.btn_typing.grid(row=7, column=2)
        ###

        self.scrt = tkst.ScrolledText(self, height=10, width=30, wrap=WORD)  # , width=50 Create a scrolledtext
        self.scrt.grid(row=8, column=0, columnspan=3, sticky=W + E)  # columnspan=3)
        # self.scrt.focus_set()  # Default focus
        ###

        self.btn_delete_scrt = Button(self, text="삭제하기", command=self.delete_from_scrt, width=10)  # , width=10)
        self.btn_delete_scrt.grid(row=10, column=1, sticky=E)

        self.btn_open_txt = Button(self, text='불러오기', command=self.btn_open_txt_func, width=10)
        self.btn_open_txt.grid(row=10, column=0)

        self.btn_save_txt = Button(self, text='저장하기', command=self.btn_save_txt_func, width=10)
        self.btn_save_txt.grid(row=10, column=2)

        self.lbl_pos_value = StringVar()
        self.lbl_pos_value.set('Click 하기')
        self.lbl_pos = Label(self, textvariable=self.lbl_pos_value, width=10)
        self.lbl_pos.grid(row=5, column=3)
        self.btn_pos = Button(self, text='좌표 받기', command=self.btn_pos_func, width=10)
        self.btn_pos.grid(row=6, column=3)
        self.btn_pos_add_once = Button(self, text='좌표 한번만추가', command=self.btn_pos_add_once_func, width=5)
        self.btn_pos_add_once.grid(row=7, column=3)
        self.btn_pos_add = Button(self, text='좌표 추가', command=self.btn_pos_add_func, width=8)
        self.btn_pos_add.grid(row=7, column=4)
        
        self.btn_run = Button(self, text='실행', command=self.btn_run_func, width=10, state='disabled')
        self.btn_run.grid(row=10, column=3)

    def btn_openxlsx_func(self):
        fname = ''
        fname = askopenfilename(filetypes=(
            # ("Template files", "*.tplate"),
            # ("HTML files", "*.html;*.htm"),
            # ("All files", "*.*"),
            # ("text files", "*.txt"),))
            ("xlsx files", "*.xlsx"),))
        if fname:
            try:
                # print(fname)
                self.btn_copy_cell.configure(state='normal')
                self.fname = fname
                # return self.fname
            except:  # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)

    def btn_noxls_func(self):
        pass


    def btn_pw_func(self):
        pw = requests.get('http://ctascy.dothome.co.kr/pw.txt').text
        print(self.ent_pw_value.get())
        if self.ent_pw_value.get()==pw:
            self.btn_run.configure(state='normal')

    def btn_add_range_func(self):
        s = self.ent_start_value.get()
        e = self.ent_end_value.get()
        try:
            s = int(s)
            e = int(e)+1
            if e >= s:
                # print(self.fname)
                self.btn_add_range.configure(text='범위 추가 완료')
                self.s = s
                self.e = e
                pass
                # print(str(s))
                # print(str(e))
                # btn_copy_cell.configure(state='normal')
                # btn_command_toscrt.configure(state='normal')
                # btn_add_range.configure(state='normal')#'disabled')
                # btn_typing.configure(state='normal')
            else:
                self.ent_end_value.set('끝 행은 시작행 보다 큰 수여야 합니다.')
        except:
            self.ent_start_value.set('행 숫자를 넣어 주세요')
            self.ent_end_value.set('행 숫자를 넣어 주세요')

    def btn_command_toscrt_func(self):
        text = self.combo_command_str.get()
        # # print(text)
        # combo_command.current(0)
        if text == self.combo_command['values'][0]:
            pass
        else:
            self.scrt.insert(INSERT, text)
            self.scrt.insert(INSERT, '\n')
            self.combo_command.current(0)

    def btn_copy_cell_func(self):
        text = self.combo_copy_cell_str.get()
        if text == self.combo_copy_cell['values'][0]:
            pass
        else:
            text = '{{CopyAndPaste(sheet,"'+text+'")}}'
            self.scrt.insert(INSERT, text)
            self.scrt.insert(INSERT, '\n')
            self.combo_copy_cell.current(0)

    def btn_typing_func(self):
        text = self.ent_typing_value.get()
        if text == '':
            pass
        else:
            text = '{{Type("'+text+'")}}'
            self.scrt.insert(INSERT, text)
            self.scrt.insert(INSERT, '\n')
    
    def btn_pos_add_once_func(self):
            text = self.lbl_pos_value.get()
            if text == '좌표':
                pass
            else:
                pos_x = text.split(":")[0]
                pos_y = text.split(":")[1]
                text = '[[CLick('+pos_x + ','+pos_y+')]]'
                self.scrt.insert(INSERT, text)
                self.scrt.insert(INSERT, '\n')

    def btn_pos_add_func(self):
        text = self.lbl_pos_value.get()
        if text == '좌표':
            pass
        else:
            pos_x = text.split(":")[0]
            pos_y = text.split(":")[1]
            text = '{{CLick('+pos_x + ','+pos_y+')}}'
            self.scrt.insert(INSERT, text)
            self.scrt.insert(INSERT, '\n')

    def delete_from_scrt(self):
        self.scrt.delete(1.0, END)

    def btn_open_txt_func(self):
        fname = askopenfilename(filetypes=(
            ("text files", "*.txt"),))
        if fname:
            try:
                with open(fname, 'r+') as f:
                    text = f.readlines()
                    print(text)
                    for t in text:
                        self.scrt.insert(INSERT, t)
            except:  # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            return

    def btn_save_txt_func(self):
        fout = asksaveasfile(mode='w', defaultextension=".txt")
        text2save = str(self.scrt.get(0.0, END))
        try:
            fout.write(text2save)
            fout.close()
        except:
            pass

    def btn_pos_func(self):
        m = PyMouse()
        pos = m.position()
        time.sleep(0.8)
        pos2 = m.position()
        if pos == pos2:
            pos_str = str(int(pos[0])) + ':' + str(int(pos[1]))
            self.lbl_pos_value.set(pos_str)
        else:
            self.btn_pos_func()

    def btn_run_func(self):
        # print(self.fname)
        text2save = str(self.scrt.get(0.0, END))
        if self.fname != '' :
            fname = self.fname
            book = openpyxl.load_workbook(filename=fname, data_only=True)
            sheet = book.active
            for i in range(self.s, self.e):
                action1.MyAction().run_func(text2save)
        else:
            for index, i in enumerate(range(self.s, self.e)):
                action1.MyAction().run_func(text2save, index=index)

            # print(self.s, self.e)
            # print(sheet["a"+str(self.s)].value)
        # print(text2save)
        # act = action.MyAction()
        # act.run_func(text2save)


if __name__ == "__main__":
    MyFrame().mainloop()