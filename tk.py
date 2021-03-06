from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import tkinter.scrolledtext as tkst
import requests
from evalpy import DoReMi


class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("서찬영 세무사")
        self.master.rowconfigure(10, weight=1)
        self.master.columnconfigure(10, weight=1)
        self.grid(sticky=W+E+N+S)

        # 지워도 됩니다. 레이블 연습용
        # self.label_hello = Label(self, text='hello')
        # self.label_hello.grid(row=1, column=12, sticky=E)

        # 버튼 for 파일 탐색기 실행
        self.btn_fileopen = Button(self, text="파일탐색", command=self.load_file, width=10)
        self.btn_fileopen.grid(row=1, column=0, sticky=W)

        self.a_list = []

        # 엑셀파일 읽는 시작과 끝 행 넘버
        self.start_rows = StringVar()
        self.lbl_start = Label(self, text='시작 라인').grid(row=2, column=0)
        self.ent_start =  Entry(self, textvariable=self.start_rows).grid(row=3, column=0)

        self.end_rows = StringVar()
        self.lbl_end = Label(self, text='끝 라인').grid(row=2, column=1)
        self.ent_end =  Entry(self,textvariable=self.end_rows).grid(row=3, column=1)

        self.btn_range = Button(self, text='영역추가', command=self.write_to_scrt_fron_entry, width=10).grid(row=3, column=2)

        # 리스트 박스 (명령어)
        self.listbox = Listbox()
        self.listbox.insert(0, "{{EnterKey()}}")
        self.listbox.insert(1, "{{SpaceBar()}}")
        self.listbox.insert(2, "{{MoveToRight()}}")
        self.listbox.insert(3, "{{MoveToLeft()}}")
        self.listbox.insert(4, "{{MoveToUp()}}")
        self.listbox.insert(5, "{{MoveToDown()}}")
        # self.listbox.delete(1, 2)
        self.listbox.grid(row=12,column=0, sticky=W)

        # 버튼 for 리스트 박스 중에 골라서 스크롤 텍스트에 넣기 위함
        self.btn_1 = Button(self, text="추가", command=self.write_to_scrt, width=10)
        self.btn_1.grid(row=10, column=0, sticky=W)

        # 스크롤텍스트
        self.scrt = tkst.ScrolledText(self, width=50, height=4, wrap=WORD) # Create a scrolledtext
        self.scrt.grid(column=0, row=4, columnspan=3)
        self.scrt.focus_set()                                                # Default focus

        self.btn_delete_scrt = Button(self, text="삭제하기", command=self.delete_from_scrt, width=10)
        self.btn_delete_scrt.grid(row=10, column=1, sticky=E)


        self.btn_sing = Button(self,text='sing a song', command=self.sing_a_song, width=10).grid(row=11,column=3)


    # 파일 탐색창 열기 link to btn_fileopen
    def load_file(self):
        fname = askopenfilename(filetypes=(
            # ("Template files", "*.tplate"),
            # ("HTML files", "*.html;*.htm"),
            # ("All files", "*.*"),
            ("text files", "*.txt"), ))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                self.work_from_file(fname)
                self.print_a_list()
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            return

    def work_from_file(self, fname):
        with open (fname, 'r+') as f:
            txt = f.read()
            self.a_list.append(txt)
            # print_a_list(self)

    def print_a_list(self):
        print(self.a_list)

    # 스크롤 텍스트에 입력 추가 하기
    def write_to_scrt(self):
        value=str((self.listbox.get(ACTIVE)))+'\n'
        self.scrt.insert(INSERT, value)                    # insert text in a scrolledtext
        self.scrt.see(END)

    # 스크롤 텍스트에 삭제 하기
    def delete_from_scrt(self,):
        self.scrt.delete(1.0, END)

    def write_to_scrt_fron_entry(self):
        start = self.start_rows.get()
        end = self.end_rows.get()
        print(1)

    def sing_a_song(self):
        song = DoReMi()
        song.run_func()


if __name__ == "__main__":
    MyFrame().mainloop()