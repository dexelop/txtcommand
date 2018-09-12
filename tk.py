from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import tkinter.scrolledtext as tkst

class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("서찬영 세무사")
        self.master.rowconfigure(5, weight=1)
        self.master.columnconfigure(5, weight=1)
        self.grid(sticky=W+E+N+S)

        # 지워도 됩니다. 레이블 연습용
        self.label_hello = Label(self, text='hello')
        self.label_hello.grid(row=1, column=5, sticky=E)

        # 버튼 for 파일 탐색기 실행
        self.btn_fileopen = Button(self, text="파일탐색", command=self.load_file, width=10)
        self.btn_fileopen.grid(row=1, column=0, sticky=W)

        self.a_list = []
        
        # 리스트 박스 (명령어)
        self.listbox = Listbox()
        self.listbox.insert(0, "1번")
        self.listbox.insert(1, "2번")
        self.listbox.insert(2, "3번")
        self.listbox.insert(3, "4번")
        self.listbox.insert(4, "5번")
        # self.listbox.delete(1, 2)
        self.listbox.grid(row=2,column=0, sticky=W)

        # 버튼 for 리스트 박스 중에 골라서 스크롤 텍스트에 넣기 위함
        self.btn_1 = Button(self, text="추가", command=self.write_to_scrt, width=10)
        self.btn_1.grid(row=10, column=0, sticky=W)

        # 스크롤텍스트
        self.scrt = tkst.ScrolledText(self, width=33, height=3, wrap=WORD) # Create a scrolledtext
        self.scrt.grid(column=0, row=2, columnspan=3)
        self.scrt.focus_set()                                                # Default focus


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

    def write_to_scrt(self):
        value=str((self.listbox.get(ACTIVE)))+'!'
        self.scrt.insert(INSERT, value)                    # insert text in a scrolledtext
        self.scrt.see(END)


if __name__ == "__main__":
    MyFrame().mainloop()