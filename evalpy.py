# ===================
# 1. sample txt 를 불러 온다
#     sample txt 에는
#       엑셀의 rows 범위를 읽어 오고, 그 범위 만큼 시도한다
#       명령어 뭉치 들이 있다
#           명령어중 command 와 위치 지정에 대한 내용이 있다.
#
#
# 1
# 1
#
# ===================



import openpyxl
from openpyxl.utils import coordinate_from_string, column_index_from_string
# xy = coordinate_from_string('A4') # returns ('A',4)
# col = column_index_from_string(xy[0]) # returns 1
# row = xy[1]


class DoReMi():
    def __init__(self, filename="sample.txt"):
        with open(filename, 'r+') as f:
            self.func_list = f.readlines()
            print('range는 ' + self.func_list[0].strip('\n')+ ' 입니다.')
            self.range = self.func_list[2].strip('\n')
            self.func_list = self.func_list[3:]

            input_range = self.range    # 'a1:d5'
            _ranges = input_range.split(':')
            _ranges_start = _ranges[0]
            _ranges_end = _ranges[1]

            start_xy = coordinate_from_string(_ranges_start)
            start_col = column_index_from_string(start_xy[0])  # returns 1
            start_row = start_xy[1]

            end_xy = coordinate_from_string(_ranges_end)
            end_col = column_index_from_string(end_xy[0])  # returns 1
            end_row = end_xy[1]

            self.start_col = start_col
            self.start_row = start_row
            self.end_col = end_col
            self.end_row = end_row

            self.alist = [ start_col, start_row, end_col, end_row]


        # print('TODO : 엑셀에서 범위를 읽어 와야 합니다.')
        print(self.range)
        print('range 를 숫자로 하면')
        print(self.alist)

    def print_f1_lst(self):
        print(f1_lst)

    def lst(self):
        print(self.func_list)

    def do(self):
        print("Do")

    def re(self):
        print("Re")

    def mi(self):
        print("Mi")
        print("구현해야 할 함수는 \n"
              "enter, spacebar, type, ctrl_v")

    def pos_click(self):
        print('I clicked!, TODO : get POS, and when you start program  move to POS and click')

    def ctrl_v(self, pos):
        print(pos.value)

    def enterkey(self, times=1):
        for i in range(times):
            print('Hit enterKey!')
            i+=1

    def type(self, t):
        print(t)

    def run_func(self):
        for func in self.func_list:
            func = func.strip("\n").lower()
            if func.startswith('{{') & func.endswith('}}'):
                func = func[2:-2]
                try:
                    exec('self.' + func)
                    # exec('self.'+func + '()')
                except:
                    print(func+' 함수는 없습니다.')
            elif func.startswith('[[') & func.endswith(']]'):
                r = 1
                pos = func[2:-2]
                try:
                    exec('self.ctrl_v(' + pos + str(r) + ')')   # self.ctrl_v(pos+r)
                    print('copy_and_paste')
                except:
                    print('붙여넣기 에러 발생')

            elif func.startswith('((') & func.endswith('))'):
                try:
                    print('excel column is ..')
                except:
                    print('exception!')


if __name__ == '__main__':
    song = DoReMi()
    song.run_func()