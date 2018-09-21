import openpyxl
from openpyxl.utils import coordinate_from_string, column_index_from_string
# xy = coordinate_from_string('A4') # returns ('A',4)
# col = column_index_from_string(xy[0]) # returns 1
# row = xy[1]

# filename = 'sample.txt'


class MyAction():
    def __init__(self):
        pass

    def enter(self, times=1):
        for i in range(times):
            print('Hit enterKey!')
            i+=1

    def spacebar(self, times=1):
        for i in range(times):
            print('Hit spacebar!')
            i+=1

    def sleep(self, times=1):
        for i in range(times):
            print('Sleep')
            i+=1

    def movetoright(self, times=1):
        for i in range(times):
            print('Hit move to right!')
            i+=1

    def movetoleft(self, times=1):
        for i in range(times):
            print('Hit move to left!')
            i+=1

    def movetoup(self, times=1):
        for i in range(times):
            print('Hit move to up!')
            i+=1

    def movetodown(self, times=1):
        for i in range(times):
            print('Hit move to down')
            i+=1

    def copyandpaste(self,sheet, column, i=1):
        # print(column)
        # print(i)
        print(column+str(i))
        print(sheet["a"+str(i)].value)

    def type(self, t):
        print(t+'를 입력하셨습니다.')

    def click(self, pos_x, pos_y):
        print(+pos_x+pos_y)

    def run_func(self, scrt_text):
        scrt_text = scrt_text.split("\n")
        # print(scrt_text)
        for func in scrt_text:
            func = func.strip("\n").lower()
            print(func)

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


"""
class MyAction2():
    def __init__(self):#, filename=filename):
        pass
        # with open(filename, 'r+') as f:
        #     self.func_list = f.readlines()
        #     print('range는 ' + self.func_list[0].strip('\n')+ ' 입니다.')
        #     self.range = self.func_list[2].strip('\n')
        #     self.func_list = self.func_list[3:]
        #
        #     input_range = self.range    # 'a1:d5'
        #     _ranges = input_range.split(':')
        #     _ranges_start = _ranges[0]
        #     _ranges_end = _ranges[1]
        #
        #     start_xy = coordinate_from_string(_ranges_start)
        #     start_col = column_index_from_string(start_xy[0])  # returns 1
        #     start_row = start_xy[1]
        #
        #     end_xy = coordinate_from_string(_ranges_end)
        #     end_col = column_index_from_string(end_xy[0])  # returns 1
        #     end_row = end_xy[1]
        #
        #     self.start_col = start_col
        #     self.start_row = start_row
        #     self.end_col = end_col
        #     self.end_row = end_row
        #
        #     self.alist = [ start_col, start_row, end_col, end_row]


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

    def click(self, pos_x, pos_y):
        print('I clicked!,'+pos_x+pos_y)

    def ctrl_v(self, pos):
        print(pos.value)


    def type(self, t):
        print(t+'를 입력하셨습니다.')

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
            
"""