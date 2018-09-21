import openpyxl
from openpyxl.utils import coordinate_from_string, column_index_from_string


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
