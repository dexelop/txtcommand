import openpyxl
from openpyxl.utils import coordinate_from_string, column_index_from_string
import pyautogui as pa
from time import sleep
import pyperclip


class MyAction():
    def __init__(self):
        pass

    def enter(self, times=1):
        for i in range(times):
            pa.hotkey("enter")
            print('Hit enter!')
            sleep(0.3)
            i+=1

    def spacebar(self, times=1):
        for i in range(times):
            pa.hotkey("space")
            print('Hit spacebar!')
            sleep(0.3)
            i+=1

    def sleep(self, times=1):
        sleep(times)
        print('Sleep')

    def movetoright(self, times=1):
        for i in range(times):
            pa.hotkey("right")
            print('Hit move to right!')
            sleep(0.3)
            i+=1
    def movetoleft(self, times=1):
        for i in range(times):
            pa.hotkey("left")
            print('Hit move to left!')
            sleep(0.3)
            i+=1

    def movetoup(self, times=1):
        for i in range(times):
            pa.hotkey("up")
            print('Hit move to up!')
            sleep(0.3)
            i+=1

    def movetodown(self, times=1):
        for i in range(times):
            pa.hotkey("down")
            print('Hit move to down')
            sleep(0.3)
            i+=1

    def copyandpaste(self, sheet, column, i=1):
        # print(column+str(i))
        # print(sheet[column+str(i)].value)
        # pa.typewrite(sheet[column + str(i)].value)
        pyperclip.copy(sheet[column + str(i)].value)
        pa.hotkey('ctrl','v')
        print("copyandpaste is Work!")
        sleep(0.3)

    def type(self, t):
        pyperclip.copy(t)
        pa.hotkey('ctrl', 'v')
        print(t+'를 입력하셨습니다.')
        sleep(0.3)

    def click(self, pos_x, pos_y):
        pa.click(pos_x, pos_y)
        print(str(pos_x)+","+str(pos_y))
        sleep(0.3)

    def move_to(self, pos_x, pos_y):
        pa.moveTo(pos_x, pos_y, duration=0.3)
        print('Move! Move!')
        sleep(0.3)

    def passing(self):
        print('Say Pass!!!')

    def run_func(self, scrt_text, sheet=None, index=None):
        scrt_text = scrt_text.split("\n")
        # print(scrt_text)
        func_list = []
        for func in scrt_text:
            func = func.strip("\n").lower()
            func_list.append(func)
            # print(func)

        for func in func_list:
            if func.startswith('{{') & func.endswith('}}'):
                func = func[2:-2]
                try:
                    print('run self.{}'.format(func))
                    exec('self.' + func)
                    # exec('self.'+func + '()')
                except:
                    print(func+' 함수는 없습니다.')
            elif func.startswith('[[') & func.endswith(']]'):
                if index == 0:
                    # print(count)
                    print('yeh it is ok')
                    func = func[2:-2]
                    try:
                        exec('self.' + func)
                    except:
                        print('붙여넣기 에러 발생')
            
        # print(func_list)
        # text2save = func_list
        # return text2save

            # elif func.startswith('[[') & func.endswith(']]'):
            #     r = 1
            #     pos = func[2:-2]
            #     try:
            #         exec('self.ctrl_v(' + pos + str(r) + ')')   # self.ctrl_v(pos+r)
            #         print('copy_and_paste')
            #     except:
            #         print('붙여넣기 에러 발생')

            # elif func.startswith('((') & func.endswith('))'):
            #     try:
            #         print('excel column is ..')
            #     except:
            #         print('exception!')
