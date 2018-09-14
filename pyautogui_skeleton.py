import pyautogui as pa
import pyperclip, time
from handle_xls_skeleton import *

print('''
************************************************

안녕하세요 서찬영 세무사입니다.
먼저 더존에서 자동전표>신용카드 또는 현금영수증 실행
(전체화면을 추천합니다)
문의 : 서찬영 세무사 : ctascy@gmail.com
정식 배포 버전이 아닙니다.

************************************************
''')

# user_delay = float(input("얼마마다 실행할까요? : "))
#
# wb = openpyxl.load_workbook(filename='hello.xlsx')
# ws = wb.active
# a1 = ws['a1'].value
# print(a1)


class MyPa:
    def __init__(self, delay=0.02, dic={}):
        self.delay = delay
        self.dic = dic
        # return self.delay, self.delay

    def click(self):

        pa.click()
        time.sleep(self.delay)

    def copy_n_paste(self):
        pa.hotkey("ctrl", "c")
        time.sleep(self.delay)
        paste = pyperclip.paste()
        time.sleep(self.delay)
        return paste

    def move_enterkey(self, n=1):
        for i in range(n):
            print('I hit ENTER')
            # pa.hotkey("enter")

            time.sleep(self.delay)
            i += 1


paaa = MyPa()
time.sleep(2)
paaa.move_enterkey(2)
print(paaa.copy_n_paste())




"""
def main():
    i=0
    click()
    while i<count:
        pos = pa.position()
        duzon = copy_paste()
        find_account(duzon)
        i +=1
        if not pa.position()==pos:
            print("커서가 움직였습니다.")
            break
        if i == count:
            print("완료했습니다.")


def click():
    pa.click()
    time.sleep(delay)


def copy_paste():
    #global duzon
    duzon = "" #있어야 하나?
    pa.hotkey("ctrl","c")
    time.sleep(delay)
    duzon = clipboard.paste()
    time.sleep(delay)
    return duzon

my_dic = {
"811":["식당","목우촌","어다리","불닭","해장국","파리바게뜨","비빔밥","한어울","나주곰탕","순두부","돈까스","커핀그루나루","김치찌개","박준뷰티랩","반점","국수","（주）춘산에프앤비","（주）춘산　관철동점","(주) 에스에프엔엠","스쿨푸드","껍데기","할머니","수산","일마지오","스테이크","조가비","양꼬치","샤브샤브","음식","보나비","곱창","꼼장어","갈비","냉면","샤브","롯데리아","칼국수","맥도날드","육개장","아비꼬","보쌈","족발","떡볶이","약국","덮밥","의원","헤어","리안","카페","김밥","덮밥","스타벅스","커피","부대찌개","이디야","피자",
"병원","카페","고깃집","비케이알","순대","맛집","쭈꾸미","만두","연탄구이","아비꼬","비비큐","봉구","아이스크림","치킨","김가네","면옥","공차","한우","애플하우스","이비인후과","삼계탕","파리바게트","명과","애플코코","탐앤탐스","편의점","찌개","스넥",
       "돈가스","본도시락","국밥","한정식","돈가스","장어구이,"],
"812":["한국스마트카드","대한항공","서울특별시시설관리공단","휴게소","한국도로공사",],
"813":["노래","호프","주점","투다리","새마을식당","피쉬","준코","비어","양꼬치","포차","CGV","막걸리",],
"814":["케이티","에스케이텔레콤",],
"821":["보험"],
"822":["LPG","주유소","에너지","충전소","SK네트웍스","석유","MOTOR","모터스",],
"824":["우정사업본부",],
"826":["조선일보","도서출판","교보문고",],
"830":["에스케이　플래닛","세븐일레븐","드림디포","자라리테일코리아주식회사","한국후지필름（주）","（주）신세계페이먼츠","이베이","인터파크","홈쇼핑","씨유","이마트","다이소","올앳","이랜드리테일","홈플러스","롯데쇼핑","하나마트","GS25","미니스톱","교보핫트랙스","이니스프리","문구","알파상사","바이더웨이","신세계"
     ,"아모레퍼시픽","아트박스","유통","이랜드","코리아세븐","공판장","지에스25","슈퍼","GS수퍼", ],
"831":["（주）케이에스넷","스튜디오",],
}

def find_account(duzon):
    for account in my_dic:
        for keyword in my_dic[account]:
            if keyword in duzon:
                move_enterkey(4)
                pa.typewrite(account)
                move_enterkey(5)
                return
    pa.hotkey("down")

def move_enterkey(times):
    for i in range(times):
        i =+1
        pa.hotkey("enter")
        time.sleep(delay)

if __name__=="__main__":
    main()
"""