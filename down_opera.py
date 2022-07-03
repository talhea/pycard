# 오페라 리포트 'trial balance'(어제 기준)를 내여받는다

import pyautogui as gui
import datetime
import os, shutil

# 어제날짜 추출
target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime(
    "%Y%m%d"
)  # 어제날짜 이름

gui.PAUSE = 0.5
gui.sleep(3)

# 오페라 작업표시줄 위치 클릭 - 활성화시킴
gui.moveTo(528, 876)
gui.click()

# 이미 로그인 된 상태이므로 PMS 버튼 클릭: PMS 로그인 화면 로딩
gui.moveTo(769, 406)
gui.click()
gui.sleep(5)

# 로그인 창에서 로그인 클릭: PMS 화면 로딩
gui.moveTo(888, 595)
gui.click()
gui.sleep(8)

# Mescelleneouse 클릭
gui.moveTo(1208, 183)
gui.click()
gui.sleep(1)

# Reports 버튼 클릭
gui.moveTo(365, 261)
gui.click()
gui.sleep(2)

# 검색창 입력
gui.moveTo(725, 305)
gui.click()
gui.sleep(1)
gui.write("%trial", interval=0.4)
gui.press("enter")
gui.sleep(1.5)

# trial_balance가 선택된 상태이므로 파일 체크 시작
gui.moveTo(752, 673)
gui.click()

# file format 선택 : XML
gui.moveTo(1027, 673)
gui.click()
gui.click(1027, 750, duration=1)

# OK버튼 핫키
gui.hotkey("alt", "o")

# file 이름 변경
gui.hotkey("alt", "f")
gui.sleep(1)
gui.moveTo(692, 546)
gui.doubleClick()
gui.sleep(1)

target_file_name = "trial_balance" + target_date     # trial_balance20220701
gui.write(target_file_name, interval=0.4)
gui.sleep(2)

gui.moveTo(1005, 547)
gui.click()
gui.sleep(2)

# 검색창 닫기
# gui.hotkey('alt' + 'c')   # 이 때는 핫키가 안먹힘
gui.moveTo(1090, 756)
gui.click()

# 파일이 저장될 시간 필요
gui.sleep(15)

# 오페라 종료
gui.moveTo(368, 175, duration=1)
gui.doubleClick()
gui.sleep(8)

# 최소화
gui.moveTo(1224, 17, duration=1)
gui.click()

gui.sleep(3)

# 자료를 다운로드 받는 down디렉토리 확인 (및 생성)
down_base_dir = "./downdata/" + target_date     # 어제날짜로 된 폴더 ./downdata/20220701/
try:
    if os.path.exists(down_base_dir) == False:  # 폴더가 없으면 생성
        os.makedirs(down_base_dir)
except:
    with open("./error.log", "a") as file:
        file.write(
            f"[down_opera.py] : {datetime.datetime.now()} - Exception: Cannot create the directory {down_base_dir}"
        )
        print("Error: Cannot create the directory {}".format(down_base_dir))

# 파일 옮기기 : downdata디렉토리로
# 디렉토리의 파일만 검색해서 .. 파일 이름에서 유추... 해당 파일을 복사/이동
target_file_name = "C:/users/fa2/" + target_file_name + ".xml"      # 파일 위치 : c:/users/fa2/trial_balance20220701.xml
shutil.move(target_file_name, down_base_dir)                        # ./downdata/20220701/ 폴더로 이동