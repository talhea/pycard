# 오페라 리포트 'trial balance'(어제 기준)를 내여받는다

import pyautogui as gui
import datetime
import os, shutil
import moving_to_folder as todown

def down():
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    
    gui.PAUSE = 0.5
    gui.sleep(3)

    # 오페라 작업표시줄 위치 클릭 - 활성화시킴
    gui.moveTo(528, 876)
    gui.click()

    # 이미 로그인 된 상태이므로 PMS 버튼 클릭: PMS 로그인 화면 로딩
    gui.moveTo(738, 390)
    gui.click()
    gui.sleep(8)

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
    gui.moveTo(1090, 756)
    gui.click()

    # 파일이 저장될 시간 필요
    gui.sleep(15)

    # 오페라 종료
    gui.moveTo(368, 175)
    gui.doubleClick()
    gui.sleep(8)

    # 최소화
    gui.moveTo(1224, 17)
    gui.click()

    gui.sleep(3)

    # 다운로드 받는 파일을 down디렉토리로 이동
    down_base_dir = "./downdata/" + target_date         # 어제날짜로 된 폴더 ./downdata/YYYYMMDD/
    target_file_name += ".xml"                          # 파일 trial_balanceYYYYMMDD.xml

    todown.to_downdata(down_base_dir, target_file_name) # 파일 이동


if __name__ == '__main__':
    down()