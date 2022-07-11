
# KICC site로부터 '신용카드 거래내역'과 '입금현황' 총 2개 파일을 내려받는다

import pyautogui as gui
import datetime
import os, shutil
import moving_to_folder as todown

def down():
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    
    gui.PAUSE = 0.5
    gui.sleep(3)

    # 작업표시줄 크롬 창 활성화
    gui.moveTo(650, 880)
    gui.click()

    # 크롬 창 두번쨰 탭, KICC 선택
    gui.moveTo(369, 17)
    gui.click()

    # KICC 로그인
    gui.moveTo(1203, 104)
    gui.click()
    gui.sleep(10)

    # 거래내역조회 더블클릭
    gui.moveTo(115, 533)
    gui.doubleClick()
    gui.sleep(6)

    # 마우스 스크롤해서 매출/입금 내역 더블클릭
    gui.moveTo(115, 533)
    gui.scroll(-55) # 클릭한번에 55 이동
    gui.scroll(-55) # 클릭한번에 55 이동
    gui.scroll(-55) # 클릭한번에 55 이동
    gui.sleep(2)
    gui.doubleClick()
    gui.scroll(55) # 클릭한번에 55 이동, 원래 위치로
    gui.scroll(55) # 클릭한번에 55 이동, 원래 위치로
    gui.scroll(55) # 클릭한번에 55 이동, 원래 위치로
    gui.sleep(6)

    # 매출/입금 내역 조회
    gui.moveTo(1514, 264)
    gui.click()
    gui.sleep(4)

    # 확인창 확인 버튼
    gui.moveTo(971, 167)
    gui.click()
    gui.sleep(2)

    # 매출/입금 내역 엑셀 저장 , '입금현황 · 일별.xlsx'
    gui.moveTo(1466, 264)
    gui.click()
    gui.sleep(3)

    # 거래내역 조회 탭 선택
    gui.moveTo(357, 134)
    gui.click()
    gui.sleep(1)

    # from 거래일시 하루 전으로 셋팅
    gui.moveTo(427, 239)
    gui.click()
    gui.sleep(1)

    gui.write(target_date, interval=0.4)    # 결론적으로 어제부터 오늘까지 내역
    gui.sleep(1)

    # 조회
    gui.moveTo(1516, 240)
    gui.click()
    gui.sleep(6)

    # 확인창 확인버튼 클릭
    gui.moveTo(971, 167)
    gui.click()

    # 엑셀 버턴 클릭
    gui.moveTo(1519, 266)
    gui.click()
    gui.sleep(3)

    # 다운로드 결과창 끄기
    gui.moveTo(1579, 832)
    gui.click()

    # 로그아웃
    gui.moveTo(53, 157)
    gui.click()
    gui.sleep(2)

    # 확인창 확인버튼
    gui.moveTo(904, 160)
    gui.click()
    gui.sleep(1)

    # 최소화
    gui.moveTo(1478, 11)
    gui.click()

    gui.sleep(2)

    # 다운로드 받는 파일을 down디렉토리로 이동
    down_base_dir = "./downdata/" + target_date             # 어제날짜로 된 폴더 ./downdata/YYYYMMDD/

    target_file_name = '입금현황 · 일별.xlsx'               # 입금파일: 오늘날짜
    todown.to_downdata(down_base_dir, target_file_name)     # 파일 이동

    target_file_name = '신용거래내역조회.xlsx'              # 신용거래내역: 어제날짜
    todown.to_downdata(down_base_dir, target_file_name)     # 파일 이동


if __name__ == '__main__':
    down()