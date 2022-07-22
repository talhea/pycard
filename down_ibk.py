# 기업은행 입금내역(오늘 기준)를 내여받는다

import imp
from tkinter import E
import pyautogui as gui
import datetime
import os, shutil
import moving_to_folder as todown
from dateutil.relativedelta import relativedelta

def down(work_date):
    """IBK 기업은행 당일 거래 내역 출력

    Args:
        work_date (datetime): 오늘 날짜
    """    
    receipts_date = work_date.strftime("%Y%m%d")                           # 오늘 입금 내역
    
    gui.PAUSE = 0.5
    gui.sleep(3)

    # 크롬 작업표시줄 위치 클릭 - 활성화시킴
    gui.moveTo(650, 880)
    gui.click()

    # 크롬 창 첫번쨰 탭,  기업은행 선택
    gui.moveTo(130, 17)
    gui.click()
    gui.sleep(5)

    # 화면 상단 로그인 버튼
    gui.moveTo(557, 112)
    gui.click()
    gui.sleep(8)

    # 공인인증서 로그인 버튼
    gui.moveTo(1068, 471)
    gui.click()
    gui.sleep(8)

    # 암호 입력창 입력
    gui.moveTo(688, 654)
    gui.click()
    gui.write('ibisbusan@1', interval=0.4)    # 인증서 암호 입력
    gui.press('enter')
    gui.sleep(5)

    # 거래내역조회 클릭
    gui.moveTo(282, 581)
    gui.click()
    gui.sleep(3)

    gui.click(300, 581)
    gui.sleep(3)

    # 입금내역 클릭
    gui.moveTo(714, 553)
    gui.click()

    # 날짜 셋팅    
    gui.click(620, 625)
    gui.press('tab')
    
    #   현재 날짜와 work_date위 날짜 차이 : diff_dt.[years, months, days]
    #   그 차이만큼 'UP'키를 눌러서 셋팅한다
    diff_dt = relativedelta(datetime.datetime.now(), work_date)
    
    if diff_dt.years != 0:
        gui.press('up', presses=diff_dt.years)
    gui.press('tab')
    if diff_dt.months != 0:
        gui.press('up', presses=diff_dt.months)
    gui.press('tab')
    if diff_dt.days != 0:
        gui.press('up', presses=diff_dt.days)
    
    gui.press('tab')
    gui.press('tab')
    
    if diff_dt.years != 0:
        gui.press('up', presses=diff_dt.years)
    gui.press('tab')
    if diff_dt.months != 0:
        gui.press('up', presses=diff_dt.months)
    gui.press('tab')
    if diff_dt.days != 0:
        gui.press('up', presses=diff_dt.days)

    # 조회버튼 클릭
    gui.moveTo(912, 714)
    gui.click()
    gui.sleep(3)

    # 마우스 스크롤 다운 1회
    gui.moveTo(912, 714)
    gui.scroll(-220)

    # 파일 저장 버튼
    gui.moveTo(723, 750)
    gui.click()
    gui.sleep(2)

    # '텍스트파일저장' 선택
    gui.moveTo(1004, 340)
    gui.click()
    gui.sleep(3)

    # 다운로드 결과창 끄기
    gui.moveTo(1579, 840)
    gui.click()

    # 파일 창 끄기
    gui.moveTo(1196, 265)
    gui.click()
    gui.sleep(1)

    # 원래 위치로 마우스 스크롤 업 1회
    gui.scroll(220)
    gui.sleep(1)

    # 로그 아웃
    gui.moveTo(762, 111)
    gui.click()
    gui.sleep(1)

    # 최소화
    gui.moveTo(1478, 11)
    gui.click()

    gui.sleep(2)


    # # 다운로드 받는 파일을 downdata디렉토리로 이동
    # down_base_dir = 'C:/Users/FA2/Downloads/'                                               # 파일이 다운로드된 디렉토리
    
    # #   source 파일명 : 당 일 입금 내역의 조회가 아니라면, 실제 내역의 날짜와는 달리 오늘 날짜로 저장되므로 이를 보정(receipts_date을 사용하지 않는 이유)
    # source_filename = '거래내역조회_입출식 예금' + datetime.datetime.now().strftime('%Y%m%d') + '.txt'  # 오늘 날짜로 된 source 파일 이름 '거래내역조회_입출식 예금YYYYMMDD.txt'
    
    # #   목적지 위치: 인수로 들어온 날짜는... 폴더 이름은 하루 전 날짜로, 파일 이름은 날짜 그대로(당일 날짜) 셋팅
    # target_date = (work_date - datetime.timedelta(1)).strftime("%Y%m%d")                    # 폴더 이름용 어제 날짜 포맷
    # target_dir = f'./data/{target_date}/downdata/거래내역조회_입출식 예금{receipts_date}.txt' # 당일 처리 내역이 아닐 경우를 위해서 파일명 지정(변경)
    #                                                                                         # 목적지 : './data/YYYYMMDD/downdata/거래내역조회_입출식 예금YYYYMMDD.txt'
    
    # #   파일 옮기기
    # todown.to_downdata(down_base_dir + source_filename, target_dir)


# if __name__ == '__main__':
#     down()