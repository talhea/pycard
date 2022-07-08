# 기업은행 입금내역(오늘 기준)를 내여받는다

import pyautogui as gui
import datetime
import os, shutil

def down():
    receips_date = datetime.datetime.now().strftime("%Y%m%d")                           # 오늘 입금 내역
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    
    gui.PAUSE = 0.5
    gui.sleep(3)

    # 오페라 작업표시줄 위치 클릭 - 활성화시킴
    gui.moveTo(650, 880)
    gui.click()

    # 크롬 창 첫번쨰 탭,  기업은행 선택
    gui.moveTo(130, 17)
    gui.click()
    gui.sleep(1)

    # 화면 상단 로그인 버튼
    gui.moveTo(557, 112)
    gui.click()
    gui.sleep(5)

    # 공인인증서 로그인 버튼
    gui.moveTo(1068, 471)
    gui.click()
    gui.sleep(5)

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

    gui.moveTo(281, 581)
    gui.click()
    gui.sleep(3)

    # 입금내역 클릭
    gui.moveTo(714, 553)
    gui.click()

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
    gui.moveTo(1579, 832)
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


    # 이동시킬 파일 이름 리스트
    # 자료를 다운로드 받는 down디렉토리 확인 (및 생성)
    target_file_name = '거래내역조회_입출식 예금' + receips_date + '.txt'    # 거래내역조회_입출식 예금20220701.txt
    down_base_dir = "./downdata/" + target_date                         # 어제날짜로 된 폴더 ./downdata/20220701/
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
    target_file_name = "C:/users/fa2/downloads/" + target_file_name     # 파일 위치 : c:/users/fa2/trial_balance20220701.xml
    shutil.move(target_file_name, down_base_dir)                        # ./downdata/20220701/ 폴더로 이동


if __name__ == '__main__':
    down()