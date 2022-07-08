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

    # 크롬 창 세번쨰 탭,  민국은행 선택
    gui.moveTo(600, 17)
    gui.click()
    gui.sleep(1)
    
    # 상단 로그인 버튼 클릭
    gui.moveTo(188, 98)
    gui.click()
    gui.sleep(4)

    # 아이디/암호 입력창 입력
    gui.moveTo(875, 610)
    gui.click()
    gui.write('ibis11', interval=0.4)    # 아이디 입력
    gui.click(1236, 658)                # 마우스 입력 선택
    gui.sleep(1)
    gui.click('./img/b.png')            # 암호 입력, 키보드 않먹힘
    gui.click('./img/c.png')
    gui.sleep(0.5)
    gui.click('./img/c.png')
    gui.click('./img/1.png')
    gui.click('./img/0.png')
    gui.click('./img/1.png')
    gui.click('./img/0.png')
    gui.click('./img/sp1.png')
    gui.click('./img/gol.png')
    gui.click('./img/sp2.png')
    gui.click('./img/1.png')
    gui.click(1120, 715)
    gui.click(3)

    # 조회메뉴 클릭
    gui.moveTo(400, 165)                # 조회 메뉴 이동만
    gui.sleep(1)
    gui.moveTo(400, 293, duration=1)    # 아래로 이동해서
    gui.click(560, 293, duration=1)     # 오른쪽 조회 버튼으로 이동 클릭
    gui.sleep(3)

    # 조회메뉴 셋팅
    gui.scroll(-220)
    gui.moveTo(630,480)
    gui.click()
    gui.sleep(1)
    gui.click(630, 525)
    gui.sleep(1)
    gui.click(630, 540)
    gui.sleep(1)
    gui.click(630, 590)
    gui.sleep(1)

    # 조회버튼 클릭
    gui.moveTo(789, 780)
    gui.click()
    gui.sleep(3)

    # 내용 확인
    gui.scroll(-220)

    # 파일 저장 버튼
    gui.moveTo(1315, 840)
    gui.click()
    gui.sleep(2)

    # 다운로드 결과창 끄기
    gui.moveTo(1579, 832)
    gui.click()

    # 원래 위치로 마우스 스크롤 업 1회
    gui.scroll(440)
    gui.sleep(1)

    # 로그 아웃
    gui.moveTo(430, 98)
    gui.click()
    gui.sleep(1)

    # 최소화
    gui.moveTo(1478, 11)
    gui.click()

    gui.sleep(2)


    # 이동시킬 파일 이름 리스트
    # 자료를 다운로드 받는 down디렉토리 확인 (및 생성)
    target_file_name = receips_date + '.xls'    # 20220701.xls
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
    target_file_name = "C:/users/fa2/downloads/" + target_file_name     # 파일 위치 : c:/users/fa2/20220701.xml
    shutil.move(target_file_name, down_base_dir)                        # ./downdata/20220701/ 폴더로 이동


if __name__ == '__main__':
    down()