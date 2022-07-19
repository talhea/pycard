# 필요 다운로드 자료를 저장하기 위한 디렉토리 만들기

from logging import shutdown
import os, datetime
import datetime
import shutil

from nbformat import read

def ready(work_date):
    """어제 날짜로 된 디렉토리 생성

    Args:
        work_date (datetime): 어제 날짜
    """
    # 0. 기존 error.log 백업
    target_date = work_date.strftime('%Y%m%d')  # 어제 날짜 포맷

    if os.path.isfile('./error.log'):
        shutil.move('./error.log', f'./error_{target_date}.log')
    
    # 1. 어제 날짜 디렉토리 생성
    yesterday_dir = f'./data/{target_date}/'            # './data/YYYYMMDD/'
    
    try:
        if os.path.exists(yesterday_dir) == False:
            os.makedirs(yesterday_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[preparation.py - downdata_dir] <{datetime.datetime.now()}> makedirs ({yesterday_dir}) error  ==> {e}\n'
            )
        raise(e)

    
    # 1-1. 자료를 다운로드 받는 down디렉토리 생성
    downdata_dir = yesterday_dir + 'downdata/'          # './data/YYYYMMDD/downdata'
    try:
        if os.path.exists(downdata_dir) == False:
            os.makedirs(downdata_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[preparation.py - downdata_dir] <{datetime.datetime.now()}> makedirs ({downdata_dir}) error  ==> {e}\n'
            )
            # return False
        raise(e)

    # 1-2. python파일에서 생성하는 data를 저장하는 df디렉토리 생성
    dfdata_dir = yesterday_dir + 'dfdata/'              # './data/YYYYMMDD/dfdata'
    try:
        if os.path.exists(dfdata_dir) == False:
            os.makedirs(dfdata_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[preparation.py - dfdata_dir] <{datetime.datetime.now()}> makedirs ({dfdata_dir}) error  ==> {e}\n'
            )
        raise(e)
    
    # print('실행 ', datetime.datetime.now().time())
    # return True
    # 이후, KICC에서 자료 다운로드, 오페라에서 자료 다운로드, 기업은행에서 자료 다운로드


if __name__ == '__main__':
    ready()