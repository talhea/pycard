'''
어제 날짜 폴더 생성('./YYYYMMDD/')
'''

import os
import datetime

def ready():
    """어제 날짜 폴더 생성('./YYYYMMDD/')
    다운로드 및 가공된 데이터들을 위한 디렉토리(downdata, dtfata)내에 어제 날짜로 된 폴더 생성
    './downdata/YYYYMMDD', './dtdata/YYYYMMDD'
    """    

    # 1. 디렉토리 생성 : 어제날짜로 된 이름들
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷

    # 1-1. 자료를 다운로드 받는 down디렉토리 생성
    down_base_dir = './downdata/' + target_date + '/'
    try:
        if os.path.exists(down_base_dir) == False:
            os.makedirs(down_base_dir)
    except:
        with open('./error.log', 'a') as file:
            file.write(f'{datetime.datetime.now()} - Exception: Cannot create the directory {down_base_dir}')
            print("Error: Cannot create the directory {}".format(down_base_dir))

    # 1-2. python파일에서 생성하는 data를 저장하는 dt디렉토리 생성
    dt_base_dir = './dtdata/' + target_date + '/'
    try:
        if os.path.exists(dt_base_dir) == False:
            os.makedirs(dt_base_dir)
    except:
        with open('./error.log', 'a') as file:
            file.write(f'{datetime.datetime.now()} - Exception: Cannot create the directory {dt_base_dir}')
            print("Error: Cannot create the directory {}".format(dt_base_dir))

    # 이후, KICC에서 자료 다운로드, 오페라에서 자료 다운로드, 기업은행에서 자료 다운로드

if __name__ == '__main__':
    ready()