# 필요 다운로드 자료를 저장하기 위한 디렉토리 만들기

import os
import datetime

from nbformat import read

def ready():
    # 1. 디렉토리 생성 : 어제날짜로 된 이름들
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷

    # 1-1. 자료를 다운로드 받는 down디렉토리 생성
    downdata_dir = './downdata/' + target_date
    try:
        if os.path.exists(downdata_dir) == False:
            os.makedirs(downdata_dir)
    except:
        with open('./error.log', 'a') as file:
            file.write(f'{datetime.datetime.now()} - Exception: Cannot create the directory {downdata_dir}')
            print("Error: Cannot create the directory {}".format(downdata_dir))

    # 1-2. python파일에서 생성하는 data를 저장하는 df디렉토리 생성
    dfdata_dir = './dfdata/' + target_date
    try:
        if os.path.exists(dfdata_dir) == False:
            os.makedirs(dfdata_dir)
    except:
        with open('./error.log', 'a') as file:
            file.write(f'{datetime.datetime.now()} - Exception: Cannot create the directory {dfdata_dir}')
            print("Error: Cannot create the directory {}".format(dfdata_dir))

    # 이후, KICC에서 자료 다운로드, 오페라에서 자료 다운로드, 기업은행에서 자료 다운로드

if __name__ == '__main__':
    ready()