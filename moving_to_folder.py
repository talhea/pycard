import datetime
import os, shutil

def to_downdata(source_file: str, target_dir: str) -> None:
    """다운로드 받은 파일을 목적지 디렉토리로 이동
    목적지 디렉토리 : './downdata/YYYYMMDD/'
    
    Args:
        sourcr_file (str): 이동시켜야할 파일
        target_dir (str): 이동 목표
    """    
    
    if (source_file == '') or (target_dir == ''):           # 함수 인자 확인
        print('moving_to_folder.py : 인자 2개 필요')
        exit()
    
    try:
        dir_name = os.path.dirname(target_dir)              # 폴더 이름 추출
        if os.path.exists(dir_name) == False:               # 폴더가 없으면 생성
            os.makedirs(dir_name)
    except Exception as e:
        with open('./error.log', 'a') as file:              # error 로그 파일에 추가
            file.write(
                f'[moving_to_folder.py - Downloading Data] <{datetime.datetime.now()}> mkdir error {target_dir} ==> {e}\n'
            )
        raise(e)

    # 3. 목적지로 파일 옮기기
    try:        
        shutil.move(source_file, target_dir)                # 이동: './data/YYYYMMDD/downdata/{file}'
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[moving_to_folder.py - Moving Data] <{datetime.datetime.now()}> move error {target_dir}  ==> {e}\n'
            )
        raise(e)


# if __name__ == '__main__':
#     to_downdata(...)