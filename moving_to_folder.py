import datetime
import os, shutil

def to_downdata(target_dir: str, target_file: str) -> None:
    """다운로드 받은 파일을 목적지 디렉토리로 이동

    Args:
        target_dir (str): 목적지 디렉토리 ('./downdata/YYYYMMDD/' 형식)
        target_file (str): 이동시켜야할 다운로드 받은 파일
    """    

    # 목적지 down디렉토리 확인하고 없으면 생성
    try:
        if os.path.exists(target_dir) == False:     # 폴더가 없으면 생성
            os.makedirs(target_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:      # error 로그 파일에 추가
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}'
            )

    # 파일 옮기기 : 'downdata/날짜' 디렉토리로
    target_file = 'C:/users/fa2/downloads/' + target_file       # 파일 위치 : c:/users/fa2/{target_file}
    
    try:
        shutil.move(target_file, target_dir)                        # ./downdata/YYYYMMDD/ 폴더로 이동
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}'
            )


if __name__ == '__main__':
    to_downdata()