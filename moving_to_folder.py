import datetime
import os, shutil

def to_downdata(source_dir: str, target_file: str) -> None:
    """다운로드 받은 파일을 목적지 디렉토리로 이동
    목적지 디렉토리 : './downdata/YYYYMMDD/'
    
    Args:
        sourcr_dir (str): 원래 파일 위치
        target_file (str): 이동시켜야할 다운로드 받은 파일
    """    
    
    if (source_dir == '') or (target_file == ''):     # 함수 인자 확인
        print('2 arguments needed')
        exit()
    
    # 목적지 down디렉토리 확인하고 없으면 생성
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    target_dir = './downdata/' + target_date + '/'                                      # './downdata/YYYYMMDD'
    
    try:
        if os.path.exists(target_dir) == False:     # 폴더가 없으면 생성
            os.makedirs(target_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:      # error 로그 파일에 추가
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {target_dir} : {e}\n'
            )
        raise(e)

    # 목적지 디렉토리로 파일 옮기기
    try:        
        shutil.move(source_dir + target_file, target_dir)       # ./downdata/YYYYMMDD/ 폴더로 이동
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> move error {target_file} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> move error {target_file} : {e}\n'
            )
        raise(e)


# if __name__ == '__main__':
#     to_downdata(...)