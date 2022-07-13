"""
다운로드 받은 KICC 입금 현황을 dataframe으로 저장한다
"""  

import pandas as pnds
import pickle, os
import datetime

def to_receipts_history_df():
    """다운로드 받은 KICC 입금 현황을 dataframe으로 저장한다
    Pandas.read_excel함수로 읽어서 dataframe으로 만들고, 필요한 컬럼만 추출해서 dt디렉토리에 저장한다
    """    

    # 1. 입금현황 엑셀 파일을 읽어서 DataFrame 생성
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    downdata_dir = './downdata/' + target_date + '/'                                   # 읽어들일 down디렉토리 './downdata/YYYYMMDD'
    receipts_filename = downdata_dir + '입금현황 · 일별.xlsx'                          # KICC 입금현황 엑셀파일
    
    try:
        receipts_history_df = pnds.read_excel(receipts_filename, header=0, thousands = ',', # 금액 천단위 구분자 ',' 고려
                                        dtype={'가맹점번호': str, '접수건수': 'int32',      # 필요 컬럼의 데이터 타입 고정
                                                '접수금액': 'int64', '합계건수': 'int32',
                                                '합계금액': 'int64', '미입금건수': 'int32',
                                                '미입금액': 'int64', '수수료': 'int64',
                                                '입금예정액': 'int64'})
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> Pandas file-reading error {receipts_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> Pandas file-reading error {receipts_filename} : {e}'
            )
        
        # 파일 읽기에 실패했으므로 이 함수 과정만 종료, 프로그램 전체를 종료시키지 않는다(raise(e)를 사용하지 않는 이유)
        return

    # 2. '입금예정일자' 컬럼명 수정 : 다른 dataframe과 통일
    receipts_history_df.rename(columns={'입금예정일자': 'date'}, inplace=True)     # '입금예정일자' 컬럼명 'date'로 수정

    # 3. 필요한 컬럼만 추출
    receipts_history_df = receipts_history_df[['카드사', 'date', '접수금액', '합계금액', '수수료', '입금예정액']]

    # 4. dt디렉토리에 dataframe 저장
    dtdata_dir = './dtdata/' + target_date + '/'       # 데이타를 저장할 dt디렉토리 './dtdata/YYYYMMDD/'
    
    # 4-1. 목적지 dt디렉토리 확인하고 없으면 생성
    try:
        if os.path.exists(dtdata_dir) == False:        # 폴더가 없으면 생성
            os.makedirs(dtdata_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:          # error 로그 파일에 추가
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {dtdata_dir} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {dtdata_dir} : {e}'
            )
        raise(e)

    # 4-2. dataframe 저장
    try:
        df_filename = 'dt_kicc_receips_' + target_date      # 저장할 dataframe 파일 이름 'dt_kicc_receipts_YYYYMMDD'

        with open(dtdata_dir + df_filename, "wb") as file:
            pickle.dump(receipts_history_df, file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle.dump error {df_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle.dump error {df_filename} : {e}'
            )
        raise(e)


if __name__ == '__main__':
    to_receipts_history_df()