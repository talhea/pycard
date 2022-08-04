"""
다운로드 받은 KICC 입금 현황을 dataframe으로 저장한다
"""  

import pandas as pnds
import pickle, os
import datetime

def to_receipts_history_df(work_date):
    """다운로드 받은 KICC 입금 현황을 Dataframe으로 저장한다
    Pandas.read_excel함수로 읽어서 Dataframe으로 만들고, 필요한 컬럼만 추출해서 dfdata디렉토리에 저장한다
    
    Args:
        work_date (datetime): 오늘 날짜
    """    

    # 1. 입금현황 엑셀 파일을 읽어서 DataFrame 생성
    target_date = (work_date - datetime.timedelta(1)).strftime("%Y%m%d")    # 어제 날짜 포맷
    
    downdata_dir = f'./data/{target_date}/downdata/'                         # 읽어들일 down디렉토리 './data/YYYYMMDD/downdata/'
    receipts_filename = downdata_dir + '입금현황 · 일별.xlsx'                   # KICC 입금현황 엑셀파일
    
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
                f'[kicc_receipts - Reading Data] <{datetime.datetime.now()}> Pandas excel-reading error ({receipts_filename}) ===> {e}\n'
            )
        # 기능 에러가 아닌 빈 거래내역 파일 조차도 없는 경우도 문제이므로 프로그램 종료
        raise(e)

    # 2. '입금예정일자' 컬럼명 'date'로 수정 : 다른 dataframe과 통일
    receipts_history_df.rename(columns={'입금예정일자': 'date'}, inplace=True)

    # 3. 필요한 컬럼만 추출
    receipts_history_df = receipts_history_df[['카드사', 'date', '접수금액', '합계금액', '수수료', '입금예정액']]

    # 4. df디렉토리에 dataframe 저장
    # 4-1. 목적지 df디렉토리 확인하고 없으면 생성
    dfdata_dir = f'./data/{target_date}/dfdata/'        # 데이타를 저장할 df디렉토리 './data/YYYYMMDD/dtdata/'
    
    try:
        if os.path.exists(dfdata_dir) == False:         # 폴더가 없으면 생성
            os.makedirs(dfdata_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:          # error 로그 파일에 추가
            file.write(
                f'[kicc_receipts - Making dfdata] <{datetime.datetime.now()}> mkdir error ({dfdata_dir}) ===> {e}\n'
            )
        raise(e)

    # 4-2. dataframe 저장
    #   입금내역의 저장 날짜는 오늘 날짜로
    receipts_date = work_date.strftime("%Y%m%d")      # 오늘 날짜
    
    try:
        df_filename = 'dt_kicc_receips_' + receipts_date      # 저장할 dataframe 파일 이름 'dt_kicc_receipts_YYYYMMDD'

        with open(dfdata_dir + df_filename, "wb") as file:
            pickle.dump(receipts_history_df, file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_receipts - Making Dataframe] <{datetime.datetime.now()}> pickle.dump error ({df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 4-3. dataframe 그대로 excel로 저장 => 추후에 사용 가능할수 있음
    try:
        xl_filename = 'df_kicc_receipts_' + receipts_date + '.xlsx'    # 저장파일 'df_kicc_receipts_YYYYMMDD.xlsx
        
        with pnds.ExcelWriter(dfdata_dir + xl_filename, mode='w', engine='openpyxl') as writer:
            receipts_history_df.to_excel(writer, sheet_name='original', index=False)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_history.py - Making Excel] <{datetime.datetime.now()}> Pandas.ExcelWriter error ({xl_filename}) ===> {e}\n'
            )
        raise(e)


if __name__ == '__main__':
    to_receipts_history_df(datetime.datetime.now())