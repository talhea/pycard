"""
다운로드 받은 KICC 신용카드 거래내역을 dataframe으로 저장한다
"""  

import pandas as pnds
import pickle, os
import datetime
import font_red_on_excel as red_font

def to_card_history_df(work_date):
    """다운로드 받은 KICC 신용카드 거래내역을 dataframe으로 저장한다
    xlsx 파일을 read_excel함수로 읽어들여 dataframe으로 만들고, 중복된 승인번호(당일 취소) 제거하고 날짜 포맷 셋팅 후 dataframe 으로 저장
    
    Args:
        work_date (datetime): 어제 날짜
    """    

    # 1. 엑셀 파일을 읽어서 DataFrame 생성
    target_date = work_date.strftime("%Y%m%d")      # 어제 날짜 포맷
    
    downdata_dir = f'./data/{target_date}/downdata/'                                        # 읽어들일 down디렉토리 './data/YYYYMMDD/downdata/'
    card_filename = downdata_dir + '신용거래내역조회.xlsx'                                  # KICC 카드 거래내역 엑셀파일
    
    try:
        card_history_df = pnds.read_excel(
                                        card_filename, header=0, thousands = ',',           # 첫번쨰 라인은 컬럼명(header), 금액 천단위 구분자 ',' 고려
                                        dtype={'거래고유번호': str,                         # 너무 큰 숫자들이므로 문자열로 셋팅
                                                '금액': 'int64',
                                                '승인번호': str},                           # 영문이나 숫자 '0'도 표시되어야 함으로 문자열 타입
                                        na_values=None
                                        )
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_history.py - Reading Data] <{datetime.datetime.now()}> Pandas excel-reading error ({card_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 2. 데이터 전처리
    # 2-1. '거래일시' 컬럼명 수정
    card_history_df.rename(columns={'거래일시▼': 'date'}, inplace=True)     # '거래일시' 컬럼명 'date'로 수정 => 다른 dataframe과 통일

    # 2-2. '거래고유번호' 컬럼에 있는 내용중에 NaN인 행 제거
    card_history_df.dropna(subset=['거래고유번호'], inplace=True)           # '거래고유번호'가 NaN 값이면 빈칸이거나 잘못된 라인

    # 2-3. '승인번호' 중복 라인제거 == 당일 취소 건 제거
    card_history_df.drop_duplicates(subset=['승인번호'], keep=False, inplace=True)    # keep= 중복 라인들 모두 제거
    
    # 2-4. '승인구분'이 '승인' 라인 추출
    card_history_df = card_history_df[card_history_df['승인구분'] == '승인']

    # 일단, 오늘 날짜 내역도 보존하자.... 대신 처리는 edi_opera.py에서!!
    # # 2-5. target_date 날짜만 추출
    # date_str = work_date.strftime("%Y-%m-%d")  # 어제 날짜 포맷
    # card_history_df = card_history_df[card_history_df['date'].str.contains(date_str, na=False)]
    
    # 2-6. 이전(2일 전) 거래내역에 포함된 내역 제거
    del_reds = red_font.read_red_color()
    if len(del_reds) != 0:              # 이전 내역에 포함된 내역이 존재할 경우
        card_history_df = card_history_df[card_history_df['거래고유번호'].isin(del_reds) == False]      # isin()결과가 False 인 것
    
    # 2-7. 필요한 컬럼만 추출
    card_history_df = card_history_df[['거래고유번호', '승인구분', 'date',
                                        '카드번호', '발급카드사', '매입카드사', '금액', '할부개월', '승인번호']]
    
    # 2-8. date 기준 sorting
    card_history_df.sort_values(by=['date'], inplace=True)
    
    # 2-9. date 정렬이후 포맷 변경: '%Y-%m-%d' (시간부분 제거)
    card_history_df['date'] = card_history_df['date'].map(lambda str_data: str_data.split()[0], na_action='ignore')

    # 3. df디렉토리에 dataframe 저장
    # 3-1. 목적지 df디렉토리 확인하고 없으면 생성
    dfdata_dir = f'./data/{target_date}/dfdata/'        # 데이타를 저장할 df디렉토리 './data/YYYYMMDD/dfdata/'
    
    try:
        if os.path.exists(dfdata_dir) == False:         # 폴더가 없으면 생성
            os.makedirs(dfdata_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:          # error 로그 파일에 추가
            file.write(
                f'kicc_history.py - Making dfdata] <{datetime.datetime.now()}> mkdir error ({dfdata_dir}) ===> {e}\n'
            )
        raise(e)
    
    # 3-2. dataframe 저장
    try:
        df_filename = 'df_kicc_history_' + target_date  # 저장할 dataframe 파일 이름 'df_kicc_history_YYYYMMDD'

        with open(dfdata_dir + df_filename, "wb") as file:
            pickle.dump(card_history_df, file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_history.py - Making Dataframe] <{datetime.datetime.now()}> pickle.dump error ({df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 3-3. dataframe 그대로 excel로 저장 => opera.py와 함께 이용        
    try:
        target_dir = f'./data/{target_date}/'
        xl_filename = 'df_kicc_history_' + target_date + '.xlsx'    # 저장파일 'opera_trial_YYYYMMDD.xlsx
        
        with pnds.ExcelWriter(target_dir + xl_filename, mode='w', engine='openpyxl') as writer:
            card_history_df.to_excel(writer, sheet_name='original', index=False)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_history.py - Making Excel] <{datetime.datetime.now()}> Pandas.ExcelWriter error ({xl_filename}) ===> {e}\n'
            )
        raise(e)


if __name__ == '__main__':
    to_card_history_df()