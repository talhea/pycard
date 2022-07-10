"""
다운로드 받은 KICC 신용카드 거래내역을 dataframe으로 저장한다
"""  

import pandas as pnds
import pickle, os
import datetime

def to_card_history_df():
    """다운로드 받은 KICC 신용카드 거래내역을 dataframe으로 저장한다
    xlsx 파일을 read_excel함수로 읽어들여 dataframe으로 만들고, 중복된 승인번호(당일 취소) 제거하고 날짜 포맷 셋팅 후 dataframe 으로 저장
    """    

    # 1. 엑셀 파일을 읽어서 DataFrame 생성
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    down_base_dir = './downdata/' + target_date + '/'                                   # 읽어들일 down디렉토리 './downdata/YYYYMMDD'
    card_filename = down_base_dir + '신용거래내역조회.xlsx'                             # KICC 카드 거래내역 엑셀파일

    card_history_df = pnds.read_excel(
                                    card_filename, header=0, thousands = ',',           # 첫번쨰 라인은 컬럼명(header), 금액 천단위 구분자 ',' 고려
                                    dtype={'거래고유번호': str,                         # 너무 큰 숫자들이므로 문자열로 셋팅
                                            '금액': 'int64',
                                            '승인번호': str}                            # 영문이나 숫자 '0'도 표시되어야 함으로 문자열 타입
                                    )
    
    # 2. 데이터 전처리
    # 2-1. '거래일시' 컬럼명 수정
    card_history_df.rename(columns={'거래일시▼': 'date'}, inplace=True)     # '거래일시' 컬럼명 'date'로 수정 => 다른 dataframe과 통일

    # 2-2. '거래고유번호' 컬럼에 있는 내용중에 NaN인 행 제거
    card_history_df.dropna(subset=['거래고유번호'], inplace=True)           # '거래고유번호'가 NaN 값이면 빈칸이거나 잘못된 라인

    # 2-3. '승인번호' 중복 라인제거 == 당일 취소 건 제거
    card_history_df.drop_duplicates(subset=['승인번호'], keep=False, inplace=True)    # keep= 중복 라인들 모두 제거

    # 2-4. 필요한 컬럼만 추출
    card_history_df = card_history_df[['거래고유번호', '승인구분', 'date', '승인번호',
                                        '카드번호', '발급카드사', '매입카드사', '금액']]

    # 2-5. date 기준 sorting
    card_history_df.sort_values(by=['date'], inplace=True)
    
    # 2-6. date 정렬이후 포맷 변경: '%Y-%m-%d'
    card_history_df['date'] = card_history_df['date'].map(lambda str_data: str_data.split()[0], na_action='ignore')

    # 3. dt디렉토리에 dataframe 저장
    dt_base_dir = './dtdata/' + target_date + '/'       # 데이타를 저장할 dt디렉토리 './dtdata/YYYYMMDD/'
    
    # 3-1. 목적지 dt디렉토리 확인하고 없으면 생성
    try:
        if os.path.exists(dt_base_dir) == False:        # 폴더가 없으면 생성
            os.makedirs(dt_base_dir)
    except Exception as e:
        with open('./error.log', 'a') as file:          # error 로그 파일에 추가
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {dt_base_dir} : {e}'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> mkdir error {dt_base_dir} : {e}'
            )
    
    # 3-2. dataframe 저장
    try:
        df_filename = 'dt_kicc_history_' + target_date      # 저장할 dataframe 파일 이름 'dt_kicc_history_YYYYMMDD'

        with open(dt_base_dir + df_filename, "wb") as file:
            pickle.dump(card_history_df, file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle.dump error {df_filename} : {e}'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle.dump error {df_filename} : {e}'
            )


if __name__ == '__main__':
    to_card_history_df()