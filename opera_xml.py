# 오페라 trial report xml data

import pandas as pnds
import xml.etree.ElementTree as et
import pickle
import datetime
#from IPython.display import display

def to_opera_df():
    # 0. dataframe과 엑셀로 각 각 저장
    target_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    down_base_dir = './downdata/' + target_date

    # 1. xml 파일로된 오페라 trial balance를 ElementTree 객체로 읽어들임
    xml_filename = down_base_dir + '/trial_balance' + target_date + '.xml'
    xtree = et.parse(xml_filename)
    xroot = xtree.getroot()

    # 2. trial balance 내용인 node(G_TRX_CODE tag : ) 추출
    g_trx_codes = xtree.findall('LIST_G_TRX_TYPE/G_TRX_TYPE/LIST_G_TRX_CODE/G_TRX_CODE')

    # 3. trial balance용 dataframe 생성
    # 3-1. dataframe에 사용될 column들
    df_cols = ('TRX_CODE', 'DESCRIPTION', 'TB_AMOUNT', 'TRX_DATE', 'NET_AMOUNT', 'GUEST_LED_DEBIT',
            'GUEST_LED_CREDIT', 'AR_LED_DEBIT', 'AR_LED_CREDIT', 'DEP_LED_DEBIT', 'DEP_LED_CREDIT', 'PACKAGE_LED_DEBIT',
            'PACKAGE_LED_CREDIT', 'NON_REVENUE_AMT', 'TODAYS_ACCRUALS', 'DEPOSIT_AT_CHECKIN')

    # 3-2. dataframe에 사용될 내용 추출
    rows = []                       # dataframe을 위한 index buffer
    for node in g_trx_codes:        # each index를 looping
        cols = {}                   # column용 딕셔너리 buffer
        for item in node:           # index 내의 each column을 looping
            key = item.tag.strip()      # column용 딕셔너리 key
            if key in df_cols:          # key가 dataframe용 column 일 경우만
                value = item.text       # column용 딕셔너리 value
                # if value is None:         # value 전처리, None이면 0...
                #     value = 0             # None: 해당없음 의미.. 따라서 값 유지

                cols.update({key: value})   # column용 딕셔너리 추가

        rows.append(cols)           # 완성된 index 한 줄을 리스트에 추가


    # 3-3. 추출된 rows 리스트와 df_cols 리스트를 이용해서 DataFrame 생성하고 원본 복사(추후에 엑셀 sheet에 저장)
    opera_df = pnds.DataFrame(rows, columns=df_cols)
    origin_df = opera_df.copy()

    # 4. 데이터 전처리 : 자료형 변환
    opera_df['TRX_DATE'] = opera_df['TRX_DATE'].astype('datetime64[ns]', errors='ignore')   # date 포맷 변경을 위해서 datetimee 타입으로 변환
    opera_df['TRX_DATE'] = opera_df['TRX_DATE'].dt.strftime('%Y-%m-%d')                     # 연산을 통해 포맷 변경 => 반환 타입은 일반 객체로 변경됨

    for col in df_cols:
        if col not in ['DESCRIPTION', 'TRX_DATE', 'TRX_CODE']:              # description과 date, trx_code 컬럼만 제외하고 모두 int형으로 변환  
            opera_df[col] = opera_df[col].astype('int', errors='ignore')    # 에러 발생 값들은 무시, 따라서 아래와 같은...
                                                                            # 일부 컬럼들 int로 변환 실패는, None 타입에의한 에러 발생으로 추측

    # 5. 실제 필요한 dataframe 추출
    # 5-1. 오페라 신용카드 transaction code 추출 : 실행파일 opera_card_trx_code.py(ipynb)에서 생성
    card_df = pnds.read_csv('./trx_card_codes.csv', sep='\t')               # {'TRX_CODE': 'Description'} 형식
    card_df['TRX_CODE'] = card_df['TRX_CODE'].astype('str', errors='ignore')# trx_code를 str 타입으로 변경
    card_df.set_index(keys=['TRX_CODE'], inplace=True)                      # transaction code를 index로 변경

    # 5-2. 추출된 transaction code에 해당되는 것들만 추출
    opera_df = opera_df.loc[opera_df['TRX_CODE'].isin(card_df.index)][['TRX_CODE','DESCRIPTION', 'TB_AMOUNT', 'NON_REVENUE_AMT', 'TRX_DATE']]
    opera_df.set_index('TRX_CODE', drop=True, inplace=True)

    # 6. dataframe과 엑셀로 각 각 저장
    dt_base_dir = './dtdata/' + target_date

    # 6-1. dataframe 저장
    df_filename = '/dt_opera_trial_' + target_date
    with open(dt_base_dir + df_filename, "wb") as file:
        pickle.dump(opera_df, file)

    # 6-2. excel 저장
    xl_filename = '/opera_trial_' + target_date + '.xlsx'
    with pnds.ExcelWriter(dt_base_dir + xl_filename) as writer:
        origin_df.to_excel(writer, sheet_name='original', index=False)

if __name__ == '__main__':
    to_opera_df()