'''EDI 템플릿 엑셀 파일에 오페라 balance trial dataframe을 입력한다

'''
import pandas as pnds
import openpyxl
import datetime
import pickle, os

def merge_edi_opera():
    """EDI 템플릿 엑셀 파일에 오페라 trial balance와 KICC 승인내역의 dataframe을 입력한다
    """    
    
    # 1. EDI 템플릿 엑셀 파일을 읽어서 DataFrame 생성
    excel_filename = './EDI-xx월.xlsx'                                      # EDI 엑셀 템플릿 파일
    
    try:
        edi_excel = openpyxl.load_workbook(excel_filename, data_only=False) # 수식파일 포함하여 엑셀파일 읽어 들임
        ws = edi_excel.active                                               # 활성화된 sheet
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> openpyxl file-reading error {excel_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> openpyxl file-reading error {excel_filename} : {e}\n'
            )
        raise(e)

    #------------------------------------------------------------------------------------------------------------------
    # 오페라 trial balance를 EDI 엑셀 파일('BOOK')에 입력
    #------------------------------------------------------------------------------------------------------------------
    
    # 2. 오페라 어제 마감된 trial report(df_trial_balance_날짜) dataframe 로딩
    yesterday = datetime.datetime.now() - datetime.timedelta(1) # 어제 날짜 추출
    target_date = yesterday.strftime("%Y%m%d")                  # 어제 날짜 포맷
    # 임시
    # target_date = '20220710'
    
    dfdata_dir = './dfdata/' + target_date + '/'
    opera_df_filename = 'df_opera_trial_' + target_date
    
    try:
        with open(dfdata_dir + opera_df_filename, "rb") as file:
            opera_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle file-reading error {dfdata_dir + opera_df_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle file-reading error {dfdata_dir + opera_df_filename} : {e}\n'
            )
        raise(e)
    
    # 3. 오페라 dataframe의 trx_code와 EDI 엑셀 book 컬럼 좌표 매칭 정보
    mapping_code_book = {
        # TRX_CODE : EDI BOOL컬럼 엑셀좌표
        '9231': 'B7',
        '9232': 'B9',
        '9233': 'B8',
        '9234': 'B15',
        '9235': 'B6',
        '9236': 'B3',
        '9237': 'B12',
        '9238': 'B14',
        '9239': 'B10',
        '9240': 'B16',
        '9241': 'B17',
        '9242': 'B11',
        '9243': 'B5',
        '9244': 'B4',
        '9245': 'B15'
    }

    # 4. trial balance 내용을 순서대로 돌면서 엑셀에 입력(mapping_code_book 참조)
    for code in opera_df.index:     # trx-code를 이용해서 ws와 opera_df에 접근하는 looping
        ws[mapping_code_book[code]] = opera_df.loc[code]['TB_AMOUNT'] * -1       # 오페라 dataframe으로부터 tb_amount 정보를 가져와, 
                                                                            # 맵핑 딕셔너리에서 추출한 좌표를 이용해서 ws(EDI엑셀)에 넣는다
    
    #------------------------------------------------------------------------------------------------------------------
    # KICC 신용카드 승인내역의 금액을 카드별로 분류하여 EDI 엑셀 파일('Actual')에 입력
    #------------------------------------------------------------------------------------------------------------------
    
    # 5. KICC로부터 생성한 신용카드 승인내역 dataframe('df_kicc_history_YYYYMMDDxlsx')을 로딩
    card_df_filename = 'df_kicc_history_' + target_date                    # 파일 이름: df_kicc_history_YYYYMMDD

    try:
        with open(dfdata_dir + card_df_filename, "rb") as file:
            card_history_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle file-reading error {dfdata_dir + card_df_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> pickle file-reading error {dfdata_dir + card_df_filename} : {e}'
            )
        raise(e)
    
    # 6. EDI 엑셀에 입력하기위한 카드 분류
    # '카드분류' 컬럼 생성
    card_history_df['카드분류'] = '분류'        # 'LT', 'KEB', 'JCB', 'VISA', 'MASTER', 'SA', 'SS', 'SH', 'BC', 'KB', 'HD', 'NH', 'CITI'

    # 'LT'
    condition = (card_history_df['매입카드사'].str.startswith('롯데', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9244'

    # 'KEB'
    condition = (card_history_df['매입카드사'] == '하나구외환')  & (card_history_df['발급카드사'].str.startswith('하나', na=False) | card_history_df['발급카드사'].str.startswith('토스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9243'

    # 'JCB'
    condition = (card_history_df['발급카드사'].str.contains('제이씨비', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9235'

    # 'VISA'
    condition = (card_history_df['발급카드사'].str.contains('해외비자', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9231'

    # 'MASTER'
    condition = (card_history_df['발급카드사'].str.contains('해외마스타', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9233'

    # 'SA'
    condition = (card_history_df['발급카드사'].str.contains('해외아멕스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9232'

    # 'SS'
    condition = (card_history_df['발급카드사'].str.startswith('삼성', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9239'

    # 'SH'
    condition = (card_history_df['매입카드사'].str.startswith('신한', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9242'

    # 'BC'
    condition = (card_history_df['매입카드사'].str.startswith('비씨', na=False))      # 'CITI'씨티카드와 중복되지만, 'CITI'는 마지막에 다시 셋팅
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9237'

    # 'KB'
    condition = (card_history_df['매입카드사'].str.startswith('국민', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9238'

    # 'HD'
    condition = (card_history_df['매입카드사'].str.startswith('현대', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9245'

    # 'NH'
    condition = (card_history_df['매입카드사'].str.startswith('NH', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9240'

    # 'CITI'
    condition = (card_history_df['발급카드사'].str.startswith('씨티', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9241'
    
    # 7. 카드분류에 따라 금액리스트에 금액 입력
    # '카드 트랙잭션 코드별 금액' 딕셔너리, 금액은 리스트 타입 value에 추가되며 엑셀 입력 서식이 적용됨(예, '=+3000+2200+..')
    actual_cards_dict = {
        # {'transaction code': [금액,],}
        '9231': ['='],
        '9232': ['='],
        '9233': ['='],
        '9234': ['='],
        '9235': ['='],
        '9237': ['='],
        '9238': ['='],
        '9239': ['='],
        '9240': ['='],
        '9241': ['='],
        '9242': ['='],
        '9243': ['='],
        '9244': ['='],
        '9245': ['=']
    }
    
    # KICC 카드내역의 금액을 EDI 엑셀 Actual 컬럼 좌표와 연동하기위한 맵핑 정보
    mapping_code_actual = {
        # TRX_CODE : EDI Actual컬럼 엑셀좌표
        '9231': 'C7',
        '9232': 'C9',
        '9233': 'C8',
        '9234': 'C15',
        '9235': 'C6',
        '9236': 'C3',
        '9237': 'C12',
        '9238': 'C14',
        '9239': 'C10',
        '9240': 'C16',
        '9241': 'C17',
        '9242': 'C11',
        '9243': 'C5',
        '9244': 'C4',
        '9245': 'C15'
    }
    # 코드별로 해당되는 각 리스트에 금액을 추가
    for index, row in card_history_df.iterrows():                       # dataframe row looping
        actual_cards_dict[row['카드분류']].append(f"+{row['금액']}")     # 해당 row의 금액을 str 타입으로 바꾸어서 코드별 리스트에 저장
    
    # 코드별로 금액리스트를 엑셀 문자열 서식으로 edi 엑셀 파일에 저장(mapping_code_actual 참조)
    for code, amounts in actual_cards_dict.items():             # trx-code를 이용해서 ws에 접근하는 looping
        ws[mapping_code_actual[code]] = ''.join(amounts)        # 금액리스트(amounts)를 엑셀 입력 서식('=+금액+금액...')으로 join
    
    # 8. sheet 이름 및 날짜 셋팅
    opera_date = yesterday.strftime('%b.%d.%Y')     # 오페라 포맷용 어제 날짜
    ws['B1'] = opera_date                           # EDI파일 날짜 셋팅
    
    ws.title = str(yesterday.day)                   # sheet 이름 변경
    
    # 9. 파일 저장
    try:
        excel_filename = dfdata_dir + 'EDI-' + target_date + '.xlsx'              # 파일 저장 시에 필요한 당 월의 엑셀파일 이름
        print(excel_filename)
        edi_excel.save(excel_filename)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[{__name__}.py] <{datetime.datetime.now()}> openpyxl wirting error {excel_filename} : {e}\n'
            )
            print(
                f'[{__name__}.py] <{datetime.datetime.now()}> openpyxl wirting error {excel_filename} : {e}'
            )
        raise(e) 

if __name__ == '__main__':
    merge_edi_opera()