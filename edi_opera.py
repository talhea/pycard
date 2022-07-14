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
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> openpyxl file-reading error ({excel_filename}) ===> {e}\n'
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
    
    dfdata_dir = f'./data/{target_date}/dfdata/'                # 읽어들일 dfdata디렉토리 './data//YYYYMMDD/dfdata/'
    opera_df_filename = 'df_opera_trial_' + target_date
    
    try:
        with open(dfdata_dir + opera_df_filename, "rb") as file:
            opera_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> pickle file-reading error ({opera_df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 3. 오페라 dataframe의 trx_code와 EDI 엑셀파일 관련 매칭 정보
    mapping_code_edi = {
        # 'TRX_CODE' : ['Book' EDI엑셀좌표, 'Actual' EDI엑셀좌표, [actual 금액 리스트]]
        '9231': {'book': 'B7', 'actual': 'C7', 'amounts': []},
        '9232': {'book': 'B9', 'actual': 'C9', 'amounts': []},
        '9233': {'book': 'B8', 'actual': 'C8', 'amounts': []},
        '9234': {'book': 'B15', 'actual': 'C15', 'amounts': []},
        '9235': {'book': 'B6', 'actual': 'C6', 'amounts': []},
        '9236': {'book': 'B3', 'actual': 'C3', 'amounts': []},
        '9237': {'book': 'B12', 'actual': 'C12', 'amounts': []},
        '9238': {'book': 'B14', 'actual': 'C14', 'amounts': []},
        '9239': {'book': 'B10', 'actual': 'C10', 'amounts': []},
        '9240': {'book': 'B16', 'actual': 'C16', 'amounts': []},
        '9241': {'book': 'B17', 'actual': 'C17', 'amounts': []},
        '9242': {'book': 'B11', 'actual': 'C11', 'amounts': []},
        '9243': {'book': 'B5', 'actual': 'C5', 'amounts': []},
        '9244': {'book': 'B4', 'actual': 'C4', 'amounts': []},
        '9245': {'book': 'B15', 'actual': 'C15', 'amounts': []}
    }

    # 4. trial balance 내용을 순서대로 돌면서 엑셀에 입력(mapping_code_edi 참조)
    for code in opera_df.index:                                         # trx-code를 이용해서 ws와 opera_df에 접근하는 looping
        ws[mapping_code_edi[code]['book']] = opera_df.loc[code]['TB_AMOUNT'] * -1       # 오페라 dataframe으로부터 tb_amount 정보를 가져와, 
                                                                                        # 맵핑 딕셔너리에서 추출한 좌표를 이용해서 ws(EDI엑셀)에 넣는다
    
    #------------------------------------------------------------------------------------------------------------------
    # KICC 신용카드 승인내역의 금액을 카드별로 분류하여 EDI 엑셀 파일('Actual')에 입력
    #------------------------------------------------------------------------------------------------------------------
    
    # 5. KICC로부터 생성한 신용카드 승인내역 dataframe
    card_df_filename = 'df_kicc_history_' + target_date                    # 파일 이름: df_kicc_history_YYYYMMDD

    try:
        with open(dfdata_dir + card_df_filename, "rb") as file:
            card_history_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> pickle file-reading error ({card_df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 6. EDI 엑셀에 입력하기위한 카드 분류
    # 6-1. '카드분류' 컬럼 생성
    card_history_df['카드분류'] = '분류'        # 'LT', 'KEB', 'JCB', 'VISA', 'MASTER', 'SA', 'SS', 'SH', 'BC', 'KB', 'HD', 'NH', 'CITI'

    # 6-2. 'LT'
    condition = (card_history_df['매입카드사'].str.startswith('롯데', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9244'

    # 6-3. 'KEB'
    condition = (card_history_df['매입카드사'] == '하나구외환')  & (card_history_df['발급카드사'].str.startswith('하나', na=False) | card_history_df['발급카드사'].str.startswith('토스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9243'

    # 6-4. 'JCB'
    condition = (card_history_df['발급카드사'].str.contains('제이씨비', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9235'

    # 6-5. 'VISA'
    condition = (card_history_df['발급카드사'].str.contains('해외비자', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9231'

    # 6-6. 'MASTER'
    condition = (card_history_df['발급카드사'].str.contains('해외마스타', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9233'

    # 6-7. 'SA'
    condition = (card_history_df['발급카드사'].str.contains('해외아멕스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9232'

    # 6-8. 'SS'
    condition = (card_history_df['발급카드사'].str.startswith('삼성', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9239'

    # 6-9. 'SH'
    condition = (card_history_df['매입카드사'].str.startswith('신한', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9242'

    # 6-10. 'BC'
    condition = (card_history_df['매입카드사'].str.startswith('비씨', na=False))      # 'CITI'씨티카드와 중복되지만, 'CITI'는 마지막에 다시 셋팅
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9237'

    # 6-11. 'KB'
    condition = (card_history_df['매입카드사'].str.startswith('국민', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9238'

    # 6-12. 'HD'
    condition = (card_history_df['매입카드사'].str.startswith('현대', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9245'

    # 6-13. 'NH'
    condition = (card_history_df['매입카드사'].str.startswith('NH', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9240'

    # 6-14. 'CITI'
    condition = (card_history_df['발급카드사'].str.startswith('씨티', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9241'
    
    # 7. datframe 카드분류 컬럼에 따라서 해당 금액을 EDI 엑셀에 수식 문자열로 입력 (예, '=+50000+10000...')
    # 7-1. dataframe 코드별 금액을 mapping_code_edi의 'amounts'리스트에 추가
    for index, row in card_history_df.iterrows():                   # dataframe row looping
        mapping_code_edi[row['카드분류']]['amounts'].append(f"+{row['금액']}")  # 금액을 엑셀 수식 문자열로 저장
    
    # 7-2. 'amounts' 금액리스트를 엑셀 서식 문자열로 edi 엑셀 파일에 저장
    for edi in mapping_code_edi.values():                   # EDI 관련 딕셔너리 리스트를 looping
        if len(edi['amounts']) != 0:                                # 해당 날짜의 카드 매출이 있는 것만
            ws[edi['actual']] = '=' + (''.join(edi['amounts']))     # 'amounts'를 엑셀 수식으로 결합해서 EDI엑셀의 'actual'좌표에 입력 (예, '=+50000+10000...')
    
    # 8. sheet 이름 및 날짜 셋팅
    opera_date = yesterday.strftime('%b.%d.%Y')     # 오페라용 어제 날짜 포맷
    ws['B1'] = opera_date                           # EDI 엑셀파일 날짜 셋팅
    
    ws.title = str(yesterday.day)                   # sheet 이름 변경
    
    # 9. 파일 저장
    target_dir = f'./data/{target_date}/'           # 날짜 폴더에 저장 '/data/YYYYMMDD/'
    
    try:
        excel_filename = target_dir + 'EDI-' + target_date + '.xlsx'              # '/data/YYYYMMDD/EDI-YYYYMMDD.xlsx'
        print(excel_filename)
        
        edi_excel.save(excel_filename)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Making Data] <{datetime.datetime.now()}> openpyxl wirting error ({excel_filename}) ===> {e}\n'
            )
        raise(e) 


if __name__ == '__main__':
    merge_edi_opera()