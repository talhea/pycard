'''EDI 템플릿 엑셀 파일에 오페라 trial balance와 KICC 승인내역의 dataframe을 입력한다

'''
import pandas as pnds
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import pickle, os
import duplicated_history

def merge_edi_opera(work_date):
    """EDI 템플릿 엑셀 파일에 오페라 trial balance와 KICC 승인내역의 dataframe을 입력한다
    trial balance dataframe을 EDI 파일의 card 분류에 맞게 'book' 컬럼에 입력하고,
    kicc 신용거래내역 dataframe도 EDI 파일의 card 분류에 맞게 분류해서 'actual'컬럼에 입력한다.
    
    Args:
        work_date (datetime): 어제 날짜
    """    
    
    # 1. EDI 템플릿 엑셀 파일을 로딩
    excel_filename = './EDI-xx월.xlsx'                                      # EDI 엑셀 템플릿 파일
    
    # 수식 포함하여 엑셀 파일 로딩
    try:
        edi_excel = openpyxl.load_workbook(excel_filename, data_only=False)
        edi_ws = edi_excel.active                                           # 활성화된 sheet
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> openpyxl file-reading error ({excel_filename}) ===> {e}\n'
            )
        raise(e)

    # 2. card code와 EDI 엑셀파일의 관련 매칭 정보
    #    이 매칭 정보의 key(card code)를 이용해서 trial balance와 kicc 신용거래내용을 EDI 파일 card 분류대로 각각 저장한다
    mapping_code_edi = {
        # 'TRX_CODE' : ['Book' EDI엑셀좌표, 'Actual' EDI엑셀좌표, [actual 금액 리스트-kicc거래내역]]
        '9231': {'book': 'B7', 'actual': 'C7', 'amounts': []},
        '9232': {'book': 'B9', 'actual': 'C9', 'amounts': []},
        '9233': {'book': 'B8', 'actual': 'C8', 'amounts': []},
        '9235': {'book': 'B6', 'actual': 'C6', 'amounts': []},
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

    #------------------------------------------------------------------------------------------------------------------
    # 오페라 trial balance를 EDI 엑셀 파일('BOOK') 'book' 컬럼에 입력
    #------------------------------------------------------------------------------------------------------------------
    
    # 3. 오페라 어제 마감된 trial report(df_trial_balance_날짜) dataframe 로딩
    target_date = work_date.strftime('%Y%m%d')                              # 어제 날짜 포맷
    
    dfdata_dir = f'./data/{target_date}/dfdata/'                            # 읽어들일 dfdata디렉토리 './data/YYYYMMDD/dfdata/'
    opera_xl_filename = 'df_opera_trial_' + target_date
    
    #   Datafarame loading
    try:
        with open(dfdata_dir + opera_xl_filename, "rb") as file:
            opera_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> pickle file-reading error ({opera_xl_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 4. trial balance 내의 각 카드별 금액을 해당되는 EDI 엑셀에 입력(mapping_code_edi 참조)
    for code in opera_df.index:                                         # trx-code(index)를 이용해서 edi_ws와 opera_df에 접근하는 looping
        edi_ws[mapping_code_edi[code]['book']] = opera_df.loc[code, 'TB_AMOUNT'] * -1       # 오페라 dataframe으로부터 금액(tb_amount) 정보를 가져와, 
                                                                                            # 맵핑 딕셔너리에서 추출한 좌표를 이용해서 edi_ws(EDI엑셀)에 넣는다
    
    #------------------------------------------------------------------------------------------------------------------
    # KICC 신용카드 승인내역의 각 금액을 카드별로 분류하여 EDI 엑셀 'Actual'에 입력
    #------------------------------------------------------------------------------------------------------------------
    
    # 5. KICC로부터 생성한 신용카드 승인내역 dataframe
    card_xl_filename = 'df_kicc_history_' + target_date                     # 파일 이름: df_kicc_history_YYYYMMDD
    
    #   KICC 카드 승인내역 Dataframe loading
    try:
        with open(dfdata_dir + card_xl_filename, "rb") as file:
            card_history_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> pickle file-reading error ({card_xl_filename}) ===> {e}\n'
            )
        raise(e)
    
    #   이전(2일 전) 거래내역 중, 이미 등록된 어제 날짜 내역 제거
    del_dup_serials = duplicated_history.get_dup_serials(work_date)
    if len(del_dup_serials) != 0:                                           # 이전 내역에 포함된 내역이 있을 경우
        card_history_df = card_history_df[card_history_df['거래고유번호'].isin(del_dup_serials) == False]           # isin()결과가 False 인 것만 추출
    
    #   내일 날짜는 제외 - 해당 날짜의 거래내역만 추출
    tomorrow = (work_date + datetime.timedelta(1)).strftime('%d')                                                   # 내일 날짜의 day: 'DD' (두자리 문자열)
    card_history_df = card_history_df[card_history_df['거래고유번호'].str.startswith(tomorrow, na=False) == False]  # 거래고유번호가 내일날자(DD)로 시작되지 않는것들만
    
    #   보정된 Dataframe을 './data/YYYYMMDD/' 폴더에 저장해서 pivot_on_excel.py에서 이용
    try:
        target_dir = f'./data/{target_date}/'               # 수정된 파일 저장할 날짜 폴더 '/data/YYYYMMDD/'
        df_filename = 'df_kicc_history_' + target_date      # 저장할 dataframe 파일 이름 'df_kicc_history_YYYYMMDD'
        print(target_dir + df_filename)
        
        with open(target_dir + df_filename, "wb") as file:
            pickle.dump(card_history_df, file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[kicc_history.py - Making Dataframe] <{datetime.datetime.now()}> pickle.dump error ({df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 6. 각 카드 내역들을 EDI 엑셀 파일의 카드 분류에따라 재분류
    # 6-1. '카드분류'(card code) 컬럼과 그 영문 단축명인 'Description' 컬럼 생성
    card_history_df['카드분류'] = '분류'
    card_history_df['Description'] = '영문명'     # 'LT', 'KEB', 'JCB', 'VISA', 'MASTER', 'SA', 'SS', 'SH', 'BC', 'KB', 'HD', 'NH', 'CITI'
    
    #   'LT'
    condition = (card_history_df['매입카드사'].str.startswith('롯데', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9244'
    card_history_df.loc[collected_cards, 'Description'] = 'LT'

    #   'KEB'
    condition = (card_history_df['매입카드사'] == '하나구외환')  & (card_history_df['발급카드사'].str.startswith('하나', na=False) | card_history_df['발급카드사'].str.startswith('토스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9243'
    card_history_df.loc[collected_cards, 'Description'] = 'KEB'

    #   'JCB'
    condition = (card_history_df['발급카드사'].str.contains('제이씨비', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9235'
    card_history_df.loc[collected_cards, 'Description'] = 'JCB'

    #   'VISA'
    condition = (card_history_df['발급카드사'].str.contains('해외비자', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9231'
    card_history_df.loc[collected_cards, 'Description'] = 'VISA'

    #   'MASTER'
    condition = (card_history_df['발급카드사'].str.contains('해외마스타', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9233'
    card_history_df.loc[collected_cards, 'Description'] = 'MASTER'

    #   'SA'
    condition = (card_history_df['발급카드사'].str.contains('해외아멕스', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9232'
    card_history_df.loc[collected_cards, 'Description'] = 'SA'

    #   'SS'
    condition = (card_history_df['발급카드사'].str.startswith('삼성', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9239'
    card_history_df.loc[collected_cards, 'Description'] = 'SS'

    #   'SH'
    condition = (card_history_df['매입카드사'].str.startswith('신한', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9242'
    card_history_df.loc[collected_cards, 'Description'] = 'SH'

    #   'BC'
    condition = (card_history_df['매입카드사'].str.startswith('비씨', na=False))      # 'CITI'씨티카드와 중복되지만, 'CITI'는 마지막에 다시 셋팅
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9237'
    card_history_df.loc[collected_cards, 'Description'] = 'BC'

    #   'KB'
    condition = (card_history_df['매입카드사'].str.startswith('국민', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9238'
    card_history_df.loc[collected_cards, 'Description'] = 'KB'

    #   'HD'
    condition = (card_history_df['매입카드사'].str.startswith('현대', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9245'
    card_history_df.loc[collected_cards, 'Description'] = 'HD'

    #   'NH'
    condition = (card_history_df['매입카드사'].str.startswith('NH', na=False))
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9240'
    card_history_df.loc[collected_cards, 'Description'] = 'NH'

    #   'CITI'
    condition = (card_history_df['발급카드사'].str.startswith('씨티', na=False)) & (card_history_df['발급카드사'].str.contains('비씨', na=False) == False)
    collected_cards = card_history_df[condition].index.tolist()
    card_history_df.loc[collected_cards, '카드분류'] = '9241'
    card_history_df.loc[collected_cards, 'Description'] = 'CITI'
    
    # 7. datframe 카드분류 컬럼에 따라 해당 금액을 EDI 엑셀에 수식 문자열로 입력 (예, '=+50000+10000...')
    # 7-1. dataframe 각 금액을 code에따라 mapping_code_edi의 'amounts'리스트에 추가
    for index, row in card_history_df.iterrows():                   # dataframe row looping
        mapping_code_edi[row['카드분류']]['amounts'].append(f"+{row['금액']}")          # 금액을 엑셀 수식 문자열로 저장(예, '+50000+10000...')
    
    # 7-2. 위에서 입력된 'amounts' 금액리스트를 엑셀 서식 문자열로 edi 엑셀 파일에 저장
    for edi in mapping_code_edi.values():                   # EDI 매핑 정보의 값들을 looping
        if len(edi['amounts']) != 0:                                            # 해당 날짜의 카드 매출이 있는 것만
            edi_ws[edi['actual']] = '=' + (''.join(edi['amounts']))             # 리스트를 엑셀 수식으로 타입으로 결합해서 EDI엑셀의 'actual'좌표에 입력 (예, '=+50000+10000...')
    
    # 8. 날짜 및 sheet명 셋팅
    edi_ws['B1'] = work_date.strftime('%b.%d.%Y')       # EDI 엑셀파일 날짜 셋팅
    edi_ws.title = str(work_date.day)                   # sheet 이름 변경
    
    # 9. EDI 엑셀 파일 저장
    
    
    try:
        excel_filename = 'EDI-' + target_date + '.xlsx' # 'EDI-YYYYMMDD.xlsx'
        print(target_dir + excel_filename)
        
        edi_excel.save(target_dir + excel_filename)     # '/data/YYYYMMDD/EDI-YYYYMMDD.xlsx'
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Making Data] <{datetime.datetime.now()}> openpyxl wirting error ({excel_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 10. grouping된 dataframe을 기존 df_kicc_history 엑셀파일에 추가 sheet로 저장
    # 10-1. 카드별 건수와 합계 금액 dataframe 생성: pivot(또는 groupby)를 이용하여 grouping
    grouped_df = pnds.pivot_table(card_history_df, index=['매입카드사', 'Description'], values=['금액'], aggfunc=['count', 'sum'], margins = True)
    
    # 10-2. 멀티 컬럼이므로 읽기 쉽도록 컬럼명 변경
    grouped_df.columns = ['갯수:건수', '합계:금액']

    # 10-3. dataframe의 행 순서를 EDI 엑셀파일 카드 순서대로 변경
    #       EDI 엑셀파일의 카드 순서 리스트를 이용하는 정렬용 key 함수
    def card_sort(series) -> pnds.Series:
        """ordered_card 리스트를 이용하는 정렬용 key함수
        인수로 전달된 카드 영문 단축명에 따라서, opdered_card의 value(정렬용 값)을 series로 반환

        Args:
            series (str): 카드 영문 단축명 series('Description' 컬럼)

        Returns:
            series (int): 인수로 들어온 series와 매칭되는 정열용 값을 가진 series
        """

        # EDI 엑셀 파일내의 카드 순서 표시
        ordered_card = {'LT': 1, 'KEB': 2, 'JCB': 3, 'VISA': 4, 'MASTER': 5, 'SA': 6,
                        'SS': 7, 'SH': 8, 'BC': 9, 'KB': 10, 'HD': 11, 'NH': 12, 'CITI': 13}
        
        return series.apply(lambda col: ordered_card.get(col, 100))
    
    #       정렬 작업 전후로, 멀티 인텍스를 임시로 해제 및 복구
    grouped_df.reset_index(inplace=True)                                    # 먼저, dataframe의 index를 리셋해서 멀티 인덱스 해제
    grouped_df.sort_values(by='Description', key=card_sort, inplace=True)   # key함수를 이용해서 정렬
    grouped_df.set_index(['매입카드사', 'Description'], inplace=True)        # 다시 원래 형태의 grouping 형태대로 index 셋팅
    
    # 10-4. 기존의 KICC 카드거래내역 엑셀 파일에 추가 모드로 저장
    try:
        df_filename = 'df_kicc_history_' + target_date + '.xlsx'            # 저장파일 'df_kicc_history_YYYYMMDD.xlsx
        print(target_dir + df_filename)

        with pnds.ExcelWriter(target_dir + df_filename, mode='a', engine='openpyxl') as writer:     # 기존 파일 추가 mode 셋팅
            grouped_df.to_excel(writer, sheet_name='groupped_calc', index=True)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Making Excel] <{datetime.datetime.now()}> Pandas.ExcelWriter error ({df_filename}) ===> {e}\n'
            )
        raise(e)


if __name__ == '__main__':
    merge_edi_opera(datetime.datetime.now() - datetime.timedelta(1))