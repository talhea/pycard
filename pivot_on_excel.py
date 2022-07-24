'''pivot table이 있는 엑셀 템플릿 파일로 '신용거래내역조회_YYMMDD.xlsx'를 생성
'''

import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pnds
import pickle
import datetime

def to_excel(work_date: datetime):
    '''pivot table이 있는 엑셀 템플릿 파일을 이용하여 KICC 카드거래내역을 pivot table로 셋팅한 엑셀 파일로 저장
    어제 날짜의 KICC 거래내역 Dataframe을 읽어서, pivot teble이 설정된 엑셀 템플릿 파일의 원본 데이터 sheet에 주입한다
    이 때, pivot table이 자동으로 refresh되도록 셋팅하든지 혹은 수동으로 파일 오픈 후 새로고침 한다.

    Args:
        work_date (datetime): 어제 날짜
    '''

    # 1. dataframe으로 저장된 KICC 카드거래내역 읽기
    target_date = work_date.strftime("%Y%m%d")                          # 어제 날짜 포맷

    dfdata_dir = f'./data/{target_date}/dfdata/'                        # 읽어들일 dfdata디렉토리 './data/YYYYMMDD/dfdata/'
    df_filename = 'df_kicc_history_' + target_date                      # 파일 이름: df_kicc_history_YYYYMMDD

    #   카드거래내역 Datafarame loading
    try:
        with open(dfdata_dir + df_filename, "rb") as file:
            card_history_df = pickle.load(file)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> pickle file-reading error ({df_filename}) ===> {e}\n'
            )
        raise(e)
    
    # 2. Pivot table 셋팅된 템플릿 엑셀 파일을 준비
    # 2-1. 수식 포함하여 템플릿 엑셀 파일 로딩
    history_template_filename = './신용거래내역조회_tmp.xlsx'               # '신용거래내역' 템플릿 파일
    
    try:
        edi_excel = openpyxl.load_workbook(history_template_filename, data_only=False)
        
        pivot_sheet = edi_excel['pivot']                                    # Pivot table 셋팅된 sheet
        pivot_sheet['A2'].value = '피벗테이블분석-테이터원본변경 필요함'        # 수동 새로고침 시 경고 문구
        pivot_sheet['A2'].style = 'Title'                                   # 경고 문구의 style
        
        # data sheet
        data_ws = edi_excel['data']
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[edi_opera.py - Reading Data] <{datetime.datetime.now()}> openpyxl file-reading error ({history_template_filename}) ===> {e}\n'
            )
        raise(e)

    # 2-2. 1행부터 데이터 행 갯수만큼의(전체) 원본 데이터 삭제
    data_ws.delete_rows(1, data_ws.max_row)

    # 3. Dataframe을 엑셀에 주입
    for row in dataframe_to_rows(card_history_df, index=False, header=True):            # index는 제외
        data_ws.append(row)
    
    # 4. Sheet 이름 변경
    data_ws.title = work_date.strftime('%y%m%b')                                        # shee 이름: 'YYMMDD'
    
    # 5. 엑셀 파일 저장
    try:
        target_dir = f'./data/{target_date}/'                                           # 저장위치 : './data/YYYYMMDD/'
        xl_filename = '신용거래내역조회_' + work_date.strftime('%y%m%b') + '.xlsx'       # 저장파일 '신용거래내역조회_YYMMDD.xlsx'
        print(target_dir + xl_filename)

        edi_excel.save(target_dir + xl_filename)
    except Exception as e:
        with open('./error.log', 'a') as file:
            file.write(
                f'[pivot_on_excel.py - Making Excel] <{datetime.datetime.now()}> openpyxl.save error ({xl_filename}) ===> {e}\n'
            )
        raise(e)


if __name__ == '__main__':
    to_excel(datetime.datetime.now() - datetime.timedelta(1))