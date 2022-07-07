"""
#   오페라 카드거래내역 청구 자동화 관련 README

1. 각 종 기본 자료 다운르드
    1-1. trial balance, 어제 날짜 기준, from OPERA
        : down_opera.py
    1-2. 신용거래내역, 어제 날짜부터 오늘까지 기준, from KICC
        : down_kicc.py
    1-3. 입금 현황 (카드), 오늘 날짜 기준, from KICC
        : down_kicc.py
    1-4. 은행 입금 내역, 오늘 날짜 기준, from 기업은행
        : down_ibk.py
    1-5. 은행 입금 내역, 오늘 날짜 기준, from NH 민국은행
        : down_nh.py
    
    downdata 폴더에 저장
    
2. 다운로드 받은 자료들 DataFrame으로 가공 저장
    2-1. dt_opera_trial_2022-07-01, 어제 날짜 기준, from opera_xml.py
    2-2. dt_kicc_history_2022-07-01, 어제 날짜 기준, from kicc_history_excel.py
    2-3. dt_kicc_recipes_2022-07-01, 오늘 날짜 기준, from kicc_receips_excel.py
    2-4. dt_bank_2022-07-01, 오늘 날짜 기준, from bank_data.py
    
"""

# 2022.07.03 home, pm 04:29
# bank_data.py 에서 'details' 관련 부부에서 'NH' 넣는 부분을 반드시 수정할것
# 수정 완료

# 2022.07.04, PM 05:38
# down_kicc.py 에서, 로그인 누르면 프로그램 설치 화면이 나올떄가 있음..
# 따라서, 차라리 무조건 처음 화면에서 로그인해서 들어가는 방법은???
# 문제는.... 프로그램 설치시 확인 누르지 못할수도...
# 수정: 일단 보안프로그램 다시 설치해서 해본다

# 2022.07.07, PM 04:33
# 기본 다운로드 및 dataframe 저장 완료