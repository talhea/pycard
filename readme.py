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

# 2022.07.07, PM 05:00
# 1. EDI 파일과 작업할 파일? => dt_opera_trial_날짜

# 2022.07.18. PM 04:50
# dataframe에서 날짜를 datetime형으로 ???? 생각 필요...

# 2022.07.23. AM 09:23
# 엑셀의 pivot을 그대로 이용하기위해서... 
# 엑셀 파일 하나에, 첫 번째 시트(pivot)에는 pivot table을 셋팅하고, 두번쨰 시트에(data)는 원본 데이터를 셋팅.
# 이후에 파이썬에서 이 파일을 로드하고, 두번째 시트에있는 원본 데이터를 변경해서 저장
# 물론.. 옵션으로 파일 로딩시에 refresh 되도록 셋팅...
# => pivot_on_escel.py

# 2022.07.25. AM 09:18
# 다운로드 데이터가 없을 경우에 대한 대책을 만련할 것!!!