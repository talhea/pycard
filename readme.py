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
