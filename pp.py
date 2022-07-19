'''
프로그램별로 실행한 결과를 바탕으로 실패한 프로그램 재실행
문제점) 한번 실패한 과정은 계속 실패만 하므로, 더 이상 진행할 필요가 없음
'''
from operator import is_
from sqlite3 import TimeFromTicks
import schedule, time, datetime
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import bank_data
import preparation
# from pytimekr import pytimekr

# 휴일 날짜 획득
# holidays = pytimekr.holidays(year=datetime.datetime.now().year)
# holidays.append(datetime.datetime.now().date())
# for day in holidays:
#     print(type(day), day)

# # 토(5), 일(6) 주말 확인
# is_weekend = datetime.datetime.now().weekday()


# 1. 휴일이나 주말일때는 process를 진행하지 않고, 그 날짜를 리스트에 저장한다.
#   이 값들은 호출되는 함수 내에서 strftime함수를 호출하거나, delta값 계산에 이용된다
# delayed_date = []               # datetime 리스트

# 2. 아침에 시작할때, 위의 리스트를 체크해서 먼저 실행하고, 당일을 실행한다
#   실행시에는, 인자로 datatime을 넘겨준다

# 금일 확인
today = datetime.datetime.now()

# if today.weekday() in [5, 6] or today.date() in holidays:       # 토, 일요일, 공유일
#     delayed_date.append(today.date())                           # 리스트에 추가
#     # 종료 또는 블럭 벗어날 것
#     exit()
# elif len(holidays) != 0:                # 이전의 작업일이 있다면 실행
#     for holiday in holidays:            # 등록된 모든 날짜대로 looping
#         wanted_date = holiday - datetime.timedelta(1)       # 해당 날의 작업 날짜(하루 전)
#         # 해당 날짜의 작업 실시
#         # schedule.every(3).minutes.do(job) 3분 마다 작업 실시

# 출근하는 오늘 날짜 처리
yesterday = today - datetime.timedelta(1)

# preparation.ready(yesterday)
# schedule.every().day.at("07:00").do(preparation.ready, yesterday)
# down_opera.down(yesterday)
# # schedule.every().day.at("07:05").do(down_opera.down, yesterday)
# down_kicc.down(today)
# schedule.every().day.at("07:10").do(down_kicc.down, today)
# opera_xml.to_opera_df(yesterday)
# # schedule.every().day.at("07:20").do(opera_xml.to_opera_df, yesterday)
# kicc_history.to_card_history_df(yesterday)
# # schedule.every().day.at("07:25").do(kicc_history.to_card_history_df, yesterday)
# kicc_receipts.to_receipts_history_df(today)
# # schedule.every().day.at("07:30").do(kicc_receipts.to_receipts_history_df, today)

# edi_opera.merge_edi_opera(yesterday)
# down_nh.down(today)
# # schedule.every().day.at("07:15").do(down_nh.down, today)
# # schedule.every().day.at("07:35").do(edi_opera.merge_edi_opera, yesterday)
# down_ibk.down(today)
# # schedule.every().day.at("13:01").do(down_ibk.down, today)
# bank_data.to_bank_df(today)
# # schedule.every().day.at("13:01").do(bank_data.to_bank_df, today)




# # 실행 결과 저장 버퍼
# result_for_processing = {
#     # 'procee 이름(모듈 이름)' : {'exex': 실행할 함수이름, 'result': 실행결과 bool}
#     'preparation': {'exec': preparation.ready, 'result': False},
#     'down_opera': {'exec': down_opera.down, 'result': False},
#     'down_kicc': {'exec': down_kicc.down, 'result': False},
#     'down_nh': {'exec': down_nh.down, 'result': False},
#     'opera_xml': {'exec': opera_xml.to_opera_df, 'result': False},
#     'kicc_history': {'exec': kicc_history.to_card_history_df, 'result': False},
#     'kicc_receipts': {'exec': kicc_receipts.to_receipts_history_df, 'result': False},
#     'edi_opera': {'exec': edi_opera.merge_edi_opera, 'result': False},
#     'down_ibk': {'exec': down_ibk.down, 'result': False},
#     'bank_data': {'exec': bank_data.to_bank_df, 'result': False},
# }

# # 시간 타이머
# start_time = datetime.datetime.now()            # 시작 시간
# end_time, diff_time = 0, 0                      # 끝 시간, 시간 차이 

# # 모든 처리결과('result')가 True일때까지  process 반복 실행
# def processing_pkg():
#     while True:
#         count = len(result_for_processing)                  # 처리 길이 만큼 count 셋팅
        
#         for proc, item in result_for_processing.items():
#             if item['result'] == False:                     # 결과가 False면 prcess 실행
#                 item['result'] = item['exec']()
#                 print(f"{proc} - {item['result']}")
#             time.sleep(2)                                   # 2초 지연

#         for item in result_for_processing.values():         # 처리 결과가 True면 count 감소
#             if item['result'] == True:
#                 count -= 1
        
#         if count <= 0:      # 0보다 작거나 같으면 처리 결과가 모두 True 이므로 재시도 종료
#             print('count가 0 보다 작거나 같음: 종료')
#             break                                           # 종료
#         else:
#             print('While 반복 지연 5초')                    # 재시작 지연 시간(초) 10분
#             time.sleep(600)
        
#         end_time = datetime.datetime.now()                  # 종료 시간
#         diff_time = end_time - start_time
#         if (diff_time.seconds / 3600) >= 1:                 # 1시간이 지나면
#             break                                           # 종료

# schedule.every().day.at("13:01").do(processing_pkg)


# while True:
#     schedule.run_pending()
#     time.sleep(1)