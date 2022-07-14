'''
프로그램별로 실행한 결과를 바탕으로 실패한 프로그램 재실행
문제점) 한번 실패한 과정은 계속 실패만 하므로, 더 이상 진행할 필요가 없음
'''
import schedule, time, datetime
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import bank_data
import preparation

# 실행 결과 저장 버퍼
result_for_processing = {
    # 'procee 이름(모듈 이름)' : {'exex': 실행할 함수이름, 'result': 실행결과 bool}
    'preparation': {'exec': preparation.ready, 'result': False},
    'down_opera': {'exec': down_opera.down, 'result': False},
    'down_kicc': {'exec': down_kicc.down, 'result': False},
    'down_nh': {'exec': down_nh.down, 'result': False},
    'opera_xml': {'exec': opera_xml.to_opera_df, 'result': False},
    'kicc_history': {'exec': kicc_history.to_card_history_df, 'result': False},
    'kicc_receipts': {'exec': kicc_receipts.to_receipts_history_df, 'result': False},
    'edi_opera': {'exec': edi_opera.merge_edi_opera, 'result': False},
    'down_ibk': {'exec': down_ibk.down, 'result': False},
    'bank_data': {'exec': bank_data.to_bank_df, 'result': False},
}

# 시간 타이머
start_time = datetime.datetime.now()            # 시작 시간
end_time, diff_time = 0, 0                      # 끝 시간, 시간 차이 

# 모든 처리결과('result')가 True일때까지  process 반복 실행
def processing_pkg():
    while True:
        count = len(result_for_processing)                  # 처리 길이 만큼 count 셋팅
        
        for proc, item in result_for_processing.items():
            if item['result'] == False:                     # 결과가 False면 prcess 실행
                item['result'] = item['exec']()
                print(f"{proc} - {item['result']}")
            time.sleep(2)                                   # 2초 지연

        for item in result_for_processing.values():         # 처리 결과가 True면 count 감소
            if item['result'] == True:
                count -= 1
        
        if count <= 0:      # 0보다 작거나 같으면 처리 결과가 모두 True 이므로 재시도 종료
            print('count가 0 보다 작거나 같음: 종료')
            break                                           # 종료
        else:
            print('While 반복 지연 5초')                    # 재시작 지연 시간(초) 10분
            time.sleep(600)
        
        end_time = datetime.datetime.now()                  # 종료 시간
        diff_time = end_time - start_time
        if (diff_time.seconds / 3600) >= 1:                 # 1시간이 지나면
            break                                           # 종료

schedule.every().day.at("13:01").do(processing_pkg)


while True:
    schedule.run_pending()
    time.sleep(1)