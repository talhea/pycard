'''
프로그램별로 실행한 결과를 바탕으로 실패한 프로그램 재실행
'''
from unittest import result
import schedule, time
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import preparation

# schedule.every().day.at("07:00").do(preparation.ready)

# schedule.every().day.at("07:05").do(down_opera.down)
# schedule.every().day.at("07:10").do(down_kicc.down)
# schedule.every().day.at("07:15").do(down_nh.down)

# schedule.every().day.at("07:20").do(opera_xml.to_opera_df)
# schedule.every().day.at("07:25").do(kicc_history.to_card_history_df)
# schedule.every().day.at("07:30").do(kicc_receipts.to_receipts_history_df)

# schedule.every().day.at("07:35").do(edi_opera.merge_edi_opera)

# schedule.every().day.at("13:01").do(down_ibk.down)

# 실행 결과를 저장, 그 저장 결과에 따라서 반복... 반복은 스케줄러

# while True:
#     schedule.run_pending()
#     time.sleep(60)

# 실행 결과 저장 버퍼
result_for_process = {
    'preparation': {'exec': preparation.ready, 'result': False},
    # 'down_opera': {'exec': down_opera.down, 'result': False},
    # 'down_kicc': {'exec': down_kicc.down, 'result': False},
    # 'down_nh': {'exec': down_nh.down, 'result': False},
    # 'opera_xml': {'exec': opera_xml.to_opera_df, 'result': False},
}

print(result_for_process)
result_for_process['preparation']['result'] = result_for_process['preparation']['exec']()
print(result_for_process['preparation']['result'])