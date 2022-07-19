
import schedule, time
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import bank_data
import preparation

schedule.every().day.at("07:00").do(preparation.ready)

schedule.every().day.at("07:05").do(down_opera.down)
schedule.every().day.at("07:10").do(down_kicc.down)

schedule.every().day.at("07:20").do(opera_xml.to_opera_df)
schedule.every().day.at("07:25").do(kicc_history.to_card_history_df)
schedule.every().day.at("07:30").do(kicc_receipts.to_receipts_history_df)

schedule.every().day.at("07:35").do(edi_opera.merge_edi_opera)

schedule.every().day.at("12:55").do(down_nh.down)
schedule.every().day.at("13:00").do(down_ibk.down)
schedule.every().day.at("13:05").do(bank_data.to_bank_df)


while True:
    schedule.run_pending()
    time.sleep(1)