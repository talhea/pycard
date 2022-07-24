
import schedule, datetime, time
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import bank_data, pivot_on_excel
import preparation

today = datetime.datetime.now()
yesterday = today - datetime.timedelta(1)

def get_date():
    today = datetime.datetime.now()
    yesterday = today - datetime.timedelta(1)

schedule.every().day.at("06:50").do(get_date)

# preparation.ready(yesterday)
schedule.every().day.at("07:00").do(preparation.ready, yesterday)

# down_opera.down(yesterday)
schedule.every().day.at("07:05").do(down_opera.down, yesterday)

# down_kicc.down(today)
schedule.every().day.at("07:10").do(down_kicc.down, today)

# opera_xml.to_opera_df(yesterday)
schedule.every().day.at("07:20").do(opera_xml.to_opera_df, yesterday)

# kicc_history.to_card_history_df(yesterday)
schedule.every().day.at("07:25").do(kicc_history.to_card_history_df, yesterday)

# kicc_receipts.to_receipts_history_df(today)
schedule.every().day.at("07:30").do(kicc_receipts.to_receipts_history_df, today)

# edi_opera.merge_edi_opera(yesterday)
schedule.every().day.at("07:35").do(edi_opera.merge_edi_opera, yesterday)

# pivot_on_excel.to_xcel(yesterday)
schedule.every().day.at("07:40").do(pivot_on_excel.to_excel, yesterday)

# down_nh.down(today)
schedule.every().day.at("12:55").do(down_nh.down, today)

# down_ibk.down(today)
schedule.every().day.at("13:00").do(down_ibk.down, today)

# bank_data.to_bank_df(today)
schedule.every().day.at("13:05").do(bank_data.to_bank_df, today)


while True:
    schedule.run_pending()
    time.sleep(1)
# schedule.run_all()