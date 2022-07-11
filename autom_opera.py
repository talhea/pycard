
import schedule, time
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import preparation

schedule.every().day.at("07:00").do(preparation.ready)

schedule.every().day.at("07:05").do(down_opera.down)
schedule.every().day.at("07:12").do(down_kicc.down)
schedule.every().day.at("07:15").do(down_nh.down)

schedule.every().day.at("07:20").do(opera_xml.to_opera_df)
schedule.every().day.at("07:25").do(edi_opera.merge_edi_opera)

schedule.every().day.at("13:01").do(down_ibk.down)

# schedule.every().day.at("13:44").do(preparation.ready)

# schedule.every().day.at("13:46").do(down_opera.down)
# schedule.every().day.at("13:48").do(down_kicc.down)
# schedule.every().day.at("13:50").do(down_nh.down)

# schedule.every().day.at("13:52").do(opera_xml.to_opera_df)
# schedule.every().day.at("13:54").do(edi_opera.merge_edi_opera)

# schedule.every().day.at("13:56").do(down_ibk.down)

while True:
    schedule.run_pending()
    time.sleep(1)