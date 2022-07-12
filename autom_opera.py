
import schedule, time
import edi_opera, opera_xml
import down_opera, down_kicc, down_ibk, down_nh
import kicc_history, kicc_receipts
import preparation

schedule.every().day.at("07:00").do(preparation.ready)

schedule.every().day.at("07:05").do(down_opera.down)
schedule.every().day.at("07:10").do(down_kicc.down)
schedule.every().day.at("07:15").do(down_nh.down)

schedule.every().day.at("07:20").do(opera_xml.to_opera_df)
schedule.every().day.at("07:25").do(kicc_history.to_card_history_df)
schedule.every().day.at("07:30").do(kicc_receipts.to_receipts_history_df)

schedule.every().day.at("07:35").do(edi_opera.merge_edi_opera)

schedule.every().day.at("13:01").do(down_ibk.down)


# schedule.every().day.at("13:44").do(preparation.ready)

# schedule.every().day.at("13:46").do(down_opera.down)
# schedule.every().day.at("13:48").do(down_kicc.down)
# schedule.every().day.at("13:50").do(down_nh.down)

# schedule.every().day.at("13:52").do(opera_xml.to_opera_df)
# schedule.every().day.at("13:54").do(edi_opera.merge_edi_opera)

# schedule.every().day.at("13:56").do(down_ibk.down)
# schedule.every().day.at("13:58").do(kicc_history.to_card_history_df)
# schedule.every().day.at("14:00").do(kicc_receipts.to_receipts_history_df)



while True:
    schedule.run_pending()
    time.sleep(1)