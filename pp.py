
import shutil
import schedule, time, datetime

def copy_folder():
    dir_date = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%Y%m%d")  # 어제 날짜 포맷
    source_dir = './dtdata/' + dir_date
    target_dir = 'C:/FA/IT/' + dir_date
    
    shutil.copytree(source_dir, target_dir)

copy_folder()