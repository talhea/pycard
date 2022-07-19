from pytimekr import pytimekr
import datetime

today = datetime.datetime.now()

print(today.date() - datetime.timedelta(1))
print(today - datetime.timedelta(1))