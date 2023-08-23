import datetime


time_now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
time_now = datetime.datetime.strptime(time_now_str, "%d.%m.%Y %H:%M")

yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
today = datetime.datetime.now().strftime('%d.%m.%Y')  # сегодня
reset_time = datetime.datetime.strptime(yesterday + " " + "00:10", "%d.%m.%Y %H:%M")



print(reset_time)
print(time_now)

print(time_now > reset_time)