import datetime


now_str = datetime.datetime.now().strftime('%H:%M')
now = datetime.datetime.strptime(now_str, '%H:%M')


reset_time = datetime.datetime.strptime('08:00', '%H:%M')
# print(type(now), type(reset_time))

# print(now < reset_time)

yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
print(yesterday)