import datetime
import os

import madmodule

reset_time_interval_start = datetime.datetime.strptime("14:00", "%H:%M")
reset_time_interval_end = datetime.datetime.strptime("14:30", "%H:%M")
time_now_str = datetime.datetime.now().strftime("%H:%M")
time_now = datetime.datetime.strptime(time_now_str, "%H:%M")
today = datetime.datetime.now().strftime('%d.%m.%Y')  # сегодня


# print(datetime.datetime.strptime("00:00", "%H:%M"))
#
#
# print(reset_time_interval_start)
# print(reset_time_interval_end)
# print(time_now)
#
# print(reset_time_interval_end > time_now > reset_time_interval_start)



print(madmodule.create_list(os.getcwd(), result_type='files', extension=f'{today}.xlsx'))
