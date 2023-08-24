import datetime
import logging
import os
import time
import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from madmodule import create_list
import shutil

logging.basicConfig(level=logging.DEBUG, filename='log.txt', filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")
logger = logging.getLogger('logger')
# logger.setLevel(logging.CRITICAL)
# logging.disable(logging.CRITICAL)


def auto_unload():
    # TODO добавить логи вместо принтов
    # инициализация словаря предыдущих значений
    rsu_previews_data = {'1-1': set(), '2-1': set(), '3-1': set(), '4-1': set(), '4-2': set(), '4-3': set(),
                         '4-4': set(), '5-1': set(), '6-1': set()}
    # бесконечный цикл
    while True:
        time_now_str = datetime.datetime.now().strftime("%H:%M")
        time_now = datetime.datetime.strptime(time_now_str, "%H:%M")
        print(f'Запуск в {time_now_str}.')
        today = datetime.datetime.now().strftime('%d.%m.%Y')  # сегодня
        # словарь РСУ
        rsu = {'1-1': {'ip': 'https://192.168.10.111/documentation/documentation.html', 'location': 'R1'},
               '2-1': {'ip': 'https://192.168.10.138/documentation/documentation.html', 'location': 'R1'},
               '3-1': {'ip': 'https://192.168.10.113/documentation/documentation.html', 'location': 'R2'},
               '4-1': {'ip': 'https://192.168.10.114/documentation/documentation.html', 'location': 'R2'},
               '4-2': {'ip': 'https://192.168.10.115/documentation/documentation.html', 'location': 'R2'},
               '4-3': {'ip': 'https://192.168.10.116/documentation/documentation.html', 'location': 'R2'},
               '4-4': {'ip': 'https://192.168.10.117/documentation/documentation.html', 'location': 'R2'},
               '5-1': {'ip': 'https://192.168.10.137/documentation/documentation.html', 'location': 'R1'},
               '6-1': {'ip': 'https://192.168.10.112/documentation/documentation.html', 'location': 'R2'},
               }
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")  # включение headless режима
        options.add_argument('--ignore-certificate-errors')  # отключение сообщений ошибки сертификатов
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(options=options)  # назначение драйвера
        for rsu_name in rsu.keys():
            result_file_name = rf"{os.getcwd()}\Выгрузка РСУ {rsu_name} за {today}.xlsx"  # имя xlsx файла
            rsu_result_list = []  # список результатов
            time_sum = 0  # сумма дуги
            try:
                print(rsu[rsu_name]['ip'])
                try:
                    req = requests.get(rsu[rsu_name]['ip'], verify=False)  # проверка включенности аппарата
                    req_status = True
                except Exception as e:
                    print(e)
                    req_status = False
                    print(rsu_name, 'Недоступен.')
                if req_status:  # если страница открылась
                    is_updated = False  # переменная факта обновления данных
                    if req.status_code == 200:  # если страница загрузилась
                        # если файл есть, то создаем, иначе открываем
                        if os.path.isfile(result_file_name):
                            res_wb = openpyxl.load_workbook(result_file_name, data_only=True)
                            res_sh = res_wb.active
                        else:
                            res_wb = openpyxl.Workbook()
                            res_sh = res_wb.active
                        driver.get(rsu[rsu_name]['ip'])
                        time.sleep(30)
                        data = driver.find_elements(By.CLASS_NAME, 'rowElement')
                        for row in data:  # обработка строк для отчёта
                            line = row.text  # текст элемента
                            if line not in rsu_previews_data[rsu_name]:
                                coma_index = line.find(',')  # индекс запятой
                                if coma_index != -1:
                                    arc_date = line[coma_index - 10:coma_index]
                                else:
                                    arc_date = 0
                                # индекс строки до первого разрыва страницы-отделение интервала времени работы аппарата
                                break_index = line.find('\n')
                                if break_index != -1:
                                    arc_period = line[coma_index + 2:break_index]
                                else:
                                    arc_period = 0
                                # индекс до второго разрыва страницы - отделение времени горения дуги
                                break_index_2 = line.find('\n', break_index + 1, len(line))
                                if break_index_2 != -1:
                                    arc_time = line[break_index + 1:break_index_2 - 2]
                                else:
                                    arc_time = 0
                                if str(arc_date).strip() == str(today).strip():
                                    rsu_result_list.append([rsu_name, arc_date, arc_period, float(arc_time), '',
                                                            line.replace('\n', '')])
                                    time_sum = time_sum + float(arc_time)
                                    res_sh.append([rsu_name, arc_date, arc_period, float(arc_time), '',
                                                   line.replace('\n', '')])
                                    rsu_previews_data[rsu_name].add(line)  # добавление строки в уже выгруженные
                                    print(f'Заполнено.')
                                    # print(f'Заполнено: {line}')
                                    is_updated = True
                            else:
                                print(f'Данные для {rsu_name} в не изменились.')
                        # print(rsu_result_list)
                        # print(rsu_previews_data[rsu_name])
                        arc_value = time_sum / 60 / 60  # значение дуги
                        print(arc_value, rsu_name)
                        time.sleep(5)
                        if is_updated:
                            res_wb.save(result_file_name)
                            print(f'Файл {result_file_name} обновлен.')
                        else:
                            print(f'Данные не изменились. Файл {result_file_name} НЕ обновлен.')
                    else:
                        print(rsu_name, 'Недоступен.')

            except Exception as e:
                print(e)
        reset_time_interval_start = datetime.datetime.strptime("00:01", "%H:%M")
        reset_time_interval_end = datetime.datetime.strptime("00:41", "%H:%M")
        yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')  # вчера
        # сохранение результатов в Temp при попадании в reset_time интервал
        if reset_time_interval_end > time_now > reset_time_interval_start:
            try:
                file_list = create_list(dir_path=os.getcwd(), result_type='files', extension=f'{yesterday}.xlsx')
                print(file_list)
                for file_name in file_list:
                    shutil.copy(file_name, r"U:\\")
                    print(f'{file_name} copied!')
                rsu_previews_data = {'1-1': set(), '2-1': set(), '3-1': set(), '4-1': set(), '4-2': set(),
                                     '4-3': set(),
                                     '4-4': set(), '5-1': set(), '6-1': set()}
                print(f'rsu_previews_data reset complete!')
            except Exception as e:
                print(e, "ФАЙЛЫ НЕ СКОПИРОВАНЫ!")
        print('Итерация пройдена, ожидание следующей. Запуск через 2 минуты.')

        time.sleep(2 * 60)


if __name__ == '__main__':
    auto_unload()
