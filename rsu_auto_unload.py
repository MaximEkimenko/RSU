import datetime
import os
import time
import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By


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
        yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%d.%m.%Y')
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

            reset_time = datetime.datetime.strptime("00:10", "%H:%M")
            temp_file_name = rf"U:\Выгрузка РСУ {rsu_name} за {yesterday}.xlsx"  # имя xlsx копии в темп
            if time_now > reset_time:  # сохранение результатов в Temp при достижении reset_time
                try:
                    res_wb.save(temp_file_name)
                    print(f'{temp_file_name} copied!')
                    rsu_previews_data = {'1-1': set(), '2-1': set(), '3-1': set(), '4-1': set(), '4-2': set(),
                                         '4-3': set(),
                                         '4-4': set(), '5-1': set(), '6-1': set()}
                    print(f'rsu_previews_data reset complete!')
                except Exception as e:
                    print(e)
        print('Итерация пройдена, ожидание следующей. Запуск через 2 минуты.')

        time.sleep(2 * 60)


if __name__ == '__main__':
    auto_unload()