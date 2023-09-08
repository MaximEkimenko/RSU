import datetime
import os
import time
import psutil
import json
# TODO подключить log


def find_and_kill_process(process_name: str, kill: bool = False) -> list:
    """
    Функция находит все pid процессы с именем процесса process_name и гасит их при kill=True
    process_name: имя процесса
    """
    pid_list = []
    for process in psutil.process_iter():
        if process.name() == process_name:
            pid_list.append(process.pid)
            if kill:
                try:
                    process.kill()
                except Exception as e:
                    pass
    return sorted(pid_list)


def restart(r_filename: str, restart_file: str, idle_time: int) -> None:
    """
    Программа анализирует последнюю дату изменения json r_filename. В случае разницы от сейчас более idle_time минут
    программа выполняет перезапуск restart_file
    """
    now = datetime.datetime.now()  # сейчас
    # разница времени создания контрольного файла от сейчас
    change_delta = now - datetime.datetime.fromtimestamp(os.path.getmtime(r_filename))
    with open(r_filename, 'r') as file:  # чтение файла - получение словаря данных процесса
        feedback_dict = json.load(file)
    idle_limit = datetime.timedelta(minutes=idle_time)  # предел отсутствия отклика
    if change_delta > idle_limit:  # если предел превышен
        print(f'Отсутствие отклика более 20 минут. '
              f'Последняя обратная связь {feedback_dict["datetime"]}. Перезапуск {now}.')
        # запуск перезагрузки
        try:
            os.system(f'taskkill /f /fi "pid eq {feedback_dict["PARENT_PID"]}"')  # остановка зависшего родит процесса
            os.system(f'taskkill /f /fi "pid eq {feedback_dict["PID"]}"')  # остановка зависшего процесса
        except Exception as err:
            print(err)
        time.sleep(1)  # пауза 1 сек
        find_and_kill_process('chrome.exe', kill=True)  # остановка всех процессов google chrome
        print('Процессы chrome.exe завершены.')
        time.sleep(1)  # пауза 1 сек
        os.startfile(restart_file)  # запуск исполнительного файла
        print(f'Перезапуск выполнен успешно {now}.\n')
        # TODO send report
    else:
        # TODO write report log
        print(f'Программа функционирует нормально {now}. Последняя обратная связь {feedback_dict["datetime"]}.'
              f'Используемые параметры PPID = {feedback_dict["PARENT_PID"]}, PID={feedback_dict["PID"]}')
        pass


if __name__ == '__main__':
    json_file = r"O:\Расчет эффективности\Выгрузки РСУ\r.json"
    restart_bat_file = r'C:\python\RSU\rsu_auto_unload.bat'
    print(f'reloader файла {restart_bat_file} запущен в {datetime.datetime.now()}\n')
    while True:
        try:
            restart(r_filename=json_file, restart_file=restart_bat_file, idle_time=20)
        except Exception as e:
            print(e)
        print(f'Итерация пройдена в {datetime.datetime.now()}\n')
        time.sleep(20*60)





