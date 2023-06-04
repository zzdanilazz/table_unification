import datetime
import os
import time
import yadisk
import win32api
import win32com.client
from win32com.universal import com_error

input_path = r'C:\Входные данные'

y = yadisk.YaDisk(token="y0_AgAAAAAF8jDjAAkUmAAAAADbNxl7Uq2hmZ4ySMC_1nIEKHNgp3zpgTE")


def tables_download():
    y.download("ОРП 30ЛП6/ОРП 30ЛП6.xlsx", "C:/Входные данные/ОРП 30ЛП6.xlsx")
    print('ОРП 30ЛП6.xlsx скачан')
    y.download("ОРП 40ЛП138А/ОРП 40ЛП138А.xlsx", "C:/Входные данные/ОРП 40ЛП138А.xlsx")
    print('ОРП 40ЛП138А.xlsx скачан')
    y.download("ОРП 40ЛП92/ОРП 40ЛП92.xlsx", "C:/Входные данные/ОРП 40ЛП92.xlsx")
    print('ОРП 40ЛП92.xlsx скачан')
    y.download("ОРП 50ВЛКСМ2/ОРП 50ВЛКСМ2.xlsx", "C:/Входные данные/ОРП 50ВЛКСМ2.xlsx")
    print('ОРП 50ВЛКСМ2.xlsx скачан')
    y.download("ОРП АБ18/ОРП АБ18.xlsx", "C:/Входные данные/ОРП АБ18.xlsx")
    print('ОРП АБ18.xlsx скачан')
    y.download("ОРП В172/ОРП В172.xlsx", "C:/Входные данные/ОРП В172.xlsx")
    print('ОРП В172.xlsx скачан')
    y.download("ОРП Г15/ОРП Г15.xlsx", "C:/Входные данные/ОРП Г15.xlsx")
    print('ОРП Г15.xlsx скачан')
    y.download("ОРП Г2/ОРП Г2.xlsx", "C:/Входные данные/ОРП Г2.xlsx")
    print('ОРП Г2.xlsx скачан')
    y.download("ОРП З33/ОРП З33.xlsx", "C:/Входные данные/ОРП З33.xlsx")
    print('ОРП З33.xlsx скачан')
    y.download("ОРП К11/ОРП К11к2.xlsx", "C:/Входные данные/ОРП К11К2.xlsx")
    print('ОРП К11К2.xlsx скачан')
    y.download("ОРП КМ287/ОРП КМ287.xlsx", "C:/Входные данные/ОРП КМ287.xlsx")
    print('ОРП КМ287.xlsx скачан')
    y.download("ОРП КМ308к1/ОРП КМ308К1.xlsx", "C:/Входные данные/ОРП КМ308К1.xlsx")
    print('ОРП КМ308К1.xlsx скачан')
    y.download("ОРП Л17/ОРП Л17.xlsx", "C:/Входные данные/ОРП Л17.xlsx")
    print('ОРП Л17.xlsx скачан')
    y.download("ОРП М103А/ОРП М103А.xlsx", "C:/Входные данные/ОРП М103А.xlsx")
    print('ОРП М103А.xlsx скачан')
    y.download("ОРП М17/ОРП М17.xlsx", "C:/Входные данные/ОРП М17.xlsx")
    print('ОРП М17.xlsx скачан')
    y.download("ОРП М8/ОРП М8.xlsx", "C:/Входные данные/ОРП М8.xlsx")
    print('ОРП М8.xlsx скачан')
    y.download("ОРП П7/ОРП П7.xlsx", "C:/Входные данные/ОРП П7.xlsx")
    print('ОРП П7.xlsx скачан')
    y.download("ОРП С50/ОРП С50.xlsx", "C:/Входные данные/ОРП С50.xlsx")
    print('ОРП С50.xlsx скачан')
    y.download("ОРП Х45/ОРП Х45.xlsx", "C:/Входные данные/ОРП Х45.xlsx")
    print('ОРП Х45.xlsx скачан')


def table_upload():
    y.upload('C:/Выходные данные/Общая таблица.xlsx', '/Общая таблица.xlsx', overwrite=True)


def kill_excel():
    os.system('TASKKILL /F /IM excel.exe')


def check_and_repair():
    # создаем COM-объект
    excel = win32com.client.DispatchEx("Excel.Application")

    # видим, что происходит с файлами
    excel.Visible = True
    excel.DisplayAlerts = True

    directory = os.fsencode(input_path)

    # перебираем входные таблицы в директории
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        try:
            wb = excel.Workbooks.Open('C:/Входные данные/{}'.format(filename))
            wb.Close()
        except com_error as reason:
            print(reason)
            excel.Quit()

            win32api.ShellExecute(0, 'open', 'C:/Входные данные/{}'.format(filename), '', '', 1)
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = True

            # восстановление
            excel.SendKeys("{LEFT}", 0)
            excel.SendKeys("{ENTER}", 0)
            excel.SendKeys("{ENTER}", 0)

            # закрытие
            excel.SendKeys("^w", 0)
            # {ENTER} не работает после предыдущей команды, пришлось юзать ~
            excel.SendKeys("~", 0)
            excel.SendKeys("~", 0)
            excel.SendKeys("{LEFT}", 0)
            excel.SendKeys("~", 0)
            time.sleep(2)

    # закрываем COM объект
    excel.Quit()
    kill_excel()


def unify():
    # создаем COM-объект
    excel = win32com.client.Dispatch("Excel.Application")

    # видим, что происходит с файлами
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    directory = os.fsencode(input_path)

    wb0 = excel.Workbooks.Open('C:/Выходные данные/Общая таблица.xlsx')
    sheet0 = wb0.ActiveSheet
    dates = [v[0] for v in sheet0.Range('A3:A21548').Value]
    orps = []
    i = 1

    for file in os.listdir(directory):
        orps.append(os.fsdecode(file)[:-5])

    # 1) проходимся по вх.таблицам
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        print('Обработка', filename, ', прогресс: {}/19'.format(i))
        wb = excel.Workbooks.Open('C:/Входные данные/{}'.format(filename))

        sheet = wb.ActiveSheet

        # 2) проходимся по строкам (начиная с 3-й) текущей таблицы
        current_string_input = 3
        while sheet.Cells(current_string_input, 1).Value is not None:
            # 3) сохраняем дату (1 столбец), начало смены (4), конец смены (5), наличные (6), эквайринг (7)
            # и название вх.таблицы в список "days"
            days = [sheet.Cells(current_string_input, 1).Value,  # дата
                    sheet.Cells(current_string_input, 4).Value,  # начало смены
                    sheet.Cells(current_string_input, 5).Value,  # конец смены
                    sheet.Cells(current_string_input, 6).Value,  # наличные
                    sheet.Cells(current_string_input, 7).Value,  # эквайринг
                    filename[:-5]]  # название вх.таблицы

            # 4) ищем дату из списка "days" в общ.таблице и сохраняем номер строки с этой датой
            # в переменную "current_string_output"

            try:
                current_string_output = dates.index(days[0]) + 3
            except ValueError:
                sheet.Cells(current_string_input, 1).Value = sheet.Cells(current_string_input - 1,
                                                                         1).Value + datetime.timedelta(days=1)
                print('\033[31m {0}{1}'.format(sheet.Cells(current_string_input, 1).Value, ': проблемная дата!'),
                      '\033[0m')
                continue

            # 5) ищем название вх.таблицы из списка "days" в общ.таблице и сохраняем
            # номер строки с этим названием в переменную "current_string_output"
            try:
                current_string_output += orps.index(days[5])
            except ValueError:
                print('Не нашли{}'.format(days[5]))
            # 6) записываем в "current_string_output"-ю строку в 3-6 столбцы данные из "days" (1-4 элементы)
            # соответственно
            sheet0.Cells(current_string_output, 3).Value = days[1]
            sheet0.Cells(current_string_output, 4).Value = days[2]
            sheet0.Cells(current_string_output, 5).Value = days[3]
            sheet0.Cells(current_string_output, 6).Value = days[4]

            current_string_input += 1
        wb.Save()
        wb.Close()

        i += 1

    wb0.Save()
    wb0.Close()
    excel.Quit()
    print('Готово!')


# excel0.Quit()

if __name__ == '__main__':
    kill_excel()
    tables_download()
    check_and_repair()
    unify()
    table_upload()
    input()
