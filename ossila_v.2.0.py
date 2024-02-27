"""
Created on Fri Nov  3 15:25:11 2023

@author: asics28

Этот скрипт тест для 4х контактного измерения + MPPT
"""

# функциональный вариант
import xlsxwriter as xls
import xtralien, time, colorama
from datetime import datetime

NAME = "Device1"  # имя образца

COM_NUMBER = 3  # выбор порта
# channel = 'smu1'  # SMU channel to use
# vsense_channel = 'vsense1'
AREA = 9.348  # площадь пикселя
POWER = 100  # мощность светового потока mW/cm2

mpp_time = 10  # Время MPPT на каждый пиксель (секунды)
mpp_time_step = 1  # Длительность одной итерации MPPT
dV = 0.005

Voc_stab_time = 10  # длительность стабилизации Voc (секунды)
qss_IV_stab_time = 10  # длительность стабилизации тока при съемке QSS-IV (секунды)

v_start = -0.05  # начальное напряжение (V)
v_end = 1.25  # конечное напряжение (V)
v_increment = 0.05  # шаг по напряжению (V)
v_time_per_point = 0.1  # длительность измерения на каждом шаге (секунды)

WIDTH = 10  # задаем ширину колонок при печати в консоль


def current_density(current):
    """Перевод тока в плотность тока"""
    return round(-current * POWER * 1000 / AREA, 4)


def pce_calc(array1, array2):
    """Функция расчета КПД"""
    power = [i * j for i, j in zip(array1, array2)]  # произведение элементов массивов напряжения и тока
    pce = round(max(power), 2)  # КПД сканирования
    return pce


def forward_scan(v_start, v_end, v_increment, v_time_per_point):
    """
    Измерения ВАХ в прямом направления для L и R
    Принимает параметры съемки
    Возвращает пару списков (напряжение и плотность тока) для каждой стороны (L и R)
    """
    # Подключение к SMU
    device.smu1.set.enabled(True, response=0), device.smu2.set.enabled(True, response=0)
    # Подключение к VS
    device.vsense1.set.enabled(1, response=0), device.vsense2.set.enabled(1, response=0)

    # forward scan
    forward_voltage_left, forward_current_left = [], []  # списки для сбора V и I для левого пикселя
    forward_voltage_right, forward_current_right = [], []  # списки для сбора V и I для правого пикселя
    forward_voltage_L, forward_voltage_R = [], []  # списки для сбора V c вольтметра для обоих пискелей

    print("Forward scan")
    while v_start <= v_end:
        # Set voltage, measure voltage and current
        voltage_left, current_left = device.smu1.oneshot(v_start)[0]
        voltage_L = device.vsense1.measure()[0]
        voltage_right, current_right = device.smu2.oneshot(v_start)[0]
        voltage_R = device.vsense2.measure()[0]

        time.sleep(v_time_per_point)

        forward_current_left.append(current_density(current_left))
        forward_current_right.append(current_density(current_right))
        forward_voltage_L.append(voltage_L)
        forward_voltage_R.append(voltage_R)

        # "Красивая" печать напряжения и плотности тока
        print(colorama.Style.BRIGHT +
              colorama.Fore.RED + f"{round(voltage_L, 3):{WIDTH}} V, {current_density(current_left):{WIDTH}} mA/cm2 |",
              colorama.Fore.BLUE + f" {round(voltage_R, 3):{WIDTH}} V, {current_density(current_right):{WIDTH}} mA/cm2"
              + colorama.Style.RESET_ALL)

        # Увеличиваем заданное напряжение
        v_start += v_increment

    # Reset output voltage and turn off SMU
    device.smu1.set.voltage(0, response=0), device.smu1.set.enabled(False, response=0)
    device.smu2.set.voltage(0, response=0), device.smu2.set.enabled(False, response=0)

    # Отключение от вольтметра
    device.vsense1.set.enabled(0, response=0), device.vsense2.set.enabled(0, response=0)

    return forward_voltage_L, forward_current_left, forward_voltage_R, forward_current_right


def reversed_scan(v_start, v_end, v_increment, v_time_per_point):
    """
    Измерения ВАХ в обратном направлении для L и R
    Принимает параметры съемки
    Возвращает пару списков (напряжение и плотность тока) для каждой стороны (L и R)
    """
    # Подключение к SMU
    device.smu1.set.enabled(True, response=0), device.smu2.set.enabled(True, response=0)
    # Подключение к VS
    device.vsense1.set.enabled(1, response=0), device.vsense2.set.enabled(1, response=0)

    reversed_voltage_left, reversed_current_left = [], []  # списки для сбора V и I для левого пикселя
    reversed_voltage_right, reversed_current_right = [], []  # списки для сбора V и I для правого пикселя
    reversed_voltage_L, reversed_voltage_R = [], []

    print(colorama.Style.RESET_ALL + "Reversed scan")
    while v_end >= (v_start - v_increment):
        voltage_left, current_left = device.smu1.oneshot(v_end)[0]
        rev_voltage_L = device.vsense1.measure()[0]
        voltage_right, current_right = device.smu2.oneshot(v_end)[0]
        rev_voltage_R = device.vsense2.measure()[0]
        time.sleep(v_time_per_point)
        reversed_current_left.append(current_density(current_left))
        reversed_current_right.append(current_density(current_right))
        reversed_voltage_L.append(rev_voltage_L)
        reversed_voltage_R.append(rev_voltage_R)
        # Print measured voltage and current density
        print(colorama.Style.BRIGHT +
              colorama.Fore.RED + f" {round(rev_voltage_L, 3):{WIDTH}} V, {current_density(current_left):{WIDTH}} mA/cm2 |",
              colorama.Fore.BLUE + f" {round(rev_voltage_R, 3):{WIDTH}} V, {current_density(current_right):{WIDTH}} mA/cm2"
              + colorama.Fore.RESET)

        v_end -= v_increment

    # Reset output voltage and turn off SMU
    device.smu1.set.voltage(0, response=0), device.smu1.set.enabled(False, response=0)
    device.smu2.set.voltage(0, response=0), device.smu2.set.enabled(False, response=0)

    # Отключение от вольтметра
    device.vsense1.set.enabled(0, response=0), device.vsense2.set.enabled(0, response=0)

    return reversed_voltage_L, reversed_current_left, reversed_voltage_R, reversed_current_right


def voltage_mpp(pce, array1, array2):
    """Возвращает элементы переданных списков array1 и array2,
    произведение которых равно заданному значению pce
    ЭТУ ФУНКЦИЮ СЛЕДУЕТ ПЕРЕДЕЛАТЬ"""
    power = [i * j for i, j in zip(array1, array2)]  # произведение элементов массивов напряжения и тока

    # Находим индексы элементов, произведение которых равно заданному значению КПД
    pce_indices = [index for index, value in enumerate(power) if round(value, 2) == pce]

    # Соответствующие элементы массивов array1 и array2
    elements_array1 = [array1[index] for index in pce_indices]

    return elements_array1[0]


def mpp_tracking(*args):
    print("Start MMP Tracking")
    # Подключение к SMU
    device.smu1.set.enabled(True, response=0), device.smu2.set.enabled(True, response=0)
    # Подключение к VS
    device.vsense1.set.enabled(1, response=0), device.vsense2.set.enabled(1, response=0)

    MPPT_L, MPPT_R = [], []  # списки для сбора данных MPP
    MPPT_time_L, MPPT_time_R = [], []  # списки для сбора данных времени MPP

    # определение Vmpp для каждого ВАХа
    Vmpp_forw_L = voltage_mpp(pce_forw_left, forward_voltage_L, forward_current_left)
    Vmpp_rev_L = voltage_mpp(pce_rev_left, reversed_voltage_L, reversed_current_left)

    Vmpp_forw_R = voltage_mpp(pce_forw_right, forward_voltage_R, forward_current_right)
    Vmpp_rev_R = voltage_mpp(pce_rev_right, reversed_voltage_R, reversed_current_right)

    voltage_mpp_L = (Vmpp_forw_L + Vmpp_rev_L) / 2  # установка начального напряжения для L
    voltage_mpp_R = (Vmpp_forw_R + Vmpp_rev_R) / 2  # установка начального напряжения для R

    p_L = (pce_rev_left + pce_forw_left) / 2  # установка начальной мощьности для L
    p_R = (pce_rev_right + pce_forw_right) / 2  # установка начальной мощьности для L

    dV_L = dV
    dV_R = dV

    start_time = time.time()
    elapsed_time = 0

    while elapsed_time < mpp_time:
        # Меняем напряжение на "небольшую" величину (на dV)
        voltage_mpp_L += dV_L
        voltage_mpp_R += dV_R

        voltage_mpp_left, current_mpp_left = device.smu1.oneshot(voltage_mpp_L)[0]
        voltage_mpp_right, current_mpp_right = device.smu2.oneshot(voltage_mpp_R)[0]

        # Определение текущей мощности
        p_current_L = round((current_density(current_mpp_left) * voltage_mpp_L), 2)
        p_current_R = round((current_density(current_mpp_right) * voltage_mpp_R), 2)

        MPPT_L.append(p_current_L), MPPT_R.append(p_current_R)

        print(colorama.Style.BRIGHT + colorama.Fore.RED + f"MPP L = {p_current_L:{WIDTH}} % |",
              colorama.Fore.BLUE + f"MPP R = {p_current_R:{WIDTH}} %" +
              colorama.Fore.RESET)

        # Обновление параметров и направления для L
        p_L, voltage_mpp_L, dV_L = update_params_and_direction(p_current_L, p_L, voltage_mpp_L, dV_L)

        # Обновление параметров и направления для R
        p_R, voltage_mpp_R, dV_R = update_params_and_direction(p_current_R, p_R, voltage_mpp_R, dV_R)

        time.sleep(mpp_time_step)
        elapsed_time = time.time() - start_time  # вычисление прошедшего времени

        # Запись времени
        MPPT_time_L.append(elapsed_time), MPPT_time_R.append(elapsed_time)

    print(colorama.Style.BRIGHT + colorama.Fore.RED + f"MPP final L   =   {MPPT_L[-1]} % |",
          colorama.Fore.BLUE + f"MPP final R   =   {MPPT_R[-1]} %" +
          colorama.Fore.RESET)

    # Reset output voltage and turn off SMU
    device.smu1.set.voltage(0, response=0), device.smu1.set.enabled(False, response=0)
    device.smu2.set.voltage(0, response=0), device.smu2.set.enabled(False, response=0)

    # Отключение от вольтметра
    device.vsense1.set.enabled(0, response=0), device.vsense2.set.enabled(0, response=0)

    return MPPT_L, MPPT_R, MPPT_time_L, MPPT_time_R


def update_params_and_direction(p_current, p_max, voltage_mpp, dV):
    """Обновляет параметры мощности и направления шага напряжения при MPPT"""
    if p_current > p_max:
        p_max = p_current
    else:
        voltage_mpp -= 2 * dV
        dV = -dV  # меняем направление шага
    return p_max, voltage_mpp, dV


def Voc_measure():
    print(colorama.Style.RESET_ALL + "Measuring Voc")
    Voc_stab_left_array, Voc_stab_right_array = [], []  # список для сбора напряжения холостого хода

    # Подключение к каналу вольтметра
    device.vsense1.set.enabled(1, response=0), device.vsense2.set.enabled(1, response=0)

    start_time = time.time()
    elapsed_time = 0

    while elapsed_time < Voc_stab_time:
        voltage_oc_left = device.vsense1.measure()[0]
        voltage_oc_right = device.vsense2.measure()[0]

        # Печать измеренного значения Voc
        print(colorama.Style.BRIGHT + colorama.Fore.RED + f"Voc L = {voltage_oc_left:{WIDTH}} V |",
              colorama.Fore.BLUE + f"Voc R = {voltage_oc_right:{WIDTH}} V" + colorama.Style.RESET_ALL)

        Voc_stab_left_array.append(voltage_oc_left)
        Voc_stab_right_array.append(voltage_oc_right)

        time.sleep(1)
        elapsed_time = time.time() - start_time  # вычисление прошедшего времени

    # отключение от канала вольтметра
    device.vsense1.set.enabled(0, response=0), device.vsense2.set.enabled(0, response=0)

    # Принимаем последнее значение в массиве за Voc
    Voc_stab_left = Voc_stab_left_array[-1]
    Voc_stab_right = Voc_stab_right_array[-1]

    print(colorama.Style.BRIGHT + colorama.Fore.RED + f"Voc stab L = {Voc_stab_left:{WIDTH}} V |",
          colorama.Fore.BLUE + f"Voc stab R = {Voc_stab_right:{WIDTH}} V")

    return Voc_stab_left, Voc_stab_right


def measure_QSS_IV(Voc_stab_left, Voc_stab_right, qss_IV_stab_time):
    # список для напряжение, где уже есть значение Voc
    voltage_qss_IV_array_left, voltage_qss_IV_array_right, voltage_qss_IV_array_L, voltage_qss_IV_array_R = [
        Voc_stab_left], [Voc_stab_right], [Voc_stab_left], [
        Voc_stab_right]
    current_qss_IV_array_left, current_qss_IV_array_right = [0], [0]

    print(colorama.Style.RESET_ALL + 'QSS-IV measaring')
    device.smu1.set.enabled(1, response=0), device.smu2.set.enabled(1, response=0)
    device.vsense1.set.enabled(1, response=0), device.vsense2.set.enabled(1, response=0)

    qss_IV_points = [0.95, 0.90, 0.875, 0.850, 0.825, 0.775, 0.7, 0.5, 0.25, 0]

    for voltage_qss_IV in qss_IV_points:
        start_time = time.time()
        elapsed_time = 0
        print(colorama.Style.RESET_ALL + f'For dot: {voltage_qss_IV}')
        while elapsed_time < qss_IV_stab_time:
            voltage_qss_left, current_qss_left = device.smu1.oneshot(voltage_qss_IV * Voc_stab_left)[0]
            voltage_qss_right, current_qss_right = device.smu2.oneshot(voltage_qss_IV * Voc_stab_right)[0]
            voltage_qss_L = device.vsense1.measure()[0]
            voltage_qss_R = device.vsense2.measure()[0]
            print(colorama.Style.BRIGHT +
                  colorama.Fore.RED +
                  f"V = {round(voltage_qss_L, 5)} V,"
                  f" {current_density(current_qss_left):{WIDTH}} mA/cm2 |",
                  colorama.Fore.BLUE +
                  f"V = {round(voltage_qss_R, 5)} V,"
                  f" {current_density(current_qss_right):{WIDTH}} mA/cm2"
                  + colorama.Style.RESET_ALL)
            time.sleep(1)
            elapsed_time = time.time() - start_time  # вычисление прошедшего времени

        voltage_qss_IV_array_L.append(round(voltage_qss_L, 2))
        voltage_qss_IV_array_R.append(round(voltage_qss_R, 2))

        current_qss_IV_array_left.append(current_density(current_qss_left))
        current_qss_IV_array_right.append(current_density(current_qss_right))

    device.smu1.set.enabled(0, response=0), device.smu2.set.enabled(0, response=0)
    device.vsense1.set.enabled(0, response=0), device.vsense2.set.enabled(0, response=0)

    return voltage_qss_IV_array_L, current_qss_IV_array_left, voltage_qss_IV_array_R, current_qss_IV_array_right


def closest_to_zero_index(array):
    """Нахождение индекса элемента в списке, который наиболее близок к 0
    Необходимо для поиска Isc и Voc"""
    return min(range(len(array)), key=lambda i: abs(array[i]))


def data_to_xslx(*params):
    '''Сохранение данных в xslx файл'''
    (NAME, side, forward_voltage, forward_current, reversed_voltage, reversed_current,
     voltage_qss_IV_array, current_qss_IV_array,
     pce_forw, pce_rev,
     MPPT_time, MPPT,
     Isc_forward, Isc_reversed,
     Voc_forward, Voc_reversed, Voc_stab,
     FF_forward, FF_reversed,
     pce_qss, Isc_qss, FF_qss) = params

    # 1. Имя, дата и время создание файла
    Sample_name = (str(NAME) + '-' + (side) + '_QSS_IV_' + str(datetime.today().strftime("%d-%m-%Y_%Hh%Mm")))
    # 2. Работа с файлом
    workbook = xls.Workbook(f'{Sample_name}.xlsx')  # создание нового файла xlsx
    worksheet = workbook.add_worksheet()  # создание нового листа
    bold = workbook.add_format({'bold': 1})
    boldred = workbook.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'bold': 1})

    # данные по имени
    worksheet.write('A1', 'Sample', boldred)
    worksheet.write('B1', f'{Sample_name}')

    # данные по forward скану
    worksheet.write('A23', 'V_forw', bold)
    worksheet.write('B23', 'I_forw', bold)
    worksheet.write_column('A24', forward_voltage)
    worksheet.write_column('B24', forward_current)

    # данные по reversed скану
    worksheet.write('C23', 'V_rev', bold)
    worksheet.write('D23', 'I_rev', bold)
    worksheet.write_column('C24', reversed_voltage)
    worksheet.write_column('D24', reversed_current)

    # "шапка" таблицы данных forward и rewersed
    worksheet.write('B3', 'PCE init', boldred)
    worksheet.write('C3', 'Isc', boldred)
    worksheet.write('D3', 'Voc', boldred)
    worksheet.write('E3', 'FF', boldred)

    # forward PCE, Isc, Voc, FF
    worksheet.write('A4', 'Forward', bold)
    worksheet.write('B4', pce_forw)
    worksheet.write('C4', Isc_forward)
    worksheet.write('D4', Voc_forward)
    worksheet.write('E4', FF_forward)

    # reverse PCE, Isc, Voc, FF
    worksheet.write('A5', 'Reversed', bold)
    worksheet.write('B5', pce_rev)
    worksheet.write('C5', Isc_reversed)
    worksheet.write('D5', Voc_reversed)
    worksheet.write('E5', FF_reversed)

    # "шапка" таблицы данных QSS-IV
    worksheet.write('H3', 'PCE final', boldred)
    worksheet.write('I3', 'Isc', boldred)
    worksheet.write('J3', 'Voc', boldred)
    worksheet.write('K3', 'FF', boldred)

    # QSS-IV PCE, Isc, Voc, FF
    worksheet.write('G4', 'QSS-IV', bold)
    worksheet.write('H4', pce_qss)
    worksheet.write('I4', Isc_qss)
    worksheet.write('J4', Voc_stab)
    worksheet.write('K4', FF_qss)

    # данные по QSS-IV
    worksheet.write('E23', 'qss_V', bold)
    worksheet.write_column('E24', voltage_qss_IV_array)
    worksheet.write('F23', 'qss_I', bold)
    worksheet.write_column('F24', current_qss_IV_array)

    # данные по MPPT
    worksheet.write('G23', 'time', bold)
    worksheet.write_column('G24', MPPT_time)
    worksheet.write('H23', 'MPP', bold)
    worksheet.write_column('H24', MPPT)

    # создание графика
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})

    # добавление данных для графика из колонок
    # forward scan
    chart.add_series({
        'categories': f'={worksheet.name}!$A$24:$A${24 + len(forward_voltage)}',
        'values': f'={worksheet.name}!$B$24:$B${24 + len(forward_current)}',
        'name': 'forward scan',
        'line': {'color': 'blue', 'width': 1},
        'marker': {
            'type': 'circle',
            'size': 2,
            'border': {'color': 'blue'},
            'fill': {'color': 'blue'}, }
    })

    # reversed scan
    chart.add_series({
        'categories': f'={worksheet.name}!$C$24:$C${24 + len(reversed_voltage)}',
        'values': f'={worksheet.name}!$D$24:$D${24 + len(reversed_current)}',
        'name': 'reversed scan',
        'line': {'color': 'red', 'width': 1},
        'marker': {
            'type': 'circle',
            'size': 2,
            'border': {'color': 'red'},
            'fill': {'color': 'red'}, }
    })

    chart.set_title({'name': 'IV curve'})  # название рисунка
    chart.set_x_axis({'name': 'Voltage (V)',  # название оси x
                      'min': 0, 'max': 1.2,
                      'major_unit': 0.2,
                      'crossing': 0,
                      'major_gridlines': {'visible': True, 'line': {'width': 1.25, 'dash_type': 'dash'}}
                      })
    chart.set_y_axis({'name': 'Current density (mA/cm2)',
                      'min': 0, 'max': 30,
                      'major_gridlines': {'visible': True, 'line': {'width': 1.2, 'dash_type': 'dash'}}
                      })  # название оси y

    worksheet.insert_chart('B7', chart)  # вставка графика в файл

    # график QSS-IV
    chart_qss = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})

    chart_qss.add_series({
        'categories': f'={worksheet.name}!$E$24:$E${24 + len(voltage_qss_IV_array)}',
        'values': f'={worksheet.name}!$F$24:$F${24 + len(current_qss_IV_array)}',
        'name': 'QSS-IV',
        'marker': {'type': 'circle', 'size': 5, 'border': {'color': 'blue'}, 'fill': {'color': 'blue'},
                   'line': {'color': 'blue', 'width': 1},
                   }})

    chart_qss.set_title({'name': 'QSS-IV curve'})  # название рисунка
    chart_qss.set_x_axis({'name': 'Voltage (V)',  # параметры оси x
                          'min': 0, 'max': 1.2,
                          'major_unit': 0.2,
                          'crossing': 0,
                          'major_gridlines': {'visible': True, 'line': {'width': 1.25, 'dash_type': 'dash'}}
                          })
    chart_qss.set_y_axis({'name': 'Current density (mA/cm2)',  # параметры оси y
                          'min': 0, 'max': 30,
                          'major_gridlines': {'visible': True, 'line': {'width': 1.2, 'dash_type': 'dash'}}
                          })

    worksheet.insert_chart('J7', chart_qss)  # вставка графика в файл

    # график MPPT
    chart_mppt = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})

    chart_mppt.add_series({
        'categories': f'={worksheet.name}!$G$24:$E${24 + len(MPPT_time)}',
        'values': f'={worksheet.name}!$H$24:$H${24 + len(MPPT)}',
        'name': 'QSS-IV',
        'marker': {'type': 'circle', 'size': 5, 'border': {'color': 'blue'}, 'fill': {'color': 'blue'},
                   'line': {'color': 'blue', 'width': 1},
                   }})

    chart_mppt.set_title({'name': 'MPPTracking curve'})  # название рисунка
    chart_mppt.set_x_axis({'name': 'time (sec)',  # параметры оси x
                           # 'min': 0, 'max': 1.2,
                           # 'major_unit': 0.2,
                           # 'crossing': 0,
                           'major_gridlines': {'visible': True, 'line': {'width': 1.25, 'dash_type': 'dash'}}
                           })
    chart_mppt.set_y_axis({'name': 'MPP (%)',  # параметры оси y
                           # 'min': 0, 'max': 30,
                           'major_gridlines': {'visible': True, 'line': {'width': 1.2, 'dash_type': 'dash'}}
                           })

    worksheet.insert_chart('R7', chart_mppt)  # вставка графика MPPT в файл

    workbook.close()  # закрытие файла

    if side == 'L':
        print(colorama.Style.BRIGHT + colorama.Fore.RED + f"Data for L saved to file: {Sample_name}")
    elif side == 'R':
        print(colorama.Style.BRIGHT + colorama.Fore.BLUE + f"Data for R saved to file: {Sample_name}")

def main():
    print(f"Start measuring {NAME}")
    # Connect to the Source Measure Unit using USB
    with xtralien.Device(f'COM{COM_NUMBER}') as device:
        # forward scan
        forward_voltage_L, forward_current_left, forward_voltage_R, forward_current_right = forward_scan(v_start,
                                                                                                         v_end,
                                                                                                         v_increment,
                                                                                                         v_time_per_point)

        time.sleep(1)

        # PCEforw calc
        pce_forw_left = pce_calc(forward_voltage_L, forward_current_left)
        pce_forw_right = pce_calc(forward_voltage_R, forward_current_right)

        print(colorama.Style.BRIGHT + colorama.Fore.RED + f"PCE forw L = {pce_forw_left:{WIDTH}} % |",
              colorama.Fore.BLUE + f"PCE forw R = {pce_forw_right:{WIDTH}} %" +
              colorama.Fore.RESET)

        # Voc forw calc
        Voc_forward_left = forward_voltage_L[closest_to_zero_index(forward_current_left)]
        Voc_forward_right = forward_voltage_R[closest_to_zero_index(forward_current_right)]

        # Isc forw calc
        Isc_forward_left = forward_current_left[closest_to_zero_index(forward_voltage_L)]
        Isc_forward_right = forward_current_right[closest_to_zero_index(forward_voltage_R)]

        # FF forw calc
        FF_forward_left = round((pce_forw_left / (Voc_forward_left * Isc_forward_left)) * 100, 1)
        FF_forward_right = round((pce_forw_right / (Voc_forward_right * Isc_forward_right)) * 100, 1)

        time.sleep(3)

        # reverse scan
        reversed_voltage_L, reversed_current_left, reversed_voltage_R, reversed_current_right = reversed_scan(v_start,
                                                                                                              v_end,
                                                                                                              v_increment,
                                                                                                              v_time_per_point)

        # PCErev calc
        pce_rev_left = pce_calc(reversed_voltage_L, reversed_current_left)
        pce_rev_right = pce_calc(reversed_voltage_R, reversed_current_right)

        print(colorama.Style.BRIGHT + colorama.Fore.RED + f"PCE rev L = {pce_rev_left:{WIDTH}} % |",
              colorama.Fore.BLUE + f" PCE rev R = {pce_rev_right:{WIDTH}} %" +
              colorama.Fore.RESET)

        # Voc rev calc
        Voc_reversed_left = reversed_voltage_L[closest_to_zero_index(reversed_current_left)]
        Voc_reversed_right = reversed_voltage_R[closest_to_zero_index(reversed_current_right)]

        # Isc rev calc
        Isc_reversed_left = reversed_current_left[closest_to_zero_index(reversed_voltage_L)]
        Isc_reversed_right = reversed_current_right[closest_to_zero_index(reversed_voltage_R)]

        # FF rev calc
        FF_reversed_left = round((pce_rev_left / (Voc_reversed_left * Isc_reversed_left)) * 100, 1)
        FF_reversed_right = round((pce_rev_right / (Voc_reversed_right * Isc_reversed_right)) * 100, 1)

        time.sleep(1)

        # MPPTracking (Pertrub and Observe)
        MPPT_L, MPPT_R, MPPT_time_L, MPPT_time_R = mpp_tracking(pce_forw_left, forward_voltage_L, forward_current_left,
                                                                pce_rev_left, reversed_voltage_L, reversed_current_left,
                                                                pce_forw_right, forward_voltage_R,
                                                                forward_current_right,
                                                                pce_rev_right, reversed_voltage_R,
                                                                reversed_current_right,
                                                                mpp_time, mpp_time_step)

        time.sleep(1)

        # Measuring stabilized Voc
        Voc_stab_left, Voc_stab_right = Voc_measure()

        time.sleep(1)

        # Measaring QSS-IV
        voltage_qss_IV_array_L, current_qss_IV_array_left, voltage_qss_IV_array_R, current_qss_IV_array_right = measure_QSS_IV(
            Voc_stab_left, Voc_stab_right, qss_IV_stab_time)

        # PCE QSS-IV calc
        pce_qss_left = pce_calc(voltage_qss_IV_array_L, current_qss_IV_array_left)
        pce_qss_right = pce_calc(voltage_qss_IV_array_R, current_qss_IV_array_right)

        # Voc qss calc не нужен, т.к его уже посчитали за Voc_stab

        # Isc rev calc
        Isc_qss_left = current_qss_IV_array_left[closest_to_zero_index(voltage_qss_IV_array_L)]
        Isc_qss_right = current_qss_IV_array_right[closest_to_zero_index(voltage_qss_IV_array_R)]

        # FF rev calc
        FF_qss_left = round((pce_qss_left / (Voc_stab_left * Isc_qss_left)) * 100, 1)
        FF_qss_right = round((pce_qss_right / (Voc_stab_right * Isc_qss_right)) * 100, 1)

        data_to_xslx(NAME, 'L',
                     forward_voltage_L, forward_current_left,
                     reversed_voltage_L, reversed_current_left,
                     voltage_qss_IV_array_L, current_qss_IV_array_left,
                     pce_forw_left, pce_rev_left,
                     MPPT_time_L, MPPT_L,
                     Isc_reversed_left, Isc_reversed_left,
                     Voc_forward_left, Voc_reversed_left, Voc_stab_left,
                     FF_forward_left, FF_reversed_left,
                     pce_qss_left, Isc_qss_left, FF_qss_left)

        data_to_xslx(NAME, "R",
                     forward_voltage_R, forward_current_right,
                     reversed_voltage_R, reversed_current_right,
                     voltage_qss_IV_array_R, current_qss_IV_array_right,
                     pce_forw_right, pce_rev_right,
                     MPPT_time_R, MPPT_R,
                     Isc_forward_right, Isc_forward_right,
                     Voc_forward_right, Voc_reversed_right, Voc_stab_right,
                     FF_forward_right, FF_reversed_right,
                     pce_qss_right, Isc_qss_right, FF_qss_right)

        print(colorama.Style.RESET_ALL + "Finish")

# запуск программы
if __name__ == '__main__':
    main()
