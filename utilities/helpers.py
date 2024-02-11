import numpy as np
import pandas as pd
from pandas import DataFrame

from configs import config


def pivot_helper(file_name: str, form_type: str) -> list:
    """Создаёт списки сводных таблиц для каждого грузополучателя, в соответствии со статьёй бюджета."""
    data = pd.read_excel(file_name, sheet_name='Sheet1')
    data.rename(columns={'Раздел ГКПЗ': 'Раздел_ГКПЗ'}, inplace=True)
    data['Завод'].replace(config.kts_factories, '7Q61', inplace=True)  # объединяем позиции для КТС
    data['Завод'].replace(config.dts_factories, '7QB1', inplace=True)  # объединяем позиции для ДТС
    data['Раздел_ГКПЗ'].replace(config.repair_budget, 'РЕМОНТ', inplace=True)
    data['Раздел_ГКПЗ'].replace(config.exploitation_budget, 'ЭКСПЛУАТАЦИЯ', inplace=True)
    data['Раздел_ГКПЗ'].replace(config.investments_budget, 'ИНВЕСТИЦИИ', inplace=True)
    # распределение позиций ЦРС по заводам
    data['Завод'] = data['Наименование МВЗ'].map(config.crs).fillna(data['Завод'])
    # получаем наименование лота и записываем его в конфиг-файл
    config.lot_name = data['Наименование лота'].iloc[0].strip()
    supply_months = get_supply_months()  # годы/месяцы поставки
    empty_rows = [config.columns.copy() for _ in supply_months]
    for index in range(len(empty_rows)):
        empty_rows[index]['Дата поставки'] = supply_months[index]
    data = data._append(empty_rows, ignore_index=True)  # фиксируем диапазон дат поставки
    data['Дата поставки'] = data['Дата поставки'].dt.strftime('%Y/%m')  # преобразование дат в формат ГГГГ/ММ
    values_for_sort = ['Завод', 'Краткий текст позиции'] if form_type == 'common' or form_type == 'nmp_info' else [
        'Краткий текст позиции']
    pivot_table_indexes = ['Раздел_ГКПЗ', 'Завод', 'Номер лота', '№ материала', 'Краткий текст позиции', 'ЕИ']
    if form_type == 'nmp_info':  # удаляет лишние столбцы для НМЦ и добавляет прогнозную цену
        pivot_table_indexes.remove('Номер лота')
        pivot_table_indexes.remove('№ материала')
        pivot_table_indexes.append('Прогнозная цена')
    pivot_table_columns = [] if form_type == 'nmp_info' else ['Дата поставки']
    pivoted_data = pd.pivot_table(data,
                                  index=pivot_table_indexes,
                                  values=['Количество'],
                                  columns=pivot_table_columns,
                                  aggfunc=np.sum).sort_values(by=values_for_sort)  # формируем общую сводную таблицу
    """Cоздаём отдельные сводные таблицы для каждого завода и раздела ГКПЗ"""
    if form_type == 'common' or form_type == 'nmp_info':
        pivots_list = [pt for budget in config.budgets
                       if (pt := pivoted_data.query(f'Раздел_ГКПЗ == ["{budget}"]')).size != 0]
    else:
        pivots_list = [pt for factory in config.factories for budget in config.budgets
                       if (pt := pivoted_data.query(f'Завод == ["{factory}"] & Раздел_ГКПЗ == ["{budget}"]')).size != 0]
    return pivots_list


def engagement_report_helper(file_name: str) -> DataFrame:
    """Подготавливает таблицу с данными для отчёта по вовлечению."""
    data = pd.read_excel(file_name, sheet_name='Sheet1')
    data.sort_values(['Завод', 'Краткий текст позиции'], inplace=True)
    data.reset_index(inplace=True)
    data.index = data.index + 1  # номера строк теперь начинаются с 1, а не с 0
    for letter in 'ABC':  # добавляет 3 пустых столбца между номером лота и № материала
        data[letter] = ''
    columns_titles = ['Завод', 'Номер лота', 'A', 'B', 'C', '№ материала', 'Краткий текст позиции', 'ЕИ',
                      'Прогнозная цена', 'Количество']
    data = data.reindex(columns=columns_titles)  # переставляет столбцы местами
    return data


def get_supply_months() -> list:
    """Создаёт список дат поставки."""
    year = config.year if config.start_month in range(1, 10) else config.year - 1
    return pd.date_range(start=f'{year}/{config.start_month}', periods=13, freq='M').to_pydatetime()
