import numpy as np
import pandas as pd

from configs import config


def pivot_helper(file_name: str, tech_task_type: str) -> list:
    """Создаёт списки сводных таблиц для каждого грузополучателя, в соответствии со статьёй бюджета"""
    supply_months = get_supply_months(config.start_month)  # годы/месяцы поставки
    data = pd.read_excel(file_name, sheet_name='Sheet1')
    data['Дата поставки'] = data['Дата поставки'].dt.strftime('%Y/%m')  # преобразование дат в формат ГГГГ/ММ
    data.rename(columns={'Раздел ГКПЗ': 'Раздел_ГКПЗ'}, inplace=True)
    data['Завод'].replace(config.kts_factories, '7Q61', inplace=True)  # объединяем позиции для КТС
    data['Завод'].replace(config.dts_factories, '7QB1', inplace=True)  # объединяем позиции для ДТС
    data['Раздел_ГКПЗ'].replace(['ИП ТПИР', 'ИП ПИП'], 'ИП_ТПИР', inplace=True)
    data['Завод'] = data['Наименование МВЗ'].map(config.crs).fillna(data['Завод'])  # распределение позиций ЦРС по заводам
    config.lot_name = data['Наименование лота'].iloc[0]  # получаем наименование лота и записываем его в конфиг-файл
    empty_rows = [config.columns for _ in supply_months]
    for index in range(len(empty_rows)):
        empty_rows[index]['Дата поставки'] = supply_months[index]
    data = data.append(empty_rows, ignore_index=True)  # фиксируем диапазон дат поставки
    values_for_sort = ['Завод', 'Краткий текст позиции'] if tech_task_type == 'common' else ['Краткий текст позиции']
    pivoted_data = pd.pivot_table(data,
                             index=['Раздел_ГКПЗ', 'Завод', 'Номер лота', '№ материала', 'Краткий текст позиции', 'ЕИ'],
                             values=['Количество'],
                             columns=['Дата поставки'],
                             aggfunc=np.sum).sort_values(by=values_for_sort)  # формируем общую сводную таблицу
    """Cоздаём отдельные сводные таблицы для каждого завода и раздела ГКПЗ"""
    if tech_task_type == 'common':
        pivots_list = [pt for budget in config.budgets
                       if (pt := pivoted_data.query(f'Раздел_ГКПЗ == ["{budget}"]')).size != 0]
    else:
        pivots_list = [pt for factory in config.factories for budget in config.budgets
                       if (pt := pivoted_data.query(f'Завод == ["{factory}"] & Раздел_ГКПЗ == ["{budget}"]')).size != 0]
    return pivots_list


def get_supply_months(start_month: int) -> list:
    """Создаёт список дат поставки"""
    return pd.date_range(start=f'{config.year - 1}/{start_month}', periods=13, freq='MS').to_pydatetime()
